#!/usr/bin/env python3
"""rpc-traffic-monitor — proxy inverse HTTP/1.1 qui journalise REQUÊTE **et**
RÉPONSE (corps complets) du trafic RPC Connect/gRPC, en local uniquement.

Topologie
---------
C'est un proxy inverse en clair : le client parle HTTP au proxy, le proxy
ré-origine vers un `--upstream` (http OU https). Selon qui on pointe dessus il
capture une des deux pattes :

  * agent ↔ host     : --upstream http://127.0.0.1:1340   (port sand-host)
                       puis pointer le client/probe sur http://127.0.0.1:<listen>
  * host  ↔ backend  : --upstream https://api2.cursor.sh
                       puis pointer la base URL backend sur http://127.0.0.1:<listen>
                       (ex. GROK_BOT_RPC_BASE_URL=http://127.0.0.1:<listen>)

Journalisation
--------------
Une ligne JSONL par échange dans logs/traffic-YYYY-MM-DD.jsonl :
  ts, direction, leg, method (path RPC), status, duree_ms,
  req_headers (rédigés), req_body (Connect dé-framé + gunzip), 
  resp_headers (rédigés), resp_body (Connect dé-framé + gunzip).
Les logs restent LOCAUX (logs/ est gitignoré). Tous les secrets/PII sont rédigés.
"""
from __future__ import annotations
import argparse, asyncio, datetime, gzip, json, os, re, ssl, sys, zlib

DIR = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(DIR, "logs")

# ------------------------- rédaction secrets / PII ---------------------------
REDACT_HEADERS = {"authorization", "cookie", "set-cookie", "x-cursor-access-token",
                  "x-api-key", "x-amz-security-token", "proxy-authorization"}
JWT_RE = re.compile(r'eyJ[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}')
BEARER_RE = re.compile(r'(?i)bearer\s+[A-Za-z0-9._-]{12,}')
EMAIL_RE = re.compile(r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}')
GH_RE = re.compile(r'gh[pousr]_[A-Za-z0-9]{20,}|github_pat_[A-Za-z0-9_]{20,}')
PKEY_RE = re.compile(r'-----BEGIN [A-Z ]*PRIVATE KEY-----')
# valeurs littérales des secrets d'env du host (rédaction exacte)
ENV_SECRET_VALUES: list[str] = []


def _load_env_secrets():
    for k, v in os.environ.items():
        if not v or len(v) < 8:
            continue
        if any(t in k.upper() for t in ("TOKEN", "SECRET", "CREDENTIAL", "KEY", "PASSWORD")):
            ENV_SECRET_VALUES.append(v)


def redact_text(s: str) -> str:
    if not s:
        return s
    for v in ENV_SECRET_VALUES:
        s = s.replace(v, "[REDACTED-ENV]")
    s = JWT_RE.sub("[REDACTED-JWT]", s)
    s = BEARER_RE.sub("Bearer [REDACTED]", s)
    s = GH_RE.sub("[REDACTED-GH]", s)
    s = PKEY_RE.sub("[REDACTED-PRIVATE-KEY]", s)
    s = EMAIL_RE.sub("[REDACTED-EMAIL]", s)
    return s


def redact_headers(headers: list[tuple[str, str]]) -> dict:
    out = {}
    for k, v in headers:
        out[k] = "[REDACTED]" if k.lower() in REDACT_HEADERS else redact_text(v)
    return out


# ------------------------- décodage corps ------------------------------------
def decompress(body: bytes, enc: str) -> bytes:
    try:
        enc = (enc or "").lower()
        if "gzip" in enc:
            return gzip.decompress(body)
        if "deflate" in enc:
            return zlib.decompress(body)
    except Exception:
        pass
    return body


def deframe_connect(body: bytes) -> object:
    """Dé-frame un flux Connect/gRPC : suite de [1 octet flag][4 octets BE len][payload].
    Renvoie une liste de messages JSON décodés, ou None si ce n'est pas framé."""
    msgs = []
    i, n = 0, len(body)
    while i + 5 <= n:
        flag = body[i]
        ln = int.from_bytes(body[i + 1:i + 5], "big")
        if i + 5 + ln > n:
            return None
        payload = body[i + 5:i + 5 + ln]
        try:
            msgs.append({"flag": flag, "data": json.loads(payload.decode("utf-8", "replace"))})
        except Exception:
            msgs.append({"flag": flag, "raw": redact_text(payload.decode("utf-8", "replace"))[:4000]})
        i += 5 + ln
    return msgs if (msgs and i == n) else None


def decode_body(body: bytes, ctype: str, cenc: str) -> object:
    if not body:
        return None
    raw = decompress(body, cenc)
    ct = (ctype or "").lower()
    if "connect" in ct or "grpc" in ct:
        framed = deframe_connect(raw)
        if framed is not None:
            return redact_json(framed)
    txt = raw.decode("utf-8", "replace")
    if "json" in ct:
        try:
            return redact_json(json.loads(txt))
        except Exception:
            pass
    return redact_text(txt)[:20000]


def redact_json(obj):
    if isinstance(obj, dict):
        return {k: ("[REDACTED]" if k.lower() in ("token", "accesstoken", "credential",
                                                   "authorization", "password", "secret")
                    else redact_json(v)) for k, v in obj.items()}
    if isinstance(obj, list):
        return [redact_json(x) for x in obj]
    if isinstance(obj, str):
        return redact_text(obj)
    return obj


# ------------------------- lecture HTTP/1.1 ----------------------------------
async def read_headers(reader: asyncio.StreamReader):
    line = await reader.readline()
    if not line:
        return None
    start = line.decode("latin1").rstrip("\r\n")
    headers = []
    while True:
        h = await reader.readline()
        if h in (b"\r\n", b"\n", b""):
            break
        k, _, v = h.decode("latin1").rstrip("\r\n").partition(":")
        headers.append((k.strip(), v.strip()))
    return start, headers


async def read_body(reader: asyncio.StreamReader, headers: list[tuple[str, str]]) -> bytes:
    hd = {k.lower(): v for k, v in headers}
    if hd.get("transfer-encoding", "").lower() == "chunked":
        buf = b""
        while True:
            size_line = await reader.readline()
            if not size_line:
                break
            try:
                size = int(size_line.strip().split(b";")[0], 16)
            except ValueError:
                break
            if size == 0:
                await reader.readline()
                break
            buf += await reader.readexactly(size)
            await reader.readexactly(2)
        return buf
    cl = hd.get("content-length")
    if cl:
        try:
            return await reader.readexactly(int(cl))
        except asyncio.IncompleteReadError as e:
            return e.partial
    return b""


# ------------------------- proxy ---------------------------------------------
class Proxy:
    def __init__(self, upstream: str, leg: str):
        from urllib.parse import urlparse
        u = urlparse(upstream)
        self.scheme = u.scheme or "http"
        self.host = u.hostname
        self.port = u.port or (443 if self.scheme == "https" else 80)
        self.leg = leg

    def logfile(self):
        os.makedirs(LOG_DIR, exist_ok=True)
        day = datetime.datetime.now().astimezone().strftime("%Y-%m-%d")
        return os.path.join(LOG_DIR, f"traffic-{day}.jsonl")

    def write_log(self, rec: dict):
        with open(self.logfile(), "a") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")

    async def handle(self, creader, cwriter):
        t0 = asyncio.get_event_loop().time()
        try:
            got = await read_headers(creader)
            if not got:
                cwriter.close(); return
            start, req_headers = got
            method, _, rest = start.partition(" ")
            path = rest.rsplit(" ", 1)[0]
            req_body = await read_body(creader, req_headers)
            hd = {k.lower(): v for k, v in req_headers}

            # connexion upstream
            ssl_ctx = ssl.create_default_context() if self.scheme == "https" else None
            ureader, uwriter = await asyncio.open_connection(self.host, self.port, ssl=ssl_ctx)
            out = [f"{method} {path} HTTP/1.1"]
            for k, v in req_headers:
                if k.lower() in ("host", "connection", "transfer-encoding", "accept-encoding"):
                    continue
                out.append(f"{k}: {v}")
            out.append(f"Host: {self.host}")
            out.append("Accept-Encoding: gzip")
            out.append("Connection: close")
            out.append(f"Content-Length: {len(req_body)}")
            uwriter.write(("\r\n".join(out) + "\r\n\r\n").encode("latin1") + req_body)
            await uwriter.drain()

            # réponse upstream
            rgot = await read_headers(ureader)
            if not rgot:
                cwriter.close(); uwriter.close(); return
            status_line, resp_headers = rgot
            resp_body = await read_body(ureader, resp_headers)
            rhd = {k.lower(): v for k, v in resp_headers}
            status = status_line.split(" ")[1] if len(status_line.split(" ")) > 1 else "?"

            # renvoi au client (corps tel quel)
            rout = [status_line]
            for k, v in resp_headers:
                if k.lower() in ("transfer-encoding", "connection"):
                    continue
                rout.append(f"{k}: {v}")
            rout.append("Connection: close")
            if "content-length" not in rhd:
                rout.append(f"Content-Length: {len(resp_body)}")
            cwriter.write(("\r\n".join(rout) + "\r\n\r\n").encode("latin1") + resp_body)
            await cwriter.drain()
            uwriter.close()

            dur = round((asyncio.get_event_loop().time() - t0) * 1000, 1)
            rec = {
                "ts": datetime.datetime.now().astimezone().isoformat(timespec="milliseconds"),
                "leg": self.leg,
                "direction": f"client->{self.host}:{self.port}",
                "method": path,
                "http_method": method,
                "status": status,
                "duree_ms": dur,
                "req_headers": redact_headers(req_headers),
                "req_body": decode_body(req_body, hd.get("content-type", ""), hd.get("content-encoding", "")),
                "resp_headers": redact_headers(resp_headers),
                "resp_body": decode_body(resp_body, rhd.get("content-type", ""), rhd.get("content-encoding", "")),
            }
            self.write_log(rec)
        except Exception as e:
            try:
                self.write_log({"ts": datetime.datetime.now().astimezone().isoformat(),
                                "leg": self.leg, "erreur": f"{type(e).__name__}: {e}"})
            except Exception:
                pass
        finally:
            try:
                cwriter.close()
            except Exception:
                pass


async def main_async(args):
    _load_env_secrets()
    proxy = Proxy(args.upstream, args.leg)
    server = await asyncio.start_server(proxy.handle, args.host, args.port)
    addr = ", ".join(str(s.getsockname()) for s in server.sockets)
    print(f"rpc-traffic-monitor en écoute sur {addr} -> {args.upstream} (leg={args.leg})",
          flush=True)
    print(f"logs: {proxy.logfile()}", flush=True)
    async with server:
        await server.serve_forever()


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=int(os.environ.get("RTM_PORT", "8787")))
    ap.add_argument("--upstream", default=os.environ.get("RTM_UPSTREAM", "https://api2.cursor.sh"))
    ap.add_argument("--leg", default=os.environ.get("RTM_LEG", "host-backend"),
                    help="étiquette: 'agent-host' ou 'host-backend'")
    args = ap.parse_args()
    try:
        asyncio.run(main_async(args))
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
