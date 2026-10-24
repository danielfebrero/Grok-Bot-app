#!/usr/bin/env python3
"""relay_agent_host.py — relai TCP loggeur PLAINTEXT pour la patte agent<->host.

La patte agent<->host de la box est du HTTP/1.1 EN CLAIR (prouvé : `curl http://
127.0.0.1:1340/` -> `HTTP/1.1 404` JSON ; handshake TLS refusé ; h2c refusé).
Aucune CA n'est donc nécessaire pour l'intercepter : il suffit de s'intercaler.

Topologie (posée par arm.sh) :
    agent / exec-daemon  ->  127.0.0.1:1340 (CE RELAI)  ->  127.0.0.1:1341 (host-main déplacé)

Le relai :
  * accepte le client, lit la requête HTTP/1.1 (headers + corps, content-length
    ou chunked), la ré-émet vers l'upstream host-main ;
  * lit la réponse, la renvoie au client ;
  * supporte le KEEP-ALIVE : plusieurs échanges par connexion (ce que
    proxy.py en mode `Connection: close` ne faisait pas) ;
  * journalise CHAQUE échange (req+resp, corps complets) en JSONL via les helpers
    de rédaction/décodage/rotation DÉJÀ écrits dans rpc-traffic-monitor/proxy.py,
    leg="agent-host". Zéro secret/PII en clair.

Usage :
    python3 relay_agent_host.py --listen 1340 --upstream-port 1341 [--upstream-host 127.0.0.1]
"""
from __future__ import annotations
import argparse, asyncio, datetime, os, sys

# importe les helpers existants (rédaction + décodage + log/rotation partagés EG)
_MONITOR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        "rpc-traffic-monitor")
if _MONITOR not in sys.path:
    sys.path.insert(0, _MONITOR)
import proxy  # noqa: E402  (rpc-traffic-monitor/proxy.py)

LEG = "agent-host"


async def pump_one(creader, cwriter, ureader, uwriter, log_dir) -> bool:
    """Traite UN échange req/resp. Renvoie True si la connexion peut rester
    ouverte (keep-alive), False si elle doit être fermée."""
    t0 = asyncio.get_event_loop().time()
    got = await proxy.read_headers(creader)
    if not got:
        return False
    start, req_headers = got
    parts = start.split(" ")
    method = parts[0] if parts else "?"
    path = parts[1] if len(parts) > 1 else "/"
    req_body = await proxy.read_body(creader, req_headers)
    hd = {k.lower(): v for k, v in req_headers}

    # ré-émission vers l'upstream (host-main déplacé), en gardant la connexion.
    out = [f"{method} {path} HTTP/1.1"]
    for k, v in req_headers:
        if k.lower() in ("connection", "transfer-encoding", "accept-encoding"):
            continue
        out.append(f"{k}: {v}")
    out.append("Accept-Encoding: gzip")
    out.append("Connection: keep-alive")
    if "content-length" not in hd and req_body:
        out.append(f"Content-Length: {len(req_body)}")
    uwriter.write(("\r\n".join(out) + "\r\n\r\n").encode("latin1") + req_body)
    await uwriter.drain()

    rgot = await proxy.read_headers(ureader)
    if not rgot:
        return False
    status_line, resp_headers = rgot
    resp_body = await proxy.read_body(ureader, resp_headers)
    rhd = {k.lower(): v for k, v in resp_headers}
    sp = status_line.split(" ")
    status = sp[1] if len(sp) > 1 else "?"

    # renvoi au client
    rout = [status_line]
    for k, v in resp_headers:
        if k.lower() in ("transfer-encoding", "connection"):
            continue
        rout.append(f"{k}: {v}")
    rout.append("Connection: keep-alive")
    if "content-length" not in rhd:
        rout.append(f"Content-Length: {len(resp_body)}")
    cwriter.write(("\r\n".join(rout) + "\r\n\r\n").encode("latin1") + resp_body)
    await cwriter.drain()

    dur = round((asyncio.get_event_loop().time() - t0) * 1000, 1)
    rec = {
        "ts": datetime.datetime.now().astimezone().isoformat(timespec="milliseconds"),
        "leg": LEG,
        "direction": "agent->host",
        "host": "127.0.0.1",
        "method": path,
        "http_method": method,
        "status": status,
        "duree_ms": dur,
        "req_headers": proxy.redact_headers(req_headers),
        "req_body": proxy.decode_body(req_body, hd.get("content-type", ""),
                                      hd.get("content-encoding", ""), path, "request"),
        "resp_headers": proxy.redact_headers(resp_headers),
        "resp_body": proxy.decode_body(resp_body, rhd.get("content-type", ""),
                                       rhd.get("content-encoding", ""), path, "response"),
    }
    proxy.write_traffic_record(rec, log_dir)

    # keep-alive sauf si l'un des deux a demandé close
    if hd.get("connection", "").lower() == "close":
        return False
    if rhd.get("connection", "").lower() == "close":
        return False
    return True


class Relay:
    def __init__(self, uhost, uport):
        self.uhost, self.uport = uhost, uport
        self.log_dir = proxy.resolve_log_dir()

    async def handle(self, creader, cwriter):
        uwriter = None
        try:
            ureader, uwriter = await asyncio.open_connection(self.uhost, self.uport)
            while True:
                keep = await pump_one(creader, cwriter, ureader, uwriter, self.log_dir)
                if not keep:
                    break
        except Exception as e:
            try:
                proxy.write_traffic_record(
                    {"ts": datetime.datetime.now().astimezone().isoformat(),
                     "leg": LEG, "erreur": f"{type(e).__name__}: {e}"}, self.log_dir)
            except Exception:
                pass
        finally:
            for w in (cwriter, uwriter):
                try:
                    if w is not None:
                        w.close()
                except Exception:
                    pass


def _split_hostport(s: str, default_host: str, default_port: int):
    """Accepte 'HOST:PORT', ':PORT', 'PORT' ou 'HOST'. Renvoie (host, port)."""
    s = (s or "").strip()
    if not s:
        return default_host, default_port
    if ":" in s:
        h, _, p = s.rpartition(":")
        return (h or default_host), (int(p) if p else default_port)
    if s.isdigit():
        return default_host, int(s)
    return s, default_port


async def main_async(a):
    proxy._load_env_secrets()
    lhost, lport = _split_hostport(a.listen, a.listen_host or "127.0.0.1", 1340)
    if a.upstream:
        uhost, uport = _split_hostport(a.upstream, a.upstream_host, a.upstream_port)
    else:
        uhost, uport = a.upstream_host, a.upstream_port
    r = Relay(uhost, uport)
    server = await asyncio.start_server(r.handle, lhost, lport)
    addr = ", ".join(str(s.getsockname()) for s in server.sockets)
    print(f"relay {LEG} en écoute {addr} -> {uhost}:{uport}", flush=True)
    print(f"logs: {proxy.logfile_path(r.log_dir)}", flush=True)
    async with server:
        await server.serve_forever()


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    # --listen accepte "PORT" ou "HOST:PORT" (rétro-compat : bare port) ;
    # --upstream accepte "HOST:PORT" (prioritaire sur --upstream-host/-port).
    ap.add_argument("--listen", default="1340", help="PORT ou HOST:PORT")
    ap.add_argument("--listen-host", default=None)
    ap.add_argument("--upstream", default=None, help="HOST:PORT (ex 172.30.0.2:1340)")
    ap.add_argument("--upstream-host", dest="upstream_host", default="127.0.0.1")
    ap.add_argument("--upstream-port", dest="upstream_port", type=int, default=1341)
    a = ap.parse_args()
    try:
        asyncio.run(main_async(a))
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
