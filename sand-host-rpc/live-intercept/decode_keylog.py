#!/usr/bin/env python3
"""Déchiffre le TLS host↔backend à partir d'un keylog NSS et journalise.

Outil : tshark (`tls.keylog_file`) sur un pcap tcpdump. Le fichier de clés est
le format NSS que Wireshark appelle SSLKEYLOGFILE. Sur le Node 22 de cette box
il est produit par `--tls-keylog` (l'env SSLKEYLOGFILE est ignoré ; cf. README).

Sélection : seul un échange dont le SNI (sinon :authority / Host) matche
HB_ALLOW_HOSTS est écrit, avec leg=host-backend. Rédaction, décompression et
dé-framing Connect/gRPC viennent de ../rpc-traffic-monitor/proxy.py ; le
classifieur de patte vient de addon_log._direction.

Le JSON tshark (corps en clair) reste en mémoire. Seul le JSONL rédigé touche
le disque, dans RPC_TRAFFIC_LOG_DIR. Rien du keylog n'est recopié.
"""
from __future__ import annotations

import argparse
import datetime
import hashlib
import os
import re
import signal
import subprocess
import sys
import time

_LI = os.path.dirname(os.path.abspath(__file__))
_MONITOR = os.path.join(os.path.dirname(_LI), "rpc-traffic-monitor")
for _p in (_MONITOR, os.path.join(_MONITOR, "mitmproxy")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import addon_log  # noqa: E402
import proxy  # noqa: E402

# Même expression que HB_ALLOW_HOSTS dans lib.sh (re.search, ancrée en fin).
DEFAULT_ALLOW = (
    r"(^|\.)(api|api2|api3|api4|repo42)\.cursor\.sh(:443)?$"
    r"|(^|\.)(www\.)?cursor\.com(:443)?$"
    r"|(^|\.)cursor\.sh(:443)?$"
)

_HEX_TOK = re.compile(r"^[0-9a-fA-F]{2}$")
_HOST_TOK = re.compile(r"^[A-Za-z0-9._-]+$")

# Champs tshark -> nom d'en-tête, si la ligne brute n'est pas disponible.
_HEADER_FIELDS = {
    "http.authorization": "Authorization",
    "http.content_type": "Content-Type",
    "http.content_encoding": "Content-Encoding",
    "http.host": "Host",
    "http.cookie": "Cookie",
    "http.set_cookie": "Set-Cookie",
    "http.server": "Server",
}

_stop = False


def _on_stop(signum, _frame):
    global _stop
    _stop = True


def _as_list(v):
    if v is None:
        return []
    if isinstance(v, list):
        return v
    return [v]


def _first(v):
    if isinstance(v, list):
        return v[0] if v else None
    return v


def find_all(obj, key):
    found = []

    def rec(o):
        if isinstance(o, dict):
            for k, v in o.items():
                if k == key:
                    found.extend(_as_list(v))
                else:
                    rec(v)
        elif isinstance(o, list):
            for i in o:
                rec(i)

    rec(obj)
    return found


def first_str(values):
    for v in values:
        if isinstance(v, str) and v != "":
            return v
        if isinstance(v, (int, float)):
            return str(v)
    return None


def parse_tshark_bytes(val) -> bytes:
    """http.file_data / http2.data.data : hex « aa:bb » ou texte déjà décodé."""
    if val is None:
        return b""
    if isinstance(val, list):
        return b"".join(parse_tshark_bytes(v) for v in val)
    if isinstance(val, (bytes, bytearray)):
        return bytes(val)
    s = str(val)
    if not s:
        return b""
    parts = s.split(":")
    if len(parts) > 1 and all(_HEX_TOK.fullmatch(p) for p in parts):
        return bytes(int(p, 16) for p in parts)
    return s.encode("utf-8", "surrogateescape")


def header_lines(val):
    out = []
    for line in _as_list(val):
        line = str(line).strip("\r\n")
        if ":" not in line:
            continue
        k, _, v = line.partition(":")
        k, v = k.strip(), v.strip()
        if k:
            out.append((k, v))
    return out


def _has_header(pairs, name: str) -> bool:
    n = name.lower()
    return any(k.lower() == n for k, _ in pairs)


def collect_headers(http_obj, kind: str):
    """kind = 'request' | 'response'. Lignes brutes d'abord, champs connus ensuite."""
    line_key = "http.request.line" if kind == "request" else "http.response.line"
    pairs = header_lines(find_all(http_obj, line_key))
    for field, name in _HEADER_FIELDS.items():
        if _has_header(pairs, name):
            continue
        val = first_str(find_all(http_obj, field))
        if val:
            pairs.append((name, val))
    return pairs


def header_get(pairs, name: str) -> str:
    n = name.lower()
    for k, v in pairs:
        if k.lower() == n:
            return v
    return ""


def bare_host(host: str) -> str:
    h = (host or "").strip().lower().rstrip(".")
    if h.startswith("[") and "]" in h:
        return h[1:h.index("]")]
    if h.count(":") == 1 and not _looks_ipv6(h):
        left, right = h.rsplit(":", 1)
        if right.isdigit():
            return left
    return h


def _looks_ipv6(h: str) -> bool:
    return h.count(":") > 1


def host_allowed(host: str, rx: re.Pattern) -> bool:
    h = (host or "").strip()
    if not h:
        return False
    if rx.search(h):
        return True
    b = bare_host(h)
    return bool(b and b != h.lower() and rx.search(b))


def find_sni(tls_layer) -> str | None:
    for v in find_all(tls_layer, "tls.handshake.extensions_server_name"):
        if isinstance(v, str) and _HOST_TOK.fullmatch(v) and ("." in v or v == "localhost"):
            return v
    return None


def endpoints(layers):
    ip = layers.get("ip") or {}
    if isinstance(ip, list):
        ip = ip[0] if ip else {}
    src = _first(ip.get("ip.src")) if isinstance(ip, dict) else None
    dst = _first(ip.get("ip.dst")) if isinstance(ip, dict) else None
    if not src:
        ip6 = layers.get("ipv6") or {}
        if isinstance(ip6, list):
            ip6 = ip6[0] if ip6 else {}
        if isinstance(ip6, dict):
            src = _first(ip6.get("ipv6.src"))
            dst = _first(ip6.get("ipv6.dst"))
    tcp = layers.get("tcp") or {}
    if isinstance(tcp, list):
        tcp = tcp[0] if tcp else {}
    if not isinstance(tcp, dict):
        tcp = {}
    return (
        str(src or ""),
        str(dst or ""),
        str(_first(tcp.get("tcp.srcport")) or ""),
        str(_first(tcp.get("tcp.dstport")) or ""),
        str(_first(tcp.get("tcp.stream")) or ""),
    )


def iter_http(layer):
    blobs = layer if isinstance(layer, list) else [layer]
    for blob in blobs:
        if isinstance(blob, dict):
            yield blob


def iter_h2(layer):
    blobs = layer if isinstance(layer, list) else [layer]
    for blob in blobs:
        if not isinstance(blob, dict):
            continue
        if "http2.stream" in blob:
            st = blob["http2.stream"]
            chunks = st if isinstance(st, list) else [st]
            for c in chunks:
                if isinstance(c, dict):
                    yield c
        elif "http2.type" in blob or "http2.data.data" in blob or "http2.header" in blob:
            yield blob


def h2_header_pairs(stream: dict):
    raw = stream.get("http2.header")
    out = []
    for h in _as_list(raw):
        if not isinstance(h, dict):
            continue
        name = h.get("http2.header.name")
        if not name:
            continue
        val = h.get("http2.header.value")
        out.append((str(name), "" if val is None else str(val)))
    return out


def _end_stream(stream: dict) -> bool:
    flags = stream.get("http2.flags_tree") or {}
    if isinstance(flags, list):
        flags = flags[0] if flags else {}
    if isinstance(flags, dict) and str(flags.get("http2.flags.end_stream", "")) == "1":
        return True
    return False


def _float_or_none(v):
    if v is None:
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


class Exchange:
    def __init__(self, seq: int):
        self.seq = seq
        self.method = None
        self.path = None
        self.status = None
        self.req_headers = []
        self.resp_headers = []
        self.req_body = b""
        self.resp_body = b""
        self.req_src = None  # "ip:port"
        self.server_port = ""
        self.dur = None
        self.epoch = None
        self.ready_req = False
        self.ready_resp = False

    def complete(self) -> bool:
        return bool(self.ready_req and self.ready_resp and self.method and self.path is not None)


class Assembler:
    def __init__(self, allow_rx: re.Pattern, tunnel: str):
        self.allow_rx = allow_rx
        self.tunnel = tunnel
        self.h1 = {}   # tcp.stream -> {queue:[Exchange]}
        self.h2 = {}   # (tcp.stream, sid) -> Exchange
        self.sni = {}  # tcp.stream -> str
        self.exclude = set()
        self._seq = 0

    def _new(self) -> Exchange:
        self._seq += 1
        return Exchange(self._seq)

    def _conn(self, stream: str):
        return self.h1.setdefault(stream, {"queue": []})

    def note_packet(self, layers):
        src, dst, sport, dport, stream = endpoints(layers)
        if not stream:
            return
        if self.tunnel and (src == self.tunnel or dst == self.tunnel):
            self.exclude.add(stream)
            return
        sni = find_sni(layers.get("tls"))
        if sni:
            self.sni.setdefault(stream, sni)
        if "http" in layers:
            self._note_http(layers["http"], src, dst, sport, dport, stream)
        if "http2" in layers:
            self._note_h2(layers["http2"], src, dst, sport, dport, stream)

    def _note_http(self, layer, src, dst, sport, dport, stream):
        if stream in self.exclude:
            return
        conn = self._conn(stream)
        for http_obj in iter_http(layer):
            method = first_str(find_all(http_obj, "http.request.method"))
            code = first_str(find_all(http_obj, "http.response.code"))
            body = parse_tshark_bytes(find_all(http_obj, "http.file_data"))
            dur = _float_or_none(first_str(find_all(http_obj, "http.time")))
            if method:
                ex = self._new()
                ex.method = method
                ex.path = first_str(find_all(http_obj, "http.request.uri")) or "/"
                ex.req_headers = collect_headers(http_obj, "request")
                ex.req_body = body
                ex.req_src = f"{src}:{sport}"
                ex.server_port = dport
                ex.ready_req = True
                if dur is not None:
                    ex.dur = dur
                conn["queue"].append(ex)
            elif code:
                ex = None
                for cand in conn["queue"]:
                    if not cand.ready_resp:
                        ex = cand
                        break
                if ex is None:
                    ex = self._new()
                    ex.req_src = f"{dst}:{dport}"
                    ex.server_port = sport
                    conn["queue"].append(ex)
                ex.status = code
                ex.resp_headers = collect_headers(http_obj, "response")
                ex.resp_body += body
                ex.ready_resp = True
                if dur is not None:
                    ex.dur = dur

    def _note_h2(self, layer, src, dst, sport, dport, stream):
        if stream in self.exclude:
            return
        for st in iter_h2(layer):
            sid = str(_first(st.get("http2.streamid")) or "")
            if sid in ("", "0"):
                continue
            typ = str(_first(st.get("http2.type")) or "")
            key = (stream, sid)
            ex = self.h2.get(key)
            if typ == "1" or (typ == "" and st.get("http2.header")):
                pairs = h2_header_pairs(st)
                if not pairs:
                    continue
                names = {k.lower(): v for k, v in pairs}
                if ex is None:
                    ex = self._new()
                    self.h2[key] = ex
                if ":method" in names:
                    ex.method = names.get(":method")
                    ex.path = names.get(":path") or "/"
                    ex.req_headers = [(k, v) for k, v in pairs if not k.startswith(":")]
                    # garde :authority pour la rédaction / le repli d'hôte, hors pseudo de transport
                    if ":authority" in names:
                        ex.req_headers = [(":authority", names[":authority"])] + ex.req_headers
                    ex.req_src = f"{src}:{sport}"
                    ex.server_port = dport
                    ex.ready_req = True
                elif ":status" in names:
                    ex.status = names.get(":status")
                    ex.resp_headers = [(k, v) for k, v in pairs if not k.startswith(":")]
                    ex.ready_resp = True
                    if not ex.server_port:
                        ex.server_port = sport
                dur = _float_or_none(first_str(find_all(st, "http2.time")))
                if dur is not None:
                    ex.dur = dur
            elif typ == "0" or "http2.data.data" in st:
                if ex is None:
                    ex = self._new()
                    self.h2[key] = ex
                data = parse_tshark_bytes(st.get("http2.data.data"))
                src_id = f"{src}:{sport}"
                if ex.req_src and src_id == ex.req_src:
                    ex.req_body += data
                elif ex.req_src:
                    ex.resp_body += data
                else:
                    # data avant headers : on la range côté requête, corrigée si besoin
                    # quand les headers arrivent (rare ; tshark ordonne headers puis data).
                    ex.req_body += data

    def records(self):
        out = []
        seen_local = set()

        def consider(ex: Exchange, stream: str):
            if not ex.complete():
                return
            if stream in self.exclude:
                return
            sni = self.sni.get(stream) or ""
            authority = header_get(ex.req_headers, ":authority") or header_get(ex.req_headers, "host")
            logical = sni or authority
            if not host_allowed(logical, self.allow_rx):
                return
            host = bare_host(logical)
            try:
                port = int(ex.server_port or "443")
            except ValueError:
                port = 443
            direction = addon_log._direction(host, port)
            if direction != "host->backend":
                return
            ident = "|".join([
                stream,
                str(ex.seq),
                ex.req_src or "",
                ex.method or "",
                ex.path or "",
                str(ex.status or ""),
                str(len(ex.req_body)),
                str(len(ex.resp_body)),
                f"{ex.epoch or ''}",
            ])
            digest = hashlib.sha256(ident.encode("utf-8", "replace")).hexdigest()
            if digest in seen_local:
                return
            seen_local.add(digest)
            req_ct = header_get(ex.req_headers, "content-type")
            req_ce = header_get(ex.req_headers, "content-encoding")
            resp_ct = header_get(ex.resp_headers, "content-type")
            resp_ce = header_get(ex.resp_headers, "content-encoding")
            dur_ms = round(ex.dur * 1000, 1) if ex.dur is not None else None
            path = ex.path or "/"
            rec = {
                "ts": datetime.datetime.now().astimezone().isoformat(timespec="milliseconds"),
                "leg": "host-backend",
                "direction": direction,
                "host": host,
                "method": path,
                "http_method": ex.method,
                "url": f"https://{host}{path if path.startswith('/') else '/' + path}",
                "status": str(ex.status),
                "duree_ms": dur_ms,
                "req_headers": proxy.redact_headers(ex.req_headers),
                "req_body": proxy.decode_body(ex.req_body, req_ct, req_ce),
                "resp_headers": proxy.redact_headers(ex.resp_headers),
                "resp_body": proxy.decode_body(ex.resp_body, resp_ct, resp_ce),
                "_digest": digest,
            }
            out.append(rec)

        for stream, conn in self.h1.items():
            for ex in conn["queue"]:
                consider(ex, stream)
        for (stream, _sid), ex in self.h2.items():
            consider(ex, stream)
        return out


def load_seen(path: str) -> set:
    if not path or not os.path.isfile(path):
        return set()
    with open(path, "r", encoding="utf-8") as f:
        return {ln.strip() for ln in f if ln.strip()}


def append_seen(path: str, digest: str) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    fd = os.open(path, os.O_APPEND | os.O_CREAT, 0o600)
    try:
        os.write(fd, (digest + "\n").encode("ascii"))
    finally:
        os.close(fd)


def tshark_cmd(pcap: str, keylog: str, tls_ports: list[str], tshark: str) -> list[str]:
    cmd = [
        tshark, "-r", pcap, "-n",
        "-o", f"tls.keylog_file:{keylog}",
        "-o", "http.decompress_body:FALSE",
        "-o", "tcp.desegment_tcp_streams:TRUE",
        "-o", "tls.desegment_ssl_records:TRUE",
        "-o", "tls.desegment_ssl_application_data:TRUE",
        "-T", "json", "--no-duplicate-keys",
        "-Y",
        "tls.handshake.extensions_server_name || http.request || http.response "
        "|| http2.header.name || http2.data.data",
    ]
    for port in tls_ports:
        if port and port != "443" and port.isdigit():
            cmd.extend(["-d", f"tcp.port=={int(port)},tls"])
    return cmd


def run_tshark(pcap: str, keylog: str, tls_ports: list[str], tshark: str):
    import json
    cmd = tshark_cmd(pcap, keylog, tls_ports, tshark)
    proc = subprocess.run(cmd, capture_output=True)
    raw = proc.stdout.strip()
    if not raw:
        return []
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        return None
    if isinstance(data, dict):
        return [data]
    if isinstance(data, list):
        return data
    return []


def decode_once(pcap: str, keylog: str, seen_path: str, tls_ports: list[str],
                allow_rx: re.Pattern, tunnel: str, tshark: str) -> int:
    if not os.path.isfile(pcap) or os.path.getsize(pcap) <= 24:
        return 0
    if not os.path.isfile(keylog) or os.path.getsize(keylog) == 0:
        return 0
    frames = run_tshark(pcap, keylog, tls_ports, tshark)
    if frames is None:
        print("decode_keylog: sortie tshark illisible, nouvel essai au passage suivant",
              file=sys.stderr)
        return 0
    asm = Assembler(allow_rx, tunnel)
    for fr in frames:
        layers = (fr.get("_source") or {}).get("layers") or {}
        if isinstance(layers, dict):
            asm.note_packet(layers)
    seen = load_seen(seen_path)
    n = 0
    log_dir = proxy.resolve_log_dir()
    for rec in asm.records():
        digest = rec.pop("_digest")
        if digest in seen:
            continue
        proxy.write_traffic_record(rec, log_dir)
        append_seen(seen_path, digest)
        seen.add(digest)
        n += 1
        print(
            f"journalisé leg={rec['leg']} host={rec['host']} "
            f"http={rec['http_method']} path={rec['method']} status={rec['status']}",
            file=sys.stderr,
        )
    return n


def _stop_tcpdump(pidfile: str) -> None:
    if not pidfile or not os.path.isfile(pidfile):
        return
    try:
        pid = int(open(pidfile, encoding="utf-8").read().strip() or "0")
    except ValueError:
        return
    if pid <= 0:
        return
    try:
        os.kill(pid, signal.SIGTERM)
    except OSError:
        pass


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Déchiffre un pcap TLS via keylog NSS et journalise (rédigé).")
    p.add_argument("--pcap", required=True)
    p.add_argument("--keylog", required=True)
    p.add_argument("--seen", required=True)
    p.add_argument("--tls-ports", default="", help="ports supplémentaires décoder comme TLS (séparés par espace ou virgule)")
    p.add_argument("--follow", action="store_true")
    p.add_argument("--interval", type=float, default=1.0)
    p.add_argument("--tshark", default=os.environ.get("HB_TSHARK") or "tshark")
    p.add_argument("--max-bytes", type=int, default=int(os.environ.get("HB_PCAP_MAX_BYTES", "67108864")))
    p.add_argument("--tcpdump-pidfile", default=os.environ.get("HB_TCPDUMP_PIDFILE", ""))
    args = p.parse_args(argv)

    tshark = args.tshark
    if not os.path.isfile(tshark):
        found = None
        for d in os.environ.get("PATH", "").split(":"):
            cand = os.path.join(d, "tshark") if d else ""
            if cand and os.path.isfile(cand) and os.access(cand, os.X_OK):
                found = cand
                break
        if found:
            tshark = found
        elif os.path.isfile("/usr/bin/tshark"):
            tshark = "/usr/bin/tshark"
        else:
            print("ERREUR: tshark introuvable (apt-get install tshark)", file=sys.stderr)
            return 2

    allow = os.environ.get("HB_ALLOW_HOSTS") or DEFAULT_ALLOW
    try:
        allow_rx = re.compile(allow)
    except re.error as e:
        print(f"ERREUR: HB_ALLOW_HOSTS invalide: {e}", file=sys.stderr)
        return 2
    tunnel = os.environ.get("HB_TUNNEL_ADDR", "198.18.0.1")
    ports = [x for x in re.split(r"[\s,]+", args.tls_ports.strip()) if x]
    proxy._load_env_secrets()

    def once() -> int:
        return decode_once(args.pcap, args.keylog, args.seen, ports, allow_rx, tunnel, tshark)

    if not args.follow:
        if not os.path.isfile(args.keylog):
            print(f"ERREUR: keylog absent: {args.keylog}", file=sys.stderr)
            return 3
        n = once()
        print(f"decode_keylog: {n} échanges journalisés", file=sys.stderr)
        return 0

    signal.signal(signal.SIGTERM, _on_stop)
    signal.signal(signal.SIGINT, _on_stop)
    oversize_stop = False
    while not _stop:
        try:
            if os.path.isfile(args.pcap) and os.path.getsize(args.pcap) > args.max_bytes:
                print(
                    "decode_keylog: pcap au-delà du cap, arrêt de tcpdump "
                    "(relancer hb-on pour continuer)",
                    file=sys.stderr,
                )
                _stop_tcpdump(args.tcpdump_pidfile)
                oversize_stop = True
                break
        except OSError:
            pass
        once()
        end = time.time() + max(0.2, args.interval)
        while time.time() < end and not _stop:
            time.sleep(0.2)
    once()
    if oversize_stop:
        return 0
    return 0


if __name__ == "__main__":
    sys.exit(main())
