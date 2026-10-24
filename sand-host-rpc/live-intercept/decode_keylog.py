#!/usr/bin/env python3
"""Déchiffre le TLS host↔backend à partir d'un keylog NSS et journalise.

Outil : tshark (`tls.keylog_file`) sur un pcap tcpdump. Le fichier de clés est
le format NSS que Wireshark appelle SSLKEYLOGFILE. Sur le Node 22 de cette box
il est produit par `--tls-keylog` (l'env SSLKEYLOGFILE est ignoré ; cf. README).

Sélection : seul un échange dont le SNI (sinon :authority / Host) matche
HB_ALLOW_HOSTS est écrit, avec leg=host-backend. Rédaction, décompression,
dé-framing Connect/gRPC et décodage protobuf viennent de
../rpc-traffic-monitor/proxy.py ; le classifieur de patte vient de
addon_log._direction.

Corps HTTP/2 : tshark réassemble les DATA et recolle le corps COMPLET dans
`http2.data.data` / `http2.body.reassembled.data` de la dernière trame, en plus
des fragments des trames précédentes. On garde ce corps réassemblé (sinon la
concaténation le compte deux fois, ou le mode --follow journalise un corps vide
avant la fin du flux). Un échange unaire n'est écrit qu'une fois la requête et
la réponse terminées (END_STREAM, ou Content-Length atteint).

Le JSON tshark (corps en clair) reste en mémoire. Seul le JSONL rédigé touche
le disque, dans RPC_TRAFFIC_LOG_DIR. Rien du keylog n'est recopié.

WebSocket : tshark 4.2 (après déchiffrement TLS, ou en clair avec
`-d tcp.port==PORT,http`) pose les trames dans la couche `websocket`. Le texte
est dans `websocket.payload.text` (déjà démasqué) ; le binaire opcode 2 est le
voisin `data.data` ; un fragment FIN=0 a un payload vide et la continuation
porte le message recollé. On écrit une ligne JSONL par message
(`protocol: websocket`), sans attendre la fin du flux — un socket CDP reste
ouvert. `--plaintext` journalise le CDP loopback (leg=host-cdp), sans keylog.
Le RSV / permessage-deflate n'est pas gonflé (voir ws_frames).

Rotation (--follow, cap --max-bytes) : arrêt de tcpdump, dernière passe de
décodage (le fichier seen évite un second JSONL pour un échange déjà écrit),
shred du pcap, puis tcpdump relancé sur le même chemin / iface / filtre.
Une session ouverte avant ce nouveau fichier reste opaque jusqu'au handshake
suivant (ClientHello TLS, ou Upgrade HTTP pour le CDP) : il doit être dans
le pcap courant. Pas de nft, pas de mitm, pas de restart host-main.
"""
from __future__ import annotations

import argparse
import datetime
import hashlib
import json
import os
import re
import shutil
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
import ws_frames  # noqa: E402

if _LI not in sys.path:
    sys.path.insert(0, _LI)
import inference_effects  # noqa: E402

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
        flags = flags[-1] if flags else {}
    if isinstance(flags, dict) and str(flags.get("http2.flags.end_stream", "")) in ("1", "true", "True"):
        return True
    raw = _first(stream.get("http2.flags"))
    if isinstance(raw, str) and raw.lower().startswith("0x"):
        try:
            # bit 0 = END_STREAM (DATA et HEADERS). SETTINGS ACK est aussi 0x1
            # mais le flux 0 est ignoré avant cet appel.
            return bool(int(raw, 16) & 0x1)
        except ValueError:
            return False
    return False


def _h2_length(stream: dict) -> int:
    try:
        return int(str(_first(stream.get("http2.length")) or "0"))
    except ValueError:
        return 0


def _byte_chunks(val) -> list:
    if val is None:
        return []
    if isinstance(val, list):
        return [parse_tshark_bytes(v) for v in val]
    return [parse_tshark_bytes(val)]


def _reassembled_bytes(stream: dict) -> bytes:
    best = b""
    fr = stream.get("http2.body.fragments")
    items = fr if isinstance(fr, list) else ([fr] if isinstance(fr, dict) else [])
    for item in items:
        if not isinstance(item, dict):
            continue
        got = parse_tshark_bytes(item.get("http2.body.reassembled.data"))
        if len(got) > len(best):
            best = got
    for got in _byte_chunks(stream.get("http2.body.reassembled.data")):
        if len(got) > len(best):
            best = got
    return best


def extract_h2_payload(stream: dict) -> tuple[bytes, str]:
    """(octets, genre) avec genre = full | fragment | empty.

    `full` est le corps HTTP déjà réassemblé par tshark (à substituer aux
    fragments). `fragment` est la charge de CETTE trame DATA, à concaténer.
    """
    reass = _reassembled_bytes(stream)
    if reass:
        return reass, "full"
    file_data = b""
    http = stream.get("http")
    https = http if isinstance(http, list) else ([http] if isinstance(http, dict) else [])
    for h in https:
        if isinstance(h, dict):
            got = parse_tshark_bytes(h.get("http.file_data"))
            if len(got) > len(file_data):
                file_data = got
    parts = [p for p in _byte_chunks(stream.get("http2.data.data")) if p]
    h2len = _h2_length(stream)
    if file_data and h2len and len(file_data) > h2len:
        return file_data, "full"
    oversized = [p for p in parts if h2len and len(p) > h2len]
    if oversized:
        return max(oversized, key=len), "full"
    if file_data and not parts:
        if h2len and len(file_data) <= h2len:
            return file_data, "fragment"
        return file_data, "full"
    if parts:
        return b"".join(parts), "fragment"
    return b"", "empty"


def _as_dicts(v):
    if v is None:
        return []
    if isinstance(v, list):
        return [x for x in v if isinstance(x, dict)]
    if isinstance(v, dict):
        return [v]
    return []


def _flag_true(v) -> bool:
    return str(v).lower() in ("1", "true")


def _parse_tshark_ws(fr: dict, data_blobs):
    """Une trame de la couche websocket de tshark 4.2 (payload déjà démasqué)."""
    try:
        opcode = int(str(_first(fr.get("websocket.opcode")) or "0").split()[0])
    except ValueError:
        return None
    fin = _flag_true(fr.get("websocket.fin"))
    mask = _flag_true(fr.get("websocket.mask"))
    rsv_raw = str(_first(fr.get("websocket.rsv")) or "0")
    try:
        rsv = int(rsv_raw, 16) if rsv_raw.lower().startswith("0x") else int(rsv_raw or "0")
    except ValueError:
        rsv = 0
    payload = fr.get("websocket.payload")
    text = None
    ping = b""
    pong = b""
    close_code = None
    close_reason = ""
    unknown = None
    if isinstance(payload, dict):
        text = payload.get("websocket.payload.text")
        if isinstance(text, list):
            text = text[-1] if text else None
        if text is not None:
            text = str(text)
        if payload.get("websocket.payload.ping") not in (None, ""):
            ping = parse_tshark_bytes(payload.get("websocket.payload.ping"))
        if payload.get("websocket.payload.pong") not in (None, ""):
            pong = parse_tshark_bytes(payload.get("websocket.payload.pong"))
        close = payload.get("websocket.payload.close")
        if isinstance(close, dict):
            try:
                close_code = int(str(close.get("websocket.payload.close.status_code") or "0"))
            except ValueError:
                close_code = None
            close_reason = str(close.get("websocket.payload.close.reason") or "")
        unknown = payload.get("websocket.payload.unknown")
    binary = b""
    if opcode == 2:
        try:
            binary = next(data_blobs)
        except StopIteration:
            binary = parse_tshark_bytes(unknown) if unknown else b""
    return {
        "opcode": opcode,
        "fin": fin,
        "mask": mask,
        "rsv": rsv,
        "text": text,
        "binary": binary,
        "ping": ping,
        "pong": pong,
        "close_code": close_code,
        "close_reason": close_reason,
    }


def extract_ws_frames(layers) -> list:
    """Trames websocket d'un paquet tshark. Les clés dupliquées sont une liste (_Merged)."""
    blobs = []
    for d in _as_dicts(layers.get("data")):
        if "data.data" in d:
            blobs.append(parse_tshark_bytes(d["data.data"]))
    if not blobs:
        for raw in find_all(layers, "data.data"):
            blobs.append(parse_tshark_bytes(raw))
    out = []
    blobs_iter = iter(blobs)
    for fr in _as_dicts(layers.get("websocket")):
        parsed = _parse_tshark_ws(fr, blobs_iter)
        if parsed is not None:
            out.append(parsed)
    return out


def _tshark_to_logical(fr: dict):
    """None si ce n'est qu'un fragment FIN=0 sans octets (tshark recolle sur la suite)."""
    opcode = int(fr.get("opcode") or 0)
    rsv = int(fr.get("rsv") or 0)
    fin = bool(fr.get("fin", True))
    mask = bool(fr.get("mask"))
    text = fr.get("text")
    if rsv and not text:
        data = fr.get("binary") or fr.get("ping") or fr.get("pong") or b""
        return ws_frames.Logical(opcode if opcode else 2, data, fin, mask, True)
    if isinstance(text, str) and text != "":
        # opcode 0 + texte = message recollé par tshark, pas seulement le dernier fragment
        op = 1 if opcode in (0, 1) else opcode
        return ws_frames.Logical(op, text.encode("utf-8", "replace"), True, mask, False)
    if opcode == 2:
        if not fin and not fr.get("binary"):
            return None
        return ws_frames.Logical(2, fr.get("binary") or b"", fin, mask, False)
    if opcode == 8:
        reason = (fr.get("close_reason") or "").encode("utf-8", "replace")
        code = fr.get("close_code")
        payload = code.to_bytes(2, "big") + reason if isinstance(code, int) else reason
        return ws_frames.Logical(8, payload, True, mask, False)
    if opcode == 9:
        return ws_frames.Logical(9, fr.get("ping") or b"", True, mask, False)
    if opcode == 10:
        return ws_frames.Logical(10, fr.get("pong") or b"", True, mask, False)
    if not fin:
        return None
    return None


def _payload_fp(obj) -> str:
    try:
        raw = json.dumps(obj, ensure_ascii=False, sort_keys=True, default=str)
    except (TypeError, ValueError):
        raw = str(obj)
    return hashlib.sha256(raw.encode("utf-8", "replace")).hexdigest()[:16]


def _ws_digest(parts: list) -> str:
    return hashlib.sha256("|".join(parts).encode("utf-8", "replace")).hexdigest()


def _content_length(pairs) -> int | None:
    raw = header_get(pairs, "content-length").strip()
    if not raw or not raw.isdigit():
        return None
    return int(raw)


def _side_done(end: bool, headers, body: bytes, full: bool) -> bool:
    """Vrai quand ce côté de l'échange unaire est complet.

    Un passage --follow qui n'a que les HEADERS (content-length déjà grand,
    corps encore vide) ne matche pas. Content-Length connu et atteint, ou
    END_STREAM avec le corps réassemblé, clos le côté.
    """
    cl = _content_length(headers)
    if cl is not None and len(body) == cl:
        return True
    if not end:
        return False
    if cl is None:
        return True
    return bool(full)


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
        self.req_end = False
        self.resp_end = False
        self.req_full = False
        self.resp_full = False
        self.ws = False
        self.ws_c2s = b""
        self.ws_s2c = b""
        self.ws_c2s_full = False
        self.ws_s2c_full = False

    def complete(self) -> bool:
        return bool(self.ready_req and self.ready_resp and self.method and self.path is not None)

    def h2_finished(self) -> bool:
        if not (self.complete() and self.status):
            return False
        return (_side_done(self.req_end, self.req_headers, self.req_body, self.req_full)
                and _side_done(self.resp_end, self.resp_headers, self.resp_body, self.resp_full))

    def store_h2_body(self, is_req: bool, payload: bytes, kind: str) -> None:
        if kind == "empty" or (not payload and kind != "full"):
            return
        if is_req:
            if kind == "full":
                self.req_body = payload
                self.req_full = True
            elif not self.req_full:
                self.req_body += payload
        else:
            if kind == "full":
                self.resp_body = payload
                self.resp_full = True
            elif not self.resp_full:
                self.resp_body += payload

    def store_ws_bytes(self, is_req: bool, payload: bytes, kind: str) -> None:
        """Octets HTTP/2 d'un flux :protocol=websocket, si tshark n'a pas de couche websocket."""
        if kind == "empty" or (not payload and kind != "full"):
            return
        if is_req:
            if kind == "full":
                self.ws_c2s = payload
                self.ws_c2s_full = True
            elif not self.ws_c2s_full:
                self.ws_c2s += payload
        else:
            if kind == "full":
                self.ws_s2c = payload
                self.ws_s2c_full = True
            elif not self.ws_s2c_full:
                self.ws_s2c += payload


class WsFlow:
    """Un handshake + les trames tshark qui suivent, sur un tcp.stream (et un sid h2)."""

    def __init__(self):
        self.path = "/"
        self.host = ""
        self.server_port = ""
        self.req_headers = []
        self.resp_headers = []
        self.status = ""
        self.http_method = "GET"
        self.frames = []
        self.saw_dissector = False
        self.h2_exchange: Exchange | None = None


class Assembler:
    def __init__(self, allow_rx: re.Pattern, tunnel: str, mode: str = "tls",
                 cdp_ports: set | None = None):
        self.allow_rx = allow_rx
        self.tunnel = tunnel
        self.mode = mode  # "tls" (host↔backend) ou "cdp" (loopback en clair)
        self.cdp_ports = {str(p) for p in (cdp_ports or set())}
        self.h1 = {}   # tcp.stream -> {queue:[Exchange]}
        self.h2 = {}   # (tcp.stream, sid) -> Exchange
        self.sni = {}  # tcp.stream -> str
        self.exclude = set()
        self._seq = 0
        self.flows = {}  # ("h1", stream) | ("h2", stream, sid) -> WsFlow

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
        has_ws = "websocket" in layers
        if "http" in layers:
            self._note_http(layers["http"], src, dst, sport, dport, stream)
        if "http2" in layers:
            self._note_h2(layers["http2"], src, dst, sport, dport, stream, parse_ws_data=not has_ws)
        if has_ws:
            self._note_ws(layers, src, dst, sport, dport, stream)

    def _flow(self, key) -> WsFlow:
        flow = self.flows.get(key)
        if flow is None:
            flow = WsFlow()
            self.flows[key] = flow
        return flow

    def _upgrade_token(self, http_obj, pairs) -> str:
        up = header_get(pairs, "upgrade")
        if not up:
            up = first_str(find_all(http_obj, "http.upgrade")) or ""
        return up

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
                if "websocket" in self._upgrade_token(http_obj, ex.req_headers).lower():
                    ex.ws = True
                    flow = self._flow(("h1", stream))
                    flow.path = ex.path
                    flow.http_method = method
                    flow.req_headers = ex.req_headers
                    flow.host = bare_host(header_get(ex.req_headers, "host") or dst)
                    flow.server_port = dport
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
                up = self._upgrade_token(http_obj, ex.resp_headers)
                if "websocket" in up.lower() or (str(code) == "101" and ex.ws):
                    ex.ws = True
                    flow = self._flow(("h1", stream))
                    flow.status = str(code)
                    flow.resp_headers = ex.resp_headers
                    if ex.path:
                        flow.path = ex.path
                    if not flow.server_port:
                        flow.server_port = ex.server_port or sport
                    if not flow.host:
                        flow.host = bare_host(header_get(ex.req_headers, "host") or dst)
                    if ex.method:
                        flow.http_method = ex.method

    def _note_h2(self, layer, src, dst, sport, dport, stream, parse_ws_data: bool = True):
        if stream in self.exclude:
            return
        for st in iter_h2(layer):
            sid = str(_first(st.get("http2.streamid")) or "")
            if sid in ("", "0"):
                continue
            typ = str(_first(st.get("http2.type")) or "")
            key = (stream, sid)
            ex = self.h2.get(key)
            end = _end_stream(st)
            has_headers = typ == "1" or bool(st.get("http2.header"))
            payload, kind = extract_h2_payload(st)
            has_data = typ == "0" or kind != "empty" or "http2.data.data" in st or bool(_reassembled_bytes(st))
            if has_headers:
                pairs = h2_header_pairs(st)
                if pairs:
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
                        if end:
                            ex.req_end = True
                        if names.get(":protocol", "").lower() == "websocket":
                            ex.ws = True
                            flow = self._flow(("h2", stream, sid))
                            flow.path = ex.path
                            flow.http_method = ex.method or "CONNECT"
                            flow.req_headers = list(ex.req_headers)
                            flow.host = bare_host(names.get(":authority") or "")
                            flow.server_port = dport
                            flow.h2_exchange = ex
                    elif ":status" in names:
                        ex.status = names.get(":status")
                        ex.resp_headers = [(k, v) for k, v in pairs if not k.startswith(":")]
                        ex.ready_resp = True
                        if not ex.server_port:
                            ex.server_port = sport
                        if end:
                            ex.resp_end = True
                        if ex.ws:
                            flow = self._flow(("h2", stream, sid))
                            flow.status = str(ex.status or "")
                            flow.resp_headers = list(ex.resp_headers)
                            flow.h2_exchange = ex
                            if not flow.server_port:
                                flow.server_port = ex.server_port
                    elif end and ex.ready_resp:
                        # trailers gRPC (HEADERS sans :status, END_STREAM)
                        ex.resp_end = True
                        for k, v in pairs:
                            if not k.startswith(":"):
                                ex.resp_headers.append((k, v))
                    elif end and ex.ready_req and not ex.ready_resp:
                        ex.req_end = True
                    dur = _float_or_none(first_str(find_all(st, "http2.time")))
                    if dur is not None:
                        ex.dur = dur
            if has_data and not has_headers:
                if ex is None:
                    ex = self._new()
                    self.h2[key] = ex
                src_id = f"{src}:{sport}"
                is_req = (not ex.req_src) or src_id == ex.req_src
                if ex.ws and parse_ws_data:
                    ex.store_ws_bytes(is_req, payload, kind)
                elif not ex.ws:
                    ex.store_h2_body(is_req, payload, kind)
                if end:
                    if is_req:
                        ex.req_end = True
                    else:
                        ex.resp_end = True
            elif has_data and has_headers:
                # une même trame tshark porte parfois HEADERS et DATA
                if ex is None:
                    ex = self._new()
                    self.h2[key] = ex
                src_id = f"{src}:{sport}"
                is_req = (not ex.req_src) or src_id == ex.req_src
                if ex.ws and parse_ws_data:
                    ex.store_ws_bytes(is_req, payload, kind)
                elif not ex.ws:
                    ex.store_h2_body(is_req, payload, kind)
                if end:
                    if is_req:
                        ex.req_end = True
                    else:
                        ex.resp_end = True

    def _note_ws(self, layers, src, dst, sport, dport, stream):
        if stream in self.exclude:
            return
        flow = self._flow_for_ws(stream, layers, dport)
        frames = extract_ws_frames(layers)
        if frames:
            flow.saw_dissector = True
            flow.frames.extend(frames)
        if not flow.server_port:
            # le port serveur est celui qui n'est pas éphémère quand on le connaît déjà ;
            # à défaut, le dport de ce paquet (souvent le client). Le handshake l'écrase.
            flow.server_port = dport

    def _flow_for_ws(self, stream: str, layers, dport: str) -> WsFlow:
        sids = []
        for s in find_all(layers, "http2.streamid"):
            sid = str(s)
            if sid not in ("", "0"):
                sids.append(sid)
        if sids:
            return self._flow(("h2", stream, sids[0]))
        h2s = [k for k in self.flows if isinstance(k, tuple) and len(k) == 3 and k[0] == "h2" and k[1] == stream]
        if len(h2s) == 1 and "http2" in layers:
            return self._flow(h2s[0])
        return self._flow(("h1", stream))

    def records(self):
        out = []
        seen_local = set()

        def consider(ex: Exchange, stream: str, h2: bool = False):
            # Un Upgrade websocket n'est pas un échange unaire : les trames partent
            # via _ws_records, y compris tant que le flux reste ouvert.
            if ex.ws:
                return
            # HTTP/2 : ne pas écrire tant que les DATA ne sont pas là. Un passage
            # --follow voyait les HEADERS (content-length déjà grand) et posait
            # resp_body null, puis le digest changeait ou le passage suivant
            # échouait — la ligne vide restait.
            if h2:
                if not ex.h2_finished():
                    return
            elif not ex.complete():
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
                "req_body": proxy.decode_body(ex.req_body, req_ct, req_ce, path, "request"),
                "resp_headers": proxy.redact_headers(ex.resp_headers),
                "resp_body": proxy.decode_body(ex.resp_body, resp_ct, resp_ce, path, "response"),
                "_digest": digest,
            }
            out.append(rec)

        for stream, conn in self.h1.items():
            for ex in conn["queue"]:
                consider(ex, stream, False)
        for (stream, _sid), ex in self.h2.items():
            consider(ex, stream, True)
        out.extend(self._ws_records(seen_local))
        return out

    def _ws_records(self, seen_local: set) -> list:
        out = []
        for key, flow in self.flows.items():
            stream = key[1] if isinstance(key, tuple) and len(key) >= 2 else ""
            sid = key[2] if isinstance(key, tuple) and len(key) == 3 else ""
            if stream in self.exclude:
                continue
            if not flow.saw_dissector and flow.h2_exchange is not None:
                self._frames_from_h2_bytes(flow)
            if not flow.status and not flow.frames:
                continue
            if not self._ws_allowed(flow, stream):
                continue
            host, port, scheme = self._ws_endpoint(flow, stream)
            path = flow.path or "/"
            if path and not path.startswith("/"):
                path = "/" + path
            url = f"{scheme}://{host}:{port}{path}"
            leg, hs_direction = self._ws_leg()
            if flow.status or flow.req_headers:
                rec = ws_frames.build_handshake_record(
                    leg=leg, direction=hs_direction, host=host, path=path, url=url,
                    status=flow.status or "101",
                    req_headers=flow.req_headers, resp_headers=flow.resp_headers,
                    http_method=flow.http_method or "GET",
                )
                digest = _ws_digest(["hs", stream, sid, path, rec["status"], flow.http_method or ""])
                if digest not in seen_local:
                    seen_local.add(digest)
                    rec["_digest"] = digest
                    out.append(rec)
            for i, fr in enumerate(flow.frames):
                built = self._frame_to_record(fr, leg, host, path, url, flow)
                if built is None:
                    continue
                digest = _ws_digest([
                    "fr", stream, sid, str(i), built.get("ws_opcode") or "",
                    built.get("direction") or "",
                    _payload_fp(built.get("req_body")),
                    _payload_fp(built.get("resp_body")),
                ])
                if digest in seen_local:
                    continue
                seen_local.add(digest)
                built["_digest"] = digest
                out.append(built)
        return out

    def _frames_from_h2_bytes(self, flow: WsFlow) -> None:
        """Repli : DATA HTTP/2 du CONNECT :protocol=websocket, sans couche websocket."""
        ex = flow.h2_exchange
        if ex is None or flow.saw_dissector:
            return
        for blob, masked in ((ex.ws_c2s, True), (ex.ws_s2c, False)):
            if not blob:
                continue
            frames, _rest = ws_frames.parse_frames(blob)
            reasm = ws_frames.Reassembler()
            for fr in frames:
                for msg in reasm.push(fr):
                    flow.frames.append({
                        "logical": msg,
                        "masked": masked,
                    })

    def _ws_leg(self) -> tuple[str, str]:
        if self.mode == "cdp":
            return "host-cdp", "host->cdp"
        return "host-backend", "host->backend"

    def _ws_allowed(self, flow: WsFlow, stream: str) -> bool:
        port = str(flow.server_port or "")
        if ws_frames.port_excluded(port):
            return False
        if self.mode == "cdp":
            if self.cdp_ports and port not in self.cdp_ports:
                return False
            return True
        sni = self.sni.get(stream) or ""
        authority = header_get(flow.req_headers, ":authority") or header_get(flow.req_headers, "host")
        logical = sni or authority or flow.host
        if not host_allowed(logical, self.allow_rx):
            return False
        host = bare_host(logical)
        try:
            pnum = int(port or "443")
        except ValueError:
            pnum = 443
        return addon_log._direction(host, pnum) == "host->backend"

    def _ws_endpoint(self, flow: WsFlow, stream: str) -> tuple[str, int, str]:
        sni = self.sni.get(stream) or ""
        authority = header_get(flow.req_headers, ":authority") or header_get(flow.req_headers, "host")
        host = bare_host(sni or authority or flow.host or "127.0.0.1")
        try:
            port = int(flow.server_port or ("443" if self.mode == "tls" else "0"))
        except ValueError:
            port = 443 if self.mode == "tls" else 0
        scheme = "ws" if self.mode == "cdp" else "wss"
        return host, port, scheme

    def _frame_to_record(self, fr: dict, leg: str, host: str, path: str, url: str, flow: WsFlow):
        if "logical" in fr:
            msg = fr["logical"]
            masked = bool(fr.get("masked", msg.masked))
            side = "request" if masked else "response"
            name, payload = ws_frames.describe(msg, path, side)
            client = masked
            return ws_frames.build_frame_record(
                leg=leg,
                direction=self._frame_direction(client),
                host=host, path=path, url=url, opcode_name=name, payload=payload,
                client_to_server=client, req_headers=flow.req_headers,
                resp_headers=flow.resp_headers, status=flow.status or "101",
                masked=masked, fin=msg.fin, compressed=msg.compressed,
            )
        msg = _tshark_to_logical(fr)
        if msg is None:
            return None
        masked = bool(fr.get("mask"))
        side = "request" if masked else "response"
        name, payload = ws_frames.describe(msg, path, side)
        return ws_frames.build_frame_record(
            leg=leg,
            direction=self._frame_direction(masked),
            host=host, path=path, url=url, opcode_name=name, payload=payload,
            client_to_server=masked, req_headers=flow.req_headers,
            resp_headers=flow.resp_headers, status=flow.status or "101",
            masked=masked, fin=msg.fin, compressed=msg.compressed,
        )

    def _frame_direction(self, client_to_server: bool) -> str:
        if self.mode == "cdp":
            return "host->cdp" if client_to_server else "cdp->host"
        return "host->backend" if client_to_server else "backend->host"


def load_seen(path: str) -> set:
    if not path or not os.path.isfile(path):
        return set()
    with open(path, "r", encoding="utf-8") as f:
        return {ln.strip() for ln in f if ln.strip()}


def append_seen(path: str, digest: str) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    # O_RDONLY vaut 0 : sans O_WRONLY, open() ouvre en lecture et write()
    # renvoie EBADF. La première ligne JSONL part, les suivantes sont perdues.
    fd = os.open(path, os.O_APPEND | os.O_CREAT | os.O_WRONLY, 0o600)
    try:
        os.write(fd, (digest + "\n").encode("ascii"))
    finally:
        os.close(fd)


def tshark_cmd(pcap: str, keylog: str, tls_ports: list[str], tshark: str,
               plaintext: bool = False, cdp_ports: list[str] | None = None) -> list[str]:
    cmd = [tshark, "-r", pcap, "-n", "-o", "tcp.desegment_tcp_streams:TRUE"]
    if plaintext:
        # CDP loopback : HTTP en clair sur un port non standard. Sans -d, tshark
        # ne voit pas l'Upgrade et ne dissèque pas les trames websocket.
        cmd += [
            "-T", "json",
            "-Y", "http.request || http.response || http.upgrade || websocket",
        ]
        for port in cdp_ports or []:
            if port and port.isdigit():
                cmd.extend(["-d", f"tcp.port=={int(port)},http"])
        return cmd
    cmd += [
        "-o", f"tls.keylog_file:{keylog}",
        "-o", "http.decompress_body:FALSE",
        "-o", "tls.desegment_ssl_records:TRUE",
        "-o", "tls.desegment_ssl_application_data:TRUE",
        "-T", "json",
        "-Y",
        "tls.handshake.extensions_server_name || http.request || http.response "
        "|| http2.header.name || http2.data.data || http2.body.reassembled.data "
        "|| http.file_data || websocket",
    ]
    for port in tls_ports:
        if port and port != "443" and port.isdigit():
            cmd.extend(["-d", f"tcp.port=={int(port)},tls"])
    return cmd


class _Merged(list):
    """Liste née de clés JSON dupliquées (plusieurs trames HTTP/2 dans un paquet)."""


def _merge_json_pairs(pairs):
    """tshark -T json répète `http2.stream` / `http2.data.data`. json.loads n'en
    garde qu'une ; on les empile pour ne pas perdre de DATA."""
    out = {}
    for k, v in pairs:
        if k not in out:
            out[k] = v
            continue
        cur = out[k]
        if isinstance(cur, _Merged):
            cur.append(v)
        else:
            out[k] = _Merged([cur, v])
    return out


def run_tshark(pcap: str, keylog: str, tls_ports: list[str], tshark: str,
               plaintext: bool = False, cdp_ports: list[str] | None = None):
    cmd = tshark_cmd(pcap, keylog, tls_ports, tshark, plaintext=plaintext, cdp_ports=cdp_ports)
    proc = subprocess.run(cmd, capture_output=True)
    raw = proc.stdout.strip()
    if not raw:
        return []
    try:
        data = json.loads(raw, object_pairs_hook=_merge_json_pairs)
    except json.JSONDecodeError:
        return None
    if isinstance(data, dict):
        return [data]
    if isinstance(data, list):
        return data
    return []


def decode_once(pcap: str, keylog: str, seen_path: str, tls_ports: list[str],
                allow_rx: re.Pattern, tunnel: str, tshark: str, *,
                plaintext: bool = False, cdp_ports: list[str] | None = None) -> int:
    if not os.path.isfile(pcap) or os.path.getsize(pcap) <= 24:
        return 0
    if not plaintext and (not keylog or not os.path.isfile(keylog) or os.path.getsize(keylog) == 0):
        return 0
    frames = run_tshark(pcap, keylog, tls_ports, tshark, plaintext=plaintext, cdp_ports=cdp_ports)
    if frames is None:
        print("decode_keylog: sortie tshark illisible, nouvel essai au passage suivant",
              file=sys.stderr)
        return 0
    asm = Assembler(
        allow_rx, tunnel,
        mode="cdp" if plaintext else "tls",
        cdp_ports=set(cdp_ports or []),
    )
    for fr in frames:
        layers = (fr.get("_source") or {}).get("layers") or {}
        if isinstance(layers, dict):
            asm.note_packet(layers)
    seen = load_seen(seen_path)
    n = 0
    relevant = False
    log_dir = proxy.resolve_log_dir()
    for rec in asm.records():
        digest = rec.pop("_digest")
        if digest in seen:
            continue
        if inference_effects.is_correlation_input(rec):
            relevant = True
        proxy.write_traffic_record(rec, log_dir)
        append_seen(seen_path, digest)
        seen.add(digest)
        n += 1
        print(
            f"journalisé leg={rec['leg']} host={rec['host']} "
            f"http={rec['http_method']} path={rec['method']} status={rec['status']}"
            + (f" proto={rec['protocol']}" if rec.get("protocol") else ""),
            file=sys.stderr,
        )
    if relevant:
        try:
            added = inference_effects.materialize(log_dir)
            if added:
                print(f"inference-effect: {added} tour(s)", file=sys.stderr)
        except Exception as e:
            print(f"decode_keylog: inference-effect ignoré ({e})", file=sys.stderr)
    return n


def _split_ports(raw: str) -> list[str]:
    return [x for x in re.split(r"[\s,]+", (raw or "").strip()) if x]


# Ports vus sur cette box (DevTools). La découverte interroge aussi les
# listeners loopback : un Chrome qui a bougé de port répond à /json/version.
KNOWN_CDP_PORTS = (9229, 9257, 9302)

# Adresses locales dans /proc/net/tcp (IPv4) et /proc/net/tcp6.
_LOCAL_IP_HEX = {
    "0100007F",  # 127.0.0.1
    "00000000",  # 0.0.0.0
    "00000000000000000000000001000000",  # ::1 (forme parfois observée)
    "00000000000000000000000000000001",  # ::1
    "00000000000000000000000000000000",  # ::
}


def parse_proc_net_listeners(text: str) -> list[int]:
    """Ports en écoute (state 0A) sur loopback ou sur toutes les interfaces."""
    ports = []
    for line in text.splitlines()[1:]:
        parts = line.split()
        if len(parts) < 4:
            continue
        local, state = parts[1], parts[3]
        if state.upper() != "0A":
            continue
        ip_hex, _, port_hex = local.partition(":")
        if ip_hex not in _LOCAL_IP_HEX:
            continue
        try:
            ports.append(int(port_hex, 16))
        except ValueError:
            continue
    return ports


def _probe_cdp(port: int) -> bool:
    """Vrai si http://127.0.0.1:port/json/version expose webSocketDebuggerUrl."""
    import urllib.request
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/version", timeout=0.25) as resp:
            raw = resp.read(8192)
        data = json.loads(raw.decode("utf-8", "replace"))
    except Exception:
        return False
    return isinstance(data, dict) and bool(data.get("webSocketDebuggerUrl"))


def discover_cdp_ports(explicit: list[str] | None = None) -> list[int]:
    """Ports à capturer.

    `CDP_PORTS` (explicit) est repris tel quel, moins 6080/6081 (noVNC).
    Sans liste, on sonde les ports connus et les listeners loopback.
    """
    if explicit:
        out = []
        for item in explicit:
            if not str(item).isdigit():
                continue
            n = int(item)
            if n in ws_frames.EXCLUDED_PORTS:
                continue
            out.append(n)
        return out
    ordered: list[int] = []
    seen: set[int] = set()

    def add(n: int) -> None:
        if n in seen or n in ws_frames.EXCLUDED_PORTS or not (0 < n < 65536):
            return
        seen.add(n)
        ordered.append(n)

    for n in KNOWN_CDP_PORTS:
        add(n)
    for path in ("/proc/net/tcp", "/proc/net/tcp6"):
        try:
            with open(path, encoding="utf-8", errors="replace") as fh:
                for n in parse_proc_net_listeners(fh.read()):
                    add(n)
        except OSError:
            pass
    from concurrent.futures import ThreadPoolExecutor
    hits: list[int] = []
    sample = ordered[:80]
    with ThreadPoolExecutor(max_workers=16) as pool:
        futs = {pool.submit(_probe_cdp, n): n for n in sample}
        for fut, n in futs.items():
            try:
                if fut.result():
                    hits.append(n)
            except Exception:
                pass
    return sorted(hits)


# Reliquat renommé le temps du shred. hb-off / cdp-off effacent aussi ce chemin.
PCAP_CLOSED_SUFFIX = ".closed"
# -B de tcpdump, en Kio. 0 = ne pas passer -B (défaut historique de tcpdump : 2048).
DEFAULT_TCPDUMP_BUFFER_KB = 8192
# Après « listening », laisser le filtre s'armer (même délai que start-capture-*.sh).
ROTATE_SETTLE_S = 0.4

ROTATE_LOG = (
    "decode_keylog: pcap au-delà du cap, rotation — nouveau pcap, capture reprise. "
    "Les sessions déjà ouvertes restent opaques jusqu'au prochain handshake "
    "(ClientHello / Upgrade doit être dans le pcap courant)."
)


def _env_int(name: str, default: int) -> int:
    raw = os.environ.get(name)
    if raw is None or not str(raw).strip():
        return default
    try:
        return int(str(raw).strip())
    except ValueError:
        return default


def read_pidfile(path: str) -> int:
    if not path or not os.path.isfile(path):
        return 0
    try:
        with open(path, encoding="utf-8") as fh:
            return int(fh.read().strip() or "0")
    except (OSError, ValueError):
        return 0


def pid_alive(pid: int) -> bool:
    """Vrai si le pid existe et n'est pas un zombie (kill -0 reste vrai sur un zombie)."""
    if pid <= 0:
        return False
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        pass
    except OSError:
        return False
    try:
        with open(f"/proc/{pid}/stat", encoding="utf-8", errors="replace") as fh:
            line = fh.read()
    except OSError:
        return True
    rest = line.rsplit(")", 1)[-1].strip()
    state = rest.split(" ", 1)[0] if rest else ""
    return bool(state) and state != "Z"


def _signal_pid(pid: int, sig: int) -> None:
    if pid <= 0:
        return
    try:
        os.kill(pid, sig)
    except OSError:
        pass


def wait_pid_dead(pid: int, timeout: float) -> bool:
    if pid <= 0:
        return True
    end = time.monotonic() + max(0.0, timeout)
    while time.monotonic() < end:
        if not pid_alive(pid):
            return True
        try:
            time.sleep(0.05)
        except InterruptedError:
            break
    if pid_alive(pid):
        _signal_pid(pid, signal.SIGKILL)
        kill_end = time.monotonic() + 1.0
        while time.monotonic() < kill_end:
            if not pid_alive(pid):
                return True
            try:
                time.sleep(0.05)
            except InterruptedError:
                break
    return not pid_alive(pid)


def _stop_tcpdump(pidfile: str) -> None:
    pid = read_pidfile(pidfile)
    if pid <= 0:
        return
    _signal_pid(pid, signal.SIGTERM)
    wait_pid_dead(pid, 5.0)


def tcpdump_argv(tcpdump: str, iface: str, pcap: str, capture_filter: str,
                 buffer_kb: int) -> list[str]:
    """Ligne tcpdump partagée par le premier lancement et chaque rotation.

    Le filtre est UN argument (les parenthèses CDP ne passent pas par le shell).
    """
    filt = (capture_filter or "").strip()
    if not tcpdump or not iface or not pcap or not filt:
        raise ValueError("tcpdump, iface, pcap ou filtre BPF vide")
    cmd = [
        tcpdump, "--immediate-mode", "-i", iface, "-n", "-s", "0", "-U", "-Z", "root",
    ]
    if buffer_kb and int(buffer_kb) > 0:
        cmd.extend(["-B", str(int(buffer_kb))])
    cmd.extend(["-w", pcap, filt])
    return cmd


def shred_file(path: str) -> None:
    """Efface un pcap sans jamais écrire son contenu (shred -n 1, sinon unlink)."""
    if not path or not os.path.lexists(path):
        return
    shred = shutil.which("shred")
    if shred:
        subprocess.run(
            [shred, "-u", "-n", "1", "--", path],
            check=False, capture_output=True,
        )
    if os.path.lexists(path):
        try:
            os.remove(path)
        except OSError:
            pass


def spawn_tcpdump(tcpdump: str, iface: str, pcap: str, logout: str,
                  capture_filter: str, pidfile: str, buffer_kb: int = 0) -> int:
    """Double-fork + setsid, comme start-capture-*.sh. Retourne le pid, ou 0.

    tcpdump est réattaché à init : un SIGTERM du décodeur ne le laisse pas zombie
    et ne l'emporte pas. Le père de ce fork attend le fils intermédiaire.
    """
    try:
        argv = tcpdump_argv(tcpdump, iface, pcap, capture_filter, buffer_kb)
    except ValueError as e:
        print(f"ERREUR: {e}", file=sys.stderr)
        return 0
    for directory in (os.path.dirname(pcap), os.path.dirname(pidfile), os.path.dirname(logout)):
        if directory:
            try:
                os.makedirs(directory, exist_ok=True)
            except OSError:
                pass
    try:
        os.remove(pidfile)
    except OSError:
        pass
    try:
        first = os.fork()
    except OSError as e:
        print(f"ERREUR: fork tcpdump: {e}", file=sys.stderr)
        return 0
    if first > 0:
        while True:
            try:
                os.waitpid(first, 0)
                break
            except ChildProcessError:
                break
            except InterruptedError:
                continue
        deadline = time.monotonic() + 5.0
        while time.monotonic() < deadline:
            pid = read_pidfile(pidfile)
            if pid > 0 and pid != os.getpid() and pid_alive(pid):
                return pid
            try:
                time.sleep(0.05)
            except InterruptedError:
                break
        return 0
    try:
        os.setsid()
        if os.fork() > 0:
            os._exit(0)
        os.umask(0o077)
        fd_pid = os.open(pidfile, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        try:
            os.write(fd_pid, f"{os.getpid()}\n".encode("ascii"))
        finally:
            os.close(fd_pid)
        fd = os.open(logout, os.O_WRONLY | os.O_CREAT | os.O_APPEND, 0o600)
        os.dup2(fd, 1)
        os.dup2(fd, 2)
        if fd > 2:
            os.close(fd)
        os.execvp(argv[0], argv)
    except Exception:
        os._exit(127)
    os._exit(127)


def wait_tcpdump_listening(logout: str, offset: int, pid: int, timeout: float = 5.0,
                           settle_s: float = ROTATE_SETTLE_S) -> bool:
    """Vrai si tcpdump a écrit un nouveau « listening on » après offset."""
    end = time.monotonic() + max(0.2, timeout)
    while time.monotonic() < end:
        if pid and not pid_alive(pid):
            return False
        chunk = ""
        try:
            with open(logout, "rb") as fh:
                fh.seek(max(0, offset))
                chunk = fh.read().decode("utf-8", "replace")
        except OSError:
            chunk = ""
        if "listening on" in chunk:
            if settle_s > 0 and not _stop:
                try:
                    time.sleep(settle_s)
                except InterruptedError:
                    pass
            return bool(pid_alive(pid))
        try:
            time.sleep(0.05)
        except InterruptedError:
            break
    return False


def rotate_capture(pcap: str, pidfile: str, tcpdump: str, iface: str,
                   capture_filter: str, logout: str, buffer_kb: int,
                   decode_fn, settle_s: float = ROTATE_SETTLE_S) -> bool:
    """Arrête tcpdump, décode le pcap plein, le shred, relance la capture.

    Pas de chevauchement des deux tcpdump : un même échange relu depuis deux
    pcaps n'a pas le même tcp.stream, donc pas le même digest. On ferme
    d'abord, on décode (seen inchangé pour ce qui est déjà journalisé), puis
    seulement on ouvre le fichier neuf.
    """
    if _stop:
        return False
    if not (tcpdump and iface and (capture_filter or "").strip() and logout and pidfile and pcap):
        print(
            "decode_keylog: rotation impossible (tcpdump, iface ou filtre manquant)",
            file=sys.stderr,
        )
        return False
    old = read_pidfile(pidfile)
    _signal_pid(old, signal.SIGTERM)
    wait_pid_dead(old, 5.0)
    if _stop:
        return False
    try:
        decode_fn()
    except Exception as e:
        print(f"decode_keylog: passe finale avant rotation ignorée ({e})", file=sys.stderr)
    if _stop:
        return False
    closed = pcap + PCAP_CLOSED_SUFFIX
    if os.path.lexists(closed):
        shred_file(closed)
        if os.path.lexists(closed):
            print(
                "decode_keylog: rotation reportée, reliquat non effacé",
                file=sys.stderr,
            )
            return False
    try:
        os.replace(pcap, closed)
    except OSError as e:
        print(f"decode_keylog: rotation impossible ({e})", file=sys.stderr)
        return False
    if _stop:
        shred_file(closed)
        return False
    offset = 0
    try:
        offset = os.path.getsize(logout) if os.path.isfile(logout) else 0
    except OSError:
        offset = 0
    new_pid = spawn_tcpdump(
        tcpdump, iface, pcap, logout, capture_filter, pidfile, buffer_kb,
    )
    listening = bool(new_pid) and wait_tcpdump_listening(
        logout, offset, new_pid, settle_s=settle_s,
    )
    if not listening or _stop:
        if new_pid:
            _signal_pid(new_pid, signal.SIGTERM)
            wait_pid_dead(new_pid, 2.0)
        shred_file(closed)
        if not _stop:
            print(
                "decode_keylog: nouveau tcpdump n'écoute pas après rotation",
                file=sys.stderr,
            )
        return False
    try:
        os.chmod(pcap, 0o600)
    except OSError:
        pass
    shred_file(closed)
    print(ROTATE_LOG, file=sys.stderr)
    return True


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Déchiffre un pcap TLS via keylog NSS et journalise (rédigé).")
    p.add_argument("--pcap", default="")
    p.add_argument("--keylog", default="", help="keylog NSS (inutile avec --plaintext)")
    p.add_argument("--seen", default="")
    p.add_argument("--plaintext", action="store_true",
                   help="pcap en clair (CDP loopback) : pas de keylog, leg=host-cdp")
    p.add_argument("--cdp-ports", default=os.environ.get("CDP_PORTS", ""),
                   help="ports CDP à décoder comme HTTP (découverte : --discover-cdp)")
    p.add_argument("--discover-cdp", action="store_true",
                   help="écrit sur stdout les ports loopback qui répondent à /json/version")
    p.add_argument("--tls-ports", default="", help="ports supplémentaires décoder comme TLS (séparés par espace ou virgule)")
    p.add_argument("--follow", action="store_true")
    p.add_argument("--interval", type=float, default=1.0)
    p.add_argument("--tshark", default=os.environ.get("HB_TSHARK") or "tshark")
    p.add_argument("--max-bytes", type=int, default=_env_int("HB_PCAP_MAX_BYTES", 536870912))
    p.add_argument("--tcpdump-pidfile", default=os.environ.get("HB_TCPDUMP_PIDFILE", ""))
    # Posés par start-capture-*.sh. Sans eux, le cap arrête tcpdump (ancien
    # comportement) : on ne devine pas l'iface ni le filtre (443 ≠ CDP).
    p.add_argument("--tcpdump", default="")
    p.add_argument("--iface", default="")
    p.add_argument("--capture-filter", default="")
    p.add_argument("--capture-out", default="")
    p.add_argument("--tcpdump-buffer-kb", type=int, default=_env_int("HB_TCPDUMP_BUFFER_KB", DEFAULT_TCPDUMP_BUFFER_KB))
    args = p.parse_args(argv)

    if args.discover_cdp:
        explicit = _split_ports(os.environ.get("CDP_PORTS", ""))
        found = discover_cdp_ports(explicit or None)
        print(" ".join(str(p) for p in found))
        return 0

    if not args.pcap or not args.seen:
        print("ERREUR: --pcap et --seen sont requis", file=sys.stderr)
        return 2

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
    ports = _split_ports(args.tls_ports)
    cdp_ports = [p for p in _split_ports(args.cdp_ports) if not ws_frames.port_excluded(p)]
    proxy._load_env_secrets()

    def once() -> int:
        return decode_once(
            args.pcap, args.keylog, args.seen, ports, allow_rx, tunnel, tshark,
            plaintext=args.plaintext, cdp_ports=cdp_ports,
        )

    if not args.follow:
        if not args.plaintext and not os.path.isfile(args.keylog):
            print(f"ERREUR: keylog absent: {args.keylog}", file=sys.stderr)
            return 3
        n = once()
        print(f"decode_keylog: {n} échanges journalisés", file=sys.stderr)
        return 0

    signal.signal(signal.SIGTERM, _on_stop)
    signal.signal(signal.SIGINT, _on_stop)

    def safe_once() -> None:
        try:
            once()
        except Exception as e:
            print(f"decode_keylog: passe ignorée ({e})", file=sys.stderr)

    can_rotate = bool(
        args.tcpdump.strip() and args.iface.strip() and args.capture_filter.strip()
        and args.capture_out.strip() and args.tcpdump_pidfile.strip()
    )
    rotate_failed = False
    while not _stop:
        too_big = False
        try:
            too_big = (
                os.path.isfile(args.pcap)
                and os.path.getsize(args.pcap) > args.max_bytes
            )
        except OSError:
            too_big = False
        if too_big and not _stop:
            if not can_rotate:
                print(
                    "decode_keylog: pcap au-delà du cap, arrêt de tcpdump "
                    "(relancer hb-on pour continuer)",
                    file=sys.stderr,
                )
                _stop_tcpdump(args.tcpdump_pidfile)
                break
            ok = rotate_capture(
                pcap=args.pcap,
                pidfile=args.tcpdump_pidfile,
                tcpdump=args.tcpdump,
                iface=args.iface,
                capture_filter=args.capture_filter,
                logout=args.capture_out,
                buffer_kb=args.tcpdump_buffer_kb,
                decode_fn=safe_once,
                settle_s=ROTATE_SETTLE_S,
            )
            if not ok:
                rotate_failed = not _stop
                break
        safe_once()
        end = time.time() + max(0.2, args.interval)
        while time.time() < end and not _stop:
            # SIGTERM interrompt sleep et lèverait InterruptedError, ce qui
            # sauterait la passe finale (pcap déjà fermé par tcpdump).
            try:
                time.sleep(0.2)
            except InterruptedError:
                break
    safe_once()
    # hb-off tue le pid lu au début. Une rotation entre-temps a pu en écrire
    # un autre : on le ferme ici, sinon il continuerait d'écrire un pcap que
    # hb-off shred. Déjà mort (cas habituel) : no-op.
    if args.tcpdump_pidfile:
        _stop_tcpdump(args.tcpdump_pidfile)
    if rotate_failed:
        return 4
    return 0


if __name__ == "__main__":
    sys.exit(main())
