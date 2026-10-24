#!/usr/bin/env python3
"""Trames WebSocket (RFC 6455) → objets JSON rédigés, même schéma JSONL que le RPC.

Utilisé par le relai agent↔host (octets en clair) et par decode_keylog (tshark).
Le décodeur ne réinvente pas les secrets : `proxy.redact_text` / `redact_json`.

Limite assumée : RSV / permessage-deflate n'est pas gonflé. Une trame avec RSV
posé est résumée (`_ws: compressed`) et le payload brut part en hex. CDP
(DevTools) ne négocie pas cette extension.
"""
from __future__ import annotations

import datetime
import json

import proto_decode
import proxy

OPCODE_NAME = {
    0: "continuation",
    1: "text",
    2: "binary",
    8: "close",
    9: "ping",
    10: "pong",
}

# Au-delà, on ne bufferise pas (une longueur annoncée énorme n'est pas une trame CDP).
MAX_FRAME_BYTES = 8_000_000
# Texte / JSON gardé dans le JSONL. Aligné sur le plafond texte de proxy.decode_body.
MAX_TEXT_CHARS = 20_000
MAX_HEX_CHARS = 4_000

# noVNC / websockify : hors périmètre, même si quelqu'un les passe dans CDP_PORTS.
EXCLUDED_PORTS = {6080, 6081}


class RawFrame:
    __slots__ = ("fin", "rsv", "opcode", "masked", "payload")

    def __init__(self, fin: bool, rsv: int, opcode: int, masked: bool, payload: bytes):
        self.fin = fin
        self.rsv = rsv
        self.opcode = opcode
        self.masked = masked
        self.payload = payload

    @property
    def compressed(self) -> bool:
        return self.rsv != 0


class Logical:
    """Message de données réassemblé, ou trame de contrôle isolée."""

    __slots__ = ("opcode", "payload", "fin", "masked", "compressed")

    def __init__(self, opcode: int, payload: bytes, fin: bool, masked: bool, compressed: bool):
        self.opcode = opcode
        self.payload = payload
        self.fin = fin
        self.masked = masked
        self.compressed = compressed


def is_client_upgrade(headers) -> bool:
    """Requête HTTP/1.1 qui demande un Upgrade: websocket."""
    hd = {k.lower(): (v or "") for k, v in headers}
    return "websocket" in hd.get("upgrade", "").lower() and "upgrade" in hd.get("connection", "").lower()


def is_switching(status, headers) -> bool:
    """Réponse 101 qui termine le handshake WebSocket."""
    if str(status) != "101":
        return False
    hd = {k.lower(): (v or "") for k, v in headers}
    if "websocket" in hd.get("upgrade", "").lower():
        return True
    return bool(hd.get("sec-websocket-accept"))


def encode_frame(opcode: int, payload: bytes = b"", *, mask: bool = False,
                 fin: bool = True, rsv: int = 0, mask_key: bytes | None = None) -> bytes:
    """Construit une trame. `mask_key` fixe la clé (tests) ; sinon 4 octets nuls si mask."""
    if mask_key is None:
        mask_key = b"\x00\x00\x00\x00"
    if len(mask_key) != 4:
        raise ValueError("mask_key")
    b0 = ((0x80 if fin else 0) | ((rsv & 7) << 4) | (opcode & 0x0F))
    n = len(payload)
    mbit = 0x80 if mask else 0
    if n < 126:
        hdr = bytes((b0, mbit | n))
    elif n < 65536:
        hdr = bytes((b0, mbit | 126)) + n.to_bytes(2, "big")
    else:
        hdr = bytes((b0, mbit | 127)) + n.to_bytes(8, "big")
    if mask:
        hdr += mask_key
        payload = bytes(b ^ mask_key[i % 4] for i, b in enumerate(payload))
    return hdr + payload


def parse_frames(buf: bytes) -> tuple[list[RawFrame], bytes]:
    """Découpe `buf`. Le reliquat est une trame incomplète (à garder pour le prochain read)."""
    frames: list[RawFrame] = []
    i = 0
    n = len(buf)
    while i + 2 <= n:
        b0 = buf[i]
        b1 = buf[i + 1]
        fin = bool(b0 & 0x80)
        rsv = (b0 >> 4) & 7
        opcode = b0 & 0x0F
        masked = bool(b1 & 0x80)
        ln = b1 & 0x7F
        hdr = 2
        if ln == 126:
            hdr = 4
            if i + hdr > n:
                break
            ln = int.from_bytes(buf[i + 2:i + 4], "big")
        elif ln == 127:
            hdr = 10
            if i + hdr > n:
                break
            ln = int.from_bytes(buf[i + 2:i + 10], "big")
        if ln > MAX_FRAME_BYTES:
            # Longueur absurde : on abandonne le buffer plutôt que de le faire grossir.
            return frames, b""
        mask_len = 4 if masked else 0
        total = hdr + mask_len + ln
        if i + total > n:
            break
        off = i + hdr
        data = buf[off + mask_len:off + mask_len + ln]
        if masked:
            key = buf[off:off + 4]
            data = bytes(b ^ key[j % 4] for j, b in enumerate(data))
        frames.append(RawFrame(fin, rsv, opcode, masked, data))
        i += total
    return frames, buf[i:]


class Reassembler:
    """Recolle les continuations (opcode 0). Les trames de contrôle passent tout de suite."""

    def __init__(self):
        self._opcode: int | None = None
        self._buf = bytearray()
        self._masked = False

    def push(self, fr: RawFrame) -> list[Logical]:
        if fr.opcode >= 8:
            return [Logical(fr.opcode, fr.payload, fr.fin, fr.masked, fr.compressed)]
        if fr.compressed:
            # Pas de déflation : on ne mélange pas ces octets avec un fragment en cours.
            return [Logical(fr.opcode, fr.payload, fr.fin, fr.masked, True)]
        if fr.opcode == 0:
            if self._opcode is None:
                return []
            self._buf += fr.payload
            if not fr.fin:
                return []
            op = self._opcode
            data = bytes(self._buf)
            masked = self._masked
            self._opcode = None
            self._buf.clear()
            return [Logical(op, data, True, masked, False)]
        if fr.opcode in (1, 2):
            if not fr.fin:
                self._opcode = fr.opcode
                self._buf = bytearray(fr.payload)
                self._masked = fr.masked
                return []
            return [Logical(fr.opcode, fr.payload, True, fr.masked, False)]
        return [Logical(fr.opcode, fr.payload, fr.fin, fr.masked, fr.compressed)]


def _now() -> str:
    return datetime.datetime.now().astimezone().isoformat(timespec="milliseconds")


def _hex(data: bytes) -> str:
    h = data.hex()
    if len(h) > MAX_HEX_CHARS:
        return h[:MAX_HEX_CHARS]
    return h


def decode_text(data: bytes):
    if len(data) > MAX_TEXT_CHARS:
        return {
            "_truncated": True,
            "text": proxy.redact_text(data[:MAX_TEXT_CHARS].decode("utf-8", "replace")),
        }
    txt = data.decode("utf-8", "replace")
    try:
        obj = json.loads(txt)
    except Exception:
        return proxy.redact_text(txt)
    return proxy.redact_json(obj)


def decode_binary(data: bytes, path: str | None, side: str | None):
    """Proto nommé seulement si le path est dans le catalogue. Sinon hex."""
    name = proto_decode.message_name_for(path, side)
    if name:
        try:
            return proxy.redact_json(proto_decode.decode_message(data, name))
        except Exception:
            return {
                "_proto_error": "decode",
                "_binary_hex": _hex(data),
                "_binary_len": len(data),
            }
    return {"_binary_hex": _hex(data), "_binary_len": len(data)}


def summarize_control(opcode: int, data: bytes) -> dict:
    name = OPCODE_NAME.get(opcode, str(opcode))
    if opcode == 8:
        code = int.from_bytes(data[:2], "big") if len(data) >= 2 else None
        reason = data[2:].decode("utf-8", "replace") if len(data) > 2 else ""
        return {"opcode": "close", "code": code, "reason": proxy.redact_text(reason)[:500]}
    return {"opcode": name, "len": len(data)}


def compressed_note(opcode: int, data: bytes) -> dict:
    return {
        "_ws": "compressed",
        "opcode": OPCODE_NAME.get(opcode, str(opcode)),
        "note": "permessage-deflate (RSV) non décompressé",
        "_binary_hex": _hex(data),
        "_binary_len": len(data),
    }


def describe(msg: Logical, path: str | None, side: str | None) -> tuple[str, object]:
    """(nom d'opcode, payload JSON-able déjà rédigé)."""
    if msg.compressed:
        name = "text" if msg.opcode in (0, 1) else OPCODE_NAME.get(msg.opcode, str(msg.opcode))
        return name, compressed_note(msg.opcode, msg.payload)
    if msg.opcode in (0, 1):
        return "text", decode_text(msg.payload)
    if msg.opcode == 2:
        return "binary", decode_binary(msg.payload, path, side)
    if msg.opcode >= 8:
        return OPCODE_NAME.get(msg.opcode, str(msg.opcode)), summarize_control(msg.opcode, msg.payload)
    return OPCODE_NAME.get(msg.opcode, str(msg.opcode)), {
        "_binary_hex": _hex(msg.payload),
        "_binary_len": len(msg.payload),
    }


def _cdp_method(payload) -> str | None:
    if isinstance(payload, dict):
        m = payload.get("method")
        if isinstance(m, str) and m and not m.startswith("_"):
            return m
    return None


def _base_record(*, leg: str, direction: str, host: str, path: str, url: str,
                 http_method: str, status: str, req_headers, resp_headers) -> dict:
    return {
        "ts": _now(),
        "leg": leg,
        "direction": direction,
        "protocol": "websocket",
        "host": host,
        "method": path or "/",
        "ws_path": path or "/",
        "http_method": http_method,
        "url": url,
        "status": str(status),
        "duree_ms": None,
        "req_headers": proxy.redact_headers(req_headers or []),
        "resp_headers": proxy.redact_headers(resp_headers or []),
    }


def build_handshake_record(*, leg: str, direction: str, host: str, path: str, url: str,
                           status: str, req_headers, resp_headers, http_method: str = "GET") -> dict:
    rec = _base_record(
        leg=leg, direction=direction, host=host, path=path or "/", url=url,
        http_method=http_method, status=status or "101",
        req_headers=req_headers, resp_headers=resp_headers,
    )
    rec["ws_event"] = "handshake"
    rec["req_body"] = None
    rec["resp_body"] = None
    return rec


def build_frame_record(*, leg: str, direction: str, host: str, path: str, url: str,
                       opcode_name: str, payload, client_to_server: bool,
                       req_headers, resp_headers, status: str = "101",
                       masked: bool | None = None, fin: bool = True,
                       compressed: bool = False) -> dict:
    shown = path or "/"
    cdp = _cdp_method(payload)
    if cdp:
        shown = cdp
    rec = _base_record(
        leg=leg, direction=direction, host=host, path=shown, url=url,
        http_method=f"WS/{opcode_name}", status=status or "101",
        req_headers=req_headers, resp_headers=resp_headers,
    )
    rec["method"] = shown
    rec["ws_path"] = path or "/"
    rec["ws_event"] = "frame"
    rec["ws_opcode"] = opcode_name
    rec["ws_fin"] = bool(fin)
    if masked is not None:
        rec["ws_masked"] = bool(masked)
    if compressed:
        rec["ws_compressed"] = True
    if client_to_server:
        rec["req_body"] = payload
        rec["resp_body"] = None
    else:
        rec["req_body"] = None
        rec["resp_body"] = payload
    return rec


def port_excluded(port) -> bool:
    try:
        return int(port) in EXCLUDED_PORTS
    except (TypeError, ValueError):
        return str(port) in {"6080", "6081"}
