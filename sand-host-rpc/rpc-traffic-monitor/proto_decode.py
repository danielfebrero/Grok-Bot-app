#!/usr/bin/env python3
"""Décode un corps protobuf (Connect unaire `application/proto`, ou trames
Connect/gRPC) en JSON à partir du catalogue `cartographie/data/rpc_catalog.json`.

Les noms de champs viennent uniquement du catalogue. Un tag absent du message
devient `_unknown_<n>` (entier, texte imprimable, ou hex). Les scalaires
utilisent les noms protobuf-es (`string`, `int32`, `sint64`, …). Un message ou
un enum imbriqué est `ref#<nom complet>`.
"""
from __future__ import annotations

import base64
import gzip
import json
import os
import struct

_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_CATALOG = os.path.normpath(
    os.path.join(_DIR, "..", "cartographie", "data", "rpc_catalog.json")
)

# Fil protobuf du scalaire (0 varint, 1 64 bits, 2 longueur, 5 32 bits).
_WIRE = {
    "int32": 0, "int64": 0, "uint32": 0, "uint64": 0,
    "sint32": 0, "sint64": 0, "bool": 0,
    "fixed64": 1, "sfixed64": 1, "double": 1,
    "string": 2, "bytes": 2,
    "fixed32": 5, "sfixed32": 5, "float": 5,
}

_MAX_DEPTH = 32
_MAX_FIELDS = 200_000
_JS_SAFE = 2 ** 53

_CACHE: dict | None = None


def load_catalog() -> dict:
    global _CACHE
    if _CACHE is not None:
        return _CACHE
    path = os.environ.get("RPC_PROTO_CATALOG") or DEFAULT_CATALOG
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, json.JSONDecodeError):
        data = {}
    data.setdefault("messages", {})
    data.setdefault("enums", {})
    data.setdefault("methods", {})
    _CACHE = data
    return data


def reset_catalog_cache() -> None:
    global _CACHE
    _CACHE = None


def is_proto_content_type(ctype: str) -> bool:
    ct = (ctype or "").lower()
    if "json" in ct:
        return False
    if "application/proto" in ct or "application/protobuf" in ct or "application/x-protobuf" in ct:
        return True
    if "grpc" in ct or ct.strip().endswith("+proto"):
        return True
    return False


def message_name_for(method: str | None, side: str | None) -> str | None:
    if not method:
        return None
    path = method.split("?", 1)[0]
    if path and not path.startswith("/"):
        path = "/" + path
    info = load_catalog()["methods"].get(path)
    if not isinstance(info, dict):
        return None
    key = "request" if side == "request" else "response"
    name = info.get(key)
    return name if isinstance(name, str) and name else None


def _read_varint(buf: bytes, i: int) -> tuple[int, int]:
    shift = 0
    n = 0
    while i < len(buf) and shift <= 63:
        b = buf[i]
        i += 1
        n |= (b & 0x7F) << shift
        if not (b & 0x80):
            return n, i
        shift += 7
    raise ValueError("varint")


def _read_value(buf: bytes, i: int, wt: int):
    if wt == 0:
        return _read_varint(buf, i)
    if wt == 1:
        if i + 8 > len(buf):
            raise ValueError("fixed64")
        return buf[i:i + 8], i + 8
    if wt == 5:
        if i + 4 > len(buf):
            raise ValueError("fixed32")
        return buf[i:i + 4], i + 4
    if wt == 2:
        ln, i = _read_varint(buf, i)
        if ln < 0 or i + ln > len(buf):
            raise ValueError("len")
        return buf[i:i + ln], i + ln
    raise ValueError(f"wire {wt}")


def _skip_group(buf: bytes, i: int) -> int:
    while i < len(buf):
        tag, i = _read_varint(buf, i)
        wt = tag & 7
        if wt == 4:
            return i
        if wt == 3:
            i = _skip_group(buf, i)
        else:
            _, i = _read_value(buf, i, wt)
    raise ValueError("groupe")


def _signed(n: int) -> int:
    return n - 2 ** 64 if n >= 2 ** 63 else n


def _zigzag(n: int) -> int:
    return (n >> 1) ^ -(n & 1)


def _num(n: int):
    """int64 hors de l'intervalle sûr JSON → chaîne (mapping proto3 JSON)."""
    if n >= _JS_SAFE or n <= -_JS_SAFE:
        return str(n)
    return n


def _finite(x: float):
    if x != x:
        return "NaN"
    if x == float("inf"):
        return "Infinity"
    if x == float("-inf"):
        return "-Infinity"
    return x


def _text_or_hex(b: bytes):
    try:
        s = b.decode("utf-8")
    except UnicodeDecodeError:
        return b.hex()
    if s and all(ch.isprintable() or ch in "\n\r\t" for ch in s):
        return s
    return b.hex()


def _decode_varint_scalar(n: int, scalar: str):
    if scalar == "bool":
        return bool(n)
    if scalar in ("sint32", "sint64"):
        return _num(_zigzag(n))
    if scalar in ("int32", "int64"):
        return _num(_signed(n))
    if scalar == "uint64":
        return _num(n) if n < _JS_SAFE else str(n)
    return _num(n)


def _decode_fixed(raw: bytes, scalar: str):
    if scalar == "double":
        return _finite(struct.unpack("<d", raw)[0])
    if scalar == "float":
        return _finite(struct.unpack("<f", raw)[0])
    if scalar == "sfixed64":
        return _num(int.from_bytes(raw, "little", signed=True))
    if scalar == "sfixed32":
        return int.from_bytes(raw, "little", signed=True)
    if scalar == "fixed64":
        n = int.from_bytes(raw, "little")
        return str(n) if n >= _JS_SAFE else n
    return int.from_bytes(raw, "little")


def _decode_scalar(raw, wt: int, scalar: str):
    expect = _WIRE.get(scalar)
    if expect == 0 and wt == 0 and isinstance(raw, int):
        return _decode_varint_scalar(raw, scalar)
    if expect == 1 and wt == 1 and isinstance(raw, (bytes, bytearray)):
        return _decode_fixed(bytes(raw), scalar)
    if expect == 5 and wt == 5 and isinstance(raw, (bytes, bytearray)):
        return _decode_fixed(bytes(raw), scalar)
    if scalar == "string" and wt == 2 and isinstance(raw, (bytes, bytearray)):
        return bytes(raw).decode("utf-8", "replace")
    if scalar == "bytes" and wt == 2 and isinstance(raw, (bytes, bytearray)):
        return base64.b64encode(bytes(raw)).decode("ascii")
    raise ValueError("scalaire")


def _unpack_packed(buf: bytes, scalar: str) -> list:
    expect = _WIRE.get(scalar)
    out = []
    i = 0
    if expect == 0:
        while i < len(buf):
            n, i = _read_varint(buf, i)
            out.append(_decode_varint_scalar(n, scalar))
        return out
    step = 8 if expect == 1 else 4 if expect == 5 else 0
    if not step or len(buf) % step:
        raise ValueError("packed")
    while i < len(buf):
        out.append(_decode_fixed(buf[i:i + step], scalar))
        i += step
    return out


def _enum_name(type_name: str, number: int):
    table = load_catalog()["enums"].get(type_name) or {}
    name = table.get(str(number))
    if name:
        return name
    signed = _signed(number) if number >= 2 ** 63 else number
    name = table.get(str(signed))
    return name if name else signed


def _store_unknown(out: dict, fno: int, wt: int, raw) -> None:
    key = f"_unknown_{fno}"
    if wt == 0 and isinstance(raw, int):
        val = _num(_signed(raw)) if raw >= 2 ** 63 else raw
    elif wt == 2 and isinstance(raw, (bytes, bytearray)):
        val = _text_or_hex(bytes(raw))
    elif isinstance(raw, (bytes, bytearray)):
        val = bytes(raw).hex()
    else:
        val = raw
    if key in out:
        cur = out[key]
        if not isinstance(cur, list):
            out[key] = [cur]
        out[key].append(val)
    else:
        out[key] = val


def _append(out: dict, name: str, value, repeated: bool) -> None:
    if repeated:
        slot = out.setdefault(name, [])
        if isinstance(value, list):
            slot.extend(value)
        else:
            slot.append(value)
    else:
        out[name] = value


def _decode_map_entry(buf: bytes, key_type: str, val_type: str, depth: int):
    key = None
    val = None
    i = 0
    while i < len(buf):
        tag, i = _read_varint(buf, i)
        fno, wt = tag >> 3, tag & 7
        raw, i = _read_value(buf, i, wt)
        if fno == 1:
            key = _decode_typed(raw, wt, key_type, depth)
        elif fno == 2:
            val = _decode_typed(raw, wt, val_type, depth)
    if isinstance(key, (dict, list)):
        key = json.dumps(key, ensure_ascii=False, sort_keys=True)
    elif key is not None and not isinstance(key, str):
        key = str(key)
    return key, val


def _decode_typed(raw, wt: int, type_token: str, depth: int):
    if type_token.startswith("ref#"):
        name = type_token[4:]
        cat = load_catalog()
        if name in cat["enums"]:
            if wt != 0 or not isinstance(raw, int):
                raise ValueError("enum")
            return _enum_name(name, raw)
        if wt != 2 or not isinstance(raw, (bytes, bytearray)):
            raise ValueError("message")
        return decode_message(bytes(raw), name if name in cat["messages"] else None, depth + 1)
    if type_token in _WIRE:
        return _decode_scalar(raw, wt, type_token)
    raise ValueError(type_token)


def _interpret(out: dict, field: dict, wt: int, raw, depth: int) -> None:
    name = field.get("name") or f"_unknown_{field.get('no')}"
    repeated = bool(field.get("repeated"))
    if field.get("type") == "map":
        if wt != 2 or not isinstance(raw, (bytes, bytearray)):
            raise ValueError("map")
        k, v = _decode_map_entry(
            bytes(raw), field.get("key") or "string", field.get("value") or "bytes", depth,
        )
        slot = out.setdefault(name, {})
        if not isinstance(slot, dict):
            slot = {}
            out[name] = slot
        if k is not None:
            slot[k] = v
        return
    type_token = field.get("type") or ""
    scalar = type_token if type_token in _WIRE else None
    if repeated and scalar and wt == 2 and _WIRE[scalar] != 2 and isinstance(raw, (bytes, bytearray)):
        _append(out, name, _unpack_packed(bytes(raw), scalar), True)
        return
    _append(out, name, _decode_typed(raw, wt, type_token, depth), repeated)


def decode_message(buf: bytes, type_name: str | None, depth: int = 0) -> dict:
    if depth > _MAX_DEPTH:
        return {"_max_depth": True}
    cat = load_catalog()
    fields = cat["messages"].get(type_name) if type_name else None
    by_no = {}
    for f in fields or []:
        if isinstance(f, dict) and isinstance(f.get("no"), int):
            by_no[f["no"]] = f
    out: dict = {}
    i = 0
    n = len(buf)
    seen = 0
    while i < n:
        seen += 1
        if seen > _MAX_FIELDS:
            out["_truncated"] = True
            break
        try:
            tag, i2 = _read_varint(buf, i)
        except ValueError:
            out["_truncated"] = True
            break
        fno, wt = tag >> 3, tag & 7
        if fno <= 0 or wt > 5:
            out["_truncated"] = True
            break
        i = i2
        if wt == 3:
            try:
                i = _skip_group(buf, i)
            except ValueError:
                out["_truncated"] = True
                break
            continue
        if wt == 4:
            out["_truncated"] = True
            break
        try:
            raw, i = _read_value(buf, i, wt)
        except ValueError:
            out["_truncated"] = True
            break
        field = by_no.get(fno)
        if not field:
            _store_unknown(out, fno, wt, raw)
            continue
        try:
            _interpret(out, field, wt, raw, depth)
        except ValueError:
            _store_unknown(out, fno, wt, raw)
    return out


def _split_envelopes(raw: bytes):
    """Trames Connect/gRPC : [flag][len BE 32][payload]. None si ce n'est pas framé."""
    if len(raw) < 5:
        return None
    i = 0
    n = len(raw)
    frames = []
    while i + 5 <= n:
        flag = raw[i]
        ln = int.from_bytes(raw[i + 1:i + 5], "big")
        if i + 5 + ln > n:
            return None
        frames.append((flag, raw[i + 5:i + 5 + ln]))
        i += 5 + ln
    if i != n or not frames:
        return None
    return frames


def _gunzip(payload: bytes) -> bytes:
    try:
        return gzip.decompress(payload)
    except Exception:
        return payload


def decode_protobuf_http_body(raw: bytes, content_type: str, method: str | None, side: str | None):
    """Corps HTTP déjà décompressé (Content-Encoding). Toujours un objet JSON-able."""
    if not raw:
        return {}
    ct = (content_type or "").lower()
    type_name = message_name_for(method, side)
    framed = None
    if "grpc" in ct or "connect" in ct:
        framed = _split_envelopes(raw)
    if framed:
        out = []
        for flag, payload in framed:
            if flag & 0x01:
                payload = _gunzip(payload)
            if flag & 0x02:
                try:
                    out.append({"_trailer": json.loads(payload.decode("utf-8"))})
                except Exception:
                    out.append({"_trailer_hex": payload.hex()})
                continue
            out.append(decode_message(payload, type_name))
        return out[0] if len(out) == 1 else out
    return decode_message(raw, type_name)
