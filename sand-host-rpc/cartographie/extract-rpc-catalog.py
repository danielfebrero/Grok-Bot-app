#!/usr/bin/env python3
"""Extrait un catalogue protobuf offline depuis host-main.cjs (ou .gz).

Les messages générés par le runtime compact (`static $()`) portent des champs
`numéro nom type` : type scalaire = numéro protobuf-es (1 double, 9 string, …),
type imbriqué = `#N` (index dans les références du descripteur). Ce script
résout `#N` en `ref#<type complet>` et associe chaque méthode Connect
`/pkg.Service/Method` à ses messages requête/réponse.

Le fichier produit ne contient aucun secret : uniquement des noms de champs.
"""
from __future__ import annotations

import argparse
import gzip
import hashlib
import json
import os
import re
import sys

SCALAR_BY_NUM = {
    1: "double",
    2: "float",
    3: "int64",
    4: "uint64",
    5: "int32",
    6: "fixed64",
    7: "fixed32",
    8: "bool",
    9: "string",
    12: "bytes",
    13: "uint32",
    15: "sfixed32",
    16: "sfixed64",
    17: "sint32",
    18: "sint64",
}

# Types bien connus @bufbuild/protobuf, référencés par identifiant JS
# (Timestamp, Struct, …) sans descripteur compact.
WKT_MESSAGES = {
    "google.protobuf.Timestamp": [
        {"no": 1, "name": "seconds", "type": "int64"},
        {"no": 2, "name": "nanos", "type": "int32"},
    ],
    "google.protobuf.Duration": [
        {"no": 1, "name": "seconds", "type": "int64"},
        {"no": 2, "name": "nanos", "type": "int32"},
    ],
    "google.protobuf.Empty": [],
    "google.protobuf.Any": [
        {"no": 1, "name": "type_url", "type": "string"},
        {"no": 2, "name": "value", "type": "bytes"},
    ],
    "google.protobuf.FieldMask": [
        {"no": 1, "name": "paths", "type": "string", "repeated": True},
    ],
    "google.protobuf.Struct": [
        {"no": 1, "name": "fields", "type": "map", "key": "string", "value": "ref#google.protobuf.Value"},
    ],
    "google.protobuf.Value": [
        {"no": 1, "name": "null_value", "type": "ref#google.protobuf.NullValue", "oneof": "kind"},
        {"no": 2, "name": "number_value", "type": "double", "oneof": "kind"},
        {"no": 3, "name": "string_value", "type": "string", "oneof": "kind"},
        {"no": 4, "name": "bool_value", "type": "bool", "oneof": "kind"},
        {"no": 5, "name": "struct_value", "type": "ref#google.protobuf.Struct", "oneof": "kind"},
        {"no": 6, "name": "list_value", "type": "ref#google.protobuf.ListValue", "oneof": "kind"},
    ],
    "google.protobuf.ListValue": [
        {"no": 1, "name": "values", "type": "ref#google.protobuf.Value", "repeated": True},
    ],
    "google.protobuf.DoubleValue": [{"no": 1, "name": "value", "type": "double"}],
    "google.protobuf.FloatValue": [{"no": 1, "name": "value", "type": "float"}],
    "google.protobuf.Int64Value": [{"no": 1, "name": "value", "type": "int64"}],
    "google.protobuf.UInt64Value": [{"no": 1, "name": "value", "type": "uint64"}],
    "google.protobuf.Int32Value": [{"no": 1, "name": "value", "type": "int32"}],
    "google.protobuf.UInt32Value": [{"no": 1, "name": "value", "type": "uint32"}],
    "google.protobuf.BoolValue": [{"no": 1, "name": "value", "type": "bool"}],
    "google.protobuf.StringValue": [{"no": 1, "name": "value", "type": "string"}],
    "google.protobuf.BytesValue": [{"no": 1, "name": "value", "type": "bytes"}],
}
WKT_ENUMS = {
    "google.protobuf.NullValue": {"0": "NULL_VALUE"},
}
WKT_SYMBOLS = {
    "Timestamp": "google.protobuf.Timestamp",
    "Duration": "google.protobuf.Duration",
    "Empty": "google.protobuf.Empty",
    "Any": "google.protobuf.Any",
    "FieldMask": "google.protobuf.FieldMask",
    "Struct": "google.protobuf.Struct",
    "Value": "google.protobuf.Value",
    "ListValue": "google.protobuf.ListValue",
    "NullValue": "google.protobuf.NullValue",
    "DoubleValue": "google.protobuf.DoubleValue",
    "FloatValue": "google.protobuf.FloatValue",
    "Int64Value": "google.protobuf.Int64Value",
    "UInt64Value": "google.protobuf.UInt64Value",
    "Int32Value": "google.protobuf.Int32Value",
    "UInt32Value": "google.protobuf.UInt32Value",
    "BoolValue": "google.protobuf.BoolValue",
    "StringValue": "google.protobuf.StringValue",
    "BytesValue": "google.protobuf.BytesValue",
}

METHOD_RE = re.compile(
    r'name:\s*"([A-Za-z0-9_]+)"\s*,\s*I:\s*([A-Za-z0-9_$.]+)\s*,\s*O:\s*([A-Za-z0-9_$.]+)'
    r'\s*,\s*kind:\s*MethodKind\.([A-Za-z]+)'
)


def load_text(path: str) -> tuple[str, str]:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        raw = f.read()
    h.update(raw)
    if path.endswith(".gz") or raw[:2] == b"\x1f\x8b":
        raw = gzip.decompress(raw)
    return raw.decode("utf-8", "replace"), h.hexdigest()


def skip_ws(s: str, i: int) -> int:
    while i < len(s) and s[i] in " \t\r\n":
        i += 1
    return i


def parse_js_value(s: str, i: int):
    i = skip_ws(s, i)
    if i >= len(s):
        raise ValueError("fin de texte")
    c = s[i]
    if c == '"':
        i += 1
        buf = []
        while i < len(s):
            if s[i] == "\\":
                buf.append(s[i + 1] if i + 1 < len(s) else "")
                i += 2
                continue
            if s[i] == '"':
                return "".join(buf), i + 1
            buf.append(s[i])
            i += 1
        raise ValueError("chaîne non fermée")
    if c == "[":
        i += 1
        arr = []
        while True:
            i = skip_ws(s, i)
            if i < len(s) and s[i] == "]":
                return arr, i + 1
            val, i = parse_js_value(s, i)
            arr.append(val)
            i = skip_ws(s, i)
            if i < len(s) and s[i] == ",":
                i += 1
                continue
            if i < len(s) and s[i] == "]":
                return arr, i + 1
            raise ValueError(f"tableau mal formé près de {s[i:i+40]!r}")
    if c == "-" or c.isdigit():
        j = i + 1
        while j < len(s) and s[j].isdigit():
            j += 1
        return int(s[i:j]), j
    j = i
    while j < len(s) and (s[j].isalnum() or s[j] in "_$."):
        j += 1
    if j == i:
        raise ValueError(f"jeton inattendu {s[i:i+40]!r}")
    return s[i:j], j


def split_call_args(s: str, open_paren: int) -> tuple[list[str], int]:
    """open_paren pointe sur '('."""
    i = open_paren + 1
    args: list[str] = []
    start = i
    depth = 0
    in_str = None
    while i < len(s):
        c = s[i]
        if in_str:
            if c == "\\":
                i += 2
                continue
            if c == in_str:
                in_str = None
            i += 1
            continue
        if c in "\"'`":
            in_str = c
            i += 1
            continue
        if c in "([{":
            depth += 1
        elif c in ")]}":
            if depth == 0 and c == ")":
                args.append(s[start:i].strip())
                return args, i + 1
            depth -= 1
        elif c == "," and depth == 0:
            args.append(s[start:i].strip())
            start = i + 1
        i += 1
    raise ValueError("appel non fermé")


def _methods_block(data: str, start: int):
    mm = data.find("methods:", start)
    if mm < 0 or mm - start > 80:
        return None
    b = data.find("{", mm)
    if b < 0:
        return None
    i = b
    depth = 0
    n = len(data)
    in_str = None
    while i < n:
        c = data[i]
        if in_str:
            if c == "\\":
                i += 2
                continue
            if c == in_str:
                in_str = None
            i += 1
            continue
        if c in "\"'`":
            in_str = c
            i += 1
            continue
        if c == "/" and i + 1 < n and data[i + 1] == "/":
            j = data.find("\n", i)
            i = n if j < 0 else j + 1
            continue
        if c == "/" and i + 1 < n and data[i + 1] == "*":
            j = data.find("*/", i)
            i = n if j < 0 else j + 2
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                return data[b:i + 1]
        i += 1
    return None


def enum_value_prefix(enum_name: str) -> str:
    """Même transformation que compact.js `enumValuePrefix` (toUpperCase)."""
    simple = enum_name.split(".")[-1]
    if not simple:
        return "_"
    rest = re.sub(r"[A-Z]", lambda m: "_" + m.group(0), simple[1:])
    return (simple[0] + rest).upper() + "_"


def parse_field(entry: str, refs: list) -> dict | None:
    toks = entry.split(" ")
    if len(toks) < 3:
        return None
    try:
        no = int(toks[0])
    except ValueError:
        return None
    name = toks[1]
    type_tok = toks[2]
    mod = toks[3] if len(toks) > 3 else None
    if type_tok and type_tok[-1] in "*?":
        mod = type_tok[-1]
        type_tok = type_tok[:-1]
    field: dict = {"no": no, "name": name}
    if mod == "*":
        field["repeated"] = True
    elif mod == "?":
        field["optional"] = True
    elif mod:
        field["oneof"] = mod
    comma = type_tok.find(",")
    if comma != -1:
        key_n = int(type_tok[:comma])
        val_tok = type_tok[comma + 1:]
        field["type"] = "map"
        field["key"] = SCALAR_BY_NUM.get(key_n, f"scalar#{key_n}")
        field["value"] = resolve_type_token(val_tok, refs)
        return field
    field["type"] = resolve_type_token(type_tok, refs)
    return field


def resolve_type_token(token: str, refs: list) -> str:
    if token.startswith("#"):
        try:
            idx = int(token[1:])
        except ValueError:
            return "ref#missing"
        if idx < 0 or idx >= len(refs):
            return "ref#missing"
        return f"ref#{refs[idx]}"
    try:
        n = int(token)
    except ValueError:
        return token
    return SCALAR_BY_NUM.get(n, f"scalar#{n}")


def extract(data: str) -> dict:
    pkg_vars = dict(re.findall(r"(__protoPackage\d*) = \"([^\"]*)\"", data))
    bases: dict[str, str] = {}
    for m in re.finditer(r"(__protoMessage\d+) = class extends CompactMessage \{", data):
        chunk = data[m.end():m.end() + 300]
        pm = re.search(r"return (__protoPackage\d*)", chunk)
        if pm:
            bases[m.group(1)] = pkg_vars.get(pm.group(1), "")

    symbols = dict(WKT_SYMBOLS)
    messages: dict[str, list] = {k: [dict(f) for f in v] for k, v in WKT_MESSAGES.items()}
    enums = {k: dict(v) for k, v in WKT_ENUMS.items()}

    # enums
    idx = 0
    enum_fail = 0
    while True:
        j = data.find("enumType(", idx)
        if j < 0:
            break
        if data[max(0, j - 15):j].rstrip().endswith("function"):
            idx = j + 9
            continue
        try:
            args, nxt = split_call_args(data, j + len("enumType") )
        except ValueError:
            enum_fail += 1
            idx = j + 9
            continue
        idx = nxt
        # args: runtime, packageVar, "name", values, optional mode
        if len(args) < 4:
            enum_fail += 1
            continue
        pkg_var = args[1].strip()
        try:
            enum_name, _ = parse_js_value(args[2], 0)
            values, _ = parse_js_value(args[3], 0)
        except ValueError:
            enum_fail += 1
            continue
        if not isinstance(enum_name, str) or not isinstance(values, list):
            enum_fail += 1
            continue
        prefix = bases_pkg = pkg_vars.get(pkg_var, "")
        full = prefix + enum_name
        mode = args[4].strip() if len(args) > 4 else ""
        table = {}
        pfx = enum_value_prefix(enum_name) if mode == "1" else ""
        for item in values:
            if not isinstance(item, list) or len(item) < 2:
                continue
            num, local = item[0], item[1]
            proto_name = (pfx + str(local)) if pfx else str(local)
            table[str(int(num))] = proto_name
        enums[full] = table
        # identifiant JS : le nom simple assigné juste avant l'appel
        pre = data[max(0, j - 400):j]
        am = re.search(r"([A-Za-z0-9_]+)\s*=\s*(?:/\*[^*]*\*/\s*)?$", pre)
        if am:
            symbols[am.group(1)] = full

    # messages compact
    msg_fail = 0
    unresolved_refs = 0
    raw_messages = []
    for m in re.finditer(r"static \$\(\) \{\s*return \[", data):
        try:
            arr, _ = parse_js_value(data, m.end() - 1)
        except ValueError:
            msg_fail += 1
            continue
        if not arr or not isinstance(arr[0], str):
            msg_fail += 1
            continue
        back = data.rfind("extends __protoMessage", 0, m.start())
        if back < 0:
            msg_fail += 1
            continue
        window = data[max(0, back - 800):back + 40]
        found = list(re.finditer(
            r"([A-Za-z0-9_]+) = class _([A-Za-z0-9_]+) extends (__protoMessage\d+)",
            window,
        ))
        # La fenêtre peut contenir la classe précédente : on prend la plus proche.
        cm = found[-1] if found else None
        if not cm:
            msg_fail += 1
            continue
        ident, _cls, base = cm.group(1), cm.group(2), cm.group(3)
        pkg = bases.get(base, "")
        if not pkg:
            msg_fail += 1
            continue
        desc = arr[0]
        bar = desc.find("|")
        short = desc if bar < 0 else desc[:bar]
        full = pkg + short
        refs = [str(x) for x in arr[1:]]
        fields = []
        if bar >= 0:
            for entry in desc[bar + 1:].split("|"):
                if not entry:
                    continue
                parsed = parse_field(entry, refs)
                if parsed:
                    fields.append(parsed)
        raw_messages.append((ident, full, fields))
        symbols.setdefault(ident, full)
        # Le descripteur référence parfois le nom de classe `_Foo`, pas l'export.
        symbols.setdefault("_" + _cls, full)

    # résoudre ref#Ident → ref#type.complet (2e passe, symbols complet)
    for ident, full, fields in raw_messages:
        resolved = []
        for f in fields:
            f = dict(f)
            f["type"] = _resolve_ref(f.get("type", ""), symbols)
            if f.get("type") == "ref#missing":
                unresolved_refs += 1
            if "value" in f:
                f["value"] = _resolve_ref(f["value"], symbols)
                if f["value"] == "ref#missing":
                    unresolved_refs += 1
            resolved.append(f)
        messages[full] = resolved

    methods = {}
    for tm in re.finditer(r'typeName:\s*"([^"]+)"', data):
        block = _methods_block(data, tm.end())
        if not block:
            continue
        svc = tm.group(1)
        for me in METHOD_RE.finditer(block):
            name, i_ty, o_ty, kind = me.groups()
            methods[f"/{svc}/{name}"] = {
                "request": symbols.get(i_ty, symbols.get(i_ty.split(".")[-1], i_ty)),
                "response": symbols.get(o_ty, symbols.get(o_ty.split(".")[-1], o_ty)),
                "kind": kind,
            }

    return {
        "version": 1,
        "note": (
            "Champs extraits des descripteurs compact host-main. "
            "Scalaires = noms protobuf-es. Imbriqué = ref#TypeComplet "
            "(le #N du bundle est résolu). Tags inconnus au décodage : _unknown_<n>."
        ),
        "messages": messages,
        "enums": enums,
        "methods": methods,
        "_stats": {
            "messages": len(messages),
            "enums": len(enums),
            "methods": len(methods),
            "enum_parse_fail": enum_fail,
            "message_parse_fail": msg_fail,
            "unresolved_refs": unresolved_refs,
        },
    }


def _resolve_ref(token: str, symbols: dict) -> str:
    if not isinstance(token, str) or not token.startswith("ref#"):
        return token
    ident = token[4:]
    full = symbols.get(ident)
    if full:
        return "ref#" + full
    if ident.startswith("google.protobuf."):
        return token
    return "ref#missing"


def main(argv: list[str] | None = None) -> int:
    here = os.path.dirname(os.path.abspath(__file__))
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("bundle", help="host-main.cjs ou host-main.cjs.gz")
    ap.add_argument(
        "-o",
        default=os.path.join(here, "data", "rpc_catalog.json"),
        help="catalogue de sortie",
    )
    args = ap.parse_args(argv)
    text, digest = load_text(args.bundle)
    cat = extract(text)
    stats = cat.pop("_stats")
    cat["bundle_sha256"] = digest
    os.makedirs(os.path.dirname(os.path.abspath(args.o)), exist_ok=True)
    with open(args.o, "w", encoding="utf-8") as f:
        json.dump(cat, f, ensure_ascii=False, separators=(",", ":"), sort_keys=True)
        f.write("\n")
    print(
        f"écrit {args.o} messages={stats['messages']} enums={stats['enums']} "
        f"methods={stats['methods']} enum_fail={stats['enum_parse_fail']} "
        f"msg_fail={stats['message_parse_fail']} unresolved={stats['unresolved_refs']}",
        file=sys.stderr,
    )
    probe = "/aiserver.v1.DashboardService/GetManagedSkills"
    print("probe", probe, cat["methods"].get(probe), file=sys.stderr)
    ok = (
        stats["message_parse_fail"] == 0
        and stats["unresolved_refs"] == 0
        and probe in cat["methods"]
    )
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
