#!/usr/bin/env python3
"""Infère un JSON Schema depuis une réponse brute d'outil, puis SUPPRIME la réponse.
Usage : infer.py <ns> <tool> <fichier_brut>  -> inbox/_inferred.json (local) ; aucune valeur conservée."""
import datetime, json, os, re, sys
DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
def merge(a, b):
    if a == b: return a
    ta, tb = a.get("type"), b.get("type")
    if ta == tb == "object":
        keys = set(a["properties"]) | set(b["properties"])
        req = sorted(set(a.get("required", [])) & set(b.get("required", [])))
        props = {}
        for k in keys:
            pa, pb = a["properties"].get(k), b["properties"].get(k)
            props[k] = merge(pa, pb) if pa and pb else (pa or pb)
        return {"type": "object", "properties": props, "required": req}
    if ta == tb == "array":
        return {"type": "array", "items": merge(a.get("items", {}), b.get("items", {})) if a.get("items") and b.get("items") else (a.get("items") or b.get("items") or {})}
    types = sorted({*(ta if isinstance(ta, list) else [ta]), *(tb if isinstance(tb, list) else [tb])} - {None})
    return {"type": types}
def infer(v):
    if v is None: return {"type": "null"}
    if isinstance(v, bool): return {"type": "boolean"}
    if isinstance(v, int): return {"type": "integer"}
    if isinstance(v, float): return {"type": "number"}
    if isinstance(v, str): return {"type": "string"}
    if isinstance(v, list):
        it = None
        for x in v: it = infer(x) if it is None else merge(it, infer(x))
        return {"type": "array", "items": it or {}}
    return {"type": "object", "properties": {k: infer(x) for k, x in v.items()}, "required": sorted(v)}
ns, tool, raw = sys.argv[1:4]
txt = open(raw).read()
txt = re.sub(r"^\s*</?cursor_untrusted_data[^>]*>\s*$", "", txt, flags=re.M).strip()
def parse(t):
    try: return json.loads(t)
    except Exception: pass
    for o, c in (("{", "}"), ("[", "]")):
        i, j = t.find(o), t.rfind(c)
        if i != -1 and j > i:
            try: return json.loads(t[i:j + 1])
            except Exception: pass
    raise ValueError
try:
    sch = infer(parse(txt)); kind = "json"
except Exception:
    sch = {"type": "string", "contentMediaType": "text/plain"}; kind = "text"
os.remove(raw)   # la réponse brute ne survit pas à l'inférence
p = os.path.join(DIR, "inbox", "_inferred.json")
d = json.load(open(p)) if os.path.exists(p) else {}
d.setdefault(ns, {})[tool] = {"source": "inferred", "inferredAt": datetime.datetime.now().astimezone().isoformat(timespec="seconds"),
                             "responseKind": kind, "schema": sch}
json.dump(d, open(p, "w"), indent=1, sort_keys=True)
print(f"{ns}/{tool}: {kind}")
