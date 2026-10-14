#!/usr/bin/env python3
"""Infère un JSON Schema depuis une réponse brute d'outil, puis SUPPRIME la réponse.
Usage : infer.py <ns> <tool> <fichier_brut> [nom=chemin.pointé ...]  -> inbox/_inferred.json (local) ; aucune valeur conservée."""
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
def get(o, path):
    for k in path.split("."):
        o = o[int(k)] if isinstance(o, list) else o[k]
    return o
try:
    obj = parse(txt); sch = infer(obj); kind = "json"
    # extractions name=chemin.pointé -> fixtures/values.local.json (local, gitignoré) pour chaîner des fixtures
    ex = dict(a.split("=", 1) for a in sys.argv[4:])
    if ex:
        vp = os.path.join(DIR, "fixtures", "values.local.json")
        vals = json.load(open(vp)) if os.path.exists(vp) else {}
        for name, path in ex.items():
            try: vals[name] = get(obj, path); print(f"  {name}={vals[name]}")
            except Exception: print(f"  {name}: introuvable")
        json.dump(vals, open(vp, "w"), indent=1)
except Exception:
    sch = {"type": "string", "contentMediaType": "text/plain"}; kind = "text"
err = (kind == "json" and isinstance(obj, dict) and ("retry_after_seconds" in obj or "error" in obj or obj.get("status", 200) >= 400 if isinstance(obj.get("status", 200), int) else "error" in obj)) \
      or (kind == "text" and re.match(r"(Error|.{0,80}not found|.{0,80}\(HTTP [45]\d\d\))", txt, re.I | re.S) is not None)
os.remove(raw)
if err:
    print(f"{ns}/{tool}: ERROR (réponse d'erreur, schéma non conservé)"); sys.exit(3)   # la réponse brute ne survit pas à l'inférence
p = os.path.join(DIR, "inbox", "_inferred.json")
d = json.load(open(p)) if os.path.exists(p) else {}
d.setdefault(ns, {})[tool] = {"source": "inferred", "inferredAt": datetime.datetime.now().astimezone().isoformat(timespec="seconds"),
                             "responseKind": kind, "schema": sch}
json.dump(d, open(p, "w"), indent=1, sort_keys=True)
empty = kind == "json" and (obj == [] or obj == {} )
print(f"{ns}/{tool}: {kind}{' (vide)' if empty else ''}")
