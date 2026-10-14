#!/usr/bin/env python3
"""mcp-schema-monitor : normalise inbox/ -> snapshots/<ts>/, compare au précédent, écrit reports/ + CHANGELOG.md.
Codes : 0 NO_CHANGE / BASELINE ; 10 CHANGED ; 1 erreur."""
import datetime, difflib, hashlib, json, os, re, shutil, subprocess, sys
DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INBOX, SNAPS, REPORTS = (os.path.join(DIR, d) for d in ("inbox", "snapshots", "reports"))
KEEP = ("tool", "title", "description", "inputSchema", "outputSchema", "annotations")  # seuls champs publiés
FIRST_PARTY = {"cursor"}                            # inclus seulement si scan propre
SAFE = re.compile(r"[^A-Za-z0-9._-]")

def dumps(o): return json.dumps(o, indent=2, sort_keys=True, ensure_ascii=False) + "\n"
def sha(s): return hashlib.sha256(s.encode()).hexdigest()

def scan_clean(path):
    r = subprocess.run([os.path.join(DIR, "scan-personal-data.sh"), "--redact", path], capture_output=True, text=True)
    return r.returncode == 0

def load_inbox():
    cat_p = os.path.join(INBOX, "_catalog.json")
    if not os.path.exists(cat_p):
        sys.exit("ERROR: inbox/_catalog.json absent (sortie de GetDynamicTools pattern='.')")
    cat = json.load(open(cat_p))
    nss = {}
    for ns, v in cat.get("namespaces", {}).items():
        nss[ns] = {"status": v.get("status", "?"), "level": "names", "tools": {t: None for t in v.get("tools", [])}}
    full = {}
    for f in sorted(os.listdir(INBOX)):
        if not f.endswith(".json") or f.startswith("_"):
            continue
        raw = json.load(open(os.path.join(INBOX, f)))
        ns = raw.get("namespace") or f[:-5]
        if ns in FIRST_PARTY and not scan_clean(os.path.join(INBOX, f)):
            print(f"WARN: namespace first-party {ns} exclu (scan non propre)", file=sys.stderr)
            continue
        tools = {}
        for t in raw.get("tools", []):
            d = {k: t[k] for k in KEEP if k in t}
            d.setdefault("outputSchema", None)      # explicite même absent
            tools[t["tool"]] = d
        full[ns] = tools
        e = nss.setdefault(ns, {"status": raw.get("namespaceStatus", "?"), "tools": {}})
        e["status"] = raw.get("namespaceStatus", e["status"])
        e["level"] = "full"
        e["tools"] = {n: sha(dumps(d)) for n, d in tools.items()}
        e["output_schema_count"] = sum(1 for d in tools.values() if d.get("outputSchema") is not None)
    for ns in list(nss):
        if ns in FIRST_PARTY and ns not in full:
            del nss[ns]
    for e in nss.values():
        e["tool_count"] = len(e["tools"])
        e.setdefault("output_schema_count", None)   # inconnu au niveau "names"
    return nss, full, cat.get("absent", {})

def snaps():
    return sorted(d for d in os.listdir(SNAPS) if os.path.isfile(os.path.join(SNAPS, d, "index.json")))

def tool_file(snap, ns, t): return os.path.join(SNAPS, snap, SAFE.sub("_", ns), SAFE.sub("_", t) + ".json")

def main():
    os.makedirs(SNAPS, exist_ok=True); os.makedirs(REPORTS, exist_ok=True)
    nss, full, absent = load_inbox()
    ts = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M")
    index = {"snapshot": ts, "namespaces": nss, "absent": absent,
             "totals": {"namespaces": len(nss), "tools": sum(e["tool_count"] for e in nss.values()),
                        "tools_full_schema": sum(len(v) for v in full.values())}}
    prev = snaps()
    prev = prev[-1] if prev else None
    old = json.load(open(os.path.join(SNAPS, prev, "index.json"))) if prev else None
    if old and old["namespaces"] == nss and old.get("absent") == absent:
        print(f"NO_CHANGE last={prev}"); return 0
    if prev == ts:
        shutil.rmtree(os.path.join(SNAPS, ts))
        prev = (snaps() or [None])[-1]
        old = json.load(open(os.path.join(SNAPS, prev, "index.json"))) if prev else None
    tmp = os.path.join(SNAPS, ts + ".partial"); shutil.rmtree(tmp, ignore_errors=True)
    for ns, tools in full.items():
        d = os.path.join(tmp, SAFE.sub("_", ns)); os.makedirs(d, exist_ok=True)
        for t, data in tools.items():
            open(os.path.join(d, SAFE.sub("_", t) + ".json"), "w").write(dumps(data))
    os.makedirs(tmp, exist_ok=True)
    open(os.path.join(tmp, "index.json"), "w").write(dumps(index))
    os.rename(tmp, os.path.join(SNAPS, ts))
    rep = os.path.join(REPORTS, ts + ".md")
    L = [f"# Schémas MCP — {ts}", ""]
    tot = index["totals"]
    if not old:
        L += [f"Baseline : {tot['namespaces']} namespaces, {tot['tools']} tools ({tot['tools_full_schema']} avec schéma complet).", "",
              "| Namespace | Statut | Tools | Niveau | outputSchema |", "|---|---|---|---|---|"]
        L += [f"| {ns} | {e['status']} | {e['tool_count']} | {e['level']} | {osc(e)} |" for ns, e in sorted(nss.items())]
        if absent: L += ["", "Non capturés : " + "; ".join(f"{k} ({v})" for k, v in absent.items())]
        open(rep, "w").write("\n".join(L) + "\n")
        changelog(ts, f"baseline — {tot['namespaces']} namespaces, {tot['tools']} tools", rep)
        print(f"BASELINE snapshot={ts} report={rep}"); return 0
    on = old["namespaces"]; add = rem = mod = 0; body = []
    for ns in sorted(set(on) | set(nss)):
        a, b = on.get(ns), nss.get(ns)
        if a == b: continue
        if not a: body.append(f"## + namespace {ns} ({b['status']}, {b['tool_count']} tools)"); add += b["tool_count"]; continue
        if not b: body.append(f"## − namespace {ns} (retiré, {a['tool_count']} tools)"); rem += a["tool_count"]; continue
        body.append(f"## {ns}")
        if a["status"] != b["status"]: body.append(f"- statut : {a['status']} → {b['status']}")
        if a.get("output_schema_count") != b.get("output_schema_count"):
            body.append(f"- tools avec outputSchema : {a.get('output_schema_count')} → {b.get('output_schema_count')}")
        if a.get("level") != b.get("level"): body.append(f"- niveau de capture : {a.get('level')} → {b.get('level')}")
        for t in sorted(set(b["tools"]) - set(a["tools"])): body.append(f"- **+ {t}**"); add += 1
        for t in sorted(set(a["tools"]) - set(b["tools"])): body.append(f"- **− {t}**"); rem += 1
        for t in sorted(set(a["tools"]) & set(b["tools"])):
            ha, hb = a["tools"][t], b["tools"][t]
            if ha and hb and ha != hb:
                mod += 1
                fa, fb = tool_file(prev, ns, t), tool_file(ts, ns, t)
                d = difflib.unified_diff((open(fa).read() if os.path.exists(fa) else "").splitlines(), open(fb).read().splitlines(),
                                         f"{prev}/{ns}/{t}.json", f"{ts}/{ns}/{t}.json", lineterm="", n=2)
                body += [f"- **~ {t}**", "", "```diff", *d, "```", ""]
    L += [f"Précédent : {prev}", "", "| | Nb |", "|---|---|", f"| Ajoutés | {add} |", f"| Retirés | {rem} |",
          f"| Modifiés | {mod} |", "", "| Namespace | Tools | outputSchema |", "|---|---|---|",
          *[f"| {ns} | {e['tool_count']} | {osc(e)} |" for ns, e in sorted(nss.items())], "", *body]
    open(rep, "w").write("\n".join(L) + "\n")
    changelog(ts, f"+{add}/−{rem}/~{mod} vs {prev}", rep)
    with open(os.path.join(REPORTS, "NOUVEAU-DIFF.flag"), "a") as fl: fl.write(f"{ts} {rep}\n")
    print(f"CHANGED old={prev} new={ts} report={rep} added={add} removed={rem} modified={mod}"); return 10

def osc(e):
    c = e.get("output_schema_count")
    return "n/a (names)" if c is None else f"{c}/{e['tool_count']}"

def changelog(ts, summary, rep):
    p = os.path.join(DIR, "CHANGELOG.md")
    old = open(p).read() if os.path.exists(p) else "# CHANGELOG — schémas MCP\n\n"
    head, _, rest = old.partition("\n\n")
    open(p, "w").write(f"{head}\n\n- **{ts}** : {summary} — [rapport](reports/{os.path.basename(rep)})\n{rest}")

if __name__ == "__main__":
    sys.exit(main())
