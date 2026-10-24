#!/usr/bin/env python3
"""gen-cartographie — catalogue EXHAUSTIF des services/méthodes RPC du host,
annoté avec le balayage lecture.

Le catalogue est (re)généré en scannant TOUS les .cjs/.js/.mjs de /opt/sand via
`lib_extract.extract()` : tous les services Connect/gRPC, tous les packages
(aiserver.v1, agent.v1, origin.v1, orbit.v1, …), toutes les méthodes. Il n'est
PAS limité au sous-ensemble de 460 méthodes audité.

Sorties :
  - services-methodes.json   (machine : par service -> méthodes annotées + stats)
  - services-methodes.md     (tableaux lisibles par service + notes IDOR)
  - data/catalogue-complet.json  (instantané plat [fqn, method, kind] du scan)

Entrées d'annotation (dans data/) :
  data/rpc-sweep-log.tsv     service  method  kind  http  outcome   (balayage lecture)
  data/all-rpc-methods.json  triplets [fqn, method, kind] du sous-ensemble AUDITÉ

Méthodo : lecture seule. Seules les méthodes à préfixe de lecture ont été
APPELÉES ; les mutations sont cataloguées mais NON appelées.
"""
import json, os, collections, datetime
import lib_extract

DIR = os.path.dirname(os.path.abspath(__file__))
SWEEP = os.path.join(DIR, "data", "rpc-sweep-log.tsv")
AUDITED = os.path.join(DIR, "data", "all-rpc-methods.json")
CATALOGUE_OUT = os.path.join(DIR, "data", "catalogue-complet.json")

READ_PREFIX = ("Get", "List", "Watch", "Check", "Validate", "Resolve", "Poll", "Query",
               "Count", "Health", "Server", "Describe", "Fetch", "Read", "Search",
               "Available", "Is", "Has", "Lookup", "Export", "Preview", "Estimate")

IDOR_WATCH = {
    ("AgentStoreService", "ListAgentStoreDirectory"),
    ("AgentStoreService", "ReadAgentStoreFile"),
    ("BackgroundComposerService", "GetBackgroundComposerInfo"),
    ("GrokBotService", "GetSandBoxUpgradeSchedule"),
    ("GrokBotService", "ListGrokBotTranscriptEntries"),
    ("GrokBotService", "ReadGrokBotAgentAttachmentChunk"),
    ("GrokBotService", "ReadGrokBotGroupChatCanvas"),
    ("GrokBotService", "ReadGrokBotMcpAppDocument"),
    ("GrokBotService", "ResolveGrokBotAttachmentReads"),
}
IDOR_TESTED_SAFE = {
    ("GrokBotService", "ListGrokBotTranscriptEntries"),
    ("GrokBotService", "ResolveGrokBotAttachmentReads"),
    ("GrokBotService", "SendGrokBotUserMessage"),
}


def short(svc):
    return svc.split(".")[-1]


def verdict(name, outcome):
    if outcome is None:
        return "mutation-non-testee" if not name.startswith(READ_PREFIX) else "non-balayee"
    o = outcome
    if o.startswith("SKIP"):
        return "mutation-non-testee"
    if o.startswith("OK"):
        return "OK_data"
    low = o.lower()
    if "not_logged_in" in low or "unauthenticated" in low:
        return "unauthenticated"
    if "invalid_argument" in low or o.startswith("http400") or "exactly one of" in low or "required" in low:
        return "passe-auth-exige-id"
    if "not_found" in low:
        return "passe-auth-exige-id"
    if "failed_precondition" in low:
        return "passe-auth-failed-precondition"
    if "permission" in low:
        return "passe-auth-refuse-permission"
    if o.startswith("exc:"):
        return "erreur-balayage"
    if "unimplemented" in low or "deprecated" in low or "outdated" in low:
        return "indisponible"
    return o


def load_sweep():
    sweep = {}
    if not os.path.exists(SWEEP):
        return sweep

    def rank(o):
        if o is None: return -1
        if o.startswith("exc:"): return 0
        if o.startswith("OK"): return 3
        return 2

    with open(SWEEP) as f:
        f.readline()
        for line in f:
            parts = line.rstrip("\n").split("\t")
            if len(parts) < 5:
                continue
            s, m, k, http, outcome = parts[:5]
            key = (s, m)
            prev = sweep.get(key)
            if prev is None or rank(outcome) >= rank(prev[1]):
                sweep[key] = (http, outcome)
    return sweep


def load_audited():
    if not os.path.exists(AUDITED):
        return set()
    out = set()
    for fqn, name, kind in json.load(open(AUDITED)):
        out.add((short(fqn), name))
    return out


def build():
    services, files = lib_extract.extract()
    sweep = load_sweep()
    audited = load_audited()

    # instantané plat reproductible du scan
    flat = []
    for fqn in services:
        for m, meta in services[fqn].items():
            flat.append([fqn, m, meta["kind"]])
    flat.sort()
    json.dump({"genere_le": datetime.datetime.now().astimezone().isoformat(timespec="seconds"),
               "fichiers_scannes": files,
               "services": len(services), "methodes": len(flat),
               "catalogue": flat},
              open(CATALOGUE_OUT, "w"), indent=0, ensure_ascii=False)

    by_svc = collections.OrderedDict()
    for fqn in services:
        s = short(fqn)
        for name, meta in services[fqn].items():
            http, outcome = sweep.get((s, name), (None, None))
            by_svc.setdefault(fqn, []).append({
                "service": s,
                "fqn": fqn,
                "method": name,
                "kind": meta["kind"],
                "I": meta["I"],
                "O": meta["O"],
                "http": http,
                "outcome_brut": outcome,
                "verdict": verdict(name, outcome),
                "lecture": name.startswith(READ_PREFIX),
                "audite": (s, name) in audited,
                "idor_watch": (s, name) in IDOR_WATCH,
                "idor_teste_non_vulnerable": (s, name) in IDOR_TESTED_SAFE,
            })
    for fqn in by_svc:
        by_svc[fqn].sort(key=lambda e: e["method"])
    return by_svc, files


def write_json(by_svc, files):
    total = sum(len(v) for v in by_svc.values())
    audited = sum(1 for es in by_svc.values() for e in es if e["audite"])
    out = {
        "genere_le": datetime.datetime.now().astimezone().isoformat(timespec="seconds"),
        "source": {
            "catalogue": "scan exhaustif de /opt/sand (lib_extract.extract)",
            "fichiers_scannes": files,
            "balayage_lecture": "data/rpc-sweep-log.tsv",
            "sous_ensemble_audite": "data/all-rpc-methods.json",
        },
        "stats": {
            "services": len(by_svc),
            "methodes": total,
            "methodes_auditees": audited,
            "packages": dict(collections.Counter(f.rsplit(".", 1)[0] for f in by_svc)),
        },
        "services": by_svc,
    }
    json.dump(out, open(os.path.join(DIR, "services-methodes.json"), "w"),
              indent=1, ensure_ascii=False)
    return total, audited


def write_md(by_svc, files):
    total = sum(len(v) for v in by_svc.values())
    audited = sum(1 for es in by_svc.values() for e in es if e["audite"])
    vcount = collections.Counter(e["verdict"] for es in by_svc.values() for e in es)
    pkgcount = collections.Counter(fqn.rsplit(".", 1)[0] for fqn in by_svc)
    kindcount = collections.Counter(e["kind"] for es in by_svc.values() for e in es)
    L = []
    L.append("# Cartographie annotée — services & méthodes RPC du host Sand/Cursor\n")
    L.append(f"_Généré le {datetime.datetime.now().astimezone().strftime('%Y-%m-%d %H:%M %Z')} "
             "par `gen-cartographie.py` — **catalogue exhaustif** (scan de tout `/opt/sand`) "
             "fusionné avec le balayage lecture._\n")
    L.append("Host `api2.cursor.sh` · protocole Connect · jeton **sand_mobile** (scopé compte) · body `{}`.\n")
    L.append("**Exhaustivité** : le catalogue est obtenu en scannant TOUS les `.cjs/.js/.mjs` de "
             "`/opt/sand` et en relevant chaque `typeName: \"<pkg>.<Service>\"` + son bloc `methods`, "
             "tous packages confondus. Fichiers porteurs de définitions :\n")
    for f in files:
        L.append(f"- `{f}`")
    L.append("")
    L.append("**Lecture seule** : seules les méthodes à préfixe de lecture (Get/List/Read/Resolve/Watch/Check/…) "
             "ont été APPELÉES pendant le balayage. Les mutations (Create/Send/Set/Delete/Save/Start/…) sont "
             "cataloguées mais **non appelées**.\n")
    L.append("## Chiffres\n")
    L.append(f"- Services : **{len(by_svc)}**")
    L.append(f"- Méthodes cataloguées (total réel) : **{total}**")
    L.append(f"- Méthodes couvertes par l'audit lecture : **{audited}**")
    L.append("\nPar package :\n")
    L.append("| Package | Services | Méthodes |\n|---|---|---|")
    for pkg in sorted(pkgcount):
        svcs = [fqn for fqn in by_svc if fqn.rsplit(".", 1)[0] == pkg]
        meth = sum(len(by_svc[fqn]) for fqn in svcs)
        L.append(f"| `{pkg}` | {len(svcs)} | {meth} |")
    L.append("")
    L.append("Par kind :\n")
    L.append("| Kind | Nombre |\n|---|---|")
    for k, n in kindcount.most_common():
        L.append(f"| `{k}` | {n} |")
    L.append("")
    L.append("Répartition par verdict d'accès observé (vide = non couvert par le balayage) :\n")
    L.append("| Verdict | Nombre |\n|---|---|")
    for v, n in vcount.most_common():
        L.append(f"| `{v}` | {n} |")
    L.append("")
    L.append("### Légende des verdicts\n")
    L.append("| Verdict | Sens |\n|---|---|")
    L.append("| `OK_data` | Jeton accepté, des données sont renvoyées. |")
    L.append("| `passe-auth-exige-id` | Auth acceptée mais un id/arg requis → **surface IDOR à vérifier**. |")
    L.append("| `passe-auth-failed-precondition` | Auth acceptée mais pré-condition non remplie. |")
    L.append("| `passe-auth-refuse-permission` | Auth acceptée mais permission refusée. |")
    L.append("| `unauthenticated` | Jeton rejeté (not-logged-in / unauthenticated). |")
    L.append("| `indisponible` | Déprécié / non implémenté / client trop ancien. |")
    L.append("| `erreur-balayage` | Exception côté client pendant le balayage. |")
    L.append("| `mutation-non-testee` | Mutation : cataloguée, jamais appelée (lecture seule). |")
    L.append("| `non-balayee` | Méthode de lecture non couverte par ce balayage. |")
    L.append("")
    L.append("## Surface IDOR\n")
    L.append("Méthodes qui **passent l'auth mais exigent un id/chemin fourni par le client** "
             "(fenêtre IDOR potentielle) :\n")
    for (s, m) in sorted(IDOR_WATCH):
        safe = " — **testé activement : NON vulnérable (scopé owner)**" if (s, m) in IDOR_TESTED_SAFE else ""
        L.append(f"- `{s}.{m}`{safe}")
    L.append("")
    L.append("Verdict IDOR global (tests actifs avec un agentId tiers fourni par l'opérateur) : "
             "`GrokBotService` **non exploitable** en lecture ET en écriture "
             "(transcript vide, attachment 404, write `REFUSED`). Résolution store-only, pas de path traversal. "
             "Voir [audit-2026-10-24.md](audit-2026-10-24.md).\n")
    L.append("## Tableaux par service\n")
    for fqn in sorted(by_svc):
        es = by_svc[fqn]
        vc = collections.Counter(e["verdict"] for e in es)
        summ = ", ".join(f"{k}:{v}" for k, v in vc.most_common())
        L.append(f"### `{fqn}`  ({len(es)} méthodes)\n")
        L.append(f"_{summ}_\n")
        L.append("| Méthode | Kind | HTTP | Verdict | Audité | Note |\n|---|---|---|---|---|---|")
        for e in es:
            note = []
            if e["idor_teste_non_vulnerable"]:
                note.append("IDOR testé → sûr")
            elif e["idor_watch"]:
                note.append("IDOR à surveiller")
            L.append(f"| `{e['method']}` | {e['kind']} | {e['http'] or '—'} | "
                     f"`{e['verdict']}` | {'✓' if e['audite'] else '—'} | {', '.join(note)} |")
        L.append("")
    with open(os.path.join(DIR, "services-methodes.md"), "w") as f:
        f.write("\n".join(L) + "\n")


def main():
    by_svc, files = build()
    total, audited = write_json(by_svc, files)
    write_md(by_svc, files)
    print(f"OK — {len(by_svc)} services, {total} méthodes ({audited} auditées) "
          "→ services-methodes.json + .md + data/catalogue-complet.json")


if __name__ == "__main__":
    main()
