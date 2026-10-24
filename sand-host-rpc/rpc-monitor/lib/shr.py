#!/usr/bin/env python3
"""shr — sand-host-rpc monitor : snapshots + diffs + changelog des MÉTHODES RPC.

Extrait l'ensemble des déclarations RPC (`@generated from rpc <FQN>.<Method>` suivies de
`{ name, I, O, kind: MethodKind.* }`) depuis /opt/sand/sand-host/*.cjs, puis suit leur dérive.

Sous-commandes :
  snapshot [--label TXT]   crée snapshots/<YYYY-MM-DD_HH-MM>/ (manifest + fingerprint sha256)
  diff [OLD NEW]           rapport reports/<new>.md (+ reports/<new>-diffs/*.diff) ; défaut = 2 plus récents
  check                    empreinte rapide vs dernier snapshot ; si changement -> snapshot + diff +
                           entrée horodatée commentée dans CHANGELOG.md ; sinon NO_CHANGE
                           codes : 0 no-change/baseline · 10 changed · 1 err · 3 lock (géré par check.sh)
  fingerprint              imprime l'empreinte rapide courante
"""
import hashlib, json, os, re, shutil, sys
from datetime import datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SNAP_DIR = os.path.join(ROOT, "snapshots")
REPORT_DIR = os.path.join(ROOT, "reports")
LOG_DIR = os.path.join(ROOT, "logs")
CHANGELOG = os.path.join(ROOT, "CHANGELOG.md")

# ---------------------------------------------------------------------------
# Extraction EXHAUSTIVE : scanne TOUS les .cjs/.js/.mjs de /opt/sand, relève
# chaque `typeName: "<pkg>.<Service>"` + son bloc `methods:{…}` (tous packages :
# aiserver.v1, agent.v1, origin.v1, orbit.v1, …), via un balayeur d'accolades
# conscient des chaînes/commentaires (évite les sur-dépassements sur minifié).
# ---------------------------------------------------------------------------
SCAN_ROOTS = ["/opt/sand"]
SCAN_EXTS = (".cjs", ".js", ".mjs")
TYPENAME_RE = re.compile(r'typeName:\s*"([A-Za-z_][A-Za-z0-9_.]*\.[A-Za-z0-9_]+)"')
METHOD_RE = re.compile(
    r'name:\s*"([A-Za-z0-9_]+)"\s*,\s*I:\s*([A-Za-z0-9_$.]+)\s*,\s*O:\s*([A-Za-z0-9_$.]+)'
    r'\s*,\s*kind:\s*MethodKind\.([A-Za-z]+)'
)
INLINE_DIFF_LINES = 200


def _methods_block(data, start):
    mm = data.find("methods:", start)
    if mm < 0 or mm - start > 80:
        return None
    b = data.find("{", mm)
    if b < 0:
        return None
    i = b; depth = 0; n = len(data)
    while i < n:
        c = data[i]
        if c in '"\'`':
            q = c; i += 1
            while i < n:
                if data[i] == "\\": i += 2; continue
                if data[i] == q: break
                i += 1
            i += 1; continue
        if c == "/" and i + 1 < n and data[i + 1] == "/":
            j = data.find("\n", i); i = n if j < 0 else j + 1; continue
        if c == "/" and i + 1 < n and data[i + 1] == "*":
            j = data.find("*/", i); i = n if j < 0 else j + 2; continue
        if c == "{": depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                return data[b:i + 1]
        i += 1
    return None


def now_local():
    return datetime.now().astimezone()


def log(msg):
    os.makedirs(LOG_DIR, exist_ok=True)
    with open(os.path.join(LOG_DIR, "monitor.log"), "a") as f:
        f.write(f"[{now_local().strftime('%Y-%m-%d %H:%M:%S %Z')}] {msg}\n")


def scan_files():
    files = []
    for root in SCAN_ROOTS:
        for dp, _, fs in os.walk(root):
            for fn in fs:
                if fn.endswith(SCAN_EXTS):
                    files.append(os.path.join(dp, fn))
    return sorted(files)


def extract():
    """Retourne {service_court: {method: {"fqn","I","O","kind","sig","source"}}}."""
    out = {}
    for path in scan_files():
        try:
            data = open(path, "r", errors="replace").read()
        except Exception:
            continue
        if "typeName:" not in data or "MethodKind" not in data:
            continue
        src = os.path.relpath(path, "/opt/sand")
        for tm in TYPENAME_RE.finditer(data):
            fqn = tm.group(1)
            block = _methods_block(data, tm.end())
            if not block:
                continue
            svc = fqn.split(".")[-1]
            for me in METHOD_RE.finditer(block):
                name, I, O, kind = me.groups()
                if name in out.get(svc, {}):
                    continue
                out.setdefault(svc, {})[name] = {
                    "fqn": fqn, "I": I, "O": O, "kind": kind,
                    "sig": f"{I}|{O}|{kind}", "source": src,
                }
    return out


def flatten(methods):
    """{svc:{m:info}} -> dict 'svc.m' -> info (ordonné)."""
    flat = {}
    for svc in sorted(methods):
        for m in sorted(methods[svc]):
            flat[f"{svc}.{m}"] = methods[svc][m]
    return flat


def fingerprint(methods):
    """sha256 global + par service, sur la signature (I|O|kind) de chaque méthode."""
    flat = flatten(methods)
    h = hashlib.sha256()
    for k in sorted(flat):
        h.update((k + "=" + flat[k]["sig"] + "\n").encode())
    per_svc = {}
    for svc in sorted(methods):
        hs = hashlib.sha256()
        for m in sorted(methods[svc]):
            hs.update((m + "=" + methods[svc][m]["sig"] + "\n").encode())
        per_svc[svc] = hs.hexdigest()
    return {
        "global": h.hexdigest(),
        "services": len(methods),
        "methodes": len(flat),
        "par_service": per_svc,
    }


def list_snapshots():
    if not os.path.isdir(SNAP_DIR):
        return []
    return sorted(d for d in os.listdir(SNAP_DIR)
                  if os.path.isdir(os.path.join(SNAP_DIR, d)) and not d.endswith(".partial"))


def take_snapshot(label="", quiet=False):
    methods = extract()
    flat = flatten(methods)
    fp = fingerprint(methods)
    name = now_local().strftime("%Y-%m-%d_%H-%M")
    dst = os.path.join(SNAP_DIR, name)
    if os.path.exists(dst):
        name = name + "-" + now_local().strftime("%S")
        dst = os.path.join(SNAP_DIR, name)
    os.makedirs(dst, exist_ok=True)
    # manifest machine : toutes les méthodes avec signature
    json.dump({"genere_le": now_local().isoformat(timespec="seconds"),
               "sources": sorted({v["source"] for sm in methods.values() for v in sm.values()}),
               "stats": {"services": fp["services"], "methodes": fp["methodes"]},
               "methodes": flat},
              open(os.path.join(dst, "manifest.json"), "w"), indent=1, ensure_ascii=False)
    # tsv lisible
    with open(os.path.join(dst, "methodes.tsv"), "w") as f:
        f.write("service\tmethode\tkind\tI\tO\tsource\n")
        for k in sorted(flat):
            svc, _, mm = k.partition(".")
            e = flat[k]
            f.write(f"{svc}\t{mm}\t{e['kind']}\t{e['I']}\t{e['O']}\t{e['source']}\n")
    json.dump(fp, open(os.path.join(dst, "fingerprint.json"), "w"), indent=1)
    with open(os.path.join(dst, "SNAPSHOT-INFO.txt"), "w") as f:
        f.write(f"name\t{name}\n")
        f.write(f"label\t{label}\n")
        f.write(f"date\t{now_local().strftime('%Y-%m-%d %H:%M:%S %Z')}\n")
        f.write(f"services\t{fp['services']}\n")
        f.write(f"methodes\t{fp['methodes']}\n")
        f.write(f"fingerprint\t{fp['global']}\n")
    log(f"snapshot {name}: {fp['services']} services, {fp['methodes']} méthodes, fp={fp['global'][:12]}")
    if not quiet:
        print(f"snapshot\t{name}")
        print(f"services\t{fp['services']}\nmethodes\t{fp['methodes']}\nfingerprint\t{fp['global']}")
    return name


def load_manifest(snap):
    return json.load(open(os.path.join(SNAP_DIR, snap, "manifest.json")))["methodes"]


def make_diff(old=None, new=None, quiet=False):
    snaps = list_snapshots()
    if old is None or new is None:
        if len(snaps) < 2:
            print(f"Il faut au moins 2 snapshots (trouvés : {len(snaps)}).")
            return None
        old, new = snaps[-2], snaps[-1]
    for s in (old, new):
        if s not in snaps:
            raise SystemExit(f"snapshot inconnu : {s}")
    A, B = load_manifest(old), load_manifest(new)
    added = sorted(set(B) - set(A))
    removed = sorted(set(A) - set(B))
    modified = sorted(k for k in set(A) & set(B) if A[k]["sig"] != B[k]["sig"])

    os.makedirs(REPORT_DIR, exist_ok=True)
    adjacent = snaps.index(new) - snaps.index(old) == 1
    rname = new if adjacent else f"{new}_vs_{old}"
    rpath = os.path.join(REPORT_DIR, rname + ".md")
    ddir = os.path.join(REPORT_DIR, rname + "-diffs")
    if os.path.isdir(ddir):
        shutil.rmtree(ddir)

    L = []
    L.append(f"# Diff méthodes RPC — `{old}` → `{new}`\n")
    L.append(f"Généré le {now_local().strftime('%Y-%m-%d %H:%M %Z')} par rpc-monitor (shr).\n")
    L.append("## Résumé\n")
    L.append(f"| | Nombre |\n|---|---|\n| Ajoutées | {len(added)} |\n"
             f"| Supprimées | {len(removed)} |\n| Modifiées (kind/schéma) | {len(modified)} |\n")

    L.append(f"## Méthodes ajoutées ({len(added)})\n")
    if added:
        L.append("| Méthode | Kind | I → O |\n|---|---|---|")
        for k in added:
            e = B[k]
            L.append(f"| `{k}` | {e['kind']} | `{e['I']}` → `{e['O']}` |")
    else:
        L.append("_Aucune._")
    L.append("")
    L.append(f"## Méthodes supprimées ({len(removed)})\n")
    if removed:
        L.append("| Méthode | Kind | I → O |\n|---|---|---|")
        for k in removed:
            e = A[k]
            L.append(f"| `{k}` | {e['kind']} | `{e['I']}` → `{e['O']}` |")
    else:
        L.append("_Aucune._")
    L.append("")
    L.append(f"## Méthodes modifiées ({len(modified)})\n")
    if modified:
        L.append("| Méthode | Avant (I\\|O\\|kind) | Après (I\\|O\\|kind) |\n|---|---|---|")
        for k in modified:
            L.append(f"| `{k}` | `{A[k]['sig']}` | `{B[k]['sig']}` |")
        os.makedirs(ddir, exist_ok=True)
        for k in modified:
            with open(os.path.join(ddir, k.replace("/", "_") + ".diff"), "w") as f:
                f.write(f"--- {old}:{k}\n+++ {new}:{k}\n")
                f.write(f"- I: {A[k]['I']}\n- O: {A[k]['O']}\n- kind: {A[k]['kind']}\n")
                f.write(f"+ I: {B[k]['I']}\n+ O: {B[k]['O']}\n+ kind: {B[k]['kind']}\n")
    else:
        L.append("_Aucune._")
    L.append("")
    with open(rpath, "w") as f:
        f.write("\n".join(L) + "\n")
    log(f"diff {old} -> {new}: +{len(added)} -{len(removed)} ~{len(modified)} report={rpath}")
    if not quiet:
        print(f"report\t{rpath}")
        print(f"added\t{len(added)}\nremoved\t{len(removed)}\nmodified\t{len(modified)}")
    return rpath, len(added), len(removed), len(modified)


def append_changelog(old, new, res):
    rpath, a, d, mo = res
    A, B = load_manifest(old), load_manifest(new)
    added = sorted(set(B) - set(A))
    removed = sorted(set(A) - set(B))
    modified = sorted(k for k in set(A) & set(B) if A[k]["sig"] != B[k]["sig"])
    ts = now_local().strftime("%Y-%m-%d %H:%M")
    lines = [f"\n## {ts} (Paris) — `{old}` → `{new}`\n"]
    # commentaire en français décrivant le changement
    bits = []
    if added:
        bits.append(f"**{len(added)} méthode(s) ajoutée(s)**")
    if removed:
        bits.append(f"**{len(removed)} supprimée(s)**")
    if modified:
        bits.append(f"**{len(modified)} modifiée(s)** (kind ou type I/O)")
    lines.append("Changement de surface RPC détecté : " + ", ".join(bits) + ".\n")
    if added:
        lines.append("Ajoutées : " + ", ".join(f"`{k}`" for k in added[:25])
                     + (" …" if len(added) > 25 else "") + "\n")
    if removed:
        lines.append("Supprimées : " + ", ".join(f"`{k}`" for k in removed[:25])
                     + (" …" if len(removed) > 25 else "") + "\n")
    if modified:
        lines.append("Modifiées : " + ", ".join(f"`{k}`" for k in modified[:25])
                     + (" …" if len(modified) > 25 else "") + "\n")
    lines.append(f"Rapport : [reports/{os.path.basename(rpath)}](reports/{os.path.basename(rpath)}).\n")
    # insertion en tête (après l'en-tête fixe)
    old_txt = open(CHANGELOG).read() if os.path.exists(CHANGELOG) else ""
    marker = "<!-- NOUVELLES-ENTREES-ICI -->"
    block = "\n".join(lines)
    if marker in old_txt:
        new_txt = old_txt.replace(marker, marker + "\n" + block, 1)
    else:
        new_txt = old_txt + "\n" + block
    open(CHANGELOG, "w").write(new_txt)


def check():
    snaps = list_snapshots()
    if not snaps:
        name = take_snapshot("baseline (auto)", quiet=True)
        print(f"BASELINE snapshot={name}")
        return 0
    last = snaps[-1]
    old_fp = json.load(open(os.path.join(SNAP_DIR, last, "fingerprint.json")))
    cur_fp = fingerprint(extract())
    if cur_fp["global"] == old_fp["global"]:
        log(f"check: no change vs {last}")
        print(f"NO_CHANGE last={last} methodes={cur_fp['methodes']}")
        return 0
    new = take_snapshot("auto: changement de surface RPC détecté", quiet=True)
    res = make_diff(last, new, quiet=True)
    append_changelog(last, new, res)
    log(f"check: CHANGED {last}->{new} +{res[1]} -{res[2]} ~{res[3]}")
    print(f"CHANGED old={last} new={new} report={res[0]} added={res[1]} removed={res[2]} modified={res[3]}")
    return 10


def main(argv):
    if not argv or argv[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    cmd, args = argv[0], argv[1:]
    if cmd == "snapshot":
        label = args[args.index("--label") + 1] if "--label" in args else ""
        take_snapshot(label)
        return 0
    if cmd == "diff":
        pos = [a for a in args if not a.startswith("-")]
        if len(pos) == 1:
            make_diff(pos[0], list_snapshots()[-1])
        elif len(pos) >= 2:
            make_diff(pos[0], pos[1])
        else:
            make_diff()
        return 0
    if cmd == "check":
        return check()
    if cmd == "fingerprint":
        print(json.dumps(fingerprint(extract()), indent=1))
        return 0
    print(__doc__)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
