#!/usr/bin/env bash
# Scan "données perso / secrets" AVANT toute publication. Exit 0 = propre, 1 = findings.
# Usage : scan-personal-data.sh [--redact] [DOSSIER]   (défaut : ce dossier, en respectant .gitignore)
#   --redact : n'affiche QUE chemins + n° de règle + nb d'occurrences (jamais le texte trouvé ni le motif)
# Motifs perso : .pii-denylist (local, gitignoré — jamais publié). Les valeurs des variables
# d'environnement SAND_/CURSOR_/BOX_/GH_ et des fichiers secrets de sand-data sont aussi
# recherchées littéralement (sans jamais être affichées).
set -uo pipefail
DIR="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"
REDACT=0
if [[ "${1:-}" == "--redact" ]]; then REDACT=1; shift; fi
TARGET="${1:-$DIR}"
[[ -e "$TARGET" ]] || { echo "cible introuvable : $TARGET" >&2; exit 2; }
exec python3 - "$DIR" "$TARGET" "$REDACT" <<'PY'
import json, os, re, subprocess, sys
mon, target, redact = sys.argv[1], sys.argv[2], sys.argv[3] == "1"
rg = ["rg", "-z", "-i", "--hidden", "--no-require-git", "-n", "--max-columns", "220", "--max-columns-preview",
      "-g", "!.pii-denylist", "-g", "!.git/"]
if os.path.abspath(target) != os.path.abspath(mon):
    rg.append("--no-ignore")
findings = 0
def show(title, out):
    global findings
    lines = [l for l in out.splitlines() if l.strip()]
    if lines:
        findings += len(lines)
        print(f"\n## {title} ({len(lines)})")
        for l in lines[:40]:
            print("  " + l[:260])
        if len(lines) > 40:
            print(f"  … {len(lines)-40} de plus")
def rgx(args):
    """rg ; en mode redact : -c (chemin:nb) au lieu des lignes."""
    a = list(args)
    if redact:
        a = [x for x in a if x not in ("-n", "--max-columns", "220", "--max-columns-preview")] + ["-c"]
    return subprocess.run(a, capture_output=True, text=True).stdout
deny = os.path.join(mon, ".pii-denylist")
if os.path.exists(deny):
    rule = 0
    for pat in open(deny):
        pat = pat.strip()
        if pat and not pat.startswith("#"):
            rule += 1
            out = rgx(rg + ["-P", "-e", pat, target])
            show(f"denylist règle #{rule}" if redact else f"denylist règle #{rule} /{pat}/", out)
else:
    print("ATTENTION : .pii-denylist absent — scan identités incomplet")
TOKENS = (r"(ghp_[A-Za-z0-9]{30,}|gh[ousr]_[A-Za-z0-9]{30,}|github_pat_[A-Za-z0-9_]{40,}|AKIA[0-9A-Z]{16}"
          r"|xox[baprs]-[A-Za-z0-9-]{10,}|sk_live_[A-Za-z0-9]{20,}|AIza[0-9A-Za-z_-]{35}"
          r"|eyJ[A-Za-z0-9_-]{20,}\.[A-Za-z0-9_-]{20,}\.[A-Za-z0-9_-]{10,}"
          r"|-----BEGIN [A-Z ]*PRIVATE KEY-----(?!\\\\?n?\\?\[?r?s?\]?n?XXXX)"
          r"|ssh-(?:rsa|ed25519) AAAA[A-Za-z0-9+/]{40,}|[Bb]earer [A-Za-z0-9._-]{30,})")
tok_rg = [x for x in rg if x != "-i"] + ["-P", "-e", TOKENS, target]
lines = subprocess.run(tok_rg, capture_output=True, text=True).stdout.splitlines()
lines = [l for l in lines if "XXXX" not in l]
if redact:
    cnt = {}
    for l in lines:
        f = l.split(":", 1)[0]
        cnt[f] = cnt.get(f, 0) + 1
    lines = [f"{f}:{n}" for f, n in cnt.items()]
show("tokens / clés", "\n".join(lines))
# valeurs secrètes connues (jamais affichées)
cands = {}
for k, v in os.environ.items():
    if k.startswith(("SAND_", "CURSOR_", "BOX_", "GH_", "GITHUB_")) and len(v) >= 8 and not v.startswith("/"):
        cands["env:" + k] = v
def walk(o, p):
    if isinstance(o, dict):
        for k, v in o.items(): walk(v, p + "." + k)
    elif isinstance(o, list):
        for i, v in enumerate(o): walk(v, f"{p}[{i}]")
    elif isinstance(o, str) and len(o) >= 8:
        cands[p] = o
for f in ["box-secrets.json", "host-secrets.json", "gateway.json", "teach-queue-key.json", "settings.json"]:
    try: walk(json.load(open("/home/box/sand-data/" + f)), "sand-data/" + f)
    except Exception: pass
for name, val in cands.items():
    r = subprocess.run([x for x in rg if x not in ("-i", "-n")] + ["-l", "-F", "--", val, target], capture_output=True, text=True)
    show(f"valeur secrète {name} (valeur masquée)", r.stdout)
env_names = [os.path.join(dp, f) for dp, _, fs in os.walk(target) for f in fs if f == "env-var-names.txt"]
for p in env_names:
    bad = [(i + 1, l) for i, l in enumerate(open(p)) if "=" in l]
    show(f"valeur d'env dans {p}", "".join(f"{p}:{i}\n" if redact else l for i, l in bad))
print(f"\nRÉSULTAT : {'PROPRE' if findings == 0 else str(findings) + ' finding(s)'} — cible {target}")
sys.exit(0 if findings == 0 else 1)
PY
