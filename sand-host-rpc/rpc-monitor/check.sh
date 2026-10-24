#!/usr/bin/env bash
# Détection d'un changement de surface RPC depuis le dernier snapshot.
#   - pas de changement  -> "NO_CHANGE ..."                                   exit 0
#   - changement         -> nouveau snapshot + rapport diff + entrée COMMENTÉE
#                           dans CHANGELOG.md ; "CHANGED ... report=<chemin>" exit 10
#   - aucun snapshot     -> crée le baseline ; "BASELINE ..."                 exit 0
#   - erreur             -> exit 1        - verrou occupé -> exit 3
set -uo pipefail
DIR="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"
mkdir -p "$DIR/logs"
exec 9>"$DIR/.lock"
flock -w 600 9 || { echo "LOCKED: autre snapshot/check en cours" >&2; exit 3; }
python3 "$DIR/lib/shr.py" check "$@" 2> >(tee -a "$DIR/logs/check.err" >&2)
rc=$?
echo "[$(date '+%F %T %Z')] check rc=$rc" >> "$DIR/logs/check.log"
exit $rc
