#!/usr/bin/env bash
# Détection automatique d'une mise à jour (box update OU hot-update du host bundle).
# Compare une empreinte rapide (versions/build + sha256 des fichiers clés) au dernier snapshot.
#   - pas de changement  -> imprime "NO_CHANGE ..."                     exit 0
#   - changement         -> nouveau snapshot + rapport diff + ligne ajoutée à
#                           reports/NOUVEAU-DIFF.flag ; imprime "CHANGED ... report=<chemin>"   exit 10
#   - aucun snapshot     -> crée le baseline, imprime "BASELINE ..."    exit 0
#   - erreur             -> exit 1 (ou 3 si verrou occupé)
# Options : --deep  (si l'empreinte rapide est identique, refait un manifeste complet pour attraper
#                    tout changement hors fichiers clés ; plus lent ~1 min)
set -uo pipefail
DIR="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"
mkdir -p "$DIR/logs"
exec 9>"$DIR/.lock"
flock -w 600 9 || { echo "LOCKED: autre snapshot/check en cours" >&2; exit 3; }
python3 "$DIR/lib/bum.py" check "$@" 2> >(tee -a "$DIR/logs/check-update.err" >&2)
rc=$?
echo "[$(date '+%F %T %Z')] check-update rc=$rc" >> "$DIR/logs/check-update.log"
exit $rc
