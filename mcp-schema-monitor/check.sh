#!/usr/bin/env bash
# Normalise inbox/ (rempli par l'agent via GetDynamicTools) -> snapshot + diff + CHANGELOG.
#   exit 0 "NO_CHANGE last=<snap>" | exit 10 "CHANGED ... report=<md>" (+ reports/NOUVEAU-DIFF.flag)
#   exit 0 "BASELINE snapshot=<snap>" | exit 3 verrou occupé | exit 1 erreur
set -uo pipefail
DIR="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"
mkdir -p "$DIR/logs"
exec 9>"$DIR/.lock"
flock -w 600 9 || { echo "LOCKED: autre check/publish en cours" >&2; exit 3; }
python3 "$DIR/lib/msm.py" "$@" 2> >(tee -a "$DIR/logs/check.err" >&2)
rc=$?
echo "[$(date '+%F %T %Z')] check rc=$rc" >> "$DIR/logs/check.log"
exit $rc
