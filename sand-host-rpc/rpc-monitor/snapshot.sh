#!/usr/bin/env bash
# Crée un snapshot horodaté des méthodes RPC : snapshots/<YYYY-MM-DD_HH-MM>/
#   manifest.json (toutes les méthodes + signature I|O|kind), methodes.tsv, fingerprint.json, SNAPSHOT-INFO.txt
# Usage : snapshot.sh [--label "texte libre"]
set -euo pipefail
DIR="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"
exec 9>"$DIR/.lock"
flock -w 600 9 || { echo "verrou occupé (autre snapshot/check en cours)" >&2; exit 3; }
exec python3 "$DIR/lib/shr.py" snapshot "$@"
