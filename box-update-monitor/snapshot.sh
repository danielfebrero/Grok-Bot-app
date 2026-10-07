#!/usr/bin/env bash
# Crée un snapshot horodaté du périmètre app Grok Bot : snapshots/<YYYY-MM-DD_HH-MM>/
#   manifest.tsv (path size sha256 mtime mode kind perimeter), files/ (copies texte), meta/, fingerprint.txt
# Usage : snapshot.sh [--label "texte libre"]
set -euo pipefail
DIR="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"
exec 9>"$DIR/.lock"
flock -w 600 9 || { echo "verrou occupé (autre snapshot/check en cours)" >&2; exit 3; }
exec python3 "$DIR/lib/bum.py" snapshot "$@"
