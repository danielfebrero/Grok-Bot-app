#!/usr/bin/env bash
# Compare deux snapshots et écrit reports/<snapshot-récent>.md (+ reports/<...>-diffs/*.diff)
# Usage :
#   diff.sh              -> les 2 snapshots les plus récents
#   diff.sh OLD          -> OLD vs le plus récent
#   diff.sh OLD NEW      -> OLD vs NEW   (noms = dossiers de snapshots/)
#   diff.sh --list       -> liste les snapshots
set -euo pipefail
DIR="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"
if [[ "${1:-}" == "--list" ]]; then
  ls -1 "$DIR/snapshots" 2>/dev/null | grep -v '\.partial$' || true; exit 0
fi
exec python3 "$DIR/lib/shr.py" diff "$@"
