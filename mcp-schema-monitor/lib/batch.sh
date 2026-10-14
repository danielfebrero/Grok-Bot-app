#!/usr/bin/env bash
# usage : batch.sh "ns tool fichier [extractions...]" ...   (fichiers relatifs à /workspace/agent-tools, sans .txt)
cd "$(dirname "$0")/.." || exit 1
for line in "$@"; do set -- $line; ns=$1 t=$2 f=$3; shift 3
  python3 lib/infer.py "$ns" "$t" "/workspace/agent-tools/$f.txt" "$@"; done
