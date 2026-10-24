#!/usr/bin/env bash
# Preuve de capture req+resp : lance le proxy vers api2.cursor.sh, envoie un
# appel Connect au travers, montre la dernière ligne JSONL (secrets rédigés).
set -euo pipefail
DIR="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"
PORT="${RTM_PORT:-8799}"
LOG="$DIR/logs/traffic-$(date +%F).jsonl"
python3 "$DIR/proxy.py" --port "$PORT" --upstream "https://api2.cursor.sh" --leg host-backend &
P=$!; trap 'kill $P 2>/dev/null || true' EXIT
sleep 1
# appel Connect unaire non authentifié -> renvoie une erreur, prouve req+resp
curl -s -o /dev/null -X POST \
  -H 'content-type: application/json' \
  -H 'connect-protocol-version: 1' \
  --data '{}' \
  "http://127.0.0.1:$PORT/aiserver.v1.DashboardService/GetTeams" || true
sleep 0.5
echo "=== dernière ligne de $LOG (secrets rédigés) ==="
tail -n1 "$LOG" | python3 -m json.tool
