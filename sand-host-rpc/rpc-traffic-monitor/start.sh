#!/usr/bin/env bash
# Démarre le proxy de journalisation en arrière-plan.
#   RTM_PORT      port d'écoute local      (défaut 8787)
#   RTM_UPSTREAM  cible ré-origine         (défaut https://api2.cursor.sh)
#   RTM_LEG       étiquette de patte       (agent-host | host-backend)
# Exemples :
#   RTM_LEG=host-backend RTM_UPSTREAM=https://api2.cursor.sh ./start.sh
#   RTM_LEG=agent-host  RTM_UPSTREAM=http://127.0.0.1:1340 RTM_PORT=8788 ./start.sh
set -euo pipefail
DIR="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"
mkdir -p "$DIR/logs"
PIDFILE="$DIR/logs/proxy.pid"
if [ -f "$PIDFILE" ] && kill -0 "$(cat "$PIDFILE")" 2>/dev/null; then
  echo "déjà lancé (pid $(cat "$PIDFILE"))"; exit 0
fi
nohup python3 "$DIR/proxy.py" >> "$DIR/logs/proxy.out" 2>&1 &
echo $! > "$PIDFILE"
sleep 0.5
echo "démarré pid $(cat "$PIDFILE") — voir logs/proxy.out"
