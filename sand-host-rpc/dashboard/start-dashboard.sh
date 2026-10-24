#!/usr/bin/env bash
# Démarre le dashboard de visualisation du trafic RPC en arrière-plan.
#   DASH_PORT            port d'écoute                (défaut 8808)
#   DASH_HOST            hôte d'écoute                (défaut 127.0.0.1)
#   RPC_TRAFFIC_LOG_DIR  dossier des JSONL rédactés   (défaut shared-eg)
set -euo pipefail
DIR="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"
PIDFILE="$DIR/dashboard.pid"
if [ -f "$PIDFILE" ] && kill -0 "$(cat "$PIDFILE")" 2>/dev/null; then
  echo "déjà lancé (pid $(cat "$PIDFILE")) sur ${DASH_HOST:-127.0.0.1}:${DASH_PORT:-8808}"; exit 0
fi
nohup python3 "$DIR/server.py" >> "$DIR/dashboard.out" 2>&1 &
echo $! > "$PIDFILE"
sleep 0.6
echo "démarré pid $(cat "$PIDFILE") — http://${DASH_HOST:-127.0.0.1}:${DASH_PORT:-8808}  (voir dashboard.out)"
