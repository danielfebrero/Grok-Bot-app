#!/usr/bin/env bash
# start-mitm.sh — lance `mitmdump -s addon_log.py` en arrière-plan (proxy EXPLICITE).
#
#   MITM_PORT   port d'écoute (défaut 8899)
#   MITM_VENV   venv mitmproxy (défaut /workspace/.venvs/mitmproxy)
#   RPC_TRAFFIC_LOG_DIR  dossier de logs (défaut shared-eg, cf. addon)
#
# Mode proxy EXPLICITE : les clients qu'on veut tracer l'utilisent via
#   HTTPS_PROXY=http://127.0.0.1:<port> HTTP_PROXY=http://127.0.0.1:<port>
# On NE détourne PAS globalement/transparemment l'egress du host (cf. README,
# décision de sécurité : éviter de casser l'egress live de l'agent / cert-pinning).
set -euo pipefail
DIR="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"
PORT="${MITM_PORT:-8899}"
VENV="${MITM_VENV:-/workspace/.venvs/mitmproxy}"
MITMDUMP="$VENV/bin/mitmdump"; [ -x "$MITMDUMP" ] || MITMDUMP="$(command -v mitmdump || true)"
[ -n "$MITMDUMP" ] || { echo "ERREUR: mitmdump introuvable (venv $VENV)"; exit 1; }
PIDFILE="$DIR/mitm.pid"
OUT="$DIR/mitm.out"
if [ -f "$PIDFILE" ] && kill -0 "$(cat "$PIDFILE")" 2>/dev/null; then
  echo "déjà lancé (pid $(cat "$PIDFILE"), port $PORT)"; exit 0
fi
# proxy explicite (mode regular), écoute loopback, addon de journalisation
nohup "$MITMDUMP" --mode regular --listen-host 127.0.0.1 --listen-port "$PORT" \
  -s "$DIR/addon_log.py" -q >> "$OUT" 2>&1 &
echo $! > "$PIDFILE"
sleep 1.5
if kill -0 "$(cat "$PIDFILE")" 2>/dev/null; then
  echo "mitmdump démarré pid $(cat "$PIDFILE") sur 127.0.0.1:$PORT (proxy explicite)"
  echo "  usage client : HTTPS_PROXY=http://127.0.0.1:$PORT <commande>"
else
  echo "ÉCHEC démarrage — voir $OUT"; tail -n5 "$OUT" 2>/dev/null || true; exit 1
fi
