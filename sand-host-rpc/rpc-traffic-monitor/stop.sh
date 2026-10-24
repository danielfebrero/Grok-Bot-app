#!/usr/bin/env bash
set -euo pipefail
DIR="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"
PIDFILE="$DIR/logs/proxy.pid"
[ -f "$PIDFILE" ] || { echo "aucun pidfile"; exit 0; }
PID="$(cat "$PIDFILE")"
if kill -0 "$PID" 2>/dev/null; then kill "$PID" && echo "arrêté pid $PID"; else echo "process $PID absent"; fi
rm -f "$PIDFILE"
