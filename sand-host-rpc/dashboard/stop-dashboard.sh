#!/usr/bin/env bash
# Arrête le dashboard.
set -euo pipefail
DIR="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"
PIDFILE="$DIR/dashboard.pid"
if [ -f "$PIDFILE" ] && kill -0 "$(cat "$PIDFILE")" 2>/dev/null; then
  kill "$(cat "$PIDFILE")" && echo "arrêté pid $(cat "$PIDFILE")"
  rm -f "$PIDFILE"
else
  echo "non lancé"; rm -f "$PIDFILE" 2>/dev/null || true
fi
