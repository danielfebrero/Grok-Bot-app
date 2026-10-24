#!/usr/bin/env bash
# start-capture-cdp.sh — capture PASSIVE du WebSocket CDP (loopback, en clair).
#
# host-main ouvre ws://127.0.0.1:<cdpPort>/devtools/... . Ce n'est ni le relai
# :1342, ni le tcp/443. tcpdump sur lo + decode_keylog.py --plaintext.
# Pas de keylog, pas de mitm, pas de nft. Root requis (tcpdump).
#
# Variables : CDP_PORTS (obligatoire, liste séparée par espace ou virgule),
# CDP_IFACE (défaut lo), CDP_PCAP, RPC_TRAFFIC_LOG_DIR. Les ports 6080 et 6081
# (noVNC) sont refusés.
# ============================================================================
set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# shellcheck source=lib.sh
. "$HERE/lib.sh"

[ "$(id -u)" = "0" ] || { echo "start-capture-cdp: root requis"; exit 90; }
[ -n "${HB_TCPDUMP:-}" ] && [ -x "$HB_TCPDUMP" ] || HB_TCPDUMP="$(command -v tcpdump || true)"
[ -n "${HB_TSHARK:-}" ] && [ -x "$HB_TSHARK" ] || HB_TSHARK="$(command -v tshark || true)"
[ -n "$HB_TCPDUMP" ] && [ -x "$HB_TCPDUMP" ] || { echo "ERREUR: tcpdump introuvable"; exit 2; }
[ -n "$HB_TSHARK" ] && [ -x "$HB_TSHARK" ] || { echo "ERREUR: tshark introuvable"; exit 2; }
[ -n "${CDP_PORTS:-}" ] || { echo "ERREUR: CDP_PORTS vide (cdp-on.sh découvre, ou exporter la liste)"; exit 2; }

ports=()
for p in ${CDP_PORTS//,/ }; do
  case "$p" in
    6080|6081)
      echo "ERREUR: port $p (noVNC/websockify) hors périmètre" >&2
      exit 2
      ;;
    ''|*[!0-9]*)
      echo "ERREUR: port CDP invalide: $p" >&2
      exit 2
      ;;
  esac
  ports+=("$p")
done
[ "${#ports[@]}" -gt 0 ] || { echo "ERREUR: aucun port CDP"; exit 2; }

filt="tcp and ("
first=1
for p in "${ports[@]}"; do
  if [ "$first" = 1 ]; then filt+="port $p"; first=0; else filt+=" or port $p"; fi
done
filt+=") and not host ${HB_TUNNEL_ADDR}"

mkdir -p "$STATE_DIR"
rm -f "$CDP_PCAP" "$CDP_CAPTURE_READY" "$CDP_TCPDUMP_PIDFILE"
: > "$CDP_CAPTURE_OUT"
chmod 0600 "$CDP_CAPTURE_OUT" 2>/dev/null || true

echo "capture CDP iface=$CDP_IFACE ports=${ports[*]}"
echo "  pcap=$CDP_PCAP"
echo "  filtre=$filt"
echo "  logs=$TRAFFIC_LOG_DIR leg=host-cdp"

python3 - "$CDP_TCPDUMP_PIDFILE" "$HB_TCPDUMP" "$CDP_IFACE" "$CDP_PCAP" "$CDP_CAPTURE_OUT" "$filt" <<'PY'
import os, sys
pidfile, tcpdump, iface, pcap, logout, filt = sys.argv[1:]
if os.fork() > 0:
    os._exit(0)
os.setsid()
if os.fork() > 0:
    os._exit(0)
with open(pidfile, "w", encoding="utf-8") as fh:
    fh.write(str(os.getpid()) + "\n")
fd = os.open(logout, os.O_WRONLY | os.O_CREAT | os.O_APPEND, 0o600)
os.dup2(fd, 1)
os.dup2(fd, 2)
if fd > 2:
    os.close(fd)
os.execvp(tcpdump, [
    tcpdump, "--immediate-mode", "-i", iface, "-n", "-s", "0", "-U", "-Z", "root",
    "-w", pcap, filt,
])
PY

ready=0
td=""
for _ in $(seq 1 50); do
  if [ -s "$CDP_TCPDUMP_PIDFILE" ]; then
    td="$(cat "$CDP_TCPDUMP_PIDFILE" 2>/dev/null || true)"
    if [ -n "$td" ] && ! hb_pid_live "$td"; then
      echo "ERREUR: tcpdump pid $td est mort (voir $CDP_CAPTURE_OUT)"
      cat "$CDP_CAPTURE_OUT" >&2 || true
      exit 6
    fi
  fi
  if grep -q "listening on" "$CDP_CAPTURE_OUT" 2>/dev/null; then
    sleep 0.4
    chmod 0600 "$CDP_PCAP" 2>/dev/null || true
    ready=1
    break
  fi
  sleep 0.1
done
if [ "$ready" != 1 ]; then
  echo "ERREUR: tcpdump CDP n'écoute pas (voir $CDP_CAPTURE_OUT)"
  cat "$CDP_CAPTURE_OUT" >&2 || true
  [ -n "$td" ] && kill -TERM "$td" 2>/dev/null || true
  exit 6
fi
touch "$CDP_CAPTURE_READY"

echo $$ > "$CDP_DECODE_PIDFILE"
export HB_TSHARK CDP_PCAP_MAX_BYTES
export HB_TCPDUMP_PIDFILE="$CDP_TCPDUMP_PIDFILE"
export HB_PCAP_MAX_BYTES="$CDP_PCAP_MAX_BYTES"
export RPC_TRAFFIC_LOG_DIR="$TRAFFIC_LOG_DIR"
exec python3 -u "$HERE/decode_keylog.py" \
    --plaintext --follow --interval "$CDP_DECODE_INTERVAL" \
    --pcap "$CDP_PCAP" --seen "$CDP_SEEN_FILE" \
    --cdp-ports "${ports[*]}" --tshark "$HB_TSHARK" \
    --max-bytes "$CDP_PCAP_MAX_BYTES" --tcpdump-pidfile "$CDP_TCPDUMP_PIDFILE"
