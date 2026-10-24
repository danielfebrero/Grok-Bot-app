#!/usr/bin/env bash
# start-capture-cdp.sh — capture PASSIVE du WebSocket CDP (loopback, en clair).
#
# host-main ouvre ws://127.0.0.1:<cdpPort>/devtools/... . Ce n'est ni le relai
# :1342, ni le tcp/443. tcpdump sur lo + decode_keylog.py --plaintext.
# Pas de keylog, pas de mitm, pas de nft. Root requis (tcpdump).
# Même rotation que le 443 : au-delà de CDP_PCAP_MAX_BYTES le décodeur shred
# et relance tcpdump. Un socket CDP déjà ouvert reste opaque tant que le
# handshake HTTP (Upgrade) n'est pas dans le nouveau pcap.
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
hb_shred "${CDP_PCAP}.closed"
: > "$CDP_CAPTURE_OUT"
chmod 0600 "$CDP_CAPTURE_OUT" 2>/dev/null || true

echo "capture CDP iface=$CDP_IFACE ports=${ports[*]}"
echo "  pcap=$CDP_PCAP"
echo "  filtre=$filt"
echo "  logs=$TRAFFIC_LOG_DIR leg=host-cdp"

python3 - "$HERE" "$CDP_TCPDUMP_PIDFILE" "$HB_TCPDUMP" "$CDP_IFACE" "$CDP_PCAP" "$CDP_CAPTURE_OUT" "$HB_TCPDUMP_BUFFER_KB" "$filt" <<'PY'
import sys
here, pidfile, tcpdump, iface, pcap, logout, buffer, filt = sys.argv[1:]
sys.path.insert(0, here)
import decode_keylog as dk
try:
    kb = int(buffer)
except ValueError:
    kb = dk.DEFAULT_TCPDUMP_BUFFER_KB
pid = dk.spawn_tcpdump(tcpdump, iface, pcap, logout, filt, pidfile, kb)
if pid <= 0:
    sys.stderr.write("ERREUR: tcpdump n'a pas démarré\n")
    raise SystemExit(2)
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
export HB_TSHARK CDP_PCAP_MAX_BYTES HB_TCPDUMP_BUFFER_KB
export HB_TCPDUMP_PIDFILE="$CDP_TCPDUMP_PIDFILE"
export HB_PCAP_MAX_BYTES="$CDP_PCAP_MAX_BYTES"
export RPC_TRAFFIC_LOG_DIR="$TRAFFIC_LOG_DIR"
exec python3 -u "$HERE/decode_keylog.py" \
    --plaintext --follow --interval "$CDP_DECODE_INTERVAL" \
    --pcap "$CDP_PCAP" --seen "$CDP_SEEN_FILE" \
    --cdp-ports "${ports[*]}" --tshark "$HB_TSHARK" \
    --max-bytes "$CDP_PCAP_MAX_BYTES" --tcpdump-pidfile "$CDP_TCPDUMP_PIDFILE" \
    --tcpdump "$HB_TCPDUMP" --iface "$CDP_IFACE" \
    --capture-filter "$filt" --capture-out "$CDP_CAPTURE_OUT" \
    --tcpdump-buffer-kb "$HB_TCPDUMP_BUFFER_KB"
