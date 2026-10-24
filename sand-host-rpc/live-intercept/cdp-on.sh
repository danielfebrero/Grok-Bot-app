#!/usr/bin/env bash
# cdp-on.sh — capture passive du WebSocket CDP (loopback). Sans restart, sans nft.
#
# Dry-run par défaut. La capture réelle :
#   CDP_ON_CONFIRM=yes sudo -E bash cdp-on.sh
# Ports : $CDP_PORTS si posé (6080/6081 refusés), sinon découverte
# GET http://127.0.0.1:<port>/json/version → webSocketDebuggerUrl.
# Ports connus de cette box, sondés en plus des listeners : 9229 9257 9302.
#
# OFF : bash cdp-off.sh   (ne touche pas le keylog 443 ni host-main)
# ============================================================================
set -uo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# shellcheck source=lib.sh
. "$HERE/lib.sh"
CONFIRM="${CDP_ON_CONFIRM:-}"
DRY=1; [ "$CONFIRM" = "yes" ] && DRY=0

RESOLVED="$(python3 "$HERE/decode_keylog.py" --discover-cdp 2>/dev/null || true)"
# shellcheck disable=SC2086
set -- $RESOLVED
PORT_LIST="$*"

plan() {
  cat <<PLAN
=== PLAN cdp-on (DRY=$DRY) — WebSocket CDP loopback, SANS restart, SANS nft ===
 ports : ${PORT_LIST:-<aucun — poser CDP_PORTS ou lancer Chrome DevTools>}
 1. tcpdump -i $CDP_IFACE -B ${HB_TCPDUMP_BUFFER_KB}  filtre tcp port <ports> (pas 6080/6081, pas ${HB_TUNNEL_ADDR})
 2. decode_keylog.py --plaintext  (tshark -d tcp.port==PORT,http, couche websocket)
    journal : $TRAFFIC_LOG_DIR   leg=host-cdp   protocol=websocket
    cap pcap : $CDP_PCAP_MAX_BYTES octets, puis rotation auto (socket déjà ouvert :
    opaque jusqu'au prochain handshake HTTP dans le pcap courant)
 noVNC/websockify 6080/6081 et le tunnel egress restent hors capture.
 OFF : bash $HERE/cdp-off.sh
PLAN
}
plan
if [ "$DRY" = 1 ]; then
  echo
  echo ">>> DRY-RUN : rien n'est capturé."
  echo ">>> Activer : CDP_ON_CONFIRM=yes sudo -E bash cdp-on.sh"
  echo ">>> Forcer les ports : CDP_PORTS='9229 9257' CDP_ON_CONFIRM=yes sudo -E bash cdp-on.sh"
  exit 0
fi

need_root
if [ -z "$PORT_LIST" ]; then
  slog "cdp-on: aucun port CDP (json/version muet, CDP_PORTS vide). Rien démarré."
  exit 4
fi
for p in $PORT_LIST; do
  case "$p" in
    6080|6081) slog "cdp-on: port $p exclu (noVNC)"; exit 2;;
  esac
done

if [ -f "$CDP_TCPDUMP_PIDFILE" ] && hb_pid_live "$(cat "$CDP_TCPDUMP_PIDFILE" 2>/dev/null || true)"; then
  slog "cdp-on: capture déjà active. Rien à faire."
  exit 0
fi
rm -f "$CDP_TCPDUMP_PIDFILE" "$CDP_DECODE_PIDFILE" "$CDP_CAPTURE_READY"
slog "cdp-on: démarrage capture CDP ports $PORT_LIST"
export CDP_PORTS="$PORT_LIST"
setsid bash "$HERE/start-capture-cdp.sh" </dev/null >>"$CDP_CAPTURE_OUT" 2>&1 &
echo $! > "$CDP_DECODE_PIDFILE"
ok=0
for _ in $(seq 1 40); do
  [ -f "$CDP_CAPTURE_READY" ] && { ok=1; break; }
  sleep 0.25
done
if [ "$ok" != 1 ]; then
  slog "cdp-on: capture non prête (voir $CDP_CAPTURE_OUT)"
  exit 6
fi
slog "cdp-on: capture CDP prête (ports $PORT_LIST, leg=host-cdp)"
exit 0
