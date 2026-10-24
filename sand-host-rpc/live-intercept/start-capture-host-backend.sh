#!/usr/bin/env bash
# start-capture-host-backend.sh — capture PASSIVE host↔backend.
#
# tcpdump (paquets bruts) + decode_keylog.py (tshark + keylog NSS).
# Ne modifie pas le trafic : pas de mitm, pas de redirect nft, pas de CA.
# Doit tourner en root (tcpdump). Le keylog est lu, jamais recopié dans les logs.
#
# Le filtre BPF par défaut ne prend que le tcp/443 hors tunnel egress. Le
# décodeur ne JOURNALISE que les SNI qui matchent HB_ALLOW_HOSTS (leg=host-backend).
# Les autres flux restent chiffrés dans le pcap (0600, détruit par hb-off).
#
# Variables (lib.sh) : HB_IFACE, HB_CAPTURE_FILTER, HB_PCAP, HB_KEYLOG_FILE,
# HB_TLS_PORTS, HB_DECODE_INTERVAL, RPC_TRAFFIC_LOG_DIR, HB_ALLOW_HOSTS.
# ============================================================================
set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# shellcheck source=lib.sh
. "$HERE/lib.sh"

[ "$(id -u)" = "0" ] || { echo "start-capture-host-backend: root requis"; exit 90; }
[ -n "${HB_TCPDUMP:-}" ] && [ -x "$HB_TCPDUMP" ] || HB_TCPDUMP="$(command -v tcpdump || true)"
[ -n "${HB_TSHARK:-}" ] && [ -x "$HB_TSHARK" ] || HB_TSHARK="$(command -v tshark || true)"
[ -n "$HB_TCPDUMP" ] && [ -x "$HB_TCPDUMP" ] || { echo "ERREUR: tcpdump introuvable (apt-get install tcpdump)"; exit 2; }
[ -n "$HB_TSHARK" ] && [ -x "$HB_TSHARK" ] || { echo "ERREUR: tshark introuvable (apt-get install tshark)"; exit 2; }
[ -f "$HB_KEYLOG_FILE" ] || { echo "ERREUR: keylog absent ($HB_KEYLOG_FILE). Lance hb-setup-keylog.sh d'abord."; exit 3; }
[ -n "$HB_CAPTURE_FILTER" ] || { echo "ERREUR: HB_CAPTURE_FILTER vide"; exit 2; }

mkdir -p "$STATE_DIR"
chmod 0600 "$HB_KEYLOG_FILE" 2>/dev/null || true
# ne pas laisser un ancien pcap (il peut contenir du chiffré hors-cible)
rm -f "$HB_PCAP"
rm -f "$HB_CAPTURE_READY"

read -r -a FILTER <<< "$HB_CAPTURE_FILTER"
echo "capture passive iface=$HB_IFACE filtre=${FILTER[*]}"
echo "  keylog=$HB_KEYLOG_FILE (contenu non affiché)"
echo "  pcap=$HB_PCAP"
echo "  allow=$HB_ALLOW_HOSTS"
echo "  logs=$TRAFFIC_LOG_DIR"

# -Z root : ne pas dropper vers l'utilisateur tcpdump (le STATE_DIR n'est pas
# forcément writable par lui). -U : un paquet = un flush, pour que tshark lise
# un pcap cohérent pendant l'écriture.
"$HB_TCPDUMP" -i "$HB_IFACE" -n -s 0 -U -Z root -w "$HB_PCAP" "${FILTER[@]}" \
    >>"$HB_CAPTURE_OUT" 2>&1 &
td=$!
echo "$td" > "$HB_TCPDUMP_PIDFILE"

ready=0
for _ in $(seq 1 50); do
  if ! kill -0 "$td" 2>/dev/null; then
    echo "ERREUR: tcpdump pid $td est mort au démarrage (voir $HB_CAPTURE_OUT)"
    exit 6
  fi
  if [ -f "$HB_PCAP" ]; then
    chmod 0600 "$HB_PCAP" 2>/dev/null || true
    ready=1
    break
  fi
  sleep 0.1
done
[ "$ready" = 1 ] || { echo "ERREUR: pcap non créé"; kill -TERM "$td" 2>/dev/null || true; exit 6; }
touch "$HB_CAPTURE_READY"

# Ce process EST le décodeur après exec : son pid est déjà celui du script.
echo $$ > "$HB_DECODE_PIDFILE"
export HB_ALLOW_HOSTS HB_TUNNEL_ADDR HB_TSHARK HB_PCAP_MAX_BYTES HB_TCPDUMP_PIDFILE
export RPC_TRAFFIC_LOG_DIR="$TRAFFIC_LOG_DIR"
exec python3 -u "$HERE/decode_keylog.py" \
    --follow --interval "$HB_DECODE_INTERVAL" \
    --pcap "$HB_PCAP" --keylog "$HB_KEYLOG_FILE" --seen "$HB_SEEN_FILE" \
    --tls-ports "$HB_TLS_PORTS" --tshark "$HB_TSHARK" \
    --max-bytes "$HB_PCAP_MAX_BYTES" --tcpdump-pidfile "$HB_TCPDUMP_PIDFILE"
