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
rm -f "$HB_CAPTURE_READY" "$HB_TCPDUMP_PIDFILE"
: > "$HB_CAPTURE_OUT"
chmod 0600 "$HB_CAPTURE_OUT" 2>/dev/null || true

read -r -a FILTER <<< "$HB_CAPTURE_FILTER"
echo "capture passive iface=$HB_IFACE filtre=${FILTER[*]}"
echo "  keylog=$HB_KEYLOG_FILE (contenu non affiché)"
echo "  pcap=$HB_PCAP"
echo "  allow=$HB_ALLOW_HOSTS"
echo "  logs=$TRAFFIC_LOG_DIR"

# tcpdump est double-forké (setsid) : son père est init, pas le décodeur.
# Sinon, après exec python, un SIGTERM laisse un zombie et kill -0 reste vrai.
# --immediate-mode : sans ça, libpcap (TPACKET_V3) ne livre les paquets qu'à
# la fin du délai de tampon (~1 s). Un hb-off dans la seconde les compte
# « received » mais en écrit 0 — le pcap reste un en-tête vide.
# -Z root : ne pas dropper vers l'utilisateur tcpdump (STATE_DIR pas writable
# par lui). -U : un paquet = un flush, pour que tshark lise un pcap cohérent.
python3 - "$HB_TCPDUMP_PIDFILE" "$HB_TCPDUMP" "$HB_IFACE" "$HB_PCAP" "$HB_CAPTURE_OUT" "${FILTER[@]}" <<'PY'
import os, sys
pidfile, tcpdump, iface, pcap, logout, *filt = sys.argv[1:]
if not filt:
    sys.stderr.write("ERREUR: filtre BPF vide\n")
    raise SystemExit(2)
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
    "-w", pcap, *filt,
])
PY

# Le fichier pcap (en-tête) existe avant que le filtre soit actif. On attend
# le « listening » de tcpdump, puis un court délai : un client lancé dans la
# foulée raterait sinon le ClientHello.
ready=0
td=""
for _ in $(seq 1 50); do
  if [ -s "$HB_TCPDUMP_PIDFILE" ]; then
    td="$(cat "$HB_TCPDUMP_PIDFILE" 2>/dev/null || true)"
    if [ -n "$td" ] && ! hb_pid_live "$td"; then
      echo "ERREUR: tcpdump pid $td est mort au démarrage (voir $HB_CAPTURE_OUT)"
      cat "$HB_CAPTURE_OUT" >&2 || true
      exit 6
    fi
  fi
  if grep -q "listening on" "$HB_CAPTURE_OUT" 2>/dev/null; then
    sleep 0.4
    chmod 0600 "$HB_PCAP" 2>/dev/null || true
    ready=1
    break
  fi
  sleep 0.1
done
if [ "$ready" != 1 ]; then
  echo "ERREUR: tcpdump n'écoute pas (voir $HB_CAPTURE_OUT)"
  cat "$HB_CAPTURE_OUT" >&2 || true
  [ -n "$td" ] && kill -TERM "$td" 2>/dev/null || true
  exit 6
fi
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
