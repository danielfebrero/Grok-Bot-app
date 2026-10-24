#!/usr/bin/env bash
# hb-revert.sh — TEARDOWN COMPLET de la patte host↔backend. Idempotent.
# --dry-run : affiche le plan, n'exige pas root, n'agit pas.
#
#   1. arrête tcpdump + décodeur, détruit le pcap.
#   2. supprime le reliquat mitm (table ip sandrpc_hb, mitmdump, marqueur CA).
#      NE TOUCHE JAMAIS ip sandrpc ni le relai agent↔host.
#   3. SIGTERM host-main -> le superviseur relance un process SANS --tls-keylog
#      et SANS SSLKEYLOGFILE. Puis shred du keylog (secrets de session).
#   4. retire le dead-man dédié HB_ARM / HB_DISARM.
# ============================================================================
set -uo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# shellcheck source=lib.sh
. "$HERE/lib.sh"
DRY=0; [ "${1:-}" = "--dry-run" ] && DRY=1

if [ "$DRY" = 1 ]; then
  echo "DRY hb-revert: arrêt capture (tcpdump+décodeur) et shred $HB_PCAP"
  echo "DRY hb-revert: supprimer reliquat nft ip $HB_NFT_TABLE + mitm :$HB_MITM_PORT (jamais ip sandrpc)"
  echo "DRY hb-revert: SIGTERM host-main (respawn superviseur SANS --tls-keylog) puis shred $HB_KEYLOG_FILE"
  echo "DRY hb-revert: rm $HB_KEYLOG_READY_FLAG $HB_ARM_FLAG $HB_DISARM_FLAG $HB_CA_READY_FLAG"
  echo "DRY hb-revert: NE PAS toucher arm-agent-host / relai / table ip sandrpc"
  exit 0
fi

need_root
slog "hb-revert: début — teardown keylog. NE touche PAS la patte agent<->host (ip sandrpc)"

hb_stop_pids
hb_shred "$HB_PCAP"
hb_drop_legacy_nft
hb_kill_legacy_mitm

hp="$(hostmain_pid)"
if [ -n "$hp" ]; then
  slog "hb-revert: SIGTERM host-main $hp (respawn superviseur sans --tls-keylog, port $HOST_PORT_ORIG)"
  kill -TERM "$hp" 2>/dev/null || true
else
  slog "hb-revert: aucun host-main vivant ; le superviseur va en lancer un"
fi
# laisser le process mourir avant de shredder le fichier qu'il pourrait recréer
sleep 1
hb_shred "$HB_KEYLOG_FILE"
rm -f "$HB_HOSTMAIN_PIDFILE" "$HB_KEYLOG_READY_FLAG" "$HB_SEEN_FILE" "$HB_CA_READY_FLAG"

if [ -f "$WATCHDOG_PIDFILE" ]; then
  wp="$(cat "$WATCHDOG_PIDFILE" 2>/dev/null || true)"
  if [ -n "${wp:-}" ] && kill -0 "$wp" 2>/dev/null; then
    slog "hb-revert: arrêt watchdog pid $wp"
    kill "$wp" 2>/dev/null || true
  fi
  rm -f "$WATCHDOG_PIDFILE"
fi
rm -f "$HB_ARM_FLAG" "$HB_DISARM_FLAG"

ok=0
for _ in $(seq 1 30); do
  if host_health "$HOST_PORT_ORIG"; then ok=1; break; fi
  sleep 2
done
if [ "$ok" = 1 ]; then
  slog "hb-revert: host-main sain sur $HOST_PORT_ORIG"
else
  slog "hb-revert: ATTENTION host-main pas (encore) sain sur $HOST_PORT_ORIG"
fi
if backend_reachable; then
  slog "hb-revert: backend joignable"
else
  slog "hb-revert: ATTENTION backend injoignable"
fi
if command -v nft >/dev/null 2>&1 && nft list table ip sandrpc >/dev/null 2>&1; then
  slog "hb-revert: table ip sandrpc (agent<->host) intacte."
fi
slog "hb-revert: FIN"
exit 0
