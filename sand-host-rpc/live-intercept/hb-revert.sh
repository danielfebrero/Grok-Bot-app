#!/usr/bin/env bash
# hb-revert.sh — TEARDOWN COMPLET de la patte host<->backend. Idempotent. --dry-run.
# ============================================================================
#  Étapes :
#   1. supprime la table ip sandrpc_hb (retire la redirection). no-op si absente.
#   2. tue mitmproxy transparent sélectif (pidfile + résiduels).
#   3. SIGTERM host-main -> le superviseur en relance un PROPRE (sans CA, api2
#      direct). Retire le marqueur CA-ready.
#   4. retire les flags dead-man dédiés (HB_ARM/HB_DISARM) + arrête son watchdog.
#
#  NE TOUCHE JAMAIS la table ip sandrpc NI le relai de la patte agent<->host.
#  (Le SIGTERM host-main : le relai agent-host pointe sur 172.30.0.2:1340 ; le
#   superviseur relance host-main sur 1340 -> la patte agent<->host se rétablit.)
# ============================================================================
set -uo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
. "$HERE/lib.sh"
DRY=0; [ "${1:-}" = "--dry-run" ] && DRY=1
run() { if [ "$DRY" = 1 ]; then echo "DRY: $*"; else eval "$@"; fi; }

need_root
slog "hb-revert: début (dry=$DRY) — NE touche PAS la patte agent<->host (ip sandrpc)"

# 1) nft : delete table dédiée host<->backend. no-op si absente.
if $SUDO nft list table ip "$HB_NFT_TABLE" >/dev/null 2>&1; then
  slog "hb-revert: suppression table ip $HB_NFT_TABLE"
  run "$SUDO nft delete table ip $HB_NFT_TABLE"
else
  slog "hb-revert: table ip $HB_NFT_TABLE absente (ok)"
fi

# 2) mitm transparent sélectif
if [ -f "$HB_MITM_PIDFILE" ]; then
  p="$(cat "$HB_MITM_PIDFILE" 2>/dev/null)"
  if [ -n "$p" ] && kill -0 "$p" 2>/dev/null; then
    slog "hb-revert: kill mitm pid $p"
    run "kill $p 2>/dev/null || $SUDO kill $p 2>/dev/null || true"
  fi
  run "rm -f $HB_MITM_PIDFILE"
fi
if [ "$DRY" = 0 ]; then
  for p in $(pgrep -f "mitmdump .*--listen-port ${HB_MITM_PORT}" 2>/dev/null); do
    slog "hb-revert: kill mitm résiduel pid $p"; kill "$p" 2>/dev/null || $SUDO kill "$p" 2>/dev/null || true
  done
fi

# 3) host-main : SIGTERM -> respawn propre par le superviseur (sans CA)
hp="$(hostmain_pid)"
if [ -n "$hp" ]; then
  slog "hb-revert: SIGTERM host-main $hp (le superviseur relance un host propre sans CA sur $HOST_PORT_ORIG)"
  run "kill -TERM $hp 2>/dev/null || $SUDO kill -TERM $hp 2>/dev/null || true"
else
  slog "hb-revert: aucun host-main vivant ; le superviseur va en lancer un"
fi
run "rm -f $HB_HOSTMAIN_PIDFILE"
run "rm -f $HB_CA_READY_FLAG"

# 4) dead-man dédié : watchdog + flags HB_ARM/HB_DISARM
if [ -f "$WATCHDOG_PIDFILE" ]; then
  wp="$(cat "$WATCHDOG_PIDFILE" 2>/dev/null)"
  if [ -n "$wp" ] && kill -0 "$wp" 2>/dev/null; then
    slog "hb-revert: arrêt watchdog pid $wp"
    run "kill $wp 2>/dev/null || $SUDO kill $wp 2>/dev/null || true"
  fi
  run "rm -f $WATCHDOG_PIDFILE"
fi
run "rm -f $HB_ARM_FLAG"
run "rm -f $HB_DISARM_FLAG"

# 5) sanity (sans toucher agent-host)
if [ "$DRY" = 0 ]; then
  ok=0
  for i in $(seq 1 30); do if host_health "$HOST_PORT_ORIG"; then ok=1; break; fi; sleep 2; done
  if [ "$ok" = 1 ]; then slog "hb-revert: host-main sain sur $HOST_PORT_ORIG"; else slog "hb-revert: ATTENTION host-main pas (encore) sain sur $HOST_PORT_ORIG"; fi
  if backend_reachable; then slog "hb-revert: backend joignable"; else slog "hb-revert: ATTENTION backend injoignable"; fi
  if $SUDO nft list table ip sandrpc >/dev/null 2>&1; then slog "hb-revert: table ip sandrpc (agent<->host) intacte."; fi
fi
slog "hb-revert: FIN (dry=$DRY)"
exit 0
