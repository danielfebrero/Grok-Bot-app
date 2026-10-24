#!/usr/bin/env bash
# revert-agent-host.sh — annule l'interception agent<->host "sans restart".
#
# NE TOUCHE JAMAIS host-main (il n'a jamais été redémarré par arm-agent-host.sh).
# Idempotent. Marche aussi appelé seul. --dry-run pour voir sans agir.
#
# Étapes :
#   1. supprime la table ip sandrpc (retire la redirection prerouting). no-op si absente.
#   2. tue le relai loggeur (pidfile + résiduels relay_agent_host.py). libère :1342.
#   3. supprime les flags d'état (ARM/DISARM) + arrête le watchdog.
set -uo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
. "$HERE/lib.sh"

DRY=0; [ "${1:-}" = "--dry-run" ] && DRY=1
run() { if [ "$DRY" = 1 ]; then echo "DRY: $*"; else eval "$@"; fi; }

need_root
slog "revert-ah: début (dry=$DRY) — host-main NON touché"

# 1) nft : delete table dédiée. no-op si absente.
if $SUDO nft list table ip "$AH_NFT_TABLE" >/dev/null 2>&1; then
  slog "revert-ah: suppression table ip $AH_NFT_TABLE"
  run "$SUDO nft delete table ip $AH_NFT_TABLE"
else
  slog "revert-ah: table ip $AH_NFT_TABLE absente (ok)"
fi

# 2) relai loggeur : pidfile puis résiduels
if [ -f "$AH_RELAY_PIDFILE" ]; then
  p="$(cat "$AH_RELAY_PIDFILE" 2>/dev/null)"
  if [ -n "$p" ] && kill -0 "$p" 2>/dev/null; then
    slog "revert-ah: kill relai pid $p"
    run "kill $p 2>/dev/null || $SUDO kill $p 2>/dev/null || true"
  fi
  run "rm -f $AH_RELAY_PIDFILE"
fi
if [ "$DRY" = 0 ]; then
  for p in $(pgrep -f 'relay_agent_host\.py' 2>/dev/null); do
    slog "revert-ah: kill relai résiduel pid $p"
    kill "$p" 2>/dev/null || $SUDO kill "$p" 2>/dev/null || true
  done
fi

# 3) watchdog + flags d'état
if [ -f "$WATCHDOG_PIDFILE" ]; then
  wp="$(cat "$WATCHDOG_PIDFILE" 2>/dev/null)"
  if [ -n "$wp" ] && kill -0 "$wp" 2>/dev/null; then
    slog "revert-ah: arrêt watchdog pid $wp"
    run "kill $wp 2>/dev/null || $SUDO kill $wp 2>/dev/null || true"
  fi
  run "rm -f $WATCHDOG_PIDFILE"
fi
run "rm -f $ARM_FLAG"
run "rm -f $DISARM_FLAG"

# sanity : plus rien ne doit écouter :1342, et host-main doit toujours répondre
if [ "$DRY" = 0 ]; then
  if host_health "$AH_HOST_PORT"; then slog "revert-ah: host-main sain sur $AH_HOST_PORT (intact)"; \
    else slog "revert-ah: NOTE host-main n'a pas répondu sur $AH_HOST_PORT (il n'a pas été touché par ce script)"; fi
fi
slog "revert-ah: FIN (dry=$DRY)"
exit 0
