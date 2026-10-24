#!/usr/bin/env bash
# revert.sh — restaure TOUT l'état d'origine de l'interception live.
#
# Idempotent. Marche même appelé seul (sans arm préalable). Exit 0 si, à la fin,
# host-main écoute sur le port d'origine (1340) ET le backend est joignable.
#
# Étapes (ordre important pour éviter un conflit de port sur 1340) :
#   1. supprime la table/chaîne netfilter dédiée (retire la redirection host<->backend)
#   2. arrête le relai agent-host (libère le port 1340)  et le mitm transparent
#   3. restaure le port host (retire le fichier d'état de port)
#   4. tue le host-main éventuellement lancé par arm (sur 1341) ->
#      le sand-supervisor le RELANCE automatiquement, propre, sur 1340, avec SON
#      env d'origine (ni NODE_EXTRA_CA_CERTS, ni SSLKEYLOGFILE, ni --tls-keylog,
#      ni port modifié). C'est le mécanisme
#      de restauration : on s'appuie sur l'auto-respawn du superviseur.
#   5. attend la santé de host-main sur 1340 + joignabilité backend.
#
# Usage : revert.sh [--dry-run]
set -uo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
. "$HERE/lib.sh"

DRY=0; [ "${1:-}" = "--dry-run" ] && DRY=1
run() { if [ "$DRY" = 1 ]; then echo "DRY: $*"; else eval "$@"; fi; }

need_root
slog "revert: début (dry=$DRY)"

# 1) netfilter : delete la table dédiée (nft, backend actif). no-op si absente.
if $SUDO nft list table ip "$NFT_TABLE" >/dev/null 2>&1; then
  slog "revert: suppression table ip $NFT_TABLE"
  run "$SUDO nft delete table ip $NFT_TABLE"
else
  slog "revert: table ip $NFT_TABLE absente (ok)"
fi

# 2) interceptors : relai + éventuel mitm résiduel + capture keylog host<->backend
for pf in "$RELAY_PIDFILE" "$MITM_PIDFILE" "$HB_TCPDUMP_PIDFILE" "$HB_DECODE_PIDFILE"; do
  if [ -f "$pf" ]; then
    p="$(cat "$pf" 2>/dev/null)"
    if [ -n "$p" ] && kill -0 "$p" 2>/dev/null; then
      slog "revert: kill interceptor pid $p ($pf)"
      run "kill $p 2>/dev/null || $SUDO kill $p 2>/dev/null || true"
    fi
    run "rm -f $pf"
  fi
done
# ceinture+bretelles : plus rien qui écoute 1340 à part host-main
if [ "$DRY" = 0 ]; then
  for p in $(pgrep -f 'relay_agent_host\.py' 2>/dev/null); do
    slog "revert: kill relai résiduel pid $p"; kill "$p" 2>/dev/null || $SUDO kill "$p" 2>/dev/null || true
  done
fi

# 3) restaure le marqueur de port d'origine
run "rm -f $ORIG_PORT_FILE"

# 4) bounce host-main pour forcer le respawn PROPRE par le superviseur.
#    On tue le host-main courant (qu'il soit l'arm-spawn sur 1341 ou l'original) ;
#    le sand-supervisor détecte la mort et relance host-main sur 1340 avec son env.
hp="$(hostmain_pid)"
if [ -n "$hp" ]; then
  slog "revert: SIGTERM host-main pid $hp (le superviseur relancera un host propre sur 1340)"
  run "kill -TERM $hp 2>/dev/null || $SUDO kill -TERM $hp 2>/dev/null || true"
else
  slog "revert: aucun host-main vivant ; le superviseur va en lancer un"
fi
run "rm -f $HOSTMAIN_SPAWN_PIDFILE"
# secrets de session TLS : seulement après le SIGTERM, sinon le process les réécrit.
if [ "$DRY" = 1 ]; then
  echo "DRY: shred $HB_PCAP $HB_KEYLOG_FILE ; rm $HB_KEYLOG_READY_FLAG"
else
  sleep 1
  hb_shred "$HB_PCAP"
  hb_shred "$HB_KEYLOG_FILE"
  rm -f "$HB_KEYLOG_READY_FLAG" "$HB_CAPTURE_READY" "$HB_SEEN_FILE" "$HB_CA_READY_FLAG"
fi

# 5) attente de restauration : health sur 1340 + backend joignable
if [ "$DRY" = 1 ]; then
  slog "revert(dry): sauterait l'attente de santé"; slog "revert: FIN (dry)"; exit 0
fi
ok=0
for i in $(seq 1 30); do
  if host_health "$HOST_PORT_ORIG"; then ok=1; break; fi
  sleep 2
done
if [ "$ok" = 1 ]; then
  slog "revert: host-main OK sur $HOST_PORT_ORIG"
else
  slog "revert: ATTENTION host-main pas (encore) sain sur $HOST_PORT_ORIG après 60s"
fi
if backend_reachable; then slog "revert: backend joignable"; else slog "revert: ATTENTION backend injoignable"; fi

# disarm le watchdog : on a fini de restaurer
run "touch $DISARM_FLAG"
slog "revert: DISARM posé ; FIN"
[ "$ok" = 1 ] && exit 0 || exit 1
