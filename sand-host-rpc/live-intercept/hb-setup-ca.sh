#!/usr/bin/env bash
# hb-setup-ca.sh — rend host-main "CA-ready" en le relançant UNE fois avec notre CA
# en ADDITIF (NODE_EXTRA_CA_CERTS). C'est le SEUL moment à risque brick de toute la
# patte host<->backend (si host-main épingle le cert api2, l'egress tombe). Après
# ça, host-main fait confiance à notre CA EN PERMANENCE ; la CA est inoffensive
# tant qu'aucune règle nft ne redirige vers mitmproxy (hb-on). On/off ultérieurs =
# nft seule, SANS restart -> plus jamais ce risque.
# ============================================================================
#  ⚠️  OPÉRATEUR UNIQUEMENT. Refuse de tourner sans le jeton explicite :
#      LI_ARM_CONFIRM=yes-i-understand-brick-risk sudo -E bash hb-setup-ca.sh
#      Sans jeton -> DRY-RUN (imprime le plan, ne touche à RIEN).
#
#  Choix CA = NODE_EXTRA_CA_CERTS additif (cf. README §3) : AJOUTE notre CA aux
#  racines Node sans les remplacer -> confiance backend d'origine préservée. Lu au
#  démarrage de node -> injecté en relançant host-main avec SON env d'origine
#  (capturé via /proc/PID/environ, en MÉMOIRE, JAMAIS sur disque) + cette variable.
#
#  MÊME PORT (1340) : on relance host-main sur son port d'origine pour laisser la
#  patte agent<->host (relai sur :1342 -> 172.30.0.2:1340) INTACTE. On s'appuie sur
#  l'adoption du superviseur (maybeAdoptOrphanHost via gateway.json) : on SIGTERM
#  l'ancien et on spawn le nôtre (même port + CA) ; le superviseur adopte le nôtre
#  au lieu d'en pondre un sans CA. Course possible -> watchdog dead-man armé AVANT.
#
#  Dead-man DÉDIÉ (HB_ARM/HB_DISARM) : n'interfère PAS avec l'état agent<->host.
# ============================================================================
set -uo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
. "$HERE/lib.sh"

GRACE="${LI_ARM_GRACE:-90}"
CONFIRM="${LI_ARM_CONFIRM:-}"
DRY=1; [ "$CONFIRM" = "yes-i-understand-brick-risk" ] && DRY=0
REVERT="$HERE/hb-revert.sh"

plan() {
  cat <<PLAN
=== PLAN hb-setup-ca (DRY=$DRY) — relance host-main CA-ready, MÊME port 1340 ===
 host-main courant : $(hostmain_pid || echo '??')  (uid $HB_UID, port $HOST_PORT_ORIG)
 CA additive       : $MITM_CA  (NODE_EXTRA_CA_CERTS, additif)
 1. mkdir $STATE_DIR ; rm $HB_DISARM_FLAG ; touch $HB_ARM_FLAG
 2. watchdog.sh --grace $GRACE --revert $REVERT --arm-flag $HB_ARM_FLAG --disarm-flag $HB_DISARM_FLAG  (détaché, AVANT le restart)
 3. capture env host-main via /proc/PID/environ (mémoire) ; NEWENV = env - NODE_EXTRA_CA_CERTS + NODE_EXTRA_CA_CERTS=$MITM_CA (port 1340 inchangé)
 4. SIGTERM ancien host-main ; spawn NOTRE host-main (même port 1340 + CA) ; attendre /health
 5. sanity backend joignable ; si OK -> touch $HB_CA_READY_FLAG (host-main CA-ready)
 Confirmer (annuler le dead-man) : touch $HB_DISARM_FLAG     Annuler : bash $REVERT
 NB : NE pose AUCUNE règle nft. La CA seule ne déchiffre RIEN. hb-on.sh pose la règle ensuite.
PLAN
}

plan
if [ "$DRY" = 1 ]; then
  echo
  echo ">>> DRY-RUN : rien modifié, host-main intact, pas de CA injectée, pas de règle nft."
  echo ">>> Pour exécuter réellement (OPÉRATEUR) :"
  echo ">>>   LI_ARM_CONFIRM=yes-i-understand-brick-risk sudo -E bash hb-setup-ca.sh"
  exit 0
fi

need_root
OLD_PID="$(hostmain_pid)"
[ -z "$OLD_PID" ] && { slog "hb-setup-ca: pas de host-main courant, abandon"; exit 2; }
[ -f "$MITM_CA" ] || { slog "hb-setup-ca: CA absente ($MITM_CA), abandon"; exit 4; }

# 1) dead-man dédié AVANT le restart ------------------------------------------
mkdir -p "$STATE_DIR"
rm -f "$HB_DISARM_FLAG"
touch "$HB_ARM_FLAG"
slog "hb-setup-ca: HB_ARM posé (host-main $OLD_PID, port $HOST_PORT_ORIG)"
bash "$HERE/watchdog.sh" --grace "$GRACE" --revert "$REVERT" \
     --arm-flag "$HB_ARM_FLAG" --disarm-flag "$HB_DISARM_FLAG"
slog "hb-setup-ca: watchdog dédié lancé (grace ${GRACE}s)"

# 2) capture env d'origine (null-séparé, MÉMOIRE uniquement) + overrides CA -----
mapfile -d '' -t OENV < "/proc/$OLD_PID/environ"
NEWENV=()
for kv in "${OENV[@]}"; do
  case "$kv" in
    NODE_EXTRA_CA_CERTS=*) continue;;   # on force la nôtre (additive)
    *) NEWENV+=("$kv");;
  esac
done
NEWENV+=("NODE_EXTRA_CA_CERTS=$MITM_CA")   # port 1340 CONSERVÉ (agent-host intact)

# 3) SIGTERM ancien + spawn le nôtre (même port, +CA) -> adoption superviseur ---
slog "hb-setup-ca: SIGTERM ancien host-main $OLD_PID puis spawn CA-ready (port $HOST_PORT_ORIG)"
kill -TERM "$OLD_PID" 2>/dev/null || true
setsid env -i "${NEWENV[@]}" \
    "$NODE_BIN" "$NODE_FLAG" "$HOST_ENTRY" \
    </dev/null >>"$STATE_DIR/hb-hostmain.out" 2>&1 &
echo $! > "$HB_HOSTMAIN_PIDFILE"
for i in $(seq 1 30); do host_health "$HOST_PORT_ORIG" && break; sleep 1; done
if ! host_health "$HOST_PORT_ORIG"; then
  slog "hb-setup-ca: host-main NE répond PAS sur $HOST_PORT_ORIG -> REVERT"
  bash "$REVERT"; exit 3
fi
slog "hb-setup-ca: host-main sain sur $HOST_PORT_ORIG"

# 4) sanity egress + marqueur CA-ready ----------------------------------------
if backend_reachable; then
  touch "$HB_CA_READY_FLAG"
  slog "hb-setup-ca: backend joignable -> host-main CA-READY. touch $HB_DISARM_FLAG pour annuler le dead-man."
  slog "hb-setup-ca: ACTIVER la capture host<->backend = bash hb-on.sh (nft seule, sans restart)."
else
  slog "hb-setup-ca: ATTENTION backend injoignable post-restart -> le watchdog va revert"
fi
