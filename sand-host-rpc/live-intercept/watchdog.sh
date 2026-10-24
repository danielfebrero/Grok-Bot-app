#!/usr/bin/env bash
# watchdog.sh — dead-man's switch DÉTACHÉ (setsid) qui SURVIT à la mort de l'agent.
#
# Rôle : après un arm, surveiller que le backend reste joignable. Si l'egress casse
# (le risque de pinning), ou si l'opérateur n'a pas confirmé l'arm à temps, il
# exécute revert.sh TOUT SEUL -> la box se répare même si l'agent est mort.
#
# Sémantique des flags (dans STATE_DIR, hors dépôt) :
#   ARM     : posé par arm.sh. Le watchdog ne fait rien tant qu'il est absent.
#   DISARM  : posé par l'opérateur (`touch`) = « je confirme, annule le dead-man,
#             garde l'interception ». Le watchdog SORT sans revert. Le teardown
#             manuel se fait alors via revert.sh (qui pose aussi DISARM).
#
# Déclencheurs de revert automatique (tant que pas de DISARM) :
#   * backend injoignable pendant --probes probes consécutives, OU
#   * --grace secondes écoulées depuis le démarrage sans DISARM (dead-man).
#
# Usage :
#   watchdog.sh [--grace SEC(def 90)] [--interval SEC(def 5)] [--probes N(def 3)]
# Se ré-exécute en setsid détaché au premier appel.
#
# Exit codes : 0 = sortie propre (DISARM ou ARM retiré) ;
#              10 = revert déclenché (backend KO) ; 11 = revert déclenché (grace).
set -uo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
. "$HERE/lib.sh"

GRACE=90; INTERVAL=5; PROBES=3
REVERT="$HERE/revert.sh"   # script de revert (surchargeable: variante agent-host)
# flags surchargeables (def = ARM/DISARM partagés) : permet un dead-man DÉDIÉ
# (ex. host<->backend : --arm-flag HB_ARM --disarm-flag HB_DISARM) qui n'interfère
# PAS avec l'état de la patte agent<->host.
ARMF="$ARM_FLAG"; DISARMF="$DISARM_FLAG"
while [ $# -gt 0 ]; do
  case "$1" in
    --grace) GRACE="$2"; shift 2;;
    --interval) INTERVAL="$2"; shift 2;;
    --probes) PROBES="$2"; shift 2;;
    --revert) REVERT="$2"; shift 2;;
    --arm-flag) ARMF="$2"; shift 2;;
    --disarm-flag) DISARMF="$2"; shift 2;;
    *) shift;;
  esac
done

# auto-détachement : se relancer sous setsid, découplé de l'agent, puis sortir.
if [ "${LI_WATCHDOG_DETACHED:-0}" != "1" ]; then
  mkdir -p "$STATE_DIR"
  LI_WATCHDOG_DETACHED=1 setsid bash "$HERE/watchdog.sh" \
      --grace "$GRACE" --interval "$INTERVAL" --probes "$PROBES" --revert "$REVERT" \
      --arm-flag "$ARMF" --disarm-flag "$DISARMF" \
      </dev/null >>"$WATCHDOG_LOG" 2>&1 &
  echo "watchdog détaché (setsid) ; log: $WATCHDOG_LOG"
  exit 0
fi

# ---- corps détaché ----------------------------------------------------------
echo $$ > "$WATCHDOG_PIDFILE"
wlog() { echo "[$(ts)] watchdog: $*"; }
wlog "démarré (grace=${GRACE}s interval=${INTERVAL}s probes=${PROBES}) pid=$$"

START=$(date +%s)
fail=0
do_revert() {
  local why="$1" code="$2"
  wlog "DÉCLENCHEMENT revert ($why) -> exécution de $REVERT"
  bash "$REVERT" >>"$WATCHDOG_LOG" 2>&1
  wlog "revert terminé (exit $?)"
  rm -f "$WATCHDOG_PIDFILE"
  exit "$code"
}

while true; do
  # ARM retiré -> plus rien à surveiller
  if [ ! -f "$ARMF" ]; then
    wlog "ARM absent ; sortie propre"; rm -f "$WATCHDOG_PIDFILE"; exit 0
  fi
  # DISARM = opérateur confirme ; on arrête de surveiller SANS revert
  if [ -f "$DISARMF" ]; then
    wlog "DISARM détecté ; le dead-man est annulé (interception conservée) ; sortie"
    rm -f "$WATCHDOG_PIDFILE"; exit 0
  fi
  # probe backend
  if backend_reachable; then
    [ "$fail" -gt 0 ] && wlog "backend de nouveau joignable (reset)"
    fail=0
  else
    fail=$((fail+1))
    wlog "backend injoignable (échec $fail/$PROBES)"
    if [ "$fail" -ge "$PROBES" ]; then
      do_revert "backend injoignable ${PROBES} probes" 10
    fi
  fi
  # dead-man : grace dépassée sans DISARM
  now=$(date +%s)
  if [ $((now - START)) -ge "$GRACE" ]; then
    do_revert "grace ${GRACE}s dépassée sans DISARM" 11
  fi
  sleep "$INTERVAL"
done
