#!/usr/bin/env bash
# arm.sh — BASCULE d'interception live des DEUX pattes. NE PAS exécuter à la légère.
# ============================================================================
#  ⚠️  CE SCRIPT PEUT COUPER L'EGRESS / BRICKER LA BOX (voir README, risque de
#      cert-pinning sur host-main). Il refuse de tourner sans le jeton de
#      confirmation explicite :  LI_ARM_CONFIRM=yes-i-understand-brick-risk
#      Sans ce jeton il se contente d'imprimer le plan (dry-run) et sort.
# ============================================================================
#
# Ce script n'arme PLUS la patte host<->backend. Plus de mitmproxy, plus de
# NODE_EXTRA_CA_CERTS, plus de redirect nft :443. Cette patte est désormais
# passive (keylog TLS) : hb-setup-keylog.sh puis hb-on.sh / hb-off.sh.
# Un arm.sh réel relance host-main SANS --tls-keylog : ne pas l'enchaîner après
# un hb-setup-keylog, sinon le superviseur (ou ce spawn) perd le flag. Pour les
# deux pattes, préférer arm-agent-host.sh (sans restart) + hb-setup-keylog.sh.
#
# Comment host-main est relancé sans que le superviseur n'en ponde un 2e :
#   le sand-supervisor ADOPTE tout process « host-main.cjs » vivant qu'il découvre
#   via gateway.json (maybeAdoptOrphanHost). On lance donc NOTRE host-main (il
#   écrit gateway.json), puis on SIGTERM l'ancien : le superviseur adopte le nôtre
#   au lieu d'en lancer un neuf. (Course possible : cf. README « risque d'adoption ».)
#
# Boucle anti-brick : revert.sh s'appuie lui sur l'auto-respawn PROPRE du
# superviseur (tuer host-main -> il en relance un sur 1340 avec l'env d'origine).
#
set -uo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
. "$HERE/lib.sh"

GRACE="${LI_ARM_GRACE:-90}"
CONFIRM="${LI_ARM_CONFIRM:-}"
DRY=1; [ "$CONFIRM" = "yes-i-understand-brick-risk" ] && DRY=0

plan() {
  cat <<PLAN
=== PLAN arm (DRY=$DRY) ===
 host-main courant : $(hostmain_pid || echo '??')  (uid $HOST_UID)
 1. mkdir $STATE_DIR ; rm DISARM ; touch ARM ; orig port=$HOST_PORT_ORIG
 2. watchdog.sh --grace $GRACE (détaché, AVANT toute coupure)
 3. (host<->backend NON armé ici — voir hb-setup-keylog.sh / hb-on.sh)
 4. spawn NOUVEAU host-main (uid box) env d'origine + SAND_HOST_PORT=$HOST_PORT_ARM
    (PAS de NODE_EXTRA_CA_CERTS, PAS de --tls-keylog)
    attendre /health sur $HOST_PORT_ARM
 5. SIGTERM ancien host-main (le superviseur adopte le nôtre via gateway.json)
 6. relai agent-host (root) sur :$RELAY_LISTEN -> 127.0.0.1:$HOST_PORT_ARM
 7. (plus de redirect nft :443 ; la table ip sandrpc n'est pas créée par ce script)
 Revert : bash revert.sh   |   Désarmer le dead-man : touch $DISARM_FLAG
PLAN
}

plan
if [ "$DRY" = 1 ]; then
  echo
  echo ">>> DRY-RUN : rien n'a été modifié. Pour armer réellement :"
  echo ">>>   LI_ARM_CONFIRM=yes-i-understand-brick-risk sudo -E bash arm.sh"
  exit 0
fi

need_root
OLD_PID="$(hostmain_pid)"
[ -z "$OLD_PID" ] && { slog "arm: pas de host-main courant, abandon"; exit 2; }

# 1) flags + watchdog AVANT toute coupure ------------------------------------
mkdir -p "$STATE_DIR"
rm -f "$DISARM_FLAG"
echo "$HOST_PORT_ORIG" > "$ORIG_PORT_FILE"
touch "$ARM_FLAG"
slog "arm: ARM posé"
bash "$HERE/watchdog.sh" --grace "$GRACE"
slog "arm: watchdog lancé (grace ${GRACE}s)"

# 2) host<->backend : plus de mitm. Rien à démarrer ici.
slog "arm: patte host<->backend NON armée (utiliser hb-setup-keylog.sh + hb-on.sh)"

# 3) capture l'env d'origine de host-main (null-séparé) et ajoute nos overrides
#    (secrets traversent la mémoire, jamais le disque/log). On retire une CA
#    résiduelle ; on n'ajoute pas SSLKEYLOGFILE (réservé à hb-setup-keylog).
mapfile -d '' -t OENV < "/proc/$OLD_PID/environ"
NEWENV=()
for kv in "${OENV[@]}"; do
  case "$kv" in
    SAND_HOST_PORT=*) continue;;
    NODE_EXTRA_CA_CERTS=*) continue;;
    SSLKEYLOGFILE=*) continue;;
    *) NEWENV+=("$kv");;
  esac
done
NEWENV+=("SAND_HOST_PORT=$HOST_PORT_ARM")

# 4) spawn NOTRE host-main (uid box) sur le nouveau port ----------------------
slog "arm: spawn host-main sur $HOST_PORT_ARM (sans CA, sans --tls-keylog)"
setsid env -i "${NEWENV[@]}" \
    "$NODE_BIN" "$NODE_FLAG" "$HOST_ENTRY" \
    </dev/null >>"$STATE_DIR/hostmain.out" 2>&1 &
echo $! > "$HOSTMAIN_SPAWN_PIDFILE"
for i in $(seq 1 20); do host_health "$HOST_PORT_ARM" && break; sleep 1; done
if ! host_health "$HOST_PORT_ARM"; then
  slog "arm: host-main NE répond PAS sur $HOST_PORT_ARM -> REVERT IMMÉDIAT"
  bash "$HERE/revert.sh"; exit 3
fi
slog "arm: host-main sain sur $HOST_PORT_ARM"

# 5) SIGTERM l'ancien host-main ; le superviseur adopte le nôtre (gateway.json)
slog "arm: SIGTERM ancien host-main $OLD_PID"
kill -TERM "$OLD_PID" 2>/dev/null || true
sleep 2

# 6) relai agent-host sur 1340 (root) ----------------------------------------
slog "arm: démarrage relai agent-host :$RELAY_LISTEN -> :$HOST_PORT_ARM"
RPC_TRAFFIC_LOG_DIR="$TRAFFIC_LOG_DIR" setsid python3 "$RELAY_PY" \
    --listen "$RELAY_LISTEN" --upstream-port "$HOST_PORT_ARM" \
    </dev/null >>"$STATE_DIR/relay.out" 2>&1 &
echo $! > "$RELAY_PIDFILE"
sleep 1

slog "arm: ARMÉ (agent<->host par déplacement de port). host<->backend non inclus."
slog "arm: Vérifie la capture puis: touch $DISARM_FLAG (confirmer) ou bash revert.sh (annuler)."
if backend_reachable; then slog "arm: backend joignable post-arm"; else
  slog "arm: ATTENTION backend injoignable post-arm -> le watchdog va revert"; fi
