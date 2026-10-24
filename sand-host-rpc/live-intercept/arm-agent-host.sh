#!/usr/bin/env bash
# arm-agent-host.sh — interception de la SEULE patte agent<->host, SANS redémarrer
# host-main. Variante la moins intrusive : on NE touche PAS host-main, PAS de CA,
# PAS de mitm, PAS de redirection 443.
# ============================================================================
#  Principe :
#   * host-main RÉEL reste intact, écoute 0.0.0.0:1340 (prouvé plaintext HTTP/1.1).
#   * on lance un relai loggeur devant lui : :1342 -> 172.30.0.2:1340.
#   * nft en PREROUTING redirige UNIQUEMENT le trafic du client agent
#     (saddr 172.30.0.1) à destination de 172.30.0.2:1340 vers :1342.
#
#  Pourquoi PAS de boucle (vérifié sur cette box) :
#   - 172.30.0.2 est une adresse LOCALE (enp0s4) ; la connexion sortante du relai
#     vers 172.30.0.2:1340 part avec saddr=172.30.0.2 (route via lo), JAMAIS
#     172.30.0.1 -> la règle (saddr 172.30.0.1) ne la matche pas.
#   - de plus un paquet localement généré vers une adresse locale passe par OUTPUT,
#     pas par PREROUTING : la connexion du relai ne traverse même pas la chaîne.
#   - skuid 0 est INOPÉRANT en prerouting (paquet entrant d'un autre netns = pas de
#     socket propriétaire) : c'est le match saddr qui protège, pas skuid. On ajoute
#     en défense le scoping `iifname` de l'iface agent.
#
#  Risque : quasi nul (on ne restart pas host-main, on ne touche pas l'egress 443).
#  Le seul effet de bord possible = la redirection nft ; le watchdog la retire
#  tout seul après GRACE s sans DISARM, et revert-agent-host.sh la retire à la main.
# ============================================================================
#  Dry-run par défaut. Pour armer réellement :
#    LI_ARM_CONFIRM=yes-i-understand-brick-risk sudo -E bash arm-agent-host.sh
# ============================================================================
set -uo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
. "$HERE/lib.sh"

GRACE="${LI_ARM_GRACE:-90}"
CONFIRM="${LI_ARM_CONFIRM:-}"
DRY=1; [ "$CONFIRM" = "yes-i-understand-brick-risk" ] && DRY=0
REVERT="$HERE/revert-agent-host.sh"

plan() {
  cat <<PLAN
=== PLAN arm-agent-host (DRY=$DRY) — SANS restart host-main ===
 host-main RÉEL (intact)   : $(hostmain_pid || echo '??')  écoute ${AH_HOST_DADDR}:${AH_HOST_PORT}
 1. mkdir $STATE_DIR ; rm DISARM ; touch ARM
 2. watchdog.sh --grace $GRACE --revert $REVERT   (détaché, AVANT toute règle nft)
 3. relai loggeur (root) : bind ${AH_RELAY_BIND}:${AH_RELAY_PORT} -> ${AH_HOST_DADDR}:${AH_HOST_PORT}  leg=agent-host
 4. nft : table ip ${AH_NFT_TABLE}
         chain prerouting { type nat hook prerouting priority -100 ; policy accept ; }
         rule iifname "${AH_IIF}" ip saddr ${AH_AGENT_SADDR} ip daddr ${AH_HOST_DADDR} tcp dport ${AH_HOST_PORT} redirect to :${AH_RELAY_PORT}
    (match saddr agent SEULEMENT -> pas de boucle ; host-main NON redémarré)
 Confirmer (garder) : touch $DISARM_FLAG   |   Annuler : bash $REVERT
PLAN
}

plan
if [ "$DRY" = 1 ]; then
  echo
  echo ">>> DRY-RUN : rien n'a été modifié, host-main intact, aucune règle nft."
  echo ">>> Pour armer réellement :"
  echo ">>>   LI_ARM_CONFIRM=yes-i-understand-brick-risk sudo -E bash arm-agent-host.sh"
  exit 0
fi

need_root
HP="$(hostmain_pid)"
[ -z "$HP" ] && { slog "arm-ah: pas de host-main courant, abandon (on ne crée rien)"; exit 2; }

# 1) flags + watchdog AVANT la moindre règle nft ------------------------------
mkdir -p "$STATE_DIR"
rm -f "$DISARM_FLAG"
touch "$ARM_FLAG"
slog "arm-ah: ARM posé (host-main $HP intact sur ${AH_HOST_DADDR}:${AH_HOST_PORT})"
bash "$HERE/watchdog.sh" --grace "$GRACE" --revert "$REVERT"
slog "arm-ah: watchdog lancé (grace ${GRACE}s, revert=$REVERT)"

# 2) relai loggeur devant le host-main RÉEL (on ne le redémarre PAS) ----------
slog "arm-ah: démarrage relai ${AH_RELAY_BIND}:${AH_RELAY_PORT} -> ${AH_HOST_DADDR}:${AH_HOST_PORT}"
RPC_TRAFFIC_LOG_DIR="$TRAFFIC_LOG_DIR" setsid python3 "$RELAY_PY" \
    --listen "${AH_RELAY_BIND}:${AH_RELAY_PORT}" \
    --upstream "${AH_HOST_DADDR}:${AH_HOST_PORT}" \
    </dev/null >>"$STATE_DIR/relay-agent-host.out" 2>&1 &
echo $! > "$AH_RELAY_PIDFILE"
sleep 1
if ! kill -0 "$(cat "$AH_RELAY_PIDFILE" 2>/dev/null)" 2>/dev/null; then
  slog "arm-ah: le relai n'a pas démarré -> REVERT"; bash "$REVERT"; exit 3
fi
slog "arm-ah: relai vivant pid $(cat "$AH_RELAY_PIDFILE")"

# 3) nft prerouting, UNIQUEMENT le client agent -------------------------------
slog "arm-ah: installation redirection nft table ip $AH_NFT_TABLE (prerouting, saddr $AH_AGENT_SADDR seulement)"
$SUDO nft add table ip "$AH_NFT_TABLE"
$SUDO nft add chain ip "$AH_NFT_TABLE" prerouting '{ type nat hook prerouting priority -100 ; policy accept ; }'
$SUDO nft add rule ip "$AH_NFT_TABLE" prerouting iifname "$AH_IIF" \
    ip saddr "$AH_AGENT_SADDR" ip daddr "$AH_HOST_DADDR" tcp dport "$AH_HOST_PORT" \
    redirect to :"$AH_RELAY_PORT"

slog "arm-ah: ARMÉ (agent<->host). host-main NON touché."
slog "arm-ah: vérifie la capture puis: touch $DISARM_FLAG (confirmer) ou bash $REVERT (annuler)."
if backend_reachable; then slog "arm-ah: backend joignable (sanity ; accessoire pour cette patte)"; fi
