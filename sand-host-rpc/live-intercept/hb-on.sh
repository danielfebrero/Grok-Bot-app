#!/usr/bin/env bash
# hb-on.sh — ACTIVE la capture host<->backend. INSTANTANÉ, SANS restart host-main.
# Pose la SEULE règle nft (table dédiée ip sandrpc_hb) qui redirige le 443 de
# host-main (uid 1000) vers mitmproxy transparent sélectif (:8900), SAUF le tunnel
# egress (198.18.0.1) laissé en direct. Démarre mitmproxy s'il ne tourne pas.
# ============================================================================
#  Pré-requis : host-main doit être CA-ready (hb-setup-ca.sh exécuté) sinon le
#  déchiffrement casse api2. hb-on REFUSE si le marqueur CA-ready est absent
#  (contournable par HB_FORCE=1, déconseillé).
#
#  Règle (chain output, nat hook, le paquet a un socket -> skuid opérant) :
#     meta skuid 1000 ip daddr != 198.18.0.1 tcp dport 443 redirect to :8900
#  Anti-boucle : mitmproxy tourne en ROOT (uid 0) ; sa reconnexion au vrai backend
#  (uid 0) n'est PAS reprise par 'skuid 1000'. Anti-collatéral : 'ip daddr !=
#  198.18.0.1' laisse le tunnel egress (donc les Chrome agents) en direct ; et
#  l'interception SÉLECTIVE de mitmproxy (--allow-hosts) ne DÉCHIFFRE que les hôtes
#  backend ciblés, PASSTHROUGH tout le reste (S3/CDN) -> zéro casse des autres 443.
#
#  Table SÉPARÉE (sandrpc_hb) : n'affecte JAMAIS la table agent<->host (ip sandrpc).
#
#  Dry-run par défaut. Activer : HB_ON_CONFIRM=yes sudo -E bash hb-on.sh
#  (ou LI_ARM_CONFIRM=yes-i-understand-brick-risk). OFF instantané : bash hb-off.sh
# ============================================================================
set -uo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
. "$HERE/lib.sh"
CONFIRM="${HB_ON_CONFIRM:-}"; [ "${LI_ARM_CONFIRM:-}" = "yes-i-understand-brick-risk" ] && CONFIRM=yes
DRY=1; [ "$CONFIRM" = "yes" ] && DRY=0

plan() {
  cat <<PLAN
=== PLAN hb-on (DRY=$DRY) — SANS restart host-main ===
 1. (si besoin) démarrer mitm transparent sélectif root sur :$HB_MITM_PORT
      intercept=$HB_ALLOW_HOSTS ; passthrough=tout le reste (tunnel $HB_TUNNEL_ADDR, S3, CDN)
 2. nft : table ip $HB_NFT_TABLE (DISTINCTE de sandrpc/agent-host)
         chain output { type nat hook output priority -100 ; policy accept ; }
         rule meta skuid $HB_UID ip daddr != $HB_TUNNEL_ADDR tcp dport 443 redirect to :$HB_MITM_PORT
 OFF instantané : bash hb-off.sh   |   Revert complet : bash hb-revert.sh
PLAN
}
plan
if [ "$DRY" = 1 ]; then
  echo; echo ">>> DRY-RUN : aucune règle posée, mitm non démarré, host-main intact."
  echo ">>> Activer : HB_ON_CONFIRM=yes sudo -E bash hb-on.sh"
  exit 0
fi

need_root
if [ ! -f "$HB_CA_READY_FLAG" ] && [ "${HB_FORCE:-0}" != "1" ]; then
  slog "hb-on: host-main PAS CA-ready ($HB_CA_READY_FLAG absent). Lance hb-setup-ca.sh d'abord (ou HB_FORCE=1)."; exit 5
fi

# 1) mitm transparent sélectif (root) -----------------------------------------
if hb_mitm_up; then
  slog "hb-on: mitm déjà en écoute sur :$HB_MITM_PORT"
else
  slog "hb-on: démarrage mitm transparent sélectif :$HB_MITM_PORT"
  setsid bash "$HERE/start-mitm-host-backend.sh" "$HB_MITM_PORT" \
      </dev/null >>"$HB_MITM_OUT" 2>&1 &
  echo $! > "$HB_MITM_PIDFILE"
  for i in $(seq 1 15); do hb_mitm_up && break; sleep 1; done
  if ! hb_mitm_up; then slog "hb-on: mitm n'a pas démarré sur :$HB_MITM_PORT -> abandon (aucune règle posée)"; exit 6; fi
  slog "hb-on: mitm vivant pid $(cat "$HB_MITM_PIDFILE") sur :$HB_MITM_PORT"
fi

# 2) nft : règle de redirection (table dédiée) --------------------------------
if hb_rule_present; then
  slog "hb-on: règle déjà posée (idempotent) ; rien à faire."
else
  slog "hb-on: pose règle nft table ip $HB_NFT_TABLE (output, skuid $HB_UID, daddr!=$HB_TUNNEL_ADDR, dport 443 -> :$HB_MITM_PORT)"
  $SUDO nft add table ip "$HB_NFT_TABLE"
  $SUDO nft add chain ip "$HB_NFT_TABLE" output '{ type nat hook output priority -100 ; policy accept ; }'
  $SUDO nft add rule ip "$HB_NFT_TABLE" output meta skuid "$HB_UID" ip daddr != "$HB_TUNNEL_ADDR" tcp dport 443 redirect to :"$HB_MITM_PORT"
fi
slog "hb-on: ACTIVÉ (host<->backend). host-main NON redémarré. OFF = bash hb-off.sh"
if backend_reachable; then slog "hb-on: backend joignable post-on"; else slog "hb-on: ATTENTION backend injoignable -> bash hb-off.sh"; fi
