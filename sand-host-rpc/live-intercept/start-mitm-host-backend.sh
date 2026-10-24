#!/usr/bin/env bash
# start-mitm-host-backend.sh — mitmproxy TRANSPARENT à INTERCEPTION SÉLECTIVE pour
# la patte host<->backend. Réutilise l'addon de log/rédaction existant
# (../rpc-traffic-monitor/mitmproxy/addon_log.py, leg="host-backend").
#
# ============================================================================
#  INTERCEPTION SÉLECTIVE (= le "tunnel alternatif" demandé)
# ----------------------------------------------------------------------------
#  --allow-hosts "$HB_ALLOW_HOSTS" : mitmproxy ne TERMINE le TLS (avec NOTRE CA)
#  que pour les hôtes host<->backend ciblés (api2/api3/api4/api/repo42.cursor.sh,
#  cursor.com, *.cursor.sh). TOUT le reste est en PASSTHROUGH TLS BRUT (tunnel TCP,
#  jamais déchiffré, jamais loggé) :
#     * 198.18.0.1  (sand-egress-tunnel  -> trafic des Chrome agents)  : passthrough
#     * S3  (*.amazonaws.com, 16.15.x / 52.217.x)                      : passthrough
#     * cloudflare (104.18.x)                                          : passthrough
#     * tout autre hôte non listé                                      : passthrough
#  => les autres clients 443 de la box ne sont JAMAIS déchiffrés et continuent de
#     fonctionner normalement (le flux passe à travers mitmproxy intact, chiffré).
#
#  En transparent, mitmproxy récupère la destination d'origine (SO_ORIGINAL_DST,
#  posée par la règle nft de hb-on.sh) et lit le SNI du ClientHello pour décider
#  intercept vs passthrough. Les hôtes DÉCHIFFRÉS exigent que host-main fasse
#  confiance à notre CA -> c'est le rôle (unique et isolé) de hb-setup-ca.sh.
#
#  Doit tourner en ROOT (uid 0) : son propre trafic sortant (uid 0) vers le vrai
#  backend n'est PAS repris par la règle nft 'meta skuid 1000' => pas de boucle.
# ============================================================================
set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
. "$HERE/lib.sh"
PORT="${1:-$HB_MITM_PORT}"
mkdir -p "$STATE_DIR"
[ "$(id -u)" = "0" ] || { echo "doit tourner en root"; exit 90; }
[ -f "$MITM_CA" ] || { echo "ERREUR: CA mitm absente: $MITM_CA (lance setup-ca)"; exit 1; }
echo "mitm transparent SÉLECTIF -> :$PORT"
echo "  intercept (déchiffré, loggé) : $HB_ALLOW_HOSTS"
echo "  passthrough (brut, non loggé): tout le reste (tunnel $HB_TUNNEL_ADDR, S3, CDN, ...)"
echo "  addon=$ADDON  logs=$TRAFFIC_LOG_DIR"
exec env RPC_TRAFFIC_LOG_DIR="$TRAFFIC_LOG_DIR" "$MITMDUMP" \
    --mode transparent --listen-host 127.0.0.1 --listen-port "$PORT" \
    --allow-hosts "$HB_ALLOW_HOSTS" \
    --set block_global=false --ssl-insecure -s "$ADDON"
