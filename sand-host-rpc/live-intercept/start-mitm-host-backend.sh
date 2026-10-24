#!/usr/bin/env bash
# start-mitm-host-backend.sh — lance mitmproxy en mode TRANSPARENT pour la patte
# host<->backend, réutilisant l'addon de log/rédaction existant (leg="host-backend").
#
# À utiliser SEULEMENT dans le cadre de arm.sh (qui pose la redirection nft des
# connexions uid-box:443 vers ce port). En transparent, mitmproxy récupère la
# destination d'origine (SO_ORIGINAL_DST) et ré-origine lui-même vers le vrai
# backend : il faut donc que host-main fasse confiance à notre CA (NODE_EXTRA_CA_CERTS).
#
# DOIT tourner en ROOT (uid 0) : ainsi son propre trafic sortant (uid 0) n'est PAS
# repris par la règle nft 'meta skuid 1000' -> pas de boucle d'interception.
set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
. "$HERE/lib.sh"
PORT="${1:-$MITM_TRANSPARENT_PORT}"
mkdir -p "$STATE_DIR"
[ "$(id -u)" = "0" ] || { echo "doit tourner en root"; exit 90; }
echo "mitm transparent -> :$PORT  addon=$ADDON  logs=$TRAFFIC_LOG_DIR"
exec env RPC_TRAFFIC_LOG_DIR="$TRAFFIC_LOG_DIR" "$MITMDUMP" \
    --mode transparent --listen-host 127.0.0.1 --listen-port "$PORT" \
    --set block_global=false --ssl-insecure -s "$ADDON"
