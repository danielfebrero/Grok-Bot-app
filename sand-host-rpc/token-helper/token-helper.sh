#!/usr/bin/env bash
# token-helper — wrapper autour du helper officiel get-inference-token.py.
# Renouvelle le jeton d'inférence du sand-host : lit SAND_INFERENCE_RENEWAL_CREDENTIAL
# depuis l'env du process sand-host, POST api2.cursor.sh/sand-box/inference-credential,
# imprime UNIQUEMENT le jeton. Ne logge JAMAIS la credential ni le jeton.
#
# Exemples :
#   eval "$(./token-helper.sh --export)"          # exporte GROK_BOT_RPC_AUTH / _BASE_URL
#   TOKEN=$(./token-helper.sh)                     # jeton sur stdout
#   ./token-helper.sh --out /tmp/tok && chmod 600 /tmp/tok
set -euo pipefail
DIR="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"
exec python3 "$DIR/get-inference-token.py" "$@"
