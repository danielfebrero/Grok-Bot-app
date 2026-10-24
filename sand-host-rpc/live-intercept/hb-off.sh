#!/usr/bin/env bash
# hb-off.sh — DÉSACTIVE la capture host<->backend. INSTANTANÉ, SANS restart.
# Supprime la SEULE règle nft dédiée (table ip sandrpc_hb). host-main reste
# CA-ready (inoffensif sans règle) -> un futur hb-on.sh ré-active sans restart.
# Laisse mitmproxy tourner (idle, ne capture rien sans la règle) ; utilise
# hb-revert.sh pour un teardown complet.
#
# NE TOUCHE JAMAIS la table ip sandrpc (patte agent<->host). Idempotent. --dry-run.
# ============================================================================
set -uo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
. "$HERE/lib.sh"
DRY=0; [ "${1:-}" = "--dry-run" ] && DRY=1
run() { if [ "$DRY" = 1 ]; then echo "DRY: $*"; else eval "$@"; fi; }

need_root
slog "hb-off: début (dry=$DRY) — supprime UNIQUEMENT la table ip $HB_NFT_TABLE"
if $SUDO nft list table ip "$HB_NFT_TABLE" >/dev/null 2>&1; then
  slog "hb-off: suppression table ip $HB_NFT_TABLE (capture host<->backend OFF)"
  run "$SUDO nft delete table ip $HB_NFT_TABLE"
else
  slog "hb-off: table ip $HB_NFT_TABLE absente (déjà OFF)"
fi
# garde-fou : on ne touche surtout pas à sandrpc (agent<->host)
if [ "$DRY" = 0 ] && $SUDO nft list table ip sandrpc >/dev/null 2>&1; then
  slog "hb-off: table ip sandrpc (agent<->host) intacte, non touchée."
fi
slog "hb-off: FIN (dry=$DRY). mitm laissé en vie (idle). Teardown complet: bash hb-revert.sh"
exit 0
