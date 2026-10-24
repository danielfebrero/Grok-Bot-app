#!/usr/bin/env bash
# cdp-off.sh — arrête la capture CDP et shred le pcap.
# Ne touche pas host-main, ni nft, ni le keylog / la capture tcp/443.
# Idempotent. --dry-run n'exige pas root.
# ============================================================================
set -uo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# shellcheck source=lib.sh
. "$HERE/lib.sh"
DRY=0; [ "${1:-}" = "--dry-run" ] && DRY=1

if [ "$DRY" = 1 ]; then
  echo "DRY cdp-off: SIGTERM tcpdump ($CDP_TCPDUMP_PIDFILE) puis décodeur ($CDP_DECODE_PIDFILE)"
  echo "DRY cdp-off: shred pcap $CDP_PCAP"
  echo "DRY cdp-off: NE PAS toucher host-main, nft, ni le keylog 443"
  exit 0
fi

need_root
slog "cdp-off: arrêt capture CDP"
cdp_stop_pids
hb_shred "$CDP_PCAP"
slog "cdp-off: FIN. host-main et la capture 443 non touchés."
exit 0
