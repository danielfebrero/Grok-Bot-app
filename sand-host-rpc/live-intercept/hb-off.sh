#!/usr/bin/env bash
# hb-off.sh — DÉSACTIVE la capture host↔backend. Sans restart host-main.
#
# SIGTERM tcpdump, laisse le décodeur faire une dernière passe, détruit le pcap.
# Le keylog et le marqueur keylog-ready RESTENT : host-main continue d'écrire
# des secrets de session (fichier 0600) et un hb-on ultérieur n'a pas besoin
# d'un restart. Pour retirer le keylog : hb-revert.sh.
#
# Supprime aussi, s'il existe, le reliquat mitm (table ip sandrpc_hb + mitmdump).
# Ne touche JAMAIS la table ip sandrpc (agent↔host).
#
# Idempotent. --dry-run n'exige pas root et n'agit pas.
# ============================================================================
set -uo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# shellcheck source=lib.sh
. "$HERE/lib.sh"
DRY=0; [ "${1:-}" = "--dry-run" ] && DRY=1

if [ "$DRY" = 1 ]; then
  echo "DRY hb-off: SIGTERM tcpdump ($HB_TCPDUMP_PIDFILE) puis décodeur ($HB_DECODE_PIDFILE)"
  echo "DRY hb-off: shred pcap $HB_PCAP ; conserver keylog $HB_KEYLOG_FILE et $HB_KEYLOG_READY_FLAG"
  echo "DRY hb-off: supprimer reliquat nft ip $HB_NFT_TABLE et mitm :$HB_MITM_PORT s'ils existent"
  echo "DRY hb-off: NE PAS toucher la table ip sandrpc ni host-main"
  exit 0
fi

need_root
slog "hb-off: début — arrêt capture passive, host-main NON redémarré"
hb_stop_pids
hb_shred "$HB_PCAP"
hb_drop_legacy_nft
hb_kill_legacy_mitm
if command -v nft >/dev/null 2>&1 && nft list table ip sandrpc >/dev/null 2>&1; then
  slog "hb-off: table ip sandrpc (agent<->host) intacte, non touchée."
fi
slog "hb-off: FIN. keylog conservé ($HB_KEYLOG_FILE). Teardown complet: bash hb-revert.sh"
exit 0
