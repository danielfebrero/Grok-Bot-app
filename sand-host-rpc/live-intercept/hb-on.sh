#!/usr/bin/env bash
# hb-on.sh — ACTIVE la capture host↔backend. Sans restart host-main.
#
# Démarre tcpdump + decode_keylog.py si besoin. Ne pose aucune règle nft, ne
# touche pas la table ip sandrpc (agent↔host), ne relance pas host-main.
#
# Pré-requis : hb-setup-keylog.sh a réussi (marqueur + --tls-keylog dans la
# cmdline du host-main vivant). Sinon les clés ne sont pas écrites et le pcap
# resterait opaque. HB_FORCE=1 saute ce contrôle (déconseillé).
#
# Les connexions TLS ouvertes AVANT ce démarrage ne sont pas déchiffrables :
# tshark a besoin du ClientHello dans le pcap. Seuls les nouveaux handshakes
# (après hb-on, ou ceux pris par la capture démarrée dans hb-setup) sortent
# dans le JSONL. Même limite après chaque rotation automatique du pcap
# (cap HB_PCAP_MAX_BYTES) : pas besoin de relancer hb-on.
#
# Dry-run par défaut.
#   HB_ON_CONFIRM=yes sudo -E bash hb-on.sh
#   (ou LI_ARM_CONFIRM=yes-i-understand-brick-risk)
# ============================================================================
set -uo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# shellcheck source=lib.sh
. "$HERE/lib.sh"
CONFIRM="${HB_ON_CONFIRM:-}"
[ "${LI_ARM_CONFIRM:-}" = "yes-i-understand-brick-risk" ] && CONFIRM=yes
DRY=1; [ "$CONFIRM" = "yes" ] && DRY=0

plan() {
  cat <<PLAN
=== PLAN hb-on (DRY=$DRY) — SANS restart host-main, SANS nft ===
 1. exiger le marqueur $HB_KEYLOG_READY_FLAG et --tls-keylog dans la cmdline
    (sauf HB_FORCE=1)
 2. si pas déjà active : tcpdump -i $HB_IFACE -B ${HB_TCPDUMP_BUFFER_KB}  filtre « $HB_CAPTURE_FILTER »
    + decode_keylog.py (tshark, keylog=$HB_KEYLOG_FILE)
    journal : $TRAFFIC_LOG_DIR  leg=host-backend  allow=$HB_ALLOW_HOSTS
    cap pcap : $HB_PCAP_MAX_BYTES octets, puis rotation auto (sessions TLS déjà
    ouvertes : opaques jusqu'au ClientHello suivant, sans relancer hb-on)
 OFF sans restart : bash $HERE/hb-off.sh
 Revert (retire le keylog, respawn host propre) : bash $HERE/hb-revert.sh
PLAN
}
plan
if [ "$DRY" = 1 ]; then
  echo
  echo ">>> DRY-RUN : capture non démarrée, host-main intact, aucune règle nft."
  echo ">>> Activer : HB_ON_CONFIRM=yes sudo -E bash hb-on.sh"
  exit 0
fi

need_root
if [ "${HB_FORCE:-0}" != "1" ]; then
  if [ ! -f "$HB_KEYLOG_READY_FLAG" ]; then
    slog "hb-on: pas keylog-ready ($HB_KEYLOG_READY_FLAG absent). Lance hb-setup-keylog.sh d'abord (ou HB_FORCE=1)."
    exit 5
  fi
  if ! hostmain_has_tls_keylog; then
    slog "hb-on: le host-main vivant n'a pas --tls-keylog vers $HB_KEYLOG_FILE. Marqueur périmé : relancer hb-setup-keylog.sh."
    exit 5
  fi
fi
[ -f "$HB_KEYLOG_FILE" ] || { slog "hb-on: keylog absent ($HB_KEYLOG_FILE)"; exit 4; }
chmod 0600 "$HB_KEYLOG_FILE" 2>/dev/null || true

if hb_capture_up && [ -f "$HB_DECODE_PIDFILE" ] && kill -0 "$(cat "$HB_DECODE_PIDFILE" 2>/dev/null)" 2>/dev/null; then
  slog "hb-on: capture déjà active (tcpdump $(cat "$HB_TCPDUMP_PIDFILE" 2>/dev/null)). Rien à faire."
else
  # un décodeur orphelin ou un tcpdump mort : on nettoie les pidfiles périmés
  # sans shredder un pcap encore utile — start-capture repart d'un pcap neuf.
  rm -f "$HB_TCPDUMP_PIDFILE" "$HB_DECODE_PIDFILE" "$HB_CAPTURE_READY"
  slog "hb-on: démarrage capture passive (sans restart host-main)"
  setsid bash "$HERE/start-capture-host-backend.sh" \
      </dev/null >>"$HB_CAPTURE_OUT" 2>&1 &
  echo $! > "$HB_DECODE_PIDFILE"
  ok=0
  for _ in $(seq 1 40); do
    [ -f "$HB_CAPTURE_READY" ] && hb_capture_up && { ok=1; break; }
    sleep 0.25
  done
  if [ "$ok" != 1 ]; then
    slog "hb-on: capture non prête (voir $HB_CAPTURE_OUT). Aucun restart, aucune règle nft."
    exit 6
  fi
fi
slog "hb-on: ACTIVÉ (host<->backend, passif). host-main NON redémarré. OFF = bash hb-off.sh"
if backend_reachable; then
  slog "hb-on: backend joignable (le chemin TLS réel n'est pas détourné)"
else
  slog "hb-on: ATTENTION backend injoignable — la capture est passive, bash hb-off.sh si besoin"
fi
