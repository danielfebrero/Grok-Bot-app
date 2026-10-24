#!/usr/bin/env bash
# hb-setup-keylog.sh — rend host-main « keylog-ready » en le relançant UNE fois.
#
# C'est le SEUL moment où host-main est redémarré pour la patte host↔backend.
# On n'injecte PAS de CA. On n'ajoute pas NODE_EXTRA_CA_CERTS (on le retire s'il
# traînait d'un ancien mitm). Le TLS reste celui du vrai backend : pas de
# pinning, pas de certificat maison.
#
# Vérifié sur /exec-daemon/node v22.23.4 (OpenSSL 3.5) : l'env SSLKEYLOGFILE
# seul ne crée aucun fichier. Node n'écrit le keylog NSS que si argv contient
# --tls-keylog=FICHIER. On passe ce flag ET SSLKEYLOGFILE= le même chemin
# (no-op sur ce build, utile si un Node futur honore l'env). Rien d'autre.
#
# L'environnement d'origine est lu dans /proc/PID/environ, en mémoire, et n'est
# jamais écrit sur disque ni journalisé.
#
# La capture (tcpdump + décodeur) est démarrée AVANT le restart, pour que les
# ClientHello du nouveau host-main soient dans le pcap (sans eux tshark ne peut
# pas associer les secrets). hb-off l'arrête ensuite sans nouveau restart.
# ============================================================================
#  ⚠️  OPÉRATEUR UNIQUEMENT. Sans jeton -> DRY-RUN.
#      LI_ARM_CONFIRM=yes-i-understand-brick-risk sudo -E bash hb-setup-keylog.sh
# ============================================================================
set -uo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# shellcheck source=lib.sh
. "$HERE/lib.sh"

GRACE="${LI_ARM_GRACE:-90}"
CONFIRM="${LI_ARM_CONFIRM:-}"
DRY=1; [ "$CONFIRM" = "yes-i-understand-brick-risk" ] && DRY=0
REVERT="$HERE/hb-revert.sh"

plan() {
  cat <<PLAN
=== PLAN hb-setup-keylog (DRY=$DRY) — un restart host-main, SANS CA ===
 host-main courant : $(hostmain_pid || echo '??')  (uid $HB_UID, port $HOST_PORT_ORIG inchangé)
 keylog            : $HB_KEYLOG_FILE  (0600, uid $HB_UID, hors logs partagés)
 injection         : argv --tls-keylog=<keylog>  + env SSLKEYLOGFILE=<même chemin>
 retiré de l'env   : NODE_EXTRA_CA_CERTS (ancien mitm) et un éventuel SSLKEYLOGFILE précédent
 1. dead-man watchdog (grace ${GRACE}s) AVANT le restart, flags $HB_ARM_FLAG / $HB_DISARM_FLAG
 2. préparer le keylog vide ; supprimer un reliquat nft ip $HB_NFT_TABLE et mitm :$HB_MITM_PORT
    (la table ip sandrpc agent<->host n'est PAS touchée)
 3. démarrer la capture passive (tcpdump + tshark) AVANT le restart
    filtre: $HB_CAPTURE_FILTER
 4. lire /proc/PID/environ en mémoire ; spawn le même host-main (port $HOST_PORT_ORIG)
    avec UNIQUEMENT --tls-keylog + SSLKEYLOGFILE ; SIGTERM l'ancien
 5. santé sur :$HOST_PORT_ORIG + cmdline contient --tls-keylog ; backend joignable
    -> touch $HB_KEYLOG_READY_FLAG
 Confirmer (annuler le dead-man, garder keylog+capture) : touch $HB_DISARM_FLAG
 Arrêter la capture SANS restart : bash $HERE/hb-off.sh
 Tout retirer (respawn host SANS keylog) : bash $REVERT
PLAN
}

abort_before_restart() {
  slog "hb-setup-keylog: $* — host-main NON modifié"
  bash "$HERE/hb-off.sh" || true
  touch "$HB_DISARM_FLAG" 2>/dev/null || true
  rm -f "$HB_ARM_FLAG"
  if [ -f "$WATCHDOG_PIDFILE" ]; then
    wp="$(cat "$WATCHDOG_PIDFILE" 2>/dev/null || true)"
    [ -n "${wp:-}" ] && kill "$wp" 2>/dev/null || true
    rm -f "$WATCHDOG_PIDFILE"
  fi
  exit 6
}

plan
if [ "$DRY" = 1 ]; then
  echo
  echo ">>> DRY-RUN : rien modifié, host-main intact, pas de keylog, pas de capture."
  echo ">>> Pour exécuter réellement (OPÉRATEUR) :"
  echo ">>>   LI_ARM_CONFIRM=yes-i-understand-brick-risk sudo -E bash hb-setup-keylog.sh"
  exit 0
fi

need_root
OLD_PID="$(hostmain_pid)"
[ -z "$OLD_PID" ] && { slog "hb-setup-keylog: pas de host-main courant, abandon"; exit 2; }
[ -x "$NODE_BIN" ] || { slog "hb-setup-keylog: node introuvable ($NODE_BIN)"; exit 2; }

mkdir -p "$STATE_DIR"
rm -f "$HB_DISARM_FLAG"
touch "$HB_ARM_FLAG"
slog "hb-setup-keylog: HB_ARM posé (host-main $OLD_PID, port $HOST_PORT_ORIG) — env NON journalisé"
bash "$HERE/watchdog.sh" --grace "$GRACE" --revert "$REVERT" \
     --arm-flag "$HB_ARM_FLAG" --disarm-flag "$HB_DISARM_FLAG"
slog "hb-setup-keylog: watchdog dédié lancé (grace ${GRACE}s)"

hb_prepare_keylog_file || { abort_before_restart "préparation keylog impossible"; }
hb_drop_legacy_nft
hb_kill_legacy_mitm

# Capture AVANT le restart : les ClientHello du nouveau process doivent être
# dans le pcap. Si tcpdump ne part pas, on n'a pas encore touché host-main.
if hb_capture_up; then
  slog "hb-setup-keylog: capture déjà active"
else
  slog "hb-setup-keylog: démarrage capture passive (avant restart)"
  setsid bash "$HERE/start-capture-host-backend.sh" \
      </dev/null >>"$HB_CAPTURE_OUT" 2>&1 &
  echo $! > "$HB_DECODE_PIDFILE"
  ok=0
  for _ in $(seq 1 40); do
    [ -f "$HB_CAPTURE_READY" ] && hb_capture_up && { ok=1; break; }
    sleep 0.25
  done
  [ "$ok" = 1 ] || abort_before_restart "capture non prête (voir $HB_CAPTURE_OUT)"
fi
slog "hb-setup-keylog: capture prête, restart host-main"

# Environ d'origine : mémoire seulement. On ne l'écrit nulle part.
mapfile -d '' -t OENV < "/proc/$OLD_PID/environ"
NEWENV=()
for kv in "${OENV[@]}"; do
  case "$kv" in
    NODE_EXTRA_CA_CERTS=*) continue;;
    SSLKEYLOGFILE=*) continue;;
    *) NEWENV+=("$kv");;
  esac
done
NEWENV+=("SSLKEYLOGFILE=$HB_KEYLOG_FILE")

slog "hb-setup-keylog: SIGTERM ancien host-main $OLD_PID puis spawn avec --tls-keylog (port $HOST_PORT_ORIG)"
kill -TERM "$OLD_PID" 2>/dev/null || true
: > "$STATE_DIR/hb-hostmain.out"
chmod 0600 "$STATE_DIR/hb-hostmain.out" 2>/dev/null || true
setsid env -i "${NEWENV[@]}" \
    "$NODE_BIN" "$NODE_FLAG" "--tls-keylog=$HB_KEYLOG_FILE" "$HOST_ENTRY" \
    </dev/null >>"$STATE_DIR/hb-hostmain.out" 2>&1 &
echo $! > "$HB_HOSTMAIN_PIDFILE"
unset OENV NEWENV

for _ in $(seq 1 30); do host_health "$HOST_PORT_ORIG" && break; sleep 1; done
if ! host_health "$HOST_PORT_ORIG"; then
  slog "hb-setup-keylog: host-main NE répond PAS sur $HOST_PORT_ORIG -> REVERT"
  bash "$REVERT"; exit 3
fi
if ! hostmain_has_tls_keylog; then
  slog "hb-setup-keylog: le host-main sain n'a PAS --tls-keylog (adoption superviseur ?) -> REVERT"
  bash "$REVERT"; exit 3
fi
slog "hb-setup-keylog: host-main sain sur $HOST_PORT_ORIG avec --tls-keylog"

if backend_reachable; then
  touch "$HB_KEYLOG_READY_FLAG"
  chmod 0600 "$HB_KEYLOG_FILE" 2>/dev/null || true
  slog "hb-setup-keylog: backend joignable -> keylog-ready. touch $HB_DISARM_FLAG pour annuler le dead-man."
  slog "hb-setup-keylog: capture DÉJÀ active. L'arrêter sans restart : bash $HERE/hb-off.sh"
else
  slog "hb-setup-keylog: ATTENTION backend injoignable post-restart -> le watchdog va revert"
fi
