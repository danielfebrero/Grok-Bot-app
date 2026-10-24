#!/usr/bin/env bash
# lib.sh — configuration + fonctions communes aux scripts live-intercept.
# Aucun secret ici. Tout ce qui est runtime (flags ARM/DISARM, pids, état) vit
# HORS du dépôt, dans STATE_DIR (gitignoré de toute façon mais placé dans /workspace).
set -uo pipefail

# ---- chemins ----------------------------------------------------------------
LI_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
RPC_ROOT="$(cd "$LI_DIR/.." && pwd)"                       # .../sand-host-rpc
MONITOR_DIR="$RPC_ROOT/rpc-traffic-monitor"               # proxy.py + addon_log.py
ADDON="$MONITOR_DIR/mitmproxy/addon_log.py"
PROXY_PY="$MONITOR_DIR/proxy.py"
RELAY_PY="$LI_DIR/relay_agent_host.py"

# état runtime (JAMAIS committé) :
STATE_DIR="${LI_STATE_DIR:-/workspace/.sand-live-intercept}"
ARM_FLAG="$STATE_DIR/ARM"
DISARM_FLAG="$STATE_DIR/DISARM"
STATE_LOG="$STATE_DIR/state.log"
WATCHDOG_LOG="$STATE_DIR/watchdog.log"
WATCHDOG_PIDFILE="$STATE_DIR/watchdog.pid"
RELAY_PIDFILE="$STATE_DIR/relay.pid"
MITM_PIDFILE="$STATE_DIR/mitm.pid"
ORIG_PORT_FILE="$STATE_DIR/orig_host_port"     # port d'origine de host-main (def 1340)
HOSTMAIN_SPAWN_PIDFILE="$STATE_DIR/hostmain_arm.pid"  # pid du host-main lancé par arm

# logs de trafic : dossier Syncthing partagé EG (écrit par la box uniquement)
TRAFFIC_LOG_DIR="${RPC_TRAFFIC_LOG_DIR:-/workspace/shared-eg/box-vers-vps/sand-host-rpc-traffic}"

# ---- topologie host ---------------------------------------------------------
HOST_ENTRY="/opt/sand/sand-host/host-main.cjs"
HOST_DIR="/opt/sand/sand-host"
NODE_BIN="/exec-daemon/node"
NODE_FLAG="--disable-warning=ExperimentalWarning"
HOST_UID=1000                      # host-main tourne en uid 1000 (box) — cf. README
GATEWAY_JSON="/home/box/sand-data/gateway.json"

# ---- ports / interception ---------------------------------------------------
HOST_PORT_ORIG=1340               # port agent<->host d'origine (plaintext HTTP/1.1)
HOST_PORT_ARM=1341                # port vers lequel arm déplace host-main
RELAY_LISTEN=1340                 # le relai loggeur prend la place sur 1340
MITM_TRANSPARENT_PORT=8900        # mitmproxy transparent (patte host<->backend)

# chaîne/table netfilter dédiée (backend ACTIF = nft, cf. README). Jamais hookée
# tant que arm ne l'a pas remplie. revert = delete table.
NFT_TABLE="sandrpc"               # table ip sandrpc

# CA mitmproxy (certificat PUBLIC ; la clé privée n'est jamais touchée/committée)
MITM_CA="${MITM_CA:-/home/box/.mitmproxy/mitmproxy-ca-cert.pem}"

# ---- venv mitmproxy ---------------------------------------------------------
MITM_VENV="/workspace/.venvs/mitmproxy"
MITMDUMP="$MITM_VENV/bin/mitmdump"

# ---- utilitaires ------------------------------------------------------------
ts() { date '+%Y-%m-%dT%H:%M:%S%z'; }

slog() {  # state log horodaté (stdout + fichier)
  local msg="[$(ts)] $*"
  echo "$msg"
  mkdir -p "$STATE_DIR" 2>/dev/null
  echo "$msg" >> "$STATE_LOG" 2>/dev/null || true
}

need_root() {
  if [ "$(id -u)" != "0" ]; then
    if sudo -n true 2>/dev/null; then SUDO="sudo -n"; else
      echo "ERREUR: root requis (sudo -n indisponible)"; exit 90; fi
  else SUDO=""; fi
}

# host-main vivant ? renvoie le pid sur stdout, code 0 si trouvé
hostmain_pid() {
  pgrep -f 'host-main\.cjs' 2>/dev/null | head -1
}

# le backend est-il joignable ? (probe TLS direct vers api2, chemin normal, SANS
# passer par le mitm). 0 = joignable.
backend_reachable() {
  timeout 6 bash -c 'exec 3<>/dev/tcp/api2.cursor.sh/443' 2>/dev/null && { exec 3>&- 3<&-; return 0; }
  return 1
}

# host-main répond-il sur un port donné ? (0 = oui)
host_health() {
  local port="${1:-1340}"
  timeout 4 curl -s "http://127.0.0.1:${port}/health" 2>/dev/null | grep -q '"ok":true'
}

# ============================================================================
#  Variante "sans restart host-main" (patte agent<->host uniquement)
#  arm-agent-host.sh / revert-agent-host.sh. Ne touche NI host-main, NI la CA,
#  NI le 443. Un simple relai loggeur devant le host-main RÉEL (intact) + une
#  redirection nft en prerouting limitée au SEUL client agent.
# ----------------------------------------------------------------------------
AH_AGENT_SADDR="${AH_AGENT_SADDR:-172.30.0.1}"   # client agent (netns distinct)
AH_HOST_DADDR="${AH_HOST_DADDR:-172.30.0.2}"     # host-main RÉEL (enp0s4), intact
AH_HOST_PORT="${AH_HOST_PORT:-1340}"             # port host-main réel (plaintext HTTP/1.1)
AH_RELAY_PORT="${AH_RELAY_PORT:-1342}"           # relai loggeur no-restart
# NB bind : la cible `redirect` réécrit la daddr vers l'adresse de l'IFACE
# d'entrée (172.30.0.2), PAS vers 127.0.0.1 (cf. sémantique REDIRECT/nat). Le
# relai doit donc être joignable sur cette adresse -> bind 0.0.0.0 par défaut.
# (La valeur "logique" visée reste 127.0.0.1:1342 ; 0.0.0.0 la couvre + l'iface.)
AH_RELAY_BIND="${AH_RELAY_BIND:-0.0.0.0}"
AH_IIF="${AH_IIF:-enp0s4}"                        # iface côté agent (scoping défensif)
AH_RELAY_PIDFILE="$STATE_DIR/relay-agent-host.pid"
AH_NFT_TABLE="${AH_NFT_TABLE:-sandrpc}"           # table ip dédiée (chain prerouting)

# le relai no-restart écoute-t-il bien ? (0 = oui)
ah_relay_up() {
  timeout 3 curl -s -o /dev/null "http://127.0.0.1:${AH_RELAY_PORT}/" 2>/dev/null
  # 000 (pas de réponse HTTP) => down ; tout code HTTP => up
  local c; c=$(timeout 3 curl -s -o /dev/null -w '%{http_code}' "http://127.0.0.1:${AH_RELAY_PORT}/" 2>/dev/null || echo 000)
  [ "$c" != "000" ]
}

# ============================================================================
#  Patte host<->backend — déchiffrement PASSIF par keylog TLS (pas de MITM).
#
#  Node 22 de cette box (/exec-daemon/node) N'honore PAS l'env SSLKEYLOGFILE.
#  Il écrit le format NSS (celui que Wireshark appelle SSLKEYLOGFILE) uniquement
#  si on le lance avec --tls-keylog=FICHIER. hb-setup-keylog.sh fait ce restart
#  UNE fois (jeton LI_ARM_CONFIRM) et pose AUSSI SSLKEYLOGFILE= même chemin pour
#  un build qui, lui, lirait l'env. Aucune CA, aucun NODE_EXTRA_CA_CERTS.
#
#  Ensuite hb-on / hb-off ne font que démarrer/arrêter tcpdump + decode_keylog.py
#  (tshark + tls.keylog_file). Aucune règle nft, aucun redirect :443.
#
#  La table agent<->host `ip sandrpc` n'est JAMAIS modifiée ici. Si une vieille
#  table mitm `ip sandrpc_hb` traîne, hb-off/hb-revert/hb-setup la suppriment.
# ----------------------------------------------------------------------------
HB_NFT_TABLE="${HB_NFT_TABLE:-sandrpc_hb}"        # reliquat mitm — cleanup seulement
HB_MITM_PORT="${HB_MITM_PORT:-8900}"              # reliquat mitm — cleanup seulement
HB_UID="${HB_UID:-1000}"                          # host-main (= uid box)
HB_TUNNEL_ADDR="${HB_TUNNEL_ADDR:-198.18.0.1}"    # egress tunnel — exclu du BPF

# SNI / :authority / Host à JOURNALISER. Le pcap peut contenir d'autre TLS (non
# déchiffrable sans ces clés, ou déchiffrable si host-main y a parlé) : le
# décodeur ne garde que ce qui matche. Aligné sur BACKEND_HOSTS de addon_log.py.
# (re.search, ancré en fin : settings.cursor.sh.evil.com est rejeté.)
HB_ALLOW_HOSTS="${HB_ALLOW_HOSTS:-(^|\\.)(api|api2|api3|api4|repo42)\\.cursor\\.sh(:443)?$|(^|\\.)(www\\.)?cursor\\.com(:443)?$|(^|\\.)cursor\\.sh(:443)?$}"

HB_KEYLOG_DIR="${HB_KEYLOG_DIR:-$STATE_DIR/keylog}"
HB_KEYLOG_FILE="${HB_KEYLOG_FILE:-$HB_KEYLOG_DIR/sslkeys.log}"
HB_PCAP="${HB_PCAP:-$STATE_DIR/hb-capture.pcap}"
HB_PCAP_MAX_BYTES="${HB_PCAP_MAX_BYTES:-536870912}"   # 512 Mio ; au-delà le décodeur coupe tcpdump
HB_IFACE="${HB_IFACE:-any}"
HB_CAPTURE_FILTER="${HB_CAPTURE_FILTER:-tcp port 443 and not host ${HB_TUNNEL_ADDR}}"
HB_TLS_PORTS="${HB_TLS_PORTS:-}"                    # ports != 443 à forcer en TLS (tests)
HB_DECODE_INTERVAL="${HB_DECODE_INTERVAL:-1}"
HB_TSHARK="${HB_TSHARK:-$(command -v tshark 2>/dev/null || true)}"
HB_TCPDUMP="${HB_TCPDUMP:-$(command -v tcpdump 2>/dev/null || true)}"

HB_TCPDUMP_PIDFILE="$STATE_DIR/hb-tcpdump.pid"
HB_DECODE_PIDFILE="$STATE_DIR/hb-decode.pid"
HB_CAPTURE_OUT="$STATE_DIR/hb-capture.out"
HB_CAPTURE_READY="$STATE_DIR/hb-capture.ready"
HB_SEEN_FILE="$STATE_DIR/hb-decode.seen"
HB_KEYLOG_READY_FLAG="$STATE_DIR/hb_keylog_ready"   # posé par hb-setup-keylog
HB_HOSTMAIN_PIDFILE="$STATE_DIR/hb-hostmain.pid"
HB_ARM_FLAG="$STATE_DIR/HB_ARM"                     # dead-man dédié (pas l'état agent-host)
HB_DISARM_FLAG="$STATE_DIR/HB_DISARM"
# reliquats mitm (cleanup) :
HB_MITM_PIDFILE="$STATE_DIR/hb-mitm.pid"
HB_CA_READY_FLAG="$STATE_DIR/hb_ca_ready"

# gid primaire de l'uid host-main (pour chown du keylog, que node puisse écrire)
hb_host_gid() {
  id -g "$HB_UID" 2>/dev/null || echo "$HB_UID"
}

# Le process host-main courant porte-t-il --tls-keylog vers NOTRE fichier ?
# (cmdline seulement ; on ne lit pas /proc/PID/environ et on ne l'affiche pas.)
hostmain_has_tls_keylog() {
  local p
  for p in $(pgrep -f 'host-main\.cjs' 2>/dev/null || true); do
    if [ -r "/proc/$p/cmdline" ] && grep -a -F -q -- "--tls-keylog=${HB_KEYLOG_FILE}" "/proc/$p/cmdline" 2>/dev/null; then
      return 0
    fi
  done
  return 1
}

# 0 si le pid existe encore et n'est pas un zombie. kill -0 reste vrai sur un
# zombie : hb-off attendrait sinon ~5 s que python (qui n'attend pas tcpdump)
# récolte un fils déjà mort.
hb_pid_live() {
  local p="$1" line state
  [ -n "${p:-}" ] || return 1
  [ -r "/proc/$p/stat" ] || return 1
  line="$(cat "/proc/$p/stat" 2>/dev/null || true)"
  state="${line##*) }"
  state="${state%% *}"
  [ -n "$state" ] && [ "$state" != "Z" ]
}

hb_capture_up() {
  local p=""
  [ -f "$HB_TCPDUMP_PIDFILE" ] || return 1
  p="$(cat "$HB_TCPDUMP_PIDFILE" 2>/dev/null || true)"
  hb_pid_live "$p"
}

# Efface un fichier de secrets (pcap ou keylog) sans jamais l'imprimer.
hb_shred() {
  local f="$1"
  [ -n "$f" ] && [ -e "$f" ] || return 0
  if command -v shred >/dev/null 2>&1; then
    shred -u -n 1 -- "$f" 2>/dev/null || rm -f -- "$f"
  else
    rm -f -- "$f"
  fi
}

# Crée le keylog vide, 0600, propriétaire uid host-main (node doit pouvoir append).
# Le contenu n'est jamais affiché. Hors dépôt, hors dossier de logs partagés.
hb_prepare_keylog_file() {
  case "$HB_KEYLOG_FILE" in
    "$TRAFFIC_LOG_DIR"|"$TRAFFIC_LOG_DIR"/*)
      echo "ERREUR: le keylog ne doit pas vivre dans les logs de trafic ($TRAFFIC_LOG_DIR)"; return 4;;
    "$LI_DIR"|"$LI_DIR"/*)
      echo "ERREUR: le keylog ne doit pas vivre dans le dépôt ($LI_DIR)"; return 4;;
  esac
  local gid; gid="$(hb_host_gid)"
  install -d -m 0750 "$STATE_DIR"
  install -d -m 0700 -o "$HB_UID" -g "$gid" "$HB_KEYLOG_DIR"
  if [ ! -e "$HB_KEYLOG_FILE" ]; then
    install -m 0600 -o "$HB_UID" -g "$gid" /dev/null "$HB_KEYLOG_FILE"
  else
    chown "$HB_UID:$gid" "$HB_KEYLOG_FILE" 2>/dev/null || true
    chmod 0600 "$HB_KEYLOG_FILE" 2>/dev/null || true
  fi
}

# Supprime UNIQUEMENT la table reliquat mitm ip sandrpc_hb. Jamais `ip sandrpc`.
hb_drop_legacy_nft() {
  command -v nft >/dev/null 2>&1 || return 0
  local sud="${SUDO:-}"
  if [ "$(id -u)" != "0" ]; then
    if [ -z "$sud" ]; then
      sudo -n true 2>/dev/null || return 0
      sud="sudo -n"
    fi
  fi
  if $sud nft list table ip "$HB_NFT_TABLE" >/dev/null 2>&1; then
    slog "hb: suppression reliquat nft table ip $HB_NFT_TABLE (ancien mitm)"
    $sud nft delete table ip "$HB_NFT_TABLE" || true
  fi
  if $sud nft list table ip sandrpc >/dev/null 2>&1; then
    slog "hb: table ip sandrpc (agent<->host) présente, NON touchée."
  fi
}

# Arrête tcpdump puis le décodeur (le décodeur fait une dernière passe sur SIGTERM).
# Ne supprime le pcap que si $1 = shred (hb-off). hb_stop_capture_keep laisse le pcap
# au décodeur ; l'appelant shred ensuite.
hb_stop_pids() {
  local pf p
  # tcpdump d'abord pour flush (-U) avant la dernière passe du décodeur.
  for pf in "$HB_TCPDUMP_PIDFILE" "$HB_DECODE_PIDFILE"; do
    [ -f "$pf" ] || continue
    p="$(cat "$pf" 2>/dev/null || true)"
    if hb_pid_live "${p:-}"; then
      slog "hb: SIGTERM pid $p ($pf)"
      kill -TERM "$p" 2>/dev/null || true
    fi
    if [ "$pf" = "$HB_TCPDUMP_PIDFILE" ]; then
      # laisser le pcap se fermer avant de signaler le décodeur
      local i
      for i in $(seq 1 25); do
        hb_pid_live "${p:-}" || break
        sleep 0.2
      done
    fi
  done
  # attendre le décodeur (passe finale)
  if [ -f "$HB_DECODE_PIDFILE" ]; then
    p="$(cat "$HB_DECODE_PIDFILE" 2>/dev/null || true)"
    local i
    for i in $(seq 1 40); do
      hb_pid_live "${p:-}" || break
      sleep 0.2
    done
    if hb_pid_live "${p:-}"; then
      slog "hb: décodeur $p encore vivant, SIGKILL"
      kill -KILL "$p" 2>/dev/null || true
    fi
  fi
  rm -f "$HB_TCPDUMP_PIDFILE" "$HB_DECODE_PIDFILE" "$HB_CAPTURE_READY"
}

hb_kill_legacy_mitm() {
  if [ -f "$HB_MITM_PIDFILE" ]; then
    local p
    p="$(cat "$HB_MITM_PIDFILE" 2>/dev/null || true)"
    if [ -n "$p" ] && kill -0 "$p" 2>/dev/null; then
      slog "hb: arrêt mitm reliquat pid $p"
      kill "$p" 2>/dev/null || true
    fi
    rm -f "$HB_MITM_PIDFILE"
  fi
  local p
  for p in $(pgrep -f "mitmdump .*--listen-port ${HB_MITM_PORT}" 2>/dev/null || true); do
    slog "hb: arrêt mitmdump reliquat pid $p"
    kill "$p" 2>/dev/null || true
  done
  rm -f "$HB_CA_READY_FLAG"
}
