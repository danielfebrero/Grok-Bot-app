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
