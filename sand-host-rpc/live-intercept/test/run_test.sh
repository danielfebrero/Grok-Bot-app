#!/usr/bin/env bash
# run_test.sh — PREUVE de capture (req+resp) + rédaction + revert, sur cibles
# SCRATCH uniquement. Ne touche JAMAIS host-main, ni nft, ni l'egress réel.
#   A) patte agent-host : relay_agent_host.py devant un faux host plaintext
#   B) patte host-backend : test-host-backend.sh (keylog NSS + tshark, pas de mitm)
#   C) revert.sh --dry-run (prouve l'innocuité)
# Puis : grep anti-secret sur le JSONL agent-host (doit être VIDE).
set -uo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LI="$(cd "$HERE/.." && pwd)"
RUN="$HERE/_run"; rm -rf "$RUN"; mkdir -p "$RUN"
export RPC_TRAFFIC_LOG_DIR="$RUN"           # logs de test isolés (hors shared-eg)
PIDS=()
cleanup(){ for p in "${PIDS[@]}"; do kill "$p" 2>/dev/null || true; done; }
trap cleanup EXIT

echo "=============== A) patte agent-host (relai plaintext) ==============="
python3 "$HERE/scratch_servers.py" plain 19341 & PIDS+=($!); sleep 1
RPC_TRAFFIC_LOG_DIR="$RUN" python3 "$LI/relay_agent_host.py" --listen 19340 --upstream-port 19341 >"$RUN/relay.out" 2>&1 & PIDS+=($!); sleep 1
curl -s -X POST http://127.0.0.1:19340/aiserver.v1.DashboardService/GetTeams \
  -H "Authorization: Bearer eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJyZXF1ZXN0LXNlY3JldCJ9.cmVxc2lnbmF0dXJlc2VjcmV0eHg" \
  -H "Content-Type: application/json" \
  -d '{"user":"agent@example.com","note":"corps requete avec secret"}' >/dev/null
sleep 1

echo "=============== B) patte host-backend (keylog NSS + tshark) ==============="
bash "$HERE/test-host-backend.sh"

echo "=============== B2) websocket (relai + CDP loopback) ==============="
python3 "$HERE/test_ws_frames.py"
echo "=============== B3) rotation pcap (cap, seen, faux tcpdump) ==============="
python3 "$HERE/test_pcap_rotate.py"
bash "$HERE/test-cdp-ws.sh"

echo "=============== C) revert.sh --dry-run (innocuité) ==============="
bash "$LI/revert.sh" --dry-run 2>&1 | sed 's/^/   /' | head -20

LOG="$(ls -1 "$RUN"/traffic-*.jsonl 2>/dev/null | head -1)"
echo
echo "=============== RÉSULTAT ==============="
if [ -z "$LOG" ] || [ ! -s "$LOG" ]; then echo "ÉCHEC: aucun log produit"; exit 1; fi
echo "log: $LOG  ($(wc -l < "$LOG") lignes)"
echo "--- legs capturés ---"; python3 - "$LOG" <<'PY'
import json,sys
for l in open(sys.argv[1]):
    r=json.loads(l)
    print(f"  leg={r.get('leg')} http={r.get('http_method')} method={r.get('method')} status={r.get('status')} req_body?={'req_body' in r and r['req_body'] is not None} resp_body?={'resp_body' in r and r['resp_body'] is not None}")
PY
echo "--- grep anti-secret (doit être VIDE) ---"
# motifs de secrets en CLAIR qui NE doivent PAS apparaître
if grep -nE 'eyJhbGciOiJIUzI1NiJ9\.eyJ|ghp_abcdef|supersecretcookie|BEGIN PRIVATE KEY|@example\.com|@corp\.com' "$LOG"; then
  echo "ÉCHEC: secret/PII en clair détecté !!"; exit 2
else
  echo "  OK: 0 secret/PII en clair"
fi
echo "--- marqueurs de rédaction présents ---"
grep -oE 'REDACTED[A-Z-]*' "$LOG" | sort | uniq -c
echo "--- classifieur de patte (addon_log._direction) ---"
python3 - "$LI/../rpc-traffic-monitor/mitmproxy" <<'PYX'
import sys
sys.path.insert(0, sys.argv[1])
import addon_log
assert addon_log._direction("api2.cursor.sh",443)=="host->backend"
assert addon_log._direction("repo42.cursor.sh",443)=="host->backend"
assert addon_log._direction("127.0.0.1",1340)=="agent->host"
print("  OK: api2/repo42->host-backend, 127.0.0.1:1340->agent-host")
PYX
echo
echo "SUCCÈS: capture req+resp (corps) via relai plaintext ET keylog host-backend ; rédaction OK ; revert dry-run sûr."
