#!/usr/bin/env bash
# test-agent-host.sh — PREUVE de la variante agent<->host "sans restart", sur
# cibles SCRATCH uniquement. Ne touche JAMAIS host-main, ni le vrai 1340/1342,
# ni nft (aucune table réelle posée), ni l'egress. Le relai est testé via les
# nouveaux arguments HOST:PORT (--listen 127.0.0.1:19342 --upstream 127.0.0.1:19340).
#
# Prouve :
#   (a) req ET resp body capturés en JSONL
#   (b) leg="agent-host"
#   (c) grep anti-secret VIDE (Bearer/ghp_/PEM BLOC COMPLET/email -> rédigés)
#   (d) revert-agent-host.sh --dry-run OK (innocuité)
set -uo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LI="$(cd "$HERE/.." && pwd)"
RUN="$HERE/_run"; rm -rf "$RUN"; mkdir -p "$RUN"
export RPC_TRAFFIC_LOG_DIR="$RUN"           # logs de test ISOLÉS (hors shared-eg)
PIDS=()
cleanup(){ for p in "${PIDS[@]:-}"; do kill "$p" 2>/dev/null || true; done; }
trap cleanup EXIT

FAKE_HOST_PORT=19340
FAKE_RELAY_PORT=19342

echo "=============== relai no-restart devant faux host (scratch) ==============="
python3 "$HERE/scratch_servers.py" plain "$FAKE_HOST_PORT" & PIDS+=($!); sleep 1
# NOUVELLE syntaxe HOST:PORT (ce que arm-agent-host.sh utilise)
RPC_TRAFFIC_LOG_DIR="$RUN" python3 "$LI/relay_agent_host.py" \
    --listen "127.0.0.1:${FAKE_RELAY_PORT}" --upstream "127.0.0.1:${FAKE_HOST_PORT}" \
    >"$RUN/relay.out" 2>&1 & PIDS+=($!); sleep 1

# corps de requête avec secrets factices (header + body), PEM MULTI-LIGNE pour
# exercer le PKEY_RE en BLOC COMPLET (BEGIN...END sur plusieurs lignes).
BODY="$RUN/req_body.txt"
{
  printf 'auth-in-body: Bearer ghp_FAKEAAAAAAAAAAAAAAAAAAAAAAAA\n'
  printf 'gh-token: ghp_FAKEBBBBBBBBBBBBBBBBBBBBBBBB\n'
  printf 'contact: victim@example.com\n'
  printf -- '-----BEGIN PRIVATE KEY-----\n'
  printf 'MIIFAKEKEYBODYLINE1aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa\n'
  printf 'MIIFAKEKEYBODYLINE2bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb\n'
  printf -- '-----END PRIVATE KEY-----\n'
} > "$BODY"

curl -s -X POST "http://127.0.0.1:${FAKE_RELAY_PORT}/aiserver.v1.DashboardService/GetTeams" \
  -H "Authorization: Bearer ghp_FAKEHEADERCCCCCCCCCCCCCCCCCCCC" \
  -H "Content-Type: text/plain" \
  --data-binary @"$BODY" >/dev/null
sleep 1

echo "=============== (d) revert-agent-host.sh --dry-run (innocuité) ==============="
bash "$LI/revert-agent-host.sh" --dry-run 2>&1 | sed 's/^/   /' | head -20
RC_DRY=${PIPESTATUS[0]}

LOG="$(ls -1 "$RUN"/traffic-*.jsonl 2>/dev/null | head -1)"
echo
echo "=============== RÉSULTAT ==============="
if [ -z "$LOG" ] || [ ! -s "$LOG" ]; then echo "ÉCHEC: aucun log produit"; exit 1; fi
echo "log: $LOG  ($(wc -l < "$LOG") lignes)"

echo "--- (a)+(b) leg / bodies capturés ---"
python3 - "$LOG" <<'PY'
import json,sys
ok_leg=ok_req=ok_resp=False
for l in open(sys.argv[1]):
    r=json.loads(l)
    leg=r.get('leg'); rb=r.get('req_body'); sb=r.get('resp_body')
    print(f"  leg={leg} http={r.get('http_method')} method={r.get('method')} status={r.get('status')} req_body?={rb is not None} resp_body?={sb is not None}")
    if leg=='agent-host': ok_leg=True
    if rb is not None: ok_req=True
    if sb is not None: ok_resp=True
assert ok_leg, "leg agent-host absent"
assert ok_req and ok_resp, "req/resp body manquant"
print("  (a) OK req+resp capturés ; (b) OK leg=agent-host")
PY

echo "--- (c) grep anti-secret (doit être VIDE) ---"
# motifs de secrets EN CLAIR qui NE doivent PAS apparaître :
#   ghp_FAKE      : token GH (header + body)
#   @example.com  : email (body + réponse scratch)
#   BEGIN PRIVATE KEY / FAKEKEYBODYLINE : PEM header ET corps base64 (bloc complet)
#   eyJhbGciOiJIUzI1NiJ9.eyJ : JWT de la réponse scratch
if grep -nE 'ghp_FAKE|@example\.com|BEGIN PRIVATE KEY|FAKEKEYBODYLINE|eyJhbGciOiJIUzI1NiJ9\.eyJ' "$LOG"; then
  echo "ÉCHEC: secret/PII en clair détecté !!"; exit 2
else
  echo "  (c) OK: 0 secret/PII en clair"
fi
echo "--- marqueurs de rédaction présents (dont PRIVATE-KEY = bloc complet) ---"
grep -oE 'REDACTED[A-Z-]*' "$LOG" | sort | uniq -c

echo "--- (d) revert dry-run ---"
if [ "${RC_DRY:-1}" = "0" ]; then echo "  (d) OK: revert-agent-host.sh --dry-run exit 0"; else echo "ÉCHEC revert dry-run"; exit 3; fi

echo
echo "SUCCÈS: relai no-restart capture req+resp (corps complets), leg=agent-host,"
echo "rédaction sans fuite (clé privée en BLOC COMPLET), revert dry-run sûr."
echo "AUCUNE table nft réelle posée, host-main jamais touché."
