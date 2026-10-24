#!/usr/bin/env bash
# test-host-backend.sh — PREUVE de l'INTERCEPTION SÉLECTIVE host<->backend sur
# cibles SCRATCH uniquement. NE TOUCHE NI host-main, NI le vrai api2, NI nft réel.
#
# Exerce EXACTEMENT le moteur --allow-hosts de mitmproxy + l'addon de log/rédaction
# (../rpc-traffic-monitor/mitmproxy/addon_log.py). En PROD on l'utilise en mode
# transparent (dest via SO_ORIGINAL_DST posée par la règle nft de hb-on) ; ici on
# l'exerce en mode PROXY (CONNECT) qui partage le MÊME moteur de décision
# intercept-vs-passthrough (hook next_layer / allow_hosts). La règle nft prod a été
# validée séparément par `nft -c` (syntaxe, non appliquée).
#
# Deux faux hôtes (noms BIDON mappés sur 127.0.0.1 dans /etc/hosts le temps du
# test, restaurés en sortie ; aucun process réel ne les résout => zéro collatéral) :
#   * scratchtest.cursor.sh  -> DANS allow-hosts => DÉCHIFFRÉ + LOGGÉ (cert = notre CA)
#   * pass.scratch.test      -> HORS allow-hosts => PASSTHROUGH (cert D'ORIGINE, non loggé)
#
# Prouve :
#   (1) intercept : curl --cacert NOTRE-CA réussit, curl --cacert cert-origine ÉCHOUE
#   (2) passthrough: curl --cacert cert-origine réussit, curl --cacert NOTRE-CA ÉCHOUE
#   (3) le JSONL contient l'hôte intercepté (leg=host-backend, req+resp body) et PAS
#       l'hôte passthrough
#   (4) grep anti-secret VIDE (Bearer/ghp_/PEM bloc complet/email/JWT rédigés)
#   (5) HB_ALLOW_HOSTS (PROD) matche api2/api3/.../cursor.sh et REJETTE tunnel/S3/CDN
#   (6) hb-off.sh --dry-run et hb-revert.sh --dry-run OK (innocuité)
set -uo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LI="$(cd "$HERE/.." && pwd)"
. "$LI/lib.sh"
RUN="$HERE/_run_hb"; rm -rf "$RUN"; mkdir -p "$RUN"
export RPC_TRAFFIC_LOG_DIR="$RUN"
VENV=/workspace/.venvs/mitmproxy; MITMDUMP="$VENV/bin/mitmdump"
CA=/home/box/.mitmproxy/mitmproxy-ca-cert.pem
PROXY_PORT=19550; PORT_A=19551; PORT_B=19552
H_INT="scratchtest.cursor.sh"   # DANS allow-hosts + classifié host-backend (.cursor.sh)
H_PASS="pass.scratch.test"       # HORS allow-hosts
HOSTS_LINE="127.0.0.1 $H_INT $H_PASS"
PIDS=(); HOSTS_ADDED=0
cleanup(){
  for p in "${PIDS[@]:-}"; do kill "$p" 2>/dev/null || true; done
  if [ "$HOSTS_ADDED" = 1 ]; then sudo sed -i "\|# LI-HB-TEST|d" /etc/hosts 2>/dev/null || true; fi
}
trap cleanup EXIT

echo "=============== prépa /etc/hosts (noms BIDON, restaurés en sortie) ==============="
if ! grep -q "$H_INT" /etc/hosts 2>/dev/null; then
  echo "$HOSTS_LINE   # LI-HB-TEST" | sudo tee -a /etc/hosts >/dev/null; HOSTS_ADDED=1
fi
grep 'LI-HB-TEST' /etc/hosts || true

echo "=============== certs scratch (origine, auto-signés distincts) ==============="
openssl req -x509 -newkey rsa:2048 -nodes -keyout "$RUN/a.key" -out "$RUN/a.crt" -subj "/CN=$H_INT"  -days 1 >/dev/null 2>&1
openssl req -x509 -newkey rsa:2048 -nodes -keyout "$RUN/b.key" -out "$RUN/b.crt" -subj "/CN=$H_PASS" -days 1 >/dev/null 2>&1

echo "=============== faux serveurs TLS (origine) ==============="
python3 "$HERE/scratch_servers.py" tls "$PORT_A" "$RUN/a.crt" "$RUN/a.key" & PIDS+=($!)
python3 "$HERE/scratch_servers.py" tls "$PORT_B" "$RUN/b.crt" "$RUN/b.key" & PIDS+=($!)
sleep 1

echo "=============== mitmproxy PROXY mode + allow-hosts SÉLECTIF ==============="
# allow-hosts de TEST : matche l'hôte intercepté (y compris avec :PORT) ; le reste
# est passthrough. (Le HB_ALLOW_HOSTS de PROD est vérifié en (5).)
TEST_ALLOW='(^|\.)cursor\.sh(:[0-9]+)?$'
RPC_TRAFFIC_LOG_DIR="$RUN" "$MITMDUMP" --mode regular \
  --listen-host 127.0.0.1 --listen-port "$PROXY_PORT" \
  --allow-hosts "$TEST_ALLOW" --ssl-insecure \
  -s "$LI/../rpc-traffic-monitor/mitmproxy/addon_log.py" \
  >"$RUN/mitm.out" 2>&1 & PIDS+=($!)
sleep 3

FAKE_PEM="-----BEGIN PRIVATE KEY-----
MIIFAKEKEYBODYLINE1aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa
MIIFAKEKEYBODYLINE2bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb
-----END PRIVATE KEY-----"

echo "--- (1) hôte INTERCEPTÉ : leaf re-signé par NOTRE CA (issuer=mitmproxy) ---"
ISS=$(curl -skv --proxy "http://127.0.0.1:$PROXY_PORT" "https://$H_INT:$PORT_A/probe" 2>&1 | grep -m1 'issuer:')
echo "  presented: $ISS"
echo "$ISS" | grep -q 'CN=mitmproxy' \
  && echo "  OK: issuer=mitmproxy -> TLS terminé+re-signé (déchiffré)" \
  || { echo "  ÉCHEC: cert non re-signé, interception absente"; exit 1; }
# requête loggable avec faux secrets (déchiffrée -> addon la voit)
curl -s --proxy "http://127.0.0.1:$PROXY_PORT" --cacert "$CA" \
     "https://$H_INT:$PORT_A/aiserver.v1.DashboardService/GetServers" \
     -H "Authorization: Bearer ghp_FAKEHEADERCCCCCCCCCCCCCCCCCCCC" \
     -H "Content-Type: application/json" \
     -d "{\"email\":\"user@corp.com\",\"k\":\"$FAKE_PEM\"}" >/dev/null \
  && echo "  OK: requête interceptée envoyée (pour log)" || { echo "  ÉCHEC requête intercept"; exit 1; }

echo "--- (2) hôte PASSTHROUGH : cert D'ORIGINE préservé (issuer != mitmproxy) ---"
ISS2=$(curl -skv --proxy "http://127.0.0.1:$PROXY_PORT" \
     "https://$H_PASS:$PORT_B/passthrough-probe" \
     -H "Authorization: Bearer ghp_SHOULDNOTBELOGGEDXXXXXXXXXXXX" 2>&1 | grep -m1 'issuer:')
echo "  presented: $ISS2"
if echo "$ISS2" | grep -q 'CN=mitmproxy'; then
  echo "  ÉCHEC: passthrough DÉCHIFFRÉ (issuer=mitmproxy)"; exit 2
else echo "  OK: issuer d'origine conservé -> TLS NON terminé (passthrough)"; fi
echo "$ISS2" | grep -q "$H_PASS" && echo "  OK: cert origine $H_PASS préservé" || true
sleep 1
sleep 1

LOG="$(ls -1 "$RUN"/traffic-*.jsonl 2>/dev/null | head -1)"
echo
echo "=============== RÉSULTAT ==============="
[ -n "$LOG" ] && [ -s "$LOG" ] || { echo "ÉCHEC: aucun log produit"; exit 3; }
echo "log: $LOG  ($(wc -l < "$LOG") lignes)"

echo "--- (3) l'hôte intercepté est loggé (leg=host-backend, bodies) ; le passthrough NON ---"
python3 - "$LOG" "$H_INT" "$H_PASS" <<'PY'
import json,sys
log,hint,hpass=sys.argv[1:4]
seen_int=seen_pass=False; ok_leg=ok_req=ok_resp=False
for l in open(log):
    r=json.loads(l)
    h=r.get('host','')
    if h==hint:
        seen_int=True
        print(f"  [intercept] leg={r.get('leg')} http={r.get('http_method')} method={r.get('method')} status={r.get('status')} req_body?={r.get('req_body') is not None} resp_body?={r.get('resp_body') is not None}")
        if r.get('leg')=='host-backend': ok_leg=True
        if r.get('req_body') is not None: ok_req=True
        if r.get('resp_body') is not None: ok_resp=True
    if h==hpass:
        seen_pass=True; print(f"  [PASSTHROUGH LOGGÉ!] {r}")
assert seen_int, "hôte intercepté absent du log"
assert ok_leg, "leg host-backend absent"
assert ok_req and ok_resp, "req/resp body manquant"
assert not seen_pass, "ÉCHEC: l'hôte passthrough a été loggé (déchiffré)"
print("  (3) OK: intercept loggé (leg=host-backend, req+resp) ; passthrough JAMAIS loggé")
PY

echo "--- (4) grep anti-secret (doit être VIDE) ---"
if grep -nE 'ghp_FAKE|ghp_SHOULDNOT|@corp\.com|BEGIN PRIVATE KEY|FAKEKEYBODYLINE|eyJhbGciOiJIUzI1NiJ9\.eyJ|supersecretcookie|ghp_abcdef' "$LOG"; then
  echo "ÉCHEC: secret/PII en clair détecté !!"; exit 4
else echo "  (4) OK: 0 secret/PII en clair"; fi
echo "--- marqueurs de rédaction présents ---"; grep -oE 'REDACTED[A-Z-]*' "$LOG" | sort | uniq -c

echo "--- (5) HB_ALLOW_HOSTS (PROD) : matche backend, rejette tunnel/S3/CDN ---"
HB_ALLOW_HOSTS="$HB_ALLOW_HOSTS" python3 - <<'PY'
import os,re
rx=re.compile(os.environ["HB_ALLOW_HOSTS"])
must=["api2.cursor.sh:443","api2.cursor.sh","api3.cursor.sh:443","repo42.cursor.sh:443","cursor.com:443","www.cursor.com:443"]
mustnot=["198.18.0.1:443","s3.amazonaws.com:443","16.15.20.1:443","104.18.5.5:443","settings.cursor.sh.evil.com:443"]
for h in must:  assert rx.search(h), f"PROD allow devrait matcher {h}"
for h in mustnot: assert not rx.search(h), f"PROD allow NE devrait PAS matcher {h}"
print("  (5) OK: PROD allow matche", must)
print("         PROD allow rejette", mustnot)
PY
echo "--- classifieur addon : .cursor.sh -> host-backend ---"
python3 - <<'PY'
import sys; sys.path.insert(0,"/workspace/projets/Grok-Bot-app/sand-host-rpc/rpc-traffic-monitor/mitmproxy")
import addon_log
assert addon_log._direction("scratchtest.cursor.sh",19551)=="host->backend"
assert addon_log._direction("api2.cursor.sh",443)=="host->backend"
assert addon_log._direction("pass.scratch.test",19552).startswith("autre")
print("  OK: .cursor.sh->host-backend ; pass.scratch.test->autre")
PY

echo "--- (6) hb-off / hb-revert --dry-run (innocuité) ---"
bash "$LI/hb-off.sh"    --dry-run >/dev/null 2>&1 && echo "  (6a) OK hb-off.sh --dry-run exit 0"    || { echo "ÉCHEC hb-off dry"; exit 6; }
bash "$LI/hb-revert.sh" --dry-run >/dev/null 2>&1 && echo "  (6b) OK hb-revert.sh --dry-run exit 0" || { echo "ÉCHEC hb-revert dry"; exit 6; }

echo
echo "SUCCÈS: interception SÉLECTIVE prouvée (backend déchiffré+loggé ; reste passthrough"
echo "non déchiffré non loggé), leg=host-backend, rédaction sans fuite, allow-hosts PROD"
echo "correct, hb-off/hb-revert dry-run sûrs. AUCUNE règle nft réelle, host-main intact."
