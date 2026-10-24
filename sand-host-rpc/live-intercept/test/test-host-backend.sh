#!/usr/bin/env bash
# test-host-backend.sh — PREUVE du déchiffrement host↔backend par keylog NSS.
# Cibles SCRATCH uniquement. NE TOUCHE NI host-main, NI api2 réel, NI nft.
#
# Ce node (/exec-daemon/node v22) n'écrit PAS de keylog si seul SSLKEYLOGFILE
# est posé. Il l'écrit avec --tls-keylog (même format NSS que SSLKEYLOGFILE).
# La preuve : tcpdump + tshark (decode_keylog.py) sur trois faux hôtes :
#   * scratchtest.cursor.sh  HTTP/1.1  -> journalisé leg=host-backend, rédigé
#   * api2.cursor.sh          HTTP/2   -> journalisé leg=host-backend, rédigé
#   * pass.scratch.test       HTTP/1.1  -> déchiffrable mais NON journalisé
#     (hors HB_ALLOW_HOSTS ; marqueur unique absent du JSONL)
#
# Puis : dry-run de hb-setup-keylog / hb-on / hb-off / hb-revert, et
# arm-agent-host.sh (toujours dry, la patte agent↔host n'est pas armée).
set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LI="$(cd "$HERE/.." && pwd)"
RUN="$HERE/_run_hb"
rm -rf "$RUN"
mkdir -p "$RUN/logs" "$RUN/state"
unset LI_ARM_CONFIRM HB_ON_CONFIRM HB_FORCE || true
export LI_STATE_DIR="$RUN/state"
export RPC_TRAFFIC_LOG_DIR="$RUN/logs"
# AVANT lib.sh : HB_IFACE="${HB_IFACE:-any}" fige la valeur au source.
# Le scratch est sur lo ; any marche aussi, mais lo évite le cooked LINUX_SLL2.
export HB_IFACE=lo
# shellcheck source=../lib.sh
. "$LI/lib.sh"
PA=19561; PB=19562; PC=19563
export HB_PCAP="$RUN/cap.pcap"
export HB_KEYLOG_FILE="$RUN/keys.log"
export HB_CAPTURE_FILTER="tcp port $PA or tcp port $PB or tcp port $PC"
export HB_TLS_PORTS="$PA $PB $PC"
export HB_DECODE_INTERVAL=0.5
export HB_ALLOW_HOSTS

if [ -x /exec-daemon/node ]; then NODE=/exec-daemon/node
elif command -v node >/dev/null 2>&1; then NODE="$(command -v node)"
else echo "ÉCHEC: node introuvable"; exit 1; fi
command -v tshark >/dev/null 2>&1 || { echo "ÉCHEC: tshark introuvable"; exit 1; }
command -v tcpdump >/dev/null 2>&1 || { echo "ÉCHEC: tcpdump introuvable"; exit 1; }
sudo -n true 2>/dev/null || { echo "ÉCHEC: sudo -n indisponible (tcpdump)"; exit 1; }

PIDS=()
cleanup() {
  sudo -n -E bash "$LI/hb-off.sh" >/dev/null 2>&1 || true
  for p in "${PIDS[@]:-}"; do kill "$p" 2>/dev/null || true; done
}
trap cleanup EXIT

echo "=============== (0) ce node ignore SSLKEYLOGFILE seul ==============="
rm -f "$RUN/env-only.log"
# serveur jetable le temps de la sonde env
openssl req -x509 -newkey rsa:2048 -nodes -keyout "$RUN/s.key" -out "$RUN/s.crt" \
  -subj "/CN=scratch" -days 1 >/dev/null 2>&1
python3 "$HERE/scratch_servers.py" tls "$PA" "$RUN/s.crt" "$RUN/s.key" >/dev/null 2>&1 &
PIDS+=($!)
sleep 0.4
SSLKEYLOGFILE="$RUN/env-only.log" "$NODE" -e '
const https=require("https");
const req=https.get({host:"127.0.0.1",port:process.argv[1],rejectUnauthorized:false,path:"/"},res=>{res.resume(); res.on("end",()=>setTimeout(()=>process.exit(0),150));});
req.on("error",e=>{console.error(e); process.exit(1);});
' "$PA"
if [ -s "$RUN/env-only.log" ]; then
  echo "  NOTE: ce binaire écrit aussi via SSLKEYLOGFILE ($(wc -c < "$RUN/env-only.log") octets)."
  echo "        hb-setup pose quand même --tls-keylog (chemin qui marche ici)."
else
  echo "  OK: SSLKEYLOGFILE seul ne crée pas de keylog sur $($NODE -v)"
fi
kill "${PIDS[0]}" 2>/dev/null || true
PIDS=()
sleep 0.2

echo "=============== serveurs scratch (h1 x2 + h2) ==============="
python3 "$HERE/scratch_servers.py" tls "$PA" "$RUN/s.crt" "$RUN/s.key" >/dev/null 2>&1 &
PIDS+=($!)
python3 "$HERE/scratch_servers.py" tls "$PB" "$RUN/s.crt" "$RUN/s.key" >/dev/null 2>&1 &
PIDS+=($!)
cat > "$RUN/h2.js" <<'JS'
const fs = require("fs");
const http2 = require("http2");
const srv = http2.createSecureServer({
  key: fs.readFileSync(process.argv[2]),
  cert: fs.readFileSync(process.argv[3]),
});
srv.on("stream", (stream) => {
  stream.on("data", () => {});
  stream.on("end", () => {
    const out = Buffer.from(JSON.stringify({
      note: "h2",
      email: "h2resp@corp.com",
      token: "eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJ0ZXN0LXNlY3JldC1wYXlsb2FkLXh4eCJ9.Zm9vYmFyYmF6cXV4c2VjcmV0c2ln",
    }));
    stream.respond({
      ":status": 200,
      "content-type": "application/json",
      "content-length": out.length,
      "set-cookie": "session=supersecretcookievalue123456",
    });
    stream.end(out);
  });
});
srv.listen(Number(process.argv[4]), "127.0.0.1");
JS
"$NODE" "$RUN/h2.js" "$RUN/s.key" "$RUN/s.crt" "$PC" >/dev/null 2>&1 &
PIDS+=($!)
sleep 0.4

install -m 0600 /dev/null "$HB_KEYLOG_FILE"
echo "=============== capture passive (script de prod, ports scratch) ==============="
sudo -n -E \
  LI_STATE_DIR="$LI_STATE_DIR" \
  RPC_TRAFFIC_LOG_DIR="$RPC_TRAFFIC_LOG_DIR" \
  HB_PCAP="$HB_PCAP" \
  HB_KEYLOG_FILE="$HB_KEYLOG_FILE" \
  HB_IFACE="$HB_IFACE" \
  HB_CAPTURE_FILTER="$HB_CAPTURE_FILTER" \
  HB_TLS_PORTS="$HB_TLS_PORTS" \
  HB_ALLOW_HOSTS="$HB_ALLOW_HOSTS" \
  HB_DECODE_INTERVAL="$HB_DECODE_INTERVAL" \
  bash "$LI/start-capture-host-backend.sh" >"$RUN/capture.out" 2>&1 &
CAP=$!
ready=0
for _ in $(seq 1 50); do
  [ -f "$HB_CAPTURE_READY" ] && { ready=1; break; }
  if ! kill -0 "$CAP" 2>/dev/null; then break; fi
  sleep 0.1
done
if [ "$ready" != 1 ]; then
  echo "ÉCHEC: capture non prête"; echo "---- capture.out ----"; cat "$RUN/capture.out" || true
  exit 1
fi
echo "  capture prête"

echo "=============== client node --tls-keylog (3 hôtes) ==============="
"$NODE" --tls-keylog="$HB_KEYLOG_FILE" - "$PA" "$PB" "$PC" <<'JS'
const https = require("https");
const http2 = require("http2");
const pa = Number(process.argv[2]);
const pb = Number(process.argv[3]);
const pc = Number(process.argv[4]);
const pem = "-----BEGIN PRIVATE KEY-----\n"
  + "MIIFAKEKEYBODYLINE1aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa\n"
  + "MIIFAKEKEYBODYLINE2bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb\n"
  + "-----END PRIVATE KEY-----";
function post(port, servername, path, headers, obj) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify(obj);
    const req = https.request({
      host: "127.0.0.1", port, servername, method: "POST", path,
      rejectUnauthorized: false,
      headers: Object.assign({
        "Content-Type": "application/json",
        "Content-Length": Buffer.byteLength(body),
      }, headers),
    }, (res) => { res.resume(); res.on("end", resolve); });
    req.on("error", reject);
    req.end(body);
  });
}
(async () => {
  await post(pa, "scratchtest.cursor.sh", "/aiserver.v1.DashboardService/GetServers",
    { Authorization: "Bearer ghp_FAKEHEADERCCCCCCCCCCCCCCCCCCCC" },
    { email: "user@corp.com", k: pem });
  await post(pb, "pass.scratch.test", "/passthrough-probe",
    { Authorization: "Bearer ghp_SHOULDNOTBELOGGEDXXXXXXXXXXXX" },
    { email: "passonly@corp.com", marker: "PASSTHROUGH_ONLY_MARKER_9f3a" });
  await new Promise((resolve, reject) => {
    const c = http2.connect("https://127.0.0.1:" + pc, {
      rejectUnauthorized: false,
      servername: "api2.cursor.sh",
    });
    c.on("error", reject);
    const body = JSON.stringify({ email: "h2req@corp.com" });
    const r = c.request({
      ":method": "POST",
      ":path": "/h2probe",
      ":authority": "api2.cursor.sh",
      "content-type": "application/json",
      authorization: "Bearer ghp_FAKEHEADERCCCCCCCCCCCCCCCCCCCC",
    });
    r.on("error", reject);
    r.on("response", () => {});
    r.resume();
    r.on("end", () => { c.close(); resolve(); });
    r.end(body);
  });
  setTimeout(() => process.exit(0), 200);
})().catch((e) => { console.error(e); process.exit(1); });
JS
echo "  client terminé ; labels keylog (pas les secrets) :"
awk '{print $1}' "$HB_KEYLOG_FILE" | sort | uniq -c
perm="$(stat -c %a "$HB_KEYLOG_FILE")"
echo "  mode keylog=$perm"
[ "$perm" = "600" ] || { echo "ÉCHEC: keylog pas en 0600"; exit 1; }

echo "=============== hb-off (passe finale + destruction du pcap) ==============="
sudo -n -E \
  LI_STATE_DIR="$LI_STATE_DIR" \
  RPC_TRAFFIC_LOG_DIR="$RPC_TRAFFIC_LOG_DIR" \
  HB_PCAP="$HB_PCAP" \
  HB_KEYLOG_FILE="$HB_KEYLOG_FILE" \
  bash "$LI/hb-off.sh"
if [ -e "$HB_PCAP" ]; then echo "ÉCHEC: pcap encore présent après hb-off"; exit 1; fi
echo "  pcap détruit, keylog conservé (setup reste armé côté fichier)"
[ -s "$HB_KEYLOG_FILE" ] || { echo "ÉCHEC: keylog disparu (hb-off ne doit pas le shredder)"; exit 1; }

LOG="$(ls -1 "$RPC_TRAFFIC_LOG_DIR"/traffic-*.jsonl 2>/dev/null | head -1)"
echo
echo "=============== RÉSULTAT ==============="
[ -n "${LOG:-}" ] && [ -s "$LOG" ] || {
  echo "ÉCHEC: aucun JSONL"; echo "---- capture.out ----"; cat "$RUN/capture.out" || true
  echo "---- tcpdump ----"; sudo -n cat "$LI_STATE_DIR/hb-capture.out" 2>/dev/null || true
  echo "---- state log ----"; cat "$LI_STATE_DIR/state.log" 2>/dev/null || true
  exit 3
}
echo "log: $LOG  ($(wc -l < "$LOG") lignes)"

echo "--- (1)(2)(3) hôtes allow journalisés (h1+h2), passthrough absent ---"
python3 - "$LOG" <<'PY'
import json, sys
log = sys.argv[1]
rows = [json.loads(l) for l in open(log) if l.strip()]
hosts = {}
for r in rows:
    hosts.setdefault(r.get("host"), []).append(r)
    print(f"  host={r.get('host')} leg={r.get('leg')} http={r.get('http_method')} method={r.get('method')} status={r.get('status')} req={r.get('req_body') is not None} resp={r.get('resp_body') is not None}")
def need(host, path):
    rs = hosts.get(host) or []
    ok = [r for r in rs if r.get("method") == path and r.get("leg") == "host-backend"
          and r.get("req_body") is not None and r.get("resp_body") is not None
          and str(r.get("status")) == "200"]
    assert ok, f"manque {host} {path}"
need("scratchtest.cursor.sh", "/aiserver.v1.DashboardService/GetServers")
need("api2.cursor.sh", "/h2probe")
assert "pass.scratch.test" not in hosts, "hôte hors allow journalisé"
assert "127.0.0.1" not in hosts, "IP locale utilisée comme host (SNI ignoré)"
print("  (1)(2)(3) OK: h1 scratchtest + h2 api2 journalisés ; pass.scratch.test absent")
PY

echo "--- (4) grep anti-secret (doit être VIDE) ---"
if grep -nE 'ghp_FAKE|ghp_SHOULDNOT|PASSTHROUGH_ONLY_MARKER_9f3a|@corp\.com|@example\.com|BEGIN PRIVATE KEY|FAKEKEYBODYLINE|eyJhbGciOiJIUzI1NiJ9\.eyJ|supersecretcookie|CLIENT_TRAFFIC_SECRET|CLIENT_HANDSHAKE_TRAFFIC_SECRET' "$LOG"; then
  echo "ÉCHEC: secret/PII ou matériel de clé en clair"
  exit 4
fi
echo "  (4) OK: 0 secret/PII/keylog en clair"
echo "--- marqueurs de rédaction ---"
grep -oE 'REDACTED[A-Z-]*' "$LOG" | sort | uniq -c

echo "--- (5) HB_ALLOW_HOSTS (PROD) : matche backend, rejette tunnel/S3/CDN/spoof ---"
HB_ALLOW_HOSTS="$HB_ALLOW_HOSTS" python3 - <<'PY'
import os, re
rx = re.compile(os.environ["HB_ALLOW_HOSTS"])
must = ["api2.cursor.sh:443","api2.cursor.sh","api3.cursor.sh:443","repo42.cursor.sh:443",
        "cursor.com:443","www.cursor.com:443","scratchtest.cursor.sh"]
mustnot = ["198.18.0.1:443","s3.amazonaws.com:443","16.15.20.1:443","104.18.5.5:443",
           "settings.cursor.sh.evil.com:443","pass.scratch.test"]
for h in must:
    assert rx.search(h), f"PROD allow devrait matcher {h}"
for h in mustnot:
    assert not rx.search(h), f"PROD allow NE devrait PAS matcher {h}"
print("  (5) OK")
PY

echo "--- classifieur addon : .cursor.sh -> host-backend ---"
python3 - <<PY
import sys
sys.path.insert(0, "$LI/../rpc-traffic-monitor/mitmproxy")
import addon_log
assert addon_log._direction("scratchtest.cursor.sh", $PA) == "host->backend"
assert addon_log._direction("api2.cursor.sh", 443) == "host->backend"
assert addon_log._direction("pass.scratch.test", $PB).startswith("autre")
print("  OK")
PY

echo "--- (6) dry-run setup/on/off/revert + stub mitm retiré + agent-host intact ---"
out="$(bash "$LI/hb-setup-keylog.sh")"
echo "$out" | grep -q 'DRY-RUN' && echo "$out" | grep -q -- '--tls-keylog' \
  && echo "  (6a) OK hb-setup-keylog dry-run" || { echo "ÉCHEC setup dry"; echo "$out"; exit 6; }
echo "$out" | grep -q 'NODE_EXTRA_CA_CERTS' \
  && echo "  (6a) le plan mentionne le retrait de NODE_EXTRA_CA_CERTS" || true
bash "$LI/hb-on.sh" >/dev/null
echo "  (6b) OK hb-on dry-run"
bash "$LI/hb-off.sh" --dry-run >/dev/null
echo "  (6c) OK hb-off dry-run"
bash "$LI/hb-revert.sh" --dry-run >/dev/null
echo "  (6d) OK hb-revert dry-run"
bash "$LI/hb-setup-ca.sh" >/dev/null 2>&1 && { echo "ÉCHEC: stub CA aurait dû sortir 1"; exit 6; }
echo "  (6e) OK hb-setup-ca.sh retiré (exit 1)"
bash "$LI/start-mitm-host-backend.sh" >/dev/null 2>&1 && { echo "ÉCHEC: stub mitm aurait dû sortir 1"; exit 6; }
echo "  (6f) OK start-mitm-host-backend.sh retiré (exit 1)"
# patte agent↔host : dry-run seulement, aucun nft, host-main non requis pour le plan
bash "$LI/arm-agent-host.sh" >/dev/null
echo "  (6g) OK arm-agent-host.sh dry-run (inchangé, non armé)"

# le keylog de test ne doit pas avoir fuité dans un autre dossier que RUN
if find "$RUN/logs" -iname '*keylog*' -o -name 'sslkeys.log' | grep -q .; then
  echo "ÉCHEC: keylog recopié dans les logs"; exit 4
fi

echo
echo "SUCCÈS: keylog NSS + tshark déchiffrent h1 et h2 sur cibles scratch,"
echo "leg=host-backend, rédaction sans fuite, hôte hors-allow non journalisé,"
echo "pcap détruit par hb-off, dry-run sûrs, mitm/CA retirés, agent↔host non armé."
