#!/usr/bin/env bash
# test-cdp-ws.sh — dry-run cdp-on/off, refus noVNC, capture réelle sur un port
# scratch (pas le Chrome de la box). cdp-on pointe tcpdump + decode_keylog
# --plaintext ; le JSONL a la méthode CDP lisible et les secrets rédigés.
set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LI="$(cd "$HERE/.." && pwd)"
MON="$(cd "$LI/../rpc-traffic-monitor" && pwd)"
RUN="$HERE/_run_cdp"
rm -rf "$RUN"
mkdir -p "$RUN/logs" "$RUN/state"
export LI_STATE_DIR="$RUN/state"
export RPC_TRAFFIC_LOG_DIR="$RUN/logs"
export LI MON
# shellcheck source=../lib.sh
. "$LI/lib.sh"

echo "=============== dry-run cdp-on / cdp-off ==============="
bash "$LI/cdp-on.sh" >"$RUN/dry-on.out"
bash "$LI/cdp-off.sh" --dry-run >"$RUN/dry-off.out"
grep -q "DRY-RUN" "$RUN/dry-on.out"
grep -q "DRY cdp-off" "$RUN/dry-off.out"

echo "=============== 6080 refusé (noVNC) ==============="
if CDP_ON_CONFIRM=yes CDP_PORTS=6080 sudo -n -E bash "$LI/cdp-on.sh" >"$RUN/novnc.out" 2>&1; then
  echo "ÉCHEC: cdp-on a accepté le port 6080"
  cat "$RUN/novnc.out"
  exit 1
fi
if [ -f "$CDP_TCPDUMP_PIDFILE" ] && kill -0 "$(cat "$CDP_TCPDUMP_PIDFILE")" 2>/dev/null; then
  echo "ÉCHEC: tcpdump CDP démarré pour 6080"
  exit 1
fi

command -v tshark >/dev/null
command -v tcpdump >/dev/null
sudo -n true

PORT="$(python3 -c 'import socket;s=socket.socket();s.bind(("127.0.0.1",0));print(s.getsockname()[1]);s.close()')"
echo "=============== capture scratch port $PORT ==============="
SPID=""
cleanup() {
  sudo -n -E bash "$LI/cdp-off.sh" >/dev/null 2>&1 || true
  if [ -n "${SPID:-}" ]; then kill "$SPID" 2>/dev/null || true; fi
}
trap cleanup EXIT

LI="$LI" MON="$MON" python3 - "$PORT" <<'PY' &
import json, os, socket, sys
sys.path.insert(0, os.environ["MON"])
import ws_frames
port = int(sys.argv[1])
srv = socket.socket()
srv.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
srv.bind(("127.0.0.1", port))
srv.listen(1)
srv.settimeout(20)
c, _ = srv.accept()
data = b""
while b"\r\n\r\n" not in data:
    data += c.recv(65536)
c.sendall(
    b"HTTP/1.1 101 Switching Protocols\r\n"
    b"Upgrade: websocket\r\nConnection: Upgrade\r\n"
    b"Sec-WebSocket-Accept: abc\r\n\r\n"
)
buf = b""
c.settimeout(5)
while True:
    try:
        chunk = c.recv(65536)
    except socket.timeout:
        break
    if not chunk:
        break
    buf += chunk
    frames, buf = ws_frames.parse_frames(buf)
    if frames:
        break
msg = json.dumps({"id": 1, "result": {"contact": "user@corp.com"}}).encode()
c.sendall(ws_frames.encode_frame(1, msg, mask=False))
c.sendall(ws_frames.encode_frame(2, b"\x01\x02\x03\x04", mask=False))
c.close()
srv.close()
PY
SPID=$!
sleep 0.2

CDP_ON_CONFIRM=yes CDP_PORTS="$PORT" sudo -n -E bash "$LI/cdp-on.sh" >"$RUN/on.out"

LI="$LI" MON="$MON" python3 - "$PORT" <<'PY'
import json, os, socket, sys
sys.path.insert(0, os.environ["MON"])
import ws_frames
port = int(sys.argv[1])
path = "/devtools/page/SCRATCH"
s = socket.create_connection(("127.0.0.1", port), 3)
s.sendall(
    f"GET {path} HTTP/1.1\r\nHost: 127.0.0.1:{port}\r\n"
    "Upgrade: websocket\r\nConnection: Upgrade\r\n"
    "Sec-WebSocket-Key: dGhlIHNhbXBsZSBub25jZQ==\r\n"
    "Sec-WebSocket-Version: 13\r\n\r\n".encode()
)
data = b""
while b"\r\n\r\n" not in data:
    data += s.recv(4096)
body = json.dumps({
    "id": 1, "method": "Page.navigate",
    "params": {"url": "https://example.com", "contact": "user@corp.com"},
}).encode()
s.sendall(ws_frames.encode_frame(1, body, mask=True, mask_key=b"\x01\x02\x03\x04"))
s.settimeout(2)
try:
    while s.recv(65536):
        pass
except Exception:
    pass
s.close()
PY

sleep 2
sudo -n -E bash "$LI/cdp-off.sh" >"$RUN/off.out"
trap - EXIT
kill "$SPID" 2>/dev/null || true

LOG="$(ls -1 "$RUN/logs"/traffic-*.jsonl 2>/dev/null | head -1 || true)"
if [ -z "${LOG:-}" ] || [ ! -s "$LOG" ]; then
  echo "ÉCHEC: pas de JSONL CDP"
  echo "---- on ----"; cat "$RUN/on.out" || true
  echo "---- off ----"; cat "$RUN/off.out" || true
  echo "---- capture ----"; cat "$CDP_CAPTURE_OUT" 2>/dev/null || true
  exit 1
fi
python3 - "$LOG" <<'PY'
import json, sys
rows = [json.loads(l) for l in open(sys.argv[1]) if l.strip()]
blob = json.dumps(rows)
assert "user@corp.com" not in blob, blob
nav = [r for r in rows if r.get("method") == "Page.navigate"]
assert nav, rows
assert nav[0]["protocol"] == "websocket"
assert nav[0]["leg"] == "host-cdp"
assert nav[0]["req_body"]["params"]["contact"] == "[REDACTED-EMAIL]"
assert any(r.get("ws_opcode") == "binary" and r["resp_body"]["_binary_hex"] == "01020304" for r in rows), rows
assert any(r.get("ws_event") == "handshake" for r in rows)
print(f"  OK {len(rows)} lignes CDP, Page.navigate lisible, binaire 01020304, secrets rédigés")
PY
if [ -e "$CDP_PCAP" ]; then
  echo "ÉCHEC: pcap non shreddé"
  exit 1
fi
echo "SUCCÈS cdp-on/off"
