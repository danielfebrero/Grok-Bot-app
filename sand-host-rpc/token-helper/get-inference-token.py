#!/usr/bin/env python3
"""Renew Sand box inference access token from the host process env.

Reads SAND_INFERENCE_RENEWAL_CREDENTIAL from the sand-host PID (port SAND_HOST_PORT
or 1340), POSTs to api2.cursor.sh/sand-box/inference-credential, prints ONLY the
token to stdout (or writes via --out). Never logs the credential.

Usage:
  eval "$(python3 /workspace/sand-helpers/get-inference-token.py --export)"
  TOKEN=$(python3 /workspace/sand-helpers/get-inference-token.py)
  python3 /workspace/sand-helpers/get-inference-token.py --out /tmp/tok && chmod 600 /tmp/tok
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import urllib.request
from pathlib import Path

DEFAULT_BACKEND = "https://api2.cursor.sh"
RENEWAL_PATH = "/sand-box/inference-credential"
HOST_PORT_ENV = "SAND_HOST_PORT"


def find_host_pid(port: int) -> int | None:
    try:
        import subprocess
        out = subprocess.check_output(["ss", "-ltnp"], text=True, stderr=subprocess.DEVNULL)
    except Exception:
        out = ""
    for line in out.splitlines():
        if f":{port}" not in line or "users:" not in line:
            continue
        if "pid=" in line:
            try:
                return int(line.split("pid=")[1].split(",")[0].split(")")[0])
            except ValueError:
                pass
    return None


def host_environ(pid: int) -> dict[str, str]:
    raw = open(f"/proc/{pid}/environ", "rb").read().split(b"\0")
    env: dict[str, str] = {}
    for e in raw:
        if b"=" not in e:
            continue
        k, v = e.split(b"=", 1)
        env[k.decode()] = v.decode(errors="replace")
    return env


def renew(credential: str, backend: str, box_namespace: str, proxy: str | None) -> dict:
    headers = {
        "content-type": "application/json",
        "user-agent": "Sand/0.1.0",
        "x-cursor-client-type": "sand",
        "x-cursor-client-version": "0.1.0",
        "x-sand-box-namespace": box_namespace,
        "x-cursor-client-os": "linux",
    }
    req = urllib.request.Request(
        backend.rstrip("/") + RENEWAL_PATH,
        data=json.dumps({"credential": credential}).encode(),
        headers=headers,
        method="POST",
    )
    handlers = []
    if proxy:
        handlers.append(
            urllib.request.ProxyHandler({"http": f"http://{proxy}", "https": f"http://{proxy}"})
        )
    opener = urllib.request.build_opener(*handlers)
    with opener.open(req, timeout=30) as r:
        return json.loads(r.read())


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--backend", default=os.environ.get("GROK_BOT_RPC_BASE_URL", DEFAULT_BACKEND))
    ap.add_argument("--host-port", type=int, default=int(os.environ.get(HOST_PORT_ENV, "1340")))
    ap.add_argument("--pid", type=int, help="sand-host PID (skip ss lookup)")
    ap.add_argument("--out", help="write token to file (0600); else stdout")
    ap.add_argument("--export", action="store_true", help="print GROK_BOT_RPC_AUTH=Bearer … for eval")
    ap.add_argument("--json", action="store_true", help="print {accessToken,expiresAtMs} JSON to stdout")
    ap.add_argument("--no-proxy", action="store_true")
    args = ap.parse_args()

    pid = args.pid or find_host_pid(args.host_port)
    if pid is None:
        print("get-inference-token: sand-host PID not found", file=sys.stderr)
        return 1
    env = host_environ(pid)
    cred = (env.get("SAND_INFERENCE_RENEWAL_CREDENTIAL") or "").strip()
    if not cred:
        print("get-inference-token: SAND_INFERENCE_RENEWAL_CREDENTIAL missing on host", file=sys.stderr)
        return 2
    proxy = None if args.no_proxy else (
        env.get("SAND_EGRESS_TUNNEL_PROXY_ADDR") or os.environ.get("SAND_EGRESS_TUNNEL_PROXY_ADDR")
    )
    ns = env.get("SAND_BOX_TENANT_ID") or ""
    try:
        parsed = renew(cred, args.backend, ns, proxy)
    except Exception as e:
        print(f"get-inference-token: renew failed: {type(e).__name__}: {e}", file=sys.stderr)
        return 3
    token = parsed.get("accessToken") or ""
    if not token:
        print("get-inference-token: empty accessToken", file=sys.stderr)
        return 4
    if args.json:
        sys.stdout.write(json.dumps({"accessToken": token, "expiresAtMs": parsed.get("expiresAtMs")}) + "\n")
        return 0
    if args.export:
        sys.stdout.write(f"export GROK_BOT_RPC_AUTH={json.dumps('Bearer ' + token)}\n")
        sys.stdout.write(f"export GROK_BOT_RPC_BASE_URL={json.dumps(args.backend.rstrip('/'))}\n")
        return 0
    if args.out:
        p = Path(args.out)
        p.write_text(token)
        p.chmod(0o600)
        print(f"wrote token ({len(token)} chars) to {p}", file=sys.stderr)
        return 0
    sys.stdout.write(token)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
