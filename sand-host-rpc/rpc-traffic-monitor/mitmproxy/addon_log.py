#!/usr/bin/env python3
"""addon_log.py — addon mitmproxy qui journalise CHAQUE échange (requête ET
réponse) en JSONL, corps complets, secrets/PII rédigés.

Chargé par `mitmdump -s addon_log.py`. Il réutilise TOUTE la logique de rédaction
et de décodage de corps déjà écrite dans ../proxy.py (rédaction headers/JWT/
Bearer/e-mails/clés GH-PEM + valeurs d'env secrètes, dé-framing Connect/gRPC,
décompression gzip/deflate/br/zstd), ainsi que la destination de log partagée
(shared-eg) et la rotation. Aucun secret ni token en clair n'atteint les logs.

Destination : <RPC_TRAFFIC_LOG_DIR>/traffic-YYYY-MM-DD.jsonl
  défaut = /workspace/shared-eg/box-vers-vps/sand-host-rpc-traffic/ (Syncthing EG).

Sens (`direction`) déduit de l'hôte cible :
  * hôtes backend (api2/api3/...cursor.sh, cursor.com, repo42...) -> "host->backend"
  * hôtes locaux / sand-host (127.0.0.1, localhost, ::1, sand-host, port 1340) -> "agent->host"
  * sinon -> "autre:<host>"
"""
from __future__ import annotations
import datetime
import os
import sys

# réutilise proxy.py (dossier parent) : rédaction + décodage + log/rotation
_PARENT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PARENT not in sys.path:
    sys.path.insert(0, _PARENT)
import proxy  # noqa: E402

BACKEND_HOSTS = (
    "api2.cursor.sh", "api3.cursor.sh", "api4.cursor.sh", "api.cursor.sh",
    "repo42.cursor.sh", "cursor.sh", "cursor.com", "www.cursor.com",
)
LOCAL_HOSTS = ("127.0.0.1", "localhost", "::1", "sand-host", "0.0.0.0")


def _direction(host: str, port: int) -> str:
    h = (host or "").lower()
    if any(h == b or h.endswith("." + b) for b in BACKEND_HOSTS):
        return "host->backend"
    if h in LOCAL_HOSTS or port == 1340:
        return "agent->host"
    if h.endswith(".cursor.sh") or h.endswith(".cursor.com"):
        return "host->backend"
    return f"autre:{h}"


class RpcTrafficLogger:
    """Hooks mitmproxy : `request` (horodatage), `response` (journalisation)."""

    def __init__(self):
        proxy._load_env_secrets()
        self.log_dir = proxy.resolve_log_dir()

    def load(self, loader):  # mitmproxy option hook — rien à déclarer
        pass

    def request(self, flow):
        # mitmproxy porte déjà timestamp_start ; on garde un repli.
        try:
            flow.metadata["rtm_t0"] = flow.request.timestamp_start
        except Exception:
            flow.metadata["rtm_t0"] = datetime.datetime.now().timestamp()

    def response(self, flow):
        try:
            self._log(flow)
        except Exception as e:
            try:
                proxy.write_traffic_record(
                    {"ts": datetime.datetime.now().astimezone().isoformat(),
                     "leg": "mitm", "erreur": f"{type(e).__name__}: {e}"},
                    self.log_dir)
            except Exception:
                pass

    # --------------------------------------------------------------------- #
    def _log(self, flow):
        req = flow.request
        resp = flow.response
        host = req.pretty_host or req.host
        port = req.port or (443 if req.scheme == "https" else 80)
        direction = _direction(host, port)
        leg = "host-backend" if direction == "host->backend" else (
            "agent-host" if direction == "agent->host" else "autre")

        # durée ms
        t0 = flow.metadata.get("rtm_t0") or req.timestamp_start
        t1 = resp.timestamp_end or datetime.datetime.now().timestamp()
        dur = round((t1 - t0) * 1000, 1) if t0 else None

        # corps bruts (non décodés par mitmproxy) + content-encoding d'origine
        req_body = req.raw_content or b""
        resp_body = resp.raw_content or b""
        req_ct = req.headers.get("content-type", "")
        req_ce = req.headers.get("content-encoding", "")
        resp_ct = resp.headers.get("content-type", "")
        resp_ce = resp.headers.get("content-encoding", "")

        rec = {
            "ts": datetime.datetime.now().astimezone().isoformat(timespec="milliseconds"),
            "leg": leg,
            "direction": direction,
            "host": host,
            "method": req.path,            # path RPC
            "http_method": req.method,
            "url": f"{req.scheme}://{host}{req.path}",
            "status": str(resp.status_code),
            "duree_ms": dur,
            "req_headers": proxy.redact_headers(list(req.headers.items())),
            "req_body": proxy.decode_body(req_body, req_ct, req_ce),
            "resp_headers": proxy.redact_headers(list(resp.headers.items())),
            "resp_body": proxy.decode_body(resp_body, resp_ct, resp_ce),
        }
        proxy.write_traffic_record(rec, self.log_dir)


addons = [RpcTrafficLogger()]
