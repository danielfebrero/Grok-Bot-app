#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dashboard de visualisation EN TEMPS RÉEL du trafic RPC capturé.

- 100% bibliothèque standard Python (http.server, json, threading).
- Lit en TAIL INCRÉMENTAL les JSONL écrits par le moniteur de trafic dans
  RPC_TRAFFIC_LOG_DIR/traffic-YYYY-MM-DD.jsonl (lignes déjà rédactées).
- Expose :
    GET /                 -> index.html
    GET /app.js /style.css-> assets statiques du dossier
    GET /api/stats        -> agrégats JSON (?days=N pour N derniers jours, défaut 1)
- Lecture SEULE des logs. N'écrit jamais dans les JSONL.

Variables d'environnement :
    DASH_PORT            port d'écoute (défaut 8808)
    DASH_HOST            hôte d'écoute (défaut 127.0.0.1)
    RPC_TRAFFIC_LOG_DIR  dossier des JSONL
                         (défaut /workspace/shared-eg/box-vers-vps/sand-host-rpc-traffic)
    DASH_MAX_RECORDS     nb max d'échanges gardés en mémoire (défaut 20000)
"""
import json
import os
import threading
import datetime
from collections import deque, Counter, defaultdict
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_LOG_DIR = "/workspace/shared-eg/box-vers-vps/sand-host-rpc-traffic"
LOG_DIR = os.environ.get("RPC_TRAFFIC_LOG_DIR", DEFAULT_LOG_DIR)
PORT = int(os.environ.get("DASH_PORT", "8808"))
HOST = os.environ.get("DASH_HOST", "127.0.0.1")
MAX_RECORDS = int(os.environ.get("DASH_MAX_RECORDS", "20000"))
BODY_TRUNC = 2000  # octets max de corps renvoyés à l'affichage

# ------------------------------------------------------------------ état tail
_lock = threading.Lock()
_records = deque(maxlen=MAX_RECORDS)          # échanges parsés (dict)
_offsets = {}                                 # path -> (inode, byte_offset)


def _log_path(day):
    return os.path.join(LOG_DIR, "traffic-%s.jsonl" % day.strftime("%Y-%m-%d"))


def _days_wanted(n):
    today = datetime.date.today()
    return [today - datetime.timedelta(days=i) for i in range(max(1, n))]


def _tail_file(path):
    """Lit uniquement les octets ajoutés depuis le dernier passage."""
    try:
        st = os.stat(path)
    except OSError:
        return
    inode = st.st_ino
    size = st.st_size
    prev_inode, prev_off = _offsets.get(path, (None, 0))
    if prev_inode != inode or size < prev_off:
        # fichier neuf / tronqué / tourné : on repart de zéro
        prev_off = 0
    if size == prev_off:
        _offsets[path] = (inode, size)
        return
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            f.seek(prev_off)
            chunk = f.read()
            new_off = f.tell()
    except OSError:
        return
    for line in chunk.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            rec = json.loads(line)
        except Exception:
            continue
        _records.append(rec)
    _offsets[path] = (inode, new_off)


def refresh(days):
    """Met à jour le buffer mémoire depuis les fichiers des N derniers jours."""
    with _lock:
        for d in _days_wanted(days):
            _tail_file(_log_path(d))


def _trunc(v):
    if v is None:
        return None
    if not isinstance(v, str):
        try:
            v = json.dumps(v, ensure_ascii=False)
        except Exception:
            v = str(v)
    if len(v) > BODY_TRUNC:
        return v[:BODY_TRUNC] + "…[tronqué %d o]" % (len(v) - BODY_TRUNC)
    return v


def _percentile(sorted_vals, pct):
    if not sorted_vals:
        return 0.0
    k = (len(sorted_vals) - 1) * (pct / 100.0)
    lo = int(k)
    hi = min(lo + 1, len(sorted_vals) - 1)
    if lo == hi:
        return round(sorted_vals[lo], 1)
    frac = k - lo
    return round(sorted_vals[lo] * (1 - frac) + sorted_vals[hi] * frac, 1)


def _parse_ts(ts):
    if not ts:
        return None
    try:
        return datetime.datetime.fromisoformat(ts)
    except Exception:
        return None


def compute_stats(days, last_n):
    with _lock:
        recs = list(_records)

    total = len(recs)
    by_leg = Counter()
    methods = Counter()
    statuses = Counter()
    hosts = Counter()
    latencies = []
    per_min = defaultdict(int)
    errors = 0

    for r in recs:
        leg = r.get("leg") or "?"
        by_leg[leg] += 1
        m = r.get("method") or r.get("path") or "?"
        methods[m] += 1
        st = str(r.get("status") or "?")
        statuses[st] += 1
        if r.get("host"):
            hosts[r["host"]] += 1
        d = r.get("duree_ms")
        if isinstance(d, (int, float)):
            latencies.append(float(d))
        # taux d'erreur : status >= 400 ou non numérique "error"
        try:
            if int(st) >= 400:
                errors += 1
        except ValueError:
            if st.lower() in ("error", "err"):
                errors += 1
        dt = _parse_ts(r.get("ts"))
        if dt:
            bucket = dt.strftime("%Y-%m-%dT%H:%M")
            per_min[bucket] += 1

    latencies.sort()
    # timeline : derniers 30 buckets-minute
    timeline = sorted(per_min.items())
    timeline = timeline[-30:]

    # échanges/min courant = dernier bucket non vide
    rate_now = timeline[-1][1] if timeline else 0

    last = recs[-last_n:][::-1]  # plus récents d'abord
    last_out = []
    for r in last:
        last_out.append({
            "ts": r.get("ts"),
            "leg": r.get("leg"),
            "direction": r.get("direction"),
            "host": r.get("host"),
            "method": r.get("method") or r.get("path"),
            "http_method": r.get("http_method"),
            "status": r.get("status"),
            "duree_ms": r.get("duree_ms"),
            "req_headers": r.get("req_headers"),
            "req_body": _trunc(r.get("req_body")),
            "resp_headers": r.get("resp_headers"),
            "resp_body": _trunc(r.get("resp_body")),
        })

    return {
        "generated_at": datetime.datetime.now().astimezone().isoformat(timespec="seconds"),
        "log_dir": LOG_DIR,
        "days": days,
        "total": total,
        "error_count": errors,
        "error_rate": round(100.0 * errors / total, 1) if total else 0.0,
        "rate_now": rate_now,
        "by_leg": dict(by_leg),
        "top_methods": methods.most_common(12),
        "status_dist": sorted(statuses.items()),
        "hosts": hosts.most_common(10),
        "latency": {
            "p50": _percentile(latencies, 50),
            "p95": _percentile(latencies, 95),
            "max": round(latencies[-1], 1) if latencies else 0.0,
            "count": len(latencies),
        },
        "timeline": timeline,
        "last": last_out,
    }


# ------------------------------------------------------------------ http
STATIC = {
    "/": ("index.html", "text/html; charset=utf-8"),
    "/index.html": ("index.html", "text/html; charset=utf-8"),
    "/app.js": ("app.js", "application/javascript; charset=utf-8"),
    "/style.css": ("style.css", "text/css; charset=utf-8"),
}


class Handler(BaseHTTPRequestHandler):
    server_version = "RpcDash/1.0"

    def log_message(self, fmt, *args):
        pass  # silencieux

    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        try:
            self.wfile.write(body)
        except BrokenPipeError:
            pass

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        if path == "/api/stats":
            q = parse_qs(parsed.query)
            days = int(q.get("days", ["1"])[0])
            last_n = int(q.get("last", ["50"])[0])
            days = max(1, min(days, 14))
            last_n = max(1, min(last_n, 500))
            refresh(days)
            try:
                stats = compute_stats(days, last_n)
                self._send(200, json.dumps(stats, ensure_ascii=False))
            except Exception as e:
                self._send(500, json.dumps({"error": str(e)}))
            return
        if path in STATIC:
            fname, ctype = STATIC[path]
            fp = os.path.join(HERE, fname)
            try:
                with open(fp, "rb") as f:
                    self._send(200, f.read(), ctype)
            except OSError:
                self._send(404, json.dumps({"error": "introuvable: %s" % fname}))
            return
        self._send(404, json.dumps({"error": "not found"}))


def main():
    # amorçage : lit tout l'historique du jour au démarrage
    refresh(1)
    srv = ThreadingHTTPServer((HOST, PORT), Handler)
    print("Dashboard RPC sur http://%s:%d  (logs: %s)" % (HOST, PORT, LOG_DIR), flush=True)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        srv.shutdown()


if __name__ == "__main__":
    main()
