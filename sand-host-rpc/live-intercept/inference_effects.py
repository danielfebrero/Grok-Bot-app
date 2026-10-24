#!/usr/bin/env python3
"""Tours d'inférence = EFFETS visibles sur la box, pas le corps du modèle.

L'inférence Grok Bot tourne dans le cloud Cursor. La box ne voit pas
AiService / StreamChat / les jetons de conversation. Elle voit seulement :

* agent↔host `POST /api/markAgentRequestStarted` et `markAgentRequestEnded`
  (corps `{requestId}` — télémétrie codebase, pas le prompt) ;
* agent↔host SSE `GET /events?channels=transcript,…`
  (événements `{channel, payload}` de type appended / updated / snapshot) ;
* host↔backend `GrokBotService.ListGrokBotTranscriptEntries`
  (request `agentId`, `sessionId`, `limit` ; response `entries[]` avec
  `entryId`, `entryKind`, `body` (JSON d'entrée en bytes), `bodyOmitted`).

Ce module corrèle ces effets en un tour synthétique
(`leg` / `protocol` = `inference-effect`, `kind` = `inference-turn`)
écrit dans le JSONL de trafic déjà rédigé. Pas de journal parallèle.
L'egress :8791 est hors sujet.
"""
from __future__ import annotations

import base64
import datetime
import fcntl
import hashlib
import json
import os
import sys
from urllib.parse import parse_qs, urlparse

_LI = os.path.dirname(os.path.abspath(__file__))
_MONITOR = os.path.join(os.path.dirname(_LI), "rpc-traffic-monitor")
if _MONITOR not in sys.path:
    sys.path.insert(0, _MONITOR)
import proxy  # noqa: E402

LEG = "inference-effect"
PROTOCOL = "inference-effect"
KIND = "inference-turn"
NOTE = (
    "Effets seulement : les jetons de conversation ne passent pas par la box "
    "(pas de corps StreamChat / AiService). Tour reconstitué depuis "
    "markAgentRequestStarted/Ended, le SSE transcript et ListGrokBotTranscriptEntries."
)
# Après markAgentRequestEnded, un List* ou un dernier événement SSE peut
# encore arriver. Au-delà, on ne l'attache pas à ce tour.
GRACE = datetime.timedelta(seconds=5)
MAX_ENTRY_IDS = 40
MAX_SNIPPETS = 8
SNIPPET_CHARS = 180
SSE_CONTENT_CHARS = 2000
SSE_MAX_ENTRIES = 40

PUBLIC_KEYS = (
    "request_id",
    "status",
    "agent_id",
    "agent_ids",
    "started_at",
    "ended_at",
    "duree_ms",
    "transcript_entry_ids",
    "snippets",
    "sse_events",
    "list_responses",
    "effects_only",
    "note",
    "start_outcome",
    "end_outcome",
)


def is_effect_record(rec: dict) -> bool:
    if not isinstance(rec, dict):
        return False
    return (
        rec.get("leg") == LEG
        or rec.get("protocol") == PROTOCOL
        or rec.get("kind") == KIND
    )


def rpc_tail(method: str) -> str:
    if not isinstance(method, str):
        return ""
    path = method.split("?", 1)[0].rstrip("/")
    if not path:
        return ""
    return path.rsplit("/", 1)[-1]


def channels_of(method: str):
    if not isinstance(method, str) or "?" not in method:
        return None
    q = parse_qs(urlparse(method).query)
    raw = (q.get("channels") or [""])[0]
    if not raw:
        return None
    found = [c.strip() for c in raw.split(",") if c.strip()]
    return found or None


def is_correlation_input(rec: dict) -> bool:
    """Vrai si la ligne peut faire partie d'un tour (et doit relancer la corrélation)."""
    if not isinstance(rec, dict) or is_effect_record(rec):
        return False
    kind = rec.get("kind")
    if kind in ("sse-event", "sse-stream"):
        return True
    if rec.get("protocol") == "sse":
        return True
    name = rpc_tail(rec.get("method") or rec.get("path") or "")
    if name in ("markAgentRequestStarted", "markAgentRequestEnded", "ListGrokBotTranscriptEntries"):
        return True
    if name == "events":
        return True
    headers = rec.get("resp_headers") or {}
    if isinstance(headers, dict):
        for k, v in headers.items():
            if str(k).lower() == "content-type" and "text/event-stream" in str(v).lower():
                return True
    return False


def _parse_ts(ts):
    if not isinstance(ts, str) or not ts:
        return None
    try:
        dt = datetime.datetime.fromisoformat(ts)
    except ValueError:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=datetime.timezone.utc)
    return dt


def _iso(dt: datetime.datetime | None):
    if dt is None:
        return None
    return dt.isoformat(timespec="milliseconds")


def unwrap_body(body):
    """JSON déjà décodé, texte JSON, ou messages Connect dé-framés."""
    if isinstance(body, str):
        text = body.strip()
        if not text:
            return None
        try:
            return unwrap_body(json.loads(text))
        except Exception:
            return None
    if isinstance(body, list) and body and all(isinstance(x, dict) and "data" in x for x in body):
        dicts = [x.get("data") for x in body if isinstance(x.get("data"), dict)]
        if not dicts:
            return None
        if len(dicts) == 1:
            return dicts[0]
        merged = {}
        entries = []
        for d in dicts:
            entries.extend(d.get("entries") or [])
            for k, v in d.items():
                if k != "entries":
                    merged[k] = v
        merged["entries"] = entries
        return merged
    if isinstance(body, dict):
        return body
    return None


def _clip(text, n: int):
    if not isinstance(text, str) or len(text) <= n:
        return text
    return text[:n] + "…"


def _shrink_entry(ent: dict) -> dict:
    out = dict(ent)
    for k in ("content", "text"):
        if k in out:
            out[k] = _clip(out[k], SSE_CONTENT_CHARS)
    msg = out.get("message")
    if isinstance(msg, dict):
        msg = dict(msg)
        for k in ("content", "text"):
            if k in msg:
                msg[k] = _clip(msg[k], SSE_CONTENT_CHARS)
        out["message"] = msg
    return out


def shrink_transcript_event(obj):
    """Garde ids + un extrait. Un snapshot complet ne doit pas remplir le JSONL."""
    if not isinstance(obj, dict):
        return obj
    payload = obj.get("payload")
    if not isinstance(payload, dict):
        return obj
    payload = dict(payload)
    if isinstance(payload.get("entry"), dict):
        payload["entry"] = _shrink_entry(payload["entry"])
    entries = payload.get("entries")
    if isinstance(entries, list):
        clipped = []
        for ent in entries[:SSE_MAX_ENTRIES]:
            if isinstance(ent, dict):
                clipped.append(_shrink_entry(ent))
        if len(entries) > SSE_MAX_ENTRIES:
            payload["entries_truncated"] = len(entries) - SSE_MAX_ENTRIES
        payload["entries"] = clipped
    return {**obj, "payload": payload}


def prepare_sse_data(data):
    """Corps d'un événement SSE, rédigé, prêt à être journalisé."""
    if data is None:
        return None
    if isinstance(data, dict):
        return shrink_transcript_event(proxy.redact_json(data))
    if not isinstance(data, str):
        data = str(data)
    try:
        parsed = json.loads(data)
    except Exception:
        parsed = None
    if isinstance(parsed, dict):
        return shrink_transcript_event(proxy.redact_json(parsed))
    return proxy.redact_text(data)[:8000]


def _snippet(text):
    if not isinstance(text, str):
        return None
    text = proxy.redact_text(text)
    text = " ".join(text.split())
    if not text:
        return None
    return _clip(text, SNIPPET_CHARS)


def snippet_from_entry(entry: dict):
    if not isinstance(entry, dict):
        return None
    content = entry.get("content")
    if isinstance(content, str) and content.strip():
        return _snippet(content)
    msg = entry.get("message")
    if isinstance(msg, dict):
        for k in ("content", "text"):
            if isinstance(msg.get(k), str) and msg[k].strip():
                return _snippet(msg[k])
    if isinstance(entry.get("text"), str) and entry["text"].strip():
        return _snippet(entry["text"])
    return None


def decode_entry_body(body):
    """`body` protobuf bytes : base64 Connect, JSON déjà décodé, ou tableau d'octets."""
    if body is None or isinstance(body, bool):
        return None
    if isinstance(body, dict):
        return body
    if isinstance(body, list) and body and all(isinstance(x, int) and 0 <= x <= 255 for x in body[:64]):
        try:
            return decode_entry_body(bytes(body).decode("utf-8", "replace"))
        except Exception:
            return None
    if isinstance(body, str):
        text = body.strip()
        if not text:
            return None
        if text[:1] in "{[":
            try:
                parsed = json.loads(text)
            except Exception:
                parsed = None
            if isinstance(parsed, dict):
                return parsed
        try:
            raw = base64.b64decode(text, validate=False)
            parsed = json.loads(raw.decode("utf-8", "replace"))
        except Exception:
            return None
        return parsed if isinstance(parsed, dict) else None
    return None


def _mark_request_id(rec: dict):
    body = unwrap_body(rec.get("req_body"))
    if not isinstance(body, dict):
        return None
    rid = body.get("requestId") or body.get("request_id")
    if isinstance(rid, str) and rid.strip():
        return rid.strip()
    return None


def _outcome(rec: dict):
    body = unwrap_body(rec.get("resp_body"))
    if isinstance(body, dict):
        outcome = body.get("outcome")
        if isinstance(outcome, str) and outcome:
            return proxy.redact_text(outcome)[:80]
    return None


def _wants_transcript(rec: dict, event_obj) -> bool:
    chans = channels_of(rec.get("method") or "")
    channel = None
    if isinstance(event_obj, dict):
        channel = event_obj.get("channel")
    if chans is not None and "transcript" not in chans:
        return False
    if isinstance(channel, str) and channel and channel != "transcript":
        return False
    if chans is None and channel not in (None, "", "transcript"):
        return False
    # Flux multiplexé ou legacy sans champ channel : on garde seulement
    # ce qui est explicitement transcript, sauf si l'abonnement est transcript
    # et que le payload a un type d'événement connu.
    if channel == "transcript":
        return True
    if chans is not None and "transcript" in chans and isinstance(event_obj, dict):
        payload = event_obj.get("payload")
        if isinstance(payload, dict) and payload.get("type") in ("appended", "updated", "snapshot", "cleared"):
            return True
    return False


def _signals_from_event_obj(event_obj, source: str) -> list[dict]:
    if not isinstance(event_obj, dict):
        return []
    payload = event_obj.get("payload") if "payload" in event_obj else event_obj
    if not isinstance(payload, dict):
        return []
    agent = payload.get("agentId") or payload.get("activeAgentId")
    if not isinstance(agent, str) or not agent:
        agent = None
    found = []

    def one(entry, agent_id):
        if not isinstance(entry, dict):
            return None
        eid = entry.get("id") or entry.get("entryId") or entry.get("entry_id")
        if not isinstance(eid, str) or not eid:
            eid = None
        kind = entry.get("kind") or entry.get("entryKind") or entry.get("entry_kind")
        if not isinstance(kind, str):
            kind = None
        ent_agent = entry.get("agentId") if isinstance(entry.get("agentId"), str) else agent_id
        return {
            "entry_id": eid,
            "entry_kind": kind,
            "agent_id": ent_agent,
            "snippet": snippet_from_entry(entry),
            "source": source,
        }

    if isinstance(payload.get("entry"), dict):
        sig = one(payload["entry"], agent)
        if sig:
            found.append(sig)
    entries = payload.get("entries")
    if isinstance(entries, list):
        for ent in entries:
            sig = one(ent, agent)
            if sig:
                found.append(sig)
    if not found and agent:
        found.append({
            "entry_id": None,
            "entry_kind": payload.get("type") if isinstance(payload.get("type"), str) else None,
            "agent_id": agent,
            "snippet": None,
            "source": source,
        })
    return found


def _iter_sse_events(rec: dict):
    """Produit des objets {channel, payload} depuis une ligne sse-event ou un corps legacy."""
    if rec.get("kind") == "sse-stream":
        return
    body = rec.get("resp_body")
    if rec.get("kind") == "sse-event" or rec.get("protocol") == "sse":
        if isinstance(body, dict):
            yield body
            return
        if isinstance(body, str):
            prepared = prepare_sse_data(body)
            if isinstance(prepared, dict):
                yield prepared
            return
        return
    # Corps legacy : texte SSE entier (éventuellement tronqué par l'ancien décodeur).
    if not isinstance(body, str) or "data:" not in body:
        return
    from sse_stream import SseAssembler
    asm = SseAssembler()
    events = asm.feed(body.encode("utf-8", "replace"))
    events.extend(asm.flush())
    for ev in events:
        prepared = prepare_sse_data(ev.get("data"))
        if isinstance(prepared, dict):
            yield prepared


def signals_from_record(rec: dict) -> list[dict]:
    """Signaux transcript (ids, extraits, agent) portés par une ligne de log."""
    name = rpc_tail(rec.get("method") or "")
    if name == "ListGrokBotTranscriptEntries":
        return _signals_from_list(rec)
    if name in ("markAgentRequestStarted", "markAgentRequestEnded"):
        return []
    if not (
        rec.get("kind") == "sse-event"
        or rec.get("protocol") == "sse"
        or name == "events"
        or is_correlation_input(rec)
    ):
        return []
    out = []
    for obj in _iter_sse_events(rec):
        if not _wants_transcript(rec, obj):
            continue
        out.extend(_signals_from_event_obj(obj, "sse"))
    return out


def _signals_from_list(rec: dict) -> list[dict]:
    req = unwrap_body(rec.get("req_body"))
    resp = unwrap_body(rec.get("resp_body"))
    agent = None
    if isinstance(req, dict):
        raw = req.get("agentId") or req.get("agent_id")
        if isinstance(raw, str) and raw:
            agent = raw
    if not isinstance(resp, dict):
        if agent:
            return [{
                "entry_id": None, "entry_kind": None, "agent_id": agent,
                "snippet": None, "source": "list",
            }]
        return []
    found = []
    entries = resp.get("entries") or []
    if not isinstance(entries, list):
        entries = []
    for ent in entries:
        if not isinstance(ent, dict):
            continue
        eid = ent.get("entryId") or ent.get("entry_id")
        if not isinstance(eid, str) or not eid:
            eid = None
        kind = ent.get("entryKind") or ent.get("entry_kind")
        if not isinstance(kind, str):
            kind = None
        omitted = ent.get("bodyOmitted") is True or ent.get("body_omitted") is True
        decoded = None if omitted else decode_entry_body(ent.get("body"))
        if eid is None and isinstance(decoded, dict):
            raw_id = decoded.get("id")
            if isinstance(raw_id, str):
                eid = raw_id
        if kind is None and isinstance(decoded, dict) and isinstance(decoded.get("kind"), str):
            kind = decoded["kind"]
        ent_agent = agent
        if isinstance(decoded, dict) and isinstance(decoded.get("agentId"), str):
            ent_agent = decoded["agentId"]
        found.append({
            "entry_id": eid,
            "entry_kind": kind,
            "agent_id": ent_agent,
            "snippet": None if omitted else snippet_from_entry(decoded or {}),
            "source": "list",
        })
    if not found and agent:
        found.append({
            "entry_id": None, "entry_kind": None, "agent_id": agent,
            "snippet": None, "source": "list",
        })
    return found


def correlate_turns(records) -> list[dict]:
    """Reconstruit les tours. `records` est chronologique (ancien → récent)."""
    turns = {}
    order = []
    for rec in records:
        if not isinstance(rec, dict) or is_effect_record(rec):
            continue
        name = rpc_tail(rec.get("method") or rec.get("path") or "")
        if name not in ("markAgentRequestStarted", "markAgentRequestEnded"):
            continue
        rid = _mark_request_id(rec)
        if not rid:
            continue
        turn = turns.get(rid)
        if turn is None:
            turn = {
                "request_id": rid,
                "start": None,
                "end": None,
                "start_outcome": None,
                "end_outcome": None,
                "signals": [],
                "sse_events": 0,
                "list_responses": 0,
            }
            turns[rid] = turn
            order.append(rid)
        ts = _parse_ts(rec.get("ts"))
        if name == "markAgentRequestStarted":
            if turn["start"] is None or (ts is not None and ts < turn["start"]):
                turn["start"] = ts
            outcome = _outcome(rec)
            if outcome:
                turn["start_outcome"] = outcome
        else:
            if turn["end"] is None or (ts is not None and (turn["end"] is None or ts > turn["end"])):
                turn["end"] = ts
            outcome = _outcome(rec)
            if outcome:
                turn["end_outcome"] = outcome

    if not turns:
        return []

    for rec in records:
        if not isinstance(rec, dict) or is_effect_record(rec):
            continue
        if not is_correlation_input(rec):
            continue
        name = rpc_tail(rec.get("method") or "")
        if name in ("markAgentRequestStarted", "markAgentRequestEnded"):
            continue
        if rec.get("kind") == "sse-stream":
            continue
        sigs = signals_from_record(rec)
        if not sigs and name != "ListGrokBotTranscriptEntries" and rec.get("kind") != "sse-event":
            continue
        ts = _parse_ts(rec.get("ts"))
        if ts is None:
            continue
        is_list = name == "ListGrokBotTranscriptEntries"
        is_sse = rec.get("kind") == "sse-event" or (
            rec.get("protocol") == "sse" and rec.get("kind") != "sse-stream"
        ) or (name == "events" and rec.get("kind") != "sse-stream")
        for turn in turns.values():
            if not _in_window(ts, turn):
                continue
            if is_list:
                turn["list_responses"] += 1
            elif is_sse and sigs:
                turn["sse_events"] += 1
            turn["signals"].extend(sigs)

    out = []
    for rid in order:
        out.append(_finalize(turns[rid]))
    return out


def _in_window(ts, turn) -> bool:
    start = turn["start"]
    end = turn["end"]
    if start is None and end is None:
        return False
    if start is not None and ts < start:
        return False
    if end is not None and ts > end + GRACE:
        return False
    if start is None and end is not None and ts < end - datetime.timedelta(seconds=120):
        return False
    return True


def _finalize(turn: dict) -> dict:
    start = turn["start"]
    end = turn["end"]
    status = "completed" if start is not None and end is not None else "open"
    duree = None
    if start is not None and end is not None:
        duree = round((end - start).total_seconds() * 1000, 1)
        if duree < 0:
            duree = 0.0
    counts = {}
    ids = []
    seen_ids = set()
    snippets = []
    seen_snips = set()
    for sig in turn["signals"]:
        agent = sig.get("agent_id")
        if isinstance(agent, str) and agent:
            counts[agent] = counts.get(agent, 0) + 1
        eid = sig.get("entry_id")
        if isinstance(eid, str) and eid and eid not in seen_ids:
            if len(ids) < MAX_ENTRY_IDS:
                ids.append(eid)
            seen_ids.add(eid)
        text = sig.get("snippet")
        if text and eid not in seen_snips and len(snippets) < MAX_SNIPPETS:
            snippets.append({
                "entry_id": eid,
                "source": sig.get("source"),
                "entry_kind": sig.get("entry_kind"),
                "text": text,
            })
            if eid:
                seen_snips.add(eid)
    agent_id = None
    if counts:
        agent_id = max(counts.items(), key=lambda kv: (kv[1], kv[0]))[0]
    agent_ids = sorted(counts)
    return {
        "request_id": turn["request_id"],
        "status": status,
        "agent_id": agent_id,
        "agent_ids": agent_ids,
        "started_at": _iso(start),
        "ended_at": _iso(end),
        "duree_ms": duree,
        "transcript_entry_ids": ids,
        "transcript_entry_count": len(seen_ids),
        "snippets": snippets,
        "sse_events": turn["sse_events"],
        "list_responses": turn["list_responses"],
        "effects_only": True,
        "note": NOTE,
        "start_outcome": turn["start_outcome"],
        "end_outcome": turn["end_outcome"],
    }


def effect_record(turn: dict, ts: str | None = None) -> dict:
    summary = {k: turn.get(k) for k in PUBLIC_KEYS}
    summary["transcript_entry_count"] = turn.get("transcript_entry_count")
    blob = json.dumps(summary, sort_keys=True, ensure_ascii=False, default=str)
    digest = hashlib.sha256(blob.encode("utf-8")).hexdigest()[:16]
    return {
        "ts": ts or turn.get("ended_at") or turn.get("started_at"),
        "leg": LEG,
        "protocol": PROTOCOL,
        "kind": KIND,
        "direction": "effects-only",
        "host": turn.get("agent_id") or "",
        "method": "inference-turn",
        "http_method": "EFFECT",
        "status": turn.get("status") or "completed",
        "duree_ms": turn.get("duree_ms"),
        "turn_id": turn.get("request_id"),
        "content_hash": digest,
        "req_body": {
            "requestId": turn.get("request_id"),
            "agentId": turn.get("agent_id"),
        },
        "resp_body": summary,
        "note": NOTE,
    }


def _read_jsonl(path: str) -> list[dict]:
    out = []
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    rec = json.loads(line)
                except Exception:
                    continue
                if isinstance(rec, dict):
                    out.append(rec)
    except OSError:
        return []
    return out


def materialize(log_dir: str) -> int:
    """Ajoute au JSONL du jour les tours *complétés* dont le contenu a changé.

    Idempotent : même `turn_id` + `content_hash` n'est pas réécrit.
    Un List* arrivé après le premier tour ajoute une ligne plus récente.
    """
    path = proxy.logfile_path(log_dir)
    if not os.path.isfile(path):
        return 0
    os.makedirs(log_dir, exist_ok=True)
    lock_path = os.path.join(log_dir, ".inference-effect.lock")
    written = 0
    with open(lock_path, "a") as lockf:
        fcntl.flock(lockf.fileno(), fcntl.LOCK_EX)
        try:
            records = _read_jsonl(path)
            seen = set()
            for rec in records:
                if is_effect_record(rec):
                    seen.add((rec.get("turn_id"), rec.get("content_hash")))
            now = datetime.datetime.now().astimezone().isoformat(timespec="milliseconds")
            for turn in correlate_turns(records):
                if turn.get("status") != "completed":
                    continue
                if not turn.get("request_id"):
                    continue
                # Un tour sans aucun effet transcript reste visible (durée seule)
                # mais on exige au moins les deux marques — déjà le cas si completed.
                rec = effect_record(turn, ts=now)
                key = (rec.get("turn_id"), rec.get("content_hash"))
                if key in seen:
                    continue
                proxy.write_traffic_record(rec, log_dir)
                seen.add(key)
                written += 1
        finally:
            fcntl.flock(lockf.fileno(), fcntl.LOCK_UN)
    return written


def maybe_materialize(log_dir: str, rec: dict) -> int:
    """Relance la corrélation sur une marque, un List*, ou la fin d'un flux SSE.

    Les `sse-event` individuels sont déjà sur disque ; on attend la fin du flux
    ou la marque de fin pour écrire le tour, afin de ne pas relire le JSONL
    à chaque ping.
    """
    if not isinstance(rec, dict) or is_effect_record(rec):
        return 0
    if rec.get("kind") == "sse-event":
        return 0
    if not is_correlation_input(rec):
        return 0
    try:
        return materialize(log_dir)
    except Exception:
        return 0


def main(argv=None) -> int:
    import argparse
    p = argparse.ArgumentParser(description="Matérialise les tours inference-effect dans le JSONL du jour.")
    p.add_argument("--materialize", metavar="LOG_DIR", required=True)
    args = p.parse_args(argv)
    n = materialize(args.materialize)
    print(f"inference-effect: {n} tour(s) ajouté(s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
