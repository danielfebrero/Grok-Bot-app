#!/usr/bin/env python3
"""Tours inference-effect : fixtures mark* + transcript, flux SSE complet, dashboard.

N'arme rien. N'écrit pas dans shared-eg. Les secrets du fixture ne doivent pas
ressortir dans le tour (extraits rédigés, bodyOmitted ignoré).
"""
from __future__ import annotations

import asyncio
import importlib.util
import json
import os
import socket
import subprocess
import sys
import tempfile
import threading
import time
import unittest
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
LI = os.path.dirname(HERE)
ROOT = os.path.dirname(LI)
if LI not in sys.path:
    sys.path.insert(0, LI)

import inference_effects  # noqa: E402
import proxy  # noqa: E402
import sse_stream  # noqa: E402

FIXTURE = os.path.join(HERE, "fixtures", "inference_turn.json")
RELAY = os.path.join(LI, "relay_agent_host.py")
DASH = os.path.join(ROOT, "dashboard", "server.py")
EMAIL = "victim@example.com"
JWT = "eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJ0ZXN0LXNlY3JldC1wYXlsb2FkLXh4eCJ9.Zm9vYmFyYmF6cXV4c2VjcmV0c2ln"
OMITTED = "SECRET-OMITTED-BODY-do-not-surface"


def load_fixture():
    with open(FIXTURE, encoding="utf-8") as f:
        return json.load(f)


def assert_redacted(blob: str) -> None:
    if EMAIL in blob or JWT in blob or OMITTED in blob:
        raise AssertionError("secret ou corps omis présent dans la sortie:\n" + blob[:500])


class CorrelateTests(unittest.TestCase):
    def test_completed_turn_ties_transcript(self):
        turns = inference_effects.correlate_turns(load_fixture())
        by_id = {t["request_id"]: t for t in turns}
        self.assertIn("req-7f3a", by_id)
        self.assertIn("req-open", by_id)
        turn = by_id["req-7f3a"]
        self.assertEqual(turn["status"], "completed")
        self.assertEqual(turn["agent_id"], "agent-bc-42")
        self.assertEqual(turn["duree_ms"], 3500.0)
        self.assertEqual(turn["started_at"][:19], "2026-10-24T09:00:00")
        self.assertEqual(turn["ended_at"][:19], "2026-10-24T09:00:03")
        self.assertIn("ent-9c2", turn["transcript_entry_ids"])
        self.assertIn("ent-omit", turn["transcript_entry_ids"])
        self.assertNotIn("ent-noise", turn["transcript_entry_ids"])
        self.assertGreaterEqual(turn["sse_events"], 1)
        self.assertGreaterEqual(turn["list_responses"], 1)
        self.assertTrue(turn["effects_only"])
        self.assertIn("StreamChat", turn["note"])
        snips = {s["entry_id"]: s for s in turn["snippets"]}
        self.assertIn("ent-9c2", snips)
        self.assertNotIn("ent-omit", snips)
        self.assertIn("diff", snips["ent-9c2"]["text"])
        self.assertIn("[REDACTED-EMAIL]", snips["ent-9c2"]["text"])
        self.assertIn("[REDACTED-JWT]", snips["ent-9c2"]["text"])
        assert_redacted(json.dumps(turn, ensure_ascii=False))
        open_turn = by_id["req-open"]
        self.assertEqual(open_turn["status"], "open")
        self.assertEqual(open_turn["transcript_entry_ids"], [])
        self.assertIsNone(open_turn["duree_ms"])

    def test_framed_list_and_grace(self):
        records = [
            {
                "ts": "2026-10-24T11:00:00.000+00:00",
                "method": "/api/markAgentRequestStarted",
                "req_body": {"requestId": "req-grace"},
                "resp_body": {"outcome": "accepted"},
            },
            {
                "ts": "2026-10-24T11:00:01.000+00:00",
                "method": "/api/markAgentRequestEnded",
                "req_body": {"requestId": "req-grace"},
                "resp_body": {"outcome": "accepted"},
            },
            {
                "ts": "2026-10-24T11:00:04.000+00:00",
                "leg": "host-backend",
                "method": "/aiserver.v1.GrokBotService/ListGrokBotTranscriptEntries",
                "req_body": [{"flag": 0, "data": {"agentId": "agent-z", "limit": 10, "sessionId": "s"}}],
                "resp_body": [{
                    "flag": 0,
                    "data": {
                        "generation": 1,
                        "entries": [{
                            "seq": "1",
                            "entryKind": "message",
                            "entryId": "ent-grace",
                            "bodyOmitted": False,
                            "body": {"id": "ent-grace", "kind": "message", "content": "après la fin"},
                        }],
                    },
                }],
            },
        ]
        turns = inference_effects.correlate_turns(records)
        self.assertEqual(len(turns), 1)
        turn = turns[0]
        self.assertEqual(turn["status"], "completed")
        self.assertEqual(turn["agent_id"], "agent-z")
        self.assertEqual(turn["transcript_entry_ids"], ["ent-grace"])
        self.assertEqual(turn["snippets"][0]["text"], "après la fin")
        self.assertEqual(turn["snippets"][0]["source"], "list")
        self.assertEqual(turn["list_responses"], 1)

    def test_materialize_idempotent_then_updates(self):
        records = load_fixture()
        with tempfile.TemporaryDirectory() as tmp:
            path = proxy.logfile_path(tmp)
            with open(path, "w", encoding="utf-8") as f:
                for rec in records:
                    f.write(json.dumps(rec, ensure_ascii=False) + "\n")
            self.assertEqual(inference_effects.materialize(tmp), 1)
            self.assertEqual(inference_effects.materialize(tmp), 0)
            with open(path, encoding="utf-8") as fh:
                lines = [json.loads(l) for l in fh if l.strip()]
            effects = [r for r in lines if r.get("leg") == "inference-effect"]
            self.assertEqual(len(effects), 1)
            eff = effects[0]
            self.assertEqual(eff["protocol"], "inference-effect")
            self.assertEqual(eff["kind"], "inference-turn")
            self.assertEqual(eff["turn_id"], "req-7f3a")
            self.assertEqual(eff["status"], "completed")
            self.assertEqual(eff["method"], "inference-turn")
            body = eff["resp_body"]
            self.assertEqual(body["agent_id"], "agent-bc-42")
            self.assertIn("ent-9c2", body["transcript_entry_ids"])
            self.assertIn("ent-omit", body["transcript_entry_ids"])
            self.assertNotIn("ent-noise", body["transcript_entry_ids"])
            assert_redacted(json.dumps(eff, ensure_ascii=False))
            self.assertTrue(all(r.get("turn_id") != "req-open" for r in effects))
            # Un List* dans la fenêtre de grâce met à jour le tour (nouvelle ligne).
            late = {
                "ts": "2026-10-24T09:00:06.500+00:00",
                "leg": "host-backend",
                "method": "/aiserver.v1.GrokBotService/ListGrokBotTranscriptEntries",
                "req_body": {"agentId": "agent-bc-42", "limit": 10, "sessionId": "sess-1"},
                "resp_body": {"entries": [{
                    "entryId": "ent-grace",
                    "entryKind": "message",
                    "bodyOmitted": False,
                    "body": {"id": "ent-grace", "kind": "message", "content": "ligne suivante"},
                }], "generation": 5},
            }
            with open(path, "a", encoding="utf-8") as f:
                f.write(json.dumps(late) + "\n")
            self.assertEqual(inference_effects.materialize(tmp), 1)
            with open(path, encoding="utf-8") as fh:
                effects = [json.loads(l) for l in fh if '"inference-effect"' in l]
            self.assertEqual(len(effects), 2)
            self.assertIn("ent-grace", effects[-1]["resp_body"]["transcript_entry_ids"])
            self.assertNotIn("ent-grace", effects[0]["resp_body"]["transcript_entry_ids"])

    def test_prepare_sse_redacts_and_keeps_id(self):
        raw = json.dumps(load_fixture()[2]["resp_body"], ensure_ascii=False)
        out = inference_effects.prepare_sse_data(raw)
        blob = json.dumps(out, ensure_ascii=False)
        assert_redacted(blob)
        self.assertEqual(out["payload"]["entry"]["id"], "ent-9c2")
        self.assertIn("[REDACTED-EMAIL]", out["payload"]["entry"]["content"])


class AssemblerTests(unittest.TestCase):
    def test_split_across_chunks_keeps_both_events(self):
        asm = sse_stream.SseAssembler()
        part1 = b'data: {"channel":"transcript","payload":{"type":"appended","entry":{"id":"ent-a"}}}'
        part2 = (
            b'\n\n:ping\n\n'
            b'data: {"channel":"transcript","payload":{"type":"appended","entry":{"id":"ent-b"}}}\n\n'
        )
        ev1 = asm.feed(part1)
        self.assertEqual(ev1, [])
        ev2 = asm.feed(part2)
        ids = []
        for ev in ev2:
            ids.append(json.loads(ev["data"])["payload"]["entry"]["id"])
        self.assertEqual(ids, ["ent-a", "ent-b"])
        self.assertEqual(asm.pings, 1)
        self.assertEqual(asm.flush(), [])


def _free_port() -> int:
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


def _read_http(reader):
    async def _inner():
        buf = b""
        while b"\r\n\r\n" not in buf:
            chunk = await reader.read(4096)
            if not chunk:
                return None
            buf += chunk
        head, rest = buf.split(b"\r\n\r\n", 1)
        lines = head.decode("latin1").split("\r\n")
        headers = {}
        for line in lines[1:]:
            if ":" in line:
                k, v = line.split(":", 1)
                headers[k.strip().lower()] = v.strip()
        cl = int(headers.get("content-length") or "0")
        body = rest
        while len(body) < cl:
            more = await reader.read(cl - len(body))
            if not more:
                break
            body += more
        return lines[0], headers, body[:cl]
    return _inner()


def _start_fake_host(port: int):
    """Faux host : POST JSON {outcome:accepted}, GET /events en SSE.

    X-Test-Mode: short-cl → Content-Length = premier événement seulement,
    puis un second événement avant EOF (le cas « premier chunk »).
    """
    loop = asyncio.new_event_loop()
    ready = threading.Event()

    async def handle(reader, writer):
        got = await _read_http(reader)
        if not got:
            writer.close()
            return
        req, headers, _body = got
        path = req.split(" ")[1] if " " in req else "/"
        mode = headers.get("x-test-mode", "")
        if path.split("?", 1)[0] == "/events":
            ev_a = (
                'data: {"channel":"transcript","payload":{"type":"appended",'
                '"agentId":"agent-bc-42","entry":{"id":"ent-a","kind":"message",'
                '"content":"un victim@example.com"}}}\n\n'
            ).encode()
            ev_b = (
                'data: {"channel":"transcript","payload":{"type":"appended",'
                '"agentId":"agent-bc-42","entry":{"id":"ent-b","kind":"message",'
                '"content":"deux"}}}\n\n'
            ).encode()
            if mode == "short-cl":
                writer.write(
                    "HTTP/1.1 200 OK\r\nContent-Type: text/event-stream\r\n"
                    f"Content-Length: {len(ev_a)}\r\nConnection: close\r\n\r\n".encode()
                    + ev_a + ev_b
                )
            else:
                writer.write(
                    b"HTTP/1.1 200 OK\r\nContent-Type: text/event-stream\r\n"
                    b"Transfer-Encoding: chunked\r\nConnection: close\r\n\r\n"
                )
                # Coupure au milieu du premier événement, puis ping, puis le second.
                cut = ev_a[:40]
                rest = ev_a[40:] + b":ping\n\n" + ev_b
                for piece in (cut, rest):
                    writer.write(f"{len(piece):x}\r\n".encode() + piece + b"\r\n")
                    await writer.drain()
                    await asyncio.sleep(0.05)
                writer.write(b"0\r\n\r\n")
        else:
            payload = b'{"outcome":"accepted"}'
            writer.write(
                b"HTTP/1.1 200 OK\r\nContent-Type: application/json\r\n"
                + f"Content-Length: {len(payload)}\r\nConnection: close\r\n\r\n".encode()
                + payload
            )
        await writer.drain()
        writer.close()
        try:
            await writer.wait_closed()
        except Exception:
            pass

    async def main():
        srv = await asyncio.start_server(handle, "127.0.0.1", port)
        ready.set()
        async with srv:
            await srv.serve_forever()

    def runner():
        asyncio.set_event_loop(loop)
        try:
            loop.run_until_complete(main())
        except Exception:
            pass

    thread = threading.Thread(target=runner, daemon=True)
    thread.start()
    if not ready.wait(3):
        raise RuntimeError("faux host pas prêt")
    return loop


class RelaySseTests(unittest.TestCase):
    def _run(self, mode: str | None):
        up = _free_port()
        lp = _free_port()
        _start_fake_host(up)
        with tempfile.TemporaryDirectory() as tmp:
            env = os.environ.copy()
            env["RPC_TRAFFIC_LOG_DIR"] = tmp
            proc = subprocess.Popen(
                [sys.executable, RELAY, "--listen", f"127.0.0.1:{lp}",
                 "--upstream", f"127.0.0.1:{up}"],
                env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
            try:
                time.sleep(0.4)
                base = f"http://127.0.0.1:{lp}"
                urllib.request.urlopen(
                    urllib.request.Request(
                        base + "/api/markAgentRequestStarted",
                        data=b'{"requestId":"req-live"}',
                        headers={"Content-Type": "application/json"},
                        method="POST",
                    ),
                    timeout=5,
                ).read()
                headers = {}
                if mode:
                    headers["X-Test-Mode"] = mode
                req = urllib.request.Request(base + "/events?channels=transcript", headers=headers)
                body = urllib.request.urlopen(req, timeout=5).read().decode("utf-8", "replace")
                urllib.request.urlopen(
                    urllib.request.Request(
                        base + "/api/markAgentRequestEnded",
                        data=b'{"requestId":"req-live"}',
                        headers={"Content-Type": "application/json"},
                        method="POST",
                    ),
                    timeout=5,
                ).read()
                time.sleep(0.3)
            finally:
                proc.terminate()
                try:
                    proc.wait(timeout=3)
                except subprocess.TimeoutExpired:
                    proc.kill()
            self.assertIn("ent-a", body)
            self.assertIn("ent-b", body)
            path = proxy.logfile_path(tmp)
            with open(path, encoding="utf-8") as fh:
                lines = [json.loads(l) for l in fh if l.strip()]
            events = [r for r in lines if r.get("kind") == "sse-event"]
            ids = []
            for ev in events:
                payload = (ev.get("resp_body") or {}).get("payload") or {}
                entry = payload.get("entry") or {}
                if entry.get("id"):
                    ids.append(entry["id"])
            self.assertEqual(ids, ["ent-a", "ent-b"])
            blob = json.dumps(lines, ensure_ascii=False)
            self.assertNotIn(EMAIL, blob)
            self.assertIn("[REDACTED-EMAIL]", blob)
            streams = [r for r in lines if r.get("kind") == "sse-stream"]
            self.assertEqual(len(streams), 1)
            self.assertEqual(streams[0]["sse_events"], 2)
            effects = [r for r in lines if r.get("leg") == "inference-effect"]
            self.assertEqual(len(effects), 1)
            turn = effects[0]["resp_body"]
            self.assertEqual(turn["status"], "completed")
            self.assertEqual(turn["request_id"], "req-live")
            self.assertEqual(turn["agent_id"], "agent-bc-42")
            self.assertEqual(turn["transcript_entry_ids"], ["ent-a", "ent-b"])
            self.assertTrue(turn["duree_ms"] is not None and turn["duree_ms"] >= 0)
            assert_redacted(json.dumps(effects[0], ensure_ascii=False))
            if mode != "short-cl":
                self.assertGreaterEqual(streams[0].get("sse_pings") or 0, 1)

    def test_chunked_sse_logs_every_event_and_turn(self):
        self._run(None)

    def test_short_content_length_does_not_drop_the_rest(self):
        self._run("short-cl")


class DashboardTests(unittest.TestCase):
    def test_stats_include_turn(self):
        spec = importlib.util.spec_from_file_location("rpc_dash_server", DASH)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        with tempfile.TemporaryDirectory() as tmp:
            path = proxy.logfile_path(tmp)
            with open(path, "w", encoding="utf-8") as f:
                for rec in load_fixture():
                    f.write(json.dumps(rec, ensure_ascii=False) + "\n")
            inference_effects.materialize(tmp)
            mod.LOG_DIR = tmp
            mod._records.clear()
            mod._offsets.clear()
            mod.refresh(1)
            stats = mod.compute_stats(1, 50)
        turns = [t for t in stats["inference_turns"] if t["request_id"] == "req-7f3a"]
        self.assertEqual(len(turns), 1)
        self.assertEqual(turns[0]["status"], "completed")
        self.assertEqual(turns[0]["agent_id"], "agent-bc-42")
        self.assertIn("ent-9c2", turns[0]["transcript_entry_ids"])
        self.assertIn("ent-omit", turns[0]["transcript_entry_ids"])
        assert_redacted(json.dumps(turns[0], ensure_ascii=False))
        self.assertIn("StreamChat", stats["inference_note"])
        legs = {r.get("leg") for r in stats["last"]}
        self.assertIn("inference-effect", legs)


if __name__ == "__main__":
    unittest.main()
