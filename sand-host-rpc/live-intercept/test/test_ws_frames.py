#!/usr/bin/env python3
"""WebSocket : trames synthétiques, relai Upgrade/101, et pcap CDP loopback.

(a) texte JSON rédigé, binaire hex ou proto catalogue, contrôle résumé,
    fragments recollés, RSV non décompressé ;
(b) le relai agent↔host ne droppe pas un Upgrade et journalise les payloads ;
(c) un pcap loopback passé dans decode_keylog --plaintext sort des lignes
    protocol=websocket lisibles (tshark 4.2 : payload.text + data.data).
"""
from __future__ import annotations

import asyncio
import json
import os
import shutil
import socket
import subprocess
import sys
import tempfile
import time
import unittest

_LI = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_MON = os.path.join(os.path.dirname(_LI), "rpc-traffic-monitor")
for _p in (_MON, os.path.join(_MON, "mitmproxy"), _LI):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import decode_keylog as dk  # noqa: E402
import proto_decode  # noqa: E402
import proxy  # noqa: E402
import relay_agent_host  # noqa: E402
import ws_frames  # noqa: E402

PAGE = "/devtools/page/TARGET123"


def _hx(b: bytes) -> str:
    return ":".join(f"{x:02x}" for x in b)


class ParserTests(unittest.TestCase):
    def test_masked_text_json_redacted_and_fragment_reassembly(self):
        body = json.dumps({
            "id": 1,
            "method": "Page.navigate",
            "params": {"url": "https://example.com", "contact": "user@corp.com",
                       "token": "sekret-token-value"},
        }).encode()
        raw = ws_frames.encode_frame(1, body[:8], mask=True, fin=False, mask_key=b"\x11\x22\x33\x44")
        raw += ws_frames.encode_frame(0, body[8:], mask=True, fin=True, mask_key=b"\x11\x22\x33\x44")
        # coupé au milieu du second en-tête, puis complété
        frames, rest = ws_frames.parse_frames(raw[:3])
        self.assertEqual(frames, [])
        self.assertTrue(rest)
        frames, rest = ws_frames.parse_frames(raw)
        self.assertEqual(rest, b"")
        reasm = ws_frames.Reassembler()
        done = []
        for fr in frames:
            done.extend(reasm.push(fr))
        self.assertEqual(len(done), 1)
        name, payload = ws_frames.describe(done[0], PAGE, "request")
        self.assertEqual(name, "text")
        self.assertEqual(payload["method"], "Page.navigate")
        self.assertEqual(payload["params"]["contact"], "[REDACTED-EMAIL]")
        self.assertEqual(payload["params"]["token"], "[REDACTED]")
        self.assertNotIn("user@corp.com", json.dumps(payload))

    def test_binary_hex_control_and_rsv_not_inflated(self):
        blob = b"\x01\x02\x03\x04"
        frames, rest = ws_frames.parse_frames(
            ws_frames.encode_frame(2, blob, mask=False)
            + ws_frames.encode_frame(9, b"", mask=False)
            + ws_frames.encode_frame(8, (1000).to_bytes(2, "big") + b"bye user@corp.com", mask=False)
            + ws_frames.encode_frame(1, b'{"a":1}', mask=False, rsv=1)
        )
        self.assertEqual(rest, b"")
        reasm = ws_frames.Reassembler()
        msgs = []
        for fr in frames:
            msgs.extend(reasm.push(fr))
        self.assertEqual(ws_frames.describe(msgs[0], "/nope", "response")[1],
                         {"_binary_hex": "01020304", "_binary_len": 4})
        self.assertEqual(ws_frames.describe(msgs[1], "/", None)[1], {"opcode": "ping", "len": 0})
        close = ws_frames.describe(msgs[2], "/", None)[1]
        self.assertEqual(close["code"], 1000)
        self.assertIn("[REDACTED-EMAIL]", close["reason"])
        name, comp = ws_frames.describe(msgs[3], "/", None)
        self.assertEqual(name, "text")
        self.assertEqual(comp["_ws"], "compressed")
        self.assertNotIn("a", comp)

    def test_binary_proto_when_catalog_maps(self):
        proto_decode._CACHE = {
            "messages": {"t.M": [{"no": 1, "name": "email", "type": "string"}]},
            "enums": {},
            "methods": {"/t.S/M": {"request": "t.M", "response": "t.M", "kind": "Unary"}},
        }
        try:
            # field 1, wire 2, string user@corp.com
            payload = b"\x0a" + bytes([len(b"user@corp.com")]) + b"user@corp.com"
            msg = ws_frames.Logical(2, payload, True, True, False)
            _name, out = ws_frames.describe(msg, "/t.S/M", "request")
            self.assertEqual(out["email"], "[REDACTED-EMAIL]")
        finally:
            proto_decode.reset_catalog_cache()

    def test_proc_listeners_and_excluded_ports(self):
        sample = (
            "  sl  local_address rem_address   st\n"
            "   0: 0100007F:2411 00000000:0000 0A\n"  # 127.0.0.1:9233
            "   1: 0100007F:17C0 00000000:0000 0A\n"  # 6080
            "   2: 010011AC:0050 00000000:0000 0A\n"  # pas loopback
        )
        ports = dk.parse_proc_net_listeners(sample)
        self.assertIn(0x2411, ports)
        self.assertIn(6080, ports)
        self.assertNotIn(80, ports)
        self.assertEqual(dk.discover_cdp_ports(["9229", "6080", "6081"]), [9229])


def _layers(stream, sport, dport, **extra):
    layers = {
        "ip": {"ip.src": "127.0.0.1", "ip.dst": "127.0.0.1"},
        "tcp": {
            "tcp.srcport": str(sport),
            "tcp.dstport": str(dport),
            "tcp.stream": str(stream),
        },
    }
    layers.update(extra)
    return layers


class AssemblerTests(unittest.TestCase):
    def setUp(self):
        proto_decode.reset_catalog_cache()

    def test_cdp_text_binary_close_from_tshark_shape(self):
        asm = dk.Assembler(dk.re.compile(dk.DEFAULT_ALLOW), "198.18.0.1",
                           mode="cdp", cdp_ports={"9333"})
        path = PAGE
        asm.note_packet(_layers(7, 40000, 9333, http={
            "http.request.method": "GET",
            "http.request.uri": path,
            "http.host": "127.0.0.1:9333",
            "http.upgrade": "websocket",
            "http.request.line": ["Host: 127.0.0.1:9333", "Upgrade: websocket", "Connection: Upgrade"],
        }))
        asm.note_packet(_layers(7, 9333, 40000, http={
            "http.response.code": "101",
            "http.upgrade": "websocket",
            "http.response.line": ["Upgrade: websocket", "Connection: Upgrade",
                                   "Sec-WebSocket-Accept: abc"],
        }))
        text = json.dumps({"id": 1, "method": "Page.navigate",
                           "params": {"contact": "user@corp.com"}})
        asm.note_packet(_layers(7, 40000, 9333, websocket={
            "websocket.fin": "1",
            "websocket.rsv": "0x00",
            "websocket.opcode": "1",
            "websocket.mask": "1",
            "websocket.payload_length": str(len(text)),
            "websocket.payload": {"websocket.payload.text": text},
        }))
        # opcode 2 : tshark 4.2 laisse websocket.payload vide et met les octets dans data.data
        asm.note_packet(_layers(7, 9333, 40000, websocket={
            "websocket.fin": "1", "websocket.rsv": "0x00", "websocket.opcode": "2",
            "websocket.mask": "0", "websocket.payload_length": "4", "websocket.payload": "",
        }, data={"data.data": "01:02:03:04", "data.len": "4"}))
        # fragment vide + continuation qui porte le message recollé, plus un ping dans le même paquet
        asm.note_packet(_layers(7, 9333, 40000, websocket=[
            {"websocket.fin": "1", "websocket.rsv": "0x00", "websocket.opcode": "9",
             "websocket.mask": "0", "websocket.payload_length": "0",
             "websocket.payload": {"websocket.payload.ping": ""}},
            {"websocket.fin": "0", "websocket.rsv": "0x00", "websocket.opcode": "1",
             "websocket.mask": "0", "websocket.payload_length": "4", "websocket.payload": ""},
            {"websocket.fin": "1", "websocket.rsv": "0x00", "websocket.opcode": "0",
             "websocket.mask": "0", "websocket.payload_length": "20",
             "websocket.payload": {"websocket.payload.text": '{"id":2,"method":"Runtime.evaluate"}'}},
        ]))
        recs = asm.records()
        kinds = [(r.get("ws_event"), r.get("ws_opcode"), r.get("method"), r.get("direction")) for r in recs]
        self.assertIn(("handshake", None, path, "host->cdp"), kinds)
        nav = [r for r in recs if r.get("method") == "Page.navigate"]
        self.assertEqual(len(nav), 1)
        self.assertEqual(nav[0]["protocol"], "websocket")
        self.assertEqual(nav[0]["leg"], "host-cdp")
        self.assertEqual(nav[0]["req_body"]["params"]["contact"], "[REDACTED-EMAIL]")
        self.assertIsNone(nav[0]["resp_body"])
        binary = [r for r in recs if r.get("ws_opcode") == "binary"]
        self.assertEqual(binary[0]["resp_body"]["_binary_hex"], "01020304")
        self.assertEqual(binary[0]["direction"], "cdp->host")
        self.assertTrue(any(r.get("ws_opcode") == "ping" for r in recs))
        ev = [r for r in recs if r.get("method") == "Runtime.evaluate"]
        self.assertEqual(len(ev), 1)
        # le fragment FIN=0 ne produit pas une ligne vide en plus
        self.assertEqual(sum(1 for r in recs if r.get("ws_event") == "frame"), 4)

    def test_novnc_and_tunnel_dropped(self):
        asm = dk.Assembler(dk.re.compile(dk.DEFAULT_ALLOW), "198.18.0.1",
                           mode="cdp", cdp_ports={"6080", "9333"})
        asm.note_packet(_layers(1, 40000, 6080, http={
            "http.request.method": "GET",
            "http.request.uri": "/websockify",
            "http.upgrade": "websocket",
            "http.request.line": ["Upgrade: websocket"],
        }))
        asm.note_packet(_layers(1, 6080, 40000, http={
            "http.response.code": "101",
            "http.upgrade": "websocket",
            "http.response.line": ["Upgrade: websocket"],
        }))
        self.assertEqual(asm.records(), [])
        asm2 = dk.Assembler(dk.re.compile(dk.DEFAULT_ALLOW), "198.18.0.1", mode="cdp", cdp_ports={"9333"})
        pkt = _layers(2, 40000, 9333, http={
            "http.request.method": "GET", "http.request.uri": PAGE,
            "http.upgrade": "websocket", "http.request.line": ["Upgrade: websocket"],
        })
        pkt["ip"] = {"ip.src": "198.18.0.1", "ip.dst": "127.0.0.1"}
        asm2.note_packet(pkt)
        self.assertEqual(asm2.records(), [])

    def test_h2_extended_connect_parses_raw_frames(self):
        asm = dk.Assembler(dk.re.compile(dk.DEFAULT_ALLOW), "198.18.0.1")
        body = json.dumps({"hello": "user@corp.com"}).encode()
        frame = ws_frames.encode_frame(1, body, mask=True, mask_key=b"\x01\x02\x03\x04")
        c, s = "10.0.0.8", "10.1.0.9"
        asm.note_packet({
            "ip": {"ip.src": c, "ip.dst": s},
            "tcp": {"tcp.srcport": "40000", "tcp.dstport": "443", "tcp.stream": "3"},
            "tls": {"tls.handshake.extensions_server_name": "api2.cursor.sh"},
            "http2": {"http2.stream": {
                "http2.streamid": "1",
                "http2.type": "1",
                "http2.flags": "0x04",
                "http2.flags_tree": {"http2.flags.end_stream": "0"},
                "http2.header": [
                    {"http2.header.name": n, "http2.header.value": v}
                    for n, v in (
                        (":method", "CONNECT"),
                        (":protocol", "websocket"),
                        (":path", "/agent.v1.AgentService/Run"),
                        (":authority", "api2.cursor.sh"),
                        (":scheme", "https"),
                    )
                ],
            }},
        })
        asm.note_packet({
            "ip": {"ip.src": s, "ip.dst": c},
            "tcp": {"tcp.srcport": "443", "tcp.dstport": "40000", "tcp.stream": "3"},
            "http2": {"http2.stream": {
                "http2.streamid": "1",
                "http2.type": "1",
                "http2.flags": "0x04",
                "http2.flags_tree": {"http2.flags.end_stream": "0"},
                "http2.header": [
                    {"http2.header.name": ":status", "http2.header.value": "200"},
                ],
            }},
        })
        asm.note_packet({
            "ip": {"ip.src": c, "ip.dst": s},
            "tcp": {"tcp.srcport": "40000", "tcp.dstport": "443", "tcp.stream": "3"},
            "http2": {"http2.stream": {
                "http2.streamid": "1",
                "http2.type": "0",
                "http2.length": str(len(frame)),
                "http2.flags": "0x00",
                "http2.flags_tree": {"http2.flags.end_stream": "0"},
                "http2.data.data": _hx(frame),
            }},
        })
        recs = [r for r in asm.records() if r.get("protocol") == "websocket"]
        self.assertTrue(any(r.get("ws_event") == "handshake" and r.get("status") == "200" for r in recs))
        frames = [r for r in recs if r.get("ws_event") == "frame"]
        self.assertEqual(len(frames), 1)
        self.assertEqual(frames[0]["leg"], "host-backend")
        self.assertEqual(frames[0]["direction"], "host->backend")
        self.assertEqual(frames[0]["req_body"]["hello"], "[REDACTED-EMAIL]")
        # pas une ligne unaire en plus, avec un corps proto vide
        self.assertTrue(all(r.get("protocol") == "websocket" for r in asm.records()))


class RelayTests(unittest.TestCase):
    def test_upgrade_frames_logged_not_dropped(self):
        log_dir = tempfile.mkdtemp(prefix="wsrelay-")
        os.environ["RPC_TRAFFIC_LOG_DIR"] = log_dir

        async def run():
            got = {}

            async def upstream(reader, writer):
                got_h = await proxy.read_headers(reader)
                self.assertIsNotNone(got_h)
                start, headers = got_h
                self.assertIn("websocket", dict((k.lower(), v) for k, v in headers).get("upgrade", ""))
                body = json.dumps({
                    "id": 1, "result": {"contact": "user@corp.com", "ok": True},
                }).encode()
                writer.write(
                    b"HTTP/1.1 101 Switching Protocols\r\n"
                    b"Upgrade: websocket\r\nConnection: Upgrade\r\n"
                    b"Sec-WebSocket-Accept: dGhlIHNhbXBsZSBub25jZQ==\r\n\r\n"
                )
                await writer.drain()
                buf = b""
                while True:
                    chunk = await reader.read(65536)
                    if not chunk:
                        break
                    buf += chunk
                    frames, buf = ws_frames.parse_frames(buf)
                    if frames:
                        got["client"] = frames[0].payload
                        break
                writer.write(ws_frames.encode_frame(1, body, mask=False))
                writer.write(ws_frames.encode_frame(2, b"\x01\x02\x03\x04", mask=False))
                writer.write(ws_frames.encode_frame(8, (1000).to_bytes(2, "big") + b"done", mask=False))
                await writer.drain()
                writer.close()

            server = await asyncio.start_server(upstream, "127.0.0.1", 0)
            sport = server.sockets[0].getsockname()[1]
            relay = relay_agent_host.Relay("127.0.0.1", sport)
            front = await asyncio.start_server(relay.handle, "127.0.0.1", 0)
            fport = front.sockets[0].getsockname()[1]

            reader, writer = await asyncio.open_connection("127.0.0.1", fport)
            req = (
                f"GET {PAGE} HTTP/1.1\r\n"
                f"Host: 127.0.0.1:{fport}\r\n"
                "Upgrade: websocket\r\n"
                "Connection: Upgrade\r\n"
                "Sec-WebSocket-Key: dGhlIHNhbXBsZSBub25jZQ==\r\n"
                "Sec-WebSocket-Version: 13\r\n\r\n"
            )
            writer.write(req.encode())
            await writer.drain()
            # 101
            blob = b""
            while b"\r\n\r\n" not in blob:
                blob += await reader.read(4096)
            self.assertIn(b"101", blob.split(b"\r\n", 1)[0])
            payload = json.dumps({
                "id": 1, "method": "Page.navigate",
                "params": {"url": "https://example.com", "token": "sekret",
                           "authorization": "Bearer ghp_FAKEAAAAAAAAAAAAAAAAAAAAAAAA"},
            }).encode()
            writer.write(ws_frames.encode_frame(1, payload, mask=True, mask_key=b"\x0a\x0b\x0c\x0d"))
            await writer.drain()
            data = b""
            try:
                while True:
                    chunk = await asyncio.wait_for(reader.read(65536), 2)
                    if not chunk:
                        break
                    data += chunk
            except asyncio.TimeoutError:
                pass
            writer.close()
            server.close()
            front.close()
            await server.wait_closed()
            await front.wait_closed()
            self.assertIn(b"Page.navigate", got.get("client", b""))
            self.assertTrue(data)

        asyncio.run(run())
        logs = [os.path.join(log_dir, n) for n in os.listdir(log_dir) if n.endswith(".jsonl")]
        self.assertTrue(logs)
        rows = []
        for path in logs:
            with open(path, encoding="utf-8") as fh:
                rows.extend(json.loads(line) for line in fh if line.strip())
        blob = json.dumps(rows)
        self.assertNotIn("user@corp.com", blob)
        self.assertNotIn("ghp_FAKE", blob)
        self.assertNotIn("sekret", blob)
        self.assertTrue(any(r.get("ws_event") == "handshake" and r.get("status") == "101" for r in rows))
        nav = [r for r in rows if r.get("method") == "Page.navigate"]
        self.assertEqual(len(nav), 1)
        self.assertEqual(nav[0]["leg"], "agent-host")
        self.assertEqual(nav[0]["protocol"], "websocket")
        self.assertEqual(nav[0]["req_body"]["params"]["token"], "[REDACTED]")
        self.assertEqual(nav[0]["req_body"]["params"]["authorization"], "[REDACTED]")
        binary = [r for r in rows if r.get("ws_opcode") == "binary"]
        self.assertEqual(binary[0]["resp_body"]["_binary_hex"], "01020304")
        self.assertEqual(binary[0]["direction"], "host->agent")
        self.assertTrue(any(r.get("ws_opcode") == "close" for r in rows))
        shutil.rmtree(log_dir, ignore_errors=True)


def _have_capture_tools() -> bool:
    if not shutil.which("tshark") or not shutil.which("tcpdump"):
        return False
    try:
        subprocess.run(["sudo", "-n", "true"], check=True, capture_output=True)
    except (subprocess.CalledProcessError, FileNotFoundError):
        return False
    return True


@unittest.skipUnless(_have_capture_tools(), "tshark/tcpdump/sudo indisponibles")
class LiveCdpPcapTests(unittest.TestCase):
    def test_loopback_cdp_frames_in_jsonl(self):
        run = tempfile.mkdtemp(prefix="cdpws-")
        probe = socket.socket()
        probe.bind(("127.0.0.1", 0))
        port = probe.getsockname()[1]
        probe.close()
        td = None
        try:
            pcap = os.path.join(run, "cap.pcap")
            td = subprocess.Popen(
                ["sudo", "-n", "tcpdump", "--immediate-mode", "-i", "lo", "-n",
                 "-s", "0", "-U", "-w", pcap, "tcp", "port", str(port)],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
            deadline = time.time() + 3
            while time.time() < deadline and (not os.path.isfile(pcap) or os.path.getsize(pcap) < 24):
                if td.poll() is not None:
                    self.fail("tcpdump mort")
                time.sleep(0.05)
            time.sleep(0.3)

            srv = socket.socket()
            srv.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            srv.bind(("127.0.0.1", port))
            srv.listen(1)
            srv.settimeout(3)

            def serve():
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
                c.settimeout(2)
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
                c.sendall(ws_frames.encode_frame(2, b"\x0a\x0b\x0c\x0d", mask=False))
                c.close()

            import threading
            th = threading.Thread(target=serve)
            th.start()
            cli = socket.create_connection(("127.0.0.1", port), 2)
            key_payload = json.dumps({
                "id": 1, "method": "Page.navigate",
                "params": {"url": "https://example.com", "contact": "user@corp.com"},
            }).encode()
            cli.sendall(
                f"GET {PAGE} HTTP/1.1\r\nHost: 127.0.0.1:{port}\r\n"
                "Upgrade: websocket\r\nConnection: Upgrade\r\n"
                "Sec-WebSocket-Key: dGhlIHNhbXBsZSBub25jZQ==\r\n"
                "Sec-WebSocket-Version: 13\r\n\r\n".encode()
            )
            acc = b""
            while b"\r\n\r\n" not in acc:
                acc += cli.recv(4096)
            cli.sendall(ws_frames.encode_frame(1, key_payload, mask=True, mask_key=b"\x01\x02\x03\x04"))
            cli.settimeout(2)
            try:
                while cli.recv(65536):
                    pass
            except socket.timeout:
                pass
            cli.close()
            th.join(3)
            srv.close()
            time.sleep(0.3)
            subprocess.call(["sudo", "-n", "kill", "-INT", str(td.pid)],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            try:
                td.wait(timeout=3)
            except subprocess.TimeoutExpired:
                td.kill()
            td = None
            self.assertGreater(os.path.getsize(pcap), 24)
            os.environ["RPC_TRAFFIC_LOG_DIR"] = os.path.join(run, "logs")
            n = dk.decode_once(
                pcap, "", os.path.join(run, "seen"), [],
                dk.re.compile(dk.DEFAULT_ALLOW), "198.18.0.1", shutil.which("tshark"),
                plaintext=True, cdp_ports=[str(port)],
            )
            self.assertGreaterEqual(n, 1, "aucune ligne CDP")
            logdir = os.path.join(run, "logs")
            rows = []
            for name in os.listdir(logdir):
                with open(os.path.join(logdir, name), encoding="utf-8") as fh:
                    rows.extend(json.loads(line) for line in fh if line.strip())
            blob = json.dumps(rows)
            self.assertNotIn("user@corp.com", blob)
            nav = [r for r in rows if r.get("method") == "Page.navigate"]
            self.assertTrue(nav, rows)
            self.assertEqual(nav[0]["protocol"], "websocket")
            self.assertEqual(nav[0]["leg"], "host-cdp")
            self.assertEqual(nav[0]["req_body"]["params"]["contact"], "[REDACTED-EMAIL]")
            self.assertIn("/devtools/page/", nav[0]["ws_path"])
            binary = [r for r in rows if r.get("ws_opcode") == "binary"]
            self.assertTrue(binary, rows)
            self.assertEqual(binary[0]["resp_body"]["_binary_hex"], "0a0b0c0d")
            self.assertTrue(any(r.get("ws_event") == "handshake" and r.get("status") == "101" for r in rows))
        finally:
            if td is not None and td.poll() is None:
                subprocess.call(["sudo", "-n", "kill", "-INT", str(td.pid)],
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            shutil.rmtree(run, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
