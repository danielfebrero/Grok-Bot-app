#!/usr/bin/env python3
"""Corps HTTP/2 réassemblés + protobuf Connect nommé depuis le catalogue.

(a) une réponse application/proto coupée en plusieurs DATA n'est pas journalisée
    tant qu'elle n'est pas finie, et le corps final n'est pas la concaténation
    fragments + corps réassemblé ;
(b) un petit payload proto connu sort avec les noms de champs du catalogue.
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import unittest

_LI = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_MON = os.path.join(os.path.dirname(_LI), "rpc-traffic-monitor")
for _p in (_MON, os.path.join(_MON, "mitmproxy"), _LI):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import decode_keylog as dk  # noqa: E402
import proto_decode  # noqa: E402
import proxy  # noqa: E402

PATH = "/aiserver.v1.DashboardService/GetManagedSkills"


def _varint(n: int) -> bytes:
    out = bytearray()
    while True:
        b = n & 0x7F
        n >>= 7
        if n:
            out.append(b | 0x80)
        else:
            out.append(b)
            break
    return bytes(out)


def _ld(field: int, payload: bytes) -> bytes:
    return _varint((field << 3) | 2) + _varint(len(payload)) + payload


def _hx(b: bytes) -> str:
    return ":".join(f"{x:02x}" for x in b)


def _skill_response(content: bytes) -> bytes:
    skill = _ld(1, b"id1") + _ld(3, content)
    return _ld(1, skill)


class ProtoDecodeTests(unittest.TestCase):
    def tearDown(self):
        proto_decode.reset_catalog_cache()

    def test_named_fields_from_catalog(self):
        raw = _ld(1, b"kit") + _ld(2, b"bc")
        out = proxy.decode_body(raw, "application/proto", "", PATH, "request")
        self.assertEqual(out["supported_canvas_kits"], ["kit"])
        self.assertEqual(out["bc_id"], "bc")
        self.assertNotIn("_unknown_1", out)

    def test_nested_response_and_unknown_tag_and_redaction(self):
        skill = _ld(1, b"id1") + _ld(3, b"user@corp.com")
        # tag 15, fil 0, valeur 7 — pas dans GetManagedSkillsResponse
        raw = _ld(1, skill) + bytes([0x78, 0x07])
        out = proxy.decode_body(raw, "application/proto", "", PATH, "response")
        self.assertEqual(out["skills"][0]["id"], "id1")
        self.assertEqual(out["skills"][0]["content"], "[REDACTED-EMAIL]")
        self.assertEqual(out["_unknown_15"], 7)

    def test_json_agent_host_path_unchanged(self):
        raw = b'{"email":"user@corp.com","ok":true}'
        out = proxy.decode_body(raw, "application/json", "")
        self.assertEqual(out, {"email": "[REDACTED-EMAIL]", "ok": True})

    def test_connect_envelope_and_packed_map_enum(self):
        proto_decode._CACHE = {
            "messages": {
                "t.M": [
                    {"no": 1, "name": "xs", "type": "int32", "repeated": True},
                    {"no": 2, "name": "tags", "type": "map", "key": "string", "value": "int32"},
                    {"no": 3, "name": "kind", "type": "ref#t.K"},
                ]
            },
            "enums": {"t.K": {"0": "K_UNSPECIFIED", "1": "K_A"}},
            "methods": {"/t.S/M": {"request": "t.M", "response": "t.M", "kind": "Unary"}},
        }
        packed = bytes([0x01, 0x02, 0xAC, 0x02])  # 1, 2, 300
        entry = _ld(1, b"a") + bytes([0x10, 0x05])
        body = _ld(1, packed) + _ld(2, entry) + bytes([0x18, 0x01])
        frame = bytes([0]) + len(body).to_bytes(4, "big") + body
        out = proxy.decode_body(frame, "application/connect+proto", "", "/t.S/M", "response")
        self.assertEqual(out["xs"], [1, 2, 300])
        self.assertEqual(out["tags"], {"a": 5})
        self.assertEqual(out["kind"], "K_A")


def _layers(src, sport, dst, dport, stream, streams, sni=None):
    layers = {
        "ip": {"ip.src": src, "ip.dst": dst},
        "tcp": {
            "tcp.srcport": str(sport),
            "tcp.dstport": str(dport),
            "tcp.stream": str(stream),
        },
        "http2": [{"http2.stream": s} for s in streams],
    }
    if sni:
        layers["tls"] = {"tls.handshake.extensions_server_name": sni}
    return layers


def _hdr(pairs, typ_flags="0x04", end="0"):
    return {
        "http2.streamid": "1",
        "http2.type": "1",
        "http2.flags": typ_flags,
        "http2.flags_tree": {"http2.flags.end_stream": end},
        "http2.header": [
            {"http2.header.name": n, "http2.header.value": v} for n, v in pairs
        ],
    }


def _data(payload: bytes, h2len: int | None = None, end: str = "0", full: bytes | None = None):
    shown = payload if full is None else full
    st = {
        "http2.streamid": "1",
        "http2.type": "0",
        "http2.length": str(len(payload) if h2len is None else h2len),
        "http2.flags": "0x01" if end == "1" else "0x00",
        "http2.flags_tree": {"http2.flags.end_stream": end},
        "http2.data.data": _hx(shown),
    }
    if full is not None:
        st["http2.body.fragments"] = {"http2.body.reassembled.data": _hx(full)}
    return st


class H2BodyTests(unittest.TestCase):
    def setUp(self):
        proto_decode.reset_catalog_cache()
        self.asm = dk.Assembler(dk.re.compile(dk.DEFAULT_ALLOW), "198.18.0.1")
        self.body = _skill_response(b"C" * 40)
        self.req = _ld(2, b"bc")

    def _feed(self, *frames):
        for layers in frames:
            self.asm.note_packet(layers)

    def _req_frames(self):
        c, s = "10.0.0.8", "10.1.0.9"
        return [
            _layers(c, 40000, s, 443, 0, [
                _hdr([
                    (":method", "POST"),
                    (":path", PATH),
                    (":authority", "api2.cursor.sh"),
                    ("content-type", "application/proto"),
                    ("content-length", str(len(self.req))),
                ], end="0"),
                _data(self.req, end="1"),
            ], sni="api2.cursor.sh"),
        ]

    def _resp_headers(self):
        return _layers("10.1.0.9", 443, "10.0.0.8", 40000, 0, [
            _hdr([
                (":status", "200"),
                ("content-type", "application/proto"),
                ("content-length", str(len(self.body))),
            ], end="0"),
        ])

    def test_incomplete_response_not_logged_and_reassembly_not_doubled(self):
        self._feed(*self._req_frames(), self._resp_headers())
        self.assertEqual(self.asm.records(), [])
        frag = self.body[:10]
        self._feed(_layers("10.1.0.9", 443, "10.0.0.8", 40000, 0, [
            _data(frag, end="0"),
        ]))
        self.assertEqual(self.asm.records(), [])
        ex = self.asm.h2[("0", "1")]
        self.assertEqual(len(ex.resp_body), 10)
        # dernière trame : http2.length = reste, http2.data.data = corps COMPLET
        # (ce que tshark 4.2 pose en plus des fragments)
        rest = self.body[10:]
        self._feed(_layers("10.1.0.9", 443, "10.0.0.8", 40000, 0, [
            _data(rest, h2len=len(rest), end="1", full=self.body),
        ]))
        self.assertEqual(len(ex.resp_body), len(self.body))
        self.assertEqual(ex.resp_body, self.body)
        recs = self.asm.records()
        self.assertEqual(len(recs), 1)
        body = recs[0]["resp_body"]
        self.assertIsInstance(body, dict)
        self.assertEqual(body["skills"][0]["id"], "id1")
        self.assertEqual(body["skills"][0]["content"], "C" * 40)
        self.assertEqual(recs[0]["req_body"]["bc_id"], "bc")

    def test_fragments_only_sum_to_content_length(self):
        self._feed(*self._req_frames(), self._resp_headers())
        a, b = self.body[:15], self.body[15:]
        self._feed(
            _layers("10.1.0.9", 443, "10.0.0.8", 40000, 0, [_data(a, end="0")]),
            _layers("10.1.0.9", 443, "10.0.0.8", 40000, 0, [_data(b, end="1")]),
        )
        ex = self.asm.h2[("0", "1")]
        self.assertEqual(ex.resp_body, self.body)
        self.assertEqual(len(self.asm.records()), 1)


def _have_capture_tools() -> bool:
    if not shutil.which("tshark") or not shutil.which("tcpdump"):
        return False
    if not os.path.isfile("/exec-daemon/node") and not shutil.which("node"):
        return False
    try:
        subprocess.run(["sudo", "-n", "true"], check=True, capture_output=True)
    except (subprocess.CalledProcessError, FileNotFoundError):
        return False
    return True


@unittest.skipUnless(_have_capture_tools(), "tshark/tcpdump/node/sudo indisponibles")
class LivePcapTests(unittest.TestCase):
    def test_pcap_proto_body_survives_tshark(self):
        """Échange unaire réel : pcap + keylog → corps proto nommé, pas null.

        tcpdump doit être en --immediate-mode. Sans ça, TPACKET_V3 ne livre
        les paquets qu'au timeout (~1 s) : un arrêt dans la seconde laisse un
        pcap d'en-tête seul (0 captured, N received) et decode_once renvoie 0.
        """
        import json
        import socket
        import tempfile
        import time
        node = "/exec-daemon/node" if os.path.isfile("/exec-daemon/node") else shutil.which("node")
        if not node or not shutil.which("tshark") or not shutil.which("tcpdump") or not shutil.which("openssl"):
            self.skipTest("node, tshark, tcpdump ou openssl manquant")
        if subprocess.call(["sudo", "-n", "true"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL) != 0:
            self.skipTest("sudo -n indisponible")
        run = tempfile.mkdtemp(prefix="h2proto-")
        probe = socket.socket()
        probe.bind(("127.0.0.1", 0))
        port = probe.getsockname()[1]
        probe.close()
        content = b"D" * 17000
        body = _skill_response(content)
        req = _ld(2, b"bc")
        td = sp = None
        try:
            subprocess.check_call([
                "openssl", "req", "-x509", "-newkey", "rsa:2048", "-nodes",
                "-keyout", os.path.join(run, "s.key"),
                "-out", os.path.join(run, "s.crt"),
                "-subj", "/CN=api2.cursor.sh", "-days", "1",
            ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            with open(os.path.join(run, "body.bin"), "wb") as fh:
                fh.write(body)
            srv = r"""
const fs = require("fs");
const http2 = require("http2");
const body = fs.readFileSync(process.argv[2]);
const srv = http2.createSecureServer({
  key: fs.readFileSync(process.argv[3]),
  cert: fs.readFileSync(process.argv[4]),
});
srv.on("stream", (stream) => {
  stream.on("data", () => {});
  stream.on("end", () => {
    stream.respond({
      ":status": 200,
      "content-type": "application/proto",
      "content-length": body.length,
    });
    stream.end(body);
  });
});
srv.listen(Number(process.argv[5]), "127.0.0.1");
"""
            with open(os.path.join(run, "h2.js"), "w") as fh:
                fh.write(srv)
            sp = subprocess.Popen(
                [node, os.path.join(run, "h2.js"), os.path.join(run, "body.bin"),
                 os.path.join(run, "s.key"), os.path.join(run, "s.crt"), str(port)],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
            deadline = time.time() + 5
            while time.time() < deadline:
                try:
                    s = socket.create_connection(("127.0.0.1", port), 0.2)
                    s.close()
                    break
                except OSError:
                    if sp.poll() is not None:
                        self.fail("serveur HTTP/2 mort au démarrage")
                    time.sleep(0.05)
            else:
                self.fail("serveur HTTP/2 pas à l'écoute")
            pcap = os.path.join(run, "cap.pcap")
            keylog = os.path.join(run, "keys.log")
            with open(keylog, "w"):
                pass
            os.chmod(keylog, 0o600)
            td = subprocess.Popen(
                ["sudo", "-n", "tcpdump", "--immediate-mode", "-i", "lo", "-n",
                 "-s", "0", "-U", "-w", pcap, f"tcp port {port}"],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
            deadline = time.time() + 3
            while time.time() < deadline and (not os.path.isfile(pcap) or os.path.getsize(pcap) < 24):
                if td.poll() is not None:
                    self.fail("tcpdump mort au démarrage")
                time.sleep(0.05)
            # le « listening » arrive après la création de l'en-tête pcap
            time.sleep(0.3)
            with open(os.path.join(run, "req.bin"), "wb") as fh:
                fh.write(req)
            client = r"""
const fs = require("fs");
const http2 = require("http2");
const req = fs.readFileSync(process.env.REQ);
const c = http2.connect("https://127.0.0.1:" + process.env.PORT, {
  rejectUnauthorized: false,
  servername: "api2.cursor.sh",
});
c.on("error", () => process.exit(1));
const r = c.request({
  ":method": "POST",
  ":path": "/aiserver.v1.DashboardService/GetManagedSkills",
  ":authority": "api2.cursor.sh",
  "content-type": "application/proto",
  "content-length": req.length,
});
r.on("error", () => process.exit(1));
r.on("data", () => {});
r.on("end", () => { c.close(); setTimeout(() => process.exit(0), 200); });
r.end(req);
"""
            env = os.environ.copy()
            env["REQ"] = os.path.join(run, "req.bin")
            env["PORT"] = str(port)
            subprocess.check_call(
                [node, "--tls-keylog=" + keylog, "-e", client],
                timeout=20, env=env,
            )
            deadline = time.time() + 3
            while time.time() < deadline and os.path.isfile(pcap) and os.path.getsize(pcap) < 1000:
                time.sleep(0.05)
            subprocess.call(
                ["sudo", "-n", "kill", "-INT", str(td.pid)],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
            try:
                td.wait(timeout=3)
            except subprocess.TimeoutExpired:
                td.kill()
            sp.terminate()
            self.assertGreater(os.path.getsize(pcap), 24, "pcap vide (tcpdump n'a pas livré les paquets)")
            self.assertGreater(os.path.getsize(keylog), 0)
            os.environ["RPC_TRAFFIC_LOG_DIR"] = os.path.join(run, "logs")
            n = dk.decode_once(
                pcap, keylog, os.path.join(run, "seen"),
                [str(port)], dk.re.compile(dk.DEFAULT_ALLOW), "198.18.0.1",
                shutil.which("tshark"),
            )
            self.assertGreaterEqual(n, 1)
            log = os.path.join(run, "logs", os.listdir(os.path.join(run, "logs"))[0])
            with open(log) as fh:
                rows = [json.loads(l) for l in fh if l.strip()]
            hit = [r for r in rows if r.get("method") == PATH and r.get("resp_body")]
            self.assertTrue(hit, rows)
            skills = hit[0]["resp_body"]["skills"]
            self.assertEqual(skills[0]["content"], "D" * 17000)
            self.assertEqual(hit[0]["req_body"]["bc_id"], "bc")
        finally:
            for proc in (td, sp):
                if proc is None or proc.poll() is not None:
                    continue
                proc.terminate()
                try:
                    proc.wait(timeout=2)
                except Exception:
                    proc.kill()
            shutil.rmtree(run, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
