#!/usr/bin/env python3
"""Rotation du pcap au cap : tcpdump relancé, seen conservé, pas de secret dans les logs.

Pas de root, pas de vrai tcpdump : un binaire factice écrit « listening on »
et un petit pcap. Le tshark factice rend un échange déjà vu et un handshake neuf.
"""
from __future__ import annotations

import contextlib
import io
import json
import os
import signal
import stat
import subprocess
import sys
import tempfile
import textwrap
import unittest
import unittest.mock
from pathlib import Path

_LI = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_MON = os.path.join(os.path.dirname(_LI), "rpc-traffic-monitor")
for _p in (_MON, os.path.join(_MON, "mitmproxy"), _LI):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import decode_keylog as dk  # noqa: E402

_SECRET = b"CLIENT_TRAFFIC_SECRET_SHOULD_NOT_LOG"
_OLD_PATH = "/aiserver.v1.Old/Call"
_NEW_PATH = "/aiserver.v1.New/Call"
_FILTER = "tcp port 443 and not host 198.18.0.1"


def _pkt(stream: str, src: str, sport: str, dst: str, dport: str, http: dict, sni: str | None = None) -> dict:
    layers = {
        "ip": {"ip.src": src, "ip.dst": dst},
        "tcp": {"tcp.srcport": sport, "tcp.dstport": dport, "tcp.stream": stream},
        "http": http,
    }
    if sni:
        layers["tls"] = {"tls.handshake.extensions_server_name": sni}
    return {"_source": {"layers": layers}}


def _exchange(stream: str, path: str) -> list:
    req = _pkt(
        stream, "10.0.0.8", "40000", "1.2.3.4", "443",
        {
            "http.request.method": "POST",
            "http.request.uri": path,
            "http.host": "api2.cursor.sh",
        },
        sni="api2.cursor.sh",
    )
    resp = _pkt(
        stream, "1.2.3.4", "443", "10.0.0.8", "40000",
        {"http.response.code": "200"},
    )
    return [req, resp]


def _write_frames(directory: str) -> None:
    old = _exchange("4", _OLD_PATH)
    both = old + _exchange("9", _NEW_PATH)
    Path(directory, "old.json").write_text(json.dumps(old), encoding="utf-8")
    Path(directory, "both.json").write_text(json.dumps(both), encoding="utf-8")


def _write_fake_tcpdump(path: str) -> None:
    body = r'''
import os, sys, time, signal
log = os.environ.get("FAKE_TCPDUMP_ARGV", "")
if log:
    with open(log, "a", encoding="utf-8") as fh:
        fh.write("\0".join(sys.argv) + "\n")
argv = sys.argv[1:]
pcap = ""
i = 0
while i < len(argv):
    if argv[i] == "-w" and i + 1 < len(argv):
        pcap = argv[i + 1]
        i += 2
        continue
    i += 1
mode = os.environ.get("FAKE_TCPDUMP_MODE", "ok")
if mode == "ok" and pcap:
    with open(pcap, "wb") as fh:
        fh.write(b"NEW-PCAP\n" + b"Z" * 40)
    sys.stdout.write("tcpdump: listening on fake, link-type EN10MB\n")
    sys.stdout.flush()
    stop = {"v": False}
    def _h(signum, frame):
        stop["v"] = True
    signal.signal(signal.SIGTERM, _h)
    signal.signal(signal.SIGINT, _h)
    while not stop["v"]:
        time.sleep(0.05)
    raise SystemExit(0)
sys.stdout.write("tcpdump: failed to start\n")
sys.stdout.flush()
time.sleep(0.4)
raise SystemExit(1)
'''
    Path(path).write_text("#!/usr/bin/env python3\n" + textwrap.dedent(body), encoding="utf-8")
    os.chmod(path, 0o755)


def _write_fake_tshark(path: str) -> None:
    script = r'''#!/usr/bin/env python3
import os, sys
pcap = sys.argv[sys.argv.index("-r") + 1]
data = open(pcap, "rb").read()
base = os.environ["FAKE_FRAMES_DIR"]
name = "old.json" if data.startswith(b"OLD-PCAP") else "both.json"
sys.stdout.write(open(os.path.join(base, name), encoding="utf-8").read())
'''
    Path(path).write_text(script, encoding="utf-8")
    os.chmod(path, 0o755)


def _jsonl_rows(log_dir: str) -> list:
    rows = []
    if not os.path.isdir(log_dir):
        return rows
    for name in sorted(os.listdir(log_dir)):
        if not name.endswith(".jsonl"):
            continue
        with open(os.path.join(log_dir, name), encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if line:
                    rows.append(json.loads(line))
    return rows


class PcapRotateTests(unittest.TestCase):
    def setUp(self):
        self._stop = dk._stop
        self._settle = dk.ROTATE_SETTLE_S
        dk._stop = False
        dk.ROTATE_SETTLE_S = 0
        self.tmp = tempfile.mkdtemp(prefix="pcap-rotate-")
        self._env = {
            "RPC_TRAFFIC_LOG_DIR": os.environ.get("RPC_TRAFFIC_LOG_DIR"),
            "HB_ALLOW_HOSTS": os.environ.get("HB_ALLOW_HOSTS"),
            "HB_TUNNEL_ADDR": os.environ.get("HB_TUNNEL_ADDR"),
            "FAKE_TCPDUMP_ARGV": os.environ.get("FAKE_TCPDUMP_ARGV"),
            "FAKE_TCPDUMP_MODE": os.environ.get("FAKE_TCPDUMP_MODE"),
            "FAKE_FRAMES_DIR": os.environ.get("FAKE_FRAMES_DIR"),
        }
        self.procs = []
        self.pidfiles = []
        os.environ["RPC_TRAFFIC_LOG_DIR"] = os.path.join(self.tmp, "logs")
        os.environ["HB_ALLOW_HOSTS"] = dk.DEFAULT_ALLOW
        os.environ["HB_TUNNEL_ADDR"] = "198.18.0.1"
        os.environ["FAKE_FRAMES_DIR"] = self.tmp
        os.environ["FAKE_TCPDUMP_ARGV"] = os.path.join(self.tmp, "argv.log")
        _write_frames(self.tmp)
        self.tshark = os.path.join(self.tmp, "tshark")
        self.tcpdump = os.path.join(self.tmp, "tcpdump")
        _write_fake_tshark(self.tshark)
        self.keylog = os.path.join(self.tmp, "keys.log")
        Path(self.keylog).write_text("fixture-not-a-secret\n", encoding="utf-8")
        self.seen = os.path.join(self.tmp, "seen")
        self.pcap = os.path.join(self.tmp, "cap.pcap")
        self.pidfile = os.path.join(self.tmp, "tcpdump.pid")
        self.logout = os.path.join(self.tmp, "capture.out")
        self.pidfiles.append(self.pidfile)

    def tearDown(self):
        dk._stop = False
        for pidfile in self.pidfiles:
            pid = dk.read_pidfile(pidfile)
            if dk.pid_alive(pid):
                try:
                    os.kill(pid, signal.SIGTERM)
                except OSError:
                    pass
                dk.wait_pid_dead(pid, 2.0)
        for proc in self.procs:
            if proc.poll() is None:
                proc.terminate()
            try:
                proc.wait(timeout=2)
            except subprocess.TimeoutExpired:
                proc.kill()
        dk._stop = self._stop
        dk.ROTATE_SETTLE_S = self._settle
        signal.signal(signal.SIGTERM, signal.SIG_DFL)
        signal.signal(signal.SIGINT, signal.SIG_DFL)
        for key, val in self._env.items():
            if val is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = val
        # le pcap factice ne contient pas de vrai secret ; on retire l'arbre
        # après avoir tué tcpdump (il tient le fichier ouvert).
        subprocess.run(["rm", "-rf", self.tmp], check=False)

    def _sleeper(self) -> subprocess.Popen:
        proc = subprocess.Popen(
            [sys.executable, "-c", "import time; time.sleep(120)"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        self.procs.append(proc)
        Path(self.pidfile).write_text(f"{proc.pid}\n", encoding="utf-8")
        return proc

    def _old_pcap(self) -> None:
        Path(self.pcap).write_bytes(b"OLD-PCAP\n" + _SECRET + b"\n" + b"Y" * 400)

    def _follow_argv(self, extra: list[str]) -> list[str]:
        return [
            "--follow", "--interval", "0.2", "--max-bytes", "100",
            "--pcap", self.pcap, "--keylog", self.keylog, "--seen", self.seen,
            "--tshark", self.tshark, "--tcpdump-pidfile", self.pidfile,
            *extra,
        ]

    def _run_main(self, argv: list[str]) -> tuple[int, str]:
        buf = io.StringIO()
        with contextlib.redirect_stderr(buf):
            rc = dk.main(argv)
        return rc, buf.getvalue()

    def test_tcpdump_argv_keeps_filter_and_optional_buffer(self):
        cmd = dk.tcpdump_argv(
            "tcpdump", "lo", "/tmp/x.pcap",
            "tcp and (port 9229 or port 9302) and not host 198.18.0.1",
            8192,
        )
        self.assertEqual(cmd[-1], "tcp and (port 9229 or port 9302) and not host 198.18.0.1")
        self.assertEqual(cmd[cmd.index("-B") + 1], "8192")
        self.assertIn("--immediate-mode", cmd)
        plain = dk.tcpdump_argv("tcpdump", "any", "/tmp/x.pcap", "tcp port 443", 0)
        self.assertNotIn("-B", plain)
        with self.assertRaises(ValueError):
            dk.tcpdump_argv("tcpdump", "lo", "/tmp/x.pcap", "  ", 8192)

    def test_rotate_decodes_old_bytes_then_respawns(self):
        _write_fake_tcpdump(self.tcpdump)
        os.environ["FAKE_TCPDUMP_MODE"] = "ok"
        self._old_pcap()
        sleeper = self._sleeper()
        seen = []

        def decode_fn():
            seen.append(Path(self.pcap).read_bytes()[:8])

        err = io.StringIO()
        with contextlib.redirect_stderr(err):
            ok = dk.rotate_capture(
                pcap=self.pcap,
                pidfile=self.pidfile,
                tcpdump=self.tcpdump,
                iface="lo",
                capture_filter=_FILTER,
                logout=self.logout,
                buffer_kb=4096,
                decode_fn=decode_fn,
                settle_s=0,
            )
        self.assertTrue(ok, err.getvalue())
        self.assertIn("ClientHello", err.getvalue())
        self.assertNotIn(_SECRET.decode("ascii"), err.getvalue())
        self.assertEqual(seen, [b"OLD-PCAP"])
        self.assertFalse(os.path.exists(self.pcap + ".closed"))
        self.assertTrue(Path(self.pcap).read_bytes().startswith(b"NEW-PCAP"))
        self.assertEqual(os.stat(self.pcap).st_mode & 0o777, 0o600)
        self.assertFalse(dk.pid_alive(sleeper.pid))
        new_pid = dk.read_pidfile(self.pidfile)
        self.assertTrue(dk.pid_alive(new_pid))
        self.assertNotEqual(new_pid, sleeper.pid)
        argv = Path(os.environ["FAKE_TCPDUMP_ARGV"]).read_text(encoding="utf-8").strip().split("\0")
        self.assertEqual(argv[argv.index("-B") + 1], "4096")
        self.assertEqual(argv[argv.index("-i") + 1], "lo")
        self.assertEqual(argv[argv.index("-w") + 1], self.pcap)
        self.assertEqual(argv[-1], _FILTER)
        self.assertNotIn(_SECRET.decode("ascii"), Path(self.logout).read_text(encoding="utf-8"))
        # le marqueur ne reste ni dans le pcap neuf ni dans un reliquat
        self.assertNotIn(_SECRET, Path(self.pcap).read_bytes())

    def test_follow_rotates_dedups_seen_and_logs_new_handshake(self):
        _write_fake_tcpdump(self.tcpdump)
        os.environ["FAKE_TCPDUMP_MODE"] = "ok"
        self._old_pcap()
        self._sleeper()
        real_sleep = dk.time.sleep

        def fast_sleep(sec):
            if sec >= 0.2:
                dk._stop = True
                return
            real_sleep(0.01 if sec > 0 else 0)

        argv = self._follow_argv([
            "--tcpdump", self.tcpdump, "--iface", "lo",
            "--capture-filter", _FILTER, "--capture-out", self.logout,
            "--tcpdump-buffer-kb", "4096",
        ])
        with unittest.mock.patch.object(dk.time, "sleep", fast_sleep):
            rc, err = self._run_main(argv)
        self.assertEqual(rc, 0, err)
        self.assertIn("rotation — nouveau pcap, capture reprise", err)
        self.assertIn("ClientHello", err)
        self.assertNotIn("relancer hb-on", err)
        self.assertNotIn(_SECRET.decode("ascii"), err)
        rows = _jsonl_rows(os.environ["RPC_TRAFFIC_LOG_DIR"])
        methods = [r.get("method") for r in rows]
        self.assertEqual(methods, [_OLD_PATH, _NEW_PATH])
        seen_lines = [ln.strip() for ln in Path(self.seen).read_text(encoding="utf-8").splitlines() if ln.strip()]
        self.assertEqual(len(seen_lines), 2)
        self.assertEqual(len(set(seen_lines)), 2)
        # rejouer le pcap neuf (ancien échange + handshake) n'ajoute rien
        n = dk.decode_once(
            self.pcap, self.keylog, self.seen, [],
            dk.re.compile(dk.DEFAULT_ALLOW), "198.18.0.1", self.tshark,
        )
        self.assertEqual(n, 0)
        self.assertEqual(
            [r.get("method") for r in _jsonl_rows(os.environ["RPC_TRAFFIC_LOG_DIR"])],
            [_OLD_PATH, _NEW_PATH],
        )
        self.assertFalse(os.path.exists(self.pcap + ".closed"))
        self.assertEqual(stat.S_IMODE(os.stat(self.pcap).st_mode), 0o600)
        spawned = Path(os.environ["FAKE_TCPDUMP_ARGV"]).read_text(encoding="utf-8").strip().splitlines()
        self.assertEqual(len(spawned), 1)

    def test_follow_rotate_failure_shreds_closed_and_returns_4(self):
        _write_fake_tcpdump(self.tcpdump)
        os.environ["FAKE_TCPDUMP_MODE"] = "die"
        self._old_pcap()
        self._sleeper()
        argv = self._follow_argv([
            "--tcpdump", self.tcpdump, "--iface", "lo",
            "--capture-filter", _FILTER, "--capture-out", self.logout,
            "--tcpdump-buffer-kb", "1024",
        ])
        rc, err = self._run_main(argv)
        self.assertEqual(rc, 4, err)
        self.assertIn("n'écoute pas", err)
        self.assertNotIn(_SECRET.decode("ascii"), err)
        self.assertFalse(os.path.exists(self.pcap + ".closed"))
        for dirpath, _dirs, files in os.walk(self.tmp):
            for name in files:
                if not name.endswith((".pcap", ".closed")):
                    continue
                blob = Path(dirpath, name).read_bytes()
                self.assertNotIn(_SECRET, blob)
        rows = _jsonl_rows(os.environ["RPC_TRAFFIC_LOG_DIR"])
        self.assertEqual([r.get("method") for r in rows], [_OLD_PATH])

    def test_follow_without_spawn_args_still_stops(self):
        _write_fake_tcpdump(self.tcpdump)
        self._old_pcap()
        sleeper = self._sleeper()
        rc, err = self._run_main(self._follow_argv([]))
        self.assertEqual(rc, 0, err)
        self.assertIn("relancer hb-on", err)
        self.assertNotIn("rotation — nouveau pcap", err)
        self.assertFalse(dk.pid_alive(sleeper.pid))
        self.assertTrue(os.path.isfile(self.pcap))
        self.assertFalse(os.path.exists(os.environ["FAKE_TCPDUMP_ARGV"]))

    def test_shred_file_removes_without_printing_contents(self):
        path = os.path.join(self.tmp, "secret.pcap")
        Path(path).write_bytes(_SECRET)
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
            dk.shred_file(path)
        self.assertFalse(os.path.exists(path))
        self.assertNotIn(_SECRET.decode("ascii"), buf.getvalue())
        dk.shred_file(path)  # absent : no-op

    def test_hb_shred_pcap_removes_closed_sibling(self):
        pcap = os.path.join(self.tmp, "hb.pcap")
        closed = pcap + ".closed"
        Path(pcap).write_bytes(b"pcap-bytes")
        Path(closed).write_bytes(b"closed-bytes")
        lib = os.path.join(_LI, "lib.sh")
        proc = subprocess.run(
            ["bash", "-c", '. "$1"; hb_shred_pcap "$2"', "bash", lib, pcap],
            capture_output=True, text=True,
        )
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertFalse(os.path.exists(pcap))
        self.assertFalse(os.path.exists(closed))
        self.assertNotIn("pcap-bytes", proc.stdout + proc.stderr)
        self.assertNotIn("closed-bytes", proc.stdout + proc.stderr)


if __name__ == "__main__":
    unittest.main()
