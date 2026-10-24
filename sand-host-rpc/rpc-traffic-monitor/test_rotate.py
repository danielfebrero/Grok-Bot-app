#!/usr/bin/env python3
"""Cap des JSONL. Le fichier du jour, seul, ne doit pas grossir sans limite."""

from __future__ import annotations

import datetime
import tempfile
import unittest
from pathlib import Path

import proxy


def _write(path: Path, body: bytes) -> None:
    path.write_bytes(body)


class RotateTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def test_old_day_is_removed_and_today_stays(self) -> None:
        old = (datetime.date.today() - datetime.timedelta(days=40)).isoformat()
        today = datetime.date.today().isoformat()
        _write(self.dir / f"traffic-{old}.jsonl", b'{"a":1}\n')
        _write(self.dir / f"traffic-{today}.jsonl", b'{"a":2}\n')
        proxy.rotate_logs(str(self.dir), keep_days=14, cap_bytes=10_000_000)
        names = sorted(p.name for p in self.dir.glob("traffic-*.jsonl"))
        self.assertEqual(names, [f"traffic-{today}.jsonl"])

    def test_oldest_file_goes_first_when_two_exceed_the_cap(self) -> None:
        today = datetime.date.today().isoformat()
        yesterday = (datetime.date.today() - datetime.timedelta(days=1)).isoformat()
        _write(self.dir / f"traffic-{yesterday}.jsonl", b"y" * 800 + b"\n")
        kept = b'{"ok":1}\n'
        _write(self.dir / f"traffic-{today}.jsonl", kept)
        proxy.rotate_logs(str(self.dir), keep_days=14, cap_bytes=100)
        names = sorted(p.name for p in self.dir.glob("traffic-*.jsonl"))
        self.assertEqual(names, [f"traffic-{today}.jsonl"])
        self.assertEqual((self.dir / f"traffic-{today}.jsonl").read_bytes(), kept)

    def test_single_file_over_cap_keeps_the_tail(self) -> None:
        today = datetime.date.today().isoformat()
        path = self.dir / f"traffic-{today}.jsonl"
        lines = [f'{{"n":{i}}}\n'.encode() for i in range(50)]
        path.write_bytes(b"".join(lines))
        original = path.read_bytes()
        proxy.rotate_logs(str(self.dir), keep_days=14, cap_bytes=80)
        got = path.read_bytes()
        self.assertLessEqual(len(got), 80)
        self.assertTrue(got.endswith(b"\n"))
        self.assertTrue(original.endswith(got))
        self.assertGreater(len(got), 0)
        self.assertFalse((self.dir / f"traffic-{today}.jsonl.cap.tmp").exists())

    def test_file_under_cap_is_untouched(self) -> None:
        today = datetime.date.today().isoformat()
        path = self.dir / f"traffic-{today}.jsonl"
        body = b'{"n":1}\n'
        path.write_bytes(body)
        proxy.rotate_logs(str(self.dir), keep_days=14, cap_bytes=10_000)
        self.assertEqual(path.read_bytes(), body)


if __name__ == "__main__":
    unittest.main()
