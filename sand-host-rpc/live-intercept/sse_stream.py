#!/usr/bin/env python3
"""Lecture SSE du relai agent↔host.

`/events` (gateway host-main) est un flux `text/event-stream` sans fin utile
avant la fermeture : plusieurs événements `data:` (souvent `channel=transcript`),
des commentaires `:ping` toutes les ~15 s, et parfois un `Content-Length` qui
ne couvre que le premier envoi. Le relai doit :

* transmettre au client au fil de l'eau (pas après la fin du flux) ;
* reassembler un événement coupé entre deux chunks ;
* exposer CHAQUE événement, pas seulement le premier bloc.

Le décodage gzip incrémental n'est là que si l'upstream l'envoie quand même.
Le relai ne demande pas gzip sur `/events` (cf. relay_agent_host).
"""
from __future__ import annotations

import asyncio
import zlib


def is_events_path(path: str) -> bool:
    """GET /events (query autorisée). /events/echo n'est pas un flux."""
    bare = (path or "").split("?", 1)[0]
    return bare == "/events"


def is_event_stream(path: str, resp_headers_lc: dict, status: str | None = None) -> bool:
    ctype = (resp_headers_lc.get("content-type") or "").lower()
    if "text/event-stream" in ctype:
        return True
    if not is_events_path(path):
        return False
    if status is not None and not str(status).startswith("2"):
        return False
    # Une erreur JSON sur /events reste une réponse courte (content-length).
    if "json" in ctype:
        return False
    return True


class SseAssembler:
    """Reassemble des blocs SSE (`\\n\\n`) à travers des chunks arbitraires."""

    def __init__(self):
        self.buf = ""
        self.pings = 0
        self.events = 0

    def feed(self, data: bytes) -> list[dict]:
        if not data:
            return []
        self.buf += data.decode("utf-8", "replace")
        self._normalize()
        out = []
        while "\n\n" in self.buf:
            block, self.buf = self.buf.split("\n\n", 1)
            ev = parse_sse_block(block)
            if ev is None:
                continue
            if ev.get("ping") and ev.get("data") is None:
                self.pings += 1
                continue
            if ev.get("data") is None:
                continue
            self.events += 1
            out.append(ev)
        return out

    def flush(self) -> list[dict]:
        self._normalize(final=True)
        if not self.buf.strip():
            self.buf = ""
            return []
        ev = parse_sse_block(self.buf)
        self.buf = ""
        if ev is None or ev.get("data") is None:
            if ev and ev.get("ping"):
                self.pings += 1
            return []
        self.events += 1
        return [ev]

    def _normalize(self, final: bool = False) -> None:
        self.buf = self.buf.replace("\r\n", "\n")
        if not final and self.buf.endswith("\r"):
            self.buf = self.buf[:-1].replace("\r", "\n") + "\r"
        else:
            self.buf = self.buf.replace("\r", "\n")


def parse_sse_block(block: str) -> dict | None:
    data_lines = []
    event = ""
    eid = ""
    saw_field = False
    ping = False
    for line in block.split("\n"):
        if line == "":
            continue
        if line.startswith(":"):
            if line == ":ping" or line.startswith(":ping "):
                ping = True
            continue
        saw_field = True
        if line.startswith("data:"):
            data_lines.append(line[5:].lstrip(" "))
        elif line.startswith("event:"):
            event = line[6:].strip()
        elif line.startswith("id:"):
            eid = line[3:].strip()
    if not saw_field:
        if ping:
            return {"ping": True, "data": None, "event": "", "id": ""}
        return None
    data = "\n".join(data_lines) if data_lines else None
    return {"ping": False, "data": data, "event": event, "id": eid}


def _gzip_decoder(headers_lc: dict):
    enc = (headers_lc.get("content-encoding") or "").lower()
    if "gzip" in enc or "x-gzip" in enc:
        return zlib.decompressobj(16 + zlib.MAX_WBITS)
    return None


async def _write_client(cwriter, blob: bytes) -> None:
    try:
        cwriter.write(blob)
        await cwriter.drain()
    except Exception:
        # Le client peut partir ; on continue de lire pour le journal.
        pass


async def forward_sse(ureader, cwriter, status_line: str, resp_headers, on_event) -> dict:
    """Transmet le flux SSE et appelle on_event(dict) pour chaque data: complet.

    Chunked : tous les chunks jusqu'au chunk terminal (pas seulement le premier).
    Sinon (Content-Length, éventuellement trop court, ou ni l'un ni l'autre) :
    lecture jusqu'à EOF. Le client reçoit alors un corps ré-encodé chunked,
    sans Content-Length d'origine, pour ne pas s'arrêter au premier envoi.
    """
    hd = {k.lower(): v for k, v in resp_headers}
    asm = SseAssembler()
    decomp = _gzip_decoder(hd)

    def consume(payload: bytes) -> None:
        if not payload:
            return
        if decomp is not None:
            try:
                payload = decomp.decompress(payload)
            except Exception:
                return
        for ev in asm.feed(payload):
            on_event(ev)

    te = hd.get("transfer-encoding", "").lower()
    if "chunked" in te:
        head = [status_line]
        for k, v in resp_headers:
            head.append(f"{k}: {v}")
        await _write_client(cwriter, ("\r\n".join(head) + "\r\n\r\n").encode("latin1"))
        while True:
            try:
                size_line = await ureader.readline()
            except (asyncio.IncompleteReadError, ConnectionError):
                break
            if not size_line:
                break
            await _write_client(cwriter, size_line)
            try:
                size = int(size_line.strip().split(b";")[0] or b"0", 16)
            except ValueError:
                break
            if size == 0:
                while True:
                    try:
                        trailer = await ureader.readline()
                    except (asyncio.IncompleteReadError, ConnectionError):
                        trailer = b""
                    await _write_client(cwriter, trailer)
                    if trailer in (b"\r\n", b"\n", b""):
                        break
                break
            try:
                chunk = await ureader.readexactly(size)
                crlf = await ureader.readexactly(2)
            except asyncio.IncompleteReadError as e:
                consume(e.partial)
                await _write_client(cwriter, e.partial)
                break
            await _write_client(cwriter, chunk + crlf)
            consume(chunk)
    else:
        # Pas de framing chunked fiable : on lit jusqu'à EOF, y compris au-delà
        # d'un Content-Length trop court, et on ré-encode en chunked pour le client.
        head = [status_line]
        for k, v in resp_headers:
            if k.lower() in ("content-length", "transfer-encoding", "connection"):
                continue
            head.append(f"{k}: {v}")
        head.append("Transfer-Encoding: chunked")
        head.append("Connection: close")
        await _write_client(cwriter, ("\r\n".join(head) + "\r\n\r\n").encode("latin1"))
        while True:
            try:
                data = await ureader.read(8192)
            except (asyncio.IncompleteReadError, ConnectionError):
                break
            if not data:
                break
            sz = f"{len(data):x}\r\n".encode("ascii")
            await _write_client(cwriter, sz + data + b"\r\n")
            consume(data)
        await _write_client(cwriter, b"0\r\n\r\n")

    if decomp is not None:
        try:
            tail = decomp.flush()
        except Exception:
            tail = b""
        if tail:
            for ev in asm.feed(tail):
                on_event(ev)
    for ev in asm.flush():
        on_event(ev)
    return {"pings": asm.pings, "events": asm.events}
