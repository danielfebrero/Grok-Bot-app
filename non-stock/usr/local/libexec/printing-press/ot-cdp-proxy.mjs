#!/usr/bin/env node
import { createHash, randomBytes } from "node:crypto";
import { writeFileSync } from "node:fs";
import { createServer } from "node:http";
import { Socket } from "node:net";

const WS_GUID = "258EAFA5-E914-47DA-95CA-C5AB0DC85B11";
const DENY_EXACT = new Set([
  "Runtime.evaluate",
  "Runtime.callFunctionOn",
  "Runtime.compileScript",
  "Runtime.addBinding",
  "Network.getCookies",
  "Network.getAllCookies",
  "Page.captureScreenshot",
  "Page.printToPDF",
  "Page.handleJavaScriptDialog",
  "Page.startScreencast",
  "Browser.close",
  "Browser.crash",
  "Browser.grantPermissions",
]);
const DENY_PREFIX = ["Input.", "Storage.", "IndexedDB.", "Fetch."];
const ALLOW_EXACT = new Set([
  "Browser.getVersion",
  "DOM.disable",
  "DOM.enable",
  "DOM.getDocument",
  "Emulation.setDeviceMetricsOverride",
  "Emulation.setTouchEmulationEnabled",
  "Emulation.setUserAgentOverride",
  "Inspector.disable",
  "Inspector.enable",
  "Log.disable",
  "Log.enable",
  "Network.disable",
  "Network.enable",
  "Network.getRequestPostData",
  "Network.getResponseBody",
  "Network.setCacheDisabled",
  "Network.setExtraHTTPHeaders",
  "Network.setUserAgentOverride",
  "Page.bringToFront",
  "Page.createIsolatedWorld",
  "Page.disable",
  "Page.enable",
  "Page.getFrameTree",
  "Page.setLifecycleEventsEnabled",
  "Page.stopLoading",
  "Performance.disable",
  "Performance.enable",
  "Runtime.disable",
  "Runtime.enable",
  "Runtime.runIfWaitingForDebugger",
  "Target.activateTarget",
  "Target.attachToTarget",
  "Target.closeTarget",
  "Target.createTarget",
  "Target.detachFromTarget",
  "Target.getTargetInfo",
  "Target.getTargets",
  "Target.setAutoAttach",
  "Target.setDiscoverTargets",
]);
const REVIEWED_PARAMS = new Set([
  "Network.setCookie",
  "Page.addScriptToEvaluateOnNewDocument",
  "Page.navigate",
]);
const HTTP_ALLOW = new Set(["/json", "/json/list", "/json/protocol", "/json/version"]);

function normalizeScript(source) {
  return String(source).replace(/\s+/g, " ").trim();
}

// https://github.com/mvanhorn/printing-press-library/blob/66fd004fc2149cfad2abd737554fda08eb1238d5/library/food-and-dining/table-reservation-goat/internal/source/opentable/chrome_avail.go
const STEALTH_SCRIPT = normalizeScript(`
  Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
  window.chrome = window.chrome || { runtime: {}, app: {} };
  Object.defineProperty(navigator, 'plugins', { get: () => [1,2,3,4,5] });
  Object.defineProperty(navigator, 'languages', { get: () => ['en-US','en'] });
  const origQuery = navigator.permissions && navigator.permissions.query;
  if (origQuery) {
    navigator.permissions.query = (p) =>
      p && p.name === 'notifications'
        ? Promise.resolve({ state: Notification.permission })
        : origQuery.call(navigator.permissions, p);
  }
`);

export function isStealthScript(params) {
  return normalizeScript(params?.source ?? "") === STEALTH_SCRIPT;
}

export function isOpenTableUrl(value) {
  if (value === "about:blank" || value === "") return true;
  if (!URL.canParse(value)) return false;
  const parsed = new URL(value);
  if (parsed.protocol !== "https:") return false;
  const host = parsed.hostname.toLowerCase();
  return host === "opentable.com" || host.endsWith(".opentable.com");
}

export function isOpenTableCookie(params) {
  const domain = String(params?.domain ?? "")
    .toLowerCase()
    .replace(/^\./, "");
  if (domain === "opentable.com" || domain.endsWith(".opentable.com")) return true;
  return typeof params?.url === "string" && isOpenTableUrl(params.url);
}

function innerCommands(message) {
  if (message?.method === "Target.sendMessageToTarget") {
    try {
      const inner = JSON.parse(String(message.params?.message ?? ""));
      if (typeof inner?.method === "string") {
        return [{ method: inner.method, params: inner.params ?? {} }];
      }
    } catch (error) {
      return [{ method: "Target.sendMessageToTarget", params: { error: String(error) } }];
    }
    return [{ method: "Target.sendMessageToTarget", params: {} }];
  }
  if (typeof message?.method === "string") {
    return [{ method: message.method, params: message.params ?? {} }];
  }
  return [];
}

export function reviewCdpCommand(message) {
  for (const command of innerCommands(message)) {
    const { method, params } = command;
    if (DENY_EXACT.has(method) || DENY_PREFIX.some((prefix) => method.startsWith(prefix))) {
      return { ok: false, method };
    }
    if (method === "Page.navigate" || method === "Target.createTarget") {
      if (!isOpenTableUrl(String(params.url ?? ""))) {
        return { ok: false, method };
      }
    }
    if (method === "Network.setCookie" && !isOpenTableCookie(params)) {
      return { ok: false, method };
    }
    if (method === "Page.addScriptToEvaluateOnNewDocument" && !isStealthScript(params)) {
      return { ok: false, method };
    }
    if (!ALLOW_EXACT.has(method) && !REVIEWED_PARAMS.has(method)) {
      return { ok: false, method };
    }
  }
  return { ok: true };
}

function rewriteDebuggerUrls(value, listenOrigin) {
  if (Array.isArray(value)) {
    return value.map((entry) => rewriteDebuggerUrls(entry, listenOrigin));
  }
  if (value && typeof value === "object") {
    const next = { ...value };
    for (const [key, child] of Object.entries(next)) {
      if (
        (key === "webSocketDebuggerUrl" || key === "devtoolsFrontendUrl") &&
        typeof child === "string"
      ) {
        if (URL.canParse(child) && URL.canParse(listenOrigin)) {
          const parsed = new URL(child);
          parsed.protocol = key === "webSocketDebuggerUrl" ? "ws:" : parsed.protocol;
          parsed.host = new URL(listenOrigin).host;
          next[key] = parsed.toString();
        } else {
          next[key] = child;
        }
      } else {
        next[key] = rewriteDebuggerUrls(child, listenOrigin);
      }
    }
    return next;
  }
  return value;
}

function wsAccept(key) {
  return createHash("sha1")
    .update(key + WS_GUID)
    .digest("base64");
}

function encodeFrame(payload, opcode = 1, mask = false) {
  const data = Buffer.from(payload);
  const length = data.length;
  const maskBit = mask ? 0x80 : 0;
  let header;
  if (length < 126) {
    header = Buffer.from([0x80 | opcode, maskBit | length]);
  } else if (length < 65536) {
    header = Buffer.alloc(4);
    header[0] = 0x80 | opcode;
    header[1] = maskBit | 126;
    header.writeUInt16BE(length, 2);
  } else {
    header = Buffer.alloc(10);
    header[0] = 0x80 | opcode;
    header[1] = maskBit | 127;
    header.writeBigUInt64BE(BigInt(length), 2);
  }
  if (!mask) return Buffer.concat([header, data]);
  const key = randomBytes(4);
  const masked = Buffer.from(data);
  for (let i = 0; i < masked.length; i += 1) masked[i] ^= key[i % 4];
  return Buffer.concat([header, key, masked]);
}

function createFrameDecoder() {
  const header = Buffer.alloc(14);
  const chunks = [];
  let headerLength = 2;
  let headerBytes = 0;
  let payloadLength = 0;
  let payloadBytes = 0;
  return (chunk) => {
    const frames = [];
    let offset = 0;
    while (offset < chunk.length) {
      if (headerBytes < headerLength) {
        const end = offset + Math.min(headerLength - headerBytes, chunk.length - offset);
        headerBytes += chunk.copy(header, headerBytes, offset, end);
        offset = end;
        if (headerBytes < headerLength) break;
        const length = header[1] & 0x7f;
        if (headerLength === 2) {
          if (length === 126) headerLength += 2;
          else if (length === 127) headerLength += 8;
          if (header[1] & 0x80) headerLength += 4;
          if (headerBytes < headerLength) continue;
        }
        payloadLength = length;
        if (length === 126) payloadLength = header.readUInt16BE(2);
        else if (length === 127) payloadLength = Number(header.readBigUInt64BE(2));
      }
      const end = offset + Math.min(payloadLength - payloadBytes, chunk.length - offset);
      if (end > offset) chunks.push(chunk.subarray(offset, end));
      payloadBytes += end - offset;
      offset = end;
      if (payloadBytes < payloadLength) break;
      const payload = Buffer.concat(chunks, payloadLength);
      if (header[1] & 0x80) {
        const mask = header.subarray(headerLength - 4, headerLength);
        for (let i = 0; i < payload.length; i += 1) payload[i] ^= mask[i % 4];
      }
      frames.push({ opcode: header[0] & 0x0f, payload });
      chunks.length = 0;
      headerLength = 2;
      headerBytes = 0;
      payloadLength = 0;
      payloadBytes = 0;
    }
    return frames;
  };
}

function connectUpstreamWs(url) {
  return new Promise((resolve, reject) => {
    const parsed = new URL(url);
    const socket = new Socket();
    const key = randomBytes(16).toString("base64");
    socket.once("error", reject);
    socket.connect(Number(parsed.port || 80), parsed.hostname, () => {
      socket.write(
        `GET ${parsed.pathname}${parsed.search} HTTP/1.1\r\n` +
          `Host: ${parsed.host}\r\n` +
          "Upgrade: websocket\r\n" +
          "Connection: Upgrade\r\n" +
          `Sec-WebSocket-Key: ${key}\r\n` +
          "Sec-WebSocket-Version: 13\r\n\r\n",
      );
    });
    const chunks = [];
    const terminator = Buffer.from("\r\n\r\n");
    let matched = 0;
    let headerLength = 0;
    const onData = (chunk) => {
      chunks.push(chunk);
      for (const byte of chunk) {
        headerLength += 1;
        if (byte === terminator[matched]) matched += 1;
        else matched = byte === terminator[0] ? 1 : 0;
        if (matched !== terminator.length) continue;
        socket.off("data", onData);
        socket.off("error", reject);
        resolve({ socket, leftover: Buffer.concat(chunks).subarray(headerLength) });
        return;
      }
    };
    socket.on("data", onData);
  });
}

async function proxyJsonHttp(req, res, upstream, origin) {
  const path = (req.url ?? "/").replace(/[?#].*$/, "").replace(/\/$/, "");
  if (req.method !== "GET" || !HTTP_ALLOW.has(path)) {
    res.writeHead(403);
    res.end();
    return;
  }
  const upstreamRes = await fetch(`${upstream.replace(/\/$/, "")}${path}`);
  const text = await upstreamRes.text();
  let body = text;
  if (text.startsWith("{") || text.startsWith("[")) {
    body = JSON.stringify(rewriteDebuggerUrls(JSON.parse(text), origin()));
  }
  res.writeHead(upstreamRes.status, { "content-type": "application/json" });
  res.end(body);
}

export async function startOtCdpProxy(upstream, listenHost = "127.0.0.1") {
  const version = await fetch(`${upstream.replace(/\/$/, "")}/json/version`);
  if (!version.ok) throw new Error(`/json/version HTTP ${version.status}`);
  const versionBody = await version.json();
  const upstreamWs = versionBody.webSocketDebuggerUrl;
  if (typeof upstreamWs !== "string" || upstreamWs.length === 0) {
    throw new Error("upstream missing webSocketDebuggerUrl");
  }

  const server = createServer((req, res) => {
    void proxyJsonHttp(req, res, upstream, () => listenOrigin).catch((error) => {
      res.writeHead(502);
      res.end(String(error));
    });
  });

  let listenOrigin = "";
  server.on("upgrade", async (req, socket) => {
    const key = req.headers["sec-websocket-key"];
    if (req.method !== "GET" || typeof key !== "string") {
      socket.destroy();
      return;
    }
    socket.write(
      "HTTP/1.1 101 Switching Protocols\r\n" +
        "Upgrade: websocket\r\n" +
        "Connection: Upgrade\r\n" +
        `Sec-WebSocket-Accept: ${wsAccept(key)}\r\n\r\n`,
    );
    let upstreamConn;
    try {
      upstreamConn = await connectUpstreamWs(upstreamWs);
    } catch (error) {
      socket.destroy(error);
      return;
    }
    const decodeClient = createFrameDecoder();
    const decodeUpstream = createFrameDecoder();
    const pumpUpstream = (chunk) => {
      for (const frame of decodeUpstream(chunk)) {
        if (frame.opcode === 8) {
          socket.end(encodeFrame(frame.payload, 8));
          upstreamConn.socket.end();
          return;
        }
        if (frame.opcode === 9) {
          socket.write(encodeFrame(frame.payload, 10));
          continue;
        }
        if (frame.opcode === 1 || frame.opcode === 2) {
          socket.write(encodeFrame(frame.payload, frame.opcode));
        }
      }
    };
    if (upstreamConn.leftover.length) pumpUpstream(upstreamConn.leftover);
    socket.on("data", (chunk) => {
      for (const frame of decodeClient(chunk)) {
        if (frame.opcode === 8) {
          upstreamConn.socket.end();
          socket.end();
          return;
        }
        if (frame.opcode === 9) {
          socket.write(encodeFrame(frame.payload, 10));
          continue;
        }
        if (frame.opcode !== 1) continue;
        let message;
        try {
          message = JSON.parse(frame.payload.toString("utf8"));
        } catch (error) {
          socket.write(encodeFrame(JSON.stringify({ error: { message: String(error) } })));
          continue;
        }
        const review = reviewCdpCommand(message);
        if (!review.ok) {
          if (message.id != null) {
            socket.write(
              encodeFrame(
                JSON.stringify({
                  id: message.id,
                  error: { message: `blocked ${review.method}` },
                }),
              ),
            );
          }
          continue;
        }
        upstreamConn.socket.write(encodeFrame(frame.payload, 1, true));
      }
    });
    upstreamConn.socket.on("data", pumpUpstream);
    const close = () => {
      socket.destroy();
      upstreamConn.socket.destroy();
    };
    socket.on("close", close);
    socket.on("error", close);
    upstreamConn.socket.on("close", close);
    upstreamConn.socket.on("error", close);
  });

  await new Promise((resolve) => server.listen(0, listenHost, resolve));
  const addr = server.address();
  listenOrigin = `http://${listenHost}:${addr.port}`;
  return { server, url: listenOrigin, close: () => server.close() };
}

function parseArgs(argv) {
  const out = {};
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === "--upstream" || arg === "--url-file") {
      out[arg.slice(2)] = argv[i + 1];
      i += 1;
    }
  }
  return out;
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const args = parseArgs(process.argv.slice(2));
  if (!args.upstream || !args["url-file"]) {
    process.stderr.write(
      "usage: ot-cdp-proxy.mjs --upstream http://127.0.0.1:PORT --url-file PATH\n",
    );
    process.exit(2);
  }
  const proxy = await startOtCdpProxy(args.upstream);
  writeFileSync(args["url-file"], proxy.url);
}
