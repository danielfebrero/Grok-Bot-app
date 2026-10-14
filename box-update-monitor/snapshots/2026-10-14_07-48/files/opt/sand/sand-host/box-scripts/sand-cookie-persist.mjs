import { randomBytes } from "node:crypto";
import {
  appendFileSync,
  constants,
  copyFileSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  renameSync,
  rmdirSync,
  rmSync,
  statSync,
  unlinkSync,
  writeFileSync,
} from "node:fs";
import { basename, dirname, join } from "node:path";
import { pathToFileURL } from "node:url";
import {
  CDP_PORT_BASE,
  connectBrowser,
  cookieFingerprint,
  cookieKey,
  cookieRecency,
  discoverMonitorPorts,
  isConnectionRefused,
  isRotatingAuthCookie,
  pushCookies,
  readCookies,
  toCookieParam,
} from "./cdp-cookies.mjs";

const SEED_PATH =
  process.env.SAND_COOKIE_SEED_PATH ?? "/home/box/sand-data/chrome-cookie-seed.json";
const CAPTURE_INTERVAL_MS = Number.parseInt(
  process.env.SAND_COOKIE_PERSIST_INTERVAL_MS ?? "5000",
  10,
);
const MAX_RESTORE_ATTEMPTS = 3;
const PRIMARY_CDP_PORT = CDP_PORT_BASE + 1;
const SHRINK_MIN_COOKIES = 30;
const SEED_VERSION = 1;
const CORRUPT_SEED_COPY_LIMIT = 3;
const HEARTBEAT_MS = 60 * 60 * 1000;
const CORRUPT_REPLACING = "captured:discarded_corrupt";
const TELEMETRY_LOG_PATH = process.env.SAND_BOX_TELEMETRY_LOG ?? "/tmp/sand-box-telemetry.log";

const CORRUPT_ENTRY_KINDS = ["copy", "stem", "unknown"];

export const PERSIST_LINES = Object.freeze({
  starting: "starting",
  "store-off": "durable box store not enabled; no-op",
  "not-box-user":
    "could not drop to the box user; refusing to run as root writing into box-owned data",
  "tick-failed": "tick failed",
  crashed: "crashed; exiting",
  "telemetry-append-failed": "telemetry append failed",
  "lock-unreadable": "seed lock dir unreadable, retrying",
  "lock-owner-kept": "seed lock owner file not removed",
  "lock-dir-kept": "seed lock dir not removed",
  "monitor-discovery-failed": "monitor discovery failed",
  "cdp-probe-failed": "CDP probe failed",
  "cdp-frame-dropped": "dropped a CDP frame that is not JSON",
  "jar-read-failed": "reading the live jar failed",
  "jar-refused-cookies": "the live jar refused seed cookies",
  "inject-failed": "injecting the seed into live Chrome failed",
  "primary-replaced": "primary Chrome was replaced; restoring the seed before the next capture",
  "guard-refused": "refusing to save over the seed after the primary Chrome changed",
  "copy-passed-over": "corrupt seed copy passed over a taken name",
  "copy-passed-over-not-file": "corrupt seed copy passed over a taken name that is not a file",
  "copies-not-pruned": "corrupt seed copies not pruned",
  ...Object.fromEntries(
    CORRUPT_ENTRY_KINDS.flatMap((kind) => [
      [`prune-skipped-${kind}`, `prune skipped ${kind} entry`],
      [`prune-skipped-${kind}-not-file`, `prune skipped ${kind} entry that is not a file`],
    ]),
  ),
  "capture-discarded-corrupt": "seed corrupt; discarded after copying it aside",
  "capture-captured": "seed holds the live jar",
  "capture-paused-unreadable": "seed unreadable; capture paused",
  "capture-paused-version": "seed version is not supported; capture paused",
  "capture-failed-lock": "capture failed at the seed lock",
  "capture-failed-copy": "capture failed copying the corrupt seed aside",
  "capture-failed-write": "capture failed writing the seed",
  "capture-failed-rename": "capture failed replacing the seed",
  "capture-failed-capture": "capture failed",
  "restore-waiting": "seed unreadable; restore waits",
  "restore-injected": "restored seed cookies into live Chrome",
  "restore-present": "the live jar already holds the seed cookies",
  "restore-empty": "no seed cookies to restore",
  "restore-incomplete": "proceeding to capture with seed cookies still absent from the live jar",
  "restore-seed-unusable": "seed corrupt or its version is not supported; nothing restored",
});

const PERSIST_COUNT_NAMES = new Set(["refused", "total", "cookies", "seed"]);

function faultCode(error) {
  const code = error?.code ?? error?.cause?.code;
  return typeof code === "string" && /^E[A-Z0-9]+$/.test(code) ? code : "unknown";
}

function lineCode(code) {
  return typeof code === "string" && /^E[A-Z0-9]+$/.test(code) ? code : "unknown";
}

export function formatPersistLine(id, code = null, counts = null) {
  const text = PERSIST_LINES[id];
  if (text == null) throw new TypeError("unknown sand-cookie-persist line id");
  let line = code == null ? text : `${text} (${lineCode(code)})`;
  for (const [name, value] of Object.entries(counts ?? {})) {
    if (PERSIST_COUNT_NAMES.has(name) && Number.isSafeInteger(value) && value >= 0) {
      line += ` ${name}=${value}`;
    }
  }
  return line;
}

function writeStderrLine(line) {
  process.stderr.write(`sand-cookie-persist ${line}\n`);
}

function isOutcomeLine(id) {
  return id.startsWith("capture-") || id.startsWith("restore-");
}

export function createPersistLog({
  write = writeStderrLine,
  now = Date.now,
  heartbeatMs = HEARTBEAT_MS,
} = {}) {
  const printed = new Map();
  return function persistLog(id, code = null, counts = null) {
    const line = formatPersistLine(id, code, counts);
    if (!isOutcomeLine(id)) {
      const at = now();
      const last = printed.get(id);
      if (last != null && last.line === line && at - last.at < heartbeatMs) return false;
      printed.set(id, { line, at });
    }
    write(line);
    return true;
  };
}

function plainPersistLog(id, code = null, counts = null) {
  writeStderrLine(formatPersistLine(id, code, counts));
}

function logLine(deps, id, code = null, counts = null) {
  (deps.log ?? plainPersistLog)(id, code, counts);
}

export function persistCdpReporters(log = plainPersistLog) {
  return {
    onUnreadable: (error) => log("monitor-discovery-failed", faultCode(error)),
    onDroppedFrame: () => log("cdp-frame-dropped"),
    onRefused: (_browser, refused, total, error) =>
      log("jar-refused-cookies", faultCode(error), { refused, total }),
  };
}

export function exitOnUncaught(log = plainPersistLog, exit = (code) => process.exit(code)) {
  process.on("uncaughtException", (error) => {
    log("crashed", faultCode(error));
    exit(1);
  });
}

function appendTelemetry(log, event) {
  try {
    appendFileSync(TELEMETRY_LOG_PATH, `${JSON.stringify(event)}\n`, "utf8");
  } catch (error) {
    log("telemetry-append-failed", faultCode(error));
  }
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export function dropToBoxUser(log = plainPersistLog) {
  if (typeof process.getuid !== "function" || process.getuid() !== 0) return true;
  try {
    if (typeof process.setgroups === "function") process.setgroups([]);
    process.setgid("box");
    process.setuid("box");
    return true;
  } catch (error) {
    log("not-box-user", faultCode(error));
    return false;
  }
}

const SEED_LOCK_TIMEOUT_MS = 5000;
const SEED_LOCK_RETRY_MS = 50;
const SEED_LOCK_STALE_MS = 30000;

function lockOwnerPath(lockDir) {
  return `${lockDir}/owner`;
}

function readLockOwner(fs, lockDir) {
  try {
    return fs.readFileSync(lockOwnerPath(lockDir), "utf8");
  } catch {
    return null;
  }
}

function lockOwnerPid(token) {
  const sep = token.indexOf(":");
  if (sep <= 0) return null;
  const pid = Number.parseInt(token.slice(0, sep), 10);
  return Number.isInteger(pid) && pid > 0 ? pid : null;
}

function processIsAlive(pid) {
  try {
    process.kill(pid, 0);
    return true;
  } catch {
    return false;
  }
}

function removeLockDir(fs, lockDir, log) {
  try {
    fs.unlinkSync(lockOwnerPath(lockDir));
  } catch (error) {
    if (error?.code !== "ENOENT") log("lock-owner-kept", faultCode(error));
  }
  try {
    fs.rmdirSync(lockDir);
    return true;
  } catch (error) {
    if (error?.code === "ENOENT") return true;
    log("lock-dir-kept", faultCode(error));
    return false;
  }
}

function lockIsStale(fs, lockDir, staleMs, log) {
  let heldSinceMs;
  try {
    heldSinceMs = fs.statSync(lockDir).mtimeMs;
  } catch (error) {
    if (error?.code !== "ENOENT") log("lock-unreadable", faultCode(error));
    return false;
  }
  if (Date.now() - heldSinceMs <= staleMs) return false;
  const existing = readLockOwner(fs, lockDir);
  const pid = existing != null ? lockOwnerPid(existing) : null;
  return pid == null || !processIsAlive(pid);
}

export async function withSeedLock(seedPath, fn, options = {}) {
  const fs = {
    mkdirSync,
    rmdirSync,
    statSync,
    writeFileSync,
    readFileSync,
    unlinkSync,
    ...(options.fs ?? {}),
  };
  const log = options.log ?? plainPersistLog;
  const now = options.now ?? Date.now;
  const wait = options.sleep ?? sleep;
  const lockDir = `${seedPath}.lock`;
  const token = `${process.pid}:${randomBytes(8).toString("hex")}`;
  const deadline = now() + (options.timeoutMs ?? SEED_LOCK_TIMEOUT_MS);
  let madeParent = false;
  for (;;) {
    try {
      fs.mkdirSync(lockDir);
    } catch (error) {
      if (error?.code === "ENOENT" && !madeParent) {
        fs.mkdirSync(dirname(lockDir), { recursive: true });
        madeParent = true;
        continue;
      }
      if (error?.code !== "EEXIST") throw error;
      const stale = lockIsStale(fs, lockDir, options.staleMs ?? SEED_LOCK_STALE_MS, log);
      if (stale && removeLockDir(fs, lockDir, log)) continue;
      if (now() >= deadline) {
        throw Object.assign(new Error("seed lock still held after timeout"), {
          code: "ETIMEDOUT",
        });
      }
      await wait(options.retryMs ?? SEED_LOCK_RETRY_MS);
      continue;
    }
    try {
      fs.writeFileSync(lockOwnerPath(lockDir), token);
    } catch (error) {
      removeLockDir(fs, lockDir, log);
      throw error;
    }
    break;
  }
  try {
    return await fn();
  } finally {
    if (readLockOwner(fs, lockDir) === token) {
      removeLockDir(fs, lockDir, log);
    }
  }
}

export function isCookiePersistEnabled(env = process.env) {
  const on = (value) => {
    const raw = value?.trim().toLowerCase();
    return raw === "1" || raw === "true" || raw === "yes";
  };
  return on(env.SAND_BOX_STORE_COPY_IN) || on(env.SAND_BOX_STORE_SYNC);
}

export function serializeSeed(cookies, now = Date.now()) {
  return JSON.stringify({ version: SEED_VERSION, savedAt: now, cookies });
}

export function parseSeed(text) {
  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch {
    return null;
  }
  const cookies = parsed?.cookies;
  return Array.isArray(cookies) ? cookies : null;
}

export function readSeed(fs, seedPath) {
  let text;
  try {
    text = fs.readFileSync(seedPath, "utf8");
  } catch (error) {
    return error?.code === "ENOENT"
      ? { kind: "absent" }
      : { kind: "unreadable", code: faultCode(error) };
  }
  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch {
    return { kind: "corrupt" };
  }
  const version = parsed?.version;
  if (typeof version === "number" && version !== SEED_VERSION) {
    return { kind: "unsupported-version", version };
  }
  const cookies = parsed?.cookies;
  return Array.isArray(cookies) ? { kind: "ok", cookies } : { kind: "corrupt" };
}

export function cookiesDigest(cookies) {
  const parts = cookies.map((c) =>
    [cookieKey(c), c.value, c.secure ? 1 : 0, c.httpOnly ? 1 : 0, c.sameSite ?? ""].join("\u0001"),
  );
  parts.sort();
  return parts.join("\u0002");
}

function latestExpiry(cookies) {
  return cookies.reduce((best, c) => (cookieRecency(c) > cookieRecency(best) ? c : best));
}

function observeJars(live) {
  return new Map(
    live.map(({ port, jar }) => [
      port,
      new Map([...jar].map(([key, c]) => [key, cookieFingerprint(c)])),
    ]),
  );
}

function changedSinceSeen(seen, key, live) {
  return live
    .filter(({ port, jar }) => {
      const before = seen.get(port);
      return before !== undefined && before.get(key) !== cookieFingerprint(jar.get(key));
    })
    .map(({ port, jar }) => jar.get(key));
}

function mergeJars(live, saved, seen, { keepUnseen }) {
  const keys = new Set(live.flatMap(({ jar }) => [...jar.keys()]));
  if (keepUnseen) for (const key of saved.keys()) keys.add(key);
  const merged = [];
  for (const key of keys) {
    const holders = live.filter(({ jar }) => jar.has(key));
    const savedCookie = saved.get(key);
    const changed = changedSinceSeen(seen, key, holders);
    const values = holders.map(({ jar }) => jar.get(key));
    let cookie = savedCookie;
    if (changed.length > 0) {
      cookie = latestExpiry(changed);
    } else if (values.length > 0) {
      cookie =
        values.find(
          (c) =>
            savedCookie !== undefined && cookieFingerprint(c) === cookieFingerprint(savedCookie),
        ) ?? latestExpiry(values);
    }
    if (!isRotatingAuthCookie(cookie.name)) merged.push(cookie);
  }
  return merged;
}

function restoreCandidates(target, others, saved, seen) {
  const missing = [];
  for (const [key, savedCookie] of saved) {
    if (target.has(key)) continue;
    const changed = changedSinceSeen(
      seen,
      key,
      others.filter(({ jar }) => jar.has(key)),
    );
    missing.push(changed.length > 0 ? latestExpiry(changed) : savedCookie);
  }
  return missing;
}

function dropsMostOfSeed(cookies, saved) {
  const mergedKeys = new Set(cookies.map(cookieKey));
  const kept = [...saved.keys()].filter((key) => mergedKeys.has(key)).length;
  return saved.size - kept >= SHRINK_MIN_COOKIES && kept * 2 <= saved.size;
}

export function classifySeedFault(code) {
  switch (code) {
    case "EACCES":
    case "EPERM":
      return "permission_denied";
    case "EISDIR":
    case "ENOTDIR":
    case "ELOOP":
      return "not_a_file";
    case "EIO":
      return "io_error";
    default:
      return "unknown";
  }
}

export function classifyRestoreOutcome(seedCount, missingAfter, corrupt = false) {
  if (corrupt) return "failed";
  if (seedCount === 0) return "empty";
  if (missingAfter <= 0) return "ok";
  if (missingAfter >= seedCount) return "failed";
  return "partial";
}

async function connectPrimary(deps) {
  try {
    return await deps.connect(PRIMARY_CDP_PORT);
  } catch (error) {
    if (isConnectionRefused(error)) return null;
    throw error;
  }
}

function monitorPorts(deps) {
  return [...new Set([PRIMARY_CDP_PORT, ...deps.discoverPorts()])].sort((a, b) => a - b);
}

async function readJar(deps, browser) {
  try {
    return { jar: await readCookies(browser), code: null };
  } catch (error) {
    const code = faultCode(error);
    logLine(deps, "jar-read-failed", code);
    return { jar: null, code };
  }
}

async function openJars(deps, ports, faults = {}) {
  const open = [];
  for (const port of ports) {
    let browser;
    try {
      browser = await deps.connect(port);
    } catch (error) {
      if (!isConnectionRefused(error)) {
        faults.code = faultCode(error);
        logLine(deps, "cdp-probe-failed", faults.code);
      }
      continue;
    }
    const { jar, code } = await readJar(deps, browser);
    if (jar === null) {
      faults.code = code;
      browser.close();
    } else {
      open.push({ port, browser, jar });
    }
  }
  return open;
}

class SeedStepError extends Error {
  constructor(step, cause) {
    super(`seed ${step} failed`, { cause });
    this.step = step;
    this.code = faultCode(cause);
  }
}

function seedStep(step, run) {
  try {
    return run();
  } catch (error) {
    throw new SeedStepError(step, error);
  }
}

function seedLockOptions(deps) {
  return { fs: deps.fs, log: deps.log, now: deps.now, sleep: deps.sleep, ...deps.seedLock };
}

// Chrome's cookie monster keeps every cookie in memory and, with persist_session_cookies_ off by default, synchronizes only non-session cookies to the SQLite store (https://chromium.googlesource.com/chromium/src/+/refs/tags/151.0.7922.169/net/cookies/cookie_monster.h), which commits every 30 seconds or 512 operations (https://chromium.googlesource.com/chromium/src/+/refs/tags/151.0.7922.169/net/extras/sqlite/sqlite_persistent_cookie_store.cc), so a copy of the `Cookies` file misses recent cookies and, by default, session cookies.
async function captureJars(deps, state, primary) {
  let held = false;
  try {
    return await withSeedLock(
      deps.seedPath,
      () => {
        held = true;
        return captureHeld(deps, state, primary);
      },
      seedLockOptions(deps),
    );
  } catch (error) {
    if (error instanceof SeedStepError) throw error;
    throw new SeedStepError(held ? "capture" : "lock", error);
  }
}

async function captureHeld(deps, state, primary) {
  const others = await openJars(
    deps,
    deps.discoverPorts().filter((port) => port !== PRIMARY_CDP_PORT),
  );
  try {
    const live = others.map(({ port, jar }) => ({ port, jar }));
    if (primary !== null) {
      const { jar } = await readJar(deps, primary);
      if (jar === null) return { outcome: null };
      live.unshift({ port: PRIMARY_CDP_PORT, jar });
    }
    if (live.length === 0) return { outcome: null };
    return mergeIntoSeed(deps, state, live, primary !== null);
  } finally {
    for (const { browser } of others) browser.close();
  }
}

function mergeIntoSeed(deps, state, live, fromPrimary) {
  const { capture } = state;
  const seed = readSeed(deps.fs, deps.seedPath);
  if (seed.kind === "unreadable") {
    return {
      outcome: "paused",
      digest: capture.digest,
      reason: classifySeedFault(seed.code),
      line: "capture-paused-unreadable",
      code: seed.code,
    };
  }
  if (seed.kind === "unsupported-version") {
    return {
      outcome: "paused",
      digest: capture.digest,
      reason: "unsupported_version",
      line: "capture-paused-version",
    };
  }
  const saved = new Map((seed.kind === "ok" ? seed.cookies : []).map((c) => [cookieKey(c), c]));
  const cookies = mergeJars(live, saved, state.seen, { keepUnseen: !fromPrimary });
  const seen = observeJars(live);
  if (fromPrimary && state.phase === "guarding" && dropsMostOfSeed(cookies, saved)) {
    logLine(deps, "guard-refused", null, { cookies: cookies.length, seed: saved.size });
    return { outcome: null, seen };
  }
  const digest = cookiesDigest(cookies);
  const phase = fromPrimary ? "capturing" : state.phase;
  const result = { outcome: "captured", digest, seedCookies: cookies.length, seen, phase };
  if (seed.kind === "ok" && capture.outcome === "captured" && digest === capture.digest) {
    return result;
  }
  const copy = seed.kind === "corrupt" ? seedStep("copy", () => copyCorruptSeed(deps)) : null;
  writeSeedFile(deps, serializeSeed(cookies, deps.now?.() ?? Date.now()));
  return copy == null ? result : { ...result, discardedCorrupt: true, keptCopy: copy };
}

function captureRow(result) {
  const row = { kind: "cookie_persist", phase: "capture", outcome: result.outcome };
  if (result.outcome === "paused") row.reason = result.reason;
  if (result.outcome === "captured") {
    row.seedCookies = result.seedCookies;
    if (result.discardedCorrupt) row.discardedCorrupt = true;
  }
  return row;
}

function captureRowKey(result) {
  if (result.outcome === "paused") return `paused:${result.reason}`;
  return result.discardedCorrupt ? CORRUPT_REPLACING : result.outcome;
}

function captureLine(result) {
  if (result.discardedCorrupt) return "capture-discarded-corrupt";
  return result.line ?? "capture-captured";
}

function settleCapture(deps, capture, result) {
  if (result.outcome == null) return capture;
  const rowKey = captureRowKey(result);
  const at = deps.now?.() ?? Date.now();
  const due = at - capture.lineAt >= HEARTBEAT_MS;
  const ships = rowKey !== capture.rowKey || (rowKey === CORRUPT_REPLACING && due);
  if (ships || due) logLine(deps, captureLine(result), result.code ?? null);
  if (ships) deps.emitTelemetry?.(captureRow(result));
  return {
    digest: result.digest,
    outcome: result.outcome,
    rowKey,
    lineAt: ships || due ? at : capture.lineAt,
  };
}

async function captureTick(deps, state, primary) {
  let result;
  try {
    result = await captureJars(deps, state, primary);
  } catch (error) {
    const step = error instanceof SeedStepError ? error.step : "capture";
    result = {
      outcome: "failed",
      digest: state.capture.digest,
      line: `capture-failed-${step}`,
      code: faultCode(error),
    };
  }
  const capture = settleCapture(deps, state.capture, result);
  if (result.keptCopy != null) pruneCorruptSeedCopies(deps, result.keptCopy);
  return {
    ...state,
    phase: result.phase ?? state.phase,
    seen: result.seen ?? state.seen,
    capture,
  };
}

function copyCorruptSeed(deps) {
  const bytes = deps.fs.readFileSync(deps.seedPath);
  const stem = `${deps.seedPath}.corrupt-${Math.trunc(deps.fs.statSync(deps.seedPath).mtimeMs)}`;
  let passedOver = false;
  for (let suffix = 0; ; suffix += 1) {
    const path = suffix === 0 ? stem : `${stem}-${suffix}`;
    try {
      deps.fs.copyFileSync(deps.seedPath, path, constants.COPYFILE_EXCL);
      return { path, stem };
    } catch (error) {
      if (error?.code !== "EEXIST") throw error;
    }
    const taken = readTakenCopy(deps, path);
    if (taken.bytes?.equals(bytes)) return { path, stem };
    if (!passedOver && (taken.notFile || taken.code != null)) {
      logLine(deps, taken.notFile ? "copy-passed-over-not-file" : "copy-passed-over", taken.code);
      passedOver = true;
    }
  }
}

function readTakenCopy(deps, path) {
  try {
    if (!deps.fs.statSync(path).isFile()) return { notFile: true };
    return { bytes: deps.fs.readFileSync(path) };
  } catch (error) {
    return { code: faultCode(error) };
  }
}

function corruptCopyKind(name, prefix, stemName) {
  if (name === stemName) return "stem";
  return /^\d+(-\d+)?$/.test(name.slice(prefix.length)) ? "copy" : "unknown";
}

function pruneCorruptSeedCopies(deps, keep) {
  const dir = dirname(deps.seedPath);
  const prefix = `${basename(deps.seedPath)}.corrupt-`;
  const stemName = basename(keep.stem);
  let names;
  try {
    names = deps.fs.readdirSync(dir);
  } catch (error) {
    logLine(deps, "copies-not-pruned", faultCode(error));
    return;
  }
  const older = [];
  for (const name of names) {
    const path = join(dir, name);
    if (!name.startsWith(prefix) || path === keep.path) continue;
    const kind = corruptCopyKind(name, prefix, stemName);
    let stat;
    try {
      stat = deps.fs.statSync(path);
    } catch (error) {
      logLine(deps, `prune-skipped-${kind}`, faultCode(error));
      continue;
    }
    if (stat.isFile()) older.push({ path, kind, mtimeMs: stat.mtimeMs });
    else logLine(deps, `prune-skipped-${kind}-not-file`);
  }
  older.sort((a, b) => b.mtimeMs - a.mtimeMs || b.path.localeCompare(a.path));
  for (const { path, kind } of older.slice(CORRUPT_SEED_COPY_LIMIT - 1)) {
    try {
      deps.fs.rmSync(path, { force: true });
    } catch (error) {
      logLine(deps, `prune-skipped-${kind}`, faultCode(error));
    }
  }
}

function writeSeedFile(deps, text) {
  const path = deps.seedPath;
  const tmp = `${path}.tmp`;
  seedStep("write", () => {
    deps.fs.mkdirSync(dirname(path), { recursive: true });
    deps.fs.writeFileSync(tmp, text);
  });
  seedStep("rename", () => deps.fs.renameSync(tmp, path));
}

export async function restoreFromSeed(deps, onlyPort, seen = new Map()) {
  const seed = readSeed(deps.fs, deps.seedPath);
  if (seed.kind === "unreadable") {
    return { injected: 0, missingAfter: 0, seedCount: 0, unreadable: seed.code };
  }
  if (seed.kind === "corrupt" || seed.kind === "unsupported-version") {
    return { injected: 0, missingAfter: 0, seedCount: 0, corrupt: true };
  }
  const seedCookies = seed.kind === "ok" ? seed.cookies : [];
  if (seedCookies.length === 0) return { injected: 0, missingAfter: 0, seedCount: 0 };

  const saved = new Map(seedCookies.map((c) => [cookieKey(c), c]));
  const faults = {};
  const reportRefused = persistCdpReporters(deps.log).onRefused;
  const onRefused = (browser, refused, total, error) => {
    faults.code = faultCode(error);
    faults.refusal = { refused, total };
    reportRefused(browser, refused, total, error);
  };
  const open = await openJars(deps, monitorPorts(deps), faults);
  let injected = 0;
  let missingAfter = seedCookies.length;
  try {
    for (const { port, browser, jar } of open) {
      if (onlyPort !== undefined && port !== onlyPort) continue;
      const others = open.filter((o) => o.port !== port);
      const missing = restoreCandidates(jar, others, saved, seen);
      try {
        if (missing.length > 0) {
          await pushCookies(browser, missing.map(toCookieParam), onRefused);
          injected += missing.length;
        }
        const after = await readCookies(browser);
        const stillMissing = [...saved.keys()].filter((key) => !after.has(key)).length;
        missingAfter = Math.min(missingAfter, stillMissing);
      } catch (error) {
        faults.code = faultCode(error);
        logLine(deps, "inject-failed", faults.code);
      }
    }
  } finally {
    for (const { browser } of open) browser.close();
  }
  return { injected, missingAfter, seedCount: seedCookies.length, ...faults };
}

function logRestoreOutcome(deps, outcome, { injected, seedCount, code, refusal }) {
  if (outcome === "ok") logLine(deps, injected > 0 ? "restore-injected" : "restore-present");
  else if (outcome === "empty") logLine(deps, "restore-empty");
  else if (seedCount === 0) logLine(deps, "restore-seed-unusable");
  else logLine(deps, "restore-incomplete", code ?? null, refusal ?? null);
}

async function isAnyChromeUp(deps) {
  for (const port of deps.discoverPorts()) {
    let browser;
    try {
      browser = await deps.connect(port);
    } catch (error) {
      if (!isConnectionRefused(error)) logLine(deps, "cdp-probe-failed", faultCode(error));
      continue;
    }
    browser.close();
    return true;
  }
  return false;
}

function restoringState(state, browserId, onlyPrimary) {
  return {
    ...state,
    phase: "restoring",
    browserId,
    onlyPrimary,
    restoreAttempts: 0,
    capture: { ...state.capture, digest: undefined },
  };
}

export function initialPersistState() {
  return {
    phase: "restoring",
    browserId: null,
    onlyPrimary: false,
    seen: new Map(),
    restoreAttempts: 0,
    restoreRowKey: null,
    capture: {},
  };
}

export async function persistTick(deps, state) {
  const primary = await connectPrimary(deps);
  try {
    const browserId = primary?.browserId ?? null;
    if (browserId != null && browserId !== state.browserId) {
      if (state.browserId != null) logLine(deps, "primary-replaced");
      const bootRestore = state.phase === "restoring" && state.browserId == null;
      return await restorePass(deps, restoringState(state, browserId, !bootRestore));
    }
    if (state.phase === "restoring") {
      if (primary == null && !(await isAnyChromeUp(deps))) return state;
      return await restorePass(deps, state);
    }
    return await captureTick(deps, state, primary);
  } finally {
    primary?.close();
  }
}

export async function runPersistTick(deps, state) {
  try {
    return await persistTick(deps, state);
  } catch (error) {
    logLine(deps, "tick-failed", faultCode(error));
    return state;
  }
}

async function restorePass(deps, state) {
  const attempt = await restoreFromSeed(
    deps,
    state.onlyPrimary ? PRIMARY_CDP_PORT : undefined,
    state.seen,
  );
  const { injected, missingAfter, seedCount, corrupt, unreadable } = attempt;
  if (unreadable != null) return settleWaiting(deps, state, unreadable);
  const restoreAttempts = state.restoreAttempts + 1;
  if (missingAfter > 0 && restoreAttempts < MAX_RESTORE_ATTEMPTS) {
    return { ...state, restoreAttempts, restoreRowKey: "attempt" };
  }
  const outcome = classifyRestoreOutcome(seedCount, missingAfter, corrupt);
  logRestoreOutcome(deps, outcome, attempt);
  deps.emitTelemetry?.({
    kind: "cookie_persist",
    phase: "restore",
    outcome,
    seedCookies: seedCount,
    injected,
    missingAfter,
    attempts: restoreAttempts,
  });
  return { ...state, phase: "guarding", restoreAttempts, restoreRowKey: "attempt" };
}

function settleWaiting(deps, state, code) {
  const reason = classifySeedFault(code);
  const rowKey = `waiting:${reason}`;
  const at = deps.now?.() ?? Date.now();
  const ships = rowKey !== state.restoreRowKey;
  const due = at - state.restoreLineAt >= HEARTBEAT_MS;
  if (ships || due) logLine(deps, "restore-waiting", code);
  if (ships) {
    deps.emitTelemetry?.({ kind: "cookie_persist", phase: "restore", outcome: "waiting", reason });
  }
  return {
    ...state,
    restoreRowKey: rowKey,
    restoreLineAt: ships || due ? at : state.restoreLineAt,
  };
}

async function main() {
  const log = createPersistLog();
  exitOnUncaught(log);
  if (!dropToBoxUser(log)) return;
  if (!isCookiePersistEnabled()) {
    log("store-off");
    return;
  }
  log("starting");
  const report = persistCdpReporters(log);
  const deps = {
    connect: (port) => connectBrowser(port, report.onDroppedFrame),
    discoverPorts: () => discoverMonitorPorts(undefined, report.onUnreadable),
    seedPath: SEED_PATH,
    fs: {
      copyFileSync,
      mkdirSync,
      readdirSync,
      readFileSync,
      renameSync,
      rmdirSync,
      rmSync,
      statSync,
      unlinkSync,
      writeFileSync,
    },
    log,
    emitTelemetry: (event) => appendTelemetry(log, event),
  };

  let state = initialPersistState();
  for (;;) {
    state = await runPersistTick(deps, state);
    await sleep(CAPTURE_INTERVAL_MS);
  }
}

if (process.argv[1] != null && import.meta.url === pathToFileURL(process.argv[1]).href) {
  void main();
}
