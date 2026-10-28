import { createHash } from "node:crypto";
import {
  appendFileSync,
  closeSync,
  constants,
  fstatSync,
  mkdtempSync,
  openSync,
  readdirSync,
  readFileSync,
  renameSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

import {
  BOX_BROWSER_ACCEPT_LANGUAGE_PATH_PREFIX,
  BOX_BROWSER_GPU_MODE_PATH_PREFIX,
  BOX_BROWSER_RUNNING_PATH_PREFIX,
  BOX_EXPERIMENTS_READY_MARKER_PATH,
  BROWSER_ACCEPT_LANGUAGE_EN_US_MARKER_PATH,
  BOX_BROWSER_COHERENT_FINGERPRINT_MARKER_PATH,
  BOX_BROWSER_FINGERPRINT_LOG_SITES_PATH,
  SAND_BOX_CDP_PORT_BASE,
  BROWSER_EGRESS_TIMEZONE_MARKER_PATH,
  BROWSER_HARDWARE_PROFILE_MARKER_PATH,
} from "./box-contract.generated.mjs";
import {
  CDP_ERROR_CODES,
  connectBrowser,
  discoverMonitorPorts,
  getBrowserVersion,
  isCdpTargetGone,
  isConnectionRefused,
} from "./cdp-cookies.mjs";
import {
  ACCEPT_LANGUAGE_EXTENSION_NAME,
  LINUX_GPU_PATCH_PROFILE,
  buildAcceptLanguageExtension,
  buildHardwareProfileDocumentScript,
  buildUserAgentOverride,
  coherentHardwareProfile,
  configuredAcceptLanguage,
  normalizeGpuMode,
  resolveCanvasSeed,
  verifyCoherentWebgl,
} from "./sand-fingerprint-profiles.mjs";
import {
  egressErrorClass,
  egressErrorCode,
  isBoxClockUtc,
  resolveEgressTimeZone,
} from "./sand-egress-timezone.mjs";

const POLL_INTERVAL_MS = 100;
const FINGERPRINT_MARKER_MAX_BYTES = 8 * 1024;
const FINGERPRINT_MARKER_MAX_HOSTS = 64;
const FINGERPRINT_MARKER_MAX_HOST_LENGTH = 253;
const FINGERPRINT_MARKER_HOST_LABEL = /^(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?)$/;
export const EGRESS_TIME_ZONE_ATTACH_WAIT_MS = 3_000;
export const EGRESS_TIME_ZONE_RETRY_MS = Object.freeze([5_000, 30_000, 120_000, 300_000]);
export const EGRESS_TIME_ZONE_GATE_WAIT_MS = 120_000;

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function errorClass(error) {
  return error?.code ?? error?.cause?.code ?? (error instanceof Error ? error.name : typeof error);
}

const reportedOnce = new Set();

function reportOnce(message) {
  if (reportedOnce.has(message)) return;
  reportedOnce.add(message);
  console.error(message);
}

export function readFingerprintMarkerCandidate({
  markerPath = BOX_BROWSER_FINGERPRINT_LOG_SITES_PATH,
  open = openSync,
  stat = fstatSync,
  readFile = readFileSync,
  effectiveUid = process.geteuid?.(),
  effectiveGroups = [process.getegid?.(), ...(process.getgroups?.() ?? [])].filter(
    Number.isInteger,
  ),
  uidSplitActive = process.env.SAND_HOST_UID_SPLIT_ACTIVE === "1",
} = {}) {
  let fd;
  try {
    fd = open(markerPath, constants.O_RDONLY | constants.O_NOFOLLOW);
    const metadata = stat(fd);
    const mode = metadata.mode & 0o777;
    const ownerPrivate =
      mode === 0o600 && (effectiveUid === undefined || metadata.uid === effectiveUid);
    const ownerShared =
      mode === 0o640 &&
      effectiveUid !== undefined &&
      metadata.uid !== effectiveUid &&
      effectiveGroups.includes(metadata.gid);
    if (
      !metadata.isFile() ||
      metadata.nlink !== 1 ||
      metadata.size <= 0 ||
      metadata.size > FINGERPRINT_MARKER_MAX_BYTES ||
      (uidSplitActive ? !ownerShared : !ownerPrivate)
    ) {
      return null;
    }
    return String(readFile(fd, "utf8"));
  } catch (error) {
    if (error?.code !== "ENOENT") {
      reportOnce(`fingerprint log marker read failed: ${errorClass(error)}`);
    }
    return null;
  } finally {
    if (fd !== undefined) closeSync(fd);
  }
}

function fingerprintMarkerHost(raw) {
  if (
    typeof raw !== "string" ||
    raw.length === 0 ||
    raw.length > FINGERPRINT_MARKER_MAX_HOST_LENGTH ||
    raw !== raw.trim() ||
    raw.includes("*") ||
    raw.includes("/") ||
    raw.includes("@")
  ) {
    return null;
  }
  let parsed;
  try {
    parsed = new URL(`https://${raw}`);
  } catch {
    return null;
  }
  if (
    parsed.username !== "" ||
    parsed.password !== "" ||
    parsed.pathname !== "/" ||
    parsed.search !== "" ||
    parsed.hash !== "" ||
    (parsed.port !== "" && parsed.port !== "443")
  ) {
    return null;
  }
  const host = parsed.hostname.toLowerCase().replace(/\.$/, "");
  const labels = host.split(".");
  if (
    host === "" ||
    host.includes(":") ||
    labels.length < 2 ||
    labels.some((label) => !FINGERPRINT_MARKER_HOST_LABEL.test(label))
  ) {
    return null;
  }
  return host;
}

export function parseFingerprintMarker(raw) {
  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return null;
  }
  if (
    parsed == null ||
    typeof parsed !== "object" ||
    Array.isArray(parsed) ||
    Object.keys(parsed).some((key) => key !== "version" && key !== "hosts") ||
    parsed.version !== 1 ||
    !Array.isArray(parsed.hosts) ||
    parsed.hosts.length === 0 ||
    parsed.hosts.length > FINGERPRINT_MARKER_MAX_HOSTS
  ) {
    return null;
  }
  const hosts = [];
  const seen = new Set();
  for (const rawHost of parsed.hosts) {
    const host = fingerprintMarkerHost(rawHost);
    if (host == null || host !== rawHost || seen.has(host)) return null;
    seen.add(host);
    hosts.push(host);
  }
  if (hosts.some((host, index) => index > 0 && hosts[index - 1] > host)) return null;
  return Object.freeze({ version: 1, hosts: Object.freeze(hosts) });
}

export function fingerprintMarkerDigest(raw) {
  return createHash("sha256").update(raw).digest("hex");
}

const FINGERPRINT_LOGGER_MODULE = "./sand-fingerprint-read-logger.mjs";

async function importFingerprintLogger() {
  return await import(FINGERPRINT_LOGGER_MODULE);
}

export async function configureFingerprintLogger(
  browser,
  {
    markerPath = BOX_BROWSER_FINGERPRINT_LOG_SITES_PATH,
    open = openSync,
    stat = fstatSync,
    readFile = readFileSync,
    load = importFingerprintLogger,
  } = {},
) {
  const raw = readFingerprintMarkerCandidate({
    markerPath,
    open,
    stat,
    readFile,
  });
  const config = raw == null ? null : parseFingerprintMarker(raw);
  browser.fingerprintLogConfigDigest = config == null ? null : fingerprintMarkerDigest(raw);
  if (config == null) return null;
  let loggerModule;
  try {
    loggerModule = await load();
  } catch (error) {
    if (error?.code !== "ERR_MODULE_NOT_FOUND") throw error;
    reportOnce(`fingerprint logger unavailable: ${errorClass(error)}`);
    return null;
  }
  browser.fingerprintReadLogger = loggerModule.createFingerprintReadLogger({
    browser,
    config,
    cdpPortBase: SAND_BOX_CDP_PORT_BASE,
  });
  return browser.fingerprintReadLogger;
}

function liveChromeProduct(browser) {
  const version = typeof browser?.chromeVersion === "string" ? browser.chromeVersion : "";
  if (/^\d+\.\d+\.\d+\.\d+$/.test(version)) return `Chrome/${version}`;
  return "Chrome/0.0.0.0";
}

export async function applyDesktopUaToTarget(browser, sessionId, browserContextId) {
  browser.acceptLanguageOverriddenSessions ??= new Set();
  const override = buildUserAgentOverride(liveChromeProduct(browser), {
    acceptLanguageEnUs: acceptLanguageEnUsEnabledForBrowser(browser),
    acceptLanguage: browser.acceptLanguage,
    acceptLanguageFromChrome: acceptLanguageFromLaunchFlag(browser, browserContextId),
    acceptLanguageOverridden: browser.acceptLanguageOverriddenSessions.has(sessionId),
  });
  if (override.acceptLanguage) {
    browser.acceptLanguageOverriddenSessions.add(sessionId);
    if (
      isOffTheRecordContext(browser, browserContextId) &&
      !acceptLanguageFromExtension(browser, browserContextId)
    ) {
      reportAcceptLanguageFallback(browser);
    }
  }
  try {
    await sendUserAgentOverride(browser, override, sessionId);
  } catch (error) {
    if (override.acceptLanguage && !isCdpTargetGone(error) && !browser.isClosed) {
      browser.acceptLanguageUnappliedSessions ??= new Set();
      browser.acceptLanguageUnappliedSessions.add(sessionId);
    }
    throw error;
  }
}

export const UA_OVERRIDE_RETRY_MS = Object.freeze([250, 1_000]);
const RETRIED_UA_OVERRIDE_ERRORS = new Set([
  CDP_ERROR_CODES.sendTimeout,
  CDP_ERROR_CODES.protocolServer,
]);

async function sendUserAgentOverride(browser, override, sessionId) {
  for (let attempt = 0; ; attempt += 1) {
    try {
      return await browser.send("Emulation.setUserAgentOverride", override, sessionId);
    } catch (error) {
      const code = errorClass(error);
      const retryMs = UA_OVERRIDE_RETRY_MS[attempt];
      const live = !browser.isClosed && browser.attachedSessions?.has(sessionId) !== false;
      if (
        !acceptLanguageEnUsEnabledForBrowser(browser) ||
        !RETRIED_UA_OVERRIDE_ERRORS.has(code) ||
        !live
      ) {
        throw error;
      }
      if (retryMs === undefined) {
        console.error(
          `ua override failed port=${browser.port} after ${attempt + 1} attempts: ${code}; this target keeps Chrome's own User-Agent and Accept-Language`,
        );
        throw error;
      }
      await delay(retryMs);
    }
  }
}

export function documentScriptMap(browser) {
  if (browser.documentScripts == null) {
    browser.documentScripts = new Map();
  }
  return browser.documentScripts;
}

const WORKER_TARGET_TYPES = new Set(["worker", "shared_worker", "service_worker"]);

export function isWorkerTargetType(targetType) {
  return WORKER_TARGET_TYPES.has(targetType);
}

const MARKER_GATE_ON = new Set(["1", "on", "true"]);
const MARKER_CONTENTS_OFF = new Set(["0", "off", "false"]);

function resolveMarkerGate({ label, envValue, markerPath, readFile, honorContents = false }) {
  if (envValue != null && String(envValue).trim() !== "") {
    return MARKER_GATE_ON.has(String(envValue).trim().toLowerCase());
  }
  try {
    const contents = readFile(markerPath, "utf8");
    return !honorContents || !MARKER_CONTENTS_OFF.has(String(contents).trim().toLowerCase());
  } catch (error) {
    if (
      error == null ||
      typeof error !== "object" ||
      !("code" in error) ||
      error.code !== "ENOENT"
    ) {
      reportOnce(
        `${label} marker read failed: ${error instanceof Error ? error.name : typeof error}`,
      );
    }
    return false;
  }
}

export function resolveHardwareProfileEnabled({
  envValue = process.env.SAND_BROWSER_HARDWARE_PROFILE,
  markerPath = BROWSER_HARDWARE_PROFILE_MARKER_PATH,
  readFile = readFileSync,
} = {}) {
  return resolveMarkerGate({
    label: "hardware profile",
    envValue,
    markerPath,
    readFile,
    honorContents: true,
  });
}

export function resolveCoherentHardwareProfileEnabled({
  envValue = process.env.SAND_BOX_BROWSER_COHERENT_FINGERPRINT,
  markerPath = BOX_BROWSER_COHERENT_FINGERPRINT_MARKER_PATH,
  readFile = readFileSync,
} = {}) {
  return resolveMarkerGate({
    label: "coherent hardware profile",
    envValue,
    markerPath,
    readFile,
    honorContents: true,
  });
}

export function resolveEgressTimeZoneGate({
  envValue = process.env.SAND_BROWSER_EGRESS_TIMEZONE,
  markerPath = BROWSER_EGRESS_TIMEZONE_MARKER_PATH,
  readyMarkerPath = BOX_EXPERIMENTS_READY_MARKER_PATH,
  readFile = readFileSync,
} = {}) {
  if (envValue != null && String(envValue).trim() !== "") {
    return MARKER_GATE_ON.has(String(envValue).trim().toLowerCase()) ? "on" : "off";
  }
  if (resolveMarkerGate({ label: "egress time zone", markerPath, readFile })) return "on";
  const ready = resolveMarkerGate({
    label: "experiments ready",
    markerPath: readyMarkerPath,
    readFile,
  });
  return ready ? "off" : "undecided";
}

export function hardwareProfileEnabledForBrowser(browser, resolve = resolveHardwareProfileEnabled) {
  if (!Object.prototype.hasOwnProperty.call(browser, "hardwareProfileEnabled")) {
    browser.hardwareProfileEnabled = resolve() === true;
  }
  return browser.hardwareProfileEnabled === true;
}

export function coherentHardwareProfileEnabledForBrowser(
  browser,
  resolve = resolveCoherentHardwareProfileEnabled,
) {
  if (!Object.prototype.hasOwnProperty.call(browser, "coherentHardwareProfileEnabled")) {
    browser.coherentHardwareProfileEnabled = resolve() === true;
  }
  return browser.coherentHardwareProfileEnabled === true;
}

export function hardwareTreatmentForBrowser(browser) {
  if (coherentHardwareProfileEnabledForBrowser(browser)) return "coherent";
  return hardwareProfileEnabledForBrowser(browser) ? "legacy" : "off";
}

export function resolveBoxBrowserGpuMode({
  port,
  gpuModePath = `${BOX_BROWSER_GPU_MODE_PATH_PREFIX}-${port}`,
  readFile = readFileSync,
} = {}) {
  try {
    return normalizeGpuMode(readFile(gpuModePath, "utf8"));
  } catch (error) {
    if (error?.code !== "ENOENT") {
      reportOnce(`box browser gpu mode read failed: ${errorClass(error)}`);
    }
    return "unknown";
  }
}

export function coherentProfileForBrowser(browser, resolve = resolveBoxBrowserGpuMode) {
  if (browser.coherentProfile == null) {
    browser.coherentProfile = coherentHardwareProfile(resolve({ port: browser.port }));
  }
  return browser.coherentProfile;
}

export async function verifyCoherentProfile(browser) {
  const profile = coherentProfileForBrowser(browser);
  let gpu;
  try {
    gpu = (await browser.send("SystemInfo.getInfo"))?.gpu;
  } catch (error) {
    if (isCdpTargetGone(error) || browser.isClosed) return null;
    console.error(
      `coherent hardware profile check failed port=${browser.port}: ${errorClass(error)}`,
    );
    return null;
  }
  const check = verifyCoherentWebgl(profile, gpu);
  console.error(
    `coherent hardware profile port=${browser.port} gpu-mode=${profile.gpuMode} webgl=${check.claimed} observed=${check.observed} ${check.verdict}`,
  );
  return check;
}

function applyCoherentHardwareToTarget(browser, sessionId) {
  return browser.send(
    "Emulation.setHardwareConcurrencyOverride",
    { hardwareConcurrency: coherentProfileForBrowser(browser).hardwareConcurrency },
    sessionId,
  );
}

export function readAcceptLanguageEnUsMarker(
  markerPath = BROWSER_ACCEPT_LANGUAGE_EN_US_MARKER_PATH,
  unreadable = "",
) {
  let fd;
  try {
    fd = openSync(markerPath, constants.O_RDONLY | constants.O_NOFOLLOW);
    return readFileSync(fd, "utf8");
  } catch (error) {
    if (error?.code === "ENOENT") return null;
    reportOnce(`accept-language marker read failed: ${errorClass(error)}`);
    return unreadable;
  } finally {
    if (fd !== undefined) closeSync(fd);
  }
}

export function acceptLanguageEnUsEnabledForBrowser(
  browser,
  resolve = readAcceptLanguageEnUsMarker,
) {
  if (!Object.prototype.hasOwnProperty.call(browser, "acceptLanguageEnUsEnabled")) {
    const marker = resolve();
    browser.acceptLanguageEnUsEnabled = marker != null;
    browser.acceptLanguage = marker ?? "";
    if (marker != null) {
      if (browser.launchAcceptLanguage === undefined) {
        browser.launchAcceptLanguage = readLaunchAcceptLanguage(browser.port);
      }
      if (browser.launchAcceptLanguage) browser.acceptLanguage = browser.launchAcceptLanguage;
    }
  }
  return browser.acceptLanguageEnUsEnabled === true;
}

export const ACCEPT_LANG_FLAG = "--accept-lang=";

function readProcFile(path) {
  try {
    return String(readFileSync(path, "utf8"));
  } catch (error) {
    if (error?.code !== "ENOENT" && error?.code !== "ESRCH") {
      reportOnce(`accept-language launch flag read failed: ${errorClass(error)}`);
    }
    return null;
  }
}

function processStartTicks(procRoot, pid) {
  const stat = readProcFile(`${procRoot}/${pid}/stat`);
  const ticks = Number(stat?.slice(stat.lastIndexOf(")") + 2).split(" ")[19]);
  return Number.isSafeInteger(ticks) ? ticks : Number.POSITIVE_INFINITY;
}

export function readLaunchAcceptLanguage(port, procRoot = "/proc") {
  let entries;
  try {
    entries = readdirSync(procRoot);
  } catch (error) {
    reportOnce(`accept-language launch flag read failed: ${errorClass(error)}`);
    return null;
  }
  const portFlag = `--remote-debugging-port=${port}`;
  let browser = null;
  for (const entry of entries) {
    if (!/^\d+$/.test(entry)) continue;
    const args = readProcFile(`${procRoot}/${entry}/cmdline`)?.split("\0");
    if (args === undefined || !args.includes(portFlag)) continue;
    if (args.some((arg) => arg.startsWith("--type="))) continue;
    const startTicks = processStartTicks(procRoot, entry);
    if (browser === null || startTicks < browser.startTicks) browser = { args, startTicks };
  }
  const flag = browser?.args.findLast((arg) => arg.startsWith(ACCEPT_LANG_FLAG));
  return flag === undefined ? null : flag.slice(ACCEPT_LANG_FLAG.length);
}

function launchFlagMatches(browser) {
  if (!acceptLanguageEnUsEnabledForBrowser(browser)) return false;
  if (browser.launchAcceptLanguage === undefined) {
    browser.launchAcceptLanguage = readLaunchAcceptLanguage(browser.port);
  }
  return browser.launchAcceptLanguage === configuredAcceptLanguage(browser.acceptLanguage);
}

export function acceptLanguageFromLaunchFlag(browser, browserContextId) {
  return (
    typeof browserContextId === "string" &&
    browserContextId === browser.defaultBrowserContextId &&
    launchFlagMatches(browser)
  );
}

async function resolveDefaultBrowserContext(browser) {
  if (
    typeof browser.defaultBrowserContextId === "string" ||
    !acceptLanguageEnUsEnabledForBrowser(browser)
  ) {
    return;
  }
  try {
    const { defaultBrowserContextId } = (await browser.send("Target.getBrowserContexts")) ?? {};
    if (typeof defaultBrowserContextId === "string") {
      browser.defaultBrowserContextId = defaultBrowserContextId;
    }
  } catch (error) {
    if (!browser.isClosed) {
      console.error(`browser contexts read failed port=${browser.port}: ${errorClass(error)}`);
    }
  }
}

export const ACCEPT_LANGUAGE_EXTENSION_ROOT = "/tmp";
const ACCEPT_LANGUAGE_EXTENSION_DIR_PREFIX = "sand-ua-governor-accept-language-";
export const ACCEPT_LANGUAGE_EXTENSION_HOLD_MS = 3_000;
export const ACCEPT_LANGUAGE_EXTENSION_LOAD_TIMEOUT_MS = 20_000;
export const ACCEPT_LANGUAGE_EXTENSION_RETRY_MS = Object.freeze([5_000, 30_000, 120_000, 300_000]);

export function isOffTheRecordContext(browser, browserContextId) {
  return (
    acceptLanguageEnUsEnabledForBrowser(browser) &&
    typeof browser.defaultBrowserContextId === "string" &&
    typeof browserContextId === "string" &&
    browserContextId !== browser.defaultBrowserContextId
  );
}

function extensionCoversList(browser) {
  const extension = browser.acceptLanguageExtension;
  return (
    extension?.phase === "ready" &&
    extension.list === configuredAcceptLanguage(browser.acceptLanguage)
  );
}

function extensionCoversOffTheRecord(browser) {
  const extension = browser.acceptLanguageExtension;
  return (
    extensionCoversList(browser) || (extension?.phase === "loading" && extension.installed != null)
  );
}

export function acceptLanguageFromExtension(browser, browserContextId) {
  return isOffTheRecordContext(browser, browserContextId) && extensionCoversOffTheRecord(browser);
}

function reportAcceptLanguageFallback(browser) {
  const extension = browser.acceptLanguageExtension;
  const reason =
    extension?.phase === "loading"
      ? `extension still loading after ${ACCEPT_LANGUAGE_EXTENSION_HOLD_MS}ms`
      : (extension?.reason ?? `extension loaded for ${extension?.list}`);
  reportOnce(
    `accept-language fallback port=${browser.port} context=off-the-record: the CDP override sends ${configuredAcceptLanguage(browser.acceptLanguage)} right after User-Agent (${reason})`,
  );
}

export async function readExtensionsCommands(port, signal) {
  const res = await fetch(`http://127.0.0.1:${port}/json/protocol`, { signal });
  if (!res.ok) throw new Error(`/json/protocol HTTP ${res.status}`);
  const domain = (await res.json())?.domains?.find((entry) => entry?.domain === "Extensions");
  return new Map(
    (domain?.commands ?? []).map((command) => [
      command?.name,
      (command?.parameters ?? []).map((parameter) => parameter?.name),
    ]),
  );
}

export function extensionsUnsupported(commands) {
  if (!commands.get("loadUnpacked")?.includes("enableInIncognito")) {
    return "Chrome lacks Extensions.loadUnpacked enableInIncognito";
  }
  const missing = ["getExtensions", "uninstall"].find((name) => !commands.has(name));
  return missing === undefined ? null : `Chrome lacks Extensions.${missing}`;
}

function acceptLanguageExtensionRoot(browser) {
  return browser.acceptLanguageExtensionRoot ?? ACCEPT_LANGUAGE_EXTENSION_ROOT;
}

function acceptLanguageExtensionDirPrefix(browser) {
  return join(
    acceptLanguageExtensionRoot(browser),
    `${ACCEPT_LANGUAGE_EXTENSION_DIR_PREFIX}${browser.port}-`,
  );
}

async function uninstallAcceptLanguageExtension(browser, id) {
  try {
    await browser.send("Extensions.uninstall", { id });
    return true;
  } catch (error) {
    if (!browser.isClosed) {
      console.error(
        `accept-language extension uninstall failed port=${browser.port} id=${id}: ${errorClass(error)}`,
      );
    }
    return false;
  }
}

const acceptLanguageExtensionPortsListed = new Set();
const acceptLanguageExtensionPortsLoaded = new Set();

function acceptLanguageExtensionDirInUse(browser, path) {
  const extension = browser.acceptLanguageExtension;
  return (
    path === extension?.installed?.dir ||
    path === extension?.attempt?.dir ||
    extension?.loadingDirs?.has(path) === true
  );
}

async function removeStaleAcceptLanguageExtensions(browser) {
  const prefix = acceptLanguageExtensionDirPrefix(browser);
  const kept = new Set();
  try {
    const { extensions } = (await browser.send("Extensions.getExtensions")) ?? {};
    for (const extension of Array.isArray(extensions) ? extensions : []) {
      if (
        extension?.name !== ACCEPT_LANGUAGE_EXTENSION_NAME ||
        typeof extension.path !== "string" ||
        !extension.path.startsWith(prefix)
      ) {
        continue;
      }
      if (
        extension.id === browser.acceptLanguageExtension?.installed?.id ||
        acceptLanguageExtensionDirInUse(browser, extension.path) ||
        !(await uninstallAcceptLanguageExtension(browser, extension.id))
      ) {
        kept.add(extension.path);
      }
    }
    for (const name of readdirSync(dirname(prefix))) {
      const path = join(dirname(prefix), name);
      if (
        path.startsWith(prefix) &&
        !kept.has(path) &&
        !acceptLanguageExtensionDirInUse(browser, path)
      ) {
        rmSync(path, { recursive: true, force: true });
      }
    }
    return kept.size === 0;
  } catch (error) {
    if (!browser.isClosed) {
      console.error(
        `accept-language extension cleanup failed port=${browser.port}: ${errorClass(error)}`,
      );
    }
    return false;
  }
}

async function installAcceptLanguageExtension(browser, extension, attempt) {
  const read = browser.readExtensionsCommands ?? readExtensionsCommands;
  const unsupported = extensionsUnsupported(await read(browser.port, attempt.abort.signal));
  if (unsupported !== null) return { unsupported };
  const dir = mkdtempSync(acceptLanguageExtensionDirPrefix(browser));
  attempt.dir = dir;
  extension.loadingDirs ??= new Set();
  extension.loadingDirs.add(dir);
  let loaded;
  try {
    for (const [name, contents] of Object.entries(buildAcceptLanguageExtension(attempt.list))) {
      writeFileSync(join(dir, name), contents, { mode: 0o644 });
    }
    loaded = await browser.send("Extensions.loadUnpacked", { path: dir, enableInIncognito: true });
    if (typeof loaded?.id !== "string") throw new Error("Extensions.loadUnpacked returned no id");
  } catch (error) {
    rmSync(dir, { recursive: true, force: true });
    throw error;
  } finally {
    extension.loadingDirs.delete(dir);
  }
  if (extension.attempt !== attempt) {
    if (await uninstallAcceptLanguageExtension(browser, loaded.id)) {
      rmSync(dir, { recursive: true, force: true });
    }
    return null;
  }
  return { installed: { id: loaded.id, dir } };
}

function settleAcceptLanguageExtension(browser, extension, attempt, outcome, now) {
  if (extension.attempt !== attempt || browser.isClosed) return;
  if (outcome.installed !== undefined) {
    extension.phase = "ready";
    extension.installed = outcome.installed;
    acceptLanguageExtensionPortsLoaded.add(browser.port);
    extension.reason = null;
    extension.attempts = 0;
    console.error(
      `accept-language extension loaded port=${browser.port} id=${outcome.installed.id}: ${attempt.list} in stock position for off-the-record contexts`,
    );
    extension.cleanup = removeStaleAcceptLanguageExtensions(browser);
    return;
  }
  extension.attempt = null;
  extension.installed = null;
  if (outcome.unsupported !== undefined) {
    extension.phase = "unsupported";
    extension.reason = outcome.unsupported;
    console.error(
      `accept-language extension unavailable port=${browser.port}: ${outcome.unsupported}`,
    );
    return;
  }
  const retryMs =
    ACCEPT_LANGUAGE_EXTENSION_RETRY_MS[
      Math.min(extension.attempts, ACCEPT_LANGUAGE_EXTENSION_RETRY_MS.length - 1)
    ];
  extension.phase = "failed";
  extension.reason = `extension load failed ${errorClass(outcome.error)}`;
  extension.attempts += 1;
  extension.nextAttemptAt = now() + retryMs;
  console.error(
    `accept-language extension load failed port=${browser.port}: ${errorClass(outcome.error)}; retrying in ${retryMs}ms`,
  );
  extension.cleanup = removeStaleAcceptLanguageExtensions(browser);
}

function reapplyOffTheRecordAcceptLanguage(browser) {
  return Promise.all(
    [...(browser.attachedSessions ?? [])].map((sessionId) => {
      const browserContextId = browser.attachedBrowserContexts?.get(sessionId);
      if (!isOffTheRecordContext(browser, browserContextId)) return null;
      return applyDesktopUaToTarget(browser, sessionId, browserContextId).catch((error) => {
        if (isCdpTargetGone(error) || !browser.attachedSessions.has(sessionId)) return;
        console.error(`accept-language re-apply failed port=${browser.port}: ${errorClass(error)}`);
      });
    }),
  );
}

function finishAcceptLanguageExtension(browser, extension, attempt, outcome, now) {
  if (extension.attempt !== attempt) return null;
  const covered = extensionCoversOffTheRecord(browser);
  settleAcceptLanguageExtension(browser, extension, attempt, outcome, now);
  extension.inFlight = null;
  attempt.releaseHold();
  if (browser.isClosed) return null;
  if (extension.phase === "ready" && !extensionCoversList(browser)) {
    void ensureAcceptLanguageExtension(browser, { now });
    if (extensionCoversOffTheRecord(browser) !== covered) {
      void reapplyOffTheRecordAcceptLanguage(browser);
    }
    return extension.inFlight;
  }
  if (extensionCoversOffTheRecord(browser) === covered) return null;
  return reapplyOffTheRecordAcceptLanguage(browser);
}

export function ensureAcceptLanguageExtension(browser, { now = Date.now } = {}) {
  browser.acceptLanguageExtension ??= { phase: "idle", attempts: 0, nextAttemptAt: 0 };
  const extension = browser.acceptLanguageExtension;
  if (extension.inFlight != null) return extension.hold;
  if (browser.isClosed || extension.phase === "unsupported") return null;
  if (extensionCoversList(browser)) return verifyAcceptLanguageExtension(browser, extension, now);
  if (extension.phase === "failed" && now() < extension.nextAttemptAt) return null;
  const attempt = {
    list: configuredAcceptLanguage(browser.acceptLanguage),
    abort: new AbortController(),
    holdUntil: now() + ACCEPT_LANGUAGE_EXTENSION_HOLD_MS,
    deadline: now() + ACCEPT_LANGUAGE_EXTENSION_LOAD_TIMEOUT_MS,
  };
  extension.attempt = attempt;
  extension.phase = "loading";
  extension.list = attempt.list;
  extension.hold = new Promise((resolve) => {
    attempt.releaseHold = resolve;
  });
  extension.inFlight = installAcceptLanguageExtension(browser, extension, attempt)
    .then(
      (outcome) => outcome ?? { error: new Error("superseded") },
      (error) => ({ error }),
    )
    .then((outcome) => finishAcceptLanguageExtension(browser, extension, attempt, outcome, now));
  return extension.hold;
}

async function acceptLanguageExtensionDropped(browser, installed) {
  try {
    const { extensions } = (await browser.send("Extensions.getExtensions")) ?? {};
    return !(Array.isArray(extensions) ? extensions : []).some(
      (entry) => entry?.id === installed.id && entry.enabled !== false,
    );
  } catch (error) {
    if (!browser.isClosed) {
      reportOnce(
        `accept-language extension check failed port=${browser.port}: ${errorClass(error)}`,
      );
    }
    return false;
  }
}

async function confirmAcceptLanguageExtension(browser, extension, verify, now) {
  const { installed } = extension;
  const dropped = await acceptLanguageExtensionDropped(browser, installed);
  if (extension.verify === verify) extension.verify = null;
  if (!dropped || extension.installed !== installed || browser.isClosed) return verify.release();
  console.error(
    `accept-language extension ${installed.id} is no longer loaded in Chrome port=${browser.port}; loading it again`,
  );
  extension.phase = "failed";
  extension.reason = "extension no longer loaded in Chrome";
  extension.installed = null;
  extension.nextAttemptAt = 0;
  verify.release(ensureAcceptLanguageExtension(browser, { now }));
  return reapplyOffTheRecordAcceptLanguage(browser);
}

function verifyAcceptLanguageExtension(browser, extension, now) {
  if (extension.verify != null) return extension.verify.hold;
  const verify = { holdUntil: now() + ACCEPT_LANGUAGE_EXTENSION_HOLD_MS };
  verify.hold = new Promise((resolve) => {
    verify.release = resolve;
  });
  extension.verify = verify;
  void confirmAcceptLanguageExtension(browser, extension, verify, now);
  return verify.hold;
}

function acceptLanguageExtensionWanted(browser) {
  return (
    acceptLanguageEnUsEnabledForBrowser(browser) &&
    typeof browser.defaultBrowserContextId === "string"
  );
}

function prepareAcceptLanguageExtension(browser) {
  if (!acceptLanguageExtensionWanted(browser)) return null;
  const hold = ensureAcceptLanguageExtension(browser);
  if (browser.acceptLanguageExtension?.verify != null) return hold;
  return extensionCoversOffTheRecord(browser) ? null : hold;
}

function hasLeftoverAcceptLanguageExtensions(browser) {
  const prefix = acceptLanguageExtensionDirPrefix(browser);
  try {
    return readdirSync(dirname(prefix)).some((name) =>
      join(dirname(prefix), name).startsWith(prefix),
    );
  } catch (error) {
    reportOnce(
      `accept-language extension cleanup failed port=${browser.port}: ${errorClass(error)}`,
    );
    return null;
  }
}

export function settleAcceptLanguageExtensionOnAttach(browser) {
  if (acceptLanguageExtensionWanted(browser)) return ensureAcceptLanguageExtension(browser);
  const { port } = browser;
  const settled = () => {
    acceptLanguageExtensionPortsListed.add(port);
    acceptLanguageExtensionPortsLoaded.delete(port);
  };
  if (
    !acceptLanguageEnUsEnabledForBrowser(browser) &&
    acceptLanguageExtensionPortsListed.has(port) &&
    !acceptLanguageExtensionPortsLoaded.has(port)
  ) {
    return null;
  }
  const leftover = hasLeftoverAcceptLanguageExtensions(browser);
  if (leftover === false) settled();
  if (leftover !== true) return null;
  browser.acceptLanguageExtensionCleanup = removeStaleAcceptLanguageExtensions(browser).then(
    (removed) => {
      if (removed) settled();
    },
  );
  return browser.acceptLanguageExtensionCleanup;
}

export function acceptLanguageExtensionHold(browser, browserContextId) {
  return isOffTheRecordContext(browser, browserContextId)
    ? prepareAcceptLanguageExtension(browser)
    : null;
}

export function tickAcceptLanguageExtension(browser, { now = Date.now } = {}) {
  const extension = browser.acceptLanguageExtension;
  if (extension == null || browser.isClosed) return null;
  const { attempt } = extension;
  if (extension.inFlight != null && attempt != null) {
    if (now() >= attempt.holdUntil) attempt.releaseHold();
    if (now() < attempt.deadline) return null;
    const timeout = Object.assign(new Error("extension load timed out"), { name: "TimeoutError" });
    attempt.abort.abort(timeout);
    return finishAcceptLanguageExtension(browser, extension, attempt, { error: timeout }, now);
  }
  if (extension.verify != null && now() >= extension.verify.holdUntil) {
    extension.verify.release();
    extension.verify = null;
  }
  if (
    extension.phase === "failed" ||
    (extension.phase === "ready" && !extensionCoversList(browser))
  ) {
    return ensureAcceptLanguageExtension(browser, { now });
  }
  return null;
}

export function recordAcceptLanguage(browser) {
  if (!acceptLanguageEnUsEnabledForBrowser(browser)) return;
  const path =
    browser.acceptLanguageRecordPath ??
    `${BOX_BROWSER_ACCEPT_LANGUAGE_PATH_PREFIX}-${browser.port}`;
  const source =
    typeof browser.defaultBrowserContextId === "string" && launchFlagMatches(browser)
      ? "flag"
      : "cdp";
  try {
    const scratch = `${path}.${process.pid}`;
    writeFileSync(scratch, `${source} ${configuredAcceptLanguage(browser.acceptLanguage)}\n`, {
      mode: 0o644,
    });
    renameSync(scratch, path);
  } catch (error) {
    reportOnce(`accept-language record failed: ${errorClass(error)}`);
  }
}

export function reapplyChangedAcceptLanguage(
  browser,
  read = () => readAcceptLanguageEnUsMarker(undefined, null),
) {
  if (browser.isClosed || browser.acceptLanguageEnUsEnabled !== true) return null;
  const marker = read();
  if (marker == null) return null;
  const current = configuredAcceptLanguage(browser.acceptLanguage);
  const next = configuredAcceptLanguage(marker);
  if (next === current) return null;
  reportOnce(
    `accept-language changed port=${browser.port} (${current}->${next}); keeping ${current} until Chrome's next cold launch`,
  );
  return null;
}

export function egressTimeZoneEnabledForBrowser(
  browser,
  { resolveGate = resolveEgressTimeZoneGate, now = Date.now } = {},
) {
  if (!Object.prototype.hasOwnProperty.call(browser, "egressTimeZoneEnabled")) {
    const gate = resolveGate();
    browser.egressTimeZoneEnabled = gate === "on";
    if (gate === "undecided") browser.egressTimeZoneUndecidedSince = now();
  }
  return browser.egressTimeZoneEnabled === true;
}

export function applyEgressTimeZoneToTarget(browser, sessionId) {
  const timezoneId = browser.egressTimeZone;
  if (typeof timezoneId !== "string") return null;
  return browser.send("Emulation.setTimezoneOverride", { timezoneId }, sessionId);
}

export function reassertEgressTimeZone(browser) {
  for (const sessionId of browser.attachedSessions ?? []) {
    applyEgressTimeZoneToTarget(browser, sessionId)?.catch((error) => {
      if (isCdpTargetGone(error) || !browser.attachedSessions.has(sessionId)) return;
      console.error(`time zone apply failed port=${browser.port}: ${errorClass(error)}`);
    });
  }
}

export function egressTimeZoneHold(browser) {
  if (browser.egressTimeZoneEnabled !== true || typeof browser.egressTimeZone === "string") {
    return null;
  }
  return browser.egressTimeZoneHold ?? null;
}

function appendBoxTelemetry(event) {
  const path = process.env.SAND_BOX_TELEMETRY_LOG ?? "/tmp/sand-box-telemetry.log";
  appendFileSync(path, `${JSON.stringify(event)}\n`, "utf8");
}

function reportEgressTimeZone(browser, lookup, event) {
  if (lookup.reported.has(event.outcome)) return;
  lookup.reported.add(event.outcome);
  try {
    lookup.report({ kind: "egress_time_zone", ...event, boxClockUtc: lookup.boxClockIsUtc() });
  } catch (error) {
    console.error(`egress time zone telemetry failed port=${browser.port}: ${errorClass(error)}`);
  }
}

export function startEgressTimeZone(
  browser,
  {
    resolveZone = ({ port }) => resolveEgressTimeZone({ port }),
    now = Date.now,
    report = appendBoxTelemetry,
    boxClockIsUtc = isBoxClockUtc,
  } = {},
) {
  let release;
  browser.egressTimeZoneHold = new Promise((resolve) => {
    release = resolve;
  });
  browser.egressTimeZoneLookup = {
    resolveZone,
    report,
    boxClockIsUtc,
    reported: new Set(),
    attempts: 0,
    nextAttemptAt: now(),
    holdUntil: now() + EGRESS_TIME_ZONE_ATTACH_WAIT_MS,
    inFlight: null,
    releaseHold: () => {
      browser.egressTimeZoneHold = null;
      release();
    },
  };
  return tickEgressTimeZone(browser, { now });
}

function turnOnEgressTimeZone(browser) {
  browser.egressTimeZoneEnabled = true;
  void (browser.startEgressTimeZoneFor ?? startEgressTimeZone)(browser);
  const reportUnlessGone = (sessionId) => (error) => {
    if (
      isCdpTargetGone(error) ||
      (sessionId != null && !browser.attachedSessions?.has(sessionId))
    ) {
      return;
    }
    console.error(`time zone attach failed port=${browser.port}: ${errorClass(error)}`);
  };
  browser
    .send("Target.setAutoAttach", browserAutoAttachParams(browser))
    .catch(reportUnlessGone(null));
  for (const sessionId of browser.attachedSessions ?? []) {
    if (browser.attachedTargetTypes?.get(sessionId) === "service_worker") continue;
    attachRelatedTargets(browser, sessionId).catch(reportUnlessGone(sessionId));
  }
}

function settleEgressTimeZoneGate(browser, { now, resolveGate }) {
  const gate = resolveGate();
  if (
    gate === "undecided" &&
    now() - browser.egressTimeZoneUndecidedSince < EGRESS_TIME_ZONE_GATE_WAIT_MS
  ) {
    return;
  }
  browser.egressTimeZoneUndecidedSince = undefined;
  if (gate === "on" && !browser.isClosed) turnOnEgressTimeZone(browser);
}

export function tickEgressTimeZone(
  browser,
  { now = Date.now, resolveGate = resolveEgressTimeZoneGate } = {},
) {
  if (browser.egressTimeZoneUndecidedSince != null) {
    settleEgressTimeZoneGate(browser, { now, resolveGate });
  }
  const lookup = browser.egressTimeZoneLookup;
  if (lookup == null) return null;
  if (browser.egressTimeZoneHold != null && now() >= lookup.holdUntil) lookup.releaseHold();
  if (
    browser.isClosed ||
    typeof browser.egressTimeZone === "string" ||
    lookup.inFlight != null ||
    now() < lookup.nextAttemptAt
  ) {
    return lookup.inFlight;
  }
  lookup.inFlight = (async () => {
    try {
      const { timeZone, provider, source } = await lookup.resolveZone({ port: browser.port });
      if (browser.isClosed) return;
      browser.egressTimeZone = timeZone;
      reassertEgressTimeZone(browser);
      console.error(
        `egress time zone resolved port=${browser.port} provider=${provider} source=${source}`,
      );
      reportEgressTimeZone(browser, lookup, { outcome: "resolved", source, provider });
    } catch (error) {
      reportOnce(`egress time zone unresolved port=${browser.port}: ${egressErrorCode(error)}`);
      reportEgressTimeZone(browser, lookup, {
        outcome: "unresolved",
        errorClass: egressErrorClass(error),
      });
      const retryIndex = Math.min(lookup.attempts, EGRESS_TIME_ZONE_RETRY_MS.length - 1);
      lookup.nextAttemptAt = now() + EGRESS_TIME_ZONE_RETRY_MS[retryIndex];
      lookup.attempts += 1;
    } finally {
      lookup.inFlight = null;
      if (browser.egressTimeZoneHold != null) lookup.releaseHold();
    }
  })();
  return lookup.inFlight;
}

export function canvasSeedForBrowser(browser) {
  if (Object.prototype.hasOwnProperty.call(browser, "canvasSeed")) {
    const seed = browser.canvasSeed;
    return typeof seed === "number" && Number.isFinite(seed) ? seed >>> 0 : null;
  }
  if (!Object.prototype.hasOwnProperty.call(browser, "resolvedCanvasSeed")) {
    browser.resolvedCanvasSeed = resolveCanvasSeed();
  }
  return browser.resolvedCanvasSeed;
}

function documentScript(browser) {
  return buildHardwareProfileDocumentScript(LINUX_GPU_PATCH_PROFILE, {
    canvasSeed: canvasSeedForBrowser(browser),
  });
}

export async function removeDocumentScript(browser, sessionId) {
  const identifier = documentScriptMap(browser).get(sessionId);
  if (identifier == null) return;
  try {
    await browser.send("Page.removeScriptToEvaluateOnNewDocument", { identifier }, sessionId);
  } catch (error) {
    console.error(
      `document script remove failed: ${error instanceof Error ? error.name : typeof error}`,
    );
  }
  documentScriptMap(browser).delete(sessionId);
}

export async function installDocumentScript(browser, sessionId, script) {
  if (documentScriptMap(browser).has(sessionId)) await removeDocumentScript(browser, sessionId);
  const [, added] = await Promise.all([
    browser.send("Page.enable", {}, sessionId),
    browser.send("Page.addScriptToEvaluateOnNewDocument", { source: script }, sessionId),
  ]);
  if (typeof added?.identifier === "string") {
    documentScriptMap(browser).set(sessionId, added.identifier);
  }
}

function relatedTargetSessions(browser) {
  if (browser.relatedTargetSessions == null) {
    browser.relatedTargetSessions = new Map();
  }
  return browser.relatedTargetSessions;
}

/*
Chrome auto-attach reaches only the direct children of the session it is
called on, so each page, frame, and worker has to attach its own
out-of-process iframes and workers
(https://chromedevtools.github.io/devtools-protocol/tot/Target/#method-setAutoAttach).
*/
const RELATED_TARGET_FILTER = [
  { type: "iframe", exclude: false },
  { type: "worker", exclude: false },
  { type: "shared_worker", exclude: false },
  { exclude: true },
];

function serviceWorkersGoverned(browser) {
  return (
    acceptLanguageEnUsEnabledForBrowser(browser) ||
    egressTimeZoneEnabledForBrowser(browser) ||
    hardwareTreatmentForBrowser(browser) === "coherent"
  );
}

function fingerprintLoggerEnabled(browser) {
  return browser.fingerprintReadLogger != null;
}

async function applyFingerprintLoggerToTarget(browser, sessionId, target) {
  const logger = browser.fingerprintReadLogger;
  if (logger == null) return null;
  return logger.prepare({
    sessionId,
    parentSessionId: target?.parentSessionId,
    target,
  });
}

function attachRelatedTargets(browser, sessionId) {
  const sessions = relatedTargetSessions(browser);
  if (sessions.has(sessionId)) return Promise.resolve();
  sessions.set(sessionId, true);
  return browser
    .send(
      "Target.setAutoAttach",
      {
        autoAttach: true,
        waitForDebuggerOnStart: true,
        flatten: true,
        filter: RELATED_TARGET_FILTER,
      },
      sessionId,
    )
    .catch((error) => {
      sessions.delete(sessionId);
      throw error;
    });
}

function startsBeforeFirstDocument(target) {
  return target?.type === "iframe" && target.waitingForDebugger === true;
}

async function settleAll(steps) {
  const failed = (await Promise.allSettled(steps)).find((step) => step.status === "rejected");
  if (failed != null) throw failed.reason;
}

export async function applyUaTreatmentToTarget(browser, sessionId, target) {
  const hardware = hardwareTreatmentForBrowser(browser);
  const acceptLanguageEnUs = acceptLanguageEnUsEnabledForBrowser(browser);
  const egressTimeZone = egressTimeZoneEnabledForBrowser(browser);
  if (hardware === "off" && !acceptLanguageEnUs && !egressTimeZone) {
    await settleAll([
      applyDesktopUaToTarget(browser, sessionId, target?.browserContextId),
      fingerprintLoggerEnabled(browser) && ["page", "worker"].includes(target?.type)
        ? attachRelatedTargets(browser, sessionId)
        : null,
      applyFingerprintLoggerToTarget(browser, sessionId, target),
    ]);
    return;
  }
  const coherent = hardware === "coherent";
  if (isWorkerTargetType(target?.type)) {
    await settleAll([
      acceptLanguageEnUs || coherent
        ? applyDesktopUaToTarget(browser, sessionId, target?.browserContextId)
        : null,
      hardware === "legacy"
        ? browser.send("Runtime.evaluate", { expression: documentScript(browser) }, sessionId)
        : null,
      coherent ? applyCoherentHardwareToTarget(browser, sessionId) : null,
      egressTimeZone ? applyEgressTimeZoneToTarget(browser, sessionId) : null,
      target.type === "service_worker" ? null : attachRelatedTargets(browser, sessionId),
      applyFingerprintLoggerToTarget(browser, sessionId, target),
    ]);
    return;
  }
  if (hardware !== "legacy") {
    await settleAll([
      applyDesktopUaToTarget(browser, sessionId, target?.browserContextId),
      coherent ? applyCoherentHardwareToTarget(browser, sessionId) : null,
      egressTimeZone ? applyEgressTimeZoneToTarget(browser, sessionId) : null,
      attachRelatedTargets(browser, sessionId),
      applyFingerprintLoggerToTarget(browser, sessionId, target),
    ]);
    return;
  }
  const script = documentScript(browser);
  await settleAll([
    applyDesktopUaToTarget(browser, sessionId, target?.browserContextId),
    startsBeforeFirstDocument(target)
      ? null
      : browser.send("Runtime.evaluate", { expression: script }, sessionId),
    installDocumentScript(browser, sessionId, script),
    egressTimeZone ? applyEgressTimeZoneToTarget(browser, sessionId) : null,
    attachRelatedTargets(browser, sessionId),
    applyFingerprintLoggerToTarget(browser, sessionId, target),
  ]);
}

function browserAutoAttachParams(browser) {
  const hardware = hardwareTreatmentForBrowser(browser);
  const acceptLanguageEnUs = acceptLanguageEnUsEnabledForBrowser(browser);
  const egressTimeZone = egressTimeZoneEnabledForBrowser(browser);
  const filter =
    hardware !== "off" || acceptLanguageEnUs || egressTimeZone || fingerprintLoggerEnabled(browser)
      ? [
          { type: "page", exclude: false },
          { type: "worker", exclude: false },
          { type: "shared_worker", exclude: false },
          { exclude: true },
        ]
      : [{ type: "page", exclude: false }, { exclude: true }];
  if (serviceWorkersGoverned(browser)) {
    filter.unshift({ type: "service_worker", exclude: false });
  }
  return { autoAttach: true, waitForDebuggerOnStart: true, flatten: true, filter };
}

export async function configureUaGovernorBrowser(
  browser,
  { startEgressTimeZoneFor, configureFingerprintLoggerFor = configureFingerprintLogger } = {},
) {
  browser.attachedSessions = new Set();
  browser.attachedTargetTypes = new Map();
  browser.attachedBrowserContexts = new Map();
  browser.attachedTargetIds = new Map();
  if (startEgressTimeZoneFor != null) browser.startEgressTimeZoneFor = startEgressTimeZoneFor;
  await configureFingerprintLoggerFor(browser);
  if (egressTimeZoneEnabledForBrowser(browser) && browser.egressTimeZoneLookup === undefined) {
    void (browser.startEgressTimeZoneFor ?? startEgressTimeZone)(browser);
  }
  if (hardwareTreatmentForBrowser(browser) === "coherent") void verifyCoherentProfile(browser);
  browser.onEvent((message) => {
    if (message.method === "Target.targetDestroyed") {
      const targetId = message.params?.targetId;
      if (typeof targetId === "string") {
        void browser.fingerprintReadLogger?.targetDestroyed(targetId).catch((error) => {
          if (!isCdpTargetGone(error) && !browser.isClosed) {
            console.error(`fingerprint drain failed port=${browser.port}: ${errorClass(error)}`);
          }
        });
      }
      return;
    }
    const sessionId = message.params?.sessionId;
    if (typeof sessionId !== "string") return;
    if (message.method === "Target.detachedFromTarget") {
      browser.attachedSessions.delete(sessionId);
      browser.attachedTargetTypes.delete(sessionId);
      browser.attachedBrowserContexts.delete(sessionId);
      browser.attachedTargetIds.delete(sessionId);
      browser.fingerprintReadLogger?.detach(sessionId);
      browser.acceptLanguageOverriddenSessions?.delete(sessionId);
      browser.acceptLanguageUnappliedSessions?.delete(sessionId);
      browser.relatedTargetSessions?.delete(sessionId);
      if (browser.egressTimeZoneEnabled === true) reassertEgressTimeZone(browser);
      return;
    }
    if (message.method !== "Target.attachedToTarget") return;
    browser.attachedSessions.add(sessionId);
    const targetType = message.params?.targetInfo?.type;
    if (typeof targetType === "string") {
      browser.attachedTargetTypes.set(sessionId, targetType);
    }
    const browserContextId = message.params?.targetInfo?.browserContextId;
    browser.attachedBrowserContexts.set(sessionId, browserContextId);
    const targetId = message.params?.targetInfo?.targetId;
    if (typeof targetId === "string") browser.attachedTargetIds.set(sessionId, targetId);
    const reportUnlessGone = (step) => (error) => {
      if (isCdpTargetGone(error) || !browser.attachedSessions.has(sessionId)) return;
      console.error(`${step} failed port=${browser.port}: ${errorClass(error)}`);
    };
    const treatAndResume = () => {
      const waitingForDebugger = message.params?.waitingForDebugger === true;
      const applying = applyUaTreatmentToTarget(browser, sessionId, {
        ...message.params?.targetInfo,
        type: targetType,
        browserContextId,
        parentSessionId: message.sessionId,
        waitingForDebugger,
      });
      const resume = () =>
        void browser
          .send("Runtime.runIfWaitingForDebugger", {}, sessionId)
          .catch(reportUnlessGone("resume"));
      const acceptLanguageMustApply =
        acceptLanguageEnUsEnabledForBrowser(browser) &&
        browser.acceptLanguageOverriddenSessions?.has(sessionId);
      if (!acceptLanguageMustApply) {
        void applying.catch(reportUnlessGone("ua apply"));
        resume();
        return;
      }
      void applying.then(resume, (error) => {
        reportUnlessGone("ua apply")(error);
        if (!acceptLanguageMustApply) {
          resume();
        } else if (!browser.acceptLanguageUnappliedSessions?.has(sessionId)) {
          resume();
        } else if (!browser.attachedSessions.has(sessionId)) {
          return;
        } else if (acceptLanguageFromExtension(browser, browserContextId)) {
          console.error(
            `accept-language override failed port=${browser.port} ${targetType ?? "target"}; resuming it, since the extension already sets its Accept-Language`,
          );
          resume();
        } else if (targetType === "page" && typeof targetId === "string") {
          console.error(
            `accept-language override failed port=${browser.port} page; closing it so it does not keep running without the list in its Accept-Language and navigator.languages`,
          );
          void browser.send("Target.closeTarget", { targetId }).catch(reportUnlessGone("close"));
        } else if (waitingForDebugger) {
          console.error(
            `accept-language override failed port=${browser.port} ${targetType ?? "target"}; holding it paused so it never runs without the list in its Accept-Language and navigator.languages`,
          );
        }
      });
    };
    const holds = [
      egressTimeZoneHold(browser),
      acceptLanguageExtensionHold(browser, browserContextId),
    ].filter((hold) => hold != null);
    if (holds.length === 0) treatAndResume();
    else void (holds.length === 1 ? holds[0] : Promise.all(holds)).then(treatAndResume);
  });
  await resolveDefaultBrowserContext(browser);
  const settling = settleAcceptLanguageExtensionOnAttach(browser);
  if (settling != null && settling === browser.acceptLanguageExtensionCleanup) {
    await Promise.race([settling, delay(ACCEPT_LANGUAGE_EXTENSION_HOLD_MS)]);
  }
  await browser.send("Target.setAutoAttach", browserAutoAttachParams(browser));
  recordAcceptLanguage(browser);
}

export function resolveBoxBrowser({
  port,
  runningPath = `${BOX_BROWSER_RUNNING_PATH_PREFIX}-${port}`,
  readFile = readFileSync,
} = {}) {
  let running;
  try {
    running = String(readFile(runningPath, "utf8")).trim();
  } catch (error) {
    if (error?.code !== "ENOENT") {
      console.error(`running box browser read failed: ${errorClass(error)}`);
    }
    return "chrome";
  }
  return running === "ungoogled-chromium" ? "ungoogled-chromium" : "chrome";
}

export async function governUnlessUngoogled(browser, resolve = resolveBoxBrowser) {
  if (resolve({ port: browser.port }) === "ungoogled-chromium") return;
  await configureUaGovernorBrowser(browser);
}

export async function reapplyUaTreatment(browsers) {
  for (const browser of browsers.values()) {
    if (browser.isClosed) continue;
    for (const sessionId of browser.attachedSessions ?? []) {
      const targetType = browser.attachedTargetTypes?.get(sessionId);
      const browserContextId = browser.attachedBrowserContexts?.get(sessionId);
      await applyUaTreatmentToTarget(browser, sessionId, {
        type: targetType,
        browserContextId,
      }).catch((error) => {
        console.error(
          `ua re-apply failed port=${browser.port}: ${error instanceof Error ? error.name : typeof error}`,
        );
      });
    }
  }
}

export function openPageCount(browser) {
  let pages = 0;
  for (const type of browser.attachedTargetTypes?.values() ?? []) {
    if (type === "page") pages += 1;
  }
  return pages;
}

export function resolveBrowserSettings({
  resolveHardwareProfile = resolveHardwareProfileEnabled,
  resolveCoherentHardwareProfile = resolveCoherentHardwareProfileEnabled,
  readAcceptLanguage = readAcceptLanguageEnUsMarker,
  resolveEgressGate = resolveEgressTimeZoneGate,
  readFingerprintMarker = readFingerprintMarkerCandidate,
} = {}) {
  const fingerprintRaw = readFingerprintMarker();
  const fingerprintConfig = fingerprintRaw == null ? null : parseFingerprintMarker(fingerprintRaw);
  return {
    hardwareProfile: resolveHardwareProfile() === true,
    coherentHardwareProfile: resolveCoherentHardwareProfile() === true,
    acceptLanguage: readAcceptLanguage(),
    egressTimeZone: resolveEgressGate(),
    fingerprintLogConfigDigest:
      fingerprintConfig == null ? null : fingerprintMarkerDigest(fingerprintRaw),
  };
}

function onOff(value) {
  return value ? "on" : "off";
}

export function changedBrowserSettings(browser, settings) {
  const changes = [];
  const has = (key) => Object.prototype.hasOwnProperty.call(browser, key);
  if (has("hardwareProfileEnabled")) {
    const before = onOff(browser.hardwareProfileEnabled === true);
    const after = onOff(settings.hardwareProfile);
    if (before !== after) changes.push(`hardware-profile ${before}->${after}`);
  }
  if (has("coherentHardwareProfileEnabled")) {
    const before = onOff(browser.coherentHardwareProfileEnabled === true);
    const after = onOff(settings.coherentHardwareProfile);
    if (before !== after) changes.push(`coherent-hardware-profile ${before}->${after}`);
  }
  if (has("acceptLanguageEnUsEnabled")) {
    const enabled = browser.acceptLanguageEnUsEnabled === true;
    const wanted = settings.acceptLanguage != null;
    if (enabled !== wanted) {
      changes.push(`accept-language ${onOff(enabled)}->${onOff(wanted)}`);
    } else if (
      wanted &&
      !browser.launchAcceptLanguage &&
      (browser.acceptLanguage ?? "") !== settings.acceptLanguage
    ) {
      changes.push("accept-language value");
    }
  }
  if (has("egressTimeZoneEnabled")) {
    const before = onOff(browser.egressTimeZoneEnabled === true);
    const after = onOff(settings.egressTimeZone === "on");
    if (before !== after) changes.push(`egress-time-zone ${before}->${after}`);
  }
  if (has("fingerprintLogConfigDigest")) {
    const before = browser.fingerprintLogConfigDigest;
    const after = settings.fingerprintLogConfigDigest ?? null;
    if (before !== after) {
      changes.push(
        before == null || after == null
          ? `fingerprint-logging ${before == null ? "off" : "on"}->${after == null ? "off" : "on"}`
          : "fingerprint-logging value",
      );
    }
  }
  return changes;
}

export function reattachIfSettingsChanged(browser, resolve = resolveBrowserSettings) {
  if (browser.isClosed || !(browser.attachedTargetTypes instanceof Map)) return [];
  const changes = changedBrowserSettings(browser, resolve());
  if (changes.some((change) => change.startsWith("fingerprint-logging "))) {
    void Promise.resolve(browser.fingerprintReadLogger?.stop()).catch((error) => {
      if (!isCdpTargetGone(error) && !browser.isClosed) {
        console.error(`fingerprint cleanup failed port=${browser.port}: ${errorClass(error)}`);
      }
    });
    browser.fingerprintReadLogger = null;
    browser.fingerprintLogConfigDigest = null;
  }
  if (openPageCount(browser) > 0) return [];
  if (changes.length === 0) return [];
  console.error(
    `browser settings changed port=${browser.port} (${changes.join(", ")}); reattaching while no page is open`,
  );
  browser.close();
  return changes;
}

const GOVERNOR_DIR = dirname(fileURLToPath(import.meta.url));
export const UA_GOVERNOR_ENTRY = "sand-ua-governor.mjs";
const LOCAL_MODULE_SPECIFIER = /\b(?:from|import)\s*\(?\s*["']\.\/([\w.-]+\.mjs)["']/g;
const OPTIONAL_GOVERNOR_MODULES = [FINGERPRINT_LOGGER_MODULE.slice(2)];

export function governorCodeFingerprint(dir = GOVERNOR_DIR, readFile = readFileSync) {
  const sources = new Map();
  const pending = [UA_GOVERNOR_ENTRY, ...OPTIONAL_GOVERNOR_MODULES];
  while (pending.length > 0) {
    const name = pending.pop();
    if (sources.has(name)) continue;
    let source;
    try {
      source = String(readFile(join(dir, name), "utf8"));
    } catch (error) {
      if (OPTIONAL_GOVERNOR_MODULES.includes(name) && error?.code === "ENOENT") continue;
      throw error;
    }
    sources.set(name, source);
    for (const match of source.matchAll(LOCAL_MODULE_SPECIFIER)) pending.push(match[1]);
  }
  const modules = [...sources.keys()].sort();
  const hash = createHash("sha256");
  for (const name of modules) hash.update(`${name}\0${sources.get(name)}\0`);
  return { fingerprint: hash.digest("hex"), modules };
}

function readGovernorFingerprint(dir) {
  try {
    return governorCodeFingerprint(dir).fingerprint;
  } catch (error) {
    reportOnce(`governor code unreadable: ${errorClass(error)}`);
    return null;
  }
}

export const UPDATE_CHECK_INTERVAL_MS = 5_000;
const UPDATE_CHECK_INTERVAL_MIN_MS = 100;

export function updateCheckIntervalMs(raw = process.env.SAND_UA_GOVERNOR_UPDATE_CHECK_MS) {
  const parsed = Number.parseInt(raw ?? "", 10);
  if (!Number.isInteger(parsed) || parsed <= 0) return UPDATE_CHECK_INTERVAL_MS;
  return Math.max(UPDATE_CHECK_INTERVAL_MIN_MS, parsed);
}

export function createCodeUpdateWatch({
  dir = GOVERNOR_DIR,
  intervalMs = UPDATE_CHECK_INTERVAL_MS,
  now = Date.now,
  fingerprint = readGovernorFingerprint,
} = {}) {
  const loaded = fingerprint(dir);
  let checkedAtMs = now();
  let lastSeen = loaded;
  let settled = null;
  return {
    loaded,
    intervalMs,
    poll() {
      if (loaded == null || now() - checkedAtMs < intervalMs) return settled;
      checkedAtMs = now();
      const seen = fingerprint(dir);
      settled = seen != null && seen === lastSeen && seen !== loaded ? seen : null;
      lastSeen = seen;
      return settled;
    },
  };
}

const DEFAULT_DESKTOP_DIR = "/tmp/sand-desktop";
export const UA_GOVERNOR_DESKTOP_GROUP = "shared";
export const UA_GOVERNOR_DESKTOP_NAME = "ua-governor";

export function readGovernorRegistration({
  desktopDir = process.env.SAND_DESKTOP_DIR || DEFAULT_DESKTOP_DIR,
  readFile = readFileSync,
} = {}) {
  const groupDir = join(desktopDir, UA_GOVERNOR_DESKTOP_GROUP);
  const descriptorPath = join(groupDir, `${UA_GOVERNOR_DESKTOP_NAME}.json`);
  let descriptor;
  try {
    descriptor = JSON.parse(String(readFile(descriptorPath, "utf8")));
  } catch (error) {
    if (error?.code !== "ENOENT") {
      reportOnce(`desktop registration unreadable: ${errorClass(error)}`);
    }
    return { descriptorPath, pidFile: null, pid: null };
  }
  const pidFile =
    typeof descriptor?.pidFile === "string" && descriptor.pidFile.length > 0
      ? descriptor.pidFile
      : join(groupDir, `${UA_GOVERNOR_DESKTOP_NAME}.pid`);
  let pid = null;
  try {
    const parsed = Number.parseInt(String(readFile(pidFile, "utf8")).trim(), 10);
    pid = Number.isInteger(parsed) && parsed > 0 ? parsed : null;
  } catch (error) {
    if (error?.code !== "ENOENT") reportOnce(`desktop pid file unreadable: ${errorClass(error)}`);
  }
  return { descriptorPath, pidFile, pid };
}

export function planGovernorRestart({ openPages, quietForMs, quietNeededMs, supervised }) {
  if (openPages > 0) return "wait-for-pages";
  if (quietForMs < quietNeededMs) return "wait-for-quiet";
  return supervised() ? "restart" : "unsupervised";
}

function shortCode(fingerprint) {
  return typeof fingerprint === "string" ? fingerprint.slice(0, 12) : "unknown";
}

export function isGovernorProcess(pid, procRoot = "/proc") {
  let argv;
  try {
    argv = String(readFileSync(`${procRoot}/${pid}/cmdline`, "utf8")).split("\0");
  } catch (error) {
    if (error?.code !== "ENOENT" && error?.code !== "ESRCH") {
      console.error(`pid ${pid} cmdline unreadable: ${errorClass(error)}`);
    }
    return false;
  }
  return (
    argv.some((arg) => arg === UA_GOVERNOR_ENTRY || arg.endsWith(`/${UA_GOVERNOR_ENTRY}`)) &&
    !argv.includes("--restart")
  );
}

const RESTART_COMMAND_WAIT_MS = 20_000;
const RESTART_COMMAND_POLL_MS = 250;

export async function restartSupervisedGovernor({
  desktopDir = process.env.SAND_DESKTOP_DIR || DEFAULT_DESKTOP_DIR,
  procRoot = "/proc",
  kill = (pid, signal) => process.kill(pid, signal),
  waitMs = RESTART_COMMAND_WAIT_MS,
  pollMs = RESTART_COMMAND_POLL_MS,
  now = Date.now,
  sleep = delay,
  print = (line) => process.stdout.write(`${line}\n`),
} = {}) {
  const registration = readGovernorRegistration({ desktopDir });
  if (registration.pidFile == null) {
    print(
      `the UA governor is not registered with the desktop supervisor (${registration.descriptorPath} is missing or unreadable), so nothing would relaunch it`,
    );
    return 1;
  }
  const { pid } = registration;
  if (pid == null || !isGovernorProcess(pid, procRoot)) {
    print(
      `no UA governor is running at the pid in ${registration.pidFile}; the desktop supervisor relaunches a missing governor on its next tick`,
    );
    return 1;
  }
  try {
    kill(pid, "SIGTERM");
  } catch (error) {
    print(`could not signal UA governor pid ${pid}: ${errorClass(error)}`);
    return 1;
  }
  print(
    `sent SIGTERM to UA governor pid ${pid}; waiting for the desktop supervisor to relaunch it`,
  );
  const deadline = now() + waitMs;
  while (now() < deadline) {
    await sleep(pollMs);
    const next = readGovernorRegistration({ desktopDir }).pid;
    if (next != null && next !== pid && isGovernorProcess(next, procRoot)) {
      print(`UA governor relaunched as pid ${next}`);
      return 0;
    }
  }
  print(
    `no relaunch within ${Math.round(waitMs / 1000)}s; check the desktop supervisor log (/tmp/sand-desktop-supervisor.log, or /tmp/sand-supervisor.log on a box without the uid split)`,
  );
  return 1;
}

function restartForCodeUpdate({ updates, openPages, quietForMs }) {
  const pending = updates.poll();
  if (pending == null) return;
  const update = `code update ${shortCode(updates.loaded)} -> ${shortCode(pending)}`;
  const plan = planGovernorRestart({
    openPages,
    quietForMs,
    quietNeededMs: updates.intervalMs,
    supervised: () => readGovernorRegistration().pid === process.pid,
  });
  if (plan === "wait-for-pages") {
    reportOnce(`${update} pending; restarting once no governed browser has a page open`);
  } else if (plan === "unsupervised") {
    reportOnce(
      `${update} pending; this governor is not under desktop supervision, so the new code loads on the next box start`,
    );
  } else if (plan === "restart") {
    console.error(
      `${update}; no governed browser has had a page open for ${quietForMs}ms; exiting so the desktop supervisor relaunches the governor on the new code`,
    );
    process.exit(0);
  }
}

async function main() {
  const updates = createCodeUpdateWatch({ intervalMs: updateCheckIntervalMs() });
  process.once("SIGTERM", () => {
    console.error(`ua governor pid=${process.pid} stopping on SIGTERM`);
    process.exit(0);
  });
  console.error(`ua governor started pid=${process.pid} code=${shortCode(updates.loaded)}`);
  const browsers = new Map();
  let lastActivityAtMs = Date.now();
  for (;;) {
    for (const port of discoverMonitorPorts()) {
      const existing = browsers.get(port);
      if (existing != null && !existing.isClosed) continue;
      try {
        const chromeVersion = await getBrowserVersion(port);
        const browser = await connectBrowser(port);
        browser.chromeVersion = chromeVersion;
        await governUnlessUngoogled(browser);
        browsers.set(port, browser);
        lastActivityAtMs = Date.now();
      } catch (error) {
        if (!isConnectionRefused(error)) {
          reportOnce(`monitor on CDP port ${port} not governed: ${errorClass(error)}`);
        }
      }
    }
    let openPages = 0;
    for (const [port, browser] of browsers) {
      if (browser.isClosed) {
        browsers.delete(port);
        continue;
      }
      void tickEgressTimeZone(browser);
      void tickAcceptLanguageExtension(browser);
      reapplyChangedAcceptLanguage(browser);
      reattachIfSettingsChanged(browser);
      if (browser.isClosed) continue;
      void browser.fingerprintReadLogger?.tick().catch((error) => {
        if (!isCdpTargetGone(error) && !browser.isClosed) {
          console.error(`fingerprint drain failed port=${browser.port}: ${errorClass(error)}`);
        }
      });
      openPages += openPageCount(browser);
    }
    if (openPages > 0) lastActivityAtMs = Date.now();
    restartForCodeUpdate({ updates, openPages, quietForMs: Date.now() - lastActivityAtMs });
    await delay(POLL_INTERVAL_MS);
  }
}

if (process.argv[1] != null && import.meta.url === pathToFileURL(process.argv[1]).href) {
  if (process.argv.includes("--restart")) {
    restartSupervisedGovernor().then(
      (code) => {
        process.exitCode = code;
      },
      (error) => {
        console.error(`restart failed: ${errorClass(error)}`);
        process.exitCode = 1;
      },
    );
  } else {
    void main();
  }
}
