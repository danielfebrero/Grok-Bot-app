import { appendFileSync, readFileSync } from "node:fs";
import { pathToFileURL } from "node:url";

import {
  BOX_BROWSER_RUNNING_PATH_PREFIX,
  BOX_EXPERIMENTS_READY_MARKER_PATH,
  BROWSER_ACCEPT_LANGUAGE_EN_US_MARKER_PATH,
  BROWSER_EGRESS_TIMEZONE_MARKER_PATH,
  BROWSER_HARDWARE_PROFILE_MARKER_PATH,
} from "./box-contract.generated.mjs";
import {
  connectBrowser,
  discoverMonitorPorts,
  getBrowserVersion,
  isCdpTargetGone,
  isConnectionRefused,
} from "./cdp-cookies.mjs";
import {
  LINUX_GPU_PATCH_PROFILE,
  buildHardwareProfileDocumentScript,
  buildUserAgentOverride,
  resolveCanvasSeed,
} from "./sand-fingerprint-profiles.mjs";
import {
  egressErrorClass,
  egressErrorCode,
  isBoxClockUtc,
  resolveEgressTimeZone,
} from "./sand-egress-timezone.mjs";

const POLL_INTERVAL_MS = 100;
export const EGRESS_TIME_ZONE_ATTACH_WAIT_MS = 3_000;
export const EGRESS_TIME_ZONE_RETRY_MS = Object.freeze([5_000, 30_000, 120_000, 300_000]);
export const EGRESS_TIME_ZONE_GATE_WAIT_MS = 120_000;

function errorClass(error) {
  return error?.code ?? error?.cause?.code ?? (error instanceof Error ? error.name : typeof error);
}

const reportedOnce = new Set();

function reportOnce(message) {
  if (reportedOnce.has(message)) return;
  reportedOnce.add(message);
  console.error(message);
}

function liveChromeProduct(browser) {
  const version = typeof browser?.chromeVersion === "string" ? browser.chromeVersion : "";
  if (/^\d+\.\d+\.\d+\.\d+$/.test(version)) return `Chrome/${version}`;
  return "Chrome/0.0.0.0";
}

export async function applyDesktopUaToTarget(browser, sessionId) {
  await browser.send(
    "Emulation.setUserAgentOverride",
    buildUserAgentOverride(liveChromeProduct(browser), {
      acceptLanguageEnUs: acceptLanguageEnUsEnabledForBrowser(browser),
      acceptLanguage: browser.acceptLanguage,
    }),
    sessionId,
  );
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

function resolveMarkerGate({ label, envValue, markerPath, readFile }) {
  if (envValue != null && String(envValue).trim() !== "") {
    return MARKER_GATE_ON.has(String(envValue).trim().toLowerCase());
  }
  try {
    readFile(markerPath);
    return true;
  } catch (error) {
    if (
      error == null ||
      typeof error !== "object" ||
      !("code" in error) ||
      error.code !== "ENOENT"
    ) {
      console.error(
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
  return resolveMarkerGate({ label: "hardware profile", envValue, markerPath, readFile });
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

export function readAcceptLanguageEnUsMarker(
  markerPath = BROWSER_ACCEPT_LANGUAGE_EN_US_MARKER_PATH,
) {
  try {
    return readFileSync(markerPath, "utf8");
  } catch (error) {
    if (error?.code === "ENOENT") return null;
    console.error(`accept-language marker read failed: ${errorClass(error)}`);
    return "";
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
  }
  return browser.acceptLanguageEnUsEnabled === true;
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

function attachRelatedTargets(browser, sessionId) {
  const sessions = relatedTargetSessions(browser);
  const serviceWorkers =
    acceptLanguageEnUsEnabledForBrowser(browser) || egressTimeZoneEnabledForBrowser(browser);
  const attached = sessions.get(sessionId);
  if (attached === true || attached === serviceWorkers) return Promise.resolve();
  sessions.set(sessionId, serviceWorkers);
  return browser
    .send(
      "Target.setAutoAttach",
      {
        autoAttach: true,
        waitForDebuggerOnStart: true,
        flatten: true,
        filter: serviceWorkers
          ? [{ type: "service_worker", exclude: false }, ...RELATED_TARGET_FILTER]
          : RELATED_TARGET_FILTER,
      },
      sessionId,
    )
    .catch((error) => {
      if (attached === undefined) sessions.delete(sessionId);
      else sessions.set(sessionId, attached);
      throw error;
    });
}

function startsBeforeFirstDocument(target) {
  return target?.type === "iframe" && target.waitingForDebugger === true;
}

export async function applyUaTreatmentToTarget(browser, sessionId, target) {
  const hardwareProfile = hardwareProfileEnabledForBrowser(browser);
  const acceptLanguageEnUs = acceptLanguageEnUsEnabledForBrowser(browser);
  const egressTimeZone = egressTimeZoneEnabledForBrowser(browser);
  if (!hardwareProfile && !acceptLanguageEnUs && !egressTimeZone) {
    await applyDesktopUaToTarget(browser, sessionId);
    return;
  }
  if (isWorkerTargetType(target?.type)) {
    await Promise.all([
      acceptLanguageEnUs ? applyDesktopUaToTarget(browser, sessionId) : null,
      hardwareProfile
        ? browser.send("Runtime.evaluate", { expression: documentScript(browser) }, sessionId)
        : null,
      egressTimeZone ? applyEgressTimeZoneToTarget(browser, sessionId) : null,
      target.type === "service_worker" ? null : attachRelatedTargets(browser, sessionId),
    ]);
    return;
  }
  if (!hardwareProfile) {
    await Promise.all([
      applyDesktopUaToTarget(browser, sessionId),
      egressTimeZone ? applyEgressTimeZoneToTarget(browser, sessionId) : null,
      attachRelatedTargets(browser, sessionId),
    ]);
    return;
  }
  const script = documentScript(browser);
  await Promise.all([
    applyDesktopUaToTarget(browser, sessionId),
    startsBeforeFirstDocument(target)
      ? null
      : browser.send("Runtime.evaluate", { expression: script }, sessionId),
    installDocumentScript(browser, sessionId, script),
    egressTimeZone ? applyEgressTimeZoneToTarget(browser, sessionId) : null,
    attachRelatedTargets(browser, sessionId),
  ]);
}

function browserAutoAttachParams(browser) {
  const hardwareProfile = hardwareProfileEnabledForBrowser(browser);
  const acceptLanguageEnUs = acceptLanguageEnUsEnabledForBrowser(browser);
  const egressTimeZone = egressTimeZoneEnabledForBrowser(browser);
  const filter =
    hardwareProfile || acceptLanguageEnUs || egressTimeZone
      ? [
          { type: "page", exclude: false },
          { type: "worker", exclude: false },
          { type: "shared_worker", exclude: false },
          { exclude: true },
        ]
      : [{ type: "page", exclude: false }, { exclude: true }];
  if (acceptLanguageEnUs || egressTimeZone) {
    filter.unshift({ type: "service_worker", exclude: false });
  }
  return { autoAttach: true, waitForDebuggerOnStart: true, flatten: true, filter };
}

export async function configureUaGovernorBrowser(browser, { startEgressTimeZoneFor } = {}) {
  browser.attachedSessions = new Set();
  browser.attachedTargetTypes = new Map();
  if (startEgressTimeZoneFor != null) browser.startEgressTimeZoneFor = startEgressTimeZoneFor;
  if (egressTimeZoneEnabledForBrowser(browser) && browser.egressTimeZoneLookup === undefined) {
    void (browser.startEgressTimeZoneFor ?? startEgressTimeZone)(browser);
  }
  browser.onEvent((message) => {
    const sessionId = message.params?.sessionId;
    if (typeof sessionId !== "string") return;
    if (message.method === "Target.detachedFromTarget") {
      browser.attachedSessions.delete(sessionId);
      browser.attachedTargetTypes.delete(sessionId);
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
    const reportUnlessGone = (step) => (error) => {
      if (isCdpTargetGone(error) || !browser.attachedSessions.has(sessionId)) return;
      console.error(`${step} failed port=${browser.port}: ${errorClass(error)}`);
    };
    const treatAndResume = () => {
      void applyUaTreatmentToTarget(browser, sessionId, {
        type: targetType,
        waitingForDebugger: message.params?.waitingForDebugger === true,
      }).catch(reportUnlessGone("ua apply"));
      void browser
        .send("Runtime.runIfWaitingForDebugger", {}, sessionId)
        .catch(reportUnlessGone("resume"));
    };
    const hold = egressTimeZoneHold(browser);
    if (hold == null) treatAndResume();
    else void hold.then(treatAndResume);
  });
  await browser.send("Target.setAutoAttach", browserAutoAttachParams(browser));
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
      await applyUaTreatmentToTarget(browser, sessionId, { type: targetType }).catch((error) => {
        console.error(
          `ua re-apply failed port=${browser.port}: ${error instanceof Error ? error.name : typeof error}`,
        );
      });
    }
  }
}

async function main() {
  const browsers = new Map();
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
      } catch (error) {
        if (!isConnectionRefused(error)) {
          reportOnce(`monitor on CDP port ${port} not governed: ${errorClass(error)}`);
        }
      }
    }
    for (const [port, browser] of browsers) {
      if (browser.isClosed) browsers.delete(port);
      else void tickEgressTimeZone(browser);
    }
    await new Promise((resolve) => setTimeout(resolve, POLL_INTERVAL_MS));
  }
}

if (process.argv[1] != null && import.meta.url === pathToFileURL(process.argv[1]).href) {
  void main();
}
