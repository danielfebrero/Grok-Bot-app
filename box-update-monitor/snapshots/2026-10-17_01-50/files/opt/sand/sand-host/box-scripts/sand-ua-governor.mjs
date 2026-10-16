import { existsSync, readFileSync } from "node:fs";
import { pathToFileURL } from "node:url";

import {
  BOX_BROWSER_RUNNING_PATH_PREFIX,
  BROWSER_ACCEPT_LANGUAGE_EN_US_MARKER_PATH,
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

const POLL_INTERVAL_MS = 100;

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

const HARDWARE_PROFILE_ON = new Set(["1", "on", "true"]);
const HARDWARE_PROFILE_OFF = new Set(["0", "off", "false"]);

export function resolveHardwareProfileEnabled({
  envValue = process.env.SAND_BROWSER_HARDWARE_PROFILE,
  markerPath = BROWSER_HARDWARE_PROFILE_MARKER_PATH,
  readFile = readFileSync,
} = {}) {
  if (envValue != null && String(envValue).trim() !== "") {
    const raw = String(envValue).trim().toLowerCase();
    if (HARDWARE_PROFILE_ON.has(raw)) return true;
    if (HARDWARE_PROFILE_OFF.has(raw)) return false;
    return false;
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
        `hardware profile marker read failed: ${error instanceof Error ? error.name : typeof error}`,
      );
    }
    return false;
  }
}

export function hardwareProfileEnabledForBrowser(browser, resolve = resolveHardwareProfileEnabled) {
  if (!Object.prototype.hasOwnProperty.call(browser, "hardwareProfileEnabled")) {
    browser.hardwareProfileEnabled = resolve() === true;
  }
  return browser.hardwareProfileEnabled === true;
}

export function resolveAcceptLanguageEnUsEnabled(
  markerPath = BROWSER_ACCEPT_LANGUAGE_EN_US_MARKER_PATH,
) {
  return existsSync(markerPath);
}

export function acceptLanguageEnUsEnabledForBrowser(
  browser,
  resolve = resolveAcceptLanguageEnUsEnabled,
) {
  if (!Object.prototype.hasOwnProperty.call(browser, "acceptLanguageEnUsEnabled")) {
    browser.acceptLanguageEnUsEnabled = resolve() === true;
  }
  return browser.acceptLanguageEnUsEnabled === true;
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
    browser.relatedTargetSessions = new Set();
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
  if (sessions.has(sessionId)) return Promise.resolve();
  sessions.add(sessionId);
  return browser
    .send(
      "Target.setAutoAttach",
      {
        autoAttach: true,
        waitForDebuggerOnStart: true,
        flatten: true,
        filter: acceptLanguageEnUsEnabledForBrowser(browser)
          ? [{ type: "service_worker", exclude: false }, ...RELATED_TARGET_FILTER]
          : RELATED_TARGET_FILTER,
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

export async function applyUaTreatmentToTarget(browser, sessionId, target) {
  const hardwareProfile = hardwareProfileEnabledForBrowser(browser);
  const acceptLanguageEnUs = acceptLanguageEnUsEnabledForBrowser(browser);
  if (!hardwareProfile && !acceptLanguageEnUs) {
    await applyDesktopUaToTarget(browser, sessionId);
    return;
  }
  if (isWorkerTargetType(target?.type)) {
    await Promise.all([
      acceptLanguageEnUs ? applyDesktopUaToTarget(browser, sessionId) : null,
      hardwareProfile
        ? browser.send("Runtime.evaluate", { expression: documentScript(browser) }, sessionId)
        : null,
      target.type === "service_worker" ? null : attachRelatedTargets(browser, sessionId),
    ]);
    return;
  }
  if (!hardwareProfile) {
    await Promise.all([
      applyDesktopUaToTarget(browser, sessionId),
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
    attachRelatedTargets(browser, sessionId),
  ]);
}

export async function configureUaGovernorBrowser(browser) {
  browser.attachedSessions = new Set();
  browser.attachedTargetTypes = new Map();
  browser.onEvent((message) => {
    const sessionId = message.params?.sessionId;
    if (typeof sessionId !== "string") return;
    if (message.method === "Target.detachedFromTarget") {
      browser.attachedSessions.delete(sessionId);
      browser.attachedTargetTypes.delete(sessionId);
      browser.relatedTargetSessions?.delete(sessionId);
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
    void applyUaTreatmentToTarget(browser, sessionId, {
      type: targetType,
      waitingForDebugger: message.params?.waitingForDebugger === true,
    }).catch(reportUnlessGone("ua apply"));
    void browser
      .send("Runtime.runIfWaitingForDebugger", {}, sessionId)
      .catch(reportUnlessGone("resume"));
  });
  const hardwareProfile = hardwareProfileEnabledForBrowser(browser);
  const acceptLanguageEnUs = acceptLanguageEnUsEnabledForBrowser(browser);
  const filter =
    hardwareProfile || acceptLanguageEnUs
      ? [
          { type: "page", exclude: false },
          { type: "worker", exclude: false },
          { type: "shared_worker", exclude: false },
          { exclude: true },
        ]
      : [{ type: "page", exclude: false }, { exclude: true }];
  if (acceptLanguageEnUs) filter.unshift({ type: "service_worker", exclude: false });
  await browser.send("Target.setAutoAttach", {
    autoAttach: true,
    waitForDebuggerOnStart: true,
    flatten: true,
    filter,
  });
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
      if (!browser.isClosed) continue;
      browsers.delete(port);
    }
    await new Promise((resolve) => setTimeout(resolve, POLL_INTERVAL_MS));
  }
}

if (process.argv[1] != null && import.meta.url === pathToFileURL(process.argv[1]).href) {
  void main();
}
