import { readFileSync } from "node:fs";
import { pathToFileURL } from "node:url";

import { BROWSER_HARDWARE_PROFILE_MARKER_PATH } from "./box-contract.generated.mjs";
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
    buildUserAgentOverride(liveChromeProduct(browser)),
    sessionId,
  );
}

export function documentScriptMap(browser) {
  if (browser.documentScripts == null) {
    browser.documentScripts = new Map();
  }
  return browser.documentScripts;
}

const WORKER_TARGET_TYPES = new Set(["worker", "shared_worker"]);

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
  const [, , added] = await Promise.all([
    browser.send("Runtime.evaluate", { expression: script }, sessionId),
    browser.send("Page.enable", {}, sessionId),
    browser.send("Page.addScriptToEvaluateOnNewDocument", { source: script }, sessionId),
  ]);
  if (typeof added?.identifier === "string") {
    documentScriptMap(browser).set(sessionId, added.identifier);
  }
}

export async function applyUaTreatmentToTarget(browser, sessionId, targetInfo) {
  if (!hardwareProfileEnabledForBrowser(browser)) {
    await applyDesktopUaToTarget(browser, sessionId);
    return;
  }
  if (isWorkerTargetType(targetInfo?.type)) {
    await browser.send("Runtime.evaluate", { expression: documentScript(browser) }, sessionId);
    return;
  }
  await Promise.all([
    applyDesktopUaToTarget(browser, sessionId),
    installDocumentScript(browser, sessionId, documentScript(browser)),
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
      return;
    }
    if (message.method !== "Target.attachedToTarget") return;
    browser.attachedSessions.add(sessionId);
    const targetType = message.params?.targetInfo?.type;
    if (typeof targetType === "string") {
      browser.attachedTargetTypes.set(sessionId, targetType);
    }
    const reportUnlessGone = (step) => (error) => {
      if (isCdpTargetGone(error)) return;
      console.error(`${step} failed port=${browser.port}: ${errorClass(error)}`);
    };
    void applyUaTreatmentToTarget(browser, sessionId, { type: targetType }).catch(
      reportUnlessGone("ua apply"),
    );
    void browser
      .send("Runtime.runIfWaitingForDebugger", {}, sessionId)
      .catch(reportUnlessGone("resume"));
  });
  const hardwareProfile = hardwareProfileEnabledForBrowser(browser);
  await browser.send("Target.setAutoAttach", {
    autoAttach: true,
    waitForDebuggerOnStart: true,
    flatten: true,
    filter: hardwareProfile
      ? [
          { type: "page", exclude: false },
          { type: "worker", exclude: false },
          { type: "shared_worker", exclude: false },
          { exclude: true },
        ]
      : [{ type: "page", exclude: false }, { exclude: true }],
  });
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
        await configureUaGovernorBrowser(browser);
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
