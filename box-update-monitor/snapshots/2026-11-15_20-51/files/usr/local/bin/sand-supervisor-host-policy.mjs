import { createHash } from "node:crypto";
import { closeSync, constants, fstatSync, openSync, readFileSync } from "node:fs";
import {
  SUPERVISOR_DIR,
  HOST_BUNDLE_SHA256_REGEX,
  COMMAND_KINDS,
  UPGRADE_MODES,
  RESTART_BACKOFF_BASE_MS,
  RESTART_BACKOFF_MAX_MS,
  MAX_DEFER_MS,
  POST_SWAP_HEALTHY_UPTIME_MS,
  POST_SWAP_MAX_QUICK_EXITS,
} from "./sand-supervisor-contract.mjs";

export const HOST_BUSY_STATES = ["no-host", "idle", "busy", "unknown"];
export const GATEWAY_DISCOVERY_RECORD_MAX_BYTES = 64 * 1024;
export const HOST_CRASH_EXIT_SIGNALS = Object.freeze([
  "SIGABRT",
  "SIGALRM",
  "SIGBUS",
  "SIGFPE",
  "SIGHUP",
  "SIGILL",
  "SIGINT",
  "SIGKILL",
  "SIGPIPE",
  "SIGQUIT",
  "SIGSEGV",
  "SIGTERM",
  "SIGTRAP",
  "SIGUSR1",
  "SIGUSR2",
  "SIGXCPU",
  "SIGXFSZ",
]);
const HOST_CRASH_EXIT_SIGNAL_SET = new Set(HOST_CRASH_EXIT_SIGNALS);

export function parseGatewayEndpoint(value) {
  if (typeof value !== "object" || value === null) return null;
  if (typeof value.port !== "number") return null;
  return {
    port: value.port,
    token: typeof value.token === "string" && value.token.length > 0 ? value.token : undefined,
  };
}

export function isSafeStagedBundlePath(path) {
  if (typeof path !== "string" || path.length === 0) return false;
  if (path.includes("\0")) return false;
  if (path.split("/").some((segment) => segment === "..")) return false;
  if (!path.startsWith(`${SUPERVISOR_DIR}/`)) return false;
  const rest = path.slice(SUPERVISOR_DIR.length + 1);
  return rest.length > 0 && rest.endsWith(".tgz");
}

export function normalizeUpgradeMode(command) {
  if (command.mode === "bundle" || command.mode === "restart") {
    return command.mode;
  }
  if (command.mode === "image") return "image";
  if (typeof command.bundlePath === "string" && isSafeStagedBundlePath(command.bundlePath)) {
    return "bundle";
  }
  return "image";
}

export function commandRequiresIdle(kind) {
  return kind === "restart" || kind === "upgrade";
}

export function decideUpgradeReadiness({ kind, busyState }) {
  if (!commandRequiresIdle(kind)) return "proceed";
  if (busyState === "idle" || busyState === "no-host") return "proceed";
  return "defer";
}

export function shouldForceHostUpgrade({ deferredForMs, maxDeferMs = MAX_DEFER_MS }) {
  if (!Number.isFinite(maxDeferMs) || maxDeferMs <= 0) return false;
  if (!Number.isFinite(deferredForMs) || deferredForMs < 0) return false;
  return deferredForMs >= maxDeferMs;
}

export function decideUpgradeAction({
  kind,
  busyState,
  deferredForMs,
  maxDeferMs = MAX_DEFER_MS,
  forceNow = false,
}) {
  if (decideUpgradeReadiness({ kind, busyState }) === "proceed") {
    return "proceed";
  }
  if (
    kind === "upgrade" &&
    (forceNow === true || shouldForceHostUpgrade({ deferredForMs, maxDeferMs }))
  ) {
    return "force";
  }
  return "defer";
}

export function isSandBoxAutoUpdateOptedOut(env) {
  const raw = env.SAND_BOX_AUTO_UPDATE?.trim().toLowerCase();
  return raw === "0" || raw === "false" || raw === "no";
}

export function shouldBootFetchHostBundle({
  hostSupervisionEnabled,
  autoUpdateOptedOut,
  bundlePresent,
  localVersion,
  imageSha,
}) {
  if (!hostSupervisionEnabled) return false;
  if (autoUpdateOptedOut) return false;
  if (!bundlePresent) return false;
  if (localVersion == null || imageSha == null) return false;
  return localVersion === imageSha;
}

export function decidePostSwapAction({
  armed,
  exitLaunchedAtMs,
  appliedAtMs,
  uptimeMs,
  quickExits,
  maxQuickExits = POST_SWAP_MAX_QUICK_EXITS,
  healthyUptimeMs = POST_SWAP_HEALTHY_UPTIME_MS,
}) {
  if (!armed) return "none";
  if (!Number.isFinite(exitLaunchedAtMs) || exitLaunchedAtMs < appliedAtMs) {
    return "none";
  }
  if (uptimeMs >= healthyUptimeMs) return "healthy";
  return quickExits + 1 >= maxQuickExits ? "rollback" : "count";
}

export function normalizeHostCrashExitSignal(signal) {
  if (signal == null) return "none";
  return HOST_CRASH_EXIT_SIGNAL_SET.has(signal) ? signal : "other";
}

export function classifyHostProcessExit({ code, signal, startedAtMs, crashedAtMs, uptimeMs }) {
  const exitSignal = normalizeHostCrashExitSignal(signal);
  if (exitSignal !== "none") {
    return {
      schemaVersion: 1,
      errorClass: "signal_exit",
      exitSignal,
      startedAtMs,
      crashedAtMs,
      uptimeMs,
    };
  }
  if (typeof code === "number") {
    return {
      schemaVersion: 1,
      errorClass: code === 0 ? "unexpected_clean_exit" : "nonzero_exit",
      exitSignal: "none",
      startedAtMs,
      crashedAtMs,
      uptimeMs,
    };
  }
  return {
    schemaVersion: 1,
    errorClass: "unobserved_exit",
    exitSignal: "unknown",
    ...(startedAtMs === undefined ? {} : { startedAtMs }),
    crashedAtMs,
    ...(uptimeMs === undefined ? {} : { uptimeMs }),
  };
}

export function classifyHostCrashMarkerWriteError(error) {
  const code = error != null && typeof error === "object" ? error.code : undefined;
  switch (code) {
    case "ENOSPC":
      return "no_space";
    case "EROFS":
      return "read_only";
    case "EACCES":
    case "EPERM":
      return "permission";
    case "ENOTDIR":
      return "not_directory";
    default:
      return "unknown";
  }
}

export function decideHostAction({
  bundlePresent,
  hostRunning,
  isBusy,
  lastExitAtMs,
  restartAttempts,
  nowMs,
}) {
  if (!bundlePresent) return "noop";
  if (hostRunning) return "noop";
  if (isBusy) return "wait";
  if (lastExitAtMs == null) return "launch";
  const backoff = nextBackoffMs(restartAttempts);
  return nowMs - lastExitAtMs >= backoff ? "restart" : "wait";
}

export function nextBackoffMs(
  attempts,
  baseMs = RESTART_BACKOFF_BASE_MS,
  maxMs = RESTART_BACKOFF_MAX_MS,
) {
  const exp = Math.min(attempts, 16);
  return Math.min(maxMs, baseMs * 2 ** exp);
}

export function shouldProcessCommand(command, state) {
  if (command == null) return false;
  return state === "pending";
}

export function buildStatus({
  nowMs,
  tickMs,
  hostBundlePresent,
  hostRunning,
  hostVersion,
  pendingUpgradeVersion,
  lastCommandId,
  lastCommandKind,
  gatewayDescriptorPresent,
  gatewayPidPresent,
}) {
  return {
    updatedAtMs: nowMs,
    tickMs,
    hostBundlePresent,
    hostRunning,
    hostVersion: hostVersion ?? null,
    pendingUpgradeVersion: pendingUpgradeVersion ?? null,
    lastCommandId: lastCommandId ?? null,
    lastCommandKind: lastCommandKind ?? null,
    gatewayDescriptorPresent: gatewayDescriptorPresent === true,
    gatewayPidPresent: gatewayPidPresent === true,
  };
}

export function parseHostBundleDigest(rawText) {
  const token = rawText.trim().split(/\s+/, 1)[0] ?? "";
  return HOST_BUNDLE_SHA256_REGEX.test(token) ? token : null;
}

export function createHostPolicy({ attempt }) {
  function parseCommand(raw) {
    const parsed = attempt(() => JSON.parse(raw));
    if (!parsed.ok) return null;
    const value = parsed.value;
    if (typeof value !== "object" || value === null) return null;
    if (typeof value.id !== "string" || value.id.length === 0) return null;
    if (typeof value.kind !== "string" || !COMMAND_KINDS.includes(value.kind)) {
      return null;
    }
    const issuedAtMs =
      typeof value.issuedAtMs === "number" && Number.isFinite(value.issuedAtMs)
        ? value.issuedAtMs
        : 0;
    return {
      id: value.id,
      kind: value.kind,
      issuedAtMs,
      reason: typeof value.reason === "string" ? value.reason : undefined,
      mode:
        typeof value.mode === "string" && UPGRADE_MODES.includes(value.mode)
          ? value.mode
          : undefined,
      version: typeof value.version === "string" ? value.version : undefined,
      bundlePath: typeof value.bundlePath === "string" ? value.bundlePath : undefined,
      sha256: typeof value.sha256 === "string" ? value.sha256 : undefined,
      forceNow: value.forceNow === true ? true : undefined,
    };
  }

  function verifyStagedBundleDigest(bundlePath, expectedSha256) {
    if (expectedSha256 == null) return { ok: false, reason: "digest_missing" };
    if (!HOST_BUNDLE_SHA256_REGEX.test(expectedSha256)) {
      return { ok: false, reason: "digest_malformed" };
    }
    const read = attempt(() => readFileSync(bundlePath));
    if (!read.ok) return { ok: false, reason: "bundle-missing" };
    const bytes = read.value;
    const actualSha256 = createHash("sha256").update(bytes).digest("hex");
    return actualSha256 === expectedSha256
      ? { ok: true, bytes }
      : { ok: false, reason: "digest_mismatch", actualSha256 };
  }

  function gatewayDiscoveryPresence(raw) {
    const parsed = raw == null ? { ok: false } : attempt(() => JSON.parse(raw));
    const pid = parsed.ok ? parsed.value?.pid : undefined;
    return {
      gatewayDescriptorPresent: parsed.ok,
      gatewayPidPresent: Number.isInteger(pid) && pid > 0,
    };
  }

  function readGatewayDiscoveryRecord(path, ownerUid = process.getuid?.()) {
    const opened = attempt(() =>
      openSync(path, constants.O_RDONLY | constants.O_NOFOLLOW | constants.O_NONBLOCK),
    );
    if (!opened.ok) return null;
    const read = attempt(() => {
      const info = fstatSync(opened.value);
      const owned = ownerUid === undefined || info.uid === ownerUid;
      const regular = info.isFile() && info.nlink === 1;
      return regular && owned && info.size <= GATEWAY_DISCOVERY_RECORD_MAX_BYTES
        ? readFileSync(opened.value, "utf8")
        : null;
    });
    attempt(() => closeSync(opened.value));
    return read.ok ? read.value : null;
  }

  return {
    parseCommand,
    verifyStagedBundleDigest,
    gatewayDiscoveryPresence,
    readGatewayDiscoveryRecord,
  };
}
