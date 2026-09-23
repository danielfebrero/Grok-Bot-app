import http from "node:http";
import https from "node:https";
import { spawn } from "node:child_process";
import { createHash } from "node:crypto";
import {
  appendFileSync,
  closeSync,
  createWriteStream,
  existsSync,
  fstatSync,
  mkdirSync,
  openSync,
  readFileSync,
  readdirSync,
  renameSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import net from "node:net";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import {
  SUPERVISOR_DIR,
  COMMAND_PATH,
  STATUS_PATH,
  ACKS_DIR,
  STAGED_BUNDLE_PATH,
  AGENT_DATA_ROOT,
  UPGRADE_APPLIED_PATH,
  HOST_CRASH_MARKER_PATH,
  HOST_DIR,
  HOST_ENTRY,
  HOST_NODE_WARNING_FLAG,
  HOST_LOG_PATH,
  HOST_VERSION_PATH,
  IMAGE_SHA_PATH,
  HOST_BUNDLE_BASE_URL_ENV,
  HOST_BUNDLE_CHANNELS,
  HOST_BUNDLE_DEFAULT_CHANNEL,
  HOST_BUNDLE_CHANNEL_ENV,
  BOOT_FETCH_GIT_SHA_REGEX,
  BOOT_FETCH_BUDGET_MS_ENV,
  BOOT_FETCH_BUDGET_DEFAULT_MS,
  BOX_SCRIPTS_SOURCE_DIR,
  BOX_SCRIPTS_BIN_DIR,
  BOX_SCRIPTS_MARKER_PATH,
  BOX_SCRIPTS_SYNC_DISABLED_ENV,
  BOX_SCRIPTS_RETRY_BACKOFF_MS,
  TICK_MS,
  HEALTH_TIMEOUT_MS,
  FORK_RFB_PROBE_TIMEOUT_MS,
  FORK_RFB_FAILURE_THRESHOLD,
  FORK_RFB_TERMINATE_GRACE_MS,
  MAX_DEFER_MS,
  POST_SWAP_HEALTHY_UPTIME_MS,
  POST_SWAP_MAX_QUICK_EXITS,
  POST_SWAP_MAX_RESTORE_ATTEMPTS,
  numberFromEnv,
  isDesktopSupervisionEnabled,
  isHostSupervisionEnabled,
  log,
} from "./sand-supervisor-contract.mjs";
import {
  parseGatewayEndpoint,
  isSafeStagedBundlePath,
  normalizeUpgradeMode,
  commandRequiresIdle,
  decideUpgradeAction,
  isSandBoxAutoUpdateOptedOut,
  shouldBootFetchHostBundle,
  decidePostSwapAction,
  classifyHostProcessExit,
  classifyHostCrashMarkerWriteError,
  decideHostAction,
  shouldProcessCommand,
  buildStatus,
  parseHostBundleDigest,
  createHostPolicy,
} from "./sand-supervisor-host-policy.mjs";
import { createProcessIdentity } from "./sand-supervisor-process-identity.mjs";
import { CGROUP_INTERACTIVE } from "./sand-supervisor-cgroup-accounting.mjs";
import { createRfbHandshake, parseForkRfbTokenRecord } from "./sand-supervisor-rfb-handshake.mjs";
import {
  hostBundleSwapTargets,
  bundleSwapTargetsClearOfDataRoot,
  createHostBundleStore,
} from "./sand-supervisor-host-bundle.mjs";
import {
  desktopLogIdentity,
  isBoundedDesktopLogSpec,
  createDesktopSupervisor,
} from "./sand-supervisor-desktop.mjs";

export {
  SUPERVISOR_DIR,
  COMMAND_PATH,
  STATUS_PATH,
  ACKS_DIR,
  STAGED_BUNDLE_PATH,
  AGENT_DATA_ROOT,
  UPGRADE_APPLIED_PATH,
  HOST_CRASH_MARKER_PATH,
  HOST_DIR,
  HOST_ENTRY,
  HOST_NODE_WARNING_FLAG,
  HOST_LOG_PATH,
  HOST_VERSION_PATH,
  IMAGE_SHA_PATH,
  HOST_BUNDLE_CHANNELS,
  HOST_BUNDLE_DEFAULT_CHANNEL,
  BOOT_FETCH_GIT_SHA_REGEX,
  HOST_BUNDLE_SHA256_REGEX,
  DESKTOP_DIR,
  DESKTOP_HEALTH_PATH,
  DESKTOP_SUPERVISION_DISABLED_ENV,
  BOX_SCRIPTS_SOURCE_DIR,
  BOX_SCRIPTS_BIN_DIR,
  BOX_SCRIPTS_MARKER_PATH,
  BOX_SCRIPTS_SYNC_DISABLED_ENV,
  BOX_SCRIPTS_RETRY_BACKOFF_MS,
  BOX_SCRIPTS_DENY,
  COMMAND_KINDS,
  UPGRADE_MODES,
  HOST_UPGRADE_MAX_DEFER_MS,
  isDesktopSupervisionEnabled,
  isHostSupervisionEnabled,
} from "./sand-supervisor-contract.mjs";
export {
  HOST_BUSY_STATES,
  HOST_CRASH_EXIT_SIGNALS,
  parseGatewayEndpoint,
  isSafeStagedBundlePath,
  normalizeUpgradeMode,
  commandRequiresIdle,
  decideUpgradeReadiness,
  shouldForceHostUpgrade,
  decideUpgradeAction,
  isSandBoxAutoUpdateOptedOut,
  shouldBootFetchHostBundle,
  decidePostSwapAction,
  normalizeHostCrashExitSignal,
  classifyHostProcessExit,
  decideHostAction,
  nextBackoffMs,
  shouldProcessCommand,
  buildStatus,
  parseHostBundleDigest,
} from "./sand-supervisor-host-policy.mjs";
export { isPidAlive, listenerArgvMatches } from "./sand-supervisor-process-identity.mjs";
export {
  CGROUP_ROOT,
  CGROUP_INTERACTIVE,
  CGROUP_AGENT,
  parseCgroupCpuStat,
  parseCgroupPressure,
  formatCgroupCpuSample,
} from "./sand-supervisor-cgroup-accounting.mjs";
export { parseForkRfbTokenRecord } from "./sand-supervisor-rfb-handshake.mjs";
export {
  hostBundleSwapTargets,
  bundleSwapTargetsClearOfDataRoot,
  pathsOverlap,
  classifySwapError,
  classifyExtractError,
} from "./sand-supervisor-host-bundle.mjs";
export {
  countRestartsInWindow,
  decideDesktopComponentAction,
  shouldDisableCompositor,
  classifyDesktopDownReason,
  desktopStateSignature,
  buildDesktopHealthSnapshot,
} from "./sand-supervisor-desktop.mjs";

const { parseCommand, verifyStagedBundleDigest } = createHostPolicy({ attempt });
const processIdentity = createProcessIdentity({ attempt, safeRead });
const {
  inspectTcpPortListener,
  isPidComm,
  isExpectedForkRfbProcess,
  isDockPid,
  findCompositorReapTargets,
  readProcStartTime,
} = processIdentity;
const { DesktopSupervisor, parseDesktopComponentDescriptor, desktopLogCursor } =
  createDesktopSupervisor({
    attempt,
    safeRead,
    atomicWrite,
    closeLog,
    processIdentity,
  });
const {
  keepExistingHostBackup,
  swapHostBundle,
  restoreHostBundleFromBackup,
  syncBoxScriptsFromDir,
} = createHostBundleStore({
  attempt,
  safeUnlink,
  safeRmRecursive,
  readVersionMarker,
  ensureDir,
});

export {
  parseCommand,
  verifyStagedBundleDigest,
  parseDesktopComponentDescriptor,
  inspectTcpPortListener,
  isPidComm,
  isDockPid,
  findCompositorReapTargets,
  readProcStartTime,
  keepExistingHostBackup,
  swapHostBundle,
  restoreHostBundleFromBackup,
  syncBoxScriptsFromDir,
};

function attempt(operation) {
  try {
    return { ok: true, value: operation() };
  } catch (error) {
    return { ok: false, error };
  }
}

function closeLog(fd) {
  try {
    closeSync(fd);
    return null;
  } catch (error) {
    return error;
  }
}

export const HOST_BUNDLE_DEFAULT_BASE_URL =
  "https://public-asphr-vm-daemon-bucket.s3.us-east-1.amazonaws.com/sand-host-bundle";

// >>> box port table (from sand/src/shared/box/box-contract.ts; regenerate: pnpm --filter sand run gen:box-ports) >>>
const FORK_RFB_TOKEN_DIR = "/tmp/sand-novnc-tokens.d";
const FORK_RFB_WEBSOCKET_PORT = 6081;
// <<< box port table <<<
export const POST_SWAP_CRASH_LOOP_ERROR_CLASS = "post-swap-crash-loop";
export const POST_SWAP_ROLLBACK_FAILED_ERROR_CLASS = "rollback-failed";
const PREPARE_UPGRADE_PATH = "/prepare-upgrade";

export function probeRfbTcp(port, timeoutMs = FORK_RFB_PROBE_TIMEOUT_MS) {
  return new Promise((resolve) => {
    const socket = net.createConnection({ host: "127.0.0.1", port });
    let settled = false;
    const settle = (healthy) => {
      if (settled) return;
      settled = true;
      for (const timer of [deadline, totalDeadline]) clearTimeout(timer);
      socket.destroy();
      resolve(healthy);
    };
    const feed = createRfbHandshake((bytes) => socket.write(bytes), settle);
    const [deadline, totalDeadline] = [timeoutMs, timeoutMs * 5].map((budget) =>
      setTimeout(() => settle(false), budget),
    );
    socket.on("data", (chunk) => {
      if (chunk.length > 0) deadline.refresh();
      feed(chunk);
    });
    socket.on("error", () => settle(false));
    socket.on("end", () => settle(false));
  });
}

export function probeRfbWebSocket({ port, token, timeoutMs = FORK_RFB_PROBE_TIMEOUT_MS }) {
  return new Promise((resolve) => {
    if (typeof WebSocket !== "function") {
      resolve(false);
      return;
    }
    const url = new URL(`ws://127.0.0.1:${port}/websockify`);
    url.searchParams.set("token", token);
    let socket;
    try {
      socket = new WebSocket(url, "binary");
    } catch {
      resolve(false);
      return;
    }
    socket.binaryType = "arraybuffer";
    let settled = false;
    const settle = (healthy) => {
      if (settled) return;
      settled = true;
      for (const timer of [deadline, totalDeadline]) clearTimeout(timer);
      try {
        socket.close();
      } catch {}
      resolve(healthy);
    };
    const feed = createRfbHandshake((bytes) => socket.send(bytes), settle);
    const [deadline, totalDeadline] = [timeoutMs, timeoutMs * 5].map((budget) =>
      setTimeout(() => settle(false), budget),
    );
    socket.addEventListener("message", (event) => {
      if (typeof event.data === "string") {
        settle(false);
        return;
      }
      const chunk = Buffer.from(event.data);
      if (chunk.length > 0) deadline.refresh();
      feed(chunk);
    });
    socket.addEventListener("error", () => settle(false));
    socket.addEventListener("close", () => settle(false));
  });
}

export async function probeForkRfbChain({
  vncPort,
  websocketPort = FORK_RFB_WEBSOCKET_PORT,
  token,
  timeoutMs = FORK_RFB_PROBE_TIMEOUT_MS,
}) {
  if (await probeRfbWebSocket({ port: websocketPort, token, timeoutMs })) {
    return "healthy";
  }
  return (await probeRfbTcp(vncPort, timeoutMs)) ? "fork-websockify" : "x11vnc";
}

export class Supervisor extends DesktopSupervisor {
  constructor(opts = {}) {
    super();
    this.markerPath = opts.markerPath ?? UPGRADE_APPLIED_PATH;
    this.crashMarkerPath = opts.crashMarkerPath ?? HOST_CRASH_MARKER_PATH;
    this.imageShaPath = opts.imageShaPath ?? IMAGE_SHA_PATH;
    this.commandPath = opts.commandPath ?? COMMAND_PATH;
    this.acksDir = opts.acksDir ?? ACKS_DIR;
    this.appliedCommandIds = new Set();
    this.stuckCommand = null;
    this.bootFetchBudgetMs =
      opts.bootFetchBudgetMs ??
      numberFromEnv(BOOT_FETCH_BUDGET_MS_ENV, BOOT_FETCH_BUDGET_DEFAULT_MS, 1_000);
    this.telemetryPath =
      opts.telemetryPath ?? process.env.SAND_BOX_TELEMETRY_LOG ?? "/tmp/sand-box-telemetry.log";
    this.bootFetch = null;
    this.child = null;
    this.adoptedHost = null;
    this.stoppedHostPids = new Set();
    this.expectedHostExits = new WeakSet();
    this.lastExitAtMs = null;
    this.restartAttempts = 0;
    this.lastCommandId = null;
    this.lastCommandKind = null;
    this.pendingUpgradeVersion = null;
    this.pendingUpgradeDeferredSinceMs = null;
    this.hostPauseRequested = false;
    this.postSwapRollback = null;
    this.lastLaunchAtMs = null;
    this.loopKeepAliveTimer = null;
    this.tickInFlight = null;
    this.stopping = new AbortController();
    this.hostLogFd = null;
    super.initializeDesktopState(opts);
    this.forkRfbTokenDir = opts.forkRfbTokenDir ?? FORK_RFB_TOKEN_DIR;
    this.forkRfbWebsocketPort = opts.forkRfbWebsocketPort ?? FORK_RFB_WEBSOCKET_PORT;
    this.forkRfbProbeTimeoutMs = opts.forkRfbProbeTimeoutMs ?? FORK_RFB_PROBE_TIMEOUT_MS;
    this.forkRfbFailureThreshold = opts.forkRfbFailureThreshold ?? FORK_RFB_FAILURE_THRESHOLD;
    this.forkRfbFailures = new Map();
    this.forkRfbProbeCursor = 0;
    this.compositorReapEscalateMs = opts.compositorReapEscalateMs ?? 2_000;
    super.initializeCgroupAccounting(opts);
  }

  start() {
    for (const dir of [SUPERVISOR_DIR, this.acksDir, AGENT_DATA_ROOT]) {
      const error = ensureDir(dir);
      if (error != null) log(`cannot create ${dir}: ${String(error)}`);
    }
    log(
      `started (tick ${TICK_MS}ms, host supervision ${
        isHostSupervisionEnabled() ? "on" : "off"
      }, host bundle ${
        this.hostBundlePresent() ? "present" : "absent — dormant"
      }, host version ${this.readHostVersion() ?? "unknown"}, desktop supervision ${
        isDesktopSupervisionEnabled() ? "on" : "off"
      })`,
    );
    if (isHostSupervisionEnabled()) this.maybeAdoptOrphanHost();
    this.maybeStartBootFetch();
    this.scheduleTick(0);
  }

  async stop() {
    this.stopping.abort();
    if (this.loopKeepAliveTimer != null) clearTimeout(this.loopKeepAliveTimer);
    await this.tickInFlight;
  }

  scheduleTick(delayMs) {
    if (this.stopping.signal.aborted) return;
    if (this.loopKeepAliveTimer != null) clearTimeout(this.loopKeepAliveTimer);
    this.loopKeepAliveTimer = setTimeout(() => {
      this.tickInFlight = this.tick();
    }, delayMs);
  }

  hostMutationAllowedNow() {
    return this.postSwapRollback?.phase !== "restoring" && this.bootFetch == null;
  }

  async tick() {
    try {
      if (isHostSupervisionEnabled()) {
        this.maybeAdoptOrphanHost();
        this.maybeForceConcludeBootFetch();
        this.runPostSwapRestore();
        if (this.hostMutationAllowedNow()) {
          await this.processCommand();
        }
        if (this.hostMutationAllowedNow()) {
          this.manageHost();
          this.maybeDisarmPostSwapRollback();
        }
      }
      try {
        this.manageDesktop();
        await this.probeForkRfb();
      } catch (error) {
        log(`desktop tick error: ${String(error)}`);
      }
      try {
        this.sampleCgroups(Date.now());
      } catch (error) {
        log(`cgroup sample error: ${String(error)}`);
      }
      try {
        this.syncBoxScripts();
      } catch (error) {
        log(`box-script sync tick error: ${String(error)}`);
      }
      this.writeStatus();
    } catch (error) {
      log(`tick error: ${String(error)}`);
    } finally {
      this.scheduleTick(TICK_MS);
    }
  }

  hostBundlePresent() {
    return existsSync(HOST_ENTRY);
  }

  runningHostPid() {
    if (this.child != null && this.child.exitCode == null && !this.child.killed) {
      return this.child.pid ?? null;
    }
    if (this.adoptedHost != null && isLiveSandHostPid(this.adoptedHost.pid)) {
      return this.adoptedHost.pid;
    }
    return null;
  }

  hostRunning() {
    if (this.child != null && this.child.exitCode == null && !this.child.killed) {
      return true;
    }
    if (this.adoptedHost != null) {
      if (isLiveSandHostPid(this.adoptedHost.pid)) return true;
      const crashedAtMs = Date.now();
      log(`adopted host pid ${this.adoptedHost.pid} is gone`);
      this.recordAdoptedHostCrash(this.adoptedHost, crashedAtMs);
      this.adoptedHost = null;
    }
    return false;
  }

  maybeAdoptOrphanHost() {
    if (this.adoptedHost != null) return;
    if (this.child != null && this.child.exitCode == null && !this.child.killed) {
      return;
    }
    const discovered = this.readGatewayDiscoveryHost();
    if (discovered == null || discovered.pid === process.pid) return;
    const { pid } = discovered;
    if (this.child != null && pid === this.child.pid) return;
    if (this.stoppedHostPids.has(pid)) {
      if (!isLiveSandHostPid(pid)) this.stoppedHostPids.delete(pid);
      return;
    }
    if (!isLiveSandHostPid(pid)) return;
    log(`adopting live orphan host pid ${pid} (from gateway discovery)`);
    this.adoptedHost = discovered;
  }

  readGatewayDiscoveryHost() {
    const raw = safeRead(join(AGENT_DATA_ROOT, "gateway.json"));
    if (raw == null) return null;
    try {
      const parsed = JSON.parse(raw);
      if (typeof parsed.pid !== "number" || !Number.isInteger(parsed.pid) || parsed.pid <= 0) {
        return null;
      }
      return {
        pid: parsed.pid,
        startedAtMs:
          typeof parsed.startedAt === "number" &&
          Number.isFinite(parsed.startedAt) &&
          parsed.startedAt >= 0
            ? parsed.startedAt
            : undefined,
      };
    } catch {
      return null;
    }
  }

  async processCommand() {
    if (!existsSync(this.commandPath)) return;
    const raw = safeRead(this.commandPath);
    const command = raw == null ? null : parseCommand(raw);
    if (command == null) {
      this.removeCommand("a malformed command");
      return;
    }
    if (!shouldProcessCommand(command, this.commandState(command.id))) {
      this.removeCommand(command.id);
      return;
    }
    let hostWasPaused = false;
    if (commandRequiresIdle(command.kind)) {
      const busyState = await this.probeBusyState();
      const nowMs = Date.now();
      if (
        command.kind === "upgrade" &&
        command.version != null &&
        this.pendingUpgradeVersion !== command.version
      ) {
        this.pendingUpgradeVersion = command.version;
        this.pendingUpgradeDeferredSinceMs = null;
      }
      const deferredForMs =
        this.pendingUpgradeDeferredSinceMs != null ? nowMs - this.pendingUpgradeDeferredSinceMs : 0;
      const action = decideUpgradeAction({
        kind: command.kind,
        busyState,
        deferredForMs,
        forceNow: command.forceNow === true,
      });
      if (action !== "proceed") {
        if (command.kind === "upgrade" && command.version != null) {
          this.pendingUpgradeVersion = command.version;
          if (this.pendingUpgradeDeferredSinceMs == null) {
            this.pendingUpgradeDeferredSinceMs = nowMs;
          }
        }
        if (action === "defer") {
          log(`deferring ${command.kind} ${command.id} (host ${busyState}; retry when idle)`);
          return;
        }
        const forcedReason =
          command.forceNow === true
            ? "on-demand force"
            : `deferred ${deferredForMs}ms ≥ ${MAX_DEFER_MS}ms`;
        const pauseReply = await this.requestHostPause();
        if (pauseReply?.quiescing === true) {
          this.hostPauseRequested = true;
        }
        const runningTurns =
          pauseReply != null && typeof pauseReply.runningTurns === "number"
            ? pauseReply.runningTurns
            : null;
        if (runningTurns !== 0) {
          log(
            `forcing ${command.kind} ${command.id}: ${forcedReason} (host ${busyState}); requested graceful pause (runningTurns=${runningTurns ?? "?"}); retry`,
          );
          return;
        }
        log(
          `forcing ${command.kind} ${command.id}: host paused (turns idle, ${forcedReason}); applying`,
        );
        hostWasPaused = true;
      }
    }
    if (this.postSwapRollback?.phase === "restoring") {
      log(`deferring ${command.kind} ${command.id} (post-swap restore in progress)`);
      return;
    }
    log(`processing ${command.kind} ${command.id}`);
    const outcome = await this.handleCommand(command);
    this.lastCommandId = command.id;
    this.lastCommandKind = command.kind;
    if (outcome === "applied") {
      this.ack(command.id);
    } else {
      log(`command ${command.id} did not apply; left un-acked for re-poke retry`);
      if (hostWasPaused || this.hostPauseRequested) {
        this.stopHost();
        this.lastExitAtMs = null;
        this.restartAttempts = 0;
        log(`command ${command.id} failed after pause; restarting the unchanged host`);
      }
    }
    this.removeCommand(command.id);
  }

  commandState(id) {
    if (this.isAcked(id)) return "acknowledged";
    return this.appliedCommandIds.has(id) ? "applied" : "pending";
  }

  removeCommand(label) {
    try {
      rmSync(this.commandPath, { force: true });
      this.stuckCommand = null;
    } catch (error) {
      if (this.stuckCommand === label) return;
      this.stuckCommand = label;
      log(
        `cannot remove ${this.commandPath} after ${label} (${String(error)}); it stays in the mailbox`,
      );
    }
  }

  async handleCommand(command) {
    switch (command.kind) {
      case "ping":
        return "applied";
      case "restart":
        log("restart: bouncing host onto the on-disk bundle");
        this.stopHost();
        this.lastExitAtMs = null;
        this.restartAttempts = 0;
        return "applied";
      case "upgrade": {
        const mode = normalizeUpgradeMode(command);
        const fromVersion = this.readHostVersion() ?? "unknown";
        const toVersion = command.version ?? "unknown";
        let swapMs = 0;
        if (mode === "bundle") {
          log(`upgrade ${command.id}: applying bundle (host ${fromVersion} -> ${toVersion})`);
          const bundlePath = command.bundlePath ?? STAGED_BUNDLE_PATH;
          const swapStartedAtMs = Date.now();
          const swapResult = await this.applyBundleUpgrade(
            bundlePath,
            command.version,
            command.sha256,
          );
          if (swapResult.ok) {
            this.armPostSwapRollback({ command, fromVersion, toVersion });
          }
          if (!swapResult.ok) {
            this.pendingUpgradeVersion = null;
            this.recordUpgradeFailed({
              commandId: command.id,
              fromVersion,
              toVersion,
              mode,
              reason: command.reason,
              issuedAtMs: command.issuedAtMs,
              swapError: swapResult.reason ?? "bundle-swap-failed",
            });
            log(
              `upgrade ${command.id}: bundle swap FAILED (${swapResult.reason ?? "bundle-swap-failed"}); host left on ${fromVersion} (will retry on re-poke)`,
            );
            return "failed";
          }
          swapMs = Date.now() - swapStartedAtMs;
        } else if (mode === "image") {
          log(`upgrade ${command.id}: image mode — pausing host for recreate`);
        } else {
          log(`upgrade ${command.id}: restart mode — bouncing on-disk bundle`);
        }
        this.stopHost();
        this.lastExitAtMs = null;
        this.restartAttempts = 0;
        this.pendingUpgradeVersion = command.version ?? null;
        if (mode === "bundle" || mode === "restart") {
          this.recordUpgradeApplied({
            commandId: command.id,
            fromVersion,
            toVersion,
            mode,
            reason: command.reason,
            issuedAtMs: command.issuedAtMs,
            swapMs,
          });
        }
        log(`upgrade ${command.id}: host stopped; relaunching onto ${toVersion}`);
        return "applied";
      }
      default:
        return "applied";
    }
  }

  async applyBundleUpgrade(bundlePath, expectedVersion, expectedSha256) {
    if (!isSafeStagedBundlePath(bundlePath)) {
      log(`refusing unsafe staged bundle path: ${bundlePath}`);
      return { ok: false, reason: "unsafe-path" };
    }
    const digest = verifyStagedBundleDigest(bundlePath, expectedSha256);
    if (!digest.ok) {
      const detail =
        digest.actualSha256 == null
          ? ""
          : `: got ${digest.actualSha256.slice(0, 12)}, expected ${expectedSha256.slice(0, 12)}`;
      log(`refusing bundle swap (${digest.reason}${detail}); nothing extracted from ${bundlePath}`);
      safeUnlink(bundlePath);
      return digest;
    }
    return this.swapVerifiedBundle(bundlePath, digest.bytes, expectedVersion, this.stopping.signal);
  }

  swapVerifiedBundle(bundlePath, bundle, expectedVersion, signal) {
    const targets = hostBundleSwapTargets(HOST_DIR);
    if (!bundleSwapTargetsClearOfDataRoot(targets, AGENT_DATA_ROOT)) {
      log(`refusing bundle swap: a swap target overlaps the agent data root ${AGENT_DATA_ROOT}`);
      return { ok: false, reason: "target-overlaps-data-root" };
    }
    return swapHostBundle({
      bundlePath,
      bundle,
      ...targets,
      expectedVersion,
      signal,
    });
  }

  readHostVersion() {
    return readVersionMarker(HOST_VERSION_PATH);
  }

  syncBoxScripts() {
    if (process.env[BOX_SCRIPTS_SYNC_DISABLED_ENV] === "1") return;
    if (this.boxScriptsRetryAtMs != null && Date.now() < this.boxScriptsRetryAtMs) {
      return;
    }
    const bundleVersion = this.readHostVersion();
    if (bundleVersion == null || readVersionMarker(BOX_SCRIPTS_MARKER_PATH) === bundleVersion) {
      return;
    }
    const result = syncBoxScriptsFromDir({
      sourceDir: BOX_SCRIPTS_SOURCE_DIR,
      binDir: BOX_SCRIPTS_BIN_DIR,
      markerPath: BOX_SCRIPTS_MARKER_PATH,
      bundleVersion,
    });
    if (result.status === "failed") {
      this.boxScriptsRetryAtMs = Date.now() + BOX_SCRIPTS_RETRY_BACKOFF_MS;
    } else {
      this.boxScriptsRetryAtMs = null;
    }
  }

  emitBoxTelemetry(event) {
    try {
      appendFileSync(this.telemetryPath, `${JSON.stringify(event)}\n`, "utf8");
    } catch (error) {
      log(`box telemetry append failed: ${String(error)}`);
    }
  }

  maybeStartBootFetch() {
    const localVersion = this.readHostVersion();
    const imageSha = readVersionMarker(this.imageShaPath);
    if (
      !shouldBootFetchHostBundle({
        hostSupervisionEnabled: isHostSupervisionEnabled(),
        autoUpdateOptedOut: isSandBoxAutoUpdateOptedOut(process.env),
        bundlePresent: this.hostBundlePresent(),
        localVersion,
        imageSha,
      })
    ) {
      return;
    }
    if (this.adoptedHost != null) {
      log(
        `boot-fetch: skipped; adopted host pid ${this.adoptedHost.pid} is already running on ${localVersion}, so there is no first launch to hold (upgrades arrive through the command mailbox)`,
      );
      return;
    }
    this.bootFetch = { startedAtMs: Date.now(), fromVersion: localVersion };
    log(
      `boot-fetch: image-baked host ${localVersion}; checking the ${bootFetchChannel()} pointer before the first launch (budget ${this.bootFetchBudgetMs}ms)`,
    );
    void this.runBootFetch(localVersion).catch((error) => {
      this.concludeBootFetch(
        { outcome: "fallback", reason: "unexpected_error" },
        `unexpected error (${String(error)})`,
      );
    });
  }

  async runBootFetch(
    localVersion,
    apply = (bundlePath, version, sha256) => this.applyBundleUpgrade(bundlePath, version, sha256),
    restore = () =>
      restoreHostBundleFromBackup({
        backupDir: hostBundleSwapTargets(HOST_DIR).backupDir,
        hostDir: HOST_DIR,
      }),
  ) {
    const startedAtMs = this.bootFetch?.startedAtMs ?? Date.now();
    const deadlineAtMs = startedAtMs + this.bootFetchBudgetMs;
    const base = bootFetchBaseUrl();
    const channel = bootFetchChannel();
    const rawPointer = await fetchSmallTextWithDeadline(
      `${base}/${bootFetchVersionFileName(channel)}`,
      deadlineAtMs - Date.now(),
    );
    if (this.bootFetch == null) return;
    const target = rawPointer?.trim();
    if (target == null || target.length === 0) {
      return this.concludeBootFetch(
        { outcome: "fallback", reason: "pointer_unreachable" },
        "pointer unreachable",
      );
    }
    if (!BOOT_FETCH_GIT_SHA_REGEX.test(target)) {
      return this.concludeBootFetch(
        { outcome: "fallback", reason: "pointer_malformed" },
        "pointer malformed",
      );
    }
    if (target === localVersion) {
      return this.concludeBootFetch(
        { outcome: "current", reason: "current", toVersion: target },
        `baked host ${localVersion} is current`,
      );
    }
    if (this.isAcked(`upgrade-${target}`)) {
      return this.concludeBootFetch(
        {
          outcome: "fallback",
          reason: "target_vetoed",
          toVersion: target,
        },
        `target ${target} is ack-vetoed`,
      );
    }
    log(
      `boot-fetch: ${channel} pointer ${target} != baked ${localVersion}; fetching its digest, then the bundle, before the first launch`,
    );
    const rawDigest = await fetchSmallTextWithDeadline(
      `${base}/${bootFetchDigestFileName(target)}`,
      deadlineAtMs - Date.now(),
    );
    if (this.bootFetch == null) return;
    if (rawDigest == null) {
      return this.concludeBootFetch(
        {
          outcome: "fallback",
          reason: "digest_missing",
          toVersion: target,
        },
        `digest for ${target} missing or unreachable; nothing downloaded; baked host launches`,
      );
    }
    const expectedSha256 = parseHostBundleDigest(rawDigest);
    if (expectedSha256 == null) {
      return this.concludeBootFetch(
        {
          outcome: "fallback",
          reason: "digest_malformed",
          toVersion: target,
        },
        `digest for ${target} malformed; nothing downloaded; baked host launches`,
      );
    }
    const downloaded = await downloadFileWithDeadline(
      `${base}/sand-host-bundle-${target}.tgz`,
      STAGED_BUNDLE_PATH,
      deadlineAtMs - Date.now(),
      expectedSha256,
    );
    if (this.bootFetch == null) return;
    if (!downloaded.ok) {
      if (downloaded.reason === "digest_mismatch") {
        return this.concludeBootFetch(
          {
            outcome: "fallback",
            reason: "digest_mismatch",
            toVersion: target,
          },
          `downloaded ${target} failed its sha256 check (digest_mismatch: got ${downloaded.actualSha256.slice(0, 12)}, expected ${expectedSha256.slice(0, 12)}); discarded before extraction; baked host launches`,
        );
      }
      return this.concludeBootFetch(
        {
          outcome: "fallback",
          reason: "download_failed",
          toVersion: target,
        },
        `download of ${target} failed/timed out`,
      );
    }
    const swapStartedAtMs = Date.now();
    this.bootFetch.swapInFlight = true;
    const swapResult = await apply(STAGED_BUNDLE_PATH, target, expectedSha256);
    if (this.bootFetch == null) return;
    if (!swapResult.ok) {
      safeUnlink(STAGED_BUNDLE_PATH);
      if (swapResult.mutated === true) {
        const restored = restore();
        return this.concludeBootFetch(
          {
            outcome: restored.ok ? "fallback" : "restore_failed",
            reason: restored.ok ? "swap_failed_restored" : "restore_failed",
            toVersion: target,
          },
          `swap failed post-commit (${swapResult.reason}); baked host ${
            restored.ok ? "restored" : `NOT restored (${restored.reason})`
          }`,
        );
      }
      return this.concludeBootFetch(
        {
          outcome: "fallback",
          reason: "swap_refused",
          toVersion: target,
        },
        `swap refused (${swapResult.reason}); baked host launches`,
      );
    }
    const swapMs = Date.now() - swapStartedAtMs;
    const commandId = `upgrade-${target}`;
    this.armPostSwapRollback({
      command: { id: commandId, reason: "boot-fetch", issuedAtMs: startedAtMs },
      fromVersion: localVersion ?? "unknown",
      toVersion: target,
    });
    this.ack(commandId);
    this.recordUpgradeApplied({
      commandId,
      fromVersion: localVersion ?? "unknown",
      toVersion: target,
      mode: "bundle",
      reason: "boot-fetch",
      issuedAtMs: startedAtMs,
      swapMs,
    });
    this.concludeBootFetch(
      {
        outcome: "applied",
        reason: "applied",
        toVersion: target,
        swapMs,
      },
      `applied ${localVersion} -> ${target} (swap ${swapMs}ms); launching the fetched host`,
    );
  }

  concludeBootFetch(report, description) {
    if (this.bootFetch == null) return;
    const heldForMs = Date.now() - this.bootFetch.startedAtMs;
    this.emitBoxTelemetry({
      kind: "host_boot_fetch",
      outcome: report.outcome,
      reason: report.reason,
      durationMs: Math.max(0, heldForMs),
      swapMs: report.swapMs,
      fromVersion: this.bootFetch.fromVersion,
      toVersion: report.toVersion,
    });
    this.bootFetch = null;
    log(`boot-fetch: ${description} (held first launch ${heldForMs}ms)`);
    this.scheduleTick(0);
  }

  maybeForceConcludeBootFetch() {
    if (this.bootFetch == null || this.bootFetch.swapInFlight === true) return;
    if (Date.now() - this.bootFetch.startedAtMs > this.bootFetchBudgetMs + TICK_MS) {
      this.concludeBootFetch(
        { outcome: "fallback", reason: "budget_exceeded" },
        "budget exceeded (backstop); baked host launches",
      );
    }
  }

  armPostSwapRollback({ command, fromVersion, toVersion }) {
    const { backupDir } = hostBundleSwapTargets(HOST_DIR);
    if (!existsSync(join(backupDir, "host-main.cjs"))) {
      this.postSwapRollback = null;
      return;
    }
    this.postSwapRollback = {
      phase: "watching",
      backupDir,
      fromVersion,
      toVersion,
      appliedAtMs: Date.now(),
      quickExits: 0,
      restoreAttempts: 0,
      commandId: command.id,
      reason: command.reason,
      issuedAtMs: command.issuedAtMs,
    };
  }

  notePostSwapHostExit(launchedAtMs, uptimeMs, stopWasRequested = false) {
    if (stopWasRequested) return;
    const armed = this.postSwapRollback;
    if (armed == null || armed.phase !== "watching") return;
    const action = decidePostSwapAction({
      armed: true,
      exitLaunchedAtMs: launchedAtMs,
      appliedAtMs: armed.appliedAtMs,
      uptimeMs,
      quickExits: armed.quickExits,
    });
    if (action === "none") return;
    if (action === "healthy") {
      this.clearPostSwapRollback(`host ran ${uptimeMs}ms on ${armed.toVersion} (healthy)`);
      return;
    }
    armed.quickExits += 1;
    if (action === "count") {
      log(
        `post-swap watch: host exited after ${uptimeMs}ms on ${armed.toVersion} (quick exit ${armed.quickExits}/${POST_SWAP_MAX_QUICK_EXITS})`,
      );
      return;
    }
    log(
      `post-swap watch: host crash-looping on ${armed.toVersion} (${armed.quickExits} quick exits); rolling back to ${armed.fromVersion}`,
    );
    armed.phase = "restoring";
    this.stopHost();
  }

  maybeDisarmPostSwapRollback() {
    const armed = this.postSwapRollback;
    if (armed == null || armed.phase !== "watching") return;
    if (
      this.hostRunning() &&
      this.lastLaunchAtMs != null &&
      this.lastLaunchAtMs >= armed.appliedAtMs &&
      Date.now() - this.lastLaunchAtMs >= POST_SWAP_HEALTHY_UPTIME_MS
    ) {
      this.clearPostSwapRollback(
        `host up ${Date.now() - this.lastLaunchAtMs}ms on ${armed.toVersion} (healthy)`,
      );
    }
  }

  clearPostSwapRollback(why) {
    const armed = this.postSwapRollback;
    if (armed == null) return;
    this.postSwapRollback = null;
    log(`post-swap watch disarmed: ${why}`);
    safeRmRecursive(armed.backupDir);
  }

  runPostSwapRestore(restore = restoreHostBundleFromBackup) {
    const armed = this.postSwapRollback;
    if (armed == null || armed.phase !== "restoring") return;
    const result = restore({
      backupDir: armed.backupDir,
      hostDir: HOST_DIR,
    });
    if (result.ok) {
      this.postSwapRollback = null;
      this.pendingUpgradeVersion = null;
      this.pendingUpgradeDeferredSinceMs = null;
      this.lastExitAtMs = null;
      this.restartAttempts = 0;
      this.recordUpgradeFailed({
        commandId: armed.commandId,
        fromVersion: armed.fromVersion,
        toVersion: armed.toVersion,
        mode: "bundle",
        reason: armed.reason,
        issuedAtMs: armed.issuedAtMs,
        swapError: POST_SWAP_CRASH_LOOP_ERROR_CLASS,
      });
      log(
        `post-swap rollback complete: host restored to ${armed.fromVersion}; ${armed.toVersion} will not be retried on this box (command already acked)`,
      );
      return;
    }
    armed.restoreAttempts += 1;
    const giveUp =
      result.reason === "backup-missing" || armed.restoreAttempts >= POST_SWAP_MAX_RESTORE_ATTEMPTS;
    if (!giveUp) {
      log(
        `post-swap rollback attempt ${armed.restoreAttempts}/${POST_SWAP_MAX_RESTORE_ATTEMPTS} failed (${result.reason}); backup kept, retrying next tick`,
      );
      return;
    }
    this.postSwapRollback = null;
    this.pendingUpgradeVersion = null;
    this.pendingUpgradeDeferredSinceMs = null;
    this.recordUpgradeFailed({
      commandId: armed.commandId,
      fromVersion: armed.fromVersion,
      toVersion: armed.toVersion,
      mode: "bundle",
      reason: armed.reason,
      issuedAtMs: armed.issuedAtMs,
      swapError: POST_SWAP_ROLLBACK_FAILED_ERROR_CLASS,
    });
    log(
      `post-swap rollback GAVE UP (${result.reason}, ${armed.restoreAttempts} attempt(s)); host left as-is on ${armed.toVersion}, relaunch backoff continues`,
    );
  }

  manageHost() {
    this.maybeAdoptOrphanHost();
    const action = decideHostAction({
      bundlePresent: this.hostBundlePresent(),
      hostRunning: this.hostRunning(),
      isBusy: false,
      lastExitAtMs: this.lastExitAtMs,
      restartAttempts: this.restartAttempts,
      nowMs: Date.now(),
    });
    if (action === "launch" || action === "restart") {
      this.launchHost();
    }
  }

  readForkRfbTargets(specs) {
    const byId = new Map(specs.map((spec) => [spec.id, spec]));
    const websockify = byId.get("shared/fork-websockify");
    if (websockify == null) return [];
    let entries;
    try {
      entries = readdirSync(this.forkRfbTokenDir, { withFileTypes: true });
    } catch {
      return [];
    }
    return entries
      .flatMap((entry) => {
        if (!entry.isFile()) return [];
        const raw = safeRead(join(this.forkRfbTokenDir, entry.name));
        const route = raw == null ? null : parseForkRfbTokenRecord(entry.name, raw);
        if (route == null) return [];
        const x11vnc = byId.get(`d${route.display}/x11vnc`);
        return x11vnc == null ? [] : [{ ...route, websockify, x11vnc }];
      })
      .sort((a, b) => a.display - b.display);
  }

  async probeForkRfb() {
    if (!isDesktopSupervisionEnabled()) return;
    const targets = this.readForkRfbTargets(this.readDesktopSpecs());
    const activeDisplays = new Set(targets.map((target) => target.display));
    for (const display of this.forkRfbFailures.keys()) {
      if (!activeDisplays.has(display)) this.forkRfbFailures.delete(display);
    }
    if (targets.length === 0) return;
    const target = targets[this.forkRfbProbeCursor % targets.length];
    const advance = () => {
      this.forkRfbProbeCursor = (this.forkRfbProbeCursor + 1) % targets.length;
    };
    const x11vncIdentity = this.readDesktopIdentity(target.x11vnc);
    const websockifyIdentity = this.readDesktopIdentity(target.websockify);
    if (x11vncIdentity == null || websockifyIdentity == null) {
      advance();
      return;
    }
    const identity = `${x11vncIdentity.pid}:${x11vncIdentity.startTime}/${websockifyIdentity.pid}:${websockifyIdentity.startTime}`;
    const result = await probeForkRfbChain({
      vncPort: target.vncPort,
      websocketPort: this.forkRfbWebsocketPort,
      token: target.token,
      timeoutMs: this.forkRfbProbeTimeoutMs,
    });
    const currentX11vncIdentity = this.readDesktopIdentity(target.x11vnc);
    const currentWebsockifyIdentity = this.readDesktopIdentity(target.websockify);
    if (
      currentX11vncIdentity?.pid !== x11vncIdentity.pid ||
      currentX11vncIdentity?.startTime !== x11vncIdentity.startTime ||
      currentWebsockifyIdentity?.pid !== websockifyIdentity.pid ||
      currentWebsockifyIdentity?.startTime !== websockifyIdentity.startTime
    ) {
      this.forkRfbFailures.delete(target.display);
      advance();
      return;
    }
    if (result === "healthy") {
      this.forkRfbFailures.delete(target.display);
      advance();
      return;
    }
    const previous = this.forkRfbFailures.get(target.display);
    const failures =
      previous?.identity === identity && previous.result === result ? previous.failures + 1 : 1;
    this.forkRfbFailures.set(target.display, { failures, identity, result });
    if (failures < this.forkRfbFailureThreshold) return;
    const spec = result === "x11vnc" ? target.x11vnc : target.websockify;
    const expectedIdentity = result === "x11vnc" ? x11vncIdentity : websockifyIdentity;
    const port = result === "x11vnc" ? target.vncPort : this.forkRfbWebsocketPort;
    if (this.requestForkRfbRestart(spec, expectedIdentity, failures, result, port)) {
      this.forkRfbFailures.delete(target.display);
    }
    advance();
  }

  requestForkRfbRestart(spec, expectedIdentity, failures, name, port) {
    const state = this.getDesktopState(spec.id);
    if (state.pid !== expectedIdentity.pid) return false;
    if (!this.wouldRestartAfterExit(state)) return false;
    const currentIdentity = this.readDesktopIdentity(spec);
    if (
      currentIdentity?.pid !== expectedIdentity.pid ||
      currentIdentity?.startTime !== expectedIdentity.startTime ||
      !isExpectedForkRfbProcess({
        ...expectedIdentity,
        name,
        port,
        procRoot: this.procRoot,
      }) ||
      !existsSync(spec.descriptorPath)
    ) {
      return false;
    }
    try {
      process.kill(expectedIdentity.pid, "SIGTERM");
    } catch {
      return false;
    }
    const escalate = setTimeout(() => {
      if (
        this.stopping.signal.aborted ||
        !isExpectedForkRfbProcess({
          ...expectedIdentity,
          name,
          port,
          procRoot: this.procRoot,
        })
      ) {
        return;
      }
      try {
        process.kill(expectedIdentity.pid, "SIGKILL");
      } catch {}
    }, FORK_RFB_TERMINATE_GRACE_MS);
    if (typeof escalate.unref === "function") escalate.unref();
    log(`desktop: RFB readiness failed ${failures} consecutive times; restarting ${spec.id}`);
    return true;
  }

  reapDisabledCompositor(spec) {
    let targets;
    try {
      targets = findCompositorReapTargets({
        display: spec.env?.DISPLAY,
        procRoot: this.procRoot,
      });
    } catch {
      return;
    }
    if (targets.length === 0) return;
    log(
      `desktop: ${spec.id} give-up: reaping surviving compositor pid(s) ${targets.join(
        ", ",
      )} so no orphan keeps compositing after disable`,
    );
    const pinned = [];
    for (const pid of targets) {
      const startTime = readProcStartTime(pid, this.procRoot);
      try {
        process.kill(pid, "SIGTERM");
      } catch {
        continue;
      }
      if (startTime != null) pinned.push({ pid, startTime });
    }
    if (pinned.length === 0) return;
    const escalate = setTimeout(() => {
      if (this.stopping.signal.aborted) return;
      let survivors;
      try {
        survivors = new Set(
          findCompositorReapTargets({
            display: spec.env?.DISPLAY,
            procRoot: this.procRoot,
          }),
        );
      } catch {
        return;
      }
      for (const { pid, startTime } of pinned) {
        if (!survivors.has(pid)) continue;
        if (readProcStartTime(pid, this.procRoot) !== startTime) continue;
        try {
          process.kill(pid, "SIGKILL");
        } catch {}
      }
    }, this.compositorReapEscalateMs);
    if (typeof escalate.unref === "function") escalate.unref();
  }

  restartDesktopComponent(spec) {
    if (!existsSync(spec.descriptorPath)) return null;
    let stdio = "ignore";
    let logFd = null;
    const boundedLog = isBoundedDesktopLogSpec(spec);
    if (boundedLog) {
      const cursor = desktopLogCursor(spec.logFile);
      this.desktopLogOffset.set(spec.id, cursor.size);
      this.desktopLogIdentity.set(spec.id, cursor.identity);
    } else if (spec.logFile != null) {
      try {
        logFd = openSync(spec.logFile, "a");
        stdio = ["ignore", logFd, logFd];
        try {
          const info = fstatSync(logFd);
          this.desktopLogOffset.set(spec.id, info.size);
          this.desktopLogIdentity.set(spec.id, desktopLogIdentity(info));
        } catch {
          this.desktopLogOffset.set(spec.id, 0);
          this.desktopLogIdentity.set(spec.id, null);
        }
      } catch {
        stdio = "ignore";
        this.desktopLogOffset.set(spec.id, 0);
        this.desktopLogIdentity.set(spec.id, null);
      }
    }
    try {
      const child = spawn(spec.argv[0], spec.argv.slice(1), {
        cwd: "/",
        detached: true,
        stdio,
        env: { ...process.env, ...spec.env },
      });
      child.on("error", (error) => log(`desktop: ${spec.id} spawn error: ${String(error)}`));
      const childPid = child.pid ?? null;
      child.on("exit", (code, signal) => {
        this.desktopExitInfo.set(spec.id, {
          pid: childPid,
          code: typeof code === "number" ? code : null,
          signal: typeof signal === "string" ? signal : null,
          atMs: Date.now(),
        });
      });
      child.unref();
      const pid = childPid;
      if (pid != null) this.placeInCgroup(CGROUP_INTERACTIVE, pid);
      if (pid != null) {
        try {
          writeFileSync(`/proc/${pid}/oom_score_adj`, "0");
        } catch {}
      }
      if (pid != null) atomicWrite(spec.pidFile, String(pid));
      return pid;
    } catch (error) {
      log(`desktop: failed to relaunch ${spec.id}: ${String(error)}`);
      return null;
    } finally {
      if (logFd != null) {
        try {
          closeSync(logFd);
        } catch {}
      }
    }
  }

  launchHost() {
    if (this.hostRunning()) return;
    this.adoptedHost = null;
    log(`launching host: node ${HOST_ENTRY}`);
    const hostLogStdio = this.openHostLog();
    const launchedAtMs = Date.now();
    const launchedAtMonotonicMs = performance.now();
    this.lastLaunchAtMs = launchedAtMs;
    const child = spawn(process.execPath, [HOST_NODE_WARNING_FLAG, HOST_ENTRY], {
      cwd: HOST_DIR,
      detached: true,
      stdio: hostLogStdio === null ? "inherit" : ["ignore", hostLogStdio, hostLogStdio],
      env: {
        ...process.env,
        SAND_PACKAGED: "1",
        SAND_DATA_ROOT: AGENT_DATA_ROOT,
        SAND_HOST_IN_BOX: "1",
        ...(hostLogStdio === null ? {} : { SAND_HOST_LOG_FILE: HOST_LOG_PATH }),
      },
    });
    this.child = child;
    if (child.pid != null) this.stoppedHostPids.delete(child.pid);
    for (const pid of this.stoppedHostPids) {
      if (!isLiveSandHostPid(pid)) this.stoppedHostPids.delete(pid);
    }
    // oom_score_adj -1000 disables OOM killing for the task entirely, so the host takes -998 and stays killable when it is the process holding the memory: https://man7.org/linux/man-pages/man5/proc_pid_oom_score_adj.5.html
    try {
      writeFileSync(`/proc/${child.pid}/oom_score_adj`, "-998");
    } catch {}
    this.hostPauseRequested = false;
    this.restartAttempts += 1;
    child.on("exit", (code, signal) => {
      const nowMs = Date.now();
      this.noteHostProcessExit({
        child,
        code,
        signal,
        startedAtMs: launchedAtMs,
        crashedAtMs: nowMs,
        uptimeMs: Math.max(0, performance.now() - launchedAtMonotonicMs),
      });
      this.lastExitAtMs = nowMs;
      log(`host exited (code ${code ?? "null"}, signal ${signal ?? "null"})`);
      if (this.child === child) this.child = null;
      try {
        this.notePostSwapHostExit(launchedAtMs, nowMs - launchedAtMs, child.killed);
      } catch (error) {
        log(`post-swap watch error: ${String(error)}`);
      }
    });
    child.on("error", (error) => {
      this.lastExitAtMs = Date.now();
      log(`host spawn error: ${String(error)}`);
      if (this.child === child) this.child = null;
    });
  }

  noteHostProcessExit({ child, code, signal, startedAtMs, crashedAtMs, uptimeMs }) {
    if (this.expectedHostExits.delete(child)) return false;
    return this.recordHostCrashMarker(
      classifyHostProcessExit({
        code,
        signal,
        startedAtMs,
        crashedAtMs,
        uptimeMs,
      }),
    );
  }

  recordHostCrashMarker(marker) {
    if (existsSync(this.crashMarkerPath)) return false;
    const partPath = `${this.crashMarkerPath}.part`;
    try {
      writeFileSync(partPath, JSON.stringify(marker));
      renameSync(partPath, this.crashMarkerPath);
      return true;
    } catch (error) {
      log(
        `host crash marker write failed (error_class=${classifyHostCrashMarkerWriteError(error)})`,
      );
      return false;
    }
  }

  recordAdoptedHostCrash(adoptedHost, crashedAtMs) {
    this.recordHostCrashMarker(
      classifyHostProcessExit({
        code: null,
        signal: null,
        startedAtMs: adoptedHost.startedAtMs,
        crashedAtMs,
        uptimeMs:
          adoptedHost.startedAtMs === undefined
            ? undefined
            : Math.max(0, crashedAtMs - adoptedHost.startedAtMs),
      }),
    );
    this.lastExitAtMs = crashedAtMs;
  }

  openHostLog() {
    if (this.hostLogFd != null) return this.hostLogFd;
    try {
      this.hostLogFd = openSync(HOST_LOG_PATH, "a");
      return this.hostLogFd;
    } catch (error) {
      log(`host log open failed (${String(error)}); inheriting stdio`);
      this.hostLogFd = null;
      return null;
    }
  }

  stopHost() {
    if (this.adoptedHost != null) {
      const adoptedHost = this.adoptedHost;
      log(`stopping adopted host pid ${adoptedHost.pid} (SIGTERM)`);
      let exitedBeforeStop = false;
      try {
        process.kill(adoptedHost.pid, "SIGTERM");
      } catch {
        if (isLiveSandHostPid(adoptedHost.pid)) return;
        exitedBeforeStop = true;
        this.recordAdoptedHostCrash(adoptedHost, Date.now());
      }
      this.stoppedHostPids.add(adoptedHost.pid);
      this.adoptedHost = null;
      if (!exitedBeforeStop) this.lastExitAtMs = Date.now();
    }
    if (this.child == null) return;
    log("stopping host (SIGTERM)");
    if (this.child.pid != null) this.stoppedHostPids.add(this.child.pid);
    this.expectedHostExits.add(this.child);
    try {
      if (!this.child.kill("SIGTERM")) {
        this.expectedHostExits.delete(this.child);
      }
    } catch {
      this.expectedHostExits.delete(this.child);
    }
  }

  async probeBusyState() {
    if (!this.hostRunning()) return "no-host";
    const endpoint = this.readGatewayEndpoint();
    if (endpoint == null) return "unknown";
    const health = await fetchHealth(endpoint.port);
    if (health == null) return "unknown";
    return health.isBusy === true ? "busy" : "idle";
  }

  readGatewayEndpoint() {
    const raw = safeRead(join(AGENT_DATA_ROOT, "gateway.json"));
    if (raw == null) return null;
    try {
      return parseGatewayEndpoint(JSON.parse(raw));
    } catch {
      return null;
    }
  }

  async requestHostPause() {
    if (!this.hostRunning()) return null;
    const endpoint = this.readGatewayEndpoint();
    if (endpoint == null) return null;
    return await postPrepareUpgrade(endpoint).catch(() => null);
  }

  isAcked(id) {
    return existsSync(join(this.acksDir, ackFilename(id)));
  }

  ack(id) {
    this.appliedCommandIds.add(id);
    try {
      writeFileSync(join(this.acksDir, ackFilename(id)), String(Date.now()));
    } catch (error) {
      log(
        `cannot write the ack for ${id} in ${this.acksDir} (${String(error)}); only this supervisor process remembers it as applied`,
      );
    }
  }

  writeStatus() {
    const hostVersion = this.readHostVersion();
    if (this.pendingUpgradeVersion != null && hostVersion === this.pendingUpgradeVersion) {
      this.pendingUpgradeVersion = null;
      this.pendingUpgradeDeferredSinceMs = null;
    }
    const status = buildStatus({
      nowMs: Date.now(),
      hostBundlePresent: this.hostBundlePresent(),
      hostRunning: this.hostRunning(),
      hostVersion,
      pendingUpgradeVersion: this.pendingUpgradeVersion,
      lastCommandId: this.lastCommandId,
      lastCommandKind: this.lastCommandKind,
    });
    atomicWrite(STATUS_PATH, JSON.stringify(status));
  }

  recordUpgradeApplied({ commandId, fromVersion, toVersion, mode, reason, issuedAtMs, swapMs }) {
    try {
      atomicWrite(
        this.markerPath,
        JSON.stringify({
          outcome: "applied",
          commandId,
          fromVersion,
          toVersion,
          mode,
          reason,
          issuedAtMs,
          appliedAtMs: Date.now(),
          swapMs,
        }),
      );
    } catch {}
  }

  recordUpgradeFailed({ commandId, fromVersion, toVersion, mode, reason, issuedAtMs, swapError }) {
    try {
      atomicWrite(
        this.markerPath,
        JSON.stringify({
          outcome: "failed",
          commandId,
          fromVersion,
          toVersion,
          mode,
          reason,
          issuedAtMs,
          failedAtMs: Date.now(),
          swapError,
        }),
      );
    } catch {}
  }
}

function ackFilename(id) {
  return id.replace(/[^a-zA-Z0-9_.-]/g, "_");
}

function ensureDir(dir) {
  try {
    mkdirSync(dir, { recursive: true });
    return null;
  } catch (error) {
    return error;
  }
}

function safeRead(path) {
  try {
    return readFileSync(path, "utf8");
  } catch {
    return null;
  }
}

function isLiveSandHostPid(pid) {
  try {
    process.kill(pid, 0);
  } catch {
    return false;
  }
  const cmdline = safeRead(`/proc/${pid}/cmdline`);
  return cmdline != null && cmdline.includes("host-main.cjs");
}

function readVersionMarker(path) {
  const raw = safeRead(path);
  if (raw == null) return null;
  const trimmed = raw.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function safeUnlink(path) {
  try {
    rmSync(path, { force: true });
  } catch {}
}

function safeRmRecursive(path) {
  try {
    rmSync(path, { recursive: true, force: true });
  } catch {}
}

function atomicWrite(path, contents) {
  const partPath = `${path}.part`;
  try {
    writeFileSync(partPath, contents);
    renameSync(partPath, path);
  } catch {}
}

function fetchHealth(port) {
  return new Promise((resolve) => {
    const req = http.get(
      { host: "127.0.0.1", port, path: "/health", timeout: HEALTH_TIMEOUT_MS },
      (res) => {
        let body = "";
        res.on("data", (chunk) => (body += chunk));
        res.on("end", () => {
          try {
            resolve(JSON.parse(body));
          } catch {
            resolve(null);
          }
        });
      },
    );
    req.on("error", () => resolve(null));
    req.on("timeout", () => {
      req.destroy();
      resolve(null);
    });
  });
}

function requestModuleFor(url) {
  return url.startsWith("http://") ? http : https;
}

export function bootFetchBaseUrl(env = process.env) {
  const override = env[HOST_BUNDLE_BASE_URL_ENV]?.trim();
  if (override != null && override.length > 0) {
    return override.replace(/\/+$/, "");
  }
  return HOST_BUNDLE_DEFAULT_BASE_URL;
}

export function bootFetchChannel(env = process.env) {
  const stamped = env[HOST_BUNDLE_CHANNEL_ENV]?.trim();
  return HOST_BUNDLE_CHANNELS.includes(stamped) ? stamped : HOST_BUNDLE_DEFAULT_CHANNEL;
}

export function bootFetchVersionFileName(channel) {
  return `sand-host-bundle-${channel}.version`;
}

export function bootFetchDigestFileName(version) {
  return `sand-host-bundle-${version}.tgz.sha256`;
}

// Node's http.request timeout is a socket inactivity timeout, not a deadline, so a separate timer bounds the whole fetch: https://nodejs.org/api/http.html#httprequesturl-options-callback
function fetchSmallTextWithDeadline(url, timeoutMs) {
  return new Promise((resolve) => {
    if (!(timeoutMs > 0)) return resolve(null);
    let settled = false;
    const settle = (value) => {
      if (settled) return;
      settled = true;
      clearTimeout(deadline);
      resolve(value);
    };
    const req = requestModuleFor(url).get(url, { timeout: timeoutMs }, (res) => {
      if (res.statusCode !== 200) {
        res.resume();
        return settle(null);
      }
      let body = "";
      res.on("data", (chunk) => {
        body += chunk;
        if (body.length > 4_096) {
          req.destroy();
          settle(null);
        }
      });
      res.on("end", () => settle(body));
      res.on("error", () => settle(null));
    });
    const deadline = setTimeout(() => {
      req.destroy();
      settle(null);
    }, timeoutMs);
    req.on("error", () => settle(null));
    req.on("timeout", () => {
      req.destroy();
      settle(null);
    });
  });
}

function downloadFileWithDeadline(url, destPath, timeoutMs, expectedSha256) {
  return new Promise((resolve) => {
    const partPath = `${destPath}.part`;
    const hash = createHash("sha256");
    let settled = false;
    const settle = (result) => {
      if (settled) return;
      settled = true;
      clearTimeout(deadline);
      if (!result.ok) safeUnlink(partPath);
      resolve(result);
    };
    const fail = () => settle({ ok: false, reason: "download_failed" });
    if (!(timeoutMs > 0)) {
      return resolve({ ok: false, reason: "download_failed" });
    }
    const req = requestModuleFor(url).get(url, { timeout: timeoutMs }, (res) => {
      if (res.statusCode !== 200) {
        res.resume();
        return fail();
      }
      let file;
      try {
        file = createWriteStream(partPath);
      } catch {
        return fail();
      }
      const failAndDrop = () => {
        try {
          file.destroy();
        } catch {}
        fail();
      };
      res.on("error", failAndDrop);
      file.on("error", failAndDrop);
      res.on("data", (chunk) => hash.update(chunk));
      res.pipe(file);
      file.on("finish", () => {
        if (settled) return;
        file.close(() => {
          if (settled) return;
          const actualSha256 = hash.digest("hex");
          if (actualSha256 !== expectedSha256) {
            return settle({ ok: false, reason: "digest_mismatch", actualSha256 });
          }
          try {
            renameSync(partPath, destPath);
            settle({ ok: true });
          } catch {
            fail();
          }
        });
      });
    });
    const deadline = setTimeout(() => {
      req.destroy();
      fail();
    }, timeoutMs);
    req.on("error", fail);
    req.on("timeout", () => {
      req.destroy();
      fail();
    });
  });
}

export function postPrepareUpgrade({ port, token }) {
  return new Promise((resolve) => {
    const req = http.request(
      {
        host: "127.0.0.1",
        port,
        path: PREPARE_UPGRADE_PATH,
        method: "POST",
        timeout: HEALTH_TIMEOUT_MS,
        headers: token != null ? { authorization: `Bearer ${token}` } : {},
      },
      (res) => {
        let body = "";
        res.on("data", (chunk) => (body += chunk));
        res.on("end", () => {
          try {
            resolve(JSON.parse(body));
          } catch {
            resolve(null);
          }
        });
      },
    );
    req.on("error", () => resolve(null));
    req.on("timeout", () => {
      req.destroy();
      resolve(null);
    });
    req.end();
  });
}

const isMain = process.argv[1] != null && fileURLToPath(import.meta.url) === process.argv[1];
export async function handOffOnSignal(supervisor, signal) {
  const hostPid = supervisor.runningHostPid();
  log(
    `received ${signal}, stopping; ${
      hostPid == null
        ? "no host is running"
        : `leaving host pid ${hostPid} running for the next supervisor to adopt`
    }`,
  );
  await supervisor.stop();
}

if (isMain) {
  const supervisor = new Supervisor();
  supervisor.start();
  const shutdown = async (signal) => {
    await handOffOnSignal(supervisor, signal);
    process.exit(0);
  };
  process.on("SIGTERM", () => void shutdown("SIGTERM"));
  process.on("SIGINT", () => void shutdown("SIGINT"));
}
