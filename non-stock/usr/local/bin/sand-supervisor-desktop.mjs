import {
  existsSync,
  fstatSync,
  openSync,
  readdirSync,
  readSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { join } from "node:path";
import {
  CGROUP_ROOT,
  CGROUP_INTERACTIVE,
  CGROUP_AGENT,
  CGROUP_SAMPLE_MS,
  parseCgroupCpuStat,
  parseCgroupPressure,
  formatCgroupCpuSample,
} from "./sand-supervisor-cgroup-accounting.mjs";
import { nextBackoffMs } from "./sand-supervisor-host-policy.mjs";
import {
  DESKTOP_DIR,
  DESKTOP_HEALTH_PATH,
  DESKTOP_BACKOFF_BASE_MS,
  DESKTOP_BACKOFF_MAX_MS,
  DESKTOP_RESTART_WINDOW_MS,
  DESKTOP_RESTART_MAX_IN_WINDOW,
  DESKTOP_COMPOSITOR_MAX_CRASHLOOPS,
  DESKTOP_COMPOSITOR_NAME,
  DESKTOP_DOCK_NAME,
  DESKTOP_LOG_TAIL_BYTES,
  DESKTOP_LOG_TAIL_LINES,
  isDesktopSupervisionEnabled,
  log,
} from "./sand-supervisor-contract.mjs";
import { isPidAlive } from "./sand-supervisor-process-identity.mjs";

export function countRestartsInWindow(timestamps, nowMs, windowMs) {
  const cutoff = nowMs - windowMs;
  return timestamps.reduce((n, t) => (t >= cutoff ? n + 1 : n), 0);
}

export function decideDesktopComponentAction({
  supervisionEnabled,
  running,
  lastExitAtMs,
  restartsInWindow,
  maxRestarts,
  nowMs,
  baseBackoffMs = DESKTOP_BACKOFF_BASE_MS,
  maxBackoffMs = DESKTOP_BACKOFF_MAX_MS,
}) {
  if (!supervisionEnabled) return "noop";
  if (running) return "noop";
  if (restartsInWindow >= maxRestarts) return "crashloop";
  if (lastExitAtMs == null) return "restart";
  const backoff = nextBackoffMs(restartsInWindow, baseBackoffMs, maxBackoffMs);
  return nowMs - lastExitAtMs >= backoff ? "restart" : "wait";
}

export function shouldDisableCompositor({
  name,
  crashloopEpisodes,
  maxEpisodes = DESKTOP_COMPOSITOR_MAX_CRASHLOOPS,
}) {
  if (name !== DESKTOP_COMPOSITOR_NAME) return false;
  return Number.isFinite(crashloopEpisodes) && crashloopEpisodes >= maxEpisodes;
}

export function classifyDesktopDownReason({ exitCode = null, signal = null, logTail = "" } = {}) {
  const tail = typeof logTail === "string" ? logTail : "";
  if (
    /could ?n[o']?t obtain listening port|address already in use|trouble binding|failed to bind/i.test(
      tail,
    )
  ) {
    return "port-in-use";
  }
  if (
    /server is already active|already active for display|cannot establish any listening sockets/i.test(
      tail,
    )
  ) {
    return "x-server-active";
  }
  if (
    /another composite manager is already running|another (?:window manager|wm)\b.*\b(?:is already running|is running)/i.test(
      tail,
    )
  ) {
    return "already-running";
  }
  if (
    /can(?:no|')t open display|cannot open display|unable to open display|couldn't open display|no display/i.test(
      tail,
    )
  ) {
    return "no-display";
  }
  if (/out of memory|cannot allocate memory|bad_alloc|std::bad_alloc/i.test(tail)) {
    return "oom";
  }
  if (/\bglx\b|opengl|\begl\b|glamor|failed to initialize.*backend|render.*backend/i.test(tail)) {
    return "glx";
  }
  if (
    /fatal io error|lost connection to x server|connection to .*(broken|reset)|\bxio\b/i.test(tail)
  ) {
    return "x-io-error";
  }
  if (typeof signal === "string" && /^SIG[A-Z0-9]+$/.test(signal)) {
    return `signal-${signal}`;
  }
  if (/fatal server error/i.test(tail)) {
    return "fatal-server";
  }
  if (Number.isInteger(exitCode) && exitCode !== 0) {
    return `exit-${exitCode}`;
  }
  return "unknown";
}

export function desktopStateSignature(components) {
  return components
    .map(
      (c) =>
        `${c.group}/${c.name}:${c.up ? 1 : 0}:${c.crashloop ? 1 : 0}:${
          c.up ? "" : (c.downReason ?? "")
        }`,
    )
    .sort()
    .join("|");
}

export function buildDesktopHealthSnapshot({ supervisionEnabled, components, nowMs, revision }) {
  const total = components.length;
  const up = components.filter((c) => c.up === true).length;
  const crashlooping = components.filter((c) => c.crashloop === true).length;
  const restartsInWindow = components.reduce(
    (n, c) => n + (Number.isFinite(c.restartsInWindow) ? c.restartsInWindow : 0),
    0,
  );
  return {
    updatedAtMs: nowMs,
    revision,
    supervisionEnabled,
    total,
    up,
    down: total - up,
    crashlooping,
    restartsInWindow,
    components: components.map((c) => {
      const component = {
        name: `${c.group}/${c.name}`,
        up: c.up === true,
        crashloop: c.crashloop === true,
        restartsInWindow: Number.isFinite(c.restartsInWindow) ? c.restartsInWindow : 0,
      };
      if (component.up !== true && typeof c.downReason === "string" && c.downReason.length > 0) {
        component.downReason = c.downReason;
      }
      return component;
    }),
  };
}

export function desktopLogIdentity(info) {
  return `${info.dev}:${info.ino}`;
}

export function isBoundedDesktopLogSpec(spec) {
  const argv = Array.isArray(spec?.argv) ? spec.argv : [];
  const runner = typeof argv[0] === "string" ? argv[0] : "";
  return (
    (runner === "/usr/local/bin/box-bounded-log" || runner.endsWith("/box-bounded-log")) &&
    argv[1] === "--run" &&
    argv[2] === spec.logFile &&
    argv[3] === "--"
  );
}

export function createDesktopSupervisor({
  attempt,
  safeRead,
  atomicWrite,
  closeLog,
  processIdentity,
}) {
  const { inspectTcpPortListener, isPidComm, isDockPid, readProcStartTime } = processIdentity;

  function parseDesktopComponentDescriptor(raw) {
    const parsed = attempt(() => JSON.parse(raw));
    if (!parsed.ok) return null;
    const value = parsed.value;
    if (typeof value !== "object" || value === null || Array.isArray(value)) {
      return null;
    }
    if (!Array.isArray(value.argv) || value.argv.length === 0) return null;
    if (!value.argv.every((a) => typeof a === "string" && a.length > 0)) {
      return null;
    }
    const order =
      typeof value.order === "number" && Number.isInteger(value.order) ? value.order : 0;
    const logFile =
      typeof value.logFile === "string" && value.logFile.length > 0 ? value.logFile : null;
    const pidFile =
      typeof value.pidFile === "string" && value.pidFile.length > 0 ? value.pidFile : null;
    const listenPort =
      Number.isInteger(value.listenPort) && value.listenPort > 0 && value.listenPort <= 65_535
        ? value.listenPort
        : null;
    const env = {};
    if (value.env != null && typeof value.env === "object" && !Array.isArray(value.env)) {
      for (const [k, v] of Object.entries(value.env)) {
        if (typeof v === "string" && v.length > 0) env[k] = v;
      }
    }
    return { argv: value.argv, env, order, logFile, pidFile, listenPort };
  }

  function desktopLogCursor(logFile) {
    if (logFile == null) return { size: 0, identity: null };
    const cursor = attempt(() => {
      const info = statSync(logFile);
      return { size: info.size, identity: desktopLogIdentity(info) };
    });
    return cursor.ok ? cursor.value : { size: 0, identity: null };
  }

  function boundedDesktopLogOwner(logFile) {
    if (logFile == null) return null;
    const raw = safeRead(`${logFile}.lock`);
    if (raw == null || !/^[1-9][0-9]*$/.test(raw.trim())) return null;
    const pid = Number(raw.trim());
    return Number.isSafeInteger(pid) ? pid : null;
  }

  function readDesktopLogTail(logFile, minOffset = 0, initialIdentity = null) {
    let fd = null;
    try {
      const tail = attempt(() => {
        fd = openSync(logFile, "r");
        const info = fstatSync(fd);
        const size = info.size;
        const generationChanged =
          initialIdentity != null && desktopLogIdentity(info) !== initialIdentity;
        const floor =
          !generationChanged && Number.isFinite(minOffset) && minOffset > 0 && minOffset <= size
            ? minOffset
            : 0;
        const start = Math.max(floor, size - DESKTOP_LOG_TAIL_BYTES);
        const len = size - start;
        if (len <= 0) return "";
        const buf = Buffer.allocUnsafe(len);
        const read = readSync(fd, buf, 0, len, start);
        const lines = buf
          .toString("utf8", 0, read)
          .split("\n")
          .flatMap((l) => {
            const trimmed = l.trimEnd();
            return trimmed.trim().length > 0 ? [trimmed] : [];
          });
        return lines.slice(-DESKTOP_LOG_TAIL_LINES).join("\n");
      });
      return tail.ok ? tail.value : "";
    } finally {
      if (fd != null) closeLog(fd);
    }
  }

  class DesktopSupervisor {
    initializeDesktopState(opts) {
      this.desktopState = new Map();
      this.desktopExitInfo = new Map();
      this.desktopLogOffset = new Map();
      this.desktopLogIdentity = new Map();
      this.desktopRevision = 0;
      this.desktopStateSig = null;
      this.desktopDir = opts.desktopDir ?? DESKTOP_DIR;
      this.desktopHealthPath = opts.desktopHealthPath ?? DESKTOP_HEALTH_PATH;
      this.desktopMaxRestarts = opts.desktopMaxRestarts ?? DESKTOP_RESTART_MAX_IN_WINDOW;
      this.desktopWindowMs = opts.desktopWindowMs ?? DESKTOP_RESTART_WINDOW_MS;
      this.desktopBackoffBaseMs = opts.desktopBackoffBaseMs ?? DESKTOP_BACKOFF_BASE_MS;
      this.desktopBackoffMaxMs = opts.desktopBackoffMaxMs ?? DESKTOP_BACKOFF_MAX_MS;
      this.desktopCompositorMaxCrashloops =
        opts.desktopCompositorMaxCrashloops ?? DESKTOP_COMPOSITOR_MAX_CRASHLOOPS;
      this.procRoot = opts.procRoot ?? "/proc";
    }

    initializeCgroupAccounting(opts) {
      this.cgroupRoot = opts.cgroupRoot ?? CGROUP_ROOT;
      this.cgroupSampleMs = opts.cgroupSampleMs ?? CGROUP_SAMPLE_MS;
      this.cgroupPrev = new Map();
      this.lastCgroupSampleAtMs = 0;
    }

    manageDesktop() {
      const nowMs = Date.now();
      if (!isDesktopSupervisionEnabled()) {
        this.writeDesktopHealth([], nowMs, false);
        return;
      }
      const specs = this.readDesktopSpecs();
      const presentIds = new Set(specs.map((s) => s.id));
      for (const id of [...this.desktopState.keys()]) {
        if (!presentIds.has(id)) this.desktopState.delete(id);
      }
      for (const id of [...this.desktopExitInfo.keys()]) {
        if (!presentIds.has(id)) this.desktopExitInfo.delete(id);
      }
      for (const id of [...this.desktopLogOffset.keys()]) {
        if (!presentIds.has(id)) this.desktopLogOffset.delete(id);
      }
      for (const id of [...this.desktopLogIdentity.keys()]) {
        if (!presentIds.has(id)) this.desktopLogIdentity.delete(id);
      }
      const groupsWithRoot = new Set(specs.filter((s) => s.order === 0).map((s) => s.group));
      const groupRootBlocked = new Map();
      const restartedThisTick = new Set();
      const health = [];
      const liveDesktopPids = [];
      const ordered = specs.slice().sort((a, b) => a.order - b.order || a.id.localeCompare(b.id));
      for (const spec of ordered) {
        const state = this.getDesktopState(spec.id);
        const restartsInWindow = countRestartsInWindow(
          state.restartTimestamps,
          nowMs,
          this.desktopWindowMs,
        );
        if (state.restartTimestamps.length > this.desktopMaxRestarts * 4) {
          state.restartTimestamps = state.restartTimestamps.filter(
            (t) => t >= nowMs - this.desktopWindowMs,
          );
        }
        let componentRestarts = restartsInWindow;
        const filePid = this.readDesktopPid(spec);
        let filePidAlive = filePid != null && isPidAlive(filePid);
        let listenerInspection = null;
        let listenerInspectionPid = null;
        let listenerIdentityMismatch = false;
        if (filePidAlive && spec.listenPort != null) {
          listenerInspection = inspectTcpPortListener(
            filePid,
            spec.listenPort,
            this.procRoot,
            spec.argv,
          );
          listenerInspectionPid = filePid;
          if (listenerInspection.reason === "pid-argv-mismatch") {
            filePidAlive = false;
            listenerIdentityMismatch = true;
            state.downReason = "port-not-listening";
          }
        }
        if (
          filePidAlive &&
          state.disabled &&
          spec.name === DESKTOP_COMPOSITOR_NAME &&
          existsSync("/proc/self/comm") &&
          !isPidComm(filePid, DESKTOP_COMPOSITOR_NAME)
        ) {
          filePidAlive = false;
        }
        if (
          filePidAlive &&
          spec.name === DESKTOP_DOCK_NAME &&
          existsSync("/proc/self/comm") &&
          !isDockPid(filePid)
        ) {
          filePidAlive = false;
        }
        let up;
        if (filePidAlive) {
          if (state.pid !== filePid) {
            const cursor = desktopLogCursor(spec.logFile);
            const ownsBoundedLog =
              isBoundedDesktopLogSpec(spec) && boundedDesktopLogOwner(spec.logFile) === filePid;
            this.desktopLogOffset.set(spec.id, ownsBoundedLog ? 0 : cursor.size);
            this.desktopLogIdentity.set(spec.id, cursor.identity);
            this.desktopExitInfo.delete(spec.id);
          }
          state.pid = filePid;
          up = true;
        } else if (state.pid != null && isPidAlive(state.pid)) {
          if (spec.listenPort != null && listenerInspectionPid !== state.pid) {
            listenerInspection = inspectTcpPortListener(
              state.pid,
              spec.listenPort,
              this.procRoot,
              spec.argv,
            );
            listenerInspectionPid = state.pid;
          }
          if (listenerInspection?.reason === "pid-argv-mismatch") {
            listenerIdentityMismatch = true;
            state.downReason = "port-not-listening";
            up = false;
          } else if (
            (state.disabled &&
              spec.name === DESKTOP_COMPOSITOR_NAME &&
              existsSync("/proc/self/comm") &&
              !isPidComm(state.pid, DESKTOP_COMPOSITOR_NAME)) ||
            (spec.name === DESKTOP_DOCK_NAME &&
              existsSync("/proc/self/comm") &&
              !isDockPid(state.pid))
          ) {
            up = false;
          } else {
            up = true;
            if (filePid !== state.pid) atomicWrite(spec.pidFile, String(state.pid));
          }
        } else {
          up = false;
        }
        const gatedByRoot =
          !up &&
          spec.order > 0 &&
          groupsWithRoot.has(spec.group) &&
          groupRootBlocked.get(spec.group) === true;
        if (!up && state.disabled) {
          state.crashloop = false;
          state.downReason = "compositor-disabled";
        } else if (gatedByRoot) {
          state.downReason = "awaiting-dependency";
        } else if (!up) {
          const action = decideDesktopComponentAction({
            supervisionEnabled: true,
            running: false,
            lastExitAtMs: state.lastExitAtMs,
            restartsInWindow,
            maxRestarts: this.desktopMaxRestarts,
            nowMs,
            baseBackoffMs: this.desktopBackoffBaseMs,
            maxBackoffMs: this.desktopBackoffMaxMs,
          });
          if (
            !listenerIdentityMismatch &&
            (action === "restart" ||
              (action === "crashloop" && !state.crashloop) ||
              state.downReason == null ||
              state.downReason === "awaiting-dependency")
          ) {
            state.downReason = this.desktopDownReason(spec, state);
          }
          if (action === "restart") {
            const newPid = this.restartDesktopComponent(spec);
            state.restartTimestamps.push(nowMs);
            state.lastExitAtMs = nowMs;
            restartedThisTick.add(spec.id);
            componentRestarts = restartsInWindow + 1;
            state.crashloop = false;
            if (newPid != null) {
              state.pid = newPid;
              up = true;
              log(
                `desktop: restarted ${spec.id} (pid ${newPid}, ${componentRestarts} restart(s) in window; down reason ${state.downReason})`,
              );
            } else {
              log(
                `desktop: relaunch of ${spec.id} FAILED (attempt ${componentRestarts} in window; down reason ${state.downReason}); backing off`,
              );
            }
          } else if (action === "crashloop") {
            if (!state.crashloop) {
              state.crashloopEpisodes = Number.isFinite(state.crashloopEpisodes)
                ? state.crashloopEpisodes + 1
                : 1;
              log(
                `desktop: ${spec.id} crash-looping (${restartsInWindow} restarts within ${this.desktopWindowMs}ms; episode ${state.crashloopEpisodes}; down reason ${state.downReason}); giving up until the storm subsides`,
              );
            }
            if (
              shouldDisableCompositor({
                name: spec.name,
                crashloopEpisodes: state.crashloopEpisodes,
                maxEpisodes: this.desktopCompositorMaxCrashloops,
              })
            ) {
              if (!state.disabled) {
                state.disabled = true;
                log(
                  `desktop: ${spec.id} hit ${state.crashloopEpisodes} crashloop episode(s); disabling the compositor for this supervisor's lifetime (the desktop runs WITHOUT a compositor from here)`,
                );
                this.reapDisabledCompositor(spec);
              }
              state.crashloop = false;
              state.downReason = "compositor-disabled";
            } else {
              state.crashloop = true;
            }
          }
        } else {
          state.downReason = null;
          this.desktopExitInfo.delete(spec.id);
          if (state.crashloop) {
            state.crashloop = false;
            log(`desktop: ${spec.id} recovered (up again)`);
          }
          if (state.disabled) {
            state.disabled = false;
            state.crashloopEpisodes = 0;
            state.restartTimestamps = [];
            state.lastExitAtMs = null;
            log(`desktop: ${spec.id} is up again externally; re-enabling supervision`);
          }
        }
        let available = up;
        if (up && spec.listenPort != null && state.pid != null) {
          const inspection =
            listenerInspectionPid === state.pid
              ? listenerInspection
              : inspectTcpPortListener(state.pid, spec.listenPort, this.procRoot, spec.argv);
          if (inspection.status === "missing") {
            available = false;
            state.downReason = "port-not-listening";
          } else if (inspection.status === "unavailable") {
            available = false;
            state.downReason = "listener-procfs-unavailable";
            this.noteDesktopListenerDiagnostic(spec, state, inspection.reason);
          }
        }
        if (spec.order === 0) {
          groupRootBlocked.set(spec.group, !up || restartedThisTick.has(spec.id));
        }
        if (up && state.pid != null) {
          liveDesktopPids.push({ id: spec.id, pid: state.pid });
        }
        health.push({
          group: spec.group,
          name: spec.name,
          up: available,
          crashloop: gatedByRoot ? false : state.crashloop,
          restartsInWindow: componentRestarts,
          downReason: available ? undefined : (state.downReason ?? undefined),
        });
      }
      this.reconcileDesktopCgroup(liveDesktopPids);
      this.writeDesktopHealth(health, nowMs, true);
    }

    reconcileDesktopCgroup(components) {
      if (components.length === 0) return;
      const membership = safeRead(join(this.cgroupRoot, CGROUP_INTERACTIVE, "cgroup.procs"));
      if (membership == null) return;
      const present = new Set(
        membership.split("\n").flatMap((line) => {
          const trimmed = line.trim();
          return trimmed.length > 0 ? [trimmed] : [];
        }),
      );
      for (const { id, pid } of components) {
        if (present.has(String(pid))) continue;
        if (this.placeInCgroup(CGROUP_INTERACTIVE, pid)) {
          log(`desktop: moved ${id} (pid ${pid}) into the ${CGROUP_INTERACTIVE} cgroup`);
        }
      }
    }

    desktopDownReason(spec, state) {
      const exit = this.desktopExitInfo.get(spec.id);
      const fresh = exit != null && exit.pid != null && exit.pid === state.pid;
      const exitCode = fresh && Number.isInteger(exit.code) ? exit.code : null;
      const signal = fresh && typeof exit.signal === "string" ? exit.signal : null;
      const logTail =
        spec.logFile != null
          ? readDesktopLogTail(
              spec.logFile,
              this.desktopLogOffset.get(spec.id) ?? 0,
              this.desktopLogIdentity.get(spec.id) ?? null,
            )
          : "";
      return classifyDesktopDownReason({ exitCode, signal, logTail });
    }

    noteDesktopListenerDiagnostic(spec, state, reason) {
      if (state.listenerDiagnosticPid === state.pid) return;
      state.listenerDiagnosticPid = state.pid;
      log(
        `desktop: ${spec.id} pid ${state.pid} listener health unavailable (${reason}); leaving the live process untouched`,
      );
    }

    readDesktopIdentity(spec) {
      const pid = this.readDesktopPid(spec);
      if (pid == null || !isPidAlive(pid)) return null;
      const startTime = readProcStartTime(pid, this.procRoot);
      return startTime == null ? null : { pid, startTime };
    }

    wouldRestartAfterExit(state) {
      const nowMs = Date.now();
      const restartsInWindow = countRestartsInWindow(
        state.restartTimestamps,
        nowMs,
        this.desktopWindowMs,
      );
      return (
        decideDesktopComponentAction({
          supervisionEnabled: true,
          running: false,
          lastExitAtMs: state.lastExitAtMs,
          restartsInWindow,
          maxRestarts: this.desktopMaxRestarts,
          nowMs,
          baseBackoffMs: this.desktopBackoffBaseMs,
          maxBackoffMs: this.desktopBackoffMaxMs,
        }) === "restart"
      );
    }

    getDesktopState(id) {
      let state = this.desktopState.get(id);
      if (state == null) {
        state = {
          pid: null,
          lastExitAtMs: null,
          restartTimestamps: [],
          crashloop: false,
          downReason: null,
          listenerDiagnosticPid: null,
          crashloopEpisodes: 0,
          disabled: false,
        };
        this.desktopState.set(id, state);
      }
      return state;
    }

    readDesktopSpecs() {
      const specs = [];
      const groups = attempt(() => readdirSync(this.desktopDir, { withFileTypes: true }));
      if (!groups.ok) return specs;
      for (const groupEntry of groups.value) {
        if (!groupEntry.isDirectory()) continue;
        const groupDir = join(this.desktopDir, groupEntry.name);
        const files = attempt(() => readdirSync(groupDir, { withFileTypes: true }));
        if (!files.ok) continue;
        for (const file of files.value) {
          if (!file.isFile() || !file.name.endsWith(".json")) continue;
          const descriptorPath = join(groupDir, file.name);
          const raw = safeRead(descriptorPath);
          if (raw == null) continue;
          const parsed = parseDesktopComponentDescriptor(raw);
          if (parsed == null) continue;
          const name = file.name.slice(0, -".json".length);
          specs.push({
            id: `${groupEntry.name}/${name}`,
            group: groupEntry.name,
            name,
            order: parsed.order,
            argv: parsed.argv,
            env: parsed.env,
            logFile: parsed.logFile,
            listenPort: parsed.listenPort,
            descriptorPath,
            pidFile: parsed.pidFile ?? join(groupDir, `${name}.pid`),
          });
        }
      }
      return specs;
    }

    readDesktopPid(spec) {
      const raw = safeRead(spec.pidFile);
      if (raw == null) return null;
      const pid = Number.parseInt(raw.trim(), 10);
      return Number.isInteger(pid) && pid > 0 ? pid : null;
    }

    placeInCgroup(group, pid) {
      if (!Number.isInteger(pid) || pid <= 0) return false;
      const placed = attempt(() =>
        writeFileSync(join(this.cgroupRoot, group, "cgroup.procs"), String(pid)),
      );
      return placed.ok;
    }

    readCgroupSample(group) {
      const stat = parseCgroupCpuStat(safeRead(join(this.cgroupRoot, group, "cpu.stat")));
      if (stat == null) return null;
      const pressure = parseCgroupPressure(safeRead(join(this.cgroupRoot, group, "cpu.pressure")));
      return {
        usageUsec: stat.usage_usec,
        throttledUsec: typeof stat.throttled_usec === "number" ? stat.throttled_usec : null,
        stallUsec: pressure?.someTotalUsec ?? null,
      };
    }

    sampleCgroups(nowMs) {
      if (nowMs - this.lastCgroupSampleAtMs < this.cgroupSampleMs) return;
      const elapsedMs = this.lastCgroupSampleAtMs === 0 ? 0 : nowMs - this.lastCgroupSampleAtMs;
      this.lastCgroupSampleAtMs = nowMs;
      for (const group of [CGROUP_INTERACTIVE, CGROUP_AGENT]) {
        const next = this.readCgroupSample(group);
        if (next == null) {
          this.cgroupPrev.delete(group);
          continue;
        }
        const line = formatCgroupCpuSample({
          group,
          prev: this.cgroupPrev.get(group) ?? null,
          next,
          elapsedMs,
        });
        if (line != null) log(line);
        this.cgroupPrev.set(group, next);
      }
    }

    writeDesktopHealth(components, nowMs, supervisionEnabled) {
      const sig = `${supervisionEnabled ? 1 : 0}#${desktopStateSignature(components)}`;
      if (sig !== this.desktopStateSig) {
        this.desktopRevision += 1;
        this.desktopStateSig = sig;
      }
      atomicWrite(
        this.desktopHealthPath,
        JSON.stringify(
          buildDesktopHealthSnapshot({
            supervisionEnabled,
            components,
            nowMs,
            revision: this.desktopRevision,
          }),
        ),
      );
    }

    restartDesktopComponent() {
      throw new Error("DesktopSupervisor requires restartDesktopComponent");
    }

    reapDisabledCompositor() {
      throw new Error("DesktopSupervisor requires reapDisabledCompositor");
    }
  }

  return { DesktopSupervisor, parseDesktopComponentDescriptor, desktopLogCursor };
}
