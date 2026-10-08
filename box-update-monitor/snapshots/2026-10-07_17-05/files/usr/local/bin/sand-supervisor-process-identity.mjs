import { readFileSync, readdirSync, readlinkSync } from "node:fs";
import { join } from "node:path";
import { DESKTOP_COMPOSITOR_NAME, DESKTOP_DOCK_COMMS } from "./sand-supervisor-contract.mjs";

// kill(2) with signal 0 sends nothing but still runs the existence and permission checks, so EPERM means the pid exists and ESRCH means it is gone: https://man7.org/linux/man-pages/man2/kill.2.html
export function isPidAlive(pid) {
  if (!Number.isInteger(pid) || pid <= 0) return false;
  try {
    process.kill(pid, 0);
    return true;
  } catch (error) {
    return error != null && error.code === "EPERM";
  }
}

export function listenerArgvMatches(actualArgv, expectedArgv) {
  if (!Array.isArray(expectedArgv) || expectedArgv.length === 0) return false;
  if (actualArgv.length < expectedArgv.length) return false;
  if (actualArgv[0] !== expectedArgv[0]) return false;
  const program = expectedArgv.slice(1);
  const tail = actualArgv.slice(actualArgv.length - program.length);
  if (!program.every((value, index) => value === tail[index])) return false;
  const injected = actualArgv.slice(1, actualArgv.length - program.length);
  return injected.every(
    (value, index) =>
      value === expectedArgv[0] ||
      value.startsWith("-") ||
      (index > 0 && injected[index - 1] === "-r"),
  );
}

export function createProcessIdentity({ attempt, safeRead }) {
  function listenerInspectionUnavailable(reason) {
    return { status: "unavailable", reason, startTime: null };
  }

  function stableListenerInspection(pid, procRoot, startTime, status, reason) {
    const after = readProcStartTime(pid, procRoot);
    if (after == null) return listenerInspectionUnavailable("pid-stat-unreadable");
    if (after !== startTime) {
      return listenerInspectionUnavailable("pid-identity-changed");
    }
    return { status, reason, startTime };
  }

  function inspectTcpPortListener(pid, port, procRoot = "/proc", expectedArgv = null) {
    if (
      !Number.isInteger(pid) ||
      pid <= 0 ||
      !Number.isInteger(port) ||
      port <= 0 ||
      port > 65_535
    ) {
      return listenerInspectionUnavailable("invalid-input");
    }
    const startTime = readProcStartTime(pid, procRoot);
    if (startTime == null) return listenerInspectionUnavailable("pid-stat-unreadable");
    if (expectedArgv != null) {
      const rawCmdline = safeRead(join(procRoot, String(pid), "cmdline"));
      if (rawCmdline == null) {
        return listenerInspectionUnavailable("pid-cmdline-unreadable");
      }
      const actualArgv = rawCmdline.split("\0");
      if (actualArgv.at(-1) === "") actualArgv.pop();
      if (!listenerArgvMatches(actualArgv, expectedArgv)) {
        return stableListenerInspection(pid, procRoot, startTime, "missing", "pid-argv-mismatch");
      }
    }

    const rawTcp = safeRead(join(procRoot, "net/tcp"));
    if (rawTcp == null) {
      return listenerInspectionUnavailable("ipv4-listener-table-unreadable");
    }
    const lines = rawTcp.split("\n");
    const header = lines.shift()?.trim().split(/\s+/) ?? [];
    if (
      header[0] !== "sl" ||
      header[1] !== "local_address" ||
      header[2] !== "rem_address" ||
      header[3] !== "st"
    ) {
      return listenerInspectionUnavailable("ipv4-listener-table-malformed");
    }

    const portHex = port.toString(16).toUpperCase().padStart(4, "0");
    const listenerInodes = new Set();
    for (const line of lines) {
      if (line.trim().length === 0) continue;
      const fields = line.trim().split(/\s+/);
      const localAddress = /^([0-9A-Fa-f]{8}):([0-9A-Fa-f]{4})$/.exec(fields[1] ?? "");
      if (
        fields.length < 10 ||
        !/^\d+:$/.test(fields[0] ?? "") ||
        localAddress == null ||
        !/^[0-9A-Fa-f]{2}$/.test(fields[3] ?? "") ||
        !/^\d+$/.test(fields[9] ?? "")
      ) {
        return listenerInspectionUnavailable("ipv4-listener-table-malformed");
      }
      if (fields[3].toUpperCase() !== "0A") continue;
      const [, address, observedPort] = localAddress;
      if (
        observedPort.toUpperCase() === portHex &&
        (address.toUpperCase() === "00000000" || address.toUpperCase() === "0100007F")
      ) {
        listenerInodes.add(fields[9]);
      }
    }
    if (listenerInodes.size === 0) {
      return stableListenerInspection(
        pid,
        procRoot,
        startTime,
        "missing",
        "usable-ipv4-listener-absent",
      );
    }

    const descriptors = attempt(() => readdirSync(join(procRoot, String(pid), "fd")));
    if (!descriptors.ok) {
      return listenerInspectionUnavailable("pid-fd-table-unreadable");
    }
    let descriptorRaced = false;
    for (const descriptor of descriptors.value) {
      const target = attempt(() => readlinkSync(join(procRoot, String(pid), "fd", descriptor)));
      if (!target.ok) {
        descriptorRaced = true;
        continue;
      }
      const socket = /^socket:\[(\d+)\]$/.exec(target.value);
      if (socket != null && listenerInodes.has(socket[1])) {
        return stableListenerInspection(
          pid,
          procRoot,
          startTime,
          "listening",
          "owned-usable-ipv4-listener",
        );
      }
    }
    if (descriptorRaced) {
      return listenerInspectionUnavailable("pid-fd-table-raced");
    }
    return stableListenerInspection(
      pid,
      procRoot,
      startTime,
      "missing",
      "usable-ipv4-listener-not-owned",
    );
  }

  function isPidComm(pid, name, procRoot = "/proc") {
    if (!Number.isInteger(pid) || pid <= 0 || typeof name !== "string" || name.length === 0) {
      return false;
    }
    const comm = attempt(() => readFileSync(join(procRoot, String(pid), "comm"), "utf8").trim());
    return comm.ok ? comm.value === name : false;
  }

  function isExpectedForkRfbProcess({ pid, startTime, name, port, procRoot }) {
    const comm = name === "x11vnc" ? "x11vnc" : "websockify";
    if (!isPidComm(pid, comm, procRoot)) return false;
    const rawCmdline = safeRead(join(procRoot, String(pid), "cmdline"));
    if (rawCmdline == null) return false;
    const argv = rawCmdline.split("\0");
    const matches =
      name === "x11vnc"
        ? argv.some((value, index) => value === "-rfbport" && argv[index + 1] === String(port))
        : argv.includes(`0.0.0.0:${port}`);
    return matches && readProcStartTime(pid, procRoot) === startTime;
  }

  function isDockPid(pid) {
    return DESKTOP_DOCK_COMMS.some((comm) => isPidComm(pid, comm));
  }

  function findCompositorReapTargets({
    display,
    procRoot = "/proc",
    ownPid = process.pid,
    comm = DESKTOP_COMPOSITOR_NAME,
  }) {
    if (typeof display !== "string" || display.length === 0) return [];
    const entries = attempt(() => readdirSync(procRoot));
    if (!entries.ok) return [];
    const targets = [];
    for (const entry of entries.value) {
      if (!/^\d+$/.test(entry)) continue;
      const pid = Number.parseInt(entry, 10);
      if (!Number.isInteger(pid) || pid <= 0 || pid === ownPid) continue;
      const procComm = attempt(() => readFileSync(join(procRoot, entry, "comm"), "utf8").trim());
      if (!procComm.ok) continue;
      if (procComm.value !== comm) continue;
      const environ = attempt(() => readFileSync(join(procRoot, entry, "environ"), "utf8"));
      if (!environ.ok) continue;
      const envDisplay = environ.value.split("\0").find((kv) => kv.startsWith("DISPLAY="));
      if (envDisplay === `DISPLAY=${display}`) targets.push(pid);
    }
    return targets;
  }

  // /proc/<pid>/stat writes comm (field 2) in parentheses unescaped, so the parse anchors on the last ')', and starttime is field 22, index 19 after the comm: https://man7.org/linux/man-pages/man5/proc_pid_stat.5.html and https://github.com/torvalds/linux/blob/v6.15/fs/proc/array.c
  function readProcStartTime(pid, procRoot = "/proc") {
    if (!Number.isInteger(pid) || pid <= 0) return null;
    const read = attempt(() => readFileSync(join(procRoot, String(pid), "stat"), "utf8"));
    if (!read.ok) return null;
    const raw = read.value;
    const close = raw.lastIndexOf(")");
    if (close === -1) return null;
    const fields = raw
      .slice(close + 1)
      .trim()
      .split(/\s+/);
    const startTime = fields[19];
    return typeof startTime === "string" && /^\d+$/.test(startTime) ? startTime : null;
  }

  return {
    inspectTcpPortListener,
    isPidComm,
    isExpectedForkRfbProcess,
    isDockPid,
    findCompositorReapTargets,
    readProcStartTime,
  };
}
