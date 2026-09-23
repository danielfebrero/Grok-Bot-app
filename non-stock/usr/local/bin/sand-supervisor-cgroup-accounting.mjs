import { numberFromEnv } from "./sand-supervisor-contract.mjs";

export const CGROUP_ROOT = process.env.SAND_CGROUP_ROOT ?? "/sys/fs/cgroup";
export const CGROUP_INTERACTIVE = "interactive";
export const CGROUP_AGENT = "agent";
export const CGROUP_SAMPLE_MS = numberFromEnv("SAND_CGROUP_SAMPLE_MS", 60_000, 1_000);

export function parseCgroupCpuStat(raw) {
  if (typeof raw !== "string") return null;
  const out = {};
  for (const line of raw.split("\n")) {
    const [key, value] = line.trim().split(/\s+/);
    if (key == null || value == null) continue;
    const parsed = Number.parseInt(value, 10);
    if (!Number.isFinite(parsed)) continue;
    out[key] = parsed;
  }
  return typeof out.usage_usec === "number" ? out : null;
}

export function parseCgroupPressure(raw) {
  if (typeof raw !== "string") return null;
  const out = { someTotalUsec: null, fullTotalUsec: null };
  for (const line of raw.split("\n")) {
    const trimmed = line.trim();
    const match = /^(some|full)\b.*\btotal=(\d+)/.exec(trimmed);
    if (match == null) continue;
    const total = Number.parseInt(match[2], 10);
    if (!Number.isFinite(total)) continue;
    if (match[1] === "some") out.someTotalUsec = total;
    else out.fullTotalUsec = total;
  }
  return out.someTotalUsec == null && out.fullTotalUsec == null ? null : out;
}

export function formatCgroupCpuSample({ group, prev, next, elapsedMs }) {
  if (prev == null || next == null) return null;
  if (!Number.isFinite(elapsedMs) || elapsedMs <= 0) return null;
  const usageDelta = next.usageUsec - prev.usageUsec;
  if (!Number.isFinite(usageDelta) || usageDelta < 0) return null;
  const elapsedUsec = elapsedMs * 1000;
  const cpuPct = (usageDelta / elapsedUsec) * 100;
  const parts = [
    `cgroup ${group}: cpu=${cpuPct.toFixed(1)}% of 1 core`,
    `over ${Math.round(elapsedMs / 1000)}s`,
  ];
  if (prev.stallUsec != null && next.stallUsec != null) {
    const stallDelta = next.stallUsec - prev.stallUsec;
    if (Number.isFinite(stallDelta) && stallDelta >= 0) {
      parts.push(`stalled=${((stallDelta / elapsedUsec) * 100).toFixed(1)}%`);
    }
  }
  if (
    prev.throttledUsec != null &&
    next.throttledUsec != null &&
    next.throttledUsec - prev.throttledUsec > 0
  ) {
    const throttleDelta = next.throttledUsec - prev.throttledUsec;
    parts.push(`throttled=${((throttleDelta / elapsedUsec) * 100).toFixed(1)}%`);
  }
  return parts.join(" ");
}
