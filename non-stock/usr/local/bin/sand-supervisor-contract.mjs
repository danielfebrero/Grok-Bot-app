import { readdirSync } from "node:fs";

const SUPERVISOR_ENTRY_NAME = "sand-supervisor.mjs";

export const SUPERVISOR_DIR = "/tmp/sand-supervisor";
export const COMMAND_PATH = `${SUPERVISOR_DIR}/command.json`;
export const STATUS_PATH = `${SUPERVISOR_DIR}/status.json`;
export const ACKS_DIR = `${SUPERVISOR_DIR}/acks`;
export const STAGED_BUNDLE_PATH = `${SUPERVISOR_DIR}/incoming-host-bundle.tgz`;
export const AGENT_DATA_ROOT = "/home/box/sand-data";
export const UPGRADE_APPLIED_PATH = `${AGENT_DATA_ROOT}/.sand-host-upgrade.json`;
export const HOST_CRASH_MARKER_PATH = `${AGENT_DATA_ROOT}/.sand-host-crash.json`;
export const HOST_DIR = "/home/box/sand-host";
export const HOST_ENTRY = `${HOST_DIR}/host-main.cjs`;
export const HOST_NODE_WARNING_FLAG = "--disable-warning=ExperimentalWarning";
export const HOST_LOG_PATH = "/tmp/sand-host.log";
export const HOST_VERSION_PATH = `${HOST_DIR}/version`;

export const IMAGE_SHA_PATH = "/etc/sand-box-image-sha";
export const HOST_BUNDLE_BASE_URL_ENV = "SAND_HOST_BUNDLE_S3_BASE_URL";
export const HOST_BUNDLE_CHANNELS = ["latest", "stable"];
export const HOST_BUNDLE_DEFAULT_CHANNEL = "latest";
export const HOST_BUNDLE_CHANNEL_ENV = "SAND_HOST_BUNDLE_CHANNEL";
export const BOOT_FETCH_GIT_SHA_REGEX = /^[0-9a-f]{7,40}$/;
export const HOST_BUNDLE_SHA256_REGEX = /^[0-9a-f]{64}$/;
export const BOOT_FETCH_BUDGET_MS_ENV = "SAND_SUPERVISOR_BOOT_FETCH_BUDGET_MS";
export const BOOT_FETCH_BUDGET_DEFAULT_MS = 20_000;

export const DESKTOP_DIR = "/tmp/sand-desktop";
export const DESKTOP_HEALTH_PATH = `${SUPERVISOR_DIR}/desktop-health.json`;
export const DESKTOP_SUPERVISION_DISABLED_ENV = "SAND_DESKTOP_SUPERVISION_DISABLED";

export const BOX_SCRIPTS_SOURCE_DIR = `${HOST_DIR}/box-scripts`;
export const BOX_SCRIPTS_BIN_DIR = "/usr/local/bin";
export const BOX_SCRIPTS_MARKER_PATH = "/usr/local/share/sand-box-scripts/version";
export const BOX_SCRIPTS_SYNC_DISABLED_ENV = "SAND_BOX_SCRIPT_SYNC_DISABLED";
export const BOX_SCRIPTS_RETRY_BACKOFF_MS = 5 * 60_000;
const IMAGE_OWNED_SCRIPT_NAMES = [
  "start-sand-box",
  "sand-exit-watch",
  SUPERVISOR_ENTRY_NAME,
  "fetch-exec-daemon",
  "sand-desktop-supervise.sh",
  "box-cgroups.sh",
  "ensure-machine-id",
  "box-xvfb",
  "box-x11vnc",
  "start-exec-daemon",
  "supervise-exec-daemon",
  "supervise-sand-supervisor",
];

export function isSupervisorModuleName(name) {
  return (
    typeof name === "string" &&
    name.startsWith("sand-supervisor-") &&
    name.endsWith(".mjs") &&
    !name.includes("/")
  );
}

export function supervisorModuleNames(directory) {
  return [SUPERVISOR_ENTRY_NAME, ...readdirSync(directory).filter(isSupervisorModuleName).sort()];
}

export function BOX_SCRIPTS_DENY(name) {
  return IMAGE_OWNED_SCRIPT_NAMES.includes(name) || isSupervisorModuleName(name);
}

export const COMMAND_KINDS = ["ping", "restart", "upgrade"];
export const UPGRADE_MODES = ["bundle", "image", "restart"];

export const TICK_MS = numberFromEnv("SAND_SUPERVISOR_TICK_MS", 5_000, 250);
export const RESTART_BACKOFF_BASE_MS = 1_000;
export const RESTART_BACKOFF_MAX_MS = 60_000;
export const HEALTH_TIMEOUT_MS = 1_500;

export const DESKTOP_BACKOFF_BASE_MS = 1_000;
export const DESKTOP_BACKOFF_MAX_MS = 30_000;
export const DESKTOP_RESTART_WINDOW_MS = numberFromEnv(
  "SAND_DESKTOP_RESTART_WINDOW_MS",
  10 * 60_000,
  1_000,
);
export const DESKTOP_RESTART_MAX_IN_WINDOW = numberFromEnv("SAND_DESKTOP_MAX_RESTARTS", 8, 1);

export const FORK_RFB_PROBE_TIMEOUT_MS = 1_000;
export const FORK_RFB_FAILURE_THRESHOLD = 2;
export const FORK_RFB_TERMINATE_GRACE_MS = 5_000;

export const DESKTOP_COMPOSITOR_MAX_CRASHLOOPS = numberFromEnv(
  "SAND_DESKTOP_COMPOSITOR_MAX_CRASHLOOPS",
  3,
  1,
);
export const DESKTOP_COMPOSITOR_NAME = "picom";
export const DESKTOP_DOCK_NAME = "dock";
export const DESKTOP_DOCK_COMMS = Object.freeze(["plank", "box-plank", "box-bounded-log"]);
export const DESKTOP_LOG_TAIL_BYTES = 8_192;
export const DESKTOP_LOG_TAIL_LINES = 6;

export const HOST_UPGRADE_MAX_DEFER_MS = 6 * 60 * 60 * 1000;
export const MAX_DEFER_MS = numberFromEnv(
  "SAND_SUPERVISOR_MAX_DEFER_MS",
  HOST_UPGRADE_MAX_DEFER_MS,
  1_000,
);
export const POST_SWAP_HEALTHY_UPTIME_MS = numberFromEnv(
  "SAND_SUPERVISOR_POST_SWAP_HEALTHY_MS",
  60_000,
  1_000,
);
export const POST_SWAP_MAX_QUICK_EXITS = numberFromEnv(
  "SAND_SUPERVISOR_POST_SWAP_MAX_QUICK_EXITS",
  3,
  1,
);
export const POST_SWAP_MAX_RESTORE_ATTEMPTS = 5;

export function numberFromEnv(name, fallback, min) {
  const raw = process.env[name];
  if (raw == null || raw.length === 0) return fallback;
  const parsed = Number.parseInt(raw, 10);
  if (!Number.isInteger(parsed) || parsed <= 0) return fallback;
  return Math.max(min, parsed);
}

export function isDesktopSupervisionEnabled(env = process.env) {
  const raw = env[DESKTOP_SUPERVISION_DISABLED_ENV];
  const v = typeof raw === "string" ? raw.trim().toLowerCase() : "";
  return !(v === "1" || v === "true" || v === "yes");
}

export function isHostSupervisionEnabled(env = process.env) {
  return env.SAND_SUPERVISOR_ENABLED === "1";
}

export function log(message) {
  process.stdout.write(`[sand-supervisor] ${message}\n`);
}
