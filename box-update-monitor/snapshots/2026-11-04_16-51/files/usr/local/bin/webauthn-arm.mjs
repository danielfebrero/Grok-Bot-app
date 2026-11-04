import {
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  renameSync,
  statSync,
  writeFileSync,
} from "node:fs";

export const SAND_WEBAUTHN_ARM_DIR = "/tmp/.sand-browser";
export const SAND_WEBAUTHN_ARM_UNTIL_PATH = `${SAND_WEBAUTHN_ARM_DIR}/webauthn-armed-until`;
export const SAND_WEBAUTHN_HANDOFF_ARM_UNTIL_PATH = "/tmp/sand-webauthn-armed-until";
export const SAND_WEBAUTHN_BROWSER_ARM_MS = 20 * 60 * 1000;
export const SAND_WEBAUTHN_HANDOFF_ARM_MS = 45 * 60 * 1000;
export const SAND_WEBAUTHN_STALE_TAB_MESSAGE =
  "Grok Bot ignored a security key request from a browser tab that isn't being used right now.";

function readArmUntil(armUntilPath) {
  if (!existsSync(armUntilPath)) {
    return undefined;
  }
  const text = readFileSync(armUntilPath, "utf8").trim();
  return /^\d{1,15}$/.test(text) ? Number(text) : undefined;
}

function armUntilIsLive(armUntilPath, nowMs) {
  const until = readArmUntil(armUntilPath);
  return until !== undefined && until > nowMs;
}

export function isWebAuthnArmed(nowMs = Date.now(), paths = {}) {
  const armDir = paths.armDir ?? SAND_WEBAUTHN_ARM_DIR;
  const armUntilPaths = [
    paths.armUntilPath ?? SAND_WEBAUTHN_ARM_UNTIL_PATH,
    paths.handoffArmUntilPath ?? SAND_WEBAUTHN_HANDOFF_ARM_UNTIL_PATH,
  ];
  for (const armUntilPath of armUntilPaths) {
    if (armUntilIsLive(armUntilPath, nowMs)) {
      return true;
    }
  }
  if (!existsSync(armDir)) {
    return false;
  }
  for (const name of readdirSync(armDir)) {
    if (!name.startsWith("views-") || !name.endsWith(".json")) {
      continue;
    }
    const stat = statSync(`${armDir}/${name}`);
    if (stat.isFile() && nowMs - stat.mtimeMs <= SAND_WEBAUTHN_BROWSER_ARM_MS) {
      return true;
    }
  }
  return false;
}

export function armWebAuthn(ttlMs, nowMs = Date.now(), paths = {}) {
  const armDir = paths.armDir ?? SAND_WEBAUTHN_ARM_DIR;
  const armUntilPath = paths.armUntilPath ?? SAND_WEBAUTHN_ARM_UNTIL_PATH;
  mkdirSync(armDir, { recursive: true });
  const existing = readArmUntil(armUntilPath);
  const until = Math.max(nowMs + ttlMs, existing ?? 0);
  const tmpPath = `${armUntilPath}.${process.pid}.tmp`;
  writeFileSync(tmpPath, `${until}\n`);
  renameSync(tmpPath, armUntilPath);
  return until;
}
