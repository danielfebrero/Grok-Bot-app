import { execFile } from "node:child_process";
import {
  chmodSync,
  copyFileSync,
  cpSync,
  existsSync,
  lstatSync,
  mkdirSync,
  readdirSync,
  renameSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { join } from "node:path";
import { promisify } from "node:util";
import {
  HOST_DIR,
  BOOT_FETCH_GIT_SHA_REGEX,
  BOX_SCRIPTS_DENY,
  numberFromEnv,
  log,
} from "./sand-supervisor-contract.mjs";

// With redirect_dir off (the overlayfs default) rename(2) of a directory that lives on a lower layer returns EXDEV, so the swap never renames the image-baked hostDir and publishes entries into it, with scratch dirs as hostDir siblings: https://docs.kernel.org/6.15/filesystems/overlayfs.html
export function hostBundleSwapTargets(hostDir = HOST_DIR) {
  return {
    hostDir,
    stageDir: `${hostDir}.stage`,
    backupDir: `${hostDir}.prev`,
  };
}

export function bundleSwapTargetsClearOfDataRoot({ hostDir, stageDir, backupDir }, dataRoot) {
  return [hostDir, stageDir, backupDir]
    .filter((target) => target != null)
    .every((target) => !pathsOverlap(target, dataRoot));
}

export function pathsOverlap(a, b) {
  const sa = pathSegments(a);
  const sb = pathSegments(b);
  if (sa.length === 0 || sb.length === 0) return true;
  const shorter = sa.length <= sb.length ? sa : sb;
  const longer = sa.length <= sb.length ? sb : sa;
  return shorter.every((segment, i) => segment === longer[i]);
}

function pathSegments(path) {
  return String(path ?? "")
    .split("/")
    .filter((segment) => segment.length > 0 && segment !== ".");
}

export function classifySwapError(error) {
  const code = error != null && typeof error === "object" ? error.code : undefined;
  switch (code) {
    case "EXDEV":
      return "rename-exdev";
    case "EACCES":
    case "EPERM":
      return "perms";
    case "ENOSPC":
      return "nospace";
    default:
      return "swap-failed";
  }
}

export function classifyExtractError(error) {
  if (error != null && typeof error === "object" && error.code === "ENOSPC") {
    return "extract-nospace";
  }
  const stderr = extractErrorStderr(error);
  if (/no space left on device/i.test(stderr)) return "extract-nospace";
  if (
    /not in gzip format|unexpected end of file|invalid compressed data|unexpected eof|damaged|corrupt/i.test(
      stderr,
    )
  ) {
    return "extract-corrupt";
  }
  return "extract-failed";
}

function extractErrorStderr(error) {
  const raw = error != null && typeof error === "object" ? error.stderr : undefined;
  let text = "";
  if (typeof raw === "string") {
    text = raw;
  } else if (raw instanceof Uint8Array) {
    text = Buffer.from(raw).toString("utf8");
  }
  return text.slice(0, 2_000);
}

const EXTRACT_TIMEOUT_MS = numberFromEnv("SAND_SUPERVISOR_STALE_S", 180, 1) * 1_000;
const execFileAsync = promisify(execFile);

function extractBundleInto(bundle, stageDir, timeoutMs, signal) {
  const extract = execFileAsync("tar", ["-xzf", "-", "-C", stageDir], {
    encoding: "buffer",
    timeout: timeoutMs,
    signal,
  });
  extract.child.stdin.on("error", (error) =>
    log(`bundle extract stdin closed early: ${error.code ?? String(error)}`),
  );
  extract.child.stdin.end(bundle);
  return extract;
}

function publishBundleEntry(srcPath, destPath) {
  try {
    renameSync(srcPath, destPath);
    return;
  } catch (error) {
    const code = error?.code;
    if (code === "ENOTEMPTY" || code === "EEXIST" || code === "EISDIR") {
      rmSync(destPath, { recursive: true, force: true });
      try {
        renameSync(srcPath, destPath);
        return;
      } catch (retryError) {
        if (retryError?.code !== "EXDEV") throw retryError;
      }
    } else if (code !== "EXDEV") {
      throw error;
    }
  }
  const tmpPath = `${destPath}.incoming-${process.pid}-${Date.now()}`;
  rmSync(tmpPath, { recursive: true, force: true });
  cpSync(srcPath, tmpPath, { recursive: true });
  try {
    renameSync(tmpPath, destPath);
  } catch (error) {
    const code = error?.code;
    if (code !== "ENOTEMPTY" && code !== "EEXIST" && code !== "EISDIR") {
      rmSync(tmpPath, { recursive: true, force: true });
      throw error;
    }
    rmSync(destPath, { recursive: true, force: true });
    renameSync(tmpPath, destPath);
  }
}

function describeStderrExcerpt(error) {
  const line = extractErrorStderr(error).replace(/\s+/g, " ").trim().slice(0, 200);
  return line.length > 0 ? ` [stderr: ${line}]` : "";
}

export function createHostBundleStore({
  attempt,
  safeUnlink,
  safeRmRecursive,
  readVersionMarker,
  ensureDir,
}) {
  function keepExistingHostBackup(backupDir, hostDir) {
    return (
      existsSync(join(backupDir, "host-main.cjs")) &&
      readVersionMarker(join(backupDir, "version")) != null &&
      readVersionMarker(join(backupDir, "version")) === readVersionMarker(join(hostDir, "version"))
    );
  }

  async function swapHostBundle({
    bundlePath,
    bundle,
    hostDir,
    stageDir,
    backupDir,
    expectedVersion,
    extractTimeoutMs = EXTRACT_TIMEOUT_MS,
    signal,
  }) {
    let entryCommitted = false;
    try {
      rmSync(stageDir, { recursive: true, force: true });
      mkdirSync(stageDir, { recursive: true });
      try {
        await extractBundleInto(bundle, stageDir, extractTimeoutMs, signal);
      } catch (error) {
        const reason = classifyExtractError(error);
        log(`bundle extract failed (${reason}): ${String(error)}${describeStderrExcerpt(error)}`);
        rmSync(stageDir, { recursive: true, force: true });
        return { ok: false, reason };
      }
      const newHostDir = join(stageDir, "sand-host");
      if (!existsSync(join(newHostDir, "host-main.cjs"))) {
        log("staged bundle missing sand-host/host-main.cjs; aborting swap");
        rmSync(stageDir, { recursive: true, force: true });
        return { ok: false, reason: "host-main-missing" };
      }
      const stagedVersion = readVersionMarker(join(newHostDir, "version"));
      const normalizedExpected =
        typeof expectedVersion === "string" && expectedVersion.trim().length > 0
          ? expectedVersion.trim()
          : null;
      if (normalizedExpected != null && stagedVersion !== normalizedExpected) {
        const reason = stagedVersion == null ? "version-unreadable" : "version-mismatch";
        log(
          `refusing swap: staged version ${
            stagedVersion == null ? "unreadable" : `"${stagedVersion}"`
          } != target "${normalizedExpected}" (${reason}); host left untouched`,
        );
        rmSync(stageDir, { recursive: true, force: true });
        return { ok: false, reason };
      }
      if (backupDir != null) {
        try {
          if (!keepExistingHostBackup(backupDir, hostDir)) {
            rmSync(backupDir, { recursive: true, force: true });
            if (existsSync(join(hostDir, "host-main.cjs"))) {
              mkdirSync(backupDir, { recursive: true });
              for (const name of readdirSync(hostDir)) {
                if (name.endsWith(".map")) continue;
                cpSync(join(hostDir, name), join(backupDir, name), {
                  recursive: true,
                });
              }
            }
          }
        } catch (error) {
          log(
            `pre-swap backup failed (${classifySwapError(error)}): ${String(error)}; refusing swap (host untouched)`,
          );
          safeRmRecursive(stageDir);
          safeRmRecursive(backupDir);
          return { ok: false, reason: "backup-failed" };
        }
      }
      mkdirSync(hostDir, { recursive: true });
      publishBundleEntry(join(newHostDir, "host-main.cjs"), join(hostDir, "host-main.cjs"));
      entryCommitted = true;
      const staged = readdirSync(newHostDir);
      const failedEntries = [];
      for (const name of staged) {
        if (name === "host-main.cjs" || name === "version") continue;
        try {
          publishBundleEntry(join(newHostDir, name), join(hostDir, name));
        } catch (error) {
          failedEntries.push({ name, error });
          log(`host bundle entry publish failed for "${name}": ${String(error)}`);
        }
      }
      if (failedEntries.length > 0) {
        const reason = classifySwapError(failedEntries[0].error);
        const failedNames = failedEntries.map((entry) => entry.name).join(", ");
        log(
          `host bundle publish incomplete: ${failedEntries.length} required entr${
            failedEntries.length === 1 ? "y" : "ies"
          } failed (${failedNames}); swap FAILED (${reason}); version NOT advanced, no prune`,
        );
        rmSync(stageDir, { recursive: true, force: true });
        return { ok: false, reason, mutated: true };
      }
      if (staged.includes("version")) {
        try {
          publishBundleEntry(join(newHostDir, "version"), join(hostDir, "version"));
        } catch (error) {
          const reason = classifySwapError(error);
          log(
            `host bundle version flip failed (${reason}): ${String(error)}; version NOT advanced, no prune`,
          );
          rmSync(stageDir, { recursive: true, force: true });
          return { ok: false, reason, mutated: true };
        }
      }
      for (const name of readdirSync(hostDir)) {
        if (name === "host-main.cjs") continue;
        if (staged.includes(name)) continue;
        try {
          rmSync(join(hostDir, name), { recursive: true, force: true });
        } catch (error) {
          log(`host bundle stale-entry prune failed for "${name}": ${String(error)}`);
        }
      }
      rmSync(stageDir, { recursive: true, force: true });
      safeUnlink(bundlePath);
      log(`swapped in new host bundle from ${bundlePath}`);
      return { ok: true };
    } catch (error) {
      const reason = classifySwapError(error);
      log(`bundle swap failed (${reason}): ${String(error)}`);
      rmSync(stageDir, { recursive: true, force: true });
      return entryCommitted ? { ok: false, reason, mutated: true } : { ok: false, reason };
    }
  }

  function restoreHostBundleFromBackup({ backupDir, hostDir }) {
    if (!existsSync(join(backupDir, "host-main.cjs"))) {
      return { ok: false, reason: "backup-missing" };
    }
    const restoreStageDir = `${hostDir}.rollback-stage`;
    try {
      rmSync(restoreStageDir, { recursive: true, force: true });
      mkdirSync(restoreStageDir, { recursive: true });
      const names = readdirSync(backupDir);
      for (const name of names) {
        cpSync(join(backupDir, name), join(restoreStageDir, name), {
          recursive: true,
        });
      }
      publishBundleEntry(join(restoreStageDir, "host-main.cjs"), join(hostDir, "host-main.cjs"));
      const failedEntries = [];
      for (const name of names) {
        if (name === "host-main.cjs" || name === "version") continue;
        try {
          publishBundleEntry(join(restoreStageDir, name), join(hostDir, name));
        } catch (error) {
          failedEntries.push({ name, error });
          log(`rollback entry publish failed for "${name}": ${String(error)}`);
        }
      }
      if (failedEntries.length === 0 && names.includes("version")) {
        publishBundleEntry(join(restoreStageDir, "version"), join(hostDir, "version"));
      }
      if (failedEntries.length > 0) {
        const reason = classifySwapError(failedEntries[0].error);
        log(
          `rollback incomplete: ${failedEntries.length} entr${
            failedEntries.length === 1 ? "y" : "ies"
          } failed (${reason}); version not restored, backup kept`,
        );
        safeRmRecursive(restoreStageDir);
        return { ok: false, reason };
      }
      for (const name of readdirSync(hostDir)) {
        if (name === "host-main.cjs") continue;
        if (names.includes(name)) continue;
        try {
          rmSync(join(hostDir, name), { recursive: true, force: true });
        } catch (error) {
          log(`rollback stale-entry prune failed for "${name}": ${String(error)}`);
        }
      }
      safeRmRecursive(restoreStageDir);
      safeRmRecursive(backupDir);
      return { ok: true };
    } catch (error) {
      const reason = classifySwapError(error);
      log(`rollback failed (${reason}): ${String(error)}; backup kept for retry`);
      safeRmRecursive(restoreStageDir);
      return { ok: false, reason };
    }
  }

  function syncBoxScriptsFromDir({
    sourceDir,
    binDir,
    markerPath,
    bundleVersion,
    denyNames = BOX_SCRIPTS_DENY,
    logFn = log,
  }) {
    if (bundleVersion == null || !BOOT_FETCH_GIT_SHA_REGEX.test(bundleVersion)) {
      return { status: "noop", reason: "no valid bundle version" };
    }
    if (readVersionMarker(markerPath) === bundleVersion) {
      return { status: "noop", reason: "marker current" };
    }
    const sourceEntries = attempt(() => readdirSync(sourceDir));
    if (!sourceEntries.ok) {
      return { status: "noop", reason: "bundle carries no box-scripts" };
    }
    const eligible = [];
    for (const name of sourceEntries.value.sort()) {
      const sourcePath = join(sourceDir, name);
      const stat = attempt(() => lstatSync(sourcePath));
      if (!stat.ok) continue;
      if (!stat.value.isFile()) {
        logFn(`box-script sync: refusing non-regular entry "${name}"`);
        continue;
      }
      const denied = typeof denyNames === "function" ? denyNames(name) : denyNames.includes(name);
      if (denied) {
        logFn(`box-script sync: refusing Zone A overlay of "${name}"`);
        continue;
      }
      eligible.push(name);
    }
    const stagedPath = (name) => join(binDir, `.${name}.sand-scripts.${process.pid}`);
    const snapshotPath = (name) => join(binDir, `.${name}.sand-scripts-floor.${process.pid}`);
    const staged = [];
    for (const name of eligible) {
      try {
        copyFileSync(join(sourceDir, name), stagedPath(name));
        chmodSync(stagedPath(name), 0o755);
        staged.push(name);
      } catch (error) {
        for (const stagedName of staged) safeUnlink(stagedPath(stagedName));
        safeUnlink(stagedPath(name));
        logFn(`box-script sync: could not stage "${name}" (${String(error)}); floor left whole`);
        return { status: "failed", reason: "stage failed" };
      }
    }
    const published = [];
    let publishFailed = null;
    for (const name of staged) {
      try {
        if (existsSync(join(binDir, name))) {
          copyFileSync(join(binDir, name), snapshotPath(name));
        }
        renameSync(stagedPath(name), join(binDir, name));
        published.push(name);
      } catch (error) {
        publishFailed = { name, error };
        break;
      }
    }
    if (publishFailed != null) {
      let rollbackFailed = null;
      for (const name of published) {
        try {
          if (existsSync(snapshotPath(name))) {
            renameSync(snapshotPath(name), join(binDir, name));
          } else {
            safeUnlink(join(binDir, name));
          }
        } catch (error) {
          rollbackFailed = name;
          logFn(
            `box-script sync: ROLLBACK of "${name}" failed (${String(error)}); floor snapshot kept at ${snapshotPath(name)}`,
          );
        }
      }
      for (const name of staged) safeUnlink(stagedPath(name));
      safeUnlink(snapshotPath(publishFailed.name));
      logFn(
        `box-script sync: could not publish "${publishFailed.name}" (${String(
          publishFailed.error,
        )}); ${
          rollbackFailed == null
            ? "rolled the floor back whole"
            : `rollback of "${rollbackFailed}" failed`
        } — marker unchanged, the next attempt re-runs whole`,
      );
      return { status: "failed", reason: "publish failed" };
    }
    for (const name of staged) safeUnlink(snapshotPath(name));
    try {
      ensureDir(join(markerPath, ".."));
      const markerTmp = `${markerPath}.${process.pid}`;
      writeFileSync(markerTmp, bundleVersion);
      renameSync(markerTmp, markerPath);
    } catch (error) {
      logFn(
        `box-script sync: synced ${staged.length} scripts to ${bundleVersion} but could not write the marker (${String(
          error,
        )}); the next tick re-runs the sync whole`,
      );
      return { status: "failed", reason: "marker write failed" };
    }
    logFn(
      `box-script sync: converged ${staged.length} script${
        staged.length === 1 ? "" : "s"
      } to bundle ${bundleVersion}`,
    );
    return { status: "synced", reason: "ok" };
  }

  return {
    keepExistingHostBackup,
    swapHostBundle,
    restoreHostBundleFromBackup,
    syncBoxScriptsFromDir,
  };
}
