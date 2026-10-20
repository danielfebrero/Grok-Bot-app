const __mod=require('node:module');const __p=require('node:path');const __depsDir=__p.join(__dirname,'..','deps');process.env.NODE_PATH=__depsDir+(process.env.NODE_PATH?__p.delimiter+process.env.NODE_PATH:'');__mod.Module._initPaths();const __import_meta_url=require('node:url').pathToFileURL(__filename).href;
"use strict";
var __create = Object.create;
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __getProtoOf = Object.getPrototypeOf;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(
  // If the importer is in node compatibility mode or this is not an ESM
  // file that has been converted to a CommonJS file using a Babel-
  // compatible transform (i.e. "__esModule" has not been set), then set
  // "default" to the CommonJS "module.exports" for node compatibility.
  isNodeMode || !mod || !mod.__esModule ? __defProp(target, "default", { value: mod, enumerable: true }) : target,
  mod
));

// src/host/extensions/box-store-sync/box-store-vacuum-worker.ts
var import_node_worker_threads = require("node:worker_threads");

// src/shared/invariant.ts
var SandInvariantViolation = class extends Error {
  constructor(message) {
    super(message);
    this.name = "SandInvariantViolation";
  }
};
var installedReporter = null;
var STRIPPED_MESSAGE = "Invariant violation (message stripped in packaged builds; the stack identifies the site)";
function messagesStripped() {
  return false;
}
var FRAME_LINE = /^at /;
function topApplicationFrame(violation) {
  const stack = violation.stack;
  if (typeof stack !== "string" || !stack.startsWith(headerOf(violation))) return null;
  for (const raw of stack.slice(headerOf(violation).length).split("\n")) {
    const frame = raw.trim();
    if (!FRAME_LINE.test(frame)) continue;
    return frame;
  }
  return null;
}
function headerOf(violation) {
  return violation.message === "" ? violation.name : `${violation.name}: ${violation.message}`;
}
function invariant(condition, message) {
  if (condition) return;
  failInvariant(message, invariant);
}
function failInvariant(message, boundary) {
  let violationMessage;
  if (messagesStripped()) {
    violationMessage = STRIPPED_MESSAGE;
  } else if (typeof message === "function") {
    violationMessage = message();
  } else {
    violationMessage = message;
  }
  const violation = new SandInvariantViolation(violationMessage);
  let frame = null;
  if ("captureStackTrace" in Error && typeof Error.captureStackTrace === "function") {
    Error.captureStackTrace(violation, boundary);
    if (installedReporter !== null) frame = topApplicationFrame(violation);
  }
  installedReporter?.({ name: violation.name, frame });
  throw violation;
}

// src/host/extensions/box-store-sync/sqlite-snapshot.ts
var import_node_sqlite2 = require("node:sqlite");

// ../dune/atomic-write/dist/internal/atomic-write.js
var import_node_fs = __toESM(require("node:fs"), 1);
var fs = typeof process.getBuiltinModule === "function" ? process.getBuiltinModule("node:fs") : import_node_fs.default;
var { O_RDONLY, O_WRONLY, O_CREAT, O_EXCL, O_DIRECTORY = 0, O_NOFOLLOW = 0, O_NONBLOCK = 0 } = fs.constants;
var TEMP_CREATE = O_WRONLY | O_CREAT | O_EXCL | O_NOFOLLOW;
var NO_FOLLOW_READ = O_RDONLY | O_NOFOLLOW | O_NONBLOCK;
var NO_FOLLOW_DIRECTORY = O_RDONLY | O_DIRECTORY | O_NOFOLLOW;
var TEMP_SUFFIX_BOUND = 2 ** 32;

// ../dune/atomic-write/dist/internal/durable-file.js
var import_node_fs2 = __toESM(require("node:fs"), 1);
var fs2 = typeof process.getBuiltinModule === "function" ? process.getBuiltinModule("node:fs") : import_node_fs2.default;

// src/shared/errors/bounded.ts
function brandedEnumOf(values, fallback) {
  const admitted = new Set(values);
  return (value) => {
    if (value === void 0) return void 0;
    return admitted.has(value) ? value : fallback;
  };
}
var DNS_LABEL = "[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?";
var HOST = new RegExp(`^(?=.{1,253}$)${DNS_LABEL}(?:\\.${DNS_LABEL})*$`);
var CAPPED_TEXT_CAPS = [200, 3276];
var CAPPED_TEXT_CEILING = Math.max(...CAPPED_TEXT_CAPS);
var CONNECT_CODE_TAGS = [
  "Canceled",
  "Unknown",
  "InvalidArgument",
  "DeadlineExceeded",
  "NotFound",
  "AlreadyExists",
  "PermissionDenied",
  "ResourceExhausted",
  "FailedPrecondition",
  "Aborted",
  "OutOfRange",
  "Unimplemented",
  "Internal",
  "Unavailable",
  "DataLoss",
  "Unauthenticated"
];
var CONNECT_CODE_FALLBACK = "Other";
var brandedConnectCode = brandedEnumOf(CONNECT_CODE_TAGS, CONNECT_CODE_FALLBACK);
var SAND_ERRNO_TAGS = [
  "ECONNREFUSED",
  "ECONNRESET",
  "ECONNABORTED",
  "ETIMEDOUT",
  "EPIPE",
  "ENETRESET",
  "ENETDOWN",
  "ENETUNREACH",
  "EHOSTUNREACH",
  "EHOSTDOWN",
  "EAI_AGAIN",
  "ENOTFOUND",
  "EADDRINUSE",
  "EACCES",
  "EPERM",
  "ENOENT",
  "ENOSPC",
  "EDQUOT",
  "EROFS",
  "EBUSY",
  "EMFILE",
  "EIO"
];
var SAND_ERRNO_FALLBACK = "E_OTHER";
var brandedErrno = brandedEnumOf(
  SAND_ERRNO_TAGS,
  SAND_ERRNO_FALLBACK
);
var BOUNDED_TELEMETRY_ERROR_NAMES = [
  "Error",
  "TypeError",
  "AbortError",
  "DeadlineExceededError",
  "ConnectError"
];
var BOUNDED_TELEMETRY_TOKENS = [
  "client-paused",
  "away-park",
  "present-park",
  "dev-induced-offline",
  "network-return",
  "unknown",
  "manual_update",
  "manual_reset",
  "auto_update",
  "remote_reset",
  "reconnect",
  ...BOUNDED_TELEMETRY_ERROR_NAMES,
  ...BOUNDED_TELEMETRY_ERROR_NAMES.flatMap(
    (name) => SAND_ERRNO_TAGS.map((errno) => `${name}/${errno}`)
  )
];
var brandedTelemetryToken = brandedEnumOf(BOUNDED_TELEMETRY_TOKENS, "unknown");

// src/host/durable-file-policy.ts
var SAND_UPGRADE_RESUME_FILE_NAME = "host-upgrade-resume.json";
var SAND_ACK_OBLIGATIONS_FILE_NAME = "ack-obligations.json";
var SAND_PENDING_WAKE_FILE_NAME = "host-pending-wakes.json";
var SAND_DISK_PRESSURE_REMINDERS_FILE_NAME = "host-disk-pressure-reminders.json";
var SAND_DATA_HOST_OWNED_FILE_NAMES = [
  "gateway.json",
  "webhook-keys.json",
  "teach-queue-key.json",
  "box-secrets.json"
];
var BOX_STORE_SAND_DATA_EXCLUDED_FILE_NAMES = [
  SAND_UPGRADE_RESUME_FILE_NAME,
  SAND_ACK_OBLIGATIONS_FILE_NAME,
  SAND_PENDING_WAKE_FILE_NAME,
  SAND_DISK_PRESSURE_REMINDERS_FILE_NAME,
  "sand-statsig-bootstrap.json",
  ...SAND_DATA_HOST_OWNED_FILE_NAMES.map((name) => `${name}.*`)
];

// ../dune/scheduling/dist/internal/policies.js
var JITTER_SPREAD = { none: 0, equal: 1 / 2, full: 1 };

// src/host/storage/store-db.ts
var import_node_sqlite = require("node:sqlite");

// ../packages/grok-bot-harness/src/host-paths.ts
var SAND_DATA_DIRNAME = "sand-data";
var SAND_BOX_HOME_DIR = "/home/box";
var SAND_BOX_DATA_ROOT = `${SAND_BOX_HOME_DIR}/${SAND_DATA_DIRNAME}`;
var SAND_BOX_MODEL_VISIBLE_DATA_ROOT = `${SAND_BOX_HOME_DIR}/agent-data`;
var SAND_BOX_REFERENCE_DIR = `${SAND_BOX_HOME_DIR}/reference`;
var LEGACY_SAND_BOX_REFERENCE_DIR = `${SAND_BOX_HOME_DIR}/sand-reference`;
var DEBUGGING_THE_BOX_FILE = "debugging-the-box.md";
var SAND_APP_UI_FILE = "app-ui.md";
var SAND_APP_ACTIONS_FILE = "app-actions.md";
var SAND_BOX_DEBUGGING_REFERENCE_PATH = `${SAND_BOX_REFERENCE_DIR}/${DEBUGGING_THE_BOX_FILE}`;
var SAND_APP_UI_REFERENCE_PATH = `${SAND_BOX_REFERENCE_DIR}/${SAND_APP_UI_FILE}`;
var SAND_APP_ACTIONS_REFERENCE_PATH = `${SAND_BOX_REFERENCE_DIR}/${SAND_APP_ACTIONS_FILE}`;

// src/host/no-follow-walk.ts
var import_node_fs3 = require("node:fs");
var NO_FOLLOW_DIRECTORY2 = import_node_fs3.constants.O_RDONLY | import_node_fs3.constants.O_DIRECTORY | import_node_fs3.constants.O_NOFOLLOW;
var FOLLOW_DIRECTORY = import_node_fs3.constants.O_RDONLY | import_node_fs3.constants.O_DIRECTORY;

// src/host/sand-data-files.ts
var SAND_DATA_JSON_MAX_BYTES = 16 * 1024 * 1024;

// src/host/storage/store-db.ts
var DB_BUSY_TIMEOUT_MS = 5e3;

// src/host/extensions/box-store-sync/sqlite-snapshot.ts
function sqliteVacuumInto(srcPath, destPath, busyTimeoutMs = DB_BUSY_TIMEOUT_MS) {
  const db = new import_node_sqlite2.DatabaseSync(srcPath, { readOnly: true });
  try {
    db.exec(`PRAGMA busy_timeout = ${busyTimeoutMs}`);
    db.exec(`VACUUM INTO '${destPath.replace(/'/g, "''")}'`);
  } finally {
    db.close();
  }
}

// src/host/extensions/box-store-sync/box-store-vacuum-worker.ts
var port = import_node_worker_threads.parentPort;
invariant(port != null, "box-store-vacuum-worker must run as a worker_thread");
port.on("message", (job) => {
  try {
    sqliteVacuumInto(job.srcPath, job.destPath, job.busyTimeoutMs);
    port.postMessage({ ok: true });
  } catch (error) {
    port.postMessage({
      ok: false,
      message: error instanceof Error ? error.message : String(error)
    });
  }
});
