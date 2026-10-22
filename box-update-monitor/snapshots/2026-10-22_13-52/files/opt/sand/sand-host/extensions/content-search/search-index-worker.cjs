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

// src/host/extensions/content-search/search-index-worker.ts
var import_node_worker_threads = require("node:worker_threads");

// src/shared/errors/system-errno.ts
function findErrorCode(error, pattern) {
  const seen = /* @__PURE__ */ new Set();
  let current = error;
  while (current != null && typeof current === "object" && !seen.has(current)) {
    seen.add(current);
    const code = current.code;
    if (typeof code === "string" && pattern.test(code)) return code;
    current = current.cause;
  }
  return void 0;
}
function findSystemErrno(error) {
  return findErrorCode(error, /^E[A-Z_]+$/);
}

// src/shared/errors/errors.ts
function errorMessage(error) {
  return error instanceof Error ? error.message : String(error);
}
function errorLogTag(error) {
  if (!(error instanceof Error)) return typeof error;
  const ownCode = error.code;
  const code = ownCode != null && ownCode.length > 0 ? ownCode : findSystemErrno(error);
  return code !== void 0 ? `${error.name} (${code})` : error.name;
}

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

// ../dune/scheduling/dist/internal/policies.js
var JITTER_SPREAD = { none: 0, equal: 1 / 2, full: 1 };

// src/host/storage/sqlite-busy.ts
var SQLITE_BUSY = 5;
var SQLITE_LOCKED = 6;
var SQLITE_NOMEM = 7;
var SQLITE_IOERR = 10;
var SQLITE_CORRUPT = 11;
var SQLITE_CANTOPEN = 14;
var SQLITE_TOOBIG = 18;
var SQLITE_MISMATCH = 20;
var SQLITE_RANGE = 25;
var SQLITE_NOTADB = 26;
var SQLITE_PRIMARY_CODES = /* @__PURE__ */ new Map([
  ["SQLITE_BUSY", SQLITE_BUSY],
  ["SQLITE_LOCKED", SQLITE_LOCKED],
  ["SQLITE_IOERR", SQLITE_IOERR],
  ["SQLITE_CORRUPT", SQLITE_CORRUPT],
  ["SQLITE_CANTOPEN", SQLITE_CANTOPEN],
  ["SQLITE_NOTADB", SQLITE_NOTADB]
]);
function sqlitePrimaryCode(error) {
  if (!(error instanceof Error)) return void 0;
  if ("errcode" in error && typeof error.errcode === "number" && Number.isInteger(error.errcode)) {
    return error.errcode & 255;
  }
  if ("code" in error && typeof error.code === "string") {
    return SQLITE_PRIMARY_CODES.get(error.code.split("_", 2).join("_"));
  }
  return void 0;
}
function sqliteResultCode(error) {
  if (!(error instanceof Error) || !("errcode" in error)) return void 0;
  const { errcode } = error;
  return typeof errcode === "number" && Number.isSafeInteger(errcode) && errcode >= 0 ? errcode : void 0;
}
function isSqliteBusyError(error) {
  const code = sqlitePrimaryCode(error);
  return code === SQLITE_BUSY || code === SQLITE_LOCKED;
}
function isSqliteInputOrMemoryError(error) {
  const code = sqlitePrimaryCode(error);
  return code === SQLITE_NOMEM || code === SQLITE_TOOBIG || code === SQLITE_MISMATCH || code === SQLITE_RANGE;
}
function isSqliteCorruptError(error) {
  const code = sqlitePrimaryCode(error);
  return code === SQLITE_CORRUPT || code === SQLITE_NOTADB;
}

// src/host/extensions/content-search/search-index-db.ts
var import_node_path5 = require("node:path");
var import_node_sqlite2 = require("node:sqlite");

// src/shared/ids.ts
function isAgentId(raw) {
  return typeof raw === "string" && raw.length > 0 && raw === raw.trim() && !raw.includes("/") && !raw.includes("\\") && !raw.includes("\0") && raw !== "." && raw !== "..";
}
function parseAgentId(raw) {
  return isAgentId(raw) ? raw : null;
}

// src/shared/media/attachment-open-policy.ts
function attachmentExtension(nameOrPath) {
  const segments = nameOrPath.split(/[/\\]/);
  const base = segments[segments.length - 1] ?? "";
  const dot = base.lastIndexOf(".");
  if (dot <= 0 || dot === base.length - 1) return null;
  return base.slice(dot + 1).toLowerCase();
}

// src/shared/media/attachment-preview.ts
var TEXT_PREVIEWABLE_EXTENSIONS = /* @__PURE__ */ new Set([
  "txt",
  "text",
  "log",
  "md",
  "markdown",
  "mdx",
  "rst",
  "adoc",
  "tex",
  "json",
  "jsonc",
  "json5",
  "ndjson",
  "csv",
  "tsv",
  "xml",
  "yaml",
  "yml",
  "toml",
  "ini",
  "cfg",
  "conf",
  "env",
  "properties",
  "plist",
  "gradle",
  "html",
  "htm",
  "css",
  "scss",
  "sass",
  "less",
  "svg",
  "js",
  "jsx",
  "mjs",
  "cjs",
  "ts",
  "tsx",
  "mts",
  "cts",
  "py",
  "pyi",
  "rb",
  "go",
  "rs",
  "java",
  "kt",
  "kts",
  "c",
  "h",
  "cc",
  "cpp",
  "cxx",
  "hpp",
  "hh",
  "cs",
  "php",
  "swift",
  "scala",
  "dart",
  "lua",
  "pl",
  "pm",
  "r",
  "sql",
  "graphql",
  "gql",
  "proto",
  "vue",
  "svelte",
  "astro",
  "sh",
  "bash",
  "zsh",
  "fish",
  "bat",
  "ps1",
  "tf",
  "tfvars",
  "dockerfile",
  "diff",
  "patch"
]);
function isTextPreviewableName(nameOrPath) {
  const ext = attachmentExtension(nameOrPath);
  return ext != null && TEXT_PREVIEWABLE_EXTENSIONS.has(ext);
}
var BINARY_SNIFF_BYTE_WINDOW = 8 * 1024;

// src/shared/media/media-extensions.ts
function extensionOf(name) {
  const base = name.slice(Math.max(name.lastIndexOf("/"), name.lastIndexOf("\\")) + 1);
  const dot = base.lastIndexOf(".");
  return dot <= 0 ? "" : base.slice(dot).toLowerCase();
}
var IMAGE_MIME_FROM_EXTENSION = {
  ".avif": "image/avif",
  ".bmp": "image/bmp",
  ".gif": "image/gif",
  ".ico": "image/x-icon",
  ".jpeg": "image/jpeg",
  ".jpg": "image/jpeg",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".webp": "image/webp"
};
var CLIENT_NATIVE_IMAGE_MIME_FROM_EXTENSION = {
  ".heic": "image/heic",
  ".heif": "image/heif"
};
var SERVABLE_IMAGE_MIME_FROM_EXTENSION = {
  ...IMAGE_MIME_FROM_EXTENSION,
  ...CLIENT_NATIVE_IMAGE_MIME_FROM_EXTENSION
};
var VIDEO_MIME_FROM_EXTENSION = {
  ".m4v": "video/mp4",
  ".mov": "video/quicktime",
  ".mp4": "video/mp4",
  ".ogv": "video/ogg",
  ".webm": "video/webm"
};
var AUDIO_MIME_FROM_EXTENSION = {
  ".aac": "audio/aac",
  ".flac": "audio/flac",
  ".m4a": "audio/mp4",
  ".mp3": "audio/mpeg",
  ".oga": "audio/ogg",
  ".ogg": "audio/ogg",
  ".opus": "audio/ogg",
  ".wav": "audio/wav",
  ".weba": "audio/webm"
};

// src/shared/media/file-preview-kind.ts
var IMAGE_EXTENSIONS = new Set(
  Object.keys(SERVABLE_IMAGE_MIME_FROM_EXTENSION).map((ext) => ext.slice(1))
);
var VIDEO_EXTENSIONS = new Set(
  Object.keys(VIDEO_MIME_FROM_EXTENSION).map((ext) => ext.slice(1))
);
var AUDIO_EXTENSIONS = new Set(
  Object.keys(AUDIO_MIME_FROM_EXTENSION).map((ext) => ext.slice(1))
);
var MARKDOWN_EXTENSIONS = /* @__PURE__ */ new Set(["md", "markdown", "mdx"]);
var JSON_EXTENSIONS = /* @__PURE__ */ new Set(["json"]);
var HTML_EXTENSIONS = /* @__PURE__ */ new Set(["html", "htm"]);
var TABLE_EXTENSIONS = /* @__PURE__ */ new Set(["csv", "tsv", "xlsx", "xls"]);
function getFilePreviewKind(nameOrPath) {
  const ext = attachmentExtension(nameOrPath);
  if (ext == null) return "unknown";
  if (IMAGE_EXTENSIONS.has(ext)) return "image";
  if (VIDEO_EXTENSIONS.has(ext)) return "video";
  if (AUDIO_EXTENSIONS.has(ext)) return "audio";
  if (ext === "pdf") return "pdf";
  if (TABLE_EXTENSIONS.has(ext)) return "table";
  if (JSON_EXTENSIONS.has(ext)) return "json";
  if (MARKDOWN_EXTENSIONS.has(ext)) return "markdown";
  if (HTML_EXTENSIONS.has(ext)) return "html";
  if (ext === "docx") return "docx";
  if (isTextPreviewableName(nameOrPath)) return "text";
  return "unknown";
}

// src/shared/media/attachment-summary.ts
var JSON_EXTENSIONS2 = /* @__PURE__ */ new Set(["json", "jsonc", "json5", "ndjson"]);
var ARCHIVE_EXTENSIONS = /* @__PURE__ */ new Set([
  "zip",
  "tar",
  "gz",
  "tgz",
  "bz2",
  "tbz2",
  "xz",
  "txz",
  "zst",
  "7z",
  "rar"
]);
var TABLE_MIME_TYPES = /* @__PURE__ */ new Set([
  "text/csv",
  "text/tab-separated-values",
  "application/vnd.ms-excel",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
]);
var DOCUMENT_MIME_TYPES = /* @__PURE__ */ new Set([
  "application/msword",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
]);
var ARCHIVE_MIME_TYPES = /* @__PURE__ */ new Set([
  "application/zip",
  "application/x-zip-compressed",
  "application/gzip",
  "application/x-tar",
  "application/x-bzip2",
  "application/x-xz",
  "application/zstd",
  "application/x-7z-compressed",
  "application/x-rar-compressed",
  "application/vnd.rar"
]);
function classifyMimeType(rawMimeType) {
  const mime = (rawMimeType.split(";")[0] ?? "").trim().toLowerCase();
  if (mime.startsWith("image/")) return "image";
  if (mime.startsWith("video/")) return "video";
  if (mime.startsWith("audio/")) return "audio";
  if (mime === "application/pdf") return "pdf";
  if (mime === "text/markdown") return "markdown";
  if (TABLE_MIME_TYPES.has(mime)) return "table";
  if (mime === "application/json" || mime.endsWith("+json")) return "json";
  if (DOCUMENT_MIME_TYPES.has(mime)) return "document";
  if (ARCHIVE_MIME_TYPES.has(mime)) return "archive";
  if (mime.startsWith("text/")) return "text";
  return null;
}
function extensionSubject(source) {
  let pathname;
  try {
    pathname = new URL(source).pathname;
  } catch {
    return source;
  }
  try {
    return decodeURIComponent(pathname);
  } catch {
    return pathname;
  }
}
function classifyPathLike(source) {
  const subject = extensionSubject(source);
  const previewKind = getFilePreviewKind(subject);
  switch (previewKind) {
    case "image":
    case "video":
    case "audio":
    case "pdf":
    case "markdown":
    case "table":
      return previewKind;
    case "docx":
      return "document";
    case "html":
      return "text";
    case "json":
      return "json";
    case "text": {
      const ext = attachmentExtension(subject);
      return ext != null && JSON_EXTENSIONS2.has(ext) ? "json" : "text";
    }
    case "unknown": {
      const ext = attachmentExtension(subject);
      return ext != null && ARCHIVE_EXTENSIONS.has(ext) ? "archive" : null;
    }
  }
  const _exhaustive = previewKind;
  return _exhaustive;
}
function classifyAttachment(source) {
  if (source.mimeType != null && source.mimeType.length > 0) {
    const byMime = classifyMimeType(source.mimeType);
    if (byMime != null) return byMime;
  }
  if (source.fileName != null && source.fileName.length > 0) {
    const byName = classifyPathLike(source.fileName);
    if (byName != null) return byName;
  }
  if (source.urlOrPath != null && source.urlOrPath.length > 0) {
    const byPath = classifyPathLike(source.urlOrPath);
    if (byPath != null) return byPath;
  }
  return "file";
}

// src/shared/media/attachments.ts
var SAND_ATTACHMENT_KINDS = [
  "image",
  "video",
  "audio",
  "pdf",
  "markdown",
  "table",
  "json",
  "text",
  "document",
  "archive",
  "file"
];
var SAND_FALLBACK_ATTACHMENT_KIND = SAND_ATTACHMENT_KINDS[10];

// src/shared/media/image-mime.ts
function imageMimeFromPath(filePath) {
  return IMAGE_MIME_FROM_EXTENSION[extensionOf(filePath)];
}
function videoMimeFromPath(filePath) {
  return VIDEO_MIME_FROM_EXTENSION[extensionOf(filePath)];
}
function audioMimeFromPath(filePath) {
  return AUDIO_MIME_FROM_EXTENSION[extensionOf(filePath)];
}

// ../packages/constants/dist/meeting-channel.js
var MEETING_REQUEST_CHANNEL_PREFIX = "meeting-request:";
var LEGACY_REQUEST_INDEX = /:\d{1,9}$/;
var MeetingRequestChannelAddress = class _MeetingRequestChannelAddress {
  constructor(meetingId) {
    this.meetingId = meetingId;
  }
  static parse(channel) {
    const trimmed = channel?.trim();
    if (trimmed === void 0 || !trimmed.startsWith(MEETING_REQUEST_CHANNEL_PREFIX)) {
      return null;
    }
    const meetingId = trimmed.slice(MEETING_REQUEST_CHANNEL_PREFIX.length).replace(LEGACY_REQUEST_INDEX, "").trim();
    return meetingId.length === 0 ? null : new _MeetingRequestChannelAddress(meetingId);
  }
};
MeetingRequestChannelAddress.platform = "meeting-request";

// ../packages/voice-call-harness/dist/prompt/channel-platform.js
var VOICE_CALL_CHANNEL_PLATFORM = "voice";

// ../packages/voice-call-harness/dist/call/voice-channel.js
var ADDRESS_PREFIX = `${VOICE_CALL_CHANNEL_PLATFORM}:`;
var OWNER_STARTED_SOURCES = ["turn", "voice-call", "handoff-resume"];
var VoiceCallChannel = class {
  static address(callId) {
    return `${ADDRESS_PREFIX}${callId}`;
  }
  static callIdOf(channel) {
    if (channel === void 0 || !channel.startsWith(ADDRESS_PREFIX))
      return null;
    const callId = channel.slice(ADDRESS_PREFIX.length).trim();
    return callId.length === 0 ? null : callId;
  }
  static isOwnerStarted(requestSource) {
    return requestSource !== void 0 && OWNER_STARTED_SOURCES.includes(requestSource);
  }
};

// src/shared/agents/bot-templates.ts
function nonEmptyTrimmed(value) {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}
function parseDraftJson(command) {
  try {
    return { kind: "parsed", value: JSON.parse(command) };
  } catch {
    return { kind: "invalid" };
  }
}
function parseBotTemplateDraftCommand(command) {
  if (command == null) return null;
  const parsed = parseDraftJson(command);
  if (parsed.kind === "invalid") return null;
  const raw = parsed.value;
  if (raw == null || typeof raw !== "object") return null;
  const name = nonEmptyTrimmed("name" in raw ? raw.name : null);
  const avatarShape = nonEmptyTrimmed("avatar_shape" in raw ? raw.avatar_shape : null);
  const avatarColor = nonEmptyTrimmed("avatar_color" in raw ? raw.avatar_color : null);
  const body = nonEmptyTrimmed("body" in raw ? raw.body : null);
  if (name == null || avatarShape == null || avatarColor == null || body == null) return null;
  const visibilityRaw = "visibility" in raw ? raw.visibility : null;
  const visibility = visibilityRaw === "public" || visibilityRaw === "team" ? visibilityRaw : void 0;
  return {
    name,
    avatarShape,
    avatarColor,
    body,
    ...visibility == null ? {} : { visibility }
  };
}

// src/shared/transcript/transcript.ts
function isExpiredLocalToolPermissionEntry(entry) {
  return entry.kind === "send-message" && entry.message.type === "local-tool-permission" && entry.message.ask.status === "expired";
}
function isBotShareAutoReviewEntry(entry) {
  return entry.kind === "send-message" && entry.message.type === "auto-review-approval" && entry.message.approval.surface === "bot_share" && parseBotTemplateDraftCommand(entry.message.approval.command) != null;
}
function isBranchedEntry(entry) {
  if (entry.kind === "message" || entry.kind === "send-message" || entry.kind === "user-attachment" || entry.kind === "notice") {
    return entry.branched === true;
  }
  return false;
}
function isOutboundAgentPeerMessageEntry(entry) {
  return entry != null && entry.kind === "message" && entry.toAgent != null;
}
function isHiddenOutboundAgentPeerMessageEntry(entry) {
  return isOutboundAgentPeerMessageEntry(entry) && entry.toAgent.kind !== "agent" && entry.toAgent.kind !== "cloud-agent";
}
function channelAddressOf(entry) {
  if (entry === void 0) return void 0;
  if (entry.kind === "message") return entry.channel;
  if (entry.kind === "send-message" && (entry.message.type === "text" || entry.message.type === "attachment" || entry.message.type === "visualization")) {
    return entry.message.channel;
  }
  return void 0;
}
function isVoiceCallChannelEntry(entry) {
  return VoiceCallChannel.callIdOf(channelAddressOf(entry)) !== null;
}
var MEETING_CHANNEL_PREFIX = "meeting:";
function isMeetingChannelEntry(entry) {
  const channel = channelAddressOf(entry);
  if (channel === void 0 || !channel.startsWith(MEETING_CHANNEL_PREFIX)) return false;
  return channel.slice(MEETING_CHANNEL_PREFIX.length).trim().length > 0;
}
function isMeetingRequestChannelEntry(entry) {
  return MeetingRequestChannelAddress.parse(channelAddressOf(entry)) !== null;
}
function isOffChatChannelEntry(entry) {
  return isVoiceCallChannelEntry(entry) || isMeetingChannelEntry(entry) || isMeetingRequestChannelEntry(entry);
}

// src/shared/transcript/silent-wake.ts
function isSuppressedEntry(entry) {
  return entry.kind === "message" && entry.suppressed != null;
}

// src/shared/transcript/timeline.ts
var SAND_TEAM_SETUP_SLOT_KINDS = [
  "plugins",
  "secrets",
  "skills",
  "memories",
  "automations",
  "files",
  "description"
];
function bySlotOrder(a, b) {
  return SAND_TEAM_SETUP_SLOT_KINDS.indexOf(a.kind) - SAND_TEAM_SETUP_SLOT_KINDS.indexOf(b.kind);
}
function teamSharedSlots(event) {
  if (event.slots !== void 0) return event.slots;
  const memories = {
    kind: "memories",
    status: event.status,
    ...event.toTeam === void 0 ? {} : { toTeam: event.toTeam },
    ...event.justYou === void 0 ? {} : { justYou: event.justYou },
    ...event.code === void 0 ? {} : { code: event.code }
  };
  const description = event.description === void 0 ? [] : [{ kind: "description", status: "done", description: event.description }];
  return [memories, ...description];
}
function teamSharedSlot(event, kind) {
  return teamSharedSlots(event).find(
    (slot) => slot.kind === kind
  );
}
function isTeamSetupSlotUnresolved(slot) {
  switch (slot.kind) {
    case "plugins":
      return slot.status === "proposing" || slot.status === "proposed";
    case "skills":
      return slot.status === "pending" || slot.status === "proposed";
    case "files":
    case "secrets":
      return slot.status === "proposed";
    case "automations":
      return slot.status === "proposed" && (slot.automations?.length ?? 0) > 0;
    case "memories":
      return slot.status === "sorting" || slot.status === "proposed";
    case "description":
      return slot.status === "pending";
  }
}
function isTeamSharedEventUnresolved(event) {
  return teamSharedSlots(event).some(isTeamSetupSlotUnresolved);
}
function teamSharedSlotRendersCard(slot) {
  switch (slot.kind) {
    case "description":
      return false;
    case "plugins":
      return slot.status !== "empty" && slot.status !== "proposing";
    case "skills":
      return slot.status !== "empty";
    case "memories":
      return slot.status !== "done";
    case "automations":
      return (slot.automations?.length ?? 0) > 0;
    case "secrets":
    case "files":
      return true;
  }
}
function isBlankTeamSharedEvent(event) {
  if (event.unshared === true || event.superseded === true) return false;
  if (orderedTeamSharedSlots(event).some(teamSharedSlotRendersCard)) return false;
  const plugins = teamSharedSlot(event, "plugins");
  if (event.setup === true && plugins?.status === "proposing") return false;
  return event.setup === true || isTeamSharedEventUnresolved(event);
}
function orderedTeamSharedSlots(event) {
  return [...teamSharedSlots(event)].sort(bySlotOrder);
}
var TEAM_SETUP_CARD_KINDS = new Set(
  SAND_TEAM_SETUP_SLOT_KINDS.filter((kind) => kind !== "description")
);

// src/shared/transcript/transcript-window-edge.ts
var BACKGROUND_TASK_STOP_NOTICE_PREFIX = "A background task was stopped ";
function isBackgroundTaskStopNotice(entry) {
  return entry.kind === "notice" && entry.text.startsWith(BACKGROUND_TASK_STOP_NOTICE_PREFIX);
}
function isEmptyStreamingPreviewEntry(entry) {
  return entry.kind === "send-message" && entry.streaming === true && entry.message.type === "text" && entry.message.content.length === 0;
}
function isTimelineHiddenEntry(entry) {
  return isEmptyStreamingPreviewEntry(entry) || isSuppressedEntry(entry) || entry.kind === "tool-call" || isBackgroundTaskStopNotice(entry) || isExpiredLocalToolPermissionEntry(entry) || isBotShareAutoReviewEntry(entry) || isHiddenOutboundAgentPeerMessageEntry(entry) || isOffChatChannelEntry(entry) || entry.kind === "event" && entry.event.type === "desktop-goal" || entry.kind === "event" && entry.event.type === "team-shared" && isBlankTeamSharedEvent(entry.event);
}

// src/host/sand-data-files.ts
var import_node_fs4 = require("node:fs");
var import_node_path4 = require("node:path");

// ../dune/atomic-write/dist/internal/atomic-write.js
var import_node_fs = __toESM(require("node:fs"), 1);
var REFUSED = {
  "target-link": "a symbolic link at the target path",
  "parent-link": "a symbolic link at the parent path",
  "not-a-file": "a non-regular file at the target path"
};
var NoFollowRefusedError = class extends Error {
  constructor(reason, path) {
    super(`refused ${REFUSED[reason]} ${path}`);
    this.name = "NoFollowRefusedError";
    this.reason = reason;
    this.path = path;
  }
};
var fs = typeof process.getBuiltinModule === "function" ? process.getBuiltinModule("node:fs") : import_node_fs.default;
var { O_RDONLY, O_WRONLY, O_CREAT, O_EXCL, O_DIRECTORY = 0, O_NOFOLLOW = 0, O_NONBLOCK = 0 } = fs.constants;
var TEMP_CREATE = O_WRONLY | O_CREAT | O_EXCL | O_NOFOLLOW;
var NO_FOLLOW_READ = O_RDONLY | O_NOFOLLOW | O_NONBLOCK;
var NO_FOLLOW_DIRECTORY = O_RDONLY | O_DIRECTORY | O_NOFOLLOW;
var TEMP_SUFFIX_BOUND = 2 ** 32;

// ../dune/atomic-write/dist/internal/durable-file.js
var import_node_fs2 = __toESM(require("node:fs"), 1);
var fs2 = typeof process.getBuiltinModule === "function" ? process.getBuiltinModule("node:fs") : import_node_fs2.default;

// ../packages/grok-bot-harness/src/host-paths.ts
var import_node_os = require("node:os");
var import_node_path2 = require("node:path");

// src/shared/node/paths.ts
var import_node_path = require("node:path");
function isPathWithin(parent, child, opts = {}) {
  const rel = (0, import_node_path.relative)((0, import_node_path.resolve)(parent), (0, import_node_path.resolve)(child));
  if (rel === "") return opts.isInclusive ?? false;
  return !rel.startsWith("..") && !(0, import_node_path.isAbsolute)(rel);
}

// src/shared/node/sand-variant.ts
function sandVariantOf(packaged, labBuild) {
  if (!packaged) return "sand-dev";
  if (labBuild) return "sand-lab";
  return "sand";
}
function getSandVariant() {
  return sandVariantOf(process.env.SAND_PACKAGED === "1", process.env.SAND_LAB === "1");
}

// ../packages/grok-bot-harness/src/host-paths.ts
var SAND_DATA_ROOT_ENV = "SAND_DATA_ROOT";
var SAND_PRODUCTION_DATA_DIRNAME = ".grokbot";
var SAND_USER_DATA_DIR_ENV = "SAND_USER_DATA_DIR";
var SAND_DATA_DIRNAME = "sand-data";
var USER_DATA_DIR_FLAG = "--user-data-dir";
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
function readUserDataDirArg(argv) {
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === USER_DATA_DIR_FLAG) {
      const next = argv[i + 1];
      return next != null && !next.startsWith("--") ? next : null;
    }
    const prefix = `${USER_DATA_DIR_FLAG}=`;
    if (arg.startsWith(prefix)) return arg.slice(prefix.length);
  }
  return null;
}
function resolveSandUserDataDir(argv = [], env = process.env, cwd = process.cwd()) {
  const raw = readUserDataDirArg(argv) ?? env[SAND_USER_DATA_DIR_ENV];
  const trimmed = raw?.trim();
  if (trimmed == null || trimmed.length === 0) return null;
  return (0, import_node_path2.isAbsolute)(trimmed) ? trimmed : (0, import_node_path2.resolve)(cwd, trimmed);
}
function getSandProductionRootDir(homeDir = (0, import_node_os.homedir)()) {
  return (0, import_node_path2.join)(homeDir, SAND_PRODUCTION_DATA_DIRNAME);
}
function resolveSandDataRootOverride(env = process.env) {
  const override = env[SAND_DATA_ROOT_ENV]?.trim();
  return override != null && override.length > 0 && (0, import_node_path2.isAbsolute)(override) ? override : null;
}
function getSandRootDir(homeDir = (0, import_node_os.homedir)()) {
  const override = resolveSandDataRootOverride();
  if (override != null) return override;
  const userDataDir = resolveSandUserDataDir([], process.env);
  if (userDataDir != null) {
    return (0, import_node_path2.join)(userDataDir, SAND_DATA_DIRNAME);
  }
  const variant = getSandVariant();
  return variant === "sand" ? getSandProductionRootDir(homeDir) : (0, import_node_path2.join)(homeDir, ".cursor", variant);
}

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
function hostNoFollowRoot(nonLinuxRoot) {
  return process.platform === "linux" ? "/" : nonLinuxRoot;
}

// src/host/no-follow-walk.ts
var import_node_fs3 = require("node:fs");
var import_node_path3 = require("node:path");
var NO_FOLLOW_DIRECTORY2 = import_node_fs3.constants.O_RDONLY | import_node_fs3.constants.O_DIRECTORY | import_node_fs3.constants.O_NOFOLLOW;
var FOLLOW_DIRECTORY = import_node_fs3.constants.O_RDONLY | import_node_fs3.constants.O_DIRECTORY;
function segmentsBelow(root, dir) {
  const rel = (0, import_node_path3.relative)(root, dir);
  if (rel === "") return [];
  invariant(
    rel !== ".." && !rel.startsWith(`..${import_node_path3.sep}`) && !(0, import_node_path3.isAbsolute)(rel),
    "a no-follow walk must stay below its root"
  );
  return rel.split(import_node_path3.sep);
}

// src/host/sand-data-files.ts
var SAND_DATA_JSON_MAX_BYTES = 16 * 1024 * 1024;
var SQLITE_FILE_SUFFIXES = ["", "-wal", "-shm", "-journal"];
function refuseDatabaseFileLinks(parentFd, path) {
  const name = (0, import_node_path4.basename)(path);
  for (const suffix of SQLITE_FILE_SUFFIXES) {
    const entry = entryOf(parentFd, `${path}${suffix}`, `${name}${suffix}`);
    if ((0, import_node_fs4.lstatSync)(entry, { throwIfNoEntry: false })?.isSymbolicLink() === true) {
      throw new NoFollowRefusedError("target-link", `${path}${suffix}`);
    }
  }
}
function refuseSandDataDatabaseLinksSync(target) {
  const path = (0, import_node_path4.resolve)(target);
  withParentNoFollowSync(path, (parentFd) => refuseDatabaseFileLinks(parentFd, path));
}
function checkSandDataDatabaseSync(db, target) {
  const path = (0, import_node_path4.resolve)(target);
  withParentNoFollowSync(path, (parentFd) => {
    refuseDatabaseFileLinks(parentFd, path);
    if (process.platform === "linux" && mainDatabaseFile(db) !== (0, import_node_path4.join)((0, import_node_fs4.readlinkSync)(descriptorDirOf(parentFd, path)), (0, import_node_path4.basename)(path))) {
      throw new NoFollowRefusedError("parent-link", path);
    }
  });
}
function mainDatabaseFile(db) {
  const file = db.prepare("PRAGMA database_list").get()?.file;
  return typeof file === "string" ? file : void 0;
}
function withParentNoFollowSync(path, use) {
  const root = sandDataNoFollowRoot(path);
  let fd = openDirectoryRefusingLinks(root, root, { followRoot: true });
  let current = root;
  try {
    for (const name of segmentsBelow(root, (0, import_node_path4.dirname)(path))) {
      current = (0, import_node_path4.join)(current, name);
      const child = openDirectoryRefusingLinks(entryOf(fd, current, name), current);
      (0, import_node_fs4.closeSync)(fd);
      fd = child;
    }
    return use(fd);
  } finally {
    (0, import_node_fs4.closeSync)(fd);
  }
}
function openDirectoryRefusingLinks(entry, shown, options = {}) {
  try {
    return (0, import_node_fs4.openSync)(
      entry,
      options.followRoot === true ? import_node_fs4.constants.O_RDONLY | import_node_fs4.constants.O_DIRECTORY : NO_FOLLOW_DIRECTORY2
    );
  } catch (error) {
    const errno = findSystemErrno(error);
    const link = errno === "ELOOP" || errno === "ENOTDIR" && (0, import_node_fs4.lstatSync)(entry, { throwIfNoEntry: false })?.isSymbolicLink() === true;
    throw link ? new NoFollowRefusedError("parent-link", shown) : error;
  }
}
function entryOf(fd, logicalPath, name) {
  return process.platform === "linux" ? `/proc/self/fd/${fd}/${name}` : logicalPath;
}
function descriptorDirOf(fd, logicalPath) {
  return process.platform === "linux" ? `/proc/self/fd/${fd}` : logicalPath;
}
function sandDataNoFollowRoot(path) {
  return isSandDataPath(path) ? sandDataAnchor() : hostNoFollowRoot((0, import_node_path4.dirname)(path));
}
function isSandDataPath(path) {
  return isPathWithin(sandDataAnchor(), path);
}
function sandDataAnchor() {
  return (0, import_node_path4.dirname)(getSandRootDir());
}

// ../packages/grok-bot-harness/src/host-diagnostics.ts
var pinnedReporter = null;
function reportHostDiagnostic(diagnostic) {
  pinnedReporter?.(diagnostic);
}

// src/host/fallback-diagnostics.ts
var reported = /* @__PURE__ */ new Set();
function reportFallback(stage, error) {
  const errorClass = errorLogTag(error);
  const key = `${stage} ${errorClass}`;
  if (reported.has(key)) return;
  reported.add(key);
  reportHostDiagnostic({ kind: "fallback_taken", stage, errorClass });
}

// src/host/storage/store-db.ts
var import_node_sqlite = require("node:sqlite");
var DB_BUSY_TIMEOUT_MS = 5e3;
function applyStorePragmas(db, options) {
  db.exec(`PRAGMA busy_timeout = ${options.busyTimeoutMs ?? DB_BUSY_TIMEOUT_MS}`);
  try {
    db.exec("PRAGMA journal_mode = WAL");
    db.exec("PRAGMA synchronous = NORMAL");
  } catch (error) {
    if (options.allowWalFailure !== true) throw error;
  }
  if (options.incrementalAutoVacuum === true) {
    db.exec("PRAGMA auto_vacuum = INCREMENTAL");
  }
}

// src/host/extensions/content-search/agent-content-search.ts
var AGENT_CONTENT_SEARCH_MAX_MATCHES_PER_AGENT = 5;
var SNIPPET_LEAD = 30;
var SNIPPET_TRAIL = 60;
var ELLIPSIS = "\u2026";
function entrySearchText(entry) {
  switch (entry.kind) {
    case "message":
      return entry.content;
    case "send-message":
      return entry.message.type === "text" ? entry.message.content : "";
    case "notice":
      return entry.text;
    default:
      return "";
  }
}
function flatten(text) {
  return text.replace(/\s+/g, " ").trim();
}
function buildContentSnippet(text, normalizedQuery) {
  if (normalizedQuery.length === 0) return null;
  const flat = flatten(text);
  const index = flat.toLowerCase().indexOf(normalizedQuery);
  if (index < 0) return null;
  const start = Math.max(0, index - SNIPPET_LEAD);
  const end = Math.min(flat.length, index + normalizedQuery.length + SNIPPET_TRAIL);
  const core = flat.slice(start, end);
  const prefix = start > 0 ? ELLIPSIS : "";
  const suffix = end < flat.length ? ELLIPSIS : "";
  return `${prefix}${core}${suffix}`;
}

// src/host/extensions/content-search/search-index-db.ts
var INDEXED_BODY_MAX_CHARS = 2e4;
var FTS_QUERY_MAX_TERMS = 8;
var SNIPPET_CONTEXT_TOKENS = 16;
var META_RECONCILE_DONE = "reconcile_done";
var ATTACHMENT_KINDS = new Set(SAND_ATTACHMENT_KINDS);
function parseAttachmentKind(kind) {
  return ATTACHMENT_KINDS.has(kind) ? kind : "file";
}
var CORE_SCHEMA = `
CREATE TABLE IF NOT EXISTS meta (
  key TEXT PRIMARY KEY,
  value TEXT NOT NULL
) STRICT;
CREATE TABLE IF NOT EXISTS agents (
  agent_id TEXT PRIMARY KEY,
  fingerprint TEXT NOT NULL
) STRICT;
CREATE TABLE IF NOT EXISTS messages (
  id INTEGER PRIMARY KEY,
  agent_id TEXT NOT NULL,
  entry_id TEXT NOT NULL,
  role TEXT NOT NULL CHECK (role IN ('user', 'assistant')),
  timestamp_ms INTEGER NOT NULL,
  body TEXT NOT NULL,
  seq INTEGER,
  in_main_chat INTEGER NOT NULL CHECK (in_main_chat IN (0, 1)),
  UNIQUE(agent_id, entry_id)
) STRICT;
CREATE INDEX IF NOT EXISTS messages_agent_recency
  ON messages(agent_id, timestamp_ms DESC);
CREATE INDEX IF NOT EXISTS messages_agent_seq ON messages(agent_id, seq);
CREATE TABLE IF NOT EXISTS media (
  id INTEGER PRIMARY KEY,
  agent_id TEXT NOT NULL,
  entry_id TEXT NOT NULL,
  file_name TEXT NOT NULL,
  ext TEXT NOT NULL,
  mime TEXT,
  kind TEXT NOT NULL,
  timestamp_ms INTEGER NOT NULL,
  width INTEGER,
  height INTEGER,
  seq INTEGER,
  UNIQUE(agent_id, entry_id)
) STRICT;
CREATE INDEX IF NOT EXISTS media_recency ON media(timestamp_ms DESC);
CREATE INDEX IF NOT EXISTS media_agent_seq ON media(agent_id, seq);
CREATE TABLE IF NOT EXISTS server_cursors (
  agent_id TEXT PRIMARY KEY,
  generation INTEGER NOT NULL,
  topped_through_seq INTEGER NOT NULL,
  backfilled_from_seq INTEGER NOT NULL
) STRICT;
`;
var FTS_SCHEMA = `
CREATE VIRTUAL TABLE IF NOT EXISTS messages_fts USING fts5(
  body,
  content='messages',
  content_rowid='id',
  tokenize='unicode61 remove_diacritics 2',
  prefix='2 3'
);
CREATE TRIGGER IF NOT EXISTS messages_fts_insert AFTER INSERT ON messages BEGIN
  INSERT INTO messages_fts(rowid, body) VALUES (new.id, new.body);
END;
CREATE TRIGGER IF NOT EXISTS messages_fts_delete AFTER DELETE ON messages BEGIN
  INSERT INTO messages_fts(messages_fts, rowid, body)
    VALUES ('delete', old.id, old.body);
END;
CREATE TRIGGER IF NOT EXISTS messages_fts_update AFTER UPDATE ON messages BEGIN
  INSERT INTO messages_fts(messages_fts, rowid, body)
    VALUES ('delete', old.id, old.body);
  INSERT INTO messages_fts(rowid, body) VALUES (new.id, new.body);
END;
CREATE VIRTUAL TABLE IF NOT EXISTS media_fts USING fts5(
  file_name,
  content='media',
  content_rowid='id',
  tokenize='unicode61 remove_diacritics 2',
  prefix='2 3'
);
CREATE TRIGGER IF NOT EXISTS media_fts_insert AFTER INSERT ON media BEGIN
  INSERT INTO media_fts(rowid, file_name) VALUES (new.id, new.file_name);
END;
CREATE TRIGGER IF NOT EXISTS media_fts_delete AFTER DELETE ON media BEGIN
  INSERT INTO media_fts(media_fts, rowid, file_name)
    VALUES ('delete', old.id, old.file_name);
END;
CREATE TRIGGER IF NOT EXISTS media_fts_update AFTER UPDATE ON media BEGIN
  INSERT INTO media_fts(media_fts, rowid, file_name)
    VALUES ('delete', old.id, old.file_name);
  INSERT INTO media_fts(rowid, file_name) VALUES (new.id, new.file_name);
END;
`;
function isFts5Available() {
  let db;
  try {
    db = new import_node_sqlite2.DatabaseSync(":memory:");
    db.exec("CREATE VIRTUAL TABLE fts5_probe USING fts5(x)");
    return true;
  } catch (error) {
    reportFallback("search_index_db", error);
    return false;
  } finally {
    db?.close();
  }
}
function openSearchIndexDb(dbPath) {
  refuseSandDataDatabaseLinksSync(dbPath);
  const db = new import_node_sqlite2.DatabaseSync(dbPath);
  try {
    checkSandDataDatabaseSync(db, dbPath);
    applyStorePragmas(db, { incrementalAutoVacuum: true });
    return db;
  } catch (error) {
    try {
      db.close();
    } catch {
    }
    throw error;
  }
}
function openSearchIndexReadDb(dbPath) {
  const db = openSearchIndexDb(dbPath);
  db.exec("PRAGMA query_only = ON");
  return db;
}
function boundRowLimit(limit) {
  const rowLimit = Math.min(Math.floor(limit), Number.MAX_SAFE_INTEGER);
  return rowLimit > 0 ? rowLimit : 0;
}
function isSearchIndexUnreadable(error) {
  return sqliteResultCode(error) !== void 0 && !isSqliteBusyError(error) && !isSqliteInputOrMemoryError(error);
}
function ensureSearchIndexSchema(db, isFtsEnabled) {
  db.exec(CORE_SCHEMA);
  if (isFtsEnabled) db.exec(FTS_SCHEMA);
}
function writeReconcileDone(db) {
  db.prepare(
    "INSERT INTO meta (key, value) VALUES (?, '1') ON CONFLICT(key) DO UPDATE SET value = '1'"
  ).run(META_RECONCILE_DONE);
}
function readServerCursor(db, agentId) {
  const row = db.prepare(
    `SELECT generation, topped_through_seq AS toppedThroughSeq,
				backfilled_from_seq AS backfilledFromSeq
			 FROM server_cursors WHERE agent_id = ?`
  ).get(agentId);
  if (typeof row?.generation !== "number" || typeof row.toppedThroughSeq !== "number" || typeof row.backfilledFromSeq !== "number") {
    return null;
  }
  return {
    generation: row.generation,
    toppedThroughSeq: row.toppedThroughSeq,
    backfilledFromSeq: row.backfilledFromSeq
  };
}
function nextServerCursor(current, range) {
  const isTop = range.beforeSeq == null;
  const restarted = {
    generation: range.generation,
    toppedThroughSeq: range.beforeSeq == null ? range.rows.reduce((max, row) => Math.max(max, row.seq), range.fromSeq - 1) : range.beforeSeq - 1,
    backfilledFromSeq: range.fromSeq
  };
  if (current == null || range.generation > current.generation) {
    return isTop ? { kind: "reset", next: restarted } : { kind: "superseded" };
  }
  if (range.generation < current.generation) return { kind: "superseded" };
  const isAdjacent = range.fromSeq <= current.toppedThroughSeq + 1 && (range.beforeSeq == null || range.beforeSeq >= current.backfilledFromSeq);
  if (!isAdjacent) return isTop ? { kind: "advance", next: restarted } : { kind: "superseded" };
  const next = {
    generation: current.generation,
    toppedThroughSeq: Math.max(current.toppedThroughSeq, restarted.toppedThroughSeq),
    backfilledFromSeq: Math.min(current.backfilledFromSeq, range.fromSeq)
  };
  return next.toppedThroughSeq === current.toppedThroughSeq && next.backfilledFromSeq === current.backfilledFromSeq ? { kind: "verify" } : { kind: "advance", next };
}
function deriveMessageRow(entry) {
  if (isHiddenOutboundAgentPeerMessageEntry(entry)) return null;
  const body = entrySearchText(entry).trim();
  if (body.length === 0) return null;
  return {
    entryId: entry.id,
    role: entry.kind === "message" ? entry.role : "assistant",
    timestampMs: wholeMs(entry.timestampMs),
    body: body.slice(0, INDEXED_BODY_MAX_CHARS),
    isInMainChat: !isBranchedEntry(entry) && !isTimelineHiddenEntry(entry)
  };
}
function wholeMs(value) {
  return typeof value === "number" && Number.isFinite(value) ? Math.round(value) : 0;
}
function wholeDimension(value) {
  return typeof value === "number" && Number.isFinite(value) ? Math.round(value) : null;
}
function mediaFileName(fileName, urlOrPath) {
  const trimmed = fileName?.trim();
  if (trimmed != null && trimmed.length > 0) return trimmed;
  let subject = urlOrPath;
  try {
    subject = new URL(urlOrPath).pathname;
  } catch {
  }
  try {
    subject = decodeURIComponent(subject);
  } catch {
  }
  return (0, import_node_path5.basename)(subject);
}
function deriveMediaRow(entry) {
  let fileName;
  let urlOrPath;
  let width = null;
  let height = null;
  if (entry.kind === "user-attachment") {
    urlOrPath = entry.file_path;
    fileName = mediaFileName(entry.file_name, urlOrPath);
    width = wholeDimension(entry.width);
    height = wholeDimension(entry.height);
  } else if (entry.kind === "send-message" && entry.message.type === "attachment") {
    urlOrPath = entry.message.url;
    fileName = mediaFileName(entry.message.file_name, urlOrPath);
  } else {
    return null;
  }
  if (fileName.length === 0) return null;
  const ext = (0, import_node_path5.extname)(fileName).toLowerCase();
  const mime = imageMimeFromPath(fileName) ?? videoMimeFromPath(fileName) ?? audioMimeFromPath(fileName) ?? null;
  return {
    entryId: entry.id,
    fileName,
    ext,
    mime,
    kind: classifyAttachment({ fileName, urlOrPath }),
    timestampMs: wholeMs(entry.timestampMs),
    width,
    height
  };
}
function searchTerms(query) {
  return query.normalize("NFKC").replaceAll("\0", " ").trim().split(/\s+/).filter((term) => term.length > 0).slice(0, FTS_QUERY_MAX_TERMS);
}
function buildFtsMatchQuery(query) {
  const terms = searchTerms(query);
  if (terms.length === 0) return null;
  return terms.map((term) => `"${term.replaceAll('"', '""')}"*`).join(" ");
}
function termConjunction(column, terms) {
  return terms.map(() => `instr(lower(${column}), lower(?)) > 0`).join(" AND ");
}
function flattenSnippet(text) {
  return text.replace(/\s+/g, " ").trim();
}
var EFFECTIVE_TIMESTAMP = `CASE
	WHEN m.timestamp_ms > 0 THEN m.timestamp_ms
	ELSE COALESCE(
		(SELECT MAX(m2.timestamp_ms) FROM messages m2 WHERE m2.agent_id = m.agent_id),
		0
	)
END`;
function messageMatchSql(rowSource, snippetColumn) {
  return `SELECT
			m.agent_id AS agentId,
			m.entry_id AS entryId,
			m.role AS role,
			matched.ts AS timestampMs,
			${snippetColumn} AS snippet
		 FROM (
			SELECT match_rowid, ts FROM (
				SELECT
					match_rowid,
					ts,
					ROW_NUMBER() OVER (
						PARTITION BY agent_id
						ORDER BY ts DESC
					) AS agent_rank
				FROM (${rowSource})
			)
			WHERE agent_rank <= ${AGENT_CONTENT_SEARCH_MAX_MATCHES_PER_AGENT}
			ORDER BY ts DESC
			LIMIT ?
		 ) AS matched
		 JOIN messages m ON m.id = matched.match_rowid`;
}
function collectMessageMatches(rows, snippetOf) {
  const results = [];
  for (const row of rows) {
    const agentId = parseAgentId(row.agentId);
    if (agentId == null || typeof row.entryId !== "string" || row.role !== "user" && row.role !== "assistant" || typeof row.timestampMs !== "number") {
      continue;
    }
    const snippet = snippetOf(row);
    if (snippet == null) continue;
    results.push({
      agentId,
      entryId: row.entryId,
      role: row.role,
      timestampMs: row.timestampMs,
      snippet
    });
  }
  return results;
}
function searchMessages(db, query, limit, isFtsEnabled) {
  const rowLimit = boundRowLimit(limit);
  if (rowLimit <= 0) return [];
  if (!isFtsEnabled) return searchMessagesPlain(db, query, rowLimit);
  const match = buildFtsMatchQuery(query);
  if (match == null) return [];
  const rows = db.prepare(
    `${messageMatchSql(
      `SELECT
					messages_fts.rowid AS match_rowid,
					m.agent_id AS agent_id,
					${EFFECTIVE_TIMESTAMP} AS ts
				FROM messages_fts
				JOIN messages m ON m.id = messages_fts.rowid
				WHERE messages_fts MATCH ?`,
      `snippet(messages_fts, 0, '', '', '\u2026', ${SNIPPET_CONTEXT_TOKENS})`
    )}
			 JOIN messages_fts ON messages_fts.rowid = matched.match_rowid
			 WHERE messages_fts MATCH ?
			 ORDER BY matched.ts DESC`
  ).all(match, rowLimit, match);
  return collectMessageMatches(
    rows,
    (row) => typeof row.snippet === "string" ? flattenSnippet(row.snippet) : null
  );
}
function searchMessagesPlain(db, query, limit) {
  const terms = searchTerms(query);
  if (terms.length === 0) return [];
  const rows = db.prepare(
    `${messageMatchSql(
      `SELECT
					m.id AS match_rowid,
					m.agent_id AS agent_id,
					${EFFECTIVE_TIMESTAMP} AS ts
				FROM messages m
				WHERE ${termConjunction("m.body", terms)}`,
      "m.body"
    )}
			 ORDER BY matched.ts DESC`
  ).all(...terms, limit);
  return collectMessageMatches(
    rows,
    (row) => typeof row.snippet === "string" ? plainSnippet(row.snippet, terms) : null
  );
}
function plainSnippet(body, terms) {
  for (const term of terms) {
    const snippet = buildContentSnippet(body, term.toLowerCase());
    if (snippet != null) return snippet;
  }
  const flat = flattenSnippet(body);
  const head = flat.slice(0, SNIPPET_CONTEXT_TOKENS * 8);
  return head.length < flat.length ? `${head}\u2026` : head;
}
function searchConversation(db, query, isFtsEnabled) {
  db.exec("BEGIN");
  try {
    return {
      matches: conversationMatches(db, query, isFtsEnabled),
      isComplete: isMirroredBelow(readServerCursor(db, query.agentId), query.beforeSeq)
    };
  } finally {
    db.exec("COMMIT");
  }
}
function isMirroredBelow(cursor, beforeSeq) {
  if (cursor === null || cursor.backfilledFromSeq !== 0) return false;
  return beforeSeq === void 0 || cursor.toppedThroughSeq >= beforeSeq - 1;
}
function conversationMatches(db, { agentId, query, beforeSeq, limit }, isFtsEnabled) {
  if (limit <= 0) return [];
  const terms = searchTerms(query);
  if (terms.length === 0) return [];
  const seqBound = beforeSeq === void 0 ? "" : "AND m.seq < ?";
  const seqArgs = beforeSeq === void 0 ? [] : [beforeSeq];
  if (!isFtsEnabled) {
    const rows2 = db.prepare(
      `SELECT
					m.entry_id AS entryId,
					m.seq AS seq,
					m.role AS role,
					m.timestamp_ms AS timestampMs,
					m.body AS snippet
				 FROM messages m
				 WHERE m.agent_id = ? AND m.seq IS NOT NULL AND m.in_main_chat = 1 ${seqBound}
					AND ${termConjunction("m.body", terms)}
				 ORDER BY m.seq DESC
				 LIMIT ?`
    ).all(agentId, ...seqArgs, ...terms, limit);
    return collectConversationMatches(rows2, (snippet) => plainSnippet(snippet, terms));
  }
  const match = buildFtsMatchQuery(query);
  if (match == null) return [];
  const rows = db.prepare(
    `SELECT
				m.entry_id AS entryId,
				m.seq AS seq,
				m.role AS role,
				m.timestamp_ms AS timestampMs,
				snippet(messages_fts, 0, '', '', '\u2026', ${SNIPPET_CONTEXT_TOKENS}) AS snippet
			 FROM messages_fts
			 JOIN messages m ON m.id = messages_fts.rowid
			 WHERE messages_fts MATCH ? AND m.agent_id = ? AND m.seq IS NOT NULL
				AND m.in_main_chat = 1 ${seqBound}
			 ORDER BY m.seq DESC
			 LIMIT ?`
  ).all(match, agentId, ...seqArgs, limit);
  return collectConversationMatches(rows, flattenSnippet);
}
function collectConversationMatches(rows, snippetOf) {
  const results = [];
  for (const row of rows) {
    if (typeof row.entryId !== "string" || typeof row.seq !== "number" || row.role !== "user" && row.role !== "assistant" || typeof row.timestampMs !== "number" || typeof row.snippet !== "string") {
      continue;
    }
    results.push({
      entryId: row.entryId,
      seq: row.seq,
      role: row.role,
      timestampMs: row.timestampMs,
      snippet: snippetOf(row.snippet)
    });
  }
  return results;
}
var MEDIA_SELECT_COLUMNS = `
	md.agent_id AS agentId,
	md.entry_id AS entryId,
	md.file_name AS fileName,
	md.ext AS ext,
	md.mime AS mime,
	md.kind AS kind,
	md.timestamp_ms AS timestampMs,
	md.width AS width,
	md.height AS height`;
function mediaMatchRows(db, query, limit, isFtsEnabled) {
  if (isFtsEnabled) {
    const match = buildFtsMatchQuery(query);
    if (match == null) return browseMediaRows(db, limit);
    return db.prepare(
      `SELECT ${MEDIA_SELECT_COLUMNS}
				 FROM media_fts
				 JOIN media md ON md.id = media_fts.rowid
				 WHERE media_fts MATCH ?
				 ORDER BY md.timestamp_ms DESC
				 LIMIT ?`
    ).all(match, limit);
  }
  const terms = searchTerms(query);
  if (terms.length === 0) return browseMediaRows(db, limit);
  return db.prepare(
    `SELECT ${MEDIA_SELECT_COLUMNS}
			 FROM media md
			 WHERE ${termConjunction("md.file_name", terms)}
			 ORDER BY md.timestamp_ms DESC
			 LIMIT ?`
  ).all(...terms, limit);
}
function browseMediaRows(db, limit) {
  return db.prepare(
    `SELECT ${MEDIA_SELECT_COLUMNS}
			 FROM media md
			 ORDER BY md.timestamp_ms DESC
			 LIMIT ?`
  ).all(limit);
}
function searchMedia(db, query, limit, isFtsEnabled) {
  const rowLimit = boundRowLimit(limit);
  if (rowLimit <= 0) return [];
  const rows = mediaMatchRows(db, query, rowLimit, isFtsEnabled);
  const results = [];
  for (const row of rows) {
    const agentId = parseAgentId(row.agentId);
    if (agentId == null || typeof row.entryId !== "string" || typeof row.fileName !== "string" || typeof row.ext !== "string" || typeof row.kind !== "string" || typeof row.timestampMs !== "number") {
      continue;
    }
    results.push({
      agentId,
      entryId: row.entryId,
      fileName: row.fileName,
      ext: row.ext,
      mime: typeof row.mime === "string" ? row.mime : null,
      kind: parseAttachmentKind(row.kind),
      timestampMs: row.timestampMs,
      width: typeof row.width === "number" ? row.width : null,
      height: typeof row.height === "number" ? row.height : null
    });
  }
  return results;
}

// src/host/extensions/content-search/search-index-writer.ts
var import_node_fs5 = require("node:fs");
var import_node_path6 = require("node:path");
var import_node_sqlite3 = require("node:sqlite");
var STORE_FILENAME = "store.db";
var INCREMENTAL_VACUUM_PAGES = 512;
function prepareStatements(db) {
  return {
    upsertMessage: db.prepare(
      `INSERT INTO messages (agent_id, entry_id, role, timestamp_ms, body, seq, in_main_chat)
			 VALUES (?, ?, ?, ?, ?, ?, ?)
			 ON CONFLICT(agent_id, entry_id) DO UPDATE SET
				role = excluded.role,
				timestamp_ms = excluded.timestamp_ms,
				body = excluded.body,
				seq = excluded.seq,
				in_main_chat = excluded.in_main_chat`
    ),
    deleteMessage: db.prepare("DELETE FROM messages WHERE agent_id = ? AND entry_id = ?"),
    deleteAgentMessages: db.prepare("DELETE FROM messages WHERE agent_id = ?"),
    deleteMessageSpan: db.prepare(
      "DELETE FROM messages WHERE agent_id = ? AND seq >= ? AND seq < ?"
    ),
    upsertMedia: db.prepare(
      `INSERT INTO media (
				agent_id, entry_id, file_name, ext, mime, kind,
				timestamp_ms, width, height, seq
			 )
			 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
			 ON CONFLICT(agent_id, entry_id) DO UPDATE SET
				file_name = excluded.file_name,
				ext = excluded.ext,
				mime = excluded.mime,
				kind = excluded.kind,
				timestamp_ms = excluded.timestamp_ms,
				width = excluded.width,
				height = excluded.height,
				seq = excluded.seq`
    ),
    deleteMedia: db.prepare("DELETE FROM media WHERE agent_id = ? AND entry_id = ?"),
    deleteAgentMedia: db.prepare("DELETE FROM media WHERE agent_id = ?"),
    deleteMediaSpan: db.prepare("DELETE FROM media WHERE agent_id = ? AND seq >= ? AND seq < ?"),
    readIsServerOwned: db.prepare("SELECT 1 AS owned FROM server_cursors WHERE agent_id = ?"),
    upsertCursor: db.prepare(
      `INSERT INTO server_cursors (agent_id, generation, topped_through_seq, backfilled_from_seq)
			 VALUES (?, ?, ?, ?)
			 ON CONFLICT(agent_id) DO UPDATE SET
				generation = excluded.generation,
				topped_through_seq = excluded.topped_through_seq,
				backfilled_from_seq = excluded.backfilled_from_seq`
    ),
    deleteCursor: db.prepare("DELETE FROM server_cursors WHERE agent_id = ?"),
    upsertFingerprint: db.prepare(
      `INSERT INTO agents (agent_id, fingerprint) VALUES (?, ?)
			 ON CONFLICT(agent_id) DO UPDATE SET fingerprint = excluded.fingerprint`
    ),
    deleteFingerprint: db.prepare("DELETE FROM agents WHERE agent_id = ?"),
    readFingerprint: db.prepare("SELECT fingerprint FROM agents WHERE agent_id = ?"),
    listIndexedAgentIds: db.prepare(
      `SELECT agent_id AS agentId FROM agents
			 UNION SELECT DISTINCT agent_id FROM messages
			 UNION SELECT DISTINCT agent_id FROM media
			 UNION SELECT agent_id FROM server_cursors`
    )
  };
}
var SandSearchIndexWriter = class {
  constructor(db, agentsRootDir) {
    this.db = db;
    this.agentsRootDir = agentsRootDir;
    this.statements = prepareStatements(db);
  }
  db;
  agentsRootDir;
  statements;
  storeConnections = /* @__PURE__ */ new Map();
  close() {
    for (const connection of this.storeConnections.values()) {
      try {
        connection.db.close();
      } catch {
      }
    }
    this.storeConnections.clear();
  }
  runJob(job) {
    switch (job.kind) {
      case "upsert-entries":
        if (!this.isServerOwned(job.agentId)) this.upsertEntries(job.agentId, job.entries);
        return void 0;
      case "delete-entry":
        if (!this.isServerOwned(job.agentId)) this.deleteEntry(job.agentId, job.entryId);
        return void 0;
      case "clear-agent":
        if (!this.isServerOwned(job.agentId)) this.clearAgent(job.agentId);
        return void 0;
      case "remove-agent":
        this.clearAgent(job.agentId);
        return void 0;
      case "reindex-agents":
        for (const agentId of job.agentIds) this.reindexAgent(agentId);
        return void 0;
      case "reconcile":
        this.reconcile();
        return void 0;
      case "apply-server-range":
        return this.applyServerRange(job.range);
      default: {
        const _exhaustive = job;
        return _exhaustive;
      }
    }
  }
  isServerOwned(agentId) {
    return this.statements.readIsServerOwned.get(agentId) != null;
  }
  applyServerRange(range) {
    const agentId = parseAgentId(range.agentId);
    if (agentId == null || !(0, import_node_fs5.existsSync)((0, import_node_path6.join)(this.agentsRootDir, agentId))) {
      return "unknown-agent";
    }
    return this.inTransaction(() => {
      const step = nextServerCursor(readServerCursor(this.db, agentId), range);
      switch (step.kind) {
        case "superseded":
          return "superseded";
        case "reset":
          this.statements.deleteAgentMessages.run(agentId);
          this.statements.deleteAgentMedia.run(agentId);
          this.statements.deleteFingerprint.run(agentId);
          break;
        case "advance":
        case "verify":
          break;
        default: {
          const _exhaustive = step;
          return _exhaustive;
        }
      }
      const beforeSeq = range.beforeSeq ?? Number.MAX_SAFE_INTEGER;
      this.statements.deleteMessageSpan.run(agentId, range.fromSeq, beforeSeq);
      this.statements.deleteMediaSpan.run(agentId, range.fromSeq, beforeSeq);
      for (const row of range.rows) this.applyEntry(agentId, row, row.seq);
      if (step.kind !== "verify") {
        this.statements.upsertCursor.run(
          agentId,
          step.next.generation,
          step.next.toppedThroughSeq,
          step.next.backfilledFromSeq
        );
      }
      return "applied";
    });
  }
  storeDbPath(agentId) {
    return (0, import_node_path6.join)(this.agentsRootDir, agentId, STORE_FILENAME);
  }
  evictStoreConnection(agentId) {
    const cached = this.storeConnections.get(agentId);
    if (cached == null) return;
    this.storeConnections.delete(agentId);
    try {
      cached.db.close();
    } catch {
    }
  }
  storeConnection(agentId) {
    const cached = this.storeConnections.get(agentId);
    if (cached != null) return cached;
    const path = this.storeDbPath(agentId);
    if (!(0, import_node_fs5.existsSync)(path)) return null;
    try {
      refuseSandDataDatabaseLinksSync(path);
      const db = new import_node_sqlite3.DatabaseSync(path, { readOnly: true });
      try {
        checkSandDataDatabaseSync(db, path);
        db.exec(`PRAGMA busy_timeout = ${DB_BUSY_TIMEOUT_MS}`);
      } catch (error) {
        db.close();
        throw error;
      }
      const connection = { db };
      this.storeConnections.set(agentId, connection);
      return connection;
    } catch {
      return null;
    }
  }
  readStoreFingerprint(agentId) {
    const connection = this.storeConnection(agentId);
    if (connection == null) return null;
    try {
      const row = connection.db.prepare(
        "SELECT COUNT(*) AS count, COALESCE(MAX(seq), 0) AS maxSeq FROM transcript_entries"
      ).get();
      if (row == null || typeof row.count !== "number" || typeof row.maxSeq !== "number") {
        return null;
      }
      return `${row.count}:${row.maxSeq}`;
    } catch {
      this.evictStoreConnection(agentId);
      return null;
    }
  }
  inTransaction(operation) {
    this.db.exec("BEGIN IMMEDIATE");
    try {
      const result = operation();
      this.db.exec("COMMIT");
      return result;
    } catch (error) {
      try {
        this.db.exec("ROLLBACK");
      } catch {
      }
      throw error;
    }
  }
  applyEntry(agentId, entry, seq) {
    let message = null;
    let media = null;
    try {
      message = deriveMessageRow(entry);
      media = deriveMediaRow(entry);
    } catch {
    }
    if (message != null) {
      this.statements.upsertMessage.run(
        agentId,
        message.entryId,
        message.role,
        message.timestampMs,
        message.body,
        seq,
        message.isInMainChat ? 1 : 0
      );
    } else {
      this.statements.deleteMessage.run(agentId, entry.id);
    }
    if (media != null) {
      this.statements.upsertMedia.run(
        agentId,
        media.entryId,
        media.fileName,
        media.ext,
        media.mime,
        media.kind,
        media.timestampMs,
        media.width,
        media.height,
        seq
      );
    } else {
      this.statements.deleteMedia.run(agentId, entry.id);
    }
  }
  refreshFingerprint(agentId) {
    const fingerprint = this.readStoreFingerprint(agentId);
    if (fingerprint == null) {
      this.statements.deleteFingerprint.run(agentId);
    } else {
      this.statements.upsertFingerprint.run(agentId, fingerprint);
    }
  }
  upsertEntries(agentId, entries) {
    if (entries.length === 0) return;
    this.inTransaction(() => {
      for (const entry of entries) this.applyEntry(agentId, entry, null);
      this.refreshFingerprint(agentId);
    });
  }
  deleteEntry(agentId, entryId) {
    this.inTransaction(() => {
      this.statements.deleteMessage.run(agentId, entryId);
      this.statements.deleteMedia.run(agentId, entryId);
      this.refreshFingerprint(agentId);
    });
  }
  clearAgent(agentId) {
    this.evictStoreConnection(agentId);
    this.inTransaction(() => {
      this.statements.deleteAgentMessages.run(agentId);
      this.statements.deleteAgentMedia.run(agentId);
      this.statements.deleteFingerprint.run(agentId);
      this.statements.deleteCursor.run(agentId);
    });
    this.db.exec(`PRAGMA incremental_vacuum(${INCREMENTAL_VACUUM_PAGES})`);
  }
  reindexAgent(agentId) {
    if (this.isServerOwned(agentId)) return;
    this.evictStoreConnection(agentId);
    if (!(0, import_node_fs5.existsSync)(this.storeDbPath(agentId))) {
      this.clearAgent(agentId);
      return;
    }
    const connection = this.storeConnection(agentId);
    if (connection == null) return;
    let rows;
    try {
      rows = connection.db.prepare("SELECT seq, entry FROM transcript_entries ORDER BY seq").all();
    } catch {
      this.evictStoreConnection(agentId);
      return;
    }
    let maxSeq = 0;
    const entries = [];
    for (const row of rows) {
      if (typeof row.seq === "number" && row.seq > maxSeq) maxSeq = row.seq;
      if (typeof row.entry !== "string") continue;
      try {
        const parsed = JSON.parse(row.entry);
        if (parsed != null && typeof parsed === "object" && typeof parsed.id === "string" && typeof parsed.kind === "string") {
          entries.push(parsed);
        }
      } catch {
      }
    }
    const fingerprint = `${rows.length}:${maxSeq}`;
    this.inTransaction(() => {
      this.statements.deleteAgentMessages.run(agentId);
      this.statements.deleteAgentMedia.run(agentId);
      for (const entry of entries) this.applyEntry(agentId, entry, null);
      this.statements.upsertFingerprint.run(agentId, fingerprint);
    });
    this.db.exec(`PRAGMA incremental_vacuum(${INCREMENTAL_VACUUM_PAGES})`);
  }
  reconcile() {
    let agentDirs;
    try {
      agentDirs = (0, import_node_fs5.readdirSync)(this.agentsRootDir, { withFileTypes: true }).flatMap((dirent) => {
        const agentId = dirent.isDirectory() ? parseAgentId(dirent.name) : null;
        return agentId == null ? [] : [agentId];
      });
    } catch {
      agentDirs = [];
    }
    const onDisk = new Set(agentDirs);
    const indexedRows = this.statements.listIndexedAgentIds.all();
    for (const row of indexedRows) {
      const agentId = parseAgentId(row.agentId);
      if (agentId === null) continue;
      if (!onDisk.has(agentId)) this.clearAgent(agentId);
    }
    for (const agentId of agentDirs) {
      if (!(0, import_node_fs5.existsSync)(this.storeDbPath(agentId))) continue;
      const storeFingerprint = this.readStoreFingerprint(agentId);
      if (storeFingerprint == null) continue;
      const indexed = this.statements.readFingerprint.get(agentId);
      if (indexed?.fingerprint !== storeFingerprint) {
        this.reindexAgent(agentId);
      }
    }
    writeReconcileDone(this.db);
  }
};

// src/host/extensions/content-search/search-index-worker.ts
function serveJobs(port2, config) {
  invariant(
    typeof config.indexDbPath === "string" && typeof config.agentsRootDir === "string",
    "search-index-worker needs indexDbPath + agentsRootDir"
  );
  const db = openSearchIndexDb(config.indexDbPath);
  ensureSearchIndexSchema(db, isFts5Available());
  const writer = new SandSearchIndexWriter(db, config.agentsRootDir);
  port2.on("message", (request) => {
    let response;
    try {
      const outcome = writer.runJob(request.job);
      response = {
        requestId: request.requestId,
        ok: true,
        ...outcome === void 0 ? {} : { outcome }
      };
    } catch (error) {
      response = {
        requestId: request.requestId,
        ok: false,
        message: errorMessage(error),
        isIndexCorrupt: isSqliteCorruptError(error)
      };
    }
    port2.postMessage(response);
  });
}
function runQuery(db, query, isFtsEnabled) {
  switch (query.kind) {
    case "messages":
      return {
        kind: query.kind,
        matches: searchMessages(db, query.query, query.limit, isFtsEnabled)
      };
    case "media":
      return { kind: query.kind, matches: searchMedia(db, query.query, query.limit, isFtsEnabled) };
    case "conversation":
      return { kind: query.kind, ...searchConversation(db, query, isFtsEnabled) };
    default: {
      const exhaustive = query;
      return exhaustive;
    }
  }
}
function serveQueries(port2, config) {
  invariant(
    typeof config.indexDbPath === "string" && typeof config.isFtsEnabled === "boolean",
    "search-index-worker reader needs indexDbPath + isFtsEnabled"
  );
  let db;
  port2.on("message", (request) => {
    let response;
    try {
      db ??= openSearchIndexReadDb(config.indexDbPath);
      response = {
        requestId: request.requestId,
        ok: true,
        ...runQuery(db, request.query, config.isFtsEnabled)
      };
    } catch (error) {
      response = {
        requestId: request.requestId,
        ok: false,
        message: errorMessage(error),
        errorClass: errorLogTag(error),
        isIndexUnreadable: isSearchIndexUnreadable(error)
      };
    }
    port2.postMessage(response);
  });
}
var port = import_node_worker_threads.parentPort;
invariant(port != null, "search-index-worker must run as a worker_thread");
var data = import_node_worker_threads.workerData;
switch (data?.role) {
  case "writer":
    serveJobs(port, data);
    break;
  case "reader":
    serveQueries(port, data);
    break;
  default:
    invariant(false, "search-index-worker needs a writer or reader role");
}
