import {
  readdirSync,
  readFileSync,
  readlinkSync,
  renameSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { request as httpRequest } from "node:http";
import { connect as tlsConnect } from "node:tls";
import { pathToFileURL } from "node:url";

import {
  exitAcceptLanguage,
  isKnownTimeZone,
  isUtcTimeZone,
  parseZoneTab,
  resolveTimeZone,
} from "./sand-locale-match.mjs";

const COUNTRY_CODE_RE = /^[A-Z]{2}$/;
const REGION_CODE_RE = /^[A-Z0-9]{1,3}$/;

function exitGeoRecord({ country, region, timeZone }) {
  return {
    country: typeof country === "string" && COUNTRY_CODE_RE.test(country) ? country : undefined,
    region: typeof region === "string" && REGION_CODE_RE.test(region) ? region : undefined,
    timeZone: isKnownTimeZone(timeZone) ? timeZone : undefined,
  };
}

export const EGRESS_GEO_PROVIDERS = Object.freeze([
  Object.freeze({
    name: "ipinfo",
    host: "ipinfo.io",
    path: "/json",
    exitGeoOf: (body) => exitGeoRecord({ country: body?.country, timeZone: body?.timezone }),
  }),
  Object.freeze({
    name: "ipwho_is",
    host: "ipwho.is",
    path: "/",
    exitGeoOf: (body) =>
      exitGeoRecord(
        body?.success === false
          ? {}
          : {
              country: body?.country_code,
              region: body?.region_code,
              timeZone: body?.timezone?.id,
            },
      ),
  }),
]);

export const EGRESS_GEO_REQUEST_TIMEOUT_MS = 5_000;
const EGRESS_GEO_MAX_BODY_BYTES = 64 * 1024;
const EGRESS_GEO_USER_AGENT = "sand-egress-timezone/1";
const ZONEINFO_DIR = "zoneinfo/";
const EGRESS_ERROR_CLASS_PREFIXES = ["no_provider(", "no_time_zone("];
const PROC_ENTRY_GONE_CODES = new Set(["ENOENT", "ESRCH", "EACCES"]);

export class EgressGeoError extends Error {
  constructor(code, options) {
    super(`egress time zone lookup failed: ${code}`, options);
    this.name = "EgressGeoError";
    this.code = code;
  }
}

export function egressErrorCode(error) {
  if (typeof error?.code === "string") return error.code;
  if (typeof error?.cause?.code === "string") return error.cause.code;
  return error instanceof Error ? error.name : typeof error;
}

export function readBoxTimeZone({ path = "/etc/localtime", readLink = readlinkSync } = {}) {
  let target;
  try {
    target = readLink(path);
  } catch (error) {
    if (error?.code === "ENOENT" || error?.code === "EINVAL") return null;
    throw error;
  }
  const index = target.lastIndexOf(ZONEINFO_DIR);
  return index === -1 ? null : target.slice(index + ZONEINFO_DIR.length);
}

export function egressErrorClass(error) {
  const code = egressErrorCode(error);
  if (code === "browser_not_found" || code === "unsupported_proxy") return code;
  const prefix = EGRESS_ERROR_CLASS_PREFIXES.find((candidate) => code.startsWith(candidate));
  return prefix == null ? "other" : prefix.slice(0, -1);
}

export function isBoxClockUtc({ readLink = readlinkSync } = {}) {
  return isUtcTimeZone(readBoxTimeZone({ readLink }));
}

export function readZoneLinks({
  path = "/usr/share/zoneinfo/tzdata.zi",
  readFile = readFileSync,
} = {}) {
  let text;
  try {
    text = String(readFile(path, "utf8"));
  } catch (error) {
    if (error?.code === "ENOENT") return new Map();
    throw error;
  }
  const links = new Map();
  for (const line of text.split("\n")) {
    const [kind, target, link] = line.split(" ");
    if (kind === "L" && target !== undefined && link !== undefined) links.set(link, target);
  }
  return links;
}

export function readZonesByCountry({
  path = "/usr/share/zoneinfo/zone.tab",
  readFile = readFileSync,
} = {}) {
  try {
    return parseZoneTab(readFile(path, "utf8"));
  } catch (error) {
    if (error?.code === "ENOENT") return new Map();
    throw error;
  }
}

export function chromeBrowserArgsForPort(
  port,
  { procRoot = "/proc", readDir = readdirSync, readFile = readFileSync } = {},
) {
  const debugPortArg = `--remote-debugging-port=${port}`;
  for (const entry of readDir(procRoot)) {
    if (!/^\d+$/.test(entry)) continue;
    let cmdline;
    try {
      cmdline = readFile(`${procRoot}/${entry}/cmdline`, "utf8");
    } catch (error) {
      if (PROC_ENTRY_GONE_CODES.has(error?.code)) continue;
      throw error;
    }
    const args = String(cmdline).split("\0");
    if (!args.includes(debugPortArg) || args.some((arg) => arg.startsWith("--type="))) continue;
    return args;
  }
  return undefined;
}

export function chromeProxyServerForPort(port, deps) {
  const args = chromeBrowserArgsForPort(port, deps);
  if (args === undefined) return undefined;
  const proxyArg = args.findLast((arg) => arg.startsWith("--proxy-server="));
  return proxyArg == null ? null : proxyArg.slice("--proxy-server=".length);
}

export function egressProxyForPort(port, proxyServerForPort = chromeProxyServerForPort) {
  const proxyServer = proxyServerForPort(port);
  if (proxyServer === undefined) throw new EgressGeoError("browser_not_found");
  const route = egressRouteFromProxyServer(proxyServer);
  if (route.kind === "unsupported") throw new EgressGeoError("unsupported_proxy");
  return route.kind === "proxy" ? { host: route.host, port: route.port } : null;
}

export function egressRouteFromProxyServer(proxyServer) {
  const value = proxyServer == null ? "" : proxyServer.trim();
  if (value === "" || value === "direct://") return { kind: "direct" };
  if (/[\s=;,]/.test(value)) return { kind: "unsupported" };
  let url;
  try {
    url = new URL(value.includes("://") ? value : `http://${value}`);
  } catch (error) {
    if (error instanceof TypeError) return { kind: "unsupported" };
    throw error;
  }
  if (
    url.protocol !== "http:" ||
    url.username !== "" ||
    url.password !== "" ||
    (url.pathname !== "" && url.pathname !== "/") ||
    url.search !== "" ||
    url.hash !== ""
  ) {
    return { kind: "unsupported" };
  }
  return {
    kind: "proxy",
    host: url.hostname.replace(/^\[(.*)\]$/, "$1"),
    port: url.port === "" ? 80 : Number(url.port),
  };
}

function openTunnel({ proxy, host, port, timeoutMs }) {
  return new Promise((resolve, reject) => {
    const authority = `${host}:${port}`;
    const request = httpRequest({
      host: proxy.host,
      port: proxy.port,
      method: "CONNECT",
      path: authority,
      headers: { host: authority },
      timeout: timeoutMs,
    });
    request.once("timeout", () => request.destroy(new EgressGeoError("connect_timeout")));
    request.once("connect", (response, socket) => {
      if (response.statusCode === 200) {
        resolve(socket);
        return;
      }
      socket.destroy();
      reject(new EgressGeoError(`connect_${response.statusCode}`));
    });
    request.once("error", reject);
    request.end();
  });
}

export async function requestJson(options) {
  const text = await requestText(options);
  try {
    return JSON.parse(text);
  } catch (error) {
    throw new EgressGeoError("bad_json", { cause: error });
  }
}

export async function requestText({
  host,
  port = 443,
  path,
  proxy = null,
  accept = "application/json",
  userAgent = EGRESS_GEO_USER_AGENT,
  timeoutMs = EGRESS_GEO_REQUEST_TIMEOUT_MS,
  secureConnect = tlsConnect,
}) {
  const tunnel = proxy == null ? null : await openTunnel({ proxy, host, port, timeoutMs });
  tunnel?.once("timeout", () => tunnel.destroy(new EgressGeoError("timeout")));
  return await new Promise((resolve, reject) => {
    const request = httpRequest(
      {
        host,
        port,
        path,
        method: "GET",
        headers: { host, accept, "user-agent": userAgent },
        timeout: timeoutMs,
        createConnection: () =>
          secureConnect(
            tunnel == null
              ? { host, port, servername: host, timeout: timeoutMs }
              : { socket: tunnel, servername: host },
          ),
      },
      (response) => {
        if (response.statusCode !== 200) {
          response.resume();
          reject(new EgressGeoError(`http_${response.statusCode}`));
          return;
        }
        const chunks = [];
        let size = 0;
        response.on("data", (chunk) => {
          size += chunk.length;
          if (size > EGRESS_GEO_MAX_BODY_BYTES) {
            request.destroy(new EgressGeoError("body_too_large"));
            return;
          }
          chunks.push(chunk);
        });
        response.on("end", () => {
          if (size > EGRESS_GEO_MAX_BODY_BYTES) return;
          resolve(Buffer.concat(chunks).toString("utf8"));
        });
        response.on("error", reject);
      },
    );
    request.once("timeout", () => request.destroy(new EgressGeoError("timeout")));
    request.on("error", (error) => {
      tunnel?.destroy();
      reject(error);
    });
    request.end();
  });
}

async function fetchExitGeo({ providers, request, proxy, timeoutMs }) {
  const failures = [];
  let partial = null;
  for (const provider of providers) {
    try {
      const body = await request({
        host: provider.host,
        port: provider.port,
        path: provider.path,
        proxy,
        timeoutMs,
      });
      const exitGeo = provider.exitGeoOf(body);
      if (exitGeo.timeZone != null) return { exitGeo, provider: provider.name };
      if (
        exitGeo.country != null &&
        (partial == null || (partial.exitGeo.region == null && exitGeo.region != null))
      ) {
        partial = { exitGeo, provider: provider.name };
      }
      failures.push(`${provider.name}=no_time_zone`);
    } catch (error) {
      failures.push(`${provider.name}=${egressErrorCode(error)}`);
    }
  }
  if (partial != null) return partial;
  throw new EgressGeoError(`no_provider(${failures.join(",")})`);
}

export async function resolveEgressTimeZone({
  port,
  providers = EGRESS_GEO_PROVIDERS,
  proxyServerForPort = chromeProxyServerForPort,
  request = requestJson,
  timeoutMs = EGRESS_GEO_REQUEST_TIMEOUT_MS,
  boxTimeZone = readBoxTimeZone,
  zonesByCountry = readZonesByCountry,
}) {
  const proxy = egressProxyForPort(port, proxyServerForPort);
  const { exitGeo, provider } = await fetchExitGeo({ providers, request, proxy, timeoutMs });
  const { timeZone, source } = resolveTimeZone({
    exitGeo,
    userTimeZone: boxTimeZone(),
    zonesByCountry: zonesByCountry(),
  });
  if (!isKnownTimeZone(timeZone)) throw new EgressGeoError(`no_time_zone(${provider})`);
  return { timeZone, provider, source };
}

export const EXIT_LOCALE_CACHE_PATH = "/tmp/sand-browser-exit-geo";
export const EXIT_LOCALE_CACHE_TTL_MS = 10 * 60_000;
export const EXIT_LOCALE_REQUEST_TIMEOUT_MS = 1_750;
export const EXIT_GEO_FLAG = "--sand-exit-geo=";
const EXIT_GEO_PROVIDER_NAMES = new Set(EGRESS_GEO_PROVIDERS.map((provider) => provider.name));
const EXIT_GEO_SOURCES = new Set([
  "exit_geo",
  "exit_region",
  "user_in_exit_country",
  "exit_country",
  "exit_country_guess",
]);

export function readExitGeoCache({
  path = EXIT_LOCALE_CACHE_PATH,
  proxyServer,
  now = Date.now,
  readFile = readFileSync,
}) {
  let cached;
  try {
    cached = JSON.parse(String(readFile(path, "utf8")));
  } catch (error) {
    if (typeof error?.code === "string" || error instanceof SyntaxError) return null;
    throw error;
  }
  if (
    cached?.proxyServer !== proxyServer ||
    typeof cached.at !== "number" ||
    now() - cached.at > EXIT_LOCALE_CACHE_TTL_MS ||
    now() < cached.at ||
    !EXIT_GEO_PROVIDER_NAMES.has(cached.provider)
  ) {
    return null;
  }
  const exitGeo = exitGeoRecord(cached.exitGeo ?? {});
  return exitGeo.country == null && exitGeo.timeZone == null
    ? null
    : { exitGeo, provider: cached.provider };
}

export function writeExitGeoCache({
  path = EXIT_LOCALE_CACHE_PATH,
  proxyServer,
  exitGeo,
  provider,
  now = Date.now,
}) {
  const scratch = `${path}.${process.pid}`;
  try {
    writeFileSync(scratch, JSON.stringify({ proxyServer, at: now(), exitGeo, provider }), {
      mode: 0o644,
    });
    renameSync(scratch, path);
  } catch (error) {
    rmSync(scratch, { force: true });
    throw error;
  }
}

export async function resolveExitLocale({
  proxyServer,
  configuredAcceptLanguage,
  providers = EGRESS_GEO_PROVIDERS,
  request = requestJson,
  timeoutMs = EXIT_LOCALE_REQUEST_TIMEOUT_MS,
  boxTimeZone = readBoxTimeZone,
  zonesByCountry = readZonesByCountry,
  readCache = readExitGeoCache,
  writeCache = writeExitGeoCache,
}) {
  const route = egressRouteFromProxyServer(proxyServer);
  if (route.kind === "unsupported") throw new EgressGeoError("unsupported_proxy");
  const proxy = route.kind === "proxy" ? { host: route.host, port: route.port } : null;
  let found = readCache({ proxyServer });
  if (found == null) {
    found = await fetchExitGeo({ providers, request, proxy, timeoutMs });
    try {
      writeCache({ proxyServer, exitGeo: found.exitGeo, provider: found.provider });
    } catch (error) {
      console.error(`exit geo cache write failed: ${egressErrorCode(error)}`);
    }
  }
  const { exitGeo, provider } = found;
  const { timeZone, source } = resolveTimeZone({
    exitGeo,
    userTimeZone: boxTimeZone(),
    zonesByCountry: zonesByCountry(),
  });
  return {
    timeZone: isKnownTimeZone(timeZone) ? timeZone : null,
    source,
    provider,
    acceptLanguage: exitAcceptLanguage({
      country: exitGeo.country,
      configured: configuredAcceptLanguage,
    }),
  };
}

export function exitLocaleLine({ timeZone, source, provider, acceptLanguage }) {
  return `${timeZone ?? "-"}\t${acceptLanguage}\t${provider}\t${source}\n`;
}

export function parseExitGeoFlag(args) {
  const flag = args?.findLast((arg) => arg.startsWith(EXIT_GEO_FLAG));
  if (flag === undefined) return null;
  const [timeZone, provider, source] = flag.slice(EXIT_GEO_FLAG.length).split(",");
  if (
    !isKnownTimeZone(timeZone) ||
    !EXIT_GEO_PROVIDER_NAMES.has(provider) ||
    !EXIT_GEO_SOURCES.has(source)
  ) {
    return null;
  }
  return { timeZone, provider, source };
}

async function main(args) {
  const [mode, proxyServer = "", configuredAcceptLanguage = ""] = args;
  if (mode !== "--exit-locale") {
    console.error("usage: sand-egress-timezone.mjs --exit-locale <proxy-server> <accept-language>");
    process.exitCode = 2;
    return;
  }
  try {
    process.stdout.write(
      exitLocaleLine(await resolveExitLocale({ proxyServer, configuredAcceptLanguage })),
    );
  } catch (error) {
    console.error(`exit locale unresolved: ${egressErrorCode(error)}`);
    process.exitCode = 1;
  }
}

if (process.argv[1] != null && import.meta.url === pathToFileURL(process.argv[1]).href) {
  await main(process.argv.slice(2));
}
