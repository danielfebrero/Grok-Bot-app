import { BOX_BROWSER_GPU_MODES } from "./box-contract.generated.mjs";

/*
The coherent profile claims only what the launched browser backs. Stock
Chrome on the GPU-less box renders WebGL on SwiftShader, on Mesa llvmpipe, or
not at all, and no real GPU's limits, extensions, and rendered pixels can be
reproduced from outside the browser in software, so WebGL is left exactly as
box-chrome launched it. Every other claim is a CDP override sent to each page, frame,
and worker target, so no realm runs injected code
(https://chromedevtools.github.io/devtools-protocol/tot/Emulation/#method-setHardwareConcurrencyOverride).
*/
const COHERENT_WEBGL_BY_GPU_MODE = Object.freeze({
  swiftshader: "swiftshader",
  "angle-gl": "none",
  stock: "none",
  llvmpipe: "angle-opengl",
  legacy: "unverified",
  unknown: "unverified",
});

export function normalizeGpuMode(value) {
  const mode = String(value ?? "").trim();
  return BOX_BROWSER_GPU_MODES.includes(mode) ? mode : "unknown";
}

export function coherentHardwareProfile(gpuMode) {
  const mode = normalizeGpuMode(gpuMode);
  return Object.freeze({
    gpuMode: mode,
    navigatorPlatform: "Linux x86_64",
    hardwareConcurrency: 8,
    webgl: COHERENT_WEBGL_BY_GPU_MODE[mode],
  });
}

const ANGLE_BACKEND_RE = /\bangle=([a-z0-9-]+)/;

export function observedWebglFromGpuInfo(gpu) {
  const status = gpu?.featureStatus?.webgl;
  if (typeof status !== "string") return "unknown";
  if (status.startsWith("disabled")) return "none";
  const parts = gpu?.auxAttributes?.glImplementationParts;
  const backend = typeof parts === "string" ? ANGLE_BACKEND_RE.exec(parts)?.[1] : undefined;
  if (backend === "swiftshader") return "swiftshader";
  return backend == null ? "unknown" : `angle-${backend}`;
}

export function verifyCoherentWebgl(profile, gpu) {
  const observed = observedWebglFromGpuInfo(gpu);
  if (profile.webgl === "unverified" || observed === "unknown") {
    return { claimed: profile.webgl, observed, verdict: "unverified" };
  }
  return {
    claimed: profile.webgl,
    observed,
    verdict: observed === profile.webgl ? "match" : "mismatch",
  };
}

// Chromium seeds the GREASE brand, its version, and the brand order with the major version (https://chromium.googlesource.com/chromium/src/+/refs/tags/154.0.8037.57/components/embedder_support/user_agent_utils.cc).
const GREASE_CHARS = [" ", "(", ":", "-", ".", "/", ")", ";", "=", "?", "_"];
const GREASE_VERSIONS = ["8", "99", "24"];
const BRAND_ORDERS = [
  [0, 1, 2],
  [0, 2, 1],
  [1, 0, 2],
  [1, 2, 0],
  [2, 0, 1],
  [2, 1, 0],
];

export function chromeProductFromUserAgentProduct(product) {
  const m = /Chrome\/(\d+)\.(\d+)\.(\d+)\.(\d+)/.exec(product ?? "");
  if (m == null) return { major: "0", full: "0.0.0.0" };
  return { major: m[1], full: `${m[1]}.${m[2]}.${m[3]}.${m[4]}` };
}

export function greaseBrandVersion(major) {
  const seed = Number.parseInt(major, 10);
  const brand = `Not${GREASE_CHARS[seed % GREASE_CHARS.length]}A${GREASE_CHARS[(seed + 1) % GREASE_CHARS.length]}Brand`;
  const short = GREASE_VERSIONS[seed % GREASE_VERSIONS.length];
  return { brand, short, full: `${short}.0.0.0` };
}

function shuffleBrands(major, entries) {
  const order = BRAND_ORDERS[Number.parseInt(major, 10) % BRAND_ORDERS.length];
  const shuffled = new Array(entries.length);
  for (let i = 0; i < order.length; i++) shuffled[order[i]] = entries[i];
  return shuffled;
}

function chromeBrandList(major, full) {
  const grease = greaseBrandVersion(major);
  return {
    brands: shuffleBrands(major, [
      { brand: grease.brand, version: grease.short },
      { brand: "Chromium", version: major },
      { brand: "Google Chrome", version: major },
    ]),
    fullVersionList: shuffleBrands(major, [
      { brand: grease.brand, version: grease.full },
      { brand: "Chromium", version: full },
      { brand: "Google Chrome", version: full },
    ]),
  };
}

function buildUserAgentMetadata(product, acceptLanguageEnUs) {
  const { major, full } = chromeProductFromUserAgentProduct(product);
  const { brands, fullVersionList } = chromeBrandList(major, full);
  return {
    brands,
    fullVersionList,
    fullVersion: full,
    platform: "Linux",
    platformVersion: "",
    architecture: "x86",
    model: "",
    mobile: false,
    bitness: "64",
    wow64: false,
    ...(acceptLanguageEnUs ? { formFactors: ["Desktop"] } : {}),
  };
}

const ACCEPT_LANGUAGE_LIST_RE =
  /^[A-Za-z]{1,8}(?:-[A-Za-z0-9]{1,8})*(?:,[A-Za-z]{1,8}(?:-[A-Za-z0-9]{1,8})*)*$/;

export function configuredAcceptLanguage(value) {
  return typeof value === "string" && value.length <= 128 && ACCEPT_LANGUAGE_LIST_RE.test(value)
    ? value
    : "en-US,en-GB,en";
}

function overrideAcceptLanguage(acceptLanguageEnUs, acceptLanguage, fromChrome, overridden) {
  if (!acceptLanguageEnUs) return { acceptLanguage: "en-US,en" };
  if (fromChrome) return overridden ? { acceptLanguage: "" } : {};
  return { acceptLanguage: configuredAcceptLanguage(acceptLanguage) };
}

export function buildUserAgentOverride(
  product,
  {
    acceptLanguageEnUs = false,
    acceptLanguage = "",
    acceptLanguageFromChrome = false,
    acceptLanguageOverridden = false,
  } = {},
) {
  const { major } = chromeProductFromUserAgentProduct(product);
  return {
    userAgent: `Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/${major}.0.0.0 Safari/537.36`,
    ...overrideAcceptLanguage(
      acceptLanguageEnUs,
      acceptLanguage,
      acceptLanguageFromChrome,
      acceptLanguageOverridden,
    ),
    platform: "Linux x86_64",
    userAgentMetadata: buildUserAgentMetadata(product, acceptLanguageEnUs),
  };
}

// https://chromium.googlesource.com/chromium/src/+/refs/tags/156.0.8078.47/net/http/http_util.cc
export function nativeAcceptLanguageHeader(list) {
  const tags = configuredAcceptLanguage(list).split(",");
  const baseOf = (tag) => tag.split("-")[0];
  const expanded = [];
  tags.forEach((tag, index) => {
    const base = baseOf(tag);
    const withBase =
      base === "x" || base === "i" || baseOf(tags[index + 1] ?? "") === base ? [tag] : [tag, base];
    for (const entry of withBase) if (!expanded.includes(entry)) expanded.push(entry);
  });
  return expanded
    .map((tag, index) => (index === 0 ? tag : `${tag};q=0.${Math.max(10 - index, 1)}`))
    .join(",");
}

export const ACCEPT_LANGUAGE_EXTENSION_NAME = "Sand Accept-Language";

const DNR_RESOURCE_TYPES = Object.freeze([
  "main_frame",
  "sub_frame",
  "stylesheet",
  "script",
  "image",
  "font",
  "object",
  "xmlhttprequest",
  "ping",
  "csp_report",
  "media",
  "websocket",
  "webtransport",
  "webbundle",
  "other",
]);

const REGULAR_TAB_ALLOW_WORKER = `let queue = Promise.resolve();
function sync() {
  queue = queue
    .then(async () => {
      const tabs = await chrome.tabs.query({});
      const regular = tabs.filter((tab) => !tab.incognito && tab.id >= 0).map((tab) => tab.id);
      await chrome.declarativeNetRequest.updateSessionRules({
        removeRuleIds: [2],
        addRules: regular.length === 0 ? [] : [{ id: 2, priority: 2, action: { type: "allow" }, condition: { tabIds: regular, resourceTypes: ${JSON.stringify(DNR_RESOURCE_TYPES)} } }],
      });
    })
    .catch(() => {});
}
chrome.tabs.onCreated.addListener(sync);
chrome.tabs.onRemoved.addListener(sync);
chrome.tabs.onReplaced.addListener(sync);
sync();
`;

// https://developer.chrome.com/docs/extensions/reference/api/declarativeNetRequest#rules
export function buildAcceptLanguageExtension(list) {
  const manifest = {
    manifest_version: 3,
    name: ACCEPT_LANGUAGE_EXTENSION_NAME,
    version: "1",
    incognito: "spanning",
    permissions: ["declarativeNetRequest"],
    host_permissions: ["<all_urls>"],
    background: { service_worker: "worker.js" },
    declarative_net_request: {
      rule_resources: [{ id: "accept_language", enabled: true, path: "rules.json" }],
    },
  };
  const rules = [
    {
      id: 1,
      priority: 1,
      action: {
        type: "modifyHeaders",
        requestHeaders: [
          { header: "accept-language", operation: "set", value: nativeAcceptLanguageHeader(list) },
        ],
      },
      condition: { resourceTypes: DNR_RESOURCE_TYPES },
    },
  ];
  return {
    "manifest.json": `${JSON.stringify(manifest)}\n`,
    "rules.json": `${JSON.stringify(rules)}\n`,
    "worker.js": REGULAR_TAB_ALLOW_WORKER,
  };
}
