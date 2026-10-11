import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";

export const LINUX_GPU_PATCH_PROFILE = {
  navigatorPlatform: "Linux x86_64",
  webglVendor: "Google Inc. (Intel)",
  webglRenderer: "ANGLE (Intel, Mesa Intel(R) UHD Graphics 630 (CFL GT2), OpenGL 4.6)",
  hardwareConcurrency: 8,
};

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

function buildUserAgentMetadata(product) {
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
  };
}

export function buildUserAgentOverride(product) {
  const { major } = chromeProductFromUserAgentProduct(product);
  return {
    userAgent: `Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/${major}.0.0.0 Safari/537.36`,
    acceptLanguage: "en-US,en",
    platform: "Linux x86_64",
    userAgentMetadata: buildUserAgentMetadata(product),
  };
}

export const CANVAS_MACHINE_ID_PATHS = Object.freeze([
  "/home/box/chrome-profile/machine-id",
  "/etc/machine-id",
]);

const CANVAS_SEED_DOMAIN = "sand-canvas-readback-v1";
const MACHINE_ID_RE = /^[0-9a-f]{32}$/;

export function machineIdFromText(text) {
  const id = String(text ?? "").trim();
  return MACHINE_ID_RE.test(id) ? id : null;
}

export function canvasSeedFromMachineId(machineId) {
  const id = machineIdFromText(machineId);
  if (id == null) return null;
  const digest = createHash("sha256").update(`${CANVAS_SEED_DOMAIN}:${id}`).digest();
  return digest.readUInt32BE(0);
}

export function readCanvasMachineId({
  paths = CANVAS_MACHINE_ID_PATHS,
  readFile = readFileSync,
} = {}) {
  for (const filePath of paths) {
    try {
      const id = machineIdFromText(readFile(filePath, "utf8"));
      if (id != null) return id;
    } catch (error) {
      if (
        error == null ||
        typeof error !== "object" ||
        !("code" in error) ||
        error.code !== "ENOENT"
      ) {
        console.error(
          `canvas seed machine-id read failed: ${error instanceof Error ? error.name : typeof error}`,
        );
      }
    }
  }
  return null;
}

export function resolveCanvasSeed({ machineId, paths, readFile } = {}) {
  const id =
    machineId != null ? machineIdFromText(machineId) : readCanvasMachineId({ paths, readFile });
  return canvasSeedFromMachineId(id);
}

function canvasReadbackBlock(seed) {
  if (typeof seed !== "number" || !Number.isFinite(seed)) return "";
  const canvasSeed = seed >>> 0;
  return `
  const CANVAS_SEED = ${canvasSeed};
  const OffscreenCanvasCtor = typeof OffscreenCanvas === "function" ? OffscreenCanvas : null;
  const ImageDataCtor = typeof ImageData === "function" ? ImageData : null;
  const htmlProto = typeof HTMLCanvasElement === "function" ? HTMLCanvasElement.prototype : null;
  const offProto = OffscreenCanvasCtor == null ? null : OffscreenCanvasCtor.prototype;
  const ctx2dProto = typeof CanvasRenderingContext2D === "function" ? CanvasRenderingContext2D.prototype : null;
  const offCtxProto = typeof OffscreenCanvasRenderingContext2D === "function" ? OffscreenCanvasRenderingContext2D.prototype : null;
  const origToDataURL = htmlProto && htmlProto.toDataURL;
  const origToBlob = htmlProto && htmlProto.toBlob;
  const origConvertToBlob = offProto && offProto.convertToBlob;
  const origHtmlGetContext = htmlProto && htmlProto.getContext;
  const origOffGetContext = offProto && offProto.getContext;
  const origHtmlPut = ctx2dProto && ctx2dProto.putImageData;
  const origOffPut = offCtxProto && offCtxProto.putImageData;
  const origOffGetImage = offCtxProto && offCtxProto.getImageData;
  const origOffDraw = offCtxProto && offCtxProto.drawImage;
  const nativeCreate = typeof document !== "undefined" && document != null ? document.createElement : null;
  const imageProto = ImageDataCtor == null ? null : ImageDataCtor.prototype;
  const protoAccessor = (proto, name, kind) => {
    for (let current = proto; current; current = Object.getPrototypeOf(current)) {
      const desc = Object.getOwnPropertyDescriptor(current, name);
      if (!desc) continue;
      const fn = desc[kind];
      return typeof fn === "function" ? fn : null;
    }
    return null;
  };
  const readAccessor = (getter, obj) => (typeof getter === "function" ? getter.call(obj) >>> 0 : 0);
  const htmlWidthGet = protoAccessor(htmlProto, "width", "get");
  const htmlWidthSet = protoAccessor(htmlProto, "width", "set");
  const htmlHeightGet = protoAccessor(htmlProto, "height", "get");
  const htmlHeightSet = protoAccessor(htmlProto, "height", "set");
  const offWidthGet = protoAccessor(offProto, "width", "get");
  const offHeightGet = protoAccessor(offProto, "height", "get");
  const imageWidthGet = protoAccessor(imageProto, "width", "get");
  const imageHeightGet = protoAccessor(imageProto, "height", "get");
  const imageDataGet = protoAccessor(imageProto, "data", "get");
  const U8 = typeof Uint8Array === "function" ? Uint8Array : null;
  const U8C = typeof Uint8ClampedArray === "function" ? Uint8ClampedArray : null;
  const taProto = U8 == null ? null : U8.prototype;
  const tcProto = U8C == null ? null : U8C.prototype;
  const tcSet = tcProto && tcProto.set;
  const taLengthGet = protoAccessor(taProto, "length", "get");
  const canvasSize = (canvas) => {
    if (typeof htmlWidthGet === "function" && typeof htmlHeightGet === "function") {
      try {
        return [htmlWidthGet.call(canvas) >>> 0, htmlHeightGet.call(canvas) >>> 0];
      } catch (e) {}
    }
    return [readAccessor(offWidthGet, canvas), readAccessor(offHeightGet, canvas)];
  };
  const mix = (n) => {
    n >>>= 0;
    n = Math.imul(n ^ (n >>> 16), 0x7feb352d);
    n = Math.imul(n ^ (n >>> 15), 0x846ca68b);
    return (n ^ (n >>> 16)) >>> 0;
  };
  const applyNoise = (data, width, height, originX, originY) => {
    let opaque = 0;
    const count = width * height;
    for (let p = 0; p < count; p++) {
      if (data[p * 4 + 3] !== 0) opaque++;
    }
    if (opaque === 0) return false;
    const pick = mix(CANVAS_SEED ^ 0x9e3779b9) % opaque;
    const channel = mix(CANVAS_SEED ^ 0x85ebca6b) % 3;
    let seen = 0;
    let chosen = 0;
    for (let y = 0; y < height; y++) {
      for (let x = 0; x < width; x++) {
        const i = (y * width + x) * 4;
        if (data[i + 3] === 0) continue;
        if (seen === pick) {
          chosen = i;
        } else {
          const h = mix(
            CANVAS_SEED ^ Math.imul(originX + x + 1, 0x9e3779b1) ^ Math.imul(originY + y + 1, 0x85ebca6b),
          );
          if ((h & 63) === 0) {
            const sparseChannel = (h >>> 6) % 3;
            data[i + sparseChannel] = (data[i + sparseChannel] & 0xfe) | ((h >>> 9) & 1);
          }
        }
        seen++;
      }
    }
    const chosenHash = mix(CANVAS_SEED ^ 0x85ebca6b);
    data[chosen + channel] = (data[chosen + channel] & 0xfe) | ((chosenHash >>> 3) & 1);
    return true;
  };
  const install = (proto, name, fn) => {
    if (!proto || typeof proto[name] !== "function" || masked.has(proto[name])) return;
    masked.add(fn);
    try {
      Object.defineProperty(proto, name, { value: fn, writable: true, configurable: true });
    } catch (e) {}
  };
  const exportImage = (canvas) => {
    if (OffscreenCanvasCtor == null || ImageDataCtor == null || U8C == null || typeof imageDataGet !== "function") return null;
    const sized = canvasSize(canvas);
    const w = sized[0];
    const h = sized[1];
    if (w < 2 || h < 2 || origOffGetContext == null || origOffDraw == null || origOffGetImage == null) {
      return null;
    }
    const scratch = new OffscreenCanvasCtor(w, h);
    const ctx = origOffGetContext.call(scratch, "2d");
    if (!ctx) return null;
    origOffDraw.call(ctx, canvas, 0, 0);
    const image = origOffGetImage.call(ctx, 0, 0, w, h);
    const pixels = imageDataGet.call(image);
    const copy = new U8C(pixels);
    if (!applyNoise(copy, w, h, 0, 0)) return null;
    return new ImageDataCtor(copy, w, h);
  };
  const htmlScratch = (image) => {
    if (nativeCreate == null || origHtmlGetContext == null || origHtmlPut == null) return null;
    if (typeof htmlWidthSet !== "function" || typeof htmlHeightSet !== "function") return null;
    const canvas = nativeCreate.call(document, "canvas");
    const w = readAccessor(imageWidthGet, image);
    const h = readAccessor(imageHeightGet, image);
    htmlWidthSet.call(canvas, w);
    htmlHeightSet.call(canvas, h);
    const ctx = origHtmlGetContext.call(canvas, "2d");
    if (!ctx) return null;
    origHtmlPut.call(ctx, image, 0, 0);
    return canvas;
  };
  const offscreenScratch = (image) => {
    if (OffscreenCanvasCtor == null || origOffGetContext == null || origOffPut == null) return null;
    const canvas = new OffscreenCanvasCtor(
      readAccessor(imageWidthGet, image),
      readAccessor(imageHeightGet, image),
    );
    const ctx = origOffGetContext.call(canvas, "2d");
    if (!ctx) return null;
    origOffPut.call(ctx, image, 0, 0);
    return canvas;
  };
  if (htmlProto && origToDataURL) {
    install(htmlProto, "toDataURL", {
      toDataURL() {
        try {
          const image = exportImage(this);
          const scratch = image && htmlScratch(image);
          if (scratch) return origToDataURL.call(scratch, arguments[0], arguments[1]);
        } catch (e) {}
        return origToDataURL.apply(this, arguments);
      },
    }.toDataURL);
  }
  if (htmlProto && origToBlob) {
    install(htmlProto, "toBlob", {
      toBlob(callback) {
        try {
          const image = exportImage(this);
          const scratch = image && htmlScratch(image);
          if (scratch) return origToBlob.call(scratch, callback, arguments[1], arguments[2]);
        } catch (e) {}
        return origToBlob.apply(this, arguments);
      },
    }.toBlob);
  }
  if (offProto && origConvertToBlob) {
    install(offProto, "convertToBlob", {
      convertToBlob() {
        try {
          const image = exportImage(this);
          const scratch = image && offscreenScratch(image);
          if (scratch) return origConvertToBlob.call(scratch, arguments[0]);
        } catch (e) {}
        return origConvertToBlob.apply(this, arguments);
      },
    }.convertToBlob);
  }
  const patchGetImageData = (proto) => {
    if (!proto || typeof proto.getImageData !== "function") return;
    const original = proto.getImageData;
    install(proto, "getImageData", {
      getImageData(sx, sy, sw, sh) {
        const image = original.apply(this, arguments);
        try {
          if (typeof imageDataGet !== "function" || U8C == null || typeof tcSet !== "function") return image;
          const data = imageDataGet.call(image);
          const w = readAccessor(imageWidthGet, image);
          const h = readAccessor(imageHeightGet, image);
          if (w < 2 || h < 2 || data == null) return image;
          const copy = new U8C(data);
          if (!applyNoise(copy, w, h, sx | 0, sy | 0)) return image;
          tcSet.call(data, copy);
        } catch (e) {}
        return image;
      },
    }.getImageData);
  };
  patchGetImageData(ctx2dProto);
  patchGetImageData(offCtxProto);
  const patchReadPixels = (proto, webgl2) => {
    if (!proto || typeof proto.readPixels !== "function") return;
    const original = proto.readPixels;
    const getParam = typeof proto.getParameter === "function" ? proto.getParameter : null;
    const bufferHeightGet = protoAccessor(proto, "drawingBufferHeight", "get");
    install(proto, "readPixels", {
      readPixels(x, y, width, height, format, type, pixels) {
        original.apply(this, arguments);
        try {
          if (pixels == null || U8 == null || taLengthGet == null) return;
          let pixelLength;
          try {
            pixelLength = taLengthGet.call(pixels);
          } catch (e) {
            return;
          }
          if (typeof pixelLength !== "number") return;
          if ((format | 0) !== 0x1908 || (type | 0) !== 0x1401) return;
          const w = width | 0;
          const h = height | 0;
          if (w < 2 || h < 2) return;
          const dstOffset = arguments.length > 7 && typeof arguments[7] === "number" ? arguments[7] >>> 0 : 0;
          const query = (pname) => getParam.call(this, pname);
          if (webgl2 && getParam != null) {
            const rowLength = query(0x0D02) | 0;
            const skipRows = query(0x0D03) | 0;
            const skipPixels = query(0x0D04) | 0;
            if (rowLength !== 0 || skipRows !== 0 || skipPixels !== 0) return;
          }
          let alignment = 4;
          try {
            const packed = getParam == null ? 4 : query(0x0D05);
            if (packed === 1 || packed === 2 || packed === 4 || packed === 8) alignment = packed;
          } catch (e) {}
          const rowBytes = w * 4;
          const stride = Math.ceil(rowBytes / alignment) * alignment;
          const byteLength = h === 0 ? 0 : (h - 1) * stride + rowBytes;
          if (dstOffset > pixelLength || byteLength > pixelLength - dstOffset) return;
          const glX = x | 0;
          const glY = y | 0;
          const bufferH = readAccessor(bufferHeightGet, this);
          const flip = bufferH >= glY + h;
          const topDown = new U8(w * h * 4);
          for (let row = 0; row < h; row++) {
            const srcRow = flip ? h - 1 - row : row;
            const src = dstOffset + srcRow * stride;
            const dst = row * rowBytes;
            for (let i = 0; i < rowBytes; i++) topDown[dst + i] = pixels[src + i];
          }
          if (!applyNoise(topDown, w, h, glX, flip ? bufferH - glY - h : glY)) return;
          for (let row = 0; row < h; row++) {
            const dstRow = flip ? h - 1 - row : row;
            const src = row * rowBytes;
            const dst = dstOffset + dstRow * stride;
            for (let i = 0; i < rowBytes; i++) pixels[dst + i] = topDown[src + i];
          }
        } catch (e) {}
      },
    }.readPixels);
  };
  try { patchReadPixels(WebGLRenderingContext.prototype, false); } catch (e) {}
  try { patchReadPixels(WebGL2RenderingContext.prototype, true); } catch (e) {}
`;
}

function navigatorNumberPin(name, value) {
  const key = JSON.stringify(name);
  return `
    try {
      if (${key} in navProto) {
        const getter = Object.getOwnPropertyDescriptor({
          get [${key}]() { return ${JSON.stringify(value)}; },
        }, ${key}).get;
        masked.add(getter);
        Object.defineProperty(navProto, ${key}, {
          get: getter, enumerable: true, configurable: true,
        });
      }
    } catch (e) {}`;
}

export function buildHardwareProfileDocumentScript(profile, { canvasSeed = null } = {}) {
  const hardwarePins = [];
  if (typeof profile.hardwareConcurrency === "number") {
    hardwarePins.push(navigatorNumberPin("hardwareConcurrency", profile.hardwareConcurrency));
  }
  if (typeof profile.deviceMemory === "number") {
    hardwarePins.push(navigatorNumberPin("deviceMemory", profile.deviceMemory));
  }
  if (typeof profile.maxTouchPoints === "number") {
    hardwarePins.push(navigatorNumberPin("maxTouchPoints", profile.maxTouchPoints));
  }
  return `(() => {
  const VENDOR = ${JSON.stringify(profile.webglVendor)};
  const RENDERER = ${JSON.stringify(profile.webglRenderer)};
  const PLATFORM = ${JSON.stringify(profile.navigatorPlatform)};
  const UNMASKED_VENDOR = 0x9245, UNMASKED_RENDERER = 0x9246;
  const nativeToString = Function.prototype.toString;
  const masked = new WeakSet();
  const fakeToString = {
    toString() {
      if (masked.has(this)) {
        return "function " + (this.name || "") + "() { [native code] }";
      }
      return nativeToString.call(this);
    },
  }.toString;
  masked.add(fakeToString);
  try {
    Object.defineProperty(Function.prototype, "toString", {
      value: fakeToString, writable: true, configurable: true,
    });
  } catch (e) {}
  const patchGetParameter = (proto) => {
    if (!proto || !proto.getParameter) return;
    const original = proto.getParameter;
    if (masked.has(original)) return;
    const wrapped = {
      getParameter(pname) {
        if (pname === UNMASKED_VENDOR) return VENDOR;
        if (pname === UNMASKED_RENDERER) return RENDERER;
        return original.call(this, pname);
      },
    }.getParameter;
    masked.add(wrapped);
    try {
      Object.defineProperty(proto, "getParameter", {
        value: wrapped, writable: true, configurable: true,
      });
    } catch (e) {}
  };
  try { patchGetParameter(WebGLRenderingContext.prototype); } catch (e) {}
  try { patchGetParameter(WebGL2RenderingContext.prototype); } catch (e) {}
${canvasReadbackBlock(canvasSeed)}
  const navProto = (() => {
    try {
      const Nav = typeof Navigator === "function"
        ? Navigator
        : typeof WorkerNavigator === "function"
          ? WorkerNavigator
          : null;
      return Nav == null ? null : Nav.prototype;
    } catch (e) {
      return null;
    }
  })();
  if (navProto) {
    try {
      if ("platform" in navProto) {
        const platformGetter = Object.getOwnPropertyDescriptor({
          get platform() { return PLATFORM; },
        }, "platform").get;
        masked.add(platformGetter);
        Object.defineProperty(navProto, "platform", {
          get: platformGetter, enumerable: true, configurable: true,
        });
      }
    } catch (e) {}
${hardwarePins.join("")}
  }
})();`;
}
