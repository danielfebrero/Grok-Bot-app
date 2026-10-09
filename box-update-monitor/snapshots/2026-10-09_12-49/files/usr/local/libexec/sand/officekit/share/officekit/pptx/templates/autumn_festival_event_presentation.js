/**
 * "Autumn Festival" event-presentation template — 25 slides, 10 x 5.625 in.
 *
 * Rebuilt from scratch with pptxgenjs. Every photograph in the source deck is
 * replaced by a flat "[image]" placeholder that keeps the original silhouette
 * (rectangle / pill / arch) and footprint.
 *
 *   node build.js   ->   <script dir>/<name>.pptx
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ============================================================== palette ==*/

const C = {
  ink: '000000',      // dk1 — headings
  body: '3F3F3F',     // long-form copy
  muted: '757070',    // secondary copy
  grey: '7F7F7F',
  silver: 'BFBFBF',
  mist: 'F2F2F2',     // photo placeholder / soft cards
  cream: 'FFF4E5',    // large tinted panels + hairlines
  blush: 'FAEBD8',    // cream over a photo
  peach: 'FEE9CD',
  amber: 'FECB86',    // accent1
  orange: 'FDA125',   // maple-leaf blade
  rust: 'C06F01',     // maple-leaf veins
  clay: 'D27350',     // maple-leaf secondary veins
  slate: '2D3847',    // dk2 — pictograms
  white: 'FFFFFF',
};

const SCRIPT = 'Dancing Script SemiBold'; // display face
const SANS = 'Raleway Light';             // text face
const INSET = [5.4, 5.4, 2.7, 2.7]; // l, r, b, t insets in points — Google-Slides default

const PHOTO = C.mist;                        // fill used for image placeholders
const LOREM = 'Lorem ipsum dolor sit amet, ';
const LOREM_LONG = LOREM + 'consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus';

/* ============================================================== drawing ==*/

function text(slide, str, opts) {
  slide.addText(str, Object.assign(
    { fontFace: SANS, color: C.ink, margin: INSET, valign: 'top', wrap: true }, opts));
}

function panel(slide, x, y, w, h, fill) {
  slide.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: fill }, line: { type: 'none' } });
}

function divider(slide, x, y, w, h, opts) {
  const o = opts || {};
  slide.addShape('line', { x: x, y: y, w: w, h: h, line: { color: o.color || C.cream, width: o.width || 2 } });
}

// 12 pt bold lead-in over 9 pt grey copy — the deck's most repeated pattern.
function captionBlock(slide, x, y, w, headline, body, opts) {
  const o = opts || {};
  text(slide, headline, { x: x, y: y, w: o.headW || w, h: 0.278, fontSize: 12, bold: true, align: o.align });
  text(slide, body, {
    x: x, y: y + (o.gap || 0.278), w: w, h: o.bodyH || 0.647, fontSize: 9, color: o.bodyColor || C.body,
    lineSpacingMultiple: 1.3, align: o.align,
  });
}

// Oversized figure over its label ("80%" / "Our Participant").
function statBlock(slide, x, y, w, value, label, opts) {
  const o = opts || {};
  text(slide, value, { x: x, y: y, w: w, h: 0.581, fontSize: o.valueSize || 30, bold: true });
  text(slide, label, {
    x: x, y: y + (o.gap || 0.625), w: o.labelW || w, h: o.labelH || 0.647, fontSize: o.labelSize || 9,
    bold: !!o.labelBold, color: o.labelColor || C.muted, lineSpacingMultiple: 1.3,
  });
}

// Silhouette helpers -------------------------------------------------------
// `kind` is one of 'rect' | 'pill' | 'hpill' | 'arch' | 'dome'.
//   pill  – both short ends fully rounded (vertical capsule)
//   hpill – both short ends fully rounded (horizontal capsule)
//   arch  – top corners rounded by w/2, square foot
//   dome  – square top, bottom corners rounded by w/2

function silhouette(slide, x, y, w, h, kind, fill, opts) {
  const o = opts || {};
  const style = { fill: { color: fill }, line: o.line || { type: 'none' } };
  if (kind === 'rect') slide.addShape('rect', Object.assign({ x: x, y: y, w: w, h: h }, style));
  else if (kind === 'pill' || kind === 'hpill') {
    slide.addShape('roundRect', Object.assign({ x: x, y: y, w: w, h: h, rectRadius: Math.min(w, h) / 2 }, style));
  } else {
    // arch / dome — half-rounded, needs a custom outline
    const r = w / 2, up = kind === 'arch';
    const pts = [];
    if (up) {
      pts.push({ x: 0, y: r, moveTo: true });
      pts.push({ x: r, y: 0, curve: { type: 'arc', wR: r, hR: r, stAng: 180, swAng: 90 } });
      pts.push({ x: w, y: r, curve: { type: 'arc', wR: r, hR: r, stAng: 270, swAng: 90 } });
      pts.push({ x: w, y: h }); pts.push({ x: 0, y: h });
    } else {
      pts.push({ x: 0, y: 0, moveTo: true }); pts.push({ x: w, y: 0 });
      pts.push({ x: w, y: h - r });
      pts.push({ x: r, y: h, curve: { type: 'arc', wR: r, hR: r, stAng: 0, swAng: 90 } });
      pts.push({ x: 0, y: h - r, curve: { type: 'arc', wR: r, hR: r, stAng: 90, swAng: 90 } });
    }
    pts.push({ close: true });
    slide.addShape('custGeom', Object.assign({ x: x, y: y, w: w, h: h, points: pts }, style));
  }
}

// Image placeholder: silhouette in placeholder grey + a centred "[image]" tag.
function photo(slide, x, y, w, h, kind, opts) {
  const o = opts || {};
  silhouette(slide, x, y, w, h, kind || 'rect', o.fill || PHOTO, o);
  if (o.label === false) return;
  slide.addText('[image]', {
    x: x, y: y + h / 2 - 0.16, w: w, h: 0.32, margin: 0,
    align: 'center', valign: 'middle', fontFace: SANS, fontSize: o.labelSize || 9, color: o.labelColor || C.grey,
  });
}

function mixHex(a, b, t) {
  let out = '';
  for (let i = 0; i < 3; i++) {
    const ca = parseInt(a.substr(i * 2, 2), 16), cb = parseInt(b.substr(i * 2, 2), 16);
    out += ('0' + Math.round(ca + (cb - ca) * t).toString(16)).slice(-2);
  }
  return out.toUpperCase();
}

// The source deck fades photos out with translucent gradient overlays.
// pptxgenjs has no gradient fill, so each wash is painted as a stack of thin
// bands whose colour is pre-composited over `base`. A wash is a list of
// [position, colour, alpha] stops with position 0 at the BOTTOM of the box;
// several washes can be layered, exactly as they are in the source file.
function washBox(slide, x, y, w, h, kind, layers, base, opts) {
  const o = opts || {};
  // Curved silhouettes need many thin bands or the arc looks like a staircase.
  const bands = o.bands || (kind === 'rect' ? 34 : Math.round(h * 40));
  const r = kind === 'rect' ? 0 : w / 2;
  const sample = (stops, u) => {
    let k = 1;
    while (k < stops.length - 1 && stops[k][0] < u) k++;
    const [p0, c0, a0] = stops[k - 1], [p1, c1, a1] = stops[k];
    const t = p1 === p0 ? 0 : Math.max(0, Math.min(1, (u - p0) / (p1 - p0)));
    return [mixHex(c0, c1, t), a0 + (a1 - a0) * t];
  };
  for (let i = 0; i < bands; i++) {
    const y0 = h * i / bands, y1 = h * (i + 1) / bands, ym = (y0 + y1) / 2;
    const u = 1 - ym / h;                 // 0 at the bottom .. 1 at the top
    let color = base;
    layers.forEach(stops => {
      const [c, a] = sample(stops, u);
      color = mixHex(color, c, a);
    });
    let inset = 0;
    if (r > 0 && (kind === 'arch' || kind === 'pill') && ym < r) inset = r - Math.sqrt(r * r - (r - ym) * (r - ym));
    if (r > 0 && (kind === 'dome' || kind === 'pill') && ym > h - r) inset = r - Math.sqrt(r * r - (r - (h - ym)) * (r - (h - ym)));
    panel(slide, x + inset, y + y0, w - 2 * inset, y1 - y0 + 0.012, color);
  }
}

// The gradient overlays used in the deck: amber held over the bottom third,
// fading to translucent grey at the top.
const WASH_SOFT = [[0, C.amber, 0.698], [0.29, C.amber, 0.698], [1, C.silver, 0.298]];
const WASH_DEEP = [[0, C.amber, 0.698], [0.29, C.amber, 0.698], [1, C.silver, 0.6]];
const WASH_ARCH = [[0, C.amber, 0.867], [0.22, C.amber, 0.867], [1, C.silver, 0.298]];
const WASH_ACCENT = [[0, C.amber, 1], [0.29, C.amber, 1], [1, C.silver, 0.298]];

/* ========================================================= maple motifs ==*/

// Outline of the maple leaf, normalised into a 1 x 1 box.
const LEAF_BODY = [[0.547,0.666],[0.383,0.755],[0.374,0.729],[0.265,0.738],[0.297,0.686],[0.08,0.477],[0,0.449],[0.252,0.423],[0.37,0.472],[0.325,0.281],[0.376,0.281],[0.372,0.179],[0.411,0.179],[0.584,0],[0.575,0.066],[0.653,0.213],[0.708,0.212],[0.679,0.313],[0.708,0.319],[0.684,0.44],[0.79,0.377],[1,0.362],[0.932,0.408],[0.872,0.531],[0.912,0.548],[0.83,0.587],[0.754,0.606],[0.814,0.656],[0.865,0.667],[0.704,0.702],[0.547,0.666]];
const LEAF_STEM = [[0.524,0.161],[0.527,0.474],[0.568,0.818],[0.543,1],[0.497,0.983],[0.533,0.904],[0.552,0.777],[0.52,0.44],[0.516,0.161],[0.524,0.161]];
const LEAF_VEINS_RUST = [
  [[0.2,0.489],[0.402,0.569],[0.545,0.651],[0.839,0.431]],
  [[0.213,0.556],[0.338,0.541],[0.282,0.466]],
  [[0.299,0.607],[0.425,0.58],[0.392,0.511]],
  [[0.369,0.688],[0.5,0.621],[0.479,0.541]],
  [[0.609,0.511],[0.603,0.607],[0.711,0.651],[0.77,0.658]],
  [[0.827,0.55],[0.672,0.556],[0.659,0.483]],
  [[0.807,0.395],[0.791,0.466]],
  [[0.377,0.332],[0.528,0.529]],
  [[0.659,0.344],[0.524,0.466]],
  [[0.629,0.297],[0.609,0.389],[0.659,0.392]],
  [[0.434,0.211],[0.518,0.322]],
  [[0.591,0.197],[0.518,0.28]],
];
const LEAF_VEINS_CLAY = [
  [[0.724,0.441],[0.724,0.516],[0.859,0.479]],
  [[0.399,0.419],[0.453,0.431],[0.453,0.353]],
  [[0.56,0.363],[0.563,0.43],[0.622,0.449]],
  [[0.472,0.211],[0.476,0.266],[0.434,0.28]],
  [[0.56,0.192],[0.56,0.232],[0.603,0.238]],
];

// The little leaf beside the "Autumn Fest" wordmark (blade + five veins).
const LOGO_BLADE = [[0.721,0.52],[0.645,0.205],[0.599,0.075],[0.55,0.004],[0.505,0.034],[0.497,0.077],[0.509,0.274],[0.494,0.349],[0.466,0.327],[0.355,0.037],[0.326,0.054],[0.312,0.213],[0.198,0.103],[0.176,0.124],[0.176,0.205],[0.098,0.133],[0.095,0.234],[0.001,0.251],[0.02,0.32],[0.116,0.422],[0.092,0.538],[0.186,0.571],[0.141,0.734],[0.163,0.742],[0.284,0.667],[0.306,0.674],[0.286,0.994],[0.315,0.992],[0.349,0.946],[0.456,0.76],[0.484,0.755],[0.49,0.778],[0.474,0.977],[0.505,0.992],[0.547,0.98],[0.595,0.935],[0.687,0.711],[0.721,0.519]];
const LOGO_VEINS = [
  [[0.144,0.303],[0.398,0.487],[0.513,0.531],[0.649,0.538],[0.762,0.489],[0.851,0.416],[0.936,0.301],[0.981,0.174],[1,0.183],[0.953,0.306],[0.873,0.414],[0.765,0.502],[0.65,0.55],[0.531,0.547],[0.396,0.499],[0.211,0.382],[0.139,0.322],[0.144,0.304]],
  [[0.548,0.117],[0.651,0.544],[0.626,0.678],[0.548,0.884],[0.628,0.659],[0.649,0.544],[0.548,0.117]],
  [[0.335,0.834],[0.423,0.643],[0.454,0.509],[0.363,0.147],[0.438,0.379],[0.458,0.526],[0.417,0.676],[0.335,0.834]],
  [[0.23,0.173],[0.292,0.307],[0.314,0.424],[0.29,0.539],[0.214,0.644],[0.288,0.536],[0.312,0.424],[0.293,0.32],[0.23,0.173]],
];

// Places `polys` (unit-square coordinates) inside the box x,y,w,h rotated by
// `rot` degrees about the box centre, as one filled or stroked custGeom shape.
function placedPolys(slide, x, y, w, h, rot, polys, style) {
  const a = (rot || 0) * Math.PI / 180, cs = Math.cos(a), sn = Math.sin(a);
  const cx = x + w / 2, cy = y + h / 2;
  const map = ([px, py]) => {
    const dx = x + px * w - cx, dy = y + py * h - cy;
    return [cx + dx * cs - dy * sn, cy + dx * sn + dy * cs];
  };
  const shapes = polys.map(p => p.map(map));
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  shapes.forEach(p => p.forEach(([px, py]) => {
    x0 = Math.min(x0, px); y0 = Math.min(y0, py); x1 = Math.max(x1, px); y1 = Math.max(y1, py);
  }));
  const pts = [];
  shapes.forEach(p => {
    p.forEach(([px, py], i) => pts.push({ x: +(px - x0).toFixed(4), y: +(py - y0).toFixed(4), moveTo: i === 0 }));
    if (style.fill) pts.push({ close: true });
  });
  slide.addShape('custGeom', Object.assign({ x: x0, y: y0, w: x1 - x0, h: y1 - y0, points: pts }, style));
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
}

function mapleLeaf(slide, x, y, w, h, rot) {
  const hair = { width: 0.75 };
  placedPolys(slide, x, y, w, h, rot, [LEAF_BODY], { fill: { color: C.orange }, line: { type: 'none' } });
  placedPolys(slide, x, y, w, h, rot, [LEAF_STEM], { fill: { color: C.rust }, line: { type: 'none' } });
  placedPolys(slide, x, y, w, h, rot, LEAF_VEINS_RUST, { line: Object.assign({ color: C.rust }, hair) });
  placedPolys(slide, x, y, w, h, rot, LEAF_VEINS_CLAY, { line: Object.assign({ color: C.clay }, hair) });
}

/* ========================================== master chrome (wordmark etc) ==*/

// Every slide except the five full-bleed ones carries the "Autumn Fest"
// wordmark top-left, its little leaf, and a page number bottom-right.
function chrome(slide, n) {
  text(slide, 'Autumn Fest', { x: 0.007, y: -0.032, w: 1.326, h: 0.379, fontFace: SCRIPT, fontSize: 18 });
  placedPolys(slide, 1.17, 0.107, 0.51, 0.279, -45, [LOGO_BLADE], { fill: { color: C.orange }, line: { type: 'none' } });
  placedPolys(slide, 1.17, 0.107, 0.51, 0.279, -45, LOGO_VEINS, { fill: { color: C.rust }, line: { type: 'none' } });
  text(slide, String(n), {
    x: 9.371, y: 5.215, w: 0.51, h: 0.215, align: 'center', fontFace: SCRIPT, fontSize: 8, color: 'A5A5A5',
  });
}

/* ============================================================ pictograms ==*/

// Small dk2 glyphs used on the agenda / chart / contact slides.
function pictogram(slide, kind, x, y, w, h, color) {
  const col = { color: color || C.slate }, none = { type: 'none' };
  if (kind === 'cloche') {                       // serving dome — "Gathering"
    slide.addShape('blockArc', {
      x: x, y: y, w: w, h: h * 2, angleRange: [180, 360], arcThicknessRatio: 0.16,
      fill: col, line: none,
    });
    slide.addShape('rect', { x: x - w * 0.06, y: y + h * 0.85, w: w * 1.12, h: h * 0.13, fill: col, line: none });
  } else if (kind === 'camera') {                // "Vacation"
    slide.addShape('roundRect', { x: x, y: y + h * 0.2, w: w, h: h * 0.8, rectRadius: 0.03, fill: col, line: none });
    slide.addShape('rect', { x: x + w * 0.3, y: y, w: w * 0.4, h: h * 0.25, fill: col, line: none });
    slide.addShape('ellipse', { x: x + w * 0.31, y: y + h * 0.36, w: w * 0.38, h: h * 0.46, fill: { color: C.white }, line: none });
  } else if (kind === 'chevron') {               // legend arrow ">"
    slide.addShape('custGeom', {
      x: x, y: y, w: w, h: h,
      points: [{ x: 0, y: 0, moveTo: true }, { x: w, y: h / 2 }, { x: 0, y: h },
               { x: w * 0.28, y: h / 2 }, { close: true }],
      fill: col, line: none,
    });
  } else if (kind === 'pin') {                   // map marker
    slide.addShape('ellipse', { x: x, y: y, w: w, h: w, fill: col, line: none });
    slide.addShape('custGeom', {
      x: x, y: y + w * 0.45, w: w, h: h - w * 0.45,
      points: [{ x: 0, y: 0, moveTo: true }, { x: w, y: 0 }, { x: w / 2, y: h - w * 0.45 }, { close: true }],
      fill: col, line: none,
    });
    slide.addShape('ellipse', { x: x + w * 0.3, y: y + w * 0.3, w: w * 0.4, h: w * 0.4, fill: { color: C.white }, line: none });
  } else if (kind === 'inbox') {                 // download tray
    slide.addShape('custGeom', {
      x: x, y: y + h * 0.45, w: w, h: h * 0.55,
      points: [{ x: 0, y: 0, moveTo: true }, { x: w, y: 0 }, { x: w, y: h * 0.55 }, { x: 0, y: h * 0.55 }, { close: true }],
      fill: col, line: none,
    });
    slide.addShape('custGeom', {
      x: x + w * 0.3, y: y, w: w * 0.4, h: h * 0.5,
      points: [{ x: w * 0.12, y: 0, moveTo: true }, { x: w * 0.28, y: 0 }, { x: w * 0.28, y: h * 0.28 },
               { x: w * 0.4, y: h * 0.28 }, { x: w * 0.2, y: h * 0.5 }, { x: 0, y: h * 0.28 },
               { x: w * 0.12, y: h * 0.28 }, { close: true }],
      fill: col, line: none,
    });
  } else if (kind === 'map') {                   // folded map
    for (let i = 0; i < 3; i++) {
      slide.addShape('custGeom', {
        x: x + i * w / 3, y: y, w: w / 3, h: h,
        points: [{ x: 0, y: i % 2 ? h * 0.15 : 0, moveTo: true }, { x: w / 3, y: i % 2 ? 0 : h * 0.15 },
                 { x: w / 3, y: i % 2 ? h * 0.85 : h }, { x: 0, y: i % 2 ? h : h * 0.85 }, { close: true }],
        fill: col, line: none,
      });
    }
  } else if (kind === 'globe') {
    slide.addShape('ellipse', { x: x, y: y, w: w, h: h, fill: col, line: none });
    slide.addShape('ellipse', { x: x + w * 0.34, y: y, w: w * 0.32, h: h, fill: { type: 'none' }, line: { color: C.white, width: 0.75 } });
    slide.addShape('line', { x: x, y: y + h / 2, w: w, h: 0, line: { color: C.white, width: 0.75 } });
  } else if (kind === 'phone') {
    slide.addShape('custGeom', {
      x: x, y: y, w: w, h: h,
      points: [{ x: 0, y: 0, moveTo: true }, { x: w * 0.42, y: 0 }, { x: w * 0.52, y: h * 0.3 },
               { x: w * 0.36, y: h * 0.44 }, { x: w * 0.62, y: h * 0.78 }, { x: w * 0.8, y: h * 0.6 },
               { x: w, y: h }, { x: w * 0.6, y: h }, { x: 0, y: h * 0.45 }, { close: true }],
      fill: col, line: none,
    });
  } else if (kind === 'envelope') {
    slide.addShape('rect', { x: x, y: y, w: w, h: h, fill: col, line: none });
    slide.addShape('custGeom', {
      x: x, y: y, w: w, h: h,
      points: [{ x: 0, y: 0, moveTo: true }, { x: w / 2, y: h * 0.6 }, { x: w, y: 0 }],
      fill: null, line: { color: C.white, width: 0.75 },
    });
  }
}

/* =============================================================== slides ==*/

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'AUTUMN', width: 10, height: 5.625 });
pptx.layout = 'AUTUMN';
pptx.author = 'Autumn Festival';
pptx.title = 'Autumn Festival — Event Presentation Template';

function newSlide(n) {
  const s = pptx.addSlide();
  s.background = { color: C.white };
  // Slides 1, 6, 7, 10 and 25 are full-bleed and hide the master furniture.
  if ([1, 6, 7, 10, 25].indexOf(n) === -1) chrome(s, n);
  return s;
}

/* -- 1 · Title ----------------------------------------------------------- */
function slide01() {
  const s = newSlide(1);
  panel(s, 5.0, 0, 5.0, 5.625, C.blush);                       // photo + peach veil
  s.addText('[image]', { x: 5.009, y: 2.65, w: 5.0, h: 0.32, align: 'center', margin: 0, fontFace: SANS, fontSize: 9, color: C.grey });
  silhouette(s, 3.609, 0.541, 2.801, 4.543, 'pill', C.amber);
  mapleLeaf(s, 2.466, 1.517, 1.555, 1.85, -60);
  text(s, 'Autumn', { x: 2.93, y: 1.628, w: 3.48, h: 1.287, fontFace: SCRIPT, fontSize: 72 });
  text(s, 'Festival', { x: 3.924, y: 2.494, w: 3.48, h: 1.287, fontFace: SCRIPT, fontSize: 72 });
  text(s, 'Event Presentation Template', { x: 0.58, y: 4.497, w: 3.029, h: 0.682, fontSize: 18, bold: true });
}

/* -- 2 · Introduction ---------------------------------------------------- */
function slide02() {
  const s = newSlide(2);
  silhouette(s, 2.355, 0.604, 2.902, 4.57, 'pill', 'E9E2D7');  // offset shadow behind the photo
  panel(s, 0, 0, 3.476, 5.625, C.cream);
  photo(s, 2.654, 0.535, 2.902, 4.57, 'pill');
  divider(s, 5.0, 2.812, 5.0, 0);
  captionBlock(s, 5.816, 1.332, 2.228, 'Autumn', LOREM + 'consectetur adipiscing elit, sed do eiusmod tempor incididunt ut', { bodyColor: C.muted });
  captionBlock(s, 7.199, 3.368, 2.228, 'A Journey Begins', LOREM + 'consectetur adipiscing elit, sed do eiusmod tempor incididunt ut', { bodyColor: C.muted });
  mapleLeaf(s, 0.431, 3.89, 1.361, 1.62, -135);
  text(s, 'Introduction', { x: -1.033, y: 2.224, w: 4.409, h: 1.186, fontFace: SCRIPT, fontSize: 66, rotate: 270 });
}

/* -- 3 · Picnic in the Park ---------------------------------------------- */
function slide03() {
  const s = newSlide(3);
  washBox(s, 0.57, 1.016, 3.59, 4.609, 'arch', [WASH_ARCH], PHOTO);
  s.addText('[image]', { x: 0.57, y: 3.16, w: 3.59, h: 0.32, align: 'center', margin: 0, fontFace: SANS, fontSize: 9, color: C.body });
  panel(s, 5.0, 0, 5.0, 5.625, C.cream);
  mapleLeaf(s, 5.275, 2.17, 1.555, 1.85, -60);
  text(s, 'Picnic in\nthe Park', { x: 5.776, y: 2.329, w: 3.48, h: 2.499, fontFace: SCRIPT, fontSize: 72 });
  captionBlock(s, 5.402, 0.457, 1.901, 'Relaxation', LOREM + 'consectetur adipiscing elit, sed do eiusmod tempor', { bodyColor: C.muted });
  captionBlock(s, 7.475, 0.456, 1.781, 'Enjoying', LOREM + 'consectetur adipiscing elit, sed do eiusmod tempor', { bodyColor: C.muted });
  text(s, 'Picnicking Year-Round', { x: 0.584, y: 4.497, w: 1.819, h: 0.682, fontSize: 18, bold: true });
}

/* -- 4 · Making Memories Together ---------------------------------------- */
function slide04() {
  const s = newSlide(4);
  mapleLeaf(s, 0.466, 0.88, 1.555, 1.85, -60);
  panel(s, 5.0, 0, 5.0, 3.984, C.cream);
  text(s, 'Making\nMemories\nTogether', { x: 0.811, y: 0.957, w: 4.189, h: 3.711, fontFace: SCRIPT, fontSize: 72 });
  captionBlock(s, 5.225, 4.187, 1.901, 'Photo Diaries', LOREM + 'consectetur adipiscing elit, sed do eiusmod tempor', { gap: 0.386 });
  captionBlock(s, 7.564, 4.187, 1.781, 'Storytelling', LOREM + 'consectetur adipiscing elit, sed do eiusmod tempor', { gap: 0.386 });
  divider(s, 5.0, 3.875, 0, 1.75);
  photo(s, 6.419, 0.488, 2.161, 2.968, 'pill');
}

/* -- 5 · Seasonal Sensations (contents) ---------------------------------- */
const CONTENTS = ['Embracing Weather', 'Kaleidoscope Colors', 'Celebrating with Family',
                  'Celebrating with Friends', 'Outdoor Adventures', 'Festivals and Foliage'];

function slide05() {
  const s = newSlide(5);
  panel(s, 5.0, 0, 5.0, 5.625, C.cream);
  mapleLeaf(s, 5.086, 0.268, 0.976, 1.161, -60);
  text(s, 'Seasonal Sensations', { x: 5.456, y: 0.459, w: 3.976, h: 0.682, fontFace: SCRIPT, fontSize: 36 });
  text(s, 'Content :', { x: 0.52, y: 1.419, w: 1.899, h: 0.379, fontSize: 18, bold: true });
  CONTENTS.forEach((item, i) => {
    text(s, item, {
      x: 1.229, y: 2.238 + i * 0.278, w: 2.271, h: 0.278, fontSize: 12, bold: true,
      bullet: { characterCode: '2022', indent: 17 },
    });
  });
  divider(s, 0, 0.451, 5.0, 0);
  divider(s, -0.004, 5.188, 5.0, 0);
  photo(s, 6.5, 1.217, 2.0, 3.949, 'pill');
}

/* -- 6 · Cozy Up --------------------------------------------------------- */
function slide06() {
  const s = newSlide(6);
  panel(s, 5.0, 0, 5.0, 2.812, C.cream);
  mapleLeaf(s, 0.464, 1.962, 1.23, 1.34, -135);
  text(s, 'Cozy Up', { x: -0.152, y: 1.195, w: 2.372, h: 0.909, fontFace: SCRIPT, fontSize: 50, rotate: 270 });
  captionBlock(s, 6.841, 0.459, 2.589, 'Autumn Picnic',
    LOREM + 'consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed', { bodyColor: C.muted });
  [['80%', 0.58], ['112+', 2.089], ['50k+', 3.597]].forEach(([v, x]) => {
    statBlock(s, x, 3.902, 1.509, v, LOREM + 'consectetuer adipiscing elit');
  });
  divider(s, 5.008, 2.783, 0, 2.842);
  photo(s, 1.489, 0, 3.092, 2.812, 'dome');
  photo(s, 6.55, 3.539, 1.361, 1.636, 'pill', { labelSize: 7 });
  photo(s, 8.1, 3.539, 1.361, 1.636, 'pill', { labelSize: 7 });
}

/* -- 7 · Festivals and Foliage ------------------------------------------- */
const FOLIAGE_ITEMS = [
  ['01.', 'Food Truck Roundup', 0.528],
  ['02.', 'Storytelling', 2.35],
  ['03.', 'Live Music', 4.184],
];

function slide07() {
  const s = newSlide(7);
  washBox(s, 5.0, 1.79, 5.0, 3.835, 'rect', [WASH_DEEP], PHOTO);
  photo(s, 0, 0, 1.548, 5.625, 'rect');
  divider(s, 1.557, 1.798, 3.443, 0);
  divider(s, 1.548, 3.724, 3.443, 0);
  panel(s, 5.0, 0, 5.0, 1.79, C.cream);
  mapleLeaf(s, 4.689, 0.529, 1.744, 2.075, -60);
  text(s, 'Festivals\nand\nFoliage', { x: 5.236, y: 0.96, w: 4.189, h: 3.711, fontFace: SCRIPT, fontSize: 72 });
  s.addText('[image]', { x: 6.5, y: 3.55, w: 2.0, h: 0.32, align: 'center', margin: 0, fontFace: SANS, fontSize: 9, color: C.body });
  FOLIAGE_ITEMS.forEach(([num, head, y]) => {
    text(s, num, { x: 1.995, y: y + 0.199, w: 0.63, h: 0.48, fontSize: 24, bold: true });
    captionBlock(s, 2.659, y, 1.902, head, LOREM + 'consectetuer adipiscing elit. Maecenas porttitor congue');
  });
}

/* -- 8 · Agenda (September calendar) ------------------------------------- */
const CAL_COLS = [0.668, 1.104, 1.54, 1.977, 2.413, 2.849, 3.286];
const CAL_ROWS = [2.79, 3.173, 3.556, 3.939, 4.323, 4.709];
// One entry per cell: [label, fill]  (null fill = outside the month)
const CAL_GRID = [
  [['25', null], ['26', null], ['27', null], ['28', null], ['29', null], ['30', null], ['31', null]],
  [['1', C.mist], ['2', C.mist], ['3', C.mist], ['4', C.mist], ['5', C.mist], ['6', C.mist], ['7', C.cream]],
  [['8', C.mist], ['9', C.mist], ['10', C.mist], ['11', C.mist], ['12', C.mist], ['13', C.mist], ['14', C.cream]],
  [['15', C.mist], ['16', C.mist], ['17', C.amber], ['18', C.mist], ['19', C.mist], ['20', C.mist], ['21', C.cream]],
  [['22', C.mist], ['23', C.mist], ['24', C.mist], ['25', C.mist], ['26', C.mist], ['27', C.mist], ['28', C.cream]],
  [['29', C.mist], ['30', C.mist], ['31', C.mist], ['1', null], ['2', null], ['3', null], ['4', null]],
];

function slide08() {
  const s = newSlide(8);
  mapleLeaf(s, 0.362, 0.379, 0.972, 1.156, -60);
  text(s, 'Agenda', { x: 0.579, y: 0.462, w: 2.372, h: 0.909, fontFace: SCRIPT, fontSize: 50 });

  // calendar card
  s.addShape('roundRect', {
    x: 0.587, y: 1.874, w: 3.214, h: 3.297, rectRadius: 0.205,
    fill: { color: C.white }, line: { color: C.amber, width: 3 },
  });
  s.addShape('roundRect', {
    x: 1.231, y: 1.877, w: 1.937, h: 0.383, rectRadius: 0.1, fill: { color: C.amber }, line: { type: 'none' },
  });
  text(s, 'September 2025', {
    x: 1.231, y: 1.877, w: 1.937, h: 0.383, fontSize: 12, bold: true, color: C.white, align: 'center', valign: 'middle',
  });
  ['M', 'T', 'W', 'T', 'F', 'S', 'S'].forEach((d, i) => {
    text(s, d, {
      x: CAL_COLS[i], y: 2.407, w: 0.436, h: 0.383, fontSize: 9, bold: true, color: C.body,
      align: 'center', valign: 'middle',
    });
  });
  CAL_GRID.forEach((row, r) => row.forEach(([label, fill], c) => {
    s.addShape('roundRect', {
      x: CAL_COLS[c], y: CAL_ROWS[r], w: 0.436, h: 0.383, rectRadius: 0.064,
      fill: fill ? { color: fill } : { type: 'none' }, line: { color: C.white, width: 0.75 },
    });
    text(s, label, {
      x: CAL_COLS[c], y: CAL_ROWS[r], w: 0.436, h: 0.383, fontSize: 9, color: C.grey,
      align: 'center', valign: 'middle',
    });
  }));

  panel(s, 5.0, 0, 5.0, 5.625, C.cream);
  mapleLeaf(s, 3.521, 1.686, 0.28, 0.393, 15);
  mapleLeaf(s, 0.778, 2.242, 0.207, 0.29, 15);
  text(s, '“ ' + LOREM + 'consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus “', {
    x: 5.403, y: 0.631, w: 4.022, h: 0.647, fontSize: 9, italic: true, color: C.body, lineSpacingMultiple: 1.3,
  });
  pictogram(s, 'cloche', 5.466, 3.256, 0.273, 0.216);
  captionBlock(s, 5.335, 3.754, 1.485, 'Gathering', LOREM + 'consectetuer adipiscing elit. ', { gap: 0.377 });
  pictogram(s, 'camera', 7.753, 3.286, 0.265, 0.237);
  captionBlock(s, 7.664, 3.754, 1.485, 'Vacation', LOREM + 'consectetuer adipiscing elit. ', { gap: 0.377 });
}

/* -- 9 · Schedule -------------------------------------------------------- */
const SCHEDULE = [
  ['Photography', 1.952, 1.109], ['Art & Craft', 1.949, 3.482],
  ['Bonfire & Storytelling', 7.534, 1.109], ['Movie Under the Stars', 7.532, 3.482],
];

function slide09() {
  const s = newSlide(9);
  mapleLeaf(s, 0.435, 3.09, 1.23, 1.34, -135);
  panel(s, 5.0, 0, 5.0, 5.625, C.cream);
  text(s, 'Schedule', { x: -0.307, y: 2.198, w: 2.69, h: 0.917, fontFace: SCRIPT, fontSize: 50, rotate: 270 });
  SCHEDULE.forEach(([head, x, y]) => {
    captionBlock(s, x, y, head === 'Movie Under the Stars' ? 2.148 : 1.901, head,
      LOREM + 'consectetur adipiscing elit, sed do eiusmod tempor', { headW: 2.148, gap: 0.386 });
  });
  photo(s, 4.117, 0, 3.033, 5.625, 'rect');
}

/* -- 10 · Gallery -------------------------------------------------------- */
// x, width, and the wash layers stacked over the placeholder (the third
// column carries an opaque accent wash with the soft one painted on top).
const GALLERY_COLS = [
  [0.0, 1.726, [WASH_SOFT]],
  [1.726, 1.726, [WASH_SOFT]],
  [3.45, 1.727, [WASH_ACCENT, WASH_SOFT]],
  [7.653, 2.347, [WASH_SOFT]],
];

function slide10() {
  const s = newSlide(10);
  panel(s, 5.177, 0, 2.476, 5.625, C.cream);
  GALLERY_COLS.forEach(([x, w, layers]) => {
    washBox(s, x, 0, w, 5.625, 'rect', layers, PHOTO);
    s.addText('[image]', { x: x, y: 2.65, w: w, h: 0.32, align: 'center', margin: 0, fontFace: SANS, fontSize: 9, color: C.body });
  });
  mapleLeaf(s, 5.329, 0.328, 1.23, 1.34, -60);
  text(s, 'Gallery', { x: 5.717, y: 0.462, w: 2.372, h: 0.909, fontFace: SCRIPT, fontSize: 50 });
  text(s, 'Best Gallery Autumn Festival 2025', { x: 5.758, y: 1.332, w: 1.902, h: 0.48, fontSize: 12, bold: true });
  text(s, '“ ' + LOREM + 'consectetur adipiscing elit, sed do eiusmod tempor ”', {
    x: 5.877, y: 4.512, w: 1.777, h: 0.647, fontSize: 9, italic: true, color: C.body, lineSpacingMultiple: 1.3,
  });
}

/* -- 11 · Best Moment Autumn --------------------------------------------- */
function slide11() {
  const s = newSlide(11);
  mapleLeaf(s, 5.045, 0.246, 1.23, 1.34, -45);
  panel(s, 0, 2.812, 10.0, 2.812, C.cream);
  text(s, 'Best Moment\nAutumn', { x: 5.421, y: 0.462, w: 4.009, h: 1.742, fontFace: SCRIPT, fontSize: 50 });
  [5.393, 7.523].forEach(x => text(s, LOREM + 'consectetur adipiscing elit, sed', {
    x: x, y: 4.724, w: 1.935, h: 0.451, fontSize: 9, color: C.body, lineSpacingMultiple: 1.3,
  }));
  divider(s, 5.0, 0, 0, 2.813);
  photo(s, 0.57, 0.451, 2.089, 3.556, 'pill');
  photo(s, 2.015, 1.619, 2.089, 3.556, 'pill');
  photo(s, 5.421, 3.116, 1.907, 1.291, 'hpill', { labelSize: 7 });
  photo(s, 7.523, 3.116, 1.907, 1.291, 'hpill', { labelSize: 7 });
}

/* -- 12 · Before the Event ------------------------------------------------ */
function slide12() {
  const s = newSlide(12);
  mapleLeaf(s, 5.663, 0.322, 1.23, 1.34, -60);
  panel(s, 0, 0, 5.791, 5.625, C.cream);
  text(s, 'Before the\nEvent', { x: 6.057, y: 0.458, w: 3.37, h: 1.742, fontFace: SCRIPT, fontSize: 50 });
  statBlock(s, 5.883, 4.206, 1.509, '112+', 'Our Participant',
    { labelSize: 14, labelBold: true, labelColor: C.ink, labelH: 0.343 });
  statBlock(s, 7.754, 4.206, 1.509, '87%', 'Total Regristation',
    { labelSize: 14, labelBold: true, labelColor: C.ink, labelH: 0.343, labelW: 1.813 });
  divider(s, 5.791, 2.812, 4.209, 0);
  photo(s, 0.57, 0.843, 2.343, 3.94, 'pill');
  photo(s, 3.24, 0.843, 2.343, 3.94, 'pill');
}

/* -- 13 · Main Event ------------------------------------------------------ */
function slide13() {
  const s = newSlide(13);
  mapleLeaf(s, 0.419, 2.521, 1.23, 1.34, -135);
  text(s, 'Main Event', { x: -0.516, y: 1.537, w: 3.1, h: 0.909, fontFace: SCRIPT, fontSize: 50, rotate: 270 });
  panel(s, 6.047, 0, 3.953, 5.625, C.cream);
  captionBlock(s, 6.788, 1.333, 2.503, 'Best of Time',
    LOREM + 'consectetur adipiscing elit, sed do eiusmod tempor', { bodyH: 0.451, gap: 0.386 });
  text(s, '2025', { x: 1.527, y: 4.035, w: 1.437, h: 0.581, fontSize: 30, bold: true });
  text(s, LOREM + 'consectetur adipiscing elit, sed do', {
    x: 2.612, y: 4.1, w: 2.388, h: 0.451, fontSize: 9, color: C.body, lineSpacingMultiple: 1.3,
  });
  photo(s, 1.489, 0.451, 1.212, 3.1, 'pill', { labelSize: 7 });
  photo(s, 2.839, 0.451, 2.959, 1.212, 'hpill');
  photo(s, 2.839, 2.339, 2.959, 1.212, 'hpill');
  photo(s, 6.79, 3.456, 1.018, 1.645, 'pill', { labelSize: 7 });
  photo(s, 8.278, 3.456, 1.018, 1.645, 'pill', { labelSize: 7 });
}

/* -- 14 · Our Leader ------------------------------------------------------ */
const SOCIAL = [['X', 1.585], ['f', 1.983], ['\u25A2', 2.422], ['in', 2.805]];

function slide14() {
  const s = newSlide(14);
  mapleLeaf(s, 0.47, 0.239, 1.23, 1.34, -45);
  panel(s, 5.0, 0, 5.0, 5.625, C.cream);
  text(s, 'Our Leader', { x: 0.839, y: 0.46, w: 3.526, h: 0.917, fontFace: SCRIPT, fontSize: 50 });
  captionBlock(s, 0.891, 1.999, 3.048, 'Biography',
    LOREM + 'consectetuer adipiscing elit. Maecenas porttitor congue massa.', { bodyH: 0.451, gap: 0.386 });
  s.addText([0, 1, 2, 3].map(() => ({
    text: LOREM + 'consectetuer',
    options: { bullet: { characterCode: '2022', indent: 10 }, breakLine: true },
  })), {
    x: 0.951, y: 3.004, w: 3.048, h: 1.041, fontFace: SANS, fontSize: 9, color: C.body,
    lineSpacingMultiple: 1.3, margin: INSET, valign: 'top',
  });
  text(s, '7+ /Years Experience', { x: 0.893, y: 4.676, w: 1.873, h: 0.278, fontSize: 12, bold: true });
  text(s, '15+ Creating Event', { x: 2.885, y: 4.676, w: 1.714, h: 0.278, fontSize: 12, bold: true });
  text(s, 'Ella Nelson', { x: 6.063, y: 4.638, w: 1.899, h: 0.278, fontSize: 12, bold: true });
  text(s, 'Founder', { x: 6.061, y: 4.916, w: 1.901, h: 0.254, fontSize: 9, color: C.muted, lineSpacingMultiple: 1.3 });
  SOCIAL.forEach(([glyph, y]) => text(s, glyph, {
    x: 8.17, y: y - 0.03, w: 0.34, h: 0.3, fontSize: 13, bold: true, color: C.amber, align: 'center', valign: 'middle',
  }));
  photo(s, 6.061, 1.075, 2.085, 3.461, 'pill');
}

/* -- 15 · Team Profile ---------------------------------------------------- */
const TEAM = [
  ['Olivia Garcia', 'Sponsorship Coordinator', 1.2, 1.092, 1.203, 3.762],
  ['Michael Smith', 'Marketing Specialist', 4.183, 2.883, 4.183, 2.032],
  ['Amelia Young', 'Festival Coordinator', 7.167, 1.092, 7.167, 3.762],
];

function slide15() {
  const s = newSlide(15);
  mapleLeaf(s, 2.762, 0.239, 1.23, 1.34, -45);
  text(s, 'Team Profile', { x: 3.196, y: 0.46, w: 3.625, h: 0.909, fontFace: SCRIPT, fontSize: 50 });
  TEAM.forEach(([name, role, px, py, tx, ty]) => {
    photo(s, px, py, 1.633, 2.292, 'pill');
    text(s, name, { x: tx, y: ty, w: 1.902, h: 0.278, fontSize: 12, bold: true });
    text(s, role, { x: tx, y: ty + 0.278, w: 1.902, h: 0.254, fontSize: 9, color: C.body, lineSpacingMultiple: 1.3 });
  });
}

/* -- 16 · Price List ------------------------------------------------------ */
const PRICING = [
  { price: '$ 120', name: 'Autumn Breeze', x: 3.306, w: 2.212, h: 3.932, fill: C.amber, tx: 3.461, px: 3.7, ticks: 5 },
  { price: '$ 90', name: 'Golden Entry', x: 5.6, w: 1.875, h: 3.183, fill: C.mist, tx: 5.674, px: 5.825, ticks: 2 },
  { price: '$ 60', name: 'Maple Entry', x: 7.556, w: 1.875, h: 3.183, fill: C.mist, tx: 7.638, px: 7.781, ticks: 2 },
];

function slide16() {
  const s = newSlide(16);
  mapleLeaf(s, 0.12, 0.254, 1.23, 1.34, -45);
  text(s, 'Price  List\nAutumn', { x: 0.579, y: 0.46, w: 3.863, h: 1.742, fontFace: SCRIPT, fontSize: 50 });
  text(s, LOREM + 'consectetur adipiscing elit, sed do', {
    x: 0.571, y: 2.326, w: 2.674, h: 0.451, fontSize: 9, color: C.body, lineSpacingMultiple: 1.3,
  });
  PRICING.forEach(p => {
    const y = p.h > 3.5 ? 1.233 : (p.x > 7 ? 1.221 : 1.233);
    s.addShape('roundRect', { x: p.x, y: y, w: p.w, h: p.h, rectRadius: 0.3125, fill: { color: p.fill }, line: { type: 'none' } });
    text(s, p.price, { x: p.px, y: 1.621, w: 1.425, h: 0.581, fontSize: 30, bold: true });
    text(s, p.name, { x: p.tx, y: 2.451, w: 1.902, h: 0.278, fontSize: 12, bold: true });
    text(s, LOREM + 'consectetuer adipiscing elit. ', {
      x: p.tx, y: p.ticks > 2 ? 2.779 : 2.812, w: 1.902, h: 0.451, fontSize: 9, color: C.body, lineSpacingMultiple: 1.3,
    });
    text(s, 'Package :', { x: p.tx, y: p.ticks > 2 ? 3.281 : 3.28, w: 1.902, h: 0.278, fontSize: 12, bold: true });
    const line = p.ticks > 2 ? 'Lorem ipsum dolor sit amet' : 'Lorem ipsum dolor sit';
    s.addText(Array.from({ length: p.ticks }, () => ({
      text: line, options: { bullet: { characterCode: '2714', indent: 10 }, breakLine: true },
    })), {
      x: p.tx, y: 3.608, w: 1.902, h: p.ticks * 0.19, fontFace: SANS, fontSize: 9, color: C.body,
      lineSpacingMultiple: 1.3, margin: INSET, valign: 'top',
    });
  });
}

/* -- 17 · Autumn Data Chart (pie) ----------------------------------------- */
// Three wedges drawn as `pie` shapes, exactly as in the source deck.
const PIE_WEDGES = [
  [C.amber, 269.6, 141.1],
  [C.silver, 140.5, 231.0],
  [C.grey, 228.7, 270.9],
];
const CHART_LEGEND = [
  ['Data One', C.amber, 2.733],
  ['Data Two', C.silver, 3.559],
  ['Data Three', C.grey, 4.407],
];

function slide17() {
  const s = newSlide(17);
  panel(s, 5.0, 0, 5.0, 5.625, C.cream);
  mapleLeaf(s, 5.008, 0.497, 1.23, 1.34, -45);
  text(s, 'Auntumn Data\nChart', { x: 5.383, y: 0.714, w: 4.05, h: 1.742, fontFace: SCRIPT, fontSize: 50 });
  PIE_WEDGES.forEach(([color, from, to]) => {
    s.addShape('pie', { x: 1.148, y: 0.58, w: 3.298, h: 3.298, angleRange: [from, to], fill: { color: color }, line: { type: 'none' } });
  });
  text(s, LOREM + 'consectetuer adipiscing elit. Maecenas porttitor congue massa.', {
    x: 1.27, y: 4.725, w: 3.048, h: 0.451, fontSize: 9, italic: true, color: C.body, lineSpacingMultiple: 1.3,
  });
  CHART_LEGEND.forEach(([label, color, y]) => {
    pictogram(s, 'chevron', 5.438, y + 0.03, 0.11, 0.16);
    text(s, label, { x: 5.572, y: y, w: 2.228, h: 0.278, fontSize: 12, bold: true, color: color });
    text(s, LOREM + 'consectetur', { x: 5.438, y: y + 0.286, w: 2.557, h: 0.254, fontSize: 9, color: C.body, lineSpacingMultiple: 1.3 });
  });
}

/* -- 18 · Step Planning --------------------------------------------------- */
const STEP_CARDS = [                       // back-to-front stacking order
  { x: 1.173, y: 2.32, fill: C.silver, num: '04.', nx: 1.943, ny: 2.867, color: C.body },
  { x: 1.77, y: 3.247, fill: C.amber, num: '01.', nx: 2.475, ny: 3.836, color: C.rust },
  { x: 3.444, y: 2.979, fill: C.body, num: '02.', nx: 4.191, ny: 3.738, color: C.mist },
  { x: 2.847, y: 1.846, fill: C.peach, num: '03.', nx: 3.621, ny: 2.442, color: C.amber },
];
const STEPS = [['01.', 'Planning One', 2.006], ['02.', 'Planning Two', 2.791],
               ['03.', 'Planning Three', 3.59], ['04.', 'Planning Four', 4.375]];

function slide18() {
  const s = newSlide(18);
  panel(s, 6.283, 0, 3.717, 5.625, C.cream);
  mapleLeaf(s, 0.191, 0.239, 1.23, 1.34, -45);
  STEP_CARDS.forEach(card => {
    s.addShape('roundRect', {
      x: card.x, y: card.y, w: 2.153, h: 2.122, rectRadius: 0.3537,
      fill: { color: card.fill }, line: { type: 'none' },
    });
  });
  STEP_CARDS.forEach(card => {
    text(s, card.num, { x: card.nx, y: card.ny, w: 0.85, h: 0.656, fontSize: 27, bold: true, color: card.color });
  });
  text(s, 'Step Planning\nAutumn Festival', { x: 0.575, y: 0.456, w: 4.425, h: 1.742, fontFace: SCRIPT, fontSize: 50 });
  STEPS.forEach(([num, head, y]) => {
    text(s, num, { x: 6.508, y: y + 0.003, w: 0.63, h: 0.48, fontSize: 24, bold: true });
    captionBlock(s, 7.102, y, 2.329, head, LOREM + 'consectetuer adipiscing elit', { bodyH: 0.451, headW: 1.902 });
  });
}

/* -- 19 · Mapping Event (Canada map) -------------------------------------- */
// Each entry is [fill, polygon] with coordinates in inches relative to the
// map's top-left corner (MAP_ORIGIN).
const MAP_ORIGIN = [4.474, 0.63];
const CANADA = [
  ['BDBDEF',[[3.493,4.376],[3.965,3.863],[4.251,3.687],[4.256,3.423],[4.355,3.382],[4.522,3.567],[4.756,3.389],[4.658,3.647],[4.766,3.594],[4.91,3.312],[4.629,3.357],[4.535,3.205],[4.412,3.26],[4.482,3.104],[4.331,3.173],[4.164,3.517],[4.283,3.054],[4.666,2.818],[4.771,2.452],[4.536,2.367],[4.461,2.533],[4.518,2.332],[4.173,2.336],[3.823,2.056],[3.796,2.34],[3.669,2.353],[3.56,2.121],[3.117,2.163],[3.241,2.414],[3.198,2.547],[3.398,2.685],[3.292,2.968],[3.454,3.209],[3.464,3.314],[3.37,3.338],[3.175,3.183],[3.114,2.992],[2.488,2.918],[2.33,2.722],[2.373,2.37],[2.495,2.256],[2.331,2.17],[2.487,2.192],[2.592,2.093],[2.617,2.009],[2.467,1.95],[2.601,1.979],[2.595,1.862],[2.762,1.823],[2.746,1.542],[2.593,1.522],[2.557,1.79],[2.212,1.349],[2.169,1.528],[2.296,1.606],[2.231,1.83],[2.203,1.733],[1.974,1.803],[1.803,1.655],[1.697,1.694],[1.813,1.706],[1.697,1.861],[1.642,1.742],[1.445,1.715],[1.456,1.577],[1.233,1.417],[1.15,1.46],[1.155,1.367],[1.085,1.42],[1.065,1.268],[0.823,1.358],[1.003,1.28],[0.726,1.306],[0.609,1.105],[0.02,1.905],[0.086,2.126],[0.227,2.149],[0.294,2.624],[0.173,2.843],[0.198,3.177],[0.411,3.499],[1.461,3.768],[2.752,3.845],[3.005,3.711],[3.22,3.919],[3.434,3.919],[3.316,3.967],[3.494,3.908],[3.587,3.975],[3.585,4.061],[3.473,4.006],[3.493,4.376]]],
  ['FECB86',[[2.74,2.199],[2.802,2.06],[2.898,2.112],[2.96,2.066],[2.726,1.949],[2.705,1.975],[2.666,1.917],[2.637,2.172],[2.696,2.144],[2.74,2.199]]],
  ['BFBFBF',[[3.045,1.963],[3.207,1.886],[3.386,1.982],[3.661,1.958],[3.45,1.863],[3.665,1.859],[3.628,1.744],[3.37,1.631],[3.398,1.555],[3.608,1.619],[3.556,1.424],[3.203,1.395],[3.191,1.236],[2.917,1.223],[2.821,1.128],[2.66,1.231],[2.6,1.062],[2.504,1.129],[2.577,1.365],[2.515,1.389],[2.452,1.082],[2.374,1.304],[2.483,1.384],[2.409,1.402],[2.451,1.463],[2.702,1.46],[2.781,1.36],[3.145,1.58],[3.167,1.792],[3.022,1.864],[3.044,1.962]]],
  ['FECB86',[[1.579,1.657],[1.749,1.638],[1.778,1.585],[1.825,1.653],[1.92,1.671],[1.964,1.623],[1.912,1.584],[2.011,1.578],[1.876,1.456],[1.829,1.18],[1.776,1.192],[1.785,1.347],[1.742,1.217],[1.585,1.212],[1.587,1.138],[1.443,1.197],[1.406,1.258],[1.422,1.307],[1.506,1.325],[1.404,1.334],[1.407,1.381],[1.552,1.41],[1.613,1.479],[1.416,1.451],[1.403,1.511],[1.579,1.657]]],
  ['BFBFBF',[[1.255,1.285],[1.33,1.274],[1.437,1.157],[1.548,1.109],[1.517,1.018],[1.472,1.034],[1.417,0.948],[1.331,0.955],[1.286,1.084],[1.203,1.151],[1.255,1.285]]],
  ['BFBFBF',[[2.207,1.324],[2.239,1.309],[2.226,1.24],[2.303,1.202],[2.299,1.104],[2.209,1.099],[2.173,1.139],[2.207,1.324]]],
  ['BFBFBF',[[2.08,1.405],[2.158,1.269],[2.08,1.225],[2.124,1.125],[2.013,1.154],[2.013,1.257],[1.935,1.242],[2.08,1.405]]],
  ['BFBFBF',[[2.366,0.995],[2.374,1.03],[2.476,1.004],[2.519,0.955],[2.556,0.993],[2.594,0.868],[2.533,0.836],[2.347,0.919],[2.287,0.874],[2.301,0.835],[2.12,0.787],[2.165,0.861],[2.235,0.874],[2.253,0.979],[2.322,1.027],[2.366,0.995]]],
  ['BFBFBF',[[1.644,1.023],[1.865,1.004],[1.883,0.908],[1.815,0.875],[1.822,0.785],[1.769,0.827],[1.793,0.926],[1.731,0.92],[1.675,0.807],[1.605,0.84],[1.585,0.883],[1.626,0.899],[1.555,0.923],[1.617,0.972],[1.649,0.928],[1.713,0.967],[1.644,1.023]]],
  ['BFBFBF',[[4.865,2.996],[5.033,2.899],[5.052,2.837],[5.098,2.853],[5.104,2.81],[5.135,2.823],[5.111,2.907],[5.141,2.9],[5.164,2.753],[5.239,2.825],[5.249,2.789],[5.295,2.805],[5.295,2.751],[5.212,2.673],[5.178,2.711],[5.166,2.633],[5.124,2.659],[5.092,2.595],[5.017,2.675],[4.902,2.651],[4.893,2.724],[4.836,2.512],[4.796,2.523],[4.854,2.881],[4.821,2.917],[4.873,2.906],[4.865,2.996]]],
  ['BFBFBF',[[2.308,0.809],[2.439,0.82],[2.52,0.76],[2.548,0.812],[2.597,0.752],[2.593,0.691],[2.471,0.681],[2.551,0.653],[2.586,0.537],[2.577,0.476],[2.493,0.469],[2.561,0.397],[2.549,0.357],[2.598,0.346],[2.611,0.125],[2.553,0.175],[2.61,0.036],[2.584,0.003],[2.373,0.034],[2.296,0.128],[2.348,0.179],[2.277,0.154],[2.189,0.24],[2.218,0.315],[2.26,0.301],[2.235,0.341],[2.307,0.367],[2.404,0.282],[2.438,0.298],[2.369,0.376],[2.404,0.43],[2.299,0.391],[2.398,0.53],[2.344,0.555],[2.323,0.628],[2.403,0.615],[2.392,0.67],[2.447,0.686],[2.335,0.659],[2.308,0.809]]],
  ['FECB86',[[0.269,3.44],[0.274,3.496],[0.361,3.548],[0.287,3.312],[0.138,3.175],[0.227,3.435],[0.269,3.44]]],
  ['BFBFBF',[[2.237,0.641],[2.302,0.613],[2.335,0.506],[2.156,0.282],[2.112,0.356],[2.143,0.382],[2.117,0.492],[2.142,0.527],[2.211,0.516],[2.165,0.553],[2.237,0.641]]],
  ['FECB86',[[0.026,2.948],[0.025,2.805],[0.082,2.756],[0.038,2.692],[0,2.815],[0.026,2.948]]],
  ['FECB86',[[0.086,1.962],[0.831,2.383],[1.49,2.559],[2.324,2.582],[2.494,2.255],[2.332,2.163],[2.495,2.189],[2.616,2.009],[2.466,1.95],[2.6,1.979],[2.594,1.861],[2.696,1.88],[2.808,1.728],[2.665,1.492],[2.542,1.781],[2.488,1.608],[2.42,1.678],[2.211,1.349],[2.167,1.528],[2.295,1.606],[2.23,1.83],[2.202,1.733],[1.973,1.803],[1.802,1.655],[1.696,1.694],[1.812,1.706],[1.696,1.861],[1.641,1.741],[1.444,1.715],[1.455,1.577],[1.232,1.417],[1.149,1.46],[1.153,1.367],[1.084,1.42],[1.064,1.268],[0.822,1.358],[1.002,1.28],[0.725,1.306],[0.608,1.104],[0.02,1.905],[0.086,1.962]]],
  ['BFBFBF',[[4.714,3.652],[4.909,3.312],[4.628,3.357],[4.535,3.205],[4.412,3.26],[4.482,3.104],[4.331,3.173],[4.163,3.517],[4.283,3.054],[4.666,2.818],[4.779,2.463],[4.535,2.367],[4.461,2.533],[4.517,2.332],[4.173,2.336],[3.822,2.056],[3.796,2.34],[3.669,2.353],[3.559,2.121],[3.303,2.091],[3.108,2.175],[3.241,2.414],[3.198,2.547],[3.398,2.685],[3.292,2.968],[3.454,3.209],[3.393,3.345],[3.199,3.216],[3.114,2.992],[2.477,2.914],[2.324,2.582],[1.491,2.559],[0.832,2.383],[0.087,1.962],[0.087,2.126],[0.221,2.128],[0.294,2.6],[0.174,2.843],[0.199,3.177],[0.522,3.543],[1.462,3.768],[2.752,3.845],[3.006,3.711],[3.22,3.919],[3.435,3.919],[3.316,3.967],[3.494,3.908],[3.587,3.975],[3.574,4.063],[3.477,3.996],[3.466,4.382],[3.735,4.198],[3.678,4.136],[4.004,3.829],[4.196,3.756],[4.257,3.423],[4.356,3.382],[4.523,3.567],[4.757,3.389],[4.646,3.579],[4.715,3.652]]],
  ['BFBFBF',[[4.056,2.779],[4.133,2.797],[4.16,2.864],[4.318,2.86],[4.302,2.73],[4.354,2.785],[4.694,2.544],[4.741,2.572],[4.779,2.464],[4.673,2.374],[4.608,2.374],[4.587,2.441],[4.579,2.374],[4.535,2.367],[4.461,2.533],[4.436,2.495],[4.517,2.332],[4.321,2.386],[4.289,2.34],[4.237,2.36],[4.163,2.328],[4.162,2.261],[3.859,2.097],[4.196,2.503],[4.181,2.555],[3.983,2.606],[4.057,2.78]]],
  ['FECB86',[[2.683,2.897],[2.476,2.913],[2.324,2.581],[1.739,2.587],[1.764,3.787],[2.473,3.836],[2.786,2.948],[2.683,2.897]]],
  ['BFBFBF',[[3.967,3.768],[3.644,3.83],[3.41,3.282],[3.393,3.345],[3.218,3.238],[3.113,2.992],[2.645,2.894],[2.36,3.337],[2.357,3.792],[2.752,3.845],[3.005,3.711],[3.22,3.919],[3.434,3.919],[3.316,3.967],[3.494,3.908],[3.587,3.975],[3.573,4.063],[3.476,3.996],[3.466,4.382],[3.735,4.198],[3.678,4.136],[3.94,3.935],[3.967,3.767]]],
  ['FECB86',[[1.963,2.597],[1.269,2.517],[0.719,2.337],[0.086,1.961],[0.086,2.126],[0.22,2.128],[0.293,2.6],[0.173,2.843],[0.198,3.177],[0.521,3.543],[1.142,3.714],[1.963,3.791],[1.963,2.597]]],
  ['FECB86',[[1.511,2.562],[0.842,2.387],[0.576,3.561],[1.292,3.744],[1.511,2.562]]],
  ['FECB86',[[0.086,1.962],[0.086,2.126],[0.22,2.128],[0.293,2.6],[0.173,2.843],[0.198,3.177],[0.41,3.499],[1.003,3.681],[0.783,3.061],[0.98,2.437],[0.086,1.961]]],
  ['FECB86',[[1.963,2.167],[1.591,2.052],[1.452,1.914],[1.234,1.604],[1.272,1.457],[1.149,1.46],[1.154,1.367],[1.084,1.42],[1.064,1.268],[0.822,1.358],[1.002,1.28],[0.725,1.306],[0.608,1.104],[0.02,1.905],[0.719,2.337],[1.269,2.517],[1.963,2.597],[1.963,2.167]]],
  ['FECB86',[[0.788,2.234],[0.624,2.064],[0.685,1.605],[0.61,1.542],[0.691,1.455],[0.608,1.406],[0.693,1.241],[0.608,1.104],[0.02,1.905],[0.796,2.369],[0.788,2.234]]],
  ['FECB86',[[1.731,1.207],[1.706,1.239],[1.614,1.19],[1.597,1.218],[1.574,1.137],[1.437,1.201],[1.419,1.303],[1.506,1.325],[1.404,1.334],[1.397,1.367],[1.552,1.41],[1.613,1.479],[1.398,1.458],[1.692,1.529],[1.732,1.207]]],
  ['BFBFBF',[[4.9,3.306],[4.826,3.284],[4.754,3.349],[4.62,3.356],[4.53,3.296],[4.535,3.205],[4.318,3.314],[4.276,3.429],[4.355,3.382],[4.522,3.568],[4.652,3.407],[4.67,3.447],[4.757,3.39],[4.646,3.579],[4.658,3.647],[4.723,3.653],[4.772,3.583],[4.757,3.499],[4.813,3.473],[4.901,3.306]]],
  ['BFBFBF',[[4.885,3.352],[4.909,3.312],[4.826,3.284],[4.754,3.349],[4.676,3.342],[4.666,3.447],[4.76,3.396],[4.68,3.494],[4.658,3.647],[4.729,3.645],[4.772,3.583],[4.757,3.499],[4.885,3.352]]],
  ['BFBFBF',[[1.731,0.92],[1.666,0.805],[1.584,0.877],[1.626,0.899],[1.554,0.913],[1.616,0.972],[1.649,0.928],[1.713,0.967],[1.649,0.989],[1.657,1.027],[1.755,0.992],[1.732,0.92]]],
];

function slide19() {
  const s = newSlide(19);
  mapleLeaf(s, 0.229, 0.345, 1.23, 1.34, -45);
  CANADA.forEach(([fill, poly]) => {
    const pts = poly.map(([px, py], i) => ({ x: +(MAP_ORIGIN[0] + px).toFixed(3), y: +(MAP_ORIGIN[1] + py).toFixed(3), moveTo: i === 0 }));
    pts.push({ close: true });
    s.addShape('custGeom', {
      x: 0, y: 0, w: 10, h: 5.625, points: pts,
      fill: { color: fill }, line: { color: C.white, width: 0.75 },
    });
  });
  pictogram(s, 'pin', 5.526, 2.727, 0.108, 0.172, C.cream);
  pictogram(s, 'pin', 8.248, 3.573, 0.108, 0.172, C.mist);
  text(s, 'Mapping Event\nAutumn Fest', { x: 0.577, y: 0.462, w: 4.05, h: 1.742, fontFace: SCRIPT, fontSize: 50 });
  text(s, LOREM_LONG, { x: 0.579, y: 2.326, w: 3.207, h: 0.647, fontSize: 9, color: C.body, lineSpacingMultiple: 1.3 });
  [['80%', 'Fest Data One', 3.45, 3.35], ['78%', 'Fest Data Two', 4.347, 4.265]].forEach(([v, head, vy, hy]) => {
    text(s, v, { x: 0.58, y: vy, w: 1.035, h: 0.631, fontSize: 33, bold: true });
    captionBlock(s, 1.617, hy, 2.637, head, LOREM + 'consectetuer adipiscing elit. Maecenas porttitor congue', { bodyH: 0.451 });
  });
}

/* -- 20 · Let's Stream on Mobile App -------------------------------------- */

// Rotating a mock-up means rotating every part about the *group* centre, so
// this returns a placement function that does exactly that for sub-rectangles.
function rotator(cx, cy, deg) {
  const a = deg * Math.PI / 180, cs = Math.cos(a), sn = Math.sin(a);
  return (x, y, w, h) => {
    const dx = x + w / 2 - cx, dy = y + h / 2 - cy;
    return { x: cx + dx * cs - dy * sn - w / 2, y: cy + dx * sn + dy * cs - h / 2, w: w, h: h, rotate: deg };
  };
}

// Phone mock-up: pale rim, black bezel, light screen (the image placeholder)
// and the notch. `deg` tilts the whole device about its centre.
function phoneMock(slide, x, y, w, h, deg, labelSize) {
  const at = rotator(x + w / 2, y + h / 2, deg || 0);
  const r = w * 0.145, rim = 0.028, bez = 0.104, none = { type: 'none' };
  slide.addShape('roundRect', Object.assign(at(x, y, w, h), { rectRadius: r, fill: { color: '7FAFCC' }, line: none }));
  slide.addShape('roundRect', Object.assign(at(x + rim, y + rim, w - 2 * rim, h - 2 * rim),
    { rectRadius: r - rim, fill: { color: '000000' }, line: none }));
  slide.addShape('roundRect', Object.assign(at(x + bez, y + bez, w - 2 * bez, h - 2 * bez),
    { rectRadius: r - bez, fill: { color: PHOTO }, line: none }));
  slide.addShape('roundRect', Object.assign(at(x + w * 0.315, y + rim, w * 0.37, 0.25),
    { rectRadius: 0.085, fill: { color: '000000' }, line: none }));
  slide.addText('[image]', Object.assign(at(x, y + h / 2 - 0.16, w, 0.32), {
    margin: 0, align: 'center', valign: 'middle', fontFace: SANS, fontSize: labelSize || 9, color: C.grey,
  }));
}

function slide20() {
  const s = newSlide(20);
  panel(s, 0, 0, 10.0, 2.812, C.cream);
  phoneMock(s, 6.24, 1.36, 1.86, 3.72, -15, 7);
  mapleLeaf(s, 0.191, 0.239, 1.23, 1.34, -45);
  text(s, 'Let’s Stream\non Mobile App', { x: 0.577, y: 0.462, w: 4.05, h: 1.742, fontFace: SCRIPT, fontSize: 50 });
  statBlock(s, 0.577, 4.206, 1.509, '345k+', 'Total Stream',
    { labelSize: 14, labelBold: true, labelColor: C.ink, labelH: 0.343 });
  statBlock(s, 2.787, 4.206, 1.509, '2,5M+', 'Total Restream',
    { labelSize: 14, labelBold: true, labelColor: C.ink, labelH: 0.343 });
  phoneMock(s, 7.194, 0.549, 2.292, 4.6, 0, 12);
}

/* -- 21 · Check Event on Desktop App -------------------------------------- */
// Laptop mock-up: thin dark lid frame around a light screen (the image
// placeholder), then the base bar with its lighter foot and trackpad notch.
function laptopMock(slide, x, y, w, lidH) {
  const none = { type: 'none' }, baseY = y + lidH;
  slide.addShape('roundRect', { x: x, y: y, w: w, h: lidH, rectRadius: 0.07, fill: { color: '3A3A3C' }, line: none });
  slide.addShape('rect', { x: x + 0.028, y: y + 0.028, w: w - 0.056, h: lidH - 0.056, fill: { color: PHOTO }, line: none });
  slide.addShape('roundRect', { x: x - 0.12, y: baseY, w: w + 0.14, h: 0.315, rectRadius: 0.04, fill: { color: '1B1B1B' }, line: none });
  slide.addShape('roundRect', { x: x - 0.12, y: baseY + 0.11, w: w + 0.54, h: 0.135, rectRadius: 0.065, fill: { color: '767780' }, line: none });
  slide.addShape('rect', { x: x + w * 0.475, y: baseY + 0.11, w: 0.72, h: 0.09, fill: { color: '454545' }, line: none });
  slide.addText('[image]', {
    x: x, y: y + lidH / 2 - 0.16, w: w, h: 0.32, margin: 0, align: 'center', valign: 'middle',
    fontFace: SANS, fontSize: 12, color: C.grey,
  });
}

function slide21() {
  const s = newSlide(21);
  panel(s, 5.0, 0, 5.0, 3.658, C.cream);
  mapleLeaf(s, 5.008, 0.272, 1.23, 1.34, -45);
  laptopMock(s, -1.9, 1.233, 6.12, 3.62);
  text(s, 'Check Event on\nDekstop App', { x: 5.386, y: 0.462, w: 4.05, h: 1.742, fontFace: SCRIPT, fontSize: 50 });
  text(s, LOREM_LONG, { x: 5.381, y: 2.424, w: 3.207, h: 0.647, fontSize: 9, color: C.body, lineSpacingMultiple: 1.3 });
  divider(s, 5.008, 3.658, 0, 1.967);
  pictogram(s, 'inbox', 5.066, 4.095, 0.337, 0.29);
  pictogram(s, 'map', 7.411, 4.095, 0.361, 0.29);
  text(s, 'You Can Download App in Here', { x: 5.006, y: 4.533, w: 1.868, h: 0.638, fontSize: 14, bold: true, lineSpacingMultiple: 1.3 });
  text(s, 'You Can Check Location Event on App', { x: 7.309, y: 4.533, w: 2.26, h: 0.638, fontSize: 14, bold: true, lineSpacingMultiple: 1.3 });
}

/* -- 22 · Testimonial ----------------------------------------------------- */
// Two donut gauges: a full track plus an accent arc for the score.
// Ring gauge: a full track plus a coloured block-arc that runs clockwise from
// `arcStart` round to 5.2° (the small gap both rings share).
function gauge(slide, x, y, size, arcStart, trackColor, arcColor, label, labelX, value) {
  const ring = { w: size, h: size, arcThicknessRatio: 0.1247, line: { type: 'none' } };
  slide.addShape('blockArc', Object.assign({ x: x, y: y, angleRange: [8.7, 5.2], rotate: -92.5, fill: { color: trackColor } }, ring));
  slide.addShape('blockArc', Object.assign({ x: x, y: y, angleRange: [arcStart, 5.2], fill: { color: arcColor } }, ring));
  slide.addText(value, {
    x: x, y: y + size / 2 - 0.19, w: size, h: 0.38, margin: 0, align: 'center', valign: 'middle',
    fontFace: SANS, fontSize: 18, bold: true, color: C.ink,
  });
  slide.addText(label, {
    x: labelX, y: y + size + 0.153, w: 1.224, h: 0.278, margin: INSET,
    fontFace: SANS, fontSize: 12, bold: true, color: C.ink,
  });
}

function slide22() {
  const s = newSlide(22);
  mapleLeaf(s, 0.129, 0.239, 1.23, 1.34, -45);
  text(s, 'Testimonal', { x: 0.579, y: 0.462, w: 4.05, h: 0.909, fontFace: SCRIPT, fontSize: 50 });
  text(s, 'Isabella', { x: 0.578, y: 2.812, w: 1.938, h: 0.278, fontSize: 12, bold: true });
  text(s, 'Participant', { x: 0.584, y: 3.09, w: 1.939, h: 0.254, fontSize: 9, color: C.body, lineSpacingMultiple: 1.3 });
  captionBlock(s, 5.816, 0.457, 3.605, 'Description',
    LOREM + 'consectetur adipiscing elit, sed do eiusmod tempor incididunt ut', { bodyH: 0.451 });
  text(s, 'Rate the Event :', { x: 5.798, y: 1.976, w: 2.637, h: 0.278, fontSize: 12, bold: true });
  gauge(s, 5.817, 2.622, 0.88, 56.2, 'FEDFB4', C.amber, 'Cleanliness', 5.798, '85%');
  gauge(s, 7.586, 2.641, 0.88, 30.3, C.mist, C.silver, 'Comfort', 7.586, '90%');
  text(s, '“ ' + LOREM_LONG + ' “', {
    x: 5.798, y: 4.529, w: 3.631, h: 0.647, fontSize: 9, italic: true, color: C.body, lineSpacingMultiple: 1.3,
  });
  photo(s, 2.516, 1.799, 1.823, 2.605, 'pill');
}

/* -- 23 · Review ---------------------------------------------------------- */
function slide23() {
  const s = newSlide(23);
  mapleLeaf(s, 0.414, 3.054, 1.23, 1.34, -135);
  text(s, 'Review', { x: -0.059, y: 2.358, w: 2.176, h: 0.909, fontFace: SCRIPT, fontSize: 50, rotate: 270 });
  photo(s, 1.938, 0.451, 1.766, 2.532, 'pill');
  text(s, 'Daniel', { x: 3.804, y: 0.922, w: 1.938, h: 0.278, fontSize: 12, bold: true });
  text(s, 'Traveler', { x: 3.81, y: 1.2, w: 1.939, h: 0.254, fontSize: 9, color: C.body, lineSpacingMultiple: 1.3 });
  text(s, '“ ' + LOREM_LONG + ' “', { x: 3.804, y: 1.618, w: 3.631, h: 0.647, fontSize: 9, italic: true, color: C.body, lineSpacingMultiple: 1.3 });
  photo(s, 7.659, 2.672, 1.766, 2.532, 'pill');
  text(s, 'Grace', { x: 5.627, y: 3.143, w: 1.938, h: 0.278, fontSize: 12, bold: true, align: 'right' });
  text(s, 'Tourist', { x: 5.633, y: 3.421, w: 1.939, h: 0.254, fontSize: 9, color: C.body, lineSpacingMultiple: 1.3, align: 'right' });
  text(s, '“ ' + LOREM_LONG + ' “', {
    x: 3.941, y: 3.912, w: 3.631, h: 0.647, fontSize: 9, italic: true, color: C.body, lineSpacingMultiple: 1.3, align: 'right',
  });
}

/* -- 24 · Get in Touch ---------------------------------------------------- */
const CONTACT = [
  ['pin', 'Office Address :', '123 Maple Avenue, Toronto', 1.151, 1.235, 0.11],
  ['phone', 'Phone :', '+1 123 4567 890', 1.929, 2.033, 0.175],
  ['envelope', 'E-Mail :', 'mail@company.com', 2.791, 2.903, 0.175],
  ['globe', 'Website :', 'www.yourwebsite.com', 3.57, 3.66, 0.175],
];

function slide24() {
  const s = newSlide(24);
  panel(s, 0, 0, 5.0, 5.625, C.cream);
  mapleLeaf(s, 0.54, 3.771, 1.23, 1.34, -150);
  text(s, 'Get in Touch', { x: -0.79, y: 2.358, w: 3.643, h: 0.909, fontFace: SCRIPT, fontSize: 50, rotate: 270 });
  text(s, 'Let Us Know How We Can Assist', { x: 5.533, y: 0.452, w: 3.852, h: 0.343, fontSize: 14, bold: true, lineSpacingMultiple: 1.3 });
  CONTACT.forEach(([icon, label, value, ly, iy, iw], i) => {
    pictogram(s, icon, 5.533, iy, iw, 0.175);
    text(s, label, { x: 5.708, y: ly, w: 3.852, h: 0.343, fontSize: 14, bold: true, lineSpacingMultiple: 1.3 });
    text(s, value, { x: 5.533, y: ly + 0.343, w: 3.914, h: 0.254, fontSize: 9, color: C.body, lineSpacingMultiple: 1.3 });
    divider(s, 5.533, [1.823, 2.631, 3.44, 4.4][i], 2.879, 0, { color: C.body, width: 0.75 });
  });
  photo(s, 1.846, 1.285, 2.747, 4.34, 'arch');
}

/* -- 25 · Thank You ------------------------------------------------------- */
function slide25() {
  const s = newSlide(25);
  panel(s, 5.0, 0, 5.0, 5.625, C.cream);
  mapleLeaf(s, 0.111, 0.349, 1.555, 1.85, -60);
  text(s, 'Thank\nYou', { x: 0.575, y: 0.46, w: 3.48, h: 2.499, fontFace: SCRIPT, fontSize: 72 });
  text(s, 'For Your Attention', { x: 0.58, y: 4.791, w: 3.029, h: 0.379, fontSize: 18, bold: true });
  text(s, '“ ' + LOREM + 'consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus “', {
    x: 6.381, y: 4.527, w: 3.039, h: 0.647, fontSize: 9, italic: true, color: C.body, lineSpacingMultiple: 1.3,
  });
  photo(s, 6.381, 0.46, 2.439, 3.822, 'pill');
}

/* ================================================================= main ==*/

[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
 slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18,
 slide19, slide20, slide21, slide22, slide23, slide24, slide25].forEach(fn => fn());

const OUT = path.join(__dirname, path.basename(__filename, '.js') + '.pptx');
pptx.writeFile({ fileName: OUT }).then(() => console.log('wrote ' + OUT));
