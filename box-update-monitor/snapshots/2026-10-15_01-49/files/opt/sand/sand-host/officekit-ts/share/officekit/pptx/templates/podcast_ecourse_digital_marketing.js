/**
 * PODSPOT - Podcast & eCourse presentation template.
 * A standalone pptxgenjs re-creation of the 36-slide reference deck (13.333 x 7.5 in).
 *
 *   node 133dc2c7-caf0-4224-bc55-68607fd21efb_grok_final.js
 *
 * Photographs in the source deck are replaced by hatched colour blocks (see IMG /
 * watermark); everything else is drawn with native pptxgenjs shapes and text runs.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ---------------------------------------------------------------- palette -- */
const WHITE     = 'FFFFFF';
const BLACK     = '000000';
const CYAN      = '00B0F0';
const AMBER     = 'FFC000';
const CYAN2     = '15C2FF';
const INK       = '171717';
const GREY15    = 'D9D9D9';
const GREY05    = 'F2F2F2';
const GREY25    = 'BFBFBF';
const GREY75    = '404040';
const GREY85    = '262626';
const GREY50    = '808080';
const GREY95    = '0D0D0D';
const NEARBLACK = '080808';
const GREY35    = 'A6A6A6';
const RED       = 'F03628';
const PHOTO_BG = 'F4F7FB', PHOTO_EDGE = 'C9D6E4', PHOTO_TEXT = '8FA3B8';
const DEVICE_EDGE = '6E7378';

/* ------------------------------------------------------------ typography --- */
const MEB = 'Montserrat ExtraBold', MSB = 'Montserrat SemiBold';
const MMD = 'Montserrat Medium',    MBK = 'Montserrat Black';
const PSB = 'Poppins SemiBold',     PMD = 'Poppins Medium', POP = 'Poppins';
const LATO = 'Lato';

/* Paragraph presets reused across the deck. */
const C   = { align: 'center' };
const RT  = { align: 'right' };
const J   = { align: 'justify', lineSpacingMultiple: 1.5 };
const CJ  = { align: 'center',  lineSpacingMultiple: 1.5 };
const RJ  = { align: 'right',   lineSpacingMultiple: 1.5 };
const LS  = { lineSpacingMultiple: 1.5 };
const NW  = { wrap: false };
const CNW = { align: 'center', wrap: false };

/* --------------------------------------------------------- filler copy ----- */
const LOREM_A =
  "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ipsum dolor sit amet consectetur. Donec enim diam vulputate ut pharetra sit amet aliquam id. Risus pretium quam vulputate dignissim suspendisse in est. Sit amet mauris commodo quis. Nam libero justo laoreet sit amet cursus sit amet. Ultrices vitae auctor eu aug.";
const LOREM_B =
  "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nulla nisi odio, faucibus at justo sagittis, tristique sollicitudin ex. Vivamus id ullamcorper justo, a condimentum diam. Pellentesque eget feugiat orci. Donec eget consequat massa, eu mollis nunc. Pellentesque pharetra convallis leo ut molestie. Vivamus convalli.";
/** First `n` characters of a filler paragraph, plus an optional tail. */
const LA = (n, tail) => LOREM_A.slice(0, n) + (tail || '');
const LB = (n, tail) => LOREM_B.slice(0, n) + (tail || '');

/* -------------------------------------------------------------- helpers ---- */
/* Text boxes in the source deck use PowerPoint's default 0.1" / 0.05" insets. */
const INSET = [7.2, 7.2, 3.6, 3.6];

/** Single-run text box. */
function T (s, text, x, y, w, h, fontSize, fontFace, color, opts) {
  s.addText(text, Object.assign(
    { x, y, w, h, fontSize, fontFace, color, margin: INSET, valign: 'top' }, opts));
}

/** Multi-run / multi-paragraph text box. */
function TX (s, runs, x, y, w, h, opts) {
  s.addText(runs, Object.assign({ x, y, w, h, margin: INSET, valign: 'top' }, opts));
}

/** Generic filled shape; no outline unless one is supplied. */
function SH (s, kind, x, y, w, h, color, opts) {
  const o = Object.assign({ x, y, w, h }, opts);
  if (o.transparency !== undefined) {
    o.fill = { color, transparency: o.transparency };
    delete o.transparency;
  } else {
    o.fill = color === null ? { type: 'none' } : { color };
  }
  s.addShape(kind, o);
}

/** Soft drop shadow: [blur, distance, direction deg, colour, opacity 0-1]. */
const shadow = (blur, offset, angle, color, opacity) =>
  ({ type: 'outer', blur, offset, angle, color, opacity });

/**
 * A glow in the source deck; pptxgenjs cannot emit `a:glow` on shapes, so the
 * halo is faked with three concentric, increasingly transparent rings drawn
 * behind the shape itself.
 */
function glow (s, x, y, w, h, color, radius) {
  for (let i = 3; i >= 1; i--) {
    const g = (radius * i) / 3;
    s.addShape('ellipse', { x: x - g, y: y - g, w: w + 2 * g, h: h + 2 * g,
      fill: { type: 'none' },
      line: { color, width: (radius * 48) / 3, transparency: 55 + i * 12 } });
  }
}

const R  = (s, x, y, w, h, color, opts) => SH(s, 'rect', x, y, w, h, color, opts);
const O  = (s, x, y, w, h, color, opts) => SH(s, 'ellipse', x, y, w, h, color, opts);
const RR = (s, x, y, w, h, color, radius, opts) =>
  SH(s, 'roundRect', x, y, w, h, color, Object.assign({ rectRadius: radius }, opts));
/** Ring / hollow circle; `thickness` is the stroke width as a fraction of the box. */
const DN = (s, x, y, w, h, color, thickness, opts) =>
  SH(s, 'donut', x, y, w, h, color,
     Object.assign({ rectRadius: thickness * Math.min(w, h) }, opts));

/** Straight connector; `dashType` may be passed through `opts`. */
function LN (s, x, y, w, h, color, width, opts) {
  const o = Object.assign({}, opts);
  const line = { color, width };
  if (o.dashType) { line.dashType = o.dashType; delete o.dashType; }
  s.addShape('line', Object.assign({ x, y, w, h, line }, o));
}

/**
 * Freeform shape. `pts` holds fractions of the bounding box:
 * [x, y] is a line/move-to, [x1, y1, x2, y2, x, y] is a cubic bezier.
 */
function poly (s, x, y, w, h, color, pts, opts) {
  const points = pts.map((p, i) => p.length === 2
    ? { x: p[0] * w, y: p[1] * h, moveTo: i === 0 }
    : { x: p[4] * w, y: p[5] * h,
        curve: { type: 'cubic', x1: p[0] * w, y1: p[1] * h, x2: p[2] * w, y2: p[3] * h } });
  points.push({ close: true });
  SH(s, 'custGeom', x, y, w, h, color, Object.assign({ points }, opts));
}

/**
 * Stand-in for a photo: a pale block with a thin border, captioned at the top
 * where PowerPoint would print its "Picture" placeholder prompt.
 */
function IMG (s, x, y, w, h) {
  s.addShape('rect',
    { x, y, w, h, fill: { color: PHOTO_BG }, line: { color: PHOTO_EDGE, width: 0.75 } });
  s.addText('[image]', { x, y: y + 0.04, w, h: 0.26, align: 'center', valign: 'top',
    fontFace: LATO, fontSize: Math.max(7, Math.min(11, h * 4)), color: PHOTO_TEXT });
}

/** Stand-in for small glyph art: a rounded marker in the icon's dominant tone. */
function ICON (s, x, y, w, h, tone) {
  const i = 0.15;
  s.addShape('roundRect', { x: x + w * i, y: y + h * i, w: w * (1 - 2 * i), h: h * (1 - 2 * i),
    fill: { color: tone }, rectRadius: Math.min(w, h) * 0.15 });
}

/** Stand-in for a laptop / phone / monitor mock-up: an empty outlined shell. */
function device (s, x, y, w, h, pts, opts) {
  const line = { color: DEVICE_EDGE, width: 2.5 };
  if (pts) poly(s, x, y, w, h, null, pts, Object.assign({ line }, opts));
  else s.addShape('roundRect',
    Object.assign({ x, y, w, h, fill: { type: 'none' }, line,
      rectRadius: Math.min(w, h) * 0.06 }, opts));
}

/** Soft-edged (blurred) disc: nested discs fading out over `blur` inches. */
function haze (s, x, y, w, h, color, blur) {
  const steps = 6;
  for (let i = 0; i < steps; i++) {
    const d = (blur * 2 * i) / (steps - 1);
    s.addShape('ellipse', { x: x + d / 2, y: y + d / 2, w: w - d, h: h - d,
      fill: { color, transparency: Math.round(88 - (88 * i) / (steps - 1)) } });
  }
}

/**
 * pptxgenjs cannot emit `a:gradFill`, so a linear gradient is drawn as a stack
 * of bands. `stops` is [[position 0-1, colour, transparency], ...] and `vertical`
 * picks the sweep direction.
 */
function gradient (s, x, y, w, h, stops, vertical) {
  const BANDS = 24;
  const at = (t) => {
    let a = stops[0], b = stops[stops.length - 1];
    for (let i = 0; i < stops.length - 1; i++) {
      if (t >= stops[i][0]) { a = stops[i]; b = stops[i + 1]; }
    }
    const k = Math.max(0, Math.min(1, b[0] === a[0] ? 0 : (t - a[0]) / (b[0] - a[0])));
    const mix = (p, q) => Math.round(p + (q - p) * k);
    const chan = (c, i) => parseInt(c.slice(i, i + 2), 16);
    const rgb = [0, 2, 4].map(i => mix(chan(a[1], i), chan(b[1], i)));
    return { color: rgb.map(v => v.toString(16).padStart(2, '0').toUpperCase()).join(''),
             transparency: mix(a[2] || 0, b[2] || 0) };
  };
  for (let i = 0; i < BANDS; i++) {
    const f = at((i + 0.5) / BANDS);
    const o = vertical
      ? { x, y: y + (h * i) / BANDS, w, h: h / BANDS + 0.005 }
      : { x: x + (w * i) / BANDS, y, w: w / BANDS + 0.005, h };
    s.addShape('rect', Object.assign(o, { fill: f }));
  }
}

/** Faint decorative photo baked into a layout - rendered as a barely-there panel. */
const watermark = (s, x, y, w, h) =>
  s.addShape('rect', { x, y, w, h, fill: { color: '9A9A9E', transparency: 94 } });

/* --- Slide 1: POD --- */
function slide01 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  R(s, -0.012, 0, 13.338, 0.235, CYAN);
  haze(s, 1.333, 2.299, 2.667, 2.667, WHITE, 0.694);
  O(s, 12.75, 3.264, 0.236, 0.236, CYAN, { rotate: 90 });
  O(s, 12.75, 3.632, 0.236, 0.236, AMBER, { rotate: 90 });
  O(s, 12.75, 4, 0.236, 0.236, WHITE, { rotate: 90 });
  TX(s, [
    { text: "POD", options: { fontSize: 88, fontFace: MEB, color: CYAN, charSpacing: 6, ...NW } },
    { text: "SPOT", options: { fontSize: 88, fontFace: MEB, color: WHITE, charSpacing: 6, ...NW } }
  ], 3.016, 2.99, 7.302, 1.582);
  T(s, "PODCAST AND eCOURSE PRESENTATION TAMPLATE", 4.04, 4.41, 5.25, 0.303, 12, LATO, WHITE, C);
  T(s, "Podcast Digital Marketing 2022", 5.248, 7.068, 2.837, 0.303, 12, LATO, CYAN, C);
  LN(s, 7.938, 7.22, 1.074, 0, CYAN, 0.5, { flipH: true, flipV: true });
  LN(s, 4.321, 7.22, 1.074, 0, CYAN, 0.5, { flipH: true, flipV: true });
  T(s, "podcast.com", 11.849, 7, 1.476, 0.303, 12, LATO, AMBER);
  T(s, "Themonre", 0.309, 7, 1.441, 0.303, 12, LATO, AMBER);
}

/* --- Slide 2: PODSPOT --- */
function slide02 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  // backdrop inherited from layout "2_Title Slide"
  O(s, 6.333, 0.583, 6.333, 6.333, CYAN, { line: { color: WHITE, width: 6 } });
  watermark(s, 0.009, 1.333, 3.222, 4.833);
  T(s, "PODSPOT", 1.281, 2.959, 8.552, 1.582, 88, MEB, WHITE, { charSpacing: 6 });
  T(s, "Podcast And eCourse Presentation Template", 1.406, 4.525, 5, 0.303, 12, MSB, WHITE);
  IMG(s, 8.584, 0.803, 4.131, 6.146);
  T(s, LB(299, "."), 1.406, 4.828, 4.594, 1.576, 12, LATO, GREY15, J);
  T(s, "WELCOME TO", 1.406, 2.808, 1.667, 0.303, 12, MSB, WHITE);
  RR(s, 1.583, 6.706, 1.102, 0.354, CYAN, 0.177);
  T(s, "START", 1.583, 6.732, 1.113, 0.303, 12, MSB, WHITE, C);
  R(s, 0.833, -0.016, 1.167, 1.08, AMBER);
  TX(s, [
    { text: "P", options: { fontSize: 24, fontFace: MBK, color: WHITE, ...NW, breakLine: true } },
    { text: "C", options: { fontSize: 24, fontFace: MBK, color: WHITE, ...NW } }
  ], 0.645, 0.158, 0.453, 0.909);
}

/* --- Slide 3: DIGITAL --- */
function slide03 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  // backdrop inherited from layout "3_Title Slide"
  watermark(s, 10.111, 1.333, 3.222, 4.833);
  R(s, 0, 0.5, 3, 6.5, CYAN);
  T(s, "WELCOME TO THE PODSPOT", 8.983, 1.25, 3.6, 0.337, 14, MSB, GREY05, RT);
  TX(s, [
    { text: "DIGITAL", options: { fontSize: 44, fontFace: MEB, color: CYAN, ...RT } },
    { text: " ", options: { fontSize: 44, fontFace: MEB, color: BLACK, ...RT } },
    { text: "MARKETING 2022", options: { fontSize: 44, fontFace: MEB, color: WHITE, ...RT } }
  ], 7.333, 1.661, 5.25, 2.322);
  T(s, "EVERYTHING YOU NEED TO KNOW TO SUCCESED", 5.758, 4.122, 6.826, 0.337, 14, MSB, GREY05, RT);
  T(s, LB(321), 6.5, 4.458, 6.083, 1.273, 12, LATO, GREY15, J);
  R(s, 10.317, 6.326, 2.167, 0.317, CYAN);
  T(s, "EXPLORE KNOW", 10.417, 6.333, 1.967, 0.303, 12, MSB, BLACK, C);
  IMG(s, 1.083, 1.333, 3.833, 4.833);
}

/* --- Slide 4: JHONE DOE --- */
function slide04 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  RR(s, 10.245, 1.2, 2.334, 1.2, WHITE, 0.2, { shadow: shadow(4, 6, 80, '9DC3E6', 0.4) });
  RR(s, 10.245, 3.15, 2.334, 1.2, WHITE, 0.2, { shadow: shadow(4, 6, 80, '9DC3E6', 0.4) });
  RR(s, 10.326, 5.1, 2.253, 1.2, CYAN, 0.2, { shadow: shadow(4, 6, 80, WHITE, 0.4) });
  T(s, "JHONE DOE", 0.878, 3.131, 2.622, 0.404, 18, MSB, WHITE);
  T(s, LB(312, "Co."), 0.883, 3.502, 3.583, 2.182, 12, LATO, GREY15, J);
  T(s, "Hallo I’m", 0.878, 2.509, 2.205, 0.438, 20, MSB, WHITE);
  R(s, 0, 0.417, 1.917, 0.75, CYAN);
  T(s, "PODSPOT", 0, 0.965, 1.917, 0.404, 18, MSB, WHITE, C);
  glow(s, 4.729, 1.112, 5.276, 5.276, CYAN, 0.167);
  DN(s, 4.729, 1.112, 5.276, 5.276, CYAN, 0.014);
  TX(s, [
    { text: "Successful", options: { fontSize: 14, fontFace: MMD, color: CYAN, breakLine: true } },
    { text: "Campaign Executed", options: { fontSize: 14, fontFace: MMD, color: CYAN } }
  ], 10.295, 1.66, 2.788, 0.572);
  T(s, "315", 10.262, 1.321, 0.972, 0.404, 18, MMD, CYAN);
  TX(s, [
    { text: "Global", options: { fontSize: 14, fontFace: MMD, color: NEARBLACK, breakLine: true } },
    { text: "Advertising Award", options: { fontSize: 14, fontFace: MMD, color: NEARBLACK } }
  ], 10.326, 3.628, 2.491, 0.572);
  T(s, "15", 10.326, 3.224, 0.44, 0.404, 18, MMD, NEARBLACK, NW);
  TX(s, [
    { text: "Year", options: { fontSize: 14, fontFace: MMD, color: NEARBLACK, breakLine: true } },
    { text: "In Advertising", options: { fontSize: 14, fontFace: MMD, color: NEARBLACK } }
  ], 10.454, 5.607, 2.124, 0.572);
  T(s, "10", 10.454, 5.203, 0.463, 0.404, 18, MMD, NEARBLACK, NW);
  IMG(s, 5.083, 1.467, 4.567, 4.567);
}

/* --- Slide 5: “ --- */
function slide05 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  // backdrop inherited from layout "5_Title Slide"
  R(s, 0.458, 0.333, 12.417, 6.833, WHITE);
  watermark(s, 9.667, 1, 3.667, 5.5);
  TX(s, [
    { text: "“", options: { fontSize: 24, fontFace: MSB, color: BLACK, bold: true, ...C } },
    { text: "The quality of a person's life is in direct proportion to their commitment to excellence, regardless of their chosen field of endeavor”", options: { fontSize: 24, fontFace: MSB, color: BLACK, bold: true, ...C } }
  ], 5.667, 1.247, 6.667, 2.121);
  T(s, "VINCE LOMBARDI", 9.417, 3.747, 2.449, 0.337, 14, MSB, BLACK, C);
  R(s, 0.458, 0.333, 2.708, 1.417, '002060');
  T(s, LB(312, "Co."), 6.25, 5.083, 6.083, 1.273, 12, LATO, '595959', J);
}

/* --- Slide 6: ABOUT --- */
function slide06 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  // backdrop inherited from layout "6_Title Slide"
  watermark(s, 10.067, 1.459, 3.333, 4.582);
  poly(s, 8.167, 0.502, 3.248, 6.496, CYAN, [
    [0, 0], [0.552, 0, 1, 0.224, 1, 0.5], [1, 0.776, 0.552, 1, 0, 1], [0, 0.993],
    [0.544, 0.993, 0.985, 0.772, 0.985, 0.5], [0.985, 0.228, 0.544, 0.007, 0, 0.007]]);
  IMG(s, 0, 0, 8, 7.5);
  IMG(s, 10.4, 1.833, 1.333, 1.333);
  IMG(s, 10.4, 4.333, 1.333, 1.333);
  R(s, 0, 0, 8.167, 7.5, BLACK, { transparency: 7 });
  O(s, 9.067, 0.317, 1.333, 1.333, AMBER);
  O(s, 9.067, 5.85, 1.333, 1.333, AMBER);
  T(s, "85 %", 9.167, 0.605, 1.067, 0.404, 18, MSB, BLACK, C);
  T(s, "STATISTIC", 8.75, 1.009, 1.9, 0.337, 14, MSB, BLACK, C);
  T(s, "95%", 9.167, 6.155, 1.067, 0.404, 18, MSB, BLACK, C);
  T(s, "STATISTIC", 8.75, 6.559, 1.9, 0.337, 14, MSB, BLACK, C);
  TX(s, [
    { text: "ABOUT", options: { fontSize: 32, fontFace: MEB, color: WHITE, breakLine: true } },
    { text: "THE BRAND", options: { fontSize: 32, fontFace: MEB, color: WHITE } }
  ], 0.667, 0.655, 5.167, 1.178);
  T(s, "TITLE TEXT HERE", 0.667, 1.852, 2.315, 0.337, 14, MMD, WHITE);
  T(s, LB(215, "."), 0.667, 3.038, 4.4, 1.273, 12, LATO, GREY15, J);
  T(s, "TITLE TEXT HERE", 0.667, 2.705, 2.083, 0.337, 14, PSB, WHITE);
  T(s, LB(215, "."), 0.667, 5.083, 4.4, 1.273, 12, LATO, GREY15, J);
  T(s, "TITLE TEXT HERE", 0.667, 4.75, 2.083, 0.337, 14, PSB, WHITE);
  DN(s, 5.65, 1.399, 4.702, 4.702, CYAN, 0.042);
  IMG(s, 5.835, 1.585, 4.331, 4.331);
}

/* --- Slide 7: PODCAST AGENDA --- */
function slide07 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  IMG(s, 1, 0, 3.417, 7.5);
  IMG(s, 6.167, 0.75, 3.417, 6);
  R(s, 2.917, 1, 3.75, 2.417, AMBER);
  R(s, 2.9, 4.083, 3.75, 2.417, '4472C4');
  T(s, "Part 1", 3.25, 1.525, 1.083, 0.337, 14, PSB, BLACK);
  T(s, "Part 2", 3.3, 4.578, 1.417, 0.337, 14, PSB, WHITE);
  R(s, 9.333, 0.758, 3.167, 2.417, CYAN);
  R(s, 9.333, 4.359, 3.167, 2.417, WHITE);
  T(s, "Part 3", 9.617, 1.189, 1.192, 0.337, 14, PSB, WHITE);
  T(s, "Part 4", 9.619, 4.69, 1.192, 0.337, 14, PSB, BLACK);
  T(s, "PODCAST AGENDA", -2.261, 3.564, 5.667, 0.572, 28, MEB, WHITE, { charSpacing: 6, rotate: 90 });
  T(s, LB(158), 3.25, 1.862, 3.131, 1.273, 12, LATO, GREY75, J);
  T(s, LB(158), 3.285, 4.854, 3.131, 1.273, 12, LATO, GREY15, J);
  T(s, LB(128), 9.619, 1.537, 2.631, 1.273, 12, LATO, GREY15, J);
  T(s, LB(128), 9.619, 5.027, 2.631, 1.273, 12, LATO, GREY75, J);
}

/* --- Slide 8: PODCAST COURSE ONLINE --- */
function slide08 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  IMG(s, 0, 2.75, 13.333, 4.75);
  R(s, -0.001, 2.758, 13.333, 4.75, BLACK, { transparency: 10 });
  T(s, "PODCAST COURSE ONLINE", 3.354, 0.837, 6.622, 0.64, 32, MEB, BLACK, NW);
  T(s, "PODSPOT", 5.86, 0.435, 1.61, 0.337, 14, MMD, BLACK, C);
  T(s, LB(215, "."), 2.019, 1.579, 9.293, 0.667, 12, LATO, GREY75, CJ);
  R(s, 0.667, 3.55, 2.417, 1.583, CYAN);
  R(s, 0.667, 5.533, 2.417, 1.583, GREY05);
  R(s, 3.917, 3.548, 2.417, 1.583, GREY05);
  R(s, 3.917, 5.533, 2.417, 1.583, GREY05);
  R(s, 7.167, 3.55, 2.417, 1.583, GREY05);
  R(s, 7.167, 5.533, 2.417, 1.583, GREY05);
  R(s, 10.25, 3.55, 2.417, 1.583, GREY05);
  R(s, 10.25, 5.533, 2.417, 1.583, AMBER);
  T(s, "Modul one", 0.949, 4.005, 1.551, 0.335, 14, PSB, BLACK);
  T(s, "Modul Two", 4.199, 4.005, 1.384, 0.337, 14, PSB, BLACK);
  T(s, "Modul Three", 7.5, 4.005, 1.667, 0.337, 14, PSB, BLACK);
  T(s, "Modul Four", 10.583, 4.003, 1.583, 0.337, 14, PSB, BLACK);
  T(s, "Modul Five", 0.949, 5.988, 1.384, 0.337, 14, PSB, BLACK);
  T(s, "Modul Six", 4.199, 5.988, 1.384, 0.337, 14, PSB, BLACK);
  T(s, "Modul Seven", 7.449, 5.988, 1.718, 0.337, 14, PSB, BLACK);
  T(s, "Modul Eight", 10.567, 5.988, 1.818, 0.337, 14, PSB, BLACK);
  T(s, "Digital Marketing Strategy", 0.949, 4.34, 1.801, 0.505, 12, LATO, BLACK);
  T(s, "Planning Digital Assets", 4.199, 4.34, 1.801, 0.505, 12, LATO, BLACK);
  T(s, "Speaking To Your audience", 7.5, 4.34, 1.801, 0.505, 12, LATO, BLACK);
  T(s, "Creating Digital Assets", 10.583, 4.34, 1.801, 0.505, 12, LATO, BLACK);
  T(s, "Direct Marketing", 0.949, 6.325, 1.801, 0.303, 12, LATO, BLACK);
  T(s, "Social Media", 4.199, 6.325, 1.801, 0.303, 12, LATO, BLACK);
  T(s, "Online Advertising", 7.449, 6.325, 1.801, 0.303, 12, LATO, BLACK);
  T(s, "Optimations", 10.567, 6.319, 1.801, 0.303, 12, LATO, BLACK);
}

/* --- Slide 9: THIS PODCAST IS FOR YOU --- */
function slide09 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  IMG(s, 0.583, 0.25, 4.083, 2.583);
  IMG(s, 8.667, 0.25, 4.083, 2.583);
  R(s, 4.208, 0.417, 4.917, 4.37, CYAN);
  T(s, "THIS PODCAST IS FOR YOU", 4.917, 0.665, 3.5, 1.043, 28, MEB, BLACK, C);
  T(s, "Title Text Here", 0.583, 5.141, 2.25, 0.337, 14, PSB, WHITE);
  T(s, "Title Text Here", 10.657, 5.155, 2.25, 0.337, 14, PSB, WHITE, RT);
  T(s, LB(215, "."), 0.583, 5.477, 4.398, 1.273, 12, LATO, GREY15, J);
  T(s, LB(128), 0.583, 3.495, 3.25, 0.97, 12, LATO, GREY15, J);
  T(s, LB(215, "."), 8.509, 5.477, 4.398, 1.273, 12, LATO, GREY15, J);
  T(s, "Title Text Here", 0.583, 3.173, 2.818, 0.337, 14, PSB, WHITE);
  T(s, "Title Text Here", 10.583, 3.17, 2.25, 0.337, 14, PSB, WHITE, RT);
  T(s, LB(157, "."), 4.751, 1.914, 3.696, 1.273, 12, LATO, GREY75, CJ);
  T(s, LB(128), 9.583, 3.495, 3.25, 0.97, 12, LATO, GREY15, RJ);
  RR(s, 6.016, 3.75, 1.167, 0.337, AMBER, 0.169);
  T(s, "Explor now", 5.919, 3.775, 1.361, 0.286, 11, MMD, BLACK, C);
}

/* --- Slide 10: THIS PODCAST IS NOT FOR YOU IF --- */
function slide10 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  R(s, 0, 0, 13.333, 2.235, BLACK);
  R(s, 8.494, 4.417, 4.336, 2.333, AMBER);
  T(s, "THIS PODCAST IS NOT FOR YOU IF", 3.294, 0.498, 6.746, 0.505, 24, MEB, WHITE, C);
  T(s, "01. Your Title Text Here", 0.417, 4.962, 2.244, 0.303, 12, PSB, BLACK);
  T(s, "01. Your Title Text Here", 3.083, 4.962, 2.241, 0.303, 12, PSB, BLACK);
  T(s, "01. Your Title Text Here", 5.75, 4.962, 2.172, 0.303, 12, PSB, BLACK);
  T(s, "01. Your Title Text Here", 8.711, 4.962, 2.447, 0.303, 12, PSB, BLACK);
  T(s, LB(119, "."), 0.417, 5.265, 2.33, 1.273, 12, LATO, GREY75, J);
  T(s, LB(119, "."), 3.08, 5.265, 2.33, 1.273, 12, LATO, GREY75, J);
  T(s, LB(119, "."), 5.744, 5.265, 2.33, 1.273, 12, LATO, GREY75, J);
  T(s, LB(187, "."), 8.711, 5.265, 3.789, 1.273, 12, LATO, GREY75, J);
  IMG(s, 0.503, 1.417, 2.33, 3.167);
  IMG(s, 3.167, 1.417, 2.33, 3.167);
  IMG(s, 5.83, 1.417, 2.33, 3.167);
  IMG(s, 8.494, 1.417, 4.336, 3.167);
}

/* --- Slide 11: 75 % --- */
function slide11 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  IMG(s, 0.333, 0.333, 12.5, 5.167);
  R(s, 9.75, 5.867, 3.083, 1.467, AMBER);
  T(s, "01. Title Text Here", 0.367, 3.908, 2.217, 0.303, 12, PSB, WHITE);
  T(s, LB(119, "."), 0.367, 4.211, 2.33, 1.273, 12, LATO, GREY15, J);
  T(s, "75 %", 10.596, 5.842, 1.391, 0.572, 28, PSB, BLACK, C);
  T(s, LB(91, "."), 9.75, 6.229, 3.083, 0.97, 12, LATO, GREY75, CJ);
  T(s, "02. Title Text Here", 3.5, 4.544, 2.167, 0.303, 12, PSB, WHITE);
  T(s, LB(119, "."), 3.5, 4.847, 2.33, 1.273, 12, LATO, GREY15, J);
  T(s, "03. Title Text Here", 6.633, 5.18, 2.2, 0.303, 12, PSB, WHITE);
  TX(s, [
    { text: LB(119), options: { fontSize: 12, fontFace: LATO, color: GREY15, ...J } },
    { text: ".", options: { fontSize: 12, fontFace: LATO, color: GREY75, ...J } }
  ], 6.633, 5.483, 2.33, 1.273);
  T(s, "WHAT YOUR LEARN IN THIS PODCAST", 7.917, 3.576, 4.682, 1.043, 28, MEB, CYAN);
}

/* --- Slide 12: YOUR COURSE PODCAST TOOLS --- */
function slide12 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  R(s, 0.692, 1.433, 3.117, 1.858, CYAN);
  R(s, 11.779, 2.944, 0.05, 1.133, CYAN);
  R(s, 9.149, 4.077, 0.05, 1.772, CYAN);
  R(s, 4.135, 2.131, 0.05, 1.772, CYAN);
  R(s, 1.53, 4.086, 0.05, 1.133, CYAN);
  RR(s, 0.583, 4.018, 4.25, 0.118, CYAN, 0.059);
  RR(s, 8.517, 4.018, 4.25, 0.118, CYAN, 0.059);
  T(s, "YOUR COURSE PODCAST TOOLS", 2.843, 0.333, 7.648, 0.572, 28, MEB, WHITE, C);
  T(s, "Here’s What You”ll Need for This Lasson or Modul", 4.352, 0.855, 4.629, 0.303, 12, PSB, GREY05, C);
  T(s, "Tool 01", 0.871, 1.634, 1.417, 0.337, 14, PSB, BLACK);
  O(s, 1.394, 3.917, 0.32, 0.32, AMBER);
  O(s, 4, 3.917, 0.32, 0.32, AMBER);
  O(s, 9.014, 3.917, 0.32, 0.32, AMBER);
  O(s, 11.62, 3.917, 0.32, 0.32, AMBER);
  R(s, 1.394, 5.219, 3.117, 1.858, CYAN);
  R(s, 9.602, 5.225, 3.117, 1.858, CYAN);
  R(s, 9.014, 1.433, 3.117, 1.858, CYAN);
  T(s, LB(102), 0.871, 1.974, 2.701, 0.97, 12, LATO, GREY75, J);
  T(s, LB(102), 1.602, 5.818, 2.701, 0.97, 12, LATO, GREY75, J);
  T(s, LB(102), 9.222, 2.054, 2.701, 0.97, 12, LATO, GREY75, J);
  T(s, LB(102), 9.81, 5.818, 2.701, 0.97, 12, LATO, GREY75, J);
  T(s, "Tool 02", 9.199, 1.634, 1.417, 0.337, 14, PSB, BLACK);
  T(s, "Tool 04", 9.805, 5.482, 1.417, 0.337, 14, PSB, BLACK);
  T(s, "Tool 03", 1.602, 5.482, 1.417, 0.337, 14, PSB, BLACK);
  glow(s, 4.463, 1.811, 4.413, 4.413, GREY05, 0.125);
  DN(s, 4.463, 1.811, 4.413, 4.413, WHITE, 0.014);
  IMG(s, 4.646, 2.042, 4.042, 4.042);
}

/* --- Slide 13: IMPORTANT PODCAST DATES --- */
function slide13 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  RR(s, 0.871, 4.281, 2.667, 1.327, CYAN, 0.197);
  RR(s, 3.833, 4.306, 2.667, 1.327, CYAN, 0.197);
  RR(s, 6.833, 4.306, 2.667, 1.327, CYAN, 0.18);
  RR(s, 9.795, 4.306, 2.667, 1.327, CYAN, 0.23);
  RR(s, 0.871, 1.667, 2.667, 2.917, GREY05, 0.194);
  RR(s, 3.833, 1.692, 2.667, 2.917, GREY05, 0.194);
  RR(s, 6.833, 1.692, 2.667, 2.917, GREY05, 0.194);
  RR(s, 9.795, 1.692, 2.667, 2.917, GREY05, 0.194);
  T(s, "IMPORTANT PODCAST DATES", 3.455, 0.737, 6.424, 0.572, 28, MEB, WHITE, NW);
  T(s, "JANUARY", 1.067, 1.882, 1.138, 0.286, 11, LATO, BLACK);
  T(s, "FEBRUARY", 4.083, 1.867, 1.346, 0.286, 11, LATO, BLACK);
  T(s, "MARCH", 7.1, 1.851, 1.067, 0.286, 11, LATO, BLACK);
  T(s, "APRIL", 10.117, 1.835, 0.954, 0.286, 11, LATO, BLACK);
  T(s, LA(55, "."), 0.977, 4.854, 2.417, 0.667, 12, LATO, GREY75, CJ);
  O(s, 2.024, 4.642, 0.313, 0.313, AMBER);
  T(s, "01", 1.9, 4.651, 0.527, 0.337, 14, PMD, BLACK, C);
  T(s, LA(55, "."), 4, 4.854, 2.417, 0.667, 12, LATO, GREY75, CJ);
  O(s, 5.047, 4.642, 0.313, 0.313, AMBER);
  T(s, "02", 4.992, 4.631, 0.437, 0.337, 14, PMD, BLACK, NW);
  T(s, LA(55, "."), 7.012, 4.854, 2.417, 0.667, 12, LATO, GREY75, CJ);
  O(s, 8.059, 4.642, 0.313, 0.313, AMBER);
  T(s, "03", 7.995, 4.631, 0.442, 0.337, 14, PMD, BLACK, NW);
  T(s, LA(55, "."), 10.024, 4.854, 2.417, 0.667, 12, LATO, GREY75, CJ);
  O(s, 11.071, 4.642, 0.313, 0.313, AMBER);
  T(s, "04", 11.003, 4.631, 0.449, 0.337, 14, PMD, BLACK, NW);
  T(s, LB(215, "."), 1.9, 6.094, 9.533, 0.667, 12, LATO, GREY15, CJ);
  IMG(s, 1.083, 2.667, 2.205, 1.614);
  IMG(s, 4.083, 2.667, 2.205, 1.614);
  IMG(s, 7.045, 2.617, 2.205, 1.614);
  IMG(s, 10.045, 2.667, 2.205, 1.614);
}

/* --- Slide 14: 5 STEP GUIDE --- */
function slide14 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  // backdrop inherited from layout "14_Title Slide"
  R(s, 5.5, 0.43, 7.333, 6.693, '19C3FF');
  IMG(s, 0.5, 0.396, 5, 6.693);
  R(s, 6.25, 1.083, 2.433, 1.833, AMBER);
  R(s, 9.733, 1.083, 2.433, 1.833, BLACK);
  R(s, 6.25, 4.583, 2.433, 1.833, '002060');
  R(s, 9.733, 4.583, 2.433, 1.833, AMBER);
  R(s, 8, 2.828, 2.417, 1.833, WHITE);
  gradient(s, 0.5, 2.25, 5, 4.873, [[0, INK, 100], [0.375, '0C0C0C', 32], [0.75, BLACK, 0]], true);
  T(s, "5 STEP GUIDE", 1.105, 3.49, 3.645, 0.572, 28, MEB, WHITE);
  T(s, LB(187, "."), 1.105, 4.062, 3.811, 1.273, 12, LATO, GREY05, J);
  RR(s, 2.083, 5.926, 1.25, 0.315, CYAN, 0.116);
  T(s, "EXPLORE NOW", 2, 5.958, 1.417, 0.252, 9, MEB, BLACK, { ...C, bold: true });
  T(s, "Step One", 6.217, 1.512, 1.4, 0.337, 14, MEB, BLACK, { ...C, bold: true });
  T(s, "Step Two", 9.717, 1.512, 1.4, 0.337, 14, MEB, WHITE, { ...C, bold: true });
  T(s, "Step Three", 8.421, 3.32, 1.4, 0.337, 14, MEB, BLACK, { ...C, bold: true });
  T(s, "Step Four", 6.25, 5.121, 1.4, 0.337, 14, MEB, WHITE, { ...C, bold: true });
  T(s, "Step Five", 9.854, 5.121, 1.4, 0.337, 14, MEB, BLACK, { bold: true });
  T(s, LA(55, "."), 6.333, 1.835, 2.258, 0.667, 12, LATO, GREY75, J);
  T(s, LA(55, "."), 9.821, 1.853, 2.258, 0.667, 12, LATO, GREY75, J);
  T(s, LA(55, "."), 8.079, 3.625, 2.258, 0.667, 12, LATO, GREY75, CJ);
  T(s, LA(55, "."), 6.333, 5.458, 2.258, 0.667, 12, LATO, GREY35, J);
  T(s, LA(55, "."), 9.854, 5.458, 2.258, 0.667, 12, LATO, GREY75, J);
}

/* --- Slide 15: 5 STEP GUIDE --- */
function slide15 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  IMG(s, 0.417, 0.424, 2.583, 2.013);
  IMG(s, 3.417, 2.743, 2.583, 2.013);
  IMG(s, 0.417, 5.083, 2.583, 2.013);
  R(s, 3.033, 0.404, 3, 2.054, CYAN);
  T(s, "5 STEP GUIDE", 7.75, 2.315, 3.058, 0.572, 28, MEB, WHITE, NW);
  T(s, LB(312, "Co."), 7.75, 2.975, 3.75, 2.182, 12, LATO, GREY35, J);
  T(s, "Part 1", 0.488, 1.564, 0.923, 0.37, 16, PSB, GREY25);
  T(s, "Title Text Here", 0.488, 1.952, 1.708, 0.337, 14, MMD, GREY25);
  R(s, 0.417, 2.743, 3, 2.013, AMBER);
  R(s, 3.033, 5.083, 3, 2.063, CYAN);
  T(s, LB(157, "."), 3.242, 0.609, 2.583, 1.576, 12, LATO, GREY85, J);
  T(s, LB(157, "."), 0.625, 2.962, 2.583, 1.576, 12, LATO, GREY85, J);
  T(s, LB(157, "."), 3.242, 5.327, 2.583, 1.576, 12, LATO, GREY85, J);
  LN(s, 6, 0.404, 6.833, 0, WHITE, 0.5);
  LN(s, 6.033, 7.1, 6.833, 0, WHITE, 0.5);
  LN(s, 12.833, 0.404, 0.033, 6.689, WHITE, 0.5);
  T(s, "Part 2", 4.684, 3.891, 1.037, 0.37, 16, PSB, GREY25, RT);
  T(s, "Title Text Here", 3.792, 4.262, 1.979, 0.337, 14, MMD, GREY25, RT);
  T(s, "Part 3", 0.555, 6.346, 1.112, 0.37, 16, PSB, GREY25);
  T(s, "Title Text Here", 0.555, 6.734, 1.641, 0.337, 14, MMD, GREY25, NW);
}

/* --- Slide 16: CASE STUDY --- */
function slide16 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  // backdrop inherited from layout "16_Title Slide"
  R(s, 0, 0, 4.5, 7.5, BLACK);
  DN(s, 5.022, 1.438, 2.458, 2.458, CYAN, 0.028);
  DN(s, 7.855, 1.438, 2.458, 2.458, CYAN, 0.028);
  DN(s, 10.688, 1.438, 2.458, 2.458, CYAN, 0.028);
  RR(s, 11.161, 4.274, 1.654, 0.354, AMBER, 0.177);
  RR(s, 8.21, 4.3, 1.654, 0.354, CYAN, 0.177);
  RR(s, 5.337, 4.268, 1.654, 0.354, '4472C4', 0.177);
  T(s, "CASE STUDY", 0.454, 1.591, 3.597, 0.572, 28, MEB, WHITE);
  T(s, "Title Text Here", 0.454, 2.499, 2.165, 0.337, 14, PMD, WHITE);
  T(s, "Case Study One", 5.23, 4.3, 1.867, 0.303, 12, MSB, BLACK, C);
  T(s, "Case Study Two", 8.103, 4.3, 1.867, 0.303, 12, MSB, BLACK, C);
  T(s, "Case Study Three", 11.054, 4.3, 1.867, 0.303, 12, MSB, BLACK, C);
  T(s, "78%", 0.798, 5.656, 1.369, 0.572, 28, MEB, WHITE);
  T(s, "Statistic", 0.877, 6.228, 1.123, 0.303, 12, MSB, WHITE);
  T(s, LB(312, "Co."), 0.454, 2.852, 3.75, 2.182, 12, LATO, GREY35, J);
  T(s, LB(157, "."), 4.958, 4.804, 2.583, 1.576, 12, LATO, GREY85, J);
  T(s, LB(157, "."), 7.792, 4.804, 2.583, 1.576, 12, LATO, GREY85, J);
  T(s, LB(157, "."), 10.635, 4.8, 2.583, 1.576, 12, LATO, GREY85, J);
  LN(s, 2, 5.656, 0, 0.866, '4472C4', 4.5);
  T(s, "Lorem ipsum dolor sit amet, co consectetur adipiscing elit. Nula .", 2.263, 5.561, 1.788, 0.97, 12, LATO, GREY25, J);
  IMG(s, 5.168, 1.585, 2.165, 2.165);
  IMG(s, 8.001, 1.585, 2.165, 2.165);
  IMG(s, 10.833, 1.585, 2.165, 2.165);
}

/* --- Slide 17: THE ULTIMATE DIGITAL MARKETING FORMULA --- */
function slide17 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  // backdrop inherited from layout "17_Title Slide"
  SH(s, 'frame', 0.525, 2.356, 12.283, 3.622, WHITE);
  IMG(s, 9.417, 2.585, 3.165, 3.165);
  IMG(s, 0.751, 2.585, 3.165, 3.165);
  R(s, 3.535, 2.585, 3.165, 3.182, CYAN);
  R(s, 6.7, 2.58, 3.165, 3.182, AMBER);
  T(s, "THE ULTIMATE DIGITAL MARKETING FORMULA", 1.62, 0.803, 10.093, 0.539, 26, MEB, WHITE, C);
  T(s, "CONSISTENCY + EFFORT = SUCCESS", 3.925, 1.346, 5.5, 0.404, 18, MSB, AMBER, C);
  T(s, "Consistency", 3.75, 2.75, 1.751, 0.37, 16, PSB, BLACK);
  T(s, "Effort", 7.833, 2.75, 1.751, 0.404, 18, PSB, BLACK, RT);
  T(s, LB(84, "re."), 4.117, 3.137, 2.351, 0.97, 12, LATO, GREY85, J);
  T(s, LB(84, "re."), 4.117, 4.255, 2.351, 0.97, 12, LATO, GREY85, J);
  T(s, LB(84, "re."), 6.983, 3.137, 2.351, 0.97, 12, LATO, GREY85, J);
  T(s, LB(84, "re."), 6.983, 4.255, 2.351, 0.97, 12, LATO, GREY85, J);
  SH(s, 'diamond', 3.842, 3.319, 0.165, 0.165, BLACK);
  SH(s, 'diamond', 3.851, 4.397, 0.165, 0.165, BLACK);
  SH(s, 'diamond', 9.42, 3.319, 0.165, 0.165, BLACK);
  SH(s, 'diamond', 9.429, 4.397, 0.165, 0.165, BLACK);
}

/* --- Slide 18: THE ULTIMATE DIGITAL MARKETING FORMULA --- */
function slide18 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  // backdrop inherited from layout "18_Title Slide"
  DN(s, 4.383, 1.467, 4.567, 4.567, AMBER, 0.057);
  IMG(s, 4.625, 1.709, 4.083, 4.083);
  glow(s, 3.517, 0.55, 6.299, 6.299, '4472C4', 0.153);
  DN(s, 3.517, 0.55, 6.299, 6.299, CYAN, 0.012);
  DN(s, 5.012, 2.046, 3.308, 3.308, CYAN, 0.02);
  glow(s, 8.523, 1.167, 0.371, 0.371, '4472C4', 0.153);
  O(s, 8.523, 1.167, 0.371, 0.371, CYAN);
  glow(s, 8.527, 5.88, 0.371, 0.371, '4472C4', 0.153);
  O(s, 8.527, 5.88, 0.371, 0.371, CYAN);
  T(s, "Consistency", 10.658, 0.797, 1.751, 0.337, 14, PMD, WHITE);
  T(s, LB(84, "re."), 10.658, 1.132, 2.351, 0.97, 12, LATO, GREY15, J);
  T(s, LB(84, "re."), 10.658, 2.25, 2.351, 0.97, 12, LATO, GREY15, J);
  T(s, "Effort", 10.658, 4.141, 1.042, 0.337, 14, PMD, WHITE);
  T(s, LB(84, "re."), 10.658, 4.477, 2.351, 0.97, 12, LATO, GREY15, J);
  T(s, LB(84, "re."), 10.658, 5.595, 2.351, 0.97, 12, LATO, GREY15, J);
  T(s, "THE ULTIMATE DIGITAL MARKETING FORMULA", 0.474, 0.55, 3.676, 1.851, 26, MEB, WHITE);
  glow(s, 3.375, 3.565, 0.371, 0.371, '4472C4', 0.153);
  O(s, 3.375, 3.565, 0.371, 0.371, CYAN);
  T(s, "Consistency + Effort = Success", 0.201, 3.464, 2.98, 0.572, 14, PMD, WHITE, C);
  R(s, 1.083, 5.88, 0.833, 1.62, WHITE);
  SH(s, 'diamond', 10.292, 1.287, 0.197, 0.197, WHITE);
  SH(s, 'diamond', 10.292, 4.636, 0.197, 0.197, WHITE);
}

/* --- Slide 19: FURTHER RESOUECES --- */
function slide19 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  // backdrop inherited from layout "19_Title Slide"
  O(s, 5.667, 3.767, 1, 1, CYAN, { line: { color: '57D3FF', width: 6 } });
  O(s, 6.4, 2.658, 1.317, 1.317, AMBER, { line: { color: 'FFD85B', width: 6 } });
  O(s, 7.717, 3.2, 1.567, 1.567, WHITE, { line: { color: GREY15, width: 6 } });
  O(s, 9.283, 2.367, 1.9, 1.9, 'ED7D31', { line: { color: 'F19E65', width: 6 } });
  O(s, 11.183, 3.45, 1.067, 1.067, '0070C0', { line: { color: '3BABFF', width: 6 } });
  LN(s, 7, 2.035, 0, 0.694, AMBER, 3, { dashType: 'sysDash' });
  T(s, "FURTHER RESOUECES", 0.583, 1.545, 3.917, 1.313, 36, MEB, WHITE);
  T(s, "Title Text Here", 0.583, 3.198, 1.833, 0.337, 14, PMD, WHITE);
  T(s, LB(312, "Co."), 0.583, 3.568, 3.75, 2.182, 12, LATO, GREY35, J);
  T(s, "Resources 2", 6.167, 0.583, 1.583, 0.337, 14, PMD, AMBER, C);
  T(s, "Resources 1", 5.375, 5.444, 1.583, 0.337, 14, PMD, CYAN, C);
  T(s, "Resources 3", 7.708, 5.444, 1.583, 0.337, 14, PMD, WHITE, C);
  T(s, "Resources 4", 9.417, 0.583, 1.583, 0.337, 14, PMD, 'ED7D31', C);
  T(s, "Resources 5", 10.958, 5.444, 1.583, 0.337, 14, PMD, '0070C0', C);
  T(s, LA(55, ". Nulla nisi odio."), 5.771, 0.92, 2.375, 0.97, 12, LATO, GREY15, CJ);
  T(s, LA(55, ". Nulla nisi odio."), 4.979, 5.78, 2.375, 0.97, 12, LATO, GREY15, CJ);
  T(s, LA(55, ". Nulla nisi odio."), 7.396, 5.78, 2.375, 0.97, 12, LATO, GREY15, CJ);
  T(s, LA(55, ". Nulla nisi odio."), 10.562, 5.78, 2.375, 0.97, 12, LATO, GREY15, CJ);
  T(s, LA(55, ". Nulla nisi odio."), 9.021, 0.92, 2.375, 0.97, 12, LATO, GREY15, CJ);
  LN(s, 6.167, 4.833, 0, 0.61, CYAN, 3, { dashType: 'sysDash' });
  LN(s, 8.5, 4.833, 0, 0.61, WHITE, 3, { dashType: 'sysDash', flipH: true });
  LN(s, 10.217, 1.905, 0, 0.428, 'ED7D31', 3, { dashType: 'sysDash' });
  LN(s, 11.75, 4.422, 0, 1.022, '0070C0', 3, { dashType: 'sysDash' });
  ICON(s, 6.667, 2.873, 0.831, 0.831, BLACK);
  ICON(s, 5.897, 4.003, 0.54, 0.54, BLACK);
  ICON(s, 8.083, 3.503, 1, 1, BLACK);
  ICON(s, 9.762, 2.835, 1, 1, BLACK);
  ICON(s, 11.445, 3.663, 0.61, 0.61, BLACK);
}

/* --- Slide 20: POD --- */
function slide20 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  TX(s, [
    { text: LA(215), options: { fontSize: 12, fontFace: LATO, color: GREY25, ...CJ } },
    { text: ",", options: { fontSize: 12, fontFace: LATO, color: GREY25, ...CJ } }
  ], 0.417, 3.383, 3.73, 1.576);
  TX(s, [
    { text: "POD", options: { fontSize: 28, fontFace: MEB, color: CYAN, ...C } },
    { text: "SPOT TEAM", options: { fontSize: 28, fontFace: MEB, color: WHITE, ...C } }
  ], 0.417, 2.667, 3.73, 0.572);
  T(s, "RUBEN CONNOR", 5.026, 5.746, 1.856, 0.421, 14, PMD, WHITE, CJ);
  T(s, "SARAH SMITH", 7.75, 5.729, 1.856, 0.421, 14, PMD, WHITE, CJ);
  T(s, "SAMUEL DUNN", 10.487, 5.744, 1.856, 0.421, 14, PMD, WHITE, CJ);
  T(s, "CEO PODSPOT", 5.162, 6.243, 1.584, 0.341, 10.5, PMD, GREY05, CJ);
  T(s, "eCOURSE MANAGER", 7.652, 6.243, 2.079, 0.341, 10.5, PMD, GREY05, CJ);
  T(s, "HEAD OF MARKETING", 10.487, 6.243, 1.856, 0.341, 10.5, PMD, GREY05, CJ);
  IMG(s, 4.664, 1.486, 2.606, 4.18);
  IMG(s, 7.388, 1.486, 2.606, 4.18);
  IMG(s, 10.112, 1.486, 2.606, 4.18);
}

/* --- Slide 21: POD --- */
function slide21 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: '1A1A1A' };
  RR(s, 0.796, 4.267, 2.662, 1.358, RED, 0.142);
  RR(s, 3.748, 4.267, 2.662, 1.358, RED, 0.142);
  RR(s, 6.7, 4.267, 2.662, 1.358, RED, 0.142);
  RR(s, 9.717, 4.267, 2.662, 1.358, AMBER, 0.142);
  poly(s, 0.796, 1.75, 2.662, 2.833, WHITE, [
    [0, 0.053], [0, 0.024, 0.025, 0, 0.057, 0], [0.943, 0], [0.975, 0, 1, 0.024, 1, 0.053],
    [1, 0.947], [1, 0.976, 0.975, 1, 0.943, 1], [0.057, 1], [0.025, 1, 0, 0.976, 0, 0.947],
    [0, 0.053]]);
  RR(s, 3.748, 1.75, 2.662, 2.833, WHITE, 0.151);
  RR(s, 6.7, 1.75, 2.662, 2.833, WHITE, 0.151);
  RR(s, 9.717, 1.75, 2.662, 2.833, WHITE, 0.151);
  TX(s, [
    { text: "POD", options: { fontSize: 28, fontFace: MEB, color: CYAN, ...C } },
    { text: "SPOT CREATIVE TEAM", options: { fontSize: 28, fontFace: MEB, color: WHITE, ...C } }
  ], 3.534, 0.594, 6.265, 0.572);
  RR(s, 1.043, 2, 1.25, 0.059, RED, 0.029);
  RR(s, 4, 2, 1.25, 0.059, RED, 0.029);
  RR(s, 6.935, 2, 1.25, 0.059, RED, 0.029);
  RR(s, 9.988, 2, 1.25, 0.059, RED, 0.029);
  TX(s, [
    { text: LA(215), options: { fontSize: 12, fontFace: LATO, color: GREY25, ...CJ } },
    { text: ",", options: { fontSize: 12, fontFace: LATO, color: GREY25, ...CJ } }
  ], 1.375, 6.211, 10.583, 0.667);
  T(s, "RUBEN CONNOR", 1.182, 4.667, 1.856, 0.421, 14, PMD, GREY95, CJ);
  T(s, "JHON SMITH", 4.168, 4.682, 1.856, 0.421, 14, PMD, GREY95, CJ);
  T(s, "MIKE TRAVOLI", 7.099, 4.667, 1.856, 0.421, 14, PMD, GREY95, CJ);
  T(s, "SAMUEL DUNN", 10.137, 4.682, 1.856, 0.421, 14, PMD, GREY95, CJ);
  O(s, 2.685, 5.105, 0.315, 0.315, CYAN2);
  ICON(s, 2.72, 5.143, 0.245, 0.24, BLACK);
  O(s, 2.204, 5.098, 0.315, 0.315, CYAN2);
  ICON(s, 2.236, 5.135, 0.252, 0.242, BLACK);
  O(s, 1.264, 5.088, 0.315, 0.315, CYAN2);
  ICON(s, 1.281, 5.098, 0.264, 0.277, BLACK);
  O(s, 1.722, 5.091, 0.315, 0.312, CYAN2);
  ICON(s, 1.748, 5.088, 0.264, 0.283, BLACK);
  O(s, 5.634, 5.135, 0.315, 0.315, CYAN2);
  ICON(s, 5.669, 5.173, 0.245, 0.24, BLACK);
  O(s, 5.153, 5.128, 0.315, 0.315, CYAN2);
  ICON(s, 5.185, 5.165, 0.252, 0.242, BLACK);
  O(s, 4.213, 5.118, 0.315, 0.315, CYAN2);
  ICON(s, 4.23, 5.128, 0.264, 0.277, BLACK);
  O(s, 4.671, 5.121, 0.315, 0.312, CYAN2);
  ICON(s, 4.697, 5.118, 0.264, 0.283, BLACK);
  O(s, 8.588, 5.135, 0.315, 0.315, CYAN2);
  ICON(s, 8.623, 5.173, 0.245, 0.24, BLACK);
  O(s, 8.107, 5.128, 0.315, 0.315, CYAN2);
  ICON(s, 8.139, 5.165, 0.252, 0.242, BLACK);
  O(s, 7.167, 5.118, 0.315, 0.315, CYAN2);
  ICON(s, 7.184, 5.128, 0.264, 0.277, BLACK);
  O(s, 7.625, 5.121, 0.315, 0.312, CYAN2);
  ICON(s, 7.651, 5.118, 0.264, 0.283, BLACK);
  O(s, 11.606, 5.118, 0.315, 0.315, WHITE);
  ICON(s, 11.641, 5.156, 0.245, 0.24, BLACK);
  O(s, 11.125, 5.111, 0.315, 0.315, WHITE);
  ICON(s, 11.156, 5.148, 0.252, 0.242, BLACK);
  O(s, 10.185, 5.101, 0.315, 0.315, WHITE);
  ICON(s, 10.202, 5.111, 0.263, 0.276, BLACK);
  O(s, 10.643, 5.104, 0.315, 0.313, WHITE);
  ICON(s, 10.669, 5.101, 0.263, 0.283, BLACK);
  gradient(s, 1.045, 2.2, 0.703, 0.208, [[0, WHITE, 0], [0.84, CYAN, 0]], false);
  gradient(s, 4.022, 2.2, 0.703, 0.208, [[0, WHITE, 0], [0.84, CYAN, 0]], false);
  gradient(s, 6.935, 2.2, 0.703, 0.208, [[0, WHITE, 0], [0.84, CYAN, 0]], false);
  gradient(s, 9.99, 2.2, 0.703, 0.208, [[0, WHITE, 0], [0.84, CYAN, 0]], false);
  IMG(s, 1.019, 2.567, 2.25, 1.667);
  IMG(s, 4, 2.6, 2.25, 1.667);
  IMG(s, 6.923, 2.583, 2.25, 1.667);
  IMG(s, 9.94, 2.567, 2.25, 1.667);
}

/* --- Slide 22: BREAK SLIDE --- */
function slide22 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: '1A1A1A' };
  IMG(s, 0, 0, 13.333, 7.5);
  T(s, "BREAK SLIDE", 1.72, 2.959, 9.893, 1.582, 88, MEB, CYAN, { ...NW, charSpacing: 6 });
  T(s, "5 MINUTS TIME TO BREAK", 5.02, 4.41, 3.294, 0.303, 12, MSB, GREY25, C);
}

/* --- Slide 23: PERSONAL BRANDING --- */
function slide23 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  IMG(s, 0.5, 0.478, 4.75, 6.31);
  T(s, "PERSONAL BRANDING", 7.719, 4.897, 3.362, 0.337, 14, PSB, GREY05, { bold: true });
  T(s, "WORK EXPERIENCE", 7.719, 1.04, 3.362, 0.337, 14, PSB, GREY05, { bold: true });
  T(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Aliquet lectus proin nibh nisl condimentum id venen. Lorem ipsum dolor sit amet, con.", 7.722, 1.376, 4.312, 1.273, 12, LATO, GREY05, J);
  T(s, "CREATIVE SOLUTION", 7.736, 3.052, 3.362, 0.337, 14, PSB, GREY05, { bold: true });
  T(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Aliquet lectus proin nibh nisl condimentum id venen. Lorem ipsum dolor sit amet, con.", 7.739, 3.389, 4.312, 1.273, 12, LATO, GREY05, J);
  T(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Aliquet lectus proin nibh nisl condimentum id venen. Lorem ipsum dolor sit amet, con.", 7.719, 5.233, 4.312, 1.273, 12, LATO, GREY05, J);
  O(s, 1.265, 2.213, 3.089, 3.089, CYAN);
  DN(s, 1.081, 2.033, 3.457, 3.457, WHITE, 0.023);
  T(s, "BUSINESS TRATEGY 2022", 1.195, 2.988, 3.229, 0.707, 18, MSB, GREY05, C);
  T(s, LA(78, "."), 1.532, 3.695, 2.555, 0.97, 12, LATO, GREY05, CJ);
  DN(s, 5.75, 4.971, 1.583, 1.583, WHITE, 0.023);
  DN(s, 5.75, 2.958, 1.583, 1.583, WHITE, 0.023);
  DN(s, 5.677, 0.946, 1.583, 1.583, WHITE, 0.023);
  IMG(s, 5.782, 1.051, 1.372, 1.372);
  IMG(s, 5.856, 3.064, 1.372, 1.372);
  IMG(s, 5.856, 5.077, 1.372, 1.372);
}

/* --- Slide 24: OUR PORTFOLIO --- */
function slide24 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  IMG(s, 4.284, 2.625, 2.667, 2.917);
  IMG(s, 7.336, 0.983, 2.667, 2.917);
  IMG(s, 10.333, 0.983, 2.667, 2.917);
  IMG(s, 10.333, 4.05, 2.667, 2.917);
  IMG(s, 7.334, 4.083, 2.667, 2.917);
  R(s, 0.919, 1, 6.049, 1.333, WHITE, { shadow: shadow(20, 1, 40, GREY75, 0.4) });
  T(s, "OUR PORTFOLIO", 1.935, 1.367, 4.016, 0.572, 28, MEB, BLACK, C);
  T(s, "Title text here", 0.919, 3.317, 2.133, 0.37, 16, PSB, WHITE);
  R(s, 4.284, 5.167, 2.667, 0.375, CYAN);
  R(s, 7.334, 3.562, 2.667, 0.375, CYAN);
  R(s, 10.333, 3.517, 2.667, 0.375, CYAN);
  R(s, 7.318, 6.625, 2.667, 0.375, CYAN);
  R(s, 10.333, 6.592, 2.667, 0.375, CYAN);
  T(s, LA(212, "."), 0.919, 3.704, 2.412, 2.485, 12, LATO, GREY25, J);
  SH(s, 'homePlate', 0.498, 3.674, 0.252, 2.492, WHITE);
  T(s, "Portfolio one", 4.698, 5.188, 1.836, 0.333, 14, PSB, BLACK, C);
  T(s, "Portfolio two", 7.733, 3.562, 1.836, 0.333, 14, PSB, BLACK, C);
  T(s, "Portfolio three", 10.749, 3.538, 1.836, 0.333, 14, PSB, BLACK, C);
  T(s, "Portfolio four", 7.749, 6.646, 1.836, 0.333, 14, PSB, BLACK, C);
  T(s, "Portfolio five", 10.732, 6.612, 1.836, 0.333, 14, PSB, BLACK, C);
}

/* --- Slide 25: POD --- */
function slide25 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  // backdrop inherited from layout "25_Title Slide"
  poly(s, 5.333, -0.017, 8.013, 7.533, WHITE, [
    [0, 0.002], [1, 0], [1, 1], [0.895, 0.998], [0.88, 0.992],
    [0.72, 0.922, 0.788, 0.504, 0.568, 0.673], [0.345, 0.842, 0.117, 0.708, 0.179, 0.479],
    [0.238, 0.298, 0.467, 0.44, 0.466, 0.213], [0.462, -0.05, 0.03, 0.185, 0.001, 0.017],
    [0, 0.002]]);
  poly(s, 5.5, 0, 7.846, 7.501, CYAN2, [
    [0, 0], [1, 0], [1, 1], [0.914, 0.996], [0.898, 0.99],
    [0.735, 0.92, 0.805, 0.501, 0.58, 0.671], [0.353, 0.84, 0.119, 0.706, 0.182, 0.477],
    [0.243, 0.296, 0.477, 0.438, 0.476, 0.211], [0.472, -0.052, 0.03, 0.183, 0.001, 0.015],
    [0, 0]]);
  TX(s, [
    { text: "POD", options: { fontSize: 20, fontFace: MEB, color: CYAN, ...C } },
    { text: "SPOT PORTFOLIO", options: { fontSize: 20, fontFace: MEB, color: WHITE, ...C } }
  ], 0.017, 0.537, 4.499, 0.438);
  T(s, LA(157), 4, 1.726, 3.248, 1.273, 12, LATO, GREY25, J);
  T(s, "Title text here", 4, 1.39, 2.133, 0.337, 14, PSB, WHITE);
  T(s, LA(157), 4, 5.144, 3.248, 1.273, 12, LATO, GREY25, J);
  T(s, "Title text here", 4, 4.807, 2.133, 0.337, 14, PSB, WHITE);
  IMG(s, 5.835, 0, 7.499, 7.485);
  IMG(s, 0.542, 1.33, 3.124, 1.917);
  IMG(s, 0.537, 4.75, 3.124, 1.917);
}

/* --- Slide 26: OUR OTHER PORTFOLIO --- */
function slide26 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  IMG(s, 0.417, 0.417, 12.5, 6.667);
  T(s, LA(121, "."), 1.53, 4.229, 2.5, 1.273, 12, LATO, GREY25, J);
  T(s, "Title Text Here", 1.53, 3.926, 1.989, 0.303, 12, PSB, GREY15);
  T(s, LA(121, "."), 5.473, 4.229, 2.5, 1.273, 12, LATO, GREY25, J);
  T(s, "Title Text Here", 5.49, 3.926, 1.989, 0.303, 12, PSB, GREY15);
  T(s, LA(121, "."), 9.417, 4.229, 2.5, 1.273, 12, LATO, GREY25, J);
  T(s, "Title Text Here", 9.417, 3.926, 1.989, 0.303, 12, PSB, GREY15);
  RR(s, 1.53, 6.062, 1.378, 0.354, WHITE, 0.177);
  O(s, 1.53, 2.551, 1.083, 1.083, AMBER);
  O(s, 5.49, 2.551, 1.083, 1.083, AMBER);
  O(s, 9.45, 2.551, 1.083, 1.083, AMBER);
  RR(s, 5.473, 6.062, 1.378, 0.354, WHITE, 0.177);
  RR(s, 9.417, 6.062, 1.378, 0.354, WHITE, 0.177);
  T(s, "OUR OTHER PORTFOLIO", 3.792, 0.656, 5.749, 0.572, 28, MEB, GREY05, C);
  RR(s, 6.125, 1.617, 1.083, 0.05, WHITE, 0.008);
  ICON(s, 1.722, 2.732, 0.699, 0.699, BLACK);
  ICON(s, 9.642, 2.726, 0.699, 0.699, BLACK);
  ICON(s, 5.688, 2.75, 0.687, 0.687, BLACK);
  T(s, "Discover more", 1.537, 6.101, 1.378, 0.278, 10.5, PSB, '020202', C);
  T(s, "Discover more", 5.481, 6.095, 1.378, 0.278, 10.5, PSB, '020202', C);
  T(s, "Discover more", 9.424, 6.09, 1.378, 0.278, 10.5, PSB, '020202', C);
}

/* --- Slide 27: Lorem ipsum dolor sit amet, consectetur adip --- */
function slide27 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  IMG(s, 5.333, 1.583, 4.417, 2.833);
  device(s, 4.75, 1.445, 5.583, 3.345);
  RR(s, 1.371, 4.425, 2.215, 0.326, RED, 0.163);
  T(s, LA(284, "."), 0.619, 2.306, 3.881, 1.879, 12, LATO, GREY25, J);
  R(s, 10.5, 1.083, 2.215, 1.121, WHITE, { shadow: shadow(11, 3, 45, BLACK, 0.4) });
  R(s, 10.5, 2.537, 2.215, 1.121, RED);
  R(s, 10.5, 3.991, 2.215, 1.121, WHITE, { shadow: shadow(11, 3, 0, BLACK, 0.4) });
  poly(s, 2.167, 5.445, 1.102, 1.102, RED, [
    [0.504, 0], [0.582, 0, 0.657, 0.018, 0.723, 0.049], [0.766, 0.074], [0.802, 0.067],
    [0.855, 0.067, 0.898, 0.109, 0.898, 0.162], [0.898, 0.214, 0.855, 0.257, 0.802, 0.257],
    [0.749, 0.257, 0.707, 0.214, 0.707, 0.162], [0.713, 0.13], [0.69, 0.117],
    [0.634, 0.09, 0.571, 0.075, 0.504, 0.075], [0.268, 0.075, 0.076, 0.265, 0.076, 0.5],
    [0.076, 0.735, 0.268, 0.925, 0.504, 0.925], [0.682, 0.925, 0.834, 0.818, 0.899, 0.665],
    [0.917, 0.608], [1, 0.581], [0.998, 0.601], [0.951, 0.829, 0.748, 1, 0.504, 1],
    [0.226, 1, 0, 0.776, 0, 0.5], [0, 0.224, 0.226, 0, 0.504, 0]]);
  T(s, "STATISTIC", 6.833, 5.075, 1.755, 0.37, 16, PMD, WHITE, C);
  TX(s, [
    { text: LA(304), options: { fontSize: 12, fontFace: LATO, color: GREY25, ...J } },
    { text: ".", options: { fontSize: 12, fontFace: LATO, color: GREY25, ...J } }
  ], 3.667, 5.445, 7.654, 0.97);
  T(s, "www.podspot.com", 1.58, 4.445, 1.797, 0.303, 12, PMD, BLACK, NW);
  T(s, "Devices Mockup", 0.619, 1.083, 2.131, 0.337, 14, PMD, WHITE);
  TX(s, [
    { text: "NOW YOU’RE THINKING TOGETHER WE ", options: { fontSize: 20, fontFace: MEB, color: WHITE } },
    { text: "CREATE", options: { fontSize: 20, fontFace: MEB, color: CYAN2 } }
  ], 0.619, 1.476, 3.833, 0.774);
  T(s, "15", 10.601, 1.239, 0.384, 0.337, 14, PMD, BLACK, NW);
  T(s, "Years In advertising", 10.583, 1.57, 2.151, 0.337, 14, PMD, BLACK, NW);
  T(s, "13", 10.594, 2.721, 0.379, 0.337, 14, PMD, BLACK, NW);
  TX(s, [
    { text: "Global", options: { fontSize: 14, fontFace: PMD, color: BLACK, breakLine: true } },
    { text: "Advertising Award", options: { fontSize: 14, fontFace: PMD, color: BLACK } }
  ], 10.55, 2.994, 2.151, 0.572);
  T(s, "2500", 10.5, 4.101, 0.682, 0.337, 14, PMD, BLACK, NW);
  T(s, "Successful Campaigns Executed", 10.477, 4.41, 2.364, 0.572, 14, PMD, BLACK);
  T(s, "85 %", 2.31, 5.778, 1.102, 0.438, 20, PMD, CYAN);
}

/* --- Slide 28: PROVIDING SIMPLEST SOLUTION --- */
function slide28 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  // backdrop inherited from layout "28_Title Slide"
  O(s, 1.092, 2.648, 2.244, 2.205, WHITE, { transparency: 65, line: { color: WHITE, width: 1 } });
  O(s, 9.997, 2.648, 2.244, 2.205, WHITE, { transparency: 65, line: { color: WHITE, width: 1 } });
  IMG(s, 4.458, 2.396, 4.417, 2.708);
  device(s, 4.25, 2.186, 4.833, 4.397);
  T(s, "PROVIDING SIMPLEST SOLUTION", 2.604, 0.609, 8.125, 0.505, 24, MBK, WHITE, { ...C, bold: true });
  T(s, LA(123), 2.859, 1.173, 7.616, 0.667, 12, LATO, GREY05, CJ);
  T(s, "DIGITAL ANALITIC", 1.055, 5.104, 2.319, 0.337, 14, PSB, GREY05, C);
  T(s, LA(89, "."), 0.811, 5.559, 2.807, 0.97, 12, LATO, GREY05, CJ);
  T(s, LA(89, "."), 9.652, 5.441, 2.935, 0.97, 12, LATO, GREY05, CJ);
  T(s, "DIGITAL MARKETING", 9.574, 5.104, 3.09, 0.337, 14, PSB, WHITE, C);
  IMG(s, 10.308, 2.939, 1.622, 1.622);
  IMG(s, 1.403, 2.939, 1.623, 1.623);
}

/* --- Slide 29: PHONE MOCKUP --- */
function slide29 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  // backdrop inherited from layout "29_Title Slide"
  poly(s, 0, 0, 7.417, 7.5, WHITE, [[0, 0], [1, 0], [0.51, 0.998], [0, 1], [0, 0]]);
  IMG(s, 3.52, 3.221, 2.494, 4.994);
  IMG(s, 1.175, -0.754, 2.494, 3.871);
  device(s, 3.35, 3.109, 2.839, 5.23, [[0, 0], [1, 0], [1, 0.748], [0.066, 1], [0, 1]], { rotate: 26.462 });
  device(s, 0.999, -0.885, 2.839, 4.189, [[0, 0.337], [1, 0], [1, 1], [0, 1]], { rotate: 26.462 });
  SH(s, 'flowChartInputOutput', 4.316, 0.351, 4.392, 1.867, AMBER);
  T(s, "PHONE MOCKUP", 8.333, 2.422, 4.167, 0.505, 24, MEB, WHITE);
  T(s, LA(386), 8.333, 3.167, 3.881, 2.485, 12, LATO, GREY15, J);
  poly(s, 0, 4.25, 3.146, 1.667, CYAN, [[0, 0], [1, 0], [0.721, 1], [0, 1]]);
  T(s, "CEO Mode Easy", 5.552, 0.5, 2.079, 0.337, 14, PSB, BLACK);
  T(s, "Creative Writing", 0.284, 4.409, 2.382, 0.337, 14, PSB, BLACK);
  T(s, LA(89, "."), 5.261, 0.89, 2.502, 0.97, 12, LATO, GREY75, J);
  TX(s, [
    { text: LA(55), options: { fontSize: 12, fontFace: LATO, color: GREY75, ...J } },
    { text: ".", options: { fontSize: 12, fontFace: LATO, color: GREY75, ...J } }
  ], 0.284, 4.734, 1.736, 0.97);
}

/* --- Slide 30: PODSPOT INFOGRAPHIC --- */
function slide30 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  DN(s, 5.04, 2.543, 3.616, 3.616, WHITE, 0.266);
  poly(s, 5.407, 2, 2.331, 1.736, CYAN, [
    [0.618, 0], [0.723, 0, 0.824, 0.021, 0.92, 0.061], [1, 0.101], [0.772, 0.897],
    [0.766, 0.893], [0.72, 0.867, 0.67, 0.853, 0.618, 0.853],
    [0.54, 0.853, 0.467, 0.885, 0.407, 0.939], [0.352, 1], [0, 0.284], [0.052, 0.232],
    [0.213, 0.086, 0.408, 0, 0.618, 0]]);
  poly(s, 4.508, 4.498, 1.739, 1.303, CYAN, [
    [0.845, 0], [0.848, 0.033], [0.869, 0.165, 0.918, 0.283, 0.987, 0.375], [1, 0.389],
    [0.269, 1], [0.219, 0.91], [0.121, 0.717, 0.051, 0.498, 0.015, 0.261], [0, 0.134],
    [0.845, 0]]);
  poly(s, 7.551, 2.72, 1.643, 1.682, CYAN, [
    [0.59, 0], [0.631, 0.041], [0.816, 0.241, 0.944, 0.493, 0.988, 0.772], [1, 0.877],
    [0.103, 1], [0.103, 0.986], [0.103, 0.873, 0.067, 0.768, 0.004, 0.682], [0, 0.677],
    [0.59, 0]]);
  poly(s, 6.822, 5.106, 1.457, 1.595, CYAN, [
    [0.342, 0], [1, 0.71], [0.924, 0.762], [0.751, 0.868, 0.555, 0.945, 0.344, 0.985],
    [0.235, 1], [0, 0.085], [0.018, 0.087], [0.135, 0.087, 0.245, 0.056, 0.338, 0.003],
    [0.342, 0]]);
  poly(s, 4.264, 2, 2.155, 3.067, 'FFD966', [
    [0.685, 0], [1, 0.521], [0.985, 0.526], [0.874, 0.579, 0.802, 0.668, 0.802, 0.769],
    [0.802, 0.799, 0.808, 0.828, 0.82, 0.856], [0.82, 0.856], [0.046, 1], [0.025, 0.941],
    [0.009, 0.886, 0, 0.828, 0, 0.769], [0, 0.444, 0.258, 0.161, 0.637, 0.016], [0.685, 0]], { transparency: 9 });
  poly(s, 7.285, 4.513, 2.117, 2.003, 'FFD966', [
    [0.226, 0], [1, 0.277], [0.96, 0.391], [0.866, 0.626, 0.709, 0.826, 0.511, 0.967],
    [0.46, 1], [0, 0.284], [0.055, 0.252], [0.134, 0.196, 0.194, 0.112, 0.223, 0.014],
    [0.226, 0]], { rotate: 357.041, transparency: 9 });
  T(s, "INFOGRAPHIC  01", 1.165, 1.878, 2.163, 0.337, 14, MSB, WHITE, RT);
  T(s, "INFOGRAPHIC  02", 1.165, 4.371, 2.163, 0.337, 14, MSB, WHITE, RT);
  T(s, "INFOGRAPHIC  03", 9.862, 1.87, 2.163, 0.337, 14, MSB, WHITE);
  T(s, "INFOGRAPHIC  04", 9.862, 4.421, 2.163, 0.337, 14, MSB, WHITE);
  T(s, "Lorem ipsum dolor sitamet, consectetur adipiscing elit, seddo eiusmod tempor incididunt ut labore et dolor.", 1.027, 2.249, 2.301, 1.273, 12, LATO, GREY50, J);
  T(s, "Lorem ipsum dolor sitamet, consectetur adipiscing elit, seddo eiusmod tempor incididunt ut labore et dolor.", 1.027, 4.794, 2.301, 1.273, 12, LATO, GREY50, J);
  T(s, "Lorem ipsum dolor sitamet, consectetur adipiscing elit, seddo eiusmod tempor incididunt ut labore et dolor.", 9.862, 2.262, 2.301, 1.273, 12, LATO, GREY50, J);
  T(s, "Lorem ipsum dolor sitamet, consectetur adipiscing elit, seddo eiusmod tempor incididunt ut labore et dolor.", 9.862, 4.806, 2.301, 1.273, 12, LATO, GREY50, J);
  T(s, "PODSPOT INFOGRAPHIC", 2.764, 0.5, 7.805, 0.707, 36, MEB, WHITE, C);
}

/* --- Slide 31: PODSPOT INFOGRAPHIC --- */
function slide31 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  RR(s, 1.812, 5.417, 9.708, 0.5, CYAN, 0.25);
  RR(s, 1.804, 5.615, 9.708, 0.079, AMBER, 0.04);
  O(s, 2.588, 2.115, 0.669, 0.669, AMBER);
  O(s, 5.177, 2.133, 0.669, 0.669, AMBER);
  O(s, 7.713, 2.115, 0.669, 0.669, AMBER);
  O(s, 10.043, 2.1, 0.669, 0.669, AMBER);
  O(s, 2.622, 5.399, 0.512, 0.512, WHITE);
  O(s, 5.268, 5.405, 0.512, 0.512, WHITE);
  O(s, 7.87, 5.405, 0.512, 0.512, WHITE);
  O(s, 10.155, 5.399, 0.512, 0.512, WHITE);
  T(s, "PODSPOT INFOGRAPHIC", 3.638, 0.668, 6.058, 0.572, 28, MEB, WHITE, C);
  T(s, "01", 2.588, 2.251, 0.669, 0.404, 18, MSB, BLACK, C);
  T(s, "02", 5.172, 2.291, 0.669, 0.404, 18, MSB, BLACK, C);
  T(s, "03", 7.707, 2.265, 0.669, 0.404, 18, MSB, BLACK, C);
  T(s, "04", 10.054, 2.233, 0.669, 0.404, 18, MSB, BLACK, C);
  T(s, "INFOGRAPHIC 1", 1.937, 2.934, 1.97, 0.337, 14, PSB, WHITE, C);
  T(s, "INFOGRAPHIC 2", 4.492, 2.934, 2.06, 0.337, 14, PSB, WHITE, C);
  T(s, "INFOGRAPHIC 3", 6.874, 2.934, 2.27, 0.337, 14, PSB, WHITE, C);
  T(s, "INFOGRAPHIC 4", 9.372, 2.934, 1.991, 0.337, 14, PSB, WHITE, C);
  T(s, LA(68, "."), 1.731, 3.235, 2.379, 0.97, 12, LATO, GREY25, CJ);
  T(s, LA(68, "."), 4.317, 3.235, 2.379, 0.97, 12, LATO, GREY25, CJ);
  T(s, LA(68, "."), 6.855, 3.235, 2.379, 0.97, 12, LATO, GREY25, CJ);
  T(s, LA(68, "."), 9.223, 3.216, 2.379, 0.97, 12, LATO, GREY25, CJ);
  LN(s, 2.879, 4.32, 0, 1.142, WHITE, 2.25);
  LN(s, 5.506, 4.32, 0, 1.142, WHITE, 2.25);
  LN(s, 8.136, 4.399, 0, 1.142, WHITE, 2.25);
  LN(s, 10.415, 4.32, 0, 1.142, WHITE, 2.25);
  O(s, 2.788, 4.274, 0.167, 0.167, CYAN);
  O(s, 5.423, 4.316, 0.167, 0.167, CYAN);
  O(s, 8.058, 4.358, 0.167, 0.167, CYAN);
  O(s, 10.327, 4.315, 0.167, 0.167, CYAN);
  T(s, LA(216), 2.092, 6.28, 9.15, 0.667, 12, LATO, GREY25, CJ);
}

/* --- Slide 32: PODSPOT INFOGRAPHIC --- */
function slide32 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  LN(s, 9.013, 4.848, 0.801, 0, '4472C4', 2.25);
  LN(s, 3.614, 4.822, 0.741, 0.012, 'FFCF5C', 2.25, { flipV: true });
  T(s, "PODSPOT INFOGRAPHIC", 2.809, 0.506, 7.715, 0.707, 36, MSB, WHITE, C);
  T(s, "INFOGRAPHIC  01", 0.972, 2.11, 2.163, 0.337, 14, PSB, WHITE, RT);
  T(s, "INFOGRAPHIC  02", 0.972, 4.603, 2.163, 0.337, 14, PSB, WHITE, RT);
  T(s, "INFOGRAPHIC  03", 10.115, 2.103, 2.163, 0.337, 14, PSB, WHITE);
  T(s, "INFOGRAPHIC  04", 10.115, 4.653, 2.163, 0.337, 14, PSB, WHITE);
  T(s, "Lorem ipsum dolor sitamet, consectetur adipiscing elit, seddo eiusmod tempor incididunt ut labore et dolor.", 0.833, 2.481, 2.301, 1.273, 12, LATO, GREY50, J);
  T(s, "Lorem ipsum dolor sitamet, consectetur adipiscing elit, seddo eiusmod tempor incididunt ut labore et dolor.", 0.833, 5.026, 2.301, 1.273, 12, LATO, GREY50, J);
  T(s, "Lorem ipsum dolor sitamet, consectetur adipiscing elit, seddo eiusmod tempor incididunt ut labore et dolor.", 10.115, 2.494, 2.301, 1.273, 12, LATO, GREY50, J);
  T(s, "Lorem ipsum dolor sitamet, consectetur adipiscing elit, seddo eiusmod tempor incididunt ut labore et dolor.", 10.115, 5.039, 2.301, 1.273, 12, LATO, GREY50, J);
  poly(s, 4.228, 2.115, 1.903, 2.092, 'F6366D', [
    [0.631, 0], [1, 0.606], [0.994, 0.609], [0.866, 0.691, 0.778, 0.821, 0.762, 0.97],
    [0.76, 1], [0, 1], [0.006, 0.899], [0.045, 0.546, 0.253, 0.238, 0.556, 0.044],
    [0.631, 0]]);
  poly(s, 7.322, 2.103, 1.912, 2.105, CYAN, [
    [0.362, 0], [0.446, 0.05], [0.748, 0.243, 0.955, 0.549, 0.994, 0.899], [1, 1],
    [0.243, 1], [0.242, 0.97], [0.225, 0.822, 0.138, 0.693, 0.011, 0.612], [0, 0.606],
    [0.362, 0]]);
  poly(s, 4.226, 4.266, 2.47, 2.484, 'CC00CC', [
    [0, 0], [0.586, 0], [0.588, 0.037], [0.608, 0.238, 0.769, 0.397, 0.971, 0.416],
    [1, 0.418], [1, 1], [0.915, 0.996], [0.435, 0.951, 0.054, 0.573, 0.005, 0.097], [0, 0]]);
  poly(s, 6.755, 4.266, 2.47, 2.484, '002F7D', [
    [0.414, 0], [1, 0], [0.995, 0.097], [0.946, 0.573, 0.565, 0.951, 0.085, 0.996], [0, 1],
    [0, 0.418], [0.029, 0.416], [0.231, 0.397, 0.392, 0.238, 0.412, 0.037], [0.414, 0]]);
  DN(s, 5.913, 3.439, 1.625, 1.625, '00B365', 0.098);
  poly(s, 4.918, 2.71, 3.615, 3.334, GREY85, [
    [0.767, 0], [0.78, 0.008], [0.913, 0.106, 1, 0.271, 1, 0.458],
    [1, 0.757, 0.776, 1, 0.5, 1], [0.224, 1, 0, 0.757, 0, 0.458],
    [0, 0.271, 0.087, 0.106, 0.22, 0.008], [0.229, 0.002], [0.335, 0.204], [0.297, 0.238],
    [0.245, 0.294, 0.213, 0.372, 0.213, 0.458], [0.213, 0.629, 0.342, 0.769, 0.5, 0.769],
    [0.658, 0.769, 0.787, 0.629, 0.787, 0.458], [0.787, 0.372, 0.755, 0.294, 0.703, 0.238],
    [0.663, 0.202], [0.767, 0]], { transparency: 75 });
  LN(s, 8.226, 2.333, 1.575, 0, CYAN, 2.25);
  LN(s, 3.568, 2.338, 1.575, 0, 'FF6464', 2.25);
  O(s, 3.412, 2.235, 0.197, 0.197, 'FF6464');
  O(s, 3.428, 4.735, 0.197, 0.197, 'FFCF5C');
  O(s, 9.716, 2.23, 0.197, 0.197, CYAN);
  O(s, 9.724, 4.75, 0.197, 0.197, '0070C0');
}

/* --- Slide 33: PRICING PLAN --- */
function slide33 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  R(s, 1.5, 1.833, 3, 5, CYAN);
  SH(s, 'hexagon', 2.235, 2.812, 1.318, 1.175, WHITE, { rotate: 90 });
  R(s, 5.199, 1.849, 3, 5, WHITE);
  R(s, 8.833, 1.833, 3, 5, 'FF9933');
  T(s, "PRICING PLAN", 4.304, 0.486, 4.719, 0.707, 36, MEB, WHITE, C);
  T(s, "Basic Plan", 2.083, 2.097, 1.706, 0.37, 16, PSB, INK, C);
  T(s, "Standar Plan", 5.667, 2.097, 2, 0.37, 16, PSB, BLACK, C);
  T(s, "Advance Plan", 9.333, 2.097, 2, 0.37, 16, PSB, INK, C);
  SH(s, 'hexagon', 6.008, 2.754, 1.318, 1.175, WHITE, { rotate: 90 });
  T(s, "350$", 2.278, 3, 1.317, 0.438, 20, PSB, INK, C);
  T(s, "/Month", 2.411, 3.399, 1.051, 0.338, 14, POP, INK, C);
  T(s, "450$", 6.008, 2.975, 1.317, 0.438, 20, PSB, INK, C);
  T(s, "/Month", 6.141, 3.374, 1.051, 0.338, 14, POP, INK, C);
  SH(s, 'hexagon', 9.676, 2.754, 1.318, 1.175, WHITE, { rotate: 90 });
  T(s, "750$", 9.677, 2.975, 1.317, 0.438, 20, PSB, INK, C);
  T(s, "/Month", 9.81, 3.374, 1.051, 0.338, 14, POP, INK, C);
  T(s, "Your Service List 1", 2.096, 4.269, 2.083, 0.337, 14, POP, INK);
  T(s, "Your Service List 2", 2.096, 4.737, 2.083, 0.337, 14, POP, INK);
  T(s, "Your Service List 3", 2.096, 5.142, 2.083, 0.337, 14, POP, INK);
  T(s, "Your Service List 4", 2.096, 5.574, 2.083, 0.337, 14, POP, INK);
  RR(s, 2.096, 6.167, 1.681, 0.464, WHITE, 0.232);
  T(s, "Get Started", 1.883, 6.262, 2.083, 0.337, 14, POP, INK, C);
  RR(s, 5.866, 6.167, 1.468, 0.464, '3BABFF', 0.173);
  T(s, "Get Started", 5.561, 6.262, 2.083, 0.337, 14, POP, INK, C);
  RR(s, 9.558, 6.167, 1.435, 0.464, WHITE, 0.232);
  T(s, "Get Started", 9.25, 6.218, 2.083, 0.337, 14, POP, INK, C);
  SH(s, 'snip2SameRect', 2.108, 1.715, 1.595, 0.118, AMBER);
  SH(s, 'snip2SameRect', 5.866, 1.739, 1.595, 0.118, WHITE);
  SH(s, 'snip2SameRect', 9.558, 1.731, 1.595, 0.118, CYAN);
  SH(s, 'hexagon', 1.88, 4.344, 0.157, 0.157, INK);
  SH(s, 'hexagon', 1.877, 4.765, 0.157, 0.157, INK);
  SH(s, 'hexagon', 1.877, 5.217, 0.157, 0.157, INK);
  SH(s, 'hexagon', 1.877, 5.653, 0.157, 0.157, INK);
  T(s, "Your Service List 1", 5.866, 4.232, 2.083, 0.337, 14, POP, GREY95);
  T(s, "Your Service List 2", 5.866, 4.7, 2.083, 0.337, 14, POP, GREY95);
  T(s, "Your Service List 3", 5.866, 5.105, 2.083, 0.337, 14, POP, GREY95);
  T(s, "Your Service List 4", 5.866, 5.537, 2.083, 0.337, 14, POP, GREY95);
  SH(s, 'hexagon', 5.649, 4.307, 0.157, 0.157, INK);
  SH(s, 'hexagon', 5.646, 4.728, 0.157, 0.157, INK);
  SH(s, 'hexagon', 5.646, 5.18, 0.157, 0.157, INK);
  SH(s, 'hexagon', 5.646, 5.616, 0.157, 0.157, INK);
  T(s, "Your Service List 1", 9.531, 4.269, 2.083, 0.337, 14, POP, INK);
  T(s, "Your Service List 2", 9.531, 4.737, 2.083, 0.337, 14, POP, INK);
  T(s, "Your Service List 3", 9.531, 5.142, 2.083, 0.337, 14, POP, INK);
  T(s, "Your Service List 4", 9.531, 5.574, 2.083, 0.337, 14, POP, INK);
  SH(s, 'hexagon', 9.314, 4.344, 0.157, 0.157, INK);
  SH(s, 'hexagon', 9.311, 4.765, 0.157, 0.157, INK);
  SH(s, 'hexagon', 9.311, 5.217, 0.157, 0.157, INK);
  SH(s, 'hexagon', 9.311, 5.653, 0.157, 0.157, INK);
}

/* --- Slide 34: TESTIMONIAL --- */
function slide34 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  DN(s, 9.244, 3.619, 2.48, 2.48, '2FC9FF', 0.079);
  DN(s, 5.427, 3.677, 2.48, 2.48, '2FC9FF', 0.079);
  DN(s, 1.609, 3.619, 2.48, 2.48, '2FC9FF', 0.079);
  T(s, "WHAT OUR CUSTOMER SAY", 4.831, 1.192, 3.672, 0.337, 14, MMD, WHITE, C);
  T(s, LA(124, "Diam quis enim lobortis."), 1.201, 1.833, 3.296, 1.273, 12, LATO, GREY05, { ...CJ, italic: true });
  T(s, LA(124, "Diam quis enim lobortis."), 4.995, 1.833, 3.296, 1.273, 12, LATO, CYAN, { ...CJ, italic: true });
  T(s, LA(124, "Diam quis enim lobortis."), 8.788, 1.84, 3.296, 1.273, 12, LATO, GREY05, { ...CJ, italic: true });
  T(s, "SAMUEL SAM", 2.089, 6.269, 1.521, 0.314, 12, PSB, WHITE, C);
  T(s, "LOUISE MARKER", 5.75, 6.28, 1.833, 0.303, 12, PSB, CYAN, C);
  T(s, "ROBERTO RIGEL", 9.651, 6.28, 1.833, 0.303, 12, PSB, WHITE, C);
  T(s, "TESTIMONIAL", 4.064, 0.553, 5.18, 0.64, 32, MEB, WHITE, { ...C, charSpacing: 6 });
  IMG(s, 1.849, 3.859, 2, 2);
  IMG(s, 5.667, 3.917, 2, 2);
  IMG(s, 9.484, 3.859, 2, 2);
}

/* --- Slide 35: Contact us --- */
function slide35 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  T(s, "Contact us", 2.332, 0.526, 3.085, 0.606, 30, MEB, WHITE, { ...C, bold: true });
  TX(s, [
    { text: "Please feel free to call us on (phone cell) or contact us by (email),", options: { fontSize: 14, fontFace: LATO, color: GREY05, ...CJ, breakLine: true } },
    { text: "if you require any further information.", options: { fontSize: 14, fontFace: LATO, color: GREY05, ...CJ } }
  ], 0.703, 1.249, 6.333, 0.761);
  T(s, "OFFICE HOURS", 1.32, 2.397, 1.596, 0.286, 11, POP, WHITE, { bold: true, charSpacing: 1 });
  TX(s, [
    { text: "Monday – Thursday", options: { fontSize: 11, fontFace: LATO, color: GREY15, ...LS, breakLine: true } },
    { text: "08:00 – 17:00", options: { fontSize: 11, fontFace: LATO, color: GREY15, ...LS } }
  ], 1.32, 2.686, 2.262, 0.62);
  TX(s, [
    { text: "(+62) 8123 4567 890", options: { fontSize: 11, fontFace: LATO, color: GREY15, ...LS, breakLine: true } },
    { text: "(0254) 50123", options: { fontSize: 11, fontFace: LATO, color: GREY15, ...LS } }
  ], 1.322, 4.117, 2.262, 0.62);
  T(s, "OUR ADDRESS", 4.583, 2.397, 1.75, 0.286, 11, POP, WHITE, { bold: true, charSpacing: 1 });
  TX(s, [
    { text: "Rakata Parkway", options: { fontSize: 11, fontFace: LATO, color: GREY15, ...LS, breakLine: true } },
    { text: "Chilegon View, IDN 1234", options: { fontSize: 11, fontFace: LATO, color: GREY15, ...LS } }
  ], 4.583, 2.705, 2.262, 0.62);
  T(s, "GET IN TOUCH", 1.322, 3.827, 1.595, 0.285, 11, POP, WHITE, { ...NW, bold: true, charSpacing: 1 });
  TX(s, [
    { text: "www.podspotcompany.com", options: { fontSize: 11, fontFace: LATO, color: GREY15, ...LS, breakLine: true } },
    { text: "office@myexample.com", options: { fontSize: 11, fontFace: LATO, color: GREY15, ...LS } }
  ], 4.583, 4.113, 2.262, 0.62);
  T(s, "FOLLOW US", 4.583, 3.823, 1.327, 0.306, 11, POP, WHITE, { ...NW, bold: true, charSpacing: 1 });
}

/* --- Slide 36: THANK YOU --- */
function slide36 (pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  IMG(s, 0, 0, 13.333, 7.5);
  R(s, 0, 2.333, 13.333, 2.833, CYAN, { transparency: 25 });
  T(s, "THANK YOU", 2.617, 2.959, 8.1, 1.582, 88, MEB, BLACK, NW);
  T(s, "FOR WATCHING", 8.667, 4.356, 1.922, 0.37, 16, PSB, BLACK, CNW);
}

/* ------------------------------------------------------------- build ------ */
const BUILDERS = [
  slide01,
  slide02,
  slide03,
  slide04,
  slide05,
  slide06,
  slide07,
  slide08,
  slide09,
  slide10,
  slide11,
  slide12,
  slide13,
  slide14,
  slide15,
  slide16,
  slide17,
  slide18,
  slide19,
  slide20,
  slide21,
  slide22,
  slide23,
  slide24,
  slide25,
  slide26,
  slide27,
  slide28,
  slide29,
  slide30,
  slide31,
  slide32,
  slide33,
  slide34,
  slide35,
  slide36
];

function build () {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE_13x75', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE_13x75';
  pptx.author = 'PODSPOT';
  pptx.title = 'PODSPOT - Podcast and eCourse Presentation Template';
  BUILDERS.forEach(fn => fn(pptx));
  return pptx.writeFile({ fileName: path.join(__dirname, '133dc2c7-caf0-4224-bc55-68607fd21efb_grok_final.pptx') });
}

build().then(f => console.log('wrote ' + f)).catch(err => { console.error(err); process.exit(1); });
