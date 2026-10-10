/*
 * "School Education Infograph" - recreated with pptxgenjs.
 * 20 slides, 13.333 x 7.5 in (16:9).
 * The raster/SVG artwork of the original deck is redrawn with native shapes.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const C = {
  navy: '2C2C7E',    // accent1 / tx2
  blue: '514DF2',    // accent2
  green: '02BA81',   // accent3
  yellow: 'FFDB01',  // accent4
  red: 'F25E6B',     // accent5
  navyDk: '21215E',
  blueDk: '1510DF',
  greenDk: '018B61',
  yellowDk: 'C0A500',
  redDk: 'E91326',
  bg: 'F2F2F2',      // lt2 - slide background
  white: 'FFFFFF',
  black: '000000',
  ink: '404040',
  slate: '595959',
  gray: '808080',    // tx1 lumMod 50%
  track: 'D9D9D9',   // bg1 lumMod 85%
  soft: 'ECECEC',    // faint background flowers
  paper: 'F7F7F8',
  gold: 'F4D05E',    // doodles / sparkles on the dark title slide
  star: 'FFF653'     // big soft starburst
};

const F = { head: 'Nunito Medium', body: 'Poppins Light' };

const CARD_SHADOW = { type: 'outer', blur: 18, offset: 5, angle: 50, color: C.black, opacity: 0.07 };

const LOREM = 'Sed ut perspiciatis unde';
const LOREM_LONG = 'Sed ut perspiciatis unde omnis iste natus';
const LOREM_SHORT = 'Sed ut perspiciatis';

/* --------------------------------------------------------------- geometry */

// Path helper. Every point is a fraction of the shape box:
//   [x, y]                            -> line (or move for the first point)
//   ['c', c1x, c1y, c2x, c2y, x, y]   -> cubic bezier
//   ['m', x, y]                       -> explicit move (new sub-path)
//   ['z']                             -> close
function poly(slide, x, y, w, h, pts, opts) {
  const points = [];
  pts.forEach((p, i) => {
    if (p[0] === 'z') points.push({ close: true });
    else if (p[0] === 'm') points.push({ x: w * p[1], y: h * p[2], moveTo: true });
    else if (p[0] === 'c') points.push({
      x: w * p[5], y: h * p[6],
      curve: { type: 'cubic', x1: w * p[1], y1: h * p[2], x2: w * p[3], y2: h * p[4] }
    });
    else points.push({ x: w * p[0], y: h * p[1], moveTo: i === 0 });
  });
  slide.addShape('custGeom', Object.assign({ x, y, w, h, points }, opts));
}

function fillPoly(slide, x, y, w, h, pts, color, opts) {
  poly(slide, x, y, w, h, pts.concat([['z']]), Object.assign({ fill: { color } }, opts));
}

// Straight stroked segment given in absolute inches.
function seg(slide, x1, y1, x2, y2, color, ptWidth) {
  const x = Math.min(x1, x2), y = Math.min(y1, y2);
  const w = Math.max(Math.abs(x2 - x1), 0.002), h = Math.max(Math.abs(y2 - y1), 0.002);
  poly(slide, x, y, w, h, [[(x1 - x) / w, (y1 - y) / h], [(x2 - x) / w, (y2 - y) / h]],
    { line: { color, width: ptWidth } });
}

function rect(slide, x, y, w, h, color, opts) {
  slide.addShape('rect', Object.assign({ x, y, w, h, fill: { color } }, opts));
}

function pill(slide, x, y, w, h, color, opts) {
  slide.addShape('roundRect', Object.assign(
    { x, y, w, h, fill: { color }, rectRadius: Math.min(w, h) / 2 }, opts));
}

function oval(slide, x, y, w, h, color, opts) {
  slide.addShape('ellipse', Object.assign({ x, y, w, h, fill: { color } }, opts));
}

function tri(slide, x, y, w, h, color, opts) {
  fillPoly(slide, x, y, w, h, [[0.5, 0], [1, 1], [0, 1]], color, opts);
}

/* ------------------------------------------------------------ decorations */

// The three yellow dashes that sit beside every section title.
function sparkle(slide, x, y, w, h, rot, color) {
  const o = { rotate: rot || 0 }, c = color || C.yellow;
  fillPoly(slide, x, y, w, h, [[0.104, 0.501], [0.000, 0.439],
    ['c', 0.15, 0.30, 0.30, 0.15, 0.430, 0.000], [0.543, 0.053],
    ['c', 0.41, 0.21, 0.26, 0.36, 0.104, 0.501]], c, o);
  fillPoly(slide, x, y, w, h, [[0.439, 0.642], [0.365, 0.560],
    ['c', 0.55, 0.47, 0.74, 0.39, 0.930, 0.320], [0.993, 0.404],
    ['c', 0.80, 0.47, 0.62, 0.55, 0.439, 0.642]], c, o);
  fillPoly(slide, x, y, w, h, [[0.315, 0.844], [0.283, 0.940], [0.603, 0.998], [0.635, 0.902]], c, o);
}

// Hand-drawn yellow doodles: 'loop' has two curls, 'wave' is a long sweep.
const CURL_LOOP = [
  [0.031, 0.997],
  ['c', -0.01, 0.89, -0.01, 0.78, 0.03, 0.69], ['c', 0.06, 0.63, 0.11, 0.59, 0.16, 0.57],
  ['c', 0.16, 0.50, 0.17, 0.42, 0.21, 0.36], ['c', 0.25, 0.29, 0.32, 0.23, 0.40, 0.22],
  ['c', 0.42, 0.20, 0.43, 0.18, 0.44, 0.17], ['c', 0.51, 0.07, 0.62, 0.01, 0.74, 0.00],
  ['c', 0.79, 0.00, 0.86, 0.00, 0.92, 0.05], ['c', 0.98, 0.10, 1.01, 0.18, 1.00, 0.25],
  [0.939, 0.236],
  ['c', 0.95, 0.19, 0.93, 0.14, 0.89, 0.10], ['c', 0.84, 0.07, 0.78, 0.06, 0.74, 0.07],
  ['c', 0.64, 0.07, 0.55, 0.13, 0.48, 0.21], ['c', 0.48, 0.22, 0.48, 0.22, 0.47, 0.23],
  ['c', 0.52, 0.24, 0.56, 0.29, 0.58, 0.34], ['c', 0.60, 0.39, 0.60, 0.45, 0.58, 0.49],
  ['c', 0.56, 0.52, 0.54, 0.53, 0.52, 0.54], ['c', 0.47, 0.57, 0.41, 0.52, 0.39, 0.47],
  ['c', 0.36, 0.41, 0.36, 0.35, 0.38, 0.29], ['c', 0.33, 0.31, 0.29, 0.35, 0.26, 0.39],
  ['c', 0.23, 0.44, 0.22, 0.50, 0.22, 0.56], ['c', 0.24, 0.56, 0.26, 0.56, 0.29, 0.57],
  ['c', 0.34, 0.58, 0.38, 0.63, 0.40, 0.68], ['c', 0.42, 0.74, 0.42, 0.79, 0.40, 0.84],
  ['c', 0.40, 0.85, 0.39, 0.86, 0.38, 0.87], ['c', 0.36, 0.89, 0.33, 0.89, 0.29, 0.87],
  ['c', 0.23, 0.82, 0.18, 0.74, 0.16, 0.65], ['c', 0.16, 0.65, 0.16, 0.64, 0.16, 0.64],
  ['c', 0.13, 0.66, 0.10, 0.69, 0.09, 0.72], ['c', 0.05, 0.79, 0.05, 0.89, 0.08, 0.97], ['z'],
  ['m', 0.221, 0.624],
  ['c', 0.24, 0.71, 0.27, 0.77, 0.33, 0.81], ['c', 0.35, 0.82, 0.35, 0.82, 0.35, 0.81],
  ['c', 0.36, 0.78, 0.36, 0.75, 0.35, 0.71], ['c', 0.33, 0.67, 0.30, 0.64, 0.27, 0.63],
  ['c', 0.25, 0.62, 0.24, 0.62, 0.22, 0.62], ['z'],
  ['m', 0.440, 0.288],
  ['c', 0.42, 0.34, 0.42, 0.39, 0.44, 0.44], ['c', 0.45, 0.47, 0.48, 0.49, 0.50, 0.48],
  ['c', 0.51, 0.48, 0.52, 0.47, 0.53, 0.46], ['c', 0.55, 0.42, 0.54, 0.39, 0.53, 0.37],
  ['c', 0.51, 0.33, 0.48, 0.30, 0.45, 0.29], ['c', 0.45, 0.29, 0.44, 0.29, 0.44, 0.29], ['z']
];
const CURL_WAVE = [
  [0.243, 0.997],
  ['c', 0.16, 1.00, 0.08, 0.98, 0.00, 0.94], [0.015, 0.890],
  ['c', 0.24, 1.00, 0.52, 0.93, 0.69, 0.72], ['c', 0.75, 0.65, 0.79, 0.58, 0.81, 0.50],
  ['c', 0.80, 0.50, 0.79, 0.49, 0.78, 0.49], ['c', 0.76, 0.46, 0.74, 0.43, 0.74, 0.39],
  ['c', 0.75, 0.36, 0.77, 0.32, 0.79, 0.32], ['c', 0.81, 0.31, 0.83, 0.31, 0.85, 0.33],
  ['c', 0.85, 0.35, 0.86, 0.37, 0.86, 0.38], ['c', 0.86, 0.41, 0.86, 0.43, 0.86, 0.46],
  ['c', 0.86, 0.46, 0.87, 0.46, 0.88, 0.45], ['c', 0.93, 0.42, 0.94, 0.34, 0.95, 0.30],
  ['c', 0.96, 0.20, 0.96, 0.11, 0.93, 0.02], [0.973, 0.000],
  ['c', 1.00, 0.10, 1.01, 0.20, 0.99, 0.31], ['c', 0.98, 0.37, 0.96, 0.46, 0.90, 0.50],
  ['c', 0.88, 0.51, 0.87, 0.51, 0.85, 0.51], ['c', 0.83, 0.60, 0.79, 0.68, 0.72, 0.76],
  ['c', 0.59, 0.91, 0.42, 1.00, 0.243, 0.997], ['z'],
  ['m', 0.809, 0.366],
  ['c', 0.80, 0.37, 0.79, 0.38, 0.79, 0.39], ['c', 0.78, 0.42, 0.80, 0.43, 0.82, 0.45],
  ['c', 0.82, 0.43, 0.82, 0.41, 0.82, 0.39], ['c', 0.82, 0.38, 0.82, 0.37, 0.809, 0.366], ['z']
];
function curl(slide, x, y, w, h, rot, variant) {
  fillPoly(slide, x, y, w, h, variant === 'loop' ? CURL_LOOP : CURL_WAVE, C.gold,
    { rotate: rot || 0 });
}

// Soft yellow starburst behind several illustrations.
function starburst(slide, x, y, w, h, rot, color) {
  fillPoly(slide, x, y, w, h, [
    [0.553, 0.540],
    ['c', 0.72, 0.53, 0.88, 0.44, 0.95, 0.32], ['c', 0.79, 0.33, 0.63, 0.37, 0.51, 0.46],
    ['c', 0.53, 0.30, 0.51, 0.15, 0.46, 0.00], ['c', 0.44, 0.17, 0.42, 0.33, 0.40, 0.50],
    [0.383, 0.542],
    ['c', 0.27, 0.42, 0.14, 0.29, 0.01, 0.18], ['c', 0.11, 0.33, 0.20, 0.49, 0.30, 0.64],
    ['c', 0.18, 0.66, 0.07, 0.71, 0.00, 0.79], ['c', 0.15, 0.75, 0.30, 0.72, 0.45, 0.69],
    ['c', 0.44, 0.79, 0.45, 0.90, 0.49, 1.00], ['c', 0.53, 0.90, 0.55, 0.79, 0.57, 0.69],
    ['c', 0.71, 0.71, 0.86, 0.72, 1.00, 0.72], ['c', 0.85, 0.66, 0.70, 0.60, 0.553, 0.540]
  ], color || C.star, { rotate: rot || 0 });
}

// Rounded 4-petal flower: small + coloured = bullet / slider knob,
// large + very light gray = the background decoration on many slides.
function flower(slide, x, y, w, h, color, rot) {
  fillPoly(slide, x, y, w, h, [
    [0.919, 0.497],
    ['c', 0.88, 0.47, 0.84, 0.46, 0.79, 0.46], ['c', 0.82, 0.45, 0.84, 0.43, 0.87, 0.41],
    ['c', 0.92, 0.37, 0.95, 0.31, 0.94, 0.25], ['c', 0.93, 0.19, 0.87, 0.13, 0.80, 0.14],
    ['c', 0.74, 0.16, 0.71, 0.23, 0.68, 0.30], ['c', 0.66, 0.35, 0.63, 0.41, 0.58, 0.43],
    ['c', 0.63, 0.38, 0.66, 0.31, 0.67, 0.24], ['c', 0.68, 0.19, 0.68, 0.14, 0.65, 0.09],
    ['c', 0.59, -0.03, 0.38, -0.03, 0.29, 0.08], ['c', 0.20, 0.20, 0.23, 0.37, 0.34, 0.47],
    ['c', 0.27, 0.45, 0.20, 0.43, 0.13, 0.45], ['c', 0.06, 0.46, -0.01, 0.52, 0.00, 0.59],
    ['c', 0.01, 0.67, 0.10, 0.71, 0.17, 0.73], ['c', 0.25, 0.75, 0.34, 0.74, 0.38, 0.69],
    ['c', 0.36, 0.74, 0.34, 0.79, 0.34, 0.84], ['c', 0.34, 0.91, 0.38, 0.98, 0.45, 1.00],
    ['c', 0.52, 1.01, 0.60, 0.95, 0.62, 0.88], ['c', 0.64, 0.81, 0.62, 0.73, 0.60, 0.65],
    ['c', 0.65, 0.73, 0.73, 0.78, 0.82, 0.81], ['c', 0.86, 0.82, 0.91, 0.82, 0.95, 0.80],
    ['c', 0.97, 0.78, 0.99, 0.74, 0.99, 0.71], ['c', 1.01, 0.63, 0.98, 0.55, 0.919, 0.497]
  ], color, { rotate: rot || 0 });
}

// Pale flower that sits behind an illustration.
function backdrop(slide, x, y, w, h, rot) {
  flower(slide, x, y, w, h, C.soft, rot);
}

/* --------------------------------------------------------------- text bits */

function header(slide, dark) {
  slide.addText('School Education Infograph', { x: 0.47, y: 0.33, w: 2.8, h: 0.32,
    fontFace: F.head, fontSize: 13, color: dark ? C.white : C.black, valign: 'middle' });
  slide.addText('Infograph Template', { x: 9.97, y: 0.34, w: 2.8, h: 0.3, align: 'right',
    fontFace: F.body, fontSize: 12, color: dark ? C.white : C.gray, valign: 'middle' });
}

// Two-tone section heading, e.g. "Phase " + "Primary".
function heading(slide, x, y, w, parts, align) {
  slide.addText(parts.map(p => ({ text: p[0], options: { color: p[1] || C.black } })),
    { x, y, w, h: 1.03, fontFace: F.head, fontSize: 48, align: align || 'center',
      valign: 'middle' });
}

function label(slide, x, y, w, text, color, align, size) {
  slide.addText(text, { x, y, w, h: 0.41, fontFace: F.head, fontSize: size || 16,
    color, align: align || 'left', valign: 'middle' });
}

function body(slide, x, y, w, text, align, size) {
  slide.addText(text, { x, y, w, h: 0.38, fontFace: F.body, fontSize: size || 12,
    color: C.gray, align: align || 'left', valign: 'middle' });
}

function cardBox(slide, x, y, w, h) {
  slide.addShape('roundRect', { x, y, w, h, fill: { color: C.white },
    rectRadius: Math.min(w, h) * 0.18, shadow: CARD_SHADOW });
}

// White card: coloured title line above a gray caption.
function titleCard(slide, x, y, w, h, title, color, align, caption) {
  cardBox(slide, x, y, w, h);
  const a = align || 'center';
  label(slide, x + 0.14, y + 0.11, w - 0.28, title, color, a);
  body(slide, x + 0.10, y + 0.43, w - 0.20, caption === undefined ? LOREM : caption, a);
}

// White card: big number above a gray caption.
function statCard(slide, x, y, w, h, value, color, align, size, caption) {
  cardBox(slide, x, y, w, h);
  const a = align || 'center';
  slide.addText(value, { x: x + 0.14, y: y + 0.07, w: w - 0.28, h: 0.64, fontFace: F.head,
    fontSize: size || 28, color, align: a, valign: 'middle' });
  body(slide, x + 0.10, y + 0.55, w - 0.20, caption === undefined ? LOREM_SHORT : caption, a);
}

// Wide white card: big number and caption side by side.
function wideCard(slide, x, y, w, h, value, color, numRight, caption) {
  cardBox(slide, x, y, w, h);
  const vw = 1.03, cw = 2.26;
  slide.addText(value, { x: numRight ? x + w - 0.16 - vw : x + 0.16, y: y + 0.05, w: vw, h: 0.64,
    fontFace: F.head, fontSize: 28, color, align: 'center', valign: 'middle' });
  body(slide, numRight ? x + w - 0.30 - vw - cw : x + 1.07, y + 0.19, cw,
    caption === undefined ? LOREM : caption, 'center');
}

/* ------------------------------------------------------- shared components */

// Slim progress bar: caption, track, coloured fill, flower knob and percentage.
function progressRow(slide, x, y, wTrack, text, pct, color, pctX) {
  body(slide, x - 0.09, y - 0.36, 1.7, text, 'left', 10);
  pill(slide, x, y, wTrack, 0.12, C.track);
  const wFill = wTrack * pct / 100;
  pill(slide, x, y, wFill, 0.12, color);
  flower(slide, x + wFill - 0.16, y - 0.10, 0.31, 0.32, color, 35);
  body(slide, pctX, y - 0.13, 0.6, pct + '%', 'left', 10);
}

// Small white pictograms - enough visual weight to read as icons.
function glyph(slide, kind, x, y, w, h, color) {
  const c = color || C.white;
  const u = (fx, fy, fw, fh, col) => rect(slide, x + w * fx, y + h * fy, w * fw, h * fh, col || c);
  switch (kind) {
    case 'doc':
      u(0.12, 0, 0.76, 1);
      u(0.26, 0.16, 0.48, 0.09, C.bg); u(0.26, 0.44, 0.48, 0.09, C.bg); u(0.26, 0.72, 0.30, 0.09, C.bg);
      break;
    case 'chart':
      u(0.04, 0.42, 0.24, 0.58); u(0.38, 0.14, 0.24, 0.86); u(0.72, 0.58, 0.24, 0.42);
      break;
    case 'lock':
      u(0.10, 0.42, 0.80, 0.58);
      poly(slide, x + w * 0.26, y + h * 0.06, w * 0.48, h * 0.40,
        [[0, 1], ['c', 0, 0.05, 1, 0.05, 1, 1]], { line: { color: c, width: 2.5 } });
      break;
    case 'calc':
      u(0.10, 0, 0.80, 1); u(0.22, 0.10, 0.56, 0.20, C.bg);
      [0.44, 0.68].forEach(row => [0.22, 0.42, 0.62].forEach(cx => u(cx, row, 0.16, 0.16, C.bg)));
      break;
    case 'home':
      fillPoly(slide, x, y, w, h,
        [[0.5, 0], [1, 0.45], [0.82, 0.45], [0.82, 1], [0.18, 1], [0.18, 0.45], [0, 0.45]], c);
      break;
    case 'calendar':
      u(0.04, 0.14, 0.92, 0.86); u(0.22, 0, 0.12, 0.26); u(0.66, 0, 0.12, 0.26);
      u(0.16, 0.44, 0.68, 0.10, C.red);
      break;
    case 'chat':
      oval(slide, x, y, w, h * 0.80, c);
      fillPoly(slide, x + w * 0.18, y + h * 0.58, w * 0.30, h * 0.42, [[0, 0], [0.9, 0], [0.1, 1]], c);
      break;
    case 'book':
      u(0.06, 0.04, 0.38, 0.92); u(0.56, 0.04, 0.38, 0.92);
      break;
    case 'mail':
      u(0.00, 0.06, 1.00, 0.88);
      fillPoly(slide, x + w * 0.08, y + h * 0.14, w * 0.84, h * 0.44,
        [[0, 0], [1, 0], [0.5, 1]], C.white);
      break;
    case 'clip':
      u(0.08, 0.10, 0.84, 0.90); u(0.30, 0, 0.40, 0.18);
      u(0.24, 0.38, 0.52, 0.08, C.yellow); u(0.24, 0.58, 0.52, 0.08, C.yellow);
      break;
    case 'pencil':
      fillPoly(slide, x, y, w, h,
        [[0.30, 1.00], [0.00, 0.70], [0.70, 0.00], [1.00, 0.30]], c);
      fillPoly(slide, x, y, w, h, [[0.30, 1.00], [0.00, 0.70], [0.00, 1.00]], c);
      break;
    case 'user':
      oval(slide, x + w * 0.26, y, w * 0.48, h * 0.46, c);
      fillPoly(slide, x, y + h * 0.52, w, h * 0.48,
        [[0.5, 0], ['c', 0.95, 0, 1, 0.7, 1, 1], [0, 1], ['c', 0, 0.7, 0.05, 0, 0.5, 0]], c);
      break;
    case 'bulb':
      oval(slide, x + w * 0.14, y, w * 0.72, h * 0.70, c); u(0.34, 0.60, 0.32, 0.34);
      break;
    case 'monitor':
      u(0, 0, 1, 0.72); u(0.42, 0.72, 0.16, 0.18); u(0.22, 0.90, 0.56, 0.10);
      break;
    default:
      u(0.1, 0.1, 0.8, 0.8);
  }
}

// Upright pencil used as a bar on slide 4.
function pencilUp(slide, x, yTop, w, yBottom, color, stripe) {
  const wood = 0.55, lead = 0.30;
  fillPoly(slide, x, yTop, w, wood, [[0.5, 0], [1, 1], [0, 1]], C.white);
  fillPoly(slide, x + w * 0.29, yTop, w * 0.42, lead, [[0.5, 0], [1, 1], [0, 1]], '262626');
  rect(slide, x, yTop + wood, w, yBottom - yTop - wood, color);
  rect(slide, x + w * 0.62, yTop + wood, w * 0.30, yBottom - yTop - wood, stripe);
}

// Closed book seen from the side, with a ribbon bookmark (slide 11).
function sideBook(slide, x, y, color) {
  const w = 3.38, h = 1.33;
  rect(slide, x + 0.10, y + 0.10, w - 0.10, h - 0.20, C.white);
  for (let i = 0; i < 11; i++) {
    seg(slide, x + 0.22 + (i % 2) * 0.05, y + 0.22 + i * 0.088,
      x + 2.80 + (i % 3) * 0.06, y + 0.22 + i * 0.088, 'BFBFBF', 1);
  }
  pill(slide, x + 0.08, y, w, 0.19, color);
  pill(slide, x + 0.08, y + h - 0.19, w, 0.19, color);
  fillPoly(slide, x, y - 0.01, 0.42, h + 0.02,
    [[0.55, 0], [1, 0], [1, 1], [0.55, 1], ['c', 0.0, 0.80, 0.0, 0.20, 0.55, 0]], color);
  fillPoly(slide, x + 1.53, y + 0.38, 0.42, 0.58, [[0, 0], [1, 0], [1, 1], [0.5, 0.70], [0, 1]], color);
}

// Isometric hardcover book for the stack on slide 8: a slab of pages seen from
// above, capped by a coloured cover whose edge shows on the two front faces.
function isoBook(slide, x, y, w, h, color) {
  const t = 0.16 / h;   // slab thickness as a fraction of the diamond height
  const side = f => [[0, 0.5], [0.5, 1], [1, 0.5], [1, 0.5 + f], [0.5, 1 + f], [0, 0.5 + f]];
  fillPoly(slide, x, y + 0.10, w, h, side(t * 1.9), color);      // cover edge
  fillPoly(slide, x - 0.02, y + 0.20, w + 0.04, h, side(t * 2.6), C.track);
  fillPoly(slide, x - 0.02, y + 0.20, w * 0.52, h,
    [[0, 0.5], [0.96, 1], [0.96, 1 + t * 2.6], [0, 0.5 + t * 2.6]], 'CFCFCF');
  fillPoly(slide, x, y, w, h, [[0, 0.5], [0.5, 0], [1, 0.5], [0.5, 1]], color);
}

// Jigsaw piece. tabs = [top, right, bottom, left]; 1 = knob, -1 = socket, 0 = flat.
// Body corners are rounded; each tab is a mushroom-shaped bezier bump.
function puzzlePiece(slide, x, y, size, color, tabs, rot) {
  const lo = 0.18, hi = 0.82, cr = 0.07;      // body inset and corner radius
  const corner = [[lo, lo], [hi, lo], [hi, hi], [lo, hi]];
  const pts = [[lo + cr, lo]];
  for (let e = 0; e < 4; e++) {
    const a = corner[e], b = corner[(e + 1) % 4];
    const dx = b[0] - a[0], dy = b[1] - a[1];
    const nx = dy, ny = -dx;                  // outward normal (clockwise path)
    const t = tabs[e];
    if (t) {
      const p1 = [a[0] + dx * 0.34, a[1] + dy * 0.34];
      const p2 = [a[0] + dx * 0.66, a[1] + dy * 0.66];
      pts.push([p1[0], p1[1]]);
      pts.push(['c',
        p1[0] + nx * 0.46 * t - dx * 0.26, p1[1] + ny * 0.46 * t - dy * 0.26,
        p2[0] + nx * 0.46 * t + dx * 0.26, p2[1] + ny * 0.46 * t + dy * 0.26,
        p2[0], p2[1]]);
    }
    const n = corner[(e + 2) % 4];            // rounded corner at b
    pts.push([b[0] - Math.sign(dx) * cr, b[1] - Math.sign(dy) * cr]);
    pts.push(['c', b[0], b[1], b[0], b[1],
      b[0] + Math.sign(n[0] - b[0]) * cr, b[1] + Math.sign(n[1] - b[1]) * cr]);
  }
  fillPoly(slide, x, y, size, size, pts, color, { rotate: rot || 0 });
}

// Triangle with rounded corners - one link in the chain on slide 6.
function roundTriangle(slide, x, y, w, h, color) {
  fillPoly(slide, x, y, w, h, [
    [0.46, 0.04], ['c', 0.50, 0.00, 0.54, 0.00, 0.57, 0.05],
    [0.96, 0.88], ['c', 1.00, 0.96, 0.98, 1.00, 0.92, 1.00],
    [0.08, 1.00], ['c', 0.02, 1.00, 0.00, 0.96, 0.04, 0.88]
  ], color);
}

// One band of the layered pyramid on slide 3: narrow top edge (bulging upward
// so it wraps over the band above), wide bottom edge.
function pyramidBand(slide, x, y, w, h, color, inset, flatTop) {
  const span = 1 - 2 * inset;
  const pts = flatTop
    ? [[inset, 0], [1 - inset, 0], [1, 1], [0, 1]]
    : [[inset, 0.165],
      ['c', inset + span * 0.25, -0.055, 1 - inset - span * 0.25, -0.055, 1 - inset, 0.165],
      [1, 1], [0, 1]];
  fillPoly(slide, x, y, w, h, pts, color);
}

// One tier of the 3-D funnel on slide 9.
function funnelCup(slide, x, y, w, h, color, dark, inset) {
  fillPoly(slide, x, y, w, h, [[0, 0],
    ['c', 0.10, 0.28, 0.30, 0.40, 0.5, 0.40], ['c', 0.70, 0.40, 0.90, 0.28, 1, 0],
    [1 - inset, 0.80], ['c', 1 - inset - 0.10, 0.96, 0.64, 1, 0.5, 1],
    ['c', 0.36, 1, inset + 0.10, 0.96, inset, 0.80]], color);
  fillPoly(slide, x + w * inset * 0.55, y + h * 0.32, w * (1 - inset * 1.1), h * 0.42,
    [[0, 0], [1, 0], [0.86, 1], [0.14, 1]], dark);
}

/* ------------------------------------------------------------------ slides */

function slide01(pres) {
  const s = pres.addSlide();
  s.background = { color: C.navy };
  header(s, true);
  curl(s, 10.998, 3.446, 4.04, 3.17, 243, 'wave');
  curl(s, -0.289, 3.061, 3.01, 2.68, 11, 'loop');
  sparkle(s, 11.194, 1.767, 0.86, 1.17, 116, C.gold);
  pill(s, 2.024, 2.28, 9.44, 1.44, C.blue, { rotate: 2.6 });
  s.addText('School  Education', { x: 2.067, y: 2.217, w: 8.95, h: 1.49, rotate: 2.6,
    fontFace: F.head, fontSize: 72, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  pill(s, 2.957, 3.647, 7.375, 1.604, C.green, { rotate: -3.1 });
  s.addText('Infographic', { x: 2.586, y: 3.486, w: 7.7, h: 1.8, rotate: -3.1,
    fontFace: F.head, fontSize: 85, color: C.white, align: 'center', valign: 'middle', margin: 0 });
  sparkle(s, 2.384, 4.993, 0.92, 1.25, 155, C.gold);
}

function slide02(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  header(s);
  starburst(s, 4.19, 0.47, 4.97, 8.09, 290);
  // open book: navy cover, two white pages, blue rule lines
  fillPoly(s, 4.33, 3.31, 4.68, 2.73, [[0.933, 0.095], [0.878, 0.000], [0.500, 0.899],
    [0.122, 0.000], [0.065, 0.106], [0.000, 0.829],
    ['c', 0.182, 0.787, 0.400, 0.900, 0.500, 1.000],
    ['c', 0.600, 0.900, 0.818, 0.787, 1.000, 0.829]], C.navy);
  fillPoly(s, 4.60, 3.31, 2.07, 2.45, [[1.000, 0.173],
    ['c', 1.000, 0.173, 0.650, 0.000, 0.145, 0.000], [0.000, 0.783],
    ['c', 0.000, 0.783, 0.677, 0.703, 1.000, 1.000]], C.white);
  fillPoly(s, 6.67, 3.31, 2.07, 2.45, [[0.000, 0.173],
    ['c', 0.000, 0.173, 0.350, 0.000, 0.855, 0.000], [1.000, 0.783],
    ['c', 1.000, 0.783, 0.323, 0.703, 0.000, 1.000]], C.white);
  for (let i = 0; i < 9; i++) {
    const y = 3.62 + i * 0.187, sag = 0.06 - i * 0.004;
    poly(s, 5.06 - i * 0.037, y, 1.47 + i * 0.037, 0.30,
      [[0, 1], ['c', 0.35, sag, 0.7, 0, 1, 0]], { line: { color: C.blue, width: 0.75 } });
    poly(s, 6.81, y, 1.47 + i * 0.037, 0.30,
      [[0, 0], ['c', 0.3, 0, 0.65, sag, 1, 1]], { line: { color: C.blue, width: 0.75 } });
  }
  heading(s, 3.89, 0.87, 5.55, [['System ', C.black], ['Overview', C.blue]], 'left');
  sparkle(s, 9.21, 0.62, 0.59, 0.8, 338);
  cardBox(s, 1.04, 3.81, 4.16, 0.99);
  label(s, 1.30, 3.92, 3.65, 'System Journey', C.blue, 'right');
  body(s, 1.30, 4.25, 3.65, LOREM_LONG, 'right');
  cardBox(s, 8.13, 3.81, 4.16, 0.99);
  label(s, 8.38, 3.92, 3.65, 'Framework Journey', C.navy, 'left');
  body(s, 8.38, 4.25, 3.65, LOREM_LONG, 'left');
}

function slide03(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  header(s);
  backdrop(s, 6.70, 0.81, 5.75, 5.89);
  pyramidBand(s, 8.76, 1.45, 1.91, 1.28, C.yellow, 0.31, true);
  pyramidBand(s, 8.16, 2.48, 3.12, 1.53, C.green, 0.193);
  pyramidBand(s, 7.56, 3.76, 4.31, 1.54, C.red, 0.139);
  pyramidBand(s, 6.96, 5.04, 5.51, 1.54, C.blue, 0.109);
  glyph(s, 'lock', 9.56, 1.77, 0.29, 0.34);
  glyph(s, 'doc', 9.55, 2.96, 0.33, 0.37);
  glyph(s, 'calc', 9.53, 4.22, 0.38, 0.42);
  glyph(s, 'chart', 9.45, 5.60, 0.52, 0.46);
  heading(s, 0.75, 1.86, 4.58, [['Phase ', C.black], ['Primary', C.green]], 'left');
  sparkle(s, 5.14, 1.67, 0.59, 0.8, 348);
  [['Focus Learning', 78, C.yellow, 3.50], ['Skill Develop', 41, '02B980', 4.10],
    ['Primary Education', 64, C.red, 4.70], ['Young Growth', 92, C.blue, 5.32]]
    .forEach(r => progressRow(s, 0.86, r[3], 4.14, r[0], r[1], r[2], 5.02));
  statCard(s, 10.22, 2.42, 2.03, 1.13, '113+', C.green);
  statCard(s, 6.67, 4.53, 2.03, 1.13, '129+', C.red);
}

function slide04(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  header(s);
  curl(s, -0.22, 3.69, 2.29, 2.05, 11, 'loop');
  curl(s, 11.32, 2.69, 3.08, 2.42, 243, 'wave');
  heading(s, 3.59, 0.78, 6.16, [['Subject ', C.black], ['Learning', C.red]]);
  sparkle(s, 9.15, 1.44, 0.59, 0.8, 63);
  [[5.03, 5.39, C.green, '3DCEB5'], [6.17, 3.09, C.red, 'F77872'],
    [7.34, 3.86, C.yellow, 'FFB14A'], [8.52, 5.39, C.blue, 'B59BC5']]
    .forEach(p => pencilUp(s, p[0], p[1], 0.46, 7.5, p[2], p[3]));
  label(s, 2.70, 4.98, 2.26, 'Core Subject', C.green, 'right');
  body(s, 2.70, 5.31, 2.26, LOREM, 'right');
  label(s, 3.75, 2.81, 2.26, 'Basic Structure', C.red, 'right');
  body(s, 3.75, 3.14, 2.26, LOREM, 'right');
  label(s, 7.94, 3.57, 2.26, 'Learning Topic', C.yellow, 'left');
  body(s, 7.94, 3.90, 2.26, LOREM, 'left');
  label(s, 9.12, 5.20, 2.26, 'Content Outline', C.blue, 'left');
  body(s, 9.12, 5.53, 2.26, LOREM, 'left');
}

function slide05(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  header(s);
  starburst(s, 3.85, 1.23, 5.03, 6.51, 270);
  [[4.78, 2.63, 3.71, C.slate], [5.15, 2.99, 2.98, C.white], [5.51, 3.36, 2.25, C.slate],
    [5.88, 3.72, 1.52, C.white], [6.24, 4.09, 0.80, C.gray]]
    .forEach(r => oval(s, r[0], r[1], r[2], r[2], r[3]));
  // three darts stuck in the board plus the big black one
  fillPoly(s, 6.64, 2.55, 0.44, 0.84, [[1, 0], [1, 1], [0, 0.55]], C.red);
  fillPoly(s, 5.95, 3.00, 0.45, 0.60, [[1, 0], [1, 1], [0, 0.45]], C.red);
  fillPoly(s, 7.83, 2.63, 0.77, 1.02, [[0.2, 0], [1, 0.25], [0.55, 1]], C.blue);
  fillPoly(s, 4.67, 4.41, 0.91, 0.34, [[0, 0.1], [1, 0], [0.6, 1]], C.yellow);
  fillPoly(s, 5.30, 4.83, 0.62, 0.49, [[0, 0], [1, 0.3], [0.35, 1]], C.yellow);
  seg(s, 5.06, 4.91, 8.60, 3.47, '3F3F3F', 14);
  oval(s, 4.90, 4.74, 0.33, 0.34, '3F3F3F');
  oval(s, 8.40, 3.23, 0.41, 0.49, '3F3F3F');
  heading(s, 3.59, 0.70, 2.51, [['Subject', C.black]], 'left');
  heading(s, 6.12, 0.70, 3.08, [['Learning', 'FF0000']]);
  sparkle(s, 9.19, 1.42, 0.59, 0.8, 63);
  statCard(s, 3.04, 2.65, 2.56, 1.11, '186+', C.red, 'right', 28, LOREM);
  statCard(s, 8.13, 4.57, 2.56, 1.11, '125+', C.blue, 'left', 28, LOREM);
  statCard(s, 3.25, 5.52, 2.56, 1.11, '128+', C.yellow, 'right', 28, LOREM);
}

function slide06(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  header(s);
  curl(s, -0.49, 1.66, 1.92, 1.72, 11, 'loop');
  curl(s, 11.67, 4.89, 2.58, 2.03, 243, 'wave');
  heading(s, 3.97, 0.96, 5.39, [['Advanced', C.blue], [' Topic', C.black]]);
  sparkle(s, 3.86, 0.69, 0.59, 0.8, 252);
  // five triangles chained together; at each seam the left triangle's knob
  // pokes through a socket cut into the right triangle's slope
  const tris = [[3.43, 3.15, C.navy], [4.56, 3.00, C.blue], [5.70, 3.15, C.green],
    [6.83, 3.00, C.yellow], [7.97, 3.15, C.red]];
  tris.forEach(t => roundTriangle(s, t[0], t[1], 1.93, 1.71, t[2]));
  tris.slice(1).forEach((t, i) => {
    const kx = t[0] + 0.42, ky = t[1] + 1.02;
    oval(s, kx - 0.19, ky - 0.19, 0.38, 0.38, C.bg);
    oval(s, kx - 0.155, ky - 0.155, 0.31, 0.31, tris[i][2]);
  });
  label(s, 0.91, 3.34, 2.26, 'Learning Topic', C.navy, 'center');
  body(s, 0.91, 3.67, 2.26, LOREM, 'center');
  label(s, 10.38, 3.34, 2.26, 'Basic Structure', C.red, 'center');
  body(s, 10.38, 3.67, 2.26, LOREM, 'center');
  label(s, 1.46, 5.49, 2.26, 'Content Outline', C.blue, 'center');
  body(s, 1.46, 5.82, 2.26, LOREM, 'center');
  label(s, 5.53, 5.69, 2.26, 'Core Subject', C.green, 'center');
  body(s, 5.53, 6.02, 2.26, LOREM, 'center');
  label(s, 9.38, 5.49, 2.26, 'Progress Track', C.yellow, 'center');
  body(s, 9.38, 5.82, 2.26, LOREM, 'center');
}

function slide07(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  header(s);
  backdrop(s, 2.44, 2.22, 4.35, 4.46);
  backdrop(s, 7.35, 2.47, 3.56, 3.64, 69);
  heading(s, 3.84, 0.82, 5.65, [['Academic', C.green], [' Tools', C.black]]);
  sparkle(s, 3.93, 0.55, 0.59, 0.8, 252);
  // pencil lying on its side, split into four coloured segments
  const yT = 3.99, hB = 1.06;
  fillPoly(s, 2.50, yT, 1.13, hB, [[1, 0], [1, 1], [0, 0.5]], C.white);
  fillPoly(s, 2.50, yT + hB * 0.34, 0.43, hB * 0.33, [[1, 0], [1, 1], [0, 0.5]], '262626');
  [[3.63, 1.73, C.blue], [5.36, 1.62, C.red], [6.98, 1.59, C.green], [8.57, 1.74, C.yellow]]
    .forEach(sg => rect(s, sg[0], yT, sg[1], hB, sg[2]));
  fillPoly(s, 10.10, yT, 0.40, hB, [[0, 0], [0.55, 0.12], [0.55, 0.88], [0, 1]], C.white);
  oval(s, 10.27, 4.36, 0.09, 0.30, C.black);
  glyph(s, 'home', 4.31, 4.36, 0.37, 0.33);
  glyph(s, 'calendar', 5.99, 4.36, 0.34, 0.34);
  glyph(s, 'chat', 7.63, 4.36, 0.34, 0.34);
  glyph(s, 'book', 9.28, 4.36, 0.34, 0.34);
  titleCard(s, 5.36, 2.71, 2.59, 0.99, 'Building Learning', C.red, 'left');
  titleCard(s, 8.67, 2.71, 2.59, 0.99, 'Basic Structure', C.yellow, 'left');
  titleCard(s, 3.71, 5.33, 2.59, 0.99, 'Content Outline', C.blue, 'left');
  titleCard(s, 7.02, 5.33, 2.59, 0.99, 'Core Subject', C.green, 'left');
}

function slide08(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  header(s);
  starburst(s, 7.16, 0.66, 4.78, 6.19, 286);
  [[C.yellow, 4.36], [C.green, 3.65], [C.blue, 2.94], [C.red, 2.22]]
    .forEach(b => isoBook(s, 8.00, b[1], 3.08, 1.55, b[0]));
  heading(s, 0.71, 2.70, 4.32, [['Student ', C.black], ['Skill', C.red]], 'left');
  sparkle(s, 4.60, 2.34, 0.59, 0.8, 343);
  flower(s, 0.86, 4.16, 0.31, 0.32, C.yellow, 35);
  s.addText('237.424+', { x: 1.24, y: 3.93, w: 2.58, h: 0.8, fontFace: F.head,
    fontSize: 36, color: C.red, valign: 'middle', margin: 0 });
  body(s, 1.24, 4.59, 3.12, 'Sed ut perspiciatis unde omnis iste', 'left');
  titleCard(s, 6.90, 1.90, 2.59, 0.99, 'Building Learning', C.red, 'right');
  titleCard(s, 6.06, 3.82, 2.59, 0.99, 'Talent Interest', C.green, 'right');
  titleCard(s, 9.88, 3.02, 2.59, 0.99, 'System Journey', C.blue, 'left');
  titleCard(s, 9.95, 4.88, 2.59, 0.99, 'Skill Growth', C.yellow, 'left');
}

function slide09(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  header(s);
  curl(s, -0.68, 2.97, 2.51, 2.24, 39, 'loop');
  curl(s, 10.37, 3.78, 3.37, 2.65, 243, 'wave');
  funnelCup(s, 4.92, 2.58, 3.50, 1.10, C.blue, '2A28A0', 0.09);
  funnelCup(s, 5.30, 3.71, 2.74, 0.94, C.green, '017551', 0.13);
  funnelCup(s, 5.65, 4.71, 2.04, 0.83, C.yellow, 'A08A00', 0.17);
  funnelCup(s, 5.98, 5.70, 1.36, 0.85, C.red, '9B0D19', 0.29);
  heading(s, 4.25, 0.72, 4.83, [['Personal ', C.black], ['Value', C.blue]], 'left');
  sparkle(s, 8.78, 0.56, 0.59, 0.8, 343);
  titleCard(s, 3.11, 2.56, 2.60, 0.99, 'Students Goal', C.blue, 'right');
  statCard(s, 7.43, 3.62, 1.96, 1.03, '176+', C.green, 'left', 24);
  titleCard(s, 3.46, 4.59, 2.60, 0.99, 'Personal Grow', C.yellow, 'right');
  statCard(s, 6.77, 5.70, 1.96, 1.03, '121+', C.red, 'left', 24);
}

function slide10(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  header(s);
  backdrop(s, 3.88, 2.07, 4.35, 4.46);
  backdrop(s, 6.91, 4.04, 3.20, 3.28, 69);
  tri(s, 6.02, 2.55, 1.97, 1.70, C.red);
  fillPoly(s, 5.86, 4.25, 2.13, 0.29, [[0.344, 0], [1, 0], [0.656, 1], [0, 1]], C.redDk);
  fillPoly(s, 5.34, 4.54, 3.33, 0.88, [[0.154, 0], [0.846, 0], [1, 1], [0, 1]], C.yellow);
  fillPoly(s, 5.17, 5.43, 3.50, 0.29, [[0.209, 0], [1, 0], [0.791, 1], [0, 1]], C.yellowDk);
  fillPoly(s, 4.66, 5.72, 4.69, 0.88, [[0.109, 0], [0.891, 0], [1, 1], [0, 1]], C.green);
  heading(s, 4.22, 0.84, 4.90, [['Teacher', C.green], [' Roles', C.black]]);
  sparkle(s, 4.04, 0.54, 0.59, 0.8, 255);
  wideCard(s, 3.32, 3.01, 3.43, 0.76, '93%', C.red, true);
  wideCard(s, 7.99, 4.58, 3.43, 0.76, '78%', C.yellow, false);
  wideCard(s, 1.92, 5.77, 3.43, 0.76, '43%', C.green, true);
}

function slide11(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  header(s);
  starburst(s, 0.42, 0.49, 5.03, 6.51, 270);
  sideBook(s, 1.20, 1.98, C.yellow);
  sideBook(s, 0.76, 3.30, C.green);
  sideBook(s, 1.46, 4.56, C.blue);
  wideCard(s, 3.44, 2.26, 3.43, 0.76, '20+', C.yellow, false);
  wideCard(s, 2.46, 3.55, 3.43, 0.76, '24+', C.green, false);
  wideCard(s, 3.31, 4.85, 3.43, 0.76, '12+', C.blue, false);
  heading(s, 7.76, 2.71, 4.02, [['Learn ', C.green], ['Space', C.black]], 'left');
  sparkle(s, 11.49, 2.48, 0.59, 0.8, 343);
  s.addText('\u201CConsistency shapes success when small daily actions align purpose.\u201D',
    { x: 7.76, y: 3.88, w: 4.46, h: 0.92, fontFace: F.body, fontSize: 16, italic: true,
      color: C.black, valign: 'middle', margin: 0, lineSpacingMultiple: 1.5 });
}

function slide12(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  header(s);
  backdrop(s, 3.52, 2.26, 3.45, 3.53);
  backdrop(s, 6.85, 2.55, 3.12, 3.20);
  // drawn right-to-left so each piece's knob overlaps its neighbour's socket
  [[8.45, 3.14, C.red, [0, 0, 0, -1]], [7.11, 3.13, C.yellow, [0, 1, 0, -1]],
    [5.77, 3.14, C.green, [0, 1, 0, -1]], [4.42, 3.13, C.blue, [0, 1, 0, -1]],
    [3.08, 3.14, C.navy, [0, 1, 0, 0]]]
    .forEach(p => puzzlePiece(s, p[0], p[1], 1.82, p[2], p[3]));
  label(s, 2.86, 1.90, 2.26, 'Learning Topic', C.navy, 'center');
  body(s, 2.86, 2.22, 2.26, LOREM, 'center');
  label(s, 5.53, 1.90, 2.26, 'Core Subject', C.green, 'center');
  body(s, 5.53, 2.22, 2.26, LOREM, 'center');
  label(s, 8.23, 1.90, 2.26, 'Basic Structure', C.red, 'center');
  body(s, 8.23, 2.22, 2.26, LOREM, 'center');
  label(s, 4.20, 5.26, 2.26, 'Content Outline', C.blue, 'right');
  body(s, 4.20, 5.58, 2.26, LOREM, 'right');
  label(s, 6.94, 5.26, 2.26, 'Progress Track', C.yellow, 'center');
  body(s, 6.94, 5.58, 2.26, LOREM, 'center');
}

function slide13(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  header(s);
  curl(s, -0.68, 4.99, 2.51, 2.24, 39, 'loop');
  curl(s, 10.37, 1.93, 3.37, 2.65, 243, 'wave');
  // target seen at an angle: stacked ellipses
  oval(s, 3.20, 4.89, 4.42, 0.73, C.slate);
  oval(s, 3.20, 4.55, 4.42, 0.89, C.ink);
  oval(s, 3.74, 4.72, 3.36, 0.57, C.white);
  oval(s, 4.07, 4.81, 2.69, 0.40, C.ink);
  oval(s, 4.43, 4.92, 1.97, 0.17, C.white);
  oval(s, 4.98, 4.96, 0.87, 0.08, C.slate);
  // three arrows planted in the middle: shaft, chevron fletching, tip
  // each arrow: shaft, three "V" fletches stepping down from the tail, and a tip
  const CHEVRON = [[0.5, 0], [1, 0.72], [1, 1], [0.5, 0.28], [0, 1], [0, 0.72]];
  [[5.42, 5.02, 5.66, 2.68, C.green], [5.18, 5.06, 3.58, 3.24, C.red],
    [5.80, 5.06, 7.32, 3.28, C.yellow]].forEach(a => {
    const [x1, y1, x2, y2, color] = a;
    seg(s, x1, y1, x2, y2, '3F3F3F', 6);
    const len = Math.hypot(x2 - x1, y2 - y1);
    const ux = (x1 - x2) / len, uy = (y1 - y2) / len;      // tail -> tip direction
    const rot = Math.atan2(x2 - x1, y1 - y2) * 180 / Math.PI;
    [0, 1, 2].forEach(i => fillPoly(s, x2 + ux * i * 0.20 - 0.27, y2 + uy * i * 0.20,
      0.54, 0.44, CHEVRON, color, { rotate: rot }));
    fillPoly(s, x1 - 0.21, y1 - 0.24, 0.42, 0.46, [[0.5, 1], [1, 0], [0, 0]], color);
  });
  wideCard(s, 4.97, 2.44, 3.43, 0.76, '51+', C.green, false);
  wideCard(s, 1.31, 3.55, 3.43, 0.76, '22+', C.red, true);
  wideCard(s, 8.60, 3.56, 3.43, 0.76, '34+', C.yellow, false);
  heading(s, 4.25, 0.89, 4.82, [['Smart ', C.green], ['System', C.black]]);
  sparkle(s, 4.06, 0.75, 0.59, 0.8, 233);
}

function slide14(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  header(s);
  // three ring gauges, each "drawn" by a little pencil
  // Each gauge is a pencil bent into a ring: the sharpened tip sits at the top
  // (arc start, 273 deg) and the eraser trails off wherever the arc ends.
  const START = 273, R = 0.95;
  [{ x: 2.15, y: 3.19, color: C.green, end: 20, pct: '43%' },
    { x: 5.91, y: 3.16, color: C.yellow, end: 175, pct: '72%' },
    { x: 9.30, y: 3.11, color: C.blue, end: 95, pct: '50%' }].forEach(g => {
    const cx = g.x + R, cy = g.y + R;
    [[0.000, 1.90, 16.5], [0.075, 1.75, 14.8], [0.225, 1.45, 10.5]].forEach(a =>
      s.addShape('arc', { x: g.x + a[0], y: g.y + a[0], w: a[1], h: a[1],
        line: { color: g.color, width: a[2] }, angleRange: [START, g.end] }));
    const t0 = START * Math.PI / 180;
    const px = cx + R * Math.cos(t0), py = cy + R * Math.sin(t0);
    rect(s, px - 0.42, py - 0.12, 0.42, 0.24, C.white);      // barrel
    rect(s, px - 0.42, py - 0.02, 0.42, 0.07, C.track);      // ferrule stripe
    fillPoly(s, px - 0.60, py - 0.11, 0.19, 0.22, [[1, 0], [1, 1], [0, 0.5]], '333333');
    const t1 = g.end * Math.PI / 180;
    const ex = cx + R * Math.cos(t1), ey = cy + R * Math.sin(t1);
    const fx = -Math.sin(t1), fy = Math.cos(t1);             // forward tangent
    rect(s, ex + fx * 0.04 - 0.17, ey + fy * 0.04 - 0.05, 0.34, 0.10, C.track,
      { rotate: g.end });
    rect(s, ex + fx * 0.26 - 0.19, ey + fy * 0.26 - 0.21, 0.38, 0.42, C.red,
      { rotate: g.end });
    s.addText(g.pct, { x: cx - 0.50, y: cy - 0.33, w: 1.0, h: 0.5, fontFace: F.head,
      fontSize: 24, color: g.color, align: 'center', valign: 'middle' });
  });
  [[3.99, 3.75, C.green, 'doc'], [7.73, 3.75, C.yellow, 'book'], [11.12, 3.76, C.blue, 'chat']]
    .forEach(c => {
      s.addShape('roundRect', { x: c[0], y: c[1], w: 0.53, h: 0.51, fill: { color: C.white },
        rectRadius: 0.12, shadow: CARD_SHADOW });
      glyph(s, c[3], c[0] + 0.15, c[1] + 0.13, 0.23, 0.25, c[2]);
    });
  titleCard(s, 2.01, 5.48, 2.60, 0.99, 'Learn Problem', C.green);
  titleCard(s, 5.48, 5.49, 2.60, 0.99, 'Students Goal', C.yellow);
  titleCard(s, 8.94, 5.49, 2.60, 0.99, 'System Issue', C.blue);
  heading(s, 4.03, 0.97, 5.26, [['School ', C.red], ['Support', C.black]]);
  sparkle(s, 3.85, 0.85, 0.59, 0.8, 233);
}

function slide15(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  header(s);
  backdrop(s, 0.96, 1.81, 4.84, 4.96);
  tri(s, 2.34, 1.61, 2.36, 2.05, C.red);
  fillPoly(s, 2.34, 3.81, 2.36, 2.05, [[0, 0], [1, 0], [0.5, 1]], C.white);
  tri(s, 0.98, 3.84, 2.36, 2.05, C.yellow);
  tri(s, 3.70, 3.84, 2.36, 2.05, C.green);
  oval(s, 3.11, 2.50, 0.82, 0.82, C.white);
  glyph(s, 'mail', 3.28, 2.76, 0.48, 0.30, C.red);
  oval(s, 1.74, 4.78, 0.82, 0.82, C.white);
  glyph(s, 'clip', 1.96, 4.95, 0.38, 0.48, C.yellow);
  oval(s, 4.47, 4.78, 0.82, 0.82, C.white);
  glyph(s, 'pencil', 4.65, 4.96, 0.46, 0.46, C.green);
  glyph(s, 'monitor', 3.31, 4.33, 0.42, 0.40, C.black);
  heading(s, 6.98, 2.28, 5.15, [['Support', C.blue], [' ', C.green], ['Quality', C.black]]);
  sparkle(s, 6.68, 2.15, 0.59, 0.8, 233);
  [[3.64, C.yellow, 'Learning Improvement Strategies'],
    [4.26, C.green, 'Performance enhancement plans'],
    [4.87, C.red, 'Education system advancement']].forEach(it => {
    flower(s, 7.11, it[0], 0.31, 0.32, it[1], 35);
    s.addText(it[2], { x: 7.48, y: it[0] - 0.08, w: 3.5, h: 0.4, fontFace: F.body,
      fontSize: 13, color: C.gray, valign: 'middle', margin: 0 });
  });
}

function slide16(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  header(s);
  curl(s, -0.68, 1.64, 2.51, 2.24, 39, 'loop');
  curl(s, 10.37, 1.93, 3.37, 2.65, 243, 'wave');
  // open book with pages fanning upwards
  fillPoly(s, 5.52, 4.01, 1.18, 2.10, [[1, 0], ['c', 0.30, 0.35, 0.05, 0.72, 0, 1], [0.60, 1]], C.track);
  fillPoly(s, 6.63, 4.28, 1.39, 1.85, [[0, 0], ['c', 0.70, 0.40, 0.95, 0.75, 1, 1], [0.40, 1]], C.track);
  fillPoly(s, 3.11, 5.29, 3.56, 1.04, [[0, 0.25], [1, 0], [1, 1], [0, 0.85]], C.track);
  fillPoly(s, 6.66, 5.29, 3.56, 1.04, [[0, 0], [1, 0.25], [1, 0.85], [0, 1]], C.track);
  fillPoly(s, 3.00, 5.86, 3.37, 0.46, [[0, 0.35], [1, 0], [1, 1], [0, 1]], C.blue);
  fillPoly(s, 6.96, 5.86, 3.37, 0.46, [[0, 0], [1, 0.35], [1, 1], [0, 1]], C.blue);
  rect(s, 3.00, 6.31, 3.37, 0.06, C.navy);
  rect(s, 6.96, 6.31, 3.37, 0.06, C.navy);
  fillPoly(s, 6.32, 6.14, 0.70, 0.23, [[0, 0.35], [0.5, 0], [1, 0.35], [1, 1], [0, 1]], C.navy);
  titleCard(s, 5.42, 2.76, 2.60, 0.99, 'Progress Road', C.green);
  titleCard(s, 1.51, 4.38, 2.60, 0.99, 'Learn Problem', C.red);
  titleCard(s, 9.23, 4.39, 2.60, 0.99, 'Growth Vision', C.blue);
  cardBox(s, 5.48, 5.80, 2.38, 0.93);
  s.addText('102.932+', { x: 5.58, y: 5.85, w: 2.18, h: 0.55, fontFace: F.head, fontSize: 24,
    color: C.blue, align: 'center', valign: 'middle', margin: 0 });
  body(s, 5.58, 6.26, 2.18, LOREM_SHORT, 'center', 11);
  heading(s, 4.46, 0.86, 4.82, [['Result ', C.green], ['System', C.black]]);
  sparkle(s, 4.16, 0.68, 0.59, 0.8, 233);
}

function slide17(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  header(s);
  backdrop(s, 2.77, 2.30, 4.03, 4.13);
  backdrop(s, 6.64, 2.46, 3.62, 3.71);
  // timeline drawn as a horizontal pencil sliced into year blocks
  const yT = 4.23, hT = 0.80, hS = 0.40;
  fillPoly(s, 2.18, yT, 0.65, hT, [[1, 0], [1, 1], [0, 0.5]], C.bg);
  fillPoly(s, 2.18, yT + hS, 0.65, hS, [[1, 0], [1, 1], [0, 0.5]], '333333');
  rect(s, 2.64, yT, 0.18, hT, C.ink);
  [[2.83, 1.11, '2017', C.navy, C.navyDk], [3.93, 1.56, '2018', C.blue, C.blueDk],
    [5.49, 2.29, '2019', C.green, C.greenDk], [7.78, 1.54, '2020', C.yellow, C.yellowDk],
    [9.32, 1.11, '2021', C.red, C.redDk]].forEach(y => {
    rect(s, y[0], yT, y[1], hT, y[3]);
    rect(s, y[0], yT + hS, y[1], hS, y[4]);
    s.addText(y[2], { x: y[0], y: yT, w: y[1], h: hT, fontFace: F.head, fontSize: 14,
      color: C.white, align: 'center', valign: 'middle', margin: 0 });
  });
  fillPoly(s, 10.43, yT, 0.72, hT, [[0, 0], [1, 0.5], [0, 1]], 'F8F2E0');
  fillPoly(s, 10.43, yT + hS, 0.72, hS, [[0, 0], [0.7, 0.5], [0, 1]], 'EDE4CC');
  fillPoly(s, 10.85, yT + 0.28, 0.30, 0.24, [[0, 0], [1, 0.5], [0, 1]], '333333');
  [[2.21, 3.27, 'Learn Problem', '2B2B7C'], [5.53, 3.27, 'Learn Problem', C.green],
    [8.79, 3.27, 'Learn Problem', 'EC5C68'], [3.58, 5.13, 'Learn Problem', C.blue],
    [7.42, 5.13, 'Learn Problem', C.yellow]].forEach(m => {
    label(s, m[0], m[1], 2.26, m[2], m[3], 'center');
    body(s, m[0], m[1] + 0.36, 2.26, LOREM, 'center');
  });
  heading(s, 4.25, 0.88, 4.82, [['Next', C.black], [' ', C.green], ['Academic', C.red]]);
  sparkle(s, 8.84, 0.72, 0.59, 0.8, 345);
}

function slide18(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  header(s);
  // ring split into four jigsaw quadrants; each seam gets a knob and a socket
  const cx = 6.67, cy = 4.54, r = 1.85, ri = r * 0.47;
  const quads = [[C.yellow, 270], [C.green, 0], [C.red, 90], [C.blue, 180]];
  quads.forEach(q => s.addShape('blockArc', { x: cx - r, y: cy - r, w: r * 2, h: r * 2,
    fill: { color: q[0] }, angleRange: [q[1], q[1] + 90], arcThicknessRatio: 0.53 }));
  // seam order N,E,S,W: the quadrant clockwise of the seam owns the knob
  const KNOB = 0.50;
  [[0, -1, C.yellow], [1, 0, C.green], [0, 1, C.red], [-1, 0, C.blue]].forEach(k => {
    oval(s, cx + k[0] * (r - KNOB * 0.45) - KNOB / 2, cy + k[1] * (r - KNOB * 0.45) - KNOB / 2,
      KNOB, KNOB, k[2]);
    oval(s, cx + k[0] * (ri + KNOB * 0.20) - KNOB / 2, cy + k[1] * (ri + KNOB * 0.20) - KNOB / 2,
      KNOB, KNOB, C.bg);
  });
  glyph(s, 'user', 5.67, 3.22, 0.27, 0.31);
  glyph(s, 'doc', 7.61, 3.53, 0.27, 0.31);
  glyph(s, 'bulb', 5.43, 5.19, 0.24, 0.31);
  glyph(s, 'chart', 7.41, 5.51, 0.31, 0.31);
  wideCard(s, 1.89, 2.88, 3.43, 0.76, '14+', C.yellow, true);
  wideCard(s, 8.02, 2.88, 3.43, 0.76, '36+', C.green, false);
  wideCard(s, 1.89, 5.41, 3.43, 0.76, '24+', C.blue, true);
  wideCard(s, 8.02, 5.44, 3.43, 0.76, '28+', C.red, false);
  heading(s, 4.25, 1.03, 4.82, [['Value ', C.green], ['Impact', C.black]]);
  sparkle(s, 4.14, 0.81, 0.59, 0.8, 233);
}

function slide19(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  header(s);
  backdrop(s, 1.30, 0.82, 4.42, 4.53, 180);
  backdrop(s, 2.04, 3.29, 3.42, 3.50, 180);
  // three-tier 3-D funnel: elliptical rim, tapered body, soft highlight
  [[1.36, 1.89, 4.32, 0.38, 1.22, C.yellow, C.yellowDk, 0.13],
    [1.92, 3.20, 3.19, 0.28, 1.19, C.green, C.greenDk, 0.18],
    [2.49, 4.44, 2.06, 0.32, 1.19, C.blue, C.blueDk, 0.28]].forEach(t => {
    const [x, y, w, hTop, hBody, color, dark, inset] = t;
    oval(s, x, y, w, hTop, dark);
    fillPoly(s, x, y + hTop / 2, w, hBody, [[0, 0], [1, 0], [1 - inset, 1], [inset, 1]], color);
    fillPoly(s, x, y + hTop / 2, w * 0.32, hBody, [[0, 0], [1, 0], [1, 1], [0.42, 1]], C.white,
      { fill: { color: C.white, transparency: 78 } });
  });
  statCard(s, 0.68, 2.67, 2.03, 1.13, '121+', C.yellow);
  statCard(s, 3.79, 4.19, 2.03, 1.13, '164+', C.blue);
  heading(s, 6.80, 2.05, 5.52, [['Overview', C.blue], [' ', C.green], ['Impact', C.black]]);
  sparkle(s, 6.43, 1.93, 0.68, 0.8, 233);
  [['Focus Learning', 78, C.yellow, 3.70], ['Skill Develop', 41, '02B980', 4.30],
    ['Primary Education', 64, C.blue, 4.91]]
    .forEach(r => progressRow(s, 7.14, r[3], 4.14, r[0], r[1], r[2], 11.30));
}

function slide20(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  header(s);
  backdrop(s, 7.28, 2.44, 4.02, 4.12);
  puzzlePiece(s, 7.85, 2.53, 2.15, C.yellow, [0, 1, -1, 0]);
  puzzlePiece(s, 9.85, 1.55, 2.20, C.green, [0, 0, 1, -1], 21);
  puzzlePiece(s, 7.99, 4.20, 2.05, C.blue, [1, -1, 0, 0]);
  puzzlePiece(s, 8.89, 4.50, 2.05, C.red, [-1, 0, 0, 1]);
  [[3.49, C.blue, 'Future Education Vision', 2.30],
    [4.12, C.yellow, 'Education System Advancement', 3.06],
    [4.73, C.green, 'Education Progress Roadmap', 2.86],
    [5.35, C.red, 'Learning Evolution Goals', 2.41]].forEach(it => {
    flower(s, 1.16, it[0], 0.31, 0.32, it[1], 35);
    s.addText(it[2], { x: 1.55, y: it[0] - 0.04, w: it[3], h: 0.37, fontFace: F.head,
      fontSize: 14, color: it[1], valign: 'middle', margin: 0 });
  });
  titleCard(s, 9.78, 1.77, 2.60, 0.99, 'Learn Problem', C.green);
  titleCard(s, 6.12, 4.85, 2.60, 0.99, 'Learn Problem', C.blue);
  heading(s, 0.81, 2.16, 5.52, [['Education', C.blue], [' ', C.green], ['Goals', C.black]]);
  sparkle(s, 0.60, 1.95, 0.59, 0.8, 233);
}

/* -------------------------------------------------------------------- main */

const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
  slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17,
  slide18, slide19, slide20];

const pres = new PptxGenJS();
pres.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pres.layout = 'WIDE';
pres.title = 'School Education Infograph';
SLIDES.forEach(fn => fn(pres));

pres.writeFile({
  fileName: path.join(__dirname, '16926811-7cc3-4706-908f-58811625e8f5_grok_final.pptx')
}).then(f => console.log('wrote', f));
