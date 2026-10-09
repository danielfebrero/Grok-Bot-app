/**
 * "Real Estate Made Easy" - 25 slide company profile deck, rebuilt with pptxgenjs.
 *
 *   node 0e0adef3-2f10-46f5-bcf8-67fdb303ad97_grok_final.js
 *
 * Every slide is a plain builder function; colours, positions (inches) and copy
 * are literals. Shared furniture (nav bar, cards, icons, the slate gradient)
 * lives in the small helper layer below.
 */
'use strict';

const path = require('path');
const pptxgen = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const PAPER = 'F1F1F1';   // page background / light type
const INK   = '121212';   // body type
const DEEP  = '032129';   // dark slate cards
const MIST  = 'A5BDC3';   // pale blue-grey
const PALE  = 'E6E7E8';   // light grey cards
const SOOT  = '231F20';   // icon outlines

const HEAD = 'Manrope Medium';
const BODY = 'Manrope Light';

/* Signature slate gradient of the deck: [position %, colour]. */
const SHEEN = [[0, '536B72'], [100, MIST]];
const SHEEN_DARK = [[0, '344C53'], [100, '536B72']];

/* ------------------------------------------------------------------ helpers */

// Plain text box. Sizes in points, positions in inches, colours are hex.
function txt(s, x, y, w, h, text, o) {
  o = o || {};
  s.addText(text, {
    x: x, y: y, w: w, h: h,
    fontFace: o.face || HEAD, fontSize: o.size || 18, color: o.color || INK,
    align: o.align || 'left', valign: 'top', rotate: o.rot || 0, bullet: o.bullet || false,
  });
}

// Rounded card; `r` is the corner radius in inches.
function card(s, x, y, w, h, r, fill, rot) {
  s.addShape('roundRect', {
    x: x, y: y, w: w, h: h, rectRadius: r, rotate: rot || 0,
    fill: typeof fill === 'string' ? { color: fill } : fill, line: { type: 'none' },
  });
}

// Square-cornered block.
function panel(s, x, y, w, h, fill) {
  s.addShape('rect', {
    x: x, y: y, w: w, h: h,
    fill: typeof fill === 'string' ? { color: fill } : fill, line: { type: 'none' },
  });
}

// Circle: filled, or outlined when o.line is given.
function disc(s, x, y, d, fill, o) {
  o = o || {};
  s.addShape('ellipse', {
    x: x, y: y, w: d, h: d,
    fill: o.line ? { type: 'none' } : (typeof fill === 'string' ? { color: fill } : fill),
    line: o.line ? { color: o.line, width: 1 } : { type: 'none' },
  });
}

// Hairline rule (width omitted = thinnest line the renderer can draw).
function rule(s, x, y, w, h, width) {
  s.addShape('line', { x: x, y: y, w: w, h: h, line: { color: DEEP, width: width || 0 } });
}

/* --- gradient cards -------------------------------------------------------
 * pptxgenjs can only fill a shape with one flat colour, so a gradient card is
 * drawn as a stack of thin slices of the rounded rectangle. Each slice is a
 * custom polygon that follows the card outline, which keeps the rounded
 * corners crisp instead of stair-stepping them.
 */
// Slice count scales with the card's length so the ramp stays smooth.
function sliceCount(len) { return Math.max(12, Math.round(len * 7)); }

// Horizontal inset of a rounded rectangle's outline at distance `t` along the
// gradient axis (`len` = axis length, `r` = corner radius).
function inset(t, len, r) {
  const d = t < r ? r - t : (t > len - r ? t - (len - r) : 0);
  return d <= 0 ? 0 : r - Math.sqrt(Math.max(0, r * r - d * d));
}

// One slice of a rounded card between `a` and `b` along the gradient axis.
// `dir` names the edge holding the first gradient stop: down/up/right/left.
function slab(s, x, y, w, h, r, dir, a, b, color, transparency) {
  const vertical = dir === 'down' || dir === 'up';
  const len = vertical ? h : w;
  const across = vertical ? w : h;
  const steps = 5;
  const near = [];
  const far = [];
  for (let i = 0; i <= steps; i++) {
    const t = a + (b - a) * i / steps;
    const off = inset(t, len, r);
    near.push([t, off]);
    far.push([t, across - off]);
  }
  const pts = near.concat(far.reverse()).map(function (p) {
    const along = p[0] - a;
    return vertical ? { x: p[1], y: along } : { x: along, y: p[1] };
  });
  pts.push({ close: true });
  s.addShape('custGeom', {
    x: vertical ? x : x + a, y: vertical ? y + a : y,
    w: vertical ? w : b - a, h: vertical ? b - a : h,
    fill: { color: color, transparency: transparency || 0 },
    line: { type: 'none' }, points: pts,
  });
}

// Gradient-filled rounded card. `dir` is 'down' | 'up' | 'right' | 'left', or
// 'diagUp' for the deck's one corner-to-corner card (dark bottom-left ->
// light top-right): a left-to-right ramp with a bottom-to-top ramp laid over it
// at half opacity, which averages to the diagonal.
function sheen(s, x, y, w, h, r, dir, stops, transparency) {
  stops = stops || SHEEN;
  if (dir === 'diagUp') {
    sheen(s, x, y, w, h, r, 'right', stops);
    sheen(s, x, y, w, h, r, 'up', stops, 50);
    return;
  }
  const len = dir === 'down' || dir === 'up' ? h : w;
  const flip = dir === 'up' || dir === 'left';
  const n = sliceCount(len);
  const step = len / n;
  for (let i = 0; i < n; i++) {
    const p = ((i + 0.5) / n) * 100;
    // slices overlap a little so no hairline of background shows between them
    slab(s, x, y, w, h, r, dir, step * i, Math.min(len, step * (i + 1.25)),
         rampColor(stops, flip ? 100 - p : p), transparency);
  }
}

// Sample a multi-stop colour ramp at position `p` (0..100).
function rampColor(stops, p) {
  let a = stops[0], b = stops[stops.length - 1];
  for (let i = 0; i < stops.length - 1; i++) {
    if (p >= stops[i][0] && p <= stops[i + 1][0]) { a = stops[i]; b = stops[i + 1]; break; }
  }
  const f = b[0] === a[0] ? 0 : (p - a[0]) / (b[0] - a[0]);
  let hex = '';
  for (let c = 0; c < 3; c++) {
    const av = parseInt(a[1].substr(c * 2, 2), 16), bv = parseInt(b[1].substr(c * 2, 2), 16);
    hex += ('0' + Math.round(av + (bv - av) * f).toString(16)).slice(-2);
  }
  return hex.toUpperCase();
}

/* --- icons ---------------------------------------------------------------- */

// Outline of the deck's diagonal arrow, as fractions of the glyph box.
const ARROW = [[0.909, 0], [0.129, 0.78], [0.129, 0.036], [0, 0.036], [0, 1],
               [0.941, 1], [0.941, 0.871], [0.221, 0.871], [1, 0.091]];

// Diagonal arrow glyph (points south-west; pass rot 180 for north-east).
function arrow(s, x, y, d, color, rot) {
  s.addShape('custGeom', {
    x: x, y: y, w: d, h: d, rotate: rot || 0,
    fill: { color: color }, line: { type: 'none' },
    points: ARROW.map(function (p) { return { x: p[0] * d, y: p[1] * d }; })
      .concat([{ close: true }]),
  });
}

// "Scroll back" badge: thin ring with the arrow inside.
function pin(s, x, y, d, color) {
  s.addShape('ellipse', {
    x: x, y: y, w: d, h: d, fill: { type: 'none' }, line: { color: color, width: 1 },
  });
  arrow(s, x + d * 0.357, y + d * 0.356, d * 0.287, color);
}

// Lightning-bolt glyph, 4:5 aspect.
const BOLT = [[0.25, 1], [0.313, 0.65], [0, 0.65], [0.563, 0], [0.688, 0],
              [0.625, 0.4], [1, 0.4], [0.375, 1]];

function bolt(s, x, y, w, color) {
  s.addShape('custGeom', {
    x: x, y: y, w: w, h: w * 1.25, fill: { color: color }, line: { type: 'none' },
    points: BOLT.map(function (p) { return { x: p[0] * w, y: p[1] * w * 1.25 }; })
      .concat([{ close: true }]),
  });
}

// Top navigation bar - identical on every slide.
const NAV = [['Home', 5.512], ['About', 6.736], ['Service', 7.96], ['Portfolio', 9.413],
             ['Contact', 10.66], ['2045', 12.159]];

function nav(s) {
  txt(s, 0.366, 0.325, 1.686, 0.37, '\u00a9COMPANY', { size: 16 });
  NAV.forEach(function (item) { txt(s, item[1], 0.367, 1.05, 0.303, item[0], { size: 12 }); });
  for (let i = 0; i < 3; i++) panel(s, 12.753, 0.47 + i * 0.033, 0.104, 0.013, INK);  // hamburger
}

/* ------------------------------------------------------------ boilerplate copy */
const LOREM12 =
  "Building Better Homes, Growing Sustainably.";
const LOREM11 =
  "Redefining Living and Investment Opportunities";
const LOREM1 =
  "The European languages are members of the same family. ";
const LOREM10 =
  "Everyone realizes why a new common language would be desirable: one";
const LOREM7 = LOREM1 +
  "Their separate";
const LOREM8 = LOREM7 +
  " existence is a";
const LOREM6 = LOREM8 +
  " myth. ";
const LOREM4 = LOREM6 +
  "For";
const LOREM9 = LOREM10 +
  " could refuse to pay expensive translators. ";
const LOREM2 = LOREM4 +
  " science, music, sport, ";
const LOREM5 =
  "“The European languages are members of the same family. Their separate existence is a myth. Europe uses the same vocabulary”";
const LOREM3 = LOREM2 +
  "etc, Europe uses the same vocabulary.\u00a0";

/* -------------------------------------------------------------- the slides */

// Slide 1 - Cover - Real Estate Made Easy
function slide01(s) {
  card(s, 1.266, 1.233, 3.454, 5.429, 0.576, { color: DEEP, transparency: 68 });
  card(s, 1.099, 1.439, 3.21, 5.045, 0.535, { color: DEEP, transparency: 68 });
  sheen(s, 6.365, 3.672, 6.388, 1.447, 0.277, 'right');
  txt(s, 5.584, 1.182, 7.847, 1.447, "Real Estate", { size: 80 });
  txt(s, 6.573, 2.193, 6.626, 1.447, "-Made Easy", { size: 80 });
  nav(s);
  pin(s, 14.91, 0.846, 0.63, SOOT);
  card(s, 0.498, 2.758, 1.749, 0.778, 0.389, { color: DEEP, transparency: 22 });
  txt(s, 0.742, 2.894, 1.337, 0.505, "©2045", { size: 24, color: PAPER });
  disc(s, 1.779, 5.913, 0.638, { color: PAPER, transparency: 78 });
  card(s, 2.469, 5.918, 2.507, 0.634, 0.317, { color: PAPER, transparency: 78 });
  arrow(s, 2.009, 6.139, 0.188, PAPER);
  txt(s, 2.662, 6.035, 2.148, 0.404, "Company Profile", { color: PAPER, align: 'center' });
  txt(s, 1.688, 1.307, 3.568, 0.707, LOREM11, { color: PAPER });
  disc(s, 5.708, 3.791, 0.417, { color: PAPER, transparency: 78 });
  bolt(s, 5.833, 3.89, 0.167, PAPER);
  txt(s, 7.83, 6.235, 1.245, 0.303, "Learn More", { size: 12, face: BODY, color: '888888', align: 'center' });
  txt(s, 6.623, 6.235, 1.398, 0.303, "Explore Now", { size: 12, align: 'center' });
  rule(s, 6.777, 6.552, 1.09, 0, 1.25);
  txt(s, 6.736, 5.409, 4.764, 0.707, LOREM3, { size: 12, face: BODY });
  txt(s, 9.541, 3.93, 3.126, 1.01, "Experience Real Estate Made Simple, Seamless, and Personal", { color: PAPER });
}

// Slide 2 - Content index
function slide02(s) {
  nav(s);
  txt(s, 0.422, 4.385, 5.79, 1.447, "Content", { size: 80 });
  txt(s, 0.422, 5.361, 3.672, 1.447, "Index", { size: 80 });
  card(s, 5.512, 1.071, 3.035, 2.198, 0.366, PALE);
  card(s, 8.677, 4.643, 3.035, 2.182, 0.364, PALE);
  card(s, 8.67, 1.071, 3.035, 3.397, 0.386, DEEP);
  txt(s, 5.84, 2.191, 2.156, 0.404, "Introduction");
  txt(s, 5.84, 2.549, 2.688, 0.505, LOREM1, { size: 12, face: BODY });
  txt(s, 8.97, 5.712, 2.305, 0.404, "Get in Touch");
  txt(s, 8.97, 6.043, 2.688, 0.505, LOREM1, { size: 12, face: BODY });
  txt(s, 5.84, 5.669, 2.498, 0.404, "Future Goals", { color: PAPER });
  txt(s, 5.84, 6.023, 2.688, 0.505, LOREM1, { size: 12, face: BODY, color: PAPER });
  txt(s, 8.97, 3.333, 2.608, 0.404, "About Us", { color: PAPER });
  txt(s, 8.97, 3.654, 2.688, 0.505, LOREM1, { size: 12, face: BODY, color: PAPER });
  txt(s, 5.84, 1.319, 0.574, 0.505, "01", { size: 24 });
  txt(s, 8.97, 4.96, 0.741, 0.505, "04", { size: 24 });
  txt(s, 5.84, 3.826, 0.728, 0.505, "03", { size: 24, color: PAPER });
  txt(s, 8.97, 1.319, 0.728, 0.505, "02", { size: 24, color: PAPER });
  txt(s, 0.422, 1.308, 3.743, 0.774, LOREM11, { size: 20 });
  pin(s, 3.971, 5.868, 0.63, SOOT);
  sheen(s, 11.828, 1.057, 1.906, 5.78, 0.376, 'down');  // runs off the right edge
}

// Slide 3 - Introducing Our Corporate Identity
function slide03(s) {
  nav(s);
  txt(s, 0.456, 0.911, 11.368, 1.447, "Introducing Our", { size: 80 });
  txt(s, 0.456, 3.733, 4.266, 0.774, LOREM12, { size: 20 });
  txt(s, 0.456, 5.465, 3.459, 0.707, LOREM2, { size: 12, face: BODY });
  txt(s, 0.456, 6.177, 3.156, 0.505, LOREM7, { size: 12, face: BODY });
  pin(s, 11.418, 1.969, 0.63, SOOT);
  card(s, 9.703, 3.759, 3.049, 3.035, 0.292, DEEP);
  txt(s, 9.999, 5.483, 2.079, 0.404, "Since 2043", { color: PAPER });
  txt(s, 9.999, 5.818, 2.754, 0.707, LOREM7, { size: 12, face: BODY, color: PAPER });
  txt(s, 0.456, 2.044, 11.368, 1.447, "Corporate Identity", { size: 80 });
  card(s, 9.855, 3.954, 2.754, 0.634, 0.317, { color: PAPER, transparency: 78 });
  txt(s, 10.154, 4.071, 2.148, 0.404, "Company Profile", { color: PAPER, align: 'center' });
}

// Slide 4 - Company Visions
function slide04(s) {
  nav(s);
  txt(s, 0.648, 4.122, 5.541, 1.447, "Company", { size: 80 });
  txt(s, 0.648, 5.092, 5.319, 1.447, "Visions", { size: 80 });
  card(s, 0.64, 1.222, 4.186, 2.754, 0.459, DEEP);
  sheen(s, 4.911, 1.214, 4.186, 2.754, 0.459, 'diagUp');
  card(s, 6.219, 4.201, 4.243, 2.517, 0.42, PALE);
  txt(s, 0.986, 2.497, 2.305, 0.707, "Connected Communities", { color: PAPER });
  txt(s, 0.986, 3.11, 2.688, 0.505, LOREM1, { size: 12, face: BODY, color: PAPER });
  txt(s, 5.236, 2.497, 2.27, 0.707, "Sustainable Solutions", { color: PAPER });
  txt(s, 5.236, 3.11, 2.688, 0.505, LOREM1, { size: 12, face: BODY, color: PAPER });
  txt(s, 6.581, 5.354, 2.1, 0.707, "Simple Ownership");
  txt(s, 6.581, 5.968, 2.688, 0.505, LOREM1, { size: 12, face: BODY });
  txt(s, 0.933, 1.608, 0.741, 0.505, "01", { size: 24, color: PAPER });
  txt(s, 5.239, 1.608, 0.728, 0.505, "02", { size: 24, color: PAPER });
  txt(s, 6.581, 4.465, 0.728, 0.505, "03", { size: 24 });
  pin(s, 11.309, 5.144, 0.63, SOOT);
}

// Slide 5 - Real Estate With a Purposeful Mission
function slide05(s) {
  nav(s);
  txt(s, 0.554, 0.939, 10.904, 1.447, "Real Estate With a", { size: 80 });
  txt(s, 0.554, 2.02, 10.904, 1.447, "Purposeful Mission", { size: 80 });
  card(s, 0.554, 3.436, 8.491, 3.397, 0.566, DEEP);
  txt(s, 1.82, 4.566, 2.351, 0.707, "Helping People Find Homes", { color: PAPER });
  txt(s, 1.82, 5.23, 2.688, 0.505, LOREM1, { size: 12, face: BODY, color: PAPER });
  txt(s, 1.269, 4.566, 0.741, 0.505, "01", { size: 24, color: PAPER });
  txt(s, 5.5, 4.566, 2.688, 0.707, "Delivering Smart Property Solutions", { color: PAPER });
  txt(s, 5.5, 5.23, 2.688, 0.505, LOREM1, { size: 12, face: BODY, color: PAPER });
  txt(s, 4.918, 4.566, 0.741, 0.505, "02", { size: 24, color: PAPER });
}

// Slide 6 - The Values We Build On
function slide06(s) {
  nav(s);
  txt(s, 6.667, 1.506, 5.836, 1.313, "The Values", { size: 72 });
  txt(s, 6.667, 2.437, 6.242, 1.313, "We Build On", { size: 72 });
  txt(s, 6.747, 5.066, 4.305, 0.707, "Experience Real Estate Made Simple, Seamless, and Personal.");
  pin(s, 9.182, 4.107, 0.463, SOOT);
  card(s, 0.664, 1.358, 5.441, 1.608, 0.268, DEEP);
  sheen(s, 0.659, 3.131, 5.802, 1.608, 0.268, 'right');
  card(s, 0.664, 4.904, 5.441, 1.608, 0.268, PALE);
  txt(s, 0.968, 1.76, 2.846, 0.404, "Honest Service", { color: PAPER });
  txt(s, 0.968, 2.06, 4.883, 0.505, LOREM2, { size: 12, face: BODY, color: PAPER });
  txt(s, 0.93, 3.532, 3.767, 0.404, "Strong Commitment ", { color: PAPER });
  txt(s, 0.93, 3.834, 4.883, 0.505, LOREM2, { size: 12, face: BODY, color: PAPER });
  txt(s, 0.968, 5.299, 2.846, 0.404, "Client First");
  txt(s, 0.968, 5.612, 4.883, 0.505, LOREM2, { size: 12, face: BODY });
  txt(s, 6.736, 5.703, 4.764, 0.707, LOREM3, { size: 12, face: BODY });
}

// Slide 7 - Full-Spectrum Real Estate Solutions
function slide07(s) {
  nav(s);
  txt(s, 0.547, 4.629, 10.213, 1.313, "Full-Spectrum Real", { size: 72 });
  txt(s, 0.539, 5.638, 8.874, 1.313, "Estate Solutions", { size: 72 });
  card(s, 3.862, 1.013, 2.887, 3.514, 0.481, PALE);
  card(s, 6.87, 1.013, 2.887, 3.514, 0.481, DEEP);
  sheen(s, 9.876, 1.013, 2.887, 3.514, 0.481, 'up');
  txt(s, 0.535, 1.219, 3.206, 0.774, "Guided by Mission, Built on Results.", { size: 20 });
  txt(s, 4.166, 3.082, 2.27, 0.707, "Buy & Sell Property");
  txt(s, 4.166, 3.733, 2.688, 0.505, LOREM1, { size: 12, face: BODY });
  txt(s, 4.168, 1.387, 0.728, 0.505, "01", { size: 24 });
  txt(s, 7.166, 3.082, 2.27, 0.707, "Manage Your Property", { color: PAPER });
  txt(s, 7.166, 3.733, 2.688, 0.505, LOREM1, { size: 12, face: BODY, color: PAPER });
  txt(s, 7.168, 1.387, 0.728, 0.505, "02", { size: 24, color: PAPER });
  txt(s, 10.145, 3.082, 2.371, 0.707, "Property Investment", { color: PAPER });
  txt(s, 10.145, 3.733, 2.688, 0.505, LOREM1, { size: 12, face: BODY, color: PAPER });
  txt(s, 10.147, 1.387, 0.728, 0.505, "03", { size: 24, color: PAPER });
  pin(s, 10.875, 5.544, 0.63, SOOT);
}

// Slide 8 - Residential Projection
function slide08(s) {
  sheen(s, 7.96, 1.443, 4.492, 1.447, 0.277, 'right');
  nav(s);
  card(s, 4.135, 2.871, 3.555, 4.796, 0.593, { color: DEEP, transparency: 68 }, 270);
  card(s, 4.067, 3.12, 3.715, 4.457, 0.619, { color: DEEP, transparency: 68 }, 270);
  txt(s, 0.478, 0.989, 7.205, 1.447, "RESIDENTIAL", { size: 80 });
  txt(s, 0.469, 2.044, 7.205, 1.447, "PROJECTION", { size: 80 });
  txt(s, 0.477, 3.63, 2.422, 0.505, LOREM1, { size: 12, face: BODY });
  txt(s, 8.757, 5.948, 3.587, 0.774, "Guiding You Home, Every Step of the Way.", { size: 20 });
  txt(s, 8.375, 2.069, 3.891, 0.505, LOREM6, { size: 12, face: BODY, color: PAPER });
  pin(s, 0.621, 5.948, 0.63, SOOT);
  txt(s, 8.375, 1.728, 3.45, 0.404, "Where Dreams Take Shape", { color: PAPER });
  card(s, 3.504, 5.4, 4.686, 1.341, 0.349, { color: PAPER, transparency: 54.424 });
}

// Slide 9 - Commercial Projection
function slide09(s) {
  panel(s, -0.023, 0.018, 5.542, 7.482, PALE);
  card(s, 10.921, 1.344, 1.521, 1.406, 0.281, DEEP);
  sheen(s, 10.921, 2.836, 1.521, 1.406, 0.281, 'down');
  nav(s);
  txt(s, 6.137, 4.242, 7.818, 1.447, "Commercial", { size: 80 });
  txt(s, 6.129, 5.298, 7.205, 1.447, "Projection", { size: 80 });
  txt(s, 1.255, 1.581, 3.692, 0.404, "Central Business Complex");
  txt(s, 1.255, 1.911, 3.707, 0.707, LOREM9, { size: 12, face: BODY });
  txt(s, 0.697, 1.581, 0.741, 0.505, "01", { size: 24 });
  txt(s, 1.255, 3.407, 3.692, 0.404, "Metro Commercial Park");
  txt(s, 1.255, 3.737, 3.707, 0.707, LOREM9, { size: 12, face: BODY });
  txt(s, 0.697, 3.407, 0.741, 0.505, "02", { size: 24 });
  txt(s, 1.255, 5.233, 3.692, 0.404, "Urban Retail Plaza");
  txt(s, 1.255, 5.563, 3.707, 0.707, LOREM9, { size: 12, face: BODY });
  txt(s, 0.697, 5.233, 0.741, 0.505, "03", { size: 24 });
  txt(s, 11.179, 1.73, 1.125, 0.505, "169+", { size: 24, color: PAPER });
  txt(s, 11.179, 2.124, 1.05, 0.269, "Total Project", { size: 10, face: BODY, color: PAPER });
  txt(s, 11.179, 3.115, 1.05, 0.505, "98%", { size: 24, color: PAPER });
  txt(s, 11.214, 3.509, 0.98, 0.438, "Business Opportunity", { size: 10, face: BODY, color: PAPER });
}

// Slide 10 - Featured Project
function slide10(s) {
  nav(s);
  card(s, 9.178, 1.112, 3.323, 2.823, 0.471, DEEP);
  txt(s, 0.555, 1.286, 5.95, 1.447, "Featured", { size: 80 });
  txt(s, 0.519, 2.359, 5.987, 1.447, "Project", { size: 80 });
  card(s, 5.728, 1.112, 3.323, 2.823, 0.471, PALE);
  txt(s, 6.068, 2.729, 2.27, 0.404, "Flagship Project");
  txt(s, 6.068, 3.083, 2.688, 0.505, LOREM1, { size: 12, face: BODY });
  txt(s, 6.07, 1.486, 0.728, 0.505, "01", { size: 24 });
  txt(s, 9.629, 2.729, 2.27, 0.404, "Unique Features", { color: PAPER });
  txt(s, 9.629, 3.083, 2.688, 0.505, LOREM1, { size: 12, face: BODY, color: PAPER });
  txt(s, 9.632, 1.486, 0.728, 0.505, "02", { size: 24, color: PAPER });
  card(s, 0.671, 4.237, 4.154, 2.546, 0.389, { color: PAPER, transparency: 54.424 });
}

// Slide 11 - Market Presence
function slide11(s) {
  nav(s);
  txt(s, 0.524, 1.024, 5.351, 1.447, "Market", { size: 80 });
  txt(s, 0.524, 1.989, 5.351, 1.447, "Presence", { size: 80 });
  txt(s, 0.524, 3.575, 2.97, 1.447, "99%", { size: 80 });
  txt(s, 0.562, 4.785, 1.49, 0.269, "Market Opportunity", { size: 10, face: BODY });
  sheen(s, 3.353, 3.575, 3.028, 3.156, 0.505, 'down');
  card(s, 6.508, 3.575, 3.028, 3.156, 0.505, DEEP);
  card(s, 9.661, 3.575, 3.028, 3.156, 0.505, PALE);
  txt(s, 3.672, 5.257, 2.381, 0.707, "Residential Properties", { color: PAPER });
  txt(s, 3.672, 5.913, 2.82, 0.505, LOREM1, { size: 12, face: BODY, color: PAPER });
  txt(s, 3.674, 3.911, 0.764, 0.505, "01", { size: 24, color: PAPER });
  txt(s, 6.818, 5.257, 2.381, 0.707, "Commercial Properties", { color: PAPER });
  txt(s, 6.818, 5.913, 2.82, 0.453, LOREM1, { size: 12, face: BODY, color: PAPER });
  txt(s, 6.821, 3.911, 0.764, 0.453, "02", { size: 24, color: PAPER });
  txt(s, 9.943, 5.257, 2.215, 0.707, "Luxury Real Estate");
  txt(s, 9.943, 5.913, 2.82, 0.453, LOREM1, { size: 12, face: BODY });
  txt(s, 9.946, 3.911, 0.764, 0.453, "03", { size: 24 });
  txt(s, 7.404, 1.561, 4.266, 0.774, LOREM11, { size: 20 });
  txt(s, 7.404, 2.295, 4.899, 0.505, LOREM2, { size: 12, face: BODY });
}

// Slide 12 - Target Client
function slide12(s) {
  nav(s);
  txt(s, 0.524, 0.979, 7.712, 1.447, "Target Client", { size: 80 });
  card(s, 0.535, 2.626, 5.995, 3.12, 0.52, DEEP);
  card(s, 6.769, 2.626, 5.995, 3.12, 0.52, PALE);
  txt(s, 0.535, 6.149, 4.632, 0.774, "Smart Property Solutions for Modern Living and Growth", { size: 20 });
  pin(s, 10.309, 1.387, 0.63, SOOT);
  txt(s, 2.492, 4.344, 3.994, 0.438, "First-Time Homebuyers", { size: 20, color: PAPER });
  txt(s, 2.492, 2.935, 0.833, 0.269, "Target 01", { size: 10, face: BODY, color: PAPER });
  txt(s, 8.769, 4.344, 3.994, 0.438, "Property Investors", { size: 20 });
  txt(s, 8.769, 2.935, 0.833, 0.269, "Target 02", { size: 10, face: BODY });
  txt(s, 2.492, 4.746, 3.994, 0.707, LOREM2, { size: 12, face: BODY, color: PAPER });
  txt(s, 8.752, 4.746, 3.994, 0.707, LOREM2, { size: 12, face: BODY });
  txt(s, 9.608, 6.283, 3.156, 0.505, LOREM7, { size: 12, face: BODY });
}

// Slide 13 - Business Model
function slide13(s) {
  nav(s);
  card(s, 2.781, 2.524, 5.82, 3.133, 0.522, { color: DEEP, transparency: 68 }, 270);
  card(s, 2.729, 2.696, 5.941, 2.911, 0.485, { color: DEEP, transparency: 68 }, 270);
  txt(s, 7.59, 4.34, 5.936, 1.447, "Business", { size: 80 });
  txt(s, 7.59, 5.295, 5.936, 1.447, "Model", { size: 80 });
  txt(s, 0.607, 1.66, 2.407, 0.774, "Earn from Property Deals", { size: 20 });
  txt(s, 0.607, 2.365, 2.688, 0.505, LOREM1, { size: 12, face: BODY });
  txt(s, 0.607, 3.505, 2.796, 0.774, "Manage Properties for Owners", { size: 20 });
  txt(s, 0.607, 4.199, 2.688, 0.505, LOREM1, { size: 12, face: BODY });
  txt(s, 0.607, 5.265, 2.66, 0.774, "Build and Sell Properties", { size: 20 });
  txt(s, 0.607, 6, 2.688, 0.505, LOREM1, { size: 12, face: BODY });
  card(s, 7.59, 1.482, 5.266, 2.621, 0.437, DEEP);
  txt(s, 8.182, 2.881, 4.323, 0.707, LOREM1 + "Their separate existence is a myth. For science, music, sport, etc, Europe uses the same vocabulary.", { size: 12, face: BODY, color: PAPER });
  txt(s, 8.182, 2.069, 3.054, 0.774, "Our Strategic Approach to Growth", { size: 20, color: PAPER });
}

// Slide 14 - Our Best Team
function slide14(s) {
  nav(s);
  txt(s, 0.386, 4.07, 5.695, 1.447, "Our Best", { size: 80 });
  txt(s, 0.384, 5.167, 3.377, 1.447, "Team", { size: 80 });
  pin(s, 11.699, 2.25, 0.63, SOOT);
  txt(s, 0.384, 1.484, 2.422, 0.505, LOREM1, { size: 12, face: BODY });
  card(s, 4.116, 2.997, 3.076, 0.634, 0.187, { color: DEEP, transparency: 30 });
  card(s, 7.665, 2.997, 3.076, 0.634, 0.187, { color: PAPER, transparency: 30 });
  card(s, 6.11, 5.789, 3.076, 0.634, 0.187, { color: PAPER, transparency: 30 });
  card(s, 9.676, 5.789, 3.076, 0.634, 0.187, { color: PAPER, transparency: 30 });
  txt(s, 4.251, 3.046, 1.968, 0.404, "Sarah Mitchell", { color: PAPER });
  txt(s, 4.251, 3.318, 1.968, 0.252, "Company’s Leader", { size: 9, face: BODY, color: PAPER });
  txt(s, 7.811, 3.046, 1.968, 0.404, "Kevin Sanders");
  txt(s, 7.811, 3.318, 1.968, 0.252, "Property Manager", { size: 9, face: BODY });
  txt(s, 6.28, 5.851, 1.968, 0.404, "Jessica Wong");
  txt(s, 6.28, 6.123, 1.968, 0.252, "Real Estate Agent", { size: 9, face: BODY });
  txt(s, 9.871, 5.851, 1.968, 0.404, "Brian Carter");
  txt(s, 9.871, 6.123, 1.968, 0.252, "Site Supervisor", { size: 9, face: BODY });
  arrow(s, 6.548, 3.22, 0.188, PAPER, 180);
  arrow(s, 10.095, 3.22, 0.188, INK, 180);
  arrow(s, 12.033, 6.029, 0.188, INK, 180);
  arrow(s, 8.562, 6.029, 0.188, INK, 180);
}

// Slide 15 - Company Milestone
function slide15(s) {
  nav(s);
  txt(s, 0.73, 0.945, 5.936, 1.447, "Company", { size: 80 });
  txt(s, 0.73, 1.9, 5.936, 1.447, "Milestone", { size: 80 });
  rule(s, 6.617, 1.345, 0, 6.155);
  disc(s, 6.567, 1.295, 0.1, DEEP);
  card(s, 7.742, 1.345, 4.845, 1.608, 0.268, DEEP);
  sheen(s, 7.776, 3.118, 4.811, 1.608, 0.268, 'right', SHEEN_DARK);
  card(s, 7.742, 4.891, 4.845, 1.608, 0.268, PALE);
  card(s, 6.981, 1.345, 0.656, 1.608, 0.328, DEEP);
  sheen(s, 6.981, 3.12, 0.656, 1.608, 0.328, 'left', [[0, MIST], [100, '49646B']]);
  card(s, 6.981, 4.886, 0.656, 1.608, 0.328, PALE);
  txt(s, 6.658, 5.438, 1.337, 0.505, "2043", { size: 24, align: 'center', rot: 270 });
  txt(s, 6.658, 3.683, 1.337, 0.505, "2044", { size: 24, color: PAPER, align: 'center', rot: 270 });
  txt(s, 6.658, 1.938, 1.337, 0.505, "2045", { size: 24, color: PAPER, align: 'center', rot: 270 });
  txt(s, 8.264, 2.07, 3.681, 0.505, LOREM8, { size: 12, face: BODY, color: PAPER });
  txt(s, 8.264, 1.724, 4.203, 0.438, "5,000+ Units Sold Milestone", { size: 20, color: PAPER });
  txt(s, 8.264, 3.854, 3.681, 0.505, LOREM8, { size: 12, face: BODY, color: PAPER });
  txt(s, 8.264, 3.509, 3.681, 0.438, "First Commercial Project ", { size: 20, color: PAPER });
  txt(s, 8.264, 5.619, 3.681, 0.505, LOREM8, { size: 12, face: BODY });
  txt(s, 8.264, 5.273, 3.845, 0.438, "Company Establishment", { size: 20 });
  txt(s, 0.763, 5.015, 4.305, 0.707, "Experience Real Estate Made Simple, Seamless, and Personal.");
  txt(s, 0.752, 5.652, 4.764, 0.707, LOREM3, { size: 12, face: BODY });
}

// Slide 16 - Take a Moment for a Break
function slide16(s) {
  card(s, 5.213, 1.513, 7.153, 2.608, 0.435, { color: DEEP, transparency: 68 });
  card(s, 5.213, 1.238, 7.031, 3.21, 0.547, { color: DEEP, transparency: 68 });
  nav(s);
  txt(s, 4.784, 4.654, 8.803, 1.447, "Take a Moment", { size: 80 });
  txt(s, 4.784, 5.609, 5.936, 1.447, "for a Break", { size: 80 });
  pin(s, 11.185, 6.017, 0.63, SOOT);
  txt(s, 1.917, 6.345, 1.245, 0.303, "Learn More", { size: 12, face: BODY, color: '888888', align: 'center' });
  txt(s, 0.71, 6.345, 1.398, 0.303, "Explore Now", { size: 12, align: 'center' });
  rule(s, 0.864, 6.661, 1.09, 0, 1.25);
  txt(s, 0.861, 4.831, 3.681, 0.505, LOREM8, { size: 12, face: BODY });
  sheen(s, 0.568, 1.481, 4.944, 2.712, 0.452, 'down');
  txt(s, 0.861, 1.846, 4.305, 0.707, "Experience Real Estate Made Simple, Seamless, and Personal", { color: PAPER });
  txt(s, 0.85, 3.225, 3.882, 0.707, LOREM2, { size: 12, face: BODY, color: PAPER });
  card(s, 11.072, 3.138, 1.749, 0.778, 0.389, { color: DEEP, transparency: 22 });
  txt(s, 11.316, 3.275, 1.337, 0.505, "©2045", { size: 24, color: PAPER });
  disc(s, 8.625, 1.359, 0.638, { color: PAPER, transparency: 78 });
  card(s, 9.316, 1.364, 2.507, 0.634, 0.317, { color: PAPER, transparency: 78 });
  arrow(s, 8.856, 1.584, 0.188, PAPER);
  txt(s, 9.508, 1.481, 2.148, 0.404, "Company Profile", { color: PAPER, align: 'center' });
}

// Slide 17 - Technology & Innovation
function slide17(s) {
  card(s, 0.674, 1.064, 4.697, 1.809, 0.302, DEEP);
  nav(s);
  txt(s, 5.687, 1.064, 6.757, 1.447, "Technology", { size: 80 });
  txt(s, 5.687, 2.189, 6.851, 1.447, "& Innovation", { size: 80 });
  txt(s, 1.146, 1.573, 3.187, 0.438, "Virtual Property Tours", { size: 20, color: PAPER });
  txt(s, 1.146, 1.942, 3.875, 0.505, LOREM6, { size: 12, face: BODY, color: PAPER });
  sheen(s, 0.674, 3.039, 4.697, 1.809, 0.302, 'down');
  txt(s, 1.146, 3.548, 3.786, 0.438, "Property System", { size: 20, color: PAPER });
  txt(s, 1.146, 3.917, 3.875, 0.505, LOREM6, { size: 12, face: BODY, color: PAPER });
  card(s, 0.674, 4.988, 4.697, 1.809, 0.302, PALE);
  txt(s, 1.146, 5.497, 4.057, 0.438, "AI-Powered Market Analysis", { size: 20 });
  txt(s, 1.146, 5.866, 3.875, 0.505, LOREM6, { size: 12, face: BODY });
  txt(s, 8.693, 6.091, 3.028, 0.707, "Integrating Technology in Real Estate");
  txt(s, 8.682, 5.428, 4.07, 0.707, LOREM1 + "Their separate existence is a myth. For science, music, sport, etc, Europe uses", { size: 12, face: BODY });
}

// Slide 18 - Sustainability Initiatives
function slide18(s) {
  card(s, 0.561, 1.15, 6.176, 3.288, 0.548, { color: DEEP, transparency: 68 });
  card(s, 0.393, 1.275, 6.343, 3.055, 0.509, { color: DEEP, transparency: 68 });
  nav(s);
  txt(s, 5.509, 4.772, 7.809, 1.447, "Sustainability", { size: 80 });
  txt(s, 5.509, 5.738, 7.809, 1.447, "Initiatives", { size: 80 });
  txt(s, 6.975, 1.92, 2.125, 0.707, "Eco-Friendly Materials");
  txt(s, 9.167, 1.92, 3.76, 0.707, LOREM1 + "Their separate existence is a myth. For science, music, sport, etc, ", { size: 12, face: BODY });
  txt(s, 6.975, 3.043, 2.125, 0.707, "Energy-Efficient Designs");
  txt(s, 9.167, 3.043, 3.76, 0.707, LOREM1 + "Their separate existence is a myth. For science, music, sport, etc, ", { size: 12, face: BODY });
  card(s, 0.393, 4.747, 4.979, 2.149, 0.358, DEEP);
  txt(s, 0.677, 5.093, 4.632, 0.774, "Smart Property Solutions for Modern Living and Growth", { size: 20, color: PAPER });
  txt(s, 0.677, 6.029, 3.156, 0.505, LOREM7, { size: 12, face: BODY, color: PAPER });
  rule(s, 6.975, 1.802, 5.798, 0);
  rule(s, 6.975, 2.969, 5.798, 0);
}

// Slide 19 - Partnership & Collaborations
function slide19(s) {
  nav(s);
  txt(s, 4.032, 1.23, 8.927, 1.447, "Partnership", { size: 80 });
  txt(s, 4.032, 2.196, 8.927, 1.447, "& Collaborations", { size: 80 });
  sheen(s, 0.538, 0.933, 3.336, 2.948, 0.491, 'down');
  card(s, 0.538, 4.014, 3.336, 2.948, 0.491, PALE);
  card(s, 4.032, 4.014, 3.336, 2.948, 0.491, DEEP);
  txt(s, 0.884, 2.42, 2.27, 0.707, "Construction Firm Partnership", { color: PAPER });
  txt(s, 0.884, 3.07, 2.688, 0.505, LOREM1, { size: 12, face: BODY, color: PAPER });
  txt(s, 0.886, 1.23, 0.728, 0.505, "01", { size: 24, color: PAPER });
  txt(s, 0.884, 5.51, 2.493, 0.707, "Banking & Finance Collaboration");
  txt(s, 0.884, 6.161, 2.688, 0.505, LOREM1, { size: 12, face: BODY });
  txt(s, 0.886, 4.321, 0.728, 0.505, "02", { size: 24 });
  txt(s, 4.442, 5.51, 2.27, 0.707, "Technology Provider Alliance", { color: PAPER });
  txt(s, 4.442, 6.161, 2.688, 0.505, LOREM1, { size: 12, face: BODY, color: PAPER });
  txt(s, 4.445, 4.321, 0.728, 0.505, "03", { size: 24, color: PAPER });
  disc(s, 9.413, 6.156, 0.638, { color: PAPER, transparency: 78 });
  card(s, 10.104, 6.161, 2.507, 0.634, 0.317, { color: PAPER, transparency: 78 });
  arrow(s, 9.643, 6.381, 0.188, PAPER);
  txt(s, 10.296, 6.278, 2.148, 0.404, "Company Profile", { color: PAPER, align: 'center' });
}

// Slide 20 - Client Testimonials
function slide20(s) {
  card(s, 0.672, 3.647, 3.932, 3.12, 0.52, DEEP);
  card(s, 8.821, 3.647, 3.932, 3.12, 0.52, PALE);
  sheen(s, 4.748, 3.647, 3.932, 3.12, 0.52, 'down');
  nav(s);
  txt(s, 1.778, 1.934, 7.042, 1.447, "Testimonials", { size: 80 });
  txt(s, 1.552, 3.991, 1.968, 0.404, "Anna Martinez", { color: PAPER });
  txt(s, 1.552, 4.263, 1.968, 0.252, "Company’s Client", { size: 9, face: BODY, color: PAPER });
  txt(s, 5.624, 3.991, 1.968, 0.404, "Rachel Kim", { color: PAPER });
  txt(s, 5.624, 4.263, 1.968, 0.252, "Company’s Client", { size: 9, face: BODY, color: PAPER });
  txt(s, 9.643, 3.991, 1.968, 0.404, "David Collins ");
  txt(s, 9.643, 4.263, 1.968, 0.252, "Company’s Client", { size: 9, face: BODY });
  txt(s, 1.011, 4.007, 0.485, 0.478, "Place Your Image Here", { size: 9, face: BODY, bullet: true });
  txt(s, 5.084, 4.007, 0.485, 0.478, "Pl", { size: 9, face: BODY, bullet: true });
  txt(s, 0.55, 0.973, 7.042, 1.447, "Client", { size: 80 });
  txt(s, 1.011, 5.469, 3.247, 0.909, LOREM5, { size: 12, face: BODY, color: PAPER });
  txt(s, 5.065, 5.469, 3.247, 0.909, LOREM5, { size: 12, face: BODY, color: PAPER });
  txt(s, 9.13, 5.469, 3.247, 0.909, LOREM5, { size: 12, face: BODY });
  txt(s, 8.795, 1.504, 4.266, 0.774, LOREM12, { size: 20 });
  txt(s, 8.795, 2.19, 4.199, 0.707, LOREM2, { size: 12, face: BODY });
  pin(s, 5.569, 1.382, 0.63, SOOT);
}

// Slide 21 - Awards & Recognition
function slide21(s) {
  nav(s);
  txt(s, 6.305, 1.282, 6.83, 1.447, "Awards &", { size: 80 });
  txt(s, 6.305, 2.199, 6.83, 1.447, "Recognition", { size: 80 });
  txt(s, 1.041, 4.563, 2.384, 0.707, "Best Property Developer Award ");
  txt(s, 3.492, 4.563, 3.072, 0.707, LOREM4, { size: 12, face: BODY });
  txt(s, 1.041, 5.75, 2.125, 0.707, "Green Building Excellence");
  txt(s, 3.492, 5.75, 3.072, 0.707, LOREM4, { size: 12, face: BODY });
  txt(s, 7.727, 4.563, 1.896, 0.707, "Top Real Estate Brand");
  txt(s, 9.784, 4.563, 3.072, 0.707, LOREM4, { size: 12, face: BODY });
  txt(s, 7.727, 5.75, 2.584, 0.707, "Commercial Excellence");
  txt(s, 9.784, 5.75, 3.072, 0.707, LOREM4, { size: 12, face: BODY });
  rule(s, 0.677, 4.325, 11.978, 0);
  rule(s, 0.677, 5.537, 11.978, 0);
}

// Slide 22 - SWOT Analysis
function slide22(s) {
  nav(s);
  txt(s, 5.44, 1.398, 3.574, 1.447, "SWOT", { size: 80 });
  card(s, 1.202, 1.302, 3.261, 5.125, 0.544, { color: DEEP, transparency: 68 });
  card(s, 1.043, 1.496, 3.03, 4.763, 0.505, { color: DEEP, transparency: 68 });
  txt(s, 5.44, 2.303, 5.041, 1.447, "Analysis", { size: 80 });
  txt(s, 5.898, 4.148, 3.115, 0.404, "Strengths");
  txt(s, 5.898, 4.478, 3.115, 0.505, LOREM10, { size: 12, face: BODY });
  txt(s, 5.341, 4.148, 0.741, 0.505, "01", { size: 24 });
  txt(s, 5.898, 5.344, 3.115, 0.404, "Opportunities");
  txt(s, 5.898, 5.673, 3.115, 0.505, LOREM10, { size: 12, face: BODY });
  txt(s, 5.341, 5.344, 0.741, 0.505, "03", { size: 24 });
  txt(s, 9.714, 4.148, 3.115, 0.404, "Weaknesses");
  txt(s, 9.714, 4.478, 3.115, 0.505, LOREM10, { size: 12, face: BODY });
  txt(s, 9.157, 4.148, 0.741, 0.505, "02", { size: 24 });
  txt(s, 9.714, 5.344, 3.115, 0.404, "Threats");
  txt(s, 9.714, 5.673, 3.115, 0.505, LOREM10, { size: 12, face: BODY });
  txt(s, 9.157, 5.344, 0.741, 0.505, "04", { size: 24 });
  pin(s, 11.185, 2.121, 0.63, SOOT);
  card(s, 0.531, 2.303, 1.749, 0.778, 0.389, { color: DEEP, transparency: 22 });
  txt(s, 0.775, 2.439, 1.337, 0.505, "©2045", { size: 24, color: PAPER });
}

// Slide 23 - Investment Opportunity
function slide23(s) {
  nav(s);
  txt(s, 0.406, 0.952, 6.33, 1.447, "Investment", { size: 80 });
  txt(s, 0.406, 5.65, 6.781, 1.447, "Opportunity", { size: 80 });
  sheen(s, 3.29, 2.447, 3.028, 3.156, 0.505, 'down');
  card(s, 6.446, 2.447, 3.028, 3.156, 0.505, DEEP);
  card(s, 9.599, 2.447, 3.028, 3.156, 0.505, PALE);
  txt(s, 3.609, 4.129, 2.381, 0.707, "Premium Residential Units", { color: PAPER });
  txt(s, 3.609, 4.784, 2.82, 0.505, LOREM1, { size: 12, face: BODY, color: PAPER });
  txt(s, 3.611, 2.783, 0.764, 0.505, "01", { size: 24, color: PAPER });
  txt(s, 6.756, 4.129, 2.189, 0.707, "Commercial Office Spaces", { color: PAPER });
  txt(s, 6.756, 4.784, 2.82, 0.453, LOREM1, { size: 12, face: BODY, color: PAPER });
  txt(s, 6.758, 2.783, 0.764, 0.453, "02", { size: 24, color: PAPER });
  txt(s, 9.881, 4.129, 2.746, 0.707, "Mixed-Use \nProjects");
  txt(s, 9.881, 4.784, 2.82, 0.453, LOREM1, { size: 12, face: BODY });
  txt(s, 9.883, 2.783, 0.764, 0.453, "03", { size: 24 });
  pin(s, 1.427, 3.63, 0.63, SOOT);
  txt(s, 9.352, 6.243, 3.156, 0.505, LOREM7, { size: 12, face: BODY });
}

// Slide 24 - Future Goals
function slide24(s) {
  nav(s);
  txt(s, 0.552, 0.936, 7.242, 1.447, "Future Goals", { size: 80 });
  txt(s, 0.47, 5.178, 2.756, 0.404, "Expand Locations");
  txt(s, 0.459, 5.499, 2.544, 0.505, LOREM1, { size: 12, face: BODY });
  txt(s, 3.627, 6.115, 2.756, 0.404, "Build Green");
  txt(s, 3.617, 6.436, 2.544, 0.505, LOREM1, { size: 12, face: BODY });
  txt(s, 6.722, 5.542, 2.756, 0.404, "Go Digital");
  txt(s, 6.712, 5.862, 2.544, 0.505, LOREM1, { size: 12, face: BODY });
  txt(s, 9.764, 6.121, 2.756, 0.404, "Team Growth");
  txt(s, 9.753, 6.442, 2.544, 0.505, LOREM1, { size: 12, face: BODY });
}

// Slide 25 - Thank You
function slide25(s) {
  nav(s);
  txt(s, 0.491, 1.019, 7.925, 1.447, "Thank You", { size: 80 });
  txt(s, 0.491, 1.936, 7.925, 1.447, "for Your Time", { size: 80 });
  txt(s, 0.491, 3.624, 4.266, 0.774, LOREM12, { size: 20 });
  txt(s, 0.491, 6.206, 3.321, 0.707, LOREM1 + "Their separate existence is a myth.", { size: 12, face: BODY });
  txt(s, 9.516, 4.196, 3.126, 0.404, "(+00) 1262 267 3621");
  txt(s, 9.516, 3.987, 2.688, 0.269, "Phone", { size: 10, face: BODY });
  txt(s, 9.516, 4.946, 3.126, 0.404, "www.example.com");
  txt(s, 9.516, 4.737, 2.688, 0.269, "Website", { size: 10, face: BODY });
  txt(s, 9.516, 5.696, 3.371, 0.707, "742 Evergreen Terrace, Springfield, IL 62704, USA");
  txt(s, 9.516, 5.487, 2.688, 0.269, "Address", { size: 10, face: BODY });
  pin(s, 10.449, 1.836, 0.63, SOOT);
}

/* ------------------------------------------------------------------- build */
const pptx = new pptxgen();
pptx.defineLayout({ name: 'DECK', width: 13.333, height: 7.5 });
pptx.layout = 'DECK';
pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
pptx.title = 'Real Estate Made Easy';

const SLIDES = [slide01, slide02, slide03, slide04, slide05,
  slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15,
  slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25];

SLIDES.forEach(function (build) {
  const s = pptx.addSlide();
  s.background = { color: PAPER };
  build(s);
});

pptx.writeFile({ fileName: path.join(__dirname, '0e0adef3-2f10-46f5-bcf8-67fdb303ad97_grok_final.pptx') })
  .then(function (f) { console.log('wrote ' + f); });
