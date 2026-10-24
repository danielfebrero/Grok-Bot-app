#!/usr/bin/env node
/**
 * Equora - "Finance Consulting Excellence" deck (16 slides, 20" x 11.25").
 *
 * Rebuilt from scratch with pptxgenjs.  Everything (geometry, colours, copy)
 * lives in plain literals and small builder functions so the deck design can be
 * read straight out of this file.  Photographs in the original are replaced by
 * grey "[image]" placeholder rectangles.
 */
const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ tokens */

const SLIDE_W = 20;
const SLIDE_H = 11.25;
const FONT = 'Manrope';

const C = {
  dark:  '2E2E38', // B1  text / background dark 1  (slide background)
  light: 'F2F2F5', // W1  text / background light 1
  navy:  '3C4660', // B2  text / background dark 2
  white: 'F9F9FC', // W2  text / background light 2
  plum:  '953C62', // P1  accent 1
  amber: 'F4B434', // P2  accent 2
  teal:  '2E8B92', // P3  accent 3
  red:   'E15A5A', // P4  accent 4
  green: '3BA76B', // P5  accent 5
  blue:  '4BA3E3', // P6  accent 6
  link:  '3D74E0',
  black: '000000',
  photo: 'CCCCCC', // stand-in fill for the removed photographs
};

const TAGLINE = '— Guiding Your Wealth, Securing Your Future';
const PAGE_NO = '001';
const LOREM = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Ut massa mi.';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet.';

/* ----------------------------------------------------------- path helpers */

// A path is a list of commands: [x, y] = line-to (first one = move-to),
// ['c', x1, y1, x2, y2, x, y] = cubic bezier.  All values are 0..1 fractions
// of the shape's bounding box.
function points(cmds, w, h, close) {
  const pts = cmds.map((c, i) => (c[0] === 'c'
    ? { x: c[5] * w, y: c[6] * h, curve: { type: 'cubic', x1: c[1] * w, y1: c[2] * h, x2: c[3] * w, y2: c[4] * h } }
    : (i === 0 ? { x: c[0] * w, y: c[1] * h, moveTo: true } : { x: c[0] * w, y: c[1] * h })));
  if (close) pts.push({ close: true });
  return pts;
}

const NO_FILL = { type: 'none' };

function freeform(s, cmds, x, y, w, h, o = {}) {
  const opts = { x, y, w, h, points: points(cmds, w, h, o.close !== false) };
  opts.fill = o.fill ? { color: o.fill, transparency: 100 - (o.alpha === undefined ? 100 : o.alpha) } : NO_FILL;
  opts.line = o.line
    ? { color: o.line, width: o.lw || 0.75, transparency: 100 - (o.lineAlpha === undefined ? 100 : o.lineAlpha) }
    : NO_FILL;
  s.addShape('custGeom', opts);
}

/* --------------------------------------------------------- shape helpers */

// Rounded "card". `adj` is the corner radius as a fraction of the short side
// (matches PowerPoint's roundRect adjust value).
function card(s, x, y, w, h, o = {}) {
  const opts = {
    x, y, w, h,
    rectRadius: (o.adj === undefined ? 0.05 : o.adj) * Math.min(w, h),
    fill: { color: o.fill || C.light, transparency: 100 - (o.alpha === undefined ? 100 : o.alpha) },
  };
  if (o.line) opts.line = { color: o.line, width: o.lw || 1.5, transparency: 100 - (o.lineAlpha === undefined ? 20 : o.lineAlpha) };
  s.addShape('roundRect', opts);
}

function pill(s, x, y, w, h, fill) {
  s.addShape('roundRect', { x, y, w, h, rectRadius: Math.min(w, h) / 2, fill: { color: fill } });
}

function circle(s, x, y, d, fill, alpha) {
  s.addShape('ellipse', { x, y, w: d, h: d, fill: { color: fill, transparency: 100 - (alpha === undefined ? 100 : alpha) } });
}

function line(s, x, y, w, h, color, o = {}) {
  s.addShape('line', { x, y, w, h, line: { color, width: o.lw || 0.75, transparency: 100 - (o.alpha === undefined ? 100 : o.alpha) } });
}

/* --------------------------------------------------------------- text */

const TEXT_BASE = { fontFace: FONT, valign: 'top', align: 'left' };

// Most of the deck's text boxes never wrap: they hold pre-broken lines and grow
// to fit (PowerPoint `wrap="none"` + `spAutoFit`).  `label`/`lines` reproduce
// that; `para` is the wrapping variant used for flowing body copy.
function label(s, x, y, w, h, text, o = {}) {
  s.addText(text, Object.assign({}, TEXT_BASE, { x, y, w, h, wrap: false, fit: 'resize' }, o));
}

function lines(s, x, y, w, h, arr, o = {}) {
  label(s, x, y, w, h, arr.map(t => ({ text: t, options: { breakLine: true } })), o);
}

function para(s, x, y, w, h, text, o = {}) {
  s.addText(text, Object.assign({}, TEXT_BASE, { x, y, w, h, wrap: true }, o));
}

function bullets(s, x, y, w, h, arr, o = {}) {
  const bullet = { characterCode: '2022', indent: 22.5 };
  para(s, x, y, w, h, arr.map(t => ({ text: t, options: { breakLine: true, bullet } })), o);
}

/* ------------------------------------------------- brand marks & badges */

// Two interlocking crescents: the Equora logo mark.  `h` drives the size.
const LOGO_LEFT = [
  [0.8078, 0.0607], ['c', 0.7067, -0.0202, 0.5429, -0.0202, 0.4418, 0.0607], [0.0758, 0.3536],
  ['c', -0.0253, 0.4344, -0.0253, 0.5656, 0.0758, 0.6464], [0.4418, 0.9393],
  ['c', 0.5429, 1.0202, 0.7067, 1.0202, 0.8078, 0.9393], [0.9824, 0.7996],
  ['c', 0.8371, 0.7547, 0.7339, 0.6418, 0.7339, 0.5096],
  ['c', 0.7339, 0.3722, 0.8454, 0.2556, 1, 0.2147],
  ['c', 0.997, 0.2121, 0.9939, 0.2096, 0.9908, 0.2071], [0.8078, 0.0607],
];
const LOGO_RIGHT = [
  [0, 0.9667], ['c', 0.0821, 0.9882, 0.1713, 1, 0.2646, 1],
  ['c', 0.6708, 1, 1, 0.7761, 1, 0.5], ['c', 1, 0.2239, 0.6708, 0, 0.2646, 0],
  ['c', 0.1838, 0, 0.1061, 0.0089, 0.0333, 0.0252],
  ['c', 0.3987, 0.2864, 0.3929, 0.6995, 0.0159, 0.9559], [0, 0.9667],
];

function logoMark(s, x, y, h, leftColor) {
  freeform(s, LOGO_LEFT, x, y, 0.799 * h, h, { fill: leftColor });
  freeform(s, LOGO_RIGHT, x + 0.787 * h, y + 0.199 * h, 0.421 * h, 0.621 * h, { fill: C.amber });
}

// Diagonal "open in new" arrow: the deck's recurring badge glyph.
const ARROW = [
  [0.1077, 1], [0, 0.8923], [0.7385, 0.1538], [0.0769, 0.1538], [0.0769, 0],
  [1, 0], [1, 0.9231], [0.8462, 0.9231], [0.8462, 0.2615], [0.1077, 1],
];

// A 40%-opacity badge over a dark card reads as this blended colour; glyph
// cut-outs are painted with it so they look punched out rather than drawn on.
function blend(hex, pct) {
  const base = [0x2e, 0x2e, 0x38];
  const mix = i => Math.round(parseInt(hex.substr(i * 2, 2), 16) * pct / 100 + base[i] * (100 - pct) / 100);
  return [0, 1, 2].map(i => mix(i).toString(16).padStart(2, '0')).join('').toUpperCase();
}

// Pale circle plus a pictogram; the glyph defaults to the diagonal arrow.
function badge(s, x, y, d, color, glyphColor, glyph) {
  circle(s, x, y, d, color, 40);
  const [gx, gy, gw, gh] = GLYPH_BOX[glyph || 'arrow'];
  drawGlyph(s, glyph || 'arrow', x + gx * d, y + gy * d, gw * d, gh * d, glyphColor || color, blend(color, 40));
}

// Each pictogram sits in its own rectangle inside the badge, expressed as
// fractions of the badge diameter: [offsetX, offsetY, width, height].
const GLYPH_BOX = {
  arrow: [0.377, 0.371, 0.231, 0.231],
  grid: [0.341, 0.335, 0.319, 0.319],
  cross: [0.323, 0.318, 0.355, 0.354],
  spark: [0.306, 0.336, 0.390, 0.301],
  cup: [0.340, 0.316, 0.319, 0.354],
  creditcard: [0.324, 0.354, 0.354, 0.283],
  gear: [0.324, 0.316, 0.372, 0.372],
  globe: [0.319, 0.323, 0.354, 0.354],
  calendar: [0.341, 0.317, 0.340, 0.364],
  user: [0.324, 0.317, 0.354, 0.354],
  team: [0.288, 0.388, 0.425, 0.213],
  folder: [0.323, 0.352, 0.354, 0.283],
};

// Pictograms, each assembled from a couple of preset shapes.  `color` is the
// glyph colour, `bg` the badge fill showing through any cut-out.
function drawGlyph(s, name, x, y, w, h, color, bg) {
  const fill = { color };
  const cut = { color: bg };
  switch (name) {
    case 'arrow':
      freeform(s, ARROW, x, y, w, h, { fill: color });
      break;
    case 'grid': // rounded plate with four square cells
      s.addShape('roundRect', { x, y, w, h, rectRadius: w * 0.18, fill });
      [[0, 0], [1, 0], [0, 1], [1, 1]].forEach(([cx, cy]) => {
        s.addShape('rect', { x: x + w * (0.19 + cx * 0.33), y: y + h * (0.19 + cy * 0.33), w: w * 0.29, h: h * 0.29, fill: cut });
      });
      break;
    case 'cross': // two crossed markers
      s.addShape('mathMultiply', { x, y, w, h, fill });
      break;
    case 'spark': // beacon: dome, base bar and radiating rays
      s.addShape('pie', { x: x + w * 0.28, y: y + h * 0.28, w: w * 0.44, h: h * 0.62, angleRange: [180, 360], fill });
      s.addShape('rect', { x: x + w * 0.22, y: y + h * 0.82, w: w * 0.56, h: h * 0.15, fill });
      [[0.5, 0, 0.5, 0.14], [0.17, 0.14, 0.3, 0.32], [0.83, 0.14, 0.7, 0.32],
        [0, 0.6, 0.14, 0.6], [1, 0.6, 0.86, 0.6]].forEach(([x1, y1, x2, y2]) => {
        s.addShape('line', {
          x: x + w * Math.min(x1, x2), y: y + h * Math.min(y1, y2),
          w: w * Math.abs(x2 - x1), h: h * Math.abs(y2 - y1),
          flipV: (x2 - x1) * (y2 - y1) < 0, line: { color, width: 1.2 },
        });
      });
      break;
    case 'cup': // tapered cup with a droplet
      s.addShape('rect', { x: x + w * 0.05, y, w: w * 0.9, h: h * 0.2, fill });
      s.addShape('trapezoid', { x: x + w * 0.08, y: y + h * 0.24, w: w * 0.84, h: h * 0.76, flipV: true, fill });
      s.addShape('teardrop', { x: x + w * 0.34, y: y + h * 0.42, w: w * 0.32, h: h * 0.3, rotate: 225, fill: cut });
      break;
    case 'creditcard':
      s.addShape('roundRect', { x, y, w, h, rectRadius: Math.min(w, h) * 0.2, fill });
      s.addShape('rect', { x: x + w * 0.12, y: y + h * 0.38, w: w * 0.4, h: h * 0.16, fill: cut });
      break;
    case 'gear': // large gear plus a smaller companion
      s.addShape('gear6', { x, y, w: w * 0.62, h: h * 0.62, fill });
      s.addShape('ellipse', { x: x + w * 0.22, y: y + h * 0.22, w: w * 0.18, h: h * 0.18, fill: cut });
      s.addShape('gear6', { x: x + w * 0.42, y: y + h * 0.42, w: w * 0.58, h: h * 0.58, fill });
      s.addShape('ellipse', { x: x + w * 0.62, y: y + h * 0.62, w: w * 0.18, h: h * 0.18, fill: cut });
      break;
    case 'globe': // filled disc with meridian / parallel cut-outs
      s.addShape('ellipse', { x, y, w, h, fill });
      s.addShape('ellipse', { x: x + w * 0.3, y, w: w * 0.4, h, fill: NO_FILL, line: { color: bg, width: 1 } });
      s.addShape('line', { x, y: y + h / 2, w, h: 0, line: { color: bg, width: 1 } });
      break;
    case 'calendar': // pad with two tabs and a tick
      s.addShape('rect', { x: x + w * 0.22, y, w: w * 0.1, h: h * 0.2, fill });
      s.addShape('rect', { x: x + w * 0.68, y, w: w * 0.1, h: h * 0.2, fill });
      s.addShape('roundRect', { x, y: y + h * 0.12, w: w * 0.82, h: h * 0.76, rectRadius: Math.min(w, h) * 0.12, fill });
      s.addShape('rect', { x: x + w * 0.1, y: y + h * 0.4, w: w * 0.62, h: h * 0.38, fill: cut });
      s.addShape('line', { x: x + w * 0.5, y: y + h * 0.66, w: w * 0.16, h: h * 0.2, line: { color, width: 1.8 } });
      s.addShape('line', { x: x + w * 0.66, y: y + h * 0.5, w: w * 0.3, h: h * 0.36, flipV: true, line: { color, width: 1.8 } });
      break;
    case 'user': // ring enclosing head and shoulders
      s.addShape('ellipse', { x, y, w, h, fill: NO_FILL, line: { color, width: 1.6 } });
      s.addShape('ellipse', { x: x + w * 0.34, y: y + h * 0.2, w: w * 0.32, h: h * 0.32, fill });
      s.addShape('ellipse', { x: x + w * 0.16, y: y + h * 0.58, w: w * 0.68, h: h * 0.5, fill });
      break;
    case 'team': // three heads above a shared bar
      s.addShape('ellipse', { x, y: y + h * 0.16, w: w * 0.22, h: h * 0.42, fill });
      s.addShape('ellipse', { x: x + w * 0.35, y, w: w * 0.3, h: h * 0.56, fill });
      s.addShape('ellipse', { x: x + w * 0.78, y: y + h * 0.16, w: w * 0.22, h: h * 0.42, fill });
      s.addShape('roundRect', { x, y: y + h * 0.62, w, h: h * 0.38, rectRadius: h * 0.14, fill });
      break;
    case 'folder': // tabbed folder holding a contact card
      s.addShape('roundRect', { x, y, w: w * 0.42, h: h * 0.28, rectRadius: Math.min(w, h) * 0.08, fill });
      s.addShape('roundRect', { x, y: y + h * 0.16, w, h: h * 0.84, rectRadius: Math.min(w, h) * 0.11, fill });
      s.addShape('ellipse', { x: x + w * 0.55, y: y + h * 0.34, w: w * 0.2, h: h * 0.26, fill: cut });
      s.addShape('ellipse', { x: x + w * 0.45, y: y + h * 0.64, w: w * 0.4, h: h * 0.3, fill: cut });
      break;
  }
}

/* ------------------------------------------------------- concave notches */

// Little "inverted radius" wedges that weld a card into the slide edge.
// The key names the corner that is carved out of the bounding box.
const NOTCH = {
  br_h: [[0, 0], [0, 1], [0.5, 1], ['c', 0.5, 0.4477, 0.724, 0, 1, 0], [0, 0]],
  br_v: [[0, 0], [1, 0], [1, 0.5], ['c', 0.4477, 0.5, 0, 0.724, 0, 1], [0, 0]],
  tl_h: [[1, 1], [1, 0], [0.5, 0], ['c', 0.5, 0.5523, 0.276, 1, 0, 1], [1, 1]],
  tl_v: [[1, 1], [0, 1], [0, 0.5], ['c', 0.5523, 0.5, 1, 0.276, 1, 0], [1, 1]],
  tr_h: [[0, 1], [0, 0], [0.5, 0], ['c', 0.5, 0.5523, 0.724, 1, 1, 1], [0, 1]],
  tr_v: [[0, 1], [1, 1], [1, 0.5], ['c', 0.4477, 0.5, 0, 0.276, 0, 0], [0, 1]],
  bl_h: [[1, 0], [1, 1], [0.5, 1], ['c', 0.5, 0.4477, 0.276, 0, 0, 0], [1, 0]],
  bl_v: [[1, 0], [0, 0], [0, 0.5], ['c', 0.5523, 0.5, 1, 0.724, 1, 1], [1, 0]],
};

function notch(s, kind, x, y, w, h, color) {
  freeform(s, NOTCH[kind], x, y, w, h, { fill: color });
}

/* --------------------------------------------------- image placeholders */

function imageBox(s, x, y, w, h, radius) {
  s.addShape('roundRect', { x, y, w, h, rectRadius: radius, fill: { color: C.photo } });
  s.addText('[image]', {
    x, y: y + h / 2 - 0.2, w, h: 0.4, fontFace: FONT, fontSize: 12, color: '666666', align: 'center', valign: 'middle',
  });
}

/* ------------------------------------------------------- slide furniture */

function chrome(s) {
  label(s, 18.523, 10.524, 0.54, 0.353, PAGE_NO, { fontSize: 15, color: C.light });
  logoMark(s, 1.042, 0.353, 0.302, C.plum);
  label(s, 1.4, 0.304, 0.959, 0.366, 'Equora', { fontSize: 15.74, bold: true, color: C.plum });
  label(s, 2.212, 0.358, 3.1, 0.271, TAGLINE, { fontSize: 10.12, color: C.light });
}

function newSlide(pptx, bg) {
  const s = pptx.addSlide();
  s.background = { color: bg || C.dark };
  return s;
}

/* --- Slide 01 - Title --- */
function slide01(pptx) {
  const s = newSlide(pptx);
  imageBox(s, 1.042, 1.026, 17.917, 9.167, 0.289);

  notch(s, 'br_h', 11.385, 1.042, 0.5, 0.25, C.dark);
  notch(s, 'br_v', 1.042, 5.594, 0.25, 0.5, C.dark);
  card(s, 1.042, 1.042, 10.594, 4.802, { fill: C.dark });
  card(s, 1.042, 1.042, 10.344, 4.552, { fill: C.light, line: C.black });

  notch(s, 'tl_h', 13.458, 9.958, 0.5, 0.25, C.dark);
  notch(s, 'tl_v', 18.708, 7.385, 0.25, 0.5, C.dark);
  card(s, 13.708, 7.635, 5.25, 2.573, { fill: C.dark });
  card(s, 13.958, 7.885, 5.0, 2.323, { fill: C.white, line: C.black });

  badge(s, 1.375, 9.269, 0.588, C.light, C.light);
  para(s, 2.113, 9.214, 3.804, 0.642, 'Trusted advisors in financial planning  and asset management',
    { fontSize: 15, color: C.light });

  para(s, 1.426, 1.552, 5.996, 2.378, 'Finance  Consulting  Excellence', { fontSize: 48, bold: true, color: C.dark });
  para(s, 1.426, 3.969, 6.412, 1.198,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Ut et  massa mi. Aliquam in hendrerit urna. '
    + 'Pellentesque sit amet  sapien fringilla, mattis ligula consectetur, ultrices mauris.',
    { fontSize: 15, color: C.navy });

  pill(s, 8.623, 1.693, 2.278, 0.652, C.plum);
  logoMark(s, 8.956, 1.818, 0.402, C.light);
  label(s, 9.468, 1.77, 1.21, 0.454, 'Equora', { fontSize: 21, bold: true, color: C.white });

  badge(s, 14.292, 8.672, 0.75, C.plum, C.plum);
  label(s, 15.192, 8.167, 1.229, 0.631, '100+', { fontSize: 31.5, bold: true, color: C.dark });
  para(s, 15.192, 8.901, 3.517, 1.057,
    'Reliable and knowledgeable  advisors who you can count on for  guidance and support.',
    { fontSize: 15, color: C.navy });

  chrome(s);
}

/* --- Slide 02 - About our firm --- */
function slide02(pptx) {
  const s = newSlide(pptx);
  card(s, 1.042, 1.073, 12.625, 1.562, { adj: 0.1, fill: C.light, line: C.black });
  card(s, 1.042, 2.844, 12.625, 3.578, { fill: C.light, alpha: 5 });
  card(s, 1.042, 6.63, 6.208, 3.578, { fill: C.light, alpha: 5 });
  card(s, 7.458, 6.63, 6.208, 3.578, { fill: C.white, line: C.black });
  card(s, 13.875, 1.073, 5.083, 2.75, { fill: C.plum });

  label(s, 1.358, 1.349, 5.074, 0.909, 'About Our Firm', { fontSize: 48, bold: true, color: C.dark });
  label(s, 9.015, 1.662, 4.371, 0.353, 'A decade of expertise in finance consulting', { fontSize: 15, color: C.navy });

  badge(s, 6.0, 7.052, 0.833, C.teal, C.light, 'grid');
  label(s, 1.358, 7.99, 2.148, 0.454, 'Asset Growth', { fontSize: 21, bold: true, color: C.light });
  para(s, 1.358, 8.506, 5.614, 1.322,
    'Equora is a trusted financial consulting firm, established  in 2010 with a clear mission: to empower '
    + 'individuals and  organizations to achieve long-term financial stability  and growth.',
    { fontSize: 15, color: C.light });

  badge(s, 12.417, 7.052, 0.833, C.red, C.red, 'cross');
  label(s, 7.775, 7.99, 1.457, 0.454, 'Planning', { fontSize: 21, bold: true, color: C.navy });
  para(s, 7.775, 8.506, 5.475, 1.217,
    'Our approach is built on data-driven insights,  personalized strategies, and transparent  communication. '
    + 'By combining global market expertise  with local knowledge.',
    { fontSize: 15, color: C.navy });

  para(s, 14.192, 1.787, 4.45, 1.474,
    'At Equora, we believe finance is more than  numbers—it’s about creating clarity,  confidence, and '
    + 'sustainable success for  every client.',
    { fontSize: 15, color: C.light });

  badge(s, 1.458, 4.115, 1.042, C.plum, C.light);
  para(s, 2.817, 3.662, 5.058, 1.65,
    'We specialize in three core areas: comprehensive  financial planning, asset management, and risk  analysis. '
    + 'Over the past decade, our team has  supported more than 200 clients worldwide, from  individuals building '
    + 'wealth to corporations  navigating complex financial landscapes.',
    { fontSize: 15, color: C.light });

  chrome(s);
  imageBox(s, 8.271, 3.256, 4.979, 2.75, 0.176);
  imageBox(s, 13.875, 4.031, 5.083, 6.177, 0.162);
}

/* --- Slide 03 - Table of content --- */
const TOC = [
  // rowY,  rowH,  title,               titleW, subtitle,                                                    subX,   subW,  numX,   numW
  [2.625, 0.906, 'Who We Are', 1.932, 'Introducing Equora and our financial consulting expertise', 3.098, 5.726, 10.660, 0.526],
  [3.740, 0.906, 'Our Services', 2.023, ' Comprehensive planning, management, and risk analysis', 3.202, 5.733, 10.608, 0.572],
  [4.854, 0.906, 'Market Analysis', 2.465, 'Key insights into financial trends and opportunities', 3.640, 5.123, 10.619, 0.568],
  [5.969, 0.896, 'Case Studies', 2.092, 'Real-world success stories from our clients', 3.275, 4.378, 10.608, 0.575],
  [7.073, 0.906, 'ROI Projections', 2.414, 'Forecasted returns and growth potential', 3.587, 4.131, 10.619, 0.568],
  [8.188, 0.906, 'Why Choose Us', 2.432, 'What makes Equora a trusted partner', 3.608, 3.822, 10.608, 0.579],
  [9.302, 0.906, 'Closing & Contact', 2.786, 'Your next step toward financial clarity', 3.962, 3.848, 10.629, 0.554],
];

function slide03(pptx) {
  const s = newSlide(pptx);
  label(s, 0.942, 1.109, 5.689, 0.909, 'Table of Content', { fontSize: 48, bold: true, color: C.light });
  label(s, 0.942, 2.068, 4.83, 0.353, 'Quick navigation through our services and value', { fontSize: 15, color: C.light });

  TOC.forEach(([y, h]) => card(s, 1.042, y, 10.375, h, { adj: 0.2, fill: C.light, alpha: 5 }));
  TOC.forEach(([y, , title, titleW, sub, subX, subW, numX, numW], i) => {
    label(s, 1.275, y + 0.199, titleW, 0.454, title, { fontSize: 21, bold: true, color: C.light });
    label(s, subX, y + 0.256, subW, 0.353, sub, { fontSize: 15, color: C.light });
    label(s, numX, y + 0.199, numW, 0.454, '0' + (i + 1), { fontSize: 21, bold: true, color: C.light });
  });

  chrome(s);
  imageBox(s, 11.622, 1.042, 7.336, 9.167, 0.233);
}

/* --- Slide 04 - Comprehensive planning services --- */
const SERVICES = [
  // cardX, cardY, cardFill, cardAlpha, cardLine, badgeColor, glyphColor, titleColor, bodyColor, title, body
  [8.164, 1.042, C.plum, 100, null, C.amber, C.light, C.white, C.light,
    'Personal Wealth Planning', 'Customized strategies to grow and protect  your personal assets', 3.866],
  [13.665, 1.042, C.light, 5, null, C.plum, C.light, C.light, C.light,
    'Retirement Roadmaps', 'Structured plans to secure your financial  independence after work life', 3.396],
  [8.164, 5.734, C.light, 5, null, C.plum, C.light, C.light, C.light,
    'Tax Efficiency Plans', 'Smart approaches to reduce liabilities and  maximize after-tax returns', 3.044],
  [13.665, 5.734, C.light, 100, C.black, C.dark, C.dark, C.dark, C.navy,
    'Corporate Advisory', 'Strategic financial guidance for businesses  and organizations', 2.984],
];

function slide04(pptx) {
  const s = newSlide(pptx);
  card(s, 1.042, 1.042, 6.914, 9.177, { fill: C.white, line: C.black, lineAlpha: 5 });
  SERVICES.forEach(([x, y, fill, alpha, ln]) => card(s, x, y, 5.293, 4.484, { fill, alpha, line: ln, lineAlpha: 5 }));

  para(s, 1.358, 1.318, 6.324, 1.65, 'Comprehensive  Planning Services', { fontSize: 48, bold: true, color: C.dark });
  label(s, 1.358, 3.006, 4.83, 0.353, 'Quick navigation through our services and value', { fontSize: 15, color: C.navy });
  label(s, 1.358, 7.355, 5.126, 0.454, 'Our core planning services include', { fontSize: 21, bold: true, color: C.navy });
  para(s, 1.358, 7.891, 6.324, 1.947,
    'At Equora, we believe every client deserves a personalized  roadmap toward financial success. Our planning '
    + 'services are  designed to align with your life goals, business ambitions, and  long-term vision. We analyze '
    + 'your financial situation, identify  opportunities, and craft strategies that bring measurable  results.',
    { fontSize: 15, color: C.navy });

  SERVICES.forEach(([x, y, , , , badgeColor, glyphColor, titleColor, bodyColor, title, body, tw]) => {
    badge(s, x + 0.419, y + 1.666, 0.833, badgeColor, glyphColor);
    label(s, x + 0.319, y + 2.865, tw, 0.454, title, { fontSize: 21, bold: true, color: titleColor });
    para(s, x + 0.319, y + 3.401, 4.724, 0.787, body, { fontSize: 15, color: bodyColor });
  });

  chrome(s);
}

/* --- Slide 05 - Asset management approach (donut) --- */
const ALLOCATION = [
  // rowY,  pct,   name,           color,   pctW,  nameW, nameX
  [3.500, '45%', 'Equity', C.amber, 1.128, 1.115, 2.900],
  [5.177, '25%', 'Bonds', C.teal, 1.122, 1.096, 2.900],
  [6.854, '20%', 'Real Estate', C.red, 1.154, 1.820, 2.931],
  [8.531, '10%', 'Alternatives', C.green, 1.084, 1.969, 2.858],
];

function slide05(pptx) {
  const s = newSlide(pptx);
  card(s, 1.042, 1.042, 17.917, 9.167, { adj: 0.02121, fill: C.plum, alpha: 95 });

  notch(s, 'br_h', 11.385, 1.042, 0.5, 0.25, C.dark);
  notch(s, 'br_v', 1.042, 3.042, 0.25, 0.5, C.dark);
  card(s, 1.042, 1.042, 10.594, 2.25, { adj: 0.08907, fill: C.dark });
  card(s, 1.042, 1.042, 10.344, 2.0, { adj: 0.0873, fill: C.white, line: C.black });
  label(s, 1.358, 1.318, 9.821, 0.909, 'Asset Management Approach', { fontSize: 48, bold: true, color: C.dark });
  label(s, 1.358, 2.276, 3.133, 0.353, 'Balancing growth and security', { fontSize: 15, color: C.navy });

  ALLOCATION.forEach(([y]) => card(s, 1.25, y, 10.375, 1.469, { adj: 0.16667, fill: C.light, line: C.black, lineAlpha: 5 }));
  ALLOCATION.forEach(([y, pct, name, color, pctW, nameW, nameX]) => {
    label(s, 1.567, y + 0.386, pctW, 0.631, pct, { fontSize: 31.5, bold: true, color });
    label(s, nameX, y + 0.308, nameW, 0.454, name, { fontSize: 21, bold: true, color: C.dark });
    label(s, nameX, y + 0.761, 6.816, 0.353, LOREM, { fontSize: 15, color: C.navy });
    badge(s, 10.375, y + 0.318, 0.833, color, color);
  });

  // Donut: white ring, amber base circle, then three wedges laid on top.
  circle(s, 11.885, 3.122, 6.878, C.light);
  circle(s, 12.13, 3.366, 6.389, C.amber);
  [[72.02, C.teal], [161.91, C.red], [233.68, C.green]].forEach(([start, color]) => {
    s.addShape('pie', { x: 12.13, y: 3.366, w: 6.389, h: 6.389, angleRange: [start, 270], fill: { color } });
  });
  [[14.069, 3.930, 1.000, '10%'], [12.796, 5.799, 1.063, '20%'],
    [14.028, 8.055, 1.035, '25%'], [16.448, 6.086, 1.040, '45%']]
    .forEach(([x, y, w, t]) => label(s, x, y, w, 0.581, t, { fontSize: 28.5, bold: true, color: C.white }));

  chrome(s);
}

/* --- Slide 06 - Risk analysis framework (radar) --- */
const RISK_NODES = [
  // badgeX, badgeY, color, glyph, titleX, titleW, title, pctX, pct, pctW
  [8.667, 1.708, C.plum, 'spark', 7.830, 1.618, 'Market Risk', 8.814, '70%', 0.624],
  [15.958, 1.708, C.amber, 'creditcard', 15.858, 1.585, 'Credit Risk ', 15.858, '60%', 0.649],
  [8.667, 6.240, C.red, 'cup', 7.625, 1.827, 'Liquidity Risk', 8.796, '40%', 0.644],
  [15.958, 6.240, C.teal, 'gear', 15.858, 2.213, 'Operational Risk', 15.858, '50%', 0.640],
];
const RISK_CARDS = [
  // x,     fill,    lined, titleColor, bodyColor, title,            line1,                      line2
  [6.333, C.plum, false, C.white, C.light, 'Market Risk', 'Exposure to stock market', 'fluctuations '],
  [9.542, C.white, true, C.dark, C.navy, 'Liquidity Risk', 'Ability to convert assets', 'into cash without loss '],
  [12.750, C.white, true, C.dark, C.navy, 'Credit Risk ', 'Probability of default from ', 'issuers or borrowers'],
  [15.958, C.white, true, C.dark, C.navy, 'Operational Risk', 'Internal processes,', 'systems '],
];

function slide06(pptx) {
  const s = newSlide(pptx);
  card(s, 1.042, 1.042, 5.083, 9.177, { fill: C.light, line: C.black, lineAlpha: 5 });
  card(s, 6.333, 1.042, 12.625, 7.354, { fill: C.light, alpha: 5 });

  // Radar grid: concentric rings around (12.6455, 4.719) plus the two axes.
  [6.335, 5.016, 3.697, 2.378, 1.059].forEach((d) => {
    s.addShape('ellipse', { x: 12.6455 - d / 2, y: 4.719 - d / 2, w: d, h: d, fill: NO_FILL, line: { color: C.light, width: 0.75, transparency: 80 } });
  });
  line(s, 9.478, 4.719, 6.335, 0, C.light, { alpha: 20 });
  line(s, 12.646, 1.551, 0, 6.336, C.light, { alpha: 20 });

  // Score polygon: market 70%, credit 60%, operational 50%, liquidity 40%.
  freeform(s, [[0.2226, 0], [1, 0.6784], [0.2226, 1], [0, 0.6784], [0.2226, 0]],
    12.116, 2.211, 2.378, 3.697, { line: C.light, lineAlpha: 30, lw: 3 });
  [[12.027, 4.629, C.red], [12.556, 5.818, C.teal], [14.405, 4.629, C.amber], [12.556, 2.121, C.plum]]
    .forEach(([x, y, color]) => circle(s, x, y, 0.179, color));

  RISK_NODES.forEach(([bx, by, color, glyph, tx, tw, title, px, pct, pctW]) => {
    badge(s, bx, by, 0.667, color, C.light, glyph);
    label(s, tx, by + 0.743, tw, 0.404, title, { fontSize: 18, bold: true, color: C.light });
    label(s, px, by + 1.141, pctW, 0.353, pct, { fontSize: 15, color: C.light });
  });

  RISK_CARDS.forEach(([x, fill, lined, titleColor, bodyColor, title, l1, l2]) => {
    card(s, x, 8.604, 3.0, 1.614, { adj: 0.1, fill, line: lined ? C.black : null, lineAlpha: 5 });
    label(s, x + 0.109, 8.857, 2.756, 0.404, title, { fontSize: 18, bold: true, color: titleColor });
    lines(s, x + 0.109, 9.256, 2.756, 0.669, [l1, l2], { fontSize: 15, color: bodyColor });
  });

  para(s, 1.358, 1.318, 4.634, 1.688, 'Risk Analysis  Framework', { fontSize: 48, bold: true, color: C.dark });
  label(s, 1.358, 3.006, 4.032, 0.353, 'Identifying and mitigating financial risks', { fontSize: 15, color: C.navy });
  para(s, 1.358, 4.917, 4.523, 1.65,
    'In finance, risks are inevitable—but they can  be anticipated, measured, and managed  effectively. '
    + 'At Equora, we use a structured  framework to assess risks across multiple  dimensions, ensuring that '
    + 'your portfolio is  both resilient and adaptable.',
    { fontSize: 15, color: C.navy });
  para(s, 1.358, 8.378, 4.51, 1.543,
    'By applying continuous monitoring  and dynamic adjustments, we  minimize vulnerabilities while  '
    + 'preserving growth potential.',
    { fontSize: 18, bold: true, color: C.navy });

  chrome(s);
}

/* --- Slide 07 - Market trends & insights (line chart) --- */
const ROI_YEARS = [
  // gridX,  labelX, labelW, label, markerX, markerY
  [3.365, 3.034, 0.661, '2019', 3.273, 6.814],
  [5.596, 5.242, 0.696, '2020', 5.505, 5.110],
  [7.828, 7.490, 0.651, '2021', 7.736, 5.962],
  [10.060, 9.696, 0.688, '2022', 9.968, 3.405],
  [12.292, 11.922, 0.684, '2023', 12.200, 1.700],
];
// ROI axis ticks: x, width, label (11 down to 5).
const ROI_TICKS = [
  [1.788, 0.367, '11'], [1.765, 0.412, '10'], [1.804, 0.332, '9'], [1.808, 0.325, '8'],
  [1.816, 0.307, '7'], [1.804, 0.332, '6'], [1.809, 0.323, '5'],
];

function slide07(pptx) {
  const s = newSlide(pptx);
  card(s, 1.042, 1.042, 17.917, 6.958, { adj: 0.04226, fill: C.white });

  label(s, 13.936, 2.17, 3.091, 0.404, 'Key Insights 2019–2023', { fontSize: 18, bold: true, color: C.dark });
  bullets(s, 13.958, 2.631, 4.75, 2.25, [
    'Global investment returns have shown  consistent upward momentum ',
    'Diversified portfolios outperform single- asset strategies ',
    'Market volatility has decreased in the past 2  years ',
    'Sustainable investments are gaining  stronger traction among investors',
  ], { fontSize: 15, color: C.navy });

  card(s, 1.042, 8.208, 12.667, 2.0, { adj: 0.14428, fill: C.light, alpha: 5 });
  label(s, 1.358, 8.484, 8.285, 0.909, 'Market Trends & Insights', { fontSize: 48, bold: true, color: C.light });
  label(s, 1.358, 9.443, 3.578, 0.353, 'Key data driving financial decisions', { fontSize: 15, color: C.light });

  notch(s, 'tl_h', 13.458, 7.75, 0.5, 0.25, C.light);
  notch(s, 'tl_v', 18.708, 6.354, 0.25, 0.5, C.light);
  card(s, 13.708, 6.604, 5.25, 1.396, { adj: 0.15693, fill: C.light });
  card(s, 13.958, 6.854, 5.0, 3.333, { fill: C.plum });
  para(s, 14.329, 7.483, 4.313, 2.26,
    'Strong financial decisions are built on  reliable market insights. At Equora, we  continuously analyze global '
    + 'and local  financial indicators to identify  opportunities and anticipate risks. This  allows us to provide '
    + 'clients with  actionable intelligence for smarter  investment strategies.',
    { fontSize: 15, color: C.light });

  label(s, 1.058, 4.344, 0.947, 0.353, 'ROI (%)', { fontSize: 15, color: C.plum, rotate: 270 });
  ROI_TICKS.forEach(([x, w, t], i) => label(s, x, 1.589 + i * 0.8507, w, 0.353, t, { fontSize: 15, color: C.navy }));

  // Plot frame: rounded L-shaped axis plus the grid.
  freeform(s, [[0, 0], [0, 0.9787], ['c', 0, 0.9904, 0.005, 1, 0.0112, 1], [1, 1]],
    2.25, 1.375, 11.125, 5.855, { line: C.dark, lw: 1.5, close: false });
  ROI_YEARS.forEach(([gx]) => line(s, gx, 1.375, 0, 5.854, C.black, { alpha: 20 }));
  [1.792, 2.644, 3.497, 4.349, 5.201, 6.054, 6.906].forEach(y => line(s, 2.25, y, 11.125, 0, C.black, { alpha: 20 }));
  ROI_YEARS.forEach(([, lx, lw, t]) => label(s, lx, 7.319, lw, 0.353, t, { fontSize: 15, color: C.navy }));

  // ROI series: 5, 7, 6, 9, 11 (%)
  freeform(s, [[0, 1], [0.25, 0.6653], [0.4988, 0.8306], [0.7488, 0.3347], [1, 0]],
    3.362, 1.789, 8.936, 5.124, { line: C.teal, lw: 3.75, close: false });
  ROI_YEARS.forEach(([, , , , mx, my]) => circle(s, mx, my, 0.184, C.teal));

  chrome(s);
}

/* --- Slide 08 - ROI projections (bars) --- */
const PROJECTIONS = [
  // cardX, color,  pct,   pctX,   pctW,  trackH, fillY,  fillH, title
  [6.333, C.plum, '8%', 9.223, 0.870, 4.112, 6.445, 0.346, '1 Year Projection'],
  [10.611, C.teal, '22%', 13.251, 1.126, 3.506, 5.839, 0.952, '3 Year Projection'],
  [14.889, C.red, '38%', 17.528, 1.128, 2.814, 5.147, 1.645, '5 Year Projection'],
];

// Vertical bar with only the top (or bottom) pair of corners rounded.
function bar(s, x, y, w, h, r, roundTop, o) {
  const cmds = roundTop
    ? [[0, r / h], ['c', 0, 0.4477 * r / h, 0.4477 * r / w, 0, r / w, 0], [1 - r / w, 0],
      ['c', 1 - 0.4477 * r / w, 0, 1, 0.4477 * r / h, 1, r / h], [1, 1], [0, 1], [0, r / h]]
    : [[0, 0], [1, 0], [1, 1 - r / h], ['c', 1, 1 - 0.4477 * r / h, 1 - 0.4477 * r / w, 1, 1 - r / w, 1],
      [r / w, 1], ['c', 0.4477 * r / w, 1, 0, 1 - 0.4477 * r / h, 0, 1 - r / h], [0, 0]];
  freeform(s, cmds, x, y, w, h, o);
}

function slide08(pptx) {
  const s = newSlide(pptx);
  card(s, 1.042, 1.042, 5.083, 9.177, { fill: C.white, line: C.black, lineAlpha: 5 });
  PROJECTIONS.forEach(([x]) => card(s, x, 1.042, 4.069, 7.5, { fill: C.light, line: C.black }));
  card(s, 6.333, 8.75, 12.625, 1.458, { adj: 0.1, fill: C.light, alpha: 5 });
  lines(s, 6.65, 9.131, 12.054, 0.606, [
    'These projections demonstrate the potential of a well-diversified portfolio, balancing growth with manageable risk. Clients ',
    'can track progress year after year and adjust strategies to remain aligned with their financial goals.',
  ], { fontSize: 15, color: C.light });

  PROJECTIONS.forEach(([x, color, pct, pctX, pctW, trackH, fillY, fillH, title]) => {
    badge(s, x + 0.417, 1.458, 0.667, color, color);
    label(s, pctX, 1.443, pctW, 0.631, pct, { fontSize: 31.5, bold: true, color });
    bar(s, x + 0.417, 2.333, 3.24, trackH, 0.1667, true, { fill: C.black, alpha: 5 });
    bar(s, x + 0.417, fillY, 3.24, fillH, 0.1667, false, { fill: color });
    label(s, x + 0.317, 6.949, 3.445, 0.454, title, { fontSize: 21, bold: true, color: C.navy });
    lines(s, x + 0.317, 7.464, 3.445, 0.606, ['Lorem ipsum dolor sit amet,', 'consectetur adipiscing elit. '],
      { fontSize: 15, color: C.navy });
  });

  para(s, 1.358, 1.443, 3.983, 1.701, 'ROI  Projections', { fontSize: 48, bold: true, color: C.dark });
  label(s, 1.358, 3.194, 2.591, 0.353, 'Forecasting your returns', { fontSize: 15, color: C.navy });
  badge(s, 1.458, 6.888, 1.042, C.plum, C.plum, 'folder');
  para(s, 1.358, 8.204, 4.388, 1.701,
    'At Equora, we believe in transparency and  measurable outcomes. By analyzing market  trends, historical '
    + 'performance, and  diversified investment strategies, we  provide clear ROI projections that help  clients '
    + 'make informed financial decisions.',
    { fontSize: 15, color: C.navy });

  chrome(s);
}

/* --- Slide 09 - Case study 1 --- */
const CASE1_STATS = [
  // cardX,  value,   valueW, caption,                  captionW
  [8.167, '$2.0M', 1.100, 'Initial Portfolio (2019)', 2.279],
  [11.765, '$2.7M', 1.063, 'After 3 Years (2022)', 2.143],
  [15.362, '+35% ', 1.038, 'Growth in value', 1.699],
];

function slide09(pptx) {
  const s = newSlide(pptx);
  imageBox(s, 1.042, 1.026, 17.917, 9.167, 0.289);

  notch(s, 'br_h', 7.719, 1.042, 0.5, 0.25, C.dark);
  notch(s, 'br_v', 1.042, 3.667, 0.25, 0.5, C.dark);
  card(s, 1.042, 1.042, 6.927, 2.875, { adj: 0.09399, fill: C.dark });
  card(s, 1.042, 1.042, 6.677, 2.625, { adj: 0.05295, fill: C.light, line: C.black });
  label(s, 1.358, 1.459, 2.027, 0.454, 'Case Study 1', { fontSize: 21, bold: true, color: C.navy });
  label(s, 1.358, 1.89, 5.116, 0.909, 'Wealth Growth', { fontSize: 48, bold: true, color: C.dark });
  label(s, 1.358, 2.849, 4.167, 0.353, 'How we increased client portfolio by 35%', { fontSize: 15, color: C.navy });

  card(s, 1.25, 4.125, 6.71, 5.875, { fill: C.plum });
  CASE1_STATS.forEach(([x]) => card(s, x, 7.396, 3.388, 2.604, { fill: C.light, line: C.black, lineAlpha: 5 }));

  label(s, 1.567, 4.493, 2.893, 0.404, 'Before & After Results', { fontSize: 18, bold: true, color: C.white });
  para(s, 1.567, 4.986, 6.154, 1.505,
    'Client A approached Equora with a goal to grow their wealth  sustainably while minimizing risks. After '
    + 'conducting a  comprehensive financial assessment, we designed a  and real estate. diversified investment '
    + 'strategy that balanced equity, bonds, ',
    { fontSize: 15, color: C.light });
  label(s, 1.567, 6.941, 2.989, 0.404, 'Key Strategies Applied', { fontSize: 18, bold: true, color: C.white });
  bullets(s, 1.567, 7.401, 6.154, 1.01, [
    'Diversification across high-performing sectors ',
    'Tax-efficient investment planning ',
    'Active monitoring with quarterly portfolio adjustments',
  ], { fontSize: 15, color: C.light });
  para(s, 1.567, 8.61, 6.154, 1.01,
    'This case demonstrates how tailored financial planning can  deliver sustainable growth even in dynamic market  conditions.',
    { fontSize: 15, color: C.light });

  CASE1_STATS.forEach(([x, value, valueW, caption, captionW]) => {
    badge(s, x + 0.333, 7.729, 0.833, C.plum, C.plum);
    label(s, x + 0.233, 8.782, valueW, 0.454, value, { fontSize: 21, bold: true, color: C.dark });
    label(s, x + 0.233, 9.318, captionW, 0.353, caption, { fontSize: 15, color: C.navy });
  });

  chrome(s);
}

/* --- Slide 10 - Case study 2 --- */
const CASE2_STEPS = [
  [8.167, 'Assess Risk', 'Identified exposure points  in equities and foreign  exchange'],
  [11.765, 'Diversify', 'Shifted 30% of assets into  bonds and alternative  investments'],
  [15.362, 'Monitor', 'Established real-time  monitoring and quarterly  reviews'],
];

function slide10(pptx) {
  const s = newSlide(pptx);
  imageBox(s, 1.042, 1.026, 17.917, 9.167, 0.289);

  notch(s, 'tr_h', 7.719, 9.958, 0.5, 0.25, C.dark);
  notch(s, 'tr_v', 1.042, 7.083, 0.25, 0.5, C.dark);
  card(s, 1.042, 7.333, 6.927, 2.875, { fill: C.dark });
  card(s, 1.042, 7.583, 6.677, 2.625, { fill: C.light, line: C.black });
  label(s, 1.358, 8.001, 2.072, 0.454, 'Case Study 2', { fontSize: 21, bold: true, color: C.navy });
  label(s, 1.358, 8.432, 4.195, 0.909, 'Risk Control', { fontSize: 48, bold: true, color: C.dark });
  label(s, 1.358, 9.391, 4.083, 0.353, 'Mitigating volatility in uncertain markets', { fontSize: 15, color: C.navy });

  CASE2_STEPS.forEach(([x], i) => card(s, x, i === 2 ? 7.385 : 7.396, 3.388, 2.604, { fill: C.light, line: C.black, lineAlpha: 5 }));
  card(s, 12.25, 3.823, 6.5, 3.344, { fill: C.plum });
  para(s, 12.478, 4.212, 6.108, 1.034,
    'Equora implemented a risk control framework to  minimize losses while preserving growth  opportunities.',
    { fontSize: 18, bold: true, color: C.white });
  para(s, 12.478, 5.464, 6.108, 1.26,
    'Client B, a mid-sized corporation, was exposed to high  market volatility during a period of economic '
    + 'uncertainty.  Their portfolio faced rapid fluctuations that threatened long- term financial stability.',
    { fontSize: 15, color: C.light });

  CASE2_STEPS.forEach(([x, title, body]) => {
    label(s, x + 0.233, 7.678, 3.004, 0.454, title, { fontSize: 21, bold: true, color: C.dark });
    para(s, x + 0.233, 8.693, 3.004, 1.181, body, { fontSize: 15, color: C.navy });
  });

  chrome(s);
}

/* --- Slide 11 - Why choose Equora --- */
function slide11(pptx) {
  const s = newSlide(pptx);
  card(s, 1.042, 1.036, 5.833, 9.177, { fill: C.light, line: C.black, lineAlpha: 5 });
  card(s, 7.083, 1.042, 5.833, 2.917, { fill: C.plum });
  card(s, 7.083, 4.167, 5.833, 2.917, { fill: C.light, alpha: 5 });
  card(s, 7.083, 7.292, 5.833, 2.917, { fill: C.light, alpha: 5 });
  card(s, 13.125, 1.036, 5.833, 4.484, { fill: C.white, line: C.black, lineAlpha: 5 });
  card(s, 13.125, 5.729, 5.833, 4.484, { fill: C.white, line: C.black, lineAlpha: 5 });

  lines(s, 1.358, 1.318, 5.308, 1.717, ['Why Choose ', 'Equora'], { wrap: true, fontSize: 48, bold: true, color: C.dark });
  label(s, 1.358, 3.006, 2.744, 0.353, 'Clients trust our expertise', { fontSize: 15, color: C.navy });
  para(s, 1.358, 8.204, 5.316, 1.729,
    'Choosing the right financial consultant can make  the difference between uncertainty and long-term  '
    + 'stability. Equora stands out because we combine  experience, data-driven strategies, and a client-first  '
    + 'approach that ensures measurable impact.',
    { fontSize: 15, color: C.navy });
  para(s, 13.442, 4.131, 5.263, 1.176,
    'We don’t offer generic solutions—we create tailored  strategies that align with your vision of financial  success.',
    { fontSize: 15, color: C.navy });

  badge(s, 7.5, 7.865, 0.833, C.plum, C.light, 'user');
  label(s, 7.4, 8.773, 4.383, 0.581, '500+ satisfied clients', { fontSize: 28.5, bold: true, color: C.light });
  label(s, 7.4, 9.454, 3.643, 0.353, 'Across individuals and corporations', { fontSize: 15, color: C.light });

  badge(s, 13.542, 7.656, 1.042, C.teal, C.teal);
  label(s, 13.442, 8.773, 4.813, 0.581, 'Personalized strategies', { fontSize: 28.5, bold: true, color: C.dark });
  label(s, 13.442, 9.454, 3.826, 0.353, 'Crafted for each unique financial goal', { fontSize: 15, color: C.navy });

  badge(s, 7.5, 4.74, 0.833, C.teal, C.light, 'calendar');
  label(s, 7.4, 5.648, 4.25, 0.581, '15+ years experience', { fontSize: 28.5, bold: true, color: C.light });
  label(s, 7.4, 6.329, 4.486, 0.353, 'Combined experience in financial consulting', { fontSize: 15, color: C.light });

  badge(s, 7.5, 1.615, 0.833, C.white, C.light, 'globe');
  label(s, 7.4, 2.523, 3.179, 0.581, 'Global insights ', { fontSize: 28.5, bold: true, color: C.white });
  label(s, 7.4, 3.204, 3.13, 0.353, 'adapted to local opportunities', { fontSize: 15, color: C.light });

  chrome(s);
  imageBox(s, 13.542, 1.443, 5.0, 2.474, 0.0787);
}

/* --- Slide 12 - Testimonials --- */
const QUOTES = [
  [2.958, C.light, 5, null, C.light, 3.275, 4.390, [
    '"Equora transformed our financial ', 'strategy, giving us clarity and ',
    'confidence in every decision." – ', 'Client X, Entrepreneur'], 2.502, '– Client X, Entrepreneur', 9.127, 3.375],
  [8.361, C.light, 100, C.black, C.navy, 8.692, 4.443, [
    '"Their proactive risk management', 'helped us stabilize during market',
    'uncertainty without losing growth ', 'potential."'], 3.428, '– Client Y, CFO of a mid-sized firm', 9.142, 8.792],
  [13.764, C.plum, 100, null, C.light, 14.108, 4.602, [
    '"With Equora’s retirement planning,', 'I now have peace of mind knowing ',
    'my dream bussiness future is ', 'secure." '], 3.258, '– Client Z, Professional Investor', 9.142, 14.208],
];

function slide12(pptx) {
  const s = newSlide(pptx);
  card(s, 1.042, 1.042, 17.917, 6.958, { adj: 0.04392, fill: C.white });
  badge(s, 1.458, 1.646, 1.042, C.plum, C.plum);
  label(s, 2.817, 1.443, 5.256, 0.909, 'ROI Projections', { fontSize: 48, bold: true, color: C.dark });
  label(s, 2.817, 2.401, 2.591, 0.353, 'Forecasting your returns', { fontSize: 15, color: C.navy });

  notch(s, 'tl_h', 2.458, 7.75, 0.5, 0.25, C.dark);
  notch(s, 'tl_v', 18.708, 3.042, 0.25, 0.5, C.dark);
  card(s, 2.708, 3.292, 16.25, 4.708, { adj: 0.05163, fill: C.dark });
  QUOTES.forEach(([x, fill, alpha, ln]) => card(s, x, 3.542, 5.194, 6.664, { fill, alpha, line: ln }));

  QUOTES.forEach(([, , , , color, tx, tw, quote, cw, credit, cy]) => {
    lines(s, tx, 7.722, tw, 1.313, quote, { fontSize: 18, bold: true, color });
    label(s, tx, cy, cw, 0.353, credit, { fontSize: 15, color });
  });

  chrome(s);
  lines(s, 14.192, 1.662, 4.46, 0.858, [
    'Our clients’ success is the true measure of',
    'our work. Here’s what some of them have to ',
    ' say about partnering with Equora:',
  ], { fontSize: 15, color: C.navy });
  QUOTES.forEach(([, , , , , , , , , , , picX]) => imageBox(s, picX, 5.354, 2.083, 2.083, 0.29));
}

/* --- Slide 13 - Our team of experts --- */
const TEAM_CARDS = [
  // x,      y,     fill,    alpha, line,    badgeColor, glyphColor, headColor, bodyColor, heading lines, headW
  [6.823, 3.542, C.light, 5, null, C.plum, C.light, C.light, C.light,
    ['A culture of innovation and', 'continuous learning '], 4.055],
  [12.995, 3.542, C.light, 100, C.black, C.plum, C.plum, C.dark, C.navy,
    ['Specialists in wealth planning,', 'investment strategy, and advisory '], 5.161],
  [12.995, 6.979, C.light, 5, null, C.plum, C.light, C.light, C.light,
    ['Commitment to client-first service', 'and measurable results '], 5.179],
];

function slide13(pptx) {
  const s = newSlide(pptx);
  card(s, 1.042, 1.042, 17.917, 6.958, { fill: C.white });
  notch(s, 'tl_h', 6.323, 7.75, 0.5, 0.25, C.dark);
  notch(s, 'tl_v', 18.708, 3.042, 0.25, 0.5, C.dark);
  card(s, 6.573, 3.292, 12.385, 4.708, { fill: C.dark });
  TEAM_CARDS.forEach(([x, y, fill, alpha, ln]) => card(s, x, y, 5.964, 3.229, { fill, alpha, line: ln }));
  card(s, 6.823, 6.979, 5.964, 3.229, { fill: C.plum });

  label(s, 1.358, 1.443, 6.788, 0.909, 'Our Team of Experts', { fontSize: 48, bold: true, color: C.dark });
  label(s, 1.358, 2.401, 4.604, 0.353, 'Experienced financial advisors at your service', { fontSize: 15, color: C.navy });
  lines(s, 14.192, 1.818, 4.385, 0.606, [
    'At Equora, we don’t just manage numbers—', 'we build trusted relationships that last.',
  ], { fontSize: 15, color: C.navy });

  badge(s, 1.417, 4.24, 1.042, C.teal, C.teal, 'team');
  lines(s, 1.358, 5.61, 4.863, 1.616, [
    'Behind every successful strategy at Equora is a',
    'team of dedicated professionals. Our ',
    'consultants bring deep expertise in financial ',
    'planning, asset management, and risk analysis—',
    'working together to craft solutions tailored for ',
    'each client.',
  ], { fontSize: 15, color: C.navy });

  // Plum "20+ Certified" tile (bottom left of the dark panel).
  label(s, 7.14, 8.45, 2.742, 0.581, '20+ Certified', { fontSize: 28.5, bold: true, color: C.light });
  lines(s, 7.14, 9.131, 4.327, 0.606, ['Certified consultants with diverse industry', 'backgrounds '],
    { fontSize: 15, color: C.light });

  TEAM_CARDS.forEach(([x, y, , , , badgeColor, glyphColor, headColor, bodyColor, heading, headW]) => {
    badge(s, x + 0.417, y + 0.583, 0.667, badgeColor, glyphColor);
    lines(s, x + 0.317, y + 1.532, headW, 0.808, heading, { fontSize: 21, bold: true, color: headColor });
    label(s, x + 0.317, y + 2.464, 2.863, 0.353, LOREM_SHORT, { fontSize: 15, color: bodyColor });
  });

  chrome(s);
  imageBox(s, 1.048, 8.212, 5.525, 1.995, 0.158);
}

/* --- Slide 14 - Next steps --- */
const NEXT_STEPS = [
  // cardX, cardFill, cardAlpha, cardLine, circleColor, numColor, titleColor, bodyColor, num, title, titleW, body
  [1.042, C.plum, 100, null, C.white, C.white, C.white, C.light, '01', 'Schedule a Consultation', 3.726,
    'Meet our experts to discuss your goals', 1.614, 0.526, 3.95],
  [7.083, C.white, 100, null, C.plum, C.plum, C.dark, C.navy, '02', 'Get a Tailored Plan', 2.868,
    'Receive a strategy aligned with your needs', 7.629, 0.572, 4.344],
  [13.125, C.light, 100, C.black, C.teal, C.teal, C.dark, C.navy, '03', 'Achieve Results', 2.498,
    'Start building wealth with confidence and clarity', 13.676, 0.568, 4.86],
];

function slide14(pptx) {
  const s = newSlide(pptx);
  card(s, 1.042, 1.042, 8.854, 4.479, { fill: C.light, alpha: 5 });
  card(s, 10.104, 1.042, 8.854, 4.479, { fill: C.white, line: C.black, lineAlpha: 5 });
  NEXT_STEPS.forEach(([x, fill, alpha, ln]) => card(s, x, 5.729, 5.833, 4.479, { fill, alpha, line: ln }));

  badge(s, 1.458, 1.458, 1.042, C.plum, C.light);
  label(s, 1.358, 2.765, 3.808, 0.909, 'Next Steps', { fontSize: 48, bold: true, color: C.light });
  label(s, 1.358, 3.724, 3.796, 0.353, 'Let’s start planning your future today', { fontSize: 15, color: C.light });
  lines(s, 1.358, 4.443, 7.465, 0.606, [
    'Together, we’ll create a roadmap that secures your future while maximizing', 'opportunities today. ',
  ], { fontSize: 15, color: C.light });

  NEXT_STEPS.forEach(([x, , , , circleColor, numColor, titleColor, bodyColor, num, title, titleW, body, numX, numW, bodyW]) => {
    circle(s, x + 0.417, 7.917, 0.833, circleColor, 40);
    label(s, numX, 8.132, numW, 0.454, num, { fontSize: 21, bold: true, color: numColor });
    label(s, x + 0.317, 8.907, titleW, 0.454, title, { fontSize: 21, bold: true, color: titleColor });
    label(s, x + 0.317, 9.443, bodyW, 0.353, body, { fontSize: 15, color: bodyColor });
  });

  label(s, 10.421, 2.381, 5.281, 0.581, 'Take the next step with us', { fontSize: 28.5, bold: true, color: C.dark });
  lines(s, 10.421, 3.062, 8.376, 0.858, [
    'PLACEHOLDER',
    'how to achieve them. At Equora, we are ready to guide you through every step—from ',
    'planning and investment to risk management and long-term growth.',
  ], { fontSize: 15, color: C.navy });
  pill(s, 10.521, 4.452, 4.168, 0.652, C.plum);
  label(s, 10.845, 4.561, 3.578, 0.454, 'Book Your Consultation', { fontSize: 21, bold: true, color: C.white });

  chrome(s);
}

/* --- Slide 15 - Thank you & contact --- */
const CONTACTS = [
  // rowY,  label,     labelY, labelW, value,             valueX,  valueY, valueW
  [6.417, 'Email', 6.732, 0.861, 'info@equora.com', 13.625, 6.761, 1.925],
  [7.681, 'Website', 8.014, 1.224, 'www.equora.com', 13.959, 8.042, 1.899],
  [8.944, 'Phone', 9.284, 0.987, '+1 234 567 890', 13.750, 9.313, 1.627],
];

function slide15(pptx) {
  const s = newSlide(pptx);
  imageBox(s, 1.042, 1.026, 17.917, 9.167, 0.289);

  notch(s, 'tr_h', 11.385, 9.958, 0.5, 0.25, C.dark);
  notch(s, 'tr_v', 1.042, 5.156, 0.25, 0.5, C.dark);
  card(s, 1.042, 5.406, 10.594, 4.802, { fill: C.dark });
  card(s, 1.042, 5.656, 10.344, 4.552, { fill: C.light, line: C.black });

  notch(s, 'bl_h', 13.458, 1.042, 0.5, 0.25, C.dark);
  notch(s, 'bl_v', 18.708, 2.427, 0.25, 0.5, C.dark);
  card(s, 13.708, 1.042, 5.25, 1.635, { adj: 0.11265, fill: C.dark });
  card(s, 13.958, 1.042, 5.0, 1.385, { adj: 0.2, fill: C.plum });

  lines(s, 1.426, 6.534, 4.171, 1.717, ['Thank You &', 'Contact '], { fontSize: 48, bold: true, color: C.dark });
  para(s, 1.426, 8.222, 7.197, 1.111,
    'Thank you for considering Equora as your financial consulting partner. We are committed to guiding your '
    + 'wealth, securing your future, and creating strategies that deliver clarity, confidence, and measurable results.',
    { fontSize: 15, color: C.navy });

  pill(s, 8.623, 6.674, 2.278, 0.652, C.plum);
  logoMark(s, 8.956, 6.799, 0.402, C.light);
  label(s, 9.468, 6.751, 1.21, 0.454, 'Equora', { fontSize: 21, bold: true, color: C.white });

  badge(s, 14.292, 1.375, 0.75, C.light, C.light);
  label(s, 15.192, 1.339, 3.554, 0.353, 'Together we build lasting financial ', { fontSize: 15, color: C.white });
  label(s, 15.192, 1.651, 1.012, 0.353, 'success', { fontSize: 15, color: C.white });

  CONTACTS.forEach(([y, name, nameY, nameW, value, valueX, valueY, valueW]) => {
    card(s, 11.833, y, 6.917, 1.056, { adj: 0.2, fill: C.white, line: C.black });
    freeform(s, ARROW, 12.316, y + 0.443, 0.173, 0.173, { fill: C.plum });
    label(s, 12.719, nameY, nameW, 0.404, name, { fontSize: 18, bold: true, color: C.plum });
    label(s, valueX, valueY, valueW, 0.353, value, { fontSize: 15, color: C.navy });
  });

  chrome(s);
}

/* --- Slide 16 - Style sheet --- */
const TYPE_SCALE = [
  [3.090, 0.909, 48, true, 'Heading 1 / Bold / 64', 6.746],
  [4.058, 0.833, 43.5, true, 'Heading 2 / Bold / 58', 6.213],
  [5.063, 0.631, 31.5, true, 'Heading 3 / Bold / 42', 4.544],
  [5.965, 0.581, 28.5, true, 'Heading 4 / Bold / 38', 4.148],
  [6.801, 0.454, 21, true, 'Heading 5 / Bold / 28', 3.105],
  [7.483, 0.404, 18, true, 'Heading 6 / Bold / 24', 2.697],
  [8.108, 0.391, 17.25, true, 'SubTitle 1 / Bold / 22', 2.563],
  [8.719, 0.366, 15.75, true, 'SubTitle 2 / Bold / 20', 2.404],
  [9.282, 0.379, 16.5, false, 'Body 1 / Regular / 20', 2.330],
  [9.856, 0.353, 15.0, false, 'Body 2 / Regular / 18', 2.137],
];

// Two rows of small greyscale swatches, then the brand palette.
const SHADE_SWATCHES = [
  // x,      key,  keyW,  pct,   pctW,  color,    alpha
  [9.025, 'S1', 0.433, '10%', 0.523, C.black, 10],
  [10.046, 'S2', 0.465, '20%', 0.554, C.black, 20],
  [11.067, 'S3', 0.462, '30%', 0.551, C.black, 30],
  [12.088, 'S4', 0.471, '40%', 0.558, C.black, 40],
  [13.108, 'S5', 0.463, '50%', 0.554, C.black, 50],
  [14.129, 'F1', 0.406, '10%', 0.523, 'FFFFFF', 10],
  [15.098, 'F2', 0.446, '20%', 0.554, 'FFFFFF', 20],
  [16.067, 'F3', 0.449, '30%', 0.551, 'FFFFFF', 30],
  [17.035, 'F4', 0.440, '40%', 0.558, 'FFFFFF', 40],
  [18.004, 'F5', 0.450, '50%', 0.554, 'FFFFFF', 50],
];
const BASE_SWATCHES = [
  [9.025, 'B1', 0.428, 'Dark 1', 0.655, C.dark],
  [10.733, 'W1', 0.496, 'Light 1', 0.687, C.light],
  [12.442, 'B2', 0.468, 'Dark 2', 0.693, C.navy],
  [14.150, 'W2', 0.537, 'Light 2', 0.725, C.white],
];
const ACCENT_SWATCHES = [
  // x,      key,  keyW,  role,       roleW, color,   captionHex, captionW
  [9.025, 'P1', 0.430, 'Accent 1', 0.858, C.plum, '6B3FA0', 0.649],
  [10.535, 'P2', 0.459, 'Accent 2', 0.896, C.amber, 'F4B434', 0.633],
  [12.046, 'P3', 0.462, 'Accent 3', 0.898, C.teal, '2E8B92', 0.666],
  [13.556, 'P4', 0.471, 'Accent 4', 0.901, C.red, 'E15A5A', 0.640],
  [15.067, 'P5', 0.463, 'Accent 5', 0.899, C.green, '3BA76B', 0.667],
  [16.577, 'P6', 0.468, 'Accent 6', 0.896, C.blue, '4BA3E3', 0.645],
];
const LINK_SWATCHES = [
  [9.025, 'H1', 0.447, 'Hyperlink', 0.945, C.link, '3D74E0', 0.642],
  [10.525, 'H2', 0.477, 'Followed Hyperlink', 1.705, C.link, '8A4DE0', 0.663],
];

function swatchCaption(s, x, y, hex, w) {
  label(s, x, y, 0.318, 0.252, '#', { fontSize: 9, color: C.black, transparency: 50 });
  label(s, x + 0.104, y, w, 0.252, hex, { fontSize: 9, color: C.black, transparency: 50 });
}

function slide16(pptx) {
  const s = newSlide(pptx, C.light);
  label(s, 1.192, 0.99, 1.431, 0.415, 'Font Used', { fontSize: 18, bold: true, color: C.black, transparency: 70 });
  label(s, 1.192, 1.441, 2.819, 0.833, 'Manrope', { fontSize: 43.5, bold: true, color: C.black });
  label(s, 1.192, 2.395, 4.649, 0.353, 'https://fonts.google.com/specimen/Manrope',
    { fontSize: 15, color: C.black, transparency: 50 });
  line(s, 1.292, 2.984, 6.917, 0, C.black, { alpha: 20, lw: 1.5 });
  TYPE_SCALE.forEach(([y, h, size, bold, text, w]) => label(s, 1.192, y, w, h, text, { fontSize: size, bold, color: C.black }));

  label(s, 9.025, 0.99, 1.547, 0.415, 'Color Used', { fontSize: 18, bold: true, color: C.black, transparency: 70 });
  SHADE_SWATCHES.forEach(([x, name, nameW, pct, pctW, color, alpha]) => {
    label(s, x, 1.756, nameW, 0.362, name, { fontSize: 15, bold: true, color: C.black, transparency: 20 });
    label(s, x, 2.059, pctW, 0.303, pct, { fontSize: 12, color: C.black, transparency: 50 });
    card(s, x + 0.1, 2.443, 0.604, 0.833, { adj: 0.16667, fill: color, alpha });
    swatchCaption(s, x, 3.309, color, color === C.black ? 0.665 : 0.581);
  });
  BASE_SWATCHES.forEach(([x, name, nameW, role, roleW, color]) => {
    label(s, x, 3.918, nameW, 0.362, name, { fontSize: 15, bold: true, color: C.black, transparency: 20 });
    label(s, x, 4.220, 1.655, 0.31, 'Text / Background ', { fontSize: 12, color: C.black, transparency: 50 });
    label(s, x, 4.428, roleW, 0.31, role, { fontSize: 12, color: C.black, transparency: 50 });
    card(s, x + 0.1, 4.812, 1.25, 0.833, { adj: 0.16667, fill: color });
    swatchCaption(s, x, 5.679, color, 0.663);
  });
  ACCENT_SWATCHES.forEach(([x, name, nameW, role, roleW, color, hex, hexW]) => {
    label(s, x, 6.287, nameW, 0.362, name, { fontSize: 15, bold: true, color: C.black, transparency: 20 });
    label(s, x, 6.590, roleW, 0.31, role, { fontSize: 12, color: C.black, transparency: 50 });
    card(s, x + 0.1, 6.974, 1.25, 0.833, { adj: 0.16667, fill: color });
    swatchCaption(s, x, 7.840, hex, hexW);
  });
  LINK_SWATCHES.forEach(([x, name, nameW, role, roleW, color, hex, hexW]) => {
    label(s, x, 8.449, nameW, 0.362, name, { fontSize: 15, bold: true, color: C.black, transparency: 20 });
    label(s, x, 8.751, roleW, 0.31, role, { fontSize: 12, color: C.black, transparency: 50 });
    card(s, x + 0.1, 9.135, 1.25, 0.833, { adj: 0.16667, fill: color });
    swatchCaption(s, x, 10.001, hex, hexW);
  });
}

const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
  slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16];

const pptx = new PptxGenJS();
pptx.author = 'Equora';
pptx.title = 'Finance Consulting Excellence';
pptx.defineLayout({ name: 'EQUORA', width: SLIDE_W, height: SLIDE_H });
pptx.layout = 'EQUORA';
SLIDES.forEach(build => build(pptx));

pptx.writeFile({ fileName: path.join(__dirname, '05c394cf-45ad-4c30-9a64-19acf6f2a3a0_grok_final.pptx') })
  .then(f => console.log('wrote ' + f));
