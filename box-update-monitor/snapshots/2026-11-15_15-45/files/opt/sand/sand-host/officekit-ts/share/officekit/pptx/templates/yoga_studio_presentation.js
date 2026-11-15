#!/usr/bin/env node
/**
 * "Jooga - Yoga Presentation" deck, rebuilt with pptxgenjs.
 *
 * Everything is plain code: a small palette, a handful of drawing helpers that
 * mirror the recurring motifs of the template (sparkle, dot grid, quarter-round
 * arc, pills, meters), and one builder function per slide.
 *
 * Run: node <this file>   ->  writes the .pptx next to the script.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

const SLIDE_W = 13.3333333;
const SLIDE_H = 7.5;

// ---------------------------------------------------------------- palette ---
const C = {
  BG: 'F9F9F9',        // page background (theme lt1)
  INK: '000000',
  GRAY: '808080',      // body copy
  GREEN: '2E9360',     // headings (accent3 shaded 75%)
  LEAF: '3FC380',      // accent3
  MINT: 'D9F3E6',      // large tint panels
  MINT_MID: 'B2E7CC',  // accent bars / meter fill
  ORANGE: 'F39C12',    // accent2
  AMBER: 'F8C471',     // accent2 tint - hairlines & squares
  SAND: 'FAD7A0',
  PURPLE: 'C39BD3',    // accent1 tint - buttons & badges
  LILAC: 'D7BDE2',
};

const F = {
  BODY: 'Open Sans',   // theme minor font
  FOOT: 'Raleway',     // running footer
  NUM: 'Roboto',       // page numbers
};

// -------------------------------------------------------- generic helpers ---
// Plain filled rectangle.
function box(s, x, y, w, h, color) {
  s.addShape('rect', { x, y, w, h, fill: { color }, line: { type: 'none' } });
}

// Rounded rectangle (buttons, badges, pricing cards).
function pill(s, x, y, w, h, color) {
  s.addShape('roundRect', { x, y, w, h, rectRadius: Math.min(w, h) * 0.16,
    fill: { color }, line: { type: 'none' } });
}

function circle(s, x, y, d, color) {
  s.addShape('ellipse', { x, y, w: d, h: d, fill: { color }, line: { type: 'none' } });
}

// 2pt horizontal hairline used under the small link labels.
function rule(s, x, y, w, color) {
  s.addShape('line', { x, y, w, h: 0, line: { color, width: 2 } });
}

// Text with the deck defaults applied. `body` is a string or an array of lines.
function text(s, body, o) {
  const lines = Array.isArray(body) ? body : [body];
  const runs = lines.map((t, i) => ({ text: t, options: { breakLine: i < lines.length - 1 } }));
  const opts = Object.assign({ fontFace: F.BODY, valign: 'top', isTextBox: true }, o);
  if (opts.lineSpacing) {
    opts.lineSpacingMultiple = opts.lineSpacing;
    delete opts.lineSpacing;
  }
  s.addText(runs, opts);
}

// ----------------------------------------------------------- deck motifs ---
// Quarter disc drawn as an outline: straight top/left edges, curved corner.
function arc(s, x, y, size, color, rotate) {
  const k = 0.5523; // circular bezier constant
  s.addShape('custGeom', {
    x, y, w: size, h: size, rotate: rotate || 0,
    fill: { type: 'none' }, line: { color, width: 1 },
    points: [
      { x: 0, y: 0 },
      { x: size, y: size, curve: { type: 'cubic', x1: size * k, y1: 0, x2: size, y2: size * (1 - k) } },
      { x: 0, y: size },
      { x: 0, y: 0 },
      { close: true },
    ],
  });
}

// Card header: rounded on the top two corners only (pricing table).
function cardHead(s, x, y, w, h, color) {
  const r = 0.2544;
  s.addShape('custGeom', {
    x, y, w, h, fill: { color }, line: { type: 'none' },
    points: [
      { x: r, y: 0 },
      { x: w - r, y: 0 },
      { x: w, y: r, curve: { type: 'cubic', x1: w - r * 0.45, y1: 0, x2: w, y2: r * 0.45 } },
      { x: w, y: h },
      { x: 0, y: h },
      { x: 0, y: r },
      { x: r, y: 0, curve: { type: 'cubic', x1: 0, y1: r * 0.45, x2: r * 0.45, y2: 0 } },
      { close: true },
    ],
  });
}

// Two four-point stars in the top-left corner.
function sparkle(s, color) {
  s.addShape('star4', { x: 0.414, y: 0.4, w: 0.425, h: 0.425, fill: { color }, line: { type: 'none' } });
  s.addShape('star4', { x: 0.708, y: 0.693, w: 0.264, h: 0.264, fill: { color }, line: { type: 'none' } });
}

// 3x3 grid of small dots in the bottom-left corner.
function dotGrid(s, color) {
  const x0 = 0.415, y0 = 6.705, step = 0.188, d = 0.063;
  for (let r = 0; r < 3; r++) {
    for (let c = 0; c < 3; c++) circle(s, x0 + c * step, y0 + r * step, d, color);
  }
}

// Pair of small squares at the top-right corner.
function corner(s, color) {
  box(s, 12.481, 0.357, 0.172, 0.172, color);
  box(s, 12.814, 0.357, 0.172, 0.172, color);
}

// Short mint underline that sits above every heading.
function accentBar(s, x, y, h) {
  box(s, x, y, 0.8, h || 0.03, C.MINT_MID);
}

// Progress meter: tinted track with a mint fill on top.
function meter(s, x, y, w, fillW, h, trackColor) {
  box(s, x, y, w, h, trackColor);
  box(s, x, y, fillW, h, C.MINT_MID);
}

// Purple rounded badge carrying a number such as "01".
function badge(s, x, y, w, h, label, size) {
  pill(s, x, y, w, h, C.PURPLE);
  text(s, label, { x: x + 0.131, y: y + 0.1, w: w - 0.226, h: h - 0.182,
    fontSize: size, color: C.BG, charSpacing: 3, wrap: false });
}

// The recurring "Learn More" button.
function learnMore(s, x, y) {
  pill(s, x, y, 1.981, 0.442, C.PURPLE);
  text(s, 'Learn More', { x: x + 0.19, y: y + 0.061, w: 1.601, h: 0.32,
    fontSize: 13, color: C.BG, charSpacing: 3, wrap: false });
}

// Running footer caption.
function footer(s, x, y, align) {
  text(s, 'Jooga Yoga Presentation', { x, y, w: 2.949, h: 0.286,
    fontSize: 11, color: C.INK, fontFace: F.FOOT, charSpacing: 3,
    align: align || 'left', wrap: false });
}

// Page number, bottom right.
function pageNumber(s, num) {
  text(s, String(num), { x: 12.31, y: 6.942, w: 0.6, h: 0.37,
    fontSize: 16, color: C.GREEN, fontFace: F.NUM, wrap: false });
}

/**
 * Draws the furniture inherited from the slide layout:
 * mint panels, sparkle, dot grid, footer hairline and page number.
 */
function chrome(s, o) {
  s.background = { color: C.BG };
  (o.panels || []).forEach(function (p) { box(s, p[0], p[1], p[2], p[3], p[4]); });
  if (o.sparkle) sparkle(s, o.sparkle);
  if (o.dots) dotGrid(s, o.dots);
  if (o.rule) rule(s, o.rule[0], 6.725, o.rule[1], o.rule[2]);
  if (o.page) pageNumber(s, o.page);
}


// Slide 1 — Title slide
function slide01(s) {
  chrome(s, { panels: [[10.015, 0, 3.318, 7.5, C.MINT]], sparkle: C.ORANGE, dots: C.ORANGE, rule: [9.982, 3.352, C.BG], page: 1 });
  arc(s, 5.458, 4.853, 1.209, C.AMBER, 180);
  arc(s, 10.437, 1.558, 1.209, C.BG);
  circle(s, 8.956, 4.126, 0.727, C.BG);
  circle(s, 9.556, 4.118, 0.727, C.BG);
  footer(s, 5.192, 6.907);
  text(s, "Jooga ", { x: 1.617, y: 2.596, w: 2.446, h: 0.942, fontSize: 50, color: C.GREEN, charSpacing: 3, wrap: false });
  accentBar(s, 1.76, 2.499, 0.079);
  text(s, "Yoga Presentation", { x: 1.541, y: 3.521, w: 2.472, h: 0.32, fontSize: 13, color: C.INK, fontFace: F.FOOT, charSpacing: 3, wrap: false });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing elit  ", { x: 1.516, y: 4.058, w: 3.066, h: 0.942, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  corner(s, C.MINT_MID);
}

// Slide 2 — Yoga Training Program
function slide02(s) {
  chrome(s, { panels: [[0, 4.136, 3.018, 3.364, C.MINT]], sparkle: C.ORANGE, dots: C.ORANGE, rule: [9.982, 3.352, C.AMBER], page: 2 });
  arc(s, 1.026, 1.239, 1.877, C.AMBER, 270);
  rule(s, 6.78, 2.915, 2.839, C.AMBER);
  rule(s, 6.78, 3.539, 2.839, C.AMBER);
  text(s, "Yoga Studio", { x: 6.736, y: 2.536, w: 1.585, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  text(s, "Yoga Training", { x: 6.72, y: 3.158, w: 1.793, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  footer(s, 5.192, 6.907);
  text(s, ["Yoga Training", "Program"], { x: 5.535, y: 4.964, w: 2.665, h: 0.841, fontSize: 22, color: C.GREEN, charSpacing: 3, wrap: false });
  accentBar(s, 5.677, 4.868);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 8.436, y: 4.964, w: 3.25, h: 0.942, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
}

// Slide 3 — Our Yoga Studio History
function slide03(s) {
  chrome(s, { panels: [[0, 0, 3.313, 7.5, C.MINT]], sparkle: C.AMBER, dots: C.AMBER, rule: [9.982, 3.352, C.AMBER], page: 3 });
  rule(s, 6.168, 5.118, 3.228, C.SAND);
  rule(s, 6.168, 5.742, 3.228, C.SAND);
  text(s, "Yoga Studio", { x: 7.92, y: 4.739, w: 1.585, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, align: 'right', wrap: false });
  text(s, "Yoga Training", { x: 7.731, y: 5.361, w: 1.793, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, align: 'right', wrap: false });
  corner(s, C.AMBER);
  text(s, ["Our Yoga", "Studio History"], { x: 6.312, y: 1.487, w: 2.825, h: 0.841, fontSize: 22, color: C.GREEN, charSpacing: 3, align: 'center', wrap: false });
  accentBar(s, 7.331, 1.343);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 6.339, y: 2.427, w: 2.784, h: 1.066, fontSize: 10, color: C.GRAY, align: 'center', lineSpacing: 1.5 });
  learnMore(s, 6.731, 3.765);
  footer(s, 5.192, 6.907);
}

// Slide 4 — Our Yoga Workshop & Events
function slide04(s) {
  chrome(s, { panels: [[0, 0, 3.313, 7.5, C.MINT]], sparkle: C.ORANGE, dots: C.ORANGE, rule: [9.982, 3.352, C.AMBER], page: 4 });
  corner(s, C.AMBER);
  text(s, ["Our Yoga", "Workshop & Events"], { x: 8.691, y: 2.107, w: 3.459, h: 0.774, fontSize: 20, color: C.GREEN, charSpacing: 3, wrap: false });
  accentBar(s, 8.808, 2.011);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 8.691, y: 3.044, w: 3.47, h: 0.942, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  rule(s, 8.778, 4.743, 3.228, C.AMBER);
  rule(s, 8.778, 5.367, 3.228, C.AMBER);
  text(s, "Yoga Studio", { x: 10.529, y: 4.364, w: 1.585, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, align: 'right', wrap: false });
  text(s, "Yoga Training", { x: 10.341, y: 4.986, w: 1.793, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, align: 'right', wrap: false });
  footer(s, 5.192, 6.907);
}

// Slide 5 — Our Yoga Teacher
function slide05(s) {
  chrome(s, { panels: [[10.812, 0, 2.521, 7.5, C.MINT]], sparkle: C.ORANGE, dots: C.ORANGE, rule: [9.982, 3.352, C.AMBER], page: 5 });
  arc(s, 8.716, 0.76, 1.877, C.AMBER, 270);
  rule(s, 5.268, 6.265, 2.913, C.AMBER);
  rule(s, 1.288, 6.272, 3.498, C.AMBER);
  text(s, "Yoga Studio", { x: 5.225, y: 5.886, w: 1.585, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  text(s, "Yoga Training", { x: 1.244, y: 5.89, w: 1.793, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  text(s, ["Our Yoga ", "Teacher"], { x: 1.245, y: 2.227, w: 1.982, h: 0.841, fontSize: 22, color: C.GREEN, charSpacing: 3, wrap: false });
  accentBar(s, 1.367, 2.089);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 4.104, y: 2.227, w: 2.754, h: 1.064, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum  ", { x: 1.245, y: 4.598, w: 2.569, h: 0.813, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  badge(s, 1.335, 3.819, 0.8, 0.586, "01", 18);
  text(s, ["Yoga ", "Experience"], { x: 2.266, y: 3.819, w: 1.631, h: 0.572, fontSize: 14, color: C.PURPLE, charSpacing: 3, wrap: false });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum  ", { x: 4.102, y: 4.598, w: 2.569, h: 0.813, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  badge(s, 4.191, 3.819, 0.8, 0.586, "02", 18);
  text(s, ["Health", "Knowledge"], { x: 5.123, y: 3.819, w: 1.588, h: 0.572, fontSize: 14, color: C.PURPLE, charSpacing: 3, wrap: false });
  corner(s, C.ORANGE);
}

// Slide 6 — The Private Yoga Class
function slide06(s) {
  chrome(s, { panels: [[0, 0, 3.313, 7.5, C.MINT]], sparkle: C.AMBER, dots: C.AMBER, rule: [9.982, 3.352, C.AMBER], page: 6 });
  corner(s, C.AMBER);
  rule(s, 5.979, 5.387, 4.004, C.AMBER);
  rule(s, 5.979, 6.011, 4.004, C.AMBER);
  text(s, "Yoga Studio", { x: 8.506, y: 5.008, w: 1.585, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, align: 'right', wrap: false });
  text(s, "Yoga Training", { x: 8.318, y: 5.63, w: 1.793, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, align: 'right', wrap: false });
  text(s, "The Private Yoga Class", { x: 5.964, y: 2.107, w: 3.741, h: 0.404, fontSize: 18, color: C.GREEN, charSpacing: 3, wrap: false });
  accentBar(s, 6.086, 1.969);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 5.964, y: 2.685, w: 3.843, h: 0.942, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  learnMore(s, 6.086, 3.941);
  footer(s, 5.192, 6.907);
}

// Slide 7 — Let’s Take A Break
function slide07(s) {
  chrome(s, { panels: [[0, 0, 6.667, 7.5, C.MINT]], sparkle: C.AMBER, dots: C.AMBER, rule: [9.982, 3.352, C.AMBER], page: 7 });
  text(s, ["Let’s Take ", "A Break"], { x: 1.333, y: 2.357, w: 2.967, h: 1.279, fontSize: 35, color: C.GREEN, charSpacing: 3, wrap: false });
  accentBar(s, 1.476, 2.175);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 1.333, y: 3.788, w: 2.784, h: 0.942, fontSize: 8.5, color: C.GRAY, lineSpacing: 1.5 });
  learnMore(s, 1.419, 4.883);
  corner(s, C.AMBER);
  footer(s, 6.964, 6.564);
}

// Slide 8 — Our Training Programs
function slide08(s) {
  chrome(s, { panels: [[0, 0, 4.169, 7.5, C.MINT]], sparkle: C.AMBER, dots: C.AMBER, rule: [9.982, 3.352, C.AMBER], page: 8 });
  arc(s, 4.645, 0.891, 1.562, C.AMBER);
  corner(s, C.MINT_MID);
  text(s, ["Our Training", "Programs"], { x: 6.835, y: 2.42, w: 2.494, h: 0.841, fontSize: 22, color: C.GREEN, charSpacing: 3, wrap: false });
  accentBar(s, 6.977, 2.323);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 9.696, y: 2.42, w: 2.784, h: 0.942, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum  ", { x: 6.835, y: 4.66, w: 2.569, h: 0.813, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  badge(s, 6.924, 3.881, 0.8, 0.586, "01", 18);
  text(s, ["Online", "Training"], { x: 7.855, y: 3.881, w: 1.27, h: 0.572, fontSize: 14, color: C.GREEN, charSpacing: 3, wrap: false });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum  ", { x: 9.77, y: 4.66, w: 2.569, h: 0.813, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  badge(s, 9.86, 3.881, 0.8, 0.586, "02", 18);
  text(s, ["Offline", "Training"], { x: 10.791, y: 3.881, w: 1.27, h: 0.572, fontSize: 14, color: C.GREEN, charSpacing: 3, wrap: false });
  footer(s, 5.192, 6.907);
  rule(s, 0.517, 3.476, 1.691, C.AMBER);
  rule(s, 0.517, 4.1, 1.691, C.AMBER);
  text(s, "Yoga Studio", { x: 0.473, y: 3.097, w: 1.585, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  text(s, "Yoga Training", { x: 0.456, y: 3.719, w: 1.793, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
}

// Slide 9 — Our Amazing Services
function slide09(s) {
  chrome(s, { panels: [[10.015, 3.75, 3.318, 3.75, C.MINT]], sparkle: C.AMBER, dots: C.MINT_MID, rule: [9.982, 3.352, C.AMBER], page: 9 });
  arc(s, 5.175, 1.723, 1.877, C.AMBER, 270);
  corner(s, C.AMBER);
  footer(s, 5.192, 6.907);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus  ", { x: 1.425, y: 3.672, w: 3.018, h: 0.653, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  badge(s, 1.514, 2.893, 0.8, 0.586, "01", 18);
  text(s, ["Online Yoga", "Training"], { x: 2.445, y: 2.893, w: 1.793, h: 0.572, fontSize: 14, color: C.GREEN, charSpacing: 3, wrap: false });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus  ", { x: 1.425, y: 5.392, w: 3.018, h: 0.586, fontSize: 9, color: C.GRAY, lineSpacing: 1.5 });
  badge(s, 1.514, 4.613, 0.8, 0.586, "02", 18);
  text(s, ["Private Yoga", "Class"], { x: 2.445, y: 4.738, w: 2.283, h: 0.572, fontSize: 14, color: C.GREEN, charSpacing: 3 });
  text(s, ["Our Amazing", "Services"], { x: 1.443, y: 1.619, w: 2.535, h: 0.841, fontSize: 22, color: C.GREEN, charSpacing: 3, wrap: false });
  accentBar(s, 1.586, 1.522);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 8.119, y: 0.683, w: 3.556, h: 0.942, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
}

// Slide 10 — Yoga Studio Facilities
function slide10(s) {
  chrome(s, { panels: [[4.991, 0, 3.352, 7.5, C.MINT]], sparkle: C.AMBER, dots: C.AMBER, rule: [9.982, 3.352, C.AMBER], page: 10 });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum  ", { x: 9.51, y: 2.729, w: 2.569, h: 0.813, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  badge(s, 9.599, 1.95, 0.8, 0.586, "01", 18);
  text(s, ["Spa", "Room"], { x: 10.531, y: 1.95, w: 0.912, h: 0.572, fontSize: 14, color: C.GREEN, charSpacing: 3, wrap: false });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum  ", { x: 9.544, y: 4.774, w: 2.569, h: 0.813, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  badge(s, 9.634, 3.995, 0.8, 0.586, "02", 18);
  text(s, ["Yoga ", "Equipment"], { x: 10.565, y: 3.995, w: 1.585, h: 0.572, fontSize: 14, color: C.GREEN, charSpacing: 3, wrap: false });
  corner(s, C.AMBER);
  text(s, ["Yoga Studio", "Facilities"], { x: 0.727, y: 2.368, w: 2.352, h: 0.841, fontSize: 22, color: C.GREEN, charSpacing: 3, wrap: false });
  accentBar(s, 0.87, 2.272);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 0.727, y: 3.412, w: 2.784, h: 0.942, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  learnMore(s, 0.807, 4.786);
  footer(s, 5.192, 6.907, 'center');
}

// Slide 11 — Our Studio Facilities
function slide11(s) {
  chrome(s, { panels: [[0, 0, 3.745, 7.5, C.MINT]], sparkle: C.AMBER, dots: C.AMBER, rule: [9.982, 3.352, C.AMBER], page: 11 });
  arc(s, 4.013, 0.544, 1.562, C.AMBER);
  text(s, ["Our Studio", "Facilities"], { x: 6.467, y: 2.196, w: 2.181, h: 0.841, fontSize: 22, color: C.GREEN, charSpacing: 3, wrap: false });
  accentBar(s, 6.61, 2.1);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 9.329, y: 2.196, w: 2.784, h: 0.942, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  corner(s, C.AMBER);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum  ", { x: 6.467, y: 4.633, w: 2.569, h: 0.813, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  badge(s, 6.556, 3.854, 0.8, 0.586, "01", 18);
  text(s, ["Yoga", "Equipment"], { x: 7.487, y: 3.854, w: 1.585, h: 0.572, fontSize: 14, color: C.GREEN, charSpacing: 3, wrap: false });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum  ", { x: 9.544, y: 4.633, w: 2.569, h: 0.813, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  badge(s, 9.634, 3.854, 0.8, 0.586, "02", 18);
  text(s, ["Spa ", "Room"], { x: 10.565, y: 3.854, w: 0.912, h: 0.572, fontSize: 14, color: C.GREEN, charSpacing: 3, wrap: false });
  footer(s, 5.192, 6.907);
  rule(s, 0.517, 3.458, 1.691, C.BG);
  rule(s, 0.517, 4.082, 1.691, C.BG);
  text(s, "Yoga Studio", { x: 0.473, y: 3.079, w: 1.585, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  text(s, "Yoga Training", { x: 0.456, y: 3.701, w: 1.793, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
}

// Slide 12 — Our Yoga Teacher (card)
function slide12(s) {
  chrome(s, { panels: [[0, 0, 2.671, 7.5, C.MINT]], sparkle: C.AMBER, dots: C.AMBER, rule: [9.982, 3.352, C.AMBER], page: 12 });
  box(s, 7.397, 3.616, 5.919, 2.383, C.MINT);
  text(s, ["Our Yoga", "Teacher"], { x: 7.44, y: 2.152, w: 1.859, h: 0.841, fontSize: 22, color: C.GREEN, charSpacing: 3, wrap: false });
  accentBar(s, 7.583, 2.056);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum   ", { x: 10.217, y: 2.152, w: 2.343, h: 0.942, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum   ", { x: 7.795, y: 4.759, w: 2.507, h: 0.653, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  badge(s, 7.885, 4.052, 0.8, 0.586, "01", 18);
  text(s, ["Yoga", "Experience"], { x: 8.816, y: 4.052, w: 1.485, h: 0.505, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum   ", { x: 10.54, y: 4.787, w: 2.465, h: 0.653, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  badge(s, 10.629, 4.079, 0.8, 0.586, "02", 18);
  text(s, ["Health", "Knowledge"], { x: 11.561, y: 4.079, w: 1.445, h: 0.505, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  corner(s, C.AMBER);
  footer(s, 5.192, 6.907);
}

// Slide 13 — Let’s Take A Break (wide)
function slide13(s) {
  chrome(s, { panels: [[0, 0, 9.854, 7.5, C.MINT]], sparkle: C.AMBER, dots: C.AMBER, rule: [9.982, 3.352, C.AMBER], page: 13 });
  arc(s, 10.306, 0.522, 1.562, C.AMBER);
  corner(s, C.AMBER);
  text(s, ["Let’s Take ", "A Break"], { x: 1.347, y: 1.79, w: 2.77, h: 1.178, fontSize: 32, color: C.GREEN, charSpacing: 3, wrap: false });
  accentBar(s, 1.49, 1.694);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 4.367, y: 1.36, w: 2.784, h: 0.942, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  learnMore(s, 4.458, 2.725);
}

// Slide 14 — Jane May Doe
function slide14(s) {
  chrome(s, { panels: [[1.916, 1.737, 5.373, 3.75, C.MINT]], sparkle: C.AMBER, dots: C.AMBER, rule: [9.982, 3.352, C.AMBER], page: 14 });
  text(s, "Jane May Doe", { x: 8.601, y: 4.407, w: 3.966, h: 0.438, fontSize: 20, color: C.GREEN, charSpacing: 3 });
  accentBar(s, 8.705, 4.247);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 8.601, y: 4.912, w: 3.966, h: 0.942, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  corner(s, C.AMBER);
  meter(s, 8.687, 3.323, 3.69, 2.877, 0.172, C.LILAC);
  text(s, "Health Knowledge", { x: 8.601, y: 2.856, w: 2.301, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  meter(s, 8.687, 2.411, 3.69, 2.877, 0.172, C.LILAC);
  text(s, "Yoga Experiences", { x: 8.601, y: 1.944, w: 2.236, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  footer(s, 5.192, 6.907);
}

// Slide 15 — John Richard Doe
function slide15(s) {
  chrome(s, { panels: [[10.265, 1.875, 3.068, 3.029, C.MINT]], sparkle: C.AMBER, dots: C.AMBER });
  corner(s, C.AMBER);
  footer(s, 5.192, 6.907);
  meter(s, 0.78, 5.657, 3.69, 2.877, 0.172, C.PURPLE);
  text(s, "Health Knowledge", { x: 0.694, y: 5.19, w: 2.301, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  meter(s, 0.78, 4.745, 3.69, 2.877, 0.172, C.PURPLE);
  text(s, "Yoga Experiences", { x: 0.694, y: 4.278, w: 2.236, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  rule(s, 1.853, 1.717, 3.228, C.AMBER);
  rule(s, 1.853, 2.342, 3.228, C.AMBER);
  text(s, "Yoga Studio", { x: 3.604, y: 1.339, w: 1.585, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, align: 'right', wrap: false });
  text(s, "Spa Room", { x: 3.874, y: 1.961, w: 1.334, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, align: 'right', wrap: false });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet ", { x: 0.694, y: 3.442, w: 3.776, h: 0.682, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  learnMore(s, 5.946, 5.387);
  text(s, ["John", "Richard Doe"], { x: 5.582, y: 4.278, w: 2.417, h: 0.841, fontSize: 22, color: C.GREEN, charSpacing: 3, align: 'right', wrap: false });
}

// Slide 16 — Jane May Doe (split)
function slide16(s) {
  chrome(s, { panels: [[0, 0, 6.667, 7.5, C.MINT]], sparkle: C.AMBER, dots: C.AMBER, rule: [9.982, 3.352, C.AMBER], page: 16 });
  meter(s, 8.029, 2.544, 4.2, 3.275, 0.172, C.PURPLE);
  text(s, "Health Knowledge", { x: 7.943, y: 2.077, w: 2.301, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  meter(s, 8.029, 1.632, 4.2, 3.275, 0.172, C.PURPLE);
  text(s, "Yoga Experiences", { x: 7.943, y: 1.165, w: 2.236, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  corner(s, C.AMBER);
  text(s, ["Jane May", "Doe"], { x: 1.032, y: 5.116, w: 1.96, h: 0.909, fontSize: 24, color: C.GREEN, charSpacing: 3, wrap: false });
  accentBar(s, 1.175, 5.02);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 3.215, y: 5.116, w: 2.784, h: 0.942, fontSize: 8.5, color: C.GRAY, lineSpacing: 1.5 });
}

// Slide 17 — Jane May Doe (right)
function slide17(s) {
  chrome(s, { panels: [[0, 0, 2.964, 7.5, C.MINT]], sparkle: C.AMBER, dots: C.AMBER, rule: [9.982, 3.352, C.AMBER], page: 17 });
  meter(s, 8.62, 5.256, 3.69, 2.877, 0.172, C.PURPLE);
  text(s, "Health Knowledge", { x: 8.534, y: 4.789, w: 2.301, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  meter(s, 8.62, 4.344, 3.69, 2.877, 0.172, C.PURPLE);
  text(s, "Yoga Experiences", { x: 8.534, y: 3.877, w: 2.236, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  text(s, "Jane May Doe", { x: 8.534, y: 2.233, w: 3.966, h: 0.438, fontSize: 20, color: C.GREEN, charSpacing: 3 });
  accentBar(s, 8.638, 2.072);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 8.534, y: 2.738, w: 3.966, h: 0.942, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  footer(s, 5.192, 6.907);
  corner(s, C.AMBER);
}

// Slide 18 — John Richard Doe (meter)
function slide18(s) {
  chrome(s, { panels: [[0, 0, 2.554, 7.5, C.MINT], [10.369, 4.711, 2.964, 2.789, C.MINT]], sparkle: C.AMBER, dots: C.AMBER });
  arc(s, 2.844, 0.899, 1.562, C.AMBER);
  footer(s, 5.192, 6.907);
  text(s, "John Richard Doe", { x: 4.817, y: 2.058, w: 3.891, h: 0.404, fontSize: 18, color: C.GREEN, charSpacing: 3 });
  accentBar(s, 4.922, 1.923);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 4.817, y: 2.563, w: 3.738, h: 0.942, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  learnMore(s, 4.922, 5.135);
  meter(s, 4.948, 4.239, 3.4, 2.651, 0.172, C.PURPLE);
  text(s, "Yoga Experiences", { x: 4.862, y: 3.772, w: 2.236, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  corner(s, C.AMBER);
}

// Slide 19 — Our Yoga Teachers
function slide19(s) {
  chrome(s, { panels: [[0.618, 0.813, 6.048, 5.873, C.MINT]], rule: [9.982, 3.352, C.AMBER], page: 19 });
  text(s, "Our Yoga Teachers", { x: 7.822, y: 2.095, w: 3.966, h: 0.438, fontSize: 20, color: C.GREEN, charSpacing: 3 });
  accentBar(s, 7.927, 1.934);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 7.822, y: 2.6, w: 3.966, h: 0.942, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  corner(s, C.AMBER);
  meter(s, 7.919, 5.23, 4, 3.119, 0.17, C.PURPLE);
  text(s, "Yoga Experiences", { x: 7.833, y: 4.763, w: 2.236, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  footer(s, 5.192, 6.932);
  badge(s, 7.883, 3.841, 0.8, 0.586, "01", 18);
  text(s, ["John", "Doe"], { x: 8.815, y: 3.841, w: 0.784, h: 0.572, fontSize: 14, color: C.GREEN, charSpacing: 3, wrap: false });
  badge(s, 9.908, 3.824, 0.8, 0.586, "02", 18);
  text(s, ["Jane", "Doe"], { x: 10.839, y: 3.824, w: 0.763, h: 0.572, fontSize: 14, color: C.GREEN, charSpacing: 3, wrap: false });
}

// Slide 20 — The Break Slides
function slide20(s) {
  chrome(s, { panels: [[0, 0, 8.265, 7.5, C.MINT]], sparkle: C.AMBER, dots: C.AMBER, rule: [9.982, 3.352, C.AMBER], page: 20 });
  text(s, ["The Break", "Slides"], { x: 9.254, y: 2.534, w: 2.447, h: 0.909, fontSize: 24, color: C.GREEN, charSpacing: 3 });
  accentBar(s, 9.358, 2.373);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 9.254, y: 3.573, w: 3.152, h: 0.942, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  learnMore(s, 9.332, 4.685);
  corner(s, C.AMBER);
}

// Slide 21 — Our Yoga Studio
function slide21(s) {
  chrome(s, { panels: [[0, 0, 2.279, 7.5, C.MINT]], sparkle: C.AMBER, dots: C.AMBER, rule: [9.982, 3.352, C.AMBER], page: 21 });
  arc(s, 5.779, 0.905, 1.562, C.AMBER);
  text(s, "Our Yoga Studio ", { x: 8.114, y: 2.087, w: 3.891, h: 0.404, fontSize: 18, color: C.GREEN, charSpacing: 3 });
  accentBar(s, 8.219, 1.953);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 8.114, y: 2.593, w: 4.261, h: 0.942, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  badge(s, 8.189, 3.837, 0.707, 0.498, "01", 14);
  text(s, ["Clean", "Studio"], { x: 9.019, y: 3.827, w: 0.958, h: 0.505, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit  ", { x: 8.114, y: 4.454, w: 2.276, h: 0.586, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  learnMore(s, 8.204, 5.363);
  badge(s, 10.47, 3.827, 0.707, 0.498, "02", 14);
  text(s, ["Best", "Equipments"], { x: 11.3, y: 3.816, w: 1.562, h: 0.505, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit  ", { x: 10.396, y: 4.444, w: 2.276, h: 0.586, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  corner(s, C.AMBER);
  footer(s, 5.192, 6.907);
}

// Slide 22 — The Yoga Practice
function slide22(s) {
  chrome(s, { panels: [[10.21, 3.519, 3.124, 3.981, C.MINT]], sparkle: C.AMBER, dots: C.AMBER, rule: [10.21, 3.124, C.AMBER], page: 22 });
  corner(s, C.AMBER);
  footer(s, 5.192, 6.907);
  text(s, "The Yoga Practice", { x: 0.884, y: 3.094, w: 3.891, h: 0.404, fontSize: 18, color: C.GREEN, charSpacing: 3 });
  accentBar(s, 0.989, 2.959);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 0.884, y: 3.599, w: 3.338, h: 0.942, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  learnMore(s, 0.989, 4.788);
}

// Slide 23 — The Satisfied Clients
function slide23(s) {
  chrome(s, { panels: [[0, 0, 3.039, 3.039, C.MINT]], sparkle: C.AMBER, dots: C.AMBER, rule: [9.982, 3.352, C.AMBER], page: 23 });
  arc(s, 10.524, 0.803, 1.562, C.AMBER);
  corner(s, C.AMBER);
  footer(s, 5.192, 6.907);
  text(s, "The Satisfied Clients", { x: 1.551, y: 4.737, w: 3.891, h: 0.404, fontSize: 18, color: C.GREEN, charSpacing: 3 });
  accentBar(s, 1.655, 4.602);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 1.551, y: 5.242, w: 3.641, h: 0.942, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  badge(s, 5.717, 4.75, 0.707, 0.498, "01", 14);
  text(s, ["Studio", "Training"], { x: 6.546, y: 4.739, w: 1.165, h: 0.505, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit  ", { x: 5.642, y: 5.367, w: 2.276, h: 0.586, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
}

// Slide 24 — Training Program
function slide24(s) {
  chrome(s, { sparkle: C.AMBER, dots: C.AMBER, rule: [9.982, 3.352, C.AMBER], page: 24 });
  arc(s, 1.844, 0.707, 1.562, C.MINT_MID, 270);
  arc(s, 9.75, 0.704, 1.562, C.AMBER);
  corner(s, C.AMBER);
  footer(s, 5.192, 6.907);
  text(s, "Training Program", { x: 2.248, y: 5.666, w: 3.891, h: 0.404, fontSize: 18, color: C.GREEN, charSpacing: 3 });
  accentBar(s, 2.353, 5.531);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 5.43, y: 5.397, w: 3.338, h: 0.942, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  learnMore(s, 9.112, 5.647);
}

// Slide 25 — Join With Us
function slide25(s) {
  chrome(s, { panels: [[0, 0, 3.039, 3.296, C.MINT]], dots: C.AMBER, rule: [9.982, 3.352, C.AMBER], page: 25 });
  corner(s, C.AMBER);
  footer(s, 5.192, 6.907);
  text(s, "Join With Us", { x: 8.762, y: 2.003, w: 3.891, h: 0.404, fontSize: 18, color: C.GREEN, charSpacing: 3 });
  accentBar(s, 8.866, 1.868);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 8.762, y: 2.508, w: 3.641, h: 0.942, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  learnMore(s, 8.866, 3.78);
  badge(s, 7.094, 4.943, 0.707, 0.498, "01", 14);
  text(s, ["Yoga ", "Training"], { x: 7.924, y: 4.932, w: 1.165, h: 0.505, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit  ", { x: 7.019, y: 5.56, w: 2.276, h: 0.586, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  badge(s, 9.741, 4.943, 0.707, 0.498, "01", 14);
  text(s, ["Spa", "& Wellness"], { x: 10.57, y: 4.932, w: 1.48, h: 0.505, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit  ", { x: 9.666, y: 5.56, w: 2.276, h: 0.586, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
}

// Slide 26 — Join With Us (split)
function slide26(s) {
  chrome(s, { panels: [[0, 4, 2.542, 3.5, C.MINT]], dots: C.BG, rule: [9.982, 3.352, C.AMBER], page: 26 });
  text(s, "Join With Us", { x: 4.5, y: 3.34, w: 3.891, h: 0.404, fontSize: 18, color: C.GREEN, charSpacing: 3 });
  accentBar(s, 4.604, 3.205);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 4.5, y: 3.845, w: 3.641, h: 0.942, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  learnMore(s, 4.604, 5.117);
  rule(s, 5.053, 1.298, 3.228, C.AMBER);
  rule(s, 5.053, 1.923, 3.228, C.AMBER);
  text(s, "Yoga Studio", { x: 6.804, y: 0.92, w: 1.585, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, align: 'right', wrap: false });
  text(s, "Spa Room", { x: 7.074, y: 1.541, w: 1.334, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, align: 'right', wrap: false });
  badge(s, 1.027, 0.889, 0.707, 0.498, "01", 14);
  text(s, "Studio Training", { x: 1.866, y: 1.006, w: 2.006, h: 0.303, fontSize: 12, color: C.GREEN, charSpacing: 3, wrap: false });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus  ", { x: 0.952, y: 1.506, w: 3.228, h: 0.586, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  footer(s, 5.192, 6.907);
}

// Slide 27 — Our Training Program
function slide27(s) {
  chrome(s, { panels: [[0, 0, 5.5, 7.5, C.MINT]], sparkle: C.AMBER, dots: C.AMBER, rule: [9.982, 3.352, C.AMBER], page: 27 });
  text(s, "Our Training Program", { x: 0.966, y: 2.631, w: 3.891, h: 0.404, fontSize: 18, color: C.GREEN, charSpacing: 3 });
  accentBar(s, 1.071, 2.496);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 0.966, y: 3.136, w: 3.645, h: 0.942, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  learnMore(s, 1.071, 4.326);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum   ", { x: 6.331, y: 2.732, w: 2.841, h: 0.653, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  badge(s, 6.421, 2.025, 0.8, 0.586, "01", 18);
  text(s, ["Title", "Number One"], { x: 7.352, y: 2.025, w: 1.82, h: 0.572, fontSize: 14, color: C.GREEN, charSpacing: 3, wrap: false });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum   ", { x: 9.748, y: 2.732, w: 2.841, h: 0.653, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  badge(s, 9.837, 2.025, 0.8, 0.586, "02", 18);
  text(s, ["Title", "Number Two"], { x: 10.768, y: 2.025, w: 1.813, h: 0.572, fontSize: 14, color: C.GREEN, charSpacing: 3, wrap: false });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum   ", { x: 6.331, y: 4.781, w: 2.841, h: 0.653, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  badge(s, 6.421, 4.074, 0.8, 0.586, "03", 18);
  text(s, ["Title", "Number Three"], { x: 7.352, y: 4.074, w: 2.048, h: 0.572, fontSize: 14, color: C.GREEN, charSpacing: 3, wrap: false });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum   ", { x: 9.748, y: 4.781, w: 2.833, h: 0.653, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  badge(s, 9.837, 4.074, 0.8, 0.586, "04", 18);
  text(s, ["Title", "Number Four"], { x: 10.768, y: 4.074, w: 1.901, h: 0.572, fontSize: 14, color: C.GREEN, charSpacing: 3, wrap: false });
  corner(s, C.MINT_MID);
  footer(s, 6.839, 6.562);
}

// Slide 28 — Our Pricing List
function slide28(s) {
  chrome(s, { panels: [[0, 1.364, 13.333, 4.771, C.MINT]], sparkle: C.AMBER, dots: C.AMBER, rule: [9.982, 3.352, C.AMBER], page: 28 });
  pill(s, 2.074, 1.999, 2.649, 3.381, C.BG);
  pill(s, 5.343, 1.999, 2.649, 3.381, C.BG);
  pill(s, 8.612, 1.999, 2.649, 3.381, C.BG);
  corner(s, C.MINT_MID);
  cardHead(s, 5.343, 1.999, 2.649, 0.975, C.PURPLE);
  text(s, ["1 Year", "Yoga Membership"], { x: 5.343, y: 2.209, w: 2.649, h: 0.505, fontSize: 12, color: C.BG, charSpacing: 3, align: 'center' });
  cardHead(s, 2.074, 1.999, 2.649, 0.975, C.LEAF);
  text(s, ["Monthly", "Yoga Membership"], { x: 2.074, y: 2.209, w: 2.649, h: 0.505, fontSize: 12, color: C.BG, charSpacing: 3, align: 'center' });
  cardHead(s, 8.612, 1.999, 2.649, 0.975, C.AMBER);
  text(s, ["3 Months", "Yoga Membership"], { x: 8.612, y: 2.209, w: 2.649, h: 0.505, fontSize: 12, color: C.BG, charSpacing: 3, align: 'center' });
  text(s, "Our Pricing List", { x: 4.721, y: 0.549, w: 3.891, h: 0.438, fontSize: 20, color: C.GREEN, charSpacing: 3, align: 'center' });
  accentBar(s, 6.267, 0.429);
  footer(s, 5.192, 6.665);
  text(s, "Free Studio Entry", { x: 2.522, y: 3.323, w: 1.719, h: 0.346, fontSize: 10, color: C.GRAY, align: 'center', lineSpacing: 1.5 });
  text(s, "Personal Trainer", { x: 2.522, y: 3.719, w: 1.719, h: 0.346, fontSize: 10, color: C.GRAY, align: 'center', lineSpacing: 1.5 });
  text(s, "Online Training", { x: 2.522, y: 4.115, w: 1.719, h: 0.346, fontSize: 10, color: C.GRAY, align: 'center', lineSpacing: 1.5 });
  text(s, "Spa Room Access", { x: 2.522, y: 4.51, w: 1.719, h: 0.346, fontSize: 10, color: C.GRAY, align: 'center', lineSpacing: 1.5 });
  pill(s, 2.663, 5.139, 1.471, 0.513, C.MINT_MID);
  text(s, "Free Studio Entry", { x: 5.814, y: 3.323, w: 1.719, h: 0.346, fontSize: 10, color: C.GRAY, align: 'center', lineSpacing: 1.5 });
  text(s, "Personal Trainer", { x: 5.814, y: 3.719, w: 1.719, h: 0.346, fontSize: 10, color: C.GRAY, align: 'center', lineSpacing: 1.5 });
  text(s, "Online Training", { x: 5.814, y: 4.115, w: 1.719, h: 0.346, fontSize: 10, color: C.GRAY, align: 'center', lineSpacing: 1.5 });
  text(s, "Spa Room Access", { x: 5.814, y: 4.51, w: 1.719, h: 0.346, fontSize: 10, color: C.GRAY, align: 'center', lineSpacing: 1.5 });
  text(s, "Free Studio Entry", { x: 9.045, y: 3.323, w: 1.719, h: 0.346, fontSize: 10, color: C.GRAY, align: 'center', lineSpacing: 1.5 });
  text(s, "Personal Trainer", { x: 9.045, y: 3.719, w: 1.719, h: 0.346, fontSize: 10, color: C.GRAY, align: 'center', lineSpacing: 1.5 });
  text(s, "Online Training", { x: 9.045, y: 4.115, w: 1.719, h: 0.346, fontSize: 10, color: C.GRAY, align: 'center', lineSpacing: 1.5 });
  text(s, "Spa Room Access", { x: 9.045, y: 4.51, w: 1.719, h: 0.346, fontSize: 10, color: C.GRAY, align: 'center', lineSpacing: 1.5 });
  text(s, "Only $50", { x: 2.809, y: 5.271, w: 1.145, h: 0.269, fontSize: 10, color: C.GREEN, bold: true, charSpacing: 3, wrap: false });
  pill(s, 5.949, 5.139, 1.471, 0.513, C.MINT_MID);
  text(s, "Only $200", { x: 6.033, y: 5.271, w: 1.268, h: 0.269, fontSize: 10, color: C.GREEN, bold: true, charSpacing: 3, align: 'center', wrap: false });
  pill(s, 9.23, 5.139, 1.471, 0.513, C.MINT_MID);
  text(s, "Only $100", { x: 9.314, y: 5.271, w: 1.268, h: 0.269, fontSize: 10, color: C.GREEN, bold: true, charSpacing: 3, align: 'center', wrap: false });
}

// Slide 29 — Contact Our Team
function slide29(s) {
  chrome(s, { panels: [[0, 0, 5.354, 7.5, C.MINT]], sparkle: C.AMBER, dots: C.AMBER, rule: [9.982, 3.352, C.AMBER], page: 29 });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum  ", { x: 6.648, y: 4.821, w: 2.569, h: 0.813, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  text(s, "Call Us", { x: 6.648, y: 4.413, w: 2.353, h: 0.337, fontSize: 14, color: C.GREEN, charSpacing: 3 });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum  ", { x: 9.716, y: 4.821, w: 2.569, h: 0.813, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  text(s, "Visit Us", { x: 9.716, y: 4.413, w: 2.353, h: 0.337, fontSize: 14, color: C.GREEN, charSpacing: 3 });
  text(s, ["Contact Our", "Team"], { x: 1.264, y: 2.366, w: 2.407, h: 0.841, fontSize: 22, color: C.GREEN, charSpacing: 3, wrap: false });
  accentBar(s, 1.386, 2.228);
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 1.264, y: 3.316, w: 2.754, h: 1.064, fontSize: 10, color: C.GRAY, lineSpacing: 1.5 });
  learnMore(s, 1.29, 4.83);
  corner(s, C.MINT_MID);
}

// Slide 30 — Thanks For Watching
function slide30(s) {
  chrome(s, { panels: [[0, 3.75, 13.333, 3.75, C.MINT]], sparkle: C.AMBER, dots: C.AMBER, rule: [9.982, 3.352, C.AMBER], page: 30 });
  text(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eu lectus dictum  Lorem ipsum dolor sit amet, consectetur adipiscing ", { x: 4.355, y: 6.053, w: 4.638, h: 0.551, fontSize: 10, color: C.GRAY, align: 'center', lineSpacing: 1.5 });
  text(s, "Thanks For Watching", { x: 4.329, y: 0.794, w: 4.675, h: 0.555, fontSize: 27, color: C.GREEN, charSpacing: 3, align: 'center', wrap: false });
  accentBar(s, 6.274, 0.65);
  footer(s, 5.192, 6.907);
  corner(s, C.AMBER);
}

// ------------------------------------------------------------------ build ---
const SLIDES = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
  slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
  slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24,
  slide25, slide26, slide27, slide28, slide29, slide30,
];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'JOOGA', width: SLIDE_W, height: SLIDE_H });
pptx.layout = 'JOOGA';
pptx.author = 'Jooga';
pptx.title = 'Jooga Yoga Presentation';

SLIDES.forEach(function (build) { build(pptx.addSlide()); });

pptx.writeFile({ fileName: path.join(__dirname, 'bd1a84ec-be1c-4db9-b69c-95af8348426d_grok_final.pptx') })
  .then(function (f) { console.log('wrote ' + f); })
  .catch(function (e) { console.error(e); process.exit(1); });
