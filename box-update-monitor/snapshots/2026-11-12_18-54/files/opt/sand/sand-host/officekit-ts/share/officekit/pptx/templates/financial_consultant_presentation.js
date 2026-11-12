/**
 * "Monety Consultant" — Financial Consultant Presentation (30 slides, 13.333 x 7.5 in)
 * Rebuilt from scratch with pptxgenjs only. Photographs in the original are empty
 * picture placeholders / decorative vector art; they are re-drawn here as native shapes.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const NAVY = '041B58'; // theme accent1
const SKY = 'C7E8F3'; // theme accent2
const WHITE = 'FFFFFF';
const HEAD = '262626'; // tx1 @ 85% luminance — headings
const BODY = '595959'; // tx1 @ 65% luminance — body copy
const FOOT = 'D9D9D9'; // bg1 @ 85% luminance — footer url
const PHONE_SHELL = '414041';
const PHONE_EDGE = '131313';
const PHONE_BTN = '919191';

/* ------------------------------------------------------------------- fonts */
const SEMI = 'Montserrat SemiBold';
const MED = 'Montserrat Medium';
const SANS = 'Hind Madurai';

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

/* ---------------------------------------------------------- text utilities */
// Titles: 36pt Montserrat SemiBold, dark grey (white on the navy slides).
function title(s, text, x, y, w, h, o = {}) {
  s.addText(text, { x, y, w, h, fontFace: SEMI, fontSize: 36, bold: true, color: HEAD, ...o });
}
// Section / item headings: 14pt Montserrat SemiBold.
function head(s, text, x, y, w, h, o = {}) {
  s.addText(text, { x, y, w, h, fontFace: SEMI, fontSize: 14, bold: true, color: HEAD, ...o });
}
// Big numbers / KPI figures: 24pt Montserrat SemiBold.
function stat(s, text, x, y, w, h, o = {}) {
  s.addText(text, { x, y, w, h, fontFace: SEMI, fontSize: 24, bold: true, color: HEAD, ...o });
}
// Body copy: 12pt Hind Madurai, 1.5 line spacing.
function body(s, text, x, y, w, h, o = {}) {
  s.addText(text, {
    x, y, w, h, fontFace: SANS, fontSize: 12, color: BODY, lineSpacingMultiple: 1.5, ...o,
  });
}

/* --------------------------------------------------------- shape utilities */
function circle(s, cx, cy, d, color, o = {}) {
  s.addShape('ellipse', { x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { color }, line: { type: 'none' }, ...o });
}
function ring(s, cx, cy, d, color, pt) {
  s.addShape('ellipse', { x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { type: 'none' }, line: { color, width: pt } });
}
// Centre-anchored primitive: the building block for every pictogram below.
function sh(s, shape, cx, cy, w, h, color, o = {}) {
  s.addShape(shape, { x: cx - w / 2, y: cy - h / 2, w, h, fill: { color }, line: { type: 'none' }, ...o });
}
function bar(s, cx, cy, w, h, color, o = {}) {
  sh(s, 'rect', cx, cy, w, h, color, o);
}
// Polygon check-mark (the deck's small "tick" bullets).
const CHECK = [[0, 0.6], [0.38, 1], [1, 0.18], [0.86, 0.05], [0.38, 0.72], [0.13, 0.46]];
function check(s, cx, cy, size, color) {
  s.addShape('custGeom', {
    x: cx - size / 2, y: cy - size / 2, w: size, h: size,
    fill: { color }, line: { type: 'none' },
    points: CHECK.map(([px, py], i) => ({ x: px * size, y: py * size, moveTo: i === 0 })).concat([{ close: true }]),
  });
}

// Currency mark used inside several of the pictograms.
function dollar(s, cx, cy, size, color) {
  s.addText('$', {
    x: cx - size, y: cy - size / 2, w: size * 2, h: size,
    fontFace: SEMI, bold: true, fontSize: size * 68, color, align: 'center', valign: 'middle', margin: 0,
  });
}

/* -------------------------------------------------------------- pictograms */
/**
 * Stand-ins for the deck's SVG icons, each assembled from native shapes.
 * Every entry draws inside a square of side `u` centred on (cx, cy).
 */
const ICONS = {
  scales: (s, cx, cy, u, c) => {
    bar(s, cx, cy - u * 0.42, u * 1.4, u * 0.07, c);
    bar(s, cx, cy + u * 0.05, u * 0.08, u * 0.85, c);
    bar(s, cx, cy + u * 0.49, u * 0.7, u * 0.09, c);
    sh(s, 'chord', cx - u * 0.55, cy - u * 0.08, u * 0.58, u * 0.34, c, { rotate: 180 });
    sh(s, 'chord', cx + u * 0.55, cy - u * 0.08, u * 0.58, u * 0.34, c, { rotate: 180 });
  },
  puzzle: (s, cx, cy, u, c) => {
    [[-1, -1], [1, -1], [-1, 1], [1, 1]].forEach(([dx, dy]) => sh(s, 'roundRect', cx + dx * u * 0.31, cy + dy * u * 0.31, u * 0.56, u * 0.56, c, { rectRadius: 0.1 }));
    [[-1, 0], [1, 0], [0, -1], [0, 1]].forEach(([dx, dy]) => sh(s, 'ellipse', cx + dx * u * 0.62, cy + dy * u * 0.62, u * 0.26, u * 0.26, c));
    [[-1, 0], [1, 0], [0, -1], [0, 1]].forEach(([dx, dy]) => sh(s, 'ellipse', cx + dx * u * 0.16, cy + dy * u * 0.16, u * 0.22, u * 0.22, c));
  },
  // Two forearms meeting in a clasp, with thin gaps standing in for fingers.
  handshake: (s, cx, cy, u, c, bg) => {
    sh(s, 'roundRect', cx - u * 0.46, cy + u * 0.18, u * 0.8, u * 0.3, c, { rotate: -18, rectRadius: 0.35 });
    sh(s, 'roundRect', cx + u * 0.46, cy - u * 0.16, u * 0.8, u * 0.3, c, { rotate: -18, rectRadius: 0.35 });
    sh(s, 'roundRect', cx, cy + u * 0.02, u * 0.66, u * 0.56, c, { rotate: -18, rectRadius: 0.3 });
    [-0.16, 0, 0.16].forEach((d) => sh(s, 'rect', cx + d * u * 0.9, cy + d * u * 0.3 + u * 0.04, u * 0.5, u * 0.04, bg, { rotate: -18 }));
    sh(s, 'triangle', cx + u * 0.44, cy - u * 0.42, u * 0.36, u * 0.3, c, { rotate: 110 });
  },
  gearGlobe: (s, cx, cy, u, c, bg) => {
    sh(s, 'gear6', cx, cy, u * 1.05, u * 1.05, c);
    circle(s, cx, cy, u * 0.62, c);
    circle(s, cx, cy, u * 0.44, bg);
    circle(s, cx, cy, u * 0.2, c);
  },
  donutBars: (s, cx, cy, u, c, bg) => {
    circle(s, cx, cy, u, c);
    circle(s, cx, cy, u * 0.76, bg);
    bar(s, cx, cy - u * 0.5, u * 0.13, u * 0.26, bg);
    bar(s, cx + u * 0.36, cy + u * 0.36, u * 0.26, u * 0.13, bg);
    bar(s, cx - u * 0.14, cy + u * 0.06, u * 0.11, u * 0.34, c);
    bar(s, cx + u * 0.05, cy, u * 0.11, u * 0.46, c);
  },
  coinCycle: (s, cx, cy, u, c, bg) => {
    circle(s, cx, cy, u * 0.98, c);
    circle(s, cx, cy, u * 0.82, bg);
    bar(s, cx - u * 0.34, cy - u * 0.34, u * 0.5, u * 0.3, bg, { rotate: 45 });
    bar(s, cx + u * 0.34, cy + u * 0.34, u * 0.5, u * 0.3, bg, { rotate: 45 });
    sh(s, 'triangle', cx + u * 0.44, cy - u * 0.34, u * 0.28, u * 0.28, c, { rotate: 135 });
    sh(s, 'triangle', cx - u * 0.44, cy + u * 0.34, u * 0.28, u * 0.28, c, { rotate: -45 });
    circle(s, cx, cy, u * 0.5, c);
    dollar(s, cx, cy, u * 0.42, bg);
  },
  coins: (s, cx, cy, u, c, bg) => {
    circle(s, cx - u * 0.28, cy - u * 0.3, u * 0.52, c);
    circle(s, cx + u * 0.3, cy - u * 0.24, u * 0.48, c);
    circle(s, cx, cy + u * 0.22, u * 0.72, c);
    dollar(s, cx, cy + u * 0.22, u * 0.55, bg);
  },
  barsUp: (s, cx, cy, u, c, bg) => {
    [[-0.6, 0.28], [-0.32, 0.48], [-0.04, 0.68]].forEach(([dx, h]) => bar(s, cx + dx * u, cy + u * (0.42 - h / 2), u * 0.18, u * h, c));
    sh(s, 'upArrow', cx + u * 0.24, cy - u * 0.14, u * 0.4, u * 0.78, c);
    circle(s, cx + u * 0.58, cy + u * 0.32, u * 0.5, c);
    dollar(s, cx + u * 0.58, cy + u * 0.32, u * 0.36, bg);
  },
  // Disc with a quarter cut away and set slightly apart, like an exploded pie chart.
  pieSlice: (s, cx, cy, u, c, bg) => {
    circle(s, cx + u * 0.06, cy - u * 0.06, u, c);
    bar(s, cx - u * 0.24, cy + u * 0.24, u * 0.52, u * 0.52, bg);
    sh(s, 'pieWedge', cx - u * 0.24, cy + u * 0.24, u * 0.44, u * 0.44, c, { rotate: 180 });
    sh(s, 'pieWedge', cx - u * 0.24, cy + u * 0.26, u * 0.3, u * 0.3, bg, { rotate: 180 });
  },
  target: (s, cx, cy, u, c, bg) => {
    circle(s, cx - u * 0.06, cy + u * 0.06, u * 0.95, c);
    circle(s, cx - u * 0.06, cy + u * 0.06, u * 0.72, bg);
    circle(s, cx - u * 0.06, cy + u * 0.06, u * 0.48, c);
    circle(s, cx - u * 0.06, cy + u * 0.06, u * 0.24, bg);
    sh(s, 'rect', cx + u * 0.16, cy - u * 0.16, u * 0.92, u * 0.14, bg, { rotate: -45 });
    sh(s, 'rect', cx + u * 0.18, cy - u * 0.18, u * 0.78, u * 0.08, c, { rotate: -45 });
    sh(s, 'triangle', cx + u * 0.56, cy - u * 0.56, u * 0.34, u * 0.3, c, { rotate: 45 });
  },
  bulb: (s, cx, cy, u, c) => {
    circle(s, cx, cy - u * 0.16, u * 0.6, c);
    bar(s, cx, cy + u * 0.28, u * 0.24, u * 0.18, c);
    bar(s, cx, cy + u * 0.44, u * 0.16, u * 0.08, c);
  },
  safe: (s, cx, cy, u, c, bg) => {
    sh(s, 'frame', cx, cy, u, u * 0.86, c);
    circle(s, cx - u * 0.12, cy, u * 0.4, c);
    circle(s, cx - u * 0.12, cy, u * 0.22, bg);
    bar(s, cx + u * 0.24, cy - u * 0.1, u * 0.22, u * 0.06, c);
    bar(s, cx + u * 0.24, cy + u * 0.1, u * 0.22, u * 0.06, c);
    bar(s, cx - u * 0.32, cy + u * 0.51, u * 0.08, u * 0.14, c);
    bar(s, cx + u * 0.32, cy + u * 0.51, u * 0.08, u * 0.14, c);
  },
  layers: (s, cx, cy, u, c) => [-0.28, 0, 0.28].forEach((d) => sh(s, 'diamond', cx, cy + d * u, u, u * 0.4, c)),
  handCoin: (s, cx, cy, u, c, bg) => {
    circle(s, cx + u * 0.24, cy - u * 0.26, u * 0.56, c);
    dollar(s, cx + u * 0.24, cy - u * 0.26, u * 0.42, bg);
    sh(s, 'roundRect', cx + u * 0.1, cy + u * 0.34, u * 0.82, u * 0.24, c, { rectRadius: 0.5 });
    sh(s, 'roundRect', cx + u * 0.46, cy + u * 0.2, u * 0.26, u * 0.2, c, { rotate: -35, rectRadius: 0.5 });
    bar(s, cx - u * 0.5, cy + u * 0.34, u * 0.32, u * 0.34, c);
  },
  personChart: (s, cx, cy, u, c) => {
    circle(s, cx - u * 0.3, cy - u * 0.3, u * 0.44, c);
    sh(s, 'blockArc', cx - u * 0.3, cy + u * 0.34, u * 0.86, u * 0.86, c);
    bar(s, cx + u * 0.34, cy + u * 0.24, u * 0.13, u * 0.36, c);
    bar(s, cx + u * 0.58, cy + u * 0.15, u * 0.13, u * 0.54, c);
  },
  headIdea: (s, cx, cy, u, c, bg) => {
    sh(s, 'flowChartDelay', cx - u * 0.2, cy + u * 0.08, u * 0.9, u * 0.9, c, { rotate: 270 });
    circle(s, cx + u * 0.36, cy - u * 0.3, u * 0.56, c);
    bar(s, cx + u * 0.36, cy - u * 0.34, u * 0.07, u * 0.2, bg);
    circle(s, cx + u * 0.36, cy - u * 0.16, u * 0.07, bg);
  },
  banknote: (s, cx, cy, u, c) => {
    sh(s, 'frame', cx + u * 0.06, cy - u * 0.12, u * 0.96, u * 0.62, c);
    dollar(s, cx + u * 0.06, cy - u * 0.12, u * 0.3, c);
    bar(s, cx - u * 0.44, cy + u * 0.24, u * 0.06, u * 0.28, c);
    bar(s, cx + u * 0.06, cy + u * 0.35, u * 0.96, u * 0.06, c);
  },
  chat: (s, cx, cy, u, c, bg) => {
    sh(s, 'wedgeRoundRectCallout', cx + u * 0.14, cy + u * 0.16, u * 0.8, u * 0.6, c);
    sh(s, 'wedgeRoundRectCallout', cx - u * 0.14, cy - u * 0.16, u * 0.8, u * 0.6, c, { flipH: true });
    dollar(s, cx - u * 0.16, cy - u * 0.22, u * 0.3, bg);
  },
  handset: (s, cx, cy, u, c) => {
    sh(s, 'moon', cx - u * 0.1, cy + u * 0.06, u * 0.8, u * 0.8, c, { rotate: 225 });
    sh(s, 'blockArc', cx + u * 0.24, cy - u * 0.2, u * 0.7, u * 0.7, c, { rotate: 135 });
  },
  gearRing: (s, cx, cy, u, c, bg) => {
    circle(s, cx, cy, u, c);
    circle(s, cx, cy, u * 0.78, bg);
    bar(s, cx + u * 0.36, cy - u * 0.36, u * 0.3, u * 0.16, bg, { rotate: -45 });
    bar(s, cx - u * 0.36, cy + u * 0.36, u * 0.3, u * 0.16, bg, { rotate: -45 });
    sh(s, 'gear6', cx, cy, u * 0.62, u * 0.62, c);
    circle(s, cx, cy, u * 0.2, bg);
  },
  facebook: (s, cx, cy, u, c) => s.addText('f', { x: cx - u * 0.5, y: cy - u * 0.62, w: u, h: u * 1.24, fontFace: SEMI, bold: true, fontSize: u * 80, color: c, align: 'center', valign: 'middle', margin: 0 }),
  twitter: (s, cx, cy, u, c) => {
    sh(s, 'moon', cx, cy + u * 0.06, u * 0.8, u * 0.8, c, { rotate: 200 });
    sh(s, 'triangle', cx + u * 0.18, cy - u * 0.14, u * 0.44, u * 0.32, c, { rotate: 135 });
  },
  instagram: (s, cx, cy, u, c) => {
    s.addShape('roundRect', { x: cx - u * 0.45, y: cy - u * 0.45, w: u * 0.9, h: u * 0.9, fill: { type: 'none' }, line: { color: c, width: 1.2 }, rectRadius: 0.3 });
    s.addShape('ellipse', { x: cx - u * 0.22, y: cy - u * 0.22, w: u * 0.44, h: u * 0.44, fill: { type: 'none' }, line: { color: c, width: 1.2 } });
    circle(s, cx + u * 0.26, cy - u * 0.26, u * 0.1, c);
  },
};
function icon(s, name, cx, cy, u, color, bg) {
  ICONS[name](s, cx, cy, u, color, bg || WHITE);
}

/* -------------------------------------------------- recurring page furniture */
// Top-left brand lockup: three rings + "Monety Consultant" wordmark.
const LOGO_RINGS = [[0.580, 0.554], [0.761, 0.608], [0.687, 0.749]];
function brand(s, onDark) {
  const ringColor = onDark ? WHITE : NAVY;
  LOGO_RINGS.forEach(([cx, cy]) => ring(s, cx, cy, 0.099, ringColor, 2));
  s.addText(
    [
      { text: 'Monety', options: { fontFace: SEMI, bold: true, color: onDark ? WHITE : NAVY } },
      { text: ' Consultant ', options: { fontFace: MED, color: onDark ? WHITE : SKY } },
    ],
    { x: 0.896, y: 0.421, w: 1.432, h: 0.505, fontSize: 12, charSpacing: 2 }
  );
}
function footer(s) {
  s.addText('www.yourgreatsite.com', {
    x: 9.753, y: 6.856, w: 3.045, h: 0.303,
    fontFace: SANS, fontSize: 12, color: FOOT, charSpacing: 3, align: 'right',
  });
}
function page(pptx, onDark) {
  const s = pptx.addSlide();
  brand(s, !!onDark);
  footer(s);
  return s;
}

/* ------------------------------------------------- reusable composite parts */
// Numbered bullet: filled disc with its ordinal, heading and paragraph beside it.
function numberedItem(s, o) {
  circle(s, o.cx, o.cy, o.d, o.disc);
  s.addText(o.num, {
    x: o.cx - o.numW / 2, y: o.cy - 0.168, w: o.numW, h: 0.337,
    fontFace: SEMI, fontSize: 14, bold: true, color: o.numColor, align: 'center',
  });
  head(s, o.title, o.tx, o.ty, o.tw, 0.337);
  body(s, o.body, o.bx, o.by, o.bw, o.bh);
}
// Circular "photo" medallion used throughout the deck (image → flat colour disc).
function medallion(s, cx, cy, d, color) {
  circle(s, cx, cy, d, color);
}

/* =================================================================== SLIDES */

function slide01(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: { color: NAVY, transparency: 15 }, line: { type: 'none' } });
  s.addShape('ellipse', { x: 3.079, y: 0.163, w: 7.175, h: 7.175, fill: { color: NAVY, transparency: 20 }, line: { type: 'none' } });
  brand(s, true);
  s.addText('Monety', {
    x: 3.524, y: 2.892, w: 6.286, h: 1.717,
    fontFace: SEMI, fontSize: 96, bold: true, color: WHITE, charSpacing: 2, align: 'center',
  });
  s.addText('Financial Consultant Presentation', {
    x: 4.619, y: 5.673, w: 4.095, h: 0.728,
    fontFace: SEMI, fontSize: 18, bold: true, color: WHITE, charSpacing: 3, align: 'center', lineSpacing: 23,
  });
}

function slide02(pptx) {
  const s = page(pptx);
  title(s, 'Perfect Financial Advisor for You', 1.104, 1.622, 5.804, 1.313);
  const LOREM = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut.';
  numberedItem(s, {
    cx: 1.551, cy: 4.097, d: 0.625, disc: NAVY, num: '01', numW: 0.534, numColor: WHITE,
    title: 'Professional Consultant', tx: 2.166, ty: 3.598, tw: 3.261,
    body: LOREM, bx: 2.166, by: 3.916, bw: 3.931, bh: 0.682,
  });
  numberedItem(s, {
    cx: 1.551, cy: 5.584, d: 0.625, disc: SKY, num: '02', numW: 0.534, numColor: HEAD,
    title: 'Best Panning', tx: 2.166, ty: 5.084, tw: 1.867,
    body: LOREM, bx: 2.166, by: 5.402, bw: 3.931, bh: 0.682,
  });
  medallion(s, 11.225, 2.5, 2.01, NAVY);
  icon(s, 'scales', 11.225, 2.14, 0.44, WHITE, NAVY);
  body(s, 'Lorem ipsum dolor adipiscing elit', 10.31, 2.439, 1.831, 0.682, { color: WHITE, align: 'center' });
}

function slide03(pptx) {
  const s = page(pptx);
  title(s, 'Meet Our Top Consultants', 2.828, 1.198, 7.677, 0.707, { align: 'center' });
  const PEOPLE = [
    { name: 'Mora Jonah', cx: 2.184, cy: 3.254, disc: NAVY, ny: 4.417, by: 4.825, sy: 5.707 },
    { name: 'Fredy Theo', cx: 5.249, cy: 3.635, disc: SKY, ny: 4.798, by: 5.206, sy: 6.088 },
    { name: 'Romaria Cloe', cx: 8.162, cy: 3.635, disc: NAVY, ny: 4.798, by: 5.206, sy: 6.088 },
    { name: 'Mike Morgan', cx: 11.151, cy: 3.254, disc: SKY, ny: 4.417, by: 4.830, sy: 5.707 },
  ];
  PEOPLE.forEach((p) => {
    medallion(s, p.cx, p.cy, 1.989, p.disc);
    head(s, p.name, p.cx - 0.979, p.ny, 1.958, 0.337, { align: 'center' });
    body(s, 'Lorem ipsum dolor sit amet, adipiscing elit, sed do.', p.cx - 1.158, p.by, 2.316, 0.682, { align: 'center' });
    ['facebook', 'twitter', 'instagram'].forEach((g, i) => {
      const bx = p.cx - 0.475 + i * 0.475;
      s.addShape('roundRect', { x: bx - 0.102, y: p.sy, w: 0.204, h: 0.233, rectRadius: 0.18, fill: { color: NAVY }, line: { type: 'none' } });
      icon(s, g, bx, p.sy + 0.117, 0.13, WHITE, NAVY);
    });
  });
}

function slide04(pptx) {
  const s = page(pptx);
  title(s, 'The Strategy to Reach Your Goals', 1.358, 1.767, 5.309, 1.313);
  medallion(s, 5.49, 4.921, 2.353, NAVY);
  icon(s, 'puzzle', 5.49, 4.504, 0.48, WHITE, NAVY);
  body(s, 'Lorem ipsum dolor sit amet adipiscing elit', 4.475, 4.817, 2.032, 0.682, { color: WHITE, align: 'center' });
  const LOREM = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ';
  [
    { dot: NAVY, dy: 2.126, hy: 2.020, by: 2.501, t: 'Reach All Target' },
    { dot: SKY, dy: 4.400, hy: 4.294, by: 4.775, t: 'Try New Things' },
  ].forEach((r) => {
    s.addShape('ellipse', { x: 7.654, y: r.dy, w: 0.124, h: 0.124, fill: { color: r.dot }, line: { type: 'none' } });
    head(s, r.t, 8.067, r.hy, 4.319, 0.337);
    body(s, LOREM, 8.067, r.by, 3.772, 0.985);
  });
}

function slide05(pptx) {
  const s = page(pptx);
  s.addShape('ellipse', { x: 5.018, y: 1.263, w: 4.974, h: 4.974, fill: { color: NAVY, transparency: 15 }, line: { type: 'none' } });
  title(s, 'For A Secure Planned Asset', 0.586, 2.636, 5.475, 1.313);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod incididunt ut enim veniam.', 0.586, 4.182, 4.326, 0.682);
  const COPY = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor et dolore magna aliqua. ';
  [
    { n: '90,0%', nx: 6.679, nw: 1.651, ny: 2.008, by: 2.597 },
    { n: '1500+', nx: 6.775, nw: 1.459, ny: 3.919, by: 4.508 },
  ].forEach((k) => {
    stat(s, k.n, k.nx, k.ny, k.nw, 0.505, { color: WHITE, align: 'center' });
    body(s, COPY, 5.796, k.by, 3.417, 0.985, { color: WHITE, align: 'center' });
  });
}

function slide06(pptx) {
  const s = page(pptx);
  title(s, 'Ready to Helping You', 3.093, 1.146, 7.148, 0.707, { align: 'center' });
  medallion(s, 3.77, 3.884, 0.89, NAVY);
  icon(s, 'handshake', 3.77, 3.884, 0.36, WHITE, NAVY);
  medallion(s, 9.563, 3.884, 0.89, SKY);
  icon(s, 'gearGlobe', 9.563, 3.884, 0.34, HEAD, SKY);
  const COPY = 'PLACEHOLDER';
  head(s, 'Management', 0.896, 3.75, 1.719, 0.337);
  body(s, COPY, 0.896, 4.232, 2.197, 1.287);
  head(s, 'Development', 10.779, 3.75, 1.719, 0.337, { align: 'right' });
  body(s, COPY, 10.301, 4.232, 2.197, 1.287, { align: 'right' });
}

// Rounded-rectangle outline as a single closed sub-path (outer CW, inner CCW),
// so the middle stays transparent instead of covering what sits behind it.
const BEZ = 0.5523;
function roundRectPath(x, y, w, h, r, clockwise) {
  const k = r * BEZ;
  return clockwise
    ? [
      { x: x + r, y, moveTo: true }, { x: x + w - r, y },
      { x: x + w, y: y + r, curve: { type: 'cubic', x1: x + w - k, y1: y, x2: x + w, y2: y + k } },
      { x: x + w, y: y + h - r },
      { x: x + w - r, y: y + h, curve: { type: 'cubic', x1: x + w, y1: y + h - k, x2: x + w - k, y2: y + h } },
      { x: x + r, y: y + h },
      { x, y: y + h - r, curve: { type: 'cubic', x1: x + k, y1: y + h, x2: x, y2: y + h - k } },
      { x, y: y + r },
      { x: x + r, y, curve: { type: 'cubic', x1: x, y1: y + k, x2: x + k, y2: y } },
      { close: true },
    ]
    : [
      { x: x + r, y, moveTo: true },
      { x, y: y + r, curve: { type: 'cubic', x1: x + k, y1: y, x2: x, y2: y + k } },
      { x, y: y + h - r },
      { x: x + r, y: y + h, curve: { type: 'cubic', x1: x, y1: y + h - k, x2: x + k, y2: y + h } },
      { x: x + w - r, y: y + h },
      { x: x + w, y: y + h - r, curve: { type: 'cubic', x1: x + w - k, y1: y + h, x2: x + w, y2: y + h - k } },
      { x: x + w, y: y + r },
      { x: x + w - r, y, curve: { type: 'cubic', x1: x + w, y1: y + k, x2: x + w - k, y2: y } },
      { close: true },
    ];
}
// Hollow rounded frame: outer ring only, the interior shows the slide through.
function hollowRoundRect(s, x, y, w, h, r, t, color) {
  s.addShape('custGeom', {
    x, y, w, h, fill: { color }, line: { type: 'none' },
    points: roundRectPath(0, 0, w, h, r, true).concat(roundRectPath(t, t, w - 2 * t, h - 2 * t, Math.max(r - t, 0.02), false)),
  });
}
// Vector phone mock-up standing in for slide 7's device artwork.
function phoneMockup(s, x, y, w, h) {
  hollowRoundRect(s, x, y, w, h, 0.3, 0.034, PHONE_SHELL);
  hollowRoundRect(s, x + 0.034, y + 0.036, w - 0.068, h - 0.072, 0.27, 0.107, PHONE_EDGE);
  s.addShape('roundRect', { x: x + w / 2 - 0.42, y: y + 0.036, w: 0.84, h: 0.2, rectRadius: 0.4, fill: { color: PHONE_EDGE }, line: { type: 'none' } });
  bar(s, x + w / 2 - 0.06, y + 0.12, 0.29, 0.035, PHONE_BTN);
  circle(s, x + w / 2 + 0.2, y + 0.12, 0.05, PHONE_BTN);
  s.addShape('roundRect', { x: x - 0.02, y: y + 0.65, w: 0.037, h: 1.107, rectRadius: 0.45, fill: { color: PHONE_BTN }, line: { type: 'none' } });
  s.addShape('roundRect', { x: x + w - 0.017, y: y + 1.11, w: 0.038, h: 0.528, rectRadius: 0.45, fill: { color: PHONE_BTN }, line: { type: 'none' } });
}

function slide07(pptx) {
  const s = page(pptx);
  medallion(s, 4.947, 3.963, 4.421, NAVY);
  phoneMockup(s, 1.425, 1.591, 2.341, 4.744);
  title(s, 'Balancing All of Your Assets', 7.529, 1.886, 4.945, 1.313);
  head(s, 'Manage Everything Well', 7.57, 3.719, 4.877, 0.337);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam.', 7.57, 4.201, 4.304, 0.985);
  body(s, 'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.', 7.58, 5.355, 4.228, 0.682);
  [
    { n: '19,8K', ny: 2.615, by: 3.083 },
    { n: '1370+', ny: 4.129, by: 4.597 },
  ].forEach((k) => {
    stat(s, k.n, 4.136, k.ny, 1.867, 0.505, { color: WHITE });
    body(s, ['Lorem ipsum dolor sit amet, ', 'adipiscing elit sed do. '].join('\n'), 4.136, k.by, 2.339, 0.682, { color: WHITE });
  });
}

function slide08(pptx) {
  const s = page(pptx);
  title(s, 'It\u2019s All About Your Bright Future', 0.896, 1.831, 5.804, 1.313);
  head(s, 'Invest Your Stock', 0.954, 3.713, 3.956, 0.337);
  stat(s, '2022', 0.954, 4.417, 1.172, 0.505, { color: NAVY });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod.', 2.126, 4.299, 3.272, 0.682);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam.', 0.954, 5.23, 4.304, 0.985);
  [
    { cx: 7.964, disc: NAVY, glyph: 'donutBars', gc: WHITE, t: 'Aspect 01', hx: 7.03, bx: 6.839 },
    { cx: 11.083, disc: SKY, glyph: 'coinCycle', gc: HEAD, t: 'Aspect 02', hx: 10.148, bx: 9.957 },
  ].forEach((c) => {
    medallion(s, c.cx, 4.119, 0.837, c.disc);
    icon(s, c.glyph, c.cx, 4.119, 0.3, c.gc, c.disc);
    head(s, c.t, c.hx, 4.843, 1.867, 0.337, { align: 'center' });
    body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod.', c.bx, 5.23, 2.25, 0.985, { align: 'center' });
  });
}

function slide09(pptx) {
  const s = page(pptx);
  s.addShape('ellipse', { x: 4.21, y: 4.098, w: 2.352, h: 2.352, fill: { color: WHITE }, line: { type: 'none' } });
  medallion(s, 1.634, 2.26, 0.9, NAVY);
  icon(s, 'coins', 1.634, 2.26, 0.4, WHITE, NAVY);
  title(s, 'Ensuring Your Financial Freedom', 6.985, 1.77, 5.804, 1.313);
  [
    { cx: 7.274, disc: NAVY, tick: WHITE, t: 'Prosperous', tx: 7.558 },
    { cx: 9.717, disc: SKY, tick: NAVY, t: 'Success', tx: 9.997 },
  ].forEach((c) => {
    circle(s, c.cx, 4.054, 0.275, c.disc);
    check(s, c.cx, 4.054, 0.154, c.tick);
    head(s, c.t, c.tx, 3.885, 1.808, 0.337);
  });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris.', 7.027, 4.557, 5.381, 0.985);
}

function slide10(pptx) {
  const s = page(pptx);
  title(s, 'Giving You Financial Power', 1.04, 2.079, 5.382, 1.313);
  const CARDS = [
    { cx: 2.091, disc: NAVY, fg: WHITE, tc: WHITE, bc: WHITE, glyph: 'barsUp', t: 'Value 01', hx: 1.158, bx: 1.083 },
    { cx: 4.778, disc: SKY, fg: HEAD, tc: HEAD, bc: BODY, glyph: 'pieSlice', t: 'Value 02', hx: 3.844, bx: 3.769 },
    { cx: 7.464, disc: NAVY, fg: WHITE, tc: WHITE, bc: WHITE, glyph: 'target', t: 'Value 03', hx: 6.53, bx: 6.455 },
  ];
  CARDS.forEach((c) => {
    medallion(s, c.cx, 5.228, 2.572, c.disc);
    icon(s, c.glyph, c.cx, 4.6, 0.4, c.fg, c.disc);
    head(s, c.t, c.hx, 5.056, 1.867, 0.337, { align: 'center', color: c.tc });
    body(s, 'Lorem ipsum dolor sit amet adipiscing.', c.bx, 5.412, 2.017, 0.682, { align: 'center', color: c.bc });
  });
}

function slide11(pptx) {
  const s = pptx.addSlide();
  // Full-width navy banner whose lower edge sweeps in a shallow arc.
  s.addShape('custGeom', {
    x: 0, y: 0, w: SLIDE_W, h: 4.528,
    fill: { color: NAVY, transparency: 15 }, line: { type: 'none' },
    points: [
      { x: 0, y: 0, moveTo: true }, { x: 13.3333, y: 0 }, { x: 13.3333, y: 3.4954 }, { x: 13.1223, y: 3.5412 },
      { x: 6.6667, y: 4.5277, curve: { type: 'cubic', x1: 11.1308, y1: 4.1823, x2: 8.9403, y2: 4.5277 } },
      { x: 0.1410, y: 3.5412, curve: { type: 'cubic', x1: 4.3931, y1: 4.5277, x2: 2.2026, y2: 4.1823 } },
      { x: 0, y: 3.4954 }, { close: true },
    ],
  });
  brand(s, true);
  footer(s);
  title(s, 'Let Your Money Grow', 2.828, 1.095, 7.677, 0.707, { align: 'center', color: WHITE });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.', 1.992, 2.303, 9.348, 0.985, { color: WHITE, align: 'center' });
  const COPY = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut.';
  [
    { dx: 1.243, disc: NAVY, num: '01.', nx: 1.175, nc: WHITE, bold: false, tx: 1.877, t: 'Strategy And Advisiory' },
    { dx: 7.45, disc: SKY, num: '02.', nx: 7.382, nc: HEAD, bold: true, tx: 8.078, t: 'Great Consultation' },
  ].forEach((r) => {
    s.addShape('ellipse', { x: r.dx, y: 5.066, w: 0.429, h: 0.429, fill: { color: r.disc }, line: { type: 'none' } });
    s.addText(r.num, { x: r.nx, y: 5.129, w: 0.564, h: 0.303, fontFace: SEMI, fontSize: 12, bold: r.bold, color: r.nc, align: 'center' });
    head(s, r.t, r.tx, 5.115, 4.279, 0.337);
    body(s, COPY, r.tx, 5.539, 4.072, 0.682);
  });
}

function slide12(pptx) {
  const s = page(pptx);
  title(s, 'We Laundry Your Money so Well', 6.72, 2.081, 5.804, 1.313);
  medallion(s, 1.43, 3.157, 0.837, NAVY);
  icon(s, 'bulb', 1.43, 3.157, 0.34, WHITE, NAVY);
  medallion(s, 1.43, 5.522, 0.837, SKY);
  icon(s, 'safe', 1.43, 5.522, 0.31, HEAD, SKY);
  circle(s, 7.119, 4.599, 0.515, NAVY);
  check(s, 7.119, 4.599, 0.27, WHITE);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam.', 7.634, 4.106, 4.237, 0.985);
  body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit sed eiusmod tempor incididunt ut labore et dolore magna aliqua. ', 6.798, 5.312, 5.462, 0.682);
}

function slide13(pptx) {
  const s = page(pptx);
  title(s, 'Keep Your Money Save', 0.7, 1.996, 4.341, 1.313);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam.', 0.7, 4.001, 3.394, 1.287);
  // Quarter-disc block (square with one rounded corner), rotated a quarter turn.
  const q = 2.56;
  s.addShape('custGeom', {
    x: 6.667, y: 3.75, w: q, h: q, rotate: 90, fill: { color: NAVY }, line: { type: 'none' },
    points: [
      { x: 0, y: 0, moveTo: true },
      { x: 0.9948 * q, y: 0.8978 * q, curve: { type: 'cubic', x1: 0.5177 * q, y1: 0, x2: 0.9436 * q, y2: 0.3935 * q } },
      { x: q, y: q }, { x: 0, y: q }, { close: true },
    ],
  });
  stat(s, '75%', 6.94, 4.147, 1.172, 0.505, { color: WHITE });
  body(s, 'Lorem ipsum dolor sit amet elit sed.', 6.94, 4.728, 1.741, 0.682, { color: WHITE });
  [
    { n: '01.', ny: 1.831, t: 'Trusted', ty: 2.382, by: 2.765 },
    { n: '02.', ny: 4.053, t: 'Guaranteed', ty: 4.605, by: 4.988 },
  ].forEach((r) => {
    stat(s, r.n, 9.753, r.ny, 1.172, 0.505, { color: NAVY });
    head(s, r.t, 9.753, r.ty, 2.651, 0.337);
    body(s, 'Lorem ipsum dolor sit amet adipiscing eiusmod.', 9.753, r.by, 2.411, 0.682);
  });
}

function slide14(pptx) {
  const s = page(pptx);
  title(s, 'Giving You Idea for Better Return', 1.37, 1.841, 5.063, 1.313);
  icon(s, 'headIdea', 1.612, 3.951, 0.33, NAVY);
  head(s, 'Create Best Strategy', 2.03, 3.782, 4.877, 0.337);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam.', 1.37, 4.424, 4.136, 0.985);
}

function slide15(pptx) {
  const s = page(pptx);
  title(s, 'Better Return in Investment', 8.683, 1.862, 4.283, 1.313);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua enim ad minim.', 8.683, 3.563, 3.893, 0.985);
  [
    { y: 4.956, t: 'Digital Wallet' },
    { y: 5.461, t: 'Choose The Right' },
  ].forEach((r) => {
    check(s, 8.891, r.y + 0.097, 0.193, NAVY);
    head(s, r.t, 9.213, r.y - 0.072, 3.18, 0.337);
  });
}

function slide16(pptx) {
  const s = page(pptx);
  title(s, 'The Secret of Great Laundering', 1.082, 1.945, 5.042, 1.313);
  stat(s, '500+', 1.082, 3.772, 1.172, 0.505, { color: NAVY });
  head(s, 'Happy Client', 1.082, 4.324, 2.651, 0.337);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation.', 1.082, 4.807, 4.777, 0.985);
  medallion(s, 8.719, 2.725, 3.329, NAVY);
  stat(s, '90,5%', 7.685, 1.688, 2.067, 0.505, { color: WHITE, align: 'center' });
  head(s, 'Project Success', 7.207, 2.316, 3.024, 0.337, { color: WHITE, align: 'center' });
  body(s, 'Lorem ipsum dolor sit amet elit sed do eiusmod tempor incididunt ut dolore magna aliqua.', 7.317, 2.777, 2.804, 0.985, { color: WHITE, align: 'center' });
}

function slide17(pptx) {
  const s = page(pptx);
  title(s, 'Financial Freedom for Your Life', 6.85, 1.893, 5.804, 1.313);
  const COPY = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud.';
  [4.045, 5.387].forEach((y) => {
    ring(s, 7.128, y + 0.128, 0.256, NAVY, 1.5);
    body(s, COPY, 7.555, y - 0.364, 4.555, 0.985);
  });
}

function slide18(pptx) {
  const s = page(pptx);
  title(s, 'Financial Dream Service Provider', 1.611, 1.622, 5.804, 1.313);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim quis nostrud exercitation ullamco.', 6.96, 1.786, 4.763, 0.985);
  const COPY = 'Lorem ipsum dolor sit amet elit, sed do eiusmod tempor labore et dolore magna aliqua. ';
  [
    { cx: 3.992, disc: NAVY, glyph: 'layers', gc: WHITE, tx: 0.769, t: 'Best Service' },
    { cx: 9.341, disc: SKY, glyph: 'handCoin', gc: HEAD, tx: 10.021, t: 'Providing Profit' },
  ].forEach((c) => {
    medallion(s, c.cx, 5.156, 0.837, c.disc);
    icon(s, c.glyph, c.cx, 5.156, 0.32, c.gc, c.disc);
    head(s, c.t, c.tx, 4.737, 2.891, 0.337);
    body(s, COPY, c.tx, 5.144, 2.543, 0.985);
  });
}

function slide19(pptx) {
  const s = page(pptx);
  title(s, 'Creating Future Specialist', 0.896, 1.831, 5.485, 1.313);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam.', 0.931, 3.676, 4.304, 0.985);
  check(s, 1.12, 5.21, 0.26, NAVY);
  check(s, 1.30, 5.21, 0.26, NAVY);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod,', 1.797, 4.869, 3.683, 0.682);
  medallion(s, 10.586, 3.75, 4.133, NAVY);
  head(s, 'Project Success', 9.073, 2.596, 3.024, 0.337, { color: WHITE, align: 'center' });
  body(s, 'Lorem ipsum dolor consectetur adipiscing do eiusmod tempor incididunt ut.', 8.95, 3.067, 3.27, 0.682, { color: WHITE, align: 'center' });
  [
    { n: '15K', nx: 9.156, t: 'Result', tx: 9.156, tw: 1.384 },
    { n: '130+', nx: 10.538, t: 'Review', tx: 10.445, tw: 1.569 },
  ].forEach((k) => {
    stat(s, k.n, k.nx, 4.015, 1.384, 0.505, { color: WHITE, align: 'center' });
    head(s, k.t, k.tx, 4.567, k.tw, 0.337, { color: WHITE, align: 'center' });
  });
}

function slide20(pptx) {
  const s = page(pptx);
  title(s, 'Confident with Your Own Money', 1.521, 1.73, 5.145, 1.313);
  medallion(s, 9.752, 3.05, 0.786, SKY);
  icon(s, 'gearGlobe', 9.752, 3.05, 0.34, HEAD, SKY);
  medallion(s, 3.623, 6.055, 0.786, NAVY);
  icon(s, 'personChart', 3.623, 6.055, 0.31, WHITE, NAVY);
  const COPY = 'Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore.';
  [
    { n: '01.', nx: 4.687, tx: 5.239, t: 'Financial Management', bx: 4.682 },
    { n: '02.', nx: 8.779, tx: 9.330, t: 'Consultation Finance', bx: 8.773 },
  ].forEach((r) => {
    head(s, r.n, r.nx, 4.457, 0.639, 0.337, { color: NAVY });
    head(s, r.t, r.tx, 4.457, 3.083, 0.337);
    body(s, COPY, r.bx, 4.908, 3.257, 0.985);
  });
}

function slide21(pptx) {
  const s = page(pptx);
  title(s, 'Company Main Services', 2.828, 1.124, 7.677, 0.707, { align: 'center' });
  medallion(s, 6.667, 3.984, 3.535, NAVY);
  [
    { t: 'First Service', hx: 1.127, hy: 5.642, bx: 1.040, by: 5.927, bh: 0.682, c: 'Lorem ipsum dolor sit amet consectetur adipiscing sed do eiusmod.' },
    { t: 'Second Service', hx: 5.241, hy: 5.972, bx: 5.154, by: 6.257, bh: 0.379, c: 'Lorem ipsum dolor sit amet consectetur.' },
    { t: 'Third Service', hx: 9.210, hy: 5.635, bx: 9.123, by: 5.921, bh: 0.682, c: 'Lorem ipsum dolor sit amet consectetur adipiscing sed do eiusmod.' },
  ].forEach((r) => {
    head(s, r.t, r.hx, r.hy, 3.083, 0.337, { align: 'center' });
    body(s, r.c, r.bx, r.by, 3.257, r.bh, { align: 'center' });
  });
}

function slide22(pptx) {
  const s = page(pptx);
  title(s, 'Bright Future for Your Best Investment', 5.32, 1.687, 5.804, 1.313);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna.', 6.993, 3.479, 4.884, 0.682);
  medallion(s, 2.249, 4.61, 1.869, NAVY);
  icon(s, 'banknote', 2.249, 4.314, 0.4, WHITE, NAVY);
  body(s, 'Lorem ipsum dolor adipiscing elit', 1.333, 4.549, 1.831, 0.682, { color: WHITE, align: 'center' });
  medallion(s, 8.217, 5.227, 1.172, SKY);
  s.addText('92%', { x: 7.631, y: 5.008, w: 1.172, h: 0.438, fontFace: SEMI, fontSize: 20, bold: true, color: HEAD, align: 'center' });
  head(s, 'Plan and Strategy', 9.111, 4.737, 3.083, 0.337);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod.', 9.111, 5.034, 3.157, 0.682);
}

function slide23(pptx) {
  const s = page(pptx);
  title(s, 'A Brand New Financial Attitude', 0.896, 1.831, 5.804, 1.313);
  const COPY = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore.';
  [
    { y: 3.744, t: 'Manage Investment', by: 4.118 },
    { y: 5.104, t: 'Financial Planning', by: 5.477 },
  ].forEach((r) => {
    circle(s, 1.177, r.y + 0.138, 0.275, NAVY);
    check(s, 1.177, r.y + 0.138, 0.139, WHITE);
    head(s, r.t, 1.612, r.y - 0.03, 3.083, 0.337);
    body(s, COPY, 0.929, r.by, 4.884, 0.682);
  });
  medallion(s, 10.228, 3.75, 4.847, NAVY);
  head(s, 'Expert Finance', 9.563, 2.536, 2.525, 0.337, { color: WHITE });
  body(s, 'Lorem ipsum dolor consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ', 9.563, 3.123, 2.779, 1.287, { color: WHITE });
  s.addText('Read More . . .', { x: 9.563, y: 4.661, w: 1.779, h: 0.303, fontFace: SEMI, fontSize: 12, bold: true, italic: true, color: WHITE });
}

function slide24(pptx) {
  const s = page(pptx);
  title(s, 'The New Perspective of Investment', 6.339, 1.371, 5.766, 1.313);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ', 6.339, 2.853, 5.766, 0.682);
  [
    { cx: 8.021, disc: NAVY, fg: WHITE, bg: WHITE, n: '3650', nx: 7.435, hx: 7.087, bx: 7.012, t: 'Detail One' },
    { cx: 10.730, disc: SKY, fg: HEAD, bg: BODY, n: '750+', nx: 10.143, hx: 9.796, bx: 9.721, t: 'Detail Two' },
  ].forEach((c) => {
    medallion(s, c.cx, 5.188, 2.403, c.disc);
    stat(s, c.n, c.nx, 4.432, 1.172, 0.505, { color: c.fg, align: 'center' });
    head(s, c.t, c.hx, 4.935, 1.867, 0.337, { color: c.fg, align: 'center' });
    body(s, 'Lorem ipsum dolor sit amet adipiscing.', c.bx, 5.291, 2.017, 0.682, { color: c.bg, align: 'center' });
  });
}

function slide25(pptx) {
  const s = page(pptx);
  title(s, 'You Will Get Good Profit', 0.617, 2.626, 4.64, 1.313);
  body(s, 'Lorem ipsum dolor sit amet, veniam adipiscing elit, sed do eiusmod tempor incididunt.', 0.607, 4.182, 3.775, 0.682);
  [
    { cx: 9.099, cy: 1.988, disc: NAVY, glyph: 'pieSlice', gc: WHITE, hx: 9.807, hy: 1.472, hw: 3.083, bw: 2.553, t: 'Maximum Profit', c: 'Lorem ipsum dolor sit amet elit, sed do eiusmod tempor.' },
    { cx: 9.753, cy: 3.750, disc: SKY, glyph: 'target', gc: HEAD, hx: 10.435, hy: 3.234, hw: 1.925, bw: 2.037, t: 'Great Result', c: 'Lorem ipsum dolor sit do eiusmod tempor.' },
    { cx: 9.204, cy: 5.513, disc: NAVY, glyph: 'gearRing', gc: WHITE, hx: 9.807, hy: 4.997, hw: 3.083, bw: 2.553, t: 'Success Planning', c: 'Lorem ipsum dolor sit amet elit, sed do eiusmod tempor.' },
  ].forEach((r) => {
    medallion(s, r.cx, r.cy, 0.837, r.disc);
    icon(s, r.glyph, r.cx, r.cy, 0.33, r.gc, r.disc);
    head(s, r.t, r.hx, r.hy, r.hw, 0.337);
    body(s, r.c, r.hx, r.hy + 0.349, r.bw, 0.682);
  });
}

function slide26(pptx) {
  const s = page(pptx);
  title(s, 'Smart Investing with Smart Ideas', 0.896, 1.831, 5.32, 1.313);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation.', 0.896, 3.488, 4.719, 0.985);
  medallion(s, 1.407, 5.343, 0.837, NAVY);
  icon(s, 'coins', 1.407, 5.343, 0.33, WHITE, NAVY);
  head(s, 'Investing Instrument', 2.061, 5.002, 3.083, 0.337);
  body(s, 'Lorem ipsum dolor amet, consectetur adipiscing.', 2.061, 5.362, 4.884, 0.379);
  const COPY = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore.';
  [
    { t: 'Target Analysis', hy: 1.630, by: 2.034 },
    { t: 'Better Strategy', hy: 4.488, by: 4.892 },
  ].forEach((r) => {
    head(s, r.t, 9.421, r.hy, 3.083, 0.337);
    body(s, COPY, 9.421, r.by, 3.23, 0.985);
  });
}

function slide27(pptx) {
  const s = page(pptx);
  title(s, 'Strategies to Grow Your Money ', 7.495, 1.658, 5.366, 1.313);
  [
    { cy: 2.742, disc: SKY, fg: HEAD, bg: BODY, n: '250K', nx: 1.505, nw: 1.172, ny: 2.018, t: 'Strategy 01', hy: 2.527, by: 2.830 },
    { cy: 5.088, disc: NAVY, fg: WHITE, bg: WHITE, n: '88,6%', nx: 1.408, nw: 1.367, ny: 4.364, t: 'Strategy 02', hy: 4.873, by: 5.176 },
  ].forEach((c) => {
    medallion(s, 2.091, c.cy, 2.168, c.disc);
    stat(s, c.n, c.nx, c.ny, c.nw, 0.505, { color: c.fg, align: 'center' });
    head(s, c.t, 1.158, c.hy, 1.867, 0.337, { color: c.fg, align: 'center' });
    body(s, 'Lorem ipsum dolor sit amet adipiscing.', 1.083, c.by, 2.017, 0.682, { color: c.bg, align: 'center' });
  });
  [
    { n: '01.', ny: 3.655, by: 3.545, c: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod incididunt ut labore.' },
    { n: '02.', ny: 4.571, by: 4.460, c: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ut labore.' },
    { n: '03.', ny: 5.497, by: 5.376, c: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut.' },
  ].forEach((r) => {
    stat(s, r.n, 7.495, r.ny, 0.908, 0.505, { color: NAVY });
    body(s, r.c, 8.403, r.by, 4.058, 0.682);
  });
}

function slide28(pptx) {
  const s = page(pptx);
  title(s, 'Growing Money Is Our Hobby', 0.896, 2.049, 5.804, 1.313);
  [
    { n: '11.500', x: 0.896 },
    { n: '5000+', x: 3.680 },
  ].forEach((k) => {
    stat(s, k.n, k.x, 4.138, 1.616, 0.505, { color: NAVY });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.', k.x, 4.672, 2.5, 0.985);
  });
  medallion(s, 6.989, 5.976, 1.049, NAVY);
  icon(s, 'chat', 6.989, 5.951, 0.36, WHITE, NAVY);
}

function slide29(pptx) {
  const s = page(pptx);
  medallion(s, 6.666, 3.908, 5.08, NAVY);
  title(s, 'Start Invest Now!', 4.756, 2.321, 3.821, 1.313, { color: WHITE, align: 'center' });
  head(s, 'For Better Future', 5.119, 4.12, 3.083, 0.337, { color: WHITE, align: 'center' });
  body(s, 'Lorem ipsum dolor amet, consectetur adipiscing elit sed do eiusmod tempor incididunt labore dolore.', 5.045, 4.524, 3.23, 0.985, { color: WHITE, align: 'center' });
}

function slide30(pptx) {
  const s = pptx.addSlide();
  // Large navy panel with a concave left edge, anchored to the right of the page.
  s.addShape('custGeom', {
    x: 3.986, y: 0, w: 9.3468, h: 7.5,
    fill: { color: NAVY }, line: { type: 'none' },
    points: [
      { x: 1.6208, y: 0, moveTo: true }, { x: 9.3468, y: 0 }, { x: 9.3468, y: 7.5 },
      { x: 1.2826, y: 7.5 }, { x: 1.1455, y: 7.3286 },
      { x: 0, y: 3.9375, curve: { type: 'cubic', x1: 0.4261, y1: 6.3830, x2: 0, y2: 5.2091 } },
      { x: 1.4979, y: 0.1267, curve: { type: 'cubic', x1: 0, y1: 2.4703, x2: 0.5672, y2: 1.1332 } },
      { close: true },
    ],
  });
  brand(s, false);
  footer(s);
  title(s, 'Contact Us and Make Big Money', 7.187, 1.957, 5.094, 1.313, { color: WHITE });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eiusmod tempor incididunt ut labore et dolore magna aliqua.', 7.228, 3.86, 5.317, 0.682, { color: WHITE });
  circle(s, 7.525, 5.337, 0.412, WHITE);
  icon(s, 'handset', 7.525, 5.337, 0.2, NAVY, WHITE);
  head(s, '(+00) 123-456-789', 7.968, 5.169, 3.083, 0.337, { color: WHITE });
}

/* ==================================================================== build */
const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'MONETY', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'MONETY';
  pptx.author = 'Monety Consultant';
  pptx.title = 'Financial Consultant Presentation';
  BUILDERS.forEach((fn) => fn(pptx));
  return pptx.writeFile({ fileName: path.join(__dirname, '07417286-b80e-4f05-8638-83806c4eab11_grok_final.pptx') });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
