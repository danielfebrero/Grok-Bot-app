/**
 * "HUGE." brand-proposal deck — rebuilt with pptxgenjs.
 * 20 slides, 13.333in x 7.5in (16:9).
 *
 * Everything below is plain data + plain pptxgenjs calls: a palette, a font
 * table, a handful of drawing helpers, a few composite blocks that the deck
 * repeats (page chrome, nav lists, testimonials, work items) and then one
 * builder function per slide.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ── palette ─────────────────────────────────────────────────────────── */
const C = {
  ink: '262626', // near-black used for type and dark panels
  black: '000000',
  white: 'FFFFFF',
  gold: '938558', // accent2
  sand: 'E8DFD6', // accent3
  cream: 'EAE6DB',
  paper: 'F2F2F2', // accent4
  taupe: 'D7C8B9', // the "2027" pill
  khaki: 'C1B696',
  rule: 'B5B5B5', // hairline grid
  ruleMid: '797979',
  ruleDark: '3C3C3C',
  ruleSoft: '3F3F3F',
  ruleFaint: 'D9D9D9',
};

/* ── fonts (all embedded in the source deck) ─────────────────────────── */
const F = {
  bebas: 'Bebas Neue',
  pop: 'Poppins',
  popMed: 'Poppins Medium',
  popLight: 'Poppins Light',
  popThin: 'Poppins Thin',
  mont: 'Montserrat',
  oswald: 'Oswald',
  inter: 'Inter SemiBold',
};

/* Text-box insets of the source deck, in points: [left, right, bottom, top]. */
const IN = [7.2, 7.2, 3.6, 3.6];

const SLIDE_W = 13.3333333;
const SLIDE_H = 7.5;

/* ── body copy ───────────────────────────────────────────────────────── */
const LOREM0 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
  'incididunt ut labore et dolore magna aliqua. Ut enim ad';
const LOREM1 = LOREM0 + ' minim veniam, quis';
const LOREM2 = LOREM1 + ' nostrud exercitation ullamco laboris nisi ut aliquip ex ea';
const LOREM3 = LOREM2 + ' commodo consequat. Duis aute irure dolor in reprehenderit in voluptate';
const LOREM4 = LOREM3 + ' velit esse cillum dolore eu fugiat nulla pariatur';
const LOREM5 =
  LOREM4 + '. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit';

const MUST = 'PLACEHOLDER';
const MUST_P = MUST + ' pleasure';
const MUST_PA = MUST_P + ' and';
const MUST_PAL = MUST_PA + ' Lorem';
const MUST_LONG = MUST_PA + ' Lorem must explain to you how all';
const MUST_TYPO = 'PLACEHOLDER';

/* ── primitive helpers ───────────────────────────────────────────────── */
function txt(s, body, o) {
  s.addText(body, {
    fontFace: F.pop,
    fontSize: 10,
    color: C.ink,
    align: 'left',
    valign: 'top',
    margin: IN,
    isTextBox: true,
    ...o,
  });
}

function rect(s, x, y, w, h, o) {
  s.addShape('rect', { x, y, w, h, ...o });
}

function bg(s, color, transparency) {
  rect(s, 0, 0, SLIDE_W, SLIDE_H, { fill: { color, transparency } });
}

function ellipse(s, x, y, w, h, o) {
  s.addShape('ellipse', { x, y, w, h, ...o });
}

function hline(s, x, y, w, color, pt, transparency) {
  s.addShape('line', { x, y, w, h: 0, line: { color, width: pt || 0.75, transparency } });
}

function vline(s, x, y, h, color, pt, transparency) {
  s.addShape('line', { x, y, w: 0, h, line: { color, width: pt || 0.75, transparency } });
}

/**
 * The deck's signature "arch": a rectangle whose top corners are rounded by a
 * full half-width radius (round2SameRect with adj1 = 50%), drawn as custGeom.
 */
function arch(s, x, y, w, h, o) {
  const r = w / 2;
  s.addShape('custGeom', {
    x,
    y,
    w,
    h,
    points: [
      { x: 0, y: h },
      { x: 0, y: r },
      { x: r, y: 0, curve: { type: 'arc', hR: r, wR: r, stAng: 180, swAng: 90 } },
      { x: w - r, y: 0 },
      { x: w, y: r, curve: { type: 'arc', hR: r, wR: r, stAng: 270, swAng: 90 } },
      { x: w, y: h },
      { close: true },
    ],
    ...o,
  });
}

/**
 * Empty picture placeholder from the source deck (the template ships with no
 * bitmaps at all — every image slot is an unfilled frame). Reproduced as a
 * transparent frame so the geometry survives without inventing artwork.
 */
function picFrame(s, x, y, w, h, shape) {
  s.addShape(shape || 'rect', { x, y, w, h, fill: { type: 'none' }, line: { type: 'none' } });
}

/* ── page chrome: PROPOSAL / 2027 / TEMPLATE tabs on every slide ─────── */
function chrome(s, o) {
  const opt = { stroke: C.ink, ink: C.ink, pill: C.taupe, pillInk: C.ink, pillY: 0.221, radius: 0, dx: 0, ...o };
  const pillOpts = opt.radius ? { rectRadius: opt.radius } : {};
  const tab = opt.radius ? 'roundRect' : 'rect';

  s.addShape(tab, {
    x: opt.radius ? 0.228 : 0.221,
    y: 0.228,
    w: opt.radius ? 1.379 : 1.385,
    h: 0.236,
    line: { color: opt.stroke, width: 1 },
    ...pillOpts,
  });
  txt(s, 'PROPOSAL', { x: 0.343, y: 0.228, w: 1.148, h: 0.236, fontSize: 8, bold: true, align: 'center', color: opt.ink });

  s.addShape(tab, { x: 5.854, y: opt.pillY, w: 0.974, h: 0.236, fill: { color: opt.pill }, ...pillOpts });
  txt(s, '2027', { x: 6.102, y: 0.228, w: 0.478, h: 0.236, fontSize: 8, bold: true, align: 'center', color: opt.pillInk });

  s.addShape(tab, {
    x: 11.15 + opt.dx,
    y: 0.228,
    w: 1.941,
    h: 0.236,
    line: { color: opt.stroke, width: 1 },
    ...pillOpts,
  });
  txt(s, 'TEMPLATE', {
    x: 11.481 + opt.dx, y: 0.228, w: 1.28, h: 0.236, fontSize: 8, bold: true, align: 'center', color: opt.ink,
  });
}

/* ── the numbered table-of-contents strip, reused on 9 slides ────────── */
/* rows are [number, number-box width, label, label-box width] */
const NAV_COL1 = [['1.', 0.375, 'Cover', 0.626], ['2.', 0.375, 'Welcome', 0.974], ['3.', 0.375, 'About Us.', 0.974]];
const NAV_COL2 = [['11.', 0.375, 'Manifesto', 1.001], ['12.', 0.404, 'Market Values', 1.001]];

/* Row baselines: three-row blocks are tighter-set than two-row ones. */
const NAV_ROWS3 = [0, 0.24161, 0.48471];
const NAV_ROWS2 = [0, 0.20987];

function navList(s, x, y, o) {
  const opt = { rows: 3, color: C.black, twoCols: true, ...o };
  const dy = opt.rows === 2 ? NAV_ROWS2 : NAV_ROWS3;
  const style = { fontFace: F.popMed, fontSize: 8, color: opt.color, h: 0.236 };
  NAV_COL1.slice(0, opt.rows).forEach(([n, nw, label, lw], i) => {
    txt(s, n, { ...style, x, y: y + dy[i], w: nw });
    txt(s, label, { ...style, x: x + 0.33316, y: y + dy[i], w: lw });
  });
  if (!opt.twoCols) return;
  NAV_COL2.forEach(([n, nw, label, lw], i) => {
    txt(s, n, { ...style, x: x + 1.50852, y: y + dy[i], w: nw });
    txt(s, label, { ...style, x: x + 1.89392, y: y + dy[i], w: lw });
  });
}

/* ── quote + name + role block ("SARAH JOHESON") ─────────────────────── */
function testimonial(s, x, y, color, o) {
  const opt = { quote: MUST, quoteW: 4.441, name: 'SARAH JOHESON', role: 'CEO / Manager', ...o };
  txt(s, opt.quote, {
    x, y, w: opt.quoteW, h: 0.486, fontSize: 10.5, color, lineSpacingMultiple: 1.3333,
  });
  txt(s, MUST_PA, {
    x, y: y + 0.59348, w: 2.39966, h: 1.44733, fontSize: 14, color, lineSpacingMultiple: 1.1428,
  });
  if (opt.name) txt(s, opt.name, { x: x + 2.35675, y: y + 0.61381, w: 1.96869, h: 0.337, fontSize: 14, color });
  if (opt.role) txt(s, opt.role, { x: x + 2.35675, y: y + 0.9229, w: 1.52939, h: 0.269, fontSize: 10, color });
}

/* ── "01 / Work Title Here / paragraph" column entry ─────────────────── */
function workItem(s, x, y, num, o) {
  const opt = { color: C.sand, bodyFont: F.pop, bodyW: 2.713, bold: false, numFont: null, ...o };
  const head = { fontSize: 11, color: opt.color, bold: opt.bold, h: 0.286 };
  txt(s, num, { ...head, x, y, w: 0.502, ...(opt.numFont ? { fontFace: opt.numFont, bold: false } : {}) });
  txt(s, 'Work Title Here', { ...head, x, y: y + 0.27, w: 1.612 });
  txt(s, MUST_P, {
    x, y: y + 0.505, w: opt.bodyW, h: 0.594,
    fontFace: opt.bodyFont, fontSize: 9, color: opt.color, lineSpacingMultiple: 1.3333,
  });
}

/* ── circled percentage badge ────────────────────────────────────────── */
function badge(s, x, y, label, dy) {
  ellipse(s, x, y, 0.794, 0.794, { fill: { color: C.sand } });
  txt(s, label, {
    x: x + 0.00863, y: y + (dy || 0.19495), w: 0.77764, h: 0.37025, fontSize: 16, bold: true, align: 'center',
  });
}

/* ── outlined circle with a big italic Bebas number ──────────────────── */
function numberCircle(s, x, y, d, num, o) {
  const opt = { stroke: C.rule, strokePt: 0.75, color: C.ruleDark, size: 35, dx: 0.2023, dy: 0.36367, textW: 0.96974, textH: 0.69, ...o };
  ellipse(s, x, y, d, d, { line: { color: opt.stroke, width: opt.strokePt } });
  txt(s, num, {
    x: x + opt.dx, y: y + opt.dy, w: opt.textW, h: opt.textH,
    fontFace: F.bebas, fontSize: opt.size, bold: true, italic: true, color: opt.color, align: 'center',
  });
}

/* ── tall vertical "TEMPLATE" ribbon on the deck's quiet slides ──────── */
function ribbon(s, x, y, stroke, o) {
  const opt = { paraDx: -1.511, tabInk: stroke, paraInk: stroke, ...o };
  ellipse(s, x, y, 0.531, 4.229, { line: { color: stroke, width: 0.75 } });
  txt(s, 'TEMPLATE', {
    x: x - 0.37497, y: y + 1.99705, w: 1.28021, h: 0.2356,
    fontSize: 8, bold: true, color: opt.tabInk, align: 'center', rotate: -90,
  });
  txt(s, LOREM4, {
    x: x + opt.paraDx, y: y + 1.735, w: 5.494, h: 0.76,
    fontFace: F.mont, fontSize: 8, color: opt.paraInk, align: 'center', lineSpacingMultiple: 1.5, rotate: -90,
  });
}

/* ── big Bebas section title ─────────────────────────────────────────── */
function title(s, x, y, w, str, o) {
  txt(s, str, {
    x, y, w, h: 0.69, fontFace: F.bebas, fontSize: 35, bold: true, color: C.black, ...o,
  });
}

/* ═════════════════════════════ slides ════════════════════════════════ */

/* 1 — cover, dark ground with a pale panel */
function slide01(s) {
  bg(s, C.ink);
  rect(s, 5.146, 0.228, 8.188, 7.05, { fill: { color: C.paper } });
  picFrame(s, 0.221, 0.221, 6.064, 7.057);
  txt(s, 'HUGE.', { x: 6.534, y: 0.51, w: 7.048, h: 4.393, fontFace: F.bebas, fontSize: 255, bold: true });
  txt(s, 'BRAND PROPOSAL', { x: 6.656, y: 4.248, w: 3.724, h: 0.404, fontFace: F.popThin, fontSize: 18 });
  txt(s, 'brand proposal', { x: 6.67, y: 4.679, w: 1.343, h: 0.269, bold: true });
  txt(s, [{ text: 'PLACEHOLDER', options: { breakLine: true } },
    { text: 'and denouncing pleasure and' }], {
    x: 6.666, y: 4.949, w: 5.366, h: 0.431, fontFace: F.popLight, fontSize: 9, lineSpacingMultiple: 1.3333,
  });
  hline(s, 6.777, 5.452, 0.65, C.ruleSoft);
  ellipse(s, 6.777, 5.68, 0.355, 0.355, { fill: { color: C.ink } });
  chrome(s, { pill: C.gold, pillInk: C.white });
}

/* 2 — same cover, inverted into the gold colourway */
function slide02(s) {
  bg(s, C.paper);
  rect(s, 5.146, 0.228, 8.188, 7.05, { fill: { color: C.gold } });
  picFrame(s, 0.221, 0.221, 6.064, 7.057);
  txt(s, 'HUGE.', { x: 6.555, y: 0.464, w: 7.048, h: 4.393, fontFace: F.bebas, fontSize: 255, bold: true, color: C.white });
  txt(s, 'BRAND PROPOSAL', { x: 6.656, y: 4.248, w: 3.724, h: 0.404, fontFace: F.popThin, fontSize: 18, color: C.white });
  txt(s, 'brand proposal', { x: 6.67, y: 4.696, w: 1.343, h: 0.269, bold: true, color: C.white });
  txt(s, [{ text: 'PLACEHOLDER', options: { breakLine: true } },
    { text: 'and denouncing pleasure and' }], {
    x: 6.666, y: 4.949, w: 5.366, h: 0.431, fontFace: F.popLight, fontSize: 9, color: C.white, lineSpacingMultiple: 1.3333,
  });
  hline(s, 6.777, 5.452, 0.65, C.white);
  ellipse(s, 6.777, 5.68, 0.355, 0.355, { fill: { color: C.khaki } });
  chrome(s, { stroke: C.white, ink: C.white, pill: C.khaki, pillInk: C.white, pillY: 0.228, radius: 0.118 });
}

/* 3 — welcome / message */
function slide03(s) {
  bg(s, C.cream, 55.295);
  hline(s, 0, 3.75, SLIDE_W, C.rule);
  arch(s, 0.961, 0.835, 4.492, 5.83, { fill: { type: 'none' }, line: { color: C.rule, width: 1 } });
  vline(s, 6.302, 0, SLIDE_H, C.rule);
  chrome(s);
  title(s, 6.892, 2.551, 3.371, 'WELCOME/MESSAGE.', { color: C.ink });
  testimonial(s, 6.92389, 3.19985, C.ink);
  picFrame(s, 1.51, 1.547, 3.395, 4.406);
}

/**
 * 4 — contents. Rows are [number, number-box w, label, label-box w, labelDy?].
 * Two right-column labels sit a hair lower than their numbers in the source
 * deck; `labelDy` keeps that quirk. The "13." box is also taller than its
 * siblings everywhere it appears.
 */
const CONTENTS_LEFT = [
  ['1.', 0.375, 'Cover', 0.626], ['2.', 0.375, 'Welcome', 0.974], ['3.', 0.375, 'About Us.', 0.974],
  ['4.', 0.375, 'Contents', 0.885], ['5.', 0.375, 'Mission & Vision', 1.414], ['6.', 0.375, 'History', 0.885],
];
const CONTENTS_MID = [
  ['7.', 0.375, 'Goals', 0.626], ['8.', 0.375, 'Brand Concept', 1.414], ['9.', 0.375, 'Case Study', 1.185],
  ['10.', 0.404, 'Brand insight', 1.273], ['11.', 0.375, 'Manifesto', 1.001], ['12.', 0.404, 'Brand Purpose', 1.414],
  ['13.', 0.375, 'Market Values', 1.414], ['14.', 0.404, 'Thank you', 1.273], ['15.', 0.492, 'Other…', 1.099],
];
const CONTENTS_RIGHT = [
  ['11.', 0.375, 'Manifesto', 1.001], ['12.', 0.404, 'Market Values', 1.414, 0.02356],
  ['13.', 0.375, 'Brand Purpose', 1.414, 0.02697], ['14.', 0.404, 'Thank you', 1.273], ['15.', 0.492, 'Other…', 0.885],
];

function contentsColumn(s, numX, labelX, y0, rows) {
  rows.forEach(([n, nw, label, lw, labelDy], i) => {
    const y = y0 + i * 0.24571;
    txt(s, n, { x: numX, y, w: nw, h: n === '13.' ? 0.438 : 0.269, bold: true, color: C.black });
    txt(s, label, { x: labelX, y: y + (labelDy || 0), w: lw, h: 0.269, bold: true, color: C.black });
  });
}

function slide04(s) {
  bg(s, C.cream, 55.295);
  hline(s, 0, 3.75, SLIDE_W, C.rule);
  vline(s, 6.302, 0, SLIDE_H, C.rule);
  rect(s, 4.051, 0.835, 4.492, 6.436, { fill: { color: C.cream } });
  chrome(s);
  contentsColumn(s, 6.592, 7.084, 1.94303, CONTENTS_LEFT);
  contentsColumn(s, 6.626, 7.118, 4.03263, CONTENTS_MID);
  contentsColumn(s, 8.752, 9.244, 1.95643, CONTENTS_RIGHT);
  title(s, 0.117, 2.926, 1.853, 'CONTENTS.');
  [['5%', 8.753], ['10%', 10.245], ['15%', 11.719]].forEach(([label, x]) => {
    txt(s, label, { x, y: 6.483, w: 0.544, h: 0.269, bold: true, color: C.black });
  });
  txt(s, LOREM3, {
    x: 0.133, y: 4.017, w: 3.232, h: 1.462, fontFace: F.popLight, fontSize: 9, color: C.black, lineSpacingMultiple: 1.5555,
  });
  [8.844, 10.338, 11.833].forEach((x) => picFrame(s, x, 4.088, 1.273, 2.0));
}

/* 5 — our overview */
function slide05(s) {
  bg(s, C.cream, 55.295);
  rect(s, 4.481, 2.274, 8.652, 5.005, { fill: { color: C.cream } });
  arch(s, 8.903, 1.034, 3.424, 4.68, { fill: { type: 'none' }, line: { color: C.rule, width: 1 } });
  chrome(s);
  txt(s, [{ text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed ' }, { text: ' ' },
    { text: 'do eiusmod tempor' }], {
    x: 0.136, y: 5.972, w: 3.431, h: 0.48, fontSize: 9, color: C.black, lineSpacingMultiple: 1.5555,
  });
  txt(s, MUST_LONG, {
    x: 9.173, y: 3.952, w: 2.883, h: 0.873, fontSize: 9, align: 'center', lineSpacingMultiple: 1.5555,
  });
  txt(s, 'Strategy', { x: 9.882, y: 3.498, w: 1.465, h: 0.404, fontSize: 18, bold: true, align: 'center' });
  vline(s, 10.615, 2.274, 0.915, C.rule);
  vline(s, 10.615, 5.111, 1.186, C.rule);
  navList(s, 0.14092, 6.66339);
  hline(s, 0, 0.464, SLIDE_W, C.rule);
  vline(s, 3.952, 0, SLIDE_H, C.rule);
  title(s, 0.117, 4.654, 2.66, 'Our OVERVIEW.');
  picFrame(s, 4.481, 1.034, 3.616, 5.264);
  picFrame(s, 0.237, 1.034, 3.19, 3.512);
}

/* 6 — about us */
function slide06(s) {
  bg(s, C.cream, 55.295);
  hline(s, 0, 3.75, SLIDE_W, C.rule);
  chrome(s);
  vline(s, 4.302, 0, SLIDE_H, C.rule);
  title(s, 0.153, 2.009, 1.994, 'About us ?.');
  txt(s, '01.(Works)', { x: 10.56, y: 3.894, w: 1.526, h: 0.334, fontSize: 11, color: C.black, lineSpacingMultiple: 1.6364 });
  txt(s, 'Lorem must explain to you how all this mist', {
    x: 10.56, y: 4.24, w: 2.743, h: 0.431, fontSize: 9, color: C.black, lineSpacingMultiple: 1.3333,
  });
  ellipse(s, 10.671, 4.66, 0.75, 0.75, { line: { color: C.gold, width: 1 } });
  txt(s, '96%', { x: 10.671, y: 4.833, w: 0.75, h: 0.404, fontSize: 18, bold: true, color: C.gold });
  navList(s, 10.57581, 2.12225, { twoCols: false });
  txt(s, LOREM1, {
    x: 0.133, y: 2.738, w: 3.232, h: 0.873, fontFace: F.popLight, fontSize: 9, color: C.black, lineSpacingMultiple: 1.5555,
  });
  navList(s, 0.14092, 3.89724);
  picFrame(s, 4.748, 0.667, 5.208, 6.612);
}

/* 7 — deliverables, full-bleed gold */
function slide07(s) {
  bg(s, C.gold);
  chrome(s, { stroke: C.sand, ink: C.sand, pillInk: C.ruleSoft, dx: 0.01 });
  workItem(s, 5.759, 4.047, '01');
  workItem(s, 5.759, 6.244, '03');
  workItem(s, 10.051, 4.047, '02');
  workItem(s, 10.051, 6.244, '04');
  hline(s, 0, 2.699, 5.854, C.sand);
  vline(s, 4.843, 0, SLIDE_H, C.sand);
  vline(s, 9.307, 4.11, 3.169, C.sand);
  vline(s, 9.046, 4.11, 3.169, C.sand);
  txt(s, 'deliverables.', {
    x: 0.112, y: 1.821, w: 3.076, h: 0.841, fontFace: F.bebas, fontSize: 44, bold: true, color: C.sand,
  });
  txt(s, LOREM1, {
    x: 0.154, y: 2.884, w: 3.232, h: 0.873, fontFace: F.popLight, fontSize: 9, color: C.sand, lineSpacingMultiple: 1.5555,
  });
  navList(s, 0.16176, 4.04307, { color: C.sand });
  hline(s, 0, 5.761, 9.046, C.sand);
  hline(s, 9.307, 5.761, 3.816, C.sand);
  txt(s, '10%', { x: 8.892, y: 5.551, w: 0.594, h: 0.286, fontSize: 11, color: C.sand, rotate: -90 });
  picFrame(s, 5.854, 0.605, 7.268, 3.262);
}

/* 8 — our objectives */
function slide08(s) {
  rect(s, 0, 0, 5.834, SLIDE_H, { fill: { color: C.paper } });
  chrome(s);
  hline(s, 0, 1.528, SLIDE_W, C.rule);
  vline(s, 5.834, 0, SLIDE_H, C.rule);
  title(s, 6.089, 1.704, 2.713, 'Our objectives.', { color: C.ink });
  navList(s, 0.14116, 1.75441, { rows: 2 });
  [['5%', 6.922], ['10%', 9.382], ['15%', 11.859]].forEach(([label, x]) => {
    txt(s, label, { x, y: 5.109, w: 0.544, h: 0.269, bold: true, color: C.black, align: 'center' });
  });
  txt(s, LOREM5, {
    x: 0.141, y: 6.343, w: 5.424, h: 0.928, fontFace: F.mont, fontSize: 8, color: C.black, lineSpacingMultiple: 1.5,
  });
  [6.533, 8.994, 11.47].forEach((x) => {
    txt(s, 'Lorem ipsum dolor consectetur', {
      x, y: 5.332, w: 1.322, h: 0.423, fontFace: F.mont, fontSize: 8, color: C.black,
      align: 'center', lineSpacingMultiple: 1.5,
    });
  });
  picFrame(s, 0.237, 2.41, 3.68, 3.748);
  [6.213, 8.674, 11.15].forEach((x) => picFrame(s, x, 2.663, 1.962, 2.179));
}

/* 9 — our goals */
function slide09(s) {
  bg(s, C.sand, 44.314);
  chrome(s);
  title(s, 5.757, 0.984, 2.713, 'Our goals.', { color: C.ink });
  txt(s, LOREM4, {
    x: 5.774, y: 1.674, w: 4.385, h: 0.928, fontFace: F.mont, fontSize: 8, color: C.black, lineSpacingMultiple: 1.5,
  });
  [['01', 3.03, 0.464, 3.38209], ['02', 4.36, 0.607, 4.71202], ['03', 5.948, 0.607, 6.30074]].forEach(([num, y, w, py]) => {
    txt(s, num, { x: 5.774, y, w, h: 0.404, fontFace: F.popMed, fontSize: 18, color: C.black });
    txt(s, LOREM2, {
      x: 5.774, y: py, w: 5.24926, h: 0.59155, fontFace: F.mont, fontSize: 8, color: C.black, lineSpacingMultiple: 1.5,
    });
  });
  hline(s, 5.363, 2.829, 7.971, C.rule);
  vline(s, 5.661, 0, SLIDE_H, C.rule);
  navList(s, 10.32336, 1.67984, { rows: 2 });
  hline(s, 0, 0.75, SLIDE_W, C.rule);
  picFrame(s, 0.237, 1.121, 5.125, 6.157);
}

/* 10 — brand concepts */
function slide10(s) {
  bg(s, C.cream, 55.295);
  picFrame(s, 1.046, 1.659, 3.977, 4.711);
  picFrame(s, 6.205, 1.659, 6.907, 4.711);
  rect(s, 1.046, 1.659, 3.977, 4.711, { fill: { type: 'none' }, line: { color: C.rule, width: 1 } });
  chrome(s);
  rect(s, 1.934, 2.711, 2.2, 2.606, { fill: { color: C.cream } });
  title(s, 6.076, 0.937, 2.93, 'Brand Concepts.');
  vline(s, 5.854, 0, SLIDE_H, C.rule);
  hline(s, 0, 1.099, 5.854, C.rule);
  hline(s, 0, 6.909, SLIDE_W, C.rule);
  hline(s, 0, 0.75, SLIDE_W, C.rule);
  title(s, 1.996, 3.103, 1.401, 'Brand');
  hline(s, 1.934, 3.793, 2.2, C.rule);
  txt(s, LOREM0, {
    x: 1.996, y: 3.867, w: 2.2, h: 0.928, fontFace: F.mont, fontSize: 8, color: C.black, lineSpacingMultiple: 1.5,
  });
  txt(s, LOREM4, {
    x: 7.645, y: 3.286, w: 4.385, h: 0.928, fontFace: F.mont, fontSize: 8, color: C.white, lineSpacingMultiple: 1.5,
  });
  navList(s, 7.65547, 4.32995, { rows: 2, color: C.white, twoCols: false });
}

/* 11 — brand schedule */
function slide11(s) {
  bg(s, C.cream, 55.295);
  [[1.805, C.ink], [3.859, C.gold], [5.914, C.sand]].forEach(([y, color]) => {
    rect(s, 7.601, y, 5.531, 0.644, { fill: { color } });
  });
  chrome(s);
  hline(s, 0, 1.088, SLIDE_W, C.rule);
  vline(s, 7.593, 0, SLIDE_H, C.rule);
  title(s, 0.134, 3.511, 3.387, 'Brand schedule.');
  txt(s, LOREM1, {
    x: 0.155, y: 1.185, w: 6.426, h: 0.48, fontFace: F.popLight, fontSize: 9, color: C.black, lineSpacingMultiple: 1.5555,
  });
  ['01', '02', '03'].forEach((num, i) => numberCircle(s, [0.23953, 2.66133, 5.06472][i], 1.784, 1.375, num));
  hline(s, 0, 3.197, SLIDE_W, C.rule);
  [[7.876, 1.975], [10.546, 4.029], [7.876, 6.084]].forEach(([x, y]) => {
    txt(s, 'concept', { x, y, w: 1.029, h: 0.304, fontSize: 12, color: C.white });
  });
  picFrame(s, 0.221, 4.201, 6.607, 3.078);
}

/* 12 — how we work, reversed out of near-black */
function slide12(s) {
  bg(s, C.ink);
  chrome(s, { stroke: C.white, ink: C.white, pill: C.white, pillInk: C.black });
  [[1.73818, 1.80805], [3.20703, 3.2539], [4.67587, 4.73942], [6.14471, 6.21894]].forEach(([cy, ty], i) => {
    numberCircle(s, 0.528, cy, 0.732, ['01', '02', '03', '04'][i], {
      color: C.white, size: 24, dx: 0.01777, dy: 0.13395, textW: 0.69643, textH: 0.50488,
    });
    txt(s, [{ text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, s', options: { breakLine: true } },
      { text: ' do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, ' }], {
      x: 1.593, y: ty, w: 3.586, h: 0.592, fontFace: F.mont, fontSize: 8, color: C.white, lineSpacingMultiple: 1.5,
    });
  });
  [1.313, 2.751, 4.272, 5.792].forEach((y) => hline(s, 0.242, y, 5.112, C.white));
  vline(s, 5.354, 0, 7.279, C.white);
  title(s, 0.134, 0.597, 2.463, 'How We work.', { color: C.white });
  ribbon(s, 11.58, 1.902, C.white);
  picFrame(s, 5.854, 0.742, 5.296, 6.536);
}

/* 13 — case study */
function slide13(s) {
  bg(s, C.cream, 55.295);
  rect(s, 5.854, 1.076, 7.258, 5.9, { fill: { color: C.gold } });
  chrome(s);
  hline(s, 0, 0.75, SLIDE_W, C.rule);
  ribbon(s, 0.443, 1.941, C.rule, { paraDx: -1.784, tabInk: C.black, paraInk: C.ink });
  title(s, 7.263, 2.091, 2.142, 'Case study.', { color: C.white });
  navList(s, 7.27643, 3.0074, { color: C.white });
  testimonial(s, 7.26254, 3.92078, C.white);
  vline(s, 5.854, 1.069, 5.907, C.white);
  vline(s, 6.229, 1.069, 5.907, C.white);
  arch(s, 6.988, 1.025, 5.451, 6.003, {
    fill: { type: 'none' }, line: { color: C.white, width: 1 }, rotate: 90,
  });
  picFrame(s, 1.94, 1.076, 3.914, 5.9);
}

/* 14 — offerings packages */
function slide14(s) {
  bg(s, C.sand);
  picFrame(s, 0.242, 1.659, 4.612, 5.62);
  chrome(s);
  hline(s, 0, 1.313, SLIDE_W, C.ruleDark);
  title(s, 0.12, 0.597, 3.505, 'Offerings packages.', { color: C.ink });
  vline(s, 5.354, 0, SLIDE_H, C.ruleMid);
  vline(s, 7.385, 0, SLIDE_H, C.ruleMid);
  [[1.66094, 1.88507], [3.33652, 3.56727], [4.99916, 5.12471]].forEach(([cy, ty], i) => {
    numberCircle(s, 5.694, cy, 1.375, ['01', '02', '03'][i], {
      stroke: C.ruleDark, size: 40, dx: 0.20307, dy: 0.36371, textW: 0.96974, textH: 0.77415,
    });
    txt(s, LOREM5, {
      x: 7.557, y: ty, w: 5.424, h: 0.928, fontFace: F.mont, fontSize: 8, color: C.black, lineSpacingMultiple: 1.5,
    });
  });
  hline(s, 5.354, 3.197, 7.979, C.ruleDark);
  hline(s, 5.354, 4.832, 7.979, C.ruleDark);
  rect(s, 1.067, 2.751, 2.901, 3.436, { fill: { color: C.sand } });
  title(s, 1.479, 3.557, 1.401, 'Brand');
  hline(s, 1.417, 4.247, 2.2, C.rule);
  txt(s, LOREM0, {
    x: 1.479, y: 4.322, w: 2.2, h: 0.928, fontFace: F.mont, fontSize: 8, color: C.black, lineSpacingMultiple: 1.5,
  });
}

/* 15 — the estimate */
function slide15(s) {
  rect(s, 0, 1.42, SLIDE_W, 5.851, { fill: { color: C.paper } });
  chrome(s);
  title(s, 0.12, 0.613, 2.515, 'The estimate.', { color: C.ink });
  [['Experience', 2.755, 1.508], ['Subtitle Here', 5.2, 1.711]].forEach(([label, x, w]) => {
    txt(s, MUST_PAL, {
      x, y: 2.956, w: 2.354, h: 0.868, fontFace: F.mont, fontSize: 9, color: C.ink, lineSpacingMultiple: 1.5555,
    });
    txt(s, label, { x, y: 2.558, w, h: 0.37, fontSize: 16, bold: true, color: C.ink });
  });
  hline(s, 2.831, 3.909, 3.18, C.ink, 1, 8.236);
  hline(s, 2.831, 4.766, 3.18, C.ink, 1, 8.236);
  hline(s, 2.815, 5.623, 3.18, C.ruleFaint, 1);
  [['Individuality', 4.08, 1.368, 4.048], ['Sustainability', 4.937, 1.528, 4.905]].forEach(([label, y, w, by]) => {
    txt(s, label, { x: 2.736, y, w, h: 0.303, fontSize: 12, bold: true, color: C.ink });
    txt(s, MUST_PAL, {
      x: 4.175, y: by, w: 4.553, h: 0.476, fontFace: F.mont, fontSize: 9, color: C.ink, lineSpacingMultiple: 1.5555,
    });
  });
  badge(s, 0.239, 5.644, '72%.');
  txt(s, MUST_PA, {
    x: 1.523, y: 5.815, w: 5.677, h: 0.486, fontFace: F.mont, fontSize: 10.5, color: C.ink, lineSpacingMultiple: 1.3333,
  });
  txt(s, 'SARAH JOHESON', { x: 0.123, y: 2.608, w: 1.969, h: 0.337, fontSize: 14, color: C.ink });
  txt(s, 'Brand portfolio', { x: 0.12, y: 4.483, w: 1.529, h: 0.269, fontSize: 10, color: C.ink });
  txt(s, MUST_PA, {
    x: 0.127, y: 2.993, w: 2.4, h: 1.447, fontSize: 14, color: C.ink, lineSpacingMultiple: 1.1428,
  });
  hline(s, 0, 2.087, SLIDE_W, C.ruleDark);
  vline(s, 8.846, 0, SLIDE_H, C.ruleMid);
  navList(s, 9.04984, 2.59371, { color: C.ink });
  testimonial(s, 9.03612, 3.50719, C.ink, { quote: MUST_TYPO, quoteW: 4.076, role: null });
}

/* 16 — the team */
const TEAM = [
  { name: 'James Vane', role: 'CEO Manager', x: 0.23197, nameX: 0.772, nameW: 1.51, roleX: 0.953, roleW: 1.149, y: 5.062 },
  { name: 'Henry Wotton', role: 'Art Design', x: 3.66349, nameX: 4.102, nameW: 1.714, roleX: 4.489, roleW: 0.939, y: 5.062 },
  { name: 'Basda auburn', role: 'Ui / X Developer', x: 7.095, nameX: 7.636, nameW: 1.51, roleX: 7.76, roleW: 1.261, y: 5.048 },
  { name: 'Vxt6sn ifhd', role: 'Ui / X Developer', x: 10.52654, nameX: 11.067, nameW: 1.51, roleX: 11.191, roleW: 1.261, y: 5.048 },
];

function slide16(s) {
  bg(s, C.cream, 55.295);
  chrome(s);
  title(s, 0.12, 0.645, 1.864, 'The team.', { color: C.ink });
  hline(s, 0, 6.909, SLIDE_W, C.ruleMid);
  hline(s, 0, 1.48, SLIDE_W, C.ruleMid);
  TEAM.forEach((m) => {
    ellipse(s, m.x, 2.142, 2.59, 2.59, { line: { color: C.ruleMid, width: 1 } });
    txt(s, m.name, { x: m.nameX, y: m.y, w: m.nameW, h: 0.303, fontSize: 12, bold: true, color: C.black, align: 'center' });
    txt(s, m.role, { x: m.roleX, y: 5.318, w: m.roleW, h: 0.252, fontSize: 9, align: 'center' });
    txt(s, [{ text: 'To avoid striking it’s ', options: { breakLine: true } },
      { text: 'important decision makers to for. Design makers' }], {
      x: m.x + 0.14952, y: 5.53973, w: 2.29166, h: 0.67213, fontSize: 9, align: 'center', lineSpacingMultiple: 1.5555,
    });
    picFrame(s, m.x + 0.21606, 2.35817, 2.159, 2.159, 'ellipse');
  });
  [3.243, 6.675, 10.106].forEach((x) => vline(s, x, 1.48, 5.429, C.ruleMid));
}

/* 17 — client feedback */
function slide17(s) {
  bg(s, C.cream, 44.314);
  s.addShape('roundRect', {
    x: 0.221, y: 1.452, w: 5.633, h: 5.451, rectRadius: 0.353,
    fill: { type: 'none' }, line: { color: C.ruleMid, width: 1 },
  });
  chrome(s);
  title(s, 5.751, 0.613, 3.007, 'Client feedback.', { color: C.ink });
  navList(s, 1.04261, 2.70075, { color: C.ink });
  testimonial(s, 1.02866, 3.61412, C.ink);
  ribbon(s, 11.579, 1.902, C.ruleMid, { tabInk: C.ink, paraInk: C.ink });
  txt(s, LOREM1, {
    x: 3.544, y: 3.987, w: 5.494, h: 0.423, fontFace: F.mont, fontSize: 8, color: C.ink,
    align: 'center', lineSpacingMultiple: 1.5, rotate: -90,
  });
  picFrame(s, 6.745, 1.452, 4.344, 5.451, 'roundRect');
}

/* 18 — the sing off */
function slide18(s) {
  rect(s, 3.746, 0, 9.587, SLIDE_H, { fill: { color: C.sand, transparency: 80 } });
  hline(s, 5.854, 1.707, 7.237, C.ink, 0.75, 35.295);
  chrome(s);
  const dark = { color: C.ink, bodyFont: F.mont, bodyW: 3.026, bold: true };
  workItem(s, 5.769, 3.233, '01', dark);
  workItem(s, 5.769, 5.43, '03', dark);
  workItem(s, 10.06, 3.233, '02', { ...dark, bodyW: 2.982 });
  workItem(s, 10.06, 5.43, '04', { ...dark, bodyW: 2.982, numFont: F.inter });
  vline(s, 9.316, 3.296, 3.169, C.ink, 0.75, 35.295);
  vline(s, 9.056, 3.296, 3.169, C.ink, 0.75, 35.295);
  hline(s, 9.316, 4.947, 3.775, C.ink, 0.75, 35.295);
  txt(s, '10%', { x: 8.902, y: 4.737, w: 0.594, h: 0.286, fontSize: 11, bold: true, rotate: -90 });
  hline(s, 5.865, 4.947, 3.191, C.ink, 0.75, 35.295);
  title(s, 3.643, 0.619, 2.246, 'The sing off.', { color: C.ink });
  badge(s, 8.82, 1.31, '72%.', 0.19495);
  txt(s, MUST_PA, {
    x: 5.769, y: 2.29, w: 5.677, h: 0.486, fontFace: F.mont, fontSize: 10.5, lineSpacingMultiple: 1.3333,
  });
  picFrame(s, 0.221, 1.319, 5.305, 5.973);
}

/* 19 — the mood */
function slide19(s) {
  bg(s, C.cream, 75.295);
  chrome(s);
  txt(s, '02', { x: 6.985, y: 2.61, w: 0.612, h: 0.438, fontFace: F.mont, fontSize: 20, bold: true, color: C.black, align: 'center' });
  txt(s, '03', { x: 10.524, y: 2.61, w: 0.612, h: 0.438, fontSize: 20, color: C.black });
  txt(s, '02', { x: 7.78, y: 5.217, w: 0.612, h: 0.438, fontFace: F.mont, fontSize: 20, bold: true, color: C.black });
  txt(s, '03', { x: 10.651, y: 5.217, w: 0.612, h: 0.438, fontFace: F.mont, fontSize: 20, bold: true, color: C.black });
  title(s, 0.134, 0.752, 1.949, 'The mood.', { color: C.ink });
  hline(s, 3.068, 1.097, 10.044, C.ink, 0.75, 35.295);
  badge(s, 7.69333, 0.7, '25%', 0.2118);
  navList(s, 0.14775, 4.44475, { color: C.ink });
  txt(s, MUST, { x: 0.134, y: 5.358, w: 4.441, h: 0.486, fontSize: 10.5, lineSpacingMultiple: 1.3333 });
  txt(s, MUST_PA, { x: 0.134, y: 5.952, w: 2.4, h: 1.447, fontSize: 14, lineSpacingMultiple: 1.1428 });
  picFrame(s, 1.491, 1.727, 5.337, 2.545);
  picFrame(s, 4.794, 4.5, 2.962, 2.778);
  picFrame(s, 7.757, 1.727, 5.19, 2.545);
  picFrame(s, 8.543, 4.507, 1.98, 2.202);
}

/* 20 — thank you / contact */
function slide20(s) {
  bg(s, C.cream, 75.295);
  picFrame(s, 0.242, 1.508, 6.945, 5.764);
  chrome(s);
  [['Email', 4.027, 0.694, 'info@yourname.com', 4.196, 1.701, F.oswald, 0.286],
    ['Phone', 4.691, 0.768, 'Phone : (100) 3345 789', 4.847, 1.801, F.oswald, 0.337],
    ['Location', 5.393, 0.926, 'New York. USA', 5.587, 1.324, F.pop, 0.286],
  ].forEach(([label, ly, lw, value, vy, vw, font, vh]) => {
    txt(s, label, { x: 7.662, y: ly, w: lw, h: 0.269, bold: true, color: C.black });
    txt(s, value, { x: 7.662, y: vy, w: vw, h: vh, fontFace: font, fontSize: 8, color: C.black, lineSpacingMultiple: 1.5 });
  });
  txt(s, 'Contact', { x: 7.621, y: 2.998, w: 2.461, h: 0.69, fontSize: 35, color: C.black });
  ellipse(s, 11.613, 2.996, 0.805, 0.805, { line: { color: C.gold, width: 2.75 } });
  s.addShape('chevron', {
    x: 11.811, y: 3.194, w: 0.409, h: 0.409, fill: { color: C.gold }, rectRadius: 0.2636, rotate: -45,
  });
  txt(s, [{ text: 'Thank', options: { color: C.sand } }, { text: ' you.', options: { color: C.ink } }], {
    x: 4.417, y: 1.193, w: 4.753, h: 1.616, fontFace: F.bebas, fontSize: 90, bold: true,
  });
}

/* ═════════════════════════════ build ═════════════════════════════════ */
const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'DECK';
  pptx.author = 'HUGE.';
  pptx.title = 'Brand Proposal 2027';
  BUILDERS.forEach((buildSlide) => buildSlide(pptx.addSlide()));
  return pptx;
}

build().writeFile({
  fileName: path.join(__dirname, '056f3379-5c42-4bf9-b2d5-91912f2d9852_grok_final.pptx'),
});
