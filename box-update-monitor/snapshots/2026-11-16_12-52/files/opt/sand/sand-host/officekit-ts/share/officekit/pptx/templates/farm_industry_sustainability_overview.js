/**
 * "Farm Industry" presentation template - rebuilt with pptxgenjs.
 *
 * Run:  node 00918b64-d487-4412-8c45-d6b8f79ed2a3_grok_final.js
 * Out:  00918b64-d487-4412-8c45-d6b8f79ed2a3_grok_final.pptx (next to this file)
 *
 * Raster artwork of the original deck (the laptop photo and the small icon
 * PNGs) is replaced by native pptxgenjs shapes standing in at the same
 * position and size. The original also carries a number of *empty* picture
 * placeholders inherited from its layouts; those hold no artwork and render as
 * nothing, so they are not reproduced.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const C = {
  dark: '122204', // accent2 - near-black green, deck background
  lime: 'D1FC55', // accent1 - signature lime
  limeDeep: 'BAF804', // accent1 lumMod 75%
  white: 'FFFFFF',
  black: '000000',
  grey: '404040', // tx1 lumMod 75% / lumOff 25%
  yellow: 'FFFF00', // accent3
  orange: 'FFC000', // accent4
  cream: 'FFF2CC', // accent4 lumMod 20% / lumOff 80%
  gridline: '626D59', // white @ 34% over the dark chart plate
  frame: '3A3B40', // laptop mockup body
  keys: '4A4B52',
};

const FONT = 'Onest';
const SLIDE_W = 13.3333;
const SLIDE_H = 7.5;

const pres = new PptxGenJS();
pres.defineLayout({ name: 'DECK', width: SLIDE_W, height: SLIDE_H });
pres.layout = 'DECK';
pres.author = 'Farm Industry';
pres.title = 'Farm Industry - Presentation Template';

const S = pres.ShapeType;

/* ---------------------------------------------------------------- helpers */

/** Solid shape. `o` carries x/y/w/h plus any extra pptxgenjs options. */
function shape(slide, type, fill, o) {
  slide.addShape(type, Object.assign({ fill: fill ? { color: fill } : { type: 'none' } }, o));
}

/** Text box. `runs` is a string or an array of run objects. */
function text(slide, runs, o) {
  slide.addText(runs, Object.assign({ fontFace: FONT, fontSize: 14, color: C.white, valign: 'top' }, o));
}

/** The three decorative dots that appear on almost every slide. */
function dots(slide, x, y, colors, size = 0.365, gap = 0.5157) {
  colors.forEach((c, i) => shape(slide, S.ellipse, c, { x: x + i * gap, y, w: size, h: size }));
}

/** Rotated page-number box (bottom-right corner). */
function pageNumber(slide, n, color, x = 11.349, y = 6.74) {
  text(slide, String(n), {
    x, y, w: 1.282, h: 0.37, color, fontSize: 16,
    align: 'right', rotate: 180, flipV: true, lineSpacingMultiple: 1,
  });
}

/** Grey page number inherited from the slide master. */
function masterPageNumber(slide, n) {
  pageNumber(slide, n, C.grey, 11.345, 6.738);
}

/** Eyebrow line + big headline, as used on the section-opener slides. */
function headline(slide, o) {
  const runs = [];
  if (o.eyebrow) runs.push({ text: o.eyebrow, options: { fontSize: o.eyebrowSize || 20, bold: true, breakLine: true } });
  runs.push({ text: o.title, options: { fontSize: o.titleSize || 44, bold: true } });
  text(slide, runs, {
    x: o.x, y: o.y, w: o.w, h: o.h, color: o.color, align: o.align || 'left',
    lineSpacingMultiple: 1.15, paraSpaceAfter: 8,
  });
}

/** Title placeholder styling: 44pt bold, 90% leading, vertically centred. */
function title(slide, str, o) {
  text(slide, str, Object.assign({ fontSize: 44, bold: true, valign: 'middle', lineSpacingMultiple: 0.9 }, o));
}

/** Rounded rectangle with an explicit corner radius. */
function rounded(slide, x, y, w, h, r, fill) {
  slide.addShape(S.roundRect, { x, y, w, h, rectRadius: r, fill: { color: fill } });
}

/** 12-petal lime flower icon used as a bullet marker (was a PNG). */
function flower(slide, x, y, size) {
  shape(slide, S.star12, C.lime, { x, y, w: size, h: size });
}

/** Outlined document icon (certificate) - stand-in for a line-art PNG. */
function documentIcon(slide, x, y, size) {
  const s = size;
  slide.addShape(S.roundRect, {
    x, y, w: s * 0.82, h: s, rectRadius: s * 0.1,
    fill: { type: 'none' }, line: { color: C.white, width: 1.25 },
  });
  [0.24, 0.42, 0.6].forEach(f => shape(slide, S.rect, C.white, { x: x + s * 0.14, y: y + s * f, w: s * 0.24, h: s * 0.06 }));
  shape(slide, S.ellipse, C.white, { x: x + s * 0.46, y: y + s * 0.22, w: s * 0.26, h: s * 0.26 });
}

/** Outlined calendar icon - stand-in for a line-art PNG. */
function calendarIcon(slide, x, y, size) {
  const s = size;
  slide.addShape(S.roundRect, {
    x, y: y + s * 0.14, w: s, h: s * 0.86, rectRadius: s * 0.1,
    fill: { type: 'none' }, line: { color: C.white, width: 1.25 },
  });
  [0.22, 0.66].forEach(f => shape(slide, S.rect, C.white, { x: x + s * f, y, w: s * 0.12, h: s * 0.24 }));
  shape(slide, S.rect, C.white, { x, y: y + s * 0.42, w: s, h: s * 0.06 });
  [0.16, 0.44, 0.72].forEach(fx => [0.56, 0.74].forEach(fy =>
    shape(slide, S.rect, C.white, { x: x + s * fx, y: y + s * fy, w: s * 0.13, h: s * 0.11 })));
}

const LOREM_LONG =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ' +
  'ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis';
const LOREM_MED = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor';

/* ------------------------------------------------------------- 1. cover */

function slide01() {
  const s = pres.addSlide();
  s.background = { color: C.dark };
  // lime panel on the right (a 7.5 x 3.71 rectangle rotated 90 degrees)
  shape(s, S.rect, C.lime, { x: 7.728, y: 1.895, w: 7.5, h: 3.71, rotate: 90 });

  text(s, [
    { text: 'Farm', options: { fontSize: 115 } },
    { text: ' Industry', options: { fontSize: 88 } },
  ], { x: 0.577, y: 0.824, w: 6.516, h: 3.517, color: C.lime });

  text(s, 'Presentation Template ', { x: 0.653, y: 4.595, w: 5.906, h: 0.572, fontSize: 28, color: C.lime });
  text(s, LOREM_LONG + ' nostrud', { x: 0.653, y: 5.313, w: 5.906, h: 1.161, lineSpacingMultiple: 1.5 });
  dots(s, 5.002, 4.692, [C.lime, C.lime, C.lime]);
}

/* --------------------------------------------------- 2. table of contents */

function slide02() {
  const s = pres.addSlide();
  s.background = { color: C.dark };
  shape(s, S.rect, C.lime, { x: 0, y: 0, w: 3.57, h: 7.5 });

  title(s, 'Table  Of Contents', { x: 6.771, y: 0.757, w: 4.728, h: 2.322, color: C.lime });
  dots(s, 11.123, 0.845, [C.lime, C.lime, C.lime]);

  const items = [
    ['Introduction ', 3.052, 3.046],
    ['Key Concepts ', 3.735, 3.046],
    ['Mission & Vision ', 4.419, 3.046],
    ['Farm Industry Strategies', 5.102, 4.664],
    ['Farm Industry Outcomes', 5.785, 4.148],
  ];
  items.forEach(([label, y, w]) => {
    flower(s, 6.886, y + 0.07, 0.365);
    text(s, label, { x: 7.34, y, w, h: 0.505, fontSize: 24 });
  });
  masterPageNumber(s, 2);
}

/* ------------------------------------------ 3. challenges and opportunities */

function slide03() {
  const s = pres.addSlide();
  s.background = { color: C.dark };
  shape(s, S.rect, C.lime, { x: 0, y: 0, w: 3.407, h: 7.5 });
  dots(s, 11.216, 0.875, [C.lime, C.lime, C.lime]);

  headline(s, {
    x: 5.722, y: 1.073, w: 6.708, h: 2.268, color: C.lime,
    eyebrow: 'Sustainability In Farming', title: 'Challenges And Opportunities',
  });
  text(s, 'Your Text Here', { x: 5.769, y: 3.697, w: 4.243, h: 0.572, fontSize: 28 });
  text(s, LOREM_LONG, { x: 5.769, y: 4.428, w: 4.243, h: 1.515, lineSpacingMultiple: 1.5 });

  [['$142.3', 4.255, 2.398], ['90%', 5.372, 2.002]].forEach(([label, y, w]) => {
    flower(s, 10.405, y + 0.055, 0.489);
    text(s, label, { x: 11.02, y, w, h: 0.655, fontSize: 36, lineSpacingMultiple: 0.9 });
  });
  masterPageNumber(s, 3);
}

/* ------------------------------------------------------- 4. innovations */

function slide04() {
  const s = pres.addSlide();
  s.background = { color: C.dark };
  shape(s, S.rect, C.lime, { x: 7.982, y: 2.15, w: 7.507, h: 3.193, rotate: 90 });
  dots(s, 5.98, 1.05, [C.lime, C.lime, C.lime]);

  headline(s, {
    x: 0.616, y: 1.182, w: 6.667, h: 2.268, color: C.lime,
    eyebrow: 'The Future Of Farming', title: 'Innovations In The Agricultural Industry',
  });

  [0.742, 4.149].forEach((x, i) => {
    (i === 0 ? documentIcon : calendarIcon)(s, [2.012, 5.418][i], 4.205, 0.472);
    text(s, 'Your Text', { x, y: 4.893, w: 3.011, h: 0.532, fontSize: 28, align: 'center', lineSpacingMultiple: 0.9 });
    text(s, 'Lorem ipsum dolor sit amet, consecte adipiscing elit. ',
      { x, y: 5.485, w: 3.011, h: 0.761, align: 'center', lineSpacingMultiple: 1.4 });
  });
  pageNumber(s, 4, C.grey);
}

/* -------------------------------------------------- 5. sustainable farming */

function slide05() {
  const s = pres.addSlide();
  s.background = { color: C.dark };
  dots(s, 11.239, 0.576, [C.lime, C.lime, C.lime]);

  text(s, 'Sustainable Farming Practices For A Healthier Future', {
    x: 1.557, y: 0.955, w: 10.27, h: 1.768, fontSize: 44, bold: true, color: C.lime,
    align: 'center', lineSpacingMultiple: 1.15, paraSpaceAfter: 8,
  });

  const cards = [
    { x: 1.467, r: 0.268, num: '01.', numX: 1.791, numY: 3.073, hdX: 2.672, hdY: 3.365, hdW: 2.737, hdH: 0.485, bodyX: 1.867, bodyW: 4.397 },
    { x: 7.21, r: 0.256, num: '02.', numX: 7.492, numY: 3.036, hdX: 8.495, hdY: 3.385, hdW: 2.375, hdH: 0.598, bodyX: 7.523, bodyW: 4.595 },
  ];
  cards.forEach(c => {
    rounded(s, c.x, 3.073, 4.998, 2.266, c.r, C.lime);
    text(s, c.num, { x: c.numX, y: c.numY, w: 1.296, h: 0.875, fontSize: 40, color: C.dark, lineSpacingMultiple: 1.3 });
    text(s, 'Your Text Here', { x: c.hdX, y: c.hdY, w: c.hdW, h: c.hdH, fontSize: 20, color: C.dark, lineSpacingMultiple: 1.2 });
    text(s, LOREM_SHORT + ' incididunt ut labore et dolore magna aliqua. ',
      { x: c.bodyX, y: 4.047, w: c.bodyW, h: 1.292, color: C.black, lineSpacingMultiple: 1.3 });
  });
  masterPageNumber(s, 5);
  pageNumber(s, 5, C.white);
}


/* ---------------------------------------------------------- 6. a new era */

function slide06() {
  const s = pres.addSlide();
  s.background = { color: C.dark };
  shape(s, S.rect, C.lime, { x: 0, y: 4.215, w: 13.333, h: 2.454 });
  dots(s, 11.154, 0.865, [C.yellow, C.orange, C.lime]);

  headline(s, {
    x: 0.682, y: 0.819, w: 9.925, h: 2.229, color: C.lime, eyebrowSize: 18,
    eyebrow: 'Revolutionizing Agriculture', title: 'A New Era In Farm Industry Practices',
  });
  text(s, LOREM_LONG, { x: 0.72, y: 3.207, w: 8.45, h: 0.808, lineSpacingMultiple: 1.5 });

  const cols = [
    { n: '01. ', x: 0.868, hy: 4.586, hw: 3.145, hh: 0.655, by: 5.243 },
    { n: '02. ', x: 4.86, hy: 4.573, hw: 3.463, hh: 0.655, by: 5.18 },
    { n: '03. ', x: 9.169, hy: 4.573, hw: 3.877, hh: 0.797, by: 5.243 },
  ];
  cols.forEach(c => {
    text(s, c.n + 'Your Text Here',
      { x: c.x, y: c.hy, w: c.hw, h: c.hh, fontSize: 24, color: C.dark, lineSpacingMultiple: 1 });
    text(s, LOREM_SHORT,
      { x: c.x, y: c.by, w: 2.996, h: 0.978, color: C.black, lineSpacingMultiple: 1.3 });
  });
  masterPageNumber(s, 6);
  pageNumber(s, 6, C.white);
}

/* ------------------------------------------------------- 7. key insights */

function slide07() {
  const s = pres.addSlide();
  s.background = { color: C.lime };
  dots(s, 11.104, 0.91, [C.dark, C.dark, C.dark]);

  const tiles = [
    { x: 0.711, y: 1.835, tx: 0.828, ty: 2.374, label: 'Traditional Tools' },
    { x: 3.645, y: 1.835, tx: 3.77, ty: 2.342, label: 'Formative Assessments' },
    { x: 0.687, y: 3.849, tx: 0.805, ty: 4.314, label: 'Rubrics and Checklists' },
    { x: 3.622, y: 3.849, tx: 3.739, ty: 4.318, label: 'Technology-Based Tools' },
  ];
  tiles.forEach(t => {
    shape(s, S.roundRect, C.dark, { x: t.x, y: t.y, w: 2.725, h: 1.776 });
    text(s, t.label, { x: t.tx, y: t.ty, w: 2.49, h: 0.838, fontSize: 24, align: 'center', lineSpacingMultiple: 0.9 });
  });

  headline(s, {
    x: 6.922, y: 1.62, w: 6.667, h: 2.268, color: C.dark,
    eyebrow: 'The State Of Farming', title: 'Key Insights And Future Directions',
  });
  text(s, LOREM_LONG, { x: 6.876, y: 4.308, w: 5.26, h: 1.161, color: C.grey, lineSpacingMultiple: 1.5 });
  pageNumber(s, 7, C.grey);
}

/* --------------------------------------------------------- 8. efficiency */

function slide08() {
  const s = pres.addSlide();
  s.background = { color: C.dark };
  shape(s, S.rect, C.lime, { x: 0, y: -0.01, w: 13.333, h: 4.528 });
  dots(s, 11.225, 0.563, [C.dark, C.dark, C.dark]);

  headline(s, {
    x: 1.392, y: 0.928, w: 10.581, h: 2.268, color: C.dark, align: 'center',
    eyebrow: 'Optimizing Farm Operations', title: 'Efficiency And Productivity In Farming',
  });

  const cards = [
    { x: 3.651, y: 3.488, fill: C.dark, num: '01.', tx: 3.836, color: C.white },
    { x: 9.52, y: 3.508, fill: C.limeDeep, num: '02.', tx: 9.706, color: C.grey },
  ];
  cards.forEach(c => {
    rounded(s, c.x, c.y, 2.505, 2.799, 0.234, c.fill);
    text(s, c.num, { x: c.tx, y: 3.895, w: 2.134, h: 0.685, fontSize: 40, color: c.color, align: 'center', lineSpacingMultiple: 1 });
    text(s, [
      { text: 'Lorem ipsum ', options: { breakLine: true } },
      { text: 'dolor sit amet, consectetur adipisicing elit, sed' },
    ], { x: c.tx, y: 4.568, w: 2.134, h: 0.978, color: c.color, align: 'center', lineSpacingMultiple: 1.3 });
  });
  masterPageNumber(s, 8);
}

/* ------------------------------------------------------------- 9. profile */

function slide09() {
  const s = pres.addSlide();
  s.background = { color: C.dark };
  title(s, 'Alexander Christie', { x: 1.116, y: 1.455, w: 5.551, h: 1.416, color: C.lime });
  dots(s, 5.578, 2.506, [C.lime, C.lime, C.lime]);

  rounded(s, 0.736, 3.254, 10.795, 3.328, 0.452, C.lime);
  text(s, 'Farmer Leader', { x: 1.262, y: 3.503, w: 3.46, h: 0.591, fontSize: 24, color: C.dark, lineSpacingMultiple: 1.3 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam',
    { x: 1.26, y: 4.104, w: 5.551, h: 1.002, color: C.dark, lineSpacingMultiple: 1.3 });

  const bars = [
    { label: 'Experience', x: 1.262, trackX: 1.342, fillX: 1.373, fillW: 2.058, knobX: 3.048, capX: 2.995, pct: '87%' },
    { label: 'Skils', x: 4.12, trackX: 4.2, fillX: 4.232, fillW: 1.496, knobX: 5.434, capX: 5.378, pct: '50%' },
  ];
  bars.forEach(b => {
    text(s, b.label, { x: b.x, y: 5.388, w: 1.273, h: 0.39, color: C.dark, lineSpacingMultiple: 1.3 });
    s.addShape(S.roundRect, { x: b.trackX, y: 5.873, w: 2.501, h: 0.184, rectRadius: 0.092, fill: { color: C.cream, transparency: 50 } });
    rounded(s, b.fillX, 5.895, b.fillW, 0.127, 0.0635, C.dark);
    shape(s, S.ellipse, C.dark, { x: b.knobX, y: 5.755, w: 0.413, h: 0.413 });
    text(s, b.pct, { x: b.capX, y: 5.8, w: 0.527, h: 0.317, fontSize: 10.5, align: 'center', lineSpacingMultiple: 1.3 });
  });
  masterPageNumber(s, 9);
}

/* ----------------------------------------------------------- 10. our team */

function slide10() {
  const s = pres.addSlide();
  s.background = { color: C.dark };
  shape(s, S.rect, C.lime, { x: 4.958, y: -0.692, w: 3.417, h: 13.333, rotate: 270, flipH: true });
  title(s, 'Our Best Team', { x: 0.851, y: 0.928, w: 11.631, h: 0.968, color: C.lime, align: 'center' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad',
    { x: 2.861, y: 2.085, w: 7.611, h: 1.126, align: 'center', lineSpacingMultiple: 1.3 });

  const team = [
    { name: 'Alex  Paula', x: 0.767, ph: 0.682, y: 5.765 },
    { name: 'James Fetcher', x: 3.983, ph: 3.837, y: 5.765 },
    { name: 'Karina', x: 7.137, ph: 6.992, y: 5.757 },
    { name: 'Tifanny', x: 10.292, ph: 10.146, y: 5.757 },
  ];
  team.forEach(m => {
    text(s, m.name, { x: m.x, y: m.y, w: 2.213, h: 0.438, fontSize: 20, color: C.dark, align: 'center' });
    text(s, 'Job Position', { x: m.x, y: m.y + 0.352, w: 2.213, h: 0.337, color: C.dark, align: 'center' });
  });
  // decorative dot trio parked off-canvas in the original file
  dots(s, -0.447, -4.104, [C.yellow, C.lime, C.orange]);
  masterPageNumber(s, 10);
  pageNumber(s, 10, C.white);
}

/* ------------------------------------------------------------ 11. pricing */

function slide11() {
  const s = pres.addSlide();
  s.background = { color: C.lime };
  title(s, 'Pricing Packages', { x: 2.157, y: 0.967, w: 9.057, h: 1.219, color: C.dark, align: 'center' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad',
    { x: 2.861, y: 2.085, w: 7.611, h: 1.126, color: C.grey, align: 'center', lineSpacingMultiple: 1.3 });
  dots(s, 0.741, 0.837, [C.dark, C.dark, C.dark]);

  const plans = [
    { card: 0.858, h: 3.48, name: 'Package 01', nameX: 1.494, nameY: 3.566, price: '$25.9', priceX: 1.143, priceY: 4.307, permX: 2.707, permY: 4.429, listX: 1.104, listY: 5.103 },
    { card: 4.962, h: 3.48, name: 'Package 02', nameX: 5.638, nameY: 3.546, price: '$45.9', priceX: 5.272, priceY: 4.307, permX: 6.836, permY: 4.429, listX: 5.208, listY: 5.103 },
    { card: 9.084, h: 3.487, name: 'Package 03', nameX: 9.773, nameY: 3.54, price: '$55.9', priceX: 9.396, priceY: 4.277, permX: 10.96, permY: 4.333, listX: 9.33, listY: 5.11 },
  ];
  const perks = ['Lorem ipsum dolor sit', 'Consectetur adipiscing', 'Sed do eiusmod tempor ', 'Ut labore et dolore'];
  plans.forEach(p => {
    rounded(s, p.card, 3.267, 3.409, p.h, 0.393, C.dark);
    text(s, p.name, { x: p.nameX, y: p.nameY, w: 1.996, h: 0.438, fontSize: 20, align: 'center' });
    text(s, p.price, { x: p.priceX, y: p.priceY, w: 1.756, h: 0.609, fontSize: 36, bold: true, color: C.lime, align: 'center', lineSpacingMultiple: 0.8 });
    text(s, '/Month', { x: p.permX, y: p.permY, w: 1.221, h: 0.404, fontSize: 20, color: C.lime, lineSpacingMultiple: 0.9 });
    text(s, perks.map(t => ({ text: t, options: { bullet: { characterCode: '2022', indent: 22.5 } } })),
      { x: p.listX, y: p.listY, w: 2.917, h: 1.326, lineSpacingMultiple: 1.3 });
  });
  pageNumber(s, 11, C.grey);
}

/* -------------------------------------------------------------- 12. break */

function slide12() {
  const s = pres.addSlide();
  s.background = { color: C.dark };
  rounded(s, 0.682, 0.757, 11.969, 5.986, 0.431, C.lime);
  text(s, 'Break Slide', {
    x: 1.98, y: 1.575, w: 9.353, h: 1.559, fontSize: 96, bold: true, color: C.dark,
    align: 'center', valign: 'middle', lineSpacingMultiple: 0.9,
  });
  text(s, '25+ Minutes Break Time', { x: 4.785, y: 3.429, w: 3.763, h: 0.563, fontSize: 20, color: C.dark, align: 'center', lineSpacingMultiple: 1.3 });
  dots(s, 5.968, 4.309, [C.dark, C.dark, C.dark]);
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor',
    { x: 3.915, y: 5.101, w: 5.504, h: 0.727, color: C.black, align: 'center', lineSpacingMultiple: 1.3 });
  masterPageNumber(s, 12);
}

/* ------------------------------------------------------ 13. global economy */

function slide13() {
  const s = pres.addSlide();
  s.background = { color: C.dark };
  dots(s, 11.254, 0.571, [C.lime, C.lime, C.lime]);
  text(s, 'The Role Of The Farm Industry In The Global Economy', {
    x: 1.537, y: 1.103, w: 10.25, h: 1.768, fontSize: 44, bold: true, color: C.lime,
    align: 'center', lineSpacingMultiple: 1.15, paraSpaceAfter: 8,
  });

  const cells = [
    { num: '01.', x: 0.927, y: 3.258, numX: 1.233, numY: 3.55, numW: 1.511, txX: 2.925, txY: 3.522 },
    { num: '02.', x: 0.927, y: 5.149, numX: 1.233, numY: 5.441, numW: 1.692, txX: 2.925, txY: 5.414 },
    { num: '03.', x: 6.881, y: 3.258, numX: 7.187, numY: 3.55, numW: 1.692, txX: 8.879, txY: 3.522 },
    { num: '04.', x: 6.881, y: 5.149, numX: 7.187, numY: 5.441, numW: 1.692, txX: 8.879, txY: 5.414 },
  ];
  cells.forEach(c => {
    shape(s, S.roundRect, C.lime, { x: c.x, y: c.y, w: 5.552, h: 1.538 });
    text(s, c.num, { x: c.numX, y: c.numY, w: c.numW, h: 1.078, fontSize: 66, color: C.black, valign: 'middle', lineSpacingMultiple: 0.9 });
    text(s, LOREM_MED, { x: c.txX, y: c.txY, w: 3.331, h: 1.132, color: C.black, lineSpacingMultiple: 1.3 });
  });
  masterPageNumber(s, 13);
}

/* ---------------------------------------------------- 14. changing demands */

function slide14() {
  const s = pres.addSlide();
  s.background = { color: C.dark };
  dots(s, 6.972, 0.577, [C.lime, C.lime, C.lime]);

  text(s, 'How The Farm Industry Is Adapting To Changing Demands',
    { x: 0.694, y: 1.018, w: 6.792, h: 3.332, fontSize: 48, bold: true, color: C.lime });
  text(s, '- By Michael Vendorven', { x: 0.758, y: 4.4, w: 3.121, h: 0.478, fontSize: 16, lineSpacingMultiple: 1.4 });

  shape(s, S.roundRect, C.lime, { x: 3.65, y: 5.153, w: 5.17, h: 1.439 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna. ',
    { x: 3.964, y: 5.35, w: 4.686, h: 1.132, color: C.grey, lineSpacingMultiple: 1.3 });
  masterPageNumber(s, 14);
  pageNumber(s, 14, C.white);
}

/* ------------------------------------------------------ 15. image gallery */

function slide15() {
  const s = pres.addSlide();
  s.background = { color: C.dark };
  title(s, 'Image Gallery', { x: 5.476, y: 1.047, w: 4.964, h: 1.208, color: C.lime });
  dots(s, 11.112, 0.898, [C.lime, C.lime, C.lime]);
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore',
    { x: 5.476, y: 2.345, w: 5.964, h: 0.727, lineSpacingMultiple: 1.3 });
  pageNumber(s, 15, C.yellow, 11.656, 6.833);
  masterPageNumber(s, 15);
}

/* ---------------------------------------------------------- 16. timeline */

function slide16() {
  const s = pres.addSlide();
  s.background = { color: C.dark };
  shape(s, S.rect, C.lime, { x: 0, y: 3.75, w: 13.333, h: 2.617 });
  title(s, 'Timeline Infographic', { x: 1.028, y: 1.18, w: 11.338, h: 1.007, color: C.lime, align: 'center' });
  dots(s, 11.153, 0.876, [C.yellow, C.lime, C.dark]);
  text(s, LOREM_LONG, { x: 2.488, y: 2.243, w: 8.357, h: 0.768, align: 'center', lineSpacingMultiple: 1.5 });

  const steps = [
    { year: '2024', pillX: 0.953, pillY: 3.466, textX: 0.81, arrowX: 2.742, arrowFill: C.lime },
    { year: '2025', pillX: 4.0, pillY: 3.466, textX: 3.911, arrowX: 5.795, arrowFill: C.lime },
    { year: '2026', pillX: 7.011, pillY: 3.465, textX: 7.011, arrowX: 8.807, arrowFill: C.limeDeep },
    { year: '2027', pillX: 10.017, pillY: 3.465, textX: 9.933 },
  ];
  steps.forEach(st => {
    if (st.arrowX) shape(s, S.rightArrow, st.arrowFill, { x: st.arrowX, y: 3.611, w: 1.136, h: 0.247 });
    s.addText(st.year, {
      shape: S.roundRect, rectRadius: 0.2575, fill: { color: C.limeDeep },
      x: st.pillX, y: st.pillY, w: 1.709, h: 0.515,
      fontFace: FONT, fontSize: 24, color: C.dark, align: 'center', valign: 'middle',
    });
    text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore',
      { x: st.textX, y: 4.268, w: 2.363, h: 1.697, color: C.grey, lineSpacingMultiple: 1.3 });
  });
  masterPageNumber(s, 16);
}

/* --------------------------------------------------------- 17. bar chart */

function slide17() {
  const s = pres.addSlide();
  s.background = { color: C.dark };
  dots(s, 0.803, 0.875, [C.lime, C.lime, C.lime]);

  const bars = [
    { x: 0.708, y: 4.114, h: 2.215, pct: '35%', labelX: 0.939, labelY: 5.276, capX: 1.163, cap: 'Chart 1' },
    { x: 3.194, y: 2.144, h: 4.185, pct: '60%', labelX: 3.425, labelY: 5.276, capX: 3.649, cap: 'Chart 2' },
    { x: 5.681, y: 0.925, h: 5.403, pct: '95%', labelX: 5.911, labelY: 5.275, capX: 6.135, cap: 'Chart 3' },
  ];
  bars.forEach(b => {
    shape(s, S.roundRect, C.lime, { x: b.x, y: b.y, w: 1.972, h: b.h });
    text(s, b.pct, { x: b.labelX, y: b.labelY, w: 1.511, h: 1.078, fontSize: 44, color: C.dark, valign: 'middle', lineSpacingMultiple: 0.9 });
    text(s, b.cap, { x: b.capX, y: 6.328, w: 1.062, h: 0.415, align: 'center', lineSpacingMultiple: 1.5 });
  });

  title(s, 'Bar Chart Slide', { x: 8.098, y: 1.877, w: 4.553, h: 1.208, color: C.lime });
  text(s, 'Your Text Here', { x: 8.104, y: 3.137, w: 3.568, h: 1.078, fontSize: 28, valign: 'middle', lineSpacingMultiple: 0.9 });
  text(s, LOREM_LONG, { x: 8.104, y: 3.918, w: 3.568, h: 1.828, lineSpacingMultiple: 1.5 });
  masterPageNumber(s, 17);
}

/* ------------------------------------------------------ 18. office chart */

function slide18() {
  const s = pres.addSlide();
  s.background = { color: C.lime };
  rounded(s, 0.682, 2.367, 6.434, 4.043, 0.307, C.dark);
  title(s, 'Office Chart Slide', { x: 1.742, y: 0.935, w: 9.823, h: 1.415, color: C.dark, align: 'center' });
  dots(s, 11.176, 0.71, [C.dark, C.dark, C.dark]);

  // Smooth-marker scatter over the twelve months of Sheet1 in the original workbook.
  s.addChart(pres.ChartType.scatter, [
    { name: 'Month', values: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12] },
    { name: 'Series 1', values: [4.3, 2.5, 3.5, 4.5, 2.4, 4.4, 1.8, 2.8, 2, 2, 3, 2] },
    { name: 'Series 2', values: [2.4, 1.3, 4.6, 6, 3.5, 2.1, 4.2, 1, 3, 4, 1.3, 3] },
  ], {
    x: 1.298, y: 2.694, w: 5.423, h: 3.496,
    layout: { x: 0.0555, y: 0.0569, w: 0.9344, h: 0.803 },
    lineSmooth: true, lineSize: 0.75, lineDataSymbolSize: 5,
    chartColors: [C.lime, C.yellow],
    showLegend: false, showTitle: false,
    catAxisLabelColor: C.white, catAxisLabelFontSize: 11, catAxisLabelFontFace: FONT,
    valAxisLabelColor: C.white, valAxisLabelFontSize: 12, valAxisLabelFontFace: FONT,
    catGridLine: { style: 'none' },
    valGridLine: { color: C.gridline, size: 0.75, style: 'solid' },
    catAxisLineShow: false, valAxisLineShow: false,
    plotArea: { fill: { color: C.dark } }, chartArea: { fill: { color: C.dark } },
  });

  const stats = [
    { value: '99.000+', x: 7.517, y: 2.367, bodyX: 7.517, bodyY: 3.148, bodyW: 4.175 },
    { value: '76.000+', x: 8.658, y: 4.226, bodyX: 8.658, bodyY: 5.007, bodyW: 3.85 },
  ];
  stats.forEach(st => {
    text(s, st.value, { x: st.x, y: st.y, w: 3.375, h: 1.078, fontSize: 28, color: C.dark, valign: 'middle', lineSpacingMultiple: 0.9 });
    text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt',
      { x: st.bodyX, y: st.bodyY, w: st.bodyW, h: 1.121, color: C.black, lineSpacingMultiple: 1.5 });
  });
  masterPageNumber(s, 18);
}

/* ----------------------------------------------------- 19. laptop mockup */

function slide19() {
  const s = pres.addSlide();
  s.background = { color: C.lime };
  laptopMockup(s);
  title(s, 'Laptop Mockup', { x: 6.667, y: 1.21, w: 4.621, h: 2.04, color: C.dark });
  dots(s, 11.109, 1.028, [C.orange, C.yellow, C.dark]);
  text(s, LOREM_LONG, { x: 7.15, y: 3.274, w: 4.243, h: 1.515, color: C.black, lineSpacingMultiple: 1.5 });
  pageNumber(s, 19, C.white);
}

/** Stand-in for the tablet-with-keyboard photograph on slide 19. */
function laptopMockup(s) {
  s.addShape(S.parallelogram, { x: 1.43, y: 4.55, w: 5.6, h: 1.75, fill: { color: C.keys }, flipH: true });
  s.addShape(S.roundRect, { x: 1.2, y: 0.72, w: 4.85, h: 3.85, rectRadius: 0.14, rotate: -2, fill: { color: C.frame } });
  s.addShape(S.roundRect, { x: 1.37, y: 0.89, w: 4.51, h: 3.51, rectRadius: 0.05, rotate: -2, fill: { color: C.lime } });
  text(s, '[image]', { x: 1.37, y: 2.45, w: 4.51, h: 0.4, align: 'center', fontSize: 12, color: C.dark });
}

/* --------------------------------------------------------- 20. thank you */

function slide20() {
  const s = pres.addSlide();
  s.background = { color: C.lime };
  s.addShape(S.roundRect, { x: 0.682, y: 0.757, w: 11.969, h: 5.986, rectRadius: 0.181, fill: { color: C.dark, transparency: 10 } });
  dots(s, 5.962, 1.356, [C.lime, C.lime, C.lime]);

  text(s, 'Thank You', {
    x: 2.267, y: 1.918, w: 8.788, h: 2.243, fontSize: 115, bold: true, color: C.lime,
    align: 'center', valign: 'middle', lineSpacingMultiple: 0.9,
  });
  s.addShape(S.line, {
    x: 3.773, y: 4.261, w: 5.787, h: 0,
    line: { color: C.yellow, width: 1, beginArrowType: 'oval', endArrowType: 'oval' },
  });

  const contacts = [
    { label: 'Address', value: '253 Mutton Town Road', x: 2.437, w: 2.762 },
    { label: 'Phone', value: '+123 456 7890', x: 5.665, w: 2.002 },
    { label: 'Website', value: 'www.website.com', x: 8.119, w: 2.777 },
  ];
  contacts.forEach(c => {
    text(s, c.label, { x: c.x, y: 4.969, w: c.w, h: 0.599, fontSize: 32, align: 'center', lineSpacingMultiple: 0.9 });
    text(s, c.value, { x: c.x, y: 5.547, w: c.w, h: 0.407, align: 'center', lineSpacingMultiple: 1.3 });
  });
  pageNumber(s, 20, C.grey);
}

/* -------------------------------------------------------------- assemble */

[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
  .forEach(build => build());

pres.writeFile({ fileName: path.join(__dirname, '00918b64-d487-4412-8c45-d6b8f79ed2a3_grok_final.pptx') })
  .then(f => console.log('wrote', f));
