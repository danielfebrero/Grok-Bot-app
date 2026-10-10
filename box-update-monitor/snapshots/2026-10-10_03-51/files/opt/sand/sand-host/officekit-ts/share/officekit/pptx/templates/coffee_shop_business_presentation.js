/**
 * "Coffee Shop" deck — rebuilt with pptxgenjs.
 * Run: node 018eeb0c-98be-466b-952d-ee27ab314be0_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Theme
 * ------------------------------------------------------------------ */

const SLIDE_W = 13.3333333;
const SLIDE_H = 7.5;

const HEAD = 'Playfair Display'; // theme major font
const BODY = 'Poppins';          // theme minor font

const INK = '000000';   // tx1
const WHITE = 'FFFFFF'; // bg1
const ORANGE = 'C66135';   // accent1
const ORANGE_DK = '944928'; // accent1, lumMod 75%
const MAROON = '8C4742';   // accent2
const TAN = 'CAAC94';      // accent3
const SILVER = 'E7E6E6';   // bg2
const YELLOW = 'FFFF00';

// outerShdw presets used by the original shapes.
// These are factories: pptxgenjs rewrites the shadow object it is handed (pt -> EMU),
// so a shared literal would be re-scaled once per shape that references it.
const SHADOW_CARD = () => ({ type: 'outer', blur: 4, offset: 3, angle: 90, color: INK, opacity: 0.04 });
const SHADOW_SOFT = () => ({ type: 'outer', blur: 50, offset: 0, angle: 45, color: INK, opacity: 0.09 });
const SHADOW_TIP = () => ({ type: 'outer', blur: 20, offset: 3, angle: 90, color: INK, opacity: 0.15 });
const SHADOW_ROUND = () => ({ type: 'outer', blur: 25, offset: 3, angle: 45, color: INK, opacity: 0.15 });
const SHADOW_BAR = () => ({ type: 'outer', blur: 35, offset: 10, angle: 50, color: INK, opacity: 0.1 });
const SHADOW_PANEL = () => ({ type: 'outer', blur: 56, offset: 3, angle: 90, color: INK, opacity: 0.13 });
const SHADOW_QUOTE = () => ({ type: 'outer', blur: 25, offset: 3, angle: 90, color: INK, opacity: 0.2 });

/* ------------------------------------------------------------------ *
 * Repeated copy
 * ------------------------------------------------------------------ */

const LOREM_SOCIIS = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis';
const LOREM_EGET = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget';
const LOREM_PENATIBUS = LOREM_SOCIIS + ' natoque penatibus ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo';
const LOREM_DULUUS = 'Lorem ipsum dolor sit amet, consectetuer legit duluus legit  adipiscing elit. Aenean commodo ligula eget dolo enean massa. Cum sociis';
const EURO = 'The European languages are members of the same family. Their separate existence';

/* ------------------------------------------------------------------ *
 * Helpers
 * ------------------------------------------------------------------ */

function txt(slide, content, o) {
  slide.addText(content, Object.assign({ fontFace: BODY, fontSize: 12, color: INK, valign: 'top' }, o));
}

/** Body copy: 12pt Poppins, 130% leading. */
function body(slide, content, o) {
  txt(slide, content, Object.assign({ lineSpacingMultiple: 1.3 }, o));
}

/** Display headline: Playfair Display. */
function head(slide, content, o) {
  txt(slide, content, Object.assign({ fontFace: HEAD, fontSize: 54 }, o));
}

function rect(slide, o) {
  slide.addShape('rect', o);
}

function hline(slide, o) {
  // 0.5pt is the theme's lnRef idx=1 weight, which the original rules inherit.
  slide.addShape('line', { x: o.x, y: o.y, w: o.w, h: 0, line: { color: o.color, width: o.width || 0.5 } });
}

/**
 * White-to-black linear fade, painted as stacked bands because pptxgenjs has no
 * gradient fill. The bands are opaque (pre-composited over the white page) and
 * slightly overlapping; translucent bands would show seams wherever they abut.
 */
function verticalFade(slide, o) {
  const bands = o.bands || 120;
  const bandH = o.h / bands;
  for (let i = 0; i < bands; i++) {
    const level = Math.round(255 * (1 - (i + 0.5) / bands));
    const hex = ('0' + level.toString(16).toUpperCase()).slice(-2);
    rect(slide, { x: o.x, y: o.y + i * bandH, w: o.w, h: bandH * 1.6, fill: { color: hex + hex + hex } });
  }
}

/** Stand-in for the small coffee line-art icons of the original deck: cup + saucer. */
function iconMark(slide, o) {
  const pen = { color: o.color, width: 1.25 };
  slide.addShape('flowChartManualOperation', {
    x: o.x + o.w * 0.16, y: o.y + o.w * 0.16, w: o.w * 0.6, h: o.w * 0.58, line: pen,
  });
  slide.addShape('line', { x: o.x + o.w * 0.06, y: o.y + o.w * 0.84, w: o.w * 0.8, h: 0, line: pen });
  slide.addShape('arc', {
    x: o.x + o.w * 0.66, y: o.y + o.w * 0.26, w: o.w * 0.3, h: o.w * 0.3,
    angleRange: [270, 90], line: pen,
  });
}

/** Row of five rating stars; `lit` of them filled with `color`. */
function stars(slide, o) {
  for (let i = 0; i < 5; i++) {
    slide.addShape('star5', {
      x: o.x + i * o.step, y: o.y, w: o.size, h: o.size,
      fill: { color: i < o.lit ? o.color : o.dim || o.color },
    });
  }
}

/** Slide furniture inherited from the slide master (rules + running heads). */
function chrome(slide, n) {
  hline(slide, { x: 1.8125, y: 0.5146, w: 10.2049, color: INK });
  hline(slide, { x: 0, y: 7.1536, w: SLIDE_W, color: INK });
  txt(slide, 'Coffee Shop', { x: 0.3203, y: 0.3464, w: 1.3333, h: 0.3365, fontFace: HEAD, fontSize: 14 });
  txt(slide, 'Slide ' + n, { x: 12.0166, y: 0.3464, w: 0.9964, h: 0.3365, fontFace: HEAD, fontSize: 14, align: 'right' });
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

// 1 — Cover
function slide01(pres) {
  const s = pres.addSlide();
  chrome(s, 1);
  verticalFade(s, { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H });
  hline(s, { x: 1.8125, y: 0.3854, w: 10.2049, color: WHITE });
  txt(s, 'Coffee Shop', { x: 0.3203, y: 0.2171, w: 1.3333, h: 0.3365, fontFace: HEAD, fontSize: 14, color: WHITE });
  txt(s, 'Slide 1', { x: 12.1364, y: 0.2171, w: 0.8766, h: 0.3365, fontFace: HEAD, fontSize: 14, color: WHITE, align: 'right' });
  txt(s, 'COFFEESHOPPRESENTATION', {
    x: 2.0866, y: 1.5311, w: 4.1529, h: 0.3365, fontSize: 14, color: WHITE, align: 'center', charSpacing: 3,
  });
  head(s, 'Coffee Shop', { x: 2.0866, y: 1.7796, w: 9.1601, h: 2.0364, fontSize: 115, color: WHITE, align: 'center' });
  body(s, LOREM_SOCIIS, { x: 7.4409, y: 3.9756, w: 3.9132, h: 0.8701, color: WHITE, align: 'justify' });
}

// 2 — Welcome / agenda
function slide02(pres) {
  const s = pres.addSlide();
  chrome(s, 2);
  head(s, 'Welcome to Our Coffee Journey', { x: 0.5967, y: 4.6183, w: 6.3826, h: 1.9186 });

  rect(s, { x: 7.7143, y: 2.3316, w: 4.1817, h: 4.2597, fill: { color: ORANGE }, shadow: SHADOW_CARD() });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum ipsum', {
    x: 7.9588, y: 2.8477, w: 3.6925, h: 0.8701, color: WHITE, align: 'justify',
  });

  const rows = [
    { n: '01.', y: 4.0831 },
    { n: '02.', y: 4.7508 },
    { n: '03.', y: 5.4185 },
  ];
  rows.forEach(function (r) {
    body(s, r.n, { x: 7.9243, y: r.y, w: 0.5701, h: 0.3722, fontFace: HEAD, fontSize: 14, color: WHITE });
    body(s, 'Lorem ipsum dolor sit amet', { x: 8.4619, y: r.y + 0.0136, w: 2.8354, h: 0.345, color: WHITE, align: 'center' });
    hline(s, { x: 7.9588, y: r.y + 0.3929, w: 3.6925, color: WHITE, width: 1 });
    s.addShape('triangle', {
      x: 11.5266, y: r.y + 0.1587, w: 0.124, h: 0.1149, rotate: 180, fill: { color: WHITE },
    });
  });
}

// 3 — Crafting every cup
function slide03(pres) {
  const s = pres.addSlide();
  chrome(s, 3);
  head(s, 'Crafting Every Cup with Care', { x: 0.6265, y: 1.1428, w: 6.1481, h: 1.9186 });

  s.addShape('roundRect', { x: 7.4396, y: 1.5851, w: 0.5, h: 0.05, rectRadius: 0.025, fill: { color: INK } });
  txt(s, 'Quality and dedication in each brew', { x: 7.9615, y: 1.4251, w: 4.6203, h: 0.3702, fontFace: HEAD, fontSize: 16 });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum ipsum adipiscing dolor sit commodo ligula', {
    x: 7.3232, y: 1.9091, w: 4.9219, h: 0.8701, align: 'justify',
  });

  rect(s, { x: 6.2142, y: 3.2352, w: 6.0306, h: 0.7364, fill: { color: ORANGE }, shadow: SHADOW_CARD() });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum', {
    x: 6.5152, y: 3.2997, w: 5.4287, h: 0.6075, color: WHITE,
  });
}

// 4 — Creating a welcoming space
function slide04(pres) {
  const s = pres.addSlide();
  chrome(s, 4);
  rect(s, { x: 2.3719, y: 1.3468, w: 10.9613, h: 5.8093, fill: { color: ORANGE } });

  head(s, 'Creating a Welcoming Space', { x: 8.3887, y: 1.5936, w: 4.3427, h: 2.8273, color: WHITE });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum ipsum adipiscing', {
    x: 8.3887, y: 4.5779, w: 4.1183, h: 0.8701, color: WHITE, align: 'justify',
  });
  body(s, LOREM_EGET, { x: 8.3887, y: 5.4318, w: 4.1183, h: 0.6075, color: WHITE, align: 'justify' });

  rect(s, { x: 0.8574, y: 5.0257, w: 6.0721, h: 1.0090, fill: { color: WHITE }, shadow: SHADOW_CARD() });
  txt(s, '45%', { x: 1.0965, y: 5.1768, w: 1.0358, h: 0.7068, fontFace: HEAD, fontSize: 36 });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit legit et Aenean et legit et', {
    x: 2.2637, y: 5.1844, w: 4.4262, h: 0.6923, fontSize: 14,
  });
}

// 5 — Product highlight
function slide05(pres) {
  const s = pres.addSlide();
  chrome(s, 5);
  head(s, 'Crafting Every Cup with Care', { x: 0.4722, y: 4.5119, w: 5.8472, h: 1.9186 });

  rect(s, { x: 5.5209, y: 1.2874, w: 3.2812, h: 2.6667, fill: { color: ORANGE } });
  s.addShape('ellipse', { x: 5.7917, y: 1.5999, w: 0.9062, h: 0.9062, fill: { color: WHITE } });
  iconMark(s, { x: 6.0058, y: 1.8135, w: 0.478, color: INK });
  body(s, LOREM_EGET, { x: 5.7917, y: 2.6383, w: 2.7290, h: 0.8701, color: WHITE, align: 'justify' });

  txt(s, '34.124K', { x: 7.0139, y: 4.7166, w: 1.6760, h: 0.6394, fontFace: HEAD, fontSize: 32, color: ORANGE });
  txt(s, 'Your Text Here', { x: 8.6903, y: 4.9517, w: 1.5938, h: 0.3365, fontSize: 14 });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer legit duluus legit  adipiscing elit. Aenean commodo ligula eget dolo enean massa. Cum sociis natoque penatibus et cum duluus', {
    x: 7.0139, y: 5.3562, w: 5.4911, h: 0.8701,
  });
}

// 6 — Customer experience (dark)
function slide06(pres) {
  const s = pres.addSlide();
  chrome(s, 6);
  verticalFade(s, { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H });
  hline(s, { x: 1.8125, y: 0.5146, w: 10.2049, color: WHITE });
  hline(s, { x: 0, y: 7.1536, w: SLIDE_W, color: WHITE });
  txt(s, 'Coffee Shop', { x: 0.3203, y: 0.3464, w: 1.3333, h: 0.3365, fontFace: HEAD, fontSize: 14, color: WHITE });
  txt(s, 'Slide 6', { x: 12.0166, y: 0.3464, w: 0.9964, h: 0.3365, fontFace: HEAD, fontSize: 14, color: WHITE, align: 'right' });

  rect(s, { x: 0, y: 2.5312, w: 5.6875, h: 4.6222, fill: { color: WHITE }, shadow: SHADOW_SOFT() });
  head(s, 'Customer Experience', { x: 0.6668, y: 3.0459, w: 4.3535, h: 1.9186 });
  body(s, LOREM_DULUUS, { x: 0.6668, y: 5.1479, w: 4.3535, h: 0.8701, align: 'justify' });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer legit duluus legit  adipiscing elit. Aenean commodo', {
    x: 0.6668, y: 6.0311, w: 4.3535, h: 0.6075, align: 'justify',
  });

  rect(s, { x: 4.8998, y: 3.4497, w: 5.0077, h: 0.6326, fill: { color: ORANGE } });
  txt(s, 'A Cozy Space to Relax, Work, and Connect', {
    x: 5.0346, y: 3.5809, w: 4.7379, h: 0.3702, fontFace: HEAD, fontSize: 16, bold: true, color: WHITE,
  });

  rect(s, { x: 5.6875, y: 4.0827, w: 2.8447, h: 0.4783, fill: { color: WHITE }, shadow: SHADOW_SOFT() });
  stars(s, { x: 6.1211, y: 4.2043, step: 0.29, size: 0.235, lit: 4, color: YELLOW, dim: SILVER });
  body(s, '4,5', { x: 7.5996, y: 4.1306, w: 0.4988, h: 0.3827, fontSize: 14, italic: true });
}

// 7 — Community engagement
function slide07(pres) {
  const s = pres.addSlide();
  chrome(s, 7);
  head(s, 'Community Engagement', { x: 4.1770, y: 1.1659, w: 8.7292, h: 1.0098 });

  body(s, 'Your Title Here', { x: 0.5959, y: 0.9746, w: 1.6748, h: 0.3825, fontSize: 14, italic: true });
  txt(s, '100+', { x: 0.6157, y: 1.3572, w: 1.7778, h: 0.8414, fontFace: HEAD, fontSize: 44, bold: true });
  body(s, LOREM_SOCIIS, { x: 0.6157, y: 2.3330, w: 3.1665, h: 1.1325, align: 'justify' });
  txt(s, 'Building Stronger Connections Through Coffee Events', {
    x: 0.5959, y: 3.8631, w: 2.4874, h: 0.9088, fontFace: HEAD, fontSize: 16, bold: true,
  });

  rect(s, { x: 1.6237, y: 5.3021, w: 5.4287, h: 0.8404, fill: { color: ORANGE }, shadow: SHADOW_SOFT() });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa.', {
    x: 1.8396, y: 5.4185, w: 5.4287, h: 0.6075, color: WHITE,
  });
}

// 8 — Meet our coffee experts
function slide08(pres) {
  const s = pres.addSlide();
  chrome(s, 8);

  const cols = [
    { x: 5.4113, num: '01', numAt: 5.4842, chip: 5.6558, name: 5.8901 },
    { x: 7.9216, num: '02', numAt: 8.0755, chip: 8.1772, name: 8.4114 },
    { x: 10.4319, num: '03', numAt: 10.7390, chip: 10.6769, name: 10.9111 },
  ];
  cols.forEach(function (c) {
    verticalFade(s, { x: c.x, y: 1.3105, w: 2.5104, h: 5.8524, bands: 80 });
  });

  rect(s, { x: 0, y: 0.8854, w: 5.1563, h: 6.2813, fill: { color: ORANGE } });
  head(s, 'Meet Our Coffee Experts', { x: 0.7499, y: 1.2177, w: 3.6555, h: 2.8273, color: WHITE });
  body(s, LOREM_SOCIIS, { x: 0.7499, y: 4.3356, w: 3.9688, h: 0.8701, color: WHITE, align: 'justify' });
  body(s, LOREM_EGET, { x: 0.7499, y: 5.2056, w: 3.9688, h: 0.6075, color: WHITE, align: 'justify' });

  cols.forEach(function (c) {
    rect(s, { x: c.chip, y: 4.8536, w: 2.2656, h: 0.4901, fill: { color: ORANGE }, shadow: SHADOW_SOFT() });
    txt(s, 'Davina Marvin', { x: c.name, y: 4.9136, w: 1.7965, h: 0.3702, fontFace: HEAD, fontSize: 16, color: WHITE, align: 'center' });
    txt(s, c.num, { x: c.numAt, y: 5.9721, w: 1.1826, h: 0.9088, fontFace: HEAD, fontSize: 48, bold: true, color: WHITE });
  });
}

// 9 — Storefront & branding showcase
function slide09(pres) {
  const s = pres.addSlide();
  chrome(s, 9);
  rect(s, { x: 0, y: 6.2648, w: SLIDE_W, h: 0.9118, fill: { color: ORANGE } });

  // stand-in for the laptop mock-up photo: screen bezel over a lens-shaped base
  s.addShape('ellipse', { x: 4.6620, y: 6.3900, w: 8.5880, h: 0.2200, fill: { color: 'EFEFEF' } });
  s.addShape('ellipse', { x: 4.8000, y: 6.4600, w: 8.3250, h: 0.1500, fill: { color: 'A5A6A9' } });
  s.addShape('roundRect', { x: 5.4664, y: 1.6667, w: 6.9930, h: 4.7639, rectRadius: 0.09, fill: { color: '202125' } });
  rect(s, { x: 5.6606, y: 1.9444, w: 6.6040, h: 4.125, fill: { color: '00BEF2' } });
  txt(s, '[image]', { x: 5.6606, y: 3.8, w: 6.6040, h: 0.4, fontSize: 14, color: WHITE, align: 'center' });

  head(s, 'Storefront & Branding Showcase', { x: 0.7812, y: 1.4576, w: 4.3438, h: 2.8273 });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque consectetuer adipiscing elit. Aenean', {
    x: 0.7812, y: 4.4691, w: 3.9915, h: 1.1325, align: 'justify',
  });

  s.addShape('roundRect', {
    x: 5.9737, y: 0.9727, w: 3.4272, h: 1.5787, rectRadius: 0.035, fill: { color: WHITE }, shadow: SHADOW_TIP(),
  });
  txt(s, 'A Glimpse of Our Ambiance, Packaging, and Design', {
    x: 6.1498, y: 1.1476, w: 3.0748, h: 0.6394, fontFace: HEAD, fontSize: 16,
  });
  txt(s, '250+', { x: 6.1498, y: 1.8707, w: 1.1706, h: 0.5052, fontFace: HEAD, fontSize: 24 });
  txt(s, 'Achievements', { x: 7.1680, y: 1.9909, w: 1.6458, h: 0.3365, fontSize: 14, italic: true });
}

// 10 — Section infographic: tilted service cards
function slide10(pres) {
  const s = pres.addSlide();
  chrome(s, 10);

  head(s, 'Section Infographic', { x: 0.9406, y: 1.4709, w: 4.6423, h: 1.7366, lineSpacingMultiple: 0.9 });
  txt(s, 'From Bean to Cup — Ethical and Sustainable Practices', {
    x: 0.9406, y: 3.7484, w: 5.0466, h: 0.3365, fontFace: HEAD, fontSize: 14,
  });
  [4.1873, 5.1590].forEach(function (y) {
    body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus dolor sit amet, consectetuer', {
      x: 0.9406, y: y, w: 5.4501, h: 0.8701, align: 'justify',
    });
  });

  const cards = [
    { fill: ORANGE, rot: 355.914, card: [7.7484, 1.4787], dot: [7.3515, 1.8848], title: [8.3980, 1.6669], copy: [8.4318, 1.9899], align: 'left', num: '1', numAt: [7.5397, 2.0700] },
    { fill: MAROON, rot: 0, card: [7.3455, 3.1010], dot: [11.5844, 3.3413], title: [8.9138, 3.2277], copy: [7.6122, 3.5983], align: 'right', num: '2', numAt: [11.7726, 3.5266] },
    { fill: TAN, rot: 2.985, card: [7.7486, 4.7189], dot: [7.3486, 4.8386], title: [8.4356, 4.8518], copy: [8.4093, 5.2551], align: 'left', num: '3', numAt: [7.5372, 5.0239] },
  ];
  cards.forEach(function (c) {
    rect(s, { x: c.card[0], y: c.card[1], w: 4.6423, h: 1.2889, rotate: c.rot, fill: { color: c.fill } });
    txt(s, 'Service Here', {
      x: c.title[0], y: c.title[1], w: 2.4062, h: 0.3702, rotate: c.rot,
      fontFace: HEAD, fontSize: 16, color: WHITE, align: c.align,
    });
    body(s, EURO, { x: c.copy[0], y: c.copy[1], w: 3.7083, h: 0.6075, rotate: c.rot, color: WHITE, align: c.align });
    s.addShape('ellipse', {
      x: c.dot[0], y: c.dot[1], w: 0.8081, h: 0.8081, rotate: c.rot, fill: { color: WHITE }, shadow: SHADOW_ROUND(),
    });
    txt(s, c.num, { x: c.numAt[0], y: c.numAt[1], w: 0.4310, h: 0.4375, fontFace: HEAD, fontSize: 20, align: 'center' });
  });
}

// 11 — Section infographic: year circles
function slide11(pres) {
  const s = pres.addSlide();
  chrome(s, 11);

  rect(s, { x: 2.7537, y: 4.4366, w: 7.8248, h: 0.7852, fill: { color: WHITE }, shadow: SHADOW_BAR() });

  const years = [
    { label: '2021', fill: ORANGE, chip: ORANGE_DK, cx: 1.0298, cy: 3.6470, d: 2.1496, sx: 1.7742, sy: 5.3512, sd: 0.6605 },
    { label: '2022', fill: MAROON, chip: MAROON, cx: 3.6851, cy: 3.3279, d: 2.7290, sx: 4.6299, sy: 5.4919, sd: 0.8388 },
    { label: '2023', fill: ORANGE, chip: ORANGE_DK, cx: 6.9192, cy: 3.6470, d: 2.1496, sx: 7.6636, sy: 5.3512, sd: 0.6605 },
    { label: '2024', fill: MAROON, chip: MAROON, cx: 9.5745, cy: 3.3279, d: 2.7290, sx: 10.5194, sy: 5.4919, sd: 0.8388 },
  ];
  years.forEach(function (y, i) {
    const mid = y.cx + y.d / 2;
    s.addShape('ellipse', { x: y.cx, y: y.cy, w: y.d, h: y.d, fill: { color: y.fill } });
    s.addShape('roundRect', {
      x: y.sx, y: y.sy, w: y.sd, h: y.sd, rectRadius: y.sd * 0.2775,
      fill: { color: y.chip }, line: { color: WHITE, width: 4.5 },
    });
    iconMark(s, { x: y.sx + y.sd * 0.23, y: y.sy + y.sd * 0.23, w: y.sd * 0.54, color: WHITE });
    txt(s, y.label, {
      x: mid - 0.6115, y: 4.0114, w: 1.2230, h: 0.4375,
      fontFace: HEAD, fontSize: 20, bold: true, color: WHITE, align: 'center',
    });
    body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ', {
      x: mid - 1.0585, y: 4.3989, w: 2.1170, h: 0.8701, color: WHITE, align: 'center',
    });
  });

  head(s, 'Section Infographic', { x: 0.9930, y: 1.1690, w: 11.3463, h: 0.9189, align: 'center', lineSpacingMultiple: 0.9 });
  body(s, LOREM_PENATIBUS, { x: 1.8659, y: 2.0876, w: 9.6009, h: 0.6075, align: 'center' });
}

// 12 — Section infographic: process pins
function slide12(pres) {
  const s = pres.addSlide();
  chrome(s, 12);

  const steps = [
    { label: 'Process 1', fill: ORANGE, x: 1.6528, y: 2.9260 },
    { label: 'Process 2', fill: MAROON, x: 4.5093, y: 3.6216 },
    { label: 'Process 3', fill: ORANGE, x: 7.3668, y: 2.9260 },
    { label: 'Process 4', fill: MAROON, x: 10.2229, y: 3.6216 },
  ];
  steps.forEach(function (p) {
    s.addShape('ellipse', { x: p.x, y: p.y, w: 1.4881, h: 1.4881, fill: { color: p.fill } });
    s.addShape('triangle', { x: p.x + 0.5677, y: p.y + 1.4360, w: 0.2904, h: 0.3116, flipV: true, fill: { color: p.fill } });
    s.addShape('ellipse', { x: p.x + 0.1532, y: p.y + 0.1532, w: 1.1818, h: 1.1818, fill: { color: WHITE } });
    s.addShape('blockArc', {
      x: p.x - 0.1058, y: p.y - 0.1058, w: 1.6996, h: 1.6996,
      angleRange: [178.705, 1.365], arcThicknessRatio: 0.0605, fill: { color: p.fill },
    });
    iconMark(s, { x: p.x + 0.4425, y: p.y + 0.4425, w: 0.603, color: p.fill });
    txt(s, p.label, {
      x: p.x - 0.5468, y: p.y + 1.9982, w: 2.5820, h: 0.3365, fontFace: HEAD, fontSize: 14, align: 'center',
    });
    body(s, 'Lorem elit ipsum congue dolor sit amet, Maecenas', {
      x: p.x - 0.5781, y: p.y + 2.3346, w: 2.5820, h: 0.6075, align: 'center',
    });
  });

  head(s, 'Section Infographic', { x: 0.9930, y: 0.9358, w: 11.3463, h: 0.9189, align: 'center', lineSpacingMultiple: 0.9 });
  body(s, LOREM_PENATIBUS, { x: 1.8659, y: 1.8547, w: 9.6009, h: 0.6075, align: 'center' });
}

// 13 — Section infographic: quarter bars
function slide13(pres) {
  const s = pres.addSlide();
  chrome(s, 13);

  const bars = [
    { label: 'Q1', fill: ORANGE, x: 6.0803, y: 3.3767, h: 3.2761, lw: 1.6287, lx: 6.2405, ix: 6.6971, iy: 2.3457 },
    { label: 'Q2', fill: MAROON, x: 8.4188, y: 2.2986, h: 4.3542, lw: 1.7879, lx: 8.4988, ix: 9.0356, iy: 1.2783 },
    { label: 'Q3', fill: TAN, x: 10.7573, y: 3.7673, h: 2.8853, lw: 1.7152, lx: 10.8730, ix: 11.3740, iy: 2.6620 },
  ];
  bars.forEach(function (b) {
    rect(s, { x: b.x, y: b.y, w: 1.9486, h: b.h, fill: { color: b.fill } });
    txt(s, b.label, { x: b.lx, y: 4.7998, w: b.lw, h: 1.3126, fontFace: HEAD, fontSize: 72, color: WHITE, align: 'center' });
    s.addShape('roundRect', { x: b.ix, y: b.iy, w: 0.7148, h: 0.7148, rectRadius: 0.0763, fill: { color: b.fill } });
    iconMark(s, { x: b.ix + 0.124, y: b.iy + 0.124, w: 0.4669, color: WHITE });
  });

  head(s, 'Section Infographic', { x: 0.6279, y: 1.2783, w: 4.7842, h: 1.7366, lineSpacingMultiple: 0.9 });
  txt(s, '480+', { x: 0.7443, y: 4.0707, w: 2.2799, h: 0.9189, fontFace: HEAD, fontSize: 54, lineSpacingMultiple: 0.9 });
  body(s, 'Your Title here', { x: 2.5559, y: 4.3576, w: 1.4831, h: 0.3450, italic: true, align: 'justify' });
  body(s, 'Lorem ipsum amet dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. ipsum dolor sit amet, consectetuer', {
    x: 0.7546, y: 5.1747, w: 4.6577, h: 0.8701, align: 'justify',
  });
  body(s, 'Lorem ipsum amet dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. ', {
    x: 0.7546, y: 6.0449, w: 4.6577, h: 0.6075, align: 'justify',
  });
}

// 14 — Section infographic: two stat panels
function slide14(pres) {
  const s = pres.addSlide();
  chrome(s, 14);

  rect(s, { x: -0.0405, y: 0, w: 5.6667, h: SLIDE_H, fill: { color: ORANGE } });
  hline(s, { x: 1.8125, y: 0.5146, w: 3.8146, color: WHITE });
  hline(s, { x: 0, y: 7.1536, w: 5.6260, color: WHITE });
  txt(s, 'Coffee Shop', { x: 0.3203, y: 0.3464, w: 1.3333, h: 0.3365, fontFace: HEAD, fontSize: 14, color: WHITE });

  head(s, 'Section Infographic', { x: 0.6616, y: 1.3648, w: 4.7842, h: 1.7366, color: WHITE, lineSpacingMultiple: 0.9 });
  body(s, LOREM_SOCIIS + ' natoque penatibus ipsum dolor sit amet, consectetuer', {
    x: 0.6616, y: 4.0770, w: 4.2450, h: 1.1325, color: WHITE, align: 'justify',
  });
  body(s, LOREM_SOCIIS + ' natoque', { x: 0.6616, y: 5.2654, w: 4.2450, h: 0.8701, color: WHITE, align: 'justify' });

  const panels = [
    { stat: '+73,9%', fill: ORANGE, x: 6.3622, tx: 6.7349, bx: 6.7395 },
    { stat: '+93,9%', fill: MAROON, x: 9.6613, tx: 10.0384, bx: 10.0424 },
  ];
  panels.forEach(function (p) {
    s.addShape('round2DiagRect', {
      x: p.x, y: 1.1119, w: 3.0102, h: 5.2755, rectRadius: 0.55, fill: { color: p.fill }, shadow: SHADOW_PANEL(),
    });
    txt(s, p.stat, { x: p.tx, y: 1.4281, w: 2.3447, h: 0.8414, fontFace: HEAD, fontSize: 44, color: WHITE });
    body(s, 'Lorem ipsum sit dolor sit amet, ligula consectetuer adipiscing sit elit. Aenean commodo elit ligula eget dolor. Aenean massa. ', {
      x: p.bx, y: 2.4768, w: 2.3392, h: 1.3950, color: WHITE, align: 'justify',
    });
    [3.9762, 4.3977, 4.8196, 5.2410].forEach(function (y) {
      body(s, 'Lorem ipsum sit dolor sit', {
        x: p.bx, y: y, w: 2.4534, h: 0.3450, color: WHITE,
        bullet: { characterCode: '2713', indent: 13.5 },
      });
    });
  });
}

// 15 — Testimonial
function slide15(pres) {
  const s = pres.addSlide();
  chrome(s, 15);

  txt(s, '\u201C', {
    x: 8.1801, y: -1.2511, w: 4.8090, h: 12.1170, fontFace: 'Archivo Black', fontSize: 714,
    color: INK, transparency: 90,
  });

  head(s, [
    { text: 'Great Coffee Builds Great Connections \u2013 It\u2019s Not Only About ' },
    { text: 'Serving Drinks, But About Creating', options: { color: ORANGE } },
    { text: ' Lasting Experiences and ' },
    { text: 'Community Growth', options: { color: ORANGE } },
  ], { x: 6.2274, y: 1.8817, w: 6.3346, h: 3.7361, fontSize: 36 });

  rect(s, { x: 0.5679, y: 0.8854, w: 5.2632, h: 5.7292, fill: { color: ORANGE } });
  s.addShape('round2DiagRect', {
    x: 2.9163, y: 4.3470, w: 2.7671, h: 1.6494, rectRadius: 0.01, fill: { color: WHITE }, shadow: SHADOW_QUOTE(),
  });
  txt(s, 'Medison Madina', { x: 3.0627, y: 4.4751, w: 2.1414, h: 0.4172, fontFace: HEAD, fontSize: 16, italic: true });
  body(s, 'Your Position', { x: 3.0627, y: 4.8072, w: 2.1414, h: 0.3450, italic: true });
  stars(s, { x: 3.0687, y: 5.2141, step: 0.241, size: 0.2203, lit: 5, color: ORANGE });
  body(s, 'Lorem ipsum dolor sit amet, ', { x: 3.0627, y: 5.5237, w: 2.5434, h: 0.3450 });
}

// 16 — Thank you
function slide16(pres) {
  const s = pres.addSlide();
  chrome(s, 16);
  verticalFade(s, { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H });
  hline(s, { x: 1.8125, y: 0.3854, w: 10.2049, color: WHITE });
  hline(s, { x: 0, y: 7.1536, w: SLIDE_W, color: WHITE });
  txt(s, 'Coffee Shop', { x: 0.3203, y: 0.2171, w: 1.3333, h: 0.3365, fontFace: HEAD, fontSize: 14, color: WHITE });
  txt(s, 'Slide 16', { x: 12.1364, y: 0.2171, w: 0.8766, h: 0.3365, fontFace: HEAD, fontSize: 14, color: WHITE, align: 'right' });

  txt(s, 'COFFEESHOPPRESENTATION', {
    x: 5.2295, y: 1.5249, w: 4.1529, h: 0.3365, fontSize: 14, color: WHITE, charSpacing: 3,
  });
  head(s, 'Thank You!!', { x: 5.2295, y: 1.8506, w: 7.8880, h: 1.7166, fontSize: 96, color: WHITE });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis ipsum dolor sit amet, consectetuer', {
    x: 5.2295, y: 3.5670, w: 7.1444, h: 0.6075, color: WHITE, align: 'justify',
  });

  const contacts = [
    { text: '+214-704-76532', x: 6.5756, w: 1.9574 },
    { text: 'thesisone@gmail.com', x: 8.4846, w: 2.0888 },
    { text: 'www.companyinfo.com', x: 10.8871, w: 2.2307 },
  ];
  contacts.forEach(function (c) {
    body(s, c.text, { x: c.x, y: 6.3733, w: c.w, h: 0.3450, color: WHITE });
  });
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */

function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'DECK', width: SLIDE_W, height: SLIDE_H });
  pres.layout = 'DECK';
  pres.title = 'Coffee Shop';
  pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
    slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16]
    .forEach(function (fn) { fn(pres); });

  return pres.writeFile({
    fileName: path.join(__dirname, '018eeb0c-98be-466b-952d-ee27ab314be0_grok_final.pptx'),
  });
}

build().then(function (f) { console.log('wrote', f); }).catch(function (e) { console.error(e); process.exit(1); });
