/**
 * Recreation of "Driving Sustainable Profit Growth" (20 slides, 13.333 x 7.5 in)
 * using pptxgenjs only. Raster photos in the source deck are replaced with
 * light-gray "[image]" placeholder shapes of the same position and size.
 *
 * Run:  node 16468c51-d384-44f9-a16b-a866ac16808d_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Design tokens
 * ------------------------------------------------------------------ */

const ORANGE = 'F0752F';
const INK = '404040'; // tx1 lumMod 75% - the deck's "black"
const WHITE = 'FFFFFF';
const PAPER = 'FBFBFB'; // master background
const MIST = 'F2F2F2'; // chip / bar background
const SILVER = 'D9D9D9';
const ASH = 'BFBFBF';
const STEEL = '808080';
const PEACH = 'FCE3D5'; // 25% orange bar segment
const PLACEHOLDER = 'F2F2F2'; // stands in for photographs
const FRAME = '262626'; // device mockup outlines

// Gradient backgrounds are flattened to their average tone (pptxgenjs has no
// gradient fill); each value was sampled from the reference render.
const NIGHT = '343434'; // tx1 gradient used on slides 1 and 13
const MIDNIGHT = '1B1B1B'; // darker gradient of slide 9
const DUSK = '414141'; // translucent gradient over a photo (slides 5, 17)
const DUSK_SOFT = '474747'; // same, slide 18 left panel

const HEAD = 'Work Sans Medium'; // theme major font
const BODY = 'Plus Jakarta Sans'; // theme minor font

const CARD_SHADOW = { type: 'outer', color: '000000', opacity: 0.05, blur: 30, offset: 10, angle: 45 };
const BADGE_SHADOW = { type: 'outer', color: '000000', opacity: 0.1, blur: 20, offset: 3, angle: 45 };

/* ------------------------------------------------------------------ *
 * Primitive helpers
 * ------------------------------------------------------------------ */

// PowerPoint text boxes are top-anchored; pptxgenjs defaults to middle.
function text(s, str, o) {
  s.addText(str, Object.assign({ valign: 'top', fontFace: BODY, color: INK }, o));
}

// Body copy: 12 pt / 130% line spacing (the deck's list-style default).
function copy(s, x, y, w, h, str, o) {
  text(s, str, Object.assign({ x, y, w, h, fontSize: 12, lineSpacingMultiple: 1.3 }, o));
}

// Display type set in the major font.
function title(s, x, y, w, h, str, size, o) {
  text(s, str, Object.assign({ x, y, w, h, fontSize: size, fontFace: HEAD }, o));
}

function rect(s, x, y, w, h, fill, o) {
  s.addShape('rect', Object.assign({ x, y, w, h, fill }, o));
}

function roundRect(s, x, y, w, h, radius, fill, o) {
  s.addShape('roundRect', Object.assign({ x, y, w, h, rectRadius: radius, fill }, o));
}

function oval(s, x, y, d, fill, o) {
  s.addShape('ellipse', Object.assign({ x, y, w: d, h: d, fill }, o));
}

/* ------------------------------------------------------------------ *
 * Recurring composite elements
 * ------------------------------------------------------------------ */

// Small orange halo dot + caption, used as an eyebrow above every headline.
// (x, y) is the top-left of the caption box.
function eyebrow(s, x, y, label, color) {
  oval(s, x - 0.237, y + 0.061, 0.181, { color: ORANGE, transparency: 80 });
  oval(s, x - 0.205, y + 0.093, 0.116, { color: ORANGE });
  text(s, label, { x, y, w: 1.758, h: 0.303, fontSize: 12, color: color || INK });
}

// Bullet dot + caption in the same style but tighter, used inside cards.
function dotLabel(s, x, y, label, dot, size, w) {
  oval(s, x, y, 0.181, { color: dot, transparency: 80 });
  oval(s, x + 0.032, y + 0.033, 0.116, { color: dot });
  text(s, label, { x: x + 0.206, y: y - 0.048, w: w || 1.129, h: 0.269, fontSize: size || 10 });
}

// Fully rounded pill tag ("Profit growth", "Revenue", ...).
function chip(s, x, y, w, h, label, fill, color, size) {
  roundRect(s, x, y, w, h, h / 2, { color: fill });
  text(s, label, {
    x, y: y + 0.011, w, h: h - 0.022, align: 'center', fontSize: size, color,
  });
}

// A row of pill tags laid out from a left edge.
function chipRow(s, x, y, tags, h, size, gap) {
  let cx = x;
  tags.forEach((t) => {
    chip(s, cx, y, t.w, h, t.label, t.fill || MIST, t.color || INK, size);
    cx += t.w + (gap === undefined ? 0.093 : gap);
  });
}

// Line-art icons rebuilt from primitive shapes. Each draws inside the unit
// square (x, y, d) of its host circle.
const ICON = {
  // Two tall arrows flanking a dollar sign, with a short arrow above.
  money(s, x, y, d, fg) {
    s.addShape('upArrow', { x: x + 0.42 * d, y: y + 0.19 * d, w: 0.11 * d, h: 0.21 * d, fill: { color: fg } });
    s.addShape('upArrow', { x: x + 0.22 * d, y: y + 0.33 * d, w: 0.12 * d, h: 0.37 * d, fill: { color: fg } });
    s.addShape('upArrow', { x: x + 0.63 * d, y: y + 0.33 * d, w: 0.12 * d, h: 0.37 * d, fill: { color: fg } });
    text(s, '$', {
      x: x + 0.35 * d, y: y + 0.38 * d, w: 0.3 * d, h: 0.3 * d, align: 'center', valign: 'middle',
      color: fg, fontSize: Math.round(d * 24), bold: true, wrap: false,
    });
  },
  // Open palm catching two rising arrows.
  hand(s, x, y, d, fg) {
    s.addShape('upArrow', { x: x + 0.27 * d, y: y + 0.19 * d, w: 0.12 * d, h: 0.22 * d, fill: { color: fg } });
    s.addShape('upArrow', { x: x + 0.5 * d, y: y + 0.19 * d, w: 0.12 * d, h: 0.22 * d, fill: { color: fg } });
    s.addShape('blockArc', {
      x: x + 0.16 * d, y: y + 0.3 * d, w: 0.48 * d, h: 0.48 * d,
      fill: { color: fg }, angleRange: [0, 180], arcThicknessRatio: 0.3,
    });
    s.addShape('roundRect', {
      x: x + 0.55 * d, y: y + 0.455 * d, w: 0.3 * d, h: 0.085 * d,
      rectRadius: 0.03, fill: { color: fg }, rotate: -24,
    });
  },
  // Rising zig-zag with an arrow head.
  trend(s, x, y, d, fg) {
    trendArrow(s, x + 0.2 * d, y + 0.36 * d, 0.6 * d, 0.28 * d, fg, Math.max(1.5, d * 3.4));
  },
  mail(s, x, y, d, fg) {
    const ln = { color: fg, width: Math.max(0.8, d * 1.9) };
    s.addShape('rect', { x: x + 0.24 * d, y: y + 0.33 * d, w: 0.52 * d, h: 0.34 * d, fill: { color: fg, transparency: 100 }, line: ln });
    s.addShape('line', { x: x + 0.24 * d, y: y + 0.34 * d, w: 0.26 * d, h: 0.16 * d, line: ln });
    s.addShape('line', { x: x + 0.5 * d, y: y + 0.5 * d, w: 0.26 * d, h: -0.16 * d, line: ln });
  },
  phone(s, x, y, d, fg) {
    text(s, '\u2706', {
      x, y, w: d, h: d, align: 'center', valign: 'middle', color: fg,
      fontSize: Math.round(d * 30), fontFace: 'DejaVu Sans', wrap: false,
    });
  },
  globe(s, x, y, d, fg) {
    const ln = { color: fg, width: Math.max(0.8, d * 1.7) };
    const clear = { color: fg, transparency: 100 };
    s.addShape('ellipse', { x: x + 0.26 * d, y: y + 0.26 * d, w: 0.48 * d, h: 0.48 * d, fill: clear, line: ln });
    s.addShape('ellipse', { x: x + 0.4 * d, y: y + 0.26 * d, w: 0.2 * d, h: 0.48 * d, fill: clear, line: ln });
    s.addShape('line', { x: x + 0.26 * d, y: y + 0.5 * d, w: 0.48 * d, h: 0, line: ln });
  },
};

function iconCircle(s, x, y, d, bg, icon, fg) {
  oval(s, x, y, d, { color: bg });
  icon(s, x, y, d, fg);
}

// Stand-in for a photograph: flat tile at the picture's own frame.
// radius 0 must use a plain rect - pptxgenjs gives roundRect a default corner.
function photo(s, x, y, w, h, o) {
  const opt = o || {};
  const radius = opt.radius === undefined ? 0.18 : opt.radius;
  const shape = opt.shape || (radius > 0 ? 'roundRect' : 'rect');
  s.addShape(shape, Object.assign(
    { x, y, w, h, fill: { color: opt.fill || PLACEHOLDER } },
    shape === 'roundRect' ? { rectRadius: radius } : {},
    opt.rotate ? { rotate: opt.rotate } : {},
  ));
  if (opt.label !== false) {
    text(s, '[image]', {
      x, y, w, h, align: 'center', valign: 'middle', fontSize: opt.labelSize || 11,
      color: opt.labelColor || ASH, rotate: opt.rotate, wrap: false,
    });
  }
}

// Phone / tablet mockup: dark bezel with a placeholder screen inside.
function device(s, x, y, w, h, radius, rotate, labelled) {
  s.addShape('roundRect', Object.assign({ x, y, w, h, rectRadius: radius, fill: { color: FRAME } },
    rotate ? { rotate } : {}));
  photo(s, x + 0.1, y + 0.1, w - 0.2, h - 0.2,
    { radius: radius - 0.02, rotate, label: labelled !== false });
}

// "PROFITERA GLOBAL" lockup: skewed cube mark + wordmark.
function brandMark(s, x, y, w, h) {
  s.addShape('cube', { x, y, w, h, fill: { color: ORANGE }, line: { color: ORANGE, width: 0.75 } });
}

function brandLockup(s, x, y, size, color) {
  const markW = size === 8 ? 0.152 : 0.318;
  const markH = size === 8 ? 0.122 : 0.254;
  brandMark(s, x, y + (size === 8 ? 0.057 : 0.058), markW, markH);
  text(s, 'PROFITERA GLOBAL', {
    x: x + (size === 8 ? 0.153 : 0.456), y, w: size === 8 ? 1.702 : 2.904, h: size === 8 ? 0.236 : 0.370,
    fontSize: size, fontFace: HEAD, color,
  });
}

// Rising zig-zag arrow: two dips then a diagonal to an arrow head.
function trendArrow(s, x, y, w, h, color, width) {
  const line = { color: color || ORANGE, width: width || 3 };
  s.addShape('line', { x, y: y + h, w: w * 0.34, h: -h * 0.62, line });
  s.addShape('line', { x: x + w * 0.34, y: y + h * 0.38, w: w * 0.2, h: h * 0.28, line });
  s.addShape('line', {
    x: x + w * 0.54, y: y + h * 0.66, w: w * 0.46, h: -h * 0.66,
    line: Object.assign({ endArrowType: 'triangle' }, line),
  });
}

// Master-slide furniture: corner lockup (top-right) and page number.
function masterLogo(s, color) {
  brandLockup(s, 11.442, 0.448, 8, color || INK);
}

function pageNumber(s, n, color) {
  text(s, String(n), {
    x: 9.904, y: 6.747, w: 3.0, h: 0.399, align: 'right', valign: 'middle',
    fontSize: 14, fontFace: HEAD, color,
  });
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

function slide01(s) {
  s.background = { color: NIGHT };
  text(s, 'www.yourwebsite.com', { x: 5.678, y: 0.501, w: 6.717, h: 0.303, fontSize: 12, align: 'right', color: ORANGE });
  brandMark(s, 1.106, 1.901, 0.235, 0.188);
  text(s, 'PROFITERA GLOBAL', { x: 1.429, y: 1.844, w: 2.223, h: 0.303, fontSize: 12, fontFace: HEAD, color: WHITE });
  title(s, 0.931, 2.825, 8.436, 2.121, 'Driving Sustainable \nProfit Growth', 60, { color: WHITE, wrap: false });
  trendArrow(s, 1.106, 5.278, 0.504, 0.252);
  text(s, 'Profit Growth Presentation Template', { x: 1.912, y: 5.236, w: 3.564, h: 0.337, fontSize: 14, color: WHITE, wrap: false });
  title(s, 8.958, 5.404, 2.498, 0.774, '90%', 40, { color: WHITE, align: 'right' });
  copy(s, 8.074, 6.255, 4.32, 0.596, 'The profit growth drives revenue and ensures sustainable business success for all.', { align: 'right', color: WHITE });
  iconCircle(s, 11.675, 5.485, 0.613, ORANGE, ICON.money, WHITE);
}

function slide02(s) {
  s.background = { color: PAPER };
  masterLogo(s);
  pageNumber(s, 2, SILVER);
  photo(s, 6.667, 0.867, 5.481, 2.434, { radius: 0.2 });
  eyebrow(s, 1.182, 1.043, 'Profit growth');
  title(s, 0.828, 1.504, 7.442, 0.774, 'Table of Content', 40);
  copy(s, 0.828, 2.53, 5.741, 0.596, 'Profit growth boosts revenue, reduces costs, and \nensures long-term sustainable business success.');

  const agenda = [
    { n: '01', label: 'Profit Growth Overview', x: 0.945, y: 4.199 },
    { n: '02', label: 'Key Growth Drivers', x: 0.945, y: 5.627 },
    { n: '03', label: 'Performance Analysis', x: 7.051, y: 4.199 },
    { n: '04', label: 'Future Business Strategies', x: 7.051, y: 5.627 },
  ];
  agenda.forEach((it) => {
    roundRect(s, it.x, it.y, 5.097, 1.0, 0.167, { color: PAPER }, { shadow: CARD_SHADOW });
    title(s, it.x + 0.24, it.y + 0.18, 1.057, 0.64, it.n, 32, { color: ORANGE });
    text(s, it.label, { x: it.x + 1.116, y: it.y + 0.298, w: 3.98, h: 0.404, fontSize: 18 });
  });
}

function slide03(s) {
  s.background = { color: PAPER };
  masterLogo(s);
  pageNumber(s, 3, SILVER);
  eyebrow(s, 1.182, 1.264, 'Profit growth');
  title(s, 0.828, 1.725, 4.675, 1.313, 'People Powering \nOur Success', 36);
  chipRow(s, 0.951, 4.772, [
    { label: 'Profitability', w: 1.26, fill: ORANGE, color: WHITE },
    { label: 'Operational', w: 1.26 },
  ], 0.312, 10, 0.161);
  copy(s, 0.844, 5.378, 5.741, 0.859, 'Profit growth refers to the increase in a company\u2019s \nprofit over a specific period, indicating improved \nperformance and efficiency.');

  const people = [
    { name: 'Sophia Carren', y: 1.088 },
    { name: 'Daniel Hidayat', y: 3.993 },
  ];
  people.forEach((p) => {
    photo(s, 6.013, p.y, 2.419, 2.419, { radius: 0.2 });
    title(s, 8.76, p.y + 0.45, 4.452, 0.64, p.name, 32);
    copy(s, 8.76, p.y + 1.373, 4.32, 0.596, 'The profit growth drives revenue and \nensures sustainable business success.');
  });
}

function slide04(s) {
  s.background = { color: PAPER };
  masterLogo(s);
  pageNumber(s, 4, SILVER);
  eyebrow(s, 1.182, 1.3, 'Profit growth');
  title(s, 0.828, 1.818, 6.04, 1.313, 'Risk Assessment\nand Mitigation', 36);
  title(s, 0.812, 4.668, 4.156, 0.909, '$25.130', 48);
  chip(s, 3.568, 5.106, 1.038, 0.258, 'Profit growth', ORANGE, WHITE, 8);
  copy(s, 0.844, 5.699, 5.741, 0.859, 'Profit growth refers to the increase in a company\u2019s net profit over \na specific period, indicating improved financial performance and \nbusiness efficiency. ');

  // Donut: four arcs sharing one 3.729" square, 48 pt stroke.
  const arcs = [
    { from: 249.85, to: 346.45, color: SILVER, transparency: 50 },
    { from: 344.41, to: 44.39, color: MIST },
    { from: 135.17, to: 251.17, color: ORANGE },
    { from: 42.83, to: 137.89, color: INK },
  ];
  arcs.forEach((a) => {
    s.addShape('arc', {
      x: 7.159, y: 1.886, w: 3.729, h: 3.729, angleRange: [a.from, a.to],
      line: { color: a.color, width: 48, transparency: a.transparency },
    });
  });
  photo(s, 7.85, 2.586, 2.335, 2.335, { shape: 'ellipse', labelSize: 10 });

  [
    { t: '10%', x: 10.594, y: 4.048, w: 0.505, c: ASH },
    { t: '20%', x: 9.693, y: 2.022, w: 0.533, c: ASH },
    { t: '35%', x: 6.918, y: 3.336, w: 0.521, c: WHITE },
    { t: '35%', x: 8.756, y: 5.478, w: 0.521, c: WHITE },
  ].forEach((p) => {
    text(s, p.t, { x: p.x, y: p.y, w: p.w, h: 0.269, fontSize: 10, bold: true, align: 'center', color: p.c, wrap: false });
  });

  const cards = [
    { x: 5.237, y: 3.9, w: 2.068, label: 'Strategic Investment', dot: ORANGE },
    { x: 9.56, y: 5.478, w: 2.068, label: 'Product Innovation', dot: INK },
    { x: 7.088, y: 1.136, w: 2.068, label: 'Customer Retention', dot: SILVER },
    { x: 10.785, y: 2.688, w: 1.871, label: 'Cost Efficiency', dot: SILVER },
  ];
  cards.forEach((c) => {
    roundRect(s, c.x, c.y, c.w, 0.887, 0.161, { color: WHITE }, { shadow: CARD_SHADOW });
    text(s, c.label, { x: c.x + 0.136, y: c.y + 0.161, w: 1.842, h: 0.286, fontSize: 10.5, fontFace: HEAD });
    dotLabel(s, c.x + 0.245, c.y + 0.504, '$75,230', c.dot);
  });
}

function slide05(s) {
  // Full-bleed photo under a near-opaque dark scrim.
  s.background = { color: DUSK };
  photo(s, 0, 0, 13.333, 7.5, { radius: 0, fill: DUSK, labelColor: '5C5C5C', labelSize: 12 });
  brandLockup(s, 11.442, 0.448, 8, SILVER);
  pageNumber(s, 5, '595959');
  eyebrow(s, 1.182, 1.037, 'Profit growth', WHITE);
  title(s, 0.828, 1.498, 7.442, 1.447, 'Industry Trends in \nGlobal Markets', 40, { color: WHITE });
  text(s, [{ text: '$', options: { color: ORANGE } }, { text: '18.250', options: { color: WHITE } }],
    { x: 0.778, y: 4.467, w: 4.156, h: 1.01, fontSize: 54, fontFace: HEAD, valign: 'top' });
  copy(s, 0.828, 5.618, 6.519, 0.596, 'Profit growth boosts revenue, reduces costs, and ensures long\n-term sustainable business success achievement advancement.', { color: WHITE });

  const stats = [
    { y: 1.286, pct: '15%', bg: WHITE, icon: ICON.money, fg: ORANGE, cy: 1.327 },
    { y: 4.514, pct: '39%', bg: ORANGE, icon: ICON.hand, fg: WHITE, cy: 4.556 },
  ];
  stats.forEach((st) => {
    iconCircle(s, 7.586, st.cy, 0.758, st.bg, st.icon, st.fg);
    title(s, 8.784, st.y, 4.143, 0.841, st.pct, 44, { color: WHITE });
    copy(s, 8.781, st.y + 0.841, 3.572, 0.859, 'The profit growth drives revenue and \nensures sustainable success and\nprofitability growth.', { color: WHITE });
  });
}

function slide06(s) {
  s.background = { color: PAPER };
  photo(s, 0, 4.527, 3.492, 2.973, { radius: 0.3 });
  photo(s, 9.917, 0, 3.417, 7.5, { radius: 0 });
  eyebrow(s, 1.182, 1.284, 'Profit growth');
  title(s, 0.852, 1.779, 9.064, 1.313, 'Cost Optimization Initiatives \nin Business Operation', 36);
  chipRow(s, 0.992, 3.478, [
    { label: 'Profit growth', w: 1.26 },
    { label: 'Revenue', w: 1.107 },
    { label: 'Profitability', w: 1.26 },
    { label: 'Operational', w: 1.26 },
  ], 0.312, 10, 0.161);
  iconCircle(s, 4.04, 4.869, 0.613, ORANGE, ICON.money, WHITE);
  title(s, 4.9, 4.822, 2.63, 0.707, '$15.230', 36);
  copy(s, 3.928, 5.756, 6.073, 0.859, 'Profit growth refers to the increase in a company\u2019s net profit over \na specific period, indicating improved financial performance and \nbusiness efficiency. ');
}

function slide07(s) {
  s.background = { color: PAPER };
  masterLogo(s);
  pageNumber(s, 7, SILVER);
  photo(s, 5.472, 0, 4.309, 7.5, { radius: 0 });
  eyebrow(s, 1.168, 1.159, 'Profit growth');
  title(s, 0.83, 1.657, 6.801, 1.313, 'Opportunities \nfor Growth', 36);
  title(s, 0.826, 4.32, 4.156, 0.909, '95.2K', 48);
  copy(s, 0.844, 5.266, 5.741, 0.596, 'Profit growth boosts revenue, reduces costs, and \nensures long-term sustainable business success.');
  chipRow(s, 0.957, 6.083, [
    { label: 'Profit growth', w: 1.038, fill: ORANGE, color: WHITE },
    { label: 'Revenue', w: 0.852, fill: INK, color: WHITE },
    { label: 'Profitability', w: 1.038 },
  ], 0.258, 8);

  const stats = [
    { y: 1.842, pct: '70%', bg: ORANGE, icon: ICON.money, cy: 1.233 },
    { y: 4.769, pct: '35%', bg: INK, icon: ICON.hand, cy: 4.16 },
  ];
  stats.forEach((st) => {
    iconCircle(s, 10.279, st.cy, 0.463, st.bg, st.icon, WHITE);
    title(s, 10.163, st.y, 1.758, 0.64, st.pct, 32);
    copy(s, 10.163, st.y + 0.64, 2.764, 0.859, 'Profit growth drives revenue \nand strengthens long-term business success.');
  });
}

function slide08(s) {
  s.background = { color: PAPER };
  masterLogo(s);
  pageNumber(s, 8, SILVER);
  eyebrow(s, 1.182, 1.3, 'Profit growth');
  title(s, 0.818, 1.76, 8.067, 0.707, 'Performance Metrics', 36);
  copy(s, 7.106, 1.515, 6.069, 0.859, 'A profit growth represents a company\u2019s capability to boost \nrevenue, streamline operations, reduce costs, and secure \nlong-term sustainable success.');

  // Horizontal stacked bar chart: one row per metric.
  const rows = [
    {
      label: 'Revenue', labelColor: ASH, y: 3.053, labelY: 3.132,
      bars: [[2.595, 2.319, MIST, 0], [4.981, 0.998, MIST, 0], [6.048, 1.405, MIST, 50]],
      badge: { x: 6.947, y: 3.191, pct: '49%', dot: SILVER },
    },
    {
      label: 'Profit', labelColor: INK, y: 4.001, labelY: 4.096,
      bars: [[2.595, 5.619, ORANGE, 0], [8.281, 1.585, PEACH, 0], [9.934, 1.011, PEACH, 50]],
      badge: { x: 10.44, y: 4.151, pct: '77%', dot: ORANGE },
    },
    {
      label: 'Annual', labelColor: INK, y: 4.949, labelY: 5.044,
      bars: [[2.595, 4.443, INK, 0], [7.106, 1.827, STEEL, 0], [9.001, 1.405, STEEL, 50]],
      badge: { x: 9.9, y: 5.098, pct: '72%', dot: INK },
    },
    {
      label: 'Return', labelColor: ASH, y: 5.896, labelY: 5.991,
      bars: [[2.595, 3.075, MIST, 0], [5.738, 0.823, MIST, 0], [6.628, 1.585, MIST, 50]],
      badge: { x: 7.708, y: 6.046, pct: '52%', dot: SILVER },
    },
  ];
  rows.forEach((r) => {
    r.bars.forEach((b) => {
      roundRect(s, b[0], r.y, b[1], 0.704, 0.147, { color: b[2], transparency: b[3] });
    });
    roundRect(s, 0.936, r.labelY, 1.323, 0.514, 0.196, { color: MIST, transparency: 50 });
    text(s, r.label, { x: 0.993, y: r.labelY + 0.105, w: 1.209, h: 0.303, fontSize: 12, bold: true, align: 'center', color: r.labelColor });
    roundRect(s, r.badge.x, r.badge.y, 1.206, 0.404, 0.202, { color: WHITE }, { shadow: BADGE_SHADOW });
    text(s, r.badge.pct, { x: r.badge.x + 0.249, y: r.badge.y + 0.05, w: 0.9, h: 0.303, fontSize: 12, bold: true, align: 'center', valign: 'middle' });
    oval(s, r.badge.x + 0.143, r.badge.y + 0.119, 0.165, { color: r.badge.dot });
  });
}

function slide09(s) {
  s.background = { color: MIDNIGHT };
  brandLockup(s, 11.442, 0.448, 8, SILVER);
  pageNumber(s, 9, '595959');
  title(s, 0.795, 0.618, 10.532, 2.036, '30', 115, { color: WHITE });
  chip(s, 3.042, 1.86, 1.198, 0.345, 'minutes', ORANGE, WHITE, 12);
  copy(s, 0.845, 2.535, 5.741, 0.859, 'Profit growth refers to the increase in a company\u2019s net profit over \na specific period, indicating improved financial performance and \nbusiness efficiency. ', { color: WHITE });
  brandMark(s, 9.466, 3.689, 0.318, 0.254);
  text(s, 'PROFITERA GLOBAL', { x: 9.919, y: 3.631, w: 2.387, h: 0.37, fontSize: 16, fontFace: HEAD, color: WHITE, align: 'right' });
  text(s, [
    { text: 'A ', options: { color: WHITE } },
    { text: 'Short Break ', options: { color: ORANGE, breakLine: true } },
    { text: 'Before We Continue', options: { color: WHITE } },
  ], { x: 1.774, y: 4.572, w: 10.532, h: 2.121, fontSize: 60, fontFace: HEAD, align: 'right', valign: 'top' });
}

function slide10(s) {
  s.background = { color: PAPER };
  masterLogo(s);
  pageNumber(s, 10, SILVER);
  photo(s, 5.891, 0, 7.442, 4.859, { radius: 0.25 });
  title(s, 0.437, 1.066, 4.156, 1.01, '95.2K', 54, { align: 'right' });
  iconCircle(s, 4.773, 1.292, 0.557, ORANGE, ICON.money, WHITE);
  copy(s, 0.749, 2.183, 4.695, 0.859, 'Profit growth refers to the increase in a company\u2019s \nprofit over a specific period, indicating improved performance and efficiency.', { align: 'right' });
  eyebrow(s, 1.182, 4.397, 'Profit growth');
  title(s, 0.828, 5.037, 7.442, 1.604, 'Historical \nProfit Performance', 44, { lineSpacing: 55 });
  copy(s, 6.761, 5.908, 5.741, 0.596, 'Profit growth boosts revenue, reduces costs, and \nensures long-term sustainable business success.', { align: 'right' });
  chipRow(s, 9.71, 5.393, [
    { label: 'Profitability', w: 1.26 },
    { label: 'Operational', w: 1.26, fill: ORANGE, color: WHITE },
  ], 0.312, 10, 0.161);
}

function slide11(s) {
  s.background = { color: PAPER };
  masterLogo(s);
  pageNumber(s, 11, SILVER);
  eyebrow(s, 1.182, 1.3, 'Profit growth');
  title(s, 0.846, 1.848, 6.801, 1.313, 'Cost Optimization Initiatives', 36);
  title(s, 0.835, 4.264, 4.156, 0.909, '$37.510', 48);
  chip(s, 3.57, 4.702, 1.038, 0.258, 'Profit growth', ORANGE, WHITE, 8);
  copy(s, 0.885, 5.32, 5.995, 0.859, 'A profit growth represents a company\u2019s capability to boost \nrevenue, streamline operations, reduce costs, and secure \nlong-term sustainable success.');

  const rows = [
    { y: 0.872, pct: '39.5%', dotY: 1.684, textY: 1.483 },
    { y: 4.007, pct: '57.2%', dotY: 4.82, textY: 4.619 },
  ];
  rows.forEach((r) => {
    photo(s, 6.381, r.y, 2.621, 2.621, { shape: 'ellipse' });
    oval(s, 9.413, r.dotY, 0.306, { color: ORANGE, transparency: 80 });
    oval(s, 9.468, r.dotY + 0.057, 0.196, { color: ORANGE });
    title(s, 9.906, r.textY, 2.63, 0.707, r.pct, 36);
    copy(s, 9.355, r.textY + 0.802, 3.572, 0.596, 'Profit growth drives revenue and \nensures sustainable success.');
  });
}

function slide12(s) {
  s.background = { color: PAPER };
  brandLockup(s, 11.442, 0.448, 8, INK);
  // Two tilted phone mockups.
  device(s, 6.556, -0.951, 3.935, 5.565, 0.55, 22);
  device(s, 9.659, 2.904, 3.935, 5.09, 0.55, -22);

  eyebrow(s, 1.182, 1.037, 'Profit growth');
  title(s, 0.828, 1.498, 7.442, 1.447, 'Industry Trends in \nGlobal Markets', 40);
  chipRow(s, 0.96, 3.271, [
    { label: 'Profit growth', w: 1.038 },
    { label: 'Revenue', w: 0.852 },
    { label: 'Profitability', w: 1.038 },
    { label: 'Profitability', w: 1.038 },
  ], 0.258, 8);
  title(s, 0.835, 4.687, 4.156, 0.909, '$25,130', 48);
  chip(s, 3.591, 5.125, 1.038, 0.258, 'Profit growth', ORANGE, WHITE, 8);

  const notes = [
    { x: 0.963, y: 5.89, tx: 1.182, ty: 5.82, w: 3.51, label: 'Profit growth drives sustainable business success.' },
    { x: 0.963, y: 6.263, tx: 1.182, ty: 6.194, w: 1.71, label: 'Increase profit margins' },
    { x: 3.104, y: 6.263, tx: 3.324, ty: 6.194, w: 1.967, label: 'Reduce operational costs' },
  ];
  notes.forEach((n) => {
    oval(s, n.x, n.y, 0.131, { color: ORANGE });
    text(s, n.label, { x: n.tx, y: n.ty, w: n.w, h: 0.269, fontSize: 10, wrap: false });
  });
  copy(s, 6.198, 5.88, 3.572, 0.596, 'Profit growth drives revenue and \nensures sustainable success.', { align: 'right' });
  iconCircle(s, 9.389, 1.498, 1.028, INK, ICON.money, WHITE);
  iconCircle(s, 10.078, 5.807, 0.742, ORANGE, ICON.hand, WHITE);
}

function slide13(s) {
  s.background = { color: NIGHT };
  brandLockup(s, 11.442, 0.448, 8, MIST);
  pageNumber(s, 13, '595959');
  // Faint concentric rings sweeping out of the lower-left corner.
  [3.483, 5.25, 7.115, 9.03, 11.357].forEach((d) => {
    s.addShape('ellipse', {
      x: 0.618 - d / 2, y: 6.968 - d / 2, w: d, h: d,
      fill: { color: NIGHT, transparency: 100 }, line: { color: '4A4A4A', width: 0.75, transparency: 55 },
    });
  });

  const cards = [
    { x: 1.772, y: 1.05, label: 'Future Strategies', value: '$93,230', px: 2.056, py: 1.274 },
    { x: 3.395, y: 3.044, label: 'Strategic Investment', value: '$75,230', px: 3.679, py: 3.268 },
    { x: 1.125, y: 5.039, label: 'Performance Analysis', value: '$32,230', px: 1.409, py: 5.262 },
  ];
  cards.forEach((c) => {
    roundRect(s, c.x, c.y, 3.52, 1.191, 0.15, { color: PAPER }, { shadow: CARD_SHADOW });
    photo(s, c.px, c.py, 0.744, 0.744, { shape: 'ellipse', labelSize: 7 });
    text(s, c.label, { x: c.x + 1.218, y: c.y + 0.268, w: 2.17, h: 0.303, fontSize: 12, fontFace: HEAD });
    dotLabel(s, c.x + 1.293, c.y + 0.682, c.value, ORANGE, 12);
  });

  title(s, 7.271, 1.066, 4.156, 1.01, '95.2K', 54, { color: WHITE, align: 'right' });
  iconCircle(s, 11.606, 1.292, 0.557, ORANGE, ICON.money, WHITE);
  copy(s, 6.667, 2.183, 5.61, 0.859, 'A profit growth represents a company\u2019s capability to boost \nrevenue, streamline operations, reduce costs, and secure \nlong-term sustainable success.', { align: 'right', color: WHITE });
  chipRow(s, 9.049, 3.383, [
    { label: 'Profit growth', w: 1.038, fill: INK, color: WHITE },
    { label: 'Revenue', w: 0.852, fill: ORANGE, color: WHITE },
    { label: 'Profitability', w: 1.038, fill: INK, color: WHITE },
  ], 0.258, 8);
  eyebrow(s, 7.38, 4.948, 'Profit growth', WHITE);
  title(s, 7.044, 5.496, 6.801, 1.178, 'Historical Profit \nPerformance Analysis', 32, { color: WHITE });
}

function slide14(s) {
  s.background = { color: PAPER };
  masterLogo(s);
  pageNumber(s, 14, SILVER);
  // Laptop mockup assembled from flat shapes: lid, screen, hinge lip, base.
  roundRect(s, 1.39, 0.81, 3.81, 2.75, 0.06, { color: '303030' });
  photo(s, 1.486, 0.896, 3.653, 2.479, { radius: 0.02 });
  roundRect(s, 1.18, 3.5, 4.24, 0.13, 0.05, { color: '8E8E8E' });
  roundRect(s, 1.39, 3.76, 3.83, 2.81, 0.09, { color: 'BEBEBE' });
  roundRect(s, 1.73, 4.0, 3.15, 1.33, 0.03, { color: '9A9A9A' });
  roundRect(s, 2.66, 5.5, 1.29, 0.83, 0.03, { color: 'A8A8A8' });

  eyebrow(s, 6.487, 1.195, 'Profit growth');
  title(s, 6.132, 1.657, 6.409, 1.447, 'Industry Trends in \nGlobal Markets', 40);
  title(s, 6.167, 4.006, 4.156, 0.909, '$25,130', 48);
  chip(s, 8.902, 4.444, 1.038, 0.258, 'Profit growth', ORANGE, WHITE, 8);
  copy(s, 6.169, 5.197, 6.409, 0.596, 'Profit growth refers to the increase in a company\u2019s net profit over a specific \nperiod, indicating improved financial performance and business efficiency. ');
  [
    { x: 6.285, tx: 6.505, w: 2.02, label: 'Increase profit margins' },
    { x: 8.836, tx: 9.055, w: 2.25, label: 'Reduce operational costs' },
  ].forEach((n) => {
    oval(s, n.x, 6.088, 0.131, { color: ORANGE });
    text(s, n.label, { x: n.tx, y: 6.002, w: n.w, h: 0.303, fontSize: 12, wrap: false });
  });
}

function slide15(s) {
  s.background = { color: PAPER };
  brandLockup(s, 11.442, 0.448, 8, INK);
  // Landscape phone mockup bleeding off the top-left corner.
  roundRect(s, -0.9, -0.85, 6.65, 4.85, 0.42, { color: FRAME });
  photo(s, -0.85, -0.8, 6.4, 4.6, { radius: 0.38 });
  photo(s, 8.0, 4.806, 5.333, 2.694, { radius: 0.2 });

  eyebrow(s, 7.021, 1.124, 'Profit growth');
  title(s, 6.667, 1.642, 6.04, 1.313, 'The Projected Profit Forecast Analysis', 36);
  chipRow(s, 6.804, 3.228, [
    { label: 'Profit growth', w: 1.038 },
    { label: 'Revenue', w: 0.852 },
    { label: 'Profitability', w: 1.038 },
    { label: 'Profitability', w: 1.038 },
  ], 0.258, 8);
  title(s, 0.778, 4.65, 4.156, 1.01, '$25.130', 54);
  chip(s, 3.831, 5.163, 1.038, 0.258, 'Profit growth', ORANGE, WHITE, 8);
  copy(s, 0.828, 5.75, 6.519, 0.859, 'The profit growth represents a company\u2019s capability to boost revenue, streamline operations, and secure long-term sustainable success and profitability expansion.');
}

function slide16(s) {
  s.background = { color: PAPER };
  photo(s, 7.635, 0.403, 5.292, 6.694, { radius: 0.35 });
  title(s, 2.04, 0.822, 4.156, 1.01, '95.2K', 54, { align: 'right' });
  iconCircle(s, 6.375, 1.048, 0.557, ORANGE, ICON.money, WHITE);
  copy(s, 1.436, 1.939, 5.61, 0.859, 'A profit growth represents a company\u2019s capability to boost \nrevenue, streamline operations, reduce costs, and secure \nlong-term sustainable success.', { align: 'right' });
  chipRow(s, 3.818, 3.139, [
    { label: 'Profit growth', w: 1.038 },
    { label: 'Revenue', w: 0.852 },
    { label: 'Profitability', w: 1.038 },
  ], 0.258, 8);
  eyebrow(s, 1.182, 4.414, 'Profit growth');
  title(s, 0.846, 4.962, 6.801, 1.717, 'Projected Profit \nForecast for Lasting \nBusiness Growth Globally', 32);
}

function slide17(s) {
  s.background = { color: DUSK };
  photo(s, 0, 0, 13.333, 7.5, { radius: 0, fill: DUSK, labelColor: '5C5C5C', labelSize: 12 });
  text(s, 'Driving efficiency and long-term success', { x: 2.896, y: 2.029, w: 7.542, h: 0.404, fontSize: 18, align: 'center', color: WHITE });
  text(s, [
    { text: 'Profit growth ', options: { color: ORANGE } },
    { text: 'reflects strategy, efficiency,', options: { color: WHITE, breakLine: true } },
    { text: 'and long-term ', options: { color: WHITE } },
    { text: 'business success', options: { color: ORANGE, breakLine: true } },
    { text: 'across industries.', options: { color: WHITE } },
  ], { x: 0.406, y: 2.82, w: 12.521, h: 2.374, fontSize: 36, bold: true, align: 'center', valign: 'top', lineSpacingMultiple: 1.3 });
  text(s, 'www.yourwebsite.com', { x: 3.308, y: 6.437, w: 6.717, h: 0.37, fontSize: 16, align: 'center', color: ORANGE });
}

function slide18(s) {
  s.background = { color: PAPER };
  photo(s, 0, 0, 5.163, 7.5, { radius: 0, fill: DUSK_SOFT, labelColor: '646464', labelSize: 12 });
  rect(s, 5.163, 0, 4.085, 7.5, { color: WHITE });
  rect(s, 9.248, 0, 4.085, 7.5, { color: MIST });

  eyebrow(s, 1.182, 1.037, 'Profit growth', WHITE);
  title(s, 0.828, 1.498, 4.966, 1.447, 'Our Vision\nand Mission', 40, { color: WHITE });
  chip(s, 0.951, 5.466, 1.64, 0.386, 'Profit growth', ORANGE, WHITE, 12);
  copy(s, 0.848, 6.002, 3.572, 0.596, 'The profit growth drives revenue and ensures sustainable success for all.', { color: WHITE });

  const cols = [
    { x: 5.672, pctX: 5.878, pct: '90%', title: 'Our Vision', badge: INK, icon: ICON.money, ovalX: 5.788 },
    { x: 9.638, pctX: 9.832, pct: '75%', title: 'Our Mission', badge: ORANGE, icon: ICON.hand, ovalX: 9.754 },
  ];
  cols.forEach((c) => {
    title(s, c.pctX, 1.368, 2.673, 1.01, c.pct, 54, { align: 'right' });
    copy(s, c.pctX - 0.018, 2.35, 2.676, 0.596, 'The profit growth drives \nsustainable business.', { align: 'right' });
    iconCircle(s, c.ovalX, 4.701, 0.463, c.badge, c.icon, WHITE);
    title(s, c.x, 5.407, 2.673, 0.505, c.title, 24);
    copy(s, c.x, 6.002, 3.118, 0.596, 'Profit growth drives revenue and strengthens long-term business.');
  });
}

function slide19(s) {
  s.background = { color: PAPER };
  photo(s, 7.014, 0, 6.319, 3.75, { radius: 0.25 });
  pageNumber(s, 19, SILVER);
  eyebrow(s, 1.182, 0.959, 'Profit growth');
  title(s, 0.828, 1.421, 7.442, 1.313, 'Our Contact \nInformation Details', 36);
  chipRow(s, 0.96, 3.03, [
    { label: 'Profit growth', w: 1.038 },
    { label: 'Revenue', w: 0.852 },
    { label: 'Profitability', w: 1.038 },
  ], 0.258, 8);
  copy(s, 8.208, 4.316, 4.272, 0.859, [
    { text: 'Profit growth ', options: { bold: true } },
    { text: 'refers to the increase in profit over ', options: { breakLine: true } },
    { text: 'a period, indicating ' },
    { text: 'improved performance ', options: { bold: true } },
    { text: 'and ' },
    { text: 'efficiency overall', options: { bold: true } },
    { text: '.' },
  ], { align: 'right' });

  const contacts = [
    { x: 0.959, tx: 0.835, vx: 0.853, bg: ORANGE, icon: ICON.phone, label: 'Phone Number', value: '+123-456-7890' },
    { x: 3.685, tx: 3.561, vx: 3.579, bg: INK, icon: ICON.mail, label: 'Email Address', value: 'addmail@yourmail.com' },
    { x: 6.41, tx: 6.286, vx: 6.305, bg: INK, icon: ICON.globe, label: 'Website Link', value: 'www.yourwebsite.com' },
  ];
  contacts.forEach((c) => {
    iconCircle(s, c.x, 5.032, 0.483, c.bg, c.icon, WHITE);
    text(s, c.label, { x: c.tx, y: 5.739, w: 2.698, h: 0.404, fontSize: 18, fontFace: HEAD });
    text(s, c.value, { x: c.vx, y: 6.238, w: 2.698, h: 0.303, fontSize: 12 });
  });
}

function slide20(s) {
  s.background = { color: PAPER };
  masterLogo(s);
  copy(s, 7.0, 2.033, 5.505, 0.859, 'Profit growth refers to the increase in a company\u2019s net profit over \na specific period, indicating improved financial performance and \nbusiness efficiency. ', { align: 'right' });
  iconCircle(s, 11.651, 1.011, 0.742, ORANGE, ICON.trend, WHITE);
  brandLockup(s, 0.939, 3.428, 16, INK);
  text(s, [
    { text: 'Thanks', options: { color: ORANGE } },
    { text: ' for Your ', options: { color: INK, breakLine: true } },
    { text: 'Attention and Support', options: { color: INK } },
  ], { x: 0.828, y: 4.369, w: 10.532, h: 2.121, fontSize: 60, fontFace: HEAD, valign: 'top' });
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */

const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE';
  pptx.title = 'Driving Sustainable Profit Growth';
  pptx.company = 'PROFITERA GLOBAL';
  BUILDERS.forEach((fn) => fn(pptx.addSlide()));
  return pptx.writeFile({
    fileName: path.join(__dirname, '16468c51-d384-44f9-a16b-a866ac16808d_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
