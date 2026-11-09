/*
 * Standalone pptxgenjs recreation of "Project Proposal" (18 slides, 13.333 x 7.5 in).
 * Raster photographs in the source deck are replaced by programmatic placeholders.
 *
 *   node 0802fd70-20ce-4d79-b85b-44fbdb11be11_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const C = {
  dk2: '262626',        // body / heading text
  white: 'FFFFFF',
  a1: '3C53A3',         // deep blue
  a2: '98B8E1',         // light blue
  a3: '8BC7BB',         // green
  a4: '50556B',         // slate
  a1dark: '2C3E7A',
  a4dark: '3C3F50',
  rule: 'DADBE2',       // hairline rules / soft grey
  mute: 'B5B8C6',       // muted grey text
  panel: 'EAF0F8',      // pale blue panel
  panel2: 'D5E2F2',
  grey1: 'F2F2F2',
  grey2: 'E8E8E8',
  grey3: 'D0CECE',
  grey4: '7F7F7F',
  ink: '3B3B3B',
};

const F = {
  reg: 'Roboto',
  light: 'Roboto Light',
  med: 'Roboto Medium',
  thin: 'Roboto Thin',
  icon: 'DejaVu Sans',   // font that carries the pictographic glyphs used as icons
};

// Glyphs standing in for the deck's vector icons.
const IC = {
  plane: '\u2708', envelope: '\u2709', phone: '\u260E', gear: '\u2699',
  star: '\u2605', crown: '\u265B', check: '\u2713', target: '\u25CE',
  ring: '\u25C9', flag: '\u2691', diamond: '\u25C6', square: '\u25A3',
  lines: '\u25A4', tree: '\u2663', badge: '\u273B', dot: '\u25CF',
  bell: '\u2740', sparkle: '\u2756', triangle: '\u25B2',
};

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

/* ------------------------------------------------------------------ helpers */

// Text-box defaults mirroring the source deck (0.1"/0.05" insets, top anchored).
const INSET = [7.2, 7.2, 3.6, 3.6]; // [left, right, bottom, top] in points

function T(s, x, y, w, h, text, o) {
  s.addText(text, Object.assign({
    x: x, y: y, w: w, h: h,
    valign: 'top', isTextBox: true, margin: INSET,
    fontFace: F.reg, fontSize: 10, color: C.dk2, align: 'left',
  }, o || {}));
}

// One paragraph inside a multi-paragraph text box.
function P(text, o) {
  return { text: text, options: Object.assign({ breakLine: true }, o || {}) };
}

// Bulleted paragraphs (bullet options have to sit on every paragraph).
function bulletList(items) {
  return items.map(function (t) {
    return P(t, { bullet: { characterCode: '2022', indent: 13.5 } });
  });
}

function rect(s, x, y, w, h, fill, o) {
  s.addShape('rect', Object.assign({ x: x, y: y, w: w, h: h, fill: fill }, o || {}));
}

function ellipse(s, x, y, w, h, fill, o) {
  s.addShape('ellipse', Object.assign({ x: x, y: y, w: w, h: h, fill: fill }, o || {}));
}

function pill(s, x, y, w, h, fill, radius, o) {
  s.addShape('roundRect', Object.assign(
    { x: x, y: y, w: w, h: h, fill: fill, rectRadius: radius }, o || {}));
}

function hline(s, x, y, w, color, width) {
  s.addShape('line', { x: x, y: y, w: w, h: 0, line: { color: color, width: width } });
}

function vline(s, x, y, h, color, width) {
  s.addShape('line', { x: x, y: y, w: 0, h: h, line: { color: color, width: width } });
}

function poly(s, x, y, w, h, pts, fill, o) {
  const points = pts.map(function (p, i) { return { x: p[0], y: p[1], moveTo: i === 0 }; });
  points.push({ close: true });
  s.addShape('custGeom', Object.assign(
    { x: x, y: y, w: w, h: h, fill: fill, points: points }, o || {}));
}

// Curved outline from a compact normalised path: ['M',x,y] / ['L',x,y] /
// ['C',x1,y1,x2,y2,x,y] / ['Z'], all coordinates in 0..1 of the bounding box.
function curveShape(s, x, y, w, h, ops, fill, o) {
  const pts = [];
  ops.forEach(function (op) {
    const k = op[0];
    if (k === 'M') pts.push({ x: op[1] * w, y: op[2] * h, moveTo: true });
    else if (k === 'L') pts.push({ x: op[1] * w, y: op[2] * h });
    else if (k === 'C') pts.push({ x: op[5] * w, y: op[6] * h,
      curve: { type: 'cubic', x1: op[1] * w, y1: op[2] * h, x2: op[3] * w, y2: op[4] * h } });
    else pts.push({ close: true });
  });
  s.addShape('custGeom', Object.assign(
    { x: x, y: y, w: w, h: h, fill: fill, points: pts }, o || {}));
}

// Icon glyph centred on a box.
function icon(s, x, y, size, glyph, color, fontSize) {
  T(s, x, y, size, size, glyph, {
    fontFace: F.icon, fontSize: fontSize || Math.round(size * 52),
    color: color, align: 'center', valign: 'middle', margin: 0,
  });
}

function mix(a, b, t) {
  const h = function (c, i) { return parseInt(c.substr(i * 2, 2), 16); };
  let out = '';
  for (let i = 0; i < 3; i++) {
    const v = Math.round(h(a, i) + (h(b, i) - h(a, i)) * t);
    out += ('0' + v.toString(16)).slice(-2);
  }
  return out.toUpperCase();
}

// Colour at position t (0..1) of a [[stop, hex], ...] ramp.
function rampColor(stops, t) {
  for (let i = 1; i < stops.length; i++) {
    if (t <= stops[i][0] || i === stops.length - 1) {
      const a = stops[i - 1], b = stops[i];
      const f = b[0] === a[0] ? 0 : (t - a[0]) / (b[0] - a[0]);
      return mix(a[1], b[1], Math.max(0, Math.min(1, f)));
    }
  }
  return stops[0][1];
}

// Linear gradient background faked with banded rectangles.
// `deg` is measured clockwise from the +x axis (90 = top to bottom).
function linearGradient(s, deg, stops, bands) {
  const n = bands || 44;
  const rad = deg * Math.PI / 180;
  const dx = Math.cos(rad), dy = Math.sin(rad);
  const corners = [[0, 0], [SLIDE_W, 0], [0, SLIDE_H], [SLIDE_W, SLIDE_H]];
  const proj = corners.map(function (c) { return c[0] * dx + c[1] * dy; });
  const pMin = Math.min.apply(null, proj), pMax = Math.max.apply(null, proj);
  const pMid = (SLIDE_W / 2) * dx + (SLIDE_H / 2) * dy;
  const step = (pMax - pMin) / n;
  const len = 2 * (SLIDE_W + SLIDE_H);
  for (let i = 0; i < n; i++) {
    const d = pMin + (i + 0.5) * step;
    const cx = SLIDE_W / 2 + dx * (d - pMid);
    const cy = SLIDE_H / 2 + dy * (d - pMid);
    const h = step * 1.6;
    rect(s, cx - len / 2, cy - h / 2, len, h,
      { color: rampColor(stops, (i + 0.5) / n) }, { rotate: (deg + 90) % 360 });
  }
}

// Radial (elliptical) gradient faked with concentric opaque ellipses.
function radialGradient(s, stops, bands) {
  const n = bands || 40;
  const rMax = 1.47;      // reaches past the slide corners
  const cx = SLIDE_W / 2, cy = SLIDE_H / 2;
  rect(s, 0, 0, SLIDE_W, SLIDE_H, { color: rampColor(stops, 1) });
  for (let i = n; i >= 1; i--) {
    const rho = rMax * i / n;
    const a = cx * rho, b = cy * rho;
    ellipse(s, cx - a, cy - b, a * 2, b * 2, { color: rampColor(stops, rho / rMax) });
  }
}

// Master furniture: "PROPOSAL" tag, year block and the hamburger mark.
function chrome(s, hamburgerColor) {
  T(s, 10.458, 0.544, 2.448, 0.236, 'PROPOSAL', { fontSize: 8, align: 'right' });
  T(s, 11.285, 6.578, 1.554, 0.411, [
    P('2025', { fontFace: F.med }),
    P('PRESENTATION', { fontFace: F.reg }),
  ], { fontSize: 8, align: 'right', lineSpacingMultiple: 1.2 });
  [6.798, 6.854, 6.91].forEach(function (y) {
    hline(s, 0.6, y, 0.249, hamburgerColor || C.rule, 2.25);
  });
}

// Section heading used on nearly every slide.
function heading(s, x, y, w, text, o) {
  T(s, x, y, w, 0.841, text, Object.assign({ fontSize: 44, bold: true }, o || {}));
}

// Stand-in for one of the deck's four raster photographs.
function photoPlaceholder(s, x, y, w, h, o) {
  const opt = o || {};
  rect(s, x, y, w, h, { color: opt.fill || C.grey2 },
    { line: { color: opt.line || C.grey3, width: 1 } });
  T(s, x, y + h / 2 - 0.2, w, 0.4, '[image]', {
    fontSize: 11, color: opt.label || C.grey4, align: 'center', valign: 'middle', margin: 0,
  });
}

/* ------------------------------------------------------- shared text recipes */

const BODY = { fontFace: F.light, fontSize: 10, lineSpacingMultiple: 1.5 };
const SUBHEAD = { fontFace: F.med, fontSize: 18 };
const LEAD = { fontFace: F.reg, fontSize: 28 };

const LOREM = {
  contents: 'Parish, looked has attachment in a to is in there go frequently to gay terminated you.',
  card: 'PLACEHOLDER',
  swot: 'Certainly elsewhere allowance address farther six hearted the its is as hundred to.',
  project: 'PLACEHOLDER',
  goals: 'Certainly elsewhere my the allowance at address farther six hearted.',
  stages: 'Parish, looked has attachment in to is in there go frequently.',
  timeline: 'Parish, gas attachment in a to is there frequently to gay terminated.',
  partners: 'Parish, looked has attachment in a to terminated you greater nay prudent looked has great it.',
  team: 'Parish, looked has attachment in a to is in there there go frequently to terminated in the.',
  peek: 'Parish, looked has great attachment in get to is there go here.',
  long: 'PLACEHOLDER'
    + 'husband in are securing off  it there occasion  daughter replying held in that feel his see so '
    + 'own to yet. Strangers the us beast artist.',
  footnote: 'Certainly elsewhere my do allowance at. the address farther six too hearted hundred to '
    + 'towards husband in are securing off occasion remember daughter.',
  metric: ['husband in are securing off  it there occasion  daughter replying',
    'held in that feel his see so own to yet'],
};

/* ------------------------------------------------------------- slide 1: cover */

function slide01(p) {
  const s = p.addSlide();
  linearGradient(s, 90, [[0, C.white], [1, C.rule]]);
  chrome(s, C.mute);

  rect(s, 10.865, 0.544, 2.041, 0.236, { color: C.white });
  T(s, 10.865, 0.544, 2.041, 0.236, 'example@gmail.com', { fontSize: 8, align: 'right' });

  T(s, 1.04, 1.303, 4.947, 1.749, [
    { text: 'Project ', options: { bold: true } },
    { text: 'Proposal', options: { fontFace: F.light } },
  ], { fontSize: 61, align: 'right', lineSpacingMultiple: 0.8 });

  T(s, 1.069, 3.063, 4.902, 0.404, 'presentation template',
    { fontSize: 18, fontFace: F.thin, color: C.mute, align: 'right' });

  T(s, 2.613, 4.14, 3.35, 0.707, [
    P('Prepared by:'),
    P('John Doe', { fontFace: F.med }),
    P('CEO template design co ', { fontFace: F.light }),
  ], { fontSize: 12, align: 'right' });

  hline(s, 0, 6.019, 4.429, C.rule, 0.75);
  T(s, 4.745, 5.706, 1.762, 0.404, [
    P('Issue date:', { fontFace: F.light }),
    P('14/09/2024'),
  ], { fontSize: 9 });
}

/* ------------------------------------------------- slide 2: table of contents */

function slide02(p) {
  const s = p.addSlide();
  chrome(s);
  heading(s, 0.462, 0.573, 6.204, 'Contents List');

  const rows = [0, 1, 2];
  const ys = [2.131, 3.679, 5.227];      // title baseline of each row
  const dots = [C.a1, C.a2, C.a3];

  const left = ['Our Team', 'Project Stages', 'Sneak Peek'];
  const right = ['About The Project', 'Major Requirements ', 'Project Goals'];
  const leftNums = ['06', '05', '04'];
  const rightNums = ['01', '02', '03'];

  rows.forEach(function (i) {
    const y = ys[i];
    // left column (right aligned toward the bullet circles)
    ellipse(s, 4.283, y + 0.119, 0.851, 0.851, { color: dots[i] },
      { line: { color: C.white, width: 2.25 } });
    T(s, 4.385, y + 0.342, 0.648, 0.404, leftNums[i],
      { fontSize: 18, color: C.white, align: 'center' });
    T(s, 1.217, y, 2.731, 0.404, left[i], Object.assign({ align: 'right' }, SUBHEAD));
    T(s, 0.583, y + 0.467, 3.365, 0.578, LOREM.contents,
      Object.assign({ align: 'right' }, BODY));

    // right column
    ellipse(s, 8.2, y + 0.119, 0.851, 0.851, { color: dots[i] },
      { line: { color: C.white, width: 2.25 } });
    T(s, 8.301, y + 0.342, 0.648, 0.404, rightNums[i],
      { fontSize: 18, color: C.white, align: 'center' });
    T(s, 9.385, y, 2.881, 0.404, right[i], SUBHEAD);
    T(s, 9.385, y + 0.467, 3.49, 0.578, LOREM.contents, BODY);
  });
}

/* ------------------------------------------------------ slide 3: our company */

function slide03(p) {
  const s = p.addSlide();
  chrome(s);
  heading(s, 0.462, 0.424, 6.204, 'Our Company');

  T(s, 3.255, 2.957, 1.26, 0.404, [P('Lorem'), P('Ipsum Dolar')], { fontSize: 9 });
  hline(s, 0, 3.27, 3.125, C.rule, 0.75);
  T(s, 5.195, 2.604, 3.24, 0.83,
    'Am no an listening depending up believing. Enough to around it remove to burton\u2019s '
    + 'agreed regret in or there Advantage mar there',
    Object.assign({ italic: true, color: C.grey4 }, BODY));

  [
    { x: 0.455, tx: 1.774, lx: 0.476, tw: 2.554, num: '01.', color: C.a1 },
    { x: 5.205, tx: 6.524, lx: 5.205, tw: 2.57, num: '02.', color: C.a3 },
  ].forEach(function (col) {
    T(s, col.x, 4.168, 1.319, 1.027, col.num,
      { fontSize: 55, fontFace: F.thin, color: col.color });
    T(s, col.tx, 4.303, col.tw, 0.83,
      'PLACEHOLDER', BODY);
    T(s, col.lx, 5.326, 1.319, 0.519, 'Lorem Ipsum Dolor',
      { fontSize: 10, color: col.color, lineSpacingMultiple: 1.3 });
  });
}

/* ----------------------------------------------- slide 4: what we are working */

function slide04(p) {
  const s = p.addSlide();
  chrome(s);

  rect(s, 3.877, 2.007, 8.704, 3.017, { color: C.panel, transparency: 30 });
  heading(s, 0.462, 0.424, 8.704, 'What We Are Working On');
  T(s, 0.462, 1.812, 3.248, 1.043, 'Your awesome subtitle text', LEAD);
  T(s, 0.462, 3.644, 3.248, 0.83, [
    P('PLACEHOLDER'),
    P('in are securing off'),
  ], { fontSize: 10, lineSpacingMultiple: 1.5 });

  vline(s, 6.669, 2.156, 2.843, C.panel2, 1);
  vline(s, 9.79, 2.156, 2.843, C.panel2, 1);

  [
    { x: 4.146, ix: 4.27, color: C.a1, glyph: IC.plane },
    { x: 7.267, ix: 7.391, color: C.a2, glyph: IC.lines },
    { x: 10.387, ix: 10.512, color: C.a3, glyph: IC.badge },
  ].forEach(function (col) {
    icon(s, col.ix, 2.319, 0.537, col.glyph, col.color, 26);
    T(s, col.x, 2.994, 1.925, 0.404, 'Insert Title',
      Object.assign({ color: col.color }, SUBHEAD));
    T(s, col.x, 3.63, 1.925, 1.083, LOREM.card, BODY);
  });

  hline(s, 0.752, 5.766, 11.83, C.panel, 1.5);
  T(s, 3.465, 5.889, 6.403, 0.578, LOREM.footnote,
    { fontSize: 10, italic: true, color: C.a2, align: 'center', lineSpacingMultiple: 1.5 });
}

/* ------------------------------------------------- slide 5: about the project */

function slide05(p) {
  const s = p.addSlide();
  chrome(s);
  heading(s, 0.462, 0.424, 6.836, 'About The Project');
  T(s, 0.571, 1.766, 4.366, 1.043, [P('Your awesome sub-'), P('title text')], LEAD);
  T(s, 0.571, 3.011, 5.738, 0.83, LOREM.long, BODY);

  [
    { x: 0.67, y: 4.477, color: C.a1 },
    { x: 3.736, y: 4.477, color: C.a3 },
    { x: 0.67, y: 5.721, color: C.a2 },
    { x: 3.736, y: 5.721, color: C.a4 },
  ].forEach(function (b) {
    ellipse(s, b.x, b.y, 0.607, 0.607, { color: b.color });
    icon(s, b.x, b.y, 0.607, IC.check, C.white, 20);
    T(s, b.x + 0.781, b.y - 0.003, 1.9, 0.578, 'Allowance address farther hearted hund', BODY);
  });
}

/* ------------------------------------------------------------- slide 6: quote */

function slide06(p) {
  const s = p.addSlide();
  chrome(s);
  T(s, 2.979, 2.832, 7.375, 0.959,
    '\u201CThis is a quote. Words full of wisdom that someone important said and can make '
    + 'the reader get inspired.\u201D',
    { fontSize: 18, align: 'center', lineSpacingMultiple: 1.5 });
  hline(s, 6.497, 4.215, 0.339, C.dk2, 2.25);
  T(s, 5.15, 4.399, 3.033, 0.269, '\u2014SOMEONE FAMOUS',
    { fontSize: 10, fontFace: F.light, align: 'center' });
}

/* ------------------------------------------------------- slide 7: now project */

function slide07(p) {
  const s = p.addSlide();
  chrome(s);
  T(s, 0.967, 1.771, 2.725, 1.582, 'Now Project', { fontSize: 44, bold: true });

  vline(s, 10.199, 1.698, 1.969, C.rule, 0.75);
  T(s, 10.269, 1.552, 1.414, 0.83, bulletList(['Minimal', 'Creative', 'Simple']),
    { fontSize: 10, lineSpacingMultiple: 1.5 });

  T(s, 2.583, 5.695, 2.625, 0.602,
    [P('You must'), P('know your product service\u2026.')],
    { fontSize: 12, color: C.a3, lineSpacingMultiple: 1.3 });
  rect(s, 6.067, 5.771, 0.05, 0.602, { color: C.rule });
  T(s, 6.975, 5.619, 4.466, 0.83,
    'Am no an listening depending up believing. Enough around it remove to burton\u2019s agreed '
    + 'regret in or it. there Advantage mar estimable beu commanded there',
    Object.assign({}, BODY, { color: C.ink }));
}

/* ---------------------------------------------------- slide 8: future project */

function slide08(p) {
  const s = p.addSlide();
  chrome(s);
  heading(s, 0.529, 0.489, 4.866, 'Future Project');
  T(s, 0.752, 1.714, 3.675, 1.043, 'Your awesome sub-title text', LEAD);
  T(s, 0.752, 3.216, 4.076, 0.83,
    'PLACEHOLDER'
    + 'husband in are securing off  its there occasion  daughter replyings', BODY);

  rect(s, 0.841, 5.016, 3.988, 1.335, { color: C.a1 });
  T(s, 1.05, 5.16, 0.6, 0.6, '\u201C',
    { fontSize: 44, color: C.a1dark, fontFace: F.reg, margin: 0 });
  T(s, 1.398, 5.4, 3.133, 0.578,
    'husband in are securing off  it there occasion daughter the own to yet.',
    { fontSize: 10, italic: true, color: C.white, lineSpacingMultiple: 1.5 });

  [
    { x: 5.258, title: 'Project 01' },
    { x: 7.777, title: 'Project 02' },
    { x: 10.295, title: 'Project 03' },
  ].forEach(function (col) {
    T(s, col.x, 4.911, 2.252, 0.404, col.title, SUBHEAD);
    T(s, col.x, 5.378, 2.252, 0.83, LOREM.project, BODY);
  });
}

/* ----------------------------------------------- slide 9: major requirements */

function slide09(p) {
  const s = p.addSlide();
  s.background = { color: C.white };
  chrome(s);
  heading(s, 0.529, 0.489, 7.63, 'Major Requirements');

  const cards = [
    { x: 0.752, color: C.a1, title: 'Strengths', glyph: IC.square, snip: false },
    { x: 3.773, color: C.a2, title: 'Weaknesses', glyph: IC.diamond, snip: false },
    { x: 6.794, color: C.a3, title: 'Opportunities', glyph: IC.flag, snip: false },
    { x: 9.815, color: C.a4, title: 'Threats', glyph: IC.lines, snip: true },
  ];
  cards.forEach(function (c) {
    s.addShape(c.snip ? 'snip1Rect' : 'rect',
      { x: c.x, y: 2.294, w: 2.766, h: 2.766, fill: { color: c.color } });
    icon(s, c.x + 0.5, 2.75, 0.53, c.glyph, C.white, 24);
    T(s, c.x + 0.338, 3.426, 1.889, 0.404, c.title,
      { fontSize: 18, bold: true, color: C.white });
    T(s, c.x + 0.338, 3.836, 2.09, 0.83, LOREM.swot,
      Object.assign({}, BODY, { color: C.white }));
  });

  hline(s, 0.752, 5.72, 11.83, C.rule, 1.25);
  T(s, 3.465, 6.201, 6.403, 0.578, LOREM.footnote,
    { fontSize: 10, italic: true, fontFace: F.light, color: C.mute,
      align: 'center', lineSpacingMultiple: 1.5 });
}

/* ------------------------------------------------------------ slide 10: budget */

function slide10(p) {
  const s = p.addSlide();
  chrome(s);
  heading(s, 0.508, 0.518, 3.208, 'Budget');
  hline(s, 0.573, 4.748, 12.19, C.rule, 1.25);

  T(s, 1.431, 1.64, 3.208, 0.602,
    [P('You must'), P('know your project budget\u2026.')],
    { fontSize: 12, italic: true, color: C.a3, align: 'right', lineSpacingMultiple: 1.3 });
  T(s, 0.8, 4.876, 3.819, 1.027, '$50,000',
    { fontSize: 55, fontFace: F.light, align: 'right' });
  T(s, 1.452, 5.821, 3.188, 0.3, 'TOTAL BUDGET',
    { fontSize: 10, color: C.mute, align: 'right', lineSpacingMultiple: 1.3 });


  // Column chart drawn from plain rectangles, exactly as the source deck does.
  T(s, 8.032, 1.058, 2.436, 0.303, 'Budget Chart',
    { fontSize: 12, fontFace: F.med, align: 'center', lineSpacingMultiple: 1.5 });
  T(s, 5.51, 1.622, 0.451, 2.289,
    ['6', '5', '4', '3', '2', '1', '0'].map(function (v) { return P(v); }),
    { fontSize: 10, fontFace: F.med, align: 'center', lineSpacingMultiple: 2 });

  const groups = [
    { labelX: 6.15, bars: [[6.052, 2.242, 1.572], [6.495, 2.972, 0.842], [6.938, 3.08, 0.733]] },
    { labelX: 7.874, bars: [[7.776, 2.89, 0.886], [8.219, 2.242, 1.535], [8.662, 3.085, 0.691]] },
    { labelX: 9.626, bars: [[9.528, 2.542, 1.222], [9.97, 2.877, 0.886], [10.413, 2.737, 1.027]] },
    { labelX: 11.338, bars: [[11.24, 2.192, 1.572], [11.683, 2.768, 0.995], [12.126, 2.021, 1.742]] },
  ];
  const series = [C.a1, C.a2, C.a3];
  groups.forEach(function (g, gi) {
    g.bars.forEach(function (b, bi) {
      rect(s, b[0], b[1], 0.443, b[2], { color: series[bi] });
    });
    T(s, g.labelX, gi === 0 ? 3.828 : (gi === 1 ? 3.791 : 3.778), 1.132, 0.269, 'Category 1',
      { fontSize: 10, fontFace: F.med, align: 'center', lineSpacingMultiple: 1.5 });
  });

  [[7.828, 7.968, 'Series 1', C.a1], [8.897, 9.036, 'Series 2', C.a2],
   [10.026, 10.165, 'Series 3', C.a3]].forEach(function (l) {
    rect(s, l[0], 4.284, 0.173, 0.173, { color: l[3] });
    T(s, l[1], 4.236, 0.761, 0.269, l[2],
      { fontSize: 10, fontFace: F.med, lineSpacingMultiple: 1.5 });
  });

  [
    { x: 5.542, amount: '$5,000', label: 'PROJECT TITLE 01', color: C.a1 },
    { x: 8.245, amount: '$15,000', label: 'PROJECT TITLE 02', color: C.a2 },
    { x: 10.949, amount: '$30,000', label: 'PROJECT TITLE 03', color: C.a3 },
  ].forEach(function (c) {
    T(s, c.x, 5.054, 2.051, 0.404, c.amount, { fontSize: 18, bold: true });
    T(s, c.x, 5.399, 2.051, 0.325, c.label,
      { fontSize: 10, fontFace: F.med, color: c.color, lineSpacingMultiple: 1.5 });
    T(s, c.x, 5.682, 2.051, 0.578,
      [P('Certainly elsewhere allowa'), P('address farther')], BODY);
  });
}

/* ------------------------------------------------------ slide 11: project goals */

function slide11(p) {
  const s = p.addSlide();
  s.background = { color: C.white };
  chrome(s);
  heading(s, 3.592, 0.489, 6.15, 'Project Goals', { align: 'center' });

  // "PROJECT GOAL" callout (sits behind the target).
  pill(s, 5.743, 2.356, 1.848, 0.503, { color: C.grey1 }, 0.25);
  poly(s, 6.55, 2.845, 0.19, 0.145, [[0, 0], [0.19, 0], [0.095, 0.145]], { color: C.grey1 });

  // Concentric target: grey cap over the top of each coloured ring.
  const cx = 6.6445, cy = 4.2945;
  const rings = [
    { r: 1.256, color: C.a1, cap: 90 },
    { r: 1.074, color: C.white, cap: 0 },
    { r: 0.900, color: C.a2, cap: 126 },
    { r: 0.726, color: C.white, cap: 0 },
    { r: 0.556, color: C.a3, cap: 52 },
    { r: 0.393, color: C.white, cap: 0 },
    { r: 0.214, color: C.a4, cap: 0 },
  ];
  rings.forEach(function (ring) {
    const d = ring.r * 2;
    ellipse(s, cx - ring.r, cy - ring.r, d, d, { color: ring.color });
    if (ring.cap) {
      s.addShape('pie', { x: cx - ring.r, y: cy - ring.r, w: d, h: d,
        fill: { color: C.grey2 }, angleRange: [269 - ring.cap / 2, 269 + ring.cap / 2] });
    }
  });

  // Arrow shaft running through the target, drawn over it.
  rect(s, 0.646, 4.204, 5.965, 0.179, { color: C.grey3 });
  rect(s, 6.7, 4.204, 5.964, 0.179, { color: C.grey3 });
  poly(s, 0.681, 3.969, 0.712, 0.29,
    [[0.2, 0.29], [0.045, 0], [0.49, 0], [0.712, 0.29]], { color: C.grey1 });
  poly(s, 0.681, 4.327, 0.7, 0.292,
    [[0.2, 0], [0.033, 0.292], [0.478, 0.292], [0.7, 0]], { color: C.grey3 });
  poly(s, 11.918, 3.969, 0.712, 0.29,
    [[0.512, 0.29], [0.667, 0], [0.222, 0], [0, 0.29]], { color: C.grey1 });
  poly(s, 11.918, 4.327, 0.7, 0.292,
    [[0.5, 0], [0.667, 0.292], [0.222, 0.292], [0, 0]], { color: C.grey3 });

  T(s, 5.856, 2.424, 1.622, 0.269, 'PROJECT GOAL', { fontSize: 10, align: 'center' });
  T(s, 5.431, 5.766, 2.472, 0.83, [
    P('Certainly elsewhere my allowance'),
    P('at address farther six hearted'),
    P('The it is a hundred.'),
  ], { fontSize: 10, italic: true, color: C.grey3, align: 'center', lineSpacingMultiple: 1.5 });

  // Four "insert title" pills: pale rounded head + solid coloured body.
  [
    { x: 1.582, color: C.a1, tint: 'BEC7E6', glyph: IC.target },
    { x: 3.495, color: C.a2, tint: 'DFE9F6', glyph: IC.square },
    { x: 8.168, color: C.a3, tint: 'DBEFEA', glyph: IC.ring },
    { x: 10.082, color: C.a4, tint: 'E2E0E0', glyph: IC.sparkle },
  ].forEach(function (c) {
    pill(s, c.x, 2.447, 1.646, 3.581, { color: c.tint }, 0.823);
    pill(s, c.x, 3.688, 1.646, 2.34, { color: c.color }, 0.823);
    rect(s, c.x, 3.688, 1.646, 0.9, { color: c.color });
    icon(s, c.x + 0.566, 2.902, 0.514, c.glyph, c.color, 24);
    T(s, c.x + 0.087, 4.048, 1.471, 0.404, 'Insert title',
      { fontSize: 18, bold: true, color: C.white, align: 'center' });
    T(s, c.x + 0.087, 4.462, 1.471, 1.083, LOREM.goals,
      Object.assign({}, BODY, { color: C.white, align: 'center' }));
  });
}

/* -------------------------------------------------- slide 12: predicted results */

function slide12(p) {
  const s = p.addSlide();
  chrome(s);
  heading(s, 0.529, 0.489, 6.836, 'Predicted Results');

  // Stacked area chart built from two filled polygons (as in the source deck).
  poly(s, 0.999, 2.598, 5.401, 3.205, [
    [0, 0], [1.374, 0], [2.684, 0.281], [4.076, 0.802],
    [5.401, 0.037], [5.401, 3.205], [0.013, 3.205],
  ], { color: C.a2 });
  poly(s, 1.006, 3.485, 5.384, 2.318, [
    [0, 0], [1.350, 0.002], [2.715, 0.275], [4.068, 1.439],
    [5.384, 1.226], [5.384, 2.308], [0.006, 2.318],
  ], { color: C.a1 });

  T(s, 2.451, 1.752, 2.436, 0.337, 'Chart Title',
    { fontSize: 14, fontFace: F.med, align: 'center' });
  T(s, 0.571, 2.194, 0.451, 3.636,
    ['50', '45', '40', '35', '30', '25', '20', '15', '10', '5', '0']
      .map(function (v) { return P(v); }),
    { fontSize: 10, fontFace: F.med, align: 'center', lineSpacingMultiple: 2 });
  [0.376, 1.723, 3.069, 4.416, 5.762].forEach(function (x) {
    T(s, x, 5.941, 1.132, 0.269, '1/5/2002',
      { fontSize: 10, fontFace: F.med, align: 'center' });
  });
  [[3.027, 3.166, 'Series 1', C.a1], [4.096, 4.235, 'Series 2', C.a2]].forEach(function (l) {
    rect(s, l[0], 6.48, 0.173, 0.173, { color: l[3] });
    T(s, l[1], 6.432, 0.761, 0.269, l[2], { fontSize: 10, fontFace: F.med });
  });

  T(s, 7.234, 1.692, 4.068, 1.043, [P('Your awesome sub-'), P('title text')], LEAD);
  T(s, 7.234, 2.824, 5.738, 0.83, LOREM.long, BODY);

  [
    { y: 4.053, color: C.a1, value: '$20,000', label: 'EXPECTED INCOME FOR 2020' },
    { y: 5.388, color: C.a2, value: '1.5K', label: 'NEW EMPLOYEES NEXT YEAR ' },
  ].forEach(function (m) {
    ellipse(s, 7.333, m.y, 0.607, 0.607, { color: m.color });
    icon(s, 7.333, m.y, 0.607, IC.check, C.white, 20);
    T(s, 8.101, m.y - 0.191, 2.051, 0.404, m.value, { fontSize: 18, bold: true });
    T(s, 8.101, m.y + 0.201, 3.008, 0.269, m.label,
      { fontSize: 10, fontFace: F.med, color: m.color });
    T(s, 8.101, m.y + 0.466, 4.557, 0.578,
      [P(LOREM.metric[0]), P(LOREM.metric[1])], BODY);
  });
}

/* ---------------------------------------------------------- slide 13: sneak peek */

function slide13(p) {
  const s = p.addSlide();
  s.background = { color: C.a4 };
  radialGradient(s, [[0, C.white], [0.05, C.white], [1, C.a4]]);
  chrome(s);

  photoPlaceholder(s, 2.73, 1.78, 7.873, 3.338,
    { fill: '8E8F96', line: 'B4B5BC', label: C.white });
  heading(s, 3.683, 0.489, 5.968, 'Sneak Peek', { align: 'center', color: C.white });

  [
    { fx: 5.687, gx: 5.928, glyph: IC.gear, tx: 3.856, bx: 2.956, align: 'right' },
    { fx: 6.886, gx: 7.113, glyph: IC.star, tx: 7.887, bx: 7.887, align: 'left' },
  ].forEach(function (c) {
    s.addShape('roundRect', { x: c.fx, y: 5.962, w: 0.76, h: 0.76,
      fill: { type: 'none' }, line: { color: C.white, width: 1.25 }, rectRadius: 0.1 });
    icon(s, c.gx - 0.09, 6.14, 0.4, c.glyph, C.dk2, 17);
    T(s, c.tx, 5.796, 1.591, 0.404, 'Insert title',
      Object.assign({ align: c.align }, SUBHEAD));
    T(s, c.bx, 6.262, 2.49, 0.578, LOREM.peek,
      Object.assign({}, BODY, { color: C.white, align: c.align }));
  });
}

/* -------------------------------------------------------- slide 14: project stages */

function slide14(p) {
  const s = p.addSlide();
  chrome(s);
  heading(s, 0.529, 0.489, 6.138, 'Project Stages');

  // Tablet photograph placeholder (rotated image in the source deck).
  s.addShape('roundRect', { x: -0.35, y: 1.834, w: 5.41, h: 4.594,
    fill: { color: C.white }, line: { color: C.grey3, width: 1 }, rectRadius: 0.22 });
  rect(s, -0.35, 2.04, 5.03, 4.12, { color: '2B2B2B' });
  T(s, 0.6, 3.9, 3.0, 0.4, '[image]',
    { fontSize: 11, color: 'A0A0A0', align: 'center', valign: 'middle', margin: 0 });

  // Right hand grey panel with a clipped corner.
  s.addShape('snip1Rect', { x: 8.063, y: 2.101, w: 4.659, h: 4.036,
    fill: { color: C.rule, transparency: 45 } });

  const rows = [
    { y: 2.101, h: 1.346, color: C.a1, glyph: IC.crown },
    { y: 3.447, h: 1.346, color: C.a2, glyph: IC.square },
    { y: 4.792, h: 1.344, color: C.a3, glyph: IC.tree },
  ];
  rows.forEach(function (r) {
    rect(s, 5.068, r.y, 2.995, r.h, { color: r.color });
    poly(s, 8.09, r.y, 1.13, r.h, [[0, 0], [1.13, r.h / 2], [0, r.h]], { color: r.color });
    poly(s, 8.063, r.y, 1.13, r.h, [[0, 0], [1.13, r.h / 2], [0, r.h]], { color: C.white });
    icon(s, 8.31, r.y + r.h / 2 - 0.2, 0.4, r.glyph, r.color, 20);
    T(s, 5.375, r.y + 0.194, 2.019, 0.404, 'Insert title',
      Object.assign({ color: C.white }, SUBHEAD));
    T(s, 5.375, r.y + 0.574, 2.381, 0.578, LOREM.stages,
      Object.assign({}, BODY, { color: C.white }));
  });

  T(s, 9.648, 2.615, 2.656, 1.043, 'Our Best sub-title text', LEAD);
  T(s, 9.648, 3.86, 2.54, 0.83,
    'PLACEHOLDER',
    BODY);
  T(s, 9.648, 4.792, 1.414, 0.83, bulletList(['Minimal', 'Creative', 'Simple']),
    { fontSize: 10, lineSpacingMultiple: 1.5 });
}

/* ------------------------------------------------------------- slide 15: timeline */

function slide15(p) {
  const s = p.addSlide();
  chrome(s);
  heading(s, 0.529, 0.489, 6.365, 'Timeline');

  // Dotted connector running through the year markers.
  [[1.492, 3.665], [4.04, 4.038], [5.787, 4.065], [7.504, 4.006],
   [9.195, 4.038], [11.675, 3.665]].forEach(function (d) {
    ellipse(s, d[0], d[1], 0.166, 0.168, { color: C.mute });
  });
  [[2.0, 3.86, 0.85, -18], [3.95, 4.18, 1.0, 6], [5.65, 4.2, 1.0, -4],
   [7.35, 4.15, 1.0, 3], [9.05, 4.16, 1.0, -8], [10.7, 3.86, 0.95, -16]]
    .forEach(function (seg) {
      s.addShape('line', { x: seg[0], y: seg[1], w: seg[2], h: 0, rotate: seg[3],
        line: { color: C.rule, width: 1.5, dashType: 'sysDot' } });
    });

  const steps = [
    { cx: 2.703, year: '2020', color: C.a4dark, day: 'Day 01', below: true,
      ix: 3.137, iy: 2.825, iglyph: IC.flag, icolor: C.dk2 },
    { cx: 4.42, year: '2021', color: C.a4, day: 'Day 02', below: false,
      ix: 4.821, iy: 5.016, iglyph: IC.lines, icolor: C.a4 },
    { cx: 6.138, year: '2022', color: C.a3, day: 'Day 03', below: true,
      ix: 6.573, iy: 2.825, iglyph: IC.bell, icolor: C.a3 },
    { cx: 7.855, year: '2023', color: C.a2, day: 'Day 04', below: false,
      ix: 8.29, iy: 5.038, iglyph: IC.square, icolor: C.a2 },
    { cx: 9.572, year: '2024', color: C.a1, day: 'Day 05', below: true,
      ix: 9.973, iy: 2.825, iglyph: IC.plane, icolor: C.a1 },
  ];
  steps.forEach(function (st) {
    ellipse(s, st.cx, 3.524, 1.183, 1.193, { color: st.color },
      { line: { color: C.white, width: 1, dashType: 'sysDot' } });
    T(s, st.cx + 0.03, 3.919, 1.123, 0.404, st.year,
      Object.assign({ color: C.white, align: 'center' }, SUBHEAD));
    icon(s, st.ix, st.iy, 0.382, st.iglyph, st.icolor, 17);

    const ty = st.below ? 5.016 : 1.925;
    T(s, st.cx - 0.1, ty, 1.457, 0.404, st.day,
      Object.assign({ align: 'center' }, SUBHEAD));
    T(s, st.cx - 0.322, ty + 0.466, 1.833, 0.83, LOREM.timeline,
      Object.assign({}, BODY, { align: 'center' }));
  });
}

/* ------------------------------------------------------------- slide 16: partners */

// Vertical bar whose right edge bows inward around the partner badge.
// Drawn twice (grey, then white nudged up/left) so only a thin crescent shows.
const NOTCHED_BAR = [
  ['M', 1, 0.999],
  ['C', 1, 0.999, 1, 0.773, 1, 0.746],
  ['C', 1, 0.725, 0.994, 0.706, 0.960, 0.690],
  ['C', 0.910, 0.666, 0.834, 0.659, 0.768, 0.647],
  ['C', 0.631, 0.621, 0.552, 0.556, 0.552, 0.499],
  ['C', 0.552, 0.443, 0.631, 0.378, 0.768, 0.352],
  ['C', 0.834, 0.340, 0.910, 0.332, 0.960, 0.309],
  ['C', 0.994, 0.292, 1, 0.273, 1, 0.252],
  ['C', 1, 0.225, 1, 0, 1, 0],
  ['L', 0, 0], ['L', 0, 1], ['L', 1, 0.999], ['Z'],
];

function slide16(p) {
  const s = p.addSlide();
  chrome(s);
  heading(s, 0.529, 0.489, 4.791, 'Our Partners');

  [
    { x: 1.229, color: C.a1, num: '1' },
    { x: 3.946, color: C.a2, num: '2' },
    { x: 6.661, color: C.a3, num: '3' },
    { x: 9.375, color: C.a4, num: '4' },
  ].forEach(function (c) {
    pill(s, c.x, 3.041, 2.457, 1.348, { color: c.color }, 0.674);
    curveShape(s, c.x - 0.232, 2.366, 1.122, 2.701, NOTCHED_BAR, { color: C.rule });
    curveShape(s, c.x - 0.331, 2.266, 1.122, 2.902, NOTCHED_BAR, { color: C.white });
    ellipse(s, c.x + 0.434, 3.359, 0.712, 0.712, { color: C.white });
    icon(s, c.x + 0.434, 3.359, 0.712, IC.tree, c.color, 22);
    T(s, c.x + 0.947, 2.409, 0.44, 0.572, c.num,
      { fontSize: 28, fontFace: F.med, color: C.grey2 });
    T(s, c.x + 1.211, 3.362, 1.122, 0.707, 'Insert subtitle',
      Object.assign({ color: C.white }, SUBHEAD));
    T(s, c.x + 0.229, 5.432, 2.457, 0.83, LOREM.partners, BODY);
  });
}

/* ------------------------------------------------------------- slide 17: our team */

function slide17(p) {
  const s = p.addSlide();
  linearGradient(s, 45, [[0, 'F8FCFB'], [0.24, 'F8FCFB'], [1, 'DBEEEA']]);
  chrome(s);
  heading(s, 0.651, 0.489, 4.223, 'Our Team');

  const skills = [
    { label: 'UI', pct: '33%', width: 1.491, color: C.a1, y: 5.635, ty: 5.603, py: 5.609 },
    { label: 'UX', pct: '75%', width: 2.218, color: C.a2, y: 5.922, ty: 5.892, py: 5.899 },
    { label: 'ID', pct: '65%', width: 1.893, color: C.a3, y: 6.208, ty: 6.181, py: 6.188 },
  ];

  [0.651, 4.776, 8.902].forEach(function (x) {
    T(s, x, 3.933, 3.254, 0.404, 'Name text here', SUBHEAD);
    T(s, x, 4.295, 2.01, 0.325, 'Position text here', { fontSize: 10 });
    T(s, x, 4.728, 3.68, 0.578, LOREM.team, BODY);

    skills.forEach(function (sk) {
      T(s, x, sk.ty, 0.519, 0.269, sk.label, { fontSize: 10 });
      pill(s, x + 0.519, sk.y, 2.559, 0.133, { color: C.white }, 0.0665);
      pill(s, x + 0.519, sk.y, sk.width, 0.133, { color: sk.color }, 0.0665);
      T(s, x + 3.161, sk.py, 0.519, 0.269, sk.pct, { fontSize: 10, align: 'right' });
    });
  });
}

/* ------------------------------------------------------------ slide 18: contact us */

function slide18(p) {
  const s = p.addSlide();
  chrome(s);
  heading(s, 6.271, 1.021, 5.396, 'Contact us');


  const blocks = [
    { ix: 6.271, tx: 6.913, y: 2.87, color: C.a1, glyph: IC.flag, title: 'Address', tw: 1.296,
      lines: ['2476 Leverton Cove Road', 'Springfield, MA, Massachusetts, 01109.'], bh: 0.83 },
    { ix: 9.308, tx: 9.966, y: 2.87, color: C.a2, glyph: IC.envelope, title: 'Email', tw: 1.142,
      lines: ['urcontactemail@mail.com', 'urpersonalemail@mail.com'], bh: 0.578 },
    { ix: 6.271, tx: 6.913, y: 4.89, color: C.a3, glyph: IC.phone, title: 'Phone', tw: 1.142,
      lines: ['+413-364-6795', '+413-426-3944'], bh: 0.578 },
    { ix: 9.308, tx: 9.966, y: 4.89, color: C.a4, glyph: IC.badge, title: 'Website', tw: 1.326,
      lines: ['www.urwebitedomain.com', 'www.uranotherwebsite.com'], bh: 0.578 },
  ];
  blocks.forEach(function (b) {
    ellipse(s, b.ix, b.y + 0.322, 0.501, 0.501, { color: b.color });
    icon(s, b.ix, b.y + 0.322, 0.501, b.glyph, C.white, 17);
    T(s, b.tx, b.y + 0.269, b.tw, 0.404, b.title, SUBHEAD);
    T(s, b.tx, b.y + 0.735, 2.105, b.bh,
      b.lines.map(function (l) { return P(l); }), BODY);
  });
}

/* ---------------------------------------------------------------------- build */

function build() {
  const pptx = new PptxGenJS();
  pptx.author = 'pptxgenjs';
  pptx.title = 'Project Proposal';
  pptx.defineLayout({ name: 'DECK', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'DECK';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
   slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18]
    .forEach(function (fn) { fn(pptx); });

  const out = path.join(__dirname, '0802fd70-20ce-4d79-b85b-44fbdb11be11_grok_final.pptx');
  return pptx.writeFile({ fileName: out }).then(function () {
    console.log('wrote ' + out);
  });
}

build().catch(function (e) { console.error(e); process.exit(1); });
