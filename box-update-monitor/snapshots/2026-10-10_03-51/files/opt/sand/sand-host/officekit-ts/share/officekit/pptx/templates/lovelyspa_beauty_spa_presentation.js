/*
 * "Lovelyspa" beauty presentation - 33 slides, 20 x 11.25 in.
 * Rebuilt with pptxgenjs only. Photographs in the original deck are replaced
 * by flat grey placeholder rectangles (see imagePlaceholder()).
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const C = {
  white: 'FFFFFF',
  grey: 'BFBFBF', // bg1 lumMod 75% - photo/panel placeholder grey
  greyLt: 'D9D9D9', // bg1 lumMod 85%
  greyDk: 'A6A6A6', // bg1 lumMod 65% - body copy on white
  ink: '262626', // tx1 lum 85/15 - headline dark
  slate: '404040', // tx1 lum 75/25 - subtitles
  mute: '7F7F7F', // tx1 lum 50/50 - small caps labels
  pink: 'EA5EA7', // accent used for small caps + outlines
  rose: 'E86A93', // mid stop of the brand gradient, used for headline accent
  roseFlat: 'E14D7C' // average of the gradient - used where charts need one colour
};

const FONT_HEAD = 'Poppins';
const FONT_BODY = 'Open Sans';

// Brand gradient: corner-to-corner rose -> deep pink (E86A93 -> DA2864).
const GRADIENT = [[232, 106, 147], [218, 40, 100]];

const LOREM_LONG =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod ' +
  'tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam. ';
const LOREM_MID =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod ' +
  'tempor incididunt ut labore';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.';
const LOREM_TINY = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et.';

/* ---------------------------------------------------------- text helpers */

// PowerPoint text boxes in the source use the default 0.1" left/right and
// 0.05" top/bottom insets, and are top-anchored (pptxgenjs centres by default).
// pptxgenjs margin arrays are ordered [left, right, bottom, top] in points.
const TEXT_BASE = { valign: 'top', margin: [7.2, 7.2, 3.6, 3.6], isTextBox: true };

function text(slide, body, opts) {
  slide.addText(body, Object.assign({}, TEXT_BASE, opts));
}

// 48pt Poppins headline. `lines` holds one entry per line, either a string or
// an array of same-line fragments. The very last fragment is the accent word
// ("Lovelyspa", "Touch", ...) and is coloured with the brand rose.
function headline(slide, x, y, w, lines, opts) {
  const o = opts || {};
  const runs = [];
  lines.forEach(function (line, i) {
    const parts = Array.isArray(line) ? line : [line];
    parts.forEach(function (part, j) {
      runs.push({ text: part, options: { color: o.color || C.ink, breakLine: j === parts.length - 1 && i < lines.length - 1 } });
    });
  });
  runs[runs.length - 1].options.color = o.accent || C.rose;
  text(slide, runs, { x: x, y: y, w: w, h: o.h || 1.717, fontFace: FONT_HEAD, fontSize: o.size || 48, bold: true });
}

// 16pt Open Sans body copy, 1.5 line spacing.
function body(slide, x, y, w, h, str, opts) {
  const o = opts || {};
  text(slide, str, {
    x: x, y: y, w: w, h: h,
    fontFace: FONT_BODY, fontSize: 16, color: o.color || C.greyDk,
    lineSpacingMultiple: 1.5, align: o.align, bold: o.bold
  });
}

// Small bold letter-spaced caps ("BEAUTY", "PRESENTATION", ...).
function caps(slide, x, y, w, runs, opts) {
  const o = opts || {};
  text(slide, runs, {
    x: x, y: y, w: w, h: 0.37,
    fontFace: FONT_BODY, fontSize: 16, bold: true,
    charSpacing: o.spc === undefined ? 3 : o.spc,
    color: o.color || C.mute, align: o.align, wrap: o.wrap
  });
}

// "Subtitle Here" style label (18pt default size, bold).
function subtitle(slide, x, y, str, color) {
  text(slide, str || 'Subtitle Here', {
    x: x, y: y, w: 1.86, h: 0.404,
    fontFace: FONT_BODY, fontSize: 18, bold: true, color: color || C.slate, wrap: false
  });
}

// Big Poppins number, e.g. "2021".
function year(slide, x, y, w, str, color) {
  text(slide, str, { x: x, y: y, w: w || 1.651, h: 0.909, fontFace: FONT_HEAD, fontSize: 48, bold: true, color: color });
}

/* --------------------------------------------------------- shape helpers */

function rect(slide, x, y, w, h, color) {
  slide.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: color }, line: { type: 'none' } });
}

// Grey stand-in for a photograph from the original deck.
function imagePlaceholder(slide, x, y, w, h, shade) {
  rect(slide, x, y, w, h, shade || C.grey);
}

// The recurring "BEAUTY PRESENTATION" chip: an outlined box with caps inside.
// dir 'h' = horizontal box, 'v' = vertical box with rotated text.
function chip(slide, x, y, dir, color) {
  const col = color || C.pink;
  const txtCol = color === C.white ? C.white : C.slate;
  if (dir === 'v') {
    slide.addShape('rect', { x: x, y: y, w: 0.917, h: 4.131, fill: { type: 'none' }, line: { color: col, width: 3 } });
    text(slide, 'BEAUTY PRESENTATION', {
      x: x + 0.917 / 2 - 4.196 / 2, y: y + 4.131 / 2 - 0.37 / 2, w: 4.196, h: 0.37, rotate: 270,
      fontFace: FONT_BODY, fontSize: 16, bold: true, charSpacing: 3, align: 'center', color: txtCol
    });
  } else {
    slide.addShape('rect', { x: x, y: y, w: 4.131, h: 0.917, fill: { type: 'none' }, line: { color: col, width: 3 } });
    caps(slide, x, y + 0.917 / 2 - 0.37 / 2, 4.131, 'BEAUTY PRESENTATION', { align: 'center', color: txtCol });
  }
}

const SPA_AND_COSMETIC = [
  { text: 'SPA', options: { color: C.pink } },
  { text: ' AND COSMETIC', options: { color: C.mute } }
];

/* ------------------------------------------------- gradient fill helpers */

// pptxgenjs has no gradient fill, so gradients are drawn as a fan of thin
// custGeom slices whose colour is sampled from the brand ramp.
function rampColor(t) {
  const a = GRADIENT[0], b = GRADIENT[1];
  const p = Math.max(0, Math.min(1, t));
  return a
    .map(function (v, k) { return Math.round(v + (b[k] - v) * p); })
    .map(function (v) { return v.toString(16).padStart(2, '0'); })
    .join('').toUpperCase();
}

// Gradient axis. PowerPoint's 45-degree "scaled" gradient runs corner to
// corner, so the two unit axes are weighted by the shape's width and height.
// `dark` names the corner holding the deepest colour.
function axisFor(dark, w, h) {
  const fx = w / (w + h), fy = h / (w + h);
  const sx = dark === 'tl' || dark === 'bl' ? -1 : 1;
  const sy = dark === 'tl' || dark === 'tr' ? -1 : 1;
  const base = (sx < 0 ? fx : 0) + (sy < 0 ? fy : 0);
  return function (p) { return base + sx * fx * p[0] + sy * fy * p[1]; };
}

// Sutherland-Hodgman half-plane clip of a unit polygon.
function clipPoly(poly, axis, limit, keepAbove) {
  const out = [];
  for (let i = 0; i < poly.length; i++) {
    const a = poly[i], b = poly[(i + 1) % poly.length];
    const da = axis(a) - limit, db = axis(b) - limit;
    const ina = keepAbove ? da >= 0 : da <= 0;
    const inb = keepAbove ? db >= 0 : db <= 0;
    if (ina) out.push(a);
    if (ina !== inb) {
      const t = da / (da - db);
      out.push([a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t]);
    }
  }
  return out;
}

// Draw a gradient-filled shape. `poly` is given in unit coordinates.
function gradient(slide, x, y, w, h, opts) {
  const o = opts || {};
  const poly = o.poly || UNIT_RECT;
  const axis = axisFor(o.dark || 'br', w, h);
  const steps = o.steps || Math.max(12, Math.min(30, Math.round(Math.max(w, h) * 3)));
  const overlap = 0.7 / steps;
  for (let i = 0; i < steps; i++) {
    const band = clipPoly(clipPoly(poly, axis, i / steps - overlap, true), axis, (i + 1) / steps + overlap, false);
    if (band.length < 3) continue;
    const pts = band.map(function (p) {
      return { x: +(p[0] * w).toFixed(3), y: +(p[1] * h).toFixed(3) };
    });
    pts.push({ close: true });
    slide.addShape('custGeom', {
      x: x, y: y, w: w, h: h, points: pts,
      fill: { color: rampColor((i + 0.5) / steps) }, line: { type: 'none' }
    });
  }
}

const UNIT_RECT = [[0, 0], [1, 0], [1, 1], [0, 1]];

function unitEllipse(n) {
  const pts = [];
  const count = n || 48;
  for (let i = 0; i < count; i++) {
    const t = (2 * Math.PI * i) / count;
    pts.push([0.5 + 0.5 * Math.cos(t), 0.5 + 0.5 * Math.sin(t)]);
  }
  return pts;
}

// Parallelogram matching prstGeom "parallelogram" with adj=16326 (skew ~16.3%).
function unitParallelogram(w, h, skew) {
  const dx = ((skew === undefined ? 16326 : skew) / 100000) * Math.min(w, h) / w;
  return [[dx, 0], [1, 0], [1 - dx, 1], [0, 1]];
}

// Teardrop = circle with one square corner; `corner` names that corner.
function unitTeardrop(corner) {
  const turns = { tr: 0, br: 1, bl: 2, tl: 3 }[corner] || 0;
  const pts = [];
  for (let i = 0; i <= 36; i++) {
    const t = (i / 36) * 1.5 * Math.PI; // 270 deg of arc, then the corner
    pts.push([0.5 * Math.cos(t), 0.5 * Math.sin(t)]);
  }
  pts.push([0.5, -0.5]);
  return pts.map(function (p) {
    let x = p[0], y = p[1];
    for (let k = 0; k < turns; k++) { const nx = -y; y = x; x = nx; }
    return [x + 0.5, y + 0.5];
  });
}

// Teardrop with the brand gradient (custGeom) or a flat colour (native shape).
function teardrop(slide, x, y, size, corner, color) {
  if (color) {
    const rotate = { tr: 0, br: 90, bl: 180, tl: 270 }[corner] || 0;
    slide.addShape('teardrop', { x: x, y: y, w: size, h: size, rotate: rotate, fill: { color: color }, line: { type: 'none' } });
  } else {
    gradient(slide, x, y, size, size, { poly: unitTeardrop(corner), dark: 'br', steps: 10 });
  }
}

/* ------------------------------------------------------------ line icons */

// Outlined lotus flower: five pointed petals over a shallow bowl.
function lotusIcon(slide, x, y, size, color) {
  const ln = { color: color, width: 1.5 };
  const petal = function (cx, w, top, h, lean) {
    slide.addShape('custGeom', {
      x: x + cx * size, y: y + top * size, w: w * size, h: h * size,
      fill: { type: 'none' }, line: ln, rotate: lean,
      points: [
        { x: (w / 2) * size, y: 0 },
        { x: (w / 2) * size, y: h * size, curve: { type: 'cubic', x1: w * size, y1: h * 0.45 * size, x2: w * size, y2: h * 0.95 * size } },
        { x: (w / 2) * size, y: 0, curve: { type: 'cubic', x1: 0, y1: h * 0.95 * size, x2: 0, y2: h * 0.45 * size } },
        { close: true }
      ]
    });
  };
  petal(0.09, 0.26, 0.30, 0.44, -42); // outer left
  petal(0.65, 0.26, 0.30, 0.44, 42); // outer right
  petal(0.19, 0.26, 0.20, 0.55, -20); // inner left
  petal(0.55, 0.26, 0.20, 0.55, 20); // inner right
  petal(0.36, 0.28, 0.10, 0.66, 0); // centre
  slide.addShape('custGeom', {
    x: x, y: y + 0.63 * size, w: size, h: 0.30 * size,
    fill: { type: 'none' }, line: ln,
    points: [
      { x: 0.50 * size, y: 0.0 },
      { x: 0.50 * size, y: 0.30 * size, curve: { type: 'cubic', x1: 1.02 * size, y1: 0.02 * size, x2: 1.02 * size, y2: 0.28 * size } },
      { x: 0.50 * size, y: 0.0, curve: { type: 'cubic', x1: -0.02 * size, y1: 0.28 * size, x2: -0.02 * size, y2: 0.02 * size } },
      { close: true }
    ]
  });
}

// Outlined bathrobe: body with crossed lapels, sleeves and a waist belt.
function robeIcon(slide, x, y, size, color) {
  const s = size;
  const ln = { color: color, width: 1.5 };
  const path = function (pts) {
    slide.addShape('custGeom', { x: x, y: y, w: s, h: s, fill: { type: 'none' }, line: ln, points: pts.map(function (p) { return { x: p[0] * s, y: p[1] * s }; }) });
  };
  path([[0.30, 0.02], [0.70, 0.02], [0.90, 0.14], [0.98, 0.56], [0.82, 0.62], [0.82, 0.98], [0.18, 0.98], [0.18, 0.62], [0.02, 0.56], [0.10, 0.14], [0.30, 0.02]]);
  path([[0.30, 0.02], [0.50, 0.52], [0.70, 0.02]]); // outer lapel V
  path([[0.40, 0.05], [0.50, 0.38], [0.60, 0.05]]); // inner lapel V
  path([[0.18, 0.60], [0.82, 0.60]]); // belt
  path([[0.18, 0.68], [0.82, 0.68]]);
  path([[0.46, 0.68], [0.46, 0.90]]); // belt knot tails
  path([[0.54, 0.68], [0.54, 0.84]]);
}

// Solid handset used on the contact slide.
function phoneIcon(slide, x, y, size, color) {
  const s = size;
  slide.addShape('custGeom', {
    x: x, y: y, w: s, h: s, fill: { color: color }, line: { type: 'none' },
    points: [
      { x: 0.10 * s, y: 0.06 * s }, { x: 0.32 * s, y: 0.02 * s }, { x: 0.46 * s, y: 0.30 * s },
      { x: 0.32 * s, y: 0.42 * s }, { x: 0.58 * s, y: 0.70 * s }, { x: 0.70 * s, y: 0.56 * s },
      { x: 0.98 * s, y: 0.70 * s }, { x: 0.92 * s, y: 0.94 * s },
      { x: 0.16 * s, y: 0.36 * s, curve: { type: 'cubic', x1: 0.40 * s, y1: 1.02 * s, x2: 0.0, y2: 0.72 * s } },
      { close: true }
    ]
  });
}

// Solid envelope with a white fold line.
function mailIcon(slide, x, y, w, h, color) {
  slide.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: color }, line: { type: 'none' } });
  slide.addShape('custGeom', {
    x: x, y: y, w: w, h: h, fill: { color: C.white }, line: { type: 'none' },
    points: [{ x: 0.10 * w, y: 0.10 * h }, { x: 0.5 * w, y: 0.48 * h }, { x: 0.90 * w, y: 0.10 * h }, { close: true }]
  });
}

// Solid map pin: round head over a pointed tail.
function pinIcon(slide, x, y, w, h, color) {
  slide.addShape('ellipse', { x: x, y: y, w: w, h: w, fill: { color: color }, line: { type: 'none' } });
  slide.addShape('triangle', { x: x + 0.16 * w, y: y + 0.55 * h, w: 0.68 * w, h: 0.45 * h, rotate: 180, fill: { color: color }, line: { type: 'none' } });
  slide.addShape('ellipse', { x: x + 0.32 * w, y: y + 0.32 * w, w: 0.36 * w, h: 0.36 * w, fill: { color: C.white }, line: { type: 'none' } });
}

/* ----------------------------------------------------------------- charts */

const CHART_CATS = ['Category 1', 'Category 2', 'Category 3', 'Category 4'];

// Shared axis styling: no legend, hairline grey gridlines, Poppins labels.
const CHART_AXIS = {
  showLegend: false,
  chartColorsOpacity: 100,
  catAxisLabelFontFace: FONT_HEAD, catAxisLabelFontSize: 12, catAxisLabelColor: '262626',
  valAxisLabelFontFace: FONT_HEAD, valAxisLabelFontSize: 12, valAxisLabelColor: '595959',
  catAxisLineColor: C.greyLt, valAxisLineShow: false,
  catGridLine: { style: 'none' }, valGridLine: { color: C.greyLt, size: 0.75 },
  catAxisMajorTickMark: 'none', valAxisMajorTickMark: 'none'
};

/* ------------------------------------------------------------- the deck  */

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'LOVELYSPA', width: 20, height: 11.25 });
pptx.layout = 'LOVELYSPA';
pptx.author = 'irfanaffian';
pptx.title = 'Lovelyspa - Beauty Presentation';

/* -- 1. Cover ------------------------------------------------------------ */
function slideCover(s) {
  rect(s, 0, 0, 20, 9.146, C.grey);
  gradient(s, 0, 9.146, 20, 2.104, { dark: 'br' });
  text(s, 'LOVELYSPA', {
    x: 4.63, y: 4.607, w: 10.604, h: 2.036,
    fontFace: FONT_HEAD, fontSize: 115, bold: true, color: C.white, charSpacing: 6, align: 'center'
  });
  chip(s, 18.449, 3.55, 'v', C.white);
  [['PRESENTATION SLIDES', 1.23, 4.297], ['BEAUTY', 9.211, 1.578], ['MINIMALIST', 16.96, 2.406]]
    .forEach(function (t) { caps(s, t[1], 0.696, t[2], t[0], { color: C.white, spc: 6, wrap: false }); });
  caps(s, 1.23, 10.013, 2.299, 'TREATMENT', { color: C.white, spc: 6, wrap: false });
  caps(s, 16.747, 10.013, 2.619, 'IRFANAFFIAN', { color: C.white, spc: 6, wrap: false });
}

/* -- 2. Welcome ---------------------------------------------------------- */
function slideWelcome(s) {
  headline(s, 1.902, 2.224, 4.772, [['Welcome to ', 'Lovelyspa']]);
  gradient(s, 16.959, -2.839, 5.581, 5.667, { poly: unitEllipse(), dark: 'br' });
  text(s, [
    { text: 'BEAUTY', options: { color: C.pink, breakLine: true } },
    { text: 'PRESENTATION', options: { color: C.mute, breakLine: true } },
    { text: 'SLIDES', options: { color: C.mute } }
  ], { x: 1.902, y: 4.863, w: 4.196, h: 1.27, fontFace: FONT_BODY, fontSize: 16, bold: true, charSpacing: 3, lineSpacingMultiple: 1.5 });
  caps(s, 1.902, 9.132, 4.196, SPA_AND_COSMETIC);
  imagePlaceholder(s, 8.025, 3.364, 6.406, 4.167);
  body(s, 12.656, 8.749, 6.412, 1.27, LOREM_LONG);
  chip(s, 13.217, 4.405, 'h');
  year(s, 6.915, 6.425, 2.101, '2021', C.pink);
}

/* -- 3. History ---------------------------------------------------------- */
function slideHistory(s) {
  headline(s, 11.866, 2.682, 4.772, ['History of', 'Lovelyspa']);
  gradient(s, 16.959, -2.839, 5.581, 5.667, { poly: unitEllipse(), dark: 'br' });
  gradient(s, 0, 0, 8.317, 11.25, { dark: 'tr' });
  year(s, 6.297, 9.332, 2.135, '20-', C.white);
  imagePlaceholder(s, 3.703, 1.619, 6.406, 6.25);
  chip(s, 1.946, 2.624, 'h', C.white);
  text(s, [
    { text: 'BEAUTY', options: { breakLine: true } },
    { text: 'PRESENTATION', options: { breakLine: true } },
    { text: 'SLIDES', options: {} }
  ], { x: 1.121, y: 8.895, w: 4.196, h: 1.27, fontFace: FONT_BODY, fontSize: 16, bold: true, charSpacing: 3, color: C.white, lineSpacingMultiple: 1.5 });
  body(s, 11.866, 4.937, 6.093, 1.27, LOREM_LONG);
  body(s, 11.866, 6.569, 6.093, 0.866, LOREM_MID);
  caps(s, 15.553, 9.794, 3.663, SPA_AND_COSMETIC);
  year(s, 0.723, 2.682, 0.911, '21', C.white);
}

/* -- 4. About Our -------------------------------------------------------- */
function slideAbout(s) {
  headline(s, 10.138, 1.619, 4.772, ['About Our', 'Lovelyspa']);
  imagePlaceholder(s, 0, 0, 8.866, 11.25);
  caps(s, 10.693, 10.204, 3.663, SPA_AND_COSMETIC);
  gradient(s, 6.847, 4.577, 8.317, 4.755, { dark: 'tr' });
  body(s, 8.117, 5.572, 6.093, 1.27, LOREM_LONG, { color: C.white });
  body(s, 8.117, 7.355, 6.093, 0.866, LOREM_MID, { color: C.white });
  chip(s, 18.326, 4.889, 'v');
  caps(s, 17.737, 10.204, 1.505, 'BEAUTY', { color: C.pink, align: 'center' });
}

/* -- 5. Place Of Service ------------------------------------------------- */
function slidePlace(s) {
  headline(s, 2.209, 1.281, 6.487, [['Place Of Service ', 'Lovelyspa']]);
  gradient(s, -0.925, 4.667, 1.851, 1.879, { poly: unitEllipse(), dark: 'br' });
  imagePlaceholder(s, 3.912, 4.173, 8.729, 4.755);
  caps(s, 13.845, 7.477, 3.663, SPA_AND_COSMETIC);
  chip(s, 1.226, 7.336, 'h');
  caps(s, 13.845, 8.253, 1.505, 'BEAUTY', { color: C.pink });
  body(s, 13.845, 4.671, 4.768, 1.674, LOREM_LONG);
  year(s, 17.254, 1.685, 1.651, '2021', C.pink);
  body(s, 5.927, 9.614, 6.714, 0.866, LOREM_TINY);
}

/* -- 6. Greetings For ---------------------------------------------------- */
function slideGreetings(s) {
  headline(s, 11.354, 2.031, 6.487, [['Greetings For ', 'Lovelyspa']]);
  gradient(s, 2.828, 5.845, 1.851, 1.879, { poly: unitEllipse(), dark: 'br' });
  imagePlaceholder(s, 3.912, 1.325, 5.739, 7.604);
  caps(s, 7.921, 10.028, 3.663, SPA_AND_COSMETIC);
  chip(s, 8.552, 7.496, 'h');
  caps(s, 12.857, 10.028, 1.505, 'BEAUTY', { color: C.pink });
  body(s, 11.584, 5.127, 6.257, 1.27, LOREM_LONG);
  year(s, 0.925, 9.744, 1.651, '2021', C.pink);
  body(s, 0.925, 1.434, 1.902, 3.289, LOREM_TINY);
  caps(s, 15.412, 10.028, 3.663, 'PRESENTATION');
  gradient(s, 19.544, 3.318, 0.847, 0.86, { poly: unitEllipse(), dark: 'br' });
}

/* -- 7. Profile Of ------------------------------------------------------- */
function slideProfile(s) {
  headline(s, 1.209, 2.322, 6.487, ['Profile Of', 'Lovelyspa']);
  imagePlaceholder(s, 12.36, 1.325, 7.64, 7.604);
  caps(s, 1.559, 9.857, 3.663, SPA_AND_COSMETIC);
  chip(s, 10.294, 2.165, 'h');
  caps(s, 6.494, 9.857, 1.505, 'BEAUTY', { color: C.pink });
  body(s, 4.923, 6.077, 6.153, 1.27, LOREM_LONG);
  caps(s, 9.049, 9.857, 3.663, 'PRESENTATION');
  gradient(s, -0.423, 7.89, 0.847, 0.86, { poly: unitEllipse(), dark: 'br' });
  gradient(s, 15.477, 6.495, 4.523, 4.755, { dark: 'tr' });
  body(s, 16.394, 7.89, 2.688, 2.078, LOREM_TINY, { color: C.white });
}

/* -- 8. Therapy Of (two panels) ------------------------------------------ */
function slideTherapyPanels(s) {
  headline(s, 6.042, 1.325, 4.2, ['Therapy Of', 'Lovelyspa']);
  gradient(s, 1.571, 0.671, 1.851, 1.879, { poly: unitEllipse(), dark: 'br' });
  imagePlaceholder(s, 12.36, 4.173, 7.64, 4.755);
  caps(s, 5.017, 9.857, 3.663, SPA_AND_COSMETIC);
  caps(s, 12.976, 9.857, 1.505, 'BEAUTY', { color: C.pink });
  caps(s, 16.748, 9.857, 2.405, 'PRESENTATION');
  gradient(s, 3.755, 2.361, 0.847, 0.86, { poly: unitEllipse(), dark: 'br' });
  gradient(s, 0, 4.173, 4.316, 4.755, { dark: 'tr' });
  imagePlaceholder(s, 4.3, 4.173, 7.64, 4.755);
  chip(s, 2.29, 7.431, 'h', C.white);
  body(s, 14.042, 1.548, 4.641, 1.27, LOREM_TINY);
  body(s, 0.814, 5.086, 2.688, 1.27, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.', { color: C.white });
}

/* -- 9. Therapy Of (stacked) --------------------------------------------- */
function slideTherapyStacked(s) {
  headline(s, 1.201, 2.456, 4.2, ['Therapy Of', 'Lovelyspa']);
  imagePlaceholder(s, 9.192, 1.981, 7.64, 2.665);
  caps(s, 10.656, 9.084, 1.505, 'BEAUTY', { color: C.pink });
  caps(s, 13.629, 9.084, 2.405, 'PRESENTATION');
  imagePlaceholder(s, 4.322, 6.495, 3.663, 4.755);
  body(s, 10.656, 6.958, 6.323, 1.27, LOREM_LONG);
  chip(s, 1.457, 5.693, 'v');
  subtitle(s, 16.034, 3.888, '1. Subtitle Here');
}

/* -- 10. Type of Theraphy ------------------------------------------------ */
function slideTypeOf(s) {
  caps(s, 9.803, 10.203, 1.505, 'BEAUTY', { color: C.pink });
  caps(s, 12.776, 10.203, 2.405, 'PRESENTATION');
  imagePlaceholder(s, 2.546, 6.495, 5.439, 4.755);
  body(s, 6.509, 3.303, 6.323, 1.27, LOREM_LONG);
  chip(s, 3.192, 3.347, 'v');
  subtitle(s, 3.643, 1.57, '1. Subtitle Here');
  headline(s, 9.803, 7.508, 6.851, ['Type of Theraphy', 'Lovelyspa']);
  body(s, 14.066, 3.303, 3.371, 1.27, LOREM_SHORT);
  subtitle(s, 14.066, 1.57);
  gradient(s, 18.957, 7.308, 1.043, 1.348, { dark: 'tr', steps: 8 });
}

/* -- 11. Nature Theraphy ------------------------------------------------- */
function slideNature(s) {
  caps(s, 16.866, 8.236, 1.505, 'BEAUTY', { color: C.pink });
  caps(s, 16.866, 6.785, 2.405, 'PRESENTATION');
  imagePlaceholder(s, 0, 0, 5.439, 6.25);
  chip(s, 1.058, 5.168, 'v');
  subtitle(s, 2.896, 7.253);
  headline(s, 7.233, 1.809, 6.851, ['Nature Theraphy', 'Lovelyspa']);
  body(s, 15.181, 2.175, 3.371, 1.27, LOREM_SHORT);
  gradient(s, 7.808, 5.875, 8.317, 3.765, { dark: 'tr' });
  body(s, 8.805, 6.569, 6.323, 1.27, LOREM_LONG, { color: C.white });
  body(s, 8.805, 8.08, 6.323, 0.866, LOREM_MID, { color: C.white });
  subtitle(s, 2.9, 8.236, 'Subtitle Here', C.greyLt);
}

/* -- 12. Theraphy of (card on grey) -------------------------------------- */
function slideCard(s) {
  rect(s, 1.063, 1.052, 17.875, 9.146, C.grey);
  rect(s, 4.502, 2.933, 10.996, 5.627, C.white);
  headline(s, 5.447, 3.655, 5.111, ['Theraphy of', 'Lovelyspa']);
  [5.447, 10.747].forEach(function (x) {
    body(s, x, 6.391, 4.196, 1.27, LOREM_SHORT);
    subtitle(s, x, 5.987);
  });
  teardrop(s, 14.161, 3.498, 0.782, 'tl');
  teardrop(s, 14.161, 4.51, 0.782, 'bl');
}

/* -- 13. We Prepare Before ----------------------------------------------- */
function slidePrepare(s) {
  caps(s, 11.283, 2.754, 1.505, 'BEAUTY', { color: C.pink });
  caps(s, 11.283, 1.808, 2.405, 'PRESENTATION');
  s.addShape('teardrop', { x: 11.283, y: 4.846, w: 5.083, h: 5.083, rotate: 270, fill: { color: C.grey }, line: { type: 'none' } });
  chip(s, 18.078, 3.82, 'v');
  headline(s, 1.84, 7.007, 6.851, ['We Prepare ', ['Before ', 'Lovelyspa']]);
  body(s, 1.84, 1.812, 3.371, 1.27, LOREM_SHORT);
  body(s, 1.84, 9.063, 6.093, 0.866, LOREM_SHORT);
  teardrop(s, 1.84, 4.843, 1.043, 'tl');
  teardrop(s, 3.217, 4.843, 1.043, 'tr');
}

// Grey parallelogram used as photo placeholder on slides 14-16 and 25.
function slantPlaceholder(s, x, y, w, h, color) {
  const poly = unitParallelogram(w, h);
  s.addShape('custGeom', {
    x: x, y: y, w: w, h: h, fill: { color: color || C.grey }, line: { type: 'none' },
    points: poly.map(function (p) { return { x: +(p[0] * w).toFixed(3), y: +(p[1] * h).toFixed(3) }; }).concat([{ close: true }])
  });
}

/* -- 14. Takecare Your Skin (right slant) -------------------------------- */
function slideTakecareRight(s) {
  subtitle(s, 3.872, 1.503);
  headline(s, 10.571, 3.163, 6.851, ['Takecare Your', ['Skin in ', 'Lovelyspa']]);
  body(s, 1.603, 8.697, 4.196, 1.27, LOREM_SHORT);
  slantPlaceholder(s, 8.569, 6.195, 9.174, 5.055);
  subtitle(s, 6.997, 1.503);
  chip(s, 7.029, 7.155, 'h');
  gradient(s, 2.211, 4.268, 1.851, 1.879, { poly: unitEllipse(), dark: 'br' });
  gradient(s, 4.395, 5.958, 0.847, 0.86, { poly: unitEllipse(), dark: 'br' });
}

/* -- 15. Takecare Your Skin (two slants) --------------------------------- */
function slideTakecareTwo(s) {
  caps(s, 6.844, 10.102, 1.505, 'BEAUTY', { color: C.pink });
  caps(s, 2.844, 10.102, 2.405, 'PRESENTATION');
  subtitle(s, 13.27, 9.547);
  headline(s, 2.018, 2.065, 6.851, ['Takecare Your', ['Skin in ', 'Lovelyspa']]);
  body(s, 11.04, 6.322, 4.196, 1.27, LOREM_SHORT);
  slantPlaceholder(s, 12.797, 1.643, 5.271, 3.652);
  subtitle(s, 16.395, 9.547);
  chip(s, 10.508, 2.126, 'h');
  slantPlaceholder(s, 4.43, 5.548, 4.827, 3.345);
  gradient(s, 4.608, 5.195, 0.847, 0.86, { poly: unitEllipse(), dark: 'br' });
  body(s, 0.834, 7.201, 2.436, 1.674, LOREM_SHORT);
}

/* -- 16. Our Service of (slants) ----------------------------------------- */
function slideServiceSlants(s) {
  caps(s, 13.719, 9.236, 1.505, 'BEAUTY', { color: C.pink });
  caps(s, 9.719, 9.236, 2.405, 'PRESENTATION');
  headline(s, 14.034, 3.116, 5.111, [['Our Service of ', 'Lovelyspa']]);
  body(s, 10.87, 7.032, 5.893, 0.866, LOREM_SHORT);
  teardrop(s, 1.323, 6.629, 0.782, 'tl');
  teardrop(s, 1.323, 7.642, 0.782, 'bl');
  slantPlaceholder(s, 5.623, -0.026, 7.947, 5.506);
  slantPlaceholder(s, 3.236, 6.715, 4.64, 3.214);
  gradient(s, -0.741, -0.826, 1.851, 1.879, { poly: unitEllipse(), dark: 'br' });
  gradient(s, 5.267, 4.979, 0.847, 0.86, { poly: unitEllipse(), dark: 'br' });
  body(s, 1.11, 2.751, 3.414, 1.27, LOREM_SHORT);
}

/* -- 17. Our Service of (circles) ---------------------------------------- */
function slideServiceCircles(s) {
  gradient(s, 3.744, 4.306, 1.851, 1.879, { poly: unitEllipse(), dark: 'br' });
  caps(s, 13.765, 1.626, 1.505, 'BEAUTY', { color: C.pink });
  caps(s, 10.483, 1.725, 2.405, 'PRESENTATION');
  gradient(s, 8.196, 8.16, 1.332, 1.352, { poly: unitEllipse(), dark: 'br' });
  gradient(s, 11.2, 4.235, 1.332, 1.352, { poly: unitEllipse(), dark: 'br' });
  headline(s, 3.204, 1.427, 5.111, [['Our Service of ', 'Lovelyspa']]);
  [[1.984, 4.483, 3.363, 3.415], [8.315, 6.826, 2.556, 2.595], [11.8, 3.807, 2.174, 2.208]]
    .forEach(function (c) { s.addShape('ellipse', { x: c[0], y: c[1], w: c[2], h: c[3], fill: { color: C.grey }, line: { type: 'none' } }); });
  [[14.63, 4.109, 3.663, 1.27], [1.984, 8.367, 3.663, 1.27], [11.561, 7.593, 4.525, 0.866]]
    .forEach(function (t) {
      subtitle(s, t[0], t[1]);
      body(s, t[0], t[1] + 0.469, t[2], t[3], LOREM_SHORT);
    });
}

/* -- 18. Portfolio of (big circle) --------------------------------------- */
function slidePortfolioCircle(s) {
  caps(s, 9.711, 1.822, 1.505, 'BEAUTY', { color: C.pink });
  body(s, 9.703, 5.918, 6.153, 1.27, LOREM_LONG);
  year(s, 17.253, 9.713, 1.651, '2021', C.pink);
  caps(s, 6.483, 1.822, 2.405, 'PRESENTATION');
  headline(s, 9.703, 3.273, 5.111, [['Portfolio of ', 'Lovelyspa']]);
  subtitle(s, 15.215, 9.949, 'Subtitle Here', C.greyLt);
  s.addShape('ellipse', { x: -2.123, y: 2.706, w: 10.042, h: 10.196, fill: { color: C.grey }, line: { type: 'none' } });
  body(s, 9.703, 7.26, 6.153, 0.866, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt.');
  chip(s, 18.078, 1.461, 'v');
}

/* -- 19. Portfolio of (offset circles) ----------------------------------- */
function slidePortfolioOffset(s) {
  caps(s, 9.703, 9.947, 3.663, SPA_AND_COSMETIC);
  caps(s, 12.942, 9.081, 1.505, 'BEAUTY', { color: C.pink });
  body(s, 9.703, 4.44, 6.153, 1.27, LOREM_LONG);
  year(s, 1.199, 1.098, 1.651, '2021', C.pink);
  caps(s, 9.713, 9.081, 2.405, 'PRESENTATION');
  gradient(s, 1.404, 2.702, 5.996, 6.089, { poly: unitEllipse(), dark: 'br' });
  headline(s, 9.703, 2.26, 5.111, [['Portfolio of ', 'Lovelyspa']]);
  subtitle(s, 3.123, 1.259, 'Subtitle Here', C.greyLt);
  s.addShape('ellipse', { x: 1.778, y: 3.082, w: 5.724, h: 5.812, fill: { color: C.grey }, line: { type: 'none' } });
  chip(s, 1.987, 9.4, 'h');
  gradient(s, 16.996, 1.025, 1.86, 1.889, { poly: unitEllipse(), dark: 'br' });
  [9.711, 13.132].forEach(function (x) {
    year(s, x, 6.331, 1.651, '2021', C.greyLt);
    caps(s, x + 0.073, 7.08, 1.505, 'BEAUTY', { color: C.pink, align: 'center' });
  });
}

/* -- 20. Portfolio of (squares + icons) ---------------------------------- */
function slidePortfolioSquares(s) {
  caps(s, 9.933, 9.812, 1.505, 'BEAUTY', { color: C.pink });
  body(s, 10.366, 5.013, 6.635, 1.27, LOREM_LONG);
  caps(s, 6.705, 9.812, 2.405, 'PRESENTATION');
  headline(s, 10.366, 2.834, 5.111, [['Portfolio of ', 'Lovelyspa']]);
  imagePlaceholder(s, 3.344, 2.364, 5.288, 5.812);
  chip(s, 14.432, 9.529, 'h');
  gradient(s, 7.671, 1.418, 1.86, 1.889, { poly: unitEllipse(), dark: 'br' });
  imagePlaceholder(s, 1.199, 6.254, 4.291, 4.996, C.greyLt);
  robeIcon(s, 13.659, 6.72, 1.278, C.pink);
  lotusIcon(s, 10.466, 6.814, 1.278, C.pink);
  year(s, 1.199, 4.816, 1.651, '2021', C.pink);
}

/* -- 21. Break slides ---------------------------------------------------- */
function slideBreak(s) {
  rect(s, 0, 2.104, 20, 9.146, C.grey);
  text(s, 'BREAK SLIDES', {
    x: 4.843, y: 5.045, w: 12.33, h: 2.036,
    fontFace: FONT_HEAD, fontSize: 115, bold: true, color: C.white, charSpacing: 6, align: 'center'
  });
  chip(s, 1.23, 3.998, 'v', C.white);
  [['PRESENTATION SLIDES', 1.23, 4.297], ['BEAUTY', 9.211, 1.578], ['MINIMALIST', 16.96, 2.406]]
    .forEach(function (t) { caps(s, t[1], 10.026, t[2], t[0], { color: C.white, spc: 6, wrap: false }); });
  gradient(s, 0, 0, 20, 2.104, { dark: 'br' });
  caps(s, 1.23, 0.867, 2.299, 'TREATMENT', { color: C.white, spc: 6, wrap: false });
  caps(s, 16.747, 0.867, 2.619, 'IRFANAFFIAN', { color: C.white, spc: 6, wrap: false });
}

// Smartphone mock-up: silver rim, thin dark bezel, grey screen, top notch.
function phoneMock(s, x, y, w, h) {
  s.addShape('roundRect', { x: x, y: y, w: w, h: h, rectRadius: 0.4, fill: { color: 'B4B4B6' }, line: { type: 'none' } });
  s.addShape('roundRect', { x: x + w * 0.022, y: y + h * 0.010, w: w * 0.956, h: h * 0.980, rectRadius: 0.38, fill: { color: '565658' }, line: { type: 'none' } });
  s.addShape('roundRect', { x: x + w * 0.048, y: y + h * 0.022, w: w * 0.904, h: h * 0.956, rectRadius: 0.35, fill: { color: C.greyLt }, line: { type: 'none' } });
  s.addShape('roundRect', { x: x + w * 0.31, y: y + h * 0.022, w: w * 0.38, h: h * 0.030, rectRadius: 0.05, fill: { color: '565658' }, line: { type: 'none' } });
}

/* -- 22. Mock Up Device (phones) ----------------------------------------- */
function slideMockPhones(s) {
  gradient(s, 0, 0, 8.317, 11.25, { dark: 'tr' });
  headline(s, 1.386, 1.79, 6.851, ['Mock Up', 'Device'], { color: C.white, accent: C.white });
  body(s, 1.343, 4.142, 6.093, 1.27, LOREM_LONG, { color: C.white });
  phoneMock(s, 10.271, 1.703, 3.946, 7.629);
  phoneMock(s, 15.081, 3.165, 3.168, 6.124);
  lotusIcon(s, 1.862, 7.2, 0.743, C.white);
  robeIcon(s, 5.055, 7.185, 0.743, C.white);
  body(s, 1.386, 8.192, 2.624, 1.27, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit', { color: C.white });
  body(s, 4.492, 8.192, 2.624, 1.27, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit', { color: C.white });
  caps(s, 17.988, 10.312, 1.505, 'BEAUTY', { color: C.pink });
  caps(s, 14.759, 10.312, 2.405, 'PRESENTATION');
}

/* -- 23. Mock Up Device (desktop) ---------------------------------------- */
function slideMockDesktop(s) {
  gradient(s, 0, 5.0, 13.833, 6.25, { dark: 'tr' });
  caps(s, 1.234, 10.022, 1.505, 'BEAUTY', { color: C.white });
  caps(s, 4.109, 10.022, 2.405, 'PRESENTATION', { color: C.white });
  // All-in-one desktop: black bezel, grey screen, silver chin and stand.
  rect(s, 1.984, 1.641, 6.602, 3.99, '050505');
  imagePlaceholder(s, 2.22, 1.88, 6.14, 3.48);
  rect(s, 1.984, 5.6, 6.602, 0.9, 'E9EAEB');
  s.addShape('trapezoid', { x: 4.44, y: 6.5, w: 1.72, h: 1.06, rotate: 180, fill: { color: 'DCDCDE' }, line: { type: 'none' } });
  s.addShape('roundRect', { x: 4.14, y: 7.5, w: 2.32, h: 0.12, rectRadius: 0.06, fill: { color: 'C9C9CC' }, line: { type: 'none' } });
  body(s, 10.654, 2.045, 6.153, 1.27, LOREM_LONG);
  [6.506, 8.81].forEach(function (y) {
    year(s, 16.271, y, 1.651, '2021', C.greyLt);
    caps(s, 16.344, y + 0.75, 1.505, 'BEAUTY', { color: C.pink, align: 'center' });
  });
  headline(s, 8.727, 8.81, 3.648, ['Mock Up', 'Device'], { color: C.white, accent: C.white });
}

/* -- 24. Mock Up Device (laptop) ----------------------------------------- */
function slideMockLaptop(s) {
  gradient(s, 0, 0, 13.833, 11.25, { dark: 'tr' });
  body(s, 6.33, 8.531, 6.153, 1.27, LOREM_LONG, { color: C.white });
  [2.645, 6.33].forEach(function (x) {
    [1.948, 3.85].forEach(function (y) {
      year(s, x, y, 1.651, '2021', C.white);
      caps(s, x + 0.073, y + 0.749, 1.505, 'BEAUTY', { color: C.white, align: 'center' });
    });
  });
  headline(s, 1.178, 8.308, 3.648, ['Mock Up', 'Device'], { color: C.white, accent: C.white });
  // Laptop mock-up: rounded dark lid, grey screen, wide silver base.
  s.addShape('roundRect', { x: 9.82, y: 1.13, w: 7.18, h: 4.36, rectRadius: 0.1, fill: { color: '0B0B0C' }, line: { type: 'none' } });
  imagePlaceholder(s, 10.06, 1.4, 6.68, 3.92);
  s.addShape('trapezoid', { x: 9.003, y: 5.49, w: 8.759, h: 0.36, fill: { color: 'D2D3D6' }, line: { type: 'none' } });
  s.addShape('roundRect', { x: 12.55, y: 5.49, w: 1.7, h: 0.16, rectRadius: 0.06, fill: { color: 'A9AAAE' }, line: { type: 'none' } });
  subtitle(s, 15.636, 7.006);
  subtitle(s, 15.636, 9.165);
}

/* -- 25. Pricelist ------------------------------------------------------- */
function slidePricelist(s) {
  chip(s, 1.568, 1.333, 'h');
  subtitle(s, 1.568, 9.613);
  headline(s, 1.44, 2.934, 6.851, ['Pricelist of', 'Lovelyspa']);
  body(s, 1.568, 8.041, 4.196, 1.27, LOREM_SHORT);
  [6.859, 12.438].forEach(function (x, i) {
    slantPlaceholder(s, x + 0.893, 1.233, 5.654, 3.115);
    gradient(s, x, 4.836, 5.958, 5.427, { poly: unitParallelogram(5.958, 5.427), dark: 'tr' });
    text(s, 'Bronze', { x: x + 1.706, y: 5.322, w: 2.403, h: 0.707, fontFace: FONT_HEAD, fontSize: 36, bold: true, color: C.white, align: 'center' });
    body(s, x + 1.278, 6.28, 3.323, 1.27, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod', { color: C.white, align: 'center' });
    body(s, x + 1.318, 7.766, 3.323, 0.462, 'Lorem ipsum dolor', { color: C.white, align: 'center', bold: true });
    body(s, x + 1.318, 8.225, 3.323, 0.462, 'Lorem ipsum dolor', { color: C.white, align: 'center', bold: true });
    s.addShape('roundRect', { x: x + 1.622, y: 9.186, w: 2.571, h: 0.746, rectRadius: 0.2, fill: { color: C.white }, line: { type: 'none' } });
    text(s, '$200', { x: x + 1.706, y: 9.311, w: 2.403, h: 0.572, fontFace: FONT_HEAD, fontSize: 28, bold: true, color: C.pink, align: 'center' });
  });
}

/* -- 26. Chart of (stacked area) ----------------------------------------- */
function slideChartArea(s) {
  const dates = ['1', '2', '3', '4', '5'];
  const series = [
    { name: 'Series 1', labels: dates, values: [32, 32, 28, 12, 15] },
    { name: 'Series 2', labels: dates, values: [12, 12, 12, 21, 28] }
  ];
  const areaOpts = { barGrouping: 'stacked', chartColors: [C.roseFlat, C.greyLt] };
  // Declared as a single-plot combo chart so the area meets the axes exactly
  // as in the original (crossBetween="midCat").
  s.addChart([{ type: 'area', data: series, options: areaOpts }],
    Object.assign({ x: 1.5, y: 1.803, w: 7.906, h: 5.271, catAxisHidden: true, valAxisMaxVal: 50 }, CHART_AXIS, areaOpts));
  chip(s, 12.999, 2.201, 'h');
  headline(s, 3.016, 7.912, 6.851, ['Chart of', 'Lovelyspa']);
  lotusIcon(s, 11.261, 4.776, 0.796, C.pink);
  robeIcon(s, 11.275, 7.856, 0.796, C.pink);
  [4.754, 7.733].forEach(function (y) {
    subtitle(s, 12.727, y);
    body(s, 12.727, y + 0.403, 6.153, 1.27, LOREM_LONG);
  });
}

/* -- 27. Chart of (line) -------------------------------------------------- */
function slideChartLine(s) {
  caps(s, 2.367, 9.262, 1.505, 'BEAUTY', { color: C.pink });
  gradient(s, 11.733, 0.984, 6.95, 9.003, { dark: 'tr' });
  s.addChart('line', [
    { name: 'Series 1', labels: CHART_CATS, values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels: CHART_CATS, values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels: CHART_CATS, values: [2.0, 2.0, 3.0, 5.0] }
  ], Object.assign({ x: 1.79, y: 4.022, w: 6.619, h: 4.413, chartColors: ['4472C4', 'ED7D31', 'A5A5A5'], lineSize: 2.25, lineDataSymbol: 'none' }, CHART_AXIS));
  headline(s, 2.392, 1.476, 6.851, ['Chart of', 'Lovelyspa']);
  body(s, 12.795, 5.091, 4.478, 1.674, LOREM_LONG, { color: C.white });
  [12.795, 15.505].forEach(function (x) {
    year(s, x, 2.902, 1.651, '2021', C.white);
    caps(s, x + 0.073, 3.652, 1.505, 'BEAUTY', { color: C.white, align: 'center' });
  });
  body(s, 12.795, 7.433, 4.478, 0.866, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod', { color: C.white });
  caps(s, 5.757, 9.263, 2.405, 'PRESENTATION');
}

/* -- 28. Chart of (100% stacked bar) ------------------------------------- */
function slideChartBar(s) {
  body(s, 11.924, 8.66, 6.435, 1.27, LOREM_LONG);
  gradient(s, 0, 2.751, 6.865, 8.499, { dark: 'tr' });
  s.addChart('bar', [
    { name: 'Series 2', labels: CHART_CATS, values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels: CHART_CATS, values: [2.0, 2.0, 3.0, 5.0] }
  ], Object.assign({
    x: 8.993, y: 1.553, w: 8.905, h: 5.937, barDir: 'bar', barGrouping: 'percentStacked',
    chartColors: [C.roseFlat, 'F2F2F2'], valAxisLabelFormatCode: '0%'
  }, CHART_AXIS));
  headline(s, 1.206, 3.832, 4.218, ['Chart of', 'Lovelyspa'], { color: C.white, accent: C.white });
  lotusIcon(s, 1.206, 6.203, 0.796, C.white);
  robeIcon(s, 1.22, 8.464, 0.796, C.white);
  [6.369, 8.528].forEach(function (y) {
    subtitle(s, 2.39, y, 'Subtitle Here', C.white);
    body(s, 2.412, y + 0.355, 3.433, 0.866, 'Lorem ipsum dolor sit amet, consectetur adipiscing', { color: C.white });
  });
  caps(s, 0.823, 1.552, 1.505, 'BEAUTY', { color: C.pink });
  caps(s, 4.213, 1.553, 2.405, 'PRESENTATION');
}

/* -- 29. Get In Touch ----------------------------------------------------- */
function slideContact(s) {
  // Gradient panel whose top-left corner is snipped off diagonally.
  gradient(s, 9.6, 0, 10.4, 7.68, {
    dark: 'br',
    poly: [[0.32, 0], [1, 0], [1, 1], [0, 1], [0, 0.44]]
  });
  caps(s, 15.971, 9.873, 1.505, 'BEAUTY', { color: C.pink });
  caps(s, 15.971, 8.625, 2.405, 'PRESENTATION');
  body(s, 15.521, 3.118, 2.997, 2.078, LOREM_MID, { color: C.white });
  imagePlaceholder(s, 6.903, 2.484, 6.875, 7.171);
  headline(s, 1.359, 2.859, 4.014, ['Get In ', 'Touch']);
  const rows = [
    ['Phone', '083 4432 135 23', 5.979, phoneIcon],
    ['Email', 'yourcompany@email.com', 7.022, mailIcon],
    ['Address', '298 Secret Street, UNA', 8.077, pinIcon]
  ];
  rows.forEach(function (r) {
    text(s, r[0], { x: 2.427, y: r[2], w: 2.924, h: 0.438, fontFace: FONT_HEAD, fontSize: 20, bold: true, color: C.slate });
    text(s, r[1], { x: 2.427, y: r[2] + 0.287, w: 4.014, h: 0.462, fontFace: FONT_BODY, fontSize: 16, color: C.mute, lineSpacingMultiple: 1.5 });
  });
  phoneIcon(s, 1.547, 6.073, 0.365, C.slate);
  mailIcon(s, 1.528, 7.181, 0.385, 0.271, C.slate);
  pinIcon(s, 1.556, 8.186, 0.284, 0.362, C.slate);
}

/* -- 30. Thank you -------------------------------------------------------- */
function slideThankYou(s) {
  rect(s, 0, 0, 14.202, 11.25, C.grey);
  [['PRESENTATION SLIDES', 1.205, 4.297], ['BEAUTY', 7.372, 1.578], ['MINIMALIST', 10.572, 2.406]]
    .forEach(function (t) { caps(s, t[1], 0.816, t[2], t[0], { color: C.white, spc: 6, wrap: false }); });
  gradient(s, 14.207, 0, 5.793, 11.25, { dark: 'bl' });
  chip(s, 15.038, 3.431, 'h', C.white);
  text(s, [
    { text: 'THANK', options: { breakLine: true } },
    { text: 'YOU', options: {} }
  ], { x: 1.092, y: 6.564, w: 7.368, h: 3.972, fontFace: FONT_HEAD, fontSize: 115, bold: true, color: C.white, charSpacing: 6, valign: 'top', margin: [3.6, 7.2, 3.6, 7.2] });
  body(s, 15.0, 5.539, 2.976, 2.078, LOREM_MID, { color: C.white });
  teardrop(s, 17.407, 9.29, 0.782, 'tl', C.white);
  teardrop(s, 18.42, 9.29, 0.782, 'tr', C.white);
}

/* -- 31-33. Icon sheets --------------------------------------------------- */
// The deck ends with three reference sheets, each an 11 x 6 grid of 66 solid
// dark-grey pictograms. Each icon here is a short recipe of native shapes in
// unit coordinates: [shape, x, y, w, h, {r: rotate, k: knock-out white}].
const ICON_COLOR = '404040';
const ICON_GRID = { x0: 2.33, dx: 1.532, y0: 2.15, dy: 1.405, size: 0.66 };

const ICONS = {
  phone: [['teardrop', 0.06, 0.06, 0.88, 0.88, { r: 225 }], ['ellipse', 0.26, 0.26, 0.66, 0.66, { k: 1 }]],
  mobile: [['roundRect', 0.28, 0.02, 0.44, 0.96], ['rect', 0.34, 0.14, 0.32, 0.62, { k: 1 }], ['ellipse', 0.44, 0.82, 0.12, 0.10, { k: 1 }]],
  mail: [['rect', 0.04, 0.24, 0.92, 0.52], ['triangle', 0.04, 0.24, 0.92, 0.34, { r: 180, k: 1 }]],
  send: [['triangle', 0.02, 0.06, 0.96, 0.88, { r: 135 }]],
  pencil: [['rect', 0.18, 0.10, 0.30, 0.74, { r: 45 }], ['triangle', 0.60, 0.62, 0.28, 0.30, { r: 180 }]],
  clip: [['roundRect', 0.24, 0.06, 0.40, 0.88], ['roundRect', 0.36, 0.20, 0.16, 0.58, { k: 1 }]],
  pin: [['ellipse', 0.18, 0.02, 0.64, 0.64], ['triangle', 0.30, 0.42, 0.40, 0.56, { r: 180 }], ['ellipse', 0.38, 0.20, 0.24, 0.24, { k: 1 }]],
  map: [['parallelogram', 0.02, 0.14, 0.96, 0.72], ['rect', 0.36, 0.14, 0.06, 0.72, { k: 1 }], ['rect', 0.62, 0.14, 0.06, 0.72, { k: 1 }]],
  compass: [['ellipse', 0.02, 0.02, 0.96, 0.96], ['triangle', 0.30, 0.20, 0.40, 0.60, { r: 30, k: 1 }]],
  target: [['ellipse', 0.02, 0.02, 0.96, 0.96], ['ellipse', 0.18, 0.18, 0.64, 0.64, { k: 1 }], ['ellipse', 0.38, 0.38, 0.24, 0.24]],
  warning: [['triangle', 0.0, 0.06, 1.0, 0.88], ['rect', 0.46, 0.36, 0.08, 0.30, { k: 1 }], ['rect', 0.46, 0.72, 0.08, 0.08, { k: 1 }]],
  share: [['ellipse', 0.62, 0.02, 0.34, 0.34], ['ellipse', 0.02, 0.32, 0.34, 0.34], ['ellipse', 0.62, 0.64, 0.34, 0.34], ['rect', 0.30, 0.26, 0.40, 0.06, { r: 20 }], ['rect', 0.30, 0.68, 0.40, 0.06, { r: -20 }]],
  star: [['star5', 0.0, 0.02, 1.0, 0.96]],
  thumb: [['roundRect', 0.36, 0.10, 0.56, 0.44], ['rect', 0.06, 0.44, 0.88, 0.48], ['rect', 0.06, 0.44, 0.24, 0.48, { k: 1 }], ['rect', 0.04, 0.46, 0.26, 0.46]],
  info: [['ellipse', 0.02, 0.02, 0.96, 0.96], ['rect', 0.44, 0.40, 0.12, 0.40, { k: 1 }], ['rect', 0.44, 0.20, 0.12, 0.12, { k: 1 }]],
  chat: [['roundRect', 0.02, 0.10, 0.96, 0.62], ['triangle', 0.20, 0.62, 0.24, 0.30, { r: 180 }]],
  quote: [['rect', 0.08, 0.18, 0.32, 0.44], ['rect', 0.56, 0.18, 0.32, 0.44], ['triangle', 0.14, 0.56, 0.20, 0.28, { r: 180 }], ['triangle', 0.62, 0.56, 0.20, 0.28, { r: 180 }]],
  question: [['ellipse', 0.02, 0.02, 0.96, 0.96], ['donut', 0.28, 0.14, 0.44, 0.44, { k: 1 }], ['rect', 0.44, 0.46, 0.12, 0.16, { k: 1 }], ['rect', 0.44, 0.70, 0.12, 0.12, { k: 1 }]],
  home: [['triangle', 0.0, 0.06, 1.0, 0.46], ['rect', 0.16, 0.48, 0.68, 0.46], ['rect', 0.42, 0.66, 0.18, 0.28, { k: 1 }]],
  copy: [['rect', 0.06, 0.06, 0.60, 0.60], ['rect', 0.34, 0.34, 0.60, 0.60], ['rect', 0.40, 0.40, 0.48, 0.48, { k: 1 }]],
  search: [['donut', 0.04, 0.04, 0.66, 0.66], ['rect', 0.62, 0.62, 0.34, 0.12, { r: 45 }]],
  flag: [['rect', 0.14, 0.02, 0.09, 0.96], ['rect', 0.23, 0.10, 0.62, 0.40]],
  gear: [['gear6', 0.0, 0.0, 1.0, 1.0], ['ellipse', 0.36, 0.36, 0.28, 0.28, { k: 1 }]],
  tools: [['rect', 0.10, 0.10, 0.18, 0.80, { r: 35 }], ['rect', 0.72, 0.10, 0.18, 0.80, { r: -35 }]],
  trophy: [['roundRect', 0.22, 0.04, 0.56, 0.52], ['rect', 0.44, 0.54, 0.12, 0.26], ['rect', 0.24, 0.80, 0.52, 0.14]],
  tag: [['diamond', 0.02, 0.02, 0.96, 0.96], ['ellipse', 0.20, 0.20, 0.18, 0.18, { k: 1 }]],
  camera: [['roundRect', 0.02, 0.22, 0.96, 0.62], ['rect', 0.28, 0.10, 0.32, 0.14], ['ellipse', 0.32, 0.34, 0.36, 0.36, { k: 1 }]],
  megaphone: [['triangle', 0.06, 0.14, 0.86, 0.72, { r: 90 }], ['rect', 0.06, 0.34, 0.16, 0.32]],
  moon: [['moon', 0.16, 0.02, 0.72, 0.96]],
  leaf: [['teardrop', 0.06, 0.06, 0.88, 0.88, { r: 135 }]],
  music: [['ellipse', 0.06, 0.62, 0.36, 0.30], ['rect', 0.38, 0.06, 0.10, 0.72], ['rect', 0.38, 0.06, 0.52, 0.16, { r: 12 }]],
  doc: [['rect', 0.14, 0.02, 0.72, 0.96], ['rect', 0.26, 0.24, 0.48, 0.07, { k: 1 }], ['rect', 0.26, 0.44, 0.48, 0.07, { k: 1 }], ['rect', 0.26, 0.64, 0.30, 0.07, { k: 1 }]],
  cup: [['trapezoid', 0.12, 0.14, 0.76, 0.72, { r: 180 }], ['rect', 0.12, 0.06, 0.76, 0.12]],
  plane: [['triangle', 0.02, 0.06, 0.96, 0.88, { r: 45 }]],
  lifebuoy: [['donut', 0.02, 0.02, 0.96, 0.96], ['ellipse', 0.30, 0.30, 0.40, 0.40, { k: 1 }], ['rect', 0.46, 0.02, 0.08, 0.30], ['rect', 0.46, 0.68, 0.08, 0.30]],
  eye: [['ellipse', 0.0, 0.22, 1.0, 0.56], ['ellipse', 0.32, 0.30, 0.36, 0.40, { k: 1 }], ['ellipse', 0.40, 0.38, 0.20, 0.24]],
  clock: [['ellipse', 0.02, 0.02, 0.96, 0.96], ['rect', 0.46, 0.18, 0.08, 0.34, { k: 1 }], ['rect', 0.46, 0.46, 0.28, 0.08, { k: 1 }]],
  mic: [['roundRect', 0.32, 0.02, 0.36, 0.56], ['donut', 0.18, 0.34, 0.64, 0.46, { k: 0 }], ['rect', 0.46, 0.76, 0.08, 0.22]],
  calendar: [['rect', 0.04, 0.12, 0.92, 0.82], ['rect', 0.12, 0.30, 0.76, 0.52, { k: 1 }], ['rect', 0.22, 0.02, 0.10, 0.20], ['rect', 0.68, 0.02, 0.10, 0.20]],
  bolt: [['lightningBolt', 0.14, 0.0, 0.72, 1.0]],
  cloud: [['cloud', 0.0, 0.16, 1.0, 0.68]],
  drop: [['teardrop', 0.10, 0.06, 0.80, 0.80, { r: 315 }]],
  gauge: [['blockArc', 0.02, 0.06, 0.96, 0.96], ['rect', 0.46, 0.28, 0.08, 0.34, { r: 30 }]],
  globe: [['ellipse', 0.02, 0.02, 0.96, 0.96], ['ellipse', 0.30, 0.02, 0.40, 0.96, { k: 1 }], ['rect', 0.02, 0.44, 0.96, 0.10, { k: 1 }]],
  key: [['donut', 0.02, 0.20, 0.56, 0.56], ['rect', 0.50, 0.42, 0.48, 0.12], ['rect', 0.80, 0.50, 0.09, 0.24]],
  bucket: [['trapezoid', 0.10, 0.16, 0.80, 0.78, { r: 180 }], ['rect', 0.28, 0.02, 0.44, 0.14]],
  magnet: [['blockArc', 0.02, 0.06, 0.96, 0.96], ['rect', 0.06, 0.72, 0.24, 0.24], ['rect', 0.70, 0.72, 0.24, 0.24]],
  drive: [['roundRect', 0.02, 0.30, 0.96, 0.40], ['ellipse', 0.72, 0.42, 0.16, 0.16, { k: 1 }]],
  trash: [['trapezoid', 0.14, 0.22, 0.72, 0.72, { r: 180 }], ['rect', 0.06, 0.10, 0.88, 0.12], ['rect', 0.36, 0.02, 0.28, 0.08]],
  rocket: [['teardrop', 0.22, 0.02, 0.56, 0.72, { r: 315 }], ['triangle', 0.02, 0.52, 0.30, 0.34, { r: 200 }], ['triangle', 0.68, 0.52, 0.30, 0.34, { r: 160 }]],
  brush: [['rect', 0.20, 0.02, 0.24, 0.62, { r: 30 }], ['triangle', 0.44, 0.62, 0.34, 0.36, { r: 200 }]],
  user: [['ellipse', 0.28, 0.02, 0.44, 0.44], ['blockArc', 0.06, 0.42, 0.88, 0.98]],
  card: [['rect', 0.02, 0.18, 0.96, 0.64], ['rect', 0.10, 0.30, 0.34, 0.08, { k: 1 }], ['rect', 0.10, 0.46, 0.34, 0.08, { k: 1 }], ['ellipse', 0.60, 0.30, 0.28, 0.28, { k: 1 }]],
  heart: [['heart', 0.02, 0.06, 0.96, 0.88]],
  plus: [['plus', 0.04, 0.04, 0.92, 0.92]],
  minus: [['rect', 0.04, 0.42, 0.92, 0.16]],
  check: [['rect', 0.06, 0.46, 0.42, 0.14, { r: 45 }], ['rect', 0.36, 0.14, 0.60, 0.14, { r: -40 }]],
  cross: [['rect', 0.04, 0.42, 0.92, 0.16, { r: 45 }], ['rect', 0.04, 0.42, 0.92, 0.16, { r: -45 }]],
  play: [['triangle', 0.06, 0.04, 0.88, 0.92, { r: 90 }]],
  arrowUp: [['upArrow', 0.18, 0.02, 0.64, 0.96]],
  arrowRight: [['rightArrow', 0.02, 0.18, 0.96, 0.64]],
  arrowBoth: [['leftRightArrow', 0.02, 0.24, 0.96, 0.52]],
  refresh: [['donut', 0.06, 0.06, 0.88, 0.88], ['rect', 0.50, 0.02, 0.48, 0.34, { k: 1 }], ['triangle', 0.56, 0.02, 0.40, 0.34, { r: 90 }]],
  sun: [['sun', 0.0, 0.0, 1.0, 1.0]],
  lock: [['rect', 0.14, 0.42, 0.72, 0.56], ['donut', 0.26, 0.02, 0.48, 0.56], ['rect', 0.26, 0.34, 0.48, 0.24, { k: 1 }], ['rect', 0.14, 0.42, 0.72, 0.56]],
  square: [['rect', 0.08, 0.08, 0.84, 0.84], ['rect', 0.22, 0.22, 0.56, 0.56, { k: 1 }]]
};

const ICON_NAMES = Object.keys(ICONS);

function drawIcon(s, name, cx, cy, size) {
  ICONS[name].forEach(function (part) {
    const shape = part[0], o = part[5] || {};
    s.addShape(shape, {
      x: cx - size / 2 + part[1] * size,
      y: cy - size / 2 + part[2] * size,
      w: part[3] * size, h: part[4] * size,
      rotate: o.r || 0,
      fill: { color: o.k ? C.white : ICON_COLOR },
      line: { type: 'none' }
    });
  });
}

// Each sheet walks the icon vocabulary from a different starting point so the
// three pages read as three distinct 66-icon reference sheets.
function slideIconSheet(s, sheet) {
  for (let r = 0; r < 6; r++) {
    for (let c = 0; c < 11; c++) {
      const name = ICON_NAMES[(sheet * 66 + r * 11 + c) % ICON_NAMES.length];
      drawIcon(s, name, ICON_GRID.x0 + c * ICON_GRID.dx, ICON_GRID.y0 + r * ICON_GRID.dy, ICON_GRID.size);
    }
  }
}

/* ----------------------------------------------------------------- build */

const BUILDERS = [
  slideCover, slideWelcome, slideHistory, slideAbout, slidePlace,
  slideGreetings, slideProfile, slideTherapyPanels, slideTherapyStacked, slideTypeOf,
  slideNature, slideCard, slidePrepare, slideTakecareRight, slideTakecareTwo,
  slideServiceSlants, slideServiceCircles, slidePortfolioCircle, slidePortfolioOffset, slidePortfolioSquares,
  slideBreak, slideMockPhones, slideMockDesktop, slideMockLaptop, slidePricelist,
  slideChartArea, slideChartLine, slideChartBar, slideContact, slideThankYou,
  function (s) { slideIconSheet(s, 0); },
  function (s) { slideIconSheet(s, 1); },
  function (s) { slideIconSheet(s, 2); }
];

BUILDERS.forEach(function (build) { build(pptx.addSlide()); });

pptx.writeFile({ fileName: path.join(__dirname, '03e16a47-b60c-4352-974a-8cb22b430e33_grok_final.pptx') })
  .then(function (f) { console.log('wrote ' + f); });
