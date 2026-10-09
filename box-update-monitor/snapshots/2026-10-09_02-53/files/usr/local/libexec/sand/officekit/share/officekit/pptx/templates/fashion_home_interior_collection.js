/**
 * "Altair — Home Decor" square (12.5in x 12.5in) presentation, 25 slides.
 * Rebuilt with pptxgenjs only. Run: node <thisfile>.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Design tokens
 * ------------------------------------------------------------------ */
const SIDE = 12.5; // slide is a 12.5in square

const INK = '3D3D3D'; // near-black body text
const SAND = 'CDC1B1'; // warm accent bands
const SHELL = 'ECE7E2'; // pale accent bands
const MIST = 'EEEEEE'; // cool light grey band
const PAPER = 'FFFFFF';
const PHONE_BODY = '000000';
const PHONE_TRIM = '1A1A1C';
const PHONE_BTN = 'A6A6A6';

const BLACK = 'Montserrat Black';
const SEMI = 'Montserrat SemiBold';
const REG = 'Montserrat';
const LIGHT = 'Montserrat Light';

/* Photo frames of each source layout: [x, y, w, h] in inches.
 * They are emitted as empty picture placeholders, exactly like the
 * reference deck, so a user can drop imagery in without re-laying out. */
const PHOTO_FRAMES = {
  L01: [[6.139, 3.385, 6.361, 9.115]],
  L02: [[2.0, 4.975, 10.5, 4.378]],
  L03: [[0.818, 4.578, 8.161, 7.922]],
  L04: [[2.016, 6.458, 3.807, 6.042], [6.839, 6.458, 5.661, 6.042]],
  L05: [[0.0, 2.243, 3.778, 8.014]],
  L06: [[0.0, 3.148, 4.889, 8.014], [5.889, 3.148, 6.611, 8.014]],
  L07: [[2.444, 8.139, 4.569, 4.361], [7.931, 8.139, 4.569, 4.361], [2.444, 2.319, 10.056, 4.859]],
  L08: [[2.611, 2.613, 7.278, 5.332]],
  L09: [[1.306, 1.306, 9.889, 6.611]],
  L10: [[7.611, 1.997, 4.889, 6.028], [0.0, 1.997, 6.694, 6.028]],
  L11: [[6.676, 2.272, 3.715, 7.957]],
  L12: [[4.292, 0.0, 3.681, 4.31], [4.292, 5.051, 3.681, 7.449], [8.819, 0.0, 3.681, 12.5]],
  L13: [[2.819, 0.0, 8.847, 12.5]],
  L14: [[0.0, 0.0, 9.403, 12.5]],
  L15: [[0.0, 6.25, 12.5, 6.25]],
  L16: [[3.986, 2.319, 8.514, 7.861]],
  L17: [[1.639, 0.0, 5.306, 10.806]],
  L18: [[7.194, 1.075, 5.306, 10.349]],
  L19: [[2.306, 0.0, 7.889, 3.125]],
  L20: [[2.306, 0.0, 7.889, 3.125], [2.306, 5.472, 7.889, 7.028]],
  L21: [[2.569, 2.569, 7.361, 7.361]],
  L22: [[1.401, 2.569, 3.604, 9.931]],
  L23: [[2.5, 0.0, 7.5, 12.5]],
  L24: [[0.0, 0.0, 10.0, 12.5]],
  BLANK: [],
};

/* ------------------------------------------------------------------ *
 * Drawing helpers
 * ------------------------------------------------------------------ */

/** Flat colour rectangle (the deck's only decorative primitive). */
function band(s, x, y, w, h, color) {
  s.addShape('rect', { x, y, w, h, fill: { color }, line: { type: 'none' } });
}

/**
 * Text box. `content` is a string, or an array of strings = one paragraph each.
 * o: { x, y, w, h, font, size, color, spc, align, line, italic, nowrap, rot, num }
 */
function txt(s, content, o) {
  const paras = Array.isArray(content) ? content : [content];
  const runOpts = o.num ? { bullet: { type: 'number', style: 'arabicPeriod', indent: 36 } } : {};
  const body = paras.map((t, i) =>
    Object.assign({ text: t }, { options: Object.assign({ breakLine: i < paras.length - 1 }, runOpts) })
  );
  s.addText(body, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: o.font || REG,
    fontSize: o.size || 18,
    color: o.color || INK,
    charSpacing: o.spc,
    align: o.align || 'left',
    valign: 'top',
    italic: o.italic,
    wrap: !o.nowrap,
    lineSpacingMultiple: o.line,
    rotate: o.rot,
    autoFit: true,
  });
}

/**
 * Flat-illustration smartphone mock-up used on slide 12, drawn from native
 * shapes. The screen is left unfilled so the sand panel behind shows through,
 * exactly as in the reference artwork. Coordinates are absolute inches.
 */
function phoneMockup(s) {
  const fill = (kind, x, y, w, h, color, rad) =>
    s.addShape(kind, { x, y, w, h, fill: { color }, line: { type: 'none' }, rectRadius: rad });

  // side buttons (silent switch, volume pair, power) alongside the shell
  [[6.425, 3.180, 0.315], [6.425, 3.780, 0.605],
   [6.425, 4.555, 0.600], [10.605, 3.990, 0.960]]
    .forEach(([x, y, h]) => fill('roundRect', x, y, 0.030, h, PHONE_BTN, 0.014));

  // shell: a hollow rounded rectangle stroked 0.21in wide, so the screen is see-through
  s.addShape('roundRect', {
    x: 6.590, y: 2.185, w: 3.880, h: 8.130, rectRadius: 0.495,
    fill: { type: 'none' }, line: { color: PHONE_BODY, width: 15.1 },
  });

  fill('roundRect', 7.510, 2.150, 2.045, 0.430, PHONE_BODY, 0.145); // notch
  fill('roundRect', 8.275, 2.344, 0.511, 0.060, PHONE_TRIM, 0.030); // earpiece grille
  fill('ellipse', 8.921, 2.314, 0.124, 0.120, PHONE_TRIM);          // front camera

  // status bar: clock at the left, signal / wi-fi / battery at the right
  s.addText('9:41', {
    x: 7.030, y: 2.390, w: 0.60, h: 0.24, margin: 0,
    fontFace: SEMI, fontSize: 11, color: PHONE_BODY, align: 'left', valign: 'middle',
  });
  [0.038, 0.057, 0.079, 0.106].forEach((barH, i) => // rising signal bars
    fill('roundRect', 9.575 + i * 0.0445, 2.566 - barH, 0.023, barH, PHONE_BODY, 0.006));
  [[9.789, 2.460, 0.146, 0.104], [9.820, 2.502, 0.084, 0.060]].forEach(([x, y, w, h]) =>
    s.addShape('blockArc', { // the two wi-fi waves
      x, y, w, h, angleRange: [200, 340], arcThicknessRatio: 0.45,
      fill: { color: PHONE_BODY }, line: { type: 'none' },
    }));
  s.addShape('triangle', { // wi-fi dot
    x: 9.848, y: 2.542, w: 0.029, h: 0.026,
    fill: { color: PHONE_BODY }, line: { type: 'none' }, flipV: true,
  });
  s.addShape('roundRect', { // battery: 40%-opacity outline around a solid cell
    x: 9.992, y: 2.478, w: 0.203, h: 0.088, rectRadius: 0.024,
    fill: { type: 'none' }, line: { color: PHONE_BODY, width: 1.4, transparency: 60 },
  });
  fill('roundRect', 10.008, 2.490, 0.170, 0.064, PHONE_BODY, 0.016);
  s.addShape('rect', {
    x: 10.205, y: 2.503, w: 0.014, h: 0.022,
    fill: { color: PHONE_BODY, transparency: 60 }, line: { type: 'none' },
  });
}

/* ------------------------------------------------------------------ *
 * Slides — each entry is [masterName, builder]
 * ------------------------------------------------------------------ */
const SLIDES = [

  // 1 — cover: right shell column, pull-quote bottom left
  ['L01', s => {
    band(s, 9.417, 0.0, 3.083, 12.625, SHELL);
    txt(s, 'Altair', { x: 0.759, y: 0.647, w: 1.45, h: 0.438, size: 20, color: SAND, spc: 6, nowrap: true });
    txt(s, 'slidestation.co', { x: 7.184, y: 0.701, w: 4.742, h: 0.337, font: LIGHT, size: 14, spc: 15, align: 'right', nowrap: true });
    txt(s, 'The interior of the house personifies the private world; the exterior of it is part of the outside world.',
      { x: 0.759, y: 8.584, w: 2.977, h: 3.227, font: SEMI, spc: 3, line: 1.5 });
  }],

  // 2 — "Home Decor" title over a wide landscape photo
  ['L02', s => {
    txt(s, 'Altair', { x: 1.968, y: 1.881, w: 1.45, h: 0.438, size: 20, color: SAND, spc: 6, nowrap: true });
    txt(s, 'Home Decor', { x: 1.905, y: 3.387, w: 5.889, h: 0.841, font: BLACK, size: 44, spc: 6 });
    txt(s, 'Fashion and interior design are one and the same.',
      { x: 1.905, y: 10.004, w: 9.104, h: 0.615, spc: 3, line: 2 });
  }],

  // 3 — quote with a shell column on the right, rotated web address
  ['L03', s => {
    band(s, 8.979, 0.0, 3.521, 12.5, SHELL);
    txt(s, 'Altair', { x: 0.762, y: 0.724, w: 1.45, h: 0.438, size: 20, color: SAND, spc: 6, nowrap: true });
    txt(s, 'Great personal style is an extreme curiosity about yourself.',
      { x: 0.762, y: 2.385, w: 6.391, h: 1.345, font: SEMI, size: 20, spc: 3, line: 2 });
    txt(s, 'slidestation.co', { x: 9.618, y: 9.68, w: 4.742, h: 0.337, font: LIGHT, size: 14, spc: 15, align: 'right', nowrap: true, rot: 90 });
  }],

  // 4 — "Home Collection", sand rule at the top, two photos below
  ['L04', s => {
    band(s, 2.016, 0.0, 10.484, 0.325, SAND);
    txt(s, 'FALL \u201822', { x: 2.0, y: 1.867, w: 2.036, h: 0.438, size: 20, color: SAND, spc: 6, nowrap: true });
    txt(s, 'Home Collection', { x: 2.0, y: 3.655, w: 7.979, h: 0.841, font: BLACK, size: 44, spc: 6 });
    txt(s, 'Fashion and interior design are one and the same.',
      { x: 2.0, y: 4.496, w: 9.431, h: 0.672, size: 20, spc: 3, line: 2 });
  }],

  // 5 — "FALL EVENT" invitation, photo strip left, sand rule right
  ['L05', s => {
    band(s, 12.167, 2.243, 0.333, 8.014, SAND);
    txt(s, 'Organized By Altair', { x: 6.035, y: 2.028, w: 3.875, h: 0.443, size: 12, spc: 3, align: 'center', line: 2, italic: true });
    txt(s, ['FALL', 'EVENT'], { x: 5.028, y: 4.986, w: 5.889, h: 1.919, font: BLACK, size: 54, spc: 6, align: 'center' });
    txt(s, '09/18/2021', { x: 6.877, y: 7.076, w: 2.19, h: 0.37, size: 16, color: SAND, spc: 6, align: 'center', nowrap: true });
    txt(s, 'Your Address Comes Here', { x: 5.306, y: 9.876, w: 5.333, h: 0.443, size: 12, spc: 3, align: 'center', line: 2, italic: true });
  }],

  // 6 — sand block top-left, statement right, two photos, sand footer rule
  ['L06', s => {
    txt(s, 'Interior design is a business of trust\u2026',
      { x: 5.889, y: 1.241, w: 5.694, h: 1.05, font: SEMI, size: 20, spc: 6, line: 1.5 });
    band(s, 0.0, 0.0, 4.889, 2.167, SAND);
    band(s, 5.889, 12.143, 6.611, 0.357, SAND);
  }],

  // 7 — full-height sand rail with two rotated labels, three photos
  ['L07', s => {
    band(s, 0.0, 0.0, 2.444, 12.5, SAND);
    txt(s, 'Home Decor', { x: 0.052, y: 9.713, w: 3.694, h: 0.546, font: SEMI, size: 20, spc: 6, align: 'right', line: 1.5, rot: 270 });
    txt(s, 'FALL \u201822', { x: 0.052, y: 3.894, w: 3.694, h: 0.546, font: SEMI, size: 20, color: SHELL, spc: 6, align: 'right', line: 1.5, rot: 270 });
  }],

  // 8 — centred pull-quote inside a big shell square
  ['BLANK', s => {
    band(s, 1.306, 1.306, 9.889, 9.889, SHELL);
    txt(s, 'slidestation.co', { x: 1.347, y: 0.481, w: 9.889, h: 0.337, size: 14, spc: 15, align: 'center' });
    txt(s, '\u201C', { x: 5.583, y: 3.901, w: 1.333, h: 1.079, font: SEMI, size: 44, spc: 3, align: 'center', line: 1.5 });
    txt(s, 'Architecture is basically the design of interiors, the art of organizing interior space.',
      { x: 2.569, y: 5.132, w: 7.361, h: 1.846, font: SEMI, size: 24, spc: 3, align: 'center', line: 1.5 });
    txt(s, ['-', 'Philip Johnson'], { x: 4.347, y: 7.452, w: 3.806, h: 0.861, font: LIGHT, size: 16, spc: 3, align: 'center', line: 1.5, italic: true });
    txt(s, 'VISIT OUR WEBSITE!', { x: 2.278, y: 11.668, w: 8.139, h: 0.337, size: 14, spc: 15, align: 'center' });
  }],

  // 9 — sand picture frame (four bands) around a centred photo
  ['L08', s => {
    [[0.0, 0.0, 12.5, 1.306], [0.0, 11.194, 12.5, 1.306],
     [11.194, 0.0, 1.306, 12.5], [0.0, 0.0, 1.306, 12.5]]
      .forEach(([x, y, w, h]) => band(s, x, y, w, h, SAND));
    txt(s, 'AUTUMN COLLECTION', { x: 4.403, y: 8.59, w: 3.694, h: 1.05, font: SEMI, size: 20, spc: 6, align: 'center', line: 1.5 });
    txt(s, '2022', { x: 5.347, y: 9.949, w: 2.083, h: 0.37, size: 16, color: SAND, spc: 20, align: 'center' });
  }],

  // 10 — photo on top, stacked centred captions underneath
  ['L09', s => {
    txt(s, 'DECOR \u201822', { x: 3.347, y: 9.298, w: 5.806, h: 0.37, size: 16, color: SAND, spc: 10, align: 'center' });
    txt(s, 'NEW COLLECTION', { x: 1.708, y: 9.744, w: 9.083, h: 0.723, font: SEMI, size: 28, spc: 6, align: 'center', line: 1.5 });
    txt(s, 'Fashion and Interior Design', { x: 1.806, y: 10.628, w: 8.889, h: 0.457, size: 16, spc: 3, align: 'center', line: 1.5 });
  }],

  // 11 — full-height sand panel right, two photos, long quote below
  ['L10', s => {
    band(s, 7.611, 0.0, 4.889, 12.5, SAND);
    txt(s, 'DECOR \u201822', { x: 0.903, y: 0.837, w: 3.681, h: 0.37, size: 16, color: SAND, spc: 10 });
    txt(s, 'There is one spectacle grander than the sea, that is the sky; there is one spectacle grander than the sky, that is the interior of the soul.',
      { x: 1.5, y: 9.339, w: 9.792, h: 1.846, size: 24, spc: 3, line: 1.5 });
  }],

  // 12 — copy + numbered remarks left, phone mock-up over a sand panel
  ['L11', s => {
    band(s, 8.569, 0.0, 3.931, 12.5, SAND);
    txt(s, 'The interior of the house personifies the private world; the exterior of it is part of the outside world.',
      { x: 1.641, y: 2.12, w: 2.668, h: 3.789, size: 16, spc: 3, line: 2 });
    txt(s, 'Remarks:', { x: 1.641, y: 7.391, w: 2.848, h: 0.501, font: SEMI, spc: 3, line: 1.5 });
    txt(s, ['Interior', 'Personifies', 'World', 'Outside', 'Private'],
      { x: 1.641, y: 8.362, w: 2.848, h: 2.117, size: 16, spc: 3, line: 1.5, num: true });
    phoneMockup(s);
  }],

  // 13 — three-column photo grid with a rotated headline
  ['L12', s => {
    txt(s, 'Home Collection', { x: -0.896, y: 5.829, w: 7.979, h: 0.841, font: BLACK, size: 44, spc: 6, align: 'center', rot: 270 });
  }],

  // 14 — edge-to-edge photo, rotated statement, sand edge right
  ['L13', s => {
    txt(s, 'Interior design is a business of trust\u2026',
      { x: -2.725, y: 6.833, w: 9.627, h: 0.546, font: SEMI, size: 20, spc: 6, line: 1.5, rot: 270 });
    band(s, 11.667, 0.0, 0.833, 12.5, SAND);
  }],

  // 15 — full-bleed photo left, two rotated labels on the white margin
  ['L14', s => {
    txt(s, 'Altair \u201822', { x: 9.845, y: 2.153, w: 2.171, h: 0.404, font: SEMI, color: SAND, spc: 6, nowrap: true, rot: 90 });
    txt(s, 'slidestation.co', { x: 8.875, y: 9.378, w: 4.111, h: 0.337, font: LIGHT, size: 14, spc: 12, align: 'right', nowrap: true, rot: 90 });
  }],

  // 16 — "MOOD" over a three-tone horizontal swatch strip
  ['L15', s => {
    txt(s, 'New Collection', { x: 1.722, y: 0.597, w: 9.056, h: 0.337, size: 14, color: SAND, spc: 12, align: 'center' });
    txt(s, 'MOOD', { x: 3.278, y: 4.579, w: 6.861, h: 0.546, font: SEMI, size: 20, spc: 60, align: 'center', line: 1.5 });
    [SHELL, SAND, MIST].forEach((c, i) => band(s, i * 4.167, 5.583, 4.167, 0.208, c));
  }],

  // 17 — sand tab left, big photo right
  ['L16', s => {
    txt(s, 'New Collection', { x: 3.917, y: 1.363, w: 7.444, h: 0.404, color: SAND, spc: 12 });
    txt(s, 'slidestation.co', { x: 3.917, y: 10.733, w: 3.915, h: 0.303, font: LIGHT, size: 12, spc: 12, nowrap: true });
    band(s, 0.0, 2.319, 0.917, 7.861, SAND);
  }],

  // 18 — tall photo centre-left, caption column right
  ['L17', s => {
    txt(s, 'slidestation.co', { x: -0.764, y: 2.086, w: 3.915, h: 0.303, font: LIGHT, size: 12, spc: 12, align: 'right', nowrap: true, rot: 270 });
    txt(s, 'New Collection', { x: 8.028, y: 5.165, w: 2.668, h: 1.221, font: SEMI, color: SAND, spc: 6, line: 2 });
    txt(s, 'The interior of the house and personifies the private world; the exterior of it is part of the outside world.',
      { x: 8.028, y: 7.114, w: 2.668, h: 3.789, size: 16, spc: 3, line: 2 });
  }],

  // 19 — "MOOD" with four vertical swatches, photo right
  ['L18', s => {
    txt(s, 'MOOD', { x: 1.653, y: 5.226, w: 4.722, h: 0.546, font: SEMI, size: 20, spc: 60, align: 'center', line: 1.5 });
    [SHELL, SAND, MIST, INK].forEach((c, i) => band(s, 2.4302 + i * 0.7083, 1.0753, 0.2438, 3.3934, c));
    txt(s, 'The interior of the house and personifies the private world; the',
      { x: 1.785, y: 9.859, w: 3.674, h: 1.635, size: 16, spc: 3, align: 'center', line: 2 });
  }],

  // 20 — banner photo, swatch rule, spaced "HOME DECOR" headline
  ['L19', s => {
    [[2.3056, 1.5642, SHELL], [4.6146, 3.2708, SAND], [8.6302, 1.5642, SHELL]]
      .forEach(([x, w, c]) => band(s, x, 3.3806, w, 0.2438, c));
    txt(s, 'HOME DECOR', { x: 0.874, y: 6.288, w: 11.389, h: 1.243, font: SEMI, size: 40, spc: 45, align: 'center', line: 2 });
    txt(s, 'The interior of the house personifies the private world; the exterior of it',
      { x: 2.552, y: 10.408, w: 7.396, h: 1.096, size: 16, spc: 6, align: 'center', line: 2 });
  }],

  // 21 — two stacked photos separated by an even three-tone swatch rule
  ['L20', s => {
    [[2.3056, SHELL], [5.2448, SAND], [8.1840, INK]]
      .forEach(([x, c]) => band(s, x, 4.1306, 2.0104, 0.2438, c));
  }],

  // 22 — square photo inset in a sand square
  ['L21', s => {
    band(s, 1.306, 1.306, 9.889, 9.889, SAND);
    txt(s, 'New Collection', { x: 2.528, y: 1.753, w: 7.444, h: 0.404, spc: 12, align: 'center' });
    txt(s, 'slidestation.co', { x: 4.293, y: 11.675, w: 3.915, h: 0.303, font: LIGHT, size: 12, spc: 12, align: 'center', nowrap: true });
  }],

  // 23 — tall photo left, long quote on a sand panel right
  ['L22', s => {
    band(s, 6.407, 1.306, 4.692, 9.889, SAND);
    txt(s, ['New', 'Collection'], { x: 1.385, y: 0.88, w: 3.621, h: 0.955, spc: 12, align: 'center', line: 1.5 });
    txt(s, 'The work of art, just like any fragment of human life considered in its deepest meaning, seems to me devoid of value if it does not offer the hardness, the rigidity, the regularity, the luster on every interior and exterior facet, of the crystal.',
      { x: 7.165, y: 3.172, w: 3.176, h: 6.155, size: 14, spc: 3, align: 'center', line: 2 });
  }],

  // 24 — full-height photo behind a wide shell banner carrying the closing quote
  ['L23', s => {
    band(s, 0.0, 3.172, 12.5, 6.155, SHELL);
    txt(s, 'slidestation.co', { x: 4.336, y: 0.249, w: 3.915, h: 0.303, font: LIGHT, size: 12, spc: 12, align: 'center', nowrap: true });
    txt(s, 'Home Decor', { x: 2.313, y: 4.931, w: 8.137, h: 0.9, font: SEMI, size: 28, spc: 12, align: 'center', line: 2 });
    txt(s, 'Inspiration comes from everywhere: books, art, people on the street. It is an interior process for me.',
      { x: 1.489, y: 6.018, w: 9.522, h: 1.221, spc: 3, align: 'center', line: 2 });
  }],

  // 25 — same closing quote on a half-width sand panel
  ['L24', s => {
    band(s, 6.25, 3.172, 6.25, 6.155, SAND);
    txt(s, 'Inspiration comes from everywhere: books, art, people on the street. It is an interior process for me.',
      { x: 7.343, y: 4.731, w: 4.063, h: 3.038, spc: 3, align: 'center', line: 2 });
  }],
];

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */
function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'SQUARE', width: SIDE, height: SIDE });
  pptx.layout = 'SQUARE';
  pptx.title = 'Altair — Home Decor';

  Object.keys(PHOTO_FRAMES).forEach(name => {
    pptx.defineSlideMaster({
      title: name,
      background: { color: PAPER },
      objects: PHOTO_FRAMES[name].map(([x, y, w, h], i) => ({
        placeholder: { options: { name: `photo${i + 1}`, type: 'image', x, y, w, h }, text: '' },
      })),
    });
  });

  SLIDES.forEach(([master, draw]) => draw(pptx.addSlide({ masterName: master })));

  return pptx.writeFile({
    fileName: path.join(__dirname, '064fd2f0-0869-4b87-8772-c296e7fd7610_grok_final.pptx'),
  });
}

build().then(f => console.log('wrote', f)).catch(e => { console.error(e); process.exit(1); });
