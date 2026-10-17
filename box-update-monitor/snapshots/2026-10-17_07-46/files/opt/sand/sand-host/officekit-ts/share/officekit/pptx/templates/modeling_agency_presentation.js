/**
 * "MODELLING" — Model Presentation Template (20 slides, 16:9).
 *
 * Standalone pptxgenjs re-creation of the reference deck.
 * Raster photos in the original are replaced with flat grey placeholder shapes
 * that keep the original silhouette (arch / pill / ellipse / rectangle).
 *
 * Run: node 0e17bc93-1630-4bbc-b9cc-3deff6f82235_grok_final.js
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

const pptx = new PptxGenJS();
const S = pptx.ShapeType;

/* ------------------------------------------------------------------ palette */

const C = {
  bg: 'FEF1E0', // theme accent3/accent6 — slide background
  peach: 'F7C799', // accent1
  maroon: '78261A', // accent2
  dark: '262626', // tx1 lumMod 85% / lumOff 15%
  black: '000000',
  white: 'FFFFFF',
  peachDark: 'F0943C', // accent1 lumMod 75%
  peachDeep: 'B9620F', // accent1 lumMod 50%
  peachSoft: 'FCE9D6', // accent1 lumMod 40% + lumOff 60%
  peachPale: 'FDF4EB', // accent1 lumMod 20% + lumOff 80%
  maroonDark: '5A1D13', // accent2 lumMod 75%
  maroonDeep: '3C130D', // accent2 lumMod 50%
  maroonSoft: 'E59387', // accent2 lumMod 40% + lumOff 60%
  maroonPale: 'F2C9C3', // accent2 lumMod 20% + lumOff 80%
  track: 'F2F2F2', // bg1 lumMod 95% — progress-bar track
  photo: 'CCCCCC', // stand-in for the deck's raster photography
  photoInk: 'E4E4E4',
};

const HEAD = 'Abril Fatface';
const BODY = 'Work Sans';

const LOREM =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
  'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud ' +
  'exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ';
const DUIS = 'Lorem ipsum dolor sit amet ipsum. Duis aute irure .';
const MAECENAS =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue porttitor massa. ';

/* ------------------------------------------------------------------ helpers */

/** New slide on the cream deck background. */
function newSlide() {
  const s = pptx.addSlide();
  s.background = { color: C.bg };
  return s;
}

/**
 * Text box. Reference text boxes are top-anchored; labels that live inside a
 * shape (`opts.shape`) are centred, matching PowerPoint's own defaults.
 */
function tbox(slide, content, opts) {
  slide.addText(content, Object.assign({ valign: opts.shape ? 'middle' : 'top' }, opts));
}

/** Solid shape. */
function box(slide, shape, x, y, w, h, color, extra) {
  slide.addShape(shape, Object.assign({ x, y, w, h, fill: { color } }, extra || {}));
}

/** Outlined (unfilled) shape. */
function outline(slide, shape, x, y, w, h, color, width, extra) {
  slide.addShape(
    shape,
    Object.assign({ x, y, w, h, fill: { type: 'none' }, line: { color, width: width || 1 } }, extra || {})
  );
}

/**
 * Closed custom shape. `pts` are fractions of w/h:
 *   [x, y]                     -> line to
 *   [x, y, c1x, c1y, c2x, c2y] -> cubic bezier to
 */
function freeform(slide, x, y, w, h, pts, extra) {
  const points = pts.map((p, i) => {
    const pt = { x: p[0] * w, y: p[1] * h };
    if (i === 0) return Object.assign(pt, { moveTo: true });
    if (p.length > 2) {
      pt.curve = { type: 'cubic', x1: p[2] * w, y1: p[3] * h, x2: p[4] * w, y2: p[5] * h };
    }
    return pt;
  });
  points.push({ close: true });
  slide.addShape(S.custGeom, Object.assign({ x, y, w, h, points }, extra || {}));
}

const K = 0.5523; // circular-arc bezier constant

/** Arch: semicircular top (radius = w/2), square bottom. */
function archPts(w, h) {
  const ry = (w / 2) / h; // the semicircle's radius, expressed in h units
  const kx = K * 0.5;
  const ky = K * ry;
  return [
    [0.5, 0],
    [1, ry, 0.5 + kx, 0, 1, ry - ky],
    [1, 1],
    [0, 1],
    [0, ry],
    [0.5, 0, 0, ry - ky, 0.5 - kx, 0],
  ];
}

/** Pill: semicircular top and bottom (radius = w/2), straight sides. */
function pillPts(w, h) {
  const ry = (w / 2) / h;
  const kx = K * 0.5;
  const ky = K * ry;
  return [
    [0.5, 0],
    [1, ry, 0.5 + kx, 0, 1, ry - ky],
    [1, 1 - ry],
    [0.5, 1, 1, 1 - ry + ky, 0.5 + kx, 1],
    [0, 1 - ry, 0.5 - kx, 1, 0, 1 - ry + ky],
    [0, ry],
    [0.5, 0, 0, ry - ky, 0.5 - kx, 0],
  ];
}

/**
 * Placeholder standing in for one of the deck's photographs.
 * kind: 'rect' | 'arch' | 'pill' | 'oval'
 */
function photo(slide, x, y, w, h, kind, label) {
  const fill = { color: C.photo };
  if (kind === 'arch') freeform(slide, x, y, w, h, archPts(w, h), { fill });
  else if (kind === 'pill') freeform(slide, x, y, w, h, pillPts(w, h), { fill });
  else if (kind === 'oval') box(slide, S.ellipse, x, y, w, h, C.photo);
  else box(slide, S.rect, x, y, w, h, C.photo);

  if (label !== false) {
    tbox(slide, '[image]', {
      x, y: y + h / 2 - 0.18, w, h: 0.36,
      align: 'center', valign: 'middle',
      fontFace: BODY, fontSize: 11, color: C.photoInk,
    });
  }
}

/** 40pt Abril Fatface headline built from coloured runs. */
function headline(slide, runs, x, y, w, h, align) {
  tbox(
    slide,
    runs.map((r) => ({ text: r[0], options: { color: r[1], breakLine: r[2] === 'br' } })),
    { x, y, w, h, fontFace: HEAD, fontSize: 40, bold: true, align: align || 'left', color: C.dark }
  );
}

/** 10.5pt Work Sans body copy, 1.5 line spacing. */
function copy(slide, text, x, y, w, h, opt) {
  const o = opt || {};
  tbox(slide, text, {
    x, y, w, h,
    fontFace: BODY, fontSize: 10.5,
    color: o.color || C.black,
    align: o.align || 'justify',
    lineSpacingMultiple: 1.5,
  });
}

/** 18pt bold Work Sans sub-heading (deck default size). */
function subhead(slide, text, x, y, w, h, color) {
  tbox(slide, text, {
    x, y, w, h, fontFace: BODY, fontSize: 18, bold: true, color: color || C.maroon, italic: false,
  });
}

/** 12pt bold Work Sans mini-heading. */
function minihead(slide, text, x, y, w, h, color) {
  tbox(slide, text, { x, y, w, h, fontFace: BODY, fontSize: 12, bold: true, color });
}

/** Centred white caption inside a coloured card. */
function cardText(slide, text, x, y, w, h, opt) {
  const o = opt || {};
  tbox(slide, text, {
    x, y, w, h, fontFace: BODY, fontSize: 10.5, color: o.color || C.white,
    align: o.align || 'center', lineSpacingMultiple: 1.5,
  });
}

/* --------------------------------------------------------- repeated furniture */

const ARROW_PTS = [
  [1, 0.5054], [0.6902, 0], [0.6902, 0.3763], [0, 0.3763],
  [0, 0.6237], [0.6902, 0.6237], [0.6902, 1],
];

function searchIcon(slide, x, y, s) {
  outline(slide, S.ellipse, x + 0.07 * s, y + 0.07 * s, 0.63 * s, 0.63 * s, C.maroon, 1.75);
  slide.addShape(S.line, {
    x: x + 0.66 * s, y: y + 0.76 * s, w: 0.3 * s, h: 0.22 * s,
    line: { color: C.maroon, width: 1.75 },
  });
}

function personIcon(slide, x, y, s) {
  box(slide, S.ellipse, x + 0.30 * s, y, 0.42 * s, 0.5 * s, C.maroon);
  box(slide, S.roundRect, x, y + 0.60 * s, s, 0.40 * s, C.maroon, { rectRadius: 0.2 * s });
}

/** Header / footer furniture that appears on every slide. */
function chrome(slide) {
  minihead(slide, 'MODELLING', 1.062, 0.345, 1.312, 0.303, C.maroon);

  [0.437, 0.4812, 0.5254].forEach((y, i) => {
    box(slide, S.rect, 6.459, y, 0.416, i === 2 ? 0.030 : 0.029, C.maroon);
  });

  searchIcon(slide, 11.442, 0.398, 0.197);
  personIcon(slide, 11.988, 0.398, 0.197);

  tbox(slide, 'MODEL PRESENTATION', {
    x: 1.062, y: 6.854, w: 2.188, h: 0.286, fontFace: BODY, fontSize: 11, color: C.maroon,
  });
  tbox(slide, 'NEXT SLIDE', {
    x: 10.493, y: 6.854, w: 1.146, h: 0.286, fontFace: BODY, fontSize: 11, color: C.maroon,
  });
  freeform(slide, 11.810, 6.938, 0.354, 0.118, ARROW_PTS, { fill: { color: C.maroon } });
}

/** The "Infographic Section" headline shared by slides 9-18. */
function infographicTitle(slide) {
  headline(slide, [['Infographic', C.maroon], [' Section', C.dark]], 3.844, 1.263, 5.582, 0.774, 'center');
}

/** Simple bank / museum glyph used on slides 10 and 15. */
function bankIcon(slide, x, y, w, h, color) {
  const fill = { color };
  freeform(slide, x, y, w, 0.36 * h, [[0, 1], [0.5, 0], [1, 1]], { fill });
  box(slide, S.rect, x, y + 0.30 * h, w, 0.07 * h, color);
  [0.10, 0.33, 0.56, 0.79].forEach((f) => box(slide, S.rect, x + f * w, y + 0.40 * h, 0.11 * w, 0.42 * h, color));
  box(slide, S.rect, x, y + 0.86 * h, w, 0.14 * h, color);
}

/* -------------------------------------------------------------------- slides */

function slide01() {
  const s = newSlide();
  photo(s, 4.847, 1.389, 3.639, 4.722, 'pill');
  chrome(s);
  outline(s, S.roundRect, 4.618, 1.151, 4.097, 5.198, C.maroon, 1, { rectRadius: 2.0485 });
  // Original strokes the middle "DELLI" instead of filling it; the visible
  // weight is identical, so a single solid run keeps the code simple.
  tbox(s, 'MODELLING', {
    x: 0.781, y: 2.732, w: 11.771, h: 2.036,
    fontFace: HEAD, fontSize: 115, bold: true, color: C.maroon, align: 'center',
  });
  tbox(s, 'Model Presentation Template', {
    x: 3.333, y: 4.965, w: 6.667, h: 0.505,
    fontFace: BODY, fontSize: 24, bold: true, italic: false, color: C.maroon, align: 'center',
  });
}

function slide02() {
  const s = newSlide();
  box(s, S.rect, 0, 0, 3.514, 7.5, C.peach);
  chrome(s);
  box(s, S.rect, 3.47, 0, 2.385, 7.5, C.maroon);

  headline(s, [['Welcome ', C.maroon, 'br'], ['To Modelling', C.dark]], 7.349, 1.45, 4.068, 1.447);
  subhead(s, 'About us', 7.349, 3.228, 1.782, 0.404);
  copy(s, LOREM, 7.349, 3.78, 4.835, 1.132);
  copy(
    s,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
      'incididunt ut labore et dolore magna aliqua. Ut laboris nisi ut aliquip ex ea commodo consequat. ',
    7.349, 5.077, 4.835, 0.867
  );

  photo(s, 1.167, 1.347, 3.708, 4.806, 'arch');
  freeform(s, 1.387, 1.347, 3.708, 4.806, archPts(3.708, 4.806), {
    fill: { type: 'none' }, line: { color: C.bg, width: 1 },
  });
}

function slide03() {
  const s = newSlide();
  outline(s, S.ellipse, 9.027, 1.322, 3.04, 4.875, C.peachDark, 1);
  chrome(s);
  box(s, S.rect, 0, 0, 0.389, 7.5, C.peach);

  headline(s, [['Vision', C.dark], [' & Mission', C.maroon]], 1.054, 1.485, 5.396, 0.774);
  copy(
    s,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
      'exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ',
    1.054, 2.354, 6.408, 0.601
  );

  const col =
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
    'incididunt ut labore et dolore magna aliqua. Ut enim, quis nostrud exercitation ' +
    'ullamco laboris nisi ut aliquip ex ea commodo consequat. ';
  [['Our Vision', 1.054], ['Our Mission', 4.789]].forEach(([title, x]) => {
    subhead(s, title, x, 3.461, 1.782, 0.404);
    copy(s, col, x, 4.013, 2.755, 1.927);
  });

  photo(s, 9.125, 1.322, 3.04, 4.875, 'oval');
}

function slide04() {
  const s = newSlide();
  outline(s, S.ellipse, 1.875, 1.473, 2.518, 4.875, C.peachDark, 1, { rotate: 116.3 });
  chrome(s);

  const facilities =
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
    'incididunt laboris nisi ut aliquip ex ea commodo consequat. ';
  [['01', 6.667, 7.582, 6.569], ['02', 9.833, 10.771, 9.757]].forEach(([num, bx, lx, tx]) => {
    tbox(s, num, {
      shape: S.roundRect, x: bx, y: 3.75, w: 0.787, h: 0.787, rectRadius: 0.3935,
      fill: { color: C.maroon }, align: 'center', valign: 'middle',
      fontFace: BODY, fontSize: 20, bold: true, color: C.white,
    });
    subhead(s, 'Facilities', lx, 3.942, 1.343, 0.404);
    copy(s, facilities, tx, 4.747, 2.549, 1.397);
  });

  headline(s, [['Modelling ', C.dark], ['Facilities', C.maroon]], 6.554, 1.485, 5.396, 0.774);
  copy(
    s,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
      'exercitation ullamco laboris commodo consequat. ',
    6.554, 2.354, 5.752, 0.601
  );

  // The single source photograph is a row of three overlapping pills.
  [[1.149, 1.375, 1.138, 3.429], [2.547, 2.033, 1.142, 3.433], [3.949, 2.682, 1.138, 3.429]]
    .forEach(([x, y, w, h], i) => photo(s, x, y, w, h, 'pill', i === 1));
}

function slide05() {
  const s = newSlide();
  chrome(s);
  const xs = [1.174, 3.996, 6.818, 9.640];
  xs.forEach((x) => {
    freeform(s, x, 2.467, 2.52, 3.673, archPts(2.52, 3.673), {
      fill: { type: 'none' }, line: { color: C.maroon, width: 1 },
    });
  });

  headline(
    s, [['The Projects ', C.maroon], ['Show This Year', C.dark]], 2.776, 1.393, 7.782, 0.774, 'center'
  );

  ['01. Project ', '02. Project ', '03. Project ', '04. Project '].forEach((t, i) => {
    tbox(s, t, {
      x: [1.564, 4.444, 7.266, 10.088][i], y: 5.536, w: 1.623, h: 0.4,
      fontFace: BODY, fontSize: 18, bold: true, color: C.maroon, align: 'center',
    });
  });

  [1.362, 4.184, 7.006, 9.828].forEach((x) => photo(s, x, 2.618, 2.143, 2.715, 'arch'));
}

function slide06() {
  const s = newSlide();
  [1.167, 4.138, 7.110, 10.081].forEach((x) => photo(s, x, 3.747, 2.083, 2.403, 'arch'));

  photo(s, 0, 0, 13.333, 2.973, 'rect', false);
  box(s, S.rect, 0, 0, 13.333, 2.973, C.dark, { fill: { color: C.dark, transparency: 30 } });

  chrome(s);

  ['ARIANI', 'LUISA', 'FERNAD', 'AMELIA'].forEach((name, i) => {
    tbox(s, name, {
      shape: S.rect, x: [1.167, 4.138, 7.110, 10.081][i], y: 5.731, w: 2.083, h: 0.419,
      fill: { color: C.maroon }, align: 'center', valign: 'middle',
      fontFace: BODY, fontSize: 16, bold: true, color: C.white, charSpacing: 2,
    });
  });

  box(s, S.roundRect, 10.568, 1.35, 2.207, 2.026, C.peach);
  tbox(s, '100%', {
    x: 11.080, y: 1.765, w: 1.229, h: 0.505,
    fontFace: BODY, fontSize: 24, bold: true, color: C.maroon, align: 'center',
  });
  tbox(s, 'Professional Model', {
    x: 10.786, y: 2.254, w: 1.818, h: 0.707, fontFace: BODY, fontSize: 18, color: C.black, align: 'center',
  });

  headline(s, [['Meet ', C.white], ['Our', C.peach], [' Models', C.white]], 1.026, 1.496, 5.012, 0.774);
}

function slide07() {
  const s = newSlide();
  photo(s, 2.408, 1.380, 2.669, 3.407, 'arch');
  freeform(s, 2.574, 1.364, 2.669, 3.407, archPts(2.669, 3.407), {
    fill: { type: 'none' }, line: { color: C.peach, width: 1 },
  });
  chrome(s);

  // photo-camera glyph
  box(s, S.roundRect, 6.667, 3.055, 0.369, 0.218, C.maroon, { rectRadius: 0.045 });
  box(s, S.rect, 6.760, 2.985, 0.120, 0.075, C.maroon);
  outline(s, S.ellipse, 6.795, 3.098, 0.115, 0.115, C.bg, 1.25);

  // video-camera glyph
  box(s, S.roundRect, 9.842, 2.985, 0.300, 0.260, C.maroon, { rectRadius: 0.05 });
  freeform(s, 10.130, 3.035, 0.130, 0.165, [[0, 0.5], [1, 0], [1, 1]], { fill: { color: C.maroon } });
  box(s, S.rect, 9.905, 3.075, 0.100, 0.080, C.bg);

  headline(s, [['Modelling ', C.dark], ['Services', C.maroon]], 6.563, 1.310, 5.396, 0.774);
  copy(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod.', 6.552, 2.113, 5.701, 0.336);

  const service =
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
    'incididunt ut labore et dolore magna exercitation ullamco laboris nisi commodo consequat. ';
  [['01', 6.563], ['02', 9.727]].forEach(([num, x]) => {
    tbox(s, 
      [
        { text: num, options: { fontSize: 28 } },
        { text: '. Service', options: { fontSize: 18 } },
      ],
      { x, y: 3.515, w: 1.782, h: 0.572, fontFace: BODY, bold: true, italic: false, color: C.maroon }
    );
    copy(s, service, x, 4.226, 2.525, 1.662);
  });

  photo(s, 1.144, 2.701, 2.669, 3.407, 'arch');
  freeform(s, 1.311, 2.701, 2.669, 3.407, archPts(2.669, 3.407), {
    fill: { type: 'none' }, line: { color: C.maroon, width: 1 },
  });
}

function slide08() {
  const s = newSlide();
  photo(s, 9.977, 1.402, 2.188, 2.358, 'arch');
  photo(s, 6.282, 1.402, 3.181, 4.743, 'arch');
  photo(s, 3.541, 3.787, 2.188, 2.358, 'arch');
  photo(s, 1.144, 3.787, 2.188, 2.358, 'arch');
  chrome(s);

  box(s, S.roundRect, 9.977, 4.118, 2.207, 2.026, C.peach);

  [['01', 1.845, 3.386], ['02', 4.241, 3.386], ['03', 7.479, 1.008], ['04', 10.687, 1.008]]
    .forEach(([num, x, y]) => {
      tbox(s, num, {
        shape: S.roundRect, x, y, w: 0.787, h: 0.787, rectRadius: 0.3935,
        fill: { color: C.maroon }, align: 'center', valign: 'middle',
        fontFace: BODY, fontSize: 18, bold: true, color: C.white,
      });
    });

  tbox(s, '187+', {
    x: 10.489, y: 4.435, w: 1.229, h: 0.505,
    fontFace: BODY, fontSize: 24, bold: true, color: C.maroon, align: 'center',
  });
  tbox(s, 'Modelling\nShow in 2024', {
    x: 10.195, y: 4.924, w: 1.818, h: 1.01, fontFace: BODY, fontSize: 18, color: C.black, align: 'center',
  });

  headline(s, [['Portfolio', C.maroon], [' Show', C.dark]], 1.048, 1.337, 4.404, 0.774);
  copy(
    s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt.',
    1.048, 2.111, 4.114, 0.601
  );
}

function slide09() {
  const s = newSlide();

  const cards = [
    [2.514, 2.962, C.peachDark], [3.914, 3.741, C.peach], [2.509, 4.515, C.maroon],
    [1.103, 3.739, C.peach], [3.903, 5.282, C.peachDark],
  ];
  cards.forEach(([x, y, color]) => box(s, S.round1Rect, x, y, 1.74, 1.5, color));

  [[1.228, 4.187], [2.640, 3.410], [4.040, 4.190], [4.029, 5.731], [2.634, 4.963]]
    .forEach(([x, y]) => cardText(s, 'Lorem ipsum dolor sit', x, y, 1.489, 0.603));

  const dots = [
    [7.426, 2.980, C.peach, 7.869, 2.849], [10.196, 2.986, C.peachDark, 10.639, 2.855],
    [7.426, 4.144, C.maroon, 7.869, 4.013], [10.196, 4.151, C.maroonDark, 10.639, 4.019],
  ];
  dots.forEach(([dx, dy, color, tx, ty]) => {
    box(s, S.ellipse, dx, dy, 0.29, 0.299, color);
    copy(s, DUIS, tx, ty, 1.595, 0.868, { color: C.dark });
  });

  [['75%', 5.476, 3.244, C.maroon], ['68%', 5.915, 2.518, C.peach]].forEach(([label, y, w, color]) => {
    tbox(s, label, {
      x: 11.74, y: y - 0.043, w: 0.552, h: 0.286, fontFace: BODY, fontSize: 11, color: C.dark, align: 'justify',
    });
    box(s, S.roundRect, 7.427, y, 4.17, 0.2, C.track, { rectRadius: 0.1 });
    box(s, S.roundRect, 7.426, y, w, 0.2, color, { rectRadius: 0.1 });
  });

  infographicTitle(s);
  chrome(s);
}

function slide10() {
  const s = newSlide();

  // Pyramid: two big faces plus the front trapezoid and their shadow triangles.
  box(s, S.rtTriangle, 7.881, 5.847, 0.489, 0.518, C.peachDeep, { flipV: true });
  box(s, S.rtTriangle, 5.541, 4.511, 0.518, 0.518, C.maroonDark, { flipV: true });
  box(s, S.triangle, 5.031, 3.047, 3.338, 2.854, C.peach, { flipH: true });
  box(s, S.triangle, 3.928, 4.506, 3.338, 2.854, C.maroon);
  box(s, S.flowChartManualOperation, 5.635, 5.893, 3.142, 1.611, C.peachDark, { flipV: true });

  box(s, S.ellipse, 5.664, 5.218, 1.27, 1.27, C.white);
  bankIcon(s, 6.007, 5.579, 0.583, 0.546, C.peach);

  cardText(s, 'Lorem ipsum dolor Duis aute irure dolor in lorem ipsum.', 5.875, 4.248, 1.711, 0.868);
  cardText(s, 'Lorem ipsum dolor Duis aute irure dolor in lorem ipsum.', 6.614, 6.454, 1.711, 0.874);
  cardText(s, 'Lorem ipsum dolor Duis aute.', 4.359, 6.370, 1.585, 0.603);

  const note = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit maecenas.';
  [[2.501, 3.271, 2.548, 3.551, C.peach], [1.036, 4.707, 1.083, 4.987, C.maroon],
   [7.938, 3.271, 7.985, 3.551, C.peach], [9.129, 4.707, 9.176, 4.987, C.maroon]]
    .forEach(([tx, ty, bx, by, color]) => {
      minihead(s, ' Title Here ', tx, ty, 1.679, 0.303, color);
      copy(s, note, bx, by, 3.097, 0.603, { color: C.dark });
    });

  infographicTitle(s);
  chrome(s);
}

function slide11() {
  const s = newSlide();

  // Four exploded doughnut wedges (paths taken from the source artwork).
  const wedges = [
    { x: 1.565, y: 2.981, w: 1.303, h: 1.212, color: C.peach, pts: [
      [0.9895, 0], [1, 0.8016], [0.9857, 0.8096], [0.978, 0.8088],
      [0.5964, 0.9787, 0.829, 0.8088, 0.6941, 0.8737], [0.5801, 1], [0, 0.5626],
      [0.9895, 0, 0.2267, 0.2152, 0.5948, 0.0059]] },
    { x: 1.233, y: 3.706, w: 1.048, h: 1.422, color: C.maroon, pts: [
      [0.287, 0], [1, 0.3769], [0.9769, 0.3975],
      [0.8622, 0.674, 0.9045, 0.4764, 0.8622, 0.5716],
      [0.8758, 0.7737, 0.8622, 0.7081, 0.8669, 0.7415], [0.8901, 0.8076], [0.0726, 1],
      [0.287, 0, -0.0747, 0.6605, 0.0048, 0.2896]] },
    { x: 1.331, y: 4.915, w: 1.462, h: 1.384, color: C.peach, pts: [
      [0.5832, 0], [0.5882, 0.017],
      [0.9345, 0.3171, 0.649, 0.169, 0.7779, 0.2832], [1, 0.3241], [0.9332, 1],
      [0, 0.1911, 0.497, 0.9519, 0.1299, 0.6337], [0.5832, 0]] },
    { x: 2.758, y: 5.070, w: 1.458, h: 1.245, color: C.maroon, pts: [
      [0.4488, 0], [1, 0.4658],
      [0, 0.9937, 0.7708, 0.8379, 0.3898, 1.039], [0.0641, 0.2383], [0.153, 0.2278],
      [0.3968, 0.0738, 0.2472, 0.2053, 0.3314, 0.1505], [0.4488, 0]] },
  ];
  wedges.forEach((wg) => freeform(s, wg.x, wg.y, wg.w, wg.h, wg.pts, { fill: { color: wg.color } }));

  [['77%', 1.817, 3.378], ['70%', 1.187, 4.238], ['78%', 2.893, 5.470], ['97%', 1.631, 5.305]]
    .forEach(([t, x, y]) => {
      tbox(s, t, {
        x, y, w: 1.043, h: 0.438, fontFace: BODY, fontSize: 20, bold: true, color: C.white, align: 'center',
      });
    });
  tbox(s, 'Title Here ', {
    x: 2.216, y: 4.346, w: 1.155, h: 0.707, fontFace: BODY, fontSize: 18, bold: true,
    color: C.peach, align: 'center',
  });

  const blurb = [
    { text: 'Lorem ipsum dolor sit amet, consectetuer adipiscing ', options: { color: C.black } },
    { text: '.aecenas porttitor.', options: { color: C.dark } },
  ];
  [['01', 5.044, 3.021, 5.110, 3.251, C.peach], ['02', 8.806, 3.017, 8.873, 3.247, C.maroon],
   ['03', 5.044, 4.211, 5.110, 4.441, C.maroon], ['04', 8.806, 4.207, 8.873, 4.436, C.peach]]
    .forEach(([num, tx, ty, bx, by, color]) => {
      minihead(s, ' Title Here ' + num, tx, ty, 1.597, 0.303, color);
      tbox(s, blurb, {
        x: bx, y: by, w: 3.338, h: 0.603, fontFace: BODY, fontSize: 10.5,
        align: 'justify', lineSpacingMultiple: 1.5,
      });
    });

  const bars = [
    [5.205, 5.456, 1.872, C.peach, '77%', 7.959], [5.205, 5.893, 1.717, C.maroon, '70%', 7.950],
    [8.964, 5.462, 2.473, C.peach, '97%', 11.735], [8.964, 5.899, 1.920, C.maroon, '78%', 11.735],
  ];
  bars.forEach(([x, y, w, color, label, lx]) => {
    box(s, S.roundRect, x, y, 2.726, 0.211, C.track, { rectRadius: 0.1055 });
    box(s, S.roundRect, x, y, w, 0.211, color, { rectRadius: 0.1055 });
    tbox(s, label, {
      x: lx, y: y - 0.04, w: 0.502, h: 0.278, fontFace: BODY, fontSize: 10.5, color: C.dark, align: 'right',
    });
  });

  infographicTitle(s);
  chrome(s);
}

function slide12() {
  const s = newSlide();

  // Slim corner brackets (the preset halfFrame is far chunkier than the original's 15% arms).
  const T = 0.151;
  const BRACKET = [[0, 1], [0, 0], [1, 0], [1 - T, T], [T, T], [T, 1 - T]];
  const tile = (x, y, color, flipH) => box(s, S.round2DiagRect, x, y, 1.444, 1.444, color, { flipH });

  // Interleaved exactly as in the original: the arms weave between the four tiles.
  tile(5.097, 2.962, C.maroon, true);
  tile(6.792, 4.684, C.maroon, false);
  freeform(s, 7.854, 3.419, 2.316, 2.316, BRACKET, { fill: { color: C.peachDark }, rotate: 315.3 });
  tile(6.782, 2.962, C.peach, false);
  freeform(s, 3.163, 3.419, 2.316, 2.316, BRACKET, {
    fill: { color: C.peachDark }, rotate: 44.7, flipH: true,
  });
  tile(5.097, 4.684, C.peach, true);

  [[5.097, 3.383], [6.782, 3.383], [5.107, 5.105], [6.792, 5.105]]
    .forEach(([x, y]) => cardText(s, 'Lorem ipsum dolor sit, ', x, y, 1.444, 0.603));

  const rows = [[3.013, C.peach], [4.329, C.maroon], [5.645, C.peach]];
  rows.forEach(([y, color]) => {
    box(s, S.flowChartDelay, 9.658, y, 0.347, 0.299, color);
    copy(s, DUIS, 10.101, y - 0.131, 2.188, 0.868, { color: C.dark });
    box(s, S.flowChartDelay, 3.346, y, 0.347, 0.299, color, { flipH: true });
    copy(s, DUIS, 1.062, y - 0.131, 2.188, 0.868, { color: C.dark, align: 'right' });
  });

  infographicTitle(s);
  chrome(s);
}

function slide13() {
  const s = newSlide();

  const TAIL = [[1, 0], [0, 0], [0, 0.4978], [0, 0.5022], [0, 1], [1, 1], [0.0045, 0.5]];
  const rows = [
    ['01', '02', 2.962, C.peach], ['03', '04', 4.150, C.maroon], ['05', '06', 5.337, C.peach],
  ];
  const note = 'Lorem ipsum dolor sit amet, consectetuer adipiscing.';

  rows.forEach(([left, right, y, color]) => {
    // left ribbon
    box(s, S.rect, 1.583, y, 3.539, 0.788, color);
    freeform(s, 1.158, y, 0.425, 0.788, TAIL, { fill: { color }, flipH: true });
    cardText(s, note, 1.836, y + 0.092, 2.904, 0.603, { align: 'justify' });
    tbox(s, left, {
      shape: S.snipRoundRect, x: 5.189, y: y + 0.003, w: 0.8, h: 0.8, flipH: true,
      fill: { color }, align: 'center', valign: 'middle',
      fontFace: BODY, fontSize: 18, bold: true, color: C.white,
    });
    // right ribbon
    box(s, S.rect, 8.222, y, 3.539, 0.788, color);
    freeform(s, 11.761, y, 0.425, 0.788, TAIL, { fill: { color } });
    cardText(s, note, 8.605, y + 0.092, 2.904, 0.603, { align: 'justify' });
    tbox(s, right, {
      shape: S.snipRoundRect, x: 7.355, y: y + 0.003, w: 0.8, h: 0.8,
      fill: { color }, align: 'center', valign: 'middle',
      fontFace: BODY, fontSize: 18, bold: true, color: C.white,
    });
    box(s, S.leftRightArrow, 6.460, y + 0.282, 0.425, 0.257, color);
  });

  infographicTitle(s);
  chrome(s);
}

function slide14() {
  const s = newSlide();

  const tape =
    [['A.', 2.974, 2.031, 3.148, 2.647, 3.166, C.maroon, false, 1.595],
     ['B.', 4.060, 2.042, 4.234, 2.658, 4.253, C.peach, true, 1.161],
     ['C.', 5.146, 2.020, 5.321, 2.636, 5.339, C.maroon, false, 1.584]];
  tape.forEach(([letter, y, lx, ly, tx, ty, color, flipH, x]) => {
    box(s, S.flowChartOnlineStorage, x, y, 4.761, 0.988, color, { flipH });
    tbox(s, letter, {
      x: lx, y: ly, w: 0.699, h: 0.64, fontFace: BODY, fontSize: 32, bold: true, color: C.white,
    });
    cardText(s, 'Lorem ipsum dolor sit amet, consectetuer dolor. ', tx, ty, 2.828, 0.603, { align: 'justify' });
  });

  [[3.013, C.maroon], [4.200, C.peach], [5.387, C.maroon]].forEach(([y, color]) => {
    [7.500, 10.159].forEach((x) => {
      box(s, S.notchedRightArrow, x, y, 0.347, 0.299, color);
      copy(s, DUIS, x + 0.443, y - 0.131, 1.595, 0.868, { color: C.dark });
    });
  });

  infographicTitle(s);
  chrome(s);
}

function slide15() {
  const s = newSlide();

  const short = 'Lorem ipsum dolor on sit amet, ipsum consectetuer. ';
  const tall =
    'Lorem ipsum dolor on sit amet, ipsum on sit consectetuer dolor. Lorem ipsum dolor on sit amet, ipsum.';

  // Five cards; "C." in the middle is taller and drawn first so its neighbours' badges overlap it.
  const cards = [
    { letter: 'C.', x: 5.433, y: 2.962, w: 2.076, h: 3.112, card: C.maroon,
      dotX: 7.116, dotY: 4.073, dotW: 0.787, dotH: 0.890, dot: C.maroonSoft,
      tx: 5.548, ty: 3.151, tw: 1.400, body: tall },
    { letter: 'A.', x: 1.167, y: 3.392, w: 1.797, h: 2.252, card: C.maroon,
      dotX: 2.623, dotY: 4.196, dotW: 0.681, dotH: 0.644, dot: C.maroonSoft,
      tx: 1.294, ty: 3.499, tw: 1.169, body: short },
    { letter: 'B.', x: 3.300, y: 3.392, w: 1.797, h: 2.252, card: C.peach,
      dotX: 4.756, dotY: 4.196, dotW: 0.681, dotH: 0.644, dot: C.peachSoft,
      tx: 3.414, ty: 3.499, tw: 1.169, body: short },
    { letter: 'D.', x: 7.898, y: 3.392, w: 1.797, h: 2.252, card: C.peach,
      dotX: 9.355, dotY: 4.196, dotW: 0.681, dotH: 0.644, dot: C.peachSoft,
      tx: 8.018, ty: 3.499, tw: 1.169, body: short },
    { letter: 'E.', x: 10.031, y: 3.392, w: 1.797, h: 2.252, card: C.maroon,
      dotX: 11.488, dotY: 4.196, dotW: 0.681, dotH: 0.644, dot: C.maroonSoft,
      tx: 10.157, ty: 3.499, tw: 1.169, body: short },
  ];

  cards.forEach((c) => {
    box(s, S.round1Rect, c.x, c.y, c.w, c.h, c.card);
    box(s, S.flowChartConnector, c.dotX, c.dotY, c.dotW, c.dotH, c.dot);
    tbox(s, c.letter, {
      x: c.tx, y: c.ty, w: 0.699, h: 0.64, fontFace: BODY, fontSize: 32, bold: true, color: C.white,
    });
    cardText(s, c.body, c.tx, c.ty + 0.639, c.tw, c.tw > 1.2 ? 1.928 : 1.398, { align: 'justify' });
  });

  [2.817, 4.952, 7.366, 9.548, 11.685].forEach((x) => bankIcon(s, x, 4.379, 0.285, 0.267, C.dark));

  infographicTitle(s);
  chrome(s);
}

function slide16() {
  const s = newSlide();

  box(s, S.diamond, 3.948, 2.962, 5.438, 3.207, C.peachPale);
  box(s, S.rtTriangle, 6.667, 3.063, 2.607, 1.506, C.maroon);
  box(s, S.rtTriangle, 4.060, 3.059, 2.607, 1.506, C.peach, { flipH: true });
  box(s, S.rtTriangle, 6.667, 4.565, 2.607, 1.506, C.peach, { flipV: true });
  box(s, S.rtTriangle, 4.060, 4.561, 2.607, 1.506, C.maroon, { flipH: true, flipV: true });
  box(s, S.flowChartConnector, 6.414, 4.315, 0.505, 0.505, C.peachPale);

  [[1.179, 2.962, C.peach], [1.179, 5.011, C.maroon], [11.514, 2.962, C.maroon], [11.514, 5.011, C.peach]]
    .forEach(([x, y, color]) => {
      tbox(s, '25%', {
        shape: S.roundRect, x, y, w: 0.658, h: 0.658, fill: { color },
        align: 'center', valign: 'middle', fontFace: BODY, fontSize: 12, bold: true, color: C.white,
      });
    });

  copy(s, MAECENAS, 1.974, 2.855, 2.158, 1.398, { color: C.dark });
  copy(s, MAECENAS, 1.974, 4.944, 2.158, 1.398, { color: C.dark });
  copy(s, MAECENAS, 9.202, 2.866, 2.175, 1.398, { color: C.dark, align: 'right' });
  copy(s, MAECENAS, 9.202, 4.955, 2.175, 1.398, { color: C.dark, align: 'right' });

  [[4.797, 4.135], [6.932, 4.135], [4.797, 4.652], [6.932, 4.652]].forEach(([x, y]) => {
    tbox(s, 'Title Here ', {
      x, y, w: 1.599, h: 0.303, fontFace: BODY, fontSize: 12, bold: true, color: C.white, align: 'center',
    });
  });

  infographicTitle(s);
  chrome(s);
}

function slide17() {
  const s = newSlide();

  const cards = [
    ['01', 1.161, 2.949, C.peach, C.peachDeep, C.peachDark],
    ['03', 1.161, 4.800, C.maroon, C.maroonDeep, C.maroonDark],
    ['02', 7.086, 2.933, C.maroon, C.maroonDeep, C.maroonDark],
    ['04', 7.086, 4.784, C.peach, C.peachDeep, C.peachDark],
  ];

  cards.forEach(([num, x, y, main, shade, badge]) => {
    box(s, S.round2SameRect, x + 0.736, y + 0.213, 4.357, 0.998, shade);
    box(s, S.flowChartDocument, x + 0.275, y + 0.151, 4.607, 1.123, main);
    box(s, S.flowChartDelay, x, y + 0.016, 1.118, 0.788, shade);
    tbox(s, num, {
      shape: S.flowChartDelay, x, y: y + 0.067, w: 1.058, h: 0.686, fill: { color: badge },
      align: 'center', valign: 'middle', fontFace: BODY, fontSize: 32, bold: true, color: C.white,
    });
    box(s, S.rtTriangle, x, y + 0.801, 0.275, 0.253, shade, { flipH: true, flipV: true });
    cardText(
      s, 'Lorem ipsum dolor on sit amet, ipsum consectetuer dolor. ',
      x + 1.218, y + 0.347, 3.305, 0.603, { align: 'justify' }
    );
  });

  infographicTitle(s);
  chrome(s);
}

function slide18() {
  const s = newSlide();

  const cols = [
    [1.167, 1.318, C.peachPale, C.peach, 1.569, 1.911],
    [5.176, 5.360, C.maroonPale, C.maroon, 5.581, 5.923],
    [9.185, 9.336, C.peachPale, C.peach, 9.590, 9.932],
  ];
  cols.forEach(([bx, fx, back, front, tx, hx]) => {
    box(s, S.flowChartMultidocument, bx, 2.981, 3.037, 3.144, back);
    box(s, S.flowChartMultidocument, fx, 3.137, 2.735, 2.832, front);
    tbox(s, 'Title Here ', {
      x: hx, y: 3.626, w: 1.544, h: 0.303, fontFace: BODY, fontSize: 12, bold: true,
      color: C.white, align: 'center',
    });
    cardText(s, MAECENAS, tx, 3.948, 2.227, 1.398);
  });

  box(s, S.rightArrow, 4.445, 4.378, 0.489, 0.35, C.peach);
  box(s, S.rightArrow, 8.455, 4.378, 0.489, 0.35, C.peach);

  infographicTitle(s);
  chrome(s);
}

function slide19() {
  const s = newSlide();
  chrome(s);

  // Tablet mock-up: rear tablet, the edge-on tablet leaning behind it,
  // the front tablet body and its grey screen.
  box(s, S.roundRect, 4.980, 1.510, 1.570, 4.400, 'D2D2D2', { rectRadius: 0.13 });
  box(s, S.roundRect, 1.404, 1.520, 0.190, 4.390, 'D5D7D6', { rectRadius: 0.09, rotate: 7.6 });
  box(s, S.roundRect, 1.790, 1.514, 2.920, 4.403, 'F5F6F8', { rectRadius: 0.13 });
  photo(s, 1.960, 1.939, 2.627, 3.548, 'rect');

  tbox(s, 'Contact us', {
    shape: S.roundRect, x: 5.071, y: 3.783, w: 2.207, h: 0.747, rectRadius: 0.3735,
    fill: { color: C.peach }, align: 'center', valign: 'middle',
    fontFace: BODY, fontSize: 18, color: C.maroon,
  });

  headline(s, [['Our Contact ', C.maroon, 'br'], ['Support', C.dark]], 8.130, 1.205, 3.803, 1.447);
  copy(
    s,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor exercitation.',
    8.130, 2.843, 4.054, 0.601
  );

  const rows = [
    ['+123 4567 8910', 3.995, 'phone'],
    ['yourmail@mail.com', 4.545, 'mail'],
    ['www.yourwebsite.com', 5.095, 'web'],
    ['123 Anywhere St., Any City', 5.645, 'home'],
  ];
  rows.forEach(([label, y, icon]) => {
    box(s, S.ellipse, 8.244, y, 0.462, 0.462, C.maroon);
    box(s, S.roundRect, 8.828, y, 2.334, 0.462, C.peach, { rectRadius: 0.231 });
    tbox(s, label, {
      x: 8.959, y: y + 0.088, w: 2.113, h: 0.286, fontFace: BODY, fontSize: 10.5, color: C.dark,
    });
    contactGlyph(s, icon, 8.355, y + 0.111, 0.24);
  });
}

/** Small pictograms inside the contact bullets on slide 19. */
function contactGlyph(s, kind, x, y, d) {
  const ink = C.bg;
  if (kind === 'phone') {
    freeform(s, x, y, d, d, [[0.05, 0.15], [0.35, 0.05], [0.5, 0.35], [0.32, 0.5],
      [0.5, 0.68], [0.65, 0.5], [0.95, 0.65], [0.85, 0.95], [0.05, 0.15, 0.35, 1.0, 0.0, 0.6]],
      { fill: { color: ink } });
  } else if (kind === 'mail') {
    box(s, S.rect, x, y + 0.15 * d, d, 0.7 * d, ink);
    freeform(s, x + 0.08 * d, y + 0.22 * d, 0.84 * d, 0.38 * d, [[0, 0], [1, 0], [0.5, 1]],
      { fill: { color: C.maroon } });
  } else if (kind === 'web') {
    outline(s, S.ellipse, x, y, d, d, ink, 1.25);
    outline(s, S.ellipse, x + 0.3 * d, y, 0.4 * d, d, ink, 1);
    s.addShape(S.line, { x, y: y + 0.5 * d, w: d, h: 0, line: { color: ink, width: 1 } });
  } else {
    freeform(s, x, y, d, 0.5 * d, [[0, 1], [0.5, 0], [1, 1]], { fill: { color: ink } });
    box(s, S.rect, x + 0.2 * d, y + 0.45 * d, 0.6 * d, 0.55 * d, ink);
  }
}

function slide20() {
  const s = newSlide();
  chrome(s);

  headline(s, [['Thanks', C.maroon, 'br'], [' For Your Attention', C.dark]], 1.062, 2.052, 3.347, 2.121);
  copy(
    s,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
      'exercitation commodo consequat. ',
    1.062, 4.279, 3.347, 0.867
  );
  copy(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing', 9.749, 5.249, 2.416, 0.601);

  photo(s, 9.184, 1.371, 3.0, 3.407, 'arch');
  photo(s, 5.902, 1.371, 3.0, 4.793, 'pill');
}

/* ---------------------------------------------------------------------- main */

pptx.defineLayout({ name: 'DECK', width: 13.3333, height: 7.5 });
pptx.layout = 'DECK';
pptx.author = 'MODELLING';
pptx.title = 'Model Presentation Template';

[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
 slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
  .forEach((build) => build());

pptx.writeFile({ fileName: path.join(__dirname, '0e17bc93-1630-4bbc-b9cc-3deff6f82235_grok_final.pptx') })
  .then((f) => console.log('wrote', f));
