/*
 * "FERTILIZER" deck — rebuilt with pptxgenjs.
 * Run: node 00f98224-a5fe-4adf-ab04-310a4097fafa_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const C = {
  green: '747823', // accent1
  rust: 'B75919', // accent2
  cream: 'E7CAB2', // lt1 / page background
  ink: '262626', // tx1 @ 85% lum
  gray: '404040', // tx1 @ 75% lum
  black: '000000',
  tan: 'E2BFA2',
  grid: 'C57F45',
  leaf: '3A3C12',
  mapFill: 'C1733B',
  screen: '404040',
  bezel: '1A1A1A',
  deck: 'E3E4E6',
  phoneEdge: '5E465D',
  phoneBody: '1D1D1B',
};

const HEAD = 'Geist Medium'; // theme major latin font
const BODY = 'Inter'; // theme minor latin font

const SLIDE_W = 13.3333; // 12192000 EMU
const SLIDE_H = 7.5; //  6858000 EMU

/* 8-pointed asterisk used as the deck's logo mark (outline, unit square) */
const STAR = [
  [1, 0.439], [1, 0.4437], [1, 0.5563], [1, 0.561], [0.6457, 0.561],
  [0.8948, 0.8114], [0.8123, 0.8943], [0.5584, 0.6393], [0.5584, 1],
  [0.4416, 1], [0.4416, 0.6393], [0.1877, 0.8943], [0.1052, 0.8114],
  [0.3544, 0.561], [0, 0.561], [0, 0.5563], [0, 0.4437], [0, 0.439],
  [0.3543, 0.439], [0.1052, 0.1886], [0.1877, 0.1057], [0.4416, 0.3607],
  [0.4416, 0], [0.5584, 0], [0.5584, 0.3607], [0.8123, 0.1057],
  [0.8948, 0.1886], [0.6456, 0.439],
];

/* thin arrow glyph, drawn pointing right then rotated 315 deg to get an up-right arrow */
const ARROW = [
  [1, 0.5001], [0.4138, 1], [0.4138, 0.8301], [0.7229, 0.5663],
  [0.002, 0.5663], [0, 0.4339], [0.7229, 0.4339], [0.4138, 0.1702],
  [0.4138, 0],
];

/* ---------------------------------------------------------------- helpers */

let pres; // set in build()

function freeform(slide, outline, o) {
  slide.addShape(pres.shapes.CUSTOM_GEOMETRY, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    rotate: o.rotate,
    fill: o.color ? { color: o.color } : { type: 'none' },
    line: o.line || { type: 'none' },
    points: outline.map(function (p) { return { x: p[0] * o.w, y: p[1] * o.h }; })
      .concat([{ close: true }]),
  });
}

/** logo asterisk; `s` is the square side in inches */
function star(slide, x, y, s, color) {
  freeform(slide, STAR, { x: x, y: y, w: s, h: s * 0.99532, color: color });
}

/** small up-right arrow; `w` is the un-rotated glyph width */
function arrow(slide, x, y, w, color, dir) {
  freeform(slide, ARROW, {
    x: x, y: y, w: w, h: w * 1.16819, color: color, rotate: dir === 'up-left' ? 225 : 315,
  });
}

function rect(slide, x, y, w, h, color, transparency) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: x, y: y, w: w, h: h,
    fill: transparency ? { color: color, transparency: transparency } : { color: color },
    line: { type: 'none' },
  });
}

function rule(slide, x, y, w, color, width) {
  slide.addShape(pres.shapes.LINE, {
    x: x, y: y, w: w, h: 0, line: { color: color, width: width || 1 },
  });
}

/** display text — theme major font */
function T(slide, text, o) {
  slide.addText(text, Object.assign({ valign: 'top', fontFace: HEAD, fontSize: 18 }, o));
}

/** paragraph text — theme minor font, 14pt, 130% leading */
function P(slide, text, o) {
  slide.addText(text, Object.assign(
    { valign: 'top', fontFace: BODY, fontSize: 14, lineSpacingMultiple: 1.3 }, o));
}

/** flat rust/green chip with a centred caption (used for tags & bar labels) */
function chip(slide, o) {
  rect(slide, o.x, o.y, o.w, o.h, o.fill);
  T(slide, o.text, {
    x: o.x, y: o.y + (o.h - 0.337) / 2, w: o.w, h: 0.337,
    fontSize: o.fontSize || 14, bold: true, color: o.color || C.cream,
    align: 'center', wrap: false,
  });
}

/** stand-in for a raster image from the source deck */
function imagePlaceholder(slide, o) {
  rect(slide, o.x, o.y, o.w, o.h, o.color);
  if (o.label) {
    T(slide, o.label, {
      x: o.x, y: o.y + o.h / 2 - 0.2, w: o.w, h: 0.4,
      fontSize: 12, color: o.labelColor || C.cream, align: 'center',
    });
  }
}

/* the source master paints the page number behind slide content, so several
   slides hide it under a full-bleed shape — hence it is added first, not via
   pptxgenjs `slideNumber` (which is always emitted last / on top) */
function pageNumber(slide, n) {
  T(slide, String(n), {
    x: 12.416, y: 0.167, w: 0.39, h: 0.286,
    fontSize: 10.5, color: C.ink, wrap: false,
  });
}

/* ------------------------------------------------------------ copy blocks */

const L = {
  welcome: 'Lorem ipsum dolor sit amet, consectetur adipis cing elit. Viva mus scelerisque tristi que. Aliquam et sem cursus massa ut.',
  farm: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Viva mus scelerisque tristique. Aliquam et sem aliquet, cursus massa ut.',
  long: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse aliquam, turpis dictum rutrum imperdiet, justo arcu consequat urna, ut consequat nisi massa et ante. ',
  stat: 'Lorem ipsum dolor sit amet, conse ctetur adipis cing elit. Viva mus scelerisque tristi que.',
  facility: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo',
  type: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Viva mus scelerisque dolor in eros.',
  short: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Viva mus scele tristique. Aliquam et sem aliquet.',
  research: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Viva mus scelerisque tristique. Aliquam et sem aliquet, cursus massa ut consectetur adipiscing.',
  product: 'Lorem ipsum dolor sit amet, consec tetur adipiscing elit. Viva mus scele risque tristique. Aliquam et sem.',
  brk: 'Lorem ipsum dolor sit amet, con sectetur adipiscing elit. Viva mus scelerisque tristique.',
  team: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Viva mus scelerisque dolor in amet, consectetur adipiscing eros.',
  tiny: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. ',
  price: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Viva mus sceleris eros.',
  priceBest: 'Lorem ipsum dolor sit amet, consectetur adipi scing elit. Viva mus sceleris eros. Viva mus scelerisque tristi que.',
  revenue: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Viva mus scele up lis tristique. Aliquam et sem consectetur adipiscing elit. Viva aliquet.',
};

/* --------------------------------------------------------- slide builders */

/* 1 — title */
function slide01(s) {
  s.background = { color: C.green };
  star(s, -0.728, 2.598, 3.354, C.cream);
  arrow(s, 12.051, 0.546, 0.665, C.rust);
  rect(s, 0, 5.063, SLIDE_W, 2.437, C.rust);
  T(s, 'FERTI', { x: 0.429, y: 0.221, w: 6.074, h: 2.423, fontSize: 138, color: C.cream });
  T(s, 'LIZER', { x: 3.179, y: 1.833, w: 6.074, h: 2.423, fontSize: 138, color: C.cream });
  rule(s, 0.656, 3.045, 2.333, C.rust, 1.5);
  T(s, 'SUPPORTING OUR VISION OF A WORLD WITHOUT HUNGER AND A PLANET RESPECTED, ' +
    'WE PURSUE A STRATEGY OF SUSTAINABLE FUTURE ',
  { x: 6.503, y: 5.77, w: 6.39, h: 1.01, color: C.cream });
  rule(s, 4.756, 5.968, 1.705, C.green, 1.5);
}

/* 2 — welcome, with three vertical rust columns on the right */
function slide02(s) {
  s.background = { color: C.green };
  star(s, 2.881, 0, 1.932, C.cream);
  T(s, 'WELCOME TO OUR FERTILIZER COMPANY ',
    { x: 3.806, y: 2.75, w: 6.111, h: 2.827, fontSize: 54, bold: true, color: C.cream });
  P(s, L.welcome, { x: 5.036, y: 5.92, w: 4.727, h: 0.99, color: C.cream });
  rule(s, 3.983, 6.125, 0.934, C.cream);

  const columns = [
    { x: 10.272, label: 'FERTILIZER PRODUCTION ' },
    { x: 11.296, label: 'PRODUCT QUALITY CONTROL' },
    { x: 12.32, label: 'FERTILIZATION EDUCATION CENTER ' },
  ];
  columns.forEach(function (col, i) {
    rect(s, col.x, 0, 1.013, SLIDE_H, C.rust);
    T(s, col.label, {
      x: 8.12 + i * 1.024, y: 4.049, w: 5.317, h: 0.404, rotate: 270, color: C.cream,
    });
    arrow(s, 10.61 + i * 1.024, 0.393, 0.338, C.cream);
  });
}

/* 3 — sustainable farm */
function slide03(s) {
  s.background = { color: C.rust };
  star(s, 11.407, 0.212, 1.565, C.cream);
  T(s, 'SUSTAINABLE FARM IN AGRICULTURE',
    { x: 4.702, y: 2.216, w: 7.761, h: 1.919, fontSize: 54, bold: true, color: C.cream });
  P(s, 'FERTILIZER SYSTEM',
    { x: 4.702, y: 0.625, w: 1.565, h: 0.269, fontSize: 10, color: C.cream, lineSpacingMultiple: 1 });
  T(s, '45,0%', { x: 6.494, y: 4.565, w: 2.367, h: 0.774, fontSize: 40, color: C.cream });
  T(s, 'MORE THAN 300+ GREAT PRODUCT WE HAVE',
    { x: 6.494, y: 5.43, w: 3.854, h: 0.707, color: C.cream });
  P(s, L.farm, { x: 6.494, y: 6.229, w: 6.705, h: 0.684, color: C.cream });
  rule(s, 4.945, 4.952, 1.367, C.cream);
}

/* 4 — from seed to harvest */
function slide04(s) {
  rect(s, 0, 0, 1.698, SLIDE_H, C.green);
  T(s, 'FROM SEED TO HARVEST ',
    { x: 1.968, y: 0.563, w: 7.558, h: 2.121, fontSize: 60, bold: true, color: C.ink });
  star(s, 11.635, 0.544, 1.402, C.green);

  rect(s, 7.733, 4.046, 2.339, 0.597, C.rust);
  T(s, 'GARDEN PLANT ',
    { x: 7.799, y: 4.143, w: 2.206, h: 0.404, bold: true, color: C.cream, align: 'center' });
  T(s, 'WAY TO PLANT IN THE GARDEN',
    { x: 7.656, y: 5.018, w: 4.476, h: 0.404, bold: true, color: C.ink });

  /* rust bullet dots are drawn as shapes: the source uses buClr=accent2,
     which pptxgenjs cannot express on a bulleted paragraph */
  const bullets = [
    'I throw myself down among the tall grass',
    'By the trickling I lie close to the earth',
    'A thousand unknown plants are noticed by me',
    'When I hear the buzz of the little world',
  ];
  bullets.forEach(function (line, i) {
    const y = 5.515 + i * 0.347;
    s.addShape(pres.shapes.OVAL, {
      x: 7.751, y: y + 0.098, w: 0.088, h: 0.088,
      fill: { color: C.rust }, line: { type: 'none' },
    });
    P(s, line, { x: 7.844, y: y, w: 4.823, h: 0.32, color: C.gray });
  });
}

/* 5 — predicting the future landscape */
function slide05(s) {
  s.background = { color: C.rust };
  T(s, 'PREDICTING THE FUTURE LANDSCAPE  ',
    { x: 0.589, y: 0.458, w: 8.374, h: 1.919, fontSize: 54, bold: true, color: C.cream });
  star(s, 11.638, 0.348, 1.287, C.cream);
  T(s, '3470+', { x: 0.589, y: 5.673, w: 2.367, h: 0.774, fontSize: 40, color: C.cream });
  T(s, 'GREEN PRODUCT ', { x: 0.589, y: 6.539, w: 2.633, h: 0.404, color: C.cream });
  P(s, L.long, { x: 6.667, y: 6.018, w: 6.512, h: 0.925, color: C.cream, lineSpacingMultiple: 1.2 });
  rule(s, 5.247, 6.219, 1.227, C.cream);
}

/* 6 — three research metrics */
function slide06(s) {
  s.background = { color: C.green };
  T(s, 'COMPANY RESEARCH METHOD SYSTEM  ',
    { x: 0.558, y: 0.561, w: 8.401, h: 1.919, fontSize: 54, bold: true, color: C.cream });
  star(s, 11.222, 0.625, 1.553, C.cream);

  const metrics = [
    { x: 0.667, value: '345+', label: 'FERTILIZER COMPOSITION PRODUCTION ' },
    { x: 4.942, value: '$378', label: 'PRODUCT REVENUE COMPLEXITY' },
    { x: 9.218, value: '50K', label: 'IMPACT IN FAMING AGRICULTURE ' },
  ];
  metrics.forEach(function (m) {
    T(s, m.value, { x: m.x, y: 3.889, w: 3.154, h: 1.01, fontSize: 54, color: C.cream });
    T(s, m.label, { x: m.x, y: 5.038, w: 3.558, h: 0.707, color: C.cream });
    P(s, L.stat, { x: m.x, y: 5.885, w: 3.433, h: 0.99, color: C.cream });
  });
}

/* 7 — boost crop fertilizer + facility list */
function slide07(s) {
  rect(s, 8.47, 0, 4.863, SLIDE_H, C.green);
  T(s, 'BOOST CROP FERTILIZER',
    { x: 0.523, y: 0.526, w: 5.005, h: 3.13, fontSize: 60, bold: true, color: C.rust });
  star(s, -0.034, 6.105, 1.402, C.green);
  T(s, '34,388+', { x: 0.606, y: 5.108, w: 3.56, h: 0.841, fontSize: 44, color: C.black });
  T(s, 'BALANCING FERTILIZER USE AND ENVIRONMENTAL IMPACT',
    { x: 0.621, y: 6.096, w: 4.329, h: 0.707, color: C.black });

  ['FACILITY ONE ', 'FACILITY TWO ', 'FACILITY THREE '].forEach(function (name, i) {
    const y = 0.904 + i * 2.1405;
    T(s, name, { x: 9.872, y: y, w: 2.161, h: 0.37, fontSize: 16, bold: true, color: C.cream });
    P(s, L.facility, { x: 9.872, y: y + 0.394, w: 2.905, h: 0.99, color: C.cream });
    rule(s, 9.317, y + 0.185, 0.439, C.cream);
  });
}

/* 8 — research development panel + three stat tiles */
function slide08(s) {
  s.background = { color: C.cream };
  rect(s, 0.667, 0.625, 7.065, 6.25, C.rust);
  star(s, 6.403, 0.884, 1.033, C.cream);
  T(s, 'RESEARCH DEVELOPMENT ',
    { x: 0.966, y: 3.045, w: 5.883, h: 1.919, fontSize: 54, bold: true, color: C.cream });
  P(s, L.long, { x: 2.086, y: 5.183, w: 4.969, h: 1.208, color: C.cream, lineSpacingMultiple: 1.2 });
  rule(s, 1.152, 5.385, 0.894, C.cream);

  const tiles = [
    { x: 8.091, y: 0.625, value: '3488+', label: 'MEDIA ENGAGEMENT ' },
    { x: 10.123, y: 2.81, value: '439K', label: 'PRODUCT PROMOTION' },
    { x: 8.089, y: 5.091, value: '56%', label: 'RESEARCH UPDATE ' },
  ];
  tiles.forEach(function (t) {
    rect(s, t.x, t.y, 2.493, 1.778, C.green);
    T(s, t.value, { x: t.x + 0.15, y: t.y + 0.122, w: 2.092, h: 0.774, fontSize: 40, color: C.cream });
    T(s, t.label, { x: t.x + 0.15, y: t.y + 0.949, w: 2.19, h: 0.707, color: C.cream });
  });
}

/* 9 — four fertilizer types */
function slide09(s) {
  s.background = { color: C.green };
  T(s, 'SUSTAINABILITY RESEARCH VALUE  ',
    { x: 0.558, y: 0.519, w: 8.2, h: 1.919, fontSize: 54, bold: true, color: C.cream });
  star(s, 11.417, 0.625, 1.359, C.cream);

  const cards = [
    { tag: 'FERTILIZER TYPE 01', bx: 0.667, by: 3.433, tx: 0.558, ty: 4.134 },
    { tag: 'FERTILIZER TYPE 02', bx: 6.654, by: 3.433, tx: 6.545, ty: 4.134 },
    { tag: 'FERTILIZER TYPE 03', bx: 0.651, by: 5.573, tx: 0.542, ty: 6.274 },
    { tag: 'FERTILIZER TYPE 04', bx: 6.654, by: 5.573, tx: 6.545, ty: 6.274 },
  ];
  cards.forEach(function (c) {
    rect(s, c.bx, c.by, 3.174, 0.529, C.rust);
    T(s, c.tag, { x: c.bx + 0.107, y: c.by + 0.063, w: 2.865, h: 0.404, bold: true, color: C.cream });
    P(s, L.type, { x: c.tx, y: c.ty, w: 4.972, h: 0.684, color: C.cream });
  });
}

/* 10 — best practice application */
function slide10(s) {
  rect(s, 0.667, 0.625, 12.0, 3.484, C.rust, 15);
  T(s, '$200,000', { x: 2.009, y: 1.811, w: 4.716, h: 1.111, fontSize: 60, color: C.cream });
  P(s, L.long, { x: 2.009, y: 3.107, w: 8.847, h: 0.643, color: C.cream, lineSpacingMultiple: 1.2 });
  rule(s, 1.009, 3.309, 0.945, C.cream);
  star(s, 11.331, 0.847, 0.963, C.green);
  T(s, 'BEST PRACTICE APPLICATION ',
    { x: 0.533, y: 5.198, w: 6.309, h: 1.919, fontSize: 54, bold: true, color: C.rust });
  T(s, 'INNOVATIVE FERTILIZER TECHNOLOGIES OF TODAY',
    { x: 7.457, y: 5.372, w: 4.109, h: 0.707, color: C.rust });
  P(s, L.short, { x: 7.457, y: 6.203, w: 5.343, h: 0.684, color: C.rust });
}

/* 11 — product research, area/line chart of monthly readings */
function slide11(s) {
  s.background = { color: C.rust };
  T(s, 'PRODUCT RESEARCH ',
    { x: 8.354, y: 3.746, w: 4.648, h: 1.919, fontSize: 54, bold: true, color: C.cream });
  star(s, 11.91, 0.345, 1.012, C.cream);
  P(s, L.product, { x: 9.165, y: 5.885, w: 3.628, h: 0.99, color: C.cream });
  rule(s, 8.551, 6.105, 0.566, C.cream);

  [{ n: '01.', x: 0.523, tx: 0.573 }, { n: '02.', x: 6.81, tx: 6.86 }].forEach(function (col) {
    T(s, col.n, { x: col.x, y: 0.519, w: 1.953, h: 1.313, fontSize: 72, bold: true, color: C.cream });
    P(s, L.research, { x: col.tx, y: 1.815, w: 5.95, h: 0.99, color: C.cream });
  });

  const series = [{
    name: 'Yield index',
    labels: ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11'],
    values: [22, 14, 35, 14, 29, 26, 8, 14, 60, 26, 26],
  }];
  s.addChart([
    { type: pres.charts.AREA, data: series, options: { chartColors: [C.green], chartColorsOpacity: 70 } },
    {
      type: pres.charts.LINE,
      data: series,
      options: {
        chartColors: [C.green], lineSize: 2.25, lineDataSymbol: 'circle',
        lineDataSymbolSize: 6, lineDataSymbolLineColor: C.leaf,
      },
    },
  ], {
    x: 0.65, y: 3.64, w: 7.36, h: 3.11,
    chartArea: { fill: { color: C.rust }, border: { color: C.rust, pt: 0 } },
    plotArea: { fill: { color: C.rust } },
    showLegend: false,
    catAxisHidden: true,
    catGridLine: { color: C.grid, size: 0.75, style: 'solid' },
    valGridLine: { color: C.grid, size: 0.75, style: 'solid' },
    valAxisLineShow: false, catAxisLineShow: false,
    valAxisMinVal: 0, valAxisMaxVal: 60, valAxisMajorUnit: 20,
    valAxisLabelColor: C.cream, valAxisLabelFontFace: HEAD, valAxisLabelFontSize: 10,
  });
}

/* 12 — gallery divider (picture frames of the source deck are empty) */
function slide12(s) {
  s.background = { color: C.green };
  T(s, 'FERTILIZER PORTFOLIO GALLERY IMAGE ',
    { x: 0.558, y: 0.519, w: 8.821, h: 1.919, fontSize: 54, bold: true, color: C.cream });
  star(s, 11.417, 0.625, 1.359, C.cream);
}

/* 13 — break slide */
function slide13(s) {
  s.background = { color: C.rust };
  T(s, 'BREAK', {
    x: 0.458, y: 0.447, w: 7.868, h: 1.847, fontSize: 115,
    color: C.cream, lineSpacingMultiple: 0.9,
  });
  T(s, 'SLIDE', {
    x: 2.428, y: 1.903, w: 5.658, h: 1.847, fontSize: 115,
    color: C.cream, lineSpacingMultiple: 0.9,
  });
  rule(s, 0.715, 2.692, 1.611, C.cream, 1.25);
  star(s, 11.519, -0.725, 2.712, C.cream);
  T(s, '60%', { x: 9.561, y: 4.667, w: 2.754, h: 1.111, fontSize: 60, color: C.cream });
  P(s, L.brk, { x: 9.561, y: 5.885, w: 3.35, h: 0.99, color: C.cream });
}

/* 14 — team */
function slide14(s) {
  T(s, 'OUR STRATEGIES TEAM ', {
    x: 5.135, y: 0.525, w: 7.682, h: 2.121, fontSize: 60, bold: true,
    color: C.ink, align: 'right',
  });
  star(s, 0.369, 6.11, 1.064, C.green);

  [{ name: 'YULIA LEE', x: 3.535, y: 2.344 }, { name: 'KALIA GILL', x: 0.664, y: 4.466 }]
    .forEach(function (person) {
      rect(s, person.x, person.y, 2.516, 0.689, C.green);
      T(s, person.name, {
        x: person.x + 0.073, y: person.y + 0.143, w: 1.621, h: 0.404, color: C.cream,
      });
      arrow(s, person.x + 2.09, person.y + 0.188, 0.268, C.rust);
    });

  [{ v: '15+', label: 'WORK EXPERIENCE ', x: 6.584, lw: 1.545, vw: 1.635 },
    { v: '560+', label: 'PRODUCT DEVELOPMENT ', x: 9.372, lw: 2.271, vw: 2.558 }]
    .forEach(function (m) {
      T(s, m.v, { x: m.x, y: 4.365, w: m.vw, h: 1.01, fontSize: 54, color: C.ink });
      T(s, m.label, { x: m.x, y: 5.356, w: m.lw, h: 0.572, fontSize: 14, color: C.ink });
    });
  P(s, L.team, { x: 6.667, y: 6.244, w: 6.298, h: 0.684, color: C.ink });
}

/* 15 — laptop mockup (raster image in the source) + headline */
function slide15(s) {
  s.background = { color: C.green };
  /* laptop mock-up: lid, screen, then the wedge of keyboard deck below */
  rect(s, -0.9, 1.15, 7.65, 5.05, C.bezel);
  imagePlaceholder(s, {
    x: -0.658, y: 1.472, w: 7.244, h: 4.5, color: C.screen, label: '[image]', labelColor: C.cream,
  });
  freeform(s, [[0, 0], [1, 0], [0.94, 1], [0.01, 1]], {
    x: -1.0, y: 6.2, w: 8.75, h: 0.46, color: C.deck,
  });
  T(s, 'THE ROLE OF FERTILIZERS IN COMBATING FOOD SECURITY CHALLENGES', {
    x: 7.331, y: 1.881, w: 5.789, h: 3.738, fontSize: 48,
    color: C.cream, lineSpacingMultiple: 0.9,
  });
  ['CORN FIELD ', 'RICE FIELD ', 'FRUIT GARDEN '].forEach(function (tag, i) {
    chip(s, {
      x: 7.463 + i * 1.84, y: 5.792, w: 1.695, h: 0.434,
      fill: C.rust, text: tag, color: C.tan,
    });
  });
  star(s, 10.999, 0.214, 1.356, C.cream);
}

/* 16 — two phone mockups with figures either side */
function slide16(s) {
  s.background = { color: C.rust };
  [{ x: 3.941, y: 0.7 }, { x: 6.833, y: 1.529 }].forEach(function (p) {
    slidePhone(s, p.x, p.y);
  });
  star(s, 11.519, 0.634, 1.147, C.cream);
  star(s, 0.666, 5.733, 1.147, C.cream);

  T(s, '75%', { x: 0.843, y: 1.157, w: 2.754, h: 1.111, fontSize: 60, color: C.cream, align: 'right' });
  T(s, 'FERTILIZER TRENDS AND FORECAST',
    { x: 0.978, y: 2.268, w: 2.619, h: 0.707, color: C.cream, align: 'right' });
  P(s, L.tiny, { x: 0.645, y: 3.025, w: 2.953, h: 0.684, color: C.cream, align: 'right' });
  arrow(s, 2.946, 0.624, 0.525, C.cream, 'up-left');

  T(s, '60%', { x: 9.617, y: 4.226, w: 2.754, h: 1.111, fontSize: 60, color: C.cream });
  T(s, 'FERTILIZER LABELS AND GRADES', { x: 9.617, y: 5.337, w: 2.619, h: 0.707, color: C.cream });
  P(s, L.tiny, { x: 9.617, y: 6.094, w: 2.953, h: 0.684, color: C.cream });
  arrow(s, 9.713, 3.693, 0.525, C.cream);
}

/* phone mock-up: coloured edge, body, then the pill-shaped camera notch */
function slidePhone(s, x, y) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: x, y: y, w: 2.559, h: 5.249, rectRadius: 0.32,
    fill: { color: C.phoneEdge }, line: { type: 'none' },
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: x + 0.013, y: y + 0.012, w: 2.534, h: 5.224, rectRadius: 0.31,
    fill: { color: '706F6F' }, line: { type: 'none' },
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: x + 0.044, y: y + 0.04, w: 2.472, h: 5.169, rectRadius: 0.3,
    fill: { color: C.phoneBody }, line: { type: 'none' },
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: x + 0.981, y: y + 0.2, w: 0.575, h: 0.167, rectRadius: 0.5,
    fill: { color: '111111' }, line: { type: 'none' },
  });
  s.addShape(pres.shapes.OVAL, {
    x: x + 1.413, y: y + 0.247, w: 0.073, h: 0.073,
    fill: { color: '666666' }, line: { type: 'none' },
  });
}

/* 17 — pricing cards */
function slide17(s) {
  T(s, 'FERTILIZER PRICING PACK',
    { x: 0.558, y: 0.519, w: 7.359, h: 1.919, fontSize: 54, bold: true, color: C.ink });
  star(s, 11.438, 0.625, 1.381, C.green);

  const cards = [
    { x: 0.714, w: 3.231, fill: C.green, price: '$63', name: 'PRODUCT 01 ', body: L.price,
      px: 0.856, py: 4.184, nx: 0.838, ny: 5.266, bx: 0.838, by: 5.755, bw: 2.984, nw: 2.444,
      ax: 3.464, ay: 3.379, ac: C.rust },
    { x: 4.271, w: 3.231, fill: C.green, price: '$75', name: 'PRODUCT 02 ', body: L.price,
      px: 4.413, py: 4.188, nx: 4.394, ny: 5.271, bx: 4.394, by: 5.76, bw: 2.984, nw: 2.444,
      ax: 7.02, ay: 3.383, ac: C.rust },
    { x: 7.829, w: 4.809, fill: C.rust, price: '$150', name: 'BEST PRODUCT 02 ', body: L.priceBest,
      px: 8.017, py: 4.23, nx: 7.998, ny: 5.312, bx: 7.998, by: 5.802, bw: 4.614, nw: 3.957,
      ax: 12.152, ay: 3.425, ac: C.cream },
  ];
  cards.forEach(function (c) {
    rect(s, c.x, 3.377, c.w, 3.498, c.fill);
    T(s, c.price, { x: c.px, y: c.py, w: 2.444, h: 1.01, fontSize: 54, color: C.cream });
    T(s, c.name, { x: c.nx, y: c.ny, w: c.nw, h: 0.438, fontSize: 20, color: C.cream });
    P(s, c.body, { x: c.bx, y: c.by, w: c.bw, h: 0.99, color: C.cream });
    arrow(s, c.ax, c.ay, 0.459, c.ac);
  });
}

/* 18 — country bar chart drawn as flat rectangles */
function slide18(s) {
  T(s, 'TOP COUNTRIES WITH BEST SELLING PRODUCT ',
    { x: 0.558, y: 0.519, w: 9.442, h: 1.919, fontSize: 54, bold: true, color: C.ink });
  star(s, 11.375, 0.625, 1.514, C.green);

  const bars = [
    { x: 0.667, top: 5.581, fill: C.green, name: 'MYANMAR', pct: '45%', px: 1.003 },
    { x: 3.77, top: 4.806, fill: C.rust, name: 'CAMBODIA ', pct: '56%', px: 4.107 },
    { x: 6.874, top: 4.194, fill: C.green, name: 'THAILAND ', pct: '75%', px: 7.21 },
    { x: 9.977, top: 3.25, fill: C.rust, name: 'INDONESIA ', pct: '85%', px: 10.314 },
  ];
  bars.forEach(function (b) {
    rect(s, b.x, b.top, 2.765, SLIDE_H - b.top, b.fill);
    chip(s, { x: b.x, y: b.top - 0.64, w: 2.765, h: 0.434, fill: b.fill, text: b.name });
    T(s, b.pct, {
      x: b.px, y: 6.632, w: 2.092, h: 0.774, fontSize: 40, color: C.cream, align: 'center',
    });
  });
  P(s, L.short, { x: 1.531, y: 2.694, w: 5.343, h: 0.684, color: C.rust });
  rule(s, 0.667, 2.903, 0.773, C.ink, 1.25);
}

/* 19 — revenue progress bar */
function slide19(s) {
  s.background = { color: C.green };
  T(s, 'TARGET FERTILIZER MONTHLY REVENUE',
    { x: 0.667, y: 0.519, w: 9.442, h: 1.919, fontSize: 54, bold: true, color: C.cream });
  star(s, 11.59, 0.625, 0.921, C.cream);

  T(s, 'JAN-FEB REVENUE UPDATE',
    { x: 0.758, y: 3.152, w: 4.408, h: 0.438, fontSize: 20, color: C.cream });
  rect(s, 0.823, 3.728, 11.688, 1.619, C.cream, 18);
  rect(s, 0.811, 3.728, 4.781, 1.619, C.rust);
  rect(s, 5.593, 3.474, 0.25, 0.25, C.rust);
  T(s, '$35,000', { x: 1.014, y: 4.081, w: 4.039, h: 0.909, fontSize: 48, color: C.cream });
  T(s, '48%', {
    x: 9.521, y: 4.113, w: 2.616, h: 0.909, fontSize: 48,
    fontFace: BODY, color: C.ink, align: 'right',
  });

  T(s, '$200,000', { x: 1.014, y: 5.775, w: 3.108, h: 0.707, fontSize: 36, color: C.cream });
  T(s, 'TOTAL TARGET REVENUE', { x: 1.014, y: 6.538, w: 2.797, h: 0.337, fontSize: 14, color: C.cream });
  P(s, L.revenue, { x: 5.312, y: 6.129, w: 7.355, h: 0.684, color: C.cream });
  rule(s, 4.122, 6.337, 1.064, C.ink, 1.25);
}

/* 20 — archipelago map */
/* 60 islands traced from the source map artwork (664 points) */
const ISLANDS = [
  [[12.09,5.16],[12.09,3.78],[11.35,3.51],[11.19,3.59],[11.20,3.66],[11.01,3.69],[10.80,3.96],[10.55,3.87],[10.50,3.96],[10.74,4.05],[10.67,4.18],[11.22,4.33],[11.47,4.50],[11.50,4.67],[11.67,4.80],[11.52,4.80],[11.68,4.85],[11.55,4.86],[11.65,4.95],[11.60,5.12],[11.89,5.03],[11.85,5.09],[12.09,5.31],[12.09,5.16]],
  [[6.15,3.42],[6.07,3.37],[6.04,3.16],[5.95,3.21],[5.95,2.99],[5.67,3.03],[5.57,3.11],[5.55,3.29],[5.28,3.34],[4.99,3.54],[5.05,3.81],[4.96,3.88],[5.15,3.87],[5.19,3.81],[5.22,4.01],[5.40,3.97],[5.50,3.85],[5.66,3.99],[5.75,3.92],[5.80,3.97],[6.15,3.42]],
  [[5.77,2.86],[5.36,2.80],[5.21,2.93],[4.90,2.96],[4.70,2.80],[4.69,2.67],[4.62,2.71],[4.52,3.08],[4.61,3.16],[4.57,3.28],[4.66,3.33],[4.61,3.37],[4.72,3.37],[4.79,3.48],[4.84,3.87],[5.05,3.81],[5.00,3.53],[5.28,3.34],[5.55,3.29],[5.57,3.11],[5.77,2.86]],
  [[6.91,2.93],[6.60,2.69],[6.67,2.59],[6.45,2.56],[6.31,2.60],[6.03,2.93],[5.77,2.83],[5.67,3.03],[5.97,3.01],[5.95,3.21],[6.04,3.16],[6.15,3.42],[6.16,3.71],[6.33,3.72],[6.25,3.59],[6.37,3.41],[6.42,3.46],[6.58,3.35],[6.54,3.14],[6.68,2.98],[6.64,2.91],[6.91,2.93]],
  [[10.50,3.96],[10.60,3.83],[10.47,3.35],[10.06,3.25],[9.79,3.36],[9.72,3.50],[9.97,3.53],[10.04,3.70],[10.43,3.66],[10.37,3.81],[9.97,3.82],[10.19,4.13],[10.42,3.85],[10.37,4.03],[10.65,4.17],[10.74,4.05],[10.50,3.96]],
  [[3.85,3.92],[3.74,3.73],[3.56,3.71],[3.49,3.82],[3.57,3.66],[3.46,3.57],[3.25,3.63],[3.24,3.73],[3.18,3.68],[2.91,3.82],[3.02,3.95],[3.13,3.97],[3.04,4.07],[3.27,4.20],[3.36,4.33],[3.67,4.05],[3.79,4.14],[3.85,3.92]],
  [[3.32,3.16],[3.21,3.04],[3.02,3.13],[3.12,3.01],[2.67,2.63],[2.64,2.74],[2.49,2.56],[2.46,3.03],[2.61,3.11],[2.62,3.24],[2.78,3.38],[3.02,3.42],[3.23,3.34],[3.27,3.27],[3.19,3.23],[3.32,3.16]],
  [[6.67,2.59],[6.49,2.41],[6.54,2.35],[6.44,2.32],[6.63,2.29],[6.56,2.18],[6.16,2.13],[6.07,2.45],[5.98,2.48],[6.01,2.58],[5.91,2.64],[5.85,2.83],[6.03,2.93],[6.20,2.81],[6.32,2.60],[6.67,2.59]],
  [[2.52,2.78],[2.48,2.54],[2.42,2.55],[2.36,2.42],[1.99,2.20],[2.01,2.15],[1.96,2.17],[1.90,2.29],[1.97,2.65],[2.14,2.77],[2.22,3.11],[2.31,3.03],[2.40,3.05],[2.35,2.96],[2.46,2.99],[2.44,2.84],[2.52,2.78]],
  [[7.81,3.38],[7.94,3.40],[7.87,3.29],[7.52,3.35],[7.29,3.48],[7.15,3.31],[7.17,3.11],[7.66,2.93],[7.35,2.85],[7.05,3.17],[7.11,3.36],[7.02,3.48],[7.72,3.93],[7.45,3.59],[7.81,3.38]],
  [[2.01,2.13],[1.83,1.93],[1.57,1.94],[1.38,1.84],[1.29,1.86],[1.37,2.08],[1.59,2.28],[1.69,2.30],[1.85,2.49],[1.89,2.63],[1.97,2.65],[1.92,2.40],[1.94,2.38],[1.90,2.28],[2.01,2.13]],
  [[3.49,3.58],[3.45,3.40],[3.32,3.42],[3.24,3.34],[3.11,3.34],[3.02,3.42],[2.87,3.38],[2.81,3.52],[2.68,3.59],[2.73,3.71],[2.91,3.82],[3.03,3.81],[3.18,3.68],[3.24,3.73],[3.25,3.63],[3.34,3.59],[3.49,3.58]],
  [[7.56,3.82],[7.24,3.60],[7.05,3.68],[7.09,3.80],[7.02,3.82],[7.05,3.90],[7.00,3.92],[7.05,4.10],[7.01,4.48],[7.25,4.49],[7.20,4.38],[7.24,3.93],[7.18,3.86],[7.32,3.78],[7.49,3.88],[7.56,3.82]],
  [[5.87,5.21],[5.81,5.16],[5.82,4.99],[5.53,4.99],[5.43,4.94],[5.39,4.78],[5.17,4.75],[5.16,4.86],[5.06,4.88],[5.08,5.04],[5.00,5.10],[5.41,5.16],[5.56,5.12],[5.87,5.21]],
  [[4.52,4.77],[4.45,4.76],[4.39,4.64],[4.23,4.64],[4.16,4.57],[4.08,4.56],[4.06,4.67],[3.94,4.66],[3.96,4.76],[3.93,4.82],[3.96,4.83],[3.92,4.89],[4.41,5.01],[4.50,4.97],[4.44,4.88],[4.52,4.77]],
  [[6.33,3.72],[6.16,3.72],[6.10,3.52],[6.12,3.48],[6.06,3.50],[6.04,3.64],[5.93,3.70],[5.80,3.97],[5.84,3.99],[5.85,3.95],[5.87,4.15],[6.19,4.02],[6.26,3.91],[6.23,3.83],[6.26,3.87],[6.27,3.76],[6.32,3.77],[6.33,3.72]],
  [[2.86,3.44],[2.62,3.24],[2.59,3.18],[2.62,3.12],[2.48,3.06],[2.46,2.98],[2.35,2.96],[2.40,3.05],[2.31,3.03],[2.22,3.11],[2.35,3.16],[2.50,3.37],[2.66,3.75],[2.73,3.71],[2.68,3.57],[2.77,3.57],[2.86,3.44]],
  [[5.17,4.75],[5.05,4.74],[5.02,4.68],[4.97,4.68],[4.89,4.80],[4.52,4.77],[4.44,4.89],[4.52,4.98],[4.79,5.03],[4.88,4.95],[5.00,5.11],[5.08,5.04],[5.06,4.88],[5.16,4.86],[5.17,4.75]],
  [[3.81,4.32],[3.79,4.14],[3.67,4.05],[3.60,4.15],[3.43,4.22],[3.43,4.32],[3.36,4.33],[3.31,4.30],[3.27,4.32],[3.50,4.56],[3.53,4.56],[3.50,4.46],[3.64,4.53],[3.66,4.45],[3.77,4.56],[3.81,4.32]],
  [[7.70,4.22],[7.83,4.20],[7.82,4.13],[7.77,4.14],[7.77,4.08],[7.66,4.01],[7.70,3.93],[7.40,3.84],[7.34,3.99],[7.52,4.12],[7.51,4.28],[7.63,4.31],[7.63,4.24],[7.70,4.22]],
  [[3.31,4.30],[3.27,4.20],[3.04,4.07],[3.12,3.96],[3.02,3.95],[2.75,3.69],[2.66,3.75],[2.98,4.12],[3.22,4.31],[3.31,4.30]],
  [[9.23,3.10],[9.08,2.98],[9.20,2.80],[8.94,2.94],[9.01,2.65],[8.88,2.88],[8.95,3.22],[9.13,3.38],[9.00,3.06],[9.23,3.10]],
  [[9.69,4.07],[9.64,3.90],[9.39,3.82],[8.99,3.91],[9.01,4.00],[9.07,3.89],[9.14,3.98],[9.47,3.95],[9.69,4.07]],
  [[7.11,3.63],[7.02,3.37],[6.99,3.63],[6.84,3.82],[6.89,4.01],[7.02,3.99],[7.11,3.63]],
  [[6.94,5.19],[6.64,5.07],[6.73,5.21],[6.48,5.14],[6.37,5.28],[6.76,5.19],[6.89,5.26],[6.83,5.23],[6.94,5.19]],
  [[7.81,5.18],[7.86,5.12],[7.82,5.07],[7.72,5.20],[7.24,5.11],[7.11,5.15],[7.09,5.23],[7.39,5.28],[7.81,5.18]],
  [[8.36,5.30],[8.27,5.29],[8.16,5.41],[8.08,5.38],[7.96,5.61],[8.19,5.57],[8.36,5.30]],
  [[8.37,2.82],[8.32,2.75],[8.22,2.88],[8.07,2.97],[7.87,2.95],[7.96,3.09],[8.13,3.08],[8.20,3.05],[8.37,2.82]],
  [[7.97,3.05],[7.85,2.94],[7.81,2.97],[7.66,2.92],[7.41,3.01],[7.45,3.06],[7.87,3.04],[7.92,3.10],[7.97,3.05]],
  [[11.54,5.11],[11.60,5.07],[11.64,4.95],[11.58,4.90],[11.48,4.91],[11.39,4.97],[11.30,5.16],[11.48,5.15],[11.54,5.11]],
  [[4.01,3.89],[4.03,3.77],[3.91,3.74],[3.85,3.54],[3.69,3.54],[3.63,3.65],[3.79,3.68],[3.83,3.83],[4.01,3.89]],
  [[7.33,5.52],[7.13,5.36],[6.94,5.38],[6.89,5.42],[7.05,5.47],[7.24,5.60],[7.33,5.52]],
  [[4.01,4.66],[4.00,4.58],[3.84,4.55],[3.79,4.69],[3.72,4.77],[3.69,4.74],[3.65,4.77],[3.93,4.82],[3.94,4.66],[4.01,4.66]],
  [[8.85,3.96],[8.73,3.89],[8.57,3.90],[8.55,3.93],[8.60,4.02],[8.72,4.07],[8.84,4.03],[8.85,3.96]],
  [[6.12,5.14],[6.02,5.07],[5.93,5.10],[5.83,5.08],[5.86,5.14],[6.00,5.21],[5.99,5.26],[6.12,5.14]],
  [[6.37,5.15],[6.26,5.11],[6.20,5.16],[6.21,5.24],[6.15,5.24],[6.33,5.27],[6.32,5.24],[6.37,5.15]],
  [[5.42,4.83],[5.42,4.85],[5.62,4.87],[5.63,4.85],[5.76,4.81],[5.72,4.78],[5.47,4.79],[5.42,4.83]],
  [[4.38,3.83],[4.32,3.78],[4.23,3.77],[4.21,3.93],[4.27,3.89],[4.32,3.93],[4.38,3.83]],
  [[7.88,4.28],[7.87,4.20],[7.83,4.22],[7.81,4.39],[7.75,4.46],[7.77,4.51],[7.90,4.42],[7.84,4.39],[7.88,4.28]],
  [[1.93,2.93],[1.83,2.82],[1.73,2.83],[1.87,3.03],[1.92,3.02],[1.93,2.93]],
  [[10.62,4.63],[10.63,4.50],[10.58,4.44],[10.49,4.51],[10.54,4.53],[10.51,4.61],[10.59,4.67],[10.62,4.63]],
  [[2.25,3.55],[2.17,3.39],[2.10,3.39],[2.08,3.45],[2.14,3.55],[2.20,3.59],[2.24,3.59],[2.25,3.55]],
  [[9.91,4.85],[9.87,4.85],[9.79,4.93],[9.80,4.98],[9.77,4.98],[9.76,5.05],[9.81,5.06],[9.88,4.98],[9.91,4.85]],
  [[9.56,3.22],[9.63,3.27],[9.72,3.25],[9.64,3.18],[9.74,3.25],[9.82,3.24],[9.70,3.16],[9.56,3.22]],
  [[4.98,5.10],[4.96,5.01],[4.90,5.01],[4.88,4.95],[4.81,4.98],[4.79,5.03],[4.98,5.10]],
  [[8.39,3.59],[8.21,3.55],[8.16,3.61],[8.22,3.64],[8.40,3.61],[8.39,3.59]],
  [[7.75,4.44],[7.75,4.38],[7.79,4.33],[7.78,4.25],[7.69,4.31],[7.71,4.36],[7.68,4.43],[7.74,4.41],[7.75,4.44]],
  [[8.75,4.98],[8.71,4.95],[8.63,4.98],[8.54,4.97],[8.50,5.06],[8.55,5.03],[8.66,5.05],[8.75,4.98]],
  [[10.78,3.34],[10.83,3.37],[10.86,3.36],[10.89,3.45],[11.01,3.43],[10.88,3.33],[10.77,3.31],[10.78,3.34]],
  [[9.06,3.54],[8.94,3.48],[8.89,3.49],[8.88,3.55],[8.92,3.58],[9.05,3.57],[9.06,3.54]],
  [[11.00,3.57],[10.78,3.54],[10.98,3.62],[11.13,3.59],[11.00,3.57]],
  [[10.56,4.74],[10.55,4.68],[10.47,4.62],[10.49,4.70],[10.47,4.69],[10.46,4.77],[10.49,4.79],[10.56,4.74]],
  [[9.08,2.64],[9.09,2.70],[9.16,2.67],[9.19,2.59],[9.17,2.55],[9.10,2.59],[9.08,2.64]],
  [[7.97,3.50],[7.96,3.45],[7.90,3.49],[7.89,3.43],[7.81,3.47],[7.82,3.54],[7.88,3.47],[7.89,3.55],[7.92,3.50],[7.97,3.50]],
  [[8.17,5.16],[8.34,5.14],[8.36,5.11],[8.20,5.08],[8.18,5.11],[8.21,5.11],[8.17,5.16]],
  [[6.27,4.04],[6.26,3.92],[6.21,3.99],[6.21,4.11],[6.27,4.08],[6.27,4.04]],
  [[9.60,3.63],[9.58,3.56],[9.43,3.61],[9.53,3.65],[9.60,3.63]],
  [[8.94,3.30],[8.93,3.24],[8.90,3.26],[8.87,3.25],[8.90,3.36],[8.94,3.34],[8.99,3.37],[8.99,3.33],[8.95,3.33],[8.94,3.30]],
  [[1.59,2.60],[1.46,2.48],[1.40,2.50],[1.42,2.54],[1.49,2.55],[1.55,2.61],[1.59,2.60]],
  [[4.40,2.25],[4.36,2.17],[4.30,2.22],[4.37,2.27],[4.33,2.29],[4.35,2.31],[4.39,2.30],[4.40,2.25]],
];

function slide20(s) {
  s.background = { color: C.rust };
  ISLANDS.forEach(function (poly) {
    const xs = poly.map(function (p) { return p[0]; });
    const ys = poly.map(function (p) { return p[1]; });
    const x = Math.min.apply(null, xs), y = Math.min.apply(null, ys);
    const w = Math.max.apply(null, xs) - x, h = Math.max.apply(null, ys) - y;
    freeform(s, poly.map(function (p) { return [(p[0] - x) / w, (p[1] - y) / h]; }), {
      x: x, y: y, w: w, h: h, color: C.mapFill, line: { color: C.cream, width: 0.75 },
    });
  });
  T(s, 'PRODUCT MAP SPREAD',
    { x: 0.69, y: 0.563, w: 9.46, h: 1.01, fontSize: 54, bold: true, color: C.cream });
  star(s, 11.519, 0.375, 1.407, C.cream);
  T(s, '2488+', { x: 0.599, y: 4.391, w: 2.389, h: 0.841, fontSize: 44, color: C.cream });

  rect(s, 0.659, 5.49, 3.561, 0.501, C.cream, 10);
  rect(s, 0.659, 5.49, 2.221, 0.501, C.green);
  T(s, '60%', { x: 0.74, y: 5.555, w: 0.841, h: 0.37, fontSize: 16, color: C.cream });

  P(s, L.short, { x: 1.531, y: 6.223, w: 5.343, h: 0.684, color: C.cream });
  rule(s, 0.667, 6.432, 0.773, C.green, 1.25);
  P(s, 'FERTILIZER SYSTEM', {
    x: 11.102, y: 6.638, w: 1.565, h: 0.269, fontSize: 10,
    color: C.cream, align: 'right', lineSpacingMultiple: 1,
  });
}

/* 21 — thank you */
function slide21(s) {
  s.background = { color: C.green };
  star(s, 0, 3.163, 4.357, C.cream);
  rect(s, 0, 0, SLIDE_W, 3.75, C.rust);

  T(s, 'GET IN TOUCH WITH US ', { x: 2.42, y: 0.714, w: 6.39, h: 0.505, fontSize: 24, color: C.cream });
  rule(s, 0.673, 0.967, 1.705, C.cream, 1.5);
  P(s, L.long, { x: 2.42, y: 1.562, w: 8.847, h: 0.643, color: C.cream, lineSpacingMultiple: 1.2 });
  [2.803, 5.546, 8.289].forEach(function (x, i) {
    T(s, 'YOUR TEXT HERE ', { x: x, y: 2.471, w: 2.202, h: 0.37, fontSize: 16, color: C.cream });
    arrow(s, 2.545 + i * 2.743, 2.557, 0.17, C.cream);
  });

  T(s, 'THANK', { x: 7.41, y: 4.436, w: 6.074, h: 1.717, fontSize: 96, color: C.cream });
  T(s, 'YOU', { x: 9.892, y: 5.605, w: 3.681, h: 1.717, fontSize: 96, color: C.cream });
  rule(s, 7.56, 6.495, 2.093, C.rust, 1.5);
}

/* -------------------------------------------------------------- assemble */

const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07,
  slide08, slide09, slide10, slide11, slide12, slide13, slide14,
  slide15, slide16, slide17, slide18, slide19, slide20, slide21,
];

function build() {
  pres = new PptxGenJS();
  pres.defineLayout({ name: 'WIDE', width: SLIDE_W, height: SLIDE_H });
  pres.layout = 'WIDE';
  pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };

  BUILDERS.forEach(function (builder, i) {
    const slide = pres.addSlide();
    slide.background = { color: C.cream };
    if (i !== 0 && i !== BUILDERS.length - 1) pageNumber(slide, i + 1);
    builder(slide);
  });

  return pres.writeFile({
    fileName: path.join(__dirname, '00f98224-a5fe-4adf-ab04-310a4097fafa_grok_final.pptx'),
  });
}

build().then(function (f) { console.log('wrote', f); }, function (e) {
  console.error(e);
  process.exit(1);
});
