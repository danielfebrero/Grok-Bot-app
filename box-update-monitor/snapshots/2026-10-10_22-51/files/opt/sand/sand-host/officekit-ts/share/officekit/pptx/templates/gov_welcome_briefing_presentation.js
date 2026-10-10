/**
 * Recreation of "138b83a3-bf31-4ba6-bb91-12c8647fe285.pptx" with pptxgenjs.
 *
 * 30 slides, 13.333 x 7.5 in.  Monochrome editorial template: charcoal panels,
 * mist-grey diagonal washes, thin rules, Lato headings / Open Sans body copy.
 * Raster photos in the original are redrawn as flat grey placeholder blocks.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const INK = '2B2B2B'; // theme dk1 - charcoal panels & headings
const PAPER = 'FFFFFF';
const MIST = 'F2F2F2'; // 5% grey - diagonal washes / soft panels
const RULE = 'CACACA'; // hairlines & progress-bar tracks
const MUTED = '959595'; // body copy
const SLATE = '404040'; // pill labels & bullet dots
const PHOTO = 'DDDDDD'; // stand-in for the deck's photographs

const HEAD = 'Lato'; // theme major font
const BODY = 'Open Sans'; // theme minor font

const W = 13.3333;
const H = 7.5;

/* ------------------------------------------------------------------ helpers */

// Every text box in the source is top-anchored with a 0.1/0.05 in inset.
function text(s, content, o) {
  s.addText(content, Object.assign({ valign: 'top', fontFace: BODY, color: INK }, o));
}

function box(s, x, y, w, h, color) {
  s.addShape('rect', { x, y, w, h, fill: { color } });
}

function hline(s, x, y, w) {
  s.addShape('line', { x, y, w, h: 0, line: { color: RULE, width: 1.5 } });
}

function vline(s, x) {
  s.addShape('line', { x, y: 0, w: 0, h: H, line: { color: RULE, width: 1.5 } });
}

// 34pt / 40pt / 44pt Lato headline, one entry per line.
function heading(s, lines, o) {
  const runs = lines.map((t, i) => ({ text: t, options: { breakLine: i < lines.length - 1 } }));
  text(s, runs, Object.assign({ fontFace: HEAD, bold: true, fontSize: 34 }, o));
}

// Justified 10.5pt grey paragraph(s) at 1.5 line spacing.
function copy(s, paras, o) {
  const runs = [].concat(paras).map((t, i, a) => ({ text: t, options: { breakLine: i < a.length - 1 } }));
  text(s, runs, Object.assign(
    { fontSize: 10.5, color: MUTED, align: 'justify', lineSpacingMultiple: 1.5 }, o));
}

// "1. First Service" + supporting sentence, the deck's recurring two-line block.
function feature(s, x, y, title, blurb, w, blurbH, color) {
  text(s, title, { x, y, w, h: 0.337, fontFace: HEAD, bold: true, fontSize: 14, color: color || INK });
  copy(s, blurb, { x: x + 0.001, y: y + 0.528, w, h: blurbH || 0.868, color: color || MUTED });
}

function dot(s, x, y) {
  s.addShape('ellipse', { x, y, w: 0.104, h: 0.104, fill: { color: SLATE } });
}

// Rounded outline "call to action" bar: dot + letter-spaced caption.
function pill(s, x, y, w, o) {
  o = o || {};
  s.addShape('roundRect', {
    x, y, w, h: 0.604, rectRadius: 0.0955,
    fill: { color: PAPER }, line: { color: INK, width: 1.5 },
  });
  dot(s, x + (o.dotDx || 0.264), y + 0.244);
  text(s, o.label || 'Write Something Here ', {
    x: x + (o.textDx || 0.715), y: y + 0.128, w: o.textW || 3.701, h: 0.304,
    fontSize: 9, charSpacing: 6, color: SLATE, align: 'justify', lineSpacingMultiple: 1.5,
  });
}

// The "75.  42.  87." figures that trail most pills.
function stats(s, x, y, values) {
  (values || ['75.', '42.', '87.']).forEach((v, i) => {
    text(s, v, {
      x: x + i * 0.4705, y, w: 0.414, h: 0.304, fontSize: 9, bold: true,
      color: SLATE, align: 'center', lineSpacingMultiple: 1.5,
    });
  });
}

// Percentage meter: grey track + charcoal fill + caption + value.
function meter(s, x, y, w, o) {
  o = o || {};
  const fill = o.fill || w * 0.863;
  s.addShape('roundRect', { x, y, w, h: 0.161, rectRadius: 0, fill: { color: RULE } });
  s.addShape('roundRect', { x, y, w: fill, h: 0.161, rectRadius: 0, fill: { color: INK } });
  copy(s, 'lorem ipsum dolor', { x: x - 0.045, y: y - 0.46, w: 1.72, h: 0.304, fontSize: 9 });
  text(s, [{ text: '81', options: { fontSize: 10.5 } }, { text: '%', options: { fontSize: 9 } }], {
    x: x + w - 0.47, y: y - 0.482, w: 0.517, h: 0.338, align: 'right', lineSpacingMultiple: 1.5,
  });
}

// Three stacked bars, the deck's "menu" glyph.
function burger(s, x, y) {
  for (let i = 0; i < 3; i++) {
    s.addShape('roundRect', { x, y: y + i * 0.0785, w: 0.293, h: 0.032, rectRadius: 0.0148, fill: { color: RULE } });
  }
}

// Charcoal vertical tab with rotated "MARCH - 35" caption (+ optional white foot).
function sideTab(s, x, y, h, o) {
  o = o || {};
  box(s, x, y, 0.606, h, INK);
  if (o.foot !== false) box(s, x, y + h, 0.606, 0.606, PAPER);
  text(s, o.label || 'MARCH - 35', {
    x: x + 0.124, y: y + (o.capDy === undefined ? 0.188 : o.capDy), w: 0.436, h: 2.573,
    vert: 'wordArtVert', align: 'center', fontSize: 7, charSpacing: 6,
    color: PAPER, lineSpacingMultiple: 1.5,
  });
  burger(s, x + 0.166, y + h + 0.236);
}

/* ------------------------------------------------- background decorations */

function photo(s, x, y, w, h) {
  box(s, x, y, w, h, PHOTO);
}

// Flat polygon from absolute corner points, in inches.
function poly(s, pts, color) {
  const xs = pts.map(function (p) { return p[0]; });
  const ys = pts.map(function (p) { return p[1]; });
  const x = Math.min.apply(null, xs);
  const y = Math.min.apply(null, ys);
  s.addShape('custGeom', {
    x, y, w: Math.max.apply(null, xs) - x, h: Math.max.apply(null, ys) - y,
    fill: { color },
    points: pts.map(function (p) { return { x: p[0] - x, y: p[1] - y }; }).concat([{ close: true }]),
  });
}

function washLine(s, x, y, w, h) {
  s.addShape('line', { x, y, w, h, flipV: true, line: { color: MIST, width: 1.5 } });
}

// Half-slide diagonal band: rectangle with the top-left corner sliced off at
// `cut` (fraction of the box), plus the hairline that trails it.
function diagWash(s, x, y, w, h, cut, line) {
  washLine(s, line[0], line[1], line[2], line[3]);
  poly(s, [[x, y + h * cut], [x + w * cut, y], [x + w, y], [x, y + h]], MIST);
}

// Full-slide slanted panel: parallelogram leaning right by `lean` inches.
function slantWash(s, x, w, lean, line) {
  washLine(s, line[0], line[1], line[2], line[3]);
  poly(s, [[x, H], [x + lean, 0], [x + w, 0], [x + w - lean, H]], MIST);
}

// Charcoal wedge in the top-left corner.
function cornerTL(s) {
  poly(s, [[0, 0], [1.204, 0], [0, 1.176]], INK);
}

// Charcoal wedge + hairline in the bottom-right corner.
function cornerBR(s, color) {
  poly(s, [[13.333, 6.32], [13.333, H], [12.125, H]], color || INK);
  s.addShape('line', { x: 11.394, y: 5.717, w: 2.064, h: 2.082, flipV: true, line: { color: RULE, width: 1.5 } });
}

/* -------------------------------------------------------- device mock-ups */
/* Photographic device renders in the source, rebuilt from native shapes.    */

function caption(s, x, y, w) {
  text(s, '[image]', { x, y, w, h: 0.3, align: 'center', fontSize: 9, color: MUTED });
}

function imac(s, x, y, w, h) {
  box(s, x + 0.107, y + 0.802, w - 0.213, 3.186, INK); // bezel
  box(s, x + 0.307, y + 1.002, w - 0.613, 2.786, PHOTO); // screen
  box(s, x + 0.107, y + 3.988, w - 0.213, 0.487, 'C4C4C4'); // chin
  box(s, x + 2.283, y + 4.475, 0.947, 0.72, 'D0D0D0'); // neck
  poly(s, [[x + 2.033, y + 5.195], [x + 3.483, y + 5.195],
           [x + 3.983, y + 5.335], [x + 1.533, y + 5.335]], 'E0E0E0'); // base
  caption(s, x + 0.307, y + 2.28, w - 0.613);
}

function phone(s, x, y, w, h) {
  s.addShape('roundRect', { x, y, w, h, rectRadius: 0.32, fill: { color: INK } });
  s.addShape('roundRect', {
    x: x + 0.169, y: y + 0.163, w: w - 0.341, h: h - 0.359, rectRadius: 0.22, fill: { color: PHOTO },
  });
  s.addShape('roundRect', {
    x: x + w / 2 - 0.62, y: y + 0.163, w: 1.24, h: 0.173, rectRadius: 0.06, fill: { color: INK },
  }); // notch
  caption(s, x, y + h / 2 - 0.15, w);
}

function laptop(s, x, y, w, h) {
  box(s, x + 0.613, y + 0.019, w - 1.226, h - 0.286, INK); // lid
  box(s, x + 0.813, y + 0.206, w - 1.626, h - 0.66, PHOTO); // screen
  poly(s, [[x + 0.613, y + h - 0.267], [x + w - 0.613, y + h - 0.267],
           [x + w, y + h - 0.101], [x, y + h - 0.101]], 'D8D8D8'); // base
  box(s, x, y + h - 0.101, w, 0.033, '6E6E6E'); // front lip
  caption(s, x + 0.813, y + h / 2 - 0.3, w - 1.626);
}

function tablet(s, x, y, w, h) {
  s.addShape('roundRect', { x, y, w, h, rectRadius: 0.22, fill: { color: INK } });
  box(s, x + 0.213, y + 0.562, w - 0.433, 4.127, PHOTO);
  s.addShape('ellipse', { x: x + w / 2 - 0.11, y: y + h - 0.42, w: 0.22, h: 0.22, fill: { color: '3D3D3D' } });
  caption(s, x + 0.213, y + h / 2 - 0.15, w - 0.433);
}

/* ------------------------------------------------------------ shared copy */

const LOREM_TITLE =
  'PLACEHOLDER';
const LOREM_LONG =
  'PLACEHOLDER';
const LOREM_ABOUT =
  'PLACEHOLDER';
const LOREM_SUB = 'misani utasisau nasat sapienan anis tinciduntan';
const LOREM_SUBA = 'misani utasisau nasat sapienana';
const SERVICE_BLURB = 'PLACEHOLDER';
const SERVICE_SHORT = 'PLACEHOLDER';
const SERVICE_GRID = 'PLACEHOLDER';
const SERVICE_WIDE = 'PLACEHOLDER';
const SERVICE_CARD = 'PLACEHOLDER';
const PORT_BLURB = 'PLACEHOLDER';
const MOCK_BLURB = 'PLACEHOLDER';
const CHART_BLURB = 'lorem ipsum dolor sitali hanisa metcon secilsain';

/* ------------------------------------------------------------ chart setup */

const CHART_GREYS = { light: 'CACACA', mid: '959595', dark: '606060', deep: '404040' };

const AXIS = {
  showLegend: false,
  chartArea: { roundedCorners: false }, // transparent - slide art shows through
  catAxisLabelColor: '757575',
  catAxisLabelFontSize: 8,
  valAxisLabelColor: '757575',
  valAxisLabelFontSize: 8,
  catAxisLineShow: false,
  valAxisLineShow: false,
  catAxisMajorTickMark: 'none',
  valAxisMajorTickMark: 'none',
  valGridLine: { color: 'D9D9D9', size: 0.75 },
  catGridLine: { style: 'none' },
};

const SER3 = [
  { name: 'Series 1', labels: ['1', '2', '3'], values: [4.3, 2.5, 3.5] },
  { name: 'Series 2', labels: ['1', '2', '3'], values: [2.4, 4.4, 1.8] },
  { name: 'Series 3', labels: ['1', '2', '3'], values: [2, 2, 3] },
];
const SER4 = [
  { name: 'Series 1', labels: ['1', '2', '3', '4'], values: [4.3, 2.5, 3.5, 4.5] },
  { name: 'Series 2', labels: ['1', '2', '3', '4'], values: [2.4, 4.4, 1.8, 2.8] },
  { name: 'Series 3', labels: ['1', '2', '3', '4'], values: [2, 2, 3, 5] },
];
// Area chart categories are date serials rendered through an "m/d/yyyy" axis
// (37377 = 5/1/2002 ... 37500 = 9/1/2002).
const AREA_DATES = ['37377', '37408', '37438', '37469', '37500'];
const AREA_RAW = [[32, 32, 28, 12, 15], [12, 12, 12, 21, 28]];
// The source uses a 100%-stacked area; pptxgenjs only emits plain "stacked",
// so the series are pre-normalised to their share of each column.
const AREA_SER = AREA_RAW.map(function (values, i) {
  return {
    name: 'Series ' + (i + 1),
    labels: AREA_DATES,
    values: values.map(function (v, j) { return v / (AREA_RAW[0][j] + AREA_RAW[1][j]); }),
  };
});

/* ---------------------------------------------------------------- slides */

const slides = [];

// 1 - opening title: charcoal left panel, oversized "GOV." straddling the edge.
slides.push(function (s) {
  box(s, 10.75, 0, 2.583, H, MIST);
  vline(s, 5.177);
  box(s, 0, 0, 4.681, H, INK);
  photo(s, 3.444, 1.052, 7.305, 5.397);

  box(s, 10.75, 1.052, 0.606, 4.791, INK);
  box(s, 10.75, 5.842, 0.606, 0.606, PAPER);
  text(s, 'PRESENTATION', {
    x: 10.874, y: 1.052, w: 0.436, h: 4.759, vert: 'wordArtVert', align: 'center',
    fontSize: 7, charSpacing: 6, color: PAPER, lineSpacingMultiple: 1.5,
  });
  burger(s, 10.925, 6.079);

  text(s, [
    { text: 'G', options: { color: PAPER } },
    { text: 'OV.', options: { color: INK } },
  ], { x: 1.474, y: 1.953, w: 4.832, h: 2.036, fontFace: HEAD, bold: true, fontSize: 115, charSpacing: 6 });
  // white half-ring completing the "O" where it crosses onto the dark panel
  s.addShape('blockArc', {
    x: 2.868, y: 2.414, w: 1.152, h: 1.176, angleRange: [90, 270], arcThicknessRatio: 0.38,
    fill: { color: PAPER },
  });
  copy(s, LOREM_TITLE, { x: 1.577, y: 3.979, w: 3.983, h: 0.868 });
});

// 2 - welcome, portrait photo left, copy right.
slides.push(function (s) {
  diagWash(s, 0, 0, 5.724, 5.773, 0.68556, [-0.136, -0.287, 3.895, 3.929]);
  cornerBR(s);
  photo(s, 2.297, 1.177, 3.875, 5.146);

  heading(s, ['Welcome Message'], { x: 7.41, y: 1.572, w: 2.815, h: 1.313, fontSize: 36 });
  copy(s,
    'PLACEHOLDER',
    { x: 7.41, y: 4.437, w: 3.983, h: 1.663 });
  pill(s, 7.255, 3.361, 4.417);
  sideTab(s, 1.691, 1.177, 4.54);
});

// 3 - mirrored welcome, photo right.
slides.push(function (s) {
  slantWash(s, 0.81, 12.524, 7.405, [6.397, 0.506, 7.016, 7.077]);
  cornerTL(s);
  photo(s, 6.069, 1.146, 5.014, 5.146);

  pill(s, 1.143, 5.67, 3.833, { label: 'Something Here ', dotDx: 0.348, textDx: 0.801, textW: 2.72 });
  sideTab(s, 11.083, 1.146, 4.54, { capDy: 0.458 });
  heading(s, ['Welcome Message'], { x: 1.378, y: 1.563, w: 2.815, h: 1.313, fontSize: 36 });
  copy(s,
    'lorem ipsum dolor sita amet consecanila sunian teturasi adipiscingesa nibuhase eget justo sedal gravida est sedi onsecasuimasil teturasanit elitas utasisau nasat sapienan tincidunt sanilma velas aliquetacursus mi ut sapien tincidunt velait asain   nibhasi elitas sedas ',
    { x: 1.378, y: 3.337, w: 3.515, h: 1.663 });
});

// 4 - landscape photo with charcoal right panel.
slides.push(function (s) {
  box(s, 8.653, 0, 4.681, H, INK);
  photo(s, 2.361, 1.329, 7.306, 3.438);

  box(s, 2.361, 5.115, 7.306, 0.888, MIST);
  hline(s, 2.361, 6.47, 7.306);
  heading(s, ['Welcome Message'], { x: 7.879, y: 1.854, w: 4.418, h: 0.707, fontSize: 36, color: PAPER, align: 'right' });
  sideTab(s, 1.356, 1.329, 4.673);
  copy(s,
    'PLACEHOLDER',
    { x: 10.287, y: 2.953, w: 1.969, h: 1.663, color: PAPER });
  dot(s, 2.787, 5.508);
  text(s, 'Write Something Here ', {
    x: 3.239, y: 5.391, w: 3.701, h: 0.304, fontSize: 9, charSpacing: 6,
    color: SLATE, align: 'justify', lineSpacingMultiple: 1.5,
  });
  stats(s, 6.693, 5.391);
});

// 5 - two staggered photos with a floating pill.
slides.push(function (s) {
  diagWash(s, 0.002, 0, 4.494, 4.472, 0.5449, [-0.264, -0.167, 2.244, 2.263]);
  cornerBR(s);
  photo(s, 1.63, 3.75, 5.037, 2.833);
  photo(s, 6.667, 0.917, 5.037, 2.833);

  copy(s,
    'PLACEHOLDER',
    { x: 7.72, y: 4.817, w: 3.983, h: 1.398 });
  heading(s, ['Welcome Message'], { x: 1.564, y: 1.535, w: 4.627, h: 0.673 });
  copy(s, LOREM_SUB, { x: 1.564, y: 2.284, w: 3.983, h: 0.338 });
  burger(s, 12.443, 0.657);
  pill(s, 4.559, 3.447, 4.215, { textW: 3.331 });
});

// 6 - vertical charcoal banner with rotated headline.
slides.push(function (s) {
  vline(s, 2.164);
  box(s, 6.667, 0, 6.667, H, MIST);
  photo(s, 1.125, 1.094, 3.789, 5.312);

  pill(s, 8.742, 3.448, 3.108, { label: 'Something Here ', dotDx: 0.298, textDx: 0.638, textW: 2.307 });
  copy(s,
    'lorem ipsum dolor sita amet consecanila sunian teturasi adipiscinges hani nibuhase eget justo sedaluin gravida estasunisa sedi onsecasuimasil teturasanit elitas utasisau velait asain   nibhasi elitas sedas ',
    { x: 8.742, y: 1.473, w: 3.108, h: 1.398 });
  meter(s, 8.815, 5.065, 2.952, { fill: 2.548 });
  meter(s, 8.815, 5.979, 2.952, { fill: 2.548 });
  box(s, 5.333, 1.094, 2.197, 5.312, INK);
  heading(s, ['Welcome Message'], {
    x: 5.009, y: 3.169, w: 2.815, h: 1.178, fontSize: 32, color: PAPER, align: 'center', rotate: 270,
  });
  burger(s, 12.443, 0.657);
});

// 7 - service intro, two feature columns.
slides.push(function (s) {
  slantWash(s, 3.506, 12.524, 7.405, [9.545, 0.506, 7.016, 7.077]);
  cornerTL(s);
  photo(s, 7.923, 1.329, 2.779, 5.279);

  sideTab(s, 11.061, 1.329, 4.673);
  heading(s, ['Service'], { x: 1.372, y: 1.691, w: 2.245, h: 0.673 });
  pill(s, 1.283, 3.422, 5.721);
  copy(s,
    'PLACEHOLDER',
    { x: 3.908, y: 1.77, w: 3.108, h: 0.868 });
  feature(s, 1.426, 4.817, '1. First Service', SERVICE_BLURB, 2.539);
  feature(s, 4.465, 4.817, '2. Second Service', SERVICE_BLURB, 2.539);
});

// 8 - charcoal service list on the right.
slides.push(function (s) {
  box(s, 8.653, 0, 4.681, H, INK);
  photo(s, 1.113, 1.271, 6.425, 2.479);

  ['1. First Service', '2. Second Service', '3. Third Service'].forEach(function (t, i) {
    feature(s, 9.681, 1.473 + i * 1.7115, t, SERVICE_SHORT, 2.539, 0.603, MIST);
  });
  copy(s,
    'PLACEHOLDER',
    { x: 1.36, y: 4.316, w: 5.932, h: 1.133 });
  burger(s, 12.443, 0.657);
  pill(s, 1.113, 5.91, 6.425);
  stats(s, 5.923, 6.038);
});

// 9 - about + 2x2 service grid.
slides.push(function (s) {
  diagWash(s, 0.011, 0.006, 4.04, 4.049, 0.5449, [-0.192, -0.146, 2, 2.018]);
  cornerBR(s);
  photo(s, 1.804, 1.339, 3.736, 0.918);

  pill(s, 6.542, 1.484, 5.387);
  [['1. First Service', 6.685, 3.005], ['2. Second Service', 9.516, 3.005],
   ['3. Third Service', 6.685, 5.014], ['4. Fourth Service', 9.516, 5.014]].forEach(function (f) {
    feature(s, f[1], f[2], f[0], SERVICE_GRID, 2.289);
  });
  heading(s, ['About ', 'Our Service'], { x: 1.74, y: 5.266, w: 3.736, h: 1.245 });
  copy(s, LOREM_LONG, { x: 1.804, y: 3.003, w: 3.736, h: 1.663 });
  burger(s, 12.443, 0.657);
});

// 10 - charcoal left column, wide intro paragraph.
slides.push(function (s) {
  vline(s, 3.245);
  box(s, 0, 0, 2.369, H, INK);
  poly(s, [[13.333, 6.32], [13.333, H], [12.125, H]], MIST);
  photo(s, 1.525, 3.284, 3.141, 3.141);

  pill(s, 5.671, 3.591, 6.25);
  feature(s, 5.856, 4.923, '1. First Service', SERVICE_WIDE, 2.691);
  feature(s, 9.026, 4.923, '2. Second Service', SERVICE_WIDE, 2.691);
  // the headline straddles the charcoal column, so its first letters reverse out
  text(s, [
    { text: 'Ab', options: { color: PAPER } },
    { text: 'out ', options: { color: INK, breakLine: true } },
    { text: 'Ou', options: { color: PAPER } },
    { text: 'r Service', options: { color: INK } },
  ], { x: 1.657, y: 1.273, w: 3.736, h: 1.245, fontFace: HEAD, bold: true, fontSize: 34 });
  burger(s, 12.443, 0.657);
  copy(s,
    'PLACEHOLDER',
    { x: 5.856, y: 1.34, w: 5.861, h: 1.133 });
});

// 11 - three-up card band on a mist panel.
slides.push(function (s) {
  photo(s, 1.22, 5.533, 5.447, 0.935);

  box(s, 1.238, 2.932, 10.921, 2.095, MIST);
  ['1. First Service', '2. Second Service', '3. Third Service'].forEach(function (t, i) {
    feature(s, 1.891 + i * 3.4915, 3.459, t, SERVICE_CARD, 2.691, 0.603);
  });
  hline(s, 1.238, 2.448, 10.921);
  pill(s, 6.686, 1.103, 5.473);
  copy(s,
    'PLACEHOLDER',
    { x: 7.312, y: 5.566, w: 4.253, h: 0.868 });
  heading(s, ['About Our Service'], { x: 1.238, y: 1.062, w: 4.845, h: 0.673 });
  burger(s, 12.443, 6.662);
});

// 12 - portfolio mosaic with charcoal right panel + rotated caption.
slides.push(function (s) {
  box(s, 10.145, 0, 3.188, H, INK);
  photo(s, 1.955, 1.756, 3.989, 3.989);
  // the four small tiles butt together into two continuous columns
  photo(s, 6.564, 1.756, 1.974, 3.947);
  photo(s, 9.159, 1.756, 1.974, 3.947);

  heading(s, ['About Our Portfolio'], { x: 2.603, y: 3.14, w: 2.691, h: 1.178, fontSize: 32, align: 'center' });
  const label = { w: 1.974, h: 0.337, fontFace: HEAD, bold: true, fontSize: 14, align: 'center' };
  text(s, '1. First Portfolio', Object.assign({ x: 6.564, y: 6.252 }, label));
  text(s, '2. SecondPortfolio', Object.assign({ x: 6.564, y: 0.939 }, label));
  // these two straddle the charcoal panel, so their tails are reversed out
  text(s, [{ text: '3. Third ' }, { text: 'Portfolio', options: { color: PAPER } }],
    Object.assign({ x: 9.18, y: 0.939 }, label));
  text(s, [{ text: '4. Fourth ' }, { text: 'Portfolio', options: { color: PAPER } }],
    Object.assign({ x: 9.159, y: 6.252 }, label));

  box(s, 1.341, 1.756, 0.613, 3.989, INK);
  text(s, 'MARCH - 35', {
    x: 1.473, y: 1.992, w: 0.436, h: 2.573, vert: 'wordArtVert', align: 'center',
    fontSize: 7, charSpacing: 6, color: PAPER, lineSpacingMultiple: 1.5,
  });
  burger(s, 1.359, 6.31);
  copy(s,
    'lorem ipsum dolor sitas amet hanis conse alcanila sunian teturasi adipiscinges haninaisa  an libuhase',
    { x: 2.285, y: 6.101, w: 3.658, h: 0.603 });
  hline(s, 1.341, 1.109, 4.602);
  // the caption group is turned a quarter turn onto the charcoal panel
  s.addShape('ellipse', { x: 11.899, y: 2.051, w: 0.104, h: 0.104, fill: { color: PAPER } });
  text(s, 'Write Something Here ', {
    x: 10.367, y: 3.951, w: 3.201, h: 0.304, rotate: 90, fontSize: 9, charSpacing: 6,
    color: PAPER, align: 'justify', lineSpacingMultiple: 1.5,
  });
});

// 13 - portfolio list, three thumbnails on the right.
slides.push(function (s) {
  diagWash(s, 0.002, 0, 4.494, 4.472, 0.5449, [-0.264, -0.167, 2.244, 2.263]);
  cornerBR(s);
  [1.302, 3.124, 4.946].forEach(function (y) { photo(s, 7.037, y, 1.625, 1.625); });

  heading(s, ['About Our Portfolio'], { x: 1.564, y: 1.535, w: 4.627, h: 0.673 });
  copy(s, LOREM_SUB, { x: 1.564, y: 2.284, w: 3.983, h: 0.338 });
  burger(s, 12.443, 0.657);
  ['1. First Portfolio', '2. Second Portfolio', '3. Third Portfolio'].forEach(function (t, i) {
    feature(s, 9.342, 1.549 + i * 1.8215, t, PORT_BLURB, 2.427, 0.603);
  });
  copy(s, LOREM_LONG, { x: 1.564, y: 3.371, w: 3.736, h: 1.663 });
  pill(s, 1.3, 5.758, 4.686, { label: 'Something Here ', dotDx: 0.349, textDx: 0.801, textW: 2.72 });
  stats(s, 4.739, 5.891, ['75.', '42.']);
});

// 14 - full-bleed portfolio grid with charcoal caption block.
slides.push(function (s) {
  photo(s, 6.487, 1.49, 4.513, 2.071);
  [[1.603, 1.49], [4.045, 1.49], [6.487, 3.938], [8.929, 3.938]].forEach(function (p) {
    photo(s, p[0], p[1], 2.071, 2.071);
  });

  box(s, 0, 3.938, 6.116, 3.562, MIST);
  box(s, 1.603, 3.938, 4.513, 2.071, INK);
  heading(s, ['Our Portfolio'], { x: 2.207, y: 4.467, w: 3.306, h: 0.673, color: PAPER, align: 'center' });
  copy(s, 'misani utasisau nasat sapienan anis tinciduni', { x: 2.207, y: 5.166, w: 3.306, h: 0.338, align: 'center' });
  const label = { w: 2.071, h: 0.303, fontFace: HEAD, bold: true, fontSize: 12, align: 'center' };
  text(s, '1. First Portfolio', Object.assign({ x: 1.603, y: 0.766 }, label));
  text(s, '2. Second Portfolio', Object.assign({ x: 4.045, y: 0.766 }, label));
  text(s, '3. Third Portfolio', Object.assign({ x: 6.487, y: 6.431 }, label));
  text(s, '4. Fourth Portfolio', Object.assign({ x: 8.929, y: 6.431 }, label));
  hline(s, 6.487, 0.917, 4.513);
  sideTab(s, 11.311, 1.49, 4.519, { foot: false, capDy: 0.459 });
});

// 15 - two tall portfolio photos, copy on the right.
slides.push(function (s) {
  slantWash(s, 0.19, 11.289, 7.405, [4.59, -0.118, 7.668, 7.735]);
  cornerTL(s);
  photo(s, 1.382, 1.582, 2.806, 3.758);
  photo(s, 4.522, 1.582, 2.806, 3.758);

  pill(s, 1.382, 5.732, 5.946);
  stats(s, 5.72, 5.866);
  heading(s, ['Our Portfolio'], { x: 8.385, y: 1.928, w: 3.857, h: 0.673 });
  copy(s, [
    'PLACEHOLDER',
    '',
    'PLACEHOLDER',
  ], { x: 8.417, y: 3.118, w: 3.736, h: 1.928 });
  text(s, '1. First Portfolio', { x: 8.385, y: 5.888, w: 1.771, h: 0.32, fontFace: HEAD, bold: true, fontSize: 13 });
  text(s, '2. Second Portfolio', { x: 10.304, y: 5.888, w: 1.84, h: 0.32, fontFace: HEAD, bold: true, fontSize: 13, align: 'right' });
  burger(s, 12.443, 0.657);
});

// 16 - six-up portfolio grid, charcoal sidebar.
slides.push(function (s) {
  box(s, 0, 0, 3.896, H, INK);
  [[5.579, 1.219], [7.773, 1.219], [9.968, 1.219],
   [5.576, 4.603], [7.771, 4.603], [9.965, 4.603]].forEach(function (p) {
    photo(s, p[0], p[1], 1.833, 1.678);
  });

  pill(s, 5.208, 3.448, 6.854);
  stats(s, 10.443, 3.581);
  burger(s, 0.624, 0.657);
  const label = { h: 0.286, fontFace: HEAD, bold: true, fontSize: 11, align: 'center' };
  text(s, '1. First Portfolio', Object.assign({ x: 5.579, y: 0.653, w: 1.833 }, label));
  text(s, '2. Second Portfolio', Object.assign({ x: 7.773, y: 0.653, w: 1.854 }, label));
  text(s, '3. Third Portfolio', Object.assign({ x: 9.968, y: 0.653, w: 1.854 }, label));
  text(s, '4. Fourth Portfolio', Object.assign({ x: 5.576, y: 6.561, w: 1.833 }, label));
  text(s, '5. Fifth Portfolio', Object.assign({ x: 7.771, y: 6.561, w: 1.854 }, label));
  text(s, '6. Sixth Portfolio', Object.assign({ x: 9.965, y: 6.561, w: 1.854 }, label));
  heading(s, ['About', 'Our ', 'Portfolio'], { x: 0.949, y: 2.593, w: 2.383, h: 1.717, fontSize: 32, color: PAPER });
  copy(s, ['misani utasisau nasas sapi', 'enan anisat'], { x: 0.949, y: 4.385, w: 2.383, h: 0.603, align: 'left' });
});

// 17 - team hero photo with meters.
slides.push(function (s) {
  diagWash(s, 0, 0, 7.683, 7.494, 0.5449, [-0.203, -0.045, 3.705, 3.738]);
  cornerBR(s);
  photo(s, 1.683, 1.222, 6, 5.056);

  meter(s, 8.572, 4.871, 2.949, { fill: 2.545 });
  meter(s, 8.572, 5.77, 2.949, { fill: 2.545 });
  pill(s, 7.178, 3.353, 4.417);
  heading(s, ['Our Team'], { x: 2.702, y: 3.296, w: 3.898, h: 0.673, color: PAPER, align: 'center' });
  copy(s,
    'PLACEHOLDER',
    { x: 8.471, y: 1.667, w: 3.179, h: 1.133 });
});

// 18 - team member card + three head-shots.
slides.push(function (s) {
  box(s, 0, 0, 4.651, H, INK);
  photo(s, 1.127, 1.396, 2.397, 2.445);
  [6.096, 8.19, 10.285].forEach(function (x) { photo(s, x, 1.396, 1.763, 1.763); });

  ['1. Add Your Name', '2. Add Your Name', '3. Add Your Name'].forEach(function (t, i) {
    text(s, t, { x: 6.096 + i * 2.0945, y: 3.592, w: 1.763, h: 0.286, fontFace: HEAD, bold: true, fontSize: 11, align: 'center' });
  });
  heading(s, ['Your Name'], { x: 1.127, y: 4.375, w: 2.397, h: 0.438, fontSize: 20, color: PAPER, align: 'center' });
  copy(s, LOREM_ABOUT, { x: 1.127, y: 5.109, w: 2.397, h: 1.398, align: 'center' });
  pill(s, 6.096, 4.52, 5.952);
  stats(s, 10.449, 4.654);
  heading(s, ['About Our Team'], { x: 6.096, y: 5.768, w: 5.952, h: 0.673, align: 'center' });
  hline(s, 6.096, 0.993, 5.952);
  burger(s, 12.443, 6.659);
});

// 19 - single team member, meters on the right.
slides.push(function (s) {
  slantWash(s, 0, 11.018, 7.405, [4.294, -0.118, 7.484, 7.735]);
  cornerBR(s);
  photo(s, 2.283, 1.246, 3.708, 5.007);

  sideTab(s, 1.677, 1.246, 4.401, { capDy: 0.26 });
  heading(s, ['Your Name'], { x: 7.463, y: 1.443, w: 4.039, h: 0.673 });
  copy(s,
    'PLACEHOLDER',
    { x: 7.463, y: 2.416, w: 3.857, h: 0.603 });
  meter(s, 7.545, 4.892, 3.656, { fill: 3.155 });
  meter(s, 7.545, 5.791, 3.656, { fill: 3.155 });
  pill(s, 7.255, 3.469, 4.417);
});

// 20 - four team portraits over paired caption blocks.
slides.push(function (s) {
  box(s, 6.886, 0, 6.448, 3.606, MIST);
  [1.729, 4.307, 6.886, 9.464].forEach(function (x) { photo(s, x, 1.466, 2.14, 2.14); });

  box(s, 6.886, 4.013, 4.718, 2.14, INK);
  box(s, 1.729, 4.019, 4.718, 2.14, MIST);
  heading(s, ['About Our Team'], { x: 1.729, y: 4.617, w: 4.718, h: 0.606, fontSize: 30, align: 'center' });
  copy(s, 'misani utasisau nasat sapienan', { x: 1.729, y: 5.223, w: 4.718, h: 0.338, align: 'center' });
  copy(s,
    'lorem ipsum dolor sitanu atenasan an hanisani amet consecanila sunian onseca  haiat suimasil velas aliqueta nisasa ladipiscinge husa ahail elit sedasuil tinciduntas velanani ',
    { x: 7.537, y: 4.516, w: 3.416, h: 1.133 });
  ['1. Your Name', '2. Your Name', '3. Your Name', '4. Your Name'].forEach(function (t, i) {
    text(s, t, { x: 1.729 + i * 2.5785, y: 0.829, w: 2.14, h: 0.303, fontFace: HEAD, bold: true, fontSize: 12, align: 'center' });
  });
  hline(s, 1.729, 6.739, 9.875);
});

// 21 - profile with a vertical meter.
slides.push(function (s) {
  diagWash(s, 0, 0, 3.429, 3.406, 0.57246, [-0.318, -0.709, 2.483, 2.505]);
  cornerBR(s);
  photo(s, 6.815, 1.663, 3.365, 4.657);

  heading(s, ['Your Name'], { x: 1.969, y: 2.049, w: 3.857, h: 0.673 });
  copy(s, [
    'PLACEHOLDER',
    'tincidunt velas aliseta nibuhase ',
  ], { x: 2.001, y: 3.164, w: 3.736, h: 1.398 });
  pill(s, 1.843, 5.318, 3.981, { dotDx: 0.263, textDx: 0.682, textW: 3.321 });
  // the meter group is rotated a quarter turn onto its side
  s.addShape('roundRect', { x: 11.104, y: 1.693, w: 0.161, h: 3.656, rectRadius: 0, fill: { color: RULE } });
  s.addShape('roundRect', { x: 11.104, y: 1.693, w: 0.161, h: 3.155, rectRadius: 0, fill: { color: INK } });
  copy(s, 'lorem ipsum dolor', { x: 10.511, y: 2.553, w: 2.131, h: 0.304, fontSize: 9, rotate: 90 });
  text(s, [{ text: '81', options: { fontSize: 10.5 } }, { text: '%', options: { fontSize: 9 } }],
    { x: 11.26, y: 4.925, w: 0.64, h: 0.338, align: 'right', rotate: 90, lineSpacingMultiple: 1.5 });
  burger(s, 12.443, 0.657);
});

// 22 - desktop mock-up, charcoal copy panel.
slides.push(function (s) {
  imac(s, 1.571, 0.365, 5.508, 5.508);
  box(s, 8.653, 0, 4.681, H, INK);

  pill(s, 1.131, 6.136, 6.401);
  stats(s, 5.987, 6.269);
  burger(s, 12.443, 0.657);
  heading(s, ['About ', 'Our ', 'Mockup'], { x: 9.768, y: 1.337, w: 2.797, h: 2.322, fontSize: 44, color: PAPER });
  copy(s, LOREM_SUBA, { x: 9.768, y: 3.803, w: 2.797, h: 0.338, align: 'left' });
  copy(s, LOREM_ABOUT, { x: 9.763, y: 5.007, w: 2.428, h: 1.398 });
});

// 23 - two phone mock-ups.
slides.push(function (s) {
  diagWash(s, 0, 0, 5.724, 5.773, 0.68556, [-0.115, -0.287, 3.895, 3.929]);
  cornerBR(s);
  phone(s, 2.523, 1.264, 2.547, 5.139);
  phone(s, 5.425, 1.264, 2.547, 5.139);

  sideTab(s, 1.568, 1.34, 4.488, { foot: false, capDy: 0.249 });
  heading(s, ['Mockup'], { x: 8.86, y: 1.693, w: 3.185, h: 0.673 });
  copy(s, [
    'PLACEHOLDER',
    '',
    'PLACEHOLDER',
  ], { x: 8.866, y: 2.77, w: 3.111, h: 2.193 });
  text(s, '1. First Mockup', { x: 8.86, y: 5.526, w: 1.404, h: 0.286, fontFace: HEAD, bold: true, fontSize: 11 });
  text(s, '2. Second Mockup', { x: 10.137, y: 5.526, w: 1.84, h: 0.286, fontFace: HEAD, bold: true, fontSize: 11, align: 'right' });
});

// 24 - laptop mock-up on a mist field.
slides.push(function (s) {
  vline(s, 10.671);
  box(s, 0, 0, 7.651, H, MIST);
  laptop(s, 5.491, 1.181, 6.477, 3.78);

  pill(s, 1.188, 4.467, 3.108, { label: 'Something Here ', dotDx: 0.298, textDx: 0.638, textW: 2.307 });
  meter(s, 1.471, 6.303, 3.293, { fill: 2.842 });
  copy(s,
    'lorem ipsum dolor sita amet consecanila sunian teturasi adipiscinges hanina nibuhase eget justo sedaluin gravida estasunisa sedisail velait asain   nibhasi elitas sedas ',
    { x: 1.385, y: 2.586, w: 3.5, h: 1.133 });
  heading(s, ['Mockup'], { x: 1.385, y: 1.558, w: 3.185, h: 0.673 });
  burger(s, 0.624, 0.657);
});

// 25 - tablet mock-up between two copy columns.
slides.push(function (s) {
  slantWash(s, 0.76, 11.289, 7.405, [5.16, -0.118, 7.668, 7.735]);
  cornerTL(s);
  tablet(s, 4.878, 1.111, 3.578, 5.278);

  pill(s, 9.299, 5.453, 3.108, { label: 'Something Here ', dotDx: 0.298, textDx: 0.638, textW: 2.307 });
  heading(s, ['Mockup'], { x: 1.066, y: 1.589, w: 3.185, h: 0.673 });
  copy(s,
    'PLACEHOLDER',
    { x: 1.072, y: 2.579, w: 2.815, h: 1.398 });
  feature(s, 9.471, 1.729, '1. First Mockup', MOCK_BLURB, 2.914, 0.603);
  feature(s, 9.471, 3.551, '2. Second Mockup', MOCK_BLURB, 2.914, 0.603);
});

// 26 - 100% stacked column chart + charcoal legend panel.
slides.push(function (s) {
  diagWash(s, 0, 0, 5.724, 5.773, 0.68556, [-0.115, -0.287, 3.895, 3.929]);

  s.addChart('bar', SER3, Object.assign({}, AXIS, {
    x: 1.107, y: 1.0, w: 4.143, h: 5.627,
    barDir: 'col', barGrouping: 'percentStacked', barGapWidthPct: 150,
    valAxisLabelFormatCode: '0%',
    chartColors: [CHART_GREYS.light, CHART_GREYS.mid, CHART_GREYS.dark],
  }));
  box(s, 5.693, 1.157, 3.178, 5.144, INK);
  burger(s, 12.443, 0.657);
  heading(s, ['About ', 'Our ', 'Chart'], { x: 9.769, y: 1.452, w: 2.797, h: 2.121, fontSize: 40 });
  copy(s, LOREM_SUBA, { x: 9.769, y: 3.749, w: 2.797, h: 0.338, align: 'left' });
  copy(s, LOREM_ABOUT, { x: 9.765, y: 4.643, w: 2.428, h: 1.398 });
  [1.778, 3.236, 4.694].forEach(function (y) {
    text(s, '1. First Chart', { x: 6.37, y, w: 1.823, h: 0.303, fontFace: HEAD, bold: true, fontSize: 12, color: PAPER });
    copy(s, CHART_BLURB, { x: 6.371, y: y + 0.423, w: 1.823, h: 0.603 });
  });
  poly(s, [[13.333, 6.627], [13.333, H], [12.439, H]], INK);
  s.addShape('line', { x: 11.858, y: 6.186, w: 1.599, h: 1.613, flipV: true, line: { color: RULE, width: 1.5 } });
});

// 27 - 100% stacked bar chart across the top.
slides.push(function (s) {
  slantWash(s, 0.688, 11.289, 7.405, [5.088, -0.118, 7.668, 7.735]);

  s.addChart('bar', SER4, Object.assign({}, AXIS, {
    x: 1.689, y: 1.009, w: 8.889, h: 3.755,
    barDir: 'bar', barGrouping: 'percentStacked', barGapWidthPct: 150,
    valAxisLabelFormatCode: '0%',
    chartColors: [CHART_GREYS.dark, CHART_GREYS.light, CHART_GREYS.mid],
  }));
  sideTab(s, 11.039, 1.42, 4.673);
  copy(s,
    'PLACEHOLDER',
    { x: 5.103, y: 5.313, w: 5.288, h: 1.133 });
  heading(s, ['Our Chart'], { x: 1.701, y: 5.493, w: 2.797, h: 0.774, fontSize: 40 });
});

// 28 - clustered bar chart, charcoal copy panel left.
slides.push(function (s) {
  box(s, 0, 0, 4.681, H, INK);

  s.addChart('bar', SER4, Object.assign({}, AXIS, {
    x: 5.864, y: 1.347, w: 6.401, h: 4.139,
    barDir: 'bar', barGrouping: 'clustered', barGapWidthPct: 182, barOverlapPct: 0,
    chartColors: [CHART_GREYS.light, CHART_GREYS.mid, CHART_GREYS.dark],
  }));
  pill(s, 5.817, 6.09, 6.401);
  stats(s, 10.673, 6.224);
  burger(s, 0.624, 0.657);
  heading(s, ['About ', 'Our ', 'Chart'], { x: 1.123, y: 1.462, w: 2.797, h: 2.121, fontSize: 40, color: PAPER });
  copy(s, LOREM_SUBA, { x: 1.123, y: 3.801, w: 2.797, h: 0.338, align: 'left' });
  copy(s, LOREM_ABOUT, { x: 1.118, y: 4.759, w: 2.428, h: 1.398 });
  s.addShape('line', { x: 6.081, y: 0.744, w: 7.283, h: 0, flipH: true, line: { color: RULE, width: 1.5 } });
});

// 29 - 100% stacked area chart.
slides.push(function (s) {
  vline(s, 8.239);
  box(s, 0, 0, 7.361, H, MIST);

  // passed as a chart-type array so pptxgenjs emits crossBetween="midCat",
  // i.e. the area meets the left and right edges of the plot as in the source
  s.addChart([{ type: 'area', data: AREA_SER }], Object.assign({}, AXIS, {
    x: 5.236, y: 1.065, w: 6.107, h: 4.324,
    barGrouping: 'stacked',
    catLabelFormatCode: 'm/d/yyyy', catAxisLabelPos: 'nextTo', catAxisMajorTickMark: 'out',
    valAxisLabelFormatCode: '0%', valAxisMinVal: 0, valAxisMaxVal: 1,
    chartColors: [CHART_GREYS.light, CHART_GREYS.deep],
  }));
  box(s, 11.393, 1.214, 0.606, 3.883, INK);
  text(s, 'MARCH - 35', {
    x: 11.525, y: 1.451, w: 0.436, h: 2.573, vert: 'wordArtVert', align: 'center',
    fontSize: 7, charSpacing: 6, color: PAPER, lineSpacingMultiple: 1.5,
  });
  burger(s, 11.566, 5.636);
  pill(s, 1.017, 4.606, 3.285, { label: 'Something Here ', dotDx: 0.299, textDx: 0.677, textW: 2.307 });
  heading(s, ['About Our Chart'], { x: 5.768, y: 5.963, w: 5.26, h: 0.64, fontSize: 32, align: 'center' });
  feature(s, 1.222, 1.284, '1. First Chart', MOCK_BLURB, 2.914, 0.603);
  feature(s, 1.222, 2.923, '2. Second Chart', MOCK_BLURB, 2.914, 0.603);
  copy(s,
    'PLACEHOLDER',
    { x: 1.191, y: 5.982, w: 4.539, h: 0.603 });
});

// 30 - closing slide, mirrors slide 1.
slides.push(function (s) {
  box(s, 10.75, 0, 2.583, H, MIST);
  vline(s, 5.177);
  box(s, 0, 0, 4.681, H, INK);
  photo(s, 3.444, 1.052, 7.305, 5.397);

  box(s, 10.75, 1.052, 0.606, 4.791, INK);
  box(s, 10.75, 5.842, 0.606, 0.606, PAPER);
  text(s, 'PRESENTATION', {
    x: 10.874, y: 1.052, w: 0.436, h: 4.759, vert: 'wordArtVert', align: 'center',
    fontSize: 7, charSpacing: 6, color: PAPER, lineSpacingMultiple: 1.5,
  });
  burger(s, 10.925, 6.079);

  text(s, [
    { text: 'TH', options: { color: PAPER } },
    { text: 'ANK.', options: { color: INK } },
  ], { x: 1.474, y: 1.953, w: 7.79, h: 2.036, fontFace: HEAD, bold: true, fontSize: 115, charSpacing: 6 });
  // charcoal stem of the "A" continuing across the panel edge
  box(s, 3.438, 2.426, 0.218, 1.152, INK);
  copy(s, LOREM_TITLE, { x: 1.577, y: 3.979, w: 3.983, h: 0.868 });
});

/* ------------------------------------------------------------------ build */

const pres = new PptxGenJS();
pres.defineLayout({ name: 'DECK', width: W, height: H });
pres.layout = 'DECK';
pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };

slides.forEach(function (build) {
  const s = pres.addSlide();
  s.background = { color: PAPER };
  build(s);
});

pres.writeFile({ fileName: path.join(__dirname, '138b83a3-bf31-4ba6-bb91-12c8647fe285_grok_final.pptx') })
  .then(function (f) { console.log('wrote ' + f); });
