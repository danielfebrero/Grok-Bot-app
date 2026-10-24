/**
 * "Stress Level" infographic deck - rebuilt with pptxgenjs.
 *
 * Slide size 26.67 x 15 in (24387175 x 13716000 EMU).
 * Raster artwork of the original deck is replaced by flat "[image]" placeholders.
 *
 * Run: node 0e8a3042-cc8c-4f61-8502-07cf2852f40c_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

const W = 26.67; // slide width  (in)
const H = 15.0;  // slide height (in)

/* ------------------------------------------------------------------ palette */
// Theme "PC - Color 10 Red": dk1 = dk2 = 44546A, accents are a red ramp.
const C = {
  text: '44546A',      // dk1/dk2 - default body + heading colour
  gray: '656D78',      // explicit grey used in the card/list paragraphs
  white: 'FFFFFF',
  a1: 'EF5350', a2: 'F44336', a3: 'E53935', a4: 'D32F2F', a5: 'C62828', a6: 'B71C1C',
  // The deck fills badges/circles with a radial accent gradient (lumMod 89% -> 70%);
  // these are its flat equivalents, measured from the reference render.
  g1: 'EA1A16', g2: 'E21B0C', g3: 'C81E1A', g4: 'AA2424', g5: '9E2020', g6: '921616',
  // Theme fill style 3 (tint 94% -> shade 78%) used by the icon discs.
  t1: 'D7514F', t2: 'DE453B', t3: 'CD3D3A', t4: 'C53939', t5: 'B73333', t6: 'C23A3A',
  ringOut: 'D1D4D5', ringIn: 'D1D4D4', ringBack: 'EDEDED', dial: 'F7F5F6',
  pill: 'D8DADD',
  art: 'F0F4FA', artLine: 'DCE6F5', artText: 'A9B8D0',
};
const ACCENTS = [C.a1, C.a2, C.a3, C.a4, C.a5, C.a6];
const GRADS = [C.g1, C.g2, C.g3, C.g4, C.g5, C.g6];
const THEME = [C.t1, C.t2, C.t3, C.t4, C.t5, C.t6];

/* -------------------------------------------------------------------- fonts */
const F = {
  light: 'Source Sans Pro Light',
  xlight: 'Source Sans Pro ExtraLight',
  black: 'Source Sans Pro Black',
  reg: 'Source Sans Pro',
  roboto: 'Roboto',
  robotoBlack: 'Roboto Black',
};

/* ----------------------------------------------------------- shared strings */
const T = {
  lorem: 'Hashtag fashion axe fingerstache, everyday carry shoreditch pinterest umami authentic brooklyn ' +
    'YOLO heirloom keytar waistcoat kickstarter. Kitsch authentic offal, narwhal tilde etsy four loko ' +
    'selvage normcore messenger bag put a bird on it heirloom gastropub.',
  loremCut: 'Hashtag fashion axe fingerstache, everyday carry shoreditch pinterest umami authentic ' +
    'brooklyn YOLO heirloom keytar waistcoat kickstarter. ',
  loremShort: 'Hashtag fashion axe fingerstache, everyday carry shoreditch pinterest umami authentic brooklyn YOLO.',
  loremMid: 'Hashtag fashion axe fingerstache, everyday carry shoreditch brooklyn YOLO.',
  loremMini: 'Hashtag fashion axe fingerstache brooklyn YOLO.',
  kitsch: 'Kitsch authentic offal, narwhal tilde etsy four loko selvage normcore messenger.',
  synth: 'Synth chartreuse XOXO, tacos brooklyn VHS plaid.',
  synthShort: 'Synth chartreuse XOXO, tacos VHS plaid.',
  bicycle: 'Bicycle rights +1 actually shoreditch, vinyl fixie small batch pop-up',
  deckTitle: 'STRESS LEVEL',
};

/* ------------------------------------------------------------------ helpers */

// Rich text: [ ['some words'], ['bold red words', C.a1, true] ] -> pptxgenjs runs
const runs = (parts) => parts.map(([text, color, bold]) =>
  ({ text, options: { color: color || C.text, bold: !!bold } }));

// Concentric box centred inside (x,y,d,d), scaled by k.
const ring = (x, y, d, k) => ({ x: x + d * (1 - k) / 2, y: y + d * (1 - k) / 2, w: d * k, h: d * k });

// Page furniture shared by every content slide (2..16).
function header(s) {
  s.addText(runs([
    ['Stress Level '], ['Infographic', C.a3], [' ', C.a1], ['Slide'],
  ]), { x: 5.074, y: 0.893, w: 16.522, h: 1.178, align: 'center', valign: 'top', fontFace: F.light, fontSize: 64 });

  s.addText(runs([
    ['Letterpress next level trust fund, '], ['before', C.a3],
    [' they sold out +1 meh gluten-free locavore tacos PBR&B tofu. '],
  ]), {
    x: 6.445, y: 2.054, w: 13.779, h: 0.505, align: 'center', valign: 'top', wrap: false,
    fontFace: F.light, fontSize: 24,
  });
}

// Stand-in for one of the deck's flat vector illustrations.
function artwork(s, x, y, w, h) {
  s.addShape('roundRect', {
    x, y, w, h, rectRadius: 0.35,
    fill: { color: C.art }, line: { color: C.artLine, width: 1 },
  });
  s.addText('[image]', {
    x, y, w, h, align: 'center', valign: 'middle',
    fontFace: F.light, fontSize: 28, color: C.artText,
  });
}

// Big quote block: 48pt Roboto with accent-coloured emphasis.
function quote(s, x, y, w, h, parts) {
  s.addText(runs(parts), { x, y, w, h, valign: 'top', fontFace: F.roboto, fontSize: 48 });
}

// "Bicycle rights ..." strap line under a quote.
function strap(s, x, y, color) {
  s.addText(T.bicycle, {
    x, y, w: 9.436, h: 1.043, valign: 'top', fontFace: F.xlight, fontSize: 28, color,
  });
}

// 21pt caption line ("Synth chartreuse ...").
function caption(s, x, y) {
  s.addText(T.synth, {
    x, y, w: 6.099, h: 0.522, valign: 'top', wrap: false,
    fontFace: F.light, fontSize: 21, color: C.text, lineSpacingMultiple: 1.3,
  });
}

// The recurring 24pt lorem paragraph.
function body(s, x, y, w, h, size) {
  s.addText(T.lorem, {
    x, y, w, h, valign: 'top', fontFace: F.light, fontSize: size || 24,
    color: C.text, lineSpacingMultiple: 1.5,
  });
}

/* --------------------------------------------------------------- widgets    */

// Flat percentage doughnut: grey track + native pie chart arc + white centre.
function donutStat(s, x, y, d, value, color) {
  s.addShape('donut', { ...ring(x, y, d, 1.00), fill: { color: C.ringBack } });
  s.addShape('donut', { ...ring(x, y, d, 0.905), fill: { color: C.ringOut } });
  s.addChart('pie', [{ name: 'Region 1', labels: ['Quarter', 'Blank'], values: [value, 130] }], {
    ...ring(x, y, d, 0.924),
    chartColors: [color, C.ringOut],
    layout: { x: 0.005, y: 0.005, w: 0.99, h: 0.9875 },
    showLegend: false, showTitle: false, dataNoEffects: true,
  });
  s.addShape('ellipse', { ...ring(x, y, d, 0.547), fill: { color: C.white } });
}

// Speedometer-style gauge (slide 2): nested rings around the same pie arc.
function gaugeRing(s, x, y, d, value, color) {
  s.addShape('ellipse', { ...ring(x, y, d, 0.896), fill: { color: C.ringBack } });
  s.addShape('donut', { ...ring(x, y, d, 1.00), fill: { color: C.ringOut } });
  s.addChart('pie', [{ name: 'Region 1', labels: ['Quarter', 'Blank'], values: [value, 130] }], {
    ...ring(x, y, d, 0.858),
    chartColors: [color, C.ringBack],
    layout: { x: 0.005, y: 0.005, w: 0.99, h: 0.9875 },
    showLegend: false, showTitle: false, dataNoEffects: true,
  });
  s.addShape('ellipse', { ...ring(x, y, d, 0.557), fill: { color: C.dial } });
  s.addShape('donut', { ...ring(x, y, d, 0.580), fill: { color: C.ringIn } });
  s.addShape('ellipse', { ...ring(x, y, d, 0.480), fill: { color: C.white } });
  s.addShape('ellipse', { ...ring(x, y, d, 0.120), fill: { color: C.ringIn } });
}

// "% + Type X" caption pair used next to donutStat.
function statLabel(s, x, y, pct, label, color) {
  s.addText(pct, { x, y, w: 0.985, h: 0.573, margin: 0, valign: 'top', fontFace: F.roboto, fontSize: 32, bold: true, color });
  s.addText(label, { x, y: y + 0.598, w: 0.985, h: 0.452, margin: 0, valign: 'top', fontFace: F.roboto, fontSize: 20, bold: true, color: C.text });
}

// Thin ring + tick glyph, drawn in `color` on a transparent background.
// `open` leaves a gap in the ring where the tick crosses it (badge variant).
function tickGlyph(s, x, y, d, color, weight, open) {
  const lw = weight || 1;
  s.addShape(open ? 'arc' : 'ellipse', {
    x, y, w: d, h: d, angleRange: [320, 260],
    fill: { type: 'none' }, line: { color, width: lw },
  });
  s.addShape('custGeom', {
    x: x + d * 0.24, y: y + d * 0.30, w: d * 0.55, h: d * 0.42,
    points: [{ x: 0, y: d * 0.22 }, { x: d * 0.20, y: d * 0.42 }, { x: d * 0.55, y: 0 }],
    fill: { type: 'none' }, line: { color, width: lw },
  });
}

// Solid accent disc with a white tick glyph (slides 4 and 6).
function tickBadge(s, x, y, d, fillColor) {
  s.addShape('ellipse', { x, y, w: d, h: d, fill: { color: fillColor } });
  tickGlyph(s, x + d * 0.25, y + d * 0.25, d * 0.50, C.white, 1.5, true);
}

// Concentric "target" icon used on slides 12 and 13.
function targetIcon(s, x, y, d, accent, grad, theme) {
  s.addShape('ellipse', { ...ring(x, y, d, 1.00), fill: { color: accent } });
  s.addShape('ellipse', { ...ring(x, y, d, 0.85), fill: { color: grad } });
  s.addShape('ellipse', { ...ring(x, y, d, 0.76), fill: { color: C.white } });
  s.addShape('ellipse', { ...ring(x, y, d, 0.56), fill: { color: theme } });
  s.addShape('ellipse', { ...ring(x, y, d, 0.47), fill: { color: C.white } });
  tickGlyph(s, x + d * 0.37, y + d * 0.37, d * 0.26, grad, 0.75, true);
}

/* ------------------------------------------------- title / closing slide art */

// The cover uses a radial red gradient; approximate it with concentric ellipses.
function redBackdrop(s) {
  const edge = [0xAF, 0x1A, 0x17];
  const core = [0xDE, 0x21, 0x1D];
  const steps = 12;
  s.addShape('rect', { x: 0, y: 0, w: W, h: H, fill: { color: 'AF1A17' } });
  for (let i = steps; i >= 1; i--) {
    const t = i / steps;
    const hex = edge.map((e, k) => Math.round(e + (core[k] - e) * (1 - t)))
      .map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
    const w = W * 1.5 * t;
    const h = H * 1.5 * t;
    s.addShape('ellipse', { x: (W - w) / 2, y: (H - h) / 2, w, h, fill: { color: hex }, line: { type: 'none' } });
  }
}

// Five near-transparent white sweeps layered over the cover background.
const SWIRLS = [
  { box: [7.068, -0.167, 19.601, 15.333], alpha: 95, pts: [[0.9999, 1.0], ['c', 0.748, 0.321, 0.192, 0.570, 0, 0], [1, 0], [1, 1], [0.9999, 1], ['z']] },
  { box: [14.094, -0.167, 12.574, 11.655], alpha: 93, pts: [[0, 0], [1, 0], [1, 0.216], [1, 1], ['c', 0.782, 0.279, 0.146, 0.613, 0, 0]] },
  { box: [0.000, 5.384, 15.281, 9.783], alpha: 95, pts: [[0, 0], ['c', 0.270, 0.415, 0.736, 0.323, 1, 1], [0, 1], [0, 0], ['z']] },
  { box: [0.000, 9.684, 11.000, 5.482], alpha: 95, pts: [[1, 1], [0, 1], [0, 0], ['c', 0.302, 0.360, 0.694, 0.451, 1, 1]] },
  { box: [12.254, -0.167, 14.414, 15.333], alpha: 95, pts: [[1, 1], [0.810, 1], ['c', 0.139, 0.828, 0.383, 0.289, 0, 0], [1, 0], [1, 1], ['z']] },
];

function swirls(s, dy) {
  SWIRLS.forEach(({ box, alpha, pts }) => {
    const [x, y, w, h] = box;
    const points = pts.map((p) => {
      if (p[0] === 'z') return { close: true };
      if (p[0] === 'c') {
        return { x: p[5] * w, y: p[6] * h, curve: { type: 'cubic', x1: p[1] * w, y1: p[2] * h, x2: p[3] * w, y2: p[4] * h } };
      }
      return { x: p[0] * w, y: p[1] * h };
    });
    s.addShape('custGeom', {
      x, y: y + dy, w, h, points,
      fill: { color: C.white, transparency: alpha }, line: { type: 'none' },
    });
  });
}

function coverSlide(s, dy, topText, bottomText, withArrow) {
  redBackdrop(s);
  swirls(s, dy);
  s.addText(topText, {
    x: 5.663, y: withArrow ? 4.973 : 4.665, w: 15.344, h: withArrow ? 0.586 : 0.767,
    align: 'center', valign: 'top', color: C.white, lineSpacingMultiple: 0.9,
    fontFace: F.reg, fontSize: withArrow ? 32 : 44,
  });
  s.addText(T.deckTitle, {
    x: 7.025, y: 6.330, w: 12.621, h: 2.339, align: 'center', valign: 'top', wrap: false,
    fontFace: F.black, fontSize: 133, bold: true, color: C.white,
  });
  s.addText(bottomText.text, {
    x: bottomText.x, y: bottomText.y, w: bottomText.w, h: 0.505, align: 'center', valign: 'top', wrap: false,
    fontFace: F.light, fontSize: 24, color: C.white,
  });
  if (withArrow) {
    s.addShape('ellipse', { x: 12.867, y: 11.393, w: 0.936, h: 0.936, fill: { type: 'none' }, line: { color: C.white, width: 2.5 } });
    s.addShape('line', {
      x: 13.335, y: 11.60, w: 0, h: 0.50,
      line: { color: C.white, width: 2.5, endArrowType: 'triangle' },
    });
  }
}

/* ------------------------------------------------------------------- slides */

function slide01(s) {
  const italic = (text, bold) => ({ text, options: { color: C.white, italic: true, bold: !!bold } });
  const plain = (text, bold) => ({ text, options: { color: C.white, bold: !!bold } });
  coverSlide(s, 0, [
    plain('We\u2019ve been '), plain('crafting', true), plain(' beautiful '),
    italic('presentation'), plain(' & making '), italic('clients', true), plain(' happy for years.'),
  ], { text: 'Letterpress next level trust fund, before. ', x: 10.517, y: 10.222, w: 5.636 }, true);
}

function slide02(s) {
  header(s);
  artwork(s, 0.651, 4.148, 12.179, 8.443);
  quote(s, 13.425, 3.769, 11.058, 3.332, [
    ['In times of '], ['great stress ', C.a1, true], ["or adversity, it's always best to keep busy, to plow "],
    ['your anger ', C.a2, true], ['and '], ['your energy ', C.a3, true], ['into', C.a4, true],
    [' something '], ['positive', C.a5, true], ['.'],
  ]);
  strap(s, 13.425, 7.106, C.a4);
  caption(s, 13.425, 8.108);
  body(s, 13.425, 8.590, 10.640, 2.461);
  [[13.425, 25, C.a1], [16.099, 75, C.a2], [18.774, 19, C.a3], [21.448, 350, C.a4]]
    .forEach(([x, v, col]) => gaugeRing(s, x, 11.437, 2.244, v, col));
}

function slide03(s) {
  header(s);
  quote(s, 2.149, 4.180, 11.224, 2.524, [
    ['Solving '], ['specific', C.a1, true], [' problems is what drives me. I am not interested in '],
    ['having', C.a2, true], [' a career. I '], ['never', C.a3, true], [' have been.'],
  ]);
  strap(s, 2.149, 6.668, C.a3);
  caption(s, 2.149, 7.675);
  body(s, 2.149, 8.162, 10.640, 2.461);
  donutStat(s, 2.149, 11.028, 1.693, 30, C.a1);
  donutStat(s, 5.032, 11.041, 1.689, 130, C.a2);
  donutStat(s, 8.336, 11.041, 1.689, 365, C.a3);
  statLabel(s, 3.944, 11.377, '19%', 'Type A', C.a1);
  statLabel(s, 6.930, 11.384, '50%', 'Type B', C.a2);
  statLabel(s, 10.127, 11.382, '74%', 'Type C', C.a3);
  artwork(s, 13.979, 4.950, 11.551, 7.780);
}

function slide04(s) {
  header(s);
  artwork(s, 0.930, 3.832, 10.602, 8.363);
  const cards = [
    { x: 12.712, y: 4.398, tx: 14.353, ty: 4.751, n: '01', accent: C.a1, grad: C.g1 },
    { x: 19.757, y: 4.398, tx: 21.398, ty: 4.776, n: '02', accent: C.a2, grad: C.g2 },
    { x: 12.712, y: 9.027, tx: 14.353, ty: 9.417, n: '03', accent: C.a4, grad: C.g4 },
    { x: 19.757, y: 9.027, tx: 21.398, ty: 9.428, n: '04', accent: C.a3, grad: C.g3 },
  ];
  cards.forEach((c) => {
    tickBadge(s, c.x, c.y, 1.311, c.grad);
    s.addText(`Analysis ${c.n}`, {
      x: c.tx, y: c.ty, w: 3.868, h: 0.558, margin: 0, valign: 'middle',
      fontFace: F.robotoBlack, fontSize: 28, bold: true, color: C.text, lineSpacingMultiple: 1.3,
    });
    s.addText([
      { text: T.synth, options: { fontSize: 22, color: c.accent, breakLine: true } },
      { text: T.loremShort, options: { fontSize: 20, color: C.gray } },
    ], {
      x: c.x, y: c.y + 1.710, w: 6.175, h: 1.732, margin: 0, valign: 'top',
      fontFace: F.light, lineSpacingMultiple: 1.5,
    });
  });
}

function slide05(s) {
  header(s);
  [[4.119, 3.423, 3.958], [5.970, 3.505, 3.654], [7.820, 3.420, 3.888],
    [9.670, 3.505, 3.621], [11.520, 3.505, 3.621]].forEach(([y, tx, tw], i) => {
    s.addShape('ellipse', { x: 2.289, y: y + 0.440, w: 1.024, h: 1.024, fill: { color: THEME[i] } });
    s.addShape('ellipse', { x: 2.371, y: y + 0.525, w: 0.858, h: 0.858, fill: { color: C.white } });
    tickGlyph(s, 2.585, y + 0.740, 0.430, ACCENTS[i], 0.75, true);
    s.addText([
      { text: T.synth, options: { fontFace: F.black, bold: true, breakLine: true } },
      { text: T.loremMini, options: { fontFace: F.light } },
    ], {
      x: tx, y, w: tw, h: 1.674, valign: 'top', fontSize: 16, color: C.text, lineSpacingMultiple: 1.5,
    });
  });
  s.addText('Stress Level Analysis', {
    x: 7.724, y: 4.124, w: 5.762, h: 0.871, margin: 0, valign: 'top',
    fontFace: F.robotoBlack, fontSize: 28, bold: true, color: C.text,
  });
  s.addText('Process Stress Level Information Analysis', {
    x: 7.724, y: 4.803, w: 5.399, h: 1.453, margin: 0, valign: 'top',
    fontFace: F.robotoBlack, fontSize: 28, bold: true, color: C.g5, lineSpacingMultiple: 1.2,
  });
  s.addText([
    { text: T.lorem, options: { breakLine: true } },
    { text: T.kitsch },
  ], {
    x: 7.724, y: 6.064, w: 4.919, h: 7.913, valign: 'top',
    fontFace: F.light, fontSize: 24, color: C.text, lineSpacingMultiple: 1.5,
  });
  artwork(s, 13.758, 4.391, 11.916, 7.798);
}

function slide06(s) {
  header(s);
  artwork(s, 0.588, 3.446, 12.640, 9.019);
  quote(s, 13.880, 3.752, 11.794, 1.717, [
    ['Stress', C.a1, true], [' is caused by being '], ["'here' ", C.a2, true],
    ['but wanting to '], ["be 'there.'", C.a3, true],
  ]);
  strap(s, 13.850, 5.533, C.a3);
  s.addText(T.lorem, {
    x: 13.850, y: 6.482, w: 12.232, h: 1.709, valign: 'top',
    fontFace: F.light, fontSize: 22, color: C.gray, lineSpacingMultiple: 1.5,
  });
  [[8.682, 8.718, 9.173, C.a1, C.g1], [11.078, 11.114, 11.570, C.a2, C.g2]].forEach(([ty, cy, by, accent, grad]) => {
    tickBadge(s, 13.993, cy, 1.496, grad);
    s.addText(T.synth, {
      x: 15.893, y: ty, w: 6.227, h: 0.467, margin: 0, valign: 'top',
      fontFace: F.light, fontSize: 21, color: accent, lineSpacingMultiple: 1.3,
    });
    s.addText(T.loremMid, {
      x: 15.914, y: by, w: 8.685, h: 1.077, margin: 0, valign: 'top',
      fontFace: F.light, fontSize: 24, color: C.gray, lineSpacingMultiple: 1.5,
    });
  });
}

function slide07(s) {
  header(s);
  quote(s, 1.698, 3.803, 10.416, 3.332, [
    ['Much of '], ['the stress ', C.a1, true], ["that people feel doesn't come from "],
    ['having', C.a2, true], [' too much to do. It comes from not '], ['finishing', C.a3, true],
    [' what '], ["they've started", C.a4, true], ['.'],
  ]);
  strap(s, 1.698, 7.135, C.a4);
  caption(s, 1.698, 8.178);
  body(s, 1.698, 8.699, 10.640, 2.461);
  [11.524, 12.521].forEach((y) => {
    tickGlyph(s, 1.721, y, 0.604, C.a4, 1.375);
    s.addText(T.synth, {
      x: 2.471, y: y + 0.011, w: 6.944, h: 0.583, valign: 'top', wrap: false,
      fontFace: F.light, fontSize: 24, color: C.text, lineSpacingMultiple: 1.3,
    });
  });
  artwork(s, 12.261, 3.497, 13.850, 9.883);
}

function slide08(s) {
  header(s);
  artwork(s, 0.720, 3.722, 12.286, 8.926);
  quote(s, 13.813, 3.959, 11.058, 3.332, [
    ['Stress', C.a1, true], [' is the trash of '], ['modern', C.a2, true], [' life \u2014 we all '],
    ['generate', C.a3, true], [" it, but if you don't dispose of it "], ['properly', C.a4, true],
    [', it will pile up and '], ['overtake', C.a5, true], [' your life.'],
  ]);
  strap(s, 13.813, 7.083, C.a4);
  caption(s, 13.813, 8.116);
  body(s, 13.813, 8.628, 10.640, 2.461);
  donutStat(s, 13.877, 11.383, 1.693, 30, C.a1);
  donutStat(s, 16.760, 11.396, 1.689, 130, C.a2);
  donutStat(s, 20.063, 11.396, 1.689, 365, C.a3);
  statLabel(s, 15.672, 11.732, '19%', 'Type A', C.a1);
  statLabel(s, 18.658, 11.739, '50%', 'Type B', C.a2);
  statLabel(s, 21.855, 11.737, '74%', 'Type C', C.a3);
}

function slide09(s) {
  header(s);
  const blocks = [
    { x: 1.818, y: 3.901, tx: 2.394, title: 'Identify Stressors', accent: C.a1 },
    { x: 8.350, y: 3.901, tx: 8.836, title: 'Stress Awareness', accent: C.a2 },
    { x: 1.818, y: 8.625, tx: 2.387, title: 'Stress Reduction', accent: C.a3 },
    { x: 8.350, y: 8.625, tx: 8.886, title: 'Problem Solving', accent: C.a4 },
  ];
  blocks.forEach((b) => {
    s.addShape('ellipse', { x: b.x, y: b.y, w: 1.181, h: 1.183, fill: { color: b.accent, transparency: 71 } });
    s.addShape('ellipse', { x: b.x + 0.103, y: b.y + 0.112, w: 0.943, h: 0.945, fill: { color: b.accent, transparency: 71 } });
    s.addShape('ellipse', { x: b.x + 0.238, y: b.y + 0.250, w: 0.668, h: 0.669, fill: { color: b.accent } });
    s.addText(b.title, {
      x: b.tx, y: b.y + 0.488, w: 4.55, h: 0.774, margin: [0, 0, 3.6, 3.6], valign: 'top', wrap: false,
      fontFace: F.robotoBlack, fontSize: 40, bold: true, color: C.text,
    });
    s.addText([
      { text: T.synthShort, options: { color: b.accent, breakLine: true } },
      { text: T.loremShort, options: { color: C.gray } },
    ], {
      x: b.tx, y: b.y + 1.633, w: 4.55, h: 2.067, margin: [0, 0, 3.6, 3.6], valign: 'top',
      fontFace: F.light, fontSize: 20, lineSpacingMultiple: 1.5,
    });
  });
  artwork(s, 14.079, 4.261, 11.699, 8.346);
}

function slide10(s) {
  header(s);
  artwork(s, 0.423, 3.791, 12.429, 8.785);
  s.addText([
    { text: 'Stress Level', options: { fontSize: 28, color: C.text, breakLine: true } },
    { text: 'Info', options: { fontSize: 28, color: C.a3, breakLine: true } },
    { text: ' ', options: { fontSize: 8, breakLine: true } },
    { text: 'STRESS ', options: { fontSize: 85.34, color: C.a3 } },
    { text: 'LEVEL', options: { fontSize: 85.34, color: C.text } },
  ], { x: 13.315, y: 4.535, w: 12.443, h: 2.614, valign: 'top', fontFace: F.robotoBlack, bold: true });
  s.addText('CONTENT TITLE', {
    x: 13.315, y: 7.190, w: 4.098, h: 0.550, valign: 'top',
    fontFace: F.robotoBlack, fontSize: 26.66, bold: true, color: C.text,
  });
  s.addText(T.lorem, {
    x: 13.315, y: 7.780, w: 12.232, h: 1.709, valign: 'top',
    fontFace: F.light, fontSize: 22, color: C.gray, lineSpacingMultiple: 1.5,
  });
  donutStat(s, 13.327, 10.139, 1.693, 30, C.a1);
  donutStat(s, 16.210, 10.152, 1.689, 130, C.a2);
  donutStat(s, 19.514, 10.152, 1.689, 365, C.a3);
  donutStat(s, 22.553, 10.106, 1.689, 280, C.a4);
  statLabel(s, 15.122, 10.488, '19%', 'Type A', C.a1);
  statLabel(s, 18.108, 10.494, '50%', 'Type B', C.a2);
  statLabel(s, 21.305, 10.493, '74%', 'Type C', C.a3);
  statLabel(s, 24.414, 10.444, '68%', 'Type D', C.a4);
}

function slide11(s) {
  header(s);
  quote(s, 2.110, 3.721, 10.912, 3.332, [
    ['Stress', C.a1, true], [' is like spice - in the right proportion, it '], ['enhances', C.a2, true],
    [' the flavor of a dish. Too little produces a bland, dull meal; too '], ['much', C.a3, true],
    [' may choke you.'],
  ]);
  s.addText('CONTENT TITLE', {
    x: 2.110, y: 6.992, w: 4.098, h: 0.550, valign: 'top',
    fontFace: F.robotoBlack, fontSize: 26.66, bold: true, color: C.text,
  });
  body(s, 2.110, 7.481, 10.640, 2.461);
  s.addShape('roundRect', { x: 2.115, y: 10.135, w: 8.513, h: 1.570, rectRadius: 0.785, fill: { color: C.pill } });
  s.addText(T.loremShort, {
    x: 2.115, y: 10.135, w: 8.513, h: 1.570, margin: [28.35, 28.35, 3.6, 3.6], valign: 'middle',
    fontFace: F.light, fontSize: 20, color: C.text, lineSpacingMultiple: 1.3,
  });
  [[2.110, '2.888 K', 'Male User', 3.583, 4.030], [6.367, '1.074 K', 'Female User', 7.840, 8.135]]
    .forEach(([x, num, label, nx, lx]) => {
      s.addShape('ellipse', { x, y: 12.210, w: 1.200, h: 1.200, fill: { color: C.t6 } });
      s.addShape('custGeom', {
        x: x + 0.395, y: 12.505, w: 0.410, h: 0.611,
        points: [{ x: 0.205, y: 0.06 }, { x: 0.315, y: 0.17 }, { x: 0.205, y: 0.28 }, { x: 0.095, y: 0.17 }, { close: true },
          { x: 0.02, y: 0.55, moveTo: true }, { x: 0.05, y: 0.34 }, { x: 0.36, y: 0.34 }, { x: 0.39, y: 0.55 }, { close: true }],
        fill: { color: C.white }, line: { type: 'none' },
      });
      s.addText(num, {
        x: nx, y: 12.155, w: 2.151, h: 0.774, valign: 'top', wrap: false,
        fontFace: F.robotoBlack, fontSize: 40, bold: true, color: C.text,
      });
      s.addText(label, {
        x: lx, y: 12.991, w: 1.560, h: 0.337, margin: 0, valign: 'top', wrap: false,
        fontFace: F.roboto, fontSize: 20, color: C.text,
      });
    });
  artwork(s, 13.044, 3.720, 12.603, 8.993);
}

function slide12(s) {
  header(s);
  artwork(s, 0.939, 3.720, 12.259, 8.747);
  const rows = [
    { ix: 13.720, iy: 5.088, tx: 15.272, ty: 4.989, tw: 3.958, accent: C.a4, grad: C.g4, theme: C.t4, bold: true },
    { ix: 19.357, iy: 5.099, tx: 20.940, ty: 4.989, tw: 3.791, accent: C.a3, grad: C.g3, theme: C.t3 },
    { ix: 13.720, iy: 7.630, tx: 15.284, ty: 7.442, tw: 3.888, accent: C.a2, grad: C.g2, theme: C.t2 },
    { ix: 19.357, iy: 7.637, tx: 20.883, ty: 7.442, tw: 3.791, accent: C.a1, grad: C.g1, theme: C.t1 },
  ];
  rows.forEach((r) => {
    targetIcon(s, r.ix, r.iy, 1.457, r.accent, r.grad, r.theme);
    s.addText([
      { text: T.synth, options: { fontFace: F.black, bold: true, breakLine: true } },
      { text: T.loremMini, options: { fontFace: F.light, bold: !!r.bold } },
    ], {
      x: r.tx, y: r.ty, w: r.tw, h: 1.674, valign: 'top',
      fontSize: 16, color: C.text, lineSpacingMultiple: 1.5,
    });
  });
  s.addText('Stress Level Analysis', {
    x: 14.327, y: 10.237, w: 5.762, h: 0.670, margin: 0, valign: 'top',
    fontFace: F.robotoBlack, fontSize: 22, bold: true, color: C.text,
  });
  s.addText('Stress Level Information Analysis', {
    x: 14.327, y: 10.907, w: 3.698, h: 2.150, margin: 0, valign: 'top',
    fontFace: F.robotoBlack, fontSize: 24, bold: true, color: C.a3, lineSpacingMultiple: 1.2,
  });
  s.addText(T.loremCut, {
    x: 18.025, y: 10.062, w: 6.706, h: 1.709, valign: 'top',
    fontFace: F.light, fontSize: 22, color: C.text, lineSpacingMultiple: 1.5,
  });
}

function slide13(s) {
  header(s);
  const items = [
    { ix: 1.085, iy: 4.720, tx: 3.438, ty: 4.966, tw: 4.464, n: '01', body: T.synth },
    { ix: 7.902, iy: 4.720, tx: 10.289, ty: 4.966, tw: 4.476, n: '02', body: T.synth },
    { ix: 1.100, iy: 7.275, tx: 3.438, ty: 7.521, tw: 4.464, n: '03', body: 'Synth chartreuase XOXO, tacos brooklyn VHS plaid.' },
    { ix: 7.966, iy: 7.275, tx: 10.289, ty: 7.521, tw: 4.476, n: '04', body: T.synth },
    { ix: 1.124, iy: 9.956, tx: 3.438, ty: 10.202, tw: 4.464, n: '05', body: T.synth },
    { ix: 7.910, iy: 9.956, tx: 10.289, ty: 10.202, tw: 4.476, n: '06', body: T.synth },
  ];
  items.forEach((it, i) => {
    targetIcon(s, it.ix, it.iy, 2.086, ACCENTS[i], GRADS[i], THEME[i]);
    s.addText([
      { text: `Analysis ${it.n}`, options: { fontFace: F.robotoBlack, fontSize: 28, bold: true, breakLine: true } },
      { text: it.body, options: { fontFace: F.light, fontSize: 21 } },
    ], { x: it.tx, y: it.ty, w: it.tw, h: 1.594, valign: 'top', color: C.text, lineSpacingMultiple: 1.3 });
  });
  artwork(s, 14.789, 4.271, 11.332, 8.085);
}

function slide14(s) {
  header(s);
  artwork(s, 0.878, 3.286, 11.655, 10.352);
  const blocks = [
    { n: '01', i: 0, bx: 15.401, by: 3.773, lx: 15.618, ly: 3.804, rx: 13.220, ry: 4.606, rw: 5.058, gx: 13.198, gy: 5.034, gw: 5.099, gh: 1.482, align: 'right' },
    { n: '02', i: 1, bx: 19.848, by: 3.773, lx: 20.053, ly: 3.813, rx: 19.866, ry: 4.533, rw: 5.026, gx: 19.888, gy: 5.024, gw: 5.005, gh: 1.077, align: 'left' },
    { n: '03', i: 2, bx: 15.460, by: 7.443, lx: 15.749, ly: 7.503, rx: 13.220, ry: 8.229, rw: 5.058, gx: 13.198, gy: 8.656, gw: 5.099, gh: 1.482, align: 'right' },
    { n: '04', i: 3, bx: 19.848, by: 7.443, lx: 20.053, ly: 7.513, rx: 19.861, ry: 8.203, rw: 5.026, gx: 19.882, gy: 8.694, gw: 5.005, gh: 1.077, align: 'left' },
    { n: '05', i: 4, bx: 15.417, by: 11.066, lx: 15.630, ly: 11.109, rx: 13.220, ry: 11.917, rw: 5.058, gx: 13.198, gy: 12.344, gw: 5.099, gh: 1.482, align: 'right' },
    { n: '06', i: 5, bx: 19.848, by: 11.066, lx: 20.053, ly: 11.136, rx: 19.866, ry: 11.816, rw: 5.026, gx: 19.888, gy: 12.308, gw: 5.005, gh: 1.077, align: 'left' },
  ];
  blocks.forEach((b) => {
    s.addShape('roundRect', { x: b.bx, y: b.by, w: 2.836, h: 0.564, rectRadius: 0.09, fill: { color: GRADS[b.i] } });
    s.addShape('roundRect', { x: b.bx + 0.065, y: b.by + 0.065, w: 2.707, h: 0.434, rectRadius: 0.07, fill: { color: ACCENTS[b.i] } });
    s.addText(`Analysis ${b.n}`, {
      x: b.lx, y: b.ly, w: 2.426, h: 0.628, margin: 0, align: 'center', valign: 'top',
      fontFace: F.robotoBlack, fontSize: 20, bold: true, color: C.white, lineSpacingMultiple: 1.3,
    });
    s.addText(T.synth, {
      x: b.rx, y: b.ry, w: b.rw, h: 0.467, margin: 0, align: b.align, valign: 'top',
      fontFace: F.light, fontSize: 18, color: ACCENTS[b.i], lineSpacingMultiple: 1.3,
    });
    s.addText(T.loremMid, {
      x: b.gx, y: b.gy, w: b.gw, h: b.gh, margin: 0, align: b.align, valign: 'top',
      fontFace: F.light, fontSize: 20, color: C.gray, lineSpacingMultiple: 1.5,
    });
  });
}

function slide15(s) {
  header(s);
  quote(s, 1.505, 4.352, 12.506, 4.140, [
    ['Life is a series of natural and '], ['spontaneous', C.a1, true],
    [" changes. Don't resist them; that only creates "], ['sorrow', C.a2, true],
    ['. Let reality be reality. Let things flow naturally '], ['forward', C.a3, true],
    [' in whatever way they like.'],
  ]);
  strap(s, 1.503, 8.444, C.a1);
  caption(s, 1.503, 9.396);
  body(s, 1.503, 9.828, 10.640, 2.461);
  artwork(s, 13.606, 3.720, 11.559, 10.018);
}

function slide16(s) {
  header(s);
  artwork(s, 0.694, 4.257, 12.641, 8.122);
  quote(s, 13.766, 4.387, 10.640, 2.524, [
    ["You can't "], ['always', C.a1, true], [' control what '], ['goes', C.a2, true], [' on '],
    ['outside', C.a3, true], [', but you can '], ['control', C.a4, true], [' what goes on '],
    ['inside', C.a5, true], ['.'],
  ]);
  strap(s, 13.766, 6.921, C.a1);
  body(s, 13.766, 7.922, 10.640, 2.461);
  [10.782, 11.774, 12.771].forEach((y) => {
    tickGlyph(s, 13.955, y, 0.604, C.a5, 1.375);
    s.addText(T.synth, {
      x: 14.705, y: y + 0.011, w: 6.944, h: 0.583, valign: 'top', wrap: false,
      fontFace: F.light, fontSize: 24, color: C.text, lineSpacingMultiple: 1.3,
    });
  });
}

function slide17(s) {
  coverSlide(s, -0.1666, [{ text: 'CLOSING SLIDE', options: { fontFace: F.reg, fontSize: 44, bold: true, color: C.white } }],
    { text: 'Pulsecolor', x: 12.526, y: 9.628, w: 1.645 }, false);
}

/* --------------------------------------------------------------------- main */
function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'STRESS', width: W, height: H });
  pptx.layout = 'STRESS';
  pptx.author = 'Pulsecolor';
  pptx.title = 'Stress Level Infographic';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
    slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17]
    .forEach((fn) => fn(pptx.addSlide()));

  return pptx.writeFile({
    fileName: path.join(__dirname, '0e8a3042-cc8c-4f61-8502-07cf2852f40c_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
