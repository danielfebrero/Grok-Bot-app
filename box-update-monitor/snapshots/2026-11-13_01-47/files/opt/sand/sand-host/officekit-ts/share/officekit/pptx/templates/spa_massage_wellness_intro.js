/**
 * VIRIA — Presentation Template  (30 slides, 13.333in x 7.5in)
 * Rebuilt with pptxgenjs. Raster photos in the source deck are replaced by
 * flat placeholder rectangles; everything else is drawn with native shapes.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const DARK   = '2B2B2B';  // theme tx1 / dk1
const INK    = '404040';  // tx1 @ 90% lum
const BODY   = '606060';  // body copy grey
const MUTED  = '959595';  // caption grey
const MIST   = 'EAEAEA';  // chart neutral
const WHITE  = 'FFFFFF';
const AMBER  = 'F49E14';  // theme accent1
const ORANGE = 'FF7427';  // theme accent3

const HEAD = 'Roboto';    // theme major font
const BASE = 'Open Sans'; // theme minor font

const INSET = [7.2, 7.2, 3.6, 3.6];   // lIns, rIns, bIns, tIns (pt) = 0.1/0.05 in
const KICKER = 'Add Seomething Here';
const BR = '\n\n';   // blank paragraph between two blocks of body copy

/* ------------------------------------------------------------------- shapes */
function newSlide(pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  return s;
}

// solid rectangle (colour blocks that make up each layout)
function band(s, o) {
  s.addShape('rect', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: { color: o.c }, line: { type: 'none' },
    rotate: o.rot || 0,
  });
}

// thin horizontal accent rule
function rule(s, o) {
  s.addShape('line', {
    x: o.x, y: o.y, w: o.w, h: 0,
    line: { color: o.c || DARK, width: 1.5 },
  });
}

// stand-in for a photo/mock-up bitmap from the source deck
function photo(s, o) {
  s.addShape('rect', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: { color: 'D9D9D9' }, line: { color: 'BFBFBF', width: 1 },
  });
  s.addText('[image]', {
    x: o.x, y: o.y + o.h / 2 - 0.16, w: o.w, h: 0.32,
    fontSize: 10, fontFace: BASE, color: '7F7F7F', align: 'center', valign: 'middle',
  });
}

// 45-degree barber-stripe block (source: "Freeform 41", drawn at -30 deg)
const STRIPE_BARS = [
  [[0.676, 0.927], [0.634, 1.000], [0.507, 0.927]],
  [[0.765, 0.772], [0.721, 0.850], [0.374, 0.850], [0.240, 0.772]],
  [[0.854, 0.618], [0.810, 0.695], [0.107, 0.695], [0.000, 0.634], [0.009, 0.618]],
  [[0.944, 0.463], [0.899, 0.541], [0.054, 0.541], [0.098, 0.463]],
  [[0.902, 0.309], [1.000, 0.365], [0.988, 0.387], [0.143, 0.387], [0.187, 0.309]],
  [[0.635, 0.154], [0.769, 0.232], [0.232, 0.232], [0.276, 0.154]],
  [[0.368, 0.000], [0.502, 0.077], [0.321, 0.077], [0.365, 0.000]],
];
function stripes(s, o) {
  const w = o.w, h = o.w * 0.999, pts = [];
  STRIPE_BARS.forEach(bar => {
    bar.forEach((p, i) => pts.push({ x: p[0] * w, y: p[1] * h, moveTo: i === 0 }));
    pts.push({ close: true });
  });
  s.addShape('custGeom', {
    x: o.x, y: o.y, w: w, h: h, rotate: -30,
    fill: { color: o.c }, line: { type: 'none' }, points: pts,
  });
}

// full-height panel whose left edge bows out (title / closing slides)
function arcPanel(s, o) {
  const w = o.w, h = o.h;
  s.addShape('custGeom', {
    x: o.x, y: o.y, w: w, h: h, fill: { color: o.c }, line: { type: 'none' },
    points: [
      { x: 0.104 * w, y: 0 },
      { x: w, y: 0 }, { x: w, y: h }, { x: 0.104 * w, y: h },
      { x: 0.099 * w, y: 0.990 * h },
      { x: 0, y: 0.500 * h, curve: { type: 'cubic', x1: 0.036 * w, y1: 0.847 * h, x2: 0, y2: 0.679 * h } },
      { x: 0.099 * w, y: 0.010 * h, curve: { type: 'cubic', x1: 0, y1: 0.321 * h, x2: 0.036 * w, y2: 0.153 * h } },
      { close: true },
    ],
  });
}

// name plate with a half-round bottom edge (team slide)
function arcBanner(s, o) {
  const w = o.w, h = o.h;
  s.addShape('custGeom', {
    x: o.x, y: o.y, w: w, h: h, fill: { color: o.c }, line: { type: 'none' },
    points: [
      { x: 0, y: 0 }, { x: w, y: 0 }, { x: 0.958 * w, y: 0.194 * h },
      { x: 0.500 * w, y: h, curve: { type: 'cubic', x1: 0.841 * w, y1: 0.692 * h, x2: 0.679 * w, y2: h } },
      { x: 0.042 * w, y: 0.194 * h, curve: { type: 'cubic', x1: 0.321 * w, y1: h, x2: 0.159 * w, y2: 0.692 * h } },
      { close: true },
    ],
  });
}

/* --------------------------------------------------------------------- text */
// every text box in the deck is top-anchored with the same 0.1in/0.05in insets;
// `o` carries only the values that differ from the style default `d`
function put(s, txt, o, d) {
  s.addText(txt, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontSize: o.sz ?? d.sz,
    fontFace: o.face ?? d.face,
    color: o.c ?? d.c,
    bold: o.b ?? d.b,
    align: o.al ?? d.al,
    lineSpacingMultiple: o.ls ?? d.ls,
    charSpacing: o.spc ?? d.spc,
    valign: 'top', margin: INSET, wrap: true,
  });
}

// big all-caps statement set in the major typeface
const title  = (s, t, o) => put(s, rich(t, o), o, { sz: 22,  face: HEAD, c: DARK,  b: true,  al: 'left', ls: 0, spc: 0 });
// bold sub-heading, e.g. "01. Service One" / "Welcome Massage"
const label  = (s, t, o) => put(s, t, o, { sz: 12,  face: HEAD, c: DARK,  b: true,  al: 'left', ls: 0, spc: 0 });
// running body copy
const para   = (s, t, o) => put(s, t, o, { sz: 10,  face: BASE, c: BODY,  b: false, al: 'justify', ls: 1.5, spc: 0 });
// small grey caption under a heading
const note   = (s, t, o) => put(s, t, o, { sz: 8,   face: BASE, c: MUTED, b: false, al: 'justify', ls: 1.5, spc: 0 });
// letter-spaced strapline
const kicker = (s, t, o) => put(s, t, o, { sz: 9,   face: BASE, c: DARK,  b: false, al: 'left', ls: 1.5, spc: 3 });

// title() accepts either a plain string or [[text, colour], ...] for mixed runs
function rich(t, o) {
  if (!Array.isArray(t)) return t;
  return t.map(r => ({ text: r[0], options: { color: r[1] } }));
}

// "12." page marker in the corner
function pageNum(s, num, o) {
  s.addText(num + '.', {
    x: o.x, y: o.y, w: 0.323, h: 0.269,
    fontSize: 10, fontFace: HEAD, bold: true, color: o.c,
    align: 'center', valign: 'top', margin: INSET, wrap: false,
  });
}

/* ------------------------------------------------------------------- charts */
const GRID = { color: 'D9D9D9', size: 0.75, style: 'solid' };
const AXIS = {
  showLegend: false, showTitle: false,
  valAxisLabelFormatCode: '0%',
  valGridLine: GRID, catGridLine: { style: 'none' },
  catAxisLineShow: false, valAxisLineShow: false,
  catAxisMajorTickMark: 'none', valAxisMajorTickMark: 'none',
  catAxisLabelFontFace: BASE, catAxisLabelFontSize: 8, catAxisLabelColor: '757575',
  valAxisLabelFontFace: BASE, valAxisLabelFontSize: 8, valAxisLabelColor: '757575',
};

function stackedChart(s, pptx, o) {
  s.addChart(pptx.ChartType.bar, CAT_SERIES, {
    ...AXIS,
    x: o.x, y: o.y, w: o.w, h: o.h,
    barDir: o.dir, barGrouping: 'percentStacked',
    barGapWidthPct: 150, barOverlapPct: 100,
    chartColors: o.colors,
  });
}

function ring(s, pptx, o) {
  s.addChart(pptx.ChartType.doughnut, [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: [8.2, 3.2] }], {
    x: o.x, y: o.y, w: o.w, h: o.h,
    holeSize: 75, chartColors: [AMBER, INK],
    dataBorder: { pct: 1.5, color: WHITE },
    showLegend: false, showTitle: false, showValue: false,
  });
}

// pptxgenjs only writes `grouping` for area charts when it is 'stacked', so the
// 100%-stacked look is reproduced by stacking pre-normalised shares.
function areaChart(s, pptx, o) {
  const totals = AREA_RAW[0].map((_, i) => AREA_RAW[0][i] + AREA_RAW[1][i]);
  const share = row => row.map((v, i) => v / totals[i]);
  s.addChart(pptx.ChartType.area, [
    { name: 'Series 1', labels: MONTHS, values: share(AREA_RAW[0]) },
    { name: 'Series 2', labels: MONTHS, values: share(AREA_RAW[1]) },
  ], {
    ...AXIS,
    x: o.x, y: o.y, w: o.w, h: o.h,
    barGrouping: 'stacked',
    valAxisMinVal: 0, valAxisMaxVal: 1, valAxisMajorUnit: 0.1,
    chartColors: [ORANGE, INK],
  });
}

/* --------------------------------------------------------------- chart data */
const CATS = ['Category 1', 'Category 2', 'Category 3'];
const CAT_SERIES = [
  { name: 'Series 1', labels: CATS, values: [4.3, 2.5, 3.5] },
  { name: 'Series 2', labels: CATS, values: [2.4, 4.4, 1.8] },
  { name: 'Series 3', labels: CATS, values: [2, 2, 3] },
];
const MONTHS = ['5/1/2002', '6/1/2002', '7/1/2002', '8/1/2002', '9/1/2002'];
const AREA_RAW = [[32, 32, 28, 12, 15], [12, 12, 12, 21, 28]];

/* ------------------------------------------------------------ text catalogue */
const T = {
  HEADLINE: "HEALTH IS NOT SIMPLY THE ABSENCE OF SICKNESS",
  HEAD_A: "HEALTH IS NOT SIMPLY THE ABSENCE ",
  HEAD_B: "OF SICKNESS",
  L01: "lorem ipsum anisai dolor atilinui sitanil amet nias",
  L02: "lorem ipsum dolor sitaniasul nisa amuit",
  L03: "PLACEHOLDER",
  L04: "PLACEHOLDER",
  L05: "PLACEHOLDER",
  L06: "PLACEHOLDER",
  L07: "lorem ipsum dolor sitanil amet conse hanit teturasi",
  L08: "lorem ipsum dolor sitaniasul nisa",
  L09: "PLACEHOLDER",
  L10: "lorem ipsum dolor sitaniasul nisanilasil anisani adipiscinge elit sed cursus misanis nisanal hanisi adipiscinge elit sed cursus misanis  ",
  L11: "PLACEHOLDER",
  L12: "PLACEHOLDER",
  L13: "PLACEHOLDER",
  L14: "PLACEHOLDER",
  L15: "PLACEHOLDER",
  L16: "PLACEHOLDER",
  L17: "PLACEHOLDER",
  L18: "PLACEHOLDER",
  L19: "PLACEHOLDER",
  L20: "PLACEHOLDER",
  L21: "PLACEHOLDER",
  L22: "PLACEHOLDER",
  L23: "PLACEHOLDER",
  L24: "PLACEHOLDER",
  L25: "PLACEHOLDER",
  L26: "PLACEHOLDER",
  L27: "PLACEHOLDER",
  L28: "PLACEHOLDER",
  L29: "PLACEHOLDER",
  L30: "PLACEHOLDER",
  L31: "PLACEHOLDER",
  L32: "PLACEHOLDER",
  L33: "PLACEHOLDER",
  L34: "PLACEHOLDER",
  L35: "PLACEHOLDER",
  L36: "PLACEHOLDER",
  L37: "lorem ipsum dolor sita amet consec teturasinis adipiscingesa ",
  L38: "misani utasisau nasat sapienan adipiscini hanin",
  L39: "PLACEHOLDER",
  L40: "PLACEHOLDER",
  L41: "PLACEHOLDER",
};

/* --------------------------------------------------------------- 30 slides */

function slide01(pptx) {
  const s = newSlide(pptx);
  arcPanel(s, { x:6.764, y:0, w:6.569, h:7.5, c:INK });
  label(s, "Welcome Massage", { x:0.806, y:0.864, w:1.691, h:0.303, c:AMBER });
  rule(s, { x:0.906, y:6.737, w:3.447 });
  title(s, [['V', DARK], ['I', AMBER], ['RIA', DARK]], { x:2.738, y:2.339, w:6.071, h:2.423, sz:138 });
  kicker(s, "Presentation Template", { x:2.42, y:4.435, w:4.274, h:0.349, sz:11, al:'center' });
  pageNum(s, 1, { x:12.481, y:6.552, c:WHITE });
}

function slide02(pptx) {
  const s = newSlide(pptx);
  band(s, { x:6.76, y:0.927, w:7.5, h:5.646, c:DARK, rot:90 });
  band(s, { x:0, y:0, w:2.042, h:7.5, c:ORANGE });
  para(s, T.L03 + BR + T.L04, { x:3.662, y:3.487, w:3.147, h:2.346 });
  rule(s, { x:0.958, y:6.747, w:2.042 });
  label(s, "Welcome Massage", { x:5.125, y:6.567, w:1.685, h:0.303, c:AMBER, al:'right' });
  title(s, T.HEADLINE, { x:3.642, y:1.268, w:3.265, h:1.161, sz:21 });
  kicker(s, KICKER, { x:3.662, y:2.565, w:3.265, h:0.281, sz:8 });
  stripes(s, { x:11.012, y:0.411, w:1.913, c:WHITE });
  pageNum(s, 2, { x:12.522, y:6.601, c:WHITE });
}

function slide03(pptx) {
  const s = newSlide(pptx);
  band(s, { x:7.806, y:-1.899, w:3.642, h:7.42, c:ORANGE, rot:90 });
  para(s, T.L21, { x:5.917, y:5.325, w:6.083, h:0.831 });
  rule(s, { x:7.906, y:6.735, w:3.987 });
  label(s, "Welcome Massage", { x:5.917, y:6.525, w:1.691, h:0.303, c:AMBER });
  title(s, T.HEADLINE, { x:0.622, y:5.357, w:4.651, h:0.841, al:'center' });
  kicker(s, KICKER, { x:0.811, y:6.526, w:4.274, h:0.304, al:'center' });
  stripes(s, { x:11.023, y:0.38, w:1.913, c:WHITE });
  pageNum(s, 3, { x:12.522, y:6.601, c:INK });
}

function slide04(pptx) {
  const s = newSlide(pptx);
  band(s, { x:9.68, y:3.739, w:2.069, h:3.342, c:DARK, rot:90 });
  title(s, T.HEAD_A + '\n' + T.HEAD_B, { x:0.916, y:1.396, w:2.507, h:1.515, sz:21 });
  kicker(s, KICKER, { x:0.937, y:3.138, w:2.507, h:0.281, sz:8 });
  para(s, T.L22 + BR + T.L23, { x:0.937, y:3.923, w:2.289, h:2.346 });
  label(s, "Welcome Massage", { x:9.872, y:4.598, w:1.685, h:0.303, c:AMBER, al:'center' });
  rule(s, { x:9.043, y:5.096, w:3.342, c:WHITE });
  para(s, T.L24, { x:9.265, y:5.355, w:2.902, h:0.831, c:WHITE, al:'center' });
}

function slide05(pptx) {
  const s = newSlide(pptx);
  band(s, { x:9.191, y:-1.961, w:2.181, h:6.103, c:ORANGE, rot:90 });
  para(s, T.L25, { x:7.922, y:3.347, w:4.469, h:0.831 });
  rule(s, { x:9.906, y:4.666, w:2.375 });
  label(s, "Welcome Massage", { x:7.922, y:4.456, w:1.957, h:0.303, c:AMBER });
  title(s, T.HEADLINE, { x:2.462, y:5.847, w:8.409, h:0.505, sz:24, al:'center' });
  kicker(s, KICKER, { x:4.53, y:6.445, w:4.274, h:0.304, al:'center' });
  pageNum(s, 5, { x:12.522, y:6.601, c:INK });
}

function slide06(pptx) {
  const s = newSlide(pptx);
  band(s, { x:9.257, y:3.403, w:2.913, h:5.281, c:DARK, rot:90 });
  band(s, { x:0, y:0, w:3.66, h:7.5, c:ORANGE });
  para(s, T.L26 + BR + T.L27, { x:7.175, y:1.804, w:5.127, h:2.094 });
  rule(s, { x:7.289, y:1.198, w:2.867 });
  stripes(s, { x:0.626, y:5.111, w:1.913, c:WHITE });
  label(s, "Welcome Massage", { x:10.042, y:1.042, w:2.26, h:0.303, c:AMBER, al:'right' });
  title(s, T.HEADLINE, { x:9.466, y:5.007, w:2.836, h:1.01, sz:18, c:WHITE });
  kicker(s, KICKER, { x:9.466, y:6.105, w:3.198, h:0.281, sz:8, c:WHITE });
  pageNum(s, 6, { x:12.512, y:6.674, c:WHITE });
}

function slide07(pptx) {
  const s = newSlide(pptx);
  band(s, { x:0, y:0, w:13.333, h:3.656, c:ORANGE });
  stripes(s, { x:0.419, y:5.256, w:1.913, c:DARK });
  title(s, T.HEAD_A + '\n' + T.HEAD_B, { x:5.409, y:4.275, w:2.289, h:1.447, sz:20 });
  para(s, T.L28, { x:8.32, y:1.662, w:4.177, h:1.336, c:WHITE });
  rule(s, { x:8.425, y:1.257, w:1.933 });
  label(s, "About Our Service", { x:10.729, y:1.101, w:1.768, h:0.303, c:WHITE, al:'right' });
  para(s, T.L14, { x:8.788, y:6.047, w:3.709, h:0.579 });
  label(s, "02. Service Two", { x:8.788, y:5.627, w:2.991, h:0.303, c:ORANGE });
  para(s, T.L14, { x:8.788, y:4.719, w:3.709, h:0.579 });
  label(s, "01. Service One", { x:8.788, y:4.299, w:2.991, h:0.303 });
  kicker(s, KICKER, { x:5.409, y:6.33, w:2.216, h:0.281, sz:8 });
  pageNum(s, 7, { x:12.645, y:6.894, c:INK });
}

function slide08(pptx) {
  const s = newSlide(pptx);
  band(s, { x:5.583, y:-0.25, w:2.167, h:13.333, c:DARK, rot:90 });
  band(s, { x:0, y:0, w:13.333, h:0.666, c:ORANGE });
  rule(s, { x:9.5, y:4.718, w:2.167 });
  para(s, "lorem ipsum dolor sitaniasul", { x:1.745, y:1.495, w:2.089, h:0.326, c:MUTED, al:'center' });
  label(s, "01. Service One", { x:1.745, y:1.189, w:2.033, h:0.303, al:'center' });
  para(s, "lorem ipsum dolor sitaniasul", { x:5.622, y:1.495, w:2.089, h:0.326, c:MUTED, al:'center' });
  label(s, "02. Service Two", { x:5.622, y:1.189, w:2.033, h:0.303, al:'center' });
  para(s, "lorem ipsum dolor sitaniasul", { x:9.5, y:1.495, w:2.089, h:0.326, c:MUTED, al:'center' });
  label(s, "03. Service Three", { x:9.5, y:1.189, w:2.033, h:0.303, al:'center' });
  rule(s, { x:1.653, y:4.707, w:2.167 });
  title(s, T.HEADLINE, { x:2.462, y:5.837, w:8.409, h:0.505, sz:24, c:WHITE, al:'center' });
  kicker(s, KICKER, { x:4.53, y:6.628, w:4.274, h:0.304, c:WHITE, al:'center' });
  label(s, "About Our Service", { x:0.522, y:6.68, w:1.768, h:0.286, sz:10.5, c:AMBER });
  pageNum(s, 8, { x:12.428, y:6.676, c:WHITE });
}

function slide09(pptx) {
  const s = newSlide(pptx);
  band(s, { x:0, y:0, w:6.667, h:4.376, c:ORANGE });
  para(s, T.L05, { x:8.048, y:1.585, w:4.177, h:0.831 });
  label(s, "01. Service One", { x:8.048, y:1.164, w:3.369, h:0.303 });
  para(s, T.L05, { x:8.048, y:3.615, w:4.177, h:0.831 });
  label(s, "02. Service Two", { x:8.048, y:3.195, w:3.369, h:0.303, c:ORANGE });
  para(s, T.L05, { x:8.048, y:5.646, w:4.177, h:0.831 });
  label(s, "03. Service Three", { x:8.048, y:5.225, w:3.369, h:0.303 });
  stripes(s, { x:-0.549, y:6.227, w:1.913, c:DARK });
  title(s, T.HEADLINE, { x:2.021, y:5.72, w:4.319, h:0.808, sz:21 });
  kicker(s, KICKER, { x:2.049, y:5.197, w:3.265, h:0.281, sz:8 });
  rule(s, { x:10.674, y:0.597, w:2.087 });
  pageNum(s, 9, { x:12.645, y:6.894, c:INK });
}

function slide10(pptx) {
  const s = newSlide(pptx);
  band(s, { x:9.728, y:0, w:3.605, h:3.75, c:ORANGE });
  band(s, { x:-2.115, y:3.031, w:6.583, h:2.354, c:DARK, rot:90 });
  para(s, T.L15, { x:4.556, y:5.428, w:3.162, h:1.084 });
  label(s, "01. Service One", { x:4.556, y:4.883, w:2.687, h:0.303 });
  para(s, T.L15, { x:8.458, y:5.428, w:3.162, h:1.084 });
  label(s, "02. Service Two", { x:8.458, y:4.883, w:2.687, h:0.303, c:ORANGE });
  title(s, T.HEADLINE, { x:3.473, y:1.141, w:4.651, h:0.841 });
  kicker(s, KICKER, { x:3.473, y:2.127, w:4.463, h:0.304 });
  para(s, T.L29, { x:3.473, y:2.984, w:5.215, h:0.831 });
  pageNum(s, 10, { x:12.512, y:6.652, c:INK });
}

function slide11(pptx) {
  const s = newSlide(pptx);
  band(s, { x:5.34, y:-3.993, w:2.654, h:13.333, c:DARK, rot:90 });
  para(s, T.L16, { x:1.014, y:2.353, w:2.993, h:1.084, c:WHITE });
  label(s, "01. Service One", { x:1.014, y:1.911, w:2.542, h:0.303, c:WHITE });
  para(s, T.L16, { x:4.405, y:2.353, w:2.993, h:1.084, c:WHITE });
  label(s, "02. Service Two", { x:4.405, y:1.911, w:2.542, h:0.303, c:WHITE });
  para(s, T.L30, { x:6.401, y:5.325, w:5.777, h:0.831 });
  rule(s, { x:8.39, y:6.735, w:3.722 });
  label(s, "About Our Service", { x:6.401, y:6.525, w:1.691, h:0.303, c:AMBER });
  title(s, T.HEADLINE, { x:1.014, y:5.357, w:4.463, h:0.841 });
  kicker(s, KICKER, { x:1.014, y:6.526, w:4.274, h:0.304 });
  pageNum(s, 11, { x:12.522, y:6.601, c:INK });
}

function slide12(pptx) {
  const s = newSlide(pptx);
  band(s, { x:0, y:5.085, w:2.653, h:2.415, c:DARK });
  band(s, { x:10.681, y:0, w:2.653, h:2.415, c:ORANGE });
  para(s, T.L31, { x:9.275, y:5.154, w:2.823, h:1.084, al:'right' });
  label(s, "02. Service Two", { x:9.275, y:4.609, w:2.823, h:0.303, al:'right' });
  para(s, T.L32, { x:1.236, y:1.985, w:2.823, h:1.084 });
  label(s, "01. Service One", { x:1.236, y:1.439, w:2.687, h:0.303 });
  title(s, T.HEAD_A + '\n' + T.HEAD_B, { x:5.638, y:2.76, w:2.289, h:1.447, sz:20, c:WHITE });
  kicker(s, KICKER, { x:5.638, y:4.479, w:2.216, h:0.281, sz:8, c:WHITE });
  rule(s, { x:0.747, y:0.737, w:1.776 });
  pageNum(s, 12, { x:12.509, y:6.628, c:INK });
}

function slide13(pptx) {
  const s = newSlide(pptx);
  band(s, { x:9.236, y:0, w:4.097, h:7.5, c:ORANGE });
  title(s, T.HEAD_A + '\n' + T.HEAD_B, { x:10.194, y:1.025, w:2.507, h:1.515, sz:21, c:WHITE });
  kicker(s, KICKER, { x:10.215, y:2.766, w:2.507, h:0.281, sz:8, c:WHITE });
  para(s, T.L33 + BR + T.L34, { x:10.215, y:3.674, w:2.289, h:1.841, c:WHITE });
  note(s, "lorem ipsum dolor sitaniasul", { x:1.141, y:6.515, w:2.089, h:0.281, al:'center' });
  label(s, "01. Portfolio One", { x:1.141, y:6.271, w:2.033, h:0.269, sz:10, al:'center' });
  note(s, "lorem ipsum dolor sitaniasul", { x:3.547, y:6.509, w:2.089, h:0.281, al:'center' });
  label(s, "02. Portfolio Two", { x:3.547, y:6.265, w:2.033, h:0.269, sz:10, al:'center' });
  note(s, "lorem ipsum dolor sitaniasul", { x:5.952, y:6.503, w:2.089, h:0.281, al:'center' });
  label(s, "03. Portfolio Three", { x:5.952, y:6.258, w:2.033, h:0.269, sz:10, al:'center' });
  rule(s, { x:10.327, y:6.702, w:1.683 });
  pageNum(s, 13, { x:12.129, y:6.567, c:WHITE });
}

function slide14(pptx) {
  const s = newSlide(pptx);
  band(s, { x:5.294, y:-1.072, w:2.744, h:13.333, c:DARK, rot:90 });
  note(s, "lorem ipsum dolor sitaniasul", { x:1.453, y:6.183, w:2.089, h:0.281, c:WHITE, al:'center' });
  label(s, "01. Portfolio One", { x:1.453, y:5.939, w:2.033, h:0.269, sz:10, c:WHITE, al:'center' });
  note(s, "lorem ipsum dolor sitaniasul", { x:4.3, y:6.19, w:2.089, h:0.281, c:WHITE, al:'center' });
  label(s, "02. Portfolio Two", { x:4.3, y:5.945, w:2.033, h:0.269, sz:10, c:WHITE, al:'center' });
  note(s, "lorem ipsum dolor sitaniasul", { x:7.046, y:6.196, w:2.089, h:0.281, c:WHITE, al:'center' });
  label(s, "03. Portfolio Three", { x:7.046, y:5.952, w:2.033, h:0.269, sz:10, c:WHITE, al:'center' });
  note(s, "lorem ipsum dolor sitaniasul", { x:9.792, y:6.203, w:2.089, h:0.281, c:WHITE, al:'center' });
  label(s, "04. Portfolio Four", { x:9.792, y:5.958, w:2.033, h:0.269, sz:10, c:WHITE, al:'center' });
  para(s, T.L06, { x:6.94, y:1.298, w:5.046, h:0.831 });
  rule(s, { x:9.016, y:0.873, w:2.865 });
  label(s, "About Our Portfolio", { x:6.932, y:0.698, w:1.691, h:0.303, c:AMBER });
  title(s, T.HEADLINE, { x:1.255, y:1.329, w:4.463, h:0.841 });
  kicker(s, KICKER, { x:1.255, y:0.657, w:4.274, h:0.304 });
  pageNum(s, 14, { x:12.429, y:0.732, c:INK });
}

function slide15(pptx) {
  const s = newSlide(pptx);
  band(s, { x:10.984, y:0, w:2.35, h:2.485, c:ORANGE });
  title(s, T.HEADLINE, { x:0.995, y:1.484, w:4.651, h:0.841 });
  kicker(s, KICKER, { x:1.023, y:2.47, w:4.463, h:0.304 });
  para(s, T.L07, { x:1.023, y:4.882, w:3.709, h:0.326 });
  label(s, "02. Portfolio Two", { x:1.023, y:4.576, w:2.991, h:0.303, c:ORANGE });
  para(s, T.L07, { x:1.023, y:3.78, w:3.709, h:0.326 });
  label(s, "01. Portfolio One", { x:1.023, y:3.474, w:2.991, h:0.303 });
  para(s, T.L07, { x:1.023, y:6.038, w:3.709, h:0.326 });
  label(s, "03. Portfolio Three", { x:1.023, y:5.733, w:2.991, h:0.303 });
  rule(s, { x:0.847, y:0.737, w:1.676 });
  stripes(s, { x:5.601, y:5.442, w:1.622, c:DARK });
  pageNum(s, 15, { x:12.461, y:6.692, c:INK });
}

function slide16(pptx) {
  const s = newSlide(pptx);
  band(s, { x:0, y:0, w:3.487, h:2.937, c:DARK });
  band(s, { x:4.044, y:6.361, w:2.388, h:1.139, c:ORANGE });
  para(s, T.L17 + BR + T.L18, { x:7.775, y:1.88, w:4.551, h:2.094 });
  rule(s, { x:7.894, y:1.315, w:2.28 });
  label(s, "About Our Portfolio", { x:10.08, y:1.16, w:2.26, h:0.303, c:AMBER, al:'right' });
  stripes(s, { x:12.034, y:6.199, w:1.913, c:DARK });
  title(s, T.HEADLINE, { x:7.771, y:4.867, w:4.463, h:0.841 });
  kicker(s, KICKER, { x:7.771, y:6.036, w:4.274, h:0.304 });
  pageNum(s, 16, { x:0.558, y:6.692, c:INK });
}

function slide17(pptx) {
  const s = newSlide(pptx);
  band(s, { x:0, y:0, w:3.792, h:7.5, c:DARK });
  title(s, T.HEADLINE, { x:7.599, y:1.227, w:4.651, h:0.841 });
  kicker(s, KICKER, { x:7.626, y:2.213, w:4.463, h:0.304 });
  para(s, T.L35, { x:7.599, y:3.085, w:4.651, h:1.084 });
  note(s, T.L02, { x:7.599, y:5.199, w:2.302, h:0.281, al:'left' });
  label(s, "01. Write Something Here", { x:7.599, y:4.892, w:2.033, h:0.269, sz:10 });
  note(s, T.L02, { x:7.599, y:6.153, w:2.302, h:0.281, al:'left' });
  label(s, "02. Write Something Here", { x:7.599, y:5.846, w:2.033, h:0.269, sz:10 });
  arcBanner(s, { x:1.705, y:5.481, w:4.194, h:0.988, c:ORANGE });
  note(s, "lorem ipsum dolor sitani amit", { x:2.64, y:5.862, w:2.302, h:0.281, c:WHITE, al:'center' });
  label(s, "Add Your Name", { x:2.64, y:5.664, w:2.302, h:0.269, sz:10, c:WHITE, al:'center' });
  pageNum(s, 17, { x:12.558, y:6.859, c:INK });
}

function slide18(pptx) {
  const s = newSlide(pptx);
  para(s, T.L06, { x:6.94, y:1.331, w:5.046, h:0.831 });
  rule(s, { x:8.767, y:0.906, w:3.114 });
  label(s, "About Our Team", { x:6.932, y:0.732, w:1.691, h:0.303, c:AMBER });
  title(s, T.HEADLINE, { x:1.088, y:1.363, w:4.463, h:0.841 });
  kicker(s, KICKER, { x:1.088, y:0.691, w:4.274, h:0.304 });
  band(s, { x:1.043, y:7.079, w:11.113, h:0.421, c:DARK });
  para(s, T.L01, { x:10.559, y:5.577, w:1.447, h:0.831, al:'center' });
  label(s, "Your Name", { x:10.559, y:5.219, w:1.447, h:0.286, sz:10.5, al:'center' });
  para(s, T.L01, { x:8.686, y:5.577, w:1.447, h:0.831, al:'center' });
  label(s, "Your Name", { x:8.686, y:5.219, w:1.447, h:0.286, sz:10.5, al:'center' });
  para(s, T.L01, { x:6.813, y:5.577, w:1.447, h:0.831, al:'center' });
  label(s, "Your Name", { x:6.813, y:5.219, w:1.447, h:0.286, sz:10.5, al:'center' });
  para(s, T.L01, { x:4.939, y:5.582, w:1.447, h:0.831, al:'center' });
  label(s, "Your Name", { x:4.939, y:5.225, w:1.447, h:0.286, sz:10.5, al:'center' });
  para(s, T.L01, { x:3.066, y:5.57, w:1.447, h:0.831, al:'center' });
  label(s, "Your Name", { x:3.066, y:5.212, w:1.447, h:0.286, sz:10.5, al:'center' });
  para(s, T.L01, { x:1.192, y:5.558, w:1.447, h:0.831, al:'center' });
  label(s, "Your Name", { x:1.192, y:5.2, w:1.447, h:0.286, sz:10.5, al:'center' });
  pageNum(s, 18, { x:12.52, y:6.67, c:INK });
}

function slide19(pptx) {
  const s = newSlide(pptx);
  band(s, { x:0, y:0, w:6.549, h:3.175, c:ORANGE });
  rule(s, { x:7.722, y:6.742, w:4.381 });
  para(s, T.L36, { x:7.619, y:4.83, w:4.597, h:1.336 });
  note(s, "lorem ipsum dolor sitaniasul", { x:7.67, y:4.01, w:2.089, h:0.281, al:'center' });
  label(s, "Add Your Name", { x:7.67, y:3.766, w:2.033, h:0.269, sz:10, al:'center' });
  note(s, "lorem ipsum dolor sitaniasul", { x:10.077, y:4.01, w:2.089, h:0.281, al:'center' });
  label(s, "Add Your Name", { x:10.077, y:3.766, w:2.033, h:0.269, sz:10, face:'Open Sans', al:'center' });
  title(s, T.HEADLINE, { x:1.025, y:6.042, w:4.651, h:0.841 });
  kicker(s, KICKER, { x:1.053, y:5.57, w:4.463, h:0.304 });
  pageNum(s, 19, { x:12.491, y:6.601, c:INK });
}

function slide20(pptx) {
  const s = newSlide(pptx);
  band(s, { x:0, y:2.093, w:13.333, h:3.67, c:DARK });
  title(s, T.HEADLINE, { x:2.462, y:0.64, w:8.409, h:0.505, sz:24, al:'center' });
  kicker(s, KICKER, { x:4.53, y:1.294, w:4.274, h:0.304, al:'center' });
  note(s, T.L08, { x:2.22, y:5.18, w:2.089, h:0.281, c:WHITE, al:'center' });
  label(s, "Add Your Name", { x:2.22, y:4.935, w:2.033, h:0.269, sz:10, c:WHITE, al:'center' });
  note(s, T.L08, { x:5.622, y:5.161, w:2.089, h:0.281, c:WHITE, al:'center' });
  label(s, "Add Your Name", { x:5.622, y:4.917, w:2.033, h:0.269, sz:10, c:WHITE, al:'center' });
  note(s, T.L08, { x:9.025, y:5.142, w:2.089, h:0.281, c:WHITE, al:'center' });
  label(s, "Add Your Name", { x:9.025, y:4.898, w:2.033, h:0.269, sz:10, c:WHITE, al:'center' });
  para(s, T.L09, { x:2.462, y:6.348, w:8.409, h:0.579, al:'center' });
  pageNum(s, 20, { x:12.522, y:6.684, c:INK });
}

function slide21(pptx) {
  const s = newSlide(pptx);
  band(s, { x:5.33, y:0, w:8.004, h:3.75, c:ORANGE });
  stripes(s, { x:-0.504, y:6.503, w:1.913, c:DARK });
  para(s, T.L03 + BR + T.L04, { x:1.079, y:3.728, w:3.147, h:2.346 });
  title(s, T.HEADLINE, { x:1.058, y:1.509, w:3.265, h:1.161, sz:21 });
  kicker(s, KICKER, { x:1.079, y:2.848, w:3.265, h:0.281, sz:8 });
  note(s, T.L02, { x:9.992, y:4.839, w:2.302, h:0.281, al:'left' });
  label(s, "01. Write Something Here", { x:9.992, y:4.533, w:2.033, h:0.269, sz:10 });
  note(s, T.L02, { x:9.992, y:5.793, w:2.302, h:0.281, al:'left' });
  label(s, "02. Write Something Here", { x:9.992, y:5.486, w:2.033, h:0.269, sz:10 });
  para(s, T.L37 + '\n' + T.L38, { x:9.992, y:2.074, w:2.302, h:1.084, c:WHITE });
  rule(s, { x:10.094, y:1.637, w:0.995 });
  label(s, "Our Team", { x:11.3, y:1.481, w:0.995, h:0.303, c:WHITE, al:'right' });
  pageNum(s, 21, { x:12.558, y:6.859, c:INK });
}

function slide22(pptx) {
  const s = newSlide(pptx);
  band(s, { x:1.667, y:0, w:4.393, h:1.192, c:ORANGE });
  photo(s, { x:1.143, y:1.05, w:5.439, h:3.174 });
  para(s, T.L06, { x:1.384, y:5.717, w:5.046, h:0.831 });
  rule(s, { x:3.459, y:5.292, w:2.865 });
  label(s, "About Our Mockup", { x:1.376, y:5.118, w:1.691, h:0.303, c:AMBER });
  title(s, T.HEADLINE, { x:7.491, y:1.547, w:4.651, h:0.909, sz:24 });
  kicker(s, KICKER, { x:7.519, y:0.959, w:4.463, h:0.304 });
  note(s, T.L19, { x:7.519, y:3.434, w:1.96, h:0.685 });
  label(s, "01. Mockup One", { x:7.519, y:3.023, w:1.731, h:0.269, sz:10 });
  note(s, T.L19, { x:10.26, y:3.434, w:1.96, h:0.685 });
  label(s, "02. Mockup Two", { x:10.26, y:3.023, w:2.033, h:0.269, sz:10 });
  pageNum(s, 22, { x:12.558, y:6.859, c:INK });
}

function slide23(pptx) {
  const s = newSlide(pptx);
  band(s, { x:0, y:0, w:6.667, h:7.5, c:DARK });
  photo(s, { x:5.327, y:1.048, w:2.678, h:5.403 });
  rule(s, { x:8.006, y:6.742, w:4.097 });
  note(s, T.L10, { x:9.31, y:1.834, w:2.72, h:0.685 });
  label(s, "01. Mockup One", { x:9.31, y:1.424, w:1.731, h:0.269, sz:10 });
  note(s, T.L10, { x:9.31, y:3.455, w:2.72, h:0.685 });
  label(s, "02. Mockup Two", { x:9.31, y:3.044, w:1.731, h:0.269, sz:10, c:ORANGE });
  note(s, T.L10, { x:9.31, y:5.075, w:2.72, h:0.685 });
  label(s, "03. Mockup Three", { x:9.31, y:4.665, w:1.731, h:0.269, sz:10 });
  para(s, T.L03 + BR + T.L04, { x:1.065, y:3.59, w:3.147, h:2.346, c:WHITE });
  title(s, T.HEADLINE, { x:1.044, y:1.37, w:3.265, h:1.161, sz:21, c:WHITE });
  kicker(s, KICKER, { x:1.065, y:2.709, w:3.265, h:0.281, sz:8, c:WHITE });
  stripes(s, { x:-0.546, y:6.503, w:1.913, c:WHITE });
  pageNum(s, 23, { x:12.491, y:6.601, c:INK });
}

function slide24(pptx) {
  const s = newSlide(pptx);
  photo(s, { x:2.112, y:1.552, w:4.39, h:4.39 });
  photo(s, { x:6.831, y:1.552, w:4.39, h:4.39 });
  title(s, T.HEADLINE, { x:2.462, y:0.692, w:8.409, h:0.505, sz:24, al:'center' });
  kicker(s, KICKER, { x:4.53, y:1.346, w:4.274, h:0.304, al:'center' });
  para(s, T.L09, { x:2.462, y:6.286, w:8.409, h:0.579, al:'center' });
  band(s, { x:12.845, y:2.19, w:0.488, h:3.52, c:ORANGE });
  band(s, { x:0, y:2.19, w:0.488, h:3.52, c:DARK });
  pageNum(s, 24, { x:12.522, y:6.684, c:INK });
}

function slide25(pptx) {
  const s = newSlide(pptx);
  band(s, { x:0, y:0, w:6.508, h:7.5, c:ORANGE });
  photo(s, { x:1.352, y:1.086, w:3.72, h:5.487 });
  stripes(s, { x:11.939, y:6.263, w:1.913, c:DARK });
  para(s, T.L17 + BR + T.L18, { x:7.632, y:1.943, w:4.551, h:2.094 });
  rule(s, { x:7.723, y:1.379, w:2.28 });
  label(s, "About Our Mcokup", { x:9.949, y:1.223, w:2.26, h:0.303, c:AMBER, al:'right' });
  title(s, T.HEADLINE, { x:7.628, y:4.899, w:4.463, h:0.841 });
  kicker(s, KICKER, { x:7.628, y:6.068, w:4.274, h:0.304 });
}

function slide26(pptx) {
  const s = newSlide(pptx);
  stackedChart(s, pptx, { x:1.006, y:1.109, w:4.327, h:5.633, dir:'col', colors:[INK, MIST, ORANGE] });
  note(s, T.L11, { x:6.263, y:1.602, w:1.731, h:0.685 });
  label(s, "01. Mockup One", { x:6.263, y:1.191, w:1.731, h:0.269, sz:10 });
  note(s, T.L11, { x:8.347, y:1.602, w:1.731, h:0.685 });
  label(s, "02. Mockup Two", { x:8.347, y:1.191, w:1.731, h:0.269, sz:10, c:ORANGE });
  note(s, T.L11, { x:10.43, y:1.602, w:1.731, h:0.685 });
  label(s, "03. Mockup Three", { x:10.43, y:1.191, w:1.731, h:0.269, sz:10 });
  rule(s, { x:8.25, y:6.631, w:3.77 });
  para(s, T.L39, { x:6.263, y:2.854, w:5.897, h:1.084 });
  label(s, "About Our Mockup", { x:6.263, y:6.462, w:1.903, h:0.303, c:AMBER });
  title(s, T.HEADLINE, { x:6.263, y:5.1, w:5.757, h:0.841 });
  kicker(s, KICKER, { x:6.291, y:4.544, w:4.463, h:0.304 });
  pageNum(s, 26, { x:12.491, y:6.49, c:INK });
}

function slide27(pptx) {
  const s = newSlide(pptx);
  ring(s, pptx, { x:0.604, y:2.045, w:4.583, h:3.056 });
  ring(s, pptx, { x:4.375, y:2.632, w:4.583, h:3.056 });
  ring(s, pptx, { x:8.146, y:2.045, w:4.583, h:3.056 });
  title(s, T.HEADLINE, { x:2.462, y:0.72, w:8.409, h:0.505, sz:24, al:'center' });
  kicker(s, KICKER, { x:4.53, y:1.374, w:4.274, h:0.304, al:'center' });
  para(s, T.L09, { x:2.462, y:6.147, w:8.409, h:0.579, al:'center' });
  note(s, T.L12, { x:2.158, y:3.404, w:1.476, h:0.685, al:'center' });
  label(s, "01. Chart One", { x:2.158, y:3.056, w:1.476, h:0.269, sz:10, al:'center' });
  note(s, T.L12, { x:5.929, y:3.991, w:1.476, h:0.685, al:'center' });
  label(s, "02. Chart Two", { x:5.929, y:3.643, w:1.476, h:0.269, sz:10, c:ORANGE, al:'center' });
  note(s, T.L12, { x:9.699, y:3.404, w:1.476, h:0.685, al:'center' });
  label(s, "03. Chart Three", { x:9.699, y:3.056, w:1.476, h:0.269, sz:10, al:'center' });
  pageNum(s, 27, { x:12.491, y:6.642, c:INK });
}

function slide28(pptx) {
  const s = newSlide(pptx);
  areaChart(s, pptx, { x:6.556, y:2.403, w:5.931, h:4.227 });
  title(s, T.HEADLINE, { x:2.462, y:0.742, w:8.409, h:0.505, sz:24, al:'center' });
  kicker(s, KICKER, { x:4.53, y:1.34, w:4.274, h:0.304, al:'center' });
  para(s, T.L40 + BR + T.L41, { x:1.021, y:3.284, w:4.551, h:1.589 });
  rule(s, { x:1.112, y:2.558, w:2.554 });
  label(s, "About Our Chart", { x:3.338, y:2.403, w:2.26, h:0.303, c:AMBER, al:'right' });
  note(s, T.L20, { x:1.021, y:5.875, w:2.027, h:0.483 });
  label(s, "01. Mockup One", { x:1.021, y:5.464, w:1.552, h:0.269, sz:10 });
  note(s, T.L20, { x:3.545, y:5.875, w:2.027, h:0.483 });
  label(s, "02. Mockup Two", { x:3.545, y:5.464, w:1.552, h:0.269, sz:10 });
  pageNum(s, 28, { x:0.616, y:6.809, c:INK });
}

function slide29(pptx) {
  const s = newSlide(pptx);
  stackedChart(s, pptx, { x:0.986, y:2.061, w:11.25, h:3.189, dir:'bar', colors:[ORANGE, MIST, INK] });
  title(s, T.HEADLINE, { x:2.462, y:0.804, w:8.409, h:0.505, sz:24, al:'center' });
  kicker(s, KICKER, { x:4.53, y:1.457, w:4.274, h:0.304, al:'center' });
  note(s, T.L13, { x:2.007, y:6.225, w:2.937, h:0.483, al:'center' });
  label(s, "01. Mockup One", { x:2.007, y:5.815, w:2.937, h:0.269, sz:10, al:'center' });
  note(s, T.L13, { x:5.198, y:6.225, w:2.937, h:0.483, al:'center' });
  label(s, "02. Mockup Two", { x:5.198, y:5.815, w:2.937, h:0.269, sz:10, c:ORANGE, al:'center' });
  note(s, T.L13, { x:8.389, y:6.225, w:2.937, h:0.483, al:'center' });
  label(s, "03. Mockup Three", { x:8.389, y:5.815, w:2.937, h:0.269, sz:10, al:'center' });
  pageNum(s, 29, { x:12.491, y:6.642, c:INK });
}

function slide30(pptx) {
  const s = newSlide(pptx);
  arcPanel(s, { x:6.764, y:0, w:6.569, h:7.5, c:INK });
  label(s, "Welcome Massage", { x:0.806, y:0.864, w:1.691, h:0.303, c:AMBER });
  rule(s, { x:0.906, y:6.737, w:3.447 });
  title(s, [['T', DARK], ['H', AMBER], ['ANK', DARK]], { x:2.959, y:2.739, w:6.845, h:1.717, sz:96 });
  kicker(s, "Presentation Template", { x:2.364, y:4.282, w:4.274, h:0.326, sz:10, al:'center' });
  pageNum(s, 30, { x:12.481, y:6.552, c:WHITE });
}

/* ---------------------------------------------------------------------- run */
const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06,
  slide07, slide08, slide09, slide10, slide11, slide12,
  slide13, slide14, slide15, slide16, slide17, slide18,
  slide19, slide20, slide21, slide22, slide23, slide24,
  slide25, slide26, slide27, slide28, slide29, slide30,
];

const pptx = new PptxGenJS();
pptx.title = 'VIRIA Presentation Template';
pptx.defineLayout({ name: 'VIRIA', width: 13.333, height: 7.5 });
pptx.layout = 'VIRIA';
pptx.theme = { headFontFace: HEAD, bodyFontFace: BASE };

BUILDERS.forEach(build => build(pptx));

pptx.writeFile({ fileName: path.join(__dirname, '0bad87bd-f04e-4b14-9c86-184f50fda456_grok_final.pptx') })
  .then(f => console.log('wrote ' + f));
