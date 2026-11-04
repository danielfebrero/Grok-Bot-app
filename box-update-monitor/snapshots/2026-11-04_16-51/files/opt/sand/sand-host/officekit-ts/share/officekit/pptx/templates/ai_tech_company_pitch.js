/**
 * Ai-Tech / GE-TECH presentation template (28 slides, 13.333in x 7.5in)
 * rebuilt with pptxgenjs. Run: node <thisfile>.js
 *
 * Design language
 *   background     : grey  A7A5A6            (slide master fill)
 *   panels         : dark  333333 rounded rectangles
 *   accents/labels : lime  ADCE35
 *   outline frame  : 3pt lime rounded rectangle covering most of the slide
 *   header chrome  : hamburger bars + "GE-TECH Presentation" + "Page NN" badge
 *   fonts          : Roboto family (Light / Regular / Medium / Black), Poppins Black
 * Photos in the source deck are empty picture placeholders; they are drawn here
 * as labelled "[image]" boxes (see imgBox).
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

const OUT_NAME = '0cd3f26a-281d-494d-92ef-b52cc1b2c09e_grok_final.pptx';

const LIME = 'ADCE35';
const DARK = '333333';
const GREY = 'A7A5A6';

const F = {
  r: 'Roboto',
  l: 'Roboto Light',
  m: 'Roboto Medium',
  b: 'Roboto Black',
  pb: 'Poppins Black',
};

// --------------------------------------------------------------------------
// primitives
// --------------------------------------------------------------------------

function newSlide(pptx) {
  const s = pptx.addSlide();
  s.background = { color: GREY };
  return s;
}

/**
 * Rounded rectangle. `adj` is the OOXML corner-radius adjust value (0..50000)
 * straight from the source deck; pptxgenjs wants it in inches.
 * style: { fill } for a filled panel, { line, lw } for an outline-only frame.
 */
function rr(s, x, y, w, h, adj, style) {
  const opts = { x, y, w, h, rectRadius: (adj / 100000) * Math.min(w, h) };
  if (style.fill) opts.fill = { color: style.fill };
  if (style.line) opts.line = { color: style.line, width: style.lw };
  s.addShape('roundRect', opts);
}

/** circle (all ellipses in this deck are circles) */
function oval(s, x, y, d, color) {
  s.addShape('ellipse', { x, y, w: d, h: d, fill: { color } });
}

/** kind is 'upArrow' or 'downArrow'; flipV mirrors an up arrow into a down one */
function arrow(s, kind, x, y, w, h, color, flipV) {
  s.addShape(kind, { x, y, w, h, fill: { color }, flipV: !!flipV });
}

/** a text run: T(text, sizePt, fontFace, color, extra) */
function T(text, fontSize, fontFace, color, extra) {
  return { text, options: Object.assign({ fontSize, fontFace, color }, extra || {}) };
}

/**
 * Text box. Every text frame in the source deck uses the PowerPoint default
 * 0.1in left/right and 0.05in top/bottom insets and is top-anchored.
 * opts: { align: 'center'|'right', lineSpacing: <multiple, e.g. 1.5> }
 */
function tx(s, x, y, w, h, runs, opts) {
  opts = opts || {};
  s.addText(runs, {
    x, y, w, h,
    margin: [7.2, 7.2, 3.6, 3.6],
    valign: 'top',
    isTextBox: true,
    align: opts.align || 'left',
    lineSpacingMultiple: opts.lineSpacing || null,
  });
}

/**
 * Stand-in for the template's picture frames. The source deck ships these as
 * *empty* picture placeholders (no bitmap is stored in the file), so this is
 * drawn as a light dashed frame rather than an opaque block.
 */
function imgBox(s, x, y, w, h) {
  s.addShape('rect', { x, y, w, h, line: { color: LIME, width: 0.75, dashType: 'dash' } });
  s.addText([T('[image]', 9, F.r, LIME)], { x, y, w, h, align: 'center', valign: 'middle' });
}

// --------------------------------------------------------------------------
// repeated template furniture
// --------------------------------------------------------------------------

/** the three stacked lime "hamburger" bars used as a decorative motif */
function bars(s, x, y) {
  [0, 0.321, 0.644].forEach(function (dy) {
    rr(s, x, y + dy, 0.767, 0.153, 50000, { fill: LIME });
  });
}

/** header logo + brand + page badge, identical on every slide */
function chrome(s, page) {
  [0.267, 0.389, 0.512].forEach(function (y) {
    rr(s, 0.54, y, 0.292, 0.058, 50000, { fill: LIME });
  });
  tx(s, 0.955, 0.226, 1.037, 0.37, [T('GE-TECH Presentation', 8, F.r, DARK)]);
  rr(s, 11.74, 0.27, 1.005, 0.302, 50000, { fill: LIME });
  oval(s, 12.464, 0.309, 0.224, DARK);
  tx(s, 11.787, 0.3, 0.502, 0.236, [T('Page', 8, F.r, DARK)]);
  tx(s, 12.395, 0.3, 0.362, 0.236, [T(page, 8, F.r, GREY)], { align: 'center' });
}

// --------------------------------------------------------------------------
// charts
// --------------------------------------------------------------------------

/** two-slice ring: dark remainder + lime slice, 75% hole (slide 22) */
function doughnut(s, x, y, w, h, values) {
  s.addChart('doughnut', [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: values }], {
    x, y, w, h,
    chartColors: [DARK, LIME],
    holeSize: 75,
    firstSliceAng: 0,
    showLegend: false,
    showValue: false,
    showPercent: false,
    chartArea: { fill: { color: DARK, transparency: 100 } },
  });
}

/**
 * Slide 23 uses a "histogram" chartEx of 76 raw samples binned into six
 * intervals; pptxgenjs has no chartEx, so the same distribution is drawn as a
 * near-gapless column chart.
 */
function histogram(s, x, y, w, h) {
  const bins = ['1-4', '5-8', '9-12', '13-16', '17-20', '21-24'];
  const counts = [5, 11, 23, 24, 9, 4];
  s.addChart('bar', [{ name: 'Series1', labels: bins, values: counts }], {
    x, y, w, h,
    barDir: 'col',
    barGapWidthPct: 10,
    chartColors: [LIME],
    showLegend: false,
    showValue: false,
    catAxisHidden: true,
    valAxisHidden: true,
    valAxisMaxVal: 25,
    valAxisMajorUnit: 5,
    valGridLine: { color: GREY, size: 1 },
    catGridLine: { style: 'none' },
    chartArea: { fill: { color: DARK, transparency: 100 } },
  });
}

/**
 * Slide 24 holds an open-high-low-close stock chart. pptxgenjs has no stock
 * chart type, so the same OHLC table is redrawn as the up/down bars that a
 * stock chart shows: a transparent base column up to the lower of open/close,
 * with a lime column spanning open -> close stacked on top of it.
 */
function stockChart(s, x, y, w, h) {
  const days = ['1/1/2002', '1/2/2002', '1/3/2002', '1/4/2002', '1/5/2002'];
  const open = [44, 25, 38, 50, 60];
  const close = [25, 38, 50, 34, 45];
  const base = open.map(function (v, i) { return Math.min(v, close[i]); });
  const body = open.map(function (v, i) { return Math.abs(v - close[i]); });
  s.addChart('bar', [
    { name: 'Base', labels: days, values: base },
    { name: 'Open-Close', labels: days, values: body },
  ], {
    x, y, w, h,
    barDir: 'col',
    barGrouping: 'stacked',
    barGapWidthPct: 150,
    chartColors: [GREY, LIME],
    showLegend: false,
    showValue: false,
    catAxisHidden: true,
    valAxisMaxVal: 70,
    valAxisMajorUnit: 10,
    valAxisLabelColor: DARK,
    valAxisLabelFontFace: F.r,
    valAxisLabelFontSize: 12,
    valGridLine: { color: DARK, size: 1 },
    catGridLine: { style: 'none' },
    chartArea: { fill: { color: DARK, transparency: 100 } },
  });
}

// -------- slide 1 - Cover --------
function slide01(pptx) {
  const s = newSlide(pptx);
  chrome(s, '01');
  rr(s, 5.303, 5.61, 1.005, 0.302, 50000, { fill: LIME });
  rr(s, 0.54, 0.927, 11.494, 3.807, 24183, { fill: DARK });
  rr(s, 1.891, 5.505, 4.578, 0.512, 50000, { line: DARK, lw: 3 });
  tx(s, 1.928, 1.967, 6.871, 1.447, [T('Ai - Tech', 80, F.b, LIME)]);
  tx(s, 2.089, 5.593, 2.193, 0.337, [T('Presentation Template', 14, F.m, DARK)], { align: 'center' });
  tx(s, 7.711, 5.346, 4.578, 0.83, [
    T('Ipsum Dolor, abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle clutch speed brake Lorem Ipsum', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  tx(s, 1.992, 3.415, 2.193, 0.774, [T('Presentation Template', 20, F.m, GREY)]);
  oval(s, 6.027, 5.649, 0.224, DARK);
  bars(s, 11.636, 3.175);
  rr(s, 1.161, 1.571, 11.715, 5.058, 11002, { line: LIME, lw: 3 });
}

// -------- slide 2 - Welcome --------
function slide02(pptx) {
  const s = newSlide(pptx);
  rr(s, 0.54, 0.925, 6.123, 2.825, 24183, { fill: DARK });
  tx(s, 2.129, 5.157, 4.918, 1.083, [
    T('Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle clutch speed brake Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour glad with a', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  tx(s, 2.129, 2.033, 3.578, 1.178, [
    T('Welcome to', 32, F.l, LIME, { breakLine: true }),
    T('Ai - Tech', 32, F.b, LIME),
  ]);
  chrome(s, '02');
  rr(s, 1.262, 1.571, 11.531, 5.179, 11002, { line: LIME, lw: 3 });
  bars(s, 6.28, 2.249);
  rr(s, 11.123, 5.769, 1.005, 0.302, 50000, { fill: LIME });
  rr(s, 7.711, 5.663, 4.578, 0.512, 50000, { line: DARK, lw: 3 });
  tx(s, 7.908, 5.751, 2.193, 0.337, [T('Ai Technology', 14, F.m, DARK)]);
  oval(s, 11.847, 5.808, 0.224, DARK);
  imgBox(s, 7.76, 1.958, 4.529, 3.277);
}

// -------- slide 3 - Let's Explore --------
function slide03(pptx) {
  const s = newSlide(pptx);
  rr(s, 0.54, 0.925, 6.395, 2.825, 24183, { fill: DARK });
  rr(s, 1.044, 2.793, 11.749, 3.957, 11002, { line: LIME, lw: 3 });
  bars(s, 6.552, 1.541);
  tx(s, 7.485, 3.539, 4.918, 1.083, [
    T('Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle clutch speed brake Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour glad with a', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  tx(s, 1.617, 1.27, 4.222, 1.178, [
    T('Let’s Explore ', 32, F.l, LIME, { breakLine: true }),
    T('Our Ai Technology', 32, F.b, LIME),
  ]);
  chrome(s, '03');
  rr(s, 11.123, 5.769, 1.005, 0.302, 50000, { fill: LIME });
  rr(s, 7.485, 5.663, 4.804, 0.512, 50000, { line: DARK, lw: 3 });
  tx(s, 7.752, 5.751, 2.193, 0.337, [T('Ai Technology', 14, F.m, DARK)]);
  oval(s, 11.847, 5.808, 0.224, DARK);
  imgBox(s, 1.617, 3.323, 4.383, 3.054);
}

// -------- slide 4 - Development in the World --------
function slide04(pptx) {
  const s = newSlide(pptx);
  rr(s, 6.667, 1.526, 5.552, 2.539, 24183, { fill: DARK });
  bars(s, 11.787, 2.133);
  tx(s, 7.214, 1.911, 3.157, 1.717, [
    T('Ai Technology Development ', 32, F.l, LIME),
    T('in the World', 32, F.b, LIME),
  ]);
  tx(s, 7.214, 5.197, 4.918, 1.083, [
    T('Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle clutch speed brake Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour glad', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  tx(s, 6.667, 4.583, 1.555, 0.438, [T('1.8K Total', 20, F.m, DARK)]);
  rr(s, 0.54, 0.95, 12.252, 5.8, 11002, { line: LIME, lw: 3 });
  chrome(s, '04');
  rr(s, 4.579, 1.631, 1.005, 0.302, 50000, { fill: LIME });
  rr(s, 1.167, 1.526, 4.578, 0.512, 50000, { line: DARK, lw: 3 });
  tx(s, 1.364, 1.614, 2.193, 0.337, [T('Ai Technology', 14, F.m, DARK)]);
  oval(s, 5.303, 1.67, 0.224, DARK);
  imgBox(s, 1.201, 2.615, 4.578, 3.665);
}

// -------- slide 5 - Introducing the company --------
function slide05(pptx) {
  const s = newSlide(pptx);
  rr(s, 0.54, 0.925, 11.524, 2.552, 24183, { fill: DARK });
  rr(s, 0.955, 1.275, 11.838, 5.475, 8645, { line: LIME, lw: 3 });
  bars(s, 11.644, 1.94);
  tx(s, 1.603, 1.749, 4.186, 1.178, [
    T('Introducing, We Are ', 32, F.l, LIME, { breakLine: true }),
    T('Ai-Tech Company', 32, F.b, LIME),
  ]);
  tx(s, 6.667, 1.749, 4.249, 1.083, [
    T('Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle clutch speed brake Lorem Ipsum Dolor Sit Amet, abstract feel guest', 10, F.r, GREY),
  ], { lineSpacing: 1.5 });
  arrow(s, 'upArrow', 6.518, 4.188, 0.674, 0.559, LIME);
  tx(s, 6.518, 4.904, 1.291, 0.438, [T('6.789K', 20, F.m, DARK)]);
  tx(s, 6.518, 5.456, 2.312, 0.757, [
    T('Lorem Ipsum Dolor Sit ', 9, F.r, DARK),
    T('Amet', 9, F.r, DARK),
    T(', abstract feel guest trust at journey work tour glad with another monthly', 9, F.r, DARK),
  ], { lineSpacing: 1.5 });
  arrow(s, 'upArrow', 9.332, 4.233, 0.674, 0.559, LIME, true);
  tx(s, 9.332, 4.904, 1.291, 0.438, [T('6.789K', 20, F.m, DARK)]);
  tx(s, 9.332, 5.454, 2.312, 0.757, [
    T('Lorem Ipsum Dolor Sit ', 9, F.r, DARK),
    T('Amet', 9, F.r, DARK),
    T(', abstract feel guest trust at journey work tour glad with another monthly', 9, F.r, DARK),
  ], { lineSpacing: 1.5 });
  chrome(s, '05');
  imgBox(s, 1.602, 3.952, 4.092, 2.262);
}

// -------- slide 6 - Business fields --------
function slide06(pptx) {
  const s = newSlide(pptx);
  rr(s, 0.54, 0.925, 6.495, 5.626, 15340, { fill: DARK });
  rr(s, 6.298, 1.571, 6.495, 4.274, 11002, { line: LIME, lw: 3 });
  bars(s, 6.667, 3.314);
  tx(s, 1.178, 1.545, 3.745, 1.178, [
    T('Ai -Tech', 32, F.l, LIME, { breakLine: true }),
    T('Business Fields', 32, F.b, LIME),
  ]);
  tx(s, 1.183, 3.756, 4.287, 1.083, [
    T('Lorem Ipsum Dolor Sit journey work tour glad with monthly chair stair class table board ink erase caps enter the grip throttle clutch speed brake Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour', 10, F.r, GREY),
  ], { lineSpacing: 1.5 });
  tx(s, 1.183, 3.242, 2.904, 0.438, [T('About Our Business', 20, F.m, GREY)]);
  chrome(s, '06');
  rr(s, 4.31, 5.417, 1.005, 0.302, 50000, { fill: LIME });
  rr(s, 1.178, 5.312, 4.287, 0.512, 50000, { line: GREY, lw: 3 });
  tx(s, 1.376, 5.4, 2.193, 0.337, [T('Ai Technology', 14, F.m, GREY)]);
  oval(s, 5.034, 5.456, 0.224, DARK);
  imgBox(s, 7.863, 1.905, 4.292, 4.646);
}

// -------- slide 7 - Market value --------
function slide07(pptx) {
  const s = newSlide(pptx);
  rr(s, 6.097, 0.925, 6.696, 5.626, 15340, { fill: DARK });
  rr(s, 0.589, 1.378, 11.635, 4.684, 11002, { line: LIME, lw: 3 });
  bars(s, 5.706, 4.715);
  tx(s, 6.55, 1.782, 5.052, 1.178, [
    T('Market Value for', 32, F.l, LIME, { breakLine: true }),
    T('Business Developments', 32, F.b, LIME),
  ]);
  tx(s, 7.057, 4.515, 4.611, 1.083, [
    T('Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle clutch speed brake Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work', 10, F.r, GREY),
  ], { lineSpacing: 1.5 });
  oval(s, 1.038, 1.91, 1.312, DARK);
  tx(s, 1.11, 2.246, 1.183, 0.64, [T('72%', 32, F.b, LIME)], { align: 'center' });
  oval(s, 1.038, 3.913, 1.312, DARK);
  tx(s, 1.11, 4.249, 1.183, 0.64, [T('80%', 32, F.b, LIME)], { align: 'center' });
  tx(s, 0.924, 5.344, 1.555, 0.337, [T('2021', 14, F.m, DARK)], { align: 'center' });
  tx(s, 0.911, 3.413, 1.555, 0.337, [T('2022', 14, F.m, DARK)], { align: 'center' });
  oval(s, 6.687, 3.471, 0.138, LIME);
  tx(s, 7.057, 3.373, 3.601, 0.337, [
    T('$ 45.678.900     ', 14, F.m, GREY),
    T('for Abstract Designs', 14, F.l, GREY),
  ]);
  oval(s, 6.687, 3.875, 0.138, LIME);
  tx(s, 7.057, 3.777, 3.476, 0.337, [T('$ 45.678.900     ', 14, F.m, GREY), T('for 3D Designs', 14, F.l, GREY)]);
  chrome(s, '07');
  imgBox(s, 2.788, 1.869, 2.669, 3.728);
}

// -------- slide 8 - Our Product --------
function slide08(pptx) {
  const s = newSlide(pptx);
  rr(s, 1.056, 0.925, 11.233, 2.552, 19829, { fill: DARK });
  rr(s, 0.54, 1.306, 12.204, 5.444, 8645, { line: LIME, lw: 3 });
  tx(s, 4.567, 1.562, 4.198, 0.64, [T('Our ', 32, F.l, LIME), T('Product', 32, F.b, LIME)], { align: 'center' });
  tx(s, 1.5, 5.562, 2.247, 0.83, [
    T('PLACEHOLDER', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  tx(s, 4.106, 5.562, 2.381, 0.83, [
    T('PLACEHOLDER', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  tx(s, 6.846, 5.562, 2.381, 0.83, [
    T('PLACEHOLDER', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  tx(s, 9.587, 5.562, 2.247, 0.83, [
    T('Lorem Ipsum Dolor Sit, abstract feel guest trust at journey work tour glad with another monthly', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  tx(s, 1.5, 5.125, 0.978, 0.438, [T('01.', 20, F.m, DARK)]);
  tx(s, 4.195, 5.081, 0.978, 0.438, [T('02.', 20, F.m, DARK)]);
  tx(s, 6.891, 5.081, 0.978, 0.438, [T('03.', 20, F.m, DARK)]);
  tx(s, 9.587, 5.081, 0.978, 0.438, [T('04.', 20, F.m, DARK)]);
  chrome(s, '08');
  imgBox(s, 1.527, 2.429, 2.22, 2.366);
  imgBox(s, 4.213, 2.429, 2.22, 2.366);
  imgBox(s, 6.9, 2.429, 2.22, 2.366);
  imgBox(s, 9.587, 2.429, 2.22, 2.366);
}

// -------- slide 9 - No similar product --------
function slide09(pptx) {
  const s = newSlide(pptx);
  rr(s, 0.54, 0.925, 6.123, 2.825, 24183, { fill: DARK });
  rr(s, 0.955, 1.324, 11.838, 5.426, 11002, { line: LIME, lw: 3 });
  bars(s, 6.28, 1.939);
  rr(s, 4.878, 4.264, 1.005, 0.302, 50000, { fill: LIME });
  rr(s, 1.467, 4.159, 4.578, 0.512, 50000, { line: DARK, lw: 3 });
  tx(s, 1.664, 4.247, 2.193, 0.337, [T('Ai Technology', 14, F.m, DARK)]);
  oval(s, 5.603, 4.303, 0.224, DARK);
  tx(s, 1.526, 1.749, 4.182, 1.178, [
    T('There are no', 32, F.l, LIME, { breakLine: true }),
    T('Similar Product', 32, F.pb, LIME),
  ]);
  tx(s, 10.336, 2.386, 2.068, 0.83, [
    T('Lorem Ipsum Dolor Sit ', 10, F.r, DARK),
    T('Amet', 10, F.r, DARK),
    T(', abstract feel guest trust at journey work tour', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  tx(s, 10.336, 1.84, 1.591, 0.438, [T('Popular', 20, F.m, DARK)]);
  tx(s, 10.336, 5.268, 2.068, 0.83, [
    T('Lorem Ipsum Dolor Sit ', 10, F.r, DARK),
    T('Amet', 10, F.r, DARK),
    T(', abstract feel guest trust at journey work tour', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  tx(s, 10.336, 4.722, 1.953, 0.438, [T('Best Selling', 20, F.m, DARK)]);
  tx(s, 1.467, 5.016, 4.918, 1.083, [
    T('Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle clutch speed brake Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour glad with another', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  imgBox(s, 7.577, 4.722, 2.361, 1.377);
  chrome(s, '09');
  imgBox(s, 7.577, 1.841, 2.361, 2.609);
}

// -------- slide 10 - Product quality --------
function slide10(pptx) {
  const s = newSlide(pptx);
  rr(s, 5.778, 1.358, 6.441, 1.697, 24183, { fill: DARK });
  bars(s, 11.787, 1.815);
  rr(s, 0.54, 0.961, 12.252, 5.789, 11002, { line: LIME, lw: 3 });
  tx(s, 6.271, 1.624, 4.523, 1.178, [
    T('Product Quality', 32, F.l, LIME, { breakLine: true }),
    T('Assurance', 32, F.b, LIME),
  ]);
  tx(s, 6.255, 3.869, 5.532, 0.83, [
    T('Lorem Ipsum Dolor Sit ', 10, F.r, DARK),
    T('Amet', 10, F.r, DARK),
    T(', abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle clutch speed brake Lorem Ipsum Dolor Sit ', 10, F.r, DARK),
    T('Amet', 10, F.r, DARK),
    T(', abstract feel', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  tx(s, 6.255, 3.42, 1.843, 0.438, [T('2020 - 2021', 20, F.m, DARK)]);
  tx(s, 6.279, 4.966, 1.843, 0.438, [T('2021 - 2022', 20, F.m, DARK)]);
  chrome(s, '10');
  tx(s, 6.255, 5.444, 5.532, 0.83, [
    T('Lorem Ipsum Dolor Sit ', 10, F.r, DARK),
    T('Amet', 10, F.r, DARK),
    T(', abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle clutch speed brake Lorem Ipsum Dolor Sit ', 10, F.r, DARK),
    T('Amet', 10, F.r, DARK),
    T(', abstract feel', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  rr(s, 4.02, 3.686, 1.005, 0.302, 50000, { fill: LIME });
  rr(s, 1.324, 3.581, 3.859, 0.512, 50000, { line: DARK, lw: 3 });
  tx(s, 1.521, 3.669, 2.193, 0.337, [T('Ai Technology Story', 14, F.m, DARK)]);
  oval(s, 4.744, 3.725, 0.224, DARK);
  imgBox(s, 1.324, 1.378, 3.859, 1.946);
  imgBox(s, 1.324, 4.324, 3.859, 1.946);
}

// -------- slide 11 - Make a choice --------
function slide11(pptx) {
  const s = newSlide(pptx);
  imgBox(s, 8.159, 1.782, 3.572, 4.517);
  rr(s, 0.54, 1.434, 6.123, 1.906, 24183, { fill: DARK });
  rr(s, 1.113, 0.925, 11.68, 5.825, 11002, { line: LIME, lw: 3 });
  bars(s, 6.28, 1.939);
  tx(s, 1.619, 1.782, 3.549, 1.178, [
    T('Make ', 32, F.l, LIME),
    T('a Choice ', 32, F.l, LIME),
    T('As Needed', 32, F.b, LIME),
  ]);
  tx(s, 1.537, 5.216, 5.51, 1.083, [
    T('Lorem Ipsum Dolor Sit ', 10, F.r, DARK),
    T('Amet', 10, F.r, DARK),
    T(', abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle clutch speed brake Lorem Ipsum Dolor Sit ', 10, F.r, DARK),
    T('Amet', 10, F.r, DARK),
    T(', abstract feel guest trust at journey work tour glad with another monthly chair stair class table boar', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  chrome(s, '11');
  rr(s, 5.031, 4.264, 1.005, 0.302, 50000, { fill: LIME });
  rr(s, 1.619, 4.159, 4.578, 0.512, 50000, { line: DARK, lw: 3 });
  tx(s, 1.816, 4.247, 2.193, 0.337, [T('Ai Technology', 14, F.m, DARK)]);
  oval(s, 5.755, 4.303, 0.224, DARK);
  oval(s, 11.151, 1.244, 1.312, DARK);
  tx(s, 11.223, 1.581, 1.183, 0.64, [T('96%', 32, F.b, LIME)], { align: 'center' });
}

// -------- slide 12 - Technology asset advantage --------
function slide12(pptx) {
  const s = newSlide(pptx);
  rr(s, 1.306, 0.925, 10.726, 2.552, 19829, { fill: DARK });
  rr(s, 0.54, 2.41, 12.204, 4.34, 8645, { line: LIME, lw: 3 });
  chrome(s, '12');
  tx(s, 4.405, 1.05, 4.523, 1.178, [
    T('Technology Asset', 32, F.l, LIME, { breakLine: true }),
    T('Advantage', 32, F.b, LIME),
  ], { align: 'center' });
  tx(s, 2.566, 5.226, 2.738, 0.438, [T('Competitive Price', 20, F.m, DARK)], { align: 'center' });
  tx(s, 1.91, 5.694, 4.313, 0.83, [
    T('Lorem Ipsum Dolor Sit ', 10, F.r, DARK),
    T('Amet', 10, F.r, DARK),
    T(', abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle', 10, F.r, DARK),
  ], { align: 'center', lineSpacing: 1.5 });
  tx(s, 8.031, 5.226, 2.738, 0.438, [T('Build Community', 20, F.m, DARK)], { align: 'center' });
  tx(s, 7.109, 5.694, 4.313, 0.83, [
    T('Lorem Ipsum Dolor Sit ', 10, F.r, DARK),
    T('Amet', 10, F.r, DARK),
    T(', abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle', 10, F.r, DARK),
  ], { align: 'center', lineSpacing: 1.5 });
  imgBox(s, 7.109, 2.681, 4.315, 2.343);
  imgBox(s, 1.91, 2.681, 4.315, 2.343);
}

// -------- slide 13 - Guide to order --------
function slide13(pptx) {
  const s = newSlide(pptx);
  rr(s, 0.54, 1.379, 7.126, 3.599, 21504, { fill: DARK });
  rr(s, 1.113, 0.925, 11.68, 5.825, 11002, { line: LIME, lw: 3 });
  bars(s, 7.283, 1.835);
  tx(s, 1.717, 1.674, 4.68, 1.178, [
    T('Guide to order at ', 32, F.l, LIME, { breakLine: true }),
    T('Ai–Tech', 32, F.b, LIME),
  ]);
  tx(s, 1.717, 5.228, 1.357, 0.438, [T('Step 1', 20, F.m, DARK)]);
  tx(s, 1.717, 5.665, 2.574, 0.578, [
    T('Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  tx(s, 5.04, 5.228, 1.357, 0.438, [T('Step 2', 20, F.m, DARK)]);
  tx(s, 5.04, 5.67, 2.541, 0.578, [
    T('Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  tx(s, 1.717, 3.875, 5.38, 0.83, [
    T('Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle clutch speed The brake Lorem Ipsum Dolor Sit abstract', 10, F.r, GREY),
  ], { lineSpacing: 1.5 });
  tx(s, 1.717, 3.349, 2.542, 0.438, [T('2 Easy Steps', 20, F.m, GREY)]);
  chrome(s, '13');
  imgBox(s, 8.378, 1.379, 3.703, 4.895);
}

// -------- slide 14 - Increase in market demand --------
function slide14(pptx) {
  const s = newSlide(pptx);
  rr(s, 5.468, 0.961, 6.714, 2.128, 24183, { fill: DARK });
  bars(s, 5.12, 1.815);
  rr(s, 0.54, 1.392, 12.225, 5.358, 11002, { line: LIME, lw: 3 });
  chrome(s, '14');
  tx(s, 6.274, 1.623, 3.582, 1.178, [T('Increase in ', 32, F.l, LIME), T('Market Demand', 32, F.b, LIME)]);
  tx(s, 5.468, 3.743, 1.764, 0.337, [T('2020 - 2021', 14, F.m, DARK)]);
  tx(s, 5.468, 4.645, 1.764, 0.337, [T('2021 - 2022', 14, F.m, DARK)]);
  tx(s, 10.904, 3.76, 0.827, 0.337, [T('63%', 14, F.m, DARK)], { align: 'right' });
  tx(s, 10.904, 4.637, 0.827, 0.337, [T('86%', 14, F.m, DARK)], { align: 'right' });
  rr(s, 7.108, 3.76, 2.609, 0.302, 50000, { fill: LIME });
  rr(s, 6.942, 3.655, 4.852, 0.512, 50000, { line: DARK, lw: 3 });
  oval(s, 9.446, 3.799, 0.224, DARK);
  rr(s, 7.108, 4.637, 3.631, 0.302, 50000, { fill: LIME });
  rr(s, 6.942, 4.532, 4.852, 0.512, 50000, { line: DARK, lw: 3 });
  oval(s, 10.448, 4.676, 0.224, DARK);
  tx(s, 5.504, 5.46, 6.678, 0.83, [
    T('Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour glad with another monthly of chair stair class table board ink erase caps enter the grip throttle clutch speed brake Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  imgBox(s, 1.151, 1.858, 3.381, 2.119);
  imgBox(s, 1.151, 4.171, 3.381, 2.119);
}

// -------- slide 15 - Three best-selling products --------
function slide15(pptx) {
  const s = newSlide(pptx);
  oval(s, 9.075, 1.784, 1.312, DARK);
  rr(s, 0.604, 0.925, 5.657, 5.825, 13095, { fill: DARK });
  rr(s, 0.955, 1.317, 11.838, 5.016, 11002, { line: LIME, lw: 3 });
  bars(s, 5.869, 1.939);
  rr(s, 4.395, 3.928, 1.005, 0.302, 50000, { fill: LIME });
  rr(s, 1.523, 3.823, 4.019, 0.512, 50000, { line: GREY, lw: 3 });
  tx(s, 1.72, 3.911, 2.193, 0.337, [T('Ai Technology', 14, F.m, GREY)]);
  oval(s, 5.119, 3.967, 0.224, DARK);
  tx(s, 1.473, 1.705, 4.221, 1.717, [
    T('Three Best-Selling ', 32, F.l, LIME),
    T('Technology Product', 32, F.b, LIME),
  ]);
  tx(s, 1.448, 4.744, 4.246, 1.083, [
    T('Lorem Ipsum Dolor Sit ', 10, F.r, GREY),
    T('Amet', 10, F.r, GREY),
    T(', abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle clutch speed brake Lorem Ipsum Dolor Sit ', 10, F.r, GREY),
    T('Amet', 10, F.r, GREY),
    T(', abstract feel guest trust', 10, F.r, GREY),
  ], { lineSpacing: 1.5 });
  tx(s, 9.14, 2.152, 1.183, 0.64, [T('71+', 32, F.b, LIME)], { align: 'center' });
  tx(s, 8.987, 3.149, 1.487, 0.337, [T('2024', 14, F.m, DARK)], { align: 'center' });
  chrome(s, '15');
  oval(s, 7.072, 4.66, 1.312, DARK);
  tx(s, 7.137, 5.027, 1.183, 0.64, [T('46+', 32, F.b, LIME)], { align: 'center' });
  tx(s, 6.985, 4.279, 1.487, 0.337, [T('2023', 14, F.m, DARK)], { align: 'center' });
  oval(s, 10.915, 4.66, 1.312, DARK);
  tx(s, 10.98, 5.027, 1.183, 0.64, [T('85+', 32, F.b, LIME)], { align: 'center' });
  tx(s, 10.828, 4.279, 1.487, 0.337, [T('2025', 14, F.m, DARK)], { align: 'center' });
  imgBox(s, 6.979, 1.794, 1.667, 2.383);
  imgBox(s, 10.741, 1.794, 1.667, 2.383);
  imgBox(s, 8.86, 3.528, 1.667, 2.383);
}

// -------- slide 16 - Visit us on the website --------
function slide16(pptx) {
  const s = newSlide(pptx);
  rr(s, 0.54, 1.383, 11.524, 2.094, 24183, { fill: DARK });
  rr(s, 1.079, 0.925, 11.714, 5.825, 7010, { line: LIME, lw: 3 });
  bars(s, 11.661, 2.032);
  tx(s, 1.529, 1.841, 3.032, 1.178, [T('Visit Us on ', 32, F.l, LIME), T('The Website', 32, F.b, LIME)]);
  tx(s, 1.529, 4.277, 4.188, 1.84, [
    T('Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle clutch', 10, F.r, DARK, { breakLine: true }),
    T('', 10, F.r, DARK, { breakLine: true }),
    T('Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle clutch', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  chrome(s, '16');
  rr(s, 10.079, 5.656, 1.005, 0.302, 50000, { fill: LIME });
  rr(s, 6.667, 5.551, 4.578, 0.512, 50000, { line: DARK, lw: 3 });
  tx(s, 6.864, 5.639, 2.193, 0.337, [T('Ai – Tech Website', 14, F.m, DARK)]);
  oval(s, 10.803, 5.695, 0.224, DARK);
  imgBox(s, 6.667, 1.949, 4.531, 2.935);
}

// -------- slide 17 - Install the app --------
function slide17(pptx) {
  const s = newSlide(pptx);
  rr(s, 6.052, 0.967, 6.761, 2.32, 24183, { fill: DARK });
  bars(s, 5.608, 1.84);
  rr(s, 0.54, 1.392, 11.863, 5.358, 11002, { line: LIME, lw: 3 });
  tx(s, 4.38, 4.049, 1.225, 1.083, [
    T('Lorem Ipsum Dolor Sit Amet, abstract feel guest', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  tx(s, 6.719, 1.649, 2.125, 1.178, [T('Install ', 32, F.l, LIME), T('The App', 32, F.b, LIME)]);
  tx(s, 6.719, 4.049, 4.82, 2.092, [
    T('Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle clutch speed brake Lorem Ipsum Dolor Sit', 10, F.r, DARK, { breakLine: true }),
    T('', 10, F.r, DARK, { breakLine: true }),
    T('Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle clutch speed brake Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour glad with another', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  chrome(s, '17');
  rr(s, 4.6, 5.669, 1.005, 0.302, 50000, { fill: LIME });
  rr(s, 1.188, 5.563, 4.578, 0.512, 50000, { line: DARK, lw: 3 });
  tx(s, 1.385, 5.651, 2.193, 0.337, [T('Ai – Tech App', 14, F.m, DARK)]);
  oval(s, 5.324, 5.707, 0.224, DARK);
  imgBox(s, 1.335, 1.815, 2.771, 3.271);
}

// -------- slide 18 - Break slide --------
function slide18(pptx) {
  const s = newSlide(pptx);
  chrome(s, '18');
  rr(s, 5.303, 5.174, 1.005, 0.302, 50000, { fill: LIME });
  rr(s, 0.54, 1.483, 11.494, 2.391, 24183, { fill: DARK });
  rr(s, 1.891, 5.068, 4.578, 0.512, 50000, { line: DARK, lw: 3 });
  tx(s, 2.089, 5.156, 2.193, 0.337, [T('Presentation Template', 14, F.m, DARK)], { align: 'center' });
  tx(s, 7.711, 4.909, 4.578, 0.83, [
    T('Ipsum Dolor, abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle clutch speed brake Lorem Ipsum', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  oval(s, 6.027, 5.212, 0.224, DARK);
  bars(s, 11.636, 2.28);
  rr(s, 1.161, 0.927, 11.715, 5.702, 11002, { line: LIME, lw: 3 });
  tx(s, 1.868, 1.955, 6.871, 1.447, [T('Break Slide', 80, F.b, LIME)]);
}

// -------- slide 19 - Quotes --------
function slide19(pptx) {
  const s = newSlide(pptx);
  rr(s, 0.955, 1.434, 4.592, 1.906, 24183, { fill: DARK });
  rr(s, 4.683, 2.583, 8.11, 4.167, 11002, { line: LIME, lw: 3 });
  bars(s, 0.54, 1.939);
  tx(s, 5.789, 3.933, 5.805, 1.784, [
    T('Technology is nothing. The important thing is that you have faith in other people, that they are basically good and smart, and if you give them tools, they will do amazing things with those tools.    ', 20, F.l, DARK, { italic: true }),
    T('-Steve Jobs', 20, F.l, DARK, { bold: true }),
  ], { align: 'center' });
  tx(s, 1.756, 1.748, 1.895, 1.178, [T('Ai -Tech ', 32, F.l, LIME), T('Quotes', 32, F.b, LIME)]);
  tx(s, 5.642, 3.549, 0.381, 0.64, [T('“', 32, F.pb, DARK)], { align: 'center' });
  tx(s, 11.359, 5.5, 0.381, 0.64, [T('“', 32, F.pb, DARK)], { align: 'center' });
  chrome(s, '19');
}

// -------- slide 20 - About our service --------
function slide20(pptx) {
  const s = newSlide(pptx);
  rr(s, 1.161, 0.925, 5.1, 5.825, 13095, { fill: DARK });
  rr(s, 0.604, 1.317, 12.189, 4.995, 11002, { line: LIME, lw: 3 });
  bars(s, 5.869, 1.939);
  tx(s, 1.772, 1.748, 3.623, 1.178, [T('About Our ', 32, F.l, LIME), T('Service', 32, F.b, LIME)]);
  tx(s, 1.772, 3.635, 3.955, 1.083, [
    T('Lorem Ipsum Dolor Sit t, abstract feel guest work tour glad with another monthly chair stair class table board ink erase caps enter throttle clutch speed brake Lorem Ipsum Dolor abstract feel guest trust at journey work tour glad with', 10, F.r, GREY),
  ], { lineSpacing: 1.5 });
  chrome(s, '20');
  rr(s, 4.395, 5.345, 1.005, 0.302, 50000, { fill: LIME });
  rr(s, 1.772, 5.24, 3.771, 0.512, 50000, { line: GREY, lw: 3 });
  tx(s, 1.923, 5.328, 2.193, 0.337, [T('Ai Technology', 14, F.m, GREY)]);
  oval(s, 5.119, 5.384, 0.224, DARK);
  imgBox(s, 7.333, 1.748, 4.635, 4.125);
}

// -------- slide 21 - After-sales service --------
function slide21(pptx) {
  const s = newSlide(pptx);
  rr(s, 6.444, 1.469, 5.959, 2.263, 20676, { fill: DARK });
  bars(s, 6.103, 2.23);
  rr(s, 0.54, 0.992, 12.307, 5.758, 11002, { line: LIME, lw: 3 });
  chrome(s, '21');
  tx(s, 7.212, 1.971, 3.623, 1.178, [T('Our After-sales ', 32, F.l, LIME), T('Service', 32, F.b, LIME)]);
  tx(s, 6.444, 5.106, 5.845, 1.083, [
    T('Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle clutch speed ', 10, F.r, DARK),
    T('yhe', 10, F.r, DARK),
    T(' brake Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour glad board ink erase caps enter the grip throttle clutch speed ', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  tx(s, 1.118, 4.061, 2.086, 0.774, [T('Trending Asset', 20, F.m, DARK)]);
  tx(s, 1.118, 5.106, 1.703, 1.083, [
    T('PLACEHOLDER', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  rr(s, 10.849, 4.32, 1.005, 0.302, 50000, { fill: LIME });
  rr(s, 6.556, 4.215, 5.46, 0.512, 50000, { line: DARK, lw: 3 });
  tx(s, 6.694, 4.303, 2.193, 0.337, [T('Ai Technology', 14, F.m, DARK)]);
  oval(s, 11.574, 4.359, 0.224, DARK);
  imgBox(s, 1.128, 1.487, 2.178, 2.263);
  imgBox(s, 3.54, 1.469, 2.178, 2.281);
  imgBox(s, 3.54, 3.993, 2.178, 2.263);
}

// -------- slide 22 - Infographic (doughnut charts) --------
function slide22(pptx) {
  const s = newSlide(pptx);
  rr(s, 0.955, 1.309, 4.905, 2.264, 20678, { fill: DARK });
  rr(s, 0.54, 0.917, 12.252, 5.833, 11002, { line: LIME, lw: 3 });
  bars(s, 5.454, 2.001);
  chrome(s, '22');
  tx(s, 1.473, 1.77, 3.792, 1.178, [
    T('Infographic', 32, F.l, LIME, { breakLine: true }),
    T('Ai Technology', 32, F.b, LIME),
  ]);
  doughnut(s, 3.769, 3.927, 3.756, 2.504, [8.2, 3.2]);  // ring behind the 73% label
  doughnut(s, 6.234, 1.253, 3.756, 2.504, [9.2, 2.2]);  // ring behind the 87% label
  tx(s, 1.473, 4.662, 1.883, 0.438, [T('Sales Data', 20, F.m, DARK)]);
  tx(s, 1.473, 5.143, 2.55, 0.83, [
    T('Lorem Ipsum Dolor Sit ', 10, F.r, DARK),
    T('Amet', 10, F.r, DARK),
    T(', abstract feel guest trust at journey work tour glad with another monthly', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  tx(s, 7.521, 2.185, 1.183, 0.64, [T('87%', 32, F.b, DARK)], { align: 'center' });
  tx(s, 9.646, 2.013, 1.999, 0.438, [T('Market Value', 20, F.m, DARK)]);
  tx(s, 9.646, 2.453, 2.643, 0.83, [
    T('Lorem Ipsum Dolor Sit ', 10, F.r, DARK),
    T('Amet', 10, F.r, DARK),
    T(', abstract feel guest trust at journey work tour glad with another monthly', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  tx(s, 5.054, 4.837, 1.183, 0.64, [T('73%', 32, F.b, DARK)], { align: 'center' });
  tx(s, 7.521, 5.143, 4.59, 0.83, [
    T('Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter feel guest trust at journey work tour glad with another', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  tx(s, 7.521, 4.618, 1.659, 0.438, [T('Statistic', 20, F.m, DARK)]);
  arrow(s, 'upArrow', 1.473, 3.966, 0.674, 0.559, LIME);
  arrow(s, 'upArrow', 9.691, 1.327, 0.674, 0.559, LIME);
}

// -------- slide 23 - Monthly progress (histogram) --------
function slide23(pptx) {
  const s = newSlide(pptx);
  rr(s, 0.538, 0.99, 11.925, 3.855, 14779, { fill: DARK });
  bars(s, 12.082, 1.995);
  rr(s, 0.922, 1.359, 5.887, 5.391, 7395, { line: LIME, lw: 3 });
  chrome(s, '23');
  histogram(s, 1.281, 1.781, 4.94, 2.738);
  tx(s, 7.468, 1.803, 3.567, 1.178, [
    T('Monthly ', 32, F.l, LIME),
    T('Progress', 32, F.b, LIME),
    T(' ', 32, F.l, LIME),
    T('Chart', 32, F.b, LIME),
  ]);
  tx(s, 1.755, 5.185, 1.975, 0.337, [T('July ', 14, F.m, DARK), T('2022', 14, F.l, DARK)]);
  tx(s, 1.755, 5.607, 1.975, 0.337, [T('August ', 14, F.m, DARK), T('2022', 14, F.l, DARK)]);
  tx(s, 1.755, 6.03, 1.975, 0.337, [T('September ', 14, F.m, DARK), T('2022', 14, F.l, DARK)]);
  tx(s, 4.505, 5.185, 1.975, 0.337, [T('October', 14, F.m, DARK), T(' 2022', 14, F.l, DARK)]);
  tx(s, 4.505, 5.607, 1.975, 0.337, [T('November ', 14, F.m, DARK), T('2022', 14, F.l, DARK)]);
  arrow(s, 'downArrow', 1.323, 5.231, 0.296, 0.251, DARK);
  arrow(s, 'upArrow', 1.323, 5.65, 0.296, 0.251, LIME);
  arrow(s, 'upArrow', 1.323, 6.07, 0.296, 0.251, LIME);
  arrow(s, 'upArrow', 4.077, 5.196, 0.296, 0.251, LIME);
  arrow(s, 'downArrow', 4.07, 5.63, 0.296, 0.251, DARK);
  tx(s, 4.505, 6.027, 1.975, 0.337, [T('December ', 14, F.m, DARK), T('2022', 14, F.l, DARK)]);
  arrow(s, 'downArrow', 4.07, 6.06, 0.296, 0.251, DARK);
  tx(s, 7.468, 5.57, 4.319, 0.83, [
    T('Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle clutch speed', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  tx(s, 7.468, 5.077, 1.843, 0.438, [T('This Year', 20, F.m, DARK)]);
  rr(s, 10.471, 4.051, 1.005, 0.302, 50000, { fill: LIME });
  rr(s, 7.468, 3.945, 4.116, 0.512, 50000, { line: GREY, lw: 3 });
  tx(s, 7.619, 4.033, 2.193, 0.337, [T('Ai Technology', 14, F.m, GREY)]);
  oval(s, 11.195, 4.09, 0.224, DARK);
}

// -------- slide 24 - This year's earning (stock chart) --------
function slide24(pptx) {
  const s = newSlide(pptx);
  rr(s, 0.589, 1.359, 5.278, 2.438, 22155, { fill: DARK });
  chrome(s, '24');
  tx(s, 1.702, 1.686, 3.623, 1.717, [T('This Year’s ', 32, F.l, LIME), T('Earning Infographic', 32, F.b, LIME)]);
  stockChart(s, 6.775, 1.359, 5.352, 3.865);
  oval(s, 1.739, 4.332, 0.138, LIME);
  tx(s, 2.109, 4.234, 3.601, 0.337, [T('$ 43.678.900     ', 14, F.m, DARK), T('August 2022', 14, F.l, DARK)]);
  oval(s, 1.739, 4.736, 0.138, LIME);
  tx(s, 2.109, 4.639, 3.476, 0.337, [T('$ 38.577.900     ', 14, F.m, DARK), T('September 2022', 14, F.l, DARK)]);
  oval(s, 1.739, 5.171, 0.138, LIME);
  tx(s, 2.109, 5.073, 3.601, 0.337, [T('$ 50.633.900     ', 14, F.m, DARK), T('October 2022', 14, F.l, DARK)]);
  oval(s, 1.739, 5.575, 0.138, LIME);
  tx(s, 2.109, 5.477, 3.476, 0.337, [T('$ 49.834.900     ', 14, F.m, DARK), T('November 2022', 14, F.l, DARK)]);
  oval(s, 1.739, 6.009, 0.138, LIME);
  tx(s, 2.109, 5.912, 3.476, 0.337, [T('$ 60.452.200     ', 14, F.m, DARK), T('December 2022', 14, F.l, DARK)]);
  rr(s, 1.044, 0.936, 11.7, 5.814, 8645, { line: LIME, lw: 3 });
  rr(s, 11.011, 5.728, 1.005, 0.302, 50000, { fill: LIME });
  rr(s, 6.829, 5.623, 5.298, 0.512, 50000, { line: DARK, lw: 3 });
  tx(s, 6.967, 5.711, 2.193, 0.337, [T('Ai Technology', 14, F.m, DARK)]);
  oval(s, 11.736, 5.767, 0.224, DARK);
  bars(s, 5.483, 1.834);
}

// -------- slide 25 - Team member --------
function slide25(pptx) {
  const s = newSlide(pptx);
  rr(s, 6.129, 0.925, 6.16, 5.284, 11569, { fill: DARK });
  tx(s, 6.835, 2.02, 4.471, 0.64, [T('Michael ', 32, F.l, LIME), T('J. Gutenberg', 32, F.b, LIME)]);
  tx(s, 6.835, 4.177, 4.855, 1.083, [
    T('Lorem Ipsum Dolor Sit ', 10, F.r, GREY),
    T('Amet', 10, F.r, GREY),
    T(', abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle clutch speed brake Lorem Ipsum Dolor Sit ', 10, F.r, GREY),
    T('Amet', 10, F.r, GREY),
    T(', abstract feel guest trust at journey work tour glad with another', 10, F.r, GREY),
  ], { lineSpacing: 1.5 });
  chrome(s, '25');
  rr(s, 0.604, 1.419, 12.189, 5.331, 11002, { line: LIME, lw: 3 });
  bars(s, 5.731, 1.939);
  rr(s, 4.118, 5.816, 1.005, 0.302, 50000, { fill: LIME });
  rr(s, 1.219, 5.711, 4.035, 0.512, 50000, { line: DARK, lw: 3 });
  tx(s, 1.416, 5.799, 2.193, 0.337, [T('CEO Founder', 14, F.m, DARK)]);
  oval(s, 4.842, 5.855, 0.224, DARK);
  imgBox(s, 1.244, 1.939, 4.01, 3.321);
}

// -------- slide 26 - Human resource --------
function slide26(pptx) {
  const s = newSlide(pptx);
  rr(s, 0.54, 0.917, 6.265, 2.617, 22821, { fill: DARK });
  rr(s, 1.028, 2.687, 11.765, 4.063, 11002, { line: LIME, lw: 3 });
  bars(s, 6.321, 1.406);
  tx(s, 1.328, 1.482, 3.942, 0.64, [T('Human ', 32, F.l, LIME), T('Resource', 32, F.b, LIME)]);
  tx(s, 4.344, 4.004, 1.585, 0.774, [
    T('Javier. J', 20, F.l, DARK),
    T(' ', 20, F.m, DARK),
    T('Robinson', 20, F.b, DARK),
  ]);
  tx(s, 4.344, 5.529, 2.322, 0.83, [
    T('PLACEHOLDER', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  tx(s, 9.983, 4.004, 1.71, 0.774, [
    T('Patrick', 20, F.l, DARK),
    T(' ', 20, F.m, DARK),
    T('Luis ', 20, F.l, DARK),
    T('Dembele', 20, F.b, DARK),
  ]);
  tx(s, 9.983, 5.529, 2.322, 0.83, [
    T('PLACEHOLDER', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  chrome(s, '26');
  rr(s, 11.41, 3.095, 1.005, 0.302, 50000, { fill: LIME });
  rr(s, 7.47, 2.989, 5.033, 0.512, 50000, { line: DARK, lw: 3 });
  tx(s, 7.609, 3.077, 2.193, 0.337, [T('Ai Technology', 14, F.m, DARK)]);
  oval(s, 12.134, 3.133, 0.224, DARK);
  tx(s, 7.47, 1.397, 5.115, 0.83, [
    T('Lorem Ipsum Dolor Sit Amet, abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle clutch speed', 10, F.r, DARK),
  ], { lineSpacing: 1.5 });
  imgBox(s, 1.736, 4.004, 2.317, 2.405);
  imgBox(s, 7.373, 4.004, 2.317, 2.405);
}

// -------- slide 27 - Keep in touch --------
function slide27(pptx) {
  const s = newSlide(pptx);
  rr(s, 0.54, 0.925, 5.515, 5.257, 16530, { fill: DARK });
  rr(s, 0.955, 1.317, 11.838, 5.433, 11002, { line: LIME, lw: 3 });
  bars(s, 5.681, 1.868);
  chrome(s, '27');
  tx(s, 1.519, 1.678, 2.273, 1.178, [T('Keep in ', 32, F.l, LIME), T('Touch', 32, F.b, LIME)]);
  tx(s, 1.594, 4.485, 1.42, 0.337, [T('Phone', 14, F.b, GREY)]);
  tx(s, 2.812, 4.485, 1.816, 0.337, [T('+123 456 7891', 14, F.r, GREY)]);
  tx(s, 1.594, 4.91, 1.42, 0.337, [T('Address', 14, F.b, GREY)]);
  tx(s, 2.812, 4.91, 2.414, 0.337, [
    T('21', 14, F.r, GREY),
    T('st', 14, F.r, GREY, { superscript: true }),
    T(' Street, New York 112', 14, F.r, GREY),
  ]);
  tx(s, 1.594, 5.335, 1.889, 0.337, [T('Email', 14, F.b, GREY)]);
  tx(s, 2.812, 5.335, 2.322, 0.337, [T('yourmail@mai.com', 14, F.r, GREY)]);
  tx(s, 1.519, 3.324, 3.898, 0.83, [
    T('PLACEHOLDER', 10, F.r, GREY),
  ], { lineSpacing: 1.5 });
  rr(s, 10.665, 5.689, 1.005, 0.302, 50000, { fill: LIME });
  rr(s, 7.209, 5.584, 4.577, 0.512, 50000, { line: DARK, lw: 3 });
  tx(s, 7.407, 5.672, 2.193, 0.337, [T('Our Contact', 14, F.m, DARK)]);
  oval(s, 11.389, 5.728, 0.224, DARK);
  imgBox(s, 7.239, 1.747, 4.463, 3.499);
}

// -------- slide 28 - Thank you --------
function slide28(pptx) {
  const s = newSlide(pptx);
  chrome(s, '28');
  rr(s, 0.54, 0.927, 11.632, 5.249, 15886, { fill: DARK });
  tx(s, 1.928, 1.967, 5.217, 2.794, [T('Thank You', 80, F.b, LIME)]);
  tx(s, 7.227, 4.734, 4.179, 0.83, [
    T('Ipsum Dolor, abstract feel guest trust at journey work tour glad with another monthly chair stair class table board ink erase caps enter the grip throttle clutch speed brake Lorem Ipsum', 10, F.r, GREY),
  ], { lineSpacing: 1.5 });
  bars(s, 11.74, 2.292);
  rr(s, 1.161, 1.571, 11.715, 5.207, 11002, { line: LIME, lw: 3 });
  rr(s, 10.229, 4.051, 1.005, 0.302, 50000, { fill: LIME });
  rr(s, 7.226, 3.945, 4.116, 0.512, 50000, { line: GREY, lw: 3 });
  tx(s, 7.377, 4.033, 2.193, 0.337, [T('Ai Technology', 14, F.m, GREY)]);
  oval(s, 10.953, 4.09, 0.224, DARK);
}


// --------------------------------------------------------------------------

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'W16x9', width: 13.333333333333334, height: 7.5 }); // 12192000 x 6858000 EMU
  pptx.layout = 'W16x9';
  pptx.title = 'Ai - Tech Presentation Template';
  const slides = [
    slide01, slide02, slide03, slide04, slide05, slide06, slide07,
    slide08, slide09, slide10, slide11, slide12, slide13, slide14,
    slide15, slide16, slide17, slide18, slide19, slide20, slide21,
    slide22, slide23, slide24, slide25, slide26, slide27, slide28,
  ];
  slides.forEach(function (fn) { fn(pptx); });
  return pptx.writeFile({ fileName: path.join(__dirname, OUT_NAME) });
}

build().then(function (f) { console.log('wrote', f); }).catch(function (e) { console.error(e); process.exit(1); });
