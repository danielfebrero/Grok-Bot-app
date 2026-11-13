/**
 * "Fashion Shop" deck - rebuilt with pptxgenjs.
 * Run: node 0c730ee8-68d5-4023-a1c9-b46f56814dd9_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Design tokens
 * ------------------------------------------------------------------ */
const SLIDE_W = 26.665;
const SLIDE_H = 15;

const PURPLE = '9E55A0';
const BLACK = '000000';
const WHITE = 'FFFFFF';
const GREY = 'BFBFBF';        // chart "secondary" series
const GREY_MID = 'A5A5A5';    // pie slice
const CHARCOAL = '404040';    // line-chart series 1

const HEAD = 'Poppins ExtraBold';
const BODY = 'Roboto Regular';
const FA_SOLID = 'Font Awesome 5 Free Solid';
const FA_BRAND = 'Font Awesome 5 Brands Regular';

/* Text presets: [fontFace, fontSize, color, lineSpacing(pt)] */
const T_TITLE = { fontFace: HEAD, fontSize: 90, color: PURPLE, lineSpacing: 100 };
const T_TITLE_W = { fontFace: HEAD, fontSize: 90, color: WHITE, lineSpacing: 100 };
const T_KICK = { fontFace: HEAD, fontSize: 50, color: PURPLE, lineSpacing: 60 };
const T_KICK_W = { fontFace: HEAD, fontSize: 50, color: WHITE, lineSpacing: 60 };
const T_BODY = { fontFace: BODY, fontSize: 36, color: BLACK, lineSpacing: 53 };
const T_BODY_W = { fontFace: BODY, fontSize: 36, color: WHITE, lineSpacing: 53 };
const T_NUM = { fontFace: HEAD, fontSize: 36, color: PURPLE, lineSpacing: 53 };
const T_MONTH = { fontFace: HEAD, fontSize: 32, color: WHITE };
const T_STEP = { fontFace: HEAD, fontSize: 66, color: WHITE };

const LOGO_BOX = [2.332, 1.225, 3.217, 1.457];
const NUM_BOX = [23.87, 13.012, 1.275, 0.799];

/* ------------------------------------------------------------------ *
 * Small helpers
 * ------------------------------------------------------------------ */
const up = (t) => (Array.isArray(t) ? t.map((s) => s.toUpperCase()) : t.toUpperCase());

/** Paragraph-per-array-entry text box; margins are always zero like the source deck. */
function text(s, body, x, y, w, h, opt = {}) {
  s.addText(Array.isArray(body) ? body.join('\n') : body, {
    x, y, w, h, margin: 0, valign: 'top', align: 'left', bullet: false, ...opt,
  });
}

const title = (s, t, x, y, w, h, o) => text(s, up(t), x, y, w, h, { ...T_TITLE, ...o });
const titleW = (s, t, x, y, w, h, o) => text(s, up(t), x, y, w, h, { ...T_TITLE_W, ...o });
const body = (s, t, x, y, w, h, o) => text(s, t, x, y, w, h, { ...T_BODY, ...o });
const bodyW = (s, t, x, y, w, h, o) => text(s, t, x, y, w, h, { ...T_BODY_W, ...o });
const kicker = (s, t, x, y, w, h, o) => text(s, up(t), x, y, w, h, { ...T_KICK, ...o });

/** Right-aligned section label + the 5pt rule underneath it (top-right corner motif). */
function tab(s, label, x, w) {
  kicker(s, label, x, 1.997, w, 0.841, { align: 'right' });
  s.addShape('line', { x, y: 3.06, w, h: 0, line: { color: PURPLE, width: 5 } });
}

function box(s, x, y, w, h, fill) {
  s.addShape('rect', { x, y, w, h, fill: { color: fill } });
}

function circle(s, x, y, d, fill) {
  s.addShape('ellipse', { x, y, w: d, h: d, fill: { color: fill } });
}

/**
 * Time/step markers.  These are body placeholders in the reference deck whose
 * geometry lives on the slide layout, so they paint as plain squares.
 */
function marker(s, x, y, d) {
  s.addShape('rect', { x, y, w: d, h: d, fill: { color: PURPLE } });
}

function outline(s, shape, x, y, w, h, width = 3, color = PURPLE) {
  s.addShape(shape, { x, y, w, h, fill: { type: 'none' }, line: { color, width } });
}

/** straight segment; `upward` draws bottom-left -> top-right */
function seg(s, x1, y1, x2, y2, width = 3, color = PURPLE) {
  s.addShape('line', {
    x: Math.min(x1, x2), y: Math.min(y1, y2),
    w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
    flipV: (x2 - x1) * (y2 - y1) < 0,
    line: { color, width },
  });
}

/* ------------------------------------------------------------------ *
 * Curved silhouettes (the deck's recurring half-disc decorations)
 * ------------------------------------------------------------------ */
const K = 0.5523; // circle -> cubic bezier constant

/** flat top edge, half-disc bulging downwards */
const DOME = (w, h) => [
  { x: 0, y: 0 },
  { x: w * 0.5, y: h, curve: { type: 'cubic', x1: w * 0.045, y1: h * 0.57, x2: w * 0.252, y2: h } },
  { x: w, y: 0, curve: { type: 'cubic', x1: w * 0.748, y1: h, x2: w * 0.955, y2: h * 0.57 } },
  { close: true },
];
/** flat bottom edge, half-disc bulging upwards */
const BOWL = (w, h) => [
  { x: w, y: h },
  { x: w * 0.5, y: 0, curve: { type: 'cubic', x1: w * 0.955, y1: h * 0.43, x2: w * 0.748, y2: 0 } },
  { x: 0, y: h, curve: { type: 'cubic', x1: w * 0.252, y1: 0, x2: w * 0.045, y2: h * 0.43 } },
  { close: true },
];
/** flat right edge, half-disc bulging to the left */
const LOBE_L = (w, h) => [
  { x: w, y: 0 },
  { x: 0, y: h * 0.5, curve: { type: 'cubic', x1: w * 0.43, y1: h * 0.045, x2: 0, y2: h * 0.252 } },
  { x: w, y: h, curve: { type: 'cubic', x1: 0, y1: h * 0.748, x2: w * 0.43, y2: h * 0.955 } },
  { close: true },
];

function curve(s, pts, x, y, w, h, fill) {
  s.addShape('custGeom', { x, y, w, h, fill: { color: fill }, points: pts(w, h) });
}

/* Quarter-discs used by the logo mark (square corner + one arc). */
const Q_BOTTOM_RIGHT = (w, h) => [
  { x: 0, y: h },
  { x: w, y: 0, curve: { type: 'cubic', x1: 0, y1: h * (1 - K), x2: w * (1 - K), y2: 0 } },
  { x: w, y: h }, { close: true },
];
const Q_TOP_RIGHT = (w, h) => [
  { x: 0, y: 0 },
  { x: w, y: h, curve: { type: 'cubic', x1: 0, y1: h * K, x2: w * (1 - K), y2: h } },
  { x: w, y: 0 }, { close: true },
];
const Q_TOP_LEFT = (w, h) => [
  { x: w, y: 0 },
  { x: 0, y: h, curve: { type: 'cubic', x1: w, y1: h * K, x2: w * K, y2: h } },
  { x: 0, y: 0 }, { close: true },
];

/* ------------------------------------------------------------------ *
 * Recurring furniture: logo lock-up + page number + image stand-ins
 * ------------------------------------------------------------------ */
function logo(s) {
  const [x, y, w, h] = LOGO_BOX;
  const gw = w * (36 / 232);       // glyph cell width
  const gh = h * (49 / 108);       // glyph cell height
  const gy = y + h * (1 / 108);
  const marks = [Q_BOTTOM_RIGHT, Q_TOP_RIGHT, Q_TOP_LEFT];
  [0, 41, 81].forEach((px, i) => {
    s.addShape('custGeom', {
      x: x + w * (px / 232), y: gy, w: gw, h: gh,
      fill: { type: 'none' }, line: { color: PURPLE, width: 3 },
      points: marks[i](gw, gh),
    });
  });
  text(s, 'FASHION SHOP', x + 0.02, y + h * (62 / 108), w, 0.42,
    { fontFace: HEAD, fontSize: 30, color: BLACK, charSpacing: 0.3 });
}

function pageNum(s, n) {
  text(s, String(n), NUM_BOX[0], NUM_BOX[1], NUM_BOX[2], NUM_BOX[3],
    { ...T_NUM, align: 'right', valign: 'middle', margin: [7.2, 7.2, 3.6, 3.6] });
}

/**
 * Photo slots.  Every picture frame in the reference deck is an *empty* picture
 * placeholder - the template ships without artwork - so the published deck shows
 * bare white in these regions.  The calls below keep the slots visible in the
 * source (position + size + silhouette) without painting anything on the slide.
 */
function imageFrame(_s, _x, _y, _w, _h, _shape = 'rect') { /* empty picture slot */ }
function imageCurve(_s, _pts, _x, _y, _w, _h) { /* empty picture slot */ }

/** Map pin: teardrop outline + inner ring. */
function pin(s, x, y, fill, stroke) {
  const w = 0.868, h = 1.163;
  s.addShape('custGeom', {
    x, y, w, h, fill: { color: fill }, line: { color: stroke, width: 4.5 },
    points: [
      { x: w * 0.5, y: h },
      { x: 0, y: h * 0.44, curve: { type: 'cubic', x1: w * 0.14, y1: h * 0.74, x2: 0, y2: h * 0.62 } },
      { x: w * 0.5, y: 0, curve: { type: 'cubic', x1: 0, y1: h * 0.2, x2: w * 0.22, y2: 0 } },
      { x: w, y: h * 0.44, curve: { type: 'cubic', x1: w * 0.78, y1: 0, x2: w, y2: h * 0.2 } },
      { x: w * 0.5, y: h, curve: { type: 'cubic', x1: w, y1: h * 0.62, x2: w * 0.86, y2: h * 0.74 } },
      { close: true },
    ],
  });
  outline(s, 'ellipse', x + w * 0.31, y + h * 0.2, w * 0.38, w * 0.38, 3, stroke === WHITE ? WHITE : PURPLE);
}

/** Purple disc with a white tick - the "list with image" bullets. */
function checkDisc(s, x, y, d) {
  circle(s, x, y, d, PURPLE);
  text(s, '\u2713', x, y, d, d,
    { fontFace: BODY, fontSize: 44, color: WHITE, align: 'center', valign: 'middle', bold: true });
}

/* ------------------------------------------------------------------ *
 * Line-art icon sets (slides 17 & 39) - drawn from primitives
 * ------------------------------------------------------------------ */
function iconGrowth(s) {                    // bar chart + rising line
  [[5.972, 9.441, 0.589, 1.205], [6.717, 9.797, 0.590, 0.849], [7.464, 9.066, 0.590, 1.580]]
    .forEach((b) => outline(s, 'rect', ...b));
  seg(s, 5.793, 9.36, 6.60, 8.62);
  seg(s, 6.60, 8.62, 7.15, 9.02);
  seg(s, 7.15, 9.02, 8.20, 8.24);
  seg(s, 8.269, 8.21, 7.75, 8.30);
  seg(s, 8.269, 8.21, 8.18, 8.72);
}
function iconFinance(s) {                   // bars + swooping arrow
  [[12.252, 9.899, 0.597, 0.780], [12.924, 9.698, 0.597, 0.984], [13.595, 8.878, 0.597, 1.800]]
    .forEach((b) => outline(s, 'rect', ...b));
  seg(s, 12.201, 9.69, 13.30, 9.10);
  seg(s, 13.30, 9.10, 13.98, 8.35);
  seg(s, 13.48, 8.30, 14.05, 8.24);
  seg(s, 14.05, 8.24, 14.13, 8.82);
}
function iconDocument(s) {                  // document + magnifier
  outline(s, 'rect', 18.887, 8.210, 1.887, 2.458);
  [9.547, 9.781, 10.012].forEach((y) => box(s, 19.089, y, 1.443, 0.075, PURPLE));
  outline(s, 'ellipse', 18.335, 8.210, 1.194, 1.196);
  seg(s, 18.60, 9.15, 17.97, 9.76, 4);
}
function iconStore(s) {                     // awning + shop front
  outline(s, 'rect', 3.156, 6.347, 3.104, 1.278);
  outline(s, 'rect', 3.484, 7.726, 2.409, 1.272);
  outline(s, 'rect', 4.717, 7.743, 0.941, 0.939);
  [4.19, 4.87, 5.55].forEach((x) => seg(s, x, 6.347, x, 7.625, 2));
}
function iconPriceTag(s) {                  // rotated square + hole
  s.addShape('rect', {
    x: 9.75, y: 6.85, w: 1.65, h: 1.65, rotate: 45,
    fill: { type: 'none' }, line: { color: PURPLE, width: 3 },
  });
  outline(s, 'ellipse', 9.90, 6.72, 0.52, 0.52);
  seg(s, 10.30, 7.95, 11.10, 7.15, 2);
  seg(s, 10.62, 8.28, 11.42, 7.48, 2);
}
function iconGift(s) {                      // lid, box, ribbon, bow
  outline(s, 'rect', 14.816, 6.941, 2.837, 0.802);
  outline(s, 'rect', 15.043, 7.743, 2.382, 1.347);
  seg(s, 16.235, 6.941, 16.235, 9.090, 3);
  outline(s, 'ellipse', 15.44, 6.26, 0.80, 0.70);
  outline(s, 'ellipse', 16.23, 6.26, 0.80, 0.70);
}
function iconTruck(s) {                     // van body + cab + wheels
  outline(s, 'rect', 20.613, 6.363, 1.851, 1.568);
  outline(s, 'rect', 22.464, 6.95, 0.510, 0.981);
  seg(s, 20.613, 7.931, 22.974, 7.931, 3);
  outline(s, 'ellipse', 21.000, 8.366, 0.615, 0.615);
  outline(s, 'ellipse', 22.422, 8.382, 0.611, 0.615);
}

/* ------------------------------------------------------------------ *
 * Chart data
 * ------------------------------------------------------------------ */
const YEARS4 = ['2016', '2017', '2018', '2019'];
const MONTHS7 = ['January', 'February', 'March', 'April', 'May', 'June', 'July'];
// The source worksheet reserves an 8th (blank) row, which shifts the bars down.
const MONTHS7_PAD = MONTHS7.concat(['']);

function progressChart(s, pptx) {
  const x = 3.096, y = 5.257, w = 20.474, h = 2.808;
  s.addChart(pptx.ChartType.bar, [
    { name: '88% growth', labels: [''], values: [88] },
    { name: '12% growth', labels: [''], values: [12] },
  ], {
    x, y, w, h,
    barDir: 'bar', barGrouping: 'percentStacked', barGapWidthPct: 10, barOverlapPct: 100,
    chartColors: [PURPLE, GREY], catAxisHidden: true, valAxisHidden: true,
    showLegend: false, showValue: false,
    valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
  });
  text(s, '88% growth', x, y, w * 0.88, h,
    { fontFace: HEAD, fontSize: 66, bold: true, color: WHITE, align: 'center', valign: 'middle' });
}

function columnChart(s, pptx) {
  s.addChart(pptx.ChartType.bar, [
    { name: 'Income', labels: YEARS4, values: [50, 60, 30, 40] },
    { name: 'Costs', labels: YEARS4, values: [50, 65, 80, 90] },
  ], {
    x: 2.421, y: 3.408, w: 11.699, h: 8.51,
    barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 50, barOverlapPct: 100,
    chartColors: [PURPLE, GREY],
    showValue: true, dataLabelPosition: 'ctr', dataLabelFormatCode: '[$$-409]#,##0',
    dataLabelFontFace: HEAD, dataLabelFontSize: 36, dataLabelColor: WHITE,
    showLegend: true, legendPos: 'r', legendFontSize: 28,
    catAxisLabelFontSize: 20, catAxisLabelColor: BLACK, catAxisLineColor: BLACK, catAxisMajorTickMark: 'none',
    valAxisLabelFontSize: 20, valAxisLabelColor: BLACK, valAxisLineColor: BLACK, valAxisMajorTickMark: 'none',
    valGridLine: { color: BLACK, size: 0.5 }, catGridLine: { style: 'none' },
  });
}

function barChart(s, pptx) {
  s.addChart(pptx.ChartType.bar, [
    { name: 'Value 2', labels: MONTHS7_PAD, values: [1200, 1400, 1800, 1350, 1220, 1800, 1050, null] },
  ], {
    x: 2.332, y: 3.11, w: 11.448, h: 8.328,
    barDir: 'bar', barGrouping: 'clustered', barGapWidthPct: 8, barOverlapPct: 8,
    chartColors: [GREY, PURPLE, GREY, PURPLE, GREY, PURPLE, GREY],
    showValue: true, dataLabelPosition: 'ctr', dataLabelFormatCode: 'General',
    dataLabelFontFace: HEAD, dataLabelFontSize: 36, dataLabelColor: WHITE,
    showLegend: false,
    catAxisLabelFontSize: 20, catAxisLabelColor: BLACK, catAxisLineColor: BLACK, catAxisMajorTickMark: 'none',
    valAxisLabelFontSize: 20, valAxisLabelColor: BLACK, valAxisLineColor: BLACK, valAxisMajorTickMark: 'none',
    valAxisMaxVal: 2000, valAxisMajorUnit: 500,
    valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
  });
}

function lineChart(s, pptx) {
  s.addChart(pptx.ChartType.line, [
    { name: ' Income', labels: YEARS4, values: [4.3, 2.5, 3.5, 4.5] },
    { name: ' Costs', labels: YEARS4, values: [2.4, 4.4, 0.5, 2.8] },
  ], {
    x: 2.507, y: 3.95, w: 10.826, h: 7.875,
    chartColors: [CHARCOAL, PURPLE], lineSize: 8, lineDataSymbol: 'none',
    showLegend: true, legendPos: 'b', legendFontSize: 28, showValue: false,
    catAxisLabelFontSize: 20, catAxisLabelColor: BLACK, catAxisLineColor: BLACK, catAxisMajorTickMark: 'none',
    valAxisLabelFontSize: 20, valAxisLabelColor: BLACK, valAxisLineColor: BLACK,
    valAxisMajorUnit: 1, valGridLine: { color: BLACK, size: 0.5 }, catGridLine: { style: 'none' },
  });
}

function areaChart(s, pptx) {
  const years5 = ['2016', '2017', '2018', '2019', '2020'];
  s.addChart(pptx.ChartType.area, [
    { name: 'Income', labels: years5, values: [14, 18, 15, 20, 16] },
    { name: 'Costs', labels: years5, values: [10, 14, 8, 12, 10] },
  ], {
    x: 2.701, y: 3.132, w: 11.418, h: 8.306,
    chartColors: [GREY, PURPLE], showLegend: true, legendPos: 'r', legendFontSize: 28,
    catAxisLabelFontSize: 20, catAxisLabelColor: BLACK, catAxisLineColor: BLACK,
    valAxisLabelFontSize: 20, valAxisLabelColor: BLACK, valAxisLineColor: BLACK,
    valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
  });
}

function doughnutChart(s, pptx) {
  s.addChart(pptx.ChartType.doughnut, [
    { name: 'Sales', labels: ['Goods', 'Services', 'Other'], values: [0.35, 0.4, 0.25] },
  ], {
    x: 6.246, y: 1.998, w: 14.175, h: 11.813,
    holeSize: 50, chartColors: [BLACK, GREY_MID, PURPLE],
    showValue: true, dataLabelFormatCode: '0%',
    dataLabelFontFace: HEAD, dataLabelFontSize: 80, dataLabelColor: WHITE,
    showLegend: true, legendPos: 'r', legendFontSize: 36,
    layout: { x: 0.02, y: 0.06, w: 0.72, h: 0.87 },
  });
}

/** Thin ring gauge + the percentage sitting in the hole. */
function ringGauge(s, pptx, x, y, value) {
  const w = 7.087, h = 7.072;
  s.addChart(pptx.ChartType.doughnut, [
    { name: 'Ratio', labels: ['done', 'rest'], values: [value, 1 - value] },
  ], {
    x, y, w, h, holeSize: 75, chartColors: [PURPLE, GREY],
    showValue: false, showLegend: false,
    layout: { x: 0.131, y: 0.131, w: 0.739, h: 0.741 },
  });
  text(s, Math.round(value * 100) + '%', x, y + h / 2 - 0.8, w, 1.6,
    { fontFace: HEAD, fontSize: 80, color: PURPLE, align: 'center', valign: 'middle' });
}

/* ------------------------------------------------------------------ *
 * Tables
 * ------------------------------------------------------------------ */
const CELL_PAD = [5.7, 28.3, 5.7, 28.3];    // pts: T R B L (source: 0.079" / 0.394")

function verticalTable(s) {
  const head = ['Name', 'Property', 'Price'];
  const rows = [
    ['Lorem ipsum', 'Integer in ex varius ', '1000.00'],
    ['Dolor', 'Consectetur lorem', '1500.00'],
    ['Consectetur', 'Duis sit amet porta nisi', '3000.00'],
    ['Adipiscing', 'Mauris non ultrices velit', '4500.00'],
    ['Praesent', 'Aliquam mattis aliquam', '4750.00'],
    ['Tristique', 'Praesent vitae tincidunt', '5000.00'],
  ];
  const none = { type: 'none' };
  const rule = { type: 'solid', color: PURPLE, pt: 1 };
  const whiteRule = { type: 'solid', color: WHITE, pt: 1 };
  const headRow = head.map((t, i) => ({
    text: up(t),
    options: {
      fill: { color: PURPLE }, color: WHITE, fontFace: HEAD, fontSize: 48,
      align: i === 2 ? 'right' : 'left', valign: 'middle', margin: CELL_PAD,
      border: [none, i === 2 ? none : whiteRule, none, i === 0 ? none : whiteRule],
    },
  }));
  const bodyRows = rows.map((r) => r.map((t, i) => ({
    text: t,
    options: {
      color: BLACK, fontFace: BODY, fontSize: 36,
      align: i === 2 ? 'right' : 'left', valign: 'middle', margin: CELL_PAD,
      border: [rule, i === 2 ? none : rule, rule, i === 0 ? none : rule],
    },
  })));
  s.addTable([headRow, ...bodyRows], {
    x: 4.069, y: 3.562, w: 18.901, colW: [4.728, 9.44, 4.733],
    rowH: [1.183, 1.177, 1.177, 1.177, 1.177, 1.177, 1.177],
  });
}

function horizontalTable(s) {
  const rows = [
    ['March', 'Integer in ex varius, congue ligula', '1000.00'],
    ['April', 'Duis sit amet porta nisi mauris non', '1500.00'],
    ['May', 'Aliquam mattis aliquam nibh praesent', '3000.00'],
    ['June', 'Donec eros nisi, sagittis non nulla', '4500.00'],
    ['July', 'Praesent placerat elit id erat rutrum ', '4750.00'],
    ['August', 'Vestibulum convallis vehicula', '5000.00'],
  ];
  const none = { type: 'none' };
  const rule = { type: 'solid', color: PURPLE, pt: 1 };
  const whiteRule = { type: 'solid', color: WHITE, pt: 1 };
  const data = rows.map(([m, d, p], r) => {
    const first = r === 0, last = r === rows.length - 1;
    return [
      {
        text: up(m),
        options: {
          fill: { color: PURPLE }, color: WHITE, fontFace: HEAD, fontSize: 48,
          align: 'left', valign: 'middle', margin: CELL_PAD,
          border: [first ? none : whiteRule, rule, last ? none : whiteRule, none],
        },
      },
      {
        text: d,
        options: {
          color: BLACK, fontFace: BODY, fontSize: 36, align: 'left', valign: 'middle', margin: CELL_PAD,
          border: [first ? none : rule, rule, last ? none : rule, rule],
        },
      },
      {
        text: p,
        options: {
          color: BLACK, fontFace: BODY, fontSize: 36, align: 'right', valign: 'middle', margin: CELL_PAD,
          border: [first ? none : rule, none, last ? none : rule, rule],
        },
      },
    ];
  });
  s.addTable(data, {
    x: 3.882, y: 4.351, w: 18.901, colW: [4.728, 10.64, 3.533],
    rowH: [1.181, 1.181, 1.181, 1.181, 1.181, 1.181],
  });
}

/* ------------------------------------------------------------------ *
 * Repeated slide skeletons
 * ------------------------------------------------------------------ */
/** Body copy flowed into two columns, as the source layout's numCol="2" does. */
function twoColBody(s, lines, x, y, w, h, gutter = 0.984) {
  const colW = (w - gutter) / 2;
  lines.forEach((para, i) => body(s, para, x + i * (colW + gutter), y, colW, h));
}

/** Big centred heading + one full-width paragraph. */
function centredHeading(s, head, headBox, para, paraBox) {
  title(s, head, ...headBox, { align: 'center' });
  body(s, para, ...paraBox);
}

/** 3 stacked "step" markers used by the vertical timelines. */
function timelineSteps(s, steps) {
  steps.forEach(([n, y]) => {
    marker(s, 13.36, y, 3.171);
    text(s, n, 13.36, y, 3.171, 1.7,
      { ...T_STEP, align: 'center', valign: 'top', margin: [0, 0, 0, 11.3] });
  });
}

/** Horizontal timeline markers (slides 45 & 46). */
function timelineRow(s, items) {
  items.forEach(([n, label, x]) => {
    marker(s, x, 9.889, 3.171);
    text(s, n, x, 9.889, 3.171, 1.7,
      { ...T_STEP, align: 'center', valign: 'top', margin: [0, 0, 0, 11.3] });
    text(s, label, x, 11.475, 3.171, 0.586, { ...T_MONTH, align: 'center' });
  });
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */
const SLIDES = [

  // 1 - Title -----------------------------------------------------------
  (s) => {
    logo(s);
    imageFrame(s, 12.571, 1.221, 11.791, 11.791, 'ellipse');
    title(s, 'More Fun, More Fashion', 2.324, 4.359, 9.45, 5.238,
      { fontSize: 116, color: BLACK, lineSpacing: 124 });
    body(s, ['yourwebsite.com', 'your@name.com'], 3.13, 11.459, 4.707, 1.575);
    body(s, ['yourcompany', '@yourcompany'], 9.413, 11.459, 4.707, 1.416);
    [['\uf245', FA_SOLID, 2.517, 11.624, 0.28], ['\uf39e', FA_BRAND, 8.818, 11.642, 0.28],
     ['\uf0e0', FA_SOLID, 2.434, 12.419, 0.449], ['\uf16d', FA_BRAND, 8.765, 12.429, 0.393]]
      .forEach(([g, f, x, y, w]) => text(s, g, x, y, w, 0.485,
        { fontFace: f, fontSize: 32, color: PURPLE, align: 'center' }));
  },

  // 2 - Table of contents ------------------------------------------------
  (s) => {
    title(s, 'Table of contents', 4.67, 4.372, 17.325, 1.402, { align: 'center' });
    const entries = [['Introduction', '02'], ['Stories', '04'], ['Infographics', '28'],
      ['Summary', '44'], ['Contacts', '48']];
    body(s, entries.map((e) => e[0]), 6.912, 5.959, 5.634, 5.259, { fontSize: 45, lineSpacing: 75 });
    body(s, entries.map((e) => e[1]), 18.057, 5.959, 1.71, 5.259,
      { fontSize: 45, lineSpacing: 75, align: 'right' });
    [[10.665, 6.712, 8.050], [9.236, 7.737, 9.479], [10.792, 8.787, 7.922],
     [10.005, 9.850, 8.710], [9.880, 10.875, 8.835]]
      .forEach(([x, y, w]) => s.addShape('line',
        { x, y, w, h: 0, line: { color: BLACK, width: 1, dashType: 'dash' } }));
  },

  // 3 - Agenda -----------------------------------------------------------
  (s) => {
    title(s, 'Agenda', 2.307, 2.811, 22.041, 1.402, { align: 'center' });
    const cols = [
      ['11:00', 'AM', 3.730, 5.167, 'Donec quis tortor tereterot rweleifend, viver raweraski sollicitudin eget ante.', 2.328, 9.095, 6.280],
      ['2:00', 'pM', 11.556, 5.137, 'Phasellus a imperdi poqer eterlectus, sit ametiwetupi sollicitudin pretiumerom.', 10.202, 9.113, 6.304],
      ['4:00', 'pM', 19.382, 5.158, 'Pellen tesque dictumoperi reterat sit amet sollicitusa din eget antecurs.', 18.076, 9.095, 6.247],
    ];
    cols.forEach(([t1, t2, cx, cy, para, px, py, pw]) => {
      marker(s, cx, cy, 3.543);
      text(s, up([t1, t2]), cx, cy, 3.543, 3.543, {
        fontFace: HEAD, fontSize: 65, color: WHITE, align: 'center', valign: 'top',
        lineSpacingMultiple: 0.9, margin: [0, 0, 0, 25.5],
      });
      body(s, para, px, py, pw, 2.27);
    });
  },

  // 4 - Heading ----------------------------------------------------------
  (s) => {
    curve(s, DOME, 8.661, 0, 9.396, 3.887, PURPLE);
    centredHeading(s, 'Fashion is new Treat', [2.307, 5.946, 22.081, 1.402],
      'Etiam lacinia fermentum est non lobortis. Maecenas non facilisis libero. Curabit urmesuada wertopit interdum scelerisque. Vestibulum sollicitudin egetante curs us pret. Curabitu werport reget sadop imperdiet tellus, venenat is sodale sertopew justo. Phasellus a imperdit eterlectus, sit amet quisasd tortor fermentum est teviverra est non lobortis lacinia fermentum.',
      [2.332, 7.541, 22.026, 2.973]);
  },

  // 5 - Heading with background -----------------------------------------
  (s) => {
    imageFrame(s, 13.215, 0, 13.446, 15);
    title(s, ['Everyday', 'living'], 2.332, 5.179, 9.425, 2.808);
    body(s, 'Vestibulum sollicitudin eget antecurs user pretium. Curabi tupeget imperdi etotelluso, venrewopenat is sodale sjusto. In hacwero rewtop dopi wero habitasse.',
      2.303, 8.338, 9.425, 2.973);
  },

  // 6 - Heading + one column --------------------------------------------
  (s) => {
    curve(s, BOWL, 2.328, 11.087, 9.453, 3.91, PURPLE);
    title(s, ['Clothes', 'are my Life'], 3.12, 5.936, 9.425, 2.808);
    body(s, ['Lorem ipsum dolorsitamet, werti rewtop consece tur adipiscing elit. Proino werop aliquet egeterwe mauri mollis eneanewr viverra quam dolor.',
      '',
      'Phasellus ullamcorper augue eter wero metupo auctor maximures. Donec dolor poritop imperdiet, exut tinocidunt ullam congue, eliter diam enean.'],
      14.144, 4.379, 9.425, 6.69);
  },

  // 7 - Heading + one column on background ------------------------------
  (s) => {
    imageFrame(s, 0, 0, 26.665, 15);
    circle(s, 7.819, 1.998, 11.014, PURPLE);
    titleW(s, ['True style', 'never dies'], 9.424, 4.362, 7.874, 2.805);
    bodyW(s, 'Quisque sed nibh mi. Pellentesque in metus ut quam aliquet ullamcor peret non tortor. Vivamus at elit ac diam venenatis dapibus.',
      9.42, 7.513, 7.85, 2.973);
  },

  // 8 - Heading + multi column ------------------------------------------
  (s) => {
    curve(s, DOME, 13.387, 0, 9.396, 3.887, PURPLE);
    title(s, ['Wide', 'Range,', 'Big fun'], 2.338, 5.933, 6.27, 4.21);
    twoColBody(s, [
      'Aenean congue, risus ut pretium accumsan, elit sem accumsan massa, a iaculis libero urna vel ipsum. Mauris scelerisque leo sapien, wer vehicula tristique tortorop euismod utenean blanditop. Enisi elit, ac pulvinar sapien',
      'suscipit vel. Aliquam aliquam erat in condimentum egestas. Etiam lacinia lectus ipsum, vel tincidunt odio ullamcorper eget. Proin volutpa venenatis magna tempus.',
    ], 10.208, 5.174, 14.15, 5.239);
  },

  // 9 - Heading + multi column on background ----------------------------
  (s) => {
    imageFrame(s, 2.332, 5.137, 6.3, 6.3, 'ellipse');
    title(s, ['Fashion as unique', 'as you are'], 10.207, 4.362, 14.147, 2.808);
    twoColBody(s, [
      'Sed nec dictum velit, non iaculis enim. Aenean ut convallis metus. Cras iaculis nibh vel tellus finibus, ac laoreet neque maximus.',
      'Integer vel nulla justo. Sed vitae aliquam augue. Lorem consectetur adipiscing elit. Fusce at nunc faucibus, susci pitnunca, porta elit.',
    ], 10.204, 7.514, 14.15, 3.716);
  },

  // 10 - Cover collage 1 --------------------------------------------------
  (s) => {
    imageFrame(s, 14.12, 2.381, 10.238, 10.238, 'ellipse');
    title(s, 'About us', 2.35, 5.954, 9.407, 1.402);
    body(s, 'Maecenas libero ex, hendrerit eu aliquet vel, consectetur sed massa. Phasellus et interdum tortor, ut facilisis lectus. Morbi condimentum laoreet finibus.',
      2.35, 7.557, 9.407, 2.973);
  },

  // 11 - Cover collage 2 --------------------------------------------------
  (s) => {
    imageFrame(s, 2.348, 2.381, 10.238, 10.238, 'ellipse');
    title(s, 'Why us?', 14.951, 5.933, 9.407, 1.442);
    body(s, 'Vestibulum sollicitudin eget antecurs user pretium. Curabi tupeget imperd etotelluso, venrewopena is sodale sjusto. In hacwero rewtop dopi wero habitasse.',
      14.951, 7.537, 9.407, 2.973);
  },

  // 12 - Right image ------------------------------------------------------
  (s) => {
    imageFrame(s, 13.608, 2.771, 13.054, 12.226);
    title(s, ['We promise', 'comfort'], 2.35, 4.362, 9.407, 2.805, { bold: true });
    body(s, 'Integer sit amet vehicula sapien. Maecena vestibulum interdum neque vel posuere. Donec accumsan, odio ac aliquam sceleri que, erat ligula condimentum augue, quis aliquet velit elit in purus.',
      2.35, 7.557, 9.407, 3.716);
  },

  // 13 - Left image -------------------------------------------------------
  (s) => {
    imageFrame(s, 0, 0, 13.054, 12.226);
    title(s, ['A planet', 'of Style'], 14.916, 4.355, 9.407, 2.808);
    body(s, ['Donec accumsan, odio ac aliquam sceleri que, erat ligula condimentum augue, quis aliquet velit elit in purus lacinia nec mauri.',
      'Integer sit amet vehicula sapien. Maecena vestibulum interdum neque.'],
      14.916, 7.526, 9.441, 3.716);
  },

  // 14 - Bottom image -----------------------------------------------------
  (s) => {
    imageCurve(s, DOME, 7.694, 9.325, 11.33, 5.672);
    centredHeading(s, 'Enjoy a New Hours', [2.332, 4.355, 22.026, 1.405],
      'Vestibulum sollicitudin eget antcurs us pretium. Curabitu reget imperdiet tellus, venenat is sodale justo. In hac rewtop habitasse platea habitas imperdiet tellus  dictumst.',
      [2.35, 5.93, 22.008, 1.416]);
    tab(s, 'Objectives', 20.054, 4.304);
  },

  // 15 - Top image --------------------------------------------------------
  (s) => {
    imageCurve(s, BOWL, 7.687, 0, 11.332, 5.672);
    centredHeading(s, 'Style that Matters', [2.307, 7.532, 22.045, 1.402],
      'Fusce laoreet, nisi id cursus egestas, purus lectus tempus mauris, id cursus neque tortor sed turpis. Pellentesque quis libero sit amet nulla faucibus rhoncus eget sit amet arcu. Proin nec porttitor sem. Aliquam mollis neque vitae facilisis dapibus.',
      [2.35, 9.111, 22.008, 2.159]);
  },

  // 16 - What we do (3 columns) ------------------------------------------
  (s) => {
    curve(s, DOME, 8.632, 0, 9.396, 3.887, PURPLE);
    title(s, 'What we do?', 2.307, 5.154, 22.094, 1.405, { align: 'center' });
    [['style', 'Donec quis tortor terer reti rweleifend, viver raweroqit nequ, volutpat justwohor vestibu lumsaterop.', 2.328, 6.280],
     ['Connection', 'Phasellus awim perdipowe eterlectus, sit ametwetopy rhonc wedui. Sed acretewt nequewe rewtyip.', 10.202, 6.280],
     ['Confidence', 'Pellentesque dictumot pot reterat sit amet eraterower aculs interdum. Sedopisad wert auguewerhre.', 18.076, 6.247]]
      .forEach(([head, para, x, w]) => {
        kicker(s, head, x, 7.254, w, 0.841);
        body(s, para, x, 8.325, w, 2.973);
      });
  },

  // 17 - Our goals --------------------------------------------------------
  (s) => {
    centredHeading(s, 'Our goals', [2.307, 4.372, 22.049, 1.405],
      'Nullam imperdiet suscipit magna nec porta. Maecenas sed malesuada leo, acsodales urnwa. Proin tempus diam in dictum consectetur. In enim dictumot reterat sitomi, varius velitve.',
      [2.328, 5.946, 22.008, 1.487]);
    iconGrowth(s); iconFinance(s); iconDocument(s);
    [['Development', 4.670, 4.726], ['Financial gain', 10.970, 4.722], ['Perfection', 17.309, 4.636]]
      .forEach(([t, x, w]) => body(s, t, x, 11.458, w, 0.743, { align: 'center' }));
  },

  // 18 - Our values -------------------------------------------------------
  (s) => {
    curve(s, BOWL, 8.622, 11.087, 9.453, 3.91, PURPLE);
    title(s, 'Our values', 2.307, 4.372, 22.049, 1.405, { align: 'center' });
    [['Phasellus lorem justo, vari usnec enim id, facilisis wert finibus felis. Donec conse quat dictum rhoncus. Null maximus felis nec.', 2.332, 5.950, 6.280],
     ['Cras ligula ante, rutrumoper malesuada arcu at, semp erelementum quis urnasad sapien. Morbi finibus cons ctetur amesagittis.', 10.207, 5.968, 6.304],
     ['Vestibulum tincidunt sapie dolor, sed faucibus risussad consequat nec. Vivamus vitae ex ante. Duis ac sem euismod, sodales.', 18.080, 5.950, 6.247]]
      .forEach(([t, x, y, w]) => body(s, t, x, y, w, 3.716));
  },

  // 19 - Our team ---------------------------------------------------------
  (s) => {
    title(s, 'OUR TEAM', 3.882, 2.812, 18.899, 1.405, { align: 'center' });
    [['Jane smith', 'Founder/Designer', 3.882, 3.905, 4.722, 4.726],
     ['Jane Doe', 'Design Assistant', 10.970, 10.992, 4.722, 4.722],
     ['John Doe', 'Shipping Manager', 18.057, 18.080, 4.726, 4.777]]
      .forEach(([name, role, tx, px, nw, rw], i) => {
        imageFrame(s, px, i === 2 ? 5.163 : 5.139, 4.724, 4.724, 'ellipse');
        kicker(s, name, tx, 10.58, nw, 0.841, { align: 'center' });
        body(s, role, tx, i === 0 ? 11.530 : 11.531, rw, 0.743, { align: 'center' });
      });
  },

  // 20 - Testimonials -----------------------------------------------------
  (s) => {
    imageFrame(s, 13.215, 0, 13.446, 15);
    title(s, ['Fashion', 'that Never', 'Sleeps'], 2.326, 5.161, 7.832, 4.21);
    kicker(s, 'Jane Smith', 2.307, 9.425, 7.875, 0.841, { align: 'right' });
  },

  // 21 - Our clients ------------------------------------------------------
  (s) => {
    title(s, 'Our clients', 3.095, 3.598, 20.499, 1.405, { align: 'center' });
    [5.960, 9.907].forEach((y, row) => {
      [0, 1, 2, 3].forEach((c) => {
        const x = (row === 0 ? 3.104 : 3.117) + c * 5.5095;
        imageFrame(s, x, y, 3.938, 2.363);
      });
    });
  },

  // 22 - Quotation --------------------------------------------------------
  (s) => {
    imageFrame(s, 0, 0, 26.665, 15);
    circle(s, 7.819, 1.998, 11.014, PURPLE);
    titleW(s, ['\u201Cpassion', 'meets     Perfection.\u201D'], 9.005, 4.372, 8.701, 5.61);
    text(s, up('Jane Doe'), 9.0, 8.659, 8.706, 0.841, { ...T_KICK_W, align: 'right' });
    bodyW(s, 'Design Assistant', 9.0, 9.479, 8.706, 0.743, { align: 'right' });
  },

  // 23 - Gallery, 2 images ------------------------------------------------
  (s) => {
    imageFrame(s, 0, 0, 13.054, 12.226);
    imageFrame(s, 13.608, 2.771, 13.054, 12.226);
  },

  // 24 - Gallery, 2 images (full bleed) -----------------------------------
  (s) => {
    imageFrame(s, 0.003, 0.003, 12.818, 15);
    imageFrame(s, 13.854, 0, 12.814, 15);
  },

  // 25 - Gallery, 3 images ------------------------------------------------
  (s) => {
    tab(s, 'Fashion that Talks', 16.805, 7.552);
    imageFrame(s, 1.974, 4.854, 6.3, 6.3, 'ellipse');
    imageFrame(s, 9.001, 3.589, 8.661, 8.661, 'ellipse');
    imageFrame(s, 18.389, 4.769, 6.3, 6.3, 'ellipse');
  },

  // 26 - Gallery, 4 images ------------------------------------------------
  (s) => {
    title(s, ['Clothes', 'that', 'Creates', 'passion'], 2.326, 5.161, 7.856, 5.613);
    [[12.23, 1.225], [18.845, 1.225], [12.23, 7.791], [18.845, 7.791]]
      .forEach(([x, y]) => imageFrame(s, x, y, 5.536, 5.536, 'ellipse'));
  },

  // 27 - Progress ---------------------------------------------------------
  (s, p) => {
    tab(s, 'Progress', 20.667, 3.69);
    progressChart(s, p);
    body(s, 'Fusce laoreet, nisi id cursus egestas, purus lectus tempus mauris, id cursus neque tortor sed turpis. Pellentesque quis libero sit amet nulla faucibus rhoncus eget sit amet arcu. Proin nec porttitor sem. Aliquam mollis neque vitae facilisis dapibus.',
      3.095, 9.093, 20.474, 2.23);
  },

  // 28 - Ratio ------------------------------------------------------------
  (s) => {
    tab(s, 'ratio', 22.186, 2.172);
    [['75%', 3.095, 9.434, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Curabitur rutrum ligula leo, et maximus ante consequat et. Mauris laoreet bibendum justo.', 3.123],
     ['94%', 14.12, 9.479, 'Aenean sed finibus elit, a vehicula purus. Morbi elit est, tristique imperdiet nisl scele risque, varius rhoncus leo. Quisque lacus velit, sagittis a purus vel.', 14.148]]
      .forEach(([pct, x, w, para, px]) => {
        text(s, pct, x, 4.372, w, 3.871,
          { fontFace: HEAD, fontSize: 230, color: PURPLE, lineSpacing: 276, align: 'center' });
        body(s, para, px, 8.318, 9.422, 2.973);
      });
  },

  // 29 - Numeric information ---------------------------------------------
  (s) => {
    curve(s, LOBE_L, 22.727, 2.734, 3.941, 9.531, PURPLE);
    title(s, '1.9 mln', 6.266, 4.378, 14.155, 1.405);
    kicker(s, ['Employment in fashion-related', 'industries'], 6.266, 5.941, 14.155, 1.683);
    body(s, 'Duis in orci ut ante bibendum malesuada vitae id nulla. Nuncon vallis lectus sit amet elit convallis gravida. Donec rhoncus dictu orci at tempor. Vestibulum mi ligula, condimentum sit amet volutpat posuere, placerat quis eros.',
      6.266, 8.002, 14.175, 2.973);
  },

  // 30 - Column chart -----------------------------------------------------
  (s, p) => {
    columnChart(s, p);
    title(s, 'Annual income', 14.944, 4.379, 9.407, 2.808);
    body(s, 'Lorem ipsum dolor sit amet, consewro ect etur adipiscing elit. Praesenot varo snisier loboportis fineytbus. Lorem ipsum dolor sit amet, cons ecte tuweri adipiscing. ',
      14.944, 7.512, 9.407, 2.973);
  },

  // 31 - Bar chart --------------------------------------------------------
  (s, p) => {
    barChart(s, p);
    title(s, ['Spending', 'comparison'], 14.928, 4.359, 9.443, 2.808);
    body(s, 'Aenean sed finibus elit, a vehicula purus. Morbi elit est, tristique imperdiet nisl scele risque, varius rhoncus leo. Quisque lacus velit, sagittis a purus vel, porttitor.',
      14.944, 7.512, 9.407, 2.973);
  },

  // 32 - Line chart -------------------------------------------------------
  (s, p) => {
    lineChart(s, p);
    title(s, ['Patient', 'increment'], 14.944, 4.379, 9.407, 2.808);
    body(s, 'Vestibulum mi ligula, condimentum sit amet volutpat posuere, placerat quis eros. Praesent rhoncus feugiat velit non tincidu. Phasellus consectetur scelerisque.',
      14.944, 7.512, 9.407, 2.973);
  },

  // 33 - Area chart -------------------------------------------------------
  (s, p) => {
    areaChart(s, p);
    title(s, 'Growth graph', 14.944, 4.379, 9.407, 2.808);
    body(s, 'Lorem ipsum dolor sit amet, consewro ect etur adipiscing elit. Praesenot varo snisier loboportis fineytbus. Lorem ipsum dolor sit amet, cons ecte tuweri adipiscing.',
      14.944, 7.512, 9.407, 2.973);
  },

  // 34 - Doughnut ---------------------------------------------------------
  (s, p) => {
    tab(s, 'Sales comparison', 17.209, 7.149);
    doughnutChart(s, p);
  },

  // 35 - Market share -----------------------------------------------------
  (s, p) => {
    tab(s, 'Market share', 18.962, 5.396);
    ringGauge(s, p, 1.897, 2.927, 0.25);
    ringGauge(s, p, 9.802, 2.927, 0.56);
    ringGauge(s, p, 17.659, 2.932, 0.80);
    [['Null maximus felis necmolls. Phasellus lorem justo, vari usnec enim id, facilisis.', 2.332, 9.875, 6.280],
     ['Morbi finibus cons ctetur nisi sit amesagittis.Cras ligu lante, rutrumoper.', 10.206, 9.872, 6.304],
     ['Vivamus vitae ex ante. Duis ac sem euismod, sodales stibulum tincidunap.', 18.079, 9.875, 6.247]]
      .forEach(([t, x, y, w]) => body(s, t, x, y, w, 2.23));
  },

  // 36 - List of items ----------------------------------------------------
  (s) => {
    curve(s, DOME, 8.661, 0, 9.396, 3.887, PURPLE);
    tab(s, 'List of items', 19.64, 4.717);
    [['01', 2.361, 4.351, 2.309, 'Vivamus cursus, liberow quis condimenpo tuwertop dictum, ante mauris reno', 2.332, 5.935, 1.487],
     ['02', 2.328, 9.113, 2.342, 'Lorem ipsum dolor sitsa amet, consectewi tuwerop adipiscing elit mokolas', 2.341, 10.676, 1.416],
     ['03', 14.929, 4.384, 2.380, 'Etiam et semper arcupo. Quisque dictum quis elite in consecte turwer', 14.942, 5.947, 1.416],
     ['04', 14.929, 9.104, 2.380, 'Aliquam neque maspos, eleifend accado viverra werorewqurtis, iaculis id', 14.900, 10.689, 1.487]]
      .forEach(([n, nx, ny, nw, para, px, py, ph]) => {
        title(s, n, nx, ny, nw, 1.402);
        body(s, para, px, py, 9.425, ph);
      });
  },

  // 37 - List with image --------------------------------------------------
  (s) => {
    imageFrame(s, 2.307, 2.773, 9.452, 9.452, 'ellipse');
    checkDisc(s, 13.337, 2.770, 1.127);
    checkDisc(s, 13.337, 8.324, 1.127);
    title(s, 'Top quality', 14.927, 2.795, 9.431, 1.405);
    body(s, 'Pellentesque eget nibh sit nulla ullamcorp eropulvinar sit amet olor sit amet, consewr ect etur adipiscing elit.',
      14.927, 4.361, 9.441, 2.23);
    title(s, 'Best price', 14.927, 8.307, 9.431, 1.405);
    body(s, 'Praesent velofelis vitae metus fini buskasd faucibus. Fusce ac purusolor sit ametotop, consewro ect etur adipiscing elit.',
      14.908, 9.891, 9.45, 2.23);
  },

  // 38 - List of images ---------------------------------------------------
  (s) => {
    tab(s, 'List of Images', 19.0, 5.357);
    imageFrame(s, 2.328, 3.600, 6.263, 6.263, 'ellipse');
    imageFrame(s, 10.182, 3.579, 6.300, 6.300, 'ellipse');
    imageFrame(s, 18.074, 3.579, 6.300, 6.300, 'ellipse');
    [['Donec quis tortor teret rwele fend, viver rawero retonequ, volutpat justwull mamus. ', 2.307, 10.667, 6.280],
     ['Phasellus a imperdip soqe eterlectus, sit ametiwesop rhonc weligu lante.', 10.182, 10.664, 6.304],
     ['Duis ac sem euismod, soda leso magna stibulum tincidu napellentesque.', 18.055, 10.667, 6.320]]
      .forEach(([t, x, y, w]) => body(s, t, x, y, w, 2.23));
  },

  // 39 - Services ---------------------------------------------------------
  (s) => {
    curve(s, BOWL, 8.622, 11.087, 9.453, 3.91, PURPLE);
    title(s, 'What we do?', 2.307, 3.583, 22.008, 1.405, { align: 'center' });
    iconStore(s); iconPriceTag(s); iconGift(s); iconTruck(s);
    [['Store', 3.486, 9.887], ['Price tag', 8.973, 9.887], ['Gift', 14.518, 9.887], ['Delivery', 20.006, 9.900]]
      .forEach(([t, x, y]) => body(s, t, x, y, 3.149, 0.743, { align: 'center' }));
  },

  // 40 - Pricing ----------------------------------------------------------
  (s) => {
    circle(s, 8.607, 3.119, 9.45, PURPLE);
    tab(s, 'Our pricing', 19.774, 4.563);
    [['Basic', '$50', 'Donec quis tortotere sad trwele fend, viverorapow wero retonequ just. ', 2.328, 5.488, 5.491, PURPLE, BLACK, 5.491],
     ['Premium', '$250', 'Phasellus a imperdip rep soqe eterlectus, sitasder ametiwesop rhonc.', 10.576, 5.497, 5.512, WHITE, WHITE, 5.906],
     ['Pro', '$100', 'Duis ac sem euismo por soda leso magnasti sadi bulum tincidu napel.', 18.845, 5.497, 5.494, PURPLE, BLACK, 5.512]]
      .forEach(([plan, price, para, x, hy, w, hc, bc, pw]) => {
        text(s, up(plan), x, hy, w, 0.841, { ...T_KICK, color: hc });
        body(s, para, x, hy === 5.497 && x > 18 ? 6.729 : 6.708, w, 2.23, { color: bc });
        text(s, price, x, 9.858, pw, 1.405, { ...T_TITLE, color: hc });
      });
  },

  // 41 - Vertical table ---------------------------------------------------
  (s) => { verticalTable(s); },

  // 42 - Horizontal table -------------------------------------------------
  (s) => {
    tab(s, 'Price table', 20.035, 4.302);
    horizontalTable(s);
  },

  // 43 - Timeline ---------------------------------------------------------
  (s) => {
    timelineSteps(s, [['12', 1.533], ['13', 5.867], ['14', 10.202]]);
    [3.135, 7.470, 11.804].forEach((y) =>
      text(s, 'Desember', 13.352, y, 3.171, 0.586, { ...T_MONTH, align: 'center' }));
    title(s, 'timeline', 3.117, 4.372, 6.258, 1.405);
    body(s, 'Lorem ipsum dolor sitente amet, consecte turewr dict adipicing elit. Praesentope mot varius nisa eros lobor tisofinibu aecenaswer sado fermentum eusit.',
      3.116, 5.935, 6.258, 4.46);
    [['Cras lacinia, nisl at condim entum viverra, nibh leowero viverra ante pretium.', 2.004],
     ['Nulla suscipit quis elitretop ultrices rutrum. In sed werit mauris at dolor dapibus.', 6.338],
     ['Etiam cursus felis dapibu, blandit tellus at, tinciduntsa quam nteger nunc.', 10.672]]
      .forEach(([t, y]) => body(s, t, 17.31, y, 6.258, 2.23, { valign: 'middle' }));
  },

  // 44 - Timeline with image ----------------------------------------------
  (s) => {
    imageFrame(s, 0, 0, 11.738, 15);
    timelineSteps(s, [['15', 1.533], ['16', 5.867], ['17', 10.202]]);
    [3.135, 7.470, 11.804].forEach((y) =>
      text(s, 'December', 13.352, y, 3.171, 0.586, { ...T_MONTH, align: 'center' }));
    [['Etiam cursus felis dapibus, blandit tellus at, tincidunt quam nteger nunc.', 2.020],
     ['Aenean eget sagittis felis. Donec lacinia pretium ege stas vitae commo.', 6.364],
     ['Donec id massa sodales, mattis velit non, posuerew sapien nisitempor.', 10.659]]
      .forEach(([t, y]) => body(s, t, 17.31, y, 6.258, 2.23, { valign: 'middle' }));
  },

  // 45 - Horizontal timeline ----------------------------------------------
  (s) => {
    curve(s, DOME, 8.661, 0, 9.396, 3.887, PURPLE);
    title(s, 'Timeline', 3.138, 5.138, 20.431, 1.405);
    body(s, 'Pellentesque sapien nisi, tempor sit amet lacus placerat, egestas mollis erat. Vestibulum vitae rutrum nisl. Morbi malesuada turpis et interdum luctus. Donec arcu ipsum, faucibus congue nibh vitae, vehicula feugiat sapien. Suspendisse molestie cursus lectus.',
      3.113, 6.719, 20.456, 2.23);
    timelineRow(s, [['10', 'Desember', 3.088], ['11', 'Desember', 8.878],
      ['12', 'December', 14.669], ['13', 'Desember', 20.459]]);
  },

  // 46 - Horizontal timeline with image -----------------------------------
  (s) => {
    imageCurve(s, DOME, 8.611, 0, 9.396, 3.887);
    timelineRow(s, [['14', 'Desember', 3.088], ['15', 'Desember', 8.878],
      ['16', 'Desember', 14.669], ['17', 'Desember', 20.459]]);
    [['Vestib ulum tinterte ipsu mprimis inerot faucibus orciret.', 2.332, 6.719],
     ['Lorem ipsum doloh sitamet, consecteki turewt adipiscing.', 8.114, 6.719],
     ['Quisque ut porttitor libero, at conguesa malesuada lorem.', 13.904, 6.712],
     ['Curabitur malesusa tad interdum sceorl esquew blandit.', 19.657, 6.737]]
      .forEach(([t, x, y]) => body(s, t, x, y, 4.701, 2.23));
  },

  // 47 - Summary ----------------------------------------------------------
  (s) => {
    imageFrame(s, 13.561, 0, 13.101, 12.226);
    title(s, 'Summary', 3.117, 5.162, 9.428, 1.442);
    body(s, 'Nam nec tortor non ex maximus tempus. Etiam molestie viverra risus, sit amet preti umligula. Donec ac egestas neque. Quisq sollicitudin vulputate ante id mattis. Aenea nonec pretium metus.',
      3.116, 6.726, 9.429, 3.716);
  },

  // 48 - Call to action ---------------------------------------------------
  (s) => {
    imageFrame(s, 0, 2.797, 13.075, 12.203);
    body(s, 'Curabitur et neque elementum, tinciduntas ligula ut, scelerisque augue. Mauris werto interdum, sapien sed feugiat fermentum, sapien ex euismod nisi, ac interdum ligula nunc vitae purus id tristique quam.',
      14.141, 3.612, 9.429, 3.716);
    title(s, ['Everyday', 'living'], 14.161, 8.053, 9.408, 2.808);
  },

  // 49 - Locations --------------------------------------------------------
  (s) => {
    imageFrame(s, 3.882, 4.372, 18.901, 7.875);
    title(s, 'Our branches', 3.882, 2.799, 18.909, 1.402, { align: 'center' });
    // pins: [x, y, filled?]
    [[8.963, 6.568, false], [12.962, 6.008, true], [17.123, 6.626, true],
     [14.693, 5.549, false], [9.611, 9.863, true],
     [7.840, 12.285, false], [14.929, 12.285, true]]
      .forEach(([x, y, filled]) => pin(s, x, y, filled ? PURPLE : WHITE, filled ? WHITE : PURPLE));
    [['New York', 8.807, 7.751, 1.374], ['Paris', 13.105, 7.238, 0.740],
     ['Moscow', 14.566, 6.792, 1.201], ['Hong Kong', 16.763, 7.845, 1.590],
     ['Buenos Aires', 9.099, 11.085, 1.893]]
      .forEach(([t, x, y, w]) => {
        box(s, x, y, w, 0.337, WHITE);
        text(s, t, x, y, w, 0.337, {
          fontFace: HEAD, fontSize: 20, color: PURPLE, lineSpacing: 24,
          align: 'center', valign: 'middle',
        });
      });
    body(s, 'Branch Offices', 8.901, 12.776, 4.431, 0.743);
    body(s, 'Field Offices', 15.901, 12.776, 3.731, 0.743);
  },

  // 50 - Follow us --------------------------------------------------------
  (s) => {
    curve(s, BOWL, 8.622, 11.087, 9.453, 3.91, PURPLE);
    title(s, 'Follow us', 3.882, 3.581, 18.899, 1.402, { align: 'center' });
    [['\uf39e', 5.351, 5.961, 1.790, 'Facebook', 3.881, 9.548, 4.714],
     ['\uf16d', 12.065, 5.958, 2.534, 'Instagram', 10.965, 9.508, 4.733],
     ['\uf099', 18.807, 5.965, 3.229, 'Twitter', 18.625, 9.508, 3.625]]
      .forEach(([glyph, gx, gy, gw, label, lx, ly, lw]) => {
        text(s, glyph, gx, gy, gw, 3.13,
          { fontFace: FA_BRAND, fontSize: 200, color: PURPLE, align: 'center' });
        body(s, label, lx, ly, lw, 0.743, { align: 'center' });
      });
  },

  // 51 - Contacts ---------------------------------------------------------
  (s) => {
    curve(s, BOWL, 8.622, 11.087, 9.453, 3.91, PURPLE);
    tab(s, 'Contact us', 19.767, 4.57);
    title(s, 'Thank you', 7.819, 5.16, 11.026, 1.402, { align: 'center' });
    body(s, ['Address:', 'Phone:', 'Website:'], 7.819, 6.735, 2.83, 2.23, { align: 'right' });
    body(s, ['123 Street, City, State 45678', '1234 56 7890', 'www.companyname.com'],
      10.915, 6.735, 7.931, 2.23);
  },
];

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */
function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'FASHION', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'FASHION';
  pptx.author = 'admin';
  pptx.title = 'PowerPoint Presentation';

  SLIDES.forEach((buildSlide, i) => {
    const slide = pptx.addSlide();
    slide.background = { color: WHITE };
    buildSlide(slide, pptx);
    logo(slide);
    if (i > 0) pageNum(slide, i + 1);
  });

  return pptx.writeFile({
    fileName: path.join(__dirname, '0c730ee8-68d5-4023-a1c9-b46f56814dd9_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
