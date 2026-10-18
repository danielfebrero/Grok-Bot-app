#!/usr/bin/env node
/**
 * "Line Chart Infographic" deck — recreated with pptxgenjs.
 * 15 slides, 13.333" x 7.5" (16:9).
 *
 * Raster icons in the source deck (lightbulb / line-chart / presentation /
 * stopwatch glyphs) are replaced with plain vector placeholders drawn from
 * pptxgenjs primitives — no embedded image data anywhere in this file.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Theme
 * ------------------------------------------------------------------ */

const C = {
  teal: '03738C',       // accent1
  blue: '117DBF',       // accent2
  darkTeal: '026873',   // accent4
  white: 'FFFFFF',
  ink: '262626',        // tx1 lumMod 85%
  body: '404040',       // tx1 lumMod 75%
  grey: '595959',       // tx1 lumMod 65%
  midGrey: '808080',
  slate: '44546A',      // tx2
  track: 'F2F2F2',      // bg1 lumMod 95% (progress-bar track)
  hair: 'D0CFCF',       // hand-drawn chart rules
  gridline: 'E8E8E8',
};

const FONT = 'Space Grotesk';

// pptxgenjs rewrites shadow objects in place while rendering, so each shape
// needs its own copy — hence factories rather than shared constants.
/** Soft drop shadow used by every white "card" in the deck. */
const cardShadow = () => ({ type: 'outer', color: '000000', opacity: 0.12, blur: 40, offset: 10, angle: 67 });
/** Tighter shadow used by callout bubbles / pills. */
const pillShadow = () => ({ type: 'outer', color: '000000', opacity: 0.15, blur: 22, offset: 5, angle: 90 });
/** Wide halo behind the circular icon badges. */
const haloShadow = () => ({ type: 'outer', color: '000000', opacity: 0.1, blur: 80, offset: 30, angle: 35 });

const LOREM = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Donec et orci feugiat erat pharetra scelerisque';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.';
const LOREM_VEHICULA = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. In vehicula sem sit amet';
const LOREM_TEMPOR = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore';
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/* ------------------------------------------------------------------ *
 * Small helpers
 * ------------------------------------------------------------------ */

/** White rounded "card" with the deck's signature soft shadow. */
function card(slide, x, y, w, h, opts = {}) {
  slide.addShape('roundRect', {
    x, y, w, h,
    fill: { color: opts.fill || C.white },
    line: { type: 'none' },
    rectRadius: opts.radius === undefined ? 0.22 : opts.radius,
    shadow: opts.shadow === undefined ? cardShadow() : opts.shadow,
  });
}

/** Text box. `t` is a string or an array of {text, options} runs. */
function text(slide, t, o) {
  // margin order in pptxgenjs is [left, right, bottom, top] in points.
  slide.addText(t, Object.assign({ fontFace: FONT, margin: [7.2, 7.2, 3.6, 3.6] }, o));
}

/** Filled circle. */
function dot(slide, x, y, d, color, shadow) {
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color }, line: { type: 'none' }, shadow });
}

/** Two-tone heading, e.g. "Financial Growth" / "Report Analysis". */
function heading(slide, first, second, o) {
  text(slide, [
    { text: first, options: { color: C.teal } },
    { text: second, options: { color: C.ink } },
  ], Object.assign({ w: 5.481, h: 0.841, fontSize: 44, bold: true, valign: 'top' }, o));
}

/**
 * "• Line Chart Infographic" eyebrow: teal disc + chart glyph + label.
 * `align` is the label alignment inside its 2.628" box.
 */
function eyebrow(slide, x, y, o = {}) {
  const rot = o.rotate || 0;
  dot(slide, x, y + 0.007, 0.357, C.teal);
  glyphChart(slide, x + 0.086, y + 0.092, 0.185, C.white);
  text(slide, 'Line Chart Infographic', {
    x: x + 0.357, y, w: 2.628, h: 0.37, fontSize: 16, bold: true, color: C.ink,
    align: o.align || 'left', wrap: false, rotate: rot, valign: 'top',
  });
}

/** Horizontal progress bar: light track + coloured fill. */
function progressBar(slide, x, y, w, h, pct, color) {
  slide.addShape('rect', { x, y, w, h, fill: { color: C.track, transparency: 12 }, line: { type: 'none' } });
  slide.addShape('rect', { x, y, w: w * pct, h, fill: { color }, line: { type: 'none' } });
}

/** Rounded pill with centred bold label (the deck's "2025" / "Your Text" chips). */
function pill(slide, x, y, w, h, label, o = {}) {
  slide.addShape('roundRect', {
    x, y, w, h,
    fill: { color: o.fill || C.white },
    line: { type: 'none' },
    rectRadius: o.radius === undefined ? h / 2 : o.radius,
    shadow: o.shadow === undefined ? cardShadow() : o.shadow,
  });
  text(slide, label, {
    x, y, w, h, align: 'center', valign: 'middle', bold: true,
    fontSize: o.fontSize || 14, color: o.color || C.teal,
  });
}

/* ------------------------------------------------------------------ *
 * Vector stand-ins for the deck's raster icons
 * ------------------------------------------------------------------ */

/** Bar-chart-in-a-frame glyph (the eyebrow badge icon). */
function glyphChart(slide, x, y, d, color) {
  const ln = { color, width: 1 };
  slide.addShape('line', { x, y, w: 0, h: d, line: ln });
  slide.addShape('line', { x, y: y + d, w: d, h: 0, line: ln });
  [[0.22, 0.42], [0.46, 0.20], [0.70, 0.55]].forEach(([fx, fh]) => {
    slide.addShape('rect', {
      x: x + d * fx, y: y + d * (0.88 - fh), w: d * 0.16, h: d * fh,
      fill: { color }, line: { type: 'none' },
    });
  });
}

/** Lightbulb glyph. */
function glyphBulb(slide, x, y, d, color) {
  slide.addShape('ellipse', {
    x: x + d * 0.2, y: y + d * 0.04, w: d * 0.6, h: d * 0.6,
    fill: { type: 'none' }, line: { color, width: Math.max(1, d * 2.2) },
  });
  for (let i = 0; i < 3; i++) {
    slide.addShape('rect', {
      x: x + d * (0.34 + i * 0.02), y: y + d * (0.66 + i * 0.11), w: d * (0.32 - i * 0.04), h: d * 0.055,
      fill: { color }, line: { type: 'none' },
    });
  }
}

/** Downward zig-zag line inside a corner frame (the "trend" icon). */
function glyphTrend(slide, x, y, d, color) {
  const ln = { color, width: Math.max(1, d * 2.2) };
  slide.addShape('line', { x, y: y + d * 0.08, w: 0, h: d * 0.84, line: ln });
  slide.addShape('line', { x, y: y + d * 0.92, w: d, h: 0, line: ln });
  freeform(slide, x + d * 0.14, y + d * 0.2, d * 0.74, d * 0.5,
    'M 0 0.15 L 0.3 0.7 L 0.52 0.34 L 1 1 M 1 1 L 1 0.32',
    { fill: { type: 'none' }, line: ln });
}

/** Presentation board on a tripod, with bars. */
function glyphBoard(slide, x, y, d, color) {
  const ln = { color, width: Math.max(1, d * 2.0) };
  slide.addShape('rect', {
    x: x + d * 0.08, y: y + d * 0.06, w: d * 0.84, h: d * 0.58,
    fill: { type: 'none' }, line: ln,
  });
  [[0.26, 0.18], [0.46, 0.34], [0.66, 0.26]].forEach(([fx, fh]) => {
    slide.addShape('rect', {
      x: x + d * fx, y: y + d * (0.56 - fh), w: d * 0.11, h: d * fh,
      fill: { color }, line: { type: 'none' },
    });
  });
  slide.addShape('line', { x: x + d * 0.5, y: y + d * 0.64, w: 0, h: d * 0.16, line: ln });
  slide.addShape('line', { x: x + d * 0.18, y: y + d * 0.96, w: d * 0.32, h: -d * 0.16, line: ln });
  slide.addShape('line', { x: x + d * 0.5, y: y + d * 0.8, w: d * 0.32, h: d * 0.16, line: ln });
}

/** Stopwatch. */
function glyphTimer(slide, x, y, d, color) {
  const ln = { color, width: Math.max(1, d * 2.0) };
  slide.addShape('ellipse', {
    x: x + d * 0.06, y: y + d * 0.14, w: d * 0.88, h: d * 0.8,
    fill: { type: 'none' }, line: ln,
  });
  slide.addShape('rect', { x: x + d * 0.38, y, w: d * 0.24, h: d * 0.1, fill: { color }, line: { type: 'none' } });
  slide.addShape('line', { x: x + d * 0.5, y: y + d * 0.54, w: 0, h: -d * 0.22, line: ln });
  slide.addShape('line', { x: x + d * 0.5, y: y + d * 0.54, w: d * 0.2, h: d * 0.12, line: ln });
}

/**
 * Freeform shape from a compact path string in unit space (0..1 on both axes),
 * scaled to the shape's own width/height.
 */
function freeform(slide, x, y, w, h, d, opts) {
  slide.addShape('custGeom', Object.assign({ x, y, w, h, points: parsePath(d, w, h) }, opts));
}

/** Circular badge = coloured disc + white glyph, as used all over the deck. */
function iconBadge(slide, glyph, x, y, d, color, o = {}) {
  dot(slide, x, y, d, color, o.shadow);
  const g = d * (o.scale || 0.44);
  glyph(slide, x + (d - g) / 2, y + (d - g) / 2, g, C.white);
}

/* ------------------------------------------------------------------ *
 * Chart helpers
 * ------------------------------------------------------------------ */

/** Common look for every chart: transparent plot area, subtle labels. */
function chartBase(extra) {
  return Object.assign({
    showLegend: false,
    showTitle: false,
    chartArea: { fill: { type: 'none' } },
    plotArea: { fill: { type: 'none' } },
    catAxisLineShow: false,
    valAxisLineShow: false,
    catGridLine: { style: 'none' },
    valGridLine: { style: 'none' },
    catAxisLabelFontFace: FONT,
    valAxisLabelFontFace: FONT,
    catAxisLabelColor: C.body,
    valAxisLabelColor: C.body,
    catAxisLabelFontSize: 12,
    valAxisLabelFontSize: 12,
  }, extra);
}

/** Inner plot-area rectangles (fractions of the chart frame) used by the deck. */
const PLOT = {
  wave: { x: 0.0495, y: 0.1172, w: 0.8688, h: 0.7459 },
  wide: { x: 0.0555, y: 0.1925, w: 0.9158, h: 0.6765 },
  full: { x: 0.0253, y: 0.0377, w: 0.9532, h: 0.8461 },
  tall: { x: 0.0448, y: 0.0978, w: 0.9230, h: 0.8292 },
  inset: { x: 0.1181, y: 0.0448, w: 0.8460, h: 0.8204 },
};

/**
 * PowerPoint's "stacked" line grouping draws each series on top of the previous
 * one; pptxgenjs has no equivalent, so the running totals are computed here.
 */
function stackSeries(series) {
  const running = series[0].values.map(() => 0);
  return series.map(s => {
    const values = s.values.map((v, i) => (running[i] += v));
    return Object.assign({}, s, { values });
  });
}

/**
 * Smoothed scatter line (the deck's "wave" charts). X values are 1..n so the
 * numeric bottom axis reads 0, 2, 4 … exactly like the original.
 */
function waveChart(slide, series, pos, opts = {}) {
  const n = series[0].values.length;
  const xs = [];
  for (let i = 1; i <= n; i++) xs.push(i);
  const data = [{ name: 'X', values: xs }].concat(series.map(s => ({ name: s.name, values: s.values })));
  slide.addChart('scatter', data, chartBase(Object.assign({
    lineSmooth: true,
    lineSize: 2.25,
    lineDataSymbol: 'none',
    chartColors: series.map(s => s.color),
    valAxisMajorUnit: 20,
    catAxisMaxVal: opts.xMax,
    catAxisMinVal: opts.xMin,
    valAxisMaxVal: opts.yMax,
    catAxisLabelFontSize: opts.labelSize || 12,
    valAxisLabelFontSize: opts.labelSize || 12,
    layout: opts.layout || PLOT.wave,
    x: pos.x, y: pos.y, w: pos.w, h: pos.h,
  }, opts.chart || {})));
}

/**
 * Multi-series line chart where each series keeps its own colour/marker —
 * built as a multi-type chart so pptxgenjs emits one <c:ser> per style.
 */
function lineChart(slide, series, pos, common) {
  const parts = series.map(s => ({
    type: 'line',
    data: [{ name: s.name, labels: [s.labels], values: s.values }],
    options: {
      chartColors: [s.color],
      lineSize: s.lineSize === undefined ? 2.25 : s.lineSize,
      lineDataSymbol: s.symbol || 'circle',
      lineDataSymbolSize: s.symbolSize || 6,
      lineDataSymbolLineColor: s.symbolLine || s.color,
      lineDataSymbolLineSize: s.symbolLineSize || 0.75,
      lineSmooth: !!s.smooth,
    },
  }));
  slide.addChart(parts, chartBase(Object.assign({ x: pos.x, y: pos.y, w: pos.w, h: pos.h }, common)));
}

/* ------------------------------------------------------------------ *
 * Repeated slide furniture
 * ------------------------------------------------------------------ */

/** Centred eyebrow + centred "Financial Growth" heading (top of many slides). */
function centredTitle(slide, first, second, y) {
  eyebrow(slide, 5.174, y, { align: 'center' });
  heading(slide, first, second, { x: 3.926, y: y + 0.428, align: 'center' });
}

/** Icon badge + bold label + grey paragraph, laid out in a row. */
function iconTextRow(slide, o) {
  iconBadge(slide, o.glyph, o.x, o.y, o.d, o.color, { shadow: o.shadow });
  if (o.value) {
    text(slide, o.value, {
      x: o.tx, y: o.vy, w: 1.085, h: 0.505, fontSize: o.valueSize || 24,
      bold: true, color: o.color, valign: 'top',
    });
  }
  if (o.label) {
    text(slide, o.label, {
      x: o.lx, y: o.ly, w: o.lw || 1.085, h: 0.37, fontSize: 16,
      bold: o.labelBold !== false, color: o.labelColor || o.color, valign: 'top',
    });
  }
  if (o.body) {
    text(slide, o.body, {
      x: o.tx, y: o.by, w: o.bw, h: o.bh || 0.607, fontSize: 12,
      color: C.body, lineSpacingMultiple: 1.3, valign: 'top',
    });
  }
}

/* ------------------------------------------------------------------ *
 * Slide 1 — Financial Growth, two wave charts
 * ------------------------------------------------------------------ */

function slide01(pptx) {
  const s = pptx.addSlide();
  centredTitle(s, 'Financial ', 'Growth', 0.993);

  const panels = [
    {
      x: 0.905, chartX: 1.049, chartY: 2.485, color: C.teal, radius: 0.216,
      values: [10, 15, 5, 38, 25, 35, 58, 48.14, 35, 20, 35, 10],
      badge: { x: 3.496, y: 2.701, label: '+10', color: C.teal, w: 0.773 },
      title: 'Data Chart 01', tx: 2.115, icon: { x: 0.905, glyph: glyphBulb, iconColor: C.teal },
    },
    {
      x: 6.873, chartX: 7.017, chartY: 2.557, color: C.blue, radius: 0.187,
      values: [3, 5, 6, 8, 10, 8, 15, 60.71, 80, 80, 86, 94.43],
      badge: { x: 10.228, y: 2.603, label: '-20', color: C.blue, w: 0.773 },
      title: 'Data Chart 02', tx: 8.084, icon: { x: 6.873, glyph: glyphTrend, iconColor: C.blue },
    },
  ];

  panels.forEach(p => {
    card(s, p.x, 2.446, 5.603, 2.613, { radius: p.radius });
    waveChart(s, [{ name: 'Series', values: p.values, color: p.color }],
      { x: p.chartX, y: p.chartY, w: 5.315, h: 2.16 });

    // Value callout bubble over the chart.
    s.addShape('wedgeRoundRectCallout', {
      x: p.badge.x, y: p.badge.y, w: p.badge.w, h: 0.406,
      fill: { color: C.white }, line: { type: 'none' }, shadow: pillShadow(),
    });
    text(s, p.badge.label, {
      x: p.badge.x, y: p.badge.y, w: p.badge.w, h: 0.406,
      align: 'center', valign: 'middle', fontSize: 14, bold: true, color: p.badge.color,
    });

    // Footer: icon badge, title, body copy.
    iconBadge(s, p.icon.glyph, p.icon.x, 5.536, 0.879, p.icon.iconColor, { shadow: haloShadow(), scale: 0.59 });
    text(s, p.title, { x: p.tx, y: 5.486, w: 1.915, h: 0.37, fontSize: 16, bold: true, color: p.color, valign: 'top' });
    text(s, LOREM, { x: p.tx, y: 5.808, w: 4.386, h: 0.607, fontSize: 12, color: C.body, lineSpacingMultiple: 1.3, valign: 'top' });
  });
}

/* ------------------------------------------------------------------ *
 * Slide 2 — Report Analysis, three-series chart + project list + bars
 * ------------------------------------------------------------------ */

function slide02(pptx) {
  const s = pptx.addSlide();

  card(s, 1.072, 0.886, 7.769, 3.466, { radius: 0.236 });
  waveChart(s, [
    { name: 'Series 1', values: [1, 5, 2.1, 4.5, 2.4, 4.4, 1.8, 2.8, 2, 2, 3, 3], color: C.teal },
    { name: 'Series 2', values: [3, 2, 0, 1.5, 5, 4, 3.8, 2.3, 1, 4.5, 3, 1], color: C.blue },
    { name: 'Series 3', values: [0, 2, 3, 5, 5, 4, 3, 2, 1, 2, 3, 4], color: C.teal },
  ], { x: 1.26, y: 0.974, w: 7.392, h: 3.274 }, {
    layout: PLOT.wide,
    chart: {
      valAxisMajorUnit: 1,
      valGridLine: { color: C.gridline, style: 'dash', size: 0.75 },
      catAxisLineShow: true,
      catAxisLineColor: C.gridline,
      catAxisLineSize: 0.75,
    },
  });
  text(s, 'Transaction In Month', { x: 1.262, y: 0.974, w: 2.87, h: 0.344, fontSize: 12, color: C.teal, valign: 'top' });

  // Right card: timeline of three projects.
  card(s, 9.001, 0.886, 3.26, 3.466, { radius: 0.236 });
  s.addShape('line', { x: 9.369, y: 1.483, w: 0, h: 1.98, line: { color: C.track, width: 1 } });
  const projects = [
    { y: 1.178, name: 'Project 01', color: C.slate, dotColor: C.slate },
    { y: 2.19, name: 'Project 02', color: C.teal, dotColor: C.teal },
    { y: 3.181, name: 'Project 03', color: C.slate, dotColor: C.slate },
  ];
  projects.forEach((pr, i) => {
    dot(s, 9.282, 1.463 + i, 0.175, pr.dotColor);
    text(s, pr.name, { x: 9.617, y: pr.y, w: 2.162, h: 0.37, fontSize: 16, bold: true, color: pr.color, valign: 'top' });
    text(s, 'Lorem ipsum dolor sit amet, ', {
      x: 9.617, y: pr.y + 0.383, w: 2.424, h: 0.344, fontSize: 12, color: C.body,
      lineSpacingMultiple: 1.3, valign: 'top',
    });
  });

  // Bottom-left: three labelled progress bars.
  const bars = [
    { y: 4.886, label: 'Marketing Data 1', pct: 0.725, value: '85%', color: C.teal },
    { y: 5.506, label: 'Marketing Data 2', pct: 0.622, value: '67%', color: C.blue },
    { y: 6.062, label: 'Marketing Data 2', pct: 0.622, value: '67%', color: C.blue },
  ];
  bars.forEach(b => {
    text(s, b.label, { x: 1.152, y: b.y, w: 2.211, h: 0.337, fontSize: 14, bold: true, color: b.color, valign: 'top' });
    text(s, b.value, { x: 5.444, y: b.y + 0.002, w: 1.328, h: 0.367, fontSize: 14, bold: true, color: b.color, align: 'right', valign: 'top' });
    progressBar(s, 1.238, b.y + 0.404, 5.45, 0.059, b.pct, b.color);
  });
  glyphBoard(s, 5.485, 5.02, 0.405, C.track);

  // Bottom-right: heading block.
  eyebrow(s, 7.33, 4.887);
  heading(s, 'Report ', 'Analysis', { x: 7.211, y: 5.321 });
  text(s, LOREM_SHORT, { x: 7.242, y: 6.255, w: 4.908, h: 0.344, fontSize: 12, color: C.body, lineSpacingMultiple: 1.3, valign: 'top' });
}

/* ------------------------------------------------------------------ *
 * Slide 3 — three data cards + Data Analysis line chart
 * ------------------------------------------------------------------ */

function slide03(pptx) {
  const s = pptx.addSlide();

  const rows = [
    { y: 0.905, radius: 0.209, value: '100K', label: 'Data 01', color: C.teal, glyph: glyphBulb },
    { y: 2.85, radius: 0.237, value: '200K', label: 'Data 02', color: C.blue, glyph: glyphTrend },
    { y: 4.794, radius: 0.237, value: '300K', label: 'Data 03', color: C.teal, glyph: glyphBoard },
  ];
  rows.forEach(r => {
    card(s, 1.008, r.y, 5.62, 1.796, { radius: r.radius });
    iconTextRow(s, {
      glyph: r.glyph, x: 1.309, y: r.y + 0.47, d: 0.618, color: r.color, shadow: haloShadow(),
      value: r.value, tx: 2.115, vy: r.y + 0.333,
      label: r.label, lx: 3.103, ly: r.y + 0.385,
      body: LOREM, by: r.y + 0.806, bw: 4.386,
    });
  });

  // Right card: "Data Analysis" three-series line chart with legend on top.
  card(s, 6.823, 0.91, 5.502, 5.68, { radius: 0.225 });
  text(s, 'Data Analysis', { x: 7.095, y: 1.096, w: 4.574, h: 0.37, fontSize: 16, bold: true, color: C.teal, align: 'center', valign: 'top' });
  lineChart(s, stackSeries([
    { name: 'Cost', labels: ['Qr 1', 'Qr 2', 'Qr 3', 'Qr 4'], values: [4.3, 2.5, 3.5, 4.5], color: C.teal, symbol: 'diamond', symbolSize: 8, lineSize: 1.75 },
    { name: 'Profit', labels: ['Qr 1', 'Qr 2', 'Qr 3', 'Qr 4'], values: [2.4, 4.4, 1.8, 2.8], color: C.blue, symbol: 'square', symbolSize: 8, lineSize: 1.75 },
    { name: 'Supply', labels: ['Qr 1', 'Qr 2', 'Qr 3', 'Qr 4'], values: [2, 2, 3, 5], color: C.teal, symbol: 'triangle', symbolSize: 8, lineSize: 1.75 },
  ]), { x: 7.107, y: 1.628, w: 4.897, h: 4.777 }, {
    valAxisMaxVal: 14,
    valAxisMajorUnit: 2,
    showLegend: true,
    legendPos: 't',
    layout: { x: 0.06, y: 0.14, w: 0.9, h: 0.78 },
    legendFontFace: FONT,
    legendFontSize: 12,
    legendColor: C.midGrey,
    catAxisLabelColor: C.midGrey,
    valAxisLabelColor: C.midGrey,
    catGridLine: { color: C.gridline, size: 0.75 },
  });

  // Value callouts pinned over the plot.
  const callouts = [
    { x: 8.566, y: 2.748, w: 1.061, h: 0.5, tri: 8.994, triY: 3.23, label: '+15%', color: C.teal, size: 16 },
    { x: 10.627, y: 3.214, w: 1.023, h: 0.482, tri: 11.036, triY: 3.687, label: '+30%', color: C.blue, size: 14 },
    { x: 8.603, y: 4.556, w: 0.984, h: 0.463, tri: 8.993, triY: 4.975, label: '-20%', color: C.teal, size: 16 },
  ];
  callouts.forEach(co => {
    s.addShape('roundRect', {
      x: co.x, y: co.y, w: co.w, h: co.h, fill: { color: co.color },
      line: { type: 'none' }, rectRadius: co.h / 6, shadow: pillShadow(),
    });
    s.addShape('triangle', {
      x: co.tri, y: co.triY, w: 0.206, h: 0.212, flipV: true,
      fill: { color: co.color }, line: { type: 'none' },
    });
    text(s, co.label, {
      x: co.x, y: co.y, w: co.w, h: co.h, align: 'center', valign: 'middle',
      fontSize: co.size, bold: true, color: C.white,
    });
  });
}

/* ------------------------------------------------------------------ *
 * Slide 4 — Data Growth (29-point chart) + stat strip, rotated title
 * ------------------------------------------------------------------ */

function slide04(pptx) {
  const s = pptx.addSlide();

  card(s, 1.288, 0.981, 8.6, 3.677, { radius: 0.238 });
  text(s, 'Data Growth', { x: 1.592, y: 1.109, w: 7.992, h: 0.404, fontSize: 18, bold: true, color: C.teal, align: 'center', valign: 'top' });

  const days = [];
  for (let i = 1; i <= 29; i++) days.push('Day ' + i);
  lineChart(s, [
    { name: 'Category A', color: C.blue, labels: days, symbol: 'circle', symbolSize: 6, symbolLine: C.blue, lineSize: 2,
      values: [12, 23, 45, 54, 76, 87, 65, 43, 23, 22, 12, 4, 32, 7, 9, 87, 65, 144, 134, 104, 110, 133, 50, 77, 40, 22, 12, 8, 10] },
    { name: 'Category B', color: C.teal, labels: days, symbol: 'circle', symbolSize: 6, symbolLine: C.teal, lineSize: 2,
      values: [33, 21, 100, 89, 65, 67, 78, 65, 150, 40, 22, 67, 7, 30, 33, 40, 80, 105, 80, 70, 40, 100, 35, 39, 37, 51, 30, 25, 20] },
  ], { x: 1.501, y: 1.669, w: 8.173, h: 2.736 }, {
    layout: PLOT.full,
    valAxisMaxVal: 160,
    valAxisMajorUnit: 40,
    catAxisLabelFrequency: '29',
    catAxisLabelFontSize: 11,
    valAxisLabelFontSize: 11,
    catAxisLabelColor: C.midGrey,
    valAxisLabelColor: C.midGrey,
    catAxisLineShow: true,
    catAxisLineColor: 'BFBFBF',
    valAxisLineShow: true,
    valAxisLineColor: 'BFBFBF',
  });

  // Two grey stats bottom-left.
  [
    { x: 1.218, value: '167K', color: C.teal },
    { x: 3.509, value: '700K', color: C.blue },
  ].forEach(st => {
    text(s, st.value, { x: st.x, y: 5.071, w: 1.339, h: 0.572, fontSize: 28, bold: true, color: st.color, valign: 'top' });
    text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. ', {
      x: st.x, y: 5.666, w: 2.072, h: 0.812, fontSize: 12, color: C.body, lineSpacingMultiple: 1.2, valign: 'top',
    });
  });

  // Teal stat strip.
  s.addShape('roundRect', {
    x: 5.8, y: 5.089, w: 6.33, h: 1.43, fill: { color: C.teal }, line: { type: 'none' }, rectRadius: 0.238,
  });
  [
    { x: 6.089, lx: 6.033, value: '1500+', label: 'Customers' },
    { x: 8.159, lx: 8.103, value: '120+', label: 'Product' },
    { x: 10.228, lx: 10.172, value: '93%', label: 'Testimonials' },
  ].forEach(st => {
    text(s, st.value, { x: st.x, y: 5.389, w: 1.613, h: 0.64, fontSize: 32, bold: true, color: C.white, align: 'center', valign: 'top' });
    text(s, st.label, { x: st.lx, y: 5.984, w: 1.724, h: 0.303, fontSize: 12, color: C.white, align: 'center', valign: 'top' });
  });

  // Rotated heading down the right edge.
  heading(s, 'Report ', 'Analysis', { x: 9.639, y: 2.084, w: 3.532, h: 1.582, rotate: 270 });
  dot(s, 10.251, 4.166, 0.357, C.teal);
  glyphChart(s, 10.336, 4.251, 0.185, C.white);
  text(s, 'Line Chart Infographic', {
    x: 9.115, y: 2.666, w: 2.628, h: 0.37, fontSize: 16, bold: true, color: C.ink, rotate: 270, wrap: false, valign: 'top',
  });
}

/* ------------------------------------------------------------------ *
 * Slide 5 — Data Sales card + options card + numbered notes
 * ------------------------------------------------------------------ */

function slide05(pptx) {
  const s = pptx.addSlide();

  // Left: "Data Sales" header pill + big card with chart.
  pill(s, 1.145, 1.011, 5.384, 0.55, 'Data Sales', { fill: C.teal, color: C.white, fontSize: 20 });
  card(s, 1.145, 1.703, 5.384, 4.921, { radius: 0.199 });
  text(s, '888,222', { x: 2.308, y: 1.896, w: 3.057, h: 0.64, fontSize: 32, bold: true, color: C.teal, align: 'center', valign: 'top' });
  text(s, [
    { text: '\u2191 ', options: { color: C.blue } },
    { text: 'Increased 50%', options: { color: C.midGrey } },
  ], { x: 2.192, y: 2.453, w: 3.289, h: 0.34, fontSize: 12, align: 'center', lineSpacingMultiple: 1.3, valign: 'top' });

  lineChart(s, [
    { name: 'Series 1', color: C.teal, labels: MONTHS, symbol: 'circle', symbolSize: 6, symbolLine: C.teal,
      values: [0, 500, 600, 500, 1000, 500, 300, 900, 874, 770, 650, 700] },
    { name: 'Series 2', color: C.blue, labels: MONTHS, symbol: 'circle', symbolSize: 6, symbolLine: C.white,
      values: [100, 800, 654, 600, 400, 200, 650, 750, 780, 600, 840, 600] },
  ], { x: 1.333, y: 2.804, w: 5.006, h: 2.956 }, {
    catAxisHidden: true,
    valAxisHidden: true,
    layout: PLOT.tall,
    valAxisMaxVal: 1000,
    valAxisMajorUnit: 100,
    valGridLine: { color: 'F2F2F2', size: 0.5 },
  });

  pill(s, 1.333, 5.948, 2.419, 0.394, 'Selling', { fill: C.teal, color: C.white, fontSize: 16 });
  pill(s, 3.921, 5.948, 2.419, 0.394, 'Buying', { fill: C.blue, color: C.white, fontSize: 16 });

  // Right: heading + options card.
  eyebrow(s, 7.138, 1.066);
  heading(s, 'Report ', 'Analysis', { x: 7.019, y: 1.493 });

  card(s, 7.019, 2.478, 5.215, 2.399, { radius: 0.232 });
  [
    { ox: 7.403, cx: 7.476, tx: 8.214, pct: '80%', title: 'Option One', color: C.teal, glyph: glyphBulb },
    { ox: 10.048, cx: 10.121, tx: 10.859, pct: '90%', title: 'Option Two', color: C.blue, glyph: glyphTrend },
  ].forEach(op => {
    iconBadge(s, op.glyph, op.cx, 2.771, 0.634, op.color, { scale: 0.61 });
    text(s, op.pct, { x: op.tx, y: 2.847, w: 0.842, h: 0.438, fontSize: 20, bold: true, color: op.color, valign: 'top' });
    text(s, op.title, { x: op.ox, y: 3.625, w: 1.56, h: 0.404, fontSize: 18, bold: true, color: op.color, wrap: false, valign: 'top' });
    text(s, 'Lorem ipsum dolor sit amet, consectetur', {
      x: op.ox, y: 4.032, w: 2.141, h: 0.607, fontSize: 12, color: C.body, lineSpacingMultiple: 1.3, valign: 'top',
    });
  });

  // Numbered notes.
  [
    { n: '01', y: 5.193, ty: 5.292, by: 5.161, color: C.teal },
    { n: '02', y: 6.027, ty: 6.145, by: 5.995, color: C.blue },
  ].forEach(nt => {
    dot(s, 7.121, nt.y, 0.549, nt.color, haloShadow());
    text(s, nt.n, { x: 7.043, y: nt.ty, w: 0.682, h: 0.344, fontSize: 12, bold: true, color: C.white, align: 'center', lineSpacingMultiple: 1.3, valign: 'top' });
    text(s, LOREM, { x: 7.862, y: nt.by, w: 4.372, h: 0.607, fontSize: 12, color: C.body, lineSpacingMultiple: 1.3, valign: 'top' });
  });
}

/* ------------------------------------------------------------------ *
 * Slide 6 — pricing card + wave chart
 * ------------------------------------------------------------------ */

function slide06(pptx) {
  const s = pptx.addSlide();
  centredTitle(s, 'Financial ', 'Growth', 0.993);

  // Teal pricing panel.
  s.addShape('roundRect', {
    x: 1.376, y: 2.466, w: 5.003, h: 2.981, fill: { color: C.teal }, line: { type: 'none' }, rectRadius: 0.262,
  });
  text(s, 'Basic Package', { x: 1.682, y: 2.823, w: 3.257, h: 0.598, fontSize: 18, bold: true, color: C.white, valign: 'top' });
  text(s, '$15', { x: 1.682, y: 3.604, w: 1.133, h: 0.774, fontSize: 40, bold: true, color: C.white, wrap: false, valign: 'top' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. ', {
    x: 3.101, y: 3.679, w: 3.161, h: 0.607, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3, valign: 'top',
  });
  text(s, 'See More', { x: 1.682, y: 4.652, w: 1.082, h: 0.337, fontSize: 14, color: C.white, wrap: false, valign: 'top' });
  s.addShape('line', {
    x: 3.286, y: 4.876, w: 0.704, h: 0,
    line: { color: C.white, width: 1, endArrowType: 'triangle' },
  });

  // Right: chart card with "787K" callout.
  card(s, 6.685, 2.466, 5.425, 3.949, { radius: 0.232 });
  waveChart(s, [
    { name: 'Series 1', values: [10, 15, 5, 38, 25, 35, 58, 35, 35, 20, 35, 10], color: C.blue },
    { name: 'Series 2', values: [4, 8, 9, 10, 15, 12, 8, 55, 65, 65, 65, 69], color: C.teal },
  ], { x: 6.952, y: 2.856, w: 4.74, h: 3.382 }, { xMin: 0, xMax: 14 });
  s.addShape('wedgeRoundRectCallout', {
    x: 10.584, y: 2.889, w: 1.046, h: 0.47, fill: { color: C.teal }, line: { type: 'none' },
  });
  text(s, '787K', { x: 10.584, y: 2.889, w: 1.046, h: 0.47, align: 'center', valign: 'middle', fontSize: 16, bold: true, color: C.white });

  // Footer chip + note.
  pill(s, 1.391, 5.913, 1.413, 0.458, '2025', { fill: C.teal, color: C.white, fontSize: 18, radius: 0.08, shadow: null });
  text(s, LOREM_VEHICULA, { x: 2.986, y: 5.808, w: 3.681, h: 0.607, fontSize: 12, color: C.body, lineSpacingMultiple: 1.3, valign: 'top' });
}

/* ------------------------------------------------------------------ *
 * Slide 7 — two hand-drawn month charts + progress list
 * ------------------------------------------------------------------ */

/** Hand-built mini chart: 3 horizontal rules, month labels, dot row, curve. */
function monthMiniChart(s, o) {
  card(s, o.x, o.y, 5.481, 2.806, { radius: o.radius });

  // Y labels (Mar / Feb / Jan, top to bottom) and the three rules.
  ['Mar', 'Feb', 'Jan'].forEach((m, i) => {
    text(s, m, {
      x: o.x + 0.15, y: o.y + 0.564 + i * 0.631, w: 0.42, h: 0.303, fontSize: 12,
      color: C.body, align: o.yAlign || 'center', wrap: false, valign: 'top',
    });
  });
  [0, 1, 2].forEach(i => {
    s.addShape('line', {
      x: o.x + 0.82, y: o.y + 0.668 + i * 0.626, w: 4.35, h: 0,
      line: { color: C.hair, width: 0.5, dashType: i === 2 ? 'solid' : 'sysDot' },
    });
  });

  // X labels 0..7 plus the baseline marker dots.
  for (let i = 0; i <= 7; i++) {
    text(s, String(i), {
      x: o.x + 0.62 + i * 0.622, y: o.y + 2.147, w: 0.243, h: 0.303, fontSize: 12,
      color: C.body, align: 'center', wrap: false, valign: 'top',
    });
    dot(s, o.x + 0.725 + i * 0.624, o.y + 1.91, 0.045, o.color);
  }

  // The curve itself + emphasised data dots along it.
  freeform(s, o.x + 0.83, o.y + 0.661, 4.34, 1.2, o.curve,
    { fill: { type: 'none' }, line: { color: o.color, width: 1.5 } });
  o.dots.forEach(([dx, dy]) => dot(s, o.x + 0.83 + dx * 4.34 - 0.038, o.y + 0.661 + dy * 1.2 - 0.045, 0.078, o.color));
}

function slide07(pptx) {
  const s = pptx.addSlide();

  monthMiniChart(s, {
    x: 1.186, y: 0.802, radius: 0.26, color: C.teal, yAlign: 'center',
    curve: 'M 0 1 L 0.0953 0.2783 C 0.1163 0.1062 0.1557 0 0.1983 0 L 0.2466 0 ' +
      'C 0.269 0 0.291 0.0296 0.3099 0.0849 L 0.3906 0.3214 ' +
      'C 0.4095 0.377 0.4315 0.4063 0.4539 0.4063 L 0.5293 0.4063 ' +
      'C 0.5564 0.4063 0.5827 0.3632 0.6037 0.2843 L 0.6385 0.1543 ' +
      'C 0.6805 0.003 0.7401 0.003 0.7835 0.1403 ' +
      'C 0.8500 0.2600 0.9200 0.1700 1 0.0700',
    dots: [[0.12, 0.06], [0.264, 0.0], [0.408, 0.34], [0.552, 0.34], [0.696, 0.02], [0.84, 0.2], [0.983, 0.0]],
  });

  monthMiniChart(s, {
    x: 1.186, y: 3.892, radius: 0.28, color: C.blue, yAlign: 'left',
    curve: 'M 0 0.7536 L 0.0893 0.9579 C 0.1199 1 0.1559 1 0.1828 0.9102 L 0.2417 0.6934 ' +
      'C 0.2725 0.5801 0.3144 0.5736 0.3467 0.677 ' +
      'C 0.3908 0.819 0.4503 0.7484 0.4762 0.5233 L 0.5103 0.227 ' +
      'C 0.5426 0.005 0.6232 0.005 0.6627 0.1783 L 0.6836 0.3148 ' +
      'C 0.7009 0.4271 0.7279 0.4933 0.7566 0.4933 L 0.8188 0.4933 ' +
      'C 0.8369 0.4933 0.8697 0.5691 1 0.9978',
    dots: [[0.12, 0.98], [0.264, 0.57], [0.408, 0.68], [0.552, 0.0], [0.696, 0.42], [0.84, 0.54]],
  });

  // Right column.
  eyebrow(s, 7.396, 1.066);
  heading(s, 'Report ', 'Analysis', { x: 7.277, y: 1.493 });
  text(s, LOREM_TEMPOR + '.', { x: 7.396, y: 2.468, w: 4.983, h: 0.607, fontSize: 12, color: C.body, valign: 'top' });

  [
    { y: 3.401, label: 'Marketing Data One', color: C.teal, n: '01' },
    { y: 5.069, label: 'Marketing Data Two', color: C.blue, n: '02' },
  ].forEach(b => {
    text(s, b.label, { x: 7.396, y: b.y, w: 3.622, h: 0.337, fontSize: 14, bold: true, color: b.color, valign: 'top' });
    text(s, '85%', { x: 11.208, y: b.y, w: 1.164, h: 0.367, fontSize: 14, bold: true, color: b.color, valign: 'top' });
    progressBar(s, 7.529, b.y + 0.549, 4.356, 0.07, 0.725, b.color);
    text(s, b.n, { x: 7.396, y: b.y + 0.802, w: 1.237, h: 0.423, fontSize: 16, bold: true, color: b.color, lineSpacingMultiple: 1.3, valign: 'top' });
    text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin convallis, ', {
      x: 7.963, y: b.y + 0.803, w: 4.356, h: 0.57, fontSize: 12, color: C.body, lineSpacingMultiple: 1.2, valign: 'top',
    });
  });
}

/* ------------------------------------------------------------------ *
 * Slide 8 — two freeform trend cards with callouts
 * ------------------------------------------------------------------ */

function slide08(pptx) {
  const s = pptx.addSlide();
  centredTitle(s, 'Financial ', 'Growth', 0.993);

  const panels = [
    {
      x: 0.901, radius: 0.22, color: C.teal, title: 'Data Chart One', titleW: 2.039,
      curveY: 3.661, curveH: 0.602,
      labels: ['2019', '2020', '2021', '2022', '2023'], labelAlign: 'left', labelX: 1.177, labelW: 0.838, step: 1.027,
      lineX: 1.308, lineW: 4.684,
      curve: 'M 0 0.7467 C 0.0274 0.5364 0.0488 0.538 0.0842 0.4139 C 0.1195 0.2898 0.174 -0.0292 0.2122 0.0022 C 0.2503 0.0336 0.2834 0.4378 0.3132 0.6023 C 0.343 0.7669 0.3603 1.0067 0.3911 0.9894 C 0.4219 0.9721 0.4665 0.5076 0.4981 0.4984 C 0.5297 0.4891 0.5457 0.8547 0.5809 0.9339 C 0.616 1.0131 0.6715 1.0141 0.7092 0.9737 C 0.7468 0.9332 0.7771 0.8427 0.8069 0.6912 C 0.8368 0.5397 0.8559 0.0959 0.8881 0.0647 C 0.9203 0.0335 0.9682 0.4954 1 0.504',
      knob: { x: 4.915, y: 3.574 }, bubble: { x: 4.631, y: 2.837, label: '+575K' },
      body: 1.171, pillX: 2.476,
    },
    {
      x: 6.928, radius: 0.23, color: C.blue, title: 'Data Chart Two', titleW: 1.883,
      curveY: 3.619, curveH: 0.687,
      labels: MONTHS.slice(0, 5), labelAlign: 'center', labelX: 7.19, labelW: 0.843, step: 1.033,
      lineX: 7.321, lineW: 4.712,
      curve: 'M 0 0.5365 C 0.0274 0.3521 0.0547 0.1678 0.0842 0.2448 C 0.1136 0.3218 0.1411 1.0389 0.1766 0.9983 C 0.212 0.9578 0.2613 0.0432 0.297 0.0017 C 0.3328 -0.0398 0.3597 0.6945 0.3911 0.7492 C 0.4224 0.8039 0.4535 0.338 0.4851 0.3299 C 0.5168 0.3218 0.547 0.6955 0.5809 0.7006 C 0.6147 0.7056 0.6504 0.3957 0.6881 0.3602 C 0.7258 0.3248 0.7709 0.4919 0.8069 0.4879 C 0.843 0.4838 0.8721 0.3633 0.9043 0.3359 C 0.9365 0.3086 0.9682 0.3162 1 0.3238',
      knob: { x: 10.541, y: 3.55 }, bubble: { x: 10.264, y: 2.804, label: '+695%' },
      body: 7.341, pillX: 8.646,
    },
  ];

  panels.forEach(p => {
    card(s, p.x, 2.33, 5.498, 2.838, { radius: p.radius });
    iconBadge(s, glyphArrowRight, p.x + 0.276, 2.63, 0.317, p.color, { scale: 0.42 });
    text(s, p.title, { x: p.x + 0.722, y: 2.598, w: p.titleW, h: 0.37, fontSize: 16, bold: true, color: p.color, valign: 'top' });

    freeform(s, p.x + 0.407, p.curveY, 4.684, p.curveH, p.curve,
      { flipH: true, fill: { type: 'none' }, line: { color: p.color, width: 5 } });
    s.addShape('line', { x: p.lineX, y: 4.521, w: p.lineW, h: 0, line: { color: C.hair, width: 1 } });
    p.labels.forEach((lb, i) => {
      text(s, lb, {
        x: p.labelX + i * p.step, y: 4.582, w: p.labelW, h: 0.3, fontSize: 12,
        color: C.body, align: p.labelAlign, valign: 'top',
      });
    });

    s.addShape('ellipse', {
      x: p.knob.x, y: p.knob.y, w: 0.173, h: 0.173,
      fill: { color: C.white }, line: { color: p.color, width: 3 },
    });
    s.addShape('wedgeRoundRectCallout', {
      x: p.bubble.x, y: p.bubble.y, w: 1.151, h: 0.602, fill: { color: p.color }, line: { type: 'none' },
    });
    text(s, p.bubble.label, {
      x: p.bubble.x, y: p.bubble.y - 0.05, w: 1.151, h: 0.602, align: 'center', valign: 'middle',
      fontSize: 16, bold: true, color: C.white,
    });

    text(s, LOREM_TEMPOR, {
      x: p.body, y: 5.393, w: 4.821, h: 0.607, fontSize: 12, color: C.body,
      align: 'center', lineSpacingMultiple: 1.3, valign: 'top',
    });
    pill(s, p.pillX, 6.155, 2.198, 0.398, 'Your Text', { color: p.color, fontSize: 12 });
  });
}

/** Right-pointing chevron used as the small badge glyph on slide 8. */
function glyphArrowRight(slide, x, y, d, color) {
  freeform(slide, x + d * 0.22, y, d * 0.56, d,
    'M 0.1643 1 L 0 0.9016 L 0.6709 0.4999 L 0 0.0984 L 0.1643 0 L 1 0.4999 L 0.1643 1 Z',
    { fill: { color }, line: { type: 'none' } });
}

/* ------------------------------------------------------------------ *
 * Slide 9 — full-bleed chart card + two stat blocks
 * ------------------------------------------------------------------ */

function slide09(pptx) {
  const s = pptx.addSlide();

  card(s, -0.424, 0.946, 7.288, 5.645, { radius: 0.262 });
  lineChart(s, [
    { name: 'Series 1', color: C.teal, labels: MONTHS, symbol: 'circle', symbolSize: 6, symbolLine: C.white, symbolLineSize: 1.5,
      values: [125, 500, 500, 750, 500, 365, 300, 200, 600, 650, 300, 450] },
    { name: 'Series 2', color: C.blue, labels: MONTHS, symbol: 'circle', symbolSize: 6, symbolLine: C.white, symbolLineSize: 1.5,
      values: [456, 700, 654, 600, 650, 300, 400, 850, 740, 500, 250, 600] },
  ], { x: 0.576, y: 1.137, w: 5.866, h: 5.226 }, {
    layout: PLOT.inset,
    valAxisMaxVal: 1000,
    valAxisMajorUnit: 100,
    valGridLine: { color: C.gridline, size: 0.5 },
    catAxisLabelColor: C.grey,
    valAxisLabelColor: C.grey,
  });

  eyebrow(s, 7.396, 1.252);
  heading(s, 'Report ', 'Analysis', { x: 7.277, y: 1.68 });
  text(s, LOREM_TEMPOR + '.', { x: 7.396, y: 2.654, w: 4.983, h: 0.607, fontSize: 12, color: C.body, valign: 'top' });

  [
    { y: 3.484, color: C.teal, value: '1500+' },
    { y: 5.108, color: C.blue, value: '700+' },
  ].forEach(b => {
    s.addShape('roundRect', {
      x: 7.396, y: b.y, w: 4.76, h: 1.424, fill: { color: b.color }, line: { type: 'none' }, rectRadius: 0.237,
    });
    text(s, b.value, { x: 7.639, y: b.y + 0.28, w: 1.353, h: 0.572, fontSize: 28, bold: true, color: C.white, align: 'center', valign: 'top' });
    text(s, 'Customers', { x: 7.592, y: b.y + 0.83, w: 1.447, h: 0.448, fontSize: 10.5, color: C.white, align: 'center', valign: 'top' });
    text(s, LOREM_SHORT, {
      x: 9.086, y: b.y + 0.379, w: 2.892, h: 0.607, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3, valign: 'top',
    });
  });
}

/* ------------------------------------------------------------------ *
 * Slide 10 — two marker charts + two tagged panels
 * ------------------------------------------------------------------ */

function slide10(pptx) {
  const s = pptx.addSlide();

  const dotValues = [7, 15, 9, 75, 55, 15, 20, 40, 35, 65, 44, 70];
  [
    { y: 0.909, chartY: 1.121, color: C.teal },
    { y: 3.856, chartY: 4.023, color: C.blue },
  ].forEach(p => {
    card(s, 1.177, p.y, 5.247, 2.735, { radius: 0.236 });
    waveChart(s, [{ name: 'Series 2', values: dotValues, color: p.color }],
      { x: 1.462, y: p.chartY, w: 4.962, h: 2.4 },
      {
        xMax: 12, labelSize: 8,
        chart: { lineSmooth: false, lineDataSymbol: 'circle', lineDataSymbolSize: 8, lineDataSymbolLineSize: 1.5 },
      });
  });

  eyebrow(s, 7.396, 1.252);
  heading(s, 'Report ', 'Analysis', { x: 7.277, y: 1.68 });

  [
    { x: 9.062, y: 2.848, tagX: 8.798, tagY: 3.097, color: C.teal },
    { x: 7.541, y: 4.826, tagX: 7.277, tagY: 5.075, color: C.blue },
  ].forEach(p => {
    s.addShape('roundRect', {
      x: p.x, y: p.y, w: 3.095, h: 1.765, fill: { color: p.color }, line: { type: 'none' },
      rectRadius: 0.088, shadow: cardShadow(),
    });
    text(s, 'Your Text Here', { x: p.x + 0.332, y: p.y + 0.658, w: 2.388, h: 0.37, fontSize: 16, bold: true, color: C.white, valign: 'top' });
    text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. In', {
      x: p.x + 0.332, y: p.y + 0.976, w: 2.762, h: 0.607, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3, valign: 'top',
    });
    pill(s, p.tagX, p.tagY, 1.568, 0.342, 'Data One', { color: p.color, fontSize: 14 });
  });
}

/* ------------------------------------------------------------------ *
 * Slide 11 — management statistic wave + data report card
 * ------------------------------------------------------------------ */

function slide11(pptx) {
  const s = pptx.addSlide();
  centredTitle(s, 'Financial ', 'Growth', 0.993);
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. ', {
    x: 2.179, y: 2.222, w: 9.071, h: 0.607, fontSize: 12, color: C.body, align: 'center', lineSpacingMultiple: 1.3, valign: 'top',
  });

  // Big left card with a single flowing curve.
  card(s, 1.14, 3.282, 7.426, 3.225, { radius: 0.108 });
  freeform(s, 1.608, 4.277, 6.546, 1.269,
    'M 0 0.88 C 0.1586 0.88 0.1241 0 0.2667 0 C 0.4092 0 0.3701 1 0.5724 1 C 0.7103 1 0.692 0.3 0.7885 0.3 C 0.8345 0.3 0.8851 0.5 1 0.5',
    { fill: { type: 'none' }, line: { color: C.teal, width: 3 } });
  text(s, [
    { text: 'Management' },
    { text: 'Statistic', options: { softBreakBefore: true } },
  ], { x: 1.473, y: 3.568, w: 1.909, h: 0.775, fontSize: 16, bold: true, color: C.grey, valign: 'top' });

  // Dotted marker on the curve + "This Week 25K sales".
  s.addShape('line', {
    x: 4.881, y: 3.899, w: 0, h: 1.394,
    line: { color: C.teal, width: 1, dashType: 'dash', beginArrowType: 'oval' },
  });
  s.addShape('ellipse', {
    x: 4.809, y: 5.414, w: 0.143, h: 0.121, fill: { color: C.teal }, line: { color: C.white, width: 2.25 },
  });
  text(s, 'This Week', { x: 5.146, y: 3.622, w: 1.684, h: 0.425, fontSize: 16, color: C.grey, valign: 'top' });
  text(s, '25K', { x: 5.21, y: 3.823, w: 0.847, h: 0.587, fontSize: 24, bold: true, color: C.darkTeal, valign: 'top' });
  text(s, 'sales', { x: 5.948, y: 3.914, w: 0.852, h: 0.425, fontSize: 16, color: C.darkTeal, valign: 'top' });

  // Right card: data report + marketing bar + two note columns.
  card(s, 8.831, 4.439, 3.362, 2.057, { radius: 0.108 });
  dot(s, 8.895, 3.306, 0.577, C.white, cardShadow());
  // Down-arrow glyph rotated 270° => points right.
  freeform(s, 9.095, 3.487, 0.176, 0.214,
    'M 0.0238 0.6353 C -0.0079 0.6092 -0.0079 0.5671 0.0238 0.541 C 0.0397 0.528 0.0603 0.5214 0.0811 0.5214 C 0.102 0.5214 0.1226 0.528 0.1385 0.541 L 0.4199 0.7724 L 0.4199 0.0667 C 0.4199 0.03 0.4561 0 0.501 0 C 0.5457 0 0.5821 0.0298 0.5821 0.0667 L 0.5821 0.7724 L 0.863 0.541 C 0.8948 0.5149 0.9459 0.5149 0.9777 0.541 C 1.0094 0.5671 1.0094 0.6092 0.9777 0.6353 L 0.5581 0.9804 C 0.5429 0.9929 0.5223 1 0.5007 1 C 0.4792 1 0.4586 0.9931 0.4434 0.9804 L 0.0238 0.6353 Z',
    { rotate: 270, fill: { color: C.teal }, line: { type: 'none' } });
  text(s, 'Data Report Financial', { x: 9.653, y: 3.271, w: 2.92, h: 0.37, fontSize: 16, bold: true, color: C.teal, valign: 'top' });
  text(s, LOREM_SHORT, { x: 9.653, y: 3.705, w: 2.92, h: 0.607, fontSize: 12, color: C.body, lineSpacingMultiple: 1.3, valign: 'top' });

  text(s, 'Marketing Data One', { x: 9.021, y: 4.681, w: 2.09, h: 0.337, fontSize: 14, bold: true, color: C.teal, valign: 'top' });
  text(s, '85%', { x: 11.314, y: 4.681, w: 0.672, h: 0.367, fontSize: 14, bold: true, color: C.teal, align: 'right', valign: 'top' });
  progressBar(s, 9.098, 5.207, 2.824, 0.068, 0.645, C.teal);
  [9.021, 10.761].forEach(x => {
    text(s, 'Lorem ipsum dolor sit amet, consectetur', {
      x, y: 5.457, w: 1.433, h: 0.814, fontSize: 12, color: C.body, lineSpacingMultiple: 1.2, valign: 'top',
    });
  });
}

/* ------------------------------------------------------------------ *
 * Slide 12 — 2024 / 2025 data report charts + two feature cards
 * ------------------------------------------------------------------ */

function slide12(pptx) {
  const s = pptx.addSlide();

  const seriesA = [0, 5000, 5200, 10000, 1300, 12000, 9500, 9500, 10600, 6500, 3000, 15000];
  const seriesB = [0, 4000, 4000, 7900, 13000, 10000, 11000, 8500, 14000, 5000, 12600, 13700];

  [
    { x: 0.97, radius: 0.224, title: '2024 Data Report', color: C.teal, lineColor: C.teal },
    { x: 6.818, radius: 0.194, title: '2025 Data Report', color: C.blue, lineColor: C.blue },
  ].forEach(p => {
    card(s, p.x, 0.778, 5.591, 3.722, { radius: p.radius });
    text(s, p.title, { x: p.x + 0.268, y: 1.039, w: 5.139, h: 0.438, fontSize: 20, bold: true, color: p.color, align: 'center', valign: 'top' });
    lineChart(s, [
      { name: 'Series 1', color: p.lineColor, labels: MONTHS, symbol: 'circle', symbolSize: 5, symbolLine: p.lineColor, values: seriesA },
      { name: 'Series 2', color: p.lineColor, labels: MONTHS, symbol: 'diamond', symbolSize: 5, symbolLine: p.lineColor, values: seriesB },
    ], { x: p.x + 0.243, y: 1.493, w: 5.164, h: 2.847 }, {
      layout: PLOT.inset,
      valAxisMaxVal: 20000,
      valAxisMajorUnit: 5000,
      valGridLine: { color: 'F0F0F0', size: 0.5 },
      catAxisLabelFontSize: 10,
      valAxisLabelFontSize: 10,
      catAxisLabelColor: C.grey,
      valAxisLabelColor: C.grey,
    });
  });

  [
    { x: 0.97, iconX: 1.423, title: 'Special Taste 01', color: C.teal, glyph: glyphBulb },
    { x: 6.818, iconX: 7.272, title: 'Special Taste 02', color: C.blue, glyph: glyphTrend },
  ].forEach(p => {
    card(s, p.x, 4.729, 5.591, 1.993, { radius: 0.222 });
    iconBadge(s, p.glyph, p.iconX, 5.241, 0.997, p.color, { scale: 0.59 });
    text(s, p.title, { x: p.x + 1.723, y: 5.187, w: 2.949, h: 0.438, fontSize: 20, bold: true, color: p.color, valign: 'top' });
    text(s, LOREM_VEHICULA, {
      x: p.x + 1.722, y: 5.685, w: 3.654, h: 0.608, fontSize: 12, color: C.body, lineSpacingMultiple: 1.3, valign: 'top',
    });
  });
}

/* ------------------------------------------------------------------ *
 * Slide 13 — Selling / Buying quarterly chart + two data cards
 * ------------------------------------------------------------------ */

function slide13(pptx) {
  const s = pptx.addSlide();
  centredTitle(s, 'Financial ', 'Growth', 1.116);

  card(s, 1.303, 2.449, 6.576, 4.088, { radius: 0.27 });
  lineChart(s, stackSeries([
    { name: 'Selling', color: C.teal, labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [4.3, 2.5, 3.5, 3],
      symbol: 'circle', symbolSize: 9, symbolLine: C.white, symbolLineSize: 1, lineSize: 2.75 },
    { name: 'Buying', color: C.blue, labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [2.4, 4.4, 1.8, 2.8],
      symbol: 'circle', symbolSize: 9, symbolLine: C.white, symbolLineSize: 1.25, lineSize: 2.5 },
  ]), { x: 1.665, y: 2.824, w: 5.928, h: 3.56 }, {
    valAxisMaxVal: 8,
    valAxisMajorUnit: 1,
    showLegend: true,
    legendPos: 'b',
    layout: { x: 0.07, y: 0.02, w: 0.9, h: 0.79 },
    legendFontFace: FONT,
    legendFontSize: 12,
    legendColor: C.grey,
    catAxisLabelColor: C.grey,
    valAxisLabelColor: C.grey,
    catAxisLineShow: true,
    catAxisLineColor: C.gridline,
    catAxisLineSize: 0.75,
    valGridLine: { color: C.gridline, size: 0.75 },
  });

  [
    { y: 2.454, value: '270K', label: 'Data 01', color: C.teal, glyph: glyphBulb },
    { y: 4.661, value: '300K', label: 'Data 02', color: C.blue, glyph: glyphTrend },
  ].forEach(p => {
    card(s, 8.159, p.y, 3.871, 1.876, { radius: 0.248 });
    iconTextRow(s, {
      glyph: p.glyph, x: 8.487, y: p.y + 0.534, d: 0.618, color: p.color, shadow: haloShadow(),
      value: p.value, tx: 9.293, vy: p.y + 0.397,
      label: p.label, lx: 10.281, ly: p.y + 0.449, labelColor: C.ink, labelBold: false,
      body: LOREM_SHORT, by: p.y + 0.871, bw: 3.046,
    });
  });
}

/* ------------------------------------------------------------------ *
 * Slide 14 — Report Analysis + Data Sales area chart card
 * ------------------------------------------------------------------ */

function slide14(pptx) {
  const s = pptx.addSlide();

  eyebrow(s, 1.305, 1.252);
  heading(s, 'Report ', 'Analysis', { x: 1.186, y: 1.68 });

  pill(s, 1.305, 2.851, 1.413, 0.458, '2024', { fill: C.teal, color: C.white, fontSize: 18, radius: 0.08 });
  text(s, LOREM_VEHICULA, { x: 2.9, y: 2.747, w: 3.688, h: 0.604, fontSize: 12, color: C.ink, lineSpacingMultiple: 1.3, valign: 'top' });

  [
    { y: 3.776, iconY: 3.896, color: C.teal, glyph: glyphBulb },
    { y: 5.312, iconY: 5.433, color: C.blue, glyph: glyphTrend },
  ].forEach(p => {
    iconBadge(s, p.glyph, 1.597, p.iconY, 0.869, p.color, { shadow: haloShadow(), scale: 0.54 });
    text(s, '270K \u2013 380K', { x: 2.915, y: p.y, w: 2.179, h: 0.438, fontSize: 20, bold: true, color: p.color, valign: 'top' });
    text(s, LOREM_VEHICULA, { x: 2.915, y: p.y + 0.437, w: 4.244, h: 0.603, fontSize: 12, color: C.ink, lineSpacingMultiple: 1.3, valign: 'top' });
  });

  // Right: header strip + big card with a filled area silhouette.
  card(s, 7.187, 0.966, 4.96, 0.559, { radius: 0.075 });
  text(s, 'DATA SALES', { x: 8.436, y: 1.048, w: 2.461, h: 0.404, fontSize: 18, bold: true, color: C.teal, align: 'center', valign: 'top' });

  card(s, 7.187, 1.669, 4.96, 4.865, { radius: 0.119 });
  text(s, '500.67+', { x: 8.559, y: 1.944, w: 2.297, h: 0.572, fontSize: 28, bold: true, color: C.teal, align: 'center', charSpacing: 3, valign: 'top' });
  text(s, 'Total Patient', { x: 8.518, y: 2.427, w: 2.38, h: 0.386, fontSize: 14, color: C.ink, align: 'center', lineSpacingMultiple: 1.3, valign: 'top' });
  s.addShape('line', { x: 7.513, y: 3.075, w: 4.307, h: 0, line: { color: 'A6A6A6', transparency: 50, width: 1, dashType: 'dash' } });

  freeform(s, 7.543, 3.904, 4.307, 1.579,
    'M 0.4515 0 L 0.5068 0 L 0.5798 0.2974 L 0.6903 0.2974 L 0.7732 0.4821 L 0.8205 0.4821 L 0.9271 0.3582 L 1 0.3582 L 1 1 L 0 1 L 0 0.5759 L 0.0055 0.5745 L 0.0489 0.5745 L 0.1081 0.6804 L 0.2206 0.4303 L 0.2739 0.4303 L 0.3252 0.1374 L 0.3982 0.1374 Z',
    { fill: { color: C.teal }, line: { color: '025669', width: 1.25 } });
  ['20', '30', '40', '50', '60', '70', '80'].forEach((lb, i) => {
    text(s, lb, {
      x: 7.648 + i * 0.5855, y: 5.503, w: 0.585, h: 0.305, fontSize: 10, color: C.grey,
      align: 'center', lineSpacingMultiple: 1.3, valign: 'top',
    });
  });
  pill(s, 7.574, 5.928, 2.036, 0.385, 'Selling', { fill: C.teal, color: C.white, fontSize: 14, radius: 0.044, shadow: null });
  pill(s, 9.783, 5.928, 2.036, 0.385, 'Buying', { fill: C.blue, color: C.white, fontSize: 14, radius: 0.044, shadow: null });
}

/* ------------------------------------------------------------------ *
 * Slide 15 — Financial Analysis, tall chart card + two stat cards
 * ------------------------------------------------------------------ */

function slide15(pptx) {
  const s = pptx.addSlide();

  card(s, 1.003, 1.115, 5.664, 6.711, { radius: 0.446 });
  text(s, 'Financial Data Report', { x: 1.919, y: 1.409, w: 3.823, h: 0.404, fontSize: 18, bold: true, color: C.teal, align: 'center', valign: 'top' });
  lineChart(s, [
    { name: 'Series 1', color: C.blue, labels: MONTHS, symbol: 'circle', symbolSize: 5, symbolLine: C.blue,
      values: [125, 500, 123, 754, 654, 365, 300, 201, 874, 650, 300, 283] },
    { name: 'Series 2', color: C.teal, labels: MONTHS, symbol: 'diamond', symbolSize: 5, symbolLine: C.teal,
      values: [456, 700, 654, 290, 400, 110, 233, 739, 740, 415, 840, 543] },
  ], { x: 1.274, y: 1.982, w: 5.11, h: 5.05 }, {
    layout: PLOT.inset,
    valAxisMaxVal: 1000,
    valAxisMajorUnit: 100,
    valGridLine: { color: 'F0F0F0', size: 0.5 },
    catAxisLabelFontSize: 10,
    valAxisLabelFontSize: 10,
    catAxisLabelColor: C.grey,
    valAxisLabelColor: C.grey,
  });

  eyebrow(s, 7.057, 1.252);
  heading(s, 'Financial ', 'Analysis', { x: 6.938, y: 1.68, w: 5.941 });

  [
    { cx: 7.057, cy: 2.825, ix: 7.487, iy: 3.236, tx: 8.255, ty: 3.298, lx: 7.414, ly: 4.091,
      pct: '85%', title: 'Data One', color: C.teal, glyph: glyphTrend, titleW: 1.305 },
    { cx: 9.895, cy: 3.767, ix: 10.325, iy: 4.178, tx: 11.093, ty: 4.24, lx: 10.252, ly: 5.032,
      pct: '50%', title: 'Data Two', color: C.blue, glyph: glyphTimer, titleW: 1.317 },
  ].forEach(p => {
    card(s, p.cx, p.cy, 2.678, 2.675, { radius: 0.238 });
    iconBadge(s, p.glyph, p.ix, p.iy, 0.634, p.color, { scale: 0.61 });
    text(s, p.pct, { x: p.tx, y: p.ty, w: 1.148, h: 0.572, fontSize: 28, bold: true, color: p.color, valign: 'top' });
    text(s, p.title, { x: p.lx, y: p.ly, w: p.titleW, h: 0.404, fontSize: 18, bold: true, color: p.color, wrap: false, valign: 'top' });
    text(s, 'Lorem ipsum dolor sit amet, consectetur', {
      x: p.lx, y: p.ly + 0.407, w: 2.321, h: 0.607, fontSize: 12, color: C.body, lineSpacingMultiple: 1.3, valign: 'top',
    });
  });
}

/* ------------------------------------------------------------------ *
 * Tiny SVG-ish path parser -> pptxgenjs `points` (unit coordinates)
 * ------------------------------------------------------------------ */

function parsePath(d, w, h) {
  const tok = d.trim().split(/\s+/);
  const px = () => +tok[i++] * w;
  const py = () => +tok[i++] * h;
  const pts = [];
  let i = 0;
  while (i < tok.length) {
    const op = tok[i++];
    if (op === 'M') pts.push({ x: px(), y: py(), moveTo: true });
    else if (op === 'L') pts.push({ x: px(), y: py() });
    else if (op === 'C') {
      const x1 = px(), y1 = py(), x2 = px(), y2 = py();
      pts.push({ x: px(), y: py(), curve: { type: 'cubic', x1, y1, x2, y2 } });
    } else if (op === 'Z') pts.push({ close: true });
  }
  return pts;
}

/* ------------------------------------------------------------------ *
 * Build & save
 * ------------------------------------------------------------------ */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE_16x9', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE_16x9';
  pptx.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pptx.title = 'Line Chart Infographic';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
    slide09, slide10, slide11, slide12, slide13, slide14, slide15].forEach(fn => fn(pptx));

  return pptx;
}

build()
  .writeFile({ fileName: path.join(__dirname, '03b335e2-8ed3-4c1b-9bfb-844eecff81f7_grok_final.pptx') })
  .then(f => console.log('wrote ' + f))
  .catch(err => { console.error(err); process.exit(1); });
