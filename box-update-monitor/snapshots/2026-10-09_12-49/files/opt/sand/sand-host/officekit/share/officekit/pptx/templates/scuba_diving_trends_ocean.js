/*
 * Matias — Diving Presentation Template (40 slides, 13.333in x 7.5in)
 * Recreated with pptxgenjs only.  Run:  node <this file>
 *
 * Photographs in the original are replaced with flat placeholder blocks
 * tinted to each image's average colour (see `photo`).
 */
'use strict';

const { join } = require('path');
const pptxgen = require('pptxgenjs');

// ---------------------------------------------------------------- palette
const BLUE = '01439A';     // brand blue
const DEEP = '10468D';     // mid gradient blue
const SKY = '2986FE';      // bright gradient blue
const SKY2 = '70ACFE';     // halo blue
const BLUE2 = '0052BD';
const BLUE3 = '003273';
const GOLD = 'FFC000';
const NAVY = '222A35';     // dark panel
const WHITE = 'FFFFFF';
const GRAY = '7F7F7F';     // body copy on light panels
const SILVER = 'D8D8D8';   // body copy on dark panels
const MIST = 'F2F2F2';
const ASH = 'BFBFBF';
const PHOTO = 'CCCCCC';    // placeholder tint for light photos
const PHOTO2 = '989898';   // placeholder tint for darker photos

// ---------------------------------------------------------------- typography
const FONTS = { '': 'Poppins', MED: 'Poppins Medium', SEMI: 'Poppins SemiBold' };
const INSET = [7.2, 7.2, 3.6, 3.6];  // l, r, b, t in points — matches the source deck
const SHADOW = { type: 'outer', color: '000000', opacity: 0.24, blur: 23, offset: 12, angle: 90 };

// Filler copy used throughout the template.
const LOREM =
  'Re paragone creatura acerbita ai guardava lasciami vi. Entro tue forza miele mazzo per pur oltre sul. ' +
  'Lo ti il gabbie quanto lancio. Fato mare arme tu anch vi mine riso. Poi affannata ami cresciuto melagrani ' +
  'una abbandona brillanti. Far aspettando nel voluttuosa sei turbamento tra. Grappoli tuo inquieta cio orribile ' +
  'dissolve scoperto. Alzeremo voi parlando pei qualcuno serbatoi mio bellezza. Impregnato voi san esaltavano ' +
  'dal dal sfaldavano. ';
const DEV =
  'Dev certa volle sopra anime animo qua. Tenta anima la no aveva ch rombo le. Se da distrutta liberarli infantile ' +
  'usignuoli. Ora afa rimorso dai braccia sentito superbe chi. Indicibili ho esaltavano raccontava un di fu ' +
  'impregnato. ho';
const cut = (s, len) => s.slice(0, len);

// ---------------------------------------------------------------- drawing helpers
/** Common shape properties shared by `sh`, `tx` and `photo`. */
function shapeOpts(x, y, w, h, o) {
  const p = { x, y, w, h };
  if (o.fill === null) p.fill = { type: 'none' };
  else if (o.fill) p.fill = o.alpha ? { color: o.fill, transparency: o.alpha } : { color: o.fill };
  if (o.line) p.line = { color: o.line, width: o.lw || 1 };
  else p.line = { type: 'none' };
  if (o.rad !== undefined) p.rectRadius = o.rad;
  if (o.arc) p.angleRange = o.arc;
  if (o.rot) p.rotate = o.rot;
  if (o.flipH) p.flipH = true;
  if (o.flipV) p.flipV = true;
  if (o.shadow) p.shadow = SHADOW;
  return p;
}

/** Plain shape. */
function sh(s, kind, x, y, w, h, o = {}) {
  s.addShape(kind, shapeOpts(x, y, w, h, o));
}

/** Stand-in block for a photograph from the original deck. */
function photo(s, x, y, w, h, kind = 'rect', o = {}) {
  s.addShape(kind, shapeOpts(x, y, w, h, Object.assign({ fill: PHOTO }, o)));
}

/**
 * Text box.  `body` is either a plain string or a list of runs:
 *   [text, color, { sz, bold, font, sup, br }]
 */
function tx(s, x, y, w, h, body, o = {}) {
  const opts = Object.assign(shapeOpts(x, y, w, h, o), {
    shape: o.shape || 'rect',
    fontFace: FONTS[o.font || ''],
    fontSize: o.sz || 12,
    bold: !!o.bold,
    color: o.color || WHITE,
    align: o.align || 'left',
    valign: o.valign || 'top',
    margin: INSET,
    wrap: true,
    isTextBox: true,
  });
  if (o.lh) opts.lineSpacingMultiple = o.lh;
  if (typeof body === 'string') { s.addText(body, opts); return; }
  const runs = body.map(([text, color, r = {}]) => ({
    text,
    options: {
      color: color || opts.color,
      fontSize: r.sz || opts.fontSize,
      fontFace: r.font !== undefined ? FONTS[r.font] : opts.fontFace,
      bold: r.bold !== undefined ? r.bold : opts.bold,
      superscript: !!r.sup,
      breakLine: !!r.br,
    },
  }));
  s.addText(runs, opts);
}

/**
 * Free-form shape.  `pts` holds fractions of the bounding box: a pair is a
 * line/move target, a sextet is a cubic bezier (c1, c2, end).
 */
function path(s, x, y, w, h, pts, o = {}) {
  const points = pts.map((p, i) => {
    if (p.length === 2) return { x: p[0] * w, y: p[1] * h, moveTo: i === 0 };
    return { x: p[4] * w, y: p[5] * h,
             curve: { type: 'cubic', x1: p[0] * w, y1: p[1] * h, x2: p[2] * w, y2: p[3] * h } };
  });
  points.push({ close: true });
  s.addShape('custGeom', Object.assign(shapeOpts(x, y, w, h, o), { points }));
}

// ---------------------------------------------------------------- gradients
// pptxgenjs cannot emit <a:gradFill>, so the brand gradient is painted as a
// fan of flat polygon bands clipped out of the shape's outline.
const BRAND = [DEEP, NAVY];   // the deck's diagonal blue wash
const BANDS = 24;

/** Linear blend between two hex colours. */
function ramp(t, stops) {
  let out = '';
  for (let k = 0; k < 3; k++) {
    const a = parseInt(stops[0].substr(k * 2, 2), 16), b = parseInt(stops[1].substr(k * 2, 2), 16);
    out += Math.round(a + (b - a) * t).toString(16).padStart(2, '0');
  }
  return out.toUpperCase();
}

const rectPts = (x, y, w, h) => [[x, y], [x + w, y], [x + w, y + h], [x, y + h]];

function ovalPts(x, y, w, h, seg = 48) {
  const out = [];
  for (let i = 0; i < seg; i++) {
    const a = (i / seg) * 2 * Math.PI;
    out.push([x + w / 2 * (1 + Math.cos(a)), y + h / 2 * (1 + Math.sin(a))]);
  }
  return out;
}

function roundPts(x, y, w, h, r, seg = 8) {
  const out = [];
  const corners = [[x + w - r, y + h - r, 0], [x + r, y + h - r, 90], [x + r, y + r, 180], [x + w - r, y + r, 270]];
  for (const [cx, cy, a0] of corners) {
    for (let i = 0; i <= seg; i++) {
      const a = (a0 + 90 * i / seg) * Math.PI / 180;
      out.push([cx + r * Math.cos(a), cy + r * Math.sin(a)]);
    }
  }
  return out;
}

/** Keep the part of `poly` on the side of the half-plane where n.p <= d. */
function clip(poly, nx, ny, d) {
  const out = [];
  for (let i = 0; i < poly.length; i++) {
    const a = poly[i], b = poly[(i + 1) % poly.length];
    const da = nx * a[0] + ny * a[1] - d, db = nx * b[0] + ny * b[1] - d;
    if (da <= 0) out.push(a);
    if ((da < 0 && db > 0) || (da > 0 && db < 0)) {
      const f = da / (da - db);
      out.push([a[0] + (b[0] - a[0]) * f, a[1] + (b[1] - a[1]) * f]);
    }
  }
  return out;
}

/** Paint `poly` with the brand gradient running along `angle` degrees. */
function grad(s, poly, o = {}) {
  const stops = o.stops || BRAND;
  const angle = (o.angle === undefined ? 150 : o.angle) * Math.PI / 180;
  const dx = Math.cos(angle), dy = Math.sin(angle);
  const proj = poly.map(p => p[0] * dx + p[1] * dy);
  const lo = Math.min(...proj), span = Math.max(...proj) - lo;
  const step = span / BANDS, bleed = step * 0.08;
  for (let i = 0; i < BANDS; i++) {
    let band = clip(poly, -dx, -dy, -(lo + step * i - bleed));
    band = clip(band, dx, dy, lo + step * (i + 1) + bleed);
    if (band.length < 3) continue;
    const xs = band.map(p => p[0]), ys = band.map(p => p[1]);
    const x0 = Math.min(...xs), y0 = Math.min(...ys);
    const w = Math.max(...xs) - x0, h = Math.max(...ys) - y0;
    if (w < 0.001 || h < 0.001) continue;
    s.addShape('custGeom', {
      x: x0, y: y0, w, h,
      fill: { color: ramp((i + 0.5) / BANDS, stops) },
      line: { type: 'none' },
      points: band.map((p, k) => ({ x: p[0] - x0, y: p[1] - y0, moveTo: k === 0 })).concat([{ close: true }]),
    });
  }
}

// ---------------------------------------------------------------- charts
const CHART_BASE = {
  showLegend: false, showTitle: false,
  chartColors: [BLUE, NAVY, GOLD],
  catAxisLabelFontFace: 'Poppins', catAxisLabelFontSize: 9, catAxisLabelColor: GRAY,
  valAxisLabelFontFace: 'Poppins', valAxisLabelFontSize: 9, valAxisLabelColor: GRAY,
  catAxisLineShow: false, valAxisLineShow: false,
  valGridLine: { style: 'solid', color: 'D9D9D9', size: 0.5 },
  catGridLine: { style: 'none' },
  barGapWidthPct: 20,
};

/** Slide 36 — alternating gold/blue histogram. */
function histogramChart(s, x, y, w, h) {
  const bins = ['[1, 5]', '(5, 9]', '(9, 13]', '(13, 17]', '(17, 21]', '(21, 25]'];
  const counts = [5, 11, 23, 24, 9, 4];
  s.addChart('bar', [{ name: 'Frequency', labels: bins, values: counts }], Object.assign({}, CHART_BASE, {
    x, y, w, h, barDir: 'col', barGapWidthPct: 8, valAxisMaxVal: 25, valAxisMajorUnit: 5,
    showTitle: true, title: 'Chart Title', titleFontFace: 'Poppins', titleFontSize: 14, titleColor: GRAY,
    chartColors: [GOLD, BLUE, GOLD, BLUE, GOLD, BLUE], varyColors: true,
  }));
}

/** Slide 37 — three-series stacked column. */
function stackedChart(s, x, y, w, h) {
  const cats = ['Category 1', 'Category 2', 'Category 3', 'Category 4'];
  s.addChart('bar', [
    { name: 'Series 1', labels: cats, values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels: cats, values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels: cats, values: [2, 2, 3, 5] },
  ], Object.assign({}, CHART_BASE, {
    x, y, w, h, barDir: 'col', barGrouping: 'stacked', valAxisMaxVal: 14, valAxisMajorUnit: 2,
    showLegend: true, legendPos: 'b', legendFontFace: 'Poppins', legendFontSize: 9, legendColor: GRAY,
  }));
}

/** Slide 38 — increase / decrease / total comparison. */
function waterfallChart(s, x, y, w, h) {
  const cats = ['Category 1', 'Category 2', 'Category 3', 'Category 4',
                'Category 5', 'Category 6', 'Category 7', 'Category 8'];
  s.addChart('bar', [
    { name: 'Increase', labels: cats, values: [100, null, 50, null, 130, null, 70, null] },
    { name: 'Decrease', labels: cats, values: [null, 20, null, 40, null, 60, null, 140] },
  ], Object.assign({}, CHART_BASE, {
    x, y, w, h, barDir: 'col', valAxisMaxVal: 180, valAxisMajorUnit: 20,
    showTitle: true, title: 'Chart Title', titleFontFace: 'Poppins', titleFontSize: 14, titleColor: GRAY,
    showLegend: true, legendPos: 't', legendFontFace: 'Poppins', legendFontSize: 9, legendColor: GRAY,
    showValue: true, dataLabelFontFace: 'Poppins', dataLabelFontSize: 8, dataLabelColor: GRAY,
    chartColors: [BLUE, GOLD],
  }));
}

// ---------------------------------------------------------------- 1. Matias
function slide01(s) {
  grad(s, rectPts(0, 0, 13.333, 3.114));

  tx(s, 6.574, 4.331, 4.157, 1.447, 'Matias', { sz: 80, bold: true, align: 'center', color: BLUE });
  tx(s, 7.026, 5.61, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', align: 'center', color: GOLD });
  sh(s, 'roundRect', 0.535, 0.45, 0.312, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 0.603, 0.544, 0.244, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 0.603, 0.637, 0.244, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  tx(s, 0.847, 0.489, 1.426, 0.269, 'Trends', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 2.266, 0.489, 1.426, 0.269, 'Diving', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 3.608, 0.489, 1.426, 0.269, 'Ocean', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 4.91, 0.489, 1.426, 0.269, 'Contact', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 11.302, 4.791, 3.035, 0.337, [['Best Day ', GOLD], ['for Diving in 2021', BLUE]], { sz: 14, bold: true, rot: 90 });
  tx(s, 0.284, 6.75, 1.982, 0.572, [['Published by: ', BLUE], ['yumnacreative', GOLD]], { sz: 14, bold: true });
  photo(s, 1.076, 0.957, 11.181, 3.214, 'roundRect', { rad: 0.239, fill: 'CBCBCB' });
}

// ---------------------------------------------------------------- 2. Introduction
function slide02(s) {
  sh(s, 'rect', 0, 0, 13.333, 7.5, { fill: NAVY });

  photo(s, 1.076, 0.864, 11.181, 5.771, 'roundRect', { rad: 0.43, fill: '979797' });
  tx(s, 0.442, 2.662, 4.799, 1.587, DEV, { align: 'justify', lh: 1.5, color: MIST });
  tx(s, 0.442, 1.215, 7.79, 1.447, 'Introduction', { sz: 80, bold: true, color: GOLD });
  sh(s, 'roundRect', 0.535, 0.45, 0.312, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 0.603, 0.544, 0.244, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 0.603, 0.637, 0.244, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  tx(s, 11.302, 2.364, 3.035, 0.337, 'Best Day for Diving in 2021', { sz: 14, bold: true, rot: 90, color: GOLD });
  tx(s, 0.284, 6.949, 3.408, 0.337, 'Published by: yumnacreative', { sz: 14, bold: true, color: GOLD });
  tx(s, 0.44, 4.355, 4.799, 0.981, cut(DEV, 142), { align: 'justify', lh: 1.5, color: MIST });
  tx(s, 9.449, 6.417, 1.78, 0.375, 'Watch More', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'roundRect', rad: 0.076, shadow: true, fill: BLUE, color: WHITE });
  tx(s, 0.847, 0.489, 1.426, 0.269, 'Trends', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 2.266, 0.489, 1.426, 0.269, 'Diving', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 3.608, 0.489, 1.426, 0.269, 'Ocean', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 4.91, 0.489, 1.426, 0.269, 'Contact', { sz: 10, font: 'MED', align: 'center', color: WHITE });
}

// ---------------------------------------------------------------- 3. Together We Can Save Our Ocean
function slide03(s) {
  grad(s, rectPts(4.1, 0, 9.233, 7.5));

  tx(s, 6.154, 2.055, 5.822, 2.193, LOREM, { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 6.154, 1.8, 2.361, 0.303, 'INSERT TEXT HERE', { font: 'SEMI', color: GOLD });
  tx(s, 6.154, 0.757, 4.655, 1.043, 'Together We Can Save Our Ocean', { sz: 28, bold: true, color: WHITE });
  tx(s, 6.154, 4.291, 5.822, 0.981, cut(LOREM, 173), { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 11.302, 5.814, 3.035, 0.337, 'Best Day for Diving in 2021', { sz: 14, bold: true, rot: 90, color: GOLD });
  tx(s, 6.606, 5.526, 1.909, 0.303, 'Insert Text Here', { font: 'SEMI', color: WHITE });
  sh(s, 'ellipse', 6.185, 5.455, 0.444, 0.444, { fill: SKY2, alpha: 60 });
  sh(s, 'ellipse', 6.233, 5.503, 0.347, 0.347, { fill: BLUE });
  sh(s, 'mathPlus', 6.301, 5.563, 0.212, 0.228, { fill: WHITE });
  tx(s, 8.936, 5.526, 1.909, 0.303, 'Insert Text Here', { font: 'SEMI', color: WHITE });
  sh(s, 'ellipse', 8.515, 5.455, 0.444, 0.444, { fill: SKY2, alpha: 60 });
  sh(s, 'ellipse', 8.563, 5.503, 0.347, 0.347, { fill: BLUE });
  sh(s, 'mathPlus', 8.631, 5.563, 0.212, 0.228, { fill: WHITE });
  tx(s, 6.606, 6.184, 1.909, 0.303, 'Insert Text Here', { font: 'SEMI', color: WHITE });
  sh(s, 'ellipse', 6.185, 6.114, 0.444, 0.444, { fill: SKY2, alpha: 60 });
  sh(s, 'ellipse', 6.233, 6.162, 0.347, 0.347, { fill: BLUE });
  sh(s, 'mathPlus', 6.301, 6.222, 0.212, 0.228, { fill: WHITE });
  tx(s, 8.936, 6.184, 1.909, 0.303, 'Insert Text Here', { font: 'SEMI', color: WHITE });
  sh(s, 'ellipse', 8.515, 6.114, 0.444, 0.444, { fill: SKY2, alpha: 60 });
  sh(s, 'ellipse', 8.563, 6.162, 0.347, 0.347, { fill: BLUE });
  sh(s, 'mathPlus', 8.631, 6.222, 0.212, 0.228, { fill: WHITE });
  photo(s, 0, 0.007, 3.886, 7.493, 'rect', { fill: 'CBCBCB' });
  photo(s, 2.471, 2.257, 2.986, 2.986, 'ellipse', { flipH: true, fill: '979797', line: WHITE, lw: 6 });
}

// ---------------------------------------------------------------- 4. Trending Diving Spot
function slide04(s) {
  sh(s, 'rect', 6.667, 0, 6.667, 7.5, { fill: NAVY });
  sh(s, 'rect', 7.179, 0.457, 1.357, 2.114, { fill: GOLD });

  tx(s, 0.453, 2.669, 5.822, 2.193, LOREM, { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 0.453, 2.414, 2.361, 0.303, 'INSERT TEXT HERE', { font: 'SEMI', color: SKY });
  tx(s, 0.453, 1.714, 4.655, 0.572, 'Trending Diving Spot', { sz: 28, bold: true, color: BLUE });
  sh(s, 'roundRect', 4.342, 5.074, 2.915, 1.721, { fill: NAVY, shadow: true, rad: 0.287 });
  grad(s, roundPts(4.342, 5.074, 2.915, 1.721, 0.287));
  tx(s, 4.764, 5.796, 2.542, 0.692, 'Amazing Trending Diving Spot 2021 ', { sz: 14, bold: true, font: 'SEMI', lh: 1.3, color: WHITE });
  tx(s, 4.764, 5.236, 1.313, 0.558, [['14', GOLD], ['th', GOLD, { sup: true }], [' Jan', GOLD]], { sz: 20, bold: true, font: 'SEMI', lh: 1.5 });
  tx(s, 0.453, 4.965, 3.418, 1.587, cut(LOREM, 173), { align: 'justify', lh: 1.5, color: GRAY });
  sh(s, 'roundRect', 0.535, 0.45, 0.312, 0.051, { rad: 0.009, shadow: true, fill: GOLD });
  sh(s, 'roundRect', 0.603, 0.544, 0.244, 0.051, { rad: 0.009, shadow: true, fill: GOLD });
  sh(s, 'roundRect', 0.603, 0.637, 0.244, 0.051, { rad: 0.009, shadow: true, fill: GOLD });
  tx(s, 0.847, 0.489, 1.426, 0.269, 'Trends', { sz: 10, font: 'MED', align: 'center', color: GOLD });
  tx(s, 2.266, 0.489, 1.426, 0.269, 'Diving', { sz: 10, font: 'MED', align: 'center', color: GOLD });
  tx(s, 3.608, 0.489, 1.426, 0.269, 'Ocean', { sz: 10, font: 'MED', align: 'center', color: GOLD });
  tx(s, 4.91, 0.489, 1.426, 0.269, 'Contact', { sz: 10, font: 'MED', align: 'center', color: GOLD });
  tx(s, 11.302, 5.814, 3.035, 0.337, 'Best Day for Diving in 2021', { sz: 14, bold: true, rot: 90, color: WHITE });
  photo(s, 7.729, 0.971, 4.543, 5.557, 'round2DiagRect', { flipH: true, fill: 'CBCBCB', line: WHITE, lw: 6 });
}

// ---------------------------------------------------------------- 5. Amazing Ocean
function slide05(s) {
  grad(s, rectPts(0, 0, 13.333, 3.453));

  photo(s, 0, 1.271, 5.991, 3.453, 'rect', { fill: 'CBCBCB' });
  photo(s, 7.343, 4.047, 5.991, 3.453, 'rect', { flipH: true, fill: 'CBCBCB' });
  tx(s, 6.522, 1.074, 5.001, 0.64, 'Amazing Ocean', { sz: 32, bold: true, color: WHITE });
  tx(s, 0.847, 0.489, 1.426, 0.269, 'Trends', { sz: 10, font: 'MED', align: 'center', color: GOLD });
  tx(s, 2.266, 0.489, 1.426, 0.269, 'Diving', { sz: 10, font: 'MED', align: 'center', color: GOLD });
  tx(s, 3.608, 0.489, 1.426, 0.269, 'Ocean', { sz: 10, font: 'MED', align: 'center', color: GOLD });
  tx(s, 6.522, 1.714, 5.822, 1.284, cut(LOREM, 236), { align: 'justify', lh: 1.5, color: SILVER });
  sh(s, 'roundRect', 0.453, 4.075, 3.774, 0.966, { fill: NAVY, shadow: true, rad: 0.117 });
  grad(s, roundPts(0.453, 4.075, 3.774, 0.966, 0.117));
  tx(s, 0.841, 4.324, 2.109, 0.303, 'Insert your text', { font: 'SEMI', color: WHITE });
  tx(s, 3.176, 4.26, 0.664, 0.375, '60%', { font: 'SEMI', align: 'right', lh: 1.5, color: WHITE });
  sh(s, 'roundRect', 0.935, 4.69, 2.8, 0.083, { rad: 0.042, fill: 'E7E6E6' });
  sh(s, 'roundRect', 0.935, 4.69, 1.862, 0.083, { rad: 0.042, fill: GOLD });
  tx(s, 0.453, 5.738, 2.991, 0.981, cut(LOREM, 96), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 0.453, 5.483, 2.361, 0.303, 'INSERT TEXT HERE', { font: 'SEMI', color: BLUE });
  tx(s, 3.643, 5.738, 2.991, 0.981, cut(LOREM, 96), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 3.643, 5.483, 2.361, 0.303, 'INSERT TEXT HERE', { font: 'SEMI', color: BLUE });
}

// ---------------------------------------------------------------- 6. Find Amazing Treasure In The Beautiful Ocean
function slide06(s) {
  grad(s, rectPts(0, 2.302, 13.333, 5.198));

  path(s, 1.526, 0.518, 3.89, 3.888, [[0.854, 0], [0.887, 0, 0.921, 0.013, 0.947, 0.038], [0.947, 0.038], [0.999, 0.089, 1, 0.173, 0.949, 0.225], [0.238, 0.948], [0.193, 0.994, 0.124, 1, 0.072, 0.967], [0.051, 0.95], [0.034, 0.929], [0, 0.878, 0.005, 0.809, 0.05, 0.763], [0.761, 0.04], [0.786, 0.014, 0.82, 0.001, 0.854, 0]], { fill: 'CBCBCB' });
  path(s, 6.323, 0.518, 3.89, 3.888, [[0.854, 0], [0.887, 0, 0.921, 0.013, 0.947, 0.038], [0.947, 0.038], [0.999, 0.089, 1, 0.173, 0.949, 0.225], [0.238, 0.948], [0.193, 0.994, 0.124, 1, 0.072, 0.967], [0.051, 0.95], [0.034, 0.929], [0, 0.878, 0.005, 0.809, 0.05, 0.763], [0.761, 0.04], [0.786, 0.014, 0.82, 0.001, 0.854, 0]], { fill: 'CBCBCB' });
  path(s, 3.125, 0.518, 3.89, 3.888, [[0.854, 0], [0.887, 0, 0.921, 0.013, 0.947, 0.038], [0.947, 0.038], [0.999, 0.089, 1, 0.173, 0.949, 0.225], [0.238, 0.948], [0.193, 0.994, 0.124, 1, 0.072, 0.967], [0.051, 0.95], [0.034, 0.929], [0, 0.878, 0.005, 0.809, 0.05, 0.763], [0.761, 0.04], [0.786, 0.014, 0.82, 0.001, 0.854, 0]], { fill: 'CBCBCB' });
  path(s, 7.922, 0.518, 3.89, 3.888, [[0.854, 0], [0.887, 0, 0.921, 0.013, 0.947, 0.038], [0.947, 0.038], [0.999, 0.089, 1, 0.173, 0.949, 0.225], [0.238, 0.948], [0.193, 0.994, 0.124, 1, 0.072, 0.967], [0.051, 0.95], [0.034, 0.929], [0, 0.878, 0.005, 0.809, 0.05, 0.763], [0.761, 0.04], [0.786, 0.014, 0.82, 0.001, 0.854, 0]], { fill: 'CBCBCB' });
  path(s, 4.719, 0.518, 3.896, 3.892, [[0.854, 0], [0.888, 0, 0.921, 0.013, 0.947, 0.038], [0.947, 0.038], [0.999, 0.089, 1, 0.173, 0.949, 0.225], [0.239, 0.947], [0.188, 0.999, 0.105, 1, 0.053, 0.949], [0.053, 0.949], [0.001, 0.898, 0, 0.814, 0.051, 0.762], [0.761, 0.04], [0.787, 0.014, 0.82, 0.001, 0.854, 0]], { fill: 'CBCBCB' });
  tx(s, 0.673, 4.908, 6.209, 1.178, 'Find Amazing Treasure In The Beautiful Ocean', { sz: 32, bold: true, color: GOLD });
  tx(s, 7.596, 4.706, 0.587, 0.589, '01', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'ellipse', fill: GOLD, color: BLUE });
  tx(s, 0.673, 6.086, 5.822, 0.981, cut(LOREM, 197), { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 8.358, 4.617, 4.229, 0.678, cut(LOREM, 90), { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 7.596, 5.578, 0.587, 0.589, '02', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'ellipse', fill: GOLD, color: BLUE });
  tx(s, 8.358, 5.489, 4.229, 0.678, cut(LOREM, 90), { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 7.596, 6.451, 0.587, 0.589, '03', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'ellipse', fill: GOLD, color: BLUE });
  tx(s, 8.358, 6.362, 4.229, 0.678, cut(LOREM, 90), { align: 'justify', lh: 1.5, color: SILVER });
  sh(s, 'ellipse', 1.774, 1.149, 1.975, 1.975, { shadow: true, fill: GOLD, line: WHITE, lw: 6 });
  tx(s, 1.847, 1.547, 1.902, 1.134, [['1', WHITE], ['st', WHITE, { sup: true }], [' ', WHITE, { br: true }], ['Amazing Diving', WHITE, { sz: 20 }]], { sz: 28, bold: true, align: 'center', lh: 0.9 });
}

// ---------------------------------------------------------------- 7. Amazing Ocean
function slide07(s) {
  sh(s, 'rect', 0, 0, 6.667, 7.5, { fill: NAVY });

  photo(s, 0.488, 0.492, 2.226, 4.778, 'roundRect', { rad: 0.14, fill: 'CBCBCB' });
  photo(s, 2.885, 0.492, 3.29, 6.603, 'roundRect', { rad: 0.207, fill: '979797' });
  tx(s, 7.342, 1.23, 4.405, 0.64, 'Amazing Ocean', { sz: 32, bold: true, color: BLUE });
  tx(s, 7.342, 3.417, 1.186, 0.909, '01', { sz: 48, bold: true, font: 'SEMI', align: 'center', color: GOLD });
  tx(s, 8.349, 3.673, 4.143, 0.678, cut(DEV, 91), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 8.349, 3.417, 2.361, 0.303, 'Insert text here', { font: 'SEMI', color: BLUE });
  sh(s, 'roundRect', 0.977, 4.835, 2.915, 1.721, { fill: NAVY, shadow: true, rad: 0.287 });
  grad(s, roundPts(0.977, 4.835, 2.915, 1.721, 0.287));
  tx(s, 1.399, 5.558, 2.542, 0.692, 'Amazing Trending Diving Spot 2021 ', { sz: 14, bold: true, font: 'SEMI', lh: 1.3, color: WHITE });
  tx(s, 1.399, 4.998, 1.313, 0.558, [['14', GOLD], ['th', GOLD, { sup: true }], [' Jan', GOLD]], { sz: 20, bold: true, font: 'SEMI', lh: 1.5 });
  tx(s, 7.342, 0.894, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', color: GOLD });
  tx(s, 7.342, 1.95, 5.277, 1.284, cut(LOREM, 211), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 7.342, 4.533, 1.186, 0.909, '02', { sz: 48, bold: true, font: 'SEMI', align: 'center', color: GOLD });
  tx(s, 8.349, 4.789, 4.143, 0.678, cut(DEV, 91), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 8.349, 4.533, 2.361, 0.303, 'Insert text here', { font: 'SEMI', color: BLUE });
  tx(s, 7.342, 5.65, 1.186, 0.909, '03', { sz: 48, bold: true, font: 'SEMI', align: 'center', color: GOLD });
  tx(s, 8.349, 5.905, 4.143, 0.678, cut(DEV, 91), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 8.349, 5.65, 2.361, 0.303, 'Insert text here', { font: 'SEMI', color: BLUE });
}

// ---------------------------------------------------------------- 8. Matias Missions
function slide08(s) {
  sh(s, 'roundRect', 0.486, 2.557, 3.819, 4.457, { fill: NAVY, shadow: true, rad: 0.298 });
  grad(s, roundPts(0.486, 2.557, 3.819, 4.457, 0.298));
  sh(s, 'roundRect', 4.743, 2.557, 3.819, 4.457, { rad: 0.298, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 9, 2.557, 3.819, 4.457, { rad: 0.298, shadow: true, fill: WHITE });

  sh(s, 'ellipse', 5.065, 5.089, 1.036, 1.036, { fill: null, line: GRAY, lw: 4 });
  sh(s, 'arc', 5.065, 5.089, 1.036, 1.036, { arc: [126, 267], fill: null, line: GOLD, lw: 4 });
  tx(s, 5.121, 5.419, 0.925, 0.444, '48%', { sz: 18, bold: true, font: 'SEMI', align: 'center', lh: 1.2, color: BLUE });
  sh(s, 'ellipse', 0.811, 5.089, 1.036, 1.036, { fill: null, line: WHITE, lw: 4 });
  sh(s, 'arc', 0.811, 5.089, 1.036, 1.036, { arc: [0, 267], fill: null, line: GOLD, lw: 4 });
  tx(s, 0.867, 5.419, 0.925, 0.444, '80%', { sz: 18, bold: true, font: 'SEMI', align: 'center', lh: 1.2, color: GOLD });
  tx(s, 4.741, 0.855, 4.405, 0.64, 'Matias Missions', { sz: 32, bold: true, align: 'center', color: BLUE });
  tx(s, 5.316, 0.519, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', align: 'center', color: GOLD });
  tx(s, 2.017, 1.575, 9.853, 0.678, cut(LOREM, 211), { align: 'center', lh: 1.5, color: GRAY });
  sh(s, 'ellipse', 9.326, 5.089, 1.036, 1.036, { fill: null, line: GRAY, lw: 4 });
  sh(s, 'arc', 9.326, 5.089, 1.036, 1.036, { arc: [97, 267], fill: null, line: GOLD, lw: 4 });
  tx(s, 9.382, 5.419, 0.925, 0.444, '50%', { sz: 18, bold: true, font: 'SEMI', align: 'center', lh: 1.2, color: BLUE });
  tx(s, 6.155, 5.35, 2.407, 1.284, cut(LOREM, 96), { lh: 1.5, color: GRAY });
  tx(s, 6.155, 5.095, 2.361, 0.303, 'INSERT TEXT HERE', { font: 'SEMI', color: BLUE });
  tx(s, 10.409, 5.35, 2.407, 1.284, cut(LOREM, 96), { lh: 1.5, color: GRAY });
  tx(s, 10.409, 5.095, 2.361, 0.303, 'INSERT TEXT HERE', { font: 'SEMI', color: BLUE });
  tx(s, 1.903, 5.35, 2.407, 1.284, cut(LOREM, 96), { lh: 1.5, color: SILVER });
  tx(s, 1.903, 5.095, 2.361, 0.303, 'INSERT TEXT HERE', { font: 'SEMI', color: GOLD });
  photo(s, 0.486, 2.557, 3.819, 2.229, 'round2SameRect', { fill: 'CBCBCB' });
  photo(s, 4.743, 2.557, 3.819, 2.229, 'round2SameRect', { fill: 'CBCBCB' });
  photo(s, 9, 2.557, 3.819, 2.229, 'round2SameRect', { fill: 'CBCBCB' });
}

// ---------------------------------------------------------------- 9. Why Choose Us
function slide09(s) {
  sh(s, 'rect', 0, 0, 13.333, 7.5, { fill: NAVY });
  grad(s, ovalPts(7.653, 1.264, 4.972, 4.972));

  tx(s, 4.134, 4.437, 2.837, 0.692, 'Best Diving Agency for Amazing Ocean 2021', { sz: 14, bold: true, font: 'SEMI', lh: 1.3, color: WHITE });
  sh(s, 'ellipse', 0.8, 4.177, 1.398, 1.398, { shadow: true, fill: BLUE });
  tx(s, 0.866, 4.624, 1.267, 0.505, '236', { sz: 24, font: 'SEMI', align: 'center', color: WHITE });
  sh(s, 'ellipse', 2.499, 4.177, 1.398, 1.398, { shadow: true, fill: GOLD });
  tx(s, 2.565, 4.624, 1.267, 0.505, '+724', { sz: 24, font: 'SEMI', align: 'center', fill: GOLD, color: WHITE });
  sh(s, 'triangle', 4.245, 5.238, 0.294, 0.254, { rot: 180, fill: GOLD });
  tx(s, 0.884, 2.311, 5.822, 1.587, cut(LOREM, 326), { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 0.884, 2.056, 2.361, 0.303, 'INSERT TEXT HERE', { font: 'SEMI', color: GOLD });
  tx(s, 0.884, 1.46, 4.655, 0.572, 'Why Choose Us', { sz: 28, bold: true, color: WHITE });
  sh(s, 'roundRect', 0.535, 0.45, 0.312, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 0.603, 0.544, 0.244, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 0.603, 0.637, 0.244, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  tx(s, 11.302, 2.364, 3.035, 0.337, 'Best Day for Diving in 2021', { sz: 14, bold: true, rot: 90, color: GOLD });
  tx(s, 0.284, 6.949, 3.408, 0.337, 'Published by: yumnacreative', { sz: 14, bold: true, color: GOLD });
  tx(s, 9.449, 6.417, 1.78, 0.375, 'Watch More', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'roundRect', rad: 0.076, shadow: true, fill: BLUE, color: WHITE });
  tx(s, 0.847, 0.489, 1.426, 0.269, 'Trends', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 2.266, 0.489, 1.426, 0.269, 'Diving', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 3.608, 0.489, 1.426, 0.269, 'Ocean', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 4.91, 0.489, 1.426, 0.269, 'Contact', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  photo(s, 7.986, 1.597, 4.306, 4.305, 'ellipse', { flipH: true, fill: 'CBCBCB', line: WHITE, lw: 6 });
}

// ---------------------------------------------------------------- 10. Our Diving Vision
function slide10(s) {
  grad(s, rectPts(0, 0, 3.529, 3.453));

  photo(s, 0.786, 0.712, 5.546, 5.545, 'teardrop', { flipH: true, fill: 'CBCBCB' });
  sh(s, 'round2SameRect', 1.441, 4.437, 1.15, 3.06, { rot: 90, flipH: true, fill: BLUE });
  tx(s, 0.778, 5.498, 1.614, 0.303, 'INSERT TITLE HERE', { font: 'MED', color: WHITE });
  tx(s, 3.034, 5.468, 0.455, 0.37, '01', { sz: 16, bold: true, align: 'center', color: WHITE });
  tx(s, 0.778, 5.759, 2.522, 0.63, cut(DEV, 56), { sz: 11, lh: 1.5, color: WHITE });
  sh(s, 'round2SameRect', 4.896, 4.457, 1.15, 3.06, { rot: 90, flipH: true, fill: GOLD });
  tx(s, 4.234, 5.518, 1.614, 0.303, 'INSERT TITLE HERE', { font: 'MED', color: WHITE });
  tx(s, 6.412, 5.488, 0.533, 0.37, '02', { sz: 16, bold: true, align: 'center', color: WHITE });
  tx(s, 4.234, 5.779, 2.522, 0.63, cut(DEV, 56), { sz: 11, lh: 1.5, color: WHITE });
  sh(s, 'round2SameRect', 8.368, 4.437, 1.15, 3.06, { rot: 90, flipH: true, fill: BLUE });
  tx(s, 7.706, 5.498, 1.614, 0.303, 'INSERT TITLE HERE', { font: 'MED', color: WHITE });
  tx(s, 9.884, 5.468, 0.533, 0.37, '03', { sz: 16, bold: true, align: 'center', color: WHITE });
  tx(s, 7.706, 5.759, 2.522, 0.63, cut(DEV, 56), { sz: 11, lh: 1.5, color: WHITE });
  tx(s, 6.945, 1.903, 4.405, 0.64, 'Our Diving Vision', { sz: 32, bold: true, color: BLUE });
  tx(s, 6.945, 1.567, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', color: GOLD });
  tx(s, 6.945, 2.568, 5.822, 2.193, LOREM, { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 10.107, 7.005, 3.035, 0.337, 'Best Day for Diving in 2021', { sz: 14, bold: true, color: GOLD });
}

// ---------------------------------------------------------------- 11. 78%
function slide11(s) {
  sh(s, 'rect', 5.986, 3.014, 7.348, 4.486, { fill: NAVY });

  photo(s, 3.314, 0.492, 8.957, 4.122, 'roundRect', { rad: 0.259, fill: 'CBCBCB' });
  sh(s, 'roundRect', 2.001, 0.711, 1.969, 1.798, { rad: 0.3, fill: GOLD });
  tx(s, 2.001, 1.069, 1.969, 0.64, '78%', { sz: 32, font: 'SEMI', align: 'center', color: WHITE });
  tx(s, 2.001, 1.656, 1.969, 0.505, 'Statistic Ocean Population', { font: 'SEMI', align: 'center', color: WHITE });
  sh(s, 'roundRect', 2.001, 2.676, 1.969, 1.798, { rad: 0.3, fill: BLUE });
  tx(s, 2.001, 3.033, 1.969, 0.64, '22%', { sz: 32, font: 'SEMI', align: 'center', color: WHITE });
  tx(s, 2.001, 3.62, 1.969, 0.505, 'Statistic Ocean Population', { font: 'SEMI', align: 'center', color: WHITE });
  tx(s, 0.453, 5.844, 5.097, 1.284, cut(LOREM, 221), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 0.453, 5.589, 2.361, 0.303, 'INSERT TEXT HERE', { font: 'SEMI', color: SKY });
  tx(s, 0.453, 4.99, 3.808, 0.572, 'Ocean Statistic', { sz: 28, bold: true, color: BLUE });
  tx(s, 6.741, 5.255, 0.587, 0.589, '01', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'ellipse', fill: GOLD, color: BLUE });
  tx(s, 7.503, 5.276, 4.229, 0.678, cut(LOREM, 90), { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 7.503, 5.086, 2.361, 0.303, 'Insert text here', { font: 'SEMI', color: GOLD });
  tx(s, 6.741, 6.308, 0.587, 0.589, '02', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'ellipse', fill: GOLD, color: BLUE });
  tx(s, 7.503, 6.329, 4.229, 0.678, cut(LOREM, 90), { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 7.503, 6.139, 2.361, 0.303, 'Insert text here', { font: 'SEMI', color: GOLD });
}

// ---------------------------------------------------------------- 12. Matias Diving
function slide12(s) {
  grad(s, [[4.24, 0], [5.3, 7.5], [0, 7.5], [0, 0]], { angle: 240 });

  sh(s, 'roundRect', 3.896, 5.4, 8.871, 1.648, { rad: 0.129, shadow: true, fill: WHITE });
  grad(s, ovalPts(6.538, 5.814, 0.815, 0.815), { stops: [BLUE2, BLUE3] });
  path(s, 6.824, 5.996, 0.242, 0.45, [[1, 0.177], [1, 0.177], [0.714, 0.177, 0.714, 0.177, 0.714, 0.177], [0.681, 0.177, 0.645, 0.197, 0.645, 0.237], [0.645, 0.354, 0.645, 0.354, 0.645, 0.354], [1, 0.354, 1, 0.354, 1, 0.354], [1, 0.511, 1, 0.511, 1, 0.511], [0.645, 0.511, 0.645, 0.511, 0.645, 0.511], [0.645, 1, 0.645, 1, 0.645, 1], [0.319, 1, 0.319, 1, 0.319, 1], [0.319, 0.511, 0.319, 0.511, 0.319, 0.511], [0, 0.511, 0, 0.511, 0, 0.511], [0, 0.354, 0, 0.354, 0, 0.354], [0.319, 0.354, 0.319, 0.354, 0.319, 0.354], [0.319, 0.257, 0.319, 0.257, 0.319, 0.257], [0.319, 0.119, 0.464, 0, 0.714, 0], [1, 0, 1, 0, 1, 0], [1, 0.177]], { fill: WHITE });
  sh(s, 'ellipse', 5.412, 5.814, 0.815, 0.815, { fill: GOLD, line: WHITE, lw: 1.75 });
  path(s, 5.618, 6.153, 0.415, 0.288, [[0.983, 0.854], [0.983, 0.927, 0.932, 0.976, 0.881, 0.976], [0.763, 1, 0.627, 1, 0.508, 1], [0.373, 1, 0.254, 1, 0.119, 0.976], [0.068, 0.976, 0.034, 0.927, 0.017, 0.854], [0, 0.732, 0, 0.61, 0, 0.512], [0, 0.39, 0, 0.268, 0.017, 0.171], [0.034, 0.073, 0.068, 0.024, 0.119, 0.024], [0.254, 0, 0.373, 0, 0.508, 0], [0.627, 0, 0.763, 0, 0.881, 0.024], [0.932, 0.024, 0.983, 0.073, 0.983, 0.171], [1, 0.268, 1, 0.39, 1, 0.512], [1, 0.61, 1, 0.732, 0.983, 0.854]], { fill: WHITE });
  path(s, 5.646, 6.202, 0.091, 0.182, [[1, 0.154], [1, 0, 1, 0, 1, 0], [0, 0, 0, 0, 0, 0], [0, 0.154, 0, 0.154, 0, 0.154], [0.385, 0.154, 0.385, 0.154, 0.385, 0.154], [0.385, 1, 0.385, 1, 0.385, 1], [0.692, 1, 0.692, 1, 0.692, 1], [0.692, 0.154, 0.692, 0.154, 0.692, 0.154], [1, 0.154]], { fill: WHITE });
  path(s, 5.681, 5.936, 0.098, 0.189, [[1, 0], [0.714, 0.593, 0.714, 0.593, 0.714, 0.593], [0.714, 1, 0.714, 1, 0.714, 1], [0.429, 1, 0.429, 1, 0.429, 1], [0.429, 0.593, 0.429, 0.593, 0.429, 0.593], [0.357, 0.519, 0.286, 0.444, 0.214, 0.296], [0.143, 0.185, 0.071, 0.111, 0, 0], [0.357, 0, 0.357, 0, 0.357, 0], [0.571, 0.407, 0.571, 0.407, 0.571, 0.407], [0.714, 0, 0.714, 0, 0.714, 0], [1, 0]], { fill: WHITE });
  path(s, 5.737, 6.252, 0.077, 0.14, [[1, 0.95], [1, 0, 1, 0, 1, 0], [0.636, 0, 0.636, 0, 0.636, 0], [0.636, 0.75, 0.636, 0.75, 0.636, 0.75], [0.545, 0.8, 0.545, 0.8, 0.455, 0.8], [0.364, 0.8, 0.364, 0.8, 0.364, 0.8], [0.364, 0.75, 0.364, 0.75, 0.364, 0.7], [0.364, 0, 0.364, 0, 0.364, 0], [0, 0, 0, 0, 0, 0], [0, 0.75, 0, 0.75, 0, 0.75], [0, 0.85, 0, 0.9, 0.091, 0.9], [0.091, 0.95, 0.182, 1, 0.273, 1], [0.364, 1, 0.545, 0.95, 0.636, 0.85], [0.636, 0.95, 0.636, 0.95, 0.636, 0.95], [1, 0.95]], { fill: WHITE });
  path(s, 5.78, 5.985, 0.077, 0.147, [[1, 0.667], [1, 0.762, 1, 0.81, 0.909, 0.857], [0.818, 0.952, 0.727, 1, 0.545, 1], [0.364, 1, 0.273, 0.952, 0.182, 0.857], [0.091, 0.81, 0, 0.762, 0, 0.667], [0, 0.333, 0, 0.333, 0, 0.333], [0, 0.238, 0.091, 0.143, 0.182, 0.095], [0.273, 0.048, 0.364, 0, 0.545, 0], [0.727, 0, 0.818, 0.048, 0.909, 0.095], [1, 0.143, 1, 0.238, 1, 0.333], [1, 0.667]], { fill: WHITE });
  path(s, 5.808, 6.006, 0.021, 0.098, [[1, 0.214], [1, 0.071, 1, 0, 0.667, 0], [0.333, 0, 0, 0.071, 0, 0.214], [0, 0.786, 0, 0.786, 0, 0.786], [0, 0.929, 0.333, 1, 0.667, 1], [1, 1, 1, 0.929, 1, 0.786], [1, 0.214]], { fill: WHITE });
  path(s, 5.836, 6.202, 0.07, 0.189, [[1, 0.481], [1, 0.407, 1, 0.333, 1, 0.333], [1, 0.259, 0.9, 0.222, 0.7, 0.222], [0.6, 0.222, 0.5, 0.259, 0.3, 0.296], [0.3, 0, 0.3, 0, 0.3, 0], [0, 0, 0, 0, 0, 0], [0, 0.963, 0, 0.963, 0, 0.963], [0.3, 0.963, 0.3, 0.963, 0.3, 0.963], [0.3, 0.889, 0.3, 0.889, 0.3, 0.889], [0.5, 0.963, 0.6, 1, 0.7, 1], [0.9, 1, 1, 0.963, 1, 0.889], [1, 0.889, 1, 0.815, 1, 0.741], [1, 0.481]], { fill: WHITE });
  path(s, 5.857, 6.273, 0.028, 0.091, [[1, 0.846], [1, 1, 0.75, 1, 0.5, 1], [0.5, 1, 0.25, 1, 0, 1], [0, 0, 0, 0, 0, 0], [0.25, 0, 0.5, 0, 0.5, 0], [0.75, 0, 1, 0.077, 1, 0.154], [1, 0.846]], { fill: WHITE });
  path(s, 5.878, 5.985, 0.077, 0.147, [[1, 0.952], [0.636, 0.952, 0.636, 0.952, 0.636, 0.952], [0.636, 0.857, 0.636, 0.857, 0.636, 0.857], [0.455, 0.952, 0.364, 1, 0.273, 1], [0.182, 1, 0.091, 0.952, 0, 0.905], [0, 0.857, 0, 0.81, 0, 0.762], [0, 0, 0, 0, 0, 0], [0.364, 0, 0.364, 0, 0.364, 0], [0.364, 0.714, 0.364, 0.714, 0.364, 0.714], [0.364, 0.762, 0.364, 0.762, 0.364, 0.762], [0.364, 0.81, 0.364, 0.81, 0.455, 0.81], [0.455, 0.81, 0.545, 0.81, 0.636, 0.714], [0.636, 0, 0.636, 0, 0.636, 0], [1, 0, 1, 0, 1, 0], [1, 0.952]], { fill: WHITE });
  path(s, 5.927, 6.245, 0.077, 0.147, [[1, 0.667], [0.636, 0.667, 0.636, 0.667, 0.636, 0.667], [0.636, 0.714, 0.636, 0.762, 0.636, 0.762], [0.636, 0.81, 0.545, 0.81, 0.545, 0.81], [0.364, 0.81, 0.364, 0.81, 0.364, 0.714], [0.364, 0.524, 0.364, 0.524, 0.364, 0.524], [1, 0.524, 1, 0.524, 1, 0.524], [1, 0.333, 1, 0.333, 1, 0.333], [1, 0.238, 1, 0.143, 0.909, 0.095], [0.818, 0.048, 0.636, 0, 0.545, 0], [0.364, 0, 0.182, 0.048, 0.091, 0.095], [0.091, 0.143, 0, 0.238, 0, 0.333], [0, 0.667, 0, 0.667, 0, 0.667], [0, 0.762, 0.091, 0.857, 0.091, 0.905], [0.182, 0.952, 0.364, 1, 0.545, 1], [0.727, 1, 0.818, 0.952, 0.909, 0.905], [0.909, 0.857, 1, 0.81, 1, 0.762], [1, 0.762, 1, 0.714, 1, 0.667]], { fill: WHITE });
  path(s, 5.955, 6.273, 0.021, 0.028, [[1, 1], [0, 1, 0, 1, 0, 1], [0, 0.5, 0, 0.5, 0, 0.5], [0, 0.25, 0, 0, 0.667, 0], [1, 0, 1, 0.25, 1, 0.5], [1, 1]], { fill: WHITE });
  tx(s, 6.945, 1.089, 4.405, 0.64, 'Matias Diving', { sz: 32, bold: true, color: BLUE });
  tx(s, 6.945, 0.752, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', color: GOLD });
  tx(s, 6.945, 1.754, 5.822, 0.678, cut(LOREM, 132), { align: 'justify', lh: 1.5, color: GRAY });
  grad(s, ovalPts(4.278, 5.814, 0.815, 0.815), { stops: [BLUE2, BLUE3] });
  path(s, 4.475, 6.199, 0.421, 0.237, [[0.812, 0.075], [0.812, 0.075], [0.812, 0.372, 0.668, 0.632, 0.501, 0.632], [0.334, 0.632, 0.186, 0.372, 0.186, 0.075], [0.186, 0.033, 0.207, 0.033, 0.207, 0], [0, 0, 0, 0, 0, 0], [0, 0.703, 0, 0.703, 0, 0.703], [0, 0.854, 0.061, 1, 0.146, 1], [0.854, 1, 0.854, 1, 0.854, 1], [0.936, 1, 1, 0.854, 1, 0.703], [1, 0, 1, 0, 1, 0], [0.793, 0, 0.793, 0, 0.793, 0], [0.793, 0.033, 0.812, 0.033, 0.812, 0.075]], { fill: WHITE });
  path(s, 4.475, 6.006, 0.421, 0.141, [[0.854, 0], [0.854, 0], [0.146, 0, 0.146, 0, 0.146, 0], [0.061, 0, 0, 0.254, 0, 0.5], [0, 1, 0, 1, 0, 1], [0.249, 1, 0.249, 1, 0.249, 1], [0.311, 0.746, 0.395, 0.627, 0.501, 0.627], [0.605, 0.627, 0.687, 0.746, 0.751, 1], [1, 1, 1, 1, 1, 1], [1, 0.5, 1, 0.5, 1, 0.5], [1, 0.254, 0.936, 0, 0.854, 0]], { fill: WHITE });
  path(s, 4.8, 6.042, 0.061, 0.06, [[1, 0.869], [1, 0.869], [1, 0.869, 1, 1, 0.855, 1], [0.274, 1, 0.274, 1, 0.274, 1], [0.145, 1, 0, 0.869, 0, 0.869], [0, 0.279, 0, 0.279, 0, 0.279], [0, 0.131, 0.145, 0, 0.274, 0], [0.855, 0, 0.855, 0, 0.855, 0], [1, 0, 1, 0.131, 1, 0.279], [1, 0.869]], { fill: WHITE });
  path(s, 4.606, 6.138, 0.158, 0.158, [[1, 0.5], [1, 0.5], [1, 0.219, 0.781, 0, 0.506, 0], [0.225, 0, 0, 0.219, 0, 0.5], [0, 0.775, 0.225, 1, 0.506, 1], [0.781, 1, 1, 0.775, 1, 0.5]], { fill: WHITE });
  tx(s, 7.636, 5.567, 2.361, 0.303, 'Insert text here', { font: 'SEMI', color: BLUE });
  tx(s, 7.636, 5.783, 4.624, 0.981, cut(LOREM, 158), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 6.945, 2.616, 5.822, 2.193, LOREM, { align: 'justify', lh: 1.5, color: GRAY });
  photo(s, 0.815, 0, 2.7, 2.507, 'rect', { fill: '979797' });
  photo(s, 3.815, 0, 2.557, 1.9, 'rect', { fill: 'CBCBCB' });
  photo(s, 3.815, 2.114, 2.557, 3.071, 'rect', { fill: '979797' });
  photo(s, 0.815, 2.786, 2.7, 3.186, 'rect', { fill: 'CBCBCB' });
}

// ---------------------------------------------------------------- 13. Expert Staff
function slide13(s) {
  grad(s, [[4.24, 7.5], [5.3, 0], [0, 0], [0, 7.5]], { angle: 120 });

  photo(s, 4.305, 0, 2.7, 6.029, 'round2SameRect', { flipH: true, fill: 'CBCBCB' });
  photo(s, 7.272, 0, 2.7, 6.029, 'round2SameRect', { flipH: true, fill: 'CBCBCB' });
  photo(s, 10.239, 0, 2.7, 6.029, 'round2SameRect', { flipH: true, fill: 'CBCBCB' });
  tx(s, 0.395, 2.272, 3.167, 0.64, 'Expert Staff', { sz: 32, bold: true, color: GOLD });
  tx(s, 0.395, 2.97, 3.359, 1.89, 'Re paragone creatura ise acerbita ai guardava lasciami vi. Entro tue forza miele mazzo per pur oltre sul. Lo ti il gabbie quantozeo lancio. Fato mare arme tu anch vi ous ve mine riso. Poi affannata ami cresciuto', { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 0.395, 4.794, 3.359, 0.678, 'Re paragone creatura ise acerbita ai guardava lasciami vi. Entro', { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 4.406, 4.962, 2.392, 0.438, 'Celsa Breen', { sz: 20, bold: true, color: WHITE });
  tx(s, 4.406, 4.625, 1.177, 0.303, 'Expert Staff', { font: 'SEMI', color: WHITE });
  tx(s, 7.402, 4.962, 2.392, 0.438, 'Shella Bock', { sz: 20, bold: true, color: WHITE });
  tx(s, 7.402, 4.625, 1.177, 0.303, 'Expert Staff', { font: 'SEMI', color: WHITE });
  tx(s, 10.369, 4.962, 2.392, 0.438, 'Season Toro', { sz: 20, bold: true, color: WHITE });
  tx(s, 10.369, 4.625, 1.177, 0.303, 'Expert Staff', { font: 'SEMI', color: WHITE });
  tx(s, 10.107, 7.005, 3.035, 0.337, 'Best Day for Diving in 2021', { sz: 14, bold: true, color: GOLD });
  tx(s, 0.41, 5.654, 1.78, 0.375, 'Watch More', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'roundRect', rad: 0.076, shadow: true, fill: BLUE, color: WHITE });
  sh(s, 'roundRect', 0.535, 0.45, 0.312, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 0.603, 0.544, 0.244, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 0.603, 0.637, 0.244, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  tx(s, 0.847, 0.489, 1.426, 0.269, 'Trends', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 2.266, 0.489, 1.426, 0.269, 'Diving', { sz: 10, font: 'MED', align: 'center', color: WHITE });
}

// ---------------------------------------------------------------- 14. Neva Vincent
function slide14(s) {
  sh(s, 'rect', 11.543, 0, 1.79, 7.5, { fill: GOLD });

  photo(s, 6.271, 0.414, 6.571, 4.356, 'roundRect', { rad: 0.274, fill: 'CBCBCB' });
  sh(s, 'roundRect', 8.163, 5.005, 3.774, 0.966, { fill: NAVY, shadow: true, rad: 0.117 });
  grad(s, roundPts(8.163, 5.005, 3.774, 0.966, 0.117));
  tx(s, 8.551, 5.254, 2.109, 0.303, 'Insert your text', { font: 'SEMI', color: WHITE });
  tx(s, 10.886, 5.19, 0.664, 0.375, '60%', { font: 'SEMI', align: 'right', lh: 1.5, color: WHITE });
  sh(s, 'roundRect', 8.645, 5.62, 2.8, 0.083, { rad: 0.042, fill: 'E7E6E6' });
  sh(s, 'roundRect', 8.645, 5.62, 1.862, 0.083, { rad: 0.042, fill: GOLD });
  sh(s, 'roundRect', 8.163, 6.15, 3.774, 0.966, { fill: NAVY, shadow: true, rad: 0.117 });
  grad(s, roundPts(8.163, 6.15, 3.774, 0.966, 0.117));
  tx(s, 8.551, 6.399, 2.109, 0.303, 'Insert your text', { font: 'SEMI', color: WHITE });
  tx(s, 10.886, 6.334, 0.664, 0.375, '56%', { font: 'SEMI', align: 'right', lh: 1.5, color: WHITE });
  sh(s, 'roundRect', 8.645, 6.765, 2.8, 0.083, { rad: 0.042, fill: 'E7E6E6' });
  sh(s, 'roundRect', 8.645, 6.765, 1.862, 0.083, { rad: 0.042, fill: GOLD });
  tx(s, 0.575, 2.195, 4.405, 0.64, 'Neva Vincent', { sz: 32, bold: true, color: BLUE });
  tx(s, 0.575, 1.859, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', color: GOLD });
  tx(s, 0.575, 2.86, 4.93, 0.981, cut(LOREM, 132), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 0.575, 3.866, 4.93, 2.799, LOREM, { align: 'justify', lh: 1.5, color: GRAY });
  sh(s, 'ellipse', 5.284, 0.22, 1.975, 1.975, { shadow: true, fill: BLUE, line: WHITE, lw: 6 });
  tx(s, 5.357, 0.618, 1.902, 1.134, [['1', WHITE], ['st', WHITE, { sup: true }], [' ', WHITE, { br: true }], ['Amazing Diving', WHITE, { sz: 20 }]], { sz: 28, bold: true, align: 'center', lh: 0.9 });
}

// ---------------------------------------------------------------- 15. Professional Staff
function slide15(s) {
  grad(s, rectPts(0, 0, 5.3, 7.5), { angle: 120 });
  sh(s, 'rect', 0, 0, 2.868, 3.014, { fill: GOLD });

  photo(s, 0.786, 0.429, 2.457, 2.771, 'roundRect', { flipH: true, rad: 0.309, fill: 'CBCBCB' });
  photo(s, 3.504, 0.936, 2.457, 2.771, 'roundRect', { flipH: true, rad: 0.309, fill: 'CBCBCB' });
  photo(s, 3.504, 4.032, 2.457, 2.771, 'roundRect', { flipH: true, rad: 0.309, fill: 'CBCBCB' });
  photo(s, 0.786, 3.525, 2.457, 2.771, 'roundRect', { flipH: true, rad: 0.309, fill: 'CBCBCB' });
  tx(s, 6.667, 2.669, 5.822, 2.193, LOREM, { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 6.667, 2.414, 2.361, 0.303, 'INSERT TEXT HERE', { font: 'SEMI', color: SKY });
  tx(s, 6.667, 1.714, 4.101, 0.572, 'Professional Staff', { sz: 28, bold: true, color: BLUE });
  tx(s, 6.667, 4.911, 5.822, 0.678, cut(LOREM, 133), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 10.107, 7.005, 3.035, 0.337, 'Best Day for Diving in 2021', { sz: 14, bold: true, color: GOLD });
  tx(s, 0.851, 5.589, 2.392, 0.438, 'Darius Pipkin', { sz: 20, bold: true, color: WHITE });
  tx(s, 0.851, 5.252, 1.177, 0.303, 'Expert Staff', { font: 'SEMI', color: WHITE });
  tx(s, 3.533, 5.996, 2.392, 0.438, 'Ilene Stacy', { sz: 20, bold: true, color: WHITE });
  tx(s, 3.533, 5.659, 1.177, 0.303, 'Expert Staff', { font: 'SEMI', color: WHITE });
  tx(s, 0.851, 2.595, 2.392, 0.438, 'Brett Hartwell', { sz: 20, bold: true, color: WHITE });
  tx(s, 0.851, 2.258, 1.177, 0.303, 'Expert Staff', { font: 'SEMI', color: WHITE });
  tx(s, 3.533, 3.002, 2.392, 0.438, 'Lesley Weir', { sz: 20, bold: true, color: WHITE });
  tx(s, 3.533, 2.665, 1.177, 0.303, 'Expert Staff', { font: 'SEMI', color: WHITE });
}

// ---------------------------------------------------------------- 16. Expert Service
function slide16(s) {
  grad(s, rectPts(8.543, 0, 4.79, 2.443), { angle: 120 });

  tx(s, 6.482, 4.778, 2.573, 0.303, 'Fisrt Diving Service', { font: 'SEMI', color: BLUE });
  tx(s, 6.482, 5.127, 2.573, 1.587, 'Dev certa volle sopra anime animo qua. Tenta anima la no aveva ch rombo le. Se da distrutta liberarli usignuoli. Ora afa rimorso', { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 0.887, 4.765, 4.405, 0.64, 'Expert Service', { sz: 32, bold: true, color: BLUE });
  tx(s, 0.887, 4.429, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', color: GOLD });
  tx(s, 0.887, 5.43, 4.93, 1.284, cut(LOREM, 211), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 9.38, 4.778, 2.573, 0.303, 'Second Diving Service', { font: 'SEMI', color: BLUE });
  tx(s, 9.38, 5.127, 2.573, 1.587, 'Dev certa volle sopra anime animo qua. Tenta anima la no aveva ch rombo le. Se da distrutta liberarli usignuoli. Ora afa rimorso', { align: 'justify', lh: 1.5, color: GRAY });
  photo(s, 0.972, 0.429, 4.671, 2.886, 'roundRect', { flipH: true, rad: 0.364, fill: 'CBCBCB' });
  photo(s, 3.643, 1.129, 4.671, 2.886, 'roundRect', { flipH: true, rad: 0.364, fill: '979797' });
  photo(s, 7.3, 0.429, 4.671, 2.886, 'roundRect', { flipH: true, rad: 0.364, fill: 'CBCBCB' });
}

// ---------------------------------------------------------------- 17. Expert Service
function slide17(s) {
  photo(s, 1.048, 0, 4.224, 5.943, 'round2SameRect', { flipH: true, fill: 'CBCBCB' });
  sh(s, 'rect', 4.057, 2.386, 3.819, 4.457, { shadow: true, fill: WHITE });
  sh(s, 'round1Rect', 0, 4.014, 3.449, 3.486, { fill: NAVY, shadow: true, rad: 0.575 });
  grad(s, roundPts(0, 4.014, 3.449, 3.486, 0.575));
  tx(s, 8.152, 1.548, 4.405, 0.64, 'Expert Service', { sz: 32, bold: true, color: BLUE });
  tx(s, 8.152, 1.211, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', color: GOLD });
  tx(s, 8.152, 2.213, 4.405, 1.587, cut(LOREM, 211), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 8.152, 3.825, 4.405, 0.981, cut(LOREM, 142), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 0.413, 4.545, 1.614, 0.303, 'INSERT TITLE HERE', { font: 'MED', color: WHITE });
  tx(s, 0.413, 4.806, 2.522, 0.63, cut(DEV, 56), { sz: 11, lh: 1.5, color: SILVER });
  tx(s, 0.413, 5.651, 1.614, 0.303, 'INSERT TITLE HERE', { font: 'MED', color: WHITE });
  tx(s, 0.413, 5.912, 2.522, 0.63, cut(DEV, 56), { sz: 11, lh: 1.5, color: SILVER });
  tx(s, 4.333, 5.912, 2.542, 0.692, 'Amazing Trending Diving Spot 2021 ', { sz: 14, bold: true, font: 'SEMI', lh: 1.3, color: GOLD });
  tx(s, 4.333, 5.352, 1.313, 0.558, [['14', BLUE], ['th', BLUE, { sup: true }], [' Jan', BLUE]], { sz: 20, bold: true, font: 'SEMI', lh: 1.5 });
  tx(s, 10.107, 7.005, 3.035, 0.337, 'Best Day for Diving in 2021', { sz: 14, bold: true, color: GOLD });
  tx(s, 10.734, 4.977, 1.78, 0.375, 'Watch More', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'roundRect', rad: 0.076, shadow: true, fill: BLUE, color: WHITE });
  photo(s, 4.333, 2.657, 3.267, 2.671, 'rect', { fill: 'CBCBCB' });
}

// ---------------------------------------------------------------- 18. Slide 18
function slide18(s) {
  grad(s, rectPts(0, 0, 5.3, 7.5), { angle: 120 });
  sh(s, 'rect', 0.396, 0.815, 12.541, 6.169, { shadow: true, fill: WHITE });
  sh(s, 'ellipse', 9.113, 1.086, 2.756, 2.756, { fill: GOLD });
  sh(s, 'ellipse', 5.289, 1.086, 2.756, 2.756, { fill: GOLD });
  sh(s, 'ellipse', 1.464, 1.086, 2.756, 2.756, { fill: GOLD });

  tx(s, 1.53, 4.194, 2.573, 0.303, 'FIRST EQUIPMENT', { font: 'SEMI', align: 'center', color: BLUE });
  tx(s, 1.53, 4.544, 2.573, 1.587, 'Dev certa volle sopra anime animo qua. Tenta anima la no aveva ch rombo le. Se da distrutta liberarli usignuoli. Ora afa rimorso', { align: 'center', lh: 1.5, color: GRAY });
  tx(s, 5.331, 4.194, 2.573, 0.303, 'SECOND EQUIPMENT', { font: 'SEMI', align: 'center', color: BLUE });
  tx(s, 5.331, 4.544, 2.573, 1.587, 'Dev certa volle sopra anime animo qua. Tenta anima la no aveva ch rombo le. Se da distrutta liberarli usignuoli. Ora afa rimorso', { align: 'center', lh: 1.5, color: GRAY });
  tx(s, 9.23, 4.194, 2.573, 0.303, 'THIRD EQUIPMENT', { font: 'SEMI', align: 'center', color: BLUE });
  tx(s, 9.23, 4.544, 2.573, 1.587, 'Dev certa volle sopra anime animo qua. Tenta anima la no aveva ch rombo le. Se da distrutta liberarli usignuoli. Ora afa rimorso', { align: 'center', lh: 1.5, color: GRAY });
  tx(s, 9.626, 6.297, 1.78, 0.375, 'Watch More', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'roundRect', rad: 0.076, shadow: true, fill: BLUE, color: WHITE });
  tx(s, 5.777, 6.297, 1.78, 0.375, 'Watch More', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'roundRect', rad: 0.076, shadow: true, fill: BLUE, color: WHITE });
  tx(s, 1.927, 6.297, 1.78, 0.375, 'Watch More', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'roundRect', rad: 0.076, shadow: true, fill: BLUE, color: WHITE });
  sh(s, 'roundRect', 0.535, 0.45, 0.312, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 0.603, 0.544, 0.244, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 0.603, 0.637, 0.244, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  tx(s, 10.997, 2.296, 3.035, 0.337, 'Best Day for Diving in 2021', { sz: 14, bold: true, rot: 90, color: GOLD });
  tx(s, 0.847, 0.489, 1.426, 0.269, 'Trends', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 2.266, 0.489, 1.426, 0.269, 'Diving', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 3.608, 0.489, 1.426, 0.269, 'Ocean', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  photo(s, 1.582, 1.203, 2.523, 2.523, 'ellipse', { flipH: true, fill: 'CBCBCB' });
  photo(s, 5.382, 1.203, 2.523, 2.523, 'ellipse', { flipH: true, fill: 'CBCBCB' });
  photo(s, 9.229, 1.203, 2.523, 2.523, 'ellipse', { flipH: true, fill: 'CBCBCB' });
}

// ---------------------------------------------------------------- 19. Expert Service
function slide19(s) {
  sh(s, 'flowChartManualInput', 6.933, 1.1, 7.5, 5.3, { rot: -90, fill: NAVY });

  sh(s, 'roundRect', 1.86, 4.094, 0.723, 0.712, { rad: 0.086, shadow: true, fill: BLUE });
  path(s, 1.952, 4.248, 0.5, 0.266, [[0.03, 0.924], [0.007, 0.955], [0, 0.965, 0, 0.98, 0.006, 0.99], [0.012, 0.999, 0.023, 1, 0.029, 0.991], [0.082, 0.919], [0.1, 0.894, 0.129, 0.894, 0.147, 0.919], [0.156, 0.932], [0.187, 0.975, 0.235, 0.975, 0.266, 0.932], [0.275, 0.919], [0.293, 0.894, 0.322, 0.894, 0.339, 0.919], [0.349, 0.932], [0.38, 0.975, 0.428, 0.975, 0.458, 0.932], [0.468, 0.919], [0.486, 0.894, 0.514, 0.894, 0.532, 0.919], [0.542, 0.932], [0.573, 0.975, 0.621, 0.975, 0.651, 0.932], [0.661, 0.919], [0.679, 0.894, 0.707, 0.894, 0.725, 0.919], [0.735, 0.932], [0.765, 0.975, 0.813, 0.975, 0.844, 0.932], [0.854, 0.919], [0.872, 0.894, 0.9, 0.894, 0.918, 0.919], [0.971, 0.991], [0.975, 0.997, 0.982, 0.999, 0.987, 0.997], [0.993, 0.994, 0.997, 0.987, 0.999, 0.979], [1, 0.97, 0.998, 0.961, 0.994, 0.955], [0.941, 0.884], [0.928, 0.867, 0.912, 0.856, 0.895, 0.853], [0.899, 0.843, 0.902, 0.832, 0.905, 0.821], [0.998, 0.446], [1, 0.439, 0.999, 0.431, 0.996, 0.424], [0.993, 0.418, 0.988, 0.414, 0.982, 0.414], [0.91, 0.414], [0.79, 0.042], [0.782, 0.016, 0.764, 0, 0.745, 0], [0.417, 0], [0.39, 0, 0.367, 0.033, 0.367, 0.073], [0.367, 0.098, 0.377, 0.12, 0.393, 0.129], [0.442, 0.157], [0.445, 0.16, 0.447, 0.165, 0.447, 0.171], [0.387, 0.608], [0.234, 0.608], [0.234, 0.511], [0.234, 0.457, 0.205, 0.414, 0.168, 0.414], [0.131, 0.414, 0.101, 0.457, 0.101, 0.511], [0.101, 0.779], [0.035, 0.779], [0.035, 0.584], [0.035, 0.57, 0.027, 0.56, 0.018, 0.56], [0.009, 0.56, 0.002, 0.57, 0.002, 0.584], [0.002, 0.876], [0.002, 0.9, 0.014, 0.92, 0.03, 0.924]], { fill: WHITE });
  path(s, 1.969, 4.468, 0.066, 0.015, [[0.239, 0.892], [0, 0.892], [0, 0], [1, 0], [1, 1], [0.787, 0.316, 0.467, 0.271, 0.239, 0.892]], { fill: WHITE });
  path(s, 2.224, 4.371, 0.207, 0.039, [[0.912, 1], [0.471, 1], [0.523, 0.785], [0.538, 0.719, 0.538, 0.615, 0.522, 0.551], [0.507, 0.487, 0.482, 0.486, 0.466, 0.549], [0.357, 1], [0.27, 1], [0.322, 0.785], [0.337, 0.719, 0.337, 0.615, 0.321, 0.551], [0.306, 0.487, 0.28, 0.486, 0.265, 0.549], [0.156, 1], [0, 1], [0.053, 0.864, 0.102, 0.71, 0.148, 0.538], [0.23, 0.25, 0.3, 0, 0.414, 0], [1, 0]], { fill: WHITE });
  path(s, 2.235, 4.287, 0.137, 0.045, [[1, 1], [0, 1], [0.095, 0], [0.801, 0]], { fill: WHITE });
  path(s, 2.152, 4.261, 0.236, 0.149, [[0.167, 0.235], [0.175, 0.184, 0.153, 0.133, 0.114, 0.114], [0.011, 0.063], [0.004, 0.06, 0, 0.052, 0, 0.043], [0, 0.019, 0.016, 0, 0.035, 0], [0.729, 0], [0.743, 0, 0.755, 0.01, 0.761, 0.025], [0.785, 0.087], [0.407, 0.087], [0.374, 0.087, 0.345, 0.115, 0.338, 0.155], [0.283, 0.46], [0.278, 0.486, 0.284, 0.513, 0.297, 0.533], [0.311, 0.554, 0.331, 0.565, 0.352, 0.565], [0.967, 0.565], [1, 0.652], [0.668, 0.652], [0.543, 0.652, 0.466, 0.734, 0.391, 0.814], [0.305, 0.904, 0.216, 0.997, 0.043, 1]], { fill: WHITE });
  path(s, 2.052, 4.423, 0.355, 0.071, [[0.305, 0], [1, 0], [0.983, 0.184], [0.164, 0.184], [0.151, 0.184, 0.141, 0.225, 0.141, 0.276], [0.141, 0.326, 0.151, 0.368, 0.164, 0.368], [0.966, 0.368], [0.948, 0.559], [0.941, 0.64, 0.93, 0.715, 0.916, 0.783], [0.906, 0.801, 0.897, 0.826, 0.889, 0.856], [0.875, 0.905], [0.85, 1, 0.81, 1, 0.784, 0.905], [0.771, 0.856], [0.728, 0.696, 0.66, 0.696, 0.617, 0.856], [0.603, 0.905], [0.578, 1, 0.538, 1, 0.513, 0.905], [0.499, 0.856], [0.456, 0.695, 0.388, 0.695, 0.345, 0.856], [0.331, 0.905], [0.306, 1, 0.266, 1, 0.241, 0.905], [0.227, 0.856], [0.184, 0.695, 0.116, 0.695, 0.073, 0.856], [0.059, 0.905], [0.044, 0.963, 0.021, 0.988, 0, 0.969], [0, 0.368], [0.07, 0.368], [0.083, 0.368, 0.094, 0.326, 0.094, 0.276], [0.094, 0.225, 0.083, 0.184, 0.07, 0.184], [0, 0.184], [0, 0]], { fill: WHITE });
  path(s, 2.019, 4.371, 0.033, 0.084, [[0, 0.154], [0, 0.069, 0.224, 0, 0.5, 0], [0.776, 0, 1, 0.069, 1, 0.154], [1, 0.462], [0.724, 0.462, 0.5, 0.53, 0.5, 0.615], [0.5, 1], [0, 1]], { fill: WHITE });
  path(s, 1.952, 4.511, 0.5, 0.042, [[0.994, 0.722], [0.941, 0.268], [0.91, 0, 0.862, 0, 0.831, 0.268], [0.821, 0.35], [0.804, 0.509, 0.775, 0.509, 0.757, 0.35], [0.748, 0.268], [0.717, 0, 0.669, 0, 0.638, 0.268], [0.629, 0.35], [0.611, 0.51, 0.582, 0.51, 0.564, 0.35], [0.555, 0.268], [0.524, 0, 0.476, 0, 0.445, 0.268], [0.436, 0.35], [0.418, 0.51, 0.389, 0.51, 0.371, 0.35], [0.362, 0.268], [0.331, 0, 0.283, 0, 0.252, 0.268], [0.243, 0.35], [0.225, 0.509, 0.196, 0.509, 0.179, 0.35], [0.169, 0.268], [0.138, 0, 0.09, 0, 0.059, 0.268], [0.006, 0.722], [0.002, 0.759, 0, 0.816, 0.001, 0.87], [0.003, 0.924, 0.007, 0.968, 0.013, 0.984], [0.018, 1, 0.025, 0.986, 0.029, 0.948], [0.082, 0.493], [0.1, 0.334, 0.128, 0.334, 0.146, 0.493], [0.156, 0.575], [0.187, 0.843, 0.235, 0.843, 0.265, 0.575], [0.275, 0.493], [0.293, 0.334, 0.321, 0.334, 0.339, 0.493], [0.349, 0.575], [0.38, 0.843, 0.428, 0.843, 0.458, 0.575], [0.468, 0.493], [0.486, 0.334, 0.514, 0.334, 0.532, 0.493], [0.542, 0.575], [0.572, 0.843, 0.62, 0.843, 0.651, 0.575], [0.661, 0.493], [0.679, 0.334, 0.707, 0.334, 0.725, 0.493], [0.735, 0.575], [0.765, 0.843, 0.813, 0.843, 0.844, 0.575], [0.854, 0.493], [0.872, 0.334, 0.9, 0.334, 0.918, 0.493], [0.971, 0.948], [0.975, 0.986, 0.982, 1, 0.987, 0.984], [0.993, 0.968, 0.997, 0.924, 0.999, 0.87], [1, 0.816, 0.998, 0.759, 0.994, 0.722]], { fill: WHITE });
  path(s, 1.952, 4.55, 0.5, 0.042, [[0.982, 0.989], [0.989, 0.989, 0.995, 0.95, 0.998, 0.891], [1, 0.832, 0.998, 0.765, 0.993, 0.722], [0.94, 0.268], [0.91, 0, 0.862, 0, 0.831, 0.268], [0.821, 0.35], [0.803, 0.509, 0.775, 0.509, 0.757, 0.35], [0.747, 0.268], [0.717, 0, 0.669, 0, 0.638, 0.268], [0.628, 0.35], [0.61, 0.51, 0.582, 0.51, 0.564, 0.35], [0.555, 0.268], [0.524, 0, 0.476, 0, 0.445, 0.268], [0.436, 0.35], [0.418, 0.51, 0.389, 0.51, 0.371, 0.35], [0.362, 0.268], [0.331, 0, 0.283, 0, 0.252, 0.268], [0.243, 0.35], [0.225, 0.509, 0.196, 0.509, 0.178, 0.35], [0.169, 0.268], [0.138, 0, 0.09, 0, 0.059, 0.268], [0.006, 0.722], [0.002, 0.759, 0, 0.816, 0.001, 0.87], [0.003, 0.924, 0.007, 0.968, 0.013, 0.984], [0.018, 1, 0.025, 0.986, 0.029, 0.948], [0.082, 0.493], [0.1, 0.334, 0.128, 0.334, 0.146, 0.493], [0.156, 0.575], [0.187, 0.843, 0.235, 0.843, 0.265, 0.575], [0.275, 0.493], [0.293, 0.334, 0.321, 0.334, 0.339, 0.493], [0.349, 0.575], [0.379, 0.843, 0.427, 0.843, 0.458, 0.575], [0.468, 0.493], [0.486, 0.334, 0.514, 0.334, 0.532, 0.493], [0.542, 0.575], [0.572, 0.843, 0.62, 0.843, 0.651, 0.575], [0.661, 0.493], [0.679, 0.334, 0.707, 0.334, 0.725, 0.493], [0.734, 0.575], [0.765, 0.843, 0.813, 0.843, 0.844, 0.575], [0.853, 0.493], [0.871, 0.334, 0.9, 0.334, 0.918, 0.493], [0.971, 0.948], [0.974, 0.974, 0.978, 0.989, 0.982, 0.989]], { fill: WHITE });
  sh(s, 'roundRect', 0.882, 5.315, 0.723, 0.712, { rad: 0.086, shadow: true, fill: GOLD });
  path(s, 1.053, 5.391, 0.409, 0.483, [[0.116, 1], [0.282, 1], [0.323, 1, 0.36, 0.978, 0.381, 0.943], [0.402, 0.978, 0.44, 1, 0.481, 1], [0.647, 1], [0.711, 1, 0.763, 0.948, 0.763, 0.883], [0.763, 0.351], [0.763, 0.314, 0.745, 0.278, 0.714, 0.256], [0.657, 0.215], [0.644, 0.206, 0.629, 0.201, 0.613, 0.2], [0.613, 0.167], [0.63, 0.167], [0.648, 0.167, 0.663, 0.152, 0.663, 0.133], [0.763, 0.133], [0.827, 0.133, 0.879, 0.186, 0.879, 0.25], [0.879, 0.518], [0.846, 0.524, 0.817, 0.546, 0.804, 0.577], [0.791, 0.608, 0.794, 0.644, 0.812, 0.672], [0.812, 0.783], [0.812, 0.802, 0.827, 0.817, 0.846, 0.817], [0.945, 0.817], [0.963, 0.817, 0.978, 0.802, 0.978, 0.783], [0.978, 0.672], [0.997, 0.644, 1, 0.608, 0.987, 0.577], [0.973, 0.546, 0.945, 0.524, 0.912, 0.518], [0.912, 0.25], [0.912, 0.167, 0.845, 0.1, 0.763, 0.1], [0.663, 0.1], [0.663, 0.082, 0.648, 0.067, 0.63, 0.067], [0.58, 0.067], [0.58, 0.033], [0.597, 0.033], [0.606, 0.033, 0.613, 0.026, 0.613, 0.017], [0.613, 0.007, 0.606, 0, 0.597, 0], [0.531, 0], [0.521, 0, 0.514, 0.007, 0.514, 0.017], [0.514, 0.026, 0.521, 0.033, 0.531, 0.033], [0.547, 0.033], [0.547, 0.067], [0.497, 0.067], [0.479, 0.067, 0.464, 0.082, 0.464, 0.1], [0.298, 0.1], [0.298, 0.082, 0.284, 0.067, 0.265, 0.067], [0.133, 0.067], [0.114, 0.067, 0.099, 0.082, 0.099, 0.1], [0.066, 0.1], [0.066, 0.083], [0.066, 0.074, 0.059, 0.067, 0.05, 0.067], [0.041, 0.067, 0.033, 0.074, 0.033, 0.083], [0.033, 0.15], [0.033, 0.159, 0.041, 0.167, 0.05, 0.167], [0.059, 0.167, 0.066, 0.159, 0.066, 0.15], [0.066, 0.133], [0.099, 0.133], [0.099, 0.152, 0.114, 0.167, 0.133, 0.167], [0.149, 0.167], [0.149, 0.2], [0.134, 0.201, 0.118, 0.206, 0.106, 0.216], [0.049, 0.257], [0.018, 0.278, 0, 0.314, 0, 0.351], [0, 0.883], [0, 0.948, 0.052, 1, 0.116, 1]], { fill: WHITE });
  path(s, 1.399, 5.731, 0.041, 0.039, [[0, 1], [0, 0], [0.309, 0.225, 0.691, 0.225, 1, 0], [1, 1]], { fill: WHITE });
  path(s, 1.392, 5.657, 0.054, 0.064, [[1, 0.5], [1, 0.776, 0.776, 1, 0.5, 1], [0.224, 1, 0, 0.776, 0, 0.5], [0, 0.224, 0.224, 0, 0.5, 0], [0.776, 0, 1, 0.224, 1, 0.5]], { fill: WHITE });
  path(s, 1.067, 5.576, 0.136, 0.209, [[0, 0], [1, 0], [1, 1], [0, 1]], { fill: WHITE });
  path(s, 1.216, 5.576, 0.136, 0.209, [[1, 1], [0, 1], [0, 0], [1, 0]], { fill: WHITE });
  path(s, 1.216, 5.802, 0.136, 0.056, [[0.75, 1], [0.25, 1], [0.112, 1, 0, 0.68, 0, 0.286], [0, 0], [1, 0], [1, 0.286], [1, 0.68, 0.888, 1, 0.75, 1]], { fill: WHITE });
  path(s, 1.216, 5.504, 0.136, 0.056, [[0.895, 0.431], [0.96, 0.563, 0.999, 0.774, 1, 1], [0, 1], [0.001, 0.774, 0.04, 0.563, 0.104, 0.431], [0.277, 0.08], [0.302, 0.028, 0.333, 0, 0.364, 0], [0.636, 0], [0.667, 0, 0.698, 0.028, 0.723, 0.08]], { fill: WHITE });
  path(s, 1.277, 5.472, 0.014, 0.016, [[1, 1], [0, 1], [0, 0], [1, 0]], { fill: WHITE });
  path(s, 1.257, 5.439, 0.054, 0.016, [[0, 0], [1, 0], [1, 1], [0, 1]], { fill: WHITE });
  path(s, 1.155, 5.456, 0.109, 0.077, [[0.187, 0], [0.813, 0], [0.813, 0.116, 0.868, 0.21, 0.938, 0.21], [1, 0.21], [1, 0.423], [0.941, 0.428, 0.884, 0.461, 0.836, 0.519], [0.621, 0.777], [0.572, 0.837, 0.531, 0.913, 0.5, 1], [0.469, 0.912, 0.428, 0.837, 0.379, 0.777], [0.164, 0.519], [0.116, 0.461, 0.059, 0.428, 0, 0.423], [0, 0.21], [0.063, 0.21], [0.132, 0.21, 0.187, 0.116, 0.187, 0]], { fill: WHITE });
  path(s, 1.107, 5.439, 0.054, 0.016, [[0, 0], [1, 0], [1, 1], [0, 1]], { fill: WHITE });
  path(s, 1.128, 5.472, 0.014, 0.016, [[1, 0], [1, 1], [0, 1], [0, 0]], { fill: WHITE });
  path(s, 1.067, 5.504, 0.136, 0.056, [[0.104, 0.431], [0.277, 0.08], [0.302, 0.028, 0.333, 0, 0.364, 0], [0.636, 0], [0.667, 0, 0.698, 0.028, 0.723, 0.08], [0.895, 0.431], [0.96, 0.563, 0.999, 0.774, 1, 1], [0, 1], [0.001, 0.774, 0.04, 0.563, 0.104, 0.431]], { fill: WHITE });
  path(s, 1.067, 5.802, 0.136, 0.056, [[0, 0], [1, 0], [1, 0.286], [1, 0.68, 0.888, 1, 0.75, 1], [0.25, 1], [0.112, 1, 0, 0.68, 0, 0.286]], { fill: WHITE });
  path(s, 1.094, 5.609, 0.081, 0.064, [[0.833, 0], [0.167, 0], [0.075, 0, 0, 0.112, 0, 0.25], [0, 0.75], [0, 0.888, 0.075, 1, 0.167, 1], [0.833, 1], [0.925, 1, 1, 0.888, 1, 0.75], [1, 0.25], [1, 0.112, 0.925, 0, 0.833, 0]], { fill: WHITE });
  path(s, 1.107, 5.625, 0.054, 0.032, [[0, 1], [0, 0], [1, 0], [1, 1]], { fill: WHITE });
  path(s, 1.094, 5.705, 0.081, 0.016, [[0.917, 0], [0.083, 0], [0.037, 0, 0, 0.224, 0, 0.5], [0, 0.776, 0.037, 1, 0.083, 1], [0.917, 1], [0.963, 1, 1, 0.776, 1, 0.5], [1, 0.224, 0.963, 0, 0.917, 0]], { fill: WHITE });
  path(s, 1.114, 5.738, 0.041, 0.016, [[0.833, 0], [0.167, 0], [0.075, 0, 0, 0.224, 0, 0.5], [0, 0.776, 0.075, 1, 0.167, 1], [0.833, 1], [0.925, 1, 1, 0.776, 1, 0.5], [1, 0.224, 0.925, 0, 0.833, 0]], { fill: WHITE });
  path(s, 1.243, 5.609, 0.081, 0.064, [[0.167, 1], [0.833, 1], [0.925, 1, 1, 0.888, 1, 0.75], [1, 0.25], [1, 0.112, 0.925, 0, 0.833, 0], [0.167, 0], [0.075, 0, 0, 0.112, 0, 0.25], [0, 0.75], [0, 0.888, 0.075, 1, 0.167, 1]], { fill: WHITE });
  path(s, 1.257, 5.625, 0.054, 0.032, [[0, 0], [1, 0], [1, 1], [0, 1]], { fill: WHITE });
  path(s, 1.243, 5.705, 0.081, 0.016, [[0.083, 1], [0.917, 1], [0.963, 1, 1, 0.776, 1, 0.5], [1, 0.224, 0.963, 0, 0.917, 0], [0.083, 0], [0.037, 0, 0, 0.224, 0, 0.5], [0, 0.776, 0.037, 1, 0.083, 1]], { fill: WHITE });
  path(s, 1.263, 5.738, 0.041, 0.016, [[0.833, 0], [0.167, 0], [0.075, 0, 0, 0.224, 0, 0.5], [0, 0.776, 0.075, 1, 0.167, 1], [0.833, 1], [0.925, 1, 1, 0.776, 1, 0.5], [1, 0.224, 0.925, 0, 0.833, 0]], { fill: WHITE });
  path(s, 1.413, 5.665, 0.014, 0.032, [[0, 0.25], [0, 0.75], [0, 0.888, 0.224, 1, 0.5, 1], [0.776, 1, 1, 0.888, 1, 0.75], [1, 0.25], [1, 0.112, 0.776, 0, 0.5, 0], [0.224, 0, 0, 0.112, 0, 0.25]], { fill: WHITE });
  sh(s, 'roundRect', 0.878, 4.066, 0.723, 0.712, { rad: 0.086, shadow: true, fill: BLUE });
  path(s, 1.068, 4.203, 0.297, 0.317, [[0.058, 0.625], [0.077, 0.625], [0.077, 0.649], [0.077, 0.679, 0.09, 0.709, 0.112, 0.728], [0.288, 0.888], [0.309, 0.906, 0.335, 0.917, 0.362, 0.917], [0.393, 0.917], [0.407, 0.942, 0.445, 1, 0.5, 1], [0.555, 1, 0.593, 0.942, 0.607, 0.917], [0.638, 0.917], [0.665, 0.917, 0.691, 0.906, 0.712, 0.888], [0.888, 0.728], [0.91, 0.709, 0.923, 0.679, 0.923, 0.649], [0.923, 0.625], [0.942, 0.625], [0.974, 0.625, 1, 0.597, 1, 0.563], [1, 0.333], [1, 0.149, 0.862, 0, 0.692, 0], [0.308, 0], [0.138, 0, 0, 0.149, 0, 0.333], [0, 0.563], [0, 0.597, 0.026, 0.625, 0.058, 0.625]], { fill: WHITE });
  path(s, 1.08, 4.362, 0.011, 0.026, [[0.5, 1], [0.224, 1, 0, 0.888, 0, 0.75], [0, 0.25], [0, 0.112, 0.224, 0, 0.5, 0], [1, 0], [1, 1]], { fill: WHITE });
  path(s, 1.194, 4.428, 0.045, 0.079, [[0.5, 1], [0.277, 1, 0.081, 0.83, 0, 0.739], [0.012, 0.707, 0.023, 0.675, 0.035, 0.642], [0.127, 0.386, 0.266, 0, 0.5, 0], [0.734, 0, 0.873, 0.386, 0.965, 0.642], [0.977, 0.675, 0.988, 0.707, 1, 0.738], [0.937, 0.808, 0.741, 1, 0.5, 1]], { fill: WHITE });
  path(s, 1.154, 4.368, 0.126, 0.112, [[0.827, 1], [0.763, 1], [0.76, 0.985, 0.758, 0.971, 0.755, 0.955], [0.708, 0.701, 0.656, 0.412, 0.5, 0.412], [0.344, 0.412, 0.292, 0.701, 0.245, 0.955], [0.242, 0.971, 0.24, 0.985, 0.237, 1], [0.173, 1], [0.131, 1, 0.09, 0.981, 0.057, 0.945], [0, 0.884], [0, 0.529], [0, 0.237, 0.183, 0, 0.409, 0], [0.591, 0], [0.817, 0, 1, 0.237, 1, 0.529], [1, 0.884], [0.943, 0.945], [0.911, 0.981, 0.869, 1, 0.827, 1]], { fill: WHITE });
  path(s, 1.103, 4.309, 0.228, 0.148, [[1, 0.677], [1, 0.716, 0.99, 0.754, 0.973, 0.779], [0.825, 1], [0.825, 0.805], [0.825, 0.533, 0.702, 0.313, 0.55, 0.313], [0.45, 0.313], [0.298, 0.313, 0.175, 0.533, 0.175, 0.805], [0.175, 1], [0.027, 0.78], [0.01, 0.754, 0, 0.716, 0, 0.677], [0, 0.224], [0, 0.1, 0.056, 0, 0.125, 0], [0.875, 0], [0.944, 0, 1, 0.1, 1, 0.224]], { fill: WHITE });
  path(s, 1.342, 4.362, 0.011, 0.026, [[1, 0.75], [1, 0.888, 0.776, 1, 0.5, 1], [0, 1], [0, 0], [0.5, 0], [0.776, 0, 1, 0.112, 1, 0.25]], { fill: WHITE });
  path(s, 1.08, 4.216, 0.274, 0.133, [[0.292, 0], [0.708, 0], [0.869, 0, 1, 0.311, 1, 0.694], [1, 1], [0.993, 0.994, 0.986, 0.991, 0.979, 0.991], [0.958, 0.991], [0.958, 0.941], [0.958, 0.75, 0.893, 0.595, 0.812, 0.595], [0.187, 0.595], [0.107, 0.595, 0.042, 0.75, 0.042, 0.941], [0.042, 0.991], [0.021, 0.991], [0.014, 0.991, 0.007, 0.994, 0, 1], [0, 0.694], [0, 0.311, 0.131, 0, 0.292, 0]], { fill: WHITE });
  path(s, 1.24, 4.203, 0.171, 0.397, [[0.9, 0], [0.845, 0, 0.8, 0.022, 0.8, 0.05], [0.8, 0.748], [0.8, 0.812, 0.695, 0.864, 0.567, 0.864], [0.507, 0.864], [0.516, 0.852, 0.524, 0.839, 0.53, 0.826], [0.539, 0.809, 0.528, 0.791, 0.503, 0.778], [0.477, 0.766, 0.44, 0.761, 0.406, 0.766], [0.371, 0.771, 0.345, 0.785, 0.336, 0.802], [0.289, 0.897, 0.193, 0.897, 0.1, 0.897], [0.045, 0.897, 0, 0.92, 0, 0.947], [0, 0.975, 0.045, 0.997, 0.1, 0.997], [0.192, 1, 0.283, 0.988, 0.361, 0.964], [0.567, 0.964], [0.806, 0.964, 1, 0.867, 1, 0.748], [1, 0.05], [1, 0.022, 0.955, 0, 0.9, 0]], { fill: WHITE });
  path(s, 1.388, 4.216, 0.011, 0.046, [[0.5, 0], [0.776, 0, 1, 0.064, 1, 0.143], [1, 1], [0, 1], [0, 0.143], [0, 0.064, 0.224, 0, 0.5, 0]], { fill: WHITE });
  path(s, 1.388, 4.276, 0.011, 0.053, [[1, 0], [1, 1], [0, 1], [0, 0]], { fill: WHITE });
  path(s, 1.251, 4.519, 0.069, 0.067, [[0, 0.901], [0, 0.846, 0.037, 0.802, 0.083, 0.802], [0.321, 0.802, 0.68, 0.802, 0.83, 0.083], [0.842, 0.031, 0.886, 0, 0.929, 0.013], [0.973, 0.026, 1, 0.079, 0.99, 0.131], [0.809, 1, 0.337, 1, 0.083, 1], [0.037, 1, 0, 0.956, 0, 0.901]], { fill: WHITE });
  path(s, 1.314, 4.342, 0.086, 0.231, [[0.268, 1], [0, 1], [0.033, 0.982, 0.063, 0.963, 0.088, 0.943], [0.268, 0.943], [0.598, 0.943, 0.866, 0.828, 0.867, 0.686], [0.867, 0], [1, 0], [1, 0.686], [1, 0.859, 0.672, 1, 0.268, 1]], { fill: WHITE });
  path(s, 1.114, 4.322, 0.029, 0.033, [[0.8, 0], [0.6, 0], [0.269, 0, 0, 0.269, 0, 0.6], [0, 0.8], [0, 0.91, 0.09, 1, 0.2, 1], [0.31, 1, 0.4, 0.91, 0.4, 0.8], [0.4, 0.6], [0.4, 0.49, 0.49, 0.4, 0.6, 0.4], [0.8, 0.4], [0.91, 0.4, 1, 0.31, 1, 0.2], [1, 0.09, 0.91, 0, 0.8, 0]], { fill: WHITE });
  path(s, 1.154, 4.322, 0.017, 0.013, [[0.667, 0], [0.333, 0], [0.149, 0, 0, 0.224, 0, 0.5], [0, 0.776, 0.149, 1, 0.333, 1], [0.667, 1], [0.851, 1, 1, 0.776, 1, 0.5], [1, 0.224, 0.851, 0, 0.667, 0]], { fill: WHITE });
  sh(s, 'ellipse', 1.194, 4.54, 0.034, 0.04, { fill: WHITE });
  sh(s, 'ellipse', 1.205, 4.553, 0.011, 0.013, { fill: WHITE });
  sh(s, 'ellipse', 1.137, 4.507, 0.046, 0.053, { fill: WHITE });
  sh(s, 'ellipse', 1.148, 4.52, 0.023, 0.026, { fill: WHITE });
  sh(s, 'ellipse', 1.171, 4.573, 0.011, 0.013, { fill: WHITE });
  sh(s, 'roundRect', 1.86, 5.315, 0.723, 0.712, { rad: 0.086, shadow: true, fill: GOLD });
  path(s, 2.006, 5.527, 0.408, 0.333, [[0.05, 0.878], [0.05, 0.898], [0.05, 0.932, 0.072, 0.959, 0.1, 0.959], [0.17, 0.959], [0.177, 0.984, 0.196, 1, 0.217, 1], [0.783, 1], [0.804, 1, 0.823, 0.984, 0.83, 0.959], [0.9, 0.959], [0.928, 0.959, 0.95, 0.932, 0.95, 0.898], [0.95, 0.878], [0.977, 0.879, 0.999, 0.854, 1, 0.821], [1, 0.281], [0.999, 0.248, 0.977, 0.223, 0.95, 0.224], [0.95, 0.204], [0.95, 0.17, 0.928, 0.143, 0.9, 0.143], [0.83, 0.143], [0.823, 0.118, 0.804, 0.102, 0.783, 0.102], [0.669, 0.102], [0.663, 0.102, 0.658, 0.098, 0.655, 0.092], [0.624, 0.03], [0.615, 0.011, 0.599, 0, 0.581, 0], [0.419, 0], [0.401, 0, 0.385, 0.011, 0.376, 0.03], [0.345, 0.092], [0.342, 0.098, 0.337, 0.102, 0.331, 0.102], [0.217, 0.102], [0.196, 0.102, 0.177, 0.118, 0.17, 0.143], [0.1, 0.143], [0.072, 0.143, 0.05, 0.17, 0.05, 0.204], [0.05, 0.224], [0.023, 0.223, 0.001, 0.248, 0, 0.281], [0, 0.821], [0.001, 0.854, 0.023, 0.879, 0.05, 0.878]], { fill: WHITE });
  path(s, 2.373, 5.615, 0.028, 0.191, [[1, 0.029], [1, 0.971], [0.98, 0.988, 0.872, 1, 0.753, 0.998], [0.365, 0.998], [0.26, 0.999, 0.162, 0.99, 0.123, 0.976], [0.042, 0.934], [0, 0.912, 0, 0.889, 0.042, 0.867], [0.132, 0.819], [0.198, 0.785, 0.198, 0.749, 0.132, 0.714], [0.042, 0.667], [0, 0.645, 0, 0.622, 0.042, 0.6], [0.132, 0.553], [0.198, 0.518, 0.198, 0.482, 0.132, 0.447], [0.042, 0.4], [0, 0.378, 0, 0.355, 0.042, 0.333], [0.132, 0.286], [0.198, 0.251, 0.198, 0.215, 0.132, 0.181], [0.042, 0.133], [0, 0.111, 0, 0.088, 0.042, 0.066], [0.123, 0.024], [0.162, 0.01, 0.26, 0.001, 0.365, 0.002], [0.753, 0.002], [0.872, 0, 0.981, 0.012, 1, 0.029]], { fill: WHITE });
  path(s, 2.347, 5.588, 0.034, 0.245, [[1, 0.028], [1, 0.057], [0.768, 0.06, 0.573, 0.082, 0.505, 0.113], [0.44, 0.146], [0.387, 0.173, 0.387, 0.202, 0.44, 0.229], [0.513, 0.265], [0.547, 0.283, 0.547, 0.301, 0.513, 0.318], [0.44, 0.355], [0.387, 0.382, 0.387, 0.41, 0.44, 0.437], [0.513, 0.474], [0.547, 0.491, 0.547, 0.509, 0.513, 0.526], [0.44, 0.563], [0.387, 0.59, 0.387, 0.618, 0.44, 0.645], [0.513, 0.682], [0.547, 0.699, 0.547, 0.718, 0.513, 0.735], [0.44, 0.771], [0.387, 0.798, 0.387, 0.827, 0.44, 0.854], [0.505, 0.887], [0.573, 0.918, 0.768, 0.94, 1, 0.943], [1, 0.972], [1, 0.988, 0.91, 1, 0.8, 1], [0, 1], [0, 0], [0.8, 0], [0.91, 0, 1, 0.012, 1, 0.028]], { fill: WHITE });
  path(s, 2.088, 5.541, 0.245, 0.306, [[0, 0.133], [0, 0.121, 0.012, 0.111, 0.028, 0.111], [0.219, 0.111], [0.248, 0.111, 0.275, 0.099, 0.29, 0.079], [0.341, 0.011], [0.346, 0.004, 0.355, 0, 0.365, 0], [0.635, 0], [0.645, 0, 0.654, 0.004, 0.659, 0.011], [0.71, 0.079], [0.725, 0.099, 0.752, 0.111, 0.781, 0.111], [0.972, 0.111], [0.988, 0.111, 1, 0.121, 1, 0.133], [1, 0.978], [1, 0.99, 0.988, 1, 0.972, 1], [0.028, 1], [0.012, 1, 0, 0.99, 0, 0.978]], { fill: WHITE });
  path(s, 2.04, 5.588, 0.034, 0.245, [[0.2, 0], [1, 0], [1, 1], [0.2, 1], [0.09, 1, 0, 0.988, 0, 0.972], [0, 0.943], [0.232, 0.94, 0.427, 0.918, 0.495, 0.887], [0.56, 0.854], [0.613, 0.827, 0.613, 0.798, 0.56, 0.771], [0.487, 0.735], [0.453, 0.717, 0.453, 0.699, 0.487, 0.682], [0.56, 0.645], [0.613, 0.618, 0.613, 0.59, 0.56, 0.563], [0.487, 0.526], [0.453, 0.509, 0.453, 0.491, 0.487, 0.474], [0.56, 0.437], [0.613, 0.41, 0.613, 0.382, 0.56, 0.355], [0.487, 0.318], [0.453, 0.301, 0.453, 0.282, 0.487, 0.265], [0.56, 0.229], [0.613, 0.202, 0.613, 0.173, 0.56, 0.146], [0.495, 0.113], [0.427, 0.082, 0.232, 0.06, 0, 0.057], [0, 0.028], [0, 0.012, 0.09, 0, 0.2, 0]], { fill: WHITE });
  path(s, 2.02, 5.615, 0.028, 0.191, [[0, 0.029], [0.02, 0.012, 0.128, 0, 0.247, 0.002], [0.635, 0.002], [0.74, 0.001, 0.838, 0.01, 0.877, 0.024], [0.958, 0.066], [1, 0.088, 1, 0.111, 0.958, 0.133], [0.868, 0.181], [0.802, 0.215, 0.802, 0.251, 0.868, 0.286], [0.958, 0.333], [1, 0.355, 1, 0.378, 0.958, 0.4], [0.868, 0.448], [0.802, 0.482, 0.802, 0.518, 0.868, 0.553], [0.958, 0.6], [1, 0.622, 1, 0.645, 0.958, 0.667], [0.868, 0.714], [0.802, 0.749, 0.802, 0.785, 0.868, 0.82], [0.958, 0.867], [1, 0.889, 1, 0.912, 0.958, 0.934], [0.878, 0.976], [0.838, 0.991, 0.74, 0.999, 0.635, 0.998], [0.247, 0.998], [0.128, 1, 0.02, 0.988, 0, 0.971]], { fill: WHITE });
  path(s, 2.183, 5.554, 0.054, 0.014, [[0.125, 1], [0.875, 1], [0.944, 1, 1, 0.776, 1, 0.5], [1, 0.224, 0.944, 0, 0.875, 0], [0.125, 0], [0.056, 0, 0, 0.224, 0, 0.5], [0, 0.776, 0.056, 1, 0.125, 1]], { fill: WHITE });
  sh(s, 'ellipse', 2.102, 5.615, 0.218, 0.218, { fill: WHITE });
  sh(s, 'ellipse', 2.115, 5.629, 0.19, 0.19, { fill: WHITE });
  sh(s, 'ellipse', 2.129, 5.643, 0.163, 0.163, { fill: WHITE });
  sh(s, 'ellipse', 2.142, 5.656, 0.136, 0.136, { fill: WHITE });
  path(s, 2.265, 5.588, 0.054, 0.041, [[0.375, 1], [0.625, 1], [0.832, 1, 1, 0.776, 1, 0.5], [1, 0.224, 0.832, 0, 0.625, 0], [0.375, 0], [0.168, 0, 0, 0.224, 0, 0.5], [0, 0.776, 0.168, 1, 0.375, 1]], { fill: WHITE });
  path(s, 2.279, 5.602, 0.027, 0.014, [[0.25, 0], [0.75, 0], [0.888, 0, 1, 0.224, 1, 0.5], [1, 0.776, 0.888, 1, 0.75, 1], [0.25, 1], [0.112, 1, 0, 0.776, 0, 0.5], [0, 0.224, 0.112, 0, 0.25, 0]], { fill: WHITE });
  path(s, 2.102, 5.595, 0.014, 0.02, [[0.5, 1], [0.776, 1, 1, 0.851, 1, 0.667], [1, 0.333], [1, 0.149, 0.776, 0, 0.5, 0], [0.224, 0, 0, 0.149, 0, 0.333], [0, 0.667], [0, 0.851, 0.224, 1, 0.5, 1]], { fill: WHITE });
  path(s, 2.122, 5.595, 0.014, 0.02, [[0.5, 1], [0.776, 1, 1, 0.851, 1, 0.667], [1, 0.333], [1, 0.149, 0.776, 0, 0.5, 0], [0.224, 0, 0, 0.149, 0, 0.333], [0, 0.667], [0, 0.851, 0.224, 1, 0.5, 1]], { fill: WHITE });
  sh(s, 'ellipse', 2.163, 5.677, 0.054, 0.054, { fill: WHITE });
  path(s, 2.204, 5.452, 0.014, 0.061, [[0.5, 1], [0.776, 1, 1, 0.95, 1, 0.889], [1, 0.111], [1, 0.05, 0.776, 0, 0.5, 0], [0.224, 0, 0, 0.05, 0, 0.111], [0, 0.889], [0, 0.95, 0.224, 1, 0.5, 1]], { fill: WHITE });
  path(s, 2.115, 5.492, 0.035, 0.035, [[0.653, 0.927], [0.729, 1, 0.85, 0.999, 0.924, 0.924], [0.999, 0.85, 1, 0.729, 0.927, 0.653], [0.347, 0.073], [0.271, 0, 0.15, 0.001, 0.076, 0.076], [0.001, 0.15, 0, 0.271, 0.073, 0.347]], { fill: WHITE });
  path(s, 2.271, 5.492, 0.035, 0.034, [[0.209, 1], [0.261, 1, 0.31, 0.979, 0.347, 0.942], [0.929, 0.351], [0.98, 0.301, 1, 0.228, 0.982, 0.159], [0.964, 0.09, 0.911, 0.036, 0.844, 0.018], [0.776, 0, 0.703, 0.021, 0.655, 0.072], [0.072, 0.663], [0.017, 0.72, 0, 0.805, 0.03, 0.878], [0.06, 0.952, 0.131, 1, 0.209, 1]], { fill: WHITE });
  path(s, 2.155, 5.465, 0.029, 0.049, [[0.517, 0.895], [0.565, 0.966, 0.701, 1, 0.821, 0.972], [0.942, 0.943, 1, 0.863, 0.952, 0.793], [0.483, 0.105], [0.435, 0.034, 0.299, 0, 0.179, 0.028], [0.058, 0.057, 0, 0.137, 0.048, 0.207]], { fill: WHITE });
  path(s, 2.238, 5.465, 0.028, 0.049, [[0.158, 0.986], [0.217, 1, 0.283, 1, 0.342, 0.985], [0.4, 0.97, 0.446, 0.943, 0.47, 0.908], [0.951, 0.21], [1, 0.139, 0.94, 0.057, 0.817, 0.029], [0.694, 0, 0.554, 0.035, 0.504, 0.106], [0.024, 0.805], [0, 0.839, 0.001, 0.877, 0.026, 0.912], [0.051, 0.946, 0.099, 0.973, 0.158, 0.986]], { fill: WHITE });
  tx(s, 0.795, 1.765, 4.405, 0.64, 'Expert Service', { sz: 32, bold: true, color: BLUE });
  tx(s, 0.795, 1.429, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', color: GOLD });
  tx(s, 0.795, 2.43, 5.025, 1.284, cut(LOREM, 211), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 2.756, 4.03, 2.573, 0.303, 'Fisrt Diving Service', { font: 'SEMI', color: BLUE });
  tx(s, 2.756, 4.333, 3.143, 0.678, cut(DEV, 65), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 2.756, 5.21, 2.573, 0.303, 'Second Diving Service', { font: 'SEMI', color: BLUE });
  tx(s, 2.756, 5.513, 3.143, 0.678, cut(DEV, 65), { align: 'justify', lh: 1.5, color: GRAY });
  photo(s, 6.667, 0.55, 2.146, 2.05, 'roundRect', { flipH: true, rad: 0.258, fill: 'CBCBCB' });
  photo(s, 6.667, 2.779, 2.146, 2.05, 'roundRect', { flipH: true, rad: 0.258, fill: 'CBCBCB' });
  photo(s, 6.667, 5.007, 2.146, 2.05, 'roundRect', { flipH: true, rad: 0.258, fill: 'CBCBCB' });
  photo(s, 9.171, 0, 2.857, 7.057, 'round2SameRect', { flipH: true, fill: 'CBCBCB' });
}

// ---------------------------------------------------------------- 20. Breaktime
function slide20(s) {
  grad(s, rectPts(0, 0, 13.333, 7.5), { angle: 120 });
  sh(s, 'rect', 0, 0, 2.868, 7.5, { fill: GOLD });

  tx(s, 5.54, 3.302, 6.225, 1.447, 'Breaktime', { sz: 80, bold: true, align: 'center', color: WHITE });
  tx(s, 7.026, 4.581, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', align: 'center', color: GOLD });
  sh(s, 'roundRect', 0.535, 0.45, 0.312, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 0.603, 0.544, 0.244, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 0.603, 0.637, 0.244, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  tx(s, 0.847, 0.489, 1.426, 0.269, 'Trends', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 2.266, 0.489, 1.426, 0.269, 'Diving', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 3.608, 0.489, 1.426, 0.269, 'Ocean', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 4.91, 0.489, 1.426, 0.269, 'Contact', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 11.329, 1.799, 3.035, 0.337, [['Best Day ', GOLD], ['for Diving in 2021', WHITE]], { sz: 14, bold: true, rot: 90 });
  tx(s, 3.211, 6.75, 1.982, 0.572, [['Published by: ', WHITE], ['yumnacreative', GOLD]], { sz: 14, bold: true });
  photo(s, 0.715, 1.738, 4.306, 4.305, 'ellipse', { flipH: true, fill: 'CBCBCB', line: WHITE, lw: 6 });
}

// ---------------------------------------------------------------- 21. Together We Can Save Our Ocean
function slide21(s) {
  grad(s, rectPts(0, 0, 13.333, 7.5), { angle: 120 });
  sh(s, 'rect', 0, 0, 13.333, 0.761, { fill: GOLD });

  tx(s, 6.464, 3.639, 5.822, 2.193, LOREM, { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 6.464, 3.383, 2.361, 0.303, 'INSERT TEXT HERE', { font: 'SEMI', color: GOLD });
  tx(s, 6.464, 2.34, 4.655, 1.043, 'Together We Can Save Our Ocean', { sz: 28, bold: true, color: WHITE });
  tx(s, 6.464, 5.874, 5.822, 0.981, cut(LOREM, 173), { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 11.302, 5.814, 3.035, 0.337, 'Best Day for Diving in 2021', { sz: 14, bold: true, rot: 90, color: GOLD });
  photo(s, -1.388, -1.108, 4.274, 4.405, 'diamond', { fill: 'CBCBCB' });
  photo(s, 1.047, 1.05, 4.274, 4.405, 'diamond', { fill: 'CBCBCB' });
  photo(s, 3.333, -1.284, 4.274, 4.405, 'diamond', { fill: 'CBCBCB' });
  photo(s, -1.239, 3.383, 4.274, 4.405, 'diamond', { fill: 'CBCBCB' });
}

// ---------------------------------------------------------------- 22. Expert Portfolio
function slide22(s) {
  sh(s, 'rect', 0.404, 0.457, 3.012, 3.699, { shadow: true, fill: WHITE });
  grad(s, [[13.333, 0], [6.248, 0], [13.333, 7.5]], { angle: 330 });
  sh(s, 'rect', 3.588, 0.457, 3.012, 3.699, { shadow: true, fill: WHITE });
  sh(s, 'rect', 6.772, 0.457, 3.012, 3.699, { shadow: true, fill: WHITE });
  sh(s, 'rect', 9.956, 0.457, 3.012, 3.699, { shadow: true, fill: WHITE });

  tx(s, 1.641, 4.997, 4.405, 0.64, 'Expert Portfolio', { sz: 32, bold: true, color: BLUE });
  tx(s, 1.641, 4.66, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', color: GOLD });
  tx(s, 1.641, 5.662, 5.025, 1.284, cut(LOREM, 211), { align: 'justify', lh: 1.5, color: GRAY });
  photo(s, 0.577, 0.656, 2.668, 3.301, 'rect', { fill: 'CBCBCB' });
  photo(s, 3.76, 0.656, 2.668, 3.301, 'rect', { fill: 'CBCBCB' });
  photo(s, 6.944, 0.656, 2.668, 3.301, 'rect', { fill: 'CBCBCB' });
  photo(s, 10.128, 0.656, 2.668, 3.301, 'rect', { fill: 'CBCBCB' });
}

// ---------------------------------------------------------------- 23. Best Galery
function slide23(s) {
  grad(s, rectPts(0, 0, 13.333, 7.5), { angle: 120 });
  sh(s, 'rect', 12.471, 0, 0.862, 7.5, { fill: GOLD });

  tx(s, 0.908, 1.967, 4.405, 0.64, 'Best Galery', { sz: 32, bold: true, color: WHITE });
  tx(s, 0.908, 1.631, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', color: GOLD });
  tx(s, 0.908, 2.632, 4.93, 0.981, cut(LOREM, 132), { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 0.908, 3.638, 4.93, 2.799, LOREM, { align: 'justify', lh: 1.5, color: SILVER });
  sh(s, 'roundRect', 0.535, 0.45, 0.312, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 0.603, 0.544, 0.244, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 0.603, 0.637, 0.244, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  tx(s, 0.847, 0.489, 1.426, 0.269, 'Trends', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 2.266, 0.489, 1.426, 0.269, 'Diving', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 3.608, 0.489, 1.426, 0.269, 'Ocean', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 4.91, 0.489, 1.426, 0.269, 'Contact', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 9.228, 6.867, 3.035, 0.337, [['Best Day ', GOLD], ['for Diving in 2021', WHITE]], { sz: 14, bold: true });
  tx(s, 0.603, 6.75, 1.982, 0.572, [['Published by: ', WHITE], ['yumnacreative', GOLD]], { sz: 14, bold: true });
  photo(s, 8.439, 0, 4.894, 4.145, 'rect', { fill: 'CBCBCB' });
  photo(s, 6.667, 1.13, 3.014, 3.014, 'ellipse', { shadow: true, fill: 'CBCBCB', line: WHITE, lw: 6 });
  photo(s, 9.681, 3.116, 3.014, 3.014, 'ellipse', { shadow: true, fill: 'CBCBCB', line: WHITE, lw: 6 });
}

// ---------------------------------------------------------------- 24. Best Events
function slide24(s) {
  sh(s, 'rect', 0, 0, 13.333, 7.5, { fill: NAVY });
  sh(s, 'rect', 0, 0, 1.257, 7.5, { fill: GOLD });

  photo(s, 1.033, 1.186, 3.838, 3.956, 'diamond', { shadow: true, fill: 'CBCBCB', line: WHITE, lw: 6 });
  photo(s, 3.219, 3.386, 3.838, 3.956, 'diamond', { shadow: true, fill: 'CBCBCB', line: WHITE, lw: 6 });
  photo(s, -1.124, -1.057, 3.838, 3.956, 'diamond', { shadow: true, fill: 'CBCBCB', line: WHITE, lw: 6 });
  photo(s, 5.376, 5.522, 3.838, 3.956, 'diamond', { shadow: true, fill: 'CBCBCB', line: WHITE, lw: 6 });
  photo(s, 7.534, 3.386, 3.838, 3.956, 'diamond', { shadow: true, fill: 'CBCBCB', line: WHITE, lw: 6 });
  photo(s, 9.615, 5.566, 3.838, 3.956, 'diamond', { shadow: true, fill: 'CBCBCB', line: WHITE, lw: 6 });
  tx(s, 5.705, 0.934, 4.405, 0.64, 'Best Events', { sz: 32, bold: true, color: WHITE });
  tx(s, 5.705, 0.597, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', color: GOLD });
  tx(s, 5.705, 1.615, 6.889, 1.284, cut(LOREM, 295), { align: 'justify', lh: 1.5, color: SILVER });
  sh(s, 'ellipse', 2.213, 4.534, 1.975, 1.975, { shadow: true, fill: BLUE, line: WHITE, lw: 6 });
  tx(s, 2.286, 4.932, 1.902, 1.134, [['1', WHITE], ['st', WHITE, { sup: true }], [' ', WHITE, { br: true }], ['Amazing Diving', WHITE, { sz: 20 }]], { sz: 28, bold: true, align: 'center', lh: 0.9 });
}

// ---------------------------------------------------------------- 25. Best Events
function slide25(s) {
  grad(s, rectPts(0, 0, 6.771, 7.5), { angle: 120 });
  sh(s, 'rect', 12.076, 0, 1.257, 7.5, { fill: GOLD });
  sh(s, 'rect', 7.273, 0.335, 5.332, 3.301, { shadow: true, fill: WHITE });
  sh(s, 'rect', 7.273, 3.813, 5.332, 3.301, { shadow: true, fill: WHITE });

  tx(s, 0.661, 2.431, 4.405, 0.64, 'Best Events', { sz: 32, bold: true, color: WHITE });
  tx(s, 0.661, 2.095, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', color: GOLD });
  tx(s, 0.661, 3.096, 5.427, 0.678, cut(LOREM, 124), { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 0.661, 3.848, 5.427, 2.193, cut(LOREM, 432), { align: 'justify', lh: 1.5, color: SILVER });
  sh(s, 'roundRect', 0.535, 0.45, 0.312, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 0.603, 0.544, 0.244, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 0.603, 0.637, 0.244, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  tx(s, 0.847, 0.489, 1.426, 0.269, 'Trends', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 2.266, 0.489, 1.426, 0.269, 'Diving', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 3.608, 0.489, 1.426, 0.269, 'Ocean', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 4.91, 0.489, 1.426, 0.269, 'Contact', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  photo(s, 7.464, 0.512, 4.951, 2.946, 'rect', { fill: 'CBCBCB' });
  photo(s, 7.464, 3.991, 4.951, 2.946, 'rect', { fill: 'CBCBCB' });
}

// ---------------------------------------------------------------- 26. Amazing Events 2021
function slide26(s) {
  grad(s, rectPts(0, 0, 13.333, 7.5), { angle: 120 });
  sh(s, 'rect', 0, 2.259, 13.333, 2.259, { fill: GOLD });
  sh(s, 'rect', 1.347, 1.4, 3.012, 3.699, { shadow: true, fill: WHITE });

  tx(s, 3.471, 0.74, 6.181, 0.64, 'Amazing Events 2021', { sz: 32, bold: true, align: 'center', color: WHITE });
  tx(s, 4.935, 0.403, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', align: 'center', color: GOLD });
  tx(s, 8.938, 7.082, 3.035, 0.337, [['Best Day ', GOLD], ['for Diving in 2021', WHITE]], { sz: 14, bold: true });
  photo(s, 1.562, 1.599, 2.668, 3.301, 'rect', { fill: 'CBCBCB' });
  photo(s, 4.634, 1.599, 4.023, 1.815, 'rect', { fill: 'CBCBCB' });
  photo(s, 4.634, 3.556, 4.023, 1.543, 'rect', { fill: 'CBCBCB' });
  photo(s, 8.932, 1.599, 2.882, 5.33, 'rect', { fill: 'CBCBCB' });
  photo(s, 6.843, 5.241, 1.814, 1.688, 'rect', { fill: 'CBCBCB' });
  photo(s, 4.634, 5.241, 2.033, 1.688, 'rect', { fill: 'CBCBCB' });
  photo(s, 1.39, 5.241, 3.012, 1.688, 'rect', { fill: 'CBCBCB' });
}

// ---------------------------------------------------------------- 27. Best Mockup
function slide27(s) {
  grad(s, rectPts(0, 0, 13.333, 7.5), { angle: 120 });
  photo(s, -0.232, -1.006, 2.77, 5.162, 'rect', { rot: 38.044, fill: 'D6D6D6' });
  photo(s, 0.464, 3.363, 2.77, 5.162, 'rect', { rot: 38.044, fill: 'D6D6D6' });
  photo(s, 3.811, -1.006, 2.77, 5.162, 'rect', { rot: 38.044, fill: 'D6D6D6' });
  photo(s, 4.507, 3.363, 2.77, 5.162, 'rect', { rot: 38.044, fill: 'D6D6D6' });

  tx(s, 8.543, 1.081, 4.405, 0.64, 'Best Mockup', { sz: 32, bold: true, color: GOLD });
  tx(s, 8.543, 1.745, 4.23, 0.981, cut(LOREM, 124), { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 8.543, 2.726, 4.23, 3.102, cut(LOREM, 432), { align: 'justify', lh: 1.5, color: SILVER });
  sh(s, 'roundRect', 8.543, 6.046, 3.774, 0.966, { rad: 0.117, shadow: true, fill: GOLD });
  tx(s, 8.93, 6.295, 2.109, 0.303, 'Insert your text', { font: 'SEMI', color: WHITE });
  tx(s, 11.266, 6.23, 0.664, 0.375, '56%', { font: 'SEMI', align: 'right', lh: 1.5, color: WHITE });
  sh(s, 'roundRect', 9.024, 6.661, 2.8, 0.083, { rad: 0.042, fill: 'E7E6E6' });
  sh(s, 'roundRect', 9.024, 6.661, 1.862, 0.083, { rad: 0.042, fill: BLUE });
  photo(s, -0.091, -0.523, 2.294, 3.984, 'rect', { rot: 37.005, fill: 'CBCBCB' });
  photo(s, 3.952, -0.523, 2.294, 3.984, 'rect', { rot: 37.005, fill: 'CBCBCB' });
  photo(s, 4.647, 3.846, 2.294, 3.984, 'rect', { rot: 37.005, fill: 'CBCBCB' });
  photo(s, 0.604, 3.846, 2.294, 3.984, 'rect', { rot: 37.005, fill: 'CBCBCB' });
}

// ---------------------------------------------------------------- 28. Best Desktop
function slide28(s) {
  grad(s, rectPts(5.674, 0, 6.771, 7.5), { angle: 120 });
  sh(s, 'round2SameRect', 0.495, 1.627, 6.754, 0.437, { fill: SILVER });
  sh(s, 'ellipse', 0.734, 1.806, 0.098, 0.098, { fill: 'FF0000' });
  sh(s, 'ellipse', 0.888, 1.806, 0.098, 0.098, { fill: GOLD });
  sh(s, 'ellipse', 1.042, 1.806, 0.098, 0.098, { fill: '92D050' });
  tx(s, 1.425, 1.745, 3.168, 0.219, 'http://www.website.com', { sz: 6, align: 'center', valign: 'middle', shape: 'roundRect', rad: 0.036, fill: WHITE, color: ASH });

  tx(s, 7.757, 1.707, 4.405, 0.64, 'Best Desktop', { sz: 32, bold: true, color: WHITE });
  tx(s, 7.757, 1.37, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', color: GOLD });
  tx(s, 7.757, 2.371, 4.291, 0.678, cut(LOREM, 90), { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 7.757, 3.123, 4.291, 3.102, cut(LOREM, 432), { align: 'justify', lh: 1.5, color: SILVER });
  sh(s, 'roundRect', 0.535, 0.45, 0.312, 0.051, { rad: 0.009, shadow: true, fill: GOLD });
  sh(s, 'roundRect', 0.603, 0.544, 0.244, 0.051, { rad: 0.009, shadow: true, fill: GOLD });
  sh(s, 'roundRect', 0.603, 0.637, 0.244, 0.051, { rad: 0.009, shadow: true, fill: GOLD });
  tx(s, 0.847, 0.489, 1.426, 0.269, 'Trends', { sz: 10, font: 'MED', align: 'center', color: GOLD });
  tx(s, 2.266, 0.489, 1.426, 0.269, 'Diving', { sz: 10, font: 'MED', align: 'center', color: GOLD });
  tx(s, 3.608, 0.489, 1.426, 0.269, 'Ocean', { sz: 10, font: 'MED', align: 'center', color: GOLD });
  tx(s, 0.603, 6.75, 1.982, 0.572, [['Published by: ', BLUE], ['yumnacreative', GOLD]], { sz: 14, bold: true });
  photo(s, 0.495, 2.065, 6.754, 3.617, 'rect', { shadow: true, fill: 'CBCBCB' });
}

// ---------------------------------------------------------------- 29. Best Watch Mockup
function slide29(s) {
  sh(s, 'rect', 0, 0, 1.257, 7.5, { fill: GOLD });
  sh(s, 'roundRect', 0.698, 0.635, 12.143, 6.23, { fill: NAVY, shadow: true, rad: 0.283 });
  grad(s, roundPts(0.698, 0.635, 12.143, 6.23, 0.283));
  photo(s, 1.177, 1.129, 4.259, 5.74, 'rect', { fill: '909090' });

  tx(s, 6.011, 1.707, 5.132, 0.64, 'Best Watch Mockup', { sz: 32, bold: true, color: WHITE });
  tx(s, 6.011, 1.37, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', color: GOLD });
  tx(s, 6.011, 3.227, 6.275, 2.193, cut(LOREM, 432), { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 6.011, 2.51, 6.275, 0.678, cut(LOREM, 142), { align: 'justify', lh: 1.5, color: SILVER });
  photo(s, 2.546, 2.134, 2.764, 3.403, 'rect', { fill: 'CBCBCB' });
}

// ---------------------------------------------------------------- 30. Best Watch Desktop
function slide30(s) {
  grad(s, rectPts(0, 0, 13.333, 4.833), { angle: 120 });
  photo(s, 7.404, 0.903, 4.694, 5.053, 'rect', { shadow: true, fill: 'ACACAE' });

  tx(s, 0.836, 1.258, 5.132, 0.64, 'Best Watch Desktop', { sz: 32, bold: true, color: WHITE });
  tx(s, 0.836, 0.922, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', color: GOLD });
  tx(s, 0.836, 2.779, 6.275, 1.587, cut(LOREM, 349), { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 0.836, 2.062, 6.275, 0.678, cut(LOREM, 142), { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 0.836, 5.23, 0.587, 0.589, '01', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'ellipse', fill: GOLD, color: BLUE });
  tx(s, 1.598, 5.141, 4.229, 0.678, cut(LOREM, 90), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 0.836, 6.094, 0.587, 0.589, '02', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'ellipse', fill: GOLD, color: BLUE });
  tx(s, 1.598, 6.005, 4.229, 0.678, cut(LOREM, 90), { align: 'justify', lh: 1.5, color: GRAY });
  photo(s, 9.806, 1.135, 2.149, 3.392, 'rect', { fill: 'CBCBCB' });
}

// ---------------------------------------------------------------- 31. Our Mobile Mockup
function slide31(s) {
  sh(s, 'rect', 0, 0, 13.333, 7.5, { fill: NAVY });
  sh(s, 'rect', 8.083, 0, 3.729, 7.5, { fill: GOLD });
  photo(s, 7.885, 0.74, 7.177, 6.76, 'rect', { fill: '946E5D' });

  photo(s, 8.451, 1.566, 2.289, 4.049, 'rect', { fill: 'CBCBCB' });
  sh(s, 'ellipse', 0.8, 4.431, 1.398, 1.398, { shadow: true, fill: BLUE });
  tx(s, 0.866, 4.878, 1.267, 0.505, '236', { sz: 24, font: 'SEMI', align: 'center', color: WHITE });
  sh(s, 'ellipse', 2.499, 4.431, 1.398, 1.398, { shadow: true, fill: GOLD });
  tx(s, 2.565, 4.878, 1.267, 0.505, '+724', { sz: 24, font: 'SEMI', align: 'center', fill: GOLD, color: WHITE });
  tx(s, 0.884, 2.565, 5.822, 1.587, cut(LOREM, 326), { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 0.884, 2.31, 2.361, 0.303, 'INSERT TEXT HERE', { font: 'SEMI', color: GOLD });
  tx(s, 0.884, 1.714, 4.655, 0.572, 'Our Mobile Mockup', { sz: 28, bold: true, color: WHITE });
  sh(s, 'roundRect', 0.535, 0.45, 0.312, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 0.603, 0.544, 0.244, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 0.603, 0.637, 0.244, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  tx(s, 11.302, 2.364, 3.035, 0.337, 'Best Day for Diving in 2021', { sz: 14, bold: true, rot: 90, color: WHITE });
  tx(s, 0.284, 6.949, 3.408, 0.337, 'Published by: yumnacreative', { sz: 14, bold: true, color: GOLD });
  tx(s, 8.705, 5.111, 1.78, 0.375, 'Watch More', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'roundRect', rad: 0.076, shadow: true, fill: BLUE, color: WHITE });
  tx(s, 0.847, 0.489, 1.426, 0.269, 'Trends', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 2.266, 0.489, 1.426, 0.269, 'Diving', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 3.608, 0.489, 1.426, 0.269, 'Ocean', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 4.91, 0.489, 1.426, 0.269, 'Contact', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  sh(s, 'ellipse', 4.246, 4.431, 1.398, 1.398, { shadow: true, fill: BLUE });
  tx(s, 4.311, 4.878, 1.267, 0.505, '2846', { sz: 24, font: 'SEMI', align: 'center', color: WHITE });
}

// ---------------------------------------------------------------- 32. Expert Mobile
function slide32(s) {
  grad(s, rectPts(0, 0, 6.771, 7.5), { angle: 120 });
  photo(s, 0.726, 1.091, 1.714, 5.059, 'rect', { rot: -17.661, flipH: true, shadow: true, fill: 'EFEDEB' });
  photo(s, 3.857, 1.108, 1.936, 5.716, 'rect', { rot: 15, flipH: true, shadow: true, fill: 'EFEDEB' });

  photo(s, 0.57, 1.859, 2.612, 3.782, 'rect', { fill: 'CBCBCB' });
  photo(s, 3.53, 2.106, 2.952, 4.273, 'rect', { rot: 32.661, fill: 'CBCBCB' });
  tx(s, 7.514, 1.648, 4.405, 0.64, 'Expert Mobile', { sz: 32, bold: true, color: BLUE });
  tx(s, 7.514, 1.311, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', color: GOLD });
  tx(s, 7.514, 2.312, 5.025, 1.284, cut(LOREM, 211), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 7.514, 3.621, 5.025, 2.799, LOREM, { align: 'justify', lh: 1.5, color: GRAY });
  sh(s, 'ellipse', 1.419, 0.867, 1.975, 1.975, { shadow: true, fill: GOLD, line: WHITE, lw: 6 });
  tx(s, 1.492, 1.265, 1.902, 1.134, [['1', WHITE], ['st', WHITE, { sup: true }], [' ', WHITE, { br: true }], ['Amazing Diving', WHITE, { sz: 20 }]], { sz: 28, bold: true, align: 'center', lh: 0.9 });
}

// ---------------------------------------------------------------- 33. Expert Mobile
function slide33(s) {
  grad(s, rectPts(0, 2.667, 13.333, 4.833), { angle: 120 });
  path(s, 8.336, 1.005, 4.007, 5.689, [[1, 0.958], [1, 0.981, 0.974, 1, 0.941, 1], [0.06, 1, 0.06, 1, 0.06, 1], [0.026, 1, 0, 0.981, 0, 0.958], [0, 0.041, 0, 0.041, 0, 0.041], [0, 0.019, 0.026, 0, 0.06, 0], [0.941, 0, 0.941, 0, 0.941, 0], [0.974, 0, 1, 0.019, 1, 0.041], [1, 0.958, 1, 0.958, 1, 0.958]], { shadow: true, fill: SILVER });
  sh(s, 'ellipse', 10.303, 1.235, 0.072, 0.072, { fill: WHITE });
  sh(s, 'ellipse', 10.303, 1.231, 0.072, 0.072, { fill: WHITE });
  sh(s, 'ellipse', 10.318, 1.245, 0.043, 0.043, { fill: WHITE });
  sh(s, 'ellipse', 10.327, 1.255, 0.024, 0.024, { fill: WHITE });
  path(s, 10.337, 1.264, 0.005, 0.005, [[1, 0], [1, 1], [0, 0], [1, 0], [1, 0]], { fill: WHITE });
  sh(s, 'rect', 10.226, 1.255, 0.034, 0.034, { fill: WHITE });
  sh(s, 'rect', 8.576, 1.481, 3.531, 4.708, { shadow: true, fill: WHITE });
  sh(s, 'ellipse', 10.193, 6.29, 0.298, 0.303, { fill: GRAY });
  path(s, 10.265, 6.362, 0.154, 0.154, [[0.938, 0.656], [0.875, 0.656, 0.875, 0.656, 0.875, 0.656], [0.875, 0.781, 0.781, 0.875, 0.656, 0.875], [0.344, 0.875, 0.344, 0.875, 0.344, 0.875], [0.219, 0.875, 0.125, 0.781, 0.125, 0.656], [0.125, 0.344, 0.125, 0.344, 0.125, 0.344], [0.125, 0.219, 0.219, 0.125, 0.344, 0.125], [0.656, 0.125, 0.656, 0.125, 0.656, 0.125], [0.781, 0.125, 0.875, 0.219, 0.875, 0.344], [0.875, 0.656, 0.875, 0.656, 0.875, 0.656], [0.938, 0.656, 0.938, 0.656, 0.938, 0.656], [1, 0.656, 1, 0.656, 1, 0.656], [1, 0.344, 1, 0.344, 1, 0.344], [1, 0.156, 0.875, 0, 0.656, 0], [0.344, 0, 0.344, 0, 0.344, 0], [0.156, 0, 0, 0.156, 0, 0.344], [0, 0.656, 0, 0.656, 0, 0.656], [0, 0.875, 0.156, 1, 0.344, 1], [0.656, 1, 0.656, 1, 0.656, 1], [0.875, 1, 1, 0.875, 1, 0.656], [0.938, 0.656]], { fill: WHITE });

  photo(s, 8.574, 1.413, 3.531, 4.781, 'rect', { fill: 'CBCBCB' });
  tx(s, 0.991, 0.773, 4.405, 0.64, 'Expert Mobile', { sz: 32, bold: true, color: BLUE });
  tx(s, 0.991, 0.437, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', color: GOLD });
  tx(s, 0.991, 1.438, 6.168, 0.981, cut(LOREM, 197), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 0.991, 3.296, 1.186, 0.909, '01', { sz: 48, bold: true, font: 'SEMI', align: 'center', color: GOLD });
  tx(s, 1.998, 3.551, 4.143, 0.678, cut(DEV, 91), { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 1.998, 3.296, 2.361, 0.303, 'Insert text here', { font: 'SEMI', color: WHITE });
  tx(s, 0.991, 4.412, 1.186, 0.909, '02', { sz: 48, bold: true, font: 'SEMI', align: 'center', color: GOLD });
  tx(s, 1.998, 4.667, 4.143, 0.678, cut(DEV, 91), { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 1.998, 4.412, 2.361, 0.303, 'Insert text here', { font: 'SEMI', color: WHITE });
  tx(s, 0.991, 5.528, 1.186, 0.909, '03', { sz: 48, bold: true, font: 'SEMI', align: 'center', color: GOLD });
  tx(s, 1.998, 5.783, 4.143, 0.678, cut(DEV, 91), { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 1.998, 5.528, 2.361, 0.303, 'Insert text here', { font: 'SEMI', color: WHITE });
  sh(s, 'roundRect', 10.088, 0.353, 2.915, 1.721, { fill: NAVY, shadow: true, rad: 0.287 });
  grad(s, roundPts(10.088, 0.353, 2.915, 1.721, 0.287));
  tx(s, 10.51, 1.076, 2.542, 0.692, 'Amazing Trending Diving Spot 2021 ', { sz: 14, bold: true, font: 'SEMI', lh: 1.3, color: WHITE });
  tx(s, 10.51, 0.516, 1.313, 0.558, [['14', GOLD], ['th', GOLD, { sup: true }], [' Jan', GOLD]], { sz: 20, bold: true, font: 'SEMI', lh: 1.5 });
}

// ---------------------------------------------------------------- 34. $125
function slide34(s) {
  grad(s, [[13.333, 0], [6.248, 0], [13.333, 7.5]], { angle: 330 });

  sh(s, 'roundRect', 9.158, 1.679, 3.592, 4.813, { rad: 0.392, shadow: true, fill: WHITE });
  tx(s, 9.524, 4.757, 2.781, 0.303, '24/7 Full Support', { align: 'center', color: BLUE });
  tx(s, 9.384, 3.274, 3, 1.01, '$125', { sz: 54, bold: true, align: 'center', color: BLUE });
  tx(s, 9.705, 2.672, 2.419, 0.41, '50% Cashback', { sz: 16, align: 'center', shape: 'roundRect', rad: 0.068, fill: GOLD, color: BLUE });
  tx(s, 9.384, 4.352, 3, 0.337, 'Features', { sz: 14, bold: true, align: 'center', color: BLUE });
  tx(s, 9.384, 1.925, 3, 0.707, 'Basic', { sz: 36, font: 'MED', align: 'center', color: BLUE });
  tx(s, 9.524, 5.129, 2.781, 0.303, 'Pro Trainer', { align: 'center', color: BLUE });
  tx(s, 9.524, 5.504, 2.781, 0.303, 'Pro Equipment', { align: 'center', color: BLUE });
  tx(s, 9.524, 5.879, 2.781, 0.303, 'Full Guaranted', { align: 'center', color: BLUE });
  sh(s, 'roundRect', 5.201, 1.679, 3.592, 4.813, { rad: 0.392, shadow: true, fill: WHITE });
  tx(s, 5.567, 4.757, 2.781, 0.303, '24/7 Full Support', { align: 'center', color: BLUE });
  tx(s, 5.428, 3.274, 3, 1.01, '$125', { sz: 54, bold: true, align: 'center', color: BLUE });
  tx(s, 5.749, 2.672, 2.419, 0.41, '50% Cashback', { sz: 16, align: 'center', shape: 'roundRect', rad: 0.068, fill: GOLD, color: BLUE });
  tx(s, 5.428, 4.352, 3, 0.337, 'Features', { sz: 14, bold: true, align: 'center', color: BLUE });
  tx(s, 5.428, 1.925, 3, 0.707, 'Premium', { sz: 36, font: 'MED', align: 'center', color: BLUE });
  tx(s, 5.567, 5.129, 2.781, 0.303, 'Pro Trainer', { align: 'center', color: BLUE });
  tx(s, 5.567, 5.504, 2.781, 0.303, 'Pro Equipment', { align: 'center', color: BLUE });
  tx(s, 5.567, 5.879, 2.781, 0.303, 'Full Guaranted', { align: 'center', color: BLUE });
  tx(s, 0.45, 1.848, 4.405, 0.64, 'Pricing Plans', { sz: 32, bold: true, color: BLUE });
  tx(s, 0.45, 1.511, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', color: GOLD });
  tx(s, 0.45, 2.512, 4.127, 0.678, cut(LOREM, 86), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 0.45, 3.287, 4.127, 3.405, LOREM, { align: 'justify', lh: 1.5, color: GRAY });
  sh(s, 'roundRect', 0.535, 0.45, 0.312, 0.051, { rad: 0.009, shadow: true, fill: BLUE });
  sh(s, 'roundRect', 0.603, 0.544, 0.244, 0.051, { rad: 0.009, shadow: true, fill: BLUE });
  sh(s, 'roundRect', 0.603, 0.637, 0.244, 0.051, { rad: 0.009, shadow: true, fill: BLUE });
  tx(s, 11.458, 2.344, 3.035, 0.337, 'Best Day for Diving in 2021', { sz: 14, bold: true, rot: 90, color: WHITE });
  tx(s, 0.284, 6.949, 3.408, 0.337, 'Published by: yumnacreative', { sz: 14, bold: true, color: GOLD });
  tx(s, 0.847, 0.489, 1.426, 0.269, 'Trends', { sz: 10, font: 'MED', align: 'center', color: BLUE });
  tx(s, 2.266, 0.489, 1.426, 0.269, 'Diving', { sz: 10, font: 'MED', align: 'center', color: BLUE });
  tx(s, 3.608, 0.489, 1.426, 0.269, 'Ocean', { sz: 10, font: 'MED', align: 'center', color: BLUE });
  tx(s, 4.91, 0.489, 1.426, 0.269, 'Contact', { sz: 10, font: 'MED', align: 'center', color: BLUE });
}

// ---------------------------------------------------------------- 35. Our Charts
function slide35(s) {
  grad(s, rectPts(0, 6.794, 13.333, 0.706), { angle: 120 });

  sh(s, 'rect', 0.917, 0.406, 4.411, 6.145, { shadow: true, fill: WHITE });
  sh(s, 'ellipse', 1.393, 1.601, 1.688, 1.689, { fill: SILVER });
  path(s, 2.29, 0.705, 1.688, 1.471, [[0.858, 0.975], [0.858, 0.973, 0.858, 0.973, 0.858, 0.973], [0.945, 0.871, 1, 0.729, 1, 0.575], [1, 0.257, 0.776, 0, 0.499, 0], [0.223, 0, 0, 0.257, 0, 0.575], [0, 0.729, 0.052, 0.869, 0.139, 0.972], [0.139, 0.973, 0.139, 0.973, 0.139, 0.973], [0.144, 0.979, 0.149, 0.984, 0.153, 0.989], [0.33, 0.785, 0.33, 0.785, 0.33, 0.785], [0.328, 0.783, 0.325, 0.78, 0.323, 0.777], [0.322, 0.776, 0.322, 0.776, 0.322, 0.776], [0.278, 0.724, 0.25, 0.653, 0.25, 0.575], [0.25, 0.416, 0.361, 0.287, 0.499, 0.287], [0.638, 0.287, 0.749, 0.416, 0.749, 0.575], [0.749, 0.652, 0.722, 0.723, 0.678, 0.775], [0.678, 0.775, 0.678, 0.775, 0.678, 0.775], [0.678, 0.775, 0.678, 0.775, 0.678, 0.775], [0.67, 0.784, 0.662, 0.795, 0.654, 0.804], [0.837, 1, 0.837, 1, 0.837, 1], [0.844, 0.991, 0.851, 0.983, 0.858, 0.975]], { fill: BLUE });
  path(s, 3.185, 1.601, 1.69, 1.689, [[0.514, 0], [0.509, 0.093, 0.481, 0.181, 0.433, 0.259], [0.454, 0.253, 0.477, 0.25, 0.5, 0.25], [0.638, 0.25, 0.749, 0.361, 0.749, 0.499], [0.749, 0.638, 0.638, 0.749, 0.5, 0.749], [0.362, 0.749, 0.251, 0.638, 0.251, 0.499], [0.251, 0.439, 0.271, 0.384, 0.306, 0.341], [0.123, 0.171, 0.123, 0.171, 0.123, 0.171], [0.046, 0.259, 0, 0.374, 0, 0.499], [0, 0.776, 0.224, 1, 0.5, 1], [0.776, 1, 1, 0.776, 1, 0.499], [1, 0.228, 0.783, 0.007, 0.514, 0]], { fill: GOLD });
  tx(s, 1.923, 4.122, 2.524, 0.678, cut(DEV, 44), { align: 'center', lh: 1.5, color: GRAY });
  tx(s, 2.892, 3.449, 0.587, 0.589, '01', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'ellipse', fill: BLUE, color: WHITE });
  tx(s, 1.923, 5.647, 2.524, 0.678, cut(DEV, 44), { align: 'center', lh: 1.5, color: GRAY });
  tx(s, 2.892, 4.974, 0.587, 0.589, '02', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'ellipse', fill: BLUE, color: WHITE });
  tx(s, 5.95, 0.961, 4.405, 0.64, 'Our Charts', { sz: 32, bold: true, color: BLUE });
  tx(s, 5.95, 0.625, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', color: GOLD });
  tx(s, 5.95, 1.626, 4.127, 0.678, cut(LOREM, 86), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 5.95, 2.401, 6.221, 2.193, LOREM, { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 5.904, 4.805, 0.587, 0.589, '01', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'ellipse', fill: GOLD, color: BLUE });
  tx(s, 6.667, 4.716, 4.229, 0.678, cut(LOREM, 90), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 5.904, 5.669, 0.587, 0.589, '02', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'ellipse', fill: GOLD, color: BLUE });
  tx(s, 6.667, 5.58, 4.229, 0.678, cut(LOREM, 90), { align: 'justify', lh: 1.5, color: GRAY });
}

// ---------------------------------------------------------------- 36. Our Charts
function slide36(s) {
  grad(s, rectPts(0, 6.794, 13.333, 0.706), { angle: 120 });

  histogramChart(s, 6.93, 0.971, 5.77, 3.261);
  tx(s, 6.93, 4.585, 5.537, 0.981, cut(DEV, 172), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 6.93, 5.7, 5.537, 0.678, cut(DEV, 123), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 0.446, 0.961, 4.405, 0.64, 'Our Charts', { sz: 32, bold: true, color: BLUE });
  tx(s, 0.446, 0.625, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', color: GOLD });
  tx(s, 0.446, 1.626, 4.127, 0.678, cut(LOREM, 86), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 0.446, 2.401, 6.221, 2.193, LOREM, { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 0.401, 4.805, 0.587, 0.589, '01', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'ellipse', fill: GOLD, color: BLUE });
  tx(s, 1.163, 4.716, 4.229, 0.678, cut(LOREM, 90), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 0.401, 5.669, 0.587, 0.589, '02', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'ellipse', fill: GOLD, color: BLUE });
  tx(s, 1.163, 5.58, 4.229, 0.678, cut(LOREM, 90), { align: 'justify', lh: 1.5, color: GRAY });
}

// ---------------------------------------------------------------- 37. Our Charts
function slide37(s) {
  grad(s, rectPts(0, 6.794, 13.333, 0.706), { angle: 120 });

  stackedChart(s, 0.926, 1.086, 5.397, 5.606);
  tx(s, 6.544, 0.961, 4.405, 0.64, 'Our Charts', { sz: 32, bold: true, color: BLUE });
  tx(s, 6.544, 0.625, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', color: GOLD });
  tx(s, 6.544, 1.626, 4.127, 0.678, cut(LOREM, 86), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 6.544, 2.401, 6.221, 2.193, LOREM, { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 6.499, 4.805, 0.587, 0.589, '01', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'ellipse', fill: GOLD, color: BLUE });
  tx(s, 7.261, 4.716, 4.229, 0.678, cut(LOREM, 90), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 6.499, 5.669, 0.587, 0.589, '02', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'ellipse', fill: GOLD, color: BLUE });
  tx(s, 7.261, 5.58, 4.229, 0.678, cut(LOREM, 90), { align: 'justify', lh: 1.5, color: GRAY });
}

// ---------------------------------------------------------------- 38. Our Charts
function slide38(s) {
  grad(s, rectPts(0, 6.794, 13.333, 0.706), { angle: 120 });

  waterfallChart(s, 7.551, 0.979, 5.071, 5.25);
  tx(s, 0.757, 0.961, 4.405, 0.64, 'Our Charts', { sz: 32, bold: true, color: BLUE });
  tx(s, 0.757, 0.625, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', color: GOLD });
  tx(s, 0.757, 1.626, 4.127, 0.678, cut(LOREM, 86), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 0.757, 2.401, 6.221, 2.193, LOREM, { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 0.711, 4.805, 0.587, 0.589, '01', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'ellipse', fill: GOLD, color: BLUE });
  tx(s, 1.474, 4.716, 4.229, 0.678, cut(LOREM, 90), { align: 'justify', lh: 1.5, color: GRAY });
  tx(s, 0.711, 5.669, 0.587, 0.589, '02', { font: 'SEMI', align: 'center', valign: 'middle', shape: 'ellipse', fill: GOLD, color: BLUE });
  tx(s, 1.474, 5.58, 4.229, 0.678, cut(LOREM, 90), { align: 'justify', lh: 1.5, color: GRAY });
}

// ---------------------------------------------------------------- 39. Our Contact
function slide39(s) {
  grad(s, rectPts(0, 2.667, 13.333, 4.833), { angle: 120 });
  sh(s, 'rect', 0, 0, 6.667, 2.667, { fill: GOLD });

  photo(s, 6.667, 0, 6.667, 5.486, 'rect', { fill: 'CBCBCB' });
  path(s, 8.421, 5.854, 0.274, 0.22, [[1, 0.121], [1, 0.121], [0.961, 0.145, 0.922, 0.169, 0.885, 0.169], [0.922, 0.145, 0.961, 0.097, 0.961, 0.027], [0.922, 0.051, 0.885, 0.072, 0.844, 0.097], [0.807, 0.051, 0.748, 0, 0.692, 0], [0.575, 0, 0.479, 0.121, 0.479, 0.263], [0.479, 0.287, 0.479, 0.311, 0.499, 0.311], [0.328, 0.311, 0.171, 0.217, 0.076, 0.051], [0.056, 0.097, 0.037, 0.145, 0.037, 0.193], [0.037, 0.263, 0.076, 0.359, 0.132, 0.405], [0.095, 0.383, 0.076, 0.383, 0.037, 0.359], [0.037, 0.477, 0.115, 0.595, 0.213, 0.619], [0.193, 0.619, 0.171, 0.619, 0.152, 0.619], [0.132, 0.619, 0.132, 0.619, 0.115, 0.619], [0.132, 0.713, 0.213, 0.788, 0.308, 0.788], [0.23, 0.855, 0.152, 0.906, 0.056, 0.906], [0.037, 0.906, 0.017, 0.906, 0, 0.906], [0.095, 0.976, 0.193, 1, 0.308, 1], [0.692, 1, 0.885, 0.619, 0.885, 0.287], [0.885, 0.263], [0.941, 0.217, 0.961, 0.193, 1, 0.121]], { fill: GOLD });
  tx(s, 8.748, 5.828, 1.471, 0.321, 'Twi.account', { font: 'SEMI', lh: 1.14, color: GOLD });
  path(s, 6.547, 5.831, 0.253, 0.255, [[0.875, 0], [0.875, 0], [0.125, 0, 0.125, 0, 0.125, 0], [0.064, 0, 0, 0.063, 0, 0.124], [0, 0.873, 0, 0.873, 0, 0.873], [0, 0.937, 0.064, 1, 0.125, 1], [0.501, 1, 0.501, 1, 0.501, 1], [0.501, 0.643, 0.501, 0.643, 0.501, 0.643], [0.376, 0.643, 0.376, 0.643, 0.376, 0.643], [0.376, 0.479, 0.376, 0.479, 0.376, 0.479], [0.501, 0.479, 0.501, 0.479, 0.501, 0.479], [0.501, 0.394, 0.501, 0.394, 0.501, 0.394], [0.501, 0.291, 0.605, 0.188, 0.708, 0.188], [0.814, 0.188, 0.814, 0.188, 0.814, 0.188], [0.814, 0.376, 0.814, 0.376, 0.814, 0.376], [0.729, 0.376, 0.729, 0.376, 0.729, 0.376], [0.687, 0.376, 0.687, 0.376, 0.687, 0.394], [0.687, 0.479, 0.687, 0.479, 0.687, 0.479], [0.814, 0.479, 0.814, 0.479, 0.814, 0.479], [0.814, 0.643, 0.814, 0.643, 0.814, 0.643], [0.687, 0.643, 0.687, 0.643, 0.687, 0.643], [0.687, 1, 0.687, 1, 0.687, 1], [0.875, 1, 0.875, 1, 0.875, 1], [0.939, 1, 1, 0.937, 1, 0.873], [1, 0.124, 1, 0.124, 1, 0.124], [1, 0.063, 0.939, 0, 0.875, 0]], { fill: GOLD });
  tx(s, 6.867, 5.828, 1.515, 0.321, 'Fac.account', { font: 'SEMI', lh: 1.14, color: GOLD });
  path(s, 10.289, 5.882, 0.242, 0.104, [[0.543, 0.967], [0.53, 0.989, 0.516, 1, 0.5, 1], [0.484, 1, 0.47, 0.989, 0.457, 0.967], [0.309, 0.708], [0.28, 0.656], [0, 0.167], [0, 0.167], [0, 0.075, 0.032, 0, 0.071, 0], [0.929, 0], [0.968, 0, 1, 0.075, 1, 0.167], [1, 0.167, 0.543, 0.967, 0.543, 0.967]], { fill: GOLD });
  path(s, 10.471, 5.91, 0.061, 0.091, [[1, 1], [0, 0.5], [1, 0], [1, 0, 1, 1, 1, 1]], { fill: GOLD });
  path(s, 10.289, 5.961, 0.242, 0.068, [[1, 0.746], [1, 0.886, 0.968, 1, 0.929, 1], [0.071, 1], [0.032, 1, 0, 0.886, 0, 0.746], [0.28, 0], [0.436, 0.416], [0.455, 0.467, 0.477, 0.492, 0.5, 0.492], [0.523, 0.492, 0.545, 0.467, 0.564, 0.416], [0.72, 0], [0.72, 0, 1, 0.746, 1, 0.746]], { fill: GOLD });
  path(s, 10.289, 5.91, 0.061, 0.091, [[0, 0], [1, 0.5], [0, 1], [0, 1, 0, 0, 0, 0]], { fill: GOLD });
  path(s, 10.272, 5.864, 0.277, 0.182, [[0.875, 0], [0.125, 0], [0.056, 0, 0, 0.085, 0, 0.19], [0, 0.809], [0, 0.915, 0.056, 1, 0.125, 1], [0.875, 1], [0.944, 1, 1, 0.915, 1, 0.809], [1, 0.19], [1, 0.085, 0.944, 0, 0.875, 0]], { fill: GOLD });
  tx(s, 10.6, 5.828, 1.61, 0.321, 'Email.Contact', { font: 'SEMI', lh: 1.14, color: GOLD });
  tx(s, 1.128, 4.216, 2.186, 0.303, 'Address Company', { bold: true, font: 'SEMI', color: GOLD });
  tx(s, 1.128, 4.45, 2.052, 0.678, [['76 United Street Ca,', SILVER, { br: true }], ['California, CA 62 71593', SILVER]], { lh: 1.5 });
  tx(s, 3.94, 4.219, 2.186, 0.303, 'Phone & Email', { bold: true, font: 'SEMI', color: GOLD });
  sh(s, 'ellipse', 0.522, 4.264, 0.473, 0.472, { shadow: true, fill: GOLD });
  path(s, 0.636, 4.393, 0.223, 0.206, [[0.994, 0.003], [0.99, 0, 0.985, 0, 0.981, 0.002], [0.009, 0.564], [0.002, 0.567, 0, 0.576, 0.003, 0.582], [0.005, 0.586, 0.007, 0.588, 0.01, 0.589], [0.308, 0.7], [0.313, 0.702], [0.398, 0.99], [0.399, 0.995, 0.402, 0.998, 0.407, 1], [0.408, 1, 0.409, 1, 0.41, 1], [0.413, 1, 0.416, 0.999, 0.419, 0.996], [0.595, 0.805], [0.854, 0.902], [0.858, 0.903, 0.862, 0.903, 0.865, 0.901], [0.868, 0.899, 0.87, 0.895, 0.871, 0.891], [0.999, 0.017], [1, 0.012, 0.998, 0.006, 0.994, 0.003]], { fill: BLUE });
  path(s, 0.647, 4.41, 0.189, 0.122, [[0, 0.828], [1, 0], [1, 0, 1, 0, 1, 0], [1, 0, 1, 0, 1, 0], [0.321, 1], [0, 0.828], [0, 0.828, 0, 0.828, 0, 0.828], [0, 0.828, 0, 0.828, 0, 0.828]], { fill: BLUE });
  path(s, 0.712, 4.417, 0.126, 0.169, [[0, 0.705], [1, 0], [1, 0, 1, 0, 1, 0], [1, 0, 1, 0, 1, 0], [0.219, 0.756], [0.217, 0.758, 0.216, 0.76, 0.215, 0.762], [0.126, 1], [0.126, 1, 0.126, 1, 0.126, 1], [0.126, 1, 0.126, 1, 0.126, 1]], { fill: BLUE });
  path(s, 0.733, 4.55, 0.03, 0.037, [[0, 0.999], [0.338, 0], [1, 0.186], [0.001, 1], [0.001, 1, 0, 1, 0, 0.999], [0, 0.999, 0, 0.999, 0, 0.999]], { fill: BLUE });
  path(s, 0.746, 4.408, 0.106, 0.165, [[0.752, 1], [0, 0.834], [0.999, 0], [1, 0, 1, 0, 1, 0], [1, 0, 1, 0, 1, 0], [0.752, 1], [0.752, 1, 0.752, 1, 0.752, 1], [0.752, 1, 0.752, 1, 0.752, 1]], { fill: BLUE });
  sh(s, 'ellipse', 3.334, 4.264, 0.473, 0.472, { shadow: true, fill: GOLD });
  path(s, 3.456, 4.385, 0.229, 0.229, [[0.985, 0.782], [0.812, 0.61], [0.791, 0.588, 0.757, 0.588, 0.735, 0.609], [0.735, 0.609, 0.735, 0.61, 0.735, 0.61], [0.642, 0.706], [0.631, 0.716, 0.615, 0.716, 0.604, 0.706], [0.294, 0.397], [0.284, 0.387, 0.284, 0.37, 0.294, 0.36], [0.391, 0.265], [0.412, 0.244, 0.412, 0.21, 0.391, 0.188], [0.391, 0.188, 0.391, 0.188, 0.391, 0.187], [0.218, 0.016], [0.208, 0.006, 0.195, 0, 0.181, 0], [0.166, 0, 0.152, 0.006, 0.142, 0.016], [0.043, 0.115], [0.018, 0.138, 0.003, 0.17, 0.002, 0.204], [0, 0.259, 0.01, 0.314, 0.031, 0.366], [0.051, 0.412, 0.076, 0.456, 0.106, 0.497], [0.207, 0.651, 0.337, 0.784, 0.487, 0.891], [0.512, 0.909, 0.538, 0.926, 0.566, 0.941], [0.628, 0.973, 0.696, 0.993, 0.766, 1], [0.77, 1, 0.774, 1, 0.778, 1], [0.822, 1, 0.863, 0.982, 0.893, 0.95], [0.983, 0.86], [0.994, 0.849, 1, 0.835, 1, 0.821], [1, 0.806, 0.995, 0.793, 0.985, 0.782]], { fill: BLUE });
  path(s, 3.485, 4.391, 0.06, 0.059, [[0.134, 0.035], [0.156, 0.014, 0.185, 0.001, 0.215, 0], [0.244, 0.001, 0.271, 0.014, 0.291, 0.035], [0.954, 0.701], [0.999, 0.745, 1, 0.817, 0.957, 0.863], [0.956, 0.863, 0.955, 0.864, 0.955, 0.865], [0.819, 1], [0, 0.171]], { fill: BLUE });
  path(s, 3.461, 4.405, 0.204, 0.203, [[0.956, 0.951], [0.956, 0.952], [0.956, 0.952], [0.928, 0.983, 0.888, 1, 0.847, 1], [0.843, 1, 0.839, 1, 0.835, 1], [0.761, 0.993, 0.688, 0.972, 0.622, 0.937], [0.592, 0.921, 0.564, 0.903, 0.537, 0.883], [0.537, 0.883], [0.537, 0.882], [0.371, 0.764, 0.228, 0.616, 0.116, 0.446], [0.084, 0.402, 0.057, 0.354, 0.035, 0.304], [0.018, 0.258, 0, 0.197, 0.004, 0.134], [0.004, 0.133], [0.004, 0.133], [0.006, 0.102, 0.019, 0.073, 0.042, 0.052], [0.042, 0.052], [0.042, 0.052], [0.094, 0], [0.333, 0.24], [0.285, 0.288], [0.262, 0.311, 0.262, 0.347, 0.285, 0.37], [0.632, 0.717], [0.655, 0.739, 0.692, 0.739, 0.715, 0.717], [0.761, 0.668], [1, 0.908]], { fill: BLUE });
  path(s, 3.62, 4.526, 0.059, 0.06, [[0.954, 0.871], [0.824, 1], [0, 0.18], [0.13, 0.046], [0.173, 0.001, 0.245, 0, 0.29, 0.043], [0.291, 0.044, 0.291, 0.045, 0.292, 0.046], [0.959, 0.709], [1, 0.756, 0.998, 0.826, 0.954, 0.871]], { fill: BLUE });
  tx(s, 3.94, 4.73, 2.186, 0.375, '@matias.official.com', { lh: 1.5, color: SILVER });
  tx(s, 3.94, 4.454, 1.978, 0.375, '+65 457 7251 6241', { lh: 1.5, color: SILVER });
  tx(s, 0.489, 3.481, 4.008, 0.572, 'Our Contact', { sz: 28, bold: true, color: WHITE });
  tx(s, 0.489, 3.109, 3.623, 0.375, 'Proposal Presentation Template', { font: 'SEMI', lh: 1.5, color: WHITE });
  sh(s, 'roundRect', 0.535, 0.45, 0.312, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 0.603, 0.544, 0.244, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 0.603, 0.637, 0.244, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  tx(s, 11.458, 2.344, 3.035, 0.337, 'Best Day for Diving in 2021', { sz: 14, bold: true, rot: 90, color: WHITE });
  tx(s, 0.847, 0.489, 1.426, 0.269, 'Trends', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 2.266, 0.489, 1.426, 0.269, 'Diving', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 3.608, 0.489, 1.426, 0.269, 'Ocean', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 4.91, 0.489, 1.426, 0.269, 'Contact', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 0.522, 5.276, 5.604, 0.981, cut(LOREM, 173), { align: 'justify', lh: 1.5, color: SILVER });
  tx(s, 0.284, 6.949, 3.408, 0.337, 'Published by: yumnacreative', { sz: 14, bold: true, color: GOLD });
}

// ---------------------------------------------------------------- 40. Thanks
function slide40(s) {
  grad(s, rectPts(0, 0, 13.333, 7.5), { angle: 120 });
  sh(s, 'rect', 8.771, 0, 4.562, 5.429, { fill: GOLD });

  tx(s, 1.22, 3.149, 4.434, 1.447, 'Thanks', { sz: 80, bold: true, align: 'center', color: WHITE });
  tx(s, 1.81, 4.428, 3.254, 0.337, 'Diving Presentation Templates', { sz: 14, font: 'SEMI', align: 'center', color: GOLD });
  sh(s, 'roundRect', 0.535, 0.45, 0.312, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 0.603, 0.544, 0.244, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  sh(s, 'roundRect', 0.603, 0.637, 0.244, 0.051, { rad: 0.009, shadow: true, fill: WHITE });
  tx(s, 0.847, 0.489, 1.426, 0.269, 'Trends', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 2.266, 0.489, 1.426, 0.269, 'Diving', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 3.608, 0.489, 1.426, 0.269, 'Ocean', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 4.91, 0.489, 1.426, 0.269, 'Contact', { sz: 10, font: 'MED', align: 'center', color: WHITE });
  tx(s, 11.329, 3.501, 3.035, 0.337, 'Best Day for Diving in 2021', { sz: 14, bold: true, rot: 90, color: WHITE });
  tx(s, 0.535, 6.75, 1.982, 0.572, [['Published by: ', WHITE], ['yumnacreative', GOLD]], { sz: 14, bold: true });
  photo(s, 6.783, 1.198, 5.429, 5.429, 'ellipse', { shadow: true, fill: 'CBCBCB', line: WHITE, lw: 6 });
}

// ---------------------------------------------------------------- deck
const SLIDES = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
  slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
  slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24,
  slide25, slide26, slide27, slide28, slide29, slide30, slide31, slide32,
  slide33, slide34, slide35, slide36, slide37, slide38, slide39, slide40,
];

const deck = new pptxgen();
deck.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
deck.layout = 'WIDE';
deck.theme = { headFontFace: 'Poppins', bodyFontFace: 'Poppins' };

for (const build of SLIDES) {
  const s = deck.addSlide();
  s.background = { color: WHITE };
  build(s);
}

deck.writeFile({ fileName: join(__dirname, '0790b00f-6a96-46dc-89c0-8b4a3425fc6e_grok_final.pptx') })
  .then(f => console.log('wrote', f));
