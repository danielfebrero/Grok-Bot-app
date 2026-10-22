/**
 * Health Sports — 30-slide presentation template, rebuilt with pptxgenjs.
 *
 * Run:  node 022ced68-2654-409d-a898-1ccb3f4eaf4b_grok_final.js
 *
 * Photographs in the original deck are replaced with flat placeholder
 * rectangles (see `photo()` / `device()` below) — no raster data is embedded.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// --- design tokens ---------------------------------------------------------
const C = {
  BG:     '0D0D0D', // slide background (theme tx1 lightened 5%)
  INK:    'FFFFFF', // primary text
  GHOST:  '262626', // outlined "HEALTH SPORTS" watermark
  PHOTO:  '464646', // image placeholder plate
  TRACK:  'BFBFBF', // progress-bar / gridline grey
  MUTED:  '808080',
  Y1:     'EECA3F', // theme accent1 .. accent6
  Y2:     'EBC721',
  Y3:     'F5C81B',
  Y4:     'FFCF00',
  Y5:     'E7CD2C',
  Y6:     'ECD75A',
};

const F = { HEAD: 'Montserrat', BODY: 'Roboto', SCRIPT: 'Euphoria Script' };

const BR = '\n';

const MONTHS = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];

// Body copy reused across the deck.
const L1  = 'PLACEHOLDER';
const L2  = 'PLACEHOLDER';
const L3  = 'lorem ipsum dolor sit amet consectetuer adipiscing elit';
const L4  = L2 + ' nullam malesuada nulla';
const L5  = L4 + ' seds massa feugiat iste scelerisque';
const L6  = 'PLACEHOLDER';
const L7  = L4 + ' seds massa feugiat iste';
const L8  = 'PLACEHOLDER';
const L9  = 'PLACEHOLDER';
const L10 = 'PLACEHOLDER';
const L11 = 'PLACEHOLDER';
const L12 = 'ipsum volutpat tincidunt dapibus at integer eleifend';
const L13 = L2 + ' nullam malesuada';

// --- primitives ------------------------------------------------------------
/** Text box. Defaults match the deck: Roboto 12pt white, top-aligned, no wrap padding. */
function T(slide, text, opts) {
  slide.addText(text, Object.assign({
    fontFace: F.BODY, fontSize: 12, color: C.INK, valign: 'top', isTextBox: true,
  }, opts));
}

/** Autoshape. */
function S(slide, shapeName, opts) {
  slide.addShape(shapeName, Object.assign({ line: { type: 'none' } }, opts));
}

/** A new slide on the dark background, optionally carrying the master navbar. */
function newSlide(deck, withNav) {
  const s = deck.addSlide();
  s.background = { color: C.BG };
  if (withNav !== false) navbar(s, 0.131, 0.17);
  return s;
}

// --- reusable furniture ----------------------------------------------------
/** Hamburger mark + Fb / Be / Tw links that the slide master paints top-left. */
function navbar(slide, x, y) {
  S(slide, 'rect', { x: x, y: y, w: 0.262, h: 0.035, fill: C.Y1 });
  S(slide, 'rect', { x: x, y: y + 0.06, w: 0.262, h: 0.035, fill: C.Y2 });
  S(slide, 'rect', { x: x, y: y + 0.12, w: 0.18, h: 0.035, fill: C.Y1 });
  S(slide, 'rect', { x: x + 0.21, y: y + 0.12, w: 0.052, h: 0.035, fill: C.Y1 });
  [['Fb', 0.59], ['Be', 1.277], ['Tw', 1.984]].forEach(function (item) {
    T(slide, item[0], {
      x: x + item[1], y: y - 0.019, w: 0.5, h: 0.185,
      fontSize: 11, fontFace: F.HEAD, margin: 0, wrap: false,
    });
  });
}

/** Placeholder standing in for a photograph in the source deck. */
function photo(slide, x, y, w, h, shape) {
  S(slide, shape === 'round' ? 'roundRect' : 'rect',
    { x: x, y: y, w: w, h: h, fill: C.PHOTO, rectRadius: 0.28 });
  T(slide, '[image]', {
    x: x, y: y + h / 2 - 0.2, w: w, h: 0.4,
    fontSize: 12, color: C.INK, transparency: 55, align: 'center', valign: 'middle',
  });
}

/** Phone / tablet / laptop mock-up shell; the screen plate is drawn on top of it. */
function device(slide, x, y, w, h) {
  S(slide, 'roundRect', {
    x: x, y: y, w: w, h: h, fill: '242424',
    line: { color: '3D3D3D', width: 1 }, rectRadius: 0.12,
  });
}

/** Oversized "HEALTH SPORTS" watermark that anchors most slides. */
function ghost(slide, x, y, w, h, size) {
  T(slide, 'HEALTH SPORTS', {
    x: x, y: y, w: w, h: h, fontSize: size, bold: true, fontFace: F.HEAD,
    color: C.GHOST, align: 'center', wrap: false,
  });
}

/** Yellow "Learn More" pill next to the outlined "Improve" pill. */
function buttons(slide, x, y) {
  S(slide, 'roundRect', { x: x, y: y, w: 1.153, h: 0.261, fill: C.Y1, rectRadius: 0.044 });
  T(slide, 'Learn More', { x: x, y: y, w: 1.153, h: 0.261, fontSize: 9, align: 'center', valign: 'middle' });
  S(slide, 'roundRect', {
    x: x + 1.43, y: y, w: 1.153, h: 0.261, fill: { type: 'none' },
    line: { color: C.INK, width: 1 }, rectRadius: 0.044,
  });
  T(slide, 'Improve', { x: x + 1.43, y: y, w: 1.153, h: 0.261, fontSize: 9, align: 'center', valign: 'middle' });
}

/** Decorative opening-quote glyph. */
function quoteMark(slide, x, y, w, h) {
  T(slide, '\u201C', {
    x: x - 0.05, y: y - 0.28, w: w + 0.2, h: h + 0.5,
    fontSize: 44, bold: true, fontFace: F.HEAD, color: C.Y1, transparency: 50,
    align: 'center', valign: 'middle',
  });
}

/** Line-art pictogram (social, sport and contact icons) reduced to a stroked badge. */
function lineIcon(slide, x, y, w, h, color) {
  S(slide, 'roundRect', {
    x: x, y: y, w: w, h: h, fill: { type: 'none' },
    line: { color: color, width: 1.5 }, rectRadius: 0.2,
  });
}

/** Tiny solid pictogram used for the inline social rows and legend markers. */
function dot(slide, x, y, w, h, color) {
  S(slide, 'ellipse', { x: x, y: y, w: w, h: h, fill: color });
}

/** Slide 27's binned-column chart (a native pptxgenjs clustered bar chart). */
function histogram(slide) {
  const bins = ['[1, 5]', '(5, 9]', '(9, 13]', '(13, 17]', '(17, 21]', '(21, 25]'];
  const counts = [5, 11, 23, 24, 9, 4];
  slide.addChart('bar', [{ name: 'Series1', labels: bins, values: counts }], {
    x: 0.708, y: 1.87, w: 5.644, h: 3.763,
    barDir: 'col', barGapWidthPct: 0,
    chartColors: [C.Y1, C.Y2, C.Y3, C.Y4, C.Y5, C.Y6],
    showLegend: false, showValue: false,
    valAxisMaxVal: 25, valAxisMajorUnit: 5,
    catAxisLabelColor: C.INK, catAxisLabelFontFace: F.BODY, catAxisLabelFontSize: 12,
    valAxisLabelColor: C.INK, valAxisLabelFontFace: F.BODY, valAxisLabelFontSize: 12,
    catAxisLineShow: true, valAxisLineShow: true,
    catAxisLineColor: C.INK, valAxisLineColor: C.INK,
    valGridLine: { color: C.MUTED, size: 1 }, catGridLine: { style: 'none' },
    plotArea: { fill: { color: C.BG } }, chartArea: { fill: { color: C.BG } },
  });
}

// --- slide 29 timeline furniture -------------------------------------------
/** Horizontal value gridlines behind the timeline columns. */
function gridLines(slide, ys) {
  ys.forEach(function (y) {
    S(slide, 'line', { x: 1.25, y: y, w: 10.827, h: 0, line: { color: C.TRACK, width: 1, transparency: 50 } });
  });
}

/** The month axis: a grey rail beaded with one white node per month. */
function monthRail(slide, x, y) {
  const STEP = 0.9025;
  for (let i = 0; i < 12; i++) {
    S(slide, 'rect', { x: x + 0.081 + i * STEP, y: y + 0.046, w: 0.903, h: 0.081, fill: C.TRACK });
  }
  for (let i = 0; i < 13; i++) {
    S(slide, 'ellipse', {
      x: x + i * STEP, y: y, w: 0.172, h: 0.172,
      fill: C.INK, line: { color: C.TRACK, width: 5 },
    });
    if (i < 12) {
      T(slide, MONTHS[i], { x: x + 0.19 + i * STEP, y: y + 0.211, w: 0.6, h: 0.303, wrap: false });
    }
  }
}

/** One value column per data point: the rounded bar, its stem and date label. */
function timelineBars(slide, bars) {
  bars.forEach(function (b) {
    S(slide, 'line', { x: b.stemX, y: b.stemTop, w: 0, h: 4.747 - b.stemTop, line: { color: C.MUTED, width: 1, transparency: 50 } });
    S(slide, 'roundRect', { x: b.x, y: b.top, w: 0.184, h: b.h, fill: b.color, rectRadius: 0.037 });
    T(slide, '14/01/2016', { x: b.labelX, y: b.labelY, w: 1.135, h: 0.303, align: 'center', wrap: false });
  });
}

/** "Description" legend chips under the timeline, one per accent colour. */
function legendRow(slide, xs, y) {
  const swatches = [C.Y1, C.Y2, C.Y3, C.Y4, C.Y5, C.Y6];
  xs.forEach(function (x, i) {
    S(slide, 'ellipse', { x: x, y: y, w: 0.206, h: 0.2, fill: swatches[i] });
    T(slide, 'Description', {
      x: x + 0.278, y: y + 0.015, w: 0.848, h: 0.168,
      fontSize: 10, bold: true, fontFace: F.HEAD, wrap: false,
    });
  });
}
// ---------------------------------------------------------------------------
// Slide 1 — Cover
// ---------------------------------------------------------------------------
function slide01(deck) {
  const s = newSlide(deck, false);
  photo(s, 0.583, 0.583, 12.167, 6.333);
  S(s, 'rect', { x:0.564, y:0.583, w:12.186, h:6.333, fill:{ color: C.BG, transparency: 20 } });
  T(s, 'SPORTS', { x:4.309, y:3.223, w:4.716, h:1.346, fontSize:80, bold:true, fontFace:F.HEAD, color:C.Y1, align:'center', wrap:false, margin:0 });
  T(s, 'HEALTH', { x:4.268, y:2.203, w:4.796, h:1.346, fontSize:80, bold:true, fontFace:F.HEAD, align:'center', wrap:false, margin:0 });
  T(s, 'PLACEHOLDER', { x:9.204, y:5.634, w:3.525, h:0.976, lineSpacingMultiple:1.5 });
  navbar(s, 0.861, 0.805);
  T(s, 'Presentation template', { x:4.341, y:4.317, w:2.428, h:0.505, fontSize:24, fontFace:F.SCRIPT, wrap:false });
}

// ---------------------------------------------------------------------------
// Slide 2 — Turn fat into fit
// ---------------------------------------------------------------------------
function slide02(deck) {
  const s = newSlide(deck);
  T(s, ['TURN FAT', 'INTO FIT'].join(BR), { x:0.887, y:1.702, w:2.858, h:1.313, fontSize:36, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L1, { x:0.887, y:3.147, w:5.185, h:1.279, align:'justify', lineSpacingMultiple:1.5 });
  T(s, 'Health Sports', { x:10.52, y:6.705, w:2.211, h:0.64, fontSize:32, fontFace:F.SCRIPT, color:C.Y1, wrap:false });
  S(s, 'rect', { x:7.708, y:4.374, w:0.318, h:2.226, fill:C.Y1 });
  buttons(s, 1.011, 4.857);
  ghost(s, 0.461, 6.347, 6.099, 0.909, 48);
  photo(s, 8.025, 0, 4.69, 6.6);
}

// ---------------------------------------------------------------------------
// Slide 3 — Just do it
// ---------------------------------------------------------------------------
function slide03(deck) {
  const s = newSlide(deck);
  quoteMark(s, 0.979, 4.828, 0.351, 0.297);
  S(s, 'rect', { x:8.222, y:0, w:0.239, h:7.5, fill:C.Y1 });
  T(s, 'JUST DO IT', { x:0.855, y:1.321, w:3.438, h:0.774, fontSize:40, bold:true, fontFace:F.HEAD, color:C.Y1, wrap:false });
  T(s, L1, { x:0.855, y:3.052, w:5.811, h:1.279, align:'justify', lineSpacingMultiple:1.5 });
  T(s, ['START CHANGE', 'OF LIFE PATTERNS'].join(BR), { x:0.855, y:2.021, w:2.97, h:0.774, fontSize:20, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L8, { x:1.5, y:4.769, w:4.793, h:0.875, bold:true, fontFace:F.HEAD, align:'justify', lineSpacingMultiple:1.5, margin:0 });
  ghost(s, 0.461, 6.347, 6.099, 0.909, 48);
  photo(s, 8.461, 0, 4.27, 7.5);
}

// ---------------------------------------------------------------------------
// Slide 4 — What features
// ---------------------------------------------------------------------------
function slide04(deck) {
  const s = newSlide(deck);
  S(s, 'rect', { x:0, y:5.298, w:6.336, h:1.077, fill:C.Y1 });
  T(s, 'WHAT FEATURES', { x:8.071, y:1.251, w:4.315, h:0.64, fontSize:32, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, L1, { x:6.892, y:2.028, w:5.494, h:1.279, align:'right', lineSpacingMultiple:1.5 });
  T(s, ['HEALTHY', 'LIFE STYLE'].join(BR), { x:4.438, y:5.483, w:1.65, h:0.707, fontSize:18, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, L3, { x:3.513, y:2.139, w:2.268, h:0.625, fontSize:11, align:'right', lineSpacingMultiple:1.5 });
  T(s, 'VECTOR BASED', { x:4.357, y:1.853, w:1.424, h:0.278, fontSize:10.5, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, L3, { x:3.513, y:3.607, w:2.268, h:0.625, fontSize:11, align:'right', lineSpacingMultiple:1.5 });
  T(s, 'SCHEME COLOR', { x:4.322, y:3.321, w:1.459, h:0.278, fontSize:10.5, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, L3, { x:0.474, y:2.139, w:2.268, h:0.625, fontSize:11, align:'right', lineSpacingMultiple:1.5 });
  T(s, 'FINE LAYOUT', { x:1.504, y:1.853, w:1.238, h:0.278, fontSize:10.5, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, L3, { x:0.474, y:3.607, w:2.268, h:0.625, fontSize:11, align:'right', lineSpacingMultiple:1.5 });
  T(s, 'PLCAEHOLDER', { x:1.364, y:3.321, w:1.378, h:0.278, fontSize:10.5, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, L9, { x:0.485, y:5.524, w:3.705, h:0.625, fontSize:11, align:'right', lineSpacingMultiple:1.5 });
  photo(s, 6.627, 4.103, 6.707, 3.397);
}

// ---------------------------------------------------------------------------
// Slide 5 — Content today
// ---------------------------------------------------------------------------
function slide05(deck) {
  const s = newSlide(deck);
  photo(s, 4.396, 0.865, 2.231, 3.167);
  T(s, 'CONTENT TODAY', { x:8.076, y:1.098, w:4.29, h:0.64, fontSize:32, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, L1, { x:7.181, y:1.894, w:5.185, h:1.279, align:'right', lineSpacingMultiple:1.5 });
  T(s, L3, { x:10.098, y:4.948, w:2.268, h:0.625, fontSize:11, align:'right', lineSpacingMultiple:1.5 });
  T(s, 'ABOUT US', { x:11.321, y:4.659, w:1.045, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, L3, { x:5.493, y:4.948, w:2.268, h:0.625, fontSize:11, align:'right', lineSpacingMultiple:1.5 });
  T(s, 'TRAINER TEAM', { x:6.333, y:4.659, w:1.427, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, L3, { x:0.887, y:4.951, w:2.268, h:0.625, fontSize:11, align:'right', lineSpacingMultiple:1.5 });
  T(s, 'IMAGE GALLERY', { x:1.628, y:4.659, w:1.527, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  S(s, 'rect', { x:4.396, y:3.638, w:2.231, h:0.393, fill:{ color: C.Y1, transparency: 20 } });
  T(s, 'HEALTH SPORTS', { x:4.463, y:3.702, w:2.097, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, charSpacing:3, align:'center', wrap:false });
  buttons(s, 9.63, 3.585);
  ghost(s, 0.524, 6.25, 7.577, 1.111, 60);
  photo(s, 0, 0.865, 4.25, 3.167);
}

// ---------------------------------------------------------------------------
// Slide 6 — Hello welcome
// ---------------------------------------------------------------------------
function slide06(deck) {
  const s = newSlide(deck);
  T(s, 'HELLO WELCOME', { x:0.646, y:1.454, w:4.448, h:0.64, fontSize:32, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L1, { x:0.646, y:2.232, w:5.179, h:1.279, align:'justify', lineSpacingMultiple:1.5 });
  S(s, 'rect', { x:9.866, y:5.073, w:3.467, h:0.434, fill:C.Y1 });
  T(s, L5, { x:1.554, y:4.395, w:3.92, h:0.802, fontSize:11, align:'justify', lineSpacingMultiple:1.5, margin:0 });
  T(s, 'JASON MARTOP', { x:1.457, y:4.109, w:1.515, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, wrap:false });
  ghost(s, 0.305, 6.347, 6.099, 0.909, 48);
  lineIcon(s, 0.896, 4.076, 0.35, 0.349, C.Y1);
  photo(s, 9.866, 0, 3.467, 4.952);
  photo(s, 6.92, 3.75, 2.82, 3.75);
}

// ---------------------------------------------------------------------------
// Slide 7 — About us
// ---------------------------------------------------------------------------
function slide07(deck) {
  const s = newSlide(deck);
  T(s, 'ABOUT US', { x:9.701, y:1.253, w:2.665, h:0.64, fontSize:32, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, L1, { x:7.181, y:3.238, w:5.185, h:1.279, align:'right', lineSpacingMultiple:1.5 });
  T(s, L8, { x:7.096, y:2.066, w:5.185, h:0.875, bold:true, fontFace:F.HEAD, align:'right', lineSpacingMultiple:1.5, margin:0 });
  T(s, 'Health Sports', { x:10.053, y:4.815, w:2.211, h:0.64, fontSize:32, fontFace:F.SCRIPT, color:C.Y1, wrap:false });
  ghost(s, 6.471, 6.219, 6.099, 0.909, 48);
  S(s, 'rect', { x:5.369, y:0.846, w:0.361, h:1.838, fill:C.Y1 });
  photo(s, 0.583, 0.846, 4.786, 3.892);
  photo(s, 3.02, 4.815, 2.349, 2.084);
  photo(s, 0.583, 4.815, 2.349, 2.084);
}

// ---------------------------------------------------------------------------
// Slide 8 — The history
// ---------------------------------------------------------------------------
function slide08(deck) {
  const s = newSlide(deck);
  photo(s, 9.382, 5.016, 3.349, 2.484);
  photo(s, 6.81, 0, 4.349, 4.857);
  ghost(s, 0.305, 6.347, 6.099, 0.909, 48);
  T(s, ['THE HISTORY', 'ABOUT US'].join(BR), { x:0.646, y:1.2, w:3.354, h:1.178, fontSize:32, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L1, { x:0.646, y:3.133, w:5.179, h:1.279, align:'justify', lineSpacingMultiple:1.5 });
  T(s, 'Health Sports', { x:0.683, y:2.309, w:2.211, h:0.64, fontSize:32, fontFace:F.SCRIPT, color:C.Y1, wrap:false });
  T(s, 'PLACEHOLDER', { x:0.646, y:4.679, w:5.179, h:0.976, align:'justify', lineSpacingMultiple:1.5 });
  T(s, '2020', { x:10.833, y:3.675, w:1.625, h:0.774, fontSize:40, bold:true, fontFace:F.HEAD, align:'center', wrap:false, rotate:90 });
  T(s, 'START AT', { x:10.731, y:2.319, w:1.618, h:0.438, fontSize:20, bold:true, fontFace:F.HEAD, color:C.Y1, align:'center', wrap:false, rotate:90 });
}

// ---------------------------------------------------------------------------
// Slide 9 — Vision & mission
// ---------------------------------------------------------------------------
function slide09(deck) {
  const s = newSlide(deck);
  photo(s, 6.627, 0.583, 6.707, 6.333);
  T(s, 'VISION & MISSION', { x:0.715, y:1.384, w:3.938, h:0.572, fontSize:28, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L1, { x:0.715, y:2.1, w:5.185, h:1.279, lineSpacingMultiple:1.5 });
  T(s, L8, { x:0.801, y:3.753, w:5.185, h:0.875, bold:true, fontFace:F.HEAD, lineSpacingMultiple:1.5, margin:0 });
  buttons(s, 0.868, 5.127);
  S(s, 'rect', { x:6.627, y:5.127, w:6.707, h:1.79, fill:{ color: C.Y1, transparency: 10 } });
  ghost(s, 0.501, 6.347, 5.614, 0.841, 44);
  T(s, L6, { x:7.204, y:5.91, w:2.363, h:0.525, fontSize:11, align:'justify', lineSpacingMultiple:1.5, margin:0 });
  T(s, 'MISSION 01', { x:7.107, y:5.623, w:1.11, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L6, { x:10.469, y:5.91, w:2.363, h:0.525, fontSize:11, align:'justify', lineSpacingMultiple:1.5, margin:0 });
  T(s, 'MISSION 01', { x:10.371, y:5.623, w:1.11, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, wrap:false });
}

// ---------------------------------------------------------------------------
// Slide 10 — What we offer (stats)
// ---------------------------------------------------------------------------
function slide10(deck) {
  const s = newSlide(deck);
  T(s, 'WHAT WE OFFER', { x:8.059, y:1.47, w:4.369, h:0.64, fontSize:32, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, L1, { x:7.243, y:2.266, w:5.185, h:1.279, align:'right', lineSpacingMultiple:1.5 });
  T(s, L10, { x:0.433, y:2.286, w:2.348, h:0.802, fontSize:11, align:'right', lineSpacingMultiple:1.5, margin:0 });
  T(s, 'OFFER NAME', { x:1.584, y:1.6, w:1.282, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, L10, { x:0.433, y:5.154, w:2.348, h:0.802, fontSize:11, align:'right', lineSpacingMultiple:1.5, margin:0 });
  T(s, 'OFFER NAME', { x:1.584, y:4.469, w:1.282, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, '60%', { x:1.928, y:1.849, w:0.938, h:0.505, fontSize:24, bold:true, fontFace:F.HEAD, color:C.Y1, align:'right', wrap:false });
  T(s, '80%', { x:1.919, y:4.718, w:0.947, h:0.505, fontSize:24, bold:true, fontFace:F.HEAD, color:C.Y2, align:'right', wrap:false });
  ghost(s, 7.309, 3.827, 5.119, 0.774, 40);
  T(s, 'OFFER DESCRIPTION', { x:10.518, y:4.779, w:1.909, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, color:C.Y1, align:'right', wrap:false });
  T(s, 'PLACEHOLDER', { x:7.442, y:5.065, w:4.868, h:0.802, fontSize:11, align:'right', lineSpacingMultiple:1.5, margin:0 });
  photo(s, 3.514, 0.955, 2.882, 2.882);
  photo(s, 3.514, 3.997, 2.882, 2.882);
}

// ---------------------------------------------------------------------------
// Slide 11 — What we offer (cards)
// ---------------------------------------------------------------------------
function slide11(deck) {
  const s = newSlide(deck);
  T(s, 'WHAT WE OFFER', { x:0.764, y:1.703, w:4.369, h:0.64, fontSize:32, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L1, { x:0.764, y:2.499, w:5.185, h:1.279, lineSpacingMultiple:1.5 });
  S(s, 'rect', { x:7.401, y:1.261, w:1.359, h:1.359, fill:C.Y1 });
  S(s, 'rect', { x:7.401, y:3.21, w:1.359, h:1.359, fill:C.Y2 });
  S(s, 'rect', { x:7.401, y:5.159, w:1.359, h:1.359, fill:C.Y3 });
  buttons(s, 0.904, 4.261);
  T(s, 'OFFER DESCRIPTION', { x:9.446, y:1.397, w:1.909, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L4, { x:9.56, y:1.683, w:3.009, h:0.802, fontSize:11, lineSpacingMultiple:1.5, margin:0 });
  T(s, 'OFFER DESCRIPTION', { x:9.446, y:3.346, w:1.909, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L4, { x:9.56, y:3.632, w:3.009, h:0.802, fontSize:11, lineSpacingMultiple:1.5, margin:0 });
  T(s, 'OFFER DESCRIPTION', { x:9.446, y:5.295, w:1.909, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L4, { x:9.56, y:5.581, w:3.009, h:0.802, fontSize:11, lineSpacingMultiple:1.5, margin:0 });
  ghost(s, 0.305, 6.347, 6.099, 0.909, 48);
  lineIcon(s, 7.734, 3.544, 0.691, 0.691, C.INK);
  lineIcon(s, 7.734, 1.595, 0.691, 0.691, C.INK);
  lineIcon(s, 7.734, 5.493, 0.691, 0.691, C.INK);
}

// ---------------------------------------------------------------------------
// Slide 12 — Special offer from us
// ---------------------------------------------------------------------------
function slide12(deck) {
  const s = newSlide(deck);
  photo(s, 6.627, 4.79, 6.104, 2.127);
  S(s, 'rect', { x:6.627, y:4.79, w:6.104, h:2.127, fill:{ color: C.Y1, transparency: 20 } });
  T(s, ['SPECIAL OFFER', 'FROM US'].join(BR), { x:0.623, y:1.384, w:3.92, h:1.178, fontSize:32, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L1, { x:0.623, y:2.69, w:5.185, h:1.279, lineSpacingMultiple:1.5 });
  T(s, L5, { x:1.462, y:4.696, w:3.92, h:0.802, fontSize:11, align:'justify', lineSpacingMultiple:1.5, margin:0 });
  T(s, 'OFFER NAME', { x:1.364, y:4.41, w:1.282, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, wrap:false });
  lineIcon(s, 0.787, 4.378, 0.418, 0.418, C.INK);
  lineIcon(s, 7.161, 5.279, 0.392, 0.392, C.INK);
  T(s, L5, { x:7.859, y:5.584, w:3.92, h:0.802, fontSize:11, align:'justify', lineSpacingMultiple:1.5, margin:0 });
  T(s, 'OFFER NAME', { x:7.762, y:5.298, w:1.282, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, wrap:false });
  ghost(s, 0.47, 6.347, 5.614, 0.841, 44);
  photo(s, 6.627, 0.583, 6.104, 4.066);
}

// ---------------------------------------------------------------------------
// Slide 13 — Sport benefits
// ---------------------------------------------------------------------------
function slide13(deck) {
  const s = newSlide(deck);
  S(s, 'rect', { x:6.43, y:2.186, w:1.492, h:1.746, fill:C.Y1 });
  T(s, ['SPORT BENEFITS', 'FOR THE BODY'].join(BR), { x:8.382, y:4.315, w:4.234, h:1.178, fontSize:32, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, L11, { x:7.431, y:5.621, w:5.185, h:0.976, align:'right', lineSpacingMultiple:1.5 });
  T(s, L7, { x:0.73, y:3.629, w:3.92, h:0.875, align:'right', lineSpacingMultiple:1.5, margin:0 });
  T(s, 'BENEFIT TWO', { x:3.432, y:3.342, w:1.341, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, '02', { x:4.777, y:3.342, w:0.398, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, color:C.Y2, wrap:false });
  T(s, L7, { x:0.73, y:1.636, w:3.92, h:0.875, align:'right', lineSpacingMultiple:1.5, margin:0 });
  T(s, 'BENEFIT ONE', { x:3.477, y:1.35, w:1.296, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, L7, { x:0.73, y:5.621, w:3.92, h:0.875, align:'right', lineSpacingMultiple:1.5, margin:0 });
  T(s, 'BENEFIT THREE', { x:3.295, y:5.335, w:1.478, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, '01', { x:4.777, y:1.35, w:0.367, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, color:C.Y1, wrap:false });
  T(s, '03', { x:4.777, y:5.335, w:0.398, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, color:C.Y3, wrap:false });
  photo(s, 6.627, 0, 6.707, 3.75);
}

// ---------------------------------------------------------------------------
// Slide 14 — Let's see what the target
// ---------------------------------------------------------------------------
function slide14(deck) {
  const s = newSlide(deck);
  S(s, 'rect', { x:10.749, y:5.008, w:2.164, h:2.092, fill:C.Y1 });
  T(s, ['LET\u2019S SEE WHAT', 'THE TARGET'].join(BR), { x:0.623, y:1.335, w:4.103, h:1.178, fontSize:32, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L1, { x:0.623, y:2.641, w:5.185, h:1.279, lineSpacingMultiple:1.5 });
  T(s, L2, { x:8.04, y:1.636, w:3.92, h:0.572, lineSpacingMultiple:1.5, margin:0 });
  T(s, 'TARGET ONE', { x:7.94, y:1.35, w:1.247, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L2, { x:8.04, y:2.983, w:3.92, h:0.572, lineSpacingMultiple:1.5, margin:0 });
  T(s, 'TARGET TWO', { x:7.94, y:2.696, w:1.292, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, wrap:false });
  ghost(s, 0.47, 6.347, 5.614, 0.841, 44);
  T(s, 'Health Sports', { x:0.683, y:4.195, w:2.211, h:0.64, fontSize:32, fontFace:F.SCRIPT, color:C.Y1, wrap:false });
  T(s, '01', { x:7.557, y:1.35, w:0.367, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, color:C.Y1, align:'right', wrap:false });
  T(s, '02', { x:7.525, y:2.696, w:0.398, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, color:C.Y2, align:'right', wrap:false });
  photo(s, 6.938, 4.062, 5.792, 2.855);
}

// ---------------------------------------------------------------------------
// Slide 15 — Health sports plan
// ---------------------------------------------------------------------------
function slide15(deck) {
  const s = newSlide(deck);
  photo(s, 0, 0, 6.09, 7.5);
  S(s, 'rect', { x:0, y:0, w:6.09, h:7.5, fill:{ color: C.BG, transparency: 30 } });
  T(s, ['HEALTH', 'SPORTS PLAN'].join(BR), { x:8.909, y:1.01, w:3.519, h:1.178, fontSize:32, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, L1, { x:6.943, y:2.348, w:5.485, h:1.279, align:'right', lineSpacingMultiple:1.5 });
  S(s, 'rect', { x:4.248, y:1.217, w:1.245, h:1.245, fill:C.Y1 });
  T(s, L4, { x:0.688, y:1.472, w:3.308, h:0.875, align:'right', lineSpacingMultiple:1.5, margin:0 });
  T(s, 'PLAN ONE', { x:3.068, y:1.186, w:1.05, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  S(s, 'rect', { x:4.248, y:3.2, w:1.245, h:1.245, fill:C.Y2 });
  T(s, L4, { x:0.688, y:3.455, w:3.308, h:0.875, align:'right', lineSpacingMultiple:1.5, margin:0 });
  T(s, 'PLAN TWO', { x:3.023, y:3.169, w:1.096, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  S(s, 'rect', { x:4.248, y:5.183, w:1.245, h:1.245, fill:C.Y3 });
  T(s, L4, { x:0.688, y:5.438, w:3.308, h:0.875, align:'right', lineSpacingMultiple:1.5, margin:0 });
  T(s, 'PLAN THREE', { x:2.886, y:5.152, w:1.233, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  quoteMark(s, 11.904, 4.072, 0.351, 0.297);
  T(s, 'PLACEHOLDER', { x:7.365, y:4.013, w:4.37, h:0.875, bold:true, fontFace:F.HEAD, align:'right', lineSpacingMultiple:1.5, margin:0 });
  T(s, 'Health Sports', { x:10.216, y:5.152, w:2.211, h:0.64, fontSize:32, fontFace:F.SCRIPT, color:C.Y1, align:'right', wrap:false });
  ghost(s, 7.476, 6.511, 5.614, 0.841, 44);
  lineIcon(s, 4.582, 3.534, 0.575, 0.575, C.INK);
  lineIcon(s, 4.486, 1.455, 0.767, 0.767, C.INK);
  lineIcon(s, 4.582, 5.517, 0.575, 0.575, C.INK);
}

// ---------------------------------------------------------------------------
// Slide 16 — Meet the trainer team
// ---------------------------------------------------------------------------
function slide16(deck) {
  const s = newSlide(deck);
  ghost(s, -0.988, 6.53, 5.614, 0.841, 44);
  T(s, ['MEET THE', 'TRAINER TEAM'].join(BR), { x:0.623, y:1.722, w:3.769, h:1.178, fontSize:32, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L1, { x:0.623, y:3.028, w:5.185, h:1.279, lineSpacingMultiple:1.5 });
  buttons(s, 0.793, 4.712);
  S(s, 'rect', { x:6.741, y:6.84, w:6.593, h:0.66, fill:C.Y1 });
  photo(s, 6.741, 0.662, 3.695, 6.01);
  photo(s, 10.603, 0.662, 2.731, 6.01);
}

// ---------------------------------------------------------------------------
// Slide 17 — The energic trainer
// ---------------------------------------------------------------------------
function slide17(deck) {
  const s = newSlide(deck);
  T(s, ['THE ENERGIC', 'TRAINER'].join(BR), { x:9.231, y:1.104, w:3.387, h:1.178, fontSize:32, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, L11, { x:7.238, y:2.441, w:5.38, h:0.976, align:'right', lineSpacingMultiple:1.5 });
  T(s, L12, { x:0.626, y:4.252, w:2.373, h:0.572, align:'center', lineSpacingMultiple:1.5, margin:0 });
  T(s, L12, { x:3.397, y:4.252, w:2.373, h:0.572, align:'center', lineSpacingMultiple:1.5, margin:0 });
  ghost(s, 0.418, 6.267, 5.614, 0.841, 44);
  S(s, 'rect', { x:0.583, y:3.714, w:2.458, h:0.333, fill:C.Y1 });
  S(s, 'rect', { x:3.354, y:3.714, w:2.458, h:0.333, fill:C.Y1 });
  T(s, 'MIKE KLORI', { x:1.021, y:3.738, w:1.583, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, charSpacing:3, align:'center', wrap:false });
  T(s, 'JHON STUF', { x:3.834, y:3.738, w:1.499, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, charSpacing:3, align:'center', wrap:false });
  dot(s, 1.712, 5.112, 0.201, 0.201, C.Y1);
  dot(s, 2.084, 5.131, 0.201, 0.164, C.Y1);
  dot(s, 1.339, 5.113, 0.201, 0.199, C.Y1);
  dot(s, 4.483, 5.112, 0.201, 0.201, C.Y1);
  dot(s, 4.855, 5.131, 0.201, 0.164, C.Y1);
  dot(s, 4.11, 5.113, 0.201, 0.199, C.Y1);
  photo(s, 7.302, 3.937, 5.429, 2.98);
  photo(s, 3.354, 1.145, 2.458, 2.458);
  photo(s, 0.583, 1.145, 2.458, 2.458);
}

// ---------------------------------------------------------------------------
// Slide 18 — Meet our special trainer
// ---------------------------------------------------------------------------
function slide18(deck) {
  const s = newSlide(deck);
  T(s, ['MEET OUR', 'SPECIAL TRAINER'].join(BR), { x:0.577, y:1.261, w:4.427, h:1.178, fontSize:32, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L1, { x:0.577, y:2.567, w:5.185, h:1.279, lineSpacingMultiple:1.5 });
  T(s, L2, { x:10.113, y:1.357, w:2.452, h:0.875, align:'justify', lineSpacingMultiple:1.5, margin:0 });
  T(s, 'GERALD MOSES', { x:10.03, y:1.07, w:1.497, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, wrap:false });
  dot(s, 10.529, 2.462, 0.201, 0.201, C.Y1);
  dot(s, 10.901, 2.481, 0.201, 0.164, C.Y1);
  dot(s, 10.156, 2.463, 0.201, 0.199, C.Y1);
  T(s, '@Geraldmoses', { x:10.03, y:2.705, w:1.252, h:0.286, fontSize:11, wrap:false });
  T(s, L2, { x:10.113, y:4.795, w:2.452, h:0.875, align:'justify', lineSpacingMultiple:1.5, margin:0 });
  T(s, 'JASON GROOM', { x:10.03, y:4.509, w:1.44, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, wrap:false });
  dot(s, 10.529, 5.9, 0.201, 0.201, C.Y1);
  dot(s, 10.901, 5.919, 0.201, 0.164, C.Y1);
  dot(s, 10.156, 5.901, 0.201, 0.199, C.Y1);
  T(s, '@Jasongroom', { x:10.03, y:6.143, w:1.212, h:0.286, fontSize:11, wrap:false });
  ghost(s, 0.326, 6.298, 5.614, 0.841, 44);
  S(s, 'rect', { x:0, y:4.891, w:6.146, h:1.077, fill:C.Y1 });
  T(s, ['HEALTHY', 'LIFE STYLE'].join(BR), { x:0.248, y:5.076, w:1.65, h:0.707, fontSize:18, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L9, { x:2.083, y:5.117, w:3.705, h:0.625, fontSize:11, lineSpacingMultiple:1.5 });
  photo(s, 7.069, 0.784, 2.494, 2.494);
  photo(s, 7.069, 4.223, 2.494, 2.494);
}

// ---------------------------------------------------------------------------
// Slide 19 — Sammy Moralio
// ---------------------------------------------------------------------------
function slide19(deck) {
  const s = newSlide(deck);
  T(s, 'SAMMY MORALIO', { x:0.671, y:1.285, w:4.394, h:0.64, fontSize:32, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L1, { x:0.671, y:2.037, w:5.185, h:1.279, lineSpacingMultiple:1.5 });
  S(s, 'rect', { x:0.746, y:4.83, w:4.426, h:0.139, fill:C.TRACK });
  S(s, 'rect', { x:0.746, y:4.83, w:3.573, h:0.139, fill:C.Y1 });
  T(s, '70%', { x:5.284, y:4.755, w:0.539, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, color:C.Y1, wrap:false });
  T(s, 'Responsibility', { x:0.692, y:4.477, w:1.243, h:0.303, wrap:false });
  S(s, 'rect', { x:0.746, y:5.559, w:4.426, h:0.139, fill:C.TRACK });
  S(s, 'rect', { x:0.746, y:5.559, w:3.573, h:0.139, fill:C.Y2 });
  T(s, '70%', { x:5.284, y:5.485, w:0.539, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, color:C.Y2, wrap:false });
  T(s, 'Responsibility', { x:0.692, y:5.207, w:1.243, h:0.303, wrap:false });
  buttons(s, 0.825, 3.765);
  T(s, 'Special Trainer', { x:6.58, y:1.34, w:2.085, h:0.572, fontSize:28, fontFace:F.SCRIPT, color:C.Y1, wrap:false, rotate:270 });
  ghost(s, 0.069, 6.537, 5.614, 0.841, 44);
  S(s, 'rect', { x:12.463, y:0.583, w:0.339, h:6.333, fill:C.Y1 });
  photo(s, 7.96, 0.583, 4.394, 6.333);
}

// ---------------------------------------------------------------------------
// Slide 20 — Break section
// ---------------------------------------------------------------------------
function slide20(deck) {
  const s = newSlide(deck);
  T(s, ['BREAK', 'SECTION'].join(BR), { x:0.61, y:1.178, w:3.719, h:1.919, fontSize:54, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L1, { x:0.61, y:3.351, w:5.185, h:1.279, lineSpacingMultiple:1.5 });
  S(s, 'rect', { x:0, y:5.84, w:6.421, h:1.077, fill:C.Y1 });
  T(s, ['HEALTHY', 'LIFE STYLE'].join(BR), { x:0.248, y:6.025, w:1.65, h:0.707, fontSize:18, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L9, { x:2.171, y:6.065, w:3.705, h:0.625, fontSize:11, lineSpacingMultiple:1.5 });
  photo(s, 6.627, 0.583, 6.707, 6.333);
}

// ---------------------------------------------------------------------------
// Slide 21 — Let's see our gallery
// ---------------------------------------------------------------------------
function slide21(deck) {
  const s = newSlide(deck);
  T(s, ['LET\u2019S SEE', 'OUR GALLERY'].join(BR), { x:8.871, y:1.073, w:3.557, h:1.178, fontSize:32, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, L1, { x:7.243, y:2.476, w:5.185, h:1.279, align:'right', lineSpacingMultiple:1.5 });
  T(s, L5, { x:8.208, y:5.106, w:4.125, h:0.802, fontSize:11, align:'right', lineSpacingMultiple:1.5, margin:0 });
  T(s, 'GALLERY DESCRIPTION', { x:10.319, y:4.819, w:2.109, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  lineIcon(s, 11.786, 4.24, 0.475, 0.475, C.INK);
  ghost(s, 7.337, 6.511, 5.614, 0.841, 44);
  photo(s, 0.583, 3.944, 5.885, 2.838);
  photo(s, 3.631, 0.912, 2.838, 2.838);
  photo(s, 0.583, 0.912, 2.838, 2.838);
}

// ---------------------------------------------------------------------------
// Slide 22 — What we have now?
// ---------------------------------------------------------------------------
function slide22(deck) {
  const s = newSlide(deck);
  T(s, ['WHAT WE', 'HAVE NOW ?'].join(BR), { x:0.577, y:1.334, w:3.326, h:1.178, fontSize:32, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L1, { x:0.577, y:2.64, w:5.185, h:1.279, lineSpacingMultiple:1.5 });
  ghost(s, 0.254, 6.445, 5.614, 0.841, 44);
  T(s, L6, { x:6.924, y:5.855, w:2.516, h:0.572, lineSpacingMultiple:1.5, margin:0 });
  T(s, 'DESCRIPTION', { x:6.801, y:5.569, w:1.322, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L6, { x:10.065, y:5.855, w:2.516, h:0.572, lineSpacingMultiple:1.5, margin:0 });
  T(s, 'DESCRIPTION', { x:9.942, y:5.569, w:1.322, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, wrap:false });
  buttons(s, 0.741, 4.326);
  photo(s, 6.65, 0.988, 2.939, 4.116);
  photo(s, 9.792, 0.988, 2.939, 4.116);
}

// ---------------------------------------------------------------------------
// Slide 23 — Come and start with us
// ---------------------------------------------------------------------------
function slide23(deck) {
  const s = newSlide(deck);
  T(s, L1, { x:7.354, y:1.313, w:5.185, h:1.279, lineSpacingMultiple:1.5 });
  T(s, '2020', { x:6.239, y:1.632, w:1.341, h:0.64, fontSize:32, bold:true, fontFace:F.HEAD, color:C.Y1, align:'center', wrap:false, rotate:270 });
  S(s, 'rect', { x:0, y:6.809, w:13.333, h:0.191, fill:C.Y1 });
  T(s, ['COME AND', 'START WITH US'].join(BR), { x:0.577, y:1.109, w:3.917, h:1.178, fontSize:32, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, 'Health Sports', { x:0.528, y:2.199, w:2.211, h:0.64, fontSize:32, fontFace:F.SCRIPT, color:C.Y1, align:'right', wrap:false });
  photo(s, 6.743, 3.377, 3.219, 3.3);
  photo(s, 10.114, 3.377, 3.219, 3.3);
  photo(s, 3.371, 3.377, 3.219, 3.3);
  photo(s, 0, 3.377, 3.219, 3.3);
}

// ---------------------------------------------------------------------------
// Slide 24 — Gallery mockup
// ---------------------------------------------------------------------------
function slide24(deck) {
  const s = newSlide(deck);
  device(s, 6.967, 1.233, 5.398, 4.099);
  S(s, 'rect', { x:10.767, y:0, w:2.566, h:7.5, fill:C.Y1 });
  T(s, ['GALLERY', 'MOCKUP'].join(BR), { x:0.758, y:1.563, w:2.369, h:1.178, fontSize:32, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L1, { x:0.758, y:2.869, w:5.185, h:1.279, lineSpacingMultiple:1.5 });
  buttons(s, 0.887, 4.597);
  ghost(s, 0.254, 6.445, 5.614, 0.841, 44);
  photo(s, 7.077, 1.343, 5.178, 3.879);
}

// ---------------------------------------------------------------------------
// Slide 25 — Gallery mockup in smart phone
// ---------------------------------------------------------------------------
function slide25(deck) {
  const s = newSlide(deck);
  device(s, 0.642, 0.709, 2.767, 5.676);
  device(s, 3.685, 2.459, 2.767, 5.151);
  S(s, 'rect', { x:0, y:1.132, w:6.405, h:2.677, fill:C.Y1 });
  T(s, ['GALLERY MOCKUP', 'IN SMART PHONE'].join(BR), { x:7.442, y:1.165, w:4.644, h:1.178, fontSize:32, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L1, { x:7.442, y:2.471, w:5.185, h:1.279, lineSpacingMultiple:1.5 });
  T(s, L7, { x:7.549, y:4.489, w:3.92, h:0.875, lineSpacingMultiple:1.5, margin:0 });
  T(s, 'DESCRIPTION', { x:7.437, y:4.203, w:1.322, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, wrap:false });
  ghost(s, 6.932, 6.341, 6.099, 0.909, 48);
  photo(s, 3.795, 2.569, 2.547, 4.931, 'round');
  photo(s, 0.752, 0.819, 2.547, 5.456, 'round');
}

// ---------------------------------------------------------------------------
// Slide 26 — Our gallery in tab mockup
// ---------------------------------------------------------------------------
function slide26(deck) {
  const s = newSlide(deck);
  device(s, 2.595, 0.764, 3.464, 4.542);
  photo(s, 0.583, 0.762, 2.479, 6.155);
  S(s, 'rect', { x:0.583, y:0.762, w:2.479, h:6.155, fill:{ color: C.GHOST, transparency: 20 } });
  T(s, [{ text: 'OUR GALLERY', options: { breakLine:true } }, { text: 'IN TAB MOCKUP', options: { color:C.INK } }], { x:8.35, y:1.073, w:4.078, h:1.178, fontSize:32, bold:true, fontFace:F.HEAD, color:C.Y1, align:'right', wrap:false });
  T(s, L1, { x:7.243, y:2.476, w:5.185, h:1.279, align:'right', lineSpacingMultiple:1.5 });
  T(s, L5, { x:8.208, y:4.534, w:4.125, h:0.802, fontSize:11, align:'right', lineSpacingMultiple:1.5, margin:0 });
  T(s, 'GALLERY DESCRIPTION', { x:10.319, y:4.248, w:2.109, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  ghost(s, 6.932, 6.341, 6.099, 0.909, 48);
  photo(s, 2.705, 0.874, 3.244, 4.322, 'round');
}

// ---------------------------------------------------------------------------
// Slide 27 — Infographic — histogram
// ---------------------------------------------------------------------------
function slide27(deck) {
  const s = newSlide(deck);
  histogram(s);
  T(s, 'INFOGRAPHIC', { x:5.095, y:0.501, w:3.144, h:0.572, fontSize:28, bold:true, fontFace:F.HEAD, align:'center', wrap:false });
  T(s, 'PLACEHOLDER', { x:0.708, y:5.895, w:5.645, h:0.673, align:'center', lineSpacingMultiple:1.5 });
  T(s, L1, { x:7.243, y:2.815, w:5.185, h:1.279, align:'right', lineSpacingMultiple:1.5 });
  T(s, 'GALLERY DESCRIPTION', { x:10.319, y:2.013, w:2.109, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, '60%', { x:11.241, y:2.263, w:1.187, h:0.64, fontSize:32, bold:true, fontFace:F.HEAD, color:C.Y1, align:'right', wrap:false });
  lineIcon(s, 12.071, 4.583, 0.143, 0.364, C.Y1);
  lineIcon(s, 12.078, 5.722, 0.129, 0.364, C.Y2);
  T(s, L13, { x:7.653, y:4.809, w:4.125, h:0.525, fontSize:11, align:'right', lineSpacingMultiple:1.5, margin:0 });
  T(s, '25.6500', { x:11.048, y:4.523, w:0.824, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, L13, { x:7.653, y:5.954, w:4.125, h:0.525, fontSize:11, align:'right', lineSpacingMultiple:1.5, margin:0 });
  T(s, '25.6500', { x:11.048, y:5.668, w:0.824, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
}

// ---------------------------------------------------------------------------
// Slide 28 — Infographic — social
// ---------------------------------------------------------------------------
function slide28(deck) {
  const s = newSlide(deck);
  T(s, 'INFOGRAPHIC', { x:5.095, y:0.501, w:3.144, h:0.572, fontSize:28, bold:true, fontFace:F.HEAD, align:'center', wrap:false });
  S(s, 'ellipse', { x:8.065, y:1.974, w:2.043, h:2.043, fill:C.Y1 });
  S(s, 'ellipse', { x:10.529, y:1.974, w:2.043, h:2.043, fill:C.Y1 });
  S(s, 'ellipse', { x:8.065, y:4.384, w:2.043, h:2.043, fill:C.Y1 });
  S(s, 'ellipse', { x:10.529, y:4.384, w:2.043, h:2.043, fill:C.Y1 });
  T(s, L2, { x:0.712, y:2.331, w:2.51, h:0.875, lineSpacingMultiple:1.5, margin:0 });
  T(s, 'INSTAGRAM', { x:0.6, y:2.044, w:1.185, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L2, { x:0.712, y:4.136, w:2.51, h:0.875, lineSpacingMultiple:1.5, margin:0 });
  T(s, 'FACEBOOK', { x:0.6, y:3.85, w:1.126, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L2, { x:4.273, y:2.331, w:2.51, h:0.875, lineSpacingMultiple:1.5, margin:0 });
  T(s, 'TWEETER', { x:4.16, y:2.044, w:0.994, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L2, { x:4.273, y:4.136, w:2.51, h:0.875, lineSpacingMultiple:1.5, margin:0 });
  T(s, 'YOUTUBE', { x:4.16, y:3.85, w:0.993, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, 'PLACEHOLDER', { x:0.6, y:5.62, w:6.369, h:0.976, lineSpacingMultiple:1.5 });
  lineIcon(s, 8.825, 4.97, 0.522, 0.522, C.INK);
  lineIcon(s, 11.29, 5.043, 0.521, 0.377, C.INK);
  lineIcon(s, 8.825, 2.56, 0.522, 0.522, C.INK);
  lineIcon(s, 11.289, 2.609, 0.522, 0.425, C.INK);
  T(s, '125k FOLLOWERS', { x:8.273, y:3.206, w:1.627, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, align:'center', wrap:false });
  T(s, '125k FOLLOWERS', { x:10.816, y:3.206, w:1.468, h:0.286, fontSize:11, bold:true, align:'center', wrap:false });
  T(s, '125k SUBCRIBER', { x:10.845, y:5.652, w:1.41, h:0.286, fontSize:11, bold:true, align:'center', wrap:false });
  T(s, '125k LIKE', { x:8.635, y:5.652, w:0.901, h:0.286, fontSize:11, bold:true, align:'center', wrap:false });
}

// ---------------------------------------------------------------------------
// Slide 29 — Infographic — timeline
// ---------------------------------------------------------------------------
function slide29(deck) {
  const s = newSlide(deck);
  gridLines(s, [1.867, 2.588, 3.307, 4.029, 4.75]);
  monthRail(s, 1.17, 4.7);
  timelineBars(s, [
      { x: 1.614, top: 3.68, h: 1.107, color: C.Y1, stemX: 1.702, stemTop: 3.405, labelX: 1.138, labelY: 3.054 },
      { x: 2.97, top: 4.076, h: 0.711, color: C.Y2, stemX: 3.059, stemTop: 3.814, labelX: 2.538, labelY: 3.475 },
      { x: 4.776, top: 3.569, h: 1.218, color: C.Y3, stemX: 4.862, stemTop: 3.237, labelX: 4.318, labelY: 2.92 },
      { x: 6.125, top: 3.68, h: 1.107, color: C.Y4, stemX: 6.211, stemTop: 3.569, labelX: 5.7, labelY: 3.221 },
      { x: 7.93, top: 3.814, h: 0.972, color: C.Y5, stemX: 8.011, stemTop: 3.405, labelX: 7.476, labelY: 3.092 },
      { x: 9.46, top: 4.226, h: 0.561, color: C.Y6, stemX: 9.542, stemTop: 3.814, labelX: 9.026, labelY: 3.47 },
      { x: 11.725, top: 3.569, h: 1.218, color: C.Y1, stemX: 11.797, stemTop: 3.405, labelX: 11.244, labelY: 3.06 },
  ]);
  legendRow(s, [2.305, 3.85, 5.213, 6.863, 8.613, 10.075], 5.59);
  T(s, 'INFOGRAPHIC', { x:5.095, y:0.501, w:3.144, h:0.572, fontSize:28, bold:true, fontFace:F.HEAD, align:'center', wrap:false });
  T(s, '31', { x:0.783, y:1.702, w:0.368, h:0.303, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, '21', { x:0.783, y:2.378, w:0.368, h:0.303, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, '7', { x:0.844, y:3.867, w:0.307, h:0.303, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, '14', { x:0.753, y:3.143, w:0.398, h:0.303, bold:true, fontFace:F.HEAD, align:'right', wrap:false });
  T(s, 'PLACEHOLDER', { x:0.733, y:6.113, w:11.867, h:0.673, align:'center', lineSpacingMultiple:1.5 });
}

// ---------------------------------------------------------------------------
// Slide 30 — Thanks for coming
// ---------------------------------------------------------------------------
function slide30(deck) {
  const s = newSlide(deck);
  T(s, ['THANKS', 'FOR COMING'].join(BR), { x:7.296, y:1.073, w:4.097, h:1.447, fontSize:40, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, L1, { x:7.296, y:2.596, w:5.185, h:1.279, lineSpacingMultiple:1.5 });
  S(s, 'rect', { x:6.972, y:4.979, w:5.759, h:2.045, fill:C.Y1 });
  T(s, 'Health Sports', { x:7.296, y:4.119, w:2.211, h:0.64, fontSize:32, fontFace:F.SCRIPT, color:C.Y1, wrap:false });
  lineIcon(s, 10.623, 5.284, 0.264, 0.264, C.INK);
  lineIcon(s, 10.623, 6.446, 0.264, 0.262, C.INK);
  T(s, 'Health_sports', { x:10.962, y:5.284, w:1.322, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, wrap:false });
  lineIcon(s, 10.613, 5.887, 0.283, 0.23, C.INK);
  T(s, 'Health_sports', { x:10.962, y:5.859, w:1.322, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, '022 654 565 123', { x:10.962, y:6.434, w:1.44, h:0.286, fontSize:11, bold:true, fontFace:F.HEAD, wrap:false });
  T(s, 'PLACEHOLDER', { x:7.296, y:5.284, w:3.069, h:1.459, fontSize:11, lineSpacingMultiple:1.5 });
  photo(s, 0.583, 0.726, 6.064, 6.299);
}

// ---------------------------------------------------------------------------
// Deck assembly
// ---------------------------------------------------------------------------
const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
];

function build() {
  const deck = new PptxGenJS();
  deck.defineLayout({ name: 'WIDE_13x7.5', width: 13.333, height: 7.5 });
  deck.layout = 'WIDE_13x7.5';
  deck.title = 'Health Sports';
  BUILDERS.forEach(function (fn) { fn(deck); });
  return deck.writeFile({ fileName: path.join(__dirname, '022ced68-2654-409d-a898-1ccb3f4eaf4b_grok_final.pptx') });
}

build().then(function (f) { console.log('wrote', f); });
