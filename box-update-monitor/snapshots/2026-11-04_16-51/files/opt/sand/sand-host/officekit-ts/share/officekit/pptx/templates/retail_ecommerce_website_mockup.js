/*
 * "SAM STYLE" - a 30-slide marketing template, rebuilt with pptxgenjs.
 *
 * Structure: a theme palette, helper functions for the furniture that repeats on
 * every page (hamburger mark, top nav, page number), a few pieces of drawn art
 * (device mock-ups and the two infographics) and one builder per slide.
 *
 * Photographs from the original are replaced by flat grey placeholder rectangles.
 *
 * Run:  node 020e628a-8b6e-4b6f-972d-f1dc03335d37_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// -------------------------------------------------------------- palette
const INK     = '2B2B2B';
const ACCENT  = 'FF5800';
const WHITE   = 'FFFFFF';
const BODY    = '959595';
const FAINT   = 'CACACA';
const MID     = '606060';
const SLATE   = '404040';
const RULE    = '808080';
const SNOW    = 'E6E6E6';
const PEACH   = 'FFDECC';
const PEACH2  = 'FFBC99';
const HAIR    = 'AAAAAA';
const PHOTO   = 'DDDDDD';
const CRIMSON = 'FF0000';

const HEAD = 'Poppins';   // theme major typeface
const TEXT = 'Open Sans'; // theme minor typeface

// ------------------------------------------------------ shared body copy
const COPY1 =
  'PLACEHOLDER';
const COPY2 =
  'PLACEHOLDER';
const COPY3 =
  'PLACEHOLDER';
const COPY4 =
  'PLACEHOLDER';
const COPY5 =
  'PLACEHOLDER';
const COPY6 =
  'PLACEHOLDER';
const COPY7 =
  'PLACEHOLDER';
const COPY8 =
  'PLACEHOLDER';
const COPY9 =
  'ban hanal anu banina atasi danusa ebetma husanes a';
const COPY10 =
  'PLACEHOLDER';
const COPY11 =
  'PLACEHOLDER';
const COPY12 =
  'PLACEHOLDER';
const COPY13 =
  'PLACEHOLDER';
const COPY14 =
  'PLACEHOLDER';
const COPY15 =
  'PLACEHOLDER';
const COPY16 =
  'PLACEHOLDER';
const COPY17 =
  'PLACEHOLDER';
const COPY18 =
  'PLACEHOLDER';
const COPY19 =
  'PLACEHOLDER';
const COPY20 =
  'PLACEHOLDER';
const COPY21 =
  'PLACEHOLDER';
const COPY22 =
  'ban hanal anuna bani atas anit at danusa  ebetan husan sanhas aduiban hanal';
const COPY23 =
  'PLACEHOLDER';
const COPY24 =
  'PLACEHOLDER';
const COPY25 =
  'PLACEHOLDER';
const COPY26 =
  'PLACEHOLDER';
const COPY27 =
  'PLACEHOLDER';
const COPY28 =
  'ban hanal anu banina atasanite han aduisa abaiai bana nabana';
const COPY29 =
  'PLACEHOLDER';

// ---------------------------------------------------------------- helpers

/** Three short grey rules: the "hamburger" mark in the page corner. */
function menu(s, x, y, scale = 1) {
  [0, 0.151, 0.3125].forEach(dy => {
    s.addShape('line', { x: x, y: y + dy * scale, w: 0.377 * scale, h: 0,
      line: { color: RULE, width: 2.25 * scale } });
  });
}

/** HOME / STORE / CONTACT US top navigation, anchored at its left edge. */
function nav(s, x, y) {
  [['HOME', 0, 0.686], ['STORE', 0.884, 0.717], ['CONTACT US', 1.769, 1.188]]
    .forEach(([label, dx, w]) => {
      s.addText(label, { x: x + dx, y: y, w: w, h: 0.353,
        align: 'center', valign: 'top', lineSpacingMultiple: 1.5,
        fontFace: HEAD, fontSize: 10, bold: true, color: FAINT });
    });
}

/** The "12." counter the slide master prints in the bottom-right corner. */
function pageNumber(s, n) {
  s.addText(n + '.', { x: 12.383, y: 6.933, w: 0.554, h: 0.337,
    align: 'center', valign: 'top', wrap: false,
    fontFace: HEAD, fontSize: 14, bold: true, color: BODY });
}

/** Stand-in for one of the deck's photographs. */
function photo(s, x, y, w, h) {
  s.addShape('rect', { x: x, y: y, w: w, h: h, fill: PHOTO });
}

/** iMac-style desktop monitor (slide 22): screen, chin, stand, foot. */
function imac(s, x, y, w, h) {
  s.addShape('roundRect', { x: x + 0.018 * w, y: y + 0.145 * h, w: 0.962 * w,
    h: 0.578 * h, fill: INK, rectRadius: 0.03 });
  s.addShape('rect', { x: x + 0.053 * w, y: y + 0.182 * h, w: 0.894 * w,
    h: 0.503 * h, fill: PHOTO });
  s.addShape('rect', { x: x + 0.018 * w, y: y + 0.723 * h, w: 0.962 * w,
    h: 0.087 * h, fill: 'C0C0C0' });
  s.addShape('trapezoid', { x: x + 0.335 * w, y: y + 0.810 * h, w: 0.330 * w,
    h: 0.125 * h, fill: 'B9B9B9' });
  s.addShape('ellipse', { x: x + 0.333 * w, y: y + 0.923 * h, w: 0.332 * w,
    h: 0.032 * h, fill: 'C6C6C6' });
}

/** Tablet, portrait (slide 23). */
function tablet(s, x, y, w, h) {
  s.addShape('roundRect', { x: x, y: y, w: w, h: h, fill: INK, rectRadius: 0.06 });
  s.addShape('rect', { x: x + 0.062 * w, y: y + 0.107 * h, w: 0.864 * w,
    h: 0.779 * h, fill: PHOTO });
  s.addShape('ellipse', { x: x + 0.462 * w, y: y + 0.920 * h, w: 0.076 * w,
    h: 0.050 * h, fill: '5A5A5A' });
}

/** Edge-to-edge phone (slide 24): thin steel rim, thin black bezel, notch. */
function phone(s, x, y, w, h) {
  s.addShape('roundRect', { x: x + 0.007 * w, y: y + 0.003 * h, w: 0.986 * w,
    h: 0.994 * h, fill: 'B0B0B0', rectRadius: 0.12 });
  s.addShape('roundRect', { x: x + 0.047 * w, y: y + 0.018 * h, w: 0.906 * w,
    h: 0.964 * h, fill: INK, rectRadius: 0.11 });
  s.addShape('roundRect', { x: x + 0.067 * w, y: y + 0.031 * h, w: 0.863 * w,
    h: 0.928 * h, fill: PHOTO, rectRadius: 0.1 });
  s.addShape('roundRect', { x: x + 0.26 * w, y: y + 0.031 * h, w: 0.48 * w,
    h: 0.026 * h, fill: INK, rectRadius: 0.5 });
}

/** Open laptop (slide 25): lid plus the light base slab it rests on. */
function laptop(s, x, y, w, h) {
  s.addShape('roundRect', { x: x + 0.092 * w, y: y + 0.005 * h, w: 0.816 * w,
    h: 0.925 * h, fill: INK, rectRadius: 0.03 });
  s.addShape('rect', { x: x + 0.119 * w, y: y + 0.056 * h, w: 0.762 * w,
    h: 0.820 * h, fill: PHOTO });
  s.addShape('roundRect', { x: x, y: y + 0.930 * h, w: w, h: 0.058 * h,
    fill: 'DEDEDE', rectRadius: 0.4 });
  s.addShape('roundRect', { x: x + 0.44 * w, y: y + 0.938 * h, w: 0.12 * w,
    h: 0.022 * h, fill: 'A8A8A8', rectRadius: 0.5 });
}

/** The small underlined kicker line that sits above nearly every headline. */
function kicker(s, o) {
  s.addText('Atasanit Hinas Atile  ', {
    x: o.x, y: o.y, w: o.w, h: o.h || 0.421, align: o.align || 'left', valign: 'top',
    lineSpacingMultiple: 1.5, fontFace: HEAD, fontSize: o.size || 14,
    color: FAINT, underline: { style: 'sng' },
  });
}

/** Two-tone headline: dark lead words, then an accent-coloured word. */
function heading(s, o) {
  const style = c => ({ fontSize: o.size, bold: true, color: c, fontFace: HEAD, align: o.align });
  const runs = [];
  if (o.lead) runs.push({ text: o.lead, options: Object.assign(style(o.leadColor || INK),
    o.stack ? { breakLine: true } : {}) });
  runs.push({ text: o.tail, options: style(o.tailColor || ACCENT) });
  s.addText(runs, { x: o.x, y: o.y, w: o.w, h: o.h, valign: 'top' });
}

/** Bold section label such as "First Service" / "Your Name". */
function subhead(s, o) {
  s.addText(o.text, { x: o.x, y: o.y, w: o.w, h: o.h, align: o.align || 'left', valign: 'top',
    lineSpacingMultiple: 1.5, fontFace: HEAD, fontSize: o.size, bold: true,
    color: o.color || MID });
}

/** A paragraph (or stack of paragraphs) of grey body copy. */
function bodyCopy(s, o) {
  const paras = [].concat(o.text);
  s.addText(paras.map((t, i) => ({ text: t, options: {
    fontSize: o.size || 11, color: o.color || BODY, fontFace: TEXT,
    align: o.align || 'left', lineSpacingMultiple: 1.5,
    breakLine: i < paras.length - 1,
  } })), { x: o.x, y: o.y, w: o.w, h: o.h, valign: 'top' });
}

/** Oversized tinted step number, e.g. "01." */
function stepNumber(s, o) {
  s.addText(o.text, { x: o.x, y: o.y, w: o.w, h: o.h, align: 'right', valign: 'top',
    lineSpacingMultiple: 1.5, fontFace: TEXT, fontSize: o.size, bold: true,
    color: o.color || PEACH });
}

/** The rotated "MODERN STYLE" wordmark running down the dark bars. */
function wordmark(s, o) {
  const style = (bold, color) => ({ fontSize: o.size, bold: bold, color: color,
    fontFace: HEAD, charSpacing: 3, align: o.align });
  const runs = [];
  if (o.dim) runs.push({ text: o.dim, options: style(true, FAINT) });
  runs.push({ text: o.bold, options: style(true, WHITE) });
  runs.push({ text: 'STYLE', options: style(false, WHITE) });
  s.addText(runs, { x: o.x, y: o.y, w: o.w, h: o.h, valign: 'top', rotate: o.rotate });
}

/** Filled circle with a centred white number, used by both infographics. */
function numberDisc(s, o) {
  s.addShape('ellipse', { x: o.x, y: o.y, w: o.w, h: o.h, fill: o.fill });
  s.addText(o.label, { x: o.x, y: o.y, w: o.w, h: o.h, align: 'center', valign: 'middle',
    fontFace: TEXT, fontSize: o.size, bold: true, color: WHITE });
}

/** Outlined "DETAILS" call-to-action button. */
function detailsButton(s, x, y) {
  s.addText('DETAILS', { x: x, y: y, w: 1.682, h: 0.435, align: 'center', valign: 'middle',
    shape: 'rect', line: { color: ACCENT, width: 1.5 },
    fontFace: TEXT, fontSize: 12, color: MID });
}

/** A "87%" skill bar: grey track, dark fill, right-aligned percentage. */
function skillBar(s, x, y, trackW, fillW, labelX, labelW, label) {
  const t = 0.139; // 10pt stroke, drawn as a rectangle so the square caps line up
  s.addShape('rect', { x: x - t / 2, y: y - t / 2, w: trackW + t, h: t, fill: FAINT });
  s.addShape('rect', { x: x - t / 2, y: y - t / 2, w: fillW + t, h: t, fill: SLATE });
  s.addText(label, { x: labelX, y: y - 0.21, w: labelW, h: 0.286,
    align: 'right', valign: 'top', fontFace: TEXT, fontSize: 11, bold: true, color: INK });
}

/** Slide 28: the open circle, its end caps, the four dots and dotted leaders. */
function infographicArc(s) {
  const cx = 3.518, cy = 3.976, rx = 2.389, ry = 2.389;
  s.addShape('custGeom', {
    x: 1.78, y: 1.587, w: 3.476, h: 4.779, line: { color: HAIR, width: 2 },
    points: [
      { x: 0, y: 4.526 },
      { x: 1.076, y: 4.779, curve: { type: 'cubic', x1: 0.323, y1: 4.688, x2: 0.69, y2: 4.779 } },
      { x: 3.476, y: 2.390, curve: { type: 'cubic', x1: 2.401, y1: 4.779, x2: 3.476, y2: 3.711 } },
      { x: 1.076, y: 0, curve: { type: 'cubic', x1: 3.476, y1: 1.068, x2: 2.401, y2: 0 } },
      { x: 0, y: 0.253, curve: { type: 'cubic', x1: 0.69, y1: 0, x2: 0.323, y2: 0.091 } },
    ],
  });
  [[0, 4.526], [0, 0.253]].forEach(([dx, dy]) => {
    s.addShape('ellipse', { x: 1.78 + dx - 0.055, y: 1.587 + dy - 0.055,
      w: 0.11, h: 0.11, fill: HAIR });
  });
  const DOTS = [
    { y: 2.0045, x: 4.2293, fill: INK,    line: [4.4864, 2.1471, 2.6044] },
    { y: 3.2254, x: 5.0265, fill: ACCENT, line: [5.3097, 3.3662, 1.7811] },
    { y: 4.4445, x: 5.0265, fill: PEACH2, line: [5.3097, 4.5854, 1.7811] },
    { y: 5.6760, x: 4.2293, fill: ACCENT, line: [4.4864, 5.8186, 2.6044] },
  ];
  DOTS.forEach(d => {
    s.addShape('line', { x: d.line[0], y: d.line[1], w: d.line[2], h: 0,
      line: { color: HAIR, width: 2, dashType: 'sysDot' } });
    s.addShape('ellipse', { x: d.x, y: d.y, w: 0.285, h: 0.285, fill: d.fill });
  });
}

/** Slide 29: the four curved arrows that link the numbered discs. */
function infographicCycle(s) {
  const ARROWS = [
    { x: 2.1187, y: 2.7158, w: 0.661, h: 0.659, from: [0, 0.659], to: [0.661, 0] },
    { x: 4.1892, y: 2.7113, w: 0.661, h: 0.659, from: [0, 0], to: [0.661, 0.659] },
    { x: 4.1915, y: 4.7751, w: 0.661, h: 0.657, from: [0.661, 0], to: [0, 0.657] },
    { x: 2.1187, y: 4.7751, w: 0.663, h: 0.661, from: [0.663, 0.661], to: [0, 0] },
  ];
  ARROWS.forEach(a => {
    const [fx, fy] = a.from, [tx, ty] = a.to;
    s.addShape('custGeom', {
      x: a.x, y: a.y, w: a.w, h: a.h,
      line: { color: FAINT, width: 4, endArrowType: 'triangle' },
      points: [
        { x: fx, y: fy },
        { x: tx, y: ty, curve: { type: 'cubic',
          x1: fx + (tx - fx) * 0.57, y1: fy + (ty - fy) * 0.22,
          x2: fx + (tx - fx) * 0.78, y2: fy + (ty - fy) * 0.43 } },
      ],
    });
  });
}

/** Slide 26: two-series stacked area chart. */
function areaChart(s, x, y, w, h) {
  const labels = ['5/1/2002', '6/1/2002', '7/1/2002', '8/1/2002', '9/1/2002'];
  // Passed as a one-entry multi-type list so the value axis crosses mid-category,
  // exactly as the original chart does.
  s.addChart([{ type: 'area', data: [
    { name: 'Series 1', labels: labels, values: [32, 32, 28, 12, 15] },
    { name: 'Series 2', labels: labels, values: [12, 12, 12, 21, 28] },
  ] }], {
    x: x, y: y, w: w, h: h,
    barGrouping: 'stacked', chartColors: [ACCENT, INK],
    showLegend: false, showTitle: false,
    catAxisLabelFontFace: TEXT, catAxisLabelFontSize: 10, catAxisLabelColor: MID,
    valAxisLabelFontFace: TEXT, valAxisLabelFontSize: 10, valAxisLabelColor: MID,
    valAxisMaxVal: 50, valAxisMajorUnit: 5,
    valGridLine: { color: SNOW, size: 0.75 },
    catAxisLineColor: SNOW, valAxisLineShow: false,
  });
}

/** Slide 27: three-series stacked column chart. */
function columnChart(s, x, y, w, h) {
  const labels = ['1', '2', '3', '4'];
  s.addChart('bar', [
    { name: 'Series 1', labels: labels, values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels: labels, values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels: labels, values: [2, 2, 3, 1.8] },
  ], {
    x: x, y: y, w: w, h: h,
    barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 150,
    chartColors: [INK, 'FF9B66', ACCENT],
    showLegend: false, showTitle: false,
    catAxisLabelFontFace: TEXT, catAxisLabelFontSize: 12, catAxisLabelColor: MID,
    valAxisLabelFontFace: TEXT, valAxisLabelFontSize: 12, valAxisLabelColor: MID,
    valAxisMaxVal: 10, valAxisMajorUnit: 1,
    valGridLine: { color: SNOW, size: 0.75 },
    catAxisLineColor: SNOW, valAxisLineShow: false,
  });
}

// --------------------------------------------------------- slide builders

function slide01(s) {
  pageNumber(s, 1);
  photo(s, 0, 0, 13.333, 7.5);
  menu(s, 0.469, 0.368);
  nav(s, 9.835, 0.463);
  heading(s, { x: 9.257, y: 5.087, w: 3.306, h: 1.717, lead: 'SAM', tail: 'STYLE', size: 48, tailColor: CRIMSON, align: 'right', stack: true });
}

function slide02(s) {
  pageNumber(s, 2);
  photo(s, 2.237, 0, 5.367, 7.5);
  s.addShape('rect', { x: 0, y: 0, w: 2.237, h: 7.5, fill: INK });
  heading(s, { x: 8.965, y: 1.837, w: 2.93, h: 1.178, lead: 'ABOUT OUR ', tail: 'PRODUCT', size: 32 });
  kicker(s, { x: 8.965, y: 2.973, w: 2.93 });
  bodyCopy(s, { x: 8.965, y: 4.039, w: 2.93, h: 1.489, text: COPY1 });
  detailsButton(s, 9.07, 6.005);
  wordmark(s, { x: -1.494, y: 3.245, w: 7.5, h: 1.01, bold: 'MODERN ', size: 54, align: 'center', rotate: 270 });
  menu(s, 0.469, 0.368);
  nav(s, 9.835, 0.463);
}

function slide03(s) {
  pageNumber(s, 3);
  s.addShape('rect', { x: 6.485, y: 0.77, w: 1.788, h: 3.98, fill: INK });
  photo(s, 7.632, 1.382, 4.375, 5.243);
  menu(s, 12.415, 0.362);
  heading(s, { x: 1.228, y: 5.986, w: 5.257, h: 0.639, lead: 'ABOUT OUR ', tail: 'PRODUCT', size: 32 });
  kicker(s, { x: 1.228, y: 5.429, w: 2.93 });
  wordmark(s, { x: 5.602, y: 2.636, w: 2.947, h: 0.438, bold: 'MODERN ', size: 20, rotate: 270 });
  nav(s, 0.542, 0.461);
  bodyCopy(s, { x: 1.228, y: 1.671, w: 4.418, h: 1.46, text: COPY2 });
  bodyCopy(s, { x: 1.228, y: 3.361, w: 4.418, h: 0.904, text: COPY3 });
}

function slide04(s) {
  pageNumber(s, 4);
  photo(s, 0, 0, 8.71, 6.103);
  s.addShape('rect', { x: 0, y: 6.103, w: 8.71, h: 1.397, fill: INK });
  wordmark(s, { x: 0.605, y: 5.598, w: 7.5, h: 1.01, bold: 'MODERN ', size: 54, align: 'center' });
  heading(s, { x: 9.861, y: 1.707, w: 2.437, h: 1.717, lead: 'ABOUT OUR ', tail: 'PRODUCT', size: 32 });
  kicker(s, { x: 9.861, y: 3.539, w: 2.437 });
  bodyCopy(s, { x: 9.861, y: 4.442, w: 2.437, h: 1.738, text: COPY4 });
  nav(s, 9.835, 0.463);
  menu(s, 0.469, 6.832);
}

function slide05(s) {
  pageNumber(s, 5);
  photo(s, 5.92, 0, 6.016, 6.603);
  s.addShape('rect', { x: 11.937, y: 0, w: 1.397, h: 6.603, fill: INK });
  nav(s, 0.542, 0.461);
  menu(s, 12.444, 0.362);
  heading(s, { x: 1.397, y: 1.744, w: 2.93, h: 1.178, lead: 'ABOUT OUR ', tail: 'PRODUCT', size: 32 });
  kicker(s, { x: 1.397, y: 2.881, w: 2.93 });
  bodyCopy(s, { x: 1.397, y: 3.77, w: 2.93, h: 1.46, text: COPY5 });
  wordmark(s, { x: 3.705, y: 5.869, w: 7.5, h: 1.01, bold: 'DERN ', dim: 'MO', size: 54, align: 'center' });
  detailsButton(s, 0.675, 6.126);
}

function slide06(s) {
  pageNumber(s, 6);
  photo(s, 5.296, 0, 3.799, 7.5);
  s.addShape('rect', { x: 4.657, y: 0.813, w: 0.639, h: 5.874, fill: INK });
  wordmark(s, { x: 3.515, y: 4.861, w: 2.962, h: 0.404, bold: 'MODERN ', size: 18, rotate: 270 });
  bodyCopy(s, { x: 0.93, y: 3.817, w: 2.93, h: 2.015, text: COPY6 });
  bodyCopy(s, { x: 0.935, y: 2.248, w: 2.93, h: 1.182, text: COPY7 });
  menu(s, 12.415, 0.362);
  nav(s, 0.542, 0.461);
  heading(s, { x: 10.008, y: 2.303, w: 2.437, h: 1.717, lead: 'ABOUT OUR ', tail: 'PRODUCT', size: 32 });
  kicker(s, { x: 10.008, y: 4.136, w: 2.437, h: 0.454 });
  detailsButton(s, 10.125, 5.133);
}

function slide07(s) {
  pageNumber(s, 7);
  photo(s, 0.979, 0, 4.604, 7.5);
  s.addShape('rect', { x: 0, y: 0, w: 0.979, h: 7.5, fill: INK });
  stepNumber(s, { x: 6.479, y: 0.422, w: 1.597, h: 1.589, text: '01.', size: 66 });
  bodyCopy(s, { x: 8.257, y: 1.678, w: 4.014, h: 0.904, text: COPY8 });
  subhead(s, { x: 8.257, y: 1.096, w: 4.014, h: 0.467, text: 'First Service', size: 16 });
  stepNumber(s, { x: 6.479, y: 2.766, w: 1.597, h: 1.61, text: '02.', size: 66 });
  bodyCopy(s, { x: 8.257, y: 4.022, w: 4.014, h: 0.904, text: COPY8 });
  subhead(s, { x: 8.257, y: 3.44, w: 4.014, h: 0.467, text: 'Second Service', size: 16 });
  heading(s, { x: 6.604, y: 6.258, w: 3.188, h: 0.639, lead: 'OUR ', tail: 'SERVICE', size: 32 });
  kicker(s, { x: 6.604, y: 5.702, w: 3.188 });
  menu(s, 12.415, 0.362);
  wordmark(s, { x: -0.971, y: 5.194, w: 2.962, h: 0.404, bold: 'MODERN ', size: 18, rotate: 270 });
}

function slide08(s) {
  pageNumber(s, 8);
  photo(s, 5.9, 0.758, 3.043, 5.984);
  stepNumber(s, { x: 0.611, y: 0.995, w: 1.281, h: 1.093, text: '01.', size: 44 });
  bodyCopy(s, { x: 2.009, y: 1.891, w: 2.168, h: 0.627, text: COPY9 });
  subhead(s, { x: 2.009, y: 1.309, w: 2.168, h: 0.421, text: 'First Service', size: 14 });
  stepNumber(s, { x: 0.611, y: 2.854, w: 1.281, h: 1.093, text: '02.', size: 44 });
  bodyCopy(s, { x: 2.009, y: 3.751, w: 2.168, h: 0.627, text: COPY9 });
  subhead(s, { x: 2.009, y: 3.169, w: 2.168, h: 0.421, text: 'Second Service', size: 14 });
  stepNumber(s, { x: 0.611, y: 4.72, w: 1.281, h: 1.093, text: '03.', size: 44 });
  bodyCopy(s, { x: 2.009, y: 5.616, w: 2.168, h: 0.627, text: COPY9 });
  subhead(s, { x: 2.009, y: 5.034, w: 2.168, h: 0.421, text: 'Third Service', size: 14 });
  s.addShape('rect', { x: 5.261, y: 0.758, w: 0.639, h: 5.984, fill: INK });
  wordmark(s, { x: 4.119, y: 4.933, w: 2.962, h: 0.37, bold: 'MODERN ', size: 16, rotate: 270 });
  menu(s, 12.415, 0.362);
  heading(s, { x: 9.991, y: 1.576, w: 2.437, h: 1.717, lead: 'ABOUT OUR ', tail: 'SERVICE', size: 32 });
  kicker(s, { x: 9.991, y: 3.409, w: 2.437 });
  bodyCopy(s, { x: 9.991, y: 4.311, w: 2.437, h: 1.738, text: COPY4 });
}

function slide09(s) {
  pageNumber(s, 9);
  s.addShape('rect', { x: 0, y: 4.167, w: 6.667, h: 3.333, fill: INK });
  photo(s, 5.583, 0.758, 7.75, 3.988);
  stepNumber(s, { x: 7.377, y: 4.869, w: 1.597, h: 1.589, text: '02.', size: 66 });
  bodyCopy(s, { x: 9.155, y: 6.125, w: 2.656, h: 0.656, text: COPY10 });
  subhead(s, { x: 9.155, y: 5.543, w: 2.656, h: 0.467, text: 'Second Service', size: 16 });
  stepNumber(s, { x: 1.26, y: 4.869, w: 1.597, h: 1.589, text: '01.', size: 66, color: RULE });
  bodyCopy(s, { x: 3.037, y: 6.125, w: 2.656, h: 0.656, text: COPY10, color: SNOW });
  subhead(s, { x: 3.037, y: 5.543, w: 2.656, h: 0.467, text: 'First Service', size: 16, color: WHITE });
  nav(s, 0.542, 0.461);
  bodyCopy(s, { x: 0.99, y: 1.564, w: 3.576, h: 1.738, text: COPY11 });
  menu(s, 0.469, 6.832);
}

function slide10(s) {
  pageNumber(s, 10);
  photo(s, 0, 0, 13.333, 7.5);
  stepNumber(s, { x: 1.635, y: 0.547, w: 1.597, h: 1.589, text: '01.', size: 66 });
  bodyCopy(s, { x: 3.413, y: 1.803, w: 2.972, h: 0.904, text: COPY12 });
  subhead(s, { x: 3.413, y: 1.221, w: 2.972, h: 0.467, text: 'First Service', size: 16 });
  stepNumber(s, { x: 1.635, y: 2.891, w: 1.597, h: 1.61, text: '03.', size: 66 });
  bodyCopy(s, { x: 3.413, y: 4.147, w: 2.972, h: 0.904, text: COPY12 });
  subhead(s, { x: 3.413, y: 3.565, w: 2.972, h: 0.467, text: 'Third Service', size: 16 });
  stepNumber(s, { x: 7.255, y: 0.547, w: 1.597, h: 1.589, text: '02.', size: 66 });
  bodyCopy(s, { x: 9.033, y: 1.803, w: 2.972, h: 0.904, text: COPY12 });
  subhead(s, { x: 9.033, y: 1.221, w: 2.972, h: 0.505, text: 'Second Service', size: 16 });
  stepNumber(s, { x: 7.255, y: 2.891, w: 1.597, h: 1.61, text: '04.', size: 66 });
  bodyCopy(s, { x: 9.033, y: 4.147, w: 2.972, h: 0.904, text: COPY12 });
  subhead(s, { x: 9.033, y: 3.565, w: 2.972, h: 0.467, text: 'Fourth Service', size: 16 });
  menu(s, 12.415, 0.362);
  heading(s, { x: 1.764, y: 6.291, w: 5.16, h: 0.639, lead: 'ABOUT OUR ', tail: 'SERVICE', size: 32 });
  kicker(s, { x: 1.764, y: 5.734, w: 3.188 });
  s.addShape('rect', { x: 0, y: 0, w: 0.639, h: 7.5, fill: INK });
  wordmark(s, { x: -1.156, y: 1.578, w: 2.962, h: 0.337, bold: 'MODERN ', size: 14, align: 'right', rotate: 270 });
}

function slide11(s) {
  pageNumber(s, 11);
  photo(s, 0, 0, 8.677, 4.812);
  s.addShape('rect', { x: 0, y: 4.812, w: 8.677, h: 2.688, fill: INK });
  stepNumber(s, { x: 0.518, y: 5.087, w: 1.281, h: 1.184, text: '01.', size: 48, color: RULE });
  bodyCopy(s, { x: 1.915, y: 6.027, w: 2.168, h: 0.904, text: COPY13, color: SNOW });
  subhead(s, { x: 1.915, y: 5.445, w: 2.168, h: 0.467, text: 'First Service', size: 16, color: WHITE });
  stepNumber(s, { x: 4.38, y: 5.087, w: 1.281, h: 1.184, text: '02.', size: 48, color: RULE });
  bodyCopy(s, { x: 5.778, y: 6.027, w: 2.129, h: 0.904, text: COPY13, color: SNOW });
  subhead(s, { x: 5.778, y: 5.445, w: 2.129, h: 0.467, text: 'Second Service', size: 16, color: WHITE });
  heading(s, { x: 9.861, y: 1.765, w: 2.437, h: 1.717, lead: 'ABOUT OUR ', tail: 'SERVICE', size: 32 });
  kicker(s, { x: 9.861, y: 3.598, w: 2.437 });
  bodyCopy(s, { x: 9.861, y: 4.5, w: 2.437, h: 1.738, text: COPY4 });
  nav(s, 9.835, 0.463);
  menu(s, 0.469, 0.368);
}

function slide12(s) {
  pageNumber(s, 12);
  photo(s, 2.237, 0.758, 2.852, 2.992);
  photo(s, 5.089, 0.758, 2.852, 2.992);
  photo(s, 5.089, 3.75, 2.852, 2.992);
  photo(s, 2.237, 3.75, 2.852, 2.992);
  s.addShape('rect', { x: 0, y: 0, w: 2.237, h: 7.5, fill: INK });
  menu(s, 0.469, 0.368);
  wordmark(s, { x: -0.958, y: 3.296, w: 5.983, h: 0.909, bold: 'MODERN ', size: 48, align: 'center', rotate: 270 });
  heading(s, { x: 9.223, y: 1.805, w: 2.93, h: 1.178, lead: 'ABOUT OUR ', tail: 'PORTFOLIO', size: 32 });
  kicker(s, { x: 9.223, y: 2.941, w: 2.93 });
  bodyCopy(s, { x: 9.223, y: 4.007, w: 2.93, h: 1.489, text: COPY1 });
  detailsButton(s, 9.328, 5.973);
  nav(s, 9.835, 0.463);
}

function slide13(s) {
  pageNumber(s, 13);
  photo(s, 0, 0, 4.742, 4.774);
  photo(s, 5.832, 4.774, 5.062, 2.742);
  s.addShape('rect', { x: 11.985, y: 0, w: 1.349, h: 6.603, fill: INK });
  wordmark(s, { x: 11.21, y: 4.477, w: 2.962, h: 0.404, bold: 'MODERN ', size: 18, rotate: 270 });
  heading(s, { x: 0.973, y: 6.149, w: 3.768, h: 0.639, lead: 'OUR ', tail: 'PORTFOLIO', size: 32 });
  kicker(s, { x: 0.973, y: 5.615, w: 2.852 });
  s.addText([
    { text: COPY14, options: { fontSize: 11, color: BODY, fontFace: TEXT, align: 'right', lineSpacingMultiple: 1.5, breakLine: true } },
    { text: '', options: { fontSize: 18, color: INK, fontFace: TEXT, align: 'right', lineSpacingMultiple: 1.5, breakLine: true } },
    { text: COPY15, options: { fontSize: 11, color: BODY, fontFace: TEXT, align: 'right', lineSpacingMultiple: 1.5 } },
  ], { x: 5.832, y: 1.414, w: 5.062, h: 2.6, valign: 'top' });
  nav(s, 8.033, 0.461);
  menu(s, 12.471, 0.362);
}

function slide14(s) {
  pageNumber(s, 14);
  photo(s, 8.365, 1.145, 4.125, 5.397);
  photo(s, 2.266, 1.145, 2.082, 2.397);
  photo(s, 4.963, 4.145, 2.082, 2.397);
  s.addShape('rect', { x: 0, y: 0, w: 0.979, h: 7.5, fill: INK });
  wordmark(s, { x: -0.971, y: 4.965, w: 2.962, h: 0.404, bold: 'MODERN ', size: 18, rotate: 270 });
  bodyCopy(s, { x: 2.266, y: 4.867, w: 2.082, h: 1.46, text: [COPY16, 'ban hanal anun'], align: 'right' });
  subhead(s, { x: 2.266, y: 4.36, w: 2.082, h: 0.398, text: 'First Portfolio', size: 13, align: 'right' });
  bodyCopy(s, { x: 4.963, y: 1.867, w: 2.082, h: 1.46, text: [COPY16, 'ban hanal anun'] });
  subhead(s, { x: 4.963, y: 1.36, w: 2.082, h: 0.398, text: 'Second Portfolio', size: 13 });
  nav(s, 9.835, 0.338);
  menu(s, 0.301, 0.368);
  heading(s, { x: 9.314, y: 5.716, w: 2.884, h: 0.639, tail: 'PORTFOLIO', size: 32, tailColor: WHITE, align: 'right' });
  kicker(s, { x: 9.314, y: 5.259, w: 2.884, align: 'right' });
}

function slide15(s) {
  pageNumber(s, 15);
  s.addShape('rect', { x: 2.328, y: 1.771, w: 8.677, h: 2.417, fill: INK });
  s.addShape('rect', { x: 12.276, y: 6.734, w: 1.057, h: 0.766, fill: WHITE });
  photo(s, 0, 1.104, 3.355, 3.75);
  photo(s, 4.989, 1.104, 3.355, 3.75);
  photo(s, 9.979, 1.104, 3.355, 3.75);
  wordmark(s, { x: 2.917, y: 2.474, w: 7.5, h: 1.01, bold: 'MODERN ', size: 54, align: 'center' });
  menu(s, 0.469, 0.368);
  nav(s, 9.856, 0.338);
  bodyCopy(s, { x: 0.846, y: 6.037, w: 2.509, h: 0.904, text: COPY17 });
  subhead(s, { x: 0.846, y: 5.476, w: 2.509, h: 0.454, text: 'First Portfolio', size: 14 });
  bodyCopy(s, { x: 9.979, y: 6.037, w: 2.509, h: 0.904, text: COPY17, align: 'right' });
  subhead(s, { x: 9.979, y: 5.476, w: 2.509, h: 0.454, text: 'Third Portfolio', size: 14, align: 'right' });
  bodyCopy(s, { x: 4.989, y: 6.037, w: 3.355, h: 0.904, text: COPY18, align: 'center' });
  subhead(s, { x: 4.989, y: 5.476, w: 3.355, h: 0.454, text: 'Second Portfolio', size: 14, align: 'center' });
}

function slide16(s) {
  pageNumber(s, 16);
  photo(s, 0, 0, 4.194, 7.5);
  photo(s, 5.948, 1.289, 2.993, 5.208);
  photo(s, 9.227, 1.289, 2.993, 5.208);
  s.addShape('rect', { x: 4.194, y: 1.289, w: 0.639, h: 5.208, fill: INK });
  wordmark(s, { x: 3.052, y: 3.049, w: 2.962, h: 0.337, bold: 'MODERN ', size: 14, align: 'right', rotate: 270 });
  nav(s, 9.835, 0.296);
  menu(s, 0.469, 0.368);
  heading(s, { x: 0.417, y: 6.588, w: 2.884, h: 0.639, tail: 'PORTFOLIO', size: 32, tailColor: WHITE });
  kicker(s, { x: 0.417, y: 6.131, w: 2.884 });
}

function slide17(s) {
  pageNumber(s, 17);
  photo(s, 0.979, 0, 6.694, 7.5);
  s.addShape('rect', { x: 0, y: 0, w: 0.979, h: 7.5, fill: INK });
  wordmark(s, { x: -0.971, y: 5.402, w: 2.962, h: 0.404, bold: 'MODERN ', size: 18, rotate: 270 });
  menu(s, 0.301, 0.368);
  nav(s, 9.835, 0.463);
  heading(s, { x: 8.918, y: 1.754, w: 2.93, h: 1.178, lead: 'Add YOUR', tail: 'NAME', size: 32, stack: true });
  kicker(s, { x: 8.918, y: 2.89, w: 2.93 });
  bodyCopy(s, { x: 8.918, y: 3.73, w: 3.32, h: 0.904, text: COPY19 });
  skillBar(s, 9.094, 5.372, 2.401, 2.134, 11.718, 0.519, '87%');
  skillBar(s, 9.094, 6.119, 2.401, 2.134, 11.718, 0.519, '87%');
}

function slide18(s) {
  pageNumber(s, 18);
  photo(s, 0, 0, 4.338, 5.771);
  photo(s, 4.338, 0, 4.338, 5.771);
  s.addShape('rect', { x: 0, y: 5.771, w: 8.677, h: 1.729, fill: INK });
  heading(s, { x: 9.861, y: 1.707, w: 2.437, h: 1.717, lead: 'ABOUT OUR ', tail: 'TEAM', size: 32 });
  kicker(s, { x: 9.861, y: 3.539, w: 2.437 });
  bodyCopy(s, { x: 9.861, y: 4.442, w: 2.437, h: 1.738, text: COPY4 });
  nav(s, 9.835, 0.463);
  menu(s, 0.469, 0.368);
  bodyCopy(s, { x: 1.118, y: 6.663, w: 2.118, h: 0.349, text: 'abaiai ebetma sanha duisa', color: SNOW });
  subhead(s, { x: 1.118, y: 6.211, w: 2.118, h: 0.467, text: 'Your Name', size: 16, color: WHITE });
  bodyCopy(s, { x: 5.456, y: 6.663, w: 2.118, h: 0.349, text: 'abaiai ebetma sanha duisa', color: SNOW });
  subhead(s, { x: 5.456, y: 6.211, w: 2.118, h: 0.467, text: 'Your Name', size: 16, color: WHITE });
}

function slide19(s) {
  pageNumber(s, 19);
  photo(s, 4.599, 0, 4.035, 7.5);
  nav(s, 9.835, 0.463);
  menu(s, 0.469, 0.368);
  skillBar(s, 1.033, 3.946, 2.023, 1.798, 3.272, 0.536, '87%');
  skillBar(s, 1.033, 4.693, 2.023, 1.798, 3.272, 0.536, '87%');
  skillBar(s, 1.033, 5.439, 2.023, 1.798, 3.272, 0.536, '87%');
  skillBar(s, 1.033, 6.186, 2.023, 1.798, 3.272, 0.536, '87%');
  heading(s, { x: 9.567, y: 1.764, w: 2.93, h: 1.178, lead: 'ADD YOUR', tail: 'NAME', size: 32, stack: true });
  kicker(s, { x: 9.567, y: 2.9, w: 2.93 });
  bodyCopy(s, { x: 9.567, y: 3.799, w: 2.93, h: 1.489, text: COPY1 });
  detailsButton(s, 9.672, 5.89);
  bodyCopy(s, { x: 0.846, y: 1.986, w: 2.962, h: 1.212, text: COPY20 });
}

function slide20(s) {
  pageNumber(s, 20);
  s.addShape('rect', { x: -0.125, y: 0, w: 3.516, h: 7.5, fill: INK });
  photo(s, 2.077, 1.354, 2.604, 2.494);
  photo(s, 2.077, 4.027, 2.604, 2.494);
  photo(s, 9.748, 1.354, 2.604, 2.494);
  photo(s, 9.748, 4.027, 2.604, 2.494);
  heading(s, { x: 5.749, y: 2.096, w: 2.93, h: 1.178, lead: 'ABOUT OUR ', tail: 'TEAM', size: 32, align: 'center' });
  kicker(s, { x: 5.749, y: 1.544, w: 2.93, align: 'center' });
  bodyCopy(s, { x: 5.596, y: 3.965, w: 3.238, h: 1.489, text: COPY21, align: 'center' });
  detailsButton(s, 6.353, 5.931);
  wordmark(s, { x: -1.697, y: 3.46, w: 6.473, h: 0.909, bold: 'MODERN ', size: 48, align: 'center', rotate: 270 });
  menu(s, 0.469, 0.368);
  nav(s, 9.835, 0.463);
}

function slide21(s) {
  pageNumber(s, 21);
  photo(s, 1.348, 1.331, 2.133, 3.011);
  photo(s, 3.704, 1.331, 2.133, 3.011);
  photo(s, 6.061, 1.331, 2.133, 3.011);
  nav(s, 9.835, 0.463);
  heading(s, { x: 9.34, y: 1.837, w: 2.93, h: 1.178, lead: 'ABOUT OUR ', tail: 'TEAM', size: 32 });
  kicker(s, { x: 9.34, y: 2.973, w: 2.93 });
  bodyCopy(s, { x: 9.34, y: 4.039, w: 2.93, h: 1.489, text: COPY1 });
  detailsButton(s, 9.445, 6.005);
  s.addShape('rect', { x: 0, y: 0, w: 0.639, h: 7.5, fill: INK });
  wordmark(s, { x: -0.813, y: 2.304, w: 2.283, h: 0.337, bold: 'MODERN ', size: 14, align: 'right', rotate: 270 });
  bodyCopy(s, { x: 1.554, y: 5.366, w: 1.822, h: 1.182, text: COPY22 });
  subhead(s, { x: 1.554, y: 4.882, w: 1.822, h: 0.429, text: 'Your Name', size: 13 });
  bodyCopy(s, { x: 3.911, y: 5.366, w: 1.822, h: 1.182, text: COPY22 });
  subhead(s, { x: 3.911, y: 4.882, w: 1.822, h: 0.429, text: 'Your Name', size: 13 });
  bodyCopy(s, { x: 6.267, y: 5.366, w: 1.822, h: 1.182, text: COPY22 });
  subhead(s, { x: 6.267, y: 4.882, w: 1.822, h: 0.429, text: 'Your Name', size: 13 });
  menu(s, 0.172, 7.014, 0.73);
}

function slide22(s) {
  pageNumber(s, 22);
  s.addShape('rect', { x: 6.692, y: 0.813, w: 2.968, h: 5.937, fill: INK });
  photo(s, 7.562, 1.596, 4.217, 2.441);
  imac(s, 7.363, 0.802, 4.607, 4.607);
  nav(s, 0.542, 0.461);
  menu(s, 12.415, 0.362);
  heading(s, { x: 0.808, y: 6.111, w: 5.257, h: 0.639, lead: 'ABOUT OUR ', tail: 'MOCKUP', size: 32 });
  kicker(s, { x: 0.808, y: 5.554, w: 2.93 });
  bodyCopy(s, { x: 1.297, y: 1.64, w: 4.418, h: 1.767, text: COPY23 });
  bodyCopy(s, { x: 1.297, y: 3.581, w: 4.418, h: 0.934, text: COPY24 });
}

function slide23(s) {
  pageNumber(s, 23);
  photo(s, 1.392, 1.758, 2.904, 3.825);
  s.addShape('rect', { x: 4.26, y: 3.342, w: 1.375, h: 3.323, fill: INK });
  tablet(s, 1.229, 1.288, 3.254, 4.8);
  nav(s, 9.835, 0.463);
  menu(s, 0.469, 0.368);
  heading(s, { x: 6.416, y: 1.663, w: 5.257, h: 0.639, lead: 'ABOUT OUR ', tail: 'MOCKUP', size: 32 });
  kicker(s, { x: 6.416, y: 2.254, w: 5.257, h: 0.454 });
  bodyCopy(s, { x: 6.416, y: 3.726, w: 2.774, h: 2.6, text: COPY25 });
  bodyCopy(s, { x: 9.689, y: 3.726, w: 2.774, h: 2.6, text: COPY26 });
  wordmark(s, { x: 3.613, y: 4.784, w: 2.829, h: 0.438, bold: 'MODERN ', size: 20, align: 'center', rotate: 270 });
}

function slide24(s) {
  pageNumber(s, 24);
  photo(s, 9.367, 1.725, 2.279, 4.817);
  nav(s, 9.835, 0.463);
  heading(s, { x: 1.553, y: 1.974, w: 5.257, h: 0.639, lead: 'ABOUT OUR ', tail: 'MOCKUP', size: 32 });
  kicker(s, { x: 1.552, y: 1.417, w: 2.93 });
  bodyCopy(s, { x: 1.552, y: 4.028, w: 2.797, h: 0.627, text: COPY27 });
  subhead(s, { x: 1.553, y: 3.446, w: 2.38, h: 0.467, text: 'First Service', size: 16 });
  bodyCopy(s, { x: 4.983, y: 4.028, w: 2.797, h: 0.627, text: COPY27 });
  subhead(s, { x: 4.983, y: 3.446, w: 2.38, h: 0.467, text: 'Second Service', size: 16 });
  bodyCopy(s, { x: 1.552, y: 5.815, w: 2.797, h: 0.627, text: COPY27 });
  subhead(s, { x: 1.553, y: 5.234, w: 2.38, h: 0.467, text: 'Third Service', size: 16 });
  bodyCopy(s, { x: 4.983, y: 5.815, w: 2.797, h: 0.627, text: COPY27 });
  subhead(s, { x: 4.983, y: 5.234, w: 2.38, h: 0.467, text: 'Fourth Service', size: 16 });
  s.addShape('rect', { x: 0, y: 0, w: 0.639, h: 7.5, fill: INK });
  wordmark(s, { x: -1.156, y: 1.578, w: 2.962, h: 0.337, bold: 'MODERN ', size: 14, align: 'right', rotate: 270 });
  phone(s, 9.261, 1.625, 2.52, 5.083);
}

function slide25(s) {
  pageNumber(s, 25);
  s.addShape('rect', { x: 0.838, y: 0.813, w: 2.968, h: 5.937, fill: INK });
  photo(s, 2.2, 1.687, 4.333, 2.737);
  laptop(s, 1.592, 1.54, 5.537, 3.231);
  menu(s, 12.415, 0.362);
  heading(s, { x: 8.694, y: 1.67, w: 2.93, h: 1.178, lead: 'ABOUT OUR ', tail: 'MOCKUP', size: 32 });
  kicker(s, { x: 8.694, y: 2.807, w: 2.93 });
  bodyCopy(s, { x: 8.694, y: 3.872, w: 2.93, h: 1.489, text: COPY1 });
  detailsButton(s, 8.799, 5.838);
}

function slide26(s) {
  pageNumber(s, 26);
  areaChart(s, 0.73, 1.417, 7.27, 5.317);
  nav(s, 0.542, 0.461);
  menu(s, 12.415, 0.362);
  heading(s, { x: 9.017, y: 1.799, w: 2.93, h: 1.178, lead: 'ABOUT OUR ', tail: 'CHART', size: 32 });
  kicker(s, { x: 9.017, y: 2.936, w: 2.93 });
  bodyCopy(s, { x: 9.017, y: 3.937, w: 2.93, h: 1.489, text: COPY1 });
  detailsButton(s, 9.122, 5.903);
}

function slide27(s) {
  pageNumber(s, 27);
  columnChart(s, 7.646, 1.308, 4.58, 5.387);
  nav(s, 9.835, 0.463);
  menu(s, 0.469, 0.368);
  heading(s, { x: 1.107, y: 1.916, w: 5.257, h: 0.639, lead: 'ABOUT OUR ', tail: 'CHART', size: 32 });
  kicker(s, { x: 1.107, y: 1.359, w: 2.93 });
  bodyCopy(s, { x: 1.107, y: 3.905, w: 2.464, h: 0.627, text: COPY28 });
  subhead(s, { x: 1.107, y: 3.323, w: 2.097, h: 0.467, text: 'First Service', size: 16 });
  bodyCopy(s, { x: 4.13, y: 3.905, w: 2.464, h: 0.627, text: COPY28 });
  subhead(s, { x: 4.13, y: 3.323, w: 2.097, h: 0.467, text: 'Second Service', size: 16 });
  bodyCopy(s, { x: 1.107, y: 5.693, w: 2.464, h: 0.627, text: COPY28 });
  subhead(s, { x: 1.107, y: 5.111, w: 2.097, h: 0.467, text: 'Third Service', size: 16 });
  bodyCopy(s, { x: 4.13, y: 5.693, w: 2.464, h: 0.627, text: COPY28 });
  subhead(s, { x: 4.13, y: 5.111, w: 2.097, h: 0.467, text: 'Fourth Service', size: 16 });
}

function slide28(s) {
  pageNumber(s, 28);
  infographicArc(s);
  numberDisc(s, { x: 6.968, y: 1.655, w: 1.014, h: 0.984, label: '01', size: 20, fill: INK });
  numberDisc(s, { x: 6.968, y: 2.874, w: 1.014, h: 0.984, label: '02', size: 20, fill: ACCENT });
  numberDisc(s, { x: 6.968, y: 4.094, w: 1.014, h: 0.984, label: '03', size: 20, fill: PEACH2 });
  numberDisc(s, { x: 6.968, y: 5.327, w: 1.014, h: 0.984, label: '04', size: 20, fill: ACCENT });
  nav(s, 9.835, 0.463);
  menu(s, 0.469, 0.368);
  heading(s, { x: 0.983, y: 3.169, w: 3.188, h: 1.178, lead: 'ABOUT ', tail: 'INFO-GRAPHIC', size: 32 });
  kicker(s, { x: 0.983, y: 4.305, w: 2.93 });
  bodyCopy(s, { x: 8.862, y: 1.834, w: 3.346, h: 0.627, text: COPY29 });
  bodyCopy(s, { x: 8.862, y: 3.053, w: 3.346, h: 0.627, text: COPY29 });
  bodyCopy(s, { x: 8.862, y: 4.202, w: 3.346, h: 0.627, text: COPY29 });
  bodyCopy(s, { x: 8.862, y: 5.505, w: 3.346, h: 0.627, text: COPY29 });
}

function slide29(s) {
  pageNumber(s, 29);
  numberDisc(s, { x: 2.781, y: 1.759, w: 1.409, h: 1.396, label: '01', size: 24, fill: 'FF3C00' });
  numberDisc(s, { x: 1.162, y: 3.372, w: 1.403, h: 1.403, label: '04', size: 24, fill: INK });
  numberDisc(s, { x: 2.781, y: 4.992, w: 1.409, h: 1.396, label: '03', size: 24, fill: 'EA0038' });
  numberDisc(s, { x: 4.407, y: 3.371, w: 1.403, h: 1.405, label: '02', size: 24, fill: 'F70000' });
  infographicCycle(s);
  nav(s, 0.542, 0.461);
  menu(s, 12.415, 0.362);
  heading(s, { x: 6.928, y: 1.851, w: 5.257, h: 0.639, lead: 'ABOUT ', tail: 'INFOGRAPHIC', size: 32 });
  kicker(s, { x: 6.928, y: 1.294, w: 2.93 });
  bodyCopy(s, { x: 6.928, y: 3.841, w: 2.464, h: 0.627, text: COPY28 });
  subhead(s, { x: 6.928, y: 3.259, w: 2.097, h: 0.467, text: 'First Service', size: 16 });
  bodyCopy(s, { x: 9.951, y: 3.841, w: 2.464, h: 0.627, text: COPY28 });
  subhead(s, { x: 9.951, y: 3.259, w: 2.097, h: 0.467, text: 'Second Service', size: 16 });
  bodyCopy(s, { x: 6.928, y: 5.629, w: 2.464, h: 0.627, text: COPY28 });
  subhead(s, { x: 6.928, y: 5.047, w: 2.097, h: 0.467, text: 'Third Service', size: 16 });
  bodyCopy(s, { x: 9.951, y: 5.629, w: 2.464, h: 0.627, text: COPY28 });
  subhead(s, { x: 9.951, y: 5.047, w: 2.097, h: 0.467, text: 'Fourth Service', size: 16 });
}

function slide30(s) {
  pageNumber(s, 30);
  photo(s, 0, 0, 13.333, 7.5);
  menu(s, 0.469, 0.368);
  nav(s, 9.835, 0.463);
  heading(s, { x: 10.127, y: 6.083, w: 2.617, h: 0.909, tail: 'THANK', size: 48, tailColor: CRIMSON, align: 'right' });
}

// ----------------------------------------------------------------- driver

const SLIDES = [
  slide01, slide02, slide03, slide04, slide05, slide06,
  slide07, slide08, slide09, slide10, slide11, slide12,
  slide13, slide14, slide15, slide16, slide17, slide18,
  slide19, slide20, slide21, slide22, slide23, slide24,
  slide25, slide26, slide27, slide28, slide29, slide30,
];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE_16x9', width: 13.3333, height: 7.5 });
pptx.layout = 'WIDE_16x9';
pptx.theme = { headFontFace: HEAD, bodyFontFace: TEXT };

SLIDES.forEach(build => build(pptx.addSlide()));

pptx.writeFile({
  fileName: path.join(__dirname, '020e628a-8b6e-4b6f-972d-f1dc03335d37_grok_final.pptx'),
});
