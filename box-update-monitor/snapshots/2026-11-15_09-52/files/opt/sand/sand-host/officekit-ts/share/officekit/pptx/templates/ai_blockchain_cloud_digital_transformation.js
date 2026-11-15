/**
 * "Digital Transformation" deck - rebuilt with PptxGenJS.
 *
 * 20 slides at 13.333 x 7.5 in (16:9). Near-black canvas with a lime / pink /
 * purple accent palette; every slide carries the same navigation bar.
 *
 * Run:  node <thisfile>.js   ->  writes <thisfile>.pptx beside the script.
 */
'use strict';
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const C = {
  dark:   '101014', // canvas
  shade:  '08080A', // canvas at 50% luminance - the muted bar-chart columns
  white:  'FEFEFE',
  lime:   'DBFD04',
  pink:   'FB54FB',
  purple: '6B0088',
  orchid: 'E25EFE',
};

/* -------------------------------------------------------------------- fonts */
const F = {
  head:       'Manrope Bold',              // theme major font
  body:       'Manrope',                   // theme minor font
  display:    'Bricolage Grotesque SemiBold',
  displayReg: 'Bricolage Grotesque Regular',
};

/* --------------------------------------------------------- shape primitives */
function box(s, x, y, w, h, o) {
  s.addShape('rect', Object.assign({ x, y, w, h }, o));
}
// Rounded rectangle, r = corner radius in inches.
function card(s, x, y, w, h, r, o) {
  s.addShape('roundRect', Object.assign({ x, y, w, h, rectRadius: r }, o));
}
// Fully rounded (stadium) shape - buttons, tags and bar-chart columns.
function pill(s, x, y, w, h, o) {
  s.addShape('roundRect', Object.assign({ x, y, w, h, rectRadius: Math.min(w, h) / 2 }, o));
}
function circle(s, x, y, w, h, o) {
  s.addShape('ellipse', Object.assign({ x, y, w, h }, o));
}
/**
 * Text box. Top-anchored with PowerPoint's default insets, like the source deck.
 * o = { size, font, color, bold, align, ls (line-spacing multiple), wrap }
 */
function txt(s, text, x, y, w, h, o) {
  o = o || {};
  s.addText(text, {
    x: x, y: y, w: w, h: h,
    fontSize: o.size || 18,
    fontFace: o.font || F.body,
    color: o.color || C.white,
    bold: !!o.bold,
    align: o.align || 'left',
    valign: 'top',
    lineSpacingMultiple: o.ls || null,
    wrap: o.wrap !== false,
    margin: [7.2, 7.2, 3.6, 3.6], // pt: 0.1in sides, 0.05in top/bottom
  });
}
/**
 * Stand-in for a photograph in the original deck. The source uses empty picture
 * placeholders that render as bare canvas, so the block is filled with the
 * canvas colour - the call still documents where the artwork belongs.
 */
function imageBox(s, x, y, w, h) {
  card(s, x, y, w, h, Math.min(w, h) * 0.064, { fill: C.dark }); // [image]
}

/* ------------------------------------------------------------ motif helpers */
// Row of three small dots.
function dots(s, x, y, d, gap, color) {
  for (let i = 0; i < 3; i++) circle(s, x + i * gap, y, d, d, { fill: color });
}
// Those three dots centred inside a filled disc (the "..." card menu).
function dotsBadge(s, x, y, d, bg, dotColor) {
  circle(s, x, y, d, d, { fill: bg });
  const r = d * 0.086, gap = d * 0.177;
  dots(s, x + d / 2 - r / 2 - gap, y + d * 0.46, r, gap, dotColor);
}
// Flame glyph (the brand mark): a teardrop stood on its point, with an inner
// lick carved out by a background-coloured crescent.
function flame(s, x, y, w, h, color, bg) {
  s.addShape('teardrop', { x: x, y: y, w: w * 1.12, h: h * 0.95, rotate: 315, fill: color });
  s.addShape('moon', { x: x - w * 0.02, y: y + h * 0.4, w: w * 0.6, h: h * 0.5, rotate: 240, fill: bg });
}
// Fingerprint glyph: nested ridges (quarter arcs) closing around a short stem.
// Each entry lists the arc start angles drawn for that ridge, outermost first.
const FP_RIDGES = [[180, 270, 0], [180, 270, 0], [180, 270], [180, 270]];
function fingerprint(s, x, y, w, h, color) {
  const cx = x + w / 2, cy = y + h / 2, lw = Math.max(0.75, w * 3.6);
  FP_RIDGES.forEach(function (angles, i) {
    const rw = w * (1 - i * 0.235), rh = h * (1 - i * 0.235);
    angles.forEach(function (a) {
      s.addShape('arc', { x: cx - rw / 2, y: cy - rh / 2, w: rw, h: rh, rotate: a, line: { color: color, width: lw } });
    });
  });
  s.addShape('line', { x: cx, y: cy + h * 0.05, w: 0, h: h * 0.32, line: { color: color, width: lw } });
}
// Four-petal mark: rounded squares arranged on a diamond.
function flower(s, x, y, w, h, color) {
  const sw = w * 0.335, sh = h * 0.345, o = { fill: color };
  card(s, x + w * 0.331, y, sw, sh, sw * 0.25, o);
  card(s, x, y + h * 0.325, sw, sh, sw * 0.25, o);
  card(s, x + w * 0.666, y + h * 0.325, sw, sh, sw * 0.25, o);
  card(s, x + w * 0.331, y + h * 0.655, sw, sh, sw * 0.25, o);
}
// A glyph centred on a filled disc; g = 'fire' | 'fp' | 'flower'.
const GLYPHS = { fire: flame, fp: fingerprint, flower: flower };
function badge(s, x, y, d, discColor, g, glyphColor, at) {
  circle(s, x, y, d, d, { fill: discColor });
  GLYPHS[g](s, at[0], at[1], at[2], at[3], glyphColor, discColor);
}
function smiley(s, x, y, d, color, faceColor) {
  s.addShape('smileyFace', { x, y, w: d, h: d, fill: color, line: { color: faceColor, width: 1.25 } });
}
// Thin stroked arrow: a shaft the width of its box plus two 45-degree barbs.
// rot = 0 points right; angles increase clockwise (PowerPoint convention).
function arrow(s, x, y, w, h, rot, color) {
  const cx = x + w / 2, cy = y + h / 2, len = Math.max(w, h);
  const ln = { color: color, width: Math.max(1, len * 9) }; // shaft = 12.5% of the box
  const r = rot * Math.PI / 180, hx = Math.cos(r) * len / 2, hy = Math.sin(r) * len / 2;
  s.addShape('line', { x: cx - hx, y: cy - hy, w: 2 * hx, h: 2 * hy, line: ln });
  [135, -135].forEach(function (d) {
    const a = (rot + d) * Math.PI / 180;
    s.addShape('line', { x: cx + hx, y: cy + hy, w: Math.cos(a) * len * 0.62, h: Math.sin(a) * len * 0.62, line: ln });
  });
}

/* ------------------------------------------------------------------- charts */
// Every chart in the deck is the same two-slice 66/33 split; only colours vary.
const SPLIT = [{ name: 'Sales', labels: [['1st Qtr', '2nd Qtr']], values: [66, 33] }];
const CHART_BASE = {
  firstSliceAng: 0, showLegend: false, showTitle: false,
  showValue: false, showPercent: false, showLabel: false,
};
function pieChart(s, x, y, w, h, colors) {
  s.addChart('pie', SPLIT, Object.assign({ x, y, w, h, chartColors: colors,
    dataBorder: { pt: 0, color: colors[0] } }, CHART_BASE));
}
function donutChart(s, x, y, w, h, colors, hole) {
  s.addChart('doughnut', SPLIT, Object.assign({ x, y, w, h, chartColors: colors, holeSize: hole,
    dataBorder: { pt: 0, color: colors[0] } }, CHART_BASE));
}

/* ------------------------------------------------------------------- navbar */
const NAV_LINKS = [
  { t: 'Home',     x: 4.556, w: 0.561 },
  { t: 'Partners', x: 5.639, w: 0.733 },
  { t: 'Features', x: 6.832, w: 0.743 },
  { t: 'Reach Us', x: 7.983, w: 0.767 },
];
// Identical header strip on all 20 slides.
function navbar(s, bold) {
  circle(s, 0.503, 0.622, 0.255, 0.255, { fill: C.purple });
  flame(s, 0.589, 0.686, 0.089, 0.128, C.white, C.purple);
  txt(s, 'Your Company', 0.83, 0.599, 1.401, 0.302, { size: 12, font: F.head, bold: bold, wrap: false });
  NAV_LINKS.forEach(function (l) {
    txt(s, l.t, l.x, 0.623, l.w, 0.253, { size: 9, font: F.head, bold: bold, align: 'center', wrap: false });
  });
  pill(s, 10.932, 0.623, 0.764, 0.253, { fill: C.pink });
  txt(s, 'Login', 11.028, 0.615, 0.575, 0.269, { size: 10, font: F.head, bold: bold, color: C.dark, align: 'center', wrap: false });
  pill(s, 12.066, 0.623, 0.764, 0.253, { fill: C.lime });
  txt(s, 'Sign Up', 12.085, 0.615, 0.722, 0.269, { size: 10, font: F.head, bold: bold, color: C.dark, align: 'center', wrap: false });
}

function slide01(s) { // Hero - AI, Blockchain, and Cloud
  navbar(s);
  imageBox(s, 2.997, 1.741, 3.184, 5.12);
  card(s, 0.625, 4.547, 1.955, 2.316, 0.236, { fill: C.pink });
  pill(s, 0.812, 5.823, 0.233, 0.884, { fill: C.white });
  pill(s, 1.479, 5.823, 0.231, 0.884, { fill: C.white });
  pill(s, 1.823, 5.823, 0.231, 0.884, { fill: C.white });
  pill(s, 2.144, 5.823, 0.231, 0.884, { fill: C.white });
  txt(s, 'AI, Blockchain, \nand Cloud Unlocking \nNew Business Frontiers', 6.972, 2.944, 6.986, 1.918, { size: 36, font: F.head });
  txt(s, 'Digital Transformation : Navigating the Digital Landscape.', 6.972, 2.264, 2.924, 0.505, { size: 12, font: F.head });
  txt(s, '95%', 9.422, 6.069, 0.835, 0.405, { size: 18, font: F.head });
  txt(s, 'Customer Happy', 9.422, 6.423, 1.38, 0.327, { size: 10, ls: 1.5 });
  txt(s, '5.0', 11.766, 6.069, 0.698, 0.405, { size: 18, font: F.head });
  txt(s, 'Client Ratings', 11.766, 6.423, 1.38, 0.327, { size: 10, ls: 1.5 });
  pill(s, 6.972, 6.193, 1.488, 0.434, { fill: C.purple });
  txt(s, 'Learn More', 7.154, 6.214, 1.124, 0.351, { size: 11, font: F.head, align: 'center', ls: 1.5 });
  txt(s, 'Ultricies tristique nulla aliquet enim tortor. A iaculis at alat pellentesque adipiscing.   ', 6.972, 5.295, 5.524, 0.271, { size: 10 });
  arrow(s, 12.311, 1.852, 0.37, 0.37, 135, C.white);
  card(s, 0.625, 1.783, 1.955, 2.316, 0.236, { fill: C.lime });
  txt(s, '298.687', 0.766, 3.457, 1.696, 0.538, { size: 26, font: F.head, color: C.dark, align: 'center' });
  txt(s, 'Data Statistics.', 1, 2.679, 1.227, 0.234, { size: 8, color: C.dark, align: 'center' });
  fingerprint(s, 0.863, 1.943, 0.4, 0.544, C.dark);
  dotsBadge(s, 2.054, 1.915, 0.335, C.dark, C.white);
  pill(s, 0.806, 6.248, 0.233, 0.458, { fill: C.shade });
  pill(s, 1.146, 5.823, 0.233, 0.884, { fill: C.shade });
  pill(s, 1.484, 6.04, 0.231, 0.667, { fill: C.shade });
  pill(s, 1.825, 6.398, 0.231, 0.309, { fill: C.shade });
  pill(s, 2.149, 5.913, 0.233, 0.793, { fill: C.shade });
  txt(s, '$2.6B', 0.67, 4.964, 1.245, 0.538, { size: 26, font: F.head, color: C.dark });
  txt(s, 'Impact On Business.', 0.672, 5.432, 1.549, 0.236, { size: 8, color: C.dark });
  dotsBadge(s, 2.102, 5.007, 0.333, C.dark, C.white);
  badge(s, 5.637, 1.878, 0.422, C.white, 'fp', C.dark, [5.759, 1.968, 0.179, 0.244]);
  pill(s, 1.106, 3.064, 0.15, 0.267, { fill: C.shade });
  pill(s, 1.325, 2.941, 0.15, 0.514, { fill: C.shade });
  pill(s, 1.544, 3.003, 0.15, 0.389, { fill: C.shade });
  pill(s, 1.762, 3.108, 0.15, 0.181, { fill: C.shade });
  pill(s, 1.972, 2.967, 0.15, 0.462, { fill: C.shade });
}

function slide02(s) { // Introduction
  navbar(s);
  imageBox(s, 0.872, 5.2, 5.309, 1.491);
  imageBox(s, 9.078, 2.111, 3.354, 2.019);
  txt(s, 'Introduction', 0.872, 1.825, 3.7, 0.774, { size: 40, font: F.head });
  txt(s, 'Sed vulputate mi sit amet mauris commodo quis imperdiet. Tempor orci eu lobortis elementum nibh tellus molestie nunc non. Auctor elite sed vulputate.', 0.872, 2.964, 5.542, 0.58, { size: 10, ls: 1.5 });
  pill(s, 0.872, 4.182, 1.488, 0.391, { fill: C.pink });
  txt(s, 'Next Slide', 1.063, 4.182, 1.106, 0.351, { size: 11, font: F.head, color: C.dark, align: 'center', ls: 1.5 });
  pill(s, 4.681, 4.182, 1.488, 0.391, { fill: C.purple });
  txt(s, 'Learn More', 4.872, 4.231, 1.106, 0.285, { size: 11, font: F.head, align: 'center' });
  badge(s, 5.625, 5.304, 0.422, C.white, 'fp', C.dark, [5.747, 5.394, 0.179, 0.244]);
  arrow(s, 5.939, 1.878, 0.37, 0.37, 135, C.white);
  card(s, 10.953, 4.488, 1.479, 2.062, 0.172, { fill: C.lime });
  donutChart(s, 10.644, 4.719, 2.087, 1.391, [C.dark, C.white], 67);
  card(s, 7.208, 4.488, 3.356, 2.061, 0.178, { fill: C.pink });
  card(s, 7.208, 2.111, 1.479, 2.042, 0.161, { line: { color: C.white, width: 1 } });
  txt(s, 'Hellow, We Are Brisk.', 7.389, 5.3, 2.986, 1.179, { size: 32, font: F.display, color: C.dark });
  smiley(s, 7.498, 4.839, 0.309, C.dark, C.pink);
  dotsBadge(s, 9.811, 4.743, 0.498, C.dark, C.white);
  pieChart(s, 7.125, 2.203, 1.634, 1.089, [C.pink, C.lime]);
  txt(s, '12.4', 7.408, 3.181, 1.028, 0.571, { size: 28, font: F.display, align: 'center' });
  txt(s, 'Statistics Point.', 7.378, 3.675, 1.144, 0.252, { size: 9, font: F.displayReg, align: 'center' });
  txt(s, '90%', 11.325, 5.247, 0.726, 0.337, { size: 14, font: F.display, color: C.dark, align: 'center' });
  txt(s, 'Professional Company.', 11.219, 6.038, 0.944, 0.405, { size: 9, font: F.displayReg, color: C.dark, align: 'center' });
  dots(s, 12.053, 4.7, 0.043, 0.088, C.lime);
  badge(s, 11.884, 2.274, 0.394, C.white, 'fire', C.dark, [12.016, 2.363, 0.137, 0.2]);
}

function slide03(s) { // Overview of Digital Landscape
  navbar(s);
  imageBox(s, 0.578, 2.016, 4.733, 4.885);
  card(s, 6.427, 4.984, 2.97, 1.788, 0.205, { fill: C.pink });
  card(s, 9.778, 4.984, 2.97, 1.788, 0.19, { fill: C.lime });
  txt(s, 'Overview of Digital Landscape.', 6.231, 2.016, 6.273, 1.448, { size: 40, font: F.head });
  txt(s, 'Nunc vel risus commodo viverra. Fermentum odio eu feugiat pretium nibh ipsum consequat. ', 6.964, 3.856, 3.95, 0.58, { size: 10, ls: 1.5 });
  txt(s, 'Evolution of Digital Technologies.', 6.571, 5.13, 2.276, 0.573, { size: 14, font: F.head, color: C.dark });
  txt(s, 'Silent vel quam elementum pulvinar etiam non. Facilisis gravida neque convallis a cras semper.', 6.571, 5.722, 2.689, 0.832, { size: 10, color: C.dark, ls: 1.5 });
  txt(s, 'Current State of Digital Transformation.', 9.863, 5.13, 2.616, 0.573, { size: 14, font: F.head, color: C.dark });
  txt(s, 'Silent vel quam elementum pulvinar etiam non. Facilisis gravida neque convallis a cras semper.', 9.887, 5.722, 2.781, 0.832, { size: 10, color: C.dark, ls: 1.5 });
  pill(s, 9.8, 2.953, 1.488, 0.391, { fill: C.purple });
  txt(s, 'Learn More', 9.983, 2.934, 1.123, 0.351, { size: 11, font: F.head, align: 'center', ls: 1.5 });
  circle(s, 6.344, 3.967, 0.387, 0.387, { fill: C.purple });
  arrow(s, 6.464, 4.089, 0.146, 0.146, 0, C.white);
  badge(s, 4.712, 2.168, 0.422, C.white, 'fp', C.dark, [4.834, 2.258, 0.179, 0.244]);
}

function slide04(s) { // Artificial Intelligence
  navbar(s);
  imageBox(s, 0.83, 1.741, 2.535, 5.238);
  imageBox(s, 3.604, 1.741, 2.535, 5.238);
  card(s, 6.771, 5.686, 2.483, 1.217, 0.164, { fill: C.lime });
  card(s, 9.488, 4.238, 2.483, 1.219, 0.165, { fill: C.lime });
  card(s, 9.488, 5.686, 2.483, 1.217, 0.164, { fill: C.pink });
  card(s, 6.753, 4.238, 2.481, 1.219, 0.165, { fill: C.pink });
  txt(s, 'Artificial Intelligence.', 6.701, 1.764, 6.278, 0.774, { size: 40, font: F.head });
  txt(s, 'Ultricies tristique nulla aliquet enim tortor. A iaculis at alat pellentesque adipiscing. Sagittis orci a scelerisque purus semper.', 7.502, 2.962, 5.01, 0.578, { size: 10, ls: 1.5 });
  txt(s, 'Definition & Overview.', 6.837, 4.366, 1.906, 0.286, { size: 11, font: F.head, color: C.dark });
  txt(s, 'AI Application & Business.', 9.639, 4.366, 2.184, 0.286, { size: 11, font: F.head, color: C.dark });
  txt(s, 'Compatibility', 9.658, 5.792, 1.623, 0.351, { size: 11, font: F.head, color: C.dark, ls: 1.5 });
  txt(s, 'Nemo enim ipsam voluptatem quia voluptas sit aut odit', 6.837, 4.674, 2.352, 0.582, { size: 10, color: C.dark, ls: 1.5 });
  txt(s, 'Nemo enim ipsam voluptatem quia voluptas sit aut odit', 9.665, 4.674, 2.352, 0.582, { size: 10, color: C.dark, ls: 1.5 });
  txt(s, 'Impact on Efficiency.', 6.856, 5.852, 2.094, 0.285, { size: 11, font: F.head, color: C.dark });
  txt(s, 'Nemo enim ipsam voluptatem quia voluptas sit aut odit', 6.837, 6.168, 2.352, 0.582, { size: 10, color: C.dark, ls: 1.5 });
  txt(s, 'Nemo enim ipsam voluptatem quia voluptas sit aut odit', 9.665, 6.158, 2.352, 0.58, { size: 10, color: C.dark, ls: 1.5 });
  badge(s, 5.559, 1.896, 0.424, C.white, 'fp', C.dark, [5.682, 1.984, 0.179, 0.244]);
  badge(s, 2.788, 1.873, 0.422, C.white, 'fp', C.dark, [2.909, 1.963, 0.179, 0.244]);
  circle(s, 6.807, 3.068, 0.389, 0.387, { fill: C.purple });
  arrow(s, 6.929, 3.188, 0.146, 0.148, 0, C.white);
}

function slide05(s) { // AI Trends
  navbar(s);
  imageBox(s, 0.67, 1.741, 3.578, 3.809);
  pill(s, 10.366, 6.401, 1.486, 0.377, { fill: C.purple });
  txt(s, 'AI Trends.', 1.123, 6.128, 3.054, 0.774, { size: 40, font: F.head });
  circle(s, 4.889, 6.361, 0.387, 0.387, { fill: C.purple });
  arrow(s, 5.009, 6.483, 0.148, 0.146, 315, C.white);
  txt(s, 'Silent vel quam elementum pulvinar etiam non. Facilisis gravida neque convallis a cras semper.', 5.439, 6.227, 3.281, 0.58, { size: 10, ls: 1.5 });
  txt(s, 'Next Slide', 10.583, 6.401, 1.052, 0.349, { size: 11, font: F.head, align: 'center', ls: 1.5 });
  arrow(s, 11.995, 6.516, 0.146, 0.148, 0, C.dark);
  card(s, 8.562, 1.717, 4.151, 3.865, 0.259, { fill: C.lime });
  txt(s, 'NLP Adoption Overtime.', 8.889, 2.031, 2.969, 0.372, { size: 16, font: F.head, color: C.dark });
  txt(s, '10%', 8.91, 2.675, 0.656, 0.304, { size: 12, font: F.head, color: C.dark, align: 'center' });
  txt(s, '2090', 8.887, 5.076, 0.656, 0.269, { size: 10, font: F.head, color: C.dark, align: 'center' });
  txt(s, '25%', 9.613, 2.675, 0.656, 0.304, { size: 12, font: F.head, color: C.dark, align: 'center' });
  txt(s, '2091', 9.589, 5.076, 0.656, 0.269, { size: 10, font: F.head, color: C.dark, align: 'center' });
  txt(s, '34%', 10.316, 2.675, 0.656, 0.304, { size: 12, font: F.head, color: C.dark, align: 'center' });
  txt(s, '2092', 10.292, 5.076, 0.656, 0.269, { size: 10, font: F.head, color: C.dark, align: 'center' });
  txt(s, '43%', 11.017, 2.675, 0.656, 0.304, { size: 12, font: F.head, color: C.dark, align: 'center' });
  txt(s, '2093', 10.995, 5.076, 0.656, 0.269, { size: 10, font: F.head, color: C.dark, align: 'center' });
  txt(s, '56%', 11.738, 2.675, 0.656, 0.304, { size: 12, font: F.head, color: C.dark, align: 'center' });
  txt(s, '2094', 11.696, 5.076, 0.658, 0.269, { size: 10, font: F.head, color: C.dark, align: 'center' });
  card(s, 4.727, 3.877, 3.394, 1.674, 0.279, { line: { color: C.white, width: 0.31 } });
  card(s, 4.727, 1.747, 3.394, 1.674, 0.279, { fill: C.pink });
  txt(s, 'Predictive Analysis.', 4.807, 4.174, 1.819, 0.538, { size: 13, font: F.head });
  txt(s, 'nisi ut aliquid ex ea commodi cons equatur Quis autem vel eum.', 4.882, 4.786, 2.201, 0.533, { size: 9, ls: 1.5 });
  txt(s, 'Machine Learning & Deep Learning.', 4.797, 2.059, 2.299, 0.538, { size: 13, font: F.head, color: C.dark });
  txt(s, 'nisi ut aliquid ex ea commodi cons equatur Quis autem vel eum.', 4.872, 2.672, 2.201, 0.533, { size: 9, color: C.dark, ls: 1.5 });
  dotsBadge(s, 7.536, 1.842, 0.418, C.dark, C.white);
  dotsBadge(s, 7.594, 3.979, 0.418, C.dark, C.white);
  badge(s, 3.661, 1.861, 0.424, C.white, 'fp', C.dark, [3.784, 1.95, 0.179, 0.244]);
  dotsBadge(s, 12.146, 1.842, 0.418, C.dark, C.white);
  fingerprint(s, 7.545, 2.595, 0.427, 0.583, C.dark);
  fingerprint(s, 7.545, 4.684, 0.428, 0.582, C.white);
  const NLP_COLUMNS = [
    { x: [9.026, 9.262], c: [[C.dark, C.dark], [C.dark, C.dark], [C.purple, C.dark], [C.purple, C.white]] },
    { x: [9.715, 9.951], c: [[C.dark, C.dark], [C.purple, C.dark], [C.purple, C.white], [C.purple, C.white]] },
    { x: [10.444, 10.681], c: [[C.dark, C.white], [C.dark, C.white], [C.purple, C.white], [C.purple, C.white]] },
    { x: [11.134, 11.368], c: [[C.dark, C.dark], [C.dark, C.white], [C.purple, C.white], [C.purple, C.white]] },
    { x: [11.839, 12.075], c: [[C.purple, C.dark], [C.purple, C.white], [C.purple, C.white], [C.purple, C.white]] },
  ];
  const NLP_ROW_Y = [3.184, 3.648, 4.08, 4.545];
  NLP_COLUMNS.forEach(function (colm) {
    colm.c.forEach(function (pair, r) {
      pill(s, colm.x[0], NLP_ROW_Y[r], 0.16, 0.382, { fill: pair[0] });
      pill(s, colm.x[1], NLP_ROW_Y[r], 0.16, 0.382, { fill: pair[1] });
    });
  });
}

function slide06(s) { // AI Examples
  navbar(s);
  imageBox(s, 8.26, 3.161, 4.589, 3.856);
  card(s, 4.405, 5.344, 3.394, 1.674, 0.279, { line: { color: C.white, width: 0.31 } });
  pill(s, 4.585, 5.833, 0.233, 0.7, { fill: C.white });
  pill(s, 5.599, 5.833, 0.231, 0.7, { fill: C.white });
  pill(s, 5.927, 5.833, 0.231, 0.7, { fill: C.white });
  pill(s, 4.924, 5.833, 0.233, 0.7, { fill: C.white });
  pill(s, 5.26, 5.833, 0.231, 0.7, { fill: C.white });
  card(s, 0.503, 3.158, 3.394, 1.674, 0.279, { line: { color: C.white, width: 0.31 } });
  dotsBadge(s, 3.372, 3.259, 0.418, C.dark, C.white);
  txt(s, 'Chatbots & Virtual Assistants.', 0.698, 3.398, 2.257, 0.538, { size: 13, font: F.head });
  txt(s, 'id est laborum et dolorum fuga. Et harum quidem rerum facilis.', 0.698, 4.069, 2.109, 0.531, { size: 9, ls: 1.5 });
  card(s, 0.503, 5.344, 3.394, 1.674, 0.279, { fill: C.lime });
  dotsBadge(s, 3.372, 5.444, 0.418, C.dark, C.white);
  txt(s, 'Soppy Chain Optimization.', 0.698, 5.583, 1.764, 0.538, { size: 13, font: F.head, color: C.dark });
  txt(s, 'id est laborum et dolorum fuga. Et harum quidem rerum facilis.', 0.698, 6.253, 2.109, 0.533, { size: 9, color: C.dark, ls: 1.5 });
  card(s, 4.405, 3.149, 3.394, 1.674, 0.279, { fill: C.lime });
  dotsBadge(s, 7.274, 3.25, 0.418, C.dark, C.white);
  txt(s, 'Personalized Marketing.', 4.599, 3.389, 1.766, 0.538, { size: 13, font: F.head, color: C.dark });
  txt(s, 'id est laborum et dolorum fuga. Et harum quidem rerum facilis.', 4.599, 4.059, 2.109, 0.533, { size: 9, color: C.dark, ls: 1.5 });
  pieChart(s, 6.845, 4.059, 1.127, 0.75, [C.pink, C.white]);
  pill(s, 4.585, 6.17, 0.233, 0.363, { fill: C.pink });
  pill(s, 5.26, 6.005, 0.231, 0.528, { fill: C.pink });
  pill(s, 5.599, 6.29, 0.231, 0.243, { fill: C.pink });
  pill(s, 5.927, 5.908, 0.231, 0.628, { fill: C.pink });
  txt(s, '198.789', 6.345, 5.845, 1.55, 0.505, { size: 24, font: F.head });
  txt(s, 'Data Statistics.', 6.347, 6.299, 1.55, 0.234, { size: 8 });
  txt(s, 'AI Examples.', 3.691, 1.632, 5.951, 1.111, { size: 60, font: F.head, align: 'center' });
  badge(s, 12.333, 3.299, 0.424, C.white, 'fp', C.dark, [12.456, 3.388, 0.179, 0.244]);
  badge(s, 7.253, 4.28, 0.311, C.dark, 'flower', C.pink, [7.329, 4.365, 0.158, 0.154]);
  pieChart(s, 2.932, 4.059, 1.127, 0.75, [C.pink, C.white]);
  badge(s, 3.34, 4.28, 0.311, C.dark, 'flower', C.pink, [3.416, 4.365, 0.158, 0.154]);
  pieChart(s, 2.932, 6.207, 1.127, 0.752, [C.pink, C.white]);
  badge(s, 3.34, 6.427, 0.311, C.dark, 'flower', C.pink, [3.416, 6.512, 0.158, 0.154]);
}

function slide07(s) { // Blockchain Technology
  navbar(s);
  imageBox(s, 4.908, 1.658, 7.38, 1.988);
  txt(s, 'Blockchain Technology.', 0.887, 1.929, 3.7, 1.446, { size: 40, font: F.head });
  card(s, 4.908, 4.179, 3.392, 1.674, 0.279, { fill: C.pink });
  txt(s, 'Decentralization & Security.', 5.068, 4.373, 1.691, 0.538, { size: 13, font: F.head, color: C.dark });
  txt(s, 'Neque porro quisquam est, qui dol\norem ipsum quia dolor sit amet.', 5.068, 5.042, 2.3, 0.531, { size: 9, color: C.dark, ls: 1.5 });
  dotsBadge(s, 7.797, 4.234, 0.418, C.dark, C.white);
  card(s, 0.922, 4.165, 3.394, 1.675, 0.279, { fill: C.lime });
  txt(s, 'Use Cases Across Industry.', 1.082, 4.359, 1.693, 0.538, { size: 13, font: F.head, color: C.dark });
  txt(s, 'Neque porro quisquam est, qui dol\norem ipsum quia dolor sit amet.', 1.082, 5.028, 2.302, 0.533, { size: 9, color: C.dark, ls: 1.5 });
  dotsBadge(s, 3.812, 4.22, 0.418, C.dark, C.white);
  card(s, 8.892, 4.165, 3.394, 1.674, 0.279, { line: { color: C.white, width: 0.31 } });
  dotsBadge(s, 11.783, 4.22, 0.417, C.dark, C.white);
  txt(s, 'Definition & Basic.', 9.087, 4.351, 1.332, 0.538, { size: 13, font: F.head });
  txt(s, 'id est laborum et dolorum fuga. Et harum quidem rerum facilis.', 9.087, 5.076, 2.109, 0.533, { size: 9, ls: 1.5 });
  pieChart(s, 0.589, 6.276, 1.325, 0.884, [C.purple, C.white]);
  txt(s, '12.4', 1.819, 6.304, 1.028, 0.573, { size: 28, font: F.head });
  txt(s, 'Statistics Point.', 1.819, 6.825, 1.146, 0.252, { size: 9 });
  badge(s, 11.726, 1.79, 0.422, C.white, 'fp', C.dark, [11.847, 1.879, 0.179, 0.244]);
  pill(s, 4.908, 6.576, 1.236, 0.325, { fill: C.lime });
  txt(s, 'Next Slide', 4.99, 6.549, 1.076, 0.349, { size: 11, font: F.head, color: C.dark, align: 'center', ls: 1.5 });
  pill(s, 11.05, 6.576, 1.236, 0.325, { fill: C.pink });
  txt(s, 'Learn More', 11.16, 6.587, 1.04, 0.286, { size: 11, font: F.head, color: C.dark, align: 'center' });
  fingerprint(s, 7.731, 5.026, 0.429, 0.582, C.dark);
  fingerprint(s, 11.703, 5.026, 0.429, 0.582, C.dark);
  fingerprint(s, 3.753, 5.026, 0.427, 0.582, C.dark);
}

function slide08(s) { // Blockchain Application
  navbar(s);
  imageBox(s, 8.092, 4.344, 4.139, 2.378);
  txt(s, 'Blockchain Application.', 8.12, 1.745, 4.434, 1.179, { size: 32, font: F.head });
  card(s, 1.085, 1.764, 3.089, 1.382, 0.187, { line: { color: C.white, width: 1 } });
  txt(s, 'Smart Contact.', 1.349, 2.016, 1.611, 0.302, { size: 12, font: F.head });
  txt(s, 'Sit amet risus nullam eget felis eget. Sit amet commodo quis imperdiet.', 1.349, 2.295, 2.608, 0.58, { size: 10, ls: 1.5 });
  card(s, 4.477, 3.562, 3.089, 1.382, 0.187, { fill: C.pink });
  txt(s, 'Cryptocurrencies', 4.75, 3.814, 1.87, 0.304, { size: 12, font: F.head, color: C.dark });
  txt(s, 'Sit amet risus nullam eget felis eget. Sit amet commodo quis imperdiet.', 4.75, 4.095, 2.608, 0.58, { size: 10, color: C.dark, ls: 1.5 });
  card(s, 1.085, 5.34, 3.089, 1.382, 0.187, { fill: C.lime });
  txt(s, 'Financial Services.', 1.349, 5.597, 2.132, 0.304, { size: 12, font: F.head, color: C.dark });
  txt(s, 'Sit amet risus nullam eget felis eget. Sit amet commodo quis imperdiet.', 1.349, 5.878, 2.608, 0.58, { size: 10, color: C.dark, ls: 1.5 });
  txt(s, 'Nunc vel risus commodo viverra. Fermentum odio eu feugiat pretium nibh ipsum consequat. Vulputate dignissim.', 8.092, 3.109, 4.056, 0.58, { size: 10, ls: 1.5 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.', 1.533, 4.191, 1.747, 0.484, { size: 8, ls: 1.5 });
  txt(s, 'The Details.', 1.512, 3.816, 1.186, 0.375, { size: 12, font: F.head, ls: 1.5 });
  badge(s, 1.075, 4.13, 0.363, C.white, 'flower', C.purple, [1.163, 4.228, 0.185, 0.18]);
  txt(s, 'AI users are increasing.', 6.109, 2.283, 1.054, 0.372, { size: 8 });
  txt(s, '16.9M+', 4.983, 2.262, 1.156, 0.438, { size: 20, font: F.head });
  badge(s, 11.724, 4.455, 0.424, C.white, 'fp', C.dark, [11.846, 4.544, 0.179, 0.244]);
  dotsBadge(s, 3.748, 1.892, 0.302, C.dark, C.white);
  dotsBadge(s, 3.748, 5.476, 0.302, C.dark, C.white);
  dotsBadge(s, 7.137, 3.674, 0.302, C.dark, C.white);
  txt(s, 'Statistics Point.', 6.109, 5.913, 1.054, 0.236, { size: 8 });
  txt(s, '12.4M+', 4.983, 5.812, 1.156, 0.438, { size: 20, font: F.head });
}

function slide09(s) { // Customer Engagement
  navbar(s);
  card(s, 9.264, 2.267, 3.227, 4.75, 0.358, { line: { color: C.white, width: 0.31 } });
  txt(s, 'Customer Engagement', 4.899, 2.436, 3.797, 1.448, { size: 40, font: F.head, align: 'center' });
  txt(s, 'SaaS.', 9.932, 2.516, 0.903, 0.304, { size: 12, font: F.head });
  circle(s, 9.694, 2.597, 0.139, 0.141, { fill: C.white });
  txt(s, 'Arcu odio ut sem nulla pharetra silent. Viverra adipiscing at in tellus.', 9.639, 2.943, 2.609, 0.58, { size: 10, ls: 1.5 });
  txt(s, 'SaaS.', 9.932, 4.203, 0.903, 0.302, { size: 12, font: F.head });
  circle(s, 9.694, 4.285, 0.139, 0.139, { fill: C.white });
  txt(s, 'Arcu odio ut sem nulla pharetra silent. Viverra adipiscing at in tellus.', 9.639, 4.63, 2.609, 0.58, { size: 10, ls: 1.5 });
  txt(s, 'SaaS.', 9.932, 5.769, 1.002, 0.304, { size: 12, font: F.head });
  circle(s, 9.694, 5.851, 0.139, 0.139, { fill: C.white });
  txt(s, 'Arcu odio ut sem nulla pharetra silent. Viverra adipiscing at in tellus.', 9.639, 6.196, 2.609, 0.58, { size: 10, ls: 1.5 });
  txt(s, 'Sed vulputate mi sit amet mauris commodo quis imperdiet. Tempor orci eu lobortis elementum \nnibh tellus molestie nunc non. ', 5.014, 4.29, 3.568, 0.833, { size: 10, align: 'center', ls: 1.5 });
  txt(s, '95%', 5.481, 5.825, 0.835, 0.405, { size: 18, font: F.head, align: 'center' });
  txt(s, 'Customer Happy', 5.22, 6.179, 1.356, 0.326, { size: 10, align: 'center', ls: 1.5 });
  txt(s, '5.0', 7.47, 5.825, 0.696, 0.405, { size: 18, font: F.head, align: 'center' });
  txt(s, 'Client Ratings', 7.267, 6.179, 1.102, 0.326, { size: 10, align: 'center', ls: 1.5 });
  card(s, 0.842, 2.267, 3.394, 1.446, 0.241, { fill: C.lime });
  txt(s, 'Definition & Key Features.', 1.002, 2.49, 2.661, 0.321, { size: 13, font: F.head, color: C.dark });
  txt(s, 'Neque porro quisquam est, qui dol\norem ipsum quia dolor sit amet.', 1.002, 2.964, 2.321, 0.533, { size: 9, color: C.dark, ls: 1.5 });
  card(s, 0.858, 5.582, 3.392, 1.446, 0.241, { fill: C.pink });
  txt(s, 'Benefits of Businesses.', 1.017, 5.804, 2.66, 0.321, { size: 13, font: F.head, color: C.dark });
  txt(s, 'Neque porro quisquam est, qui dol\norem ipsum quia dolor sit amet.', 1.017, 6.278, 2.319, 0.533, { size: 9, color: C.dark, ls: 1.5 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.', 1.58, 4.592, 1.747, 0.484, { size: 8, ls: 1.5 });
  txt(s, 'The Details.', 1.58, 4.219, 1.186, 0.373, { size: 12, font: F.head, ls: 1.5 });
  circle(s, 1.054, 4.531, 0.363, 0.363, { fill: C.white });
  flower(s, 1.142, 4.63, 0.186, 0.18, C.purple);
  fingerprint(s, 3.697, 6.286, 0.38, 0.517, C.dark);
  badge(s, 3.693, 3.036, 0.389, C.dark, 'flower', C.pink, [3.788, 3.142, 0.198, 0.193]);
}

function slide10(s) { // Cloud Computing Trends
  navbar(s);
  imageBox(s, 1.269, 3.797, 6.502, 3.108);
  card(s, 4.96, 1.712, 3.394, 1.674, 0.279, { fill: C.pink });
  dotsBadge(s, 7.767, 1.844, 0.418, C.dark, C.white);
  txt(s, 'Serverless Architectures.', 5.155, 1.951, 1.766, 0.538, { size: 13, font: F.head, color: C.dark });
  txt(s, 'id est laborum et dolorum fuga. Et harum quidem rerum facilis.', 5.155, 2.622, 2.109, 0.533, { size: 9, color: C.dark, ls: 1.5 });
  card(s, 1.257, 1.71, 3.394, 1.674, 0.279, { fill: C.lime, line: { color: C.white, width: 0.31 } });
  dotsBadge(s, 4.062, 1.842, 0.418, C.dark, C.white);
  txt(s, 'Edge Computing.', 1.451, 1.95, 1.337, 0.538, { size: 13, font: F.head, color: C.dark });
  txt(s, 'id est laborum et dolorum fuga. Et harum quidem rerum facilis.', 1.451, 2.62, 2.109, 0.533, { size: 9, color: C.dark, ls: 1.5 });
  card(s, 8.66, 1.712, 3.394, 1.674, 0.279, { line: { color: C.white, width: 0.31 } });
  dotsBadge(s, 11.465, 1.844, 0.418, C.dark, C.white);
  txt(s, 'Hybrid & Multi Cloud Adoption.', 8.854, 1.951, 1.764, 0.538, { size: 13, font: F.head });
  txt(s, 'id est laborum et dolorum fuga. Et harum quidem rerum facilis.', 8.854, 2.622, 2.109, 0.533, { size: 9, ls: 1.5 });
  badge(s, 6.917, 3.955, 0.63, C.white, 'fp', C.dark, [7.099, 4.088, 0.267, 0.364]);
  txt(s, 'Cloud Computing Trends.', 8.474, 3.837, 3.972, 1.111, { size: 30, font: F.head });
  txt(s, 'similique sunt in culpa qui officia deserunt mollitia animi, id est laboret dolorum fuga.', 8.51, 5.158, 3.194, 0.531, { size: 9, ls: 1.5 });
  pill(s, 8.595, 6.186, 1.236, 0.323, { fill: C.lime });
  txt(s, 'Next Slide', 8.677, 6.156, 1.078, 0.349, { size: 11, font: F.head, color: C.dark, align: 'center', ls: 1.5 });
  pill(s, 10.467, 6.186, 1.236, 0.323, { fill: C.pink });
  txt(s, 'Learn More', 10.578, 6.196, 1.04, 0.285, { size: 11, font: F.head, color: C.dark, align: 'center' });
  fingerprint(s, 3.963, 2.525, 0.428, 0.582, C.dark);
  fingerprint(s, 7.712, 2.524, 0.429, 0.583, C.dark);
  fingerprint(s, 11.378, 2.525, 0.428, 0.582, C.white);
}

function slide11(s) { // Cloud Computing Examples
  navbar(s);
  imageBox(s, 0.618, 4.592, 4.465, 1.674);
  card(s, 9.411, 2.41, 3.392, 1.674, 0.279, { fill: C.lime });
  pieChart(s, 11.696, 3.076, 1.318, 0.878, [C.pink, C.white]);
  txt(s, 'Cloud Computing Examples.', 0.45, 2.161, 5.259, 1.448, { size: 40, font: F.head });
  txt(s, 'optio cumque nihil impedit quo minus id quod maxime placeat facere possimus.', 5.925, 3.054, 2.818, 0.533, { size: 9, ls: 1.5 });
  dotsBadge(s, 12.222, 2.507, 0.418, C.dark, C.white);
  txt(s, 'Infrastructure Scalability.', 9.606, 2.649, 1.502, 0.538, { size: 13, font: F.head, color: C.dark });
  txt(s, 'But I must explain to you how all\nthis mistaken idea of denouncing.', 9.606, 3.319, 2.262, 0.533, { size: 9, color: C.dark, ls: 1.5 });
  card(s, 5.526, 4.592, 3.394, 1.674, 0.279, { line: { color: C.white, width: 0.31 } });
  dotsBadge(s, 8.337, 4.689, 0.418, C.dark, C.white);
  txt(s, 'Data Storage Backups.', 5.72, 4.832, 1.764, 0.538, { size: 13, font: F.head });
  txt(s, 'id est laborum et dolorum fuga. Et harum quidem rerum facilis.', 5.72, 5.502, 2.109, 0.533, { size: 9, ls: 1.5 });
  pieChart(s, 7.811, 5.257, 1.318, 0.88, [C.pink, C.white]);
  card(s, 9.41, 4.592, 3.394, 1.674, 0.279, { fill: C.pink });
  dotsBadge(s, 12.222, 4.689, 0.418, C.dark, C.white);
  txt(s, 'Data Storage Backups.', 9.604, 4.832, 1.766, 0.538, { size: 13, font: F.head, color: C.dark });
  txt(s, 'id est laborum et dolorum fuga. Et harum quidem rerum facilis.', 9.604, 5.502, 2.109, 0.533, { size: 9, color: C.dark, ls: 1.5 });
  pieChart(s, 11.694, 5.257, 1.318, 0.88, [C.purple, C.white]);
  badge(s, 4.503, 4.698, 0.441, C.white, 'fp', C.dark, [4.63, 4.791, 0.187, 0.255]);
  pill(s, 3.932, 3.116, 1.236, 0.323, { fill: C.pink });
  txt(s, 'Learn More', 4.042, 3.127, 1.042, 0.285, { size: 11, font: F.head, color: C.dark, align: 'center' });
  txt(s, 'The Details.', 6.658, 2.45, 1.186, 0.373, { size: 12, font: F.head, ls: 1.5 });
  badge(s, 5.924, 2.507, 0.363, C.white, 'flower', C.pink, [6.012, 2.605, 0.185, 0.18]);
  badge(s, 8.288, 5.516, 0.363, C.dark, 'flower', C.pink, [8.377, 5.614, 0.185, 0.18]);
  badge(s, 12.172, 5.516, 0.363, C.dark, 'flower', C.pink, [12.261, 5.614, 0.185, 0.18]);
  badge(s, 12.174, 3.333, 0.363, C.dark, 'flower', C.pink, [12.263, 3.431, 0.185, 0.18]);
}

function slide12(s) { // Impact of Digital Transformation
  navbar(s);
  imageBox(s, 8.356, 3.127, 4.372, 3.891);
  txt(s, 'Impact of Digital Transformation.', 1.866, 1.655, 9.601, 0.774, { size: 40, font: F.head, align: 'center' });
  card(s, 0.587, 3.134, 3.394, 1.674, 0.279, { fill: C.pink });
  txt(s, 'Competitive Advantage.', 0.731, 3.446, 1.464, 0.538, { size: 13, font: F.head, color: C.dark });
  txt(s, 'nisi ut aliquid ex ea commodi cons equatur Quis autem vel eum.', 0.731, 4.059, 2.201, 0.533, { size: 9, color: C.dark, ls: 1.5 });
  dotsBadge(s, 3.42, 3.212, 0.418, C.dark, C.white);
  card(s, 4.49, 3.127, 3.394, 1.674, 0.279, { line: { color: C.white, width: 0.31 } });
  txt(s, 'Operational Efficiency.', 4.644, 3.424, 1.627, 0.538, { size: 13, font: F.head });
  txt(s, 'nisi ut aliquid ex ea commodi cons equatur Quis autem vel eum.', 4.644, 4.036, 2.201, 0.531, { size: 9, ls: 1.5 });
  dotsBadge(s, 7.38, 3.212, 0.418, C.dark, C.white);
  card(s, 0.606, 5.342, 3.394, 1.674, 0.279, { fill: C.lime });
  txt(s, 'Enhanced Custom-\ner Experiences.', 0.762, 5.639, 2.024, 0.538, { size: 13, font: F.head, color: C.dark });
  txt(s, 'nisi ut aliquid ex ea commodi cons equatur Quis autem vel eum.', 0.762, 6.252, 2.201, 0.533, { size: 9, color: C.dark, ls: 1.5 });
  dotsBadge(s, 3.497, 5.427, 0.418, C.dark, C.white);
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.', 5.859, 5.995, 1.747, 0.484, { size: 8, ls: 1.5 });
  txt(s, 'The Details.', 5.84, 5.622, 1.184, 0.373, { size: 12, font: F.head, ls: 1.5 });
  badge(s, 5.064, 5.934, 0.363, C.white, 'flower', C.pink, [5.154, 6.033, 0.185, 0.18]);
  badge(s, 12.066, 3.326, 0.472, C.white, 'fire', C.dark, [12.226, 3.432, 0.163, 0.24]);
  fingerprint(s, 3.351, 4.014, 0.429, 0.582, C.dark);
  fingerprint(s, 3.434, 6.177, 0.429, 0.582, C.dark);
  fingerprint(s, 7.317, 4.013, 0.428, 0.582, C.white);
}

function slide13(s) { // Challenges of Digital Transformation
  navbar(s);
  imageBox(s, 9.431, 1.976, 3.418, 1.774);
  card(s, 5.035, 3.755, 3.394, 1.675, 0.279, { line: { color: C.white, width: 0.31 } });
  txt(s, 'Integration Complexity.', 5.194, 4.005, 1.491, 0.538, { size: 13, font: F.head });
  txt(s, 'Neque porro quisquam est, qui dol\norem ipsum quia dolor sit amet.', 5.194, 4.618, 2.368, 0.533, { size: 9, ls: 1.5 });
  dotsBadge(s, 7.884, 3.856, 0.418, C.dark, C.white);
  card(s, 0.491, 5.372, 3.394, 1.674, 0.279, { fill: C.lime });
  txt(s, 'Talent Gap.', 0.651, 5.62, 1.071, 0.54, { size: 13, font: F.head, color: C.dark });
  txt(s, 'Neque porro quisquam est, qui dolorem ipsum quia dolor sit.', 0.651, 6.234, 2.023, 0.531, { size: 9, color: C.dark, ls: 1.5 });
  dotsBadge(s, 3.339, 5.472, 0.418, C.dark, C.white);
  card(s, 9.455, 5.33, 3.394, 1.674, 0.279, { fill: C.pink });
  dotsBadge(s, 12.311, 5.431, 0.418, C.dark, C.white);
  txt(s, 'Security Concerns.', 9.658, 5.569, 1.521, 0.538, { size: 13, font: F.head, color: C.dark });
  txt(s, 'id est laborum et dolorum fuga. Et harum quidem rerum facilis.', 9.658, 6.241, 2.108, 0.531, { size: 9, color: C.dark, ls: 1.5 });
  txt(s, 'Challenges of Digital Transformation.                ', 0.424, 1.773, 6.057, 1.448, { size: 40, font: F.head });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.', 6.148, 6.181, 1.747, 0.486, { size: 8, ls: 1.5 });
  txt(s, 'The Details.', 6.127, 5.807, 1.186, 0.373, { size: 12, font: F.head, ls: 1.5 });
  badge(s, 5.689, 6.122, 0.363, C.white, 'flower', C.pink, [5.778, 6.219, 0.185, 0.18]);
  pill(s, 1.224, 4.474, 0.085, 0.233, { fill: C.orchid });
  pill(s, 1.347, 4.257, 0.085, 0.45, { fill: C.purple });
  pill(s, 1.47, 4.368, 0.085, 0.339, { fill: C.orchid });
  pill(s, 1.594, 4.55, 0.085, 0.156, { fill: C.purple });
  pill(s, 1.712, 4.306, 0.085, 0.403, { fill: C.orchid });
  txt(s, '198.789', 1.892, 4.08, 1.549, 0.505, { size: 24, font: F.head });
  txt(s, 'Data Statistics.', 1.903, 4.582, 1.549, 0.236, { size: 8 });
  badge(s, 12.33, 2.073, 0.444, C.white, 'fire', C.dark, [12.479, 2.174, 0.155, 0.224]);
  pill(s, 5.269, 2.809, 1.234, 0.325, { fill: C.purple });
  txt(s, 'Learn More', 5.378, 2.819, 1.04, 0.285, { size: 11, font: F.head, align: 'center' });
  pieChart(s, 7.392, 4.451, 1.318, 0.88, [C.pink, C.white]);
  badge(s, 7.87, 4.71, 0.363, C.dark, 'flower', C.pink, [7.959, 4.808, 0.185, 0.18]);
  pieChart(s, 2.694, 6.097, 1.318, 0.878, [C.pink, C.white]);
  badge(s, 3.172, 6.354, 0.363, C.dark, 'flower', C.pink, [3.261, 6.452, 0.185, 0.18]);
  pieChart(s, 11.786, 6.045, 1.318, 0.878, [C.lime, C.white]);
  badge(s, 12.264, 6.302, 0.363, C.dark, 'flower', C.pink, [12.353, 6.4, 0.185, 0.18]);
}

function slide14(s) { // Case Studies
  navbar(s);
  imageBox(s, 10.106, 1.78, 2.743, 5.238);
  txt(s, 'Case Studies.', 6.425, 1.78, 2.809, 1.446, { size: 40, font: F.head });
  txt(s, 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloren mque laudantium, totam rem aperiam.', 6.425, 3.26, 3.24, 0.759, { size: 9, ls: 1.5 });
  card(s, 0.504, 1.779, 2.436, 3.201, 0.313, { fill: C.pink });
  txt(s, 'Brisk\nCompany.', 0.74, 3.128, 1.227, 0.54, { size: 13, font: F.head, color: C.dark });
  txt(s, 'PLACEHOLDER', 0.74, 3.766, 2.201, 0.759, { size: 9, color: C.dark, ls: 1.5 });
  dotsBadge(s, 2.269, 1.979, 0.486, C.dark, C.white);
  txt(s, '89%', 1.036, 5.672, 1.37, 0.722, { size: 37, font: F.head, align: 'center' });
  txt(s, 'Strength Rate.', 1.102, 6.354, 1.238, 0.328, { size: 10, align: 'center', ls: 1.5 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.', 7.497, 6.137, 1.747, 0.484, { size: 8, ls: 1.5 });
  txt(s, 'The Details.', 7.477, 5.764, 1.184, 0.373, { size: 12, font: F.head, ls: 1.5 });
  badge(s, 7.038, 6.076, 0.363, C.white, 'flower', C.pink, [7.127, 6.175, 0.185, 0.18]);
  card(s, 3.417, 1.77, 2.436, 3.201, 0.295, { fill: C.lime });
  txt(s, 'Mandak Company.', 3.651, 3.128, 1.227, 0.54, { size: 13, font: F.head, color: C.dark });
  txt(s, 'PLACEHOLDER', 3.651, 3.766, 2.201, 0.759, { size: 9, color: C.dark, ls: 1.5 });
  dotsBadge(s, 5.181, 1.979, 0.486, C.dark, C.white);
  txt(s, '67%', 3.95, 5.672, 1.37, 0.722, { size: 37, font: F.head, align: 'center' });
  txt(s, 'Strength Rate.', 4.016, 6.354, 1.238, 0.328, { size: 10, align: 'center', ls: 1.5 });
  fingerprint(s, 3.652, 1.994, 0.683, 0.929, C.dark);
  fingerprint(s, 0.764, 1.994, 0.683, 0.929, C.dark);
  pill(s, 6.425, 4.627, 1.236, 0.323, { fill: C.purple });
  txt(s, 'Next Slide', 6.507, 4.597, 1.078, 0.349, { size: 11, font: F.head, align: 'center', ls: 1.5 });
  pill(s, 8.297, 4.627, 1.236, 0.323, { fill: C.purple });
  txt(s, 'Learn More', 8.406, 4.637, 1.042, 0.285, { size: 11, font: F.head, align: 'center' });
  badge(s, 12.247, 1.917, 0.474, C.white, 'fire', C.dark, [12.406, 2.023, 0.163, 0.238]);
}

function slide15(s) { // Best Practices for Digital Transformation
  navbar(s, true);
  imageBox(s, 1.097, 2.109, 3.394, 2.668);
  txt(s, 'Best Practices for Digital Transformation.', 4.95, 2.033, 8.486, 1.448, { size: 40, font: F.head });
  txt(s, 'Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut odit. Nemo enim ipsam voluptatem quia voluptas sit aspernatur.', 5.05, 3.875, 5.387, 0.628, { size: 11, ls: 1.5 });
  card(s, 1.097, 5.227, 3.394, 1.674, 0.279, { fill: C.pink });
  txt(s, 'Establish Clear Goals.', 1.257, 5.476, 1.366, 0.538, { size: 13, font: F.head, color: C.dark });
  txt(s, 'But I must explain to you how allth\nis mistaken idea of denouncing.', 1.257, 6.09, 2.333, 0.533, { size: 9, color: C.dark, ls: 1.5 });
  dotsBadge(s, 3.931, 5.358, 0.418, C.dark, C.white);
  card(s, 4.95, 5.222, 3.394, 1.674, 0.279, { line: { color: C.white, width: 0.31 } });
  txt(s, 'Invest in Talent & Training.', 5.109, 5.47, 1.642, 0.538, { size: 13, font: F.head });
  txt(s, 'But I must explain to you how allth\nis mistaken idea of denouncing.', 5.109, 6.085, 2.333, 0.531, { size: 9, ls: 1.5 });
  dotsBadge(s, 7.783, 5.352, 0.418, C.dark, C.white);
  card(s, 8.842, 5.207, 3.394, 1.674, 0.279, { fill: C.lime });
  dotsBadge(s, 11.675, 5.337, 0.418, C.dark, C.white);
  txt(s, 'Foster a Culture Innovation.', 9.036, 5.446, 1.608, 0.538, { size: 13, font: F.head, color: C.dark });
  txt(s, 'id est laborum et dolorum fuga. Et harum quidem rerum facilis.', 9.036, 6.118, 2.109, 0.531, { size: 9, color: C.dark, ls: 1.5 });
  badge(s, 3.842, 2.314, 0.472, C.white, 'fire', C.dark, [4, 2.42, 0.163, 0.238]);
  badge(s, 3.804, 6.226, 0.479, C.dark, 'flower', C.pink, [3.921, 6.356, 0.244, 0.239]);
  badge(s, 7.738, 6.302, 0.363, C.lime, 'flower', C.pink, [7.827, 6.401, 0.185, 0.18]);
  // (embedded spreadsheet chart object in the original - visually empty)
  badge(s, 11.569, 6.226, 0.481, C.dark, 'flower', C.pink, [11.687, 6.356, 0.245, 0.239]);
  pill(s, 10.438, 2.998, 1.236, 0.325, { fill: C.purple });
  txt(s, 'Learn More', 10.547, 3.009, 1.042, 0.286, { size: 11, font: F.head, align: 'center' });
}

function slide16(s) { // Regulatory Considerations
  navbar(s);
  imageBox(s, 9.519, 2.049, 3.33, 4.969);
  card(s, 5.22, 2.05, 3.816, 2.224, 0.371, { fill: C.lime });
  dotsBadge(s, 8.434, 2.167, 0.47, C.dark, C.white);
  txt(s, 'Regulatory Considerations.', 0.438, 1.856, 4.156, 1.312, { size: 36, font: F.head });
  txt(s, 'Neque porro quisquam est, qui dolorem ipsum quia dolor sit amet, consectetur, adipisci velit, sed qui.', 0.438, 3.774, 3.545, 0.533, { size: 9, ls: 1.5 });
  txt(s, 'Legal Complications of New Technologies.', 5.382, 2.418, 2.264, 0.571, { size: 14, font: F.head, color: C.dark });
  txt(s, 'But I must explain to you how all\nthis mistaken idea of denouncing.', 5.382, 3.3, 2.264, 0.533, { size: 9, color: C.dark, ls: 1.5 });
  card(s, 5.22, 4.811, 3.816, 2.226, 0.371, { fill: C.pink });
  dotsBadge(s, 8.434, 4.929, 0.47, C.dark, C.white);
  txt(s, 'Data Privacy & Compliance.', 5.382, 5.181, 2.264, 0.571, { size: 14, font: F.head, color: C.dark });
  txt(s, 'But I must explain to you how all\nthis mistaken idea of denouncing.', 5.382, 6.128, 2.264, 0.533, { size: 9, color: C.dark, ls: 1.5 });
  fingerprint(s, 8.304, 3.415, 0.429, 0.583, C.dark);
  fingerprint(s, 8.305, 6.244, 0.428, 0.582, C.dark);
  badge(s, 12.25, 2.193, 0.472, C.white, 'fire', C.dark, [12.408, 2.299, 0.163, 0.24]);
  pill(s, 0.507, 5.878, 1.236, 0.325, { fill: C.pink });
  txt(s, 'Next Slide', 0.589, 5.851, 1.078, 0.349, { size: 11, font: F.head, color: C.dark, align: 'center', ls: 1.5 });
  pill(s, 2.378, 5.878, 1.236, 0.325, { fill: C.lime });
  txt(s, 'Learn More', 2.49, 5.889, 1.04, 0.286, { size: 11, font: F.head, color: C.dark, align: 'center' });
}

function slide17(s) { // Future Trends
  navbar(s);
  txt(s, 'Future Trends.', 4.915, 1.911, 4.503, 0.774, { size: 40, font: F.head });
  txt(s, 'Quis autem vel eum iure reprehenderit qui in ea voluptate velit esse quam nihil molestiae.', 4.915, 2.799, 3.226, 0.533, { size: 9, ls: 1.5 });
  card(s, 1.196, 1.865, 3.392, 1.674, 0.279, { fill: C.lime });
  txt(s, 'Integration of AI, Blockchain, & Cloud.', 1.34, 2.177, 2.014, 0.538, { size: 13, font: F.head, color: C.dark });
  txt(s, 'Ut enim ad minima veniam, quis trum exercitationem ullam.', 1.34, 2.792, 2.2, 0.531, { size: 9, color: C.dark, ls: 1.5 });
  dotsBadge(s, 4.014, 1.97, 0.418, C.dark, C.white);
  card(s, 4.998, 3.849, 3.394, 1.675, 0.279, { line: { color: C.white, width: 0.31 } });
  txt(s, 'Emergence of Industry Sulutions.', 5.155, 4.146, 1.955, 0.538, { size: 13, font: F.head });
  txt(s, 'Ut enim ad minima veniam, quis trum exercitationem ullam.', 5.155, 4.76, 2.2, 0.531, { size: 9, ls: 1.5 });
  dotsBadge(s, 7.875, 3.962, 0.418, C.dark, C.white);
  card(s, 8.819, 3.877, 3.394, 1.675, 0.279, { fill: C.pink });
  txt(s, 'Ethical & Social Implications.', 8.976, 4.174, 2.024, 0.538, { size: 13, font: F.head, color: C.dark });
  txt(s, 'nisi ut aliquid ex ea commodi cons equatur Quis autem vel eum.', 8.976, 4.788, 2.2, 0.531, { size: 9, color: C.dark, ls: 1.5 });
  dotsBadge(s, 11.696, 3.99, 0.418, C.dark, C.white);
  badge(s, -4.905, -4.391, 0.521, C.white, 'fp', C.dark, [-4.754, -4.281, 0.222, 0.301]);
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.', 6.009, 6.323, 1.747, 0.486, { size: 8, ls: 1.5 });
  txt(s, 'The Details.', 5.988, 5.95, 1.186, 0.373, { size: 12, font: F.head, ls: 1.5 });
  badge(s, 5.55, 6.264, 0.361, C.white, 'flower', C.purple, [5.639, 6.362, 0.185, 0.18]);
  pieChart(s, 8.054, 6.095, 1.082, 0.72, [C.pink, C.white]);
  txt(s, '12.4', 8.854, 6.083, 0.783, 0.438, { size: 20, font: F.head });
  txt(s, 'Statistics Point.', 8.854, 6.519, 1.144, 0.252, { size: 9 });
  txt(s, 'AI users are increasing.', 10.875, 6.441, 1.056, 0.37, { size: 8 });
  txt(s, '16.9M+', 10.847, 6.007, 1.158, 0.438, { size: 20, font: F.head });
  badge(s, 4, 3.995, 0.472, C.white, 'fire', C.dark, [4.16, 4.101, 0.163, 0.238]);
  pill(s, 9.08, 2.922, 1.236, 0.325, { fill: C.purple });
  txt(s, 'Next Slide', 9.161, 2.892, 1.078, 0.351, { size: 11, font: F.head, align: 'center', ls: 1.5 });
  pill(s, 10.951, 2.922, 1.236, 0.325, { fill: C.purple });
  txt(s, 'Learn More', 11.061, 2.932, 1.042, 0.286, { size: 11, font: F.head, align: 'center' });
  imageBox(s, 1.196, 3.851, 3.392, 2.92);
}

function slide18(s) { // In Conclusion
  navbar(s);
  imageBox(s, 1.28, 1.609, 4.901, 2.012);
  imageBox(s, 7.078, 1.609, 4.901, 3.543);
  txt(s, 'In Conclusion', 1.28, 3.981, 4.438, 0.774, { size: 40, font: F.head });
  txt(s, 'Accumsan in nisl nisi scelerisque eu ultrices vitae auctor eu. Pharetra pharetra massa massa ultricies mi quis hendrerit dolor magna. Arcu non sodales neque sodales ut etiam sit amet nisl.', 1.28, 4.988, 4.615, 0.825, { size: 10, ls: 1.5 });
  card(s, 9.965, 5.891, 2.014, 1.082, 0.23, { fill: C.pink });
  card(s, 7.142, 5.891, 2.012, 1.082, 0.25, { fill: C.lime });
  pieChart(s, 7.078, 6.113, 1.082, 0.722, [C.purple, C.white]);
  txt(s, '12.4', 7.878, 6.101, 0.783, 0.438, { size: 20, font: F.head, color: C.dark });
  txt(s, 'Statistics Point.', 7.878, 6.536, 1.144, 0.252, { size: 9, color: C.dark });
  txt(s, 'AI users are increasing.', 11.097, 6.247, 1.054, 0.37, { size: 8, color: C.dark });
  txt(s, '16.9M+', 10.035, 6.226, 1.158, 0.438, { size: 20, font: F.head, color: C.dark });
  badge(s, 5.587, 1.76, 0.472, C.white, 'fire', C.dark, [5.745, 1.866, 0.163, 0.24]);
  badge(s, 11.396, 1.76, 0.472, C.white, 'fire', C.dark, [11.554, 1.866, 0.163, 0.24]);
  pill(s, 1.28, 6.329, 1.236, 0.322, { fill: C.purple });
  txt(s, 'Next Slide', 1.361, 6.299, 1.078, 0.348, { size: 11, font: F.head, align: 'center', ls: 1.5 });
  pill(s, 4.821, 6.326, 1.234, 0.325, { fill: C.purple });
  txt(s, 'Learn More', 4.931, 6.337, 1.04, 0.286, { size: 11, font: F.head, align: 'center' });
}

function slide19(s) { // Q&A Session
  navbar(s);
  imageBox(s, 0.503, 2.066, 3.274, 4.885);
  imageBox(s, 9.575, 2.066, 3.274, 4.885);
  txt(s, 'Q&A Session', 3.943, 2.59, 5.448, 0.773, { size: 40, font: F.head, align: 'center' });
  txt(s, 'Open floor for questions and answers', 4.474, 3.719, 4.333, 0.417, { size: 14, font: F.head, align: 'center', ls: 1.5 });
  txt(s, 'Accumsan in nisl nisi scelerisque eu ultrices vitae auctor eu. Pharetra pharetra massa massa ultricies mi quisoto hendrerit dolor magna non sodales neque sodales.', 4.474, 4.491, 4.333, 0.825, { size: 10, align: 'center', ls: 1.5 });
  badge(s, 3.151, 2.236, 0.472, C.white, 'fire', C.dark, [3.309, 2.342, 0.165, 0.24]);
  badge(s, 12.227, 2.236, 0.472, C.white, 'fire', C.dark, [12.385, 2.342, 0.165, 0.24]);
  pill(s, 5.028, 5.849, 1.236, 0.323, { fill: C.purple });
  txt(s, 'Next Slide', 5.109, 5.819, 1.078, 0.349, { size: 11, font: F.head, align: 'center', ls: 1.5 });
  pill(s, 7.359, 5.849, 1.236, 0.323, { fill: C.purple });
  txt(s, 'Learn More', 7.47, 5.859, 1.04, 0.285, { size: 11, font: F.head, align: 'center' });
}

function slide20(s) { // Thank you
  navbar(s);
  imageBox(s, 1.163, 1.688, 11.167, 2.036);
  txt(s, 'Thankyou', 1.003, 5.047, 10.233, 2.036, { size: 115, font: F.head });
  txt(s, 'Et netus et malesuada fames ac turpis egestas. Vestibulum sed arcu non odio euismod lacinia at. Elementum facilisis leo vel.', 3.021, 4.344, 4.441, 0.58, { size: 10, ls: 1.5 });
  pill(s, 1.163, 4.427, 1.227, 0.415, { fill: C.purple });
  txt(s, [{ text: 'Learn ', options: { fontSize: 11, fontFace: F.head, color: C.white, align: 'center', lineSpacingMultiple: 1.5 } }, { text: 'More', options: { fontSize: 12, fontFace: F.head, color: C.white, align: 'center', lineSpacingMultiple: 1.5 } }], 1.189, 4.408, 1.175, 0.372);
  badge(s, 11.441, 4.425, 0.418, C.pink, 'fire', C.dark, [11.582, 4.528, 0.144, 0.212]);
  badge(s, 11.896, 4.417, 0.434, C.lime, 'fp', C.dark, [12.021, 4.509, 0.185, 0.251]);
  badge(s, 10.986, 4.425, 0.417, C.purple, 'flower', C.white, [11.089, 4.531, 0.213, 0.207]);
  arrow(s, 11.606, 5.766, 0.601, 0.599, 223.82, C.white);
}

const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20];


/* ------------------------------------------------------------------- output */
const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'DECK', width: 13.333, height: 7.5 });
pptx.layout = 'DECK';
pptx.theme = { headFontFace: F.head, bodyFontFace: F.body };

SLIDES.forEach(function (build) {
  const s = pptx.addSlide();
  s.background = { color: C.dark };
  build(s);
});

pptx.writeFile({ fileName: path.join(__dirname, path.basename(__filename, '.js') + '.pptx') })
  .then(function (f) { console.log('wrote ' + f); });
