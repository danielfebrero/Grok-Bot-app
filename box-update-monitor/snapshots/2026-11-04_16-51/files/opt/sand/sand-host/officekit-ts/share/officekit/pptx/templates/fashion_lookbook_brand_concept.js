/*
 * LARSO - Lookbook Concept  (30 slides, 13.333 x 7.5 in)
 * Rebuilt with pptxgenjs.  Raster photos in the source deck are replaced by
 * flat colour placeholders; everything else is drawn with native shapes/text.
 *
 *   node 0716d450-4ad8-4698-a7ac-23cc0bda03f8_grok_final.js   ->   0716d450-4ad8-4698-a7ac-23cc0bda03f8_grok_final.pptx
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette --
const C = {
  purple:   '655396',   // theme accent1
  mid:      'A194C4',   // accent1 @ 60% luminance
  light:    'C0B8D8',   // accent1 @ 40% luminance
  lilac:    'E0DBEB',   // accent1 @ 20% luminance - the big background panels
  deep:     '322A4B',   // accent1 darkened - cover wordmark
  ink:      '404040',   // headline grey
  sub:      '595959',   // sub-heading grey
  body:     '808080',   // body copy grey
  track:    'D9D9D9',   // progress-bar track / arrows
  foot:     'BFBFBF',
  footBold: 'A6A6A6',
  navText:  'F2F2F2',
  white:    'FFFFFF',
  photo:    '464646',   // image placeholder fill
  photoTag: '8C8C8C',
};

const HEAD = 'Alata';   // major (headline) typeface
const BODY = 'Roboto';  // minor (body) typeface

// Body copy repeated throughout the deck.
const T = {
  A: 'PLACEHOLDER',
  B: 'PLACEHOLDER',
  C: 'PLACEHOLDER',
  D: 'PLACEHOLDER',
  E: 'PLACEHOLDER',
  F: 'PLACEHOLDER',
  G: 'Asuscipit eros iste metus auctor id dapibus quam aliq',
  H: 'PLACEHOLDER',
  I: 'PLACEHOLDER',
  J: 'PLACEHOLDER',
  K: 'PLACEHOLDER',
  L: 'PLACEHOLDER',
};

// ------------------------------------------------------------------ atoms --

// Flat colour band / panel.
function band(s, x, y, w, h, color, o = {}) {
  s.addShape('rect', { x, y, w, h, fill: { color, transparency: o.transparency || 0 } });
}

// Placeholder standing in for a photograph in the original deck.
function photo(s, x, y, w, h) {
  s.addShape('rect', { x, y, w, h, fill: { color: C.photo } });
  s.addText('[image]', {
    x: x + 0.1, y: y + 0.08, w: 1, h: 0.2, align: 'left', valign: 'top', margin: 0,
    fontFace: BODY, fontSize: 11, color: C.photoTag,
  });
}

// Device mock-up: dark bezel (`outer`) wrapped around a screen placeholder
// (`screen`), each given as [x, y, w, h].  `rotate` reproduces the tilt of the
// photographed devices in the original deck.
function device(s, outer, screen, rotate = 0) {
  const [ox, oy, ow, oh] = outer;
  const [x, y, w, h] = screen;
  s.addShape('roundRect', {
    x: ox, y: oy, w: ow, h: oh, rotate,
    rectRadius: Math.min(ow, oh) * 0.06, fill: { color: '2B2B2B' },
  });
  s.addShape('rect', { x, y, w, h, rotate, fill: { color: C.photo } });
  s.addText('[image]', {
    x, y, w, h, rotate, margin: 0, align: 'center', valign: 'middle',
    fontFace: BODY, fontSize: 11, color: C.photoTag,
  });
}

// Big Alata headline; `lines` is one paragraph per array entry.
function title(s, lines, x, y, w, h, o = {}) {
  s.addText(lines.map((t, i) => ({ text: t, options: { breakLine: i < lines.length - 1 } })), {
    x, y, w, h, margin: o.inset ? [3.6, 7.2, 3.6, 7.2] : 0, valign: 'top', wrap: o.wrap === true,
    fontFace: HEAD, fontSize: o.fontSize || 40, color: o.color || C.ink,
    transparency: o.transparency, italic: o.italic, charSpacing: o.charSpacing,
    align: o.align || 'left', lineSpacing: (o.fontSize || 40) * 1.2,
  });
}

// Paragraph of body copy (Roboto 12 / 150% leading by default).
function body(s, text, x, y, w, h, o = {}) {
  const size = o.fontSize || 12;
  const lines = Array.isArray(text) ? text : [text];
  s.addText(lines.map((t, i) => ({ text: t, options: { breakLine: i < lines.length - 1 } })), {
    x, y, w, h, margin: o.inset ? [3.6, 7.2, 3.6, 7.2] : 0, valign: 'top', wrap: o.wrap !== false,
    fontFace: o.fontFace || BODY, fontSize: size, color: o.color || C.body,
    bold: o.bold === true, italic: o.italic, align: o.align || 'left',
    charSpacing: o.charSpacing, lineSpacingMultiple: o.lineSpacing || 1.5,
  });
}

// Small bold caption ("Description", "Skill Name", ...).
function label(s, text, x, y, w, h, o = {}) {
  body(s, text, x, y, w, h, Object.assign({ fontSize: 14, color: C.sub }, o));
}

// Thin connector line.
function rule(s, x, y, w, color) {
  s.addShape('line', { x, y, w, h: 0, line: { color, width: 1 } });
}

// White pictogram placeholder standing in for a vector icon.
function glyph(s, x, y, w, h, shape, color) {
  s.addShape(shape, { x, y, w, h, fill: { color: color || C.white } });
}

// ------------------------------------------------------- shared furniture --

// Nav pill (top-left) + copyright footer, present on every slide.
function chrome(s) {
  s.addShape('rect', { x: 0, y: 0.106, w: 1.355, h: 0.278, fill: { color: C.purple } });
  s.addText('HOME  I  ABOUT', {
    x: 0.221, y: 0.136, w: 0.913, h: 0.18, margin: 0, wrap: false,
    fontFace: BODY, fontSize: 8, color: C.navText,
  });
  s.addText([
    { text: '\u00a9 2021  ', options: { fontFace: BODY, color: C.foot } },
    { text: 'larso ', options: { fontFace: HEAD, bold: true, color: C.footBold } },
    { text: 'all rights reserved', options: { fontFace: BODY, color: C.foot } },
  ], { x: 5.678, y: 7.175, w: 1.979, h: 0.168, margin: 0, wrap: false, align: 'center', fontSize: 10 });
}

// ------------------------------------------------------------- set pieces --

// Slide 26: exploded pie, four quarters of "Sales".
function pieChart(s, x, y, w, h) {
  s.addChart('pie', [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'],
    values: [8.2, 3.2, 1.4, 1.2] }], {
    x, y, w, h,
    chartColors: [C.purple, C.lilac, C.light, C.mid],
    dataBorder: { pt: 1.5, color: C.white },
    showLegend: false, showTitle: false, showValue: false,
    chartArea: { fill: { color: C.white, transparency: 100 } },
    plotArea: { fill: { color: C.white, transparency: 100 } },
  });
}

// Filled polygon from normalised (0-1) outline points.
function polygon(s, x, y, w, h, outline, o = {}) {
  const points = outline.map(([px, py]) => ({ x: px * w, y: py * h }));
  points.push({ close: true });
  s.addShape('custGeom', { x, y, w, h, points, fill: { color: o.color },
    line: o.line ? { color: o.line, width: 0.75 } : undefined });
}

// Slide 29: Africa silhouette with four highlighted countries and pin numbers.
function africaMap(s) {
  polygon(s, 8.228, 2.275, 3.802, 4.204, AFRICA, { color: 'F2F2F2' });
  polygon(s, 8.685, 2.289, 1.179, 1.263, PURPLE_0, { color: C.purple });
  polygon(s, 10.461, 3.33, 0.888, 1.096, PURPLE_1, { color: C.purple });
  polygon(s, 9.712, 2.595, 0.902, 0.902, LILAC_0, { color: C.lilac });
  polygon(s, 9.962, 4.343, 1.068, 0.999, LILAC_1, { color: C.lilac });
  MAP_PINS.forEach(([px, py, n]) => {
    s.addShape('ellipse', { x: px, y: py, w: 0.224, h: 0.224, fill: { color: C.white } });
    s.addText(n, { x: px, y: py, w: 0.224, h: 0.224, margin: 0, align: 'center', valign: 'middle',
      fontFace: BODY, fontSize: 7, bold: true, color: C.sub });
  });
}
const MAP_PINS = [[10.056, 2.933, '02'], [10.8, 3.772, '03'], [10.452, 4.69, '04'], [9.229, 2.897, '01']];

// Slide 29: circular badge holding a pictogram.
function dot(s, x, y, fill, mark, shape) {
  s.addShape('ellipse', { x, y, w: 0.385, h: 0.385, fill: { color: fill } });
  s.addShape(shape, { x: x + 0.1, y: y + 0.1, w: 0.182, h: 0.182, fill: { color: mark } });
}

// ----------------------------------------------------------- the 30 slides --

// Normalised map outlines (x, y as a fraction of each shape's own box).
const AFRICA = [[0.15, 0.026], [0.146, 0.053], [0.102, 0.086], [0.091, 0.149], [0.033, 0.195], [0, 0.261], [0.018, 0.317], [0.015, 0.396], [0.077, 0.449], [0.091, 0.482], [0.164, 0.518], [0.332, 0.485], [0.354, 0.512], [0.394, 0.512], [0.416, 0.528], [0.405, 0.584], [0.449, 0.63], [0.474, 0.7], [0.449, 0.795], [0.533, 0.957], [0.533, 0.99], [0.642, 0.977], [0.715, 0.908], [0.723, 0.875], [0.759, 0.851], [0.752, 0.805], [0.828, 0.749], [0.821, 0.617], [0.942, 0.502], [0.996, 0.389], [0.894, 0.409], [0.876, 0.373], [0.832, 0.34], [0.752, 0.185], [0.748, 0.099], [0.653, 0.102], [0.566, 0.073], [0.54, 0.086], [0.536, 0.116], [0.478, 0.102], [0.471, 0.083], [0.412, 0.073], [0.398, 0.003], [0.259, 0.01], [0.215, 0.036]];
const PURPLE_0 = [[0.447, 0.033], [0.294, 0.121], [0.353, 0.297], [0, 0.516], [0.012, 0.571], [0.624, 0.989], [0.741, 0.956], [0.988, 0.769], [0.894, 0.692], [0.859, 0.385], [0.753, 0.22], [0.776, 0.011]];
const PURPLE_1 = [[0.875, 0], [0.172, 0.013], [0.172, 0.114], [0.109, 0.139], [0.125, 0.354], [0, 0.519], [0.094, 0.734], [0.375, 0.949], [0.609, 0.987], [0.781, 0.886], [0.672, 0.734], [0.859, 0.456], [0.891, 0.266], [0.984, 0.203], [0.906, 0.139]];
const LILAC_0 = [[0.108, 0], [0, 0.215], [0.062, 0.646], [0.292, 0.769], [0.462, 0.723], [0.923, 0.985], [0.969, 0.954], [0.938, 0.062], [0.677, 0.015], [0.646, 0.169], [0.569, 0.231], [0.369, 0.138], [0.323, 0.046]];
const LILAC_1 = [[0.338, 0.028], [0.182, 0.528], [0, 0.639], [0.208, 0.625], [0.26, 0.736], [0.455, 0.681], [0.494, 0.889], [0.727, 0.917], [0.818, 0.986], [0.766, 0.917], [0.779, 0.764], [0.87, 0.694], [0.818, 0.639], [0.805, 0.375], [0.909, 0.167], [0.896, 0.097], [0.987, 0.069], [0.649, 0], [0.481, 0.069]];

function slide01(s) {
  photo(s, 0.583, 1.151, 12.167, 5.198);
  title(s, ['LARSO'], 2.362, 2.075, 8.609, 3.349, { fontSize: 199, color: C.lilac, transparency: 61, align: 'center' });
  band(s, 1.134, 4.319, 3.805, 3.201, C.lilac, { transparency: 20 });
  band(s, 10.171, -0.02, 1.9, 2.668, C.lilac, { transparency: 20 });
  title(s, ['LARSO'], 1.565, 4.631, 2.943, 1.144, { fontSize: 67.99, color: C.deep, align: 'center' });
  label(s, 'LOOKBOOK CONCEPT', 1.356, 5.931, 3.361, 0.236, { fontSize: 14, color: C.deep, charSpacing: 6, fontFace: HEAD, align: 'center', lineSpacing: 1, wrap: false });
}

function slide02(s) {
  title(s, ['LARSO'], 1.202, 1.968, 1.729, 0.673);
  label(s, ['Simple', 'Presentation', 'Template'], 1.202, 2.817, 1.499, 0.808, { fontSize: 16, color: C.sub, fontFace: HEAD, lineSpacing: 1, wrap: false });
  body(s, T.D, 1.202, 3.8, 4.782, 0.875, { align: 'justify' });
  body(s, T.H, 1.202, 4.96, 4.782, 0.572, { color: C.sub, bold: true, fontFace: HEAD, align: 'justify' });
  band(s, 9.188, 0, 4.148, 4.729, C.lilac);
  s.addText('2021', { x: 10.962, y: 2.283, w: 3.062, h: 0.9, rotate: 90, margin: 0,
      fontFace: HEAD, fontSize: 46, charSpacing: 4, color: C.mid, transparency: 58,
      align: 'center', valign: 'middle', wrap: false });
  photo(s, 7.646, 1.05, 4.606, 5.4);
}

function slide03(s) {
  band(s, 0, 3.817, 5.808, 3.683, C.lilac);
  photo(s, 0.662, 1.26, 3.375, 4.981);
  title(s, ['HALLO ', 'MY NAME', 'IS VICKY'], 4.51, 1.391, 2.512, 2.02);
  photo(s, 4.51, 4.174, 1.855, 2.044);
  body(s, T.A, 7.748, 3.235, 4.347, 0.875, { align: 'justify' });
  body(s, T.F, 7.75, 4.851, 3.629, 0.525, { fontSize: 11 });
  label(s, 'The Owner', 7.748, 4.413, 0.924, 0.314, { fontSize: 14, color: C.sub, bold: true, wrap: false });
  title(s, ['If you want to achieve excellence, you can get there today. As of this second, quit doing less-than-excellent work'], 7.748, 2.124, 4.347, 0.808, { fontSize: 16, color: C.purple, italic: true, wrap: true });
}

function slide04(s) {
  band(s, 0, 0, 5.808, 3.75, C.lilac);
  photo(s, 2.889, 0.942, 10.444, 4.122);
  title(s, ['WELCOME', 'MESSAGE'], 1.319, 1.696, 2.649, 1.346);
  body(s, T.E, 1.319, 5.446, 9.601, 0.572, { align: 'justify' });
}

function slide05(s) {
  band(s, 9.462, 0, 3.871, 4.981, C.lilac);
  photo(s, 9.012, 1.151, 3.375, 4.981);
  photo(s, 6.144, 2.641, 2.602, 3.84);
  title(s, ['OUR SIMPLE', 'CONCEPT'], 1.145, 1.696, 3.078, 1.346);
  body(s, T.A, 1.145, 3.318, 4.347, 0.875, { align: 'justify' });
  body(s, T.F, 1.147, 5.002, 2.901, 0.802, { fontSize: 11, align: 'justify' });
  label(s, 'Description', 1.145, 4.564, 0.978, 0.314, { fontSize: 14, color: C.sub, bold: true, wrap: false });
}

function slide06(s) {
  band(s, 0, 1.771, 7.49, 5.729, C.lilac);
  body(s, T.G, 4.637, 4.915, 2.291, 0.572, { align: 'justify' });
  label(s, 'Image Gallery', 4.637, 4.537, 1.222, 0.314, { fontSize: 14, color: C.sub, bold: true, wrap: false });
  body(s, T.G, 4.637, 3.7, 2.291, 0.572, { align: 'justify' });
  label(s, 'Meet Our Teams', 4.637, 3.322, 1.473, 0.314, { fontSize: 14, color: C.sub, bold: true, wrap: false });
  body(s, T.G, 4.637, 2.485, 2.291, 0.572, { align: 'justify' });
  label(s, 'Company Profile', 4.637, 2.106, 1.473, 0.314, { fontSize: 14, color: C.sub, bold: true, wrap: false });
  body(s, T.G, 4.615, 6.091, 2.291, 0.572, { align: 'justify' });
  label(s, 'Infographics', 4.615, 5.712, 1.115, 0.314, { fontSize: 14, color: C.sub, bold: true, wrap: false });
  photo(s, 0.81, 1.012, 3.264, 5.247);
  title(s, ['CONTENT', 'LIST'], 8.403, 2.061, 2.505, 1.346);
  body(s, T.B, 8.403, 3.687, 3.787, 0.875, { align: 'justify' });
  body(s, 'PLACEHOLDER', 8.403, 4.868, 3.787, 0.571, { color: C.sub, fontFace: HEAD, align: 'justify' });
}

function slide07(s) {
  band(s, 9.606, 3.388, 3.729, 4.112, C.lilac);
  photo(s, 6.827, 1.612, 3.734, 4.737);
  title(s, ['01'], 10.839, 4.488, 1.967, 2.322, { fontSize: 138, color: C.light, transparency: 38, align: 'center' });
  title(s, ['SECTION'], 11.032, 4.078, 1.581, 0.471, { fontSize: 28, color: C.light, transparency: 38, align: 'center' });
  title(s, ['COMPANY', 'PROFILE'], 1.261, 2.129, 2.742, 1.346);
  body(s, T.A, 1.261, 3.751, 4.347, 0.875, { align: 'justify' });
  body(s, T.C, 1.261, 4.8, 4.253, 0.571, { bold: true, fontFace: HEAD, align: 'justify' });
}

function slide08(s) {
  band(s, -0.007, 2.236, 5.59, 5.264, C.lilac);
  photo(s, 0, 0, 3.583, 5.264);
  photo(s, 3.856, 3.387, 2.972, 2.962);
  title(s, ['FEW WORDS', 'ABOUT US'], 4.202, 1.547, 3.252, 1.346);
  body(s, T.D, 8.073, 2.208, 4.006, 1.178, { align: 'justify' });
  body(s, T.C, 8.073, 3.601, 4.006, 0.572, { color: C.sub, fontFace: HEAD, align: 'justify' });
  body(s, T.B, 8.073, 4.417, 4.006, 0.875, { align: 'justify' });
}

function slide09(s) {
  band(s, 7.417, 2.236, 5.917, 5.264, C.lilac);
  photo(s, 5.097, 1.034, 3.787, 5.433);
  title(s, ['SIMPLE', 'ELEGANCE', 'MODERN'], 9.498, 2.74, 2.589, 2.02);
  body(s, T.D, 1.087, 3.266, 3.177, 1.481, { align: 'justify' });
  label(s, 'Description', 1.087, 2.753, 0.978, 0.314, { fontSize: 14, color: C.sub, bold: true, wrap: false });
}

function slide10(s) {
  band(s, 0, 0, 2.444, 7.5, C.lilac);
  photo(s, 8.944, 1.528, 4.389, 4.444);
  photo(s, 0, 1.528, 3.167, 4.444);
  title(s, ['THE PLANNING', 'FUTURE'], 3.882, 2.501, 3.764, 1.346);
  body(s, T.A, 3.882, 4.123, 4.347, 0.875, { align: 'justify' });
}

function slide11(s) {
  band(s, 0, 2.153, 4.944, 5.347, C.lilac);
  photo(s, 4, 1.151, 3.389, 4.737);
  body(s, 'PLACEHOLDER', 1.147, 4.644, 2.298, 0.802, { fontSize: 11, align: 'justify' });
  label(s, 'Description', 1.145, 4.206, 0.978, 0.314, { fontSize: 14, color: C.sub, bold: true, wrap: false });
  body(s, T.B, 8.174, 4.456, 4.146, 0.875, { align: 'justify' });
  title(s, ['SEE OUR', 'AMAZING', 'WORK'], 8.174, 2.169, 2.575, 2.02);
}

function slide12(s) {
  band(s, 0, 5.861, 13.333, 1.639, C.lilac);
  photo(s, 4.528, 2.46, 8.806, 4.193);
  title(s, ['ABOUT', 'OUR', 'VISION'], 1.684, 2.74, 1.881, 2.02);
  label(s, '“Aliquam quaerat voluptatem. Ut enim ad minima veniam, quis nostrum exercitationem ullam suscipit.”', 4.528, 1.293, 5.861, 0.64, { fontSize: 16, color: C.sub, fontFace: HEAD, lineSpacing: 1, inset: true });
}

function slide13(s) {
  band(s, 0, 3.388, 3.729, 4.112, C.lilac);
  photo(s, 2.774, 1.612, 3.734, 4.737);
  title(s, ['02'], 0.262, 4.488, 2.502, 2.322, { fontSize: 138, color: C.light, transparency: 38, align: 'center' });
  title(s, ['SECTION'], 0.722, 4.078, 1.581, 0.471, { fontSize: 28, color: C.light, transparency: 38, align: 'center' });
  title(s, ['MEET OUR', 'TEAM'], 7.727, 2.129, 2.649, 1.346);
  body(s, T.A, 7.727, 3.751, 4.347, 0.875, { align: 'justify' });
  body(s, T.C, 7.727, 4.8, 4.253, 0.571, { bold: true, fontFace: HEAD, align: 'justify' });
}

function slide14(s) {
  band(s, 9.604, 0, 3.729, 7.5, C.lilac);
  band(s, 2.453, 4.355, 4.478, 0.21, C.track);
  band(s, 2.453, 5.199, 4.478, 0.21, C.track);
  band(s, 1.443, 4.355, 4.563, 0.21, C.purple);
  band(s, 1.445, 5.199, 3.835, 0.21, C.purple);
  label(s, '90%', 7.022, 4.359, 0.312, 0.202, { fontSize: 12, align: 'justify', lineSpacing: 1, wrap: false });
  label(s, '75%', 7.022, 5.203, 0.312, 0.202, { fontSize: 12, align: 'justify', lineSpacing: 1, wrap: false });
  label(s, 'Skill Name', 1.443, 4.042, 0.812, 0.202, { fontSize: 12, color: C.ink, bold: true, lineSpacing: 1, wrap: false });
  label(s, 'Skill Name', 1.443, 4.885, 0.812, 0.202, { fontSize: 12, color: C.ink, bold: true, lineSpacing: 1, wrap: false });
  label(s, '@Insert Name Here', 1.415, 5.796, 1.462, 0.202, { fontSize: 12, align: 'justify', lineSpacing: 1, wrap: false });
  label(s, 'your_domain@mail.com', 3.176, 5.796, 1.807, 0.202, { fontSize: 12, lineSpacing: 1, wrap: false });
  photo(s, 8.118, 1.208, 4.132, 5.084);
  title(s, ['MEET THE', 'LEADERSHIP'], 1.415, 1.208, 2.991, 1.346);
  body(s, T.B, 1.415, 2.834, 3.787, 0.875, { align: 'justify' });
}

function slide15(s) {
  photo(s, 0, 1.25, 2.93, 3.389);
  photo(s, 6.098, 1.25, 2.944, 3.389);
  title(s, ['THE', 'CREATIVE', 'TEAM'], 3.292, 1.935, 2.458, 2.02, { align: 'center' });
  photo(s, 9.252, 1.25, 2.944, 3.389);
  body(s, T.E, 1.866, 5.446, 9.601, 0.572, { align: 'center' });
  band(s, 12.405, 1.25, 0.928, 3.389, C.lilac);
}

function slide16(s) {
  band(s, 0, 3.847, 4.653, 3.653, C.lilac);
  photo(s, 9.696, 3.13, 2.693, 3.1);
  photo(s, 6.779, 3.13, 2.693, 3.1);
  photo(s, 0.943, 3.13, 2.693, 3.1);
  title(s, ['SUPPORTING ', 'TEAM'], 0.943, 1.208, 3.496, 1.346);
  body(s, 'PLACEHOLDER', 5.026, 1.269, 6.863, 1.178, { align: 'justify' });
  photo(s, 3.861, 3.13, 2.693, 3.1);
}

function slide17(s) {
  band(s, 0, 0, 8.278, 7.5, C.lilac);
  photo(s, 0, 0.799, 9.097, 4.459);
  title(s, ['BREAK', 'SESSION'], 9.902, 1.896, 2.174, 1.346);
  label(s, ['TAKE ONE HOUR ', 'TO BREAK'], 9.902, 3.495, 1.523, 0.666, { fontSize: 14, color: C.sub, italic: true, fontFace: HEAD, wrap: false });
  body(s, 'PLACEHOLDER', 1.623, 5.654, 7.474, 0.572, { align: 'justify' });
}

function slide18(s) {
  band(s, 9.606, 3.388, 3.729, 4.112, C.lilac);
  photo(s, 6.827, 1.612, 3.734, 4.737);
  title(s, ['03'], 10.58, 4.488, 2.484, 2.322, { fontSize: 138, color: C.light, transparency: 38, align: 'center' });
  title(s, ['SECTION'], 11.032, 4.078, 1.581, 0.471, { fontSize: 28, color: C.light, transparency: 38, align: 'center' });
  title(s, ['IMAGE', 'GALLERY'], 1.261, 2.129, 2.218, 1.346);
  body(s, T.A, 1.261, 3.751, 4.347, 0.875, { align: 'justify' });
  body(s, T.C, 1.261, 4.8, 4.253, 0.571, { bold: true, fontFace: HEAD, align: 'justify' });
}

function slide19(s) {
  band(s, 0, 2.519, 5.583, 4.981, C.lilac);
  photo(s, 0.909, 1.071, 2.879, 4.83);
  photo(s, 3.955, 3.031, 2.879, 3.579);
  title(s, ['OUR ', 'PORTOFOLIO'], 7.687, 1.696, 3.482, 1.346);
  body(s, T.A, 7.687, 3.318, 4.347, 0.875, { align: 'justify' });
  body(s, T.F, 7.688, 5.002, 2.901, 0.802, { fontSize: 11, align: 'justify' });
  label(s, 'Description', 7.687, 4.564, 0.978, 0.314, { fontSize: 14, color: C.sub, bold: true, wrap: false });
}

function slide20(s) {
  band(s, 10.889, 0, 2.444, 7.5, C.lilac);
  device(s, [4.13, 1.27, 2.32, 4.9], [4.22, 1.36, 2.14, 4.72], -30);
  device(s, [10.394, 1.522, 1.967, 4.996], [10.477, 1.604, 1.884, 4.564]);
  body(s, T.F, 1.263, 4.646, 3.022, 0.802, { fontSize: 11 });
  label(s, 'Description', 1.262, 4.208, 0.978, 0.314, { fontSize: 14, color: C.sub, bold: true, wrap: false });
  title(s, ['MOCKUP', 'COLLECTION’S'], 6.347, 1.712, 3.725, 1.346);
}

function slide21(s) {
  band(s, 0, 0, 2.444, 7.5, C.lilac);
  device(s, [3.06, 0.98, 3.97, 2.87], [3.15, 1.07, 3.79, 2.69], 2);
  device(s, [1.14, 3.89, 3.88, 2.53], [1.23, 3.98, 3.7, 2.35], 2);
  body(s, T.F, 5.678, 5.044, 3.57, 0.525, { fontSize: 11 });
  label(s, 'Description', 5.676, 4.606, 0.978, 0.314, { fontSize: 14, color: C.sub, bold: true, wrap: false });
  title(s, ['MOCKUP', 'COLLECTION’S'], 7.892, 1.416, 3.725, 1.346);
  body(s, T.A, 7.892, 3.038, 4.347, 0.875, { align: 'justify' });
}

function slide22(s) {
  band(s, 0, 6.154, 13.333, 1.346, C.lilac);
  device(s, [6.667, 0.996, 7.187, 5.507], [6.902, 1.241, 6.431, 4.878]);
  title(s, ['MOCKUP', 'COLLECTION’S'], 1.423, 1.696, 3.725, 1.346);
  body(s, T.A, 1.423, 3.318, 4.347, 0.875, { align: 'justify' });
  body(s, T.F, 1.425, 5.002, 2.901, 0.802, { fontSize: 11, align: 'justify' });
  label(s, 'Description', 1.423, 4.564, 0.978, 0.314, { fontSize: 14, color: C.sub, bold: true, wrap: false });
}

function slide23(s) {
  band(s, 0, 0, 5.514, 7.5, C.lilac);
  band(s, 0, 6.2, 7.76, 0.24, '9E9E9E');
  device(s, [-0.19, 1.21, 7.16, 4.85], [0, 1.396, 6.789, 4.482]);
  title(s, ['MOCKUP', 'COLLECTION’S'], 8.173, 1.696, 3.725, 1.346);
  body(s, T.B, 8.173, 3.318, 4.05, 0.875, { align: 'justify' });
  body(s, 'PLACEHOLDER', 8.175, 5.002, 3.02, 0.525, { fontSize: 11, align: 'justify' });
  label(s, 'Description', 8.173, 4.564, 0.978, 0.314, { fontSize: 14, color: C.sub, bold: true, wrap: false });
}

function slide24(s) {
  band(s, 0, 3.388, 3.729, 4.112, C.lilac);
  photo(s, 2.774, 1.612, 3.734, 4.737);
  title(s, ['04'], 0.254, 4.488, 2.517, 2.322, { fontSize: 138, color: C.light, transparency: 38, align: 'center' });
  title(s, ['SECTION'], 0.722, 4.078, 1.581, 0.471, { fontSize: 28, color: C.light, transparency: 38, align: 'center' });
  title(s, ['DATA', 'INFOGRAPHIC'], 7.727, 2.129, 3.618, 1.346);
  body(s, T.A, 7.727, 3.751, 4.347, 0.875, { align: 'justify' });
  body(s, T.C, 7.727, 4.8, 4.253, 0.571, { bold: true, fontFace: HEAD, align: 'justify' });
}

function slide25(s) {
  s.addShape('upDownArrow', { x: 6.46, y: 2.236, w: 0.445, h: 4.423, fill: { color: C.track },
      line: { color: C.white, width: 1 } });
  s.addShape('roundRect', { x: 6.939, y: 4.71, w: 1.578, h: 1.578, fill: { color: C.mid }, rectRadius: 0.18 });
  s.addShape('roundRect', { x: 4.825, y: 4.71, w: 1.578, h: 1.578, fill: { color: C.light }, rectRadius: 0.18 });
  s.addShape('roundRect', { x: 6.939, y: 2.577, w: 1.578, h: 1.578, fill: { color: C.lilac }, rectRadius: 0.18 });
  s.addShape('roundRect', { x: 4.825, y: 2.577, w: 1.578, h: 1.578, fill: { color: C.purple }, rectRadius: 0.18 });
  s.addShape('leftRightArrow', { x: 4.472, y: 4.224, w: 4.422, h: 0.447, fill: { color: C.track },
      line: { color: C.white, width: 1 } });
  body(s, T.J, 1.23, 5.207, 3.107, 0.875, { align: 'right' });
  label(s, 'Description', 3.313, 4.821, 1.024, 0.311, { fontSize: 14, color: C.sub, bold: true, align: 'right' });
  body(s, T.J, 1.23, 3.074, 3.107, 0.875, { align: 'right' });
  label(s, 'Description', 3.313, 2.688, 1.024, 0.311, { fontSize: 14, color: C.sub, bold: true, align: 'right' });
  body(s, T.J, 8.998, 5.207, 3.107, 0.875);
  label(s, 'Description', 8.998, 4.821, 1.024, 0.311, { fontSize: 14, color: C.sub, bold: true });
  body(s, T.J, 8.998, 3.074, 3.107, 0.875);
  label(s, 'Description', 8.998, 2.688, 1.024, 0.311, { fontSize: 14, color: C.sub, bold: true });
  glyph(s, 5.262, 3.015, 0.703, 0.703, 'wedgeRoundRectCallout');
  glyph(s, 7.377, 5.245, 0.703, 0.508, 'triangle');
  glyph(s, 5.262, 5.148, 0.703, 0.703, 'flowChartManualInput');
  glyph(s, 7.377, 3.016, 0.703, 0.702, 'pie');
  title(s, ['DATA INFOGRAPHIC'], 4.058, 1.047, 5.217, 0.673);
}

function slide26(s) {
  pieChart(s, 4.792, 2.563, 3.749, 3.58);
  body(s, T.I, 1.385, 3.021, 2.663, 0.802, { fontSize: 11 });
  label(s, 'Description', 1.75, 2.628, 1.013, 0.314, { fontSize: 14, color: C.sub, bold: true, wrap: false });
  band(s, 1.383, 2.739, 0.196, 0.196, C.purple);
  body(s, T.I, 1.385, 4.822, 2.663, 0.802, { fontSize: 11 });
  label(s, 'Description', 1.75, 4.429, 1.013, 0.314, { fontSize: 14, color: C.sub, bold: true, wrap: false });
  band(s, 1.383, 4.539, 0.196, 0.196, C.mid);
  body(s, T.I, 9.286, 3.021, 2.666, 0.802, { fontSize: 11, align: 'right' });
  label(s, 'Description', 10.57, 2.628, 1.013, 0.314, { fontSize: 14, color: C.sub, bold: true, align: 'right', wrap: false });
  band(s, 11.755, 2.739, 0.196, 0.196, C.light);
  body(s, T.I, 9.286, 4.822, 2.666, 0.802, { fontSize: 11, align: 'right' });
  label(s, 'Description', 10.57, 4.429, 1.013, 0.314, { fontSize: 14, color: C.sub, bold: true, align: 'right', wrap: false });
  band(s, 11.755, 4.539, 0.196, 0.196, C.lilac);
  label(s, '12,67 Pts', 6.977, 5.114, 0.84, 0.314, { fontSize: 14, color: C.white, bold: true, wrap: false });
  glyph(s, 7.174, 4.626, 0.446, 0.445, 'flowChartManualOperation');
  title(s, ['DATA INFOGRAPHIC'], 4.058, 1.047, 5.217, 0.673);
}

function slide27(s) {
  s.addShape('donut', { x: 3.933, y: 3.969, w: 2.649, h: 2.649, fill: { color: C.purple } });
  s.addShape('donut', { x: 5.342, y: 2.562, w: 2.649, h: 2.649, fill: { color: C.mid } });
  s.addShape('donut', { x: 6.747, y: 3.969, w: 2.653, h: 2.653, fill: { color: C.purple } });
  title(s, ['01'], 6.252, 3.439, 0.758, 0.841, { fontSize: 44, color: C.mid, inset: true });
  title(s, ['02'], 4.845, 4.872, 0.882, 0.841, { fontSize: 44, color: C.purple, inset: true });
  title(s, ['03'], 7.661, 4.872, 0.882, 0.841, { fontSize: 44, color: C.purple, inset: true });
  glyph(s, 6.481, 2.76, 0.368, 0.29, 'flowChartManualOperation');
  glyph(s, 7.919, 6.122, 0.359, 0.374, 'cube');
  glyph(s, 5.081, 6.073, 0.352, 0.403, 'funnel');
  rule(s, 4.845, 3.191, 1.016, C.mid);
  rule(s, 3.723, 5.295, 0.701, C.purple);
  rule(s, 8.931, 6.115, 3.021, C.purple);
  label(s, 'Description', 3.548, 2.632, 1.18, 0.415, { fontSize: 14, bold: true, align: 'right', wrap: false, inset: true });
  label(s, T.L, 2.079, 3.044, 2.649, 0.625, { fontSize: 11, align: 'right', inset: true });
  label(s, 'Description', 2.359, 4.766, 1.18, 0.415, { fontSize: 14, bold: true, align: 'right', wrap: false, inset: true });
  label(s, T.L, 0.802, 5.146, 2.737, 0.625, { fontSize: 11, align: 'right', inset: true });
  label(s, 'Description', 9.657, 4.956, 1.18, 0.415, { fontSize: 14, bold: true, wrap: false, inset: true });
  label(s, T.L, 9.657, 5.382, 2.572, 0.625, { fontSize: 11, inset: true });
  label(s, 'Media Analysis', 8.487, 2.547, 1.554, 0.337, { fontSize: 14, color: C.purple, bold: true, lineSpacing: 1, wrap: false, inset: true });
  label(s, T.H, 8.531, 2.92, 3.188, 0.903, { fontSize: 11, inset: true });
  title(s, ['DATA INFOGRAPHIC'], 4.058, 1.047, 5.217, 0.673);
}

function slide28(s) {
  band(s, 7.761, 2.05, 4.197, 0.816, C.purple);
  band(s, 7.094, 3.054, 4.197, 0.816, C.lilac);
  band(s, 7.761, 4.058, 4.197, 0.816, C.purple);
  band(s, 7.094, 5.061, 4.197, 0.816, C.lilac);
  label(s, '12,617 pts ~ 18,817 pts', 8.36, 3.293, 2.098, 0.236, { fontSize: 14, color: C.sub, bold: true, align: 'center', lineSpacing: 1, wrap: false });
  label(s, '12,617 pts ~ 18,817 pts', 9.027, 4.297, 2.098, 0.236, { fontSize: 14, color: C.white, bold: true, align: 'center', lineSpacing: 1, wrap: false });
  label(s, '12,617 pts ~ 18,817 pts', 8.36, 5.301, 2.098, 0.236, { fontSize: 14, color: C.sub, bold: true, align: 'center', lineSpacing: 1, wrap: false });
  label(s, '12,617 pts ~ 18,817 pts', 9.027, 2.289, 2.098, 0.236, { fontSize: 14, color: C.white, bold: true, align: 'center', lineSpacing: 1, wrap: false });
  label(s, 'PLACEHOLDER', 1.398, 3.985, 4.344, 0.572, { fontSize: 12 });
  label(s, 'PLACEHOLDER', 1.398, 4.712, 4.344, 1.178, { fontSize: 12 });
  label(s, 'description', 1.348, 3.52, 1.013, 0.314, { fontSize: 14, color: C.sub, bold: true, align: 'right', wrap: false });
  title(s, ['12,6 Pts'], 1.948, 2.657, 1.904, 0.606, { fontSize: 36, color: C.sub });
  s.addShape('upArrow', { x: 1.387, y: 2.669, w: 0.367, h: 0.456, fill: { color: C.purple } });
  glyph(s, 8.413, 4.296, 0.34, 0.339, 'wedgeEllipseCallout', C.white);
  glyph(s, 7.747, 3.292, 0.339, 0.339, 'flowChartOnlineStorage', C.sub);
  glyph(s, 8.414, 2.288, 0.339, 0.339, 'flowChartManualInput', C.white);
  glyph(s, 7.747, 5.3, 0.339, 0.339, 'flowChartMultidocument', C.sub);
  title(s, ['DATA INFOGRAPHIC'], 4.058, 1.047, 5.217, 0.673);
}

function slide29(s) {
  africaMap(s);
  body(s, T.K, 1.972, 3.11, 2.535, 0.807, { fontSize: 11 });
  label(s, 'Description', 1.972, 2.758, 1.013, 0.314, { fontSize: 14, color: C.sub, bold: true, wrap: false });
  dot(s, 1.41, 4.925, C.lilac, C.sub, 'flowChartMagneticDisk');
  body(s, T.K, 5.273, 3.11, 2.535, 0.807, { fontSize: 11 });
  label(s, 'Description', 5.273, 2.758, 1.013, 0.314, { fontSize: 14, color: C.sub, bold: true, wrap: false });
  dot(s, 4.711, 2.772, C.lilac, C.sub, 'flowChartDocument');
  body(s, T.K, 1.972, 5.263, 2.535, 0.807, { fontSize: 11 });
  label(s, 'Description', 1.972, 4.911, 1.013, 0.314, { fontSize: 14, color: C.sub, bold: true, wrap: false });
  dot(s, 1.41, 2.772, C.purple, C.white, 'flowChartConnector');
  body(s, T.K, 5.273, 5.263, 2.535, 0.807, { fontSize: 11 });
  label(s, 'Description', 5.273, 4.911, 1.013, 0.314, { fontSize: 14, color: C.sub, bold: true, wrap: false });
  dot(s, 4.711, 4.925, C.purple, C.white, 'flowChartManualOperation');
  title(s, ['DATA INFOGRAPHIC'], 4.058, 1.047, 5.217, 0.673);
}

function slide30(s) {
  photo(s, 0.583, 1.151, 12.167, 5.198);
  title(s, ['THANK YOU'], 2.23, 2.889, 8.874, 1.935, { fontSize: 115, color: C.lilac, transparency: 50, align: 'center' });
  band(s, 0, 4.237, 7.212, 3.283, C.lilac, { transparency: 20 });
  band(s, 10.171, -0.02, 1.9, 2.668, C.lilac, { transparency: 20 });
  title(s, ['THANK YOU'], 1.355, 4.403, 5.249, 1.144, { fontSize: 67.99, color: C.deep, align: 'center' });
  label(s, 'SEE YOU TO THE NEXTPROJECT', 1.355, 5.761, 4.924, 0.236, { fontSize: 14, color: C.deep, charSpacing: 6, fontFace: HEAD, align: 'center', lineSpacing: 1, wrap: false });
}

// --------------------------------------------------------------------------
const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
];

const pres = new PptxGenJS();
pres.defineLayout({ name: 'DECK', width: 13.3333333, height: 7.5 });
pres.layout = 'DECK';
pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };
pres.author = 'LARSO';
pres.title = 'LARSO Lookbook Concept';

BUILDERS.forEach(build => {
  const s = pres.addSlide();
  s.background = { color: 'FFFFFF' };
  build(s);
  chrome(s);
});

pres.writeFile({ fileName: path.join(__dirname, '0716d450-4ad8-4698-a7ac-23cc0bda03f8_grok_final.pptx') })
  .then(f => console.log('wrote ' + f));
