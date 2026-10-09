/**
 * "Fasting Blessings" — Ramadan presentation rebuilt with pptxgenjs.
 * Slide size 20 x 11.25 in (widescreen, 16:9). 23 slides.
 * Photographs in the original deck are redrawn as flat grey [image] placeholders.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const TEAL = '007970';
const TEAL_DARK = '03423D';
const TEAL_MID = '045F58';
const INK = '2E2E2E';
const CREAM = 'F4F4F4';
const WHITE = 'FFFFFF';
const GOLD = 'DE932E';
const GOLD_LT = 'D6B07C';
const AMBER = 'FC9F0B';
const PHOTO_BG = 'E9E9E9';
const PHOTO_INK = 'BFBFBF';
const GREY_CARD = 'EDECEC';

const F_TITLE = 'El Messiri Bold';
const F_BOLD = 'Livvic Bold';
const F_BODY = 'Livvic';

/* --------------------------------------------------------------- text styles */

const S = {
  h1: { fontFace: F_TITLE, fontSize: 48, bold: true, color: TEAL, lineSpacing: 67.19 },
  h2: { fontFace: F_BOLD, fontSize: 24, bold: true, color: INK, lineSpacing: 33.59 },
  h3: { fontFace: F_BOLD, fontSize: 21, bold: true, color: INK, lineSpacing: 29.4 },
  h4: { fontFace: F_BOLD, fontSize: 18, bold: true, color: INK, lineSpacing: 25.2 },
  body: { fontFace: F_BODY, fontSize: 15.99, color: INK, lineSpacing: 28.79 },
  note: { fontFace: F_BODY, fontSize: 15.99, color: INK, lineSpacing: 26.87 },
  card: { fontFace: F_BODY, fontSize: 16, color: INK, lineSpacing: 26.88, align: 'center' },
  cap: { fontFace: F_BOLD, fontSize: 15.99, bold: true, color: INK, lineSpacing: 22.39, align: 'center' },
  tiny: { fontFace: F_BODY, fontSize: 15.99, color: INK, lineSpacing: 27.03 },
  num: { fontFace: F_BOLD, fontSize: 19.99, bold: true, color: CREAM, lineSpacing: 27.99, align: 'center' },
  foot: { fontFace: F_BOLD, fontSize: 14, bold: true, color: INK, lineSpacing: 19.6 },
};

/* ------------------------------------------------------------- filler copy */

const A = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit,sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.';
const A2 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.';
const B = 'Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.';
const C = 'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.';
const D = 'Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.';
const E = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit';

/* ------------------------------------------------------------------ helpers */

/** Text box; the deck uses zero insets and top alignment everywhere. */
function txt(s, style, text, x, y, w, h, extra) {
  s.addText(text, Object.assign({ x, y, w, h, margin: 0, valign: 'top', align: 'left' }, style, extra));
}

/** Flat rectangle. */
function box(s, x, y, w, h, color, extra) {
  s.addShape('rect', Object.assign({ x, y, w, h, fill: { color } }, extra));
}

/** Colour of the deck's teal gradient at fraction `t` (0 = top, 1 = bottom). */
function tealAt(t) {
  const from = [0x00, 0x79, 0x70];
  const to = [0x03, 0x42, 0x3d];
  return from
    .map((c, k) => Math.round(c + (to[k] - c) * t).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
}

/** Vertical teal gradient block, painted as a stack of interpolated bands. */
function gradBox(s, x, y, w, h, bands) {
  const n = bands || 16;
  for (let i = 0; i < n; i++) {
    box(s, x, y + (h * i) / n, w, h / n + 0.02, tealAt(i / (n - 1)));
  }
}

/** Grey stand-in for a photograph. */
function photo(s, x, y, w, h, shape) {
  s.addShape(shape || 'rect', { x, y, w, h, fill: { color: PHOTO_BG } });
  txt(s, { fontFace: F_BODY, fontSize: 11, color: PHOTO_INK, align: 'center' }, '[image]',
    x, y + h / 2 - 0.16, w, 0.32, { valign: 'middle' });
}

/** Left half of the gold name-plate outline: pointed tip, notch, flat top. */
const PLATE_EDGE = [
  [0.000, 0.479], [0.025, 0.301], [0.050, 0.164], [0.075, 0.123], [0.100, 0.068],
  [0.125, 0.000], [0.150, 0.068], [0.500, 0.068],
];

/** Ornate gold name-plate used for captions ("Read More", team names, ...). */
function plate(s, x, y, w, h, label) {
  const pts = [];
  PLATE_EDGE.forEach(function (p) { pts.push({ x: p[0] * w, y: p[1] * h }); });
  PLATE_EDGE.slice().reverse().forEach(function (p) { pts.push({ x: (1 - p[0]) * w, y: p[1] * h }); });
  PLATE_EDGE.forEach(function (p) { pts.push({ x: (1 - p[0]) * w, y: (1 - p[1]) * h }); });
  PLATE_EDGE.slice().reverse().forEach(function (p) { pts.push({ x: p[0] * w, y: (1 - p[1]) * h }); });
  pts.push({ close: true });
  s.addShape('custGeom', { x, y, w, h, points: pts, fill: { color: WHITE }, line: { color: GOLD, width: 1.25 } });
  txt(s, S.cap, label, x + 0.15, y + h / 2 - 0.145, w - 0.3, 0.29, { valign: 'middle' });
}

/**
 * Radius of the deck's scalloped lozenge, sampled every 3.75 deg over one
 * quadrant and mirrored around both axes.
 */
const LOZENGE_R = [
  0.999, 0.966, 0.936, 0.910, 0.885, 0.860, 0.831, 0.817,
  0.846, 0.874, 0.859, 0.837, 0.818, 0.837, 0.859, 0.874,
  0.846, 0.817, 0.831, 0.860, 0.885, 0.910, 0.936, 0.966,
];

/** Teal lozenge holding a step number. */
function badge(s, x, y, w, h, label) {
  const q = LOZENGE_R.length;
  const pts = [];
  for (let i = 0; i < q * 4; i++) {
    const t = (2 * Math.PI * i) / (q * 4);
    const r = LOZENGE_R[i % q] / 2;
    pts.push({ x: (0.5 + r * Math.cos(t)) * w, y: (0.5 + r * Math.sin(t)) * h });
  }
  pts.push({ close: true });
  s.addShape('custGeom', { x, y, w, h, points: pts, fill: { color: TEAL }, line: { color: GOLD_LT, width: 2.5 } });
  txt(s, S.num, label, x - 0.12, y + h / 2 - 0.19, w + 0.24, 0.38, { valign: 'middle' });
}

/* ------------------------------------------------- decorative page furniture */

/** Hanging lantern: cord, gold caps, amber body with lit panels. */
function lamp(s, x, y, w, h) {
  box(s, x + w * 0.47, y, w * 0.06, h * 0.66, GOLD);
  const by = y + h * 0.62;
  const bh = h * 0.38;
  s.addShape('trapezoid', { x: x + w * 0.22, y: by, w: w * 0.56, h: bh * 0.12, fill: { color: GOLD } });
  s.addShape('hexagon', { x, y: by + bh * 0.1, w, h: bh * 0.62, fill: { color: AMBER }, line: { color: GOLD, width: 0.75 } });
  [0.22, 0.44, 0.66].forEach(function (u) { box(s, x + w * u, y + h * 0.72, w * 0.14, bh * 0.34, WHITE); });
  s.addShape('trapezoid', { x: x + w * 0.22, y: by + bh * 0.7, w: w * 0.56, h: bh * 0.13, fill: { color: GOLD }, flipV: true });
  s.addShape('diamond', { x: x + w * 0.36, y: by + bh * 0.84, w: w * 0.28, h: bh * 0.16, fill: { color: GOLD } });
}

/** Hanging crescent moon. */
function moon(s, x, y, w, h, flip) {
  box(s, x + w * 0.37, y, w * 0.05, h * 0.79, GOLD);
  s.addShape('moon', { x: x + w * 0.02, y: y + h * 0.76, w: w * 0.88, h: h * 0.24, fill: { color: AMBER }, rotate: 200, flipH: !!flip });
}

/** Four-pointed sparkle. */
function spark(s, x, y, w, h) {
  s.addShape('star4', { x, y, w, h, fill: { color: AMBER } });
}

const DECO_FN = { lamp: lamp, moon: moon, spark: spark };

/**
 * Lanterns, moons and sparkles floating around the edge of a slide:
 * [kind, x, y, w, h, flipH].
 */
const DECO = {
  2: [['spark', 17.444, 1.73, 0.437, 0.474], ['lamp', 17.68, -1.404, 0.402, 2.807], ['lamp', 18.322, -1.017, 0.764, 4.283], ['spark', 19.238, 0.628, 0.358, 0.388], ['spark', 0.946, 0.823, 0.358, 0.388], ['lamp', 0.304, -1.017, 0.508, 3.545], ['spark', 18.666, 3.598, 0.257, 0.279]],
  3: [['lamp', 17.68, -1.404, 0.402, 2.807], ['lamp', 18.322, -1.017, 0.764, 4.283], ['spark', 19.238, 0.628, 0.358, 0.388], ['moon', 0.312, -1.928, 0.813, 3.794], ['spark', 1.199, 0.823, 0.358, 0.388], ['spark', 17.68, 1.793, 0.437, 0.474]],
  4: [['lamp', 11.678, -0.857, 0.611, 3.427], ['moon', 18.875, -2.161, 0.813, 3.794, 1], ['lamp', 0.396, 0, 0.402, 2.807], ['spark', 18.438, 0.382, 0.437, 0.474], ['spark', 11.129, 1.699, 0.261, 0.283], ['spark', 12.46, 0.857, 0.261, 0.283], ['spark', 0.135, 2.807, 0.261, 0.283]],
  5: [['lamp', 18.628, -0.842, 0.611, 3.427], ['spark', 18.079, 1.714, 0.261, 0.283], ['spark', 19.41, 0.872, 0.261, 0.283], ['moon', 8.195, -1.902, 0.813, 3.794, 1], ['spark', 7.758, 0.64, 0.437, 0.474], ['spark', 9.176, 1.57, 0.306, 0.333], ['lamp', 0.396, 0, 0.402, 2.807], ['spark', 0.135, 2.807, 0.261, 0.283]],
  6: [['spark', 17.444, 1.73, 0.437, 0.474], ['lamp', 17.68, -1.404, 0.402, 2.807], ['lamp', 18.322, -1.017, 0.764, 4.283], ['spark', 19.238, 0.628, 0.358, 0.388], ['spark', 18.666, 3.598, 0.257, 0.279], ['moon', 2.624, -1.902, 0.813, 3.794], ['spark', 2.187, 0.64, 0.437, 0.474, 1], ['spark', 3.605, 1.57, 0.306, 0.333, 1]],
  7: [['lamp', 19.136, -1.108, 0.402, 2.807], ['spark', 19.538, 1.699, 0.261, 0.283], ['moon', 1.263, -1.902, 0.813, 3.794], ['spark', 0.826, 0.64, 0.437, 0.474, 1], ['spark', 2.244, 1.57, 0.306, 0.333, 1], ['lamp', 11.678, -0.857, 0.611, 3.427], ['spark', 11.129, 1.699, 0.261, 0.283], ['spark', 12.46, 0.857, 0.261, 0.283]],
  8: [['spark', 11.7, 1.567, 0.391, 0.425], ['lamp', 11.911, -1.24, 0.36, 2.514], ['lamp', 12.486, -0.893, 0.684, 3.836], ['spark', 13.306, 0.58, 0.32, 0.348], ['spark', 12.794, 3.24, 0.23, 0.249], ['moon', 0.194, -1.841, 0.695, 3.246], ['spark', 0.952, 0.513, 0.306, 0.332]],
  9: [['spark', 4.377, 6.285, 0.437, 0.474], ['spark', 0.472, 5.761, 0.257, 0.279]],
  10: [['moon', 1.263, -1.902, 0.813, 3.794], ['spark', 0.826, 0.64, 0.437, 0.474, 1], ['spark', 2.244, 1.57, 0.306, 0.333, 1], ['lamp', 18.414, -1.252, 0.402, 2.807], ['lamp', 19.056, -0.864, 0.764, 4.283], ['spark', 17.737, 1.147, 0.437, 0.474, 1], ['spark', 19.576, 0.474, 0.306, 0.333, 1]],
  11: [['spark', 13.164, 1.472, 0.37, 0.402], ['lamp', 13.363, -1.181, 0.341, 2.376], ['lamp', 13.907, -0.853, 0.647, 3.625], ['spark', 14.682, 0.539, 0.303, 0.329], ['spark', 14.198, 3.054, 0.217, 0.236], ['moon', 1.817, -1.626, 0.662, 3.089], ['spark', 1.461, 0.445, 0.356, 0.386, 1], ['spark', 2.616, 1.201, 0.25, 0.271, 1]],
  12: [['lamp', 10.055, -1.284, 0.611, 3.427], ['moon', 18.952, -2.019, 0.735, 3.433, 1], ['lamp', 0.556, -0.53, 0.402, 2.807], ['spark', 18.557, 0.282, 0.395, 0.429], ['spark', 9.506, 1.273, 0.261, 0.283], ['spark', 10.837, 0.43, 0.261, 0.283], ['spark', 0.295, 2.277, 0.261, 0.283]],
  13: [['spark', 11.7, 1.567, 0.391, 0.425], ['lamp', 11.911, -1.24, 0.36, 2.514], ['lamp', 12.486, -0.893, 0.684, 3.836], ['spark', 13.306, 0.58, 0.32, 0.348], ['spark', 12.794, 3.24, 0.23, 0.249], ['moon', 0.194, -1.841, 0.695, 3.246], ['spark', 0.952, 0.513, 0.306, 0.332]],
  14: [['lamp', 8.991, -1.331, 0.611, 3.427], ['moon', 18.875, -2.161, 0.813, 3.794, 1], ['spark', 18.438, 0.382, 0.437, 0.474], ['spark', 8.442, 1.225, 0.261, 0.283], ['spark', 9.772, 0.382, 0.261, 0.283]],
  15: [['moon', 12.098, -1.523, 0.735, 3.433, 1], ['spark', 11.702, 0.777, 0.395, 0.429], ['spark', 12.986, 1.618, 0.277, 0.301], ['lamp', 0.526, -1.171, 0.402, 2.807], ['spark', 0.265, 1.636, 0.261, 0.283]],
  16: [['spark', 0.601, 1.48, 0.314, 0.341], ['lamp', 0.77, -0.774, 0.289, 2.02], ['lamp', 1.232, -0.496, 0.549, 3.081], ['spark', 1.891, 0.687, 0.257, 0.279], ['spark', 1.48, 2.824, 0.185, 0.2], ['moon', 10.12, -1.662, 0.709, 3.312, 1], ['spark', 9.738, 0.558, 0.382, 0.414]],
  17: [['lamp', 18.657, -1.134, 0.539, 3.022], ['spark', 18.173, 1.12, 0.23, 0.25], ['spark', 19.347, 0.377, 0.23, 0.25], ['moon', 8.328, -1.902, 0.728, 3.401, 1], ['spark', 7.936, 0.377, 0.392, 0.425], ['spark', 9.208, 1.21, 0.275, 0.298], ['lamp', 0.723, -0.743, 0.402, 2.807], ['spark', 0.462, 2.064, 0.261, 0.283]],
  18: [['moon', 0.981, -1.683, 0.722, 3.371], ['spark', 0.593, 0.576, 0.388, 0.421, 1], ['spark', 1.853, 1.402, 0.272, 0.296, 1], ['lamp', 10.236, -0.754, 0.543, 3.045], ['spark', 9.748, 1.517, 0.232, 0.252], ['spark', 10.93, 0.768, 0.232, 0.252]],
  19: [['lamp', 18.657, -1.134, 0.539, 3.022], ['spark', 18.173, 1.12, 0.23, 0.25], ['spark', 19.347, 0.377, 0.23, 0.25], ['moon', 8.328, -1.902, 0.728, 3.401, 1], ['spark', 7.936, 0.377, 0.392, 0.425], ['spark', 9.208, 1.21, 0.275, 0.298], ['lamp', 0.723, -0.743, 0.402, 2.807], ['spark', 0.462, 2.064, 0.261, 0.283]],
  20: [['lamp', 11.013, -1.192, 0.611, 3.427], ['moon', 19.009, -1.897, 0.707, 3.3, 1], ['lamp', 0.51, -0.42, 0.402, 2.807], ['spark', 18.629, 0.315, 0.38, 0.413], ['spark', 10.464, 1.364, 0.261, 0.283], ['spark', 11.795, 0.521, 0.261, 0.283], ['spark', 0.249, 2.387, 0.261, 0.283]],
  21: [['lamp', 18.657, -1.134, 0.539, 3.022], ['spark', 18.173, 1.12, 0.23, 0.25], ['spark', 19.347, 0.377, 0.23, 0.25], ['moon', 8.328, -1.902, 0.728, 3.401, 1], ['spark', 7.936, 0.377, 0.392, 0.425], ['spark', 9.208, 1.21, 0.275, 0.298], ['lamp', 0.723, -0.743, 0.402, 2.807], ['spark', 0.462, 2.064, 0.261, 0.283]],
  22: [['moon', 18.875, -2.161, 0.813, 3.794, 1], ['lamp', 0.396, 0, 0.402, 2.807], ['spark', 18.438, 0.382, 0.437, 0.474], ['spark', 0.135, 2.807, 0.261, 0.283], ['lamp', 9.753, -1.094, 0.611, 3.427], ['spark', 9.204, 1.462, 0.261, 0.283], ['spark', 10.535, 0.619, 0.261, 0.283]],
};

function deco(s, n) {
  (DECO[n] || []).forEach(function (d) { DECO_FN[d[0]](s, d[1], d[2], d[3], d[4], d[5]); });
}

/** Ketupat glyph + "Page / NN" strip repeated at the foot of slides 2-22. */
function footer(s, n) {
  const label = 'Page / ' + String(n).padStart(2, '0');
  ICON.ketupat(s, 1.125, 10.481, 0.256, 0.419, TEAL);
  txt(s, S.foot, label, 1.509, 10.553, 1.302, 0.253);
}

/* ------------------------------------------------------- large ornament art */

/** Half of the title-slide cartouche outline, sampled left edge -> top centre. */
const CARTOUCHE = [
  [0.000, 0.496], [0.016, 0.486], [0.031, 0.473], [0.047, 0.449], [0.062, 0.404], [0.078, 0.337],
  [0.094, 0.277], [0.109, 0.242], [0.125, 0.227], [0.141, 0.219], [0.156, 0.219], [0.172, 0.220],
  [0.188, 0.220], [0.203, 0.210], [0.219, 0.193], [0.234, 0.176], [0.250, 0.173], [0.266, 0.173],
  [0.281, 0.177], [0.297, 0.177], [0.312, 0.153], [0.328, 0.140], [0.344, 0.130], [0.359, 0.123],
  [0.375, 0.117], [0.391, 0.111], [0.406, 0.106], [0.422, 0.100], [0.438, 0.093], [0.453, 0.085],
  [0.469, 0.074], [0.484, 0.056], [0.500, 0.008],
];

/** Scalloped ornamental frame: teal body, dark rim and an inset gold keyline. */
function cartouche(s, x, y, w, h) {
  const ring = function (inset, opts) {
    const at = function (u, v) {
      return { x: (0.5 + (u - 0.5) * inset) * w, y: (0.5 + (v - 0.5) * inset) * h };
    };
    const pts = [];
    CARTOUCHE.forEach(function (p) { pts.push(at(p[0], p[1])); });
    CARTOUCHE.slice().reverse().forEach(function (p) { pts.push(at(1 - p[0], p[1])); });
    CARTOUCHE.forEach(function (p) { pts.push(at(1 - p[0], 1 - p[1])); });
    CARTOUCHE.slice().reverse().forEach(function (p) { pts.push(at(p[0], 1 - p[1])); });
    pts.push({ close: true });
    s.addShape('custGeom', Object.assign({ x, y, w, h, points: pts }, opts));
  };
  ring(1.0, { fill: { color: TEAL }, line: { color: TEAL_MID, width: 5 } });
  ring(0.955, { line: { color: GOLD_LT, width: 3.5 } });
  ring(0.925, { fill: { color: TEAL } });
}

/** Tiled corner ornament: teal panel cut by a scalloped diagonal edge. */
const CORNER_EDGE = [
  [0.000, 0.998], [0.018, 0.968], [0.036, 0.944], [0.054, 0.925], [0.071, 0.909], [0.089, 0.895],
  [0.107, 0.883], [0.125, 0.871], [0.143, 0.860], [0.161, 0.849], [0.179, 0.717], [0.196, 0.685],
  [0.214, 0.662], [0.232, 0.644], [0.250, 0.629], [0.268, 0.616], [0.286, 0.606], [0.304, 0.596],
  [0.321, 0.587], [0.339, 0.580], [0.357, 0.429], [0.375, 0.388], [0.393, 0.322], [0.411, 0.321],
  [0.429, 0.318], [0.446, 0.314], [0.464, 0.308], [0.482, 0.302], [0.500, 0.296], [0.518, 0.293],
  [0.536, 0.292], [0.554, 0.292], [0.571, 0.294], [0.589, 0.299], [0.607, 0.305], [0.625, 0.305],
  [0.643, 0.267], [0.661, 0.237], [0.679, 0.212], [0.696, 0.191], [0.714, 0.175], [0.732, 0.162],
  [0.750, 0.151], [0.768, 0.144], [0.786, 0.139], [0.804, 0.136], [0.821, 0.136], [0.839, 0.136],
  [0.857, 0.138], [0.875, 0.138], [0.893, 0.122], [0.911, 0.097], [0.929, 0.070], [0.946, 0.048],
  [0.964, 0.029], [0.982, 0.012], [1.000, 0.000],
];

function corner(s, x, y, w, h, flip) {
  const at = function (u, v) {
    return flip ? { x: (1 - u) * w, y: (1 - v) * h } : { x: u * w, y: v * h };
  };
  const pts = [at(0, 0), at(1, 0)];
  CORNER_EDGE.slice().reverse().forEach(function (p) { pts.push(at(p[0], p[1])); });
  pts.push({ close: true });
  s.addShape('custGeom', { x, y, w, h, points: pts, fill: { color: TEAL }, line: { color: GOLD_LT, width: 3 } });

  // Girih tiling inside the panel: a white lattice studded with gold stars.
  const edgeAt = function (u) {
    for (let i = 1; i < CORNER_EDGE.length; i++) {
      if (CORNER_EDGE[i][0] >= u) {
        const a = CORNER_EDGE[i - 1];
        const b = CORNER_EDGE[i];
        return a[1] + ((b[1] - a[1]) * (u - a[0])) / (b[0] - a[0]);
      }
    }
    return 0;
  };
  const cols = 8;
  const rows = 8;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const u = (c + 0.5) / cols;
      const v = (r + 0.5) / rows;
      if (v > edgeAt(u) - 0.06) continue;
      const p = at(u, v);
      const d = Math.min(w / cols, h / rows);
      s.addShape('star8', { x: x + p.x - d * 0.5, y: y + p.y - d * 0.5, w: d, h: d, line: { color: WHITE, width: 1.5 } });
      s.addShape('star8', { x: x + p.x - d * 0.16, y: y + p.y - d * 0.16, w: d * 0.32, h: d * 0.32, fill: { color: '9A6B25' } });
    }
  }
}

/** Fine gold rosette peeking in from a slide corner. */
function mandala(s, cx, cy, r) {
  const ring = function (shape, k, color) {
    s.addShape(shape, { x: cx - r * k, y: cy - r * k, w: r * k * 2, h: r * k * 2, line: { color, width: 0.75 } });
  };
  ring('star24', 1.0, GOLD);
  ring('star16', 0.86, GOLD_LT);
  ring('ellipse', 0.7, GOLD);
  ring('star16', 0.62, GOLD_LT);
  ring('star8', 0.42, GOLD);
  ring('ellipse', 0.26, GOLD_LT);
  s.addShape('star8', { x: cx - r * 0.13, y: cy - r * 0.13, w: r * 0.26, h: r * 0.26, fill: { color: GOLD } });
  for (let i = 0; i < 8; i++) {
    const a = (i * Math.PI) / 4 + Math.PI / 8;
    s.addShape('ellipse', {
      x: cx + Math.cos(a) * r * 0.52 - r * 0.09, y: cy + Math.sin(a) * r * 0.52 - r * 0.09,
      w: r * 0.18, h: r * 0.18, line: { color: GOLD, width: 0.75 },
    });
  }
}

/** Full-bleed panel whose left edge bows out into a half circle. */
function arcPanel(s, x, y, w, h, color) {
  const cu = function (x1, y1, x2, y2, ex, ey) {
    return { x: ex * w, y: ey * h, curve: { type: 'cubic', x1: x1 * w, y1: y1 * h, x2: x2 * w, y2: y2 * h } };
  };
  s.addShape('custGeom', {
    x, y, w, h, fill: { color },
    points: [
      { x: 0.35 * w, y: 0 },
      cu(0.15, 0.05, 0.0, 0.24, 0.0, 0.5),
      cu(0.0, 0.76, 0.15, 0.95, 0.35, 1.0),
      { x: w, y: h }, { x: w, y: 0 }, { close: true },
    ],
  });
}

/* ------------------------------------------------------------- small icons */

const ICON = {
  /** Domed mosque with a crescent finial, flanked by two minarets. */
  mosque: function (s, x, y, w, h, c) {
    s.addShape('moon', { x: x + w * 0.44, y: y, w: w * 0.12, h: h * 0.13, fill: { color: c }, rotate: 200 });
    s.addShape('chord', { x: x + w * 0.24, y: y + h * 0.06, w: w * 0.52, h: h * 0.62, fill: { color: c }, angleRange: [180, 0] });
    box(s, x + w * 0.24, y + h * 0.36, w * 0.52, h * 0.46, c);
    s.addShape('roundRect', { x: x + w * 0.06, y: y + h * 0.44, w: w * 0.14, h: h * 0.38, fill: { color: c }, rectRadius: 0.04 });
    s.addShape('roundRect', { x: x + w * 0.8, y: y + h * 0.44, w: w * 0.14, h: h * 0.38, fill: { color: c }, rectRadius: 0.04 });
    box(s, x, y + h * 0.82, w, h * 0.18, c);
    s.addShape('roundRect', { x: x + w * 0.42, y: y + h * 0.58, w: w * 0.16, h: h * 0.24, fill: { color: WHITE }, rectRadius: 0.05 });
  },
  /** Open Quran: two page blocks either side of a gap; `bg` is the page interior. */
  book: function (s, x, y, w, h, c, bg) {
    box(s, x, y, w * 0.46, h, c);
    box(s, x + w * 0.54, y, w * 0.46, h, c);
    box(s, x + w * 0.07, y + h * 0.16, w * 0.32, h * 0.68, bg || WHITE);
    box(s, x + w * 0.61, y + h * 0.16, w * 0.32, h * 0.68, bg || WHITE);
  },
  /** Community / "halal bil halal" heart with three figures. */
  heart: function (s, x, y, w, h, c) {
    s.addShape('heart', { x, y, w, h, fill: { color: c } });
    s.addShape('ellipse', { x: x + w * 0.36, y: y + h * 0.2, w: w * 0.28, h: h * 0.24, fill: { color: WHITE } });
    s.addShape('ellipse', { x: x + w * 0.17, y: y + h * 0.3, w: w * 0.24, h: h * 0.2, fill: { color: WHITE } });
    s.addShape('ellipse', { x: x + w * 0.59, y: y + h * 0.3, w: w * 0.24, h: h * 0.2, fill: { color: WHITE } });
    s.addShape('ellipse', { x: x + w * 0.3, y: y + h * 0.44, w: w * 0.4, h: h * 0.34, fill: { color: WHITE } });
  },
  /** Wire globe. */
  globe: function (s, x, y, w, h, c) {
    s.addShape('ellipse', { x, y, w, h, fill: { color: c } });
    box(s, x, y + h * 0.26, w, h * 0.09, WHITE);
    box(s, x, y + h * 0.62, w, h * 0.09, WHITE);
    box(s, x + w * 0.46, y, w * 0.08, h, WHITE);
    s.addShape('ellipse', { x: x + w * 0.3, y, w: w * 0.4, h, line: { color: WHITE, width: 2 } });
  },
  /** Ketupat glyph (woven rice parcel with two tails). */
  ketupat: function (s, x, y, w, h, c) {
    s.addShape('diamond', { x, y, w, h: h * 0.68, fill: { color: c } });
    box(s, x + w * 0.34, y + h * 0.62, w * 0.1, h * 0.38, c);
    box(s, x + w * 0.56, y + h * 0.62, w * 0.1, h * 0.38, c);
  },
  /** Map pin. */
  pin: function (s, x, y, w, h, c) {
    s.addShape('ellipse', { x, y, w, h: h * 0.78, fill: { color: c } });
    s.addShape('triangle', { x: x + w * 0.18, y: y + h * 0.42, w: w * 0.64, h: h * 0.58, fill: { color: c }, flipV: true });
    s.addShape('ellipse', { x: x + w * 0.3, y: y + h * 0.2, w: w * 0.4, h: h * 0.32, fill: { color: WHITE } });
  },
  /** Handset, drawn as a thick open arc. */
  phone: function (s, x, y, w, h, c) {
    s.addShape('blockArc', { x, y, w, h, angleRange: [130, 50], arcThicknessRatio: 0.55, fill: { color: c }, rotate: 315 });
  },
  /** Envelope. */
  mail: function (s, x, y, w, h, c) {
    box(s, x, y, w, h, c);
    box(s, x + w * 0.1, y + h * 0.16, w * 0.8, h * 0.68, WHITE);
    s.addShape('triangle', { x: x + w * 0.1, y: y + h * 0.16, w: w * 0.8, h: h * 0.56, fill: { color: WHITE }, line: { color: c, width: 2 }, flipV: true });
  },
  /** Circled tick used on the pricing table. */
  tick: function (s, x, y, w, h, disc, mark) {
    s.addShape('ellipse', { x, y, w, h, fill: { color: disc } });
    txt(s, { fontFace: F_BODY, fontSize: 15, color: mark, align: 'center' }, '\u2713', x, y, w, h, { valign: 'middle' });
  },
};

/* ------------------------------------------------------ composite fragments */

/** Numbered lozenge + heading + paragraph, the deck's workhorse list row. */
function numberedRow(s, n, bx, by, bw, tx, ty, tw, heading, bodyY, bodyW, bodyH, bodyText) {
  badge(s, bx, by, bw, bw, n);
  txt(s, S.h3, heading, tx, ty, tw, 0.39);
  txt(s, S.note, bodyText, tx, bodyY, bodyW, bodyH);
}

/** White service card: picture on top, bold title, small caption. */
function whiteCard(s, x, y, w, h, px, py, pw, ph, title, caption) {
  box(s, x, y, w, h, WHITE, { shadow: { type: 'outer', color: '000000', blur: 10, offset: 2, angle: 90, opacity: 0.12 } });
  photo(s, px, py, pw, ph);
  txt(s, Object.assign({}, S.h4, { align: 'center' }), title, x + 0.357, y + 2.609, 3.25, 0.325);
  txt(s, S.card, caption, x + 0.357, y + 3.038, 3.25, 0.685);
}

/** Phone mock-up: dark body with a screen cut-out. */
function phoneMock(s, x, y, w, h, sx, sy, sw, sh) {
  s.addShape('roundRect', { x, y, w, h, fill: { color: '555555' }, rectRadius: w * 0.12 });
  s.addShape('roundRect', { x: x + w * 0.02, y: y + h * 0.005, w: w * 0.96, h: h * 0.99, fill: { color: '000000' }, rectRadius: w * 0.11 });
  box(s, x + w * 0.43, y + h * 0.038, w * 0.13, h * 0.008, '555555');
  box(s, x - w * 0.012, y + h * 0.132, w * 0.012, h * 0.041, INK);
  box(s, x - w * 0.012, y + h * 0.204, w * 0.012, h * 0.074, INK);
  box(s, x - w * 0.012, y + h * 0.294, w * 0.012, h * 0.075, INK);
  box(s, x + w, y + h * 0.229, w * 0.012, h * 0.119, INK);
  photo(s, sx, sy, sw, sh);
}

/* ------------------------------------------------------------ slide builders */

const SLIDES = [];

// 1 — cover
SLIDES.push(function (s) {
  cartouche(s, 1.597, 1.285, 16.612, 8.68);
  corner(s, 0, 0, 5.263, 4.987);
  corner(s, 14.737, 6.228, 5.263, 4.987, true);
  mandala(s, 0, 9.16, 1.786);
  mandala(s, 20, 2.061, 1.786);
  txt(s, { fontFace: F_TITLE, fontSize: 104, bold: true, color: CREAM, lineSpacing: 117.52, align: 'center' },
    'FASTING BLESSINGS', 3.785, 3.791, 12.236, 3.278);
  txt(s, { fontFace: F_BODY, fontSize: 14, color: CREAM, lineSpacing: 21, align: 'center' },
    '"Ramadan is not just a month of fasting; it is a month of reflection, a month of giving, and a month to strengthen your bond with Allah."',
    5.263, 7.098, 9.265, 0.567);
});

// 2 — Introducing Ramadan
SLIDES.push(function (s) {
  gradBox(s, 16.484, 6.171, 3.516, 5.079);
  deco(s, 2);
  ICON.mosque(s, 1.116, 8.002, 1.362, 1.112, GOLD);
  txt(s, S.h1, 'Introducing Ramadan', 1.125, 2.1, 8.875, 0.898);
  txt(s, S.h3, 'Discover The Spiritual Meaning Of Ramadan', 1.125, 3.236, 7.988, 0.39);
  txt(s, S.body, A + ' ' + B + ' ' + C + ' ', 1.125, 4.062, 8.342, 1.94, { align: 'justify' });
  txt(s, S.body, A + ' ' + B + ' ', 1.125, 6.264, 8.342, 1.148, { align: 'justify' });
  txt(s, S.h3, 'Ramadan Kareem', 2.81, 7.902, 6.899, 0.39);
  txt(s, S.body, A, 2.81, 8.373, 6.657, 0.752, { align: 'justify' });
  footer(s, 2);
  photo(s, 11.342, 0.823, 3.269, 4.755);
  photo(s, 11.333, 5.792, 3.269, 4.247);
  photo(s, 14.85, 3.598, 3.269, 5.519);
});

// 3 — Origin Of Ramadan
SLIDES.push(function (s) {
  photo(s, 12.479, 1.793, 4.769, 8.332);
  deco(s, 3);
  gradBox(s, 15.885, 7.006, 4.115, 1.852);
  txt(s, S.h1, 'Origin Of Ramadan', 1.125, 2.008, 9.505, 0.898);
  badge(s, 1.125, 3.499, 0.862, 0.86, '01');
  txt(s, S.h2, 'Ramadan Kareem', 2.314, 3.447, 6.899, 0.444);
  txt(s, S.body, A + ' ' + B + ' ' + C + ' ' + D, 2.314, 4.043, 9.171, 1.94, { align: 'justify' });
  badge(s, 1.125, 6.576, 0.862, 0.86, '02');
  txt(s, S.h2, 'Nuzulul Quran', 2.314, 6.524, 6.899, 0.444);
  txt(s, S.body, A + ' ' + B + ' ' + C + ' ' + D, 2.314, 7.185, 9.171, 1.94, { align: 'justify' });
  txt(s, Object.assign({}, S.note, { color: CREAM, align: 'center' }), A, 16.057, 7.183, 3.771, 1.425);
  footer(s, 3);
});

// 4 — The Virtues Of Fasting
SLIDES.push(function (s) {
  photo(s, 13.875, 2.174, 5.0, 9.076);
  deco(s, 4);
  txt(s, S.h1, 'The Virtues Of Fasting', 1.125, 2.069, 9.505, 0.898);
  txt(s, S.body, A + ' ' + B + ' ' + C, 1.125, 3.236, 9.081, 1.544);
  gradBox(s, 1.125, 5.811, 3.729, 3.259);
  gradBox(s, 9.188, 5.811, 3.729, 3.259);
  [
    ['Full Fasting 30 Days', 1.477, 3.025, CREAM],
    ['Taraweeh & Tadarus', 5.463, 3.116, INK],
    ['Eid Al-Fitr', 9.441, 3.222, CREAM],
  ].forEach(function (col) {
    txt(s, Object.assign({}, S.h3, { color: col[3], align: 'center' }), col[0], col[1], 6.225, col[2], 0.39);
    txt(s, Object.assign({}, S.note, { color: col[3], align: 'center' }), A, col[1], 6.826, col[2], 1.789);
  });
  footer(s, 4);
});

// 5 — Suhoor And Iftar
SLIDES.push(function (s) {
  photo(s, 10.955, 1.557, 4.523, 6.661);
  deco(s, 5);
  txt(s, S.h1, 'Suhoor And Iftar', 1.135, 2.115, 6.27, 0.898);
  [['01', 'Suhoor', 3.63], ['02', 'Iftar', 5.605], ['02', 'Imsak', 7.477]].forEach(function (r, i) {
    numberedRow(s, r[0], 1.193, 3.672 + i * 1.9675, 0.941, 2.46, r[2], 4.237, r[1], r[2] + 0.556, 7.023, 0.695, A + ' ');
  });
  gradBox(s, 10.955, 7.722, 4.522, 0.824);
  txt(s, { fontFace: F_BODY, fontSize: 21, color: WHITE, lineSpacing: 29.4, align: 'center' },
    'Suhoor & Iftar', 11.415, 7.918, 3.602, 0.39);
  footer(s, 5);
  photo(s, 15.477, 4.589, 4.523, 6.661);
});

// 6 — Taraweeh & Tadarus
SLIDES.push(function (s) {
  photo(s, 11.776, 1.137, 5.231, 10.113);
  deco(s, 6);
  gradBox(s, 14.76, 8.273, 4.115, 1.852);
  ICON.mosque(s, 1.265, 3.705, 0.922, 0.868, GOLD);
  ICON.book(s, 1.32, 6.749, 0.866, 0.612, GOLD);
  txt(s, S.h1, 'Taraweeh & Tadarus', 1.125, 2.308, 9.505, 0.898);
  txt(s, S.h2, 'Taraweeh Prayer Together', 2.577, 3.652, 6.899, 0.444);
  txt(s, S.body, A + ' ' + B + ' ' + C, 2.577, 4.313, 8.482, 1.544, { align: 'justify' });
  txt(s, S.h2, 'Tadarus Al-Quran', 2.577, 6.633, 6.899, 0.444);
  txt(s, S.body, A + ' ' + B + ' ' + C, 2.577, 7.294, 8.482, 1.544, { align: 'justify' });
  txt(s, Object.assign({}, S.note, { color: CREAM, align: 'center' }), A, 14.932, 8.45, 3.771, 1.425);
  footer(s, 6);
});

// 7 — Lailatur Qadr
SLIDES.push(function (s) {
  photo(s, 6.491, 1.127, 5.84, 10.123);
  deco(s, 7);
  txt(s, S.h1, 'Lailatur Qadr', 1.125, 2.313, 4.753, 0.898);
  txt(s, S.body, A + ' \n\n' + B + '\n\n ' + C, 1.125, 3.677, 4.753, 4.315);
  [['01', 'Read Al-Quran', 3.594], ['02', 'Tahajud Prayer', 5.666], ['03', 'Dzikir & Doa', 7.737]].forEach(function (r, i) {
    badge(s, 11.811, 3.636 + i * 2.096, 1.04, 1.036, r[0]);
    txt(s, S.h3, r[1], 13.274, r[2], 4.237, 0.39);
    txt(s, S.note, A + ' ', 13.274, r[2] + 0.512, 5.601, 1.06);
  });
  footer(s, 7);
});

// 8 — Benefits Of Fasting In Ramadan
SLIDES.push(function (s) {
  photo(s, 14.068, -0.021, 5.932, 6.193);
  deco(s, 8);
  txt(s, Object.assign({}, S.h1, { lineSpacing: 56.64 }), 'Benefits Of Fasting In Ramadan', 1.125, 1.577, 7.969, 1.565);
  [[3.74, 3.663], [5.505, 5.463], [7.305, 7.298], [8.953, 8.98]].forEach(function (r, i) {
    badge(s, 1.258, r[0], 1.04, 1.036, '0' + (i + 1));
    txt(s, S.h3, 'Description Here', 2.652, r[1], 4.237, 0.39);
    txt(s, S.note, A + ' ', 2.652, r[1] + 0.47, 6.781, 0.695);
  });
  footer(s, 8);
  photo(s, 10.76, 4.517, 5.731, 5.556);
});

// 9 — Eid Al-Fitr
SLIDES.push(function (s) {
  photo(s, -0.024, 0, 20.024, 5.469);
  gradBox(s, 12.371, 4.579, 5.748, 1.852);
  deco(s, 9);
  ICON.heart(s, 10.893, 7.654, 1.057, 0.944, GOLD);
  txt(s, Object.assign({}, S.note, { color: CREAM, align: 'center' }), A, 12.754, 4.939, 4.981, 1.06);
  txt(s, S.h1, 'Eid Al-Fitr', 1.125, 6.509, 8.522, 0.898);
  txt(s, S.h3, 'Eid Mubarak', 1.125, 7.561, 6.504, 0.39);
  txt(s, S.body, A + ' ' + B + ' ' + C, 1.125, 8.133, 8.522, 1.544);
  txt(s, S.h3, 'Halal Bil Halal', 12.371, 7.561, 6.504, 0.39);
  txt(s, S.body, A + '  ' + B.slice(0, -1), 12.371, 8.032, 6.504, 1.544, { align: 'justify' });
  footer(s, 9);
});

// 10 — Our Service
SLIDES.push(function (s) {
  arcPanel(s, 8.788, -0.58, 12.921, 13.465, TEAL_MID);
  deco(s, 10);
  txt(s, S.h1, 'Our Service', 1.125, 2.696, 6.873, 0.898);
  txt(s, S.h3, 'Our Positive Program In Ramadan', 1.183, 4.009, 6.899, 0.39);
  txt(s, S.body, A + ' ' + B + ' \n\n' + C + ' ' + D, 1.183, 4.762, 6.551, 3.523);
  whiteCard(s, 10.677, 1.902, 3.964, 4.008, 10.92, 2.142, 3.477, 2.12, 'Colorful Ramadan', E);
  whiteCard(s, 14.911, 1.902, 3.964, 4.008, 15.167, 2.131, 3.477, 2.12, 'Holly Jolly Month', E);
  whiteCard(s, 10.677, 6.191, 3.964, 3.934, 10.92, 6.462, 3.477, 2.12, 'Charity & Donation', E);
  whiteCard(s, 14.911, 6.191, 3.964, 3.934, 15.154, 6.444, 3.477, 2.12, 'Zakat & Umrah', E);
  plate(s, 1.158, 8.764, 2.786, 0.614, 'Read More');
  footer(s, 10);
});

// 11 — Colorful Ramadan
SLIDES.push(function (s) {
  photo(s, 11.415, 7.033, 4.26, 3.092);
  photo(s, 15.981, 1.054, 4.019, 7.766);
  deco(s, 11);
  txt(s, S.h1, 'Colorful Ramadan', 1.125, 2.17, 8.024, 0.898);
  txt(s, S.h2, 'Hafiz Quran', 1.125, 3.437, 6.899, 0.444);
  txt(s, S.body, A + ' ' + B + ' ' + C, 1.125, 4.033, 8.466, 1.544);
  txt(s, S.h2, 'Iftar Gathering', 1.125, 6.089, 6.899, 0.444);
  txt(s, S.body, A + ' ' + B + ' ' + C, 1.125, 6.685, 8.466, 1.544);
  gradBox(s, 12.684, 4.718, 3.981, 2.845);
  txt(s, Object.assign({}, S.card, { color: CREAM }), A2 + ' Ut enim ad minim veniam,  ', 13.058, 5.22, 3.247, 1.779);
  footer(s, 11);
});

// 12 — Holly Jolly Month
SLIDES.push(function (s) {
  deco(s, 12);
  txt(s, S.h1, 'Holly Jolly Month', 1.125, 2.456, 8.989, 0.898);
  txt(s, S.body, A + ' ' + B + ' ' + C, 1.125, 3.674, 9.177, 1.544);
  numberedRow(s, '01', 1.167, 5.755, 0.941, 2.621, 5.714, 4.237, 'Halal Bihalal', 6.225, 5.817, 1.06, A + ' ');
  numberedRow(s, '02', 1.167, 7.823, 0.941, 2.573, 7.66, 4.237, 'Islamic Studies', 8.145, 5.817, 1.06, A + ' ');
  gradBox(s, 15.476, 3.0, 3.383, 3.065);
  ICON.book(s, 16.796, 3.455, 0.743, 0.607, WHITE, tealAt(0.16));
  txt(s, Object.assign({}, S.h4, { color: CREAM, align: 'center' }), 'Islamic Studies', 15.53, 4.326, 3.242, 0.325);
  txt(s, Object.assign({}, S.card, { color: CREAM }), E, 15.609, 4.88, 3.25, 0.685);
  gradBox(s, 11.822, 6.394, 3.383, 3.02);
  ICON.ketupat(s, 13.259, 6.674, 0.513, 0.837, WHITE);

  txt(s, Object.assign({}, S.h4, { color: CREAM, align: 'center' }), 'Halal Bihalal', 11.89, 7.763, 3.242, 0.325);
  txt(s, Object.assign({}, S.card, { color: CREAM }), E, 11.89, 8.307, 3.25, 0.685);
  footer(s, 12);
  photo(s, 11.819, 3.003, 3.384, 3.066);
  photo(s, 15.476, 6.412, 3.384, 2.985);
});

// 13 — Charity & Donation
SLIDES.push(function (s) {
  photo(s, 14.016, -0.061, 5.984, 10.186);
  deco(s, 13);
  txt(s, S.h1, 'Charity & Donation', 1.125, 3.046, 6.657, 0.898);
  txt(s, S.body, A + ' ' + B + ' ' + C, 1.125, 4.254, 6.657, 2.335, { align: 'justify' });
  txt(s, S.body, D, 1.125, 6.808, 6.657, 0.752, { align: 'justify' });
  gradBox(s, 14.015, 6.59, 4.169, 3.535);
  txt(s, Object.assign({}, S.h3, { color: CREAM, align: 'center' }), 'Description Here', 14.478, 7.325, 3.242, 0.39);
  txt(s, Object.assign({}, S.card, { color: CREAM }), A2 + ' ', 14.295, 7.934, 3.61, 1.414);
  footer(s, 13);
  photo(s, 9.845, 6.59, 4.17, 3.535);
});

// 14 — Zakat & Umrah Program
SLIDES.push(function (s) {
  photo(s, 0, -0.099, 4.696, 7.488);
  photo(s, 4.667, 2.899, 4.63, 7.226);
  deco(s, 14);
  txt(s, Object.assign({}, S.h1, { lineSpacing: 56.16 }), 'Zakat & Umrah Program ', 10.495, 2.903, 9.505, 0.78);
  txt(s, S.h3, 'Our Best Umrah Program During Ramadan', 10.495, 3.984, 6.899, 0.39);
  txt(s, S.body, A + ' ' + B + ' ' + C, 10.495, 4.624, 7.304, 1.94, { align: 'justify' });
  txt(s, S.body, A + ' ' + B + ' ', 10.495, 6.813, 7.304, 1.544, { align: 'justify' });
  gradBox(s, 2.27, 9.301, 4.697, 0.824);
  txt(s, { fontFace: F_BODY, fontSize: 21, color: WHITE, lineSpacing: 29.4, align: 'center' },
    'Zakat & Umroh Program', 2.772, 9.498, 3.693, 0.39);
  footer(s, 14);
});

// 15 — Taraweh Prayer Together
SLIDES.push(function (s) {
  gradBox(s, 15.167, 0, 4.833, 11.25);
  photo(s, 13.45, 1.076, 5.425, 9.049);
  deco(s, 15);
  txt(s, S.h1, 'Taraweh Prayer Together', 1.125, 1.805, 9.505, 0.898);
  txt(s, S.body, A + ' ' + B + ' ' + C, 1.125, 2.908, 8.875, 1.544, { align: 'justify' });
  [['01', 'Full 30 Days ', 5.035, 5.044], ['02', '12 Rakaat', 6.83, 6.788], ['02', 'Tahajud Prayer', 8.579, 8.542]].forEach(function (r) {
    numberedRow(s, r[0], 1.245, r[2], 0.941, 2.512, r[3], 4.237, r[1], r[3] + 0.504, 7.023, 0.695, A + ' ');
  });
  footer(s, 15);
});

// 16 — Our Best Gallery
SLIDES.push(function (s) {
  photo(s, 11.342, 0.58, 4.177, 5.127);
  photo(s, 11.342, 6.01, 4.177, 5.24);
  photo(s, 15.831, -0.144, 4.177, 5.096);
  photo(s, 15.804, 5.254, 4.177, 5.3);
  deco(s, 16);
  plate(s, 12.037, 4.707, 2.786, 0.614, 'Islamic Studies');
  plate(s, 16.599, 4.054, 2.786, 0.614, 'Iftar Gathering');
  plate(s, 16.599, 9.797, 2.786, 0.614, 'Donation');
  plate(s, 12.037, 10.247, 2.786, 0.614, 'Umrah');
  txt(s, S.h1, 'Our Best Gallery', 2.577, 3.207, 6.994, 0.898);
  txt(s, S.body, A + ' ' + B + ' ' + C, 2.577, 4.357, 6.994, 2.335);
  gradBox(s, 2.577, 7.332, 6.994, 1.377);
  txt(s, Object.assign({}, S.note, { color: WHITE }), A + ' ', 2.952, 7.636, 6.243, 0.695);
  footer(s, 16);
});

// 17 — Introduce Our Website
SLIDES.push(function (s) {
  deco(s, 17);
  txt(s, S.h1, 'Introduce Our Website', 1.125, 2.865, 9.505, 0.898);
  gradBox(s, 13.44, 2.541, 5.435, 4.656);
  s.addShape('ellipse', { x: 9.733, y: 8.0, w: 8.329, h: 0.9, fill: { color: 'DFDFDF' } });
  // laptop
  s.addShape('roundRect', { x: 10.696, y: 3.855, w: 6.406, h: 4.296, fill: { color: '000000' }, rectRadius: 0.12 });
  photo(s, 10.891, 4.108, 6.012, 3.762);
  box(s, 13.371, 8.152, 1.051, 0.096, 'CCCCCC');
  box(s, 10.098, 8.344, 7.6, 0.06, 'CCCCCC');
  phoneMock(s, 16.498, 5.653, 1.782, 3.638, 16.59, 5.745, 1.597, 3.458);
  ICON.globe(s, 1.474, 7.335, 0.866, 0.857, GOLD);
  txt(s, S.body, A + ' ' + B + ' ' + C, 1.125, 4.309, 7.134, 1.94);
  txt(s, Object.assign({}, S.note, { lineSpacing: 27.99 }), A2, 2.633, 7.163, 5.277, 1.117, { align: 'justify' });
  footer(s, 17);
});

// 18 — Download Our App
SLIDES.push(function (s) {
  deco(s, 18);
  phoneMock(s, 12.078, 1.406, 4.133, 8.435, 12.296, 1.622, 3.701, 8.016);
  phoneMock(s, 16.809, 1.406, 4.133, 8.435, 17.026, 1.591, 3.701, 8.046);
  txt(s, S.h1, 'Download Our App', 1.125, 2.843, 9.505, 0.898);
  txt(s, Object.assign({}, S.h4, { lineSpacing: 30.42 }),
    'The application is available for download on both the Play Store and the App Store.', 1.125, 4.135, 8.47, 0.794);
  txt(s, S.body, A + ' ' + B + ' ' + C, 1.125, 5.315, 8.47, 1.544, { align: 'justify' });
  txt(s, S.body, A + ' ', 1.125, 7.037, 8.47, 0.752, { align: 'justify' });
  plate(s, 1.125, 8.273, 2.786, 0.614, 'Go To Play Store');
  footer(s, 18);
});

// 19 — Get A Costumized Plan
SLIDES.push(function (s) {
  deco(s, 19);
  txt(s, S.h1, 'Get A Costumized Plan', 5.248, 2.018, 9.505, 0.898, { align: 'center' });
  txt(s, S.body,
    A + ' ' + B.slice(0, -1), 3.269, 3.183, 13.833, 0.752, { align: 'center' });
  const PLANS = [
    { x: 2.371, w: 4.412, name: 'STARTER', price: '$20', ink: CREAM, dark: true, disc: WHITE, mark: TEAL },
    { x: 7.792, w: 4.417, name: 'REGULER', price: '$25', ink: TEAL, dark: false, disc: TEAL, mark: WHITE },
    { x: 13.217, w: 4.417, name: 'PREMIUM', price: '$40', ink: CREAM, dark: true, disc: WHITE, mark: TEAL },
  ];
  PLANS.forEach(function (p) {
    if (p.dark) gradBox(s, p.x, 4.722, p.w, 5.403);
    else box(s, p.x, 4.722, p.w, 5.403, GREY_CARD);
    txt(s, { fontFace: F_BODY, fontSize: 36, color: p.ink, lineSpacing: 43.2, align: 'center' },
      p.name, p.x + 0.912, 5.146, 2.594, 0.583);
    txt(s, { fontFace: F_BOLD, fontSize: 36, bold: true, color: p.ink, lineSpacing: 43.2, align: 'center' },
      p.price, p.x + 1.404, 6.169, 1.609, 0.583);
    txt(s, { fontFace: F_BODY, fontSize: 15.99, color: p.ink, lineSpacing: 22.39, align: 'center' },
      'The Benefits You Will Get :', p.x, 7.258, p.w, 0.289);
    for (let i = 0; i < 4; i++) {
      ICON.tick(s, p.x + 0.713, 7.795 + i * 0.5365, 0.37, 0.37, p.disc, p.mark);
      txt(s, Object.assign({}, S.tiny, { color: p.ink }), 'Your Awesome Service', p.x + 1.207, 7.763 + i * 0.5365, 2.498, 0.349);
    }
  });
  footer(s, 19);
});

// 20 — Leader Team
SLIDES.push(function (s) {
  deco(s, 20);
  s.addShape('star8', { x: 12.251, y: 2.741, w: 5.627, h: 5.627, fill: { color: TEAL }, line: { color: WHITE, width: 2 } });
  photo(s, 12.539, 3.03, 5.051, 5.05, 'star8');
  plate(s, 13.671, 8.77, 2.786, 0.614, 'Umar Fazad');
  txt(s, S.h1, 'Leader Team', 1.553, 2.653, 8.911, 0.898);
  txt(s, Object.assign({}, S.h4, { lineSpacing: 30.42 }), 'Our Best Team Leader', 1.553, 3.753, 5.739, 0.378);
  txt(s, S.body, A + ' ' + B + ' ', 1.553, 4.459, 8.317, 1.148);
  txt(s, S.body, A + ' ' + B + ' ', 1.553, 5.794, 8.317, 1.148);
  gradBox(s, 1.553, 7.369, 8.317, 1.123);
  txt(s, Object.assign({}, S.note, { color: CREAM, align: 'center' }),
    '"Ramadan is the month whose beginning is mercy, whose middle is forgiveness, and whose end is freedom from the fire."',
    1.887, 7.557, 7.567, 0.695);
  footer(s, 20);
});

// 21 — Member Of Our Organizer
SLIDES.push(function (s) {
  deco(s, 21);
  txt(s, S.h1, 'Member Of Our Organizer', 5.247, 2.008, 9.505, 0.898, { align: 'center' });
  const TEAM = [
    { n: '01', name: 'Ghazal Habib', px: 1.536, bx: 1.615, lx: 1.805, tx: 2.118, cx: 1.495, ly: 7.343, ny: 7.479 },
    { n: '02', name: 'Ayesha Bilqis', px: 6.061, bx: 6.14, lx: 6.339, tx: 6.652, cx: 6.02, ly: 7.343, ny: 7.479 },
    { n: '03', name: 'Anwar Mazed', px: 10.586, bx: 10.725, lx: 10.873, tx: 11.185, cx: 10.563, ly: 7.332, ny: 7.469 },
    { n: '04', name: 'Nadya Qamara', px: 15.111, bx: 15.219, lx: 15.412, tx: 15.721, cx: 15.099, ly: 7.322, ny: 7.458 },
  ];
  TEAM.forEach(function (m) {
    photo(s, m.px, 3.615, 3.323, 3.321, 'ellipse');
    badge(s, m.bx, 3.614, 0.941, 0.938, m.n);
    plate(s, m.lx, m.ly, 2.786, 0.614, m.name);
    txt(s, S.cap, 'Your Position', m.tx, 8.146, 2.161, 0.289);
    txt(s, Object.assign({}, S.tiny, { align: 'center' }), E + ',', m.cx, 8.569, 3.406, 0.724);
  });
  footer(s, 21);
});

// 22 — Contact Information
SLIDES.push(function (s) {
  arcPanel(s, 9.11, -0.708, 12.921, 13.465, TEAL_MID);
  s.addShape('ellipse', { x: 9.851, y: 8.363, w: 9.172, h: 1.366, fill: { color: '0A6E66' } });
  deco(s, 22);
  photo(s, 10.0, 2.765, 8.875, 5.991);
  txt(s, S.h1, 'Contact Information', 1.683, 2.317, 8.375, 0.898);
  const ROWS = [
    ['pin', 'Address', '123 Anywhere St., Any City, ST 12345', 1.683, 3.568, 0.236, 0.337, 2.197, 3.527, 2.201, 1.683, 4.1],
    ['globe', 'Website', 'www.reallygreatsite.com ', 1.683, 5.095, 0.337, 0.337, 2.197, 5.066, 1.955, 1.683, 5.653],
    ['phone', 'Telephone', '+123-456-7890', 1.711, 6.646, 0.337, 0.337, 2.224, 6.617, 2.06, 1.711, 7.204],
    ['mail', 'Email', 'reallygreatsite@gmail.com', 1.711, 8.273, 0.37, 0.277, 2.267, 8.156, 2.06, 1.736, 8.71],
  ];
  ROWS.forEach(function (r) {
    ICON[r[0]](s, r[3], r[4], r[5], r[6], TEAL);
    txt(s, { fontFace: F_BOLD, fontSize: 21.99, bold: true, color: INK, lineSpacing: 30.79 }, r[1], r[7], r[8], r[9], 0.408);
    txt(s, { fontFace: F_BODY, fontSize: 21, color: INK, lineSpacing: 29.4 }, r[2], r[10], r[11], 5.554, 0.39);
  });
  footer(s, 22);
});

// 23 — Thank You
SLIDES.push(function (s) {
  cartouche(s, 1.597, 1.285, 16.612, 8.68);
  corner(s, 0, 0, 5.263, 4.987);
  corner(s, 14.737, 6.228, 5.263, 4.987, true);
  mandala(s, 0, 9.16, 1.786);
  mandala(s, 20, 2.061, 1.786);
  txt(s, { fontFace: F_TITLE, fontSize: 104, bold: true, color: WHITE, lineSpacing: 114.4, align: 'center' },
    'THANK YOU', 4.885, 4.865, 10.229, 1.624);
  txt(s, { fontFace: F_BODY, fontSize: 18, color: WHITE, lineSpacing: 27, align: 'center' },
    '"May your Ramadan be filled with joy and happiness."', 5.72, 6.539, 8.366, 0.352);
});

/* ---------------------------------------------------------------------- main */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'RAMADAN_20x11_25', width: 20, height: 11.25 });
  pptx.layout = 'RAMADAN_20x11_25';
  pptx.title = 'Fasting Blessings';

  SLIDES.forEach(function (draw) {
    const slide = pptx.addSlide();
    slide.background = { color: WHITE };
    draw(slide);
  });

  return pptx.writeFile({
    fileName: path.join(__dirname, '0c89788b-7a76-4845-8111-43d34112e0a6_grok_final.pptx'),
  });
}

build().then(function (f) { console.log('wrote ' + f); });
