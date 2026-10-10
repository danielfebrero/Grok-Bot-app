/**
 * Xectol — Modern Real Estate Presentation (18 slides, 13.333in x 7.5in)
 * Recreated with pptxgenjs. Raster/photo placeholders are drawn as grey
 * rounded rectangles labelled "[image]"; every other element is a native shape.
 *
 *   node 131f850a-07da-4896-b3df-5ab098ec465c_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const BLACK = '000000';
const WHITE = 'FFFFFF';
const LIME = 'E1FF01';
const NAVY = '0A0635';
const PHOTO_BG = 'AFB0B0';
const PHOTO_FG = '4D4D4D';

/* -------------------------------------------------------------------- fonts */
const F = 'Bricolage Grotesque';
const FB = 'Bricolage Grotesque SemiBold';
const FS = 'Sarabun SemiBold';
const FSR = 'Sarabun';
const BAR = 'Barlow SemiBold';
const BARR = 'Barlow regular';

/* ------------------------------------------------------- boilerplate copy */
const L_FULL = 'There are many variations of passages of Lorem Ipsum available, but the majority have suffered.';
const L_SHORT = 'There are many variations of passages of Lorem Ipsum.';
const L_VAR = 'variations of passages of Lorem Ipsum available, but the majority.';
const L_VARS = 'variations of passages of Lorem Ipsum available, but the majority have suffered.';
const L_MANY = 'many variations of passages of Lorem Ipsum available, but the majority have suffered.';
const L_MANY_S = 'many variations of passages of Lorem Ipsum available, but the majority.';
const L_THERE = 'There are many variations of passages of Lorem Ipsum available, but the majority.';
const TWO_LINE = [{ text: 'There are many variat', options: { breakLine: true } }, { text: 'ions of passages.' }];

/* ---------------------------------------------------------------- helpers */

/** roundRect adjust value (OOXML "adj", 0..50000) -> pptxgenjs rectRadius in inches */
const adj = (val, w, h) => (val * Math.min(w, h)) / 100000;

/** plain text box; reference boxes are top-anchored with the default 0.1in inset */
function T(s, text, x, y, w, h, o) {
  s.addText(text, Object.assign({
    x, y, w, h, fontFace: F, fontSize: 10, color: BLACK,
    align: 'left', valign: 'top', margin: [7.2, 7.2, 3.6, 3.6], isTextBox: true,
  }, o));
}

/** filled rounded rectangle (the deck's card / chip primitive) */
function panel(s, x, y, w, h, color, adjVal) {
  s.addShape('roundRect', { x, y, w, h, fill: { color }, rectRadius: adj(adjVal || 8045, w, h) });
}

function rect(s, x, y, w, h, color) {
  s.addShape('rect', { x, y, w, h, fill: { color } });
}

function dot(s, cx, cy, r, color) {
  s.addShape('ellipse', { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, fill: { color } });
}

function ring(s, cx, cy, r, thick, color) {
  s.addShape('donut', { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, fill: { color }, rectRadius: thick });
}

function hline(s, x, y, w) {
  s.addShape('line', { x, y, w, h: 0, line: { color: BLACK, width: 1 } });
}

function vline(s, x, y, h) {
  s.addShape('line', { x, y, w: 0, h, line: { color: BLACK, width: 0.5 } });
}

/** closed custom outline; pts are fractions of w/h, 'q' entries add a quadratic control */
function poly(s, x, y, w, h, color, pts, o) {
  const P = pts.map((p, i) => {
    const pt = { x: p[0] * w, y: p[1] * h };
    if (i === 0) pt.moveTo = true;
    if (p.length === 4) pt.curve = { type: 'quadratic', x1: p[2] * w, y1: p[3] * h };
    if (p.length === 6) pt.curve = { type: 'cubic', x1: p[2] * w, y1: p[3] * h, x2: p[4] * w, y2: p[5] * h };
    return pt;
  });
  s.addShape('custGeom', Object.assign({ x, y, w, h, fill: { color }, points: P.concat([{ close: true }]) }, o));
}

/** the four-point concave "sparkle" that marks every rule intersection */
function sparkle(s, x, y, d) {
  poly(s, x, y, d, d, BLACK, [[0.5, 0], [1, 0.5, 0.56, 0.44], [0.5, 1, 0.56, 0.56], [0, 0.5, 0.44, 0.56], [0.5, 0, 0.44, 0.44]]);
}

/** grey stand-in for a photograph */
function photo(s, x, y, w, h) {
  s.addShape('roundRect', { x, y, w, h, fill: { color: PHOTO_BG }, rectRadius: adj(8095, w, h) });
  T(s, '[image]', x, y + h / 2 - 0.16, w, 0.32, { align: 'center', color: PHOTO_FG, fontSize: 12 });
}

/** black or lime pill button with its caption */
function button(s, x, y, dark, label, tx, ty, tw, adjVal) {
  s.addShape('roundRect', {
    x, y, w: 1.293, h: 0.359, fill: { color: dark ? BLACK : LIME },
    rectRadius: adj(adjVal || 18313, 1.293, 0.359),
  });
  T(s, label, tx, ty, tw, 0.269, { fontFace: FS, fontSize: 10, color: dark ? WHITE : BLACK, align: 'center' });
}

/** big number + "There are many" caption pair */
function stat(s, x, y, value, ly, color) {
  T(s, value, x, y, 1.07, 0.438, { fontFace: FB, fontSize: 20, color: color || BLACK });
  T(s, 'There are many', x, ly, 1.073, 0.236, { fontSize: 8, color: color || BLACK });
}

/** card = rounded panel + 13/14pt heading + 8pt body */
function card(s, box, head, body, o) {
  o = o || {};
  const fg = o.dark ? WHITE : BLACK;
  panel(s, box[0], box[1], box[2], box[3], o.dark ? BLACK : LIME);
  T(s, head.t, head.x, head.y, head.w, 0.32, { fontFace: FB, fontSize: head.size || 13, color: fg });
  T(s, body.t, body.x, body.y, body.w, 0.484, { fontSize: 8, color: fg, lineSpacingMultiple: 1.5 });
}

/* ------------------------------------------------------------------- icons */
/* Every glyph draws inside the unit box (x,y,d,d) of its badge circle. */
const ICON = {
  // crescent + dot brand mark
  cdot: (s, x, y, d, c) => {
    dot(s, x + 0.44 * d, y + 0.56 * d, 0.26 * d, c);
    dot(s, x + 0.72 * d, y + 0.32 * d, 0.13 * d, c);
  },
  flame: (s, x, y, d, c) => {
    // outer blaze: a tall right tongue with a shorter one curling off to the left
    poly(s, x + 0.26 * d, y + 0.16 * d, 0.48 * d, 0.68 * d, c, [
      [0.55, 0], [0.22, 0.34, 0.52, 0.18, 0.3, 0.24], [0.34, 0.5, 0.14, 0.45, 0.32, 0.46],
      [0.0, 0.62, 0.14, 0.56, 0.0, 0.5], [0.5, 1.0, 0.0, 0.86, 0.2, 1.0],
      [1.0, 0.62, 0.86, 1.0, 1.0, 0.85], [0.78, 0.2, 1.0, 0.42, 0.84, 0.3],
      [0.72, 0.42, 0.82, 0.34, 0.78, 0.4], [0.55, 0, 0.7, 0.26, 0.72, 0.08]]);
    // white core notch that splits the base into two licks of fire
    poly(s, x + 0.4 * d, y + 0.52 * d, 0.22 * d, 0.3 * d, WHITE, [
      [0.5, 0], [0.0, 0.55, 0.1, 0.3], [0.5, 1.0, 0.0, 0.9], [1.0, 0.55, 1.0, 0.9], [0.5, 0, 0.9, 0.3]]);
  },
  eye: (s, x, y, d, c) => {
    poly(s, x + 0.13 * d, y + 0.35 * d, 0.74 * d, 0.3 * d, WHITE,
      [[0, 0.5], [1, 0.5, 0.5, -0.55], [0, 0.5, 0.5, 1.55]], { line: { color: c, width: 1.25 } });
    dot(s, x + 0.5 * d, y + 0.5 * d, 0.115 * d, c);
  },
  target: (s, x, y, d, c) => {
    ring(s, x + 0.46 * d, y + 0.54 * d, 0.32 * d, 0.045 * d, c);
    ring(s, x + 0.46 * d, y + 0.54 * d, 0.2 * d, 0.045 * d, c);
    dot(s, x + 0.46 * d, y + 0.54 * d, 0.06 * d, c);
    // dart flying in from the upper right
    s.addShape('rect', { x: x + 0.46 * d, y: y + 0.5 * d, w: 0.36 * d, h: 0.04 * d, fill: { color: c }, rotate: -45 });
    poly(s, x + 0.66 * d, y + 0.16 * d, 0.2 * d, 0.2 * d, c, [[1, 0], [1, 0.55], [0.45, 1], [0, 0.5], [0.5, 0.5]]);
  },
  atom: (s, x, y, d, c) => {
    [0, 60, 120].forEach((rotate) => s.addShape('donut', {
      x: x + 0.06 * d, y: y + 0.31 * d, w: 0.88 * d, h: 0.38 * d,
      fill: { color: c }, rectRadius: 0.022 * d, rotate,
    }));
    dot(s, x + 0.5 * d, y + 0.5 * d, 0.11 * d, WHITE);
    dot(s, x + 0.5 * d, y + 0.5 * d, 0.085 * d, c);
  },
  search: (s, x, y, d, c) => {
    ring(s, x + 0.43 * d, y + 0.41 * d, 0.26 * d, 0.05 * d, c);
    s.addShape('rect', { x: x + 0.6 * d, y: y + 0.63 * d, w: 0.24 * d, h: 0.055 * d, fill: { color: c }, rotate: 45 });
  },
  hourglass: (s, x, y, d, c) => {
    const stroke = { line: { color: c, width: 1.25 } };
    poly(s, x + 0.29 * d, y + 0.25 * d, 0.42 * d, 0.25 * d, WHITE, [[0, 0], [1, 0], [0.5, 1]], stroke);
    poly(s, x + 0.29 * d, y + 0.5 * d, 0.42 * d, 0.25 * d, WHITE, [[0.5, 0], [1, 1], [0, 1]], stroke);
    rect(s, x + 0.33 * d, y + 0.28 * d, 0.34 * d, 0.09 * d, c);          // sand, upper bulb
    poly(s, x + 0.38 * d, y + 0.62 * d, 0.24 * d, 0.12 * d, c, [[0.5, 0], [1, 1], [0, 1]]);
    rect(s, x + 0.22 * d, y + 0.19 * d, 0.56 * d, 0.06 * d, c);          // top plate
    rect(s, x + 0.22 * d, y + 0.75 * d, 0.56 * d, 0.06 * d, c);          // base plate
  },
  rocket: (s, x, y, d, c) => {
    poly(s, x + 0.3 * d, y + 0.18 * d, 0.5 * d, 0.5 * d, c,
      [[1, 0], [0.15, 0.7, 0.7, 0.2, 0.4, 0.45], [0, 1], [0.3, 0.85, 0.05, 0.95],
        [0.8, 0.3, 0.55, 0.5, 0.8, 0.3]], { line: { color: c, width: 1 } });
    dot(s, x + 0.62 * d, y + 0.36 * d, 0.06 * d, WHITE);
    poly(s, x + 0.2 * d, y + 0.42 * d, 0.2 * d, 0.22 * d, c, [[1, 0], [0, 1], [0.85, 0.85]]);
    poly(s, x + 0.44 * d, y + 0.62 * d, 0.22 * d, 0.2 * d, c, [[1, 0], [0.15, 0.15], [0, 1]]);
  },
  badge: (s, x, y, d, c) => {
    poly(s, x + 0.28 * d, y + 0.62 * d, 0.2 * d, 0.24 * d, c, [[0, 0], [1, 0], [1, 1], [0.5, 0.7], [0, 1]]);
    poly(s, x + 0.52 * d, y + 0.62 * d, 0.2 * d, 0.24 * d, c, [[0, 0], [1, 0], [1, 1], [0.5, 0.7], [0, 1]]);
    ring(s, x + 0.5 * d, y + 0.4 * d, 0.28 * d, 0.05 * d, c);
    poly(s, x + 0.36 * d, y + 0.32 * d, 0.28 * d, 0.2 * d, c,
      [[0, 0.45], [0.1, 0.3], [0.4, 0.62], [0.9, 0], [1, 0.15], [0.4, 0.95]]);
  },
  coins: (s, x, y, d, c) => {
    // back stack, front stack, then a dollar coin overlapping at the left
    [[0.36, 0.28], [0.46, 0.5]].forEach(([cx, top]) => {
      s.addShape('can', {
        x: x + cx * d, y: y + top * d, w: 0.36 * d, h: 0.26 * d,
        fill: { color: WHITE }, line: { color: c, width: 1 },
      });
    });
    s.addShape('ellipse', {
      x: x + 0.18 * d, y: y + 0.42 * d, w: 0.28 * d, h: 0.28 * d,
      fill: { color: WHITE }, line: { color: c, width: 1.25 },
    });
    T(s, '$', x + 0.18 * d, y + 0.45 * d, 0.28 * d, 0.22 * d,
      { fontFace: FB, fontSize: 6, align: 'center', color: c, margin: 0 });
  },
  hand: (s, x, y, d, c) => {
    s.addShape('ellipse', {
      x: x + 0.35 * d, y: y + 0.15 * d, w: 0.32 * d, h: 0.32 * d,
      fill: { color: WHITE }, line: { color: c, width: 1.25 },
    });
    rect(s, x + 0.495 * d, y + 0.21 * d, 0.04 * d, 0.2 * d, c);
    // open palm cupping the coin: fingers sweep right, sleeve cuff bottom-left
    poly(s, x + 0.14 * d, y + 0.5 * d, 0.74 * d, 0.32 * d, c,
      [[0.1, 0.3], [0.55, 0.35, 0.3, 0.62], [0.95, 0.0, 0.82, 0.05],
        [1, 0.2, 1.06, 0.06], [0.5, 0.72, 0.8, 0.68], [0.18, 0.62]]);
    poly(s, x + 0.1 * d, y + 0.56 * d, 0.2 * d, 0.26 * d, c, [[0.5, 0], [1, 0.5], [0.5, 1], [0, 0.5]]);
  },
  // handshake: two sleeve cuffs angled in from the sides, hands clasped in between
  handshake: (s, x, y, d, c) => {
    poly(s, x + 0.28 * d, y + 0.36 * d, 0.44 * d, 0.34 * d, c,
      [[0, 0.1], [0.35, 0], [1, 0.55], [0.75, 1], [0.1, 0.55]]);
    [[0.12, 0.3, -22], [0.62, 0.3, 22]].forEach(([cx, cy, rotate]) => {
      s.addShape('roundRect', {
        x: x + cx * d, y: y + cy * d, w: 0.26 * d, h: 0.22 * d, rotate,
        fill: { color: c }, rectRadius: 0.05 * d,
      });
      dot(s, x + (cx + 0.07) * d, y + (cy + 0.11) * d, 0.03 * d, WHITE);
    });
  },
  fingerprint: (s, x, y, d, c) => {
    // nested open ridges, each a partial arc, so the print reads as a thumb
    [[0.36, 200], [0.27, 230], [0.18, 260], [0.09, 300]].forEach(([r, sweep]) => {
      s.addShape('blockArc', {
        x: x + (0.5 - r) * d, y: y + (0.5 - r) * d, w: 2 * r * d, h: 2 * r * d,
        fill: { color: c }, angleRange: [270 - sweep / 2, 270 + sweep / 2],
        arcThicknessRatio: 0.05 / r,
      });
    });
  },
};

/** white/black/lime badge circle with a glyph inside */
function icon(s, x, y, d, bg, glyph, fg) {
  dot(s, x + d / 2, y + d / 2, d / 2, bg);
  ICON[glyph](s, x, y, d, fg);
}

/* -------------------------------------------------- shared top navigation */
function nav(s, o) {
  o = o || {};
  // lime "Z" brand mark (traced from the original 211 x 191 outline)
  poly(s, 1.15, 0.561, 0.263, 0.221, LIME, [
    [0.038, 0], [0.009, 0.042, 0, 0.026], [0.114, 0.209, 0.09, 0.194],
    [0.692, 0.209], [0.133, 0.832], [0.147, 0.958, 0.157, 0.859],
    [0.284, 1.0], [1, 0.209], [1, 0]]);
  rect(s, 1.283, 0.751, 0.13, 0.052, LIME);

  T(s, 'Xectol', 1.488, 0.497, 1.057, 0.337, { fontFace: FB, fontSize: 14, align: 'center' });
  [['Home', 2.891, 0.726], ['Featuras', 3.962, 0.931], ['About Us', 5.238, 0.882],
    ['Agants', 6.467, 0.882], ['ENG', 9.896, 0.545]].forEach(([t, x, w]) => {
    T(s, t, x, 0.556, w, 0.286, { fontSize: 11, align: 'center' });
  });
  s.addShape('roundRect', {
    x: 10.752, y: 0.457, w: 1.427, h: 0.455, fill: { color: BLACK },
    rectRadius: adj(17431, 1.427, 0.455),
  });
  T(s, 'Contact', 11.026, 0.53, 0.878, 0.286, { fontFace: FB, fontSize: 11, color: WHITE, align: 'center' });

  hline(s, 1.155, 1.111, 11.024);
  sparkle(s, 9.024, 0.969, 0.285);
  if (!o.noTick) vline(s, 9.167, 0, 1.111);
}

/* ============================================================== the slides */
const SLIDES = [
  /* 1 — hero */
  (s) => {
    nav(s);
    photo(s, 6.603, 2.394, 2.268, 3.038);
    photo(s, 9.167, 1.781, 3.012, 3.642);
    T(s, 'Xectol', 1.059, 1.587, 4.141, 1.446, { fontFace: FB, fontSize: 80 });
    T(s, 'Modern Real Estate Presentation', 1.059, 3.075, 3.351, 0.286, { fontSize: 11 });
    T(s, L_FULL, 1.059, 3.443, 3.733, 0.606, { lineSpacingMultiple: 1.5 });

    panel(s, 1.155, 4.394, 1.682, 1.033, LIME);
    T(s, 'Total Balance', 1.389, 4.582, 1.214, 0.252, { fontSize: 9, align: 'center' });
    T(s, '$267.08K', 1.158, 4.802, 1.677, 0.438, { fontFace: FB, fontSize: 20, align: 'center' });
    panel(s, 3.042, 4.394, 1.684, 1.033, LIME);
    T(s, '92%', 3.17, 4.592, 1.069, 0.438, { fontFace: FB, fontSize: 20 });
    T(s, 'Return Customers', 3.17, 4.993, 1.387, 0.252, { fontSize: 9 });
    icon(s, 4.203, 4.535, 0.354, WHITE, 'eye', BLACK);

    hline(s, 1.155, 5.807, 11.024);
    vline(s, 5.705, 5.807, 1.693);
    sparkle(s, 5.562, 5.67, 0.285);

    icon(s, 1.193, 6.095, 0.303, BLACK, 'cdot', WHITE);
    T(s, '1.9k', 1.656, 6.028, 0.802, 0.403, { fontFace: FB, fontSize: 18 });
    T(s, TWO_LINE, 1.672, 6.41, 1.448, 0.505, { fontSize: 8, lineSpacingMultiple: 1.5 });
    icon(s, 3.259, 6.092, 0.304, LIME, 'cdot', BLACK);
    T(s, 'The Details.', 3.616, 6.078, 1.28, 0.302, { fontFace: FB, fontSize: 12 });
    T(s, TWO_LINE, 3.616, 6.394, 1.448, 0.505, { fontSize: 8, lineSpacingMultiple: 1.5 });

    button(s, 6.035, 6.262, true, 'Explore Now', 6.188, 6.306, 0.986);
    button(s, 7.481, 6.266, false, 'Learn Now', 7.663, 6.311, 0.929, 19765);
    T(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque leo eros, pretium..',
      9.167, 6.062, 2.925, 0.531, { fontSize: 9, lineSpacingMultiple: 1.5 });

    icon(s, 11.719, 2.062, 0.355, LIME, 'flame', BLACK);
    icon(s, 6.766, 2.578, 0.356, BLACK, 'cdot', WHITE);
  },

  /* 2 — Introduction */
  (s) => {
    nav(s);
    photo(s, 7.371, 1.781, 4.808, 5.118);
    T(s, 'Introduction', 1.059, 1.977, 3.733, 0.774, { fontFace: FB, fontSize: 40 });
    T(s, L_FULL, 1.059, 3.003, 3.455, 0.606, { lineSpacingMultiple: 1.5 });

    panel(s, 1.155, 4.125, 1.906, 1.273, BLACK);
    T(s, '95K+', 1.53, 4.259, 1.156, 0.573, { fontFace: BAR, fontSize: 28, color: WHITE, align: 'center' });
    T(s, 'Sed utThere are many varia tions of passages of Loremi.', 1.271, 4.786, 1.674, 0.477,
      { fontFace: BARR, fontSize: 8, color: WHITE, align: 'center', lineSpacingMultiple: 1.5 });

    panel(s, 3.293, 4.125, 1.906, 1.273, LIME);
    [[3.436, 4.477, 0.786, BLACK], [3.611, 4.898, 0.366, WHITE], [3.788, 4.681, 0.583, BLACK],
      [3.964, 4.552, 0.712, WHITE], [4.139, 4.851, 0.413, BLACK]].forEach(([x, y, h, c]) => {
      panel(s, x, y, 0.133, h, c, 24559);
    });
    T(s, 'Total Courses', 4.271, 4.243, 0.943, 0.234, { fontSize: 8 });
    T(s, '280+', 4.271, 4.443, 0.885, 0.438, { fontFace: FB, fontSize: 20 });

    vline(s, 6.667, 1.111, 6.389);
    hline(s, 1.155, 5.913, 5.512);
    vline(s, 3.856, 5.913, 1.587);
    sparkle(s, 3.714, 5.771, 0.285);

    button(s, 1.302, 6.231, false, 'Explore Now', 1.457, 6.276, 0.986);
    icon(s, 2.939, 6.207, 0.408, BLACK, 'cdot', WHITE);
    stat(s, 4.181, 6.116, '185K', 6.495);
    stat(s, 5.359, 6.099, '103 M', 6.495);
    icon(s, 11.549, 2.062, 0.356, LIME, 'flame', NAVY);
  },

  /* 3 — Commercial Real Estate */
  (s) => {
    nav(s);
    photo(s, 1.155, 3.75, 4.984, 3.149);
    T(s, 'Commercial Real Estate', 1.075, 1.927, 5.214, 0.606, { fontFace: FB, fontSize: 30 });
    T(s, L_FULL, 1.075, 2.743, 3.717, 0.606, { lineSpacingMultiple: 1.5 });
    icon(s, 1.424, 3.95, 0.355, BLACK, 'flame', WHITE);

    vline(s, 6.667, 1.111, 6.389);
    hline(s, 6.667, 5.37, 6.667);
    sparkle(s, 6.524, 5.227, 0.285);

    card(s, [7.194, 1.781, 2.358, 1.389], { x: 7.438, y: 2.165, w: 1.37, t: 'Financing' },
      { x: 7.438, y: 2.497, w: 1.736, t: L_SHORT }, { dark: 1 });
    card(s, [9.812, 1.792, 2.358, 1.389], { x: 10.056, y: 2.175, w: 1.368, t: 'Consulting' },
      { x: 10.056, y: 2.507, w: 1.736, t: L_SHORT });
    card(s, [7.203, 3.429, 2.358, 1.389], { x: 7.446, y: 3.812, w: 1.368, t: 'Construction' },
      { x: 7.446, y: 4.144, w: 1.736, t: L_SHORT });
    card(s, [9.821, 3.439, 2.358, 1.389], { x: 10.064, y: 3.823, w: 1.634, t: 'Manangement' },
      { x: 10.064, y: 4.155, w: 1.736, t: L_SHORT }, { dark: 1 });
    icon(s, 8.847, 1.892, 0.42, WHITE, 'flame', BLACK);
    icon(s, 11.5, 1.898, 0.42, WHITE, 'badge', BLACK);
    icon(s, 8.849, 3.594, 0.42, WHITE, 'rocket', BLACK);
    icon(s, 11.521, 3.589, 0.447, WHITE, 'handshake', BLACK);

    panel(s, 7.203, 5.696, 1.266, 0.951, BLACK);
    stat(s, 7.299, 5.865, '185K', 6.245, WHITE);
    panel(s, 8.8, 5.696, 1.264, 0.951, LIME);
    stat(s, 8.898, 5.856, '103 M', 6.252);
    button(s, 10.441, 6.064, true, 'Explore Now', 10.594, 6.109, 0.986);
  },

  /* 4 — Industrial Real Estate */
  (s) => {
    nav(s);
    photo(s, 1.155, 1.781, 2.274, 1.969);
    photo(s, 7.223, 4.683, 4.956, 2.217);
    icon(s, 1.253, 1.901, 0.356, BLACK, 'flame', WHITE);
    T(s, 'Investment', 3.698, 2.123, 2.042, 0.503, { fontFace: FB, fontSize: 24 });
    T(s, L_VAR, 3.698, 2.776, 2.455, 0.58, { lineSpacingMultiple: 1.5 });

    vline(s, 6.668, 1.111, 6.389);
    sparkle(s, 6.526, 3.898, 0.285);
    T(s, 'Industrial Real Estate', 7.222, 1.934, 3.733, 1.448, { fontFace: FB, fontSize: 40 });
    T(s, L_FULL, 7.222, 3.729, 3.53, 0.606, { lineSpacingMultiple: 1.5 });
    icon(s, 11.661, 4.833, 0.356, LIME, 'cdot', BLACK);

    hline(s, 1.155, 4.038, 5.514);
    card(s, [1.155, 4.375, 2.405, 1.45], { x: 1.323, y: 4.665, w: 1.688, t: 'Industrial One', size: 14 },
      { x: 1.323, y: 5.031, w: 2.069, t: L_VAR }, { dark: 1 });
    card(s, [3.911, 4.349, 2.405, 1.448], { x: 4.08, y: 4.637, w: 1.66, t: 'Industrial Two', size: 14 },
      { x: 4.08, y: 5.003, w: 2.068, t: L_VAR });
    stat(s, 1.311, 6.212, '185K', 6.59);
    stat(s, 2.49, 6.194, '103 M', 6.589);
    button(s, 3.91, 6.309, false, 'Explore Now', 4.062, 6.352, 0.986);
    icon(s, 5.514, 6.24, 0.409, BLACK, 'cdot', WHITE);
  },

  /* 5 — Retail Real Estate */
  (s) => {
    nav(s, { noTick: true });
    photo(s, 6.958, 1.781, 5.22, 2.748);
    T(s, 'Retail Real Estate', 1.092, 1.795, 5.052, 0.774, { fontFace: FB, fontSize: 40 });
    T(s, L_FULL, 1.092, 2.771, 3.7, 0.606, { lineSpacingMultiple: 1.5 });
    button(s, 1.155, 3.988, true, 'Explore Now', 1.309, 4.033, 0.984);
    button(s, 2.602, 3.993, false, 'Learn Now', 2.785, 4.038, 0.929, 19765);
    stat(s, 4.257, 3.793, '185K', 4.172);

    vline(s, 6.462, 1.111, 6.389);
    hline(s, 1.155, 4.875, 5.307);
    sparkle(s, 6.323, 4.733, 0.285);
    icon(s, 11.684, 1.979, 0.355, BLACK, 'flame', WHITE);

    card(s, [1.167, 5.217, 2.358, 1.389], { x: 1.517, y: 5.503, w: 1.368, t: 'Retail One', size: 14 },
      { x: 1.517, y: 5.835, w: 1.745, t: L_SHORT });
    card(s, [3.807, 5.217, 2.358, 1.389], { x: 4.16, y: 5.503, w: 1.368, t: 'Reatil Two', size: 14 },
      { x: 4.16, y: 5.835, w: 1.745, t: L_SHORT }, { dark: 1 });

    panel(s, 6.958, 4.889, 5.22, 2.01, LIME);
    [[7.542, 5.486, 25926], [8.026, 5.691, 19382], [8.51, 6.097, 25926], [8.99, 5.486, 21564]].forEach(
      ([x, fy, a]) => {
        panel(s, x, 5.2, 0.319, 1.389, WHITE, 24404);
        s.addShape('round2SameRect', {
          x, y: fy, w: 0.319, h: 6.589 - fy, fill: { color: BLACK }, rotate: 180,
          rectRadius: adj(a, 0.319, 6.589 - fy),
        });
      });
    T(s, 'Revenue Last Quarter', 9.684, 5.201, 1.661, 0.269, { fontSize: 10 });
    T(s, '+350%', 9.684, 5.5, 2.22, 0.707, { fontFace: FB, fontSize: 36 });
    T(s, 'variations of passages of Lorem.', 9.684, 6.234, 2.417, 0.328, { lineSpacingMultiple: 1.5 });
  },

  /* 6 — Investment Strategies */
  (s) => {
    nav(s);
    photo(s, 1.155, 3.0, 4.984, 2.425);
    T(s, 'Real Estate Investment Strategies', 1.059, 1.656, 5.061, 1.111, { fontFace: FB, fontSize: 30 });
    icon(s, 1.332, 3.141, 0.355, LIME, 'flame', BLACK);
    vline(s, 6.667, 1.111, 4.583);

    card(s, [7.073, 1.781, 2.41, 1.656], { x: 7.212, y: 2.238, w: 1.502, t: 'Safer Option' },
      { x: 7.212, y: 2.604, w: 2.069, t: L_VAR }, { dark: 1 });
    card(s, [9.769, 1.781, 2.41, 1.656], { x: 9.91, y: 2.238, w: 1.5, t: 'Leveraging' },
      { x: 9.91, y: 2.604, w: 2.068, t: L_VAR });
    card(s, [7.073, 3.767, 2.41, 1.658], { x: 7.212, y: 4.226, w: 1.955, t: 'Inflation Hedging' },
      { x: 7.212, y: 4.59, w: 2.069, t: L_VAR });
    card(s, [9.769, 3.767, 2.41, 1.658], { x: 9.91, y: 4.226, w: 1.714, t: 'Passive Income' },
      { x: 9.91, y: 4.59, w: 2.068, t: L_VAR }, { dark: 1 });
    icon(s, 8.714, 2.005, 0.496, WHITE, 'coins', BLACK);
    icon(s, 11.415, 2.002, 0.496, WHITE, 'target', BLACK);
    icon(s, 8.903, 3.936, 0.365, WHITE, 'atom', BLACK);
    icon(s, 11.564, 3.91, 0.413, WHITE, 'eye', BLACK);

    hline(s, 1.155, 5.694, 11.024);
    sparkle(s, 6.524, 5.554, 0.285);
    T(s, L_FULL, 1.075, 6.049, 3.455, 0.606, { lineSpacingMultiple: 1.5 });
    button(s, 5.146, 6.191, false, 'Explore Now', 5.3, 6.234, 0.984);
    button(s, 6.594, 6.194, true, 'Learn Now', 6.776, 6.24, 0.929, 19765);
    stat(s, 8.503, 6.028, '185K', 6.408);
    stat(s, 9.682, 6.01, '103 M', 6.406);
    icon(s, 11.03, 6.095, 0.409, LIME, 'cdot', BLACK);
  },

  /* 7 — Factors Influencing */
  (s) => {
    nav(s);
    [['Industry Sector', 1.592, 1.762, 1.707, 1],
      ['Financial Performance', 4.172, 4.314, 2.175, 0],
      ['Investor Sentiment', 6.752, 6.922, 1.965, 1],
      ['Investor Perception', 9.332, 9.502, 1.965, 0]].forEach(([t, bx, tx, tw, dark]) => {
      card(s, [bx, 1.958, 2.41, 1.656], { x: tx, y: 2.349, w: tw, t },
        { x: tx, y: 2.717, w: 2.069, t: L_VAR }, { dark });
    });

    hline(s, 1.155, 4.069, 11.024);
    vline(s, 6.082, 4.069, 3.431);
    sparkle(s, 5.939, 3.929, 0.285);
    T(s, 'Factors Influencing Real Estate Market', 1.592, 4.55, 4.51, 1.111, { fontFace: FB, fontSize: 30 });
    T(s, L_FULL, 1.592, 5.965, 3.608, 0.606, { lineSpacingMultiple: 1.5 });
    photo(s, 6.752, 4.436, 4.984, 2.227);
    icon(s, 11.21, 4.576, 0.356, BLACK, 'flame', WHITE);
  },

  /* 8 — Valuation Methods */
  (s) => {
    nav(s);
    card(s, [1.149, 1.781, 2.41, 1.656], { x: 1.29, y: 2.288, w: 1.5, t: 'Activities', size: 14 },
      { x: 1.29, y: 2.655, w: 2.068, t: L_VAR }, { dark: 1 });
    card(s, [3.847, 1.781, 2.408, 1.656], { x: 3.986, y: 2.288, w: 1.5, t: 'Customers', size: 14 },
      { x: 3.986, y: 2.655, w: 2.069, t: L_VAR });
    icon(s, 2.882, 1.974, 0.486, WHITE, 'hourglass', BLACK);
    icon(s, 5.538, 1.944, 0.486, WHITE, 'hand', BLACK);

    vline(s, 6.542, 1.111, 6.389);
    hline(s, 1.155, 3.941, 5.387);
    sparkle(s, 6.401, 3.799, 0.285);

    T(s, 'Real Estate Valuation Methods', 7.061, 1.667, 4.389, 1.111, { fontFace: FB, fontSize: 30 });
    T(s, L_FULL, 7.082, 3.031, 3.53, 0.604, { lineSpacingMultiple: 1.5 });
    photo(s, 7.061, 4.236, 4.843, 2.663);
    icon(s, 11.363, 4.352, 0.355, LIME, 'flame', BLACK);

    photo(s, 1.165, 4.353, 2.374, 2.605);
    icon(s, 1.318, 4.571, 0.355, LIME, 'cdot', BLACK);
    T(s, 'The Best Valuation', 3.847, 5.056, 2.214, 0.37, { fontFace: FB, fontSize: 16 });
    T(s, L_THERE, 3.847, 5.59, 2.207, 0.858, { lineSpacingMultiple: 1.5 });
  },

  /* 9 — Financing Options */
  (s) => {
    nav(s);
    T(s, 'Real Estate Financing Options', 1.059, 1.974, 4.141, 1.111, { fontFace: FB, fontSize: 30 });
    T(s, L_VAR, 5.2, 2.479, 2.562, 0.58, { lineSpacingMultiple: 1.5 });

    icon(s, 7.925, 2.252, 0.303, BLACK, 'cdot', WHITE);
    T(s, '1.9k', 8.389, 2.184, 0.802, 0.405, { fontFace: FB, fontSize: 18 });
    T(s, TWO_LINE, 8.403, 2.566, 1.448, 0.505, { fontSize: 8, lineSpacingMultiple: 1.5 });
    icon(s, 10.191, 2.273, 0.303, LIME, 'cdot', BLACK);
    T(s, 'The Details.', 10.549, 2.259, 1.2, 0.302, { fontFace: FB, fontSize: 12 });
    T(s, TWO_LINE, 10.549, 2.575, 1.448, 0.505, { fontSize: 8, lineSpacingMultiple: 1.5 });

    hline(s, 1.155, 3.491, 11.024);
    vline(s, 6.542, 3.491, 4.009);
    sparkle(s, 6.399, 3.349, 0.285);

    card(s, [1.155, 3.861, 2.358, 1.389], { x: 1.363, y: 4.137, w: 1.368, t: 'Bank Loans' },
      { x: 1.363, y: 4.469, w: 1.941, t: L_SHORT }, { dark: 1 });
    card(s, [3.762, 3.872, 2.358, 1.389], { x: 3.988, y: 4.148, w: 1.906, t: 'Invoice Financing' },
      { x: 3.988, y: 4.479, w: 1.941, t: L_SHORT });
    card(s, [1.163, 5.509, 2.358, 1.389], { x: 1.372, y: 5.785, w: 1.778, t: 'Crowdfunding' },
      { x: 1.372, y: 6.116, w: 1.941, t: L_SHORT });
    card(s, [3.771, 5.519, 2.358, 1.389], { x: 3.892, y: 5.795, w: 2.236, t: 'Equipment Financing' },
      { x: 3.892, y: 6.127, w: 1.941, t: L_SHORT }, { dark: 1 });

    T(s, 'The Best Financing', 6.898, 4.814, 1.562, 0.774, { fontFace: FB, fontSize: 20 });
    T(s, L_THERE, 6.898, 5.767, 2.208, 0.858, { lineSpacingMultiple: 1.5 });
    photo(s, 9.309, 3.861, 2.87, 3.038);
    icon(s, 11.66, 4.038, 0.356, BLACK, 'flame', WHITE);
  },

  /* 10 — Development Process */
  (s) => {
    nav(s);
    T(s, 'Real Estate Development Process', 1.059, 1.674, 5.061, 1.111, { fontFace: FB, fontSize: 30 });
    T(s, L_FULL, 1.069, 2.998, 3.766, 0.606, { lineSpacingMultiple: 1.5 });
    photo(s, 6.999, 1.785, 5.18, 1.998);
    icon(s, 11.727, 1.941, 0.355, LIME, 'flame', NAVY);

    hline(s, 1.155, 4.12, 11.024);
    vline(s, 6.542, 4.12, 3.38);
    sparkle(s, 6.399, 3.979, 0.285);

    card(s, [1.155, 4.49, 2.358, 1.389], { x: 1.38, y: 4.878, w: 1.368, t: 'Maintenance' },
      { x: 1.38, y: 5.208, w: 1.906, t: L_SHORT }, { dark: 1 });
    card(s, [3.762, 4.502, 2.358, 1.387], { x: 3.988, y: 4.889, w: 1.906, t: 'Planning' },
      { x: 3.988, y: 5.22, w: 1.906, t: L_SHORT });
    icon(s, 2.882, 4.641, 0.395, WHITE, 'fingerprint', BLACK);
    icon(s, 5.455, 4.628, 0.389, WHITE, 'search', BLACK);
    T(s, L_VARS, 1.059, 6.16, 3.243, 0.606, { lineSpacingMultiple: 1.5 });
    button(s, 4.826, 6.408, true, 'Explore Now', 4.979, 6.451, 0.986);

    photo(s, 6.999, 4.523, 2.423, 1.345);
    icon(s, 7.125, 4.599, 0.356, LIME, 'flame', BLACK);
    photo(s, 9.763, 4.523, 2.423, 1.345);
    icon(s, 9.896, 4.599, 0.356, BLACK, 'flame', WHITE);
    icon(s, 7.003, 6.062, 0.303, LIME, 'cdot', BLACK);
    T(s, '1.9k', 7.467, 5.993, 0.802, 0.405, { fontFace: FB, fontSize: 18 });
    T(s, TWO_LINE, 7.481, 6.375, 1.608, 0.505, { fontSize: 8, lineSpacingMultiple: 1.5 });
    icon(s, 9.762, 6.082, 0.304, BLACK, 'cdot', WHITE);
    T(s, 'The Details.', 10.12, 6.068, 1.297, 0.302, { fontFace: FB, fontSize: 12 });
    T(s, TWO_LINE, 10.12, 6.384, 1.608, 0.505, { fontSize: 8, lineSpacingMultiple: 1.5 });
  },

  /* 11 — Trends (doughnut chart) */
  (s, pres) => {
    nav(s);
    [['Economic Growth', 1.771, 2.184, 2.507, 1, 'hourglass', 1.92, 0.488],
      ['NRI Investments', 3.561, 3.974, 4.297, 0, 'target', 3.707, 0.496],
      ['Loan Interst Rates', 5.352, 5.766, 6.087, 1, 'atom', 5.615, 0.435]].forEach(
      ([t, by, ty, cy, dark, glyph, iy, id]) => {
        panel(s, 1.155, by, 3.637, 1.547, dark ? BLACK : LIME);
        T(s, t, 1.62, ty, 1.964, 0.337, { fontFace: FB, fontSize: 14, color: dark ? WHITE : BLACK });
        T(s, L_MANY, 1.62, cy, 2.705, 0.484, { fontSize: 8, color: dark ? WHITE : BLACK, lineSpacingMultiple: 1.5 });
        icon(s, 3.97, iy, id, WHITE, glyph, BLACK);
      });

    vline(s, 5.441, 1.111, 6.389);
    hline(s, 5.441, 3.861, 6.738);
    sparkle(s, 5.299, 3.719, 0.285);

    T(s, 'Real Estate Trends', 6.556, 1.696, 3.3, 0.505, { fontFace: FB, fontSize: 24 });
    T(s, L_FULL, 6.556, 2.321, 3.406, 0.606, { lineSpacingMultiple: 1.5 });
    button(s, 6.667, 3.247, true, 'Explore Now', 6.821, 3.292, 0.984);
    button(s, 8.115, 3.252, false, 'Learn Now', 8.297, 3.297, 0.929, 19765);

    panel(s, 6.667, 4.222, 2.196, 1.259, LIME);
    [[6.878, 4.795, 0.431, BLACK], [7.054, 4.46, 0.766, WHITE], [7.231, 4.885, 0.34, BLACK],
      [7.406, 4.613, 0.613, WHITE], [7.583, 4.844, 0.382, BLACK]].forEach(([x, y, h, c]) => {
      panel(s, x, y, 0.114, h, c, 27391);
    });
    T(s, '2K+', 7.865, 4.58, 0.809, 0.505, { fontFace: FB, fontSize: 24, align: 'center' });
    T(s, 'Total Courses', 7.778, 5.021, 0.983, 0.236, { fontSize: 8, align: 'center' });

    panel(s, 6.667, 5.644, 2.196, 1.259, BLACK);
    T(s, '260 M', 7.01, 5.793, 1.509, 0.571, { fontFace: FB, fontSize: 28, color: WHITE, align: 'center' });
    T(s, L_SHORT, 6.844, 6.248, 1.842, 0.486, { fontSize: 8, color: WHITE, align: 'center', lineSpacingMultiple: 1.5 });

    panel(s, 9.109, 4.227, 2.196, 2.672, LIME);
    s.addChart(pres.ChartType.doughnut, [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: [66, 33] }], {
      x: 9.075, y: 4.369, w: 2.349, h: 1.566,
      holeSize: 75, showLegend: false, showValue: false, showTitle: false,
      chartColors: [WHITE, 'D9D9D9'], dataBorder: { pt: 0, color: WHITE },
    });
    T(s, '88%', 9.715, 5.056, 1.061, 0.269, { fontFace: FB, fontSize: 16, align: 'center', margin: 0 });
    T(s, 'On Progress', 9.583, 6.038, 1.248, 0.168, { fontFace: FB, fontSize: 10, align: 'center', margin: 0 });
    T(s, 'passages of Lorem Ipsum available, but the majority.', 9.196, 6.22, 2.023, 0.484,
      { fontSize: 8, align: 'center', lineSpacingMultiple: 1.5 });
  },

  /* 12 — Challenges */
  (s) => {
    nav(s);
    photo(s, 1.155, 1.781, 2.957, 5.118);
    icon(s, 1.332, 1.934, 0.356, BLACK, 'cdot', WHITE);

    vline(s, 4.486, 1.111, 6.389);
    T(s, 'Real Estate Challenges', 4.906, 1.922, 4.891, 0.606, { fontFace: FB, fontSize: 30 });
    T(s, L_MANY, 4.906, 2.719, 3.285, 0.606, { lineSpacingMultiple: 1.5 });
    button(s, 8.658, 2.898, true, 'Explore Now', 8.812, 2.943, 0.984);
    button(s, 10.106, 2.903, false, 'Learn Now', 10.288, 2.946, 0.929, 19765);

    hline(s, 4.486, 3.75, 7.859);
    vline(s, 8.333, 3.75, 3.75);
    sparkle(s, 4.344, 3.608, 0.285);
    sparkle(s, 8.191, 3.608, 0.285);

    card(s, [4.903, 3.95, 3.111, 1.389], { x: 5.13, y: 4.24, w: 1.368, t: 'Support', size: 14 },
      { x: 5.13, y: 4.571, w: 2.241, t: L_MANY_S });
    card(s, [4.903, 5.552, 3.111, 1.389], { x: 5.13, y: 5.842, w: 1.368, t: 'Proccess', size: 14 },
      { x: 5.13, y: 6.174, w: 2.241, t: L_MANY_S }, { dark: 1 });
    card(s, [8.658, 3.908, 3.111, 1.389], { x: 8.885, y: 4.198, w: 1.368, t: 'Productivity', size: 14 },
      { x: 8.885, y: 4.53, w: 2.247, t: L_MANY_S }, { dark: 1 });
    card(s, [8.658, 5.51, 3.111, 1.389], { x: 8.885, y: 5.8, w: 1.368, t: 'Finandal', size: 14 },
      { x: 8.885, y: 6.132, w: 2.24, t: L_MANY_S });
    icon(s, 7.248, 4.122, 0.42, WHITE, 'flame', BLACK);
    icon(s, 7.248, 5.74, 0.389, WHITE, 'search', BLACK);
    icon(s, 10.964, 4.12, 0.448, WHITE, 'handshake', BLACK);
    icon(s, 11.01, 5.668, 0.421, WHITE, 'badge', BLACK);
  },

  /* 13 — Opportunities */
  (s) => {
    nav(s);
    T(s, 'Real Estate Opportunities', 1.059, 1.837, 5.431, 0.606, { fontFace: FB, fontSize: 30 });
    T(s, L_MANY, 1.113, 2.691, 3.243, 0.606, { lineSpacingMultiple: 1.5 });
    button(s, 4.788, 2.938, true, 'Explore Now', 4.941, 2.983, 0.986);

    card(s, [1.155, 3.75, 2.358, 1.389], { x: 1.434, y: 4.026, w: 1.368, t: 'Motivational' },
      { x: 1.434, y: 4.358, w: 1.799, t: L_SHORT }, { dark: 1 });
    card(s, [3.762, 3.76, 2.358, 1.389], { x: 3.988, y: 4.036, w: 1.906, t: 'Action Oriented' },
      { x: 3.988, y: 4.368, w: 1.799, t: L_SHORT });
    card(s, [1.155, 5.5, 2.358, 1.389], { x: 1.434, y: 5.776, w: 1.799, t: 'Discover Insights' },
      { x: 1.434, y: 6.108, w: 1.799, t: L_SHORT });
    card(s, [3.762, 5.51, 2.358, 1.389], { x: 3.894, y: 5.786, w: 2.094, t: 'Embedded Coaching' },
      { x: 3.894, y: 6.118, w: 1.799, t: L_SHORT }, { dark: 1 });

    vline(s, 7.073, 1.111, 6.389);
    sparkle(s, 6.931, 0.969, 0.285);
    photo(s, 7.657, 1.781, 4.522, 2.346);
    icon(s, 11.708, 1.934, 0.356, LIME, 'flame', BLACK);
    photo(s, 7.657, 4.554, 4.522, 2.346);
    icon(s, 7.849, 4.795, 0.355, BLACK, 'cdot', WHITE);
  },

  /* 14 — Case Studies */
  (s) => {
    nav(s);
    card(s, [1.155, 1.781, 2.358, 1.389], { x: 1.451, y: 2.057, w: 1.368, t: 'Motivational' },
      { x: 1.451, y: 2.389, w: 1.764, t: L_SHORT });
    card(s, [3.762, 1.792, 2.358, 1.389], { x: 3.988, y: 2.068, w: 1.906, t: 'Action Oriented' },
      { x: 3.988, y: 2.399, w: 1.764, t: L_SHORT }, { dark: 1 });
    card(s, [1.155, 3.531, 2.358, 1.389], { x: 1.434, y: 3.807, w: 1.799, t: 'Discover Insights' },
      { x: 1.434, y: 4.139, w: 1.764, t: L_SHORT }, { dark: 1 });
    card(s, [3.762, 3.542, 2.358, 1.389], { x: 3.894, y: 3.818, w: 2.094, t: 'Embedded Coaching' },
      { x: 3.894, y: 4.149, w: 1.764, t: L_SHORT });

    vline(s, 6.611, 1.111, 6.389);
    hline(s, 6.611, 3.528, 5.568);
    hline(s, 1.155, 5.292, 5.457);
    sparkle(s, 6.469, 3.385, 0.285);
    sparkle(s, 6.469, 5.158, 0.285);

    T(s, 'Case Studies', 7.139, 1.688, 3.069, 0.606, { fontFace: FB, fontSize: 30 });
    T(s, L_FULL, 7.139, 2.446, 3.613, 0.606, { lineSpacingMultiple: 1.5 });
    photo(s, 7.139, 4.056, 5.04, 2.844);
    icon(s, 11.635, 4.253, 0.356, LIME, 'flame', BLACK);

    panel(s, 1.155, 5.618, 2.582, 1.273, LIME);
    icon(s, 1.389, 5.964, 0.61, WHITE, 'eye', BLACK);
    T(s, '274%', 2.156, 5.799, 1.155, 0.505, { fontFace: FB, fontSize: 24 });
    T(s, 'Sed utThere are many varia tions of passages.', 2.156, 6.255, 1.552, 0.484,
      { fontSize: 8, lineSpacingMultiple: 1.5 });
    icon(s, 4.201, 5.844, 0.303, BLACK, 'cdot', WHITE);
    T(s, '1.9k', 4.665, 5.776, 0.8, 0.405, { fontFace: FB, fontSize: 18 });
    T(s, TWO_LINE, 4.679, 6.158, 1.448, 0.505, { fontSize: 8, lineSpacingMultiple: 1.5 });
  },

  /* 15 — Legal Considerations */
  (s) => {
    nav(s);
    T(s, 'Legal Considerations', 1.059, 1.953, 3.302, 1.076, { fontFace: FB, fontSize: 29 });
    T(s, L_MANY, 4.743, 2.384, 3.241, 0.606, { lineSpacingMultiple: 1.5 });
    panel(s, 8.571, 2.036, 1.476, 0.953, BLACK);
    stat(s, 8.773, 2.205, '2.5k', 2.584, WHITE);
    panel(s, 10.248, 2.049, 1.476, 0.953, LIME);
    stat(s, 10.45, 2.21, '15.0 M', 2.604);

    hline(s, 1.155, 3.292, 11.024);
    vline(s, 4.023, 3.292, 4.208);
    vline(s, 7.503, 3.292, 4.208);
    sparkle(s, 3.88, 3.149, 0.285);
    sparkle(s, 7.361, 3.149, 0.285);

    card(s, [1.155, 3.851, 2.358, 1.389], { x: 1.505, y: 4.127, w: 1.656, t: 'Due Diligence' },
      { x: 1.505, y: 4.458, w: 1.799, t: L_SHORT }, { dark: 1 });
    card(s, [1.155, 5.601, 2.358, 1.389], { x: 1.434, y: 5.877, w: 1.799, t: 'Legal Assitance' },
      { x: 1.434, y: 6.208, w: 1.799, t: L_SHORT });
    card(s, [4.545, 3.868, 2.358, 1.389], { x: 4.896, y: 4.146, w: 1.37, t: 'Insurance' },
      { x: 4.896, y: 4.476, w: 1.799, t: L_SHORT });
    card(s, [4.545, 5.618, 2.358, 1.389], { x: 4.825, y: 5.896, w: 1.799, t: 'Requirements' },
      { x: 4.825, y: 6.226, w: 1.799, t: L_SHORT }, { dark: 1 });

    photo(s, 8.104, 3.75, 4.074, 3.149);
    icon(s, 11.635, 3.95, 0.356, LIME, 'cdot', BLACK);
  },

  /* 16 — Real Estate Ethics */
  (s) => {
    nav(s);
    photo(s, 1.176, 1.785, 4.796, 3.271);
    icon(s, 1.424, 1.997, 0.355, BLACK, 'flame', WHITE);
    T(s, 'Real Estate Ethics', 1.036, 5.382, 3.882, 0.606, { fontFace: FB, fontSize: 30 });
    T(s, L_FULL, 1.036, 6.148, 3.755, 0.606, { lineSpacingMultiple: 1.5 });

    vline(s, 6.448, 1.111, 6.389);
    hline(s, 6.448, 5.281, 6.885);
    sparkle(s, 6.306, 5.139, 0.285);

    card(s, [7.214, 1.781, 2.358, 1.389], { x: 7.535, y: 2.068, w: 1.368, t: 'Values' },
      { x: 7.535, y: 2.399, w: 1.715, t: L_SHORT }, { dark: 1 });
    card(s, [9.821, 1.792, 2.358, 1.389], { x: 10.047, y: 2.078, w: 1.906, t: 'Rules Regulations' },
      { x: 10.047, y: 2.41, w: 1.715, t: L_SHORT });
    card(s, [7.214, 3.531, 2.358, 1.389], { x: 7.493, y: 3.818, w: 1.799, t: 'Ethical Practices' },
      { x: 7.493, y: 4.149, w: 1.715, t: L_SHORT });
    card(s, [9.821, 3.542, 2.358, 1.389], { x: 10.142, y: 3.828, w: 1.715, t: 'Moral Principies' },
      { x: 10.142, y: 4.16, w: 1.715, t: L_SHORT }, { dark: 1 });

    panel(s, 7.214, 5.594, 2.358, 1.306, BLACK);
    T(s, '260.01 M', 7.637, 5.804, 1.51, 0.405,
      { fontFace: FS, fontSize: 24, color: WHITE, align: 'center', margin: 0, wrap: false });
    T(s, 'variations of passages Ipsum available, but the majority.', 7.576, 6.212, 1.632, 0.477,
      { fontFace: FSR, fontSize: 8, color: WHITE, align: 'center', lineSpacingMultiple: 1.5 });
    icon(s, 10.229, 5.946, 0.303, LIME, 'cdot', BLACK);
    T(s, 'The Details.', 10.587, 5.932, 1.318, 0.302, { fontFace: FB, fontSize: 12 });
    T(s, TWO_LINE, 10.587, 6.248, 1.448, 0.505, { fontSize: 8, lineSpacingMultiple: 1.5 });
  },

  /* 17 — Conclusion */
  (s) => {
    nav(s);
    photo(s, 1.155, 1.905, 2.608, 1.305);
    icon(s, 1.283, 2.005, 0.356, LIME, 'flame', BLACK);
    T(s, 'The Best Conclusion', 4.019, 2.111, 1.851, 0.306, { fontFace: FB, fontSize: 12 });
    T(s, 'There are many variations of passages of Lorem Ipsum available.', 4.019, 2.464, 2.062, 0.484,
      { fontSize: 8, lineSpacingMultiple: 1.5 });
    hline(s, 1.141, 3.458, 5.807);

    card(s, [1.155, 3.75, 2.358, 1.389], { x: 1.46, y: 4.026, w: 1.368, t: 'Motivational' },
      { x: 1.46, y: 4.358, w: 1.747, t: L_SHORT }, { dark: 1 });
    card(s, [3.762, 3.76, 2.358, 1.389], { x: 3.988, y: 4.036, w: 1.906, t: 'Action Oriented' },
      { x: 3.988, y: 4.368, w: 1.747, t: L_SHORT });
    card(s, [1.155, 5.509, 2.358, 1.389], { x: 1.434, y: 5.785, w: 1.799, t: 'Discover Insights' },
      { x: 1.434, y: 6.116, w: 1.747, t: L_SHORT });
    card(s, [3.762, 5.51, 2.358, 1.389], { x: 3.894, y: 5.786, w: 2.094, t: 'Embedded Coaching' },
      { x: 3.894, y: 6.118, w: 1.747, t: L_SHORT }, { dark: 1 });

    vline(s, 6.948, 1.111, 6.389);
    photo(s, 7.556, 1.785, 4.345, 2.046);
    icon(s, 11.394, 1.984, 0.355, LIME, 'cdot', BLACK);
    hline(s, 6.948, 4.181, 5.231);
    T(s, 'Conclusion', 7.458, 4.365, 3.356, 0.774, { fontFace: FB, fontSize: 40 });
    T(s, L_FULL, 7.458, 5.316, 3.568, 0.606, { lineSpacingMultiple: 1.5 });
    button(s, 7.557, 6.314, false, 'Explore Now', 7.71, 6.359, 0.986);
    button(s, 9.003, 6.319, true, 'Learn Now', 9.186, 6.363, 0.929, 19765);
  },

  /* 18 — Thank You */
  (s) => {
    nav(s);
    T(s, 'Thank You', 1.059, 2.903, 3.733, 0.91, { fontFace: FB, fontSize: 48 });
    T(s, L_MANY, 1.062, 4.205, 3.241, 0.606, { lineSpacingMultiple: 1.5 });
    button(s, 1.179, 5.399, true, 'Explore Now', 1.333, 5.443, 0.984);
    button(s, 2.627, 5.403, false, 'Learn Now', 2.809, 5.448, 0.929, 19765);

    photo(s, 5.037, 1.781, 3.6, 5.118);
    icon(s, 8.092, 1.931, 0.355, LIME, 'cdot', BLACK);

    [['Email :', 3.441, 1.252], ['Website:', 4.125, 1.227], ['Location:', 4.809, 1.227],
      ['Telephone:', 5.493, 1.373]].forEach(([t, y, w]) => {
      T(s, t, 8.96, y, w, 0.372, { fontFace: FB, fontSize: 16 });
    });
    T(s, 'Yourname@email.com', 10.26, 3.554, 1.814, 0.269);
    T(s, 'Yourwebsite.com', 10.26, 4.174, 1.542, 0.269);
    T(s, '24 Street Your City, Office', 10.26, 4.793, 1.328, 0.438);
    T(s, '+(423) 345-32-345', 10.26, 5.582, 1.533, 0.269);
  },
];

/* ---------------------------------------------------------------- assemble */
function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'DECK', width: 13.333, height: 7.5 });
  pres.layout = 'DECK';
  pres.author = 'Xectol';
  pres.title = 'Modern Real Estate Presentation';

  SLIDES.forEach((draw) => {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    draw(s, pres);
  });

  return pres.writeFile({ fileName: path.join(__dirname, '131f850a-07da-4896-b3df-5ab098ec465c_grok_final.pptx') });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
