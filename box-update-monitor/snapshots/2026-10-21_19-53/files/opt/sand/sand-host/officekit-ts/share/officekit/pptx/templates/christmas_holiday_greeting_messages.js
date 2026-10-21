/**
 * "Sharing the Christmas Spirit" - 20 slide deck, rebuilt with pptxgenjs.
 *
 * Slide size 13.333 x 7.5 in (16:9).  Theme fonts: Staatliches (display) / Kanit (text).
 * Raster artwork in the source deck (device mock-ups, sparkles, icons) is redrawn with
 * native shapes; the source deck's picture placeholders are empty and render as nothing,
 * so they are intentionally not reproduced.
 */
const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ palette */
// Flattened from the theme (accent3 570009 / accent4 BA0111 / accent5 C5010C / accent6 FF7F00)
// with the lumMod / lumOff modifiers the deck applies.  Two-stop gradients are
// represented by their mid tone because pptxgenjs shapes take a solid fill.
const C = {
  bgDark: '620006', // accent5 lumMod 50%  - top-left of the slide background
  bgLite: 'C10014', // accent3 lumMod 75% / lumOff 25% - bottom-right
  red: 'C5010C', // accent5, the snowflake red
  redMid: 'A10917', // 81000D -> BA0111 button / card gradient
  orange: 'FF7F00', // accent6
  orangeMid: 'E07408', // BF5F00 -> FF7F00 card gradient
  orangeLite: 'FF9933', // FFB266 -> FF7F00 card gradient
  glass: '890810', // translucent dark panels sitting on the red background
  white: 'FFFFFF',
  black: '000000',
  cream: 'F8EFDC', // sparkle stars
  salmon: 'EB2626', // accent2 lumMod 50% / lumOff 50%
  ink: '0A0A0A', // device mock-up bodies
};

const FONT_DISPLAY = 'STaatliches'; // theme major latin
const FONT_TEXT = 'Kanit'; // theme minor latin
const SLIDE_W = 13.333;
const SLIDE_H = 7.5;
const NO_LINE = { type: 'none' };

/* ------------------------------------------------------------------ helpers */
const mix = (a, b, t) =>
  [0, 2, 4]
    .map((i) => {
      const v = Math.round(parseInt(a.substr(i, 2), 16) * (1 - t) + parseInt(b.substr(i, 2), 16) * t);
      return v.toString(16).padStart(2, '0').toUpperCase();
    })
    .join('');

/** Solid shape with the deck's default "no outline". */
function shape(slide, kind, opts) {
  slide.addShape(kind, Object.assign({ line: NO_LINE }, opts));
}

/** roundRect where `adj` is the OOXML adjust value (fraction of the short side). */
function card(slide, x, y, w, h, color, adj = 0.16667) {
  shape(slide, 'roundRect', { x, y, w, h, fill: { color }, rectRadius: adj * Math.min(w, h) });
}

/** Straight line between two points (pptxgenjs encodes direction with flipH/flipV). */
function line(slide, x1, y1, x2, y2, opts) {
  shape(slide, 'line', {
    x: Math.min(x1, x2),
    y: Math.min(y1, y2),
    w: Math.abs(x2 - x1),
    h: Math.abs(y2 - y1),
    flipH: x2 < x1,
    flipV: y2 < y1,
    line: opts,
  });
}

/** custGeom polygon; `pts` are 0..1 fractions of the w x h box. [x, y] or [x, y, 1] to start a subpath. */
function poly(slide, x, y, w, h, pts, opts) {
  const points = pts.map((p) => (p.close ? { close: true } : { x: p[0] * w, y: p[1] * h, moveTo: p[2] === 1 }));
  shape(slide, 'custGeom', Object.assign({ x, y, w, h, points }, opts));
}

/* --------------------------------------------------------------- text blocks */
function text(slide, content, opts) {
  slide.addText(content, Object.assign({ fontFace: FONT_TEXT, fontSize: 18, color: C.black, valign: 'top' }, opts));
}

/** The orange "Happy Merry Christmas" kicker that sits above most headlines. */
function kicker(slide, x, y) {
  text(slide, ' Happy Merry Christmas', { x, y, w: 2.131, h: 0.286, fontSize: 11, color: C.orange });
}

function heading(slide, str, x, y, w, h, opts) {
  text(slide, str, Object.assign({ x, y, w, h, fontSize: 44, fontFace: FONT_DISPLAY }, opts));
}

function paragraph(slide, str, x, y, w, h, opts) {
  text(slide, str, Object.assign({ x, y, w, h, fontSize: 10.5, lineSpacingMultiple: 1.5 }, opts));
}

/** Pill / rounded button: filled shape plus a centred caption box. */
function button(slide, o) {
  const lw = o.labelW || 1.368;
  const lh = o.labelH || 0.342;
  card(slide, o.x, o.y, o.w, o.h, o.fill, o.adj === undefined ? 0.5 : o.adj);
  text(slide, o.label, {
    x: o.x + (o.w - lw) / 2,
    y: o.y + (o.h - lh) / 2,
    w: lw,
    h: lh,
    fontSize: o.fontSize || 14,
    color: C.white,
    align: 'center',
  });
}

/* ------------------------------------------------------------- backgrounds */
/**
 * The six dark slides use a 45-degree linear gradient (620006 -> C10014).
 * pptxgenjs has no gradient fill, so it is banded into strips drawn
 * perpendicular to the gradient axis.
 */
function gradientBackground(slide, steps = 48) {
  const span = (SLIDE_W + SLIDE_H) / Math.SQRT2; // length of the slide along the 45 deg axis
  const band = span / steps;
  for (let i = 0; i < steps; i++) {
    const along = ((i + 0.5) * band) / Math.SQRT2;
    const t = Math.max(0, ((i + 0.5) / steps - 0.13) / 0.87); // first stop sits at 13%
    shape(slide, 'rect', {
      x: along + (SLIDE_W - SLIDE_H) / 4 - span / 2,
      y: along - (SLIDE_W - SLIDE_H) / 4 - band / 2,
      w: span,
      h: band * 1.04, // slight overlap so the seams do not show
      rotate: -45,
      fill: { color: mix(C.bgDark, C.bgLite, Math.min(1, t)) },
    });
  }
}

/* -------------------------------------------------------------- snowflakes */
/**
 * The eight-armed shard snowflake used all over the deck.  One quadrant is
 * described below; the remaining three are 90-degree copies of it.
 * `box` is [x, y, w, h] as a fraction of the flake, `pts` the shard outline.
 */
const FLAKE_QUADRANT = [
  { box: [0.570, 0.000, 0.083, 0.188], pts: [[0.81, 0], [0, 0.44], [0.19, 1], [1, 0.561]] },
  {
    box: [0.487, 0.173, 0.149, 0.275],
    pts: [[0.904, 0.07], [0.522, 0.326], [0.432, 0], [0, 0.294], [0.181, 1], [1, 0.443]],
  },
  {
    box: [0.354, 0.243, 0.084, 0.148],
    pts: [[0.855, 0.764], [0.781, 0.61], [0.716, 0.465], [0.558, 0.367], [0.381, 0.251], [-0.015, -0.002],
      [0.169, 0.319], [0.308, 0.602], [0.677, 0.833], [0.881, 0.946], [0.985, 0.998]],
  },
  {
    // thin diagonal band with a hollow leaf punched out of it (second sub-path)
    box: [0.324, 0.192, 0.149, 0.261],
    pts: [[0.743, 0.418], [-0.008, -0.001], [0.241, 0.58], [0.992, 0.999], { close: true },
      [0.673, 0.746, 1], [0.558, 0.682], [0.314, 0.563], [0.309, 0.56], [0.309, 0.554], [0.237, 0.364],
      [0.188, 0.192], [0.424, 0.304], [0.553, 0.365], [0.679, 0.431], [0.688, 0.436], [0.688, 0.444],
      [0.724, 0.532], [0.758, 0.621], [0.818, 0.801], [0.831, 0.844], [0.78, 0.813]],
  },
];

function snowflake(slide, o) {
  const s = o.size;
  const fill = { color: o.color, transparency: o.transparency || 0 };
  FLAKE_QUADRANT.forEach(({ box: [bx, by, bw, bh], pts }) => {
    for (let q = 0; q < 4; q++) {
      const angle = (o.rotate || 0) + q * 90;
      const rad = (angle * Math.PI) / 180;
      // spin the shard's centre around the middle of the flake, then spin the shard itself
      const dx = bx + bw / 2 - 0.5;
      const dy = by + bh / 2 - 0.5;
      const cx = 0.5 + dx * Math.cos(rad) - dy * Math.sin(rad);
      const cy = 0.5 + dx * Math.sin(rad) + dy * Math.cos(rad);
      poly(slide, o.x + (cx - bw / 2) * s, o.y + (cy - bh / 2) * s, bw * s, bh * s, pts, { fill, rotate: angle });
    }
  });
}

/** The thin white line-art snowflake (8 arms, each with a pair of tip branches). */
function lineSnowflake(slide, o) {
  const stroke = { color: C.white, width: 1.4 };
  const cx = o.x + o.size / 2;
  const cy = o.y + o.size / 2;
  const r = o.size / 2;
  for (let i = 0; i < 8; i++) {
    const a = (i * Math.PI) / 4;
    const tip = [cx + r * Math.cos(a), cy + r * Math.sin(a)];
    const fork = [cx + 0.59 * r * Math.cos(a), cy + 0.59 * r * Math.sin(a)];
    line(slide, cx, cy, tip[0], tip[1], stroke);
    [-1, 1].forEach((side) => {
      const b = a + (side * Math.PI) / 4;
      line(slide, fork[0], fork[1], fork[0] + 0.41 * r * Math.cos(b), fork[1] + 0.41 * r * Math.sin(b), stroke);
    });
  }
}

/** Four-point sparkle stars scattered over the title slide. */
function sparkle(slide, x, y, w, h) {
  shape(slide, 'star4', { x, y, w, h, fill: { color: C.cream } });
}

/* ------------------------------------------------------------------- icons */
// Simple white pictograms that sit inside the orange / red tiles.
function iconTree(slide, x, y, w, h, color = C.white) {
  poly(slide, x, y, w, h * 0.82, [[0.5, 0], [1, 1], [0, 1]], { fill: { color } });
  shape(slide, 'rect', { x: x + w * 0.42, y: y + h * 0.8, w: w * 0.16, h: h * 0.2, fill: { color } });
}

function iconSnowman(slide, x, y, w, h, color = C.white) {
  shape(slide, 'ellipse', { x: x + w * 0.18, y, w: w * 0.64, h: h * 0.46, fill: { color } });
  shape(slide, 'ellipse', { x, y: y + h * 0.4, w, h: h * 0.6, fill: { color } });
}

function iconHat(slide, x, y, w, h, color = C.white) {
  poly(slide, x + w * 0.1, y, w * 0.9, h * 0.72, [[1, 0], [0.35, 0.55], [0, 1], [0.9, 1]], { fill: { color } });
  shape(slide, 'roundRect', { x, y: y + h * 0.66, w, h: h * 0.34, fill: { color }, rectRadius: h * 0.17, line: NO_LINE });
}

function iconStocking(slide, x, y, w, h, color = C.white) {
  poly(slide, x, y, w, h, [[0.28, 0], [1, 0], [1, 0.2], [0.62, 0.2], [0.62, 0.62],
    [1, 0.82], [0.95, 1], [0.2, 1], [0, 0.78], [0.28, 0.2]], { fill: { color } });
}

function iconGift(slide, x, y, w, h, color = C.white) {
  shape(slide, 'rect', { x, y: y + h * 0.3, w, h: h * 0.7, fill: { color } });
  shape(slide, 'rect', { x, y: y + h * 0.18, w, h: h * 0.16, fill: { color } });
  shape(slide, 'rect', { x: x + w * 0.42, y: y + h * 0.18, w: w * 0.16, h: h * 0.82, fill: { color: C.orangeMid } });
  shape(slide, 'ellipse', { x: x + w * 0.16, y, w: w * 0.34, h: h * 0.22, fill: { color } });
  shape(slide, 'ellipse', { x: x + w * 0.5, y, w: w * 0.34, h: h * 0.22, fill: { color } });
}

function iconPin(slide, x, y, w, h, color = C.white) {
  poly(slide, x, y, w, h, [[0.5, 1], [0.05, 0.42], [0.16, 0.12], [0.5, 0], [0.84, 0.12], [0.95, 0.42]],
    { fill: { color } });
  shape(slide, 'ellipse', { x: x + w * 0.31, y: y + h * 0.22, w: w * 0.38, h: h * 0.34, fill: { color: C.redMid } });
}

function iconLink(slide, x, y, w, h, color = C.white) {
  [-1, 1].forEach((side) => {
    shape(slide, 'roundRect', {
      x: x + w * (side < 0 ? 0.0 : 0.4),
      y: y + h * (side < 0 ? 0.4 : 0.0),
      w: w * 0.6,
      h: h * 0.6,
      rotate: 45,
      fill: { type: 'none' },
      rectRadius: Math.min(w, h) * 0.28,
      line: { color, width: 2.2 },
    });
  });
}

function iconPhone(slide, x, y, w, h, color = C.white) {
  shape(slide, 'roundRect', {
    x: x + w * 0.18, y, w: w * 0.64, h,
    fill: { type: 'none' }, rectRadius: Math.min(w, h) * 0.16, line: { color, width: 2.2 },
  });
  shape(slide, 'rect', { x: x + w * 0.36, y: y + h * 0.08, w: w * 0.28, h: h * 0.05, fill: { color } });
}

function iconHolly(slide, x, y, w, h, color = C.black) {
  poly(slide, x, y, w * 0.55, h * 0.7, [[0, 0.5], [0.45, 0], [1, 0.35], [0.55, 1]], { fill: { color } });
  poly(slide, x + w * 0.42, y, w * 0.58, h * 0.75, [[0.1, 0.15], [0.7, 0], [1, 0.6], [0.35, 1]], { fill: { color } });
  [[0.34, 0.6], [0.24, 0.78], [0.52, 0.78]].forEach(([bx, by]) =>
    shape(slide, 'ellipse', { x: x + w * bx, y: y + h * by, w: w * 0.22, h: h * 0.22, fill: { color } }));
}

/* --------------------------------------------------------- device mock-ups */
// Both mock-ups keep the measured geometry of the photographs they replace.
function tabletMockup(slide, x, y, w, h) {
  const bezel = 0.21;
  shape(slide, 'roundRect', { x, y, w, h, fill: { color: C.ink }, rectRadius: 0.3 });
  shape(slide, 'rect', {
    x: x + bezel, y: y + 0.39, w: w - 2 * bezel, h: h - 0.39 - 0.41, fill: { color: C.white },
  });
  shape(slide, 'ellipse', { x: x + w / 2 - 0.13, y: y + h - 0.34, w: 0.26, h: 0.26, fill: { color: '1B1B1B' } });
  shape(slide, 'ellipse', { x: x + w / 2 - 0.03, y: y + 0.19, w: 0.06, h: 0.06, fill: { color: '3A3A3A' } });
}

function monitorMockup(slide, x, y, w, h) {
  const bezel = 0.275;
  const screenH = h * 0.847; // bezel block; the stand occupies the rest
  shape(slide, 'roundRect', { x, y, w, h: screenH, fill: { color: C.ink }, rectRadius: 0.07 });
  shape(slide, 'rect', {
    x: x + bezel, y: y + bezel, w: w - 2 * bezel, h: screenH - 2 * bezel, fill: { color: C.white },
  });
  // neck + flared foot in one outline
  poly(slide, x + w / 2 - 1.28, y + screenH, 2.56, h - screenH,
    [[0.149, 0], [0.849, 0], [0.87, 0.548], [1, 1], [0, 1], [0.131, 0.548]], { fill: { color: C.ink } });
}

/** Reindeer peeking in from the bottom-left corner of the title slide. */
function reindeer(slide, x, y, w, h) {
  const brown = '8A5A12';
  [[0.30, 0.06], [0.62, 0.02]].forEach(([ax, ay]) => {
    line(slide, x + ax * w, y + (ay + 0.30) * h, x + ax * w, y + ay * h, { color: brown, width: 3 });
    line(slide, x + ax * w, y + (ay + 0.12) * h, x + (ax - 0.16) * w, y + (ay + 0.02) * h, { color: brown, width: 3 });
    line(slide, x + ax * w, y + (ay + 0.22) * h, x + (ax + 0.14) * w, y + (ay + 0.10) * h, { color: brown, width: 3 });
  });
  shape(slide, 'ellipse', { x: x + 0.22 * w, y: y + 0.30 * h, w: 0.56 * w, h: 0.44 * h, fill: { color: C.orange } });
  shape(slide, 'ellipse', { x: x + 0.10 * w, y: y + 0.62 * h, w: 0.86 * w, h: 0.52 * h, fill: { color: C.orange } });
  shape(slide, 'ellipse', { x: x + 0.55 * w, y: y + 0.50 * h, w: 0.26 * w, h: 0.20 * h, fill: { color: C.white } });
  shape(slide, 'ellipse', { x: x + 0.47 * w, y: y + 0.41 * h, w: 0.07 * w, h: 0.10 * h, fill: { color: C.black } });
  shape(slide, 'rect', { x: x + 0.24 * w, y: y + 0.71 * h, w: 0.5 * w, h: 0.05 * h, fill: { color: 'E0A21B' } });
}

/* -------------------------------------------------------- shared copy text */
const LOREM = {
  sollicitudin: 'sollicitudin eleifend at ut felis. Vivamus commodo porttitor, rutrum urna lacinia non. Vivamus',
  curabitur:
    'Lorem ipsum dolor  amet, consectetur adipiscing Curabitur hendrerit dictum. Praesent lobortis magna ' +
    'vestibulum erat sollicitudin mollis. Aenean at nibh metus. Suspendisse',
  namSed:
    'Nam sed dolor nulla. Morbi iaculis, mauris quis pretium consequat, lacus nisl viverra dui, nec commodo ' +
    'libero lectus mattis ante. Aliquam interdum lacus pharetra sapien efficitur, non porta libero vehicula. ' +
    'Nulla hendrerit',
  auctor:
    'auctor lectus Aliquam congue nibh dignissim, nec dignissim elit rutrum. Ut accumsanlectus Aliquam',
  metusFelis: 'metus felis Vivamus ultricies ante leo, id aliquam tellus lacinia quis. Nam',
  nonPorta: 'non porta libero vehicula. Nulla hendrerit dui, nec commodo libero lectus mattis ante. Aliquam',
};

/* ------------------------------------------------------------------ slides */

// 1 - title slide
function slide01(s) {
  gradientBackground(s);
  snowflake(s, { x: -1.328, y: -2.442, size: 6.478, rotate: 0, color: '7B0008', transparency: 50 });
  snowflake(s, { x: 10.243, y: 4.064, size: 5.101, rotate: 139.6, color: C.red, transparency: 25 });
  snowflake(s, { x: 10.474, y: -0.317, size: 1.589, rotate: 139.6, color: C.red, transparency: 59 });
  snowflake(s, { x: 5.973, y: 6.465, size: 1.589, rotate: 139.6, color: C.red, transparency: 59 });
  reindeer(s, -0.718, 6.615, 1.436, 1.388);
  lineSnowflake(s, { x: 5.246, y: 5.292, size: 1.015 });
  sparkle(s, 12.615, 0.554, 0.178, 0.186);
  sparkle(s, 6.492, 1.051, 0.266, 0.278);
  sparkle(s, 9.549, 6.743, 0.266, 0.278);
  sparkle(s, 0.451, 4.170, 0.178, 0.186);
  heading(s, 'Sharing the Christmas Spirit', 6.323, 2.381, 6.445, 2.3, {
    fontSize: 72, color: C.white, lineSpacingMultiple: 0.9,
  });
  button(s, {
    x: 6.426, y: 4.848, w: 2.563, h: 0.631, fill: C.orangeMid,
    label: 'Lean more', fontSize: 20, labelW: 1.595, labelH: 0.438,
  });
}

// 2 - "The Spirit of Christmas"
function slide02(s) {
  snowflake(s, { x: 11.077, y: 6.621, size: 1.666, rotate: 325.6, color: C.orangeMid });
  snowflake(s, { x: -0.724, y: -1.106, size: 2.553, rotate: 0, color: C.red });
  kicker(s, 0.961, 1.325);
  heading(s, 'The Spirit of Christmas A Message of Joy', 0.965, 1.692, 6.167, 1.717, { fontSize: 48 });
  button(s, {
    x: 0.961, y: 3.721, w: 1.258, h: 0.32, fill: C.orange,
    label: 'Learn more', fontSize: 11, labelW: 1.016, labelH: 0.286,
  });
  paragraph(s, 'sollicitudin eleifend at ut felis. Vivamus commodo porttitor, rutrum urna lacinia non. ' +
    'Vivamus commodo porttitor, ', 0.942, 4.943, 1.967, 1.404);
  text(s, 'Christmas Message', { x: 8.466, y: 5.174, w: 3.444, h: 0.438, fontSize: 20 });
  text(s, '85%', { x: 8.449, y: 5.667, w: 1.554, h: 1.01, fontSize: 54, fontFace: FONT_DISPLAY, color: C.red });
  text(s, 'Lorem ipsum dolor consectetur adipiscing elit. Lorem ipsum dolor amet, consectetur adipiscing',
    { x: 9.970, y: 5.645, w: 2.847, h: 0.873, fontSize: 12, lineSpacingMultiple: 1.3 });
}

// 3 - "Merry Vibes Only"
function slide03(s) {
  snowflake(s, { x: 11.440, y: -1.150, size: 2.553, rotate: 331.8, color: C.red });
  snowflake(s, { x: -0.453, y: 5.646, size: 1.666, rotate: 325.6, color: C.orangeMid });
  kicker(s, 6.908, 1.306);
  heading(s, 'Merry Vibes Only A Christmas Message of Positivity', 6.908, 1.659, 5.674, 2.121);
  paragraph(s, 'urabitur hendrerit. Praesent lobortis magna vestibulum erat sollicitudin mollis. ' +
    'Aenean metus. Suspendisse', 1.467, 6.174, 4.066, 0.609);
}

// 4 - "The Light That Never Fades"
function slide04(s) {
  snowflake(s, { x: -0.226, y: -1.109, size: 2.340, rotate: 277.7, color: C.red });
  snowflake(s, { x: 11.171, y: 5.645, size: 1.666, rotate: 325.6, color: C.orangeMid });
  kicker(s, 1.283, 1.436);
  heading(s, 'The Light That Never Fades A Christmas', 1.283, 1.789, 4.1, 2.322);
  paragraph(s, LOREM.curabitur + ' dolor  amet, consectetur adipiscing Curabitur ', 1.283, 4.110, 4.233, 1.139);
  button(s, { x: 1.283, y: 5.603, w: 1.778, h: 0.46, fill: C.redMid, label: 'Lean more' });
}

// 5 - "The Heart of Christmas" with three stacked cards
function slide05(s) {
  snowflake(s, { x: -0.018, y: 6.526, size: 1.666, rotate: 325.6, color: C.orangeMid });
  snowflake(s, { x: 12.357, y: 0.169, size: 1.631, rotate: 321.4, color: C.red });
  const cards = [
    { y: 1.464, fill: C.redMid, title: 'Christmas Cheer', titleW: 2.242 },
    { y: 3.083, fill: C.orangeMid, title: 'Christmas Message', titleW: 2.457 },
    { y: 4.703, fill: C.orangeLite, title: 'Finding Peace', titleW: 2.242 },
  ];
  cards.forEach((c) => {
    card(s, 0.720, c.y, 3.267, 1.333, c.fill);
    text(s, c.title, { x: 1.062, y: c.y + 0.248, w: c.titleW, h: 0.404, color: C.white });
    paragraph(s, 'etur adipiscing elit. Curabitur hendrerit n lacinia dictum. Praesent lobortis ',
      1.062, c.y + 0.548, 2.583, 0.536, { fontSize: 9, color: C.white });
  });
  kicker(s, 8.711, 1.358);
  heading(s, 'The Heart of Christmas Love Hope  and Faith', 8.711, 1.697, 3.867, 2.322);
  paragraph(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Curabitur hendrerit n lacinia dictum. ' +
    'Praesent lobortis magna vestibulum erat sollicitudin mollis. Aenean at nibh metus. Suspendisse',
    8.705, 4.209, 3.636, 1.139);
  button(s, { x: 8.781, y: 5.682, w: 1.778, h: 0.46, fill: C.redMid, label: 'Lean more' });
}

// 6 - "The Meaning Behind Christmas Traditions" (dark)
function slide06(s) {
  gradientBackground(s);
  snowflake(s, { x: 9.910, y: -1.231, size: 4.134, rotate: 139.6, color: C.red, transparency: 59 });
  snowflake(s, { x: -1.391, y: 4.595, size: 5.101, rotate: 139.6, color: C.red, transparency: 61 });
  snowflake(s, { x: 4.029, y: -0.460, size: 1.589, rotate: 139.6, color: C.red, transparency: 59 });
  lineSnowflake(s, { x: 7.818, y: 6.070, size: 1.015 });
  heading(s, 'The Meaning Behind Christmas Traditions', 1.356, 1.280, 4.811, 2.524,
    { fontSize: 48, color: C.white });
  [{ y: 4.120, icon: iconHat, iy: 4.364, is: 0.346 },
   { y: 5.387, icon: iconSnowman, iy: 5.596, is: 0.415 }].forEach((row) => {
    card(s, 1.415, row.y, 0.833, 0.833, C.orangeMid);
    row.icon(s, 1.415 + (0.833 - row.is) / 2, row.iy, row.is, row.is);
    paragraph(s, LOREM.sollicitudin, 2.552, row.y + 0.113, 3.556, 0.609, { color: C.white });
  });
}

// 7 - "How Christmas Was Banned and Revived" with two stat cards
function slide07(s) {
  snowflake(s, { x: 12.336, y: 6.133, size: 1.566, rotate: 308.0, color: C.red });
  snowflake(s, { x: 3.521, y: -0.706, size: 2.191, rotate: 325.6, color: C.orangeMid });
  kicker(s, 0.776, 1.296);
  heading(s, 'How Christmas Was Banned and Revived', 0.776, 1.607, 5.092, 1.582);
  [{ x: 5.322, y: 4.090, fill: C.redMid, pct: '85%', tw: 2.115 },
   { x: 8.200, y: 4.100, fill: C.orangeMid, pct: '92%', tw: 2.141 }].forEach((c) => {
    card(s, c.x, c.y, 2.597, 2.826, c.fill, 0.0555);
    const tx = c.x + 0.317;
    text(s, 'Merry Christmas', { x: tx, y: c.y + 0.489, w: 1.657, h: 0.335, fontSize: 14, fontFace: FONT_DISPLAY, color: C.white });
    text(s, c.pct, { x: tx, y: c.y + 0.824, w: 1.326, h: 0.774, fontSize: 40, fontFace: FONT_DISPLAY, color: C.white });
    text(s, 'sollicitudin eleifend at ut Vivamus commodo porttitor, rutrum urna lacinia non. ',
      { x: tx, y: c.y + 1.561, w: c.tw, h: 0.777, fontSize: 10.5, lineSpacingMultiple: 1.3, color: C.white });
  });
  paragraph(s, 'sollicitudin eleifend at ut felis. Vivamus commodo porttitor, rutrum urna',
    11.271, 5.155, 1.374, 0.991, { fontSize: 9 });
}

// 8 - "How Christmas Music Shapes the Holiday Mood"
function slide08(s) {
  snowflake(s, { x: 11.448, y: 5.874, size: 2.191, rotate: 325.6, color: C.orangeMid });
  snowflake(s, { x: -0.564, y: 0.316, size: 1.566, rotate: 278.7, color: C.red });
  kicker(s, 7.953, 1.427);
  heading(s, 'How Christmas Music Shapes the Holiday Mood', 7.953, 1.737, 4.590, 2.322);
  paragraph(s, LOREM.curabitur, 7.953, 4.083, 4.350, 0.874);
  const tiles = [
    { x: 8.013, fill: C.redMid, icon: iconHat, ix: 8.233, iy: 5.565, iw: 0.333, ih: 0.248 },
    { x: 8.939, fill: C.redMid, icon: iconSnowman, ix: 9.180, iy: 5.534, iw: 0.282, ih: 0.311 },
    { x: 9.865, fill: C.orangeMid, icon: iconTree, ix: 10.121, iy: 5.511, iw: 0.252, ih: 0.357 },
  ];
  tiles.forEach((t) => {
    card(s, t.x, 5.309, 0.764, 0.764, t.fill, 0.0555);
    t.icon(s, t.ix, t.iy, t.iw, t.ih);
  });
}

// 9 - "The Birth of Jesus"
function slide09(s) {
  snowflake(s, { x: 11.768, y: 4.175, size: 2.599, rotate: 325.6, color: C.orangeMid });
  snowflake(s, { x: -0.446, y: -0.714, size: 1.859, rotate: 278.7, color: C.red });
  // orange panel: rounded on the right edge only
  card(s, 0, 4.583, 6.5, 2.328, C.orangeMid, 0.0555);
  shape(s, 'rect', { x: 0, y: 4.583, w: 3.25, h: 2.328, fill: { color: C.orangeMid } });
  kicker(s, 1.133, 1.192);
  heading(s, 'The Birth of Jesus A Gift to Humanity', 1.133, 1.616, 4.650, 1.582);
  button(s, { x: 1.133, y: 3.438, w: 1.778, h: 0.46, fill: C.redMid, label: 'Lean more' });
  text(s, 'What It Really Means',
    { x: 1.133, y: 5.031, w: 3.317, h: 0.505, fontSize: 24, fontFace: FONT_DISPLAY, color: C.white });
  paragraph(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Curabitur hendrerit n lacinia dictum. ' +
    'Praesent lobortis magna vestibulum erat sollicitudin mollis. Aenean at nibh metus. Suspendisse',
    1.133, 5.536, 4.650, 0.874, { color: C.white });
  text(s, 'Witnessing the Savior', { x: 7.378, y: 5.035, w: 2.380, h: 0.303, fontSize: 12 });
  text(s, '6.816 M2', { x: 7.378, y: 5.348, w: 2.656, h: 1.01, fontSize: 54, fontFace: FONT_DISPLAY, color: C.red });
  text(s, 'consectetur adipiscing Curabitur hendrerit di Praesent lobortis the magna vestibulum erat. ' +
    'consectetur adipiscing Curabitur hendrerit ',
    { x: 7.378, y: 6.342, w: 5.131, h: 0.568, fontSize: 10.5, lineSpacingMultiple: 1.3 });
}

// 10 - "Break / Slide" divider (dark)
function slide10(s) {
  gradientBackground(s);
  snowflake(s, { x: 8.293, y: 5.391, size: 3.833, rotate: 0, color: C.orangeMid });
  snowflake(s, { x: -1.866, y: 3.546, size: 4.134, rotate: 129.5, color: C.red, transparency: 59 });
  snowflake(s, { x: 10.919, y: 0.095, size: 4.134, rotate: 128.8, color: C.red, transparency: 59 });
  snowflake(s, { x: 2.022, y: 0.093, size: 1.247, rotate: 139.6, color: C.red, transparency: 59 });
  const big = { fontSize: 166, bold: true, color: C.white, fontFace: FONT_DISPLAY, lineSpacingMultiple: 0.8 };
  text(s, 'Break', Object.assign({ x: 0.661, y: 1.280, w: 5.494, h: 2.448 }, big));
  text(s, 'Slide', Object.assign({ x: 7.811, y: 1.286, w: 4.861, h: 2.448 }, big));
  text(s, 'YOUR TEXT HERE',
    { x: 0.661, y: 3.917, w: 1.539, h: 0.279, fontSize: 11, color: C.white, lineSpacingMultiple: 0.95 });
  text(s, '60,000+', {
    x: 0.669, y: 4.181, w: 2.423, h: 0.806,
    fontSize: 44, fontFace: FONT_DISPLAY, color: C.white, lineSpacingMultiple: 0.95,
  });
  text(s, 'Lorem ipsum dolor consectetur adipiscing Vivamus commodo porttitor, rutrum urna lacinia non. Vivamus',
    { x: 0.669, y: 4.986, w: 3.181, h: 0.777, fontSize: 10.5, color: C.white, lineSpacingMultiple: 1.3 });
  button(s, {
    x: 0.675, y: 6.043, w: 1.584, h: 0.413, fill: C.orangeMid,
    label: 'Lean more', fontSize: 12, labelW: 0.993, labelH: 0.303,
  });
  card(s, 10.897, 5.152, 1.782, 1.783, C.glass, 0.0555);
  iconGift(s, 11.480, 5.735, 0.615, 0.615);
  card(s, 9.704, 4.018, 1.107, 1.108, C.glass, 0.0555);
  iconGift(s, 10.057, 4.371, 0.402, 0.402);
}

// 11 - "Celebrating Progress" stats + numbered cards (dark)
function slide11(s) {
  gradientBackground(s);
  snowflake(s, { x: 2.032, y: 4.740, size: 4.134, rotate: 139.6, color: C.red, transparency: 59 });
  snowflake(s, { x: 0.270, y: 0.157, size: 1.589, rotate: 139.6, color: C.red, transparency: 59 });
  snowflake(s, { x: 8.048, y: -1.575, size: 4.134, rotate: 139.6, color: C.red, transparency: 59 });
  snowflake(s, { x: 11.883, y: 4.543, size: 1.589, rotate: 139.6, color: C.red });
  text(s, [
    { text: 'Celebrating Progress ', options: { bold: true } },
    { text: 'Inspiring Growth', options: { bold: false } },
  ], {
    x: 0.833, y: 1.083, w: 6.653, h: 1.555,
    fontSize: 48, fontFace: FONT_DISPLAY, color: C.white, lineSpacingMultiple: 0.9,
  });
  text(s, 'The Holiday Mood',
    { x: 7.951, y: 1.130, w: 2.399, h: 0.393, color: C.white, lineSpacingMultiple: 0.95 });
  text(s, 'Vivamus ultricies ante leo, id aliquam tellus lacinia quis. Nam tempor, condimentum aliquam, metus ' +
    'felis Vivamus ultricies ante leo, id aliquam tellus lacinia quis. Nam tempor, condimentum aliquam, metus felis',
    { x: 7.951, y: 1.575, w: 4.549, h: 1.015, fontSize: 11, color: C.white, lineSpacingMultiple: 1.25 });
  [{ y: 3.614, amount: '$791,642,0' }, { y: 5.280, amount: '$024,619,7' }].forEach((row) => {
    text(s, row.amount, {
      x: 0.833, y: row.y, w: 2.434, h: 0.717,
      fontSize: 40, fontFace: FONT_DISPLAY, color: C.salmon, lineSpacingMultiple: 0.9,
    });
    text(s, LOREM.metusFelis, { x: 0.833, y: row.y + 0.677, w: 4.017, h: 0.640, fontSize: 16, color: C.white });
  });
  [{ x: 5.847, fill: C.orangeMid, num: '01', title: 'Where Hope Was Born', tw: 2.603 },
   { x: 9.299, fill: C.glass, num: '02', title: 'The Shepherd\u2019s Joy', tw: 2.918 }].forEach((c) => {
    card(s, c.x, 3.744, 3.202, 2.842, c.fill, 0.12513);
    text(s, c.num, {
      x: c.x + 0.143, y: 3.981, w: 0.862, h: 0.655, lineSpacingMultiple: 0.9,
      fontSize: 36, fontFace: FONT_DISPLAY, color: C.white, align: 'center',
    });
    text(s, c.title,
      { x: c.x + 0.143, y: 5.113, w: c.tw, h: 0.438, fontSize: 20, fontFace: FONT_DISPLAY, color: C.white });
    text(s, LOREM.auctor,
      { x: c.x + 0.141, y: 5.580, w: 2.920, h: 0.784, fontSize: 11, color: C.white, lineSpacingMultiple: 1.25 });
  });
}

// 12 - "How Christmas Changed Through the Ages"
function slide12(s) {
  snowflake(s, { x: 3.352, y: 6.244, size: 2.191, rotate: 325.6, color: C.orangeMid });
  snowflake(s, { x: 12.033, y: 0.276, size: 1.566, rotate: 278.7, color: C.red });
  heading(s, 'How Christmas Changed Through the Ages', 6.667, 1.205, 5.867, 1.582);
  paragraph(s, 'Vivamus commodo porttitor arcu, vel rutrum urna lacinia non. Morbi cursus id ligula non ' +
    'efficitur. Sed iaculis purus non purus ultrices aliquet ac quis massa. ', 6.667, 2.786, 5.867, 0.609);
  text(s, '2.775 m\u00b2',
    { x: 0.800, y: 4.878, w: 2.656, h: 1.01, fontSize: 54, fontFace: FONT_DISPLAY, color: C.red });
  paragraph(s, 'egestas. mauris ullamcorper, tincidunt nisl at, tristique magna. Aenean aliquam cursus ' +
    'sagittis. Suspendisse sit amet iaculis dolor. Nullam at vestibulum velit, sed', 0.800, 5.804, 3.964, 0.874);
  card(s, 7.807, 5.208, 4.726, 1.192, C.orangeMid, 0.19464);
  text(s, '$439,229',
    { x: 7.955, y: 5.400, w: 1.558, h: 0.572, fontSize: 28, fontFace: FONT_DISPLAY, color: C.white });
  text(s, 'Cost Production ',
    { x: 7.955, y: 5.838, w: 1.628, h: 0.370, fontSize: 16, fontFace: FONT_DISPLAY, color: C.white });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. ',
    { x: 9.911, y: 5.499, w: 2.475, h: 0.611, fontSize: 12, lineSpacingMultiple: 1.3, color: C.white });
}

// 13 - "How To Keep The Spirit Of Christmas"
function slide13(s) {
  snowflake(s, { x: 10.778, y: 5.962, size: 2.191, rotate: 325.6, color: C.orangeMid });
  snowflake(s, { x: 2.746, y: -0.267, size: 1.566, rotate: 278.7, color: C.red });
  kicker(s, 1.009, 1.459);
  heading(s, 'How To Keep The Spirit Of Christmas', 1.009, 1.758, 3.257, 2.322);
  paragraph(s, LOREM.namSed, 1.009, 4.131, 3.257, 1.404);
  button(s, { x: 1.056, y: 5.794, w: 1.778, h: 0.46, fill: C.orangeMid, label: 'Lean more' });
  text(s, '67,8K', {
    x: 10.832, y: 1.745, w: 1.339, h: 0.774,
    fontSize: 40, bold: true, fontFace: FONT_DISPLAY, color: C.red,
  });
  text(s, 'Morbi iaculis, mauris quis pretium consequat, lacus nisl viverra dui, nec commodo libero',
    { x: 10.832, y: 2.467, w: 1.937, h: 0.940, fontSize: 10.5, lineSpacingMultiple: 1.2 });
  card(s, 8.488, 4.152, 4.281, 1.874, C.redMid, 0.11229);
  text(s, '$99.00.00', {
    x: 9.032, y: 4.470, w: 2.436, h: 0.707,
    fontSize: 36, bold: true, fontFace: FONT_DISPLAY, color: C.white,
  });
  text(s, 'mattis ante. Aliquam interdum pharetra sapien efficitur, non porta libero',
    { x: 9.032, y: 5.096, w: 3.193, h: 0.611, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3 });
}

// 14 - "The Music of Joy"
function slide14(s) {
  snowflake(s, { x: -0.706, y: 6.019, size: 2.191, rotate: 318.8, color: C.orangeMid });
  snowflake(s, { x: 11.347, y: -0.161, size: 1.566, rotate: 278.7, color: C.red });
  kicker(s, 7.445, 1.245);
  heading(s, 'The Music of Joy Sounds of the Season', 7.445, 1.531, 5.233, 1.582);
  text(s, 'Tooth Surgery  ',
    { x: 0.852, y: 4.441, w: 1.494, h: 0.337, fontSize: 14, fontFace: FONT_DISPLAY });
  text(s, '$457,90,0', {
    x: 0.836, y: 4.790, w: 2.331, h: 0.774,
    fontSize: 40, bold: true, fontFace: FONT_DISPLAY, color: C.red,
  });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin risus sit amet venenatis tristique' +
    ' dolor sit amet, consectetur adipiscing elit. Proin ',
    { x: 0.835, y: 5.552, w: 3.212, h: 1.007, fontSize: 10.5, lineSpacingMultiple: 1.3 });
  card(s, 4.550, 3.957, 2.896, 2.826, C.orangeMid, 0.0328);
  shape(s, 'ellipse', { x: 4.841, y: 4.334, w: 0.634, h: 0.634, fill: { color: C.white } });
  iconHolly(s, 4.990, 4.520, 0.336, 0.262);
  text(s, 'Security Services',
    { x: 4.841, y: 5.055, w: 1.860, h: 0.383, fontSize: 16, fontFace: FONT_DISPLAY, color: C.white });
  text(s, 'The Lorem ipsum dolor sit amet, adielit. Sed Lorem ipsum amet, adipiscing elit. Sed Proin risus ' +
    'sit amet Sed Lorem ipsum ',
    { x: 4.841, y: 5.399, w: 2.314, h: 1.007, fontSize: 10.5, lineSpacingMultiple: 1.3, color: C.white });
}

// 15 - "How Different Cultures Celebrate Christmas"
function slide15(s) {
  snowflake(s, { x: -1.095, y: -0.648, size: 2.191, rotate: 325.6, color: C.orangeMid });
  snowflake(s, { x: 12.156, y: 5.042, size: 1.803, rotate: 318.5, color: C.red });
  kicker(s, 0.832, 1.088);
  heading(s, 'How Different Cultures Celebrate Christmas', 0.832, 1.549, 4.892, 2.322);
  paragraph(s, LOREM.namSed + ' dui, nec commodo libero lectus mattis ante. Aliquam interdum lacus ' +
    'pharetra sapien', 0.832, 4.047, 4.609, 1.404);
  button(s, { x: 0.832, y: 5.952, w: 1.778, h: 0.46, fill: C.redMid, label: 'Lean more' });
  card(s, 10.599, 1.970, 1.902, 1.902, C.orangeMid, 0.09415);
  iconStocking(s, 11.261, 2.459, 0.571, 0.922);
}

// 16 - "Victorian Christmas"
function slide16(s) {
  snowflake(s, { x: 11.651, y: 3.291, size: 1.850, rotate: 313.3, color: C.orangeMid });
  snowflake(s, { x: -0.043, y: 0.135, size: 1.566, rotate: 278.7, color: C.red });
  // red panel: rounded on the left edge only
  card(s, 6.850, 4.667, 6.5, 2.328, C.redMid, 0.0555);
  shape(s, 'rect', { x: 10.1, y: 4.667, w: 3.25, h: 2.328, fill: { color: C.redMid } });
  kicker(s, 7.650, 1.299);
  heading(s, 'Victorian Christmas Birth of Modern Traditions', 7.650, 1.635, 4.867, 2.322);
  text(s, 'Why We Exchange Gifts',
    { x: 7.650, y: 5.183, w: 3.533, h: 0.505, fontSize: 24, fontFace: FONT_DISPLAY, color: C.white });
  paragraph(s, 'libero lectus mattis ante. Aliquam interdum lacus pharetra sapien efficitur, non porta libero ' +
    'vehicula. Nulla hendrerit dui, nec commodo libero lectus mattis ante. Aliquam interdum lacus pharetra sapien',
    7.650, 5.604, 5.112, 0.874, { color: C.white });
  text(s, '1500+',
    { x: 0.824, y: 5.259, w: 1.614, h: 0.909, fontSize: 48, fontFace: FONT_DISPLAY, color: C.orange });
  text(s, 'Customers', { x: 0.824, y: 5.998, w: 1.402, h: 0.404, fontFace: FONT_DISPLAY });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Lorem ipsum dolor sit amet, ' +
    'consectetur adipiscing elit. ', { x: 2.621, y: 5.394, w: 3.289, h: 0.873, fontSize: 12, lineSpacingMultiple: 1.3 });
}

// 17 - "Traditions That Travel" with a tablet mock-up
function slide17(s) {
  snowflake(s, { x: 6.252, y: 5.913, size: 2.145, rotate: 325.6, color: C.orangeMid });
  snowflake(s, { x: -0.102, y: -0.565, size: 1.566, rotate: 331.7, color: C.red });
  // red block behind the tablet: square on top, rounded at the bottom
  card(s, 6.593, 0, 1.946, 2.048, C.redMid, 0.11149);
  shape(s, 'rect', { x: 6.593, y: 0, w: 1.946, h: 1.024, fill: { color: C.redMid } });
  tabletMockup(s, 7.613, 0.993, 4.298, 5.954);
  kicker(s, 1.342, 1.224);
  heading(s, 'Traditions That Travel Global Christmas Stories', 1.342, 1.544, 4.568, 2.322);
  [{ y: 4.152, fill: C.orangeMid, icon: iconSnowman, ix: 1.614, iy: 4.378, iw: 0.385, ih: 0.424 },
   { y: 5.385, fill: C.redMid, icon: iconTree, ix: 1.618, iy: 5.587, iw: 0.343, ih: 0.486 }].forEach((row) => {
    card(s, 1.342, row.y, 0.891, 0.890, row.fill, 0.09415);
    row.icon(s, row.ix, row.iy, row.iw, row.ih);
    paragraph(s, LOREM.nonPorta, 2.563, row.y + 0.141, 3.700, 0.609);
  });
}

// 18 - "The Gift of Time" with a monitor mock-up
function slide18(s) {
  shape(s, 'rect', { x: 0, y: 0, w: 2.667, h: 2.917, fill: { color: C.orangeMid } });
  snowflake(s, { x: 0.254, y: 6.111, size: 1.950, rotate: 325.6, color: C.orangeMid });
  snowflake(s, { x: 8.093, y: -0.235, size: 1.566, rotate: 278.7, color: C.red });
  monitorMockup(s, 0.986, 1.340, 6.957, 5.007);
  kicker(s, 9.237, 1.596);
  heading(s, 'The Gift of Time with Loved Ones', 9.237, 1.932, 2.918, 2.322);
  paragraph(s, 'nulla. Morbi iaculis, mauris  pretium consequat, lacus nisl viverra dui, nec commodo libero ' +
    'lectus mattis ante. Aliquam interdum lacus pharetra sapien efficitur, non porta', 9.237, 4.255, 3.363, 1.139);
  button(s, { x: 9.237, y: 5.611, w: 1.778, h: 0.46, fill: C.redMid, label: 'Lean more' });
}

// 19 - "Contact Us" (dark)
function slide19(s) {
  gradientBackground(s);
  snowflake(s, { x: -0.626, y: -1.244, size: 4.134, rotate: 139.6, color: C.red, transparency: 59 });
  snowflake(s, { x: 10.139, y: 4.184, size: 4.134, rotate: 139.6, color: C.red, transparency: 7 });
  snowflake(s, { x: 9.541, y: 0.302, size: 1.589, rotate: 139.6, color: C.red, transparency: 59 });
  snowflake(s, { x: 2.456, y: 6.035, size: 1.352, rotate: 139.6, color: C.red, transparency: 59 });
  text(s, 'Contact  Us ', {
    x: 0.979, y: 1.433, w: 5.625, h: 1.456,
    fontSize: 88, bold: true, fontFace: FONT_DISPLAY, color: C.white, lineSpacingMultiple: 0.9,
  });
  const rows = [
    { y: 3.063, fill: C.redMid, icon: iconPin, label: '123 Anywhere St., Any City', w: 3.603 },
    { y: 4.259, fill: C.orangeMid, icon: iconLink, label: 'reallygreatsite.com', w: 2.739 },
    { y: 5.454, fill: C.redMid, icon: iconPhone, label: '123-456-7890', w: 2.320 },
  ];
  rows.forEach((row) => {
    shape(s, 'ellipse', { x: 1.083, y: row.y, w: 0.888, h: 0.888, fill: { color: row.fill } });
    row.icon(s, 1.083 + 0.244, row.y + 0.228, 0.4, 0.432);
    text(s, row.label, { x: 2.439, y: row.y + 0.276, w: row.w, h: 0.438, fontSize: 20, color: C.white });
  });
  [4.127, 5.367].forEach((y) => line(s, 2.201, y, 7.239, y, { color: C.white, width: 0.75, transparency: 45 }));
  card(s, 10.719, 1.344, 1.635, 1.635, C.orangeMid, 0.09415);
  iconGift(s, 11.206, 1.781, 0.661, 0.769);
  paragraph(s, 'lectus mattis ante. Aliquam interdum lacus pharetra', 11.368, 4.720, 1.490, 0.874,
    { color: C.white });
}

// 20 - closing slide (dark)
function slide20(s) {
  gradientBackground(s);
  snowflake(s, { x: -1.353, y: -1.802, size: 6.080, rotate: 139.6, color: C.red, transparency: 59 });
  snowflake(s, { x: 9.806, y: 4.682, size: 4.134, rotate: 139.6, color: C.red, transparency: 7 });
  snowflake(s, { x: 0.277, y: 5.357, size: 1.779, rotate: 139.6, color: C.red, transparency: 59 });
  snowflake(s, { x: 9.602, y: -0.286, size: 1.779, rotate: 139.6, color: C.red, transparency: 59 });
  heading(s, 'Thank You For Your Time', 7.246, 2.151, 5.617, 2.3, {
    fontSize: 72, color: C.white, lineSpacingMultiple: 0.9,
  });
  button(s, {
    x: 7.341, y: 4.718, w: 3.047, h: 0.631, fill: C.orangeMid,
    label: 'End presentation', fontSize: 20, labelW: 2.563, labelH: 0.438,
  });
}

/* -------------------------------------------------------------------- build */
const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'DECK', width: SLIDE_W, height: SLIDE_H });
pptx.layout = 'DECK';
pptx.title = 'Sharing the Christmas Spirit';
pptx.theme = { headFontFace: FONT_DISPLAY, bodyFontFace: FONT_TEXT };

BUILDERS.forEach((build) => build(pptx.addSlide()));

pptx.writeFile({ fileName: path.join(__dirname, '0cec476b-f759-48a0-bd7b-74f50b5f9ea1_grok_final.pptx') });
