/*
 * Recreation of "Education" deck (20 slides, 13.333 x 7.5 in) with pptxgenjs.
 *
 * Raster photos in the original are replaced by flat colour placeholder blocks
 * (see `imagePlaceholder`); every other element is a native pptxgenjs shape.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const C = {
  orange: 'F6511D',   // accent1
  yellow: 'FFB400',   // accent2
  blue: '00A6ED',     // accent3
  green: '7FB800',    // accent4
  rust: 'DD3B09',     // accent5
  text: '3C3C3C',     // tx1
  slate: '313C41',    // tx2
  white: 'FFFFFF',
  grey05: 'F2F2F2',   // white lumMod 95%
  grey15: 'D9D9D9',   // white lumMod 85%
  greyIcon: 'A5B4BB', // tx2 lumMod 40% lumOff 60%
  greenLt: 'C2FF3B',
  blueLt: '5BCEFF',
  yellowLt: 'FFD266',
  ph: 'E8E8E8',       // image-placeholder fill
  phText: '9A9A9A',
};

const HEAD = 'Poppins';   // major latin
const BODY = 'Lato';      // minor latin
const SYM = 'DejaVu Sans'; // font used for the glyph icons

const ACCENTS = [C.orange, C.yellow, C.blue, C.green, C.rust];

/* --------------------------------------------------------------- helpers */

// Body copy: 12 pt Lato, 150 % line spacing — the deck's default paragraph.
function body(slide, text, opts = {}) {
  slide.addText(text, Object.assign({
    fontFace: BODY, fontSize: 12, color: C.text,
    lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6],
  }, opts));
}

// Hanging-indent list where the bullet glyph carries its own colour, as in the
// reference (pptxgenjs `bullet` has no colour option, so the glyph is a run).
function bulletList(slide, items, glyph, glyphColor, opts = {}) {
  const runs = [];
  items.forEach((t, i) => {
    runs.push({ text: glyph + '  ', options: { color: glyphColor } });
    runs.push({ text: t, options: { breakLine: i < items.length - 1 } });
  });
  body(slide, runs, opts);
}

// Bold Poppins label ("Description", names, headings inside cards).
function label(slide, text, opts = {}) {
  slide.addText(text, Object.assign({
    fontFace: HEAD, fontSize: 16, bold: true, color: C.orange,
    valign: 'middle', margin: [7.2, 7.2, 3.6, 3.6],
  }, opts));
}

// Slide title: Poppins bold 44 pt, 90 % line spacing (master titleStyle).
function title(slide, text, opts = {}) {
  slide.addText(text, Object.assign({
    fontFace: HEAD, fontSize: 44, bold: true, color: C.orange,
    lineSpacingMultiple: 0.9, valign: 'top',
    x: 0.686, y: 0.762, w: 9.952, h: 0.963,
    margin: [7.2, 7.2, 3.6, 3.6],
  }, opts));
}

// Yellow rounded tab + white number in the bottom-right corner (from the master).
function slideNumber(slide, n) {
  slide.addShape('roundRect', {
    x: 12.67, y: 6.767, w: 0.81, h: 0.81, fill: { color: C.yellow },
    line: { type: 'none' }, rectRadius: 0.124,
  });
  slide.addText(String(n), {
    x: 12.806, y: 6.904, w: 0.447, h: 0.399, align: 'center', valign: 'middle',
    fontFace: BODY, fontSize: 12, color: C.white,
  });
}

/* ------------------------------------------------------------- pictograms */

// Each icon is drawn from small shapes on top of its parent disc. Coordinates
// are fractions of the disc diameter `d`, measured from its centre (cx, cy).
// `fg` is the icon colour, `bg` the disc colour used to punch holes.
function box(P, u, v, uw, uh, extra) {
  return Object.assign({
    x: P.cx + u * P.d, y: P.cy + v * P.d, w: uw * P.d, h: uh * P.d,
    fill: { color: P.fg }, line: { type: 'none' },
  }, extra || {});
}

const G = {
  // Graduation cap: diamond mortarboard over a trapezoid crown.
  cap: (s, P) => {
    s.addShape('trapezoid', box(P, -0.17, -0.02, 0.34, 0.18, { flipV: true }));
    s.addShape('diamond', box(P, -0.32, -0.22, 0.64, 0.26));
  },
  // Trophy: flared cup on a stem and base.
  trophy: (s, P) => {
    s.addShape('flowChartManualOperation', box(P, -0.19, -0.28, 0.38, 0.30));
    s.addShape('rect', box(P, -0.035, 0.0, 0.07, 0.15));
    s.addShape('rect', box(P, -0.16, 0.15, 0.32, 0.08));
  },
  star: (s, P) => s.addShape('star5', box(P, -0.25, -0.25, 0.5, 0.5)),
  heart: (s, P) => s.addShape('heart', box(P, -0.25, -0.24, 0.5, 0.48)),
  // Pencil: slanted barrel with a pointed tip.
  pencil: (s, P) => {
    s.addShape('rect', box(P, -0.055, -0.26, 0.11, 0.36, { rotate: 45 }));
    s.addShape('triangle', box(P, -0.055, 0.10, 0.11, 0.14, { rotate: 225 }));
  },
  // Magnifier: ring plus a slanted handle.
  search: (s, P) => {
    s.addShape('ellipse', box(P, -0.26, -0.28, 0.38, 0.38,
      { fill: { type: 'none' }, line: { color: P.fg, width: 2.5 } }));
    s.addShape('roundRect', box(P, 0.06, 0.06, 0.06, 0.22,
      { rectRadius: 0.02 * P.d, rotate: -45 }));
  },
  // Light bulb: glass dome over a screw base.
  bulb: (s, P) => {
    s.addShape('ellipse', box(P, -0.17, -0.28, 0.34, 0.34));
    s.addShape('rect', box(P, -0.09, 0.02, 0.18, 0.14));
    s.addShape('rect', box(P, -0.09, 0.19, 0.18, 0.06));
  },
  // Cog wheel with a hub punched out in the disc colour.
  gears: (s, P) => {
    s.addShape('gear6', box(P, -0.30, -0.30, 0.60, 0.60));
    s.addShape('ellipse', box(P, -0.09, -0.09, 0.18, 0.18, { fill: { color: P.bg } }));
  },
  // Painter's palette: disc with paint dabs and a thumb hole.
  palette: (s, P) => {
    s.addShape('ellipse', box(P, -0.28, -0.28, 0.56, 0.56));
    [[-0.13, -0.12], [0.02, -0.17], [0.16, -0.05]].forEach(([u, v]) => {
      s.addShape('ellipse', box(P, u - 0.045, v - 0.045, 0.09, 0.09, { fill: { color: P.bg } }));
    });
    s.addShape('ellipse', box(P, 0.0, 0.06, 0.20, 0.20, { fill: { color: P.bg } }));
  },
  // Dumbbell: two weights joined by a bar.
  dumbbell: (s, P) => {
    [-0.30, 0.18].forEach(u => {
      s.addShape('roundRect', box(P, u, -0.16, 0.12, 0.32, { rectRadius: 0.04 * P.d }));
    });
    s.addShape('rect', box(P, -0.20, -0.055, 0.40, 0.11));
  },
  // Rocket: teardrop nose, body, side fins and a porthole.
  rocket: (s, P) => {
    s.addShape('teardrop', box(P, -0.15, -0.30, 0.34, 0.34, { rotate: 315 }));
    s.addShape('rect', box(P, -0.10, -0.05, 0.20, 0.20));
    s.addShape('triangle', box(P, -0.28, 0.0, 0.16, 0.20, { rotate: -20 }));
    s.addShape('triangle', box(P, 0.12, 0.0, 0.16, 0.20, { rotate: 20 }));
    s.addShape('ellipse', box(P, -0.055, -0.16, 0.11, 0.11, { fill: { color: P.bg } }));
  },
  // Crossed tools with a pivot dot.
  tools: (s, P) => {
    [35, -35].forEach(rotate => {
      s.addShape('roundRect', box(P, -0.05, -0.28, 0.10, 0.56,
        { rectRadius: 0.03 * P.d, rotate }));
    });
    s.addShape('ellipse', box(P, -0.07, -0.06, 0.14, 0.14, { fill: { color: P.bg } }));
  },
  // Two-person group: round heads over rounded shoulders.
  people: (s, P) => {
    [-1, 1].forEach(side => {
      const u = side * 0.135;
      s.addShape('ellipse', box(P, u - 0.075, -0.26, 0.15, 0.15));
      s.addShape('round2SameRect', box(P, u - 0.105, -0.08, 0.21, 0.32,
        { rectRadius: 0.09 * P.d }));
    });
  },
  // Telephone handset (crescent).
  phone: (s, P) => s.addShape('moon', box(P, -0.14, -0.28, 0.36, 0.56, { rotate: 135 })),
  // Envelope: rectangle with the flap knocked out in the disc colour.
  mail: (s, P) => {
    s.addShape('rect', box(P, -0.28, -0.19, 0.56, 0.38));
    s.addShape('triangle', box(P, -0.28, -0.19, 0.56, 0.24,
      { fill: { color: P.bg }, flipV: true }));
  },
  // Globe: disc with meridian ellipse and equator line.
  globe: (s, P) => {
    s.addShape('ellipse', box(P, -0.28, -0.28, 0.56, 0.56));
    s.addShape('ellipse', box(P, -0.12, -0.28, 0.24, 0.56,
      { fill: { type: 'none' }, line: { color: P.bg, width: 1.25 } }));
    s.addShape('line', box(P, -0.28, 0, 0.56, 0,
      { fill: undefined, line: { color: P.bg, width: 1.25 } }));
  },
};

// Coloured disc with a pictogram drawn in `fg` (white unless overridden).
function icon(slide, x, y, d, color, glyph, fg) {
  slide.addShape('flowChartConnector', {
    x, y, w: d, h: d, fill: { color }, line: { type: 'none' },
  });
  if (glyph) glyph(slide, { cx: x + d / 2, cy: y + d / 2, d, fg: fg || C.white, bg: color });
}

// Programmatic stand-in for the one raster photo in the reference deck.
function imagePlaceholder(slide, x, y, w, h, opts = {}) {
  slide.addShape('rect', {
    x, y, w, h, fill: { color: opts.fill || C.ph }, line: { type: 'none' },
  });
  slide.addText(opts.label || '[image]', {
    x, y, w, h, align: 'center', valign: 'middle',
    fontFace: BODY, fontSize: 12, color: opts.labelColor || C.phText,
  });
}

/* ------------------------------------------- reusable decorative confetti */

// "Rainbow" half-ring (custGeom) — the green swoosh that recurs on many slides.
function arcSwoosh(slide, x, y, w, h, color, rotate) {
  slide.addShape('custGeom', {
    x, y, w, h, rotate: rotate || 0, fill: { color }, line: { type: 'none' },
    points: [
      { x: 0.500 * w, y: 0 },
      { curve: { type: 'cubic', x1: 0.776 * w, y1: 0, x2: w, y2: 0.448 * h }, x: w, y: h },
      { x: 0.754 * w, y: h },
      { curve: { type: 'cubic', x1: 0.754 * w, y1: 0.724 * h, x2: 0.640 * w, y2: 0.5 * h }, x: 0.5 * w, y: 0.5 * h },
      { curve: { type: 'cubic', x1: 0.360 * w, y1: 0.5 * h, x2: 0.246 * w, y2: 0.724 * h }, x: 0.246 * w, y: h },
      { x: 0, y: h },
      { curve: { type: 'cubic', x1: 0, y1: 0.448 * h, x2: 0.224 * w, y2: 0 }, x: 0.5 * w, y: 0 },
      { close: true },
    ],
  });
}

// Thick "X" mark with chevron notches — the deck's signature confetti glyph.
// Traced from the reference custGeom (unit coordinates x fractions of w/h).
const X_MARK = [
  [0.2056, 0], [0.5, 0.2763], [0.7944, 0], [1, 0.2545], [0.7384, 0.5],
  [1, 0.7455], [0.7944, 1], [0.5, 0.7237], [0.2056, 1], [0, 0.7455],
  [0.2616, 0.5], [0, 0.2545],
];

function cross(slide, x, y, w, h, color, rotate) {
  slide.addShape('custGeom', {
    x, y, w, h, rotate: rotate || 0, fill: { color }, line: { type: 'none' },
    points: X_MARK.map(([u, v]) => ({ x: u * w, y: v * h })).concat([{ close: true }]),
  });
}

// Plain multiplication sign (a preset autoshape in the reference).
function multiply(slide, x, y, w, h, color) {
  slide.addShape('mathMultiply', { x, y, w, h, fill: { color }, line: { type: 'none' } });
}

// Ring (donut).
function ring(slide, x, y, w, h, color) {
  slide.addShape('donut', { x, y, w, h, fill: { color }, line: { type: 'none' } });
}

// Division sign: bar between two dots (cover-slide confetti).
function divide(slide, x, y, w, h, color) {
  const r = 0.1703 * w, cx = x + w / 2;
  slide.addShape('ellipse', { x: cx - r, y: y + 0.6922 * h, w: 2 * r, h: 0.3078 * h, fill: { color }, line: { type: 'none' } });
  slide.addShape('rect', { x, y: y + 0.3461 * h, w, h: 0.3078 * h, fill: { color }, line: { type: 'none' } });
  slide.addShape('ellipse', { x: cx - r, y, w: 2 * r, h: 0.3078 * h, fill: { color }, line: { type: 'none' } });
}

/* ------------------------------------------------------ shared text blocks */

const LOREM_LONG =
  'PLACEHOLDER' +
  'eget est amet est plac in egestas erat imperdet sed nam at lec';
const LOREM_MED =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed non risus. Suspendisse ' +
  'lectus tortor, dignissim sit amet, adipiscing nec, ultricies sed, dolor. Cras elementum ultrices diam';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed non Suspendisse.';
const DIAM =
  'PLACEHOLDER' +
  'luctus venenatis vitae auctor eu augue lectus';

/* ================================================================ slides */

// 1 — Cover: full-bleed orange, big title, yellow banner, confetti shapes.
function slide01(pres) {
  const s = pres.addSlide();
  s.background = { color: C.orange };

  ring(s, 4.572, 6.11, 1.009, 1.004, C.green);
  cross(s, 0.819, 1.238, 0.847, 0.785, C.blue);
  arcSwoosh(s, 7.375, 1.238, 1.918, 0.943, C.green, 145.89);
  divide(s, 7.864, 5.599, 0.875, 0.969, C.blue);

  s.addShape('roundRect', {
    x: -0.513, y: 6.502, w: 5.712, h: 1.223, fill: { color: C.yellow },
    line: { type: 'none' }, rectRadius: 0.405,
  });

  s.addText('Education', {
    x: 0.674, y: 2.93, w: 6.659, h: 1.82, fontFace: HEAD, fontSize: 88, bold: true,
    color: C.white, lineSpacingMultiple: 1.0, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6],
  });
  body(s, 'Empowering Education for Every Child', {
    x: 0.743, y: 4.264, w: 5.852, h: 0.667, fontSize: 16, color: C.white,
    lineSpacingMultiple: 0.9,
  });
  s.addText('Presented by: Jhon Doe', {
    x: 0.744, y: 6.662, w: 4.526, h: 0.452, fontFace: BODY, fontSize: 16, bold: true,
    color: C.white, lineSpacingMultiple: 1.5, valign: 'middle', margin: [7.2, 7.2, 3.6, 3.6],
  });
  return s;
}

// 2 — About Us: intro paragraph plus two icon/description columns.
function slide02(pres) {
  const s = pres.addSlide();

  s.addShape('roundRect', {
    x: 8.324, y: 1.083, w: 5.75, h: 4.667, fill: { color: C.yellow },
    line: { type: 'none' }, rectRadius: 0.281,
  });

  arcSwoosh(s, -0.461, 0.726, 1.3, 0.639, C.green, 145.89);
  cross(s, 8.333, 5.75, 0.719, 0.667, C.blue);

  title(s, 'About Us', { x: 1.02, y: 1.314, w: 5.98, h: 1.436, valign: 'middle' });
  body(s,
    'PLACEHOLDER' +
    'amet consectetur adipiscing elit sed eiusmod tempor incididunt labore  dolor magna aliqua ' +
    'potenti nullam tortor vitae purus faucibus ornar volutpat tincidu vitae semper.',
    { x: 1.015, y: 2.267, w: 5.652, h: 1.273 });

  [
    { x: 1.133, dx: 1.02, color: C.orange, glyph: G.cap },
    { x: 4.235, dx: 4.129, color: C.yellow, glyph: G.trophy },
  ].forEach((col, i) => {
    icon(s, col.x, 3.785, 0.665, col.color, col.glyph);
    label(s, 'Description', { x: col.dx, y: 4.509 + i * 0.03, w: 1.607, h: 0.378, color: col.color });
    body(s, 'Pharetra massa ultrici quisan dolores magnia ipsum sit',
      { x: col.dx, y: 4.817 + i * 0.036, w: i ? 2.409 : 2.607, h: 0.667 });
  });
  return s;
}

// 3 — Our Vision: orange side panel, two stacked icon rows.
function slide03(pres) {
  const s = pres.addSlide();

  s.addShape('roundRect', {
    x: -0.052, y: -0.349, w: 3.384, h: 8.234, fill: { color: C.orange },
    line: { type: 'none' }, rectRadius: 0.518,
  });

  cross(s, 5.201, 0.918, 0.847, 0.785, C.blue);
  ring(s, 1.665, 5.901, 1.009, 1.004, C.green);

  title(s, 'Our Vision', { x: 6.693, y: 1.417, w: 5.75, h: 1.458 });
  body(s, LOREM_LONG, { x: 6.687, y: 2.111, w: 5.98, h: 0.667 });

  [
    { y: 3.227, ty: 3.107, color: C.orange, glyph: G.star },
    { y: 4.755, ty: 4.626, color: C.yellow, glyph: G.pencil },
  ].forEach(row => {
    icon(s, 6.815, row.y, 0.664, row.color, row.glyph);
    label(s, 'Description', { x: 7.573, y: row.ty, w: 2.333, h: 0.378, color: row.color });
    body(s, LOREM_MED, { x: 7.573, y: row.ty + 0.382, w: 5.002, h: 0.97, align: 'justify' });
  });
  return s;
}

// 4 — Our Mission: mirror of slide 3 with a yellow quote pill.
function slide04(pres) {
  const s = pres.addSlide();

  s.addShape('rect', { x: 9.998, y: 0, w: 3.335, h: 1.75, fill: { color: C.yellow }, line: { type: 'none' } });
  arcSwoosh(s, 11.889, 1.726, 1.3, 0.639, C.green, 145.89);

  title(s, 'Our Mission', { x: 1.02, y: 1.083, w: 5.75, h: 1.0 });
  body(s,
    'PLACEHOLDER' +
    'eget est amet est plac in egestas erat lec',
    { x: 1.024, y: 1.782, w: 5.002, h: 0.667 });

  [
    { y: 2.804, ty: 2.683, x: 1.786, color: C.orange, glyph: G.cap },
    { y: 4.473, ty: 4.351, x: 1.791, color: C.yellow, glyph: G.pencil },
  ].forEach(row => {
    icon(s, 0.998, row.y, 0.664, row.color, row.glyph);
    label(s, 'Description', { x: row.x, y: row.ty, w: 2.333, h: 0.378, color: row.color });
    body(s,
      'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed non risus. Suspendisse ' +
      'lectus tortor, dignissim sit amet, adipiscing nec, ultricies sed',
      { x: row.x, y: row.ty + 0.393, w: 3.994, h: 0.97, align: 'justify' });
  });

  s.addShape('roundRect', {
    x: 6.667, y: 5.738, w: 3.335, h: 1.012, fill: { color: C.yellow },
    line: { type: 'none' }, rectRadius: 0.169,
  });
  s.addText('“Education is the key to unlock future”', {
    x: 6.77, y: 5.844, w: 3.0, h: 0.823, align: 'center', valign: 'middle',
    fontFace: HEAD, fontSize: 16, bold: true, italic: true, color: C.white,
    lineSpacingMultiple: 1.0, margin: [7.2, 7.2, 3.6, 3.6],
  });
  return s;
}

// World map traced from the reference (slide-inch polygons, drawn light grey).
const MAP_LAND = [
  [[0.58, 1.86], [0.73, 2.03], [0.52, 2.06], [0.77, 2.13], [0.59, 2.29], [0.89, 2.44], [0.79, 2.59], [1.28, 2.31], [1.73, 2.45], [2.05, 2.9], [2.05, 2.76], [2.21, 2.73], [2.14, 3.35], [2.57, 3.94], [2.47, 3.67], [2.83, 4.15], [3.8, 4.55], [3.7, 5.09], [4.1, 5.54], [3.91, 6.79], [3.98, 7.05], [4.23, 7.1], [4.12, 6.89], [4.27, 6.45], [5.15, 5.68], [5.24, 5.41], [5.04, 5.28], [5.19, 5.2], [5.28, 5.32], [5.38, 5.08], [4.71, 4.84], [4.68, 4.66], [4.36, 4.48], [3.63, 4.53], [3.61, 4.33], [3.42, 4.29], [3.47, 4.08], [3.2, 4.2], [3.09, 3.9], [2.94, 3.94], [2.87, 3.81], [3.51, 3.71], [3.71, 3.92], [3.67, 3.68], [3.94, 3.28], [4.42, 3.09], [4.22, 2.86], [4.57, 2.8], [4.51, 2.98], [4.73, 3], [4.61, 2.71], [4.3, 2.35], [4.15, 2.51], [4.08, 2.32], [3.82, 2.24], [3.81, 2.85], [3.62, 2.91], [3.64, 2.63], [3.19, 2.43], [3.46, 2.09], [3.43, 1.85], [3.21, 1.66], [3.32, 1.51], [3.17, 1.52], [3.05, 1.92], [2.77, 1.87], [2.96, 1.78], [2.82, 1.53], [2.58, 1.61], [2.22, 1.47], [2.06, 1.64], [2.31, 1.61], [2.41, 1.85], [0.91, 1.7]],
  [[6.22, 0.9], [5.87, 0.93], [5.86, 0.81], [5.7, 0.74], [5.41, 0.73], [5.27, 0.78], [4.98, 0.78], [4.95, 0.84], [4.82, 0.83], [4.82, 0.88], [4.44, 0.89], [4.44, 0.95], [4.27, 0.96], [4.27, 0.88], [4.42, 0.84], [3.85, 0.78], [3.65, 0.86], [3.34, 0.89], [3.43, 0.95], [3.42, 1], [3.26, 0.93], [3.14, 1.04], [3.29, 1.17], [3.47, 1.15], [3.45, 1.27], [3.39, 1.3], [3.71, 1.33], [3.79, 1.3], [3.79, 1.2], [3.91, 1.14], [3.82, 1.13], [3.82, 1.08], [3.92, 1.11], [3.96, 1.05], [4.26, 0.94], [4.2, 1.01], [4.29, 1.03], [4.28, 1.11], [4, 1.19], [4.09, 1.21], [4.07, 1.27], [4.15, 1.33], [4.42, 1.33], [4.59, 1.43], [4.67, 1.57], [4.64, 1.66], [4.78, 1.66], [4.78, 1.72], [4.71, 1.72], [4.67, 1.79], [4.83, 1.77], [4.72, 2], [4.9, 2.28], [5.08, 2.35], [5.21, 2.04], [5.83, 1.77], [5.88, 1.7], [5.84, 1.58], [5.75, 1.59], [5.74, 1.54], [5.93, 1.52], [5.97, 1.36], [5.89, 1.29], [6.01, 1.27], [5.89, 1.2], [5.9, 1.14], [6.01, 1.01]],
  [[3.4, 1.57], [3.41, 1.72], [3.54, 1.76], [3.5, 1.93], [3.67, 1.94], [3.62, 1.76], [3.82, 1.75], [3.86, 1.94], [3.89, 1.84], [4, 1.88], [4.01, 1.98], [3.83, 2.12], [3.97, 2.1], [4.23, 2.25], [4.29, 2.16], [4.19, 2], [4.33, 2.06], [4.41, 1.98], [4.16, 1.87], [4.17, 1.76], [3.94, 1.7], [3.82, 1.54], [3.63, 1.52], [3.49, 1.61], [3.5, 1.52]],
  [[3.09, 1.15], [3.1, 1.2], [3.18, 1.23], [3.15, 1.29], [3.29, 1.34], [3.31, 1.44], [3.73, 1.44], [3.73, 1.4], [3.68, 1.38], [3.39, 1.39], [3.35, 1.32], [3.19, 1.28], [3.19, 1.24], [3.24, 1.23], [3.18, 1.22], [3.17, 1.16]],
  [[2.79, 1.37], [2.72, 1.37], [2.69, 1.35], [2.67, 1.29], [2.65, 1.32], [2.66, 1.33], [2.66, 1.39], [2.6, 1.4], [2.55, 1.35], [2.49, 1.35], [2.46, 1.32], [2.44, 1.35], [2.41, 1.35], [2.43, 1.36], [2.43, 1.4], [2.38, 1.41], [2.49, 1.41], [2.49, 1.45], [2.51, 1.46], [2.77, 1.41]],
];

const MAP_PINS = [
  [3.469, 1.935, C.orange], [3.594, 2.717, C.yellow], [2.833, 2.484, C.orange],
  [3.295, 3.141, C.orange], [2.787, 2.104, C.yellow], [2.531, 3.055, C.orange],
  [1.898, 2.206, C.orange], [4.08, 2.293, C.orange], [2.894, 3.731, C.yellow],
  [2.049, 2.735, C.yellow], [4.674, 4.638, C.orange], [5.06, 5.222, C.yellow],
  [4.384, 5.97, C.orange], [4.059, 5.059, C.yellow], [4.695, 5.626, C.yellow],
];

// 5 — Why Choose Us: world map with pins + two stat pills.
function slide05(pres) {
  const s = pres.addSlide();

  MAP_LAND.forEach(poly => {
    s.addShape('custGeom', {
      x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: C.grey15 }, line: { type: 'none' },
      points: poly.map(([x, y]) => ({ x, y })).concat([{ close: true }]),
    });
  });
  MAP_PINS.forEach(([x, y, color]) => {
    s.addShape('ellipse', {
      x, y, w: 0.206, h: 0.204, fill: { color },
      line: { color: C.white, width: 2.25 },
    });
  });

  cross(s, 12.335, 0.449, 0.847, 0.785, C.blue);
  arcSwoosh(s, -0.627, 6.038, 1.918, 0.943, C.green, 145.89);

  title(s, 'Why Choose Us?', { x: 6.687, y: 1.782, w: 5.648, h: 0.968, valign: 'middle' });
  body(s, [
    { text: LOREM_LONG, options: { breakLine: true } },
    { text: 'tus urna dictum amet justo donec ornare. Sed persp unde omi iste natu vol' },
  ], { x: 6.676, y: 2.552, w: 5.95, h: 0.97 });

  [
    { x: 6.667, bx: 6.852, color: C.orange, num: '756K', cap: 'Female Student' },
    { x: 9.338, bx: 9.519, color: C.yellow, num: '629K', cap: 'Male Student ' },
  ].forEach(card => {
    s.addShape('roundRect', {
      x: card.x, y: 3.742, w: 2.345, h: 1.329, fill: { color: card.color },
      line: { type: 'none' }, rectRadius: 0.222,
    });
    s.addText(card.num, {
      x: card.bx, y: 3.916, w: 1.973, h: 0.774, align: 'center', valign: 'middle',
      fontFace: HEAD, fontSize: 40, bold: true, color: C.white, margin: [7.2, 7.2, 3.6, 3.6],
    });
    s.addText(card.cap, {
      x: card.bx, y: 4.561, w: 1.973, h: 0.337, align: 'center', valign: 'middle',
      fontFace: HEAD, fontSize: 14, bold: true, color: C.white, margin: [7.2, 7.2, 3.6, 3.6],
    });
  });

  body(s, LOREM_LONG.replace(/lec$/, 'lec.'), { x: 6.667, y: 5.318, w: 5.96, h: 0.667 });
  return s;
}

// 6 — Meet Our Educators: four portrait cards with coloured caption blocks.
const EDUCATORS = [
  { x: 0.68, w: 2.666, color: C.orange, name: 'Zoe Morgan', role: 'Creative Educator', ny: 4.86 },
  { x: 3.678, w: 2.67, color: C.yellow, name: 'Mira Valentina', role: 'Learning Facilitator', ny: 4.798 },
  { x: 6.675, w: 2.667, color: C.blue, name: 'Arden Reyes', role: 'Preschool Guide', ny: 4.795 },
  { x: 9.678, w: 2.673, color: C.green, name: 'Isla Nakamura', role: 'Class Instructor', ny: 4.784 },
];

function slide06(pres) {
  const s = pres.addSlide();
  title(s, 'Meet Our Educators', { x: 0.686, y: 0.762, w: 6.627, h: 0.963 });
  arcSwoosh(s, 11.525, -0.249, 3.223, 1.585, C.green, 145.89);

  EDUCATORS.forEach((p, i) => {
    s.addShape('round2SameRect', {
      x: p.x, y: 4.759, w: p.w, h: 1.657, fill: { color: p.color },
      line: { type: 'none' }, rectRadius: 0.476,
    });
    s.addText(p.name, {
      x: p.x + 0.167, y: p.ny, w: 2.333, h: 0.378, align: 'center', valign: 'middle',
      fontFace: HEAD, fontSize: 16, bold: true, color: C.white,
      lineSpacingMultiple: 1.5, margin: [7.2, 7.2, 3.6, 3.6],
    });
    s.addText(p.role, {
      x: p.x + 0.17, y: p.ny + 0.297, w: 2.327, h: 0.349, align: 'center', valign: 'top',
      fontFace: BODY, fontSize: 12, bold: true, color: C.white,
      lineSpacingMultiple: 1.5, margin: [7.2, 7.2, 3.6, 3.6],
    });
    body(s, 'Nemo voluptate aspernatu odit fugit voluptat nesciuna', {
      x: p.x - 0.045, y: [5.51, 5.47, 5.471, 5.471][i], w: 2.751, h: 0.675,
      align: 'center', color: C.white,
    });
  });
  return s;
}

// 7 — Our Learning Method: circle diagram, four icon+text quadrants.
const METHOD_QUADRANTS = [
  { ix: 4.77, iy: 2.422, color: C.orange, glyph: G.pencil, tx: 1.216, ty: 2.616, lx: 2.497, ly: 2.313, align: 'right' },
  { ix: 7.911, iy: 2.422, color: C.yellow, glyph: G.search, tx: 8.79, ty: 2.616, lx: 8.79, ly: 2.313, align: 'left' },
  { ix: 4.77, iy: 5.214, color: C.blue, glyph: G.bulb, tx: 1.216, ty: 5.396, lx: 2.497, ly: 5.1, align: 'right' },
  { ix: 7.916, iy: 5.214, color: C.green, glyph: G.gears, tx: 8.79, ty: 5.396, lx: 8.79, ly: 5.1, align: 'left' },
];

function slide07(pres) {
  const s = pres.addSlide();
  s.addShape('ellipse', {
    x: 4.608, y: 2.114, w: 4.118, h: 4.118, fill: { type: 'none' },
    line: { color: C.grey05, width: 10 },
  });
  title(s, 'Our Learning Method');
  cross(s, 12.305, 0.25, 1.557, 1.445, C.blue);
  ring(s, -0.504, 6.998, 1.009, 1.004, C.green);

  METHOD_QUADRANTS.forEach(q => {
    icon(s, q.ix, q.iy, 0.667, q.color, q.glyph);
    label(s, 'Description', {
      x: q.lx, y: q.ly, w: 2.035, h: 0.378, color: q.color, align: q.align,
    });
    body(s,
      'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed non risus. Suspendisse ' +
      'lectus tortor, dignisim sit amet, adipiscing',
      { x: q.tx, y: q.ty, w: 3.316, h: 0.97, align: q.align });
  });
  return s;
}

// 8 — Curriculum Overview: five outlined cards with header pill, blurb, bullets, icon.
const CURRICULUM = [
  { x: 0.998, color: C.orange, name: 'Basics', glyph: G.rocket },
  { x: 3.355, color: C.yellow, name: 'Techniques', glyph: G.pencil },
  { x: 5.677, color: C.blue, name: 'Practice', glyph: G.gears },
  { x: 8.016, color: C.green, name: 'Fixes', glyph: G.tools },
  { x: 10.339, color: C.rust, name: 'Finale', glyph: G.cap },
];
const CURRICULUM_BULLETS = [
  'At vero eos et accu', 'Samus raesentium ', 'Voluptatum atque',
  'Pharetra massa rici', 'Quisan dolores sit',
];

function slide08(pres) {
  const s = pres.addSlide();
  title(s, 'Curriculum Overview');
  multiply(s, 12.34, 0.091, 1.374, 1.342, C.orange);

  CURRICULUM.forEach((card, i) => {
    const y = [2.061, 2.076, 2.074, 2.072, 2.069][i];
    s.addShape('roundRect', {
      x: card.x, y, w: 2.0, h: 4.0, fill: { type: 'none' },
      line: { color: card.color, width: 2 }, rectRadius: 0.212,
    });
    s.addShape('roundRect', {
      x: card.x + 0.125, y: y + 0.136, w: 1.758, h: 0.788, fill: { color: card.color },
      line: { type: 'none' }, rectRadius: 0.161,
    });
    s.addText(card.name, {
      x: card.x - 0.144, y: y + 0.327, w: 2.333, h: 0.396, align: 'center', valign: 'middle',
      fontFace: HEAD, fontSize: 16, bold: true, color: C.white,
      lineSpacingMultiple: 1.0, margin: [7.2, 7.2, 3.6, 3.6],
    });
    body(s, 'Lorem ipsum dolor stu amet, consectetuai', {
      x: card.x - 0.069, y: y + 1.011, w: 2.162, h: 0.667, align: 'center',
      bold: true, color: card.color,
    });
    bulletList(s, CURRICULUM_BULLETS, '•', card.color, {
      x: card.x + 0.094, y: y + 1.817, w: 1.9, h: 1.879,
    });
    icon(s, card.x + 0.673, y + 3.669, 0.667, card.color, card.glyph);
  });
  return s;
}

// 9 — Subjects We Offer: 2x2 icon grid plus a yellow corner block.
const SUBJECTS = [
  { x: 0.687, y: 3.189, tx: 1.349, color: C.orange, glyph: G.heart },
  { x: 4.348, y: 3.196, tx: 5.014, color: C.yellow, glyph: G.palette },
  { x: 0.687, y: 4.633, tx: 1.349, color: C.blue, glyph: G.star },
  { x: 4.349, y: 4.636, tx: 5.018, color: C.green, glyph: G.dumbbell },
];

function slide09(pres) {
  const s = pres.addSlide();
  s.addShape('roundRect', {
    x: 10.668, y: -0.611, w: 3.137, h: 3.694, fill: { color: C.yellow },
    line: { type: 'none' }, rectRadius: 0.48,
  });

  title(s, 'Subjects We Offer', { x: 0.687, y: 1.083, w: 6.647, h: 1.0 });
  body(s,
    'PLACEHOLDER' +
    'PLACEHOLDER',
    { x: 0.664, y: 1.91, w: 6.369, h: 0.667 });

  SUBJECTS.forEach(sub => {
    icon(s, sub.x, sub.y, 0.667, sub.color, sub.glyph);
    label(s, 'Description', {
      x: sub.tx, y: sub.y - 0.026, w: 1.62, h: 0.378, color: sub.color, lineSpacingMultiple: 1.0,
    });
    body(s, 'Cum sociis natoque penati magnis parturient nascetur',
      { x: sub.tx + 0.005, y: sub.y + 0.312, w: 2.65, h: 0.667 });
  });
  return s;
}

// 10 — Class Levels: 3-D style staircase with dashed leader lines.
const LEVELS = [
  { color: C.green, light: C.greenLt, sx: 8.933, sy: 5.323, ty: 5.65, ox: 9.399, oy: 5.115,
    lx: 7.201, ly: 5.439, lw: 2.197, glyph: G.dumbbell, dx: 5.028, dy: 5.165, bx: 1.725, by: 5.481, bw: 5.338,
    ba: 'center',
    text: 'PLACEHOLDER' },
  { color: C.blue, light: C.blueLt, sx: 9.867, sy: 4.201, ty: 4.528, ox: 10.269, oy: 3.936,
    lx: 7.868, ly: 4.254, lw: 2.395, glyph: G.star, dx: 5.699, dy: 3.977, bx: 2.063, by: 4.301, bw: 5.672,
    ba: 'right',
    text: 'PLACEHOLDER' },
  { color: C.yellow, light: C.yellowLt, sx: 10.801, sy: 3.08, ty: 3.406, ox: 11.258, oy: 2.803,
    lx: 8.868, ly: 3.104, lw: 2.397, glyph: G.palette, dx: 6.691, dy: 2.831, bx: 3.055, by: 3.146, bw: 5.671,
    ba: 'right',
    text: 'PLACEHOLDER' },
];

function slide10(pres) {
  const s = pres.addSlide();
  title(s, 'Class Levels ');
  cross(s, -0.21, 2.206, 1.014, 0.94, C.blue);

  // grey staircase body
  s.addShape('custGeom', {
    x: 9.457, y: 3.395, w: 3.212, h: 3.378, fill: { color: C.grey15 }, line: { type: 'none' },
    points: [
      { x: 3.212, y: 0.012 }, { x: 3.212, y: 2.221 }, { x: 0.410, y: 3.378 },
      { x: 0, y: 2.606 }, { x: 1.062, y: 1.423 }, { x: 2.335, y: 0 }, { x: 3.212, y: 0.012 },
      { close: true },
    ],
  });

  LEVELS.forEach(L => {
    // top face (light) + front face (solid)
    s.addShape('custGeom', {
      x: L.sx, y: L.sy, w: 1.868, h: 0.654, fill: { color: L.light }, line: { type: 'none' },
      points: [
        { x: 0, y: 0.327 }, { x: 0.671, y: 0.052 },
        { curve: { type: 'cubic', x1: 0.839, y1: -0.018, x2: 1.028, y2: -0.018 }, x: 1.196, y: 0.052 },
        { x: 1.868, y: 0.327 }, { x: 1.196, y: 0.602 },
        { curve: { type: 'cubic', x1: 1.028, y1: 0.671, x2: 0.839, y2: 0.671 }, x: 0.671, y: 0.602 },
        { x: 0, y: 0.327 }, { close: true },
      ],
    });
    s.addShape('custGeom', {
      x: L.sx + 0.001, y: L.ty, w: 0.934, h: 1.122, fill: { color: L.color }, line: { type: 'none' },
      points: [
        { x: 0.933, y: 1.122 }, { x: 0.106, y: 0.784 },
        { curve: { type: 'cubic', x1: 0.042, y1: 0.757, x2: 0, y2: 0.683 }, x: 0, y: 0.598 },
        { x: 0, y: 0 }, { x: 0.723, y: 0.296 },
        { curve: { type: 'cubic', x1: 0.849, y1: 0.347, x2: 0.934, y2: 0.495 }, x: 0.934, y: 0.661 },
        { close: true },
      ],
    });
    icon(s, L.ox, L.oy, 0.667, L.color, L.glyph);
    s.addShape('line', {
      x: L.lx, y: L.ly, w: L.lw, h: 0,
      line: { color: L.color, width: 1, dashType: 'dash' },
    });
    label(s, 'Description', { x: L.dx, y: L.dy, w: 2.035, h: 0.378, color: L.color, align: 'right' });
    body(s, L.text, { x: L.bx, y: L.by, w: L.bw, h: 0.667, align: L.ba });
  });

  // orange arrow head crowning the staircase
  s.addShape('custGeom', {
    x: 11.542, y: 1.69, w: 1.306, h: 1.716, fill: { color: C.orange }, line: { type: 'none' },
    points: [
      { x: 0.648, y: 0 }, { x: 1.306, y: 1.247 }, { x: 1.120, y: 1.172 }, { x: 1.124, y: 1.194 },
      { curve: { type: 'cubic', x1: 1.127, y1: 1.214, x2: 1.128, y2: 1.234 }, x: 1.128, y: 1.255 },
      { x: 1.127, y: 1.716 }, { x: 0.300, y: 1.378 },
      { curve: { type: 'cubic', x1: 0.236, y1: 1.352, x2: 0.194, y2: 1.277 }, x: 0.194, y: 1.193 },
      { x: 0.194, y: 0.797 }, { x: 0, y: 0.719 }, { close: true },
    ],
  });
  s.addShape('line', {
    x: 10.361, y: 1.902, w: 1.64, h: 0, line: { color: C.orange, width: 1, dashType: 'dash' },
  });
  label(s, 'Description', { x: 8.184, y: 1.63, w: 2.035, h: 0.378, color: C.orange, align: 'right' });
  body(s,
    'PLACEHOLDER' +
    'eget est amet est plac in egestas erat imperdet sed nam risus',
    { x: 4.548, y: 1.953, w: 5.671, h: 0.667, align: 'right' });
  return s;
}

// 11 — Weekly Schedule: five ruled rows (number, day pill, blurb, time).
const WEEK = [
  { num: '01', day: 'Monday', color: C.orange, y: 2.352, nx: 0.835, ny: 2.228, nw: 0.974, nh: 0.757 },
  { num: '02', day: 'Tuesday', color: C.yellow, y: 3.211, nx: 0.825, ny: 3.138, nw: 0.974, nh: 0.655 },
  { num: '03', day: 'Wednesday', color: C.blue, y: 4.044, nx: 0.837, ny: 3.972, nw: 0.974, nh: 0.655 },
  { num: '04', day: 'Thursday', color: C.green, y: 4.868, nx: 0.754, ny: 4.795, nw: 1.179, nh: 0.655 },
  { num: '05', day: 'Friday', color: C.rust, y: 5.694, nx: 0.756, ny: 5.622, nw: 1.177, nh: 0.653 },
];

function slide11(pres) {
  const s = pres.addSlide();
  title(s, 'Weekly Schedule');
  ring(s, 12.002, -0.879, 2.307, 2.296, C.green);

  [2.178, 3.033, 3.862, 4.7, 5.514, 6.358].forEach(y => {
    s.addShape('line', { x: 0.83, y, w: 8.99, h: 0, line: { color: C.grey15, width: 1 } });
  });

  WEEK.forEach((row, i) => {
    s.addText(row.num, {
      x: row.nx, y: row.ny, w: row.nw, h: row.nh, align: 'center', valign: 'middle',
      fontFace: HEAD, fontSize: 44, bold: true, color: row.color,
      lineSpacingMultiple: 1.0, margin: [7.2, 7.2, 3.6, 3.6],
    });
    s.addShape('roundRect', {
      x: 1.979, y: row.y, w: 1.668, h: 0.492, fill: { color: row.color },
      line: { type: 'none' }, rectRadius: 0.246,
    });
    s.addText(row.day, {
      x: 1.979, y: row.y + 0.077, w: 1.671, h: 0.337, align: 'center', valign: 'middle',
      fontFace: HEAD, fontSize: 14, bold: true, color: C.white, margin: [7.2, 7.2, 3.6, 3.6],
    });
    body(s, LOREM_SHORT, { x: 3.97, y: [2.26, 3.119, 3.953, 4.776, 5.602][i], w: 3.75, h: 0.675 });
    s.addText('08.30 AM-09.30 AM', {
      x: 7.827, y: row.y + 0.077, w: 1.989, h: 0.33, align: 'center', valign: 'middle',
      fontFace: BODY, fontSize: 12, color: C.text, lineSpacingMultiple: 1.5,
      margin: [7.2, 7.2, 3.6, 3.6],
    });
  });
  return s;
}

// 12 — Break Time: full-bleed orange section divider.
function slide12(pres) {
  const s = pres.addSlide();
  s.background = { color: C.orange };
  s.addShape('rect', {
    x: 0, y: 0.013, w: 13.333, h: 7.5, fill: { color: C.orange, transparency: 17 },
    line: { type: 'none' },
  });
  multiply(s, 11.65, -0.616, 2.659, 2.315, C.blue);
  multiply(s, -0.814, 5.592, 2.659, 2.315, C.yellow);

  s.addText('Break Time', {
    x: 2.998, y: 3.09, w: 7.337, h: 1.66, align: 'center', fontFace: HEAD, fontSize: 88,
    bold: true, color: C.white, lineSpacingMultiple: 0.9, valign: 'top',
    margin: [7.2, 7.2, 3.6, 3.6],
  });
  s.addText(
    'PLACEHOLDER' +
    'enean massa sociis natoque penatibus', {
      x: 2.998, y: 4.272, w: 7.337, h: 1.133, align: 'center', fontFace: BODY, fontSize: 16,
      color: C.white, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6],
    });
  return s;
}

// 13 — Core Thinking Skills: four-lobe brain + labelled callouts.
// Lobe outlines traced from the reference; points are fractions of each bbox.
const BRAIN = [
  { color: C.orange, x: 4.439, y: 2.085, w: 2.475, h: 2.56, pts: [[0.816, 0.035], [0.733, 0.002], [0.616, 0.006], [0.517, 0.049], [0.448, 0.127], [0.275, 0.164], [0.214, 0.334], [0.135, 0.373], [0.117, 0.463], [0.067, 0.506], [0.077, 0.574], [0, 0.744], [0.044, 0.871], [0.257, 0.865], [0.263, 0.943], [0.311, 0.951], [0.321, 0.912], [0.277, 0.857], [0.418, 0.738], [0.442, 0.691], [0.402, 0.66], [0.418, 0.646], [0.501, 0.678], [0.43, 0.759], [0.689, 0.836], [0.891, 0.998], [0.891, 0.798], [0.82, 0.818], [0.788, 0.755], [0.83, 0.715], [0.891, 0.736], [0.891, 0.434], [0.972, 0.453], [0.998, 0.418], [0.97, 0.389], [0.891, 0.41], [0.893, 0.168]] },
  { color: C.yellow, x: 6.444, y: 2.085, w: 2.475, h: 2.56, pts: [[0.174, 0.043], [0.105, 0.174], [0.105, 0.391], [0.178, 0.371], [0.21, 0.432], [0.168, 0.475], [0.105, 0.451], [0.105, 0.757], [0.028, 0.736], [0, 0.763], [0.028, 0.8], [0.105, 0.779], [0.105, 0.998], [0.224, 0.884], [0.4, 0.798], [0.36, 0.748], [0.386, 0.695], [0.465, 0.713], [0.461, 0.779], [0.564, 0.759], [0.493, 0.68], [0.578, 0.646], [0.558, 0.703], [0.671, 0.832], [0.806, 0.877], [0.956, 0.869], [0.998, 0.728], [0.921, 0.574], [0.933, 0.508], [0.881, 0.463], [0.867, 0.377], [0.786, 0.334], [0.768, 0.244], [0.709, 0.18], [0.727, 0.164], [0.549, 0.127], [0.487, 0.053], [0.384, 0.006], [0.255, 0.004]] },
  { color: C.blue, x: 4.509, y: 4.084, w: 2.405, h: 2.265, pts: [[0.023, 0.128], [0.002, 0.252], [0.058, 0.351], [0.05, 0.437], [0, 0.457], [0.058, 0.47], [0.148, 0.57], [0.15, 0.695], [0.202, 0.781], [0.274, 0.821], [0.362, 0.83], [0.422, 0.934], [0.543, 0.991], [0.709, 0.985], [0.817, 0.905], [0.888, 0.742], [0.888, 0.466], [0.973, 0.49], [0.998, 0.453], [0.965, 0.417], [0.888, 0.442], [0.888, 0.289], [0.807, 0.185], [0.632, 0.066], [0.416, 0], [0.281, 0.099], [0.326, 0.148], [0.32, 0.194], [0.285, 0.223], [0.22, 0.196], [0.222, 0.121]] },
  { color: C.green, x: 6.704, y: 3.905, w: 2.15, h: 2.445, pts: [[0.972, 0.19], [0.73, 0.184], [0.623, 0.139], [0.544, 0.071], [0.393, 0.098], [0.393, 0.016], [0.335, 0.006], [0.321, 0.04], [0.367, 0.104], [0.151, 0.202], [0, 0.341], [0, 0.462], [0.067, 0.437], [0.121, 0.48], [0.081, 0.548], [0, 0.525], [0, 0.748], [0.098, 0.926], [0.23, 0.991], [0.412, 0.987], [0.53, 0.934], [0.591, 0.842], [0.691, 0.834], [0.774, 0.795], [0.828, 0.724], [0.833, 0.603], [0.933, 0.509], [0.998, 0.497], [0.942, 0.476], [0.933, 0.398], [0.995, 0.304]] },
];

const SKILLS = [
  { name: 'Critical Thinking', color: C.orange, px: 0.791, py: 2.283, tx: 0.674, ty: 2.762,
    align: 'left', lx: 3.142, ly: 2.516, lw: 2.075, num: '1', nx: 5.424, ny: 2.922 },
  { name: 'Creativity', color: C.yellow, px: 10.158, py: 2.283, tx: 9.353, ty: 2.749,
    align: 'right', lx: 8.152, ly: 2.516, lw: 2.075, num: '2', nx: 7.263, ny: 2.922 },
  { name: 'Focus', color: C.blue, px: 0.789, py: 4.884, tx: 0.676, ty: 5.36,
    align: 'left', lx: 3.142, ly: 5.117, lw: 1.49, num: '3', nx: 5.424, ny: 4.673 },
  { name: 'Adaptability', color: C.green, px: 10.158, py: 4.884, tx: 9.354, ty: 5.352,
    align: 'right', lx: 8.736, ly: 5.117, lw: 1.49, num: '4', nx: 7.263, ny: 4.673 },
];

function slide13(pres) {
  const s = pres.addSlide();
  title(s, 'Core Thinking Skills');

  BRAIN.forEach(lobe => {
    s.addShape('custGeom', {
      x: lobe.x, y: lobe.y, w: lobe.w, h: lobe.h, fill: { color: lobe.color },
      line: { type: 'none' },
      points: lobe.pts.map(([u, v]) => ({ x: u * lobe.w, y: v * lobe.h })).concat([{ close: true }]),
    });
  });

  SKILLS.forEach(sk => {
    s.addShape('roundRect', {
      x: sk.px, y: sk.py, w: 2.383, h: 0.466, fill: { color: sk.color },
      line: { type: 'none' }, rectRadius: 0.233,
    });
    s.addText(sk.name, {
      x: sk.px - 0.117, y: sk.py + 0.059, w: 2.618, h: 0.37, align: 'center', valign: 'middle',
      fontFace: HEAD, fontSize: 16, bold: true, color: C.white, margin: [7.2, 7.2, 3.6, 3.6],
    });
    s.addShape('line', {
      x: sk.lx, y: sk.ly, w: sk.lw, h: 0,
      line: { color: sk.color, width: 1.25, dashType: 'dash' },
    });
    body(s, DIAM, { x: sk.tx, y: sk.ty, w: 3.315, h: 0.985, align: sk.align });
    s.addShape('ellipse', {
      x: sk.nx, y: sk.ny, w: 0.739, h: 0.739, fill: { color: C.white }, line: { type: 'none' },
    });
    s.addText(sk.num, {
      x: sk.nx, y: sk.ny, w: 0.739, h: 0.739, align: 'center', valign: 'middle',
      fontFace: BODY, fontSize: 24, bold: true, color: sk.color, margin: 0,
    });
  });
  return s;
}

// 14 — Class Allocation: orange stat panel + pie chart + 2x2 legend.
const ALLOCATION = [
  { name: 'Theory', color: C.orange, ix: 0.84, iy: 4.128, tx: 1.483, ty: 4.074, glyph: G.pencil },
  { name: 'Group Discussion', color: C.yellow, ix: 4.187, iy: 4.145, tx: 4.816, ty: 4.038, glyph: G.people, w: 2.382 },
  { name: 'Practice', color: C.blue, ix: 0.844, iy: 5.361, tx: 1.485, ty: 5.298, glyph: G.gears },
  { name: 'Evaluation', color: C.green, ix: 4.188, iy: 5.369, tx: 4.816, ty: 5.288, glyph: G.star },
];

function slide14(pres) {
  const s = pres.addSlide();
  s.addShape('roundRect', {
    x: -0.457, y: -0.703, w: 7.79, h: 4.453, fill: { color: C.orange, transparency: 15 },
    line: { type: 'none' }, rectRadius: 0.449,
  });
  ring(s, 12.002, -0.879, 2.307, 2.296, C.orange);

  title(s, 'Class Allocation', { color: C.white });
  s.addText('Female Student', {
    x: 0.612, y: 1.676, w: 1.973, h: 0.337, align: 'center', valign: 'middle',
    fontFace: HEAD, fontSize: 14, bold: true, color: C.white, margin: [7.2, 7.2, 3.6, 3.6],
  });
  s.addText('756K', {
    x: 0.612, y: 1.908, w: 1.973, h: 0.774, align: 'center', valign: 'middle',
    fontFace: HEAD, fontSize: 40, bold: true, color: C.white, margin: [7.2, 7.2, 3.6, 3.6],
  });
  body(s, LOREM_SHORT, { x: 2.583, y: 1.891, w: 3.75, h: 0.667, color: C.white });
  body(s, DIAM, { x: 0.752, y: 2.555, w: 5.255, h: 0.667, color: C.white });

  s.addChart(pres.ChartType.pie, [{
    name: 'Sales',
    labels: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'],
    values: [8.2, 3.2, 1.4, 1.2],
  }], {
    x: 7.573, y: 1.417, w: 5.405, h: 5.139,
    chartColors: [C.orange, C.yellow, C.blue, C.green],
    dataBorder: { pt: 1.5, color: C.white },
    showLegend: false, showTitle: false, showValue: false, firstSliceAng: 0,
  });

  ALLOCATION.forEach(a => {
    icon(s, a.ix, a.iy, 0.606, a.color, a.glyph);
    s.addText(a.name, {
      x: a.tx, y: a.ty, w: a.w || 1.821, h: 0.429, valign: 'middle',
      fontFace: HEAD, fontSize: 16, bold: true, color: a.color,
      lineSpacingMultiple: 1.5, margin: [7.2, 7.2, 3.6, 3.6],
    });
    body(s, 'Diam in arcu cursus euismode quis viverra nibh pulvu ut',
      { x: a.tx + 0.004, y: a.ty + 0.326, w: 2.412, h: 0.667, valign: 'middle' });
  });
  return s;
}

// 15 — Online & Offline Learning: laptop mock-up placeholder + two paragraphs.
function slide15(pres) {
  const s = pres.addSlide();
  cross(s, 10.728, 0.056, 2.396, 2.223, C.blue);
  s.addShape('round1Rect', {
    x: 5.703, y: 1.745, w: 1.524, h: 2.941, fill: { color: C.yellow },
    line: { type: 'none' }, rectRadius: 0.762, flipH: true,
  });
  // laptop mock-up photo (the deck's only embedded raster) -> grey block
  imagePlaceholder(s, 5.751, 1.08, 6.719, 4.444, { fill: 'D6D6D6', labelColor: '6E6E6E' });
  s.addShape('ellipse', {
    x: 11.383, y: 3.763, w: 0.75, h: 0.73, fill: { color: C.green }, line: { type: 'none' },
  });
  s.addShape('blockArc', {
    x: 8.005, y: 6.133, w: 1.333, h: 1.298, fill: { color: C.orange }, line: { type: 'none' },
    angleRange: [180, 0], arcThicknessRatio: 0.47,
  });

  s.addText('Online & Offline Learning', {
    x: 1.02, y: 1.083, w: 4.313, h: 2.333, fontFace: HEAD, fontSize: 44, bold: true,
    color: C.orange, lineSpacingMultiple: 0.9, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6],
  });
  body(s, [
    { text: 'Vel risus commodo viverra maecenas ac', options: { breakLine: true } },
    { text: 'PLACEHOLDER', options: { breakLine: true } },
    { text: '', options: { breakLine: true } },
    { text: 'Nam at lectus urna dictum amet justo do', options: { breakLine: true } },
    { text: 'nec ornare. Sed persp unde omi iste natu vol uptat accusantiui doloru laudai totam' },
  ], { x: 0.998, y: 3.596, w: 3.333, h: 2.182 });
  return s;
}

// 16 — Education Features: three grey cards with coloured headers and check lists.
const FEATURES = [
  { x: 0.665, w: 3.69, color: C.orange, title: 'Theory-Based Learning', on: [1, 1, 0, 0] },
  { x: 4.832, w: 3.69, color: C.yellow, title: 'Hands-On Practice', on: [1, 1, 1, 0] },
  { x: 9.0, w: 3.668, color: C.blue, title: 'Interactive Discussion', on: [1, 1, 1, 1] },
];
const FEATURE_ITEMS = [
  'Diam in arcu cursus euismo quis nibh ',
  'Purusa sit amet luctus venenat auctor ',
  'Accusantium doloremque laudantium ',
  'Diam in arcu cursus euismo quis nibh ',
];

function slide16(pres) {
  const s = pres.addSlide();
  title(s, 'Education Features');
  s.addShape('roundRect', {
    x: 11.689, y: -0.131, w: 1.911, h: 0.881, fill: { color: C.green },
    line: { type: 'none' }, rectRadius: 0.147,
  });

  FEATURES.forEach((f, i) => {
    const hy = [3.748, 3.749, 3.751][i];
    s.addShape('round2SameRect', {
      x: f.x, y: hy + 0.328, w: f.w, h: 2.662, fill: { color: C.grey05 },
      line: { type: 'none' }, rectRadius: 0.319, flipV: true,
    });
    s.addShape('round2SameRect', {
      x: f.x, y: hy, w: f.w, h: 0.672, fill: { color: f.color },
      line: { type: 'none' }, rectRadius: 0.112,
    });
    s.addText(f.title, {
      x: f.x + 0.333, y: hy + 0.13 + i * 0.01, w: 3.002, h: 0.37, align: 'center', valign: 'middle',
      fontFace: HEAD, fontSize: 16, bold: true, color: C.white, margin: [7.2, 7.2, 3.6, 3.6],
    });
    FEATURE_ITEMS.forEach((item, j) => {
      const y = [4.742, 5.154, 5.569, 5.988][j] + (i === 1 ? 0.003 : 0) - (i === 2 ? 0.001 : 0);
      const on = f.on[j];
      s.addShape('ellipse', {
        x: f.x + 0.206, y, w: 0.26, h: 0.26,
        fill: { color: on ? f.color : C.greyIcon }, line: { type: 'none' },
      });
      s.addText(on ? '✓' : '✕', {
        x: f.x + 0.206, y, w: 0.26, h: 0.26, align: 'center', valign: 'middle',
        fontFace: SYM, fontSize: 9, bold: true, color: C.white, margin: 0,
      });
      body(s, item, {
        x: f.x + 0.528, y: y - 0.021, w: 3.142, h: 0.303, color: C.slate, lineSpacingMultiple: 1.0,
      });
    });
  });
  return s;
}

// 17 — Our Activities: intro paragraph + 2x2 icon grid.
const ACTIVITIES = [
  { name: 'Project', color: C.orange, ix: 0.78, iy: 3.597, tx: 1.408, ty: 3.433, glyph: G.pencil },
  { name: 'Collaboration', color: C.yellow, ix: 3.407, iy: 3.596, tx: 4.027, ty: 3.442, glyph: G.people, w: 2.605 },
  { name: 'Simulation', color: C.blue, ix: 0.784, iy: 4.981, tx: 1.409, ty: 4.814, glyph: G.gears },
  { name: 'Reflection', color: C.green, ix: 3.407, iy: 4.97, tx: 4.036, ty: 4.804, glyph: G.star },
];

function slide17(pres) {
  const s = pres.addSlide();
  title(s, 'Our Activities', { x: 0.687, y: 1.408, w: 4.98, h: 1.0 });
  ring(s, 6.012, 1.115, 0.667, 0.667, C.green);
  multiply(s, 11.875, 3.404, 1.168, 1.14, C.orange);
  s.addShape('roundRect', {
    x: -0.492, y: 7.083, w: 2.49, h: 0.838, fill: { color: C.yellow },
    line: { type: 'none' }, rectRadius: 0.419,
  });

  body(s,
    'PLACEHOLDER' +
    'PLACEHOLDER',
    { x: 0.687, y: 2.257, w: 5.15, h: 0.97 });

  ACTIVITIES.forEach(a => {
    icon(s, a.ix, a.iy, 0.627, a.color, a.glyph);
    s.addText(a.name, {
      x: a.tx, y: a.ty, w: a.w || 1.46, h: 0.442, valign: 'middle',
      fontFace: HEAD, fontSize: 16, bold: true, color: a.color,
      lineSpacingMultiple: 1.5, margin: [7.2, 7.2, 3.6, 3.6],
    });
    body(s, 'Diam in arcu cursus euismode quis viverra',
      { x: a.tx - 0.004, y: a.ty + 0.368, w: 1.875, h: 0.626, valign: 'middle' });
  });
  return s;
}

// 18 — Our Pricelist: three price bars, each with a divider and two check columns.
const PRICES = [
  { y: 1.417, color: C.orange, price: '$90', right: 1 },
  { y: 3.083, color: C.yellow, price: '$100', right: 2 },
  { y: 4.75, color: C.blue, price: '$120', right: 3 },
];
const PRICE_ITEMS = ['At vero eos et accusamus', 'Praesentium voluptatum', 'Deleniti atque corrupti'];
const NOTES = [
  'Diam in arcu cursus euismo quis viverra nibh ',
  'Purusa sit amet luctus venenat vitae auctor ',
  'Accusantium doloremque laudantium ',
];

function slide18(pres) {
  const s = pres.addSlide();
  arcSwoosh(s, -0.316, 6.496, 1.295, 0.637, C.green, 16.56);
  cross(s, 4.765, 0.417, 0.469, 0.436, C.blue);

  s.addText('Our Pricelist', {
    x: 0.687, y: 1.083, w: 4.797, h: 2.458, fontFace: HEAD, fontSize: 44, bold: true,
    color: C.orange, lineSpacingMultiple: 0.9, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6],
  });
  body(s,
    'PLACEHOLDER' +
    'eget est amet est plac in egestas erat imperdet sed nam at lectus urna dictum amet justo don.',
    { x: 0.712, y: 2.673, w: 3.745, h: 1.273 });

  label(s, 'Notes', { x: 0.712, y: 4.469, w: 2.035, h: 0.333, lineSpacingMultiple: 1.5 });
  NOTES.forEach((n, i) => {
    const y = 4.904 + i * 0.415;
    s.addShape('ellipse', { x: 0.83, y, w: 0.26, h: 0.26, fill: { color: C.orange }, line: { type: 'none' } });
    s.addText('✓', {
      x: 0.83, y, w: 0.26, h: 0.26, align: 'center', valign: 'middle',
      fontFace: SYM, fontSize: 9, bold: true, color: C.white, margin: 0,
    });
    body(s, n, { x: 1.152, y: y - 0.021, w: 3.652, h: 0.303, lineSpacingMultiple: 1.0 });
  });

  PRICES.forEach(p => {
    s.addShape('roundRect', {
      x: 5.674, y: p.y, w: 6.987, h: 1.335, fill: { color: p.color },
      line: { type: 'none' }, rectRadius: 0.125,
    });
    s.addText('Package', {
      x: 5.667, y: p.y + 0.146, w: 1.868, h: 0.362, align: 'center', valign: 'middle',
      fontFace: HEAD, fontSize: 16, bold: true, color: C.white,
      lineSpacingMultiple: 1.0, margin: [7.2, 7.2, 3.6, 3.6],
    });
    s.addText(p.price, {
      x: 5.67, y: p.y + 0.489, w: 1.861, h: 0.707, align: 'center',
      fontFace: HEAD, fontSize: 36, bold: true, color: C.white, margin: [7.2, 7.2, 3.6, 3.6],
    });
    s.addShape('line', {
      x: 7.54, y: p.y + 0.166, w: 0, h: 1.0, line: { color: C.white, width: 1 },
    });
    bulletList(s, PRICE_ITEMS, '✓', C.white,
      { x: 7.725, y: p.y + 0.146, w: 2.328, h: 0.97, color: C.white });
    bulletList(s, PRICE_ITEMS.slice(0, p.right), '✓', C.white,
      { x: 10.204, y: p.y + 0.146, w: 2.328, h: 0.97, color: C.white });
  });
  return s;
}

// 19 — Contact Us: three contact rows plus a yellow "Have a question?" bubble.
const CONTACTS = [
  { name: 'Official Phone', color: C.orange, glyph: G.phone, ix: 1.016, y: 3.09, lx: 1.877, ty: 3.102, bx: 1.903 },
  { name: 'Official Email', color: C.yellow, glyph: G.mail, ix: 0.995, y: 4.08, lx: 1.875, ty: 4.099, bx: 1.9, w: 2.442 },
  { name: 'Official Website', color: C.blue, glyph: G.globe, ix: 1.0, y: 5.087, lx: 1.891, ty: 5.096, bx: 1.917 },
];

function slide19(pres) {
  const s = pres.addSlide();
  s.addShape('rect', { x: 7.333, y: -0.287, w: 5.005, h: 4.037, fill: { color: C.orange }, line: { type: 'none' } });

  title(s, 'Contact Us', { x: 1.02, y: 1.083, w: 4.31, h: 1.0 });
  body(s,
    'PLACEHOLDER',
    { x: 1.016, y: 1.77, w: 3.984, h: 0.667 });

  icon(s, 7.014, 2.738, 0.679, C.green, G.cap);
  icon(s, 11.349, 4.425, 0.983, C.blue, G.trophy);

  CONTACTS.forEach(c => {
    icon(s, c.ix, c.y, 0.679, c.color, c.glyph);
    label(s, c.name, { x: c.lx, y: c.ty, w: c.w || 2.035, h: 0.333, color: c.color, lineSpacingMultiple: 1.5 });
    body(s, 'Sed ut persp unde omnis iste natusi', { x: c.bx, y: c.ty + 0.267, w: 2.969, h: 0.364 });
  });

  s.addShape('ellipse', {
    x: 7.014, y: 4.063, w: 2.353, h: 2.353, fill: { color: C.yellow }, line: { type: 'none' },
  });
  s.addText([
    { text: 'Have a question?', options: { breakLine: true } },
    { text: 'We are here to help' },
  ], {
    x: 7.132, y: 4.673, w: 2.194, h: 1.134, align: 'center', valign: 'middle',
    fontFace: HEAD, fontSize: 18, bold: true, color: C.white,
    lineSpacingMultiple: 1.0, margin: [7.2, 7.2, 3.6, 3.6],
  });
  return s;
}

// 20 — Closing: full-bleed orange, big title, yellow contact bar.
const FOOTER = [
  { name: 'Official Phone', ox: 0.671, gx: 0.748, tx: 1.138, tw: 2.144, glyph: G.phone },
  { name: 'Official Email', ox: 5.586, gx: 5.701, tx: 6.055, tw: 1.709, glyph: G.mail },
  { name: 'Official Website', ox: 10.212, gx: 10.311, tx: 10.701, tw: 2.136, glyph: G.globe },
];

function slide20(pres) {
  const s = pres.addSlide();
  s.background = { color: C.orange };
  s.addShape('rect', {
    x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: C.orange, transparency: 10 }, line: { type: 'none' },
  });

  s.addShape('blockArc', {
    x: 9.043, y: -0.202, w: 1.918, h: 1.887, fill: { color: C.yellow }, line: { type: 'none' },
    angleRange: [180, 0], arcThicknessRatio: 0.5,
  });
  s.addShape('blockArc', {
    x: -1.069, y: 1.758, w: 1.918, h: 1.887, fill: { color: C.yellow }, line: { type: 'none' },
    angleRange: [180, 0], arcThicknessRatio: 0.5, rotate: 150,
  });
  ring(s, 1.447, -0.358, 1.26, 1.254, C.green);
  multiply(s, 11.204, 1.319, 1.294, 1.127, C.blue);

  s.addText('Let’s Learn Together', {
    x: 2.998, y: 2.04, w: 7.337, h: 3.044, align: 'center', fontFace: HEAD, fontSize: 88,
    bold: true, color: C.white, lineSpacingMultiple: 0.9, valign: 'top',
    margin: [7.2, 7.2, 3.6, 3.6],
  });
  s.addText('Vulputate eu scelerisque imperdiet proin fermentu tristique', {
    x: 2.998, y: 4.932, w: 7.337, h: 0.659, align: 'center', fontFace: BODY, fontSize: 16,
    color: C.white, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6],
  });

  s.addShape('roundRect', {
    x: 0.333, y: 6.417, w: 12.385, h: 0.667, fill: { color: C.yellow },
    line: { type: 'none' }, rectRadius: 0.333,
  });
  FOOTER.forEach(f => {
    s.addShape('ellipse', {
      x: f.ox, y: 6.52, w: 0.462, h: 0.462, fill: { color: C.white }, line: { type: 'none' },
    });
    f.glyph(s, { cx: f.ox + 0.231, cy: 6.751, d: 0.462, fg: C.yellow, bg: C.white });
    s.addText(f.name, {
      x: f.tx, y: 6.599, w: f.tw, h: 0.337, fontFace: HEAD, fontSize: 14, color: C.white,
      valign: 'top', margin: [7.2, 7.2, 3.6, 3.6],
    });
  });
  return s;
}

/* ================================================================== build */

const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
];

// Slides 2-11 and 13-19 carry the master's yellow page-number badge.
const NUMBERED = new Set([2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 14, 15, 16, 17, 18, 19]);

function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
  pres.layout = 'WIDE';
  pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };
  pres.title = 'Education';

  BUILDERS.forEach((fn, i) => {
    const slide = fn(pres);
    if (NUMBERED.has(i + 1)) slideNumber(slide, i + 1);
  });

  return pres.writeFile({
    fileName: path.join(__dirname, '07d4b1a0-ed59-4ef2-91cd-a1d487416db7_grok_final.pptx'),
  });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
