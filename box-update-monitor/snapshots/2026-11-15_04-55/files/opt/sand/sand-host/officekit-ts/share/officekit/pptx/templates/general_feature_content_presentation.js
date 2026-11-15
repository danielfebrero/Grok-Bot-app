/**
 * Rebuild of "0b8e3f34-7223-4d05-93df-74bb98f2e424.pptx" with PptxGenJS.
 *
 * The source deck is a 26.6 x 15 in "screen presentation" template: a soft grey
 * page wash, violet 45-degree gradient panels and photographic device mock-ups.
 * The photographs are replaced here by programmatic stand-ins built from native
 * shapes - nothing raster is embedded.
 *
 *   node 0b8e3f34-7223-4d05-93df-74bb98f2e424_grok_final.js
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const SLIDE_W = 26.6006944; // 24323675 EMU
const SLIDE_H = 15.0; // 13716000 EMU

const WHITE = 'FFFFFF';
const DIM_WHITE = 'BFBFBF'; // scheme bg1 at lumMod 75%
const ACCENT3 = '5D739A';
const ACCENT4 = '6997AF';
const ACCENT5 = '84ACB6';
const ACCENT6 = '6F8183';

// Violet panels: accent2 gradient, dark at the bottom-left, light at top-right.
const PANEL_STOPS = [
  [0, '4D4C7A'],
  [1, '6C6AA9'],
];
// Page wash: white in the top-left corner fading to grey in the bottom-right.
const PAGE_STOPS = [
  [0.0, 'FFFFFF'],
  [0.22, 'FFFFFF'],
  [0.5, 'F5F5F5'],
  [0.85, 'E8E8E8'],
  [1.0, 'E8E8E8'],
];

const BODY_FONT = 'Lato';
const HEAD_FONT = 'Lato Black';

/* ------------------------------------------------------------- lorem text */

const LEAD = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. ';
const TAIL =
  'Aliquam tincidunt ante nec sem congue convallis. Pellentesque vel mauris quis nisl ornare ' +
  'rutrum in id risus. Proin vehicula ut sem et tempus. Interdum et malesuada fames ac ante ' +
  'ipsum primis in faucibus. ';
const SHORT_TAIL = 'Aliquam tincidunt ante nec sem congue convallis. ';
const RISUS = 'Pellentesque vel mauris quis nisl ornare rutrum in id risus. ';

const T = {};
T.a = LEAD + TAIL; // slides 1, 6, 9, 13
T.b = T.a + 'Pellentesque.'; // the deck's default paragraph
T.c = T.b + ' ' + TAIL + 'Pellentesque.'; // slides 10, 11, 15
T.d = T.b + ' ' + SHORT_TAIL; // slide 14
T.e = T.b + ' consectetur adipiscing elit. ' + TAIL + 'Pellentesque.'; // slides 17, 18
T.f = T.c + LEAD + SHORT_TAIL + RISUS; // slides 23, 28, 33-36, 38
T.g = T.b + ' ' + SHORT_TAIL + RISUS + 'nec sem congue convallis. ' + RISUS; // slide 24
T.h = LEAD + TAIL.slice(0, -2); // slide 25 caption (no closing period)
T.i = LEAD + 'Aliquam tincidunt ante nec sem congue convallis'; // banner strap line
T.j = T.c + LEAD; // slide 32
T.k = T.c + 'Lorem'; // slide 37
T.check =
  'PLACEHOLDER' +
  'Data analysis which summarizes. Methodologies.';
T.checkShort = 'Lorem Ipsum has two main data statisticalv';

/* --------------------------------------------------------------- plumbing */

/** Colour at position `k` (0..1) along a list of [stop, hex] pairs. */
function sample(stops, k) {
  let i = 1;
  while (i < stops.length - 1 && stops[i][0] < k) i++;
  const [k0, c0] = stops[i - 1];
  const [k1, c1] = stops[i];
  const f = k1 === k0 ? 0 : (k - k0) / (k1 - k0);
  let hex = '';
  for (let ch = 0; ch < 6; ch += 2) {
    const a = parseInt(c0.substr(ch, 2), 16);
    const b = parseInt(c1.substr(ch, 2), 16);
    hex += Math.round(a + (b - a) * f)
      .toString(16)
      .padStart(2, '0');
  }
  return hex.toUpperCase();
}

const poly = (pts) => pts.map(([x, y]) => ({ x, y })).concat([{ close: true }]);

/** Clip a polygon against the half plane `value(pt) >= level` (or `<=`). */
function clipHalfPlane(pts, value, level, keepAbove) {
  const out = [];
  for (let i = 0; i < pts.length; i++) {
    const a = pts[i];
    const b = pts[(i + 1) % pts.length];
    const fa = (keepAbove ? 1 : -1) * (value(a) - level);
    const fb = (keepAbove ? 1 : -1) * (value(b) - level);
    if (fa >= 0) out.push(a);
    if (fa >= 0 !== fb >= 0) {
      const s = fa / (fa - fb);
      out.push([a[0] + (b[0] - a[0]) * s, a[1] + (b[1] - a[1]) * s]);
    }
  }
  return out;
}

/**
 * PptxGenJS cannot emit gradient fills, so a 45-degree gradient is painted as a
 * fan of solid bands running perpendicular to the box diagonal.
 *   dir 'up'    - stop 0 at the bottom-left corner, stop 1 at the top-right
 *   dir 'down'  - stop 0 at the top-left corner, stop 1 at the bottom-right
 *   dir 'right' - stop 0 at the left edge, stop 1 at the right edge
 */
function gradientBox(slide, x, y, w, h, stops, dir, steps) {
  const span = dir === 'right' ? w : w + h;
  const t =
    dir === 'up'
      ? (p) => (p[0] + h - p[1]) / span
      : dir === 'right'
        ? (p) => p[0] / span
        : (p) => (p[0] + p[1]) / span;
  const bleed = 0.75 / steps; // bands overlap slightly so no seams show
  for (let i = 0; i < steps; i++) {
    let band = [
      [0, 0],
      [w, 0],
      [w, h],
      [0, h],
    ];
    band = clipHalfPlane(band, t, i / steps - bleed, true);
    band = clipHalfPlane(band, t, (i + 1) / steps + bleed, false);
    if (band.length < 3) continue;
    slide.addShape('custGeom', {
      x,
      y,
      w,
      h,
      points: poly(band),
      fill: { color: sample(stops, (i + 0.5) / steps) },
    });
  }
}

const panel = (slide, x, y, w, h) => gradientBox(slide, x, y, w, h, PANEL_STOPS, 'up', 16);

function newSlide(pres) {
  const slide = pres.addSlide();
  slide.background = { color: 'FFFFFF' };
  gradientBox(slide, 0, 0, SLIDE_W, SLIDE_H, PAGE_STOPS, 'down', 22);
  return slide;
}

/* ------------------------------------------------------------ text pieces */

/**
 * White body copy: 24pt Lato, 130% leading, vertically centred in its box.
 * Several source text frames end with a blank paragraph, which lifts the copy
 * inside a centred box - `pad` reproduces that.
 */
function bodyText(slide, text, x, y, w, h, opts = {}) {
  const runs = [{ text, options: { breakLine: true } }];
  if (opts.pad) runs.push({ text: '' });
  slide.addText(runs, {
    x,
    y,
    w,
    h,
    fontFace: BODY_FONT,
    fontSize: 24,
    color: opts.color || WHITE,
    align: opts.align || 'left',
    valign: 'middle',
    lineSpacingMultiple: 1.3,
    margin: [19.2, 19.2, 9.6, 9.6],
  });
}

/** Heading set in Lato Black, anchored to the top of its box. */
function heading(slide, text, x, y, w, h, size, color = WHITE) {
  slide.addText(text, {
    x,
    y,
    w,
    h,
    fontFace: HEAD_FONT,
    fontSize: size,
    bold: true,
    color,
    valign: 'top',
  });
}

/** The violet "PRESENTATION" flag tucked into a corner of most slides. */
function headerTag(slide, x = 0, y = -0.018) {
  gradientBox(slide, x, y, 5.217, 0.977, PANEL_STOPS, 'up', 10);
  slide.addText('PRESENTATION', {
    x: x + 0.134,
    y: y + 0.018,
    w: 5.029,
    h: 0.875,
    fontFace: HEAD_FONT,
    fontSize: 40,
    bold: true,
    color: WHITE,
    valign: 'top',
    margin: [14.4, 14.4, 7.2, 7.2],
  });
}

/** Facebook / Twitter / LinkedIn / Skype buttons: coloured discs with a glyph. */
function socialRow(slide, x, y) {
  const D = 0.969;
  [
    { dx: 0.0, color: ACCENT6, glyph: 'f' },
    { dx: 1.078, color: ACCENT3, glyph: 't' },
    { dx: 2.203, color: ACCENT4, glyph: 'in' },
    { dx: 3.281, color: ACCENT6, glyph: 'S' },
  ].forEach((ic) => {
    slide.addShape('ellipse', { x: x + ic.dx, y, w: D, h: D, fill: { color: ic.color } });
    slide.addText(ic.glyph, {
      x: x + ic.dx,
      y: y + 0.02,
      w: D,
      h: D,
      fontFace: HEAD_FONT,
      fontSize: 34,
      bold: true,
      color: WHITE,
      align: 'center',
      valign: 'middle',
      margin: 0,
    });
  });
}

/** A tick inside a coloured disc with a grey caption to its right. */
function checkRow(slide, o) {
  const D = 1.014;
  slide.addShape('ellipse', { x: o.cx, y: o.cy, w: D, h: D, fill: { color: o.color } });
  const s = D * 0.54;
  slide.addShape('custGeom', {
    x: o.cx + (D - s) / 2,
    y: o.cy + (D - s) / 2,
    w: s,
    h: s,
    points: poly(
      [
        [0.0, 0.46],
        [0.16, 0.3],
        [0.38, 0.55],
        [0.85, 0.03],
        [1.0, 0.19],
        [0.38, 0.9],
      ].map(([u, v]) => [u * s, v * s])
    ),
    fill: { color: WHITE },
  });
  slide.addText(o.text, {
    x: o.tx,
    y: o.ty,
    w: o.tw,
    h: 1.05,
    fontFace: BODY_FONT,
    fontSize: 24,
    color: DIM_WHITE,
    valign: 'top',
    margin: [17.28, 17.28, 8.64, 8.64],
  });
}

/* ------------------------------------------- device mock-up stand-ins */

/** Desktop all-in-one: black display with a glass glare, light chin and stand. */
function iMac(slide, x, y, w, h) {
  const chin = 0.7145 * h; // top of the aluminium chin
  const body = 0.822 * h; // bottom of the display housing
  slide.addShape('roundRect', {
    x,
    y,
    w,
    h: body,
    rectRadius: 0.022 * w,
    fill: { color: '040404' },
  });
  // Glass reflection: a diagonal band across the top-right of the screen that
  // fades out downwards. Drawn as slabs, each starting further to the right.
  const GLARE_BOTTOM = 0.68; // fraction of the screen the reflection reaches
  const GLARE_SLABS = 14;
  for (let i = 0; i < GLARE_SLABS; i++) {
    const v0 = (i / GLARE_SLABS) * GLARE_BOTTOM;
    const v1 = ((i + 1) / GLARE_SLABS) * GLARE_BOTTOM + 0.004;
    const edge = (v) => Math.min(1, 0.5 + 0.72 * v);
    const tone = Math.round(70 - (62 * i) / GLARE_SLABS);
    slide.addShape('custGeom', {
      x,
      y,
      w,
      h: chin,
      points: poly(
        [
          [edge(v0), v0],
          [1, v0],
          [1, v1],
          [edge(v1), v1],
        ].map(([u, v]) => [u * w, v * chin])
      ),
      fill: { color: tone.toString(16).padStart(2, '0').repeat(3).toUpperCase() },
    });
  }
  // Brushed aluminium chin, lighter towards the right.
  gradientBox(
    slide,
    x,
    y + chin,
    w,
    body - chin,
    [
      [0, '96989C'],
      [0.4, 'C6C8CB'],
      [1, 'D1D3D5'],
    ],
    'right',
    30
  );
  slide.addShape('custGeom', {
    x: x + 0.37 * w,
    y: y + body,
    w: 0.26 * w,
    h: 0.125 * h,
    points: poly([
      [0.016 * w, 0],
      [0.241 * w, 0],
      [0.256 * w, 0.125 * h],
      [0.0, 0.125 * h],
    ]),
    fill: { color: 'B7B9BD' },
  });
  slide.addShape('roundRect', {
    x: x + 0.329 * w,
    y: y + 0.945 * h,
    w: 0.34 * w,
    h: 0.014 * h,
    rectRadius: 0.007 * h,
    fill: { color: 'BBBDC1' },
  });
}

/** Laptop: black lid with a hairline bezel over a silver wedge base. */
function macBook(slide, x, y, w, h) {
  slide.addShape('roundRect', {
    x: x + 0.125 * w,
    y,
    w: 0.748 * w,
    h: 0.909 * h,
    rectRadius: 0.014 * w,
    fill: { color: 'C2C3C7' },
  });
  slide.addShape('roundRect', {
    x: x + 0.128 * w,
    y: y + 0.006 * h,
    w: 0.742 * w,
    h: 0.898 * h,
    rectRadius: 0.012 * w,
    fill: { color: '050505' },
  });
  slide.addShape('custGeom', {
    x,
    y: y + 0.902 * h,
    w,
    h: 0.05 * h,
    points: poly([
      [0.125 * w, 0],
      [0.873 * w, 0],
      [0.962 * w, 0.05 * h],
      [0.038 * w, 0.05 * h],
    ]),
    fill: { color: 'CFD0D4' },
  });
  slide.addShape('roundRect', {
    x: x + 0.038 * w,
    y: y + 0.947 * h,
    w: 0.924 * w,
    h: 0.024 * h,
    rectRadius: 0.012 * h,
    fill: { color: 'ADAEB2' },
  });
  slide.addShape('roundRect', {
    x: x + 0.44 * w,
    y: y + 0.947 * h,
    w: 0.12 * w,
    h: 0.017 * h,
    rectRadius: 0.006 * h,
    fill: { color: 'C4C5C9' },
  });
}

/**
 * Tablet. Geometry is written in the portrait frame; passing `landscape`
 * rotates it 90 degrees clockwise into the given (wide) box, which is how the
 * source deck uses it on slide 12.
 */
function tablet(slide, x, y, w, h, landscape = false) {
  const pw = landscape ? h : w; // device width  in inches
  const ph = landscape ? w : h; // device height in inches
  const at = (u, v, du, dv) =>
    landscape
      ? { x: x + w * (1 - v - dv), y: y + h * u, w: w * dv, h: h * du }
      : { x: x + w * u, y: y + h * v, w: w * du, h: h * dv };
  const disc = (d) => [d / pw, d / ph]; // a circle of diameter d, normalised

  slide.addShape('roundRect', {
    ...at(0.029, 0.0, 0.945, 0.998),
    rectRadius: 0.05 * pw,
    fill: { color: '3C3C3E' },
  });
  slide.addShape('roundRect', {
    ...at(0.036, 0.004, 0.93, 0.99),
    rectRadius: 0.045 * pw,
    fill: { color: '141416' },
  });
  slide.addShape('rect', { ...at(0.039, 0.05, 0.925, 0.886), fill: { color: '070707' } });
  const cam = disc(0.012 * pw);
  slide.addShape('ellipse', {
    ...at(0.494, 0.023, cam[0], cam[1]),
    fill: { color: '2E2E30' },
  });
  const home = disc(0.06 * pw);
  slide.addShape('ellipse', {
    ...at(0.47, 0.952, home[0], home[1]),
    fill: { color: '2A2A2C' },
  });
}

/** Smartphone, portrait. */
function phone(slide, x, y, w, h) {
  slide.addShape('roundRect', {
    x: x + 0.014 * w,
    y,
    w: 0.981 * w,
    h: 0.995 * h,
    rectRadius: 0.1 * w,
    fill: { color: '3A3A3C' },
  });
  slide.addShape('roundRect', {
    x: x + 0.025 * w,
    y: y + 0.004 * h,
    w: 0.959 * w,
    h: 0.987 * h,
    rectRadius: 0.09 * w,
    fill: { color: '0D0D0F' },
  });
  slide.addShape('rect', {
    x: x + 0.042 * w,
    y: y + 0.095 * h,
    w: 0.925 * w,
    h: 0.772 * h,
    fill: { color: '040404' },
  });
  slide.addShape('ellipse', {
    x: x + 0.485 * w,
    y: y + 0.026 * h,
    w: 0.03 * w,
    h: 0.03 * w,
    fill: { color: '2B2B2D' },
  });
  slide.addShape('roundRect', {
    x: x + 0.39 * w,
    y: y + 0.05 * h,
    w: 0.22 * w,
    h: 0.011 * h,
    rectRadius: 0.005 * h,
    fill: { color: '2B2B2D' },
  });
  slide.addShape('ellipse', {
    x: x + 0.435 * w,
    y: y + 0.892 * h,
    w: 0.13 * w,
    h: 0.13 * w,
    fill: { color: '1D1D1F' },
  });
  slide.addShape('roundRect', {
    x: x + 0.466 * w,
    y: y + 0.906 * h,
    w: 0.068 * w,
    h: 0.068 * w,
    rectRadius: 0.015 * w,
    fill: { color: '0D0D0F' },
  });
}

/** Flat-illustration monitor (slides 19 and 20). */
function flatMonitor(slide, x, y) {
  const W = 14.325;
  slide.addShape('roundRect', { x, y, w: W, h: 8.69, rectRadius: 0.16, fill: { color: '101010' } });
  slide.addShape('rect', { x, y: y + 0.16, w: W, h: 0.301, fill: { color: '1E1E1E' } });
  slide.addShape('roundRect', {
    x,
    y: y + 8.69,
    w: W,
    h: 1.284,
    rectRadius: 0.16,
    fill: { color: 'DFE4E5' },
  });
  slide.addShape('rect', { x, y: y + 8.69, w: W, h: 0.4, fill: { color: 'DFE4E5' } });
  slide.addShape('custGeom', {
    x: x + 2.436,
    y: y + 9.941,
    w: 4.806,
    h: 1.712,
    points: poly([
      [0.0, 1.712],
      [0.0, 1.547],
      [0.56, 1.481],
      [0.954, 0.0],
      [3.852, 0.0],
      [4.246, 1.481],
      [4.806, 1.547],
      [4.806, 1.712],
    ]),
    fill: { color: 'DFE4E5' },
  });
  slide.addShape('rect', {
    x: x + 3.358,
    y: y + 9.941,
    w: 2.996,
    h: 0.263,
    fill: { color: 'D8D8D8' },
  });
}

/** Flat-illustration laptop (slides 21 and 22). */
function flatLaptop(slide, x, y) {
  const W = 12.737;
  slide.addShape('roundRect', { x, y, w: W, h: 8.774, rectRadius: 0.2, fill: { color: '101010' } });
  slide.addShape('rect', { x, y: y + 0.2, w: W, h: 0.525, fill: { color: '1E1E1E' } });
  slide.addShape('ellipse', {
    x: x + 6.231,
    y: y + 0.216,
    w: 0.274,
    h: 0.274,
    fill: { color: '1E1E1E' },
  });
  slide.addShape('ellipse', {
    x: x + 6.251,
    y: y + 0.235,
    w: 0.235,
    h: 0.235,
    fill: { color: '101010' },
  });
  const bx = x - 1.998; // the base is wider than the lid on both sides
  const by = y + 8.324;
  slide.addShape('roundRect', {
    x: bx,
    y: by,
    w: 16.733,
    h: 0.685,
    rectRadius: 0.14,
    fill: { color: 'F0F0F0' },
  });
  slide.addShape('roundRect', {
    x: bx + 0.02,
    y: by + 0.353,
    w: 16.693,
    h: 0.333,
    rectRadius: 0.14,
    fill: { color: 'C9C9C9' },
  });
  slide.addShape('rect', { x: bx + 5.428, y: by, w: 5.879, h: 0.137, fill: { color: 'C9C9C9' } });
  [0, 0.392, 0.784].forEach((dx) => {
    slide.addShape('ellipse', {
      x: bx + 0.372 + dx,
      y: by + 0.059,
      w: 0.255,
      h: 0.255,
      fill: { color: 'C9C9C9' },
    });
    slide.addShape('ellipse', {
      x: bx + 0.411 + dx,
      y: by + 0.098,
      w: 0.176,
      h: 0.176,
      fill: { color: '676767' },
    });
  });
}

/** Flat-illustration phone (slides 23 and 24). */
function flatPhone(slide, x, y) {
  const W = 7.626;
  const H = 13.466;
  slide.addShape('roundRect', { x, y, w: W, h: H, rectRadius: 0.35, fill: { color: '101010' } });
  slide.addShape('rect', { x, y: y + 0.35, w: W, h: 0.515, fill: { color: '1E1E1E' } });
  slide.addShape('roundRect', {
    x: x + 2.965,
    y: y + 0.7,
    w: 1.647,
    h: 0.124,
    rectRadius: 0.062,
    fill: { color: '818181' },
  });
  slide.addShape('ellipse', {
    x: x + 3.5,
    y: y + 12.354,
    w: 0.824,
    h: 0.824,
    fill: { color: '232323' },
  });
  slide.addShape('ellipse', {
    x: x + 3.56,
    y: y + 12.414,
    w: 0.704,
    h: 0.704,
    fill: { color: '101010' },
  });
}

/** Round "drop a photo here" frame from the source layout of slide 25. */
function pictureFrame(slide, x, y, w, h, caption) {
  slide.addShape('ellipse', { x, y, w, h, fill: { color: 'F6F6F6' } });
  slide.addText(caption, {
    x,
    y: y + 0.05,
    w,
    h: 1.0,
    fontFace: BODY_FONT,
    fontSize: 54,
    color: '111111',
    align: 'center',
    valign: 'top',
  });
  slide.addText('[image]', {
    x: x + w / 2 - 1.2,
    y: y + h / 2 - 0.35,
    w: 2.4,
    h: 0.7,
    fontFace: BODY_FONT,
    fontSize: 20,
    color: 'A9AFB5',
    align: 'center',
    valign: 'middle',
  });
}

/* -------------------------------------------------- composite slide blocks */

/**
 * The recurring violet card: an optional oversized SCREEN / PRESENT headline,
 * a FEATURE CONTENT (or INFORMATION HERE) sub head and a body paragraph.
 */
function textPanel(slide, o) {
  panel(slide, o.panel[0], o.panel[1], o.panel[2], o.panel[3]);
  ['line1', 'line2', 'info'].forEach((key) => {
    const l = o[key];
    if (l) heading(slide, l.t, l.x, l.y, l.w, l.h, l.size, l.color);
  });
  if (o.feature) heading(slide, 'FEATURE CONTENT', o.feature[0], o.feature[1], 7.637, 0.841, 44);
  if (o.body) bodyText(slide, o.body.t, o.body.x, o.body.y, o.body.w, o.body.h, o.body);
}

/** SCREEN / PRESENT headline pair. */
const screenPresent = (sx, sy, px, py, o = {}) => ({
  line1: {
    t: o.t1 || 'SCREEN',
    x: sx,
    y: sy,
    w: o.w1 || 9.139,
    h: o.h1 || 2.878,
    size: o.s1 || 165,
  },
  line2: {
    t: o.t2 || 'PRESENT',
    x: px,
    y: py,
    w: o.w2 || 7.964,
    h: o.h2 || 2.121,
    size: o.s2 || 120,
  },
});

/** Full-width violet banner used on slides 33-38. */
function banner(slide, box, headAt, bodyBox, text, pad = true) {
  panel(slide, box[0], box[1], box[2], box[3]);
  heading(slide, 'INFORMATION HERE', headAt[0], headAt[1], 7.792, 0.69, 35);
  bodyText(slide, text, bodyBox[0], bodyBox[1], bodyBox[2], bodyBox[3], { pad });
}

/** Narrow violet strap line at the top of slides 25 and 26. */
function topBanner(slide) {
  panel(slide, -0.015, 0, 17.399, 2.08);
  heading(slide, 'INFORMATION HERE', 0.505, 0.454, 6.487, 0.69, 35);
  bodyText(slide, T.i, 0.402, 0.917, 16.982, 0.794, { align: 'justify' });
}

/* --------------------------------------------------------------- the deck */

const slides = [];

// 1 - iMac and laptop below a violet caption card
slides.push((s) => {
  iMac(s, 12.009, 3.093, 12.651, 10.619);
  macBook(s, 1.55, 7.111, 12.185, 6.639);
  textPanel(s, {
    panel: [-0.033, 2.25, 11.333, 3.944],
    feature: [0.811, 2.63],
    body: { t: T.a, x: 0.634, y: 3.38, w: 10.5, h: 2.37 },
  });
  headerTag(s);
});

// 2 - iMac left, full height violet panel right
slides.push((s) => {
  iMac(s, 0.967, 1.463, 15.234, 12.787);
  textPanel(s, {
    panel: [16.55, 1.379, 10.05, 12.455],
    ...screenPresent(16.884, 1.639, 17.086, 3.799),
    feature: [17.119, 5.986],
    body: { t: T.b, x: 17.108, y: 6.716, w: 8.025, h: 3.42 },
  });
  socialRow(s, 17.203, 10.447);
  headerTag(s);
});

// 3 - single iMac
slides.push((s) => {
  iMac(s, 6.066, 1.713, 15.234, 12.787);
  headerTag(s);
});

// 4 - iMac right, panel left
slides.push((s) => {
  iMac(s, 10.566, 1.796, 15.234, 12.787);
  textPanel(s, {
    panel: [-0.033, 2.25, 10.031, 10.5],
    ...screenPresent(0.578, 2.25, 0.669, 4.411),
    feature: [0.703, 6.597],
    body: { t: T.b, x: 0.692, y: 7.327, w: 8.025, h: 3.42 },
  });
  socialRow(s, 0.787, 11.058);
  headerTag(s);
});

// 5 - iMac and laptop, wide panel top right
slides.push((s) => {
  iMac(s, 2.015, 2.427, 12.651, 10.619);
  macBook(s, 12.949, 7.5, 12.185, 6.639);
  textPanel(s, {
    panel: [15.267, 0.879, 11.333, 6.018],
    line2: { t: 'PRESENT', x: 15.8, y: 0.963, w: 7.964, h: 2.121, size: 120 },
    feature: [15.91, 2.906],
    body: { t: T.b, x: 15.8, y: 3.605, w: 10.235, h: 2.895 },
  });
  headerTag(s);
});

// 6 - two iMacs below a full-width bar
slides.push((s) => {
  iMac(s, 1.717, 4.106, 11.202, 9.402);
  iMac(s, 13.8, 4.083, 11.202, 9.402);
  textPanel(s, {
    panel: [1.3, 0.879, 25.3, 2.704],
    info: { t: 'INFORMATION HERE', x: 1.82, y: 1.333, w: 6.487, h: 0.69, size: 35 },
    body: { t: T.a, x: 1.717, y: 1.892, w: 23.285, h: 1.319, align: 'justify' },
  });
});

// 7 - laptop left, panel right
slides.push((s) => {
  macBook(s, 0.967, 3.167, 17.844, 9.722);
  textPanel(s, {
    panel: [17.0, 0.667, 9.634, 8.869],
    ...screenPresent(17.078, 0.552, 17.28, 2.713),
    feature: [17.33, 4.847],
    body: { t: T.b, x: 17.225, y: 5.577, w: 8.909, h: 3.42 },
  });
  socialRow(s, 21.967, 10.031);
  headerTag(s);
});

// 8 - laptop right, panel left
slides.push((s) => {
  macBook(s, 8.75, 3.417, 17.844, 9.722);
  textPanel(s, {
    panel: [-0.033, 2.083, 10.031, 9.778],
    ...screenPresent(0.578, 2.083, 0.669, 4.244),
    feature: [0.703, 6.43],
    body: { t: T.b, x: 0.692, y: 7.161, w: 8.025, h: 3.42 },
  });
  socialRow(s, 0.787, 10.75);
  headerTag(s);
});

// 9 - iMac and tablet below a full-width bar
slides.push((s) => {
  iMac(s, 2.432, 4.106, 11.202, 9.402);
  tablet(s, 17.3, 3.833, 6.924, 9.372);
  textPanel(s, {
    panel: [1.3, 0.833, 25.3, 2.704],
    info: { t: 'INFORMATION HERE', x: 1.82, y: 1.287, w: 6.487, h: 0.69, size: 35 },
    body: { t: T.a, x: 1.717, y: 1.846, w: 23.285, h: 1.319, align: 'justify' },
  });
});

// 10 - tablet left, wide panel right
slides.push((s) => {
  tablet(s, 1.34, 1.239, 9.674, 13.094);
  textPanel(s, {
    panel: [11.718, 2.667, 14.883, 8.886],
    ...screenPresent(11.879, 2.583, 12.082, 4.744, { t2: 'PRESENTATION', w2: 14.135 }),
    feature: [12.132, 6.878],
    body: { t: T.c, x: 12.026, y: 7.639, w: 14.191, h: 3.945, pad: true },
  });
  socialRow(s, 11.789, 12.281);
  headerTag(s);
});

// 11 - tablet right, wide panel left
slides.push((s) => {
  tablet(s, 15.929, 0.798, 9.674, 13.094);
  textPanel(s, {
    panel: [-0.033, 2.25, 15.257, 9.0],
    ...screenPresent(0.578, 2.25, 0.669, 4.411),
    feature: [0.703, 6.597],
    body: { t: T.c, x: 0.563, y: 7.305, w: 14.154, h: 3.945, pad: true },
  });
  socialRow(s, 0.8, 11.854);
  headerTag(s);
});

// 12 - tablet turned on its side, tall panel right
slides.push((s) => {
  tablet(s, 1.3, 2.417, 14.5, 10.712, true);
  textPanel(s, {
    panel: [16.55, 1.379, 10.05, 11.038],
    ...screenPresent(16.884, 1.639, 17.086, 3.799),
    feature: [17.119, 5.986],
    body: { t: T.b, x: 17.108, y: 6.716, w: 8.025, h: 3.42 },
  });
  socialRow(s, 17.203, 10.447);
  headerTag(s);
});

// 13 - two tablets, small panel left
slides.push((s) => {
  tablet(s, 15.929, 0.798, 9.674, 13.094);
  tablet(s, 6.283, 0.836, 9.674, 13.094);
  socialRow(s, 0.527, 9.583);
  textPanel(s, {
    panel: [0, 1.25, 6.135, 7.833],
    info: { t: 'INFORMATION\nHERE', x: 0.524, y: 1.917, w: 5.49, h: 1.616, size: 45 },
    body: { t: T.a, x: 0.384, y: 3.505, w: 5.482, h: 4.995 },
  });
  headerTag(s);
});

// 14 - iMac, laptop and tablet cluster
slides.push((s) => {
  iMac(s, 1.384, 1.446, 10.976, 9.213);
  macBook(s, 6.847, 8.028, 12.185, 6.639);
  tablet(s, 2.863, 8.696, 4.267, 5.776);
  textPanel(s, {
    panel: [13.092, 1.0, 13.509, 6.686],
    ...screenPresent(13.253, 0.917, 13.217, 2.378, {
      h1: 1.784,
      s1: 100,
      t2: 'PRESENTATION',
      w2: 14.135,
      h2: 1.784,
      s2: 100,
    }),
    feature: [13.248, 4.135],
    body: { t: T.d, x: 13.092, y: 4.804, w: 12.959, h: 2.37 },
  });
  socialRow(s, 22.05, 8.114);
  headerTag(s);
});

// 15 - two phones right, panel left
slides.push((s) => {
  phone(s, 19.8, 3.204, 4.62, 9.72);
  phone(s, 13.431, 3.204, 4.62, 9.72);
  textPanel(s, {
    panel: [-0.033, 2.167, 13.13, 10.0],
    ...screenPresent(0.578, 2.167, 0.669, 4.327),
    feature: [0.703, 6.514],
    body: { t: T.c, x: 0.563, y: 7.167, w: 11.904, h: 4.47, pad: true },
  });
  socialRow(s, 0.634, 12.667);
  headerTag(s);
});

// 16 - tablet and phone left, panel right
slides.push((s) => {
  tablet(s, 2.634, 1.75, 8.741, 11.832);
  phone(s, 12.217, 3.895, 4.62, 9.72);
  textPanel(s, {
    panel: [17.17, 1.791, 9.431, 11.038],
    ...screenPresent(17.503, 2.051, 17.55, 3.75, { h1: 2.289, s1: 130 }),
    feature: [17.561, 6.25],
    body: { t: T.b, x: 17.55, y: 6.98, w: 8.025, h: 3.42 },
  });
  socialRow(s, 17.823, 10.859);
  headerTag(s);
});

// 17 - phone left, MOBILE PHONE PRESENTATION panel right
slides.push((s) => {
  phone(s, 2.024, 0.884, 6.48, 13.634);
  textPanel(s, {
    panel: [9.8, 2.045, 16.717, 8.869],
    ...screenPresent(10.417, 2.306, 10.417, 4.0, {
      t1: 'MOBILE PHONE',
      w1: 16.0,
      h1: 2.121,
      s1: 120,
      t2: 'PRESENTATION',
      w2: 14.417,
    }),
    feature: [10.344, 6.333],
    body: { t: T.e, x: 10.25, y: 7.164, w: 15.917, h: 3.42, pad: true },
  });
  socialRow(s, 9.913, 11.447);
  headerTag(s, 21.417, 0.419);
});

// 18 - mirror of slide 17
slides.push((s) => {
  phone(s, 18.866, 0.8, 6.48, 13.634);
  textPanel(s, {
    panel: [-0.033, 3.352, 16.717, 8.869],
    ...screenPresent(0.583, 3.612, 0.583, 5.307, {
      t1: 'MOBILE PHONE',
      w1: 16.0,
      h1: 2.121,
      s1: 120,
      t2: 'PRESENTATION',
      w2: 14.417,
    }),
    feature: [0.511, 7.64],
    body: { t: T.e, x: 0.417, y: 8.47, w: 15.917, h: 3.42, pad: true },
  });
  socialRow(s, 12.467, 12.697);
  headerTag(s);
});

// 19 - flat monitor, full height panel right
slides.push((s) => {
  flatMonitor(s, 1.226, 2.264);
  textPanel(s, {
    panel: [16.55, 1.379, 10.05, 12.455],
    ...screenPresent(16.884, 1.639, 17.086, 3.799),
    feature: [17.119, 5.986],
    body: { t: T.b, x: 17.108, y: 6.716, w: 8.025, h: 3.42 },
  });
  socialRow(s, 17.203, 10.447);
  headerTag(s);
});

// 20 - flat monitor right, panel left
slides.push((s) => {
  flatMonitor(s, 11.05, 2.167);
  textPanel(s, {
    panel: [0.016, 3.5, 10.05, 9.416],
    ...screenPresent(0.35, 3.76, 0.552, 5.921),
    feature: [0.586, 8.107],
    body: { t: T.b, x: 0.575, y: 8.837, w: 8.025, h: 3.42 },
  });
  socialRow(s, 5.817, 13.281);
  headerTag(s);
});

// 21 - flat laptop, panel right
slides.push((s) => {
  flatLaptop(s, 2.715, 3.833);
  textPanel(s, {
    panel: [16.55, 1.379, 10.05, 9.108],
    ...screenPresent(16.884, 1.639, 17.086, 3.799),
    feature: [17.119, 5.986],
    body: { t: T.b, x: 17.108, y: 6.716, w: 8.025, h: 3.42 },
  });
  socialRow(s, 21.855, 10.833);
  headerTag(s);
});

// 22 - flat laptop right, panel left
slides.push((s) => {
  flatLaptop(s, 11.298, 3.5);
  textPanel(s, {
    panel: [0.016, 1.917, 10.05, 9.416],
    ...screenPresent(0.35, 2.177, 0.552, 4.337),
    feature: [0.586, 6.524],
    body: { t: T.b, x: 0.575, y: 7.517, w: 9.238, h: 2.895 },
  });
  socialRow(s, 0.3, 11.781);
  headerTag(s);
});

// 23 - flat phone left, panel and three ticked notes right
slides.push((s) => {
  flatPhone(s, 2.467, 0.784);
  [
    { cx: 11.134, cy: 9.146, tx: 12.55, ty: 9.0, color: ACCENT5 },
    { cx: 11.134, cy: 10.901, tx: 12.55, ty: 10.755, color: ACCENT4 },
    { cx: 11.104, cy: 12.796, tx: 12.52, ty: 12.65, color: ACCENT3 },
  ].forEach((r) => checkRow(s, { ...r, tw: 11.25, text: T.check }));
  textPanel(s, {
    panel: [10.668, 1.111, 15.932, 7.008],
    info: { t: 'INFORMATION HERE', x: 11.561, y: 1.833, w: 12.489, h: 1.111, size: 60 },
    body: { t: T.f, x: 11.437, y: 2.833, w: 14.28, h: 4.995, pad: true },
  });
});

// 24 - mirror of slide 23
slides.push((s) => {
  flatPhone(s, 16.988, 0.861);
  [
    { cx: 3.035, cy: 11.593, tx: 4.452, ty: 11.447, color: ACCENT4 },
    { cx: 3.005, cy: 12.993, tx: 4.422, ty: 12.847, color: ACCENT3 },
  ].forEach((r) => checkRow(s, { ...r, tw: 11.25, text: T.check }));
  textPanel(s, {
    panel: [0, 4.906, 15.932, 5.167],
    info: { t: 'INFORMATION HERE', x: 0.893, y: 5.181, w: 12.489, h: 1.111, size: 60 },
    body: { t: T.g, x: 0.757, y: 6.222, w: 14.28, h: 3.945, pad: true },
  });
  headerTag(s);
});

// 25 - two round picture frames with a caption below
slides.push((s) => {
  pictureFrame(s, 2.884, 2.589, 9.25, 9.244, 'Drag  Your Picture Here');
  pictureFrame(s, 14.134, 2.589, 9.25, 9.244, 'Drag  Your Picture Here');
  heading(s, 'INFORMATION HERE', 3.939, 12.343, 7.792, 0.69, 35, ACCENT5);
  bodyText(s, T.h, 3.783, 12.931, 20.018, 1.319, { color: DIM_WHITE });
  topBanner(s);
});

// 26 - three empty picture slots with ticked notes on the right
slides.push((s) => {
  [
    { cx: 20.581, cy: 6.575, tx: 21.807, ty: 6.417, color: ACCENT5 },
    { cx: 20.599, cy: 7.95, tx: 21.744, ty: 7.792, color: ACCENT4 },
    { cx: 20.581, cy: 9.342, tx: 21.726, ty: 9.184, color: ACCENT3 },
  ].forEach((r) => checkRow(s, { ...r, tw: 4.094, text: T.checkShort }));
  topBanner(s);
});

// 27 - panel only, left hand side
slides.push((s) => {
  textPanel(s, {
    panel: [0.016, 0.667, 10.05, 8.495],
    ...screenPresent(0.35, 0.755, 0.552, 2.915),
    feature: [0.586, 4.995],
    body: { t: T.b, x: 0.467, y: 5.745, w: 9.238, h: 2.895 },
  });
  socialRow(s, 0.3, 10.359);
});

// 28 - panel with three ticked notes, right hand side
slides.push((s) => {
  [
    { cx: 12.682, cy: 9.319, tx: 14.099, ty: 9.173, color: ACCENT5 },
    { cx: 12.682, cy: 10.834, tx: 14.099, ty: 10.688, color: ACCENT4 },
    { cx: 12.652, cy: 12.313, tx: 14.069, ty: 12.167, color: ACCENT3 },
  ].forEach((r) => checkRow(s, { ...r, tw: 11.25, text: T.check }));
  textPanel(s, {
    panel: [12.217, 1.492, 14.384, 7.008],
    info: { t: 'INFORMATION HERE', x: 13.11, y: 1.766, w: 12.489, h: 1.111, size: 60 },
    body: { t: T.f, x: 12.986, y: 2.737, w: 12.363, h: 5.52, pad: true },
  });
  headerTag(s);
});

// 29 - full height panel on the right
slides.push((s) => {
  textPanel(s, {
    panel: [16.55, 1.379, 10.05, 12.455],
    ...screenPresent(16.884, 1.639, 17.086, 3.799),
    feature: [17.119, 5.986],
    body: { t: T.b, x: 17.108, y: 6.716, w: 8.025, h: 3.42 },
  });
  socialRow(s, 17.203, 10.447);
  headerTag(s);
});

// 30 - full height panel on the left
slides.push((s) => {
  textPanel(s, {
    panel: [0.039, 1.75, 10.05, 12.455],
    ...screenPresent(0.373, 2.01, 0.575, 4.171),
    feature: [0.609, 6.357],
    body: { t: T.b, x: 0.598, y: 7.087, w: 8.025, h: 3.42 },
  });
  socialRow(s, 0.692, 10.819);
  headerTag(s);
});

// 31 - as slide 29 but without the corner flag
slides.push((s) => {
  textPanel(s, {
    panel: [16.55, 1.417, 10.05, 12.455],
    ...screenPresent(16.884, 1.677, 17.086, 3.837),
    feature: [17.119, 6.024],
    body: { t: T.b, x: 17.108, y: 6.754, w: 8.025, h: 3.42 },
  });
  socialRow(s, 17.203, 10.485);
});

// 32 - headline panel above grey copy and two ticked notes
slides.push((s) => {
  bodyText(s, T.j, 11.494, 7.261, 14.28, 3.42, { color: DIM_WHITE });
  heading(s, 'INFORMATION HERE', 11.618, 6.745, 7.792, 0.69, 35, ACCENT5);
  [
    { cx: 11.577, cy: 11.396, tx: 12.994, ty: 11.25, color: ACCENT5 },
    { cx: 11.577, cy: 12.811, tx: 12.994, ty: 12.665, color: ACCENT4 },
  ].forEach((r) => checkRow(s, { ...r, tw: 11.25, text: T.check }));
  textPanel(s, {
    panel: [10.488, 1.333, 16.112, 4.667],
    ...screenPresent(11.161, 1.333, 11.248, 3.494, { t2: 'PRESENTATION', w2: 14.135 }),
  });
});

// 33-38 - full width banners at different heights on the page
slides.push((s) => {
  banner(s, [0.029, 2.583, 25.772, 4.667], [2.174, 3.0], [2.05, 3.6, 22.583, 3.42], T.f);
  headerTag(s);
});
slides.push((s) => {
  banner(s, [0, 1.639, 25.772, 4.667], [2.146, 2.056], [2.022, 2.656, 22.583, 3.42], T.f);
  headerTag(s);
});
slides.push((s) => {
  banner(s, [0, 9.25, 25.772, 4.667], [2.146, 9.667], [2.022, 10.266, 22.583, 3.42], T.f);
});
slides.push((s) => {
  banner(s, [0, 9.417, 25.772, 4.667], [2.146, 9.833], [2.022, 10.433, 22.583, 3.42], T.f);
  headerTag(s);
});
slides.push((s) => {
  banner(s, [0, 10.478, 22.467, 4.022], [0.924, 10.937], [0.8, 11.583, 21.083, 2.37], T.k, false);
  headerTag(s);
});
slides.push((s) => {
  banner(s, [0, 0.917, 25.772, 4.667], [2.146, 1.333], [2.022, 1.933, 22.583, 3.42], T.f);
});

/* ----------------------------------------------------------------- render */

const pres = new PptxGenJS();
pres.defineLayout({ name: 'DECK', width: SLIDE_W, height: SLIDE_H });
pres.layout = 'DECK';
pres.title = 'Screen Presentation';

slides.forEach((build) => build(newSlide(pres)));

pres
  .writeFile({
    fileName: path.join(__dirname, '0b8e3f34-7223-4d05-93df-74bb98f2e424_grok_final.pptx'),
  })
  .then((f) => console.log('wrote ' + f));
