#!/usr/bin/env node
/**
 * "Online Education" — 15-slide deck rebuilt with pptxgenjs.
 *
 * Slide size 13.333 x 7.5 in (16:9). Headings use Poppins, body copy Open Sans.
 * The source deck's picture placeholders are empty frames (no bitmaps), so they
 * are recreated as tinted `photoFrame()` shapes, and its icon artwork is redrawn
 * from native shapes -- no embedded raster data anywhere.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- design tokens

const C = {
  blue: '2964E9',
  red: 'FD493F',
  yellow: 'FACC16',
  black: '000000',
  white: 'FFFFFF',
  paper: 'FEFDFB', // deck background
  ink: '262626', // dark icon strokes
  frame: '4472C4', // picture-placeholder tint, painted at 5% density (see photoFrame)
};

// Text-box insets used throughout the source deck: 0.1" sides, 0.05" top/bottom.
const INSET = [7.2, 7.2, 3.6, 3.6];

const FONT = { head: 'Poppins', body: 'Open Sans' };

const SLIDE_W = 13.3333333;
const SLIDE_H = 7.5;

// Radius fractions taken from the source shapes' `adj` values.
const R_CARD = 0.11143; // rounded content cards
const R_PILL = 0.5; // fully rounded buttons
const R_PHOTO = 0.13233; // rounded picture frames

const LOREM = {
  full:
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eget elit ut neque ' +
    'vulputate hendrerit id vel nibh. Donec commodo felis id ante facilisis consectetur. ' +
    'Ut malesuada eget mauris vitae consequat. Proin commodo lorem quis nisi laoreet ' +
    'euismod. Morbi malesuada eu diam quis tempor.',
  noTempor:
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eget elit ut neque ' +
    'vulputate hendrerit id vel nibh. Donec commodo felis id ante facilisis consectetur. ' +
    'Ut malesuada eget mauris vitae consequat. Proin commodo lorem quis nisi laoreet euismod. ',
  noPeriod:
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eget elit ut neque ' +
    'vulputate hendrerit id vel nibh. Donec commodo felis id ante facilisis consectetur. ' +
    'Ut malesuada eget mauris vitae consequat. Proin commodo lorem quis nisi laoreet euismod',
  noEuismod:
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eget elit ut neque ' +
    'vulputate hendrerit id vel nibh. Donec commodo felis id ante facilisis consectetur. ' +
    'Ut malesuada eget mauris vitae consequat. Proin commodo lorem quis nisi laoreet.',
  short:
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eget elit ut neque ' +
    'vulputate hendrerit id vel nibh. Donec commodo felis id ante facilisis consectetur. ',
  tiny: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam eget elit ut neque vulputate hendrerit id vel nibh.',
  mini: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.',
};

// ---------------------------------------------------------------- shape helpers

/** New slide on the deck's off-white paper background. */
function newSlide(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.paper };
  return s;
}

/** Flat colour block. */
function rect(slide, x, y, w, h, color) {
  slide.addShape('rect', { x, y, w, h, fill: { color } });
}

/** Rounded block; `frac` is the corner radius as a fraction of the short side. */
function roundBlock(slide, x, y, w, h, color, frac = R_CARD) {
  slide.addShape('roundRect', {
    x, y, w, h,
    fill: { color },
    rectRadius: frac * Math.min(w, h),
  });
}

/**
 * Picture placeholder from the source deck. The originals hold no bitmap -- they
 * are empty frames washed with a 5%-density accent pattern, reproduced here as a
 * near-transparent tint so the silhouette reads without hiding what sits behind.
 */
function photoFrame(slide, x, y, w, h, shape = 'rect', frac = null) {
  const opts = { x, y, w, h, fill: { color: C.frame, transparency: 97 } };
  if (frac) opts.rectRadius = frac * Math.min(w, h);
  slide.addShape(shape, opts);
}

/** Poppins display heading; one paragraph per array entry. */
function heading(slide, lines, x, y, w, h, { size = 48, color = C.red, align = 'left' } = {}) {
  const runs = [].concat(lines).map((t, i, all) => ({
    text: t,
    options: { breakLine: i < all.length - 1 },
  }));
  slide.addText(runs, {
    x, y, w, h,
    fontFace: FONT.head, fontSize: size, color,
    align, valign: 'top', wrap: false, margin: INSET, fit: 'resize',
  });
}

/** Open Sans paragraph. */
function bodyText(slide, text, x, y, w, h, { size = 14, color = C.black, align = 'justify', bold = false } = {}) {
  slide.addText(text, {
    x, y, w, h,
    fontFace: FONT.body, fontSize: size, color, bold,
    align, valign: 'top', wrap: true, margin: INSET, fit: 'resize',
  });
}

/** Circle with a centred Poppins number, e.g. "01". */
function numberBadge(slide, x, y, w, h, fill, text) {
  slide.addShape('ellipse', { x, y, w, h, fill: { color: fill } });
  slide.addText(text, {
    x, y, w, h,
    fontFace: FONT.head, fontSize: 32, color: C.white,
    align: 'center', valign: 'middle',
  });
}

// ---------------------------------------------------------------- icon library
//
// The source deck's icons are Office SVG glyphs drawn on a 96 x 96 grid; each is
// rebuilt here from that same grid as a list of native shapes:
//   [shape, x, y, w, h, opts?]   opts: {o} outline only, {v} knock-out (bg colour),
//                                      {r} rotation, {rr} corner radius, {a} angleRange
const ICONS = {
  // open book resting on a lectern
  openBook: [
    ['rect', 4, 71, 88, 7], ['rect', 4, 22, 6, 56], ['rect', 86, 22, 6, 56],
    ['rect', 39, 78, 18, 3],
    ['rect', 14, 16, 68, 6], ['rect', 14, 61, 68, 6],
    ['rect', 14, 16, 6, 51], ['rect', 76, 16, 6, 51], ['rect', 46, 16, 4, 45],
    ['rect', 56, 32, 14, 3], ['rect', 56, 38, 14, 3], ['rect', 56, 44, 10, 3],
  ],
  // rolled diploma: tube seen side-on, curled at the left, wax seal and ribbon
  diploma: [
    ['can', 30, -5, 34, 82, { o: true, r: 90 }], // upright can, quarter-turned onto its side
    ['donut', 12, 32, 18, 18, { o: true }], // curled end of the sheet
    ['rect', 40, 50, 20, 24], ['triangle', 40, 66, 20, 8, { v: true }],
    ['star12', 36, 22, 28, 28],
    ['ellipse', 41, 27, 18, 18, { v: true }], ['ellipse', 41, 27, 18, 18, { o: true }],
  ],
  // teacher at a whiteboard facing three students
  classroom: [
    ['rect', 24, 10, 57, 4], ['rect', 24, 47, 57, 4],
    ['rect', 24, 10, 4, 41], ['rect', 77, 10, 4, 41],
    ['ellipse', 15, 20, 13, 13],
    ['roundRect', 12, 34, 18, 30, { rr: 5 }], ['rect', 13, 60, 6, 27], ['rect', 23, 60, 6, 27],
    ['rect', 30, 32, 30, 4, { r: 318 }], // pointer, hand to board
    ['ellipse', 36, 63, 10, 10], ['pie', 32, 73, 18, 20, { a: [180, 360] }],
    ['ellipse', 56, 63, 10, 10], ['pie', 52, 73, 18, 20, { a: [180, 360] }],
    ['ellipse', 76, 63, 10, 10], ['pie', 72, 73, 18, 20, { a: [180, 360] }],
  ],
  // row of books on a shelf
  books: [
    ['rect', 14, 76, 68, 6],
    ['rect', 14, 14, 8, 58],
    ['rect', 26, 20, 14, 52], ['rect', 30, 24, 6, 6, { v: true }],
    ['rect', 44, 28, 16, 44],
    ['ellipse', 50, 32, 4, 4, { v: true }], ['ellipse', 50, 64, 4, 4, { v: true }],
    ['rect', 64, 18, 18, 8], ['rect', 64, 30, 18, 30], ['rect', 64, 64, 18, 8],
  ],
  // graduation cap
  gradCap: [
    ['round2SameRect', 20, 52, 56, 20, { r: 180, rr: 10 }],
    ['diamond', 4, 24, 88, 31],
    ['rect', 10, 40, 4, 22],
  ],
  // desk telephone
  phone: [
    ['blockArc', 9, 19, 78, 50, { a: [180, 360] }],
    ['trapezoid', 18, 41, 60, 35],
    ['rect', 32, 32, 32, 10],
    ['rect', 38, 48, 4, 4, { v: true }], ['rect', 46, 48, 4, 4, { v: true }], ['rect', 54, 48, 4, 4, { v: true }],
    ['rect', 38, 56, 4, 4, { v: true }], ['rect', 46, 56, 4, 4, { v: true }], ['rect', 54, 56, 4, 4, { v: true }],
    ['rect', 38, 64, 4, 4, { v: true }], ['rect', 46, 64, 4, 4, { v: true }], ['rect', 54, 64, 4, 4, { v: true }],
  ],
  // laptop showing a globe
  laptop: [
    ['rect', 14, 20, 68, 6], ['rect', 14, 60, 68, 6],
    ['rect', 14, 20, 6, 46], ['rect', 76, 20, 6, 46],
    ['roundRect', 2, 70, 92, 6, { rr: 3 }],
    ['ellipse', 34, 29, 28, 28, { o: true }],
    ['ellipse', 42, 29, 12, 28, { o: true }],
    ['rect', 34, 42, 28, 2],
  ],
  // envelope: box plus the V-fold and the two flap creases
  envelope: [
    ['rect', 8, 20, 80, 6], ['rect', 8, 70, 80, 6],
    ['rect', 8, 20, 6, 56], ['rect', 82, 20, 6, 56],
    ['rect', 9, 38, 45, 5, { r: 40 }], ['rect', 42, 38, 45, 5, { r: 320 }],
    ['rect', 10, 57, 30, 5, { r: 318 }], ['rect', 56, 57, 30, 5, { r: 42 }],
  ],
};

const ICON_GRID = 96;

/** Draw an ICONS entry scaled into a `size` x `size` box at (x, y). */
function icon(slide, name, x, y, size, color, bg) {
  const u = size / ICON_GRID;
  ICONS[name].forEach(([shape, ix, iy, iw, ih, o = {}]) => {
    const opts = { x: x + ix * u, y: y + iy * u, w: iw * u, h: ih * u };
    if (o.o) opts.line = { color, width: 4 * u * 72 };
    else opts.fill = { color: o.v ? bg : color };
    if (o.r) opts.rotate = o.r;
    if (o.rr) opts.rectRadius = o.rr * u;
    if (o.a) opts.angleRange = o.a;
    slide.addShape(shape, opts);
  });
}

/**
 * The bookend stack shared by the opening and closing slides: two 80pt words in
 * blue then red, a 24pt strapline and a yellow pill button. `dx`/`dy` shift the
 * whole stack (slide 15 repeats it 7.034" to the right and 0.048" up).
 */
function titleStack(slide, dx, dy, { line1, w1, line2, w2, strap, button, buttonW, labelW, arrow }) {
  heading(slide, line1, dx + 0.989, dy + 1.237, w1, 1.447, { size: 80, color: C.blue });
  heading(slide, line2, dx + 0.989, dy + 2.181, w2, 1.447, { size: 80, color: C.red });
  bodyText(slide, strap, dx + 1.019, dy + 3.396, 4.343, 0.505, { size: 24, align: 'left' });
  roundBlock(slide, dx + 1.109, dy + 4.01, buttonW, 0.66, C.yellow, R_PILL);
  bodyText(slide, button, dx + 1.326, dy + 4.121, labelW, 0.438, { size: 20, align: 'left' });
  if (arrow) {
    slide.addShape('rightArrow', { x: dx + 3.901, y: dy + 4.201, w: 0.285, h: 0.278, fill: { color: C.black } });
  }
}

/** The four picture frames used as a 2 x 2 decorative cluster on slides 1 and 15. */
function frameCluster(slide, x, y) {
  photoFrame(slide, x, y, 2.243, 2.243, 'ellipse');
  photoFrame(slide, x + 2.42, y, 2.243, 2.243, 'teardrop');
  photoFrame(slide, x + 2.42, y + 2.426, 2.243, 2.243, 'round2SameRect');
  photoFrame(slide, x, y + 2.417, 2.243, 2.243, 'round2DiagRect');
}

// ---------------------------------------------------------------- slide builders

// 1 — Title slide
function slide01(pptx) {
  const s = newSlide(pptx);
  rect(s, 0, 5.635, 13.333, 1.865, C.yellow);
  titleStack(s, 0, 0, {
    line1: 'Online', w1: 3.948,
    line2: 'Education', w2: 6.026,
    strap: 'Learning Without Limits',
    button: 'Start Presentation', buttonW: 3.321, labelW: 2.668, arrow: true,
  });
  frameCluster(s, 7.696, 0.662);
}

// 2 — Introducing / speaker
function slide02(pptx) {
  const s = newSlide(pptx);
  photoFrame(s, 1.074, 2.141, 4.422, 4.422, 'roundRect', 0.16667);
  rect(s, 6.627, 0, 6.707, 7.5, C.blue);
  heading(s, 'Introducing', 1.074, 0.899, 4.776, 1.01, { size: 54 });
  bodyText(s, LOREM.noPeriod, 7.375, 5.275, 5.531, 1.279, { color: C.white });
  roundBlock(s, 1.543, 5.313, 3.484, 0.858, C.yellow, R_PILL);
  bodyText(s, 'Michelle Halim', 1.918, 5.438, 2.734, 0.37, { size: 16, bold: true, align: 'center' });
  bodyText(s, 'CEO of Your Brand', 1.927, 5.696, 2.734, 0.37, { size: 16, align: 'center' });
  photoFrame(s, 7.496, 0.899, 5.29, 4.185, 'round2SameRect');
}

// 3 — Why Online Education Matters?
function slide03(pptx) {
  const s = newSlide(pptx);
  rect(s, 9.289, 0, 4.044, 7.5, C.red);
  photoFrame(s, 6.091, 0.596, 6.526, 6.309, 'roundRect', R_PHOTO);
  heading(s, ['Why Online', 'Education', 'Matters?'], 1.03, 1.853, 4.181, 2.524);
  bodyText(s, LOREM.short, 1.03, 4.378, 4.172, 1.043);

  const cards = [
    { x: 5.502, icon: 'openBook', ix: 6.141, iy: 4.657, is: 0.843, tx: 5.605, label: 'Lorem Ipsum Dolor Sit Ametc' },
    { x: 7.768, icon: 'diploma', ix: 8.332, iy: 4.583, is: 1.0, tx: 7.873, label: 'Lorem Ipsum Dolor Sit Amet' },
    { x: 10.035, icon: 'classroom', ix: 10.681, iy: 4.663, is: 0.829, tx: 10.132, label: 'Lorem Ipsum Dolor Sit Amet' },
  ];
  cards.forEach((c) => roundBlock(s, c.x, 4.474, 2.122, 1.939, C.yellow));
  cards.forEach((c) => icon(s, c.icon, c.ix, c.iy, c.is, C.ink, C.yellow));
  cards.forEach((c) => bodyText(s, c.label, c.tx, 5.493, 1.928, 0.64, { size: 16, bold: true, align: 'center' }));
}

// 4 — Key Benefits (three stacked bars)
function slide04(pptx) {
  const s = newSlide(pptx);
  heading(s, ['Key', 'Benefits'], 8.414, 2.265, 3.019, 1.717);

  const rows = [
    { y: 0.746, bar: C.yellow, dot: C.red, n: '01' },
    { y: 2.851, bar: C.red, dot: C.yellow, n: '02' },
    { y: 4.958, bar: C.yellow, dot: C.red, n: '03' },
  ];
  rows.forEach((r) => {
    roundBlock(s, 1.211, r.y, 6.294, 1.939, r.bar);
    bodyText(s, LOREM.noTempor, 1.96, r.y + 0.33, 5.214, 1.279);
    numberBadge(s, 0.773, r.y + 0.484, 0.989, 0.971, r.dot, r.n);
  });

  bodyText(s, LOREM.short, 8.414, 3.959, 4.146, 1.043);
}

// 5 — Education Platforms (blue panel + three step cards)
function slide05(pptx) {
  const s = newSlide(pptx);
  rect(s, -0.04, 0, 6.707, 7.5, C.blue);
  heading(s, ['Education', 'Platforms'], 1.176, 2.033, 3.698, 1.717, { color: C.white });
  bodyText(s, LOREM.full, 1.176, 3.737, 4.172, 1.986, { color: C.white });

  // card origin -> badge / title / copy are all placed at fixed offsets
  const cards = [
    { x: 7.479, y: 1.296, n: '01', title: 'Platform 01' },
    { x: 10.39, y: 2.941, n: '02', title: 'Platform 02' },
    { x: 7.479, y: 4.596, n: '03', title: 'Platform 03' },
  ];
  cards.forEach((c) => {
    roundBlock(s, c.x, c.y, 2.354, 2.093, C.yellow);
    numberBadge(s, c.x + 0.694, c.y - 0.485, 0.989, 0.971, C.red, c.n);
    bodyText(s, c.title, c.x + 0.213, c.y + 0.581, 1.928, 0.404, { size: 18, bold: true, align: 'center' });
    bodyText(s, LOREM.mini, c.x + 0.111, c.y + 0.945, 2.154, 0.808, { align: 'center' });
  });
}

// 6 — Interactive Learning
function slide06(pptx) {
  const s = newSlide(pptx);
  rect(s, 0, 5.635, 13.333, 1.865, C.yellow);
  heading(s, ['Interactive', 'Learning'], 1.343, 1.591, 4.003, 1.717);
  bodyText(s, LOREM.full, 1.343, 3.294, 4.172, 1.986);
  photoFrame(s, 6.478, 0.589, 6.278, 6.322, 'round2SameRect');
}

// 7 — Learning Styles
function slide07(pptx) {
  const s = newSlide(pptx);
  heading(s, ['Learning', 'Styles'], 1.154, 2.047, 3.247, 1.717);
  bodyText(s, LOREM.full, 1.154, 3.75, 4.457, 1.75);
  rect(s, 9.289, 0, 4.044, 7.5, C.red);
  photoFrame(s, 6.224, 0.711, 6.526, 3.017, 'roundRect', R_PHOTO);
  photoFrame(s, 6.224, 3.846, 6.526, 3.017, 'roundRect', R_PHOTO);
}

// 8 — Online Education (full-width banner)
function slide08(pptx) {
  const s = newSlide(pptx);
  heading(s, ['Online', 'Education'], 1.143, 5.242, 3.698, 1.717);
  bodyText(s, LOREM.full, 5.175, 5.664, 7.38, 1.043);
  rect(s, 0, 0, 13.333, 1.4, C.blue);
  photoFrame(s, 0, 1.4, 13.333, 3.344);
}

// 9 — Traditional Education
function slide09(pptx) {
  const s = newSlide(pptx);
  rect(s, 6.085, 4.398, 7.248, 1.939, C.yellow);
  heading(s, ['Traditional', 'Education'], 0.976, 1.384, 4.011, 1.717);
  bodyText(s, LOREM.full, 6.553, 4.728, 6.214, 1.279);
  photoFrame(s, 0.976, 3.467, 4.768, 3.017, 'roundRect', R_PHOTO);
  photoFrame(s, 6.296, 0.802, 3.104, 3.104, 'ellipse');
  photoFrame(s, 9.563, 0.802, 3.104, 3.104, 'ellipse');
}

// 10 — Challenges (two teardrop cards)
function slide10(pptx) {
  const s = newSlide(pptx);
  photoFrame(s, 7.489, 0, 5.844, 7.5);
  heading(s, 'Challenges', 1.165, 1.462, 4.141, 0.909);

  const cards = [
    { x: 1.165, title: 'Challenges 01', icon: 'books', ix: 1.395, is: 0.858 },
    { x: 5.559, title: 'Challenges 02', icon: 'gradCap', ix: 5.755, is: 0.952 },
  ];
  cards.forEach((c) => {
    roundBlock(s, c.x, 2.741, 4.035, 3.687, C.yellow);
    s.addShape('teardrop', { x: c.x, y: 2.741, w: 1.379, h: 1.354, fill: { color: C.red }, flipH: true });
    bodyText(s, c.title, c.x + 0.659, 4.389, 2.576, 0.438, { size: 20, bold: true, align: 'center' });
    bodyText(s, LOREM.tiny, c.x + 0.181, 4.864, 3.554, 1.043, { align: 'center' });
  });
  cards.forEach((c) => icon(s, c.icon, c.ix, 2.918, c.is, C.paper, C.red));
}

// 11 — Break slide
function slide11(pptx) {
  const s = newSlide(pptx);
  photoFrame(s, 0, 0, 13.333, 4.556);
  heading(s, 'Break Slide', 1.065, 5.037, 7.253, 1.582, { size: 88 });
  rect(s, 8.867, 4.556, 4.467, 2.944, C.yellow);
  rect(s, 0, 0, 4.467, 2.311, C.blue);
}

// 12 — Overcoming Challenges
function slide12(pptx) {
  const s = newSlide(pptx);
  rect(s, 0, 5.75, 13.333, 1.75, C.yellow);
  heading(s, ['Overcoming', 'Challenges'], 0.936, 1.685, 4.514, 1.717);
  bodyText(s, LOREM.full, 0.936, 3.389, 4.613, 1.75);
  photoFrame(s, 6.244, 0.663, 3.104, 3.104, 'ellipse');
  photoFrame(s, 9.533, 0.663, 3.104, 3.104, 'ellipse');
  photoFrame(s, 6.244, 3.956, 6.393, 3.017, 'roundRect', R_PHOTO);
}

// 13 — Conclusion (three colour cards)
function slide13(pptx) {
  const s = newSlide(pptx);
  heading(s, 'Conclusion', 4.622, 1.068, 4.115, 0.909, { align: 'center' });

  //          card x   fill      title                 title x/y      body x/y
  const cards = [
    { x: 0.487, color: C.yellow, title: 'Conclusion 01', tx: 1.217, ty: 3.565, bx: 0.787, by: 4.061 },
    { x: 4.702, color: C.red, title: 'Conclusion 02', tx: 5.431, ty: 3.565, bx: 5.002, by: 4.061 },
    { x: 8.917, color: C.blue, title: 'Conclusion 03', tx: 9.687, ty: 3.506, bx: 9.257, by: 4.002 },
  ];
  cards.forEach((c) => roundBlock(s, c.x, 3.249, 4.035, 3.183, c.color));
  bodyText(s, LOREM.full, 2.039, 1.977, 9.361, 0.808, { align: 'center' });
  cards.forEach((c) => {
    bodyText(s, LOREM.noEuismod, c.bx, c.by, 3.436, 1.986, { align: 'center' });
    bodyText(s, c.title, c.tx, c.ty, 2.576, 0.438, { size: 20, bold: true, align: 'center' });
  });
}

// 14 — Our Contact
function slide14(pptx) {
  const s = newSlide(pptx);
  rect(s, 6.65, 3.75, 6.683, 3.75, C.red);
  rect(s, 6.65, 0, 6.683, 3.75, C.yellow);
  heading(s, 'Our Contact', 1.343, 2.658, 4.423, 0.909);

  const rows = [
    { text: '+123-4567-890', w: 3.898, y: 3.614, icon: 'phone', ix: 1.343, iy: 3.567, is: 0.5 },
    { text: 'www.yourwebsite.com', w: 4.376, y: 4.074, icon: 'laptop', ix: 1.343, iy: 4.022, is: 0.5 },
    { text: 'name@yourwebsite.com', w: 4.647, y: 4.548, icon: 'envelope', ix: 1.366, iy: 4.522, is: 0.454 },
  ];
  rows.forEach((r) => {
    bodyText(s, r.text, 1.925, r.y, r.w, 0.404, { size: 18, align: 'left' });
    icon(s, r.icon, r.ix, r.iy, r.is, C.blue, C.paper);
  });

  photoFrame(s, 6.65, 0.6, 5.916, 6.3, 'round2DiagRect');
}

// 15 — Thank You (slide 1's stack and frame cluster, mirrored left to right)
function slide15(pptx) {
  const s = newSlide(pptx);
  rect(s, 0, 5.635, 13.333, 1.865, C.yellow);
  titleStack(s, 7.034, -0.048, {
    line1: 'Thank ', w1: 4.071,
    line2: 'You!', w2: 2.865,
    strap: 'For Your Attention',
    button: 'www.yourwebsite.com', buttonW: 3.613, labelW: 3.306,
  });
  frameCluster(s, 1.682, 0.651);
}

const SLIDES = [
  slide01, slide02, slide03, slide04, slide05,
  slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15,
];

// ---------------------------------------------------------------- build

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'REF', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'REF';
  pptx.title = 'Online Education';

  SLIDES.forEach((fn) => fn(pptx));

  const out = path.join(__dirname, '11f1bf71-2bbb-4bd0-91c1-b22ccb7d97a5_grok_final.pptx');
  return pptx.writeFile({ fileName: out }).then(() => console.log('wrote', out));
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
