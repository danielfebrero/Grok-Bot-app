/**
 * "MEDICAL PITCH DECK" — 20-slide deck rebuilt with pptxgenjs.
 * Slide size 10 x 5.625 in (16:9). Raster photos in the source deck are
 * replaced by programmatic placeholders (grey boxes / device outlines).
 *
 *   node 084de3d2-fb08-412f-87c1-ac19e0fc217d_grok_final.js
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Palette, fonts, shared metrics
 * ------------------------------------------------------------------ */

const TEAL = '008080'; // theme accent1
const TEAL_D = '006060'; // darker teal (cards, funnel, gridlines)
const TEAL_DD = '004040'; // darkest teal (funnel top)
const CYAN = 'B2FFFF'; // pale cyan band / progress ring
const WHITE = 'FFFFFF';
const GREY_TX = '262626'; // body copy on white
const GREY_TX2 = '3F3F3F'; // body copy, alt slides
const PH_BG = 'F2F2F2'; // image placeholder fill
const PH_ICON = '0A0A0A'; // image placeholder glyph
const GREY1 = '7F7F7F';
const GREY2 = 'BFBFBF';
const GREY3 = 'D9D9D9';
const GREY4 = 'D8D8D8';
const DEVICE = '0A0A0A'; // phone / laptop bezel
const DEVICE_2 = '4D4D4D'; // laptop base

const BEBAS = 'Bebas Neue';
const LIGHT = 'Roboto Light';
const SEMI = 'Roboto SemiBold';

// Text-box insets in points [left, right, bottom, top] — the source deck uses 0.075" / 0.0375".
const INSET = [5.4, 5.4, 2.7, 2.7];
const NOINSET = [0, 0, 0, 0];

const LOREM =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ' +
  'ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco ' +
  'laboris nisi ut aliquip ex ea commodo consequat.\u00a0';
const LOREM_SHORT =
  'Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea ' +
  'commodo consequat.\u00a0';
const LOREM_CARD =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ' +
  'ut labore et dolore magna aliqua. Ut enim ad minim.';
const LOREM_TINY = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor';

/* Custom geometries copied from the source deck (path units -> normalised 0..1). */
const GEOM = {
  // right-pointing arrow used by every "Learn more" link
  arrow: {
    w: 877993,
    h: 400205,
    pts: [
      [677890, 0], [877993, 200103], [677890, 400205], [642248, 365187],
      [782319, 225116], [781685, 225116], [781685, 225580], [0, 225580],
      [0, 174780], [781685, 174780], [781685, 175090], [782319, 175090],
      [642248, 35018],
    ],
  },
  // tick mark inside the round bullet badges
  check: {
    w: 155257,
    h: 114538,
    pts: [
      [54293, 114538], [0, 60246], [13573, 46673], [54293, 87392],
      [141684, 0], [155258, 13573],
    ],
  },
  // slanted tail of the header ribbon
  ribbonTail: {
    w: 1368955,
    h: 250506,
    pts: [[0, 0], [1368955, 0], [1163638, 250506], [0, 250506]],
  },
  // one band of the go-to-market funnel
  funnel: {
    w: 4809363,
    h: 1537430,
    pts: [
      [2404682, 0], [0, 0], [334137, 1139762], [2404682, 1537430],
      [4475131, 1139762], [4809363, 0],
    ],
  },
};

/* ------------------------------------------------------------------ *
 * Low-level helpers
 * ------------------------------------------------------------------ */

function rect(slide, x, y, w, h, fill, extra) {
  slide.addShape('rect', Object.assign({ x, y, w, h, fill: { color: fill }, line: { type: 'none' } }, extra));
}

/** Draw one of the GEOM custom paths, scaled into the box x/y/w/h. */
function poly(slide, geom, x, y, w, h, fill) {
  const pts = geom.pts.map(function (p) {
    return { x: (p[0] / geom.w) * w, y: (p[1] / geom.h) * h };
  });
  pts.push({ close: true });
  slide.addShape('custGeom', { x, y, w, h, points: pts, fill: { color: fill }, line: { type: 'none' } });
}

function text(slide, content, o) {
  slide.addText(content, {
    x: o.x,
    y: o.y,
    w: o.w,
    h: o.h,
    fontFace: o.face || LIGHT,
    fontSize: o.size,
    color: o.color,
    bold: !!o.bold,
    align: o.align || 'left',
    valign: o.valign || 'top',
    margin: o.margin || INSET,
    lineSpacingMultiple: o.line,
    rotate: o.rotate,
    isTextBox: true,
    wrap: true,
    paraSpaceBefore: 0,
    paraSpaceAfter: 0,
  });
}

/** Bebas Neue headline (the deck renders every Bebas run in caps). */
function head(slide, str, o) {
  text(slide, str, Object.assign({ face: BEBAS, color: TEAL }, o));
}

/** 9 pt Roboto Light body copy at 150 % leading. */
function body(slide, content, o) {
  text(slide, content, Object.assign({ face: LIGHT, size: 9, color: GREY_TX, line: 1.5 }, o));
}

/** Two or three paragraphs separated by a blank line. */
function paras(list) {
  const out = [];
  list.forEach(function (p, i) {
    out.push({ text: p, options: { breakLine: i < list.length - 1 } });
  });
  return out;
}

/* ------------------------------------------------------------------ *
 * Repeated design components
 * ------------------------------------------------------------------ */

/** Arrow + "Learn more" caption. `x`,`y` = top-left of the 1.374 x 0.328 group. */
function learnMore(slide, x, y, color) {
  poly(slide, GEOM.arrow, x, y + 0.073, 0.424, 0.193, color);
  head(slide, 'Learn more', { x: x + 0.424, y, w: 0.95, h: 0.328, size: 15, color });
}

/**
 * Header ribbon: hairline rule, tag block and slanted tail.
 * `dark` = ribbon painted in teal (default is the white variant).
 */
function ribbon(slide, opts) {
  const o = opts || {};
  const bar = o.dark ? TEAL : WHITE;
  const label = o.dark ? WHITE : TEAL;
  rect(slide, 0.275, 0.265, 9.451, 0.037, bar);
  poly(slide, GEOM.ribbonTail, 0.783, 0.279, 1.123, 0.205, bar);
  rect(slide, 0.275, 0.265, 0.917, 0.22, bar);
  text(slide, 'MEDICAL PITCH DECK', {
    x: 0.309, y: 0.27, w: 1.436, h: 0.227, face: SEMI, size: 9, color: label,
  });
  if (o.page) {
    text(slide, 'Page ' + o.page, {
      x: 8.776, y: 0.336, w: 0.95, h: 0.202, face: BEBAS, size: 8, color: bar, align: 'right', line: 1,
    });
  }
}

/** Round badge with a white tick — used for the bulleted feature lists. */
function checkBadge(slide, x, y, d, ring, tick) {
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color: ring }, line: { type: 'none' } });
  poly(slide, GEOM.check, x + d * 0.125, y + d * 0.2225, d * 0.752, d * 0.5549, tick);
}

/**
 * Grey box standing in for a photograph, with the same "no image" glyph the source
 * deck uses (a tilted card behind a picture frame). The glyph scales with the box the
 * way the original bitmap does, i.e. off the covered 16:9 area.
 */
function photo(slide, x, y, w, h) {
  rect(slide, x, y, w, h, PH_BG);

  const d = 0.1068 * Math.max(h, (w * 9) / 16); // glyph size, from the source bitmap
  const cx = x + w / 2;
  const cy = y + h / 2;
  slide.addShape('roundRect', { // card tucked in behind
    x: cx - d * 0.5, y: cy - d * 0.28, w: d * 0.86, h: d * 0.66, rectRadius: d * 0.1,
    rotate: 352, fill: { color: PH_ICON }, line: { type: 'none' },
  });
  slide.addShape('roundRect', { // picture frame
    x: cx - d * 0.34, y: cy - d * 0.46, w: d * 0.8, h: d * 0.66, rectRadius: d * 0.1,
    fill: { color: PH_BG }, line: { color: PH_ICON, width: d * 8 },
  });
  slide.addShape('ellipse', { // sun
    x: cx - d * 0.22, y: cy - d * 0.34, w: d * 0.13, h: d * 0.13,
    fill: { color: PH_ICON }, line: { type: 'none' },
  });
  slide.addShape('triangle', { // hill
    x: cx - d * 0.16, y: cy - d * 0.19, w: d * 0.56, h: d * 0.35,
    fill: { color: PH_ICON }, line: { type: 'none' },
  });
}

/**
 * Vertical scrim fading from transparent to `color`. pptxgenjs cannot emit gradient
 * fills, so the fade is built from abutting translucent bands.
 */
function fadeDown(slide, x, y, w, h, color) {
  const n = 40;
  for (let i = 0; i < n; i++) {
    const top = y + (h * i) / n;
    slide.addShape('rect', {
      x,
      y: top,
      w,
      h: y + (h * (i + 1)) / n - top,
      fill: { color, transparency: Math.round(100 - (100 * (i + 0.5)) / n) },
      line: { type: 'none' },
    });
  }
}

/** Solid device body; the screen placeholder is drawn on top of it afterwards. */
function deviceBody(slide, x, y, w, h, radius) {
  slide.addShape('roundRect', {
    x, y, w, h, rectRadius: radius, fill: { color: DEVICE }, line: { type: 'none' },
  });
}

/* ------------------------------------------------------------------ *
 * Slide builders
 * ------------------------------------------------------------------ */

/* 1 — Cover */
function slide01(pptx) {
  const s = pptx.addSlide();
  s.background = { color: TEAL };
  head(s, 'MEDICAL', { x: -0.202, y: 0.585, w: 8.081, h: 3.698, size: 215, color: WHITE });
  head(s, 'PITCH DECK', { x: 3.948, y: 3.33, w: 6.313, h: 2.171, size: 125, color: WHITE });
  text(s, 'Presentation Template', { x: 1.905, y: 4.008, w: 2.313, h: 0.884, size: 24, color: WHITE });
  ribbon(s, {});
}

/* 2 — Problem Statement */
function slide02(pptx) {
  const s = pptx.addSlide();
  photo(s, 0, 0, 10, 2.812);
  ribbon(s, { page: 2 });
  head(s, 'Problem Statement', { x: 0.275, y: 2.949, w: 6.221, h: 1.186, size: 66 });
  body(s, LOREM, { x: 0.271, y: 3.945, w: 4.018, h: 0.984 });
  body(s, LOREM, { x: 5.0, y: 3.945, w: 4.018, h: 0.984 });
  learnMore(s, 8.352, 5.023, TEAL);
}

/* 3 — Market Opportunity */
function slide03(pptx) {
  const s = pptx.addSlide();
  ribbon(s, { dark: true, page: 3 });
  head(s, 'Market Opportunity', { x: 0.275, y: 0.657, w: 4.225, h: 2.297, size: 66 });
  body(s, paras([LOREM, '', LOREM_SHORT]), { x: 0.271, y: 2.954, w: 4.018, h: 1.666 });
  learnMore(s, 0.271, 5.023, TEAL);

  const stats = [
    { label: 'TAM', value: '$659,97', y: 0.657 },
    { label: 'TRENDS', value: '78,98%', y: 2.276 },
    { label: 'GROWTH ', value: '345,672++', y: 3.895 },
  ];
  stats.forEach(function (st) {
    text(s, st.label, { x: 7.319, y: st.y, w: 2.406, h: 0.328, size: 15, color: TEAL });
    head(s, st.value, { x: 7.319, y: st.y + 0.263, w: 2.406, h: 0.757, size: 41 });
    body(s, LOREM_TINY, { x: 7.319, y: st.y + 0.924, w: 2.406, h: 0.53 });
  });
  photo(s, 4.5, 0.657, 2.406, 4.694);
}

/* 4 — Your Solution */
function slide04(pptx) {
  const s = pptx.addSlide();
  s.background = { color: TEAL };
  ribbon(s, { page: 4 });
  head(s, 'Your Solution', { x: 0.275, y: 0.657, w: 4.725, h: 1.186, size: 66, color: WHITE });
  body(s, LOREM, { x: 5.0, y: 0.784, w: 4.725, h: 0.757, color: WHITE });

  // three recessed cards plus one raised white card
  const cards = [
    { title: 'device', x: 0, y: 2.326, tx: 0.275, ax: 0.271, bg: TEAL_D, fg: WHITE },
    { title: 'platform', x: 5.107, y: 2.326, tx: 5.382, ax: 5.378, bg: TEAL_D, fg: WHITE },
    { title: 'drug', x: 7.661, y: 2.326, tx: 7.935, ax: 7.932, bg: TEAL_D, fg: WHITE },
    { title: 'service', x: 2.554, y: 2.052, tx: 2.828, ax: 2.824, bg: WHITE, fg: TEAL },
  ];
  cards.forEach(function (c) {
    rect(s, c.x, c.y, 2.339, 3.299, c.bg);
    head(s, c.title, { x: c.tx, y: c.y + 0.295, w: 2.065, h: 0.757, size: 41, color: c.fg });
    body(s, LOREM_CARD, { x: c.tx, y: c.y + 1.107, w: 1.868, h: 1.212, color: c.fg });
    learnMore(s, c.ax, c.y + 2.697, c.fg);
  });
}

/* 5 — Product Overview */
function slide05(pptx) {
  const s = pptx.addSlide();
  ribbon(s, { dark: true, page: 5 });
  rect(s, 3.254, 0.938, 6.746, 2.062, CYAN);
  deviceBody(s, 0.347, 0.694, 2.986, 6.0, 0.42); // phone chassis (runs off the bottom)
  photo(s, 0.454, 0.799, 2.8, 5.845); // phone screen, runs off the bottom edge
  s.addShape('roundRect', { // camera notch hanging from the top bezel
    x: 1.271, y: 0.62, w: 1.139, h: 0.401, rectRadius: 0.11,
    fill: { color: DEVICE }, line: { type: 'none' },
  });
  head(s, 'Product Overview', { x: 3.604, y: 0.95, w: 6.121, h: 1.287, size: 72 });
  body(s, LOREM, { x: 3.604, y: 2.065, w: 6.121, h: 0.757 });

  [
    { title: 'Key features 01', x: 2.812, tx: 3.075, ax: 3.157 },
    { title: 'Key features 02', x: 6.16, tx: 6.422, ax: 6.505 },
  ].forEach(function (c) {
    rect(s, c.x, 3.27, 3.062, 2.081, TEAL);
    head(s, c.title, { x: c.tx, y: 3.43, w: 2.8, h: 0.631, size: 33, color: WHITE });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore.',
      { x: c.tx, y: 3.989, w: 2.66, h: 0.757, color: WHITE });
    learnMore(s, c.ax, 4.863, WHITE);
  });
}

/* 6 — Clinical Impact */
function slide06(pptx) {
  const s = pptx.addSlide();
  ribbon(s, { dark: true, page: 6 });
  head(s, 'Clinical Impact', { x: 0.275, y: 0.515, w: 9.451, h: 1.527, size: 86 });

  [
    { title: 'Clinical results', y: 1.938 },
    { title: 'patient outcomes', y: 3.135 },
    { title: 'accuracy', y: 4.333 },
  ].forEach(function (it) {
    checkBadge(s, 6.118, it.y + 0.104, 0.319, TEAL, WHITE);
    head(s, it.title, { x: 6.604, y: it.y, w: 3.121, h: 0.631, size: 33 });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore.',
      { x: 6.604, y: it.y + 0.49, w: 3.121, h: 0.53 });
  });
  photo(s, 0, 1.938, 5.583, 3.688);
}

/* 7 — Competitive Advantage */
function slide07(pptx) {
  const s = pptx.addSlide();
  s.background = { color: TEAL };
  ribbon(s, { page: 7 });
  head(s, 'Competitive Advantage',
    { x: 0.275, y: 0.515, w: 9.451, h: 1.527, size: 86, color: WHITE, align: 'center' });

  [
    { pct: '47', cx: 0.693, px: 0.94, tx: 2.685, label: 'Income Company A', value: '+24.018' },
    { pct: '58', cx: 5.33, px: 5.577, tx: 7.321, label: 'Income Company B', value: '+37.678' },
  ].forEach(function (d) {
    s.addShape('ellipse', {
      x: d.cx, y: 2.232, w: 1.641, h: 1.641,
      fill: { type: 'none' }, line: { color: TEAL_D, width: 20 },
    });
    s.addShape('arc', { // progress sweep, drawn clockwise from 12 o'clock
      x: d.cx, y: 2.232, w: 1.641, h: 1.641, angleRange: [270, 63.9],
      fill: { type: 'none' }, line: { color: CYAN, width: 20 },
    });
    s.addText(
      [
        { text: d.pct, options: { fontSize: 50 } },
        { text: '%', options: { fontSize: 50, superscript: true } },
      ],
      {
        x: d.px, y: 2.915, w: 1.193, h: 0.478, fontFace: BEBAS, color: WHITE,
        align: 'center', valign: 'middle', margin: NOINSET, lineSpacingMultiple: 1,
        isTextBox: true,
      }
    );
    text(s, d.label, { x: d.tx, y: 2.219, w: 1.985, h: 0.328, size: 15, color: WHITE, line: 1 });
    head(s, d.value, { x: d.tx, y: 2.587, w: 1.516, h: 0.682, size: 36, color: WHITE, line: 1 });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt.',
      { x: d.tx, y: 3.154, w: 1.985, h: 0.732, color: WHITE });
  });

  rect(s, 0, 4.381, 10, 1.244, WHITE);
  head(s, '68,7K', { x: 1.905, y: 4.439, w: 2.043, h: 1.186, size: 66 });
  text(s, 'Income Our Company', { x: 0.51, y: 4.771, w: 1.395, h: 0.581, size: 15, color: TEAL, line: 1 });
  body(s, LOREM, { x: 4.052, y: 4.603, w: 5.673, h: 0.757 });
}

/* 8 — Business Model */
function slide08(pptx) {
  const s = pptx.addSlide();
  photo(s, 0, 0, 3.688, 5.625);
  head(s, 'Business Model',
    { x: 1.979, y: 2.321, w: 4.875, h: 1.186, size: 66, align: 'center', rotate: 270 });
  ribbon(s, { dark: true, page: 8 });
  body(s, LOREM, { x: 4.874, y: 0.567, w: 3.156, h: 1.212 });

  [
    { title: 'B2B', y: 1.883 },
    { title: 'subscriptions', y: 3.117 },
    { title: 'licensing', y: 4.351 },
  ].forEach(function (it) {
    checkBadge(s, 6.154, it.y + 0.16, 0.26, TEAL, WHITE);
    head(s, it.title, { x: 6.49, y: it.y, w: 3.081, h: 0.581, size: 30 });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore.',
      { x: 6.49, y: it.y + 0.475, w: 3.236, h: 0.53 });
  });
}

/* 9 — Go-to-Market Strategy */
function slide09(pptx) {
  const s = pptx.addSlide();
  ribbon(s, { dark: true, page: 9 });
  head(s, 'Go-to-Market Strategy', { x: 0.275, y: 0.66, w: 4.875, h: 2.297, size: 66 });
  body(s, paras([LOREM, '', LOREM_SHORT]), { x: 1.082, y: 3.038, w: 3.918, h: 1.666 });
  learnMore(s, 1.082, 5.023, TEAL);

  // funnel bands, drawn bottom-up so each one overlaps the band below
  [
    { label: 'insurance', x: 5.928, y: 4.074, w: 2.868, ty: 4.372, th: 0.581, size: 30, fill: TEAL, alpha: 32 },
    { label: 'patients', x: 5.678, y: 3.023, w: 3.37, ty: 3.29, th: 0.682, size: 36, fill: TEAL },
    { label: 'doctors', x: 5.387, y: 1.945, w: 3.952, ty: 2.207, th: 0.757, size: 41, fill: TEAL_D },
    { label: 'HOSpitals', x: 5.0, y: 0.715, w: 4.725, ty: 0.946, th: 0.833, size: 45, fill: TEAL_DD, h: 1.448 },
  ].forEach(function (b) {
    const h = b.h || 1.261;
    const pts = GEOM.funnel.pts.map(function (p) {
      return { x: (p[0] / GEOM.funnel.w) * b.w, y: (p[1] / GEOM.funnel.h) * h };
    });
    pts.push({ close: true });
    s.addShape('custGeom', {
      x: b.x, y: b.y, w: b.w, h, points: pts,
      fill: b.alpha ? { color: b.fill, transparency: b.alpha } : { color: b.fill },
      line: { type: 'none' },
    });
    head(s, b.label, { x: b.x, y: b.ty, w: b.w, h: b.th, size: b.size, color: WHITE, align: 'center' });
  });
}

/* 10 — Traction & Milestones */
function slide10(pptx) {
  const s = pptx.addSlide();
  photo(s, 0, 0, 10, 5.625);
  fadeDown(s, 0, 2.721, 10, 2.904, TEAL);
  head(s, 'Traction & Milestones', { x: 0.275, y: 3.006, w: 7.329, h: 1.186, size: 66, color: WHITE });
  ribbon(s, { page: 10 });
  body(s, LOREM, { x: 0.275, y: 3.928, w: 7.829, h: 0.53, color: WHITE });

  [
    { title: 'trials', x: 1.027, tw: 1.405 },
    { title: 'pilot programs', x: 3.274, tw: 1.697 },
    { title: 'revenue', x: 5.521, tw: 1.405 },
    { title: 'patents', x: 7.768, tw: 1.405 },
  ].forEach(function (it) {
    checkBadge(s, it.x, 4.745, 0.26, WHITE, TEAL);
    head(s, it.title, { x: it.x + 0.292, y: 4.706, w: it.tw, h: 0.429, size: 21, color: WHITE });
    body(s, 'Lorem ipsum dolor sit amet', { x: it.x + 0.292, y: 5.043, w: 1.723, h: 0.303, color: WHITE });
  });
}

/* 11 — Technology or Science */
function slide11(pptx) {
  const s = pptx.addSlide();
  s.background = { color: TEAL };
  photo(s, 0, 0, 3.823, 5.625);
  ribbon(s, { page: 11 });
  head(s, 'Technology or Science', { x: 2.583, y: 0.591, w: 7.142, h: 1.186, size: 66, color: WHITE });
  body(s, LOREM, { x: 2.583, y: 1.666, w: 7.142, h: 0.53, color: WHITE });

  ['one', 'TWO', 'THREE'].forEach(function (word, i) {
    const x = 3.704 + i * 2.0875;
    head(s, paras(['Process', word]), { x, y: 2.827, w: 1.833, h: 1.085, size: 30, color: WHITE });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam.',
      { x, y: 3.913, w: 1.761, h: 1.439, color: WHITE });
    poly(s, GEOM.arrow, x + 0.917, 3.421, 0.717, 0.327, WHITE);
  });
}

/* 12 — Regulatory Pathway */
function slide12(pptx) {
  const s = pptx.addSlide();
  ribbon(s, { dark: true, page: 12 });
  head(s, 'Regulatory Pathway', { x: 0.275, y: 0.494, w: 9.451, h: 1.186, size: 66 });
  rect(s, 0, 2.883, 10, 1.062, TEAL);

  // steps alternate below / above the teal band
  [
    { num: '001', label: 'approvals', x: 0.275, ny: 3.231, ly: 4.082, dy: 4.589 },
    { num: '002', label: 'compliance', x: 2.551, ny: 2.723, ly: 2.292, dy: 1.524 },
    { num: '003', label: 'next steps', x: 4.827, ny: 3.231, ly: 4.082, dy: 4.589 },
    { num: '004', label: 'FINISH', x: 7.103, ny: 2.723, ly: 2.292, dy: 1.524 },
  ].forEach(function (st) {
    head(s, st.num, { x: st.x, y: st.ny, w: 1.147, h: 0.985, size: 54, color: WHITE });
    head(s, st.label, { x: st.x, y: st.ly, w: 2.276, h: 0.581, size: 30 });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore',
      { x: st.x, y: st.dy, w: 2.276, h: 0.757, color: GREY_TX2 });
  });
}

/* 13 — Market Adoption Strategy */
function slide13(pptx) {
  const s = pptx.addSlide();
  ribbon(s, { dark: true, page: 13 });
  head(s, 'Market Adoption Strategy',
    { x: 0.275, y: 0.654, w: 9.451, h: 1.287, size: 72, align: 'center' });

  ['Early adopters', 'pilot clients', 'channl partners'].forEach(function (title, i) {
    const x = 1.363 + i * 2.918;
    photo(s, x, 1.941, 1.866, 1.902);
    head(s, title, { x, y: 3.958, w: 2.524, h: 0.581, size: 30 });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ',
      { x, y: 4.589, w: 2.798, h: 0.682, color: GREY_TX2 });
  });
}

/* 14 — Team */
function slide14(pptx) {
  const s = pptx.addSlide();
  ribbon(s, { dark: true, page: 14 });
  head(s, 'Team', { x: -0.401, y: 0.983, w: 2.619, h: 1.818, size: 104, rotate: 270 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam.',
    { x: 0.294, y: 3.543, w: 1.992, h: 1.212, color: GREY_TX2 });
  learnMore(s, 0.294, 5.023, TEAL);

  [
    { name: 'Jack', role: 'Founders', px: 2.712, cx: 2.712, tx: 2.863 },
    { name: 'EMMa', role: 'Medical advisors', px: 5.145, cx: 5.147, tx: 5.3 },
    { name: 'James', role: 'Key hires', px: 7.585, cx: 7.582, tx: 7.74 },
  ].forEach(function (m) {
    photo(s, m.px, 0.769, 2.14, 3.566);
    rect(s, m.cx, 4.335, 2.143, 1.29, TEAL);
    head(s, m.name, { x: m.tx, y: 4.335, w: 1.99, h: 0.985, size: 54, color: WHITE });
    body(s, m.role, { x: m.tx - 0.002, y: 5.049, w: 1.99, h: 0.303, color: WHITE });
  });
}

/* 15 — Partnerships */
function slide15(pptx) {
  const s = pptx.addSlide();
  ribbon(s, { dark: true, page: 15 });

  // photo tile, gradient scrim and caption for each partner group
  [
    { label: 'research institutes', size: 41, px: 5.24, py: 0.719, pw: 4.486, ph: 2.948, gy: 2.021, gh: 1.646, tx: 5.4, ty: 2.909, tw: 4.326, th: 0.757 },
    { label: 'Hospitals', size: 33, px: 0.275, py: 2.193, pw: 4.59, ph: 1.844, gy: 2.721, gh: 1.315, tx: 0.525, ty: 3.405, tw: 4.326, th: 0.631 },
    { label: 'universities', size: 27, px: 5.24, py: 4.036, pw: 4.486, ph: 1.315, gy: 4.612, gh: 0.739, tx: 5.4, ty: 4.821, tw: 4.326, th: 0.53 },
    { label: 'pharma', size: 21, px: 0.275, py: 4.365, pw: 1.913, ph: 0.987, gy: 4.565, gh: 0.786, gw: 1.92, tx: 0.403, ty: 4.922, tw: 1.785, th: 0.429 },
    { label: 'Etc', size: 21, px: 2.57, py: 4.365, pw: 2.295, ph: 0.987, gy: 4.565, gh: 0.786, tx: 2.698, ty: 4.922, tw: 2.166, th: 0.429 },
  ].forEach(function (t) {
    photo(s, t.px, t.py, t.pw, t.ph);
    fadeDown(s, t.px, t.gy, t.gw || t.pw, t.gh, TEAL);
    head(s, t.label, { x: t.tx, y: t.ty, w: t.tw, h: t.th, size: t.size, color: WHITE });
  });
  head(s, 'Partnerships', { x: 0.275, y: 0.905, w: 4.892, h: 1.287, size: 72 });
}

/* 16 — Financial Projections */
function slide16(pptx) {
  const s = pptx.addSlide();
  s.background = { color: TEAL };
  ribbon(s, { page: 16 });
  head(s, 'Financial Projections',
    { x: 0.275, y: 0.51, w: 9.451, h: 1.085, size: 60, color: WHITE, align: 'right' });
  body(s, paras([LOREM, '', LOREM_SHORT]), { x: 0.275, y: 1.382, w: 2.6, h: 2.348, color: WHITE });
  learnMore(s, 0.271, 5.023, WHITE);
  head(s, '+798K', { x: 0.275, y: 3.801, w: 2.6, h: 1.085, size: 60, color: WHITE });
  body(s, '/ month', { x: 2.168, y: 4.344, w: 0.707, h: 0.303, color: WHITE });

  // hand-drawn column chart: grid, axis labels, then the staircase of bars
  const AXIS_Y = [1.586, 1.922, 2.257, 2.592, 2.928, 3.263, 3.598, 3.933, 4.269, 4.604, 4.939];
  const AXIS_V = ['50', '45', '40', '35', '30', '25', '20', '15', '10', '5', '0'];
  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Des'];
  const BAR_H = [0.012, 0.051, 0.126, 0.163, 0.203, 0.22, 0.239, 0.316, 0.393, 0.383, 0.402, 0.534];
  const BAR_Y = [5.062, 5.003, 4.877, 4.714, 4.511, 4.292, 4.053, 3.737, 3.344, 2.96, 2.559, 2.023];

  for (let i = 0; i < 11; i++) {
    s.addShape('line', {
      x: 3.772, y: 1.728 + i * 0.3351, w: 5.852, h: 0,
      line: { color: TEAL_D, width: 0.75 },
    });
  }
  s.addShape('line', { x: 3.774, y: 1.724, w: 0, h: 3.357, line: { color: TEAL_D, width: 0.75 } });

  AXIS_V.forEach(function (v, i) {
    body(s, v, { x: 3.288, y: AXIS_Y[i], w: 0.493, h: 0.273, color: WHITE, align: 'right', line: 1.3 });
  });
  MONTHS.forEach(function (m, i) {
    const x = 3.772 + i * 0.4878;
    body(s, m, { x, y: 5.081, w: 0.493, h: 0.273, color: WHITE, align: 'center', line: 1.3 });
    rect(s, x + 0.005, BAR_Y[i], 0.487, BAR_H[i], WHITE);
  });
}

/* 17 — Funding Ask */
function slide17(pptx) {
  const s = pptx.addSlide();
  ribbon(s, { dark: true, page: 17 });
  head(s, 'Funding Ask', { x: 0.275, y: 0.51, w: 4.725, h: 1.186, size: 66 });
  body(s, LOREM, { x: 0.275, y: 1.583, w: 4.061, h: 0.984, color: GREY_TX2 });

  [
    { pct: '48%', color: TEAL, x: 0.814, y: 2.77 },
    { pct: '25%', color: GREY1, x: 3.019, y: 2.77 },
    { pct: '18%', color: GREY2, x: 0.814, y: 4.241 },
    { pct: '9%', color: GREY4, x: 3.019, y: 4.241 },
  ].forEach(function (d) {
    head(s, d.pct, { x: d.x, y: d.y, w: 1.765, h: 0.757, size: 41, color: d.color });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit,',
      { x: d.x, y: d.y + 0.58, w: 1.765, h: 0.53, color: GREY_TX2 });
  });

  // donut assembled from four pie wedges plus a white hub
  [
    { fill: TEAL, range: [98.7, 270] },
    { fill: GREY1, range: [10.0, 99.3] },
    { fill: GREY2, range: [303.2, 10.3] },
    { fill: GREY3, range: [268.9, 304.0] },
  ].forEach(function (w) {
    s.addShape('pie', {
      x: 5.11, y: 0.7, w: 4.454, h: 4.454, angleRange: w.range, flipH: true,
      fill: { color: w.fill }, line: { color: WHITE, width: 1.5 },
    });
  });
  s.addShape('ellipse', {
    x: 5.801, y: 1.391, w: 3.072, h: 3.072, fill: { color: WHITE }, line: { type: 'none' },
  });
  // banknote icon: note behind (L), front note frame, inner rule and the coin
  poly(s, { w: 100, h: 100, pts: [[0, 0], [12.5, 0], [12.5, 87.5], [100, 87.5], [100, 100], [0, 100]] },
    6.968, 2.34, 0.638, 0.436, TEAL);
  s.addShape('roundRect', {
    x: 7.135, y: 2.272, w: 0.537, h: 0.336, rectRadius: 0.067,
    fill: { type: 'none' }, line: { color: TEAL, width: 4.8 },
  });
  s.addShape('rect', {
    x: 7.169, y: 2.306, w: 0.47, h: 0.269,
    fill: { type: 'none' }, line: { color: TEAL, width: 0.75 },
  });
  s.addShape('ellipse', {
    x: 7.303, y: 2.34, w: 0.201, h: 0.201, fill: { color: TEAL }, line: { type: 'none' },
  });
  text(s, paras(['INCOME ', 'IN QUARTER 4']), {
    x: 5.876, y: 2.933, w: 2.921, h: 0.682, size: 18, bold: true, color: TEAL, align: 'center',
  });
}

/* 18 — Exit Strategy */
function slide18(pptx) {
  const s = pptx.addSlide();
  ribbon(s, { dark: true, page: 18 });
  head(s, 'Exit Strategy', { x: 5.792, y: 0.922, w: 3.904, h: 2.297, size: 66 });
  deviceBody(s, -0.553, 0.941, 5.764, 3.892, 0.07); // laptop lid
  photo(s, -0.433, 1.142, 5.524, 3.492); // laptop screen
  poly(s, { w: 100, h: 100, pts: [[0, 0], [100, 0], [96.6, 100], [3.4, 100]] },
    -0.578, 4.87, 5.814, 0.135, DEVICE_2); // tapered base
  body(s, LOREM, { x: 5.792, y: 3.152, w: 3.904, h: 0.959, color: GREY_TX2 });
  learnMore(s, 5.883, 4.375, TEAL);
}

/* 19 — Call to Action */
function slide19(pptx) {
  const s = pptx.addSlide();
  ribbon(s, { dark: true, page: 19 });
  photo(s, 0, 0.604, 10, 2.812);
  head(s, 'Call to Action', { x: 0.275, y: 2.338, w: 6.036, h: 1.527, size: 86, color: WHITE });
  head(s, 'LOcation', { x: 6.037, y: 4.265, w: 1.688, h: 0.581, size: 30 });
  text(s, 'Paul and Mary Moore, 1313 E Main St, Portage MI 49024-2001',
    { x: 6.037, y: 4.821, w: 3.659, h: 0.53, size: 14, color: GREY_TX2 });
}

/* 20 — Thank You */
function slide20(pptx) {
  const s = pptx.addSlide();
  s.background = { color: TEAL };
  head(s, 'THANK', { x: -0.202, y: 0.585, w: 6.268, h: 3.698, size: 215, color: WHITE });
  head(s, 'YOU', { x: 3.784, y: 3.33, w: 2.542, h: 2.171, size: 125, color: WHITE });
  text(s, paras(['For your ', 'Attention']),
    { x: 1.905, y: 4.008, w: 2.313, h: 0.884, size: 24, color: WHITE });
  ribbon(s, {});
  photo(s, 6.326, 0, 3.674, 5.625);
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK', width: 10, height: 5.625 });
  pptx.layout = 'DECK';
  pptx.title = 'Medical Pitch Deck';
  pptx.theme = { headFontFace: BEBAS, bodyFontFace: LIGHT };

  [
    slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  ].forEach(function (fn) {
    fn(pptx);
  });

  return pptx.writeFile({
    fileName: path.join(__dirname, '084de3d2-fb08-412f-87c1-ac19e0fc217d_grok_final.pptx'),
  });
}

build().then(function (f) {
  console.log('wrote ' + f);
}).catch(function (e) {
  console.error(e);
  process.exit(1);
});
