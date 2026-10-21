/**
 * "Electric Cars" — 30 slide deck rebuilt with pptxgenjs.
 * Slide size 13.333 x 7.5 in (16:9). Run: node <thisfile>.js
 *
 * The original deck uses gradients, soft edges and vector clip-art that
 * pptxgenjs cannot express directly; those are approximated with stacks of
 * translucent native shapes (see `swoosh`, `dome`, `gauge`, ...).
 * Picture placeholders of the original are drawn as flat [image] boxes.
 */
'use strict';

const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ theme */

const C = {
  bg: 'F6FCFA',      // page background
  ink: '262626',     // headline text (tx1 lumMod 85%)
  black: '0D0D0D',
  body: '595959',    // body text (tx1 lumMod 65% / lumOff 35%)
  a1: 'FF5400',      // accent1 orange
  a2: 'F97737',
  a3: 'FFBD00',      // accent3 yellow
  a4: 'FFCE42',
  a5: 'D00000',      // accent5 deep red
  a6: 'FF1B1B',
  red: 'E60D0D',     // accent5 -> accent6 gradient, flattened
  white: 'FFFFFF',
  grey: 'D9D9D9',
  photo: 'F1F5F3',   // empty picture placeholder
  photoInk: 'C4CCC9'
};

const HEAD = 'Work Sans';   // major latin font of the theme
const BODY = 'Inter';       // minor latin font of the theme

const NOLINE = { type: 'none' };
const CARD_LINE = { color: C.body, transparency: 80, width: 1 };
const TILE_SHADOW = { type: 'outer', color: C.a5, opacity: 0.45, blur: 14, offset: 3, angle: 90 };

/* -------------------------------------------------------------- utilities */

/** blend two hex colours, t = 0 -> a, t = 1 -> b */
function mix (a, b, t) {
  const ch = i => Math.round(parseInt(a.substr(i, 2), 16) * (1 - t) + parseInt(b.substr(i, 2), 16) * t);
  return [ch(0), ch(2), ch(4)].map(v => v.toString(16).padStart(2, '0').toUpperCase()).join('');
}

/** preset-geometry "adj" value expressed the way pptxgenjs wants it (inches) */
function adj (val, w, h) { return (val / 100000) * Math.min(w, h); }

function shape (slide, kind, o) { slide.addShape(kind, o); }

/** plain rectangle */
function rect (slide, x, y, w, h, fill, more) {
  shape(slide, 'rect', Object.assign({ x, y, w, h, fill, line: NOLINE }, more));
}

/** the light "card" panel used all over the deck */
function card (slide, x, y, w, h, radius) {
  shape(slide, radius ? 'roundRect' : 'rect', {
    x, y, w, h, fill: { color: C.bg }, line: CARD_LINE,
    rectRadius: radius || undefined
  });
}

/**
 * text with the deck defaults; run arrays are copied because pptxgenjs
 * writes the resolved options back into the objects it is handed.
 */
function tx (slide, text, o) {
  const body = Array.isArray(text)
    ? text.map(function (r) { return { text: r.text, options: Object.assign({}, r.options) }; })
    : text;
  slide.addText(body, Object.assign({
    fontFace: BODY, fontSize: 14, color: C.body, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6]
  }, o));
}

/** italic headline */
function heading (slide, text, o) {
  tx(slide, text, Object.assign({
    fontFace: HEAD, fontSize: 48, bold: true, italic: true, color: C.ink, lineSpacingMultiple: 0.8
  }, o));
}

/**
 * Orange parallelogram "swoosh": in the original this is one parallelogram
 * (adj 53170) filled with a left-to-right alpha ramp of accent1 —
 * transparent at 0%, 20% opaque at 55%, solid at 100%.
 * Rebuilt as vertical slices of the same parallelogram, one per ramp step.
 */
const SWOOSH_SLICES = 16;
function swooshAlpha (t) { return t < 0.55 ? t / 0.55 * 0.2 : 0.2 + (t - 0.55) / 0.45 * 0.8; }
function swoosh (slide, x, y, w, h) {
  const off = 0.5317 * Math.min(w, h);   // parallelogram skew, in inches
  const span = w - off;
  for (let i = 0; i < SWOOSH_SLICES; i++) {
    const t0 = i / SWOOSH_SLICES, t1 = (i + 1.02) / SWOOSH_SLICES;
    shape(slide, 'custGeom', {
      x, y, w, h,
      fill: { color: C.a1, transparency: 100 - 100 * swooshAlpha((t0 + t1) / 2) }, line: NOLINE,
      points: [
        { x: off + t0 * span, y: 0 }, { x: off + t1 * span, y: 0 },
        { x: t1 * span, y: h }, { x: t0 * span, y: h }, { close: true }
      ]
    });
  }
}

/** three stacked chevrons, the deck's section marker */
function chevrons (slide, x, y, size) {
  const s = size || 0.333;
  [40, 60, 0].forEach(function (transparency, i) {
    shape(slide, 'chevron', {
      x: x + i * s * 0.781, y, w: s, h: s,
      fill: { color: C.a1, transparency }, line: NOLINE
    });
  });
}

/** white car pictogram — stands in for the icon artwork of the original */
function carIcon (slide, x, y, s, color) {
  const col = color || C.white;
  shape(slide, 'roundRect', {
    x: x + 0.04 * s, y: y + 0.40 * s, w: 0.92 * s, h: 0.30 * s, rectRadius: 0.09 * s,
    fill: { color: col }, line: NOLINE
  });
  shape(slide, 'trapezoid', {
    x: x + 0.22 * s, y: y + 0.18 * s, w: 0.56 * s, h: 0.26 * s,
    fill: { color: col }, line: NOLINE
  });
  [0.22, 0.62].forEach(function (wx) {
    shape(slide, 'ellipse', {
      x: x + wx * s, y: y + 0.64 * s, w: 0.16 * s, h: 0.16 * s, fill: { color: col }, line: NOLINE
    });
  });
}

/** white double chevron, the glyph on the title slide badge */
function doubleChevron (slide, x, y, s) {
  [0, 0.42].forEach(function (dx) {
    shape(slide, 'chevron', {
      x: x + dx * s, y, w: s * 0.58, h: s, fill: { color: C.white }, line: NOLINE
    });
  });
}

/** small red icon tile (a rounded red square with a lighter top half) */
function tile (slide, x, y, size, icon) {
  shape(slide, 'roundRect', {
    x, y, w: size, h: size, rectRadius: size * 0.2,
    fill: { color: C.red }, line: NOLINE, shadow: TILE_SHADOW
  });
  shape(slide, 'roundRect', {
    x: x + size * 0.06, y: y - size * 0.03, w: size * 0.88, h: size * 0.52, rectRadius: size * 0.26,
    fill: { color: 'E36666', transparency: 45 }, line: NOLINE
  });
  if (icon === 'chevron') doubleChevron(slide, x + size * 0.30, y + size * 0.26, size * 0.46);
  else if (icon !== false) carIcon(slide, x + size * 0.18, y + size * 0.2, size * 0.62);
}

/** red label chip ("230 Units", "Planning", ...) */
function chip (slide, x, y, w, h, text, fontSize) {
  shape(slide, 'rect', { x, y, w, h, fill: { color: C.red }, line: NOLINE, shadow: TILE_SHADOW });
  tx(slide, text, {
    x, y, w, h, align: 'center', valign: 'middle', color: C.white,
    fontFace: HEAD, bold: true, fontSize: fontSize || 18
  });
}

/**
 * Picture region of the original template. All of them are empty in the
 * reference; only the two layouts that carry a prompt paint anything, so
 * unlabelled regions are left as bare page.
 */
function photo (slide, x, y, w, h, label) {
  if (!label) return;
  rect(slide, x, y, w, h, { color: C.photo });
  tx(slide, label, {
    x, y: y + 0.09, w, h: 0.3, align: label === 'Image placeholder' ? 'left' : 'center',
    color: C.photoInk, fontSize: 12
  });
  // the "mountain" glyph PowerPoint shows inside an empty picture placeholder
  const g = 0.62, gx = x + w / 2 - g / 2, gy = y + h / 2 - g * 0.36;
  shape(slide, 'rect', { x: gx, y: gy, w: g, h: g * 0.72, fill: { color: C.white }, line: { color: C.photoInk, width: 0.75 } });
  shape(slide, 'triangle', { x: gx + 0.1, y: gy + 0.2, w: g * 0.6, h: g * 0.34, fill: { color: '9FC3E8' }, line: NOLINE });
  shape(slide, 'ellipse', { x: gx + 0.08, y: gy + 0.07, w: 0.1, h: 0.1, fill: { color: 'F2C14E' }, line: NOLINE });
}

/** free-form polygon from absolute [x, y] pairs (inches) */
function polygon (slide, pts, fill, more) {
  const xs = pts.map(function (p) { return p[0]; });
  const ys = pts.map(function (p) { return p[1]; });
  const x = Math.min.apply(null, xs), y = Math.min.apply(null, ys);
  const w = Math.max.apply(null, xs) - x, h = Math.max.apply(null, ys) - y;
  shape(slide, 'custGeom', Object.assign({
    x, y, w, h, fill, line: NOLINE,
    points: pts.map(function (p) { return { x: p[0] - x, y: p[1] - y }; }).concat([{ close: true }])
  }, more));
}

/** white check mark drawn as a free form */
function check (slide, x, y, w, h, color) {
  shape(slide, 'custGeom', {
    x, y, w, h, fill: { color: color || C.white }, line: NOLINE,
    points: [
      { x: 0, y: h * 0.45 }, { x: w * 0.16, y: h * 0.28 }, { x: w * 0.38, y: h * 0.55 },
      { x: w * 0.85, y: 0 }, { x: w, y: h * 0.18 }, { x: w * 0.38, y: h }, { close: true }
    ]
  });
}

/**
 * Ring segment (a thick arc) as a free form. Angles in degrees, 0 = up,
 * growing clockwise — the same convention the source deck uses.
 */
function arcBand (slide, cx, cy, rOut, rIn, a0, a1, fill, more) {
  const steps = Math.max(6, Math.round(Math.abs(a1 - a0) / 6));
  const pts = [];
  for (let i = 0; i <= steps; i++) {
    const a = (a0 + (a1 - a0) * i / steps) * Math.PI / 180;
    pts.push([cx + Math.sin(a) * rOut, cy - Math.cos(a) * rOut]);
  }
  for (let i = steps; i >= 0; i--) {
    const a = (a0 + (a1 - a0) * i / steps) * Math.PI / 180;
    pts.push([cx + Math.sin(a) * rIn, cy - Math.cos(a) * rIn]);
  }
  polygon(slide, pts, fill, more);
}

/** downward triangle (arrow head) centred on cx */
function arrowHead (slide, cx, yTop, w, h, color) {
  polygon(slide, [[cx - w / 2, yTop], [cx + w / 2, yTop], [cx, yTop + h]], { color });
}

/** big decorative quote glyph */
function quoteGlyph (slide, x, y, size, transparency) {
  tx(slide, '\u201C', {
    x, y, w: size * 1.6, h: size * 1.4, fontFace: HEAD, bold: true,
    fontSize: size * 78, color: C.ink, transparency: transparency
  });
}

/* --------------------------------------------------------------- the deck */

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'DECK', width: 13.333, height: 7.5 });
pptx.layout = 'DECK';
pptx.title = 'Electric Cars';
pptx.author = 'Anthony Louis';
pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };

/**
 * Every page carries: the page colour, a very faint orange dome at the
 * bottom (a blurred free form in the original master), the running header
 * and the page number rule. Slide 12 uses the one layout of the template
 * that switches the master graphics off (`bare`).
 */
function newSlide (bare) {
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  if (!bare) {
    for (let i = 0; i < 8; i++) {           // nested ellipses fake the soft dome
      const k = 1 - i * 0.09;
      shape(slide, 'ellipse', {
        x: 6.667 - 6.19 * k, y: 5.799 - 2.96 * k, w: 12.377 * k, h: 5.919 * k,
        fill: { color: C.a1, transparency: 98.95 }, line: NOLINE
      });
    }
    tx(slide, [
      { text: 'Electric Cars //', options: { bold: true, color: C.a1 } },
      { text: ' Electric cars presentation', options: { color: C.body } }
    ], { x: 0.581, y: 0.267, w: 5.259, h: 0.278, fontSize: 10.5 });
    shape(slide, 'line', { x: 0, y: 0.41, w: 0.438, h: 0, line: { color: C.body, transparency: 60, width: 0.5 } });
  }
  shape(slide, 'line', { x: 12.896, y: 7.09, w: 0.438, h: 0, line: { color: C.body, transparency: 60, width: 0.5 } });
  tx(slide, String(pptx.slides.length), {
    x: 12.137, y: 6.951, w: 0.615, h: 0.278, align: 'right', fontSize: 10.5, color: 'A8B3BB'
  });
  return slide;
}

/* ---------------------------------------------------- recurring fragments */

const LOREM = 'PLACEHOLDER' +
  'doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis ' +
  'et quasi architecto beatae vitae dicta sunt explicabo.';
const FAR_SHORT = 'Far far away, behind the word mountains, far from the countries Vokalia ' +
  'and Consonantia, there live the blind texts.';
const FAR_LONG = FAR_SHORT + ' Separated they live in Bookmarksgrove right at the coast of ' +
  'the Semantics, a large language ocean.';
const LEVERAGE = 'Leverage agile frameworks to provide a robust synopsis for high level overviews. ';
const SERENITY = 'A wonderful serenity has taken possession of my entire soul, like these sweet.';
const GRAMMAR = 'Lorem Ipsum decided to leave for the far World of Grammar. ';
const QUOTE_RUNS = [
  { text: '\u201CBehind the word mountains, far from the countries\u201D ', options: { italic: true, color: C.a1 } },
  { text: 'Vocalic and Consonantal, there live the blind texts. Separated they live in Bookmarks.', options: { color: C.body } }
];
const EARTH_RUNS = [
  { text: 'Electric cars are the cars of ' },
  { text: 'the future', options: { bold: true, color: C.a1 } },
  { text: ', cars that are environmentally friendly and ' },
  { text: 'keep the earth safe.', options: { bold: true, color: C.a1 } }
];

/** the centred "Behind the word mountains" footnote */
function footQuote (slide, y) {
  tx(slide, QUOTE_RUNS, {
    x: 2.949, y, w: 7.436, h: 0.676, align: 'center', lineSpacingMultiple: 1.3
  });
}

/* --------------------------------------------------------------- slide 01 */

function slide01 () {
  const s = newSlide();
  swoosh(s, 7.796, -0.377, 5.704, 2.019);
  swoosh(s, 6.542, 0.639, 3.361, 0.433);
  swoosh(s, -1.163, 6.821, 7.816, 2.766);
  swoosh(s, 2.056, 6.574, 6.755, 0.494);
  tx(s, 'ELECTRIC CARS', {
    x: 1.76, y: 2.403, w: 6.755, h: 2.767, fontFace: HEAD, fontSize: 88,
    bold: true, italic: true, color: C.ink, lineSpacingMultiple: 0.9
  });
  card(s, 6.411, 3.945, 5.162, 0.867);
  tile(s, 6.634, 4.157, 0.443, 'chevron');
  tx(s, [
    { text: 'Presentation by ', options: { bold: true, color: C.a1 } },
    { text: 'Anthony Louis', options: { color: C.ink } }
  ], { x: 7.299, y: 4.181, w: 3.997, h: 0.404, fontFace: HEAD, fontSize: 18 });
}

/* --------------------------------------------------------------- slide 02 */

function slide02 () {
  const s = newSlide();
  heading(s, 'Pure Performance Goes Electric', {
    x: 2.957, y: 1.023, w: 7.42, h: 1.561, fontSize: 54, align: 'center'
  });
  card(s, 1.154, 3.024, 5.809, 2.564);
  chevrons(s, 1.316, 2.855);
  tx(s, LOREM, { x: 1.693, y: 3.56, w: 4.732, h: 1.491, lineSpacingMultiple: 1.2 });
  tx(s, GRAMMAR, {
    x: 7.741, y: 3.303, w: 4.438, h: 0.854, fontFace: HEAD, fontSize: 18, bold: true,
    color: C.a1, lineSpacingMultiple: 1.3
  });
  tx(s, FAR_SHORT, { x: 7.741, y: 4.318, w: 4.438, h: 0.99, lineSpacingMultiple: 1.3 });
  footQuote(s, 6.027);
}

/* --------------------------------------------------------------- slide 03 */

function slide03 () {
  const s = newSlide();
  swoosh(s, 7.5, 0, 5.833, 0.878);
  swoosh(s, 7.347, -0.004, 3.868, 1.514);
  swoosh(s, 7.241, 6.25, 4.051, 0.33);
  heading(s, 'Two Paragraph Text', { x: 1.154, y: 1.478, w: 9.165, h: 0.916, fontSize: 60 });
  chevrons(s, 1.316, 2.656);
  card(s, 0.858, 2.829, 11.618, 3.587);
  tx(s, FAR_LONG + ' A small river named Duden flows by their place and supplies it with the necessary regelialia.',
    { x: 1.837, y: 3.666, w: 5.065, h: 2.215, lineSpacingMultiple: 1.3 });
  tx(s, 'It is a paradisematic country, in which roasted parts of sentences fly into your mouth. ' +
    'Even the all-powerful Pointing has no control about the blind texts it is an almost unorthographic life.',
    { x: 7.347, y: 3.666, w: 4.056, h: 1.603, lineSpacingMultiple: 1.3 });
  tile(s, 10.622, 2.293, 0.894);
}

/* --------------------------------------------------------------- slide 04 */

function slide04 () {
  const s = newSlide();
  swoosh(s, 0.441, 3.428, 3.868, 0.763);
  swoosh(s, 8.604, 3.772, 4.062, 0.27);
  tx(s, [{ text: '\u201C', options: { bold: true } }].concat(EARTH_RUNS, [{ text: '\u201D', options: { bold: true } }]), {
    x: 3.022, y: 1.134, w: 7.289, h: 1.373, align: 'center', fontFace: HEAD, fontSize: 28,
    italic: true, color: C.ink, lineSpacingMultiple: 0.9
  });
  tile(s, 5.951, 3.026, 1.431);
}

/* --------------------------------------------------------------- slide 05 */

function slide05 () {
  const s = newSlide();
  swoosh(s, 7.014, 0.224, 5.685, 0.855);
  swoosh(s, -1.637, 6.559, 8.986, 1.352);
  quoteGlyph(s, 1.35, 0.55, 3.1, 96);
  tx(s, EARTH_RUNS, {
    x: 2.462, y: 1.882, w: 8.408, h: 3.736, fontFace: HEAD, fontSize: 48,
    color: C.ink, lineSpacingMultiple: 0.9
  });
  shape(s, 'line', { x: 6.667, y: 5.858, w: 1.306, h: 0, line: { color: C.body, transparency: 50 } });
  tx(s, 'Gareth Peterson', { x: 8.132, y: 5.671, w: 2.299, h: 0.374, fontFace: HEAD, fontSize: 18 });
}

/* --------------------------------------------------------------- slide 06 */

function slide06 () {
  const s = newSlide();
  photo(s, 0.378, 0, 3.75, 6.333, false);
  photo(s, 4.33, 1.167, 3.469, 6.333, false);
  swoosh(s, -4.058, 5.714, 7.819, 0.806);
  swoosh(s, 1.515, 6.117, 3.222, 0.433);
  card(s, 6.667, 2.552, 5.809, 2.564);
  chevrons(s, 6.828, 2.384);
  tx(s, LOREM, { x: 7.205, y: 3.089, w: 4.732, h: 1.491, lineSpacingMultiple: 1.2 });
}

/* --------------------------------------------------------------- slide 07 */

function slide07 () {
  const s = newSlide();
  photo(s, 0, 2.263, 6.281, 5.237, 'Image placeholder');
  rect(s, 6.577, 1.356, 5.756, 6.144, { color: C.white });
  swoosh(s, 7.087, 2.083, 4.896, 0.875);
  swoosh(s, 2.458, -0.117, 4.198, 0.433);
  swoosh(s, 4.792, 2.668, 4.198, 0.433);
  heading(s, 'Creative, Smart and Modern', {
    x: 0.843, y: 0.723, w: 11.648, h: 1.01, fontSize: 54, align: 'center', lineSpacingMultiple: 1
  });
  tx(s, FAR_LONG, { x: 7.086, y: 3.453, w: 5.318, h: 1.297, lineSpacingMultiple: 1.3 });
  [['Feature One', 7.086], ['Feature Two', 9.961]].forEach(function (f) {
    tile(s, f[1] + 0.079, 5.288, 0.447);
    tx(s, f[0], { x: f[1] + 0.583, y: 5.341, w: 1.788, h: 0.404, fontFace: HEAD, fontSize: 18, bold: true, color: C.ink });
    tx(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing.', {
      x: f[1], y: 5.824, w: 2.567, h: 0.988, lineSpacingMultiple: 1.3
    });
  });
}

/* --------------------------------------------------------------- slide 08 */

function slide08 () {
  const s = newSlide();
  photo(s, 0.923, 1.103, 5.654, 5.295, false);
  card(s, 6.577, 1.09, 5.834, 5.319);
  tile(s, 6.144, 3.284, 0.894);
  [
    ['Point One', 1.496, C.a1, 0.572],
    ['Point Two', 3.111, C.a2, 0.808],
    ['Point Three', 4.726, C.a3, 0.808]
  ].forEach(function (p, i) {
    const y = p[1];
    shape(s, 'roundRect', {
      x: 7.295, y: y + 0.172, w: 0.169, h: 0.169, rectRadius: 0.04,
      fill: { color: p[2] }, line: NOLINE
    });
    tx(s, p[0], { x: 7.691, y, w: 2.509, h: 0.404, fontFace: HEAD, fontSize: 18, bold: true, color: C.ink });
    tx(s, LEVERAGE, { x: 7.691, y: y + 0.47, w: 3.95, h: p[3] });
    if (i < 2) {
      shape(s, 'line', { x: 7.766, y: 2.9 + i * 1.615, w: 3.09, h: 0, line: { color: C.grey } });
    }
  });
}

/* --------------------------------------------------------------- slide 09 */

function slide09 () {
  const s = newSlide();
  photo(s, 6.935, 0, 6.399, 7.5, false);
  swoosh(s, 7.26, 0.956, 4.229, 1.311);
  swoosh(s, 5.292, 0.737, 4.198, 0.433);
  heading(s, 'Battery Level on dashboard', { x: 1.413, y: 1.468, w: 7.179, h: 1.717, fontSize: 60 });
  chevrons(s, 1.838, 3.605);
  card(s, 1.413, 3.771, 6.728, 2.261);
  tx(s, 'Automotive Project', {
    x: 1.783, y: 4.117, w: 2.926, h: 0.404, fontFace: HEAD, fontSize: 18, bold: true, color: C.ink
  });
  tx(s, 'Sit amet consectetur elit sed do eiusmod tempor consectetur.', {
    x: 1.783, y: 4.521, w: 5.989, h: 0.381, lineSpacingMultiple: 1.3
  });
  [['340', C.a1, 4.808], ['876', C.a3, 6.348]].forEach(function (v) {
    tx(s, v[0], { x: v[2], y: 4.976, w: 1.297, h: 0.572, align: 'right', fontSize: 28, color: v[1] });
    tx(s, 'km/h', { x: v[2] + 0.384, y: 5.345, w: 0.913, h: 0.341, align: 'right', fontSize: 12 });
  });
  tx(s, 'Hype Automotive', { x: 8.904, y: 3.837, w: 2.986, h: 0.438, fontFace: HEAD, fontSize: 20, color: C.ink });
  tx(s, 'December. 07. 2022', { x: 8.904, y: 4.193, w: 2.986, h: 0.321, fontSize: 11 });
  tx(s, LEVERAGE, { x: 8.904, y: 4.772, w: 3.016, h: 0.983, lineSpacingMultiple: 1.3 });
}

/* --------------------------------------------------------------- slide 10 */

function slide10 () {
  const s = newSlide();
  photo(s, 0, 0, 3.128, 7.5, false);
  photo(s, 3.271, 0, 4.983, 7.5, false);
  swoosh(s, 3.694, 6.965, 7.819, 0.806);
  swoosh(s, -0.583, -0.117, 4.198, 0.433);
  heading(s, 'The Future has been our destination', {
    x: 1.154, y: 2.811, w: 4.442, h: 1.878, fontSize: 44
  });
  chevrons(s, 1.27, 4.79);
  [['All Features', '+239 Users', 1.459], ['New Models', '+159 Buyers', 4.12]].forEach(function (b) {
    card(s, 6.174, b[2], 1.945, 1.921, 0.12);
    tx(s, b[0], { x: 6.334, y: b[2] + 0.242, w: 1.641, h: 0.37, fontFace: HEAD, fontSize: 16, bold: true, color: C.ink });
    tx(s, b[1], { x: 6.333, y: b[2] + 0.585, w: 1.482, h: 0.303, fontSize: 12 });
    tile(s, 6.395, b[2] + 1.084, 0.652);
  });
}

/* --------------------------------------------------------------- slide 11 */

function slide11 () {
  const s = newSlide();
  swoosh(s, 7.26, 0.956, 4.229, 1.311);
  swoosh(s, 5.292, 0.737, 4.198, 0.433);
  heading(s, 'The Design Interior', {
    x: 3.222, y: 1.304, w: 6.888, h: 0.753, align: 'center'
  });
  chevrons(s, 6.245, 2.336);
  [['Modern', 1.179, 1.389, 1.371], ['Complete', 4.94, 5.168, 5.108], ['Comfort', 8.702, 8.979, 8.898]]
    .forEach(function (col) {
      photo(s, col[1], 3.255, 3.453, 2.056, 'Click icon to add picture');
      chip(s, col[2], 4.669, 1.473, 0.429, col[0], 16);
      tx(s, 'Leverage agile frameworks a robust. ', {
        x: col[3], y: 5.519, w: 2.576, h: 0.676, lineSpacingMultiple: 1.3
      });
    });
}

/* --------------------------------------------------------------- slide 12 */

function slide12 () {
  const s = newSlide(true);
  photo(s, 0, 0, 3.128, 7.5, false);
  photo(s, 3.271, 0, 4.983, 7.5, false);
  card(s, 5.99, 0.986, 6.213, 5.319);
  chevrons(s, 6.919, 0.854);
  heading(s, 'This Amazing Cars Future ', { x: 6.876, y: 1.727, w: 4.836, h: 1.393 });
  tile(s, 5.613, 3.18, 0.894);
  tx(s, GRAMMAR, {
    x: 6.919, y: 3.56, w: 4.438, h: 0.854, fontFace: HEAD, fontSize: 18, bold: true,
    color: C.a1, lineSpacingMultiple: 1.3
  });
  tx(s, FAR_SHORT, { x: 6.919, y: 4.574, w: 4.438, h: 0.99, lineSpacingMultiple: 1.3 });
}

/* --------------------------------------------------------------- slide 13 */

function slide13 () {
  const s = newSlide();
  heading(s, 'Smile From Costumers', {
    x: 1.776, y: 0.842, w: 9.782, h: 1.111, fontSize: 60, align: 'center', lineSpacingMultiple: 1
  });
  [1.314, 5.121, 8.939].forEach(function (x) {
    rect(s, x, 2.393, 3.116, 3.051, { color: '0C1212' });
    tx(s, 'Leverage agile frameworks a robust. ', {
      x: x + 0.262, y: 3.581, w: 2.576, h: 0.676, align: 'center', color: C.white, lineSpacingMultiple: 1.3
    });
  });
  footQuote(s, 5.981);
  tx(s, [
    { text: 'Poweride //', options: { bold: true, color: C.a1, charSpacing: 3 } },
    { text: ' Electric cars presentation', options: { color: C.body, charSpacing: 3 } }
  ], { x: 0.581, y: 6.947, w: 5.259, h: 0.278, fontSize: 10.5 });
}

/* --------------------------------------------------------------- slide 14 */

function slide14 () {
  const s = newSlide();
  heading(s, 'Power Battery Electric Cars', {
    x: 1.817, y: 1.212, w: 9.7, h: 0.747, align: 'center'
  });
  [
    ['Costumer Oriented', 'Sold for 346 Units', 1.391, 2.475],
    ['List Them For Sale!', 'Sold for 2473 Units', 4.968, 2.482]
  ].forEach(function (col) {
    const x = col[2], y = col[3];
    card(s, x, y, 3.32, 3.806);
    tx(s, col[0], { x: x + 0.26, y: y + 0.595, w: 2.9, h: 0.909, fontFace: HEAD, fontSize: 24, bold: true, color: C.ink });
    tx(s, col[1], { x: x + 0.26, y: y + 1.7, w: 2.358, h: 0.37, fontFace: HEAD, fontSize: 16, bold: true, color: C.a1 });
    tx(s, SERENITY, { x: x + 0.26, y: y + 2.27, w: 2.893, h: 0.992, lineSpacingMultiple: 1.3 });
    shape(s, 'ellipse', { x: x + 2.755, y: y + 0.231, w: 0.308, h: 0.308, fill: { color: C.a1 }, line: NOLINE });
    check(s, x + 2.82, y + 0.315, 0.18, 0.14);
  });
  card(s, 8.607, 2.429, 3.948, 3.86);
  chevrons(s, 10.112, 2.265);
  tx(s, '2,334,786,98+', {
    x: 9.264, y: 3.01, w: 2.54, h: 0.505, align: 'center', fontFace: HEAD, fontSize: 24, bold: true, color: C.ink
  });
  tx(s, 'People in the world', { x: 9.264, y: 3.495, w: 2.54, h: 0.375, align: 'center', lineSpacingMultiple: 1.3 });
  tx(s, SERENITY, { x: 9.135, y: 4.013, w: 2.893, h: 0.992, align: 'center', lineSpacingMultiple: 1.3 });
  chip(s, 9.588, 5.292, 1.987, 0.625, 'Planning', 16);
}

/* --------------------------------------------------------------- slide 15 */

function slide15 () {
  const s = newSlide();
  swoosh(s, -1.604, 0, 6.761, 1.208);
  swoosh(s, 3.326, 0, 3.868, 0.691);
  swoosh(s, 8.026, 5.733, 4.896, 0.875);
  heading(s, 'The Future has been our destination', { x: 1.751, y: 2.73, w: 4.484, h: 2.04 });
  chevrons(s, 1.898, 4.979);
  [['Our Mission', 1.271], ['Our Vision', 4.007]].forEach(function (b) {
    card(s, 7.634, b[1], 3.948, 2.222, 0.29);
    tx(s, b[0], { x: 8.487, y: b[1] + 0.421, w: 2.54, h: 0.505, fontFace: HEAD, fontSize: 24, bold: true, color: C.ink });
    tx(s, 'A wonderful serenity has taken soul, like these sweet.', {
      x: 8.487, y: b[1] + 1.105, w: 2.893, h: 0.682, lineSpacingMultiple: 1.3
    });
    tile(s, 7.263, b[1] + 0.728, 0.767);
  });
}

/* --------------------------------------------------------------- slide 16 */

function slide16 () {
  const s = newSlide();
  tx(s, EARTH_RUNS, {
    x: 1.935, y: 1.842, w: 9.463, h: 3.009, align: 'center', fontFace: HEAD, fontSize: 48,
    color: C.ink, lineSpacingMultiple: 0.9
  });
  shape(s, 'line', { x: 4.785, y: 5.471, w: 1.306, h: 0, line: { color: C.body, transparency: 50 } });
  tx(s, 'Gareth Peterson', { x: 6.249, y: 5.285, w: 2.299, h: 0.374, fontFace: HEAD, fontSize: 18 });
}

/* --------------------------------------------------------------- slide 17 */

/** the accent ramp of the serpentine arrow, keyed on horizontal position */
const SNAKE_RAMP = [[3.16, C.a1], [4.90, C.a2], [6.60, C.a3], [8.40, C.a4], [10.14, C.a5]];
function snakeColor (x) {
  for (let i = 1; i < SNAKE_RAMP.length; i++) {
    if (x <= SNAKE_RAMP[i][0] || i === SNAKE_RAMP.length - 1) {
      const a = SNAKE_RAMP[i - 1], b = SNAKE_RAMP[i];
      return mix(a[1], b[1], Math.max(0, Math.min(1, (x - a[0]) / (b[0] - a[0]))));
    }
  }
  return C.a1;
}

function slide17 () {
  const s = newSlide();
  // seven vertical bars linked by alternating half rings
  const BAR = 0.63, R_OUT = 0.895, R_IN = R_OUT - BAR, TOP1 = 2.075, TOP = 1.69, BOT = 2.735;
  const X = [3.155, 4.32, 5.49, 6.65, 7.82, 8.98, 10.135];
  [[0, TOP1], [2, TOP], [4, TOP]].forEach(function (h) {          // upper humps
    const cx = (X[h[0]] + X[h[0] + 1]) / 2;
    arcBand(s, cx, h[1], R_OUT, R_IN, -90, 90, { color: snakeColor(cx) });
  });
  [1, 3, 5].forEach(function (i) {                                 // lower U turns
    const cx = (X[i] + X[i + 1]) / 2;
    arcBand(s, cx, BOT, R_OUT, R_IN, 90, 270, { color: snakeColor(cx) });
  });
  [[0, TOP1, 4.414], [1, TOP1, BOT], [2, TOP, BOT], [3, TOP, BOT], [4, TOP, BOT], [5, TOP, BOT]]
    .forEach(function (b) {
      rect(s, X[b[0]] - BAR / 2, b[1], BAR, b[2] - b[1], { color: snakeColor(X[b[0]]) });
    });
  arrowHead(s, 3.164, 4.414, 1.1, 0.55, C.a1);                     // tail of the run
  arrowHead(s, 5.496, 1.777, 0.94, 0.44, snakeColor(5.496));       // mid arrow heads
  arrowHead(s, 7.828, 1.777, 0.94, 0.44, snakeColor(7.828));
  shape(s, 'downArrow', {                                          // head of the run
    x: 9.619, y: 0, w: 1.1, h: 2.26, fill: { color: C.a5 }, line: NOLINE
  });
  [
    ['22', 'Process one', C.a1, 0.388, 3.076],
    ['30', 'Process Two', C.a2, 4.207, 4.173],
    ['35', 'Process Three', C.a4, 7.582, 3.840],
    ['48', 'Process Four', C.a5, 10.777, 2.615]
  ].forEach(function (p) {
    tx(s, [
      { text: p[0], options: { fontSize: 36 } },
      { text: '%', options: { fontSize: 28 } }
    ], { x: p[3] + 0.19, y: p[4], w: 1.614, h: 0.702, align: 'center', fontFace: HEAD, bold: true, color: p[2] });
    tx(s, p[1], { x: p[3] - 0.2, y: p[4] + 0.667, w: 2.391, h: 0.337, align: 'center', fontFace: HEAD, fontSize: 20, bold: true, color: C.ink });
    tx(s, 'Totally ', { x: p[3] + 0.304, y: p[4] + 1.061, w: 1.384, h: 0.236, align: 'center', fontSize: 14 });
  });
  footQuote(s, 6.099);
}

/* --------------------------------------------------------------- slide 18 */

function slide18 () {
  const s = newSlide();
  heading(s, 'Revenue Achievement', {
    x: 2.047, y: 0.782, w: 9.239, h: 0.909, align: 'center', lineSpacingMultiple: 1
  });
  /**
   * Four bent arrows rising out of the bottom edge: the stems carry a
   * top-to-bottom alpha ramp, so they are painted as fading slices.
   * [x, yTop, colour, headPointsLeft, stemX]
   */
  const BEND_H = 1.42, STEM_W = 0.45, ARROW_W = 1.916;
  [
    [4.707, 2.185, C.a3, true, 6.15],
    [6.735, 3.004, C.a4, false, 6.70],
    [4.100, 3.888, C.a2, true, 5.59],
    [7.351, 4.441, C.a5, false, 7.26]
  ].forEach(function (a) {
    const top = a[1] + BEND_H, slices = 22, seg = (7.5 - top) / slices;
    for (let i = 0; i < slices; i++) {
      rect(s, a[4], top + i * seg, STEM_W, seg + 0.01,
        { color: a[2], transparency: Math.round(100 * i / slices) });
    }
    shape(s, 'bentArrow', {
      x: a[0], y: a[1], w: ARROW_W, h: BEND_H, flipH: a[3],
      fill: { color: a[2] }, line: NOLINE
    });
  });
  [
    ['927,2M', 1.468, 2.294, 'right', 3.427],
    ['627,2M', 1.214, 4.421, 'right', 3.173],
    ['577,2M', 9.094, 2.865, 'left', 9.230],
    ['377,2M', 9.602, 4.929, 'left', 9.738]
  ].forEach(function (v) {
    const alignRight = v[3] === 'right';
    tx(s, 'Value Title', { x: v[1], y: v[2], w: 2.444, h: 0.381, align: v[3], color: C.ink });
    tx(s, v[0], {
      x: v[1] + (alignRight ? 0.039 : 0.747), y: v[2] + 0.444, w: 1.697, h: 0.572,
      align: v[3], fontFace: HEAD, fontSize: 28, bold: true, color: C.ink
    });
    tile(s, v[4], v[2] + 0.521, 0.447);
  });
}

/* --------------------------------------------------------------- slide 19 */

function slide19 () {
  const s = newSlide();
  swoosh(s, -0.462, -0.428, 5.685, 0.855);
  swoosh(s, 6.236, 6.818, 5.685, 0.855);
  chevrons(s, 6.245, 0.771);
  heading(s, 'Company Progressions', { x: 2.047, y: 1.419, w: 9.239, h: 0.747, align: 'center' });
  // four interlocking hexagon outlines in the accent ramp
  [[1.51, C.a1], [4.06, C.a2], [6.61, C.a3], [9.08, C.a5]].forEach(function (hx) {
    shape(s, 'hexagon', {
      x: hx[0], y: 2.671, w: 2.72, h: 2.416,
      fill: { type: 'none' }, line: { color: hx[1], width: 14 }
    });
  });
  [['Step One', 1.953, 2.428], ['Step Two', 4.591, 4.954], ['Step Three', 7.128, 7.501], ['Step Four', 9.704, 9.996]]
    .forEach(function (b) {
      tile(s, b[2], 3.357, 0.894);
      card(s, b[1], 5.587, 1.781, 0.494);
      tx(s, b[0], {
        x: b[1], y: 5.587, w: 1.781, h: 0.494, align: 'center', valign: 'middle',
        fontFace: HEAD, fontSize: 16, bold: true, color: C.ink
      });
    });
}

/* --------------------------------------------------------------- slide 20 */

function slide20 () {
  const s = newSlide();
  swoosh(s, -1.264, -0.367, 10.778, 0.867);
  swoosh(s, 6.417, 7.094, 4.875, 0.653);
  heading(s, 'Company Schedule', {
    x: 2.727, y: 0.968, w: 7.879, h: 0.909, align: 'center', lineSpacingMultiple: 1
  });
  shape(s, 'line', { x: 6.682, y: 2.738, w: 0, h: 3.581, line: { color: C.grey, width: 0.75 } });
  [
    ['10:00', 1.406, 2.525, false], ['11:00', 1.406, 3.925, true], ['12:00', 1.406, 5.326, false],
    ['13:00', 7.191, 2.525, false], ['14:00', 7.191, 3.925, false], ['15:00', 7.191, 5.326, false]
  ].forEach(function (r) {
    const x = r[1], y = r[2], on = r[3];
    if (on) {
      shape(s, 'rect', { x, y, w: 4.736, h: 1.206, fill: { color: C.red }, line: NOLINE, shadow: TILE_SHADOW });
      shape(s, 'roundRect', {
        x: x + 0.005, y: y + 0.025, w: 4.729, h: 0.771, rectRadius: 0.16,
        fill: { color: 'E36666', transparency: 55 }, line: NOLINE
      });
    } else {
      card(s, x, y, 4.736, 1.206);
    }
    tx(s, r[0], {
      x: x + 0.314, y: y + 0.283, w: 1.666, h: 0.64, fontFace: HEAD, fontSize: 32, bold: true,
      color: on ? C.white : C.ink
    });
    tx(s, 'Far far away, behind the word mountains.', {
      x: x + 1.864, y: y + 0.261, w: 2.657, h: 0.684, lineSpacingMultiple: 1.3,
      color: on ? C.white : C.body
    });
  });
}

/* --------------------------------------------------------------- slide 21 */

function slide21 () {
  const s = newSlide();
  swoosh(s, 2.032, -0.062, 9.417, 0.97);
  swoosh(s, 1.199, -0.117, 4.198, 0.433);
  chevrons(s, 1.634, 1.569);
  heading(s, 'Agenda Schedule', { x: 1.445, y: 2.174, w: 4.548, h: 1.724, fontSize: 60 });
  tx(s, 'Hype Automotive', { x: 1.559, y: 4.014, w: 2.986, h: 0.438, fontFace: HEAD, fontSize: 20, color: C.ink });
  tx(s, 'December. 21. 2021', { x: 1.559, y: 4.369, w: 2.986, h: 0.321, fontSize: 11 });
  tx(s, LEVERAGE, { x: 1.559, y: 4.948, w: 3.016, h: 0.983, lineSpacingMultiple: 1.3 });
  card(s, 6.401, 1.563, 5.487, 4.103);
  // calendar: header strip then a 6 x 7 grid of day numbers
  const CW = 0.746, CH = 0.526, CX = 6.55, CY = 1.699;
  'SMTWTFS'.split('').forEach(function (d, i) {
    rect(s, CX + i * CW, CY, CW, CH, { color: i === 0 ? 'BF3F00' : C.a1 });
    tx(s, d, { x: CX + i * CW, y: CY, w: CW, h: CH, align: 'center', valign: 'middle', fontFace: HEAD, color: C.white });
  });
  const FIRST = 3;   // the 1st falls on a Wednesday
  for (let day = 1; day <= 31; day++) {
    const idx = FIRST + day - 1;
    const cx = CX + (idx % 7) * 0.7525, cy = 2.328 + Math.floor(idx / 7) * 0.5985;
    if (day === 21) {
      rect(s, cx, cy, 0.706, 0.564, { color: C.a1 }, { shadow: TILE_SHADOW });
    }
    tx(s, String(day), {
      x: cx, y: cy, w: 0.706, h: 0.564, align: 'center', valign: 'middle', fontSize: 12,
      color: day === 21 ? C.white : (idx % 7 === 0 ? 'BFBFBF' : C.body)
    });
  }
  // call-out card hanging under the 21st
  rect(s, 5.519, 4.796, 3.554, 1.141, { color: C.white });
  shape(s, 'triangle', {
    x: 8.453, y: 4.5, w: 0.339, h: 0.388, rotate: 335, fill: { color: C.white }, line: NOLINE
  });
  tx(s, 'Dec, 21', { x: 5.879, y: 5.008, w: 1.309, h: 0.438, fontFace: HEAD, fontSize: 20, bold: true, color: C.a1 });
  tx(s, '4:00 PM', { x: 5.883, y: 5.384, w: 0.963, h: 0.341, fontSize: 12 });
  tx(s, 'Release New Product', {
    x: 7.234, y: 5.023, w: 1.478, h: 0.686, fontFace: HEAD, bold: true, color: C.ink, lineSpacingMultiple: 1.3
  });
}

/* --------------------------------------------------------------- slide 22 */

/** world map, drawn as the grid of dots used by the original artwork */
const WORLD_MAP = [
  '               ##### ####                                       ',
  '             ##############     ##                              ',
  '            # ##  #########                     ##              ',
  '        ### ####    ######              #    ######             ',
  '         ## # ###    #####         #      ############### #   # ',
  ' ##############  #   ###         ####  #########################',
  ' ##############  #   ##         ##############################  ',
  ' ############   #              ## ######################### #   ',
  '      ########   ##            #  #####################   #     ',
  '       ######## ###          # ########################   #     ',
  '        ############         ##########################         ',
  '        ####### ##            ## ### # ###############          ',
  '        ########            ##    ##### #############           ',
  '         #######             ###     ###############            ',
  '         #### #             ########################            ',
  '           #               ######### ###   #########            ',
  '           ##              #############   ### ###              ',
  '                           ###########      #   #               ',
  '                 ###        ###########                         ',
  '                #####           #######         # ##            ',
  '                ######          #####                 ##        ',
  '                ########        #####                           ',
  '                 ######         #####                           ',
  '                 ######         ##### #             #####       ',
  '                 #####           ###               #######      ',
  '                 ####            ###               #######      ',
  '                 ###                                  ###       ',
  '                 ##                                          #  ',
  '                 #                                          #   ',
  '                 #                                              ',
  '                 #                                              '
];

function dotMap (slide, grid, x, y, w, h, color, transparency) {
  const cw = w / grid[0].length, ch = h / grid.length, d = Math.min(cw, ch) * 0.55;
  grid.forEach(function (row, r) {
    for (let c = 0; c < row.length; c++) {
      if (row[c] !== '#') continue;
      shape(slide, 'ellipse', {
        x: x + c * cw + (cw - d) / 2, y: y + r * ch + (ch - d) / 2, w: d, h: d,
        fill: { color, transparency }, line: NOLINE
      });
    }
  });
}

function slide22 () {
  const s = newSlide();
  swoosh(s, -0.462, -0.428, 5.685, 0.855);
  swoosh(s, 7.177, -0.155, 5.685, 0.855);
  dotMap(s, WORLD_MAP, 2.484, 2.46, 8.365, 4.079, C.ink, 82);
  heading(s, 'Electric Car Dealers', {
    x: 2.047, y: 1.053, w: 9.239, h: 0.909, align: 'center', lineSpacingMultiple: 1
  });
  chevrons(s, 6.219, 0.671);
  [['230 Units', 1.739, 2.817], ['120 Units', 6.058, 3.261], ['22 Units', 9.002, 3.748],
    ['810 Units', 5.141, 4.275], ['70 Units', 9.836, 5.345]]
    .forEach(function (m) { chip(s, m[1], m[2], 1.748, 0.495, m[0]); });
  // "High Value" call-out
  rect(s, 2.398, 4.491, 2.274, 1.901, { color: C.white });
  shape(s, 'triangle', {
    x: 2.047, y: 4.438, w: 0.426, h: 0.535, rotate: 299, fill: { color: C.white }, line: NOLINE
  });
  tx(s, 'High Value', { x: 2.376, y: 4.692, w: 2.078, h: 0.447, fontFace: HEAD, fontSize: 18, bold: true, color: C.ink });
  shape(s, 'line', { x: 2.576, y: 5.219, w: 1.078, h: 0, line: { color: C.body, transparency: 68, width: 1 } });
  tx(s, '$452,378,000', { x: 2.376, y: 5.337, w: 1.612, h: 0.39, color: C.a1, lineSpacingMultiple: 1.3 });
  tx(s, '$322,378,000', { x: 2.376, y: 5.694, w: 1.612, h: 0.37, color: C.a4, lineSpacingMultiple: 1.3 });
}

/* --------------------------------------------------------------- slide 23 */

/** United States, blocked out row by row from the original outline */
const US_MAP = [
  '  #######                                              ## ',
  '  ##############################                      ####',
  '  ################################# #               ##### ',
  ' ##################################### ###          ###   ',
  ' ##################################### ####    ####  ###  ',
  ' ##################################### ####   ########    ',
  '#####################################################     ',
  '####################################################      ',
  ' ###################################################      ',
  ' ###################################################      ',
  '  #  ###############################################      ',
  '       ############################################       ',
  '     ############################################         ',
  '      ##########################################          ',
  '           ####################################           ',
  '                   #############################          ',
  '                    ##################      # ###         ',
  '                         ######               ###         ',
  '                          ###                  ###        ',
  '                           ##                             '
];

/** paint contiguous runs of a character grid as rectangles */
function blockMap (slide, grid, x, y, w, h, color) {
  const cw = w / grid[0].length, ch = h / grid.length;
  grid.forEach(function (row, r) {
    let c = 0;
    while (c < row.length) {
      if (row[c] !== '#') { c++; continue; }
      let e = c;
      while (e < row.length && row[e] === '#') e++;
      rect(slide, x + c * cw, y + r * ch, (e - c) * cw, ch * 1.02, { color });
      c = e;
    }
  });
}

/** concentric halo used for the map pins */
function pin (slide, cx, cy, r, color, label) {
  [[1, 75], [0.813, 55], [0.644, 0]].forEach(function (ring) {
    const rr = r * ring[0];
    shape(slide, 'ellipse', {
      x: cx - rr, y: cy - rr, w: 2 * rr, h: 2 * rr,
      fill: { color, transparency: ring[1] }, line: NOLINE
    });
  });
  tx(slide, label, {
    x: cx - r, y: cy - 0.185, w: 2 * r, h: 0.37, align: 'center', fontFace: HEAD,
    fontSize: 16, color: C.white
  });
}

function slide23 () {
  const s = newSlide();
  blockMap(s, US_MAP, 0.965, 2.101, 6.216, 3.622, 'DCDCDC');
  pin(s, 2.157, 2.351, 0.575, C.a1, '82%');
  pin(s, 2.616, 4.192, 0.575, C.a2, '55%');
  pin(s, 4.957, 3.424, 0.575, C.a3, '34%');
  pin(s, 5.855, 4.889, 0.575, C.a4, '20%');
  chevrons(s, 7.536, 1.569);
  heading(s, 'United States Map', { x: 7.347, y: 2.174, w: 5.021, h: 1.724, fontSize: 60 });
  tx(s, 'Hype Automotive', { x: 7.461, y: 4.014, w: 2.986, h: 0.438, fontFace: HEAD, fontSize: 20, bold: true, color: C.ink });
  tx(s, 'December. 21. 2021', { x: 7.461, y: 4.369, w: 2.986, h: 0.321, fontSize: 11 });
  tx(s, LEVERAGE, { x: 7.461, y: 4.948, w: 3.016, h: 0.983, lineSpacingMultiple: 1.3 });
}

/* --------------------------------------------------------------- slide 24 */

/** icon bubble on a stalk (roadmap markers) */
function bubble (slide, cx, cy, r, color, stalkTo, glyph) {
  shape(slide, 'line', { x: cx, y: cy, w: 0, h: stalkTo - cy, line: { color, width: 1.5 } });
  shape(slide, 'ellipse', { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, fill: { color }, line: NOLINE });
  shape(slide, 'ellipse', {
    x: cx - r * 0.85, y: cy - r * 0.85, w: 1.7 * r, h: 1.7 * r, fill: { color: C.white }, line: NOLINE
  });
  tx(slide, glyph, {
    x: cx - r, y: cy - r * 0.55, w: 2 * r, h: r * 1.1, align: 'center', valign: 'middle',
    fontSize: r * 34, color
  });
}

/**
 * The winding road of slide 24, sampled off the original artwork:
 * left kerb from the top of the slide down, right kerb bottom-up.
 */
const ROAD_LEFT = [
  [8.80, 1.74], [8.40, 2.35], [8.95, 2.80], [10.40, 3.25], [10.95, 3.60], [11.30, 4.00],
  [10.35, 4.40], [9.86, 4.80], [9.31, 5.20], [8.73, 5.60], [8.55, 6.05], [7.90, 6.55],
  [6.72, 7.05], [6.10, 7.50]
];
const ROAD_RIGHT = [
  [9.90, 7.50], [10.24, 7.00], [10.72, 6.50], [11.18, 6.00], [11.54, 5.60], [11.89, 5.20],
  [12.22, 4.80], [12.31, 4.40], [12.47, 4.00], [12.39, 3.60], [11.99, 3.25], [10.93, 2.80],
  [9.12, 2.40], [9.62, 1.74]
];

function slide24 () {
  const s = newSlide();
  polygon(s, ROAD_LEFT.concat(ROAD_RIGHT), { color: '444C55' });
  // white kerb stripes and the dashed centre line
  ROAD_LEFT.forEach(function (p, i) {
    const r = ROAD_RIGHT[ROAD_RIGHT.length - 1 - i];
    if (i % 2 === 0 && i > 1) {
      const mx = (p[0] + r[0]) / 2, my = (p[1] + r[1]) / 2, sc = 0.05 + i * 0.028;
      polygon(s, [[mx - sc, my], [mx + sc, my], [mx + sc * 1.4, my + 0.13], [mx - sc * 0.6, my + 0.13]],
        { color: C.white, transparency: 25 });
    }
  });
  carIcon(s, 8.30, 1.44, 0.55, C.a3);
  carIcon(s, 10.35, 3.45, 1.05, C.a1);
  carIcon(s, 7.20, 5.30, 1.60, C.a1);
  bubble(s, 8.163, 1.157, 0.32, C.a3, 2.36, '\u25D4');
  bubble(s, 11.767, 2.024, 0.47, C.a2, 4.04, '\u2692');
  bubble(s, 7.025, 3.617, 0.54, C.a1, 6.20, '\u26B7');
  chevrons(s, 1.316, 1.105);
  heading(s, 'Roadmap Infographic', { x: 1.154, y: 1.704, w: 5.024, h: 1.724, fontSize: 60 });
  [['Good Services', 3.721], ['Users Convenience', 5.241]].forEach(function (f) {
    shape(s, 'rect', {
      x: 1.367, y: f[1] + 0.135, w: 0.55, h: 0.55, fill: { color: C.red }, line: NOLINE, shadow: TILE_SHADOW
    });
    check(s, 1.482, f[1] + 0.298, 0.319, 0.224);
    tx(s, f[0], { x: 2.274, y: f[1], w: 3.111, h: 0.447, fontFace: HEAD, fontSize: 18, bold: true, color: C.ink });
    tx(s, 'Leverage agile frameworks to provide a robust.', {
      x: 2.274, y: f[1] + 0.47, w: 2.95, h: 0.684, lineSpacingMultiple: 1.3
    });
  });
}

/* --------------------------------------------------------------- slide 25 */

function slide25 () {
  const s = newSlide();
  swoosh(s, -0.462, -0.428, 5.685, 0.855);
  swoosh(s, 2.96, 5.976, 5.685, 0.855);
  heading(s, 'Testimo-nials', { x: 1.229, y: 2.888, w: 3.664, h: 1.724, fontSize: 60 });
  [
    ['Servina Rose', 1.296, 7.269, 2.797],
    ['George Ham', 3.984, 7.295, 5.482]
  ].forEach(function (t, i) {
    const y = t[1];
    card(s, 5.802, y, 6.297, 2.22);
    quoteGlyph(s, 6.10, y + 0.15, 0.72, 88);
    tx(s, 'Behind the word mountains there live the blind texts. Separated they live in Bookmarks.', {
      x: t[2], y: y + 0.297, w: 4.155, h: 0.983, italic: true, align: i ? 'right' : 'left', lineSpacingMultiple: 1.3
    });
    chip(s, i ? 9.361 : 7.269, t[3], 2.089, 0.422, t[0], 14);
  });
}

/* --------------------------------------------------------------- slide 26 */

function slide26 () {
  const s = newSlide();
  swoosh(s, 2.125, 0, 5.89, 0.607);
  swoosh(s, 2.208, 6.565, 9.069, 0.935);
  // donut built from three "pie" wedges of decreasing radius
  [
    [1.476, 1.484, 4.401, [270, 88], C.a1],
    [1.681, 1.689, 3.99, [88, 165], C.a5],
    [1.886, 1.894, 3.58, [165, 270], C.a3]
  ].forEach(function (w) {
    shape(s, 'pie', {
      x: w[0], y: w[1], w: w[2], h: w[2], angleRange: w[3],
      fill: { color: w[4] }, line: NOLINE
    });
  });
  shape(s, 'ellipse', {
    x: 2.397, y: 2.405, w: 2.559, h: 2.559, fill: { color: C.bg }, line: NOLINE,
    shadow: { type: 'outer', color: '000000', opacity: 0.1, blur: 12, offset: 2, angle: 90 }
  });
  tx(s, 'Circle', { x: 2.392, y: 3.186, w: 2.606, h: 0.707, align: 'center', fontFace: HEAD, fontSize: 36, bold: true, color: C.ink });
  tx(s, 'Infographic', { x: 2.392, y: 3.778, w: 2.606, h: 0.404, align: 'center', fontSize: 18 });
  [['35%', 1.273, 2.003], ['43%', 4.968, 1.74], ['21%', 2.829, 5.142]].forEach(function (b) {
    shape(s, 'ellipse', { x: b[1], y: b[2], w: 1.125, h: 1.125, fill: { color: C.bg }, line: CARD_LINE });
    tx(s, b[0], {
      x: b[1], y: b[2], w: 1.125, h: 1.125, align: 'center', valign: 'middle',
      fontFace: HEAD, fontSize: 18, bold: true, color: C.ink
    });
  });
  heading(s, 'Roadmap Infographic', { x: 7.036, y: 2.165, w: 5.024, h: 1.724, fontSize: 60 });
  shape(s, 'rect', {
    x: 7.248, y: 4.316, w: 0.55, h: 0.55, fill: { color: C.red }, line: NOLINE, shadow: TILE_SHADOW
  });
  check(s, 7.364, 4.48, 0.319, 0.224);
  tx(s, 'Good Services', { x: 8.156, y: 4.181, w: 3.111, h: 0.447, fontFace: HEAD, fontSize: 18, bold: true, color: C.ink });
  tx(s, 'Leverage agile frameworks to provide a robust.', {
    x: 8.156, y: 4.651, w: 2.883, h: 0.684, lineSpacingMultiple: 1.3
  });
}

/* --------------------------------------------------------------- slide 27 */

function slide27 () {
  const s = newSlide();
  heading(s, 'Custom Infographic', {
    x: 2.047, y: 0.942, w: 9.239, h: 0.909, align: 'center', lineSpacingMultiple: 1
  });
  chevrons(s, 6.219, 0.538);
  // three S-curves: an upper-left "C" arrow and a lower-right one per node
  [2.204, 5.622].forEach(function (x) {
    shape(s, 'curvedRightArrow', {
      x, y: 2.027, w: 2.692, h: 1.428, rotate: 180, flipV: true,
      fill: { color: C.a1 }, line: NOLINE
    });
    shape(s, 'curvedRightArrow', {
      x: x + 2.833, y: 3.619, w: 2.654, h: 1.342,
      fill: { color: C.a1 }, line: NOLINE
    });
  });
  shape(s, 'curvedRightArrow', {
    x: 8.983, y: 2.095, w: 2.126, h: 0.828, rotate: 180, flipV: true,
    fill: { color: C.a1 }, line: NOLINE
  });
  shape(s, 'curvedRightArrow', { x: 2.206, y: 4.251, w: 2.092, h: 0.723, fill: { color: C.a1 }, line: NOLINE });
  [[2.529, C.a1], [5.993, C.a2], [9.444, C.a3]].forEach(function (n) {
    shape(s, 'ellipse', { x: n[0], y: 2.87, w: 1.345, h: 1.345, fill: { color: n[1] }, line: NOLINE });
    tx(s, 'Your text goes here', {
      x: n[0] + 0.065, y: 3.29, w: 1.211, h: 0.505, align: 'center', fontFace: HEAD,
      fontSize: 12, color: C.white
    });
  });
  [['34%', 4.59], ['62%', 7.99]].forEach(function (b) {
    shape(s, 'ellipse', {
      x: b[1], y: 3.165, w: 0.741, h: 0.741, fill: { color: C.white }, line: NOLINE,
      shadow: { type: 'outer', color: '000000', opacity: 0.15, blur: 8, offset: 2, angle: 90 }
    });
    tx(s, b[0], { x: b[1] + 0.038, y: 3.369, w: 0.665, h: 0.37, align: 'center', fontFace: HEAD, fontSize: 16 });
  });
  [
    ['Automotive', 1.71, 2.197, '\u265B', C.a1],
    ['New design', 5.173, 5.661, '\u2600', C.a2],
    ['Car dealer', 8.623, 9.11, '\u25D4', '22232B']
  ].forEach(function (col) {
    tx(s, col[3], { x: col[1] + 1.22, y: 5.19, w: 0.55, h: 0.42, align: 'center', fontSize: 18, color: col[4] });
    tx(s, col[0], { x: col[2], y: 5.648, w: 2.011, h: 0.37, align: 'center', fontFace: HEAD, fontSize: 16, bold: true, color: C.ink });
    tx(s, 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem', {
      x: col[1], y: 5.994, w: 2.986, h: 0.563, align: 'center', fontSize: 12, lineSpacingMultiple: 1.2
    });
  });
}

/* --------------------------------------------------------------- slide 28 */

function slide28 () {
  const s = newSlide();
  chevrons(s, 6.219, 1.005);
  heading(s, 'Electric Cars Services', { x: 2.047, y: 1.612, w: 9.239, h: 0.747, align: 'center' });
  [
    ['8.009', C.a1, 2.26, 2.805, 1.564, 'right'],
    ['32.500', C.a3, 1.236, 4.526, 0.706, 'right'],
    ['30.222', C.a4, 9.801, 2.94, 9.801, 'left'],
    ['14.020', C.a5, 10.258, 4.747, 10.258, 'left']
  ].forEach(function (v) {
    tx(s, [
      { text: v[0], options: { fontSize: 24 } },
      { text: '/km', options: { fontSize: 18 } }
    ], { x: v[2], y: v[3], w: 1.834, h: 0.505, align: v[5], fontFace: HEAD, bold: true, color: v[1] });
    tx(s, 'A wonderful serenity has taken possession.', {
      x: v[4], y: v[3] + 0.55, w: 2.37, h: 0.563, align: v[5], fontSize: 12, lineSpacingMultiple: 1.2
    });
  });
}

/* --------------------------------------------------------------- slide 29 */

/** speedometer style ring with tick marks */
function gauge (slide, x, y, size, color) {
  shape(slide, 'ellipse', { x, y, w: size, h: size, fill: { color }, line: NOLINE });
  const g1 = size * 0.082;
  shape(slide, 'ellipse', {
    x: x + g1, y: y + g1, w: size - 2 * g1, h: size - 2 * g1, fill: { color: 'E9E9E9' }, line: NOLINE
  });
  const cx = x + size / 2, cy = y + size / 2, rOut = size * 0.40, rIn = size * 0.34;
  for (let i = 0; i < 40; i++) {
    const a = (i / 40) * 2 * Math.PI;
    const x1 = cx + Math.sin(a) * rIn, y1 = cy - Math.cos(a) * rIn;
    const x2 = cx + Math.sin(a) * rOut, y2 = cy - Math.cos(a) * rOut;
    shape(slide, 'line', {
      x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
      flipH: (x2 - x1) * (y2 - y1) < 0, line: { color, width: 1 }
    });
  }
  const g2 = size * 0.245;
  shape(slide, 'ellipse', {
    x: x + g2, y: y + g2, w: size - 2 * g2, h: size - 2 * g2, fill: { color: C.white }, line: NOLINE,
    shadow: { type: 'outer', color: '000000', opacity: 0.12, blur: 10, offset: 2, angle: 90 }
  });
}

function slide29 () {
  const s = newSlide();
  swoosh(s, -1.049, -0.122, 9.069, 0.935);
  swoosh(s, -0.324, 0.519, 5.89, 0.607);
  heading(s, 'Custom Infographic', {
    x: 2.52, y: 0.71, w: 8.294, h: 0.909, align: 'center', lineSpacingMultiple: 1
  });
  [
    ['40', 'Auto parts one', 1.379, C.a1],
    ['82', 'Auto parts two', 5.194, C.a2],
    ['60', 'Auto parts three', 9.008, C.a3]
  ].forEach(function (col) {
    const x = col[2];
    gauge(s, x, 2.013, 2.946, col[3]);
    tx(s, [
      { text: col[0], options: { fontSize: 20 } },
      { text: '% data', options: { fontSize: 14 } }
    ], { x: x + 0.871, y: 3.201, w: 1.203, h: 0.616, align: 'center', fontFace: HEAD, bold: true, lineSpacingMultiple: 0.9 });
    shape(s, 'triangle', {
      x: x + 1.31, y: 4.894, w: 0.326, h: 0.326, rotate: 180, fill: { color: col[3] }, line: NOLINE
    });
    tx(s, col[1], { x, y: 5.518, w: 2.946, h: 0.37, align: 'center', fontFace: HEAD, fontSize: 16, bold: true, color: C.ink });
    tx(s, 'PLACEHOLDER', {
      x, y: 5.985, w: 2.946, h: 0.805, align: 'center', fontSize: 12, lineSpacingMultiple: 1.2
    });
  });
}

/* --------------------------------------------------------------- slide 30 */

function slide30 () {
  const s = newSlide();
  swoosh(s, 9.079, -0.645, 4.254, 1.29);
  swoosh(s, 5.617, 3.446, 5.89, 0.607);
  swoosh(s, 4.545, 6.94, 9.069, 0.935);
  swoosh(s, -1.312, 6.368, 3.361, 0.433);
  chevrons(s, 6.219, 1.588);
  tx(s, [
    { text: 'Thank', options: { color: C.a1 } },
    { text: 'You', options: { color: C.ink } }
  ], {
    x: 1.183, y: 1.924, w: 10.967, h: 2.121, align: 'center', fontFace: HEAD, fontSize: 120,
    bold: true, italic: true, charSpacing: -3
  });
  tx(s, 'See you on next presentation!', {
    x: 4.095, y: 4.912, w: 5.143, h: 0.438, align: 'center', fontFace: HEAD, fontSize: 20, color: C.ink
  });
  tx(s, 'PLACEHOLDER' +
    'laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi.', {
    x: 2.64, y: 5.349, w: 8.054, h: 0.563, align: 'center', fontSize: 12, lineSpacingMultiple: 1.2
  });
}

/* ----------------------------------------------------------------- output */

[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30]
  .forEach(function (build) { build(); });

pptx.writeFile({
  fileName: path.join(__dirname, '122138bf-b31d-4d3d-b7b4-fec8b301efdf_grok_final.pptx')
}).then(function (f) { console.log('wrote ' + f); });
