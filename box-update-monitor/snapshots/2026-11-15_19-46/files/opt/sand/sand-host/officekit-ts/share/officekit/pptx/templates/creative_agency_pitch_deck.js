/**
 * "BizClear" creative-agency deck - 24 slides, 13 1/3 x 7.5 in (16:9).
 *
 * Standalone rebuild with pptxgenjs. Run `node <this file>` to emit
 * 19513ccb-cab8-4cbe-aadb-380349609893_grok_final.pptx next to the script.
 *
 * The source deck contains no raster media: every "Picture Placeholder" in it
 * is empty and therefore invisible when rendered, so nothing is drawn for them.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const ORANGE = 'F99E02'; // accent1
const DARK = '262626'; // tx1 lumMod 85%
const GRAY = '808080'; // tx1 lumMod 50%
const WHITE = 'FFFFFF';

const HEAD = 'Montserrat Medium'; // major latin font
const BODY = 'Open Sans'; // minor latin font

const pptx = new PptxGenJS();
// 13 1/3 x 7.5 in == 12192000 x 6858000 EMU, the source deck's exact slide size.
pptx.defineLayout({ name: 'WIDE_16X9', width: 40 / 3, height: 7.5 });
pptx.layout = 'WIDE_16X9';

/* ------------------------------------------------------- shape primitives */

/** Filled rectangle, no outline. `alpha` is 0-100 transparency. */
function rect(s, x, y, w, h, color, alpha) {
  s.addShape(pptx.ShapeType.rect, {
    x: x, y: y, w: w, h: h,
    fill: alpha ? { color: color, transparency: alpha } : { color: color },
    line: { type: 'none' },
  });
}

/** Thin orange decorative circle outline (mostly bleeding off-slide). */
function ring(s, x, y, d) {
  s.addShape(pptx.ShapeType.ellipse, {
    x: x, y: y, w: d, h: d,
    fill: { type: 'none' },
    line: { color: ORANGE, width: 1 },
  });
}

/**
 * The four-square "grid" motif. `g` is the overall box size; the deck always
 * uses squares that are 1/2.4882 of the box with an equal gap between them.
 */
function dots(s, x, y, g, color) {
  const c = g / 2.4882;
  const d = g - c;
  [[0, 0], [d, 0], [d, d], [0, d]].forEach(function (o) {
    rect(s, x + o[0], y + o[1], c, c, color);
  });
}

/** Straight arrow connector. */
function arrow(s, x, y, w, color) {
  s.addShape(pptx.ShapeType.line, {
    x: x, y: y, w: w, h: 0,
    line: { color: color, width: 1, endArrowType: 'triangle' },
  });
}

/**
 * Text box: top-anchored, Open Sans 11pt by default. `fit: 'resize'` mirrors
 * the `<a:spAutoFit/>` carried by every text box in the source deck, which is
 * what keeps unwrapped (`wrap: false`) labels sitting where the original does.
 */
function txt(s, x, y, w, h, runs, opts) {
  s.addText(runs, Object.assign(
    { x: x, y: y, w: w, h: h, valign: 'top', fit: 'resize',
      fontFace: BODY, fontSize: 11, color: DARK, isTextBox: true },
    opts || {}
  ));
}

/** Two-tone section heading: dark/white lead-in + orange tail. */
function title(s, x, y, w, h, lead, tail, opts) {
  const o = opts || {};
  txt(s, x, y, w, h, [
    { text: lead, options: { color: o.color || DARK } },
    { text: tail, options: { color: ORANGE } },
  ], { fontFace: HEAD, fontSize: o.size || 28, align: o.align, wrap: o.wrap !== false });
}

/** Italic letter-spaced kicker under a heading. */
function subtitle(s, x, y, text, color, w) {
  txt(s, x, y, w || 2.286, 0.303, text,
    { fontSize: 12, italic: true, color: color || GRAY, charSpacing: 3, wrap: false });
}

/** 11pt paragraph at 1.5 line spacing - the deck's body-copy style. */
function para(s, x, y, w, h, text, color, align) {
  txt(s, x, y, w, h, text,
    { color: color || GRAY, lineSpacingMultiple: 1.5, align: align });
}

/** Small 14pt Montserrat card heading. */
function cardTitle(s, x, y, w, text, color, align) {
  txt(s, x, y, w, 0.337, text, { fontFace: HEAD, fontSize: 14, color: color, align: align, wrap: false });
}

/** "Learn More" pill: filled block plus centred letter-spaced label. */
function button(s, x, y, w, h, fill, label) {
  rect(s, x, y, w, h, fill);
  txt(s, x + (w - 1.454) / 2, y + (h - 0.286) / 2, 1.454, 0.286, 'Learn More',
    { color: label, charSpacing: 3, align: 'center', wrap: false });
}

/* ---------------------------------------------------------- page furniture */

// Orange quarter-round tab wedged into the top-left corner of every slide.
const CORNER_TAB = [
  { x: 0, y: 0, moveTo: true },
  { x: 0.544, y: 0 },
  { x: 0.557, y: 0.042 },
  { x: 0.573, y: 0.198, curve: { type: 'cubic', x1: 0.567, y1: 0.093, x2: 0.573, y2: 0.144 } },
  { x: 0.102, y: 0.908, curve: { type: 'cubic', x1: 0.573, y1: 0.517, x2: 0.379, y2: 0.791 } },
  { x: 0, y: 0.939 },
  { close: true },
];

const NAV = [
  ['About', 8.42, 0.637],
  ['Service', 9.674, 0.71],
  ['Team', 11.002, 0.603],
  ['Portfolio', 12.222, 0.821],
];

// Per-slide colours for the shared header/footer. `nav` may be one colour for
// all four links or an array of four. `bg` is only set on the dark slides.
const CHROME = {
  1: { bg: DARK, brand: WHITE, nav: WHITE, site: WHITE, page: ORANGE },
  2: { brand: DARK, nav: GRAY, site: GRAY, page: DARK },
  3: { brand: DARK, nav: WHITE, site: GRAY, page: WHITE },
  4: { brand: DARK, nav: GRAY, site: GRAY, page: ORANGE },
  5: { brand: WHITE, nav: GRAY, site: WHITE, page: ORANGE },
  6: { bg: DARK, brand: WHITE, nav: WHITE, site: WHITE, page: ORANGE },
  7: { brand: DARK, nav: GRAY, site: GRAY, page: DARK },
  8: { brand: WHITE, nav: GRAY, site: WHITE, page: WHITE },
  9: { brand: DARK, nav: GRAY, site: WHITE, page: ORANGE },
  10: { brand: DARK, nav: GRAY, site: GRAY, page: WHITE },
  11: { brand: DARK, nav: GRAY, site: GRAY, page: DARK },
  12: { brand: DARK, nav: GRAY, site: GRAY, page: DARK },
  13: { bg: DARK, brand: WHITE, nav: WHITE, site: WHITE, page: WHITE },
  14: { brand: DARK, nav: GRAY, site: GRAY, page: DARK },
  15: { brand: DARK, nav: GRAY, site: GRAY, page: DARK },
  16: { brand: DARK, nav: [GRAY, WHITE, WHITE, WHITE], site: GRAY, page: ORANGE },
  17: { brand: DARK, nav: GRAY, site: GRAY, page: DARK },
  18: { bg: DARK, brand: DARK, nav: WHITE, site: WHITE, page: ORANGE },
  19: { brand: DARK, nav: GRAY, site: GRAY, page: DARK },
  20: { brand: WHITE, nav: GRAY, site: WHITE, page: DARK },
  21: { brand: DARK, nav: [GRAY, GRAY, GRAY, WHITE], site: GRAY, page: DARK },
  22: { brand: DARK, nav: GRAY, site: WHITE, page: ORANGE },
  23: { brand: DARK, nav: GRAY, site: GRAY, page: DARK },
  24: { bg: DARK, brand: WHITE, nav: WHITE, site: WHITE, page: ORANGE },
};

/**
 * Creates slide `n`: page background, then the optional `backdrop` painter
 * (full-bleed panels that must sit *under* the header/footer, as they do in
 * the source deck), then the shared chrome.
 */
function newSlide(n, backdrop) {
  const c = CHROME[n];
  const s = pptx.addSlide();
  if (c.bg) s.background = { color: c.bg };
  if (backdrop) backdrop(s);

  s.addShape(pptx.ShapeType.custGeom, {
    x: 0, y: 0, w: 0.573, h: 0.939,
    fill: { color: ORANGE }, line: { type: 'none' }, points: CORNER_TAB,
  });
  txt(s, 0.78, 0.221, 1.312, 0.286, 'Agency Name', { fontFace: HEAD, color: c.brand, wrap: false });
  NAV.forEach(function (item, i) {
    txt(s, item[1], 0.221, item[2], 0.286, item[0],
      { color: Array.isArray(c.nav) ? c.nav[i] : c.nav, align: 'right', wrap: false });
  });
  txt(s, 0.286, 6.993, 2.258, 0.286, 'Your Website Here', { color: c.site, charSpacing: 3, wrap: false });
  // fit:'none' keeps the box at full width so the number stays flush right.
  txt(s, 12.3, 6.993, 0.743, 0.286, ('0' + n).slice(-2),
    { fontFace: HEAD, color: c.page, align: 'right', wrap: false, fit: 'none' });
  return s;
}

/* ---------------------------------------------- body copy reused verbatim */

const L = {
  // Long "sit amet ... commodo consequat." paragraph (slides 3, 7, 11, 14, 20, 23)
  a: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip commodo consequat. ',
  // Two-line follow-up (slides 3, 11, 23)
  b: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed ullamco laboris nisi ut aliquip commodo consequat. ',
  // Three-line follow-up (slides 7, 14, 20)
  c: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip commodo consequat. ',
  // Card copy (slides 10, 11, 12, 19, 23)
  d: 'Lorem ipsum dolor amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore',
  // Wide centred paragraph (slides 9, 15; slide 2 appends `dolorFugiat`)
  e: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ',
  dolorFugiat: 'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ',
  // "ex commodo consequat" variant (slides 4, 5, 12)
  f: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex commodo consequat. Duis aute irure dolor reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ',
  excepteur: 'Excepteur sint occaecat cupidatat culpa qui officia deserunt mollit anim id est laborum.',
  // Dark-slide intro paragraph (slides 6, 13)
  g: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim minim veniam, quis nostrud exercitation ullamco laboris nisi aliquip commodo aute irure dolor in reprehenderit in voluptate velit esse',
  // Pricing/service tile copy (slides 9, 22)
  h: 'PLACEHOLDER',
  // Narrow column beside a heading (slides 8, 21)
  i: 'Lorem ipsum dolor sit consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore dolore magna aliqua. Ut enim ad minim veniam, nostrud exercitati ullamco laboris nisi ut aliquip ex ea',
  j: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed ullamco laboris nisi ut', // slide 4
  k: 'Lorem ipsum dolor, consectetur adipiscing elit, sed do', // slide 15
  m: 'Lorem ipsum dolor consectetur adipiscing elit, sed do', // slide 17
  n: 'Lorem ipsum dolor consectetur adipiscing eli sed do', // slide 19
  o: 'Lorem ipsum dolor sit amet, consectetur', // slide 5
  p: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod', // slide 8
};

const SUBTITLE = 'Your Subtitle Here';

/* --------------------------------------------------------------- slide 01 */
// Dark cover: oversized "BizClear" wordmark, intro copy and a CTA.
function slide01() {
  const s = newSlide(1);
  ring(s, -1.599, -2.46, 5.363);
  para(s, 5.944, 5.245, 4.662, 1.182,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt labore et dolore magna aliqua. Utenim ad minim veniam, quis nostrud exercitation ullamcosis laboris nisi ut aliquip ex ea commodo consequat. ',
    WHITE);
  button(s, 11.445, 5.553, 1.889, 0.567, ORANGE, WHITE);
  txt(s, 1.313, 5.061, 4.081, 1.212, [
    { text: 'Biz', options: { color: WHITE } },
    { text: 'Clear', options: { color: ORANGE } },
  ], { fontFace: HEAD, fontSize: 66, wrap: false });
  txt(s, 1.388, 6.088, 3.273, 0.337, 'Creative Agency Presnetation',
    { fontSize: 14, color: WHITE, charSpacing: 1, wrap: false });
  dots(s, 1.739, 1.467, 0.352, WHITE);
}

/* --------------------------------------------------------------- slide 02 */
// Table of contents: dark panel with a 2x2 grid of section links.
function slide02() {
  const s = newSlide(2);
  ring(s, -1.599, -2.46, 5.363);
  title(s, 4.868, 0.832, 3.598, 0.572, 'Table Of ', 'Content', { align: 'center', wrap: false });
  para(s, 1.389, 5.764, 10.555, 0.904, L.e + L.dolorFugiat, GRAY, 'center');

  rect(s, 7.971, 2.274, 5.363, 2.619, DARK);
  dots(s, 8.12, 2.424, 0.2, ORANGE);
  dots(s, 8.12, 4.547, 0.2, ORANGE);

  [
    { label: 'About', lw: 0.817, x: 8.507, y: 2.762 },
    { label: 'Service', lw: 0.924, x: 8.512, y: 3.713 },
    { label: 'Team', lw: 0.763, x: 10.83, y: 2.757 },
    { label: 'Portfolio', lw: 1.057, x: 10.83, y: 3.713 },
  ].forEach(function (e) {
    cardTitle(s, e.x, e.y, e.lw, e.label, ORANGE);
    para(s, e.x, e.y + 0.337, 2.017, 0.349, 'Lorem ipsum dolor amet ', WHITE);
  });
}

/* --------------------------------------------------------------- slide 03 */
// "Welcome To Our Presentation" - copy left, dark/orange blocks right.
function slide03() {
  const s = newSlide(3, function (s) {
    rect(s, 6.944, 0, 6.389, 2.635, DARK);
    rect(s, 11.429, 5.992, 1.905, 1.508, ORANGE);
  });
  title(s, 1.289, 1.854, 3.839, 1.043, 'Welcome To Our ', 'Presentation');
  subtitle(s, 1.289, 2.897, SUBTITLE);
  para(s, 1.289, 3.532, 4.437, 1.182, L.a);
  para(s, 1.289, 5.019, 4.437, 0.627, L.b);
  dots(s, 7.783, 1.998, 0.272, ORANGE);
}

/* --------------------------------------------------------------- slide 04 */
// "About Our Creatice Agency" - orange callout card over a dark band.
function slide04() {
  const s = newSlide(4, function (s) {
    ring(s, -1.599, 4.598, 5.363);
    rect(s, 5.185, 5.495, 8.148, 2.005, DARK);
  });
  title(s, 1.206, 1.044, 3.839, 1.043, 'About Our ', 'Creatice Agency');
  subtitle(s, 1.206, 2.087, SUBTITLE);
  para(s, 5.468, 0.987, 6.671, 1.46, L.f + L.excepteur);

  rect(s, 8.148, 3.434, 5.185, 2.352, ORANGE);
  dots(s, 8.311, 3.596, 0.272, WHITE);
  cardTitle(s, 8.682, 3.85, 1.124, 'About Us', WHITE);
  para(s, 8.686, 4.187, 1.896, 1.182, L.j, WHITE);
  para(s, 8.053, 6.007, 4.169, 0.627, L.j, WHITE);
}

/* --------------------------------------------------------------- slide 05 */
// "History Of Our Agency" - full-height dark rail on the left.
function slide05() {
  const s = newSlide(5, function (s) {
    rect(s, 12.016, 6.473, 1.317, 1.027, DARK);
    rect(s, 0, 0, 2.897, 7.5, DARK);
  });
  title(s, 5.624, 1.225, 4.794, 0.572, 'History Of Our ', 'Agency');
  subtitle(s, 5.624, 1.797, SUBTITLE);
  para(s, 5.624, 2.258, 6.669, 1.182, L.f);

  rect(s, 8.021, 4.349, 3.995, 2.124, ORANGE);
  [4.697, 5.439].forEach(function (y) {
    cardTitle(s, 8.5, y, 1.124, 'About Us', WHITE);
    para(s, 8.504, y + 0.336, 3.141, 0.349, L.o, WHITE);
  });
  dots(s, 8.136, 6.125, 0.233, WHITE);
  dots(s, 11.666, 4.464, 0.233, WHITE);
  ring(s, -1.599, 4.598, 5.363);
}

/* --------------------------------------------------------------- slide 06 */
// Dark "The Best Creative Agency" with CTA and arrow.
function slide06() {
  const s = newSlide(6);
  ring(s, -1.599, 4.598, 5.363);
  title(s, 6.92, 1.432, 4.05, 1.043, 'The Best Creative ', 'Agency', { color: WHITE });
  subtitle(s, 6.92, 2.475, SUBTITLE, WHITE);
  para(s, 6.92, 3.082, 5.357, 1.182, L.g, WHITE);
  para(s, 6.92, 4.377, 5.357, 0.627,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod aute irure dolor in reprehenderit in voluptate velit esse',
    WHITE);
  button(s, 7.009, 5.501, 1.889, 0.567, ORANGE, WHITE);
  arrow(s, 10.971, 5.785, 1.172, WHITE);
  dots(s, 1.781, 1.432, 0.33, ORANGE);
}

/* --------------------------------------------------------------- slide 07 */
// "Let's Talk About Our Company" - copy column framed by two dot motifs.
function slide07() {
  const s = newSlide(7);
  ring(s, -2.365, -1.23, 9.96);
  title(s, 1.249, 1.715, 3.648, 1.043, 'Let\u2019s Talk About ', 'Our Company');
  subtitle(s, 1.249, 2.758, SUBTITLE);
  para(s, 1.249, 3.393, 4.437, 1.182, L.a);
  para(s, 1.249, 4.881, 4.437, 0.904, L.c);
  dots(s, 0.748, 1.579, 0.272, ORANGE);
  dots(s, 0.748, 5.947, 0.272, ORANGE);
}

/* --------------------------------------------------------------- slide 08 */
// "Our Vision And Mission" - orange panel with two labelled rows.
function slide08() {
  const s = newSlide(8, function (s) {
    rect(s, 9.213, 5.261, 4.12, 2.239, DARK);
    rect(s, 0, 0, 2.831, 3.386, DARK);
    rect(s, 0, 6.772, 2.831, 0.728, ORANGE);
  });
  title(s, 5.637, 1.434, 3.515, 1.043, 'Our Vision And ', 'Mission');
  subtitle(s, 5.637, 2.478, SUBTITLE);
  para(s, 9.626, 1.239, 2.691, 1.737, L.i);

  rect(s, 6.292, 4.086, 5.842, 2.35, ORANGE);
  dots(s, 11.722, 4.236, 0.262, WHITE);
  [
    { label: 'Our Vision', lw: 1.238, y: 4.711 },
    { label: 'Our Mission', lw: 1.384, y: 5.474 },
  ].forEach(function (e) {
    cardTitle(s, 6.866, e.y, e.lw, e.label, WHITE);
    para(s, 8.473, e.y - 0.145, 3.117, 0.627, L.p, WHITE);
  });
}

/* --------------------------------------------------------------- slide 09 */
// "What We Do?" - three translucent orange service tiles.
const SERVICE_TILES = [
  { x: 1.444, label: 'Creative Agency', lw: 1.838 },
  { x: 4.992, label: 'Content Creator', lw: 1.817 },
  { x: 8.54, label: 'Social Media', lw: 1.454 },
];

function slide09() {
  const s = newSlide(9, function (s) {
    rect(s, 0, 4.151, 13.333, 3.349, DARK);
    SERVICE_TILES.forEach(function (t) { rect(s, t.x, 2.1, 3.349, 3.349, ORANGE, 15); });
  });
  title(s, 4.909, 0.832, 3.515, 0.572, 'What We ', 'Do?', { align: 'center' });
  subtitle(s, 5.523, 1.404, SUBTITLE);
  para(s, 1.389, 6.042, 10.555, 0.627, L.e, WHITE, 'center');

  SERVICE_TILES.forEach(function (t) {
    dots(s, t.x + 1.539, 2.522, 0.272, WHITE);
    cardTitle(s, t.x + (3.349 - t.lw) / 2, 3.646, t.lw, t.label, WHITE, 'center');
    para(s, t.x + 0.334, 3.983, 2.681, 0.904, L.h, WHITE, 'center');
    button(s, t.x + 0.8015, 5.211, 1.746, 0.467, WHITE, GRAY);
  });
}

/* --------------------------------------------------------------- slide 10 */
// "Our Creative Services" - wide dark panel with two service columns.
function slide10() {
  const s = newSlide(10, function (s) {
    ring(s, 9.575, -2.46, 5.363);
    rect(s, 0, 1.111, 8.695, 5.278, DARK);
    rect(s, 12.222, 6.389, 1.111, 1.111, ORANGE);
  });
  title(s, 1.066, 1.824, 4.747, 0.572, 'Our Creative ', 'Services', { color: WHITE });
  subtitle(s, 1.066, 2.397, SUBTITLE, WHITE);
  para(s, 1.066, 2.969, 6.669, 1.182,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ea commodo consequat. Duis aute irure dolor reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ',
    WHITE);
  [
    { x: 1.066, label: 'Creative Agency', lw: 1.838 },
    { x: 4.769, label: 'Content Creator', lw: 1.817 },
  ].forEach(function (e) {
    cardTitle(s, e.x, 4.434, e.lw, e.label, ORANGE);
    para(s, e.x, 4.771, 2.966, 0.904, L.d, WHITE);
  });
  dots(s, 8.12, 1.458, 0.224, ORANGE);
}

/* --------------------------------------------------------------- slide 11 */
// "This Is Our Work Services" - two stacked dark cards on the right.
function slide11() {
  const s = newSlide(11);
  title(s, 1.289, 1.411, 3.419, 1.043, 'This Is Our ', 'Work Services');
  subtitle(s, 1.289, 2.455, SUBTITLE);
  para(s, 1.289, 3.09, 4.437, 1.182, L.a);
  para(s, 1.289, 4.482, 4.437, 0.627, L.b);
  button(s, 1.372, 5.622, 1.746, 0.467, ORANGE, WHITE);
  arrow(s, 4.825, 5.855, 0.821, ORANGE);

  [
    { y: 1.762, label: 'Creative Agency', lw: 1.838 },
    { y: 3.83, label: 'Content Creator', lw: 1.817 },
  ].forEach(function (e) {
    rect(s, 9.571, e.y, 3.762, 1.907, DARK);
    cardTitle(s, 9.983, e.y + 0.334, e.lw, e.label, ORANGE);
    para(s, 9.983, e.y + 0.67, 2.966, 0.904, L.d, WHITE);
    dots(s, 12.931, e.y + 0.194, 0.224, ORANGE);
  });
}

/* --------------------------------------------------------------- slide 12 */
// "Our Excellent Services" - dark banner on top, heading and copy below.
function slide12() {
  const s = newSlide(12);
  ring(s, -1.599, -2.46, 5.363);
  title(s, 1.331, 5.448, 3.419, 1.043, 'Our Excellent ', 'Services');
  para(s, 5.332, 5.378, 6.671, 1.182, L.f);

  rect(s, 5.771, 1.58, 7.562, 2.301, DARK);
  dots(s, 12.714, 4.08, 0.399, ORANGE);
  [
    { x: 6.477, label: 'Creative Agency' },
    { x: 9.828, label: 'Content Creator' },
  ].forEach(function (e) {
    cardTitle(s, e.x, 2.11, 1.838, e.label, ORANGE);
    para(s, e.x, 2.446, 2.966, 0.904, L.d, WHITE);
  });
}

/* --------------------------------------------------------------- slide 13 */
// Dark "Break Slide" divider with an orange 78% stat bar.
function slide13() {
  const s = newSlide(13);
  txt(s, 6.92, 1.26, 3.738, 0.841, [
    { text: 'Break ', options: { color: WHITE } },
    { text: 'Slide', options: { color: ORANGE } },
  ], { fontFace: HEAD, fontSize: 44, wrap: false });
  para(s, 6.92, 2.293, 5.357, 1.182, L.g, WHITE);

  rect(s, 5.802, 4.397, 7.531, 1.627, ORANGE);
  dots(s, 6.521, 5.011, 0.399, WHITE);
  txt(s, 7.399, 4.924, 1.021, 0.572, '78%', { fontFace: HEAD, fontSize: 28, color: WHITE, wrap: false });
  para(s, 8.746, 4.758, 4.175, 0.904,
    'Lorem ipsum dolor sit amet, consectetur adipiscing eli sed do eiusmod tempor incididunt ut labore et dolore magna aliqu enim minim veniam, ',
    WHITE);
}

/* --------------------------------------------------------------- slide 14 */
// "Meet Our Creative Team" - copy column on the right of a big arc.
function slide14() {
  const s = newSlide(14);
  title(s, 7.608, 1.715, 3.648, 1.043, 'Meet Our ', 'Creative Team');
  subtitle(s, 7.608, 2.758, SUBTITLE);
  para(s, 7.608, 3.393, 4.437, 1.182, L.a);
  para(s, 7.608, 4.881, 4.437, 0.904, L.c);
  dots(s, 7.107, 1.579, 0.272, ORANGE);
  dots(s, 7.107, 5.947, 0.272, ORANGE);
  ring(s, 5.706, -1.23, 9.96);
}

/* --------------------------------------------------------------- slide 15 */
// "Our Creative Team" - two mirrored orange name cards on a dark band.
function slide15() {
  const s = newSlide(15);
  rect(s, -0.005, 3.75, 13.338, 3.022, DARK);
  ring(s, 1.567, -8.064, 10.2);
  title(s, 4.262, 0.832, 4.809, 0.572, 'Our Creative ', 'Team', { align: 'center' });
  subtitle(s, 5.523, 1.404, SUBTITLE);
  para(s, 1.389, 5.582, 10.555, 0.627, L.e, WHITE, 'center');
  dots(s, 0.625, 4.615, 0.401, WHITE);
  dots(s, 12.312, 4.615, 0.401, WHITE);

  rect(s, 1.021, 2.885, 3.349, 1.73, ORANGE);
  cardTitle(s, 2.31, 3.268, 1.639, 'Lester Bradley', WHITE, 'right');
  para(s, 1.393, 3.605, 2.539, 0.627, L.k, WHITE, 'right');

  rect(s, 8.964, 2.885, 3.349, 1.73, ORANGE);
  cardTitle(s, 9.405, 3.268, 1.646, 'Kimberly Graff', WHITE);
  para(s, 9.405, 3.605, 2.539, 0.627, L.k, WHITE);
}

/* --------------------------------------------------------------- slide 16 */
// "Lester Bradley" profile with two horizontal skill bars.
function slide16() {
  const s = newSlide(16, function (s) {
    rect(s, 9.319, 0, 4.014, 7.5, DARK);
    ring(s, 9.575, 4.598, 5.363);
  });
  title(s, 1.224, 1.596, 3.072, 0.572, 'Lester ', 'Bradley', { wrap: false });
  subtitle(s, 1.224, 2.168, 'CEO / Founder', GRAY, 1.813);
  para(s, 1.224, 2.7, 4.437, 1.46,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip commodo consequat. Duis aute irure dolor in reprehenderit in voluptate');

  [
    { label: 'Skill One', lw: 0.984, pct: '85%', bar: 3.401, y: 4.437 },
    { label: 'Skill Two', lw: 0.954, pct: '90%', bar: 3.707, y: 5.324 },
  ].forEach(function (e) {
    txt(s, 1.224, e.y, e.lw, 0.303, e.label, { fontFace: HEAD, fontSize: 12, wrap: false });
    txt(s, 5.101, e.y, 0.56, 0.303, e.pct, { fontFace: HEAD, fontSize: 12, align: 'right', wrap: false });
    rect(s, 1.332, e.y + 0.333, 4.227, 0.247, DARK);
    rect(s, 1.332, e.y + 0.333, e.bar, 0.247, ORANGE);
  });
  dots(s, 11.151, 1.546, 0.372, ORANGE);
}

/* --------------------------------------------------------------- slide 17 */
// "Kimberly Graff" profile with a dark two-column skills panel.
function slide17() {
  const s = newSlide(17);
  ring(s, 9.575, -2.46, 5.363);
  rect(s, 1.268, 4.498, 6.248, 1.808, DARK);
  title(s, 1.159, 1.194, 3.087, 0.572, 'Kimberly ', 'Graff', { wrap: false });
  subtitle(s, 1.159, 1.766, 'CEO / Founder', GRAY, 1.813);
  para(s, 1.159, 2.283, 5.443, 1.182,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim minim veniam, quis nostrud exercitation ullamco laboris nisu aliquip commodo aute irure dolor in reprehenderit in voluptate velit esse');

  [{ x: 1.783, label: 'Skill One' }, { x: 4.577, label: 'Skill Two' }].forEach(function (e) {
    txt(s, e.x, 4.921, 1.838, 0.337, e.label, { fontFace: HEAD, fontSize: 14, color: ORANGE });
    para(s, e.x, 5.257, 2.454, 0.627, L.m, WHITE);
  });
  dots(s, 7.144, 4.686, 0.184, ORANGE);
}

/* --------------------------------------------------------------- slide 18 */
// Dark "Our Creative Portfolio" opener with a 78% stat strip.
function slide18() {
  const s = newSlide(18, function (s) {
    ring(s, 9.575, -2.46, 5.363);
    rect(s, 11.928, 0, 1.406, 0.728, ORANGE);
  });
  title(s, 1.269, 1.086, 3.648, 1.043, 'Our Creative ', 'Portfolio', { color: WHITE });
  subtitle(s, 1.269, 2.129, SUBTITLE, WHITE);
  para(s, 5.085, 4.224, 7.074, 1.182,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, do eiusmod tempor incididunt ut labore dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi aliquip ex ea commodo consequat. ' + L.dolorFugiat,
    WHITE);

  rect(s, 8.694, 5.828, 4.64, 0.944, ORANGE);
  dots(s, 9.021, 6.14, 0.319, WHITE);
  txt(s, 9.506, 6.081, 0.787, 0.438, '78%', { fontFace: HEAD, fontSize: 20, color: WHITE, wrap: false });
  para(s, 10.426, 5.986, 2.687, 0.627, 'Lorem ipsum dolor sit consectetur adipiscing eli sed do', WHITE);
}

/* --------------------------------------------------------------- slide 19 */
// "Our Best Portfolio" - dark stats panel left, two orange cards right.
function slide19() {
  const s = newSlide(19);
  title(s, 4.262, 0.832, 4.809, 0.572, 'Our Best ', 'Portfolio', { align: 'center' });
  subtitle(s, 5.523, 1.404, SUBTITLE);

  rect(s, 0, 2.095, 6.952, 4.392, DARK);
  para(s, 0.735, 2.65, 3.416, 1.737,
    'Lorem ipsum dolor sit consectetur adipiscing elit, do eiusmod tempor incididunt labore dolore magna aliqua. Ut enim minim veniam, quis nostrud exercitation ullao laboris nisi aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in',
    WHITE);
  [{ y: 4.641, pct: '81%' }, { y: 5.365, pct: '91%' }].forEach(function (e) {
    txt(s, 0.735, e.y, 0.723, 0.438, e.pct, { fontFace: HEAD, fontSize: 20, color: ORANGE, wrap: false });
    para(s, 1.625, e.y - 0.095, 2.527, 0.627, L.n, WHITE);
  });

  [
    { y: 2.095, label: 'Creative Agency', lw: 1.838 },
    { y: 4.368, label: 'Content Creator', lw: 1.817 },
  ].forEach(function (e) {
    rect(s, 9.405, e.y, 3.929, 2.119, ORANGE);
    cardTitle(s, 9.925, e.y + 0.439, e.lw, e.label, WHITE);
    para(s, 9.925, e.y + 0.776, 2.966, 0.904, L.d, WHITE);
    dots(s, 12.973, e.y + 0.168, 0.193, WHITE);
  });
}

/* --------------------------------------------------------------- slide 20 */
// "Take A Look At Our Portfolio" - dark panel covering the left 62%.
function slide20() {
  const s = newSlide(20, function (s) {
    rect(s, 0, 0, 8.23, 7.5, DARK);
  });
  title(s, 1.227, 1.715, 3.648, 1.043, 'Take A Look At ', 'Our Portfolio', { color: WHITE });
  subtitle(s, 1.227, 2.758, SUBTITLE, WHITE);
  para(s, 1.227, 3.393, 4.437, 1.182, L.a, WHITE);
  para(s, 1.227, 4.881, 4.437, 0.904, L.c, WHITE);
  dots(s, 6.365, 1.256, 0.272, ORANGE);
  dots(s, 6.365, 5.972, 0.272, ORANGE);
}

/* --------------------------------------------------------------- slide 21 */
// "Our Best Project Portfolio" - dark caption card bottom right.
function slide21() {
  const s = newSlide(21, function (s) {
    rect(s, 11.928, 0, 1.406, 0.728, ORANGE);
  });
  title(s, 1.06, 1.446, 3.648, 1.043, 'Our Best Project ', 'Portfolio');
  subtitle(s, 1.06, 2.49, SUBTITLE);
  para(s, 5.17, 1.251, 2.691, 1.737, L.i);

  rect(s, 8.194, 4.474, 5.139, 2.302, DARK);
  dots(s, 8.351, 4.63, 0.193, ORANGE);
  dots(s, 8.351, 6.423, 0.193, ORANGE);
  cardTitle(s, 8.682, 4.866, 1.838, 'Creative Agency', ORANGE);
  para(s, 8.682, 5.197, 2.237, 1.182,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, deiusmod tempor incididunt ut labore et dolore aliqua.',
    WHITE);
}

/* --------------------------------------------------------------- slide 22 */
// "Our Best Pricing Plan" - three translucent orange price tiles.
const PRICE_TILES = [
  { x: 1.444, price: '$15', pw: 0.728, label: 'Package One', lw: 1.54 },
  { x: 4.992, price: '$25', pw: 0.796, label: 'Package Two', lw: 1.538 },
  { x: 8.54, price: '$35', pw: 0.796, label: 'Package Three', lw: 1.692 },
];

function slide22() {
  const s = newSlide(22, function (s) {
    rect(s, 0, 5.028, 13.333, 2.472, DARK);
    ring(s, 1.567, -8.064, 10.2);
  });
  title(s, 4.262, 0.832, 4.809, 0.572, 'Our Best ', 'Pricing Plan', { align: 'center' });
  subtitle(s, 5.523, 1.404, SUBTITLE);

  PRICE_TILES.forEach(function (t) {
    rect(s, t.x, 2.755, 3.349, 3.349, ORANGE, 15);
    dots(s, t.x + 1.539, 3.177, 0.272, WHITE);
    txt(s, t.x + (3.349 - t.pw) / 2, 3.745, t.pw, 0.505, t.price,
      { fontFace: HEAD, fontSize: 24, color: WHITE, align: 'center', wrap: false });
    cardTitle(s, t.x + (3.349 - t.lw) / 2, 4.253, t.lw, t.label, WHITE, 'center');
    para(s, t.x + 0.334, 4.589, 2.681, 0.904, L.h, WHITE, 'center');
    button(s, t.x + 0.8015, 5.866, 1.746, 0.467, WHITE, GRAY);
  });
}

/* --------------------------------------------------------------- slide 23 */
// "Contact Information" - CTA on the left, two dark contact cards right.
function slide23() {
  const s = newSlide(23);
  title(s, 1.289, 1.649, 4.486, 0.572, 'Contact ', 'Information');
  subtitle(s, 1.289, 2.217, SUBTITLE);
  para(s, 1.289, 2.852, 4.437, 1.182, L.a);
  para(s, 1.289, 4.244, 4.437, 0.627, L.b);
  button(s, 1.372, 5.384, 1.746, 0.467, ORANGE, WHITE);
  arrow(s, 4.825, 5.617, 0.821, ORANGE);

  [
    { y: 1.558, label: 'Phone Number', lw: 1.743 },
    { y: 3.831, label: 'Our Address', lw: 1.441 },
  ].forEach(function (e) {
    rect(s, 6.926, e.y, 3.929, 2.119, DARK);
    cardTitle(s, 7.447, e.y + 0.439, e.lw, e.label, ORANGE);
    para(s, 7.447, e.y + 0.775, 2.966, 0.904, L.d, WHITE);
    dots(s, 10.494, e.y + 0.167, 0.193, ORANGE);
  });
}

/* --------------------------------------------------------------- slide 24 */
// Dark closing slide mirroring the cover: "Thank You".
function slide24() {
  const s = newSlide(24);
  ring(s, -1.599, 4.598, 5.363);
  para(s, 7.043, 1.191, 3.865, 1.182,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore dolore magna aliqua. Ut enim adminim venia, quis nostrud exercitation ullamco laboris nisi',
    WHITE);
  button(s, 11.445, 1.498, 1.889, 0.567, ORANGE, WHITE);
  txt(s, 1.313, 1.006, 5.193, 1.212, [
    { text: 'Thank ', options: { color: WHITE } },
    { text: 'You', options: { color: ORANGE } },
  ], { fontFace: HEAD, fontSize: 66, wrap: false });
  txt(s, 1.388, 2.033, 3.471, 0.337, 'For Watching This Presnetation',
    { fontSize: 14, color: WHITE, charSpacing: 1, wrap: false });
  dots(s, 1.739, 3.167, 0.352, WHITE);
}

/* ------------------------------------------------------------------- main */

[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
  slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
  slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24]
  .forEach(function (build) { build(); });

const outFile = path.join(__dirname, '19513ccb-cab8-4cbe-aadb-380349609893_grok_final.pptx');
pptx.writeFile({ fileName: outFile }).then(function () {
  console.log('wrote ' + outFile);
});
