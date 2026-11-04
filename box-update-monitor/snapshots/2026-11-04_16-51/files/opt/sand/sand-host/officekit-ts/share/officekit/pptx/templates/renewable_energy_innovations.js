/**
 * Recreation of "Energy" presentation template (20 slides, 16:9 / 13.333 x 7.5 in)
 * with pptxgenjs. Raster photos in the original deck are replaced by flat
 * placeholder rectangles labelled "[image]".
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */
const GREEN = '004231'; // accent1  - deep forest green
const MINT = 'B3FCA7'; // accent2  - light mint
const NEAR_BLACK = '0C0C0C'; // accent3
const DARK_GREY = '262626'; // accent4
const MID_GREY = '3F3F3F'; // accent5
const WHITE = 'FFFFFF';
const INK = '0D0D0D'; // tx1 lumMod 95% - body copy colour
const TRACK_LIGHT = 'F2F2F2'; // white, lum -5%
const TRACK_GREY = 'BFBFBF'; // white, lum -25%

const HEAD = 'Fira Sans'; // major latin font
const BODY = 'Work Sans'; // minor latin font

const W = 13.333;
const H = 7.5;

/* --------------------------------------------------------------- helpers */

/**
 * Five faint vertical rules that decorate every slide: accent green at 25%
 * opacity on the light slides, solid mint on the dark cover.
 */
function stripes(slide, x0, o) {
  const opt = o || {};
  for (let i = 0; i < 5; i++) {
    slide.addShape('line', {
      x: x0 + i * 0.8288, y: 0, w: 0, h: H,
      line: { color: opt.color || GREEN, width: 1, transparency: opt.transparency === undefined ? 75 : opt.transparency },
    });
  }
}

/** Page number, bottom right (comes from the slide master in the original). */
function pageNumber(slide, n) {
  slide.addText(String(n), {
    x: 12.749, y: 7.055, w: 0.462, h: 0.286,
    fontFace: HEAD, fontSize: 11, color: INK, align: 'center', valign: 'top',
  });
}

/** Plain top-anchored text box (the deck's default). */
function text(slide, str, o) {
  slide.addText(str, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: o.face || HEAD,
    fontSize: o.size || 18,
    color: o.color || INK,
    align: o.align || 'left',
    valign: o.valign || 'top',
    bold: !!o.bold,
    lineSpacingMultiple: o.ls,
  });
}

/** 14pt paragraph copy (line spacing 130%). */
function para(slide, str, x, y, w, h, color) {
  text(slide, str, { x, y, w, h, face: BODY, size: 14, color: color || INK, ls: 1.3 });
}

/** Big display heading (line spacing 90%). */
function heading(slide, str, x, y, w, h, size, align) {
  text(slide, str, { x, y, w, h, size, color: INK, ls: 0.9, align });
}

/** Rounded rectangle; radius given in inches. */
function card(slide, x, y, w, h, fill, radius, opts) {
  slide.addShape('roundRect', Object.assign({
    x, y, w, h, fill: { color: fill }, line: { type: 'none' },
    rectRadius: radius === undefined ? 0.25 : radius,
  }, opts || {}));
}

/** Circled check mark used as a bullet glyph. */
function checkBullet(slide, x, y, color) {
  const d = 0.217;
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color } });
  slide.addShape('custGeom', {
    x: x + 0.045, y: y + 0.062, w: 0.127, h: 0.093,
    points: [{ x: 0, y: 0.05 }, { x: 0.045, y: 0.093 }, { x: 0.127, y: 0 }],
    line: { color: color === MINT ? GREEN : WHITE, width: 1.5 },
  });
}

/**
 * Picture frame. Every picture placeholder in the source deck is empty — no
 * photograph is embedded anywhere in it — so each frame is recreated the way
 * the template paints an unfilled placeholder: a barely-there wash of accent1
 * (or grey, on the layouts that use bg1) behind nothing.
 */
function imagePlaceholder(slide, x, y, w, h, o) {
  const opt = o || {};
  slide.addShape('roundRect', {
    x, y, w, h, rectRadius: opt.radius === undefined ? 0.24 : opt.radius,
    fill: { color: opt.grey ? '808080' : GREEN, transparency: 96 },
    line: { type: 'none' },
  });
}

/** Stat block: big number, label, paragraph — used on many slides. */
function statBlock(slide, s) {
  text(slide, s.value, { x: s.x, y: s.y, w: s.vw || 1.989, h: s.vh || 0.841, size: s.vsize || 44, color: s.numColor });
  if (s.label) text(slide, s.label, { x: s.x, y: s.labelY, w: s.lw || 2.286, h: 0.404, color: s.labelColor });
  if (s.body) para(slide, s.body, s.x, s.bodyY, s.bw || 2.314, s.bh || 0.992, s.bodyColor);
}

/* ------------------------------------------------------------- copy text */
const LOREM_LONG =
  'Lorem ipsum dolor amet, consectur adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus.';
const LOREM_MED = 'Lorem ipsum dolor amet, consectur adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. ';
const LOREM_CARD = 'Lorem ips dolor amet, consectur adipiscing elit. Aenean com.';
const LOREM_SHORT = 'Lorem ipsum dolor amet consec adipiscing elit. ';
const LOREM_BULLET = 'Lorem ipsum dolor amet consectur adipiscing';
const LOREM_WIDE =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa Cum sociis natoque penatibus et magnis dis parturient montes, nasceturCum sociis natoque penatibus.';

/* ---------------------------------------------------------------- slides */

function slide01(pptx) {
  const s = pptx.addSlide();
  s.background = { color: GREEN };
  stripes(s, 9.241, { color: MINT, transparency: 0 });
  text(s, 'Energy', { x: 0.427, y: 0.801, w: 7.985, h: 2.895, size: 166, color: WHITE });
  para(s, LOREM_LONG, 0.651, 3.948, 6.862, 0.686, WHITE);
  card(s, 0.735, 5.191, 4.245, 0.633, MINT, 0.3165);
  text(s, 'Presentation Template', {
    x: 1.009, y: 5.255, w: 3.696, h: 0.505, size: 24, color: GREEN, align: 'center',
  });
}

function slide02(pptx) {
  const s = pptx.addSlide();
  stripes(s, 9.241);
  heading(s, 'Innovations in Renewable Energy', 0.765, 0.601, 7.95, 2.1, 66);
  text(s, 'Energizing the Future', { x: 9.051, y: 1.042, w: 3.256, h: 0.404 });
  para(s, LOREM_MED, 9.051, 1.561, 3.639, 0.992);
  imagePlaceholder(s, 0.76, 3.75, 5.906, 3.149);
  card(s, 6.963, 3.75, 2.649, 3.149, GREEN, 0.2412);
  card(s, 9.908, 3.75, 2.649, 3.149, MINT, 0.2412);
  statBlock(s, {
    x: 7.13, y: 4.057, value: '75,2%', numColor: WHITE, label: 'Innovation One', labelY: 5.081,
    labelColor: WHITE, body: LOREM_CARD, bodyY: 5.6, bodyColor: WHITE,
  });
  statBlock(s, {
    x: 10.075, y: 4.057, value: '80,5%', numColor: GREEN, label: 'Innovation Two', labelY: 5.081,
    labelColor: GREEN, body: LOREM_CARD, bodyY: 5.6,
  });
  pageNumber(s, 2);
}

function slide03(pptx) {
  const s = pptx.addSlide();
  stripes(s, 9.241);
  heading(s, 'Advances in Clean Energy Technologies', 0.765, 0.601, 9.584, 2.1, 66);
  imagePlaceholder(s, 7.016, 3.75, 5.557, 3.149);
  const rows = [
    { y: 3.75, x: 0.765, num: '01', bg: GREEN, numColor: WHITE, title: 'The Energy Equation', bw: 4.998 },
    { y: 5.548, x: 0.725, num: '02', bg: MINT, numColor: GREEN, title: 'Energizing the Future', bw: 5.038 },
  ];
  rows.forEach((r) => {
    s.addShape('ellipse', { x: r.x, y: r.y, w: 0.59, h: 0.59, fill: { color: r.bg } });
    text(s, r.num, { x: r.x + 0.046, y: r.y + 0.093, w: 0.498, h: 0.404, color: r.numColor, align: 'center' });
    text(s, r.title, { x: r.x + 0.789, y: r.y + 0.01, w: 2.767, h: 0.404 });
    para(s, 'Lorem ipsum dolor amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. ',
      r.x + 0.789, r.y + 0.564, r.bw, 0.686);
  });
  pageNumber(s, 3);
}

function slide04(pptx) {
  const s = pptx.addSlide();
  stripes(s, 0.76);
  heading(s, 'From Fossil Fuels to Green Energy', 0.76, 0.601, 6.748, 1.919, 60);
  para(s, 'Lorem ipsum dolor amet consectetur adipiscing elit. Aenean commo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus.',
    0.76, 2.777, 6.851, 0.686);
  imagePlaceholder(s, 8.369, 0.601, 4.964, 6.899, { radius: 0.3, grey: true });
  const bars = [
    { y: 4.226, bg: GREEN, value: '76,2%', numColor: WHITE, bodyColor: WHITE },
    { y: 5.728, bg: MINT, value: '80,5%', numColor: GREEN, bodyColor: INK },
  ];
  bars.forEach((b) => {
    // pill-ended banner running off the left edge of the slide
    s.addShape('roundRect', {
      x: 0.002, y: b.y, w: 7.609, h: 1.171, fill: { color: b.bg },
      line: { type: 'none' }, rectRadius: 0.5855,
    });
    s.addShape('rect', { x: 0.002, y: b.y, w: 1.0, h: 1.171, fill: { color: b.bg }, line: { type: 'none' } });
    text(s, b.value, { x: 0.472, y: b.y + 0.198, w: 1.764, h: 0.774, size: 40, color: b.numColor });
    para(s, 'Lorem ipsum dolor amet consectetur adipiscing elit. Aenean commo ligula eget dolor. Aenean',
      2.296, b.y + 0.243, 4.845, 0.686, b.bodyColor);
  });
  pageNumber(s, 4);
}

function slide05(pptx) {
  const s = pptx.addSlide();
  stripes(s, 0.76);
  heading(s, 'The Role of Batteries in Energy Storage', 0.76, 0.601, 8.176, 1.919, 60);
  imagePlaceholder(s, 9.444, 0, 3.889, 7.5, { radius: 0, grey: true });
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus magnis dis parturient montes.',
    0.76, 6.214, 8.684, 0.686);
  const cards = [
    { x: 0.76, bg: GREEN, fg: WHITE, title: 'Energy Economics', num: '01' },
    { x: 5.911, bg: MINT, fg: GREEN, title: 'Energy Efficiency', num: '02' },
  ];
  cards.forEach((c) => {
    card(s, c.x, 3.623, 4.643, 2.237, c.bg, 0.2724);
    text(s, c.title, { x: c.x + 0.35, y: 4.218, w: 2.951, h: 0.404, color: c.fg });
    para(s, 'Lorem ipsum dolor amet consectetuer adipiscing elit. Aenean commo ligula eget dolor. Aenean massa sociis.',
      c.x + 0.35, 4.621, 3.944, 0.992, c.fg);
    text(s, c.num, { x: c.x + 3.301, y: 3.87, w: 0.918, h: 0.64, size: 32, color: c.fg, align: 'right' });
  });
  pageNumber(s, 5);
}

function slide06(pptx) {
  const s = pptx.addSlide();
  stripes(s, 9.268);
  imagePlaceholder(s, 0.75, 0.601, 2.926, 6.299, { radius: 0.3 });
  imagePlaceholder(s, 8.935, 3.75, 3.637, 3.149);
  heading(s, 'Efficiency Meets Sustainability', 3.924, 0.601, 7.012, 2.1, 66);
  card(s, 4.025, 3.761, 3.434, 0.466, GREEN, 0.233);
  text(s, 'The Renewable Revolution', { x: 4.131, y: 3.792, w: 3.221, h: 0.404, color: WHITE, align: 'center' });
  para(s, 'Lorem ipsum dolor amet, consectur adipiscing elit. Aenean com ligula eget dolor Aenean.',
    3.924, 4.391, 4.662, 0.686);
  [{ y: 5.17, dot: GREEN, value: '72,3%' }, { y: 6.063, dot: MINT, value: '84,2%' }].forEach((r) => {
    checkBullet(s, 4.025, r.y + 0.234, r.dot);
    text(s, r.value, { x: 4.32, y: r.y + 0.023, w: 1.497, h: 0.64, size: 32 });
    para(s, 'Lorem ipsum dolor amet consectur adipiscing.', 5.777, r.y, 2.661, 0.686);
  });
  pageNumber(s, 6);
}

function slide07(pptx) {
  const s = pptx.addSlide();
  stripes(s, 9.268);
  heading(s, 'The Role of Clean Energy', 0.76, 0.601, 6.167, 2.1, 66);
  para(s, 'Lorem ipsum dolor amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa Cum sociis.',
    0.769, 2.903, 6.167, 0.686);
  imagePlaceholder(s, 7.356, 0.601, 5.228, 2.988);
  const cols = [
    { x: 0.76, w: 3.779, bg: GREEN, fg: WHITE, value: '52,9%' },
    { x: 4.765, w: 3.779, bg: MINT, fg: GREEN, value: '62,9%', bodyColor: INK },
    { x: 8.77, w: 3.814, bg: GREEN, fg: WHITE, value: '80,5%' },
  ];
  cols.forEach((c) => {
    card(s, c.x, 4.662, c.w, 2.237, c.bg, 0.2724);
    text(s, c.value, { x: c.x + 0.244, y: 4.868, w: 1.764, h: 0.774, size: 40, color: c.fg });
    para(s, 'Lorem ipsum dolor amet consec adipiscing elit. Aene commodo ligula eget dolor. Aenean.',
      c.x + 0.237, 5.701, 3.305, 0.992, c.bodyColor || c.fg);
  });
  pageNumber(s, 7);
}

function slide08(pptx) {
  const s = pptx.addSlide();
  stripes(s, 0.76);
  heading(s, 'The Rise of Wind Energy', 0.76, 0.601, 6.167, 2.1, 66);
  card(s, 6.667, 0.601, 5.906, 2.1, GREEN, 0.2531);
  text(s, 'Opportunities in Nuclear Energy', { x: 7.073, y: 0.932, w: 4.772, h: 0.404, color: WHITE });
  [1.433, 1.99].forEach((y) => {
    checkBullet(s, 7.14, y + 0.081, MINT);
    para(s, LOREM_BULLET, 7.459, y, 4.707, 0.38, WHITE);
  });
  imagePlaceholder(s, 0.75, 3.75, 11.833, 3.139);
  pageNumber(s, 8);
}

function slide09(pptx) {
  const s = pptx.addSlide();
  stripes(s, 9.268);
  heading(s, 'Exploring Sustainable Solutions Energy', 4.1, 0.601, 8.473, 1.919, 60);
  imagePlaceholder(s, 0.75, 0.601, 3.01, 6.299, { radius: 0.3 });
  imagePlaceholder(s, 4.229, 3.75, 3.01, 3.149);
  const rows = [
    { x: 7.764, y: 3.75, tx: 8.034, bg: GREEN, fg: WHITE, num: '01.' },
    { x: 7.753, y: 5.483, tx: 8.024, bg: MINT, fg: GREEN, num: '02.', bodyColor: INK },
  ];
  rows.forEach((r) => {
    card(s, r.x, r.y, 4.819, 1.417, r.bg, 0.2362);
    text(s, r.num, { x: r.tx, y: r.y + 0.254, w: 1.085, h: 0.909, size: 48, color: r.fg });
    para(s, 'Lorem ipsum dolor amet conse adipiscing elit. Aene commo.',
      r.tx + 1.082, r.y + 0.365, 3.197, 0.686, r.bodyColor || r.fg);
  });
  pageNumber(s, 9);
}

function slide10(pptx) {
  const s = pptx.addSlide();
  stripes(s, 9.268);
  imagePlaceholder(s, 0.76, 0.601, 4.381, 6.299, { radius: 0.3 });
  heading(s, 'Powering Growth Without Pollution', 5.592, 0.601, 6.981, 1.919, 60);
  para(s, LOREM_LONG, 5.592, 2.633, 6.981, 0.686);
  const cards = [
    { cx: 5.639, tx: 5.895, bg: GREEN, fg: WHITE, r: 0.253, value: '75,2%', vw: 1.676, label: 'Energy Revolution' },
    { cx: 9.082, tx: 9.338, bg: MINT, fg: GREEN, r: 0.295, value: '80,5%', vw: 1.828, label: 'Harnessing Nature', textColor: INK },
  ];
  cards.forEach((c) => {
    card(s, c.cx, 4.004, 3.139, 2.891, c.bg, c.r);
    text(s, c.value, { x: c.tx, y: 4.399, w: c.vw, h: 0.774, size: 40, color: c.fg });
    text(s, c.label, { x: c.tx, y: 5.282, w: 2.444, h: 0.403, color: c.textColor || c.fg });
    para(s, LOREM_SHORT, c.tx, 5.814, 2.627, 0.686, c.textColor || c.fg);
  });
  pageNumber(s, 10);
}

function slide11(pptx) {
  const s = pptx.addSlide();
  stripes(s, 9.268);
  imagePlaceholder(s, 0.76, 0.601, 4.906, 3.149);
  heading(s, 'Transitioning to Clean Energy', 6.061, 0.816, 6.523, 1.919, 60);
  para(s, 'Lorem ipsum dolor amet consectur adipiscing elit. Aenean com ligula eget dolor. Aene massa. Cum sociis natoque penatibus.',
    6.061, 2.849, 6.523, 0.686);
  const cards = [
    { cx: 0.76, tx: 1.016, bg: GREEN, fg: WHITE, value: '5492+', vw: 2.143, label: 'Energy Revolution' },
    { cx: 4.26, tx: 4.516, bg: MINT, fg: GREEN, value: '7329+', vw: 1.934, label: 'Energy Innovations', textColor: INK },
  ];
  cards.forEach((c) => {
    card(s, c.cx, 4.383, 3.139, 2.514, c.bg, 0.22);
    text(s, c.value, { x: c.tx, y: 4.59, w: c.vw, h: 0.774, size: 40, color: c.fg });
    text(s, c.label, { x: c.tx, y: 5.473, w: 2.444, h: 0.403, color: c.textColor || c.fg });
    para(s, LOREM_SHORT, c.tx, 6.005, 2.627, 0.686, c.textColor || c.fg);
  });
  text(s, '2025', { x: 7.76, y: 4.223, w: 4.016, h: 1.111, size: 60 });
  text(s, 'Commitment to Sustainability', { x: 7.746, y: 5.447, w: 4.837, h: 0.404 });
  [5.996, 6.494].forEach((y) => {
    checkBullet(s, 7.859, y + 0.082, GREEN);
    para(s, 'Lorem ipsum dolor amet conse adipiscing.', 8.179, y, 4.335, 0.38);
  });
  pageNumber(s, 11);
}

function slide12(pptx) {
  const s = pptx.addSlide();
  stripes(s, 9.268);
  imagePlaceholder(s, 0.75, 0.601, 4.494, 3.149);
  imagePlaceholder(s, 8.521, 4.381, 4.062, 2.518);
  heading(s, 'The Renewable Energy Puzzle', 5.727, 0.816, 6.523, 1.919, 60);
  para(s, 'Lorem ipsum dolor amet consectur adipiscing elit. Aenean com ligula eget dolor. Aene massa. Cum sociis natoque penatibus.',
    5.727, 2.849, 6.523, 0.686);
  const cards = [
    { x: 0.76, bg: GREEN, fg: WHITE, dot: MINT, title: 'Global Energy Trends' },
    { x: 4.641, bg: MINT, fg: GREEN, dot: GREEN, title: 'The New Age of Energy', bodyColor: INK },
  ];
  cards.forEach((c) => {
    card(s, c.x, 4.379, 3.583, 2.518, c.bg, 0.2204);
    text(s, c.title, { x: c.x + 0.301, y: 4.68, w: 2.901, h: 0.404, color: c.fg });
    [5.154, 5.91].forEach((y) => {
      checkBullet(s, c.x + 0.374, y + 0.081, c.dot);
      para(s, LOREM_BULLET, c.x + 0.693, y, 2.59, 0.686, c.bodyColor || c.fg);
    });
  });
  pageNumber(s, 12);
}

function slide13(pptx) {
  const s = pptx.addSlide();
  imagePlaceholder(s, 0, 4.08, 12.573, 3.42, { radius: 0.8295, grey: true });
  s.addShape('round1Rect', {
    x: 0, y: 4.08, w: 12.573, h: 3.42, rectRadius: 0.8295,
    fill: { color: GREEN, transparency: 20 }, line: { type: 'none' },
  });
  stripes(s, 9.268);
  heading(s, 'Integrating Technology for Better Outcomes', 0.75, 0.601, 9.273, 1.919, 60);
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes nascetur.',
    0.75, 2.633, 9.273, 0.686);
  [['62,5%', 0.76, 0.75], ['75,8%', 4.225, 4.215], ['87,4%', 7.69, 7.68]].forEach(([value, vx, bx]) => {
    text(s, value, { x: vx, y: 4.798, w: 2.855, h: 0.909, size: 48, color: MINT });
    para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula.',
      bx, 5.809, 3.194, 0.992, WHITE);
  });
  pageNumber(s, 13);
}

function slide14(pptx) {
  const s = pptx.addSlide();
  stripes(s, 9.268);
  imagePlaceholder(s, 0, 0.601, 4.842, 6.899, { radius: 0.3, grey: true });
  heading(s, 'Resilient Energy Systems for a Changing Climate', 5.429, 0.764, 7.144, 2.827, 60);
  para(s, 'Lorem ipsum dolor amet consectur adipiscing elit. Aenean com ligula eget dolor. Aene massa. Cum sociis natoque penatibus.',
    5.429, 3.804, 7.252, 0.686);
  card(s, 5.556, 5.347, 6.847, 1.552, GREEN, 0.2971);
  text(s, '80,5%', { x: 6.024, y: 5.669, w: 2.072, h: 0.909, size: 48, color: MINT });
  text(s, 'Energy Access and Equity', { x: 8.195, y: 5.578, w: 3.74, h: 0.404, face: BODY, color: WHITE });
  para(s, 'Lorem ipsum dolor amet consectur adipiscing elit. Aenean com ligula.', 8.195, 5.982, 3.74, 0.686, WHITE);
  pageNumber(s, 14);
}

/**
 * Slide 15 — four-armed pinwheel. Each quadrant is built from three flat
 * pieces: a horizontal arm ending in a point, a vertical stem whose far end is
 * cut at 45 degrees, and a diagonal bar bridging the two, so each pair reads
 * as a stylised "4". The base outlines below point left/up and are mirrored
 * per quadrant.
 */
const ARM_OUTLINE = [[0, 0.09], [1, 0], [1, 1], [0.2, 0.95]];
const STEM_OUTLINE = [[0, 0], [1, 0.19], [1, 1], [0, 1]];

function mirrored(outline, flipX, flipY) {
  return outline
    .map(([x, y]) => ({ x: flipX ? 1 - x : x, y: flipY ? 1 - y : y }))
    .concat([{ close: true }]);
}

/** Draw one outline scaled into the box (x, y, w, h). */
function polygon(s, outline, box, color, flipX, flipY) {
  s.addShape('custGeom', {
    x: box[0], y: box[1], w: box[2], h: box[3],
    fill: { color }, line: { type: 'none' },
    points: mirrored(outline, flipX, flipY).map((p) => (p.close
      ? p : { x: p.x * box[2], y: p.y * box[3] })),
  });
}

const PINWHEEL = [
  { color: '003A2B', flipX: false, flipY: false, arm: [4.889, 3.528, 1.778, 0.415], stem: [5.930, 2.518, 0.356, 1.801], bar: [4.801, 2.995, 315] },
  { color: '4FE04A', flipX: true, flipY: false, arm: [6.664, 3.528, 1.780, 0.415], stem: [7.048, 2.515, 0.356, 1.795], bar: [7.068, 2.992, 45] },
  { color: '0A0A0A', flipX: true, flipY: true, arm: [6.664, 4.665, 1.780, 0.355], stem: [7.048, 4.318, 0.356, 1.746], bar: [7.068, 5.235, 135] },
  { color: '1E1E1E', flipX: false, flipY: true, arm: [4.887, 4.666, 1.781, 0.355], stem: [5.927, 4.318, 0.356, 1.746], bar: [4.799, 5.235, 225] },
];

// small pill caps where the green/black stems butt together in the middle
const PINWHEEL_CAPS = [
  { x: 6.001, y: 4.019, rotate: 90, color: GREEN },
  { x: 6.001, y: 4.224, rotate: 90, color: DARK_GREY },
  { x: 7.131, y: 4.007, rotate: 90, color: MINT },
  { x: 7.131, y: 4.212, rotate: 90, color: NEAR_BLACK },
  { x: 6.461, y: 4.633, rotate: 0, color: DARK_GREY },
  { x: 6.667, y: 4.633, rotate: 0, color: NEAR_BLACK },
];

function pinwheel(s) {
  PINWHEEL.forEach((q) => {
    s.addShape('rect', {
      x: q.bar[0], y: q.bar[1], w: 1.464, h: 0.352, rotate: q.bar[2],
      fill: { color: q.color }, line: { type: 'none' },
    });
    polygon(s, STEM_OUTLINE, q.stem, q.color, q.flipX, q.flipY);
    polygon(s, ARM_OUTLINE, q.arm, q.color, q.flipX, q.flipY);
  });
  PINWHEEL_CAPS.forEach((c) => s.addShape('roundRect', {
    x: c.x, y: c.y, w: 0.208, h: 0.415, rotate: c.rotate, rectRadius: 0.104,
    fill: { color: c.color }, line: { type: 'none' },
  }));
}

function slide15(pptx) {
  const s = pptx.addSlide();
  stripes(s, 5.009);
  heading(s, 'Global Energy Trends', 1.954, 0.611, 9.634, 1.101, 66, 'center');
  pinwheel(s);
  const badges = [
    { x: 4.976, y: 2.63, tx: 5.062, ty: 2.758, num: '01', color: GREEN },
    { x: 7.697, y: 2.644, tx: 7.783, ty: 2.773, num: '02', color: MINT },
    { x: 7.697, y: 5.249, tx: 7.783, ty: 5.378, num: '03', color: NEAR_BLACK },
    { x: 4.976, y: 5.249, tx: 5.062, ty: 5.378, num: '04', color: DARK_GREY },
  ];
  badges.forEach((b) => {
    s.addShape('ellipse', { x: b.x, y: b.y, w: 0.661, h: 0.661, fill: { color: WHITE } });
    text(s, b.num, { x: b.tx, y: b.ty, w: 0.488, h: 0.404, color: b.color, align: 'center' });
  });
  const blurbs = [
    { x: 8.969, y: 2.663, align: 'left', title: 'Enegery Two' },
    { x: 8.969, y: 5.184, align: 'left', title: 'Enegery Three' },
    { x: 1.126, y: 2.663, align: 'right', title: 'Enegery Two' },
    { x: 1.126, y: 5.184, align: 'right', title: 'Enegery Four' },
  ];
  blurbs.forEach((b) => {
    text(s, b.title, { x: b.x, y: b.y, w: 3.295, h: 0.404, face: BODY, align: b.align });
    text(s, 'Lorem ipsum dolor amet conse adipiscing elit. Aenean comod ligula eget dolo Aenean', {
      x: b.x, y: b.y + 0.432, w: 3.295, h: 0.992, face: BODY, size: 14, ls: 1.3, align: b.align,
    });
  });
  pageNumber(s, 15);
}

function slide16(pptx) {
  const s = pptx.addSlide();
  stripes(s, 5.009);
  heading(s, 'Sustainable Energy Systems', 1.187, 0.611, 10.959, 1.01, 60, 'center');
  const cards = [
    { x: 1.267, y: 2.631, bg: NEAR_BLACK, fg: WHITE, track: TRACK_GREY, bar: WHITE, suffix: 'A', valueAlign: 'left' },
    { x: 6.789, y: 2.635, bg: MINT, fg: GREEN, track: TRACK_LIGHT, bar: GREEN, suffix: 'B', valueAlign: 'right' },
    { x: 1.267, y: 4.871, bg: GREEN, fg: WHITE, track: TRACK_LIGHT, bar: MINT, suffix: 'C', valueAlign: 'right' },
    { x: 6.789, y: 4.871, bg: DARK_GREY, fg: WHITE, track: TRACK_GREY, bar: WHITE, suffix: 'D', valueAlign: 'right' },
  ];
  const rows = [
    { dy: 0, label: 'Value Energy A', value: '75%', frac: 3.11 / 4.65 },
    { dy: 0.547, label: 'Value Energy B', value: '36%', frac: 1.879 / 4.65 },
  ];
  cards.forEach((c) => {
    card(s, c.x, c.y, 5.278, 2.029, c.bg, 0.127);
    s.addText('Resilient Energy Systems ' + c.suffix, {
      x: c.x + 0.247, y: c.y + 0.248, w: 4.272, h: 0.37,
      fontFace: HEAD, fontSize: 16, color: c.fg, valign: 'bottom',
    });
    rows.forEach((r) => {
      s.addText(r.label, {
        x: c.x + 0.222, y: c.y + 0.772 + r.dy, w: 1.983, h: 0.303,
        fontFace: BODY, fontSize: 12, color: c.fg, valign: 'top',
      });
      s.addText(r.value, {
        x: c.x + 4.182, y: c.y + 0.772 + r.dy, w: 0.766, h: 0.303,
        fontFace: BODY, fontSize: 12, color: c.fg, align: c.valueAlign, valign: 'top',
      });
      s.addShape('rect', { x: c.x + 0.326, y: c.y + 1.134 + r.dy, w: 4.65, h: 0.079, fill: { color: c.track } });
      s.addShape('rect', { x: c.x + 0.326, y: c.y + 1.134 + r.dy, w: 4.65 * r.frac, h: 0.079, fill: { color: c.bar } });
    });
  });
  pageNumber(s, 16);
}

/** Slide 17 — three semicircular double gauges with elbow callouts. */
function gauge(s, g) {
  s.addShape('arc', {
    x: g.x, y: g.y, w: 2.228, h: 2.228, rotate: 315,
    angleRange: g.outer, line: { color: MINT, width: 34.75 }, fill: { type: 'none' },
  });
  s.addShape('arc', {
    x: g.x + 0.128, y: g.y + 0.128, w: 1.972, h: 1.972, rotate: 315,
    angleRange: g.inner, line: { color: GREEN, width: 16.25 }, fill: { type: 'none' },
  });
}

/**
 * Leader line: a right-angled polyline with a dot at each end. `corner` names
 * where the bend sits, so 'tl' runs up the left side then right along the top.
 */
const ELBOW_W = 1.018;
const ELBOW_H = 0.77;
const ELBOW_PATHS = {
  tl: [[0, ELBOW_H], [0, 0], [ELBOW_W, 0]],
  bl: [[0, 0], [0, ELBOW_H], [ELBOW_W, ELBOW_H]],
  br: [[ELBOW_W, 0], [ELBOW_W, ELBOW_H], [0, ELBOW_H]],
};

function elbow(s, x, y, corner) {
  const pts = ELBOW_PATHS[corner];
  s.addShape('custGeom', {
    x, y, w: ELBOW_W, h: ELBOW_H,
    points: pts.map(([px, py]) => ({ x: px, y: py })),
    line: { color: '000000', width: 2 }, fill: { type: 'none' },
  });
  [pts[0], pts[2]].forEach(([px, py]) => s.addShape('ellipse', {
    x: x + px - 0.045, y: y + py - 0.045, w: 0.09, h: 0.09, fill: { color: '000000' },
  }));
}

function slide17(pptx) {
  const s = pptx.addSlide();
  stripes(s, 5.009);
  heading(s, 'Creating a Greener Planet', 1.46, 0.611, 10.413, 1.01, 60, 'center');
  text(s, LOREM_WIDE, { x: 1.223, y: 1.783, w: 10.767, h: 0.686, face: BODY, size: 14, ls: 1.3, align: 'center' });

  gauge(s, { x: 1.726, y: 2.549, outer: [44.9, 224.7], inner: [101.5, 223.7] });
  text(s, '70,2%  ', { x: 3.398, y: 3.12, w: 1.348, h: 0.404, align: 'right' });
  elbow(s, 3.636, 3.474, 'tl');
  text(s, '60,2%  ', { x: 0.849, y: 4.677, w: 1.348, h: 0.404 });
  elbow(s, 0.849, 3.859, 'br');

  gauge(s, { x: 5.492, y: 2.549, outer: [24.4, 224.7], inner: [24.4, 184.3] });
  text(s, '70,2%  ', { x: 5.807, y: 3.309, w: 1.348, h: 0.404, align: 'right' });
  elbow(s, 6.045, 3.663, 'tl');
  text(s, '60,2%  ', { x: 7.483, y: 4.552, w: 1.348, h: 0.404, align: 'right' });
  elbow(s, 7.794, 3.707, 'bl');

  gauge(s, { x: 9.481, y: 2.549, outer: [60.3, 189.1], inner: [59.6, 188.7] });
  text(s, '40,2%  ', { x: 11.136, y: 3.081, w: 1.348, h: 0.404, align: 'right' });
  elbow(s, 11.391, 3.474, 'tl');
  text(s, '40,2%  ', { x: 9.302, y: 3.345, w: 1.348, h: 0.404, align: 'right' });
  elbow(s, 9.674, 3.699, 'tl');

  [
    { tx: 1.74, bx: 1.591, title: 'Innovation One', tw: 2.115 },
    { tx: 5.609, bx: 5.46, title: 'Innovation Two', tw: 2.115 },
    { tx: 9.627, bx: 9.486, title: 'Innovation Three', tw: 2.132 },
  ].forEach((c) => {
    text(s, c.title, { x: c.tx, y: 5.593, w: c.tw, h: 0.404, align: 'center' });
    text(s, 'Lorem ipsum dolor sit amet consectetuer.', {
      x: c.bx, y: 6.077, w: 2.413, h: 0.686, face: BODY, size: 14, ls: 1.3, align: 'center',
    });
  });
  pageNumber(s, 17);
}

/**
 * Slide 18 — 2x2 jigsaw. Each piece is a square body whose edges either grow a
 * round tab (a neck plus a circle straddling the edge, in the piece's colour)
 * or carry a bite — the same shape sunk inwards and painted white.
 * Bodies first, then tabs, then bites, so tabs sit on top of their neighbours.
 */
const PIECE = 1.916; // body edge length
const KNOB_D = 0.54; // tab / bite circle diameter
const KNOB_OUT = 0.35; // tab circle centre, outside the body edge
const KNOB_IN = 0.33; // bite circle centre, inside the body edge
const NECK = 0.373; // width of the strip joining a circle to the body

/** Neck + circle on one edge of a piece; `dist` is signed (+ out, - in). */
function knob(s, p, side, dist, color) {
  const paint = { fill: { color }, line: { type: 'none' } };
  const half = PIECE / 2;
  const centre = {
    top: [p.x + half, p.y - dist],
    bottom: [p.x + half, p.y + PIECE + dist],
    left: [p.x - dist, p.y + half],
    right: [p.x + PIECE + dist, p.y + half],
  }[side];
  const edge = { top: p.y, bottom: p.y + PIECE, left: p.x, right: p.x + PIECE }[side];
  const horizontal = side === 'left' || side === 'right';
  s.addShape('rect', Object.assign({
    x: horizontal ? Math.min(centre[0], edge) : centre[0] - NECK / 2,
    y: horizontal ? centre[1] - NECK / 2 : Math.min(centre[1], edge),
    w: horizontal ? Math.abs(centre[0] - edge) : NECK,
    h: horizontal ? NECK : Math.abs(centre[1] - edge),
  }, paint));
  s.addShape('ellipse', Object.assign({
    x: centre[0] - KNOB_D / 2, y: centre[1] - KNOB_D / 2, w: KNOB_D, h: KNOB_D,
  }, paint));
}

const PUZZLE = [
  { x: 7.909, y: 1.833, color: GREEN, tabs: ['top', 'right'], bites: ['left'] },
  { x: 9.815, y: 1.833, color: MINT, tabs: ['right', 'bottom'], bites: ['top'] },
  { x: 7.909, y: 3.749, color: '000000', tabs: ['left', 'top'], bites: ['bottom'] },
  { x: 9.815, y: 3.749, color: NEAR_BLACK, tabs: ['left', 'bottom'], bites: ['right'] },
];

function puzzle(s) {
  PUZZLE.forEach((p) => s.addShape('rect', {
    x: p.x, y: p.y, w: PIECE, h: PIECE, fill: { color: p.color }, line: { type: 'none' },
  }));
  PUZZLE.forEach((p) => p.tabs.forEach((side) => knob(s, p, side, KNOB_OUT, p.color)));
  PUZZLE.forEach((p) => p.bites.forEach((side) => knob(s, p, side, -KNOB_IN, WHITE)));
}

function slide18(pptx) {
  const s = pptx.addSlide();
  stripes(s, 0.76);
  heading(s, 'Ensuring Equity and Accessibility', 0.886, 1.318, 6.119, 1.737, 54);

  const bars = [
    { labelY: 3.366, barY: 3.761, valueY: 3.349, label: 'Innovation One', value: '75%', w: 3.11, color: GREEN },
    { labelY: 4.097, barY: 4.526, valueY: 4.107, label: 'Innovation Two', value: '50%', w: 2.362, color: MINT },
    { labelY: 4.862, barY: 5.291, valueY: 4.872, label: 'Innovation Three', value: '90%', w: 3.976, color: NEAR_BLACK },
    { labelY: 5.695, barY: 6.056, valueY: 5.637, label: 'Innovation Four', value: '80%', w: 3.504, color: '000000' },
  ];
  bars.forEach((b) => {
    s.addText(b.label, {
      x: 0.966, y: b.labelY, w: 1.983, h: 0.37,
      fontFace: HEAD, fontSize: 16, color: INK, align: 'center', valign: 'bottom',
    });
    s.addText(b.value, {
      x: 4.926, y: b.valueY, w: 0.766, h: 0.37,
      fontFace: HEAD, fontSize: 16, color: INK, align: 'right', valign: 'bottom',
    });
    s.addShape('rect', { x: 1.07, y: b.barY, w: 4.65, h: 0.126, fill: { color: TRACK_GREY, transparency: 75 } });
    s.addShape('rect', { x: 1.07, y: b.barY, w: b.w, h: 0.126, fill: { color: b.color } });
  });

  puzzle(s);
  [
    { x: 9.124, y: 1.832, num: '01', color: MINT },
    { x: 11.042, y: 1.832, num: '02', color: GREEN },
    { x: 9.124, y: 3.75, num: '03', color: WHITE },
    { x: 11.041, y: 3.749, num: '04', color: WHITE },
  ].forEach((n) => {
    s.addText(n.num, {
      x: n.x, y: n.y, w: 0.702, h: 0.572,
      fontFace: BODY, fontSize: 28, color: n.color, align: 'right', valign: 'top',
    });
  });
  pageNumber(s, 18);
}

function slide19(pptx) {
  const s = pptx.addSlide();
  stripes(s, 5.009);
  heading(s, 'Energy Data Table', 1.46, 0.611, 10.413, 1.01, 60, 'center');
  s.addShape('roundRect', {
    x: 1.106, y: 2.046, w: 11.121, h: 3.709, rectRadius: 0.1876,
    fill: { color: WHITE }, line: { type: 'none' },
    shadow: { type: 'outer', color: '000000', opacity: 0.1, blur: 35, offset: 3, angle: 45 },
  });

  const COLX = [3.127, 4.943, 6.758, 8.574, 10.389]; // data column left edges
  const CELLW = 1.704;
  const rowFills = [GREEN, MINT, NEAR_BLACK, DARK_GREY, MID_GREY];
  const rowText = [MINT, GREEN, WHITE, WHITE, WHITE];
  const headText = [WHITE, GREEN, WHITE, WHITE, WHITE];
  // which data columns are ticked, per row
  const ticks = [[0, 1, 3, 4], [1, 2, 4], [0, 3], [0, 2, 4], [1, 3]];

  // header strip
  s.addShape('roundRect', { x: 1.273, y: 2.205, w: 1.759, h: 0.493, fill: { color: '808080' }, line: { type: 'none' }, rectRadius: 0.0822 });
  text(s, 'Energy', { x: 1.273, y: 2.264, w: 1.759, h: 0.404, color: WHITE, align: 'center' });
  COLX.forEach((x, i) => {
    s.addShape('roundRect', { x, y: 2.205, w: CELLW, h: 0.493, fill: { color: rowFills[i] }, line: { type: 'none' }, rectRadius: 0.0822 });
    text(s, 'Data 0' + (i + 1), { x, y: 2.264, w: CELLW, h: 0.404, color: headText[i], align: 'center' });
  });

  // body rows
  ticks.forEach((cols, r) => {
    const y = 2.789 + r * 0.584;
    s.addShape('roundRect', { x: 3.127, y, w: 8.966, h: 0.493, fill: { color: TRACK_LIGHT }, line: { type: 'none' }, rectRadius: 0.0822 });
    s.addShape('roundRect', { x: 1.273, y, w: 1.759, h: 0.493, fill: { color: rowFills[r] }, line: { type: 'none' }, rectRadius: 0.0822 });
    text(s, 'Energy 0' + (r + 1), { x: 1.273, y: y + 0.056, w: 1.759, h: 0.404, color: rowText[r], align: 'center' });
    cols.forEach((c) => {
      const cx = COLX[c] + CELLW / 2;
      s.addShape('ellipse', { x: cx - 0.157, y: y + 0.095, w: 0.314, h: 0.302, fill: { color: rowFills[r] } });
      s.addShape('custGeom', {
        x: cx - 0.075, y: y + 0.19, w: 0.15, h: 0.11,
        points: [{ x: 0, y: 0.06 }, { x: 0.055, y: 0.11 }, { x: 0.15, y: 0 }],
        line: { color: r === 1 ? GREEN : WHITE, width: 1.75 },
      });
    });
  });

  text(s, LOREM_WIDE, { x: 1.283, y: 6.192, w: 10.767, h: 0.686, face: BODY, size: 14, ls: 1.3, align: 'center' });
  pageNumber(s, 19);
}

function slide20(pptx) {
  const s = pptx.addSlide();
  stripes(s, 9.268);
  heading(s, 'Thank You For Attention', 0.76, 0.486, 8.008, 2.767, 88);
  imagePlaceholder(s, 6.667, 3.362, 3.1, 3.098, { radius: 1.549 });
  imagePlaceholder(s, 9.473, 1.178, 3.1, 3.098, { radius: 1.549 });
  card(s, 0.76, 4.517, 6.323, 2.388, GREEN, 0.288);
  const cols = [
    { x: 1.252, tx: 1.697, title: 'Phone:', tw: 0.982, lines: ['+123 837 4521', '+123 923 5148'], lw: 1.721, icon: 'phone' },
    { x: 3.748, tx: 4.194, title: 'Website & Email:', tw: 2.112, lines: ['www.energy.com', 'energy@email.com'], lw: 2.398, icon: 'mail' },
  ];
  cols.forEach((c) => {
    s.addShape('ellipse', { x: c.x, y: 5.251, w: 0.373, h: 0.373, fill: { color: MINT } });
    if (c.icon === 'phone') {
      s.addShape('custGeom', {
        x: c.x + 0.105, y: 5.356, w: 0.163, h: 0.163,
        points: [{ x: 0.02, y: 0 }, { x: 0.07, y: 0.05 }, { x: 0.04, y: 0.08 },
          { x: 0.09, y: 0.13 }, { x: 0.12, y: 0.1 }, { x: 0.163, y: 0.15 },
          { x: 0.12, y: 0.163 }, { x: 0.02, y: 0.06 }, { close: true }],
        fill: { color: GREEN }, line: { type: 'none' },
      });
    } else {
      s.addShape('rect', { x: c.x + 0.098, y: 5.375, w: 0.177, h: 0.125, fill: { color: GREEN } });
      s.addShape('line', { x: c.x + 0.118, y: 5.42, w: 0.137, h: 0, line: { color: MINT, width: 1 } });
    }
    text(s, c.title, { x: c.tx, y: 5.236, w: c.tw, h: 0.404, color: MINT });
    c.lines.forEach((line, i) => {
      text(s, line, { x: c.tx, y: 5.646 + i * 0.411, w: c.lw, h: 0.404, color: WHITE });
    });
  });
  card(s, 0.752, 4.269, 2.872, 0.479, MINT, 0.2395);
  text(s, 'Get in Touch With Us', { x: 0.912, y: 4.302, w: 2.552, h: 0.412, color: GREEN, align: 'center' });
  pageNumber(s, 20);
}

/* ------------------------------------------------------------------ main */
function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: W, height: H });
  pptx.layout = 'WIDE';
  pptx.title = 'Energy — Presentation Template';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach((fn) => fn(pptx));

  return pptx.writeFile({
    fileName: path.join(__dirname, '0c62c766-2918-4aee-bac8-703c839b82f7_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote ' + f)).catch((e) => { console.error(e); process.exit(1); });
