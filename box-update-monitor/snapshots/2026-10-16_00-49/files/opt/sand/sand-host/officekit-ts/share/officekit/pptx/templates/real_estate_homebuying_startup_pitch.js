/**
 * "Real Estate" / Luxure homebuying-startup deck - rebuilt with pptxgenjs.
 *
 * 27 slides, 13.333 x 7.5 in.  Theme "Luxure": Archivo (headings) + Manrope (body).
 * The source deck contains no bitmaps on any slide (its only media file is an
 * unused layout texture), so nothing here needs an image; every mark, icon and
 * map is drawn with native pptxgenjs shapes.
 *
 * Run:  node 0e9090b5-74b8-4c7c-b770-721d32eda74c_grok_final.js
 */
'use strict';

const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ theme */

const C = {
  ink: '3C3C3C',      // tx1
  white: 'FFFFFF',
  orange: 'F95114',   // accent1
  amber: 'F9B25D',    // accent2
  teal: '518984',     // accent3
  coral: 'FA7D50',    // accent4
  sand: 'FBC98D',     // accent5
  peach: 'FB9772',    // accent1 @ 60% luminance - chart bars
  grey: 'B2B2B2',
  lightGrey: 'DADADA',
  paleGrey: 'EBEBEB',
  gridline: 'FDDCD0', // accent1 @ 20% luminance
  axis: 'E2E2E2',
};
const HEAD = 'Archivo';   // major latin
const BODY = 'Manrope';   // minor latin

const ACCENTS = [C.orange, C.amber, C.teal];
const SUBTITLE = 'Homebuying Startup Presentation Template';

/* ---------------------------------------------------------------- helpers */

const noLine = { type: 'none' };

/** Filled / outlined rectangle (fill = null for outline-only). */
function rect(s, x, y, w, h, fill, opt = {}) {
  s.addShape('rect', Object.assign({
    x, y, w, h,
    fill: fill ? { color: fill } : { type: 'none' },
    line: noLine,
  }, opt));
}

/** Straight connector. */
function line(s, x, y, w, h, color, width = 1, opt = {}) {
  s.addShape('line', Object.assign({ x, y, w, h, line: { color, width } }, opt));
}

function oval(s, x, y, w, h, fill, opt = {}) {
  s.addShape('ellipse', Object.assign({ x, y, w, h, fill: { color: fill }, line: noLine }, opt));
}

/** Body copy. */
function tx(s, text, o) {
  s.addText(text, Object.assign({ fontFace: BODY, fontSize: 12, color: C.ink }, o));
}

/** Heading copy (Archivo, bold by default). */
function hx(s, text, o) {
  s.addText(text, Object.assign({ fontFace: HEAD, fontSize: 18, bold: true, color: C.ink }, o));
}

/** Turn a list of strings into paragraphs for a single text frame. */
function paras(list, opt = {}) {
  return list.map((t, i) => ({ text: t, options: Object.assign({ breakLine: i < list.length - 1 }, opt) }));
}

const BULLET = { characterCode: '2022', indent: 13.5 };

/* --------------------------------------------------------- brand elements */

/**
 * The theme's "wide downward diagonal" pattern fill, redrawn as thin 45 deg
 * rules over a solid ground (pptxgenjs has no pattern fills).
 */
function hatch(s, x, y, w, h, ground = C.white) {
  rect(s, x, y, w, h, ground);
  for (let c = -h; c < w; c += 0.0525) {
    const ax = Math.max(c, 0), ay = Math.max(-c, 0);
    const bx = Math.min(w, c + h), by = Math.min(h, w - c);
    if (bx - ax < 0.01) continue;
    line(s, x + ax, y + ay, bx - ax, by - ay, C.amber, 0.25);
  }
}

/** Luxure house mark: a chevron roof over a small square. */
function houseMark(s, x, y, size, color = C.orange) {
  const roof = [[0.5, 0], [1, 0.47], [1, 0.93], [0.5, 0.46], [0, 0.93], [0, 0.47]];
  s.addShape('custGeom', {
    x, y, w: size, h: size, fill: { color }, line: noLine,
    points: roof.map(([px, py]) => ({ x: px * size, y: py * size })).concat([{ close: true }]),
  });
  rect(s, x + 0.393 * size, y + 0.786 * size, 0.214 * size, 0.207 * size, color);
}

/** The three stacked accent chips (coral / amber / teal). */
function chips(s, x, y, w = 0.205, h = 0.223, horizontal = false) {
  [C.coral, C.amber, C.teal].forEach((col, i) => {
    if (horizontal) rect(s, x + i * w, y, w, h, col);
    else rect(s, x, y + i * h, w, h, col);
  });
}

/**
 * Furniture inherited from the slide master: accent chips on the left edge,
 * the house mark + wordmark top right and the hatched corner square.
 */
function deco(s, o = {}) {
  const opt = Object.assign({ chips: true, mark: true, word: false, corner: true, ink: C.ink }, o);
  if (opt.chips) chips(s, 0, 0.524);
  if (opt.mark) houseMark(s, 12.435, 0.472, 0.247);
  if (opt.word) hx(s, 'Luxure', { x: 11.653, y: 0.444, w: 0.801, h: 0.303, fontSize: 12, color: opt.ink, valign: 'middle' });
  if (opt.corner) hatch(s, 12.683, 6.848, 0.651, 0.652, opt.cornerGround || C.white);
}

/** Standard slide header: kicker line + big title. */
function header(s, title, o = {}) {
  const opt = Object.assign({ size: 40, h: 0.776, color: C.ink, subColor: C.ink, sub: SUBTITLE }, o);
  if (opt.sub) tx(s, opt.sub, { x: 0.651, y: 0.445, w: 12.031, h: 0.303, color: opt.subColor, valign: 'middle', lineSpacingMultiple: 1 });
  s.addText(title, {
    x: 0.651, y: 0.675, w: opt.w || 12.031, h: opt.h,
    fontFace: HEAD, fontSize: opt.size, bold: true, color: opt.color,
    valign: 'middle', lineSpacingMultiple: 1,
  });
}

/* ------------------------------------------------------------------ icons */

/** Small flat glyphs standing in for the deck's detailed vector icons. */
function icon(s, kind, x, y, w, color) {
  const h = w;
  switch (kind) {
    case 'home':
      s.addShape('triangle', { x, y, w, h: 0.5 * h, fill: { color }, line: noLine });
      rect(s, x + 0.16 * w, y + 0.45 * h, 0.68 * w, 0.55 * h, color);
      rect(s, x + 0.3 * w, y + 0.6 * h, 0.15 * w, 0.15 * h, C.white);
      rect(s, x + 0.55 * w, y + 0.6 * h, 0.15 * w, 0.15 * h, C.white);
      break;
    case 'people':
      [0.0, 0.36, 0.72].forEach((dx) => oval(s, x + dx * w, y, 0.28 * w, 0.28 * h, color));
      [0.0, 0.36, 0.72].forEach((dx) => rect(s, x + dx * w, y + 0.38 * h, 0.28 * w, 0.34 * h, color));
      rect(s, x + 0.1 * w, y + 0.86 * h, 0.8 * w, 0.1 * h, color);
      break;
    case 'chart':
      rect(s, x, y + 0.12 * h, w, 0.76 * h, color);
      rect(s, x + 0.16 * w, y + 0.5 * h, 0.14 * w, 0.24 * h, C.white);
      rect(s, x + 0.4 * w, y + 0.38 * h, 0.14 * w, 0.36 * h, C.white);
      rect(s, x + 0.64 * w, y + 0.26 * h, 0.14 * w, 0.48 * h, C.white);
      break;
    case 'board':
      rect(s, x, y, w, 0.62 * h, color);
      rect(s, x + 0.14 * w, y + 0.14 * h, 0.72 * w, 0.34 * h, C.white);
      rect(s, x + 0.46 * w, y + 0.62 * h, 0.08 * w, 0.38 * h, color);
      rect(s, x + 0.2 * w, y + 0.92 * h, 0.6 * w, 0.08 * h, color);
      break;
    case 'city':
      rect(s, x, y + 0.24 * h, 0.44 * w, 0.76 * h, color);
      rect(s, x + 0.52 * w, y, 0.48 * w, h, color);
      rect(s, x + 0.62 * w, y + 0.16 * h, 0.1 * w, 0.12 * h, C.white);
      rect(s, x + 0.8 * w, y + 0.16 * h, 0.1 * w, 0.12 * h, C.white);
      rect(s, x + 0.62 * w, y + 0.44 * h, 0.1 * w, 0.12 * h, C.white);
      rect(s, x + 0.8 * w, y + 0.44 * h, 0.1 * w, 0.12 * h, C.white);
      break;
    case 'pin':
      oval(s, x, y, w, 0.8 * h, color);
      s.addShape('triangle', { x: x + 0.25 * w, y: y + 0.5 * h, w: 0.5 * w, h: 0.5 * h, fill: { color }, line: noLine, rotate: 180 });
      oval(s, x + 0.33 * w, y + 0.26 * h, 0.34 * w, 0.28 * h, C.white);
      break;
    case 'mail':
      rect(s, x, y + 0.12 * h, w, 0.76 * h, color);
      s.addShape('triangle', { x: x + 0.1 * w, y: y + 0.2 * h, w: 0.8 * w, h: 0.4 * h, fill: { color: C.white }, line: noLine, rotate: 180 });
      break;
    case 'globe':
      oval(s, x, y, w, h, color);
      oval(s, x + 0.33 * w, y, 0.34 * w, h, C.white);
      rect(s, x + 0.04 * w, y + 0.44 * h, 0.92 * w, 0.1 * h, C.white);
      break;
    case 'phone':
      oval(s, x, y + 0.2 * h, w, 0.7 * h, color);
      rect(s, x + 0.28 * w, y, 0.44 * w, 0.3 * h, color);
      oval(s, x + 0.36 * w, y + 0.42 * h, 0.28 * w, 0.26 * h, C.white);
      break;
    default:
      oval(s, x, y, w, h, color);
  }
}

/** Accent disc with a white glyph inside. */
function iconDisc(s, kind, cx, cy, d, color, glyph = 0.62) {
  oval(s, cx - d / 2, cy - d / 2, d, d, color);
  icon(s, kind, cx - (d * glyph) / 2, cy - (d * glyph) / 2, d * glyph, C.white);
}

/** Small disc with a white arrow - the deck's "next step" marker. */
function arrowDot(s, x, y, d, color) {
  oval(s, x, y, d, d, color);
  s.addShape('rightArrow', {
    x: x + 0.2 * d, y: y + 0.3 * d, w: 0.6 * d, h: 0.4 * d,
    fill: { color: C.white }, line: noLine,
  });
}

/** Small disc with a white check - used on the comparison table. */
function checkDot(s, x, y, d, color) {
  oval(s, x, y, d, d, color);
  s.addShape('custGeom', {
    x: x + 0.22 * d, y: y + 0.3 * d, w: 0.56 * d, h: 0.4 * d,
    fill: { type: 'none' }, line: { color: C.white, width: 1.5 },
    points: [{ x: 0, y: 0.5 * 0.4 * d }, { x: 0.35 * 0.56 * d, y: 0.4 * d }, { x: 0.56 * d, y: 0 }],
  });
}

/** Legend swatch: coloured dot + caption. */
function legendDot(s, x, y, color, text) {
  oval(s, x, y, 0.119, 0.119, color);
  tx(s, text, { x: x + 0.113, y: y - 0.075, w: 0.686, h: 0.269, fontSize: 10, valign: 'middle' });
}

/* -------------------------------------------------------------- the build */

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'W16x9', width: 13.333, height: 7.5 });
pptx.layout = 'W16x9';
pptx.author = 'Luxure';
pptx.title = 'Real Estate';

/* ------------------------------------------------------- 1. title slide */

function slide01() {
  const s = pptx.addSlide();
  hatch(s, 0, 6.329, 13.333, 1.171);
  deco(s);
  tx(s, 'Homebuying Startup', { x: 0.907, y: 0.74, w: 2.975, h: 0.337, fontSize: 14, valign: 'middle' });
  line(s, 3.222, 0.908, 7.205, 0, C.amber, 0.75);
  tx(s, '13 October, 2027', { x: 9.452, y: 0.74, w: 2.975, h: 0.337, fontSize: 14, align: 'right', valign: 'middle' });
  s.addText('Real Estate', {
    x: 0.804, y: 1.366, w: 7.551, h: 1.449,
    fontFace: HEAD, fontSize: 80, bold: true, color: C.ink, valign: 'top', lineSpacingMultiple: 0.9,
  });
  hx(s, 'Betty W. Brunner', { x: 9.604, y: 1.769, w: 2.499, h: 0.404, color: C.orange });
  tx(s, 'Presentation Speaker', { x: 9.604, y: 2.112, w: 2.499, h: 0.303 });
  chips(s, 12.222, 1.757);
}

/* ------------------------------------------------------------ 2. agenda */

const AGENDA = [
  'Sed vulputate mi sit non sed ameit',
  'Donec felis ultricies pellentesq neu sociis natoque penatibus',
  'PLACEHOLDER',
  'PLACEHOLDER',
  'Donec felis ultricies pellentesq neu sociis natoque penatibus',
  'Donec felis ultricies pellentesq neu sociis natoque penatibus',
  'Donec felis ultricies pellentesq neu sociis natoque penatibus',
  'Donec felis ultricies pellentesq neu sociis natoque penatibus',
];

function slide02() {
  const s = pptx.addSlide();
  deco(s, { word: true });
  header(s, 'Agenda for Today');
  const cols = [0.651, 3.695, 6.738, 9.782];
  const rows = [1.962, 4.553];
  AGENDA.forEach((body, i) => {
    const x = cols[i % 4], y = rows[Math.floor(i / 4)], col = ACCENTS[i % 3];
    s.addText([
      { text: String(i + 1).padStart(2, '0'), options: { fontSize: 48 } },
      { text: '/08', options: { fontSize: 28 } },
    ], { x, y, w: 2.9, h: 0.909, fontFace: HEAD, bold: true, color: col, valign: 'middle' });
    hx(s, 'Description', { x, y: y + 0.789, w: 2.9, h: 0.438, fontSize: 20, color: col, valign: 'middle' });
    tx(s, body, { x, y: y + 1.108, w: 2.9, h: 1.05, lineSpacingMultiple: 1.5, valign: 'top' });
  });
}

/* ------------------------------------------------------- 3. mission hero */

function slide03() {
  const s = pptx.addSlide();
  deco(s);
  chips(s, 0, 4.419);
  tx(s, SUBTITLE, { x: 0.651, y: 4.379, w: 12.031, h: 0.303, valign: 'middle' });
  s.addText('Luxure is Building the World\u2019s Best Homebuying Experience', {
    x: 0.651, y: 4.605, w: 9.045, h: 1.583,
    fontFace: HEAD, fontSize: 44, bold: true, color: C.ink, valign: 'middle', lineSpacingMultiple: 1,
  });
  s.addShape('roundRect', {
    x: 10.051, y: 5.057, w: 2.9, h: 1.181, rectRadius: 0.59,
    fill: { color: C.orange }, line: noLine,
  });
  hx(s, 'Luxure', { x: 10.548, y: 5.396, w: 1.42, h: 0.505, fontSize: 24, color: C.white, valign: 'middle' });
  houseMark(s, 11.968, 5.405, 0.485, C.white);
  hx(s, 'Description', { x: 0.652, y: 6.196, w: 2.9, h: 0.438, fontSize: 20, color: C.orange, valign: 'middle' });
  tx(s, 'PLACEHOLDER'
    + 'PLACEHOLDER',
  { x: 2.414, y: 6.196, w: 7.186, h: 0.675, lineSpacingMultiple: 1.5, valign: 'top' });
}

/* --------------------------------------------------------- 4. about page */

function slide04() {
  const s = pptx.addSlide();
  deco(s, { word: true });
  header(s, 'About Homebuying Startup');
  hx(s, 'Detail Information', { x: 4.752, y: 2.142, w: 3.733, h: 0.438, fontSize: 20, color: C.orange });
  const colA = [
    'PLACEHOLDER'
    + 'PLACEHOLDER',
    '',
    'PLACEHOLDER'
    + 'PLACEHOLDER',
  ];
  const colB = [
    'PLACEHOLDER'
    + 'PLACEHOLDER',
    '',
    'PLACEHOLDER',
  ];
  const bodyOpt = { align: 'justify', lineSpacingMultiple: 1.5 };
  tx(s, paras(colA, bodyOpt), { x: 4.752, y: 2.483, w: 4.1, h: 3.099, valign: 'top' });
  tx(s, paras(colB, bodyOpt), { x: 8.852, y: 2.483, w: 3.733, h: 2.493, valign: 'top' });

  const places = [
    { x: 4.897, tx: 5.363, w: 2.512, col: C.orange, head: 'First Location', body: 'Description, London, England' },
    { x: 8.219, tx: 8.684, w: 3.024, col: C.amber, head: 'Second Location', body: 'Detail Information, London, England' },
  ];
  places.forEach((p) => {
    icon(s, 'pin', p.x, 5.789, 0.324, p.col);
    hx(s, p.head, { x: p.tx, y: 5.674, w: p.w, h: 0.438, fontSize: 20, color: p.col });
    tx(s, p.body, { x: p.tx, y: 5.986, w: p.w, h: 0.373, lineSpacingMultiple: 1.5, valign: 'top' });
  });
}

/* -------------------------------------------------- 5. how homebuying works */

function slide05() {
  const s = pptx.addSlide();
  deco(s, { word: true });
  header(s, 'How Homebuying Works');
  const steps = [
    { x: 0.822, col: C.orange, ico: 'home' },
    { x: 3.892, col: C.amber, ico: 'people' },
    { x: 6.962, col: C.teal, ico: 'chart' },
    { x: 10.032, col: C.orange, ico: 'board' },
  ];
  steps.forEach((st, i) => {
    s.addShape('ellipse', {
      x: st.x, y: 2.21, w: 2.479, h: 2.479,
      fill: { color: st.col, transparency: 70 }, line: noLine,
    });
    iconDisc(s, st.ico, st.x + 1.2395, 3.45, 1.132, st.col, 0.34);
    hx(s, 'Description', { x: st.x + 0.003, y: 4.967, w: 2.476, h: 0.438, fontSize: 20, color: st.col, align: 'center', valign: 'middle' });
    tx(s, 'PLACEHOLDER',
      { x: st.x + 0.003, y: 5.286, w: 2.476, h: 0.978, align: 'center', lineSpacingMultiple: 1.5, valign: 'top' });
    if (i < 3) arrowDot(s, st.x + 2.647, 3.322, 0.256, ACCENTS[i]);
  });
}

/* ------------------------------------------------------ 6. three columns */

function slide06() {
  const s = pptx.addSlide();
  deco(s, { word: true });
  header(s, 'We Have Created a Unique Homebuying Experience That People Love', { h: 1.449 });
  const cards = [
    {
      x: 0.665, w: 3.668, col: C.orange, ico: 'home', head: 'Structural Inefficiencies', tx: 0.945,
      bullets: ['PLACEHOLDER', 'Velit esseu quam nihil consequatu'],
    },
    {
      x: 4.667, w: 4.0, col: C.amber, ico: 'people', head: 'Seamless Experience', tx: 5.112,
      bullets: ['PLACEHOLDER', 'Sequi nesciu neque porro quisqiu'],
    },
    {
      x: 9.0, w: 3.668, col: C.teal, ico: 'chart', head: 'Tech Driven', tx: 9.28,
      bullets: ['PLACEHOLDER', 'Tempora incidunt ut labore dolore'],
    },
  ];
  cards.forEach((c) => {
    rect(s, c.x, 2.438, c.w, 0.669, c.col);
    rect(s, c.x, 2.438, c.w, 4.312, null, { line: { color: c.col, width: 1 } });
    hx(s, c.head, { x: c.x, y: 2.587, w: c.w, h: 0.37, fontSize: 16, color: C.white, align: 'center', valign: 'top' });
    line(s, c.x, 4.978, c.w, 0, c.col, 1);
    iconDisc(s, c.ico, c.tx + 0.51, 4.043, 1.021, c.col, 0.34);
    hx(s, 'Description', { x: c.tx + 1.116, y: 3.554, w: 1.992, h: 0.404, color: c.col, valign: 'middle' });
    tx(s, 'Donu ultrices pellentes sed natoque magnis',
      { x: c.tx + 1.116, y: 3.855, w: 1.992, h: 0.675, lineSpacingMultiple: 1.5, valign: 'top' });
    hx(s, 'Description', { x: c.tx, y: 5.215, w: 3.109, h: 0.438, fontSize: 20, color: c.col });
    tx(s, paras(c.bullets, { bullet: BULLET, lineSpacingMultiple: 1.5 }),
      { x: c.tx, y: 5.534, w: 3.109, h: 1.1, valign: 'top' });
  });
}

/* --------------------------------------------------- 7. homebuyers come last */

function slide07() {
  const s = pptx.addSlide();
  deco(s);
  header(s, 'Homebuyers Have Always Come Last In Real Estate', { w: 7.35, h: 1.449 });
  iconDisc(s, 'home', 9.6715, 1.3995, 1.021, C.orange, 0.34);
  [
    { x: 0.665, tx: 0.916, col: C.orange },
    { x: 3.667, tx: 3.918, col: C.amber },
    { x: 6.667, tx: 6.918, col: C.teal },
  ].forEach((c) => {
    rect(s, c.x, 5.417, 2.667, 1.333, c.col);
    hx(s, 'Description', { x: c.tx, y: 5.591, w: 2.165, h: 0.404, color: C.white, valign: 'middle' });
    tx(s, 'Dolores eum enimu ipsam voluptat quia aspernie',
      { x: c.tx, y: 5.893, w: 2.165, h: 0.675, color: C.white, lineSpacingMultiple: 1.5, valign: 'top' });
  });
}

/* ------------------------------------------------- 8. why we are different */

function slide08() {
  const s = pptx.addSlide();
  deco(s);
  header(s, 'Why We Are Different');
  const rows = [
    { y: 1.922, col: C.orange, ico: 'home', big: '+6.000' },
    { y: 3.276, col: C.amber, ico: 'people', big: '$45m+' },
    { y: 4.63, col: C.teal, ico: 'chart', big: '7531m\u00B2' },
  ];
  rows.forEach((r) => {
    hx(s, r.big, { x: 0.651, y: r.y + 0.123, w: 2.299, h: 0.774, fontSize: 40, color: r.col, valign: 'middle' });
    iconDisc(s, r.ico, 3.4905, r.y + 0.5105, 1.021, r.col, 0.34);
    hx(s, 'Description', { x: 4.097, y: r.y + 0.022, w: 2.903, h: 0.404, color: r.col, valign: 'middle' });
    tx(s, 'PLACEHOLDER',
      { x: 4.097, y: r.y + 0.323, w: 2.903, h: 0.675, lineSpacingMultiple: 1.5, valign: 'top' });
  });
  icon(s, 'pin', 0.822, 6.256, 0.324, C.orange);
  hx(s, 'Real Estate Location', { x: 1.288, y: 6.141, w: 4.512, h: 0.438, fontSize: 20, color: C.orange });
  tx(s, 'Ultrices pellentes sed natoque sed non magnis, London, England',
    { x: 1.288, y: 6.453, w: 5.712, h: 0.373, lineSpacingMultiple: 1.5, valign: 'top' });
  hatch(s, 7.0, 6.848, 0.651, 0.652);
}

/* --------------------------------------------------- 9. market snapshot */

function slide09() {
  const s = pptx.addSlide();
  deco(s, { word: true });
  header(s, 'Real Estate \u2013 Market Snapshot');
  const panels = [
    { y: 1.748, col: C.orange, tag: 'January 2028', pct: '13%', big1: '5.873', big2: '$241k' },
    { y: 4.417, col: C.amber, tag: 'January 2032', pct: '17%', big1: '7,348', big2: '$525k' },
  ];
  panels.forEach((p) => {
    rect(s, -0.167, p.y + 0.832, 2.332, 0.668, p.col, { rotate: 270 });
    hx(s, p.tag, { x: -0.167, y: p.y + 0.981, w: 2.332, h: 0.37, fontSize: 16, color: C.white, align: 'center', rotate: 270 });
    rect(s, 0.665, p.y, 12.003, 2.333, null, { line: { color: p.col, width: 1 } });
    s.addShape('upArrow', {
      x: 6.34, y: p.y + 0.389, w: 1.484, h: 1.556,
      fill: { color: p.col }, line: noLine,
    });
    icon(s, 'home', 6.941, p.y + 0.969, 0.281, C.white);
    hx(s, p.pct, { x: 6.34, y: p.y + 1.271, w: 1.484, h: 0.438, fontSize: 20, color: C.white, align: 'center', valign: 'middle' });
    [
      { dy: 0.329, big: p.big1, note: 'Aliqu commodi consue quis' },
      { dy: 1.164, big: p.big2, note: 'Ultrices pellent sed natoque' },
    ].forEach((r) => {
      arrowDot(s, 8.2, p.y + r.dy + 0.263, 0.315, p.col);
      hx(s, r.big, { x: 8.635, y: p.y + r.dy, w: 2.299, h: 0.841, fontSize: 44, color: p.col, valign: 'middle' });
      tx(s, r.note, { x: 10.626, y: p.y + r.dy + 0.083, w: 1.369, h: 0.675, lineSpacingMultiple: 1.5, valign: 'top' });
    });
  });
}

/* ----------------------------------------------------- 10. average price */

function slide10() {
  const s = pptx.addSlide();
  deco(s, { word: true });
  header(s, 'Real Estate \u2013 Average Price (in dollars)');
  const cards = [
    { x: 1.012, col: C.orange, ico: 'home', pct: '+24%', bx: 0.76, big: '$321.986' },
    { x: 4.014, col: C.amber, ico: 'city', pct: '+36%', bx: 3.762, big: '$224.763' },
    { x: 7.014, col: C.teal, ico: 'home', pct: '+57%', bx: 6.764, big: '$547.872' },
    { x: 10.014, col: C.orange, ico: 'city', pct: '+82%', bx: 9.765, big: '$453.879' },
  ];
  cards.forEach((c) => {
    rect(s, c.x, 1.755, 2.319, 1.333, c.col, { line: { color: C.white, width: 1.5 } });
    icon(s, c.ico, c.x + 1.022, 1.969, 0.276, C.white);
    hx(s, c.pct, { x: c.x + 0.302, y: 2.23, w: 1.716, h: 0.64, fontSize: 32, color: C.white, align: 'center', valign: 'middle' });
    hx(s, c.big, { x: c.bx, y: 5.604, w: 2.823, h: 0.774, fontSize: 40, color: c.col, align: 'center', valign: 'middle' });
    hx(s, 'Detail Information', { x: c.bx, y: 6.276, w: 2.823, h: 0.404, color: c.col, align: 'center', valign: 'middle' });
    tx(s, 'Aliqu commodi consesu eum iure',
      { x: c.bx, y: 6.544, w: 2.823, h: 0.373, align: 'center', lineSpacingMultiple: 1.5, valign: 'top' });
  });
}

/* --------------------------------------------------- 11. comparison table */

const COMPARE_ROWS = [
  ['01. Nis aliqu commodi consequatur quis autem reprehenderi', 3],
  ['02. Qui in voluptate velit esseu quam molestiae consequatur', 3],
  ['03. Illum dolores eum enim ipsam non voluptatem sit amet', 3],
  ['04. Sequi nesciu neque porro quisqiu esti dolorem sed quia', 2],
  ['05. Numqua eius modi tempora incident  consecteti velit', 1],
  ['06. Illum dolores eum enim ipsam non voluptatem sit amet', 1],
];

function slide11() {
  const s = pptx.addSlide();
  deco(s, { word: true });
  header(s, 'We Have a Stronger Business Model');
  const border = [{ type: 'solid', color: C.sand, pt: 1 }];
  const head = (t, fill) => ({
    text: t,
    options: { fill: { color: fill }, color: C.white, fontFace: HEAD, fontSize: 16, align: 'center', valign: 'middle' },
  });
  const rows = [[
    { text: '', options: { border: [{ type: 'none' }, border[0], border[0], { type: 'none' }] } },
    head('Luxure', C.orange), head('Other 1', C.amber), head('Other 2', C.teal),
  ]];
  COMPARE_ROWS.forEach(([label]) => {
    rows.push([
      { text: label, options: { fontSize: 14, valign: 'middle', align: 'left' } }, '', '', '',
    ]);
  });
  s.addTable(rows, {
    x: 0.665, y: 1.75, w: 12.003, colW: [6.502, 1.834, 1.834, 1.834], rowH: 0.714,
    border, fontFace: BODY, color: C.ink, margin: [0.05, 0.1, 0.05, 0.1],
  });
  const marks = [7.896, 9.735, 11.574];
  const cols = [C.orange, C.amber, C.teal];
  const ys = [2.664, 3.37, 4.081, 4.803, 5.525, 6.236];
  COMPARE_ROWS.forEach(([, count], r) => {
    for (let cIdx = 0; cIdx < count; cIdx += 1) checkDot(s, marks[cIdx], ys[r], 0.315, cols[cIdx]);
  });
}

/* ------------------------------------------------- 12. market activity (dark) */

function slide12() {
  const s = pptx.addSlide();
  deco(s);
  rect(s, 0, 0, 13.333, 7.5, '585C5E');
  chips(s, 0, 0.524);
  houseMark(s, 12.435, 0.472, 0.247);
  hx(s, 'Luxure', { x: 11.653, y: 0.444, w: 0.801, h: 0.303, fontSize: 12, color: C.white, valign: 'middle' });
  hatch(s, 12.683, 6.848, 0.651, 0.652, C.ink);
  header(s, 'Real Estate \u2013 Market Activity', { color: C.white, subColor: C.white });
  const cards = [
    { x: 0.665, w: 3.668, tx: 0.904, col: C.orange, ico: 'home', ix: 2.295 },
    { x: 4.667, w: 4.0, tx: 5.071, col: C.amber, ico: 'city', ix: 6.447 },
    { x: 9.0, w: 3.668, tx: 9.239, col: C.teal, ico: 'home', ix: 10.632 },
  ];
  cards.forEach((c) => {
    rect(s, c.x, 1.767, c.w, 1.346, c.col);
    icon(s, c.ico, c.ix, 1.914, 0.408, C.white);
    hx(s, 'Description', { x: c.tx, y: 2.313, w: 3.19, h: 0.404, color: C.white, align: 'center', valign: 'middle' });
    tx(s, 'Aliqu commodi consesu eum sed iure',
      { x: c.tx, y: 2.593, w: 3.19, h: 0.373, color: C.white, align: 'center', lineSpacingMultiple: 1.5, valign: 'top' });
  });
  [
    { x: 2.999, col: C.orange, pct: '+54%' },
    { x: 7.337, col: C.amber, pct: '+37%' },
    { x: 11.331, col: C.teal, pct: '+28%' },
  ].forEach((b) => {
    rect(s, b.x, 5.405, 1.335, 1.346, b.col);
    s.addShape('upArrow', { x: b.x + 0.558, y: 5.723, w: 0.218, h: 0.228, fill: { color: C.white }, line: noLine });
    hx(s, b.pct, { x: b.x, y: 5.969, w: 1.335, h: 0.505, fontSize: 24, color: C.white, align: 'center', valign: 'middle' });
  });
}

/* ------------------------------------------------------ chart defaults */

const CHART_BASE = {
  chartArea: { fill: { color: C.white, transparency: 100 } },
  showLegend: false,
  catAxisLabelFontFace: BODY, catAxisLabelFontSize: 12, catAxisLabelColor: C.ink,
  valAxisLabelFontFace: BODY, valAxisLabelFontSize: 12, valAxisLabelColor: C.ink,
  catAxisLineShow: true, catAxisLineColor: C.axis,
  valAxisLineShow: false,
  valGridLine: { style: 'solid', size: 1, color: C.gridline },
  catGridLine: { style: 'none' },
  border: { pt: 0, color: C.white },
};

/* ------------------------------------------------------ 13. current demand */

function slide13() {
  const s = pptx.addSlide();
  deco(s, { word: true });
  header(s, 'Real Estate \u2013 Current Demand');
  rect(s, 0.665, 1.752, 6.002, 0.665, C.orange);
  rect(s, 7.0, 1.752, 5.668, 0.665, C.amber);
  hx(s, 'Fundamental Factors Expected to Drive Growth',
    { x: 0.788, y: 1.899, w: 5.755, h: 0.37, fontSize: 16, color: C.white, align: 'center', valign: 'top' });
  hx(s, 'Current Residental Data',
    { x: 6.997, y: 1.899, w: 5.668, h: 0.37, fontSize: 16, color: C.white, align: 'center', valign: 'top' });

  [
    { y: 2.417, col: C.amber, tag: 'Hospitality', body: 'PLACEHOLDER' },
    { y: 3.761, col: C.orange, tag: 'Retail', body: 'PLACEHOLDER' },
    { y: 5.083, col: C.teal, tag: 'Commercial', body: 'PLACEHOLDER' },
  ].forEach((r) => {
    rect(s, 0.665, r.y, 2.236, 1.0, r.col);
    rect(s, 0.665, r.y, 6.002, 1.0, null, { line: { color: r.col, width: 1 } });
    hx(s, r.tag, { x: 0.711, y: r.y + 0.298, w: 2.144, h: 0.404, color: C.white, align: 'center', valign: 'middle' });
    tx(s, r.body, { x: 3.333, y: r.y + 0.162, w: 3.031, h: 0.675, lineSpacingMultiple: 1.5, valign: 'top' });
  });

  tx(s, 'Square Feet in Millions', { x: 6.997, y: 2.493, w: 5.668, h: 0.303, align: 'center', valign: 'top' });
  const cats = ['2027', '2028', '2029'];
  s.addChart(pptx.ChartType.bar, [
    { name: 'Series 1', labels: cats, values: [100, 330, 300] },
    { name: 'Series 2', labels: cats, values: [240, 200, 215] },
    { name: 'Series 3', labels: cats, values: [313, 290, 225] },
  ], Object.assign({}, CHART_BASE, {
    x: 6.997, y: 2.866, w: 5.668, h: 3.884,
    barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 20, barOverlapPct: 100,
    chartColors: [C.amber, C.teal, C.orange],
    valAxisMaxVal: 1000, valAxisMinVal: 0, valAxisMajorUnit: 100,
    valAxisLabelFormatCode: '$#,##0',
  }));

  tx(s, '*Ultrices pellentes sed natoque sed non magnis, London, England',
    { x: 0.651, y: 6.266, w: 5.712, h: 0.373, lineSpacingMultiple: 1.5, valign: 'top' });
}

/* -------------------------------------------------------- 14. sales growth */

function slide14() {
  const s = pptx.addSlide();
  deco(s, { word: true });
  header(s, 'Real Estate \u2013 Sales Growth');
  s.addChart(pptx.ChartType.bar, [{
    name: 'Series 1',
    labels: ['2027', '2028', '2029', '2030', '2031'],
    values: [4000, 6300, 6790, 7755, 9234],
  }], Object.assign({}, CHART_BASE, {
    x: 0.822, y: 1.681, w: 11.689, h: 5.167,
    barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 5, barOverlapPct: -14,
    chartColors: [C.peach, C.peach, C.peach, C.peach, C.orange],
    valAxisMaxVal: 10000, valAxisMinVal: 1000, valAxisMajorUnit: 1000,
    valAxisLabelFormatCode: '$#,##0',
    showValue: true, dataLabelPosition: 'inBase', dataLabelFormatCode: '$#,##0',
    dataLabelColor: C.white, dataLabelFontFace: BODY, dataLabelFontSize: 16,
  }));
  // house markers riding on top of each column
  [
    { x: 2.317, y: 4.14 }, { x: 4.45, y: 2.977 }, { x: 6.583, y: 2.724 },
    { x: 8.717, y: 2.239 }, { x: 10.846, y: 1.495 },
  ].forEach((h, i) => icon(s, 'home', h.x, h.y, 0.889, i === 4 ? C.orange : C.peach));
}

/* ------------------------------------------------- 15. market opportunity */

function slide15() {
  const s = pptx.addSlide();
  deco(s, { word: true });
  header(s, 'Massive Market Opportunity');
  [
    { x: 7.558, y: 1.898, d: 4.952, col: C.orange, ty: 2.117, label: '$3.4 Trillion' },
    { x: 7.943, y: 2.667, d: 4.182, col: C.amber, ty: 2.935, label: '$721 Billion' },
    { x: 8.366, y: 3.512, d: 3.336, col: C.teal, ty: 3.78, label: '$542 Billion' },
    { x: 8.789, y: 4.358, d: 2.49, col: C.orange, ty: 5.401, label: '$267 Million' },
  ].forEach((o) => {
    oval(s, o.x, o.y, o.d, o.d, o.col);
    hx(s, o.label, { x: 8.789, y: o.ty, w: 2.491, h: 0.404, color: C.white, align: 'center', valign: 'middle' });
  });
  const rows = [
    { y: 2.056, col: C.orange, ico: 'people', lead: 2.521, leadW: 1.611 },
    { y: 3.262, col: C.amber, ico: 'home', lead: 3.772, leadW: 0.984 },
    { y: 4.467, col: C.teal, ico: 'board', lead: 4.977, leadW: 0.984 },
    { y: 5.668, col: C.orange, ico: 'city', lead: 6.179, leadW: 1.611 },
  ];
  rows.forEach((r) => {
    iconDisc(s, r.ico, 1.3765, r.y + 0.5105, 1.021, r.col, 0.34);
    hx(s, 'Description', { x: 1.983, y: r.y + 0.006, w: 4.228, h: 0.438, fontSize: 20, color: r.col, valign: 'middle' });
    tx(s, 'PLACEHOLDER',
      { x: 1.983, y: r.y + 0.324, w: 4.472, h: 0.675, lineSpacingMultiple: 1.5, valign: 'top' });
    line(s, 6.456, r.lead, r.leadW, 0, r.col, 1);
  });
}

/* ---------------------------------------------- 16. future market challenges */

function slide16() {
  const s = pptx.addSlide();
  deco(s, { word: true });
  header(s, 'Future Market Challenges');
  rect(s, -1.489, 3.915, 5.0, 0.669, C.orange, { rotate: 270 });
  hx(s, 'Future Marketing Challenges',
    { x: -1.489, y: 4.065, w: 5.0, h: 0.37, fontSize: 16, color: C.white, align: 'center', rotate: 270 });
  [
    { y: 1.756, col: C.orange, ico: 'people' },
    { y: 3.083, col: C.amber, ico: 'home' },
    { y: 4.417, col: C.teal, ico: 'board' },
    { y: 5.75, col: C.orange, ico: 'city' },
  ].forEach((r) => {
    s.addShape('rightArrow', {
      x: 1.665, y: r.y, w: 3.002, h: 1.0,
      fill: { color: r.col }, line: noLine,
    });
    hx(s, 'Detail Information', { x: 1.665, y: r.y + 0.331, w: 2.595, h: 0.337, fontSize: 14, color: C.white, align: 'center', valign: 'middle' });
    rect(s, 4.122, r.y, 8.535, 1.0, null, { line: { color: r.col, width: 1 } });
    iconDisc(s, r.ico, 5.624, r.y + 0.5, 0.726, r.col, 0.34);
    hx(s, 'Detail Information', { x: 6.113, y: r.y + 0.176, w: 3.733, h: 0.404, color: r.col });
    tx(s, 'PLACEHOLDER',
      { x: 6.113, y: r.y + 0.45, w: 6.021, h: 0.373, lineSpacingMultiple: 1.5, valign: 'top' });
  });
}

/* ------------------------------------------------------- 17. big number */

function slide17() {
  const s = pptx.addSlide();
  deco(s);
  rect(s, 0, 0, 13.333, 7.5, C.orange, { fill: { color: C.orange, transparency: 9 } });
  hatch(s, 0, 0, 0.651, 0.652, C.orange);
  hatch(s, 12.683, 6.848, 0.651, 0.652, C.orange);
  houseMark(s, 6.934, 0.472, 0.247, C.white);
  hx(s, 'Luxure', { x: 6.152, y: 0.444, w: 0.801, h: 0.303, fontSize: 12, color: C.white, valign: 'middle' });
  tx(s, 'How Big is Real Estate Market?',
    { x: 4.059, y: 2.679, w: 5.215, h: 0.438, fontSize: 20, color: C.white, align: 'center' });
  s.addText('$78,9856,000,000', {
    x: 0.917, y: 3.033, w: 11.5, h: 1.45,
    fontFace: HEAD, fontSize: 80, bold: true, color: C.white, align: 'center', valign: 'top', lineSpacingMultiple: 1,
  });
  tx(s, '*Data Based on Annual Property Transactions Volume',
    { x: 3.078, y: 4.484, w: 7.178, h: 0.337, fontSize: 14, color: C.white, align: 'center' });
}

/* -------------------------------------------- europe / usa map placeholders */

/** Blocky stand-in for the deck's vector maps: a scatter of rounded tiles. */
function mapBlocks(s, x, y, w, h, cells, palette) {
  cells.forEach(([cx, cy, cw, ch, tone]) => {
    rect(s, x + cx * w, y + cy * h, cw * w, ch * h, palette[tone]);
  });
}

const EUROPE_CELLS = [
  [0.02, 0.34, 0.14, 0.30, 0], [0.14, 0.10, 0.10, 0.20, 0], [0.20, 0.30, 0.12, 0.22, 1],
  [0.26, 0.05, 0.14, 0.22, 2], [0.30, 0.52, 0.14, 0.26, 0], [0.34, 0.26, 0.16, 0.24, 0],
  [0.42, 0.02, 0.12, 0.20, 2], [0.46, 0.48, 0.12, 0.30, 1], [0.50, 0.22, 0.18, 0.24, 2],
  [0.58, 0.52, 0.16, 0.24, 0], [0.62, 0.04, 0.18, 0.20, 2], [0.66, 0.26, 0.16, 0.22, 2],
  [0.76, 0.46, 0.20, 0.26, 0], [0.80, 0.16, 0.18, 0.26, 2], [0.10, 0.62, 0.14, 0.22, 0],
  [0.40, 0.74, 0.14, 0.20, 1], [0.60, 0.78, 0.16, 0.18, 0],
];

const USA_CELLS = [
  [0.02, 0.20, 0.14, 0.26, 0], [0.02, 0.46, 0.12, 0.26, 0], [0.14, 0.14, 0.12, 0.24, 0],
  [0.14, 0.40, 0.14, 0.26, 0], [0.10, 0.66, 0.16, 0.22, 0], [0.26, 0.12, 0.14, 0.24, 0],
  [0.26, 0.38, 0.14, 0.26, 0], [0.26, 0.64, 0.16, 0.26, 0], [0.40, 0.10, 0.14, 0.26, 0],
  [0.40, 0.38, 0.14, 0.24, 0], [0.42, 0.62, 0.16, 0.30, 0], [0.54, 0.12, 0.14, 0.24, 0],
  [0.54, 0.38, 0.14, 0.22, 0], [0.58, 0.60, 0.14, 0.22, 0], [0.68, 0.14, 0.14, 0.22, 0],
  [0.68, 0.38, 0.12, 0.24, 0], [0.72, 0.60, 0.14, 0.20, 0], [0.80, 0.18, 0.14, 0.22, 0],
  [0.80, 0.40, 0.14, 0.22, 0], [0.84, 0.62, 0.12, 0.26, 0], [0.86, 0.80, 0.10, 0.16, 0],
];

/* ---------------------------------------------- 18. renting vs homebuying */

function slide18() {
  const s = pptx.addSlide();
  deco(s, { word: true });
  header(s, [
    { text: 'Others Focus on Renting Real Estate, ', options: { breakLine: true } },
    { text: '', options: { breakLine: true } },
    { text: 'but Homebuying is the Key Problem' },
  ], { h: 1.449 });
  rect(s, 0.665, 2.417, 5.002, 0.67, C.orange);
  rect(s, 6.0, 2.417, 6.668, 0.67, C.amber);
  hx(s, 'Supply/demand Imbalance', { x: 0.665, y: 2.566, w: 5.002, h: 0.37, fontSize: 16, color: C.white, align: 'center', valign: 'top' });
  hx(s, 'Average Days on Market/Number of Offers', { x: 6.0, y: 2.566, w: 6.668, h: 0.37, fontSize: 16, color: C.white, align: 'center', valign: 'top' });
  rect(s, 0.665, 3.087, 5.002, 3.663, null, { line: { color: C.orange, width: 1 } });
  line(s, 0.665, 4.918, 5.002, 0, C.orange, 1);

  [
    { y: 3.305, title: '2028 Expected Buyer Activity', dots: [C.amber, C.teal], pal: [C.lightGrey, C.amber, C.teal], my: 3.088 },
    { y: 5.131, title: '2029 Expected Seller Activity', dots: ['6D6D6D', C.grey], pal: [C.lightGrey, '6D6D6D', C.grey], my: 4.915 },
  ].forEach((p) => {
    mapBlocks(s, 3.2, p.my, 2.47, 1.806, EUROPE_CELLS, p.pal);
    hx(s, p.title, { x: 1.05, y: p.y, w: 1.805, h: 0.64, fontSize: 16, color: C.orange, valign: 'middle' });
    arrowDot(s, 1.145, p.y + 0.695, 0.236, C.orange);
    legendDot(s, 1.146, p.y + 1.201, p.dots[0], 'Details');
    legendDot(s, 2.056, p.y + 1.195, p.dots[1], 'Details');
  });

  rect(s, 6.0, 3.083, 6.668, 3.667, null, { line: { color: C.amber, width: 1 } });
  const cats = ['2027', '2028', '2029'];
  s.addChart(pptx.ChartType.bar, [
    { name: 'Series 1', labels: cats, values: [100, 130, 300] },
    { name: 'Series 2', labels: cats, values: [240, 200, 115] },
  ], Object.assign({}, CHART_BASE, {
    x: 6.33, y: 3.237, w: 6.009, h: 2.715,
    barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 20, barOverlapPct: 100,
    chartColors: [C.amber, C.teal],
    valAxisMaxVal: 500, valAxisMinVal: 0, valAxisMajorUnit: 50,
    valAxisLabelFormatCode: '$#,##0',
  }));
  rect(s, 6.0, 6.083, 6.668, 0.667, C.amber);
  icon(s, 'pin', 6.968, 6.278, 0.225, C.white);
  hx(s, 'First Location', { x: 7.215, y: 6.215, w: 2.089, h: 0.404, color: C.white, valign: 'middle' });
  tx(s, 'Description, London, England', { x: 9.018, y: 6.265, w: 2.448, h: 0.303, color: C.white, valign: 'middle' });
}

/* ---------------------------------------------------------- 19. traction */

function slide19() {
  const s = pptx.addSlide();
  deco(s, { word: true });
  header(s, 'Early Traction Has Been Exciting');
  s.addChart(pptx.ChartType.bar, [{
    name: 'Series 1',
    labels: ['January', 'February', 'March', 'April', 'Mei', 'June', 'July'],
    values: [46, 63, 67, 75, 94, 44, 26],
  }], Object.assign({}, CHART_BASE, {
    x: 0.822, y: 1.681, w: 11.689, h: 5.167,
    barDir: 'bar', barGrouping: 'clustered', barGapWidthPct: 11,
    chartColors: [C.peach, C.peach, C.peach, C.peach, C.orange, C.peach, C.peach],
    valAxisMaxVal: 100, valAxisMinVal: 0, valAxisHidden: true,
    showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '0" Units"',
    dataLabelColor: C.ink, dataLabelFontFace: BODY, dataLabelFontSize: 12,
    valGridLine: { style: 'solid', size: 1, color: C.gridline },
  }));
  for (let i = 0; i < 7; i += 1) icon(s, 'home', 1.994, 2.084 + i * 0.6935, 0.233, C.white);
  s.addShape('wedgeRectCallout', {
    x: 9.644, y: 1.57, w: 2.867, h: 1.218,
    fill: { color: C.orange }, line: noLine,
  });
  hx(s, 'With $45.7M Gross Profit, Fastest within the First 3 Weeks',
    { x: 10.019, y: 1.775, w: 2.117, h: 0.808, fontSize: 14, color: C.white, valign: 'middle' });
}

/* ------------------------------------------------------------ 20. leads */

function slide20() {
  const s = pptx.addSlide();
  deco(s, { word: true });
  header(s, 'Generate Quality Leads and Convert them');
  s.addChart(pptx.ChartType.doughnut, [{
    name: 'Sales', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr'], values: [8.2, 3.2, 1.4],
  }], Object.assign({}, CHART_BASE, {
    x: 0.822, y: 2.48, w: 6.117, h: 4.078,
    chartColors: [C.orange, C.amber, C.teal],
    holeSize: 52, firstSliceAng: 0, showValue: false, dataBorder: { pt: 0, color: C.white },
  }));
  iconDisc(s, 'home', 3.881, 4.178, 0.66, C.white, 0.6);
  icon(s, 'home', 3.681, 4.006, 0.4, C.orange);
  hx(s, paras(['Data ', 'Convert'], { align: 'center' }),
    { x: 3.25, y: 4.392, w: 1.261, h: 0.64, fontSize: 16, color: C.orange, align: 'center', valign: 'middle' });
  [
    { x: 5.924, y: 4.092, col: C.orange, pct: '82%', align: 'left' },
    { x: 0.651, y: 3.565, col: C.amber, pct: '32%', align: 'right' },
    { x: 2.073, y: 1.842, col: C.teal, pct: '14%', align: 'right' },
  ].forEach((l) => {
    hx(s, l.pct, { x: l.x, y: l.y, w: 1.253, h: 0.64, fontSize: 32, color: l.col, align: l.align, valign: 'middle' });
    tx(s, 'Description', { x: l.x, y: l.y + 0.529, w: 1.253, h: 0.303, align: l.align, valign: 'middle' });
  });
  legendDot(s, 2.57, 6.66, C.orange, 'Details');
  legendDot(s, 3.482, 6.66, C.amber, 'Details');
  legendDot(s, 4.392, 6.655, C.teal, 'Details');

  rect(s, 7.333, 1.75, 5.335, 0.668, C.orange);
  rect(s, 7.333, 2.418, 5.335, 4.332, null, { line: { color: C.orange, width: 1 } });
  hx(s, 'High Real Buyer Conversion',
    { x: 7.333, y: 1.899, w: 5.335, h: 0.37, fontSize: 16, color: C.white, align: 'center', valign: 'top' });
  s.addChart(pptx.ChartType.bar, [{
    name: 'Series 1', labels: ['Description', 'Description', 'Description'], values: [14, 32, 82],
  }], Object.assign({}, CHART_BASE, {
    x: 7.534, y: 2.576, w: 4.933, h: 4.017,
    barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 11,
    chartColors: [C.teal, C.amber, C.orange],
    valAxisMaxVal: 100, valAxisMinVal: 0, valAxisHidden: true,
    showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '0" Units"',
    dataLabelColor: C.ink, dataLabelFontFace: BODY, dataLabelFontSize: 12,
  }));
}

/* ------------------------------------------------------- 21. expansion map */

/** Pill + arrow-head marker used on the expansion map, with its leader line. */
function mapTag(s, x, y, color, leadX, leadY, leadW, flip) {
  rect(s, x, y, 1.607, 0.472, color);
  s.addShape('triangle', {
    x: x + 1.437, y: y + 0.036, w: 0.472, h: 0.4,
    fill: { color }, line: noLine, rotate: 90,
  });
  tx(s, 'Description', { x, y: y + 0.084, w: 1.607, h: 0.303, color: C.white, align: 'center', valign: 'middle' });
  line(s, leadX, leadY, leadW, 0, color, 1, { line: { color, width: 1, beginArrowType: 'oval', endArrowType: 'oval' }, flipV: flip });
}

function slide21() {
  const s = pptx.addSlide();
  deco(s, { word: true });
  header(s, [
    { text: 'Our Expansion Plan Could Take Us ', options: { breakLine: true } },
    { text: 'to $1.6B in Revenue' },
  ], { h: 1.449 });
  mapBlocks(s, 5.206, 2.324, 7.393, 4.932, USA_CELLS, [C.paleGrey]);
  [
    { y: 2.463, col: C.orange, big: '$45M' },
    { y: 3.564, col: C.amber, big: '178%' },
    { y: 4.664, col: C.teal, big: '$1.1B' },
    { y: 5.765, col: C.orange, big: '6783' },
  ].forEach((r) => {
    rect(s, 0.701, r.y, 2.965, 0.98, r.col);
    hx(s, r.big, { x: 0.871, y: r.y + 0.171, w: 1.551, h: 0.64, fontSize: 32, color: C.white, valign: 'middle' });
    tx(s, 'Sequi nesciu neque eus', { x: 2.221, y: r.y + 0.238, w: 1.194, h: 0.505, color: C.white, valign: 'middle' });
  });
  [
    { x: 4.793, y: 2.52, col: C.teal, lx: 6.401, ly: 2.756, lw: 0.591, flip: true },
    { x: 4.444, y: 3.687, col: C.orange, lx: 6.195, ly: 3.923, lw: 0.591, flip: true },
    { x: 8.87, y: 2.902, col: C.amber, lx: 10.558, ly: 3.138, lw: 1.113, flip: false },
    { x: 8.572, y: 4.789, col: C.orange, lx: 10.323, ly: 5.026, lw: 0.591, flip: true },
    { x: 5.383, y: 4.688, col: C.amber, lx: 7.072, ly: 4.924, lw: 0.591, flip: true },
    { x: 7.264, y: 3.869, col: C.teal, lx: 8.953, ly: 4.105, lw: 0.591, flip: true },
    { x: 9.176, y: 5.874, col: C.amber, lx: 10.864, ly: 6.11, lw: 0.591, flip: true },
  ].forEach((t) => mapTag(s, t.x, t.y, t.col, t.lx, t.ly, t.lw, t.flip));
  legendDot(s, 4.719, 5.916, C.orange, 'Details');
  legendDot(s, 5.631, 5.916, C.amber, 'Description');
  legendDot(s, 4.719, 6.297, C.teal, 'Details');
  legendDot(s, 5.631, 6.297, C.orange, 'Description');
}

/* -------------------------------------------------------- 22. milestones */

const MILESTONES = [
  { x: 0.822, col: C.orange, date: 'Jan 2028', bullets: ['Aliqu sed consiesi quis autem eum ', 'Iure reprehendiriu voluptate non'] },
  { x: 3.3, col: C.amber, date: 'Jul 2029', bullets: ['Aliqu sed consiesi quis autem eum ', 'Iure reprehendriu'] },
  { x: 5.783, col: C.teal, date: 'Mei 2030', bullets: ['Aliqu sed consiesi quis autem eum ', 'Velit esseu consq'] },
  { x: 8.261, col: C.orange, date: 'Sep 2031', bullets: ['Aliqu sed consiesi quis autem eum ', 'Iure reprehendiriu voluptate non'] },
  { x: 10.744, col: C.amber, date: 'Dec 2032', bullets: ['Aliqu sed consiesi quis autem eum ', 'Iure reprehendiriu Velit esseu consq'] },
];

function slide22() {
  const s = pptx.addSlide();
  deco(s, { word: true });
  header(s, 'Future Company Milestones');
  line(s, 1.745, 4.767, 9.843, 0, C.orange, 1);
  MILESTONES.forEach((m, i) => {
    hx(s, m.date, { x: m.x, y: 1.898, w: 1.767, h: 0.438, fontSize: 20, color: m.col, align: 'center' });
    arrowDot(s, [1.55, 4.027, 6.511, 8.994, 11.477][i], 4.611, 0.312, m.col);
    hx(s, 'Description', { x: m.x, y: 5.092, w: 1.767, h: 0.404, color: m.col });
    tx(s, paras(m.bullets, { bullet: BULLET, lineSpacingMultiple: 1.5 }),
      { x: m.x, y: 5.399, w: i === 0 ? 1.903 : 1.767, h: 1.281, valign: 'top' });
  });
}

/* ----------------------------------------------------- 23. org structure */

const ORG_TOP = [
  { x: 3.183, name: 'Ashley P. Albert' },
  { x: 7.906, name: 'David M. Jones' },
];
const ORG_MID = ['Ashley P. Albert', 'Jeanette J. Clark', 'Leon D. Morgan', 'Mildred G. Grene', 'Mark J. Ferguson'];
const ORG_BOT = [
  { x: 2.003, name: 'Ruby J. Bergman', role: 'Head of Marketing' },
  { x: 4.364, name: 'Lucas B. Boswell', role: 'Head of Creative' },
  { x: 6.725, name: 'Mary M. Warners', role: 'Head of Resources' },
  { x: 9.087, name: 'Stephen N. Wenz', role: 'Head of Brand Sales' },
];

function orgCard(s, x, y, color, name, role) {
  rect(s, x, y, 2.244, 1.101, C.white, { line: { color, width: 1 } });
  hx(s, name, { x, y: y + 0.434, w: 2.244, h: 0.37, fontSize: 16, color, align: 'center' });
  tx(s, role, { x, y: y + 0.642, w: 2.244, h: 0.35, fontSize: 11, align: 'center', lineSpacingMultiple: 1.5, valign: 'top' });
}

function slide23() {
  const s = pptx.addSlide();
  deco(s, { word: true });
  header(s, 'Organizational Structure');
  line(s, 0.411, 6.185, 12.511, 0, C.orange, 1);
  ORG_TOP.forEach((t) => orgCard(s, t.x, 2.211, C.orange, t.name, 'General Manager'));
  ORG_BOT.forEach((b) => orgCard(s, b.x, 5.635, C.teal, b.name, b.role));
  line(s, 0.411, 4.474, 12.512, 0, C.orange, 1);
  ORG_MID.forEach((n, i) => orgCard(s, 0.822 + i * 2.361, 3.924, C.amber, n, 'Project Manager'));
  line(s, 4.305, 3.312, 0, 0.186, C.orange, 1);
  line(s, 9.028, 3.312, 0, 0.186, C.orange, 1);
  line(s, 0.411, 4.463, 0, 1.722, C.orange, 1);
  line(s, 12.922, 4.463, 0, 1.722, C.orange, 1);
}

/* ------------------------------------------------------- 24. partnerships */

const PARTNERS = [
  { x: 3.44, y: 1.898, col: C.orange, ico: 'people', cap: 'Quia numqua modi tempora' },
  { x: 6.543, y: 1.898, col: C.amber, ico: 'home', cap: 'Incidunt labore dolor magnam' },
  { x: 9.646, y: 1.898, col: C.teal, ico: 'chart', cap: 'Aliqu sed consiesi quis autem' },
  { x: 3.44, y: 4.492, col: C.teal, ico: 'board', cap: 'Incidunt labore dolor magnam' },
  { x: 6.543, y: 4.492, col: C.amber, ico: 'city', cap: 'Aliqu sed consiesi quis autem' },
  { x: 9.646, y: 4.492, col: C.orange, ico: 'globe', cap: 'Quia numqua modi tempora' },
];

function slide24() {
  const s = pptx.addSlide();
  deco(s, { word: true });
  header(s, 'Global Leading Landlords Partnerships');
  PARTNERS.forEach((p) => {
    rect(s, p.x, p.y, 2.865, 2.356, null, { line: { color: p.col, width: 0.75 } });
    iconDisc(s, p.ico, p.x + 1.432, p.y + 0.924, 1.006, p.col, 0.34);
    tx(s, p.cap, { x: p.x, y: p.y + 1.562, w: 2.865, h: 0.373, align: 'center', lineSpacingMultiple: 1.5, valign: 'top' });
  });
  [
    { y: 2.119, col: C.orange },
    { y: 4.713, col: C.teal },
  ].forEach((b) => {
    hx(s, 'Description', { x: 0.651, y: b.y, w: 2.424, h: 0.438, fontSize: 20, color: b.col });
    tx(s, paras([
      'Aliqu sed consi quis autem eum  iure reprehendiriu ',
      'Velit esseu consqie doneic felis ultricies pellentesq ',
      'Neu sociis magnis partur',
    ], { bullet: BULLET, lineSpacingMultiple: 1.5 }), { x: 0.651, y: b.y + 0.329, w: 2.526, h: 1.584, valign: 'top' });
  });
}

/* ------------------------------------------------------- 25. any question */

function slide25() {
  const s = pptx.addSlide();
  deco(s, { chips: false, mark: true });
  s.addShape('rect', {
    x: 0, y: 0.652, w: 13.333, h: 6.197,
    fill: { color: C.orange, transparency: 9 }, line: noLine,
  });
  hatch(s, 0, 0, 0.651, 0.652);
  hatch(s, 12.683, 6.848, 0.651, 0.652);
  s.addText('Any Other\n Question?', {
    x: 1.42, y: 2.322, w: 11.5, h: 3.424,
    fontFace: HEAD, fontSize: 88, bold: true, color: C.white, valign: 'top', lineSpacingMultiple: 1,
  });
  hx(s, 'Description', { x: 9.075, y: 2.944, w: 2.424, h: 0.438, fontSize: 20, color: C.white });
  tx(s, paras([
    'Aliqu sed consi quis autem eum  ',
    'Velit esseu consqie doneic felis ultricies pellentesq  repreh',
    'Neu sociis magnis sed partur',
  ], { bullet: BULLET, lineSpacingMultiple: 1.5, color: C.white }),
  { x: 9.075, y: 3.274, w: 2.838, h: 1.281, color: C.white, valign: 'top' });
}

/* ------------------------------------------------------------ 26. contact */

function slide26() {
  const s = pptx.addSlide();
  rect(s, 0, 0, 13.333, 1.898, '2D2D2D');
  chips(s, 0, 0.524);
  houseMark(s, 12.435, 0.472, 0.247);
  hx(s, 'Luxure', { x: 11.653, y: 0.444, w: 0.801, h: 0.303, fontSize: 12, color: C.white, valign: 'middle' });
  hatch(s, 12.683, 6.848, 0.651, 0.652);
  header(s, 'Contact Us Here Now', { color: C.white, subColor: C.white });
  tx(s, 'PLACEHOLDER',
    { x: 0.651, y: 5.817, w: 3.016, h: 0.978, lineSpacingMultiple: 1.5, valign: 'top' });
  [
    { ix: 4.131, iy: 5.88, ico: 'mail', tx: 4.567, w: 2.121, head: 'Our Email', lines: ['company@email.com', 'branch@email.com'] },
    { ix: 7.224, iy: 5.851, ico: 'globe', tx: 7.656, w: 1.987, head: 'Our Website', lines: ['www.company.com', 'www.branch.com'] },
    { ix: 10.278, iy: 5.876, ico: 'phone', tx: 10.692, w: 1.987, head: 'Our Phone', lines: ['+1 623 7243 9834', '+3 724 7745 8560'] },
  ].forEach((c) => {
    icon(s, c.ico, c.ix, c.iy, 0.3, C.orange);
    hx(s, c.head, { x: c.tx, y: 5.764, w: c.w, h: 0.404, color: C.orange });
    tx(s, paras(c.lines, { lineSpacingMultiple: 1.5 }), { x: c.tx, y: 6.077, w: c.w, h: 0.771, fontSize: 14, valign: 'middle' });
  });
}

/* ---------------------------------------------------------- 27. thank you */

function slide27() {
  const s = pptx.addSlide();
  rect(s, 0, 0, 13.333, 7.5, '212629');
  hatch(s, 0, 0, 0.651, 0.652, C.ink);
  hatch(s, 12.683, 6.848, 0.651, 0.652, C.ink);
  houseMark(s, 6.424, 2.268, 0.485);
  hx(s, 'Luxure', { x: 5.863, y: 2.754, w: 1.607, h: 0.505, fontSize: 24, color: C.white, align: 'center', valign: 'middle' });
  s.addText('Thank You', {
    x: 4.089, y: 3.28, w: 5.156, h: 1.012,
    fontFace: HEAD, fontSize: 54, bold: true, color: C.white, align: 'center', valign: 'middle', lineSpacingMultiple: 1,
  });
  tx(s, 'PLACEHOLDER',
    { x: 4.089, y: 4.126, w: 5.155, h: 0.675, color: C.white, align: 'center', lineSpacingMultiple: 1.5, valign: 'top' });
  chips(s, 6.564, 4.795, 0.223, 0.205, true);
}

/* ---------------------------------------------------------------- output */

[
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
  slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18,
  slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27,
].forEach((build) => build());

pptx.writeFile({
  fileName: path.join(__dirname, '0e9090b5-74b8-4c7c-b770-721d32eda74c_grok_final.pptx'),
}).then((f) => console.log('wrote', f));
