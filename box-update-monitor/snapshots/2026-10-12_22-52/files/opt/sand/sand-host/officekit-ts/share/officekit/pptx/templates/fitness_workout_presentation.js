/**
 * "WorkOut" presentation template - 30 slides, 16:9 (13.333 x 7.5 in).
 *
 * Rebuilt with pptxgenjs only. Photographic placeholders from the original
 * deck are drawn as flat tinted rectangles labelled "[image]".
 *
 * Run:  node 0d1753e7-24c2-4b8b-b741-a3450abc6cd0_grok_final.js
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const INK = '262626'; // near-black headings (tx1 lum 85%)
const GRAY = '808080'; // body copy (tx1 lum 50%)
const BLUE = '6D9CE1'; // accent1
const WHITE = 'FFFFFF';
const MIST = 'F2F2F2'; // bg1 lum 95%
const SILVER = 'D9D9D9'; // bg1 lum 85%
const PHOTO = 'D9D9D9'; // stand-in for an embedded photograph

const HEAD = 'Montserrat'; // +mj-lt
const BODY = 'Open Sans'; // +mn-lt

// pptxgenjs rewrites shadow objects in place, so hand out a fresh one each time.
const softShadow = () => ({ type: 'outer', color: 'A6A6A6', blur: 30, offset: 0.0001, angle: 90, opacity: 0.4 });
const dropShadow = () => ({ type: 'outer', color: INK, blur: 31, offset: 15, angle: 135, opacity: 0.4 });

/* ------------------------------------------------------------------- copy */

const LONG =
  'PLACEHOLDER' +
  'PLACEHOLDER' +
  'hanase baninanit sanisina nasilet atasa atebetmen';
/** first `n` words of the long lorem paragraph */
const clip = (n) => LONG.split(' ').slice(0, n).join(' ');

const INTRO = 'PLACEHOLDER';
const INTRO_LONG = INTRO + ' atasa atebetm sanhanasen hunai';
const INTRO_ALT = 'PLACEHOLDER';
const DARK_PARA =
  'PLACEHOLDER' +
  'PLACEHOLDER' +
  'sanisina nasilet atasa atebetmen';

const SVC = 'sanisinaia at cona salanasie atasanit atile nalatauit nuba';
const SUB = 'sanisi lana atasa';
const SUB2 = 'sanisi lana atasa hinu';
const CAPTION = 'hanaban duisanai nasan';
const TITLE = 'Exercise in the morning';
const PORTFOLIO = 'Portfolio Slides';

/* ---------------------------------------------------------------- helpers */

/** Solid rectangle, no outline. */
function box(s, x, y, w, h, color, opts) {
  s.addShape('rect', Object.assign({ x, y, w, h, fill: { color }, line: { type: 'none' } }, opts));
}

/** White card with the deck's soft drop shadow. */
function card(s, x, y, w, h, color) {
  box(s, x, y, w, h, color || WHITE, { shadow: softShadow() });
}

/**
 * Picture frames inherited from the template. They are empty in the source
 * deck, so they are drawn as plain white plates underneath everything else -
 * their coordinates document where artwork would sit.
 */
function frames(s, boxes) {
  boxes.forEach(([x, y, w, h]) => s.addShape('rect', {
    x, y, w, h, fill: { color: WHITE }, line: { type: 'none' },
  }));
}

/** Stand-in for a photograph that was embedded in the original deck. */
function photo(s, x, y, w, h, rotate) {
  s.addShape('rect', { x, y, w, h, rotate, fill: { color: PHOTO }, line: { type: 'none' } });
  s.addText('[image]', {
    x, y, w, h, rotate, fontFace: BODY, fontSize: 10, color: '8C8C8C',
    align: 'center', valign: 'middle',
  });
}

/** Montserrat display heading (slide titles). */
function heading(s, text, x, y, w, h, size, color, align, rotate) {
  s.addText(text, {
    x, y, w, h, rotate, fontFace: HEAD, fontSize: size, bold: true, color: color || INK,
    align: align || 'left', valign: 'top',
  });
}

/** 20pt bold person name on the profile slides (body face, not the display face). */
function name(s, text, x, y, w, color) {
  s.addText(text, {
    x, y, w, h: 0.438, fontFace: BODY, fontSize: 20, bold: true, color: color || INK, valign: 'top',
  });
}

/** 11pt justified body copy at 1.5 line spacing. */
function body(s, text, x, y, w, h, o) {
  o = o || {};
  s.addText(text, {
    x, y, w, h, fontFace: BODY, fontSize: o.size || 11, color: o.color || GRAY,
    align: o.align || 'justify', valign: 'top', lineSpacingMultiple: 1.5,
  });
}

/** Two paragraphs separated by a blank line. */
function bodyPair(s, first, second, x, y, w, h, color) {
  body(s, [
    { text: first, options: { breakLine: true } },
    { text: '', options: { breakLine: true } },
    { text: second },
  ], x, y, w, h, { color });
}

/** 14pt bold service label. */
function label(s, text, x, y, w, color, size) {
  s.addText(text, {
    x, y, w: w || 1.699, h: 0.337, fontFace: BODY, fontSize: size || 14, bold: true,
    color: color || INK, valign: 'top',
  });
}

/** 9pt sub-label under a service label. */
function tiny(s, text, x, y, w, color) {
  s.addText(text, {
    x, y, w: w || 1.691, h: 0.304, fontFace: BODY, fontSize: 9, color: color || GRAY,
    valign: 'top', lineSpacingMultiple: 1.5,
  });
}

/** label + sub-label (+ optional paragraph) block used across the deck. */
function service(s, x, y, o) {
  label(s, o.title, x, y, o.tw, o.color, o.tsize);
  tiny(s, o.sub || SUB, x + 0.008, y + (o.sdy || 0.256), o.sw, o.color);
  if (o.body) body(s, o.body, x, y + (o.dy || 0.624), o.bw || 2.196, o.bh || 0.627, { color: o.color });
}

/** 16pt "service" word above a small grey caption (chart / gallery captions). */
function serviceCaption(s, x, y, text) {
  s.addText('service', {
    x, y, w: 1.096, h: 0.37, fontFace: HEAD, fontSize: 16, color: INK, align: 'center', valign: 'top',
  });
  s.addText(text || CAPTION, {
    x: x - 0.453, y: y + 0.327, w: 2.002, h: 0.349, fontFace: BODY, fontSize: 11, color: GRAY,
    align: 'center', valign: 'top', lineSpacingMultiple: 1.5,
  });
}

/** Blue diamond (45deg rounded square) with its number. */
function diamond(s, x, y, size, num) {
  s.addShape('roundRect', {
    x, y, w: size, h: size, rotate: 45, rectRadius: size / 6,
    fill: { color: BLUE }, line: { type: 'none' },
  });
  const big = size > 0.65;
  s.addText(num, {
    x: x + (big ? 0.029 : 0.051), y: y + (big ? 0.18 : 0.099),
    w: big ? 0.66 : 0.476, h: big ? 0.37 : 0.349,
    fontFace: BODY, fontSize: big ? 16 : 11, bold: true, color: WHITE,
    align: 'center', valign: 'top', lineSpacingMultiple: big ? undefined : 1.5,
  });
}

/** Three-bar "hamburger" mark. */
function menu(s, x, y, color) {
  [0, 0.225, 0.45].forEach((dy) => box(s, x, y + dy, 0.342, 0.05, color || BLUE));
}

const NAV_ITEMS = [['Title', 0.51, 0], ['Fitness', 0.698, 0.988], ['About', 0.63, 2.037]];

/** "Title  Fitness  About" mini nav; `rot` may be 0, 90 or 270. */
function nav(s, cx, cy, color, rot) {
  NAV_ITEMS.forEach(([text, w, off]) => {
    const mx = rot ? cx : cx + off;
    const my = rot === 270 ? cy - off : rot === 90 ? cy + off : cy;
    s.addText(text, {
      x: mx - w / 2, y: my - 0.1345, w, h: 0.269, rotate: rot || undefined,
      fontFace: BODY, fontSize: 10, bold: true, color: color || GRAY,
      align: 'center', valign: 'top', wrap: false,
    });
  });
}

/** "13 : 20" clock mark; the hours carry an extra outline in the original. */
function clock(s, x, y, color, big, small) {
  big = big || 28;
  s.addText([
    { text: '13 ', options: { fontSize: big, color, outline: { size: 1.5, color } } },
    { text: ': 20', options: { fontSize: small || Math.round(big * 0.643), color } },
  ], {
    x, y, w: big === 28 ? 1.177 : 1.054, h: big === 28 ? 0.572 : 0.505,
    fontFace: BODY, bold: true, color, align: 'center', valign: 'top', wrap: false,
  });
}

/** Thin accent rule. */
function rule(s, x, y, w, color) {
  s.addShape('line', { x, y, w, h: 0, line: { color: color || BLUE, width: 1.5 } });
}

/* ------------------------------------------------------------- the slides */

const SLIDES = [];

// 1 - cover
SLIDES.push((s) => {
  frames(s, [[0, 0, 13.333, 7.5]]);
  box(s, 2.365, 2.125, 8.604, 3.292, WHITE, { fill: { color: WHITE, transparency: 12 } });
  s.addText('WorkOut', {
    x: 3.635, y: 2.68, w: 6.062, h: 1.582, fontFace: HEAD, fontSize: 88, bold: true,
    underline: { style: 'sng' }, color: INK, align: 'center', valign: 'top', wrap: false,
  });
  s.addText('Presentation Template', {
    x: 7.66, y: 4.817, w: 3.201, h: 0.417, fontFace: BODY, fontSize: 14, color: GRAY,
    charSpacing: 3, valign: 'top', lineSpacingMultiple: 1.5,
  });
  nav(s, 1.026, 0.7935, WHITE);
  body(s, INTRO + ' atasa', 2.274, 5.808, 4.282, 0.555, { size: 9, color: WHITE, align: 'left' });
  clock(s, 11.751, 0.314, WHITE);
});

// 2 - title left, dark quote card right
SLIDES.push((s) => {
  frames(s, [[5.354, 1.51, 4.042, 4.479]]);
  heading(s, TITLE, 1.09, 1.629, 2.868, 2.121, 40);
  body(s, INTRO, 1.09, 5.045, 3.056, 0.934);
  clock(s, 11.751, 0.314, BLUE);
  nav(s, 8.55, 6.7455, GRAY);
  rule(s, 11.751, 6.778, 1.582);
  box(s, 9.396, 1.51, 3.938, 2.948, INK);
  body(s, DARK_PARA, 9.924, 2.093, 2.88, 1.767, { color: WHITE });
  menu(s, 5.354, 0.508);
});

// 3 - blue caption block between two photo columns
SLIDES.push((s) => {
  frames(s, [[0.677, 5.083, 3.354, 2.417], [4.99, 2.75, 3.354, 4.75], [9.302, 3.75, 3.354, 3.75]]);
  heading(s, TITLE, 9.233, 1.052, 3.491, 1.178, 32);
  nav(s, 0.828, 0.6205, GRAY);
  clock(s, 1.766, 4.043, BLUE);
  body(s, INTRO, 0.826, 2.408, 3.056, 0.934);
  box(s, 4.99, 1.238, 3.354, 1.512, BLUE);
  menu(s, 9.302, 2.641);
  body(s, INTRO.split(' ').slice(0, 7).join(' ') + ' duisan', 5.426, 1.666, 2.481, 0.656, {
    color: WHITE, align: 'center',
  });
  rule(s, 3.623, 0.623, 3.044);
});

// 4 - white card left, title right
SLIDES.push((s) => {
  frames(s, [[4.21, 0.958, 1.915, 4.288], [11.498, 0, 1.836, 7.5]]);
  card(s, 1.062, 0.958, 5.062, 5.583);
  clock(s, 1.564, 1.418, BLUE);
  nav(s, 3.453, 6.1415, GRAY);
  body(s, LONG, 1.564, 2.647, 2.145, 2.6);
  menu(s, 7.31, 1.149);
  heading(s, TITLE, 7.143, 2.116, 3.336, 1.313, 36);
  body(s, INTRO, 7.143, 4.843, 3.336, 0.934);
});

// 5 - big photo bottom-right, white text card on top
SLIDES.push((s) => {
  frames(s, [[3.396, 2.292, 9.938, 5.208]]);
  box(s, 5.728, 0, 7.606, 2.292, BLUE);
  card(s, 0.589, 0.51, 7.385, 3.688);
  body(s, LONG, 1.376, 1.647, 5.811, 0.934);
  body(s, clip(21), 1.376, 2.968, 5.811, 0.656);
  nav(s, 1.631, 0.9895, GRAY);
  rule(s, 4.426, 0.992, 2.647);
  heading(s, TITLE, 3.982, 6.429, 5.872, 0.64, 32, WHITE);
  clock(s, 0.293, 6.626, BLUE);
  menu(s, 12.472, 0.51, WHITE);
});

// 6 - title left, blue banner + photo right
SLIDES.push((s) => {
  frames(s, [[5.671, 1.429, 5.99, 4.974]]);
  box(s, 5.671, 0, 7.663, 3.542, BLUE);
  bodyPair(s, INTRO_LONG, INTRO_ALT, 1.178, 3.574, 3.201, 2.322);
  heading(s, TITLE, 1.178, 1.57, 3.336, 1.313, 36);
  clock(s, 11.751, 0.314, WHITE);
  nav(s, 10.656, 7.1095, GRAY);
  menu(s, 12.61, 5.929);
});

// 7 - dark banner + three numbered services
SLIDES.push((s) => {
  frames(s, [[9.118, 0, 4.216, 2.603], [0, 3.759, 2.253, 2.603]]);
  box(s, 0, 0, 9.118, 2.603, INK);
  diamond(s, 3.978, 3.854, 0.73, '01');
  service(s, 3.826, 4.953, { title: 'first service ', body: SVC });
  service(s, 6.804, 4.953, { title: 'second service ', tw: 1.799, sub: SUB2, body: SVC });
  service(s, 9.692, 4.953, { title: 'third service ', tw: 1.799, sub: SUB2, body: SVC, bw: 2.253 });
  diamond(s, 6.955, 3.854, 0.73, '02');
  diamond(s, 9.843, 3.854, 0.73, '03');
  nav(s, 0.531, 7.1095, GRAY);
  menu(s, 0.444, 0.408, WHITE);
  heading(s, TITLE, 1.398, 1.528, 6.126, 0.64, 32, WHITE);
  body(s, clip(11), 1.431, 0.998, 5.559, 0.379, { color: WHITE });
  clock(s, 11.932, 6.824, BLUE);
});

// 8 - portrait photo left, white card with two services right
SLIDES.push((s) => {
  frames(s, [[2.125, 1.438, 4.688, 6.062]]);
  card(s, 5.392, 3.48, 7.385, 3.491);
  service(s, 6.276, 4.984, { title: 'first service ', body: SVC });
  service(s, 9.697, 4.984, { title: 'second service ', tw: 1.799, sub: SUB2, body: SVC });
  s.addText('01', { x: 6.284, y: 4.222, w: 0.66, h: 0.505, fontFace: BODY, fontSize: 24, bold: true, color: INK, valign: 'top' });
  s.addText('02', { x: 9.705, y: 4.222, w: 0.66, h: 0.505, fontFace: BODY, fontSize: 24, bold: true, color: INK, valign: 'top' });
  heading(s, TITLE, 7.518, 1.422, 5.121, 0.572, 28);
  body(s, clip(19), 7.522, 2.067, 5.117, 0.656);
  clock(s, 0.293, 6.626, BLUE);
  menu(s, 0.444, 0.408);
  nav(s, 0.61, 4.8155, GRAY, 270);
});

// 9 - three diamond captions under a white card
SLIDES.push((s) => {
  frames(s, [[0, 0, 3.27, 4.431]]);
  [[5.207, '01', 4.962, 4.51], [8.017, '02', 7.773, 7.32], [10.827, '03', 10.583, 10.13]].forEach(([dx, num, sx]) => {
    diamond(s, dx, 4.794, 0.577, num);
    serviceCaption(s, sx, 5.631);
  });
  heading(s, TITLE, 4.055, 1.33, 2.612, 2.121, 40);
  menu(s, 0.465, 6.569);
  card(s, 7.185, 1.254, 5.733, 2.276);
  clock(s, 11.741, 0.232, BLUE);
  body(s, LONG, 7.887, 1.741, 4.329, 1.212);
  nav(s, 10.705, 7.1095, GRAY);
});

// 10 - blue banner with two white services, photo below
SLIDES.push((s) => {
  frames(s, [[1.103, 0, 1.959, 7.5], [3.509, 3.958, 5.749, 3.542]]);
  box(s, 3.509, 0, 5.749, 3.542, BLUE);
  service(s, 4.255, 1.39, {
    title: 'first service ', body: 'sanisinaia at salanasie atasanit atile nalatauit hanaban',
    color: WHITE, bw: 1.799, bh: 0.934,
  });
  service(s, 6.713, 1.39, {
    title: 'Second service ', tw: 1.799, sub: SUB2,
    body: 'sanisinaia at salanasie atasanit atile nalatauit hanaban', color: WHITE, bw: 1.799, bh: 0.934,
  });
  heading(s, TITLE, 10.03, 3.812, 2.721, 2.121, 40);
  clock(s, 11.741, 0.232, BLUE);
  menu(s, 0.382, 0.408);
  nav(s, 0.547, 6.9435, GRAY, 270);
  s.addText('Sanisinaia sala nasile atasanit atile atasa', {
    x: 4.255, y: 0.419, w: 4.256, h: 0.454, fontFace: BODY, fontSize: 14, bold: true,
    color: WHITE, valign: 'top', lineSpacingMultiple: 1.5,
  });
});

// 11 - 2x2 service grid on a white card
SLIDES.push((s) => {
  frames(s, [[8.444, 1.638, 4.889, 5.862]]);
  box(s, 7.375, 1.638, 1.069, 5.862, BLUE);
  card(s, 0.517, 0.506, 6.858, 4.077);
  [[1.285, 1.078, 'One'], [1.285, 2.72, 'Two'], [4.366, 1.078, 'Three'], [4.366, 2.72, 'Four']]
    .forEach(([x, y, word]) => service(s, x, y, { title: word + ' service ', body: SVC }));
  heading(s, TITLE, 1.328, 5.406, 5.121, 0.572, 28);
  body(s, clip(19), 1.331, 6.051, 5.117, 0.656);
  clock(s, 11.741, 0.232, BLUE);
});

// 12 - mist panel, arrow rule and two diamonds
SLIDES.push((s) => {
  frames(s, [[2.802, 0, 3.062, 3.75]]);
  box(s, 4.333, 0, 9, 7.5, MIST);
  heading(s, TITLE, 6.989, 1.595, 3.907, 1.313, 36);
  service(s, 6.989, 5.063, { title: 'One service ', body: SVC });
  service(s, 10.07, 5.063, { title: 'Three service ', body: SVC });
  diamond(s, 7.117, 4.119, 0.577, '01');
  diamond(s, 10.189, 4.119, 0.577, '02');
  clock(s, 0.293, 6.626, BLUE);
  menu(s, 0.444, 0.408);
  nav(s, 0.61, 4.8155, GRAY, 270);
  s.addShape('line', { x: 4.498, y: 0.683, w: 3.146, h: 0, line: { color: BLUE, width: 1.5, endArrowType: 'triangle' } });
  card(s, 1.853, 2.838, 1.893, 3.475);
  body(s, SVC, 2.246, 3.213, 1.298, 1.133, { size: 10, align: 'left' });
  body(s, 'sanisinaia at cona salanasie atasanit atile nalatauit', 2.246, 4.815, 1.298, 1.133, { size: 10, align: 'left' });
});

// 13 - full-bleed blue panel with white service copy
SLIDES.push((s) => {
  frames(s, [[0, 4.917, 6.146, 2.583], [9.167, 0, 2.938, 3.75]]);
  box(s, 4.958, 0, 8.375, 6.125, BLUE);
  heading(s, TITLE, 0.913, 1.674, 3.363, 2.121, 40);
  menu(s, 0.444, 0.408);
  service(s, 5.979, 1.05, {
    title: 'One service ', color: WHITE, bh: 1.182,
    body: SVC + ' cona salanasie atasanit atile nalatauit nuba',
  });
  body(s, clip(19), 6.787, 4.779, 5.95, 0.656, { color: WHITE, align: 'center' });
  nav(s, 10.705, 7.1095, GRAY);
});

// 14 - wide white title band over a photo
SLIDES.push((s) => {
  frames(s, [[0.625, 1.156, 6.771, 5.188]]);
  box(s, 0, 0, 4.057, 5.188, SILVER);
  card(s, 3.564, 1.67, 7.441, 1.684);
  heading(s, TITLE, 4.057, 2.136, 6.678, 0.707, 36);
  body(s, LONG, 8.804, 4.262, 2.967, 1.767);
  menu(s, 12.569, 0.408);
  nav(s, 10.584, 7.0545, GRAY);
  clock(s, 0.293, 6.647, BLUE);
});

// 15 - blue column left with a white service block
SLIDES.push((s) => {
  frames(s, [[1.521, 3.125, 5.812, 4.375], [4.271, 0, 3.833, 2.333], [8.557, 0, 2.829, 2.333]]);
  box(s, 0, 0, 5.542, 7.5, BLUE);
  heading(s, TITLE, 8.557, 3.436, 3.464, 1.313, 36);
  body(s, clip(19) + ' nasan banan', 8.557, 5.153, 3.464, 1.212);
  service(s, 0.891, 0.785, {
    title: 'One service ', color: WHITE, bh: 0.934,
    body: SVC + ' cona salanasie atasanit',
  });
  menu(s, 12.569, 0.408);
  nav(s, 0.422, 6.9745, WHITE, 270);
});

// 16 - two identical portfolio cards
SLIDES.push((s) => {
  frames(s, [[4.772, 0.958, 1.915, 4.288], [10.481, 0.958, 1.915, 4.288]]);
  [1.624, 7.333].forEach((x) => {
    card(s, x, 0.958, 5.062, 5.583);
    clock(s, x + 0.502, 1.418, BLUE);
    nav(s, x + 2.391, 6.1415, GRAY);
    body(s, LONG, x + 0.502, 2.647, 2.145, 2.6);
  });
  heading(s, PORTFOLIO, -1.523, 3.369, 4.609, 0.707, 36, INK, 'center', 270);
  menu(s, 0.444, 0.408);
  clock(s, 0.293, 6.647, BLUE);
});

// 17 - blue header band, dark service tab
SLIDES.push((s) => {
  frames(s, [[2.312, 0.542, 6.208, 2.542], [8.812, 0.542, 4.521, 2.542]]);
  box(s, 5.604, 0, 7.729, 3.083, BLUE);
  heading(s, TITLE, 0.994, 3.958, 7.047, 0.707, 36);
  menu(s, 0.444, 0.408);
  body(s, clip(22), 4.518, 5.292, 3.315, 1.212);
  body(s, clip(22), 8.812, 5.292, 3.315, 1.212);
  clock(s, 0.293, 6.647, BLUE);
  box(s, 8.812, 3.083, 4.521, 1.271, INK);
  s.addText('One service ', {
    x: 10.223, y: 3.55, w: 1.699, h: 0.337, fontFace: BODY, fontSize: 14, bold: true,
    color: WHITE, align: 'center', valign: 'top',
  });
});

// 18 - blue corner + dark footer panel
SLIDES.push((s) => {
  frames(s, [[5.792, 0, 3.667, 1.625], [5.792, 2.125, 5.208, 1.625], [1, 3.75, 4.271, 3.75]]);
  box(s, 9.458, 0, 3.875, 3.75, BLUE);
  heading(s, TITLE, 1.167, 1.227, 3.464, 1.313, 36);
  box(s, 5.271, 3.75, 8.062, 3.75, INK);
  service(s, 9.964, 0.41, { title: 'One service ', color: WHITE, body: SVC, bh: 0.656 });
  body(s, clip(24) + ' banin', 6.31, 4.762, 2.442, 1.767, { color: WHITE });
  body(s, clip(16), 9.782, 4.762, 2.442, 1.212, { color: WHITE });
  menu(s, 0.444, 0.408);
  clock(s, 11.751, 6.729, WHITE);
});

// 19 - blue card centre, three captions along the bottom
SLIDES.push((s) => {
  frames(s, [[0, 0.708, 3.688, 1.792], [0, 2.938, 3.688, 1.792]]);
  box(s, 3.688, 0.708, 3.5, 4.021, BLUE);
  heading(s, TITLE, 7.967, 1.562, 3.325, 2.121, 40);
  service(s, 4.396, 1.332, {
    title: 'One service ', color: WHITE, dy: 0.711, bh: 1.182,
    body: SVC + ' salanasie atasanit atile nalatauit nuba',
  });
  body(s, 'sanisinaia at cona salanasie atasanit atile nalatauit', 4.396, 3.43, 2.196, 0.656, { color: WHITE });
  clock(s, 11.776, 6.626, BLUE);
  menu(s, 12.611, 0.408);
  nav(s, 12.818, 2.7185, GRAY, 90);
  [3.309, 6.119, 8.929].forEach((x) => serviceCaption(s, x, 6.31));
  body(s, clip(22), 2.871, 5.233, 7.607, 0.656, { align: 'center' });
});

// 20 - "Portfolio One" caption bar
SLIDES.push((s) => {
  frames(s, [[6.832, 0, 4.125, 7.5], [2.067, 0, 4.125, 2.021], [2.067, 4, 4.125, 3.5]]);
  heading(s, PORTFOLIO, -1.523, 3.369, 4.609, 0.707, 36, INK, 'center', 270);
  menu(s, 0.444, 0.408);
  clock(s, 0.293, 6.647, BLUE);
  box(s, 9.165, 2.631, 3.584, 4.303, WHITE, { shadow: dropShadow() });
  body(s, LONG, 9.884, 3.46, 2.145, 2.6);
  box(s, 2.067, 3.01, 4.125, 0.99, INK);
  s.addText('Portfolio One', {
    x: 2.995, y: 3.325, w: 2.271, h: 0.404, fontFace: BODY, fontSize: 18, bold: true,
    color: WHITE, align: 'center', valign: 'top',
  });
});

// 21 - vertical diamond list on the right edge
SLIDES.push((s) => {
  frames(s, [[5.833, 0, 2.771, 4.057], [0, 1.811, 2.487, 5.689], [2.775, 2.956, 2.771, 3.419]]);
  ['01', '02', '03'].forEach((num, i) => {
    const y = 0.74 + i * 2.2475;
    diamond(s, 11.496, y, 0.73, num);
    service(s, 11.345, y + 1.099, { title: 'first service ' });
  });
  heading(s, PORTFOLIO, 0.761, 0.54, 4.609, 0.707, 36);
  clock(s, 6.63, 4.617, INK);
  card(s, 8.295, 1.08, 2.426, 5.583);
  body(s, clip(15), 8.805, 1.592, 1.512, 2.045, { align: 'left' });
  body(s, clip(15), 8.805, 4.057, 1.512, 2.045, { align: 'left' });
});

// 22 - four service labels across the top
SLIDES.push((s) => {
  frames(s, [[0, 1.646, 7.375, 2.104], [7.75, 1.646, 2.579, 2.104], [10.704, 1.646, 2.629, 2.104]]);
  [0.949, 2.907, 4.857, 6.816].forEach((x) => service(s, x, 0.589, { title: 'first service ' }));
  heading(s, TITLE, 0.957, 4.72, 4.343, 1.313, 36);
  nav(s, 10.584, 7.0545, GRAY);
  clock(s, 0.293, 6.647, BLUE);
  menu(s, 12.611, 0.408);
  body(s, clip(24) + ' banin', 6.392, 4.762, 2.442, 1.767);
  body(s, clip(16), 9.864, 4.762, 2.442, 1.212);
});

// 23 - "Handi Jaka" profile, blue + dark panels
SLIDES.push((s) => {
  frames(s, [[6.875, 0, 3.292, 4.229]]);
  box(s, 0, 0, 5.938, 7.5, BLUE);
  box(s, 4.104, 0, 6.062, 6.417, INK);
  heading(s, 'Handi Jaka', 2.52, 2.501, 3.418, 0.707, 36, WHITE, 'right');
  body(s, clip(24) + ' banin', 0.774, 4.121, 2.442, 1.767, { color: WHITE });
  s.addText([
    { text: '13 ', options: { fontSize: 28, color: WHITE, outline: { size: 1.5, color: WHITE } } },
    { text: ': 20', options: { fontSize: 18, color: WHITE } },
  ], {
    x: 3.985, y: 5.306, w: 1.177, h: 0.572, rotate: 90, fontFace: BODY, bold: true,
    color: WHITE, align: 'center', valign: 'top', wrap: false,
  });
  menu(s, 0.444, 0.408, WHITE);
  heading(s, TITLE, 11.055, 1.189, 1.535, 2.794, 32);
  nav(s, 10.584, 7.0545, GRAY);
  service(s, 6.779, 5.651, { title: 'one service ', tsize: 12, tw: 1.195, sw: 1.326, sdy: 0.214, color: WHITE });
  service(s, 8.726, 5.651, { title: 'two service ', tsize: 12, tw: 1.207, sw: 1.326, sdy: 0.214, color: WHITE });
});

// 24 - full blue right panel with rotated service label
SLIDES.push((s) => {
  frames(s, [[3.938, 0, 2.458, 3.312], [0, 4.188, 5.167, 3.312]]);
  box(s, 5.167, 0, 8.167, 7.5, BLUE);
  name(s, 'nani bobo', 7.197, 0.76, 1.765, WHITE);
  body(s, clip(16), 7.197, 1.365, 2.442, 1.212, { color: WHITE });
  clock(s, 11.751, 6.729, WHITE);
  heading(s, TITLE, 8.807, 3.942, 3.533, 1.313, 36, WHITE);
  menu(s, 12.611, 0.408, WHITE);
  s.addText('first service ', {
    x: 5.378, y: 5.675, w: 1.699, h: 0.337, rotate: 90, fontFace: BODY, fontSize: 14, bold: true,
    color: WHITE, align: 'center', valign: 'top',
  });
  s.addText(SUB, {
    x: 5.062, y: 5.696, w: 1.691, h: 0.304, rotate: 90, fontFace: BODY, fontSize: 9,
    color: WHITE, align: 'center', valign: 'top', lineSpacingMultiple: 1.5,
  });
  nav(s, 0.547, 2.8095, GRAY, 270);
  rule(s, 8.961, 5.844, 4.372, WHITE);
});

// 25 - two named bios on the blue panel
SLIDES.push((s) => {
  frames(s, [[6.5, 0.592, 2.712, 2.801], [6.5, 4.065, 2.712, 2.801]]);
  box(s, 5.625, 0, 7.708, 7.5, BLUE);
  [['nani bobo', 1.01, 1.615, 1.765], ['gendra nes', 4.463, 5.067, 2.042]].forEach(([label, ty, by, tw]) => {
    name(s, label, 9.958, ty, tw, WHITE);
    body(s, clip(16), 9.958, by, 2.442, 1.212, { color: WHITE });
  });
  bodyPair(s, INTRO_LONG, INTRO_ALT, 1.178, 3.533, 3.201, 2.322);
  heading(s, TITLE, 1.178, 1.528, 3.336, 1.313, 36);
  menu(s, 0.444, 0.408);
  clock(s, 0.293, 6.647, BLUE);
});

// 26 - two bios on white, wide title right
SLIDES.push((s) => {
  frames(s, [[4.458, 0, 3.25, 3.75], [4.458, 4.125, 3.25, 3.375]]);
  box(s, 6.5, 1.375, 1.708, 0.958, WHITE);
  [['nani bobo', 0.938, 1.542], ['deni lai', 4.583, 5.188]].forEach(([label, ty, by]) => {
    name(s, label, 1, ty, 1.765);
    body(s, clip(16), 1, by, 2.442, 1.212);
  });
  heading(s, TITLE, 6.679, 1.523, 6.53, 0.64, 32);
  bodyPair(s, INTRO_LONG, INTRO_ALT, 9.033, 3.533, 3.201, 2.322);
  nav(s, 10.584, 7.0545, GRAY);
  menu(s, 12.611, 0.408);
  clock(s, 1, 3.399, BLUE, 24, 16);
});

// 27 - clustered column chart
SLIDES.push((s) => {
  frames(s, [[10.349, 0, 2.984, 7.5]]);
  card(s, 0.81, 5.381, 8.444, 1.425);
  s.addChart('bar', [
    { name: 'Series 1', labels: ['1', '2', '3'], values: [4.3, 2.5, 3.5] },
    { name: 'Series 2', labels: ['1', '2', '3'], values: [2.4, 4.4, 1.8] },
    { name: 'Series 3', labels: ['1', '2', '3'], values: [2, 2, 3] },
  ], {
    x: 1, y: 0.67, w: 5.63, h: 4.308,
    barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 219, barOverlapPct: -27,
    chartColors: [INK, BLUE, 'A6A6A6'],
    showLegend: false, showTitle: false,
    catAxisLabelFontFace: BODY, catAxisLabelFontSize: 12, catAxisLabelColor: '595959',
    valAxisLabelFontFace: BODY, valAxisLabelFontSize: 9, valAxisLabelColor: '595959',
    catAxisLineColor: 'D9D9D9', valAxisLineShow: false,
    valGridLine: { color: 'D9D9D9', size: 1 }, catGridLine: { style: 'none' },
  });
  [1.677, 4.487, 7.297].forEach((x) => serviceCaption(s, x, 5.727));
  heading(s, TITLE, 7.636, 1.37, 1.942, 3.13, 36);
});

// 28 - tablet mockup with two numbered services
SLIDES.push((s) => {
  frames(s, [[1.601, 1.609, 3.208, 4.25]]);
  photo(s, 1.382, 1.046, 3.667, 5.409);
  heading(s, TITLE, 8.376, 2.88, 3.107, 1.178, 32);
  clock(s, 11.835, 6.647, BLUE);
  [[1.217, '01', 2.315], [4.796, '02', 5.895]].forEach(([dy, num, ly]) => {
    diamond(s, 6.14, dy, 0.73, num);
    service(s, 5.989, ly, { title: 'first service ' });
  });
  bodyPair(s, INTRO_LONG, INTRO_ALT, 8.376, 4.409, 3.709, 2.045);
  menu(s, 0.444, 0.408);
  card(s, 8.376, 0.569, 4.957, 1.425);
  nav(s, 9.806, 1.2615, GRAY);
});

// 29 - three phone mockups fanned out
SLIDES.push((s) => {
  frames(s, [[7.179, 1.59, 2.208, 4.724], [3.885, 1.59, 2.208, 4.724], [5.567, 1.334, 2.208, 4.724]]);
  photo(s, 3.731, 1.42, 2.528, 5.099, 345);
  photo(s, 7.013, 1.42, 2.528, 5.099, 15);
  heading(s, TITLE, 0.888, 1.495, 2.047, 2.255, 32);
  clock(s, 11.776, 6.626, BLUE);
  menu(s, 12.611, 0.408);
  nav(s, 12.818, 2.7185, GRAY, 90);
  body(s, clip(16) + ' atile atasa', 10.157, 4.434, 1.854, 1.767);
  body(s, clip(16) + ' atile atasa', 0.888, 4.434, 1.854, 1.767);
  photo(s, 5.403, 1.18, 2.528, 5.099);
});

// 30 - closing slide
SLIDES.push((s) => {
  frames(s, [[0, 0, 13.333, 7.5]]);
  box(s, 2.365, 2.125, 8.604, 3.292, WHITE, { fill: { color: WHITE, transparency: 12 } });
  s.addText('Thanks', {
    x: 4.347, y: 2.68, w: 4.639, h: 1.582, fontFace: HEAD, fontSize: 88, bold: true,
    underline: { style: 'sng' }, color: INK, align: 'center', valign: 'top', wrap: false,
  });
  s.addText('Presentation Template', {
    x: 7.66, y: 4.817, w: 3.201, h: 0.417, fontFace: BODY, fontSize: 14, color: GRAY,
    charSpacing: 3, valign: 'top', lineSpacingMultiple: 1.5,
  });
  nav(s, 1.026, 0.7935, WHITE);
  body(s, INTRO + ' atasa', 2.274, 5.808, 4.282, 0.555, { size: 9, color: WHITE, align: 'left' });
  clock(s, 11.751, 0.314, WHITE);
});

/* ------------------------------------------------------------------ build */

const pptx = new PptxGenJS();
pptx.title = 'WorkOut';
pptx.defineLayout({ name: 'WIDE_16x9', width: 13.333333, height: 7.5 });
pptx.layout = 'WIDE_16x9';

SLIDES.forEach((build) => build(pptx.addSlide()));

pptx.writeFile({
  fileName: path.join(__dirname, '0d1753e7-24c2-4b8b-b741-a3450abc6cd0_grok_final.pptx'),
}).then((f) => console.log('wrote', f));
