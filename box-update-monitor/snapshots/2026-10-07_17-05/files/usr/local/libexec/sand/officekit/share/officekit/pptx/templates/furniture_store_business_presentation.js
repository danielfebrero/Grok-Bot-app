#!/usr/bin/env node
/**
 * "Bútorüzlet" — Furniture Business Presentation Template
 * 20 slides, 13.333in x 7.5in, rebuilt with pptxgenjs only.
 *
 * Photographs / logos / icons of the original deck are not embedded; each one is
 * drawn as a flat colour placeholder at the same position and size.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ---------------------------------------------------------------- palette -- */
const GREEN = '456934'; // brand green (backgrounds, panels, rules)
const DARK = '1B2613'; // headline green-black
const GRAY = '7C7D7A'; // body copy
const GRAY2 = '7B7C75'; // body copy (infographic slides)
const PALE = 'DCE3D7'; // pale green, used on green panels
const LIME = '9EAC43';
const OLIVE = '6E8959';
const LEAF = '8FC34D';
const DEEP = '12381E';
const CREAM = 'FEFFFA';
const PRICE = '20220D';
const WHITE = 'FFFFFF';
const PHOTO = 'DBDBDB'; // placeholder tone for photographs
const PHOTO_DK = '494949'; // placeholder tone for the dark phone screen

/* ------------------------------------------------------------------ fonts -- */
const SERIF = 'Playfair Display';
const SERIF_MED = 'Playfair Display Medium';
const SERIF_BLK = 'Playfair Display Black';
const SANS = 'Source Sans Pro';
const LATO = 'Lato';

/** text-box insets of the source deck (pptxgenjs order: left, right, bottom, top) */
const INSET = [7.2, 7.2, 3.6, 3.6];

const NONE = { type: 'none' };

/* ---------------------------------------------------------------- helpers -- */
const txt = (s, content, o) =>
  s.addText(
    content,
    Object.assign(
      { fontFace: SANS, fontSize: 12, color: GRAY, align: 'left', valign: 'top', margin: INSET, wrap: true },
      o
    )
  );

/** headline: 36pt bold italic serif */
const heading = (s, content, x, y, w, h, o) =>
  txt(s, content, Object.assign({ x, y, w, h, fontFace: SERIF, fontSize: 36, bold: true, italic: true, color: DARK }, o));

/** body copy: 12pt sans, 150% leading */
const body = (s, content, x, y, w, h, o) =>
  txt(s, content, Object.assign({ x, y, w, h, lineSpacingMultiple: 1.5 }, o));

const rect = (s, x, y, w, h, color, o) =>
  s.addShape('rect', Object.assign({ x, y, w, h, fill: { color }, line: NONE }, o));

const oval = (s, x, y, w, h, color) => s.addShape('ellipse', { x, y, w, h, fill: { color }, line: NONE });

const ring = (s, x, y, d, color) =>
  s.addShape('ellipse', { x, y, w: d, h: d, fill: NONE, line: { color, width: 1 } });

/** pill shaped bar (roundRect with adj = 50%) */
const pill = (s, x, y, w, h, color) =>
  s.addShape('roundRect', { x, y, w, h, fill: { color }, line: NONE, rectRadius: Math.min(w, h) / 2 });

const hLine = (s, x, y, w, color, o) =>
  s.addShape('line', Object.assign({ x, y, w, h: 0, line: { color, width: 0.75 } }, o));

const vLine = (s, x, y, h, color) => s.addShape('line', { x, y, w: 0, h, line: { color, width: 0.75 } });

/** placeholder for a photograph — same box as the bitmap it replaces */
const photo = (s, x, y, w, h, o) => {
  const p = o || {};
  s.addShape(p.shape || 'rect', { x, y, w, h, fill: { color: p.color || PHOTO }, line: NONE });
};

/** placeholder for a pictogram / icon — a small glyph-sized block inside the icon box */
const icon = (s, x, y, size, color) => {
  const g = size * 0.55;
  s.addShape('roundRect', { x: x + (size - g) / 2, y: y + (size - g) / 2, w: g, h: g, fill: { color }, line: NONE, rectRadius: g * 0.2 });
};

/* ------------------------------------------------- shared "website" chrome -- */
const NAV = [
  { label: 'Home', x: 6.776, w: 0.698 },
  { label: 'Pricing ', x: 8.131, w: 0.812 },
  { label: 'Products', x: 9.891, w: 0.942 },
  { label: 'About Us', x: 11.748, w: 0.945 }
];

function navBar(s) {
  NAV.forEach((n) => txt(s, n.label, { x: n.x, y: 0.472, w: n.w, h: 0.337, fontSize: 14, color: WHITE }));
  // caret next to "Products"
  s.addShape('triangle', { x: 10.905, y: 0.605, w: 0.185, h: 0.072, fill: { color: WHITE }, line: NONE, rotate: 180 });
}

function logo(s) {
  // picnic-table mark drawn from primitives (placeholder for the logo bitmap):
  // orange spark, white table top, two splayed legs, green bench bar
  [-32, 0, 32].forEach((a) => rect(s, 0.994, 0.3, 0.018, 0.15, 'E38C41', { rotate: a }));
  rect(s, 0.826, 0.538, 0.348, 0.052, WHITE);
  rect(s, 0.895, 0.59, 0.05, 0.26, WHITE, { rotate: 20 });
  rect(s, 1.055, 0.59, 0.05, 0.26, WHITE, { rotate: -20 });
  rect(s, 0.714, 0.683, 0.469, 0.05, LEAF);
  txt(s, 'Bútorüzlet', { x: 1.201, y: 0.487, w: 1.338, h: 0.404, fontFace: SERIF, fontSize: 18, italic: true, color: WHITE });
}

/* ------------------------------------------------------------ slide bodies -- */

// 1 — cover
function slide01(pptx) {
  const s = pptx.addSlide();
  s.background = { color: GREEN };
  photo(s, 6.667, 0, 6.667, 7.5);
  rect(s, 6.667, 1.278, 2.052, 4.945, GREEN);
  navBar(s);
  logo(s);
  hLine(s, 1.857, 6.204, 4.791, WHITE);
  hLine(s, 0, 1.303, 4.791, WHITE);
  photo(s, 4.821, 1.562, 3.691, 4.376);
  txt(s, 'Bútorüzlet', {
    x: 0.709, y: 2.81, w: 6.067, h: 1.447,
    fontFace: SERIF_MED, fontSize: 80, bold: true, italic: true, color: WHITE
  });
  txt(s, 'Furniture Business Presentation Template', { x: 0.709, y: 4.354, w: 5.34, h: 0.337, fontSize: 14, color: WHITE });
}

// 2 — "Hello, Welcome To Bútorüzlet"
function slide02(pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  rect(s, 0.6, 0, 2.498, 7.5, GREEN); // panel inherited from the layout

  txt(s, [
    { text: 'Hello, ', options: { breakLine: true } },
    { text: 'Welcome To Bútorüzlet' }
  ], { x: 7.1, y: 1.472, w: 5.633, h: 1.313, fontFace: SERIF_MED, fontSize: 36, bold: true, italic: true, color: DARK });

  body(s,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do loreiopui ipsim eiusmodialo tempor ' +
    'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco ' +
    'laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit ' +
    'esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa ' +
    'qui officia deserunt mollit anim id est laborum. Lorem ipsum dolor sit amet, consectetur adipiscing.a',
    7.1, 2.934, 5.389, 2.19);

  // "Learn More" button
  s.addShape('roundRect', { x: 7.1, y: 5.386, w: 2.938, h: 0.642, fill: NONE, line: { color: LIME, width: 1 }, rectRadius: 0.321 });
  txt(s, 'Learn More', { x: 7.498, y: 5.522, w: 1.681, h: 0.37, fontFace: SERIF_MED, fontSize: 16 });
  s.addShape('line', { x: 9.26, y: 5.707, w: 0.33, h: 0, line: { color: GRAY, width: 1.25, endArrowType: 'triangle' } });

  hLine(s, 4.169, 5.784, 2.5, GREEN);
  hLine(s, 4.167, 0.36, 2.5, GREEN);
  vLine(s, 6.914, 0.6, 2.35, GREEN);
  vLine(s, 6.914, 3.151, 2.35, GREEN);
  hLine(s, 2.995, 1.703, 0.984, GREEN);
  hLine(s, 2.995, 6.976, 0.984, GREEN);

  [[1.482, 1.905], [1.482, 4.453], [4.169, 3.148], [4.169, 0.6]].forEach(([x, y]) => photo(s, x, y, 2.498, 2.353));
}

// 3 — quote
function slide03(pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  photo(s, 0, 0, 13.333, 7.5);
  rect(s, 2.773, -1.556, 7.787, 4.313, GREEN);
  [2.38, 6.273, 10.166].forEach((x) => oval(s, x, 2.363, 0.787, 0.787, WHITE));
  [[2.38, 2.756], [6.273, 2.779], [10.166, 2.751]].forEach(([x, y]) => hLine(s, x, y, 0.787, GREEN));
  [[2.773, 2.363], [6.667, 2.363], [10.56, 2.357]].forEach(([x, y]) => vLine(s, x, y, 0.787, GREEN));
  txt(s, '\u201C Simplicity carried to an extreme becomes elegance.\u201D \u2013 Unknown', {
    x: 2.865, y: 0.316, w: 7.602, h: 1.616,
    fontSize: 32, bold: true, italic: true, color: WHITE, align: 'center', lineSpacingMultiple: 1.5
  });
}

// 4 — "Best Furniture For People who love their home"
function slide04(pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  rect(s, 0.6, 1.328, 12.733, 2.93, GREEN); // band inherited from the layout

  heading(s, 'Best Furniture For People who love their home', 1.141, 1.558, 6.679, 1.313, { fontFace: SERIF_MED, color: WHITE });
  body(s,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do loreiopui ipsim lorem eiusmodialo tempor ' +
    'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ',
    1.141, 3.039, 6.079, 0.978, { color: WHITE });

  hLine(s, 7.945, 6.666, 4.791, GREEN);
  hLine(s, 7.942, 0.395, 4.791, GREEN);
  vLine(s, 6.915, 4.454, 3.051, GREEN);

  photo(s, 0.6, 4.453, 6.079, 3.047);
  [0.602, 2.125, 3.648, 5.171].forEach((y) => photo(s, 7.945, y, 4.788, 1.326));
}

// 5 — "Why You Should Choose Us?"
function slide05(pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  rect(s, 8.25, 0, 4.483, 7.5, GREEN);

  heading(s, 'Why You Should Choose Us?', 0.6, 1.563, 5.666, 1.313);
  body(s,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do loreiopui ipsim lorem eiusmodialo tempor ' +
    'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ',
    0.6, 3.138, 5.851, 0.978);

  const PERKS = [
    ['Pleasant Staff', 0.6, 4.565, 2.074], ['Very Fast Delivery', 4.138, 4.565, 2.59],
    ['Best quality', 0.6, 5.194, 1.846], ['Convenient location', 4.138, 5.194, 2.798],
    ['Individual Design', 0.6, 5.823, 2.537], ['Non-Standard Offers', 4.138, 5.823, 2.919]
  ];
  PERKS.forEach(([label, x, y, w]) =>
    txt(s, label, { x, y, w, h: 0.404, fontFace: SERIF_MED, fontSize: 18, color: DARK, bullet: { indent: 22.5 } }));

  hLine(s, 6.677, 4.065, 2.079, GREEN);
  vLine(s, 6.482, 0.6, 3.272, GREEN);
  photo(s, 6.667, 0.6, 2.083, 3.275);
  photo(s, 8.966, 1.874, 3.052, 3.753);
}

// 6 — "For People Who Love Their Home"
function slide06(pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  rect(s, 0.6, 0, 2.498, 1.406, GREEN);

  heading(s, 'For People Who Love Their Home', 6.975, 1.588, 5.633, 1.313);
  body(s,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do loreiopui ipsim eiusmodialo tempor ' +
    'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam,',
    6.975, 3.05, 5.389, 0.978);

  [['Comfortable and Clean', 4.637, 3.172], ['Spacious and Strategic', 5.266, 3.084], ['Sturdy Build', 5.895, 1.957]]
    .forEach(([label, y, w]) =>
      txt(s, label, { x: 8.965, y, w, h: 0.404, fontFace: SERIF_MED, fontSize: 18, color: DARK, bullet: { indent: 22.5 } }));

  hLine(s, 3.371, 4.221, 2.5, GREEN);
  hLine(s, 3.371, 1.529, 2.5, GREEN);
  vLine(s, 6.059, 1.701, 2.35, GREEN);
  vLine(s, 8.564, 4.364, 2.35, GREEN);
  vLine(s, 3.294, 4.364, 2.35, GREEN);

  [[0.6, 1.674], [0.6, 4.302], [3.373, 1.674], [5.871, 4.302]].forEach(([x, y]) => photo(s, x, y, 2.498, 2.353));
}

// 7 — "Lighting Tips" + progress bars
function slide07(pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  rect(s, 8.294, 3.047, 2.498, 1.406, GREEN);

  heading(s, 'Lighting Tips Interior design for a better home', 0.6, 0.6, 5.806, 1.313);
  body(s,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do loreiopui ipsimip eiusmodialo tempor ' +
    'incididunt ut labore et dolore magna aliqua. Ut enim adispi minim veniam, quis nostrud exercitation ullamco ' +
    'laboris nisi ut aliquip ex ealore commodo consequat. Duis aute irure dolor in reprehenderit in voluptate ' +
    'velilore esse cillum dolore.',
    0.6, 2.185, 5.806, 1.584);

  const BARS = [
    { label: 'Design interior', labelY: 4.062, labelW: 3.038, y: 4.855, done: 4.083 },
    { label: 'Furniture Design', labelY: 5.524, labelW: 4.365, y: 6.247, done: 3.278 }
  ];
  BARS.forEach((b) => {
    txt(s, b.label, { x: 0.6, y: b.labelY, w: b.labelW, h: 0.504, fontFace: SERIF_MED, fontSize: 18, color: DARK, lineSpacingMultiple: 1.5 });
    pill(s, 0.6, b.y, 5.129, 0.307, PALE);
    pill(s, 0.6, b.y, b.done, 0.307, OLIVE);
  });

  vLine(s, 10.976, 0.635, 2.409, GREEN);
  hLine(s, 6.668, 5.784, 2.5, GREEN);
  hLine(s, 6.669, 3.087, 2.5, GREEN);
  vLine(s, 9.505, 3.741, 3.752, GREEN);
  hLine(s, 10.233, 3.554, 2.5, GREEN);

  photo(s, 6.671, 3.264, 2.498, 2.353);
  photo(s, 9.681, 3.741, 3.052, 3.753);
  photo(s, 8.291, 0.634, 2.498, 2.407);
}

// 8 — "Minimalist room specifications"
function slide08(pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  rect(s, 4.337, 3.75, 8.996, 3.75, GREEN);

  heading(s, 'Minimalist room specifications', 7.251, 0.953, 5.633, 1.313);
  body(s,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do loreiopui ipsim eiusmodialo tempor ' +
    'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam,',
    7.251, 2.415, 5.389, 0.978);

  const FEATURES = [
    { title: 'Simple design', tw: 1.792, y: 4.463 },
    { title: 'Good quality', tw: 1.661, y: 5.887 }
  ];
  FEATURES.forEach((f) => {
    rect(s, 4.681, f.y, 0.701, 0.708, WHITE);
    icon(s, 4.732, f.y + 0.055, 0.598, DARK);
    txt(s, f.title, { x: 5.682, y: f.y + 0.003, w: f.tw, h: 0.404, fontFace: SERIF_MED, fontSize: 18, color: WHITE });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.', { x: 5.682, y: f.y + 0.505, w: 2.274, h: 0.505, color: WHITE });
  });

  vLine(s, 3.884, 2.948, 4.559, GREEN);
  hLine(s, 0.6, 2.767, 3.059, GREEN);
  vLine(s, 6.869, 0.6, 3.15, GREEN);
  hLine(s, 4.338, 0.414, 2.331, GREEN);

  photo(s, 0.6, 2.947, 3.052, 4.555);
  photo(s, 4.338, 0.6, 2.329, 3.146);
  photo(s, 8.074, 4.505, 4.66, 2.353);
}

// 9 — "Our Catalogue"
function slide09(pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  vLine(s, 8.767, 0.6, 6.3, LIME); // vertical dividers from the layout
  vLine(s, 4.567, 0.6, 6.3, LIME);

  const CARDS = [
    { x: 0.871, capY: 3.442, imgY: 3.75, imgH: 3.75, tagX: 0.876, tagY: 6.45, priceX: 1.749, priceY: 6.531, priceW: 0.751, price: '$225' },
    { x: 4.954, capY: 2.135, imgY: 2.442, imgH: 5.058, tagX: 4.951, tagY: 6.3, priceX: 5.829, priceY: 6.381, priceW: 0.742, price: '$235' },
    { x: 9.037, capY: 1.038, imgY: 1.346, imgH: 5.554, tagX: 9.037, tagY: 6.0, priceX: 9.909, priceY: 6.075, priceW: 0.754, price: '$245' }
  ];
  CARDS.forEach((c) => rect(s, c.x, c.capY, 3.426, 0.308, GREEN));
  CARDS.forEach((c) => photo(s, c.x, c.imgY, 3.426, c.imgH));

  heading(s, 'Our Catalogue', 0.6, 0.6, 3.977, 0.707);
  body(s,
    'Lorem ipsum dolor sit amet, consectetur adipiscingioi elit, sed do loreiopui ipsimip eiusmodialo tempor ' +
    'lorei incididunt ut labore et dolore magna aliqua. Ut enim il adispiscing.',
    0.6, 1.65, 3.977, 1.281);

  CARDS.forEach((c) => {
    rect(s, c.tagX, c.tagY, 2.498, 0.6, CREAM);
    txt(s, c.price, { x: c.priceX, y: c.priceY, w: c.priceW, h: 0.438, fontFace: SERIF_MED, fontSize: 20, color: PRICE, align: 'center' });
  });
}

// 10 — designer profile
function slide10(pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  photo(s, 0, 0, 3.654, 7.5);
  photo(s, 3.815, 0, 2.596, 7.5);

  heading(s, 'Melvin Alexandria', 7.34, 1.202, 4.745, 0.707);
  body(s, 'DESIGNER FURNITURE', 7.34, 2.018, 3.063, 0.417, { fontSize: 14 });
  body(s,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed dolorei loreiopui ipsimip eiusmodialo tempor ' +
    'incididunt ut labore etimpsui dolore magna aliqua. Ut enim adispi minim veniam, quis nostrudloi exercitation.',
    7.34, 2.817, 4.882, 1.281);

  [{ y: 4.521, rule: 5.099, pct: '88%', pctX: 11.655, pctW: 0.582 },
   { y: 5.716, rule: 6.298, pct: '90%', pctX: 11.659, pctW: 0.589 }].forEach((r) => {
    hLine(s, 7.482, r.rule, 4.74, GREEN);
    txt(s, 'Lorem Ipsum Dolor', { x: 7.482, y: r.y, w: 2.059, h: 0.337, fontSize: 14, bold: true, color: DARK });
    txt(s, r.pct, { x: r.pctX, y: r.y, w: r.pctW, h: 0.337, fontFace: SERIF, fontSize: 14, bold: true, italic: true, color: DARK });
  });

  rect(s, 6.131, 0, 0.559, 4.542, GREEN);
}

// 11 — "Amazing Team Work"
function slide11(pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };

  heading(s, 'Amazing Team Work', 0.629, 1.397, 5.7, 0.707);
  body(s,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do loreiopui ipsimip eiusmodialo tempor ' +
    'incididunt ut labore et dolore magna aliqua. Ut enim adispi minim veniam, quis nostrud exercitation ' +
    'ullamco laboris nisi ut aliquip ex ealore commodo consequat.',
    0.629, 2.447, 6.067, 1.281);

  // six framed portraits: rule above, hairline to the right
  const TEAM = [[0.6, 4.547], [3.802, 4.547], [7.004, 4.547], [10.207, 4.547], [10.207, 1.397], [7.004, 1.397]];
  TEAM.forEach(([x, y]) => {
    hLine(s, x + 0.029, y - 0.233, 2.5, GREEN);
    vLine(s, x + 2.717, y - 0.06, 2.35, GREEN);
  });
  TEAM.forEach(([x, y]) => photo(s, x, y, 2.498, 2.353));
}

// 12 — break slide
function slide12(pptx) {
  const s = pptx.addSlide();
  s.background = { color: GREEN };
  hLine(s, 0, 1.303, 10.591, WHITE);
  hLine(s, 1.857, 6.204, 5.915, WHITE);
  photo(s, 7.773, 0, 5.561, 7.5);
  navBar(s);
  logo(s);
  txt(s, 'Break Slide ', {
    x: 0.709, y: 2.81, w: 6.067, h: 1.447,
    fontFace: SERIF_MED, fontSize: 80, bold: true, italic: true, color: WHITE
  });
  txt(s, 'Furniture Business Presentation Template', { x: 0.709, y: 4.354, w: 5.34, h: 0.337, fontSize: 14, color: WHITE });
}

// 13 — six services
function slide13(pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  heading(s, 'Get Our Great Service', 3.769, 0.6, 5.796, 0.707, { align: 'center' });
  body(s,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et ' +
    'dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ' +
    'ex ea commodo consequat',
    1.478, 1.539, 10.378, 0.675, { align: 'center' });

  const SERVICES = [
    { name: 'Service One', ix: 2.095, iy: 2.397, ty: 3.393, by: 3.767, tx: 1.624, bx: 1.169 },
    { name: 'Service Two', ix: 6.273, iy: 2.378, ty: 3.374, by: 3.748, tx: 5.801, bx: 5.347 },
    { name: 'Service Three', ix: 10.451, iy: 2.411, ty: 3.407, by: 3.781, tx: 9.979, bx: 9.525 },
    { name: 'Service Four', ix: 2.095, iy: 4.822, ty: 5.818, by: 6.191, tx: 1.624, bx: 1.169 },
    { name: 'Service Five', ix: 6.273, iy: 4.802, ty: 5.799, by: 6.172, tx: 5.801, bx: 5.347 },
    { name: 'Service Six', ix: 10.451, iy: 4.835, ty: 5.832, by: 6.205, tx: 9.979, bx: 9.525 }
  ];
  SERVICES.forEach((v) => {
    icon(s, v.ix, v.iy, 0.787, GREEN);
    txt(s, v.name, { x: v.tx, y: v.ty, w: 1.731, h: 0.404, fontSize: 18, bold: true, color: DARK, align: 'center' });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit', v.bx, v.by, 2.64, 0.675, { align: 'center' });
  });
}

// 14 — "Our Best Product" (3 x 3 photo grid)
function slide14(pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  rect(s, 0, 0, 6.286, 7.5, GREEN);

  heading(s, 'Our Best Product', 0.6, 0.6, 4.331, 0.707, { color: WHITE });
  const para1 =
    'Lorem ipsum dolor sit amet, consectetur adipiscing elitei, sed do eiusmod tempor incididunt ut labore et ' +
    'dolorelore magna aliqua. Ut enim ad minim veniam, quis nostrudiui exercitation ullamco laboris nisi ut ' +
    'aliquip ex ealoremipu commodo consequat. Duis aute irure dolor inloremipsum reprehenderit in voluptate ' +
    'velit esse cillum dolore eudolr fugiat nulla pariatur. Excepteur sint occaecat cupidatatio non proident, ' +
    'sunt in culpa qui officia deserunt mollit lori anim id est laborum. ';
  const para2 =
    'Lorem ipsum dolor sit amet, consectetur adipiscing elitei, sed do eiusmod tempor incididunt ut labore et ' +
    'dolorelore magna aliqua. Ut enim ad minim veniam, quis nostrudiui exercitation ullamco laboris nisi ut ' +
    'aliquip ex ealoremipu commodo consequat. Duis aute irure dolor inloremipsum reprehenderit in voluptate ' +
    'velit esse cillum dolore eudolr fugiat nulla pariatur.';
  txt(s, [{ text: para1, options: { breakLine: true } }, { text: para2 }],
    { x: 0.6, y: 1.587, w: 4.21, h: 4.916, color: PALE, lineSpacingMultiple: 1.5 });

  [-0.453, 2.047, 4.547].forEach((y) => [5.034, 7.635, 10.235].forEach((x) => photo(s, x, y, 2.498, 2.353)));
}

// 15 — pricing tables
function slide15(pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  vLine(s, 8.767, 1.625, 5.275, GREEN);
  vLine(s, 4.567, 1.625, 5.275, GREEN);

  const PLANS = [
    { name: 'Basic Package', nameW: 1.778, price: '$200', card: 0.871, ring: 1.702, disc: 1.881, nameX: 1.655, priceX: 1.501, monthX: 2.877, listX: 1.518 },
    { name: 'Standard Package', nameW: 2.23, price: '$300', card: 4.959, ring: 5.785, disc: 5.969, nameX: 5.594, priceX: 5.441, monthX: 6.816, listX: 5.457 },
    { name: 'Premium Package', nameW: 2.22, price: '$400', card: 9.048, ring: 9.879, disc: 10.058, nameX: 9.698, priceX: 9.545, monthX: 10.92, listX: 9.562 }
  ];
  const FEATURES = ['3GB Disk Space', 'Free SSL & CDN', 'Unlimited Projects', '300 Images'];

  PLANS.forEach((p) => rect(s, p.card, 2.587, 3.426, 4.313, GREEN));
  PLANS.forEach((p) => ring(s, p.ring, 1.706, 1.763, GREEN));

  PLANS.forEach((p) => {
    txt(s, p.name, { x: p.nameX, y: 3.466, w: p.nameW, h: 0.404, fontSize: 18, bold: true, color: WHITE });
    txt(s, p.price, { x: p.priceX, y: 3.874, w: 1.565, h: 0.897, fontFace: SERIF_BLK, fontSize: 36, bold: true, color: WHITE, align: 'center', lineSpacingMultiple: 1.5 });
    txt(s, '/month', { x: p.monthX, y: 4.458, w: 0.789, h: 0.364, fontFace: LATO, color: PALE, lineSpacingMultiple: 1.5 });
    // breakLine on every item so pptxgenjs emits one bulleted paragraph per feature
    txt(s, FEATURES.map((f) => ({ text: f, options: { breakLine: true, bullet: { characterCode: '27A2', indent: 13.5 } } })),
      { x: p.listX, y: 4.998, w: 2.13, h: 1.651, color: PALE, lineSpacingMultiple: 2 });
  });

  heading(s, 'Pricing Tables section', 3.769, 0.6, 5.796, 0.707, { align: 'center' });
  PLANS.forEach((p) => photo(s, p.disc, 1.885, 1.406, 1.406, { shape: 'ellipse' }));
}

// 16 — app mock-up
function slide16(pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  rect(s, 7.047, 0, 6.286, 7.5, GREEN);

  // two phone mock-ups replacing the mockup bitmaps: white body, screen, notch
  const phone = (x, y, w, h, screen) => {
    s.addShape('roundRect', { x, y, w, h, fill: { color: 'FAFAFA' }, line: { color: 'EDEDED', width: 0.75 }, rectRadius: 0.3 });
    s.addShape('roundRect', { x: x + 0.119, y: y + 0.12, w: w - 0.238, h: h - 0.24, fill: { color: screen }, line: NONE, rectRadius: 0.22 });
    s.addShape('roundRect', { x: x + w / 2 - 0.42, y: y + 0.12, w: 0.84, h: 0.2, fill: { color: 'FAFAFA' }, line: NONE, rectRadius: 0.09 });
    s.addShape('roundRect', { x: x + w / 2 - 0.13, y: y + 0.19, w: 0.26, h: 0.035, fill: { color: 'D5D5D5' }, line: NONE, rectRadius: 0.017 });
  };
  phone(2.999, 0.62, 2.533, 5.364, PHOTO_DK);
  phone(1.49, 1.516, 2.533, 5.349, PHOTO);

  heading(s, "Let's Join the exclusive Bútorüzlet community", 7.486, 0.74, 5.248, 1.919, { color: WHITE });
  body(s,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elitei, sed dolor siteu eiusmod tempor incididunt ut ' +
    'labore et dolorelore magna aliqua. Utioloi enim ad minim veniam, quis nostrudiui exercitation ullamco ' +
    'laboris nisi ut aliquip ex ealoremipu commodo consequat. Duis aute irure dolorsitie inloremipsum ' +
    'reprehenderit in voluptate velit esse cillum dolore eudolr fugiat nulla pariatur. ',
    7.486, 2.853, 5.248, 1.887, { color: PALE });

  [7.489, 9.191, 10.894].forEach((x) => ring(s, x, 5.017, 1.763, PALE));
  vLine(s, 5.981, 1.022, 4.614, DEEP);
  vLine(s, 1.3, 1.988, 4.614, DEEP);
  hLine(s, 1.859, 1.233, 0.984, GREEN);
  hLine(s, 4.253, 6.278, 0.984, GREEN);
  [7.667, 9.37, 11.072].forEach((x) => photo(s, x, 5.196, 1.406, 1.406, { shape: 'ellipse' }));
}

// 17 — infographic: stacked rings
function slide17(pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  heading(s, 'Infographic Mockups', 3.769, 0.6, 5.796, 0.707, { align: 'center' });

  // stack of four rings (placeholder for the 3-D artwork), drawn back to front.
  // Each ring = extruded outer wall + light top face + shaded bore + open hole.
  const WALL = 0.208;
  for (let i = 0; i < 4; i++) {
    const outerY = 1.739 + i * 0.963;
    const holeY = 2.236 + i * 0.963;
    oval(s, 1.972, outerY + WALL, 3.389, 1.944, '456934');
    oval(s, 1.972, outerY, 3.389, 1.944, '749A61');
    oval(s, 2.447, holeY + WALL, 2.44, 0.95, '2B4320');
    oval(s, 2.447, holeY, 2.44, 0.95, WHITE);
  }

  const STEPS = [
    { name: 'Step One', tx: 8.509, y: 1.894, accent: GRAY, dot: '.' },
    { name: 'Step Two', tx: 8.53, y: 3.221, accent: GREEN, dot: '.' },
    { name: 'Step Three', tx: 8.53, y: 4.548, accent: GRAY, dot: '.' },
    { name: 'Step Four', tx: 8.53, y: 5.875, accent: GREEN, dot: '.' }
  ];
  STEPS.forEach((v) => {
    oval(s, 7.283, v.y, 0.899, 0.899, v.accent);
    hLine(s, 6.058, v.y + 0.449, 1.009, v.accent, { line: { color: v.accent, width: 0.75, beginArrowType: 'oval', endArrowType: 'oval' } });
    icon(s, 7.339, v.y + 0.057, 0.787, WHITE);
    txt(s, v.name, { x: v.tx, y: v.y - 0.075, w: 2.528, h: 0.404, fontSize: 18, bold: true, color: DARK });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed dolored lorem ipsum sed dolor' + v.dot,
      v.tx, v.y + 0.299, 3.856, 0.675);
  });
}

// 18 — infographic: hexagon steps
function slide18(pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  heading(s, 'Infographic Mockups', 3.769, 0.6, 5.796, 0.707, { align: 'center' });

  const NODES = [
    { ix: 4.346, iy: 3.296, title: 'Step One', tx: 0.899, ty: 3.109, bx: 0.742, by: 3.482, align: 'right' },
    { ix: 5.711, iy: 4.961, title: 'Step Two', tx: 1.705, ty: 4.956, bx: 1.549, by: 5.33, align: 'right' },
    { ix: 7.076, iy: 2.414, title: 'Step Three', tx: 9.34, ty: 2.185, bx: 9.34, by: 2.559, align: 'left' },
    { ix: 8.441, iy: 4.176, title: 'Step Four', tx: 10.115, ty: 4.032, bx: 10.115, by: 4.406, align: 'left' }
  ];
  NODES.forEach((n) => {
    const cx = n.ix + 0.315;
    const cy = n.iy + 0.315;
    // isometric block: extruded dark side, lighter top face, white icon disc
    s.addShape('hexagon', { x: cx - 1.104, y: cy - 0.183, w: 2.208, h: 1.085, fill: { color: '344C29' }, line: NONE });
    s.addShape('hexagon', { x: cx - 1.104, y: cy - 0.493, w: 2.208, h: 1.085, fill: { color: '769A64' }, line: NONE });
    oval(s, cx - 0.443, cy - 0.373, 0.82, 0.82, WHITE);
    icon(s, n.ix, n.iy, 0.63, DARK);
    txt(s, n.title, { x: n.tx, y: n.ty, w: 2.32, h: 0.404, fontSize: 18, bold: true, color: DARK, align: n.align });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed dolored lorem ipsum sed dolor.',
      n.bx, n.by, 2.477, 0.978, { color: GRAY2, align: n.align });
  });
}

// 19 — contact
function slide19(pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  rect(s, 0.6, 3.395, 6.067, 4.105, GREEN);

  heading(s, 'Get In Touch With Us', 5.621, 1.697, 4.268, 1.313);

  const CONTACTS = [
    { title: 'Phone', tw: 0.942, y: 3.561, accent: GREEN, lines: [{ text: '133-347888-30', options: { fontFace: LATO } }] },
    { title: 'Email', tw: 0.898, y: 4.827, accent: GRAY, lines: [
      { text: 'furniture' }, { text: 'bútorüzlet ', options: { italic: true } }, { text: '@gmail.you ' }] },
    { title: 'Opening Time', tw: 1.871, y: 6.092, accent: GREEN, lines: [{ text: '09.00 AM – 20.00 PM' }] }
  ];
  CONTACTS.forEach((c) => {
    oval(s, 8.147, c.y, 0.802, 0.802, c.accent);
    icon(s, 8.323, c.y + 0.177, 0.449, WHITE);
    txt(s, c.title, { x: 9.273, y: c.y - 0.048, w: c.tw, h: 0.404, fontFace: SERIF_BLK, fontSize: 18, bold: true, color: DARK });
    txt(s, c.lines, { x: 9.273, y: c.y + 0.416, w: 2.58, h: 0.372, lineSpacingMultiple: 1.5 });
  });

  hLine(s, 1.921, 1.037, 3.429, GREEN);
  hLine(s, 3.821, 4.314, 2.5, GREEN);
  vLine(s, 6.509, 4.487, 2.35, GREEN);
  vLine(s, 9.889, 0, 2.35, GREEN);
  hLine(s, 10.164, 2.568, 2.5, GREEN);

  photo(s, 1.921, 1.364, 3.426, 5.536);
  photo(s, 10.164, 0, 2.498, 2.353);
}

// 20 — thank you
function slide20(pptx) {
  const s = pptx.addSlide();
  s.background = { color: GREEN };
  photo(s, 0, 0, 13.333, 7.5);
  s.addShape('rect', { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: GREEN, transparency: 6 }, line: NONE });
  hLine(s, 0, 6.105, 10.59, WHITE);
  txt(s, 'Thank You!', {
    x: 4.523, y: 3.498, w: 6.067, h: 1.447,
    fontFace: SERIF_MED, fontSize: 80, bold: true, italic: true, color: WHITE
  });
  txt(s, 'Furniture Business Presentation Template', { x: 4.523, y: 5.042, w: 5.34, h: 0.337, fontSize: 14, color: WHITE });
  navBar(s);
  logo(s);
}

/* -------------------------------------------------------------------- build */
const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'W16x9', width: 13.333, height: 7.5 });
  pptx.layout = 'W16x9';
  pptx.title = 'Bútorüzlet — Furniture Business Presentation Template';
  BUILDERS.forEach((fn) => fn(pptx));
  return pptx;
}

build()
  .writeFile({ fileName: path.join(__dirname, '187cad69-6aa3-4a30-a11a-214cba867bdd_grok_final.pptx') })
  .then((f) => console.log('wrote', f))
  .catch((e) => {
    console.error(e);
    process.exit(1);
  });
