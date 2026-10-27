/**
 * Santana - Fashion Presentation Template
 * Standalone pptxgenjs rebuild of the 36-slide reference deck (13.333 x 7.5 in).
 * Raster photos in the original are redrawn as labelled placeholder rectangles.
 *
 *   node <thisfile>.js   ->   writes the .pptx next to this file
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const GOLD = 'A28E6A';      // accent3 - the deck's signature tan
const CREAM = 'E9E5DC';     // lt2 - soft panel background
const WHITE = 'FFFFFF';
const INK = '000000';
const SLATE = '3A405A';
const NEAR_INK = '0C0C0C';
const GREY = '3F3F3F';
const DEEP = '463E2C';      // dark brown used by the infographic pictograms
const UMBER = '524733';     // slightly warmer variant on slide 31
const BAND = 'EEEEEE';      // timeline background strip
const RULE = '7F7F7F';      // hairline rules

/* -------------------------------------------------------------------- fonts */
const TITLE_F = 'Work Sans Medium';
const BODY_F = 'Poppins';
const MED_F = 'Poppins Medium';
const SEMI_F = 'Poppins SemiBold';
const OS_F = 'Open Sans SemiBold';

/* --------------------------------------------------- recurring lorem blocks */
const L = {
  a: "There are many variations of passages of Lorem Ipsum available, but the majority have suffered " +
     "alteration in some form, by injected humour, or randomised words which don't look even slightly believable.",
  b: "If you are going to use a passage of Lorem Ipsum, you need to be sure there isn't anything " +
     "embarrassing hidden in the middle of text.",
  c: "All the Lorem Ipsum generators on the Internet tend to repeat predefined chunks as necessary, " +
     "making this the first true generator on the Internet.",
  d: "It uses a dictionary of over 200 Latin words, combined with a handful of model sentence structures, " +
     "to generate Lorem Ipsum which looks reasonable.",
  e: "The generated Lorem Ipsum is therefore always free from repetition, injected humour, " +
     "or non-characteristic words etc.",
  f: "Lorep  ipsum duis aute irure dolor in kauselih oiluek epreh deriti vols esse cill inure " +
     "dolorlaboru sit amet. Duis autelo irusitakus rephenderi Voluptate",
};

/* ------------------------------------------------------------------ helpers */

/** Flat colour rectangle (panels, columns, strips). */
function band(s, x, y, w, h, color) {
  s.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: color }, line: { type: 'none' } });
}

/** Text box. Defaults match the deck: top anchored, no autofit, 0.1" insets. */
function text(s, body, x, y, w, h, opt) {
  opt = opt || {};
  s.addText(body, {
    x: x, y: y, w: w, h: h,
    align: opt.align || 'left',
    valign: opt.valign || 'top',
    fontFace: opt.fontFace || BODY_F,
    fontSize: opt.fontSize || 10,
    color: opt.color || INK,
    bold: opt.bold || false,
    italic: opt.italic || false,
    rotate: opt.rotate,
    lineSpacingMultiple: opt.lineSpacing,
    margin: opt.margin || [7.2, 7.2, 3.6, 3.6],   // [l, r, b, t] in points, as in the source deck
    wrap: true,
  });
}

/** The "+" ornament of the cover slides (0.556" box, 20% arm thickness). */
function plus(s, x, y, color) {
  const size = 0.556, arm = 0.111;
  band(s, x, y + (size - arm) / 2, size, arm, color);
  band(s, x + (size - arm) / 2, y, arm, size, color);
}

/** Fully rounded "pill" tag. */
function pill(s, x, y, w, h, color) {
  s.addShape('roundRect', { x: x, y: y, w: w, h: h, rectRadius: h / 2, fill: { color: color }, line: { type: 'none' } });
}

/** Placeholder standing in for a pictogram of the original deck.
 *  Drawn at 64% of the slot so its ink mass matches a typical glyph. */
function icon(s, x, y, w, h, color) {
  const k = 0.64;
  s.addShape('roundRect', {
    x: x + w * (1 - k) / 2, y: y + h * (1 - k) / 2, w: w * k, h: h * k,
    rectRadius: Math.min(w, h) * k * 0.25,
    fill: { color: color }, line: { type: 'none' },
  });
}

/** Placeholder standing in for a raster photo of the original deck. */
function photo(s, x, y, w, h) {
  s.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: 'C8C8C8' }, line: { color: 'A9A29A', width: 0.75 } });
  s.addText('[image]', {
    x: x, y: y, w: w, h: h, align: 'center', valign: 'middle',
    fontFace: BODY_F, fontSize: 11, color: '5A5A5A',
  });
}

/** Hairline horizontal rule. */
function rule(s, x, y, w, color) {
  s.addShape('line', { x: x, y: y, w: w, h: 0, line: { color: color, width: 0.75 } });
}

/** Vertical dashed leader with a dot at the end pointing away from the strip. */
function dashLine(s, cx, y, h, dotAtTop) {
  s.addShape('line', { x: cx, y: y, w: 0, h: h, line: { color: DEEP, width: 1.5, dashType: 'dash' } });
  const dy = dotAtTop ? y : y + h;
  s.addShape('ellipse', { x: cx - 0.048, y: dy - 0.048, w: 0.097, h: 0.097, fill: { color: DEEP }, line: { type: 'none' } });
}

/** Pointy-top hexagon badge (infographic cards). */
function hexagon(s, x, y, w, h, color) {
  s.addShape('hexagon', {
    x: x + (w - h) / 2, y: y + (h - w) / 2, w: h, h: w, rotate: 90,
    fill: { color: color }, line: { color: WHITE, width: 2.9 },
  });
}

/** Quarter-circle pie wedge; the wedge sweeps 90.4 deg from due west. */
function wedge(s, x, y, size, rotate, color) {
  s.addShape('pie', {
    x: x, y: y, w: size, h: size, rotate: rotate,
    angleRange: [179.6, 270], fill: { color: color }, line: { type: 'none' },
  });
}

/** One quadrant of a split circle. quad: 0=UL 1=UR 2=LR 3=LL, pulled apart by a gap. */
function quarter(s, cx, cy, r, quad, color) {
  const gap = 0.075;
  const dx = (quad === 0 || quad === 3) ? -gap : gap;
  const dy = (quad === 0 || quad === 1) ? -gap : gap;
  const range = [[180, 270], [270, 360], [0, 90], [90, 180]][quad];
  s.addShape('pie', {
    x: cx + dx - r, y: cy + dy - r, w: 2 * r, h: 2 * r,
    angleRange: range, fill: { color: color }, line: { type: 'none' },
  });
}

/** Ring node: thick coloured disc with a lighter core. */
function donut(s, cx, cy, rOuter, rInner, outer, inner) {
  s.addShape('ellipse', { x: cx - rOuter, y: cy - rOuter, w: 2 * rOuter, h: 2 * rOuter, fill: { color: outer }, line: { type: 'none' } });
  s.addShape('ellipse', { x: cx - rInner, y: cy - rInner, w: 2 * rInner, h: 2 * rInner, fill: { color: inner }, line: { type: 'none' } });
}

/** Slanted "Text Here" tab of the timeline strip (bottom edge pushed right). */
function chevron(s, x, y, w, h, color) {
  s.addShape('parallelogram', {
    x: x, y: y, w: w, h: h, rectRadius: 0.337, flipV: true,
    fill: { color: color }, line: { type: 'none' },
  });
}

/** Right-hand sidebar shared by the cover, break and closing slides. */
function sidebar(s) {
  band(s, 12.467, 0, 0.866, 1.488, CREAM);
  band(s, 12.467, 1.488, 0.866, 6.012, GOLD);
  icon(s, 12.662, 0.497, 0.439, 0.438, INK);                       // rocket mark
  text(s, 'PREMIUM DESIGN', 12.249, 2.927, 1.302, 0.202,
    { align: 'center', fontSize: 6, color: WHITE, rotate: -90 });
  text(s, 'Santana', 12.327, 4.585, 1.154, 0.337,
    { align: 'center', fontSize: 14, fontFace: SEMI_F, color: WHITE, rotate: -90 });
  for (let i = 0; i < 3; i++) band(s, 12.784, 6.957 + i * 0.106, 0.229, 0.046, WHITE);   // hamburger mark
}

/** Cover / break / closing layout: big title, kicker, paragraph and two plus marks. */
function coverSlide(s, title, titleBox, plusB, ink) {
  sidebar(s);
  text(s, title, titleBox[0], titleBox[1], titleBox[2], 1.582,
    { align: 'center', fontSize: 88, fontFace: MED_F, color: ink });
  text(s, 'Fashion Presentation Template', 0.978, 4.154, 4.247, 0.37,
    { fontSize: 16, fontFace: MED_F, color: ink });
  text(s, L.f, 0.229, 6.455, 5.581, 0.582,
    { align: 'justify', fontSize: 10, color: ink, lineSpacing: 1.5 });
  plus(s, 0.641, 2.102, ink);
  plus(s, plusB[0], plusB[1], ink);
}

/* ------------------------------------------------------------------- slides */

function slide01(s) {   // Cover
  coverSlide(s, 'Santana', [0.809, 2.526, 5.866], [6.203, 4.014], INK);
}

function slide02(s) {   // Welcome To Santana
  band(s, 4.262, 1.333, 1.5, 4.873, GOLD);
  text(s, L.a + ' ' + L.b, 6.457, 3.483, 5.491, 1.339, { align: 'justify', fontSize: 10, fontFace: BODY_F, color: SLATE, lineSpacing: 1.5 });
  text(s, 'It uses a dictionary of over 200 Latin words, combined with a handful of model sentence structures, to generate Lorem Ipsum which looks reasonable the generated Lorem Ipsum.', 6.457, 5.319, 5.491, 0.834, { align: 'justify', fontSize: 10, fontFace: BODY_F, color: SLATE, lineSpacing: 1.5 });
  text(s, 'Welcome To Santana', 6.457, 1.166, 5.116, 1.717, { fontSize: 48, fontFace: TITLE_F });
}

function slide03(s) {   // About Of Santana
  band(s, 6.549, 1.18, 1.687, 4.873, GOLD);
  text(s, L.a + ' ' + L.b, 0.913, 3.415, 4.975, 1.592, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, L.d + ' ' + 'The generated Lorem Ipsum is therefore always free from repetition, injected humour, or non-characteristic', 0.913, 5.188, 4.956, 1.087, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'About Of Santana', 0.913, 0.811, 5.116, 1.717, { fontSize: 48, fontFace: TITLE_F });
  text(s, 'About Us', 0.913, 2.988, 4.247, 0.37, { fontSize: 16, fontFace: TITLE_F });
}

function slide04(s) {   // History Santana
  band(s, 2.5, 2.19, 3.092, 5.31, GOLD);
  text(s, L.d + ' ' + 'The generated Lorem Ipsum is therefore always free from repetition, injected humour, or non-characteristic', 7.568, 3.419, 4.449, 1.339, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour, or randomised words which don\'t look even slightly.', 7.568, 5.032, 4.449, 1.087, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'History Santana', 7.568, 0.655, 3.795, 1.717, { fontSize: 48, fontFace: TITLE_F });
  text(s, 'History of Us', 7.568, 2.986, 4.247, 0.37, { fontSize: 16, fontFace: TITLE_F });
}

function slide05(s) {   // Santana Collaboration
  band(s, 10.286, 0, 3.048, 7.5, GOLD);
  band(s, 0.563, 5.091, 6.703, 2.107, CREAM);
  text(s, 'Santana Collaboration', 0.625, 1.032, 5.355, 1.717, { fontSize: 48, fontFace: TITLE_F });
  text(s, L.b + ' ' + 'All the Lorem Ipsum generators on the Internet tend to repeat predefined chunks as necessary.', 0.625, 3.619, 5.355, 1.136, { align: 'justify', fontSize: 10.5, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'Santana Collaboration', 0.625, 3.225, 4.247, 0.37, { fontSize: 16, fontFace: TITLE_F });
  text(s, L.b + ' ' + 'All the Lorem Ipsum generators on the Internet.', 1.123, 5.939, 5.337, 0.834, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'Style Of Santana', 1.123, 5.498, 2.183, 0.37, { fontSize: 16, fontFace: TITLE_F });
}

function slide06(s) {   // Product Quality
  band(s, 0, 0.556, 3.048, 6.389, GOLD);
  text(s, 'Product Quality', 6.667, 0.596, 5.461, 1.717, { fontSize: 48, fontFace: TITLE_F });
  text(s, L.a + ' ' + L.b, 6.667, 3.461, 5.461, 1.339, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'Product Santana', 6.667, 3.027, 4.247, 0.37, { fontSize: 16, fontFace: TITLE_F });
  text(s, L.d + ' ' + 'The generated Lorem Ipsum is therefore always free from repetition.', 6.667, 5.063, 5.461, 1.087, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
}

function slide07(s) {   // Why Choose Us?
  band(s, 0, 1.127, 1.683, 4.429, GOLD);
  band(s, 8.571, 0, 4.762, 5.171, CREAM);
  text(s, 'Why Choose Us?', 2.345, 0.445, 5.155, 1.717, { fontSize: 48, fontFace: TITLE_F });
  text(s, L.b + ' ' + L.c, 2.345, 2.884, 5.258, 1.087, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'Choose Us ', 2.345, 2.49, 4.247, 0.37, { fontSize: 16, fontFace: TITLE_F });
  text(s, L.d + ' ' + 'The generated Lorem Ipsum is therefore always free from repetition, injected humour, or non-characteristic', 2.345, 4.188, 5.258, 1.087, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour, or randomised words which don\'t look even slightly.', 2.345, 6.08, 5.163, 1.148, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'Prodcut Santana', 2.345, 5.71, 2.16, 0.37, { fontSize: 16, fontFace: TITLE_F });
}

function slide08(s) {   // Who We Are
  band(s, 7.746, 0, 4.698, 7.5, CREAM);
  band(s, 0, 5.651, 0.692, 1.849, GOLD);
  text(s, 'Who We Are', 0.787, 0.923, 4.92, 0.909, { fontSize: 48, fontFace: TITLE_F });
  text(s, L.a + ' ' + L.b, 0.723, 2.928, 4.247, 1.931, { align: 'justify', fontSize: 10.5, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'About Santana', 0.723, 2.486, 4.247, 0.37, { fontSize: 16, fontFace: TITLE_F });
  text(s, L.c + ' ', 0.876, 5.883, 4.247, 1.133, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'New Collor Fashion', 0.876, 5.508, 2.65, 0.37, { fontSize: 16, fontFace: TITLE_F });
}

function slide09(s) {   // Vision & Mission
  band(s, 0, 0, 5.671, 5, CREAM);
  band(s, 2.751, 5.971, 2.556, 1.529, GOLD);
  text(s, 'VISION & MISSION', 5.912, 0.64, 7.102, 0.909, { fontSize: 48, fontFace: TITLE_F });
  text(s, L.a + ' ', 6.017, 1.833, 5.298, 0.834, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, L.b + ' ' + 'All the Lorem Ipsum generators on the Internet tend to repeat predefined chunks as necessary.', 6.536, 3.856, 5.254, 1.087, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'Vision 01', 6.536, 3.338, 1.632, 0.467, { fontSize: 16, fontFace: TITLE_F, lineSpacing: 1.5 });
  text(s, 'Mission 01', 6.536, 5.41, 1.632, 0.46, { fontSize: 16, fontFace: TITLE_F, lineSpacing: 1.5 });
  text(s, L.d + ' ' + 'The generated Lorem Ipsum is therefore always free from repetition, injected humour.', 6.536, 5.852, 5.298, 1.087, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
}

function slide10(s) {   // New Team Fashion
  band(s, 0, 4.129, 13.333, 3.371, CREAM);
  band(s, 0, 0, 0.893, 0.7, GOLD);
  band(s, 12.44, 2.886, 0.893, 0.7, GOLD);
  text(s, 'New Team Fashion', 1.521, 0.755, 7.257, 0.909, { fontSize: 48, fontFace: TITLE_F });
  text(s, L.f + ' ' + L.f + '.', 1.521, 2.341, 7.828, 0.834, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'New Team', 1.521, 1.947, 2.117, 0.37, { fontSize: 16, fontFace: TITLE_F });
}

function slide11(s) {   // Best Team
  band(s, 0, 4.029, 13.333, 3.471, CREAM);
  band(s, 12.714, 0, 0.619, 2.657, GOLD);
  text(s, L.f + ' ' + 'loremil. ipsum duis aute irure dolor in kauselih oiluek Lorep  ipsum duis aute irure', 2.126, 1.528, 8.832, 0.606, { align: 'center', fontSize: 10.5, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'Best Team', 4.705, 0.418, 4.283, 0.909, { fontSize: 48, fontFace: TITLE_F });
  text(s, [
      { text: 'Cleveland, ' },
      { text: '03/14/1990', options: { fontFace: OS_F } },
    ], 5.571, 6.56, 1.942, 0.344, { align: 'center', fontSize: 10.5, fontFace: SEMI_F, italic: true, lineSpacing: 1.5 });
  text(s, 'Safira Clara', 5.302, 6.224, 2.48, 0.404, { align: 'center', fontSize: 18, fontFace: MED_F });
  text(s, [
      { text: 'Minnesota, ' },
      { text: '03/14/1990', options: { fontFace: OS_F } },
    ], 8.759, 6.56, 1.942, 0.344, { align: 'center', fontSize: 10.5, fontFace: SEMI_F, italic: true, lineSpacing: 1.5 });
  text(s, 'Anastasya Kristi', 8.49, 6.224, 2.48, 0.404, { align: 'center', fontSize: 18, fontFace: MED_F });
  text(s, [
      { text: 'Oklahama, ' },
      { text: '03/14/1990', options: { fontFace: OS_F } },
    ], 2.254, 6.56, 1.942, 0.344, { align: 'center', fontSize: 10.5, fontFace: SEMI_F, italic: true, lineSpacing: 1.5 });
  text(s, 'Winddy Rara', 1.985, 6.224, 2.48, 0.404, { align: 'center', fontSize: 18, fontFace: MED_F });
}

function slide12(s) {   // New Team
  band(s, 4.414, 0, 8.919, 3.086, CREAM);
  band(s, 1.538, 6.357, 2.719, 1.143, GOLD);
  text(s, L.f + ' ' + 'loremil. ipsum duis aute irure dolor in kauselih oiluek Lorep  ipsum duis aute irure', 5.416, 4.543, 5.921, 0.834, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, L.f + ' ' + 'loremil. ', 5.416, 5.775, 5.921, 0.582, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, L.f + ' ' + 'loremil.', 5.416, 1.395, 5.921, 0.582, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'New Team', 5.414, 0.423, 4.054, 0.909, { fontSize: 48, fontFace: TITLE_F });
  text(s, [
      { text: 'Boston, ' },
      { text: '03/14/1990', options: { fontFace: OS_F } },
    ], 5.416, 4.047, 1.942, 0.344, { fontSize: 10.5, fontFace: SEMI_F, italic: true, lineSpacing: 1.5 });
  text(s, 'Winddy Rara', 5.416, 3.71, 2.48, 0.404, { fontSize: 18, fontFace: MED_F });
  text(s, [
      { text: 'New Jersey, ' },
      { text: '03/14/1990', options: { fontFace: OS_F } },
    ], 5.416, 2.588, 1.942, 0.344, { fontSize: 10.5, fontFace: SEMI_F, italic: true, lineSpacing: 1.5 });
  text(s, 'Dandi Fazza', 5.416, 2.252, 2.48, 0.404, { fontSize: 18, fontFace: MED_F });
}

function slide13(s) {   // New Founder
  band(s, 0, 5.243, 6.5, 1.6, CREAM);
  band(s, 6.614, 5.243, 3.614, 1.6, GOLD);
  text(s, L.a, 0.838, 2.459, 5.508, 0.834, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'New Founder', 0.838, 0.604, 5.056, 0.909, { fontSize: 48, fontFace: TITLE_F });
  text(s, [
      { text: 'Indiana, ' },
      { text: '03/14/1990', options: { fontFace: OS_F } },
    ], 7.324, 6.081, 1.942, 0.344, { align: 'center', fontSize: 10.5, fontFace: SEMI_F, color: WHITE, italic: true, lineSpacing: 1.5 });
  text(s, 'Fabio Kross ', 7.055, 5.744, 2.48, 0.404, { align: 'center', fontSize: 18, fontFace: MED_F, color: WHITE });
  text(s, L.b + ' ' + L.c, 0.838, 3.582, 5.508, 1.087, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'Founder', 0.838, 2.025, 2.117, 0.37, { fontSize: 16, fontFace: TITLE_F });
  text(s, L.a + ' ', 0.663, 5.744, 5.43, 0.834, { align: 'justify', fontSize: 10, fontFace: BODY_F, color: NEAR_INK, lineSpacing: 1.5 });
  text(s, 'Models', 0.663, 5.374, 2.117, 0.37, { fontSize: 16, fontFace: TITLE_F });
}

function slide14(s) {   // New Service
  band(s, 0, 2.841, 5.357, 4.675, CREAM);
  band(s, 12.214, 6.329, 1.119, 1.171, GOLD);
  text(s, 'There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form.', 6.961, 6.225, 4.986, 0.606, { align: 'justify', fontSize: 10.5, fontFace: BODY_F, color: NEAR_INK, lineSpacing: 1.5 });
  text(s, 'Fashion 01', 6.961, 5.754, 1.745, 0.505, { fontSize: 16, fontFace: MED_F, lineSpacing: 1.5 });
  text(s, L.c, 6.961, 4.502, 4.986, 0.834, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'Fashion 01', 6.961, 4.031, 1.745, 0.505, { fontSize: 16, fontFace: MED_F, lineSpacing: 1.5 });
  icon(s, 6.049, 5.95, 0.592, 0.591, GOLD);
  icon(s, 6.172, 4.227, 0.345, 0.591, GOLD);
  text(s, 'New Service', 0.344, 3.901, 4.749, 0.909, { fontSize: 48, fontFace: TITLE_F });
  text(s, L.d + ' ' + L.e, 0.344, 5.189, 4.288, 1.339, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
}

function slide15(s) {   // Best Fashion Service
  band(s, 10.186, 3.143, 2.543, 4.373, CREAM);
  text(s, L.b, 2.131, 4.49, 5.074, 0.582, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'New Service', 2.131, 3.988, 2.355, 0.518, { fontSize: 18, fontFace: MED_F, lineSpacing: 1.5 });
  icon(s, 1.183, 4.416, 0.478, 0.445, GOLD);
  text(s, 'New Service', 2.131, 5.526, 2.355, 0.518, { fontSize: 18, fontFace: MED_F, lineSpacing: 1.5 });
  icon(s, 1.133, 5.771, 0.531, 0.531, GOLD);
  text(s, L.c, 2.131, 6.106, 5.074, 0.834, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'Best Fashion Service', 0.63, 0.671, 5.014, 1.717, { fontSize: 48, fontFace: TITLE_F });
  text(s, L.b + ' ' + L.c, 0.63, 2.586, 5.456, 1.087, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
}

function slide16(s) {   // Our Service
  band(s, 0, 3.043, 7.6, 4.473, CREAM);
  icon(s, 0.769, 4.286, 0.446, 0.507, GOLD);
  icon(s, 0.769, 5.859, 0.525, 0.382, GOLD);
  text(s, 'All the Lorem Ipsum generators on the Internet tend to repeat predefined chunks as necessary, making this the first true generator on the Internet', 1.526, 4.286, 4.501, 0.834, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, L.b, 1.526, 5.928, 4.501, 0.834, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'New Service', 1.526, 5.388, 1.784, 0.471, { fontSize: 16, fontFace: MED_F, lineSpacing: 1.5 });
  text(s, 'New Service', 1.526, 3.815, 1.784, 0.471, { fontSize: 16, fontFace: MED_F, lineSpacing: 1.5 });
  text(s, 'Our Service', 0.769, 0.539, 5.014, 0.909, { fontSize: 48, fontFace: TITLE_F });
  text(s, L.d + ' ' + L.e, 0.769, 1.683, 5.014, 1.087, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
}

function slide17(s) {   // New Service (gold panel)
  band(s, 7.342, 0, 5.991, 6.794, GOLD);
  band(s, 7.342, 6.794, 5.991, 0.722, CREAM);
  text(s, 'New Service', 7.944, 0.458, 4.897, 0.909, { fontSize: 48, fontFace: TITLE_F, color: WHITE });
  text(s, [
      { text: L.d + ' ' + 'The generated Lorem Ipsum is therefore always free from repetition, injected humour, or non-characteristic words etc' },
      { text: '.', options: { fontSize: 10.5 } },
    ], 7.944, 1.728, 4.738, 1.351, { align: 'justify', fontSize: 10, fontFace: BODY_F, color: WHITE, lineSpacing: 1.5 });
  text(s, L.c, 8.795, 3.563, 3.835, 0.858, { align: 'justify', fontSize: 10, fontFace: BODY_F, color: WHITE, lineSpacing: 1.5 });
  text(s, '01', 7.742, 3.563, 0.89, 0.934, { fontSize: 36, fontFace: SEMI_F, color: WHITE, lineSpacing: 1.5 });
  text(s, L.d + ' ' + 'PLACEHOLDER', 8.795, 4.896, 3.887, 1.363, { align: 'justify', fontSize: 10, fontFace: BODY_F, color: WHITE, lineSpacing: 1.5 });
  text(s, '02', 7.742, 4.896, 0.89, 0.934, { fontSize: 36, fontFace: SEMI_F, color: WHITE, lineSpacing: 1.5 });
}

function slide18(s) {   // Break slide
  coverSlide(s, 'Breakslide', [0.866, 2.594, 7.063], [7.617, 4.097], WHITE);
}

function slide19(s) {   // Our Gallery
  band(s, 6.587, 0.27, 6.385, 3.952, GOLD);
  text(s, L.a + ' ' + 'If you are going to use a passage of Lorem Ipsum, ', 7.874, 1.961, 4.613, 1.339, { align: 'justify', fontSize: 10, fontFace: BODY_F, color: WHITE, lineSpacing: 1.5 });
  text(s, 'Our Gallery', 7.874, 0.655, 4.516, 0.909, { fontSize: 48, fontFace: TITLE_F, color: WHITE });
  text(s, L.b + ' ' + L.e, 0.532, 5.549, 3.587, 1.616, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'New Collor Fashion', 0.532, 5.084, 3.034, 0.37, { fontSize: 16, fontFace: TITLE_F });
}

function slide20(s) {   // Our Portfolio
  band(s, 0, 0.964, 6.365, 5.706, GOLD);
  text(s, L.a + ' ' + L.b, 0.704, 3.626, 4.875, 1.592, { align: 'justify', fontSize: 10, fontFace: BODY_F, color: WHITE, lineSpacing: 1.5 });
  text(s, 'Our Portfolio', 0.704, 1.698, 5.06, 0.909, { fontSize: 48, fontFace: TITLE_F, color: WHITE });
  text(s, 'New Portfolio', 0.704, 3.212, 2.117, 0.37, { fontSize: 16, fontFace: TITLE_F, color: WHITE });
}

function slide21(s) {   // Our Portfolio
  band(s, 0, 0.317, 0.556, 6.921, GOLD);
  text(s, 'Our Portfolio ', 1.104, 0.437, 3.562, 1.717, { fontSize: 48, fontFace: TITLE_F });
  text(s, L.b + ' ' + L.e, 1.104, 2.939, 4.04, 1.339, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'New Portfolio', 1.104, 2.525, 2.117, 0.37, { fontSize: 16, fontFace: TITLE_F });
  text(s, 'There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour, or randomised words which don\'t look even slightly believable', 1.104, 5.477, 4.04, 1.087, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'About Gallery', 1.104, 5.09, 2.117, 0.37, { fontSize: 16, fontFace: TITLE_F });
}

function slide22(s) {   // Gallery Collection
  band(s, 9.524, 0, 3.81, 7.516, CREAM);
  text(s, L.a + ' ' + 'If you are going to use a passage of Lorem Ipsum, you need to be sure there isn’t  embarrassing hidden in the middle of text.', 1.246, 5.998, 6.662, 1.087, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'About Gallery', 1.246, 5.589, 2.117, 0.37, { fontSize: 16, fontFace: TITLE_F });
  text(s, 'Gallery Collection ', 1.957, 0.316, 9.419, 0.909, { align: 'center', fontSize: 48, fontFace: TITLE_F });
}

function slide23(s) {   // New Portfolio Workspace
  band(s, 8.048, 3.508, 5.286, 4.008, CREAM);
  text(s, 'New Portfolio Workspace', 6.667, 0.691, 5.667, 1.717, { fontSize: 48, fontFace: TITLE_F });
  text(s, L.d + ' ' + L.e, 8.453, 4.623, 4.393, 1.339, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'About Workspace', 8.453, 4.204, 2.363, 0.37, { fontSize: 16, fontFace: TITLE_F });
  text(s, 'If you are going to use a passage of Lorem Ipsum, you need to be sure there isn\'t anything embarrassing hidden in the middle of text', 8.453, 6.272, 4.393, 0.834, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
}

function slide24(s) {   // New Collection
  band(s, 6.635, 0.397, 6.698, 3.476, GOLD);
  text(s, 'New Collection', 7.12, 0.887, 6.023, 0.909, { fontSize: 48, fontFace: TITLE_F, color: WHITE });
  text(s, 'Lorep  ipsum duis aute irure dolor in kauselih oiluek epreh deriti vols esse cill inure dolorlaboru sit amet. Duis autelo irusitaku.', 7.12, 2.585, 4.921, 0.582, { align: 'justify', fontSize: 10, fontFace: BODY_F, color: WHITE, lineSpacing: 1.5 });
  text(s, 'About Gallery', 7.12, 2.166, 2.117, 0.37, { fontSize: 16, fontFace: TITLE_F, color: WHITE });
  text(s, L.a + ' ' + L.b, 9.026, 4.769, 3.889, 1.844, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, '2024', 0.419, 5.146, 1.871, 0.841, { align: 'center', fontSize: 44, fontFace: SEMI_F });
  text(s, L.e, 0.195, 6.02, 2.305, 0.89, { align: 'center', fontSize: 8, fontFace: BODY_F, lineSpacing: 1.5 });
}

function slide25(s) {   // New Mockup Dekstop
  band(s, 6.04, 0, 7.294, 7.5, CREAM);
  photo(s, 1.19, 2.023, 4.815, 4.213);
  text(s, L.c + ' ' + 'Then It uses a dictionary of over 200 Latin words another, combined with a handful of model sentence structures, to generate Lorem Ipsum which looks reasonable. The generated Lorem Ipsum is therefore always free from repetition, injected humour, or non-characteristic words etc.', 6.779, 4.129, 5.364, 1.592, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'Mockup Dekstop', 6.779, 3.75, 2.117, 0.37, { fontSize: 16, fontFace: TITLE_F });
  text(s, 'New Mockup Dekstop', 6.779, 1.294, 5.23, 1.717, { fontSize: 48, fontFace: TITLE_F });
}

function slide26(s) {   // New Mockup Laptop
  band(s, 6.047, 5.134, 7.286, 2.387, GOLD);
  photo(s, 1.002, 2.219, 5.572, 3.225);
  text(s, L.a + ' ' + L.b, 6.779, 3.687, 5.364, 1.339, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'Mockup Laptop', 6.779, 3.207, 2.117, 0.37, { fontSize: 16, fontFace: TITLE_F });
  text(s, 'New Mockup Laptop', 6.759, 1.057, 5.364, 1.717, { fontSize: 48, fontFace: TITLE_F });
}

function slide27(s) {   // New Mockup Tablet
  band(s, 0, 1.794, 4.079, 5.727, GOLD);
  photo(s, 0.749, 1.022, 4.964, 5.353);
  text(s, L.b + ' ' + L.e, 6.489, 4.209, 5.364, 1.087, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'Mockup Tablet', 6.451, 3.839, 2.117, 0.37, { fontSize: 16, fontFace: TITLE_F });
  text(s, 'New Mockup Tablet', 6.489, 1.357, 5.364, 1.717, { fontSize: 48, fontFace: TITLE_F });
}

function slide28(s) {   // New Mockup Phone
  band(s, 2.143, 4.868, 11.19, 2.652, GOLD);
  photo(s, 1.836, 0.979, 2.813, 5.542);
  text(s, L.a + ' ' + 'If you are going to use a passage of Lorem Ipsum.', 6.176, 3.457, 5.509, 1.087, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'Mockup Phone', 6.176, 3.086, 2.117, 0.37, { fontSize: 16, fontFace: TITLE_F });
  text(s, 'New Mockup Phone', 6.176, 0.979, 5.509, 1.717, { fontSize: 48, fontFace: TITLE_F });
}

function slide29(s) {   // Infographic - 4 hex cards
  text(s, 'Infographic', 4.475, 0.725, 4.383, 0.841, { align: 'center', fontSize: 44, fontFace: TITLE_F });
  band(s, 4.597, 2.883, 2.056, 3.091, GOLD);
  hexagon(s, 5.105, 2.272, 1.024, 1.132, GOLD);
  icon(s, 5.389, 2.598, 0.481, 0.479, WHITE);
  band(s, 8.805, 2.883, 2.056, 3.091, GOLD);
  hexagon(s, 9.327, 2.272, 1.024, 1.132, GOLD);
  icon(s, 9.6, 2.627, 0.479, 0.421, WHITE);
  band(s, 2.468, 2.883, 2.068, 3.091, CREAM);
  hexagon(s, 2.97, 2.272, 1.024, 1.132, CREAM);
  icon(s, 3.272, 2.598, 0.421, 0.479, INK);
  band(s, 6.696, 2.883, 2.056, 3.091, CREAM);
  hexagon(s, 7.218, 2.272, 1.024, 1.132, CREAM);
  icon(s, 7.491, 2.598, 0.485, 0.479, INK);
  text(s, 'Lorem ipsum dolor sit amet. Et fugiat exercitationem qui amet quas et dolore suscipit. ', 4.711, 4.52, 1.827, 0.707, { align: 'center', fontSize: 9, fontFace: BODY_F, color: WHITE });
  text(s, 'Lorem ipsum dolor sit amet. Et fugiat exercitationem qui amet quas et dolore suscipit. ', 6.834, 4.52, 1.827, 0.707, { align: 'center', fontSize: 9, fontFace: BODY_F });
  text(s, 'Lorem ipsum dolor sit amet. Et fugiat exercitationem qui amet quas et dolore suscipit. ', 8.92, 4.52, 1.827, 0.707, { align: 'center', fontSize: 9, fontFace: BODY_F, color: WHITE });
  text(s, 'Lorem ipsum dolor sit amet. Et fugiat exercitationem qui amet quas et dolore suscipit. ', 2.589, 4.52, 1.827, 0.707, { align: 'center', fontSize: 9, fontFace: BODY_F });
  text(s, '01', 2.468, 3.772, 2.068, 0.578, { align: 'center', fontSize: 40, fontFace: MED_F, valign: 'middle', margin: [7.2, 7.2, 7.2, 7.2] });
  text(s, '02', 4.597, 3.772, 2.056, 0.578, { align: 'center', fontSize: 40, fontFace: MED_F, color: WHITE, valign: 'middle', margin: [7.2, 7.2, 7.2, 7.2] });
  text(s, '03', 6.696, 3.772, 2.056, 0.578, { align: 'center', fontSize: 40, fontFace: MED_F, valign: 'middle', margin: [7.2, 7.2, 7.2, 7.2] });
  text(s, '04', 8.805, 3.772, 2.056, 0.578, { align: 'center', fontSize: 40, fontFace: MED_F, color: WHITE, valign: 'middle', margin: [7.2, 7.2, 7.2, 7.2] });
}

function slide30(s) {   // Infographic - ring of four
  // four gold nodes wired together by a diamond-shaped band
  const cx = 6.58, cy = 4.53, reach = 1.53;
  s.addShape('rect', {
    x: cx - 1.089, y: cy - 1.089, w: 2.178, h: 2.178, rotate: 45,
    fill: { type: 'none' }, line: { color: GOLD, width: 23.8 },
  });
  [[cx, cy - reach], [cx + reach, cy], [cx, cy + reach], [cx - reach, cy]].forEach(function (p) {
    donut(s, p[0], p[1], 0.82, 0.52, GOLD, CREAM);
  });
  icon(s, 6.346, 2.799, 0.531, 0.356, INK);
  icon(s, 4.805, 4.269, 0.443, 0.413, INK);
  icon(s, 7.883, 4.322, 0.516, 0.515, INK);
  icon(s, 6.346, 5.811, 0.531, 0.424, INK);
  text(s, 'Lorep  ipsum duis aute irure dolor in kauselih oiluek eprehfa hreer deriti vols daqeesse cill.', 9.112, 5.518, 2.57, 0.858, { fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  pill(s, 9.106, 4.96, 1.501, 0.392, CREAM);
  band(s, 9.239, 5.064, 1.236, 0.236, CREAM);
  text(s, 'NEW TEXT', 9.239, 5.064, 1.236, 0.236, { align: 'center', fontSize: 8, fontFace: MED_F });
  text(s, 'Lorep  ipsum duis aute irure dolor in kauselih oiluek eprehfa hreer deriti vols daqeesse cill.', 9.112, 3.09, 2.57, 0.858, { fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  pill(s, 9.106, 2.533, 1.501, 0.392, CREAM);
  band(s, 9.239, 2.637, 1.236, 0.236, CREAM);
  text(s, 'NEW TEXT', 9.239, 2.637, 1.236, 0.236, { align: 'center', fontSize: 8, fontFace: MED_F });
  text(s, 'Lorep  ipsum duis aute irure dolor in kauselih oiluek eprehfa hreer deriti vols daqeesse cill.', 1.092, 3.09, 2.57, 0.858, { align: 'right', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  pill(s, 2.161, 2.533, 1.501, 0.392, CREAM);
  band(s, 2.293, 2.637, 1.236, 0.236, CREAM);
  text(s, 'NEW TEXT', 2.293, 2.637, 1.236, 0.236, { align: 'center', fontSize: 8, fontFace: MED_F });
  text(s, 'Lorep  ipsum duis aute irure dolor in kauselih oiluek eprehfa hreer deriti vols daqeesse cill.', 1.092, 5.517, 2.57, 0.858, { align: 'right', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  pill(s, 2.161, 4.96, 1.501, 0.392, CREAM);
  band(s, 2.293, 5.064, 1.236, 0.236, CREAM);
  text(s, 'NEW TEXT', 2.293, 5.064, 1.236, 0.236, { align: 'center', fontSize: 8, fontFace: MED_F });
  text(s, 'Infographic', 4.475, 0.725, 4.383, 0.841, { align: 'center', fontSize: 44, fontFace: TITLE_F });
}

function slide31(s) {   // Infographic - four quadrants
  // circle split into four quadrants, gold and cream alternating
  [[0, CREAM], [1, GOLD], [2, CREAM], [3, GOLD]].forEach(function (q) {
    quarter(s, 6.69, 4.125, 1.828, q[0], q[1]);
  });
  icon(s, 7.217, 3.018, 0.584, 0.594, WHITE);
  icon(s, 5.597, 4.629, 0.574, 0.57, WHITE);
  icon(s, 7.217, 4.629, 0.569, 0.528, UMBER);
  icon(s, 5.598, 3.044, 0.575, 0.575, UMBER);
  text(s, 'Lorep  ipsum duis aute irure dolor in kauselih oiluek eprehfa hreer deriti vols daqeesse cill.', 9.112, 2.955, 2.57, 0.858, { fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'Lorep  ipsum duis aute irure dolor in kauselih oiluek eprehfa hreer deriti vols daqeesse cill.', 9.112, 5.073, 2.57, 0.858, { fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  pill(s, 9.106, 2.327, 1.501, 0.392, GOLD);
  band(s, 9.239, 2.431, 1.236, 0.236, GOLD);
  text(s, 'NEW TEXT', 9.239, 2.431, 1.236, 0.236, { align: 'center', fontSize: 8, fontFace: MED_F, color: WHITE });
  pill(s, 9.106, 4.516, 1.501, 0.392, CREAM);
  band(s, 9.239, 4.619, 1.236, 0.236, CREAM);
  text(s, 'NEW TEXT', 9.239, 4.619, 1.236, 0.236, { align: 'center', fontSize: 8, fontFace: MED_F });
  pill(s, 2.722, 4.516, 1.501, 0.392, CREAM);
  band(s, 2.855, 4.619, 1.236, 0.236, CREAM);
  text(s, 'NEW TEXT', 2.855, 4.619, 1.236, 0.236, { align: 'center', fontSize: 8, fontFace: MED_F });
  text(s, 'Lorep  ipsum duis aute irure dolor in kauselih oiluek eprehfa hreer deriti vols daqeesse cill.', 1.653, 5.073, 2.57, 0.858, { align: 'right', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  pill(s, 2.722, 2.328, 1.501, 0.392, GOLD);
  band(s, 2.855, 2.432, 1.236, 0.236, GOLD);
  text(s, 'NEW TEXT', 2.855, 2.432, 1.236, 0.236, { align: 'center', fontSize: 8, fontFace: MED_F, color: WHITE });
  text(s, 'Lorep  ipsum duis aute irure dolor in kauselih oiluek eprehfa hreer deriti vols daqeesse cill.', 1.653, 2.886, 2.57, 0.858, { align: 'right', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'Infographic', 4.475, 0.725, 4.383, 0.841, { align: 'center', fontSize: 44, fontFace: TITLE_F });
}

function slide32(s) {   // Infographic - timeline
  band(s, 0, 4.144, 13.312, 0.512, BAND);
  dashLine(s, 3.98, 4.249, 1.21, false);
  chevron(s, 3.061, 4.138, 1.839, 0.513, GOLD);
  text(s, 'Text Here', 3.061, 4.138, 1.839, 0.513, { align: 'center', fontSize: 14, fontFace: MED_F, color: WHITE, valign: 'middle', margin: [7.2, 7.2, 7.2, 7.2] });
  icon(s, 3.721, 3.464, 0.518, 0.516, DEEP);
  text(s, 'dolrlasue boru sit amet. Duis  volui repiehenderiti', 2.889, 5.57, 2.182, 0.779, { align: 'center', fontSize: 12, fontFace: BODY_F, margin: [7.2, 7.2, 7.2, 7.2] });
  dashLine(s, 7.439, 4.249, 1.21, false);
  chevron(s, 6.521, 4.138, 1.839, 0.513, GOLD);
  text(s, 'Text Here', 6.521, 4.138, 1.839, 0.513, { align: 'center', fontSize: 14, fontFace: MED_F, color: WHITE, valign: 'middle', margin: [7.2, 7.2, 7.2, 7.2] });
  icon(s, 7.183, 3.461, 0.515, 0.519, DEEP);
  text(s, 'dolrlasue boru sit amet. Duis  volui repiehenderiti', 6.349, 5.57, 2.182, 0.779, { align: 'center', fontSize: 12, fontFace: BODY_F, margin: [7.2, 7.2, 7.2, 7.2] });
  dashLine(s, 10.901, 4.249, 1.21, false);
  chevron(s, 9.994, 4.138, 1.813, 0.513, GOLD);
  text(s, 'Text Here', 9.994, 4.138, 1.813, 0.513, { align: 'center', fontSize: 14, fontFace: MED_F, color: WHITE, valign: 'middle', margin: [7.2, 7.2, 7.2, 7.2] });
  icon(s, 10.642, 3.465, 0.516, 0.515, DEEP);
  text(s, 'dolrlasue boru sit amet. Duis  volui repiehenderiti', 9.809, 5.57, 2.182, 0.779, { align: 'center', fontSize: 12, fontFace: BODY_F, margin: [7.2, 7.2, 7.2, 7.2] });
  dashLine(s, 2.25, 3.318, 1.21, true);
  chevron(s, 1.331, 4.138, 1.839, 0.513, CREAM);
  text(s, 'Text Here', 1.331, 4.138, 1.839, 0.513, { align: 'center', fontSize: 14, fontFace: MED_F, valign: 'middle', margin: [7.2, 7.2, 7.2, 7.2] });
  icon(s, 1.993, 4.815, 0.515, 0.515, DEEP);
  text(s, 'dolrlasue boru sit amet. Duis  volui repiehenderiti', 1.159, 2.428, 2.182, 0.779, { align: 'center', fontSize: 12, fontFace: BODY_F, margin: [7.2, 7.2, 7.2, 7.2] });
  dashLine(s, 5.709, 3.318, 1.21, true);
  chevron(s, 4.791, 4.138, 1.839, 0.513, CREAM);
  text(s, 'Text Here', 4.791, 4.138, 1.839, 0.513, { align: 'center', fontSize: 14, fontFace: MED_F, valign: 'middle', margin: [7.2, 7.2, 7.2, 7.2] });
  icon(s, 5.449, 4.815, 0.522, 0.518, DEEP);
  text(s, 'dolrlasue boru sit amet. Duis  volui repiehenderiti', 4.618, 2.428, 2.182, 0.779, { align: 'center', fontSize: 12, fontFace: BODY_F, margin: [7.2, 7.2, 7.2, 7.2] });
  dashLine(s, 9.17, 3.318, 1.21, true);
  chevron(s, 8.251, 4.138, 1.839, 0.513, CREAM);
  text(s, 'Text Here', 8.251, 4.138, 1.839, 0.513, { align: 'center', fontSize: 14, fontFace: MED_F, valign: 'middle', margin: [7.2, 7.2, 7.2, 7.2] });
  icon(s, 8.912, 4.815, 0.516, 0.518, DEEP);
  text(s, 'dolrlasue boru sit amet. Duis  volui repiehenderiti', 8.078, 2.428, 2.182, 0.779, { align: 'center', fontSize: 12, fontFace: BODY_F, margin: [7.2, 7.2, 7.2, 7.2] });
  text(s, 'Infographic', 4.475, 0.725, 4.383, 0.841, { align: 'center', fontSize: 44, fontFace: TITLE_F });
}

function slide33(s) {   // Infographic - pinwheel
  text(s, 'Infographic', 4.475, 0.725, 4.383, 0.841, { align: 'center', fontSize: 44, fontFace: TITLE_F });
  wedge(s, 5.218, 0.912, 2.967, 180, GOLD);
  wedge(s, 5.148, 2.395, 2.967, 0, CREAM);
  wedge(s, 5.148, 2.456, 2.967, -90, GOLD);
  wedge(s, 5.218, 3.944, 2.967, 90, CREAM);
  icon(s, 5.698, 4.225, 0.668, 0.663, WHITE);
  icon(s, 7.002, 2.721, 0.647, 0.647, WHITE);
  icon(s, 6.962, 4.413, 0.665, 0.644, INK);
  icon(s, 5.698, 2.939, 0.66, 0.66, INK);
  text(s, 'Lorep  ipsum duis aute irure dolor in kauselih oilusioi reprehenderitilores voluptates esse', 1.412, 3.055, 2.579, 0.688, { fontSize: 8, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'TEXT TITTLE HERE', 1.412, 2.469, 1.593, 0.286, { fontSize: 10.5, fontFace: SEMI_F, color: NEAR_INK });
  rule(s, 1.512, 2.904, 2.109, RULE);
  text(s, 'Lorep  ipsum duis aute irure dolor in kauselih oilusioi reprehenderitilores voluptates esse', 1.412, 4.988, 2.579, 0.688, { fontSize: 8, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'TEXT TITTLE HERE', 1.412, 4.402, 1.593, 0.286, { fontSize: 10.5, fontFace: SEMI_F, color: NEAR_INK });
  rule(s, 1.512, 4.837, 2.109, RULE);
  text(s, 'Lorep  ipsum duis aute irure dolor in kauselih oilusioi reprehenderitilores voluptates esse', 9.306, 3.055, 2.579, 0.688, { align: 'right', fontSize: 8, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'TEXT TITTLE HERE', 10.41, 2.469, 1.593, 0.286, { align: 'right', fontSize: 10.5, fontFace: SEMI_F, color: NEAR_INK });
  rule(s, 9.765, 2.904, 2.109, RULE);
  text(s, 'Lorep  ipsum duis aute irure dolor in kauselih oilusioi reprehenderitilores voluptates esse', 9.249, 4.99, 2.579, 0.688, { align: 'right', fontSize: 8, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'TEXT TITTLE HERE', 10.353, 4.403, 1.593, 0.286, { align: 'right', fontSize: 10.5, fontFace: SEMI_F, color: NEAR_INK });
  rule(s, 9.708, 4.838, 2.109, RULE);
}

function slide34(s) {   // Quotes of Santana
  band(s, 11.538, 0, 1.795, 0.615, CREAM);
  band(s, 7.211, 0, 4.334, 0.615, GOLD);
  text(s, 'Quotes of Santana', 1.301, 0.932, 5.265, 1.582, { fontSize: 44, fontFace: TITLE_F });
  text(s, 'Lorem Ipsum', 4.754, 6.396, 1.651, 0.33, { fontSize: 10, fontFace: MED_F, color: GREY, lineSpacing: 1.5 });
  text(s, 'Text Tittle Here', 4.754, 6.086, 1.651, 0.352, { fontSize: 11, fontFace: SEMI_F, color: NEAR_INK, lineSpacing: 1.5 });
  text(s, 'Lorep  ipsum duis aute irure dolor in kauselih kauilis oilusioalio repres henderiti voluptates esse cill inure dolorlaboru sits amet. Duis aute irusita dolori inalisk reprehenderit inskalieos voluptate.', 1.301, 3.357, 5.366, 1.481, { align: 'justify', fontSize: 14, fontFace: BODY_F, color: '262626', lineSpacing: 1.5 });
  text(s, '“', 1.301, 2.896, 2.861, 0.841, { fontSize: 44, fontFace: SEMI_F });
  text(s, 'Santana Tatiana', 1.301, 5.402, 1.651, 0.375, { fontSize: 12, fontFace: BODY_F, color: NEAR_INK, lineSpacing: 1.5 });
}

function slide35(s) {   // Contact Us
  band(s, 12.524, 5.063, 0.81, 2.437, GOLD);
  band(s, 5.952, 5.063, 6.571, 2.437, CREAM);
  text(s, L.a + ' ', 6.602, 1.953, 5.295, 0.834, { align: 'justify', fontSize: 10, fontFace: BODY_F, lineSpacing: 1.5 });
  text(s, 'OUR ADDRESS', 6.602, 4.008, 1.534, 0.303, { align: 'center', fontSize: 12, fontFace: SEMI_F });
  text(s, [
      { text: '123, Rev Avenue, ', options: { breakLine: true } },
      { text: 'Kolabagan LA , 267' },
    ], 6.663, 4.375, 1.412, 0.555, { align: 'center', fontSize: 9, fontFace: BODY_F, lineSpacing: 1.5 });
  icon(s, 7.205, 3.213, 0.358, 0.359, GOLD);
  text(s, 'GET IN TOUCH', 8.788, 4.038, 1.551, 0.303, { fontSize: 12, fontFace: SEMI_F });
  text(s, [
      { text: '(+11) 185 6554 3435', options: { breakLine: true } },
      { text: '(+11) 189 6398 3432' },
    ], 8.788, 4.404, 1.443, 0.534, { fontSize: 9, fontFace: BODY_F, lineSpacing: 1.5 });
  icon(s, 9.351, 3.215, 0.407, 0.41, GOLD);
  text(s, 'FOLLOW US', 11.125, 4.005, 1.279, 0.303, { align: 'center', fontSize: 12, fontFace: SEMI_F });
  text(s, [
      { text: 'www.yourtext.com', options: { breakLine: true } },
      { text: 'example@gmail.com' },
    ], 10.958, 4.372, 1.613, 0.555, { align: 'center', fontSize: 9, fontFace: BODY_F, lineSpacing: 1.5 });
  icon(s, 11.571, 3.211, 0.388, 0.361, GOLD);
  text(s, 'Contact Us', 6.602, 0.669, 5.295, 0.909, { fontSize: 48, fontFace: TITLE_F });
}

function slide36(s) {   // Thank You
  coverSlide(s, 'Thank You', [0.871, 2.594, 7.055], [7.617, 4.097], WHITE);
}

/* --------------------------------------------------------------------- main */
const builders = [
  slide01,
  slide02,
  slide03,
  slide04,
  slide05,
  slide06,
  slide07,
  slide08,
  slide09,
  slide10,
  slide11,
  slide12,
  slide13,
  slide14,
  slide15,
  slide16,
  slide17,
  slide18,
  slide19,
  slide20,
  slide21,
  slide22,
  slide23,
  slide24,
  slide25,
  slide26,
  slide27,
  slide28,
  slide29,
  slide30,
  slide31,
  slide32,
  slide33,
  slide34,
  slide35,
  slide36
];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'SANTANA', width: 13.333, height: 7.5 });
pptx.layout = 'SANTANA';
pptx.title = 'Santana - Fashion Presentation Template';

builders.forEach(function (build) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  build(s);
});

pptx.writeFile({ fileName: path.join(__dirname, '1a0ea5ed-a8a2-402a-8e28-42b28ee95022_grok_final.pptx') })
  .then(function (f) { console.log('wrote ' + f); })
  .catch(function (e) { console.error(e); process.exit(1); });
