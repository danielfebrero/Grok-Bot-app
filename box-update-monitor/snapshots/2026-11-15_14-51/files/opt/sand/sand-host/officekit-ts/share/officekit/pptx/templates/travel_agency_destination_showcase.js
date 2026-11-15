/**
 * "Trove." travel presentation - recreated with pptxgenjs.
 * Raster photos from the source deck are replaced by flat colour placeholders.
 *
 *   node 03f83b15-59f0-468c-89e0-b0c879612a52_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ tokens */

const C = {
  bg: 'F2F2F2',        // page wash (15% opaque over white)
  white: 'FFFFFF',
  dark: '262626',
  body: '374151',
  black: '000000',
  orange: 'FF9801',
  teal: '04A7BE',
  star: 'FFC000',
  photo: '3C3C3C',     // average colour of the replaced photographs
  photoPale: 'D9D7D8',
  frame: '656564',   // brushed-metal edge of the phone mockup
};

const F = { head: 'Poppins SemiBold', ui: 'Poppins', body: 'Open Sans' };

// pptxgenjs rewrites the shadow object it is handed, so hand it a fresh one every time
const shadow = (color) => ({ type: 'outer', color: color || 'A5A5A5', blur: 15, offset: 3, angle: 45, opacity: 0.4 });

// text insets of the original deck: 0.075" sides / 0.0375" top+bottom, in points
const MARGIN = [5.4, 5.4, 2.7, 2.7];

// every rounded rectangle in the deck uses the same 16.667% corner radius
const RR = 0.16667;
const rad = (w, h) => Math.min(w, h) * RR;

const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipiscing';
const LOREM_LONG = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod ' +
  'tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud';
const CAPTURE = 'Capture the beauty of your travels through the lens, preserving moments that last a lifetime';
const TITLE_TRAVEL = 'Travel Beyond Boundaries: A Journey of Discovery Awaits';
const TITLE_FOODIE = 'Foodie Adventures: Culinary Delights from Around the Globe';

/* ----------------------------------------------------------------- helpers */

function text(slide, str, o) {
  slide.addText(str, Object.assign({
    fontFace: F.body, fontSize: 11, color: C.dark,
    align: 'left', valign: 'top', margin: MARGIN,
  }, o));
}

/** 24pt section heading (Poppins SemiBold) */
const heading = (s, x, y, w, h, str, o) =>
  text(s, str, Object.assign({ x, y, w, h, fontFace: F.head, fontSize: 24, color: C.dark }, o));

/** 8pt eyebrow label (Poppins) */
const label = (s, x, y, w, h, str, o) =>
  text(s, str, Object.assign({ x, y, w, h, fontFace: F.ui, fontSize: 8, color: C.orange }, o));

/** 11pt running copy (Open Sans) */
const para = (s, x, y, w, h, str, o) =>
  text(s, str, Object.assign({ x, y, w, h, color: C.body }, o));

/** 9pt / 7pt caption with the 150% leading used all over the deck */
const caption = (s, x, y, w, h, str, o) =>
  text(s, str, Object.assign({ x, y, w, h, fontSize: 9, lineSpacingMultiple: 1.5 }, o));

/** page wash - drawn first on every slide */
const bgRect = (s) =>
  s.addShape('rect', { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.bg, transparency: 85 } });

/** white rounded card with the deck's standard drop shadow */
function card(s, x, y, w, h, o) {
  o = o || {};
  const opts = { x, y, w, h, fill: { color: o.fill || C.white }, rectRadius: rad(w, h) };
  if (!o.flat) opts.shadow = shadow(o.shadowColor);
  s.addShape('roundRect', opts);
}

/** placeholder standing in for a photograph of the source deck */
function photo(s, x, y, w, h, o) {
  o = o || {};
  const shape = o.shape || 'roundRect';
  const opts = { x, y, w, h, fill: { color: o.color || C.photo } };
  if (shape === 'roundRect') opts.rectRadius = o.radius || rad(w, h);
  if (o.rotate) opts.rotate = o.rotate;
  s.addShape(shape, opts);
}

/** "Trove." wordmark + About / Service / Gallery menu */
function logoNav(s) {
  text(s, 'Trove.', { x: 0.4733, y: 0.4286, w: 0.6195, h: 0.2524, fontFace: F.ui, fontSize: 11, bold: true });
  [['About', 7.2051, 0.4657], ['Service', 7.982, 0.5341], ['Gallery', 8.8588, 0.5209]]
    .forEach(([str, x, w]) => label(s, x, 0.4286, w, 0.202, str, { color: C.dark }));
}

/** white "123+ / Destination" chip (1.435 x 0.807) */
function chipDestination(s, x, y) {
  card(s, x, y, 1.4349, 0.8065);
  text(s, '123+', { x: x + 0.0322, y: y + 0.0878, w: 1.3704, h: 0.5301, fontFace: F.ui, fontSize: 27, bold: true, align: 'center' });
  label(s, x + 0.0322, y + 0.5105, 1.3704, 0.2083, 'Destination', { align: 'center' });
}

/** teal "782+ / Client" chip (text boxes overhang the tile on both sides) */
function chipClient(s, x, y) {
  card(s, x + 0.2131, y, 0.9607, 0.591, { fill: C.teal });
  text(s, '782+', { x: x + 0.113, y: y + 0.0558, w: 1.1867, h: 0.3787, fontFace: F.ui, fontSize: 18, bold: true, color: C.white, align: 'center' });
  label(s, x, y + 0.3121, 1.3704, 0.2083, 'Client', { color: C.white, align: 'center' });
}

/** white card carrying one 9pt lorem line (used as a floating note) */
function noteCard(s, x, y, w, h) {
  card(s, x, y, w, h);
  caption(s, x + 0.2713, y + 0.1386, 2.2948, 0.5059, LOREM_SHORT);
}

/** small "title + 7pt lorem" card, 2.510 x 0.876 */
function serviceCard(s, x, y, title, o) {
  o = o || {};
  const fg = o.fill === C.orange ? C.white : C.dark;
  card(s, x, y, 2.5104, 0.8762, { fill: o.fill, shadowColor: o.fill === C.orange ? C.orange : undefined });
  text(s, title, { x: x + 0.2892, y: y + 0.1409, w: o.titleW || 0.7983, h: 0.2524, fontFace: F.ui, fontSize: 11, bold: true, color: fg });
  caption(s, x + 0.2892, y + 0.3103, 2.2948, 0.3983, LOREM_SHORT, { fontSize: 7, color: fg });
}

/* ------------------------------------------------------------------ slides */

function slide01(s) {                                     // cover
  bgRect(s);
  s.addShape('chord', { x: 1.1091, y: 1.4707, w: 3.7977, h: 3.7977, fill: { color: C.orange }, angleRange: [45, 270], rotate: 112.59 });
  card(s, 1.0432, 4.0774, 3.8318, 0.8762);
  card(s, 1.2708, 4.2603, 1.0625, 0.5104, { fill: C.teal, flat: true });
  logoNav(s);
  text(s, 'Read More', { x: 1.4115, y: 4.4082, w: 0.7813, h: 0.2146, fontFace: F.ui, fontSize: 8, bold: true, color: C.white, align: 'center' });
  caption(s, 2.6044, 4.2647, 2.2948, 0.5059, LOREM_SHORT);
  text(s, 'Trove.', { x: 5.6801, y: 1.3273, w: 3.3635, h: 1.2875, fontFace: F.ui, fontSize: 72, bold: true });
  caption(s, 5.6801, 2.6148, 4.4605, 0.2787, 'TRAVEL PRESENTATION TEMPLATE');
  caption(s, 5.6756, 4.3698, 3.7041, 0.5059,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt uta');
}

function slide02(s) {                                     // about us, photo left
  bgRect(s);
  s.addShape('rect', { x: 0, y: 3.8197, w: 1.8929, h: 1.8053, fill: { color: C.teal } });
  card(s, 1.079, 1.1854, 3.302, 3.2541);
  logoNav(s);
  heading(s, 5.1337, 1.8821, 4.2459, 1.2875, TITLE_FOODIE, { align: 'right' });
  text(s, 'About Us', { x: 7.982, y: 1.6084, w: 1.3704, h: 0.2524, fontFace: F.ui, color: C.orange, align: 'right' });
  para(s, 7.7708, 3.6879, 1.6088, 0.9593, CAPTURE, { align: 'right' });
  para(s, 5.8229, 3.6879, 1.6088, 0.9593, CAPTURE, { align: 'right' });
  photo(s, 1.2049, 1.3096, 3.05, 3.0058);
}

function slide03(s) {                                     // about us, orange pillar
  bgRect(s);
  s.addShape('round2SameRect', { x: 7.5614, y: 1.8509, w: 1.8929, h: 3.7907, fill: { color: C.orange } });
  logoNav(s);
  heading(s, 0.4733, 1.9526, 4.4045, 1.2875, TITLE_TRAVEL);
  para(s, 0.4733, 3.5942, 3.7041, 0.4292, CAPTURE);
  text(s, 'About Us', { x: 0.4733, y: 1.5984, w: 1.3704, h: 0.2524, fontFace: F.ui, color: C.orange });
  chipDestination(s, 5.5807, 4.0744);
  chipClient(s, 7.8416, 3.0938);
}

function slide04(s) {                                     // portrait photo + inset
  bgRect(s);
  s.addShape('round2SameRect', { x: 0, y: 1.8929, w: 2.7456, h: 3.7321, fill: { color: C.orange } });
  card(s, 1.0928, 0.9856, 2.9725, 3.9192);
  card(s, 3.6861, 3.4486, 1.5996, 1.5764);
  photo(s, 1.2188, 1.1097, 2.7456, 3.6916);
  photo(s, 3.7471, 3.5088, 1.4775, 1.4561, { color: C.photoPale });
  logoNav(s);
  heading(s, 5.0401, 1.6011, 4.5493, 1.2875, TITLE_TRAVEL);
  para(s, 5.7908, 4.0222, 3.7041, 0.4292, CAPTURE, { color: C.dark });
  chipClient(s, 3.3297, 2.66);
}

function slide05(s) {                                     // mirrored variant of 04
  bgRect(s);
  s.addShape('round2SameRect', { x: 5.3116, y: 1.8929, w: 2.7456, h: 3.7321, fill: { color: C.orange } });
  logoNav(s);
  card(s, 6.508, 1.1682, 2.9725, 3.9192);
  card(s, 0.6484, 4.1139, 0.9091, 0.896);
  card(s, 2.0622, 4.1139, 0.9091, 0.896);
  heading(s, 0.5195, 1.1076, 4.5493, 1.2875, TITLE_TRAVEL);
  para(s, 0.6203, 3.2323, 3.7041, 0.4292, CAPTURE, { color: C.dark });
  photo(s, 6.634, 1.2924, 2.7456, 3.6916);
  photo(s, 0.6814, 4.1481, 0.8398, 0.8276);
  photo(s, 2.098, 4.1481, 0.8398, 0.8276);
  chipClient(s, 5.8145, 4.1481);
}

function slide06(s) {                                     // wide photo card
  bgRect(s);
  s.addShape('rect', { x: 0, y: 3.8197, w: 4.2459, h: 1.8053, fill: { color: C.teal } });
  logoNav(s);
  card(s, 0.5075, 3.3176, 4.7976, 1.5764);
  heading(s, 0.5075, 1.525, 4.2459, 1.2875, TITLE_FOODIE);
  para(s, 6.2383, 3.8029, 3.1414, 0.6059, CAPTURE);
  chipDestination(s, 8.0576, 1.1218);
  chipClient(s, 6.4539, 2.4033);
  photo(s, 0.5686, 3.3778, 4.6658, 1.4561);
}

function slide07(s) {                                     // teal pillar with two tiles
  bgRect(s);
  s.addShape('round2SameRect', { x: 6.1615, y: 1.0388, w: 2.7456, h: 4.5862, fill: { color: C.teal } });
  card(s, 7.7162, 1.2962, 1.5996, 1.5764);
  card(s, 7.7162, 3.3379, 1.5996, 1.5764);
  logoNav(s);
  heading(s, 0.9383, 2.0504, 4.0574, 1.2875, TITLE_TRAVEL);
  para(s, 0.9609, 3.5603, 3.1414, 0.6059, CAPTURE);
  photo(s, 7.7773, 3.398, 1.4775, 1.4561);
  photo(s, 7.7773, 1.3564, 1.4775, 1.4561);
  noteCard(s, 5.4919, 1.7309, 2.5104, 0.8762);
  noteCard(s, 5.4919, 3.71, 2.5104, 0.8762);
}

function slide08(s) {                                     // break slide
  bgRect(s);
  logoNav(s);
  text(s, 'Break \nSlide.', { x: 0.4733, y: 1.4783, w: 4.1834, h: 2.4992, fontFace: F.ui, fontSize: 72, bold: true });
  caption(s, 0.4733, 4.1294, 3.1928, 0.2787, 'Time To Drink');
  s.addShape('round2SameRect', { x: 5.9509, y: 1.8929, w: 2.7456, h: 3.7321, fill: { color: C.orange } });
  chipClient(s, 4.7672, 1.1744);
  chipDestination(s, 3.93, 4.121);
  photo(s, 5.7887, 2.2239, 3.0701, 3.0701, { shape: 'ellipse' });
}

function slide09(s) {                                     // tilted photo cards
  bgRect(s);
  s.addShape('rect', { x: 5.7541, y: 1.2142, w: 4.2459, h: 1.8053, fill: { color: C.teal } });
  [[4.5739, 1.8903], [7.267, 1.0679]].forEach(([x, y]) =>
    s.addShape('roundRect', {
      x, y, w: 2.1232, h: 2.7995, rotate: 18.704,
      fill: { color: C.white }, rectRadius: rad(2.1232, 2.7995), shadow: shadow(),
    }));
  logoNav(s);
  heading(s, 0.4733, 1.7064, 4.0574, 0.8835, 'About Our \nService');
  para(s, 0.4959, 2.6842, 3.1414, 0.6059, CAPTURE);
  [[7.3481, 1.1492], [4.655, 1.9716]].forEach(([x, y]) =>
    photo(s, x, y, 1.9612, 2.6369, { rotate: 18.704, radius: 0.2013 }));
  serviceCard(s, 4.2167, 4.5182, 'Service 1');
  serviceCard(s, 6.9771, 3.7021, 'Service 1');
}

function slide10(s) {                                     // service list
  bgRect(s);
  s.addShape('round2SameRect', { x: 2.1498, y: 3.7326, w: 1.8495, h: 1.8924, fill: { color: C.teal } });
  logoNav(s);
  card(s, 0.4733, 1.1097, 1.9925, 1.5764);
  card(s, 3.1265, 3.219, 3.1986, 1.5764);
  heading(s, 2.9713, 1.2289, 4.0574, 0.8835, 'About Our \nService');
  para(s, 2.9391, 2.2969, 2.954, 0.6059, CAPTURE);
  [1.4602, 2.6894, 3.9192].forEach((y) => serviceCard(s, 6.8849, y, 'Service 1'));
  photo(s, 0.5311, 1.1699, 1.877, 1.4561);
  photo(s, 3.2192, 3.2791, 3.0132, 1.4561);
}

function slide11(s) {                                     // guide portrait, dark banner
  bgRect(s);
  s.addShape('round2SameRect', { x: 0.2737, y: 2.0517, w: 1.3451, h: 1.8924, fill: { color: C.teal }, rotate: 90 });
  photo(s, 1.5192, 1.5865, 1.877, 3.1906);
  card(s, 1.4614, 1.5264, 1.9925, 3.3158);
  photo(s, 0.6875, 1.4974, 8.625, 3.569, { shape: 'rect' });
  logoNav(s);
  serviceCard(s, 2.9801, 2.9653, 'Team Work', { titleW: 1.0192 });
  serviceCard(s, 2.9801, 4.1558, 'Leadership', { titleW: 0.989, fill: C.orange });
  heading(s, 6.1697, 1.8563, 2.8943, 0.4796, 'Alexander Baiu');
  para(s, 6.1378, 3.5235, 3.1414, 0.6059, CAPTURE, { color: C.dark });
  label(s, 6.1677, 1.5865, 2.375, 0.2083, 'About Our Guide');
}

function slide12(s) {                                     // about team
  bgRect(s);
  card(s, 4.7793, 3.6199, 3.1986, 1.5764);
  photo(s, 4.872, 3.6801, 3.0132, 1.4561);
  logoNav(s);
  // avatar glyph: orange tile + white head and shoulders
  card(s, 0.5018, 1.8195, 0.5911, 0.5794, { fill: C.orange, shadowColor: C.orange });
  s.addShape('chord', { x: 0.6123, y: 2.1064, w: 0.37, h: 0.1697, fill: { color: C.white }, angleRange: [180, 360] });
  s.addShape('ellipse', { x: 0.7176, y: 1.9423, w: 0.1595, h: 0.1595, fill: { color: C.white } });
  card(s, 1.1855, 1.5136, 3.2312, 1.0776);
  label(s, 1.6136, 1.734, 2.375, 0.2083, 'Trove Travel presentation', { align: 'center' });
  text(s, 'About Team', { x: 1.1504, y: 1.8808, w: 3.3013, h: 0.5301, fontFace: F.ui, fontSize: 27, bold: true, align: 'center' });
  para(s, 1.0928, 3.0742, 2.954, 0.6059, CAPTURE);
  chipClient(s, 3.8895, 3.8703);
  card(s, 6.9644, 0.868, 2.4285, 3.2019);
  photo(s, 7.0571, 0.961, 2.2431, 3.0159);
}

function slide13(s) {                                     // expert team
  bgRect(s);
  card(s, 0.4733, 1.3308, 1.9925, 3.3567);
  card(s, 6.5293, 3.8993, 3.1986, 1.5764);
  photo(s, 0.5311, 1.3856, 1.877, 3.247);
  photo(s, 6.622, 3.9594, 3.0132, 1.4561);
  logoNav(s);
  heading(s, 2.9391, 1.7257, 5.488, 0.4796, 'We are the Expert team');
  para(s, 2.9269, 2.7252, 3.1414, 0.6059, CAPTURE, { color: C.dark });
  label(s, 2.9371, 1.4559, 2.375, 0.2083, 'Trove Travel presentation');
  noteCard(s, 4.6947, 4.2494, 2.5104, 0.8762);
  para(s, 6.6763, 2.6864, 2.8504, 0.9593, LOREM_LONG, { color: C.black });
  chipDestination(s, 1.6613, 4.2294);
}

function slide14(s) {                                     // meet all guide
  bgRect(s);
  s.addShape('round2SameRect', { x: 4.9363, y: 0.5912, w: 2.7456, h: 7.3906, fill: { color: C.orange }, rotate: 270 });
  logoNav(s);
  const guides = [
    { x: 4.0066, name: 'Adam', nameX: 4.3861, nameW: 0.8407, textX: 3.9758, textY: 4.0392 },
    { x: 5.9037, name: 'Bianca', nameX: 6.2832, nameW: 0.8407, textX: 5.8729, textY: 4.0365 },
    { x: 7.7801, name: 'Lorenac', nameX: 8.2017, nameW: 0.7563, textX: 7.8352, textY: 4.0365 },
  ];
  guides.forEach((g) => card(s, g.x, 1.929, 1.5996, 1.5764));
  heading(s, 0.4733, 1.1097, 3.5333, 0.4796, 'Meet All Guide');
  guides.forEach((g) =>
    caption(s, g.textX, g.textY, 1.6613, 0.7331, LOREM_SHORT, { color: C.white, align: 'center' }));
  guides.forEach((g) =>
    text(s, g.name, { x: g.nameX, y: g.name === 'Lorenac' ? 3.7207 : 3.7513, w: g.nameW, h: 0.2524, fontFace: F.ui, fontSize: 11, bold: true, align: 'center' }));
  guides.forEach((g) => photo(s, g.x + 0.0611, 1.9892, 1.4775, 1.4561));
}

function slide15(s) {                                     // best destination
  bgRect(s);
  logoNav(s);
  card(s, 0.4733, 1.3308, 1.9925, 3.3567);
  heading(s, 2.9391, 1.3856, 3.5252, 0.8835, 'Best \nDestination');
  card(s, 2.8464, 3.5468, 3.1986, 1.5764);
  card(s, 6.7544, 2.6681, 2.4285, 2.5078);
  para(s, 2.9795, 2.605, 2.8419, 0.6059, CAPTURE, { color: C.dark });
  photo(s, 0.5311, 1.3856, 1.877, 3.247);
  photo(s, 6.8471, 2.7409, 2.2431, 2.3621);
  photo(s, 2.9391, 3.6069, 3.0132, 1.4561);
  chipDestination(s, 7.2512, 2.2018);
}

function slide16(s) {                                     // recommendation destination
  bgRect(s);
  s.addShape('round2SameRect', { x: 7.814, y: 0.971, w: 2.0268, h: 2.3453, fill: { color: C.orange }, rotate: 270 });
  logoNav(s);
  heading(s, 0.7728, 2.0404, 3.7048, 0.8835, 'Rekomendation \nDestination');
  para(s, 0.8132, 3.2598, 2.8419, 0.6059, CAPTURE, { color: C.dark });
  card(s, 4.5163, 0.7865, 2.4285, 2.5078);
  card(s, 7.0877, 2.1436, 1.8694, 2.5078);
  photo(s, 7.1591, 2.2165, 1.7268, 2.3621);
  photo(s, 4.609, 0.8593, 2.2431, 2.3621);
  chipClient(s, 6.1669, 3.5731);
}

function slide17(s) {                                     // three destination cards
  bgRect(s);
  logoNav(s);
  const cols = [1.0928, 3.7833, 6.4857];
  const textX = [1.117, 3.8551, 6.6051];   // caption boxes drift right of their cards
  cols.forEach((x) => card(s, x, 1.8011, 2.4285, 2.5078));
  textX.forEach((x) => {
    text(s, LOREM_SHORT, { x, y: 4.8564, w: 2.3313, h: 0.3787, fontSize: 9, align: 'center' });
    text(s, 'Destination', { x: x + 0.5733, y: 4.5422, w: 1.1846, h: 0.2524, fontFace: F.ui, fontSize: 11, bold: true, align: 'center' });
  });
  heading(s, 2.2261, 1.0251, 5.5478, 0.4796, 'Rekomendation Destination', { align: 'center' });
  cols.forEach((x) => photo(s, x + 0.0927, 1.8739, 2.2431, 2.3621));
}

function slide18(s) {                                     // our gallery, orange panel
  bgRect(s);
  card(s, 0.6309, 1.1097, 2.4285, 2.5078);
  photo(s, 0.7236, 1.1825, 2.2431, 2.3621);
  photo(s, 7.2468, 2.9885, 2.2431, 2.3621);
  card(s, 7.1541, 2.9157, 2.4285, 2.5078);
  logoNav(s);
  card(s, 2.2128, 2.4364, 4.1999, 2.5078, { fill: C.orange, flat: true });
  para(s, 2.5419, 2.9078, 3.5416, 0.7826, LOREM_LONG, { color: C.black });
  para(s, 2.5419, 3.8349, 3.5416, 0.7826, LOREM_LONG, { color: C.black });
  heading(s, 5.5129, 1.4777, 2.8554, 0.4796, 'Our Galley');
  photo(s, 7.2468, 2.9885, 2.2431, 2.3621);
}

function slide19(s) {                                     // gallery mosaic
  bgRect(s);
  logoNav(s);
  card(s, 0.6266, 3.4962, 4.7976, 1.5764);
  card(s, 0.5339, 0.9283, 2.4285, 2.5078);
  card(s, 5.5168, 1.7159, 1.9925, 3.3567);
  heading(s, 8.0773, 2.8125, 1.5629, 0.8835, 'Our \nGalley');
  para(s, 3.2841, 1.7708, 1.911, 0.9593, CAPTURE, { color: C.dark });
  photo(s, 0.6266, 1.0011, 2.2431, 2.3621);
  photo(s, 5.5745, 1.7708, 1.877, 3.247);
  photo(s, 0.6876, 3.5564, 4.6658, 1.4561);
}

function slide20(s) {                                     // phone mockup
  bgRect(s);
  s.addShape('round2SameRect', { x: -0.009, y: -2.2989, w: 3.8646, h: 7.3906, fill: { color: C.orange }, rotate: 180 });
  logoNav(s);
  // phone mockup: metallic frame with a dark screen inside
  photo(s, 0.6505, 2.1237, 2.5456, 4.5148, { color: C.frame, radius: 0.32 });
  photo(s, 0.7831, 2.1795, 2.2606, 4.3548, { radius: 0.2662 });
  heading(s, 4.3129, 1.571, 5.488, 0.8835, 'About Mockup\nTrove Travel');
  para(s, 4.3006, 2.8032, 2.2708, 0.7826, CAPTURE, { color: C.dark });
  para(s, 7.2051, 3.7116, 2.2708, 1.136, LOREM_LONG, { color: C.black });
  chipDestination(s, 4.3129, 4.1288);
  chipClient(s, 2.458, 1.884);
}

function slide21(s) {                                     // pricing table
  bgRect(s);
  logoNav(s);
  const plans = [
    { // Standard - centre, taller, orange
      outer: [3.8327, 1.737, 2.3887, 3.3567], inner: [3.8904, 1.7919, 2.2502, 3.247], fill: C.orange,
      name: 'Standard', nameBox: [4.425, 2.0858, 1.1811, 0.3029], nameSize: 14,
      price: '$45/Pack', priceBox: [4.0846, 2.6598, 1.8618, 0.4292], priceSize: 21,
      rows: [[4.2989, 3.3601, 1.8418], [4.2989, 3.9509, 1.9225]], rowH: 0.4701, rowSize: 8,
      dots: [[4.2513, 3.4633], [4.2513, 4.054]],
    },
    { // Basic - left
      outer: [1.4499, 2.4872, 1.846, 2.5941], inner: [1.5223, 2.5296, 1.739, 2.5093], fill: C.teal,
      name: 'Basic', nameBox: [1.9653, 2.778, 0.8407, 0.2524], nameSize: 11,
      price: '$25/Pack', priceBox: [1.555, 3.195, 1.6613, 0.3787], priceSize: 18,
      rows: [[1.7556, 3.7507, 1.6613], [1.7556, 4.1826, 1.6613]], rowH: 0.3983, rowSize: 7,
      dots: [[1.708, 3.8538], [1.708, 4.2857]],
    },
    { // Premium - right
      outer: [6.7289, 2.4448, 1.846, 2.5941], inner: [6.8013, 2.4872, 1.739, 2.5093], fill: C.teal,
      name: 'Premium', nameBox: [7.1113, 2.778, 1.1694, 0.2524], nameSize: 11,
      price: '$65/Pack', priceBox: [6.8653, 3.195, 1.6613, 0.3787], priceSize: 18,
      rows: [[7.0658, 3.7507, 1.6613], [7.0658, 4.1826, 1.6613]], rowH: 0.3983, rowSize: 7,
      dots: [[7.0183, 3.8538], [7.0183, 4.2857]],
    },
  ];
  // shells first (matching the source z-order), then all labels
  plans.forEach((p) => {
    card(s, p.outer[0], p.outer[1], p.outer[2], p.outer[3]);
    card(s, p.inner[0], p.inner[1], p.inner[2], p.inner[3], { fill: p.fill, flat: true });
  });
  heading(s, 3.2489, 0.8182, 3.5333, 0.4796, 'Pricing Package', { align: 'center' });
  plans.slice(1).concat(plans.slice(0, 1)).forEach((p) => {
    text(s, p.name, { x: p.nameBox[0], y: p.nameBox[1], w: p.nameBox[2], h: p.nameBox[3], fontFace: F.ui, fontSize: p.nameSize, bold: true, color: C.white, align: 'center' });
    text(s, p.price, { x: p.priceBox[0], y: p.priceBox[1], w: p.priceBox[2], h: p.priceBox[3], fontFace: F.ui, fontSize: p.priceSize, bold: true, color: C.white, align: 'center' });
    p.rows.forEach(([x, y, w], i) => {
      caption(s, x, y, w, p.rowH, LOREM_SHORT, { fontSize: p.rowSize, color: C.white });
      const [dx, dy] = p.dots[i];
      s.addShape('ellipse', { x: dx, y: dy, w: 0.0625, h: 0.0625, fill: { color: C.white } });
    });
  });
}

function slide22(s) {                                     // testimonials
  bgRect(s);
  s.addShape('round2SameRect', { x: 7.814, y: 3.439, w: 2.0268, h: 2.3453, fill: { color: C.orange }, rotate: 270 });
  logoNav(s);
  const STAR_DX = [0, 0.2539, 0.5157, 0.7754, 1.0235];
  [1.3133, 5.388].forEach((x) => {
    card(s, x, 2.9566, 3.2206, 1.8538);
    STAR_DX.forEach((dx) =>
      s.addShape('star5', { x: x + 0.77 + dx, y: 3.5085, w: 0.1875, h: 0.1875, fill: { color: C.star } }));
    text(s, 'Lorence. A', { x: x + 0.6973, y: 3.1689, w: 1.2837, h: 0.2524, fontFace: F.ui, fontSize: 11, bold: true });
    para(s, x + 0.6973, 3.8835, 2.2708, 0.6059,
      '\u201CCapture the beauty of your travels through the lens, preserving moments\u201D', { color: C.dark });
  });
  heading(s, 0.9092, 1.1779, 4.5825, 0.8835, 'About Testimonial\nTrove Travel');
  para(s, 5.5383, 1.3168, 3.2206, 0.6059, CAPTURE, { color: C.dark });
  photo(s, 5.1086, 2.6544, 0.8542, 0.8542, { shape: 'ellipse' });
  photo(s, 1.0338, 2.6544, 0.8542, 0.8542, { shape: 'ellipse' });
}

function slide23(s) {                                     // contact
  bgRect(s);
  card(s, 7.0439, 1.0147, 2.4285, 2.5078);
  photo(s, 7.1365, 1.0875, 2.2431, 2.3621);
  logoNav(s);
  card(s, 3.9524, 2.6527, 4.1999, 2.5078, { fill: C.orange, flat: true });
  heading(s, 0.5277, 1.3627, 3.5252, 0.4796, 'Get In Touch');
  para(s, 0.5277, 2.6671, 2.1364, 0.7826, CAPTURE, { color: C.dark });
  chipDestination(s, 1.9149, 3.9066);
  const columns = [
    { x: 4.3086, y: 3.1482, title: 'Website', lines: [[4.3086, 3.4497, 'www.youraddress.com'], [4.3086, 3.6059, 'www.traveliane.com']] },
    { x: 6.5117, y: 3.1482, title: 'Telephone', lines: [[6.5664, 3.4497, '+098 97897 9089'], [6.5664, 3.6059, '+979 89692 9878']] },
    { x: 4.3086, y: 4.1169, title: 'Email', lines: [[4.3086, 4.3794, 'Traveliane@gmail.com'], [4.3086, 4.5356, 'yourname@traveliane.com']] },
    { x: 6.5117, y: 4.1169, title: 'Application', lines: [[6.5117, 4.4155, 'www.youraddress.com'], [6.5117, 4.5718, 'www.Traveliane.com']] },
  ];
  columns.forEach((c) =>
    text(s, c.title, { x: c.x, y: c.y, w: 1.0131, h: 0.2524, fontFace: F.head, fontSize: 11 }));
  columns.forEach((c) => c.lines.forEach(([x, y, str]) =>
    text(s, str, { x, y, w: 1.4714, h: 0.1893, fontSize: 7 })));
}

function slide24(s) {                                     // thanks
  bgRect(s);
  s.addShape('round2SameRect', { x: 7.624, y: 3.249, w: 1.8301, h: 2.9219, fill: { color: C.orange }, rotate: 270 });
  logoNav(s);
  text(s, 'Thanks Slide.', { x: 0.4733, y: 1.6204, w: 8.3854, h: 1.2875, fontFace: F.ui, fontSize: 72, bold: true });
  caption(s, 0.6043, 2.9785, 4.5981, 0.2787, 'Trove Travel Presentation');
  chipClient(s, 1.7077, 0.8873);
  card(s, 6.1811, 3.1564, 3.1986, 1.5764);
  photo(s, 6.2738, 3.2291, 3.0132, 1.4561);
  chipDestination(s, 5.2024, 4.1747);
}

/* -------------------------------------------------------------------- main */

const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
  slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
  slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24,
];

const pres = new PptxGenJS();
pres.defineLayout({ name: 'TROVE', width: 10, height: 5.625 });
pres.layout = 'TROVE';
pres.author = 'Trove';
pres.title = 'Trove Travel Presentation';

BUILDERS.forEach((build) => build(pres.addSlide()));

pres.writeFile({ fileName: path.join(__dirname, '03f83b15-59f0-468c-89e0-b0c879612a52_grok_final.pptx') })
  .then((f) => console.log('wrote', f));
