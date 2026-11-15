/**
 * "YOUR STORE" new-year sale story deck — 10 portrait slides (11.25in x 20in).
 * Recreated with pptxgenjs only. Photographs in the source deck are stood in
 * for by flat dark placeholder panels (see `photo()`).
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Deck constants
 * ------------------------------------------------------------------ */

const SLIDE_W = 11.25;
const SLIDE_H = 20;

const C = {
  red: 'E54142',
  blue: '2E4CAE',
  blueMid: '215FC3',
  blueLight: '385AC6',
  orange: 'FF7049',
  white: 'FFFFFF',
  offWhite: 'F3F1F2',
  black: '000000',
  photo: '3E3D3E', // average tone of the replaced photographs
  photoLabel: '4A494B', // barely-there caption, like the watermark it replaces
};

const F = {
  mont: 'Montserrat',
  montMed: 'Montserrat Medium',
  montLight: 'Montserrat Light',
  pop: 'Poppins',
  popXB: 'Poppins ExtraBold',
  popLight: 'Poppins Light',
  sans: 'Open Sans',
};

/* ------------------------------------------------------------------ *
 * Small helpers
 * ------------------------------------------------------------------ */

// Every text box in the source uses Google-Slides insets and top anchoring.
const INSET = { margin: [7.2, 7.2, 3.6, 3.6] }; // = 91425 / 45700 EMU

function txt(slide, content, o) {
  slide.addText(content, Object.assign(
    { color: C.white, fontFace: F.mont, valign: 'top', align: 'left', ...INSET },
    o
  ));
}

function rect(slide, o) {
  slide.addShape('rect', o);
}

// Solid stand-in for a photograph, with a discreet caption.
// `label: null` leaves a full-bleed panel bare so headline copy stays legible.
function photo(slide, o) {
  const { label = '[image]', shape = 'rect', ...pos } = o;
  slide.addShape(shape, { ...pos, fill: { color: C.photo } });
  if (!label) return;
  slide.addText(label, {
    x: pos.x, y: pos.y + pos.h / 2 - 0.45, w: pos.w, h: 0.9,
    align: 'center', valign: 'middle', color: C.photoLabel,
    fontFace: F.mont, fontSize: Math.max(14, Math.min(40, pos.w * 4)),
  });
}

function hline(slide, x, y, w, width) {
  slide.addShape('line', { x, y, w, h: 0, line: { color: C.white, width } });
}

/**
 * pptxgenjs only exposes a preset shape's "adj" guide through `rectRadius`,
 * which it converts back via adj = rectRadius / min(w, h). Undo that here so
 * call sites can state the OOXML adj value from the source deck.
 */
function adj(val, w, h) {
  return (val / 100000) * Math.min(w, h);
}

// Turn a path in 0..1 shape-local units into pptxgenjs custGeom points.
function pathPoints(cmds, w, h) {
  return cmds.map(([op, ...n]) => {
    if (op === 'Z') return { close: true };
    if (op === 'M') return { x: n[0] * w, y: n[1] * h, moveTo: true };
    if (op === 'L') return { x: n[0] * w, y: n[1] * h };
    return {
      x: n[4] * w, y: n[5] * h,
      curve: { type: 'cubic', x1: n[0] * w, y1: n[1] * h, x2: n[2] * w, y2: n[3] * h },
    };
  });
}

/* ------------------------------------------------------------------ *
 * Reusable artwork
 * ------------------------------------------------------------------ */

// Shopping-bag pictogram (custom geometry, normalised to its bounding box).
const BAG_PATH = [
  ['M', 0.8168, 0.9000], ['L', 0.8590, 0.4159], ['L', 0.9028, 0.8375], ['L', 0.8168, 0.9000], ['Z'],
  ['M', 0.7197, 0.9318], ['L', 0.0939, 0.9318], ['L', 0.1487, 0.2955], ['L', 0.2175, 0.2955],
  ['L', 0.2175, 0.3864], ['C', 0.2175, 0.3989, 0.2316, 0.4091, 0.2488, 0.4091],
  ['C', 0.2660, 0.4091, 0.2801, 0.3989, 0.2801, 0.3864], ['L', 0.2801, 0.2955], ['L', 0.5930, 0.2955],
  ['L', 0.5930, 0.3864], ['C', 0.5930, 0.3989, 0.6071, 0.4091, 0.6243, 0.4091],
  ['C', 0.6415, 0.4091, 0.6556, 0.3989, 0.6556, 0.3864], ['L', 0.6556, 0.2955], ['L', 0.7745, 0.2955],
  ['L', 0.7197, 0.9318], ['Z'],
  ['M', 0.2801, 0.1818], ['C', 0.2801, 0.1295, 0.3286, 0.0864, 0.3943, 0.0727],
  ['C', 0.3709, 0.0977, 0.3583, 0.1273, 0.3583, 0.1591], ['L', 0.3583, 0.2273], ['L', 0.2801, 0.2273],
  ['L', 0.2801, 0.1818], ['Z'],
  ['M', 0.4772, 0.0716], ['C', 0.5445, 0.0841, 0.5930, 0.1284, 0.5930, 0.1818], ['L', 0.5930, 0.2273],
  ['L', 0.4209, 0.2273], ['L', 0.4209, 0.1591], ['C', 0.4209, 0.1239, 0.4428, 0.0932, 0.4772, 0.0716],
  ['Z'],
  ['M', 0.5774, 0.0455], ['C', 0.6634, 0.0455, 0.7338, 0.0966, 0.7338, 0.1591], ['L', 0.7338, 0.2273],
  ['L', 0.6556, 0.2273], ['L', 0.6556, 0.1818], ['C', 0.6556, 0.1250, 0.6149, 0.0750, 0.5523, 0.0466],
  ['C', 0.5602, 0.0466, 0.5696, 0.0455, 0.5774, 0.0455], ['Z'],
  ['M', 0.9372, 0.2705], ['C', 0.9357, 0.2466, 0.9075, 0.2273, 0.8746, 0.2273], ['L', 0.8074, 0.2273],
  ['C', 0.8387, 0.2273, 0.8653, 0.2443, 0.8699, 0.2659],
  ['C', 0.8653, 0.2443, 0.8402, 0.2273, 0.8074, 0.2273], ['L', 0.7964, 0.2273], ['L', 0.7964, 0.1591],
  ['C', 0.7964, 0.0716, 0.6978, 0.0000, 0.5774, 0.0000],
  ['C', 0.5351, 0.0000, 0.4960, 0.0091, 0.4632, 0.0239],
  ['C', 0.4538, 0.0227, 0.4460, 0.0227, 0.4366, 0.0227],
  ['C', 0.3161, 0.0227, 0.2175, 0.0943, 0.2175, 0.1818], ['L', 0.2175, 0.2273], ['L', 0.1205, 0.2273],
  ['C', 0.0877, 0.2273, 0.0595, 0.2455, 0.0579, 0.2705], ['L', 0.0001, 0.9523],
  ['C', -0.0015, 0.9784, 0.0267, 1.0000, 0.0626, 1.0000], ['L', 0.7510, 1.0000], ['L', 0.7870, 1.0000],
  ['C', 0.8042, 1.0000, 0.8214, 0.9943, 0.8324, 0.9852], ['L', 0.9810, 0.8784],
  ['C', 0.9935, 0.8693, 1.0014, 0.8557, 0.9998, 0.8432], ['L', 0.9372, 0.2705], ['Z'],
];

// Rounded "petal" blob: flat right/bottom edges, one big rounded corner.
const BLOB_PATH = [
  ['M', 0.7988, 0.0000],
  ['C', 0.8677, 0.0000, 0.9346, 0.0108, 0.9984, 0.0310],
  ['L', 1.0000, 0.0315], ['L', 1.0000, 1.0000], ['L', 0.0003, 1.0000], ['L', 0.0000, 0.9837],
  ['C', 0.0000, 0.4404, 0.3576, 0.0000, 0.7988, 0.0000],
  ['Z'],
];

const BAG_W = 0.5134;
const BAG_H = 0.7068;

// "YOUR / STORE" lock-up. (x, y) is the top-left of the bag pictogram.
function logo(slide, x, y) {
  slide.addShape('custGeom', {
    x, y, w: BAG_W, h: BAG_H,
    fill: { color: C.white }, line: { type: 'none' },
    points: pathPoints(BAG_PATH, BAG_W, BAG_H),
  });
  txt(slide, [
    { text: 'YOUR', options: { breakLine: true } },
    { text: 'STORE', options: { bold: true } },
  ], { x: x + 0.5723, y: y + 0.0432, w: 1.1854, h: 0.7742, fontSize: 20 });
}

/* ------------------------------------------------------------------ *
 * Shared copy
 * ------------------------------------------------------------------ */

const LOREM_LONG =
  'Lorem ipsum dolor sit consectetur adipiscing elit. Pellentesque amet, ' +
  'consectetur adipiscing elit. ';
const LOREM_SHORT =
  'Lorem ipsum dolor sit consectetur adipiscing elit. Pellentesque amet, ';

// The body copy is always Open Sans 20pt on 1.5 line spacing.
const BODY = { fontFace: F.sans, fontSize: 20, lineSpacingMultiple: 1.5 };

// Fresh run objects each call — pptxgenjs writes resolved options back into them.
const address = (city = 'Melbourne, AUS') => ([
  { text: '25 Evlyn Street.', options: { bold: true, breakLine: true } },
  { text: city },
]);

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

function slide1(pres) {
  const s = pres.addSlide();
  rect(s, { x: 0, y: 0.0043, w: SLIDE_W, h: 19.9957, fill: { color: C.red } });
  txt(s, 'NEW YEAR EVENT',
    { x: 1.5074, y: 2.7102, w: 8.2353, h: 0.9088, fontSize: 48, align: 'center' });
  txt(s, 'BIG SALE!',
    { x: 0, y: 3.3982, w: SLIDE_W, h: 2.4234, fontFace: F.pop, fontSize: 138, bold: true, align: 'center' });
  logo(s, 4.8168, 1.1628);
  txt(s, 'Get 50% Off All Collection',
    { x: 1.0965, y: 15.8003, w: 3.2223, h: 1.1781, fontFace: F.montMed, fontSize: 32 });
  txt(s, '01. 01. 21',
    { x: 7.9921, y: 15.8006, w: 2.0079, h: 0.6395, fontFace: F.montMed, fontSize: 32, align: 'right' });
  hline(s, 1.228, 15.5459, 8.772, 0.75);
  txt(s, 'SWIPE  FOR  SHOP',
    { x: 1.0965, y: 17.9002, w: 8.9035, h: 0.8415, fontFace: F.montLight, fontSize: 44, align: 'center' });
  hline(s, 1.1491, 17.7706, 8.772, 0.75);
  photo(s, { x: 1.2396, y: 6.276, w: 8.7708, h: 8.3559 });
}

function slide2(pres) {
  const s = pres.addSlide();
  rect(s, { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: { color: C.blue } });
  s.addShape('ellipse',
    { x: 3.7574, y: 5.2886, w: 4.7111, h: 11.1834, rotate: -40.594, fill: { color: C.red } });
  s.addShape('custGeom', {
    x: 5.625, y: 15.1579, w: 5.6049, h: 4.8421, fill: { color: C.red },
    points: pathPoints(BLOB_PATH, 5.6049, 4.8421),
  });
  txt(s, [
    { text: 'NEW YEAR', options: { breakLine: true } },
    { text: 'SALE' },
  ], { x: 1.0647, y: 2.5188, w: 9.1206, h: 7.0684, fontFace: F.popXB, fontSize: 138, bold: true });
  txt(s, 'Get 50% Off All Items',
    { x: 6.8268, y: 3.183, w: 2.6115, h: 1.9186, fontSize: 36, bold: true });
  hline(s, 6.9483, 2.9344, 2.7253, 1);
  logo(s, 1.1862, 1.1616);
  txt(s, LOREM_LONG, { x: 1.1183, y: 15.813, w: 7.0142, h: 1.057, ...BODY });
  txt(s, 'SWIPE  UP NOW!', { x: 1.0647, y: 17.5789, w: 8.9035, h: 0.8415, fontSize: 44 });
  photo(s, { x: 1.1858, y: 10.6719, w: 5.7622, h: 4.4323 });
  photo(s, { x: 6.9479, y: 7.3924, w: 4.3021, h: 3.2795 });
}

function slide3(pres) {
  const s = pres.addSlide();
  rect(s, { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: { color: C.blue } });
  txt(s, 'FLASH SALE!',
    { x: -0.0184, y: 3.2956, w: SLIDE_W, h: 2.0364, fontFace: F.popXB, fontSize: 115, align: 'center' });
  txt(s, 'ONLY THIS DAY UNTIL 7PM',
    { x: 2.6239, y: 5.0287, w: 6.0358, h: 0.5722, fontSize: 28, align: 'center' });
  logo(s, 4.8168, 1.8471);
  txt(s, LOREM_LONG, { x: 1.53, y: 17.096, w: 8.1533, h: 1.057, align: 'center', ...BODY });
  photo(s, { x: 0, y: 6.8142, w: SLIDE_W, h: 6.9219 });

  // Three concentric outlines behind a solid red badge.
  const ring = { w: 9.3967, h: 3.4053, fill: { type: 'none' }, line: { color: C.white, width: 3 } };
  s.addShape('ellipse', { ...ring, x: 0.9083, y: 11.5841, h: 3.1249 });
  s.addShape('ellipse', { ...ring, x: 0.9435, y: 12.9587, h: 2.9631 });
  s.addShape('ellipse', { ...ring, x: 0.9259, y: 13.5572, h: 2.9631 });
  s.addShape('ellipse', { x: 0.9435, y: 12.1053, w: 9.3967, h: 3.4053, fill: { color: C.red } });
  txt(s, '80% OFF',
    { x: 1.8402, y: 12.9459, w: 7.7701, h: 2.0364, fontFace: F.pop, fontSize: 115, align: 'center' });
}

function slide4(pres) {
  const s = pres.addSlide();
  photo(s, { x: 0, y: 0, w: SLIDE_W, h: 9.4479 });
  rect(s, { x: 0, y: 9.4474, w: SLIDE_W, h: 10.5526, fill: { color: C.blueMid } });
  s.addShape('custGeom', {
    x: 3.0264, y: 7.3027, w: 6.0789, h: 10.3684, rotate: -90, fill: { color: C.orange },
    points: pathPoints(BLOB_PATH, 6.0789, 10.3684),
  });
  txt(s, '50%',
    { x: 3.2296, y: 10.6216, w: 8.2321, h: 3.45, fontFace: F.popXB, fontSize: 199, align: 'center' });
  txt(s, 'NEW STORE SALE!',
    { x: 0.8816, y: 15.7122, w: 6.2072, h: 2.3225, fontFace: F.popXB, fontSize: 66, bold: true });
  txt(s, address(), { x: 3.9852, y: 17.0791, w: 2.629, h: 0.7742, fontSize: 20 });
  txt(s, 'GET', { x: 2.6245, y: 11.0238, w: 2.5983, h: 1.1107, fontFace: F.popXB, fontSize: 60 });
  txt(s, 'OFF',
    { x: 8.3403, y: 13.2867, w: 2.028, h: 1.1107, fontFace: F.popXB, fontSize: 60, align: 'right' });
  txt(s, LOREM_SHORT, { x: 7.4072, y: 16.2982, w: 3.2402, h: 1.5618, ...BODY });
  logo(s, 1.1336, 1.3985);
}

function slide5(pres) {
  const s = pres.addSlide();
  rect(s, { x: 0, y: 0.0043, w: SLIDE_W, h: 19.9957, fill: { color: C.orange } });
  txt(s, 'NEW ARRIVAL!',
    { x: 0, y: 2.7199, w: SLIDE_W, h: 1.582, fontFace: F.pop, fontSize: 88, bold: true, align: 'center' });
  logo(s, 4.8168, 1.1628);
  txt(s, 'Get 50% Off',
    { x: 2.4348, y: 4.1721, w: 6.433, h: 0.6395, fontFace: F.montMed, fontSize: 32, align: 'center' });

  // Three product cards, each with a rotated tab and a brand label.
  const brands = [
    { y: 5.5521, tabY: 7.0191, labelX: 6.3527, labelY: 8.5832, name: 'BALENCIANAGA' },
    { y: 9.8299, tabY: 11.2399, labelX: 6.3527, labelY: 12.8172, name: 'EVRI TOMAN' },
    { y: 14.1076, tabY: 15.5338, labelX: 6.4778, labelY: 17.1038, name: 'WADIMOR' },
  ];
  brands.forEach(b => photo(s, { x: 1.5434, y: b.y, w: 8.217, h: 3.5382 }));
  brands.forEach(b => {
    s.addShape('trapezoid', {
      x: 7.4017, y: b.tabY, w: 1.1838, h: 3.5321, rotate: -90,
      fill: { color: C.blue }, rectRadius: adj(17857, 1.1838, 3.5321),
    });
    txt(s, b.name,
      { x: b.labelX, y: b.labelY, w: 3.2107, h: 0.4376, fontSize: 20, bold: true, align: 'center' });
  });
}

function slide6(pres) {
  const s = pres.addSlide();
  rect(s, { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: { color: C.blue } });
  s.addShape('round1Rect', {
    x: 3.7991, y: 7.1, w: 4.8111, h: 8.874, rotate: -120,
    fill: { color: C.orange }, rectRadius: adj(50000, 4.8111, 8.874),
  });
  s.addShape('ellipse',
    { x: -0.1683, y: 2.6727, w: 10.2756, h: 4.483, rotate: 15.588, fill: { color: C.red } });
  txt(s, 'ALWAYS WEAR WHAT YOU WANT TO WEAR',
    { x: 1.0768, y: 3.605, w: 8.8179, h: 11.7133, fontFace: F.pop, fontSize: 115, bold: true, align: 'center' });
  txt(s, address(), { x: 1.0768, y: 16.6921, w: 2.6235, h: 0.7068, fontSize: 18 });
  txt(s, 'www.store.com', { x: 4.1419, y: 16.6928, w: 3.1776, h: 0.5049, fontSize: 24, bold: true });
  txt(s, '@yourstore',
    { x: 7.6462, y: 16.6845, w: 2.3538, h: 0.5049, fontSize: 24, align: 'right' });
  txt(s, 'SWIPE    UP    NOW!',
    { x: 1.0965, y: 17.9002, w: 8.9035, h: 0.8415, fontSize: 44, bold: true, align: 'justify' });
  hline(s, 1.1491, 17.7706, 8.772, 0.75);
  logo(s, 4.8168, 1.1628);
}

function slide7(pres) {
  const s = pres.addSlide();
  rect(s, { x: 0, y: 0, w: SLIDE_W, h: 4.6989, fill: { color: C.red } });
  txt(s, 'BIG SALE!',
    { x: 0.7646, y: 1.8154, w: 7.7582, h: 2.0364, fontFace: F.popXB, fontSize: 115 });
  rect(s, { x: 0, y: 4.6989, w: SLIDE_W, h: 2.8537, fill: { color: C.offWhite } });
  txt(s, [
    { text: 'Get 50% Off  ', options: { bold: true } },
    { text: 'All Collection' },
  ], { x: 0.7646, y: 5.4727, w: 5.933, h: 0.6395, fontFace: F.pop, fontSize: 32, color: C.black });
  txt(s, '01. 01. 21', {
    x: 7.6987, y: 5.4275, w: 2.5502, h: 0.7068,
    fontFace: F.montMed, fontSize: 36, color: C.black, align: 'right',
  });
  logo(s, 8.7171, 2.565);
  txt(s, 'Lorem ipsum dolor sit amet',
    { x: 0.7693, y: 6.2383, w: 8.4612, h: 0.5521, color: C.black, ...BODY });
  photo(s, { x: 0, y: 7.5521, w: SLIDE_W, h: 12.4479 });
}

function slide8(pres) {
  const s = pres.addSlide();
  photo(s, { x: 0, y: 0, w: SLIDE_W, h: 10, label: null });
  photo(s, { x: -0.016, y: 10, w: SLIDE_W, h: 10, label: null });
  s.addShape('ellipse', { x: 0, y: 4.375, w: 11.25, h: 11.25, fill: { color: C.blueLight } });
  s.addShape('ellipse', { x: 1.4438, y: 5.8188, w: 8.3626, h: 8.3626, fill: { color: C.blue } });
  txt(s, '01.01', { x: 3.3516, y: 6.5095, w: 4.3778, h: 2.0364, fontSize: 115, align: 'center' });
  txt(s, 'BIG SALE!',
    { x: 1.7729, y: 8.2834, w: 7.7041, h: 1.7166, fontFace: F.pop, fontSize: 96, bold: true, align: 'center' });
  txt(s, 'Get 50% Off All Items',
    { x: 2.7169, y: 10.6907, w: 5.8162, h: 1.4473, fontSize: 40, align: 'center' });
  txt(s, [
    { text: 'Shop now at', options: { breakLine: true } },
    { text: 'www.yourstore.com', options: { bold: true } },
  ], { x: 3.627, y: 13.3453, w: 3.996, h: 0.9088, fontSize: 24, align: 'center' });
  s.addShape('ellipse', { x: 2.3994, y: 16.6459, w: 6.4512, h: 1.7166, fill: { color: C.red } });
  txt(s, 'SWIPE UP!',
    { x: 3.221, y: 17.1171, w: 4.808, h: 0.7742, fontFace: F.pop, fontSize: 40, align: 'center' });
}

function slide9(pres) {
  const s = pres.addSlide();
  rect(s, { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: { color: C.red } });
  s.addShape('parallelogram', {
    x: 0, y: 0, w: SLIDE_W, h: 10.4474,
    fill: { color: C.blue }, rectRadius: adj(52967, SLIDE_W, 10.4474),
  });
  txt(s, 'GARAGE SALE -',
    { x: 0.7744, y: 2.5155, w: 7.4098, h: 3.9717, fontFace: F.pop, fontSize: 115, bold: true });
  txt(s, address('MelbourneAUS'),
    { x: 8.2843, y: 4.6624, w: 2.0587, h: 1.7166, fontSize: 24 });
  txt(s, '80% OFF',
    { x: 5.1125, y: 7.6485, w: 5.2306, h: 1.582, fontFace: F.popLight, fontSize: 88 });
  logo(s, 0.9422, 1.3542);
  txt(s, LOREM_SHORT, { x: 0.9069, y: 7.4962, w: 3.2402, h: 1.5618, ...BODY });
  photo(s, {
    x: -0.016, y: 10.4479, w: SLIDE_W, h: 9.5521,
    shape: 'parallelogram', rectRadius: adj(51844, SLIDE_W, 9.5521),
  });
}

function slide10(pres) {
  const s = pres.addSlide();
  photo(s, { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, label: null });
  rect(s, { x: 0, y: 0, w: 3.2368, h: 4.4474, fill: { color: C.red } });
  logo(s, 0.7843, 2.4595);
  txt(s, 'MANCING PERKORO',
    { x: 1.041, y: 8.0141, w: 8.3323, h: 3.9717, fontFace: F.pop, fontSize: 115, bold: true });
  txt(s, LOREM_SHORT, { x: 1.1199, y: 12.7593, w: 3.2402, h: 1.5618, ...BODY });
  txt(s, 'Get 50% Off All Collection',
    { x: 1.1199, y: 16.8842, w: 3.8509, h: 1.3127, fontFace: F.montMed, fontSize: 36 });
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */

const pres = new PptxGenJS();
pres.defineLayout({ name: 'STORY', width: SLIDE_W, height: SLIDE_H });
pres.layout = 'STORY';

[slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10]
  .forEach(build => build(pres));

pres.writeFile({
  fileName: path.join(__dirname, '0f3cda14-b978-4a2f-8a7d-64ac521b2dab_grok_final.pptx'),
});
