/**
 * "Miela." - Furniture Catalogue deck (30 slides, 13.333 x 7.5 in / 16:9).
 *
 * A faithful pptxgenjs rebuild of the original PowerPoint template: a dark
 * cocoa background, mustard accent blocks, an 8-pointed asterisk motif and
 * grey photo panels.  Photographs in the source are stand-in "insert your
 * picture here" frames; they are re-created here as plain grey placeholder
 * rectangles (no embedded bitmaps).
 *
 * Run:  node e31ae818-5a03-42a1-9b54-42a909ca10f4_grok_final.js
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const BROWN = '3A2C00';   // background / text on yellow  (theme bg1, tx2)
const CREAM = 'FFFFFF';   // body text on brown            (theme tx1, bg2)
const YELLOW = 'FFD965';  // accent blocks                 (theme accent1)
const PANEL = 'E0E0E0';   // photo placeholder fill
const PANEL_LINE = 'CCCCCC';
const PANEL_TEXT = '333333';

const HEAD = 'Plus Jakarta Sans ExtraBold';  // theme major font
const BODY = 'Montserrat';                   // theme minor font

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

/* Text styles used across the deck (see txt() below). */
const DISPLAY = { size: 75, face: HEAD, color: CREAM };          // page headlines
const HEADING = { size: 20, face: HEAD, color: CREAM };          // section / product titles
const ON_YELLOW = { size: 20, face: HEAD, color: BROWN };        // titles sitting on a yellow block
const PRICE = { size: 20, face: BODY, color: CREAM, lh: 1.5, align: 'right' };
const BODY_TEXT = { size: 12, face: BODY, color: CREAM, lh: 1.5 };
const STEP_NUM = { size: 20, face: HEAD, color: CREAM, lh: 1.5 };  // "01." markers
const STEP_NAME = { size: 12, face: HEAD, color: CREAM, lh: 1.5 };

/* Boilerplate copy, reused verbatim on many slides. */
const SERENITY_ALONE =
  'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. \u00A0I am alone, and feel the charm of existence in this spot.';
const SERENITY_ALONE_BLISS =
  'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. \u00A0I am alone, and feel the charm of existence in this spot. which was created for the bliss of souls like mine.';
const SERENITY_ALONE_PLAIN =
  'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence in this spot.';
const SERENITY =
  'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart.';
const HAPPY_TALENTS =
  '\u00A0I am so happy, my dear friend, so absorbed in the exquisite sense of mere tranquil existence, that I neglect my talents. ';
const HAPPY_ARTIST =
  '\u00A0I am so happy, my dear friend, so absorbed in the exquisite sense of mere tranquil existence, that I neglect my talents. I should be incapable of drawing a single stroke at the present moment; and yet I feel that I never was a greater artist than now.';
const HAPPY_DOT = '\u00A0I am so happy, my dear friend, so absorbed in the exquisite sense of mere tranquil existence.';
const HAPPY_TALENTS_2 =
  'I am so happy, my dear friend, so absorbed in the exquisite sense of mere tranquil existence, that I neglect my talents. ';
const HAPPY_SHORT = 'I am so happy, my dear friend, so absorbed.';
const HAPPY = 'I am so happy, my dear friend, so absorbed in the exquisite sense of mere tranquil existence';
const DETAILS = 'Product Details';

/* The 8-pointed asterisk logo mark, as a closed path normalised to a 1x1 box. */
const STAR_PATH = [
  [1.0000, 0.4585], [0.6000, 0.4585], [0.8829, 0.1757], [0.8243, 0.1171],
  [0.5415, 0.4000], [0.5415, 0.0000], [0.4585, 0.0000], [0.4585, 0.4000],
  [0.1757, 0.1171], [0.1171, 0.1757], [0.4000, 0.4585], [0.0000, 0.4585],
  [0.0000, 0.5415], [0.4000, 0.5415], [0.1171, 0.8243], [0.1757, 0.8829],
  [0.4585, 0.6000], [0.4585, 1.0000], [0.5415, 1.0000], [0.5415, 0.6000],
  [0.8243, 0.8829], [0.8829, 0.8243], [0.6000, 0.5415], [1.0000, 0.5415]
];

/* --------------------------------------------------------------- helpers */

/** New slide: brown background plus the yellow "0NN / 030" page tab. */
function newSlide(pres, n) {
  const s = pres.addSlide();
  s.background = { color: BROWN };
  s.addShape(pres.ShapeType.rect, { x: 0, y: 0.623, w: 2.243, h: 0.382, fill: { color: YELLOW }, line: { type: 'none' } });
  s.addText(String(n).padStart(3, '0') + ' / 030', {
    x: n <= 9 ? 0.917 : 0.964, y: 0.613, w: 1.131, h: 0.399,
    fontFace: BODY, fontSize: 12, color: BROWN, valign: 'middle',
  });
  return s;
}

/** Mustard accent block. `rotate` is the PowerPoint rotation in degrees. */
function block(s, x, y, w, h, rotate) {
  s.addShape('rect', { x: x, y: y, w: w, h: h, rotate: rotate || 0, fill: { color: YELLOW }, line: { type: 'none' } });
}

/** Photo placeholder: grey panel, faint corner-to-corner cross, "[image]" caption. */
function photo(s, x, y, w, h) {
  s.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: PANEL }, line: { type: 'none' } });
  s.addShape('line', { x: x, y: y, w: w, h: h, line: { color: PANEL_LINE, width: 0.75 } });
  s.addShape('line', { x: x, y: y, w: w, h: h, flipV: true, line: { color: PANEL_LINE, width: 0.75 } });
  s.addText('[image]', {
    x: x, y: y, w: w, h: h, align: 'center', valign: 'middle',
    fontFace: BODY, fontSize: 14, bold: true, color: PANEL_TEXT,
  });
}

/** Thin horizontal / vertical hairline (the deck's price separators). */
function rule(s, x, y, w, color) {
  s.addShape('line', { x: x, y: y, w: w, h: 0, line: { color: color, width: 0.5 } });
}

function vRule(s, x, y, h, color) {
  s.addShape('line', { x: x, y: y, w: 0, h: h, line: { color: color, width: 0.5 } });
}

/** Asterisk motif, drawn from STAR_PATH at the requested square size. */
function star(s, x, y, size) {
  s.addShape('custGeom', {
    x: x, y: y, w: size, h: size, fill: { color: CREAM }, line: { type: 'none' },
    points: STAR_PATH.map(function (p, i) {
      return i === 0 ? { x: p[0] * size, y: p[1] * size, moveTo: true } : { x: p[0] * size, y: p[1] * size };
    }).concat([{ close: true }]),
  });
}

/**
 * Text box: one array entry per paragraph.
 * `style` is one of the presets above; `over` overrides individual properties.
 */
function txt(s, style, x, y, w, h, lines, over) {
  const o = Object.assign({}, style, over || {});
  s.addText(lines.map(function (line) { return { text: line, options: { breakLine: true } }; }), {
    x: x, y: y, w: w, h: h,
    fontFace: o.face, fontSize: o.size, color: o.color,
    align: o.align || 'left', valign: 'top',
    lineSpacingMultiple: o.lh || null,
  });
}

/* ---------------------------------------------------------------- slides */


/* 01 - Miela. */
function slide01(s) {
  block(s, 9.806, 3.798, 3.527, 3.75);
  photo(s, 7.446, 0.613, 4.729, 6.323);
  block(s, 6.508, 0.001, 1.877, 3.386);
  txt(s, DISPLAY, 0.825, 4.737, 4.475, 1.784, ['Miela.'], { size: 100 });
  txt(s, PRICE, 0.875, 6.486, 3.246, 0.438, ['Furniture Catalogue'], { align: 'left', lh: null });
  block(s, 11.786, 0.606, 0.78, 0.661);
  star(s, 1.002, 1.978, 0.843);
  txt(s, ON_YELLOW, 6.654, 2.434, 1.877, 0.774, ['Luxurious', 'Collection']);
}

/* 02 - Introduction */
function slide02(s) {
  block(s, 9.806, 5.323, 3.527, 2.177);
  txt(s, DISPLAY, 0.946, 1.278, 6.635, 1.363, ['Introduction']);
  photo(s, 6.065, 2.969, 6.339, 3.951);
  block(s, 4.383, 3.264, 3.061, 1.047);
  txt(s, BODY_TEXT, 0.946, 4.97, 4.746, 1.581, [HAPPY_ARTIST]);
  txt(s, ON_YELLOW, 4.512, 3.386, 1.665, 0.774, ['Luxurious', 'Collection']);
  rule(s, 8.667, 2.075, 3.737, CREAM);
  star(s, 1.002, 3.351, 0.843);
}

/* 03 - Table Of Content */
function slide03(s) {
  photo(s, 6.667, 0, 6.667, 3.75);
  txt(s, DISPLAY, 0.946, 1.617, 4.769, 2.625, ['Table Of', 'Content']);
  block(s, 9.944, 3.273, 3.39, 1.019);
  txt(s, BODY_TEXT, 0.946, 5.967, 1.626, 0.976, [HAPPY_SHORT]);
  txt(s, BODY_TEXT, 3.367, 5.967, 1.626, 0.976, [HAPPY_SHORT]);
  txt(s, BODY_TEXT, 5.861, 5.967, 1.626, 0.976, [HAPPY_SHORT]);
  txt(s, BODY_TEXT, 8.318, 5.967, 1.626, 0.976, [HAPPY_SHORT]);
  txt(s, BODY_TEXT, 10.777, 5.967, 1.626, 0.976, [HAPPY_SHORT]);
  txt(s, STEP_NUM, 0.946, 5.09, 0.803, 0.539, ['01.']);
  txt(s, STEP_NAME, 0.946, 5.525, 1.471, 0.364, ['Product Insight']);
  txt(s, STEP_NUM, 3.367, 5.09, 0.729, 0.539, ['02.']);
  txt(s, STEP_NAME, 3.367, 5.525, 1.206, 0.364, ['Product Set']);
  txt(s, STEP_NUM, 5.861, 5.09, 0.729, 0.539, ['03.']);
  txt(s, STEP_NAME, 5.861, 5.525, 1.626, 0.364, ['Room Accessory']);
  txt(s, STEP_NUM, 8.303, 5.09, 0.729, 0.539, ['04.']);
  txt(s, STEP_NAME, 8.303, 5.525, 1.471, 0.364, ['Special Edition']);
  txt(s, STEP_NUM, 10.785, 5.09, 0.729, 0.539, ['05.']);
  txt(s, STEP_NAME, 10.785, 5.525, 1.777, 0.364, ['Limited Time Offer']);
  rule(s, 1.604, 5.418, 1.618, CREAM);
  rule(s, 4.096, 5.418, 1.618, CREAM);
  rule(s, 6.59, 5.418, 1.618, CREAM);
  rule(s, 9.032, 5.418, 1.618, CREAM);
  txt(s, ON_YELLOW, 10.106, 3.383, 1.626, 0.774, ['Luxurious', 'Collection']);
  block(s, 6.277, 0.463, 0.78, 0.661);
  star(s, 6.277, 3.352, 0.843);
}

/* 04 - Product Insight */
function slide04(s) {
  block(s, 9.813, 3.026, 3.521, 1.305);
  photo(s, 4.635, 0, 8.698, 3.75);
  txt(s, DISPLAY, 0.834, 4.331, 4.769, 2.625, ['Product', 'Insight']);
  star(s, 5.848, 5.338, 0.843);
  txt(s, BODY_TEXT, 7.333, 6.284, 5.206, 0.673, [HAPPY_TALENTS]);
  block(s, 0.898, 2.175, 5.747, 0.781);
  txt(s, BODY_TEXT, 7.333, 5.035, 5.206, 0.974, [SERENITY]);
  txt(s, ON_YELLOW, 1.059, 2.252, 5.383, 0.539,
      ['Luxurious Comfort, Impeccable Style'], { align: 'center', lh: 1.5 });
}

/* 05 - Living Room Set */
function slide05(s) {
  photo(s, 0, 1.73, 6.667, 5.77);
  txt(s, DISPLAY, 9.07, 3.116, 3.446, 3.888, ['Living Room Set'], { align: 'right' });
  star(s, 11.619, 1.424, 0.843);
  block(s, 6.277, 1.424, 0.78, 0.661);
  block(s, 5.447, 5.384, 2.44, 1.792, 90);
  txt(s, ON_YELLOW, 5.876, 5.268, 1.811, 0.774, ['Luxurious', 'Collection']);
}

/* 06 - Living Room Set */
function slide06(s) {
  block(s, 8.887, 1.779, 4.446, 1.174);
  txt(s, HEADING, 5.789, 3.198, 1.531, 0.774, ['Living Room Set']);
  txt(s, PRICE, 11.378, 3.401, 1.085, 0.549, ['$2099']);
  rule(s, 7.92, 3.756, 2.883, CREAM);
  star(s, 0.994, 1.779, 0.843);
  txt(s, BODY_TEXT, 0.994, 5.988, 3.179, 0.976, [HAPPY_DOT]);
  txt(s, BODY_TEXT, 0.994, 4.378, 3.306, 1.279, [SERENITY]);
  txt(s, HEADING, 4.475, 1.848, 3.863, 0.774, ['Exquisite Craftsmanship, Unparalleled Luxury']);
  txt(s, HEADING, 0.994, 3.693, 2.347, 0.438, [DETAILS]);
  photo(s, 5.079, 2.952, 8.254, 4.548);
  block(s, 3.604, 5.544, 2.952, 0.959, 90);
}

/* 07 - Transform Your Space, Elevate Your Lifestyle */
function slide07(s) {
  block(s, 2.259, 5.339, 3.016, 1.306, 90);
  photo(s, 3.769, 1.371, 5.348, 6.129);
  star(s, 1.014, 1.661, 0.843);
  txt(s, HEADING, 0.982, 5.517, 2.034, 1.447, ['Transform Your Space, Elevate Your Lifestyle']);
  txt(s, BODY_TEXT, 9.39, 5.689, 3.179, 1.279, [HAPPY_TALENTS_2]);
  txt(s, BODY_TEXT, 9.39, 3.645, 2.884, 1.581, [SERENITY]);
  txt(s, HEADING, 9.39, 2.942, 2.347, 0.438, [DETAILS]);
  txt(s, HEADING, 4.198, 6.269, 1.306, 0.774, ['Modern', 'Shelf']);
  txt(s, PRICE, 7.531, 6.382, 0.916, 0.549, ['$499']);
  rule(s, 5.749, 6.656, 1.362, CREAM);
  block(s, 8.809, 1, 0.78, 0.661);
}

/* 08 - Modern 4 Seat Sofa */
function slide08(s) {
  photo(s, 7.524, 4.533, 4.918, 2.967);
  photo(s, 0, 1.403, 7.524, 3.129);
  block(s, 6.698, 6.327, 1.587, 0.76, 90);
  txt(s, STEP_NUM, 0.926, 1.59, 2.986, 0.539, ['Modern 4 Seat Sofa']);
  txt(s, PRICE, 6.162, 1.587, 0.909, 0.549, ['$799']);
  rule(s, 4.127, 1.907, 1.816, CREAM);
  txt(s, BODY_TEXT, 7.821, 2.481, 4.426, 1.581, [SERENITY_ALONE_BLISS]);
  txt(s, HEADING, 7.839, 1.797, 2.347, 0.438, [DETAILS]);
  txt(s, HEADING, 7.821, 4.822, 1.921, 0.774, ['Mini Modern Table']);
  rule(s, 9.871, 5.071, 1.235, CREAM);
  txt(s, PRICE, 11.195, 4.749, 0.909, 0.549, ['$299']);
  star(s, 11.404, 0.743, 0.843);
  block(s, 1.919, 2.256, 0.747, 4.584, 90);
  txt(s, BODY_TEXT, 0.926, 5.991, 5.574, 0.976, [SERENITY_ALONE]);
  txt(s, HEADING, 0.926, 5.359, 2.347, 0.438, [DETAILS]);
}

/* 09 - Bedroom Set */
function slide09(s) {
  photo(s, 6.667, 0, 6.667, 7.5);
  txt(s, DISPLAY, 0.89, 1.463, 5.062, 2.625, ['Bedroom', 'Set']);
  block(s, 5.123, 5.061, 3.087, 1.792, 90);
  txt(s, ON_YELLOW, 5.751, 4.676, 1.811, 0.774, ['Luxurious', 'Collection'], { align: 'center' });
  txt(s, BODY_TEXT, 0.946, 6.07, 4.295, 0.976, [HAPPY_TALENTS_2]);
  star(s, 0.959, 4.626, 0.843);
  block(s, 6.154, -0.383, 1.026, 1.792, 90);
  txt(s, HEADING, 9.826, 1.463, 2.582, 0.774, ['Miela Bedroom Set'], { align: 'right' });
  txt(s, PRICE, 7.152, 1.337, 1.181, 0.549, ['$2499']);
  rule(s, 8.638, 1.671, 1.169, CREAM);
}

/* 10 - Luxury Furnishings for Discerning Tastes */
function slide10(s) {
  photo(s, 8.222, 0, 5.111, 3.75);
  block(s, 5.869, 5.682, 2.421, 1.215, 90);
  txt(s, HEADING, 0.952, 1.875, 3.333, 0.774, ['Luxury Furnishings for Discerning Tastes']);
  txt(s, BODY_TEXT, 8.107, 4.836, 4.426, 1.279, [SERENITY_ALONE]);
  txt(s, HEADING, 8.107, 4.183, 2.347, 0.438, [DETAILS]);
  txt(s, BODY_TEXT, 8.107, 6.32, 4.426, 0.673, [HAPPY]);
  star(s, 5.824, 1.875, 0.843);
  block(s, 7.012, 0.603, 2.421, 1.215, 90);
  txt(s, HEADING, 10.885, 3.057, 1.488, 0.438, ['Armchair'], { align: 'right' });
  txt(s, PRICE, 8.406, 2.945, 1.025, 0.549, ['$299']);
  rule(s, 9.667, 3.254, 1.122, CREAM);
  txt(s, HEADING, 0.952, 6.555, 1.984, 0.438, ['Double-Bed']);
  txt(s, PRICE, 5.75, 6.443, 1.025, 0.549, ['$799'], { align: 'left' });
  rule(s, 3.1, 6.752, 2.406, CREAM);
  photo(s, 0, 4.095, 7.083, 3.405);
  block(s, 0, 3.628, 4.286, 0.774, 180);
}

/* 11 - Stand Lamp */
function slide11(s) {
  block(s, 5.813, 4.984, 0.588, 2.516, 180);
  photo(s, 6.097, 1.903, 4.274, 5.597);
  txt(s, HEADING, 6.426, 2.324, 1.122, 0.774, ['Stand Lamp']);
  txt(s, PRICE, 6.426, 3.969, 1.025, 0.549, ['$199'], { align: 'left' });
  vRule(s, 6.743, 3.279, 0.587, CREAM);
  txt(s, HEADING, 0.948, 1.793, 3.536, 1.111, ['Experience the Perfect Blend of Comfort and Elegance.']);
  block(s, 9.878, 2.192, 0.78, 0.661);
  star(s, 11.21, 5.577, 1.33);
  txt(s, BODY_TEXT, 0.948, 4.733, 4.426, 1.279, [SERENITY_ALONE]);
  txt(s, HEADING, 0.948, 4.08, 2.347, 0.438, [DETAILS]);
  txt(s, BODY_TEXT, 0.948, 6.32, 4.426, 0.673, [HAPPY]);
  block(s, 12.003, 0, 1.33, 1.91, 180);
}

/* 12 - Kitchen Set */
function slide12(s) {
  block(s, 0, 5.465, 2.809, 1);
  photo(s, 0, 1.856, 5.903, 3.922);
  txt(s, DISPLAY, 8.339, 1.807, 4.154, 2.625, ['Kitchen', 'Set'], { align: 'right' });
  txt(s, BODY_TEXT, 8.339, 5.793, 3.682, 0.976, [HAPPY]);
  block(s, 4.855, 3.938, 2.809, 1.323);
  txt(s, ON_YELLOW, 5.804, 4.195, 1.811, 0.774, ['Luxurious', 'Collection']);
  star(s, 5.432, 1.471, 0.943);
}

/* 13 - Where Quality Meets Impeccable Craftsmanship. */
function slide13(s) {
  block(s, 11.384, -0.185, 1.183, 2.75, 90);
  photo(s, 7.274, 1.371, 6.059, 5.562);
  txt(s, HEADING, 0.954, 1.748, 2.722, 1.447, ['Where Quality Meets Impeccable Craftsmanship.']);
  star(s, 5.432, 1.974, 0.943);
  txt(s, BODY_TEXT, 0.948, 4.733, 4.746, 1.279, [SERENITY_ALONE]);
  txt(s, HEADING, 0.948, 4.08, 2.347, 0.438, [DETAILS]);
  txt(s, BODY_TEXT, 0.948, 6.32, 4.426, 0.673, [HAPPY]);
  block(s, 6.667, 3.75, 1.183, 3.75);
  txt(s, HEADING, 8.148, 5.932, 1.803, 0.774, ['Miela Kitchen Set']);
  txt(s, PRICE, 11.939, 6.158, 1.106, 0.549, ['$3999']);
  rule(s, 10.22, 6.477, 1.53, CREAM);
}

/* 14 - Product Details */
function slide14(s) {
  block(s, 7.095, 1.048, 2.222, 3.523);
  photo(s, 4.175, 0, 3.786, 3.75);
  photo(s, 8.524, 1.714, 4.809, 5.786);
  block(s, 12.238, 1.384, 0.78, 0.661);
  txt(s, BODY_TEXT, 3.667, 5.685, 4.679, 1.279, [SERENITY_ALONE]);
  txt(s, HEADING, 3.667, 5.022, 2.347, 0.438, [DETAILS], { align: 'right' });
  star(s, 1.003, 6.02, 0.943);
  txt(s, HEADING, 0.946, 2.212, 2.435, 1.111, ['Your Home, Your Sanctuary, Our Expertise.']);
  txt(s, HEADING, 8.791, 6.242, 1.53, 0.774, ['Modern', 'Wastafel']);
  txt(s, PRICE, 11.504, 6.452, 0.909, 0.549, ['$499'], { align: 'left' });
  rule(s, 10.429, 6.733, 0.871, CREAM);
  block(s, 3.801, 3.426, 0.78, 0.661);
}

/* 15 - Bathroom Set */
function slide15(s) {
  photo(s, 3.542, 3.75, 9.792, 3.75);
  txt(s, DISPLAY, 0.952, 1.604, 5.715, 2.625, ['Bathroom', 'Set']);
  star(s, 11.91, 0.78, 0.943);
  block(s, 1.177, 5.693, 4.032, 1.216);
  block(s, 9.301, 3.336, 4.032, 1.022);
  txt(s, ON_YELLOW, 1.412, 5.882, 1.668, 0.774, ['Luxurious Collection']);
  txt(s, HEADING, 9.301, 2.143, 3.339, 0.774, ['Crafted with Passion, Inspired by You.']);
}

/* 16 - Product Details */
function slide16(s) {
  txt(s, BODY_TEXT, 8.267, 3.356, 4.426, 1.279, [SERENITY_ALONE]);
  txt(s, HEADING, 8.267, 2.692, 2.347, 0.438, [DETAILS], { align: 'right' });
  block(s, 0, 1.616, 3.604, 2.134);
  txt(s, HEADING, 9.356, 5.864, 2.758, 1.111, ['Designing Exquisite Comfort for Your Home']);
  photo(s, 0, 2.188, 7.871, 5.312);
  block(s, 7.532, 5.339, 1.231, 2.161);
  txt(s, HEADING, 0.892, 6.6, 2.908, 0.438, ['Miela Bathroom Set']);
  txt(s, PRICE, 5.941, 6.488, 1.106, 0.549, ['$2799']);
  rule(s, 4.081, 6.818, 1.628, CREAM);
  star(s, 11.193, 0.902, 1.077);
}

/* 17 - Product Details */
function slide17(s) {
  block(s, 3.309, 0.715, 0.831, 2.155, 270);
  photo(s, 0.569, 2.224, 4.233, 5.292);
  photo(s, 5.396, 3.75, 7.938, 3.75);
  block(s, 4.483, 5.339, 1.231, 2.161);
  txt(s, BODY_TEXT, 5.396, 2.041, 4.77, 1.279, [SERENITY_ALONE]);
  txt(s, HEADING, 5.396, 1.377, 2.347, 0.438, [DETAILS]);
  txt(s, HEADING, 0.959, 2.523, 1.39, 0.774, ['Modern Sink']);
  txt(s, PRICE, 3.53, 2.635, 0.898, 0.549, ['$399']);
  rule(s, 2.526, 2.943, 0.726, BROWN);
  txt(s, HEADING, 6.184, 4.378, 2.402, 0.438, ['Modern Bathub']);
  txt(s, PRICE, 11.266, 4.322, 1.106, 0.549, ['$399']);
  rule(s, 8.788, 4.614, 2.409, CREAM);
  star(s, 11.193, 1.67, 1.077);
}

/* 18 - Room Accessory */
function slide18(s) {
  photo(s, 7.421, 0, 5.913, 6.125);
  txt(s, DISPLAY, 0.952, 1.917, 5.715, 2.625, ['Room', 'Accessory']);
  block(s, 10.78, 4.188, 1.231, 3.875, 270);
  txt(s, ON_YELLOW, 9.647, 5.757, 1.668, 0.774, ['Luxurious Collection']);
  star(s, 5.59, 5.684, 1.077);
  txt(s, BODY_TEXT, 0.948, 5.896, 3.448, 0.976, [HAPPY]);
  block(s, 6.285, 0.381, 2.429, 1.667, 270);
}

/* 19 - 01. */
function slide19(s) {
  txt(s, BODY_TEXT, 0.946, 2.123, 2.434, 0.673, [HAPPY_SHORT]);
  txt(s, STEP_NUM, 0.946, 1.245, 0.803, 0.539, ['01.']);
  txt(s, STEP_NAME, 0.946, 1.68, 1.471, 0.364, ['Miela Furniture']);
  txt(s, BODY_TEXT, 0.946, 4.185, 2.434, 0.673, [HAPPY_SHORT]);
  txt(s, STEP_NUM, 0.946, 3.307, 0.803, 0.539, ['03.']);
  txt(s, STEP_NAME, 0.946, 3.742, 1.471, 0.364, ['Miela Furniture']);
  txt(s, BODY_TEXT, 3.379, 2.123, 2.434, 0.673, [HAPPY_SHORT]);
  txt(s, STEP_NUM, 3.379, 1.245, 0.803, 0.539, ['02.']);
  txt(s, STEP_NAME, 3.379, 1.68, 1.471, 0.364, ['Miela Furniture']);
  txt(s, BODY_TEXT, 3.379, 4.209, 2.434, 0.673, [HAPPY_SHORT]);
  txt(s, STEP_NUM, 3.379, 3.332, 0.803, 0.539, ['04.']);
  txt(s, STEP_NAME, 3.379, 3.766, 1.471, 0.364, ['Miela Furniture']);
  star(s, 6.299, 1.784, 0.945);
  txt(s, HEADING, 1.094, 6.074, 3.757, 0.774, ['Your Home, Your Sanctuary, Our Expertise. ']);
  photo(s, 6.546, 3.763, 2.095, 3.029);
  photo(s, 9.188, 3.763, 3.854, 3.029);
  photo(s, 10.813, 1.346, 2.229, 1.967);
  photo(s, 8.006, 1.346, 2.229, 1.967);
  block(s, 9.845, 1.016, 0.78, 0.661);
  block(s, 8.835, 6.461, 0.78, 0.661);
}

/* 20 - Product Details */
function slide20(s) {
  block(s, 9.246, 5.412, 4.087, 1.507);
  txt(s, BODY_TEXT, 1.116, 5.638, 4.663, 1.279, [SERENITY_ALONE]);
  txt(s, HEADING, 1.116, 4.975, 2.347, 0.438, [DETAILS]);
  txt(s, HEADING, 1.116, 2.138, 2.347, 0.774, ['Furniture That Tells your story']);
  photo(s, 6.667, 0, 5.417, 6.125);
  block(s, 5.464, 0.316, 2.138, 1.507, 90);
  star(s, 4.53, 3.33, 0.945);
  txt(s, HEADING, 7.167, 4.918, 1.507, 0.774, ['Modern Armchair']);
  txt(s, PRICE, 10.578, 5.03, 1.106, 0.549, ['$399']);
  rule(s, 8.728, 5.338, 1.85, BROWN);
}

/* 21 - Product Details */
function slide21(s) {
  photo(s, 0, 1.5, 6.667, 2.854);
  block(s, 7.339, 5.968, 2.161, 0.903, 90);
  txt(s, BODY_TEXT, 7.339, 2.615, 4.663, 1.279, [SERENITY_ALONE]);
  txt(s, HEADING, 7.339, 1.952, 2.347, 0.438, [DETAILS]);
  star(s, 11.336, 1.007, 0.945);
  txt(s, HEADING, 9.632, 5.709, 2.649, 1.111, ['Crafted with Passion, Inspired by You.']);
  photo(s, 0, 4.645, 8.29, 2.855);
  block(s, 0, 4, 3.129, 0.903);
  txt(s, HEADING, 0.947, 1.721, 1.751, 0.438, ['2 Seat Sofa']);
  txt(s, PRICE, 5.24, 1.603, 0.922, 0.549, ['$599']);
  rule(s, 3.044, 1.94, 1.85, CREAM);
  txt(s, HEADING, 0.947, 6.623, 1.751, 0.438, ['3 Seat Sofa']);
  txt(s, PRICE, 6.926, 6.505, 0.922, 0.549, ['$599']);
  rule(s, 3.044, 6.842, 3.623, CREAM);
}

/* 22 - Product Details */
function slide22(s) {
  block(s, 10.417, 5.747, 2.917, 1.232);
  txt(s, BODY_TEXT, 0.972, 2.615, 4.663, 1.279, [SERENITY_ALONE]);
  txt(s, HEADING, 0.972, 1.952, 2.347, 0.438, [DETAILS]);
  star(s, 1.06, 5.986, 0.945);
  txt(s, BODY_TEXT, 0.972, 4.267, 3.631, 0.976, [HAPPY]);
  txt(s, HEADING, 3.333, 6.119, 3.158, 0.774, ['Elevate Your Living, One Piece at a Time']);
  photo(s, 7.492, 0, 5.841, 6.5);
  block(s, 6.854, 1.11, 0.999, 4.28, 180);
  txt(s, HEADING, 8.121, 0.513, 1.467, 0.774, ['Simple ', 'Armchair']);
  txt(s, PRICE, 11.483, 0.598, 0.922, 0.549, ['$699']);
  rule(s, 9.925, 0.9, 1.171, CREAM);
}

/* 23 - Creating Spaces that Inspire. */
function slide23(s) {
  block(s, 8.271, 3.75, 5.062, 0.775, 180);
  txt(s, DISPLAY, 0.952, 1.288, 8.673, 2.625, ['Creating Spaces that Inspire.']);
  star(s, 10.983, 1.827, 1.143);
  photo(s, 0, 4.155, 13.333, 3.345);
  block(s, 0, 6.857, 5.812, 0.648, 180);
}

/* 24 - Product Details */
function slide24(s) {
  txt(s, BODY_TEXT, 6.097, 5.086, 6.523, 0.976, [SERENITY_ALONE]);
  txt(s, HEADING, 6.097, 4.46, 2.351, 0.438, [DETAILS]);
  txt(s, HEADING, 0.983, 1.604, 2.351, 0.438, ['Special Edition.']);
  txt(s, BODY_TEXT, 6.097, 6.292, 4.982, 0.673, [HAPPY]);
  block(s, 11.874, 3.117, 0.805, 2.109, 90);
  star(s, 6.097, 1.827, 1.143);
  photo(s, 0, 2.542, 5.083, 4.958);
  photo(s, 8.448, 0, 4.885, 3.75);
  block(s, 4.681, 2.542, 0.805, 2.109, 180);
  block(s, 8.031, -0.025, 0.805, 2.109, 180);
  txt(s, HEADING, 0.958, 6.19, 1.577, 0.774, ['Mini Table', 'Set']);
  txt(s, PRICE, 3.847, 6.292, 1.042, 0.549, ['$2699']);
  rule(s, 2.652, 6.628, 0.941, CREAM);
}

/* 25 - Product Details */
function slide25(s) {
  block(s, 5.113, 5.042, 0.998, 2.109, 90);
  star(s, 1.058, 1.403, 1.143);
  txt(s, BODY_TEXT, 1.058, 4.787, 2.704, 2.187, [SERENITY_ALONE_PLAIN]);
  txt(s, HEADING, 1.058, 4.144, 2.351, 0.438, [DETAILS]);
  photo(s, 5.151, 0, 3.914, 6.097);
  photo(s, 9.571, 1.403, 3.914, 6.097);
  block(s, 9.524, 0.184, 0.833, 2.465, 90);
}

/* 26 - Product Details */
function slide26(s) {
  txt(s, BODY_TEXT, 0.952, 5.981, 6.523, 0.976, [SERENITY_ALONE]);
  txt(s, HEADING, 0.952, 5.354, 2.351, 0.438, [DETAILS]);
  block(s, 10.523, 3.262, 0.976, 4.646, 90);
  star(s, 1.058, 1.403, 1.143);
  txt(s, HEADING, 8.687, 1.386, 3.06, 0.774, ['Where Style Meets Functionality.']);
  txt(s, HEADING, 0.952, 3.691, 1.465, 0.774, ['Special Edition.']);
  photo(s, 4.354, 2.379, 8.979, 3.038);
  block(s, 5.507, -0.02, 0.976, 4.646, 90);
}

/* 27 - Limited Time Offer. */
function slide27(s) {
  star(s, 11.137, 1.403, 1.143);
  txt(s, HEADING, 6.909, 1.975, 1.778, 0.774, ['Limited Time Offer.']);
  txt(s, BODY_TEXT, 7.326, 4.769, 5.111, 1.279, [SERENITY_ALONE]);
  txt(s, HEADING, 7.326, 4.143, 2.351, 0.438, [DETAILS]);
  block(s, 0.974, 5.245, 0.793, 2.741, 90);
  txt(s, BODY_TEXT, 7.326, 6.292, 4.982, 0.673, [HAPPY]);
  photo(s, 0, 1.379, 6.437, 5.204);
  block(s, 6.116, 3.379, 0.793, 2.741, 180);
}

/* 28 - Product Details */
function slide28(s) {
  block(s, 6.285, 3.466, 3.065, 0.813, 180);
  photo(s, 8.643, 0, 4.69, 6.286);
  star(s, 1.043, 5.714, 1.143);
  txt(s, BODY_TEXT, 4.271, 5.276, 3.854, 1.581, [SERENITY_ALONE]);
  txt(s, HEADING, 4.271, 4.65, 2.351, 0.438, [DETAILS]);
  txt(s, BODY_TEXT, 0.959, 2.214, 2.708, 2.187, [SERENITY_ALONE]);
  txt(s, HEADING, 0.959, 1.588, 2.351, 0.438, [DETAILS]);
  block(s, 10.268, 5.941, 3.065, 0.689, 180);
  photo(s, 4.271, 0, 4.118, 3.75);
  block(s, 3.482, 0.088, 1.577, 0.689, 90);
}

/* 29 - Discover the Art of Living Beautifully. */
function slide29(s) {
  txt(s, DISPLAY, 1.167, 1.66, 6.667, 3.888, ['Discover the Art of Living Beautifully.']);
  block(s, 9.972, 4.432, 2.194, 3.068, 180);
  block(s, 1.167, 6.417, 2.104, 0.5, 180);
  star(s, 11.086, 1.66, 1.143);
  txt(s, ON_YELLOW, 10.23, 4.774, 1.68, 0.774, ['Luxurious', 'Collection']);
}

/* 30 - Miela. */
function slide30(s) {
  photo(s, 7.224, 0, 6.109, 3.75);
  txt(s, DISPLAY, 0.994, 2.177, 4.661, 1.784, ['Miela.'], { size: 100 });
  txt(s, PRICE, 1.046, 3.961, 3.092, 0.438, ['Furniture Catalogue'], { align: 'left', lh: null });
  star(s, 1.046, 5.666, 1.143);
  txt(s, BODY_TEXT, 5.453, 6.609, 2.311, 0.303, ['www.mielafurniture.com'], { lh: null });
  txt(s, BODY_TEXT, 10.055, 5.645, 1.66, 0.303, ['+11 222 3333 4444'], { lh: null });
  txt(s, BODY_TEXT, 10.055, 6.051, 2.16, 0.505, ['2701 Willow Oaks Lane Lake Charles, LA'], { lh: null });
  txt(s, HEADING, 5.453, 5.664, 2.16, 0.774, ['Visit Our Shop', 'For More']);
  txt(s, BODY_TEXT, 10.055, 6.609, 2.424, 0.303, ['yourcompany@email.com'], { lh: null });
  block(s, 9.379, 3.215, 3.954, 1.151);
}

const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10, slide11,
  slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20, slide21, slide22,
  slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
];

/* ----------------------------------------------------------------- build */

const pres = new PptxGenJS();
pres.defineLayout({ name: 'MIELA', width: SLIDE_W, height: SLIDE_H });
pres.layout = 'MIELA';
pres.author = 'Miela';
pres.title = 'Miela. Furniture Catalogue';

BUILDERS.forEach(function (buildSlide, i) { buildSlide(newSlide(pres, i + 1)); });

pres.writeFile({ fileName: path.join(__dirname, 'e31ae818-5a03-42a1-9b54-42a909ca10f4_grok_final.pptx') })
  .then(function (name) { console.log('wrote ' + name); })
  .catch(function (err) { console.error(err); process.exit(1); });
