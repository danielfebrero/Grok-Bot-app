/**
 * "Healthcare" presentation template - rebuilt with pptxgenjs.
 *
 * 35 slides on a 26.667 x 15 inch (16:9) canvas.  Every position is in inches,
 * every colour is a hex literal from the palette below and every slide is built
 * by its own `slideNN()` function so the deck's design can be read off the code.
 *
 * Photographs in the original are represented by flat grey `[image]` boxes.
 *
 *   node <this-file>.js   ->   writes the .pptx next to this script
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- design tokens
const C = {
  teal:     '24A3A1',   // primary brand colour, full-bleed panels
  tealSoft: '5EBABA',   // secondary panels / cards
  ice:      'DAF3FA',   // pale blue diamonds and page numbers
  sand:     'FCE077',   // yellow accent
  rose:     'FBB7C8',   // pink accent
  white:    'FFFFFF',
  ink:      '000000',   // icon glyphs
  body:     '1A1A1A',   // default body copy
  star:     'FFC000',   // rating stars
  photo:    'D8D8D8'    // image placeholder fill
};

const F = { display: 'Mulish', body: 'Roboto' };

const SLIDE_W = 26.6667;
const SLIDE_H = 15;

// Body copy is reused verbatim all over the deck.
const LOREM =       "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book.";
const LOREM_B =     'It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged.';
const LOREM_PRINT = "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer.";
const LOREM_SCRAM = "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled.";
const LOREM_NP =    "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer";
const LOREM_UNK =   "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown.";
const LOREM_1500 =  "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s.";
const LOREM_SHORT = "Lorem Ipsum has been the industry's standard dummy text ever since.";
const LOREM_TINY =  "Lorem Ipsum has been the industry's standard dummy.";
const LOREM_LINE =  'Lorem Ipsum has been';
const STAT_BODY =   "It's your first one, and your hands are already sweating at the thought of having to speak in front of a crowd.";
const QUOTE =       '“This is a quote. Words full of wisdom that someone important said”';
const TESTIMONIAL = "“Lorem Ipsum has been the industry's standard dummy.”";

// ------------------------------------------------------------------- helpers
/** Plain rectangle. */
function band(s, x, y, w, h, color) {
  s.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: color }, line: { type: 'none' } });
}

/** Rounded rectangle card (PowerPoint's default 1/6-of-short-side radius). */
function panel(s, x, y, w, h, color) {
  s.addShape('roundRect', { x: x, y: y, w: w, h: h, fill: { color: color }, line: { type: 'none' } });
}

/** Rounded rectangle with an explicit corner radius. */
function softBar(s, x, y, w, h, color, radius) {
  s.addShape('roundRect', {
    x: x, y: y, w: w, h: h, rectRadius: radius === undefined ? 0.411 : radius,
    fill: { color: color }, line: { type: 'none' }
  });
}

/** Ellipse. */
function bulb(s, x, y, w, h, color) {
  s.addShape('ellipse', { x: x, y: y, w: w, h: h, fill: { color: color }, line: { type: 'none' } });
}

/** The pale blue diamond watermark: a rounded square turned 45 degrees. */
function iceDiamond(s, x, y, size) {
  s.addShape('roundRect', {
    x: x, y: y, w: size, h: size, rotate: 45,
    fill: { color: C.ice, transparency: 10 }, line: { type: 'none' }
  });
}

/** Rounded square turned 45 degrees, sized by the diamond's bounding box. */
function roundDiamond(s, x, y, box, color) {
  const side = box / Math.SQRT2;
  s.addShape('roundRect', {
    x: x + (box - side) / 2, y: y + (box - side) / 2, w: side, h: side, rotate: 45,
    rectRadius: 0.14, fill: { color: color }, line: { type: 'none' }
  });
}

/** Grey stand-in for a photograph. */
function photo(s, x, y, w, h) {
  s.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: C.photo }, line: { type: 'none' } });
  s.addText('[image]', {
    x: x, y: y + h / 2 - 0.35, w: w, h: 0.7,
    fontFace: F.body, fontSize: 20, color: '8C8C8C', align: 'center', valign: 'middle', margin: 0
  });
}

/**
 * Text box. `body` is a string or an array of paragraph strings
 * (an empty string yields a blank line).  `o.bullet` is a 4-digit unicode
 * code point for a bulleted list.
 */
function text(s, x, y, w, h, body, o) {
  o = o || {};
  const runs = (Array.isArray(body) ? body : [body]).map((line, i, all) => ({
    text: line,
    options: { breakLine: i < all.length - 1, bullet: o.bullet ? { characterCode: o.bullet } : false }
  }));
  s.addText(runs, {
    x: x, y: y, w: w, h: h,
    fontFace: o.font || F.body,
    fontSize: o.size,
    bold: !!o.bold,
    color: o.color || C.body,
    align: o.align || 'left',
    valign: o.valign || 'top',
    lineSpacingMultiple: o.lineSpacing,
    margin: o.margin === undefined ? [7.2, 7.2, 3.6, 3.6] : o.margin
  });
}

/** Oversized translucent page number in the corner. */
function pageNo(s, label, x, y, color, align) {
  text(s, x, y, 7.172, 3.467, label,
       { font: F.display, size: 200, bold: true, color: color, align: align || 'right' });
}

/** Pink diamond bullet + eyebrow line + big display title. */
function header(s, x, y, eyebrow, lines, color, titleW, titleH) {
  const w = titleW || 11.69;
  const h = titleH || [0, 1.447, 2.794, 4.14][lines.length];
  s.addShape('roundRect', {
    x: x, y: y, w: 0.724, h: 0.724, rotate: 45,
    fill: { color: C.rose, transparency: 10 }, line: { type: 'none' }
  });
  text(s, x + 1.098, y + 0.042, 8.716, 0.64, eyebrow, { font: F.display, size: 32, color: color });
  text(s, x - 0.15, y + 1.491, w, h, lines, { font: F.display, size: 80, color: color });
}

/** Five-star rating row. */
function stars(s, x, y) {
  for (let i = 0; i < 5; i++) {
    s.addShape('star5', {
      x: x + i * 0.534, y: y, w: 0.392, h: 0.392,
      fill: { color: C.star }, line: { type: 'none' }
    });
  }
}

/**
 * Slide 21's callout: rounded right end with a pointed tab on the left edge.
 * Path coordinates are fractions of the shape's own box.
 */
function tagBanner(s, x, y, w, h, color) {
  const P = (fx, fy) => ({ x: fx * w, y: fy * h });
  const cubic = (x1, y1, x2, y2, fx, fy) => ({
    x: fx * w, y: fy * h,
    curve: { type: 'cubic', x1: x1 * w, y1: y1 * h, x2: x2 * w, y2: y2 * h }
  });
  s.addShape('custGeom', {
    x: x, y: y, w: w, h: h, fill: { color: color }, line: { type: 'none' },
    points: [
      P(0.8603, 0), P(0.1713, 0),
      cubic(0.1517, 0, 0.1357, 0.0571, 0.1357, 0.1269),
      P(0.1357, 0.3058), P(0, 0.5), P(0.1357, 0.6942), P(0.1357, 0.8731),
      cubic(0.1357, 0.9429, 0.1517, 1, 0.1713, 1),
      P(0.86, 1),
      cubic(0.9371, 1, 1, 0.7767, 1, 0.5),
      cubic(1, 0.2234, 0.9375, 0, 0.8603, 0),
      { close: true }
    ]
  });
}

/** Slide 15's smartphone: dark shell, white screen, speaker notch. */
function phoneMockup(s, x, y, w, h) {
  s.addShape('roundRect', { x: x, y: y, w: w, h: h, rectRadius: 1.05,
    fill: { color: '111111' }, line: { type: 'none' } });
  s.addShape('roundRect', { x: x + 0.34, y: y + 0.34, w: w - 0.68, h: h - 0.68, rectRadius: 0.8,
    fill: { color: C.white }, line: { type: 'none' } });
  s.addShape('roundRect', { x: x + w / 2 - 1.35, y: y + 0.34, w: 2.7, h: 0.62, rectRadius: 0.3,
    fill: { color: '111111' }, line: { type: 'none' } });
}

// ------------------------------------------------------------------- icon set
/** First-aid case with a star-of-life cross. */
function medkitIcon(s, x, y, ink, bg) {
  const hole = bg || C.sand;
  s.addShape('roundRect', { x: x + 0.30, y: y, w: 0.33, h: 0.22, rectRadius: 0.07,
    fill: { color: ink }, line: { type: 'none' } });
  s.addShape('rect', { x: x + 0.365, y: y + 0.07, w: 0.20, h: 0.15,
    fill: { color: hole }, line: { type: 'none' } });
  s.addShape('roundRect', { x: x + 0.02, y: y + 0.17, w: 0.885, h: 0.64, rectRadius: 0.07,
    fill: { color: ink }, line: { type: 'none' } });
  [0, 60, 120].forEach(angle => {
    s.addShape('rect', { x: x + 0.405, y: y + 0.32, w: 0.115, h: 0.34, rotate: angle,
      fill: { color: hole }, line: { type: 'none' } });
  });
}

/** Telephone handset: a thick arc open towards the upper right. */
function phoneIcon(s, x, y, ink) {
  s.addShape('blockArc', {
    x: x, y: y, w: 0.8, h: 0.8, angleRange: [45, 225], arcThicknessRatio: 0.55,
    fill: { color: ink }, line: { type: 'none' } });
}

/** Three people. */
function peopleIcon(s, x, y, ink) {
  const f = { color: ink }, none = { type: 'none' };
  s.addShape('ellipse', { x: x + 0.225, y: y + 0.03, w: 0.20, h: 0.20, fill: f, line: none });
  s.addShape('roundRect', { x: x + 0.185, y: y + 0.26, w: 0.28, h: 0.34, rectRadius: 0.09, fill: f, line: none });
  s.addShape('ellipse', { x: x + 0.02, y: y + 0.14, w: 0.155, h: 0.155, fill: f, line: none });
  s.addShape('roundRect', { x: x + 0.01, y: y + 0.33, w: 0.155, h: 0.27, rectRadius: 0.06, fill: f, line: none });
  s.addShape('ellipse', { x: x + 0.455, y: y + 0.14, w: 0.155, h: 0.155, fill: f, line: none });
  s.addShape('roundRect', { x: x + 0.465, y: y + 0.33, w: 0.155, h: 0.27, rectRadius: 0.06, fill: f, line: none });
}

/** Filled disc with a rising zig-zag arrow knocked out of it. */
function trendIcon(s, x, y, ink, bg) {
  const d = 0.689;
  s.addShape('ellipse', { x: x, y: y, w: d, h: d, fill: { color: ink }, line: { type: 'none' } });
  // Arrow path, in inches relative to the shape's own top-left corner.
  s.addShape('custGeom', {
    x: x + 0.10, y: y + 0.18, w: 0.50, h: 0.34, fill: { color: bg || C.rose }, line: { type: 'none' },
    points: [
      { x: 0.00, y: 0.30 }, { x: 0.15, y: 0.15 }, { x: 0.25, y: 0.25 }, { x: 0.40, y: 0.10 },
      { x: 0.31, y: 0.10 }, { x: 0.31, y: 0.00 }, { x: 0.50, y: 0.00 }, { x: 0.50, y: 0.19 },
      { x: 0.40, y: 0.19 }, { x: 0.40, y: 0.19 }, { x: 0.25, y: 0.34 }, { x: 0.15, y: 0.24 },
      { x: 0.06, y: 0.34 },
      { close: true }
    ]
  });
}

/** Interlocking heart. */
function heartIcon(s, x, y, ink) {
  s.addShape('heart', { x: x, y: y, w: 0.824, h: 0.733, fill: { color: ink }, line: { type: 'none' } });
}

/** Globe: disc plus knocked-out meridian and equator. */
function globeIcon(s, x, y, ink, bg) {
  const d = 0.646;
  bg = bg || C.rose;
  s.addShape('ellipse', { x: x, y: y, w: d, h: d, fill: { color: ink }, line: { type: 'none' } });
  s.addShape('ellipse', { x: x + 0.21, y: y, w: 0.22, h: d,
    fill: { type: 'none' }, line: { color: bg, width: 1.5 } });
  s.addShape('rect', { x: x + 0.02, y: y + d / 2 - 0.02, w: d - 0.04, h: 0.045,
    fill: { color: bg }, line: { type: 'none' } });
}

// ---------------------------------------------------------------- slide 1
/** Cover */
function slide01(s) {
  s.background = { color: C.white };
  band(s, 0, 0, 8.538, 15, C.teal);
  iceDiamond(s, 6.308, 1.61, 7.774);
  text(s, 11.738, 7.772, 12.66, 2.457, 'Healthcare', { font: F.display, size: 140, bold: true, align: 'right' });
  text(s, 11.738, 10.693, 12.66, 0.774, 'PRESENTATION TEMPLATE', { size: 40, align: 'right' });
  pageNo(s, '01', 17.225, 2.412, C.ice);
  photo(s, 3.577, 3.308, 8.538, 8.383);
  panel(s, 11.399, 4.638, 1.664, 1.664, C.sand);
  medkitIcon(s, 11.768, 5.065, C.ink, C.sand);
  panel(s, 2.229, 10.229, 5.138, 2.515, C.tealSoft);
  text(s, 2.479, 10.568, 4.648, 0.774, 'CONTACT US', { size: 40, bold: true, color: C.white, align: 'center' });
  panel(s, 2.469, 11.426, 4.657, 1.044, C.sand);
  text(s, 2.469, 11.561, 4.657, 0.774, '+123 456 7890', { size: 40, bold: true, align: 'center' });
}

// ---------------------------------------------------------------- slide 2
/** About us - text left */
function slide02(s) {
  iceDiamond(s, 15.221, 3.992, 7.774);
  header(s, 2.379, 2.532, 'ABOUT US, Healthcare', ['Our History Will ', 'Make Changes ', 'for The Future'], C.body);
  panel(s, 2.425, 9.483, 7.89, 3.159, C.tealSoft);
  text(s, 3.099, 10.126, 6.541, 2.036, LOREM, { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  pageNo(s, '02', 18.851, 11.144, C.ice);
  photo(s, 11.549, 7.5, 7.559, 7.5);
  photo(s, 19.108, 0, 7.559, 7.5);
}

// ---------------------------------------------------------------- slide 3
/** About us - two photos */
function slide03(s) {
  band(s, 0, 8.701, 26.704, 6.299, C.teal);
  header(s, 15.571, 2.452, 'ABOUT US, Healthcare', ['Our History Will ', 'Make Changes ', 'for The Future'], C.body);
  iceDiamond(s, 5.04, 2.778, 7.774);
  pageNo(s, '03', 0.992, 0.741, C.ice, 'left');
  photo(s, 6.318, 6.013, 7.559, 6.716);
  photo(s, 15.421, 8.701, 8.976, 6.716);
  panel(s, 2.425, 9.483, 7.89, 3.159, C.tealSoft);
  text(s, 3.099, 10.126, 6.541, 2.036, LOREM, { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  panel(s, 12.213, 4.861, 1.664, 1.664, C.sand);
  medkitIcon(s, 12.582, 5.288, C.ink, C.sand);
}

// ---------------------------------------------------------------- slide 4
/** About us - teal */
function slide04(s) {
  s.background = { color: C.teal };
  iceDiamond(s, 10.928, 1.861, 6.219);
  header(s, 2.497, 6.415, 'ABOUT US, Healthcare', ['Our History Will ', 'Make Changes ', 'for The Future'], C.white);
  pageNo(s, '04', 3.149, 0.857, C.tealSoft);
  photo(s, 13.333, 3.116, 11.448, 6.778);
  panel(s, 13.335, 9.261, 5.544, 3.159, C.sand);
  text(s, 14.01, 10.098, 4.183, 1.649, LOREM_UNK, { size: 23, valign: 'middle', margin: 3.6 });
  panel(s, 19.238, 9.3, 5.544, 3.159, C.tealSoft);
  text(s, 19.913, 10.136, 4.183, 1.649, LOREM_UNK, { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
}

// ---------------------------------------------------------------- slide 5
/** About us - side panel */
function slide05(s) {
  band(s, 16.231, 0, 10.473, 15, C.teal);
  iceDiamond(s, 13.878, 7.3, 6.219);
  header(s, 2.497, 2.452, 'ABOUT US, Healthcare', ['Our History Will Make', 'Changes for The Future'], C.body, 15.117);
  text(s, 18.238, 2.654, 6.459, 3.585, [LOREM, '', LOREM_B], { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  pageNo(s, '05', 17.881, 8.552, C.tealSoft);
  photo(s, 2.41, 8.121, 7.628, 4.577);
  photo(s, 10.61, 8.121, 7.628, 4.577);
  panel(s, 1.839, 11.218, 4.053, 1.48, C.sand);
  text(s, 1.839, 11.638, 4.053, 0.64, 'LOREM IPSUM', { size: 32, bold: true, align: 'center' });
}

// ---------------------------------------------------------------- slide 6
/** Why choose us - options */
function slide06(s) {
  band(s, 17.988, 0, 8.716, 15, C.teal);
  iceDiamond(s, 13.475, 4.276, 7.774);
  header(s, 2.379, 2.532, 'WHY CHOOSE US, Healthcare', ['We Will Provide ', 'The Best Service ', 'for Your'], C.body);
  text(s, 2.229, 9.52, 10.425, 1.262, LOREM, { size: 23, valign: 'middle', margin: 3.6 });
  panel(s, 2.229, 11.481, 5.537, 1.48, C.tealSoft);
  text(s, 2.229, 11.902, 5.537, 0.64, 'OPTION', { size: 32, bold: true, color: C.white, align: 'center' });
  panel(s, 7.852, 11.481, 5.537, 1.48, C.sand);
  text(s, 7.852, 11.902, 5.537, 0.64, 'OPTION', { size: 32, bold: true, align: 'center' });
  pageNo(s, '06', 17.68, 10.477, C.tealSoft);
  photo(s, 14.545, 2.382, 9.931, 8.003);
}

// ---------------------------------------------------------------- slide 7
/** Why choose us - option cards */
function slide07(s) {
  iceDiamond(s, 1.193, 1.223, 7.422);
  header(s, 11.431, 2.504, 'WHY CHOOSE US, Healthcare', ['We Will Provide The Best Service for Your'], C.body, 13.642, 2.794);
  text(s, 17.416, 7.8, 6.613, 3.585, [LOREM, '', LOREM_B], { size: 23, valign: 'middle', margin: 3.6 });
  pageNo(s, '07', 18.812, 11.059, C.ice);
  photo(s, 2.812, 2.546, 6.886, 6.329);
  photo(s, 9.698, 7.8, 6.886, 8.353);
  panel(s, 2.812, 8.402, 12.348, 4.112, C.tealSoft);
  text(s, 9.379, 9.206, 3.405, 0.639, 'OPTION 2', { size: 32, bold: true, color: C.white, valign: 'middle', margin: 3.6 });
  text(s, 9.379, 10.066, 4.949, 1.262, LOREM_SHORT, { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  text(s, 4.037, 9.206, 3.405, 0.639, 'OPTION 1', { size: 32, bold: true, color: C.white, valign: 'middle', margin: 3.6 });
  text(s, 4.037, 10.066, 4.949, 1.262, LOREM_SHORT, { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
}

// ---------------------------------------------------------------- slide 8
/** Why choose us - teal */
function slide08(s) {
  s.background = { color: C.teal };
  iceDiamond(s, 17.449, 3.738, 7.422);
  header(s, 2.419, 8.658, 'WHY CHOOSE US, Healthcare', ['We Will Provide The Best Service for Your'], C.white, 13.642, 2.794);
  pageNo(s, '08', 18.812, 11.059, C.tealSoft);
  photo(s, 2.269, 0, 10.769, 7.5);
  photo(s, 13.706, 0, 10.769, 7.5);
  panel(s, 17.388, 7.794, 5.118, 3.936, C.white);
  text(s, 17.911, 9.372, 4.221, 1.649, LOREM_PRINT, { size: 23, valign: 'middle', margin: 3.6 });
  text(s, 17.911, 8.585, 4.221, 0.64, 'OPTION', { size: 32, bold: true });
  panel(s, 17.388, 3.694, 5.118, 3.936, C.sand);
  text(s, 17.911, 5.273, 4.221, 1.649, LOREM_PRINT, { size: 23, valign: 'middle', margin: 3.6 });
  text(s, 17.911, 4.485, 4.221, 0.64, 'OPTION', { size: 32, bold: true });
}

// ---------------------------------------------------------------- slide 9
/** Testimonials - cards */
function slide09(s) {
  band(s, 0, 9.167, 26.704, 5.833, C.teal);
  header(s, 13.562, 2.871, 'WHY CHOOSE US, Healthcare', ['What They Say ', 'About Us'], C.body, 12.104);
  iceDiamond(s, 21.805, -2.458, 7.422);
  pageNo(s, '09', 18.82, 9.943, C.tealSoft);
  text(s, 13.412, 10.852, 8.011, 1.649, LOREM, { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  photo(s, 5.606, 0, 6.577, 7.115);
  photo(s, 5.606, 7.885, 6.577, 7.115);
  panel(s, 2.347, 3.812, 5.118, 6.688, C.tealSoft);
  text(s, 2.901, 5.496, 4.557, 0.875, TESTIMONIAL, { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  stars(s, 2.972, 6.45);
  text(s, 2.972, 4.62, 4.494, 0.639, 'NAME HERE', { size: 32, bold: true, color: C.white, valign: 'middle', margin: 3.6 });
  text(s, 2.977, 8.292, 4.557, 0.875, TESTIMONIAL, { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  stars(s, 3.048, 9.245);
  text(s, 3.048, 7.416, 4.494, 0.639, 'NAME HERE', { size: 32, bold: true, color: C.white, valign: 'middle', margin: 3.6 });
}

// ---------------------------------------------------------------- slide 10
/** Testimonials - side panel */
function slide10(s) {
  iceDiamond(s, 1.698, 8.545, 7.422);
  header(s, 2.497, 2.532, 'WHY CHOOSE US, Healthcare', ['What They Say ', 'About Us'], C.body, 10.986);
  band(s, 14.356, 0, 12.348, 15, C.teal);
  text(s, 15.93, 10.612, 4.557, 0.875, TESTIMONIAL, { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  stars(s, 16.001, 11.565);
  text(s, 16.001, 9.735, 4.494, 0.639, 'NAME HERE', { size: 32, bold: true, color: C.white, valign: 'middle', margin: 3.6 });
  text(s, 16.001, 2.618, 7.458, 3.198, [LOREM, '', LOREM_B], { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  pageNo(s, '10', 17.8, 6.052, C.tealSoft);
  photo(s, 4.231, 9.115, 10.126, 5.885);
}

// ---------------------------------------------------------------- slide 11
/** Service - two cards */
function slide11(s) {
  s.background = { color: C.teal };
  iceDiamond(s, 12.675, 0.021, 7.774);
  header(s, 2.419, 6.092, 'SERVICE, Healthcare', ['The Best Service for Our Patients'], C.white, 11.69, 2.794);
  text(s, 2.269, 11.436, 9.308, 1.262, LOREM, { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  pageNo(s, '11', 3.887, 2.074, C.tealSoft);
  photo(s, 15.636, 3.269, 7.194, 11.731);
  panel(s, 19.388, 8.762, 5.118, 3.936, C.white);
  text(s, 19.911, 10.341, 4.221, 1.649, LOREM_PRINT, { size: 23, valign: 'middle', margin: 3.6 });
  text(s, 19.911, 9.553, 4.221, 0.64, 'SERVICE', { size: 32, bold: true });
  panel(s, 19.388, 4.662, 5.118, 3.936, C.sand);
  text(s, 19.911, 6.241, 4.221, 1.649, LOREM_PRINT, { size: 23, valign: 'middle', margin: 3.6 });
  text(s, 19.911, 5.453, 4.221, 0.64, 'SERVICE', { size: 32, bold: true });
}

// ---------------------------------------------------------------- slide 12
/** Service - footer band */
function slide12(s) {
  band(s, 3.846, 8.183, 22.858, 6.817, C.teal);
  iceDiamond(s, 19.717, 5.015, 7.774);
  header(s, 2.379, 2.532, 'SERVICE, Healthcare', ['The Best Service for Our Patients'], C.body, 11.69, 2.794);
  text(s, 11.046, 10.662, 5.614, 2.036, LOREM, { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  text(s, 4.979, 10.662, 5.614, 2.036, LOREM, { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  text(s, 4.979, 9.559, 5, 0.64, 'SERVICE', { size: 32, bold: true, color: C.white });
  text(s, 11.046, 9.559, 5, 0.64, 'SERVICE', { size: 32, bold: true, color: C.white });
  pageNo(s, '12', 9.749, 4.75, C.ice);
  photo(s, 17.214, 0, 7.194, 12.557);
}

// ---------------------------------------------------------------- slide 13
/** Service - four tiles */
function slide13(s) {
  iceDiamond(s, -1.135, -2.27, 7.774);
  iceDiamond(s, 21.468, 10.33, 7.774);
  panel(s, 2.229, 1.718, 5.118, 5.118, C.teal);
  header(s, 14.691, 3.292, 'SERVICE, Healthcare', ['The Best Service for Our Patients'], C.body, 11.69, 2.794);
  text(s, 2.752, 4.479, 4.221, 1.649, LOREM_PRINT, { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  text(s, 2.752, 3.691, 4.221, 0.64, 'LOREM IPSUM', { size: 32, bold: true, color: C.white });
  text(s, 2.752, 2.388, 1.691, 1.111, 'A', { size: 60, bold: true, color: C.white });
  panel(s, 2.229, 7.028, 5.118, 5.118, C.rose);
  text(s, 2.752, 9.789, 4.221, 1.649, LOREM_PRINT, { size: 23, valign: 'middle', margin: 3.6 });
  text(s, 2.752, 9.001, 4.221, 0.64, 'LOREM IPSUM', { size: 32, bold: true });
  text(s, 2.752, 7.698, 1.691, 1.111, 'A', { size: 60, bold: true });
  panel(s, 7.496, 2.997, 5.118, 5.118, C.sand);
  text(s, 8.019, 5.758, 4.221, 1.649, LOREM_PRINT, { size: 23, valign: 'middle', margin: 3.6 });
  text(s, 8.019, 4.971, 4.221, 0.64, 'LOREM IPSUM', { size: 32, bold: true });
  text(s, 8.019, 3.668, 1.691, 1.111, 'A', { size: 60, bold: true });
  panel(s, 7.496, 8.307, 5.118, 5.118, C.tealSoft);
  text(s, 8.019, 11.068, 4.221, 1.649, LOREM_PRINT, { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  text(s, 8.019, 10.281, 4.221, 0.64, 'LOREM IPSUM', { size: 32, bold: true, color: C.white });
  text(s, 8.019, 8.978, 1.691, 1.111, 'A', { size: 60, bold: true, color: C.white });
  text(s, 14.541, 8.821, 6.613, 3.585, [LOREM, '', LOREM_B], { size: 23, valign: 'middle', margin: 3.6 });
  pageNo(s, '13', 18.255, 1.273, C.ice);
}

// ---------------------------------------------------------------- slide 14
/** Service - teal + cards */
function slide14(s) {
  s.background = { color: C.teal };
  band(s, 20.692, 0, 6.615, 15, C.white);
  iceDiamond(s, 16.098, 1.61, 7.774);
  header(s, 2.419, 2.532, 'SERVICE, Healthcare', ['The Best Service ', 'for Our Patients'], C.white);
  panel(s, 2.269, 8.863, 5.118, 3.936, C.sand);
  text(s, 2.792, 10.442, 4.221, 1.649, LOREM_PRINT, { size: 23, valign: 'middle', margin: 3.6 });
  text(s, 2.792, 9.654, 4.221, 0.64, 'SERVICE', { size: 32, bold: true });
  pageNo(s, '14', 18.638, 9.873, C.ice);
  photo(s, 12.178, 3.5, 8.514, 9.198);
  panel(s, 7.762, 8.762, 5.118, 3.936, C.tealSoft);
  text(s, 8.284, 10.341, 4.221, 1.649, LOREM_PRINT, { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  text(s, 8.284, 9.553, 4.221, 0.64, 'SERVICE', { size: 32, bold: true, color: C.white });
}

// ---------------------------------------------------------------- slide 15
/** Service - phone mockup */
function slide15(s) {
  iceDiamond(s, 9.314, 10.568, 7.774);
  band(s, 18.462, 0, 8.243, 15, C.teal);
  header(s, 2.419, 3.561, 'SERVICE, Healthcare', ['You Can Check ', 'Your Health Using ', 'Your Phone'], C.body);
  text(s, 2.312, 10.314, 7.642, 1.649, LOREM, { size: 23, valign: 'middle', margin: 3.6 });
  pageNo(s, '15', 18.997, 0.776, C.tealSoft);
  photo(s, 14.938, 2.816, 6.77, 14.508);
  phoneMockup(s, 14.508, 2.387, 7.615, 15.319);
  panel(s, 21.023, 5.393, 1.664, 1.664, C.sand);
  medkitIcon(s, 21.393, 5.821, C.ink, C.sand);
  panel(s, 20.992, 7.372, 4.053, 1.48, C.tealSoft);
  text(s, 20.992, 7.792, 4.053, 0.64, 'LOREM IPSUM', { size: 32, bold: true, color: C.white, align: 'center' });
  panel(s, 20.992, 9.199, 4.053, 1.48, C.white);
  text(s, 20.992, 9.619, 4.053, 0.64, 'LOREM IPSUM', { size: 32, bold: true, align: 'center' });
  panel(s, 20.992, 11.042, 4.053, 1.48, C.rose);
  text(s, 20.992, 11.462, 4.053, 0.64, 'LOREM IPSUM', { size: 32, bold: true, align: 'center' });
}

// ---------------------------------------------------------------- slide 16
/** Team - teal */
function slide16(s) {
  s.background = { color: C.teal };
  iceDiamond(s, 11.544, 6.115, 7.682);
  header(s, 2.419, 2.532, 'TEAM, Healthcare', ['Amazing Team for', 'Your Health'], C.white);
  text(s, 2.781, 8.693, 6.613, 3.585, [LOREM, '', LOREM_B], { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  pageNo(s, '16', 19.054, 0.315, C.tealSoft);
  photo(s, 13.727, 2.382, 8.514, 10.316);
  panel(s, 20.423, 11.218, 4.053, 1.48, C.sand);
  text(s, 20.423, 11.638, 4.053, 0.64, 'NAME HERE', { size: 32, bold: true, align: 'center' });
}

// ---------------------------------------------------------------- slide 17
/** Team - two portraits */
function slide17(s) {
  header(s, 14.188, 5.778, 'TEAM, Healthcare', ['Amazing Team for', 'Your Health'], C.body);
  text(s, 14.038, 11.022, 7.642, 1.649, LOREM, { size: 23, valign: 'middle', margin: 3.6 });
  iceDiamond(s, 6.957, -0.568, 6.9);
  pageNo(s, '17', 17.552, 2.116, C.ice);
  photo(s, 0, 2.302, 5.962, 10.316);
  photo(s, 6.231, 2.302, 5.962, 10.316);
  panel(s, 0, 10.539, 5.962, 2.615, C.tealSoft);
  text(s, 0.558, 11.087, 4.942, 1.649, LOREM_NP, { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  panel(s, 6.231, 11.698, 4.053, 1.48, C.sand);
  text(s, 6.231, 12.118, 4.053, 0.64, 'NAME HERE', { size: 32, bold: true, align: 'center' });
}

// ---------------------------------------------------------------- slide 18
/** Team - photo pair */
function slide18(s) {
  iceDiamond(s, 13.275, 2.318, 7.682);
  header(s, 2.419, 2.532, 'TEAM, Healthcare', ['Amazing Team for', 'Your Health'], C.body);
  band(s, 0, 8.183, 14.538, 6.817, C.teal);
  text(s, 2.389, 9.614, 9.174, 2.423, [LOREM, '', LOREM_B], { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  pageNo(s, '18', 18.471, 0.84, C.ice);
  photo(s, 12.206, 5.852, 5.859, 5.923);
  photo(s, 18.617, 5.852, 5.859, 5.923);
  panel(s, 13.063, 10.951, 4.053, 1.48, C.sand);
  text(s, 13.063, 11.371, 4.053, 0.64, 'NAME HERE', { size: 32, bold: true, align: 'center' });
  panel(s, 19.472, 10.941, 4.053, 1.48, C.sand);
  text(s, 19.472, 11.361, 4.053, 0.64, 'NAME HERE', { size: 32, bold: true, align: 'center' });
}

// ---------------------------------------------------------------- slide 19
/** Gallery - teal */
function slide19(s) {
  s.background = { color: C.teal };
  iceDiamond(s, 12.291, 4.444, 10.715);
  header(s, 2.379, 2.676, 'GALLERY, Healthcare', ['The Best Gallery ', 'of Our Hospital'], C.white, 12.562);
  pageNo(s, '19', 0.342, 11.267, C.tealSoft, 'left');
  text(s, 2.229, 8.978, 7.529, 1.649, LOREM, { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  photo(s, 13.938, 2.944, 4.154, 4.078);
  photo(s, 12.975, 7.5, 5.294, 5.198);
  photo(s, 18.538, 1.577, 5.859, 5.923);
  photo(s, 18.538, 8.013, 7, 5.923);
  panel(s, 16.605, 7.313, 1.664, 1.664, C.white);
  medkitIcon(s, 16.974, 7.741, C.ink, C.white);
  panel(s, 18.5, 12.698, 4.053, 1.48, C.sand);
  text(s, 18.5, 13.118, 4.053, 0.64, 'LOREM IPSUM', { size: 32, bold: true, align: 'center' });
}

// ---------------------------------------------------------------- slide 20
/** Gallery - light */
function slide20(s) {
  iceDiamond(s, -0.178, 2.912, 10.715);
  header(s, 14.826, 8.034, 'GALLERY, Healthcare', ['The Best Gallery ', 'of Our Hospital'], C.body, 12.562);
  pageNo(s, '20', 17.303, 2.032, C.ice);
  photo(s, 0, 9.998, 5.028, 5.002);
  photo(s, 5.18, 2.829, 7.092, 7.17);
  panel(s, 10.956, 3.571, 7.89, 3.159, C.tealSoft);
  text(s, 11.63, 4.214, 6.541, 2.036, LOREM, { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
}

// ---------------------------------------------------------------- slide 21
/** Infographic - syringe */
function slide21(s) {
  iceDiamond(s, 17.367, 8.607, 7.682);
  // Syringe illustration, stacked bottom-up
  band(s, 13.963, 2.364, 0.311, 2.234, C.tealSoft);   // plunger rod
  band(s, 13.693, 4.229, 0.851, 0.369, C.tealSoft);   // plunger step
  band(s, 13.491, 4.473, 1.255, 0.369, C.tealSoft);   // plunger flange
  band(s, 13.276, 4.954, 1.685, 2.478, C.rose);       // barrel, upper half
  band(s, 13.276, 7.548, 1.685, 2.478, C.sand);       // barrel, lower half
  band(s, 13.012, 10.142, 2.212, 0.424, C.teal);      // barrel collar
  band(s, 13.744, 10.142, 0.751, 1.997, C.teal);      // needle
  band(s, 13.276, 12.045, 1.685, 0.571, C.teal);      // needle tip guard
  tagBanner(s, 15.166, 2.348, 9.033, 2.530, C.tealSoft);
  bulb(s, 22.04, 2.713, 1.798, 1.798, C.white);
  tagBanner(s, 15.166, 4.932, 9.033, 2.530, C.rose);
  bulb(s, 22.039, 5.295, 1.798, 1.798, C.white);
  tagBanner(s, 15.166, 7.522, 9.033, 2.530, C.sand);
  bulb(s, 22.038, 7.886, 1.798, 1.798, C.white);
  tagBanner(s, 15.166, 10.116, 9.033, 2.530, C.teal);
  bulb(s, 22.039, 10.479, 1.798, 1.798, C.white);
  text(s, 21.745, 3.019, 2.387, 1.178, '1', { size: 70, bold: true, align: 'center', margin: 0 });
  text(s, 21.744, 5.603, 2.387, 1.178, '2', { size: 70, bold: true, align: 'center', margin: 0 });
  text(s, 21.746, 8.239, 2.387, 1.178, '3', { size: 70, bold: true, align: 'center', margin: 0 });
  text(s, 21.744, 10.835, 2.387, 1.178, '4', { size: 70, bold: true, align: 'center', margin: 0 });
  text(s, 17.271, 2.621, 3.405, 0.819, 'LOREM IPSUM', { size: 32, bold: true, color: C.white, valign: 'middle', margin: 3.6, lineSpacing: 1.5 });
  text(s, 17.272, 3.397, 4.472, 1.161, LOREM_SHORT, { size: 23, color: C.white, margin: 0 });
  text(s, 17.269, 5.079, 3.405, 0.819, 'LOREM IPSUM', { size: 32, bold: true, valign: 'middle', margin: 3.6, lineSpacing: 1.5 });
  text(s, 17.271, 5.855, 4.472, 1.161, LOREM_SHORT, { size: 23, margin: 0 });
  text(s, 17.345, 7.758, 3.405, 0.819, 'LOREM IPSUM', { size: 32, bold: true, valign: 'middle', margin: 3.6, lineSpacing: 1.5 });
  text(s, 17.346, 8.534, 4.472, 1.161, LOREM_SHORT, { size: 23, margin: 0 });
  text(s, 17.343, 10.34, 3.405, 0.819, 'LOREM IPSUM', { size: 32, bold: true, color: C.white, valign: 'middle', margin: 3.6, lineSpacing: 1.5 });
  text(s, 17.345, 11.115, 4.472, 1.161, LOREM_SHORT, { size: 23, color: C.white, margin: 0 });
  header(s, 2.473, 3.3, 'INFOGRAPHIC, Healthcare', ['Our Hospital ', 'in Infographic'], C.body, 9.033);
  text(s, 4.964, 8.957, 6.613, 3.585, [LOREM, '', LOREM_B], { size: 23, valign: 'middle', margin: 3.6 });
  pageNo(s, '21', 0.866, 11.068, C.ice, 'left');
}

// ---------------------------------------------------------------- slide 22
/** Infographic - thermometers */
function slide22(s) {
  iceDiamond(s, 19.766, 8.542, 7.682);
  text(s, 1.959, 11.455, 4.561, 1.262, LOREM_1500, { size: 23, align: 'center', valign: 'middle', margin: 3.6 });
  band(s, 7.855, 3.825, 2.071, 5.571, C.ice);
  band(s, 7.855, 6.146, 2.071, 3.25, C.sand);
  bulb(s, 7.391, 8.003, 3, 2.964, C.sand);
  band(s, 12.412, 3.825, 2.071, 5.571, C.ice);
  band(s, 12.412, 5.36, 2.071, 4.036, C.tealSoft);
  bulb(s, 11.947, 8.003, 3, 2.964, C.tealSoft);
  band(s, 3.181, 3.825, 2.071, 5.571, C.ice);
  band(s, 3.181, 7.146, 2.071, 2.25, C.rose);
  bulb(s, 2.717, 8.003, 3, 2.964, C.rose);
  text(s, 6.61, 11.455, 4.561, 1.262, LOREM_1500, { size: 23, align: 'center', valign: 'middle', margin: 3.6 });
  text(s, 11.255, 11.408, 4.561, 1.262, LOREM_1500, { size: 23, align: 'center', valign: 'middle', margin: 3.6 });
  peopleIcon(s, 3.925, 9.082, C.ink);
  trendIcon(s, 8.546, 9.141, C.ink, C.sand);
  heartIcon(s, 12.923, 9.198, C.ink);
  header(s, 16.426, 2.452, 'INFOGRAPHIC, Healthcare', ['Our Hospital ', 'in Infographic'], C.body, 9.033);
  panel(s, 16.504, 8.654, 7.894, 3.998, C.tealSoft);
  text(s, 17.881, 9.711, 5.823, 2.036, LOREM, { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  pageNo(s, '22', 0.61, 0.569, C.ice, 'left');
}

// ---------------------------------------------------------------- slide 23
/** Infographic - arrow rows */
function slide23(s) {
  band(s, 0, 0, 26.667, 8.754, C.teal);
  iceDiamond(s, 18.453, 5.249, 6.526);
  roundDiamond(s, 2.347, 7.116, 2.791, C.sand);
  softBar(s, 4.597, 7.5, 7.618, 2.023, C.tealSoft);
  softBar(s, 3.613, 7.5, 2.023, 2.023, C.rose);
  peopleIcon(s, 4.396, 8.125, C.ink);
  text(s, 6.635, 8.044, 4.797, 0.875, LOREM_TINY, { font: 'Open Sans Light', size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  roundDiamond(s, 2.347, 10.291, 2.791, C.sand);
  softBar(s, 4.597, 10.675, 7.618, 2.023, C.tealSoft);
  softBar(s, 3.613, 10.675, 2.023, 2.023, C.rose);
  text(s, 6.635, 11.219, 4.797, 0.875, LOREM_TINY, { font: 'Open Sans Light', size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  globeIcon(s, 4.302, 11.336, C.ink, C.rose);
  header(s, 2.497, 2.452, 'INFOGRAPHIC, Healthcare', ['Our Hospital in Infographic'], C.white, 15.715);
  roundDiamond(s, 13.973, 7.069, 2.791, C.sand);
  softBar(s, 16.223, 7.453, 7.618, 2.023, C.tealSoft);
  softBar(s, 15.239, 7.453, 2.023, 2.023, C.rose);
  text(s, 18.261, 7.997, 4.797, 0.875, LOREM_TINY, { font: 'Open Sans Light', size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  roundDiamond(s, 13.973, 10.244, 2.791, C.sand);
  softBar(s, 16.223, 10.628, 7.618, 2.023, C.tealSoft);
  softBar(s, 15.239, 10.628, 2.023, 2.023, C.rose);
  text(s, 18.261, 11.172, 4.797, 0.875, LOREM_TINY, { font: 'Open Sans Light', size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  trendIcon(s, 15.905, 8.12, C.ink, C.rose);
  heartIcon(s, 15.837, 11.296, C.ink);
  pageNo(s, '23', 17.431, 0.848, C.tealSoft, 'left');
}

// ---------------------------------------------------------------- slide 24
/** Infographic - two stats */
function slide24(s) {
  iceDiamond(s, 12.681, 5.795, 10.715);
  panel(s, 15.422, 7.87, 5.886, 5.091, C.tealSoft);
  text(s, 16.434, 8.55, 3.42, 0.607, 'GRAPHIC DATA', { size: 32, bold: true, color: C.white, margin: 2.37, lineSpacing: 1 });
  text(s, 16.434, 9.562, 4.331, 1.455, STAT_BODY, { size: 21, color: C.white, margin: 2.37, lineSpacing: 1 });
  text(s, 16.434, 11.257, 2.384, 0.871, '3.5K', { size: 66, bold: true, color: C.white, margin: 2.37, lineSpacing: 1 });
  header(s, 2.473, 3.3, 'INFOGRAPHIC, Healthcare', ['Our Hospital ', 'in Infographic'], C.body, 9.033);
  text(s, 2.324, 8.589, 8.138, 3.198, [LOREM, '', LOREM_B], { size: 23, valign: 'middle', margin: 3.6 });
  panel(s, 15.31, 2.073, 5.886, 5.091, C.sand);
  text(s, 16.322, 2.754, 3.42, 0.607, 'GRAPHIC DATA', { size: 32, bold: true, margin: 2.37, lineSpacing: 1 });
  text(s, 16.322, 3.765, 4.331, 1.455, STAT_BODY, { size: 21, margin: 2.37, lineSpacing: 1 });
  text(s, 16.322, 5.461, 2.384, 0.871, '2.5K', { size: 66, bold: true, margin: 2.37, lineSpacing: 1 });
  pageNo(s, '24', 18.599, 1.186, C.ice);
}

// ---------------------------------------------------------------- slide 25
/** Infographic - three stats */
function slide25(s) {
  iceDiamond(s, 0.671, 4.412, 6.584);
  panel(s, 2.347, 5.318, 5.886, 5.091, C.tealSoft);
  text(s, 3.125, 7.257, 3.42, 0.607, 'GRAPHIC DATA', { size: 32, bold: true, color: C.white, margin: 2.37, lineSpacing: 1 });
  text(s, 3.125, 8.269, 4.331, 1.455, STAT_BODY, { size: 21, color: C.white, margin: 2.37, lineSpacing: 1 });
  text(s, 3.125, 6.018, 2.384, 0.871, '3.5K', { size: 66, bold: true, color: C.white, margin: 2.37, lineSpacing: 1 });
  panel(s, 8.459, 2.302, 5.886, 5.091, C.tealSoft);
  text(s, 9.236, 4.241, 3.42, 0.607, 'GRAPHIC DATA', { size: 32, bold: true, color: C.white, margin: 2.37, lineSpacing: 1 });
  text(s, 9.236, 5.252, 4.331, 1.455, STAT_BODY, { size: 21, color: C.white, margin: 2.37, lineSpacing: 1 });
  text(s, 9.236, 3.002, 2.384, 0.871, '3.5K', { size: 66, bold: true, color: C.white, margin: 2.37, lineSpacing: 1 });
  panel(s, 8.459, 7.66, 5.886, 5.091, C.tealSoft);
  text(s, 9.236, 9.599, 3.42, 0.607, 'GRAPHIC DATA', { size: 32, bold: true, color: C.white, margin: 2.37, lineSpacing: 1 });
  text(s, 9.236, 10.61, 4.331, 1.455, STAT_BODY, { size: 21, color: C.white, margin: 2.37, lineSpacing: 1 });
  text(s, 9.236, 8.359, 2.384, 0.871, '3.5K', { size: 66, bold: true, color: C.white, margin: 2.37, lineSpacing: 1 });
  header(s, 15.981, 4.571, 'INFOGRAPHIC, Healthcare', ['Our Hospital ', 'in Infographic'], C.body, 9.033);
  pageNo(s, '25', 18.484, 0.741, C.ice);
  text(s, 15.876, 9.446, 8.138, 3.198, [LOREM, '', LOREM_B], { size: 23, valign: 'middle', margin: 3.6 });
}

// ---------------------------------------------------------------- slide 26
/** Price - single card */
function slide26(s) {
  s.background = { color: C.teal };
  iceDiamond(s, 9.778, 1.002, 7.682);
  header(s, 2.497, 9.373, 'PRICE, Healthcare', ['Price and Package'], C.white, 12.562);
  pageNo(s, '26', 3.865, 3.109, C.tealSoft, 'left');
  text(s, 15.133, 9.757, 9.186, 2.423, [LOREM, '', LOREM_B], { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  photo(s, 13.335, 0, 13.332, 7.5);
  panel(s, 13.335, 5.155, 4.704, 3.732, C.sand);
  text(s, 14.279, 5.832, 3.338, 1.717, '$39', { size: 96, bold: true, valign: 'middle', margin: 3.6 });
  text(s, 14.279, 7.672, 3.338, 0.639, 'VIP', { size: 32, bold: true, valign: 'middle', margin: 3.6 });
}

// ---------------------------------------------------------------- slide 27
/** Price - three columns */
function slide27(s) {
  iceDiamond(s, 17.394, 1.579, 7.682);
  panel(s, 19.029, 4.701, 5.369, 8.194, C.teal);
  panel(s, 13.464, 5.52, 5.369, 7.465, C.rose);
  panel(s, 7.821, 6.454, 5.369, 6.426, C.sand);
  text(s, 14.43, 6.431, 3.813, 1.717, '$39', { size: 96, bold: true, valign: 'middle', margin: 3.6 });
  text(s, 8.722, 7.196, 3.028, 1.717, '$29', { size: 96, bold: true });
  text(s, 14.43, 9.194, 3.543, 2.81, LOREM_SCRAM, { size: 23, valign: 'middle', margin: 3.6 });
  text(s, 8.798, 9.969, 3.455, 2.036, LOREM_PRINT, { size: 23, valign: 'middle', margin: 3.6 });
  text(s, 8.798, 8.965, 3.028, 0.64, 'CLASS 1', { size: 32, bold: true });
  text(s, 14.43, 8.271, 3.813, 0.639, 'VIP', { size: 32, bold: true, valign: 'middle', margin: 3.6 });
  text(s, 19.991, 5.675, 3.028, 1.717, '$49', { size: 96, bold: true, color: C.white });
  text(s, 20.068, 8.807, 3.455, 3.198, LOREM, { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  text(s, 20.068, 7.444, 3.028, 0.64, 'VVIP', { size: 32, bold: true, color: C.white });
  header(s, 2.379, 2.532, 'PRICE, Healthcare', ['Price and Package'], C.body, 12.562);
  text(s, 2.37, 8.392, 3.455, 2.036, LOREM_PRINT, { size: 23, valign: 'middle', margin: 3.6 });
  pageNo(s, '27', 1.093, 10.923, C.ice, 'left');
}

// ---------------------------------------------------------------- slide 28
/** Price - two columns */
function slide28(s) {
  band(s, 0, 9.503, 15.231, 5.377, C.teal);
  iceDiamond(s, 10.568, 6.422, 5.998);
  panel(s, 2.347, 3.158, 5.369, 8.194, C.tealSoft);
  text(s, 3.095, 5.252, 4.002, 1.717, '$29', { size: 96, bold: true, color: C.white });
  text(s, 3.139, 7.192, 3.746, 3.198, [LOREM_LINE, LOREM_LINE, LOREM_LINE, LOREM_LINE, LOREM_LINE, LOREM_LINE, LOREM_LINE, ''], { size: 23, color: C.white, valign: 'middle', margin: 3.6, bullet: '27A2' });
  text(s, 3.139, 4.193, 3.446, 0.64, 'CLASS 1', { size: 32, bold: true, color: C.white });
  panel(s, 7.889, 3.095, 5.369, 8.194, C.sand);
  text(s, 8.637, 5.189, 4.002, 1.717, '$39', { size: 96, bold: true });
  text(s, 8.681, 7.128, 3.746, 3.198, [LOREM_LINE, LOREM_LINE, LOREM_LINE, LOREM_LINE, LOREM_LINE, LOREM_LINE, LOREM_LINE, ''], { size: 23, valign: 'middle', margin: 3.6, bullet: '27A2' });
  text(s, 8.681, 4.13, 3.446, 0.64, 'CLASS 2', { size: 32, bold: true });
  header(s, 14.841, 2.558, 'PRICE, Healthcare', ['Price and Package'], C.body, 12.562);
  pageNo(s, '28', 18.682, 10.964, C.ice);
  text(s, 17.809, 7.007, 6.613, 3.585, [LOREM, '', LOREM_B], { size: 23, valign: 'middle', margin: 3.6 });
}

// ---------------------------------------------------------------- slide 29
/** Contact - footer band */
function slide29(s) {
  band(s, 0, 8.423, 19.885, 6.577, C.teal);
  header(s, 2.379, 2.532, 'CONTACT, Healthcare', ['Contact Us to ', 'Check Your Health'], C.body, 12.562);
  text(s, 2.38, 11.161, 3.928, 0.774, [LOREM_LINE, "the industry's standard"], { size: 23, color: C.white, margin: 0 });
  text(s, 2.38, 10.438, 3.928, 0.539, 'PHONE', { size: 32, bold: true, color: C.white, margin: 0 });
  text(s, 11.069, 11.159, 3.928, 0.774, [LOREM_LINE, "the industry's standard"], { size: 23, color: C.white, margin: 0 });
  text(s, 11.069, 10.436, 3.928, 0.539, 'ADDRESS', { size: 32, bold: true, color: C.white, margin: 0 });
  text(s, 6.724, 11.159, 3.928, 0.774, [LOREM_LINE, "the industry's standard"], { size: 23, color: C.white, margin: 0 });
  text(s, 6.724, 10.436, 3.928, 0.539, 'EMAIL', { size: 32, bold: true, color: C.white, margin: 0 });
  iceDiamond(s, 17.436, 1.433, 7.774);
  pageNo(s, '29', 18.682, 10.964, C.ice);
  photo(s, 15.282, 2.382, 7.185, 8.852);
  panel(s, 21.635, 8.529, 1.664, 1.664, C.sand);
  phoneIcon(s, 22.067, 8.961, C.ink);
}

// ---------------------------------------------------------------- slide 30
/** Contact - side card */
function slide30(s) {
  iceDiamond(s, 1.244, 1.306, 7.774);
  iceDiamond(s, 21.459, 10.679, 7.774);
  header(s, 16.494, 2.869, 'CONTACT, Healthcare', ['Contact Us to ', 'Check Your ', 'Health'], C.body, 8.886);
  pageNo(s, '30', 16.263, 9.523, C.ice, 'left');
  panel(s, 14.052, 8.237, 1.664, 1.664, C.sand);
  phoneIcon(s, 14.485, 8.669, C.ink);
  photo(s, 6.445, 1.415, 8.514, 12.17);
  panel(s, 2.269, 2.382, 5.154, 10.316, C.teal);
  text(s, 3.188, 5.084, 3.928, 0.774, [LOREM_LINE, "the industry's standard"], { size: 23, color: C.white, margin: 0 });
  text(s, 3.188, 4.361, 3.928, 0.539, 'PHONE', { size: 32, bold: true, color: C.white, margin: 0 });
  text(s, 3.188, 9.793, 3.928, 0.774, [LOREM_LINE, "the industry's standard"], { size: 23, color: C.white, margin: 0 });
  text(s, 3.188, 9.069, 3.928, 0.539, 'ADDRESS', { size: 32, bold: true, color: C.white, margin: 0 });
  text(s, 3.188, 7.393, 3.928, 0.774, [LOREM_LINE, "the industry's standard"], { size: 23, color: C.white, margin: 0 });
  text(s, 3.188, 6.669, 3.928, 0.539, 'EMAIL', { size: 32, bold: true, color: C.white, margin: 0 });
}

// ---------------------------------------------------------------- slide 31
/** Contact - teal */
function slide31(s) {
  s.background = { color: C.teal };
  header(s, 2.379, 2.532, 'CONTACT, Healthcare', ['Contact Us to ', 'Check Your Health'], C.white, 12.562);
  iceDiamond(s, 17.974, 6.876, 7.774);
  panel(s, 15.532, 2.573, 1.664, 1.664, C.sand);
  phoneIcon(s, 15.964, 3.005, C.ink);
  panel(s, 17.489, 3.193, 4.657, 1.044, C.white);
  text(s, 17.489, 3.328, 4.657, 0.774, '+123 456 7890', { size: 40, bold: true, align: 'center' });
  pageNo(s, '31', 15.532, 4.833, C.tealSoft, 'left');
  photo(s, 0, 8.846, 6.808, 6.154);
  photo(s, 7.321, 8.846, 6.808, 6.154);
  photo(s, 14.641, 8.846, 6.808, 6.154);
}

// ---------------------------------------------------------------- slide 32
/** Break - quote right */
function slide32(s) {
  band(s, 9.701, 0, 16.987, 6.577, C.teal);
  iceDiamond(s, 13.404, 5.657, 8.089);
  header(s, 2.379, 8.051, 'BREAK, Healthcare', ['Time to Break,', 'Enjoy for A Moment'], C.body, 12.562);
  pageNo(s, '32', 2.17, 1.991, C.ice, 'left');
  photo(s, 8.51, 2.382, 5.451, 4.971);
  photo(s, 14.24, 4.091, 8.504, 7.755);
  panel(s, 18.308, 9.141, 6.09, 3.557, C.tealSoft);
  panel(s, 8.508, 1.991, 4.053, 1.48, C.sand);
  text(s, 8.508, 2.411, 4.053, 0.64, 'LOREM IPSUM', { size: 32, bold: true, align: 'center' });
  text(s, 19.017, 9.709, 5.134, 2.524, QUOTE, { size: 36, color: C.white, valign: 'middle', margin: 3.6 });
}

// ---------------------------------------------------------------- slide 33
/** Break - quote left */
function slide33(s) {
  iceDiamond(s, 11.675, 10.753, 8.089);
  header(s, 11.986, 2.832, 'BREAK, Healthcare', ['Time to Break,', 'Enjoy for A Moment'], C.body, 12.562);
  iceDiamond(s, 20.465, -4.045, 8.089);
  pageNo(s, '33', 17.225, 9.747, C.ice);
  photo(s, 0, 0, 9.769, 15);
  panel(s, 6.724, 9.18, 6.09, 3.557, C.tealSoft);
  text(s, 7.433, 9.747, 5.134, 2.524, QUOTE, { size: 36, color: C.white, valign: 'middle', margin: 3.6 });
}

// ---------------------------------------------------------------- slide 34
/** Break - teal */
function slide34(s) {
  s.background = { color: C.teal };
  iceDiamond(s, 4.505, 8.692, 8.089);
  header(s, 2.419, 2.532, 'BREAK, Healthcare', ['Time to Break,', 'Enjoy for A Moment'], C.white, 12.562);
  pageNo(s, '34', 17.647, 9.764, C.tealSoft);
  text(s, 17.078, 9.093, 4.053, 1.649, LOREM_PRINT, { size: 23, color: C.white, valign: 'middle', margin: 3.6 });
  photo(s, 8.923, 7.627, 6.808, 6.154);
  photo(s, 15.731, 1.346, 6.808, 6.154);
  panel(s, 4.724, 9.18, 6.09, 3.557, C.tealSoft);
  text(s, 5.433, 9.747, 5.134, 2.524, QUOTE, { size: 36, color: C.white, valign: 'middle', margin: 3.6 });
  panel(s, 15.731, 6.538, 4.053, 1.48, C.sand);
  text(s, 15.731, 6.958, 4.053, 0.64, 'LOREM IPSUM', { size: 32, bold: true, align: 'center' });
}

// ---------------------------------------------------------------- slide 35
/** Thank you */
function slide35(s) {
  iceDiamond(s, 9.327, 6.993, 7.774);
  text(s, 2.269, 2.502, 12.66, 2.457, 'Thank You', { font: F.display, size: 140, bold: true });
  text(s, 2.654, 4.968, 12.66, 0.774, 'TAKE CARE OF YOUR HEALTH', { size: 40 });
  pageNo(s, '35', 2.269, 9.717, C.ice, 'left');
  panel(s, 16.421, 3.621, 1.664, 1.664, C.teal);
  medkitIcon(s, 16.79, 4.048, C.white, C.teal);
  text(s, 18.711, 3.621, 4.749, 2.423, LOREM, { size: 23, valign: 'middle', margin: 3.6 });
  photo(s, 12.389, 8.846, 6.808, 6.154);
  photo(s, 19.859, 8.846, 6.808, 6.154);
  panel(s, 12.355, 7.161, 5.138, 2.515, C.tealSoft);
  text(s, 12.605, 7.5, 4.648, 0.774, 'CONTACT US', { size: 40, bold: true, color: C.white, align: 'center' });
  panel(s, 12.595, 8.358, 4.657, 1.044, C.sand);
  text(s, 12.595, 8.493, 4.657, 0.774, '+123 456 7890', { size: 40, bold: true, align: 'center' });
}

// ---------------------------------------------------------------------- build
const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07,
  slide08, slide09, slide10, slide11, slide12, slide13, slide14,
  slide15, slide16, slide17, slide18, slide19, slide20, slide21,
  slide22, slide23, slide24, slide25, slide26, slide27, slide28,
  slide29, slide30, slide31, slide32, slide33, slide34, slide35,
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'HEALTHCARE', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'HEALTHCARE';
  pptx.author = 'pptxgenjs';
  pptx.title = 'Healthcare Presentation Template';

  BUILDERS.forEach(fn => fn(pptx.addSlide()));

  const out = path.join(__dirname, '12202acc-0d13-4cf1-afb0-3d583cca057b_grok_final.pptx');
  return pptx.writeFile({ fileName: out }).then(() => console.log('wrote ' + out));
}

build().catch(err => { console.error(err); process.exit(1); });
