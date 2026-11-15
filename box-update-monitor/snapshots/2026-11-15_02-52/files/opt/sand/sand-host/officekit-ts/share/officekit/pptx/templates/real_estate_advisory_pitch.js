/**
 * Real Property — 30-slide real-estate template deck, rebuilt with pptxgenjs.
 *
 * Run:  node 156f36fb-924b-432d-bc0c-68087cc1640d_grok_final.js
 * Writes 156f36fb-924b-432d-bc0c-68087cc1640d_grok_final.pptx next to this file.
 *
 * The original deck's photos live in empty picture placeholders (no image data),
 * so they render as blank areas. The few real raster/vector graphics it does
 * contain are redrawn here with native pptxgenjs shapes.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette ---
const BG = 'F3F3F5';    // page background
const DARK = '18181A';  // headline / logo ink
const GREEN = '529933';  // brand accent
const GRAY = '808080';  // body copy
const WHITE = 'FFFFFF';
const TRACK = 'D9D9D9';  // progress-bar track

const HEAD_FONT = 'Poppins SemiBold';
const BODY_FONT = 'Open Sans';

// pptxgenjs mutates the shadow object it is given, so hand it a fresh one each time.
const cardShadow = () => ({ type: 'outer', color: '000000', opacity: 0.4, blur: 4, offset: 3, angle: 45 });

// ---------------------------------------------------------------- helpers ---
function text(s, str, x, y, w, h, o) {
  s.addText(str, Object.assign({ x, y, w, h, valign: 'top', isTextBox: true }, o));
}

/** Page furniture: stacked logo squares + "Real Property" + "Template 2024". */
function chrome(s, o) {
  o = o || {};
  if (o.logo !== false) {
    s.addShape('rect', { x: 1.021, y: 0.622, w: 0.147, h: 0.143, fill: { color: DARK } });
    s.addShape('rect', { x: 1.094, y: 0.693, w: 0.147, h: 0.143, fill: { color: GREEN } });
    text(s, [
      { text: 'Real ', options: { color: DARK } },
      { text: 'Property', options: { color: GREEN } },
    ], 1.384, 0.578, 1.733, 0.303, { fontFace: HEAD_FONT, fontSize: 12, bold: true });
  }
  if (o.tag !== false) {
    text(s, 'Template 2024', 11.019, 0.637, 1.733, 0.286,
      { fontFace: HEAD_FONT, fontSize: 11, bold: true, color: DARK });
  }
}

/** Two-tone section headline: dark first phrase, green second phrase. */
function heading(s, dark, green, x, y, w, h, o) {
  o = o || {};
  text(s, [
    { text: dark, options: { color: DARK } },
    { text: green, options: { color: GREEN } },
  ], x, y, w, h, {
    fontFace: HEAD_FONT, fontSize: o.size || 32, bold: true, align: o.align || 'left',
  });
}

/** Open Sans paragraph copy (150% leading unless `tight`). */
function body(s, str, x, y, w, h, o) {
  o = o || {};
  text(s, str, x, y, w, h, {
    fontFace: BODY_FONT, fontSize: o.size || 12, color: o.color || GRAY,
    italic: !!o.italic, align: o.align || 'left',
    lineSpacingMultiple: o.tight ? undefined : 1.5,
  });
}

/** Poppins caption / stat / button label. */
function label(s, str, x, y, w, h, o) {
  o = o || {};
  text(s, str, x, y, w, h, {
    fontFace: HEAD_FONT, fontSize: o.size || 16, color: o.color || GREEN,
    bold: !!o.bold, align: o.align || 'left',
  });
}

/** Filled panel / button / stat block. */
function card(s, x, y, w, h, color, shadow) {
  s.addShape('rect', {
    x, y, w, h, fill: { color }, shadow: shadow ? cardShadow() : undefined,
  });
}

/** 0.13" square bullet marker. */
function tick(s, x, y) {
  s.addShape('rect', { x, y, w: 0.13, h: 0.124, fill: { color: GREEN } });
}

/** Rounded progress-bar segment. */
function bar(s, x, y, w, h, color) {
  s.addShape('roundRect', { x, y, w, h, fill: { color }, rectRadius: h / 2 });
}

/** Envelope glyph (replaces the deck's small mail icon graphics).
 *  `behind` is the colour showing through the flap notch. */
function mailIcon(s, x, y, size, color, behind) {
  const top = y + size * 0.17;
  const height = size * 0.66;
  s.addShape('roundRect', { x, y: top, w: size, h: height, fill: { color }, rectRadius: 0.02 });
  s.addShape('triangle', {
    x: x + size * 0.06, y: top + height * 0.08, w: size * 0.88, h: height * 0.5,
    flipV: true, fill: { color: behind },
  });
}

/** Globe glyph: filled disc, hollow ring, meridian + equator. */
function globeIcon(s, x, y, size) {
  s.addShape('ellipse', { x, y, w: size, h: size, fill: { color: GREEN } });
  s.addShape('ellipse', {
    x: x + size * 0.13, y: y + size * 0.13, w: size * 0.74, h: size * 0.74,
    fill: { color: BG },
  });
  s.addShape('ellipse', {
    x: x + size * 0.31, y: y + size * 0.13, w: size * 0.38, h: size * 0.74,
    fill: { color: GREEN },
  });
  s.addShape('ellipse', {
    x: x + size * 0.36, y: y + size * 0.13, w: size * 0.28, h: size * 0.74,
    fill: { color: BG },
  });
  s.addShape('rect', {
    x: x + size * 0.13, y: y + size * 0.44, w: size * 0.74, h: size * 0.12,
    fill: { color: GREEN },
  });
}

/** House glyph: pitched roof over a body with a doorway. */
function homeIcon(s, x, y, w, h) {
  s.addShape('triangle', { x, y, w, h: h * 0.55, fill: { color: GREEN } });
  s.addShape('rect', {
    x: x + w * 0.19, y: y + h * 0.5, w: w * 0.62, h: h * 0.5, fill: { color: GREEN },
  });
  s.addShape('rect', {
    x: x + w * 0.41, y: y + h * 0.72, w: w * 0.18, h: h * 0.28, fill: { color: BG },
  });
}

/** Stand-in for a raster photo (the original embeds a bitmap here). */
function imagePlaceholder(s, x, y, w, h) {
  s.addShape('rect', { x, y, w, h, fill: { color: 'E4E4E6' }, line: { color: TRACK, width: 1 } });
  text(s, '[image]', x, y + h / 2 - 0.18, w, 0.36, {
    fontFace: BODY_FONT, fontSize: 12, color: GRAY, align: 'center', valign: 'middle',
  });
}

/** Big green diagonal wedge anchored to the bottom-left corner (closing slide). */
function cornerWedge(s) {
  s.addShape('custGeom', {
    x: 0, y: 2.351, w: 5.565, h: 5.149, fill: { color: GREEN },
    points: [{ x: 0, y: 0 }, { x: 5.565, y: 5.149 }, { x: 0, y: 5.149 }, { close: true }],
  });
}

function slide01(s) {
  chrome(s);
  heading(s, 'Your Trusted Real Estate ', 'Advisors', 8.8, 1.983, 3.805, 1.919, { size: 36 });
  card(s, 8.905, 5.139, 1.105, 0.378, GREEN);
  label(s, 'Start Slide', 8.822, 5.185, 1.272, 0.286, { align: 'center', size: 11, color: WHITE });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod', 8.8, 4.082, 3.805, 0.675);
  label(s, 'www.yourwebsite.com', 0.828, 6.521, 2.122, 0.286, { size: 11, color: DARK, bold: true });
}

function slide02(s) {
  chrome(s);
  card(s, 1.384, 4.557, 2.605, 1.642, GREEN);
  heading(s, 'Welcome The Page ', 'About Real Estate', 7.002, 1.709, 5.254, 1.178);
  label(s, 'Short Info About Real Estate', 7.002, 3.232, 3.332, 0.37);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, velit esse cillum ', 7.002, 3.887, 5.147, 0.978);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore ', 7.002, 5.117, 4.991, 0.675);
}

function slide03(s) {
  chrome(s);
  heading(s, 'We Provide To Best ', 'The Real Estate', 1.239, 1.597, 4.735, 1.178);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore', 1.239, 3.041, 4.94, 0.675);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, nostrud exercitation', 1.239, 4.003, 5.147, 0.978);
  card(s, 1.337, 5.405, 1.523, 0.497, GREEN);
  label(s, 'Read More', 1.361, 5.469, 1.475, 0.37, { align: 'center', color: WHITE });
}

function slide04(s) {
  chrome(s);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna ', 6.667, 3.134, 5.147, 0.675);
  tick(s, 6.785, 4.406);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore', 7.194, 4.13, 4.315, 0.675);
  heading(s, 'Real Estate Investment ', 'Understanding Here', 6.667, 1.665, 5.497, 1.178);
  tick(s, 6.792, 5.436);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore', 7.201, 5.16, 4.315, 0.675);
}

function slide05(s) {
  chrome(s);
  card(s, 8.021, 4.518, 2.605, 1.642, GREEN);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore ', 1.257, 3.322, 4.913, 0.675);
  heading(s, 'About The We Have ', 'Development Design', 1.258, 1.896, 5.11, 1.178);
  label(s, 'Description 1', 1.257, 4.356, 2.254, 0.37);
  body(s, 'Lorem ipsum dolor sit amet, consectetur', 1.273, 4.929, 2.254, 0.675);
  label(s, 'Description 2', 3.739, 4.356, 2.254, 0.37);
  body(s, 'Lorem ipsum dolor sit amet, consectetur', 3.755, 4.929, 2.254, 0.675);
  label(s, '1500 +', 8.569, 4.737, 1.509, 0.505, { align: 'center', size: 24, color: WHITE, bold: true });
  body(s, 'Lorem ipsum dolor sit amet, consectetur', 8.236, 5.243, 2.175, 0.675, { align: 'center', color: WHITE });
}

function slide06(s) {
  chrome(s, { tag: false });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, nostrud', 1.334, 4.446, 4.96, 0.978);
  heading(s, 'The Vision That We ', 'Run And Implement', 1.358, 2.077, 4.912, 1.178);
  tick(s, 1.426, 3.834);
  label(s, 'Description 1', 1.75, 3.71, 1.798, 0.37, { color: DARK });
  tick(s, 3.889, 3.838);
  label(s, 'Description 2', 4.213, 3.715, 2.254, 0.37, { color: DARK });
  card(s, 10.073, 2.126, 2.229, 1.353, GREEN);
  label(s, '2500 +', 10.467, 2.223, 1.509, 0.505, { align: 'center', size: 24, color: WHITE, bold: true });
  body(s, 'Lorem ipsum dolor sit amet, consectetur ', 10.195, 2.687, 1.985, 0.675, { align: 'center', color: WHITE });
  card(s, 10.086, 3.939, 2.229, 1.353, GREEN);
  label(s, '3500 +', 10.48, 4.037, 1.509, 0.505, { align: 'center', size: 24, color: WHITE, bold: true });
  body(s, 'Lorem ipsum dolor sit amet, consectetur ', 10.208, 4.5, 1.985, 0.675, { align: 'center', color: WHITE });
}

function slide07(s) {
  chrome(s);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et ', 7.24, 3.652, 4.776, 0.675);
  heading(s, 'The Mission We Set In ', 'Development Here', 6.785, 1.525, 5.231, 1.178);
  tick(s, 6.917, 3.205);
  label(s, 'Description One', 7.24, 3.082, 2.116, 0.37);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et ', 7.24, 5.301, 4.776, 0.675);
  tick(s, 6.917, 4.854);
  label(s, 'Description Two', 7.24, 4.731, 2.116, 0.37);
}

function slide08(s) {
  chrome(s);
  card(s, 8.355, 5.499, 2.636, 0.769, GREEN);
  heading(s, 'Our Real Estate ', 'Manager Expert', 1.45, 1.627, 4.387, 1.178);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore dolore magna aliqua. Ut enim ad minim ', 1.482, 3.042, 4.387, 0.978);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore', 1.482, 4.257, 4.387, 0.675);
  card(s, 1.589, 5.375, 1.523, 0.497, GREEN);
  label(s, 'See More', 1.613, 5.439, 1.475, 0.37, { align: 'center', color: WHITE });
  label(s, 'Daniel Gallego', 8.614, 5.545, 2.116, 0.404, { align: 'center', size: 18, color: WHITE });
  body(s, 'Manager Expert', 8.717, 5.836, 1.911, 0.372, { align: 'center', color: WHITE, italic: true });
}

function slide09(s) {
  chrome(s);
  card(s, 0, 5.5, 13.333, 2, GREEN);
  heading(s, 'The Luxury Of Our ', 'Real Estate', 2.974, 1.297, 7.385, 0.639, { align: 'center' });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo ', 1.932, 2.218, 9.469, 0.675, { align: 'center' });
  card(s, 11.238, 3.439, 1.296, 0.773, GREEN);
  card(s, 11.238, 4.367, 1.296, 0.773, GREEN);
}

function slide10(s) {
  chrome(s);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna ', 1.308, 3.331, 5.146, 0.675);
  heading(s, 'We Are The Best ', 'Property Real Estate', 1.308, 1.806, 5.216, 1.178);
  tick(s, 1.393, 4.501);
  label(s, 'Description One', 1.796, 4.409, 2.116, 0.37);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore dolore', 1.796, 5.02, 4.657, 0.675);
}

function slide11(s) {
  chrome(s);
  heading(s, 'Meet Our Real ', 'Estate Team', 2.974, 1.297, 7.385, 0.639, { align: 'center' });
  card(s, 1.81, 2.454, 2.615, 3.612, WHITE, true);
  card(s, 5.359, 2.454, 2.615, 3.612, GREEN);
  card(s, 8.909, 2.454, 2.615, 3.612, WHITE, true);
  label(s, 'Avery Davis', 1.936, 5.255, 1.74, 0.37);
  body(s, 'Assistant CEO', 1.936, 5.58, 1.911, 0.372, { italic: true });
  mailIcon(s, 3.899, 5.462, 0.291, 'CFAF96', WHITE);
  label(s, 'Chiaki Sato', 9.034, 5.255, 1.911, 0.37);
  body(s, 'Assistant CEO', 9.034, 5.58, 1.911, 0.372, { italic: true });
  mailIcon(s, 10.997, 5.462, 0.291, 'CFAF96', WHITE);
  label(s, 'Claudia Alves', 5.487, 5.25, 1.911, 0.37, { color: WHITE });
  body(s, 'Assistant CEO', 5.487, 5.575, 1.911, 0.372, { color: WHITE, italic: true });
  mailIcon(s, 7.45, 5.457, 0.291, WHITE, GREEN);
}

function slide12(s) {
  chrome(s);
  card(s, 1.329, 3.785, 2.392, 2.126, GREEN);
  card(s, 3.802, 1.589, 2.392, 2.126, GREEN);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna ', 6.95, 3.301, 5.216, 0.675);
  heading(s, 'The Finest Residential ', 'Properties  To Market', 6.918, 1.883, 5.216, 1.178);
  tick(s, 7.058, 4.549);
  label(s, 'Description 1', 7.381, 4.426, 1.798, 0.37);
  tick(s, 7.07, 5.37);
  label(s, 'Description 3', 7.393, 5.247, 1.798, 0.37);
  tick(s, 9.754, 4.549);
  label(s, 'Description 2', 10.077, 4.426, 1.798, 0.37);
  tick(s, 9.766, 5.37);
  label(s, 'Description 4', 10.09, 5.247, 1.798, 0.37);
  label(s, '45%', 4.25, 2.026, 1.509, 0.505, { align: 'center', size: 24, color: WHITE, bold: true });
  body(s, 'Lorem ipsum dolor sit amet, consectetur, ', 3.802, 2.671, 2.392, 0.675, { align: 'center', color: WHITE });
  label(s, '85%', 1.78, 4.209, 1.509, 0.505, { align: 'center', size: 24, color: WHITE, bold: true });
  body(s, 'Lorem ipsum dolor sit amet, consectetur, ', 1.329, 4.854, 2.392, 0.675, { align: 'center', color: WHITE });
}

function slide13(s) {
  chrome(s);
  card(s, 8.853, 2.931, 2.836, 3.382, WHITE, true);
  card(s, 5.252, 2.931, 2.836, 3.382, WHITE, true);
  card(s, 1.62, 2.931, 2.836, 3.382, WHITE, true);
  heading(s, 'Our Real Estate ', 'Pricing Table', 2.331, 1.278, 8.588, 0.639, { align: 'center' });
  card(s, 2.2, 2.488, 1.675, 0.886, GREEN);
  card(s, 5.829, 2.488, 1.675, 0.886, GREEN);
  card(s, 9.458, 2.488, 1.675, 0.886, GREEN);
  label(s, '$150', 2.486, 2.645, 1.105, 0.572, { align: 'center', size: 28, color: WHITE, bold: true });
  label(s, '$250', 6.074, 2.649, 1.186, 0.572, { align: 'center', size: 28, color: WHITE, bold: true });
  label(s, '$450', 9.655, 2.647, 1.282, 0.572, { align: 'center', size: 28, color: WHITE, bold: true });
  label(s, 'Basic Package', 2.062, 3.645, 1.976, 0.37, { align: 'center', bold: true });
  label(s, 'Premium Package', 5.479, 3.645, 2.375, 0.37, { align: 'center', bold: true });
  label(s, 'Platinum Package', 8.816, 3.651, 2.959, 0.37, { align: 'center', bold: true });
  tick(s, 2.184, 4.412);
  label(s, 'Strategic', 2.508, 4.289, 1.675, 0.37, { color: DARK });
  tick(s, 2.184, 5.001);
  label(s, 'Affordable ', 2.508, 4.877, 1.675, 0.37, { color: DARK });
  card(s, 2.393, 5.615, 1.314, 0.387, GREEN);
  label(s, 'Book Now', 2.393, 5.64, 1.314, 0.337, { align: 'center', size: 14, color: WHITE });
  tick(s, 5.808, 4.412);
  label(s, 'Minimalist', 6.131, 4.289, 1.675, 0.37, { color: DARK });
  tick(s, 5.808, 5.001);
  label(s, 'Best Quality', 6.131, 4.877, 1.675, 0.37, { color: DARK });
  card(s, 6.016, 5.615, 1.314, 0.387, GREEN);
  label(s, 'Book Now', 6.016, 5.64, 1.314, 0.337, { align: 'center', size: 14, color: WHITE });
  tick(s, 9.414, 4.412);
  label(s, 'Negotiable', 9.738, 4.289, 1.675, 0.37, { color: DARK });
  tick(s, 9.414, 5.001);
  label(s, 'The Best', 9.738, 4.877, 1.675, 0.37, { color: DARK });
  card(s, 9.623, 5.615, 1.314, 0.387, GREEN);
  label(s, 'Book Now', 9.623, 5.64, 1.314, 0.337, { align: 'center', size: 14, color: WHITE });
}

function slide14(s) {
  chrome(s);
  card(s, 3.727, 1.857, 2.199, 1.531, GREEN);
  label(s, 'The Ability', 3.898, 2.075, 1.881, 0.37, { align: 'center', color: WHITE, bold: true });
  body(s, 'Lorem ipsum dolor sit amet, consectetur', 3.727, 2.545, 2.199, 0.675, { align: 'center', color: WHITE });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod ', 8.454, 4.755, 3.963, 0.675);
  heading(s, 'Real Estate Market ', 'Cycles A Deep Dive', 6.945, 2.07, 4.883, 1.178);
  label(s, '1255 +', 6.945, 4.84, 1.509, 0.505, { size: 24 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore dolore magna aliqua. Ut enim ad minim ', 6.945, 3.571, 4.883, 0.978);
}

function slide15(s) {
  chrome(s);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore ', 1.316, 3.228, 5.047, 0.675);
  heading(s, 'Strategy Made By ', 'Our Real Estate', 1.316, 1.834, 4.883, 1.178);
  tick(s, 1.417, 4.47);
  label(s, '1400 +', 1.74, 4.305, 1.487, 0.505, { size: 24 });
  tick(s, 3.76, 4.47);
  label(s, '2500 +', 4.084, 4.305, 1.487, 0.505, { size: 24 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore ', 1.302, 5.055, 4.74, 0.675);
}

function slide16(s) {
  chrome(s);
  heading(s, 'We Provide Many ', 'Excellent References', 7.311, 2.109, 5.147, 1.178);
  label(s, '75%', 7.272, 3.765, 1.043, 0.505, { size: 24 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor', 8.315, 3.642, 3.803, 0.675);
  label(s, '95%', 7.272, 4.839, 1.043, 0.505, { size: 24 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor', 8.315, 4.717, 3.803, 0.675);
  card(s, 4.271, 4.175, 2.237, 1.4, GREEN);
  label(s, '1700 +', 4.646, 4.311, 1.487, 0.505, { align: 'center', size: 24, color: WHITE, bold: true });
  body(s, 'Lorem ipsum dolor sit amet, consectetur', 4.378, 4.8, 2.023, 0.675, { align: 'center', color: WHITE });
}

function slide17(s) {
  chrome(s, { tag: false });
  card(s, 7.076, 1.542, 3.062, 2.104, GREEN);
  label(s, '95%', 7.864, 1.775, 1.487, 0.572, { align: 'center', size: 28, color: WHITE });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor', 7.311, 2.474, 2.594, 0.978, { align: 'center', color: WHITE });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna', 1.119, 3.403, 5.075, 0.675);
  heading(s, 'The Dream House Of ', 'A Harmonious Family', 1.119, 1.979, 5.147, 1.178);
  tick(s, 1.229, 4.598);
  label(s, 'Description 1', 1.553, 4.475, 1.798, 0.37);
  tick(s, 3.652, 4.598);
  label(s, 'Description 2', 3.975, 4.475, 1.798, 0.37);
  tick(s, 1.229, 5.365);
  label(s, 'Description 3', 1.553, 5.242, 1.798, 0.37);
  tick(s, 3.652, 5.365);
  label(s, 'Description 4', 3.975, 5.242, 1.798, 0.37);
}

function slide18(s) {
  chrome(s);
  heading(s, 'Many Attractive And ', 'Dream Homes', 2.331, 1.444, 8.588, 0.639, { align: 'center' });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute ', 1.551, 2.3, 10.149, 0.675, { align: 'center' });
}

function slide19(s) {
  chrome(s);
  heading(s, 'The Dream Building ', 'Must Be Obtained', 1.222, 1.644, 5.147, 1.178);
  label(s, '01', 1.222, 3.236, 0.728, 0.505, { size: 24, color: DARK });
  body(s, 'Lorem ipsum dolor sit amet, consectetur elit, sed do eiusmod tempor incididunt ut labore et', 1.95, 3.151, 4.05, 0.675);
  label(s, '02', 1.222, 4.251, 0.728, 0.505, { size: 24 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur elit, sed do eiusmod tempor incididunt ut labore et', 1.95, 4.166, 4.05, 0.675);
  label(s, '03', 1.222, 5.266, 0.728, 0.505, { size: 24, color: DARK });
  body(s, 'Lorem ipsum dolor sit amet, consectetur elit, sed do eiusmod tempor incididunt ut labore et', 1.95, 5.181, 4.05, 0.675);
}

function slide20(s) {
  chrome(s, { logo: false });
  heading(s, 'Real Estate Explained From Property To Have ', 'Valuation Transactions', 6.96, 1.605, 5.75, 1.717);
  label(s, 'Explanation Of Home Financing', 1.178, 4.718, 4.211, 0.37);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore', 1.178, 5.309, 4.934, 0.675);
}

function slide21(s) {
  chrome(s);
  heading(s, 'A Beautiful Home ', 'For A Bright Future', 1.292, 1.852, 4.861, 1.178);
  label(s, 'Is Clearly The Best', 1.292, 3.346, 2.481, 0.37);
  tick(s, 1.384, 4.292);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore', 1.724, 4.017, 4.429, 0.675);
  tick(s, 1.384, 5.249);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore', 1.724, 4.974, 4.429, 0.675);
  card(s, 9.805, 4.892, 2.237, 1.4, GREEN);
  label(s, '1700 +', 10.18, 5.027, 1.487, 0.505, { align: 'center', size: 24, color: WHITE, bold: true });
  body(s, 'Lorem ipsum dolor sit amet, consectetur', 9.912, 5.517, 2.023, 0.675, { align: 'center', color: WHITE });
}

function slide22(s) {
  chrome(s, { logo: false });
  card(s, 1.321, 5.015, 4.165, 2.485, GREEN);
  heading(s, 'Neat And Attractive ', 'Building Estate', 7.276, 1.6, 4.861, 1.178);
  label(s, '01', 7.254, 3.125, 0.728, 0.505, { size: 24 });
  label(s, 'Description One', 7.981, 3.192, 2.481, 0.37, { color: DARK });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore ', 7.269, 3.717, 5.011, 0.675);
  label(s, '02', 7.254, 4.805, 0.728, 0.505, { size: 24 });
  label(s, 'Description Two', 7.981, 4.873, 2.481, 0.37, { color: DARK });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore ', 7.269, 5.413, 5.011, 0.675);
  card(s, 4.595, 2.91, 1.783, 0.719, GREEN);
  label(s, '1500 +', 4.799, 3.018, 1.374, 0.505, { align: 'center', size: 24, color: WHITE, bold: true });
  label(s, 'About Explanation', 2.163, 5.468, 2.481, 0.37, { align: 'center', color: WHITE });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et', 1.546, 6.104, 3.715, 0.978, { align: 'center', color: WHITE });
}

function slide23(s) {
  chrome(s);
  heading(s, 'The Role Of Real Estate ', 'Agents What To Expert', 1.417, 1.599, 5.787, 1.178);
  card(s, 4.837, 4.139, 3.224, 2.153, GREEN);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ', 1.417, 2.952, 6.266, 0.675);
  label(s, 'The Best Real Estate', 5.208, 4.572, 2.481, 0.37, { align: 'center', color: WHITE });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit', 5.208, 5.183, 2.481, 0.675, { align: 'center', color: WHITE });
}

function slide24(s) {
  chrome(s);
  card(s, 2.519, 4.799, 2.189, 1.701, GREEN);
  label(s, 'S', 3.209, 4.791, 0.81, 1.717, { align: 'center', size: 96, color: WHITE, bold: true });
  heading(s, 'Strength Of Building ', 'Materials In The Estate', 6.917, 2.064, 5.604, 1.178);
  tick(s, 7.05, 3.747);
  label(s, 'We Use The Best Ingredients', 7.403, 3.633, 3.721, 0.37);
  tick(s, 7.05, 4.468);
  label(s, 'We Always Do Our Best', 7.403, 5.066, 3.721, 0.37);
  tick(s, 7.05, 5.189);
  label(s, 'Beautiful Dream House', 7.403, 4.344, 3.721, 0.37);
}

function slide25(s) {
  chrome(s);
  card(s, 8.625, 4.799, 2.189, 1.701, GREEN);
  label(s, 'W', 9.315, 4.791, 0.81, 1.717, { align: 'center', size: 96, color: WHITE, bold: true });
  heading(s, 'Weakness In Marketing ', 'Real Estate Marketing', 1.075, 2.228, 5.787, 1.178);
  label(s, '95%', 1.075, 3.872, 1.07, 0.505, { size: 24 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.', 1.075, 4.598, 5.256, 0.675);
  label(s, 'Accelerated Development', 2.234, 3.936, 3.721, 0.37);
}

function slide26(s) {
  chrome(s);
  card(s, 2.519, 4.799, 2.189, 1.701, GREEN);
  label(s, 'O', 3.209, 4.791, 0.81, 1.717, { align: 'center', size: 96, color: WHITE, bold: true });
  heading(s, 'Opportunities In The ', 'Field Of Increased', 6.999, 2.049, 5.155, 1.178);
  bar(s, 7.095, 4.226, 4.332, 0.142, TRACK);
  bar(s, 7.095, 4.226, 3.554, 0.142, GREEN);
  label(s, 'Progress ', 6.986, 3.626, 2.23, 0.37, { bold: true });
  label(s, '95%', 10.81, 3.626, 0.897, 0.37, { bold: true });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore ', 7.007, 4.859, 4.965, 0.675);
}

function slide27(s) {
  chrome(s);
  card(s, 8.625, 4.799, 2.189, 1.701, GREEN);
  label(s, 'T', 9.315, 4.791, 0.81, 1.717, { align: 'center', size: 96, color: WHITE, bold: true });
  heading(s, 'There Many Threats In ', 'The Real Estate Area', 1.108, 2.009, 5.641, 1.178);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ', 1.136, 3.452, 5.147, 0.675);
  label(s, '125 +', 1.57, 4.568, 1.043, 0.505, { size: 24 });
  tick(s, 1.218, 4.759);
  label(s, 'The Ability', 1.57, 5.12, 1.607, 0.37, { color: DARK });
  label(s, '225 +', 3.866, 4.568, 1.278, 0.505, { size: 24 });
  tick(s, 3.515, 4.759);
  label(s, 'The Dexterity', 3.866, 5.12, 1.831, 0.37, { color: DARK });
}

function slide28(s) {
  chrome(s);
  imagePlaceholder(s, 1.255, 2.293, 4.931, 3.125);
  heading(s, 'Proper And Timely ', 'Development Here', 7.226, 2.165, 4.931, 1.178);
  label(s, 'Learn More …', 7.226, 4.965, 2.044, 0.37);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, velit esse cillum ', 7.226, 3.665, 4.931, 0.978);
}

function slide29(s) {
  chrome(s);
  heading(s, 'More Contact ', 'Information Here', 2.811, 1.243, 7.711, 0.639, { align: 'center' });
  label(s, 'Website', 2.785, 5.341, 1.265, 0.37);
  globeIcon(s, 2.327, 5.358, 0.281);
  body(s, 'www.yourwebsite.com', 2.785, 5.704, 2.178, 0.372);
  label(s, 'Email', 5.801, 5.341, 1.265, 0.37);
  body(s, 'youremail@gmail.com', 5.801, 5.704, 2.178, 0.372);
  mailIcon(s, 5.389, 5.358, 0.283, GREEN, BG);
  label(s, 'Address', 8.932, 5.358, 1.265, 0.37);
  body(s, '0123 Street, Country 0123', 8.932, 5.773, 2.551, 0.303, { tight: true });
  homeIcon(s, 8.388, 5.393, 0.331, 0.267);
}

function slide30(s) {
  chrome(s);
  cornerWedge(s);
  heading(s, 'Thank ', 'You !', 5.477, 2.146, 7.002, 1.447, { size: 80 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, velit ', 5.626, 3.764, 6.679, 0.675);
  card(s, 5.695, 4.977, 1.732, 0.378, GREEN);
  label(s, 'See You Next Time', 5.695, 5.022, 1.732, 0.286, { align: 'center', size: 11, color: WHITE });
  label(s, 'www.yourwebsite.com', 10.586, 6.214, 2.122, 0.286, { size: 11, color: DARK, bold: true });
}

// ------------------------------------------------------------------ build ---
const SLIDES = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
];

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE'; // 13.333in x 7.5in

SLIDES.forEach(build => {
  const slide = pptx.addSlide();
  slide.background = { color: BG };
  build(slide);
});

pptx.writeFile({ fileName: path.join(__dirname, '156f36fb-924b-432d-bc0c-68087cc1640d_grok_final.pptx') })
  .then(f => console.log('wrote', f));
