/*
 * "Chielo - Creative presentation template" - 30 widescreen slides.
 * Rebuilt from scratch with pptxgenjs.
 *
 * The original deck uses photo placeholders and three device mock-up images.
 * Raster content is not embedded here: every photo becomes a flat gray
 * "[image]" rectangle at the same position and size, and the phone / tablet /
 * laptop mock-ups are redrawn with native rounded rectangles.
 *
 *   node 146b6b9b-dd62-47b5-9006-bdf7fc4a1276_grok_final.js
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Design tokens (theme "Red Dark" + custom font scheme of the source) *
 * ------------------------------------------------------------------ */

const SLIDE_W = 13.33507; // 12193588 EMU, the exact stage width of the source
const SLIDE_H = 7.5;

const HEAD_FONT = 'Petit formal script'; // theme major latin
const BODY_FONT = 'quicksand';           // theme minor latin

const NAVY = '44546A';       // tx2 - headings
const NAVY_SOFT = '8497B0';  // tx2 lum 60/40 - quote attribution
const RED = '950E0E';        // accent1 - body copy
const RED_SOFT = 'ED4141';   // accent1 lum 60/40
const WHITE = 'FFFFFF';
const GRAY_25 = '404040';    // bg1 lum 25%
const GRAY_50 = '808080';    // bg1 lum 50%
const PHOTO_GRAY = 'A7A7A7'; // stand-in for the "replace your image" photos
const DEVICE_GRAY = '7F7F7F';

// The two decorative squares, both painted at 80% opacity in the original.
const PINK_PALE = { color: 'F9C0C0', transparency: 20 };
const PINK_MID = { color: 'F38080', transparency: 20 };

// Palette strip on the "Colours palette" slide.
const SWATCHES = [
  { fill: RED, name: 'Ice blue accent 1', hex: 'ffefe0', rgb: 'rgb(255,239,224)' },
  { fill: 'C71313', name: 'Ice blue accent 2', hex: 'fed9ca', rgb: 'rgb(254,217,202)' },
  { fill: '4C6574', name: 'Ice blue accent 1', hex: 'c5c5c5', rgb: 'rgb(197,197,197)' },
  { fill: 'FFC000', name: 'Ice blue accent 2', hex: '7d7d7d', rgb: 'rgb(125,125,125)' },
];

/* ------------------------------------------------------------------ *
 * Body copy re-used across the deck                                   *
 * ------------------------------------------------------------------ */

const LOREM_LONG =
  'PLACEHOLDER' +
  'PLACEHOLDER' +
  'PLACEHOLDER' +
  'PLACEHOLDER';
const LOREM_MED =
  'PLACEHOLDER' +
  'PLACEHOLDER';
const LOREM_SHORT =
  'PLACEHOLDER';
const LOREM_AND =
  'PLACEHOLDER';
const LOREM_MOCKUP =
  'PLACEHOLDER' +
  'adipiscing ipsum dolor sit amet consectetuer ';
const FUSCE =
  'PLACEHOLDER' +
  'malesuada fames ';
const SUITABLE = 'suitable as for all category, it lorem ipsum is not simply random text. ';
const FOOTER_TEXT = '\u00A9 2019\u00A0chielo \u2013 creative template';

/* ------------------------------------------------------------------ *
 * Element helpers - every slide is built from these                   *
 * ------------------------------------------------------------------ */

// Flat decorative square (pale pink or salmon).
function square(slide, fill, x, y, w, h) {
  slide.addShape('rect', { x, y, w, h, fill, line: { type: 'none' } });
}
const pale = (s, x, y, w, h) => square(s, PINK_PALE, x, y, w, h);
const mid = (s, x, y, w, h) => square(s, PINK_MID, x, y, w, h);

// Stand-in for a photo placeholder of the original deck.
function photo(slide, x, y, w, h) {
  slide.addText('[image]', {
    x, y, w, h,
    fill: { color: PHOTO_GRAY },
    color: WHITE, fontFace: BODY_FONT, fontSize: 14,
    align: 'center', valign: 'middle',
  });
}

// Script-face section title (40 pt unless overridden).
function heading(slide, text, x, y, w, h, fontSize) {
  slide.addText(text, {
    x, y, w, h,
    fontFace: HEAD_FONT, fontSize: fontSize || 40, color: NAVY,
    align: 'center', valign: 'top', wrap: false, autoFit: true,
  });
}

// Justified 14 pt red body copy with 1.5 line spacing.
function body(slide, text, x, y, w, h, color) {
  slide.addText(text, {
    x, y, w, h,
    fontFace: BODY_FONT, fontSize: 14, color: color || RED,
    align: 'justify', lineSpacingMultiple: 1.5, valign: 'top',
  });
}

// Rotated copyright strip running up the left edge of every slide.
function footer(slide) {
  slide.addText(FOOTER_TEXT, {
    x: -1.128, y: 3.624, w: 3.477, h: 0.252, rotate: 270,
    fontFace: BODY_FONT, fontSize: 9, charSpacing: 3, color: RED,
    align: 'right', valign: 'top', wrap: false, autoFit: true,
  });
}

// Free-form shape. `vb` is the source view box; `d` holds either a straight
// point [x, y] or a cubic segment [x1, y1, x2, y2, x, y], both in view-box
// units, which are scaled into inches relative to the shape origin.
function freeform(slide, spec) {
  const sx = spec.w / spec.vb[0];
  const sy = spec.h / spec.vb[1];
  const points = spec.d.map(seg => (seg.length === 2
    ? { x: seg[0] * sx, y: seg[1] * sy }
    : {
      x: seg[4] * sx, y: seg[5] * sy,
      curve: { type: 'cubic', x1: seg[0] * sx, y1: seg[1] * sy, x2: seg[2] * sx, y2: seg[3] * sy },
    }));
  points.push({ close: true });
  slide.addShape('custGeom', {
    x: spec.x, y: spec.y, w: spec.w, h: spec.h,
    fill: { color: spec.fill }, line: { type: 'none' }, points,
  });
}

/* ------------------------------------------------------------------ *
 * Title / closing slides                                              *
 * ------------------------------------------------------------------ */

function coverSlide(slide, title, box, subtitleY) {
  mid(slide, 0, 4.84, 2.917, 2.66);
  pale(slide, 8.164, 0, 5.171, 4.715);
  heading(slide, title, box.x, box.y, box.w, box.h, box.fontSize);
  slide.addText('Creative presentation template', {
    x: 6.518, y: subtitleY, w: 3.838, h: 0.404,
    fontFace: BODY_FONT, fontSize: 18, color: RED,
    align: 'right', valign: 'top', wrap: false, autoFit: true,
  });
  footer(slide);
}

/* ------------------------------------------------------------------ *
 * Slides 2 - 23: heading + copy + photo collage                       *
 * Each builder lists its shapes back-to-front, as in the source deck. *
 * ------------------------------------------------------------------ */

function slide02(s) { // welcome
  pale(s, 0.763, 0.719, 2.52, 2.298);
  photo(s, 7.389, 0.719, 5.224, 6.063);
  heading(s, 'welcome', 2.651, 1.598, 2.672, 0.774);
  body(s, LOREM_LONG, 1.625, 3.003, 4.648, 3.241);
  mid(s, 6.668, 5.413, 1.501, 1.368);
  footer(s);
}

function slide03(s) { // about us
  pale(s, 1.673, 3.416, 2.52, 2.298);
  heading(s, 'about us', 5.27, 2.424, 2.793, 0.774);
  body(s, LOREM_LONG, 2.258, 3.822, 8.86, 1.485);
  mid(s, 8.794, 1.998, 1.501, 1.368);
  footer(s);
}

function slide04(s) { // history
  mid(s, 6.668, 5.413, 1.501, 1.368);
  pale(s, 10.093, 2.333, 2.52, 2.298);
  heading(s, 'history', 8.572, 1.781, 2.257, 0.774);
  body(s, LOREM_LONG, 7.376, 3.003, 4.648, 3.241);
  photo(s, 1.155, 0.719, 5.224, 6.063);
  footer(s);
}

function slide05(s) { // focus
  pale(s, 2.967, 0.719, 2.52, 2.298);
  heading(s, 'focus', 9.005, 1.781, 1.715, 0.774);
  body(s, LOREM_MED, 8.951, 2.854, 3.636, 1.839);
  footer(s);
  photo(s, 4.723, 2.147, 3.887, 3.205);
  mid(s, 8.164, 4.852, 1.501, 1.368);
}

function slide06(s) { // achievment
  mid(s, 8.089, 4.473, 1.501, 1.368);
  pale(s, 9.581, 2.175, 2.52, 2.298);
  heading(s, 'achievment', 4.722, 1.427, 3.591, 0.774);
  body(s, LOREM_MED, 5.6, 2.618, 5.398, 1.132);
  body(s, FUSCE, 5.6, 4.025, 5.398, 0.778);
  footer(s);
  photo(s, 1.474, 2.766, 3.887, 4.016);
}

function slide07(s) { // solution
  pale(s, 1.16, 0.719, 2.202, 2.008);
  heading(s, 'solution', 5.002, 1.421, 2.577, 0.774);
  body(s, LOREM_MED, 5.002, 2.49, 5.398, 1.132);
  mid(s, 11.117, 2.569, 1.501, 1.368);
  footer(s);
  photo(s, 4.62, 3.916, 7.998, 2.865);
}

function slide08(s) { // vision
  heading(s, 'vision', 8.95, 1.956, 1.929, 0.774);
  body(s, LOREM_MED, 8.558, 3.037, 2.714, 2.545);
  body(s, FUSCE, 3.99, 5.685, 7.446, 0.778);
  pale(s, 0.763, 4.515, 2.52, 2.298);
  mid(s, 11.112, 0.719, 1.501, 1.368);
  footer(s);
  photo(s, 3.99, 0.719, 3.887, 4.274);
}

function slide09(s) { // mission
  pale(s, 0.762, 2.341, 2.834, 2.818);
  heading(s, 'mission', 4.069, 0.836, 2.493, 0.774);
  body(s, LOREM_MED, 3.958, 1.892, 2.714, 2.545);
  body(s, FUSCE, 3.958, 4.619, 2.709, 1.485);
  mid(s, 11.696, 5.945, 0.917, 0.836);
  footer(s);
  photo(s, 7.219, 0.719, 3.887, 6.062);
}

function slide10(s) { // service
  pale(s, 0.764, 0.719, 2.99, 2.727);
  heading(s, 'service', 7.254, 1.309, 2.115, 0.774);
  body(s, LOREM_MED, 7.187, 2.397, 4.914, 1.485);
  body(s, FUSCE, 7.187, 4.021, 4.914, 0.767);
  photo(s, 0.762, 3.455, 5.905, 3.326);
  mid(s, 11.273, 5.58, 1.34, 1.202);
  footer(s);
}

function slide11(s) { // prestation
  pale(s, 1.156, 0.719, 2.106, 1.92);
  heading(s, 'prestation', 7.528, 0.994, 3.14, 0.774);
  body(s, LOREM_MED, 7.492, 2.017, 4.914, 1.485);
  mid(s, 7.06, 5.58, 1.26, 1.202);
  footer(s);
  body(s, FUSCE + 'pulvinar ultricies purus lectus ', 7.502, 3.523, 4.914, 1.12);
  photo(s, 1.155, 2.648, 5.905, 4.134);
}

function slide12(s) { // team
  pale(s, 0.762, 0.719, 5.905, 6.062);
  heading(s, 'team', 7.304, 2.023, 1.706, 0.774);
  body(s, LOREM_MED, 7.304, 3.046, 4.914, 1.485, RED_SOFT);
  mid(s, 11.353, 5.383, 1.26, 1.202);
  footer(s);
  photo(s, 1.57, 1.23, 4.289, 5.04);
}

function slide13(s) { // creator
  pale(s, 11.392, 0.722, 1.227, 1.149);
  heading(s, 'creator', 7.037, 2.128, 2.248, 0.774);
  body(s, LOREM_MED, 7.061, 3.007, 4.914, 1.485);
  mid(s, 1.147, 5.029, 1.712, 1.752);
  footer(s);
  body(s, FUSCE, 7.061, 4.598, 4.914, 0.767);
  photo(s, 2.151, 0.719, 3.887, 4.975);
}

function slide14(s) { // editor
  photo(s, 7.89, 3.747, 3.257, 3.06);
  mid(s, 6.251, 5.482, 1.639, 1.324);
  pale(s, 5.408, 0.722, 1.227, 1.149);
  heading(s, 'editor', 10.526, 0.96, 1.908, 0.774);
  body(s, LOREM_SHORT, 2.009, 4.101, 3.436, 1.12);
  photo(s, 2.151, 0.719, 3.257, 3.06);
  body(s, LOREM_SHORT, 7.89, 2.204, 3.257, 1.12);
  footer(s);
}

function slide15(s) { // talent
  photo(s, 6.51, 3.727, 3.257, 3.06);
  mid(s, 10.985, 0.719, 1.639, 1.324);
  pale(s, 0.762, 4.852, 2.783, 1.929);
  heading(s, 'talent', 10.132, 0.96, 1.932, 0.774);
  body(s, LOREM_AND + ' ', 5.841, 1.512, 2.402, 1.473);
  footer(s);
  body(s, LOREM_AND, 10.223, 4.521, 2.402, 1.473);
  photo(s, 2.1, 0.719, 3.257, 3.06);
}

function slide16(s) { // talent (3 photos)
  photo(s, 9.198, 0.719, 3.257, 3.06);
  mid(s, 8.133, 3.117, 1.933, 1.561);
  pale(s, 2.999, 0.719, 2.218, 2.098);
  heading(s, 'talent', 1.499, 5.302, 1.932, 0.774);
  body(s, LOREM_SHORT, 8.505, 5.129, 3.436, 1.12);
  body(s, LOREM_SHORT, 4.061, 5.129, 3.257, 1.12);
  photo(s, 3.99, 1.515, 3.257, 3.163);
  footer(s);
}

function slide17(s) { // portfolio (3 photos)
  photo(s, 5.141, 1.279, 3.613, 3.163);
  mid(s, 8.055, 3.356, 1.933, 1.561);
  pale(s, 0.761, 0.719, 2.218, 2.098);
  heading(s, 'portfolio', 2.573, 5.509, 2.625, 0.774);
  body(s, LOREM_SHORT, 5.47, 5.488, 5.749, 0.767);
  photo(s, 1.313, 1.236, 3.576, 3.163);
  photo(s, 9.006, 1.267, 3.576, 3.163);
  footer(s);
}

function slide18(s) { // portfolio (2 photos)
  photo(s, 9.033, 0.719, 3.594, 6.062);
  mid(s, 3.765, 1.895, 2.392, 2.419);
  pale(s, 1.208, 1.108, 2.557, 2.483);
  heading(s, 'portfolio', 2.1, 2.668, 2.625, 0.774);
  body(s, LOREM_SHORT, 1.785, 3.811, 3.255, 1.12);
  photo(s, 5.486, 2.884, 3.307, 3.897);
  footer(s);
}

function slide19(s) { // gallery (3 photos)
  photo(s, 0.998, 0.719, 3.738, 6.062);
  footer(s);
  mid(s, 10.939, 5.089, 1.674, 1.693);
  pale(s, 9.103, 4.376, 1.868, 1.814);
  heading(s, 'gallery', 9.2, 1.268, 2.299, 0.774);
  body(s, LOREM_SHORT, 8.722, 2.411, 3.255, 1.12);
  photo(s, 4.988, 0.719, 3.378, 2.813);
  photo(s, 4.988, 3.75, 3.378, 3.031);
}

function slide20(s) { // gallery (2 photos)
  photo(s, 8.888, 0.719, 3.738, 6.062);
  footer(s);
  heading(s, 'gallery', 3.044, 1.289, 2.299, 0.774);
  body(s, LOREM_SHORT, 1.883, 2.275, 4.62, 0.767);
  photo(s, 0.998, 3.75, 5.67, 3.031);
  mid(s, 7.815, 2.904, 1.616, 1.693);
  pale(s, 6.116, 4.144, 2.198, 2.134);
}

function slide21(s) { // gallery (2 wide photos)
  photo(s, 6.928, 0.732, 5.691, 3.154);
  pale(s, 7.856, 5.189, 1.639, 1.591);
  footer(s);
  body(s, LOREM_SHORT, 1.514, 5.246, 4.62, 0.767);
  heading(s, 'gallery', 2.675, 4.26, 2.299, 0.774);
  mid(s, 9.266, 4.144, 1.496, 1.262);
  photo(s, 0.998, 0.725, 5.67, 3.154);
}

function slide22(s) { // gallery (one banner photo)
  mid(s, 7.14, 0.733, 5.474, 3.332);
  pale(s, 0.996, 2.708, 0.946, 0.919);
  footer(s);
  body(s, LOREM_SHORT, 1.915, 1.942, 4.62, 0.767);
  heading(s, 'gallery', 3.045, 0.956, 2.299, 0.774);
  photo(s, 0.998, 3.627, 11.615, 3.154);
}

function slide23(s) { // gallery (wide + narrow photo)
  photo(s, 0.998, 2.805, 7.639, 3.979);
  footer(s);
  pale(s, 7.14, 0.719, 5.473, 0.774);
  body(s, LOREM_SHORT, 1.35, 2.104, 8.625, 0.413);
  mid(s, 1.035, 0.719, 1.149, 0.97);
  heading(s, 'gallery', 1.381, 1.122, 2.299, 0.774);
  photo(s, 8.897, 2.803, 3.716, 3.979);
}

/* ------------------------------------------------------------------ *
 * Slides 24 - 26: device mock-ups                                     *
 * ------------------------------------------------------------------ */

// Gray body of a device, drawn under its screen placeholder.
function deviceBody(slide, x, y, w, h, radius) {
  slide.addShape('roundRect', {
    x, y, w, h, rectRadius: radius,
    fill: { color: DEVICE_GRAY }, line: { type: 'none' },
  });
}

function slide24(s) { // mockup - phone
  pale(s, 6.668, 0.719, 5.945, 6.062);
  footer(s);
  body(s, LOREM_MOCKUP, 1.424, 2.672, 2.876, 1.827);
  mid(s, 0.762, 5.089, 0.931, 1.693);
  heading(s, 'mockup', 9.345, 3.199, 2.535, 0.774);
  deviceBody(s, 5.453, 1.124, 2.554, 5.356, 0.28);
  photo(s, 5.486, 1.151, 2.52, 5.315);
}

function slide25(s) { // mockup - tablet
  pale(s, 6.668, 0.719, 5.945, 6.062);
  mid(s, 0.762, 5.089, 0.931, 1.693);
  deviceBody(s, 1.36, 1.263, 6.452, 4.954, 0.22);
  photo(s, 1.47, 1.388, 6.221, 4.646);
  footer(s);
  body(s, LOREM_MOCKUP, 8.849, 3.039, 2.876, 1.827);
  heading(s, 'mockup', 9.019, 1.86, 2.535, 0.774);
}

function slide26(s) { // mockup - laptop
  mid(s, 4.377, 2.904, 0.931, 1.693);
  s.addShape('ellipse', { // base the screen rests on
    x: 4.135, y: 6.133, w: 9.144, h: 0.25,
    fill: { color: DEVICE_GRAY }, line: { type: 'none' },
  });
  deviceBody(s, 4.954, 1.166, 7.507, 5.078, 0.12);
  photo(s, 5.25, 1.512, 6.93, 4.365);
  pale(s, 0.762, 0.719, 3.622, 6.062);
  footer(s);
  body(s, LOREM_MOCKUP, 1.221, 3.039, 2.876, 1.827);
  heading(s, 'mockup', 1.391, 1.86, 2.535, 0.774);
}

/* ------------------------------------------------------------------ *
 * Slide 27: six-petal infographic wheel                               *
 * ------------------------------------------------------------------ */

// Petal outlines, straight from the source free-forms.
const PETALS = [
  { fill: 'C71313', x: 6.998, y: 2.651, w: 2.028, h: 2.806, vb: [537, 743],
    d: [[479, 401], [438, 329, 432, 246, 458, 173], [417, 217, 359, 243, 297, 243],
      [259, 243, 221, 233, 188, 214], [110, 169, 70, 84, 79, 0], [0, 191, 7, 415, 118, 607],
      [120, 610, 120, 610, 120, 610], [177, 709, 305, 743, 404, 685],
      [503, 628, 537, 501, 479, 402], [479, 401, 479, 401, 479, 401]] },
  { fill: '4C6574', x: 7.624, y: 4.327, w: 2.863, h: 1.926, vb: [758, 510],
    d: [[550, 94], [550, 94, 550, 94, 550, 94], [468, 94, 394, 57, 344, 0],
      [355, 38, 355, 78, 345, 117], [330, 174, 294, 222, 243, 251],
      [209, 271, 171, 281, 133, 281], [84, 281, 38, 265, 0, 236], [126, 402, 324, 509, 547, 510],
      [550, 510, 550, 510, 550, 510], [665, 510, 758, 417, 758, 302],
      [758, 187, 665, 94, 550, 94]] },
  { fill: 'FFC000', x: 9.91, y: 3.656, w: 2.269, h: 2.579, vb: [601, 683],
    d: [[467, 57], [367, 0, 240, 35, 183, 134], [183, 134, 183, 134, 183, 134],
      [142, 205, 74, 251, 0, 266], [94, 291, 164, 377, 164, 479], [164, 571, 107, 651, 26, 683],
      [234, 658, 429, 539, 542, 343], [544, 340, 544, 340, 544, 340],
      [601, 241, 566, 114, 467, 57]] },
  { fill: '4472C4', x: 10.382, y: 1.838, w: 2.028, h: 2.809, vb: [537, 744],
    d: [[57, 342], [98, 413, 104, 494, 80, 566], [121, 523, 178, 498, 239, 498],
      [277, 498, 315, 508, 348, 527], [399, 556, 435, 604, 451, 660],
      [458, 688, 460, 717, 457, 744], [537, 553, 530, 329, 419, 136],
      [417, 133, 417, 133, 417, 133], [359, 34, 232, 0, 133, 58], [34, 115, 0, 242, 57, 341],
      [57, 342, 57, 342, 57, 342]] },
  { fill: '70AD47', x: 8.92, y: 1.042, w: 2.859, h: 1.923, vb: [757, 509],
    d: [[412, 392], [427, 335, 463, 287, 514, 258], [547, 239, 585, 228, 624, 228],
      [673, 228, 719, 244, 757, 273], [632, 108, 434, 1, 211, 0], [207, 0, 207, 0, 207, 0],
      [93, 0, 0, 93, 0, 208], [0, 323, 93, 416, 207, 416], [208, 416, 208, 416, 208, 416],
      [290, 416, 363, 452, 413, 509], [402, 471, 401, 431, 412, 392]] },
  { fill: '950E0E', x: 7.228, y: 1.06, w: 2.273, h: 2.575, vb: [602, 682],
    d: [[435, 203], [435, 112, 491, 34, 570, 0], [364, 26, 171, 145, 59, 338],
      [57, 342, 57, 342, 57, 342], [0, 441, 35, 568, 134, 625], [233, 682, 360, 648, 418, 548],
      [418, 548, 418, 548, 418, 548], [459, 477, 527, 431, 602, 416],
      [506, 393, 435, 306, 435, 203]] },
];

// Caption sitting inside each petal.
const PETAL_LABELS = [
  { t: 'Content 1', x: 7.715, y: 2.572, w: 0.861 },
  { t: 'Content 6', x: 7.633, y: 4.488, w: 0.879 },
  { t: 'Content 5', x: 9.069, y: 5.537, w: 0.865 },
  { t: 'Content 4', x: 10.681, y: 4.782, w: 0.873 },
  { t: 'Content 3', x: 10.96, y: 2.995, w: 0.873 },
  { t: 'Content 1', x: 9.412, y: 1.901, w: 0.861 },
];

// A pointed leaf: two mirrored cubic arcs between opposite corners.
function leaf(slide, x, y, w, h, rotate) {
  slide.addShape('custGeom', {
    x, y, w, h, rotate, fill: { color: WHITE }, line: { type: 'none' },
    points: [
      { x: 0, y: h },
      { x: w, y: 0, curve: { type: 'cubic', x1: 0, y1: h * 0.4, x2: w * 0.4, y2: 0 } },
      { x: 0, y: h, curve: { type: 'cubic', x1: w, y1: h * 0.6, x2: w * 0.6, y2: h } },
      { close: true },
    ],
  });
}

// The six white pictograms, approximated with native shapes.
function petalIcons(s) {
  const draw = (shape, x, y, w, h, fill, extra) =>
    s.addShape(shape, Object.assign(
      { x, y, w, h, fill: { color: fill }, line: { type: 'none' } }, extra));

  // leaf pair - dark red petal
  leaf(s, 7.875, 2.213, 0.28, 0.28, 315);
  leaf(s, 8.11, 2.077, 0.31, 0.4, 20);
  // apple - green petal
  draw('ellipse', 9.669, 1.545, 0.348, 0.299, WHITE);
  draw('rect', 9.833, 1.456, 0.02, 0.12, WHITE, { rotate: 12 });
  leaf(s, 9.87, 1.456, 0.146, 0.1, 15);
  // framed picture - blue petal
  draw('roundRect', 11.206, 2.369, 0.416, 0.416, WHITE, { rectRadius: 0.07 });
  draw('rect', 11.262, 2.425, 0.304, 0.304, '4472C4');
  draw('ellipse', 11.295, 2.458, 0.08, 0.08, WHITE);
  draw('triangle', 11.312, 2.58, 0.205, 0.149, WHITE);
  // crosshair target - red petal
  draw('donut', 7.756, 3.866, 0.481, 0.481, WHITE);
  draw('ellipse', 7.929, 4.039, 0.135, 0.135, WHITE);
  draw('rect', 7.719, 4.088, 0.555, 0.036, WHITE);
  draw('rect', 7.978, 3.829, 0.036, 0.555, WHITE);
  // document with bookmark - amber petal
  draw('snip1Rect', 10.933, 4.183, 0.361, 0.43, WHITE);
  draw('rect', 11.075, 4.183, 0.082, 0.185, 'FFC000');
  draw('triangle', 11.075, 4.306, 0.082, 0.062, 'FFC000', { rotate: 180 });
  // microscope - slate petal: C-shaped arm, angled eyepiece, stage and base
  draw('blockArc', 9.395, 5.0, 0.305, 0.305, WHITE,
    { angleRange: [270, 90], arcThicknessRatio: 0.32 });
  draw('roundRect', 9.44, 4.983, 0.075, 0.26, WHITE, { rotate: 32, rectRadius: 0.03 });
  draw('rect', 9.49, 5.19, 0.075, 0.09, WHITE);
  draw('rect', 9.339, 5.28, 0.23, 0.04, WHITE);
  draw('rect', 9.44, 5.3, 0.05, 0.06, WHITE);
  draw('rect', 9.339, 5.36, 0.361, 0.04, WHITE);
}

function slide27(s) { // infographic
  mid(s, 9.716, 3.264, 2.897, 3.517);
  pale(s, 0.967, 0.719, 8.749, 2.545);
  footer(s);
  body(s, LOREM_MOCKUP, 1.896, 1.925, 4.701, 1.12);
  heading(s, 'infographic', 1.924, 0.933, 3.603, 0.774);

  PETALS.forEach(p => freeform(s, p));
  PETAL_LABELS.forEach(l => s.addText(l.t, {
    x: l.x, y: l.y, w: l.w, h: 0.269,
    fontFace: HEAD_FONT, fontSize: 10, bold: true, color: WHITE,
    align: 'center', valign: 'top', wrap: false, autoFit: true,
  }));
  petalIcons(s);

  // Two columns of numbered notes under the pale banner.
  for (let i = 0; i < 6; i += 1) {
    const col = i < 3 ? 1.36 : 3.985;
    const row = (i % 3) * 1.027;
    s.addText('content ' + (i + 1), {
      x: col, y: 3.674 + row, w: 1.0, h: 0.303,
      fontFace: HEAD_FONT, fontSize: 12, bold: true, color: NAVY,
      align: 'center', valign: 'top', wrap: false, autoFit: true,
    });
    s.addText(SUITABLE, {
      x: col, y: 3.951 + row, w: 1.965, h: 0.631,
      fontFace: BODY_FONT, fontSize: 10.5, color: RED,
      align: 'justify', valign: 'top',
    });
  }
}

/* ------------------------------------------------------------------ *
 * Slides 28 - 30                                                      *
 * ------------------------------------------------------------------ */

function slide28(s) { // Colours palette
  mid(s, 11.476, 0.728, 1.137, 0.895);
  footer(s);
  s.addText('Colours palette', {
    x: 8.262, y: 0.874, w: 4.955, h: 0.707,
    fontFace: HEAD_FONT, fontSize: 36, color: GRAY_25, align: 'left', valign: 'top',
  });
  SWATCHES.forEach((sw, i) => {
    const x = 1.239 + i * 2.714;
    s.addShape('rect', { x, y: 3.699, w: 2.714, h: 0.472, fill: { color: sw.fill }, line: { type: 'none' } });
  });
  s.addText('Colours style used for design', {
    x: 1.224, y: 2.493, w: 2.714, h: 0.774,
    fontFace: HEAD_FONT, fontSize: 20, color: GRAY_25, align: 'left', valign: 'top',
  });
  const labelX = [1.224, 4.016, 6.693, 9.449];
  SWATCHES.forEach((sw, i) => {
    s.addText(sw.name, {
      x: labelX[i], y: 4.357, w: 1.979, h: 0.774,
      fontFace: BODY_FONT, fontSize: 20, color: GRAY_25, align: 'left', valign: 'top',
    });
    s.addText(sw.hex, {
      x: labelX[i], y: 5.201, w: 0.9, h: 0.337,
      fontFace: BODY_FONT, fontSize: 14, color: GRAY_50,
      align: 'left', valign: 'top', wrap: false, autoFit: true,
    });
    s.addText(sw.rgb, {
      x: labelX[i], y: 5.64, w: 1.708, h: 0.337,
      fontFace: BODY_FONT, fontSize: 14, color: GRAY_50,
      align: 'left', valign: 'top', wrap: false, autoFit: true,
    });
  });
}

function slide29(s) { // pull quote
  pale(s, 3.334, 0.719, 6.667, 6.062);
  footer(s);
  s.addText('"Don\'t be afraid to give up the good to go for the great."', {
    x: 3.334, y: 2.517, w: 6.667, h: 1.919,
    fontFace: BODY_FONT, fontSize: 36, italic: true, color: NAVY,
    align: 'center', valign: 'top', autoFit: true,
  });
  s.addText('--John D. Rockefeller', {
    x: 5.079, y: 4.579, w: 3.175, h: 0.404,
    fontFace: HEAD_FONT, fontSize: 18, color: NAVY_SOFT,
    align: 'center', valign: 'top', wrap: false, autoFit: true,
  });
  mid(s, 11.664, 0.719, 0.949, 0.747);
  mid(s, 0.762, 6.033, 0.949, 0.747);
}

/* ------------------------------------------------------------------ *
 * Assembly                                                            *
 * ------------------------------------------------------------------ */

const BUILDERS = [
  s => coverSlide(s, 'Chielo', { x: 3.543, y: 2.538, w: 6.25, h: 2.423, fontSize: 138 }, 4.82),
  slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
  slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17,
  slide18, slide19, slide20, slide21, slide22, slide23, slide24, slide25,
  slide26, slide27, slide28, slide29,
  s => coverSlide(s, 'Thanks', { x: 3.268, y: 2.737, w: 6.799, h: 2.036, fontSize: 115 }, 4.695),
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'CHIELO', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'CHIELO';
  pptx.author = 'chielo';
  pptx.title = 'Chielo - Creative presentation template';

  BUILDERS.forEach(builder => {
    const slide = pptx.addSlide();
    slide.background = { color: WHITE };
    builder(slide);
  });

  return pptx.writeFile({
    fileName: path.join(__dirname, '146b6b9b-dd62-47b5-9006-bdf7fc4a1276_grok_final.pptx'),
  });
}

build().then(f => console.log('wrote ' + f)).catch(err => {
  console.error(err);
  process.exit(1);
});
