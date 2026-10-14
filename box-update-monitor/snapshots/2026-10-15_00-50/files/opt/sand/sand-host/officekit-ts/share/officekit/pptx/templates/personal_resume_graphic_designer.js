/**
 * Amatya - Personal Resume  |  30 slides, 13.333 x 7.5 in (16:9)
 * Rebuilt with pptxgenjs. Photos in the source deck are flat grey
 * placeholder boxes; they are re-created here as grey rounded shapes.
 *
 *   node 0fd4d301-e53b-4b2e-84f9-5d418807f7ce_grok_final.js
 */
'use strict';
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette
const C = {
  blue:   '19C5FE',  // accent
  ink:    '242527',  // body text / near-black
  dark:   '393A3E',  // dark panels
  gray:   '585B60',  // muted text
  body:   '6D7076',  // paragraph grey
  white:  'FFFFFF',
  bar:    'F2F2F2',  // top navigation bar
  track:  'E8E9EA',  // progress track
  line2:  'D9D9D9',  // hairline / light copy
  hexLn:  'C9C9C9',  // hexagon outlines
  mist:   'C6C8CB',
  photo:  'CCCCCC',  // image placeholder grey
  mid:    '8E9197',
  slate:  '5C5C5C',
  slate2: '535353',
  silver: 'A7A7A7',
  sage:   'D2DDD9',
  coal:   '3F3F3F',
  stone:  '808080',
};
const F = { head: 'Poppins', body: 'Open Sans' };

// Body copy of the template - the same lorem sentences repeat throughout.
const L = {
  L1: 'Nunc venenatis, elit vel facilisis blandit, purus turpis dictum diam, faucibus volutpat mauris nisl sed nulla. Quisque tempus luctus feugiat. Suspendisse at cursus ipsum.  Praesent pharetra elit et gravida blandit.  ',
  L2: 'Nunc venenatis, elit vel facilisis blandit, purus turpis dictum diam, faucibus volutpat  ',
  L3: 'Nunc venenatis, elit vel facilisis blandit, purus turpis dictum diam, faucibus volutpat mauris nisl sed nulla. Quisque tempus luctus feugiat. Suspendisse at cursus ipsum.   ',
  L4: 'Nunc venenatis, elit vel facilisis blandit, purus turpis dictum diam, faucibus volutpat mauris nisl sed nulla. Quisque tempus luctus feugiat. Suspendisse at cursus  ',
  L5: 'Nunc venenatis, elit vel facilisis blandit, purus turpis dictum diam, faucibus volutpat mauris nisl sed nulla. Quisque tempus luctus  ',
  L6: 'Nunc venenatis, elit vel facilisis blandit, purus turpis dictum diam, faucibus volutpat mauris nisl sed nulla.  ',
  L7: 'Nunc venenatis, elit vel facilisis blandit, purus turpis dictum diam, faucibus volutpat mauris  ',
  L8: 'Nunc venenatis, elit vel facilisis blandit, purus turpis dictum diam, faucibus volutpat mauris nisl sed nulla. Quisque tempus luctus feugiat. ',
  L9: 'Nunc venenatis, elit vel facilisis blandit, purus turpis dictum diam, faucibus volutpat mauris nisl sed nulla. Quisque tempus luctus feugiat. Suspendisse at  ',
  L10: 'Nunc venenatis, elit vel facilisis blandit, purus turpis dictum diam, faucibus volutpat ',
  L11: 'Nunc venenatis, elit vel facilisis blandit, purus turpis dictum diam, faucibus volutpat mauris nisl sed nulla. Quisque tempus luctus feugiat. Suspendisse at cursus ipsum.  Praesent pharetra elit et gravida  ',
  L12: 'Nunc venenatis, elit vel facilisis blandit, purus turpis dictum diam. Lorem ipsum ',
  L13: 'Nunc venenatis, elit vel facilisis blandit, purus turpis dictum diam, faucibus volutpat mauris nisl sed nulla. Quisque tempus luctus feugiat.  ',
  L14: 'Nunc venenatis, elit vel facilisis blandit, purus turpis dictum diam, faucibus volutpat mauris nisl sed nulla. Quisque  ',
  L15: 'Nunc venenatis, elit vel facilisis blandit, purus turpis dictum diam',
};

// ------------------------------------------------------------ tiny helpers
// Poppins label / heading
const t = (s, str, x, y, w, h, o = {}) =>
  s.addText(str, { x, y, w, h, fontFace: F.head, fontSize: 12, color: C.ink,
                   valign: 'top', wrap: false, fit: 'resize', ...o });

// Open Sans paragraph (150% leading is the template default)
const p = (s, str, x, y, w, h, o = {}) =>
  s.addText(str, { x, y, w, h, fontFace: F.body, fontSize: 10, color: C.body,
                   valign: 'top', lineSpacingMultiple: 1.5, fit: 'resize', ...o });

// any autoshape
const box = (s, shape, x, y, w, h, o = {}) => s.addShape(shape, { x, y, w, h, ...o });

// flat grey stand-in for a photo placeholder (the template's photos are
// plain grey boxes with softly rounded corners)
const photo = (s, x, y, w, h, r = 0) =>
  box(s, r ? 'roundRect' : 'rect', x, y, w, h,
      { fill: { color: C.photo }, rectRadius: r, objectName: '[image]' });

// the two small outlined circles used above every section title
const dots = (s, x, y) => {
  box(s, 'ellipse', x, y, 0.137, 0.137, { line: { color: C.blue, width: 0.75 } });
  box(s, 'ellipse', x + 0.296, y, 0.137, 0.137, { line: { color: C.blue, width: 0.75 } });
};

// faint hexagon outline used as background decoration
const hex = (s, x, y, w, h, rot) =>
  box(s, 'hexagon', x, y, w, h, { rotate: rot, line: { color: C.hexLn, width: 0.75 } });

// the pair of hexagons that bleeds off the top-right corner
const hexPair = s => {
  hex(s, 11.964, -0.556, 4.974, 4.288, 351.2);
  hex(s, 10.738, -1.541, 4.974, 4.288, 320.49);
};

// ------------------------------------------------- persistent page chrome
// [label, x, width, active?] - the highlighted item changes per section
const NAV = {
  NAV1: [['About Me', 8.172, 1.124, 1], ['Skills', 9.642, 0.63], ['Portfolio', 10.706, 0.933], ['Contact', 12.002, 0.935]],
  NAV2: [['About Me', 8.1, 1.068], ['Services', 9.431, 0.951], ['Portfolio', 10.671, 1.005, 1], ['Contact', 11.898, 0.935]],
  NAV3: [['About Me', 8.072, 1.124, 1], ['Services', 9.431, 0.951], ['Portfolio', 10.706, 0.933], ['Contact', 11.898, 0.935]],
  NAV4: [['About Me', 8.1, 1.068], ['Services', 9.401, 1.01, 1], ['Portfolio', 10.706, 0.933], ['Contact', 11.898, 0.935]],
};

function chrome(s, nav, page, dark) {
  const dy = page >= 22 && page <= 28 ? 0.002 : 0;   // this run sits 0.002" lower
  box(s, 'rect', 0, dy, 13.333, 0.824, { fill: { color: C.bar } });
  t(s, [{ text: 'Personal ', options: { bold: true } }, { text: 'Resume' }],
    1.144, 0.243 + dy, 1.952, 0.337, { fontSize: 14 });
  nav.forEach(([label, x, w2, on]) =>
    t(s, label, x, 0.258 + dy, w2, 0.32, { fontSize: 13, bold: !!on, color: C.gray, align: 'center' }));
  box(s, 'roundRect', 0.511, 0.167, 0.475, 0.455, { fill: { color: C.blue }, rectRadius: 0.086 });
  t(s, 'Logo', 0.485, 0.229, 0.514, 0.309, { fontSize: 9, bold: true, align: 'center', lineSpacingMultiple: 1.5 });
  // bottom-right ribbon + page badge
  box(s, 'custGeom', 9.018, 7.062, 4.333, 0.438, { fill: { color: C.blue },
    points: [{ x: 0.207, y: 0 }, { x: 4.333, y: 0 }, { x: 4.333, y: 0.438 }, { x: 0, y: 0.438 }, { close: true }] });
  t(s, 'Penelope Doe', 9.56, 7.164, 1.422, 0.303, { bold: true, color: C.white });
  t(s, 'Resume Presentation', 11.024, 7.187, 1.529, 0.252, { fontSize: 9, color: C.white });
  box(s, 'roundRect', 12.697, 6.844, 0.475, 0.455, { fill: { color: dark ? C.white : C.gray }, rectRadius: 0.086,
    shadow: { type: 'outer', blur: 13, offset: 3, angle: 90, color: '000000', opacity: 0.38 } });
  t(s, String(page), 12.622, 6.842, 0.626, 0.459, { fontFace: F.body, fontSize: 11, bold: true,
    color: dark ? C.dark : C.white, align: 'center', valign: 'middle', fit: 'none' });
}

// Slide 26 embeds a small clustered column chart (5 one-point series).
const BAR_SERIES = [1, 2, 3, 4, 5].map((i, k) => ({
  name: 'Series ' + i, labels: ['Category 1'], values: [[4.3, 2.4, 2, 3, 4][k]],
}));

// ================================================================ slides

// --- 01  cover
function slide01(s) {
  photo(s, 0, 0, 13.333, 7.5);
  box(s, 'rect', 0, 0, 13.333, 7.5, { fill: { color: C.body, transparency: 25 } });
  box(s, 'rect', 4.382, 2.471, 8.951, 2.779, { fill: { color: C.white }, line: { color: C.mist, width: 0.75 } });
  t(s, 'PERSONAL RESUME', 5.245, 4.168, 3.221, 0.337, { fontSize: 14, color: C.gray, charSpacing: 6 });
  t(s, 'Amatya ', 5.197, 3.343, 3.391, 0.942, { fontSize: 50, bold: true, lineSpacing: 60 });
  box(s, 'triangle', 7.436, 3.216, 0.659, 0.215, { fill: { color: C.blue } });
  p(s, L.L1, 8.738, 3.406, 4.177, 1.084);
  box(s, 'custGeom', 11.999, 2.48, 1.335, 0.828, { fill: { color: C.white }, line: { color: C.mist, width: 0.75 }, points: [{ x: 0, y: 0 }, { x: 1.335, y: 0 }, { x: 1.335, y: 0.772 }, { x: 1.267, y: 0.828 }, { x: 0.068, y: 0.296 }, { close: true }] });
  box(s, 'custGeom', 12.442, 2.471, 0.891, 1.247, { line: { color: C.mist, width: 0.75 }, points: [{ x: 0.057, y: 0 }, { x: 0.891, y: 0 }, { x: 0.891, y: 1.227 }, { x: 0.759, y: 1.247 }, { x: 0, y: 0.178 }, { close: true }] });
  box(s, 'custGeom', 4.382, 4.5, 0.982, 0.75, { line: { color: C.mist, width: 0.75 }, points: [{ x: 0.45, y: 0 }, { x: 0.982, y: 0.75 }, { x: 0, y: 0.75 }, { x: 0, y: 0.07 }, { close: true }] });
}

// --- 02  about me
function slide02(s) {
  photo(s, 1.743, 1.755, 4.856, 4.647, 0.183);
  hexPair(s);
  chrome(s, NAV.NAV1, 2);
  box(s, 'roundRect', 1.217, 1.763, 0.966, 4.654, { fill: { color: C.dark } });
  t(s, 'About Me', 0.935, 3.868, 1.618, 0.438, { fontSize: 20, bold: true, color: C.white, align: 'center', rotate: 270 });
  t(s, 'Penelope Doe', 7.21, 2.157, 3.247, 0.606, { fontSize: 30, bold: true, lineSpacing: 36 });
  dots(s, 7.359, 1.909);
  p(s, L.L1, 7.224, 3.086, 5.127, 0.831);
  box(s, 'roundRect', 7.357, 4.409, 5.407, 1.94, { line: { color: C.line2, width: 0.75 }, rectRadius: 0.227 });
  box(s, 'triangle', 1.481, 2, 0.438, 0.143, { fill: { color: C.blue } });
  box(s, 'roundRect', 7.649, 4.795, 0.458, 0.438, { fill: { color: C.blue }, rectRadius: 0.072 });
  t(s, 'GRAPHIC DESIGNER', 7.224, 2.646, 2.956, 0.37, { fontSize: 16, charSpacing: 3 });
  t(s, '01', 7.675, 4.886, 0.383, 0.303, { bold: true, color: C.white, align: 'center' });
  t(s, 'Graphic Designer', 8.16, 4.894, 1.753, 0.303, { bold: true, lineSpacing: 14.4 });
  p(s, L.L6, 7.562, 5.283, 2.655, 0.758, { fontSize: 9 });
  box(s, 'roundRect', 10.164, 4.795, 0.458, 0.438, { fill: { color: C.blue }, rectRadius: 0.072 });
  t(s, '02', 10.19, 4.886, 0.409, 0.303, { bold: true, color: C.white, align: 'center' });
  t(s, 'Web Designer', 10.676, 4.894, 1.445, 0.303, { bold: true, lineSpacing: 14.4 });
  p(s, L.L6, 10.078, 5.283, 2.655, 0.758, { fontSize: 9 });
  hex(s, -2.992, 6.405, 4.974, 4.288, 351.2);
  box(s, 'rect', 1.939, 1.767, 0.255, 4.647, { fill: { color: C.dark } });
}

// --- 03  table of contents
function slide03(s) {
  chrome(s, NAV.NAV1, 3);
  t(s, 'Table Of Contents', 1.045, 2.175, 4.209, 0.606, { fontSize: 30, bold: true });
  p(s, L.L1, 1.105, 2.781, 5.127, 0.831);
  hex(s, -2.992, 6.405, 4.974, 4.288, 351.2);
  t(s, 'About Me', 8.321, 1.76, 1.266, 0.353, { fontSize: 15, bold: true, lineSpacing: 18 });
  p(s, L.L2, 8.321, 1.972, 3.075, 0.579);
  box(s, 'line', 7.096, 2.685, 5.239, 0, { line: { color: C.track, width: 0.75 } });
  t(s, 'Skills', 8.321, 2.895, 0.798, 0.353, { fontSize: 15, bold: true, lineSpacing: 18 });
  p(s, L.L2, 8.321, 3.109, 3.075, 0.579);
  box(s, 'line', 7.096, 3.926, 5.239, 0, { line: { color: C.track, width: 0.75 } });
  t(s, 'Portfolio', 8.321, 4.138, 1.131, 0.353, { fontSize: 15, bold: true, lineSpacing: 18 });
  p(s, L.L2, 8.321, 4.354, 3.075, 0.579);
  box(s, 'line', 7.096, 5.171, 5.239, 0, { line: { color: C.track, width: 0.75 } });
  t(s, 'Contact', 8.321, 5.376, 1.096, 0.353, { fontSize: 15, bold: true, lineSpacing: 18 });
  p(s, L.L2, 8.321, 5.584, 3.075, 0.579);
  t(s, 'Page', 11.752, 1.794, 0.651, 0.303, { bold: true, align: 'center' });
  t(s, '01', 11.801, 2.083, 0.502, 0.438, { fontSize: 20, bold: true, color: C.blue, align: 'center' });
  t(s, 'Page', 11.752, 3.019, 0.651, 0.303, { bold: true, align: 'center' });
  t(s, '02', 11.78, 3.308, 0.544, 0.438, { fontSize: 20, bold: true, color: C.blue, align: 'center' });
  t(s, 'Page', 11.776, 4.207, 0.651, 0.303, { bold: true, align: 'center' });
  t(s, '03', 11.799, 4.496, 0.554, 0.438, { fontSize: 20, bold: true, color: C.blue, align: 'center' });
  t(s, 'Page', 11.776, 5.479, 0.651, 0.303, { bold: true, align: 'center' });
  t(s, '04', 11.789, 5.767, 0.574, 0.438, { fontSize: 20, bold: true, color: C.blue, align: 'center' });
  photo(s, 1.118, 3.865, 5.127, 2.333, 0.092);
  photo(s, 7.114, 2.857, 0.853, 0.817, 0.105);
  photo(s, 7.114, 4.103, 0.853, 0.817, 0.105);
  photo(s, 7.114, 5.332, 0.853, 0.817, 0.105);
  photo(s, 7.114, 1.721, 0.853, 0.817, 0.105);
}

// --- 04  my biodata
function slide04(s) {
  box(s, 'rect', -0.053, 2.947, 4.737, 4.579, { fill: { color: C.dark } });
  hex(s, -2.487, 5.773, 4.974, 4.288, 351.2);
  chrome(s, NAV.NAV1, 4);
  box(s, 'roundRect', 7.146, 3.835, 5.407, 2.678, { line: { color: C.line2, width: 0.75 }, rectRadius: 0.212 });
  t(s, 'My Biodata', 7.146, 2.148, 2.749, 0.606, { fontSize: 30, bold: true, lineSpacing: 36 });
  dots(s, 7.243, 1.843);
  t(s, 'Name', 8.214, 4.297, 0.74, 0.303, { bold: true });
  p(s, 'Penelope Doe', 9.321, 4.293, 1.301, 0.303, { fontSize: 12, color: C.gray, lineSpacingMultiple: 1, wrap: false });
  t(s, 'Birth', 8.214, 4.668, 0.63, 0.303, { bold: true });
  p(s, '4 March 1997', 9.321, 4.664, 1.266, 0.303, { fontSize: 12, color: C.gray, lineSpacingMultiple: 1, wrap: false });
  t(s, 'Address', 8.214, 5.013, 0.921, 0.303, { bold: true });
  p(s, 'Lorem ipsum, dolor sit amet 12345  ', 9.321, 5.009, 3.026, 0.303, { fontSize: 12, lineSpacingMultiple: 1, wrap: false });
  t(s, 'Email', 8.214, 5.366, 0.707, 0.303, { bold: true });
  p(s, 'loremipsum@com', 9.321, 5.361, 1.636, 0.303, { fontSize: 12, color: C.gray, lineSpacingMultiple: 1, wrap: false });
  t(s, 'Phone', 8.214, 5.738, 0.761, 0.303, { bold: true });
  p(s, '123-456-789', 9.321, 5.734, 1.168, 0.303, { fontSize: 12, color: C.gray, lineSpacingMultiple: 1, wrap: false });
  p(s, L.L1, 7.146, 2.743, 5.407, 0.831);
  box(s, 'roundRect', 7.146, 4.293, 0.529, 1.899, { fill: { color: C.dark }, rectRadius: 0.179 });
  box(s, 'rect', 7.146, 4.293, 0.32, 1.899, { fill: { color: C.dark } });
  t(s, 'Biodata', 6.958, 5.023, 0.903, 0.303, { bold: true, color: C.white, rotate: 270 });
  photo(s, 0.999, 1.895, 5.583, 4.647, 0.183);
}

// --- 05  a short story about me
function slide05(s) {
  box(s, 'rect', 0, 0.865, 6.667, 6.635, { fill: { color: C.dark } });
  hex(s, -2.487, 5.773, 4.974, 4.288, 351.2);
  hexPair(s);
  hex(s, 11.964, -0.556, 4.974, 4.288, 351.2);
  chrome(s, NAV.NAV1, 5);
  t(s, [{ text: 'A Short Story', options: { breakLine: true } }, { text: 'About Me' }], 8.234, 3.018, 3.121, 1.111, { fontSize: 30, bold: true, lineSpacing: 36 });
  dots(s, 8.384, 2.77);
  p(s, L.L1, 8.234, 4.249, 4, 1.084);
  photo(s, 1.144, 1.859, 6.503, 4.647, 0.183);
}

// --- 06  my expertise
function slide06(s) {
  box(s, 'rect', 0, 0.865, 5.667, 6.635, { fill: { color: C.dark } });
  chrome(s, NAV.NAV1, 6);
  t(s, 'Graphic Design', 7.056, 4.804, 1.922, 0.353, { fontSize: 15, bold: true, align: 'center' });
  t(s, 'My Expertise', 1.144, 3.255, 3.017, 0.606, { fontSize: 30, bold: true, color: C.white });
  dots(s, 1.293, 3.007);
  p(s, L.L1, 1.144, 3.962, 3.197, 1.336, { color: C.line2 });
  p(s, L.L4, 6.484, 5.182, 3.056, 1.084, { align: 'center' });
  t(s, 'Web Design', 10.422, 4.804, 1.534, 0.353, { fontSize: 15, bold: true, align: 'center' });
  p(s, L.L4, 9.656, 5.182, 3.056, 1.084, { align: 'center' });
  hex(s, -1.82, 6.096, 3.962, 3.415, 351.2);
  photo(s, 9.77, 1.818, 2.897, 2.649, 0.182);
  photo(s, 6.53, 1.82, 2.897, 2.649, 0.182);
}

// --- 07  hobbies & interests
function slide07(s) {
  s.background = { color: C.white };
  box(s, 'rect', -0.038, 0.698, 7.723, 6.853, { fill: { color: C.dark } });
  chrome(s, NAV.NAV1, 7);
  t(s, 'Photography', 9.377, 2.313, 1.659, 0.353, { fontSize: 15, bold: true });
  p(s, L.L4, 9.377, 2.666, 3.282, 1.084);
  t(s, 'Movie Enthusiast', 9.377, 4.541, 2.122, 0.353, { fontSize: 15, bold: true });
  p(s, L.L4, 9.377, 4.894, 3.282, 1.084);
  box(s, 'roundRect', -0.047, 3.125, 0.529, 2.247, { fill: { color: C.white }, rectRadius: 0.179 });
  box(s, 'rect', -0.064, 3.125, 0.414, 2.247, { fill: { color: C.white } });
  t(s, 'Hobbies & Interest', -0.739, 4.067, 1.839, 0.303, { bold: true, color: C.dark, rotate: 270 });
  hex(s, -1.315, 6.529, 3.596, 3.1, 351.2);
  t(s, 'Hobbies &  Interests', 1.144, 3.383, 4.653, 0.606, { fontSize: 30, bold: true, color: C.white });
  dots(s, 1.293, 3.135);
  p(s, L.L1, 1.144, 3.962, 4.718, 1.084, { color: C.line2 });
  photo(s, 6.605, 1.956, 2.025, 2.121, 0.139);
  photo(s, 6.607, 4.307, 2.025, 2.121, 0.139);
  box(s, 'roundRect', 6.271, 6.083, 0.672, 0.643, { fill: { color: C.blue }, rectRadius: 0.106 });
  t(s, '02', 6.378, 6.254, 0.46, 0.353, { fontSize: 15, bold: true, color: C.white });
  box(s, 'roundRect', 8.292, 1.676, 0.672, 0.643, { fill: { color: C.blue }, rectRadius: 0.106 });
  t(s, '01', 8.414, 1.833, 0.428, 0.353, { fontSize: 15, bold: true, color: C.white });
}

// --- 08  my education
function slide08(s) {
  box(s, 'rect', -0.038, 0.698, 4.678, 6.853, { fill: { color: C.dark } });
  chrome(s, NAV.NAV1, 8);
  t(s, 'My Education', 0.964, 4.485, 2.337, 0.454, { fontSize: 21, bold: true, color: C.track, align: 'center' });
  p(s, L.L3, 0.563, 4.935, 3.113, 1.084, { color: C.line2, align: 'center' });
  t(s, '2010-2011', 7.17, 2.155, 1.245, 0.353, { fontSize: 15, bold: true });
  p(s, 'School Name', 7.126, 2.425, 1.326, 0.32, { fontSize: 13, color: C.blue, align: 'center', lineSpacingMultiple: 1, wrap: false });
  p(s, L.L3, 8.64, 1.969, 4.057, 0.831, { color: C.gray });
  box(s, 'roundRect', 5.622, 2.385, 0.169, 0.162, { fill: { color: C.blue }, rectRadius: 0.027 });
  box(s, 'roundRect', 5.638, 3.518, 0.169, 0.162, { fill: { color: C.blue }, rectRadius: 0.027 });
  box(s, 'roundRect', 5.638, 4.744, 0.169, 0.162, { fill: { color: C.blue }, rectRadius: 0.027 });
  box(s, 'roundRect', 5.638, 5.845, 0.169, 0.162, { fill: { color: C.blue }, rectRadius: 0.027 });
  t(s, '2011-2012', 7.17, 3.294, 1.229, 0.353, { fontSize: 15, bold: true });
  p(s, 'School Name', 7.17, 3.565, 1.326, 0.32, { fontSize: 13, color: C.blue, align: 'center', lineSpacingMultiple: 1, wrap: false });
  p(s, L.L3, 8.64, 3.108, 4.02, 0.831, { color: C.gray });
  t(s, '2012-2013', 7.17, 4.435, 1.268, 0.353, { fontSize: 15, bold: true });
  p(s, 'School Name', 7.129, 4.706, 1.326, 0.32, { fontSize: 13, color: C.blue, align: 'center', lineSpacingMultiple: 1, wrap: false });
  p(s, L.L3, 8.64, 4.249, 4.02, 0.831, { color: C.gray });
  t(s, '2013-2014', 7.17, 5.576, 1.289, 0.353, { fontSize: 15, bold: true });
  p(s, 'School Name', 7.12, 5.847, 1.326, 0.32, { fontSize: 13, color: C.blue, align: 'center', lineSpacingMultiple: 1, wrap: false });
  p(s, L.L3, 8.64, 5.39, 4.02, 0.831, { color: C.gray });
  hex(s, -1.315, 6.529, 3.596, 3.1, 351.2);
  photo(s, 1.083, 2.229, 2.097, 2.007, 0.259);
  photo(s, 6.196, 3.23, 0.777, 0.743, 0.096);
  photo(s, 6.196, 4.375, 0.777, 0.743, 0.096);
  photo(s, 6.196, 5.516, 0.777, 0.743, 0.096);
  photo(s, 6.196, 2.094, 0.777, 0.743, 0.096);
}

// --- 09  university
function slide09(s) {
  hexPair(s);
  box(s, 'rect', -0.038, 0.698, 4.678, 6.853, { fill: { color: C.dark } });
  chrome(s, NAV.NAV1, 9);
  box(s, 'line', 5.627, 3.297, 0, 3.144, { line: { color: C.line2, width: 0.75 } });
  t(s, '2014-2015', 7.088, 3.437, 1.298, 0.353, { fontSize: 15, bold: true });
  p(s, 'University Name', 7.056, 3.708, 1.389, 0.286, { fontSize: 11, color: C.blue, align: 'center', lineSpacingMultiple: 1, wrap: false });
  p(s, L.L3, 8.558, 3.251, 4.057, 0.831, { color: C.gray });
  box(s, 'roundRect', 5.541, 3.668, 0.169, 0.162, { fill: { color: C.blue }, rectRadius: 0.027 });
  box(s, 'roundRect', 5.556, 4.8, 0.169, 0.162, { fill: { color: C.blue }, rectRadius: 0.027 });
  box(s, 'roundRect', 5.556, 6.026, 0.169, 0.162, { fill: { color: C.blue }, rectRadius: 0.027 });
  t(s, '2015-2016', 7.088, 4.576, 1.289, 0.353, { fontSize: 15, bold: true });
  p(s, 'University Name', 7.056, 4.847, 1.389, 0.286, { fontSize: 11, color: C.blue, align: 'center', lineSpacingMultiple: 1, wrap: false });
  p(s, L.L3, 8.558, 4.39, 4.02, 0.831, { color: C.gray });
  t(s, '2016-2017', 7.088, 5.717, 1.27, 0.353, { fontSize: 15, bold: true });
  p(s, 'University Name', 7.056, 5.988, 1.389, 0.286, { fontSize: 11, color: C.blue, align: 'center', lineSpacingMultiple: 1, wrap: false });
  p(s, L.L3, 8.558, 5.532, 4.02, 0.831, { color: C.gray });
  hex(s, -1.315, 6.529, 3.596, 3.1, 351.2);
  t(s, 'University', 0.511, 2.072, 2.46, 0.606, { fontSize: 30, bold: true, color: C.white, lineSpacing: 36 });
  dots(s, 0.66, 1.823);
  p(s, L.L8, 0.511, 2.592, 3.378, 0.831, { color: C.line2 });
  p(s, L.L13, 5.472, 1.761, 2.764, 1.084, { color: C.gray });
  box(s, 'line', 5.627, 3.297, 0, 3.144, { line: { color: C.line2, width: 0.75 } });
  photo(s, 0.608, 3.673, 3.378, 1.432, 0.098);
  photo(s, 0.608, 5.313, 3.378, 1.432, 0.098);
  photo(s, 8.64, 1.544, 4.004, 1.432, 0.098);
  photo(s, 6.145, 3.377, 0.777, 0.743, 0.096);
  photo(s, 6.145, 4.57, 0.777, 0.743, 0.096);
  photo(s, 6.145, 5.657, 0.777, 0.743, 0.096);
}

// --- 10  award & certificate
function slide10(s) {
  box(s, 'rect', -0.038, 4.286, 13.372, 3.266, { fill: { color: C.dark } });
  chrome(s, NAV.NAV1, 10);
  p(s, L.L9, 1.069, 5.157, 3.652, 0.831, { color: 'A6A6A6' });
  t(s, 'Graphic Design', 6.75, 1.663, 1.922, 0.353, { fontSize: 15, bold: true, lineSpacing: 18 });
  p(s, L.L10, 6.729, 1.955, 2.107, 0.758, { fontSize: 9, color: C.gray });
  t(s, 'Web Design', 10.467, 1.663, 1.534, 0.353, { fontSize: 15, bold: true, lineSpacing: 18 });
  p(s, L.L10, 10.446, 1.955, 2.107, 0.758, { fontSize: 9, color: C.gray });
  hex(s, -1.836, 6.584, 3.596, 3.1, 351.2);
  t(s, [{ text: 'Award & ', options: { breakLine: true } }, { text: 'Certificate' }], 1.089, 1.881, 2.93, 1.279, { fontSize: 35, bold: true, lineSpacing: 42 });
  dots(s, 1.218, 1.572);
  p(s, L.L9, 1.089, 3.072, 4, 0.831);
  t(s, 'Award Description', 1.069, 4.841, 2.26, 0.353, { fontSize: 15, bold: true, color: C.white, lineSpacing: 18 });
  photo(s, 5.392, 3.211, 7.018, 3.16, 0.156);
  photo(s, 5.433, 1.703, 1.02, 0.976, 0.126);
  photo(s, 9.15, 1.703, 1.02, 0.976, 0.126);
}

// --- 11  let’s take a break
function slide11(s) {
  s.background = { color: C.dark };
  hexPair(s);
  hex(s, 11.964, -0.556, 4.974, 4.288, 351.2);
  chrome(s, NAV.NAV1, 11, 1);
  t(s, 'Let\u2019s Take A Break', 8.106, 2.729, 4.245, 0.606, { fontSize: 30, bold: true, color: C.white, lineSpacing: 36 });
  dots(s, 8.256, 2.414);
  p(s, L.L3, 8.099, 3.4, 4.245, 0.831, { color: C.line2 });
  box(s, 'roundRect', 0.485, 6.945, 0.17, 0.162, { fill: { color: C.blue }, rectRadius: 0.027 });
  box(s, 'roundRect', 0.974, 6.945, 0.17, 0.162, { fill: { color: C.blue }, rectRadius: 0.027 });
  box(s, 'roundRect', 1.464, 6.945, 0.17, 0.162, { fill: { color: C.blue }, rectRadius: 0.027 });
  box(s, 'roundRect', 8.22, 4.677, 4.333, 1.644, { line: { color: C.line2, width: 0.75 }, rectRadius: 0.098 });
  t(s, 'More Info About Me', 8.55, 4.936, 1.943, 0.303, { bold: true, color: C.white });
  p(s, L.L8, 8.55, 5.199, 3.652, 0.831, { color: C.line2 });
  box(s, 'custGeom', -0.048, 1.939, 7.234, 4.519, { fill: { color: C.photo }, objectName: '[image]', points: [{ x: 0.021, y: 0 }, { x: 0.36, y: 0 }, { x: 1.702, y: 0 }, { x: 6.891, y: 0 }, { x: 7.234, y: 0.343, curve: { type: 'cubic', x1: 7.081, y1: 0, x2: 7.234, y2: 0.153 } }, { x: 7.234, y: 4.176 }, { x: 6.891, y: 4.519, curve: { type: 'cubic', x1: 7.234, y1: 4.366, x2: 7.081, y2: 4.519 } }, { x: 1.702, y: 4.519 }, { x: 0.36, y: 4.519 }, { x: 0.021, y: 4.519 }, { x: 0, y: 4.517 }, { x: 0, y: 0.002 }, { close: true }] });
}

// --- 12  my services
function slide12(s) {
  t(s, 'My Services', 1.144, 2.645, 2.863, 0.606, { fontSize: 30, bold: true, lineSpacing: 36 });
  dots(s, 1.293, 2.397);
  p(s, L.L1, 1.144, 3.185, 4, 1.084);
  chrome(s, NAV.NAV4, 12);
  hex(s, -2.992, 6.405, 4.974, 4.288, 351.2);
  t(s, 'Graphic Design', 5.475, 4.619, 1.922, 0.353, { fontSize: 15, bold: true, align: 'center' });
  p(s, L.L2, 5.301, 4.972, 2.271, 0.831, { align: 'center' });
  t(s, 'Web Design', 8.048, 4.619, 1.534, 0.353, { fontSize: 15, bold: true, align: 'center' });
  p(s, L.L2, 7.68, 4.972, 2.271, 0.831, { align: 'center' });
  t(s, 'UI UX Design', 10.424, 4.619, 1.629, 0.353, { fontSize: 15, bold: true, align: 'center' });
  p(s, L.L2, 10.103, 4.972, 2.271, 0.831, { align: 'center' });
  photo(s, 1.232, 4.443, 3.519, 1.319, 0.207);
  photo(s, 5.475, 2.191, 2.025, 2.121, 0.139);
  photo(s, 7.869, 2.191, 2.025, 2.121, 0.139);
  photo(s, 10.263, 2.191, 2.025, 2.121, 0.139);
}

// --- 13  latest projects
function slide13(s) {
  box(s, 'rect', 6.16, 0.865, 7.173, 6.635, { fill: { color: C.dark } });
  t(s, 'Latest Projects', 7.328, 2.42, 3.459, 0.606, { fontSize: 30, bold: true, color: C.white, lineSpacing: 36 });
  dots(s, 7.477, 2.172);
  p(s, L.L1, 7.342, 3.349, 5.127, 0.831, { color: C.track });
  t(s, 'GRAPHIC DESIGNER', 7.342, 2.909, 2.956, 0.37, { fontSize: 16, color: C.white, charSpacing: 3 });
  t(s, 'Project One', 2.324, 2.281, 1.489, 0.353, { fontSize: 15, bold: true, lineSpacing: 18 });
  p(s, L.L2, 2.324, 2.56, 3.075, 0.579);
  t(s, 'Project Two', 2.324, 3.572, 1.489, 0.353, { fontSize: 15, bold: true, lineSpacing: 18 });
  p(s, L.L2, 2.324, 3.851, 3.075, 0.579);
  t(s, 'Project Three', 2.324, 4.862, 1.673, 0.353, { fontSize: 15, bold: true, lineSpacing: 18 });
  p(s, L.L2, 2.324, 5.141, 3.075, 0.579);
  hex(s, -2.992, 6.405, 4.974, 4.288, 351.2);
  chrome(s, NAV.NAV4, 13);
  photo(s, 0.961, 3.475, 1.153, 1.104, 0.143);
  photo(s, 0.961, 4.765, 1.153, 1.104, 0.143);
  photo(s, 0.961, 2.185, 1.153, 1.104, 0.143);
  photo(s, 7.342, 4.429, 4.841, 1.279, 0.088);
}

// --- 14  client review
function slide14(s) {
  box(s, 'rect', -0.038, 4.02, 5.924, 3.531, { fill: { color: C.dark } });
  hex(s, -2.992, 6.405, 4.974, 4.288, 351.2);
  p(s, L.L3, 7.658, 5.392, 4.566, 0.831, { color: C.gray });
  t(s, 'Client Review', 7.658, 2.3, 3.231, 0.606, { fontSize: 30, bold: true, lineSpacing: 36 });
  dots(s, 7.808, 2.052);
  t(s, 'John Doe', 7.658, 5.073, 1.04, 0.303, { bold: true });
  t(s, 'Jessica Doe', 9.274, 5.077, 1.261, 0.303, { bold: true });
  t(s, 'Riana Doe', 10.89, 5.056, 1.114, 0.303, { bold: true });
  p(s, L.L5, 7.658, 2.847, 4.566, 0.579, { color: C.gray });
  chrome(s, NAV.NAV4, 14);
  photo(s, 1.432, 2.029, 5.572, 4.211, 0.208);
  photo(s, 9.315, 3.645, 1.146, 1.201, 0.144);
  photo(s, 10.903, 3.645, 1.146, 1.201, 0.144);
  photo(s, 7.726, 3.645, 1.146, 1.201, 0.144);
}

// --- 15  my services (cards)
function slide15(s) {
  box(s, 'rect', 0, 3.75, 13.333, 3.75, { fill: { color: C.dark } });
  t(s, 'Graphic Design', 2.285, 5.298, 1.922, 0.353, { fontSize: 15, bold: true, color: C.white, align: 'center' });
  p(s, L.L2, 2.111, 5.651, 2.271, 0.831, { color: C.mist, align: 'center' });
  t(s, 'Web Design', 5.982, 5.298, 1.534, 0.353, { fontSize: 15, bold: true, color: C.white, align: 'center' });
  p(s, L.L2, 5.614, 5.651, 2.271, 0.831, { color: C.mist, align: 'center' });
  t(s, 'UI/UX Design', 9.321, 5.293, 1.661, 0.353, { fontSize: 15, bold: true, color: C.white, align: 'center' });
  p(s, L.L2, 9.016, 5.646, 2.271, 0.831, { color: C.mist, align: 'center' });
  hex(s, -2.992, 6.405, 4.974, 4.288, 351.2);
  t(s, 'My Services', 1.999, 1.76, 2.863, 0.606, { fontSize: 30, bold: true, lineSpacing: 36 });
  dots(s, 2.149, 1.512);
  p(s, L.L6, 5.08, 1.675, 4.134, 0.579, { color: C.gray });
  box(s, 'roundRect', 9.384, 1.697, 1.903, 0.545, { fill: { color: C.blue }, rectRadius: 0.09 });
  t(s, 'Learn More', 9.705, 1.833, 1.309, 0.303, { bold: true, color: C.white, align: 'center', lineSpacing: 14.4, wrap: true });
  chrome(s, NAV.NAV4, 15);
  photo(s, 1.985, 2.662, 2.436, 2.236, 0.281);
  photo(s, 5.435, 2.662, 2.436, 2.236, 0.281);
  photo(s, 8.884, 2.662, 2.436, 2.236, 0.281);
}

// --- 16  let’s take a break
function slide16(s) {
  s.background = { color: C.dark };
  hexPair(s);
  chrome(s, NAV.NAV4, 16);
  t(s, 'Let\u2019s Take A Break', 7.527, 2.851, 4.245, 0.606, { fontSize: 30, bold: true, color: C.white });
  dots(s, 7.677, 2.537);
  p(s, L.L1, 7.483, 3.522, 4.245, 1.084, { color: C.line2 });
  box(s, 'roundRect', 0.491, 6.056, 0.137, 0.143, { fill: { color: C.blue }, rectRadius: 0.023 });
  box(s, 'roundRect', 0.491, 6.468, 0.137, 0.143, { fill: { color: C.blue }, rectRadius: 0.023 });
  box(s, 'roundRect', 0.491, 6.882, 0.137, 0.143, { fill: { color: C.blue }, rectRadius: 0.023 });
  t(s, '200+', 10.545, 5.073, 1.052, 0.505, { fontSize: 24, bold: true, color: C.blue });
  t(s, 'Project', 10.56, 5.466, 0.979, 0.353, { fontSize: 15, bold: true, color: C.white });
  box(s, 'custGeom', 1.163, 2.051, 5.572, 5.449, { fill: { color: C.photo }, objectName: '[image]', points: [{ x: 0.275, y: 0 }, { x: 5.297, y: 0 }, { x: 5.572, y: 0.275, curve: { type: 'cubic', x1: 5.449, y1: 0, x2: 5.572, y2: 0.123 } }, { x: 5.572, y: 5.449 }, { x: 0, y: 5.449 }, { x: 0, y: 0.275 }, { x: 0.275, y: 0, curve: { type: 'cubic', x1: 0, y1: 0.123, x2: 0.123, y2: 0 } }, { close: true }] });
  photo(s, 7.635, 5.031, 1.192, 0.824, 0.103);
  photo(s, 9.085, 5.031, 1.192, 0.824, 0.103);
}

// --- 17  skills overview - web
function slide17(s) {
  box(s, 'roundRect', 7.371, 4.2, 4.132, 0.328, { fill: { color: C.track }, rectRadius: 0.164 });
  box(s, 'rect', 0, 0.824, 4.037, 6.676, { fill: { color: C.dark } });
  hexPair(s);
  chrome(s, NAV.NAV3, 17);
  t(s, 'Skills Overview', 7.292, 2.351, 3.582, 0.606, { fontSize: 30, bold: true, lineSpacing: 36 });
  dots(s, 7.441, 2.103);
  p(s, L.L5, 7.292, 3.239, 4.566, 0.579, { color: C.gray });
  hex(s, -2.992, 6.405, 4.974, 4.288, 351.2);
  t(s, 'Web Design Skills', 7.3, 2.841, 2.006, 0.353, { fontSize: 15, color: C.gray });
  box(s, 'roundRect', 7.371, 4.2, 3.129, 0.309, { fill: { color: C.blue }, rectRadius: 0.154 });
  t(s, '70%', 11.633, 4.194, 0.638, 0.353, { fontSize: 15, bold: true, color: C.gray });
  t(s, 'Drawing Skills', 7.453, 4.24, 1.264, 0.269, { fontSize: 10, bold: true, color: C.white, lineSpacing: 12 });
  box(s, 'roundRect', 7.384, 4.788, 4.132, 0.328, { fill: { color: C.track }, rectRadius: 0.164 });
  box(s, 'roundRect', 7.384, 4.788, 3.789, 0.309, { fill: { color: C.blue }, rectRadius: 0.154 });
  t(s, '90%', 11.645, 4.781, 0.647, 0.353, { fontSize: 15, bold: true, color: C.gray });
  t(s, 'Layout Design Skills', 7.465, 4.827, 1.696, 0.269, { fontSize: 10, bold: true, color: C.white, lineSpacing: 12 });
  box(s, 'roundRect', 7.371, 5.394, 4.132, 0.328, { fill: { color: C.track }, rectRadius: 0.164 });
  box(s, 'roundRect', 7.371, 5.394, 3.53, 0.328, { fill: { color: C.blue }, rectRadius: 0.164 });
  t(s, '80%', 11.633, 5.387, 0.647, 0.353, { fontSize: 15, bold: true, color: C.gray });
  t(s, 'Web Design Skills', 7.453, 5.433, 1.529, 0.269, { fontSize: 10, bold: true, color: C.white, lineSpacing: 12 });
  photo(s, 1.435, 1.887, 5.22, 4.4, 0.405);
}

// --- 18  skills overview - graphic
function slide18(s) {
  box(s, 'rect', 0, 0.824, 8.22, 6.676, { fill: { color: C.dark } });
  hexPair(s);
  chrome(s, NAV.NAV3, 18);
  hex(s, -2.992, 6.405, 4.974, 4.288, 351.2);
  t(s, 'Skills Overview', 1.177, 2.344, 3.582, 0.606, { fontSize: 30, bold: true, color: C.white, lineSpacing: 36 });
  dots(s, 1.327, 2.096);
  p(s, L.L1, 1.177, 3.232, 4.923, 0.831, { color: C.line2 });
  t(s, 'Graphic Design Skills', 1.185, 2.835, 2.36, 0.353, { fontSize: 15, color: C.line2 });
  box(s, 'roundRect', 3.103, 4.598, 0.253, 0.253, { fill: { color: C.blue }, rectRadius: 0.047 });
  box(s, 'roundRect', 3.527, 4.598, 0.253, 0.253, { fill: { color: C.blue }, rectRadius: 0.047 });
  box(s, 'roundRect', 3.948, 4.598, 0.253, 0.253, { fill: { color: C.blue }, rectRadius: 0.047 });
  box(s, 'roundRect', 4.371, 4.598, 0.253, 0.253, { fill: { color: C.track }, rectRadius: 0.047 });
  box(s, 'roundRect', 4.795, 4.598, 0.253, 0.253, { fill: { color: C.track }, rectRadius: 0.047 });
  box(s, 'roundRect', 3.103, 5.182, 0.253, 0.253, { fill: { color: C.blue }, rectRadius: 0.047 });
  box(s, 'roundRect', 3.527, 5.182, 0.253, 0.253, { fill: { color: C.blue }, rectRadius: 0.047 });
  box(s, 'roundRect', 3.948, 5.182, 0.253, 0.253, { fill: { color: C.blue }, rectRadius: 0.047 });
  box(s, 'roundRect', 4.371, 5.182, 0.253, 0.253, { fill: { color: C.blue }, rectRadius: 0.047 });
  box(s, 'roundRect', 4.795, 5.182, 0.253, 0.253, { fill: { color: C.track }, rectRadius: 0.047 });
  box(s, 'roundRect', 3.103, 5.767, 0.253, 0.253, { fill: { color: C.blue }, rectRadius: 0.047 });
  box(s, 'roundRect', 3.527, 5.767, 0.253, 0.253, { fill: { color: C.blue }, rectRadius: 0.047 });
  box(s, 'roundRect', 3.948, 5.767, 0.253, 0.253, { fill: { color: C.blue }, rectRadius: 0.047 });
  box(s, 'roundRect', 4.371, 5.767, 0.253, 0.253, { fill: { color: C.blue }, rectRadius: 0.047 });
  box(s, 'roundRect', 4.795, 5.767, 0.253, 0.253, { fill: { color: C.blue }, rectRadius: 0.047 });
  t(s, 'Experience', 1.262, 4.598, 1.417, 0.353, { fontSize: 15, bold: true, color: C.white });
  t(s, 'Leadership', 1.262, 5.197, 1.433, 0.353, { fontSize: 15, bold: true, color: C.white });
  t(s, 'Hard Work', 1.28, 5.794, 1.392, 0.353, { fontSize: 15, bold: true, color: C.white });
  t(s, '3/5', 5.314, 4.528, 0.609, 0.387, { fontSize: 17, bold: true, color: C.bar, align: 'center' });
  t(s, '4/5', 5.317, 5.138, 0.626, 0.387, { fontSize: 17, bold: true, color: C.bar, align: 'right' });
  t(s, '5/5', 5.334, 5.739, 0.626, 0.387, { fontSize: 17, bold: true, color: C.bar, align: 'right' });
  photo(s, 7.109, 1.588, 5.43, 4.633, 0.33);
}

// --- 19  language fluency
function slide19(s) {
  hexPair(s);
  chrome(s, NAV.NAV3, 19);
  t(s, 'Language Fluency', 7.79, 2.494, 3.903, 0.555, { fontSize: 27, bold: true, lineSpacing: 32.4 });
  dots(s, 7.94, 2.245);
  p(s, L.L14, 7.79, 3.309, 4.076, 0.579, { color: C.gray });
  t(s, 'My Language Fluency', 7.798, 2.984, 2.456, 0.353, { fontSize: 15, color: C.gray });
  box(s, 'roundRect', 7.867, 4.162, 1.932, 0.931, { line: { color: C.line2, width: 0.75 }, rectRadius: 0.107 });
  box(s, 'roundRect', 9.941, 4.162, 1.932, 0.931, { line: { color: C.line2, width: 0.75 }, rectRadius: 0.107 });
  box(s, 'roundRect', 7.86, 5.234, 1.932, 0.931, { line: { color: C.line2, width: 0.75 }, rectRadius: 0.107 });
  box(s, 'roundRect', 9.935, 5.234, 1.932, 0.931, { line: { color: C.line2, width: 0.75 }, rectRadius: 0.107 });
  t(s, 'English', 8.282, 4.349, 1.07, 0.37, { fontSize: 16, bold: true, align: 'center', lineSpacing: 19.2 });
  t(s, 'Fluent', 8.415, 4.669, 0.823, 0.353, { fontSize: 15, color: C.blue, align: 'center', lineSpacing: 18 });
  t(s, 'Spanish', 10.351, 4.349, 1.171, 0.37, { fontSize: 16, bold: true, align: 'center', lineSpacing: 19.2 });
  t(s, 'Fluent', 10.535, 4.669, 0.823, 0.353, { fontSize: 15, color: C.blue, align: 'center', lineSpacing: 18 });
  t(s, 'French', 8.313, 5.429, 1.019, 0.37, { fontSize: 16, bold: true, align: 'center', lineSpacing: 19.2 });
  t(s, 'Proficient', 8.285, 5.749, 1.166, 0.353, { fontSize: 15, color: C.blue, align: 'center', lineSpacing: 18 });
  t(s, 'German', 10.358, 5.429, 1.168, 0.37, { fontSize: 16, bold: true, align: 'center', lineSpacing: 19.2 });
  t(s, 'Intermediate', 10.181, 5.749, 1.541, 0.353, { fontSize: 15, color: C.blue, align: 'center', lineSpacing: 18 });
  hex(s, -2.992, 6.405, 4.974, 4.288, 351.2);
  photo(s, 1.147, 1.846, 2.829, 4.633, 0.201);
  photo(s, 4.364, 1.846, 2.829, 4.633, 0.201);
}

// --- 20  award & recognition
function slide20(s) {
  box(s, 'rect', 0, 0.824, 6.969, 6.676, { fill: { color: C.dark } });
  chrome(s, NAV.NAV3, 20);
  t(s, 'Award & Recognition', 1.144, 1.95, 4.86, 0.606, { fontSize: 30, bold: true, color: C.white, lineSpacing: 36 });
  dots(s, 1.293, 1.701);
  p(s, L.L1, 1.144, 2.496, 4.981, 0.831, { color: C.line2 });
  t(s, 'Award One', 9.532, 1.785, 1.429, 0.353, { fontSize: 15, bold: true, lineSpacing: 18 });
  p(s, L.L2, 9.532, 1.997, 3.075, 0.579);
  box(s, 'line', 7.768, 3.035, 4.838, 0, { line: { color: C.track, width: 0.75 } });
  t(s, 'Award Two', 9.532, 3.43, 1.429, 0.353, { fontSize: 15, bold: true, lineSpacing: 18 });
  p(s, L.L2, 9.532, 3.642, 3.075, 0.579);
  box(s, 'line', 7.768, 4.68, 4.81, 0, { line: { color: C.track, width: 0.75 } });
  t(s, 'Award Three', 9.531, 5.075, 1.613, 0.353, { fontSize: 15, bold: true, lineSpacing: 18 });
  p(s, L.L2, 9.531, 5.287, 3.075, 0.579);
  box(s, 'line', 7.767, 6.325, 4.811, 0, { line: { color: C.track, width: 0.75 } });
  hex(s, -2.992, 6.405, 4.974, 4.288, 351.2);
  photo(s, 1.141, 3.79, 4.981, 2.649, 0.188);
  photo(s, 7.765, 3.35, 1.487, 1.038, 0.134);
  photo(s, 7.765, 4.977, 1.487, 1.038, 0.134);
  photo(s, 7.767, 1.705, 1.487, 1.038, 0.134);
}

// --- 21  let’s take a break
function slide21(s) {
  s.background = { color: C.dark };
  box(s, 'line', 7.112, 3.479, 4.838, 0, { line: { color: C.mid, width: 0.75 } });
  hexPair(s);
  chrome(s, NAV.NAV3, 21);
  hex(s, -2.992, 6.405, 4.974, 4.288, 351.2);
  t(s, 'Let\u2019s Take A Break', 7.82, 4.237, 4.245, 0.606, { fontSize: 30, bold: true, color: C.white });
  dots(s, 7.969, 3.922);
  p(s, L.L1, 7.775, 4.907, 4.245, 1.084, { color: C.line2 });
  photo(s, 1.174, 1.816, 6.048, 4.541, 0.323);
}

// --- 22  project history
function slide22(s) {
  hexPair(s);
  chrome(s, NAV.NAV2, 22);
  box(s, 'line', 8.835, 4.066, 3.357, 0, { line: { color: C.track, width: 0.75 } });
  t(s, 'Project History', 7.09, 2.12, 3.468, 0.606, { fontSize: 30, bold: true });
  dots(s, 7.187, 1.815);
  p(s, L.L1, 7.09, 2.715, 5.407, 0.831);
  box(s, 'can', 7.245, 4.801, 1.072, 1.268, { fill: { color: '3E4043' } });
  box(s, 'can', 8.537, 4.414, 1.072, 1.656, { fill: { color: C.blue } });
  box(s, 'can', 9.829, 4.801, 1.072, 1.268, { fill: { color: '757679' } });
  box(s, 'can', 11.12, 5.06, 1.072, 1.01, { fill: { color: '73767B' } });
  t(s, 'Project Finished', 7.103, 3.966, 1.632, 0.303, { bold: true });
  t(s, [{ text: '40%', options: { bold: true, breakLine: true } }, { text: 'Graphic', options: { breakLine: true } }, { text: 'Dsign' }], 8.706, 5.288, 0.756, 0.606, { fontSize: 10, color: C.white, align: 'center' });
  t(s, [{ text: '20%', options: { bold: true, breakLine: true } }, { text: 'Web Dsign' }], 7.309, 5.456, 0.949, 0.438, { fontSize: 10, color: C.white, align: 'center' });
  t(s, [{ text: '20%', options: { bold: true, breakLine: true } }, { text: 'Illustration' }], 9.912, 5.456, 0.917, 0.438, { fontSize: 10, color: C.white, align: 'center' });
  t(s, [{ text: '10%', options: { bold: true, breakLine: true } }, { text: 'Mobile UI' }], 11.241, 5.456, 0.83, 0.438, { fontSize: 10, color: C.white, align: 'center' });
  hex(s, -2.992, 6.405, 4.974, 4.288, 351.2);
  photo(s, 1.142, 1.718, 3.377, 4.717, 0.17);
  photo(s, 4.753, 1.718, 1.653, 2.606, 0.213);
  photo(s, 4.738, 4.495, 1.653, 1.941, 0.213);
}

// --- 23  portfolio
function slide23(s) {
  box(s, 'rect', 0, 0.824, 8.394, 6.676, { fill: { color: C.dark } });
  box(s, 'line', 1.144, 3.474, 3.357, 0, { line: { color: C.mid, width: 0.75 } });
  chrome(s, NAV.NAV2, 23);
  hex(s, -2.992, 6.405, 4.974, 4.288, 351.2);
  t(s, 'Portfolio', 1.028, 2.318, 2.062, 0.606, { fontSize: 30, bold: true, color: C.white });
  dots(s, 1.178, 2.07);
  p(s, L.L1, 1.036, 3.952, 4.923, 0.831, { color: C.line2 });
  t(s, 'Graphic Design Work', 1.036, 2.808, 2.392, 0.353, { fontSize: 15, color: C.line2 });
  t(s, '200+', 1.435, 5.326, 1.052, 0.505, { fontSize: 24, bold: true, color: C.blue, align: 'center' });
  t(s, 'Project', 1.503, 5.718, 0.914, 0.353, { fontSize: 15, color: C.white });
  t(s, '100+', 2.628, 5.325, 1.001, 0.505, { fontSize: 24, bold: true, color: C.blue, align: 'center' });
  t(s, 'On Going', 2.532, 5.718, 1.18, 0.353, { fontSize: 15, color: C.white, align: 'center' });
  box(s, 'roundRect', 1.144, 5.092, 5.081, 1.196, { line: { color: C.line2, width: 0.75 }, rectRadius: 0.138 });
  p(s, L.L7, 3.84, 5.311, 2.113, 0.758, { fontSize: 9, color: C.line2 });
  photo(s, 10.262, 3.864, 2.178, 2.465, 0.11);
  photo(s, 10.214, 1.612, 2.178, 2.127, 0.107);
  photo(s, 6.714, 1.612, 3.377, 4.717, 0.17);
  photo(s, 4.365, 1.612, 2.178, 2.127, 0.107);
}

// --- 24  my portfolio
function slide24(s) {
  hex(s, -2.992, 6.405, 4.974, 4.288, 351.2);
  chrome(s, NAV.NAV2, 24);
  t(s, 'My Portfolio', 1.326, 1.976, 2.856, 0.606, { fontSize: 30, bold: true, lineSpacing: 36 });
  dots(s, 1.423, 1.671);
  p(s, L.L11, 4.984, 1.808, 6.889, 0.579, { color: C.gray });
  t(s, 'Web Design', 1.326, 5.533, 1.534, 0.353, { fontSize: 15, bold: true, lineSpacing: 18 });
  p(s, L.L2, 1.326, 5.783, 3.075, 0.579);
  t(s, 'UI UX Design', 4.984, 5.533, 1.629, 0.353, { fontSize: 15, bold: true, lineSpacing: 18 });
  p(s, L.L2, 4.984, 5.783, 3.075, 0.579);
  t(s, 'Graphic Design', 8.669, 5.533, 1.922, 0.353, { fontSize: 15, bold: true, lineSpacing: 18 });
  p(s, L.L2, 8.669, 5.783, 3.075, 0.579);
  photo(s, 1.326, 2.933, 3.338, 2.417, 0.122);
  photo(s, 4.998, 2.933, 3.338, 2.417, 0.122);
  photo(s, 8.669, 2.933, 3.338, 2.417, 0.122);
}

// --- 25  work in progress
function slide25(s) {
  photo(s, 4.145, 1.74, 2.907, 4.717, 0.146);
  photo(s, 0.999, 1.74, 2.907, 4.717, 0.146);
  hexPair(s);
  chrome(s, NAV.NAV2, 25);
  box(s, 'custGeom', 8.431, 4.424, 4.496, 1.526, { fill: { color: C.stone }, line: { color: C.white, width: 6 }, points: [{ x: 0, y: 0.764 }, { x: 0, y: 0.764 }, { x: 0, y: 0.764 }, { close: true }, { x: 3.733, y: 0 }, { x: 4.496, y: 0.763, curve: { type: 'cubic', x1: 4.154, y1: 0, x2: 4.496, y2: 0.342 } }, { x: 3.733, y: 1.526, curve: { type: 'cubic', x1: 4.496, y1: 1.184, x2: 4.154, y2: 1.526 } }, { x: 3.676, y: 1.523 }, { x: 3.62, y: 1.526 }, { x: 0.762, y: 1.526 }, { x: 0.015, y: 0.917, curve: { type: 'cubic', x1: 0.394, y1: 1.526, x2: 0.086, y2: 1.265 } }, { x: 0, y: 0.764 }, { x: 0.015, y: 0.61 }, { x: 0.762, y: 0.002, curve: { type: 'cubic', x1: 0.086, y1: 0.263, x2: 0.394, y2: 0.002 } }, { x: 3.62, y: 0.002 }, { x: 3.661, y: 0.004 }, { close: true }] });
  box(s, 'custGeom', 5.771, 4.387, 3.735, 1.526, { fill: { color: C.slate2 }, line: { color: C.white, width: 6 }, points: [{ x: 0, y: 0.764 }, { x: 0, y: 0.764 }, { x: 0, y: 0.764 }, { close: true }, { x: 2.972, y: 0 }, { x: 3.735, y: 0.763, curve: { type: 'cubic', x1: 3.394, y1: 0, x2: 3.735, y2: 0.342 } }, { x: 2.972, y: 1.526, curve: { type: 'cubic', x1: 3.735, y1: 1.184, x2: 3.394, y2: 1.526 } }, { x: 2.916, y: 1.523 }, { x: 2.859, y: 1.526 }, { x: 0.762, y: 1.526 }, { x: 0.015, y: 0.917, curve: { type: 'cubic', x1: 0.394, y1: 1.526, x2: 0.087, y2: 1.265 } }, { x: 0, y: 0.764 }, { x: 0.015, y: 0.61 }, { x: 0.762, y: 0.002, curve: { type: 'cubic', x1: 0.087, y1: 0.263, x2: 0.394, y2: 0.002 } }, { x: 2.859, y: 0.002 }, { x: 2.9, y: 0.004 }, { close: true }] });
  t(s, 'Web Design Project', 6.227, 4.783, 2.205, 0.325, { fontSize: 13, bold: true, color: C.white, lineSpacing: 15.6, wrap: true });
  p(s, L.L12, 6.25, 5.03, 2.802, 0.531, { fontSize: 9, color: C.white });
  t(s, 'Graphic Design Project', 9.891, 4.785, 2.661, 0.325, { fontSize: 13, bold: true, color: C.white, lineSpacing: 15.6, wrap: true });
  p(s, L.L12, 9.915, 5.032, 2.802, 0.531, { fontSize: 9, color: C.white });
  hex(s, -2.992, 6.405, 4.974, 4.288, 351.2);
  t(s, 'Work In Progress', 7.577, 2.672, 3.982, 0.606, { fontSize: 30, bold: true, lineSpacing: 36 });
  dots(s, 7.674, 2.367);
  p(s, L.L1, 7.577, 3.267, 5.099, 0.831);
}

// --- 26  project statistics
function slide26(s) {
  box(s, 'rect', 0, 0.824, 5.843, 6.676, { fill: { color: C.dark } });
  box(s, 'rect', 4.178, 1.811, 1.666, 2.311, { fill: { color: C.slate } });
  box(s, 'roundRect', 5.843, 4.111, 6.258, 2.404, { line: { color: C.line2, width: 0.75 }, rectRadius: 0.278 });
  hexPair(s);
  chrome(s, NAV.NAV2, 26);
  t(s, 'Project Statistics', 6.454, 2.199, 3.961, 0.606, { fontSize: 30, bold: true });
  dots(s, 6.551, 1.894);
  p(s, L.L1, 6.454, 2.794, 5.581, 0.831);
  t(s, '200+', 4.496, 2.07, 1.052, 0.505, { fontSize: 24, bold: true, color: C.blue, align: 'center' });
  t(s, 'Project', 4.528, 2.462, 0.914, 0.353, { fontSize: 15, color: C.white, align: 'center' });
  t(s, '100+', 4.522, 2.946, 1.001, 0.505, { fontSize: 24, bold: true, color: C.blue, align: 'center' });
  t(s, 'Finished', 4.492, 3.338, 1.049, 0.353, { fontSize: 15, color: C.white, align: 'center' });
  hex(s, -2.992, 6.405, 4.974, 4.288, 351.2);
  // clustered column chart: 5 one-point series, hidden category axis
  s.addChart('bar', BAR_SERIES, {
    x: 9.496, y: 4.454, w: 2.246, h: 1.706,
    barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 219, barOverlapPct: -27,
    chartColors: [C.blue, C.slate2, C.silver, C.sage, C.coal],
    showLegend: false, showTitle: false, catAxisHidden: true,
    valAxisLineShow: false, valAxisLabelFontSize: 10, valAxisLabelColor: C.mid,
    valAxisLabelFontFace: F.body, valGridLine: { color: C.line2, size: 0.75 },
    plotArea: { fill: { color: C.white } },
  });
  p(s, L.L5, 7.094, 4.93, 2.275, 0.986, { fontSize: 9 });
  t(s, 'Penelope Doe', 7.115, 4.658, 1.422, 0.303, { bold: true });
  photo(s, 0.999, 1.811, 2.907, 4.717, 0.146);
  photo(s, 4.191, 4.111, 2.488, 2.417, 0.122);
}

// --- 27  on going projects
function slide27(s) {
  box(s, 'roundRect', 4.918, 4.509, 4.835, 1.852, { line: { color: C.line2, width: 0.75 }, rectRadius: 0.217 });
  hexPair(s);
  chrome(s, NAV.NAV2, 27);
  t(s, 'On Going Projects', 4.918, 2.541, 4.185, 0.606, { fontSize: 30, bold: true, lineSpacing: 36 });
  dots(s, 5.015, 2.236);
  p(s, L.L1, 4.918, 3.136, 5.581, 0.831);
  t(s, '100+', 10.894, 3.125, 1.001, 0.505, { fontSize: 24, bold: true, color: C.blue, align: 'center' });
  t(s, 'On Going', 10.797, 3.519, 1.18, 0.353, { fontSize: 15, color: C.dark, align: 'center' });
  t(s, 'About The Client', 5.31, 4.974, 2.072, 0.353, { fontSize: 15, bold: true, lineSpacing: 18 });
  p(s, L.L2, 5.31, 5.224, 2.406, 0.831);
  photo(s, 1.144, 1.738, 3.247, 4.717, 0.163);
  photo(s, 8.204, 4.508, 1.928, 1.883, 0.095);
  photo(s, 10.306, 4.493, 1.928, 1.883, 0.095);
}

// --- 28  client reviews
function slide28(s) {
  box(s, 'roundRect', 1.495, 4.92, 5.13, 1.731, { line: { color: C.line2, width: 0.75 }, rectRadius: 0.203 });
  photo(s, 1.488, 3.176, 5.154, 2.468, 0.124);
  photo(s, 10.479, 3.418, 1.164, 1.151, 0.219);
  photo(s, 10.479, 5.257, 1.164, 1.151, 0.219);
  p(s, L.L5, 1.718, 5.82, 4.51, 0.579);
  chrome(s, NAV.NAV2, 28);
  t(s, 'Client Reviews', 1.488, 2.182, 3.464, 0.606, { fontSize: 30, bold: true, lineSpacing: 36 });
  dots(s, 1.585, 1.877);
  p(s, L.L11, 5.302, 1.856, 5.48, 0.831);
  t(s, '200+', 10.863, 1.941, 1.052, 0.505, { fontSize: 24, bold: true, color: C.blue });
  t(s, 'Project', 10.878, 2.334, 0.979, 0.353, { fontSize: 15, bold: true, color: C.dark });
  box(s, 'roundRect', 7.064, 3.176, 4.835, 1.637, { line: { color: C.line2, width: 0.75 }, rectRadius: 0.192 });
  box(s, 'roundRect', 7.064, 5.015, 4.835, 1.637, { line: { color: C.line2, width: 0.75 }, rectRadius: 0.192 });
  p(s, L.L7, 7.442, 3.892, 2.959, 0.531, { fontSize: 9 });
  t(s, 'Jonathan Doe', 7.453, 3.609, 1.776, 0.353, { fontSize: 15, bold: true, color: C.dark, lineSpacing: 18 });
  p(s, L.L7, 7.453, 5.767, 2.959, 0.531, { fontSize: 9 });
  t(s, 'Penelope Doe', 7.463, 5.484, 1.725, 0.353, { fontSize: 15, bold: true, color: C.dark, lineSpacing: 18 });
  box(s, 'roundRect', 2.65, 3.977, 2.819, 0.866, { fill: { color: C.white, transparency: 31 }, rectRadius: 0.142 });
  t(s, 'Project Name', 3.203, 4.271, 1.713, 0.353, { fontSize: 15, bold: true, color: C.dark, align: 'center' });
}

// --- 29  get in touch
function slide29(s) {
  hexPair(s);
  chrome(s, NAV.NAV3, 29);
  t(s, [{ text: 'Get In Touch', options: { breakLine: true } }, { text: 'With Me' }], 1.52, 2.319, 3.019, 1.111, { fontSize: 30, bold: true, lineSpacing: 36 });
  dots(s, 1.67, 2.07);
  p(s, L.L1, 4.763, 2.32, 5.618, 0.831);
  box(s, 'roundRect', 1.588, 3.677, 4.939, 1.196, { line: { color: C.line2, width: 0.75 }, rectRadius: 0.138 });
  box(s, 'roundRect', 1.588, 5.119, 4.939, 1.196, { line: { color: C.line2, width: 0.75 }, rectRadius: 0.138 });
  box(s, 'roundRect', 6.753, 3.679, 4.939, 1.196, { line: { color: C.line2, width: 0.75 }, rectRadius: 0.138 });
  box(s, 'roundRect', 6.753, 5.122, 4.939, 1.196, { line: { color: C.line2, width: 0.75 }, rectRadius: 0.138 });
  t(s, 'My Email', 1.885, 4.122, 1.161, 0.337, { fontSize: 14, bold: true, lineSpacing: 16.8 });
  t(s, 'Office Address', 1.89, 5.598, 1.692, 0.337, { fontSize: 14, bold: true, lineSpacing: 16.8 });
  t(s, 'Phone Number', 7.066, 4.134, 1.757, 0.337, { fontSize: 14, bold: true, lineSpacing: 16.8 });
  t(s, 'Social Media', 7.03, 5.582, 1.543, 0.337, { fontSize: 14, bold: true, lineSpacing: 16.8 });
  p(s, 'Mail (at) website com', 3.129, 4.133, 1.694, 0.269, { lineSpacing: 12 });
  p(s, 'Mail (at) website com', 4.646, 4.133, 1.694, 0.269, { lineSpacing: 12 });
  p(s, '123 456 789 0', 9.048, 4.143, 1.694, 0.269, { lineSpacing: 12 });
  p(s, '123 456 789 0', 10.222, 4.15, 1.262, 0.269, { lineSpacing: 12 });
  p(s, L.L15, 3.634, 5.409, 2.643, 0.579);
  box(s, 'roundRect', 8.898, 5.514, 0.475, 0.455, { fill: { color: C.track }, rectRadius: 0.086 });
  t(s, 'A', 8.982, 5.596, 0.321, 0.303, { bold: true, align: 'center' });
  box(s, 'roundRect', 9.528, 5.514, 0.475, 0.455, { fill: { color: C.track }, rectRadius: 0.086 });
  t(s, 'B', 9.612, 5.596, 0.321, 0.303, { bold: true, align: 'center' });
  box(s, 'roundRect', 10.158, 5.514, 0.475, 0.455, { fill: { color: C.track }, rectRadius: 0.086 });
  t(s, 'C', 10.237, 5.596, 0.33, 0.303, { bold: true, align: 'center' });
  box(s, 'roundRect', 10.786, 5.514, 0.475, 0.455, { fill: { color: C.track }, rectRadius: 0.086 });
  t(s, 'D', 10.867, 5.596, 0.326, 0.303, { bold: true, align: 'center' });
  hex(s, -2.992, 6.405, 4.974, 4.288, 351.2);
}

// --- 30  thank you
function slide30(s) {
  s.background = { color: C.dark };
  photo(s, 0, 0, 13.333, 7.5);
  box(s, 'rect', 0.018, 0, 13.333, 7.5, { fill: { color: C.dark, transparency: 25 } });
  chrome(s, NAV.NAV1, 30, 1);
  box(s, 'roundRect', 0.485, 6.945, 0.17, 0.162, { fill: { color: C.blue }, rectRadius: 0.027 });
  box(s, 'roundRect', 0.974, 6.945, 0.17, 0.162, { fill: { color: C.blue }, rectRadius: 0.027 });
  box(s, 'roundRect', 1.464, 6.945, 0.17, 0.162, { fill: { color: C.blue }, rectRadius: 0.027 });
  t(s, 'Thank You For Watching', 1.188, 3.195, 5.174, 0.555, { fontSize: 27, bold: true, color: C.white });
  dots(s, 1.338, 2.88);
  p(s, L.L1, 1.144, 3.814, 5.174, 0.831, { color: C.line2 });
  box(s, 'roundRect', 1.258, 5.046, 3.05, 0.69, { fill: { color: C.blue }, rectRadius: 0.113 });
  t(s, 'Visit Our Website', 1.642, 5.245, 2.253, 0.37, { fontSize: 16, bold: true, color: C.dark, align: 'center' });
}

const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30];

const pres = new PptxGenJS();
pres.defineLayout({ name: 'W16x9', width: 13.333, height: 7.5 });
pres.layout = 'W16x9';
pres.author = 'Penelope Doe';
pres.title = 'Personal Resume';
SLIDES.forEach(build => build(pres.addSlide()));
pres.writeFile({ fileName: path.join(__dirname, '0fd4d301-e53b-4b2e-84f9-5d418807f7ce_grok_final.pptx') })
  .then(f => console.log('wrote ' + f));
