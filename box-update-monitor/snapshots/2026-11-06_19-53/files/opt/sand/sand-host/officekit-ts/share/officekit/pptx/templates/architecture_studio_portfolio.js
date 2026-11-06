/**
 * "Archinea Studio" presentation template - 32 slides, 13.335 x 7.5 in (16:9).
 *
 * Rebuilt from scratch with pptxgenjs. Every position is in inches and matches
 * the source deck. Photographs in the original are stand-ins here: each one is
 * drawn as a light-gray rectangle labelled [image] at the photo's own frame.
 *
 * Run: node 15180b4e-e785-4af2-bdb0-5d91dc838c4f_grok_final.js
 */
const path = require('path');
const pptxgen = require('pptxgenjs');

// ---------------------------------------------------------------- palette --
const INK = '262626';      // tx1 @ 85% luminance - headings, display numerals
const MID = '595959';      // tx1 @ 65% - occasional body copy
const GRAY = '808080';     // tx1 @ 50% - body copy, labels, rules
const PANEL = 'F2F2F2';    // bg1 @ 95% - the big off-white content panels
const TILE = 'D9D9D9';     // bg1 @ 85% - light half of the infographic tiles
const PHOTO = 'DDDDDD';    // stand-in fill for photographs
const WHITE = 'FFFFFF';

// ------------------------------------------------------------------ fonts --
const DISPLAY = 'Roboto Medium';  // headings + numerals
const SANS = 'Lato';              // kickers, footers
const LIGHT = 'Lato Light';       // body copy
const HEAVY = 'Lato Black';       // bold labels

// Soft 20%-black drop shadow used by every off-white panel. pptxgenjs mutates
// the options object while writing XML, so hand each shape its own copy.
// Angle 360 rather than 0: pptxgenjs treats a falsy angle as "unset" and
// substitutes 270, which would throw the shadow downwards instead of right.
function panelShadow(angle, offset) {
  return { type: 'outer', color: '000000', opacity: 0.2, blur: 16, offset: offset || 8, angle: angle || 360 };
}

// The off-white content panel that anchors most layouts.
function panel(s, x, y, w, h, shadow, transparency) {
  s.addShape('rect', { x, y, w, h, fill: { color: PANEL, transparency: transparency || 0 }, shadow });
}

// --------------------------------------------------------------- primitives --
// PowerPoint text boxes anchor text at the top; pptxgenjs defaults to middle.
function text(s, body, opts) {
  s.addText(body, Object.assign({ valign: 'top', isTextBox: true, fit: 'resize' }, opts));
}

function heading(s, x, y, w, h, body, extra) {
  text(s, body, Object.assign({ x, y, w, h, fontFace: DISPLAY, fontSize: 36, color: INK }, extra));
}

function stat(s, x, y, w, h, body, extra) {
  text(s, body, Object.assign({ x, y, w, h, fontFace: DISPLAY, fontSize: 28, color: INK }, extra));
}

function body(s, x, y, w, h, copy, extra) {
  text(s, copy, Object.assign({ x, y, w, h, fontFace: LIGHT, fontSize: 9, color: GRAY,
    lineSpacingMultiple: 1.5 }, extra));
}

function label(s, x, y, w, h, copy, extra) {
  text(s, copy, Object.assign({ x, y, w, h, fontFace: HEAVY, fontSize: 9, color: GRAY, bold: true,
    lineSpacingMultiple: 1.5 }, extra));
}

// Wingdings check-mark bullet list (slide 12).
function checklist(s, x, y, w, h, lines) {
  text(s, lines.map((line, i) => ({
    text: line,
    options: { bullet: { characterCode: '2713', indent: 13.5 }, breakLine: i < lines.length - 1 },
  })), { x, y, w, h, fontFace: LIGHT, fontSize: 9, color: GRAY, lineSpacingMultiple: 1.5 });
}

// Eyebrow line that sits above almost every headline.
function kicker(s, x, y) {
  text(s, 'ARCHINEA STUDIO PRESENTATION TEMPLATE',
    { x, y, w: 3.268, h: 0.269, fontFace: SANS, fontSize: 10, color: GRAY, wrap: false });
}

// "07/32" page marker.
function pageNum(s, x, y, page) {
  text(s, [{ text: page, options: { color: INK } }, { text: '/32', options: { color: GRAY } }],
    { x, y, w: 0.877, h: 0.404, fontFace: DISPLAY, fontSize: 18, wrap: false });
}

// "2021 (c) Archinea Studio / www.archineastudio.com" footer block.
function credit(s, x, y) {
  text(s, [
    { text: '2021', options: { bold: true, fontFace: HEAVY, breakLine: false } },
    { text: ' \u00A9 Archinea Studio', options: { breakLine: true } },
    { text: 'www.archineastudio.com ' },
  ], { x, y, w: 1.865, h: 0.404, fontFace: SANS, fontSize: 9, color: GRAY });
}

// Hairline divider (vertical in footers, horizontal on the cover).
function rule(s, x, y, w, h, color) {
  s.addShape('line', { x, y, w, h, line: { color: color || GRAY, width: 0.75 } });
}

// Full-bleed / cropped photo stand-in.
function photo(s, x, y, w, h) {
  s.addShape('rect', { x, y, w, h, fill: { color: PHOTO } });
  text(s, '[image]', { x, y: y + h / 2 - 0.2, w, h: 0.4, align: 'center', valign: 'middle',
    fontFace: SANS, fontSize: 12, color: GRAY });
}

// One quadrant of the four-part arrow pinwheel on slide 30.
// Path is the source freeform, normalised to the 1.803 x 2.008 in shape box.
const ARROW_TILE = [
  { x: 0, y: 0.742, moveTo: true }, { x: 0, y: 0.205 }, { x: 0.5, y: 0.205 }, { x: 0.705, y: 0 },
  { x: 0.911, y: 0.205 }, { x: 1.413, y: 0.205 }, { x: 1.803, y: 0.595 }, { x: 1.803, y: 1.043 },
  { x: 1.543, y: 1.303 }, { x: 1.803, y: 1.563 }, { x: 1.803, y: 2.008 }, { x: 1.267, y: 2.008 },
  { close: true },
];

function arrowTile(s, x, y, rotate, color) {
  s.addShape('custGeom', {
    x, y, w: 1.803, h: 2.008, rotate, points: ARROW_TILE, fill: { color }, line: { type: 'none' },
  });
}

const OUTPUT_NAME = '15180b4e-e785-4af2-bdb0-5d91dc838c4f_grok_final.pptx';
const CHART_CATEGORIES = ['Category 1', 'Category 2', 'Category 3', 'Category 4'];

// ------------------------------------------------------------------ slides --
function slide01(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: 13.335, h: 7.5, fill: { color: WHITE, transparency: 70 } });
  text(s, "Archinea", { x: 0.84, y: 2.581, w: 6.35, h: 1.952, fontSize: 110, fontFace: DISPLAY, color: INK, wrap: false });
  text(s, "ARCHITECTURE AND INTERIOR DESIGN STUDIO", { x: 0.84, y: 4.585, w: 2.205, h: 0.62, fontSize: 11, fontFace: SANS, color: INK, lineSpacingMultiple: 1.5 });
  text(s, [{ text: "© Archinea Studio", options: { breakLine: true } }, { text: "www.archineastudio.com " }], { x: 0.84, y: 6.296, w: 1.865, h: 0.525, fontSize: 9, fontFace: SANS, color: INK, lineSpacingMultiple: 1.5 });
  text(s, "2021", { x: 0.84, y: 5.959, w: 0.651, h: 0.337, fontSize: 14, fontFace: DISPLAY, color: INK, wrap: false });
  rule(s, 1.491, 6.127, 1.554, 0, INK);
}

function slide02(pptx) {
  const s = pptx.addSlide();
  panel(s, 0, 0, 5.408, 7.5, panelShadow());
  heading(s, 0.919, 1.5, 3.622, 2.524, "We Shape Our Buildings, Therefore They Shape Us");
  body(s, 0.919, 4.424, 3.622, 0.98, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ");
  kicker(s, 0.919, 0.994);
  credit(s, 2.204, 6.585);
  pageNum(s, 0.919, 6.585, "02");
  rule(s, 2.021, 6.506, 0, 0.561);
}

function slide03(pptx) {
  const s = pptx.addSlide();
  panel(s, 8.321, 0, 5.014, 7.5, panelShadow(180));
  body(s, 9.27, 4.73, 3.307, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud", { color: MID });
  heading(s, 9.266, 1.288, 3.334, 3.13, "Recognizing The Need Is The Primary Condition For Design");
  kicker(s, 9.266, 0.836);
  credit(s, 10.551, 6.585);
  pageNum(s, 9.266, 6.585, "03");
  rule(s, 10.369, 6.506, 0, 0.561);
}

function slide04(pptx) {
  const s = pptx.addSlide();
  panel(s, 0, 0, 9.03, 7.5, panelShadow());
  heading(s, 0.999, 1.408, 3.229, 3.13, "Everything Is Designed. Few Things Are Designed Well");
  body(s, 5.227, 0.915, 3.018, 3.933, [{ text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ", options: { breakLine: true } }, { text: "", options: { breakLine: true } }, { text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum." }]);
  kicker(s, 0.999, 0.915);
  credit(s, 6.512, 6.585);
  pageNum(s, 5.227, 6.585, "04");
  rule(s, 6.329, 6.506, 0, 0.561);
  heading(s, 0.999, 5.529, 1.865, 0.707, "4.3K+");
  body(s, 0.999, 6.236, 1.865, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor");
}

function slide05(pptx) {
  const s = pptx.addSlide();
  panel(s, 4.899, 0, 8.437, 7.5, panelShadow(180));
  heading(s, 6.116, 1.268, 6.221, 2.524, "We Should Attempt To Bring Nature, Houses, And Human Beings Together In A Higher Unity");
  kicker(s, 6.116, 0.775);
  stat(s, 6.118, 4.301, 1.708, 0.572, "4.846");
  body(s, 6.116, 4.87, 1.708, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor");
  credit(s, 7.401, 6.585);
  pageNum(s, 6.116, 6.585, "05");
  rule(s, 7.219, 6.506, 0, 0.561);
  stat(s, 8.375, 4.301, 1.708, 0.572, "2.8K+");
  body(s, 8.373, 4.87, 1.708, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor");
  stat(s, 10.632, 4.301, 1.708, 0.572, "84K+");
  body(s, 10.63, 4.87, 1.708, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor");
}

function slide06(pptx) {
  const s = pptx.addSlide();
  body(s, 8.51, 4.957, 4.063, 0.98, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ");
  label(s, 8.51, 4.659, 4.063, 0.298, "LOREM IPSUM DOLOR SIT AMET, CONSECTETUR ADIPISCING");
  credit(s, 9.795, 6.585);
  pageNum(s, 8.51, 6.585, "06");
  rule(s, 9.613, 6.506, 0, 0.561);
  panel(s, 0, 0, 7.534, 4.089, panelShadow());
  heading(s, 0.84, 1.332, 6.142, 1.919, "All Fine Architectural Values Are Human Values, Else Not Valuable");
  kicker(s, 0.84, 0.839);
}

function slide07(pptx) {
  const s = pptx.addSlide();
  panel(s, 4.306, 0, 9.029, 7.5, panelShadow(180));
  label(s, 9.502, 3.067, 3.105, 0.298, "LOREM IPSUM DOLOR SIT AMET");
  body(s, 9.502, 3.367, 3.105, 1.661, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ");
  label(s, 9.502, 1.173, 3.105, 0.298, "LOREM IPSUM DOLOR SIT AMET CONSECTETUR");
  body(s, 9.502, 1.472, 3.072, 1.207, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ");
  credit(s, 10.787, 6.585);
  pageNum(s, 9.502, 6.585, "07");
  rule(s, 10.605, 6.506, 0, 0.561);
  heading(s, 5.269, 1.619, 3.28, 2.524, "A Design Isn't Finished Until Someone Is Using It");
  kicker(s, 5.264, 1.173);
  heading(s, 5.264, 5.757, 1.865, 0.707, "630+");
  body(s, 5.264, 6.463, 2.43, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor");
  label(s, 5.259, 5.458, 2.43, 0.298, "LOREM IPSUM DOLOR SIT AMET");
}

function slide08(pptx) {
  const s = pptx.addSlide();
  heading(s, 10.783, 0.835, 1.708, 0.707, "4.846");
  body(s, 10.781, 1.496, 1.951, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor");
  heading(s, 10.781, 2.943, 1.708, 0.707, "2,8K+");
  body(s, 10.781, 3.65, 1.951, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor");
  heading(s, 10.783, 5.096, 1.708, 0.707, "84K+");
  body(s, 10.781, 5.803, 1.951, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor");
  panel(s, 0, 0, 5.576, 7.5, panelShadow());
  body(s, 0.919, 4.503, 3.622, 0.98, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ", { align: "justify" });
  heading(s, 0.919, 1.578, 3.622, 2.524, "We Shape Our Buildings, Therefore They Shape Us");
  kicker(s, 0.919, 1.073);
  credit(s, 2.204, 6.585);
  pageNum(s, 0.919, 6.585, "08");
  rule(s, 2.021, 6.506, 0, 0.561);
}

function slide09(pptx) {
  const s = pptx.addSlide();
  heading(s, 0.919, 5.408, 5.827, 1.313, "Design Is Where Science And Art Break Even");
  kicker(s, 0.919, 4.914);
  credit(s, 9.527, 6.585);
  pageNum(s, 8.243, 6.585, "09");
  rule(s, 9.345, 6.506, 0, 0.561);
  heading(s, 8.247, 4.914, 4.159, 0.707, [{ text: "02 - 07 " }, { text: "/08", options: { fontSize: 24, color: GRAY } }]);
  body(s, 8.243, 5.621, 4.164, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim");
  label(s, 8.243, 4.616, 4.164, 0.298, "LOREM IPSUM DOLOR SIT AMET CONSECTETUR ADIPISCING");
}

function slide10(pptx) {
  const s = pptx.addSlide();
  panel(s, 5.408, 1.073, 7.324, 6.427, panelShadow(45, 16), 10);
  body(s, 6.239, 5.202, 5.666, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ");
  heading(s, 6.235, 2.302, 5.67, 2.524, "PLACEHOLDER");
  kicker(s, 6.235, 1.831);
  credit(s, 7.52, 6.585);
  pageNum(s, 6.235, 6.585, "10");
  rule(s, 7.337, 6.506, 0, 0.561);
}

function slide11(pptx) {
  const s = pptx.addSlide();
  panel(s, 0.84, 0, 7.534, 6.427, panelShadow(45, 16), 10);
  heading(s, 1.601, 1.195, 6.221, 2.524, "We Should Attempt To Bring Nature, Houses, And Human Beings Together In A Higher Unity");
  kicker(s, 1.595, 0.709);
  credit(s, 2.892, 5.512);
  pageNum(s, 1.607, 5.512, "11");
  rule(s, 2.709, 5.434, 0, 0.561);
  heading(s, 1.595, 4.067, 1.865, 0.707, [{ text: "21 " }, { text: "/08", options: { fontSize: 24, color: GRAY } }]);
  body(s, 3.203, 4.37, 4.619, 0.298, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor");
  label(s, 3.203, 4.067, 4.164, 0.298, "LOREM IPSUM DOLOR SIT AMET CONSECTETUR ADIPISCING");
}

function slide12(pptx) {
  const s = pptx.addSlide();
  panel(s, 0, 0, 4.856, 7.5, panelShadow());
  text(s, "LOREM IPSUM DOLOR SIT AMET", { x: 5.254, y: 4.83, w: 2.492, h: 0.298, fontSize: 9, fontFace: SANS, color: GRAY, align: "justify", lineSpacingMultiple: 1.5 });
  text(s, "Adam L. Mckinney", { x: 5.25, y: 4.38, w: 2.525, h: 0.438, fontSize: 20, fontFace: DISPLAY, color: INK, wrap: false });
  body(s, 5.267, 5.36, 3.425, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna");
  text(s, "LOREM IPSUM DOLOR SIT AMET", { x: 9.385, y: 4.83, w: 2.492, h: 0.298, fontSize: 9, fontFace: SANS, color: GRAY, lineSpacingMultiple: 1.5 });
  text(s, "Henrietta Stevenson", { x: 9.385, y: 4.392, w: 2.758, h: 0.438, fontSize: 20, fontFace: DISPLAY, color: INK, wrap: false });
  body(s, 9.385, 5.36, 3.425, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna");
  body(s, 0.849, 4.725, 3.28, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis", { align: "justify" });
  label(s, 0.849, 4.427, 3.28, 0.298, "LOREM IPSUM DOLOR SIT AMET CONSECTETUR");
  heading(s, 0.849, 1.361, 3.28, 2.524, "A Design Isn't Finished Until Someone Is Using It");
  kicker(s, 0.844, 0.915);
  credit(s, 2.125, 6.585);
  pageNum(s, 0.84, 6.585, "12");
  rule(s, 1.943, 6.506, 0, 0.561);
  checklist(s, 5.267, 6.46, 3.425, 0.525, ["Lorem ipsum dolor sit amet, consectetur adipiscing elit", "Sed do eiusmod tempor incididunt ut labore et dolore"]);
  checklist(s, 9.385, 6.46, 3.425, 0.525, ["Lorem ipsum dolor sit amet, consectetur adipiscing elit", "Sed do eiusmod tempor incididunt ut labore et dolore"]);
  rule(s, 5.329, 6.191, 3.425, 0);
  rule(s, 9.464, 6.191, 3.425, 0);
}

function slide13(pptx) {
  const s = pptx.addSlide();
  panel(s, 0, 0, 8.085, 7.5, panelShadow());
  text(s, "LOREM IPSUM DOLOR SIT AMET", { x: 1.183, y: 1.198, w: 2.492, h: 0.298, fontSize: 9, fontFace: SANS, color: GRAY, align: "right", lineSpacingMultiple: 1.5 });
  text(s, "Trevor J. Roberson", { x: 1.08, y: 0.707, w: 2.595, h: 0.438, fontSize: 20, fontFace: DISPLAY, color: INK, align: "right", wrap: false });
  body(s, 0.367, 1.582, 3.307, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore", { color: MID, align: "right" });
  s.addShape('triangle', { x: 3.345, y: 0.406, w: 0.197, h: 0.17, rotate: 90, fill: { color: GRAY } });
  text(s, "LOREM IPSUM DOLOR SIT AMET", { x: 1.183, y: 6.2, w: 2.492, h: 0.298, fontSize: 9, fontFace: SANS, color: GRAY, align: "right", lineSpacingMultiple: 1.5 });
  text(s, "Henrietta Stevenson", { x: 0.917, y: 5.762, w: 2.758, h: 0.438, fontSize: 20, fontFace: DISPLAY, color: INK, align: "right", wrap: false });
  body(s, 0.367, 6.526, 3.307, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore", { color: MID, align: "right" });
  s.addShape('triangle', { x: 3.345, y: 5.462, w: 0.197, h: 0.17, rotate: 90, fill: { color: GRAY } });
  text(s, "LOREM IPSUM DOLOR SIT AMET", { x: 4.463, y: 3.717, w: 2.492, h: 0.298, fontSize: 9, fontFace: SANS, color: GRAY, lineSpacingMultiple: 1.5 });
  text(s, "Johnathan Bishop", { x: 4.463, y: 3.279, w: 2.477, h: 0.438, fontSize: 20, fontFace: DISPLAY, color: INK, wrap: false });
  body(s, 4.463, 4.017, 3.307, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore", { color: MID });
  s.addShape('triangle', { x: 4.567, y: 2.971, w: 0.197, h: 0.17, rotate: 270, fill: { color: GRAY } });
  body(s, 9.191, 4.73, 3.307, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud", { color: MID });
  heading(s, 9.188, 1.288, 3.334, 3.13, "Recognizing The Need Is The Primary Condition For Design");
  kicker(s, 9.188, 0.836);
  credit(s, 10.472, 6.585);
  pageNum(s, 9.188, 6.585, "13");
  rule(s, 10.29, 6.506, 0, 0.561);
}

function slide14(pptx) {
  const s = pptx.addSlide();
  panel(s, 5.124, 0, 8.211, 7.5, panelShadow(180));
  s.addShape('rect', { x: 5.58, y: 2.852, w: 3.514, h: 0.779, fill: { color: GRAY } });
  s.addShape('rect', { x: 9.365, y: 2.85, w: 3.514, h: 0.779, fill: { color: GRAY } });
  s.addShape('rect', { x: 5.58, y: 6.262, w: 3.514, h: 0.779, fill: { color: GRAY } });
  s.addShape('rect', { x: 9.365, y: 6.262, w: 3.514, h: 0.779, fill: { color: GRAY } });
  heading(s, 0.787, 1.457, 3.675, 0.707, "Meet Our Team");
  text(s, "TREVOR J. ROBERSON", { x: 6.464, y: 2.96, w: 1.746, h: 0.252, fontSize: 9, fontFace: HEAVY, color: WHITE, bold: true, align: "center" });
  body(s, 5.939, 3.213, 2.794, 0.298, "Lorem ipsum dolor sit amet, consectetur", { color: WHITE, align: "center" });
  text(s, "JOHNATHAN BISHOP", { x: 10.268, y: 2.951, w: 1.706, h: 0.252, fontSize: 9, fontFace: HEAVY, color: WHITE, bold: true, align: "center" });
  body(s, 9.724, 3.203, 2.794, 0.298, "Lorem ipsum dolor sit amet, consectetur", { color: WHITE, align: "center" });
  text(s, "HENRIETTA STEVENSON", { x: 6.386, y: 6.376, w: 1.902, h: 0.252, fontSize: 9, fontFace: HEAVY, color: WHITE, bold: true, align: "center" });
  body(s, 5.939, 6.628, 2.794, 0.298, "Lorem ipsum dolor sit amet, consectetur", { color: WHITE, align: "center" });
  text(s, "ADAM L. MCKINNEY", { x: 10.311, y: 6.376, w: 1.621, h: 0.252, fontSize: 9, fontFace: HEAVY, color: WHITE, bold: true, align: "center" });
  body(s, 9.724, 6.628, 2.794, 0.298, "Lorem ipsum dolor sit amet, consectetur", { color: WHITE, align: "center" });
  kicker(s, 0.787, 0.994);
  credit(s, 2.072, 6.585);
  pageNum(s, 0.787, 6.585, "14");
  rule(s, 1.89, 6.506, 0, 0.561);
  body(s, 0.787, 2.85, 3.64, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis");
  label(s, 0.787, 2.552, 3.28, 0.298, "LOREM IPSUM DOLOR SIT AMET CONSECTETUR");
  stat(s, 0.789, 4.122, 1.169, 0.572, "4.84");
  body(s, 1.89, 4.117, 2.441, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor");
  stat(s, 0.789, 4.91, 1.169, 0.572, "1.8K");
  body(s, 1.89, 4.905, 2.441, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor");
}

function slide15(pptx) {
  const s = pptx.addSlide();
  panel(s, 0, 0, 7.815, 7.5, panelShadow());
  heading(s, 1.077, 1.5, 4.602, 0.707, "Marianne Jefferson");
  kicker(s, 1.077, 0.994);
  label(s, 1.076, 2.542, 3.105, 0.298, "FOUNDER AND CEO OF ARCHINEA STUDIO");
  body(s, 1.076, 2.84, 5.67, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo ");
  stat(s, 1.073, 4.459, 2.43, 0.572, [{ text: "91.6 " }, { text: "/100", options: { fontSize: 18, color: GRAY } }]);
  body(s, 1.076, 5.036, 2.43, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor");
  credit(s, 2.361, 6.585);
  pageNum(s, 1.076, 6.585, "15");
  rule(s, 2.179, 6.506, 0, 0.561);
  label(s, 1.076, 4.155, 2.43, 0.298, "LOREM IPSUM DOLOR SIT AMET");
  stat(s, 4.317, 4.459, 2.43, 0.572, [{ text: "88.9 " }, { text: "/100", options: { fontSize: 18, color: GRAY } }]);
  body(s, 4.317, 5.032, 2.43, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor");
  label(s, 4.317, 4.163, 2.43, 0.298, "LOREM IPSUM DOLOR SIT AMET");
}

function slide16(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: 13.335, h: 7.5, fill: { color: WHITE, transparency: 30 } });
  heading(s, 0.919, 1.151, 6.064, 3.736, "“As An Architect You Design For The Present, With An Awareness Of The Past, For A Future Which Is Essentially Unknown”");
  text(s, "— NORMAN FOSTER", { x: 0.919, y: 5.292, w: 1.538, h: 0.269, fontSize: 10, fontFace: SANS, color: GRAY, wrap: false });
  kicker(s, 0.919, 0.679);
  credit(s, 2.204, 6.585);
  pageNum(s, 0.919, 6.585, "16");
  rule(s, 2.021, 6.506, 0, 0.561);
}

function slide17(pptx) {
  const s = pptx.addSlide();
  panel(s, 0, 0, 4.62, 7.5, panelShadow());
  heading(s, 0.722, 1.122, 3.268, 3.13, "Everything Is Designed. Few Things Are Designed Well");
  kicker(s, 0.722, 0.659);
  text(s, "Lorem Ipsum Building", { x: 5.014, y: 6.221, w: 2.951, h: 0.438, fontSize: 20, fontFace: DISPLAY, color: INK, wrap: false });
  body(s, 5.014, 6.66, 3.596, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna");
  text(s, "Lorem Ipsum Building", { x: 9.214, y: 6.221, w: 2.951, h: 0.438, fontSize: 20, fontFace: DISPLAY, color: INK, wrap: false });
  body(s, 9.214, 6.66, 3.596, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna");
  credit(s, 2.007, 6.585);
  pageNum(s, 0.722, 6.585, "17");
  rule(s, 1.824, 6.506, 0, 0.561);
  heading(s, 0.727, 4.644, 1.767, 0.707, [{ text: "21 " }, { text: "/08", options: { fontSize: 24, color: GRAY } }]);
  body(s, 0.723, 5.351, 1.771, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, ");
}

function slide18(pptx) {
  const s = pptx.addSlide();
  heading(s, 0.604, 1.168, 4.017, 1.919, "Simplicity Is The Ultimate Sophistication");
  kicker(s, 0.604, 0.663);
  label(s, 5.305, 4.519, 3.105, 0.298, "LOREM IPSUM DOLOR SIT AMET CONSECT");
  body(s, 5.305, 4.818, 3.188, 0.98, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip");
  credit(s, 6.59, 6.585);
  pageNum(s, 5.305, 6.585, "18");
  rule(s, 6.408, 6.506, 0, 0.561);
}

function slide19(pptx) {
  const s = pptx.addSlide();
  heading(s, 0.919, 1.343, 3.859, 1.919, "People Ignore Design That Ignores People");
  body(s, 1.785, 3.75, 2.993, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam");
  kicker(s, 0.919, 0.836);
  credit(s, 2.204, 6.585);
  pageNum(s, 0.919, 6.585, "19");
  rule(s, 2.021, 6.506, 0, 0.561);
  stat(s, 0.919, 3.768, 0.945, 0.572, "01.");
  body(s, 1.785, 4.699, 2.993, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam");
  stat(s, 0.919, 4.716, 0.945, 0.572, "02.");
}

function slide20(pptx) {
  const s = pptx.addSlide();
  heading(s, 4.778, 4.972, 3.849, 1.919, "Design Is Where Science And Art Break Even");
  kicker(s, 4.778, 4.517);
  credit(s, 1.967, 2.884);
  pageNum(s, 0.683, 2.884, "20");
  rule(s, 1.785, 2.805, 0, 0.561);
  heading(s, 0.687, 1.056, 3.255, 0.707, [{ text: "21 - 25 " }, { text: "/08", options: { fontSize: 24, color: GRAY } }]);
  body(s, 0.684, 1.763, 3.26, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore");
  label(s, 0.683, 0.758, 3.26, 0.298, "LOREM IPSUM DOLOR SIT AMET CONSECTETUR");
}

function slide21(pptx) {
  const s = pptx.addSlide();
  heading(s, 8.873, 1.137, 3.537, 2.524, "Everything Is Designed. Few Things Are Designed Well");
  kicker(s, 8.873, 0.679);
  credit(s, 10.157, 6.585);
  pageNum(s, 8.873, 6.585, "21");
  rule(s, 9.975, 6.506, 0, 0.561);
  stat(s, 8.875, 4.07, 1.283, 0.572, "4.4K+");
  body(s, 10.243, 4.065, 1.865, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor");
  stat(s, 8.875, 5.076, 1.283, 0.572, "1.8K+");
  body(s, 10.243, 5.071, 1.865, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor");
}

function slide22(pptx) {
  const s = pptx.addSlide();
  body(s, 6.983, 6.236, 5.725, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum");
  credit(s, 1.81, 6.585);
  pageNum(s, 0.525, 6.585, "22");
  rule(s, 1.628, 6.506, 0, 0.561);
  heading(s, 0.525, 3.784, 4.646, 1.313, "An Idea Is Salvation By Imagination");
  kicker(s, 0.525, 3.278);
  body(s, 0.525, 5.43, 5.591, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris", { align: "justify" });
}

function slide23(pptx) {
  const s = pptx.addSlide();
  panel(s, 0, 3.829, 5.014, 3.671, panelShadow());
  body(s, 7.038, 4.521, 5.714, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut");
  label(s, 7.029, 4.222, 4.635, 0.298, "LOREM IPSUM DOLOR SIT AMET, CONSECTETUR ADIPISCING ELIT");
  body(s, 7.038, 5.587, 5.714, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut");
  label(s, 7.029, 5.289, 4.635, 0.298, "LOREM IPSUM DOLOR SIT AMET, CONSECTETUR ADIPISCING ELIT");
  stat(s, 5.869, 4.222, 1.088, 0.572, "92.7");
  stat(s, 5.869, 5.294, 1.088, 0.572, "74.8");
  heading(s, 0.525, 4.958, 4.017, 1.919, "Simplicity Is The Ultimate Sophistication");
  kicker(s, 0.525, 4.452);
  credit(s, 7.154, 6.585);
  pageNum(s, 5.869, 6.585, "23");
  rule(s, 6.972, 6.506, 0, 0.561);
}

function slide24(pptx) {
  const s = pptx.addSlide();
  panel(s, 0, 0, 9.607, 7.5, panelShadow());
  heading(s, 4.384, 1.107, 4.804, 3.13, "We Should Attempt To Bring Nature, Houses, And Human Beings Together In A Higher Unity");
  kicker(s, 4.384, 0.6);
  credit(s, 5.669, 6.585);
  pageNum(s, 4.384, 6.585, "24");
  rule(s, 5.486, 6.506, 0, 0.561);
  stat(s, 4.39, 4.633, 1.708, 0.572, "2.8K+");
  body(s, 4.388, 5.202, 2.103, 0.525, "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ");
  stat(s, 6.929, 4.633, 1.708, 0.572, "8.4K+");
  body(s, 6.927, 5.202, 2.103, 0.525, "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ");
}

function slide25(pptx) {
  const s = pptx.addSlide();
  panel(s, 5.565, 0, 7.77, 7.5, panelShadow(180));
  photo(s, 0.7, 0.68, 4.05, 6.82);   // phone mockup, bleeds off the bottom
  heading(s, 6.668, 1.18, 5.749, 1.919, "Recognizing The Need Is The Primary Condition For Design");
  kicker(s, 6.668, 0.679);
  stat(s, 6.667, 3.443, 1.708, 0.572, "4.846");
  body(s, 6.665, 4.012, 1.708, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit,");
  stat(s, 8.688, 3.443, 1.708, 0.572, "2.8K+");
  body(s, 8.686, 4.012, 1.708, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit,");
  stat(s, 10.709, 3.443, 1.708, 0.572, "84K+");
  body(s, 10.706, 4.012, 1.708, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit,");
  body(s, 6.665, 5.202, 5.749, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut");
  label(s, 6.664, 4.904, 4.635, 0.298, "LOREM IPSUM DOLOR SIT AMET, CONSECTETUR ADIPISCING ELIT");
  credit(s, 7.949, 6.585);
  pageNum(s, 6.664, 6.585, "25");
  rule(s, 7.766, 6.506, 0, 0.561);
}

function slide26(pptx) {
  const s = pptx.addSlide();
  panel(s, 9.817, 0.033, 3.518, 7.5, panelShadow(180));
  photo(s, 0, 0.94, 5.45, 5.62);     // laptop mockup, cropped by the left edge
  heading(s, 6.116, 1.397, 3.268, 1.919, "An Idea Is Salvation By Imagination");
  body(s, 10.356, 2.604, 2.441, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore", { align: "center" });
  label(s, 10.356, 2.305, 2.441, 0.298, "LOREM IPSUM DOLOR SIT AMET", { align: "center" });
  body(s, 10.356, 5.911, 2.441, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore", { align: "center" });
  label(s, 10.356, 5.613, 2.441, 0.298, "LOREM IPSUM DOLOR SIT AMET", { align: "center" });
  kicker(s, 6.116, 0.915);
  stat(s, 6.122, 4.149, 1.865, 0.572, "630+");
  body(s, 6.122, 4.721, 2.43, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor");
  label(s, 6.122, 3.85, 2.43, 0.298, "LOREM IPSUM DOLOR SIT AMET");
  credit(s, 7.519, 6.181);
  pageNum(s, 6.234, 6.181, "26");
  rule(s, 7.337, 6.102, 0, 0.561);
}

function slide27(pptx) {
  const s = pptx.addSlide();
  photo(s, 8.4, 1.08, 4.935, 5.75);  // desktop mockup, cropped by the right edge
  body(s, 4.784, 1.672, 2.978, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ");
  label(s, 5.437, 1.107, 2.325, 0.525, "LOREM IPSUM DOLOR SIT AMET CONSECTETUR");
  body(s, 0.761, 4.466, 3.279, 1.207, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ");
  credit(s, 2.046, 6.585);
  pageNum(s, 0.761, 6.585, "27");
  rule(s, 1.864, 6.506, 0, 0.561);
  heading(s, 0.761, 1.594, 3.279, 2.524, "A Design Isn't Finished Until Someone Is Using It");
  kicker(s, 0.761, 1.106);
  stat(s, 4.778, 1.106, 0.945, 0.572, "01.");
  body(s, 4.784, 3.391, 2.978, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ");
  label(s, 5.437, 2.827, 2.325, 0.525, "LOREM IPSUM DOLOR SIT AMET CONSECTETUR");
  stat(s, 4.778, 2.825, 0.945, 0.572, "02.");
  body(s, 4.784, 5.045, 2.978, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ");
  label(s, 5.437, 4.48, 2.325, 0.525, "LOREM IPSUM DOLOR SIT AMET CONSECTETUR");
  stat(s, 4.778, 4.479, 0.945, 0.572, "03.");
}

function slide28(pptx) {
  const s = pptx.addSlide();
  panel(s, 0, 0, 5.486, 7.5, panelShadow());
  // Clustered columns: 3 series x 4 categories, no legend, no category axis.
  s.addChart('bar', [
    { name: 'Series 1', labels: CHART_CATEGORIES, values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels: CHART_CATEGORIES, values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels: CHART_CATEGORIES, values: [2.0, 2.0, 3.0, 5.3] },
  ], {
    x: 6.353, y: 0.787, w: 6.379, h: 5.927,
    barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 180, barOverlapPct: -40,
    chartColors: [INK, GRAY, TILE],
    showLegend: false, showTitle: false,
    catAxisHidden: true, catAxisLineShow: false,
    valAxisLineShow: false, valAxisLabelFontFace: LIGHT, valAxisLabelFontSize: 9,
    valAxisLabelColor: GRAY, valAxisMajorUnit: 1, valAxisMaxVal: 6,
    valGridLine: { style: 'solid', color: TILE, size: 0.75 },
  });
  heading(s, 0.802, 1.257, 4.369, 1.919, "Archinea Studio Chart and Infographic");
  body(s, 0.802, 3.523, 4.095, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip");
  credit(s, 2.087, 6.585);
  pageNum(s, 0.802, 6.585, "28");
  rule(s, 1.905, 6.506, 0, 0.561);
  kicker(s, 0.802, 0.784);
  stat(s, 0.804, 4.624, 1.708, 0.572, "83.64%");
  body(s, 0.802, 5.193, 1.708, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit,");
  stat(s, 3.19, 4.624, 1.708, 0.572, "8.49K+");
  body(s, 3.188, 5.193, 1.708, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit,");
}

function slide29(pptx) {
  const s = pptx.addSlide();
  panel(s, 0, 0, 7.534, 7.5, panelShadow());
  heading(s, 8.674, 1.422, 4.136, 1.919, "Simplicity Is The Ultimate Sophistication");
  // Doughnut with an 80% hole; the percentage sits in the middle (next shape).
  s.addChart('doughnut', [
    { name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: [8, 5] },
  ], {
    x: 1.326, y: 0.6, w: 4.882, h: 4.962,
    holeSize: 80, firstSliceAng: 0, chartColors: [GRAY, TILE],
    dataBorder: { pt: 1.5, color: WHITE },
    showLegend: false, showTitle: false, showValue: false,
  });
  text(s, "62.73%", { x: 2.394, y: 2.66, w: 2.745, h: 0.841, fontSize: 44, fontFace: DISPLAY, color: INK, align: "center" });
  body(s, 1.326, 5.92, 4.882, 0.98, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore", { color: MID, align: "center" });
  credit(s, 9.955, 6.585);
  pageNum(s, 8.67, 6.585, "29");
  rule(s, 9.773, 6.506, 0, 0.561);
  kicker(s, 8.67, 0.915);
  heading(s, 8.67, 3.974, 1.943, 0.707, "68.4K+");
  body(s, 8.67, 4.68, 3.518, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud");
  label(s, 8.67, 3.675, 3.518, 0.298, "LOREM IPSUM DOLOR SIT AMET CONSECTETUR");
}

function slide30(pptx) {
  const s = pptx.addSlide();
  arrowTile(s, 1.022, 1.761, 90, INK);
  arrowTile(s, 2.907, 1.864, 180, TILE);
  arrowTile(s, 2.804, 3.73, 270, INK);
  arrowTile(s, 0.919, 3.627, 0, TILE);
  heading(s, 1.384, 2.683, 1.009, 0.707, "01.", { color: WHITE, align: "right" });
  heading(s, 3.21, 2.643, 1.009, 0.707, "02.", { color: GRAY });
  heading(s, 3.21, 4.137, 1.009, 0.707, "03.", { color: WHITE });
  heading(s, 1.384, 4.141, 1.009, 0.707, "04.", { color: GRAY, align: "right" });
  stat(s, 5.821, 3.195, 0.884, 0.572, "01.");
  body(s, 6.584, 3.493, 2.433, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore");
  label(s, 6.584, 3.195, 2.433, 0.298, "LOREM IPSUM DOLOR SIT AMET");
  stat(s, 5.821, 4.585, 0.884, 0.572, "03.");
  body(s, 6.584, 4.883, 2.433, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore");
  label(s, 6.584, 4.585, 2.433, 0.298, "LOREM IPSUM DOLOR SIT AMET");
  stat(s, 9.456, 3.195, 0.884, 0.572, "02.");
  body(s, 10.22, 3.493, 2.433, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore");
  label(s, 10.22, 3.195, 2.433, 0.298, "LOREM IPSUM DOLOR SIT AMET");
  stat(s, 9.456, 4.561, 0.884, 0.572, "04.");
  body(s, 10.22, 4.86, 2.433, 0.753, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore");
  label(s, 10.22, 4.561, 2.433, 0.298, "LOREM IPSUM DOLOR SIT AMET");
  heading(s, 5.821, 1.509, 4.701, 1.313, "Archia Studio Chart and Infographic");
  kicker(s, 5.821, 1.06);
  credit(s, 7.106, 6.585);
  pageNum(s, 5.821, 6.585, "30");
  rule(s, 6.924, 6.506, 0, 0.561);
}

function slide31(pptx) {
  const s = pptx.addSlide();
  panel(s, 6.668, 3.986, 6.668, 3.514, panelShadow(180));
  body(s, 7.841, 5.026, 1.732, 0.525, [{ text: "1047 Birch Street, El Paso", options: { breakLine: true } }, { text: "Texas 79922, United States" }]);
  label(s, 7.842, 4.728, 1.732, 0.298, "ADDRESS");
  body(s, 7.836, 6.233, 1.732, 0.525, [{ text: "www.archiastudio.com ", options: { breakLine: true } }, { text: "info@ archiastudio.com " }]);
  label(s, 7.836, 5.935, 1.732, 0.298, "WEBSITE");
  body(s, 10.356, 5.033, 1.732, 0.525, [{ text: "+000 1234 5678 90", options: { breakLine: true } }, { text: "+000 1234 5678 90" }]);
  label(s, 10.356, 4.735, 1.732, 0.298, "PHONE");
  body(s, 10.356, 6.226, 1.811, 0.525, [{ text: "@archineastudio", options: { breakLine: true } }, { text: "Facebook, LinkedIn, Instagram ", options: { italic: true } }]);
  label(s, 10.356, 5.927, 1.732, 0.298, "FOLLOW US");
  heading(s, 0.84, 5.194, 4.93, 0.707, "Get In Touch With Us");
  kicker(s, 0.84, 4.722);
  credit(s, 2.125, 6.585);
  pageNum(s, 0.84, 6.585, "31");
  rule(s, 1.943, 6.506, 0, 0.561);
}

function slide32(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: 13.335, h: 7.5, fill: { color: WHITE, transparency: 30 } });
  text(s, "Thank You", { x: 0.84, y: 2.482, w: 6.082, h: 1.582, fontSize: 88, fontFace: DISPLAY, color: INK, wrap: false });
  text(s, [{ text: "© Archinea Studio", options: { breakLine: true } }, { text: "www.archineastudio.com " }], { x: 0.84, y: 6.296, w: 1.865, h: 0.525, fontSize: 9, fontFace: SANS, color: INK, lineSpacingMultiple: 1.5 });
  text(s, "2021", { x: 0.84, y: 5.959, w: 0.651, h: 0.337, fontSize: 14, fontFace: DISPLAY, color: INK, wrap: false });
  rule(s, 1.491, 6.127, 1.554, 0, INK);
  body(s, 0.84, 4.749, 5.591, 0.525, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi", { color: INK });
  text(s, "ARCHINEA STUDIO PRESENTATION TEMPLATE", { x: 0.84, y: 2.018, w: 3.268, h: 0.269, fontSize: 10, fontFace: SANS, color: INK, wrap: false });
}

// ----------------------------------------------------------------- render --
const SLIDES = [
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
  slide32
];

function build() {
  const pptx = new pptxgen();
  pptx.defineLayout({ name: 'ARCHINEA', width: 13.335, height: 7.5 });
  pptx.layout = 'ARCHINEA';
  pptx.author = 'Archinea Studio';
  pptx.title = 'Archinea Studio Presentation Template';
  SLIDES.forEach(fn => fn(pptx));
  return pptx.writeFile({ fileName: path.join(__dirname, OUTPUT_NAME) });
}

build().then(f => console.log('wrote ' + f));
