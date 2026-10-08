/*
 * Madog.Corp - "Managing Properties, (Delivering) Possibilities!"
 * A 25-slide, 16:9 project-management deck rebuilt with pptxgenjs.
 *
 *     node 16efaec1-b9d0-4da4-af28-fdd9f3d4cfdc_grok_final.js
 *
 * Layout is expressed in inches on a 13.333 x 7.5 canvas. Everything is drawn
 * with native pptxgenjs primitives; the raster artwork of the source deck is
 * stood in for by the flat `photoFrame` rectangles.
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette
const ORANGE = 'FF6D00'; // theme accent 1 - every highlight in the deck
const INK = '212121'; // theme dark 1 - body copy, rules, dots
const GREY = 'F2F2F2'; // light card fill (bg1 at 95% luminance)
const WHITE = 'FFFFFF';
const PHOTO = 'FAF4F0'; // stand-in tint for the deck's hatched image frames
const HAIRLINE = 0.01; // rules inherit the theme's thinnest stroke

const HEAD = 'Space Grotesk Medium'; // theme major font
const BODY = 'Inter Light'; // theme minor font

// Placeholder copy that the template repeats throughout the deck.
const LOREM1 =
  'The European languages are members of the same family. Their separate existence is a myth. For science, music, sport, etc.';
const LOREM2 =
  'The European languages are members of the same family. Their separate existence is a myth.';
const LOREM3 =
  'The European languages are members of the same family. Their separate existence is a myth. For science, music, sport, etc, Europe uses the same vocabulary.';
const LOREM4 =
  'The European languages are members of the same family. Their separate.';
const LOREM5 =
  'The European languages are members of the same family.';

// ------------------------------------------------------------------ text
// Every text box in the deck is top-anchored, keeps PowerPoint's implicit
// insets (0.1" / 0.05" - deliberately not restated, since renderers treat an
// explicit inset differently from an inherited one) and is set to "resize
// shape to fit text". `txt` centralises that so the builders below stay
// purely declarative; `extra` overrides any of it.
function txt(s, text, x, y, w, h, extra) {
  s.addText(text, Object.assign({
    x: x, y: y, w: w, h: h,
    fontFace: BODY, fontSize: 12, color: INK,
    align: 'left', valign: 'top', autoFit: true, wrap: false,
  }, extra));
}

// Headline words are set with wrapping off so the auto-fitted box hugs the
// glyphs. The few that must reflow pass `{ wrap: true }` (and, on the closing
// statement slides, `{ align: 'justify' }`) through `extra`.

// 48pt (or 36pt) headline word in the body face.
function title(s, text, x, y, w, h, size, extra) {
  txt(s, text, x, y, w, h, Object.assign({ fontSize: size || 48 }, extra));
}

// The bracketed orange word that punctuates every headline.
function accent(s, text, x, y, w, h, size, extra) {
  txt(s, text, x, y, w, h,
    Object.assign({ fontSize: size || 48, fontFace: HEAD, color: ORANGE }, extra));
}

// 16pt sub-head in the display face.
function heading(s, text, x, y, w, h, extra) {
  txt(s, text, x, y, w, h, Object.assign({ fontSize: 16, fontFace: HEAD }, extra));
}

// 16pt list entry in the body face.
function item(s, text, x, y, w, h) {
  txt(s, text, x, y, w, h, { fontSize: 16 });
}

// 12pt justified intro paragraph.
function lead(s, text, x, y, w, h) {
  txt(s, text, x, y, w, h, { fontSize: 12, align: 'justify', lineSpacing: 18, wrap: true });
}

// 10pt justified paragraph, tight leading (used inside the stat cards).
function note(s, text, x, y, w, h, color) {
  txt(s, text, x, y, w, h,
    { fontSize: 10, color: color || INK, align: 'justify', lineSpacing: 14, wrap: true });
}

// 10pt justified paragraph, open leading (used in the two-column bodies).
function para(s, text, x, y, w, h) {
  txt(s, text, x, y, w, h, { fontSize: 10, align: 'justify', lineSpacing: 18, wrap: true });
}

// ---------------------------------------------------------------- shapes
// Rounded rectangle; `r` is the corner radius in inches.
function panel(s, x, y, w, h, color, r) {
  s.addShape('roundRect', {
    x: x, y: y, w: w, h: h, rectRadius: r,
    fill: { color: color }, line: { type: 'none' },
  });
}

function dot(s, x, y, d) {
  s.addShape('ellipse', { x: x, y: y, w: d, h: d, fill: { color: INK }, line: { type: 'none' } });
}

function hRule(s, x, y, w) {
  s.addShape('line', { x: x, y: y, w: w, h: 0, line: { color: INK, width: HAIRLINE } });
}

function vRule(s, x, y, h) {
  s.addShape('line', { x: x, y: y, w: 0, h: h, line: { color: INK, width: HAIRLINE } });
}

// Draws a filled outline from sub-paths whose points are normalised to 0..1.
function vector(s, x, y, w, h, subpaths, color, extra) {
  const pts = [];
  subpaths.forEach(function (sub) {
    sub.forEach(function (p, i) { pts.push({ x: w * p[0], y: h * p[1], moveTo: i === 0 }); });
    pts.push({ close: true });
  });
  s.addShape('custGeom', Object.assign({
    x: x, y: y, w: w, h: h, points: pts,
    fill: { color: color }, line: { type: 'none' },
  }, extra));
}

// Square brackets around a diagonal arrow - the Madog.Corp monogram. Slides
// flip it horizontally / vertically to point the arrow into the layout.
const MONOGRAM = [
  [[0, 1], [0, 0.003], [0.145, 0.003], [0.145, 0.112],
   [0.07, 0.112], [0.07, 0.891], [0.145, 0.891], [0.145, 1]],
  [[1, 0.003], [1, 1], [0.855, 1], [0.855, 0.891],
   [0.93, 0.891], [0.93, 0.112], [0.855, 0.112], [0.855, 0.003]],
  [[0.754, 0.919], [0.343, 0.113], [0.756, 0.113], [0.756, 0], [0.245, 0],
   [0.245, 0.998], [0.302, 0.998], [0.302, 0.193], [0.714, 0.998]],
];

function monogram(s, o) {
  vector(s, o.x, o.y, o.w || 0.9531, o.h || 0.4869, MONOGRAM, ORANGE,
    { flipH: !!o.flipH, flipV: !!o.flipV });
}

// Eight-pointed star, the badge in the corner of every stat card.
const STAR = [[[1, 0.432], [0.665, 0.432], [0.902, 0.195], [0.805, 0.098], [0.569, 0.335],
  [0.569, 0], [0.432, 0], [0.432, 0.335], [0.195, 0.098], [0.098, 0.195], [0.335, 0.432],
  [0, 0.432], [0, 0.569], [0.335, 0.569], [0.098, 0.805], [0.195, 0.902], [0.432, 0.665],
  [0.432, 1], [0.569, 1], [0.569, 0.665], [0.805, 0.902], [0.902, 0.805], [0.665, 0.569],
  [1, 0.569]]];

function asterisk(s, x, y, size, color) {
  vector(s, x, y, size, size, STAR, color);
}

// Small hooked arrow that terminates the plain list rows; `rotate` aims it.
const ARROW = [[[1, 0.932], [0.164, 0.096], [0.966, 0.096], [0.966, 0], [0, 0], [0, 0.966],
  [0.096, 0.966], [0.096, 0.164], [0.932, 1]]];

function cornerArrow(s, x, y, rotate) {
  vector(s, x, y, 0.134, 0.134, ARROW, INK, { rotate: rotate });
}

// ----------------------------------------------------------- compositions
// Rounded panel + big figure + supporting sentence + star badge. Light by
// default; `dark: true` inverts it to the orange-on-white variant.
function statCard(s, c) {
  const fg = c.dark ? WHITE : INK;
  const pad = c.pad === undefined ? 0.123 : c.pad;
  const top = c.valueTop === undefined ? c.h - 1.279 : c.valueTop;
  panel(s, c.x, c.y, c.w, c.h, c.dark ? ORANGE : GREY, c.r);
  txt(s, c.value, c.x + pad, c.y + top, c.valW, 0.64,
    { fontSize: 32, fontFace: HEAD, color: fg });
  note(s, c.caption, c.x + pad - 0.003, c.y + top + 0.64, c.capW, 0.478, fg);
  asterisk(s, c.x + c.w - (c.starRight === undefined ? 0.455 : c.starRight),
    c.y + (c.starTop === undefined ? 0.242 : c.starTop), 0.189, c.starColor || fg);
}

// Stand-in for a photo: flat tinted panel carrying the template's own prompt.
function photoFrame(s, x, y, w, h, r) {
  panel(s, x, y, w, h, PHOTO, r);
  txt(s, 'Put Your Image Here!', x, y + 0.06, w, 0.32, { fontSize: 12, align: 'center' });
  txt(s, '[image]', x, y + h / 2 - 0.16, w, 0.32,
    { fontSize: 11, align: 'center', color: 'B9A99A' });
}

// Running head, repeated verbatim on all 25 slides.
function header(s) {
  txt(s, 'Madog.Corp', 0.5676, 0.6769, 1.0119, 0.2692, { fontSize: 10 });
  txt(s, 'Project Management', 2.5628, 0.6769, 1.5657, 0.2692, { fontSize: 10 });
  txt(s, [{ text: 'January 18' }, { text: 'th', options: { superscript: true } }, { text: ', 2030' }],
    11.3401, 0.6769, 1.4262, 0.2692, { fontSize: 10, align: 'right' });
}

// --------------------------------------------------------------- slides
function slide1(s) {
  header(s);
  statCard(s, { x: 10.3556, y: 4.8439, w: 2.3111, h: 1.9061, r: 0.0708, value: '85%', valW: 1.0873, caption: 'Exceed Their Initial Budget.', capW: 1.3846, pad: 0.2077, starRight: 0.397, starTop: 0.1551 });
  statCard(s, { x: 7.7505, y: 4.8439, w: 2.3111, h: 1.9061, r: 0.0708, value: '90%', valW: 1.11, caption: 'Clients Expect Detailed Documentation.', capW: 2.0174, pad: 0.1502, starTop: 0.1551 });
  statCard(s, { x: 5.2311, y: 4.8439, w: 2.3111, h: 1.9061, r: 0.0708, value: '80%', valW: 1.11, caption: 'Successful Projects Meet Scope.', capW: 1.6768, pad: 0.1503, starTop: 0.1551 });
  monogram(s, { x: 11.7135, y: 1.7654 });
  photoFrame(s, 0.6667, 4.8438, 4.3562, 1.9061, 0.0455);
  title(s, 'Managing Properties,', 0.5424, 1.5589, 6.9898, 0.9088);
  accent(s, '(Delivering)', 1.9299, 2.5703, 3.9412, 0.9088);
  title(s, 'Possibilities!', 6.0181, 2.5674, 4.1008, 0.9088);
  lead(s, LOREM2, 1.9299, 3.6482, 5.3703, 0.5831);
}

function slide2(s) {
  header(s);
  statCard(s, { x: 0.7908, y: 4.4352, w: 2.3111, h: 2.3148, r: 0.0858, value: '85%', valW: 1.0873, caption: 'Real Estate Projects Exceed.', capW: 1.554, pad: 0.1502, starTop: 0.3079 });
  statCard(s, { x: 3.2488, y: 4.4352, w: 2.3111, h: 2.3148, r: 0.0858, value: '60%', valW: 1.11, caption: 'Projects Face Unexpected Cost Increases.', capW: 1.9563, dark: true, pad: 0.1503, starTop: 0.3079 });
  monogram(s, { x: 0.7988, y: 1.9504, flipH: true, flipV: true });
  hRule(s, 6.3658, 5.5686, 6.0712);
  title(s, 'Budgeting', 3.0908, 1.7498, 3.4293, 0.9088);
  lead(s, LOREM2, 3.1553, 2.8409, 5.3703, 0.5831);
  accent(s, '(Strategies)', 6.5481, 1.7749, 4.1411, 0.9088);
  heading(s, 'Budgeting Plan.', 6.2981, 4.3192, 1.9322, 0.3703);
  note(s, LOREM2, 6.2981, 4.6929, 3.9935, 0.4782);
  heading(s, 'Accurate Cost Estimation.', 8.549, 5.971, 3.0472, 0.3703);
  note(s, LOREM2, 8.549, 6.3448, 3.9935, 0.4782);
}

function slide3(s) {
  header(s);
  statCard(s, { x: 7.7336, y: 1.6698, w: 2.3013, h: 1.775, r: 0.0635, value: '65%', valW: 1.089, caption: 'Risk Management Is The Most Critical Part.', capW: 1.9563, pad: 0.1117, valueTop: 0.5131, starRight: 0.3297, starTop: 0.145 });
  statCard(s, { x: 10.2419, y: 1.6698, w: 2.3013, h: 1.775, r: 0.0635, value: '70+', valW: 1.0206, caption: 'Project Managed By John Doe Since 2025.', capW: 1.7786, pad: 0.1117, valueTop: 0.5131, starRight: 0.3297, starTop: 0.145 });
  monogram(s, { x: 7.7336, y: 6.1127, flipH: true });
  photoFrame(s, 10.2419, 3.6325, 2.3011, 2.9666, 0.0715);
  hRule(s, 0.9246, 4.7554, 4.5444);
  hRule(s, 1.6808, 5.4746, 4.5444);
  hRule(s, 2.664, 6.1797, 4.5444);
  accent(s, '(Project Manager)', 0.7901, 1.7604, 6.1133, 0.9088);
  txt(s, 'John Doe Experience', 0.841, 4.2809, 2.2057, 0.3366, { fontSize: 14, fontFace: HEAD });
  lead(s, LOREM5, 0.8234, 4.9497, 4.7357, 0.3306);
  lead(s, LOREM5, 1.5796, 5.6689, 4.7357, 0.3306);
  lead(s, LOREM5, 2.5628, 6.374, 4.7357, 0.3306);
  lead(s, LOREM2, 0.8234, 2.8887, 5.3703, 0.5831);
}

function slide4(s) {
  header(s);
  statCard(s, { x: 1.2556, y: 2.8986, w: 2.3111, h: 1.4833, r: 0.0551, value: '2025', valW: 1.2959, caption: 'Land Acquisition & Feasibility Study.', capW: 1.554, pad: 0.1502, valueTop: 0.1828, starTop: 0.3331, starColor: 'FF6D00' });
  statCard(s, { x: 4.0926, y: 2.8986, w: 2.3111, h: 1.4833, r: 0.0551, value: '2026', valW: 1.3029, caption: 'Design, Permits & Pre-Construction.', capW: 1.554, pad: 0.1503, valueTop: 0.1828, starTop: 0.3331, starColor: 'FF6D00' });
  statCard(s, { x: 6.9296, y: 2.8986, w: 2.3111, h: 1.4833, r: 0.0551, value: '2027', valW: 1.2783, caption: 'Construction & Quality Control.', capW: 1.554, pad: 0.1502, valueTop: 0.1828, starTop: 0.3331, starColor: 'FF6D00' });
  statCard(s, { x: 9.7667, y: 2.8986, w: 2.3111, h: 1.4833, r: 0.0551, value: '2028', valW: 1.3011, caption: 'Project Completion & Handover.', capW: 1.554, pad: 0.1502, valueTop: 0.1828, starTop: 0.3331, starColor: 'FF6D00' });
  monogram(s, { x: 11.0997, y: 5.277, flipH: true, flipV: true });
  hRule(s, 2.5028, 2.0654, 8.3278);
  dot(s, 5.1618, 1.9738, 0.1833);
  dot(s, 2.3194, 1.9738, 0.1833);
  dot(s, 7.9882, 1.9738, 0.1833);
  dot(s, 10.8306, 1.9738, 0.1833);
  title(s, 'Project', 1.1311, 5.0659, 2.4248, 0.9088);
  accent(s, '(Timeline)', 3.6502, 5.091, 3.4083, 0.9088);
  lead(s, LOREM3, 1.1477, 6.2399, 7.0238, 0.5831);
}

function slide5(s) {
  header(s);
  statCard(s, { x: 10.235, y: 4.1518, w: 2.3111, h: 2.3148, r: 0.0858, value: '55%', valW: 1.082, caption: 'Transactions Are Expected To Use Blockchain.', capW: 1.9807, dark: true, pad: 0.1503, starTop: 0.3079 });
  monogram(s, { x: 11.5929, y: 1.9887, flipV: true });
  asterisk(s, 0.9196, 4.3372, 0.16, 'FF6D00');
  asterisk(s, 0.9196, 6.3019, 0.16, 'FF6D00');
  asterisk(s, 5.3304, 4.3372, 0.16, 'FF6D00');
  asterisk(s, 5.3304, 6.3019, 0.16, 'FF6D00');
  hRule(s, 1.2578, 4.1879, 2.8173);
  hRule(s, 1.2578, 6.1526, 3.088);
  hRule(s, 5.6686, 4.1879, 2.8173);
  hRule(s, 5.6686, 6.1526, 2.8173);
  heading(s, 'AI-Powered Management.', 1.1559, 4.2462, 3.0261, 0.3703);
  title(s, 'Technology', 0.7873, 1.7777, 3.8992, 0.9088);
  accent(s, '(Trends)', 4.6518, 1.8028, 2.97, 0.9088);
  note(s, LOREM4, 1.1337, 3.4671, 2.982, 0.4782);
  heading(s, 'Blockchain for Transactions.', 1.1559, 6.211, 3.3277, 0.3703);
  note(s, LOREM4, 1.1337, 5.4318, 2.982, 0.4782);
  heading(s, 'Digital Twins & BIM.', 5.5667, 4.2462, 2.2969, 0.3703);
  note(s, LOREM4, 5.5445, 3.4671, 2.982, 0.4782);
  heading(s, 'Smart Construction Sites.', 5.5667, 6.211, 3.0472, 0.3703);
  note(s, LOREM4, 5.5445, 5.4318, 2.982, 0.4782);
}

function slide6(s) {
  header(s);
  statCard(s, { x: 1.0954, y: 4.5532, w: 2.3111, h: 2.0069, r: 0.0579, value: '35%', valW: 1.0855, caption: 'Projects Fail Due To Poor Risk Management.', capW: 1.9807, dark: true, pad: 0.1502 });
  statCard(s, { x: 3.8294, y: 4.5532, w: 2.3111, h: 2.0069, r: 0.0579, value: '45%', valW: 1.096, caption: 'Delays Result From Legal And Compliance Issues.', capW: 1.9807, pad: 0.1502 });
  cornerArrow(s, 8.0289, 4.9425, 180);
  cornerArrow(s, 8.0289, 5.612, 180);
  cornerArrow(s, 8.0289, 6.2732, 180);
  monogram(s, { x: 11.4431, y: 2.0027, flipV: true });
  hRule(s, 8.5268, 5.3373, 3.8694);
  hRule(s, 8.5268, 5.9873, 3.8694);
  title(s, 'Risk', 0.9371, 1.7916, 1.4817, 0.9088);
  accent(s, '(Management)', 2.4761, 1.8167, 5.0176, 0.9088);
  title(s, 'Strategies', 7.5511, 1.7916, 3.4136, 0.9088);
  lead(s, LOREM1, 2.7659, 2.9979, 5.9406, 0.5831);
  item(s, 'Identifying and Mitigating Risks.', 8.4146, 4.8219, 3.5363, 0.3703);
  item(s, 'Financial Risk Management.', 8.4187, 5.4895, 3.1103, 0.3703);
  item(s, 'Operational Risk Management.', 8.4146, 6.157, 3.4013, 0.3703);
}

function slide7(s) {
  header(s);
  asterisk(s, 4.8792, 6.2548, 0.16, 'FF6D00');
  asterisk(s, 9.1864, 6.2548, 0.16, 'FF6D00');
  monogram(s, { x: 11.3803, y: 1.8136, flipV: true });
  photoFrame(s, 0.9506, 1.8137, 3.1562, 4.5977, 0.0398);
  hRule(s, 5.2174, 6.1055, 2.8173);
  hRule(s, 9.5246, 6.1055, 2.8173);
  title(s, 'Managing -', 4.7249, 1.6672, 3.787, 0.9088);
  accent(s, '(Stakeholders)', 5.6715, 2.6705, 5.0316, 0.9088);
  heading(s, 'Transparent Communication.', 5.1155, 6.1638, 3.4223, 0.3703);
  note(s, LOREM4, 5.0933, 5.3847, 2.982, 0.4782);
  heading(s, 'Conflict Resolution.', 9.4228, 6.1638, 2.3565, 0.3703);
  note(s, LOREM4, 9.4005, 5.3847, 2.982, 0.4782);
  lead(s, LOREM1, 5.1155, 4.0235, 5.9406, 0.5831);
}

function slide8(s) {
  header(s);
  statCard(s, { x: 7.8976, y: 1.5388, w: 2.3111, h: 2.0069, r: 0.0579, value: '30%', valW: 1.1065, caption: 'Project Delays Are Due To Rising Material.', capW: 1.9807, dark: true, pad: 0.1503 });
  statCard(s, { x: 10.3556, y: 1.5388, w: 2.3111, h: 2.0069, r: 0.0579, value: '40%', valW: 1.1171, caption: 'Funding Shortfalls Due To Rising Interest Rates.', capW: 1.8985, pad: 0.1502 });
  monogram(s, { x: 11.3803, y: 4.3219, flipV: true });
  photoFrame(s, 0.6667, 1.5387, 7.0836, 2.0069, 0.0642);
  hRule(s, 0.6705, 6.3664, 4.7611);
  hRule(s, 7.7962, 6.3664, 4.7611);
  title(s, 'Impact Of', 0.5094, 4.1494, 3.2856, 0.9088);
  accent(s, '(Inflation)', 3.8598, 4.1744, 3.4714, 0.9088);
  heading(s, 'Rising Construction Costs.', 0.5524, 6.4939, 3.0963, 0.3703);
  para(s, LOREM1, 0.5674, 5.6123, 4.9698, 0.5763);
  heading(s, 'Higher Interest Rates.', 7.6781, 6.4939, 2.5809, 0.3703);
  para(s, LOREM1, 7.693, 5.6123, 4.9698, 0.5763);
}

function slide9(s) {
  header(s);
  statCard(s, { x: 7.6852, y: 4.6372, w: 2.3111, h: 2.0069, r: 0.0579, value: '45%', valW: 1.096, caption: 'Projects Face Legal Disputes Due To Non-compliance.', capW: 2.0947 });
  statCard(s, { x: 10.1922, y: 4.6372, w: 2.4745, h: 2.0069, r: 0.0579, value: '30%', valW: 1.1065, caption: 'Roper Zoning Compliance Reduces Project Approval Delays.', capW: 2.4059, dark: true, pad: 0.0377, starRight: 0.3845 });
  cornerArrow(s, 5.0686, 6.248, 270);
  cornerArrow(s, 5.0686, 5.5993, 270);
  cornerArrow(s, 5.0686, 4.9493, 270);
  photoFrame(s, 7.6847, 2.035, 4.982, 2.39, 0.0764);
  hRule(s, 0.744, 5.3121, 4.4581);
  hRule(s, 0.744, 5.9621, 4.4581);
  title(s, 'Compliance', 0.5901, 1.8711, 3.9114, 0.9088);
  accent(s, '(Ensure)', 4.468, 1.8961, 2.9332, 0.9088);
  lead(s, LOREM1, 0.6318, 2.9715, 5.9406, 0.5831);
  item(s, '2026: Secure Permits.', 0.6318, 4.7967, 2.5248, 0.3703);
  item(s, '2027: Conduct Site Assessments.', 0.6359, 5.4643, 3.7414, 0.3703);
  item(s, '2028: Ensure Proper Documentation.', 0.6318, 6.1318, 4.0797, 0.3703);
}

function slide10(s) {
  header(s);
  statCard(s, { x: 10.3556, y: 2.3555, w: 2.3111, h: 2.0069, r: 0.0579, value: '70%', valW: 1.0855, caption: 'Projects Use Traditional Project Management.', capW: 1.8712, dark: true });
  monogram(s, { x: 11.7558, y: 5.1878, flipV: true });
  vRule(s, 3.1856, 1.7337, 0.4475);
  vRule(s, 3.1856, 2.8243, 0.4475);
  vRule(s, 3.1856, 3.9149, 0.4475);
  title(s, 'Traditional', 0.5358, 4.9797, 3.5293, 0.9088);
  accent(s, '(Project)', 4.0922, 5.0076, 3.0752, 0.9088);
  title(s, 'Management', 7.1946, 4.9797, 4.3322, 0.9088);
  heading(s, 'Feasibility Study & Approvals.', 0.5524, 1.6348, 2.2542, 0.6395, { wrap: true });
  para(s, LOREM1, 3.6155, 1.6943, 4.9698, 0.5763);
  heading(s, 'Budgeting & Resource Allocation.', 0.5524, 2.7255, 2.4365, 0.6395, { wrap: true });
  para(s, LOREM1, 3.6155, 2.7849, 4.9698, 0.5763);
  heading(s, 'Construction & Final Handover.', 0.5524, 3.8161, 2.2542, 0.6395, { wrap: true });
  para(s, LOREM1, 3.6155, 3.8755, 4.9698, 0.5763);
  lead(s, LOREM1, 0.5707, 6.2476, 5.9406, 0.5831);
}

function slide11(s) {
  header(s);
  statCard(s, { x: 0.675, y: 4.7431, w: 2.3111, h: 2.0069, r: 0.0579, value: '65%', valW: 1.089, caption: 'Failed Projects Lacked A Proper Feasibility Study.', capW: 1.8712, dark: true });
  vRule(s, 7.1459, 4.8812, 0.4475);
  vRule(s, 7.1459, 6.2053, 0.4475);
  hRule(s, 3.82, 5.7592, 8.8383);
  title(s, 'Importance Of', 0.5108, 1.8799, 4.7056, 0.9088);
  accent(s, '(Feasibility)', 5.2711, 1.9075, 4.0587, 0.9088);
  title(s, 'Studies', 9.3846, 1.8797, 2.5616, 0.9088);
  heading(s, 'Why is a Feasibility Study Important?', 3.7299, 4.7824, 2.2542, 0.6395, { wrap: true });
  para(s, LOREM1, 7.7958, 4.8418, 4.9698, 0.5763);
  heading(s, 'Common Risks Without a Feasibility Study.', 3.7299, 6.1065, 2.7701, 0.6395, { wrap: true });
  para(s, LOREM1, 7.7958, 6.1659, 4.9698, 0.5763);
  lead(s, LOREM1, 0.5594, 3.1279, 5.9406, 0.5831);
}

function slide12(s) {
  header(s);
  statCard(s, { x: 10.3556, y: 4.4653, w: 2.3111, h: 2.0069, r: 0.0579, value: '35%', valW: 1.0855, caption: 'Pricing Models Improve Property Valuation.', capW: 1.8712, dark: true });
  statCard(s, { x: 10.3556, y: 2.1137, w: 2.3111, h: 2.0069, r: 0.0579, value: '70%', valW: 1.0855, caption: 'Invest In Ai-powered Project Management Tools.', capW: 2.0328 });
  monogram(s, { x: 0.6667, y: 5.9855, flipH: true });
  hRule(s, 3.1493, 4.3136, 4.7611);
  hRule(s, 3.1343, 6.0873, 4.7611);
  title(s, 'Role Of', 0.5108, 1.9112, 2.5055, 0.9088);
  accent(s, '(AI)', 3.0162, 1.9388, 1.3187, 0.9088);
  title(s, 'In Real Estate', 4.3349, 1.9111, 4.4864, 0.9088);
  heading(s, 'Project Management with AI.', 3.0312, 4.4411, 3.3364, 0.3703);
  para(s, LOREM1, 3.0461, 3.5595, 4.9698, 0.5763);
  heading(s, 'Enhancing Efficiency & Sustainability.', 3.0162, 6.2149, 4.3409, 0.3703);
  para(s, LOREM1, 3.0312, 5.3332, 4.9698, 0.5763);
}

function slide13(s) {
  header(s);
  statCard(s, { x: 0.6643, y: 2.009, w: 2.3111, h: 2.0069, r: 0.0579, value: '10%', valW: 1.0294, caption: 'Sustainable Projects Increase Property Value.', capW: 1.8712, dark: true });
  vRule(s, 7.1567, 2.2346, 0.4475);
  hRule(s, 3.8308, 3.1126, 8.8383);
  vRule(s, 7.1567, 3.5848, 0.4475);
  hRule(s, 7.899, 6.3665, 4.7611);
  title(s, 'Sustainable Project', 0.5215, 5.1881, 6.4499, 0.9088);
  accent(s, '(Management)', 0.5461, 6.0835, 5.0176, 0.9088);
  heading(s, 'Eco-Friendly Building Materials.', 3.7406, 2.1358, 2.2542, 0.6395, { wrap: true });
  para(s, LOREM1, 7.8066, 2.1953, 4.9698, 0.5763);
  heading(s, 'Energy-Efficient Design.', 3.7406, 3.4859, 2.2542, 0.6395, { wrap: true });
  para(s, LOREM1, 7.8066, 3.5454, 4.9698, 0.5763);
  heading(s, 'Impact of Sustainability.', 7.7809, 6.494, 2.8876, 0.3703);
  para(s, LOREM1, 7.7958, 5.6124, 4.9698, 0.5763);
}

function slide14(s) {
  header(s);
  statCard(s, { x: 7.4952, y: 4.3982, w: 2.3111, h: 2.0069, r: 0.0579, value: '45%', valW: 1.096, caption: 'Projects Face Legal Disputes Due To Non-compliance.', capW: 2.0947 });
  statCard(s, { x: 10.0022, y: 4.3982, w: 2.4745, h: 2.0069, r: 0.0579, value: '30%', valW: 1.1065, caption: 'Roper Zoning Compliance Reduces Project Approval Delays.', capW: 2.4059, dark: true, pad: 0.0377, starRight: 0.3845 });
  monogram(s, { x: 0.8313, y: 5.6707, flipH: true });
  hRule(s, 3.0542, 5.7798, 3.3067);
  hRule(s, 7.7466, 3.0347, 3.3067);
  title(s, 'Common', 0.692, 1.836, 3.0314, 0.9088);
  accent(s, '(Cause)', 3.736, 1.8612, 2.7036, 0.9088);
  lead(s, LOREM1, 0.7531, 2.9387, 5.6866, 0.5831);
  heading(s, 'Poor Planning & Scheduling.', 2.9361, 5.9074, 3.2313, 0.3703);
  para(s, LOREM2, 2.9511, 5.0257, 3.5099, 0.5763);
  heading(s, 'Regulatory & Permit Issues.', 7.6285, 3.1623, 3.2313, 0.3703);
  para(s, LOREM2, 7.6435, 2.2806, 3.5099, 0.5763);
}

function slide15(s) {
  header(s);
  statCard(s, { x: 0.9129, y: 1.9158, w: 2.3111, h: 2.0069, r: 0.0579, value: '45%', valW: 1.096, caption: 'Project Delays Are Caused By Vendor-related Issues.', capW: 2.0947, dark: true });
  monogram(s, { x: 11.3673, y: 1.9158, flipV: true });
  photoFrame(s, 0.9128, 4.131, 2.3111, 2.3078, 0.0671);
  hRule(s, 4.4951, 3.8492, 3.3067);
  hRule(s, 9.0137, 3.8492, 3.3067);
  title(s, 'Vendor', 4.3749, 4.8469, 2.4774, 0.9088);
  accent(s, '(Management)', 6.9076, 4.8721, 5.0176, 0.9088);
  lead(s, LOREM1, 4.436, 5.9496, 5.6866, 0.5831);
  heading(s, 'Vendor Selection & Evaluation.', 4.377, 3.9767, 3.5223, 0.3703);
  para(s, LOREM2, 4.3919, 3.0951, 3.5099, 0.5763);
  heading(s, 'Contract Negotiation.', 8.8956, 3.9767, 2.5914, 0.3703);
  para(s, LOREM2, 8.9106, 3.0951, 3.5099, 0.5763);
}

function slide16(s) {
  header(s);
  statCard(s, { x: 10.3556, y: 4.4729, w: 2.3111, h: 2.0069, r: 0.0579, value: '80%', valW: 1.1083, caption: 'Legal Disputes Arise Due To Improper Documentation.', capW: 2.0947, dark: true });
  photoFrame(s, 6.5668, 4.4728, 3.558, 2.0069, 0.0584);
  photoFrame(s, 2.0039, 4.4728, 4.2965, 2.0069, 0.0584);
  panel(s, 2.0042, 3.1068, 0.05, 0.3729, 'FF6D00', 0.0083);
  title(s, 'Real Estate', 0.5094, 1.7853, 3.7309, 0.9088);
  accent(s, '(Project)', 4.3294, 1.8105, 3.0752, 0.9088);
  lead(s, LOREM3, 2.1296, 2.9851, 8.0426, 0.5831);
  title(s, 'Documentation', 7.4938, 1.7853, 5.0071, 0.9088);
}

function slide17(s) {
  header(s);
  cornerArrow(s, 12.5331, 6.6022, 270);
  cornerArrow(s, 12.5331, 5.9535, 270);
  cornerArrow(s, 12.5331, 5.3035, 270);
  monogram(s, { x: 0.6667, y: 6.2632, flipH: true });
  hRule(s, 8.2086, 5.6663, 4.4581);
  hRule(s, 8.2086, 6.3163, 4.4581);
  hRule(s, 0.6667, 2.4928, 3.3067);
  hRule(s, 9.3589, 2.4928, 3.3067);
  hRule(s, 5.0223, 2.4928, 3.3067);
  title(s, 'Construction', 0.5094, 3.5921, 4.2568, 0.9088);
  accent(s, '(Safety)', 4.7179, 3.6173, 2.9315, 0.9088);
  title(s, 'Management', 7.6494, 3.5921, 4.3322, 0.9088);
  item(s, 'Personal Protective Equipment.', 8.0963, 5.1509, 3.4855, 0.3703);
  item(s, 'Regular Safety Training.', 8.1005, 5.8185, 2.7053, 0.3703);
  item(s, 'Site Inspections & Audits.', 8.0963, 6.486, 2.8631, 0.3703);
  heading(s, 'Personal Protective Equipment.', 0.5486, 2.6203, 3.6362, 0.3703);
  para(s, LOREM2, 0.5635, 1.7387, 3.5099, 0.5763);
  heading(s, 'Site Inspections & Audits.', 9.2408, 2.6203, 2.9823, 0.3703);
  para(s, LOREM2, 9.2558, 1.7387, 3.5099, 0.5763);
  heading(s, 'Regular Safety Training.', 4.9043, 2.6203, 2.8333, 0.3703);
  para(s, LOREM2, 4.9192, 1.7387, 3.5099, 0.5763);
}

function slide18(s) {
  header(s);
  statCard(s, { x: 10.3556, y: 1.8466, w: 2.3111, h: 2.0069, r: 0.0579, value: '85%', valW: 1.0873, caption: 'Project Failures Are Due To Poor Communication.', capW: 1.8712, dark: true });
  monogram(s, { x: 11.716, y: 5.137, flipV: true });
  vRule(s, 3.9716, 2.0534, 0.4475);
  hRule(s, 0.6457, 2.9314, 8.8383);
  vRule(s, 3.9716, 3.4035, 0.4475);
  panel(s, 0.6667, 6.171, 0.05, 0.3729, 'FF6D00', 0.0083);
  title(s, 'Communication', 0.5201, 4.8942, 5.114, 0.9088);
  accent(s, '(Project)', 5.6342, 4.9193, 3.0752, 0.9088);
  heading(s, 'Use Centralized Communication Platforms.', 0.5556, 1.9546, 3.1503, 0.6395, { wrap: true });
  para(s, LOREM1, 4.6215, 2.014, 4.9698, 0.5763);
  heading(s, 'Regular Progress Meetings & Reports.', 0.5556, 3.3047, 2.5222, 0.6395, { wrap: true });
  para(s, LOREM1, 4.6215, 3.3641, 4.9698, 0.5763);
  lead(s, LOREM3, 0.7921, 6.0493, 8.0426, 0.5831);
}

function slide19(s) {
  header(s);
  statCard(s, { x: 9.9833, y: 4.6463, w: 2.6833, h: 2.0069, r: 0.0579, value: '50%', valW: 1.103, caption: 'Report Financial Losses Due To Unplanned Project Modifications.', capW: 2.3539, dark: true, starRight: 0.3982 });
  statCard(s, { x: 7.5095, y: 4.6463, w: 2.3111, h: 2.0069, r: 0.0579, value: '70%', valW: 1.0855, caption: 'Projects Face Scope Changes During Execution.', capW: 2.0947 });
  monogram(s, { x: 11.7135, y: 1.9999 });
  hRule(s, 2.8876, 6.0174, 3.3067);
  title(s, 'Management', 0.5093, 1.7749, 4.3322, 0.9088);
  accent(s, '(Change)', 4.8415, 1.8, 3.1944, 0.9088);
  heading(s, 'Why is Change Management Essential?', 0.5652, 3.2294, 4.4566, 0.3703);
  para(s, LOREM3, 0.5691, 3.7366, 5.7252, 0.5763);
  heading(s, 'Market Fluctuations.', 2.7695, 6.1449, 2.4722, 0.3703);
  para(s, LOREM2, 2.7844, 5.2633, 3.5099, 0.5763);
}

function slide20(s) {
  header(s);
  statCard(s, { x: 10.3556, y: 1.5658, w: 2.3111, h: 2.0069, r: 0.0579, value: '90%', valW: 1.11, caption: 'Lean Projects Meet Or Exceed Delivery Deadlines.', capW: 2.0947, dark: true });
  statCard(s, { x: 7.7778, y: 1.5658, w: 2.432, h: 2.0069, r: 0.0579, value: '30%', valW: 1.1065, caption: 'Construction Materials Go To Waste In Traditional Projects.', capW: 2.2229, pad: 0.1079 });
  title(s, 'Benefit Of', 0.5093, 4.094, 3.382, 0.9088);
  accent(s, '(Lean Construction)', 3.8913, 4.114, 6.6976, 0.9088, 48, { wrap: true });
  heading(s, 'How Lean Construction Lowers Costs?', 0.5652, 2.3283, 4.4566, 0.3703);
  para(s, LOREM3, 0.5691, 2.8356, 5.7252, 0.5763);
  heading(s, 'Minimized Material Waste.', 8.0147, 5.5779, 3.0472, 0.3703);
  para(s, LOREM2, 8.0186, 6.0851, 3.6689, 0.5763);
  heading(s, 'Efficient Labor Utilization.', 0.5652, 5.5779, 3.0822, 0.3703);
  para(s, LOREM3, 0.5691, 6.0851, 5.7252, 0.5763);
}

function slide21(s) {
  header(s);
  statCard(s, { x: 10.3556, y: 4.3282, w: 2.3111, h: 2.0069, r: 0.0579, value: '80%', valW: 1.1083, caption: 'Firms Using Cloud-based Tools Report.', capW: 2.0947, dark: true });
  vRule(s, 3.9716, 4.5375, 0.4475);
  hRule(s, 0.6457, 5.4155, 8.8383);
  vRule(s, 3.9716, 5.8876, 0.4475);
  title(s, 'Automating', 0.5493, 1.7598, 3.8168, 0.9088);
  accent(s, '(Project Management)', 4.3661, 1.7798, 7.484, 0.9088, 48, { wrap: true });
  para(s, 'The European languages are members of the same family. Their separate existence is a myth. For science, music, sport, etc, Europe uses the same vocabulary. The languages only differ in their grammar, their pronunciation and their most common words.', 0.5557, 2.9293, 11.1776, 0.5763);
  heading(s, 'Benefits of Automation in Real Estate Project.', 0.5556, 4.4387, 3.1503, 0.6395, { wrap: true });
  para(s, LOREM1, 4.6215, 4.4981, 4.9698, 0.5763);
  heading(s, 'Key Technologies Driving Automation.', 0.5556, 5.7888, 2.5222, 0.6395, { wrap: true });
  para(s, LOREM1, 4.6215, 5.8482, 4.9698, 0.5763);
}

function slide22(s) {
  header(s);
  statCard(s, { x: 10.1722, y: 1.7431, w: 2.4944, h: 2.0069, r: 0.0579, value: '80%', valW: 1.1083, caption: 'Disputes Are Due To Contractual Misunderstandings.', capW: 2.2888, pad: 0.1063 });
  monogram(s, { x: 11.7135, y: 4.5863, flipV: true });
  photoFrame(s, 0.6667, 1.7431, 9.3427, 2.0069, 0.0501);
  title(s, 'Disputes', 0.5063, 4.383, 2.9543, 0.9088);
  accent(s, '(Handle)', 3.4606, 4.403, 2.9543, 0.9088, 48, { wrap: true });
  title(s, 'Projects', 6.4148, 4.383, 2.7702, 0.9088);
  heading(s, 'Contract Misinterpretation.', 0.5652, 5.7602, 3.2242, 0.3703);
  para(s, LOREM3, 0.5691, 6.2674, 5.7252, 0.5763);
  heading(s, 'Quality & Compliance Issues.', 7.0352, 5.7602, 3.2961, 0.3703);
  para(s, LOREM3, 7.039, 6.2674, 5.7252, 0.5763);
}

function slide23(s) {
  header(s);
  statCard(s, { x: 0.8222, y: 4.4932, w: 3.2889, h: 2.0069, r: 0.0579, value: '5D BIM', valW: 1.6868, caption: 'Budgeting Experience 25% Fewer Financial Discrepancies.', capW: 2.2888, dark: true, pad: 0.1789, valueTop: 0.7069, starRight: 0.4266 });
  monogram(s, { x: 0.8222, y: 2.1582, flipH: true, flipV: true });
  panel(s, 4.2842, 4.4932, 8.2269, 2.0069, 'F2F2F2', 0.0579);
  title(s, 'Role Of', 4.2842, 1.9605, 2.5055, 0.9088, 48, { wrap: true });
  accent(s, '(BIM)', 6.7897, 1.9805, 1.9192, 0.9088, 48, { wrap: true });
  title(s, 'In Project', 8.7089, 1.9605, 3.1804, 0.9088, 48, { wrap: true });
  para(s, LOREM3, 4.3287, 3.0959, 5.7252, 0.5763);
  heading(s, '3D Visualization & Planning.', 4.6683, 4.9603, 3.1787, 0.3703);
  para(s, LOREM2, 4.6721, 5.4568, 3.8034, 0.5763);
  heading(s, 'Enhanced Collaboration.', 9.2589, 4.9603, 2.8666, 0.3703);
  para(s, 'The European languages are members of the same family. Their.', 9.2627, 5.4568, 2.609, 0.5763);
}

function slide24(s) {
  header(s);
  monogram(s, { x: 10.3944, y: 2.7299, w: 0.6935, h: 0.3542, flipV: true });
  photoFrame(s, 7.557, 5.6197, 1.1801, 1.1303, 0.0476);
  title(s, 'A Well-managed Project Is Not Just About Meeting Deadlines, it\'s About', 0.5677, 1.7572, 12.1979, 1.4981, 36, { wrap: true, align: 'justify', lineSpacing: 52 });
  accent(s, '(Delivering Value)', 5.5712, 2.5651, 4.5325, 0.7068, 36, { wrap: true });
  title(s, 'Minimizing Risks, And Ensuring Lasting Success.', 0.5677, 3.283, 11.4968, 0.7068, 36, { wrap: true, align: 'justify' });
  txt(s, 'Morgan Davis – 2K27', 0.5955, 4.3094, 2.439, 0.3703, { fontSize: 16, fontFace: HEAD, color: 'FF6D00', wrap: true });
  heading(s, 'Funfact about Morgan Davis.', 8.9584, 5.7501, 3.3645, 0.3703);
  para(s, LOREM2, 8.9622, 6.2466, 3.8034, 0.5763);
}

function slide25(s) {
  header(s);
  statCard(s, { x: 10.4177, y: 1.7386, w: 2.3111, h: 1.9061, r: 0.0708, value: '800+', valW: 1.3327, caption: 'Projects Done Completely!', capW: 1.3846, dark: true, pad: 0.2077, starRight: 0.397, starTop: 0.1551, starColor: 'F2F2F2' });
  monogram(s, { x: 4.7211, y: 4.8055, flipV: true });
  photoFrame(s, 10.418, 3.8551, 2.3111, 2.447, 0.0891);
  hRule(s, 0.7288, 3.2726, 2.261);
  hRule(s, 3.978, 3.2726, 2.261);
  hRule(s, 7.2273, 3.2726, 2.261);
  title(s, 'Thank You', 0.6045, 4.599, 3.5521, 0.9088);
  accent(s, '(For All Your)', 0.6298, 5.6312, 4.3742, 0.9088);
  title(s, 'Attention', 5.0491, 5.6075, 3.0963, 0.9088);
  heading(s, 'Business Number', 0.6107, 3.3918, 2.0707, 0.3703);
  lead(s, '(+00) 0000 0000 0000 0000', 0.6257, 2.8061, 2.4572, 0.3306);
  heading(s, 'Mail Address', 3.86, 3.3918, 1.5851, 0.3703);
  lead(s, 'madogco@mail.com', 3.8749, 2.8061, 1.7993, 0.3306);
  heading(s, 'Website Address', 7.1092, 3.3918, 2.0286, 0.3703);
  lead(s, 'www.madog.com/project', 7.1241, 2.8061, 2.1755, 0.3306);
}

// ---------------------------------------------------------------- output
const deck = new PptxGenJS();
deck.defineLayout({ name: 'W16x9', width: 13.333, height: 7.5 });
deck.layout = 'W16x9';
deck.theme = { headFontFace: HEAD, bodyFontFace: BODY };
deck.author = 'Madog.Corp';
deck.title = 'Managing Properties, (Delivering) Possibilities!';

  slide1(deck.addSlide());
  slide2(deck.addSlide());
  slide3(deck.addSlide());
  slide4(deck.addSlide());
  slide5(deck.addSlide());
  slide6(deck.addSlide());
  slide7(deck.addSlide());
  slide8(deck.addSlide());
  slide9(deck.addSlide());
  slide10(deck.addSlide());
  slide11(deck.addSlide());
  slide12(deck.addSlide());
  slide13(deck.addSlide());
  slide14(deck.addSlide());
  slide15(deck.addSlide());
  slide16(deck.addSlide());
  slide17(deck.addSlide());
  slide18(deck.addSlide());
  slide19(deck.addSlide());
  slide20(deck.addSlide());
  slide21(deck.addSlide());
  slide22(deck.addSlide());
  slide23(deck.addSlide());
  slide24(deck.addSlide());
  slide25(deck.addSlide());

deck.writeFile({ fileName: path.join(__dirname, '16efaec1-b9d0-4da4-af28-fdd9f3d4cfdc_grok_final.pptx') })
  .then(function (f) { console.log('wrote ' + f); });