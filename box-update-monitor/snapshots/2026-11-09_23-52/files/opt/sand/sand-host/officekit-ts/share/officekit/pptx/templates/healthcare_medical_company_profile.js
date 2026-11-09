/**
 * "Medic" medical presentation deck (20 slides, 13.333 x 7.5 in) rebuilt with pptxgenjs.
 * Photographs in the source deck are replaced with flat grey placeholder rectangles.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

const pres = new PptxGenJS();
pres.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pres.layout = 'WIDE';
const S = pres.ShapeType;

/* ------------------------------------------------------------------ palette */
const NAVY = '000E44';
const TEAL = '00DBC3';
const WHITE = 'FFFFFF';
const GRAY = '262626';
const LIGHT = 'EBEBEB';
const TRACK = 'D9D9D9';
const PHOTO = 'CECECE'; // stand-in for the deck's raster photographs

const HEAD = 'Inter';
const BODY = 'Work Sans';

const NO_LINE = { type: 'none' };

// pptxgenjs rewrites shadow objects in place, so hand out a fresh one every time.
function cardShadow() {
  return { type: 'outer', color: '000000', opacity: 0.2, blur: 50, offset: 20, angle: 45 };
}

/* ------------------------------------------------------------------ helpers */
function rect(slide, x, y, w, h, color, opts) {
  const o = Object.assign({}, opts);
  const fill = { color };
  if (o.transparency) { fill.transparency = o.transparency; delete o.transparency; }
  if (o.shadow) o.shadow = cardShadow();
  slide.addShape(S.rect, Object.assign({ x, y, w, h, fill, line: NO_LINE }, o));
}

/** `radius` is the corner radius in inches (what pptxgenjs's rectRadius expects). */
function roundRect(slide, x, y, w, h, radius, color, opts) {
  slide.addShape(S.roundRect, Object.assign({
    x, y, w, h, rectRadius: radius, fill: { color }, line: NO_LINE,
  }, opts));
}

function ellipse(slide, x, y, w, h, color, opts) {
  const o = Object.assign({}, opts);
  const fill = { color };
  if (o.transparency) { fill.transparency = o.transparency; delete o.transparency; }
  slide.addShape(S.ellipse, Object.assign({ x, y, w, h, fill, line: NO_LINE }, o));
}

function poly(slide, x, y, color, points, opts) {
  const xs = points.map((p) => p.x);
  const ys = points.map((p) => p.y);
  slide.addShape(S.custGeom, Object.assign({
    x, y, w: Math.max.apply(null, xs), h: Math.max.apply(null, ys),
    fill: { color }, line: NO_LINE,
    points: points.map((p, i) => (i === 0 ? { x: p.x, y: p.y, moveTo: true } : p)).concat([{ close: true }]),
  }, opts));
}

/** Body copy / labels. Defaults match the deck's "Work Sans 10.5pt justified" style. */
function text(slide, str, x, y, w, h, opts) {
  slide.addText(str, Object.assign({
    x, y, w, h, fontFace: BODY, fontSize: 10.5, color: GRAY, valign: 'top', align: 'justify',
  }, opts));
}

/** Filled box with centred label (the deck's "01" badges and pill buttons). */
function labelBox(slide, str, x, y, w, h, fill, opts) {
  slide.addText(str, Object.assign({
    shape: S.rect, x, y, w, h, fill: { color: fill }, line: NO_LINE,
    fontFace: BODY, fontSize: 10.5, bold: true, color: WHITE, align: 'center', valign: 'middle',
  }, opts));
}

/* ------------------------------------------------------- reusable slide chrome */
const NAV = [
  { label: 'Home', x: 8.656, w: 0.579 },
  { label: 'About', x: 9.425, w: 0.581 },
  { label: 'Service', x: 10.186, w: 0.66 },
  { label: 'Contact', x: 10.988, w: 0.716 },
];

function plusMark(slide, x, y, w, h, color) {
  rect(slide, x, y + h * 0.39, w, h * 0.25, color); // horizontal arm
  rect(slide, x + w * 0.378, y, w * 0.243, h, color); // vertical arm
}

function hamburger(slide, x, y, color) {
  [0, 0.0577, 0.1154].forEach((dy) => roundRect(slide, x, y + dy, 0.232, 0.0216, 0.011, color));
}

function addChrome(slide, opts) {
  const o = opts || {};
  ellipse(slide, 0.492, 0.342, 0.222, 0.222, NAVY);
  plusMark(slide, 0.547, 0.399, 0.111, 0.107, WHITE);
  text(slide, 'Medic', 0.769, 0.302, 0.707, 0.303,
    { fontFace: HEAD, fontSize: 12, bold: true, color: NAVY, align: 'left', wrap: false });
  if (o.navPlate) rect(slide, 8.448, 0.302, 4.394, 0.375, WHITE, { shadow: true });
  NAV.forEach((item) => {
    text(slide, item.label, item.x, 0.367, item.w, 0.252,
      { fontSize: 9, color: NAVY, align: 'right', wrap: false });
  });
  hamburger(slide, 12.332, 0.417, NAVY);
}

/** Big section heading sitting on a teal highlight bar. */
function heading(slide, str, o) {
  rect(slide, o.barX, o.barY, o.barW, 0.248, TEAL);
  text(slide, str, o.x, o.y, o.w, 0.909, {
    fontFace: HEAD, fontSize: 48, bold: true, color: NAVY,
    align: o.align || 'left', wrap: o.wrap !== false,
  });
}

/** The heading shared by every "Infographic Slide" (slides 10-18). */
function infographicHeading(slide) {
  heading(slide, 'Infographic Slide',
    { barX: 3.842, barY: 1.761, barW: 5.65, x: 3.508, y: 1.205, w: 6.317, align: 'center' });
}

/** Grey block standing in for a photograph in the original deck. */
function photo(slide, x, y, w, h) {
  rect(slide, x, y, w, h, PHOTO);
  slide.addText('[image]', {
    x: Math.max(0, x), y: y + h / 2 - 0.18, w: Math.min(w, 13.333 - Math.max(0, x)), h: 0.36,
    fontFace: BODY, fontSize: 11, color: WHITE, align: 'center', valign: 'middle',
  });
}

/* ------------------------------------------------------------------- icons */
/** ">" glyph drawn as a stroked outline of thickness `t`. */
function chevronRight(slide, x, y, w, h, color) {
  const t = w * 0.5;
  poly(slide, x, y, color, [
    { x: 0, y: t * 0.6 }, { x: t * 0.6, y: 0 }, { x: w, y: h / 2 },
    { x: t * 0.6, y: h }, { x: 0, y: h - t * 0.6 }, { x: w - t, y: h / 2 },
  ]);
}

function globeSolid(slide, x, y, size, color, landColor) {
  ellipse(slide, x, y, size, size, color);
  poly(slide, x + size * 0.12, y + size * 0.14, landColor, [
    { x: 0, y: size * 0.16 }, { x: size * 0.22, y: 0 }, { x: size * 0.42, y: size * 0.2 },
    { x: size * 0.3, y: size * 0.34 }, { x: size * 0.34, y: size * 0.56 }, { x: size * 0.16, y: size * 0.4 },
  ]);
  poly(slide, x + size * 0.5, y + size * 0.5, landColor, [
    { x: 0, y: 0 }, { x: size * 0.26, y: size * 0.1 }, { x: size * 0.14, y: size * 0.34 },
  ]);
}

function globeWire(slide, x, y, size, color) {
  const ring = { fill: NO_LINE, line: { color, width: 2 } };
  slide.addShape(S.ellipse, Object.assign({ x, y, w: size, h: size }, ring));
  slide.addShape(S.ellipse, Object.assign({ x: x + size * 0.29, y, w: size * 0.42, h: size }, ring));
  rect(slide, x + size * 0.02, y + size * 0.47, size * 0.96, size * 0.05, color);
  rect(slide, x + size * 0.14, y + size * 0.24, size * 0.72, size * 0.045, color);
  rect(slide, x + size * 0.14, y + size * 0.71, size * 0.72, size * 0.045, color);
}

function checkCircle(slide, x, y, size, color) {
  ellipse(slide, x, y, size, size, color);
  poly(slide, x + size * 0.22, y + size * 0.3, WHITE, [
    { x: 0, y: size * 0.2 }, { x: size * 0.09, y: size * 0.11 }, { x: size * 0.22, y: size * 0.24 },
    { x: size * 0.47, y: 0 }, { x: size * 0.56, y: size * 0.09 }, { x: size * 0.22, y: size * 0.42 },
  ]);
}

function microscopeIcon(slide, x, y, w, h, color) {
  ellipse(slide, x + w * 0.28, y + h * 0.18, w * 0.72, h * 0.7, color); // objective arm
  ellipse(slide, x + w * 0.42, y + h * 0.32, w * 0.44, h * 0.42, WHITE); // arm cut-out
  roundRect(slide, x + w * 0.19, y + h * 0.06, w * 0.19, h * 0.46, 0.03, color, { rotate: -32 }); // eyepiece
  rect(slide, x, y + h * 0.55, w * 0.55, h * 0.09, color); // stage
  rect(slide, x + w * 0.2, y + h * 0.6, w * 0.11, h * 0.24, color); // stage post
  roundRect(slide, x, y + h * 0.84, w, h * 0.16, 0.03, color); // base
}

function clipboardIcon(slide, x, y, w, h, color) {
  roundRect(slide, x + w * 0.05, y + h * 0.09, w * 0.9, h * 0.91, 0.02, color); // board
  rect(slide, x + w * 0.28, y + h * 0.02, w * 0.44, h * 0.14, color); // clip
  ellipse(slide, x + w * 0.4, y, w * 0.2, h * 0.16, color);
  ellipse(slide, x + w * 0.44, y + h * 0.035, w * 0.12, h * 0.1, WHITE);
  rect(slide, x + w * 0.15, y + h * 0.3, w * 0.7, h * 0.55, WHITE); // paper
  poly(slide, x + w * 0.26, y + h * 0.44, color, [
    { x: 0, y: h * 0.16 }, { x: w * 0.08, y: h * 0.08 }, { x: w * 0.2, y: h * 0.2 },
    { x: w * 0.42, y: 0 }, { x: w * 0.5, y: h * 0.08 }, { x: w * 0.2, y: h * 0.36 },
  ]);
}

function medkitIcon(slide, x, y, w, h, color) {
  rect(slide, x + w * 0.3, y + h * 0.03, w * 0.4, h * 0.2, color); // handle
  rect(slide, x + w * 0.37, y + h * 0.1, w * 0.26, h * 0.16, WHITE);
  roundRect(slide, x, y + h * 0.2, w, h * 0.8, 0.03, color); // case
  plusMark(slide, x + w * 0.34, y + h * 0.4, w * 0.32, h * 0.42, WHITE);
}

function peopleIcon(slide, x, y, w, h, color) {
  ellipse(slide, x + w * 0.02, y + h * 0.14, w * 0.26, h * 0.26, color); // side heads
  ellipse(slide, x + w * 0.72, y + h * 0.14, w * 0.26, h * 0.26, color);
  roundRect(slide, x, y + h * 0.44, w * 0.3, h * 0.3, 0.03, color); // side bodies
  roundRect(slide, x + w * 0.7, y + h * 0.44, w * 0.3, h * 0.3, 0.03, color);
  ellipse(slide, x + w * 0.28, y + h * 0.04, w * 0.44, h * 0.44, color); // front head
  roundRect(slide, x + w * 0.2, y + h * 0.54, w * 0.6, h * 0.46, 0.05, color); // front body
}

function planeIcon(slide, x, y, w, h, color) {
  poly(slide, x, y, color, [
    { x: w, y: 0 }, { x: w * 0.36, y: h }, { x: w * 0.26, y: h * 0.6 }, { x: 0, y: h * 0.52 },
  ]);
}

function chartIcon(slide, x, y, w, h, color) {
  rect(slide, x, y, w * 0.1, h, color); // axes
  rect(slide, x, y + h * 0.9, w, h * 0.1, color);
  [0.24, 0.48, 0.72].forEach((dx, i) => {
    rect(slide, x + w * dx, y + h * (0.88 - 0.19 * (i + 1)), w * 0.16, h * 0.19 * (i + 1), color);
  });
  poly(slide, x + w * 0.2, y + h * 0.06, color, [
    { x: 0, y: h * 0.42 }, { x: w * 0.3, y: h * 0.14 }, { x: w * 0.48, y: h * 0.3 },
    { x: w * 0.8, y: 0 }, { x: w * 0.8, y: h * 0.14 }, { x: w * 0.5, y: h * 0.44 },
    { x: w * 0.32, y: h * 0.28 }, { x: w * 0.08, y: h * 0.52 },
  ]);
}

/** Two stacked pages. Drawn white on a navy card, so the separating gap is navy. */
function docIcon(slide, x, y, w, h, color) {
  roundRect(slide, x, y, w * 0.6, h * 0.68, 0.02, color); // back page
  rect(slide, x + w * 0.12, y + h * 0.11, w * 0.36, h * 0.06, NAVY); // ruled line
  slide.addShape(S.snip1Rect, { // front page, gap punched out with a navy halo
    x: x + w * 0.34, y: y + h * 0.26, w: w * 0.7, h: h * 0.78,
    fill: { color: NAVY }, line: NO_LINE, flipH: true,
  });
  slide.addShape(S.snip1Rect, {
    x: x + w * 0.38, y: y + h * 0.3, w: w * 0.62, h: h * 0.7,
    fill: { color }, line: NO_LINE, flipH: true,
  });
}

function envelopeIcon(slide, x, y, w, h, color) {
  rect(slide, x, y, w, h, color);
  poly(slide, x + w * 0.04, y + h * 0.06, WHITE, [
    { x: 0, y: 0 }, { x: w * 0.92, y: 0 }, { x: w * 0.46, y: h * 0.62 },
  ]);
}

function pinIcon(slide, x, y, w, h, color) {
  ellipse(slide, x, y, w, w, color);
  poly(slide, x + w * 0.16, y + w * 0.55, color, [
    { x: 0, y: 0 }, { x: w * 0.68, y: 0 }, { x: w * 0.34, y: h - w * 0.55 },
  ]);
  ellipse(slide, x + w * 0.3, y + w * 0.28, w * 0.4, w * 0.4, WHITE);
}

function phoneIcon(slide, x, y, w, h, color) {
  ellipse(slide, x, y + h * 0.3, w, h * 0.7, color); // rotary body
  roundRect(slide, x + w * 0.08, y, w * 0.84, h * 0.28, 0.04, color); // handset
  ellipse(slide, x + w * 0.28, y + h * 0.5, w * 0.44, h * 0.36, WHITE); // dial
}

function hourglassIcon(slide, x, y, w, h, color) {
  rect(slide, x, y, w, h * 0.1, color);
  rect(slide, x, y + h * 0.9, w, h * 0.1, color);
  poly(slide, x + w * 0.08, y + h * 0.1, color, [
    { x: 0, y: 0 }, { x: w * 0.84, y: 0 }, { x: w * 0.42, y: h * 0.4 },
  ]);
  poly(slide, x + w * 0.08, y + h * 0.5, color, [
    { x: w * 0.42, y: 0 }, { x: w * 0.84, y: h * 0.4 }, { x: 0, y: h * 0.4 },
  ]);
}

/* ================================================================= slide 1 */
function slide01() {
  const s = pres.addSlide();
  rect(s, 7.444, 2.159, 3.238, 5.341, NAVY);
  photo(s, 8.284, 1.373, 5.05, 5.493);
  addChrome(s);
  text(s, 'Medical Presentation', 1.103, 2.341, 2.775, 0.303,
    { fontSize: 12, color: NAVY, charSpacing: 3, align: 'left', wrap: false });
  rect(s, 1.17, 3.59, 3.643, 0.409, TEAL);
  text(s, 'Medic', 1.032, 2.644, 3.915, 1.582,
    { fontFace: HEAD, fontSize: 88, bold: true, color: NAVY, align: 'left', wrap: false });
  slide01Button(s, 1.17, 'Start Presentation');
}

function slide01Button(s, x, label) {
  s.addShape(S.rect, { x, y: 4.508, w: 2.173, h: 0.436, fill: NO_LINE, line: { color: NAVY, width: 1.5 } });
  text(s, label, x - 0.046, 4.529, 2.264, 0.338,
    { fontSize: 10.5, bold: true, color: NAVY, align: 'center', lineSpacingMultiple: 1.5 });
}

/* ================================================================= slide 2 */
const LOREM_LONG = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue '
  + 'massa. Fusce posuere congue massa. Fusce posuere, magna sed pulvinar osuere, posuere Fusce '
  + 'posuere congue massa. Fusce posuere, congue massa. ';

function slide02() {
  const s = pres.addSlide();
  rect(s, 3.218, 0.97, 2.347, 5.582, NAVY);
  photo(s, 0.0, 1.373, 5.05, 4.776);
  addChrome(s);
  heading(s, 'Introduction', { barX: 6.667, barY: 2.207, barW: 4.153, x: 6.627, y: 1.651, w: 4.28, wrap: false });
  text(s, LOREM_LONG + 'Maecenas porttitor congue massa. Fusce posuere congue massa sed pulvinar osuere.',
    6.603, 3.116, 5.101, 1.398, { lineSpacingMultiple: 1.5 });
  rect(s, 6.667, 5.07, 0.629, 0.619, TEAL, { shadow: true });
  s.addShape(S.line, {
    x: 6.842, y: 5.379, w: 0.279, h: 0,
    line: { color: WHITE, width: 4, endArrowType: 'triangle' },
  });
  text(s, 'Company Brief History', 7.636, 4.909, 5.101, 0.338, { bold: true, lineSpacingMultiple: 1.5 });
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ',
    7.636, 5.246, 4.068, 0.603, { lineSpacingMultiple: 1.5 });
}

/* ================================================================= slide 3 */
function slide03() {
  const s = pres.addSlide();
  rect(s, 7.786, 5.404, 1.323, 1.251, NAVY);
  photo(s, 8.284, 1.373, 5.05, 4.776);
  addChrome(s);
  heading(s, 'Vision & Mission', { barX: 1.125, barY: 2.153, barW: 5.444, x: 1.03, y: 1.597, w: 5.617, wrap: false });
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. '
    + 'Fusce posuere congue massa. Fusce posuere, magna sed pulvinar osuere, posuere.',
    1.057, 3.062, 5.61, 0.867, { lineSpacingMultiple: 1.5 });
  [
    { x: 1.057, title: 'Company Vision' },
    { x: 4.39, title: 'Company Mission' },
  ].forEach((col) => {
    text(s, col.title, col.x, 4.388, 2.257, 0.303, { fontSize: 12, bold: true, color: NAVY });
    text(s, 'Lorem ipsum dolor sit amet, cons ectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce.',
      col.x, 4.771, 2.277, 1.132, { lineSpacingMultiple: 1.5 });
  });
}

/* ================================================================= slide 4 */
function slide04() {
  const s = pres.addSlide();
  rect(s, 3.922, 2.958, 1.665, 3.661, NAVY);
  photo(s, 1.122, 3.349, 3.973, 4.151);
  addChrome(s);
  heading(s, 'Our Facilities', { barX: 1.122, barY: 2.153, barW: 4.336, x: 1.03, y: 1.597, w: 4.532, wrap: false });
  text(s, 'Enjoy Various Advanced Facilites', 7.541, 1.495, 5.101, 0.37,
    { fontSize: 12, bold: true, lineSpacingMultiple: 1.5 });
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere.',
    7.541, 1.95, 4.627, 0.603, { lineSpacingMultiple: 1.5 });
  ['01', '02', '03'].forEach((num, i) => {
    const y = 2.987 + i * 1.2495;
    labelBox(s, num, 7.636, y, 0.629, 0.619, TEAL, { fontSize: 14, shadow: cardShadow() });
    text(s, 'Lorem ipsum dolor sit amet, cons ectetuer adipiscing elit. Maecenas porttitor congue.',
      8.562, y + 0.008, 3.605, 0.603, { lineSpacingMultiple: 1.5 });
  });
}

/* ================================================================= slide 5 */
function slide05() {
  const s = pres.addSlide();
  heading(s, 'Our Products', { barX: 7.382, barY: 2.331, barW: 4.424, x: 7.301, y: 1.775, w: 4.586, wrap: false });
  rect(s, 1.424, 3.177, 2.202, 0.884, NAVY); // accent blocks tucked behind the photos
  rect(s, 11.752, 4.362, 0.884, 2.202, NAVY);
  addChrome(s);
  text(s, LOREM_LONG, 1.09, 4.982, 4.855, 1.133, { lineSpacingMultiple: 1.5 });
  photo(s, 7.382, 3.619, 4.818, 3.881);
  photo(s, 0.0, 1.393, 5.876, 2.227);
}

/* ================================================================= slide 6 */
function slide06() {
  const s = pres.addSlide();
  addChrome(s);
  heading(s, 'Our Doctors', { barX: 1.169, barY: 1.89, barW: 4.03, x: 1.081, y: 1.334, w: 4.206, wrap: false });
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. '
    + 'Fusce posuere congue massa. Fusce posuere, magna sed pulvinar osuere,. ',
    1.054, 2.629, 4.899, 0.868, { color: NAVY, lineSpacingMultiple: 1.5 });
  [
    { title: 'Doctor Profile 1', ty: 3.786, by: 4.169 },
    { title: 'Doctor Profile 2', ty: 5.059, by: 5.442 },
  ].forEach((row) => {
    text(s, row.title, 1.054, row.ty, 4.795, 0.303, { fontSize: 12, bold: true, color: NAVY });
    text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam vitae venenatis neque. ',
      1.03, row.by, 4.922, 0.601, { color: NAVY, lineSpacingMultiple: 1.5 });
  });
  [
    { x: 7.381, plate: 7.298, role: 'Doctor Profile 1', name: 'Francisco Andrade' },
    { x: 10.004, plate: 9.935, role: 'Doctor Profile 2', name: 'Brigitte Schwartz' },
  ].forEach((card) => {
    photo(s, card.x, 1.496, 2.174, 4.323);
    rect(s, card.plate, 5.354, 2.369, 0.76, NAVY); // name plate (rotated bar in source)
    text(s, card.role, card.plate + 0.15, 5.424, 2.302, 0.278, { color: LIGHT });
    text(s, card.name, card.plate + 0.15, 5.663, 2.302, 0.303, { fontSize: 12, bold: true, color: WHITE });
  });
}

/* ================================================================= slide 7 */
function slide07() {
  const s = pres.addSlide();
  rect(s, 1.169, 5.161, 11.038, 0.401, NAVY); // wide navy band behind the bottom photo strip
  photo(s, 0.0, 5.387, 13.333, 2.113);
  addChrome(s);
  heading(s, 'Our Services', { barX: 4.539, barY: 1.761, barW: 4.256, x: 4.441, y: 1.205, w: 4.451, align: 'center' });
  const services = [
    { x: 1.092, title: 'Service 1', tx: 1.826, icon: microscopeIcon, ix: 2.559, iy: 2.723, iw: 0.362, ih: 0.411, dy: 0 },
    { x: 4.992, title: 'Service 2', tx: 5.726, icon: clipboardIcon, ix: 6.459, iy: 2.733, iw: 0.358, ih: 0.414, dy: 0.002 },
    { x: 8.892, title: 'Service 3', tx: 9.627, icon: medkitIcon, ix: 10.36, iy: 2.811, iw: 0.411, ih: 0.333, dy: -0.003 },
  ];
  services.forEach((sv) => {
    sv.icon(s, sv.ix, sv.iy, sv.iw, sv.ih, NAVY);
    text(s, sv.title, sv.tx, 3.467 + sv.dy, 1.881, 0.303,
      { fontSize: 12, bold: true, color: NAVY, align: 'center' });
    text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam vitae venenatis',
      sv.x, 3.874 + sv.dy, 3.349, 0.601, { color: NAVY, align: 'center', lineSpacingMultiple: 1.5 });
  });
}

/* ================================================================= slide 8 */
function slide08() {
  const s = pres.addSlide();
  addChrome(s);
  heading(s, 'Portfolio', { barX: 1.75, barY: 1.89, barW: 2.869, x: 1.664, y: 1.334, w: 3.04, align: 'center', wrap: false });
  const frames = [
    { fx: 1.163, fy: 2.729, fw: 4.066, fh: 3.389, px: 1.261, py: 2.847, pw: 3.87, ph: 3.154 },
    { fx: 5.458, fy: 1.417, fw: 3.633, fh: 4.701, px: 5.556, py: 1.496, pw: 3.437, ph: 4.543 },
    { fx: 9.235, fy: 1.399, fw: 2.891, fh: 2.249, px: 9.334, py: 1.498, pw: 2.693, ph: 2.051 },
    { fx: 9.235, fy: 3.869, fw: 2.891, fh: 2.249, px: 9.334, py: 3.968, pw: 2.693, ph: 2.051 },
  ];
  frames.forEach((f) => rect(s, f.fx, f.fy, f.fw, f.fh, WHITE, { shadow: cardShadow() }));
  frames.forEach((f) => photo(s, f.px, f.py, f.pw, f.ph));
}

/* ================================================================= slide 9 */
function pricingCard(s, o) {
  rect(s, o.cardX, o.cardY, o.cardW, o.cardH, o.cardFill, o.cardShadow ? { shadow: cardShadow() } : undefined);
  rect(s, o.bannerX, o.bannerY, o.bannerW, o.bannerH, o.bannerFill,
    o.bannerAlpha ? { transparency: o.bannerAlpha } : undefined);
  text(s, o.plan, o.textX, o.planY, 2.148, o.planH, {
    fontSize: o.planSize, bold: true, color: o.planColor, align: 'center',
  });
  text(s, o.price, o.priceX, o.priceY, 2.148, 0.337,
    { fontSize: 14, bold: true, color: o.priceColor, align: 'center' });
  ['Our Service Class 001', 'Our Service Class 002', 'Our Service Class 003'].forEach((line, i) => {
    const y = o.listY + i * 0.331;
    ellipse(s, o.bulletX, y + 0.095, 0.087, 0.087, o.bulletColor);
    text(s, line, o.bulletX + 0.105, y, 1.883, o.listH, { fontSize: o.listSize, color: o.listColor });
  });
  rect(s, o.btnX, o.btnY, 2.198, 0.542, WHITE, o.btnShadow ? { shadow: cardShadow() } : undefined);
  text(s, 'EXPLORE NOW', o.btnX + 0.315, o.btnY + 0.128, 1.5, 0.286,
    { fontSize: 11, bold: true, color: NAVY });
  chevronRight(s, o.btnX + 1.731, o.btnY + 0.203, 0.092, 0.142, NAVY);
}

function slide09() {
  const s = pres.addSlide();
  addChrome(s);
  heading(s, 'Pricing Table', { barX: 4.479, barY: 1.761, barW: 4.375, x: 4.24, y: 1.205, w: 4.854, align: 'center' });
  pricingCard(s, {
    cardX: 1.712, cardY: 2.664, cardW: 2.892, cardH: 3.384, cardFill: NAVY,
    bannerX: 2.029, bannerY: 3.065, bannerW: 2.194, bannerH: 0.514, bannerFill: WHITE, bannerAlpha: 85,
    plan: 'REGULAR CLASS', textX: 2.049, planY: 3.154, planH: 0.337, planSize: 14, planColor: WHITE,
    price: '$50 /Month', priceX: 2.049, priceY: 3.672, priceColor: WHITE,
    listY: 4.096, listH: 0.278, listSize: 10.5, listColor: WHITE, bulletX: 2.234, bulletColor: WHITE,
    btnX: 2.024, btnY: 5.186, btnShadow: true,
  });
  pricingCard(s, {
    cardX: 5.046, cardY: 2.66, cardW: 3.248, cardH: 3.705, cardFill: WHITE, cardShadow: true,
    bannerX: 5.412, bannerY: 3.059, bannerW: 2.509, bannerH: 0.588, bannerFill: NAVY,
    plan: 'MEDIUM CLASS', textX: 5.592, planY: 3.168, planH: 0.37, planSize: 16, planColor: WHITE,
    price: '$75 /Month', priceX: 5.558, priceY: 3.741, priceColor: GRAY,
    listY: 4.23, listH: 0.286, listSize: 11, listColor: GRAY, bulletX: 5.745, bulletColor: NAVY,
    btnX: 5.558, btnY: 5.407, btnShadow: true,
  });
  pricingCard(s, {
    cardX: 8.736, cardY: 2.664, cardW: 2.892, cardH: 3.384, cardFill: TEAL,
    bannerX: 9.085, bannerY: 3.065, bannerW: 2.194, bannerH: 0.514, bannerFill: WHITE, bannerAlpha: 85,
    plan: 'PREMIUM CLASS', textX: 9.106, planY: 3.154, planH: 0.337, planSize: 14, planColor: WHITE,
    price: '$100 /Month', priceX: 9.106, priceY: 3.672, priceColor: WHITE,
    listY: 4.096, listH: 0.278, listSize: 10.5, listColor: WHITE, bulletX: 9.29, bulletColor: WHITE,
    btnX: 9.081, btnY: 5.186,
  });
}

/* ================================================================ slide 10 */
function slide10() {
  const s = pres.addSlide();
  addChrome(s);
  infographicHeading(s);
  ['01', '02', '03', '04'].forEach((num, i) => {
    const x = 1.173 + i * 2.8215;
    const fill = i % 2 === 0 ? NAVY : TEAL;
    rect(s, x, 2.648, 2.544, 2.544, fill);
    globeSolid(s, x + 1.07, 3.187, 0.404, WHITE, fill);
    text(s, 'Title Here', x + 0.384, 3.771, 1.777, 0.303,
      { fontSize: 12, bold: true, color: WHITE, align: 'center' });
    text(s, 'Lorem ipsum dolor sit amet, dolor', x + 0.384, 4.041, 1.777, 0.627,
      { fontSize: 11, color: WHITE, align: 'center', lineSpacingMultiple: 1.5 });
    labelBox(s, num, x, 5.58 + i * 0.0065, 0.515, 0.515, fill, { fontSize: 12 });
    text(s, 'Lorem ipsum dolor massa consect adipiscing', x + 0.619, 5.508 + i * 0.0065, 2.076, 0.601,
      { lineSpacingMultiple: 1.5 });
  });
}

/* ================================================================ slide 11 */
function slide11() {
  const s = pres.addSlide();
  addChrome(s);
  infographicHeading(s);
  ['01', '02', '03'].forEach((num, i) => {
    const cx = 1.603 + i * 3.7945;
    const tx = 2.084 + i * 3.8205;
    rect(s, cx, 3.095, 3.009, 2.686, NAVY, { shadow: cardShadow() });
    rect(s, cx - 0.471, 2.674, 2.613, 1.025, WHITE, { shadow: cardShadow() });
    labelBox(s, num, cx - 0.289, 2.867, 0.64, 0.64, NAVY, { fontSize: 12 });
    text(s, 'Title Here', tx, 2.884, 1.024, 0.303, { fontSize: 12, bold: true, wrap: false });
    for (let k = 0; k < 5; k += 1) {
      s.addShape(S.star5, {
        x: tx + 0.098 + k * 0.1835, y: 3.242, w: 0.133, h: 0.133, fill: { color: NAVY }, line: NO_LINE,
      });
    }
    text(s, 'Lorem ipsum dolor sit am et, consectetuer adip ipsum dolor sit am et, consecip sum dolor sit am',
      cx + 0.427, 4.031, 2.334, 1.181, { fontSize: 11, color: WHITE, lineSpacingMultiple: 1.5 });
    labelBox(s, 'More Information', cx + 0.644, 5.525, 1.722, 0.512, WHITE, { color: NAVY, shadow: cardShadow() });
  });
}

/* ================================================================ slide 12 */
function slide12() {
  const s = pres.addSlide();
  addChrome(s);
  infographicHeading(s);
  const bars = [
    { dx: 0.529, top: 3.207, h: 1.104 },
    { dx: 0.958, top: 3.717, h: 0.594 },
    { dx: 1.387, top: 3.207, h: 1.104 },
    { dx: 1.811, top: 3.481, h: 0.83 },
  ];
  ['01', '02', '03', '04'].forEach((num, i) => {
    const x = 1.167 + i * 2.8125;
    const fill = i % 2 === 0 ? NAVY : TEAL;
    rect(s, x, 2.648, 2.569, 2.375, fill);
    bars.forEach((b) => {
      rect(s, x + b.dx, 2.984, 0.227, 1.326, WHITE, { transparency: 85 });
      rect(s, x + b.dx, b.top, 0.227, b.h, WHITE);
    });
    text(s, 'Statistic Here', x + 0.161, 4.408, 2.248, 0.349,
      { fontSize: 11, color: WHITE, align: 'center', lineSpacingMultiple: 1.5 });
    labelBox(s, num, x, 5.395 + i * 0.016, 0.571, 0.571, fill, { fontSize: 14 });
    text(s, 'Lorem ipsum dolor sit amet, consect', x + 0.755, 5.379 + i * 0.016, 1.654, 0.603,
      { lineSpacingMultiple: 1.5 });
  });
}

/* ================================================================ slide 13 */
function slide13() {
  const s = pres.addSlide();
  addChrome(s);
  infographicHeading(s);
  [
    { pct: '47%', fill: NAVY },
    { pct: '87%', fill: NAVY },
    { pct: '90%', fill: TEAL },
  ].forEach((card, i) => {
    const x = 1.157 + i * 3.8175;
    rect(s, x, 2.667, 3.385, 3.139, card.fill);
    ellipse(s, x + 2.079, 3.028, 1.027, 1.027, WHITE, { transparency: 80 });
    globeWire(s, x + 2.303, 3.252 - i * 0.002, 0.579, WHITE);
    text(s, 'Title Here', x + 0.355, 3.137, 1.159, 0.337, { fontSize: 14, bold: true, color: WHITE });
    text(s, card.pct, x + 0.355, 3.421, 1.433, 0.774, { fontSize: 40, bold: true, color: WHITE });
    text(s, 'Lorem ipsum dolor sit am ipsum dolor sit am ipsum dolor', x + 0.355, 4.482, 2.751, 0.627,
      { fontSize: 11, color: WHITE, lineSpacingMultiple: 1.5 });
    labelBox(s, 'More Information', x + 0.831, 5.539, 1.722, 0.512, WHITE, { color: NAVY, shadow: cardShadow() });
  });
}

/* ================================================================ slide 14 */
function slide14() {
  const s = pres.addSlide();
  addChrome(s);
  infographicHeading(s);
  ['2020', '2021', '2022', '2023', '2024', '2025'].forEach((year, i) => {
    const x = 1.201 + i * 1.8702;
    rect(s, x, 2.648, 1.604, 1.604, i % 2 === 0 ? NAVY : TEAL);
    text(s, year, x + 0.121, 3.099, 1.361, 0.505,
      { fontSize: 24, bold: true, color: WHITE, align: 'center' });
    text(s, i === 4 ? 'Our Members' : 'Our Sales', x + 0.121, 3.526, 1.361, 0.252,
      { fontSize: 9, color: WHITE, align: 'center' });
  });
  [
    { x: 1.201, num: '01.', fill: NAVY },
    { x: 6.811, num: '02.', fill: TEAL },
  ].forEach((panel) => {
    rect(s, panel.x, 4.483, 5.345, 1.583, panel.fill);
    text(s, panel.num, panel.x + 0.4, 4.856, 0.96, 0.505, { fontSize: 24, bold: true, color: WHITE });
    text(s, 'Title Here', panel.x + 1.122, 4.97, 1.096, 0.303, { fontSize: 12, bold: true, color: WHITE });
    text(s, 'Lorem ipsum dolor sit amet, ipsum dolor sit amet, ipsum dolor sit amet, consect ip',
      panel.x + 2.289, 4.82, 2.583, 0.867, { color: WHITE, lineSpacingMultiple: 1.5 });
  });
}

/* ================================================================ slide 15 */
function slide15() {
  const s = pres.addSlide();
  addChrome(s);
  infographicHeading(s);
  rect(s, 1.178, 2.648, 5.489, 3.443, NAVY);
  [
    { label: 'Your Title Here 01', pct: '60%', barW: 2.395, knobX: 3.93, labelW: 2.395 },
    { label: 'Your Title Here 02', pct: '72%', barW: 3.101, knobX: 4.653, labelW: 2.596 },
    { label: 'Your Title Here 03', pct: '93%', barW: 3.525, knobX: 5.081, labelW: 2.395 },
  ].forEach((row, i) => {
    const dy = i * 0.938;
    slide15Track(s, 1.6485, 3.5365 + dy);
    s.addShape(S.roundRect, {
      x: 1.649, y: 3.52 + dy, w: row.barW, h: 0.118, rectRadius: 0.5,
      fill: { color: WHITE }, line: NO_LINE,
    });
    ellipse(s, row.knobX, 3.476 + dy, 0.197, 0.197, WHITE);
    text(s, row.label, 1.572, 3.106 + dy, row.labelW, 0.345,
      { fontSize: 12, color: WHITE, lineSpacingMultiple: 1.3 });
    text(s, row.pct, 5.568, 3.411 + dy, 0.704, 0.345,
      { fontSize: 12, color: WHITE, lineSpacingMultiple: 1.3 });
  });
  ['01', '02', '03'].forEach((num, i) => {
    const y = 2.663 + i * 1.2985;
    labelBox(s, num, 7.579, y, 0.634, 0.634, i === 1 ? TEAL : NAVY, { fontSize: 14 });
    text(s, 'Lorem ipsum dolor sit amet, consect etuer dolor sit amet, consect etuer adipiscing elit. '
      + 'dolor sit etuer adipiscing etuer adipiscing', 8.339, y + 0.001, 3.903, 0.868,
    { lineSpacingMultiple: 1.5 });
  });
}

function slide15Track(s, x, y) {
  s.addShape(S.roundRect, {
    x, y, w: 3.919, h: 0.102, rectRadius: 0.5,
    fill: { color: TRACK, transparency: 50 }, line: NO_LINE,
  });
}

/* ================================================================ slide 16 */
function slide16() {
  const s = pres.addSlide();
  addChrome(s);
  infographicHeading(s);
  [
    { x: 1.174, y: 2.648, num: '01.', titleY: 3.773, fill: NAVY },
    { x: 4.031, y: 2.667, num: '02.', titleY: 3.78, fill: TEAL },
    { x: 6.889, y: 2.687, num: '03.', titleY: 3.787, fill: NAVY },
    { x: 9.746, y: 2.707, num: '04.', titleY: 3.793, fill: TEAL },
  ].forEach((card) => {
    rect(s, card.x, card.y, 2.438, 3.45, card.fill);
    text(s, card.num, card.x + 0.429, card.y + 0.465, 1.579, 0.64,
      { fontSize: 32, bold: true, color: WHITE });
    text(s, 'Title Here', card.x + 0.429, card.titleY, 1.476, 0.303,
      { fontSize: 12, bold: true, color: WHITE });
    text(s, 'Lorem ipsum dolor ctetuer adipiscing ipsum dolo', card.x + 0.429, card.y + 1.406, 1.579, 0.868,
      { color: WHITE, lineSpacingMultiple: 1.5 });
    labelBox(s, 'Learn More', card.x + 0.532, card.y + 2.472, 1.476, 0.523, WHITE, { color: NAVY });
  });
}

/* ================================================================ slide 17 */
function slide17() {
  const s = pres.addSlide();
  addChrome(s);
  infographicHeading(s);
  [
    { x: 1.188, year: '2023', checkX: 1.586 },
    { x: 5.028, year: '2024', checkX: 5.438 },
    { x: 8.868, year: '2025', checkX: 9.29 },
  ].forEach((card) => {
    rect(s, card.x, 2.653, 3.303, 2.198, WHITE, { shadow: cardShadow() });
    text(s, card.year, card.x + 0.409, 2.958, 2.484, 0.707,
      { fontSize: 36, bold: true, color: NAVY, align: 'center' });
    text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing', card.x + 0.409, 3.616, 2.484, 0.627,
      { fontSize: 11, align: 'center', lineSpacingMultiple: 1.5 });
    labelBox(s, 'More Information', card.x + 0.642, 4.6, 2.019, 0.484, NAVY, { shadow: cardShadow() });
    checkCircle(s, card.checkX, 5.531, 0.239, NAVY);
    text(s, 'Lorem ipsu dolo sit amet, consec etuer adipisc', card.checkX + 0.339, 5.441, 2.157, 0.603,
      { lineSpacingMultiple: 1.5 });
  });
}

/* ================================================================ slide 18 */
function slide18() {
  const s = pres.addSlide();
  addChrome(s);
  infographicHeading(s);
  [
    { x: 1.264, y: 2.592, fill: NAVY, icon: peopleIcon },
    { x: 1.264, y: 4.5, fill: TEAL, icon: chartIcon },
    { x: 6.768, y: 2.592, fill: TEAL, icon: planeIcon },
    { x: 6.768, y: 4.5, fill: NAVY, icon: docIcon },
  ].forEach((card) => {
    rect(s, card.x, card.y, 5.305, 1.695, card.fill);
    rect(s, card.x + 0.284, card.y + 0.234, 0.117, 1.167, WHITE, { transparency: 80 });
    rect(s, card.x + 0.503, card.y + 0.234, 1.167, 1.167, WHITE, { transparency: 80 });
    card.icon(s, card.x + 0.877, card.y + 0.608, 0.42, 0.42, WHITE);
    text(s, 'Title Here', card.x + 1.933, card.y + 0.375, 1.871, 0.303,
      { fontSize: 12, bold: true, color: WHITE });
    text(s, 'Lorem ipsum dolor sit amet, adipiscing elit. Mae cenas porttmassa. ',
      card.x + 1.933, card.y + 0.663, 3.084, 0.625, { fontSize: 11, color: WHITE, lineSpacingMultiple: 1.5 });
  });
}

/* ================================================================ slide 19 */
function slide19() {
  const s = pres.addSlide();
  rect(s, 8.073, 2.792, 0.884, 3.333, NAVY);
  photo(s, 8.448, 1.417, 4.886, 6.083);
  addChrome(s, { navPlate: true });
  heading(s, 'Contact Us', { barX: 1.193, barY: 1.89, barW: 3.756, x: 1.133, y: 1.334, w: 3.876, wrap: false });
  [
    {
      title: 'Email', tx: 1.64, ty: 2.99, vx: 1.186, vy: 3.449, align: 'left',
      lines: ['emailname1@mail.com', 'emailname2@mail.com'],
      icon: envelopeIcon, ix: 1.295, iy: 3.052, iw: 0.274, ih: 0.171,
    },
    {
      title: 'Address', tx: 4.423, ty: 2.967, vx: 4.142, vy: 3.427, align: 'justify',
      lines: ['Company Street, Company State, 1234'],
      icon: pinIcon, ix: 4.247, iy: 3.044, iw: 0.148, ih: 0.217,
    },
    {
      title: 'Phone Number', tx: 1.64, ty: 4.795, vx: 1.193, vy: 5.254, align: 'left',
      lines: ['Company (+123) 456 789', 'Personal (+123) 456 789'],
      icon: phoneIcon, ix: 1.292, iy: 4.785, iw: 0.277, ih: 0.261,
    },
    {
      title: 'Office Hours', tx: 4.42, ty: 4.795, vx: 4.142, vy: 5.254, align: 'left',
      lines: ['Monday - Saturday', '08.00 AM \u2013 03.00 PM'],
      icon: hourglassIcon, ix: 4.25, iy: 4.817, iw: 0.171, ih: 0.235,
    },
  ].forEach((block) => {
    block.icon(s, block.ix, block.iy, block.iw, block.ih, NAVY);
    text(s, block.title, block.tx, block.ty, 1.739, 0.303, { fontSize: 12, bold: true, color: NAVY });
    s.addText(block.lines.map((line, i) => ({
      text: line, options: { breakLine: i < block.lines.length - 1 },
    })), {
      x: block.vx, y: block.vy, w: 2.225, h: 0.606, fontFace: BODY, fontSize: 10.5, color: NAVY,
      align: block.align, valign: 'top', lineSpacingMultiple: 1.5,
    });
  });
}

/* ================================================================ slide 20 */
function slide20() {
  const s = pres.addSlide();
  rect(s, 2.542, 2.159, 3.238, 5.341, NAVY);
  photo(s, 0.0, 1.373, 5.031, 5.493);
  addChrome(s);
  text(s, 'Medical Presentation', 6.738, 2.341, 2.775, 0.303,
    { fontSize: 12, color: NAVY, charSpacing: 3, align: 'left', wrap: false });
  rect(s, 6.736, 3.59, 5.383, 0.409, TEAL);
  text(s, 'Thank You', 6.667, 2.752, 5.521, 1.313,
    { fontFace: HEAD, fontSize: 72, bold: true, color: NAVY, align: 'center' });
  slide01Button(s, 6.805, 'End Presentation');
}

/* ------------------------------------------------------------------- build */
[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
  .forEach((build) => build());

pres.writeFile({
  fileName: path.join(__dirname, '1383e333-10e0-4f90-a225-706dfd4e7a30_grok_final.pptx'),
}).then((f) => console.log('wrote', f));
