/**
 * Recreation of a 40-slide 20x11.25in deck with pptxgenjs.
 * Run: node 06c68e40-2433-409c-a130-152def107fd6_grok_final.js
 */
'use strict';
const pptxgen = require('pptxgenjs');
const path = require('path');

// ---------------------------------------------------------------- palette
const BG        = 'E7E6E6';  // slide background (theme lt2)
const GREEN     = '429675';  // accent1
const BLUE      = '8496B0';  // accent5
const BLUE_DARK = '5A6F8D';  // accent5, 75% luminance
const BEIGE     = 'DBD2C9';  // accent2
const WHITE     = 'FFFFFF';
const BLACK     = '000000';
const NUM_TEXT  = 'F2F2F2';  // slide-number digits

const TITLE_FONT = 'Poppins';
const BODY_FONT  = 'Sora ExtraLight';

// arrow-icon directions (degrees, clockwise from pointing right)
const RIGHT = 0, DOWN = 90, LEFT = 180, UP = 270;

// ------------------------------------------------------------ body copy
const FASHION_1 =
  "Fashion is a dynamic form of self-expression that reflects individual identity, cultural influences, and social trends. More than just clothing.";
const FASHION_2 =
  "Fashion is a dynamic form of self-expression that reflects individual identity, cultural influences, and social trends. More than just clothing, fashion represents a visual language through which people.";
const FASHION_3 =
  "Fashion is a dynamic form of self-expression that reflects individual identity, cultural influences, and social trends. More than just clothing, fashion represents a visual language through which people communicate.";
const FASHION_4 =
  "Fashion is a dynamic form of self-expression that reflects individual identity, cultural influences, and social trends. More than just clothing, fashion represents a visual language through which people communicate their personality.";
const FASHION_5 =
  "Fashion is a dynamic form of self-expression that reflects individual identity, cultural influences, and social trends. More than just clothing, fashion represents a visual language through which people communicate their personality, values, and even emotions. ";
const FASHION_6 =
  "Fashion is a dynamic form of self-expression that reflects individual identity, cultural influences, and social trends. More than just clothing, fashion represents a visual language through which people communicate their personality, values, and even emotions. Styles change with time.";
const FASHION_7 =
  "Fashion is a dynamic form of self-expression that reflects individual identity, cultural influences, and social trends. More than just clothing, fashion represents a visual language through which people communicate their personality, values, and even emotions. Styles change with time, often influenced by historical events, popular culture.";
const OBSERVE_1 =
  "Observation allows individuals to see connections.";
const OBSERVE_2 =
  "Observation allows individuals to see connections and patterns.";
const OBSERVE_3 =
  "Observation allows individuals to see connections and patterns that others might overlook.";
const OBSERVE_4 =
  "Observation allows individuals to see connections and patterns that others might overlook, leading to innovative.";
const VISION_1 =
  "A compelling future vision acts as a guiding force.";
const VISION_2 =
  "A compelling future vision acts as a guiding force, enabling individuals, organizations, and societies to navigate.";
const BIOENERGY_1 =
  "The future of bioenergy is poised to move beyond traditional, first-generation fuels derived from food crops. Significant advancements are focused.";
const BIOENERGY_2 =
  "The future of bioenergy is poised to move beyond traditional, first-generation fuels derived from food crops. Significant advancements are focused on developing advanced biofuels.";
const BIOSHIFT =
  "This shift is crucial for addressing concerns about land use and food security, while also maximizing the use of waste streams. Innovations in thermochemical.";
const BIOPRODUCTS =
  "Styles change with time. This is a model that uses biomass to produce not just energy, but also a range of other valuable bioproducts like bioplastics, industrial chemicals, and lubricants. ";
const CLOTHING =
  "More than just clothing, fashion represents a visual language through which people communicate their personality, values, and even emotions. Styles change with time.";

// -------------------------------------------------------------- helpers
const NO_LINE = { type: 'none' };
const rect = (b, color) => ({ x: b[0], y: b[1], w: b[2], h: b[3], fill: { color }, line: NO_LINE });

function txt(s, content, b, o) {
  const runs = Array.isArray(content)
    ? content.map((t, i) => ({ text: t, options: { breakLine: i < content.length - 1 } }))
    : content;
  s.addText(runs, Object.assign({ x: b[0], y: b[1], w: b[2], h: b[3], margin: 0, valign: 'top', fit: 'none' }, o));
}

const title = (s, c, b, o) => txt(s, c, b, Object.assign(
  { fontFace: TITLE_FONT, fontSize: 72, bold: true, charSpacing: -1.5, lineSpacingMultiple: 0.75, color: BLACK }, o));

const heading = (s, c, b, o) => txt(s, c, b, Object.assign(
  { fontFace: TITLE_FONT, fontSize: 20, bold: true, lineSpacingMultiple: 0.75, color: BLACK }, o));

const body = (s, c, b, o) => txt(s, c, b, Object.assign(
  { fontFace: BODY_FONT, fontSize: 18, lineSpacingMultiple: 1.2, color: BLACK }, o));

const num = (s, c, b, o) => txt(s, c, b, Object.assign(
  { fontFace: TITLE_FONT, fontSize: 44, bold: true, charSpacing: -1.5, lineSpacingMultiple: 0.85, color: GREEN }, o));

// outlined pill used for the keyword list on slide 23
const tag = (s, c, b) => s.addText(c, {
  x: b[0], y: b[1], w: b[2], h: b[3], shape: 'rect', line: { color: BLACK, width: 1 },
  fontFace: TITLE_FONT, fontSize: 20, charSpacing: 3, color: BLACK, align: 'center', valign: 'middle', fit: 'none',
});

// Circle enclosing an arrow; `dir` rotates the arrow inside the fixed circle.
const ARROW_PTS = [[0.197, 0.5, 0.803, 0.5], [0.487, 0.215, 0.803, 0.5], [0.487, 0.785, 0.803, 0.5]];
function arrowIcon(s, x, y, size, color, dir) {
  const stroke = { color, width: 3.5 };
  s.addShape('ellipse', { x, y, w: size, h: size, line: stroke });
  const a = (dir * Math.PI) / 180, cos = Math.cos(a), sin = Math.sin(a);
  const map = (px, py) => [x + size * (0.5 + (px - 0.5) * cos - (py - 0.5) * sin),
                           y + size * (0.5 + (px - 0.5) * sin + (py - 0.5) * cos)];
  ARROW_PTS.forEach(([x1, y1, x2, y2]) => {
    const [ax, ay] = map(x1, y1), [bx, by] = map(x2, y2);
    s.addShape('line', {
      x: Math.min(ax, bx), y: Math.min(ay, by), w: Math.abs(bx - ax), h: Math.abs(by - ay),
      flipV: (bx - ax) * (by - ay) < 0, line: stroke,
    });
  });
}

// green card with the two bottom corners rounded (slide 11)
const bottomRoundedCard = (s, b) => s.addShape('round2SameRect', {
  x: b[0], y: b[1], w: b[2], h: b[3], flipV: true, fill: { color: GREEN }, line: NO_LINE,
});

// green banner rotated on its side so a single corner stays rounded (slide 22)
const cornerCard = (s, b, rotate, flipH) => s.addShape('round1Rect', {
  x: b[0], y: b[1], w: b[2], h: b[3], rotate, flipH, flipV: true,
  rectRadius: 0.467, fill: { color: GREEN }, line: NO_LINE,
});

// Panel whose right half dissolves into the slide background (slide 36).
// pptxgenjs has no gradient fill, so the fade is a stack of translucent strips.
function fadePanel(s, b, color) {
  s.addShape('roundRect', Object.assign(rect(b, color), { rectRadius: 0.6 }));
  const steps = 18, start = b[0] + b[2] * 0.51, sw = (b[2] * 0.49) / steps;
  for (let i = 0; i < steps; i++) {
    s.addShape('rect', {
      x: start + i * sw, y: b[1], w: sw + 0.02, h: b[3],
      fill: { color: BG, transparency: 100 - Math.round(((i + 1) / steps) * 100) }, line: NO_LINE,
    });
  }
}

// --------------------------------------------------------- outline icons
// Each icon is a stroked outline drawn from primitives inside its bounding box.
const strokeBox = (s, x, y, w, h, color, radius) => s.addShape(radius ? 'roundRect' : 'rect',
  Object.assign({ x, y, w, h, line: { color, width: 2 } }, radius ? { rectRadius: radius } : {}));
const strokeLine = (s, x1, y1, x2, y2, color) => s.addShape('line', {
  x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
  flipV: (x2 - x1) * (y2 - y1) < 0, line: { color, width: 2 },
});

function iconGrid(s, b, color) {          // four rounded squares in a 2x2 grid
  const [x, y, w, h] = b, cw = w * 0.42, ch = h * 0.42;
  [[0, 0], [1, 0], [0, 1], [1, 1]].forEach(([cx, cy]) =>
    strokeBox(s, x + cx * w * 0.58, y + cy * h * 0.58, cw, ch, color, cw * 0.14));
}
function iconCamera(s, b, color) {        // film camera: body, two reels, lens wedge
  const [x, y, w, h] = b;
  strokeBox(s, x, y, w * 0.7, h, color, h * 0.1);
  strokeLine(s, x, y + h * 0.36, x + w * 0.7, y + h * 0.36, color);
  s.addShape('ellipse', { x: x + w * 0.11, y: y + h * 0.48, w: w * 0.21, h: h * 0.34, line: { color, width: 2 } });
  s.addShape('ellipse', { x: x + w * 0.38, y: y + h * 0.48, w: w * 0.21, h: h * 0.34, line: { color, width: 2 } });
  s.addShape('custGeom', { x: x + w * 0.7, y: y + h * 0.28, w: w * 0.3, h: h * 0.44, line: { color, width: 2 },
    points: [{ x: 0, y: h * 0.22 }, { x: w * 0.3, y: 0 }, { x: w * 0.3, y: h * 0.44 }, { close: true }] });
}
function iconTv(s, b, color) {            // television: screen, antennae, foot
  const [x, y, w, h] = b;
  strokeBox(s, x, y + h * 0.3, w, h * 0.58, color, h * 0.08);
  strokeLine(s, x + w * 0.28, y + h * 0.02, x + w * 0.5, y + h * 0.3, color);
  strokeLine(s, x + w * 0.5, y + h * 0.3, x + w * 0.74, y + h * 0.02, color);
  strokeLine(s, x + w * 0.12, y + h * 0.96, x + w * 0.88, y + h * 0.96, color);
}
function iconMonitor(s, b, color) {       // desktop monitor: screen, bezel, stand
  const [x, y, w, h] = b;
  strokeBox(s, x, y, w, h * 0.78, color, h * 0.07);
  strokeLine(s, x, y + h * 0.64, x + w, y + h * 0.64, color);
  strokeLine(s, x + w * 0.42, y + h * 0.78, x + w * 0.42, y + h * 0.92, color);
  strokeLine(s, x + w * 0.22, y + h * 0.95, x + w * 0.78, y + h * 0.95, color);
}
function iconScreen(s, b, color) {        // projector screen on a tripod
  const [x, y, w, h] = b;
  strokeLine(s, x, y + h * 0.08, x + w, y + h * 0.08, color);
  strokeBox(s, x + w * 0.05, y + h * 0.12, w * 0.9, h * 0.6, color, 0);
  strokeLine(s, x + w * 0.5, y + h * 0.72, x + w * 0.5, y + h * 0.86, color);
  strokeLine(s, x + w * 0.3, y + h * 0.99, x + w * 0.5, y + h * 0.82, color);
  strokeLine(s, x + w * 0.5, y + h * 0.82, x + w * 0.7, y + h * 0.99, color);
}

// rounded badge in the top-right corner holding the slide number
function slideNumber(s, n, color) {
  s.addShape('roundRect', Object.assign(rect([19.082, 0.312, 0.558, 0.558], color), { rectRadius: 0.093 }));
  s.addText(String(n), { x: 19.019, y: 0.291, w: 0.686, h: 0.599, fontFace: TITLE_FONT, fontSize: 16,
    color: NUM_TEXT, align: 'center', valign: 'middle', margin: 0, fit: 'none' });
}

// ----------------------------------------------------------- slide bodies

// Slide 1 - Cover
function slide01(s) {
  title(s, ["AUGMENTED REALITY ", "AND THE EVERYDAY"], [1.3323, 4.7176, 6.1886, 3.6907]);
  arrowIcon(s, 1.3323, 3.1625, 1.0432, GREEN, RIGHT);
  heading(s, "JOHN DOE", [1.3464, 8.6136, 3.5303, 0.375], { fontSize: 28, color: GREEN });
}

// Slide 2 - A sustainable future
function slide02(s) {
  s.addShape('rect', rect([0, 7.106, 20, 4.144], BLUE));
  title(s, "A SUSTAINABLE FUTURE", [3.1108, 1.3898, 13.7784, 0.9643], { align: 'center' });
  heading(s, "NEXT DECADE", [2.7769, 8.9476, 2.3053, 0.3214], { fontSize: 24, color: WHITE, align: 'center' });
  body(s, VISION_2, [1.3739, 9.406, 5.1112, 1.1199], { color: WHITE, align: 'center' });
  heading(s, "TRANSFORMING IDEAS ", [8.1854, 8.9476, 3.8375, 0.3214], { fontSize: 24, color: WHITE, align: 'center' });
  body(s, VISION_2, [7.5485, 9.406, 5.1112, 1.1199], { color: WHITE, align: 'center' });
  heading(s, "THE VISIONARIES", [14.8666, 8.9476, 2.8242, 0.3214], { fontSize: 24, color: WHITE, align: 'center' });
  body(s, VISION_2, [13.7231, 9.406, 5.1112, 1.1199], { color: WHITE, align: 'center' });
  iconGrid(s, [3.598, 7.917, 0.664, 0.67], WHITE);
  iconCamera(s, [9.635, 7.892, 0.938, 0.596], WHITE);
  iconTv(s, [15.882, 7.848, 0.794, 0.758], WHITE);
}

// Slide 3 - When every individual is a creator
function slide03(s) {
  title(s, "WHEN EVERY INDIVIDUAL IS A CREATOR", [1.3531, 2.8431, 6.8643, 2.7819]);
  arrowIcon(s, 1.3531, 1.4582, 1.0432, GREEN, RIGHT);
  heading(s, "NEW PLANET", [1.3531, 8.3429, 3.5303, 0.2668]);
  body(s, BIOENERGY_1, [1.3531, 8.7372, 6.8643, 1.0694]);
  heading(s, "SUSTAINABLE FUTURE", [1.3531, 6.2201, 3.5303, 0.2679]);
  body(s, BIOENERGY_1, [1.3531, 6.6143, 6.8643, 1.0694]);
}

// Slide 4 - Charting the course for energy
function slide04(s) {
  s.addShape('rect', rect([13.864, 0, 6.136, 11.25], BLUE));
  title(s, "CHARTING THE COURSE FOR ENERGY", [0.8731, 3.0554, 6.2481, 2.7819]);
  body(s, FASHION_7, [0.8731, 6.8348, 5.8178, 2.9081]);
  heading(s, "NEXT DECADE", [14.4697, 1.7199, 2.3053, 0.3214], { fontSize: 24, color: WHITE });
  body(s, VISION_2, [14.4697, 2.1783, 5.1112, 1.1199], { color: WHITE });
  heading(s, "TRANSFORMING IDEAS ", [14.4697, 4.8189, 3.8375, 0.3214], { fontSize: 24, color: WHITE });
  body(s, VISION_2, [14.4697, 5.2774, 5.1112, 1.1199], { color: WHITE });
  heading(s, "THE VISIONARIES", [14.4697, 7.9179, 2.8242, 0.3214], { fontSize: 24, color: WHITE });
  body(s, VISION_2, [14.4697, 8.3764, 5.1112, 1.1199], { color: WHITE });
  arrowIcon(s, 0.8731, 1.5751, 1.0432, GREEN, RIGHT);
}

// Slide 5 - From products to systems
function slide05(s) {
  title(s, "FROM PRODUCTS TO SYSTEMS", [11.4147, 2.5605, 6.3978, 2.7819]);
  arrowIcon(s, 11.4147, 1.1756, 1.0432, GREEN, LEFT);
  heading(s, "GLOBAL ENERGY", [11.4147, 8.1962, 3.5303, 0.2679]);
  body(s, OBSERVE_3, [11.4147, 8.5905, 3.5303, 1.4329]);
  heading(s, "SUSTAINABLE FUTURE", [15.4858, 8.1962, 3.5303, 0.2679]);
  body(s, OBSERVE_3, [15.4858, 8.5905, 3.5303, 1.4329]);
  body(s, FASHION_6, [11.4147, 5.8276, 7.6015, 1.7964]);
}

// Slide 6 - Charting the course - four points
function slide06(s) {
  heading(s, "NEW ERA", [8.5455, 9.0747, 3.5303, 0.2668]);
  body(s, OBSERVE_4, [8.5455, 9.4689, 4.865, 0.9694], { fontSize: 16 });
  num(s, "02", [8.5455, 8.2966, 1.1212, 0.6425]);
  heading(s, "GLOBAL ENERGY", [8.5455, 6.3278, 3.5303, 0.2679]);
  body(s, OBSERVE_4, [8.5455, 6.722, 4.865, 0.9694], { fontSize: 16 });
  num(s, "01", [8.5455, 5.5497, 1.1212, 0.6425]);
  heading(s, "NEW PLANET", [14.1818, 9.0747, 3.5303, 0.2668]);
  body(s, OBSERVE_4, [14.1818, 9.4689, 4.865, 0.9694], { fontSize: 16 });
  num(s, "04", [14.1818, 8.2966, 1.1212, 0.6425]);
  heading(s, "SUSTAINABLE FUTURE", [14.1818, 6.3278, 3.5303, 0.2679]);
  body(s, OBSERVE_4, [14.1818, 6.722, 4.865, 0.9694], { fontSize: 16 });
  num(s, "03", [14.1818, 5.5497, 1.1212, 0.6425]);
  title(s, "CHARTING THE COURSE FOR ENERGY", [0.8731, 5.625, 6.2481, 2.7819]);
  arrowIcon(s, 0.8731, 8.5107, 1.0432, GREEN, RIGHT);
}

// Slide 7 - Man behind project
function slide07(s) {
  s.addShape('rect', rect([0, 0, 20, 5.788], BEIGE));
  title(s, "MAN BEHIND PROJECT", [3.1108, 1.3898, 13.7784, 0.9643], { align: 'center' });
  heading(s, "TAN MALAKA", [7.8106, 8.1936, 4.3788, 0.3214], { fontSize: 24, align: 'center' });
  body(s, OBSERVE_2, [7.5675, 9.0793, 4.865, 0.6275], { fontSize: 16, align: 'center' });
  heading(s, "PROJECT SPECIALIST", [7.8106, 8.5748, 4.3788, 0.2679], { bold: false, align: 'center' });
  heading(s, "COMBIS HATARA", [13.75, 8.1936, 4.3788, 0.3214], { fontSize: 24, align: 'center' });
  body(s, OBSERVE_2, [13.5069, 9.0793, 4.865, 0.6275], { fontSize: 16, align: 'center' });
  heading(s, "DESIGN SPECIALIST", [13.75, 8.5748, 4.3788, 0.2679], { bold: false, align: 'center' });
  heading(s, "HATTA DJASARA", [1.8712, 8.1936, 4.3788, 0.3214], { fontSize: 24, align: 'center' });
  body(s, OBSERVE_2, [1.6281, 9.0793, 4.865, 0.6275], { fontSize: 16, align: 'center' });
  heading(s, "PROJECT MANAGER", [1.8712, 8.5748, 4.3788, 0.2679], { bold: false, align: 'center' });
}

// Slide 8 - Access in a clean energy world
function slide08(s) {
  s.addShape('rect', rect([10, 1.5, 5, 8.25], BLUE));
  title(s, "ACCESS IN A CLEAN ENERGY WORLD", [0.9473, 3.7837, 3.4163, 2.2554], { fontSize: 44 });
  body(s, FASHION_1, [0.995, 6.379, 3.4163, 2.1599]);
  arrowIcon(s, 0.9473, 2.5455, 0.8969, GREEN, RIGHT);
  title(s, "COURSE FOR GLOBAL ENERGY", [10.768, 3.7837, 3.4163, 2.2554], { fontSize: 44, color: WHITE });
  body(s, FASHION_1, [10.8157, 6.379, 3.4163, 2.1599], { color: WHITE });
  arrowIcon(s, 10.768, 2.5455, 0.8969, WHITE, RIGHT);
}

// Slide 9 - Modular reactors
function slide09(s) {
  title(s, "MODULAR REACTORS IN CLEAN ENERGY", [1.2879, 6.7113, 5.8178, 3.6907]);
  arrowIcon(s, 1.2879, 5.109, 1.0432, GREEN, 322);
  body(s, FASHION_6, [13.1489, 7.7275, 5.5632, 2.5234], { align: 'right' });
}

// Slide 10 - Power of ocean currents
function slide10(s) {
  s.addShape('rect', rect([10.47, 0, 9.53, 11.25], BEIGE));
  title(s, "POWER OF OCEAN CURRENTS", [1.114, 3.3746, 5.8178, 2.7819]);
  body(s, FASHION_6, [1.114, 7.0026, 5.3103, 2.5234]);
  arrowIcon(s, 1.114, 1.9898, 1.0432, GREEN, RIGHT);
  heading(s, "NEW PLANET", [13.5133, 4.9078, 3.5303, 0.2668]);
  body(s, OBSERVE_4, [13.5133, 5.302, 4.865, 0.9694], { fontSize: 16 });
  heading(s, "SUSTAINABLE FUTURE", [13.5133, 1.92, 3.5303, 0.2679]);
  body(s, OBSERVE_4, [13.5133, 2.3142, 4.865, 0.9694], { fontSize: 16 });
  heading(s, "FUTURE OF POWER", [13.5133, 7.9665, 3.5303, 0.2668]);
  body(s, OBSERVE_4, [13.5133, 8.3607, 4.865, 0.9694], { fontSize: 16 });
}

// Slide 11 - Clean fuels for aviation
function slide11(s) {
  s.addShape('rect', rect([10.47, 0, 9.53, 11.25], BEIGE));
  bottomRoundedCard(s, [8.258, 7.97, 4.515, 1.879]);
  bottomRoundedCard(s, [13.515, 7.97, 4.515, 1.879]);
  title(s, "CLEAN FUELS FOR AVIATION", [1.114, 3.3746, 5.8178, 2.7819]);
  arrowIcon(s, 1.114, 1.9898, 1.0432, GREEN, RIGHT);
  body(s, FASHION_6, [1.114, 7.0026, 5.3103, 2.5234]);
  heading(s, "THERMAL SOLUTIONS", [8.75, 8.3695, 3.5303, 0.2679], { color: WHITE });
  body(s, OBSERVE_1, [8.75, 8.7637, 3.5303, 0.6275], { fontSize: 16, color: WHITE });
  heading(s, "GLOBAL ENERGY", [14.0076, 8.3788, 3.5303, 0.2679], { color: WHITE });
  body(s, OBSERVE_1, [14.0076, 8.773, 3.5303, 0.6275], { fontSize: 16, color: WHITE });
}

// Slide 12 - Creating spaces that heal
function slide12(s) {
  heading(s, "NEW ERA", [8.3715, 9.2051, 3.5303, 0.2668]);
  body(s, OBSERVE_4, [8.3715, 9.5993, 4.865, 0.9694], { fontSize: 16 });
  num(s, "02", [8.3715, 8.427, 1.1212, 0.6425]);
  heading(s, "GLOBAL ENERGY", [8.3715, 6.4582, 3.5303, 0.2679]);
  body(s, OBSERVE_4, [8.3715, 6.8524, 4.865, 0.9694], { fontSize: 16 });
  num(s, "01", [8.3715, 5.6801, 1.1212, 0.6425]);
  heading(s, "NEW PLANET", [14.0079, 9.2051, 3.5303, 0.2668]);
  body(s, OBSERVE_4, [14.0079, 9.5993, 4.865, 0.9694], { fontSize: 16 });
  num(s, "04", [14.0079, 8.427, 1.1212, 0.6425]);
  heading(s, "SUSTAINABLE FUTURE", [14.0079, 6.4582, 3.5303, 0.2679]);
  body(s, OBSERVE_4, [14.0079, 6.8524, 4.865, 0.9694], { fontSize: 16 });
  num(s, "03", [14.0079, 5.6801, 1.1212, 0.6425]);
  title(s, "CREATING SPACES THAT HEAL", [1.0905, 5.7554, 6.2481, 2.7819]);
  arrowIcon(s, 1.0905, 8.6412, 1.0432, GREEN, RIGHT);
}

// Slide 13 - The rise of microgrids
function slide13(s) {
  s.addShape('roundRect', Object.assign(rect([1.39, 1.469, 17.22, 8.312], BG), { rectRadius: 0.369 }));
  title(s, "THE RISE OF MICROGRIDS AND COMMUNITY ", [2.2506, 2.814, 6.2652, 3.6907]);
  arrowIcon(s, 2.32, 6.7166, 1.0432, GREEN, RIGHT);
  body(s, [BIOENERGY_2, ' ', BIOSHIFT], [9.2893, 2.6685, 3.9947, 5.068]);
  body(s, [FASHION_5, ' ', BIOPRODUCTS], [13.8415, 2.6685, 3.9947, 6.1586]);
}

// Slide 14 - Recycling materials
function slide14(s) {
  s.addShape('rect', rect([14.182, 3.894, 5.818, 4.902], GREEN));
  s.addShape('rect', rect([0, 3.894, 5.818, 4.902], BEIGE));
  title(s, "RECYCLING MATERIALS AND MINIMIZING WASTE", [0.5382, 4.6098, 4.9618, 3.8329], { fontSize: 60 });
  arrowIcon(s, 14.972, 4.4555, 1.0432, WHITE, LEFT);
  body(s, FASHION_4, [15.0189, 5.7247, 4.4428, 2.5234], { color: WHITE });
}

// Slide 15 - Course for global energy
function slide15(s) {
  title(s, "COURSE FOR GLOBAL ENERGY", [1.114, 2.7383, 5.8178, 3.6907]);
  body(s, FASHION_6, [1.114, 7.1086, 4.6875, 2.8869]);
  arrowIcon(s, 1.114, 1.3534, 1.0432, GREEN, RIGHT);
  heading(s, "NEW PLANET", [14.1061, 5.3322, 3.5303, 0.2668]);
  body(s, OBSERVE_4, [14.1061, 5.7265, 4.865, 0.9694], { fontSize: 16 });
  num(s, "02", [14.1061, 4.5542, 1.1212, 0.6425]);
  heading(s, "SUSTAINABLE FUTURE", [14.1061, 2.211, 3.5303, 0.2679]);
  body(s, OBSERVE_4, [14.1061, 2.6052, 4.865, 0.9694], { fontSize: 16 });
  num(s, "01", [14.1061, 1.433, 1.1212, 0.6425]);
  heading(s, "FUTURE OF POWER", [14.1061, 8.4535, 3.5303, 0.2668]);
  body(s, OBSERVE_4, [14.1061, 8.8477, 4.865, 0.9694], { fontSize: 16 });
  num(s, "03", [14.1061, 7.6754, 1.1212, 0.6425]);
}

// Slide 16 - Project portfolio - eight tiles
function slide16(s) {
  title(s, "PROJECT PORTFOLIO", [1.114, 2.3898, 5.8178, 1.8731]);
  arrowIcon(s, 1.114, 1.0049, 1.0432, GREEN, RIGHT);
  body(s, FASHION_6, [9.5303, 2.6099, 9.4697, 1.4329]);
}

// Slide 17 - Transition for all community
function slide17(s) {
  title(s, "TRANSITION FOR ALL COMMUNITY", [1.2427, 1.7854, 6.4758, 2.7819]);
  body(s, FASHION_2, [8.7661, 1.7854, 4.6324, 2.1599]);
  body(s, CLOTHING, [14.4462, 1.7854, 4.6324, 1.7964]);
}

// Slide 18 - Power of ocean currents - split
function slide18(s) {
  title(s, "POWER OF OCEAN CURRENTS", [1.3314, 1.6362, 5.8178, 2.7819]);
  heading(s, "GLOBAL ENERGY", [11.1324, 8.2219, 3.5303, 0.2679]);
  body(s, OBSERVE_4, [11.1324, 8.6161, 3.5303, 1.2737], { fontSize: 16 });
  heading(s, "SUSTAINABLE FUTURE", [15.3557, 8.2219, 3.5303, 0.2679]);
  body(s, OBSERVE_4, [15.3557, 8.6161, 3.5303, 1.2737], { fontSize: 16 });
  arrowIcon(s, 11.1324, 6.9544, 0.7603, GREEN, LEFT);
  arrowIcon(s, 15.3557, 6.9418, 0.7603, GREEN, UP);
}

// Slide 19 - Sense of touch
function slide19(s) {
  title(s, "SENSE OF TOUCH", [3.2879, 8.3164, 5.8178, 1.8731], { color: WHITE });
  arrowIcon(s, 1.7828, 8.3167, 1.0432, WHITE, UP);
  body(s, FASHION_6, [9.6607, 8.5148, 9.1871, 1.4329], { color: WHITE });
}

// Slide 20 - Bridge to net-zero
function slide20(s) {
  s.addShape('rect', rect([0, 0, 20, 6.065], BG));
  title(s, "BRIDGE TO NET-ZERO", [2.8314, 1.4333, 5.4512, 1.8731]);
  arrowIcon(s, 1.4434, 1.39, 1.0432, GREEN, DOWN);
  body(s, FASHION_6, [8.9723, 1.4333, 9.4697, 1.4329]);
}

// Slide 21 - Cities as living organisms
function slide21(s) {
  title(s, "CITIES AS LIVING ORGANISMS", [1.2879, 3.5591, 6.6469, 2.7819], { color: WHITE });
  arrowIcon(s, 1.2879, 1.9569, 1.0432, WHITE, RIGHT);
  body(s, FASHION_6, [2.8062, 6.8982, 5.3103, 2.5234], { color: WHITE });
}

// Slide 22 - The future of work and leisure
function slide22(s) {
  s.addShape('rect', rect([6.59, 0, 6.819, 11.25], BEIGE));
  cornerCard(s, [2.76, 6.301, 1.848, 5.812], 90, false);
  cornerCard(s, [15.392, 6.301, 1.848, 5.812], 270, true);
  heading(s, "THERMAL SOLUTIONS", [1.837, 8.691, 3.5303, 0.2679], { color: WHITE });
  body(s, OBSERVE_1, [1.837, 9.0852, 3.5303, 0.6275], { fontSize: 16, color: WHITE });
  heading(s, "GLOBAL ENERGY", [14.5505, 8.7003, 3.5303, 0.2679], { color: WHITE, align: 'right' });
  body(s, OBSERVE_1, [14.5505, 9.0946, 3.5303, 0.6275], { fontSize: 16, color: WHITE, align: 'right' });
  title(s, "THE FUTURE OF WORK AND LEISURE", [7.3531, 2.8635, 5.8178, 3.6907]);
  arrowIcon(s, 7.3531, 1.3916, 1.0432, GREEN, DOWN);
  body(s, FASHION_6, [7.3531, 7.2762, 5.3103, 2.5234]);
}

// Slide 23 - The architect of the digital twin
function slide23(s) {
  tag(s, "SIGNIFICANCE", [10.6111, 7.9294, 3.0417, 0.6028]);
  tag(s, " SHIFT TOWARDS", [13.8361, 7.9294, 3.3709, 0.6028]);
  tag(s, "CHALLENGING CONVENTIONS", [10.6111, 8.6914, 5.1438, 0.6028]);
  tag(s, "FASHIONABLY", [15.9614, 8.6914, 2.9043, 0.6028]);
  tag(s, "ACCESSORIES", [10.6111, 9.4533, 2.8626, 0.6028]);
  tag(s, "TRENDSPOTTING", [13.6528, 9.4533, 3.3709, 0.6028]);
  title(s, ["THE ARCHITECT OF THE ", "DIGITAL TWIN"], [10.6111, 2.8431, 7.6015, 2.7819]);
  arrowIcon(s, 10.6111, 1.4582, 1.0432, GREEN, LEFT);
  body(s, FASHION_4, [10.6111, 5.9648, 7.6015, 1.4329]);
}

// Slide 24 - Ensuring a transition for all
function slide24(s) {
  title(s, "ENSURING A TRANSITION FOR ALL", [1.3314, 2.5605, 6.8643, 2.7819]);
  arrowIcon(s, 1.3314, 1.1756, 1.0432, GREEN, RIGHT);
  heading(s, "GLOBAL ENERGY", [1.3314, 8.1962, 3.5303, 0.2679]);
  body(s, OBSERVE_3, [1.3314, 8.5905, 3.5303, 1.4329]);
  heading(s, "SUSTAINABLE FUTURE", [5.4025, 8.1962, 3.5303, 0.2679]);
  body(s, OBSERVE_3, [5.4025, 8.5905, 3.5303, 1.4329]);
  body(s, FASHION_6, [1.3314, 5.8276, 7.6015, 1.7964]);
}

// Slide 25 - Modeling our future world
function slide25(s) {
  title(s, "MODELING OUR FUTURE WORLD", [1.2879, 1.4406, 6.2556, 2.7819]);
  heading(s, "GLOBAL ENERGY", [8.3245, 2.5698, 4.9051, 0.3214], { fontSize: 24 });
  body(s, OBSERVE_4, [8.3245, 2.964, 4.9051, 1.0694]);
  heading(s, "SUSTAINABLE FUTURE", [13.9374, 2.5698, 4.9051, 0.3214], { fontSize: 24 });
  body(s, OBSERVE_4, [13.9374, 2.964, 4.9051, 1.0694]);
  iconMonitor(s, [8.324, 1.412, 0.824, 0.816], GREEN);
  iconScreen(s, [13.937, 1.405, 0.824, 0.831], GREEN);
}

// Slide 26 - Project portfolio - four values
function slide26(s) {
  s.addShape('rect', rect([0, 5.625, 20, 5.625], BLUE));
  title(s, "PROJECT PORTFOLIO", [1.114, 2.3028, 5.8178, 1.8731]);
  arrowIcon(s, 1.114, 0.918, 1.0432, GREEN, RIGHT);
  body(s, FASHION_6, [9.5303, 2.5229, 9.4697, 1.4329]);
  heading(s, "BE HUMAN", [2.2821, 8.6124, 1.7829, 0.3214], { fontSize: 24, color: WHITE, align: 'center' });
  body(s, VISION_1, [1.114, 9.0709, 4.119, 0.7563], { color: WHITE, align: 'center' });
  heading(s, "WELL-BEING", [6.7291, 8.6124, 2.0669, 0.3214], { fontSize: 24, color: WHITE, align: 'center' });
  body(s, VISION_1, [5.703, 9.0709, 4.119, 0.7563], { color: WHITE, align: 'center' });
  heading(s, "PERSONAL GROWTH", [10.6738, 8.6124, 3.3554, 0.3214], { fontSize: 24, color: WHITE, align: 'center' });
  body(s, VISION_1, [10.292, 9.0709, 4.119, 0.7563], { color: WHITE, align: 'center' });
  heading(s, "NEXT DECADE", [15.7878, 8.6124, 2.3053, 0.3214], { fontSize: 24, color: WHITE, align: 'center' });
  body(s, VISION_1, [14.881, 9.0709, 4.119, 0.7563], { color: WHITE, align: 'center' });
}

// Slide 27 - Project portfolio - filmstrip
function slide27(s) {
  title(s, "PROJECT PORTFOLIO", [3.2879, 8.1289, 5.8178, 1.8731]);
  arrowIcon(s, 1.7828, 8.1292, 1.0432, GREEN, UP);
  body(s, FASHION_6, [9.6607, 8.3273, 9.1871, 1.4329]);
}

// Slide 28 - Project portfolio - six tiles
function slide28(s) {
  title(s, "PROJECT PORTFOLIO", [1.114, 4.2376, 5.8178, 1.8731]);
  arrowIcon(s, 1.114, 2.744, 1.0432, GREEN, RIGHT);
  body(s, FASHION_2, [1.114, 6.5592, 4.6324, 2.1599]);
  heading(s, "PORTFOLIO", [8.4954, 4.7271, 1.867, 0.3214], { fontSize: 24, align: 'center' });
  body(s, "A compelling future vision,", [7.6957, 5.0947, 3.4665, 0.3928], { align: 'center' });
  heading(s, "PORTFOLIO", [12.3573, 4.7271, 1.867, 0.3214], { fontSize: 24, align: 'center' });
  body(s, "A compelling future vision,", [11.5576, 5.0947, 3.4665, 0.3928], { align: 'center' });
  heading(s, "PORTFOLIO", [16.2193, 4.7271, 1.867, 0.3214], { fontSize: 24, align: 'center' });
  body(s, "A compelling future vision,", [15.4196, 5.0947, 3.4665, 0.3928], { align: 'center' });
  heading(s, "PORTFOLIO", [8.4954, 9.6503, 1.867, 0.3214], { fontSize: 24, align: 'center' });
  body(s, "A compelling future vision,", [7.6957, 10.0179, 3.4665, 0.3928], { align: 'center' });
  heading(s, "PORTFOLIO", [12.3573, 9.6503, 1.867, 0.3214], { fontSize: 24, align: 'center' });
  body(s, "A compelling future vision,", [11.5576, 10.0179, 3.4665, 0.3928], { align: 'center' });
  heading(s, "PORTFOLIO", [16.2193, 9.6503, 1.867, 0.3214], { fontSize: 24, align: 'center' });
  body(s, "A compelling future vision,", [15.4196, 10.0179, 3.4665, 0.3928], { align: 'center' });
}

// Slide 29 - Turning waste into power
function slide29(s) {
  title(s, "TURNING WASTE INTO POWER", [1.114, 3.1355, 6.8208, 2.7819]);
  body(s, FASHION_6, [1.114, 6.4534, 5.3103, 2.5234]);
  arrowIcon(s, 1.114, 1.7506, 1.0432, GREEN, RIGHT);
}

// Slide 30 - The future of work and leisure - right
function slide30(s) {
  title(s, "THE FUTURE OF WORK AND LEISURE", [11.0705, 3.0374, 5.8178, 3.6907]);
  arrowIcon(s, 11.0705, 1.4351, 1.0432, GREEN, LEFT);
  body(s, FASHION_6, [13.1574, 7.3197, 5.3103, 2.5234]);
}

// Slide 31 - Ensuring a transition for all - right
function slide31(s) {
  title(s, "ENSURING A TRANSITION FOR ALL", [11.3096, 2.5605, 6.8643, 2.7819]);
  arrowIcon(s, 11.3096, 1.1756, 1.0432, GREEN, LEFT);
  heading(s, "GLOBAL ENERGY", [11.3096, 8.1962, 3.5303, 0.2679]);
  body(s, OBSERVE_3, [11.3096, 8.5905, 3.5303, 1.4329]);
  heading(s, "SUSTAINABLE FUTURE", [15.3808, 8.1962, 3.5303, 0.2679]);
  body(s, OBSERVE_3, [15.3808, 8.5905, 3.5303, 1.4329]);
  body(s, FASHION_6, [11.3096, 5.8276, 7.6015, 1.7964]);
}

// Slide 32 - Tapping earth's unseen power
function slide32(s) {
  title(s, "TAPPING EARTH'S UNSEEN POWER", [1.3531, 3.4301, 7.6015, 2.7819]);
  arrowIcon(s, 1.3531, 1.9365, 1.0432, GREEN, RIGHT);
  heading(s, "GLOBAL ENERGY", [1.3531, 6.9136, 3.5303, 0.2679]);
  body(s, OBSERVE_3, [1.3531, 7.3078, 3.5303, 1.4329]);
  heading(s, "SUSTAINABLE FUTURE", [5.4242, 6.9136, 3.5303, 0.2679]);
  body(s, OBSERVE_3, [5.4242, 7.3078, 3.5303, 1.4329]);
}

// Slide 33 - Three numbered points
function slide33(s) {
  heading(s, "NEW PLANET", [1.4057, 5.3322, 3.5303, 0.3214], { fontSize: 24 });
  body(s, FASHION_1, [1.4057, 5.7265, 6.3334, 1.0694]);
  num(s, "02", [1.4057, 4.5542, 1.1212, 0.7035], { fontSize: 48 });
  heading(s, "SUSTAINABLE FUTURE", [1.4057, 2.211, 4.5958, 0.3214], { fontSize: 24 });
  body(s, FASHION_1, [1.4057, 2.6052, 6.3334, 1.0694]);
  num(s, "01", [1.4057, 1.433, 1.1212, 0.7035], { fontSize: 48 });
  heading(s, "FUTURE OF POWER", [1.4057, 8.4535, 3.5303, 0.3214], { fontSize: 24 });
  body(s, FASHION_1, [1.4057, 8.8477, 6.3334, 1.0694]);
  num(s, "03", [1.4057, 7.6754, 1.1212, 0.7035], { fontSize: 48 });
}

// Slide 34 - The future of holistic design
function slide34(s) {
  s.addShape('rect', rect([13.864, 0, 6.136, 11.25], BLUE));
  title(s, "THE FUTURE OF HOLISTIC DESIGN", [7.0398, 3.0554, 6.2481, 2.7819]);
  body(s, FASHION_7, [7.0398, 6.8348, 5.8178, 2.9081]);
  heading(s, "NEXT DECADE", [14.4697, 1.7199, 2.3053, 0.3214], { fontSize: 24, color: WHITE });
  body(s, VISION_2, [14.4697, 2.1783, 5.1112, 1.1199], { color: WHITE });
  heading(s, "TRANSFORMING IDEAS ", [14.4697, 4.8189, 3.8375, 0.3214], { fontSize: 24, color: WHITE });
  body(s, VISION_2, [14.4697, 5.2774, 5.1112, 1.1199], { color: WHITE });
  heading(s, "THE VISIONARIES", [14.4697, 7.9179, 2.8242, 0.3214], { fontSize: 24, color: WHITE });
  body(s, VISION_2, [14.4697, 8.3764, 5.1112, 1.1199], { color: WHITE });
  arrowIcon(s, 7.0398, 1.5751, 1.0432, GREEN, LEFT);
}

// Slide 35 - Accessible tools for everyone
function slide35(s) {
  title(s, "ACCESSIBLE TOOLS FOR EVERYONE", [1.114, 3.3746, 5.8178, 2.7819]);
  arrowIcon(s, 1.114, 1.9898, 1.0432, GREEN, RIGHT);
  body(s, FASHION_6, [1.114, 7.0026, 5.3103, 2.5234]);
}

// Slide 36 - Clean fuels for aviation - panel
function slide36(s) {
  fadePanel(s, [1.087, 1.038, 8.913, 9.174], BEIGE);
  title(s, "CLEAN FUELS FOR AVIATION", [2.3748, 3.3746, 5.8178, 2.7819]);
  arrowIcon(s, 2.3748, 1.9898, 1.0432, GREEN, RIGHT);
  body(s, FASHION_3, [3.6792, 6.7634, 5.1686, 2.1599]);
}

// Slide 37 - Solutions for global challenges
function slide37(s) {
  title(s, "SOLUTIONS FOR GLOBAL CHALLENGES", [11.8323, 2.8431, 6.8643, 2.7819]);
  arrowIcon(s, 11.8323, 1.4582, 1.0432, GREEN, LEFT);
  heading(s, "NEW PLANET", [11.8323, 8.3429, 3.5303, 0.2668]);
  body(s, BIOENERGY_1, [11.8323, 8.7372, 6.8643, 1.0694]);
  heading(s, "SUSTAINABLE FUTURE", [11.8323, 6.2201, 3.5303, 0.2679]);
  body(s, BIOENERGY_1, [11.8323, 6.6143, 6.8643, 1.0694]);
}

// Slide 38 - Bridge to net-zero - lower
function slide38(s) {
  title(s, "BRIDGE TO NET-ZERO", [2.8314, 8.1624, 5.4512, 1.8731]);
  arrowIcon(s, 1.4434, 8.1192, 1.0432, GREEN, UP);
  body(s, FASHION_6, [8.9723, 8.1624, 9.4697, 1.4329]);
}

// Slide 39 - Three columns with dividers
function slide39(s) {
  s.addShape('line', { x: 7.0167, y: 7.5833, w: 0, h: 2.5619, line: { color: BLACK, width: 1.75 } });
  s.addShape('line', { x: 13.2863, y: 7.5833, w: 0, h: 2.5619, line: { color: BLACK, width: 1.75 } });
  heading(s, "GLOBAL ENERGY", [2.1167, 8.6977, 3.5303, 0.3214], { fontSize: 24, align: 'center' });
  body(s, OBSERVE_3, [1.6786, 9.1336, 4.4065, 1.0694], { align: 'center' });
  heading(s, "ACCESSIBLE TOOLS ", [8.3863, 8.6977, 3.5303, 0.3214], { fontSize: 24, align: 'center' });
  body(s, OBSERVE_3, [7.9482, 9.1336, 4.4065, 1.0694], { align: 'center' });
  heading(s, "SUSTAINABLE FUTURE", [14.561, 8.6977, 3.9684, 0.3214], { fontSize: 24, align: 'center' });
  body(s, OBSERVE_3, [14.342, 9.1336, 4.4065, 1.0694], { align: 'center' });
  arrowIcon(s, 16.1651, 7.5683, 0.7603, GREEN, UP);
  arrowIcon(s, 9.7713, 7.5683, 0.7603, GREEN, UP);
  arrowIcon(s, 3.5017, 7.5683, 0.7603, GREEN, UP);
}

// Slide 40 - Resilient and adaptive aesthetics
function slide40(s) {
  title(s, "RESILIENT AND ADAPTIVE AESTHETICS", [1.4355, 1.5737, 7.9603, 2.7819]);
  heading(s, "GLOBAL ENERGY", [1.4355, 8.3261, 3.5303, 0.2679]);
  body(s, OBSERVE_4, [1.4355, 8.7203, 3.5303, 1.2737], { fontSize: 16 });
  heading(s, "SUSTAINABLE FUTURE", [5.6589, 8.3261, 3.5303, 0.2679]);
  body(s, OBSERVE_4, [5.6589, 8.7203, 3.5303, 1.2737], { fontSize: 16 });
  heading(s, "CRADLE TO CRADLE", [1.4334, 5.5982, 3.5303, 0.2679]);
  body(s, OBSERVE_4, [1.4334, 5.9924, 3.5303, 1.2737], { fontSize: 16 });
  heading(s, "MIND AND SENSES", [5.6567, 5.5982, 3.5303, 0.2679]);
  body(s, OBSERVE_4, [5.6567, 5.9924, 3.5303, 1.2737], { fontSize: 16 });
  num(s, "02", [1.4358, 7.5365, 1.1212, 0.6425]);
  num(s, "01", [1.4358, 4.7896, 1.1212, 0.6425]);
  num(s, "04", [5.6567, 7.5365, 1.1212, 0.6425]);
  num(s, "03", [5.6567, 4.7896, 1.1212, 0.6425]);
}

// ------------------------------------------------------------------ build
const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30, slide31, slide32, slide33, slide34, slide35, slide36, slide37, slide38, slide39, slide40];
const GREEN_BG_SLIDES = [19, 21];   // these two slides invert to a green field
const NO_BADGE_SLIDES = [13, 20];   // layouts drawn from the second master
const DARK_BADGE_SLIDES = [4, 34];

const pptx = new pptxgen();
pptx.defineLayout({ name: 'WIDE20', width: 20, height: 11.25 });
pptx.layout = 'WIDE20';

BUILDERS.forEach((build, i) => {
  const n = i + 1;
  const s = pptx.addSlide();
  s.background = { color: GREEN_BG_SLIDES.includes(n) ? GREEN : BG };
  build(s);
  if (!NO_BADGE_SLIDES.includes(n)) slideNumber(s, n, DARK_BADGE_SLIDES.includes(n) ? BLUE_DARK : BLUE);
});

pptx.writeFile({ fileName: path.join(__dirname, '06c68e40-2433-409c-a130-152def107fd6_grok_final.pptx') })
  .then(f => console.log('wrote ' + f));
