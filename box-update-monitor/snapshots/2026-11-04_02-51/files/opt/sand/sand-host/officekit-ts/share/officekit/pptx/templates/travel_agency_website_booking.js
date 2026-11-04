/**
 * "Discovery Travel" deck — rebuilt with pptxgenjs.
 *
 * 16 slides, 13.333 x 7.5 in.  Raster artwork in the source deck is replaced by
 * native pptxgenjs shape placeholders (see the `art` section near the bottom).
 *
 * Run:  node 18ce58ca-0168-4d8a-975f-d730e5fa43d0_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const C = {
  bg: 'FAF8ED',        // lt1 / bg1 — page cream
  orange: 'F66F4D',    // accent1
  yellow: 'FFC302',    // accent2
  green: '00B46F',     // accent3
  ink: '000000',       // tx1
  slate: '44546A',     // tx2
  silver: 'E7E6E6',    // bg2
  gray: 'ADACAC',      // bg2 @ 75% luminance — muted captions
  darkGray: '404040',  // tx1 @ 75% lum / 25% offset — body copy on infographics
  mapTint: 'F5F3E9',   // faint world-map watermark
  mapKhaki: 'EAE2B4',  // bg1 @ 85% luminance — solid world map (slide 14)
  quoteGray: 'E0DED4',
  white: 'FFFFFF',
};

const MAJOR = 'Montserrat SemiBold';  // theme major latin font
const MINOR = 'Poppins';              // theme minor latin font

// Soft drop shadow used by every card / pill in the deck.
// Returns a fresh object per call — pptxgenjs rewrites shadow props in place.
const cardShadow = () => ({ type: 'outer', angle: 90, blur: 18, offset: 3, color: '000000', opacity: 0.11 });

const LOREM =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. ';
const LOREM_LONG =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. ' +
  'Aenean massa. Cum sociis natoque penatibus ipsum dolor sit amet, ';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula';

/* ---------------------------------------------------------------- helpers */

let S; // pptxgenjs ShapeType map, filled in main()

/** New slide on the deck's cream background. */
function page(pptx) {
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  return slide;
}

/** Free-form polygon from normalised (0..1) points, scaled into x/y/w/h. */
function polygon(slide, x, y, w, h, pts, o) {
  slide.addShape(S.custGeom, Object.assign({
    x, y, w, h,
    points: pts.map(([px, py]) => ({ x: px * w, y: py * h })).concat([{ close: true }]),
  }, o));
}

/** Closed free-form from absolute-inch points; the box is derived from them. */
function pathShape(slide, pts, o) {
  const xs = pts.map((p) => p[0]), ys = pts.map((p) => p[1]);
  const x = Math.min(...xs), y = Math.min(...ys);
  slide.addShape(S.custGeom, Object.assign({
    x, y, w: Math.max(...xs) - x, h: Math.max(...ys) - y,
    points: pts.map(([px, py]) => ({ x: px - x, y: py - y })).concat([{ close: true }]),
  }, o));
}

/** Text with the deck's defaults (Poppins, 12pt, black, top-anchored). */
function text(slide, str, o) {
  slide.addText(str, Object.assign({ fontFace: MINOR, fontSize: 12, color: C.ink, valign: 'top' }, o));
}

/** Heading text (Montserrat SemiBold). */
function head(slide, str, o) {
  text(slide, str, Object.assign({ fontFace: MAJOR }, o));
}

/** Rounded card with the deck's signature soft shadow. */
function card(slide, o) {
  slide.addShape(S.roundRect, Object.assign({ fill: { color: C.bg }, line: { type: 'none' }, shadow: cardShadow() }, o));
}

/** Fully rounded "pill": outlined (default) or solid orange. */
function pill(slide, x, y, w, h, solid) {
  slide.addShape(S.roundRect, {
    x, y, w, h, rectRadius: h / 2,
    fill: solid ? { color: C.orange } : { type: 'none' },
    line: solid ? { type: 'none' } : { color: C.orange, width: 1 },
    shadow: solid ? cardShadow() : undefined,
  });
}

/** Label + centered caption pill, e.g. "Adventures That Enrich Life". */
function tagPill(slide, x, y, w, h, label) {
  pill(slide, x, y, w, h, false);
  head(slide, label, { x: x + 0.114, y: y + 0.063, w: w - 0.228, h: 0.344, align: 'center', color: C.orange, lineSpacingMultiple: 1.3 });
}

/** Orange "Learn More >" button (a grouped shape in the source deck). */
function learnMore(slide, x, y) {
  slide.addShape(S.roundRect, { x, y, w: 1.702, h: 0.434, rectRadius: 0.217, fill: { color: C.orange }, line: { type: 'none' }, shadow: cardShadow() });
  slide.addShape(S.ellipse, { x: x + 1.322, y: y + 0.065, w: 0.303, h: 0.303, fill: { color: C.bg }, line: { type: 'none' } });
  head(slide, 'Learn More', { x: x + 0.006, y: y + 0.065, w: 1.316, h: 0.303, align: 'center', color: C.bg });
  slide.addShape(S.line, { x: x + 1.384, y: y + 0.217, w: 0.18, h: 0, line: { color: C.orange, width: 1, endArrowType: 'arrow' } });
}

// "<" chevron drawn as a thin outline stroke, normalised 0..1.
const CHEVRON = [[0.62, 0.00], [1.00, 0.00], [0.38, 0.50], [1.00, 1.00], [0.62, 1.00], [0.00, 0.50]];

/** Carousel "< >" control: outlined circle then a filled orange circle. */
function navArrows(slide, x, y, scale) {
  const d = 0.423 * scale;
  slide.addShape(S.ellipse, { x, y, w: d, h: d, fill: { type: 'none' }, line: { color: C.orange, width: 1 } });
  slide.addShape(S.ellipse, { x: x + 0.62 * scale, y, w: d, h: d, fill: { color: C.orange }, line: { type: 'none' } });
  const cw = 0.075 * scale, ch = 0.135 * scale;
  polygon(slide, x + d / 2 - cw * 0.6, y + (d - ch) / 2, cw, ch, CHEVRON, { fill: { color: C.orange }, line: { type: 'none' } });
  polygon(slide, x + 0.62 * scale + d / 2 - cw * 0.4, y + (d - ch) / 2, cw, ch, CHEVRON, { fill: { color: C.bg }, line: { type: 'none' }, flipH: true });
}

/** Five orange rating stars followed by a "4.5" caption. */
function stars(slide, x, y, rating) {
  for (let i = 0; i < 5; i++) {
    slide.addShape(S.star5, { x: x + i * 0.3255, y, w: 0.229, h: 0.229, fill: { color: C.orange }, line: { type: 'none' } });
  }
  text(slide, rating, { x: x + 1.652, y: y - 0.012, w: 0.541, h: 0.303, italic: true });
}

/** Boarding-pass card: "CGK - BDO / Pesawat" + flight details row. */
function flightCard(slide, x, y) {
  card(slide, { x, y, w: 3.475, h: 0.762, rectRadius: 0.057 });
  head(slide, 'CGK - BDO', { x: x + 0.132, y: y + 0.097, w: 1.531, h: 0.303 });
  text(slide, 'Pesawat', { x: x + 1.139, y: y + 0.122, w: 0.785, h: 0.252, fontSize: 9, italic: true, color: C.slate });
  const cells = [['ID-762 ( Batik Air) ', 0.132, 1.323], ['Ekonomi', 1.607, 0.785], ['13:00WIB', 2.544, 0.8]];
  cells.forEach(([t, dx, w]) => text(slide, t, { x: x + dx, y: y + 0.412, w, h: 0.252, fontSize: 9 }));
  [1.531, 2.468].forEach((dx) =>
    slide.addShape(S.line, { x: x + dx, y: y + 0.448, w: 0, h: 0.182, line: { color: C.silver, width: 1 } }));
}

/** Hotel card: "Rome, Italty" + room details row. */
function hotelCard(slide, x, y) {
  card(slide, { x, y, w: 3.895, h: 0.762, rectRadius: 0.057 });
  head(slide, 'Rome, Italty', { x: x + 0.743, y: y + 0.114, w: 1.531, h: 0.303 });
  const cells = [['1 room', 0.748, 0.608], ['2 persons', 1.432, 0.898], ['Check-in 13:00WIB', 2.406, 1.361]];
  cells.forEach(([t, dx, w]) => text(slide, t, { x: x + dx, y: y + 0.429, w, h: 0.252, fontSize: 9 }));
  [1.394, 2.368].forEach((dx) =>
    slide.addShape(S.line, { x: x + dx, y: y + 0.465, w: 0, h: 0.182, line: { color: C.silver, width: 1 } }));
}

/** "5000+ / Customers" stat card. */
function statCard(slide, x, y) {
  card(slide, { x, y, w: 2.231, h: 0.87, rectRadius: 0.06 });
  iconGlyph(slide, x + 0.18, y + 0.182, 0.505, C.ink);
  head(slide, '5000+', { x: x + 0.743, y: y + 0.07, w: 1.308, h: 0.505, fontSize: 24 });
  text(slide, 'Customers', { x: x + 0.743, y: y + 0.475, w: 1.081, h: 0.325, fontSize: 10.5, italic: true });
}

/** "Mark Medison / Your Position" name block. */
function personName(slide, x, y, w, color) {
  head(slide, 'Mark Medison', { x, y, w, h: 0.337, fontSize: 14, color });
  text(slide, 'Your Position', { x, y: y + 0.298, w: 1.361, h: 0.252, fontSize: 9, italic: true, color });
}

/** "Trip To Rome / 40% Completed" progress block. */
function tripLabel(slide, x, y, color) {
  head(slide, 'Trip To Rome  ', { x, y, w: 1.577, h: 0.337, fontSize: 14, color });
  text(slide, '40% Completed', { x, y: y + 0.337, w: 1.361, h: 0.252, fontSize: 9, color });
}

/* -------------------------------------------------- image stand-ins (art) */

/** Line-art icon stand-in: rounded square outline with a dot, mimics the
 *  small monoline travel glyphs used throughout the source deck. */
function iconGlyph(slide, x, y, size, color, rotate) {
  slide.addShape(S.roundRect, {
    x, y, w: size, h: size, rectRadius: size * 0.18, rotate,
    fill: { type: 'none' }, line: { color: color || C.ink, width: 1 },
  });
  slide.addShape(S.ellipse, {
    x: x + size * 0.33, y: y + size * 0.33, w: size * 0.34, h: size * 0.34, rotate,
    fill: { color: color || C.ink }, line: { type: 'none' },
  });
}

/** Orange map pin (teardrop + eye), stands in for the pin graphics. */
function mapPin(slide, x, y, size) {
  slide.addShape(S.teardrop, { x, y, w: size, h: size, rotate: 135, fill: { color: C.orange }, line: { type: 'none' } });
  slide.addShape(S.ellipse, { x: x + size * 0.34, y: y + size * 0.28, w: size * 0.32, h: size * 0.32, fill: { color: C.bg }, line: { type: 'none' } });
}

// Aeroplane seen from above: nose up, swept wings, tailplane. Normalised 0..1.
const PLANE = [
  [0.50, 0.00], [0.58, 0.10], [0.58, 0.34], [1.00, 0.63], [1.00, 0.75], [0.58, 0.64],
  [0.58, 0.85], [0.72, 0.96], [0.72, 1.00], [0.50, 0.94], [0.28, 1.00], [0.28, 0.96],
  [0.42, 0.85], [0.42, 0.64], [0.00, 0.75], [0.00, 0.63], [0.42, 0.34], [0.42, 0.10],
];
function plane(slide, x, y, w, h, rotate) {
  polygon(slide, x, y, w, h, PLANE, { rotate: (rotate || 0) + 90, fill: { color: C.orange }, line: { type: 'none' } });
}

// Pale seams sketched across the globe: [startAngle, endAngle, bow].  Angles are
// degrees around the disc; `bow` pulls the mid-point toward the centre so each
// curve stays inside the sphere.  Sampled as polylines in normalised 0..1 space.
const GLOBE_SEAMS = [[200, 340, 0.50], [250, 55, 0.30], [95, 200, 0.40], [10, 130, 0.25]];

/** Textured orange globe: the deck's recurring "planet" photo. */
function globe(slide, x, y, d) {
  slide.addShape(S.ellipse, { x, y, w: d, h: d, fill: { color: C.orange }, line: { type: 'none' } });
  const on = (deg) => [0.5 + 0.5 * Math.cos((deg * Math.PI) / 180), 0.5 + 0.5 * Math.sin((deg * Math.PI) / 180)];
  GLOBE_SEAMS.forEach(([a1, a2, bow]) => {
    const [x1, y1] = on(a1), [x2, y2] = on(a2);
    const mx = 0.5 + (x1 + x2 - 1) * bow, my = 0.5 + (y1 + y2 - 1) * bow;
    const pts = [];
    for (let t = 0; t <= 1.0001; t += 0.1) {
      const u = 1 - t;
      pts.push([u * u * x1 + 2 * u * t * mx + t * t * x2, u * u * y1 + 2 * u * t * my + t * t * y2]);
    }
    slide.addShape(S.custGeom, {
      x, y, w: d, h: d,
      points: pts.map(([px, py]) => ({ x: px * d, y: py * d })),
      fill: { type: 'none' }, line: { color: 'FBD4C6', width: 1 },
    });
  });
}

// Continent outlines for the world-map backdrop, normalised 0..1 over the map's
// bounding box: Eurasia+Africa, the Americas, the northern ice cap, Alaska.
const CONTINENTS = [
  [[1, 0.106], [0.824, 0.073], [0.846, 0.344], [0.684, 0.349], [0.647, 0.115], [0.62, 0.101], [0.593, 0.161], [0.591, 0.096],
    [0.532, 0.083], [0.493, 0.216], [0.463, 0.17], [0.431, 0.436], [0.564, 0.436], [0.583, 0.752], [0.522, 0.757], [0.542, 0.927],
    [0.6, 0.844], [0.686, 0.431], [0.74, 0.472], [0.755, 0.573], [0.794, 0.463], [0.863, 0.569], [0.86, 0.477], [0.9, 0.44],
    [0.892, 0.326], [0.922, 0.353], [0.929, 0.28], [0.9, 0.174], [0.956, 0.142], [0.975, 0.216]],
  [[0.091, 0.138], [0.054, 0.349], [0.1, 0.5], [0.189, 0.587], [0.174, 0.697], [0.218, 0.803], [0.245, 1], [0.324, 0.826],
    [0.343, 0.697], [0.24, 0.56], [0.169, 0.564], [0.157, 0.472], [0.123, 0.486], [0.132, 0.404], [0.172, 0.394], [0.184, 0.431],
    [0.201, 0.362], [0.147, 0.362], [0.115, 0.151]],
  [[0.456, 0.009], [0.426, 0], [0.306, 0.005], [0.272, 0.014], [0.265, 0.032], [0.248, 0.023], [0.238, 0.037], [0.218, 0.032],
    [0.233, 0.046], [0.294, 0.041], [0.304, 0.028], [0.338, 0.041], [0.336, 0.128], [0.341, 0.147], [0.355, 0.147], [0.373, 0.11], [0.417, 0.087]],
  [[0.071, 0.073], [0.056, 0.078], [0.039, 0.096], [0.02, 0.11], [0.02, 0.128], [0.005, 0.138], [0, 0.147], [0.007, 0.17],
    [0.015, 0.17], [0.027, 0.151], [0.051, 0.147], [0.051, 0.128], [0.064, 0.11], [0.078, 0.115]],
];
const MAP_ASPECT = 6.79 / 3.64;

/** World map centred in the given frame, keeping the artwork's aspect ratio. */
function worldMap(slide, fx, fy, fw, fh, color) {
  const w = Math.min(fw, fh * MAP_ASPECT), h = w / MAP_ASPECT;
  const x = fx + (fw - w) / 2, y = fy + (fh - h) / 2;
  CONTINENTS.forEach((pts) => polygon(slide, x, y, w, h, pts, { fill: { color }, line: { type: 'none' } }));
}

// Wallpaper bands sampled off the reference screenshot: [colour, top, height].
const PHONE_SCREEN = [
  ['3057AA', 0.00, 0.14], ['E52877', 0.00, 0.30], ['E63E6B', 0.30, 0.16],
  ['E96D53', 0.46, 0.18], ['F3C7B4', 0.64, 0.14], ['F8F1E9', 0.78, 0.22],
];

/** Phone mock-up (slide 9): dark bezel over the deck's gradient wallpaper. */
function phoneMock(slide, x, y, w, h) {
  slide.addShape(S.roundRect, { x, y, w, h, rectRadius: 0.35, fill: { color: '15161A' }, line: { type: 'none' }, shadow: cardShadow() });
  const sx = x + 0.09, sy = y + 0.09, sw = w - 0.18, sh = h - 0.18;
  PHONE_SCREEN.forEach(([color, fy, fh], i) =>
    slide.addShape(S.rect, {
      x: i === 0 ? sx + sw * 0.62 : sx, y: sy + fy * sh,
      w: i === 0 ? sw * 0.38 : sw, h: fh * sh,
      fill: { color }, line: { type: 'none' },
    }));
  slide.addShape(S.roundRect, { x: sx, y: sy, w: sw, h: sh, rectRadius: 0.28, fill: { type: 'none' }, line: { color: '15161A', width: 7 } });
  slide.addShape(S.roundRect, { x: x + w * 0.34, y: y + 0.11, w: w * 0.32, h: 0.13, rectRadius: 0.065, fill: { color: '15161A' }, line: { type: 'none' } });
}

/** Traveller illustration stand-in (slide 10): flat-vector tourist with a map,
 *  a bus-stop sign and luggage.  [shape, fx, fy, fw, fh, fill, radius] */
const TRAVELLER = [
  ['cloud', 0.02, 0.11, 0.92, 0.84, 'FFEDE5'],
  ['sun', 0.70, 0.02, 0.18, 0.20, 'FFD9C7'],
  ['rect', 0.19, 0.73, 0.026, 0.22, 'F67E7C'],                   // sign post
  ['ellipse', 0.13, 0.93, 0.14, 0.05, 'F67E7C'],                 // sign foot
  ['rect', 0.10, 0.375, 0.196, 0.36, 'FFFFFF'],                  // sign board
  ['roundRect', 0.37, 0.73, 0.12, 0.19, 'FFB997', 0.06],         // legs
  ['roundRect', 0.35, 0.51, 0.16, 0.23, '6C3552', 0.05],         // shorts
  ['roundRect', 0.34, 0.19, 0.18, 0.34, 'F67E7C', 0.20],         // shirt
  ['ellipse', 0.385, 0.055, 0.095, 0.135, 'FFB997'],             // face
  ['chord', 0.375, 0.030, 0.115, 0.090, '3B2430'],               // hair / beret
  ['roundRect', 0.40, 0.29, 0.20, 0.25, 'FDFCFA', 0.02],         // folded map
  ['roundRect', 0.35, 0.90, 0.09, 0.055, 'F67E7C', 0.03],        // shoes
  ['roundRect', 0.45, 0.90, 0.09, 0.055, 'F67E7C', 0.03],
  ['roundRect', 0.635, 0.42, 0.055, 0.21, 'E4756D', 0.03],       // suitcase handle
  ['roundRect', 0.57, 0.62, 0.20, 0.33, 'E4756D', 0.05],         // suitcase
  ['roundRect', 0.76, 0.73, 0.12, 0.22, 'F0907F', 0.06],         // backpack
  ['ellipse', 0.04, 0.93, 0.90, 0.055, 'F3DED8'],                // ground shadow
];
function travellerArt(slide, x, y, w, h) {
  TRAVELLER.forEach(([shape, fx, fy, fw, fh, fill, radius]) =>
    slide.addShape(S[shape], {
      x: x + fx * w, y: y + fy * h, w: fw * w, h: fh * h, rectRadius: radius,
      fill: { color: fill }, line: shape === 'rect' && fill === 'FFFFFF' ? { color: 'F79A93', width: 1.5 } : { type: 'none' },
    }));
  text(slide, 'MAP', { x: x + 0.105 * w, y: y + 0.40 * h, w: 0.19 * w, h: 0.28, fontSize: 13, bold: true, color: 'E4756D', align: 'center' });
}

/** Scooter illustration stand-in (slide 13), drawn back-to-front.
 *  Fractions were measured off the reference render. */
const SCOOTER = [
  ['cloud', 0.02, 0.02, 0.26, 0.34, C.silver],                   // exhaust puffs
  ['cloud', 0.14, 0.55, 0.30, 0.34, C.silver],
  ['roundRect', 0.66, 0.02, 0.16, 0.05, C.orange, 0.025],        // handlebar
  ['roundRect', 0.70, 0.04, 0.06, 0.20, '586572', 0.03],         // steering column
  ['ellipse', 0.79, 0.03, 0.09, 0.16, C.orange],                 // headlight
  ['roundRect', 0.23, 0.33, 0.13, 0.03, '313F4E', 0.015],        // rear rack
  ['roundRect', 0.33, 0.35, 0.03, 0.14, '313F4E', 0.015],        // rack stay
  ['ellipse', 0.26, 0.44, 0.04, 0.08, C.orange],                 // rear light
  ['roundRect', 0.66, 0.48, 0.11, 0.24, C.green, 0.04],          // front fork
  ['roundRect', 0.71, 0.48, 0.26, 0.45, '586572', 0.09],         // front fender
  ['roundRect', 0.29, 0.42, 0.42, 0.31, C.green, 0.09],          // body
  ['roundRect', 0.50, 0.31, 0.19, 0.13, '455566', 0.05],         // seat
  ['roundRect', 0.23, 0.60, 0.24, 0.14, '586572', 0.07],         // exhaust
  ['roundRect', 0.53, 0.73, 0.18, 0.08, C.orange, 0.02],         // footboard trim
  ['roundRect', 0.83, 0.58, 0.07, 0.22, C.green, 0.03],          // front suspension
];
const WHEELS = [0.304, 0.774]; // left edges; both sit on the same baseline
function scooterArt(slide, x, y, w, h) {
  SCOOTER.forEach(([shape, fx, fy, fw, fh, fill, radius]) =>
    slide.addShape(S[shape], {
      x: x + fx * w, y: y + fy * h, w: fw * w, h: fh * h, rectRadius: radius,
      fill: { color: fill }, line: { type: 'none' },
    }));
  WHEELS.forEach((fx) => {
    const fy = 0.582;
    slide.addShape(S.donut, { x: x + fx * w, y: y + fy * h, w: 0.223 * w, h: 0.415 * h, fill: { color: '313F4E' }, line: { type: 'none' } });
    slide.addShape(S.ellipse, { x: x + (fx + 0.050) * w, y: y + (fy + 0.093) * h, w: 0.123 * w, h: 0.229 * h, fill: { color: '586572' }, line: { type: 'none' } });
    slide.addShape(S.ellipse, { x: x + (fx + 0.089) * w, y: y + (fy + 0.166) * h, w: 0.045 * w, h: 0.083 * h, fill: { color: '313F4E' }, line: { type: 'none' } });
  });
}

/** Sticky note with ruled lines, pin and badge (slide 14). */
const NOTE_LINES = [
  [0.172, 0.287, 0.562, 0.030], [0.189, 0.426, 0.101, 0.030], [0.384, 0.426, 0.232, 0.030],
  [0.657, 0.426, 0.095, 0.030], [0.219, 0.604, 0.426, 0.029], [0.686, 0.604, 0.095, 0.029],
  [0.237, 0.728, 0.195, 0.029], [0.503, 0.728, 0.296, 0.029],
];
function stickyNote(slide, x, y, w, h, fill, lineColor, pin, badge) {
  slide.addShape(S.rect, { x, y, w, h, fill: { color: fill }, line: { type: 'none' }, rotate: 2 });
  NOTE_LINES.forEach(([fx, fy, fw, fh]) =>
    slide.addShape(S.rect, { x: x + fx * w, y: y + fy * h, w: fw * w, h: fh * h, fill: { color: lineColor }, line: { type: 'none' } }));
  slide.addShape(S.ellipse, { x: pin[0], y: pin[1], w: 0.256, h: 0.214, fill: { color: '455566' }, line: { type: 'none' } });
  slide.addShape(S.ellipse, { x: badge[0], y: badge[1], w: 0.591, h: 0.591, fill: { color: C.bg }, line: { type: 'none' }, shadow: cardShadow() });
  slide.addShape(S.ellipse, { x: badge[0] + 0.147, y: badge[1] + 0.14, w: 0.297, h: 0.312, fill: { color: badge[2] }, line: { type: 'none' } });
}

/** One quadrant of the slide-12 cycle: a thick arc ending in an arrow head
 *  whose tip sits at `tipAngle` (degrees, 0 = 3 o'clock, clockwise). */
function cycleArrow(slide, ring, tipAngle) {
  const { cx, cy, rOut, rIn, barb } = ring;
  const rad = (a) => (a * Math.PI) / 180;
  const at = (r, a) => [cx + r * Math.cos(rad(a)), cy + r * Math.sin(rad(a))];
  const a1 = tipAngle - 108, a2 = tipAngle - 24;
  const pts = [];
  for (let a = a1; a <= a2; a += 4) pts.push(at(rOut, a));
  pts.push(at(rOut + barb, a2), at((rOut + rIn) / 2, tipAngle), at(rIn - barb, a2));
  for (let a = a2; a >= a1; a -= 4) pts.push(at(rIn, a));
  pathShape(slide, pts, { fill: { color: C.orange }, line: { color: C.bg, width: 2.25 } });
}

/** Big decorative quotation mark (slide 8). */
function quoteMark(slide, x, y, w, h) {
  [0, w * 0.54].forEach((dx) =>
    slide.addShape(S.teardrop, {
      x: x + dx, y, w: w * 0.46, h, rotate: 200,
      fill: { color: C.quoteGray }, line: { type: 'none' },
    }));
}

/* ------------------------------------------------------- master furniture */

const NAV_ITEMS = [['Home', 3.56], ['About Us', 5.263], ['Service', 6.966], ['Contact', 8.669]];

/** Header nav + footer that the slide master paints on every slide. */
function chrome(slide, pageNo) {
  head(slide, 'Discovery Travel', { x: 0.364, y: 0.418, w: 2.5, h: 0.303 });
  NAV_ITEMS.forEach(([label, x]) =>
    head(slide, label, { x, y: 0.417, w: 1.104, h: 0.345, align: 'center', lineSpacingMultiple: 1.3 }));
  pill(slide, 11.617, 0.355, 1.352, 0.469, false);
  head(slide, 'Book Now', { x: 11.741, y: 0.417, w: 1.104, h: 0.345, align: 'center', color: C.orange, lineSpacingMultiple: 1.3 });
  text(slide, 'www.TravelTourism.com', { x: 0.364, y: 6.91, w: 3.5, h: 0.286, fontSize: 11 });
  head(slide, String(pageNo), { x: 11.89, y: 6.91, w: 1.079, h: 0.286, fontSize: 11, bold: true, align: 'right' });
}

/* -------------------------------------------------------- slide 1: hero */

function slide01(pptx) {
  const s = page(pptx);
  chrome(s, 1);
  head(s, 'Discover the Best Lovely Places', { x: 0.588, y: 1.185, w: 6.836, h: 3.433, fontSize: 66 });
  text(s, LOREM, { x: 0.588, y: 4.733, w: 5.456, h: 0.608, lineSpacingMultiple: 1.3 });
  learnMore(s, 0.588, 5.618);

  globe(s, 7.802, 2.372, 3.938);
  iconGlyph(s, 7.989, 1.961, 0.505, C.ink, 17.5);
  iconGlyph(s, 11.938, 2.998, 0.505, C.ink, 341.6);

  card(s, { x: 6.582, y: 3.569, w: 2.214, h: 1.772, rectRadius: 0.123 });
  head(s, 'Dubai moruvumi', { x: 6.847, y: 4.7, w: 1.684, h: 0.339, align: 'center', lineSpacingMultiple: 1.3 });

  card(s, { x: 9.084, y: 4.207, w: 1.369, h: 0.526, rectRadius: 0.263 });
  mapPin(s, 9.184, 4.266, 0.409);
  head(s, 'Dubai', { x: 9.56, y: 4.319, w: 0.794, h: 0.303 });
}

/* -------------------------------------------- slide 2: welcome + itinerary */

function slide02(pptx) {
  const s = page(pptx);
  worldMap(s, 0.201, -0.617, 12.75, 8.989, C.mapTint);
  chrome(s, 2);
  tagPill(s, 5.074, 1.995, 3.186, 0.469, 'Adventures That Enrich Life');
  head(s, 'Welcome to Discovery Travel', { x: 3.566, y: 2.591, w: 6.201, h: 1.582, fontSize: 44, align: 'center' });
  text(s, LOREM, { x: 4.207, y: 4.301, w: 4.918, h: 0.608, align: 'center', lineSpacingMultiple: 1.3 });

  flightCard(s, 0.724, 5.229);
  hotelCard(s, 8.714, 5.229);
  s.addShape(S.line, { x: 4.199, y: 5.61, w: 4.515, h: 0, line: { color: C.orange, width: 2, dashType: 'dash' } });
  plane(s, 6.323, 5.327, 0.687, 0.603, 8.9);
}

/* ------------------------------------------------ slide 3: our experience */

function slide03(pptx) {
  const s = page(pptx);
  chrome(s, 3);
  tagPill(s, 6.577, 1.764, 1.776, 0.469, 'Our Experience');
  head(s, 'Our Stories Have Adventures', { x: 6.466, y: 2.413, w: 5.653, h: 1.582, fontSize: 44 });
  text(s, LOREM + 'Cum sociis ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo',
    { x: 6.466, y: 4.174, w: 5.832, h: 0.87, align: 'justify', lineSpacingMultiple: 1.3 });
  learnMore(s, 6.466, 5.301);

  globe(s, 1.035, 2.49, 3.735);
  iconGlyph(s, 0.996, 1.746, 0.505, C.ink, 17.5);
  iconGlyph(s, 4.944, 2.783, 0.505, C.ink, 341.6);
  statCard(s, 3.234, 4.63);
}

/* --------------------------------------------- slide 4: destination cards */

function slide04(pptx) {
  const s = page(pptx);
  chrome(s, 4);
  head(s, 'Find Popular Destination', { x: 0.758, y: 1.238, w: 8.44, h: 0.841, fontSize: 44 });
  navArrows(s, 11.832, 1.447, 1);

  // Four tour cards; the first has a solid "Book Now" button, the rest outlined.
  [0.758, 4.204, 7.649, 11.095].forEach((cx, i) => {
    card(s, { x: cx, y: 2.51, w: 3.294, h: 4.023, rectRadius: 0.175 });
    const tx = cx + 0.167;
    head(s, 'Muntain Hiking Tour', { x: tx, y: 5.169, w: 2.601, h: 0.337, fontSize: 14 });
    text(s, 'Muntain Hiking Tour', { x: tx, y: 5.413, w: 1.709, h: 0.284, fontSize: 9, italic: true, color: C.gray, align: 'justify', lineSpacingMultiple: 1.3 });
    head(s, '$89', { x: tx, y: 6.026, w: 0.622, h: 0.337, fontSize: 14 });
    text(s, '/Person', { x: tx + 0.405, y: 6.052, w: 0.767, h: 0.284, fontSize: 9, italic: true, color: C.gray, align: 'justify', lineSpacingMultiple: 1.3 });
    const solid = i === 0;
    pill(s, tx + 1.608, 6.015, 1.352, 0.358, solid);
    head(s, 'Book Now', { x: tx + 1.732, y: 6.033, w: 1.104, h: 0.32, fontSize: 11, align: 'center', color: solid ? C.bg : C.orange, lineSpacingMultiple: 1.3 });
  });
}

/* ------------------------------------------------ slide 5: top destinations */

const FILTERS = ['London', 'Bangkok', 'England', 'Singapore', 'Italy'];
const RATING_PILLS = [[0.713, 3.328], [4.087, 3.328], [7.556, 3.328], [7.556, 5.066], [9.766, 5.066], [0.713, 5.066]];

function slide05(pptx) {
  const s = page(pptx);
  chrome(s, 5);
  head(s, 'Top Destinations', { x: 3.436, y: 1.076, w: 6.461, h: 0.841, fontSize: 44, align: 'center' });
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean',
    { x: 3.751, y: 1.918, w: 5.832, h: 0.345, align: 'center', lineSpacingMultiple: 1.3 });

  FILTERS.forEach((label, i) => {
    const x = 2.86 + i * 1.5653, solid = i === 0;
    pill(s, x, 2.46, 1.352, 0.358, solid);
    head(s, label, { x: x + 0.124, y: 2.496, w: 1.104, h: 0.286, fontSize: 11, align: 'center', color: solid ? C.bg : C.orange });
  });

  RATING_PILLS.forEach(([x, y]) => {
    pill(s, x, y, 0.678, 0.358, true);
    head(s, '3.5', { x: x + 0.118, y: y + 0.037, w: 0.442, h: 0.286, fontSize: 11, align: 'center', color: C.bg });
  });
}

/* --------------------------------------------- slide 6: trusted travellers */

function slide06(pptx) {
  const s = page(pptx);
  chrome(s, 6);
  tagPill(s, 0.773, 1.307, 3.186, 0.406, 'Adventures That Enrich Life');
  head(s, 'Trusted by Travelers', { x: 0.773, y: 1.88, w: 6.771, h: 0.841, fontSize: 44 });
  text(s, LOREM + 'Cum sociis natoque penatibus ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. ',
    { x: 0.884, y: 2.88, w: 6.595, h: 0.87, align: 'justify', lineSpacingMultiple: 1.3 });

  // Three overlapping trip chips — the first one is highlighted.
  const chips = [[5.431, 4.475, C.orange, C.bg], [8.083, 4.475, C.bg, C.ink], [6.499, 5.462, C.bg, C.ink]];
  chips.forEach(([x, y, fill, fg]) => {
    card(s, { x, y, w: 2.504, h: 0.854, rectRadius: 0.064, fill: { color: fill } });
    tripLabel(s, x + 0.679, y + 0.132, fg);
  });
}

/* ------------------------------------------------ slide 7: explore together */

function slide07(pptx) {
  const s = page(pptx);
  chrome(s, 7);
  card(s, { x: 0.936, y: 1.874, w: 3.082, h: 4.396, rectRadius: 0.147, rotate: 356 });
  card(s, { x: 3.231, y: 1.481, w: 2.907, h: 4.146, rectRadius: 0.138, rotate: 6.9 });
  flightCard(s, 1.601, 4.321);

  head(s, 'Explore the World Together', { x: 6.965, y: 1.788, w: 5.432, h: 1.582, fontSize: 44 });
  s.addShape(S.roundRect, { x: 6.965, y: 3.59, w: 0.612, h: 0.612, rectRadius: 0.055, fill: { color: C.orange }, line: { type: 'none' }, shadow: cardShadow() });
  iconGlyph(s, 7.096, 3.722, 0.349, C.bg);
  head(s, 'Explore the World Beyond Boundaries and Expectations', { x: 7.768, y: 3.647, w: 3.382, h: 0.505 });
  text(s, LOREM_LONG, { x: 6.965, y: 4.484, w: 5.234, h: 0.87, align: 'justify', lineSpacingMultiple: 1.3 });
  text(s, LOREM + 'Cum', { x: 6.965, y: 5.355, w: 5.234, h: 0.608, align: 'justify', lineSpacingMultiple: 1.3 });
}

/* ---------------------------------------------------- slide 8: testimonial */

function slide08(pptx) {
  const s = page(pptx);
  chrome(s, 8);
  head(s, 'Meet Our Travel Experts', { x: 4.635, y: 1.132, w: 8.39, h: 0.841, fontSize: 44, align: 'center' });

  // Two expert tabs above the review card.
  card(s, { x: 5.097, y: 2.408, w: 2.825, h: 1.07, rectRadius: 0.166, fill: { color: C.orange } });
  card(s, { x: 8.133, y: 2.408, w: 2.825, h: 1.07, rectRadius: 0.166 });
  personName(s, 6.072, 2.668, 1.688, C.bg);
  personName(s, 9.151, 2.668, 1.688, C.ink);

  card(s, { x: 5.097, y: 3.575, w: 7.466, h: 3.029, rectRadius: 0.137 });
  quoteMark(s, 10.279, 3.789, 1.737, 1.439);
  stars(s, 5.434, 3.824, '4.5');
  text(s, LOREM + 'Cum sociis natoque penatibus ipsum dolor sit amet, massa. Cum sociis natoque massa. Cum sociis natoque',
    { x: 5.434, y: 4.247, w: 6.566, h: 0.87, align: 'justify', lineSpacingMultiple: 1.3 });
  text(s, LOREM + 'Cum sociis natoque penatibus ipsum',
    { x: 5.434, y: 5.134, w: 6.566, h: 0.608, align: 'justify', lineSpacingMultiple: 1.3 });
  personName(s, 5.434, 5.874, 1.688, C.ink);
  navArrows(s, 10.978, 5.925, 0.753);
}

/* ------------------------------------------------- slide 9: journey showcase */

function slide09(pptx) {
  const s = page(pptx);
  worldMap(s, 5.917, 1.524, 7.615, 5.369, 'EFEDE2');
  chrome(s, 9);

  // Dashed flight paths curving around the phone.
  s.addShape(S.arc, { x: 7.076, y: 3.03, w: 3.159, h: 3.159, angleRange: [236.3, 0], line: { color: C.orange, width: 2, dashType: 'dash' }, fill: { type: 'none' } });
  s.addShape(S.arc, { x: 8.742, y: 3.303, w: 3.159, h: 3.159, rotate: 94.2, angleRange: [236.3, 0], line: { color: C.orange, width: 2, dashType: 'dash' }, fill: { type: 'none' } });
  phoneMock(s, 8.289, 1.413, 2.872, 5.59);
  [[7.385, 2.873], [11.337, 3.547], [7.482, 4.383], [11.468, 5.487]].forEach(([x, y]) => mapPin(s, x, y, 0.5));
  plane(s, 7.353, 3.662, 0.562, 0.494, 265.6);
  plane(s, 11.64, 4.634, 0.515, 0.452, 277.7);

  head(s, 'Journey Showcase', { x: 0.756, y: 2.098, w: 6.483, h: 0.841, fontSize: 44 });
  text(s, LOREM_LONG, { x: 0.756, y: 3.122, w: 5.234, h: 0.87, align: 'justify', lineSpacingMultiple: 1.3 });
  text(s, LOREM + 'Cum', { x: 0.756, y: 3.992, w: 5.234, h: 0.608, align: 'justify', lineSpacingMultiple: 1.3 });
  learnMore(s, 0.756, 5.146);
}

/* --------------------------------------------- slide 10: travel infographic */

function slide10(pptx) {
  const s = page(pptx);
  chrome(s, 10);
  travellerArt(s, 0.927, 1.635, 6.378, 4.921);

  head(s, 'Travel Infographic', { x: 7.496, y: 1.379, w: 4.66, h: 1.582, fontSize: 44 });
  text(s, LOREM + 'Cum sociis natoque penatibus ipsum', { x: 7.496, y: 3.112, w: 4.66, h: 0.87, align: 'justify', lineSpacingMultiple: 1.3 });

  [4.361, 5.349].forEach((y) => {
    s.addShape(S.roundRect, { x: 7.496, y, w: 0.612, h: 0.612, rectRadius: 0.055, fill: { color: C.orange }, line: { type: 'none' }, shadow: cardShadow() });
    iconGlyph(s, 7.62, y + 0.124, 0.365, C.bg);
    text(s, LOREM_SHORT, { x: 8.24, y: y + 0.002, w: 3.916, h: 0.608, align: 'justify', lineSpacingMultiple: 1.3 });
  });
}

/* ------------------------------------------- slide 11: serpentine timeline */

// [circle x, circle y, icon x, icon y, label x, label y, above?]
const WAVE_STOPS = [
  [2.839, 4.145, 3.110, 4.416, 2.605, 2.659, true],
  [4.503, 4.145, 4.774, 4.416, 4.253, 5.804, false],
  [6.189, 4.072, 6.460, 4.342, 5.902, 2.659, true],
  [7.792, 4.145, 8.063, 4.416, 7.550, 5.804, false],
  [9.447, 4.171, 9.717, 4.441, 9.198, 2.659, true],
];
// [x, y, rotation, start angle, end angle, thickness ratio]
const WAVE_ARCS = [
  [2.419, 3.711, 187.6, 189.26, 6.72, 0.283],
  [4.070, 3.726, 0, 179.43, 0.31, 0.265],
  [5.755, 3.593, 180, 177.71, 0.31, 0.277],
  [7.381, 3.736, 359.4, 178.76, 359.86, 0.295],
  [9.013, 3.692, 180, 177.71, 0.31, 0.277],
];

function slide11(pptx) {
  const s = page(pptx);
  chrome(s, 11);
  head(s, 'Section Infographic', { x: 3.436, y: 1.336, w: 6.461, h: 0.767, fontSize: 44, align: 'center', lineSpacingMultiple: 0.9 });

  WAVE_ARCS.forEach(([x, y, rotate, a1, a2, thick]) =>
    s.addShape(S.blockArc, {
      x, y, w: 1.902, h: 1.902, rotate, flipV: true,
      angleRange: [a1, a2], arcThicknessRatio: thick,
      fill: { color: C.orange }, line: { type: 'none' },
    }));

  WAVE_STOPS.forEach(([cx, cy, ix, iy, lx, ly, above]) => {
    s.addShape(S.ellipse, { x: cx, y: cy, w: 1.046, h: 1.046, fill: { color: C.bg }, line: { type: 'none' }, shadow: cardShadow() });
    iconGlyph(s, ix, iy, 0.505, C.orange);
    head(s, 'Title Here', { x: lx, y: ly, w: 1.53, h: 0.348, fontSize: 16, bold: true, color: C.orange, align: 'center', valign: 'bottom' });
    text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing. ',
      { x: lx - 0.44, y: ly + 0.27, w: 2.41, h: 0.564, fontSize: 11, align: 'center', lineSpacing: 18.1 });
    void above;
  });
}

/* ------------------------------------------------- slide 12: SWOT cycle */

// [label, label x, label y, label w, body x, body y, body w, align]
const SWOT = [
  ['Strengths', 2.993, 2.580, 1.348, 1.656, 2.950, 2.685, 'right'],
  ['Weakness', 8.992, 2.580, 1.403, 8.992, 2.950, 2.820, 'left'],
  ['Threats', 3.268, 5.121, 1.073, 1.656, 5.491, 2.685, 'right'],
  ['Opportunity', 8.992, 5.121, 1.631, 8.992, 5.491, 2.466, 'left'],
];
const SWOT_NODES = [[4.699, 2.712, 4.923, 2.936], [7.779, 2.712, 8.003, 2.936], [4.699, 5.253, 4.923, 5.477], [7.779, 5.253, 8.003, 5.477]];

function slide12(pptx) {
  const s = page(pptx);
  chrome(s, 12);
  head(s, 'Section Infographic', { x: 3.436, y: 1.336, w: 6.461, h: 0.767, fontSize: 44, align: 'center', lineSpacingMultiple: 0.9 });

  // Four arrow-headed quadrants forming the cycle ring, tips at 12/3/6/9 o'clock.
  const RING = { cx: 6.667, cy: 4.410, rOut: 1.790, rIn: 1.250, barb: 0.30 };
  [0, 90, 180, 270].forEach((tip) => cycleArrow(s, RING, tip));

  SWOT_NODES.forEach(([cx, cy, ix, iy]) => {
    s.addShape(S.ellipse, { x: cx, y: cy, w: 0.856, h: 0.856, fill: { color: C.bg }, line: { type: 'none' }, shadow: cardShadow() });
    iconGlyph(s, ix, iy, 0.408, C.orange);
  });

  SWOT.forEach(([label, lx, ly, lw, bx, by, bw, align]) => {
    head(s, label, { x: lx, y: ly, w: lw, h: 0.37, fontSize: 16, color: C.orange, align, wrap: false });
    text(s, LOREM_SHORT, { x: bx, y: by, w: bw, h: 0.755, fontSize: 11, color: C.darkGray, align, lineSpacingMultiple: 1.2 });
  });
}

/* -------------------------------------------------- slide 13: booking stats */

function slide13(pptx) {
  const s = page(pptx);
  chrome(s, 13);
  scooterArt(s, 4.784, 3.107, 6.853, 3.69);

  // This slide's text boxes carry a bold list style, so the title renders bold.
  head(s, 'Booking Success Travel Solutions for a Thriving', { x: 0.764, y: 1.397, w: 8.343, h: 1.286, fontSize: 44, bold: true, lineSpacingMultiple: 0.8 });
  text(s, LOREM + 'Cum sociis', { x: 0.764, y: 2.944, w: 3.841, h: 0.87, lineSpacingMultiple: 1.3 });
  statCard(s, 9.029, 1.592);

  // Two numbered callouts: 01 (orange) and 02 (cream).
  [['01', 1.644, 4.517, 0.932, C.orange, C.bg, 1.061], ['02', 7.269, 3.834, 1.112, C.bg, C.ink, 1.148]].forEach(
    ([num, x, y, numW, fill, fg, dx]) => {
      card(s, { x, y, w: 4.613, h: 1.052, rectRadius: 0.073, fill: { color: fill } });
      head(s, num, { x: x + 0.118, y: y + 0.106, w: numW, h: 0.841, fontSize: 44, color: fg, lineSpacingMultiple: 1.0 });
      text(s, LOREM_SHORT, { x: x + dx, y: y + 0.26, w: 3.435, h: 0.533, fontSize: 11, color: fg, lineSpacingMultiple: 1.2 });
    });
}

/* ------------------------------------------- slide 14: map + sticky notes */

const PROJECTS = [
  ['Your Project One', '+76,1%', 0.885, 2.614, 2.386, C.orange, C.orange],
  ['Your Project One', '+85,3%', 9.648, 2.255, 2.576, C.ink, C.green],
  ['Your Project One', '+24,6%', 7.507, 5.255, 2.569, C.ink, C.yellow],
];

function slide14(pptx) {
  const s = page(pptx);
  chrome(s, 14);
  head(s, 'Section Infographic', { x: 2.022, y: 1.047, w: 9.289, h: 0.767, fontSize: 44, align: 'center', lineSpacingMultiple: 0.9 });

  worldMap(s, 3.151, 2.562, 7.021, 3.827, C.mapKhaki);
  stickyNote(s, 3.918, 2.899, 1.441, 1.163, C.orange, C.bg, [4.472, 2.865], [3.684, 2.691, C.orange]);
  stickyNote(s, 7.540, 2.830, 1.441, 1.163, C.green, C.bg, [8.094, 2.796], [7.285, 2.562, C.green]);
  stickyNote(s, 5.749, 4.328, 1.441, 1.163, C.yellow, '455566', [6.303, 4.294], [5.525, 4.079, C.yellow]);

  PROJECTS.forEach(([label, value, x, y, w, labelColor, valueColor]) => {
    head(s, label, { x, y, w: 2.041, h: 0.37, fontSize: 16, color: labelColor, wrap: false });
    head(s, value, { x, y: y + 0.495, w, h: 0.841, fontSize: 44, color: valueColor, wrap: false });
  });
}

/* ----------------------------------------------- slide 15: customer quote */

function slide15(pptx) {
  const s = page(pptx);
  chrome(s, 15);
  globe(s, 1.428, 2.49, 3.735);
  iconGlyph(s, 1.374, 2.267, 0.505, C.ink, 17.5);
  iconGlyph(s, 5.322, 3.304, 0.505, C.ink, 341.6);

  head(s, 'A Customer Said About Us', { x: 6.344, y: 1.507, w: 4.9, h: 1.582, fontSize: 44 });
  text(s, LOREM_LONG, { x: 6.344, y: 3.237, w: 5.234, h: 0.87, align: 'justify', lineSpacingMultiple: 1.3 });
  text(s, LOREM + 'Cum', { x: 6.344, y: 4.108, w: 5.234, h: 0.608, align: 'justify', lineSpacingMultiple: 1.3 });
  stars(s, 6.464, 4.941, '4.5');
  personName(s, 6.338, 5.373, 1.871, C.ink);
  navArrows(s, 10.201, 5.437, 1);
}

/* ------------------------------------------------------- slide 16: thanks */

const CONTACT = ['+123 456 7890', 'www. Storytelling.com', '12 Your Street Name'];

function slide16(pptx) {
  const s = page(pptx);
  chrome(s, 16);
  globe(s, 7.802, 2.372, 3.938);
  iconGlyph(s, 7.989, 1.961, 0.505, C.ink, 17.5);
  iconGlyph(s, 11.938, 2.998, 0.505, C.ink, 341.6);

  head(s, 'Thank You — Adventure Awaits!', { x: 0.796, y: 1.271, w: 6.454, h: 3.433, fontSize: 66 });
  text(s, LOREM + 'Cum', { x: 0.796, y: 4.862, w: 5.183, h: 0.608, align: 'justify', lineSpacingMultiple: 1.3 });
  learnMore(s, 0.796, 5.794);

  card(s, { x: 6.432, y: 3.482, w: 2.267, h: 1.489, rectRadius: 0.103 });
  CONTACT.forEach((line, i) =>
    text(s, line, { x: 6.598, y: 3.655 + i * 0.395, w: 1.936, h: 0.352, fontSize: 11, lineSpacingMultiple: 1.5 }));
}

/* -------------------------------------------------------------------- main */

function main() {
  const pptx = new PptxGenJS();
  S = pptx.ShapeType;

  pptx.defineLayout({ name: 'WIDE_16x9', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE_16x9';
  pptx.theme = { headFontFace: MAJOR, bodyFontFace: MINOR };
  pptx.defineSlideMaster({ title: 'DISCOVERY', background: { color: C.bg } });

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
    slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16]
    .forEach((build) => build(pptx));

  return pptx.writeFile({ fileName: path.join(__dirname, '18ce58ca-0168-4d8a-975f-d730e5fa43d0_grok_final.pptx') });
}

main().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
