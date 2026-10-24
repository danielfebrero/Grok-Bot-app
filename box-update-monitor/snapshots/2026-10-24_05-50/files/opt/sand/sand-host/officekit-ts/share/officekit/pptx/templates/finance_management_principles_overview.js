/*
 * Financial Management — 30-slide deck rebuilt with pptxgenjs.
 *
 * Run with no arguments; the .pptx lands next to this file.
 *
 * Design notes
 *   * 16:9 at 10 x 5.625 in, every slide on the near-black navy background.
 *   * Photographs in the source deck are flat crops of one solid red image,
 *     so they are reproduced as red shapes (`photo`, `photoCircle`,
 *     `photoRound`).  The two product renders become `monitorMock` /
 *     `phoneMock`, built from plain grey rectangles.
 *   * The recurring furniture (kicker, three dots, page number, outline
 *     triangles) lives in the helpers at the top; each slideNN() function
 *     then only holds what makes that slide different.
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

const SLIDE_W = 10;      // inches
const SLIDE_H = 5.625;

const C = {
  navy:  '040D1C',   // background / theme tx2
  white: 'FFFFFF',
  blue:  '2BA9DF',   // theme accent1
  red:   'FF002A',   // the colour of every photo crop in the source deck
  slate: '2C2E31',   // device bezels
  silver: '979A9F',  // device housings
  chrome: 'CFD1D4'   // device bases and keys
};

const F = {
  head: 'Raleway Black',   // theme major font
  body: 'Open Sans'        // theme minor font
};

// ---------------------------------------------------------------- text helpers

/** Display heading: Raleway Black, white, top aligned. */
function title(s, x, y, w, h, text, opt = {}) {
  s.addText(text, {
    x, y, w, h, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6],
    fontFace: F.head, fontSize: opt.fontSize || 24, color: C.white
  });
}

/** Body copy: Open Sans 8.25 pt, justified, 1.5 line spacing. */
function body(s, x, y, w, h, text, opt = {}) {
  const size = opt.fontSize || 8.25;
  s.addText(text, {
    x, y, w, h, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6],
    fontFace: F.body, fontSize: size, color: opt.color || C.white,
    bold: !!opt.bold, align: opt.align || 'justify',
    lineSpacingMultiple: opt.lineSpacing || 1.5,
    bullet: opt.bullet ? { characterCode: '25AA', indent: 10 } : false
  });
}

/** Small blue eyebrow in the top-left corner of every interior slide. */
function kicker(s) {
  s.addText('- Financial Management -', {
    x: 0.2286, y: 0.1956, w: 1.4868, h: 0.2272, valign: 'top',
    margin: [7.2, 7.2, 3.6, 3.6],
    fontFace: F.head, fontSize: 7.5, color: C.blue
  });
}

/** Three white dots in the top-right corner. */
function dots(s) {
  [9.577, 9.6572, 9.7374].forEach(x =>
    s.addShape('ellipse', { x, y: 0.2757, w: 0.0486, h: 0.0486, fill: { color: C.white } }));
}

/** Page number; interior slides 2-13 sit left, 14-29 sit right-aligned. */
function pageNum(s, label, side) {
  if (side === 'left') {
    s.addText(label, {
      x: 9.6285, y: 5.2135, w: 0.3715, h: 0.3003, valign: 'top',
      margin: [7.2, 7.2, 3.6, 3.6], fontFace: F.body, fontSize: 8.25, color: C.white
    });
  } else {
    s.addText(label, {
      x: 9.4955, y: 5.2135, w: 0.372, h: 0.2995, valign: 'middle', align: 'right',
      margin: [5.4, 5.4, 2.7, 2.7], fontFace: F.body, fontSize: 8.25, color: C.white
    });
  }
}

// --------------------------------------------------------------- shape helpers

function rect(s, x, y, w, h, opt = {}) {
  s.addShape('rect', {
    x, y, w, h, fill: { color: opt.fill || C.blue }, rotate: opt.rotate,
    shadow: opt.shadow
      ? { type: 'outer', blur: opt.shadow, offset: 0, angle: 0,
          color: '000000', opacity: opt.shadowOpacity }
      : undefined
  });
}

function oval(s, x, y, w, h, opt = {}) {
  s.addShape('ellipse', {
    x, y, w, h, fill: { color: opt.fill || C.white }, rotate: opt.rotate,
    shadow: opt.shadow
      ? { type: 'outer', blur: opt.shadow, offset: 0, angle: 0,
          color: '000000', opacity: opt.shadowOpacity }
      : undefined
  });
}

/**
 * Slide 2's layout puts a big blue disc behind the red one; only its left
 * half is on the page, so a plain oval reproduces it exactly.
 */
function bluePill(s) {
  s.addShape('ellipse', { x: 6.3928, y: 0.5706, w: 4.2081, h: 4.6189, fill: { color: C.blue } });
}

/** Hollow outline triangle — the deck's signature scatter motif. */
function tri(s, x, y, w, h, rotate, lw) {
  s.addShape('triangle', {
    x, y, w, h, rotate, fill: { type: 'none' },
    line: { color: C.white, width: lw || 2.25 }
  });
}

/** Solid triangle (used once, as a tiny blue bullet marker). */
function solidTri(s, x, y, w, h, opt = {}) {
  s.addShape('triangle', { x, y, w, h, fill: { color: opt.fill || C.blue } });
}

/**
 * The recurring "double triangle" badge: a small outline triangle nested
 * inside a larger one.  Both are fixed in size; only the placement and the
 * badge's own rotation change from slide to slide.
 */
const BADGE_W = 0.3077;
const BADGE_H = 0.2653;
const BADGE_TILT = 151.638;

function doubleTri(s, x, y, rotate) {
  const cx = x + BADGE_W / 2;
  const cy = y + BADGE_H / 2;
  const rad = (rotate * Math.PI) / 180;
  const place = (dx, dy, w, h, width) => {
    const px = dx + w / 2 - cx;
    const py = dy + h / 2 - cy;
    s.addShape('triangle', {
      x: cx + px * Math.cos(rad) - py * Math.sin(rad) - w / 2,
      y: cy + px * Math.sin(rad) + py * Math.cos(rad) - h / 2,
      w, h, rotate: BADGE_TILT + rotate,
      fill: { type: 'none' }, line: { color: C.white, width }
    });
  };
  place(x + 0.0702, y + 0.0437, 0.1556, 0.1342, 1);
  place(x, y, BADGE_W, BADGE_H, 2.25);
}

/** Thin rule, arrow-headed unless told otherwise. */
function connector(s, x, y, w, h, opt = {}) {
  const arrow = opt.arrow !== false;
  s.addShape(arrow ? 'straightConnector1' : 'line', {
    x, y, w, h,
    line: { color: opt.color || C.white, width: opt.lw || 0.5,
            endArrowType: arrow ? 'triangle' : undefined }
  });
}

// ------------------------------------------------------- image stand-in shapes
// Every photo in the source is a crop of one flat red image, so a red shape of
// the same footprint is a faithful stand-in — no bitmap needed.

function photo(s, x, y, w, h) {
  s.addShape('rect', { x, y, w, h, fill: { color: C.red } });
}

function photoCircle(s, x, y, d) {
  s.addShape('ellipse', { x, y, w: d, h: d, fill: { color: C.red } });
}

function photoRound(s, x, y, w, h, r) {
  s.addShape('roundRect', { x, y, w, h, rectRadius: r, fill: { color: C.red } });
}

// The two product renders (an iMac and a pair of phones) are rebuilt from
// native shapes instead of embedded bitmaps.  The tables below hold each part
// as a fraction of the original image box, so the mock-ups scale with it.

const MONITOR_PARTS = [
  // [ x0,    y0,    x1,    y1,   colour,   cornerRadiusFrac ]
  [0.094, 0.110, 0.908, 0.583, C.slate,  0],      // display bezel
  [0.094, 0.583, 0.908, 0.693, C.silver, 0],      // chin below the display
  [0.404, 0.693, 0.598, 0.775, C.silver, 0],      // neck of the stand
  [0.284, 0.775, 0.730, 0.820, C.chrome, 0.012],  // foot and keyboard
  [0.793, 0.775, 0.948, 0.817, C.chrome, 0.03]    // mouse
];

const PHONE_PARTS = [
  [0.207, 0.045, 0.419, 0.955, C.chrome, 0.05],   // rear phone, silver back
  [0.420, 0.045, 0.795, 0.955, C.slate,  0.05]    // front phone, dark frame
];

function mock(s, parts, x, y, w, h) {
  parts.forEach(([x0, y0, x1, y1, color, r]) => {
    s.addShape(r ? 'roundRect' : 'rect', {
      x: x + x0 * w, y: y + y0 * h, w: (x1 - x0) * w, h: (y1 - y0) * h,
      rectRadius: r ? r * w : undefined, fill: { color }
    });
  });
}

function monitorMock(s, x, y, w, h) { mock(s, MONITOR_PARTS, x, y, w, h); }
function phoneMock(s, x, y, w, h) { mock(s, PHONE_PARTS, x, y, w, h); }

// ------------------------------------------------- title / closing arc ribbons
// Six nested crescents, drawn as custom geometry in fractional coordinates.
// The closing slide mirrors the same band across the page.

const ARC_MIRROR = 10.0742;

const ARCS = [
  { x: 0.5883, y: 0.5195, w: 9.527, h: 5.1062, fill: C.blue,
    path: [["M", 0.3532, 0.0226], ["C", 0.4004, 0.0079, 0.4506, 0, 0.5027, 0], ["C", 0.7196, 0, 0.9044, 0.1376, 0.9749, 0.3303], ["L", 0.9788, 0.342], ["L", 1, 0.5767], ["L", 0.9996, 0.58], ["C", 0.983, 0.6883, 0.932, 0.7853, 0.8581, 0.8593], ["L", 0.851, 0.8658], ["L", 0.4241, 1], ["L", 0.4014, 0.9965], ["C", 0.1723, 0.9496, 0, 0.7466, 0, 0.5034], ["C", 0, 0.2775, 0.1486, 0.0864, 0.3532, 0.0226], ["Z"]] },
  { x: 0.7963, y: 0.5024, w: 9.3397, h: 5.1255, fill: C.navy,
    path: [["M", 0.3603, 0.0225], ["C", 0.4085, 0.0079, 0.4597, 0, 0.5128, 0], ["C", 0.6898, 0, 0.8458, 0.0877, 0.938, 0.2211], ["L", 0.9635, 0.2622], ["L", 1, 0.6572], ["L", 0.9852, 0.6967], ["C", 0.9593, 0.7567, 0.9217, 0.8107, 0.8753, 0.8561], ["L", 0.8482, 0.8802], ["L", 0.4579, 1], ["L", 0.4094, 0.9928], ["C", 0.1758, 0.946, 0, 0.7438, 0, 0.5015], ["C", 0, 0.2764, 0.1516, 0.086, 0.3603, 0.0225], ["Z"]] },
  { x: 1.1351, y: 0.4957, w: 9.1287, h: 5.1349, fill: C.blue,
    path: [["M", 0.3686, 0.0225], ["C", 0.4179, 0.0079, 0.4703, 0, 0.5246, 0], ["C", 0.6876, 0, 0.8332, 0.0709, 0.9294, 0.1822], ["L", 0.9527, 0.2118], ["L", 1, 0.7113], ["L", 0.9859, 0.7391], ["C", 0.9536, 0.7959, 0.9101, 0.846, 0.8583, 0.8868], ["L", 0.8504, 0.8925], ["L", 0.4913, 1], ["L", 0.471, 0.9985], ["C", 0.2064, 0.9729, 0, 0.7597, 0, 0.5006], ["C", 0, 0.2759, 0.1551, 0.0859, 0.3686, 0.0225], ["Z"]] },
  { x: 1.3093, y: 0.5061, w: 8.9847, h: 5.1349, fill: C.navy,
    path: [["M", 0.3686, 0.0225], ["C", 0.4179, 0.0079, 0.4703, 0, 0.5246, 0], ["C", 0.6876, 0, 0.8332, 0.0709, 0.9294, 0.1822], ["L", 0.9527, 0.2118], ["L", 1, 0.7113], ["L", 0.9859, 0.7391], ["C", 0.9536, 0.7959, 0.9101, 0.846, 0.8583, 0.8868], ["L", 0.8504, 0.8925], ["L", 0.4913, 1], ["L", 0.471, 0.9985], ["C", 0.2064, 0.9729, 0, 0.7597, 0, 0.5006], ["C", 0, 0.2759, 0.1551, 0.0859, 0.3686, 0.0225], ["Z"]] },
  { x: 1.6556, y: 0.4534, w: 8.7484, h: 5.1255, fill: C.blue,
    path: [["M", 0.3805, 0.0225], ["C", 0.4314, 0.0079, 0.4855, 0, 0.5415, 0], ["C", 0.6887, 0, 0.8222, 0.0544, 0.9198, 0.1427], ["L", 0.9411, 0.164], ["L", 1, 0.7675], ["L", 0.9906, 0.7819], ["C", 0.9711, 0.8085, 0.949, 0.8334, 0.9245, 0.8561], ["L", 0.8958, 0.8802], ["L", 0.4835, 1], ["L", 0.4357, 0.9934], ["C", 0.1873, 0.9478, 0, 0.7449, 0, 0.5015], ["C", 0, 0.2764, 0.1601, 0.086, 0.3805, 0.0225], ["Z"]] },
  { x: 1.7676, y: 0.417, w: 8.6188, h: 5.1349, fill: C.navy,
    path: [["M", 0.386, 0.0225], ["C", 0.4376, 0.0079, 0.4925, 0, 0.5494, 0], ["C", 0.6988, 0, 0.8342, 0.0543, 0.9332, 0.1424], ["L", 0.9364, 0.1455], ["L", 1, 0.7865], ["L", 0.9734, 0.819], ["C", 0.951, 0.8437, 0.926, 0.8664, 0.8989, 0.8868], ["L", 0.8906, 0.8925], ["L", 0.5145, 1], ["L", 0.4932, 0.9985], ["C", 0.2162, 0.9729, 0, 0.7597, 0, 0.5006], ["C", 0, 0.2759, 0.1624, 0.0859, 0.386, 0.0225], ["Z"]] },
];

function arcBand(s, mirrored) {
  ARCS.forEach(a => {
    s.addShape('custGeom', {
      x: mirrored ? ARC_MIRROR - a.x - a.w : a.x,
      y: a.y, w: a.w, h: a.h,
      rotate: mirrored ? -9.562 : 9.562,
      flipH: mirrored || undefined,
      fill: { color: a.fill },
      points: a.path.map(p =>
        p[0] === 'M' ? { x: p[1] * a.w, y: p[2] * a.h, moveTo: true }
      : p[0] === 'L' ? { x: p[1] * a.w, y: p[2] * a.h }
      : p[0] === 'C' ? { x: p[5] * a.w, y: p[6] * a.h,
                         curve: { type: 'cubic', x1: p[1] * a.w, y1: p[2] * a.h,
                                  x2: p[3] * a.w, y2: p[4] * a.h } }
      : { close: true })
    });
  });
}

// --------------------------------------------------------------- slide bodies

function slide01(s) {
  arcBand(s, false);
  title(s, 4.8997, 1.5665, 4.2083, 1.3127, "FINANCIAL MANAGEMENT", { fontSize: 36 });
  tri(s, 7.2764, 0.7343, 0.381, 0.3284, 75);
  tri(s, 9.3817, 5.0361, 0.1556, 0.1342, 151.638);
  doubleTri(s, 4.8462, 4.3207, 0);
  body(s, 4.8997, 2.9256, 4.4162, 0.7258, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id vulputate volutpat. Vestibu lum eros arcu, maximus ut molestie vel, ");
  photoCircle(s, 1.7031, 1.125, 2.875);
}

function slide02(s) {
  kicker(s);
  dots(s);
  pageNum(s, "02", "left");
  bluePill(s);
  title(s, 1.1875, 1.2102, 4.2083, 1.3127, "What Is The \nImportance of \nFinancial Management ?");
  tri(s, 5.2054, 0.8195, 0.381, 0.3284, 75);
  tri(s, 0.5323, 3.2928, 0.1556, 0.1342, 151.638);
  body(s, 1.1875, 2.8125, 4.4162, 1.1423, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id vulputate volutpat. Vestibu lum eros arcu, maximus ut molestie vel, dignissim vitae elit. Vestibulum consectetur adipiscing elit et urna ac semper. Mauris finibus augue id vulputate volutpat. Vestibu lum eros arcu, maximus ut molestie vel.");
  body(s, 4.0504, 4.1649, 1.5206, 0.3092, "Read More......", { bold: true, align: "right" });
  doubleTri(s, 2.7092, 4.9109, 0);
  photoCircle(s, 6.4464, 0.4464, 4.7321);
}

function slide03(s) {
  kicker(s);
  dots(s);
  pageNum(s, "03", "left");
  rect(s, 0, 3.846, 2.9531, 1.2188, { fill: C.blue });
  photo(s, 0.5597, 1.2727, 3.284, 3.3614);
  title(s, 4.4531, 1.2727, 4.2083, 0.9088, "What Is Financial Management ?");
  body(s, 4.4531, 2.4658, 4.4162, 0.934, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id vulputate volutpat. Vestibu lum eros arcu, maximus ut molestie vel, semper. Mauris finibus augue id vulputate volutpat. Vestibu lum eros arcu, ");
  tri(s, 3.6928, 4.4146, 0.381, 0.3284, 75);
  tri(s, 9.4759, 2.1852, 0.1556, 0.1342, 151.638);
  doubleTri(s, 5.6309, 0.6238, 0);
  connector(s, 5.7665, 3.9906, 0.3916, 0);
  body(s, 6.3473, 3.8326, 1.8006, 0.5175, "Lorem ipsum dolor sit amet, consectetur adipiscing");
  body(s, 4.4531, 3.8326, 1.166, 0.3092, "Creating Real", { bold: true });
}

function slide04(s) {
  kicker(s);
  dots(s);
  pageNum(s, "04", "left");
  photo(s, 5.3888, 0.8095, 3.7453, 4.8155);
  rect(s, 6.7857, 2.6581, 3.2143, 0.3088, { fill: C.blue });
  title(s, 1.0657, 1.2534, 4.2083, 0.9088, "What Is Financial Management ?");
  connector(s, 7.0067, 2.8125, 0.3916, 0);
  body(s, 7.5084, 2.6581, 2.2897, 0.3092, "Financial Management Is Important", { bold: true, align: "left" });
  body(s, 1.0657, 2.3874, 3.72, 0.934, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id vulputate volutpat. Vestibu lum eros arcu, maximus ut molestie vel, semper. Mauris ");
  body(s, 1.0657, 3.7956, 3.4903, 0.3092, "Maintaining enough supply of funds for the organisation;", { bullet: true });
  body(s, 1.0657, 4.0309, 3.4903, 0.3092, "Optimum and efficient utilization of funds;", { bullet: true });
  body(s, 1.0657, 3.5603, 3.4903, 0.3092, "Creating real and safe investment opportunities to invest in", { bullet: true });
  tri(s, 0.3757, 4.9427, 0.381, 0.3284, 75);
  tri(s, 3.9088, 4.8134, 0.1556, 0.1342, 151.638);
  doubleTri(s, 4.369, 0.6357, 75);
}

function slide05(s) {
  kicker(s);
  dots(s);
  pageNum(s, "05", "left");
  rect(s, 0.3635, 1.6886, 1.9942, 2.7213, { fill: C.blue, rotate: 90 });
  photo(s, 0.4713, 1.4435, 2.1597, 3.4371);
  title(s, 5.3779, 0.9279, 3.5605, 1.3127, "Maintaining Enough Supply Of Funds For The Organisation");
  body(s, 5.3779, 3.3863, 3.474, 1.1423, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id vulputate volutpat. Vestibu lum eros arcu, maximus ut molestie vel, semper. Mauris finibus augue id vulputate volutpat. Vestibu lum eros arcu, ", { align: "left" });
  body(s, 5.3779, 3.0058, 1.3889, 0.366, "Discription Here -", { fontSize: 10.5, color: C.blue });
  tri(s, 4.1115, 0.529, 0.1556, 0.1342, 151.638);
  doubleTri(s, 8.6628, 2.447, 75);
  tri(s, 2.5308, 4.7164, 0.381, 0.3284, 75);
  photo(s, 2.7213, 1.044, 2.1597, 3.4371);
}

function slide06(s) {
  kicker(s);
  dots(s);
  pageNum(s, "06", "left");
  photo(s, 5.0172, 1.0546, 4.0476, 2.0327);
  title(s, 1.1548, 3.1021, 4.4197, 0.5049, "Financial Planing");
  rect(s, 6.7857, 2.9799, 3.2143, 0.3088, { fill: C.blue });
  connector(s, 7.0067, 3.1346, 0.3916, 0);
  body(s, 7.5084, 2.9799, 2.2775, 0.3092, "Financial Management Is Important", { bold: true });
  body(s, 1.1548, 3.7362, 3.72, 0.934, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id vulputate volutpat. Vestibu lum eros arcu, maximus ut molestie vel, semper. Mauris ");
  body(s, 5.181, 3.7362, 3.72, 0.934, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id vulputate volutpat. Vestibu lum eros arcu, maximus ut molestie vel, semper. Mauris ");
  body(s, 2.8552, 1.5968, 1.4894, 0.5175, "Lorem ipsum dolor sit amet, consectetur");
  connector(s, 2.2952, 1.7618, 0.3916, 0);
  body(s, 1.1548, 1.5968, 1.166, 0.3092, "Creating Real", { bold: true });
  tri(s, 3.4679, 0.5829, 0.2639, 0.2275, 75);
  tri(s, 7.0287, 1.0067, 0.1556, 0.1342, 117.422);
  doubleTri(s, 0.7104, 2.4969, 75);
}

function slide07(s) {
  kicker(s);
  dots(s);
  pageNum(s, "07", "left");
  photo(s, 0, 3.0379, 5, 2.5871);
  title(s, 1.1548, 0.8717, 4.4197, 0.9088, "Determining The Amount Of Capital Required");
  body(s, 1.1548, 1.8738, 6.0952, 0.7258, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id vulputate volutpat. Vestibu lum eros arcu, maximus ut Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id vulputate volutpat. ");
  connector(s, 6.7149, 3.6174, 0.3916, 0);
  body(s, 5.5745, 3.4628, 1.166, 0.3092, "Financial Plan", { bold: true });
  body(s, 5.5745, 3.987, 3.3065, 0.7258, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et Mauris commodo et urna ac semper. Mauris");
  tri(s, 4.8888, 4.8229, 0.381, 0.3284, 75);
  tri(s, 7.0287, 1.0067, 0.1556, 0.1342, 117.422);
  doubleTri(s, 8.5592, 2.2649, 75);
  body(s, 7.3894, 3.4628, 1.166, 0.3092, "2021", { bold: true });
}

function slide08(s) {
  kicker(s);
  dots(s);
  pageNum(s, "08", "left");
  rect(s, 3.881, 2.4486, 3.2143, 0.1743, { fill: C.blue });
  photo(s, 6.4142, 0.5054, 3.5858, 2.3783);
  title(s, 1.1699, 0.9493, 4.8153, 1.3127, "Framing Of The Organisation’s Financial Policies And Regulations.");
  body(s, 4.2108, 3.7669, 4.7163, 0.934, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id vulputate volutpat. Vestibu lum eros arcu, maximus ut Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id vulputate volutpat. ");
  body(s, 4.2108, 3.092, 1.0695, 0.4796, "85.36%", { fontSize: 15 });
  body(s, 4.2108, 3.4342, 2.3113, 0.366, "Desciption Here", { fontSize: 10.5 });
  tri(s, 5.2976, 5.0493, 0.381, 0.3284, 75);
  tri(s, 9.1281, 2.8166, 0.1556, 0.1342, 117.422);
  doubleTri(s, 0.5554, 2.4031, 75);
  photo(s, 0, 2.9051, 3.7031, 2.1393);
}

function slide09(s) {
  kicker(s);
  dots(s);
  pageNum(s, "09", "left");
  rect(s, 7.7106, 0.806, 1.6157, 4.0131, { fill: C.blue });
  title(s, 0.8648, 2.1799, 2.6025, 1.3127, "More About\nFinancial Control");
  body(s, 3.5179, 1.5663, 2.6507, 0.934, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id vulputate volutpat. ");
  body(s, 3.5179, 1.2335, 2.3113, 0.366, "About Think #1", { fontSize: 10.5, bold: true });
  body(s, 3.5179, 3.5049, 2.6507, 0.934, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id vulputate volutpat. ");
  body(s, 3.5179, 3.1722, 2.3113, 0.366, "About Think #2", { fontSize: 10.5, bold: true });
  connector(s, 3.1786, 1.2573, 0, 3.1579, { arrow: false });
  photo(s, 6.7363, 0.9795, 2.3113, 3.666);
}

function slide10(s) {
  kicker(s);
  dots(s);
  pageNum(s, "010", "left");
  title(s, 4.8768, 1.1151, 4.1646, 1.3127, "Are The Organisation’s Assets Being Used Competently?");
  tri(s, 4.8923, 5.0387, 0.381, 0.3284, 75);
  tri(s, 4.7017, 0.6381, 0.1556, 0.1342, 117.422);
  doubleTri(s, 8.0849, 2.7147, 75);
  body(s, 1.252, 3.7603, 3.1739, 0.934, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id vulputate volutpat. Donec commodo et urna ac semper. Mauris");
  body(s, 1.252, 3.4276, 2.3113, 0.366, "About Think #1", { fontSize: 10.5, bold: true });
  body(s, 5.2301, 3.7603, 1.3791, 0.3092, "Financial Control");
  connector(s, 6.573, 3.9149, 0.3916, 0);
  body(s, 7.2625, 3.7603, 1.3791, 0.3092, "Description.", { bold: true });
  body(s, 5.2301, 4.0238, 3.1739, 0.5175, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo");
  photo(s, 0, 0.7052, 4.3305, 2.1073);
}

function slide11(s) {
  kicker(s);
  dots(s);
  pageNum(s, "11", "left");
  rect(s, 0, 1.0102, 4.5784, 4.0131, { fill: C.blue });
  body(s, 0.8266, 2.0812, 1.8712, 0.7258, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Viva mus vel euismod leo. ");
  body(s, 0.8266, 1.7485, 2.3113, 0.366, "01. Descriptions.", { fontSize: 10.5, bold: true });
  body(s, 0.8266, 3.561, 1.8712, 0.7258, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Viva mus vel euismod leo. ");
  body(s, 0.8266, 3.2283, 2.3113, 0.366, "02. Descriptions.", { fontSize: 10.5, bold: true });
  title(s, 6.1436, 1.5929, 2.7943, 1.3127, "Are The Organisation’s Assets Secure?");
  body(s, 6.1436, 3.6433, 1.3791, 0.3092, "Financial Control");
  connector(s, 7.4067, 3.7979, 0.3916, 0);
  body(s, 8.0175, 3.6433, 1.3791, 0.3092, "Read More.", { bold: true });
  body(s, 6.1436, 3.9068, 3.1739, 0.5175, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo");
  photo(s, 3.0504, 1.6567, 2.5634, 2.7202);
}

function slide12(s) {
  kicker(s);
  dots(s);
  pageNum(s, "12", "left");
  rect(s, 4.1161, 2.5428, 2.981, 0.1771, { fill: C.blue });
  photo(s, 1.1892, 2.174, 3.6071, 1.8465);
  title(s, 1.1246, 0.8838, 3.1093, 0.9088, "About Financial Decision Making");
  body(s, 4.8971, 1.2219, 4.0144, 0.7258, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id vulputate volutpat. Donec commodo et urna ac semper. Mauris");
  body(s, 4.8971, 0.8892, 2.3113, 0.366, [{ text: "Financial Management " }, { text: "-", bold: true }, { text: " " }, { text: "2021", bold: true }], { fontSize: 10.5 });
  body(s, 1.1246, 4.2431, 1.3791, 0.3092, "Financial Control");
  connector(s, 2.4676, 4.3977, 0.3916, 0);
  body(s, 3.1571, 4.2431, 1.3791, 0.3092, "Description.", { bold: true });
  body(s, 1.1246, 4.5067, 3.6717, 0.5175, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo");
  body(s, 6.1922, 3.2795, 2.9037, 0.7258, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id");
  body(s, 6.1922, 3.0769, 2.3113, 0.3092, "About Think #1", { bold: true });
  body(s, 5.6748, 3.0485, 0.7061, 0.366, "85%", { fontSize: 10.5, bold: true });
  tri(s, 0.9695, 2.7876, 0.381, 0.3284, 75);
  tri(s, 4.4583, 0.5891, 0.1556, 0.1342, 117.422);
  doubleTri(s, 7.6581, 4.8707, 75);
}

function slide13(s) {
  kicker(s);
  dots(s);
  pageNum(s, "13", "left");
  rect(s, 2.6994, 3, 7.3006, 1.2273, { fill: C.blue });
  title(s, 1.1892, 1.0948, 3.1093, 0.9088, "Calculating The Capital Required");
  body(s, 5, 1.0365, 1.3791, 0.3092, "Financial Control");
  connector(s, 6.3429, 1.1911, 0.3916, 0);
  body(s, 7.0325, 1.0365, 1.3791, 0.3092, "Description.", { bold: true });
  body(s, 5, 1.3, 3.6717, 0.7258, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo", { align: "left" });
  body(s, 7.4747, 3.3786, 1.8538, 0.7258, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Viva mus vel euismod leo. ", { align: "left" });
  body(s, 7.4747, 3.1465, 2.3113, 0.3092, "About Think #1", { bold: true });
  tri(s, 0.7452, 4.5879, 0.381, 0.3284, 75);
  tri(s, 9.1552, 1.4783, 0.1556, 0.1342, 117.422);
  doubleTri(s, 7.6581, 4.8707, 75);
  photo(s, 0.8797, 2.4844, 2.9908, 2.3125);
  photo(s, 4.1059, 2.4844, 2.9908, 2.3125);
}

function slide14(s) {
  kicker(s);
  dots(s);
  pageNum(s, "14", "right");
  rect(s, 6.369, 3.353, 3.631, 1.3063, { fill: C.blue });
  oval(s, 0.6056, 1.172, 3.3635, 3.3635, { fill: C.navy, shadow: 35, shadowOpacity: 0.3 });
  title(s, 4.2116, 1.1844, 3.1093, 0.9088, "Formation Of Capital Structure");
  body(s, 4.2116, 2.2059, 4.0272, 0.934, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo.  Donec commodo sit Vivamus vel euismod leo amet, consectetur adipiscing elit.. Donec commodo consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo sit", { align: "left" });
  body(s, 6.7474, 3.6031, 1.3791, 0.3092, "Financial Control");
  connector(s, 7.9306, 3.7577, 0.2709, 0);
  body(s, 8.3267, 3.6031, 1.3791, 0.3092, "Detail.", { bold: true });
  body(s, 6.7474, 3.8633, 2.8782, 0.5175, "Lorem ipsum dolor sit amet, cons ectetur adi piscing elit. Lorem ipsum dolor sit");
  tri(s, 8.2184, 1.2049, 0.1556, 0.1342, 117.422);
  doubleTri(s, 4.8371, 3.598, 75);
  photoCircle(s, 0.8348, 1.4012, 2.9051);
  tri(s, 1.3542, 3.961, 0.381, 0.3284, 75);
}

function slide15(s) {
  kicker(s);
  dots(s);
  pageNum(s, "15", "right");
  rect(s, 0, 4.3292, 2.9167, 0.3684, { fill: C.blue });
  title(s, 0.9719, 1.3684, 3.1093, 0.9088, "Investing The Capital Here");
  body(s, 0.9719, 2.4865, 3.1068, 1.1423, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo.  Donec commodo sit Vivamus vel euismod leo amet, consectetur adipiscing elit.. Donec commodo consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo sit", { align: "left" });
  body(s, 3.4497, 4.2583, 1.3791, 0.3092, "Financial Control");
  connector(s, 4.6924, 4.413, 0.2709, 0);
  body(s, 5.029, 4.2583, 1.3791, 0.3092, "Detail.", { bold: true });
  body(s, 3.4497, 4.5066, 4.8837, 0.4953, "Lorem ipsum dolor sit amet, cons ectetur adi piscing elit. Lorem ipsum dolor si Lorem ipsum dolor sit amet, cons ectetur adi piscing t", { align: "left" });
  photo(s, 8.2133, 1.0119, 1.7867, 2.9405);
  photo(s, 6.3388, 1.0119, 1.7867, 2.9405);
  photo(s, 4.4643, 1.0119, 1.7867, 2.9405);
}

function slide16(s) {
  kicker(s);
  dots(s);
  pageNum(s, "16", "right");
  rect(s, 1.9609, 3.581, 1.125, 0.1065, { fill: C.blue });
  body(s, 5.5189, 3.0508, 1.3791, 0.3092, "Financial Control");
  connector(s, 6.8166, 3.2054, 0.2709, 0);
  body(s, 7.1532, 3.0508, 1.3791, 0.3092, "Detail.", { bold: true });
  body(s, 5.5189, 3.4755, 3.6866, 0.934, "Lorem ipsum dolor sit amet, cons ectetur adi piscing elit. Lorem ipsum dolor si Lorem ipsum dolor sit amet, cons ectetur adi piscing amet, cons ectetur adi piscing elit. Lorem dolor si Lorem ipsum dolor sit amet, cons ectetur adi piscing amet, cons ");
  title(s, 5.5189, 1.2629, 2.6061, 1.3127, "Effective Management Of Money");
  tri(s, 8.5496, 1.7787, 0.381, 0.3284, 75);
  tri(s, 6.6434, 4.9409, 0.1556, 0.1342, 117.422);
  doubleTri(s, 0.5791, 4.8899, 75);
  body(s, 0.3394, 3.4958, 1.3791, 0.3092, "Financial Control");
  body(s, 0.3394, 3.7423, 1.8973, 0.5175, "Lorem ipsum dolor sit amet, cons ectetur adi piscing");
  photo(s, 0, 0.8606, 2.4772, 2.1684);
  photo(s, 2.6625, 0.8606, 2.3931, 3.9477);
}

function slide17(s) {
  kicker(s);
  dots(s);
  pageNum(s, "17", "right");
  photo(s, 5.1211, 2.5897, 4.8789, 1.825);
  title(s, 1.1757, 1.1205, 4.4611, 0.9088, "Why Is Financial Management Important ? ");
  body(s, 1.1757, 3.8225, 1.3791, 0.3092, "First Thing Here");
  connector(s, 2.4184, 3.9772, 0.2709, 0);
  body(s, 2.755, 3.8225, 1.3791, 0.3092, "Detail.", { bold: true });
  body(s, 1.1757, 2.5261, 3.254, 1.1423, "Lorem ipsum dolor sit amet, cons ectetur adi piscing elit. Lorem ipsum dolor si Lorem ipsum dolor sit amet, cons ectetur adi piscing amet, cons ectetur adi piscing elit. Lorem dolor si Lorem ipsum dolor sit amet, cons ectetur adi piscing amet, cons ");
  body(s, 1.1757, 4.2117, 1.3791, 0.3092, "Second Thing Here");
  connector(s, 2.4184, 4.3663, 0.2709, 0);
  body(s, 2.755, 4.2117, 1.3791, 0.3092, "Detail.", { bold: true, color: C.blue });
  tri(s, 8.9122, 2.3619, 0.381, 0.3284, 75);
  tri(s, 6.2684, 0.8452, 0.1556, 0.1342, 117.422);
  doubleTri(s, 0.5791, 4.8899, 75);
  body(s, 6.9842, 1.3801, 1.672, 0.5175, "About Creative Presentation Templates", { align: "left" });
  connector(s, 6.6406, 1.513, 0.2344, 0, { arrow: false, color: C.blue, lw: 3 });
}

function slide18(s) {
  kicker(s);
  dots(s);
  pageNum(s, "18", "right");
  rect(s, 1.2262, 2.6602, 6.0275, 0.2151, { fill: C.blue });
  tri(s, 0.6791, 5.0465, 0.1556, 0.1342, 117.422);
  doubleTri(s, 5.42, 0.88, 75);
  title(s, 1.174, 1.1956, 4.4611, 1.3127, "Break Slide\nPresentation ", { fontSize: 36 });
  body(s, 1.174, 3.3346, 3.1068, 1.1423, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo.  Donec commodo sit Vivamus vel euismod leo amet, consectetur adipiscing elit.. Donec commodo consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo sit");
  body(s, 5.5738, 1.6169, 1.3791, 0.3092, "Financial Managemetn");
  connector(s, 7.1104, 1.7716, 0.2709, 0);
  body(s, 7.5388, 1.6169, 1.3791, 0.3092, "Next Slide", { bold: true });
  tri(s, 9.0709, 2.3188, 0.381, 0.3284, 75);
  photo(s, 5, 2.8125, 3.7738, 1.7695);
}

function slide19(s) {
  kicker(s);
  dots(s);
  pageNum(s, "19", "right");
  title(s, 1.0648, 3.3526, 3.2905, 1.3127, "Why Study Financial Management?");
  body(s, 6.5785, 3.9492, 2.3382, 0.7258, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo.  Donec commodo sit Vivamus vel");
  body(s, 6.5785, 3.6705, 2.3382, 0.3282, "Description #02", { fontSize: 9, bold: true });
  body(s, 3.9278, 3.9492, 2.3382, 0.7258, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo.  Donec commodo sit Vivamus vel");
  body(s, 3.9278, 3.6705, 2.3382, 0.3282, "Description #01", { fontSize: 9, bold: true });
  body(s, 3.9278, 3.3429, 2.3382, 0.3567, "+518", { fontSize: 10.13, bold: true });
  body(s, 6.5785, 3.3429, 2.3382, 0.3567, "89%", { fontSize: 10.13, bold: true });
  tri(s, 0.6791, 5.0465, 0.1556, 0.1342, 117.422);
  doubleTri(s, 4.2015, 0.4114, 75);
  tri(s, 9.0709, 2.3188, 0.381, 0.3284, 75);
  photo(s, 1.0833, 0.9435, 2.5134, 1.7695);
  photo(s, 3.7433, 0.9435, 2.5134, 1.7695);
  photo(s, 6.4033, 0.9435, 2.5134, 1.7695);
}

function slide20(s) {
  kicker(s);
  dots(s);
  pageNum(s, "20", "right");
  rect(s, 8.4739, 3.6916, 1.5261, 0.2995, { fill: C.blue });
  title(s, 4.5549, 1.0548, 3.2905, 1.3127, "Strategic vs. Tactical Financial Management");
  body(s, 4.5549, 2.3962, 3.6951, 0.7258, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo.  Donec commodo sit Vivamus vel euismod leo amet, consectetur adipiscing elit.. Donec commodo");
  body(s, 2.8613, 3.6768, 1.3791, 0.3092, "Strategic Financial ");
  connector(s, 4.0515, 3.8314, 0.2709, 0);
  body(s, 4.3602, 3.6768, 0.8141, 0.3092, "Detail", { bold: true });
  body(s, 5.5126, 3.6768, 1.3791, 0.3092, "Tactical Financial");
  connector(s, 6.62, 3.8314, 0.2709, 0);
  body(s, 6.956, 3.6768, 0.8141, 0.3092, "Detail", { bold: true });
  body(s, 2.8613, 3.9798, 2.3329, 0.5175, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel");
  body(s, 5.5126, 4.0058, 2.3329, 0.5175, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel");
  tri(s, 5.1164, 5.1465, 0.1556, 0.1342, 117.422);
  doubleTri(s, 7.6223, 0.8691, 75);
  tri(s, 0.9258, 1.7587, 0.381, 0.3284, 75);
  photoCircle(s, 8.835, 0.7455, 1.9019);
  photoCircle(s, 2.0746, 1.0548, 2.1343);
  photoCircle(s, 0.7257, 3.1147, 1.7527);
}

function slide21(s) {
  kicker(s);
  dots(s);
  pageNum(s, "21", "right");
  photo(s, 5.1477, 2.8125, 3.8404, 1.7695);
  title(s, 1.0833, 1.2927, 4.2738, 0.9088, "Objectives of Financial Management");
  body(s, 5.55, 1.2927, 3.25, 0.934, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo.  Donec commodo sit Vivamus vel euismod leo amet, consectetur adipiscing elit.. Donec commodo sit amet, consectetur adipiscing elit. ");
  tri(s, 0.6791, 5.0465, 0.1556, 0.1342, 117.422);
  doubleTri(s, 4.2961, 2.0466, 75);
  tri(s, 8.8137, 2.6483, 0.381, 0.3284, 75);
  photo(s, 1.0833, 2.8125, 3.8404, 1.7695);
}

function slide22(s) {
  kicker(s);
  dots(s);
  pageNum(s, "22", "right");
  title(s, 5, 1.1529, 3.1963, 0.9088, "Tracking Liquidity And Cash Flow");
  body(s, 5, 3.2588, 3.8629, 1.1423, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo.  Donec commodo sit Vivamus vel euismod leo amet, consectetur adipiscing elit.. Donec commodo sit amet, consectetur adipiscing elit. consectetur adipiscing elit. Vivamus vel euismod leo.  Donec commodo sit Vivamus vel euismod leo amet, ");
  body(s, 5, 2.7982, 3.4903, 0.3092, "About Our Cashflow", { bullet: true });
  body(s, 5, 2.4561, 3.4903, 0.3092, "First Financial Tracking Liquidity", { bullet: true });
  tri(s, 0.7988, 5.2323, 0.1556, 0.1342, 117.422);
  doubleTri(s, 8.7091, 2.0306, 75);
  tri(s, 6.612, 4.8311, 0.381, 0.3779, 75);
  photo(s, 0, 0.8095, 4.3391, 4.1778);
}

function slide23(s) {
  kicker(s);
  dots(s);
  pageNum(s, "23", "right");
  rect(s, 0.75, 2.9664, 3.8274, 1.3498, { fill: C.blue });
  photo(s, 3.75, 2.5952, 6.25, 2.1895);
  title(s, 1.131, 1.141, 3.7411, 0.9088, "Developing Financial Scenarios");
  body(s, 1.131, 3.2393, 1.3791, 0.3092, "Tactical Financial");
  connector(s, 2.2384, 3.3939, 0.2709, 0);
  body(s, 2.5744, 3.2393, 0.8141, 0.3092, "Detail", { bold: true });
  body(s, 1.131, 3.5683, 2.3329, 0.5175, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel");
  body(s, 4.872, 1.141, 3.753, 0.934, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo.  Donec commodo sit Vivamus vel euismod leo amet, conse ctetur adipiscing elit.. Donec commodo sit amet, consectetur adipiscing elit. consectetur adipiscing");
  tri(s, 0.6791, 5.0465, 0.1556, 0.1342, 117.422);
  doubleTri(s, 3.5962, 0.9648, 75);
  tri(s, 6.4766, 4.6205, 0.381, 0.3284, 75);
}

function slide24(s) {
  kicker(s);
  dots(s);
  pageNum(s, "24", "right");
  photo(s, 3.1369, 3.0446, 2.771, 2.5804);
  title(s, 3.4167, 0.8642, 3.5238, 0.9088, "Scope of Financial Management");
  body(s, 3.4167, 1.9263, 5.3698, 0.7258, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo.  Donec commodo sit Vivamus vel euismod leo amet, conse ctetur adipiscing elit.. Donec commodo sit amet, consectetur adipiscing elit. consectetur adipiscing");
  body(s, 6.3375, 3.0165, 1.3791, 0.3092, "Planing Here");
  connector(s, 7.5811, 3.1711, 0.2709, 0);
  body(s, 8.146, 3.0165, 0.8141, 0.3092, "Detail", { bold: true });
  body(s, 6.3375, 3.4428, 1.3791, 0.3092, "Budgeting ");
  connector(s, 7.5811, 3.5974, 0.2709, 0);
  body(s, 8.146, 3.4428, 0.8141, 0.3092, "Detail", { bold: true });
  body(s, 6.3375, 3.9576, 2.673, 0.7258, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo.  Donec commodo sit Vivamus vel");
  tri(s, 0.6791, 5.0465, 0.1556, 0.1342, 117.422);
  doubleTri(s, 8.4719, 1.0761, 75);
  tri(s, 5.8257, 4.9831, 0.381, 0.3284, 75);
  photo(s, 0, 0.8095, 2.9286, 4.1071);
}

function slide25(s) {
  kicker(s);
  dots(s);
  pageNum(s, "25", "right");
  photo(s, 5.76, 1.3851, 3.0774, 2.9315);
  title(s, 1.1626, 1.3571, 3.5238, 0.9088, "Managing And Assessing Risk");
  rect(s, 8.5156, 3.9893, 0.6436, 0.6131, { fill: C.blue });
  body(s, 1.1626, 2.667, 3.9419, 0.5175, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo.  Donec commodo sit Vivamus vel euismod", { bullet: true });
  body(s, 1.1626, 3.2897, 3.9419, 0.5175, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo.  Donec commodo sit Vivamus vel euismod", { bullet: true });
  body(s, 1.1626, 3.9124, 3.9419, 0.5175, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo.  Donec commodo sit Vivamus vel euismod", { bullet: true });
  tri(s, 2.8467, 4.8898, 0.1556, 0.1342, 117.422);
  doubleTri(s, 0.7368, 1.1099, 75);
  tri(s, 5.5696, 1.6347, 0.381, 0.3284, 75);
}

function slide26(s) {
  kicker(s);
  dots(s);
  pageNum(s, "26", "right");
  rect(s, 0, 4.5923, 3.5485, 0.3772, { fill: C.blue });
  photoCircle(s, 1.1715, 2.381, 1.7527);
  title(s, 1.1514, 0.9715, 4.0203, 0.9088, "Functions of Financial Management");
  body(s, 5.2385, 0.9715, 3.2438, 0.934, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo.  Donec commodo sit Vivamus vel euismod leo amet, conse ctetur adipiscing elit.. Donec commodo sit amet, consectetur adipiscing");
  body(s, 0.5682, 4.6241, 1.3791, 0.3092, "Planing Here");
  connector(s, 1.6546, 4.7773, 0.4046, 0);
  body(s, 2.3768, 4.6241, 0.8141, 0.3092, "Detail", { bold: true });
  body(s, 3.832, 4.5466, 4.765, 0.5175, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo.  Donec commodo sit Vivamus vel euismod leo amet, conse ctetur");
  tri(s, 4.9984, 3.8731, 0.1556, 0.1342, 117.422);
  doubleTri(s, 9.2795, 2.8628, 75);
  tri(s, 1.3759, 2.2892, 0.381, 0.3284, 75);
  photoCircle(s, 7.0758, 2.3136, 1.7527);
  photoCircle(s, 5.1077, 2.3136, 1.7527);
  photoCircle(s, 3.1396, 2.3473, 1.7527);
}

function slide27(s) {
  kicker(s);
  dots(s);
  pageNum(s, "27", "right");
  rect(s, 6.9067, 2.3697, 3.0933, 2.4344, { fill: C.blue });
  monitorMock(s, 5.1716, 0.7781, 4.1194, 4.0688);
  title(s, 1.1514, 1.269, 4.0203, 0.9088, "Financial Management for Startups");
  body(s, 1.1514, 4.0484, 1.3791, 0.3092, "Planing Here");
  connector(s, 2.4171, 4.2031, 0.2709, 0);
  body(s, 2.9821, 4.0484, 0.8141, 0.3092, "Detail", { bold: true });
  body(s, 3.7962, 4.0484, 2.0023, 0.7258, "Lorem ipsum dolor sit amet, cons ectetur adipiscing elit. Vivamus vel euismod leo.  ");
  body(s, 1.1514, 2.3697, 3.9046, 1.1423, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo.  Donec commodo sit Vivamus vel euismod leo amet, conse ctetur adipiscing elit.. Donec commodo sit amet, consectetur adipiscing consectetur adipiscing elit. Vivamus vel euismod leo.  Donec commodo sit Vivamus vel euismod leo");
  tri(s, 1.8811, 5.076, 0.1556, 0.1342, 117.422);
  doubleTri(s, 0.7368, 1.1099, 75);
  tri(s, 5.608, 1.1019, 0.381, 0.3284, 75);
  photo(s, 5.6641, 1.3359, 3.1677, 1.8359);
}

function slide28(s) {
  kicker(s);
  dots(s);
  pageNum(s, "28", "right");
  rect(s, 0, 3.3058, 2.6567, 2.3192, { fill: C.blue });
  phoneMock(s, 0.2286, 0.8095, 4.006, 4.006);
  title(s, 4.0799, 1.3026, 3.2591, 0.5049, "Maximizing Profits");
  body(s, 6.4624, 2.1942, 2.5897, 1.1423, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo.  Donec commodo sit Vivamus vel euismod leo amet, conse ctetur adipiscing elit.. Donec commodo sit amet, ");
  body(s, 4.0799, 2.1942, 1.3791, 0.3092, "Planing Here");
  connector(s, 5.1041, 2.3488, 0.2709, 0);
  body(s, 5.5144, 2.1942, 0.8141, 0.3092, "Detail", { bold: true });
  body(s, 4.0799, 2.543, 2.053, 0.7258, "Lorem ipsum dolor sit amet, consect etur adipiscing elit. Viva mus vel euismod leo.  Donec .");
  rect(s, 5.1041, 3.872, 4.8959, 0.6783, { fill: C.navy, shadow: 20, shadowOpacity: 0.6 });
  body(s, 5.6552, 3.9966, 0.6733, 0.4039, "80.39", { fontSize: 12, bold: true, align: "left" });
  solidTri(s, 5.5144, 4.1038, 0.1096, 0.0945, { fill: C.blue });
  body(s, 6.232, 3.9916, 3.1056, 0.5175, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo.  Donec");
  tri(s, 3.859, 4.2759, 0.1556, 0.1342, 117.422);
  doubleTri(s, 0.4476, 2.426, 75);
  tri(s, 8.5568, 1.4469, 0.381, 0.3284, 75);
  photoRound(s, 1.7153, 1.069, 1.6093, 3.4813, 0.1787);
}

function slide29(s) {
  kicker(s);
  dots(s);
  pageNum(s, "29", "right");
  rect(s, 6.6169, 1.3618, 2.4314, 3.6656, { fill: C.blue });
  photo(s, 5.8147, 0.1317, 2.7575, 4.237);
  title(s, 1.125, 1.3008, 4.125, 0.9088, "Personal Financial Planners Here");
  body(s, 1.125, 3.4407, 3.9151, 1.1423, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo.  Donec commodo sit Vivamus vel euismod leo amet, conse ctetur adipiscing elit.. Donec commodo sit amet, consectetur adipiscing consectetur adipiscing elit. Vivamus vel euismod leo.  Donec commodo sit Vivamus vel euismod leo");
  rect(s, 0, 2.5163, 4.2, 0.6783, { fill: C.navy, shadow: 20, shadowOpacity: 0.6 });
  body(s, 1.1498, 2.5633, 0.7539, 0.4039, "50%", { fontSize: 12, bold: true, color: C.blue, align: "center" });
  body(s, 2.1997, 2.5633, 0.7539, 0.4039, "75%", { fontSize: 12, bold: true, align: "center" });
  body(s, 1.125, 2.8033, 0.8035, 0.3092, "Aspect #1", { color: C.blue, align: "center" });
  body(s, 2.1748, 2.8033, 0.8035, 0.3092, "Aspect #2", { align: "center" });
  body(s, 3.2105, 2.5633, 0.7539, 0.4039, "85%", { fontSize: 12, bold: true, align: "center" });
  body(s, 3.1857, 2.8033, 0.8035, 0.3092, "Aspect #3", { align: "center" });
  tri(s, 3.9873, 5.1962, 0.1556, 0.1342, 117.422);
  doubleTri(s, 2.8639, 0.6354, 75);
  tri(s, 5.7013, 4.1448, 0.381, 0.3284, 75);
}

function slide30(s) {
  arcBand(s, true);
  title(s, 1.1497, 1.8166, 4.2083, 1.1107, "THANKS FOR\nWATCHING", { fontSize: 30 });
  tri(s, 0.5815, 0.1953, 0.381, 0.3284, 75);
  tri(s, 5.431, 4.3863, 0.1556, 0.1342, 151.638);
  doubleTri(s, 1.0962, 4.3207, 0);
  body(s, 1.1497, 3.1566, 3.2103, 0.7258, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id");
  photoCircle(s, 5.0922, 1.125, 2.875);
}


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
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'DECK';
  pptx.theme = { headFontFace: F.head, bodyFontFace: F.body };
  pptx.title = 'Financial Management';

  SLIDES.forEach(draw => {
    const s = pptx.addSlide();
    s.background = { color: C.navy };
    draw(s);
  });

  return pptx.writeFile({ fileName: path.join(__dirname, "0be36227-c2c5-472c-8f44-409431da794f_grok_final.pptx") });
}

build().then(f => console.log('wrote ' + f)).catch(e => { console.error(e); process.exit(1); });
