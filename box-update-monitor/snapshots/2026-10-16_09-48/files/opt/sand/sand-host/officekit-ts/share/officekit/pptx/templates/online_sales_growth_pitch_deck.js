'use strict';
/*
 * "Optimizing Online Sales for Maximum Growth" — 20-slide pitch deck (13.333" x 7.5").
 * Rebuilt with pptxgenjs only. Raster photos in the original are replaced by flat
 * light-grey "[image]" rectangles at the same position and size.
 */
const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ palette */
const C = {
  green: '8FE916',      // theme accent1
  purple: 'B782F5',     // theme accent2
  greenLt: 'E9FBD0',    // accent1 lumMod 20% / lumOff 80%
  purpleLt: 'F1E6FD',   // accent2 lumMod 20% / lumOff 80%
  purpleMid: 'C7A5F9',  // upper wave band on purple cards
  grey85: 'D9D9D9',     // white lumMod 85%
  grey95: 'F2F2F2',     // white lumMod 95%
  black: '000000',
  white: 'FFFFFF',
  photo: 'E0E0E0'       // stand-in for the deck's photo placeholders
};
const FONT = 'Syne';

/* -------------------------------------------------------------- primitives */

// Every text frame in the source is a plain top-anchored auto-fit text box.
function text(s, str, o) {
  s.addText(str, Object.assign({
    fontFace: FONT, fontSize: 18, color: C.black, valign: 'top',
    isTextBox: true, autoFit: true
  }, o));
}

// roundRect whose corner radius is given as the OOXML "adj" value (1/100000).
function card(s, x, y, w, h, fill, adj) {
  s.addShape('roundRect', {
    x, y, w, h, fill: { color: fill },
    rectRadius: ((adj === undefined ? 8033 : adj) / 100000) * Math.min(w, h)
  });
}

// Fully-rounded capsule (adj = 50000).
function pill(s, x, y, w, h, fill, rotate) {
  const o = { x, y, w, h, fill: { color: fill }, rectRadius: Math.min(w, h) / 2 };
  if (rotate) o.rotate = rotate;
  s.addShape('roundRect', o);
}

// Outlined capsule used for the small "Project" / "Income" tags.
function tag(s, x, y, w, h, lineColor) {
  s.addShape('roundRect', {
    x, y, w, h, fill: { type: 'none' },
    line: { color: lineColor, width: 1 }, rectRadius: Math.min(w, h) / 2
  });
}

// Progress meter: full-width track plus a shorter value capsule on top.
function meter(s, x, y, w, h, filled, trackColor, valueColor) {
  pill(s, x, y, w, h, trackColor);
  pill(s, x, y, filled, h, valueColor);
}

// Replacement for a photo placeholder from the original deck.
function photo(s, x, y, w, h, adj) {
  card(s, x, y, w, h, C.photo, adj === undefined ? 5323 : adj);
  text(s, '[image]', {
    x, y: y + h / 2 - 0.18, w, h: 0.36, fontSize: 11, color: '8A8A8A', align: 'center'
  });
}

/* ------------------------------------------------------------------- icons */

// Eight-lobed "flower" blob that appears all over the deck.
// NOTE: custGeom points are measured from the shape's own top-left corner.
function blob(s, x, y, size, color) {
  const R = size / 2, pts = [];
  for (let i = 0; i <= 128; i++) {
    const t = (2 * Math.PI * i) / 128;
    // Flattened cosine: broad rounded lobes separated by narrow notches.
    const u = Math.cos(8 * t - 0.157);
    const r = R * (0.945 + 0.067 * Math.sign(u) * Math.pow(Math.abs(u), 0.35));
    pts.push({ x: R + r * Math.cos(t), y: R + r * Math.sin(t) });
  }
  pts.push({ close: true });
  s.addShape('custGeom', { x, y, w: size, h: size, points: pts, fill: { color } });
}

// Thick arrow drawn as three rounded strokes. It points right at deg = 0;
// the deck also uses deg = -45 for the north-east variant.
function arrow(s, x, y, size, color, deg) {
  deg = deg || 0;
  const t = size * 0.135, cx = x + size / 2, cy = y + size / 2, rad = (deg * Math.PI) / 180;
  const strokes = [                       // endpoints relative to the arrow centre
    [[-0.48, 0.00], [0.44, 0.00]],        // shaft
    [[0.44, 0.00], [-0.06, -0.44]],       // upper chevron arm
    [[0.44, 0.00], [-0.06, 0.44]]         // lower chevron arm
  ];
  strokes.forEach(function (seg) {
    const ax = seg[0][0] * size, ay = seg[0][1] * size;
    const bx = seg[1][0] * size, by = seg[1][1] * size;
    const len = Math.hypot(bx - ax, by - ay);
    const ang = (Math.atan2(by - ay, bx - ax) * 180) / Math.PI;
    const mx = (ax + bx) / 2, my = (ay + by) / 2;
    pill(s,
      cx + mx * Math.cos(rad) - my * Math.sin(rad) - (len + t) / 2,
      cy + mx * Math.sin(rad) + my * Math.cos(rad) - t / 2,
      len + t, t, color, ang + deg);
  });
}

// Award rosette: scalloped medal over two ribbon tails. `bg` = card colour
// showing through the medal's centre.
function iconAward(s, x, y, w, h, color, bg) {
  const d = Math.min(w, h * 0.72);
  s.addShape('custGeom', {                                 // left ribbon tail
    x, y, w, h, fill: { color },
    points: [
      { x: w * 0.26, y: h * 0.58 }, { x: w * 0.44, y: h * 0.70 },
      { x: w * 0.26, y: h }, { x: w * 0.17, y: h * 0.87 },
      { x: w * 0.03, y: h * 0.90 }, { close: true }
    ]
  });
  s.addShape('custGeom', {                                 // right ribbon tail
    x, y, w, h, fill: { color },
    points: [
      { x: w * 0.74, y: h * 0.58 }, { x: w * 0.56, y: h * 0.70 },
      { x: w * 0.74, y: h }, { x: w * 0.83, y: h * 0.87 },
      { x: w * 0.97, y: h * 0.90 }, { close: true }
    ]
  });
  blob(s, x + (w - d) / 2, y, d, color);
  s.addShape('ellipse', {
    x: x + (w - d) / 2 + d * 0.24, y: y + d * 0.24, w: d * 0.52, h: d * 0.52,
    fill: { color: bg }
  });
}

function iconBrain(s, x, y, w, h, color) {
  s.addShape('cloud', { x, y, w, h: h * 0.78, fill: { color } });
  s.addShape('rect', { x: x + w * 0.40, y: y + h * 0.62, w: w * 0.20, h: h * 0.38, fill: { color } });
}

function iconFlame(s, x, y, w, h, color) {
  s.addShape('custGeom', {
    x, y, w, h, fill: { color },
    points: [
      { x: w * 0.52, y: 0 },
      { curve: { type: 'cubic', x1: w * 0.62, y1: h * 0.18, x2: w * 0.86, y2: h * 0.24 }, x: w * 0.94, y: h * 0.40 },
      { curve: { type: 'cubic', x1: w * 1.02, y1: h * 0.56, x2: w * 0.86, y2: h * 0.66 }, x: w * 0.70, y: h * 0.52 },
      { curve: { type: 'cubic', x1: w * 0.66, y1: h * 0.66, x2: w * 0.50, y2: h * 0.62 }, x: w * 0.34, y: h * 0.74 },
      { curve: { type: 'cubic', x1: w * 0.24, y1: h * 0.86, x2: w * 0.34, y2: h * 0.94 }, x: w * 0.44, y: h },
      { curve: { type: 'cubic', x1: w * 0.14, y1: h * 0.94, x2: w * 0.02, y2: h * 0.82 }, x: w * 0.06, y: h * 0.60 },
      { curve: { type: 'cubic', x1: w * 0.10, y1: h * 0.34, x2: w * 0.44, y2: h * 0.28 }, x: w * 0.52, y: 0 },
      { close: true }
    ]
  });
}

function iconThumb(s, x, y, w, h, color) {
  card(s, x, y + h * 0.38, w * 0.24, h * 0.62, color, 22000);
  card(s, x + w * 0.34, y + h * 0.30, w * 0.66, h * 0.70, color, 26000);
  s.addShape('roundRect', {
    x: x + w * 0.36, y: y, w: w * 0.26, h: h * 0.52,
    fill: { color }, rectRadius: w * 0.13, rotate: 14
  });
}

// Outlined origami dart, nose up and to the right, with a centre crease.
function iconPlane(s, x, y, w, h, color) {
  const lw = Math.max(1, Math.round(w * 2.2));
  s.addShape('custGeom', {
    x, y, w, h, fill: { type: 'none' }, line: { color, width: lw },
    points: [
      { x: 0, y: h * 0.48 }, { x: w, y: 0 }, { x: w * 0.56, y: h },
      { x: w * 0.40, y: h * 0.64 }, { close: true }
    ]
  });
  s.addShape('custGeom', {                                 // fold crease
    x, y, w, h, fill: { type: 'none' }, line: { color, width: lw },
    points: [{ x: w * 0.40, y: h * 0.64 }, { x: w, y: 0 }]
  });
}

// Seedling: stem, four alternating leaves and a soil mound.
// Each leaf entry is [x, y, w, h, rotate]; teardrops point up-right at rotate 0.
function iconPlant(s, x, y, w, h, color) {
  s.addShape('rect', { x: x + w * 0.44, y: y + h * 0.08, w: w * 0.10, h: h * 0.82, fill: { color } });
  [[0.52, 0.00, 0.36, 0.24, 0], [0.14, 0.12, 0.34, 0.22, -90],
   [0.52, 0.30, 0.48, 0.28, 45], [0.00, 0.24, 0.48, 0.28, -135]]
    .forEach(function (l) {
      s.addShape('teardrop', {
        x: x + w * l[0], y: y + h * l[1], w: w * l[2], h: h * l[3],
        fill: { color }, rotate: l[4]
      });
    });
  s.addShape('ellipse', { x: x + w * 0.04, y: y + h * 0.84, w: w * 0.92, h: h * 0.20, fill: { color } });
}

// Outlined light bulb with radiating rays.
function iconBulb(s, x, y, w, h, color) {
  const lw = Math.max(1, Math.round(w * 3));
  s.addShape('ellipse', {
    x: x + w * 0.26, y: y + h * 0.14, w: w * 0.48, h: h * 0.50,
    fill: { type: 'none' }, line: { color, width: lw }
  });
  s.addShape('rect', { x: x + w * 0.38, y: y + h * 0.62, w: w * 0.24, h: h * 0.07, fill: { color } });
  s.addShape('rect', { x: x + w * 0.40, y: y + h * 0.74, w: w * 0.20, h: h * 0.07, fill: { color } });
  s.addShape('ellipse', { x: x + w * 0.42, y: y + h * 0.30, w: w * 0.16, h: h * 0.16, fill: { color } });
  for (let a = 0; a < 360; a += 45) {
    const rad = (a * Math.PI) / 180;
    s.addShape('rect', {
      x: x + w * (0.48 + 0.42 * Math.cos(rad)), y: y + h * (0.36 + 0.42 * Math.sin(rad)),
      w: w * 0.16, h: h * 0.045, fill: { color }, rotate: a
    });
  }
}

// Concentric target rings.
function iconTarget(s, x, y, w, h, color) {
  s.addShape('donut', { x, y, w, h, fill: { color } });
  s.addShape('donut', { x: x + w * 0.22, y: y + h * 0.22, w: w * 0.56, h: h * 0.56, fill: { color } });
  s.addShape('ellipse', { x: x + w * 0.40, y: y + h * 0.40, w: w * 0.20, h: h * 0.20, fill: { color } });
}

// Head silhouette in profile with two cogs cut into it.
function iconHead(s, x, y, w, h, color) {
  s.addShape('ellipse', { x, y, w: w * 0.86, h: h * 0.74, fill: { color } });
  s.addShape('rect', { x: x + w * 0.10, y: y + h * 0.50, w: w * 0.56, h: h * 0.50, fill: { color } });
  s.addShape('triangle', { x: x + w * 0.62, y: y + h * 0.32, w: w * 0.30, h: h * 0.34, fill: { color }, rotate: 160 });
  s.addShape('gear6', { x: x + w * 0.34, y: y + h * 0.10, w: w * 0.28, h: h * 0.26, fill: { color: C.purple } });
  s.addShape('gear6', { x: x + w * 0.16, y: y + h * 0.32, w: w * 0.26, h: h * 0.24, fill: { color: C.purple } });
}

/* ---------------------------------------------- repeated page furniture */

function chrome(s, o) {
  o = o || {};
  text(s, 'Yourcompany', { x: 0.876, y: 0.63, w: 1.583, h: 0.269, fontSize: 10, bold: true });
  s.addShape('ellipse', { x: 0.667, y: 0.666, w: 0.209, h: 0.209, fill: { color: C.black } });
  text(s, o.year || '2032', {
    x: o.yearX || 6.149, y: o.yearY || 0.63, w: o.yearW || 0.595, h: 0.269,
    fontSize: 10, bold: true, align: o.yearAlign || 'right'
  });
  text(s, 'Pitch Deck', {
    x: o.deckX || 11.892, y: 0.63, w: o.deckW || 0.881, h: o.deckH || 0.438,
    fontSize: 10, bold: true, align: 'right'
  });
}

// Grey card whose lower part is a coloured "hill" wave (the deck's stat cards).
// Path proportions are taken straight from the original freeforms; `band` gives
// the [top, height] of the coloured area as fractions of the card height.
const BAND_TALL = [0.142, 0.851];   // slides 2 and 4's large cards
const BAND_SHORT = [0.262, 0.731];  // slides 4, 11 and 13's compact cards
function waveCard(s, x, y, w, h, dark, light, band) {
  band = band || BAND_TALL;
  card(s, x, y, w, h, C.grey85, 8033);
  const wy = y + h * band[0], wh = h * band[1];            // coloured band
  const rx = w * 0.074, ry = wh * 0.095;                   // bottom corner radius
  s.addShape('custGeom', {                                 // pale crest (mostly overlapped)
    x, y: wy + wh * 0.027, w, h: wh * 0.554, fill: { color: light },
    points: [
      { x: w, y: wh * 0.554 * 0.437 },
      { curve: { type: 'cubic', x1: w * 0.536, y1: wh * 0.554 * 0.362, x2: w * 0.268, y2: -wh * 0.554 * 0.215 },
        x: w * 0.100, y: wh * 0.554 * 0.086 },
      { curve: { type: 'cubic', x1: w * 0.067, y1: wh * 0.554 * 0.146, x2: w * 0.029, y2: wh * 0.554 * 0.241 },
        x: 0, y: wh * 0.554 * 0.402 },
      { x: 0, y: wh * 0.554 }, { x: w, y: wh * 0.554 }, { close: true }
    ]
  });
  s.addShape('custGeom', {                                 // main coloured hill
    x, y: wy, w, h: wh, fill: { color: dark },
    points: [
      { x: w, y: wh * 0.25 }, { x: w, y: wh - ry },
      { curve: { type: 'cubic', x1: w, y1: wh, x2: w - rx * 0.7, y2: wh }, x: w - rx, y: wh },
      { x: rx, y: wh },
      { curve: { type: 'cubic', x1: rx * 0.7, y1: wh, x2: 0, y2: wh }, x: 0, y: wh - ry },
      { x: 0, y: wh * 0.25 },
      { curve: { type: 'cubic', x1: 0, y1: wh * 0.25, x2: w * 0.194, y2: wh * 0.271 },
        x: w * 0.50, y: wh * 0.063 },
      { curve: { type: 'cubic', x1: w * 0.815, y1: -wh * 0.151, x2: w, y2: wh * 0.25 },
        x: w, y: wh * 0.25 },
      { close: true }
    ]
  });
}

/* --------------------------------------------------------------- slides */

function slide01(s) {
  card(s, 4.275, 1.745, 3.343, 3.083, C.green);
  card(s, 0.667, 1.721, 3.343, 3.083, C.purple);
  blob(s, 5.841, 1.867, 1.652, C.greenLt);
  arrow(s, 2.582, 1.804, 1.40, C.purpleLt, -43);
  photo(s, 7.882, 1.745, 4.784, 5.088, 5128);
  text(s, 'Optimizing Online Sales for Maximum Growth',
    { x: 0.573, y: 5.292, w: 7.405, h: 1.717, fontSize: 48 });
  text(s, '12m+', { x: 0.573, y: 3.586, w: 1.633, h: 0.83, fontSize: 32, color: C.white, align: 'center', lineSpacingMultiple: 1.5 });
  text(s, '86k+', { x: 4.275, y: 3.591, w: 1.497, h: 0.821, fontSize: 32, align: 'center', lineSpacingMultiple: 1.5 });
  text(s, 'View', { x: 0.634, y: 4.352, w: 1.236, h: 0.269, fontSize: 10, color: C.white, align: 'center' });
  text(s, 'Project', { x: 4.346, y: 4.352, w: 1.236, h: 0.269, fontSize: 10, align: 'center' });
  text(s, 'The shoreline meets the vast expanse',
    { x: 2.081, y: 4.009, w: 1.619, h: 0.581, fontSize: 10, color: C.white, lineSpacingMultiple: 1.5 });
  pill(s, 11.096, 6.256, 1.236, 0.409, C.white);
  text(s, 'See More', { x: 11.096, y: 6.326, w: 1.236, h: 0.269, fontSize: 10, align: 'center' });
  chrome(s);
}

function slide02(s) {
  waveCard(s, 4.249, 3.72, 3.343, 3.083, C.green, C.greenLt);
  waveCard(s, 0.667, 3.714, 3.343, 3.083, C.purple, C.purpleMid);
  photo(s, 7.831, 1.97, 4.835, 3.345, 5128);
  card(s, 7.831, 5.589, 4.835, 1.259, C.purple);
  arrow(s, 8.583, 5.828, 0.81, C.purpleLt);
  text(s, 'Opening Remarks', { x: 0.595, y: 1.97, w: 4.372, h: 1.582, fontSize: 44 });
  text(s, 'PLACEHOLDER',
    { x: 4.608, y: 2.522, w: 2.526, h: 0.578, fontSize: 10, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, 'Market', { x: 0.745, y: 5.018, w: 1.146, h: 0.329, fontSize: 10, bold: true, color: C.white, lineSpacingMultiple: 1.5 });
  text(s, 'In the heart of the verdant ', { x: 0.753, y: 5.246, w: 1.516, h: 0.582, fontSize: 10, color: C.white, lineSpacingMultiple: 1.5 });
  text(s, '531k', { x: 2.919, y: 5.802, w: 1.871, h: 0.825, fontSize: 32, color: C.white, lineSpacingMultiple: 1.5 });
  text(s, 'Sales', { x: 4.309, y: 4.984, w: 1.146, h: 0.329, fontSize: 10, bold: true, lineSpacingMultiple: 1.5 });
  text(s, 'In the heart of the verdant ', { x: 4.317, y: 5.213, w: 1.516, h: 0.582, fontSize: 10, lineSpacingMultiple: 1.5 });
  text(s, '+83%', { x: 6.185, y: 5.834, w: 1.871, h: 0.825, fontSize: 32, lineSpacingMultiple: 1.5 });
  text(s, '+1.421.21', { x: 11.505, y: 5.892, w: 1.133, h: 0.329, fontSize: 10, bold: true, color: C.white, lineSpacingMultiple: 1.5 });
  text(s, '$2.641.51', { x: 10.529, y: 5.811, w: 2.924, h: 0.83, fontSize: 32, color: C.white, lineSpacingMultiple: 1.5 });
  chrome(s, { year: '2031', yearX: 5.921, yearY: 0.664, yearW: 1.183, yearAlign: 'center', deckX: 11.59, deckW: 1.183, deckH: 0.269 });
}

function slide03(s) {
  card(s, 5.754, 1.938, 6.913, 2.307, C.purple, 5203);
  card(s, 5.754, 4.469, 3.343, 2.392, C.purple);
  card(s, 9.324, 4.469, 3.343, 2.4, C.green);
  photo(s, 0.667, 4.563, 4.843, 2.307, 5730);
  arrow(s, 8.749, 2.606, 0.92, C.purpleLt);
  arrow(s, 8.032, 4.563, 0.92, C.purpleLt, -43);
  blob(s, 11.457, 4.584, 1.083, C.greenLt);
  text(s, 'Problem\nStatement', { x: 0.561, y: 2.211, w: 4.572, h: 1.582, fontSize: 44 });
  text(s, '$3.145.21', { x: 6.149, y: 2.69, w: 2.746, h: 0.83, fontSize: 32, color: C.white, lineSpacingMultiple: 1.5 });
  text(s, '-2.532', { x: 7.298, y: 2.583, w: 1.133, h: 0.329, fontSize: 10, bold: true, color: C.white, lineSpacingMultiple: 1.5 });
  text(s, '$3.131.52', { x: 10.234, y: 2.606, w: 2.746, h: 0.83, fontSize: 32, color: C.white, lineSpacingMultiple: 1.5 });
  text(s, '-4.232', { x: 11.382, y: 2.499, w: 1.133, h: 0.329, fontSize: 10, bold: true, color: C.white, lineSpacingMultiple: 1.5 });
  tag(s, 5.95, 4.834, 0.882, 0.243, C.white);
  text(s, 'Project', { x: 5.972, y: 4.821, w: 0.868, h: 0.269, fontSize: 10, color: C.white, align: 'center' });
  text(s, '72+', { x: 5.89, y: 5.128, w: 1.031, h: 0.83, fontSize: 32, color: C.white, lineSpacingMultiple: 1.5 });
  text(s, 'High above, the towering mountains pierce the sky, their snow-capped peaks touching the clouds. ',
    { x: 5.911, y: 5.799, w: 2.458, h: 0.834, fontSize: 10, color: C.white, align: 'justify', lineSpacingMultiple: 1.5 });
  tag(s, 9.562, 4.855, 0.882, 0.243, C.black);
  text(s, 'Income', { x: 9.621, y: 4.856, w: 0.868, h: 0.269, fontSize: 10, align: 'justify' });
  text(s, '82%', { x: 9.589, y: 5.097, w: 1.031, h: 0.83, fontSize: 32, lineSpacingMultiple: 1.5 });
  text(s, 'High above, the towering mountains pierce the sky, their snow-capped peaks touching the clouds. ',
    { x: 9.562, y: 5.803, w: 2.534, h: 0.834, fontSize: 10, align: 'justify', lineSpacingMultiple: 1.5 });
  chrome(s);
}

function slide04(s) {
  card(s, 5.762, 1.823, 6.883, 2.307, C.purple, 5203);
  waveCard(s, 9.296, 4.303, 3.35, 2.53, C.green, C.greenLt, BAND_SHORT);
  waveCard(s, 5.763, 4.303, 3.343, 2.53, C.purple, C.purpleMid, BAND_SHORT);
  photo(s, 9.533, 1.954, 2.941, 2.004, 4950);
  text(s, 'Market Opportunity', { x: 0.583, y: 2.148, w: 4.572, h: 1.582, fontSize: 44 });
  text(s, '70%', { x: 7.084, y: 1.957, w: 2.006, h: 0.64, fontSize: 32, color: C.white, align: 'center' });
  text(s, 'Market Trends', { x: 5.915, y: 2.537, w: 1.276, h: 0.352, fontSize: 11, color: C.white, align: 'center', lineSpacingMultiple: 1.5 });
  text(s, 'High above, the towering mountains pierce the sky, their snow-capped peaks touching the clouds.',
    { x: 5.963, y: 2.867, w: 2.455, h: 0.834, fontSize: 10, color: C.white, align: 'justify', lineSpacingMultiple: 1.5 });
  tag(s, 0.677, 4.503, 1.209, 0.243, C.black);
  text(s, 'Marketing', { x: 0.677, y: 4.494, w: 1.237, h: 0.269, fontSize: 10, align: 'center' });
  text(s, '96%', { x: 3.571, y: 4.004, w: 1.072, h: 0.83, fontSize: 32, align: 'right', lineSpacingMultiple: 1.5 });
  text(s, 'High above, the towering mountains pierce the sky', { x: 0.594, y: 4.735, w: 2.044, h: 0.581, fontSize: 10, lineSpacingMultiple: 1.5 });
  tag(s, 0.679, 5.951, 1.209, 0.243, C.black);
  text(s, 'Client', { x: 0.619, y: 5.925, w: 1.368, h: 0.269, fontSize: 10, align: 'center' });
  text(s, '85+', { x: 3.571, y: 5.363, w: 1.072, h: 0.83, fontSize: 32, align: 'right', lineSpacingMultiple: 1.5 });
  text(s, 'High above, the towering mountains pierce the sky', { x: 0.602, y: 6.171, w: 2.069, h: 0.581, fontSize: 10, lineSpacingMultiple: 1.5 });
  text(s, '82%', { x: 6.426, y: 5.528, w: 2.006, h: 0.841, fontSize: 44, color: C.white, align: 'center' });
  text(s, 'Market Potentials', { x: 6.313, y: 6.175, w: 2.268, h: 0.352, fontSize: 11, color: C.white, align: 'center', lineSpacingMultiple: 1.5 });
  text(s, '86%', { x: 9.978, y: 5.488, w: 2.006, h: 0.841, fontSize: 44, align: 'center' });
  text(s, 'Target Audience', { x: 10.11, y: 6.11, w: 1.735, h: 0.346, fontSize: 11, align: 'center', lineSpacingMultiple: 1.5 });
  chrome(s);
}

function slide05(s) {
  card(s, 6.667, 2.663, 6.009, 1.719, C.green, 10860);
  card(s, 6.667, 4.717, 2.659, 2.094, C.purple);
  photo(s, 0.667, 4.74, 5.661, 2.094, 5323);
  blob(s, 7.063, 2.91, 1.275, C.greenLt);
  blob(s, 8.481, 2.91, 1.275, C.greenLt);
  text(s, 'Solution \nOverview', { x: 0.547, y: 2.379, w: 3.825, h: 1.582, fontSize: 44 });
  text(s, 'Social Media Marketing', { x: 6.667, y: 1.681, w: 2.43, h: 0.774, fontSize: 20, bold: true });
  text(s, 'Feedback', { x: 10.189, y: 2.918, w: 2.12, h: 0.329, fontSize: 10, bold: true, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, 'High above, the towering mountains pierce the sky', { x: 10.156, y: 3.222, w: 2.12, h: 0.578, fontSize: 10, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, 'See More', { x: 10.156, y: 3.743, w: 2.314, h: 0.33, fontSize: 10, bold: true, underline: { style: 'sng' }, align: 'justify', lineSpacingMultiple: 1.5 });
  tag(s, 6.771, 4.887, 1.209, 0.243, C.white);
  text(s, 'Reporting', { x: 6.771, y: 4.874, w: 1.237, h: 0.269, fontSize: 10, color: C.white, align: 'center' });
  text(s, 'High above, the towering mountains pierce the sky, their snow-capped peaks touching the clouds. ',
    { x: 6.739, y: 5.28, w: 2.404, h: 0.834, fontSize: 10, color: C.white, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, '85+', { x: 8.395, y: 5.981, w: 1.209, h: 0.83, fontSize: 32, color: C.white, lineSpacingMultiple: 1.5 });
  tag(s, 10.215, 4.942, 1.209, 0.243, C.black);
  text(s, 'Insights', { x: 10.156, y: 4.916, w: 1.368, h: 0.269, fontSize: 10, align: 'center' });
  text(s, 'High above, the towering mountains pierce the sky, their snow-capped peaks touching the clouds. ',
    { x: 10.215, y: 5.302, w: 2.452, h: 0.834, fontSize: 10, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, '96%', { x: 11.797, y: 6.106, w: 1.209, h: 0.83, fontSize: 32, lineSpacingMultiple: 1.5 });
  chrome(s);
}

function slide06(s) {
  card(s, 6.609, 4.381, 6.075, 2.468, C.purple, 7905);
  card(s, 0.667, 4.381, 5.635, 2.468, C.green, 8117);
  photo(s, 6.678, 1.671, 5.988, 2.38, 5323);
  blob(s, 4.801, 5.567, 1.095, C.greenLt);
  arrow(s, 11.305, 5.578, 1.08, C.purpleLt, -45);
  text(s, 'Unique Value Proposition', { x: 0.593, y: 1.921, w: 4.755, h: 1.582, fontSize: 44 });
  text(s, 'Overview', { x: 0.663, y: 3.908, w: 3.593, h: 0.303, fontSize: 12, bold: true });
  text(s, 'Value ', { x: 1.125, y: 4.939, w: 1.07, h: 0.269, fontSize: 10, bold: true });
  meter(s, 1.963, 5.003, 2.911, 0.134, 2.156, C.grey95, C.black);
  text(s, 'The shoreline meets the vast expanse of the ocean with an eternal ebb and flow. ',
    { x: 1.125, y: 5.388, w: 2.911, h: 0.578, fontSize: 10, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, 'Project', { x: 5.093, y: 4.746, w: 0.796, h: 0.329, fontSize: 10, lineSpacingMultiple: 1.5 });
  text(s, '11.2+', { x: 5.067, y: 4.882, w: 1.026, h: 0.511, fontSize: 18, lineSpacingMultiple: 1.5 });
  text(s, 'Income', { x: 7.153, y: 4.888, w: 1.07, h: 0.329, fontSize: 10, bold: true, color: C.white, align: 'justify', lineSpacingMultiple: 1.5 });
  meter(s, 8.443, 5.023, 2.911, 0.134, 2.156, C.black, C.white);
  text(s, 'The shoreline meets the vast expanse of the ocean with an eternal ebb and flow. ',
    { x: 7.2, y: 5.422, w: 2.911, h: 0.578, fontSize: 10, color: C.white, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, 'Market', { x: 11.574, y: 4.784, w: 0.796, h: 0.329, fontSize: 10, color: C.white, lineSpacingMultiple: 1.5 });
  text(s, '+23%', { x: 11.504, y: 4.948, w: 0.963, h: 0.511, fontSize: 18, color: C.white, lineSpacingMultiple: 1.5 });
  chrome(s);
}

function slide07(s) {
  // Four "step" cards with an icon, a heading and a paragraph.
  const steps = [
    { x: 0.667, fill: C.green, label: 'Step One', fg: C.black, tx: 0.916, bx: 0.916, by: 3.387 },
    { x: 3.773, fill: C.purple, label: 'Step Two', fg: C.white, tx: 3.985, bx: 3.98, by: 3.401 },
    { x: 6.785, fill: C.green, label: 'Step Three', fg: C.black, tx: 7.044, bx: 7.044, by: 3.422 },
    { x: 9.902, fill: C.purple, label: 'Step Four', fg: C.white, tx: 10.108, bx: 10.108, by: 3.38 }
  ];
  steps.forEach(function (st) { card(s, st.x, 1.933, 2.765, 2.977, st.fill, 8117); });
  iconAward(s, 2.26, 2.112, 0.784, 1.133, C.greenLt, C.green);
  iconBrain(s, 5.233, 2.21, 1.056, 0.9, C.purpleLt);
  iconFlame(s, 8.55, 2.106, 0.777, 1.136, C.greenLt);
  iconThumb(s, 11.426, 2.204, 1.056, 0.924, C.purpleLt);
  steps.forEach(function (st) {
    text(s, st.label, { x: st.tx, y: 3.038, w: 2.22, h: 0.286, fontSize: 11, bold: true, color: st.fg });
    text(s, 'The shoreline meets the vast expanse of the ocean with an eternal ebb and flow. ',
      { x: st.bx, y: st.by, w: 1.977, h: 0.834, fontSize: 10, color: st.fg, align: 'justify', lineSpacingMultiple: 1.5 });
  });
  text(s, 'How It Works', { x: 4.044, y: 5.591, w: 4.805, h: 0.841, fontSize: 44, align: 'center' });
  chrome(s);
}

function slide08(s) {
  card(s, 3.664, 2.228, 3.167, 1.989, C.purple, 8117);
  arrow(s, 5.97, 2.395, 0.71, C.purpleLt, -45);
  text(s, 'Features and Benefits', { x: 8.646, y: 2.365, w: 4.704, h: 1.582, fontSize: 44 });
  text(s, 'Feature 2', { x: 0.632, y: 2.902, w: 1.965, h: 0.438, fontSize: 20 });
  text(s, 'PLACEHOLDER',
    { x: 0.632, y: 3.382, w: 1.904, h: 0.834, fontSize: 10, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, 'Benefits 2', { x: 3.984, y: 2.907, w: 1.996, h: 0.438, fontSize: 20, color: C.white });
  text(s, 'The crisp air carries the promise of adventure',
    { x: 3.984, y: 3.339, w: 1.729, h: 0.581, fontSize: 10, color: C.white, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, 'Feature 1', { x: 0.632, y: 4.823, w: 2.213, h: 0.438, fontSize: 20 });
  text(s, 'PLACEHOLDER',
    { x: 0.632, y: 5.323, w: 2.308, h: 0.834, fontSize: 10, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, 'Benefits 1', { x: 3.848, y: 4.791, w: 2.213, h: 0.438, fontSize: 20 });
  text(s, 'PLACEHOLDER',
    { x: 3.848, y: 5.346, w: 2.324, h: 0.834, fontSize: 10, align: 'justify', lineSpacingMultiple: 1.5 });
  // Two labelled progress rows on the right.
  text(s, 'Items', { x: 8.808, y: 4.324, w: 0.868, h: 0.269, fontSize: 10, bold: true, align: 'justify' });
  meter(s, 9.698, 4.379, 1.575, 0.112, 1.078, C.black, C.green);
  text(s, '72+', { x: 11.301, y: 4.216, w: 0.737, h: 0.326, fontSize: 10, lineSpacingMultiple: 1.5 });
  text(s, 'The crisp air carries the promise', { x: 8.821, y: 4.542, w: 2.324, h: 0.329, fontSize: 10, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, 'Quality ', { x: 8.808, y: 5.317, w: 0.868, h: 0.269, fontSize: 10, bold: true, align: 'justify' });
  meter(s, 9.671, 5.407, 1.575, 0.112, 1.439, C.black, C.purple);
  text(s, '93%', { x: 11.315, y: 5.217, w: 0.737, h: 0.326, fontSize: 10, lineSpacingMultiple: 1.5 });
  text(s, 'The crisp air carries the promise', { x: 8.821, y: 5.608, w: 2.324, h: 0.329, fontSize: 10, align: 'justify', lineSpacingMultiple: 1.5 });
  chrome(s);
}

function slide09(s) {
  card(s, 3.794, 4.324, 4.116, 2.512, C.purple, 4861);
  card(s, 8.205, 4.324, 4.461, 2.512, C.green, 4861);
  photo(s, 0.678, 4.324, 2.815, 2.51, 5323);
  iconPlane(s, 6.288, 4.535, 1.217, 1.125, C.purpleLt);
  iconPlant(s, 11.334, 4.58, 0.86, 1.176, C.greenLt);
  text(s, 'Business Model', { x: 0.557, y: 2.168, w: 3.678, h: 1.582, fontSize: 44 });
  tag(s, 3.914, 2.712, 0.882, 0.243, C.black);
  text(s, 'Project', { x: 4.018, y: 2.695, w: 0.868, h: 0.269, fontSize: 10, align: 'justify' });
  text(s, '75+', { x: 5.218, y: 2.209, w: 1.112, h: 0.921, fontSize: 36, lineSpacingMultiple: 1.5 });
  meter(s, 3.914, 3.119, 2.065, 0.112, 1.488, C.grey85, C.green);
  text(s, '+23%', { x: 3.794, y: 3.098, w: 1.871, h: 0.636, fontSize: 24, lineSpacingMultiple: 1.5 });
  text(s, 'High above', { x: 5.341, y: 3.293, w: 1.402, h: 0.329, fontSize: 10, lineSpacingMultiple: 1.5 });
  tag(s, 8.277, 2.761, 0.882, 0.243, C.black);
  text(s, 'Income', { x: 8.297, y: 2.744, w: 0.868, h: 0.269, fontSize: 10, align: 'center' });
  text(s, '86%', { x: 9.595, y: 2.281, w: 1.112, h: 0.921, fontSize: 36, lineSpacingMultiple: 1.5 });
  meter(s, 8.277, 3.168, 2.224, 0.112, 1.83, C.grey85, C.purple);
  text(s, '11.2+', { x: 8.266, y: 3.13, w: 1.582, h: 0.636, fontSize: 24, lineSpacingMultiple: 1.5 });
  text(s, 'High above', { x: 9.677, y: 3.318, w: 1.402, h: 0.329, fontSize: 10, lineSpacingMultiple: 1.5 });
  text(s, 'Pricing', { x: 3.95, y: 5.507, w: 1.521, h: 0.438, fontSize: 20, color: C.white });
  text(s, 'In the heart of the verdant wilderness, where the symphony of leaves rustling',
    { x: 3.959, y: 6.004, w: 2.724, h: 0.581, fontSize: 10, color: C.white, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, 'Revenue', { x: 8.429, y: 5.507, w: 1.602, h: 0.438, fontSize: 20 });
  text(s, 'In the heart of the verdant wilderness, where the symphony of leaves rustling',
    { x: 8.429, y: 6.004, w: 2.729, h: 0.581, fontSize: 10, align: 'justify', lineSpacingMultiple: 1.5 });
  chrome(s);
}

function slide10(s) {
  card(s, 0.667, 1.781, 2.849, 2.512, C.purple, 4861);
  card(s, 6.115, 4.37, 2.849, 2.512, C.green, 4861);
  photo(s, 7.74, 1.744, 4.926, 2.362, 5323);
  iconAward(s, 2.801, 2.121, 0.383, 0.554, C.purpleLt, C.purple);
  iconBrain(s, 6.366, 2.192, 0.516, 0.44, C.black);
  iconFlame(s, 8.305, 4.895, 0.38, 0.555, C.greenLt);
  iconThumb(s, 12.109, 4.999, 0.516, 0.452, C.black);
  text(s, 'Customized Services For Your Unique Needs', { x: 0.513, y: 4.648, w: 5.636, h: 2.322, fontSize: 44 });
  const svc = [
    { n: 'Services 1 ', x: 0.822, ty: 2.513, by: 2.925, bw: 2.22, fg: C.white },
    { n: 'Services 3', x: 4.091, ty: 2.527, by: 2.94, bw: 2.245, fg: C.black },
    { n: 'Services 2', x: 6.366, ty: 5.314, by: 5.714, bw: 2.244, fg: C.black },
    { n: 'Services 4', x: 9.885, ty: 5.314, by: 5.714, bw: 1.985, fg: C.black }
  ];
  svc.forEach(function (v) {
    text(s, v.n, { x: v.x, y: v.ty, w: 2.129, h: 0.438, fontSize: 20, color: v.fg });
    text(s, 'PLACEHOLDER',
      { x: v.x, y: v.by, w: v.bw, h: 0.907, fontSize: 11, color: v.fg, align: 'justify', lineSpacingMultiple: 1.5 });
  });
  chrome(s);
}

function slide11(s) {
  // Four wave cards; each holds a big number and a caption.
  const cards = [
    { x: 0.667, y: 4.513, w: 2.616, h: 2.321, dark: C.green, light: C.greenLt,
      num: '573k+', nx: 1.182, ny: 5.509, nw: 1.685, cap: 'View', cx: 1.511, cy: 6.257, fg: C.black },
    { x: 3.855, y: 4.513, w: 2.605, h: 2.323, dark: C.purple, light: C.purpleMid,
      num: '72%', nx: 4.384, ny: 5.509, nw: 1.55, cap: 'Income', cx: 4.739, cy: 6.257, fg: C.white },
    { x: 6.902, y: 1.789, w: 2.616, h: 2.321, dark: C.green, light: C.greenLt,
      num: '22-', nx: 7.524, ny: 2.775, nw: 1.553, cap: 'Expanse', cx: 7.707, cy: 3.524, fg: C.black },
    { x: 10.006, y: 1.789, w: 2.605, h: 2.323, dark: C.purple, light: C.purpleMid,
      num: '141+', nx: 10.619, ny: 2.812, nw: 1.447, cap: 'Project', cx: 10.619, cy: 3.56, fg: C.white }
  ];
  cards.forEach(function (v) { waveCard(s, v.x, v.y, v.w, v.h, v.dark, v.light, BAND_SHORT); });
  photo(s, 6.873, 4.506, 5.794, 2.327, 5323);
  cards.forEach(function (v) {
    text(s, v.num, { x: v.nx, y: v.ny, w: v.nw, h: 0.83, fontSize: 32, bold: true, color: v.fg, align: 'center', lineSpacingMultiple: 1.5 });
    text(s, v.cap, { x: v.cx, y: v.cy, w: v.cap === 'Project' ? 1.311 : (v.cap === 'Expanse' ? 1.056 : 0.892), h: 0.353,
      fontSize: 11, bold: true, color: v.fg, align: 'center', lineSpacingMultiple: 1.5 });
  });
  text(s, 'Future Plans', { x: 0.531, y: 2.146, w: 5.691, h: 0.841, fontSize: 44 });
  text(s, 'PLACEHOLDER',
    { x: 0.577, y: 2.861, w: 2.863, h: 0.629, fontSize: 11, align: 'justify', lineSpacingMultiple: 1.5 });
  chrome(s);
}

function slide12(s) {
  const strat = [
    { x: 6.32, y: 1.754, fill: C.green, t: 'Strategy one', tx: 6.528, ty: 2.298, bx: 6.528, by: 2.634, bw: 1.871, fg: C.black, icon: 'blob', ix: 8.397, iy: 3.194 },
    { x: 9.657, y: 1.756, fill: C.purple, t: 'Strategy Three', tx: 9.924, ty: 2.279, bx: 9.924, by: 2.643, bw: 1.809, fg: C.white, icon: 'arrow', ix: 11.818, iy: 3.251 },
    { x: 6.284, y: 4.476, fill: C.purple, t: 'Strategy Two', tx: 6.563, ty: 4.752, bx: 6.563, by: 5.145, bw: 1.871, fg: C.white, icon: 'arrow', ix: 8.414, iy: 5.958 },
    { x: 9.657, y: 4.444, fill: C.green, t: 'Strategy Four', tx: 10.001, ty: 4.734, bx: 10.001, by: 5.164, bw: 1.865, fg: C.black, icon: 'blob', ix: 11.795, iy: 5.901 }
  ];
  strat.forEach(function (v) { card(s, v.x, v.y, 3.009, 2.389, v.fill, 8835); });
  strat.forEach(function (v) {
    if (v.icon === 'blob') blob(s, v.ix, v.iy, 0.791, C.greenLt);
    else arrow(s, v.ix, v.iy, 0.75, C.purpleLt, -45);
  });
  photo(s, 0.667, 4.476, 4.5, 2.357, 7753);
  text(s, 'Marketing \nStrategy', { x: 0.581, y: 2.179, w: 5.151, h: 1.447, fontSize: 40 });
  strat.forEach(function (v) {
    text(s, v.t, { x: v.tx, y: v.ty, w: 2.882, h: 0.438, fontSize: 20, bold: true, color: v.fg });
    text(s, 'PLACEHOLDER',
      { x: v.bx, y: v.by, w: v.bw, h: 0.834, fontSize: 10, color: v.fg, align: 'justify', lineSpacingMultiple: 1.5 });
  });
  chrome(s);
}

function slide13(s) {
  waveCard(s, 0.684, 1.808, 2.605, 2.323, C.purple, C.purpleMid, BAND_SHORT);
  waveCard(s, 0.662, 4.513, 2.616, 2.321, C.green, C.greenLt, BAND_SHORT);
  waveCard(s, 7.218, 4.513, 2.605, 2.323, C.purple, C.purpleMid, BAND_SHORT);
  text(s, 'Financial Projections', { x: 7.13, y: 2.198, w: 4.395, h: 1.582, fontSize: 44 });
  text(s, '$1.201.802', { x: 0.9, y: 3.238, w: 2.136, h: 0.572, fontSize: 28, align: 'center' });
  tag(s, 3.552, 2.428, 1.022, 0.282, C.black);
  text(s, 'Profit', { x: 3.398, y: 2.441, w: 1.33, h: 0.269, fontSize: 10, align: 'center' });
  text(s, 'In the heart of the verdant wilderness, where the symphony',
    { x: 3.458, y: 2.796, w: 2.233, h: 0.582, fontSize: 10, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, '$892.234', { x: 1.135, y: 5.977, w: 1.704, h: 0.572, fontSize: 28, align: 'center' });
  tag(s, 3.575, 5.06, 1.022, 0.282, C.black);
  text(s, 'Revenue', { x: 3.458, y: 5.076, w: 1.33, h: 0.269, fontSize: 10, align: 'center' });
  text(s, 'In the heart of the verdant wilderness, where the symphony',
    { x: 3.508, y: 5.544, w: 2.233, h: 0.582, fontSize: 10, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, '$734.195', { x: 7.789, y: 5.856, w: 1.704, h: 0.572, fontSize: 28 });
  tag(s, 10.044, 5.031, 1.022, 0.282, C.black);
  text(s, 'Expenses', { x: 9.933, y: 5.044, w: 1.33, h: 0.269, fontSize: 10, align: 'center' });
  text(s, 'High above, the towering mountains pierce the sky, their snow-capped peaks touching the clouds. ',
    { x: 9.961, y: 5.56, w: 2.605, h: 0.834, fontSize: 10, lineSpacingMultiple: 1.5 });
  chrome(s);
}

function slide14(s) {
  const funds = [
    { x: 0.653, fill: C.green, label: 'Development', fg: C.black, lx: 1.4, bx: 1.403, icon: iconBulb, ix: 0.822, iy: 4.417, iw: 0.507, ih: 0.549, ic: C.black },
    { x: 3.687, fill: C.purple, label: 'Marketing', fg: C.white, lx: 4.508, bx: 4.508, icon: iconTarget, ix: 3.982, iy: 4.484, iw: 0.436, ih: 0.436, ic: C.white },
    { x: 6.745, fill: C.green, label: 'Operations', fg: C.black, lx: 7.543, bx: 7.51, icon: iconPlant, ix: 7.052, iy: 4.45, iw: 0.323, ih: 0.441, ic: C.black },
    { x: 9.798, fill: C.purple, label: 'Facility', fg: C.white, lx: 10.632, bx: 10.632, icon: iconHead, ix: 10.167, iy: 4.417, iw: 0.375, ih: 0.445, ic: C.white }
  ];
  funds.forEach(function (v) { card(s, v.x, 3.589, 2.869, 2.734, v.fill, 8835); });
  text(s, 'Use Of Funds', { x: 4.067, y: 1.457, w: 5.011, h: 0.841, fontSize: 44 });
  funds.forEach(function (v) {
    v.icon(s, v.ix, v.iy, v.iw, v.ih, v.ic);
    text(s, v.label, { x: v.lx, y: 4.495, w: 2.344, h: 0.404, fontSize: 18, color: v.fg });
    text(s, 'PLACEHOLDER',
      { x: v.bx, y: 4.898, w: 1.95, h: 0.834, fontSize: 10, color: v.fg, lineSpacingMultiple: 1.5 });
  });
  chrome(s);
}

function slide15(s) {
  card(s, 8.953, 2.571, 3.714, 1.901, C.green, 8835);
  card(s, 5.136, 4.86, 2.869, 1.963, C.purple, 8835);
  photo(s, 0.667, 2.581, 3.965, 4.243, 5231);
  blob(s, 11.778, 2.826, 0.791, C.greenLt);
  arrow(s, 7.192, 6.123, 0.65, C.purpleLt, -45);
  text(s, 'Investment Opportunity', { x: 4.942, y: 2.709, w: 4.39, h: 1.582, fontSize: 44 });
  text(s, 'Acquisition', { x: 9.332, y: 3.017, w: 1.784, h: 0.404, fontSize: 18 });
  text(s, 'High above, the towering mountains pierce the sky.',
    { x: 9.344, y: 3.461, w: 1.784, h: 0.581, fontSize: 10, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, 'IPO', { x: 5.386, y: 5.104, w: 1.784, h: 0.404, fontSize: 18, color: C.white });
  text(s, 'High above, the towering mountains pierce the sky.',
    { x: 5.365, y: 5.522, w: 1.827, h: 0.578, fontSize: 10, color: C.white, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, 'Return On Investment', { x: 8.268, y: 5.226, w: 2.528, h: 0.707, fontSize: 18 });
  text(s, 'PLACEHOLDER',
    { x: 8.296, y: 5.868, w: 2.92, h: 0.581, fontSize: 10, lineSpacingMultiple: 1.5 });
  // Off-canvas call-to-action block, parked above the slide in the original.
  text(s, 'In the heart of the verdant wilderness, where the symphony of leaves rustling and birdsong intertwine.',
    { x: 4.5, y: -2.841, w: 4.77, h: 0.581, fontSize: 10, lineSpacingMultiple: 1.5 });
  pill(s, 4.578, -2.012, 1.311, 0.377, C.black);
  text(s, 'Get Started', { x: 4.632, y: -1.956, w: 1.236, h: 0.269, fontSize: 10, color: C.white, align: 'center' });
  text(s, 'Cancel', { x: 6.123, y: -1.956, w: 1.236, h: 0.269, fontSize: 10, align: 'center' });
  chrome(s);
}

function slide16(s) {
  card(s, 6.667, 2.504, 2.869, 2.734, C.purple, 8835);
  card(s, 9.798, 2.519, 2.869, 2.734, C.purple, 8835);
  card(s, 6.667, 5.495, 6.0, 1.308, C.green, 8835);
  photo(s, 0.667, 4.715, 5.013, 2.108, 5231);
  blob(s, 7.074, 5.71, 0.942, C.greenLt);
  text(s, 'Market Challenges and Risks', { x: 0.575, y: 2.226, w: 4.502, h: 2.322, fontSize: 44 });
  text(s, 'Risks ', { x: 6.772, y: 3.173, w: 1.07, h: 0.269, fontSize: 10, bold: true, color: C.white });
  text(s, '2026 - 2027', { x: 6.772, y: 3.539, w: 1.308, h: 0.269, fontSize: 10, color: C.white });
  text(s, 'Client', { x: 6.776, y: 3.89, w: 0.892, h: 0.353, fontSize: 11, color: C.white, lineSpacingMultiple: 1.5 });
  text(s, '573k+', { x: 8.575, y: 3.951, w: 0.792, h: 0.326, fontSize: 10, color: C.white, align: 'right', lineSpacingMultiple: 1.5 });
  meter(s, 6.879, 4.328, 2.385, 0.112, 1.094, C.black, C.green);
  text(s, 'The crisp air carries the promise ', { x: 6.776, y: 4.574, w: 2.703, h: 0.33, fontSize: 10, color: C.white, lineSpacingMultiple: 1.5 });
  text(s, '2026 - 2027', { x: 9.971, y: 3.617, w: 1.308, h: 0.269, fontSize: 10, color: C.white });
  text(s, 'Project', { x: 9.987, y: 3.89, w: 1.311, h: 0.353, fontSize: 11, color: C.white, lineSpacingMultiple: 1.5 });
  text(s, '141+', { x: 11.783, y: 4.034, w: 0.679, h: 0.326, fontSize: 10, color: C.white, align: 'right', lineSpacingMultiple: 1.5 });
  meter(s, 10.079, 4.384, 2.286, 0.112, 1.094, C.black, C.green);
  text(s, 'The crisp air carries the promise ', { x: 9.987, y: 4.574, w: 3.289, h: 0.33, fontSize: 10, color: C.white, lineSpacingMultiple: 1.5 });
  text(s, 'High above, the towering mountains pierce the sky, their snow-capped peaks touching the clouds. ',
    { x: 8.284, y: 5.838, w: 3.727, h: 0.582, fontSize: 10, lineSpacingMultiple: 1.5 });
  chrome(s);
}

function slide17(s) {
  card(s, 3.503, 4.584, 2.869, 2.249, C.purple, 8835);
  card(s, 7.464, 3.081, 2.816, 2.02, C.green, 8835);
  arrow(s, 5.536, 4.807, 0.65, C.purpleLt, -45);
  blob(s, 9.364, 3.226, 0.791, C.greenLt);
  text(s, 'Principle Of Success', { x: 0.612, y: 2.518, w: 4.085, h: 1.582, fontSize: 44 });
  const steps = [
    { t: 'Step One', tx: 0.648, ty: 5.693, bx: 0.648, by: 5.999, bw: 2.297, bh: 0.581, fg: C.black },
    { t: 'Step Two', tx: 3.592, ty: 5.693, bx: 3.592, by: 5.96, bw: 2.297, bh: 0.581, fg: C.white },
    { t: 'Step Three', tx: 7.587, ty: 3.748, bx: 7.587, by: 4.054, bw: 2.028, bh: 0.831, fg: C.black },
    { t: 'Step Four', tx: 10.693, ty: 3.592, bx: 10.693, by: 3.898, bw: 2.028, bh: 0.831, fg: C.black }
  ];
  steps.forEach(function (v) {
    text(s, v.t, { x: v.tx, y: v.ty, w: 2.033, h: 0.269, fontSize: 10, bold: true, color: v.fg });
    text(s, 'PLACEHOLDER',
      { x: v.bx, y: v.by, w: v.bw, h: v.bh, fontSize: 10, color: v.fg, align: 'justify', lineSpacingMultiple: 1.5 });
  });
  text(s, 'Success', { x: 7.477, y: 5.652, w: 1.311, h: 0.353, fontSize: 11, lineSpacingMultiple: 1.5 });
  text(s, '98%', { x: 10.44, y: 5.679, w: 0.679, h: 0.326, fontSize: 10, align: 'right', lineSpacingMultiple: 1.5 });
  meter(s, 7.583, 6.07, 3.516, 0.112, 3.333, C.black, C.green);
  text(s, 'High above, the towering mountains pierce the sky.',
    { x: 7.464, y: 6.212, w: 4.16, h: 0.33, fontSize: 10, align: 'justify', lineSpacingMultiple: 1.5 });
  chrome(s);
}

function slide18(s) {
  card(s, 8.486, 2.613, 4.214, 1.968, C.green, 6835);
  card(s, 8.486, 4.865, 4.167, 1.968, C.purple, 6835);
  blob(s, 8.698, 3.022, 1.134, C.greenLt);
  photo(s, 0.68, 5.544, 1.282, 1.273, 7674);
  photo(s, 6.282, 3.78, 1.282, 1.273, 7674);
  text(s, 'Our Clients Say', { x: 0.606, y: 2.55, w: 3.342, h: 1.582, fontSize: 44 });
  text(s, '\u201C', { x: 5.555, y: 3.547, w: 0.143, h: 0.774, fontSize: 40, align: 'right' });
  text(s, 'Fanny Price', { x: 3.735, y: 4.036, w: 2.271, h: 0.404, fontSize: 18, align: 'right' });
  text(s, 'In the heart of the verdant wilderness, where the symphony.',
    { x: 3.186, y: 4.441, w: 2.819, h: 0.581, fontSize: 10, align: 'right', lineSpacingMultiple: 1.5 });
  text(s, '\u201C', { x: 2.205, y: 5.391, w: 0.143, h: 0.774, fontSize: 40, align: 'right' });
  text(s, 'John Smith', { x: 2.205, y: 5.802, w: 2.819, h: 0.404, fontSize: 18 });
  text(s, 'In the heart of the verdant wilderness, where the symphony.',
    { x: 2.205, y: 6.207, w: 3.294, h: 0.581, fontSize: 10, lineSpacingMultiple: 1.5 });
  text(s, 'PLACEHOLDER',
    { x: 9.997, y: 3.306, w: 2.493, h: 0.581, fontSize: 10, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, 'Client', { x: 8.698, y: 5.076, w: 0.892, h: 0.353, fontSize: 11, color: C.white, lineSpacingMultiple: 1.5 });
  text(s, '573k+', { x: 11.679, y: 5.085, w: 0.792, h: 0.326, fontSize: 10, color: C.white, align: 'right', lineSpacingMultiple: 1.5 });
  meter(s, 8.801, 5.515, 3.56, 0.112, 2.682, C.black, C.green);
  text(s, 'Project', { x: 8.717, y: 5.91, w: 1.311, h: 0.353, fontSize: 11, color: C.white, lineSpacingMultiple: 1.5 });
  text(s, '141+', { x: 11.696, y: 6.002, w: 0.679, h: 0.326, fontSize: 10, color: C.white, align: 'right', lineSpacingMultiple: 1.5 });
  meter(s, 8.809, 6.405, 3.462, 0.112, 1.952, C.black, C.green);
  chrome(s);
}

function slide19(s) {
  const team = [
    { x: 0.667, fill: C.green, name: 'Linda Brown', nx: 1.111, nw: 1.267, role: 'CEO', rx: 1.208, rw: 0.978, ry: 5.948, fg: C.black },
    { x: 3.86, fill: C.purple, name: 'Michael Brown', nx: 3.773, nw: 2.402, role: 'Marketing Project', rx: 4.181, rw: 1.586, ry: 5.931, fg: C.white },
    { x: 7.32, fill: C.green, name: 'James Vane', nx: 7.333, nw: 2.074, role: 'Business Analysis', rx: 7.333, rw: 2.074, ry: 5.901, fg: C.black },
    { x: 10.525, fill: C.purple, name: 'Nelly Dean', nx: 10.721, nw: 1.811, role: 'Brand Manager', rx: 10.605, rw: 2.074, ry: 5.919, fg: C.white }
  ];
  team.forEach(function (v, i) {
    card(s, v.x, i < 2 ? 5.577 : 5.574, 2.154, 0.916, v.fill, 13657);
    photo(s, [0.667, 3.857, 7.32, 10.508][i], 3.271, 2.151, 2.139, 8453);
    text(s, v.name, { x: v.nx, y: i === 0 ? 5.731 : (i === 1 ? 5.687 : (i === 2 ? 5.684 : 5.729)), w: v.nw, h: 0.303, fontSize: 12, color: v.fg, align: 'center' });
    text(s, v.role, { x: v.rx, y: v.ry, w: v.rw, h: 0.352, fontSize: 11, color: v.fg, align: 'center', lineSpacingMultiple: 1.5 });
  });
  text(s, 'Super Team', { x: 4.142, y: 1.732, w: 5.05, h: 0.841, fontSize: 44, align: 'center' });
  chrome(s);
}

function slide20(s) {
  card(s, 8.539, 2.308, 2.935, 1.678, C.purple, 11510);
  card(s, 5.878, 4.216, 5.595, 1.982, C.green, 8835);
  photo(s, 1.591, 2.308, 3.575, 3.89, 8164);
  text(s, 'Thank You', { x: 5.717, y: 2.023, w: 2.822, h: 2.121, fontSize: 60 });
  text(s, 'Get In Touch', { x: 8.746, y: 2.452, w: 2.615, h: 0.505, fontSize: 24 });
  text(s, 'In conclusion', { x: 8.751, y: 2.872, w: 1.814, h: 0.352, fontSize: 11, bold: true, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, 'In the heart of the verdant.', { x: 8.751, y: 3.17, w: 1.232, h: 0.629, fontSize: 11, align: 'justify', lineSpacingMultiple: 1.5 });
  const contact = [
    { t: 'Our Address', v: '123 Main Street, Any City, State, Country 123456', x: 6.234, ty: 4.498, vy: 4.742, tw: 2.725, vw: 2.685, vh: 0.471 },
    { t: 'Our Telephone', v: '+11 222 3333 4444', x: 9.366, ty: 4.491, vy: 4.735, tw: 2.058, vw: 1.99, vh: 0.286 },
    { t: 'Our Email', v: 'Yourcompany@email.com', x: 6.234, ty: 5.42, vy: 5.638, tw: 1.888, vw: 2.975, vh: 0.286 },
    { t: 'Our Website', v: 'www.Yourcompany.com', x: 9.366, ty: 5.419, vy: 5.636, tw: 2.274, vw: 2.274, vh: 0.286 }
  ];
  contact.forEach(function (v) {
    text(s, v.t, { x: v.x, y: v.ty, w: v.tw, h: 0.286, fontSize: 11, bold: true, align: 'justify' });
    text(s, v.v, { x: v.x, y: v.vy, w: v.vw, h: v.vh, fontSize: 11, align: 'justify' });
  });
  chrome(s);
}

/* ---------------------------------------------------------------- build */

const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07,
  slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15,
  slide16, slide17, slide18, slide19, slide20];

const pres = new PptxGenJS();
pres.defineLayout({ name: 'DECK', width: 13.333, height: 7.5 });
pres.layout = 'DECK';
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };

BUILDERS.forEach(function (build) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  build(s);
});

pres.writeFile({ fileName: path.join(__dirname, '182b88c9-2b88-4a90-b83e-4c177a4b4396_grok_final.pptx') })
  .then(function (f) { console.log('wrote ' + f); });
