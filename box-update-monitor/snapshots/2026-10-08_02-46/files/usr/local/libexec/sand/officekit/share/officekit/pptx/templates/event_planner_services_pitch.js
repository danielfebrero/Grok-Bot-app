/**
 * "Event Planner" presentation template - 30 slides, 10 x 5.625 in.
 * Rebuilt with pptxgenjs only. Raster artwork from the original deck is
 * replaced by native vector placeholders.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const C = {
  pink: 'FF3A65',   // accent1
  blue: '2A3AFF',   // accent2
  yellow: 'FFC229', // accent3
  white: 'FFFFFF',
  gray: '808080',
  silver: 'A6A6A6',
  smoke: 'D9D9D9',
  light: 'F2F2F2',
  ink: '000000',
};

const HEAD = 'Merriweather'; // major latin font
const BODY = 'Open Sans';    // minor latin font

// pptxgenjs rewrites the shadow object in place while serialising, so each
// shape needs its own copy.
const SHADOW_SM = () => ({ type: 'outer', color: C.ink, opacity: 0.2, blur: 25, offset: 0, angle: 90 });
const SHADOW_LG = () => ({ type: 'outer', color: C.ink, opacity: 0.3, blur: 35, offset: 0, angle: 90 });
const NOLINE = { type: 'none' };

/** PowerPoint's default roundRect corner radius (adj = 16.667% of the short side). */
const rounded = (w, h) => 0.16667 * Math.min(w, h);

/** "RRGGBB" or "RRGGBB@aa" (aa = alpha in percent) -> pptxgenjs fill object. */
function col(spec) {
  const [hex, alpha] = String(spec).split('@');
  return alpha === undefined ? { color: hex } : { color: hex, transparency: 100 - Number(alpha) };
}

/* -------------------------------------------------- confetti / ribbon art */

// Normalised outlines (0..1 of the shape box) of the three scattered
// "party ribbon" glyphs that decorate every slide.
const RIBBON = [ // big curled streamer, box ratio 0.732
  ['M', 0.670, 0.000], ['C', 0.627, 0.044, 0.522, 0.096, 0.384, 0.137],
  ['C', 0.589, 0.212, 0.707, 0.330, 0.776, 0.434], ['C', 0.677, 0.424, 0.544, 0.405, 0.370, 0.370],
  ['C', 0.370, 0.370, 0.271, 0.391, 0.241, 0.545], ['C', 0.238, 0.564, 0.238, 0.576, 0.238, 0.576],
  ['C', 0.387, 0.651, 0.490, 0.713, 0.563, 0.762], ['C', 0.410, 0.790, 0.210, 0.799, 0.000, 0.731],
  ['C', 0.047, 0.835, 0.056, 0.926, 0.037, 0.980], ['C', 0.514, 1.058, 0.715, 0.884, 0.715, 0.884],
  ['C', 0.715, 0.884, 0.872, 0.729, 0.847, 0.679], ['C', 0.797, 0.657, 0.750, 0.635, 0.707, 0.614],
  ['C', 0.758, 0.624, 0.808, 0.638, 0.864, 0.653], ['C', 0.934, 0.634, 0.997, 0.448, 0.997, 0.448],
  ['C', 0.997, 0.448, 1.053, 0.219, 0.670, 0.000], ['Z'],
  ['M', 0.784, 0.701], ['C', 0.746, 0.714, 0.696, 0.731, 0.634, 0.746],
  ['C', 0.556, 0.690, 0.455, 0.631, 0.337, 0.569], ['C', 0.389, 0.569, 0.464, 0.572, 0.557, 0.586],
  ['C', 0.615, 0.621, 0.692, 0.661, 0.784, 0.701], ['Z'],
  ['M', 0.935, 0.431], ['L', 0.930, 0.445], ['C', 0.907, 0.445, 0.881, 0.443, 0.847, 0.442],
  ['C', 0.789, 0.343, 0.688, 0.229, 0.516, 0.142], ['C', 0.580, 0.118, 0.634, 0.090, 0.675, 0.063],
  ['C', 0.956, 0.239, 0.941, 0.406, 0.935, 0.431], ['Z'],
];
const SWIRL = [ // small curled streamer, box ratio 0.712
  ['M', 0.969, 0.421], ['C', 0.808, 0.360, 0.572, 0.325, 0.372, 0.305],
  ['C', 0.500, 0.211, 0.648, 0.125, 0.769, 0.123], ['C', 0.616, 0.023, 0.399, 0.000, 0.399, 0.000],
  ['C', 0.182, 0.060, 0.005, 0.281, 0.005, 0.281], ['C', -0.022, 0.428, 0.079, 0.561, 0.079, 0.561],
  ['C', 0.135, 0.553, 0.192, 0.549, 0.251, 0.549], ['C', 0.355, 0.549, 0.456, 0.561, 0.550, 0.577],
  ['C', 0.409, 0.656, 0.273, 0.770, 0.202, 0.930], ['C', 0.204, 0.923, 0.522, 1.000, 0.522, 1.000],
  ['C', 0.522, 1.000, 0.643, 0.770, 0.942, 0.684], ['C', 0.942, 0.684, 1.053, 0.516, 0.969, 0.421], ['Z'],
  ['M', 0.251, 0.479], ['C', 0.222, 0.479, 0.190, 0.481, 0.162, 0.482],
  ['C', 0.197, 0.449, 0.241, 0.411, 0.293, 0.370], ['C', 0.520, 0.389, 0.705, 0.419, 0.840, 0.456],
  ['C', 0.789, 0.472, 0.725, 0.495, 0.656, 0.526], ['C', 0.532, 0.500, 0.392, 0.479, 0.251, 0.479], ['Z'],
];
const FLAKE = [['M', 1, 0], ['L', 0.273, 0.273], ['L', 0, 1], ['L', 0.727, 0.727], ['Z']];

const GLYPH = { S: RIBBON, M: SWIRL, D: FLAKE };

/** Turn a normalised outline into pptxgenjs custGeom points scaled to w x h. */
function outline(pathData, w, h) {
  return pathData.map((seg) => {
    if (seg[0] === 'Z') return { close: true };
    if (seg[0] === 'M') return { x: seg[1] * w, y: seg[2] * h, moveTo: true };
    if (seg[0] === 'L') return { x: seg[1] * w, y: seg[2] * h };
    return {
      x: seg[5] * w, y: seg[6] * h,
      curve: { type: 'cubic', x1: seg[1] * w, y1: seg[2] * h, x2: seg[3] * w, y2: seg[4] * h },
    };
  });
}

/** Draw one confetti piece: [glyph, x, y, w, h, rotation, flipH]. */
function piece(s, [g, x, y, w, h, rot, flipH], color) {
  s.addShape('custGeom', {
    x, y, w, h, rotate: rot || 0, flipH: !!flipH,
    fill: col(color), line: NOLINE, points: outline(GLYPH[g], w, h),
  });
}

// The twelve confetti positions reused across the deck; each slide picks a
// subset and gives every spot its own colour.
const SPOT = {
  A: [['S', 8.894, 1.448, 0.236, 0.323, 195]],
  B: [['S', 2.754, 4.722, 0.175, 0.240, 0]],
  C: [['D', 7.171, 0.303, 0.066, 0.066, 0], ['M', 6.980, 0.298, 0.239, 0.336, 0], ['D', 6.999, 0.687, 0.088, 0.088, 0]],
  D: [['M', 5.964, 4.550, 0.179, 0.252, 0], ['D', 6.155, 4.554, 0.066, 0.066, 0], ['D', 5.978, 4.842, 0.066, 0.066, 0]],
  E: [['S', 7.665, 2.150, 0.319, 0.436, 321.7]],
  F: [['M', 0.959, 3.400, 0.317, 0.445, 21.3], ['D', 1.337, 3.505, 0.117, 0.117, 21.3], ['D', 0.861, 3.865, 0.117, 0.117, 21.3]],
  G: [['S', 8.564, 4.325, 0.236, 0.323, 0]],
  H: [['S', 0.988, 1.077, 0.319, 0.436, 0]],
  I: [['S', 3.117, 3.101, 0.319, 0.436, 0, true]],
  J: [['M', 5.874, 2.493, 0.302, 0.424, 21.3], ['D', 6.234, 2.593, 0.112, 0.112, 21.3], ['D', 5.781, 2.936, 0.112, 0.112, 21.3]],
  K: [['S', 3.857, 0.873, 0.319, 0.436, 0]],
  L: [['M', 2.523, 0.317, 0.179, 0.252, 345], ['D', 2.590, 0.613, 0.066, 0.066, 345]],
};

/** `spots` maps spot letter -> colour; `extras` are one-off [pieces, colour]. */
function confetti(s, spots, extras) {
  Object.keys(spots).forEach((k) => SPOT[k].forEach((p) => piece(s, p, spots[k])));
  (extras || []).forEach(([group, color]) => group.forEach((p) => piece(s, p, color)));
}

/* ----------------------------------------------------------- text helpers */

/** Body copy: Open Sans, 8.25pt grey, 150% leading, top aligned. */
function txt(s, text, o) {
  s.addText(text, Object.assign({
    fontFace: BODY, fontSize: 8.25, color: C.gray, valign: 'top',
    lineSpacingMultiple: 1.5, align: 'left',
  }, o));
}

/** Section heading: Merriweather bold 27pt, black. */
function heading(s, text, o) {
  s.addText(text, Object.assign({
    fontFace: HEAD, fontSize: 27, bold: true, color: C.ink, valign: 'top', align: 'left',
  }, o));
}

/** White (or coloured) card with the deck's soft drop shadow. */
function card(s, x, y, w, h, o) {
  o = o || {};
  // a roundRect with a zero adjust is just a rectangle
  const shape = o.shape === 'roundRect' && !o.rectRadius ? 'rect' : (o.shape || 'rect');
  s.addShape(shape, Object.assign({
    x, y, w, h, fill: col(o.fill || C.white), line: NOLINE,
    shadow: o.shadow === undefined ? SHADOW_SM() : o.shadow,
  }, o.rotate !== undefined ? { rotate: o.rotate } : {},
     o.flipH ? { flipH: true } : {},
     o.rectRadius !== undefined ? { rectRadius: o.rectRadius } : {}));
}

/** Pill button, e.g. "Read More" / "More Info". */
function pill(s, x, y, w, h, text, o) {
  o = o || {};
  s.addShape('roundRect', {
    x, y, w, h, rectRadius: o.radius === undefined ? rounded(w, h) : o.radius,
    fill: o.fill ? col(o.fill) : { type: 'none' },
    line: o.line ? { color: o.line, width: 1 } : NOLINE,
  });
  if (text) {
    s.addText(text, {
      x, y, w, h, align: 'center', valign: 'middle',
      fontFace: BODY, fontSize: o.size || 8.25, bold: o.bold !== false, color: o.color || C.white,
    });
  }
}

/* --------------------------------------------- page chrome (every slide) */

/** Small social badge: rounded pill plus four dots. */
function badge(s, x, y, w, h, fill, dot) {
  s.addShape('roundRect', { x, y, w, h, rectRadius: h / 2, fill: col(fill), line: NOLINE });
  const d = w * 0.119;
  [0.099, 0.315, 0.519, 0.734].forEach((f) => {
    s.addShape('ellipse', { x: x + w * f, y: y + (h - d) / 2, w: d, h: d, fill: col(dot), line: NOLINE });
  });
}

/**
 * Standard furniture: brand label (top left), page number (bottom left),
 * "Event Planner --o-- +52 ..." footer (bottom right) and the social badge.
 */
function chrome(s, o) {
  o = o || {};
  const foot = o.foot || C.gray;
  const dash = o.dash || C.silver;
  const dy = o.dy || 0;
  if (o.label !== false) {
    s.addText('Event Planner', {
      x: 0.11, y: 0.2, w: 0.994, h: 0.252, fontFace: HEAD, fontSize: 9,
      color: o.label || C.pink, valign: 'top', wrap: false,
    });
  }
  if (o.num) {
    s.addText(o.num, {
      x: 0.11, y: 5.013, w: 0.577, h: 0.505, fontFace: HEAD, fontSize: 24,
      color: o.numColor || C.yellow, valign: 'top', wrap: false,
    });
  }
  s.addText('Event Planner', {
    x: 7.108, y: 5.231, w: 0.931, h: 0.24, align: 'right', fontFace: BODY, fontSize: 8.25,
    color: foot, valign: 'top', wrap: false,
  });
  s.addShape('line', { x: 8.101, y: 5.338 + dy, w: 0.545, h: 0, line: { color: dash, width: 1.5 } });
  s.addText('+52 1921 1823 12', {
    x: 8.686, y: 5.231, w: 1.143, h: 0.24, align: 'right', fontFace: BODY, fontSize: 8.25,
    color: o.tel || C.pink, valign: 'top', wrap: false,
  });
  s.addShape('ellipse', {
    x: 8.182, y: 5.293 + dy, w: 0.09, h: 0.09,
    fill: col(o.ovalFill || C.pink), line: { color: o.ovalLine || C.silver, width: 1.5 },
  });
  badge(s, 9.081, 0.224, 0.749, 0.203, o.badge || C.blue, o.badgeDot || C.white);
}

/* ------------------------------------------------------- vector icon set */

/** Monochrome pictograms that stand in for the deck's icon artwork. */
function icon(s, kind, x, y, sz, color) {
  const f = col(color);
  const add = (shape, ox, oy, ow, oh, extra) =>
    s.addShape(shape, Object.assign({ x: x + ox * sz, y: y + oy * sz, w: ow * sz, h: oh * sz, fill: f, line: NOLINE }, extra));
  if (kind === 'people') {
    add('ellipse', 0.34, 0.14, 0.32, 0.32);
    add('ellipse', 0.02, 0.24, 0.24, 0.24);
    add('ellipse', 0.74, 0.24, 0.24, 0.24);
    add('roundRect', 0.26, 0.50, 0.48, 0.34, { rectRadius: 0.08 * sz });
    add('roundRect', 0.00, 0.55, 0.24, 0.26, { rectRadius: 0.06 * sz });
    add('roundRect', 0.76, 0.55, 0.24, 0.26, { rectRadius: 0.06 * sz });
  } else if (kind === 'bulb') {
    add('ellipse', 0.22, 0.10, 0.56, 0.56);
    add('rect', 0.38, 0.60, 0.24, 0.20);
    add('rect', 0.40, 0.84, 0.20, 0.08);
  } else if (kind === 'photo') {
    add('rect', 0.30, 0.05, 0.62, 0.62, { rotate: 12 });
    add('rect', 0.05, 0.25, 0.62, 0.62, { rotate: -8 });
  } else if (kind === 'award') {
    add('ellipse', 0.16, 0.00, 0.68, 0.68);
    add('rect', 0.28, 0.60, 0.16, 0.40, { rotate: 10 });
    add('rect', 0.56, 0.60, 0.16, 0.40, { rotate: -10 });
  } else if (kind === 'city') {
    add('rect', 0.02, 0.34, 0.28, 0.62);
    add('rect', 0.36, 0.10, 0.28, 0.86);
    add('rect', 0.70, 0.46, 0.28, 0.50);
  } else if (kind === 'globe') {
    add('ellipse', 0.00, 0.00, 1.00, 1.00);
  } else if (kind === 'burst') {
    for (let i = 0; i < 8; i++) {
      const a = (i * Math.PI) / 4;
      add('ellipse', 0.46 + 0.40 * Math.cos(a), 0.46 + 0.40 * Math.sin(a), 0.09, 0.09);
    }
    add('ellipse', 0.40, 0.40, 0.20, 0.20);
  }
}

/* ------------------------------------------------------------------ deck */

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'EVENT', width: 10, height: 5.625 });
pptx.layout = 'EVENT';
pptx.author = 'Event Planner';
pptx.title = 'Event Planner - Presentation Templates';

const LOREM = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. '
  + 'Donec commodo et urna ac semper. Mauris finibus augue id vulputate';
const LOREM_MOLESTIE = LOREM + ' volutpat. Vestibu lum eros arcu, maximus ut molestie vel, ';
const LOREM_DIGNISSIM = LOREM_MOLESTIE + 'dignissim vitae elit. ';
const LOREM_CONSEC = LOREM + ' volutpat. amet, consectetur consectetur adipi scing elit. Vivamus vel euismod leo. Donec urna ';
const LOREM_VIVAMUS = LOREM + ' sit amet, consectetur adipiscing elit. Vivamus  euismod leo. Donec commodo et urna ac ';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipis cing elit. Vivamus vel euismod leo. ';
const LOREM_CARD = 'Lorem ipsum dolor sit amet, consecte tur adipiscing elit. Vivamus';

/* =========================================================== slide 01/30 */
function slide01() {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: 10, h: 5.625, fill: col(C.pink + '@80'), line: NOLINE });
  badge(s, 9.081, 0.224, 0.749, 0.203, C.blue, C.white);

  // footer sits bottom-left on the cover
  s.addText('Event Planner', { x: 0.112, y: 5.217, w: 0.931, h: 0.24, align: 'right', fontFace: BODY, fontSize: 8.25, color: C.white, valign: 'top', wrap: false });
  s.addShape('line', { x: 1.089, y: 5.321, w: 0.545, h: 0, line: { color: C.white, width: 1.5 } });
  s.addText('+52 1921 1823 12', { x: 1.69, y: 5.217, w: 1.143, h: 0.24, align: 'right', fontFace: BODY, fontSize: 8.25, color: C.white, valign: 'top', wrap: false });
  s.addShape('ellipse', { x: 1.186, y: 5.28, w: 0.09, h: 0.09, fill: col(C.yellow), line: { color: C.white, width: 1.5 } });

  icon(s, 'burst', 4.646, 0.901, 0.708, C.yellow);
  s.addShape('line', { x: 2.346, y: 3.5, w: 5.307, h: 0, line: { color: C.white, width: 1 } });
  s.addText('EVENT PLANNER', { x: 2.829, y: 1.728, w: 4.343, h: 1.616, align: 'center', valign: 'top', fontFace: HEAD, fontSize: 45, bold: true, color: C.white });
  txt(s, 'PRESENTATION TEMPLATES', { x: 3.632, y: 3.585, w: 2.737, h: 0.328, align: 'center', fontSize: 9, color: C.white });
  txt(s, LOREM.replace(' id vulputate', ''), { x: 2.832, y: 3.9, w: 4.336, h: 0.518, align: 'center', color: C.white });

  confetti(s, {}, [
    [[['S', 8.808, 1.294, 0.236, 0.323, 195]], C.white],
    [SPOT.B, C.ink + '@20'],
    [[['D', 2.557, 0.603, 0.066, 0.066, 0], ['M', 2.366, 0.598, 0.239, 0.336, 0], ['D', 2.385, 0.987, 0.088, 0.088, 0]], C.white],
    [[['S', 7.588, 4.735, 0.236, 0.323, 0]], C.white],
    [[['M', 7.487, 0.656, 0.179, 0.252, 345], ['D', 7.554, 0.952, 0.066, 0.066, 345]], C.ink + '@20'],
    [[['S', 0.776, 1.350, 0.319, 0.436, 0]], C.blue],
    [[['S', 2.405, 2.088, 0.319, 0.436, 0, true]], C.pink],
    [SPOT.E, C.ink + '@20'],
    [[['M', 0.740, 3.735, 0.302, 0.424, 21.3], ['D', 1.100, 3.835, 0.112, 0.112, 21.3], ['D', 0.648, 4.177, 0.112, 0.112, 21.3]], C.yellow],
    [[['M', 5.874, 2.303, 0.302, 0.424, 21.3], ['D', 6.234, 2.402, 0.112, 0.112, 21.3], ['D', 5.781, 2.745, 0.112, 0.112, 21.3]], C.blue],
    [[['D', 9.185, 3.499, 0.066, 0.066, 0], ['M', 8.994, 3.494, 0.239, 0.336, 0], ['D', 9.013, 3.883, 0.088, 0.088, 0]], C.yellow],
    [[['M', 2.187, 2.277, 0.179, 0.252, 44.5], ['D', 2.074, 2.482, 0.066, 0.066, 44.5]], C.ink + '@20'],
  ]);
}

/* =========================================================== slide 02/30 */
function slide02() {
  const s = pptx.addSlide();
  confetti(s, { A: C.light, B: C.pink, F: C.light, C: C.light, G: C.blue, L: C.yellow, H: C.light, I: C.light, E: C.light, K: C.light, D: C.light, J: C.light },
    [[[['M', 7.038, 2.144, 0.302, 0.424, 21.3], ['D', 7.398, 2.243, 0.112, 0.112, 21.3], ['D', 6.945, 2.586, 0.112, 0.112, 21.3]], C.smoke]]);
  chrome(s, { num: '02' });

  card(s, 1.728, 1.298, 1.806, 3.03);
  card(s, 6.466, 1.364, 1.806, 3.03, { flipH: true });
  card(s, 3.303, 0.813, 3.395, 4.0, { fill: C.pink });
  s.addText('Welcome Message', { x: 3.303, y: 1.231, w: 3.395, h: 1.01, align: 'center', valign: 'top', fontFace: HEAD, fontSize: 27, bold: true, color: C.white });
  s.addShape('line', { x: 4.599, y: 2.429, w: 0.9, h: 0, line: { color: C.white, width: 2.25 } });
  txt(s, LOREM_MOLESTIE, { x: 3.632, y: 2.642, w: 2.834, h: 1.142, align: 'center', color: C.white });
  pill(s, 4.424, 3.95, 1.25, 0.316, 'More Info', { line: C.white, radius: 0.054, size: 9 });
}

/* =========================================================== slide 03/30 */
function slide03() {
  const s = pptx.addSlide();
  confetti(s, { B: C.blue, F: C.light, K: C.light, G: C.blue, D: C.light, L: C.yellow, H: C.light, I: C.light, J: C.light, E: C.light });
  chrome(s, { num: '03' });

  heading(s, 'The Best Event Planner', { x: 1.094, y: 1.236, w: 3.783, h: 1.01 });
  txt(s, LOREM + ' volutpat. Vestibu lum eros arcu, maximus ut amet, ', { x: 1.094, y: 2.438, w: 3.298, h: 0.934, align: 'justify' });
  txt(s, 'Included Services :', { x: 1.094, y: 3.451, w: 3.298, h: 0.328, align: 'justify', bold: true, fontSize: 9, color: C.ink });
  const check = { breakLine: true, bullet: { characterCode: '2713', indent: 10.15 } };
  txt(s, ['Organizing Services', 'Event Consultation ', 'Creative Planner'].map((t) => ({ text: t, options: check })), {
    x: 1.094, y: 3.772, w: 3.298, h: 0.783, align: 'justify', fontSize: 9,
  });

  card(s, 4.852, 1.435, 1.989, 1.347);
  card(s, 4.852, 2.987, 1.989, 1.347, { fill: C.pink });
  icon(s, 'photo', 5.695, 1.667, 0.303, C.gray);
  icon(s, 'bulb', 5.668, 3.104, 0.357, C.white);
  txt(s, 'Memorable', { x: 5.225, y: 1.959, w: 1.242, h: 0.309, align: 'center', bold: true });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing', { x: 4.951, y: 2.164, w: 1.791, h: 0.518, align: 'center' });
  txt(s, 'Amazing Concepts ', { x: 5.225, y: 3.479, w: 1.242, h: 0.309, align: 'center', bold: true, color: C.white });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing', { x: 4.951, y: 3.706, w: 1.791, h: 0.518, align: 'center', color: C.white });
}

/* =========================================================== slide 04/30 */
function slide04() {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: 3.083, h: 5.625, fill: col(C.blue), line: NOLINE });
  confetti(s, { A: C.light, C: C.pink, K: C.light, G: C.blue, D: C.light, L: C.ink + '@20', H: C.ink + '@20', J: C.light, E: C.light, F: C.light, B: C.ink + '@20' });
  chrome(s, { label: C.white, num: '04', numColor: C.white });

  heading(s, 'What We Do ?', { x: 6.404, y: 1.324, w: 1.786, h: 1.01 });
  txt(s, LOREM_DIGNISSIM, { x: 6.404, y: 2.421, w: 2.913, h: 1.142, align: 'justify' });
  pill(s, 6.498, 3.889, 1.137, 0.345, 'Read More', { fill: C.smoke, color: C.gray, radius: 0.1725 });

  // left "photo" stack stand-ins
  card(s, 0, 3.111, 1.452, 1.072);
  [[3.316, 0.957, 0.131], [3.606, 1.143, 0.102], [3.815, 1.010, 0.115]].forEach(([y, w, h]) =>
    s.addShape('rect', { x: 0, y, w, h, fill: col(C.smoke), line: NOLINE }));

  card(s, 1.583, 3.316, 2.271, 1.207);
  txt(s, 'GRADUATE PARTY', { x: 1.3, y: 3.383, w: 2.913, h: 0.328, align: 'center', bold: true, fontSize: 9, color: C.yellow });
  txt(s, 'Lorem ipsum dolor sit amet, consec te tur adipiscing elit. Vivamus vel euismod leo. Donec', { x: 1.701, y: 3.661, w: 2.153, h: 0.726, align: 'center' });

  card(s, 3.957, 3.222, 2.104, 1.072);
  txt(s, 'WEDDING PARTY', { x: 3.562, y: 3.383, w: 2.913, h: 0.328, align: 'center', bold: true, fontSize: 9 });
  txt(s, 'Lorem ipsum dolor sit amet, consec tetur adipiscing', { x: 4.027, y: 3.661, w: 2.022, h: 0.518, align: 'center' });
}

/* =========================================================== slide 05/30 */
function slide05() {
  const s = pptx.addSlide();
  confetti(s, { A: C.light, B: C.blue, F: C.light, C: C.yellow, K: C.light, G: C.light, D: C.light, L: C.light, H: C.light, I: C.light, E: C.light });
  s.addShape('rect', { x: 6.981, y: 2.239, w: 3.019, h: 1.346, fill: col(C.pink), line: NOLINE });
  chrome(s, { num: '05' });

  heading(s, 'Why Choose Us ?', { x: 0.667, y: 1.361, w: 3.544, h: 0.555 });
  txt(s, LOREM_DIGNISSIM, { x: 0.667, y: 2.129, w: 3.544, h: 0.934, align: 'justify' });
  txt(s, 'Simple Text Here', { x: 0.667, y: 3.254, w: 3.544, h: 0.309, align: 'justify', bold: true, color: C.ink });
  txt(s, LOREM + ' volutpat. ', { x: 0.667, y: 3.585, w: 3.544, h: 0.726, align: 'justify' });

  const rows = [
    ['SOLID TEAM', 1.178, 1.45, C.gray, 'people', 9.034, 1.194, 0.226],
    ['INNOVATIVE', 2.505, 2.777, C.white, 'bulb', 9.032, 2.488, 0.229],
    ['THE BEST', 3.832, 4.104, C.gray, 'award', 8.985, 3.794, 0.322],
  ];
  rows.forEach(([label, yTitle, yBody, color, kind, ix, iy, isz]) => {
    txt(s, label, { x: 7.538, y: yTitle, w: 1.856, h: 0.328, align: 'justify', bold: true, fontSize: 9, color });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit', { x: 7.538, y: yBody, w: 1.856, h: 0.518, align: 'justify', color });
    icon(s, kind, ix, iy, isz, color);
  });
}

/* =========================================================== slide 06/30 */
function slide06() {
  const s = pptx.addSlide();
  confetti(s, { J: C.light, A: C.yellow, B: C.light, F: C.light, C: C.light, K: C.light, G: C.light, D: C.pink, L: C.blue, E: C.light },
    [[[['S', 1.012, 1.077, 0.319, 0.436, 0]], C.light]]);
  chrome(s, { num: '06' });

  heading(s, 'About Event Planner Costing', { x: 6.286, y: 2.145, w: 3.544, h: 1.01 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur', { x: 6.286, y: 3.114, w: 2.706, h: 0.309, align: 'justify' });

  const tags = [
    ['$ 100', 0.757, 0.917, C.yellow, 0.463, 3.432, 0.315, 3.665, C.yellow],
    ['$ 180', 4.763, 0.917, C.pink, 4.469, 3.432, 4.321, 3.665, C.pink],
    ['$ 159', 2.736, 4.279, C.blue, 2.442, 1.399, 2.294, 1.632, C.blue],
  ];
  tags.forEach(([price, px, py, fill, tx, ty, bx, by, tc]) => {
    s.addShape('roundRect', { x: px, y: py, w: 0.83, h: 0.534, rectRadius: 0.0626, fill: col(fill), line: NOLINE });
    s.addText(price, { x: px, y: py, w: 0.83, h: 0.534, align: 'center', valign: 'middle', fontFace: BODY, fontSize: 10.5, bold: true, color: C.white });
    txt(s, 'Simple Text Here', { x: tx, y: ty, w: 1.418, h: 0.309, align: 'center', bold: true, color: tc });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod', { x: bx, y: by, w: 1.714, h: 0.726, align: 'center' });
  });
}

/* =========================================================== slide 07/30 */
function slide07() {
  const s = pptx.addSlide();
  confetti(s, { B: C.blue, F: C.light, C: C.yellow, K: C.light, L: C.light, H: C.light, I: C.light },
    [[SPOT.D, C.light]]);
  chrome(s, { num: '07' });

  heading(s, [{ text: 'About Our ', options: { breakLine: true } }, { text: 'Vision Company' }], { x: 1.139, y: 1.243, w: 3.544, h: 1.01 });
  txt(s, LOREM + ' volutpat. amet, consectetur', { x: 1.139, y: 2.632, w: 3.847, h: 0.726, align: 'justify' });
  txt(s, LOREM_CONSEC, { x: 1.139, y: 3.495, w: 3.847, h: 0.934, align: 'justify' });

  s.addShape('ellipse', { x: 5.439, y: 3.286, w: 1.304, h: 1.304, fill: col(C.pink), line: NOLINE, shadow: SHADOW_SM() });
  s.addText('+516', { x: 5.432, y: 3.61, w: 1.318, h: 0.454, align: 'center', valign: 'top', fontFace: HEAD, fontSize: 21, bold: true, color: C.white });
  txt(s, 'Growth Business', { x: 5.423, y: 3.993, w: 1.336, h: 0.309, align: 'center', bold: true, color: C.white });
}

/* =========================================================== slide 08/30 */
function slide08() {
  const s = pptx.addSlide();
  confetti(s, { A: C.light, B: C.light, F: C.light, C: C.light, K: C.yellow, G: C.pink, D: C.light, L: C.yellow, I: C.light, J: C.light, E: C.light },
    [[[['S', 0.988, 1.158, 0.319, 0.436, 0]], C.light]]);
  chrome(s, { num: '08' });

  s.addShape('rect', { x: 0.73, y: 1.087, w: 3.081, h: 3.594, fill: col(C.blue), line: NOLINE });
  card(s, 2.884, 0.767, 2.584, 4.092);
  icon(s, 'people', 1.195, 1.469, 0.375, C.white);
  txt(s, 'Good Services', { x: 1.123, y: 1.857, w: 1.405, h: 0.328, align: 'justify', bold: true, fontSize: 9, color: C.white });
  txt(s, 'Lorem ipsum dolor sit amet, consect etur', { x: 1.123, y: 2.13, w: 1.445, h: 0.518, align: 'justify', color: C.white });
  icon(s, 'city', 1.165, 2.985, 0.362, C.white);
  txt(s, 'Free Consultation', { x: 1.123, y: 3.324, w: 1.405, h: 0.328, align: 'justify', bold: true, fontSize: 9, color: C.white });
  txt(s, 'Lorem ipsum dolor sit amet, consect etur', { x: 1.123, y: 3.601, w: 1.445, h: 0.518, align: 'justify', color: C.white });

  heading(s, [{ text: 'Event Planner', options: { breakLine: true } }, { text: 'Mission Here' }], { x: 6.062, y: 1.236, w: 3.544, h: 1.01 });
  txt(s, 'Simple Text Here', { x: 6.062, y: 2.441, w: 3.544, h: 0.328, align: 'justify', bold: true, fontSize: 9, color: C.ink });
  txt(s, LOREM, { x: 6.062, y: 2.812, w: 3.217, h: 0.726, align: 'justify' });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et', { x: 6.053, y: 3.533, w: 3.217, h: 0.518, align: 'justify' });
  pill(s, 6.146, 4.208, 1.137, 0.345, 'Read More', { fill: C.smoke, color: C.gray, radius: 0.1725 });
}

/* =========================================================== slide 09/30 */
function slide09() {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 6.583, y: 0, w: 3.417, h: 5.625, fill: col(C.yellow), line: NOLINE });
  confetti(s, { F: C.light, I: C.light, A: C.light, B: C.pink, C: C.ink + '@30', K: C.blue, G: C.ink + '@30', L: C.light, H: C.light, E: C.light, D: C.light, J: C.light });
  chrome(s, { num: '09', foot: C.white, tel: C.white, dash: C.white, ovalFill: C.yellow, ovalLine: C.white });

  heading(s, 'About Our Services', { x: 1.139, y: 1.243, w: 2.56, h: 1.01 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipis cing elit. Vivamus vel euismod ', { x: 1.129, y: 2.437, w: 2.652, h: 0.518, align: 'justify', italic: true });
  txt(s, LOREM + ' sit amet, consectetur adipiscing elit. Vivamus  euismod leo. Donec commodo et urna ac semper. M', { x: 1.139, y: 3.188, w: 2.387, h: 1.161, align: 'justify', fontSize: 7 });

  card(s, 5.327, 0.847, 4.172, 3.536);
  card(s, 4.281, 2.525, 2.422, 2.314);
}

/* =========================================================== slide 10/30 */
function slide10() {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: -0.03, w: 5, h: 5.655, fill: col(C.yellow), line: NOLINE });
  s.addShape('ellipse', { x: 2.685, y: 0.497, w: 4.631, h: 4.631, fill: col(C.white), line: NOLINE, shadow: SHADOW_SM() });
  confetti(s, { A: C.light, F: C.ink + '@20', C: C.light, G: C.blue, L: C.pink, H: C.ink + '@20', I: C.light, E: C.light, B: C.ink + '@20' });
  chrome(s, { label: C.white, num: '10', numColor: C.white });

  // four service cards: [cardX, cardY, align, tabX, heartX, titleRuns, textX]
  const items = [
    { x: 0.676, y: 1.417, align: 'right', tab: 3.481, tabFill: C.pink, heart: 0.817, heartFill: C.pink, tx: 0.744, ty: 1.608, bx: 0.762, by: 1.883,
      title: [{ text: 'About', options: { color: C.pink } }, { text: ' Service 01', options: { color: C.gray } }] },
    { x: 6.428, y: 1.417, align: 'left', tab: 6.35, tabFill: C.smoke, heart: 8.976, heartFill: C.smoke, tx: 6.795, ty: 1.608, bx: 6.795, by: 1.883,
      title: 'About Service 02' },
    { x: 1.2, y: 3.135, align: 'right', tab: 4.004, tabFill: C.smoke, heart: 1.37, heartFill: C.smoke, tx: 1.25, ty: 3.329, bx: 1.285, by: 3.59,
      title: 'About Service 03' },
    { x: 5.837, y: 3.049, align: 'left', tab: 5.76, tabFill: C.smoke, heart: 8.44, heartFill: C.smoke, tx: 6.204, ty: 3.243, bx: 6.204, by: 3.503,
      title: 'About Service 04' },
  ];
  items.forEach((it) => {
    card(s, it.x, it.y, 2.896, 1.159);
    s.addShape('roundRect', { x: it.tab, y: it.y + 0.274, w: 0.155, h: 0.612, rectRadius: rounded(0.155, 0.612), fill: col(it.tabFill), line: NOLINE });
    heart(s, it.heart, it.y + 0.16, 0.179, it.heartFill);
    txt(s, it.title, { x: it.tx, y: it.ty, w: 2.528, h: 0.366, align: it.align, bold: true, fontSize: 10.5 });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel', { x: it.bx, y: it.by, w: 2.528, h: 0.518, align: it.align });
  });
}

/** Small heart marker used on the service cards. */
function heart(s, x, y, w, color) {
  s.addShape('heart', { x, y, w, h: w * 0.8, fill: col(color), line: NOLINE });
}

/* =========================================================== slide 11/30 */
function slide11() {
  const s = pptx.addSlide();
  confetti(s, { A: C.light, B: C.blue, F: C.light, C: C.light, K: C.yellow, D: C.light, L: C.yellow, H: C.light, I: C.light, J: C.light, E: C.light });
  chrome(s, { num: '11' });

  s.addShape('rect', { x: 3.34, y: 0, w: 2.142, h: 5.625, fill: col(C.pink), line: NOLINE });
  heading(s, 'Worldwide Events With A Personal Touch', { x: 0.692, y: 1.245, w: 2.744, h: 1.919 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipis cing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. ', { x: 0.692, y: 3.253, w: 2.429, h: 0.726, align: 'justify' });
  pill(s, 0.769, 4.107, 1.137, 0.345, 'Read More', { fill: C.smoke, color: C.gray, radius: 0.1725 });

  icon(s, 'globe', 4.278, 0.956, 0.266, C.white);
  txt(s, 'Event Planner 01', { x: 3.437, y: 1.308, w: 1.948, h: 0.309, align: 'center', bold: true, color: C.white });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipis cing elit. Vivamus vel euismod', { x: 3.437, y: 1.612, w: 1.948, h: 0.726, align: 'center', color: C.white });
  s.addShape('line', { x: 4.175, y: 2.461, w: 0.472, h: 0, line: { color: C.white, width: 2.25 } });

  s.addShape('custGeom', { // bookmark ribbon
    x: 4.362, y: 3.111, w: 0.099, h: 0.253, fill: col(C.white), line: NOLINE,
    points: outline([['M', 0, 0], ['L', 1, 0], ['L', 1, 1], ['L', 0.5, 0.78], ['L', 0, 1], ['Z']], 0.099, 0.253),
  });
  txt(s, 'Event Planner 02', { x: 3.437, y: 3.45, w: 1.948, h: 0.309, align: 'center', bold: true, color: C.white });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipis cing elit. Vivamus vel euismod', { x: 3.437, y: 3.794, w: 1.948, h: 0.726, align: 'center', color: C.white });
  s.addShape('line', { x: 4.175, y: 4.669, w: 0.472, h: 0, line: { color: C.white, width: 2.25 } });
}

/* =========================================================== slide 12/30 */
function slide12() {
  const s = pptx.addSlide();
  confetti(s, { A: C.light, B: C.light, C: C.pink, G: C.light, D: C.blue, L: C.yellow, I: C.light, J: C.light, E: C.light });
  chrome(s, { num: '12' });

  heading(s, 'Meet Our Creative Event Planner Here', { x: 5.774, y: 1.323, w: 3.151, h: 1.464 });
  txt(s, LOREM_CONSEC, { x: 5.791, y: 2.998, w: 3.364, h: 1.142, align: 'justify' });
  s.addShape('line', { x: 5.909, y: 4.304, w: 3.083, h: 0, line: { color: 'BFBFBF', width: 1 } });
  s.addShape('line', { x: 7.238, y: 4.312, w: 1.418, h: 0, line: { color: C.yellow, width: 4.5 } });

  // notched name tags
  const tag = (x, y, fill, name, role) => {
    s.addShape('custGeom', {
      x, y, w: 1.406, h: 0.512, fill: col(fill), line: NOLINE,
      points: outline([['M', 0, 0], ['L', 1, 0], ['L', 1, 1], ['L', 0, 1], ['L', 0.075, 0.504], ['Z']], 1.406, 0.512),
    });
    txt(s, name, { x: x + 0.139, y: y + 0.03, w: 1.177, h: 0.328, align: 'right', bold: true, fontSize: 9, color: C.white });
    txt(s, role, { x: x + 0.139, y: y + 0.198, w: 1.177, h: 0.309, align: 'right', italic: true, color: C.white });
  };
  tag(1.792, 2.128, C.pink, 'Lussiana Maria', 'Ceo & Founder');
  tag(3.813, 2.128, C.yellow, 'Christian Key', 'Marketing Manager');
  tag(1.064, 4.106, C.blue, 'Britany Gerogia', 'Event Planner 01');
  tag(3.850, 4.106, C.pink, 'Violita Swift', 'Event Planner 02');
}

/* =========================================================== slide 13/30 */
function slide13() {
  const s = pptx.addSlide();
  s.background = { color: C.pink };
  confetti(s, { A: C.ink + '@20', B: C.pink, F: C.blue, C: C.ink + '@20', K: C.yellow, G: C.ink + '@20', D: C.ink + '@20', L: C.light, H: C.ink + '@20', J: C.ink + '@20', E: C.ink + '@20', I: C.light },
    [[[['D', 7.037, 1.717, 0.043, 0.043, 0], ['M', 6.915, 1.714, 0.239, 0.336, 0], ['D', 6.934, 2.103, 0.088, 0.088, 0]], C.ink + '@25']]);
  chrome(s, { label: C.white, num: '13', numColor: C.white, foot: C.white, tel: C.white, dash: C.white, ovalFill: C.blue, ovalLine: C.white, badge: C.white, badgeDot: C.pink });

  // three tilted photo cards
  card(s, 0.76, 1.099, 2.633, 1.852, { rotate: -11.6, shadow: SHADOW_LG() });
  card(s, 2.541, 1.62, 2.633, 1.852, { rotate: 11.3, shadow: SHADOW_LG() });
  card(s, 1.565, 2.866, 2.633, 1.852, { shadow: SHADOW_LG() });

  s.addText([{ text: 'Wedding Party', options: { breakLine: true } }, { text: 'Projects Here' }], {
    x: 5.605, y: 1.023, w: 3.1, h: 1.01, valign: 'top', fontFace: HEAD, fontSize: 27, bold: true, color: C.white });
  txt(s, LOREM_VIVAMUS, { x: 5.605, y: 2.216, w: 3.711, h: 0.934, align: 'justify', color: C.white });

  statCard(s, 5.936, 3.384, 4.068, '+300', C.yellow, 'Completed Wedding Party', 5.851, 6.083, 7.060);
}

/**
 * The recurring "big number + caption" white bar with a coloured side tab.
 */
function statCard(s, x, y, w, value, accent, caption, tabX, valueX, textX) {
  s.addShape('roundRect', { x, y, w, h: 0.989, rectRadius: 0, fill: col(C.white), line: NOLINE });
  s.addShape('roundRect', { x: tabX, y: y + 0.251, w: 0.15, h: 0.515, rectRadius: rounded(0.15, 0.515), fill: col(accent), line: NOLINE });
  txt(s, value, { x: valueX, y: y + 0.111, w: 1.457, h: 0.707, align: 'justify', bold: true, fontSize: 24, color: accent });
  txt(s, caption, { x: textX, y: y + 0.11, w: 1.811, h: 0.309, align: 'justify', bold: true, color: accent });
  txt(s, LOREM_CARD, { x: textX, y: y + 0.324, w: 2.168, h: 0.518, align: 'justify' });
}

/* =========================================================== slide 14/30 */
function slide14() {
  const s = pptx.addSlide();
  confetti(s, { A: C.pink, B: C.blue, F: C.light, C: C.yellow, K: C.light, G: C.blue, D: C.light, L: C.light, H: C.light, I: C.light, J: C.light, E: C.light });
  chrome(s, { num: '14' });

  card(s, 4.632, 2.49, 4.862, 2.237);
  heading(s, 'About Wedding Planning Service ', { x: 0.626, y: 3.208, w: 3.532, h: 1.01 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus', { x: 0.695, y: 4.239, w: 3.68, h: 0.309, align: 'justify' });
  s.addShape('line', { x: 7.073, y: 2.812, w: 0, h: 1.635, line: { color: C.smoke, width: 2.25, dashType: 'dash' } });

  [['Wedding Organizer', C.pink, 5.038, 5.127, null], ['Wedding Packages', C.gray, 7.333, 7.436, C.smoke]].forEach(([title, color, tx, bx, btnFill]) => {
    txt(s, title, { x: tx, y: 2.939, w: 1.926, h: 0.328, align: 'justify', bold: true, fontSize: 9, color });
    txt(s, LOREM_SHORT, { x: tx, y: 3.273, w: 1.789, h: 0.726, align: 'justify' });
    s.addShape('rect', { x: bx, y: 4.085, w: 0.979, h: 0.273, fill: col(btnFill || C.pink), line: NOLINE });
    s.addText('View More', { x: bx, y: 4.085, w: 0.979, h: 0.273, align: 'center', valign: 'middle', fontFace: BODY, fontSize: 8.25, bold: true, color: btnFill ? C.gray : C.white });
  });
}

/** Linear blend of two "RRGGBB" colours. */
function mix(a, b, t) {
  const ch = (i) => Math.round(parseInt(a.substr(i, 2), 16) * (1 - t) + parseInt(b.substr(i, 2), 16) * t);
  return [ch(0), ch(2), ch(4)].map((v) => v.toString(16).padStart(2, '0').toUpperCase()).join('');
}

/** Vertical colour ramp built from solid slices (pptxgenjs has no gradients). */
function vRamp(s, x, y, w, h, stops, steps) {
  for (let i = 0; i < steps; i++) {
    const t = (i / (steps - 1)) * (stops.length - 1);
    const k = Math.min(Math.floor(t), stops.length - 2);
    s.addShape('rect', {
      x, y: y + (h * i) / steps, w, h: h / steps + 0.012,
      fill: col(mix(stops[k], stops[k + 1], t - k)), line: NOLINE,
    });
  }
}

// Right edge of slide 15's rounded backdrop, sampled top (y=0) to bottom (y=5.625).
const BLOB_EDGE = [4.402, 4.790, 5.099, 5.350, 5.557, 5.726, 5.862, 5.969, 6.048, 6.102,
  6.132, 6.138, 6.121, 6.080, 6.014, 5.922, 5.801, 5.650, 5.463, 5.235, 4.960];

/**
 * Slide 15's backdrop: a rounded slab whose fill fades to white toward the
 * top-right. Painted as solid 45-degree bands, then masked to the outline.
 */
function fadingBlob(s, base, bands) {
  const K = Math.SQRT1_2;
  const project = (x, y) => x * K - y * K;   // distance along the fade axis
  const from = project(0, 5.625), to = project(6.139, 0);
  const span = project(10, 0) - from;        // covers the whole slide
  const step = span / bands;
  for (let i = 0; i < bands; i++) {
    const p = from + (i + 0.5) * step;
    const t = Math.max(0, ((p - from) / (to - from) - 0.23) / 0.77);
    s.addShape('rect', {
      x: p * K - 8.5, y: -p * K - step / 2, w: 17, h: step + 0.06, rotate: 45,
      fill: col(mix(base, C.white, Math.min(1, t))), line: NOLINE,
    });
  }
  const rowH = 5.625 / (BLOB_EDGE.length - 1);
  s.addShape('custGeom', {
    x: 0, y: 0, w: 10, h: 5.625, fill: col(C.white), line: NOLINE,
    points: BLOB_EDGE.map((edge, r) => ({ x: edge, y: r * rowH }))
      .concat([{ x: 10, y: 5.625 }, { x: 10, y: 0 }, { close: true }]),
  });
}

/* =========================================================== slide 15/30 */
function slide15() {
  const s = pptx.addSlide();
  fadingBlob(s, C.yellow, 70);
  confetti(s, { A: C.light, G: C.light, B: C.blue, C: C.light, D: C.light, L: C.light, I: C.ink + '@20', J: C.light, E: C.light }, [
    [[['M', 0.606, 3.851, 0.317, 0.445, 21.3], ['D', 0.984, 3.956, 0.117, 0.117, 21.3], ['D', 0.509, 4.316, 0.117, 0.117, 21.3]], C.white + '@40'],
    [[['S', 3.195, 1.063, 0.319, 0.436, 0]], C.pink],
    [[['S', 0.509, 1.024, 0.319, 0.436, 0]], C.white + '@40'],
  ]);
  chrome(s, { label: C.white, num: '15', numColor: C.white });

  card(s, 4.428, 1.163, 4.651, 3.333);
  s.addText([{ text: 'Good Concept', options: { breakLine: true } }, { text: 'Great Production' }], {
    x: 0.769, y: 2.097, w: 3.568, h: 1.01, valign: 'top', fontFace: HEAD, fontSize: 27, bold: true, color: C.white });
  txt(s, LOREM_SHORT, { x: 0.769, y: 3.178, w: 2.637, h: 0.518, align: 'justify', color: C.white });

  const longer = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivam us vel euismod leo. Donec comm odo et urna ac semper. ';
  txt(s, longer + 'Mauris finibus augue id vulputate sit amet, consectetur', { x: 4.765, y: 1.711, w: 2.077, h: 1.351, align: 'justify' });
  txt(s, longer, { x: 4.752, y: 3.064, w: 2.077, h: 0.934, align: 'justify' });
  txt(s, 'Wedding Concepts', { x: 6.944, y: 1.659, w: 1.926, h: 0.328, align: 'justify', bold: true, fontSize: 9 });
  txt(s, LOREM_SHORT, { x: 6.944, y: 1.993, w: 1.789, h: 0.726, align: 'justify' });
  txt(s, 'Wedding Productions', { x: 6.991, y: 2.987, w: 1.926, h: 0.328, align: 'justify', bold: true, fontSize: 9, color: C.blue });
  txt(s, LOREM_SHORT, { x: 6.991, y: 3.321, w: 1.789, h: 0.726, align: 'justify' });
}

/* =========================================================== slide 16/30 */
function slide16() {
  const s = pptx.addSlide();
  confetti(s, { A: C.light, B: C.blue, F: C.light, C: C.light, K: C.yellow, G: C.blue, D: C.light, L: C.light, H: C.light, I: C.light, E: C.light });
  chrome(s, { num: '16', ovalFill: C.yellow, ovalLine: C.silver });

  heading(s, 'Our Newest Decoration And Theming Design', { x: 0.849, y: 1.29, w: 3.468, h: 1.464 });
  s.addShape('line', { x: 0, y: 2.924, w: 4.017, h: 0, line: { color: C.silver, width: 1 } });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivam us vel euismod leo. Donec comm odo et urna ac semper. Mauris finibus augue id', { x: 0.849, y: 3.113, w: 3.097, h: 0.726, align: 'justify' });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivam us vel euismod leo. Donec', { x: 0.849, y: 3.865, w: 3.097, h: 0.518, align: 'justify' });

  const priceDot = (cx, cy, fill, price, priceX, priceY, title, titleX, titleY, textX, textY) => {
    s.addShape('ellipse', { x: cx, y: cy, w: 0.978, h: 0.978, fill: col(fill), line: NOLINE });
    txt(s, price, { x: priceX, y: priceY, w: 0.936, h: 0.48, align: 'center', bold: true, fontSize: 15, color: C.white });
    txt(s, title, { x: titleX, y: titleY, w: 1.926, h: 0.328, align: 'center', bold: true, fontSize: 9, color: fill });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur adipis cing', { x: textX, y: textY, w: 1.789, h: 0.518, align: 'center' });
  };
  priceDot(6.422, 0.915, C.pink, '$ 398', 6.443, 1.153, 'Wedding Decoration', 7.512, 1.049, 7.58, 1.383);
  priceDot(6.767, 3.341, C.yellow, '$ 100', 6.77, 3.608, 'Wedding Theme ', 4.919, 3.367, 4.987, 3.701);
}

/* =========================================================== slide 17/30 */
function slide17() {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 3.323, w: 10, h: 2.302, fill: col(C.pink), line: NOLINE });
  card(s, 6.381, 2.162, 2.273, 1.599, { rotate: 11.6, flipH: true, shadow: SHADOW_LG() });
  card(s, 1.367, 2.136, 2.273, 1.599, { rotate: -11.6, shadow: SHADOW_LG() });
  card(s, 6.204, 3.375, 2.273, 1.599, { rotate: -15, flipH: true, shadow: SHADOW_LG() });
  card(s, 1.709, 3.375, 2.273, 1.599, { rotate: 15, shadow: SHADOW_LG() });
  card(s, 3.457, 2.31, 3.097, 2.563, { shadow: SHADOW_LG() });
  confetti(s, { J: C.light, E: C.light, A: C.light, B: C.pink, F: C.ink + '@20', C: C.light, K: C.light, G: C.ink + '@20', L: C.yellow, H: C.light, I: C.light, D: C.blue },
    [[[['S', 6.804, 2.609, 0.298, 0.407, 12.5]], C.smoke]]);
  chrome(s, { num: '17', numColor: C.white, foot: C.white, tel: C.white, dash: C.white, ovalLine: C.white });

  s.addText('Our Wedding Gallery', { x: 2.91, y: 0.515, w: 4.181, h: 0.555, align: 'center', valign: 'top', fontFace: HEAD, fontSize: 27, bold: true, color: C.ink });
  txt(s, LOREM_CONSEC, { x: 1.032, y: 1.163, w: 7.935, h: 0.518, align: 'center' });
}

/* =========================================================== slide 18/30 */
function slide18() {
  const s = pptx.addSlide();
  card(s, 0, 3.674, 10, 1.951, { fill: C.light });
  confetti(s, { F: C.ink + '@20', B: C.pink, C: C.pink, K: C.ink + '@20', G: C.ink + '@20', D: C.ink + '@20', L: C.ink + '@20', H: C.blue, I: C.light, J: C.light, E: C.ink + '@20' },
    [[[['S', 8.977, 1.823, 0.236, 0.323, 195]], C.yellow]]);
  chrome(s, { num: '18' });

  card(s, 2.123, 1.369, 5.754, 2.888, { fill: C.yellow, shadow: SHADOW_LG() });
  s.addText('BREAK SLIDE PRESENTATION', { x: 2.552, y: 1.848, w: 4.896, h: 1.313, align: 'center', valign: 'top', fontFace: HEAD, fontSize: 36, bold: true, color: C.white });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing', { x: 2.858, y: 3.06, w: 4.285, h: 0.309, align: 'center', color: C.white });
  card(s, 2.123, 3.491, 5.754, 0.49, { fill: C.blue, shadow: SHADOW_LG() });
  s.addText('TAKE A SAVERAL MINUTES FOR BREAK SECTION', { x: 2.123, y: 3.491, w: 5.754, h: 0.49, align: 'center', valign: 'middle', fontFace: BODY, fontSize: 10.13, color: C.white });
}

/* =========================================================== slide 19/30 */
function slide19() {
  const s = pptx.addSlide();
  s.background = { color: C.yellow };
  confetti(s, { A: C.ink + '@20', B: C.pink, F: C.ink + '@20', C: C.ink + '@20', K: C.ink + '@20', G: C.ink + '@20', D: C.blue, L: C.blue, H: C.ink + '@20', I: C.ink + '@20', J: C.ink + '@20', E: C.ink + '@20' });
  chrome(s, { label: C.white, num: '19', numColor: C.white, foot: C.white, tel: C.white, dash: C.white, ovalFill: C.yellow, ovalLine: C.white, badge: C.white, badgeDot: C.yellow });

  card(s, 6.548, 1.282, 2.537, 3.399, { rotate: 21.5, shadow: SHADOW_LG() });
  card(s, 4.839, 1.234, 2.273, 3.046, { rotate: -10, shadow: SHADOW_LG() });

  s.addText([{ text: 'About Private ', options: { breakLine: true } }, { text: 'Parties Project' }], {
    x: 0.74, y: 1.195, w: 3.728, h: 1.01, valign: 'top', fontFace: HEAD, fontSize: 27, bold: true, color: C.white });
  statCard(s, -0.032, 2.506, 4.241, '+110', C.pink, 'Private Parties Projects', 4.084, 0.740, 1.680);
  txt(s, LOREM + ' sit amet' + '\u2026\u2026\u2026. ', { x: 0.74, y: 3.733, w: 3.728, h: 0.726, align: 'justify', color: C.white });
  txt(s, 'Read More >', { x: 2.988, y: 4.518, w: 1.307, h: 0.309, align: 'right', bold: true, color: C.white });
}

/* =========================================================== slide 20/30 */
function slide20() {
  const s = pptx.addSlide();
  confetti(s, { A: C.yellow, B: C.white + '@20', F: C.white + '@20', C: C.pink, K: C.light, G: C.light, D: C.light, L: C.white + '@10', H: C.white + '@10', I: C.white + '@10', J: C.light, E: C.light });
  chrome(s, { num: '20', numColor: C.white });

  heading(s, 'We Work Around  Your Schedule', { x: 4.151, y: 0.994, w: 3.737, h: 1.01 });

  // month chip + agenda text, four in a staggered grid
  const agenda = [
    ['AUG,', C.white, C.gray, 3.064, 2.362, 3.132, 3.218, 'First Agenda', C.gray, 4.185, 2.357, 4.185, 2.606],
    ['SEPT,', C.pink, C.white, 6.152, 2.362, 6.221, 6.306, 'Second Agenda', C.pink, 7.229, 2.357, 7.229, 2.606],
    ['NOV,', C.blue, C.white, 3.755, 3.766, 3.823, 3.908, 'Third Agenda', C.blue, 4.843, 3.762, 4.843, 4.011],
    ['DEC,', C.yellow, C.white, 6.789, 3.766, 6.857, 6.943, 'Fourth Agenda', C.yellow, 7.887, 3.762, 7.887, 4.011],
  ];
  agenda.forEach(([m, chip, ink, bx, by, mx, dx, title, tcol, tx, ty, px, py]) => {
    card(s, bx, by, 0.985, 0.902, { fill: chip, shadow: SHADOW_LG() });
    txt(s, m, { x: mx, y: by + 0.029, w: 0.848, h: 0.555, align: 'center', bold: true, fontSize: 18, color: ink, wrap: false });
    txt(s, '28 2021', { x: dx, y: by + 0.438, w: 0.677, h: 0.328, align: 'center', bold: true, fontSize: 9, color: ink });
    txt(s, title, { x: tx, y: ty, w: 1.926, h: 0.328, align: 'justify', bold: true, fontSize: 9, color: tcol });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur adip is cing elit. Vivamus', { x: px, y: py, w: 1.494, h: 0.726, align: 'justify' });
  });
}

/* =========================================================== slide 21/30 */
function slide21() {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: -0.03, w: 5.177, h: 5.655, fill: col(C.blue), line: NOLINE });
  card(s, 4.354, 0.87, 4.299, 3.351);
  confetti(s, { A: C.pink, B: C.ink + '@20', F: C.ink + '@20', C: C.ink + '@20', K: C.yellow, G: C.ink + '@20', D: C.pink, L: C.ink + '@20', H: C.ink + '@20', I: C.ink + '@20', J: C.ink + '@20', E: C.ink + '@20' });
  chrome(s, { label: C.white, num: '21', numColor: C.white });

  s.addText([{ text: 'About Vanue', options: { breakLine: true } }, { text: 'Sourcing ' }], {
    x: 0.806, y: 1.404, w: 3.737, h: 1.01, valign: 'top', fontFace: HEAD, fontSize: 27, bold: true, color: C.white });
  txt(s, LOREM + ' s ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec', { x: 0.806, y: 2.541, w: 3.001, h: 1.142, align: 'justify', color: C.white });
  pill(s, 0.877, 3.905, 1.25, 0.316, 'More Info', { line: C.white, radius: 0.054, size: 9 });

  card(s, 7.079, 3.137, 2.418, 1.456, { fill: C.yellow });
  txt(s, '$ 502.00,-', { x: 7.274, y: 3.349, w: 1.926, h: 0.357, align: 'justify', bold: true, fontSize: 10.13, color: C.white });
  s.addShape('line', { x: 8.373, y: 3.526, w: 0.827, h: 0, line: { color: C.white, width: 1 } });
  txt(s, 'Lorem ipsum dolor sit amet, conse ctetur adip is cing elit. Viva mus Lorem ipsum dolor sit ', { x: 7.287, y: 3.687, w: 2.029, h: 0.726, align: 'justify', color: C.white });
}

/* =========================================================== slide 22/30 */
function slide22() {
  const s = pptx.addSlide();
  confetti(s, { B: C.pink, F: C.ink + '@20', E: C.ink + '@20' });
  chrome(s, { num: '22' });
  s.addShape('rect', { x: 0.946, y: 2.337, w: 6.781, h: 2.419, fill: col(C.white), line: NOLINE });
  s.addShape('rect', { x: 7.632, y: 2.337, w: 1.422, h: 2.419, fill: col(C.yellow), line: NOLINE });
  card(s, 7.02, 2.795, 1.558, 1.513, { fill: C.pink, shadow: SHADOW_LG() });

  s.addText([{ text: 'Best Deal For ', options: { breakLine: true } }, { text: 'Gourmet & Catering' }], {
    x: 1.447, y: 2.702, w: 4.769, h: 1.01, valign: 'top', fontFace: HEAD, fontSize: 27, bold: true, color: C.ink });
  txt(s, LOREM + ' sit amet, consectetur adipiscing elit. Vivamus  euismod leo. Donec commodo et', { x: 1.447, y: 3.753, w: 5.241, h: 0.726, align: 'justify' });
  txt(s, 'Discount Up to', { x: 6.839, y: 3.068, w: 1.92, h: 0.309, align: 'center', bold: true, color: C.white });
  txt(s, '15%', { x: 6.839, y: 3.218, w: 1.92, h: 0.783, align: 'center', bold: true, fontSize: 27, color: C.white });
  txt(s, 'For Your Special Day', { x: 6.839, y: 3.857, w: 1.92, h: 0.309, align: 'center', bold: true, color: C.white });

  confetti(s, { G: C.blue, D: C.light, I: C.light, J: C.blue });
}

/* =========================================================== slide 23/30 */
function slide23() {
  const s = pptx.addSlide();
  confetti(s, { B: C.light, C: C.pink, K: C.light, D: C.light, L: C.light, H: C.yellow }, [
    [[['S', 3.117, 3.044, 0.319, 0.436, 0, true]], C.light],
    [[['M', 1.787, 3.344, 0.317, 0.445, 21.3], ['D', 2.164, 3.448, 0.117, 0.117, 21.3], ['D', 1.689, 3.808, 0.117, 0.117, 21.3]], C.light],
    [[['S', 7.751, 4.269, 0.236, 0.323, 0]], C.light],
  ]);
  chrome(s, { num: '23' });

  s.addText('Private Parties Services', { x: 2.656, y: 0.547, w: 4.688, h: 0.555, align: 'center', valign: 'top', fontFace: HEAD, fontSize: 27, bold: true, color: C.ink });

  [['Theming', C.pink, 1.171, 1.485, 1.217, 1.882], ['Productions', C.blue, 3.712, 4.061, 3.768, 4.458], ['Entertaining ', C.yellow, 6.242, 6.557, 6.288, 6.953]]
    .forEach(([title, fill, cx, tx, bx, btnX]) => {
      card(s, cx, 2.998, 2.576, 1.91, { fill });
      s.addText(title, { x: tx, y: 3.214, w: 1.947, h: 0.271, align: 'center', valign: 'top', fontFace: BODY, fontSize: 10.13, bold: true, color: C.white });
      txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac', { x: bx, y: 3.59, w: 2.43, h: 0.726, align: 'center', color: C.white });
      pill(s, btnX, 4.342, 1.154, 0.323, 'More Detail', { line: C.white, radius: 0.1615, size: 7.88, bold: false });
    });
}

/* =========================================================== slide 24/30 */
function slide24() {
  const s = pptx.addSlide();
  s.background = { color: C.blue };
  s.addShape('rect', { x: 0, y: 0, w: 2.554, h: 5.625, fill: col(C.yellow), line: NOLINE });
  confetti(s, { H: C.ink + '@20', G: C.ink + '@20', D: C.ink + '@20', J: C.ink + '@20', E: C.ink + '@20', B: C.ink + '@20', I: C.ink + '@20', A: C.pink, F: C.pink, C: C.yellow, K: C.ink + '@20', L: C.ink + '@20' },
    [[[['D', 5.443, 2.648, 0.095, 0.074, 330], ['M', 5.260, 2.727, 0.239, 0.336, 330], ['D', 5.419, 3.108, 0.088, 0.088, 330]], C.ink + '@25']]);
  chrome(s, { label: C.white, num: '24', numColor: C.white, foot: C.white, tel: C.white, dash: C.white, ovalFill: C.blue, ovalLine: C.white, badge: C.white, badgeDot: C.blue });

  card(s, 0.76, 1.323, 4.093, 3.264, { shadow: SHADOW_LG() });
  s.addText('About Corporate Event Projects', { x: 5.155, y: 1.18, w: 3.394, h: 1.01, valign: 'top', fontFace: HEAD, fontSize: 27, bold: true, color: C.white });
  statCard(s, 4.798, 2.395, 4.068, '+200', C.yellow, 'Completed Corporate Event', 8.747, 5.192, 6.168);
  statCard(s, 5.734, 3.598, 4.266, '95%', C.pink, 'Client Satishfied', 5.648, 5.977, 6.857);
}

/* =========================================================== slide 25/30 */
function slide25() {
  const s = pptx.addSlide();
  confetti(s, { B: C.light, F: C.light, C: C.pink, K: C.pink, G: C.blue, D: C.light, L: C.yellow, H: C.light, I: C.light, J: C.light });
  chrome(s, { num: '25' });

  heading(s, 'Technical Productions', { x: 1.397, y: 1.392, w: 3.23, h: 1.01 });
  card(s, 1.052, 2.75, 5.469, 2.152, { fill: C.blue });
  card(s, 6.495, 2.75, 0.57, 2.152, { fill: C.yellow });
  s.addShape('line', { x: 3.786, y: 3.058, w: 0, h: 1.635, line: { color: C.smoke, width: 2.25, dashType: 'dash' } });

  icon(s, 'people', 2.164, 3.074, 0.48, C.white);
  txt(s, 'Productions Team ', { x: 1.442, y: 3.478, w: 1.926, h: 0.328, align: 'center', bold: true, fontSize: 9, color: C.white });
  const body = 'Lorem ipsum dolor sit amet, consec tetur adipis cing elit. Vivamus dolor sit amet, consec tetur adipis';
  txt(s, body, { x: 1.307, y: 3.812, w: 2.141, h: 0.726, align: 'center', color: C.white });

  s.addText('$150.00', { x: 3.998, y: 3.193, w: 2.272, h: 0.353, align: 'center', valign: 'top', fontFace: HEAD, fontSize: 15, bold: true, color: C.white });
  txt(s, 'Maximum Budget', { x: 4.171, y: 3.478, w: 1.926, h: 0.328, align: 'center', bold: true, fontSize: 9, color: C.white });
  txt(s, body, { x: 4.035, y: 3.812, w: 2.131, h: 0.726, align: 'center', color: C.white });
  txt(s, 'CORPORTE EVENT PROJECT', { x: 5.734, y: 3.665, w: 2.063, h: 0.328, align: 'center', fontSize: 9, color: C.white, rotate: 270 });
}

/* =========================================================== slide 26/30 */
function slide26() {
  const s = pptx.addSlide();
  confetti(s, { A: C.light, F: C.light, C: C.pink, D: C.light, L: C.light, H: C.light, E: C.light });
  chrome(s, { num: '26' });

  heading(s, 'View Event Corporate Gallery', { x: 5.488, y: 0.913, w: 3.668, h: 1.01 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit', { x: 5.488, y: 1.916, w: 3.668, h: 0.309, align: 'justify' });

  // agenda cards with a vertical date tab on the side
  const agendaCard = (cardX, cardY, tabX, tabY, tabFill, dateLabel, dateX, textX, textY, title) => {
    card(s, cardX, cardY, 4.068, 2.053, { shape: 'roundRect', rectRadius: 0 });
    s.addShape('roundRect', { x: tabX, y: tabY, w: 0.396, h: 1.328, rectRadius: rounded(0.396, 1.328), fill: col(tabFill), line: NOLINE });
    txt(s, dateLabel, { x: dateX, y: tabY + 0.5, w: 1.386, h: 0.328, align: 'center', bold: true, fontSize: 9, color: C.white, rotate: 270 });
    txt(s, title, { x: textX, y: textY, w: 1.386, h: 0.328, align: 'justify', bold: true, fontSize: 9 });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur adip sci ng elit. Vivamus vel euismod leo\u2026..', { x: textX, y: textY + 0.26, w: 1.386, h: 0.934, align: 'justify', italic: true });
    txt(s, 'Read More >', { x: textX + 0.597, y: textY + 1.186, w: 0.823, h: 0.518, align: 'right', bold: true, italic: true });
  };
  agendaCard(0.791, 0.656, 4.512, 1.015, C.yellow, '04 AUGUST', 3.999, 1.024, 0.981, 'Agenda 01');
  agendaCard(0.791, 2.897, 4.512, 3.282, C.pink, '20 UGUST', 4.008, 1.024, 3.218, 'Agenda 02');
  agendaCard(5.141, 2.897, 8.863, 3.282, C.blue, '20 SEPTEMBER', 8.351, 5.430, 3.218, 'Agenda 02');
}

/* =========================================================== slide 27/30 */
function slide27() {
  const s = pptx.addSlide();
  confetti(s, { A: C.pink, F: C.pink, C: C.light, D: C.light, L: C.light, H: C.light },
    [[[['S', 7.043, 2.442, 0.319, 0.436, 321.7]], C.light]]);
  chrome(s, { num: '27' });

  card(s, 1.368, 0.622, 7.075, 2.048, { shape: 'roundRect', rectRadius: 0 });
  card(s, 1.557, 2.912, 7.075, 2.048, { shape: 'roundRect', rectRadius: 0, flipH: true });

  s.addShape('roundRect', { x: 8.057, y: 0.807, w: 0.576, h: 1.652, rectRadius: rounded(0.576, 1.652), fill: col(C.yellow), line: NOLINE });
  txt(s, '27 SEPTEMBER', { x: 7.627, y: 1.259, w: 1.386, h: 0.707, align: 'center', bold: true, fontSize: 12, color: C.white, rotate: 270 });
  s.addShape('roundRect', { x: 1.368, y: 3.098, w: 0.576, h: 1.652, rectRadius: rounded(0.576, 1.652), fill: col(C.blue), line: NOLINE });
  txt(s, '30 SEPTEMBER', { x: 0.987, y: 3.55, w: 1.386, h: 0.707, align: 'center', bold: true, fontSize: 12, color: C.white, rotate: 90 });

  [[1.670, 0.990, 3.820, 0.899, '+571 ', 'Agenda 04', 1.724, 1.416, 1.487],
   [5.053, 3.250, 7.202, 3.159, '+1K ', 'Agenda 05', 5.107, 3.676, 3.747]]
    .forEach(([tx, ty, cx, cy, count, title, lx, ly, by]) => {
      txt(s, title, { x: tx, y: ty, w: 1.386, h: 0.328, align: 'justify', bold: true, fontSize: 9 });
      txt(s, [{ text: count, options: { bold: true, fontSize: 10.13 } }, { text: 'Peoples ', options: { fontSize: 10.5 } }],
        { x: cx, y: cy, w: 1.386, h: 0.366, align: 'justify' });
      s.addShape('line', { x: lx, y: ly, w: 3.147, h: 0, line: { color: C.smoke, width: 1 } });
      txt(s, LOREM + ' sit amet, consectetur adipiscing elit. Vivamus  euismod', { x: tx, y: by, w: 3.201, h: 0.934, align: 'justify' });
    });
}

/* =========================================================== slide 28/30 */
function slide28() {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: 2.784, h: 5.625, fill: col(C.pink), line: NOLINE });
  confetti(s, { A: C.light, F: C.blue, C: C.yellow, G: C.light, D: C.light, L: C.ink + '@20', H: C.ink + '@20', J: C.light, E: C.light });
  chrome(s, { label: C.white, num: '28', numColor: C.white });

  phoneMockup(s, 1.496, 0.346, 2.514, 4.933);

  heading(s, 'Our Social Media Mockup Section', { x: 4.69, y: 0.99, w: 3.668, h: 1.01 });
  [2.255, 2.993, 3.731].forEach((y) => {
    txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. ', {
      x: 4.666, y, w: 4.129, h: 0.518, align: 'justify', bullet: { characterCode: '2713', indent: 10.15 },
    });
  });
  badge(s, 4.779, 4.475, 1.627, 0.441, C.yellow, C.white);
}

/** Stand-in for the phone photo: rounded body, screen and speaker notch. */
function phoneMockup(s, x, y, w, h) {
  s.addShape('roundRect', { x, y, w, h, rectRadius: 0.28, fill: col('1A1A1A'), line: NOLINE, shadow: SHADOW_LG() });
  vRamp(s, x + 0.11, y + 0.11, w - 0.22, h - 0.22,
    ['E7286E', 'E63B7C', 'E44A78', 'E85C63', 'F0975E', 'F7DCC8', 'EEE6EA'], 40);
  s.addShape('roundRect', { x: x + w / 2 - 0.32, y: y + 0.09, w: 0.64, h: 0.16, rectRadius: 0.08, fill: col('1A1A1A'), line: NOLINE });
}

/* =========================================================== slide 29/30 */
function slide29() {
  const s = pptx.addSlide();
  confetti(s, { A: C.pink, B: C.blue, C: C.yellow, G: C.light, D: C.light, L: C.light, J: C.light, E: C.light },
    [[[['S', 5.772, 4.331, 0.319, 0.436, 0, true]], C.light]]);
  chrome(s, { num: '29', dy: 0.013 });

  heading(s, 'Get In Touch ', { x: 5.827, y: 1.387, w: 3.668, h: 0.555 });
  txt(s, LOREM, { x: 5.837, y: 1.955, w: 3.244, h: 0.726, align: 'justify' });

  card(s, 2.106, 3.287, 2.215, 1.23, { shape: 'roundRect', rectRadius: 0 });
  card(s, 4.475, 3.274, 2.267, 1.23, { shape: 'roundRect', rectRadius: 0, fill: C.pink });
  card(s, 6.892, 3.274, 2.128, 1.23, { shape: 'roundRect', rectRadius: 0 });

  txt(s, 'Offices Location', { x: 2.357, y: 3.541, w: 1.209, h: 0.309, align: 'center', bold: true, fill: col(C.light) });
  txt(s, 'Velodrome Streets, The West Office 28 District A-148', { x: 2.321, y: 3.801, w: 1.788, h: 0.518, align: 'justify' });

  txt(s, 'Our Contact', { x: 4.688, y: 3.5, w: 0.899, h: 0.309, align: 'center', bold: true, color: C.pink, fill: col(C.white) });
  txt(s, '+52 1921 1823 12 (Contact Peson)', { x: 4.632, y: 3.849, w: 2.11, h: 0.24, color: C.white, lineSpacingMultiple: 1 });
  txt(s, '0001 293 124 1(Fax)', { x: 4.632, y: 4.031, w: 2.11, h: 0.24, color: C.white, lineSpacingMultiple: 1 });

  txt(s, 'Email Address', { x: 7.145, y: 3.478, w: 1.022, h: 0.309, align: 'center', bold: true, fill: col(C.light) });
  txt(s, 'Event.planner@domain.com', { x: 7.145, y: 3.788, w: 1.813, h: 0.309 });
  txt(s, 'planner.evnt@domain.com', { x: 7.145, y: 4.023, w: 1.813, h: 0.309 });

  // contact pictograms: map pin, phone handset, envelope
  s.addShape('custGeom', {
    x: 2.83, y: 3.42, w: 0.22, h: 0.30, fill: col(C.gray), line: NOLINE,
    points: outline([['M', 0.5, 1.0], ['C', 0.5, 1.0, 0.0, 0.5, 0.0, 0.32],
      ['C', 0.0, 0.14, 0.22, 0.0, 0.5, 0.0], ['C', 0.78, 0.0, 1.0, 0.14, 1.0, 0.32],
      ['C', 1.0, 0.5, 0.5, 1.0, 0.5, 1.0], ['Z']], 0.22, 0.30),
  });
  s.addShape('ellipse', { x: 2.885, y: 3.485, w: 0.11, h: 0.11, fill: col(C.white), line: NOLINE });
  s.addShape('custGeom', {
    x: 6.328, y: 3.455, w: 0.252, h: 0.25, fill: col(C.white), line: NOLINE,
    points: outline([['M', 0.579, 0.580], ['C', 0.500, 0.659, 0.399, 0.738, 0.360, 0.700],
      ['C', 0.300, 0.639, 0.259, 0.598, 0.140, 0.700], ['C', 0.000, 0.799, 0.099, 0.878, 0.160, 0.919],
      ['C', 0.218, 0.998, 0.459, 0.939, 0.698, 0.700], ['C', 0.937, 0.460, 0.998, 0.219, 0.937, 0.138],
      ['C', 0.878, 0.079, 0.818, 0.000, 0.718, 0.120], ['C', 0.619, 0.239, 0.660, 0.278, 0.718, 0.341],
      ['C', 0.759, 0.377, 0.680, 0.479, 0.579, 0.580], ['Z']], 0.252, 0.25),
  });
  s.addShape('custGeom', {
    x: 8.639, y: 3.504, w: 0.244, h: 0.151, fill: col(C.gray), line: NOLINE,
    points: outline([['M', 0.039, 0.095], ['L', 0.500, 0.467], ['L', 0.961, 0.095], ['L', 0.961, 0.000],
      ['L', 0.039, 0.000], ['Z'], ['M', 0.020, 0.281], ['L', 0.500, 0.625], ['L', 0.980, 0.281],
      ['L', 0.980, 0.996], ['L', 0.020, 0.996], ['Z']], 0.244, 0.151),
  });
}

/* =========================================================== slide 30/30 */
function slide30() {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: 10, h: 4.812, fill: col(C.yellow), line: NOLINE });
  card(s, 0.917, 1.115, 4.365, 4.51, { shadow: SHADOW_LG() });
  confetti(s, { I: C.light }, [
    [[['M', 3.732, 0.374, 0.302, 0.424, 21.3], ['D', 4.092, 0.473, 0.112, 0.112, 21.3], ['D', 3.639, 0.816, 0.112, 0.112, 21.3]], C.ink + '@20'],
    [[['S', 8.786, 1.012, 0.236, 0.323, 195]], C.ink + '@20'],
    [[['D', 6.553, 0.172, 0.066, 0.066, 0], ['M', 6.362, 0.167, 0.239, 0.336, 0], ['D', 6.381, 0.556, 0.088, 0.088, 0]], C.pink],
    [[['S', 8.231, 4.326, 0.236, 0.323, 0]], C.blue],
    [[['M', 0.462, 1.146, 0.179, 0.252, 345], ['D', 0.529, 1.442, 0.066, 0.066, 345]], C.white],
    [[['S', 0.335, 3.584, 0.319, 0.436, 321.7]], C.ink + '@20'],
    [[['M', 5.991, 4.044, 0.179, 0.252, 0], ['D', 6.182, 4.048, 0.066, 0.066, 0], ['D', 6.005, 4.336, 0.066, 0.066, 0]], C.light],
  ]);
  chrome(s, { label: C.white, dy: 0.013 });

  txt(s, 'End Of Presentation', { x: 5.765, y: 1.504, w: 3.244, h: 0.309, align: 'justify', color: C.white });
  s.addText('THANK YOU FOR WATCHING', { x: 5.727, y: 1.804, w: 3.832, h: 1.111, valign: 'top', fontFace: HEAD, fontSize: 30, bold: true, color: C.white });
  txt(s, LOREM, { x: 5.787, y: 2.894, w: 3.244, h: 0.726, align: 'justify', color: C.white });
}

/* ----------------------------------------------------------------- build */

[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
 slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
 slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30]
  .forEach((build) => build());

const outFile = path.join(__dirname, '0bc31e45-447b-4daf-9231-7121e16d121a_grok_final.pptx');
pptx.writeFile({ fileName: outFile }).then(() => console.log('wrote ' + outFile));
