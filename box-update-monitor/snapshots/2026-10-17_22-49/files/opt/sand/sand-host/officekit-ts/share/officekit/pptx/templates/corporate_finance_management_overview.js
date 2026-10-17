/**
 * "Corporate Financial" deck - rebuilt with pptxgenjs.
 * Raster photos in the source deck are replaced by flat "[image]" placeholder boxes.
 *
 *   node 13c4e6b0-5899-41ae-be06-55db357138eb_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const C = {
  purple: 'A06CFF',
  yellow: 'FFFF6D',
  mint: '28F8C4',
  cyan: '4DE3EB',
  magenta: 'A02B93',
  black: '000000',
  ink: '0D0D0D', // near-black panels / slide-1 background
  night: '111827', // heading ink
  white: 'FFFFFF',
  grey: '595959', // body copy on white
  greyMid: '808080',
  greyDark: '404040',
  greyLine: 'BFBFBF',
  greyLt: 'D9D9D9', // body copy on black
  greyXl: 'F2F2F2', // body copy on near-black
  photo: 'CECECE', // flat stand-in for the grey stock photos
  photoWarm: '8E6654', // flat stand-in for the hand-holding-phone photo
  shadeYellow: 'BFBF52',
  shadePurple: '7851BF',
};

const HEAD = 'Questrial'; // theme major font
const BODY = 'Open Sans'; // theme minor font

const KAPPA = 0.5523; // circle -> cubic bezier control offset

const EYEBROW = 'Understanding Financial Management in Corporations';
const EYEBROW_CAPS = 'UNDERSTANDING FINANCIAL MANAGEMENT IN CORPORATIONS';
const LOREM_LONG =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
  'Fusce posuere magna sed pulvinar ultricies, purus lectus malesuada';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue';
const LOREM_STEP = 'Lorem dolor sit amet consectetur adipiscing elit sed do';
const LOREM_INFO = 'Lorem ipsum dolor sit amet, elit adipiscing elit consectetuer';
const LOREM_CARD = 'PLACEHOLDER';

/* ---------------------------------------------------------------- text tools */

// Every text box in the source deck is top-anchored with PowerPoint's default insets.
function text(slide, content, opts) {
  slide.addText(content, Object.assign({ fontFace: BODY, valign: 'top', color: C.black }, opts));
}

function eyebrow(slide, x, y, opts) {
  text(slide, EYEBROW, Object.assign({ x, y, w: 4.483, h: 0.286, fontSize: 11, color: C.purple }, opts));
}

function heading(slide, content, x, y, w, opts) {
  text(slide, content, Object.assign({ x, y, w, h: 0.707, fontSize: 36, bold: true, color: C.night, fontFace: HEAD }, opts));
}

function bodyCopy(slide, content, x, y, w, h, opts) {
  text(slide, content, Object.assign({ x, y, w, h, fontSize: 12, color: C.grey, lineSpacingMultiple: 1.5 }, opts));
}

/* ------------------------------------------------------- custom geometry kit */

// pptxgenjs custGeom points are inches measured from the shape's own top-left.
// Path templates below are unit-square (0..1) and get scaled by the shape box.
function scalePath(tmpl, w, h) {
  return tmpl.map(function (seg) {
    if (seg[0] === 'Z') return { close: true };
    if (seg[0] === 'M') return { x: seg[1] * w, y: seg[2] * h, moveTo: true };
    if (seg[0] === 'L') return { x: seg[1] * w, y: seg[2] * h };
    return {
      x: seg[5] * w,
      y: seg[6] * h,
      curve: { type: 'cubic', x1: seg[1] * w, y1: seg[2] * h, x2: seg[3] * w, y2: seg[4] * h },
    };
  });
}

function addPath(slide, tmpl, o) {
  slide.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    points: scalePath(tmpl, o.w, o.h),
    fill: o.fill ? { color: o.fill } : { type: 'none' },
    line: o.line,
    flipH: o.flipH, flipV: o.flipV, rotate: o.rotate,
  });
}

// Quarter disc; `corner` marks the square corner (the arc's centre).
const QUARTER = {
  tr: [['M', 0, 0], ['C', 0, KAPPA, 1 - KAPPA, 1, 1, 1], ['L', 1, 0], ['Z']],
  bl: [['M', 1, 1], ['C', 1, 1 - KAPPA, 1 - KAPPA, 0, 0, 0], ['L', 0, 1], ['Z']],
  tl: [['M', 1, 0], ['C', 1, KAPPA, KAPPA, 1, 0, 1], ['L', 0, 0], ['Z']],
  br: [['M', 0, 1], ['C', 0, 1 - KAPPA, KAPPA, 0, 1, 0], ['L', 1, 1], ['Z']],
};

// Half-ring arch, opening downwards (2 : 1 box).
const ARCH = [
  ['M', 0.5, 1],
  ['C', 0.776, 0.999, 0.999, 0.552, 1, 0],
  ['L', 0.739, 0],
  ['C', 0.739, 0.264, 0.632, 0.478, 0.5, 0.478],
  ['C', 0.368, 0.478, 0.261, 0.264, 0.261, 0],
  ['L', 0, 0],
  ['C', 0, 0.553, 0.224, 1, 0.5, 1],
  ['Z'],
];

// Square with a quarter-disc bite; `corner` marks the bitten corner.
const K2 = 0.5 * KAPPA;
const NOTCH = {
  tl: [['M', 0.5, 0], ['L', 1, 0], ['L', 1, 1], ['L', 0, 1], ['L', 0, 0.5],
    ['C', K2, 0.5, 0.5, K2, 0.5, 0], ['Z']],
  tr: [['M', 0.5, 0], ['L', 0, 0], ['L', 0, 1], ['L', 1, 1], ['L', 1, 0.5],
    ['C', 1 - K2, 0.5, 0.5, K2, 0.5, 0], ['Z']],
  bl: [['M', 0.5, 1], ['L', 1, 1], ['L', 1, 0], ['L', 0, 0], ['L', 0, 0.5],
    ['C', K2, 0.5, 0.5, 1 - K2, 0.5, 1], ['Z']],
  br: [['M', 0.5, 1], ['L', 0, 1], ['L', 0, 0], ['L', 1, 0], ['L', 1, 0.5],
    ['C', 1 - K2, 0.5, 0.5, 1 - K2, 0.5, 1], ['Z']],
};

// Quarter ring (band width 0.43); `corner` marks the ring's centre.
const RI = 0.571; // inner radius as a fraction of the box
const RING = {
  bl: [['M', 0, 0], ['C', 0.552, 0, 1, 0.448, 1, 1], ['L', RI, 1],
    ['C', RI, 1 - RI * 0.552, RI * 0.552, 1 - RI, 0, 1 - RI], ['L', 0, 0], ['Z']],
  br: [['M', 1, 0], ['C', 0.448, 0, 0, 0.448, 0, 1], ['L', 1 - RI, 1],
    ['C', 1 - RI, 1 - RI * 0.552, 1 - RI * 0.552, 1 - RI, 1, 1 - RI], ['L', 1, 0], ['Z']],
  tl: [['M', 0, 1], ['C', 0.552, 1, 1, 0.552, 1, 0], ['L', RI, 0],
    ['C', RI, RI * 0.552, RI * 0.552, RI, 0, RI], ['L', 0, 1], ['Z']],
  tr: [['M', 1, 1], ['C', 0.448, 1, 0, 0.552, 0, 0], ['L', 1 - RI, 0],
    ['C', 1 - RI, RI * 0.552, 1 - RI * 0.552, RI, 1, RI], ['L', 1, 1], ['Z']],
};

// Four-lobe "clover" motif with two star-shaped voids.
const BLOB = [
  ['M', 0.448, 0.337], ['L', 0.423, 0.345], ['C', 0.37, 0.367, 0.333, 0.419, 0.333, 0.48],
  ['L', 0.333, 0.487], ['L', 0.33, 0.457], ['C', 0.319, 0.4, 0.274, 0.355, 0.217, 0.344],
  ['L', 0.189, 0.341], ['L', 0.217, 0.338], ['C', 0.283, 0.325, 0.333, 0.266, 0.333, 0.195],
  ['L', 0.336, 0.224], ['C', 0.346, 0.272, 0.379, 0.311, 0.423, 0.329], ['Z'],
  ['M', 1, 0.854], ['L', 1, 0.479], ['C', 1, 0.399, 0.935, 0.333, 0.854, 0.333],
  ['L', 0.667, 0.333], ['L', 0.854, 0.333], ['C', 0.935, 0.333, 1, 0.268, 1, 0.187],
  ['L', 1, 0], ['L', 0.813, 0], ['C', 0.732, 0, 0.667, 0.065, 0.667, 0.146],
  ['L', 0.667, 0.154], ['L', 0.664, 0.124], ['C', 0.65, 0.058, 0.591, 0.007, 0.521, 0.007],
  ['L', 0.333, 0.007], ['L', 0.146, 0.008], ['C', 0.065, 0.008, 0, 0.073, 0, 0.154],
  ['L', 0, 0.528], ['C', 0, 0.589, 0.037, 0.64, 0.089, 0.663], ['L', 0.114, 0.67],
  ['L', 0.089, 0.678], ['C', 0.037, 0.7, 0, 0.752, 0, 0.813], ['L', 0, 1],
  ['L', 0.187, 1], ['C', 0.268, 1, 0.333, 0.935, 0.333, 0.854], ['L', 0.333, 0.853],
  ['L', 0.336, 0.883], ['C', 0.35, 0.949, 0.409, 1, 0.479, 1], ['L', 0.667, 1],
  ['L', 0.667, 0.812], ['C', 0.667, 0.742, 0.617, 0.683, 0.55, 0.669], ['L', 0.523, 0.666],
  ['L', 0.546, 0.664], ['C', 0.605, 0.654, 0.652, 0.608, 0.664, 0.55], ['L', 0.667, 0.521],
  ['L', 0.67, 0.55], ['C', 0.682, 0.608, 0.728, 0.654, 0.787, 0.664], ['L', 0.813, 0.667],
  ['L', 0.783, 0.67], ['C', 0.717, 0.683, 0.667, 0.742, 0.667, 0.813], ['L', 0.667, 1],
  ['L', 0.854, 1], ['C', 0.935, 1, 1, 0.935, 1, 0.854], ['Z'],
];

// Dashed elbow leader line used on the fan chart.
const ELBOW = [['M', 1, 1], ['L', 0, 1], ['L', 0, 0]];

/* -------------------------------------------------------------- decorations */

// 3 x 3 tile logo: quarter discs, a solid square and an arch.
// Cells are [col, row, kind, colorKey]; `kind` is a QUARTER corner, 'rect' or 'arch'.
const LOGO_TILES = [
  [0, 0, 'bl', 'purple'],
  [1, 0, 'br', 'white'],
  [2, 1, 'bl', 'purple'],
  [1, 1, 'rect', 'yellow'],
  [0, 1, 'tr', 'mint'],
  [0, 2, 'br', 'purple'],
  [1, 2, 'arch', 'white'],
];
const MIRROR_CORNER = { bl: 'br', br: 'bl', tl: 'tr', tr: 'tl' };

function logoCluster(slide, ox, oy, u, opt) {
  const o = opt || {};
  LOGO_TILES.forEach(function (t) {
    let col = t[0];
    let kind = t[2];
    if (o.noArch && kind === 'arch') return;
    const color = C[o.colors && o.colors[t[3]] ? o.colors[t[3]] : t[3]];
    const wide = kind === 'arch' ? 2 : 1;
    if (o.mirror) {
      col = 3 - wide - col;
      kind = MIRROR_CORNER[kind] || kind;
    }
    const box = { x: ox + col * u, y: oy + t[1] * u, w: wide * u, h: u, fill: color };
    if (kind === 'rect') slide.addShape('rect', { x: box.x, y: box.y, w: box.w, h: box.h, fill: { color: color } });
    else if (kind === 'arch') addPath(slide, ARCH, box);
    else addPath(slide, QUARTER[kind], box);
  });
}

// Clover motif surrounded by a notched square, quarter discs, a ring and a bar.
// Entries are [x, y, w, h, kind, colorKey] as fractions of the cluster unit,
// listed back to front. Portrait on slides 3/5/6, turned 90 deg on slide 8.
const BLOB_CLUSTER = [
  [0.2871, -0.0667, 0.6451, 0.779, 'blob', 'dark', { rotate: 270 }],
  [0.5192, 0.6504, 0.4803, 0.4467, 'notch:tl', 'purple'],
  [0.5213, 0.6457, 0.2299, 0.2163, 'quarter:tl', 'mint'],
  [0.6598, 1.0997, 0.3396, 0.3396, 'quarter:tr', 'dark'],
  [-0.001, 0.5764, 0.5223, 0.5202, 'ring:tr', 'yellow'],
  [0.0016, 0.0047, 0.2283, 0.5712, 'rect', 'purple'],
];
const BLOB_CLUSTER_TURNED = [
  [0.7937, 0, 0.6451, 0.779, 'blob', 'dark', { flipV: true }],
  [0.3249, 0.0168, 0.4803, 0.4467, 'notch:br', 'purple'],
  [0.5706, 0.2557, 0.2299, 0.2163, 'quarter:br', 'mint'],
  [0, 0, 0.3396, 0.3396, 'quarter:tr', 'dark'],
  [0.3412, 0.4793, 0.5223, 0.5202, 'ring:tr', 'yellow'],
  [0.8635, 0.779, 0.58, 0.2189, 'rect', 'purple'],
];

function blobCluster(slide, ox, oy, s, opt) {
  const o = opt || {};
  (o.tiles || BLOB_CLUSTER).forEach(function (t) {
    const kind = t[4].split(':');
    const color = t[5] === 'dark' ? o.dark : C[t[5]];
    const box = Object.assign({ x: ox + t[0] * s, y: oy + t[1] * s, w: t[2] * s, h: t[3] * s, fill: color }, t[6]);
    if (kind[0] === 'rect') slide.addShape('rect', { x: box.x, y: box.y, w: box.w, h: box.h, fill: { color: color } });
    else if (kind[0] === 'blob') addPath(slide, BLOB, box);
    else if (kind[0] === 'ring') addPath(slide, RING[kind[1]], box);
    else if (kind[0] === 'notch') addPath(slide, NOTCH[kind[1]], box);
    else addPath(slide, QUARTER[kind[1]], box);
  });
}

/* --------------------------------------------------------- image stand-ins */

function imageBox(slide, x, y, w, h, opt) {
  const o = opt || {};
  slide.addShape(o.shape || 'rect', {
    x: x, y: y, w: w, h: h,
    fill: { color: o.color || C.photo },
    rectRadius: o.radius,
    flipH: o.flipH, flipV: o.flipV,
  });
  if (o.label === false) return;
  text(slide, '[image]', {
    x: x, y: y + h / 2 - 0.18, w: w, h: 0.36,
    align: 'center', valign: 'middle', fontSize: 12, color: C.white,
  });
}

/* --------------------------------------------------------------- components */

// White card (soft drop shadow) with a coloured chip, caption and paragraph.
// pptxgenjs rewrites the shadow object in place, so hand each shape its own.
function cardShadow() {
  return { type: 'outer', color: '7F7F7F', opacity: 0.11, blur: 27, offset: 0.01, angle: 90 };
}

function featureCard(slide, o) {
  slide.addShape('rect', { x: o.x, y: o.y, w: o.w, h: 1.595, fill: { color: C.white }, shadow: cardShadow() });
  slide.addShape('rect', { x: o.x, y: o.y + 0.292, w: 0.369, h: 0.418, fill: { color: o.chip } });
  text(slide, o.title, {
    x: o.x + 0.569, y: o.y + 0.325, w: 3.31, h: 0.337,
    fontSize: 14, bold: true, fontFace: HEAD,
  });
  bodyCopy(slide, o.body, o.x + 0.569, o.y + 0.628, o.bodyW, 0.675);
}

// Small pictogram approximations for the icon glyphs of the source deck.
// Three overlapping bank cards, fanned out; the two front cards carry
// a magnetic stripe and a chip drawn in the surrounding background colour.
function cardsIcon(slide, x, y, w, h, color, bg) {
  slide.addShape('roundRect', { x: x + 0.32 * w, y: y + 0.04 * h, w: 0.62 * w, h: 0.17 * h, fill: { color: color }, rectRadius: 0.012, rotate: 12 });
  slide.addShape('roundRect', { x: x + 0.3 * w, y: y + 0.26 * h, w: 0.58 * w, h: 0.42 * h, fill: { color: color }, rectRadius: 0.02, rotate: 12 });
  slide.addShape('rect', { x: x + 0.67 * w, y: y + 0.36 * h, w: 0.13 * w, h: 0.14 * h, fill: { color: bg }, rotate: 12 });
  slide.addShape('roundRect', { x: x + 0.02 * w, y: y + 0.4 * h, w: 0.66 * w, h: 0.46 * h, fill: { color: bg }, rectRadius: 0.02, rotate: 12 });
  slide.addShape('roundRect', { x: x + 0.04 * w, y: y + 0.42 * h, w: 0.62 * w, h: 0.42 * h, fill: { color: color }, rectRadius: 0.02, rotate: 12 });
  slide.addShape('rect', { x: x + 0.07 * w, y: y + 0.51 * h, w: 0.57 * w, h: 0.11 * h, fill: { color: bg }, rotate: 12 });
  slide.addShape('rect', { x: x + 0.12 * w, y: y + 0.68 * h, w: 0.12 * w, h: 0.12 * h, fill: { color: bg }, rotate: 12 });
}

// Bar chart: an outlined bar, a solid bar and a hatched bar.
function chartIcon(slide, x, y, w, h, color, bg) {
  slide.addShape('roundRect', { x: x, y: y + 0.28 * h, w: 0.26 * w, h: 0.72 * h, fill: { color: color }, rectRadius: 0.012 });
  slide.addShape('rect', { x: x + 0.06 * w, y: y + 0.36 * h, w: 0.14 * w, h: 0.56 * h, fill: { color: bg } });
  slide.addShape('roundRect', { x: x + 0.34 * w, y: y + 0.05 * h, w: 0.28 * w, h: 0.95 * h, fill: { color: color }, rectRadius: 0.012 });
  slide.addShape('roundRect', { x: x + 0.7 * w, y: y + 0.42 * h, w: 0.26 * w, h: 0.58 * h, fill: { color: color }, rectRadius: 0.012 });
  slide.addShape('rect', {
    x: x + 0.75 * w, y: y + 0.49 * h, w: 0.16 * w, h: 0.44 * h,
    fill: { color: bg, type: 'solid' }, line: { color: bg, width: 0 },
  });
  for (let i = 0; i < 4; i++) {
    slide.addShape('line', {
      x: x + 0.75 * w, y: y + (0.55 + i * 0.11) * h, w: 0.16 * w, h: 0.1 * h,
      line: { color: color, width: 1 }, flipV: true,
    });
  }
}

// Head-and-shoulders with a small check badge.
function personIcon(slide, x, y, w, h, color, bg) {
  slide.addShape('ellipse', { x: x + 0.27 * w, y: y, w: 0.36 * w, h: 0.38 * h, fill: { color: color } });
  slide.addShape('arc', { x: x + 0.02 * w, y: y + 0.42 * h, w: 0.86 * w, h: 1.1 * h, fill: { color: color }, angleRange: [180, 360] });
  slide.addShape('triangle', { x: x + 0.38 * w, y: y + 0.42 * h, w: 0.14 * w, h: 0.5 * h, fill: { color: bg }, flipV: true });
  slide.addShape('ellipse', { x: x + 0.66 * w, y: y + 0.6 * h, w: 0.34 * w, h: 0.36 * h, fill: { color: color }, line: { color: bg, width: 1 } });
}

// Floppy-disk "save" glyph with a clipped top-right corner.
function saveIcon(slide, x, y, w, h, color, bg) {
  slide.addShape('snip1Rect', { x: x, y: y, w: w, h: h, fill: { color: color }, rotate: 270, flipV: true });
  slide.addShape('rect', { x: x + 0.14 * w, y: y + 0.1 * h, w: 0.52 * w, h: 0.24 * h, fill: { color: bg } });
  slide.addShape('ellipse', { x: x + 0.31 * w, y: y + 0.48 * h, w: 0.34 * w, h: 0.34 * h, fill: { color: bg } });
}

// Map pin holding a cog.
function pinIcon(slide, x, y, w, h, color, bg) {
  slide.addShape('teardrop', { x: x + 0.03 * w, y: y, w: 0.94 * w, h: 0.75 * h, fill: { color: color }, rotate: 135 });
  slide.addShape('gear9', { x: x + 0.12 * w, y: y + 0.08 * h, w: 0.76 * w, h: 0.6 * h, fill: { color: bg } });
  slide.addShape('ellipse', { x: x + 0.32 * w, y: y + 0.23 * h, w: 0.36 * w, h: 0.28 * h, fill: { color: color } });
}

function docIcon(slide, x, y, w, h, color) {
  slide.addShape('rect', { x: x, y: y, w: 0.72 * w, h: h, fill: { color: color } });
  slide.addShape('rect', { x: x + 0.12 * w, y: y + 0.12 * h, w: 0.48 * w, h: 0.76 * h, fill: { color: C.white } });
  [0.24, 0.42, 0.6].forEach(function (fy) {
    slide.addShape('rect', { x: x + 0.18 * w, y: y + fy * h, w: 0.36 * w, h: 0.08 * h, fill: { color: color } });
  });
  slide.addShape('rect', { x: x + 0.5 * w, y: y + 0.18 * h, w: 0.18 * w, h: 0.6 * h, fill: { color: color }, rotate: 28 });
}

function speakerIcon(slide, x, y, w, h, color) {
  slide.addShape('rect', { x: x, y: y + 0.33 * h, w: 0.18 * w, h: 0.34 * h, fill: { color: color } });
  slide.addShape('triangle', { x: x + 0.02 * w, y: y + 0.06 * h, w: 0.88 * h, h: 0.42 * w, fill: { color: color }, rotate: 90 });
  slide.addShape('arc', { x: x + 0.4 * w, y: y + 0.24 * h, w: 0.34 * w, h: 0.52 * h, line: { color: color, width: 2 }, angleRange: [300, 60] });
  slide.addShape('arc', { x: x + 0.56 * w, y: y + 0.06 * h, w: 0.44 * w, h: 0.88 * h, line: { color: color, width: 2 }, angleRange: [300, 60] });
}

/* -------------------------------------------------------------- the slides */

// 1 - cover
function slide01(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.ink };
  [
    [0, -0.002, 0.728, 'bl', C.purple], [0.733, -0.002, 0.728, 'br', C.white],
    [0, 0.722, 0.728, 'tr', C.mint], [0, 1.45, 0.728, 'br', C.yellow],
    [0.001, 4.44, 0.695, 'tl', C.mint], [12.709, 0, 0.623, 'br', C.yellow],
    [12.68, 5.524, 0.653, 'tl', C.purple], [12.68, 6.177, 0.653, 'bl', C.mint],
    [12.023, 6.826, 0.653, 'tl', C.white], [12.68, 6.826, 0.652, 'tr', C.purple],
  ].forEach(function (q) {
    addPath(s, QUARTER[q[3]], { x: q[0], y: q[1], w: q[2], h: q[2], fill: q[4] });
  });
  text(s, EYEBROW_CAPS, {
    x: 3.196, y: 1.998, w: 6.941, h: 0.269,
    align: 'center', fontSize: 10, charSpacing: 3, color: C.white,
  });
  text(s, [
    { text: 'Corporate ', options: { breakLine: true } },
    { text: 'Financial', options: { color: C.yellow } },
  ], {
    x: 1.838, y: 2.47, w: 9.658, h: 3.205,
    align: 'center', fontSize: 115, bold: true, color: C.white,
    fontFace: HEAD, lineSpacingMultiple: 0.8,
  });
}

// 2 - introduction
function slide02(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 10.208, y: 0, w: 3.125, h: 7.499, fill: { color: C.ink } });
  imageBox(s, 6.161, 0, 5.027, 7.499, { shape: 'round1Rect', radius: 1.03, flipV: true });
  logoCluster(s, 11.403, 0, 0.643, { mirror: true });

  heading(s, 'Introduction to Corporate Finance', 0.705, 1.332, 5.962, { h: 1.447, fontSize: 40 });
  eyebrow(s, 0.705, 0.953);
  text(s, 'Objectives Of Corporate Finance', { x: 0.705, y: 3.079, w: 3.31, h: 0.337, fontSize: 14, bold: true, fontFace: HEAD });
  bodyCopy(s, LOREM_LONG, 0.705, 3.507, 4.483, 0.978);

  featureCard(s, { x: 0.705, y: 4.952, w: 4.247, bodyW: 3.476, chip: C.purple, title: 'Management Finance', body: LOREM_SHORT });

  text(s, 'FINANCE', {
    x: 10.908, y: 4.165, w: 3.755, h: 1.01,
    fontSize: 54, bold: true, color: '1A1A1A', fontFace: HEAD, rotate: 90,
  });
}

// 3 - key concepts
function slide03(pptx) {
  const s = pptx.addSlide();
  imageBox(s, 0, 0, 3.644, 7.499);
  s.addShape('rect', { x: 3.644, y: 5.272, w: 9.689, h: 2.249, fill: { color: C.ink } });
  blobCluster(s, 11.429, 0.007, 1.905, { dark: C.ink });

  heading(s, 'Key Concepts in Corporate Finance', 4.433, 1.258, 5.4, { h: 1.447, fontSize: 40 });
  eyebrow(s, 4.433, 0.88);
  text(s, 'Cost of Capital', { x: 4.433, y: 3.027, w: 3.396, h: 0.337, fontSize: 14, bold: true, fontFace: HEAD });
  bodyCopy(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
    'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.',
  4.433, 3.452, 7.588, 0.978);

  [['Capital Structure', 4.433, 1.912], ['Financial Management', 8.708, 2.701]].forEach(function (t, i) {
    const dy = i * 0.068;
    text(s, t[0], { x: t[1], y: 5.789 + dy, w: t[2], h: 0.337, fontSize: 14, bold: true, color: C.yellow, fontFace: HEAD });
    bodyCopy(s, LOREM_SHORT, t[1], 6.217 + dy, 3.836, 0.675, { color: C.greyLt });
  });
}

// 4 - sources of finance
function slide04(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: 5.667, h: 7.499, fill: { color: C.ink } });
  s.addShape('rect', { x: 5.667, y: 5.25, w: 3.833, h: 2.249, fill: { color: C.purple } });
  imageBox(s, 9.5, 5.25, 3.833, 2.249);
  logoCluster(s, 0, 0, 0.835);

  heading(s, 'Sources of Finance', 6.589, 1.073, 5.4);
  eyebrow(s, 6.589, 0.694);
  bodyCopy(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
    'Fusce posuere magna sed pulvinar.', 6.589, 1.926, 5.997, 0.675);

  s.addShape('rect', { x: 6.667, y: 2.953, w: 5.997, h: 1.595, fill: { color: C.white }, shadow: cardShadow() });
  s.addShape('rect', { x: 6.667, y: 3.245, w: 0.369, h: 0.418, fill: { color: C.mint } });
  [
    ['Equity Financing', 'Retained Earnings'],
    ['Debt Financing', 'Personal Savings'],
    ['Hybrid Instruments', 'Government Grants'],
  ].forEach(function (row, i) {
    row.forEach(function (label, col) {
      text(s, label, {
        x: col ? 10.007 : 7.236, y: 3.124 + i * 0.458, w: 2.201, h: 0.326,
        fontSize: 12, lineSpacingMultiple: 1.2, paraSpaceAfter: 6,
        bullet: { characterCode: '006F', indent: 22.5 },
      });
    });
  });

  text(s, [
    { text: '50% ', options: { color: C.yellow } },
    { text: 'Explain Issuing Stocks And Retaining Earnings' },
  ], { x: 0.512, y: 3.327, w: 4.644, h: 1.919, fontSize: 36, bold: true, color: C.white, fontFace: HEAD });
  bodyCopy(s, 'Lorem ipsum dolor sit amet, c onsectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
    'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada', 0.512, 5.499, 4.644, 0.975, { color: C.greyXl });

  text(s, 'Capital Structure', { x: 5.824, y: 5.823, w: 2.565, h: 0.37, fontSize: 16, bold: true, color: C.yellow, fontFace: HEAD });
  bodyCopy(s, LOREM_SHORT, 5.824, 6.251, 3.519, 0.675, { color: C.white });
}

// 5 - investment decisions
function slide05(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: 13.333, h: 3.583, fill: { color: C.ink } });
  imageBox(s, 0.595, 4.074, 5.966, 2.802, { shape: 'round2DiagRect' });
  imageBox(s, 6.745, 4.074, 2.901, 2.802, { shape: 'round2DiagRect', flipH: true });
  imageBox(s, 9.83, 4.074, 2.901, 2.802, { shape: 'round2DiagRect' });
  blobCluster(s, 11.7735, 0.0101, 1.5607, { dark: C.white });

  heading(s, 'Investment Decisions', 0.748, 1.059, 5.4, { color: C.white });
  eyebrow(s, 0.748, 0.68);
  bodyCopy(s, 'Lorem ipsum dolor sit amet, c onsectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
    'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada', 0.748, 1.928, 5.106, 0.975, { color: C.greyXl });

  text(s, '60%', { x: 6.675, y: 0.681, w: 2.095, h: 1.212, fontSize: 66, bold: true, color: C.yellow, wrap: false });
  text(s, 'Efficiency Finance', { x: 8.841, y: 0.852, w: 0.943, h: 0.505, fontSize: 12, color: C.greyXl, lineSpacingMultiple: 1 });
  text(s, 'Capital Budgeting', { x: 6.675, y: 1.766, w: 1.561, h: 0.303, fontSize: 12, color: C.white, fontFace: HEAD, wrap: false });
  bodyCopy(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
    'Fusce posuere, magna sed pulvinar ultricies.', 6.677, 2.228, 5.939, 0.675, { color: C.greyXl });
}

// 6 - financial analysis
function slide06(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 8.104, y: 0, w: 5.229, h: 7.5, fill: { color: C.ink } });
  imageBox(s, 8.104, 0, 5.229, 5.0, { shape: 'round1Rect', radius: 1.02, flipV: true });
  cardsIcon(s, 8.776, 5.381, 0.513, 0.457, C.white, C.ink);
  text(s, 'Corporate Finance', { x: 8.635, y: 6.007, w: 2.558, h: 0.438, fontSize: 20, color: C.white, fontFace: HEAD, wrap: false });
  bodyCopy(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ',
    8.635, 6.444, 4.167, 0.675, { color: C.greyXl });

  heading(s, 'Financial Analysis', 0.904, 1.058, 4.483);
  eyebrow(s, 0.904, 0.679);
  bodyCopy(s, LOREM_LONG, 0.904, 1.904, 5.917, 0.978);
  featureCard(s, { x: 0.904, y: 3.331, w: 5.917, bodyW: 5.022, chip: C.purple, title: 'Financial Statements',
    body: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed' });
  featureCard(s, { x: 0.904, y: 5.226, w: 5.917, bodyW: 5.022, chip: C.mint, title: 'Ratio Analysis',
    body: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed' });

  blobCluster(s, 6.5865, 0.0151, 1.5607, { dark: C.ink });
}

// 7 - portfolio grid
function slide07(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0.602, y: 0.601, w: 2.939, h: 6.316, fill: { color: C.ink } });
  imageBox(s, 3.708, 2.773, 7.0, 1.996);
  imageBox(s, 10.875, 2.773, 1.856, 4.143);
  imageBox(s, 3.708, 4.893, 2.939, 1.996);
  imageBox(s, 6.814, 4.92, 3.894, 1.996);
  logoCluster(s, 0.602, 0.583, 0.617);

  heading(s, 'Corporate Finance Portfolio', 4.021, 1.296, 6.687);
  eyebrow(s, 4.021, 0.917);
  text(s, 'Capital Structure', { x: 0.872, y: 4.433, w: 1.912, h: 0.337, fontSize: 14, bold: true, color: C.yellow, fontFace: HEAD });
  bodyCopy(s, LOREM_SHORT, 0.872, 4.861, 2.4, 0.978, { color: C.greyLt });

  s.addShape('roundRect', { x: 0.945, y: 6.114, w: 1.094, h: 0.298, fill: { type: 'none' }, line: { color: C.white, width: 0.75 }, rectRadius: 0.149 });
  text(s, 'Learn More', { x: 0.977, y: 6.12, w: 1.03, h: 0.286, align: 'center', fontSize: 11, color: C.white, fontFace: HEAD });
}

// 8 - team
function slide08(pptx) {
  const s = pptx.addSlide();
  const cards = [
    [0.602, C.ink, C.white, 'Michael Johnson'],
    [4.7, C.purple, C.white, 'Charlotte Nelson'],
    [8.797, C.yellow, C.black, 'Amelia Mitchell'],
  ];
  cards.forEach(function (c) {
    imageBox(s, c[0], 3.357, 3.934, 3.464);
    s.addShape('rect', { x: c[0], y: 6.18, w: 3.934, h: 0.642, fill: { color: c[1] } });
    text(s, c[3], {
      x: c[0], y: 6.282, w: 3.934, h: 0.438,
      align: 'center', fontSize: 20, color: c[2], fontFace: HEAD,
    });
  });

  heading(s, 'Meet Our Expert Team', 0.602, 1.225, 5.35);
  eyebrow(s, 0.602, 0.846);
  bodyCopy(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
    'Fusce posuere magna sed pulvinar ultricies, ', 0.602, 2.071, 5.917, 0.675);

  blobCluster(s, 10.592, -0.005, 1.905, { dark: C.ink, tiles: BLOB_CLUSTER_TURNED });
}

// 9 - mockup
function slide09(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: 3.562, h: 7.5, fill: { color: C.ink } });
  // Stand-in for the "hand holding a phone" photo: the hand, the phone
  // body, and the grey screen placeholder.
  s.addShape('roundRect', { x: -0.6, y: 3.1, w: 6.6, h: 3.8, fill: { color: C.photoWarm }, rectRadius: 0.9, rotate: 300 });
  s.addShape('roundRect', { x: 2.35, y: 1.05, w: 2.75, h: 5.65, fill: { color: '1C1C1C' }, rectRadius: 0.3 });
  imageBox(s, 2.646, 1.854, 2.135, 3.812);
  logoCluster(s, 0, 0.011, 0.617);

  heading(s, 'Corporate Financial Mockup Design', 6.564, 1.422, 5.35, { h: 1.313 });
  eyebrow(s, 6.564, 1.043);
  bodyCopy(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
    'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo', 6.564, 2.86, 5.917, 0.978);

  cardsIcon(s, 6.705, 4.34, 0.513, 0.457, C.black, C.white);
  text(s, 'Corporate Finance', { x: 6.564, y: 4.966, w: 2.558, h: 0.438, fontSize: 20, fontFace: HEAD, wrap: false });
  bodyCopy(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed',
    6.564, 5.404, 5.746, 0.675);
}

// Shared heading block of the four "Infographic Section" slides.
function infographicTitle(slide) {
  text(slide, 'Infographic Section', {
    x: 3.992, y: 1.359, w: 5.35, h: 0.707,
    align: 'center', fontSize: 36, bold: true, color: C.night, fontFace: HEAD,
  });
  eyebrow(slide, 4.425, 0.98, { align: 'center' });
}

// 10 - pinwheel infographic
function slide10(pptx) {
  const s = pptx.addSlide();
  infographicTitle(s);

  [['bl', C.black, 4.88, 4.191], ['br', C.cyan, 6.183, 4.191],
    ['tr', C.purple, 6.187, 2.884], ['tl', C.yellow, 4.88, 2.888]].forEach(function (r) {
    addPath(s, RING[r[0]], { x: r[2], y: r[3], w: 2.271, h: 2.263, fill: r[1] });
  });

  // [circle x, circle fill, label x, label align, label colour]
  const nodes = [
    [4.618, 2.632, C.yellow, 1.763, 3.083, 'right', C.black],
    [4.618, 5.57, C.black, 1.763, 3.083, 'right', C.black],
    [7.584, 2.632, C.purple, 8.982, 8.982, 'left', C.purple],
    [7.584, 5.57, C.cyan, 8.982, 8.982, 'left', C.cyan],
  ];
  nodes.forEach(function (n) {
    s.addShape('ellipse', { x: n[0], y: n[1], w: 1.131, h: 1.131, fill: { color: n[2] } });
    text(s, 'Your Data', {
      x: n[4], y: n[1] - 0.149, w: 1.268, h: 0.337,
      align: n[5], fontSize: 14, color: n[6],
    });
    bodyCopy(s, LOREM_INFO, n[3], n[1] + 0.373, 2.589, 0.978, { align: n[5], color: C.greyDark });
  });

  chartIcon(s, 4.988, 3.008, 0.391, 0.379, C.black, C.yellow);
  personIcon(s, 7.934, 2.971, 0.431, 0.454, C.white, C.purple);
  saveIcon(s, 5.02, 5.972, 0.328, 0.328, C.white, C.black);
  pinIcon(s, 7.967, 5.905, 0.366, 0.461, C.white, C.cyan);
}

// 11 - pricing table
function slide11(pptx) {
  const s = pptx.addSlide();
  infographicTitle(s);

  const bullets = [
    ['Mauris quam dolor', 'cursus at porta et, luctus Nunc tempor luctus libero accumsan eu, viverra et '],
    ['Mauris quam dolor', 'cursus at porta et, luctus Nunc tempor luctus libero accumsan eu, viverra et ',
      'Nunc rhoncus tellus in '],
    ['Mauris quam dolor', 'cursus at porta et, luctus Nunc tempor luctus libero',
      'Nunc rhoncus tellus in ', 'ipsum molestie et gravida '],
    ['Mauris quam dolor', 'cursus at porta et, luctus Nunc tempor luctus libero',
      'Nunc rhoncus tellus in ', 'ipsum molestie et gravida '],
  ];
  const columns = [
    [1.253, 2.778, C.black, C.white, '$3000', 'BASIC'],
    [3.996, 2.773, C.purple, C.white, '$6000', 'MODERATE'],
    [6.742, 2.773, C.mint, C.white, '$8000', 'PRO'],
    [9.478, 2.774, C.yellow, C.black, '$10000', 'EXTREME'],
  ];

  columns.forEach(function (col, i) {
    const x = col[0];
    const y = col[1];
    s.addShape('rect', { x: x, y: y, w: 2.602, h: 3.765, fill: { type: 'none' }, line: { color: col[2], width: 1.5 } });
    s.addShape('rect', { x: x + 0.234, y: y - 0.217, w: 2.134, h: 0.811, fill: { color: col[2] } });
    text(s, col[4], {
      x: x + 0.234, y: y - 0.217, w: 2.134, h: 0.811,
      align: 'center', valign: 'middle', margin: 0, fontSize: 24, bold: true, color: col[3], lineSpacingMultiple: 0.9,
    });
    s.addShape('rect', { x: x + 0.596, y: y + 3.571, w: 1.41, h: 0.378, fill: { color: col[2] } });
    text(s, col[5], {
      x: x + 0.596, y: y + 3.571, w: 1.41, h: 0.378,
      align: 'center', valign: 'middle', margin: 0, fontSize: 14, bold: true, color: col[3], lineSpacingMultiple: 0.9,
    });
    const dot = { characterCode: '2022', indent: 13.5 };
    text(s, bullets[i].map(function (line, k) {
      return { text: line, options: { bullet: dot, breakLine: k < bullets[i].length - 1 } };
    }), {
      x: x + 0.233, y: 3.76, w: 2.137, h: 2.156,
      fontSize: 12, color: C.grey, lineSpacingMultiple: 1.5,
      margin: [4.8, 4.8, 2.4, 2.4],
    });
  });
}

// 12 - five-step fan
function slide12(pptx) {
  const s = pptx.addSlide();
  infographicTitle(s);

  // Fan segments, drawn back to front.
  const outline = { color: C.white, width: 1 };
  addPath(s, [['M', 0.811, 0], ['L', 0, 1], ['L', 1, 1], ['C', 1, 0.647, 0.937, 0.293, 0.811, 0], ['Z']],
    { x: 7.206, y: 5.182, w: 2.253, h: 1.337, fill: C.purple, line: outline });
  addPath(s, [['M', 0.383, 0], ['L', 0, 1], ['L', 1, 0.383], ['C', 0.846, 0.199, 0.626, 0.068, 0.383, 0], ['Z']],
    { x: 7.206, y: 4.162, w: 2.014, h: 2.372, fill: C.magenta, line: outline });
  addPath(s, [['M', 0.503, 0], ['C', 0.328, 0, 0.159, 0.016, 0, 0.049], ['L', 0.503, 1], ['L', 1, 0.049],
    ['C', 0.841, 0.016, 0.677, 0, 0.503, 0], ['Z']],
  { x: 6.375, y: 3.814, w: 1.677, h: 2.72, fill: C.cyan, line: outline });
  addPath(s, [['M', 0, 0.384], ['L', 1, 1], ['L', 0.616, 0], ['C', 0.358, 0.073, 0.146, 0.21, 0, 0.384], ['Z']],
    { x: 4.84, y: 3.71, w: 2.38, h: 2.809, fill: C.yellow, line: outline });
  addPath(s, [['M', 0, 1], ['L', 1, 1], ['L', 0.191, 0], ['C', 0.07, 0.278, 0, 0.627, 0, 1], ['Z']],
    { x: 4.063, y: 4.655, w: 3.157, h: 1.864, fill: C.black, line: outline });

  // Dashed leader lines.
  const leader = { color: C.greyLine, width: 1, dashType: 'dash' };
  s.addShape('line', { x: 3.829, y: 5.745, w: 0.319, h: 0, line: leader, flipH: true });
  s.addShape('line', { x: 9.299, y: 5.745, w: 0.625, h: 0, line: leader });
  addPath(s, ELBOW, { x: 4.371, y: 4.001, w: 0.938, h: 0.294, line: leader, flipH: true, flipV: true });
  addPath(s, ELBOW, { x: 6.229, y: 2.945, w: 0.938, h: 0.883, line: leader, flipH: true, flipV: true });
  addPath(s, ELBOW, { x: 8.61, y: 4.001, w: 0.938, h: 0.491, line: leader, flipV: true });

  // Step numbers sitting on the segments.
  [['03', 6.825, 4.471, 0.712, C.white], ['02', 5.793, 4.739, 0.714, C.black],
    ['04', 7.848, 4.739, 0.729, C.white], ['01', 5.302, 5.698, 0.677, C.white],
    ['05', 8.357, 5.698, 0.729, C.white]].forEach(function (n) {
    text(s, [
      { text: n[0], options: { fontSize: 32, breakLine: true } },
      { text: 'STEP', options: { fontSize: 14 } },
    ], { x: n[1], y: n[2], w: n[3], h: 0.71, align: 'center', color: n[4], lineSpacingMultiple: 0.8, wrap: false });
  });

  // Callouts around the fan.
  [[2.36, 3.737, 1.729, 4.113, 'right'], [1.839, 5.219, 1.208, 5.595, 'right'],
    [4.188, 2.536, 3.556, 2.912, 'right'], [9.944, 3.695, 9.944, 4.071, 'left'],
    [10.365, 5.219, 10.365, 5.587, 'left']].forEach(function (t) {
    text(s, 'Add title here', { x: t[0], y: t[1], w: 1.55, h: 0.337, align: t[4], fontSize: 14, bold: true, wrap: false });
    text(s, LOREM_STEP, {
      x: t[2], y: t[3], w: 2.181, h: 0.719,
      align: t[4], fontSize: 11, color: C.greyMid, lineSpacingMultiple: 1.1,
    });
  });
}

// 13 - two pill cards
function slide13(pptx) {
  const s = pptx.addSlide();
  infographicTitle(s);

  // Pill body: rectangle + rounded cap + shaded lip + white disc.
  const capCurve = [['M', 1, 0], ['L', 0, 0], ['L', 0, 0.408],
    ['C', 0.096, 0.761, 0.284, 1, 0.499, 1], ['C', 0.716, 1, 0.903, 0.761, 1, 0.408], ['L', 1, 0], ['Z']];
  [[4.364, C.yellow, C.shadeYellow, 2.083], [6.591, C.purple, C.shadePurple, 2.081]].forEach(function (p) {
    s.addShape('rect', { x: p[0], y: 3.982, w: p[3], h: 2.538, fill: { color: p[1] } });
    addPath(s, capCurve, { x: p[0], y: 3.982, w: p[3], h: 1.036, fill: p[2] });
    s.addShape('ellipse', { x: p[0], y: 2.946, w: p[3], h: 2.072, fill: { color: p[1] } });
    s.addShape('ellipse', { x: p[0] + 0.15, y: 3.094, w: p[3] - 0.302, h: 1.775, fill: { color: C.white } });
  });
  docIcon(s, 5.234, 3.777, 0.343, 0.341, C.black);
  speakerIcon(s, 7.42, 3.775, 0.422, 0.334, C.purple);

  [[4.738, 4.468, C.black], [6.985, 6.715, C.white]].forEach(function (p) {
    text(s, 'Insert title', { x: p[0], y: 5.281, w: 1.294, h: 0.37, align: 'center', fontSize: 16, color: p[2], wrap: false });
    text(s, 'Sed ut perspiciat unde omni iste natu', {
      x: p[1], y: 5.668, w: 1.835, h: 0.604,
      align: 'center', fontSize: 12, color: p[2], lineSpacingMultiple: 1.3,
    });
  });

  // Left and right descriptions.
  text(s, '01', { x: 2.895, y: 3.247, w: 0.78, h: 0.707, align: 'right', fontSize: 36, bold: true, color: C.yellow, wrap: false });
  text(s, 'Insert title here', { x: 1.273, y: 3.991, w: 2.402, h: 0.438, align: 'right', fontSize: 20, bold: true, color: C.greyMid, wrap: false });
  bodyCopy(s, LOREM_CARD, 1.02, 4.537, 2.909, 0.978, { align: 'right', color: C.greyMid });
  text(s, '02', { x: 9.151, y: 3.366, w: 0.78, h: 0.707, fontSize: 36, bold: true, color: C.purple, wrap: false });
  text(s, 'Insert title here', { x: 9.151, y: 4.109, w: 2.402, h: 0.438, fontSize: 20, bold: true, color: C.greyMid, wrap: false });
  bodyCopy(s, LOREM_CARD, 9.151, 4.655, 2.909, 0.978, { color: C.greyMid });

  [[1.924, C.yellow, C.black], [9.242, C.purple, C.white]].forEach(function (b) {
    s.addShape('rect', { x: b[0], y: 5.885, w: 1.602, h: 0.509, fill: { color: b[1] } });
    text(s, 'Show All', { x: b[0], y: 5.885, w: 1.602, h: 0.509, align: 'center', valign: 'middle', fontSize: 14, color: b[2] });
  });
}

// 14 - contact
function slide14(pptx) {
  const s = pptx.addSlide();
  imageBox(s, 6.688, 0, 6.642, 7.5);
  s.addShape('rect', { x: 0.025, y: 3.953, w: 6.642, h: 3.547, fill: { color: C.ink } });

  heading(s, 'Contact Information\u2019s', 0.762, 1.19, 5.446, { color: C.black });
  eyebrow(s, 0.762, 0.811);
  bodyCopy(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
    'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet', 0.762, 2.165, 5.204, 0.978);

  const columns = [
    [0.744, 'Office Hours', ['Monday \u2013 Saturday', '08.00 AM \u2013 08.00 PM'],
      'Get In Touch', ['(+62) 000 0000 0000', '(0725) 00000']],
    [3.492, 'Address', ['123 Street Name, City Name'],
      'Follow Us', ['www.yoursitehere.com', 'office@ yoursitehere.com']],
  ];
  columns.forEach(function (col) {
    [[4.432, col[1], col[2]], [5.906, col[3], col[4]]].forEach(function (blk) {
      text(s, blk[1], { x: col[0], y: blk[0], w: 2.456, h: 0.467, fontSize: 16, color: C.yellow, lineSpacingMultiple: 1.5 });
      text(s, blk[2].map(function (line, i) {
        return { text: line, options: { breakLine: i < blk[2].length - 1 } };
      }), { x: col[0], y: blk[0] + 0.437, w: 2.456, h: 0.678, fontSize: 12, color: C.white, lineSpacingMultiple: 1.5 });
    });
  });
}

// 15 - thank you
function slide15(pptx) {
  const s = pptx.addSlide();
  logoCluster(s, 10.979, 0, 0.784, { mirror: true, noArch: true, colors: { white: 'black' } });
  addPath(s, QUARTER.bl, { x: 0, y: 3.602, w: 0.784, h: 0.784, fill: C.purple });

  text(s, EYEBROW_CAPS, { x: 0.615, y: 5.063, w: 6.941, h: 0.286, fontSize: 11, charSpacing: 3 });
  text(s, [
    { text: 'Thank ' },
    { text: 'You', options: { color: C.purple } },
  ], { x: 0.615, y: 5.216, w: 9.059, h: 2.036, fontSize: 115, bold: true, fontFace: HEAD });
}

/* --------------------------------------------------------------------- main */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK', width: 13.333, height: 7.5 });
  pptx.layout = 'DECK';
  pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
    slide09, slide10, slide11, slide12, slide13, slide14, slide15].forEach(function (fn) {
    fn(pptx);
  });

  const out = path.join(__dirname, '13c4e6b0-5899-41ae-be06-55db357138eb_grok_final.pptx');
  return pptx.writeFile({ fileName: out }).then(function () {
    console.log('wrote ' + out);
  });
}

build().catch(function (err) {
  console.error(err);
  process.exit(1);
});
