/**
 * Renogen Presentation — recreated with pptxgenjs.
 * Reference deck: 10 slides, 13.333in x 7.5in (16:9).
 *
 * The raster artwork of the original (quote marks, social glyphs, rating stars,
 * eco icons) is redrawn here as native vector shapes; no image data is embedded.
 *
 *   node 07d5493b-a6dd-4892-87f8-2142669afd2f_grok_final.js
 */

const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ================================================================== */
/* Design tokens                                                       */
/* ================================================================== */

const C = {
  accent: 'C3FF3D',   // theme accent1 - lime
  olive: '92BF2E',    // numerals on the philosophy slide
  ink: '262626',      // theme tx1, lightened 15%
  body: '808080',     // theme bg1, darkened 50%
  bodyDark: '404040', // body copy on light / lime cards
  light: 'F2F2F2',    // theme bg1, darkened 5%
  mid: 'D9D9D9',
  gray: 'A6A6A6',
  white: 'FFFFFF',
  orange: 'FFAE00',
  socialGray: 'BFBFBF',
};

// Both eco icons are placed with a luminance boost in the source deck, so the
// palette below is the base artwork colour already brightened.
const BARREL_C = {
  body: 'BDF05C', soft: 'AEE34A', leaf: '99D144', vein: '8EC240', cap: 'FFF452',
};
const BIN_C = {
  lid: 'BEFB4D', rim: 'A7E335', body: '94D631', shade: '88C42D',
  wheel: '7D7D7D', hub: 'FFFFFF',
};

const HEAD = 'Montserrat SemiBold';   // theme major latin font
const BODY = 'Open Sans';             // theme minor latin font

// Text-box insets used throughout the source deck (0.1in l/r, 0.05in t/b), in points.
const INSET = [7.2, 7.2, 3.6, 3.6];

// Soft drop shadow of the source deck. Returns a fresh object each call because
// pptxgenjs rewrites shadow properties in place while generating the XML.
const cardShadow = () => ({ type: 'outer', blur: 50, offset: 3, angle: 90, color: '000000', opacity: 0.1 });

/* ================================================================== */
/* Generic helpers                                                     */
/* ================================================================== */

/** Heading text: Montserrat SemiBold, dark ink, top-aligned. */
function heading(s, text, x, y, w, h, size, opts = {}) {
  s.addText(text, {
    x, y, w, h, fontFace: HEAD, fontSize: size, color: C.ink,
    align: 'left', valign: 'top', margin: INSET, ...opts,
  });
}

/** Small 14pt Montserrat caption. */
function caption(s, text, x, y, w, opts = {}) {
  heading(s, text, x, y, w, 0.337, 14, { bold: true, ...opts });
}

/** Body copy: 11pt Open Sans, 150% leading. */
function bodyText(s, text, x, y, w, h, opts = {}) {
  s.addText(text, {
    x, y, w, h, fontFace: BODY, fontSize: 11, color: C.body,
    lineSpacingMultiple: 1.5, align: 'left', valign: 'top', margin: INSET, ...opts,
  });
}

/** Rounded rectangle; `adj` is the OOXML corner-radius fraction of the short side. */
function roundRect(s, x, y, w, h, adj, opts = {}) {
  s.addShape('roundRect', { x, y, w, h, rectRadius: adj * Math.min(w, h), ...opts });
}

/**
 * Compact path syntax -> pptxgenjs custGeom points.
 * ['M', x, y] move, ['L', x, y] line, ['C', x1,y1, x2,y2, x,y] cubic, ['Z'] close.
 * All coordinates are fractions of the shape box.
 */
function pathPoints(cmds, w, h) {
  return cmds.map((c) => {
    if (c[0] === 'Z') return { close: true };
    if (c[0] === 'M') return { moveTo: true, x: c[1] * w, y: c[2] * h };
    if (c[0] === 'L') return { x: c[1] * w, y: c[2] * h };
    return {
      x: c[5] * w, y: c[6] * h,
      curve: { type: 'cubic', x1: c[1] * w, y1: c[2] * h, x2: c[3] * w, y2: c[4] * h },
    };
  });
}

function freeform(s, cmds, x, y, w, h, color, opts = {}) {
  s.addShape('custGeom', { x, y, w, h, points: pathPoints(cmds, w, h), fill: { color }, ...opts });
}

/** Closed polygon from a flat list of [x, y] fractions. */
function polygon(s, pts, x, y, w, h, color, opts = {}) {
  const cmds = pts.map((p, i) => [i === 0 ? 'M' : 'L', p[0], p[1]]);
  cmds.push(['Z']);
  freeform(s, cmds, x, y, w, h, color, opts);
}

/**
 * Draw a multi-part icon. Each part is [fx, fy, fw, fh, shape, color?] where the
 * f* values are fractions of the icon box and `shape` is either a path array or
 * one of the strings below.
 */
function icon(s, parts, x, y, size, defaultColor) {
  parts.forEach(([fx, fy, fw, fh, shape, color]) => {
    const box = [x + fx * size, y + fy * size, fw * size, fh * size];
    const fill = { color: color || defaultColor };
    if (shape === 'rect') s.addShape('rect', { x: box[0], y: box[1], w: box[2], h: box[3], fill });
    else if (shape === 'round') roundRect(s, ...box, 0.4, { fill });
    else if (shape === 'ellipse') s.addShape('ellipse', { x: box[0], y: box[1], w: box[2], h: box[3], fill });
    else if (shape === 'donut') s.addShape('donut', { x: box[0], y: box[1], w: box[2], h: box[3], fill });
    else if (shape === 'trapezoid') {
      s.addShape('trapezoid', { x: box[0], y: box[1], w: box[2], h: box[3], rotate: 180, fill });
    } else if (Array.isArray(shape[0])) freeform(s, shape, ...box, fill.color);
    else polygon(s, shape, ...box, fill.color);
  });
}

/* ================================================================== */
/* Vector artwork                                                      */
/* ================================================================== */

// Wire-frame cube brand mark: hexagon shell with four cut-out faces.
const LOGO_CUBE = [
  ['M', 0.5000, 0.0000], ['L', 1.0000, 0.2317], ['L', 1.0000, 0.7680],
  ['L', 0.5000, 1.0000], ['L', 0.0000, 0.7680], ['L', 0.0000, 0.2317], ['Z'],
  ['M', 0.5513, 0.5181], ['L', 0.5513, 0.8770], ['L', 0.8513, 0.7544], ['Z'],
  ['M', 0.4487, 0.5181], ['L', 0.1487, 0.7544], ['L', 0.4487, 0.8770], ['Z'],
  ['M', 0.8977, 0.3054], ['L', 0.5974, 0.4281], ['L', 0.8977, 0.6644], ['Z'],
  ['M', 0.1023, 0.3054], ['L', 0.1023, 0.6644], ['L', 0.4029, 0.4281], ['Z'],
  ['M', 0.5002, 0.1018], ['L', 0.1814, 0.2317], ['L', 0.5002, 0.3613], ['L', 0.8186, 0.2317], ['Z'],
];

// Rounded ">" glyph: nav tick, rating arrow and the two bullet arrows.
const CHEVRON = [
  ['M', 0.3017, 0.0306],
  ['C', 0.3464, 0.0612, 0.9441, 0.4252, 0.9441, 0.4252],
  ['C', 0.9441, 0.4558, 0.9944, 0.4830, 0.9944, 0.4830],
  ['C', 0.9944, 0.5136, 0.9441, 0.5442, 0.9441, 0.5748],
  ['C', 0.9441, 0.5748, 0.3464, 0.9354, 0.3017, 0.9660],
  ['C', 0.2458, 0.9966, 0.1508, 0.9966, 0.0503, 0.9660],
  ['C', 0.0000, 0.9354, 0.0000, 0.8741, 0.0503, 0.8163],
  ['L', 0.5978, 0.4830],
  ['C', 0.0503, 0.1497, 0.0503, 0.1497, 0.0503, 0.1497],
  ['C', 0.0000, 0.1224, 0.0000, 0.0612, 0.0503, 0.0306],
  ['C', 0.1508, 0.0000, 0.2458, 0.0000, 0.3017, 0.0306], ['Z'],
];

// Rounded play triangle inside the dark circle on the cover.
const PLAY = [
  ['M', 0.9444, 0.4355],
  ['C', 0.6481, 0.2903, 0.6481, 0.2903, 0.6481, 0.2903],
  ['C', 0.5926, 0.2581, 0.4815, 0.2097, 0.4259, 0.1774],
  ['C', 0.1296, 0.0323, 0.1296, 0.0323, 0.1296, 0.0323],
  ['C', 0.0556, 0.0000, 0.0000, 0.0323, 0.0000, 0.0968],
  ['L', 0.0000, 0.9032],
  ['C', 0.0000, 0.9677, 0.0556, 1.0000, 0.1296, 0.9677],
  ['C', 0.4074, 0.8226, 0.4074, 0.8226, 0.4074, 0.8226],
  ['C', 0.4815, 0.7903, 0.5926, 0.7419, 0.6481, 0.7097],
  ['C', 0.9444, 0.5645, 0.9444, 0.5645, 0.9444, 0.5645],
  ['C', 1.0000, 0.5323, 1.0000, 0.4677, 0.9444, 0.4355], ['Z'],
];

// Human head silhouette — first feature row of "Various Appliance Power Plants".
const HEAD_ICON = [
  ['M', 1.0000, 0.3965],
  ['C', 1.0000, 0.1775, 0.8020, 0.0000, 0.5577, 0.0000],
  ['C', 0.3134, 0.0000, 0.1154, 0.1775, 0.1154, 0.3965],
  ['C', 0.1154, 0.4023, 0.0000, 0.6206, 0.0000, 0.6206],
  ['L', 0.1158, 0.6206], ['L', 0.1155, 0.8275],
  ['C', 0.1154, 0.8656, 0.1498, 0.8965, 0.1923, 0.8966],
  ['L', 0.3654, 0.8966], ['L', 0.3654, 1.0000],
  ['L', 0.8846, 1.0000], ['L', 0.8846, 0.8437],
  ['C', 0.8846, 0.7763, 0.9009, 0.7096, 0.9325, 0.6483],
  ['L', 0.9517, 0.6111],
  ['C', 0.9835, 0.5492, 0.9999, 0.4819, 0.9995, 0.4138],
  ['C', 0.9998, 0.4080, 1.0000, 0.4023, 1.0000, 0.3965], ['Z'],
];

// Water drop ringed by three recycling arrows — second feature row.
const RECYCLE = [
  [0.311, 0.208, 0.378, 0.588, [ // drop
    ['M', 0.500, 0.000], ['L', 0.057, 0.532],
    ['C', -0.070, 0.688, 0.022, 0.881, 0.263, 0.963],
    ['C', 0.336, 0.988, 0.418, 1.001, 0.500, 1.000],
    ['C', 0.773, 1.002, 0.997, 0.861, 1.000, 0.685],
    ['C', 1.001, 0.632, 0.981, 0.579, 0.943, 0.532], ['Z']]],
  [0.535, 0.289, 0.465, 0.706, [ // arrow, lower right
    ['M', 1.000, 0.303],
    ['C', 1.000, 0.198, 0.966, 0.095, 0.898, 0.000],
    ['L', 0.847, 0.102], ['L', 0.708, 0.080],
    ['C', 0.892, 0.366, 0.690, 0.697, 0.257, 0.818],
    ['C', 0.208, 0.832, 0.158, 0.842, 0.106, 0.850],
    ['L', 0.000, 0.939], ['L', 0.118, 1.000],
    ['C', 0.628, 0.938, 0.999, 0.645, 1.000, 0.303], ['Z']]],
  [0.095, 0.000, 0.784, 0.272, [ // arrow, top
    ['M', 0.518, 0.000],
    ['C', 0.313, -0.002, 0.120, 0.283, 0.000, 0.764],
    ['L', 0.088, 0.764], ['L', 0.099, 0.995], ['L', 0.102, 1.000],
    ['C', 0.262, 0.339, 0.577, 0.180, 0.805, 0.644],
    ['C', 0.840, 0.715, 0.872, 0.797, 0.899, 0.891],
    ['L', 1.000, 0.891], ['L', 1.000, 0.643],
    ['C', 0.879, 0.236, 0.703, 0.001, 0.518, 0.000], ['Z']]],
  [0.000, 0.311, 0.465, 0.689, [ // arrow, lower left
    ['M', 1.000, 0.850],
    ['C', 0.531, 0.823, 0.184, 0.543, 0.224, 0.226],
    ['C', 0.229, 0.183, 0.242, 0.141, 0.261, 0.100],
    ['L', 0.202, 0.000], ['L', 0.059, 0.039],
    ['C', -0.135, 0.418, 0.163, 0.833, 0.724, 0.964],
    ['C', 0.801, 0.982, 0.881, 0.994, 0.962, 1.000],
    ['L', 0.873, 0.909], ['L', 1.000, 0.853], ['Z']]],
];

// Dripping tap on the lime tile of the philosophy slide.
const TAP = [
  [0.224, 0.030, 0.552, 0.113, 'round'],   // cross handle
  [0.466, 0.000, 0.069, 0.259, 'rect'],    // handle stem
  [0.397, 0.259, 0.207, 0.103, 'rect'],    // neck
  [0.310, 0.362, 0.379, 0.276, 'rect'],    // body
  [0.690, 0.397, 0.190, 0.207, 'rect'],    // back arm
  [0.879, 0.328, 0.121, 0.345, 'rect'],    // wall flange
  [0.000, 0.414, 0.310, 0.310, [           // spout elbow
    ['M', 1.000, 0.000], ['L', 1.000, 0.556], ['L', 0.667, 0.556],
    ['C', 0.605, 0.556, 0.556, 0.605, 0.556, 0.667],
    ['L', 0.556, 1.000], ['L', 0.000, 1.000], ['L', 0.000, 0.667],
    ['C', 0.000, 0.298, 0.298, 0.000, 0.667, 0.000], ['Z']]],
  [0.018, 0.793, 0.136, 0.207, [           // falling drop
    ['M', 0.500, 0.000], ['L', 0.049, 0.552],
    ['C', -0.066, 0.706, 0.030, 0.892, 0.263, 0.968],
    ['C', 0.337, 0.992, 0.418, 1.003, 0.500, 0.999],
    ['C', 0.760, 1.010, 0.983, 0.880, 0.999, 0.708],
    ['C', 1.004, 0.655, 0.988, 0.601, 0.952, 0.552], ['Z']]],
];

// Oil-drum with a leaf — "Purpose One".
const BARREL = [
  [0.234, 0.047, 0.531, 0.906, 'rect', BARREL_C.body],   // drum body
  [0.203, 0.047, 0.594, 0.063, 'round', BARREL_C.body],  // top rim
  [0.203, 0.469, 0.594, 0.063, 'round', BARREL_C.body],  // middle rim
  [0.203, 0.891, 0.594, 0.063, 'round', BARREL_C.body],  // bottom rim
  [0.234, 0.109, 0.531, 0.047, 'rect', BARREL_C.soft],   // band under top rim
  [0.234, 0.531, 0.531, 0.047, 'rect', BARREL_C.soft],   // band under middle rim
  [0.430, 0.641, 0.335, 0.312, [                         // shading on lower half
    ['M', 1.000, 0.000], ['L', 1.000, 1.000], ['L', 0.000, 1.000], ['Z']], BARREL_C.soft],
  [0.375, 0.344, 0.281, 0.281, [                         // leaf blade
    ['M', 0.667, 0.000], ['L', 1.000, 0.000], ['L', 1.000, 0.333],
    ['C', 1.000, 0.701, 0.701, 1.000, 0.333, 1.000],
    ['L', 0.000, 1.000], ['L', 0.000, 0.667],
    ['C', 0.000, 0.299, 0.299, 0.000, 0.667, 0.000], ['Z']], BARREL_C.leaf],
  [0.328, 0.391, 0.281, 0.281, [                         // leaf stem
    ['M', 0.920, 0.000], ['L', 1.000, 0.080], ['L', 0.080, 1.000],
    ['L', 0.000, 0.920], ['Z']], BARREL_C.vein],
  [0.500, 0.109, 0.188, 0.172, [                         // dripping cap
    ['M', 0.000, 0.000], ['L', 0.667, 0.000], ['L', 0.667, 0.818],
    ['L', 0.333, 0.818], ['L', 0.333, 0.364], ['L', 0.000, 0.364], ['Z']], BARREL_C.cap],
];

// Recycling wheelie bin — "Purpose Two".
const BIN = [
  [0.234, 0.051, 0.594, 0.078, 'round', BIN_C.lid],      // handle tab
  [0.141, 0.125, 0.719, 0.094, 'round', BIN_C.rim],      // lid
  [0.203, 0.219, 0.594, 0.734, [                         // tapered body
    ['M', 0.000, 0.000], ['L', 1.000, 0.000], ['L', 0.895, 1.000],
    ['L', 0.105, 1.000], ['Z']], BIN_C.body],
  [0.203, 0.219, 0.594, 0.047, 'rect', BIN_C.shade],     // rim shadow
  [0.263, 0.364, 0.521, 0.589, [                         // shading on the lower right
    ['M', 1.000, 0.000], ['L', 0.860, 1.000], ['L', 0.000, 1.000], ['Z']], BIN_C.shade],
  [0.332, 0.430, 0.344, 0.250, [                         // recycling arrows
    ['M', 0.500, 0.000], ['L', 1.000, 1.000], ['L', 0.000, 1.000], ['Z'],
    ['M', 0.500, 0.300], ['L', 0.200, 0.860], ['L', 0.800, 0.860], ['Z']], BIN_C.lid],
  [0.629, 0.738, 0.215, 0.215, 'ellipse', BIN_C.wheel],  // wheel
  [0.703, 0.812, 0.066, 0.066, 'ellipse', BIN_C.hub],
];

// One of the two comma marks on the testimonial avatars.
const QUOTE_COMMA = [
  ['M', 0.7857, 0.0000], ['L', 0.2143, 0.0000],
  ['C', 0.0960, 0.0000, 0.0000, 0.0480, 0.0000, 0.1071],
  ['L', 0.0000, 0.3929],
  ['C', 0.0000, 0.4520, 0.0960, 0.5000, 0.2143, 0.5000],
  ['L', 0.5714, 0.5000], ['L', 0.5714, 0.6429],
  ['C', 0.5714, 0.7217, 0.4433, 0.7857, 0.2857, 0.7857],
  ['L', 0.2500, 0.7857],
  ['C', 0.1906, 0.7857, 0.1429, 0.8096, 0.1429, 0.8393],
  ['L', 0.1429, 0.9464],
  ['C', 0.1429, 0.9761, 0.1906, 1.0000, 0.2500, 1.0000],
  ['L', 0.2857, 1.0000],
  ['C', 0.6804, 1.0000, 1.0000, 0.8402, 1.0000, 0.6429],
  ['L', 1.0000, 0.1071],
  ['C', 1.0000, 0.0480, 0.9040, 0.0000, 0.7857, 0.0000], ['Z'],
];

// Five-pointed rating star.
const STAR = [
  [0.494, 0.001], [0.335, 0.312], [0.001, 0.395], [0.234, 0.641], [0.191, 0.995],
  [0.500, 0.843], [0.802, 0.998], [0.766, 0.641], [0.999, 0.381], [0.664, 0.312],
];

// Twitter bird.
const BIRD = [
  [0.972, 0.019], [0.843, 0.080], [0.608, 0.022], [0.506, 0.147], [0.492, 0.310],
  [0.248, 0.228], [0.069, 0.047], [0.046, 0.230], [0.132, 0.383], [0.040, 0.352],
  [0.054, 0.446], [0.204, 0.602], [0.112, 0.607], [0.157, 0.701], [0.303, 0.782],
  [0.000, 0.886], [0.342, 0.999], [0.661, 0.864], [0.843, 0.584], [0.897, 0.249],
  [1.000, 0.119], [0.883, 0.157],
];

// Facebook "f".
const FACEBOOK = [
  [0.999, 0.012], [0.674, 0.001], [0.465, 0.041], [0.315, 0.161], [0.297, 0.381],
  [0.000, 0.381], [0.000, 0.562], [0.297, 0.562], [0.297, 1.000], [0.661, 1.000],
  [0.661, 0.562], [0.934, 0.562], [0.985, 0.383], [0.662, 0.381], [0.664, 0.241],
  [0.757, 0.175], [0.999, 0.166],
];

/* ================================================================== */
/* Repeated chrome                                                     */
/* ================================================================== */

function chevron(s, x, y, w, h, color, rotate) {
  freeform(s, CHEVRON, x, y, w, h, color, rotate ? { rotate } : {});
}

// Four staggered bars = the hamburger button; [xOffset, yOffset, width] fractions.
const BURGER_BARS = [
  [0.047, 0.000, 0.906], [0.234, 0.296, 0.766],
  [0.047, 0.593, 0.906], [0.281, 0.889, 0.719],
];

/** Browser-style bar across the top of every slide. */
function navBar(s, links = []) {
  roundRect(s, 0.812, 0.138, 12.290, 0.457, 0.5, { fill: { color: C.light } });

  s.addShape('ellipse', { x: 0.232, y: 0.131, w: 0.464, h: 0.464, fill: { color: C.accent } });
  freeform(s, LOGO_CUBE, 0.333, 0.226, 0.260, 0.274, C.ink);
  s.addText('Renogen', {
    x: 0.972, y: 0.195, w: 1.473, h: 0.337,
    fontFace: HEAD, fontSize: 14, color: C.ink, valign: 'top', margin: INSET,
  });

  links.forEach(([label, x]) => {
    s.addText(label, {
      x, y: 0.226, w: 1.301, h: 0.303, align: 'center',
      fontFace: HEAD, fontSize: 12, color: C.ink, valign: 'top', margin: INSET,
    });
  });

  s.addText('Menu', {
    x: 11.419, y: 0.226, w: 0.763, h: 0.303,
    fontFace: HEAD, fontSize: 12, color: C.ink, valign: 'top', margin: INSET,
  });
  s.addShape('ellipse', { x: 12.105, y: 0.307, w: 0.155, h: 0.155, fill: { color: C.ink } });
  chevron(s, 12.158, 0.347, 0.050, 0.082, C.accent, 90);
  BURGER_BARS.forEach(([bx, by, bw]) => {
    roundRect(s, 12.594 + bx * 0.2365, 0.268 + by * 0.2004, bw * 0.2365, 0.0223,
      0.5, { fill: { color: C.ink } });
  });
}

/** Scattered confetti squares; each entry is [x, y, size, rotation, colour]. */
function diamonds(s, list) {
  list.forEach(([x, y, size, rot, color]) => {
    s.addShape('roundRect', { x, y, w: size, h: size, rotate: rot, fill: { color } });
  });
}

// Triangle whose top edge bulges into a circle — the big lime shape of the
// opening and closing slides, taken from the source layout.
const BLOB = [
  ['M', 1.0000, 0.3122], ['L', 0.6056, 1.0000], ['L', 0.0000, 0.3122],
  ['L', 0.2513, 0.3122], ['L', 0.2520, 0.3122],
  ['C', 0.2662, 0.3103, 0.2784, 0.2980, 0.2844, 0.2807],
  ['L', 0.2848, 0.2794], ['L', 0.2873, 0.2689], ['L', 0.2883, 0.2620],
  ['L', 0.2885, 0.2576], ['L', 0.2889, 0.2576], ['L', 0.2889, 0.2625],
  ['L', 0.2914, 0.2521],
  ['C', 0.3017, 0.2156, 0.3160, 0.1806, 0.3344, 0.1485],
  ['C', 0.4328, -0.0231, 0.6113, -0.0499, 0.7332, 0.0885],
  ['C', 0.7713, 0.1318, 0.7993, 0.1861, 0.8169, 0.2455],
  ['L', 0.8253, 0.2784], ['L', 0.8287, 0.2882],
  ['C', 0.8344, 0.3003, 0.8433, 0.3089, 0.8533, 0.3121],
  ['L', 0.8537, 0.3122], ['Z'],
];

/** Lime blob plus the white circle punched through it. */
function limeBlob(s, x, rotate, flipH, ovalX) {
  freeform(s, BLOB, x, 1.884, 9.932, 7.057, C.accent, { rotate, flipH, shadow: cardShadow() });
  s.addShape('ellipse', { x: ovalX, y: 1.955, w: 5.058, h: 5.058, fill: { color: C.white } });
}

/** instagram / facebook / twitter row underneath each purpose card. */
function socialRow(s, x, y) {
  const line = { color: C.socialGray, width: 1.5 };
  roundRect(s, x, y, 0.203, 0.231, 0.28, { fill: { type: 'none' }, line });
  s.addShape('ellipse', { x: x + 0.055, y: y + 0.062, w: 0.093, h: 0.107, fill: { type: 'none' }, line });
  s.addShape('ellipse', { x: x + 0.148, y: y + 0.037, w: 0.024, h: 0.024, fill: { color: C.socialGray } });
  polygon(s, FACEBOOK, x + 0.406, y + 0.017, 0.169, 0.198, C.ink);
  polygon(s, BIRD, x + 0.752, y + 0.043, 0.237, 0.146, C.accent);
}

/* ================================================================== */
/* Shared copy                                                         */
/* ================================================================== */

const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do '
  + 'eiusmod tempor incididunt ut labore et dolore';
const LOREM_MAGNA = LOREM_SHORT + ' magna';
const LOREM_TESTIMONIAL = '“Lorem ipsum dolor sit amet elit, sed do lorem tempor incididunt '
  + 'ut labore et dolore magna aliqua. Ut enim ad et dolore magna ';
const LOREM_YEAR = 'Lorem ipsum dolor sit amet, conseur adipisicing elit, sed do eiusmod';
const LOREM_PHILOSOPHY = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do '
  + 'tempor incididunt ut labore minim enim.';
const LOREM_PROGRESS = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do '
  + 'eiusmod tempor incididunt ut labore';

/* ================================================================== */
/* Slide builders                                                      */
/* ================================================================== */

function slide1Cover(pres) {
  const s = pres.addSlide();
  limeBlob(s, 5.335, 321.0966, false, 7.767);
  navBar(s, [['Menu', 4.696], ['About Us', 6.016], ['Services', 7.487]]);

  heading(s, 'Renogen Presentation', 1.387, 2.270, 5.383, 1.717, 48);
  bodyText(s, LOREM_SHORT, 1.450, 4.307, 4.566, 0.627);

  roundRect(s, 1.549, 5.245, 1.503, 0.367, 0.5, { fill: { color: C.accent } });
  s.addText('Read More', {
    x: 1.549, y: 5.273, w: 1.503, h: 0.303, align: 'center',
    fontFace: HEAD, fontSize: 12, color: C.ink, valign: 'top', margin: INSET,
  });

  s.addShape('ellipse', { x: 3.552, y: 5.245, w: 0.365, h: 0.365, fill: { color: C.ink } });
  freeform(s, PLAY, 3.683, 5.356, 0.126, 0.145, C.accent);
  bodyText(s, 'Lorem ipsum dolor sit amet, consectetur', 3.964, 5.088, 2.052, 0.627);

  diamonds(s, [
    [5.534, 1.362, 0.265, 44.62, C.accent],
    [0.640, 4.483, 0.114, 44.62, C.ink],
    [3.873, 6.693, 0.183, 287.74, C.accent],
  ]);
}

function slide2Appliances(pres) {
  const s = pres.addSlide();
  navBar(s);

  heading(s, 'Various Appliance Power Plants', 1.450, 1.561, 5.117, 1.313, 36);

  // row 1 — head icon on a light tile
  roundRect(s, 1.554, 3.398, 0.502, 0.503, 0.125, { fill: { color: C.light } });
  freeform(s, HEAD_ICON, 1.654, 3.495, 0.271, 0.302, C.ink);
  caption(s, 'Second Appliance', 2.252, 3.489, 2.996);
  bodyText(s, LOREM_MAGNA + ' amet', 1.450, 3.999, 4.716, 0.627);

  // row 2 — recycling icon on a dark tile
  roundRect(s, 1.554, 5.106, 0.502, 0.503, 0.139, { fill: { color: C.ink } });
  icon(s, RECYCLE, 1.651, 5.199, 0.306, C.accent);
  caption(s, 'First Appliance', 2.248, 5.213, 2.721);
  bodyText(s, LOREM_MAGNA + ' aliqua. Ut enim ad minim veniam, quis nostrud dolore magna ',
    1.450, 5.747, 4.716, 0.904);

  diamonds(s, [
    [12.676, 5.664, 0.265, 44.62, C.accent],
    [0.662, 4.948, 0.114, 44.62, C.ink],
    [6.575, 3.230, 0.183, 287.74, C.accent],
  ]);
}

function slide3Journey(pres) {
  const s = pres.addSlide();
  navBar(s);

  heading(s, 'The Brief Journey of Our Company', 7.079, 1.865, 5.136, 1.313, 36);
  bodyText(s, LOREM_SHORT, 7.079, 3.691, 4.555, 0.627);

  // 2021 sits on a lime band, 2020 on the plain background
  roundRect(s, 6.817, 5.420, 4.946, 0.946, 0.0859, { fill: { color: C.accent } });
  [
    ['2020', 4.634, 4.545, 3.544, C.body],
    ['2021', 5.640, 5.551, 3.233, C.bodyDark],
  ].forEach(([year, yearY, textY, textW, textColor]) => {
    heading(s, year, 7.079, yearY, 1.057, 0.505, 24);
    bodyText(s, LOREM_YEAR, 8.280, textY, textW, 0.627, { color: textColor });
  });

  diamonds(s, [
    [0.572, 4.754, 0.265, 44.62, C.accent],
    [12.617, 4.734, 0.114, 44.62, C.ink],
    [8.014, 1.073, 0.183, 287.74, C.accent],
  ]);
}

function slide4Testimonials(pres) {
  const s = pres.addSlide();
  navBar(s);

  heading(s, 'The Renewable Energy for Better Environtment',
    2.378, 1.003, 8.577, 1.313, 36, { align: 'center' });

  const cards = [
    { x: 2.207, fill: C.accent, avatar: C.ink, mark: C.accent, copy: C.bodyDark,
      name: 'Thomas Eduardo', nameX: 2.486, quoteX: 2.480, avatarX: 5.678, avatarY: 2.946 },
    { x: 7.051, fill: C.light, avatar: C.gray, mark: C.light, copy: C.body,
      name: 'Robert Costa', nameX: 7.330, quoteX: 7.323, avatarX: 10.522, avatarY: 2.962 },
  ];
  cards.forEach((card) => {
    roundRect(s, card.x, 2.768, 4.075, 2.089, 0.0584, { fill: { color: card.fill } });
    s.addShape('ellipse', {
      x: card.avatarX, y: card.avatarY, w: 0.434, h: 0.434, fill: { color: card.avatar },
    });
    [0, 0.107].forEach((dx) => {
      freeform(s, QUOTE_COMMA, card.avatarX + 0.121 + dx, card.avatarY + 0.129,
        0.083, 0.166, card.mark);
    });
    caption(s, card.name, card.nameX, 3.163, 2.475, { bold: false });
    bodyText(s, LOREM_TESTIMONIAL, card.quoteX, 3.591, 3.803, 0.904,
      { italic: true, color: card.copy });
  });

  diamonds(s, [
    [12.418, 3.030, 0.265, 44.62, C.accent],
    [0.696, 5.712, 0.183, 287.74, C.ink],
  ]);
}

function slide5People(pres) {
  const s = pres.addSlide();
  navBar(s);

  heading(s, 'Important People of The Company', 1.437, 2.118, 5.122, 1.313, 36);
  caption(s, 'Dream Team', 1.437, 3.912, 2.605, { bold: false });
  bodyText(s, LOREM_SHORT, 1.450, 4.427, 4.549, 0.627);
  bodyText(s, LOREM_MAGNA + ' aliqua. Ut enim ad minim veniam, quis nostrud',
    1.450, 5.207, 4.549, 0.904);

  [['Alex Carlos', 7.318, 7.094], ['Lopes Holand', 9.890, 9.666]].forEach(([name, nameX, textX]) => {
    caption(s, name, nameX, 5.035, 1.835, { align: 'center' });
    s.addText(
      [
        { text: 'Lorem ipsum dolor sit amet, consectetur adipiscing ', options: { breakLine: true } },
        { text: 'elit, sed do' },
      ],
      { x: textX, y: 5.405, w: 2.283, h: 0.904, fontFace: BODY, fontSize: 11, color: C.body,
        align: 'center', valign: 'top', lineSpacingMultiple: 1.5, margin: INSET },
    );
  });

  diamonds(s, [
    [4.053, 1.271, 0.265, 44.62, C.accent],
    [0.639, 4.111, 0.114, 44.62, C.ink],
    [5.606, 6.805, 0.183, 287.74, C.accent],
    [12.663, 5.320, 0.171, 307.63, C.ink],
  ]);
}

function slide6Purposes(pres) {
  const s = pres.addSlide();
  navBar(s);

  heading(s, 'The Main Purposes of The Projects', 6.843, 1.291, 5.243, 1.313, 36);
  bodyText(s, LOREM_MAGNA, 6.829, 2.852, 4.716, 0.627);

  const cards = [
    { x: 5.334, art: BARREL, iconX: 6.633, title: 'Purpose One', titleX: 5.927,
      textX: 5.637, socialX: 6.410 },
    { x: 8.683, art: BIN, iconX: 9.983, title: 'Purpose Two', titleX: 9.277,
      textX: 8.987, socialX: 9.760 },
  ];
  cards.forEach((card) => {
    roundRect(s, card.x, 3.884, 3.141, 2.994, 0.0444,
      { fill: { color: C.white }, shadow: cardShadow() });
    icon(s, card.art, card.iconX, 4.241, 0.543);
    caption(s, card.title, card.titleX, 4.991, 1.953, { align: 'center' });
    bodyText(s, 'Lorem ipsum dolor sit amet elit, sed do eiusmod tempor',
      card.textX, 5.328, 2.535, 0.627, { align: 'center' });
    socialRow(s, card.socialX, 6.267);
  });

  diamonds(s, [[12.579, 3.939, 0.265, 44.62, C.accent]]);
}

function slide7Philosophy(pres) {
  const s = pres.addSlide();
  navBar(s);

  heading(s, 'Philosophy of Renewable Energy', 1.414, 1.180, 5.337, 1.313, 36);
  bodyText(s, LOREM_MAGNA, 1.446, 2.678, 4.716, 0.627);

  roundRect(s, 5.862, 5.312, 0.746, 0.746, 0.0922, { fill: { color: C.accent } });
  icon(s, TAP, 6.042, 5.492, 0.385, C.ink);

  const rows = [
    { num: '01', numX: 7.612, numY: 4.036, numW: 0.770, titleY: 4.188,
      title: 'Philosophy One', textX: 7.641, textY: 4.684 },
    { num: '02', numX: 7.641, numY: 5.601, numW: 0.971, titleY: 5.752,
      title: 'Philosophy Two', textX: 7.630, textY: 6.249 },
  ];
  rows.forEach((r) => {
    heading(s, r.num, r.numX, r.numY, r.numW, 0.640, 32, { bold: true, color: C.olive });
    caption(s, r.title, r.numX + 0.77, r.titleY, 2.285);
    bodyText(s, LOREM_PHILOSOPHY, r.textX, r.textY, 4.200, 0.628);
  });

  diamonds(s, [
    [12.553, 5.053, 0.265, 44.62, C.accent],
    [0.640, 4.483, 0.114, 44.62, C.ink],
  ]);
}

function slide8Rating(pres) {
  const s = pres.addSlide();
  navBar(s);

  heading(s, 'Success full Projects Bring People Trust', 1.423, 2.089, 5.675, 1.313, 36);
  caption(s, 'Top Company Rating', 1.451, 4.039, 2.679, { bold: false });
  chevron(s, 4.130, 4.166, 0.085, 0.140, C.ink);

  for (let i = 0; i < 5; i += 1) {
    polygon(s, STAR, 4.607 + i * 0.3887, 4.056, 0.276, 0.264, C.orange);
  }

  bodyText(s, LOREM_MAGNA + ' aliqua. Ut enim ad minim veniam, quis nostrud '
    + 'exercitation ullamco. Ut enim ad minim ', 1.471, 4.534, 5.627, 0.905);

  [
    { label: 'Detail One', x: 1.549, bullet: C.accent, textX: 1.874 },
    { label: 'Detail Two', x: 3.625, bullet: C.light, textX: 3.951 },
  ].forEach((d) => {
    s.addShape('ellipse', { x: d.x, y: 5.775, w: 0.194, h: 0.194, fill: { color: d.bullet } });
    chevron(s, d.x + 0.074, 5.830, 0.051, 0.083, C.ink);
    caption(s, d.label, d.textX, 5.699, 1.751, { bold: false });
  });

  diamonds(s, [
    [4.083, 1.106, 0.265, 44.62, C.accent],
    [0.617, 4.477, 0.114, 44.62, C.ink],
    [5.818, 6.808, 0.183, 287.74, C.accent],
  ]);
}

function slide9Progress(pres) {
  const s = pres.addSlide();
  navBar(s);

  heading(s, 'The Progress of Current Powerplant', 1.468, 1.334, 5.964, 1.313, 36);
  bodyText(s, LOREM_MAGNA + ' aliqua. Ut enim ad', 1.468, 2.882, 5.389, 0.627);

  const cards = [
    { y: 1.241, cardFill: C.light, cardAdj: 0.0748, title: 'Powerplant A', pct: '90%',
      track: C.mid, barX: 8.138, fillW: 2.297, titleX: 8.064, titleY: 1.653, barY: 2.152,
      pctX: 10.892, textX: 8.047, textY: 2.559, copy: C.body },
    { y: 3.731, cardFill: C.accent, cardAdj: 0.0628, title: 'Powerplant B', pct: '95%',
      track: C.white, barX: 8.140, fillW: 2.481, titleX: 8.065, titleY: 4.151, barY: 4.650,
      pctX: 10.893, textX: 8.049, textY: 5.056, copy: C.bodyDark },
  ];
  cards.forEach((c) => {
    roundRect(s, 7.809, c.y, 4.312, 2.344, c.cardAdj, { fill: { color: c.cardFill } });
    caption(s, c.title, c.titleX, c.titleY, 2.230);
    roundRect(s, c.barX, c.barY, 2.662, 0.306, 0.1667, { fill: { color: c.track } });
    roundRect(s, c.barX, c.barY, c.fillW, 0.306, 0.1667, { fill: { color: C.ink } });
    caption(s, c.pct, c.pctX, c.barY - 0.009, 0.667);
    bodyText(s, LOREM_PROGRESS, c.textX, c.textY, 3.918, 0.627, { color: c.copy });
  });

  diamonds(s, [
    [0.617, 2.452, 0.265, 44.62, C.accent],
    [12.735, 2.528, 0.114, 44.62, C.ink],
  ]);
}

function slide10ThankYou(pres) {
  const s = pres.addSlide();
  limeBlob(s, -1.935, 38.9034, true, 0.496);
  navBar(s);

  heading(s, 'Thank You For Attention', 7.164, 2.333, 4.921, 1.717, 48);
  bodyText(s, LOREM_SHORT, 7.226, 4.300, 4.566, 0.627);

  [
    { label: 'Read More', x: 7.296, fill: C.accent, color: C.ink },
    { label: 'End Slide', x: 9.082, fill: C.ink, color: C.accent },
  ].forEach((btn) => {
    roundRect(s, btn.x, 5.265, 1.503, 0.367, 0.5, { fill: { color: btn.fill } });
    s.addText(btn.label, {
      x: btn.x, y: 5.298, w: 1.503, h: 0.303, align: 'center',
      fontFace: HEAD, fontSize: 12, color: btn.color, valign: 'top', margin: INSET,
    });
  });

  diamonds(s, [
    [7.563, 1.310, 0.265, 44.62, C.accent],
    [12.735, 4.750, 0.114, 44.62, C.ink],
    [9.742, 6.808, 0.183, 287.74, C.accent],
  ]);
}

/* ================================================================== */
/* Build                                                               */
/* ================================================================== */

function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'RENOGEN', width: 13.333333, height: 7.5 });
  pres.layout = 'RENOGEN';
  pres.author = 'Renogen';
  pres.title = 'Renogen Presentation';

  [
    slide1Cover, slide2Appliances, slide3Journey, slide4Testimonials, slide5People,
    slide6Purposes, slide7Philosophy, slide8Rating, slide9Progress, slide10ThankYou,
  ].forEach((builder) => builder(pres));

  const out = path.join(__dirname, '07d5493b-a6dd-4892-87f8-2142669afd2f_grok_final.pptx');
  return pres.writeFile({ fileName: out }).then(() => console.log('wrote', out));
}

build().catch((err) => { console.error(err); process.exit(1); });
