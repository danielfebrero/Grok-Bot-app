/**
 * FINAXY - Financial Presentation Template (25 slides, 20" x 11.25")
 * Recreated with pptxgenjs. Raster images in the original are replaced by
 * programmatic placeholders / native shapes.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const C = {
  bg: 'F5F5F5',
  white: 'FFFFFF',
  navy: '203590',
  blue: '2B47BF',
  orange: 'F58C00',
  redOrange: 'F54300',
  head: '0D0D0D',      // tx1 lumMod 95%
  head85: '262626',    // tx1 lumMod 85%
  body: '808080',      // tx1 lumMod 50%
  ghost: 'BFBFBF',     // bg1 lumMod 75%
  track: 'E9EAEC',
};

const F = { head: 'Poppins SemiBold', black: 'Poppins ExtraBold', plain: 'Poppins', body: 'Rubik' };

// PowerPoint gradients are not exposed by pptxgenjs, so each gradient is
// flattened to the visually dominant mid tone of its two stops.
const GRAD = { orange: 'F26A00', blue: '253DA7' };

// pptxgenjs emits an empty <a:ln/> for `{type:'none'}`, which renderers fill in
// with a default outline, so borderless shapes get an explicitly invisible one.
const NO_LINE = { color: 'FFFFFF', transparency: 100, width: 0.25 };

const SHADOW_CARD = { type: 'outer', color: '000000', opacity: 0.09, blur: 48, offset: 30, angle: 90 };
const SHADOW_TILE = { type: 'outer', color: '000000', opacity: 0.12, blur: 30, offset: 3, angle: 135 };
const SHADOW_DOT = { type: 'outer', color: '000000', opacity: 0.27, blur: 4, offset: 3, angle: 45 };

/* -------------------------------------------------- reusable custom outlines */
// Point lists are normalised 0..1 over the shape box:
//   [x, y]                 -> lineTo (first entry is the moveTo)
//   [x, y, 1]              -> moveTo
//   [x, y, [x1,y1,x2,y2]]  -> cubic bezier with the two control points
//   'z'                    -> close
const PATH_CORNER_BLOB = [
  [0, 0, 1], [1, 0], [0.9998, 0.0011],
  [0.1836, 1, [0.8653, 0.5881, 0.5505, 1]],
  [0.0051, 0.9668, [0.1225, 1, 0.0628, 0.9886]],
  [0, 0.9643], 'z',
];
const PATH_BIG_BLOB = [
  [0.1542, 0, 1], [1, 0], [1, 0.9698], [0.9923, 0.9719],
  [0.7648, 1, [0.9204, 0.9902, 0.844, 1]],
  [0.0039, 0.439, [0.3688, 1, 0.0431, 0.7541]],
  [0, 0.3751],
  [0.0039, 0.3112],
  [0.1306, 0.0258, [0.017, 0.2062, 0.0619, 0.1089]], 'z',
];
const PATH_CHEVRON = [
  [0.2711, 0, 1], [0.3, 0],
  [0.4408, 0.0429, [0.3525, 0, 0.4031, 0.0154]], [0.9182, 0.391],
  [0.9182, 0.6088, [1, 0.4506, 1, 0.5492]], [0.4408, 0.9569],
  [0.3, 1, [0.403, 0.9846, 0.3525, 1]], [0.2711, 1],
  [0.1302, 0.7393, [0.0895, 1, 0, 0.8343]], [0.309, 0.6089],
  [0.309, 0.3911, [0.3907, 0.5493, 0.3907, 0.4507]], [0.1302, 0.2607],
  [0.2711, 0, [0, 0.1658, 0.0895, 0]], 'z',
];
const PATH_TRAIL = [
  [0.2707, 1, 1],
  [0.0523, 0.7975, [0.1581, 0.9524, 0.0781, 0.8783]],
  [0.13, 0.5615, [0.0264, 0.717, 0.0549, 0.6306]],
  [0.3047, 0.4775, [0.172, 0.5235, 0.2305, 0.4892]],
  [0.504, 0.5213, [0.3784, 0.4656, 0.4682, 0.4812]],
  [0.3739, 0.6614, [0.5531, 0.5758, 0.4758, 0.6486]],
  [0.101, 0.5918, [0.2721, 0.6739, 0.1671, 0.6396]],
  [0.0165, 0.3515, [0.0098, 0.5261, -0.0228, 0.4334]],
  [0.3007, 0.1734, [0.0554, 0.2697, 0.1653, 0.2008]],
  [0.4558, 0.1692, [0.3507, 0.1631, 0.4062, 0.1586]],
  [0.542, 0.2399, [0.5054, 0.1798, 0.5465, 0.2088]],
  [0.4343, 0.2986, [0.5375, 0.2715, 0.4875, 0.2965]],
  [0.298, 0.2582, [0.3816, 0.3008, 0.3298, 0.2832]],
  [0.3132, 0.0812, [0.2332, 0.2074, 0.2466, 0.1312]],
  [0.5948, 0.0017, [0.3794, 0.0315, 0.4879, 0.006]],
  [1, 0.0276, [0.7092, -0.0072, 0.9156, 0.0222]],
];
const PATH_PAPER_PLANE = [
  [0.978, 0, 1], [1, 0.9], [0.631, 0.757], [0.456, 1], [0.333, 0.657], [0, 0.683], 'z',
];

function scalePath(pts, x, y, w, h) {
  return pts.map(p => {
    if (p === 'z') return { close: true };
    const pt = { x: x + p[0] * w, y: y + p[1] * h };
    if (p[2] === 1) pt.moveTo = true;
    else if (Array.isArray(p[2])) {
      pt.curve = { type: 'cubic', x1: x + p[2][0] * w, y1: y + p[2][1] * h, x2: x + p[2][2] * w, y2: y + p[2][3] * h };
    }
    return pt;
  });
}

/* ------------------------------------------------------------ tiny helpers */
function freeform(slide, pts, box, opts) {
  slide.addShape('custGeom', Object.assign({
    x: box.x, y: box.y, w: box.w, h: box.h,
    points: scalePath(pts, 0, 0, box.w, box.h),
    line: NO_LINE,
  }, opts));
}

function rr(slide, x, y, w, h, radius, opts) {
  slide.addShape('roundRect', Object.assign({ x, y, w, h, rectRadius: radius, line: NO_LINE }, opts));
}

function pill(slide, x, y, w, h, opts) { rr(slide, x, y, w, h, Math.min(w, h) / 2, opts); }

function tx(slide, text, o) {
  slide.addText(text, Object.assign({ fontFace: F.body, fontSize: 16, color: C.body, valign: 'top' }, o));
}

function heading(slide, x, y, w, str, size, extra) {
  tx(slide, str, Object.assign({ x, y, w, h: size / 33, fontFace: F.head, fontSize: size, color: C.head, lineSpacing: size * 1.2 }, extra));
}

function kicker(slide, x, y, str) {
  tx(slide, str || 'Financial Presentation', { x, y, w: 3.763, h: 0.438, fontFace: F.head, fontSize: 20, color: C.blue });
}

function bodyText(slide, x, y, w, str, extra) {
  tx(slide, str, Object.assign({ x, y, w, h: 1.269, fontSize: 16, lineSpacingMultiple: 1.5 }, extra));
}

/* ------------------------------------------------------- decorative pieces */
// Quarter-circle "petal" that hugs a slide corner. dir: tl | tr | bl | br
function cornerBlob(slide, dir, color, x, y, w, h) {
  w = w || 2.396; h = h || 1.298;
  const place = {
    tl: { x: 0, y: 0 },
    tr: { x: 20 - w, y: 0, flipH: true },
    br: { x: 20 - w, y: 11.25 - h, rotate: 180 },
    bl: { x: 0, y: 11.25 - h, rotate: 180, flipH: true },
  }[dir];
  freeform(slide, PATH_CORNER_BLOB, { x: x === undefined ? place.x : x, y: y === undefined ? place.y : y, w, h },
    { fill: { color }, rotate: place.rotate, flipH: place.flipH });
}

function bigBlob(slide, x, y, w, h, color, opts) {
  freeform(slide, PATH_BIG_BLOB, { x, y, w: w || 5.708, h: h || 6.987 }, Object.assign({ fill: { color } }, opts));
}

// Scattered accent circles: [x, y, diameter, color]
function dots(slide, list) {
  list.forEach(([x, y, d, color]) => slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color }, line: NO_LINE }));
}

/* ------------------------------------------------------------ icon glyphs */
// Fractional boxes inside the icon square: [x, y, w, h, kind, rotate]
//   R filled rect · P filled pill · E filled ellipse · O outlined ellipse · B outlined rounded box
const GLYPHS = {
  screen: [[0.40, 0.08, 0.40, 0.40, 'B'], [0.50, 0.21, 0.20, 0.08, 'R'],
    [0.12, 0.16, 0.18, 0.07, 'R'], [0.05, 0.31, 0.25, 0.07, 'R'], [0.08, 0.62, 0.74, 0.28, 'Q']],
  nodes: [[0.08, 0.10, 0.46, 0.68, 'B'], [0.24, 0.30, 0.22, 0.08, 'R'], [0.24, 0.50, 0.22, 0.08, 'R'],
    [0.62, 0.42, 0.14, 0.06, 'R'], [0.70, 0.20, 0.20, 0.20, 'O'], [0.53, 0.62, 0.20, 0.20, 'O']],
  people: [[0.02, 0.10, 0.22, 0.22, 'E'], [0.00, 0.44, 0.28, 0.22, 'P'],
    [0.76, 0.10, 0.22, 0.22, 'E'], [0.72, 0.44, 0.28, 0.22, 'P'],
    [0.35, 0.02, 0.30, 0.30, 'E'], [0.24, 0.40, 0.52, 0.34, 'P']],
  search: [[0.04, 0.04, 0.62, 0.62, 'O'], [0.56, 0.60, 0.40, 0.13, 'P', 45]],
  bulb: [[0.22, 0.00, 0.56, 0.56, 'O'], [0.38, 0.50, 0.24, 0.10, 'R'],
    [0.34, 0.64, 0.32, 0.07, 'R'], [0.34, 0.76, 0.32, 0.07, 'R'], [0.40, 0.88, 0.20, 0.07, 'R']],
  link: [[0.00, 0.32, 0.58, 0.32, 'Q', -30], [0.42, 0.36, 0.58, 0.32, 'Q', -30]],
  puzzle: [[0.14, 0.14, 0.72, 0.72, 'R'], [0.38, 0.00, 0.26, 0.26, 'E'],
    [0.74, 0.36, 0.26, 0.26, 'E'], [0.00, 0.36, 0.26, 0.26, 'E'], [0.38, 0.74, 0.26, 0.26, 'E']],
};

function glyph(slide, name, x, y, s, color, stroke) {
  GLYPHS[name].forEach(([fx, fy, fw, fh, kind, rot]) => {
    const box = { x: x + fx * s, y: y + fy * s, w: fw * s, h: fh * s, rotate: rot };
    if (kind === 'R') slide.addShape('rect', Object.assign({ fill: { color }, line: NO_LINE }, box));
    else if (kind === 'E') slide.addShape('ellipse', Object.assign({ fill: { color }, line: NO_LINE }, box));
    else if (kind === 'O') slide.addShape('ellipse', Object.assign({ fill: { type: 'none' }, line: { color, width: stroke } }, box));
    else if (kind === 'P') slide.addShape('roundRect', Object.assign({ rectRadius: fh * s / 2, fill: { color }, line: NO_LINE }, box));
    else if (kind === 'Q') slide.addShape('roundRect', Object.assign({ rectRadius: fh * s / 2, fill: { type: 'none' }, line: { color, width: stroke } }, box));
    else slide.addShape('roundRect', Object.assign({ rectRadius: 0.1 * s, fill: { type: 'none' }, line: { color, width: stroke } }, box));
  });
}

// Gradient-filled rounded tile carrying a white line icon.
function iconTile(slide, x, y, w, h, color, name) {
  rr(slide, x, y, w, h, Math.min(w, h) / 6, { fill: { color }, shadow: SHADOW_TILE });
  const s = h * 0.63;
  glyph(slide, name, x + (w - s) / 2, y + (h - s) / 2, s, C.white, 1.5);
}

/* --------------------------------------------------------- content blocks */
// White "search field" bar with the orange MORE button pinned to its right.
function searchBar(slide, x, y, w) {
  pill(slide, x, y, w, 1.024, { fill: { color: C.white }, shadow: SHADOW_CARD });
  tx(slide, ' Your Text Here', { x: x + 0.1, y, w: w - 2.4, h: 1.024, fontFace: F.plain, fontSize: 18, color: C.ghost, valign: 'middle' });
  pill(slide, x + w - 2.177, y + 0.215, 1.908, 0.621, { fill: { color: GRAD.orange } });
  tx(slide, 'MORE', { x: x + w - 2.177, y: y + 0.215, w: 1.908, h: 0.621, fontFace: F.head, fontSize: 18, color: C.white, align: 'center', valign: 'middle' });
}

// "892+ / Lorem Ipsum" statistic.
function stat(slide, x, y) {
  tx(slide, '892+', { x, y, w: 1.97, h: 0.774, fontFace: F.black, fontSize: 40, color: C.navy, align: 'center' });
  tx(slide, 'Lorem Ipsum', { x: x - 0.09, y: y + 0.774, w: 2.15, h: 0.404, fontFace: F.head, fontSize: 18, color: C.head, align: 'center' });
}

const CARD_BODY = 'Lorem ipsum dolorasw amet, consectetur';

// Floating white card with an icon tile hanging off its top edge.
function sectionCard(slide, x, y, color, icon, title) {
  rr(slide, x, y, 3.25, 2.524, 0.413, { fill: { color: C.white }, shadow: SHADOW_CARD });
  tx(slide, title || 'Section Content', { x: x + 0.297, y: y + 0.663, w: 2.655, h: 0.438, fontFace: F.head, fontSize: 20, color: C.head, align: 'center' });
  tx(slide, CARD_BODY, { x: x + 0.297, y: y + 1.195, w: 2.655, h: 0.865, fontSize: 16, align: 'center', lineSpacingMultiple: 1.5 });
  iconTile(slide, x + 1.204, y - 0.383, 0.841, 0.793, color, icon);
}

// Icon tile with the heading + copy laid out to its right.
function inlineFeature(slide, x, y, color, icon, title, body, w) {
  iconTile(slide, x, y, 0.644, 0.607, color, icon);
  tx(slide, title || 'Section Content', { x: x + 0.822, y: y - 0.041, w: w || 2.655, h: 0.438, fontFace: F.head, fontSize: 20, color: C.head });
  tx(slide, body || CARD_BODY, { x: x + 0.822, y: y + 0.491, w: w || 2.655, h: 0.865, fontSize: 16, lineSpacingMultiple: 1.5 });
}

// "Section Financial One/Two" two column footer copy.
function financialCols(slide, x, y, gap, body, color, w) {
  ['Section Financial One', 'Section Financial Two'].forEach((title, i) => {
    tx(slide, title, { x: x + i * gap, y, w: 3.279, h: 0.438, fontFace: F.head, fontSize: 20, color: color || C.blue });
    tx(slide, body, { x: x + i * gap, y: y + 0.66, w: w || 3.726, h: 0.865, fontSize: 16, lineSpacingMultiple: 1.5 });
  });
}

// Vertical bar fading from `color` at the top to white at the bottom.
// Drawn as overlapping slices, each one lighter and painted over the previous.
const FADE_STEPS = 10;
function fadeBar(slide, x, y, w, h, color) {
  const step = h / FADE_STEPS;
  for (let i = 0; i < FADE_STEPS; i++) {
    const last = i === FADE_STEPS - 1;
    rr(slide, x, y + i * step, w, last ? step : step * 2, i === 0 ? w / 2 : 0,
      { fill: { color: mix(color, i / (FADE_STEPS - 1)) } });
  }
}

function mix(hex, toWhite) {
  let out = '';
  for (let i = 0; i < 3; i++) {
    const v = parseInt(hex.substr(i * 2, 2), 16);
    out += Math.round(v + (255 - v) * toWhite).toString(16).padStart(2, '0');
  }
  return out.toUpperCase();
}

// Tall pill-shaped card holding a 5 bar / 5 track mini chart ("Charts Data").
const CHART_TRACKS = [0.604, 0.952, 1.298, 1.658, 1.997];
const CHART_BARS = [[0.604, 1.024, GRAD.orange], [0.956, 0.948, GRAD.blue], [1.305, 1.346, GRAD.orange],
  [1.651, 0.735, GRAD.blue], [2.001, 0.948, GRAD.orange]];

function chartsDataCard(slide, x, y) {
  rr(slide, x, y, 2.718, 4.816, 1.359, { fill: { color: C.white }, shadow: SHADOW_CARD });
  CHART_TRACKS.forEach(dx => rr(slide, x + dx, y + 0.847, 0.18, 1.255, 0.09, { fill: { color: C.track } }));
  CHART_BARS.forEach(([dx, h, color]) => fadeBar(slide, x + dx, y + 2.63 - h, 0.18, h, color));
  tx(slide, 'Charts Data', { x: x + 0.31, y: y + 2.748, w: 2.085, h: 0.438, fontFace: F.head, fontSize: 20, color: C.head, align: 'center' });
  tx(slide, 'Lorem ipsum dora amet, consecte', { x: x + 0.194, y: y + 3.185, w: 2.33, h: 0.865, fontSize: 16, align: 'center', lineSpacingMultiple: 1.5 });
}

// Squat card with a 4 bar mini chart ("Data Charts"). `tall` = slide 2 variant.
const DATA_BARS = [[0.647, 0.579, GRAD.orange], [1.167, 0.821, GRAD.blue], [1.676, 0.743, GRAD.orange], [2.212, 1.05, GRAD.blue]];

function dataChartsCard(slide, x, y, tall) {
  const h = tall ? 3.963 : 3.553;
  const d = tall ? 0.298 : 0;
  rr(slide, x, y, 3.25, h, tall ? 1.625 : 0.534, { fill: { color: C.white }, shadow: SHADOW_CARD });
  DATA_BARS.forEach(([dx, bh, color]) => rr(slide, x + dx, y + d + 1.49 - bh, 0.327, bh, 0.055, { fill: { color } }));
  tx(slide, 'Data Charts', { x: x + 0.539, y: y + d + 1.683, w: 2.172, h: 0.471, fontFace: F.head, fontSize: 22, color: C.head, align: 'center' });
  tx(slide, 'Lorem ipsum dolora amet, consectetur', { x: x + 0.483, y: y + d + 2.182, w: 2.33, h: 0.865, fontSize: 16, align: 'center', lineSpacingMultiple: 1.5 });
}

/* ------------------------------------------------------------ copy blocks */
const LOREM = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas pora ttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus ma lesuada libero, sit amet commodo magna eros quis urna.';
const LOREM_W = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas poras ttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purusw lectus ma lesuada libero, sit amet commodo magna eros quis urna.';
const LOREM_COL = 'Lorem ipsum dolor sit amet, cons ectetuer adipiscing elit. Maecena pora ttitor congue massa';
const LOREM_COL_S = 'Lorem ipsum dolor sit amet, cons ectetuer adipiscing elit. Maecena';
const LOREM_PROC = 'Lorem ipsum dolor sit amet, consec tetuer adipiscing elit. Maecenas por ttitor congue massa. ';
const LOREM_CIRC = 'Lorem ipsum dolor sit amet, tetuer adipiscing elit. Maecr';

/* ------------------------------------------------------------- the slides */
const build = {};

build[1] = s => {
  slideCover(s, 'FINAXY', 6.211, 7.66);
};

build[25] = s => {
  slideCover(s, 'Thank You', 4.459, 11.082);
};

// Slides 1 / 25 share the four-petal cover frame.
function slideCover(s, title, titleX, titleW) {
  s.addShape('rect', { x: 0, y: 0, w: 20, h: 11.25, fill: { color: C.white, transparency: 10 }, line: NO_LINE });
  cornerBlob(s, 'tl', GRAD.orange); cornerBlob(s, 'br', GRAD.orange);
  cornerBlob(s, 'tr', GRAD.blue); cornerBlob(s, 'bl', GRAD.blue);
  tx(s, 'Financial Presentation Template ', { x: 6.322, y: 3.559, w: 7.11, h: 0.572, fontFace: F.head, fontSize: 28, color: C.orange, align: 'center' });
  tx(s, title, { x: titleX, y: 4.249, w: titleW, h: 2.423, fontFace: F.head, fontSize: 138, color: C.navy, align: 'center' });
  dots(s, [[1.541, 3.413, 0.524, GRAD.blue], [18.015, 3.559, 0.489, GRAD.orange],
    [14.812, 9.037, 0.915, GRAD.blue], [4.345, 9.07, 0.489, GRAD.orange]]);
  searchBar(s, 7.357, 6.926, 5.286);
}

build[2] = s => {
  bigBlob(s, 15.154, 0, 4.846, 5.932, GRAD.blue);
  rr(s, 11.561, 4.963, 4.492, 4.492, 2.246, { fill: { color: GRAD.orange } });
  cornerBlob(s, 'tl', GRAD.orange);
  tx(s, 'Financial Presentation Template ', { x: 1.636, y: 2.418, w: 7.11, h: 0.572, fontFace: F.head, fontSize: 28, color: C.orange });
  tx(s, 'FINAXY', { x: 1.525, y: 3.108, w: 7.66, h: 2.423, fontFace: F.head, fontSize: 138, color: C.navy });
  bodyText(s, 1.525, 6.01, 7.952, LOREM);
  searchBar(s, 1.412, 8.139, 5.286);
  stat(s, 7.125, 8.045);
  dataChartsCard(s, 14.804, 1.508, true);
  chartsDataCard(s, 11.121, 5.731);
  dots(s, [[9.137, 1.043, 0.524, GRAD.blue], [1.263, 9.974, 0.524, GRAD.blue], [14.31, 0.531, 0.641, GRAD.orange],
    [18.157, 1.865, 0.863, GRAD.orange], [9.661, 9.163, 0.489, GRAD.orange], [18.588, 10.406, 0.524, GRAD.blue]]);
};

build[3] = s => {
  bigBlob(s, 0.543, -0.543, 4.846, 5.932, GRAD.blue, { rotate: 270 });
  rr(s, 0.116, 2.979, 4.492, 4.492, 2.246, { fill: { color: GRAD.orange } });
  cornerBlob(s, 'tr', GRAD.orange);
  chartsDataCard(s, 6.152, 1.302);
  kicker(s, 10.827, 2.252);
  heading(s, 10.827, 2.848, 7.66, 'In the realm of assets, financial insight drives growth.', 48);
  bodyText(s, 10.827, 6.29, 7.952, LOREM);
  searchBar(s, 10.714, 8.419, 5.286);
  stat(s, 16.426, 8.325);
  dots(s, [[1.68, 0.554, 0.863, GRAD.orange], [10.0, 0.776, 0.641, GRAD.orange], [16.164, 1.062, 0.524, GRAD.blue],
    [19.297, 9.847, 0.489, GRAD.orange], [9.796, 10.246, 0.524, GRAD.blue]]);
};

build[4] = s => {
  rr(s, 12.431, 2.543, 3.226, 3.002, 0.585, { fill: { color: GRAD.orange }, shadow: SHADOW_CARD });
  rr(s, 15.886, 6.486, 3.226, 3.002, 0.585, { fill: { color: GRAD.blue }, shadow: SHADOW_CARD });
  cornerBlob(s, 'tl', GRAD.orange);
  kicker(s, 1.515, 1.939);
  heading(s, 1.515, 2.535, 7.952, 'Navigating the journey to wealth with financial strategy.', 48);
  bodyText(s, 1.515, 5.689, 7.952, LOREM);
  financialCols(s, 1.515, 7.678, 4.226, LOREM_COL);
  sectionCard(s, 11.073, 1.55, GRAD.blue, 'screen');
  sectionCard(s, 15.338, 8.175, GRAD.orange, 'nodes');
  dots(s, [[1.263, 9.974, 0.524, GRAD.blue], [8.699, 0.955, 0.641, GRAD.orange], [9.657, 9.485, 0.489, GRAD.orange],
    [19.007, 10.119, 0.524, GRAD.blue], [14.722, 0.361, 0.524, GRAD.blue]]);
};

build[5] = s => {
  rr(s, 2.233, 6.342, 3.226, 3.002, 0.585, { fill: { color: GRAD.orange }, shadow: SHADOW_CARD });
  rr(s, 0.806, 1.054, 3.226, 3.002, 0.585, { fill: { color: GRAD.blue }, shadow: SHADOW_CARD });
  cornerBlob(s, 'tr', GRAD.orange);
  sectionCard(s, 5.93, 3.101, GRAD.blue, 'screen');
  sectionCard(s, 1.798, 5.954, GRAD.orange, 'nodes');
  kicker(s, 10.929, 2.241);
  heading(s, 10.929, 2.837, 7.66, 'Financial planning: where dreams become achievable.', 44);
  bodyText(s, 10.929, 5.985, 7.952, LOREM);
  bodyText(s, 10.929, 7.664, 7.952, LOREM);
  dots(s, [[1.263, 9.974, 0.524, GRAD.blue], [8.699, 0.955, 0.641, GRAD.orange], [18.789, 9.34, 0.489, GRAD.orange],
    [12.195, 10.236, 0.524, GRAD.blue], [14.722, 0.361, 0.524, GRAD.blue]]);
};

build[6] = s => {
  rr(s, 12.022, 5.204, 3.226, 3.002, 0.585, { fill: { color: GRAD.orange }, shadow: SHADOW_CARD });
  rr(s, 16.308, 3.703, 3.226, 3.002, 0.585, { fill: { color: GRAD.blue }, shadow: SHADOW_CARD });
  cornerBlob(s, 'tl', GRAD.orange);
  kicker(s, 1.388, 1.939);
  heading(s, 1.388, 2.535, 7.952, 'Navigating financial decisions with precision.', 44);
  bodyText(s, 1.388, 4.684, 7.952, LOREM);
  inlineFeature(s, 1.483, 6.923, GRAD.blue, 'screen');
  inlineFeature(s, 5.61, 6.927, GRAD.orange, 'nodes');
  bodyText(s, 1.388, 8.878, 7.952, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas pora ttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus');
  dots(s, [[8.699, 0.955, 0.641, GRAD.orange], [19.043, 8.483, 0.489, GRAD.orange],
    [11.162, 9.145, 0.524, GRAD.blue], [14.722, 0.361, 0.524, GRAD.blue]]);
};

build[7] = s => {
  rr(s, 0.989, 6.67, 3.226, 3.002, 0.585, { fill: { color: GRAD.orange }, shadow: SHADOW_CARD });
  cornerBlob(s, 'tr', GRAD.orange);
  sectionCard(s, 5.298, 6.631, GRAD.blue, 'screen');
  kicker(s, 10.929, 2.241);
  heading(s, 10.929, 2.837, 7.66, 'Financial planning: where dreams become achievable.', 44);
  bodyText(s, 10.929, 5.723, 7.952, LOREM);
  searchBar(s, 10.929, 7.666, 7.66);
  dots(s, [[1.263, 9.974, 0.524, GRAD.blue], [8.699, 0.955, 0.641, GRAD.orange], [14.722, 0.361, 0.524, GRAD.blue],
    [9.25, 9.044, 0.489, GRAD.orange], [18.277, 9.585, 0.524, GRAD.blue]]);
};

build[8] = s => {
  rr(s, 1.141, 6.603, 17.719, 3.275, 0.369, { fill: { color: C.white }, shadow: SHADOW_CARD });
  cornerBlob(s, 'tl', GRAD.orange); cornerBlob(s, 'tr', GRAD.orange);
  kicker(s, 1.792, 1.765);
  heading(s, 1.792, 2.361, 7.952, 'Navigating financial decisions with precision.', 44);
  bodyText(s, 1.792, 4.509, 7.952, LOREM);
  sectionCard(s, 11.073, 1.55, GRAD.blue, 'screen');
  sectionCard(s, 11.073, 5.438, GRAD.orange, 'nodes');
  financialCols(s, 1.792, 7.365, 4.226, LOREM_COL_S);
  dots(s, [[18.62, 6.218, 0.641, GRAD.orange], [11.37, 9.449, 0.524, GRAD.blue],
    [0.774, 9.974, 0.489, GRAD.orange], [6.186, 0.512, 0.524, GRAD.blue]]);
};

build[9] = s => {
  cornerBlob(s, 'tr', GRAD.orange, 12.087, 0, 7.913, 4.286);
  rr(s, 3.403, 1.248, 14.841, 9.276, 0.99, { fill: { color: C.white, transparency: 8 }, shadow: SHADOW_CARD });
  kicker(s, 4.983, 2.324);
  heading(s, 4.983, 2.921, 11.16, '\u201CIn the world of finance, smart decisions pave the way to prosperity.\u201D', 60, { italic: true });
  bodyText(s, 4.983, 6.372, 11.682, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas pora ttitor congue massa. Fusce posu, magna sed pulvinar ultricies, purus lectus ma lesuada libero, sit amet commodo magna eros quis urna.');
  searchBar(s, 4.983, 8.015, 11.097);
  dots(s, [[10.0, 0.776, 0.641, GRAD.orange], [16.783, 1.881, 0.524, GRAD.blue],
    [19.297, 9.847, 0.489, GRAD.orange], [9.796, 10.246, 0.524, GRAD.blue]]);
};

build[10] = s => {
  bigBlob(s, 16.053, 0, 3.947, 4.831, GRAD.blue);
  rr(s, 12.545, 7.762, 3.011, 3.011, 1.505, { fill: { color: GRAD.orange } });
  cornerBlob(s, 'tl', GRAD.orange);
  kicker(s, 1.388, 2.653);
  heading(s, 1.388, 3.25, 7.952, 'Navigating the markets with financial expertise.', 44);
  bodyText(s, 1.388, 5.398, 7.952, LOREM);
  inlineFeature(s, 1.483, 7.637, GRAD.blue, 'screen');
  inlineFeature(s, 5.61, 7.641, GRAD.orange, 'nodes');
  dots(s, [[8.699, 0.955, 0.641, GRAD.orange], [10.289, 8.73, 0.524, GRAD.blue], [18.318, 0.635, 0.641, GRAD.orange],
    [15.857, 2.012, 2.461, GRAD.orange], [1.198, 9.959, 0.438, GRAD.orange]]);
};

build[11] = s => {
  cornerBlob(s, 'tr', GRAD.orange);
  sectionCard(s, 4.707, 0.995, GRAD.blue, 'screen');
  sectionCard(s, 4.707, 4.363, GRAD.orange, 'nodes');
  sectionCard(s, 4.707, 7.731, GRAD.blue, 'screen');
  kicker(s, 9.484, 2.241);
  heading(s, 9.484, 2.837, 9.14, 'In the realm of assets, financial insight drives growth.', 40);
  bodyText(s, 9.484, 4.829, 9.305, LOREM + ' elit. Maecenas pora ttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus ma lesuada libero, sit amet commodo magna eros quis urna.');
  [9.54, 12.638, 15.97].forEach(x => {
    tx(s, 'Section Content', { x, y: 7.585, w: 2.655, h: 0.438, fontFace: F.head, fontSize: 20, color: C.head });
    tx(s, CARD_BODY, { x, y: 8.117, w: 2.655, h: 0.865, fontSize: 16, lineSpacingMultiple: 1.5 });
  });
  dots(s, [[1.263, 9.974, 0.524, GRAD.blue], [8.699, 0.955, 0.641, GRAD.orange], [18.789, 9.34, 0.489, GRAD.orange],
    [12.195, 10.236, 0.524, GRAD.blue], [14.722, 0.361, 0.524, GRAD.blue]]);
};

build[12] = s => {
  cornerBlob(s, 'tl', GRAD.orange); cornerBlob(s, 'br', GRAD.orange);
  cornerBlob(s, 'tr', GRAD.blue); cornerBlob(s, 'bl', GRAD.blue);
  searchBar(s, 6.322, 2.901, 7.23);
  tx(s, 'Break Slides', { x: 3.853, y: 4.249, w: 12.293, h: 2.423, fontFace: F.head, fontSize: 138, color: C.navy, align: 'center' });
  tx(s, 'Financial Presentation Template ', { x: 6.322, y: 6.947, w: 7.11, h: 0.572, fontFace: F.head, fontSize: 28, color: C.orange, align: 'center' });
  dots(s, [[1.541, 3.413, 0.524, GRAD.blue], [18.015, 3.559, 0.489, GRAD.orange],
    [14.812, 9.037, 0.915, GRAD.blue], [4.345, 9.07, 0.489, GRAD.orange]]);
};

build[13] = s => {
  cornerBlob(s, 'tl', GRAD.orange); cornerBlob(s, 'tr', GRAD.orange);
  kicker(s, 1.292, 1.551);
  heading(s, 1.292, 2.148, 7.952, 'Our Services Finaxy Business Financial', 54);
  bodyText(s, 1.292, 4.342, 7.952, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas pora ttitor congue massa. Fusce posuere, magna sed pulvinar');
  heading(s, 10.512, 2.113, 6.408, 'Fusce posuere, magna sed pulvinar ultricies', 32);
  bodyText(s, 10.512, 3.97, 7.952, LOREM);
  rr(s, 1.292, 6.312, 17.5, 3.696, 0.56, { fill: { color: C.white }, shadow: SHADOW_CARD });
  const svcBody = ['Lorem ipsum dolor sita amet, consectetuer ad ipiscing elit. ', 'Lorem ipsum dolor sita amet, consectetuer ad ipiscing elit. ',
    'Lorem ipsum dolor sita amet, consectetuer ad ipiscing elit. ', 'Lorem ipsum dolor sita amet, consectetuer ad ipiscing elit. '];
  [2.198, 6.324, 10.382, 14.509].forEach((x, i) => {
    inlineFeature(s, x, i % 2 ? 7.336 : 7.332, i % 2 ? GRAD.orange : GRAD.blue, i % 2 ? 'nodes' : 'screen', null, svcBody[i]);
  });
  dots(s, [[18.802, 2.177, 0.524, GRAD.blue], [0.254, 6.388, 0.641, GRAD.orange],
    [8.872, 10.254, 0.524, GRAD.blue], [18.464, 10.059, 0.321, GRAD.orange]]);
};

build[14] = s => {
  cornerBlob(s, 'tr', GRAD.orange);
  // Cards are stacked back-to-front: two, three, then one on top.
  const SERVICES = [
    { card: [4.266, 4.321], title: 'Section Services Two', tX: 4.952, tY: 5.28, bX: 4.679, bY: 5.811, tile: [6.206, 4.018], color: GRAD.orange, icon: 'nodes' },
    { card: [0.871, 6.767], title: 'Section Services Three', tX: 1.558, tY: 7.725, bX: 1.284, bY: 8.257, tile: [2.811, 6.464], color: GRAD.blue, icon: 'screen' },
    { card: [1.079, 1.227], title: 'Section Services One', tX: 1.766, tY: 2.185, bX: 1.492, bY: 2.717, tile: [3.019, 0.924], color: GRAD.blue, icon: 'screen' },
  ];
  SERVICES.forEach(v => {
    rr(s, v.card[0], v.card[1], 4.849, 3.271, 0.301, { fill: { color: C.white }, shadow: SHADOW_CARD });
    tx(s, v.title, { x: v.tX, y: v.tY, w: 3.476, h: 0.438, fontFace: F.head, fontSize: 20, color: C.head, align: 'center' });
    tx(s, 'Lorem ipsum dolor sit amet, consec tetuer adipiscing elit. Maecenas poa ttitor congue massa. ',
      { x: v.bX, y: v.bY, w: 4.024, h: 1.269, fontSize: 16, align: 'center', lineSpacingMultiple: 1.5 });
    iconTile(s, v.tile[0], v.tile[1], 0.97, 0.914, v.color, v.icon);
  });
  kicker(s, 10.442, 1.817);
  heading(s, 10.442, 2.413, 7.952, 'Our Services Finaxy Business Financial', 54);
  bodyText(s, 10.442, 5.095, 8.335, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas pora ttitr congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus ma lesuada libero, sit amet commodo magna eros quis urna. elit. Maecenas pora ttitor congue massa. Fusce posuere, magna');
  [10.498, 13.595].forEach(x => {
    tx(s, 'Section Content', { x, y: 7.851, w: 2.655, h: 0.438, fontFace: F.head, fontSize: 20, color: C.head });
    tx(s, CARD_BODY, { x, y: 8.382, w: 2.655, h: 0.865, fontSize: 16, lineSpacingMultiple: 1.5 });
  });
  stat(s, 16.551, 7.86);
  dots(s, [[8.699, 0.955, 0.641, GRAD.orange], [17.604, 10.416, 0.489, GRAD.orange],
    [8.816, 9.34, 0.524, GRAD.blue], [14.722, 0.361, 0.524, GRAD.blue]]);
};

// Slides 15-18: the SWOT letter slides. side = 'right' puts the letter block right.
function swotSlide(s, cfg) {
  bigBlob(s, cfg.blob[0], cfg.blob[1], 5.708, 6.987, cfg.blobColor, { rotate: cfg.blob[2], flipH: cfg.blob[3] });
  dots(s, cfg.blobDots);
  rr(s, cfg.shadowCard[0], cfg.shadowCard[1], 5.417, 5.417, 0.82, { fill: { color: C.white }, shadow: SHADOW_CARD });
  rr(s, cfg.letterCard[0], cfg.letterCard[1], 5.417, 5.417, 0.82, { fill: { color: cfg.letterColor }, shadow: SHADOW_CARD });
  tx(s, cfg.letter, { x: cfg.letterCard[0], y: cfg.letterCard[1], w: 5.417, h: 5.417, fontFace: F.black, fontSize: 287, color: C.white, align: 'center', valign: 'middle' });
  cornerBlob(s, cfg.corner, GRAD.orange);
  dataChartsCard(s, cfg.chart[0], cfg.chart[1]);
  kicker(s, cfg.textX, 2.18);
  heading(s, cfg.textX, 2.776, 7.952, cfg.title, 72);
  bodyText(s, cfg.bodyX, 5.685, 7.952, LOREM_W);
  dots(s, cfg.dots);
}

build[15] = s => {
  swotSlide(s, {
    blob: [14.292, 0], blobColor: GRAD.blue, blobDots: [[18.318, 0.635, 0.641, GRAD.orange], [16.063, 1.57, 1.458, GRAD.orange]],
    shadowCard: [13.0, 4.987], letterCard: [12.104, 4.279], letterColor: GRAD.orange, letter: 'S',
    corner: 'tl', chart: [10.927, 1.523], textX: 1.525, bodyX: 1.562, title: 'Strengths Analysis Slides',
    dots: [[8.699, 0.955, 0.641, GRAD.orange], [10.289, 8.73, 0.524, GRAD.blue], [1.198, 9.959, 0.438, GRAD.orange]],
  });
  financialCols(s, 1.562, 7.674, 4.226, LOREM_COL_S);
};

build[16] = s => {
  swotSlide(s, {
    blob: [0, 0, 0, true], blobColor: GRAD.orange, blobDots: [[1.058, 1.046, 0.693, GRAD.blue], [2.12, 1.555, 1.837, GRAD.blue]],
    shadowCard: [1.583, 4.987], letterCard: [2.479, 4.279], letterColor: GRAD.blue, letter: 'W',
    corner: 'tr', chart: [5.177, 1.215], textX: 10.426, bodyX: 10.463, title: 'Weaknessess Analysis Slides',
    dots: [[9.006, 1.174, 0.641, GRAD.orange], [17.604, 10.416, 0.489, GRAD.orange],
      [8.816, 9.34, 0.524, GRAD.blue], [15.476, 0.784, 0.524, GRAD.blue]],
  });
  [10.498, 13.595].forEach(x => {
    tx(s, 'Section Content', { x, y: 7.719, w: 2.655, h: 0.438, fontFace: F.head, fontSize: 20, color: C.head });
    tx(s, CARD_BODY, { x, y: 8.251, w: 2.655, h: 0.865, fontSize: 16, lineSpacingMultiple: 1.5 });
  });
  stat(s, 16.551, 7.86);
};

build[17] = s => {
  swotSlide(s, {
    blob: [14.292, 4.263, 180, true], blobColor: GRAD.orange,
    blobDots: [[18.249, 9.512, 0.693, GRAD.blue], [16.043, 7.858, 1.837, GRAD.blue]],
    shadowCard: [13.0, 0.846], letterCard: [12.104, 1.555], letterColor: GRAD.blue, letter: 'O',
    corner: 'tl', chart: [11.208, 6.142], textX: 1.525, bodyX: 1.562, title: 'Opportunities Analysis Slides',
    dots: [[8.699, 0.955, 0.641, GRAD.orange], [10.031, 5.618, 0.524, GRAD.blue], [1.198, 9.959, 0.438, GRAD.orange]],
  });
  inlineFeature(s, 1.636, 7.716, GRAD.blue, 'screen');
  inlineFeature(s, 5.941, 7.72, GRAD.orange, 'nodes');
};

build[18] = s => {
  swotSlide(s, {
    blob: [0, 4.263, 180], blobColor: GRAD.blue,
    blobDots: [[1.041, 9.974, 0.641, GRAD.orange], [2.479, 8.222, 1.458, GRAD.orange]],
    shadowCard: [1.583, 0.846], letterCard: [2.468, 1.555], letterColor: GRAD.orange, letter: 'T',
    corner: 'tr', chart: [5.177, 6.142], textX: 10.426, bodyX: 10.463, title: 'Threatss Analysis Slides',
    dots: [[9.006, 1.174, 0.641, GRAD.orange], [17.604, 10.416, 0.489, GRAD.orange],
      [9.327, 9.116, 0.524, GRAD.blue], [15.476, 0.784, 0.524, GRAD.blue]],
  });
  searchBar(s, 10.463, 7.661, 7.66);
};

build[19] = s => {
  cornerBlob(s, 'tl', GRAD.orange); cornerBlob(s, 'tr', GRAD.orange);
  // Four arrow ribbons; each is a run of chevrons sized from the ribbon height.
  const RIBBONS = [
    { x: 1.047, y: 6.652, w: 2.979, h: 0.524, n: 10, color: GRAD.orange },
    { x: 4.213, y: 6.481, w: 3.96, h: 0.871, n: 8, color: GRAD.blue },
    { x: 8.361, y: 6.235, w: 4.655, h: 1.365, n: 6, color: GRAD.orange },
    { x: 13.203, y: 6.102, w: 5.536, h: 1.624, n: 6, color: GRAD.blue },
  ];
  RIBBONS.forEach(r => {
    const cw = 0.6234 * r.h;
    const pitch = (r.w - cw) / (r.n - 1);
    for (let i = 0; i < r.n; i++) freeform(s, PATH_CHEVRON, { x: r.x + i * pitch, y: r.y, w: cw, h: r.h }, { fill: { color: r.color } });
  });
  const STEPS = [
    { year: '2023', yx: 1.524, yy: 8.132, yw: 2.025, tx: 1.356, ty: 3.996, bx: 0.898, by: 4.595 },
    { year: '2024', yx: 5.15, yy: 4.713, yw: 2.086, tx: 4.924, ty: 8.26, bx: 4.466, by: 8.859 },
    { year: '2025', yx: 9.659, yy: 8.132, yw: 2.058, tx: 9.522, ty: 3.996, bx: 9.064, by: 4.595 },
    { year: '2026', yx: 14.588, yy: 4.713, yw: 2.043, tx: 14.226, ty: 8.26, bx: 13.768, by: 8.859 },
  ];
  STEPS.forEach(p => {
    tx(s, p.year, { x: p.yx, y: p.yy, w: p.yw, h: 1.01, fontFace: F.black, fontSize: 54, color: C.blue, align: 'center' });
    tx(s, 'Process One', { x: p.tx, y: p.ty, w: 2.765, h: 0.505, fontFace: F.head, fontSize: 24, color: C.head85, align: 'center' });
    tx(s, LOREM_PROC, { x: p.bx, y: p.by, w: 3.681, h: 1.128, fontSize: 14, align: 'center', lineSpacingMultiple: 1.5 });
  });
  tx(s, 'Financial Presentation', { x: 8.119, y: 0.981, w: 3.763, h: 0.438, fontFace: F.head, fontSize: 20, color: C.blue, align: 'center' });
  heading(s, 4.293, 1.732, 11.414, 'ProcessTimeline Infographic ', 54, { align: 'center' });
  dots(s, [[18.802, 2.177, 0.524, GRAD.blue], [8.119, 3.338, 0.389, GRAD.orange],
    [12.744, 10.233, 0.458, GRAD.orange], [0.898, 2.714, 0.524, GRAD.blue]]);
};

build[20] = s => {
  cornerBlob(s, 'tl', GRAD.orange); cornerBlob(s, 'tr', GRAD.orange);
  // Five concentric-circle nodes; ring / core sizes are fixed ratios of the outer disc.
  const NODES = [
    { x: 1.258, y: 3.926, d: 4.346, color: GRAD.blue, icon: 'people', arrow: [3.431, 6.958, 1.601], label: 'Section One', tx: 2.049, bx: 1.82 },
    { x: 5.21, y: 4.112, d: 3.973, color: GRAD.orange, icon: 'search', arrow: [7.201, 6.865, 1.693], label: 'Section Two', tx: 5.815, bx: 5.586 },
    { x: 8.804, y: 4.292, d: 3.614, color: GRAD.blue, icon: 'bulb', arrow: [10.566, 6.813, 1.746], label: 'Section Three', tx: 9.147, bx: 8.918 },
    { x: 12.146, y: 4.451, d: 3.296, color: GRAD.orange, icon: 'link', arrow: [13.782, 6.727, 1.832], label: 'Section Four', tx: 12.406, bx: 12.177 },
    { x: 15.277, y: 4.629, d: 2.941, color: GRAD.blue, icon: 'puzzle', arrow: [16.747, 6.68, 1.756], label: 'Section Five', tx: 15.506, bx: 15.277 },
  ];
  NODES.forEach(n => {
    s.addShape('ellipse', { x: n.x, y: n.y, w: n.d, h: n.d, fill: { color: n.color, transparency: 30 }, line: NO_LINE });
    s.addShape('ellipse', { x: n.x + 0.2336 * n.d, y: n.y + 0.2336 * n.d, w: 0.5326 * n.d, h: 0.5326 * n.d, fill: { color: '000000', transparency: 80 }, line: NO_LINE });
    const id = 0.3948 * n.d;
    s.addShape('ellipse', { x: n.x + 0.3023 * n.d, y: n.y + 0.3023 * n.d, w: id, h: id, fill: { color: n.color }, line: NO_LINE, shadow: SHADOW_DOT });
    glyph(s, n.icon, n.x + 0.3023 * n.d + id * 0.25, n.y + 0.3023 * n.d + id * 0.25, id * 0.5, C.white, 2.5);
    s.addShape('line', { x: n.arrow[0], y: n.arrow[1], w: 0, h: n.arrow[2], line: { color: n.color, width: 2.5, endArrowType: 'triangle' } });
    tx(s, n.label, { x: n.tx, y: 8.817, w: 2.765, h: 0.505, fontFace: F.head, fontSize: 24, color: C.head85, align: 'center' });
    tx(s, LOREM_CIRC, { x: n.bx, y: 9.415, w: 3.223, h: 0.865, fontSize: 16, align: 'center', lineSpacingMultiple: 1.5 });
  });
  kicker(s, 2.614, 1.045);
  heading(s, 2.614, 1.641, 7.035, 'Decreasing Circles Infographic ', 48);
  financialCols(s, 10.472, 1.535, 4.226, LOREM_COL_S);
  dots(s, [[8.06, 0.296, 0.389, GRAD.orange], [0.898, 2.714, 0.524, GRAD.blue], [18.897, 3.853, 0.524, GRAD.blue]]);
};

build[21] = s => {
  cornerBlob(s, 'tl', GRAD.orange); cornerBlob(s, 'tr', GRAD.orange);
  freeform(s, PATH_TRAIL, { x: 5.518, y: -0.515, w: 6.968, h: 13.287 },
    { fill: { type: 'none' }, rotate: 253.274, line: { color: 'BFBFBF', width: 1.75, dashType: 'dash' } });
  freeform(s, PATH_PAPER_PLANE, { x: 14.58, y: 1.95, w: 3.6, h: 3.0 }, { fill: { color: GRAD.blue } });
  const STEPS = [
    { n: '1', x: 1.373, y: 4.088, color: GRAD.blue, tx: 3.05, ty: 4.104, by: 4.703, label: 'Process One' },
    { n: '2', x: 6.938, y: 9.034, color: GRAD.orange, tx: 9.006, ty: 8.643, by: 9.242, label: 'Process Two' },
    { n: '3', x: 8.708, y: 5.304, color: GRAD.blue, tx: 10.754, ty: 4.927, by: 5.525, label: 'Process Three' },
    { n: '4', x: 15.018, y: 5.206, color: GRAD.orange, tx: 15.535, ty: 7.04, by: 7.638, label: 'Process Four' },
  ];
  STEPS.forEach(p => {
    s.addShape('ellipse', { x: p.x, y: p.y, w: 1.276, h: 1.276, fill: { color: p.color }, line: NO_LINE });
    tx(s, p.n, { x: p.x, y: p.y, w: 1.276, h: 1.276, fontFace: F.black, fontSize: 48, color: C.white, align: 'center', valign: 'middle' });
    tx(s, p.label, { x: p.tx, y: p.ty, w: 2.765, h: 0.505, fontFace: F.head, fontSize: 24, color: C.head85 });
    tx(s, LOREM_PROC, { x: p.tx, y: p.by, w: 3.681, h: 1.128, fontSize: 14, lineSpacingMultiple: 1.5 });
  });
  kicker(s, 3.005, 0.919);
  heading(s, 3.005, 1.515, 11.326, 'Proccess Paper Airplane And Trail Infographic', 48);
  dots(s, [[8.06, 0.296, 0.389, GRAD.orange], [0.898, 2.714, 0.524, GRAD.blue], [15.535, 10.631, 0.524, GRAD.blue],
    [1.19, 9.477, 0.447, GRAD.orange], [7.663, 4.551, 0.447, GRAD.orange]]);
};

build[22] = s => {
  cornerBlob(s, 'tr', GRAD.orange);
  // Overlapping hexagon band (drawn right to left so each notch tucks under).
  [[6.896, GRAD.blue], [4.553, GRAD.orange], [2.21, GRAD.blue]].forEach(([x, color]) =>
    s.addShape('hexagon', { x, y: 4.384, w: 3.795, h: 3.271, fill: { color }, line: NO_LINE }));
  freeform(s, [[0.03, 0, 1], [0.734, 0], [1, 0.5], [0.734, 1], [0.03, 1], [0, 0.944], [0.237, 0.5], [0, 0.056], 'z'],
    { x: 0.594, y: 4.384, w: 3.069, h: 3.271 }, { fill: { color: GRAD.orange } });
  const NODES = [
    { disc: 1.597, icon: 'people', color: GRAD.orange, dot: [0.844, 2.208, GRAD.orange], line: [1.03, 2.579, 1.576, C.orange], tx: 1.545, ty: 2.113, by: 2.712, label: 'Process One' },
    { disc: 3.974, icon: 'search', color: GRAD.blue, dot: [3.187, 9.459, GRAD.blue], line: [3.372, 7.841, 1.619, C.navy], tx: 3.969, ty: 7.96, by: 8.559, label: 'Process Two' },
    { disc: 6.35, icon: 'link', color: GRAD.orange, dot: [5.53, 2.208, GRAD.orange], line: [5.715, 2.579, 1.576, C.orange], tx: 6.205, ty: 2.113, by: 2.712, label: 'Process Three' },
    { disc: 8.727, icon: 'bulb', color: GRAD.blue, dot: [7.872, 9.459, GRAD.blue], line: [8.058, 7.841, 1.619, C.navy], tx: 8.63, ty: 7.96, by: 8.559, label: 'Process Four' },
  ];
  NODES.forEach(n => {
    s.addShape('ellipse', { x: n.dot[0], y: n.dot[1], w: 0.371, h: 0.371, fill: { color: n.dot[2] }, line: NO_LINE });
    s.addShape('line', { x: n.line[0], y: n.line[1], w: 0, h: n.line[2], line: { color: n.line[3], width: 1 } });
    s.addShape('ellipse', { x: n.disc, y: 5.349, w: 1.3, h: 1.3, fill: { color: C.white }, line: NO_LINE });
    glyph(s, n.icon, n.disc + 0.42, 5.769, 0.46, n.color, 2.5);
    tx(s, n.label, { x: n.tx, y: n.ty, w: 2.765, h: 0.505, fontFace: F.head, fontSize: 24, color: C.head85 });
    tx(s, LOREM_PROC, { x: n.tx, y: n.by, w: 3.681, h: 1.128, fontSize: 14, lineSpacingMultiple: 1.5 });
  });
  kicker(s, 12.96, 2.113);
  tx(s, [{ text: 'Proccess ', options: { breakLine: true } }, { text: 'Hexagons' }],
    { x: 12.96, y: 2.709, w: 6.164, h: 2.524, fontFace: F.head, fontSize: 72, color: C.head, lineSpacing: 86 });
  tx(s, LOREM_W, { x: 12.997, y: 5.618, w: 6.164, h: 1.672, fontSize: 16, lineSpacingMultiple: 1.5 });
  [13.032, 16.129].forEach(x => {
    tx(s, 'Section Content', { x, y: 7.938, w: 2.655, h: 0.438, fontFace: F.head, fontSize: 20, color: C.head });
    tx(s, CARD_BODY, { x, y: 8.47, w: 2.655, h: 0.865, fontSize: 16, lineSpacingMultiple: 1.5 });
  });
  dots(s, [[17.604, 10.416, 0.489, GRAD.orange], [11.587, 1.039, 0.524, GRAD.blue]]);
};

build[23] = s => {
  bigBlob(s, 14.292, 0, 5.708, 6.987, GRAD.blue);
  dots(s, [[18.318, 0.635, 0.641, GRAD.orange], [16.546, 2.788, 2.628, GRAD.orange]]);
  phoneMockup(s, 12.45, 3.75, 2.2, 5.4, 327);
  phoneMockup(s, 15.2, 1.25, 2.5, 5.3, 11);
  tx(s, '[image]', { x: 12.0, y: 9.9, w: 6.4, h: 0.4, fontSize: 16, color: 'A6ABB8', align: 'center' });
  cornerBlob(s, 'tl', GRAD.orange);
  kicker(s, 1.388, 2.13);
  heading(s, 1.388, 2.726, 7.952, 'Easy Access call & message', 60);
  bodyText(s, 1.388, 5.398, 7.952, LOREM);
  inlineFeature(s, 1.483, 7.637, GRAD.blue, 'screen');
  inlineFeature(s, 5.61, 7.641, GRAD.orange, 'nodes');
  dots(s, [[8.699, 0.955, 0.641, GRAD.orange], [10.289, 8.73, 0.524, GRAD.blue], [1.198, 9.959, 0.438, GRAD.orange]]);
};

build[24] = s => {
  rr(s, 1.141, 5.063, 17.719, 5.175, 0.582, { fill: { color: C.white }, shadow: SHADOW_CARD });
  cornerBlob(s, 'tl', GRAD.orange); cornerBlob(s, 'tr', GRAD.orange);
  laptopMockup(s, 0.623, 2.731, 7.95, 4.569);
  kicker(s, 10.492, 1.834);
  heading(s, 10.492, 2.43, 7.952, 'Prewiew Finaxy Business Financial with Laptop', 44);
  financialCols(s, 10.492, 5.806, 4.056, LOREM_COL_S, C.head85);
  bodyText(s, 10.492, 7.912, 7.952, LOREM);
  inlineFeature(s, 1.921, 8.028, GRAD.blue, 'screen');
  inlineFeature(s, 6.227, 8.032, GRAD.orange, 'nodes');
  dots(s, [[8.882, 2.537, 0.389, GRAD.orange], [1.79, 1.251, 0.524, GRAD.blue], [16.552, 0.815, 0.524, GRAD.blue],
    [17.485, 10.224, 0.389, GRAD.orange], [1.727, 10.049, 0.389, GRAD.orange], [9.968, 9.981, 0.524, GRAD.blue]]);
};

/* ------------------------- raster photos -> native shape stand-ins [image] */
const DEVICE_FRAME = '1C1C1E';

// Stand-in for the product photo of two tilted smartphones. Kept as a single
// shape so the tilt does not pull a nested screen out of alignment.
function phoneMockup(slide, x, y, w, h, rotate) {
  rr(slide, x, y, w, h, w * 0.14, { fill: { color: C.white }, rotate, shadow: SHADOW_CARD, line: { color: DEVICE_FRAME, width: 7 } });
}

// Stand-in for the open-laptop photo: bezel, screen and base.
function laptopMockup(slide, x, y, w, h) {
  rr(slide, x + w * 0.09, y, w * 0.82, h * 0.9, 0.08, { fill: { color: DEVICE_FRAME } });
  slide.addShape('rect', { x: x + w * 0.115, y: y + h * 0.045, w: w * 0.77, h: h * 0.79, fill: { color: C.white }, line: NO_LINE });
  rr(slide, x, y + h * 0.9, w, h * 0.07, h * 0.035, { fill: { color: 'C9CCD2' } });
  tx(slide, '[image]', { x: x + w * 0.115, y: y + h * 0.045, w: w * 0.77, h: h * 0.79, fontSize: 16, color: 'C3C7D0', align: 'center', valign: 'middle' });
}

/* --------------------------------------------------------------- assemble */
const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'FINAXY', width: 20, height: 11.25 });
pptx.layout = 'FINAXY';
pptx.author = 'FINAXY';
pptx.title = 'Financial Presentation Template';

for (let i = 1; i <= 25; i++) {
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  build[i](slide);
}

pptx.writeFile({ fileName: path.join(__dirname, '05024229-a47f-4378-b776-bf7912e3dcd0_grok_final.pptx') })
  .then(f => console.log('wrote', f));
