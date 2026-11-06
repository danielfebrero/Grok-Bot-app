/*
 * "Sustainable Energy Resolutions" — 20-slide deck rebuilt with pptxgenjs.
 * Raster artwork in the source deck (photo/illustration PNGs and icon glyphs)
 * is replaced here by native pptxgenjs shape placeholders in the same spot.
 *
 *   node 056b9c22-275f-4b66-aa56-df0f54541315_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const SLIDE_W = 26.6597;
const SLIDE_H = 15.0;

const FONT = 'Poppins';
const FONT_MED = 'Poppins Medium';

const C = {
  navy: '0C0F35',   // tx1 – body/heading text
  black: '000000',  // accent5
  white: 'FEFFFF',  // bg1 / accent6
  green: '9BC33D',  // accent1
  orange: 'F6B128', // accent2
  forest: '448042', // accent3
  yellow: 'FEDF20', // accent4
  pale: 'EBF3D8',   // accent1 @ 20% lum + 80% off — page background tint
  leaf: 'C3DB8B',   // accent1 @ 60% lum + 40% off
  moss: 'D7E7B1',
  sand: 'FBE0A9',   // accent2 @ 40% lum + 60% off — building windows
  cyan: '00FEFE',   // accent6 @ 50% lum — globe ocean
  ice: 'E5FFFF',    // bg1 @ 95% lum — unfilled slice of the mini pies
  road: '595959',
  shadow: '808080',
  skin: 'FBC0B9',
  skinDark: 'EAA19B',
  hair: '8E4F39',
  denim: '152A39',
  cream: 'FFF9D2',
  mustard: 'BEA718',
  darkGreen: '224021',
  brown: '7B5814',
  tan: 'A67C52',
  grey: 'A9A9A9',
  charcoal: '404040',
};

/* lorem building blocks, assembled per slide */
const L1 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.';
const L2 = ' Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.';
const L3 = ' Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.';
const L4 = ' Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.';
const LOREM_FULL = L1 + L2 + L3 + L4;
const LOREM_MED = L1 + L2 + ' ';
const LOREM_TINY = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed';
const SED_UT = 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi architecto beatae vitae dicta sunt explicabo. ';
const QUOTE = '"The switch to green energy, such as solar power, helps reduce our carbon footprint, saves money in the long run, and benefits the environment."';

/* ---------------------------------------------------------------- helpers */

const NONE = { type: 'none' };

/* Every non-title text box in the source carries spc="160" (1.6 pt tracking). */
function txt(s, text, o) {
  s.addText(text, Object.assign({ fontFace: FONT, color: C.navy, margin: 0, valign: 'top', charSpacing: 1.6 }, o));
}
/** Section/card heading — 36 pt in this deck. */
function heading(s, text, x, y, w, o) {
  txt(s, text, Object.assign({ x, y, w, h: 0.61, fontSize: 36 }, o));
}
/** Running copy — 18 pt at 1.5 line spacing unless overridden. */
function body(s, text, x, y, w, h, o) {
  txt(s, text, Object.assign({ x, y, w, h, fontSize: 18, lineSpacingMultiple: 1.5 }, o));
}
/** Slide title from the master: 80 pt, single line spacing, no tracking. */
function title(s, text, x, y, w, h, o) {
  txt(s, text, Object.assign({ x, y, w, h, fontSize: 80, lineSpacingMultiple: 1.0, charSpacing: 0 }, o));
}

function rect(s, x, y, w, h, color, o) {
  s.addShape('rect', Object.assign({ x, y, w, h, fill: { color }, line: NONE }, o));
}
function ell(s, x, y, w, h, color, o) {
  s.addShape('ellipse', Object.assign({ x, y, w, h, fill: { color }, line: NONE }, o));
}
/** Rounded rectangle; `r` is the corner radius in inches. */
function rrect(s, x, y, w, h, r, o) {
  s.addShape('roundRect', Object.assign({ x, y, w, h, rectRadius: r, line: NONE }, o));
}
/** Outlined card used all over the deck (orange or navy hairline). */
function card(s, x, y, w, h, r, color, width) {
  s.addShape('roundRect', { x, y, w, h, rectRadius: r, fill: NONE, line: { color, width } });
}
function hline(s, x, y, w, color, width) {
  s.addShape('line', { x, y, w, h: 0, line: { color, width } });
}

/** Smooth "staircase" stroke: flats joined by rises, cubic-interpolated. */
function waveRun(x0, y0, w, h) {
  const dx = w / 5;
  const dy = h / 2;
  const way = [[0, 0], [1, 0], [2, -1], [3, -1], [4, -2], [5, -2]]
    .map(([a, b]) => ({ x: x0 + a * dx, y: y0 + b * dy }));
  const pts = [{ x: way[0].x, y: way[0].y, moveTo: true }];
  for (let i = 1; i < way.length; i++) {
    const a = way[i - 1];
    const b = way[i];
    pts.push({
      x: b.x, y: b.y,
      curve: { type: 'cubic', x1: a.x + (b.x - a.x) * 0.7, y1: a.y, x2: b.x - (b.x - a.x) * 0.7, y2: b.y },
    });
  }
  return pts;
}

/**
 * The deck's recurring orange double-wave doodle (a PNG in the original).
 * (x, y) is the top-left of the original 3.15" graphic frame.
 */
function squiggle(s, x, y, rotate) {
  const size = 3.15;
  const w = 0.765 * size;
  const h = 0.606 * size;
  s.addShape('custGeom', {
    x: x + 0.134 * size, y: y + 0.197 * size, w, h, rotate,
    fill: NONE, line: { color: C.orange, width: 7 },
    points: waveRun(0, h * 0.55, w * 0.80, h * 0.52)
      .concat(waveRun(w * 0.20, h, w * 0.80, h * 0.52)),
  });
}

/** Stand-in for a raster icon: a thin outlined tile with a diagonal accent. */
function iconMark(s, x, y, w, h, color) {
  s.addShape('roundRect', { x, y, w, h, rectRadius: Math.min(w, h) * 0.18, fill: NONE, line: { color, width: 1.5 } });
  s.addShape('line', { x: x + w * 0.2, y: y + h * 0.72, w: w * 0.6, h: -h * 0.44, line: { color, width: 1.5 } });
}

/* ------------------------------------------------ illustration plumbing -- */
/*
 * Illustrations are described as flat tables of normalised parts,
 *   [kind, x, y, w, h, color]   with x/y/w/h as fractions of the frame,
 * where kind is 'e' ellipse, 'r' rectangle, 'q' rounded rectangle,
 * 't' triangle, 'd' half-dome and 'o' outlined ellipse. `figure()` scales a
 * table into place, so the same drawing can be reused at any size.
 */
function figure(s, x, y, w, h, parts) {
  parts.forEach((p) => {
    const [kind, fx, fy, fw, fh, color, rot] = p;
    const o = { x: x + fx * w, y: y + fy * h, w: fw * w, h: fh * h };
    if (rot) o.rotate = rot;
    if (kind === 'e') ell(s, o.x, o.y, o.w, o.h, color, o);
    else if (kind === 'r') rect(s, o.x, o.y, o.w, o.h, color, o);
    else if (kind === 'q') rrect(s, o.x, o.y, o.w, o.h, Math.min(o.w, o.h) * 0.28, Object.assign(o, { fill: { color } }));
    else if (kind === 't') s.addShape('triangle', Object.assign(o, { fill: { color }, line: NONE }));
    else if (kind === 'd') s.addShape('pie', Object.assign(o, { h: o.h * 2, angleRange: [180, 360], fill: { color }, line: NONE }));
    else if (kind === 'o') s.addShape('ellipse', Object.assign(o, { fill: NONE, line: { color, width: 3 } }));
  });
}

/*
 * Rider on a left-facing scooter — slides 1 and 10. Listed back to front;
 * the body shell is drawn after the rider so his legs disappear behind it.
 */
const SCOOTER = [
  ['q', 0.53, 0.10, 0.28, 0.32, C.brown],         // rider torso
  ['e', 0.58, 0.00, 0.14, 0.12, C.skin],          // head
  ['e', 0.57, -0.03, 0.16, 0.08, C.navy],         // hair
  ['q', 0.34, 0.22, 0.32, 0.10, C.skin],          // arm on the bars
  ['e', 0.23, 0.17, 0.13, 0.11, C.green],         // handlebar knob
  ['r', 0.28, 0.26, 0.03, 0.14, C.charcoal],      // steering column
  ['q', 0.40, 0.37, 0.34, 0.24, C.charcoal],      // saddle
  ['e', 0.05, 0.74, 0.26, 0.24, C.navy],          // front wheel
  ['e', 0.74, 0.68, 0.25, 0.23, C.navy],          // rear wheel
  ['q', 0.53, 0.32, 0.45, 0.34, C.leaf],          // rear hump
  ['q', 0.07, 0.36, 0.80, 0.46, C.leaf],          // body shell
  ['e', 0.02, 0.50, 0.24, 0.14, C.green],         // front fender
  ['q', 0.60, 0.47, 0.30, 0.24, C.tan],           // side panel
  ['e', 0.20, 0.53, 0.13, 0.13, C.green],         // shell trim
  ['e', 0.10, 0.80, 0.13, 0.12, C.white],         // front hub
];

/*
 * Boy in a cap holding a potted plant — slides 2 and 20.
 * Coordinates traced from the original group at 5.84 x 12.81 in.
 */
const GARDENER = [
  ['q', 0.238, 0.738, 0.690, 0.262, C.darkGreen], // trousers
  ['e', 0.456, 0.860, 0.479, 0.140, C.forest],    // shoe
  ['r', 0.491, 0.180, 0.175, 0.105, C.skin],      // neck
  ['q', 0.218, 0.244, 0.777, 0.522, C.mustard],   // shirt
  ['q', 0.378, 0.542, 0.589, 0.162, C.cream],     // sleeve cuff
  ['r', 0.262, 0.441, 0.111, 0.215, C.cream],     // placket
  ['e', 0.406, 0.376, 0.163, 0.090, C.cream],     // chest badge
  ['e', 0.298, 0.055, 0.443, 0.194, C.skin],      // face
  ['e', 0.288, 0.000, 0.437, 0.108, C.mustard],   // cap
  ['r', 0.700, 0.072, 0.140, 0.034, '7F7010'],    // cap brim
  ['e', 0.601, 0.094, 0.068, 0.038, C.denim],     // eye
  ['e', 0.199, 0.412, 0.161, 0.048, C.forest],    // seedling leaves
  ['e', 0.053, 0.412, 0.161, 0.048, C.forest],
  ['q', 0.000, 0.505, 0.384, 0.155, C.navy],      // plant pot
  ['e', 0.063, 0.604, 0.348, 0.103, C.skin],      // hands
  ['e', 0.014, 0.602, 0.154, 0.084, C.skinDark],
];

/*
 * Woman holding a plant, standing in front of the globe — slide 6.
 * Coordinates are relative to the 9.25 x 7.48 in group frame, so the
 * figure occupies only the left half of it.
 */
const PLANTER = [
  ['e', 0.156, 0.013, 0.216, 0.281, C.hair],      // hair mass
  ['e', 0.150, 0.120, 0.048, 0.078, C.skinDark],  // ears
  ['e', 0.330, 0.124, 0.048, 0.078, C.skinDark],
  ['e', 0.186, 0.083, 0.156, 0.196, C.skin],      // face
  ['e', 0.156, 0.013, 0.216, 0.100, C.hair],      // fringe
  ['e', 0.223, 0.150, 0.022, 0.038, C.denim],     // eyes
  ['e', 0.279, 0.152, 0.022, 0.036, C.denim],
  ['e', 0.240, 0.212, 0.045, 0.027, C.denim],     // mouth
  ['r', 0.238, 0.250, 0.062, 0.110, C.skin],      // neck
  ['r', 0.004, 0.468, 0.098, 0.495, C.skin],      // arms
  ['r', 0.491, 0.468, 0.076, 0.495, C.skin],
  ['q', 0.000, 0.334, 0.545, 0.666, C.orange],    // shirt
  ['e', 0.230, 0.401, 0.110, 0.055, C.forest],    // seedling leaves
  ['e', 0.150, 0.401, 0.110, 0.055, C.forest],
  ['q', 0.177, 0.446, 0.206, 0.230, C.navy],      // plant pot
  ['e', 0.134, 0.642, 0.110, 0.070, C.skin],      // hands cradling the pot
  ['e', 0.329, 0.642, 0.110, 0.070, C.skin],
];

/* Factory + house with rooftop solar panels — slide 9. */
const FACTORY = [
  ['r', 0.285, 0.026, 0.123, 0.055, C.shadow],  // chimney caps
  ['r', 0.285, 0.026, 0.123, 0.494, C.road],
  ['r', 0.432, 0.068, 0.123, 0.055, C.shadow],
  ['r', 0.432, 0.068, 0.123, 0.494, C.road],
  ['r', 0.573, 0.174, 0.424, 0.125, C.sand],    // office tower + parapet
  ['r', 0.573, 0.214, 0.424, 0.649, C.orange],
  ['r', 0.598, 0.125, 0.366, 0.143, '7F8080'],  // rooftop solar panel
  ['r', 0.623, 0.136, 0.310, 0.114, '02BBD7'],
  ['r', 0.044, 0.566, 0.656, 0.403, '74922E'],  // house body
  ['t', 0.000, 0.500, 0.462, 0.100, '262626'],  // gable
  ['r', 0.000, 0.515, 0.462, 0.195, '262626'],  // roof slab
  ['r', 0.238, 0.432, 0.471, 0.265, '262626'],
  ['r', 0.286, 0.427, 0.370, 0.227, '02BBD7'],  // house solar panel
  ['r', 0.072, 0.750, 0.131, 0.250, '262626'],  // doorways
  ['r', 0.219, 0.750, 0.131, 0.250, '262626'],
];

/* Earth-with-sprout badge — slide 3. */
const EARTH = [
  ['e', 0.00, 0.20, 0.98, 0.80, C.cyan],
  ['e', 0.03, 0.28, 0.49, 0.41, C.forest],
  ['e', 0.48, 0.27, 0.52, 0.25, C.forest],
  ['e', 0.66, 0.55, 0.30, 0.30, C.forest],
  ['e', 0.00, 0.85, 0.57, 0.11, C.forest],
  ['r', 0.63, 0.05, 0.04, 0.20, C.forest],      // stem
  ['e', 0.74, 0.00, 0.25, 0.12, C.forest],      // leaves
  ['e', 0.51, 0.00, 0.25, 0.12, C.forest],
];

/* Sprout on a mound of soil — slide 4. */
const SPROUT = [
  ['d', 0.00, 0.51, 1.00, 0.49, C.brown],
  ['r', 0.42, 0.18, 0.14, 0.42, C.forest],
  ['e', 0.44, 0.04, 0.30, 0.21, C.forest],
  ['e', 0.19, 0.00, 0.30, 0.22, C.forest],
];

/** Light bulb sprouting a seedling — slides 2, 7 and 13. */
function bulb(s, x, y, w, h) {
  figure(s, x, y, w, h, [
    ['o', 0.22, 0.16, 0.56, 0.62, C.navy],       // glass
    ['r', 0.47, 0.42, 0.06, 0.32, C.forest],     // stem
    ['e', 0.29, 0.36, 0.21, 0.11, C.forest],     // leaves
    ['e', 0.50, 0.36, 0.21, 0.11, C.forest],
    ['r', 0.32, 0.78, 0.36, 0.18, C.grey],       // screw base
  ]);
  // Radiating rays: [fx, fy, dx, dy] in frame fractions.
  [[0.02, 0.55, 0.11, -0.05], [0.87, 0.55, 0.11, 0.05], [0.06, 0.16, 0.10, 0.09],
   [0.84, 0.16, 0.10, -0.09], [0.36, 0.00, 0.04, 0.10], [0.60, 0.00, -0.04, 0.10]]
    .forEach(([fx, fy, dx, dy]) => {
      s.addShape('line', { x: x + fx * w, y: y + fy * h, w: dx * w, h: dy * h, line: { color: C.navy, width: 2 } });
    });
}

/** Globe with green land masses (slides 6 & 12). */
function globe(s, x, y, d) {
  ell(s, x, y, d, d, C.cyan);
  [[-0.01, 0.10, 0.42, 0.46], [0.10, 0.05, 0.34, 0.30], [0.52, 0.00, 0.36, 0.30],
   [0.60, 0.14, 0.34, 0.34], [0.62, 0.40, 0.38, 0.36], [0.06, 0.58, 0.40, 0.40]]
    .forEach(([fx, fy, fw, fh]) => ell(s, x + fx * d, y + fy * d, fw * d, fh * d, C.forest));
}

/**
 * Three navy arrows circling the globe. Each entry gives the arc's bounding
 * box, its sweep (degrees clockwise from 3 o'clock) and where the head sits.
 */
function orbitArrows(s, cx, cy, r) {
  const arcs = [
    { ax: cx - r * 0.34, ay: cy - r * 1.30, aw: r * 0.92, ah: r * 0.42, range: [200, 350], tx: 0.98, ty: 0.42, rot: 135 },
    { ax: cx + r * 0.92, ay: cy - r * 0.16, aw: r * 0.40, ah: r * 0.86, range: [290, 100], tx: 0.42, ty: 0.99, rot: 200 },
    { ax: cx - r * 1.32, ay: cy - r * 0.16, aw: r * 0.40, ah: r * 0.86, range: [100, 270], tx: 0.50, ty: 0.02, rot: 20 },
  ];
  arcs.forEach((a) => {
    s.addShape('arc', { x: a.ax, y: a.ay, w: a.aw, h: a.ah, fill: NONE, line: { color: C.navy, width: 4 }, angleRange: a.range });
    s.addShape('triangle', {
      x: a.ax + a.tx * a.aw - r * 0.09, y: a.ay + a.ty * a.ah - r * 0.09,
      w: r * 0.18, h: r * 0.18, fill: { color: C.navy }, line: NONE, rotate: a.rot,
    });
  });
}

/* ============================================================== slides === */

function slide01(pres) {
  const s = pres.addSlide();
  s.background = { color: C.pale };

  title(s, 'Sustainable Energy Resolutions', 1.89, 3.69, 10.17, 6.7, { fontSize: 120 });
  rrect(s, 2.01, 10.83, 5.03, 1.47, 0.735, { fill: { color: C.forest } });
  txt(s, 'GET STARTED', { x: 2.39, y: 11.25, w: 4.26, h: 0.64, fontSize: 32, bold: true, color: C.white, align: 'center' });

  // City skyline: five towers with window grids
  const towers = [
    { x: 12.45, y: 4.93, w: 2.12, h: 5.69, wx: 12.74, ww: 1.54, wh: 0.4, wy: 5.24, step: 0.708, rows: 7, cols: 1, cstep: 0 },
    { x: 15.30, y: 4.29, w: 1.31, h: 6.45, wx: 15.61, ww: 0.26, wh: 0.26, wy: 4.52, step: 0.383, rows: 16, cols: 2, cstep: 0.43 },
    { x: 17.18, y: 6.17, w: 1.57, h: 4.29, wx: 17.39, ww: 1.14, wh: 0.3, wy: 6.40, step: 0.534, rows: 7, cols: 1, cstep: 0 },
    { x: 18.93, y: 7.57, w: 2.98, h: 4.24, wx: 19.15, ww: 0.33, wh: 0.33, wy: 7.79, step: 0.494, rows: 8, cols: 4, cstep: 0.565 },
    { x: 22.14, y: 4.92, w: 2.01, h: 6.45, wx: 22.37, ww: 1.54, wh: 0.4, wy: 5.20, step: 0.708, rows: 7, cols: 1, cstep: 0 },
  ];
  towers.forEach((t) => {
    rect(s, t.x, t.y, t.w, t.h, C.orange);
    for (let r = 0; r < t.rows; r++) {
      for (let c = 0; c < t.cols; c++) {
        rect(s, t.wx + c * t.cstep, t.wy + r * t.step, t.ww, t.wh, C.sand);
      }
    }
  });

  // Clouds
  [[13.03, 4.14, 1.45, 0.57], [14.72, 3.87, 0.83, 0.29], [17.31, 3.83, 0.87, 0.31],
   [18.52, 4.51, 1.45, 0.57], [20.26, 3.87, 1.08, 0.38], [22.46, 3.68, 1.45, 0.57],
   [15.23, 5.37, 0.91, 0.32], [22.07, 5.38, 0.91, 0.32]]
    .forEach((c) => ell(s, c[0], c[1], c[2], c[3], C.white));

  rect(s, 12.01, 10.44, 12.76, 2.54, C.road);                  // roadway
  ell(s, 13.99, 10.54, 5.49, 1.49, C.shadow);
  ell(s, 20.88, 10.68, 2.91, 0.54, C.shadow);

  // Charging post
  rrect(s, 21.36, 7.60, 1.98, 2.82, 0.2, { fill: { color: C.forest } });
  rect(s, 21.67, 7.90, 1.36, 1.09, C.navy);
  s.addShape('line', { x: 21.75, y: 8.85, w: 1.2, h: -0.8, line: { color: C.white, width: 4 } });
  s.addShape('line', { x: 21.75, y: 9.15, w: 1.2, h: -0.8, line: { color: C.white, width: 4 } });
  ell(s, 21.95, 9.29, 0.83, 0.83, C.green);
  rect(s, 21.09, 10.41, 2.53, 0.44, C.navy);
  s.addShape('line', { x: 19.87, y: 9.60, w: 1.5, h: 0.85, line: { color: C.navy, width: 4 } });

  // Battery sign
  rrect(s, 20.40, 5.53, 1.00, 1.50, 0.12, { fill: { color: C.white } });
  s.addShape('roundRect', { x: 20.40, y: 5.53, w: 1.00, h: 1.50, rectRadius: 0.12, fill: NONE, line: { color: C.navy, width: 2 } });
  rect(s, 20.67, 5.40, 0.45, 0.15, C.navy);
  rect(s, 20.51, 6.36, 0.76, 0.51, C.yellow);
  s.addShape('lightningBolt', { x: 20.70, y: 5.79, w: 0.37, h: 0.66, fill: { color: C.navy }, line: NONE });

  figure(s, 14.0, 5.21, 5.55, 6.35, SCOOTER);

  squiggle(s, 25.01, 11.85);
  squiggle(s, -1.53, 2.44);
}

function slide02(pres) {
  const s = pres.addSlide();
  rect(s, 0, 0, 11.33, 15.03, C.green);
  title(s, 'Advantages of Going Green', 1.22, 1.18, 8.67, 3.27);

  const bullets = [
    ['Reduces greenhouse gas emissions and helps combat climate change.', 1.29, 10.41, 1.55],
    ['Saves money on energy bills in the long run.', 3.67, 11.53, 0.74],
    ['Encourages technological innovation and job creation in the green energy sector.', 5.25, 11.33, 1.55],
  ];
  bullets.forEach(([t, y, w, h]) => {
    txt(s, t, {
      x: 13.33, y, w, h, fontSize: 32, lineSpacingMultiple: 1.5,
      bullet: { characterCode: '2022', indent: 22.5 }, indentLevel: 0,
    });
  });

  card(s, 12.7, 8.17, 12.63, 5.33, 0.668, C.orange, 2.25);
  heading(s, 'We are green energy company', 13.33, 8.73, 9.85);
  hline(s, 13.33, 9.49, 7.09, C.yellow, 4);
  body(s, LOREM_FULL, 13.33, 9.94, 11.33, 2.67);

  // Countryside vignette in the green panel
  rect(s, 0, 5.68, 11.33, 9.31, C.pale);
  [[0.46, 6.05, 2.54, 0.77], [1.14, 7.61, 1.83, 0.55], [6.74, 5.91, 2.54, 0.77], [7.54, 7.56, 3.42, 1.04]]
    .forEach((c) => ell(s, c[0], c[1], c[2], c[3], C.white));
  ell(s, 0.01, 8.90, 4.95, 4.26, C.white);
  ell(s, 3.41, 10.75, 7.95, 4.27, C.moss);
  ell(s, -0.01, 11.84, 6.56, 3.18, C.leaf);
  ell(s, 0.0, 13.0, 5.57, 2.03, C.forest);
  ell(s, 8.27, 13.32, 3.09, 1.71, C.forest);
  bulb(s, 2.61, 6.64, 4.32, 4.68);
  figure(s, 6.71, 6.49, 4.64, 8.53, GARDENER);

  squiggle(s, 9.88, -1.15);
}

function slide03(pres) {
  const s = pres.addSlide();
  rect(s, 12.66, 6.83, 14.0, 8.17, C.green);
  title(s, 'About Our\nCompany', 13.92, 2.53, 6.74, 3.5, { fontFace: FONT_MED });
  heading(s, 'We are green energy company', 14.01, 8.31, 9.96, { h: 0.83 });
  body(s, LOREM_FULL, 14.01, 9.8, 10.66, 3.14);

  figure(s, 23.06, 4.72, 2.27, 2.78, EARTH);

  squiggle(s, 10.63, 11.5);
}

function slide04(pres) {
  const s = pres.addSlide();
  s.background = { color: C.pale };
  title(s, 'Why Choose Us', 1.22, 1.18, 9.03, 1.86, { color: C.black });

  const rows = [
    ['Eco-friendly Materials', LOREM_MED, 4.03, 4.19],
    ['High-quality, Durable Materials', LOREM_MED, 7.36, 7.54],
    ['Eco-friendly Form And Function',
      'Lorem Ipsum Dolor Sit Amet, Consectetur Adipiscing Elit, Sed Do Eiusmod Tempor Incididunt Ut Labore Et Dolore Magna Aliqua. Ut Enim Ad Minim Veniam, Quis Nostrud Exercitation Ullamco Laboris Nisi Ut Aliquip Ex Ea Commodo Consequat. ',
      10.71, 10.87],
  ];
  rows.forEach(([h, b, ty, iy], i) => {
    rrect(s, 1.33, iy, 2.0, 1.99, 0.33, { fill: { color: C.yellow } });
    iconMark(s, 1.82, iy + 0.5, 1.02, 1.0, C.black);
    heading(s, h, 4.0, ty, 9.03, { color: C.black });
    body(s, b, 4.0, ty + 0.81, i === 2 ? 10.82 : 10.67, 1.78, { color: C.black });
  });

  figure(s, 15.63, 2.15, 1.78, 2.14, SPROUT);

  squiggle(s, 23.5, 11.01);
}

function slide05(pres) {
  const s = pres.addSlide();
  title(s, 'Our Services', 1.45, 1.18, 8.11, 1.65);

  const services = [
    ['Service 01', 2.0, 4.20, 5.33, 4.05],
    ['Service 02', 14.0, 4.20, 17.33, 4.05],
    ['Service 03', 2.0, 9.51, 5.33, 9.36],
    ['Service 04', 14.0, 9.51, 17.33, 9.36],
  ];
  services.forEach(([label, ix, iy, tx, ty]) => {
    rrect(s, ix, iy, 2.67, 2.67, 0.445, { fill: { color: C.yellow } });
    iconMark(s, ix + 0.7, iy + 0.7, 1.27, 1.27, C.navy);
    heading(s, label, tx, ty, 4.69);
    body(s, LOREM_MED, tx, ty + 0.92, 7.33, 2.21);
  });

  squiggle(s, 24.66, 11.85);
  squiggle(s, -1.57, -0.32);
}

function slide06(pres) {
  const s = pres.addSlide();
  title(s, 'Our Process', 9.44, 2.5, 5.23, 2.75, { align: 'right' });

  [['Step 01', 1.99, 2.87], ['Step 02', 8.66, 6.88], ['Step 03', 1.99, 9.50]].forEach(([label, x, y]) => {
    card(s, x, y, 5.33, 4.0, 0.667, C.orange, 2.25);
    heading(s, label, x + 0.57, y + 0.67, 3.09);
    body(s, L1 + '  ', x + 0.57, y + 1.58, 4.35, 2.23);
  });

  globe(s, 17.07, 2.14, 6.55);
  orbitArrows(s, 20.35, 5.42, 3.3);
  rect(s, 18.34, 10.71, 3.64, 4.29, '262626');                 // trousers, cropped by the slide edge
  figure(s, 17.55, 3.59, 9.25, 7.48, PLANTER);

  squiggle(s, -1.03, 0.39);
}

function slide07(pres) {
  const s = pres.addSlide();
  rect(s, 0, 6.17, 26.66, 8.86, C.green);
  title(s, 'Meet Our Engineer', 1.22, 10.5, 8.67, 3.5, { fontSize: 96, color: C.black });

  heading(s, 'Tomas Kozak', 16.66, 2.73, 4.98, { color: C.black });
  hline(s, 16.68, 3.34, 3.17, C.yellow, 4);
  txt(s, 'Green Energy Engineer', { x: 16.68, y: 3.53, w: 4.33, h: 0.35, fontSize: 21, color: C.black });

  heading(s, 'Caption Here', 16.66, 6.83, 5.46, { color: C.black });
  body(s, LOREM_FULL + '\n\n' + SED_UT, 16.66, 7.75, 8.67, 5.87, { color: C.black });

  bulb(s, 23.04, 4.54, 2.19, 2.37);

  squiggle(s, 5.98, 0.74);
}

function slide08(pres) {
  const s = pres.addSlide();
  title(s, 'Our Team', 9.0, 6.52, 8.67, 1.65, { align: 'center' });
  txt(s, '05', { x: 23.5, y: 0.56, w: 2.0, h: 1.11, fontSize: 60, align: 'right' });

  const team = [
    { name: 'Hendrick Larsonn', role: 'Chief Executive Officer', x: 2.0, y: 2.02, align: 'left', lx: 2.0, lw: 4.10 },
    { name: 'Claudio Olmo', role: 'Chief Executive Marketing', x: 19.7, y: 2.02, align: 'right', lx: 21.47, lw: 3.21 },
    { name: 'Gregg Muller', role: 'Researcher', x: 1.98, y: 9.43, align: 'left', lx: 2.0, lw: 3.05 },
    { name: 'Lisa Malcolm', role: 'Engineer', x: 19.7, y: 9.43, align: 'right', lx: 21.47, lw: 3.21 },
  ];
  team.forEach((m) => {
    const right = m.align === 'right';
    heading(s, m.name, m.x, m.y, 4.98, { align: m.align });
    hline(s, m.lx, m.y + 0.76, m.lw, C.yellow, 4);
    txt(s, m.role, { x: right ? 20.34 : 2.0, y: m.y + 0.85, w: 4.33, h: 0.3, fontSize: 18, italic: true, align: m.align });
    body(s, L1 + ' ', right ? 20.31 : 2.0, m.y + (right ? 1.70 : 1.55), 4.37, 2.23, { align: m.align });
  });

  squiggle(s, -1.15, 6.32);
}

function slide09(pres) {
  const s = pres.addSlide();
  title(s, 'Our Vision\n& Mission', 1.22, 1.18, 6.56, 3.5);
  rrect(s, 10.65, 3.5, 7.09, 9.67, 0.396, { fill: { color: C.pale } });

  heading(s, 'Our vision', 18.66, 3.4, 5.97);
  body(s, LOREM_FULL, 18.66, 4.44, 6.67, 4.96);
  heading(s, 'Our Mission', 1.32, 7.5, 4.62);
  body(s, L1 + L2 + L3, 1.33, 8.54, 6.56, 3.6);

  figure(s, 8.93, 4.13, 8.19, 8.67, FACTORY);
  for (let r = 0; r < 6; r++) rect(s, 13.80, 6.69 + r * 0.67, 1.46, 0.73, C.sand);   // tower windows
  for (let r = 0; r < 5; r++) rect(s, 15.86, 6.33 + r * 0.82, 1.09, 1.05, C.sand);

  squiggle(s, 15.89, 11.24);
}

function slide10(pres) {
  const s = pres.addSlide();
  s.background = { color: C.pale };
  title(s, 'Our Recent Project', 1.22, 1.18, 6.78, 1.32, { color: C.black });
  rrect(s, 10.66, 2.17, 13.33, 4.19, 0.698, { fill: { color: C.yellow } });
  body(s, LOREM_FULL, 11.33, 2.68, 12.56, 2.99, { fontSize: 20, color: C.black });

  heading(s, 'Project  A - Denver', 1.46, 13.05, 7.2, { color: C.black });
  [['Project  B - Mumbai', 10.0], ['Project  C - Jakarta', 15.33], ['Project  D - Tokyo', 20.66]]
    .forEach(([t, x]) => txt(s, t, { x, y: 13.12, w: 4.67, h: 0.47, fontSize: 28, align: 'center', color: C.black }));

  ell(s, 21.19, 8.76, 3.58, 0.97, C.shadow);
  figure(s, 21.19, 5.28, 3.63, 4.45, SCOOTER);
  squiggle(s, 9.18, 0.67);
}

function slide11(pres) {
  const s = pres.addSlide();
  rect(s, 12.66, 0, 14.0, 15.03, C.green);
  title(s, 'What They Say?', 1.22, 1.18, 6.56, 3.09);

  const quotes = [
    { x: 1.34, qy: 5.33, qw: 7.99, qh: 1.47, size: 20, name: 'Ami Takahasi', ny: 7.34, lw: 3.08, lx: 1.32, color: C.navy },
    { x: 1.35, qy: 10.01, qw: 7.99, qh: 1.47, size: 20, name: 'Nuno Mendes', ny: 12.03, lw: 3.25, lx: 1.34, color: C.navy },
    { x: 15.34, qy: 8.86, qw: 7.35, qh: 2.37, size: 24, name: 'Rodrigo Santos', ny: 12.07, lw: 3.61, lx: 15.33, color: C.black },
  ];
  quotes.forEach((q) => {
    body(s, QUOTE, q.x, q.qy, q.qw, q.qh, { fontSize: q.size, color: q.color });
    heading(s, q.name, q.x - 0.01, q.ny, 4.98, { color: q.color });
    hline(s, q.lx, q.ny + 0.75, q.lw, C.yellow, 4);
    txt(s, 'Company Name', { x: q.x, y: q.ny + 0.92, w: 4.33, h: 0.35, fontSize: 21, italic: true, color: q.color });
  });
}

function slide12(pres) {
  const s = pres.addSlide();
  title(s, 'Types of Green Energy Sources', 1.22, 1.18, 9.94, 3.5);

  const sources = [
    ['Wind Energy', 'Generated by wind turbines that convert kinetic energy into electricity.', 1.36, 6.87, 4.0, 6.95],
    ['Solar Energy', 'Harnessed by using photovoltaic panels that convert sunlight into electricity.', 1.36, 10.87, 4.0, 10.95],
    ['Hydro Energy', 'Produced by using the power of moving water to turn turbines and generate electricity.', 11.37, 10.84, 13.99, 10.73],
  ];
  sources.forEach(([h, b, ix, iy, tx, ty]) => {
    rrect(s, ix, iy, 2.0, 1.99, 0.33, { fill: { color: C.yellow } });
    iconMark(s, ix + 0.5, iy + 0.5, 1.0, 1.0, C.navy);
    heading(s, h, tx, ty, 4.69);
    body(s, b, tx, ty + 0.92, 5.31, 1.33);
  });

  globe(s, 16.53, 2.95, 6.55);
  orbitArrows(s, 19.81, 6.22, 3.3);
  squiggle(s, 24.26, 11.25);
}

function slide13(pres) {
  const s = pres.addSlide();
  title(s, 'Solar Energy', 1.22, 1.18, 12.0, 1.32);

  rrect(s, 6.66, 9.5, 9.44, 3.33, 0.555, { fill: { color: C.green } });
  iconMark(s, 7.34, 10.11, 1.66, 1.66, C.navy);
  body(s, 'Solar energy is a renewable source of power that harnesses energy from the sun. Solar panels convert sunlight into electricity, reducing reliance on fossil fuels.',
    9.68, 9.95, 5.93, 2.48, { fontSize: 20 });

  [['89%', 3.5, 3.82, 5.32], ['76%', 8.83, 9.15, 10.65]].forEach(([pct, cy, ny, by]) => {
    card(s, 16.7, cy, 8.0, 4.0, 0.667, C.orange, 2.25);
    txt(s, pct, { x: 17.29, y: ny, w: 5.47, h: 1.35, fontSize: 80, lineSpacingMultiple: 1.0 });
    body(s, L1, 17.43, by, 6.67, 1.98, { fontSize: 20 });
  });

  bulb(s, 23.67, 7.86, 1.98, 2.15);

  squiggle(s, 11.09, 2.24, 60);
}

function slide14(pres) {
  const s = pres.addSlide();
  title(s, 'Clustered Chart', 5.04, 1.18, 16.58, 1.66, { align: 'center' });

  const cats = ['Category 1', 'Category 2', 'Category 3', 'Category 4'];
  s.addChart(pres.ChartType.bar, [
    { name: 'Series 1', labels: cats, values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels: cats, values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels: cats, values: [2, 2, 3, 5] },
  ], {
    x: 9.33, y: 3.41, w: 15.61, h: 10.41,
    barDir: 'col', barGapWidthPct: 219, barOverlapPct: -27,
    chartColors: [C.forest, C.green, C.yellow],
    showLegend: true, legendPos: 'b', legendFontFace: FONT, legendFontSize: 24, legendColor: C.navy,
    catAxisLabelFontFace: FONT, catAxisLabelFontSize: 24, catAxisLabelColor: C.navy,
    valAxisLabelFontFace: FONT, valAxisLabelFontSize: 24, valAxisLabelColor: C.navy,
    catAxisLineColor: C.orange, catAxisLineSize: 1, valAxisLineShow: false,
    catAxisMajorTickMark: 'none', valAxisMajorTickMark: 'none',
    catAxisLabelPos: 'nextTo',
    valGridLine: { color: 'D6D7DC', style: 'solid', size: 1 },
    plotArea: { fill: { color: C.white } },
  });

  [[4.20, 5.18, C.forest], [7.54, 8.52, C.green], [10.87, 11.85, C.yellow]].forEach(([cy, oy, color]) => {
    card(s, 2.04, cy, 6.66, 2.63, 0.439, C.orange, 2.25);
    ell(s, 2.56, oy, 0.67, 0.67, color);
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, ', 3.76, oy - 0.19, 4.46, 0.97, { fontSize: 20 });
  });

  squiggle(s, 24.42, 11.59);
  squiggle(s, -1.11, 1.54);
}

function slide15(pres) {
  const s = pres.addSlide();
  rect(s, 0, 0, 11.33, 15.0, C.green);
  title(s, 'Wind Energy', 1.22, 1.18, 12.0, 1.32);
  body(s, 'Wind energy is a renewable, clean source of power that can reduce carbon emissions and help combat climate change. Wind turbines harness the power of wind to generate electricity.',
    12.91, 2.52, 10.59, 1.47, { fontSize: 20 });

  heading(s, 'Caption Here', 1.35, 4.69, 6.37);
  body(s, LOREM_FULL + '\n\n' + SED_UT +
    'Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut odit aut fugit, sed quia consequuntur magni dolores eos qui ratione voluptatem sequi nesciunt.',
    1.35, 5.74, 8.86, 7.23);

  // Wind turbine: slim tapered mast with three swept blades
  s.addShape('triangle', { x: 10.85, y: 2.18, w: 0.72, h: 3.32, fill: { color: '7F7F7F' }, line: NONE });
  s.addShape('triangle', { x: 10.87, y: 3.79, w: 0.68, h: 1.71, fill: { color: C.charcoal }, line: NONE });
  rect(s, 10.64, 5.16, 1.10, 0.34, C.charcoal);                // foot
  [[11.14, 0.83, 0.24, 1.40, 8], [11.30, 2.14, 0.92, 0.28, -18], [10.20, 1.92, 0.98, 0.30, 197]]
    .forEach(([bx, by, bw, bh, rot]) =>
      s.addShape('triangle', { x: bx, y: by, w: bw, h: bh, fill: { color: 'CBCBCB' }, line: NONE, rotate: rot }));
  ell(s, 11.10, 2.10, 0.22, 0.22, C.charcoal);                 // hub

  squiggle(s, 23.49, 11.57);
}

function slide16(pres) {
  const s = pres.addSlide();
  title(s, 'Pie Chart Data', 1.22, 1.18, 7.25, 1.85);
  body(s, L1, 1.31, 4.18, 7.25, 1.47, { fontSize: 20 });

  card(s, 1.31, 6.83, 6.69, 6.67, 0.5, C.orange, 2.5);
  [[7.38, C.green], [8.88, C.forest], [10.54, C.orange], [12.06, C.yellow]].forEach(([y, color], i) => {
    ell(s, 1.94, y, 0.95, 0.95, color);
    body(s, 'Lorem ipsum amet', 3.28, y + 0.09, 4.46, 0.56, { fontSize: 24 });
  });

  card(s, 9.33, 1.5, 6.67, 12.0, 0.5, C.orange, 2.5);
  const mini = [['10%', 2.13, 2.17, C.yellow], ['20%', 4.89, 4.86, C.forest], ['30%', 7.82, 7.68, C.orange], ['40%', 10.52, 10.38, C.green]];
  mini.forEach(([pct, ty, cy, color], i) => {
    txt(s, pct, { x: 9.85, y: ty, w: 2.08, h: 0.83, fontSize: 40 });
    body(s, 'Nam aliquam euismod dui et tincidunt. ', 9.97, ty + 0.75, 1.86, 1.78);
    s.addChart(pres.ChartType.pie, [{
      name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: [(i + 1) * 10, 100 - (i + 1) * 10],
    }], { x: 11.7, y: cy, w: 3.53, h: 2.35, chartColors: [color, C.ice], showLegend: false, dataBorder: { pt: 1.5, color: 'FFFFFF' } });
  });

  s.addChart(pres.ChartType.doughnut, [{
    name: 'Sales', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'], values: [40, 30, 20, 10],
  }], {
    x: 14.74, y: 2.96, w: 12.89, h: 8.6, holeSize: 63, showLegend: false,
    chartColors: [C.green, C.orange, C.forest, C.yellow],
  });
  ell(s, 18.99, 5.06, 4.41, 4.41, C.white);

  // Leader lines from the mini pies to the doughnut
  [[3.54, 4.66], [5.63, 3.46], [9.04, 3.46], [11.91, 8.78]].forEach(([y, w]) => hline(s, 15.43, y, w, C.green, 1.75));
  s.addShape('line', { x: 24.21, y: 10.03, w: 0, h: 1.88, line: { color: C.green, width: 1.75 } });

  squiggle(s, 24.42, 11.59);
  squiggle(s, -1.45, -0.14);
}

function slide17(pres) {
  const s = pres.addSlide();
  rect(s, 13.97, 0, 12.69, 15.0, C.pale);
  title(s, 'Your Text Here', 14.86, 1.82, 8.82, 1.61);

  const steps = [
    ['RESEARCH', 4.78, 5.69, C.orange, 4.85, 5.16, '1', C.white],
    ['ANALYZE', 7.91, 8.67, C.forest, 7.96, 8.27, '2', C.white],
    ['BUSINESS', 10.93, 11.71, C.yellow, 11.01, 11.32, '3', C.navy],
  ];
  steps.forEach(([label, ly, by, color, oy, ny, num, numColor]) => {
    ell(s, 14.96, oy, 1.70, 1.69, color);
    txt(s, num, { x: 14.96, y: ny, w: 1.70, h: 1.01, fontSize: 60, align: 'center', color: numColor });
    txt(s, label, { x: 17.31, y: ly, w: 3.72, h: 0.61, fontSize: 36 });
    body(s, LOREM_TINY, 17.31, by, 5.56, 0.87);
  });

  // Three-segment process ring; angles run clockwise from 3 o'clock
  const RX = 1.33;
  const RY = 2.14;
  const RD = 11.36;
  [[C.forest, 272, 26], [C.yellow, 34, 146], [C.orange, 154, 266]].forEach(([color, a1, a2]) => {
    s.addShape('pie', { x: RX, y: RY, w: RD, h: RD, angleRange: [a1, a2], fill: { color }, line: NONE });
  });
  ell(s, RX + RD * 0.29, RY + RD * 0.29, RD * 0.42, RD * 0.42, C.white);
  iconMark(s, 2.68, 4.76, 1.68, 1.69, C.navy);                 // pie-chart icon
  iconMark(s, 10.06, 5.39, 1.76, 1.76, C.white);               // gear icon
  iconMark(s, 5.96, 10.97, 1.88, 1.68, C.navy);                // briefcase icon

  squiggle(s, 24.42, 11.59);
  squiggle(s, -1.11, 1.54);
}

function slide18(pres) {
  const s = pres.addSlide();
  title(s, 'Your Text Here', 5.04, 1.18, 16.58, 1.66, { align: 'center' });

  const cols = [
    { year: '1987', label: 'Option 1', chevX: 2.48, chevW: 5.72, color: C.green, yearX: 3.48, yearW: 3.24, cardX: 2.72, ovalX: 4.32, ovalY: 8.03, labelX: 3.17, labelY: 9.96, bodyX: 3.37, bodyY: 10.75, cardY: 7.53 },
    { year: '1998', label: 'Option 2', chevX: 7.90, chevW: 5.68, color: C.forest, yearX: 8.91, yearW: 3.18, cardX: 8.13, ovalX: 9.73, ovalY: 8.01, labelX: 8.58, labelY: 9.92, bodyX: 8.78, bodyY: 10.71, cardY: 7.53 },
    { year: '2005', label: 'Option 3', chevX: 13.29, chevW: 5.68, color: C.yellow, yearX: 14.34, yearW: 3.23, cardX: 13.59, ovalX: 15.19, ovalY: 7.97, labelX: 14.05, labelY: 9.95, bodyX: 14.24, bodyY: 10.74, cardY: 7.52 },
    { year: '2017', label: 'Option 1', chevX: 18.68, chevW: 5.68, color: C.orange, yearX: 19.89, yearW: 3.27, cardX: 19.02, ovalX: 20.62, ovalY: 8.08, labelX: 19.47, labelY: 9.95, bodyX: 19.67, bodyY: 10.74, cardY: 7.52 },
  ];
  cols.forEach((c, i) => {
    s.addShape(i === 0 ? 'homePlate' : 'chevron', {
      x: c.chevX, y: 5.53, w: c.chevW, h: 1.31, fill: { color: c.color }, line: NONE,
    });
    txt(s, c.year, { x: c.yearX, y: 3.96, w: c.yearW, h: 1.01, fontSize: 60, align: 'center' });
    card(s, c.cardX, c.cardY, 4.64, 5.30, 0.664, C.navy, 2.25);
    s.addShape('ellipse', { x: c.ovalX, y: c.ovalY, w: 1.43, h: 1.44, fill: NONE, line: { color: C.navy, width: 2.63 } });
    iconMark(s, c.ovalX + 0.32, c.ovalY + 0.33, 0.80, 0.78, C.navy);
    txt(s, c.label, { x: c.labelX, y: c.labelY, w: 3.72, h: 0.61, fontSize: 36, align: 'center' });
    body(s, LOREM_TINY, c.bodyX, c.bodyY, 3.33, 1.32, { align: 'center' });
  });

  squiggle(s, 24.42, 11.59);
  squiggle(s, -1.11, 1.54);
}

function slide19(pres) {
  const s = pres.addSlide();
  title(s, 'Weekly Task Timeline', 5.04, 1.18, 16.58, 1.66, { align: 'center' });

  // ring x, ring y, ring colour, caption y, arrow y
  const weeks = [
    { n: 'Week\n1', rx: 2.69, ry: 4.36, color: C.yellow, capY: 10.74, capX: 2.72, txtX: 2.48, txtW: 3.86, arrowX: 4.41, arrowY: 8.22 },
    { n: 'Week 4', rx: 7.16, ry: 9.64, color: C.orange, capY: 4.39, capX: 7.19, txtX: 6.99, txtW: 3.76, arrowX: 8.88, arrowY: 6.95 },
    { n: 'Week 2', rx: 11.62, ry: 4.36, color: C.forest, capY: 10.74, capX: 11.66, txtX: 11.46, txtW: 3.76, arrowX: 13.34, arrowY: 8.22 },
    { n: 'Week 5', rx: 16.09, ry: 9.64, color: C.yellow, capY: 4.39, capX: 16.12, txtX: 15.93, txtW: 3.76, arrowX: 17.81, arrowY: 6.95 },
    { n: 'Week 3', rx: 20.56, ry: 4.36, color: C.orange, capY: 10.74, capX: 20.59, txtX: 20.39, txtW: 3.76, arrowX: 22.28, arrowY: 8.22 },
  ];
  weeks.forEach((w) => {
    ell(s, w.rx, w.ry, 3.44, 3.44, w.color);
    ell(s, w.rx + 0.64, w.ry + 0.64, 2.16, 2.16, C.white);
    s.addShape('ellipse', { x: w.rx + 0.79, y: w.ry + 0.79, w: 1.85, h: 1.85, fill: NONE, line: { color: C.orange, width: 3 } });
    txt(s, w.n, { x: w.rx + 0.79, y: w.ry + 1.28, w: 1.85, h: 0.91, fontSize: 24, align: 'center' });
    s.addShape('line', {
      x: w.arrowX, y: w.arrowY, w: 0, h: 2.13,
      line: { color: C.orange, width: 2.25, beginArrowType: 'arrow', endArrowType: 'arrow' },
    });
    txt(s, 'Text Here', { x: w.capX, y: w.capY, w: 3.37, h: 0.61, fontSize: 36, align: 'center', lineSpacingMultiple: 1.0 });
    body(s, 'Lorem ipsum dolor sit amet, consectetur elit, sed do eiusmod ut. ', w.txtX, w.capY + 0.73, w.txtW, 1.33, { align: 'center' });
  });

  squiggle(s, 24.42, 11.59);
  squiggle(s, -1.11, 1.54);
}

function slide20(pres) {
  const s = pres.addSlide();
  s.background = { color: C.pale };
  txt(s, 'Daun Hijau Company', { x: 1.23, y: 0.68, w: 5.37, h: 0.64, fontSize: 32, color: C.black });
  txt(s, '2025 Presentation Template', { x: 17.33, y: 0.7, w: 8.06, h: 0.71, fontSize: 36, align: 'right', color: C.black });

  title(s, 'Get in Touch', 2.56, 3.12, 8.93, 1.59, { fontSize: 96 });
  hline(s, 2.7, 5.11, 6.06, C.orange, 4);

  const contacts = [
    ['1234, Name Street, Building, Your City, Your Country.', 6.46, 2.67, 6.65, 0.70, 0.64, 5.17],
    ['+123 4567 890\n+123 4567', 8.57, 2.67, 8.92, 0.70, 0.70, 5.17],
    ['yourinfo@email.com \nyour.email@mail.com', 10.68, 2.68, 11.15, 0.67, 0.41, 6.01],
  ];
  contacts.forEach(([t, ty, ix, iy, iw, ih, tw]) => {
    iconMark(s, ix, iy, iw, ih, C.navy);
    body(s, t, 4.02, ty, tw, 1.16, { fontSize: 24 });
  });

  figure(s, 15.07, 2.21, 5.84, 12.81, GARDENER);
  // Bushes along the bottom edge
  [[-1.0, 12.19, 8.79], [21.33, 12.08, 6.33]].forEach(([bx, by, bw]) => ell(s, bx, by, bw, 3.5, C.forest));
  [[0.4, 12.9, 2.4], [3.2, 12.6, 2.6], [22.2, 12.7, 2.4]].forEach(([bx, by, bw]) => ell(s, bx, by, bw, 2.2, C.forest));

  squiggle(s, -1.39, 2.21);
}

/* ------------------------------------------------------------------ build */

function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'DECK', width: SLIDE_W, height: SLIDE_H });
  pres.layout = 'DECK';
  pres.author = 'Daun Hijau Company';
  pres.title = 'Sustainable Energy Resolutions';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
   slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach((fn) => fn(pres));

  const out = path.join(__dirname, '056b9c22-275f-4b66-aa56-df0f54541315_grok_final.pptx');
  return pres.writeFile({ fileName: out }).then(() => console.log('wrote', out));
}

build().catch((err) => { console.error(err); process.exit(1); });
