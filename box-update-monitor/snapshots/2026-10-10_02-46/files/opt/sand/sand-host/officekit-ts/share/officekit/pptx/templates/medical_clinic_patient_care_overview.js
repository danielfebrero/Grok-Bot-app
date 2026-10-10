/**
 * "Medical Clinic" healthcare deck — rebuilt with pptxgenjs.
 *
 * Run:  node 10ab7517-2c79-4d71-a6ea-ed7429e103c6_grok_final.js
 * Out:  10ab7517-2c79-4d71-a6ea-ed7429e103c6_grok_final.pptx  (next to this file)
 *
 * Photographs in the source deck are stand-in grey boxes; they are reproduced
 * here as grey placeholder shapes labelled "[image]". Vector icons are redrawn
 * from native PowerPoint shapes.
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const C = {
  navy: '214FA4',   // accent1
  teal: '01E4B7',   // accent2
  blue: '357EFE',   // accent3
  deep: '002060',   // accent6
  band: '0099FF',   // decorative sweep on the title / closing slides
  ink: '262626',    // body copy (tx1, 85% luminance)
  page: '2E2E2E',   // slide-number grey
  white: 'FFFFFF',
  ph: 'CCCCCC',     // image placeholder fill
  phInk: 'DDDDDD',  // image placeholder caption
  grid: 'D9D9D9',
  tick: 'A5A5A5',
  softBlue: 'D7E5FF', // accent3 @ 20% lum
  midBlue: 'AECBFF',  // accent3 @ 40% lum
  darkBlue: '013899', // accent3 @ 50% lum
  richBlue: '0154E5', // accent3 @ 75% lum
  pin: 'E8E8E8',
  track: 'E6E6E6',
};

const HEAD = 'Bebas Neue'; // major latin font
const BODY = 'Open Sans';  // minor latin font

const NONE = { type: 'none' };
/** OOXML roundRect "adj" value -> pptxgenjs rectRadius (inches). */
const rr = (adj, minSide) => (adj / 100000) * minSide;
/**
 * Shadow presets. pptxgenjs rewrites the object it is handed, so each shape
 * needs its own copy — hence factory functions rather than shared literals.
 */
const CARD_SHADOW = () => ({ type: 'outer', blur: 20, offset: 0.5, angle: 90, color: '000000', opacity: 0.1 });
const LIFT_SHADOW = () => ({ type: 'outer', blur: 41, offset: 1, angle: 90, color: '000000', opacity: 0.1 });
const PIN_SHADOW = () => ({ type: 'outer', blur: 17, offset: 7, angle: 90, color: '000000', opacity: 0.13 });

/* --------------------------------------------------------- shared strings */

const LOREM_A = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nost';
const LOREM_B = 'exercitation ullamco laboris nisi ut aliquip ex ea.';
const LOREM_MED = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore';
const LOREM_ADIC = 'Lorem ipsum dolor sit amet, consectetur adic adipiscing elit.';
const EAQUE = 'PLACEHOLDER';

/* ------------------------------------------------------------- primitives */

/** Text box. Defaults mirror PowerPoint: top-anchored, left-aligned. */
function text(slide, runs, o) {
  slide.addText(runs, Object.assign({ fontFace: BODY, color: C.ink, valign: 'top' }, o));
}

/** Solid shape helper. */
function shape(slide, kind, o) {
  slide.addShape(kind, Object.assign({ line: NONE }, o));
}

/** Grey stand-in for a photograph, with a faint "[image]" caption. */
function imagePlaceholder(slide, o) {
  const { x, y, w, h, kind = 'roundRect', radius = 0, label = true } = o;
  shape(slide, kind, { x, y, w, h, fill: { color: C.ph }, rectRadius: radius });
  if (label && w > 1.1 && h > 0.6) {
    text(slide, '[image]', {
      x, y: y + h / 2 - 0.2, w, h: 0.4, align: 'center', valign: 'middle',
      fontSize: 12, color: C.phInk,
    });
  }
}

/**
 * Closed polygon that fills from a sampled curve out to one side of the slide.
 * `curve` is a list of [y, x] samples, top to bottom.
 */
function sweep(slide, curve, color, side) {
  const edgeX = side === 'left' ? 0 : 13.3333;
  const pts = curve.map(([y, x]) => ({ x, y }));
  pts.push({ x: edgeX, y: curve[curve.length - 1][0] });
  pts.push({ x: edgeX, y: curve[0][0] });
  pts.push({ close: true });
  shape(slide, 'custGeom', { x: 0, y: 0, w: 13.3333, h: 7.5, points: pts, fill: { color } });
}

/** Freeform from normalised path commands: ['M',x,y] ['L',x,y] ['C',x1,y1,x2,y2,x,y] ['Z']. */
function freeform(slide, cmds, o) {
  const { x, y, w, h } = o;
  const pts = cmds.map((c) => {
    if (c[0] === 'Z') return { close: true };
    if (c[0] === 'M') return { x: c[1] * w, y: c[2] * h, moveTo: true };
    if (c[0] === 'L') return { x: c[1] * w, y: c[2] * h };
    return { x: c[5] * w, y: c[6] * h, curve: { type: 'cubic', x1: c[1] * w, y1: c[2] * h, x2: c[3] * w, y2: c[4] * h } };
  });
  slide.addShape('custGeom', Object.assign({ points: pts, line: NONE }, o));
}

/** Straight polyline through [x,y] points (used for chart grid lines / series). */
function polyline(slide, points, o) {
  const x0 = Math.min(...points.map((p) => p[0]));
  const y0 = Math.min(...points.map((p) => p[1]));
  const w = Math.max(0.01, Math.max(...points.map((p) => p[0])) - x0);
  const h = Math.max(0.01, Math.max(...points.map((p) => p[1])) - y0);
  slide.addShape('custGeom', Object.assign({
    x: x0, y: y0, w, h, fill: NONE,
    points: points.map((p, i) => ({ x: p[0] - x0, y: p[1] - y0, moveTo: i === 0 })),
  }, o));
}

/* ------------------------------------------------------------------ icons */

/**
 * Flat medical icons, each assembled from a couple of built-in shapes inside a
 * unit box. `hole` is the colour showing through cut-outs (the plate behind).
 */
function icon(slide, name, x, y, s, color, hole) {
  const put = (kind, dx, dy, dw, dh, o) => slide.addShape(kind, Object.assign(
    { x: x + dx * s, y: y + dy * s, w: dw * s, h: dh * s, fill: { color }, line: NONE }, o));
  const cut = { fill: { color: hole } };
  /** Ring drawn as filled disc + punched-out disc (gives exact stroke width). */
  const ring = (dx, dy, d, t, c) => {
    put('ellipse', dx, dy, d, d, c ? { fill: { color: c } } : undefined);
    put('ellipse', dx + t, dy + t, d - 2 * t, d - 2 * t, cut);
  };
  /** Greek cross from two bars. */
  const cross = (dx, dy, d, t, o) => {
    put('rect', dx, dy + (d - t) / 2, d, t, o);
    put('rect', dx + (d - t) / 2, dy, t, d, o);
  };
  switch (name) {
    case 'crossRing':                                             // outlined cross badge
      ring(0, 0, 1, 0.085);
      cross(0.235, 0.235, 0.53, 0.2);
      break;
    case 'crossSolid':                                            // logo mark
      ring(0, 0, 1, 0.1);
      cross(0.19, 0.19, 0.62, 0.24);
      break;
    case 'search':                                                // magnifier
      ring(0, 0, 0.78, 0.1);
      put('rect', 0.58, 0.66, 0.4, 0.13, { rotate: 45 });
      break;
    case 'medbag':                                                // doctor's bag
      put('rect', 0.36, 0.06, 0.28, 0.14);
      put('rect', 0.42, 0.11, 0.16, 0.11, cut);
      put('roundRect', 0.08, 0.2, 0.84, 0.64, { rectRadius: 0.07 * s });
      put('star6', 0.33, 0.34, 0.34, 0.36, cut);
      break;
    case 'medkit':                                                // flat first-aid kit
      put('rect', 0.36, 0.08, 0.28, 0.14);
      put('rect', 0.42, 0.13, 0.16, 0.11, cut);
      put('roundRect', 0.04, 0.22, 0.92, 0.6, { rectRadius: 0.07 * s });
      put('star6', 0.34, 0.36, 0.32, 0.34, cut);
      break;
    case 'bottle':                                                // pill bottle
      put('roundRect', 0.26, 0.02, 0.4, 0.18, { rectRadius: 0.04 * s });
      put('roundRect', 0.18, 0.2, 0.56, 0.62, { rectRadius: 0.06 * s });
      put('rect', 0.24, 0.38, 0.44, 0.22, cut);
      put('roundRect', 0.5, 0.7, 0.46, 0.24, { rectRadius: 0.12 * s });
      break;
    case 'bed':                                                   // hospital bed
      put('ellipse', 0.13, 0.3, 0.19, 0.19);
      put('rect', 0.03, 0.44, 0.11, 0.34);
      put('rect', 0.14, 0.54, 0.24, 0.12);
      put('roundRect', 0.36, 0.5, 0.6, 0.28, { rectRadius: 0.05 * s });
      put('rect', 0.62, 0.06, 0.34, 0.24);
      put('rect', 0.65, 0.11, 0.28, 0.14, cut);
      break;
    case 'hospital':                                              // clinic building
      put('triangle', 0.02, 0.04, 0.96, 0.26);
      cross(0.4, 0.06, 0.2, 0.07, cut);
      put('rect', 0.04, 0.3, 0.92, 0.66);
      put('rect', 0.42, 0.66, 0.16, 0.3, cut);
      [0.12, 0.26, 0.62, 0.76].forEach((cx) => [0.4, 0.54, 0.68].forEach((cy) => put('rect', cx, cy, 0.1, 0.08, cut)));
      break;
    case 'syringe':
      put('rect', 0.26, 0.44, 0.58, 0.15, { rotate: -45 });
      put('rect', 0.7, 0.16, 0.3, 0.07, { rotate: -45 });
      put('rect', 0.06, 0.7, 0.26, 0.09, { rotate: -45 });
      [0.2, 0.32, 0.44].forEach((d) => put('rect', 0.42 + d, 0.3 + d, 0.02, 0.15, { rotate: -45, fill: { color: hole } }));
      break;
    case 'stetho':                                                // stethoscope
      put('rect', 0.06, 0.04, 0.09, 0.3);
      put('rect', 0.55, 0.04, 0.09, 0.3);
      put('blockArc', 0.0, 0.1, 0.7, 0.62, { angleRange: [0, 180], arcThicknessRatio: 0.24 });
      put('rect', 0.31, 0.4, 0.08, 0.3);
      ring(0.6, 0.56, 0.4, 0.09);
      break;
    case 'appleCup':
      put('ellipse', 0.04, 0.44, 0.42, 0.42);
      put('rect', 0.22, 0.34, 0.06, 0.14, { rotate: 20 });
      put('trapezoid', 0.44, 0.36, 0.54, 0.56, { flipV: true });
      put('rect', 0.52, 0.06, 0.44, 0.1, { rotate: -30 });
      break;
    case 'appleBag':
      put('ellipse', 0.0, 0.48, 0.42, 0.42);
      put('rect', 0.18, 0.38, 0.06, 0.14, { rotate: 20 });
      put('trapezoid', 0.36, 0.3, 0.62, 0.64, { fill: NONE, line: { color, width: 0.08 * s * 72 } });
      put('rect', 0.44, 0.02, 0.42, 0.1, { rotate: -25 });
      break;
    case 'ivBag':
      put('roundRect', 0.18, 0.06, 0.6, 0.58, { rectRadius: 0.09 * s });
      put('rect', 0.44, 0.0, 0.07, 0.1);
      put('ellipse', 0.42, -0.02, 0.11, 0.11);
      put('rect', 0.45, 0.64, 0.06, 0.24);
      put('ellipse', 0.4, 0.86, 0.15, 0.15);
      break;
    case 'handHeart':
      put('heart', 0.4, 0.08, 0.5, 0.42);
      put('rect', 0.16, 0.58, 0.8, 0.15, { rotate: -10 });
      put('triangle', -0.02, 0.5, 0.34, 0.34, { rotate: -125 });
      break;
    case 'heartPulse':
      put('heart', 0.04, 0.14, 0.92, 0.78);
      put('rect', 0.1, 0.47, 0.8, 0.07, cut);
      put('rect', 0.36, 0.31, 0.07, 0.28, cut);
      put('rect', 0.56, 0.42, 0.07, 0.3, cut);
      break;
    case 'doctor':                                                // person + stethoscope
      put('ellipse', 0.29, 0.04, 0.42, 0.42);
      put('trapezoid', 0.04, 0.52, 0.92, 0.44);
      put('rect', 0.46, 0.46, 0.08, 0.28, cut);
      break;
    default:
      put('ellipse', 0.1, 0.1, 0.8, 0.8);
  }
}

/* ------------------------------------------------------ master decorations */

/** "MEDICALCLINIC" logo, top-left. */
function logo(slide, dark) {
  icon(slide, 'crossSolid', 0.393, 0.29, 0.358, dark ? C.white : C.ink, dark ? C.navy : C.white);
  text(slide, [
    { text: 'Medical', options: { color: dark ? C.white : C.ink } },
    { text: 'clinic', options: { color: C.teal } },
  ], { x: 0.71, y: 0.302, w: 1.6, h: 0.37, fontFace: HEAD, fontSize: 16, wrap: false });
}

/** "Healthcare Presentation" search pill, top-right. */
function searchPill(slide) {
  shape(slide, 'roundRect', { x: 10.593, y: 0.314, w: 2.428, h: 0.337, rectRadius: 0.169, fill: { color: C.teal } });
  text(slide, 'Healthcare Presentation', {
    x: 10.742, y: 0.373, w: 1.985, h: 0.219, align: 'center',
    fontSize: 7, charSpacing: 1, color: C.white,
  });
  shape(slide, 'ellipse', { x: 12.732, y: 0.36, w: 0.245, h: 0.245, fill: { color: C.deep } });
  icon(slide, 'search', 12.79, 0.418, 0.128, C.white, C.deep);
}

/** Everything inherited from the slide master. */
function chrome(slide, opts) {
  const o = opts || {};
  logo(slide, !!o.dark);
  searchPill(slide);
  if (o.page) {
    text(slide, String(o.page), {
      x: 12.574, y: 6.818, w: 0.469, h: 0.399, align: 'right', valign: 'middle',
      fontSize: 12, color: C.page,
    });
  }
}

/* ------------------------------------------- repeated content furniture */

/** Small blue kicker line above a headline. */
function eyebrow(slide, txt, x, y, w, o) {
  text(slide, txt, Object.assign({ x, y, w, h: 0.303, fontSize: 12, color: C.blue, wrap: false }, o));
}

/** Rounded pill button. */
function button(slide, txt, x, y, w, color, textX, textW) {
  shape(slide, 'roundRect', { x, y, w, h: 0.337, rectRadius: 0.1685, fill: { color } });
  text(slide, txt, {
    x: textX, y: y + 0.019, w: textW, h: 0.337, align: 'center', wrap: false,
    fontFace: HEAD, fontSize: 14, color: C.white,
  });
}

/** The "Learn More" / "Next Page" button pair. */
function buttonPair(slide, x, y) {
  button(slide, 'Learn More', x, y, 1.346, C.teal, x + 0.198, 0.951);
  button(slide, 'Next Page', x + 1.522, y, 1.346, C.blue, x + 1.773, 0.844);
}

/** White circle badge holding an icon. */
function badge(slide, name, x, y, d, plate, ink, iconScale) {
  shape(slide, 'ellipse', { x, y, w: d, h: d, fill: { color: plate }, shadow: CARD_SHADOW() });
  const is = d * (iconScale || 0.635);
  icon(slide, name, x + (d - is) / 2, y + (d - is) / 2, is, ink, plate);
}

/** Body paragraph at 150% leading. */
function para(slide, runs, x, y, w, h, o) {
  text(slide, runs, Object.assign({ x, y, w, h, fontSize: 10.5, lineSpacingMultiple: 1.5 }, o));
}

/** Two-tone Bebas headline; `parts` = [[text, color?], ...]. */
function headline(slide, parts, x, y, w, h, size, o) {
  const runs = parts.map((p, i) => ({
    text: p[0],
    options: Object.assign({ color: p[1] || C.ink }, p[2] || {}, i < parts.length - 1 && p[3] ? { breakLine: true } : {}),
  }));
  text(slide, runs, Object.assign({ x, y, w, h, fontFace: HEAD, fontSize: size }, o));
}

/* ==================================================================== deck */

const pres = new PptxGenJS();
pres.defineLayout({ name: 'DECK', width: 13.3333, height: 7.5 });
pres.layout = 'DECK';
pres.author = 'Medical Clinic';
pres.title = 'Healthcare Presentation';

/* --------------------------------------------------------------- slide 1 */

// Decorative sweeps: sampled [y, x] outlines of the two ribbons.
const S1_TEAL = [[0, 3.234], [0.417, 3.262], [0.833, 3.29], [1.25, 3.332], [1.667, 3.387], [2.083, 3.457],
  [2.5, 3.54], [2.917, 3.63], [3.333, 3.72], [3.75, 3.804], [4.16, 3.88], [4.576, 3.942], [4.993, 3.977],
  [5.41, 4.005], [5.826, 3.977], [6.243, 3.935], [6.66, 3.873], [7.076, 3.79], [7.5, 3.686]];
const S1_BLUE = [[0, 2.36], [0.417, 2.388], [0.833, 2.415], [1.25, 2.457], [1.667, 2.513], [2.083, 2.582],
  [2.5, 2.665], [2.917, 2.756], [3.333, 2.846], [3.75, 2.929], [4.16, 3.005], [4.576, 3.068], [4.993, 3.103],
  [5.41, 3.13], [5.826, 3.103], [6.243, 3.061], [6.66, 2.998], [7.076, 2.915], [7.5, 2.818]];

function slide1() {
  const s = pres.addSlide();
  imagePlaceholder(s, { x: 0, y: 0, w: 13.3333, h: 7.5, kind: 'rect', label: false });
  shape(s, 'rect', { x: 0, y: 0, w: 13.3333, h: 7.5, fill: { color: C.navy, transparency: 10 } });
  sweep(s, S1_TEAL, C.teal, 'left');
  sweep(s, S1_BLUE, C.band, 'left');

  shape(s, 'ellipse', { x: 1.693, y: 1.605, w: 4.316, h: 4.316, fill: { color: C.white } });
  imagePlaceholder(s, { x: 1.9, y: 1.818, w: 3.902, h: 3.902, kind: 'ellipse' });

  eyebrow(s, 'Compassionate Care For Your Community', 6.955, 1.857, 4.2, { color: C.white });
  headline(s, [['Comprehensive Patient-Centered Healthcare Solutions', C.white]],
    6.929, 2.221, 5.702, 2.524, 48);
  shape(s, 'roundRect', { x: 9.641, y: 3.984, w: 1.524, h: 0.432, rectRadius: 0.216, fill: { color: C.teal } });
  text(s, 'start New Page', { x: 9.799, y: 4.032, w: 1.3, h: 0.337, fontFace: HEAD, fontSize: 14, color: C.white, wrap: false });
  para(s, LOREM_A.replace('quis nost', 'PLACEHOLDER'),
    6.929, 4.814, 5.702, 0.868, { color: C.white });

  badge(s, 'crossRing', 5.024, 1.983, 1.06, C.white, C.blue, 0.667);
  badge(s, 'medbag', 1.63, 4.542, 1.152, C.white, C.blue, 0.614);
  chrome(s, { dark: true });
}

/* --------------------------------------------------------------- slide 2 */

function slide2() {
  const s = pres.addSlide();
  shape(s, 'round2SameRect', { x: 7.291, y: 4.633, w: 5.065, h: 2.867, fill: { color: C.white }, shadow: CARD_SHADOW() });
  imagePlaceholder(s, { x: 7.492, y: 4.835, w: 4.673, h: 2.656, kind: 'round2SameRect' });
  shape(s, 'roundRect', { x: 7.291, y: 1.114, w: 5.065, h: 3.228, rectRadius: rr(8432, 3.228), fill: { color: C.white }, shadow: CARD_SHADOW() });
  imagePlaceholder(s, { x: 7.473, y: 1.291, w: 4.691, h: 2.873, radius: 0.218 });

  eyebrow(s, 'Established To Provide Accessible', 1.169, 1.183, 3.2);
  headline(s, [['Committed to ', C.ink, {}, true], ['Excellence in ', C.ink, {}, true], ['Patient Care', C.teal]],
    1.169, 1.486, 4.356, 2.827, 54);
  para(s, [{ text: LOREM_A, options: { breakLine: true } }, { text: LOREM_B }], 1.169, 4.385, 4.112, 1.133);
  buttonPair(s, 1.252, 5.953);

  // Blue statistic card
  shape(s, 'roundRect', { x: 6.266, y: 1.958, w: 2.429, h: 4.332, rectRadius: rr(7354, 2.429), fill: { color: C.blue } });
  shape(s, 'roundRect', { x: 6.496, y: 2.284, w: 0.47, h: 0.47, rectRadius: 0.078, fill: { color: C.white }, shadow: CARD_SHADOW() });
  icon(s, 'bottle', 6.537, 2.325, 0.388, C.blue, C.white);
  text(s, '+4560', { x: 6.406, y: 3.014, w: 1.6, h: 0.707, fontFace: HEAD, fontSize: 36, color: C.white });
  text(s, 'quality healthcare to all families', { x: 6.406, y: 3.699, w: 1.758, h: 0.572, fontFace: HEAD, fontSize: 14, color: C.white });
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam.',
    6.406, 4.301, 2.046, 1.589, { fontSize: 10, color: C.white });
  chrome(s, { page: 2 });
}

/* --------------------------------------------------------------- slide 3 */

function slide3() {
  const s = pres.addSlide();
  // Four-up quilt: two photo tiles, two coloured copy tiles.
  const tileR = rr(9893, 2.617);
  [[0.828, 4.076], [3.729, 1.057]].forEach(([x, y]) => {
    shape(s, 'roundRect', { x, y, w: 2.617, h: 2.751, rectRadius: tileR, fill: { color: C.white }, shadow: CARD_SHADOW() });
  });
  imagePlaceholder(s, { x: 0.947, y: 4.223, w: 2.38, h: 2.474, radius: 0.235 });
  imagePlaceholder(s, { x: 3.847, y: 1.205, w: 2.38, h: 2.474, radius: 0.235 });

  const copyTile = (x, y, color, iconName, tail) => {
    shape(s, 'roundRect', { x, y, w: 2.617, h: 2.751, rectRadius: tileR, fill: { color }, shadow: CARD_SHADOW() });
    icon(s, iconName, x + 1.038, y + 0.319, 0.506, C.white, color);
    text(s, 'quality healthcare to all', { x: x + 0.272, y: y + 0.959, w: 2.126, h: 0.337, align: 'center', fontFace: HEAD, fontSize: 14, color: C.white });
    para(s, 'Lorem ipsum dolor amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et' + tail,
      x + 0.219, y + 1.3, 2.126, 1.133, { align: 'center', color: C.white });
  };
  copyTile(0.828, 1.057, C.teal, 'bed', '..');
  copyTile(3.729, 4.076, C.blue, 'hospital', '.');

  eyebrow(s, 'Guiding Our Care And Growth', 7.561, 1.564, 3.0);
  headline(s, [['Delivering patient-centered healthcare ', C.ink], ['aligned community', C.teal]],
    7.561, 1.867, 4.654, 2.322, 44);
  para(s, [{ text: LOREM_A, options: { breakLine: true } }, { text: LOREM_B }], 7.561, 4.396, 4.112, 1.133);
  buttonPair(s, 7.645, 5.965);
  chrome(s, { page: 3 });
}

/* --------------------------------------------------------------- slide 4 */

function slide4() {
  const s = pres.addSlide();
  shape(s, 'roundRect', { x: 7.291, y: 0.991, w: 5.065, h: 5.517, rectRadius: rr(8432, 5.065), fill: { color: C.white }, shadow: CARD_SHADOW() });
  imagePlaceholder(s, { x: 7.473, y: 1.169, w: 4.691, h: 4.185, radius: 0.318 });
  para(s, LOREM_MED, 7.785, 5.631, 4.112, 0.603, { align: 'center' });

  eyebrow(s, 'From Preventive Care To Treatment', 1.233, 1.187, 3.2);
  headline(s, [['Comprehensive Healthcare ', C.ink], ['Under One Roof', C.teal]], 1.233, 1.49, 3.823, 2.524, 48);
  para(s, LOREM_MED, 1.233, 4.124, 4.112, 0.603);

  const chip = (x, color, title) => {
    shape(s, 'roundRect', { x, y: 5.158, w: 2.283, h: 1.156, rectRadius: rr(6182, 1.156), fill: { color } });
    text(s, title, { x: x + 0.179, y: 5.307, w: 1.9, h: 0.337, fontFace: HEAD, fontSize: 14, color: C.white, wrap: false });
    para(s, 'Lorem ipsum dolor sit amet, nostrud elit.', x + 0.179, 5.587, 1.925, 0.578, { fontSize: 10, color: C.white });
  };
  chip(1.274, C.teal, 'General Checkup');
  chip(3.824, C.blue, 'Vaccinations');

  badge(s, 'syringe', 7.002, 2.35, 0.817, C.blue, C.white, 0.53);
  chrome(s, { page: 4 });
}

/* --------------------------------------------------------------- slide 5 */

function slide5() {
  const s = pres.addSlide();
  eyebrow(s, 'Quality Care And Compassionate Staff', 1.12, 1.014, 3.4);
  headline(s, [['Your Trusted and Healthcare Partner', C.ink, {}, true], ['& Hispital facility', C.teal]],
    1.12, 1.317, 4.871, 2.524, 48);

  // Two feature rows, top-right
  const featureRow = (y, iconName, color, title) => {
    icon(s, iconName, 7.342, y + 0.103, 0.587, color, C.white);
    text(s, title, { x: 8.138, y, w: 3.6, h: 0.438, fontFace: HEAD, fontSize: 20, color: C.ink });
    para(s, LOREM_ADIC, 8.138, y + 0.456, 4.483, 0.338);
  };
  featureRow(1.433, 'appleCup', C.blue, 'experienced doctors Healthcare');
  featureRow(2.627, 'appleBag', C.teal, 'modern equipment Healthcare');

  // Three cards along the bottom
  const infoCard = (x, color, iconName, title) => {
    shape(s, 'roundRect', { x, y: 4.37, w: 4.079, h: 2.132, rectRadius: rr(8432, 2.132), fill: { color: C.white }, shadow: CARD_SHADOW() });
    shape(s, 'roundRect', { x: x + 0.344, y: 4.611, w: 0.47, h: 0.47, rectRadius: 0.078, fill: { color }, shadow: CARD_SHADOW() });
    icon(s, iconName, x + 0.465, 4.731, 0.229, C.white, color);
    text(s, title, { x: x + 0.256, y: 5.2, w: 3.3, h: 0.505, fontFace: HEAD, fontSize: 24, color: C.ink, wrap: false });
    para(s, 'Lorem ipsum dolor sit amet, consectetur adic adipiscing elit, sed tempor incididunt ut  ', x + 0.256, 5.658, 3.567, 0.603);
  };
  infoCard(0.776, C.teal, 'stetho', 'compassionate staff');
  infoCard(5.211, C.blue, 'bed', 'commitment health');
  shape(s, 'roundRect', { x: 9.59, y: 4.37, w: 2.968, h: 2.132, rectRadius: rr(8432, 2.132), fill: { color: C.white }, shadow: CARD_SHADOW() });
  imagePlaceholder(s, { x: 9.685, y: 4.487, w: 2.748, h: 1.898, radius: 0.144 });
  chrome(s, { page: 5 });
}

/* --------------------------------------------------------------- slide 6 */

function slide6() {
  const s = pres.addSlide();
  shape(s, 'roundRect', { x: 1.122, y: 1.04, w: 4.345, h: 5.978, rectRadius: rr(8432, 4.345), fill: { color: C.white }, shadow: CARD_SHADOW() });
  imagePlaceholder(s, { x: 1.29, y: 1.195, w: 4.024, h: 5.678, radius: 0.307 });

  eyebrow(s, 'Compassionate Care for Your Community', 6.469, 1.554, 4.0);
  headline(s, [['Hospital Experienced and ', C.ink], ['Compassionate Professionals', C.teal]],
    6.469, 1.857, 5.887, 1.447, 40);

  shape(s, 'roundRect', { x: 4.463, y: 4.424, w: 7.666, h: 2.132, rectRadius: rr(8432, 2.132), fill: { color: C.white }, shadow: CARD_SHADOW() });
  imagePlaceholder(s, { x: 9.278, y: 4.541, w: 2.733, h: 1.898, radius: 0.144 });
  text(s, 'Dedicated to providing quality healthcare services', { x: 4.795, y: 4.781, w: 3.442, h: 0.774, fontFace: HEAD, fontSize: 20, color: C.ink });
  para(s, 'Lorem ipsum dolor sit amet, consectetur adic adipiscing elitLorem ipsum dolor sit amet, consectetur adic',
    4.795, 5.595, 4.129, 0.603);

  badge(s, 'ivBag', 0.825, 1.857, 0.817, C.blue, C.white, 0.61);
  chrome(s, { page: 6 });
}

/* --------------------------------------------------------------- slide 7 */

function slide7() {
  const s = pres.addSlide();
  [1.181, 4.117].forEach((y) => {
    shape(s, 'roundRect', { x: 7.061, y, w: 4.968, h: 2.519, rectRadius: rr(8432, 2.519), fill: { color: C.white }, shadow: CARD_SHADOW() });
    imagePlaceholder(s, { x: 7.199, y: y + 0.138, w: 4.691, h: 2.243, radius: 0.17 });
  });

  eyebrow(s, 'Ensuring Patient Comfort And Safety ', 1.443, 1.32, 3.3);
  headline(s, [['Equipped for Comprehensive ', C.ink], ['Healthcare', C.teal]], 1.443, 1.623, 3.823, 2.524, 48);
  para(s, LOREM_MED, 1.443, 4.257, 4.112, 0.603);

  const stat = (x, color, value) => {
    shape(s, 'roundRect', { x, y: 5.309, w: 1.835, h: 0.868, rectRadius: 0.145, fill: { color } });
    text(s, value, { x: x + 0.24, y: 5.423, w: 1.355, h: 0.64, align: 'center', fontFace: HEAD, fontSize: 32, bold: true, color: C.white });
  };
  stat(1.443, C.teal, '$76,0');
  stat(3.522, C.blue, '850K');

  badge(s, 'handHeart', 6.791, 1.624, 0.817, C.blue, C.white, 0.635);
  badge(s, 'heartPulse', 11.482, 5.36, 0.817, C.teal, C.white, 0.635);
  chrome(s, { page: 7 });
}

/* --------------------------------------------------------------- slide 8 */

function slide8() {
  const s = pres.addSlide();
  eyebrow(s, 'Ensuring A Smooth, Comfortable Healthcare Experience', 4.381, 0.91, 4.571, { align: 'center' });
  headline(s, [['Streamlined, Patient-', C.ink], ['Centered', C.teal], [' ', C.ink], ['Processes', C.teal]],
    2.592, 1.249, 8.15, 0.774, 40, { align: 'center' });

  // card x, photo x, badge x, copy x, colour, number, name
  const profiles = [
    [0.524, 0.808, 3.758, 3.661, C.blue, '01', 'Agustina falvia'],
    [6.249, 7.118, 10.119, 10.006, C.teal, '02', 'Michelino Ferry'],
  ];
  profiles.forEach(([cx, px, bx, tx, color, num, name]) => {
    shape(s, 'roundRect', { x: cx, y: 2.515, w: 5.975, h: 2.81, rectRadius: rr(8432, 2.81), fill: { color: C.white }, shadow: CARD_SHADOW() });
    imagePlaceholder(s, { x: px, y: 2.699, w: 2.227, h: 2.441, radius: 0.169 });
    shape(s, 'roundRect', { x: bx, y: 2.828, w: 0.594, h: 0.594, rectRadius: 0.099, fill: { color }, shadow: CARD_SHADOW() });
    text(s, num, { x: bx + 0.061, y: 2.906, w: 0.472, h: 0.505, align: 'center', fontFace: HEAD, fontSize: 24, color: C.white, wrap: false });
    text(s, name, { x: tx, y: 3.642, w: 2.4, h: 0.438, fontFace: HEAD, fontSize: 20, color: C.ink });
    para(s, LOREM_ADIC, tx, 4.056, 2.055, 0.868);
  });

  para(s, [{ text: LOREM_A, options: { breakLine: true } }, { text: LOREM_B }],
    1.095, 5.987, 11.143, 0.603, { align: 'center' });
  chrome(s, { page: 8 });
}

/* --------------------------------------------------------------- slide 9 */

function slide9() {
  const s = pres.addSlide();
  eyebrow(s, 'Insights into who we serve to improve', 5.08, 0.825, 3.173, { align: 'center' });
  headline(s, [['Specialized ', C.ink], ['departments', C.teal]], 2.592, 1.164, 8.15, 0.774, 40, { align: 'center' });

  ['Immanuel Johnson', 'Slavina Cistra', 'Alex Wlilliam'].forEach((name, i) => {
    const x = 1.156 + i * 3.912;
    shape(s, 'roundRect', { x, y: 2.202, w: 3.197, h: 4.802, rectRadius: rr(8432, 3.197), fill: { color: C.white }, shadow: CARD_SHADOW() });
    imagePlaceholder(s, { x: x + 0.218, y: 2.386, w: 2.76, h: 2.88, radius: 0.207 });
    text(s, name, { x: x + 0.571, y: 5.684, w: 2.055, h: 0.438, align: 'center', fontFace: HEAD, fontSize: 20, color: C.ink });
    para(s, LOREM_ADIC, x + 0.381, 6.097, 2.407, 0.603, { align: 'center' });
  });

  badge(s, 'handHeart', 0.935, 2.7, 0.817, C.blue, C.white, 0.635);
  badge(s, 'heartPulse', 11.551, 4.104, 0.817, C.teal, C.white, 0.635);
  chrome(s, { page: 9 });
}

/* -------------------------------------------------------------- slide 10 */

function slide10() {
  const s = pres.addSlide();
  shape(s, 'roundRect', { x: 0.682, y: 1.193, w: 7.602, h: 5.113, rectRadius: rr(5694, 5.113), fill: { color: C.white }, shadow: CARD_SHADOW() });

  const cats = ['Kategori 1', 'Kategori 2', 'Kategori 3', 'Kategori 4'];
  s.addChart(pres.ChartType.bar, [
    { name: 'Seri 1', labels: cats, values: [4.31, 2.5, 3.51, 4.5] },
    { name: 'Seri 2', labels: cats, values: [2.4, 4.41, 1.81, 2.8] },
    { name: 'Seri 3', labels: cats, values: [2.0, 2.0, 3.01, 5.0] },
  ], {
    // Frame matches the white card; `layout` pins the plot area so the grid
    // lands on the same coordinates as the original hand-drawn chart.
    x: 0.682, y: 1.193, w: 7.602, h: 5.113,
    layout: { x: 0.1776, y: 0.0698, w: 0.7597, h: 0.8021 },
    barDir: 'bar', barGrouping: 'clustered', barGapWidthPct: 176, barOverlapPct: 0,
    chartColors: [C.navy, C.teal, C.blue],
    showLegend: false, showTitle: false,
    valAxisMinVal: 0, valAxisMaxVal: 6, valAxisMajorUnit: 1,
    valAxisLabelFontFace: BODY, valAxisLabelFontSize: 12, valAxisLabelColor: '595959',
    catAxisLabelFontFace: BODY, catAxisLabelFontSize: 12, catAxisLabelColor: '595959',
    valGridLine: { color: C.grid, size: 0.75 }, catGridLine: { style: 'none' },
    valAxisLineColor: C.grid, catAxisLineColor: C.grid,
    dataBorder: { pt: 0, color: 'FFFFFF' },
  });

  eyebrow(s, 'Guiding Our Care And Growth', 9.175, 1.789, 3.0);
  headline(s, [['Our Healthcare', C.ink, {}, true], ['Chart', C.teal]], 9.175, 2.092, 3.477, 1.582, 44);
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, ',
    9.175, 3.8, 3.179, 1.133);
  buttonPair(s, 9.253, 5.355);
  chrome(s, { page: 10 });
}

/* -------------------------------------------------------------- slide 11 */

// Stethoscope illustration, traced from the original freeforms.
const STETH_TUBE = [['M', 0.1478, 0], ['C', 0.1471, 0.3555, 0.1498, 0.5393, 0.1485, 0.6908], ['C', 0.148, 0.747, 0.1453, 0.8648, 0.1251, 0.9351], ['C', 0.1075, 0.9956, 0.0855, 1, 0.0758, 0.9984], ['C', 0.0653, 0.9967, 0.0349, 0.9836, 0.0157, 0.8969], ['C', 0, 0.8266, 0.0033, 0.7465, 0.0043, 0.7236], ['C', 0.0092, 0.6052, 0.0371, 0.5382, 0.0458, 0.5169], ['C', 0.0686, 0.4618, 0.0919, 0.4504, 0.1086, 0.4433], ['C', 0.1242, 0.4362, 0.1501, 0.4313, 0.2163, 0.4896], ['C', 0.2308, 0.5027, 0.245, 0.5164, 0.3001, 0.5785], ['C', 0.3652, 0.6521, 0.3838, 0.6756, 0.4269, 0.7236], ['C', 0.4816, 0.7841, 0.509, 0.8135, 0.5227, 0.8239], ['C', 0.6819, 0.9411, 0.8368, 0.7497, 0.8879, 0.6581], ['C', 0.9469, 0.5523, 1, 0.4024, 1, 0.2219]];
const STETH_ARM_L = [['M', 1, 0.9722], ['C', 0.7707, 1, 0.7707, 1, 0.7707, 1], ['C', 0.7707, 1, 0.5173, 0.8154, 0.3333, 0.5961], ['C', 0.152, 0.381, 0.056, 0.2935, 0.0427, 0.2422], ['C', 0.0267, 0.1908, 0, 0.1124, 0.2293, 0.0507], ['C', 0.416, 0, 0.7707, 0.0028, 0.896, 0.0028], ['C', 0.944, 0.0028, 0.9573, 0.0028, 0.9573, 0.0028], ['C', 0.9573, 0.0597, 0.9573, 0.0597, 0.9573, 0.0597], ['C', 0.9573, 0.0597, 0.9013, 0.0597, 0.8507, 0.0597], ['C', 0.7093, 0.0597, 0.464, 0.0604, 0.36, 0.102], ['C', 0.2427, 0.1471, 0.2587, 0.1846, 0.2667, 0.2241], ['C', 0.2773, 0.2637, 0.4107, 0.415, 0.5867, 0.6225], ['C', 0.8, 0.8744, 1, 0.9722, 1, 0.9722], ['Z']];
const STETH_ARM_R = [['M', 0, 0.9722], ['C', 0.2293, 1, 0.2293, 1, 0.2293, 1], ['C', 0.2293, 1, 0.4827, 0.8154, 0.6667, 0.5961], ['C', 0.848, 0.381, 0.9413, 0.2935, 0.9573, 0.2422], ['C', 0.9707, 0.1908, 1, 0.1124, 0.7707, 0.0507], ['C', 0.584, 0, 0.2293, 0.0028, 0.104, 0.0028], ['C', 0.056, 0.0028, 0.0427, 0.0028, 0.0427, 0.0028], ['C', 0.0427, 0.0597, 0.0427, 0.0597, 0.0427, 0.0597], ['C', 0.0427, 0.0597, 0.096, 0.0597, 0.1493, 0.0597], ['C', 0.2907, 0.0597, 0.5333, 0.0604, 0.64, 0.102], ['C', 0.7547, 0.1471, 0.7413, 0.1846, 0.7333, 0.2241], ['C', 0.7227, 0.2637, 0.5893, 0.415, 0.4133, 0.6225], ['C', 0.2, 0.8744, 0, 0.9722, 0, 0.9722], ['Z']];
const STETH_TIP_L = [['M', 0.7259, 0.0223], ['C', 0.7259, 0.0223, 0.4407, 0, 0.3556, 0.0223], ['C', 0.2741, 0.0391, 0.0556, 0.1397, 0.0148, 0.2235], ['C', 0, 0.2626, 0, 0.3017, 0, 0.324], ['C', 0, 0.3687, 0, 0.4358, 0.0037, 0.4972], ['C', 0, 0.5642, 0, 0.6313, 0, 0.676], ['C', 0, 0.6927, 0, 0.7374, 0.0148, 0.7709], ['C', 0.0556, 0.8547, 0.2741, 0.9609, 0.3556, 0.9777], ['C', 0.4407, 1, 0.7259, 0.9777, 0.7259, 0.9777], ['C', 0.8778, 0.9777, 1, 0.7933, 1, 0.5642], ['C', 1, 0.2067, 0.8778, 0.0223, 0.7259, 0.0223], ['Z']];
const STETH_TIP_R = [['M', 0.2741, 0.0223], ['C', 0.2741, 0.0223, 0.5593, 0, 0.6444, 0.0223], ['C', 0.7259, 0.0391, 0.9444, 0.1397, 0.9815, 0.2235], ['C', 1, 0.2626, 0.9963, 0.3017, 0.9963, 0.324], ['C', 0.9963, 0.3687, 0.9963, 0.4358, 0.9963, 0.4972], ['C', 0.9963, 0.5642, 0.9963, 0.6313, 0.9963, 0.676], ['C', 0.9963, 0.6927, 1, 0.7374, 0.9815, 0.7709], ['C', 0.9444, 0.8547, 0.7259, 0.9609, 0.6444, 0.9777], ['C', 0.5593, 1, 0.2741, 0.9777, 0.2741, 0.9777], ['C', 0.1222, 0.9777, 0, 0.7933, 0, 0.5642], ['C', 0, 0.2067, 0.1222, 0.0223, 0.2741, 0.0223], ['Z']];
const STETH_BELL = [['M', 0.9668, 0.0424], ['C', 0.9442, 0.0212, 0.925, 0.0042, 0.897, 0.0014], ['C', 0.8691, 0, 0.8482, 0.0607, 0.8482, 0.0607], ['C', 0.8072, 0.2203, 0.7007, 0.7444, 0.5061, 0.7444], ['C', 0.5044, 0.7444, 0.5017, 0.7444, 0.5, 0.7444], ['C', 0.4974, 0.7444, 0.4956, 0.7444, 0.493, 0.7444], ['C', 0.2984, 0.7444, 0.1928, 0.2203, 0.1518, 0.0607], ['C', 0.1518, 0.0607, 0.1309, 0, 0.103, 0.0014], ['C', 0.075, 0.0042, 0.055, 0.0212, 0.0323, 0.0424], ['C', 0, 0.0734, 0.0148, 0.1497, 0.0148, 0.1497], ['C', 0.1108, 0.7105, 0.2801, 1, 0.493, 1], ['C', 0.4956, 1, 0.4974, 1, 0.5, 1], ['C', 0.5017, 1, 0.5044, 1, 0.5061, 1], ['C', 0.7199, 1, 0.8892, 0.7105, 0.9852, 0.1497], ['C', 0.9852, 0.1497, 1, 0.0734, 0.9668, 0.0424], ['Z']];
// Teardrop marker that holds each infographic icon.
const PIN = [['M', 0.9035, 0.3442], ['C', 0.8263, 0.2837, 0.7529, 0.2047, 0.6911, 0.1209], ['C', 0.6178, 0.0465, 0.5212, 0, 0.417, 0], ['C', 0.1853, 0, 0, 0.2233, 0, 0.5023], ['C', 0, 0.7767, 0.1853, 1, 0.417, 1], ['C', 0.5212, 1, 0.6178, 0.9535, 0.6911, 0.8744], ['C', 0.7529, 0.7953, 0.8263, 0.7209, 0.9035, 0.6558], ['C', 0.9575, 0.6093, 1, 0.5581, 1, 0.5023], ['C', 1, 0.4419, 0.9575, 0.3907, 0.9035, 0.3442], ['Z']];

/** Teardrop pin + coloured disc + white icon. */
function pinMarker(slide, x, y, color, iconName) {
  freeform(slide, PIN, { x, y, w: 1.026, h: 0.842, fill: { color: C.pin }, shadow: PIN_SHADOW() });
  shape(slide, 'ellipse', { x: x + 0.099, y: y + 0.096, w: 0.659, h: 0.651, fill: { color }, shadow: PIN_SHADOW() });
  icon(slide, iconName, x + 0.242, y + 0.235, 0.372, C.white, color);
}

function slide11() {
  const s = pres.addSlide();
  headline(s, [['Our Healthy ', C.ink], ['Infographic', C.teal]], 4.043, 0.81, 5.247, 0.841, 44, { align: 'center', wrap: false });

  freeform(s, STETH_TUBE, { x: 4.375, y: 4.522, w: 7.304, h: 2.089, fill: NONE, line: { color: C.softBlue, width: 10 } });
  freeform(s, STETH_ARM_L, { x: 4.525, y: 2.549, w: 0.427, h: 1.494, fill: { color: C.midBlue } });
  freeform(s, STETH_TIP_L, { x: 4.929, y: 2.506, w: 0.306, h: 0.186, fill: { color: C.darkBlue } });
  freeform(s, STETH_ARM_R, { x: 5.956, y: 2.549, w: 0.425, h: 1.494, fill: { color: C.midBlue } });
  freeform(s, STETH_TIP_R, { x: 5.673, y: 2.506, w: 0.306, h: 0.186, fill: { color: C.darkBlue } });
  freeform(s, STETH_BELL, { x: 4.801, y: 3.955, w: 1.302, h: 0.734, fill: { color: C.richBlue } });
  shape(s, 'ellipse', { x: 11.191, y: 3.987, w: 1.03, h: 1.03, fill: { color: C.midBlue } });
  shape(s, 'ellipse', { x: 11.303, y: 4.098, w: 0.806, h: 0.806, fill: { color: C.blue } });
  shape(s, 'ellipse', { x: 11.546, y: 4.343, w: 0.317, h: 0.317, fill: { color: C.white } });

  text(s, 'Your Title Here', { x: 1.205, y: 3.466, w: 2.668, h: 0.505, fontFace: HEAD, fontSize: 24, color: C.ink });
  para(s, 'In voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non',
    1.205, 3.979, 2.74, 0.868);

  pinMarker(s, 7.308, 2.835, C.navy, 'doctor');
  para(s, 'Illo inventore veritatis beatae dicta', 8.514, 2.919, 1.794, 0.603);
  pinMarker(s, 7.308, 4.343, C.teal, 'doctor');
  para(s, 'Illo inventore veritatis beatae dicta', 8.514, 4.45, 1.859, 0.603);
  chrome(s, { page: 11 });
}

/* -------------------------------------------------------------- slide 12 */

const THERM_SHELL = [['M', 0.9992, 0.0915], ['C', 0.9843, 0.0405, 0.7677, 0, 0.5, 0], ['C', 0.2323, 0, 0.0157, 0.0405, 0.0008, 0.0915], ['L', 0, 0.6963], ['C', 0, 0.7563, 0.2026, 0.7857, 0.2026, 0.8274], ['L', 0.2026, 0.9429], ['C', 0.2043, 0.9744, 0.3353, 1, 0.5, 1], ['C', 0.6647, 1, 0.7957, 0.9744, 0.7974, 0.9429], ['L', 0.7974, 0.8274], ['C', 0.7974, 0.7857, 1, 0.7563, 1, 0.6963], ['L', 1, 0.0915], ['Z']];
const THERM_INNER = [['M', 0.9965, 0.0694], ['C', 0.9816, 0.031, 0.7673, 0.0005, 0.5, 0], ['C', 0.2339, 0.0005, 0.0184, 0.031, 0.0035, 0.0694], ['L', 0.0023, 0.7189], ['C', 0, 0.7617, 0.2477, 0.7798, 0.2477, 0.8529], ['L', 0.2477, 0.9631], ['C', 0.2477, 0.9834, 0.3594, 0.9998, 0.5, 1], ['C', 0.6406, 0.9998, 0.7523, 0.9834, 0.7535, 0.9631], ['L', 0.7535, 0.8529], ['C', 0.7535, 0.7798, 1, 0.7617, 0.9965, 0.7189], ['L', 0.9977, 0.0694], ['Z']];
const THERM_TRACK = [['M', 1, 0.0273], ['C', 0.9853, 0.0122, 0.7684, 0, 0.5, 0], ['C', 0.2353, 0, 0.0184, 0.0122, 0.0037, 0.0273], ['L', 0, 0.971], ['C', 0.0037, 0.987, 0.2243, 1, 0.5, 1], ['C', 0.7757, 1, 1, 0.987, 1, 0.971], ['L', 1, 0.0273], ['Z']];
const THERM_BULB = [['M', 0, 0], ['L', 1, 0], ['L', 1, 0.2101], ['C', 1, 0.4534, 1, 0.7509, 1, 0.7509], ['C', 0.9977, 0.888, 0.7768, 0.9989, 0.5034, 1], ['C', 0.4966, 1, 0.4966, 1, 0.4943, 1], ['C', 0.221, 0.9989, 0, 0.888, 0, 0.7509], ['C', 0, 0.7509, 0, 0.4534, 0, 0.2101], ['Z']];

function slide12() {
  const s = pres.addSlide();
  headline(s, [['Our Healthy ', C.ink], ['Infographic', C.teal]], 4.043, 0.81, 5.247, 0.841, 44, { align: 'center', wrap: false });

  // Scale labels
  [['High', 2.188, C.navy], ['Medium', 3.762, C.blue], ['Low', 5.242, C.teal]].forEach(([label, y, color]) => {
    text(s, label, { x: 4.066, y, w: 1.479, h: 0.471, align: 'right', fontSize: 22, bold: true, color, lineSpacingMultiple: 1 });
  });
  // Tick marks: long ones line up with the labels, short ones in between.
  [[2.382, 'long'], [2.797, 'short'], [3.183, 'short'], [3.598, 'short'], [3.972, 'long'],
    [4.318, 'short'], [4.704, 'short'], [5.119, 'short'], [5.449, 'long']].forEach(([y, kind]) => {
    const long = kind === 'long';
    shape(s, 'rect', { x: long ? 5.723 : 6.005, y, w: long ? 0.497 : 0.214, h: 0.054, fill: { color: '000000' } });
  });

  // Thermometer body
  freeform(s, THERM_SHELL, { x: 6.437, y: 1.79, w: 0.891, h: 5.173, fill: { color: C.navy } });
  freeform(s, THERM_INNER, { x: 6.565, y: 1.941, w: 0.636, h: 4.916, fill: { color: C.white } });
  freeform(s, THERM_TRACK, { x: 6.783, y: 2.258, w: 0.2, h: 4.183, fill: { color: C.track } });
  freeform(s, THERM_BULB, { x: 6.722, y: 6.129, w: 0.321, h: 0.728, fill: { color: C.teal } });
  shape(s, 'roundRect', { x: 6.783, y: 2.754, w: 0.2, h: 3.632, rectRadius: 0.1, fill: { color: C.teal } });

  para(s, EAQUE + ' dicta sunt explicabo. Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut odit aut fugit, sed quia consequuntur magni dolores eos qui ratione ',
    0.851, 2.979, 2.58, 1.928);

  [[2.219, C.navy, 'medkit', 2.152], [3.808, C.blue, 'bottle', 3.74], [5.397, C.teal, 'ivBag', 5.329]]
    .forEach(([y, color, iconName, ty], i) => {
      pinMarker(s, i === 0 ? 8.173 : 8.154, y, color, iconName);
      para(s, EAQUE, 9.62, ty, 2.38, 0.868);
    });
  chrome(s, { page: 12 });
}

/* -------------------------------------------------------------- slide 13 */

// Tapered aluminium base of the laptop, widening towards the front edge.
const LAPTOP_BASE = [['M', 0.056, 0], ['L', 0.944, 0], ['L', 1, 1], ['L', 0, 1], ['Z']];

/** Open laptop: black lid with a notch, screen well, tapered aluminium base. */
function laptop(slide, x, y, w) {
  const u = w / 6.597; // artwork was authored 6.597in wide
  const at = (dx, dy, dw, dh) => ({ x: x + dx * u, y: y + dy * u, w: dw * u, h: dh * u });
  shape(slide, 'roundRect', Object.assign(at(0.953, 0.799, 4.699, 3.217), { rectRadius: 0.05 * u, fill: { color: '111111' } }));
  imagePlaceholder(slide, Object.assign(at(1.046, 0.899, 4.512, 2.93), { radius: 0.02 * u }));
  shape(slide, 'round2SameRect', Object.assign(at(3.045, 0.799, 0.52, 0.2), { rotate: 180, fill: { color: '111111' } }));
  freeform(slide, LAPTOP_BASE, Object.assign(at(0.599, 4.008, 5.397, 0.25), { fill: { color: 'A5A39A' } }));
  shape(slide, 'roundRect', Object.assign(at(0.599, 4.24, 5.397, 0.15), { rectRadius: 0.075 * u, fill: { color: '8B8880' } }));
}

function slide13() {
  const s = pres.addSlide();
  headline(s, [['Our Health ', C.ink, { bold: true }, true], ['Mockup', C.teal, { bold: true }]],
    1.157, 0.953, 4.118, 1.919, 54);

  laptop(s, 6.143, 0.288, 6.597);
  laptop(s, 0.289, 2.345, 6.597);

  // Floating summary card
  shape(s, 'roundRect', { x: 4.779, y: 3.622, w: 6.142, h: 1.74, rectRadius: rr(9693, 1.74), fill: { color: C.white }, shadow: LIFT_SHADOW() });
  text(s, 'Health Mockup', { x: 5.101, y: 3.917, w: 2.678, h: 0.303, margin: 0, fontFace: HEAD, fontSize: 18, bold: true, color: C.ink });
  text(s, '$678,2 M', { x: 4.985, y: 4.284, w: 2.108, h: 0.572, fontFace: HEAD, fontSize: 28, color: C.teal });
  text(s, 'Track and Print Report', { x: 5.088, y: 4.89, w: 1.9, h: 0.177, margin: 0, fontSize: 10.5, color: C.ink });

  const deviceRow = (y, color, label) => {
    shape(s, 'roundRect', { x: 7.779, y, w: 2.936, h: 0.601, rectRadius: rr(35582, 0.601), fill: { color, transparency: 85 } });
    shape(s, 'ellipse', { x: 7.947, y: y + 0.254, w: 0.092, h: 0.092, fill: { color } });
    text(s, label, { x: 8.169, y: y + 0.208, w: 1.387, h: 0.185, margin: 0, fontSize: 11, bold: true, color });
    shape(s, 'roundRect', { x: 9.876, y: y + 0.191, w: 0.645, h: 0.218, rectRadius: 0.109, fill: { color } });
    text(s, '+657%', { x: 9.811, y: y + 0.182, w: 0.774, h: 0.236, align: 'center', valign: 'middle', fontFace: HEAD, fontSize: 8, color: 'E8E8E8' });
  };
  deviceRow(3.812, C.teal, 'Device 01');
  deviceRow(4.571, C.blue, 'Device 02');

  text(s, 'Serving Beyond the Clinic Walls', { x: 7.109, y: 5.892, w: 4.002, h: 0.337, fontFace: HEAD, fontSize: 14, color: C.blue });
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere',
    7.109, 6.235, 3.884, 0.579, { fontSize: 10 });
  chrome(s, { page: 13 });
}

/* -------------------------------------------------------------- slide 14 */

function slide14() {
  const s = pres.addSlide();
  // Phone mock-up: dark body, grey screen, notch. Runs off the bottom edge.
  shape(s, 'roundRect', { x: 4.356, y: 2.607, w: 5.014, h: 7.203, rectRadius: 0.42, fill: { color: '111111' } });
  imagePlaceholder(s, { x: 4.462, y: 2.7, w: 4.794, h: 6.996, radius: 0.34 });
  shape(s, 'round2SameRect', { x: 6.307, y: 2.7, w: 1.045, h: 0.193, rotate: 180, fill: { color: '111111' } });

  headline(s, [['Our Health ', C.ink, { bold: true }, true], ['Mockup', C.teal, { bold: true }]],
    0.902, 0.987, 4.118, 1.919, 54);
  para(s, 'Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada',
    0.902, 3.36, 3.202, 0.97, { fontSize: 12 });

  // Trend card (top right)
  shape(s, 'roundRect', { x: 8.484, y: 1.177, w: 3.326, h: 2.573, rectRadius: rr(7703, 2.573), fill: { color: C.white }, shadow: LIFT_SHADOW() });
  const gridY = [1.526, 1.703, 1.874, 2.046, 2.218, 2.39, 2.562, 2.734, 2.906, 3.078, 3.245];
  gridY.forEach((y, i) => {
    const edge = i === 0 || i === gridY.length - 1 || i === 3;
    polyline(s, [[8.84, y], [11.455, y]], { line: { color: edge ? 'DCDCDC' : 'F2F2F2', width: 0.75 } });
  });
  const series = [[9.106, 2.729], [9.626, 2.385], [10.147, 2.219], [10.667, 1.875], [11.188, 2.562]];
  polyline(s, series, { line: { color: C.blue, width: 1.25 } });
  series.forEach(([px, py]) => shape(s, 'ellipse', { x: px - 0.037, y: py - 0.037, w: 0.073, h: 0.073, fill: { color: C.blue } }));
  ['2020', '2021', '2022', '2023', '2024'].forEach((yr, i) => {
    text(s, yr, { x: 8.915 + i * 0.5224, y: 3.258, w: 0.5, h: 0.249, fontFace: HEAD, fontSize: 8.25, color: '404040', wrap: false });
  });

  // Numbered list (bottom right)
  [['01', 4.637], ['02', 5.684]].forEach(([num, y]) => {
    text(s, num, { x: 9.782, y, w: 0.917, h: 0.707, valign: 'middle', fontFace: HEAD, fontSize: 36, bold: true, color: C.blue });
    text(s, 'Inventory and Order Management', { x: 10.534, y: y + 0.067, w: 2.059, h: 0.572, fontSize: 14, color: C.ink });
  });

  // KPI card with progress ring
  shape(s, 'roundRect', { x: 0.856, y: 5.002, w: 4.13, h: 1.498, rectRadius: rr(11691, 1.498), fill: { color: C.white }, shadow: LIFT_SHADOW() });
  text(s, 'Device Mockup', { x: 1.335, y: 5.199, w: 1.933, h: 0.236, fontFace: HEAD, fontSize: 14, bold: true, color: C.ink });
  text(s, '$44.568', { x: 1.335, y: 5.487, w: 2.115, h: 0.606, fontFace: HEAD, fontSize: 36, color: C.teal });
  text(s, '-134 for last month', { x: 1.335, y: 6.119, w: 1.9, h: 0.185, fontSize: 11, color: C.ink });
  shape(s, 'arc', {
    x: 3.666, y: 5.331, w: 0.841, h: 0.841, angleRange: [269, 199],
    fill: NONE, line: { color: C.navy, width: 15 },
  });
  shape(s, 'ellipse', { x: 3.77, y: 5.434, w: 0.634, h: 0.634, fill: { color: C.white }, shadow: LIFT_SHADOW() });
  text(s, '80%', { x: 3.739, y: 5.659, w: 0.695, h: 0.185, align: 'center', valign: 'middle', fontFace: HEAD, fontSize: 11, bold: true, color: C.navy });
  chrome(s, { page: 14 });
}

/* -------------------------------------------------------------- slide 15 */

const S15_TEAL = [[0, 10.39], [0.417, 10.62], [0.833, 10.793], [1.25, 10.904], [1.667, 10.967], [2.083, 10.932],
  [2.5, 10.835], [2.917, 10.675], [3.333, 10.467], [3.75, 10.224], [4.16, 9.974], [4.576, 9.717], [4.993, 9.481],
  [5.41, 9.266], [5.826, 9.092], [6.243, 8.961], [6.66, 8.856], [7.076, 8.801], [7.5, 8.78]];
const S15_BLUE = [[0, 11.015], [0.417, 11.237], [0.833, 11.411], [1.25, 11.529], [1.667, 11.584], [2.083, 11.556],
  [2.5, 11.459], [2.917, 11.3], [3.333, 11.091], [3.75, 10.849], [4.16, 10.599], [4.576, 10.342], [4.993, 10.106],
  [5.41, 9.891], [5.826, 9.717], [6.243, 9.585], [6.66, 9.481], [7.076, 9.426], [7.5, 9.405]];

function slide15() {
  const s = pres.addSlide();
  imagePlaceholder(s, { x: 0, y: 0, w: 13.3333, h: 7.5, kind: 'rect', label: false });
  shape(s, 'rect', { x: 0, y: 0, w: 13.3333, h: 7.5, fill: { color: C.navy, transparency: 10 } });
  // The search pill sits under the ribbons on the closing slide, so it is hidden.
  sweep(s, S15_TEAL, C.teal, 'right');
  sweep(s, S15_BLUE, C.band, 'right');

  eyebrow(s, 'Book your appointment or reach out with any questions', 0.996, 2.21, 5.0, { color: C.white });
  headline(s, [['We\u2019re Here for Your Health & Thank You', C.white]], 0.969, 2.574, 5.702, 1.717, 48);
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea.',
    0.969, 4.429, 5.145, 0.868, { color: C.white });

  shape(s, 'ellipse', { x: 7.551, y: 1.589, w: 4.316, h: 4.316, fill: { color: C.white } });
  imagePlaceholder(s, { x: 7.761, y: 1.798, w: 3.898, h: 3.898, kind: 'ellipse' });
  badge(s, 'crossRing', 10.82, 1.937, 1.06, C.white, C.blue, 0.667);
  badge(s, 'medbag', 7.574, 4.523, 0.985, C.white, C.blue, 0.614);
  logo(s, true);
}

/* ------------------------------------------------------------------ write */

[slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8,
  slide9, slide10, slide11, slide12, slide13, slide14, slide15].forEach((fn) => fn());

const outFile = path.join(__dirname, '10ab7517-2c79-4d71-a6ea-ed7429e103c6_grok_final.pptx');
pres.writeFile({ fileName: outFile }).then(() => console.log('wrote', outFile));
