/**
 * "Cybersecurity Essentials" — 16-slide deck rebuilt with pptxgenjs.
 *
 * Slide size 13.333 x 7.5 in (LAYOUT_WIDE), black background, teal/green accents.
 * Headings use Inconsolata (theme major font), body copy uses Poppins Light.
 *
 * Raster artwork from the original deck is replaced by programmatic placeholders,
 * and the original's 45-degree teal->black gradients are approximated with stacked
 * translucent layers (pptxgenjs only emits solid fills).
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const C = {
  black: '000000',
  white: 'FFFFFF',
  teal: '73B6BF', // theme accent1
  green: '51B49B', // theme accent2
  blue: '5B9BD5', // theme accent5
  gold: 'FFC000', // theme accent4
  grey: 'A5A5A5', // theme accent3
  navy: '51647E', // theme dk2, lightened
  bone: 'E7E6E6', // theme lt2
  skin: 'FFBC96',
  ink: '262626',
  soot: '0D0D0D',
  dim: '808080', // 50% white over the black background (rules, page numbers)
};

const FONT = { head: 'Inconsolata', body: 'Poppins Light', quote: 'Anton' };

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

/* --------------------------------------------------------------- copy decks */

const LOREM = {
  labore: 'Lorem ipsum dolor sit elit sed do amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore.',
  laboreTempor:
    'Lorem ipsum dolor sit elit sed do amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore eiusmod tempor.',
  eiusmod: 'Lorem ipsum dolor sit elit sed do amet, consectetur adipiscing elit, sed do eiusmod.',
  sedDo: 'Lorem ipsum dolor sit amet elit sed do amet, consectetur adipiscing elit, sed do',
  consectetur: 'Lorem ipsum dolor sit elit sed do amet, consectetur',
  tempor: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.',
  exercitation:
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation.',
  nostrud: 'Lorem ipsum dolor sit amet, dolore nostrud adipiscing elit,',
  aenean: 'Lorem ipsum dolor sit amet, consectetuer adipiscing. Aenean commodo ligula eget dolor. Aenean massa it Cum sociis ligula',
  aeneanShort: 'Lorem ipsum dolor sit amet, ligula eget adipiscing Aenean commodo.',
  aeneanLigula: 'Lorem ipsum dolor sit amet, consectetuer adipiscing. Aenean commodo ligula.',
  adipiscing: 'Lorem ipsum dolor sit amet, consectetuer adipiscing. ',
  edet: 'Lorem ipsum dolor sit amet, dolor consectetuer adipiscing. edet Aenean commodo ligula eget',
  utEnim:
    'Lorem ipsum dolor sit sed amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim.',
  labore2: 'Lorem ipsum dolor do elit, sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore',
};

/* ------------------------------------------------------------ text builders */

/** Base text writer: PowerPoint text boxes anchor at the top, pptxgenjs does not. */
function text(slide, str, o) {
  slide.addText(str, Object.assign({ fontFace: FONT.body, fontSize: 18, color: C.white, valign: 'top' }, o));
}

/** Big gradient-filled section title of the original; flattened to solid accent1. */
function title(slide, str, o) {
  text(slide, str, Object.assign({ fontFace: FONT.head, fontSize: 44, color: C.teal }, o));
}

/** 12 pt justified Poppins Light paragraph, 130% leading — the deck's body style. */
function body(slide, str, o) {
  text(slide, str, Object.assign({ fontSize: 12, align: 'justify', lineSpacingMultiple: 1.3 }, o));
}

/** 16 pt Inconsolata sub-heading that captions every card. */
function subhead(slide, str, o) {
  text(slide, str, Object.assign({ fontFace: FONT.head, fontSize: 16 }, o));
}

/* ----------------------------------------------------------- shape builders */

function rect(slide, x, y, w, h, o) {
  slide.addShape('rect', Object.assign({ x, y, w, h }, o));
}

function disc(slide, cx, cy, d, color, o) {
  slide.addShape('ellipse', Object.assign({ x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { color } }, o));
}

function ring(slide, cx, cy, d, color, pt) {
  slide.addShape('ellipse', { x: cx - d / 2, y: cy - d / 2, w: d, h: d, line: { color, width: pt } });
}

/**
 * Free-form outline. `pts` holds unit coordinates (0..1 of w/h):
 *   [x, y]                          -> moveTo / lineTo
 *   ['c', x1, y1, x2, y2, x, y]     -> cubic bezier
 *   'z'                             -> close
 */
function poly(slide, x, y, w, h, pts, o) {
  const points = pts.map((p, i) => {
    if (p === 'z') return { close: true };
    if (p[0] === 'c') {
      return { x: p[5] * w, y: p[6] * h, curve: { type: 'cubic', x1: p[1] * w, y1: p[2] * h, x2: p[3] * w, y2: p[4] * h } };
    }
    return { x: p[0] * w, y: p[1] * h, moveTo: i === 0 };
  });
  slide.addShape('custGeom', Object.assign({ x, y, w, h, points }, o));
}

/** The chevron arrow used by the "Next Page" / "More Information" buttons. */
const ARROW_PTS = [
  [0.761, 0.436], [0.425, 0.091], [0.514, 0], [1, 0.5], [0.514, 1],
  [0.425, 0.909], [0.761, 0.564], [0, 0.564], [0, 0.436], 'z',
];

function arrowGlyph(slide, x, y, w, h, color) {
  poly(slide, x, y, w, h, ARROW_PTS, { fill: { color } });
}

/** Sutherland-Hodgman clip of a convex polygon against the half-plane f(p) <= 0. */
function clipHalfPlane(pts, f) {
  const out = [];
  pts.forEach((p, i) => {
    const q = pts[(i + 1) % pts.length];
    const fp = f(p);
    const fq = f(q);
    if (fp <= 0) out.push(p);
    if ((fp < 0 && fq > 0) || (fp > 0 && fq < 0)) {
      const t = fp / (fp - fq);
      out.push([p[0] + (q[0] - p[0]) * t, p[1] + (q[1] - p[1]) * t]);
    }
  });
  return out;
}

/** `hex` scaled towards black — the deck's gradients all sit on a black slide. */
function shade(hex, k) {
  return [0, 2, 4].map((i) => Math.round(parseInt(hex.substr(i, 2), 16) * k).toString(16).padStart(2, '0')).join('').toUpperCase();
}

/**
 * Brightness of the deck's 45-degree ramp at diagonal fraction `s`, where s=0 is
 * the top-left corner. `teal` runs accent1 -> black; `glass` runs black -> 30%
 * white, brightest at the bottom-right.
 */
function rampLevel(isTeal, s) {
  if (!isTeal) return 0.3 * s;
  const t = Math.min(s / 0.9, 1);
  return (1 - t) * (1 - 0.3 * t);
}

/**
 * Rounded card carrying that gradient. pptxgenjs writes solid fills only, so the
 * ramp is painted as opaque slices perpendicular to the diagonal: each slice
 * covers everything from its own edge outwards, so they nest and never seam.
 */
const BANDS = 14;

function gradientSlices(slide, x, y, w, h, outline, isTeal, mirror, o) {
  const peak = isTeal ? C.teal : C.white;
  for (let k = 0; k < BANDS; k++) {
    const edge = (isTeal ? BANDS - k : k) / BANDS; // slice boundary along the diagonal
    const level = rampLevel(isTeal, (isTeal ? BANDS - k - 0.5 : k + 0.5) / BANDS);
    let pts = k === 0 ? outline : clipHalfPlane(outline, (p) => (isTeal ? 1 : -1) * (p[0] + p[1] - 2 * edge));
    if (pts.length < 3) continue;
    if (mirror) pts = pts.map((p) => [1 - p[0], p[1]]);
    poly(slide, x, y, w, h, pts, Object.assign({}, o, { fill: { color: shade(peak, level) }, line: null }));
  }
}

function card(slide, x, y, w, h, radius, tone, mirror) {
  const chX = radius / w; // corners are chamfered so slices follow the rounding
  const chY = radius / h;
  const outline = [[chX, 0], [1 - chX, 0], [1, chY], [1, 1 - chY], [1 - chX, 1], [chX, 1], [0, 1 - chY], [0, chY]];
  gradientSlices(slide, x, y, w, h, outline, tone === 'teal', mirror, null);
  slide.addShape('roundRect', {
    x, y, w, h,
    rectRadius: radius,
    fill: { type: 'none' },
    line: { color: C.white, transparency: 50, width: 1 },
  });
}

/**
 * Free-form panel in the same teal glass as `card`. Curves cannot be clipped
 * against the diagonal, so its convex hull carries the ramp and the outline
 * itself is drawn last as the border.
 */
function glassPoly(slide, x, y, w, h, pts, hull, o) {
  gradientSlices(slide, x, y, w, h, hull, true, !!o.flipH, o);
  poly(slide, x, y, w, h, pts,
    Object.assign({}, o, { fill: { type: 'none' }, line: { color: C.white, transparency: 50, width: 1 } }));
}

/** Stand-in for a bitmap in the source deck. */
function imagePlaceholder(slide, x, y, w, h, radius) {
  slide.addShape('roundRect', {
    x, y, w, h,
    rectRadius: radius || 0.25,
    fill: { color: C.ink },
    line: { color: C.dim, width: 1 },
  });
  text(slide, '[image]', { x, y: y + h / 2 - 0.2, w, h: 0.4, align: 'center', fontSize: 12, color: C.dim });
}

/* ------------------------------------------------------------------- icons */

/**
 * Line-art icons drawn from primitives inside a 0..1 unit box:
 *   ['d', cx, cy, d]      filled disc      ['o', cx, cy, d]      ring
 *   ['r', x, y, w, h]     filled rect      ['b', x, y, w, h]     rect outline
 *   ['R', x, y, w, h]     rounded rect     ['t', x, y, w, h, ch] glyph
 */
const ICONS = {
  target: [['o', 0.5, 0.5, 0.94], ['o', 0.5, 0.5, 0.56], ['d', 0.5, 0.5, 0.2]],
  gauge: [['o', 0.5, 0.5, 0.94], ['r', 0.47, 0.2, 0.06, 0.34], ['d', 0.5, 0.62, 0.16]],
  money: [['b', 0.04, 0.24, 0.92, 0.52], ['o', 0.5, 0.5, 0.3], ['r', 0.1, 0.44, 0.06, 0.12], ['r', 0.84, 0.44, 0.06, 0.12]],
  book: [['b', 0.06, 0.16, 0.88, 0.68], ['r', 0.47, 0.16, 0.06, 0.68]],
  safe: [['b', 0.06, 0.12, 0.88, 0.76], ['o', 0.5, 0.5, 0.42], ['r', 0.47, 0.42, 0.06, 0.16]],
  board: [['b', 0.06, 0.1, 0.88, 0.6], ['r', 0.47, 0.7, 0.06, 0.22], ['r', 0.2, 0.86, 0.6, 0.06]],
  bulb: [['o', 0.5, 0.38, 0.6], ['r', 0.38, 0.68, 0.24, 0.08], ['r', 0.4, 0.82, 0.2, 0.08]],
  gear: [['o', 0.5, 0.5, 0.9], ['d', 0.5, 0.5, 0.3], ['r', 0.45, 0, 0.1, 0.16], ['r', 0.45, 0.84, 0.1, 0.16]],
  people: [['d', 0.28, 0.24, 0.3], ['d', 0.72, 0.24, 0.3], ['R', 0.08, 0.46, 0.38, 0.42], ['R', 0.54, 0.46, 0.38, 0.42]],
  rocket: [['o', 0.5, 0.36, 0.52], ['r', 0.42, 0.06, 0.16, 0.2], ['r', 0.1, 0.74, 0.8, 0.08], ['d', 0.5, 0.36, 0.16]],
  diamond: [['b', 0.1, 0.34, 0.8, 0.52], ['r', 0.3, 0.14, 0.4, 0.12], ['d', 0.5, 0.6, 0.22]],
  cloud: [['o', 0.5, 0.4, 0.8], ['r', 0.14, 0.5, 0.72, 0.1], ['r', 0.42, 0.62, 0.16, 0.3]],
  handshake: [['r', 0.06, 0.4, 0.4, 0.16], ['r', 0.54, 0.4, 0.4, 0.16], ['d', 0.5, 0.48, 0.34], ['r', 0.2, 0.68, 0.6, 0.08]],
  hand: [['o', 0.5, 0.34, 0.5], ['r', 0.06, 0.66, 0.88, 0.1], ['r', 0.24, 0.8, 0.52, 0.08]],
  clipboard: [['b', 0.14, 0.1, 0.72, 0.82], ['r', 0.34, 0.02, 0.32, 0.14], ['r', 0.28, 0.36, 0.44, 0.07], ['r', 0.28, 0.56, 0.44, 0.07]],
  trophy: [['b', 0.24, 0.06, 0.52, 0.44], ['r', 0.44, 0.5, 0.12, 0.24], ['r', 0.22, 0.78, 0.56, 0.12], ['o', 0.5, 0.24, 0.9]],
  briefcase: [['b', 0.04, 0.26, 0.92, 0.6], ['r', 0.34, 0.1, 0.32, 0.12], ['r', 0.44, 0.44, 0.12, 0.2]],
  key: [['o', 0.24, 0.42, 0.44], ['r', 0.42, 0.38, 0.54, 0.09], ['r', 0.78, 0.46, 0.08, 0.16], ['r', 0.2, 0.74, 0.66, 0.08]],
  chart: [['b', 0.08, 0.08, 0.84, 0.56], ['r', 0.24, 0.28, 0.08, 0.24], ['r', 0.46, 0.2, 0.08, 0.32], ['r', 0.68, 0.34, 0.08, 0.18], ['r', 0.2, 0.78, 0.6, 0.08]],
  news: [['b', 0.02, 0.1, 0.96, 0.8], ['r', 0.12, 0.24, 0.36, 0.26], ['r', 0.56, 0.24, 0.32, 0.07], ['r', 0.56, 0.42, 0.32, 0.07], ['r', 0.12, 0.62, 0.76, 0.07]],
  cast: [['b', 0.04, 0.1, 0.92, 0.66], ['o', 0.3, 0.5, 0.5], ['d', 0.16, 0.7, 0.14]],
  phone: [['R', 0.24, 0.02, 0.52, 0.96], ['r', 0.38, 0.1, 0.24, 0.05], ['d', 0.5, 0.86, 0.14]],
  pin: [['o', 0.5, 0.36, 0.72], ['d', 0.5, 0.36, 0.24], ['r', 0.46, 0.6, 0.08, 0.4]],
};

function icon(slide, name, cx, cy, size, color) {
  const stroke = Math.max(0.012, size * 0.075);
  (ICONS[name] || []).forEach((p) => {
    const X = (u) => cx - size / 2 + u * size;
    const Y = (u) => cy - size / 2 + u * size;
    if (p[0] === 'd') disc(slide, X(p[1]), Y(p[2]), p[3] * size, color);
    else if (p[0] === 'o') ring(slide, X(p[1]), Y(p[2]), p[3] * size, color, stroke * 72);
    else if (p[0] === 'r') rect(slide, X(p[1]), Y(p[2]), p[3] * size, p[4] * size, { fill: { color } });
    else if (p[0] === 'b') rect(slide, X(p[1]), Y(p[2]), p[3] * size, p[4] * size, { fill: { type: 'none' }, line: { color, width: stroke * 72 } });
    else if (p[0] === 'R') {
      slide.addShape('roundRect', { x: X(p[1]), y: Y(p[2]), w: p[3] * size, h: p[4] * size, rectRadius: size * 0.1, fill: { color } });
    }
  });
}

/** Coloured disc with a white icon on top — used on almost every infographic. */
function iconBadge(slide, name, cx, cy, d, fill, glyphColor) {
  disc(slide, cx, cy, d, fill);
  icon(slide, name, cx, cy, d * 0.56, glyphColor || C.white);
}

/* ------------------------------------------------------------ slide builders */

function slide01(s) {
  title(s, 'Cybersecurity Essentials', {
    x: 0.514, y: 2.804, w: 12.306, h: 3.198, fontSize: 115, align: 'center', lineSpacingMultiple: 0.8,
  });
}

function slide02(s) {
  title(s, 'Innovative Cybersecurity Strategies For A Safer Connected World', { x: 0.608, y: 0.489, w: 12.024, h: 1.582 });
  const cols = [
    { x: 0.715, tone: 'teal', n: '01', h: 1.648, copy: 'PLACEHOLDER' },
    { x: 4.789, tone: 'glass', n: '02', h: 1.255, copy: 'Smart Cybersecurity Solutions to Secure the Digital Economy' },
    { x: 8.863, tone: 'glass', n: '03', h: 1.648, copy: 'Empowering Organizations with Next-Generation Cybersecurity Tools' },
  ];
  cols.forEach((c) => {
    card(s, c.x, 2.565, 3.755, 3.437, 0.114, c.tone);
    text(s, c.n, { x: c.x + 0.301, y: 2.774, w: 1.203, h: 1.01, fontFace: FONT.head, fontSize: 54 });
    text(s, c.copy, { x: c.x + 0.301, y: 4.014, w: 3.064, h: c.h, lineSpacingMultiple: 1.3 });
  });
}

function slide03(s) {
  card(s, 7.861, 0.622, 4.962, 5.422, 0.149, 'teal');
  title(s, 'Strengthening Global Cybersecurity for Digital Trust', { x: 0.395, y: 0.434, w: 4.445, h: 3.803 });
  subhead(s, 'Protecting Data Integrity and Privacy Worldwide', { x: 8.792, y: 1.348, w: 3.653, h: 0.64 });
  body(s, LOREM.labore, { x: 8.799, y: 2.042, w: 3.458, h: 0.87 });
  subhead(s, 'Building Resilience Against Evolving Cyber Threats', { x: 8.799, y: 3.544, w: 3.653, h: 0.64 });
  body(s, LOREM.labore, { x: 8.806, y: 4.238, w: 3.458, h: 0.87 });

  rect(s, 0.512, 4.47, 2.142, 0.583, { fill: { color: C.teal } });
  text(s, 'Next Page', { x: 0.711, y: 4.604, w: 1.197, h: 0.337, fontSize: 14 });
  arrowGlyph(s, 2.096, 4.619, 0.316, 0.307, C.white);
}

function slide04(s) {
  title(s, 'Next-Generation Cybersecurity Solutions', { x: 0.522, y: 0.619, w: 7.426, h: 1.582 });
  [
    { x: 8.182, tone: 'teal', pct: '88%' },
    { x: 10.414, tone: 'glass', pct: '91%' },
  ].forEach((k) => {
    card(s, k.x, 0.759, 2.119, 1.363, 0.105, k.tone);
    text(s, k.pct, { x: k.x + 0.38, y: 0.909, w: 1.417, h: 0.572, fontFace: FONT.head, fontSize: 28, align: 'center' });
    text(s, 'Data Protection Assurance', {
      x: k.x + 0.056, y: 1.387, w: 1.976, h: 0.505, fontFace: FONT.head, fontSize: 12, align: 'center',
    });
  });
}

function slide05(s) {
  title(s, 'From Digital Threats to Digital Trust', { x: 0.441, y: 0.521, w: 6.573, h: 1.582 });
  body(s, LOREM.laboreTempor, { x: 0.483, y: 2.339, w: 5.878, h: 0.608 });
  [
    { x: 0.538, ic: 'news', ix: 0.969, iy: 4.18 },
    { x: 4.858, ic: 'cast', ix: 5.334, iy: 4.242 },
  ].forEach((k, i) => {
    card(s, k.x, 3.75, 4.153, 1.569, 0.095, 'teal');
    icon(s, k.ic, k.ix, k.iy, 0.4, C.white);
    subhead(s, 'Your Text Here', { x: k.x + 0.978, y: 4.02 + i * 0.013, w: 1.889, h: 0.347, fontSize: 14 });
    body(s, LOREM.eiusmod, { x: k.x + 0.12, y: 4.492 + i * 0.013, w: 3.913, h: 0.608 });
  });
}

function slide06(s) {
  title(s, 'Meet Our Team 2025', { x: 3.54, y: 0.438, w: 6.253, h: 0.841, align: 'center' });
  [
    { x: 1.102, n: '01', dy: 0, mirror: true },
    { x: 9.201, n: '02', dy: 0.031, mirror: false },
  ].forEach((k) => {
    card(s, k.x, 3.315, 3.03, 2.773, 0.092, 'teal', k.mirror);
    text(s, k.n, { x: k.x + 0.206, y: 3.479 + k.dy, w: 1.203, h: 1.01, fontFace: FONT.head, fontSize: 54 });
    subhead(s, 'Olivia Johnson', { x: k.x + 0.229, y: 4.479 + k.dy + 0.01, w: 2.181, h: 0.385 });
    body(s, LOREM.sedDo, { x: k.x + 0.236, y: 4.851 + k.dy + 0.009, w: 2.621, h: 0.87 });
  });
}

function slide07(s) {
  card(s, 6.722, 3.432, 2.544, 2.221, 0.073, 'teal');
  card(s, 9.572, 3.432, 2.544, 2.221, 0.073, 'glass');
  imagePlaceholder(s, 1.401, 1.194, 4.34, 5.4, 0.5); // phone mock-up photo
  title(s, 'Mobile App Mockup', { x: 6.667, y: 0.829, w: 5.985, h: 0.841 });
  body(s, LOREM.exercitation, { x: 6.667, y: 1.85, w: 5.401, h: 0.87 });
  [6.722, 9.572].forEach((cx, i) => {
    text(s, '459,769K', {
      x: cx + 0.331, y: 3.806, w: 1.881, h: 0.533, fontFace: FONT.head, fontSize: 31.65, valign: 'middle', wrap: false,
    });
    body(s, LOREM.nostrud, { x: [6.95, 9.798][i], y: 4.403, w: 2.091, h: 0.87 });
  });
}

/** Slide 8: four-wedge donut built from native PIE shapes. */
function donut(s) {
  const wedges = [
    { x: 4.417, y: 1.681, d: 4.498, a: [89.9, 180.3], color: C.teal },
    { x: 4.579, y: 1.806, d: 4.242, a: [180.0, 270.0], color: C.blue },
    { x: 4.543, y: 1.807, d: 4.246, a: [315.3, 90.3], color: C.green },
    { x: 5.025, y: 2.289, d: 3.283, a: [271.0, 316.5], color: C.teal },
  ];
  wedges.forEach((w) => s.addShape('pie', { x: w.x, y: w.y, w: w.d, h: w.d, angleRange: w.a, fill: { color: w.color } }));

  disc(s, 6.7, 3.927, 1.857, C.bone);
  disc(s, 6.7, 3.927, 1.406, C.bone);
  disc(s, 6.7, 3.927, 1.021, C.navy);
  text(s, 'S', { x: 6.2, y: 3.55, w: 1.0, h: 0.75, fontFace: FONT.head, fontSize: 26, bold: true, color: C.teal, align: 'center', valign: 'middle' });

  const labels = [
    { t: '16.7%', x: 6.701, y: 2.548, w: 0.934, h: 0.444, sz: 11 },
    { t: '29.5%', x: 7.45, y: 4.014, w: 1.315, h: 0.642, sz: 20 },
    { t: '28.2%', x: 4.777, y: 4.535, w: 1.492, h: 0.642, sz: 20 },
    { t: '25.6%', x: 5.061, y: 2.677, w: 1.127, h: 0.598, sz: 18 },
  ];
  labels.forEach((l) => {
    text(s, l.t, {
      x: l.x, y: l.y, w: l.w, h: l.h, fontFace: FONT.head, fontSize: l.sz, bold: true,
      align: 'center', valign: 'middle', lineSpacingMultiple: 1.5,
    });
  });
}

function slide08(s) {
  const quads = [
    { cx: 1.99, cy: 1.537, card: 0.792, cardY: 1.383, icon: 'people', fill: C.green, tx: 1.421, ty: 2.416, bx: 1.061, by: 2.769 },
    { cx: 1.99, cy: 4.179, card: 0.792, cardY: 3.994, icon: 'rocket', fill: C.teal, tx: 1.437, ty: 5.103, bx: 1.077, by: 5.455 },
    { cx: 10.58, cy: 1.537, card: 9.362, cardY: 1.333, icon: 'diamond', fill: C.teal, tx: 10.007, ty: 2.448, bx: 9.647, by: 2.801 },
    { cx: 10.58, cy: 4.179, card: 9.362, cardY: 3.944, icon: 'gear', fill: C.green, tx: 10.007, ty: 5.061, bx: 9.647, by: 5.414 },
  ];
  quads.forEach((q) => {
    card(s, q.card, q.cardY, 3.18, 2.329, 0.141, 'teal');
    iconBadge(s, q.icon, q.cx + 0.372, q.cy + 0.372, 0.744, q.fill);
    subhead(s, 'Your Text Here', { x: q.tx, y: q.ty, w: 1.889, h: 0.37, align: 'center' });
    body(s, LOREM.consectetur, { x: q.bx, y: q.by, w: 2.609, h: 0.608, align: 'center' });
  });
  donut(s);
  title(s, 'Section Infographic.', { x: 3.326, y: 0.391, w: 6.972, h: 0.767, align: 'center', lineSpacingMultiple: 0.9 });
}

function slide09(s) {
  const bars = [
    { x: 1.063, top: 4.79, v: '+25.2', color: C.teal },
    { x: 2.063, top: 4.673, v: '+32.7', color: C.green },
    { x: 3.064, top: 4.27, v: '+37.5', color: C.teal },
    { x: 4.064, top: 3.888, v: '+50.9', color: C.green },
    { x: 5.064, top: 3.241, v: '+72.4', color: C.teal },
  ];
  // rising trend arrow behind the columns
  poly(s, 1.063, 2.001, 4.652, 2.241, [
    [0, 1], ['c', 0.35, 0.9, 0.6, 0.63, 0.78, 0.29], [0.7, 0.16], [1, 0.03], [0.93, 0.62], [0.84, 0.44],
    ['c', 0.6, 0.78, 0.3, 0.96, 0.03, 1], 'z',
  ], { fill: { color: C.navy } });

  bars.forEach((b) => {
    s.addShape('roundRect', {
      x: b.x, y: b.top, w: 0.876, h: 6.149 - b.top, rectRadius: 0.084, fill: { color: b.color },
    });
    text(s, b.v, {
      x: b.x, y: 5.533, w: 0.876, h: 0.466, fontFace: FONT.head, fontSize: 12, bold: true,
      align: 'center', lineSpacingMultiple: 1.5,
    });
  });
  s.addShape('roundRect', { x: 0.738, y: 6.079, w: 5.526, h: 0.07, rectRadius: 0.035, fill: { color: C.navy } });

  const rows = [
    { y: 0.908, tone: 'teal', icon: 'cloud', badge: C.bone, glyph: C.green, iy: 1.168, ty: 1.163, by: 1.52, tx: 8.416, bx: 8.514 },
    { y: 2.714, tone: 'glass', icon: 'handshake', badge: C.teal, glyph: C.white, iy: 2.962, ty: 2.994, by: 3.352, tx: 8.402, bx: 8.5 },
    { y: 4.496, tone: 'glass', icon: 'hand', badge: C.teal, glyph: C.white, iy: 4.784, ty: 4.764, by: 5.121, tx: 8.388, bx: 8.486 },
  ];
  rows.forEach((r) => {
    card(s, 6.961, r.y, 5.51, 1.528, 0.092, r.tone);
    iconBadge(s, r.icon, 7.69, r.iy + 0.505, 1.009, r.badge, r.glyph);
    subhead(s, 'Your Text Here', { x: r.tx, y: r.ty, w: 1.889, h: 0.37, align: 'center' });
    body(s, LOREM.tempor, { x: r.bx, y: r.by, w: 3.639, h: 0.608 });
  });
  title(s, 'Section Infographic.', { x: 0.93, y: 0.516, w: 4.362, h: 1.434, lineSpacingMultiple: 0.9 });
}

/** Slide 10: magnifier whose lens is split into six labelled wedges. */
function magnifier(s) {
  const cx = 6.679;
  const cy = 3.787;
  const lens = 3.484;

  // dashed orbit around the lens
  [[246.6, 293.0], [66.9, 113.7], [126.9, 173.5], [186.6, 232.9], [6.9, 53.3], [306.6, 353.3]].forEach((a) => {
    s.addShape('arc', { x: 4.557, y: 1.668, w: 4.243, h: 4.243, angleRange: a, fill: { type: 'none' }, line: { color: '44546A', width: 2 } });
  });
  [[5.601, 1.946, C.teal], [7.734, 1.946, C.green], [4.535, 3.791, C.blue], [8.8, 3.791, C.teal],
   [5.601, 5.634, C.teal], [7.734, 5.634, C.green]].forEach((m) => {
    ring(s, m[0], m[1], 0.4, m[2], 3);
    disc(s, m[0], m[1], 0.157, m[2]);
  });

  // handle (light collar over a dark shaft), then the white lens body
  rect(s, 6.44, 5.6, 0.46, 0.45, { fill: { color: 'F2F2F2' } });
  rect(s, 6.44, 6.0, 0.46, 1.5, { fill: { color: C.ink } });
  disc(s, cx, cy, lens + 0.34, C.white);

  const wedges = [
    { a: [210, 270], color: C.teal, icon: 'target' },
    { a: [270, 330], color: C.green, icon: 'gauge' },
    { a: [330, 30], color: C.teal, icon: 'book' },
    { a: [30, 90], color: C.green, icon: 'board' },
    { a: [90, 150], color: C.teal, icon: 'safe' },
    { a: [150, 210], color: C.blue, icon: 'money' },
  ];
  wedges.forEach((w) => {
    s.addShape('pie', {
      x: cx - lens / 2, y: cy - lens / 2, w: lens, h: lens, angleRange: w.a,
      fill: { color: w.color }, line: { color: C.white, width: 4 },
    });
    const mid = ((w.a[0] + (w.a[1] < w.a[0] ? w.a[1] + 360 : w.a[1])) / 2) * (Math.PI / 180);
    icon(s, w.icon, cx + Math.cos(mid) * lens * 0.28, cy + Math.sin(mid) * lens * 0.28, 0.46, C.white);
  });
}

function slide10(s) {
  magnifier(s);
  const left = [
    { ty: 1.489, tx: 2.006, by: 1.81, bx: 0.408, bw: 3.793, bh: 0.87, copy: LOREM.aenean },
    { ty: 3.253, tx: 1.702, by: 3.573, bx: 0.658, bw: 3.239, bh: 0.608, copy: LOREM.aeneanShort },
    { ty: 5.011, tx: 2.102, by: 5.331, bx: 0.503, bw: 3.793, bh: 0.87, copy: LOREM.aenean },
  ];
  const right = [
    { ty: 1.439, tx: 9.098, by: 1.859, bx: 9.094, bw: 3.793, bh: 0.87, copy: LOREM.aenean },
    { ty: 3.363, tx: 9.5, by: 3.714, bx: 9.514, bw: 3.239, bh: 0.608, copy: LOREM.aeneanShort },
    { ty: 4.911, tx: 9.096, by: 5.331, bx: 9.091, bw: 3.793, bh: 0.87, copy: LOREM.aenean },
  ];
  left.forEach((r) => {
    subhead(s, 'Your Title Here', { x: r.tx, y: r.ty, w: 2.172, h: 0.37, align: 'right' });
    body(s, r.copy, { x: r.bx, y: r.by, w: r.bw, h: r.bh, align: 'right' });
  });
  right.forEach((r) => {
    subhead(s, 'Your Title Here', { x: r.tx, y: r.ty, w: 2.172, h: 0.37 });
    body(s, r.copy, { x: r.bx, y: r.by, w: r.bw, h: r.bh, align: 'left' });
  });
  title(s, 'Section Infographic.', { x: 3.326, y: 0.391, w: 6.972, h: 0.767, align: 'center', lineSpacingMultiple: 0.9 });
}

function slide11(s) {
  const cx = 6.63;
  const cy = 6.287; // flat bottom edge of the arch
  const outer = 2.739; // radius of the coloured band; the white shell is slightly wider
  // solid white half-disc shell, then three coloured step bands ringed on top
  s.addShape('pie', { x: cx - 2.942, y: cy - 2.942, w: 5.884, h: 5.884, angleRange: [180, 360], fill: { color: C.white } });
  [
    { a: [182, 237], color: C.teal }, { a: [243, 297], color: C.green }, { a: [303, 358], color: C.teal },
  ].forEach((seg) => {
    s.addShape('blockArc', {
      x: cx - outer, y: cy - outer, w: outer * 2, h: outer * 2, angleRange: seg.a,
      arcThicknessRatio: (outer - 1.732) / outer, fill: { color: seg.color },
    });
  });
  // step captions ride the band mid-line, rotated tangentially
  [{ t: 'STEP 01', a: 209.5 }, { t: 'STEP 02', a: 270 }, { t: 'STEP 03', a: 330.5 }].forEach((k) => {
    const rad = (k.a * Math.PI) / 180;
    text(s, k.t, {
      x: cx + 2.19 * Math.cos(rad) - 1.074, y: cy + 2.19 * Math.sin(rad) - 0.537,
      w: 2.147, h: 1.074, rotate: (k.a + 90) % 360, fontFace: FONT.head, fontSize: 32, bold: true,
      align: 'center', valign: 'middle', lineSpacingMultiple: 1.3,
    });
  });

  const spokes = [
    { icon: 'target', cx: 2.022, cy: 3.529, fill: C.teal, tx: 0.976, ty: 4.129, bx: 0.771, by: 4.499 },
    { icon: 'bulb', cx: 6.774, cy: 1.025, fill: C.green, tx: 5.688, ty: 1.594, bx: 5.482, by: 1.965 },
    { icon: 'gear', cx: 11.249, cy: 3.573, fill: C.teal, tx: 10.162, ty: 4.123, bx: 9.956, by: 4.493 },
  ];
  spokes.forEach((k) => {
    iconBadge(s, k.icon, k.cx, k.cy, 0.769, k.fill);
    subhead(s, 'Your Title Here', { x: k.tx, y: k.ty, w: 2.172, h: 0.37, align: 'center' });
    body(s, LOREM.aeneanLigula, { x: k.bx, y: k.by, w: 2.583, h: 0.87, align: 'center' });
  });
  title(s, 'Section Infographic.', { x: 0.7, y: 0.576, w: 4.362, h: 1.434, lineSpacingMultiple: 0.9 });
}

/** Simplified stand-in for the flat-illustration characters in the source deck. */
function figure(s, x, yHead, h, shirt) {
  disc(s, x, yHead, h * 0.2, C.skin);
  poly(s, x - h * 0.09, yHead - h * 0.13, h * 0.18, h * 0.14, [[0, 1], ['c', 0, 0, 1, 0, 1, 1], 'z'], { fill: { color: C.ink } });
  s.addShape('roundRect', {
    x: x - h * 0.16, y: yHead + h * 0.08, w: h * 0.32, h: h * 0.4, rectRadius: h * 0.09, fill: { color: shirt },
  });
  rect(s, x - h * 0.14, yHead + h * 0.44, h * 0.12, h * 0.44, { fill: { color: C.ink } });
  rect(s, x + h * 0.02, yHead + h * 0.44, h * 0.12, h * 0.44, { fill: { color: C.ink } });
}

/** Lens-shaped leaf, used for the foliage clusters in the source illustrations. */
function leaf(s, x, y, w, h, color, rotate) {
  poly(s, x, y, w, h, [[0, 1], ['c', 0.05, 0.35, 0.4, 0.05, 1, 0], ['c', 0.6, 0.42, 0.3, 0.72, 0, 1], 'z'],
    { fill: { color }, rotate });
}

function slide12(s) {
  s.addShape('ellipse', { x: 0.48, y: 0.835, w: 5.673, h: 5.426, fill: { color: C.soot } });
  [[3.999, C.white], [3.413, C.blue], [2.251, C.white], [2.08, C.blue], [0.977, C.white], [0.806, C.blue]]
    .forEach((r) => disc(s, 2.9295, 3.8345, r[0], r[1]));
  poly(s, 2.929, 1.332, 2.18, 2.18, [[0.93, 0], [1, 0.07], [0.07, 1], [0, 0.93], 'z'], { fill: { color: C.ink } });
  poly(s, 4.39, 1.33, 0.72, 0.72, [[1, 0], [1, 0.65], [0.62, 0.38], [0.35, 0], 'z'], { fill: { color: C.ink } });

  figure(s, 1.45, 3.49, 2.85, C.grey);
  figure(s, 5.37, 3.43, 2.75, C.ink);
  figure(s, 4.25, 4.05, 2.0, C.grey);
  rect(s, 3.03, 5.10, 2.42, 0.09, { fill: { color: C.green } });
  [3.19, 5.13].forEach((lx) => {
    poly(s, lx, 5.19, 0.34, 0.79, [[0, 0], [0.08, 0], [1, 1], [0.92, 1], 'z'], { fill: { color: C.green } });
    poly(s, lx, 5.19, 0.34, 0.79, [[0.92, 0], [1, 0], [0.08, 1], [0, 1], 'z'], { fill: { color: C.green } });
  });

  const stats = [
    { cx: 7.075, cy: 2.478, icon: 'clipboard', badge: C.teal, val: '$ 59.6m', vx: 7.468, vy: 1.942, vw: 1.897, vc: C.teal, bx: 6.704, by: 2.979 },
    { cx: 10.032, cy: 2.478, icon: 'board', badge: C.green, val: '$ 42.8m', vx: 10.56, vy: 1.914, vw: 2.023, vc: C.green, bx: 9.765, by: 2.979 },
    { cx: 7.075, cy: 4.49, icon: 'money', badge: C.green, val: '$ 39.5m', vx: 7.468, vy: 4.006, vw: 1.897, vc: C.green, bx: 6.718, by: 5.01 },
    { cx: 10.032, cy: 4.49, icon: 'safe', badge: C.teal, val: '$ 62.7m', vx: 10.646, vy: 4.006, vw: 1.897, vc: C.teal, bx: 9.778, by: 5.01 },
  ];
  stats.forEach((k) => {
    iconBadge(s, k.icon, k.cx, k.cy, 0.806, k.badge);
    text(s, k.val, {
      x: k.vx, y: k.vy, w: k.vw, h: 0.817, fontFace: FONT.head, fontSize: 28, color: k.vc,
      valign: 'middle', lineSpacingMultiple: 1.5,
    });
    body(s, LOREM.adipiscing, { x: k.bx, y: k.by, w: 2.661, h: 0.608, align: 'left' });
  });
  title(s, 'Section Infographic.', { x: 6.274, y: 0.75, w: 6.548, h: 0.767, lineSpacingMultiple: 0.9 });
}

function slide13(s) {
  s.addShape('ellipse', { x: 0.541, y: 1.706, w: 4.931, h: 4.715, fill: { color: C.soot } });
  [[1.867, C.white], [1.806, C.blue], [1.173, C.white], [1.1, C.blue], [0.506, C.white], [0.427, C.blue]]
    .forEach((r) => disc(s, 4.0145, 3.0975, r[0], r[1]));
  poly(s, 4.173, 1.781, 0.945, 1.149, [[0.9, 0], [1, 0.08], [0.1, 1], [0, 0.92], 'z'], { fill: { color: C.navy } });
  poly(s, 4.812, 1.645, 0.444, 0.613, [[1, 0], [0.86, 0.72], [0, 1], [0.16, 0.28], 'z'], { fill: { color: '70AD47' } });

  // foliage fan behind the runner
  [[0.75, 3.36, 0.99, 2.35, C.teal, 12], [1.21, 3.49, 0.52, 2.17, 'ABD3D9', 6],
   [3.64, 2.70, 1.73, 3.27, C.green, 4], [1.68, 3.97, 1.67, 1.99, C.green, 22],
   [0.75, 4.63, 1.49, 1.35, '97D2C3', 34], [4.47, 2.80, 0.84, 1.03, '97D2C3', 8],
   [2.92, 4.22, 1.26, 1.75, '97D2C3', 16]].forEach((l) => leaf(s, l[0], l[1], l[2], l[3], l[4], l[5]));

  // coins
  [[2.664, 2.293, 0.389], [3.13, 1.806, 0.678], [3.401, 2.577, 0.699], [3.871, 1.951, 0.511]].forEach((c) => {
    disc(s, c[0], c[1], c[2], C.white);
    disc(s, c[0], c[1], c[2] * 0.85, C.gold);
    text(s, '$', {
      x: c[0] - c[2] / 2, y: c[1] - c[2] / 2, w: c[2], h: c[2], fontFace: FONT.head, fontSize: c[2] * 44,
      bold: true, color: C.white, align: 'center', valign: 'middle',
    });
  });
  s.addShape('gear6', { x: 1.21, y: 2.441, w: 0.99, h: 0.988, fill: { color: C.teal } });
  s.addShape('gear6', { x: 0.681, y: 1.991, w: 0.551, h: 0.55, fill: { color: C.grey } });

  // trophy + runner
  poly(s, 2.645, 3.163, 0.997, 0.194, [[0, 0], [1, 0], [0.94, 1], [0.06, 1], 'z'], { fill: { color: 'FFD966' } });
  poly(s, 2.735, 3.35, 0.797, 0.84, [[0, 0], [1, 0], [0.8, 1], [0.2, 1], 'z'], { fill: { color: C.gold } });
  [[2.555, 3.707, 0.471], [3.247, 3.707, 0.47]].forEach((h) => {
    s.addShape('arc', { x: h[0], y: h[1], w: h[2], h: 0.43, angleRange: [270, 90], line: { color: 'FFD966', width: 6 } });
  });
  rect(s, 3.06, 4.19, 0.15, 0.24, { fill: { color: C.gold } });
  rect(s, 2.872, 4.4, 0.521, 0.14, { fill: { color: 'FFD966' } });
  figure(s, 2.15, 3.9, 2.3, C.grey);

  [
    { x: 5.818, h: 3.851, r: 0.192, tone: 'teal', icon: 'trophy', badge: C.teal, cx: 7.415, vx: 6.535, vy: 3.463, bx: 6.351, by: 4.389 },
    { x: 9.176, h: 3.835, r: 0.105, tone: 'glass', icon: 'target', badge: C.green, cx: 10.766, vx: 9.961, vy: 3.476, bx: 9.777, by: 4.401 },
  ].forEach((k) => {
    card(s, k.x, 1.937, 3.18, k.h, k.r, k.tone);
    iconBadge(s, k.icon, k.cx, 2.61, 0.863, k.badge);
    text(s, '$8,5M', {
      x: k.vx, y: k.vy, w: 1.765, h: 0.799, fontFace: FONT.head, fontSize: 47.48, valign: 'middle', wrap: false,
    });
    body(s, LOREM.nostrud, { x: k.bx, y: k.by, w: 2.091, h: 0.87, align: 'center' });
  });
  title(s, 'Section Infographic.', { x: 5.832, y: 0.581, w: 6.548, h: 0.767, align: 'center', lineSpacingMultiple: 0.9 });
}

function slide14(s) {
  const steps = [
    { x: 0.831, icon: 'briefcase', color: C.teal, shade: '48939D', n: '1', tx: 0.692, ty: 4.17, bx: 0.817, by: 4.771 },
    { x: 3.882, icon: 'key', color: C.green, shade: '3B8975', n: '2', tx: 3.811, ty: 4.154, bx: 3.936, by: 4.754 },
    { x: 6.933, icon: 'chart', color: C.teal, shade: '48939D', n: '3', tx: 6.804, ty: 4.19, bx: 6.929, by: 4.79 },
    { x: 9.984, icon: 'trophy', color: C.green, shade: '3B8975', n: '4', tx: 9.845, ty: 4.19, bx: 9.97, by: 4.79 },
  ];
  steps.forEach((k, i) => {
    s.addShape('roundRect', { x: k.x + 0.055, y: 1.986, w: 2.349, h: 1.707, rectRadius: 0.18, fill: { color: k.shade } });
    s.addShape('roundRect', { x: k.x, y: 2.098, w: 2.317, h: 1.677, rectRadius: 0.18, fill: { color: k.color } });
    icon(s, k.icon, k.x + 1.16, 2.88, 0.85, C.white);
    disc(s, k.tx + 0.256, 3.822, 0.511, i % 2 === 0 ? C.teal : C.green, { line: { color: C.white, width: 4.5 } });
    text(s, k.n, {
      x: k.tx, y: 3.566, w: 0.511, h: 0.511, fontSize: 14, bold: true, align: 'center', valign: 'middle',
    });
    text(s, 'Your title here', {
      x: k.tx, y: k.ty, w: 2.317, h: 0.601, fontFace: FONT.head, bold: true, lineSpacingMultiple: 1.5,
    });
    body(s, LOREM.edet, { x: k.bx, y: k.by, w: 2.317, h: 1.133 });
  });
  title(s, 'Section Infographic.', { x: 3.393, y: 0.554, w: 6.548, h: 0.767, align: 'center', lineSpacingMultiple: 0.9 });
}

const PANEL_PTS = [
  [0.6885, 0], [1, 0], [1, 1], [0.7202, 1], [0.7074, 0.9967],
  ['c', 0.5097, 0.9457, 0.1388, 0.8469, 0.0267, 0.8125],
  ['c', -0.0018, 0.7982, 0.0071, 0.7873, 0, 0.775],
  ['c', 0.0062, 0.6697, 0.0062, 0.3333, 0.0027, 0.2153],
  ['c', 0.0116, 0.206, 0.0031, 0.1954, 0.0267, 0.1847], 'z',
];

const PANEL_HULL = [
  [0, 0.798], [0.003, 0.215], [0.027, 0.185], [0.689, 0], [1, 0], [1, 1],
  [0.72, 1], [0.51, 0.946], [0.139, 0.847], [0.027, 0.813],
];

const BUBBLE_PTS = [
  [0.4429, 0.0891], [0.9942, 0.7937],
  ['c', 1.006, 0.8087, 0.9992, 0.828, 0.9791, 0.8367], [0.6148, 0.9957],
  ['c', 0.5947, 1.0045, 0.5689, 0.9994, 0.5571, 0.9844], [0.0058, 0.2799],
  ['c', -0.006, 0.2648, 0.0008, 0.2455, 0.0209, 0.2368], [0.3035, 0.1135],
  [0.3337, 0], [0.3577, 0.0898], [0.3852, 0.0778],
  ['c', 0.4053, 0.069, 0.4312, 0.0741, 0.4429, 0.0891], 'z',
];

const BUBBLE_HULL = [
  [0, 0.265], [0.334, 0], [0.431, 0.074], [1, 0.809], [0.595, 1], [0.006, 0.28],
];

function slide15(s) {
  // angled full-height panel (mirrored, so the slanted edge faces right)
  glassPoly(s, 0, 0, 5.806, 7.5, PANEL_PTS, PANEL_HULL, { flipH: true });

  title(s, 'What Our Clients Say', { x: 0.784, y: 1.268, w: 4.049, h: 1.582, color: C.white });
  body(s, LOREM.utEnim, { x: 0.87, y: 3.052, w: 3.391, h: 1.133 });

  // testimonial speech bubble
  glassPoly(s, 8.912, 2.388, 3.292, 4.408, BUBBLE_PTS, BUBBLE_HULL, { rotate: 300.3 });

  subhead(s, 'Sophia Bennett', { x: 8.961, y: 4.03, w: 2.917, h: 0.404, fontSize: 18 });
  body(s, LOREM.labore2, { x: 8.961, y: 4.446, w: 3.391, h: 0.87 });
  text(s, '\u201C', { x: 11.877, y: 3.597, w: 0.769, h: 1.313, fontFace: FONT.quote, fontSize: 72 });

  s.addShape('roundRect', { x: 0.944, y: 4.856, w: 2.987, h: 0.638, rectRadius: 0.319, fill: { color: C.white } });
  subhead(s, 'More Information', { x: 1.2, y: 4.998, w: 2.103, h: 0.37, color: C.black });
  arrowGlyph(s, 3.308, 5.051, 0.256, 0.249, C.black);
}

function slide16(s) {
  title(s, 'Thank You', { x: 0.735, y: 2.148, w: 5.5, h: 2.822, fontSize: 115, lineSpacingMultiple: 0.7 });
  text(s, '\u201C', { x: 7.419, y: 0.889, w: 0.769, h: 1.313, fontFace: FONT.quote, fontSize: 72 });
  body(s, LOREM.utEnim, { x: 8.188, y: 0.889, w: 4.408, h: 0.87 });
  text(s, 'Contact Us :', { x: 0.842, y: 4.97, w: 2.375, h: 0.404, fontFace: FONT.head, bold: true });
  [
    { icon: 'phone', cx: 1.197, cy: 5.782, tx: 1.475, ty: 5.573, tw: 1.545, label: '+123 456 7890' },
    { icon: 'pin', cx: 3.358, cy: 5.786, tx: 3.719, ty: 5.568, tw: 1.932, label: '12 Your Street Name' },
  ].forEach((k) => {
    icon(s, k.icon, k.cx, k.cy, 0.26, C.white);
    text(s, k.label, { x: k.tx, y: k.ty, w: k.tw, h: 0.375, fontSize: 12, lineSpacingMultiple: 1.5 });
  });
}

/* -------------------------------------------------------------------- build */

const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
  slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
];

function build() {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.theme = { headFontFace: FONT.head, bodyFontFace: FONT.body };
  pptx.title = 'Cybersecurity Essentials';

  pptx.defineSlideMaster({
    title: 'DARK',
    background: { color: C.black },
    objects: [
      { line: { x: 0.431, y: 6.611, w: 12.472, h: 0, line: { color: C.white, transparency: 50, width: 1 } } },
      {
        text: {
          text: 'www.yourwebsitw.com',
          options: { x: 10.458, y: 6.806, w: 2.5, h: 0.337, align: 'right', valign: 'top', fontFace: FONT.body, fontSize: 14, color: C.dim },
        },
      },
    ],
  });

  BUILDERS.forEach((fn, i) => {
    const slide = pptx.addSlide({ masterName: 'DARK' });
    fn(slide);
    text(slide, 'Page ' + (i + 1), { x: 0.375, y: 6.806, w: 1.403, h: 0.336, fontSize: 14, color: C.dim });
  });

  return pptx.writeFile({ fileName: path.join(__dirname, '096b3ad8-5530-40d3-a78f-7081aeaee16e_grok_final.pptx') });
}

build().then((f) => console.log('wrote ' + f)).catch((e) => {
  console.error(e);
  process.exit(1);
});
