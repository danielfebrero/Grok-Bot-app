#!/usr/bin/env node
/**
 * "Patiesche" therapy pitch deck - 15 slides, 13.333 x 7.5 in.
 * Rebuilt with pptxgenjs only. Photographs in the source deck are replaced
 * by flat grey "[image]" placeholder blocks of the same size and position.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Palette / typography
 * ------------------------------------------------------------------ */
const PURPLE = '6B3DA2';   // theme accent1
const PEACH = 'EAA797';   // theme accent2
const WHITE = 'FFFFFF';
const INK = '0D0D0D';   // text1 @ 95% lumMod  (headings)
const INK_SOFT = '262626';   // text1 @ 85% lumMod
const MUTED = '7F7F7F';   // white @ 50% lumMod (body copy)
const SILVER = 'BFBFBF';   // white @ 75% lumMod (empty star)
const IMG_FILL = 'CECECE';   // photo placeholder block
const IMG_LABEL = 'E8E8E8';

const F_BOLD = 'Quicksand Bold';
const F_SEMI = 'Quicksand SemiBold';
const F_BODY = 'Roboto';
const F_NAV = 'Roboto SemiBold';

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

const NO_LINE = { type: 'none' };

/* Repeated body copy of the original deck */
const LOREM_FULL = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Curabitur cursus sapien eget egestas vestibulum. Nam fringilla neque varius turpis';
const LOREM_MED = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Curabitur cursus sapien eget egestas vestibulum. ';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Curabitur cursus sapien';
const LOREM_NEC = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. In id iaculis ligula. In nec';

/* ------------------------------------------------------------------ *
 * Geometry helpers - polygons + 45 deg banded gradients
 * (pptxgenjs has no gradient fill, so gradients are drawn as a stack of
 *  diagonal slices clipped to the shape outline)
 * ------------------------------------------------------------------ */

/** Corner-rounded rectangle as a point list. `r` may be a single radius or
 *  [topLeft, topRight, bottomRight, bottomLeft]. */
function roundRectPoints(x, y, w, h, r, seg = 8) {
  const rr = Array.isArray(r) ? r : [r, r, r, r];
  const corners = [
    [x + w - rr[1], y + rr[1], rr[1], -90],   // top-right
    [x + w - rr[2], y + h - rr[2], rr[2], 0],  // bottom-right
    [x + rr[3], y + h - rr[3], rr[3], 90],  // bottom-left
    [x + rr[0], y + rr[0], rr[0], 180], // top-left
  ];
  const pts = [];
  for (const [cx, cy, rad, a0] of corners) {
    if (rad <= 0) { pts.push([cx, cy]); continue; }
    for (let i = 0; i <= seg; i++) {
      const a = ((a0 + (90 * i) / seg) * Math.PI) / 180;
      pts.push([cx + rad * Math.cos(a), cy + rad * Math.sin(a)]);
    }
  }
  return pts;
}

/** Sutherland-Hodgman clip of a polygon against the half plane n.p <= d */
function clipHalfPlane(poly, nx, ny, d) {
  const out = [];
  for (let i = 0; i < poly.length; i++) {
    const a = poly[i];
    const b = poly[(i + 1) % poly.length];
    const da = nx * a[0] + ny * a[1] - d;
    const db = nx * b[0] + ny * b[1] - d;
    if (da <= 0) out.push(a);
    if ((da < 0 && db > 0) || (da > 0 && db < 0)) {
      const t = da / (da - db);
      out.push([a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t]);
    }
  }
  return out;
}

function mixHex(from, to, t) {
  const rgb = (h) => [0, 2, 4].map((i) => parseInt(h.substr(i, 2), 16));
  const a = rgb(from);
  const b = rgb(to);
  return a
    .map((v, i) => Math.round(v + (b[i] - v) * t).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
}

/** Draw a filled free-form polygon. */
function addPoly(slide, pts, opts) {
  slide.addShape('custGeom', Object.assign({
    x: 0, y: 0, w: SLIDE_W, h: SLIDE_H,
    points: pts.map((p) => ({ x: p[0], y: p[1] })).concat([{ close: true }]),
    line: NO_LINE,
  }, opts));
}

/** Fill a polygon with the deck's signature 45 deg purple -> peach gradient. */
function addGradPoly(slide, pts, opts = {}) {
  const from = opts.from || PURPLE;
  const to = opts.to || PEACH;
  const n = 1 / Math.SQRT2;                       // 45 deg (down-right) axis
  const proj = pts.map((p) => (p[0] + p[1]) * n);
  const lo = Math.min(...proj);
  const hi = Math.max(...proj);
  // one band every ~0.16in keeps the stepping invisible without bloating the file
  const bands = opts.bands || Math.max(8, Math.min(64, Math.round((hi - lo) / 0.16)));
  for (let i = 0; i < bands; i++) {
    const a = lo + ((hi - lo) * i) / bands;
    const b = lo + ((hi - lo) * (i + 1)) / bands;
    let slice = clipHalfPlane(pts, n, n, b + 0.012);
    slice = clipHalfPlane(slice, -n, -n, -a);
    if (slice.length < 3) continue;
    addPoly(slide, slice, { fill: { color: mixHex(from, to, (i + 0.5) / bands) } });
  }
}

/** Every card in the deck sits on the same soft purple glow, drawn as its own
 *  caster shape underneath so that gradient-filled cards (built from many
 *  slices) still get a single clean halo. The original effect is scaled to 88%
 *  of the shape, which makes the halo escape only past the top and left edges. */
function glowBehind(slide, x, y, w, h, r, opacity) {
  if (!opacity) return;
  slide.addShape('roundRect', {
    x, y, w: w * 0.88, h: h * 0.88, rectRadius: r * 0.88,
    fill: { color: PURPLE }, line: NO_LINE,
    shadow: { type: 'outer', blur: 36, offset: 0, angle: 45, color: PURPLE, opacity },
  });
}

/** Rounded rectangle with the 45 deg gradient (plus optional drop glow). */
function gradCard(slide, x, y, w, h, r, shadowOpacity) {
  glowBehind(slide, x, y, w, h, r, shadowOpacity);
  addGradPoly(slide, roundRectPoints(x, y, w, h, r));
}

/** White card with the thin gradient outline used throughout the deck. */
function whiteCard(slide, x, y, w, h, r, shadowOpacity) {
  glowBehind(slide, x, y, w, h, r, shadowOpacity);
  slide.addShape('roundRect', {
    x, y, w, h, rectRadius: r, fill: { color: WHITE }, line: { color: PURPLE, width: 0.75 },
  });
}

/* ------------------------------------------------------------------ *
 * Image placeholders (all photos in the source are stock grey mock-ups)
 * ------------------------------------------------------------------ */
function photo(slide, x, y, w, h, r = 0.25, label = true) {
  if (r > 0) {
    slide.addShape('roundRect', { x, y, w, h, rectRadius: r, fill: { color: IMG_FILL }, line: NO_LINE });
  } else {
    slide.addShape('rect', { x, y, w, h, fill: { color: IMG_FILL }, line: NO_LINE });
  }
  if (label && w > 1 && h > 0.6) {
    slide.addText('[image]', {
      x, y: y + h / 2 - 0.25, w, h: 0.5, align: 'center', valign: 'middle',
      fontFace: F_BODY, fontSize: 14, color: IMG_LABEL,
    });
  }
}

/** Photo block whose top corners only are rounded. */
function photoTopRounded(slide, x, y, w, h, r) {
  addPoly(slide, roundRectPoints(x, y, w, h, [r, r, 0, 0]), { fill: { color: IMG_FILL } });
  slide.addText('[image]', {
    x, y: y + h / 2 - 0.25, w, h: 0.5, align: 'center', valign: 'middle',
    fontFace: F_BODY, fontSize: 14, color: IMG_LABEL,
  });
}

/* ------------------------------------------------------------------ *
 * Text helpers - the deck only uses a handful of type styles
 * ------------------------------------------------------------------ */
const T = (slide, text, opts) => slide.addText(text, Object.assign({ valign: 'top' }, opts));

const title = (s, x, y, w, h, text, extra) =>
  T(s, text, Object.assign({ x, y, w, h, fontFace: F_BOLD, fontSize: 32, color: INK }, extra));

const heading = (s, x, y, w, h, text, extra) =>   // 18pt default size
  T(s, text, Object.assign({ x, y, w, h, fontFace: F_BOLD, fontSize: 18, color: INK }, extra));

const bodyCopy = (s, x, y, w, h, text, extra) =>
  T(s, text, Object.assign({
    x, y, w, h, fontFace: F_BODY, fontSize: 11, color: MUTED, lineSpacingMultiple: 1.5,
  }, extra));

const bigNumber = (s, x, y, w, h, text, extra) =>
  T(s, text, Object.assign({ x, y, w, h, fontFace: F_BOLD, fontSize: 44, color: PURPLE }, extra));

/* ------------------------------------------------------------------ *
 * Chrome shared by every slide: logo mark + top navigation
 * ------------------------------------------------------------------ */
function navBar(slide, color) {
  // logo mark: a dot plus two rounded slashes leaning up to the right
  slide.addShape('ellipse', { x: 0.570, y: 0.537, w: 0.060, h: 0.060, fill: { color }, line: NO_LINE });
  slide.addShape('roundRect', { x: 0.551, y: 0.593, w: 0.207, h: 0.052, rectRadius: 0.026, rotate: 306, fill: { color }, line: NO_LINE });
  slide.addShape('roundRect', { x: 0.636, y: 0.642, w: 0.149, h: 0.052, rectRadius: 0.026, rotate: 307, fill: { color }, line: NO_LINE });

  T(slide, 'Patiesche', { x: 0.841, y: 0.49, w: 1.201, h: 0.269, fontFace: F_BOLD, fontSize: 10, color });
  T(slide, 'Dashboard', { x: 9.253, y: 0.49, w: 1.078, h: 0.269, fontFace: F_NAV, fontSize: 10, color, align: 'center' });
  T(slide, 'Details', { x: 10.522, y: 0.49, w: 0.931, h: 0.269, fontFace: F_BODY, fontSize: 10, color, align: 'center' });
  T(slide, 'Products', { x: 11.629, y: 0.49, w: 0.931, h: 0.269, fontFace: F_BODY, fontSize: 10, color, align: 'center' });
}

/** Full-bleed 45 deg purple -> peach slide background (slides 1 and 15). */
function gradientBackground(slide) {
  addGradPoly(slide, [[0, 0], [SLIDE_W, 0], [SLIDE_W, SLIDE_H], [0, SLIDE_H]], { bands: 48 });
}

/* ------------------------------------------------------------------ *
 * Small vector icons that stand in for the deck's icon graphics
 * ------------------------------------------------------------------ */
function badgeIcon(slide, x, y, w, h) {           // "certified" thumbs-up rosette
  const P = { color: PURPLE };
  const cx = x + w / 2;
  const cy = y + w / 2;
  const R = w / 2;
  // scalloped medal edge: radius wobbles around the circle
  const scallop = [];
  for (let i = 0; i < 160; i++) {
    const a = (i / 160) * 2 * Math.PI;
    const r = R * (0.94 + 0.06 * Math.cos(16 * a));
    scallop.push([cx + r * Math.cos(a), cy + r * Math.sin(a)]);
  }
  slide.addShape('triangle', { x: x + w * 0.1, y: y + h * 0.66, w: w * 0.26, h: h * 0.3, rotate: 205, fill: P, line: NO_LINE });
  slide.addShape('triangle', { x: x + w * 0.64, y: y + h * 0.66, w: w * 0.26, h: h * 0.3, rotate: 155, fill: P, line: NO_LINE });
  addPoly(slide, scallop, { fill: P });
  slide.addShape('ellipse', { x: cx - R * 0.62, y: cy - R * 0.62, w: R * 1.24, h: R * 1.24, fill: { color: WHITE }, line: NO_LINE });
  // thumbs-up: fist + raised thumb
  slide.addShape('roundRect', { x: cx - R * 0.42, y: cy - R * 0.08, w: R * 0.28, h: R * 0.5, rectRadius: 0.008, fill: P, line: NO_LINE });
  slide.addShape('roundRect', { x: cx - R * 0.12, y: cy - R * 0.16, w: R * 0.58, h: R * 0.58, rectRadius: 0.012, fill: P, line: NO_LINE });
  slide.addShape('roundRect', { x: cx - R * 0.06, y: cy - R * 0.56, w: R * 0.18, h: R * 0.44, rectRadius: 0.012, fill: P, line: NO_LINE });
}

function trophyIcon(slide, x, y, w, h) {
  const P = { color: PURPLE };
  const handle = { color: PURPLE, width: 1.6 };
  slide.addShape('ellipse', { x: x + w * 0.02, y: y + h * 0.1, w: w * 0.3, h: h * 0.32, fill: NO_LINE, line: handle });
  slide.addShape('ellipse', { x: x + w * 0.68, y: y + h * 0.1, w: w * 0.3, h: h * 0.32, fill: NO_LINE, line: handle });
  slide.addShape('rect', { x: x + w * 0.22, y: y + h * 0.02, w: w * 0.56, h: h * 0.08, fill: P, line: NO_LINE });   // rim
  // bowl: straight sides that taper to a rounded base
  addPoly(slide, [
    [x + w * 0.24, y + h * 0.08], [x + w * 0.76, y + h * 0.08],
    [x + w * 0.68, y + h * 0.44], [x + w * 0.58, y + h * 0.56],
    [x + w * 0.42, y + h * 0.56], [x + w * 0.32, y + h * 0.44],
  ], { fill: P });
  slide.addShape('rect', { x: x + w * 0.44, y: y + h * 0.56, w: w * 0.12, h: h * 0.16, fill: P, line: NO_LINE });
  slide.addShape('roundRect', { x: x + w * 0.28, y: y + h * 0.68, w: w * 0.44, h: h * 0.14, rectRadius: 0.03, fill: P, line: NO_LINE });
  slide.addShape('rect', { x: x + w * 0.16, y: y + h * 0.86, w: w * 0.68, h: h * 0.14, fill: P, line: NO_LINE });
}

function therapistIcon(slide, x, y, w, h) {
  const P = { color: PURPLE };
  const tube = { color: WHITE, width: 1.1 };
  slide.addShape('ellipse', { x: x + w * 0.26, y: y + h * 0.05, w: w * 0.48, h: h * 0.44, fill: P, line: NO_LINE });   // hair
  slide.addShape('ellipse', { x: x + w * 0.33, y: y + h * 0.16, w: w * 0.34, h: h * 0.4, fill: { color: WHITE }, line: { color: PURPLE, width: 1.1 } }); // face
  slide.addShape('chord', { x: x + w * 0.26, y: y + h * 0.05, w: w * 0.48, h: h * 0.26, angleRange: [180, 360], fill: P, line: NO_LINE });              // fringe
  addPoly(slide, roundRectPoints(x + w * 0.06, y + h * 0.6, w * 0.88, h * 0.4, [0.055, 0.055, 0.008, 0.008]), { fill: P }); // shoulders
  addPoly(slide, [[x + w * 0.38, y + h * 0.58], [x + w * 0.62, y + h * 0.58], [x + w * 0.5, y + h * 0.74]],
    { fill: { color: WHITE } });                                                                                          // collar
  slide.addShape('arc', { x: x + w * 0.24, y: y + h * 0.58, w: w * 0.26, h: h * 0.3, angleRange: [0, 180], line: tube });
  slide.addShape('line', { x: x + w * 0.58, y: y + h * 0.58, w: w * 0.08, h: h * 0.2, line: tube });
  slide.addShape('ellipse', { x: x + w * 0.59, y: y + h * 0.77, w: w * 0.14, h: h * 0.14, fill: NO_LINE, line: tube });
}

function calendarIcon(slide, x, y, w, h) {
  const W = { color: WHITE };
  slide.addShape('roundRect', { x: x + w * 0.24, y, w: w * 0.09, h: h * 0.26, rectRadius: 0.018, fill: W, line: NO_LINE });
  slide.addShape('roundRect', { x: x + w * 0.67, y, w: w * 0.09, h: h * 0.26, rectRadius: 0.018, fill: W, line: NO_LINE });
  slide.addShape('rect', { x, y: y + h * 0.17, w, h: h * 0.83, fill: NO_LINE, line: { color: WHITE, width: 1.25 } });
  slide.addShape('rect', { x, y: y + h * 0.17, w, h: h * 0.2, fill: W, line: NO_LINE });
  slide.addShape('line', { x: x + w * 0.26, y: y + h * 0.63, w: w * 0.14, h: h * 0.14, line: { color: WHITE, width: 1.4 } });
  slide.addShape('line', { x: x + w * 0.4, y: y + h * 0.52, w: w * 0.28, h: h * 0.25, flipV: true, line: { color: WHITE, width: 1.4 } });
}

/** small downward "accordion" chevron */
function chevron(slide, x, y, color) {
  slide.addShape('triangle', { x, y, w: 0.143, h: 0.068, rotate: 180, fill: { color }, line: NO_LINE });
}

/* ================================================================== *
 * Slide builders
 * ================================================================== */

/* 1 - Cover ---------------------------------------------------------- */
function slide01(pres) {
  const s = pres.addSlide();
  gradientBackground(s);
  photo(s, 7.071, 0.778, 5.365, 6.722, 0);
  navBar(s, WHITE);

  T(s, 'A Safe Space For Healing & Growth.', {
    x: 1.355, y: 1.792, w: 6.181, h: 2.794, fontFace: F_BOLD, fontSize: 60, color: WHITE, lineSpacing: 64,
  });
  bodyCopy(s, 1.45, 4.535, 4.35, 0.625,
    'provide a holistic approach to emotional well-being, offering a range of therapies tailored to individual needs.',
    { color: WHITE });

  // "Start Presentation" pill
  whiteCard(s, 1.45, 5.551, 2.7, 0.559, 0.28, 0.6);
  T(s, 'Start Presentation', { x: 1.683, y: 5.65, w: 2.7, h: 0.37, fontFace: F_BOLD, fontSize: 16, color: INK });

  // avatar cluster
  whiteCard(s, 6.667, 5.405, 2.029, 0.894, 0.227, 0.6);
  [6.829, 7.214, 7.599].forEach((x) =>
    s.addShape('ellipse', { x, y: 5.577, w: 0.549, h: 0.549, fill: { color: IMG_FILL }, line: NO_LINE }));
  addGradPoly(s, roundRectPoints(7.97, 5.577, 0.549, 0.549, 0.2745), { bands: 10 });
  T(s, '2k', { x: 8.008, y: 5.657, w: 0.548, h: 0.37, fontFace: F_BOLD, fontSize: 16, color: WHITE });
}

/* 2 - Intro / stats -------------------------------------------------- */
function slide02(pres) {
  const s = pres.addSlide();
  photoTopRounded(s, 6.454, 1.489, 5.176, 6.011, 0.342);
  navBar(s, INK);

  title(s, 0.865, 1.701, 4.548, 1.714, 'Discover Clarity, Confidence, And Emotional Wellness');
  badgeIcon(s, 1.016, 4.826, 0.319, 0.401);
  T(s, 'Certified Therapist', { x: 1.533, y: 4.831, w: 3.48, h: 0.404, fontFace: F_SEMI, fontSize: 18, color: INK });
  bodyCopy(s, 0.929, 5.347, 3.952, 0.903, LOREM_FULL);

  gradCard(s, 10.407, 2.267, 2.122, 1.833, 0.258, 0.6);
  T(s, '50K+', { x: 10.4, y: 2.692, w: 2.135, h: 0.841, fontFace: F_BOLD, fontSize: 44, color: WHITE, align: 'center' });
  T(s, 'Supported', { x: 10.616, y: 3.426, w: 1.704, h: 0.404, fontFace: F_SEMI, fontSize: 18, color: WHITE, align: 'center' });

  whiteCard(s, 10.417, 4.501, 2.122, 1.833, 0.258, 0.6);
  T(s, '85%', { x: 10.395, y: 4.927, w: 2.165, h: 0.841, fontFace: F_BOLD, fontSize: 44, color: PURPLE, align: 'center' });
  T(s, 'Outcomes', { x: 10.455, y: 5.661, w: 2.046, h: 0.404, fontFace: F_SEMI, fontSize: 18, color: INK, align: 'center' });
}

/* 3 - Three numbered benefits ---------------------------------------- */
function slide03(pres) {
  const s = pres.addSlide();
  photo(s, 0.692, 1.413, 7.57, 2.857, 0.3);
  gradCard(s, 0.857, 3.7, 2.984, 3.31, 0.301, 0.6);
  navBar(s, INK);

  title(s, 8.956, 1.701, 4.091, 1.178, 'Helps You Manage Stress');

  const items = [
    // x,   number, heading,                 numberColor, headColor, bodyColor, headY, bodyY, numY
    [1.186, '01', 'Heal From Within', WHITE, WHITE, WHITE, 5.225, 5.701, 4.49, 2.671, LOREM_SHORT],
    [4.942, '02', 'Grow Stronger', PURPLE, INK, MUTED, 5.427, 5.903, 4.692, 3.016, LOREM_MED],
    [8.972, '03', 'Mental Health Journey', PURPLE, INK, MUTED, 5.427, 5.903, 4.692, 3.016, LOREM_MED],
  ];
  items.forEach(([x, n, head, nc, hc, bc, headY, bodyY, numY, bw, copy]) => {
    bigNumber(s, x, numY, 0.992, 0.841, n, { color: nc });
    heading(s, x - 0.015, headY, 3.48, 0.404, head, { color: hc });
    bodyCopy(s, x - 0.015, bodyY, bw, 0.903, copy, { color: bc });
  });
}

/* 4 - Tailored to individual needs ----------------------------------- */
function slide04(pres) {
  const s = pres.addSlide();
  photo(s, 5.58, 1.413, 7.039, 2.857, 0.3);
  gradCard(s, 5.58, 4.703, 7.039, 2.091, 0.288, 0.6);
  navBar(s, INK);

  title(s, 0.841, 1.701, 4.175, 1.178, 'Tailored To Individual Needs');
  heading(s, 0.835, 5.19, 3.48, 0.404, 'Heal From Within');
  bodyCopy(s, 0.815, 5.665, 3.952, 0.903, LOREM_FULL);

  whiteCard(s, 5.793, 4.931, 1.064, 1.613, 0.258, 0.6);
  trophyIcon(s, 6.128, 5.565, 0.394, 0.367);
  s.addShape('line', { x: 7.252, y: 4.931, w: 0, h: 1.613, line: { color: WHITE, width: 1 } });

  heading(s, 7.677, 5.11, 3.48, 0.404, 'Heal From Within', { color: WHITE });
  bodyCopy(s, 7.677, 5.518, 3.952, 0.903, LOREM_FULL, { color: WHITE });
}

/* 5 - Three feature cards -------------------------------------------- */
function slide05(pres) {
  const s = pres.addSlide();
  photoTopRounded(s, 0.002, 5.29, 13.33, 2.21, 0.58);
  addPoly(s, roundRectPoints(0.002, 5.312, 13.33, 2.188, [0.53, 0.53, 0, 0]),
    { fill: { color: PURPLE, transparency: 75 } });
  navBar(s, INK);

  title(s, 2.115, 1.522, 9.103, 0.64, 'Find Balance, Embrace Life', { align: 'center' });

  const cards = [
    { x: 1.729, w: 3.016, grad: true, imgX: 1.983, headX: 1.983, headW: 2.508, head: 'Personalized Wellness Plans', bodyX: 1.935, bodyW: 2.605, lineX: 2.81 },
    { x: 5.151, w: 3.031, grad: false, imgX: 5.413, headX: 5.167, headW: 2.853, head: '24/7 Support\nCommunity', bodyX: 5.383, bodyW: 2.485, lineX: 6.22 },
    { x: 8.588, w: 3.031, grad: false, imgX: 8.82, headX: 8.604, headW: 2.853, head: 'Interactive\nSelf-Care Tools', bodyX: 8.82, bodyW: 2.485, lineX: 9.698 },
  ];
  cards.forEach((c) => {
    if (c.grad) gradCard(s, c.x, 2.675, c.w, 4.336, 0.301, 0.6);
    else whiteCard(s, c.x, 2.675, c.w, 4.336, 0.291, 0.6);
    photo(s, c.imgX, 2.971, 2.508, 1.105, 0.184, false);
    heading(s, c.headX, 4.39, c.headW, 0.707, c.head, { align: 'center', color: c.grad ? WHITE : INK });
    s.addShape('line', { x: c.lineX, y: 5.29, w: 0.812, h: 0, line: { color: c.grad ? WHITE : PURPLE, width: 1 } });
    bodyCopy(s, c.bodyX, 5.484, c.bodyW, 0.903, LOREM_SHORT, { align: 'center', color: c.grad ? WHITE : MUTED });
  });
}

/* 6 - Holistic therapy accordion ------------------------------------- */
function slide06(pres) {
  const s = pres.addSlide();
  photo(s, 0.656, 3.231, 12.016, 3.546, 0.35);
  gradCard(s, 8.372, 1.71, 4.299, 2.032, 0.262, 0.6);
  whiteCard(s, 7.762, 4.248, 4.299, 2.032, 0.262, 0.6);
  navBar(s, INK);

  title(s, 0.865, 1.582, 6.321, 1.178, 'Support Your Mental Health With Holistic Therapy');

  heading(s, 8.659, 2.036, 3.48, 0.404, 'Individual Therapy', { color: WHITE });
  bodyCopy(s, 8.629, 2.517, 3.952, 0.903, LOREM_FULL, { color: WHITE });
  chevron(s, 12.1, 2.204, WHITE);

  heading(s, 8.071, 4.578, 3.48, 0.404, 'Cognitive Behavioral');
  bodyCopy(s, 8.042, 5.03, 3.952, 0.903, LOREM_FULL);
  chevron(s, 11.453, 4.782, PURPLE);
}

/* 7 - Therapy process ------------------------------------------------ */
function slide07(pres) {
  const s = pres.addSlide();
  photo(s, 0.662, 1.229, 12.01, 3.696, 0.273);
  navBar(s, INK);

  whiteCard(s, 8.693, 3.309, 3.659, 2.044, 0.188, 0.23);
  gradCard(s, 8.693, 2.098, 3.659, 0.802, 0.176, 0.23);
  whiteCard(s, 8.693, 5.761, 3.659, 0.817, 0.197, 0.23);

  heading(s, 8.895, 2.282, 3.109, 0.404, '01. Appointment', { color: WHITE });
  chevron(s, 11.887, 2.46, WHITE);

  heading(s, 8.895, 3.619, 2.705, 0.404, '02. Consultation');
  s.addShape('triangle', { x: 11.887, y: 3.795, w: 0.143, h: 0.068, fill: { color: PURPLE }, line: NO_LINE });
  bodyCopy(s, 8.895, 4.111, 3.169, 0.903,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Curabitur cursus sapien eget egestas vestibulum. Nam fringilla');

  heading(s, 8.895, 5.971, 2.705, 0.404, '03. Session & Review');
  chevron(s, 11.887, 6.147, PURPLE);

  title(s, 0.865, 5.424, 4.548, 1.178, 'Mental Health Therapy Process');
  bodyCopy(s, 5.194, 5.561, 3.06, 0.903, LOREM_MED);
}

/* 8 - Meet the therapists -------------------------------------------- */
function slide08(pres) {
  const s = pres.addSlide();
  navBar(s, INK);
  title(s, 2.758, 1.511, 7.818, 0.64, 'Meet Our Senior Therapist', { align: 'center' });

  // card frames come from the slide layout
  const people = [
    { cx: 0.637, cw: 3.016, r: 0.305, grad: true, imgX: 0.842, nameX: 0.891, nameW: 2.508, roleX: 0.842, roleW: 2.605, name: 'Mia Thompson' },
    { cx: 4.059, cw: 3.031, r: 0.292, grad: false, imgX: 4.272, nameX: 4.075, nameW: 2.853, roleX: 4.291, roleW: 2.485, name: 'Noah Martinez' },
    { cx: 7.545, cw: 3.031, r: 0.292, grad: false, imgX: 7.758, nameX: 7.561, nameW: 2.853, roleX: 7.777, roleW: 2.485, name: 'Noah Martinez' },
  ];
  people.forEach((p) => {
    if (p.grad) gradCard(s, p.cx, 2.557, p.cw, 4.336, p.r, 0.6);
    else whiteCard(s, p.cx, 2.557, p.cw, 4.336, p.r, 0.6);
    photo(s, p.imgX, 2.798, 2.605, 2.986, 0.204, false);
    T(s, p.name, { x: p.nameX, y: 5.923, w: p.nameW, h: 0.404, fontFace: F_BOLD, fontSize: 18, align: 'center', color: p.grad ? WHITE : INK });
    T(s, 'Psychologist', { x: p.roleX, y: 6.23, w: p.roleW, h: 0.37, fontFace: F_BODY, fontSize: 12, align: 'center', lineSpacingMultiple: 1.5, color: p.grad ? WHITE : MUTED });
  });

  therapistIcon(s, 11.053, 3.64, 0.594, 0.598);
  heading(s, 10.915, 4.452, 2.513, 0.707, 'Experienced\nTherapist');
  bodyCopy(s, 10.903, 5.327, 1.794, 0.903, 'Lorem ipsum dolor \nsit amet, consectetur adipiscing elit. ');
}

/* 9 - Personalised care ---------------------------------------------- */
function slide09(pres) {
  const s = pres.addSlide();
  photo(s, 0.656, 1.307, 5.847, 3.657, 0.348);
  photo(s, 6.955, 4.407, 3.724, 2.24, 0.27);
  photo(s, 11.151, 4.407, 1.503, 2.24, 0.253);
  gradCard(s, 0.68, 5.391, 5.803, 1.144, 0.244, 0.23);
  navBar(s, INK);

  title(s, 7.535, 1.558, 5.143, 1.178, 'Provide Personalized Care to Help You ');
  bodyCopy(s, 7.535, 2.932, 4.324, 0.903, LOREM_FULL);
  heading(s, 0.911, 5.61, 5.589, 0.707, 'Overcoming Anxiety Through Cognitive Behavioral Therapy (CBT)', { color: WHITE });
}

/* 10 - Testimonials -------------------------------------------------- */
function slide10(pres) {
  const s = pres.addSlide();
  photoTopRounded(s, 0.693, 1.229, 5.974, 6.271, 0.4);
  navBar(s, INK);
  title(s, 7.499, 1.685, 4.548, 1.178, 'We Create A Safe & Supportive Space');

  const quotes = [
    { cx: 3.46, tx: 3.734, grad: true, num: '01', name: 'Alexander Davis', stars: 5 },
    { cx: 6.51, tx: 6.784, grad: false, num: '02', name: 'James Adams', stars: 4 },
    { cx: 9.553, tx: 9.854, grad: false, num: '03', name: 'John Doe', stars: 4 },
  ];
  quotes.forEach((q) => {
    if (q.grad) gradCard(s, q.cx, 3.489, 2.606, 2.988, 0.245, 0.37);
    else whiteCard(s, q.cx, 3.486, 2.619, 2.988, 0.263, 0.37);
    T(s, q.num, { x: q.tx, y: 3.83, w: 1.012, h: 0.774, fontFace: F_SEMI, fontSize: 40, color: q.grad ? WHITE : PURPLE });
    T(s, q.name, { x: q.tx, y: 4.561, w: 1.865, h: 0.415, fontFace: F_SEMI, fontSize: 14, lineSpacingMultiple: 1.5, color: q.grad ? WHITE : INK });
    for (let i = 0; i < 5; i++) {
      const on = i < q.stars;
      const color = q.grad ? WHITE : (on ? PURPLE : SILVER);
      s.addShape('star5', { x: q.tx + 0.1 + i * 0.1613, y: 4.986, w: 0.132, h: 0.132, fill: { color }, line: NO_LINE });
    }
    bodyCopy(s, q.tx, 5.201, 2.145, 0.903, LOREM_NEC, { color: q.grad ? WHITE : MUTED });
  });
}

/* 11 - Photo collage ------------------------------------------------- */
function slide11(pres) {
  const s = pres.addSlide();
  photo(s, 6.388, 1.229, 3.48, 4.007, 0.312);
  photo(s, 10.331, 2.229, 1.676, 1.706, 0.284);
  photo(s, 10.331, 4.35, 2.396, 2.485, 0.215);
  photo(s, 7.6, 5.705, 2.268, 1.123, 0.224);
  navBar(s, INK);

  title(s, 0.865, 1.701, 4.548, 1.714, 'Transforming Lives Through Mental Wellness.');
  heading(s, 0.929, 4.831, 3.48, 0.404, 'Based on Trusted Fact');
  bodyCopy(s, 0.929, 5.347, 3.952, 0.903, LOREM_FULL);
}

/* 12 - Appointment form ---------------------------------------------- */
function slide12(pres) {
  const s = pres.addSlide();
  // phone mock-up (from the slide layout): black bezel inside a brushed metal
  // rail, grey screen and a notch hanging from the top edge
  s.addShape('roundRect', {
    x: 3.774, y: 1.222, w: 3.871, h: 6.278, rectRadius: 0.52,
    fill: { color: '111111' }, line: { color: 'C9CFCB', width: 2 },
  });
  photoTopRounded(s, 3.918, 1.351, 3.608, 6.149, 0.4);
  addPoly(s, roundRectPoints(4.967, 1.222, 1.471, 0.44, [0, 0, 0.1, 0.1]), { fill: { color: '111111' } });
  navBar(s, INK);

  title(s, 8.702, 1.701, 3.64, 1.717, 'Book Your Appointment Through App');
  bodyCopy(s, 8.766, 3.634, 3.577, 0.903,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Curabitur cursus sapien eget egestas vestibulum. Nam fringilla neque');
  bigNumber(s, 8.798, 5.005, 3.577, 0.841, '4 Steps');
  heading(s, 8.782, 5.74, 3.48, 0.404, 'Register');

  // form card
  whiteCard(s, 0.841, 2.953, 3.697, 3.799, 0.188, 0.23);
  gradCard(s, 0.841, 1.742, 3.659, 0.802, 0.176, 0.23);
  heading(s, 1.043, 1.926, 3.109, 0.404, 'Set the Appointment', { color: WHITE });
  chevron(s, 4.035, 2.104, WHITE);

  const field = (label, y, filled) => {
    if (filled) {
      gradCard(s, 1.203, y, 3.016, 0.613, 0.3065, 0.09);
      calendarIcon(s, 3.671, y + 0.187, 0.214, 0.239);
    } else {
      whiteCard(s, 1.203, y, 3.016, 0.613, 0.3065, 0.09);
    }
    T(s, label, { x: 1.505, y: y + 0.086, w: 1.876, h: 0.37, fontFace: F_BODY, fontSize: 12, italic: true, lineSpacingMultiple: 1.5, color: filled ? WHITE : MUTED });
    if (!filled) s.addShape('line', { x: 1.505, y: y + 0.456, w: 2.38, h: 0, line: { color: PURPLE, width: 0.75 } });
  };

  T(s, 'Name', { x: 1.203, y: 3.251, w: 1.255, h: 0.37, fontFace: F_BODY, fontSize: 12, italic: true, lineSpacingMultiple: 1.5, color: MUTED });
  s.addShape('line', { x: 1.144, y: 3.707, w: 3.035, h: 0, line: { color: PURPLE, width: 0.75 } });
  field('Date mm/dd/yy', 4.012, true);
  field('Confirm Your Email', 4.876, false);
  field('Subject', 5.74, false);
}

/* 13 - Pricing ------------------------------------------------------- */
function slide13(pres) {
  const s = pres.addSlide();
  navBar(s, INK);

  title(s, 0.865, 1.701, 4.548, 1.178, 'Tailored Therapy Plans For You');
  bigNumber(s, 0.865, 4.408, 3.577, 0.841, '1 Month');
  heading(s, 0.841, 5.16, 3.48, 0.404, 'Free Trial');
  bodyCopy(s, 0.841, 5.564, 3.952, 0.903, LOREM_FULL);

  const plans = [
    {
      grad: true, cx: 5.747, cw: 3.085, r: 0.245, tag: 'Recommended', price: '$100',
      tagX: 6.022, priceX: 6.419, copyX: 6.022, f1X: 5.589, f2X: 5.891,
      features: ['Personalized assessment', 'action plan for self-care', 'Weekly health check-ins'],
      btnX: 6.219, btnY: 6.33, btnTxtX: 6.355, btnTxtY: 6.417,
    },
    {
      grad: false, cx: 9.368, cw: 3.1, r: 0.263, tag: 'Regular', price: '$450',
      tagX: 9.665, priceX: 10.061, copyX: 9.665, f1X: 9.231, f2X: 9.533,
      features: ['Personalized assessment', 'action plan for self-care'],
      btnX: 9.785, btnY: 5.777, btnTxtX: 9.987, btnTxtY: 5.849,
    },
  ];

  plans.forEach((p) => {
    if (p.grad) gradCard(s, p.cx, 1.389, p.cw, 5.343, p.r, 0.37);
    else whiteCard(s, p.cx, 1.385, p.cw, 5.343, p.r, 0.37);
    const fg = p.grad ? WHITE : INK;

    T(s, p.tag, { x: p.tagX, y: 1.906, w: 2.503, h: 0.348, fontFace: F_BODY, fontSize: 11, align: 'center', lineSpacingMultiple: 1.5, color: fg });
    T(s, p.price, { x: p.priceX, y: 2.26, w: 1.711, h: 0.841, fontFace: F_SEMI, fontSize: 44, align: 'center', color: fg });
    T(s, 'Per Session', { x: p.tagX, y: 2.968, w: 2.503, h: 0.348, fontFace: F_BODY, fontSize: 11, align: 'center', lineSpacingMultiple: 1.5, color: fg });
    bodyCopy(s, p.copyX, 3.419, 2.503, 0.903, LOREM_NEC + ' sem', { align: 'center', color: fg });

    p.features.forEach((f, i) => {
      const x = i === 0 ? p.f1X : p.f2X;
      const w = i === 0 ? 3.37 : 2.766;
      T(s, f, { x, y: 4.644 + i * 0.4835, w, h: 0.337, fontFace: F_SEMI, fontSize: 14, align: 'center', color: fg });
    });

    if (p.grad) {
      whiteCard(s, p.btnX, p.btnY, 2.234, 0.577, 0.2885, 0.37);
      T(s, 'Get Started', { x: p.btnTxtX, y: p.btnTxtY, w: 1.83, h: 0.404, fontFace: F_BOLD, fontSize: 18, align: 'center', color: INK_SOFT });
    } else {
      gradCard(s, p.btnX, p.btnY, 2.234, 0.577, 0.2885, 0.37);
      T(s, 'Get Started', { x: p.btnTxtX, y: p.btnTxtY, w: 1.83, h: 0.404, fontFace: F_BOLD, fontSize: 18, align: 'center', color: WHITE });
    }
  });
}

/* 14 - Contact ------------------------------------------------------- */
function slide14(pres) {
  const s = pres.addSlide();
  photo(s, 0.637, 1.229, 12.035, 4.076, 0.294);
  navBar(s, INK);

  title(s, 0.841, 5.604, 5.56, 1.178, 'Empowering Minds, Enhancing Lives.');

  T(s, 'Address', { x: 6.667, y: 5.69, w: 1.755, h: 0.337, fontFace: F_SEMI, fontSize: 14, color: INK });
  bodyCopy(s, 6.667, 6.049, 2.694, 0.625, '123 Street Name, Anywhere 400, Any City, State 12345');

  T(s, 'Phones', { x: 10.114, y: 5.69, w: 1.755, h: 0.337, fontFace: F_SEMI, fontSize: 14, color: INK });
  T(s, '+1 555-123-4567', { x: 10.114, y: 6.072, w: 1.656, h: 0.303, fontFace: F_BODY, fontSize: 12, color: MUTED });
  T(s, '+1 555-234-5678', { x: 10.107, y: 6.371, w: 1.656, h: 0.303, fontFace: F_BODY, fontSize: 12, color: MUTED });
}

/* 15 - Closing ------------------------------------------------------- */
function slide15(pres) {
  const s = pres.addSlide();
  gradientBackground(s);

  T(s, 'See You', { x: 0.841, y: 1.751, w: 6.492, h: 1.717, fontFace: F_BOLD, fontSize: 96, color: WHITE });
  photo(s, 2.042, 1.079, 7.902, 6.421, 0);

  // peach corner blob (free-form, rotated)
  s.addShape('custGeom', {
    x: 11.195, y: -0.48, w: 2.374, h: 2.697, rotate: 340,
    points: [
      { x: 0, y: 0 },
      { x: 2.374, y: 0.851 },
      { x: 1.712, y: 2.697 },
      { x: 0.686, y: 2.697 },
      { curve: { type: 'cubic', x1: 0.307, y1: 2.697, x2: 0.0, y2: 2.389 }, x: 0, y: 2.011 },
      { close: true },
    ],
    fill: { color: PEACH }, line: NO_LINE,
  });

  navBar(s, WHITE);

  whiteCard(s, 0.841, 3.513, 2.388, 0.559, 0.2795, 0.6);
  T(s, 'End Presentation', { x: 1.075, y: 3.613, w: 2.7, h: 0.37, fontFace: F_BOLD, fontSize: 16, color: INK });

  T(s, 'Next Time', { x: 5.937, y: 4.073, w: 8.048, h: 1.717, fontFace: F_BOLD, fontSize: 96, color: WHITE });
}

/* ================================================================== */
function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'DECK', width: SLIDE_W, height: SLIDE_H });
  pres.layout = 'DECK';
  pres.author = 'Patiesche';
  pres.title = 'A Safe Space For Healing & Growth';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
    slide09, slide10, slide11, slide12, slide13, slide14, slide15]
    .forEach((fn) => fn(pres));

  return pres.writeFile({
    fileName: path.join(__dirname, '12b4836d-d104-4bb8-ac70-3d90a94be0b1_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
