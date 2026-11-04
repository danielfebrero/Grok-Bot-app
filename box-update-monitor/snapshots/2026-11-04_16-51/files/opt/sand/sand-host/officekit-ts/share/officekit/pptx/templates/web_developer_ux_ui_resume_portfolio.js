#!/usr/bin/env node
/**
 * "Nickolas Hoard" CV / portfolio deck - 30 slides, 13.333 x 7.5 in (16:9).
 * Rebuilt with pptxgenjs only. Raster photos in the original are drawn here as
 * labelled grey placeholder rectangles.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const BLUE = '4044ED';
const INK = '262626';
const GREY = '808080';
const WHITE = 'FFFFFF';
const TRACK = 'F2F2F2';
const PLACEHOLDER = 'D9D9D9';

const DISPLAY = 'Montserrat';
const BODY = 'Open Sans';
const TAGLINE = 'Web Developer & UX/UI Designer';

const NO_LINE = { type: 'none' };

/** Soft drop shadow used by every white "floating card" in the deck.
 *  A fresh object per call: pptxgenjs rewrites these values in place. */
function soft() {
  return { type: 'outer', blur: 20, offset: 0.0001, angle: 0.0001, color: 'BFBFBF', opacity: 0.4 };
}

/* ------------------------------------------------- reusable body copy blocks */

const T = {
  lead: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et',
  para: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ',
  cillum: 'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ',
  culpa: 'Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt',
  card: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, ',
  service: 'Lorem ipsum dolor sit consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.',
  tile: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor',
  project: 'Lorem ipsum dolor sit consectetur adipiscing elit sed do eiusmod tempor incididunt labore et dolore magna aliqua.',
  projectCard: 'PLACEHOLDER',
  node: 'Lorem ipsum dolor sit consectetur adipiscing elit sed eiusmod',
  skill: 'Lorem ipsum dolor sit consectetur adipiscing elit, sed do',
  work: 'Lorem ipsum dolor sit amet consectet adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore',
  service3: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore',
  contact: 'Lorem ipsum dolor sit amet, consectetur adipis elit, sed do eiusmod tempor incididunt ',
};

/* ------------------------------------------------------- geometry primitives */

const K = 0.5523; // circle -> cubic bezier magic number

/** One closed ellipse sub-path; dir > 0 = clockwise, dir < 0 = counter-clockwise. */
function ellipseSubpath(cx, cy, rx, ry, dir) {
  const ox = K * rx;
  const oy = K * ry;
  const arcs = dir > 0
    ? [[cx + ox, cy - ry, cx + rx, cy - oy, cx + rx, cy],
       [cx + rx, cy + oy, cx + ox, cy + ry, cx, cy + ry],
       [cx - ox, cy + ry, cx - rx, cy + oy, cx - rx, cy],
       [cx - rx, cy - oy, cx - ox, cy - ry, cx, cy - ry]]
    : [[cx - ox, cy - ry, cx - rx, cy - oy, cx - rx, cy],
       [cx - rx, cy + oy, cx - ox, cy + ry, cx, cy + ry],
       [cx + ox, cy + ry, cx + rx, cy + oy, cx + rx, cy],
       [cx + rx, cy - oy, cx + ox, cy - ry, cx, cy - ry]];
  const pts = [{ x: cx, y: cy - ry, moveTo: true }];
  arcs.forEach(function (a) {
    pts.push({ x: a[4], y: a[5], curve: { type: 'cubic', x1: a[0], y1: a[1], x2: a[2], y2: a[3] } });
  });
  pts.push({ close: true });
  return pts;
}

/** Tapered ring: full outer circle with a slightly smaller circle punched out
 *  off-centre, so the stroke fades from thick (left) to hairline (right). */
function ringPoints(w, h) {
  return ellipseSubpath(0.5 * w, 0.5 * h, 0.5 * w, 0.5 * h, 1)
    .concat(ellipseSubpath(0.535 * w, 0.5 * h, 0.466 * w, 0.466 * h, -1));
}

/** Rectangle with individually rounded corners. `r` = {tl, tr, br, bl} inches. */
function roundedPoints(w, h, r) {
  const tl = r.tl || 0;
  const tr = r.tr || 0;
  const br = r.br || 0;
  const bl = r.bl || 0;
  const pts = [{ x: tl, y: 0, moveTo: true }, { x: w - tr, y: 0 }];
  if (tr) pts.push({ x: w, y: tr, curve: { type: 'cubic', x1: w - tr + K * tr, y1: 0, x2: w, y2: tr - K * tr } });
  pts.push({ x: w, y: h - br });
  if (br) pts.push({ x: w - br, y: h, curve: { type: 'cubic', x1: w, y1: h - br + K * br, x2: w - br + K * br, y2: h } });
  pts.push({ x: bl, y: h });
  if (bl) pts.push({ x: 0, y: h - bl, curve: { type: 'cubic', x1: bl - K * bl, y1: h, x2: 0, y2: h - bl + K * bl } });
  pts.push({ x: 0, y: tl });
  if (tl) pts.push({ x: tl, y: 0, curve: { type: 'cubic', x1: 0, y1: tl - K * tl, x2: tl - K * tl, y2: 0 } });
  pts.push({ close: true });
  return pts;
}

/** Half-circle top that narrows into a flat bottom edge (the "What I Do?" tiles). */
function domePoints(w, h) {
  const c = [[0.776, 0.000, 1.000, 0.288, 1.000, 0.644],
             [1.000, 0.755, 0.978, 0.860, 0.940, 0.951]];
  const c2 = [[0.022, 0.860, 0.000, 0.755, 0.000, 0.644],
              [0.000, 0.288, 0.224, 0.000, 0.500, 0.000]];
  const cub = function (a) {
    return { x: a[4] * w, y: a[5] * h, curve: { type: 'cubic', x1: a[0] * w, y1: a[1] * h, x2: a[2] * w, y2: a[3] * h } };
  };
  return [{ x: 0.5 * w, y: 0, moveTo: true }, cub(c[0]), cub(c[1]),
          { x: 0.916 * w, y: h }, { x: 0.084 * w, y: h }, { x: 0.060 * w, y: 0.951 * h },
          cub(c2[0]), cub(c2[1]), { close: true }];
}

/* ------------------------------------------------------------ shape helpers */

/** Rounded / plain panel. `r` may be omitted for a square rectangle. */
function panel(s, x, y, w, h, fill, r, shadow) {
  const opt = { x: x, y: y, w: w, h: h, fill: { color: fill }, line: NO_LINE };
  if (shadow) opt.shadow = soft();
  if (r) {
    opt.points = roundedPoints(w, h, r);
    s.addShape('custGeom', opt);
  } else {
    s.addShape('rect', opt);
  }
}

/** White floating card with the deck's signature soft shadow. */
function card(s, x, y, w, h, r) {
  panel(s, x, y, w, h, WHITE, r, true);
}

function circle(s, x, y, d, fill, shadow) {
  const opt = { x: x, y: y, w: d, h: d, fill: { color: fill }, line: NO_LINE };
  if (shadow) opt.shadow = soft();
  s.addShape('ellipse', opt);
}

/** Numbered step bubble ("01", "02", ...) - 0.645in circle + centred label. */
function badge(s, x, y, label, fill, ink) {
  circle(s, x, y, 0.645, fill, true);
  s.addText(label, {
    x: x + 0.0725, y: y + 0.145, w: 0.5, h: 0.37, align: 'center', valign: 'top',
    fontFace: DISPLAY, fontSize: 16, bold: true, color: ink, wrap: false, fit: 'resize',
  });
}

/** Stand-in for a photograph from the original deck. `radius` mimics the
 *  rounded bezel of the phone / monitor mock-ups it replaces. */
function imagePlaceholder(s, x, y, w, h, radius) {
  const opt = { x: x, y: y, w: w, h: h, fill: { color: PLACEHOLDER }, line: NO_LINE };
  if (radius) {
    s.addShape('roundRect', Object.assign(opt, { rectRadius: radius }));
  } else {
    s.addShape('rect', opt);
  }
  s.addText('[image]', {
    x: x, y: y + h / 2 - 0.18, w: w, h: 0.36, align: 'center', valign: 'middle',
    fontFace: BODY, fontSize: 11, color: GREY,
  });
}

/* ------------------------------------------------------------- text helpers */

/** 40pt display heading + the recurring 12pt tagline underneath it.
 *  `wrap` mirrors the original text box: non-wrapping boxes grow around their
 *  anchor (so the words sit centred in `w`), wrapping ones start flush at `x`. */
function sectionTitle(s, x, y, w, str, o) {
  o = o || {};
  const wrap = o.wrap === true;
  s.addText(str, {
    x: x, y: y, w: w, h: 0.774, fontFace: DISPLAY, fontSize: 40, bold: true,
    color: o.color || INK, align: o.align || 'left', valign: 'top', wrap: wrap, fit: 'resize',
  });
  s.addText(o.tagline || TAGLINE, {
    x: o.subX === undefined ? x : o.subX, y: y + 0.774, w: o.subW || 3.009, h: 0.303,
    fontFace: DISPLAY, fontSize: 12, color: o.subColor || (o.color === WHITE ? WHITE : BLUE),
    align: o.align || 'left', valign: 'top', wrap: wrap, fit: 'resize',
  });
}

/** 11pt Open Sans paragraph, 1.5 line spacing (the deck's only body style). */
function body(s, x, y, w, h, str, o) {
  o = o || {};
  s.addText(str, {
    x: x, y: y, w: w, h: h, fontFace: BODY, fontSize: 11, color: o.color || GREY,
    align: o.align || 'justify', lineSpacingMultiple: 1.5, valign: 'top', fit: 'resize',
  });
}

/** 16pt bold item title, sized from TITLE_W so it lands where the original does. */
function itemTitle(s, x, y, str, color) {
  label(s, x, y, TITLE_W[str] || 2.036, 0.37, str, 16, color);
}

/** Bold Montserrat run - used for item titles, stat numbers and small labels. */
function label(s, x, y, w, h, str, size, color, o) {
  o = o || {};
  s.addText(str, {
    x: x, y: y, w: w, h: h, fontFace: DISPLAY, fontSize: size, color: color,
    bold: o.bold !== false, align: o.align || 'left', valign: 'top',
    wrap: o.wrap === true, fit: 'resize', charSpacing: o.charSpacing,
  });
}

/** "250+ / Project Done" pair. `vw`/`cw` are the value/caption box widths -
 *  they hug the glyphs in the original, which is what keeps them left-aligned. */
function stat(s, x, y, vw, cw, value, caption, o) {
  o = o || {};
  label(s, x, y, vw, 0.505, value, 24, o.color || INK, { align: o.align });
  label(s, x, y + 0.505, cw, 0.303, caption, 12, o.captionColor || BLUE, { align: o.align, bold: false });
}

/** Box widths (inches) of the recurring 16pt item titles. The originals are
 *  non-wrapping boxes shrink-wrapped to the text, so the width matters for
 *  placement - too wide and the renderer re-centres the words. */
const TITLE_W = {
  'Web Developer': 2.036,
  'UI/UX Design': 1.767,
  'Senior Designer': 2.083,
  'Design Art': 1.459,
  'Photoshop': 1.489,
  'Web Design': 1.65,
};

/** The "Name / Birth / Address" fact list, reused on slides 4 and 18. */
const DETAILS = [
  { label: 'Name     :', lw: 1.273, value: 'Nickolas Hoard', vw: 1.906 },
  { label: 'Birth       :', lw: 1.292, value: 'September 18, 1990', vw: 2.370 },
  { label: 'Address :', lw: 1.285, value: 'New York, USA', vw: 1.867 },
];

/** Renders the fact list at (x, y) with `step` inches between rows. */
function detailList(s, labelX, valueX, y, step) {
  DETAILS.forEach(function (r, i) {
    label(s, labelX, y + i * step, r.lw, 0.37, r.label, 16, INK);
    label(s, valueX, y + i * step, r.vw, 0.37, r.value, 16, GREY, { bold: false });
  });
}

/** The three head-line figures, always in the same order and box widths. */
const STATS = {
  projects: { value: '250+', caption: 'Project Done', vw: 1.029, cw: 1.308 },
  clients: { value: '200+', caption: 'Happy Client', vw: 1.057, cw: 1.312 },
  years: { value: '15+', caption: 'Year Of Existence', vw: 0.735, cw: 1.689 },
};

function statBlock(s, x, y, key, o) {
  const d = STATS[key];
  stat(s, x, y, d.vw, d.cw, d.value, d.caption, o);
}

/** Same pair, centred on `cx` instead of left-aligned at `x`. */
function statBlockCentered(s, cx, y, key, o) {
  const d = STATS[key];
  o = Object.assign({}, o, { align: 'center' });
  label(s, cx - d.vw / 2, y, d.vw, 0.505, d.value, 24, o.color || INK, { align: 'center' });
  label(s, cx - d.cw / 2, y + 0.505, d.cw, 0.303, d.caption, 12, o.captionColor || BLUE, { align: 'center', bold: false });
}

/** Progress bar with a name on the left and a percentage on the right. */
function skillBar(s, x, y, w, filled, name, pct, o) {
  o = o || {};
  const barY = y + (o.gap || 0.447);
  const h = o.h || 0.103;
  const size = o.size || 14;
  const lh = o.labelH || 0.337;
  panel(s, x, barY, w, h, TRACK);
  panel(s, x, barY, filled, h, BLUE);
  label(s, x - 0.119, y, 1.516, lh, name, size, INK, { wrap: true });
  label(s, x + w - 0.714, y, 0.714, lh, pct, size, INK, { align: 'right', wrap: true });
}

/* ------------------------------------------------------------------- slides */

// 1 - cover: tapered blue ring, white name plate.
function slide01(s) {
  s.addShape('custGeom', { x: 1.393, y: 1.376, w: 4.736, h: 4.747, points: ringPoints(4.736, 4.747), fill: { color: BLUE }, line: NO_LINE });
  body(s, 7.898, 5.871, 4.689, 0.627, T.lead);
  card(s, 5.146, 2.63, 8.188, 2.241, { tl: 0.373, bl: 0.373 });
  label(s, 5.794, 3.043, 6.89, 1.111, 'Nickolas Hoard', 60, INK);
  label(s, 5.794, 4.154, 4.271, 0.303, TAGLINE, 12, BLUE, { bold: false, charSpacing: 3 });
}

// 2 - About Me with a blue right column and a stats strip.
function slide02(s) {
  panel(s, 9.018, 0, 4.315, 7.5, BLUE);
  panel(s, 5.788, 4.928, 7.546, 1.857, WHITE, null, true);
  circle(s, 9.789, 0.715, 2.773, WHITE);
  label(s, 9.959, 3.75, 2.434, 0.438, 'Nickolas Hoard', 20, WHITE, { align: 'center' });
  label(s, 9.671, 4.188, 3.009, 0.303, TAGLINE, 12, WHITE, { bold: false });
  statBlock(s, 11.251, 5.452, 'clients');
  statBlock(s, 9.095, 5.452, 'projects');
  statBlock(s, 6.559, 5.452, 'years');
  sectionTitle(s, 1.176, 0.9, 3.066, 'About Me');
  body(s, 1.174, 2.179, 6.667, 1.183, T.para + 'Duis aute irure dolor in occaecat cupidatat non proident, sunt in culpa qui officia');
  body(s, 1.174, 3.4, 6.667, 0.627, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt occaecat cupidatat non proident, sunt in culpa qui officia');
  body(s, 1.174, 5.404, 3.588, 0.905, 'Lorem ipsum dolor sit consectetur adipiscing, sed do eiusmod tempor incididunt occaecat cupidatat non proident, sunt in culpa');
}

// 3 - big name over a solid blue stats band.
function slide03(s) {
  // The original box is narrower than the word "Nickolas" at 66pt, so the
  // break is forced here rather than left to the renderer.
  s.addText([{ text: 'Nickolas', options: { breakLine: true } }, { text: 'Hoard' }], {
    x: 7.041, y: 1.037, w: 4.3, h: 2.124, fontFace: DISPLAY, fontSize: 66, bold: true,
    color: INK, valign: 'top', lineSpacing: 83, margin: 0,
  });
  label(s, 6.942, 3.372, 3.896, 0.303, TAGLINE + ', New York', 12, BLUE, { bold: false });
  panel(s, 4.593, 4.722, 7.389, 1.539, BLUE);
  const onBlue = { color: WHITE, captionColor: WHITE };
  statBlock(s, 9.977, 5.088, 'clients', onBlue);
  statBlock(s, 5.285, 5.088, 'projects', onBlue);
  statBlock(s, 7.441, 5.088, 'years', onBlue);
}

// 4 - About Me with a rounded detail card over a blue corner block.
function slide04(s) {
  panel(s, 5.526, 3.75, 7.807, 3.75, BLUE);
  card(s, 2.117, 4.133, 6.8, 2.983, { tl: 0.497, tr: 0.497, br: 0.497, bl: 0.497 });
  detailList(s, 4.146, 5.576, 4.763, 0.679);
  sectionTitle(s, 1.176, 1.336, 3.066, 'About Me');
  body(s, 5.583, 1.145, 6.667, 1.461, T.para + T.cillum + T.culpa);
  const onBlue = { color: WHITE, captionColor: WHITE };
  statBlockCentered(s, 11.125, 4.68, 'projects', onBlue);
  statBlockCentered(s, 11.125, 5.763, 'clients', onBlue);
}

// 5 - three stacked paragraphs, blue caption block bottom-right.
function slide05(s) {
  sectionTitle(s, 1.176, 1.734, 3.066, 'About Me');
  body(s, 1.176, 2.981, 6.667, 1.183, T.para + T.cillum);
  body(s, 1.176, 4.199, 6.667, 0.905, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ');
  body(s, 1.176, 5.139, 6.667, 0.627, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ');
  panel(s, 9.07, 5.351, 4.263, 2.149, BLUE);
  label(s, 9.697, 5.773, 3.009, 0.37, 'Nickolas Hoard', 16, WHITE, { align: 'center' });
  label(s, 9.697, 6.145, 3.009, 0.303, TAGLINE, 12, WHITE, { bold: false, align: 'center' });
  body(s, 9.579, 6.45, 3.246, 0.627, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed', { color: WHITE, align: 'center' });
}

// 6 - Education: two columns over a wide blue band with arrow bullets.
function slide06(s) {
  sectionTitle(s, 1.176, 0.85, 3.196, 'Education');
  body(s, 1.176, 2.11, 5.224, 1.183, T.para);
  body(s, 7.093, 2.11, 5.224, 1.183, T.para);
  panel(s, 0, 4.143, 12.233, 2.43, BLUE, { br: 0.405 });
  [1.157, 6.686].forEach(function (x, i) {
    s.addShape('triangle', { x: x, y: 4.719, w: 0.279, h: 0.24, rotate: 90, fill: { color: WHITE }, line: NO_LINE });
    label(s, x + 0.488, 4.671, 1.468, 0.337, 'Web Design', 14, WHITE);
    body(s, i === 0 ? 1.176 : 6.705, 5.134, 4.44, 0.905, T.card, { color: WHITE });
  });
  [0.563, 6.091].forEach(function (x) {
    s.addShape('bentConnector2', { x: x, y: 4.144, w: 0.613, h: 0.695, rotate: 180, fill: NO_LINE, line: { color: WHITE, width: 0.75 } });
  });
}

// 7 - Education: white band left, two blue cards stacked right.
function slide07(s) {
  panel(s, 0, 1.776, 13.333, 3.998, WHITE, null, true);
  [[1.249, { tl: 0.303 }], [3.828, { bl: 0.303 }]].forEach(function (row) {
    const y = row[0];
    panel(s, 6.874, y, 6.459, 2.423, BLUE, row[1]);
    s.addShape('triangle', { x: 7.884, y: y + 0.575, w: 0.279, h: 0.24, rotate: 90, fill: { color: WHITE }, line: NO_LINE });
    label(s, 8.372, y + 0.527, 1.468, 0.337, 'Web Design', 14, WHITE);
    body(s, 7.903, y + 0.991, 4.44, 0.905, T.card, { color: WHITE });
    s.addShape('bentConnector2', { x: 7.29, y: y, w: 0.613, h: 0.695, rotate: 180, fill: NO_LINE, line: { color: WHITE, width: 1 } });
  });
  sectionTitle(s, 1.176, 2.529, 3.196, 'Education');
  body(s, 1.176, 3.788, 4.561, 1.183, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo');
}

// 8 - Education timeline: blue title block left, three numbered steps right.
function slide08(s) {
  panel(s, 0, 2.54, 5.387, 2.422, BLUE, { tr: 0.404, br: 0.404 });
  const steps = [['01', 1.588], ['02', 3.429], ['03', 5.267]];
  steps.forEach(function (st) {
    badge(s, 6.667, st[1], st[0], WHITE, BLUE);
    itemTitle(s, 7.903, st[1] - 0.315, 'Web Design', BLUE);
    body(s, 7.903, st[1] + 0.055, 4.44, 0.905, T.card);
  });
  s.addShape('bentConnector3', { x: 6.667, y: 1.911, w: 0.014, h: 1.84, rotate: 180, flipV: true, fill: NO_LINE, line: { color: BLUE, width: 1 } });
  s.addShape('bentConnector3', { x: 7.311, y: 3.751, w: 0.014, h: 1.838, fill: NO_LINE, line: { color: BLUE, width: 1 } });
  sectionTitle(s, 1.064, 3.213, 3.196, 'Education', { color: WHITE });
}

// 9 - Experience timeline: blue text block left, three roles right.
function slide09(s) {
  panel(s, 0, 2.704, 5.494, 3.658, BLUE, { tr: 0.404, br: 0.404 });
  sectionTitle(s, 1.159, 0.961, 3.491, 'Experience');
  const roles = [['01', 2.704, 'Web Developer'], ['02', 4.211, 'UI/UX Design'], ['03', 5.718, 'Senior Designer']];
  roles.forEach(function (r) {
    badge(s, 6.507, r[1], r[0], WHITE, BLUE);
    itemTitle(s, 7.637, r[1] - 0.176, r[2], BLUE);
    body(s, 7.637, r[1] + 0.194, 4.44, 0.627, T.lead);
  });
  [3.349, 4.856].forEach(function (y) {
    s.addShape('line', { x: 6.83, y: y, w: 0, h: 0.862, line: { color: BLUE, width: 1 } });
  });
  body(s, 1.159, 3.526, 3.287, 2.016, T.para + 'Duis deserunt mollit anim id', { color: WHITE });
}

// 10 - Experience: wide white card with two numbered entries.
function slide10(s) {
  card(s, 0.89, 4.416, 11.553, 2.415, { tl: 0.497, tr: 0.497, br: 0.497, bl: 0.497 });
  [[1.927, '01', 3.057, 'Web Developer', '2005-2009'],
   [7.123, '02', 8.252, 'UI/UX Design', '2009-2012']].forEach(function (e) {
    badge(s, e[0], 5.297, e[1], BLUE, WHITE);
    itemTitle(s, e[2], 4.97, e[3], INK);
    label(s, e[2], 5.34, 1.13, 0.303, e[4], 12, BLUE);
    body(s, e[2], 5.643, 3.21, 0.627, T.tile);
  });
  sectionTitle(s, 1.263, 0.961, 3.491, 'Experience');
  body(s, 1.263, 2.177, 6.667, 1.183, T.para + T.cillum);
  body(s, 8.793, 2.177, 3.304, 1.183, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna cillum dolore eu fugiat nulla pariatur. ');
}

// 11 - Experience: two white cards left, blue title band right.
function slide11(s) {
  panel(s, 0, 2.54, 7.925, 2.422, BLUE, { tr: 0.404, br: 0.404 });
  panel(s, 6.917, 2.54, 5.387, 2.422, BLUE, { tr: 0.404, br: 0.404 });
  const round = { tl: 0.497, tr: 0.497, br: 0.497, bl: 0.497 };
  [[1.209, 3.876, '02', 4.761, 4.433, 'UI/UX Design', '2009-2012'],
   [1.209, 1.209, '01', 2.09, 1.762, 'Web Developer', '2005-2009']].forEach(function (c) {
    card(s, c[0], c[1], 5.708, 2.415, round);
    badge(s, 1.893, c[3], c[2], BLUE, WHITE);
    itemTitle(s, 3.023, c[4], c[5], INK);
    label(s, 3.023, c[4] + 0.371, 1.13, 0.303, c[6], 12, BLUE);
    body(s, 3.023, c[4] + 0.674, 3.21, 0.627, T.tile);
  });
  sectionTitle(s, 7.865, 3.211, 3.491, 'Experience', { color: WHITE });
  body(s, 9.19, 5.871, 3.276, 0.627, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod');
}

// 12 - Language: blue title block plus three progress bars.
function slide12(s) {
  panel(s, 0, 0.664, 5.387, 2.422, BLUE, { tr: 0.404, br: 0.404 });
  sectionTitle(s, 1.122, 1.336, 3.144, 'Language', { color: WHITE });
  body(s, 6.62, 1.145, 5.612, 1.461, T.para + T.cillum);
  skillBar(s, 1.45, 4.138, 4.597, 4.297, 'English', '95%');
  skillBar(s, 1.45, 5.018, 4.597, 3.413, 'Spanish', '80%');
  skillBar(s, 1.45, 5.898, 4.597, 2.932, 'Russian', '70%');
}

// 13 - Language: bars on the right, blue "15+" tag bottom-left.
function slide13(s) {
  sectionTitle(s, 1.263, 0.961, 3.144, 'Language');
  body(s, 1.263, 2.177, 10.834, 0.905, T.para + T.cillum);
  skillBar(s, 7.5, 4.138, 4.597, 4.297, 'English', '95%');
  skillBar(s, 7.5, 5.018, 4.597, 3.413, 'Spanish', '80%');
  skillBar(s, 7.5, 5.898, 4.597, 2.932, 'Russian', '70%');
  panel(s, -0.008, 5.001, 2.542, 1.539, BLUE);
  statBlockCentered(s, 1.263, 5.367, 'years', { color: WHITE, captionColor: WHITE });
}

// 14 - What I Do?: three circular tiles with pill captions.
function slide14(s) {
  const tiles = [[1.003, 1.324, 'UI/UX Design', 1.401], [5.058, 5.379, 'Web Developer', 5.457], [9.113, 9.434, 'Design Art', 9.512]];
  tiles.forEach(function (t) { circle(s, t[0], 3.569, 3.197, WHITE, true); });
  tiles.forEach(function (t) {
    s.addShape('roundRect', { x: t[1], y: 3.769, w: 2.554, h: 0.762, fill: { color: BLUE }, line: NO_LINE });
    label(s, t[1], 3.971, 2.554, 0.37, t[2], 16, WHITE, { align: 'center' });
  });
  sectionTitle(s, 1.263, 1.246, 3.494, 'What I Do?');
  body(s, 5.583, 1.054, 6.667, 1.461, T.para + T.cillum + T.culpa);
  tiles.forEach(function (t) { body(s, t[3], 4.839, 2.4, 0.905, T.tile, { align: 'center' }); });
  panel(s, 0, 6.052, 13.333, 1.448, BLUE);
}

// 15 - What I Do?: two dome tiles left, copy right.
function slide15(s) {
  s.addShape('custGeom', { x: 1.611, y: 1.137, w: 3.197, h: 2.483, points: domePoints(3.197, 2.483), fill: { color: WHITE }, line: NO_LINE, shadow: soft() });
  s.addShape('roundRect', { x: 1.933, y: 1.336, w: 2.554, h: 0.762, fill: { color: BLUE }, line: NO_LINE });
  label(s, 1.933, 1.538, 2.554, 0.37, 'UI/UX Design', 16, WHITE, { align: 'center' });
  body(s, 2.01, 2.406, 2.4, 0.905, T.tile, { align: 'center' });
  s.addShape('custGeom', { x: 1.611, y: 3.881, w: 3.197, h: 2.483, points: domePoints(3.197, 2.483), rotate: 180, fill: { color: WHITE }, line: NO_LINE, shadow: soft() });
  s.addShape('roundRect', { x: 1.933, y: 5.402, w: 2.554, h: 0.762, fill: { color: BLUE }, line: NO_LINE });
  label(s, 1.933, 5.592, 2.554, 0.37, 'Web Developer', 16, WHITE, { align: 'center' });
  body(s, 2.01, 4.189, 2.4, 0.905, T.tile, { align: 'center' });
  sectionTitle(s, 6.091, 1.635, 3.494, 'What I Do?');
  body(s, 6.091, 2.845, 6.077, 1.461, T.para + T.cillum + 'Excepteur sint occaecat cupidatat non');
  body(s, 6.091, 4.305, 6.077, 0.905, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis fugiat nulla pariatur. Excepteur sint occaecat cupidatat non');
  body(s, 6.091, 5.21, 6.077, 0.627, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor nulla pariatur. Excepteur sint occaecat cupidatat non');
}

// 16 - Skills & Knowledge: three tabbed cards down the left edge.
function slide16(s) {
  panel(s, 0, -0.024, 2.054, 7.524, BLUE);
  sectionTitle(s, 5.324, 1.635, 5.814, 'Skills & Knowledge');
  const alt = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco voluptate velit esse cillum dolore eu fugiat nulla pariatur. ';
  body(s, 5.324, 2.845, 6.845, 1.183, T.para + T.cillum);
  body(s, 5.324, 4.028, 6.845, 0.905, alt);
  body(s, 5.324, 4.933, 6.845, 0.905, alt);
  [[0.554, 'UI/UX Design'], [2.865, 'Web Design'], [5.177, 'Photoshop']].forEach(function (c) {
    card(s, 0, c[0], 4.125, 1.746, { tr: 0.359, br: 0.359 });
    itemTitle(s, 0.661, c[0] + 0.374, c[1], INK);
    body(s, 0.661, c[0] + 0.744, 2.803, 0.627, T.skill);
  });
}

// 17 - Skills & Knowledge: compact bar card bottom-right.
function slide17(s) {
  card(s, 7.542, 4.419, 5.051, 2.411, { tr: 0.44, br: 0.44 });
  skillBar(s, 8.577, 5.017, 2.984, 2.671, 'UI/UX Design', '90%', { gap: 0.27, h: 0.194, size: 10, labelH: 0.269 });
  skillBar(s, 8.577, 5.77, 2.984, 1.97, 'Photoshop', '80%', { gap: 0.269, h: 0.194, size: 10, labelH: 0.269 });
  sectionTitle(s, 1.176, 0.813, 5.814, 'Skills & Knowledge');
  body(s, 1.176, 2.023, 10.981, 0.905, T.para + T.cillum);
}

// 18 - The Services I Provide: three pill cards plus a detail list.
function slide18(s) {
  panel(s, 0.568, 2.743, 12.238, 2.024, WHITE, null, true);
  [[1.324, 'UI/UX Design', 1.263], [5.163, 'Web Developer', 5.102], [9.002, 'Design Art', 8.941]].forEach(function (c) {
    s.addShape('roundRect', { x: c[0], y: 2.362, w: 2.554, h: 0.762, fill: { color: BLUE }, line: NO_LINE });
    label(s, c[0], 2.563, 2.554, 0.37, c[1], 16, WHITE, { align: 'center' });
    body(s, c[2], 3.492, 3.171, 0.905, T.service3);
  });
  sectionTitle(s, 1.263, 0.813, 6.658, 'The Services I Provide');
  detailList(s, 8.577, 10.007, 5.455, 0.492);
}

// 19 - The Services I Provide: 2x2 numbered grid on one big white card.
function slide19(s) {
  panel(s, 0.265, 0.292, 12.803, 6.916, WHITE, null, true);
  sectionTitle(s, 1.263, 1.141, 6.658, 'The Services I Provide');
  const items = [[1.336, 3.173, '01', 2.572, 2.706, 'UI/UX Design'],
                 [7.223, 3.173, '03', 8.457, 2.706, 'Web Developer'],
                 [1.336, 5.25, '02', 2.571, 4.775, 'Design Art'],
                 [7.223, 5.25, '04', 8.457, 4.775, 'Photoshop']];
  items.forEach(function (it) {
    badge(s, it[0], it[1], it[2], BLUE, WHITE);
    itemTitle(s, it[3], it[4], it[5], INK);
    label(s, it[3], it[4] + 0.371, 1.103, 0.303, 'My Service', 12, BLUE, { bold: false });
    body(s, it[3], it[4] + 0.674, 3.663, 0.905, T.service);
  });
}

// 20 - The Services I Provide: blue banner with two numbered items.
function slide20(s) {
  body(s, 1.176, 5.294, 5.059, 1.183, T.para);
  body(s, 7.061, 5.294, 5.059, 1.183, T.para);
  sectionTitle(s, 1.176, 4.084, 6.658, 'The Services I Provide');
  panel(s, 0, 0.473, 13.333, 2.58, BLUE);
  [[1.336, '01', 2.572, 'UI/UX Design'], [7.223, '02', 8.457, 'Web Developer']].forEach(function (it) {
    badge(s, it[0], 1.44, it[1], WHITE, BLUE);
    itemTitle(s, it[2], 0.974, it[3], WHITE);
    label(s, it[2], 1.344, 1.103, 0.303, 'My Service', 12, WHITE, { bold: false });
    body(s, it[2], 1.647, 3.663, 0.905, T.service, { color: WHITE });
  });
}

// 21 - Project Samples: centred intro plus three captioned thumbnails.
function slide21(s) {
  sectionTitle(s, 4.164, 0.738, 4.983, 'Project Samples', { align: 'center', subX: 5.151 });
  body(s, 1.012, 2.015, 11.31, 0.627, T.para, { align: 'center' });
  panel(s, 4.322, 3.846, 0.747, 0.411, BLUE);
  panel(s, 8.242, 3.846, 0.747, 0.411, BLUE);
  [[2.134, 1.012, 'Project 1'], [6.031, 4.932, 'Project 2'], [9.951, 8.852, 'Project 3']].forEach(function (p) {
    label(s, p[0], 5.46, 1.249, 0.37, p[2], 16, INK, { align: 'center' });
    body(s, p[1], 5.83, 3.447, 0.905, T.project, { align: 'center' });
  });
}

// 22 - Project Samples: two rounded cards wired to the title with elbows.
function slide22(s) {
  [[1.603, 'Project 1', 2.189], [3.911, 'Project 2', 2.166]].forEach(function (c) {
    card(s, 1.037, c[0], 3.506, 1.906, { tl: 0.221, tr: 0.221, br: 0.221, bl: 0.221 });
    label(s, c[2], c[0] + 0.316, 1.249, 0.37, c[1], 16, BLUE, { align: 'center' });
    body(s, 1.42, c[0] + 0.686, 2.741, 0.905, T.projectCard, { align: 'center' });
  });
  s.addShape('bentConnector3', { x: 4.544, y: 4.71, w: 1.037, h: 1.15, fill: NO_LINE, line: { color: BLUE, width: 1 } });
  s.addShape('bentConnector2', { x: 4.544, y: 3.288, w: 6.147, h: 0.933, fill: NO_LINE, line: { color: BLUE, width: 1 } });
  sectionTitle(s, 6.369, 1.105, 4.983, 'Project Samples');
}

// 23 - Project Samples: three left-anchored cards with blue tabs.
function slide23(s) {
  [0.535, 2.757, 4.979].forEach(function (y) { card(s, 0, y, 3.506, 1.906, { tr: 0.221, br: 0.221 }); });
  [1.283, 3.504, 5.726].forEach(function (y) { panel(s, 3.506, y, 0.863, 0.411, BLUE); });
  [[0.851, 'Project 1', 1.152], [3.072, 'Project 2', 1.129], [5.294, 'Project 3', 1.129]].forEach(function (c) {
    label(s, c[2], c[0], 1.249, 0.37, c[1], 16, BLUE, { align: 'center' });
    body(s, 0.383, c[0] + 0.37, 2.741, 0.905, T.projectCard, { align: 'center' });
  });
  sectionTitle(s, 7.261, 3.172, 4.983, 'Project Samples');
}

// 24 - My Vision: blue title block linked to three stacked note cards.
function slide24(s) {
  panel(s, 0, 2.499, 5.387, 2.422, BLUE, { tr: 0.404, br: 0.404 });
  sectionTitle(s, 1.083, 3.172, 3.042, 'My Vision', { color: WHITE });
  s.addShape('line', { x: 5.387, y: 3.71, w: 2.995, h: 0, line: { color: BLUE, width: 1 } });
  [1.087, 2.896, 4.705].forEach(function (y) {
    card(s, 8.382, y, 3.506, 1.629, { tl: 0.221, tr: 0.221, br: 0.221, bl: 0.221 });
    label(s, 9.154, y + 0.349, 1.962, 0.337, 'Simple Text Here', 14, BLUE, { align: 'center' });
    body(s, 8.765, y + 0.685, 2.741, 0.627, T.node, { align: 'center' });
  });
  s.addShape('bentConnector3', { x: 8.382, y: 1.886, w: 0.014, h: 3.618, fill: NO_LINE, line: { color: BLUE, width: 1 } });
}

// 25 - My Mission: blue header branching down to three note cards.
function slide25(s) {
  [2.738, 6.515, 10.292].forEach(function (x) { panel(s, x, 4.705, 0.303, 0.303, WHITE); });
  s.addShape('custGeom', { x: 4.397, y: 1.086, w: 4.539, h: 1.981, points: roundedPoints(4.539, 1.981, { tl: 0.27, tr: 0.27, br: 0.27, bl: 0.27 }), fill: { color: BLUE }, line: NO_LINE });
  sectionTitle(s, 4.938, 1.538, 3.457, 'My Mission', { color: WHITE, align: 'center', subX: 5.162 });
  [1.137, 4.913, 8.69].forEach(function (x) {
    card(s, x, 4.705, 3.506, 1.629, { tl: 0.221, tr: 0.221, br: 0.221, bl: 0.221 });
    label(s, x + 0.772, 5.054, 1.962, 0.337, 'Simple Text Here', 14, BLUE, { align: 'center' });
    body(s, x + 0.383, 5.391, 2.741, 0.627, T.node, { align: 'center' });
  });
  s.addShape('bentConnector4', { x: 2.32, y: 2.628, w: 2.646, h: 1.507, rotate: 90, flipH: true, flipV: true, fill: NO_LINE, line: { color: BLUE, width: 1 } });
  s.addShape('line', { x: 6.667, y: 3.067, w: 0.007, h: 1.638, flipV: true, line: { color: BLUE, width: 1 } });
  s.addShape('bentConnector2', { x: 8.384, y: 2.646, w: 2.611, h: 1.507, rotate: 270, flipV: true, fill: NO_LINE, line: { color: BLUE, width: 1 } });
}

// 26 - Work Samples: phone mock-ups on a blue column.
function slide26(s) {
  panel(s, 10.392, 0, 2.941, 7.5, BLUE);
  sectionTitle(s, 1.225, 1.175, 4.677, 'Work Samples', { subW: 3.132, wrap: true });
  body(s, 1.225, 2.39, 6.828, 0.905, T.para);
  card(s, 0, 4.47, 7.627, 1.746, { tr: 0.359, br: 0.359 });
  [0.691, 4.122].forEach(function (x) {
    label(s, x, 4.844, 1.589, 0.37, 'Simple Text', 16, BLUE);
    body(s, x, 5.214, 2.803, 0.627, T.skill);
  });
  imagePlaceholder(s, 9.17, 1.284, 2.445, 4.932, 0.3);
  imagePlaceholder(s, 11.899, 1.284, 2.445, 4.932, 0.3);
}

// 27 - Work Samples: three pill captions above a monitor mock-up.
function slide27(s) {
  imagePlaceholder(s, 0.961, 3.793, 3.664, 3.081);
  panel(s, 0, 0.737, 13.333, 2.024, WHITE, null, true);
  [[1.534, 1.027], [5.588, 5.081], [9.643, 9.136]].forEach(function (c) {
    body(s, c[1], 1.106, 3.171, 0.905, T.work, { align: 'center' });
    s.addShape('roundRect', { x: c[0], y: 2.38, w: 2.157, h: 0.762, fill: { color: BLUE }, line: NO_LINE });
    label(s, c[0], 2.57, 2.157, 0.37, 'Simple Text', 16, WHITE, { align: 'center' });
  });
  sectionTitle(s, 5.783, 4.264, 4.677, 'Work Samples', { subW: 3.132, wrap: true });
  body(s, 5.783, 5.479, 6.524, 0.905, T.para);
}

// 28 - Work Samples: blue title bar beside a laptop mock-up.
function slide28(s) {
  panel(s, 1.189, 4.08, 6.514, 2.422, BLUE, { tl: 0.404, bl: 0.404 });
  sectionTitle(s, 2.316, 4.752, 4.677, 'Work Samples', { color: WHITE, subW: 3.132, wrap: true });
  body(s, 1.393, 1.58, 6.524, 0.905, T.para);
  panel(s, 9.152, 0, 2.618, 2.769, WHITE, null, true);
  statBlock(s, 9.809, 0.462, 'projects');
  statBlock(s, 9.806, 1.492, 'clients');
  imagePlaceholder(s, 7.197, 3.75, 5.281, 3.081);
}

// 29 - My Contact: blue contact bar over two overlapping white panels.
function slide29(s) {
  panel(s, 3.918, 2.03, 8.263, 4.444, WHITE, null, true);
  sectionTitle(s, 5.103, 3.824, 4.677, 'My Contact', { subW: 3.132, wrap: true });
  body(s, 5.103, 5.039, 5.996, 0.627, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ');
  panel(s, 1.152, 1.026, 9.605, 1.99, BLUE);
  [[2.022, 'Phone Number', 2.027], [6.315, 'Email Address', 6.315]].forEach(function (c) {
    label(s, c[0], 1.54, 2.758, 0.337, c[1], 14, WHITE, { wrap: true });
    body(s, c[2], 1.876, 3.572, 0.627, T.contact, { color: WHITE });
  });
  panel(s, 0.861, 3.365, 2.618, 2.769, WHITE, null, true);
  statBlock(s, 1.518, 3.827, 'projects');
  statBlock(s, 1.515, 4.857, 'clients');
}

// 30 - closing card, mirrors the cover.
function slide30(s) {
  s.addShape('custGeom', { x: 1.393, y: 1.376, w: 4.736, h: 4.747, points: ringPoints(4.736, 4.747), fill: { color: BLUE }, line: NO_LINE });
  body(s, 7.898, 5.871, 4.689, 0.627, T.lead);
  card(s, 5.146, 2.63, 8.188, 2.241, { tl: 0.373, bl: 0.373 });
  label(s, 6.178, 2.993, 6.122, 1.212, 'Thank You', 66, INK, { align: 'center', charSpacing: 6 });
  label(s, 7.104, 4.204, 4.271, 0.303, TAGLINE, 12, BLUE, { bold: false, align: 'center', charSpacing: 3 });
}

/* --------------------------------------------------------------------- main */

const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
                slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
                slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24,
                slide25, slide26, slide27, slide28, slide29, slide30];

function build() {
  const pptx = new PptxGenJS();
  pptx.author = 'Nickolas Hoard';
  pptx.title = 'Nickolas Hoard - Web Developer & UX/UI Designer';
  pptx.defineLayout({ name: 'WIDE_16x9', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE_16x9';

  SLIDES.forEach(function (builder) {
    const slide = pptx.addSlide();
    slide.background = { color: WHITE };
    builder(slide);
  });

  return pptx.writeFile({ fileName: path.join(__dirname, '0e931695-9139-434c-ac2b-1ba90cac214c_grok_final.pptx') });
}

build().then(function (f) { console.log('wrote ' + f); }, function (e) { console.error(e); process.exit(1); });
