/*
 * CONTRUZ - construction presentation template  (30 slides, 13.333 x 7.5 in)
 * Rebuilt from scratch with pptxgenjs. Raster photos in the original deck are
 * replaced by flat "[image]" placeholder rectangles of the same geometry.
 *
 *   node 09b7ed50-e230-4f0a-ac22-0d8173b155c1_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const C = {
  accent: 'FFB202',   // brand amber (theme accent1)
  amber2: 'FB9900',   // theme accent2
  ink:    '1C1C1C',   // theme bg2 - headline black
  gray:   '555555',   // bg2 @ 75% luminance
  mid:    '8E8E8E',   // body copy grey
  line:   'C6C6C6',   // bg2 @ 25% luminance - hairlines
  pale:   'E8E8E8',
  faint:  'F2F2F2',
  white:  'FFFFFF',
};

const F = { head: 'Montserrat', body: 'Open Sans' };

/* Text presets: size / bold / italic / colour / face / align / line-height. */
const S = {
  title:   { size: 24, bold: true, color: C.ink,    face: F.head },
  logo:    { size: 24, bold: true, color: C.ink,    face: F.head, align: 'right' },
  mark:    { size: 20, bold: true, color: C.accent, face: F.head, align: 'right' },
  kicker:  { size: 9,  bold: true, color: C.mid,    face: F.head, lh: 1.5 },
  item:    { size: 11, bold: true, color: C.gray,   face: F.head, lh: 1.5 },
  eyebrow: { size: 10, ital: true, color: C.mid,                  lh: 1.5 },
  body:    { size: 10,             color: C.mid,                  lh: 1.5 },
  note:    { size: 8,              color: C.mid,                  lh: 1.5 },
  tag:     { size: 7,  ital: true, color: C.gray,                 lh: 1.5 },
  badge:   { size: 24, color: C.white, face: F.head, align: 'center', mid: true, nowrap: true },
};

/* Repeated placeholder copy (the template ships with lorem-style filler). */
const TITLE   = 'Architecture is the thoughtful making of space';
const TITLE_A = 'Architecture is the thoughtful making ';
const TITLE_B = 'of space';
const EYE     = 'aebesma hanabul banesi anisesan';
const EYE2    = 'aebesma hanabul anisesan';
const TAG     = 'contruction presentation';
const LOGO    = 'CONTRUZ';

const T = {
  p1: 'PLACEHOLDER',
  p2: 'PLACEHOLDER',
  p3: 'PLACEHOLDER',
  p4: 'PLACEHOLDER',
  p5: 'PLACEHOLDER',
  p6: 'PLACEHOLDER',
  p7: 'usateanhanab hana sanisin balanasei',
  p8: 'PLACEHOLDER',
  p9: 'usateanhanab hana sanisin balanasei anubane baniani andalisil',
  p10: 'PLACEHOLDER',
  p11: 'PLACEHOLDER',
  p12: 'PLACEHOLDER',
  p13: 'usateanhanab hana ban sanisinaia salana nusal ',
  p14: 'PLACEHOLDER',
  p15: 'bala nasil dani atasan ebetmit nunanes danatansa',
  p16: 'PLACEHOLDER',
  p17: 'PLACEHOLDER',
  p18: 'PLACEHOLDER',
  p19: 'PLACEHOLDER',
};

/* ----------------------------------------------------------- primitives */

// Solid rectangle; `alpha` is percent transparency (PowerPoint semantics).
function box(s, x, y, w, h, color, alpha) {
  s.addShape('rect', { x, y, w, h, fill: alpha ? { color, transparency: alpha } : { color }, line: { type: 'none' } });
}

// Full-bleed amber panel (a zero-skew parallelogram in the source deck).
function band(s, x, y, w, h) {
  box(s, x, y, w, h, C.accent);
}

function hline(s, x, y, w, color, width) {
  s.addShape('line', { x, y, w, h: 0, line: { color, width } });
}

function vline(s, x, y, h, color, width) {
  s.addShape('line', { x, y, w: 0, h, line: { color, width } });
}

// Text block. `st` is one of the S.* presets, `more` overrides individual keys.
function txt(s, body, x, y, w, h, st, more) {
  const o = Object.assign({}, st);
  if (more) Object.keys(more).forEach(k => { if (more[k] !== undefined) o[k] = more[k]; });
  const paras = Array.isArray(body) ? body : [body];
  const runs = paras.map((t, i) => ({
    text: t,
    options: { breakLine: i < paras.length - 1 },
  }));
  s.addText(runs, {
    x, y, w, h,
    fontSize: o.size,
    fontFace: o.face || F.body,
    color: o.color || C.white,
    bold: !!o.bold,
    italic: !!o.ital,
    charSpacing: o.spc,
    align: o.align || 'left',
    valign: o.mid ? 'middle' : 'top',
    wrap: !o.nowrap,
    lineSpacingMultiple: o.lh,
    rotate: o.rot,
  });
}

// Short vertical rule at the slide edge plus the sideways strapline below it.
function sideTag(s, x, y, o) {
  vline(s, x, y, o.h, o.rule, 2.25);
  txt(s, TAG, x - 0.9, y + o.h + 0.838, 1.808, 0.259, S.tag, { color: o.color, align: o.align, rot: 90 });
}

// "1) first service" + hairline + grey note paragraph - the deck's list unit.
function item(s, x, y, label, note, o) {
  o = o || {};
  txt(s, label, x, y, o.lw, 0.347, S.item, { color: o.lc });
  if (o.rw) hline(s, x + o.rx, y + 0.247, o.rw, o.rc || C.line, 1.5);
  txt(s, note, x, y + (o.nd || 0.447), o.nw, o.nh || 0.685, S.note, { color: o.nc });
}

// Stand-in for a photo/mock-up bitmap from the original deck.
function photo(s, x, y, w, h) {
  box(s, x, y, w, h, C.pale);
  s.addText('[image]', {
    x, y, w, h, fontSize: 11, fontFace: F.body, color: C.mid,
    align: 'center', valign: 'middle',
  });
}

/* -------------------------------------------------------------- charts */

// Ring chart used twice on the "our chart" slide (75% hole, no legend).
function donut(s, x, y, w, h, colors) {
  s.addChart('doughnut', [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr'], values: [8.2, 3.2, 1.4] }], {
    x, y, w, h,
    holeSize: 75,
    chartColors: colors,
    dataBorder: { pt: 1.5, color: C.white },
    showLegend: false, showTitle: false, showValue: false,
    firstSliceAng: 0,
  });
}

// Stacked-area chart on the statistics slide.
function areaChart(s, x, y, w, h) {
  const months = ['5/1/2002', '6/1/2002', '7/1/2002', '8/1/2002', '9/1/2002'];
  const data = [
    { name: 'Series 1', labels: months, values: [32, 32, 28, 12, 15] },
    { name: 'Series 2', labels: months, values: [12, 12, 12, 21, 28] },
  ];
  const opts = {
    x, y, w, h,
    barGrouping: 'stacked',
    chartColors: [C.accent, C.gray],
    showLegend: false, showTitle: false,
    catAxisLabelFontSize: 8, catAxisLabelColor: C.gray, catAxisLabelFontFace: F.body,
    catAxisLineColor: C.pale, catGridLine: { style: 'none' },
    valAxisLabelFontSize: 8, valAxisLabelColor: C.gray, valAxisLabelFontFace: F.body,
    valAxisMaxVal: 50, valAxisMajorUnit: 5, valAxisLineColor: C.line,
    valGridLine: { color: C.line, style: 'solid', size: 0.75 },
  };
  // The array form makes pptxgenjs emit crossBetween="midCat" so the plot
  // touches both ends of the axis, exactly like the source chart.
  s.addChart([{ type: 'area', data, options: opts }], opts);
}

/* ------------------------------------------------- infographic artwork */

const DARK_AMBER = 'C18600';   // shaded amber used for 3-D faces / hub
const DARK_GRAY  = '333333';   // shaded grey, likewise

function poly(s, points, color) {
  s.addShape('custGeom', {
    x: 0, y: 0, w: 13.333, h: 7.5, fill: { color }, line: { type: 'none' },
    points: points.map(p => ({ x: p[0], y: p[1] })).concat([{ close: true }]),
  });
}

function disc(s, cx, cy, r, color) {
  s.addShape('ellipse', { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, fill: { color }, line: { type: 'none' } });
}

// The gear has 9 teeth on a 40-degree pitch; each quadrant of the graphic is a
// 90-degree wedge cut out of that profile, so teeth straddle the seams.
const GEAR = { tip: 2.04, root: 1.53, pitch: 40, width: 22, phase: -1.6 };

function gearRadius(deg) {
  const t = ((deg - GEAR.phase) % GEAR.pitch + GEAR.pitch) % GEAR.pitch;
  return t < GEAR.width ? GEAR.tip : GEAR.root;
}

function gearQuadrant(s, cx, cy, a0, color) {
  const at = (r, d) => [cx + r * Math.cos(d * Math.PI / 180), cy + r * Math.sin(d * Math.PI / 180)];
  const pts = [at(gearRadius(a0), a0)];
  for (let d = Math.ceil(a0); d <= a0 + 90; d++) {
    if (gearRadius(d) !== gearRadius(d - 1)) {
      pts.push(at(gearRadius(d - 1), d - 0.5), at(gearRadius(d), d - 0.5));
    }
  }
  pts.push(at(gearRadius(a0 + 89.9), a0 + 90), [cx, cy]);
  poly(s, pts, color);
}

// Slide 28: four gear quadrants around a recessed hub, joined by puzzle knobs.
function gearPuzzle(s, cx, cy) {
  const quads = [
    { a0: 180, color: C.gray,   dark: DARK_GRAY },   // 02, top-left
    { a0: 270, color: C.accent, dark: DARK_AMBER },  // 03, top-right
    { a0: 0,   color: C.gray,   dark: DARK_GRAY },   // 04, bottom-right
    { a0: 90,  color: C.accent, dark: DARK_AMBER },  // 01, bottom-left
  ];
  quads.forEach(q => gearQuadrant(s, cx, cy, q.a0, q.color));
  // recessed hub: a darker ring of the same four colours, then the white bore
  quads.forEach(q => s.addShape('pie', {
    x: cx - 0.71, y: cy - 0.71, w: 1.42, h: 1.42,
    angleRange: [q.a0, q.a0 + 90], fill: { color: q.dark }, line: { type: 'none' },
  }));
  disc(s, cx, cy, 0.44, C.white);
  // puzzle knobs straddling the four seams
  disc(s, cx, cy - 1.11, 0.27, C.accent);
  disc(s, cx - 1.11, cy, 0.27, C.gray);
  disc(s, cx, cy + 1.11, 0.27, C.accent);
  disc(s, cx + 1.11, cy, 0.27, C.gray);
}

// Slide 29: four-piece jigsaw shaped like a house (two roof halves + two tiles).
function housePuzzle(s) {
  const MID = 6.81, APEX = 2.41, EAVE = 4.78;      // ridge line / apex / eave
  const LTIP = 4.81, RTIP = 8.74, RTIP_Y = 4.24;   // roof tips
  const BL = 5.56, BR = 8.02, BB = 6.33;           // base tiles
  const DX = -0.25, DY = 0.17;                     // oblique 3-D offset

  // extruded faces behind the front surfaces
  poly(s, [[LTIP, EAVE], [LTIP + DX, EAVE + DY], [MID + DX, EAVE + DY], [MID, EAVE]], DARK_AMBER);
  poly(s, [[MID, APEX], [MID + DX, APEX + DY], [LTIP + DX, EAVE + DY], [LTIP, EAVE]], DARK_AMBER);
  poly(s, [[BL, EAVE], [BL + DX, EAVE + DY], [BL + DX, BB + DY], [BL, BB]], DARK_GRAY);
  poly(s, [[BL, BB], [BL + DX, BB + DY], [MID + DX, BB + DY], [MID, BB]], DARK_GRAY);
  poly(s, [[MID, BB], [MID + DX, BB + DY], [BR + DX, BB + DY], [BR, BB]], DARK_AMBER);

  // 02 - left roof half, stepped for the chimney
  poly(s, [[MID, APEX], [MID, EAVE], [LTIP, EAVE], [5.74, 3.62], [5.74, 3.00], [6.16, 2.80]], C.accent);
  // 03 - right roof half
  poly(s, [[MID, APEX], [RTIP, RTIP_Y], [8.05, RTIP_Y], [8.05, EAVE], [MID, EAVE]], C.gray);
  // 01 / 04 - base tiles
  poly(s, [[BL, EAVE], [MID, EAVE], [MID, BB], [BL, BB]], C.gray);
  poly(s, [[MID, EAVE], [BR, EAVE], [BR, BB], [MID, BB]], C.accent);

  disc(s, MID, 3.45, 0.27, C.gray);      // 03 into 02
  disc(s, 6.15, EAVE, 0.27, C.accent);   // 02 into 01
  disc(s, 7.42, EAVE, 0.27, C.accent);   // 04 into 03
  disc(s, MID, 5.62, 0.27, C.gray);      // 01 into 04
}

/* --------------------------------------------------------------- slides */

function slide01(s) {
  box(s, 0, 0, 13.333, 7.5, C.gray, 15);
  txt(s, LOGO, 4.095, 2.917, 5.143, 1.212, S.mark, { size: 66, align: 'center' });
  txt(s, 'presentation template', 4.259, 3.988, 3.559, 0.417, { size: 14, color: C.pale, spc: 3, lh: 1.5 });
  sideTag(s, 12.796, 0, { h: 0.599, rule: C.white, color: C.faint, align: 'right' });
}

function slide02(s) {
  band(s, 11.181, 0, 2.153, 7.5);
  hline(s, 0.85, 6.833, 8.95, C.gray, 2.25);
  txt(s, 'about our company', 0.744, 6.45, 4.317, 0.303, S.kicker);
  txt(s, TITLE, 6.103, 1.615, 3.559, 1.313, S.title);
  txt(s, EYE, 6.103, 1.203, 3.104, 0.326, S.eyebrow);
  txt(s, LOGO, 11.865, 6.3, 1.146, 0.909, S.logo);
  sideTag(s, 12.925, 0, { h: 0.901, rule: C.gray, color: C.gray });
  txt(s, [T.p3, '', T.p2, '', T.p15], 6.103, 3.465, 3.559, 2.598, S.body);
}

function slide03(s) {
  box(s, 0.917, 0.843, 1.908, 4.889, C.gray, 20);
  txt(s, LOGO, 1.17, 4.618, 1.146, 0.909, S.title, { color: C.accent });
  sideTag(s, 1.296, 0.843, { h: 0.599, rule: C.white, color: C.faint, align: 'right' });
  txt(s, [TITLE_A, TITLE_B], 10.027, 1.296, 2.455, 2.121, S.title);
  txt(s, EYE2, 10.027, 0.916, 2.455, 0.326, S.eyebrow);
  txt(s, T.p4, 10.027, 3.749, 2.455, 1.589, S.body);
  hline(s, 10.071, 5.704, 2.287, C.gray, 2.25);
  txt(s, 'PLACEHOLDER', 3.177, 6.447, 6.979, 0.483, S.note, { ital: true, align: 'center' });
}

function slide04(s) {
  band(s, 10.514, -0, 2.153, 6.847);
  txt(s, LOGO, 11.198, 5.588, 1.146, 0.909, S.logo);
  sideTag(s, 12.258, 0, { h: 0.901, rule: C.gray, color: C.gray });
  txt(s, TITLE, 0.934, 4.89, 3.559, 1.313, S.title);
  txt(s, EYE, 0.934, 4.479, 3.104, 0.326, S.eyebrow);
  txt(s, 'PLACEHOLDER', 5.439, 5.063, 3.825, 1.084, S.body);
  hline(s, 0, 6.819, 9.264, C.gray, 2.25);
}

function slide05(s) {
  box(s, 4.667, 5.443, 5.01, 1.351, C.gray, 20);
  txt(s, TITLE, 0.987, 1.304, 5.54, 0.909, S.title);
  txt(s, EYE, 0.987, 0.893, 3.104, 0.326, S.eyebrow);
  txt(s, LOGO, 11.83, 6.319, 1.146, 0.909, S.mark, { size: 24 });
  txt(s, ['PLACEHOLDER', 'PLACEHOLDER', '', 'PLACEHOLDER'], 0.987, 2.711, 5.346, 2.094, S.body);
  hline(s, 0, 6.777, 4.23, C.gray, 2.25);
  txt(s, 'about our company', 2.227, 6.262, 2.109, 0.303, S.kicker, { align: 'right' });
  txt(s, T.p16, 5.122, 5.675, 4.101, 0.887, S.note, { ital: true, color: C.faint });
}

function slide06(s) {
  band(s, -0, 0, 2.153, 7.5);
  txt(s, LOGO, 0.323, 6.3, 1.146, 0.909, S.title);
  sideTag(s, 0.409, 0, { h: 0.901, rule: C.gray, color: C.gray, align: 'right' });
  txt(s, [TITLE_A, TITLE_B], 10.056, 1.711, 2.455, 2.121, S.title);
  txt(s, EYE2, 10.056, 1.331, 2.455, 0.326, S.eyebrow);
  txt(s, T.p4, 10.056, 4.348, 2.455, 1.589, S.body);
  hline(s, 10.1, 6.494, 2.287, C.gray, 2.25);
}

function slide07(s) {
  band(s, 11.181, 0, 2.153, 7.5);
  sideTag(s, 11.637, 0, { h: 0.901, rule: C.gray, color: C.gray, align: 'right' });
  item(s, 1.048, 2.979, '1) first service', T.p10, { lw: 1.863, nw: 3.497, nh: 0.483, rx: 1.377, rw: 1.998 });
  item(s, 1.048, 4.36, '2) second service', T.p10, { lw: 1.863, nw: 3.497, nh: 0.483, rx: 1.69, rw: 1.685 });
  item(s, 1.048, 5.742, '3) third service', T.p10, { lw: 1.863, nw: 3.559, nh: 0.483, rx: 1.461, rw: 1.913 });
  txt(s, TITLE, 1.032, 1.144, 6.392, 1.043, S.title, { size: 28 });
  txt(s, EYE, 1.032, 0.732, 3.104, 0.326, S.eyebrow);
  txt(s, ['PLACEHOLDER', 'nianitasi hanili sanibinaiat atil'], 8.727, 1.299, 2.037, 0.887, S.note, { ital: true, align: 'right' });
}

function slide08(s) {
  item(s, 5.368, 3.616, '1) first service', T.p6, { lw: 1.821, nw: 2.912, rx: 1.337, rw: 1.383 });
  item(s, 9.113, 3.616, '2) second service', T.p6, { lw: 1.821, nw: 2.912, rx: 1.657, rw: 1.07 });
  item(s, 5.368, 5.263, '3) third service', T.p6, { lw: 1.821, nw: 2.912, rx: 1.426, rw: 1.294 });
  item(s, 9.113, 5.263, '4) fourth service', T.p6, { lw: 1.821, nw: 2.912, rx: 1.554, rw: 1.173 });
  txt(s, TITLE, 5.394, 1.509, 5.54, 0.909, S.title);
  txt(s, EYE, 5.394, 1.097, 3.104, 0.326, S.eyebrow);
  hline(s, 0, 2.223, 4.087, C.gray, 2.25);
  txt(s, 'about our service', 2.083, 1.787, 2.109, 0.303, S.kicker, { align: 'right' });
  band(s, 12.261, -0, 1.072, 2.223);
  box(s, 2.161, 5.49, 1.922, 2.01, C.gray, 10);
  txt(s, LOGO, 2.65, 5.739, 1.146, 0.909, S.mark, { size: 24 });
}

function slide09(s) {
  box(s, 7.448, 0.688, 1.908, 4.889, C.gray, 10);
  txt(s, LOGO, 7.701, 4.463, 1.146, 0.909, S.title, { color: C.accent });
  sideTag(s, 7.828, 0.688, { h: 0.599, rule: C.white, color: C.faint, align: 'right' });
  item(s, 10.233, 1.227, '1) first service', T.p1, { lw: 1.431, nw: 2.232, rx: 1.343, rw: 0.811 });
  item(s, 10.233, 2.693, '2) second service', T.p1, { lw: 1.744, nw: 2.232, rx: 1.64, rw: 0.514 });
  item(s, 10.233, 4.206, '3) third service', T.p1, { lw: 1.431, nw: 2.271, rx: 1.479, rw: 0.676 });
  txt(s, [TITLE_A, TITLE_B], 1.097, 1.567, 2.455, 2.121, S.title);
  txt(s, EYE2, 1.097, 1.187, 2.455, 0.326, S.eyebrow);
  txt(s, T.p4, 1.097, 4.224, 2.455, 1.589, S.body);
  hline(s, 1.142, 6.412, 2.287, C.gray, 2.25);
}

function slide10(s) {
  band(s, -0.001, 0.845, 2.153, 5.81);
  txt(s, LOGO, 0.323, 5.455, 1.146, 0.909, S.title);
  sideTag(s, 0.443, 0.845, { h: 0.901, rule: C.gray, color: C.gray, align: 'right' });
  item(s, 2.922, 5.455, '1) first service', T.p5, { lw: 1.821, nw: 2.912, nh: 0.483, rx: 1.338, rw: 1.383 });
  item(s, 6.271, 5.455, '2) second service', T.p5, { lw: 1.821, nw: 2.912, nh: 0.483, rx: 1.657, rw: 1.07 });
  item(s, 9.619, 5.437, '3) third service', T.p5, { lw: 1.821, nw: 2.912, nh: 0.483, rx: 1.448, rw: 1.281 });
  txt(s, TITLE, 8.876, 1.611, 3.559, 1.313, S.title);
  txt(s, EYE, 8.876, 1.2, 3.104, 0.326, S.eyebrow);
  txt(s, ['PLACEHOLDER', 'PLACEHOLDER'], 8.876, 3.322, 3.559, 1.084, S.body);
}

function slide11(s) {
  band(s, 8.093, 0, 5.241, 7.5);
  txt(s, LOGO, 11.865, 6.3, 1.146, 0.909, S.logo);
  sideTag(s, 12.925, 0, { h: 0.901, rule: C.gray, color: C.gray });
  item(s, 4.596, 1.459, '1) first service', T.p11, { lw: 1.431, nw: 2.232 });
  item(s, 4.596, 3.03, '2) second service', T.p11, { lw: 1.744, nw: 2.232 });
  item(s, 4.596, 4.601, '3) third service', T.p11, { lw: 1.431, nw: 2.271 });
  txt(s, [TITLE_A, TITLE_B], 1.083, 1.74, 2.455, 2.121, S.title);
  txt(s, EYE2, 1.083, 1.36, 2.455, 0.326, S.eyebrow);
  txt(s, ['PLACEHOLDER', 'asaiai contanseci'], 1.083, 4.397, 2.663, 1.336, S.body);
  hline(s, 0, 6.77, 6.75, C.gray, 2.25);
}

function slide12(s) {
  box(s, 10.571, 6.788, 1.788, 0.783, C.gray);
  box(s, 6.638, 0.917, 1.788, 1.788, C.accent);
  hline(s, 0, 6.812, 8.425, C.gray, 2.25);
  txt(s, 'about our portfolio', 4.708, 6.386, 1.725, 0.303, S.kicker, { align: 'right' });
  txt(s, TITLE, 1.385, 1.775, 3.559, 1.313, S.title);
  txt(s, EYE, 1.385, 1.363, 3.104, 0.326, S.eyebrow);
  txt(s, [T.p3, '', T.p2], 1.385, 3.574, 3.559, 2.094, S.body);
  txt(s, LOGO, 10.892, 3.316, 1.146, 0.909, S.title, { align: 'center' });
  item(s, 6.901, 1.143, '1) portfolio', T.p12, { lw: 1.303, nw: 1.303, nh: 0.887 });
  item(s, 10.886, 1.143, '2) portfolio', T.p12, { lw: 1.303, nw: 1.303, nh: 0.887 });
  item(s, 8.878, 5.207, '3) portfolio', T.p12, { lw: 1.303, nw: 1.303, nh: 0.887 });
}

function slide13(s) {
  txt(s, 'PLACEHOLDER', 1.007, 6.653, 8.538, 0.304, S.eyebrow, { size: 9 });
  item(s, 6.633, 1.994, '1) first portfolio', T.p5, { lw: 1.821, nw: 2.912, nh: 0.483, rx: 1.456, rw: 1.264 });
  item(s, 6.633, 3.466, '2) second portfolio', T.p5, { lw: 1.821, nw: 2.912, nh: 0.483, rx: 1.725, rw: 1.002 });
  txt(s, TITLE, 1.084, 1.158, 3.781, 1.313, S.title);
  txt(s, EYE, 1.084, 0.747, 3.104, 0.326, S.eyebrow);
  box(s, 10.516, 0, 2.817, 3.382, C.gray);
  item(s, 6.633, 4.984, '3) third portfolio', T.p5, { lw: 1.821, nw: 2.912, nh: 0.483, rx: 1.527, rw: 1.201 });
  txt(s, LOGO, 11.119, 6.586, 1.898, 0.438, S.mark);
  sideTag(s, 12.838, 0, { h: 0.599, rule: C.white, color: C.faint });
}

function slide14(s) {
  band(s, -0, 0, 2.153, 7.5);
  txt(s, LOGO, 0.323, 6.3, 1.146, 0.909, S.title);
  sideTag(s, 0.409, 0, { h: 0.901, rule: C.gray, color: C.gray, align: 'right' });
  item(s, 9.189, 1.189, '1) first portfolio', [T.p13, T.p14], { lw: 1.821, nw: 2.912, rx: 1.456, rw: 1.264 });
  item(s, 9.189, 3.206, '2) second portfolio', [T.p13, T.p14], { lw: 1.821, nw: 2.912, rx: 1.725, rw: 1.002 });
  item(s, 9.189, 5.179, '3) third portfolio', [T.p13, T.p14], { lw: 1.821, nw: 2.912, rx: 1.526, rw: 1.201 });
}

function slide15(s) {
  item(s, 4.453, 3.819, '1) portfolio', T.p17, { lw: 1.13, nw: 2.158, rx: 1.13, rw: 0.917 });
  item(s, 4.451, 5.364, '2) portfolio', T.p17, { lw: 1.13, nw: 2.158, rx: 1.131, rw: 0.917 });
  txt(s, TITLE, 7.569, 1.498, 5.288, 0.909, S.title);
  txt(s, EYE, 7.569, 1.086, 3.104, 0.326, S.eyebrow);
  hline(s, 7.188, 3.175, 6.146, C.gray, 2.25);
}

function slide16(s) {
  band(s, 6.229, 3.958, 3.061, 3.542);
  hline(s, 0, 6.741, 4.991, C.gray, 2.25);
  txt(s, TITLE, 1.316, 1.775, 3.559, 1.313, S.title);
  txt(s, EYE, 1.316, 1.363, 3.104, 0.326, S.eyebrow);
  txt(s, [T.p3, '', T.p2], 1.316, 3.574, 3.559, 2.094, S.body);
  item(s, 9.691, 0.769, '1) first portfolio', 'hani hanalai banes hanal banis anuban banianitasi', { lw: 1.821, nw: 2.912, nh: 0.281, nd: 0.392, rx: 1.455, rw: 1.264 });
  item(s, 6.68, 4.838, '2) portfolio', ['PLACEHOLDER', '', 'PLACEHOLDER'], { lw: 1.13, nw: 2.158, nh: 1.493, nd: 0.489, rx: 1.131, rw: 0.917, rc: C.gray, nc: C.gray });
}

function slide17(s) {
  box(s, 3.625, 0.792, 1.908, 4.889, C.gray, 10);
  txt(s, LOGO, 3.878, 4.567, 1.146, 0.909, S.title, { color: C.accent });
  sideTag(s, 4.005, 0.792, { h: 0.599, rule: C.white, color: C.faint, align: 'right' });
  txt(s, [TITLE_A, TITLE_B], 9.777, 1.77, 2.455, 2.121, S.title);
  txt(s, EYE2, 9.777, 1.391, 2.455, 0.326, S.eyebrow);
  txt(s, T.p4, 9.777, 4.287, 2.455, 1.589, S.body);
  hline(s, 9.821, 6.401, 2.287, C.gray, 2.25);
  item(s, 0.838, 1.391, '1) our service', T.p18, { lw: 1.431, nw: 2.079, nh: 1.089, nd: 0.516, rx: 1.342, rw: 0.639 });
  item(s, 0.838, 3.881, '2) our service', T.p18, { lw: 1.431, nw: 2.079, nh: 1.089, nd: 0.517, rx: 1.342, rw: 0.639 });
}

function slide18(s) {
  band(s, 11.181, 0.792, 2.153, 6);
  txt(s, LOGO, 11.865, 5.592, 1.146, 0.909, S.logo);
  sideTag(s, 12.925, 0.792, { h: 0.901, rule: C.gray, color: C.gray });
  txt(s, TITLE, 6.3, 1.835, 3.559, 1.313, S.title);
  txt(s, EYE, 6.3, 1.424, 3.104, 0.326, S.eyebrow);
  txt(s, [T.p3, '', T.p2], 6.3, 3.727, 3.559, 2.094, S.body);
  hline(s, 4.979, 6.764, 4.75, C.gray, 2.25);
  box(s, 1.415, 5.647, 3.087, 0.839, C.gray, 10);
  item(s, 1.847, 5.746, '1) your name', T.p7, { lw: 1.264, nw: 2.237, nh: 0.281, nd: 0.291, rx: 1.229, rw: 0.819, rc: C.pale, lc: C.accent, nc: C.line });
}

function slide19(s) {
  item(s, 7.743, 4.266, '2) your name', T.p8, { lw: 1.264, nw: 1.74, nh: 1.089, nd: 0.51, rx: 1.229, rw: 0.403 });
  item(s, 10.37, 4.266, '3) your name', T.p8, { lw: 1.264, nw: 1.74, nh: 1.089, nd: 0.51, rx: 1.229, rw: 0.403 });
  txt(s, 'PLACEHOLDER', 7.535, 6.483, 3.856, 0.304, S.eyebrow, { size: 9, align: 'right' });
  txt(s, ['CON', 'TRUZ'], 11.844, 6.453, 1.146, 0.774, S.mark);
  txt(s, [TITLE_A, TITLE_B], 0.919, 1.567, 2.455, 2.121, S.title);
  txt(s, EYE2, 0.919, 1.187, 2.455, 0.326, S.eyebrow);
  txt(s, T.p4, 0.919, 4.276, 2.455, 1.589, S.body);
  hline(s, 0, 6.693, 3.354, C.gray, 2.25);
  box(s, 4.187, 5.892, 2.838, 0.839, C.gray, 10);
  item(s, 4.55, 5.99, '1) your name', T.p7, { lw: 1.264, nw: 2.237, nh: 0.281, nd: 0.291, rx: 1.23, rw: 0.819, rc: C.pale, lc: C.accent, nc: C.line });
}

function slide20(s) {
  hline(s, 4.013, 6.714, 4.383, C.gray, 2.25);
  txt(s, TITLE, 4.887, 1.694, 3.559, 1.313, S.title);
  txt(s, EYE, 4.887, 1.283, 3.104, 0.326, S.eyebrow);
  txt(s, ['PLACEHOLDER', '', T.p2, '', T.p15], 4.887, 3.544, 3.559, 2.346, S.body);
  box(s, 0.926, 5.89, 3.087, 0.839, C.gray, 10);
  item(s, 1.358, 5.989, '1) your name', T.p7, { lw: 1.264, nw: 2.237, nh: 0.281, nd: 0.291, rx: 1.229, rw: 0.819, rc: C.pale, lc: C.accent, nc: C.line });
  box(s, 9.32, 5.89, 3.087, 0.839, C.gray, 10);
  item(s, 9.752, 5.989, '2) your name', T.p7, { lw: 1.264, nw: 2.237, nh: 0.281, nd: 0.291, rx: 1.264, rw: 0.784, rc: C.pale, lc: C.accent, nc: C.line });
}

function slide21(s) {
  item(s, 1.036, 5.919, '1) your name', T.p9, { lw: 1.264, nw: 2.237, nh: 0.483, nd: 0.467, rx: 1.229, rw: 0.819 });
  item(s, 4.044, 5.919, '2) your name', T.p9, { lw: 1.264, nw: 2.237, nh: 0.483, nd: 0.467, rx: 1.229, rw: 0.819 });
  item(s, 7.052, 5.919, '3) your name', T.p9, { lw: 1.264, nw: 2.237, nh: 0.483, nd: 0.467, rx: 1.229, rw: 0.819 });
  item(s, 10.06, 5.919, '4) your name', T.p9, { lw: 1.418, nw: 2.237, nh: 0.483, nd: 0.467, rx: 1.253, rw: 0.796 });
  txt(s, TITLE, 0.747, 1.057, 5.54, 0.909, S.title);
  txt(s, EYE, 0.747, 0.565, 3.104, 0.326, S.eyebrow);
  txt(s, LOGO, 11.312, 0.355, 1.677, 0.438, S.mark);
}

function slide22(s) {
  photo(s, 1.023, 1.099, 5.752, 3.357);
  band(s, 11.181, 0, 2.153, 7.5);
  txt(s, LOGO, 11.865, 6.3, 1.146, 0.909, S.logo);
  sideTag(s, 12.925, 0, { h: 0.901, rule: C.gray, color: C.gray });
  item(s, 7.988, 1.111, '1) first service', T.p1, { lw: 1.431, nw: 2.232, rx: 1.343, rw: 0.811 });
  item(s, 7.988, 2.577, '2) second service', T.p1, { lw: 1.744, nw: 2.232, rx: 1.639, rw: 0.514 });
  item(s, 7.988, 4.091, '3) third service', T.p1, { lw: 1.431, nw: 2.271, rx: 1.478, rw: 0.676 });
  item(s, 7.988, 5.6, '4) fourth service', T.p1, { lw: 1.561, nw: 2.271, rx: 1.561, rw: 0.593 });
  txt(s, TITLE, 1.248, 5.725, 5.465, 0.909, S.title);
  txt(s, EYE, 1.248, 5.272, 3.104, 0.326, S.eyebrow);
}

function slide23(s) {
  photo(s, 9.339, 0.858, 2.868, 5.785);
  box(s, 5.73, 5.443, 5.01, 1.351, C.gray, 20);
  hline(s, 0, 6.777, 5.118, C.gray, 2.25);
  txt(s, 'about our mockup', 3.114, 6.262, 2.109, 0.303, S.kicker, { align: 'right' });
  txt(s, T.p16, 6.185, 5.675, 4.101, 0.887, S.note, { ital: true, color: C.faint });
  item(s, 6.217, 2.039, '1) first service', T.p1, { lw: 1.431, nw: 2.232, rx: 1.343, rw: 0.811 });
  item(s, 6.217, 3.695, '2) second service', T.p1, { lw: 1.744, nw: 2.232, rx: 1.64, rw: 0.514 });
  txt(s, TITLE, 1.203, 1.457, 3.559, 1.313, S.title);
  txt(s, EYE, 1.203, 1.046, 3.104, 0.326, S.eyebrow);
  txt(s, [T.p3, '', T.p2], 1.203, 3.349, 3.559, 2.094, S.body);
}

function slide24(s) {
  photo(s, 2.94, 1.022, 3.699, 5.457);
  band(s, -0.001, 0.845, 2.153, 5.81);
  txt(s, LOGO, 0.323, 5.455, 1.146, 0.909, S.title);
  sideTag(s, 0.443, 0.845, { h: 0.901, rule: C.gray, color: C.gray, align: 'right' });
  item(s, 7.521, 5.089, '1) first service', T.p1, { lw: 1.431, nw: 2.232, rx: 1.343, rw: 0.811 });
  item(s, 10.295, 5.089, '2) second service', T.p1, { lw: 1.744, nw: 2.232, rx: 1.64, rw: 0.514 });
  txt(s, TITLE, 7.521, 1.782, 3.699, 1.313, S.title);
  txt(s, EYE, 7.521, 1.371, 3.104, 0.326, S.eyebrow);
  txt(s, ['PLACEHOLDER', 'PLACEHOLDER'], 7.521, 3.492, 5.006, 1.084, S.body);
}

function slide25(s) {
  photo(s, 6.065, 0.632, 5.665, 5.665);
  box(s, 10.495, 4.722, 2.068, 2.049, C.gray, 10);
  txt(s, LOGO, 11.191, 5.625, 1.146, 0.909, S.mark, { size: 24 });
  txt(s, TITLE, 1.27, 1.642, 3.559, 1.313, S.title);
  txt(s, EYE, 1.27, 1.231, 3.104, 0.326, S.eyebrow);
  txt(s, [T.p3, '', T.p2], 1.27, 3.493, 3.559, 2.094, S.body);
  hline(s, 0, 6.777, 9.789, C.gray, 2.25);
}

function slide26(s) {
  band(s, 11.181, 0, 2.153, 7.5);
  txt(s, LOGO, 11.865, 6.3, 1.146, 0.909, S.logo);
  sideTag(s, 12.925, 0, { h: 0.901, rule: C.gray, color: C.gray });
  donut(s, 0.191, 0.787, 4.24, 2.827, [C.gray, C.amber2, C.line]);
  donut(s, 0.191, 3.899, 4.24, 2.827, [C.accent, C.gray, C.line]);
  item(s, 4.29, 1.312, '1) our chart', T.p8, { lw: 1.264, nw: 1.74, nh: 1.089, nd: 0.51, rx: 1.12, rw: 0.513 });
  item(s, 4.29, 4.589, '2) our chart', T.p8, { lw: 1.264, nw: 1.74, nh: 1.089, nd: 0.509, rx: 1.15, rw: 0.483 });
  txt(s, [TITLE_A, TITLE_B], 7.32, 1.692, 2.455, 2.121, S.title);
  txt(s, EYE2, 7.32, 1.312, 2.455, 0.326, S.eyebrow);
  txt(s, T.p4, 7.32, 4.209, 2.455, 1.589, S.body);
  hline(s, 7.365, 6.725, 3.816, C.gray, 2.25);
  txt(s, '01', 1.981, 1.88, 0.66, 0.64, { size: 32, color: C.accent, mid: true, nowrap: true });
  txt(s, '02', 1.934, 4.992, 0.752, 0.64, { size: 32, color: C.gray, mid: true, nowrap: true });
}

function slide27(s) {
  areaChart(s, 0.792, 0.694, 7.716, 3.984);
  txt(s, TITLE, 1.125, 5.548, 5.733, 0.909, S.title);
  txt(s, EYE, 1.172, 5.136, 3.104, 0.326, S.eyebrow);
  txt(s, 'PLACEHOLDER', 7.471, 5.625, 4.654, 0.831, S.body);
  hline(s, 0, 6.915, 12.514, C.gray, 2.25);
  item(s, 9.535, 1.267, '1) first service', T.p1, { lw: 1.431, nw: 2.232, rx: 1.343, rw: 0.811 });
  item(s, 9.535, 2.897, '2) second service', T.p1, { lw: 1.744, nw: 2.232, rx: 1.64, rw: 0.514 });
}

function slide28(s) {
  gearPuzzle(s, 6.771, 4.427);
  txt(s, '02', 5.637, 3.471, 0.616, 0.505, { size: 24, mid: true, nowrap: true });
  txt(s, '03', 7.209, 3.338, 0.614, 0.505, { size: 24, mid: true, nowrap: true });
  txt(s, '01', 5.773, 5.077, 0.546, 0.505, { size: 24, mid: true, nowrap: true });
  txt(s, '04', 7.293, 4.944, 0.647, 0.505, { size: 24, mid: true, nowrap: true });
  txt(s, TITLE, 0.795, 1.072, 5.54, 0.909, S.title);
  txt(s, EYE, 0.795, 0.581, 3.104, 0.326, S.eyebrow);
  item(s, 9.875, 3.054, '3) third service', T.p1, { lw: 1.431, nw: 2.232, rx: 1.38, rw: 0.774 });
  item(s, 9.875, 5.18, '4) fourth service', T.p1, { lw: 1.744, nw: 2.232, rx: 1.513, rw: 0.641 });
  item(s, 1.591, 3.054, '2) second service', T.p1, { lw: 1.744, nw: 2.232, rx: 1.576, rw: 0.578 });
  item(s, 1.591, 5.18, '1) first service', T.p1, { lw: 1.31, nw: 2.232, rx: 1.292, rw: 0.862 });
  txt(s, T.p19, 8.16, 1.065, 4.221, 0.831, S.body);
}

function slide29(s) {
  housePuzzle(s);
  txt(s, '01', 5.874, 5.69, 0.546, 0.505, S.badge);
  txt(s, '03', 7.059, 3.472, 0.614, 0.505, S.badge);
  txt(s, '04', 7.264, 5.175, 0.647, 0.505, S.badge);
  txt(s, '02', 5.661, 3.944, 0.616, 0.505, S.badge);
  txt(s, TITLE, 0.795, 1.072, 5.54, 0.909, S.title);
  txt(s, EYE, 0.795, 0.581, 3.104, 0.326, S.eyebrow);
  txt(s, T.p19, 8.16, 1.065, 4.221, 0.831, S.body);
  item(s, 9.875, 3.139, '3) third service', T.p1, { lw: 1.431, nw: 2.232, rx: 1.38, rw: 0.774 });
  item(s, 9.875, 5.264, '4) fourth service', T.p1, { lw: 1.744, nw: 2.232, rx: 1.513, rw: 0.641 });
  item(s, 1.591, 3.139, '2) second service', T.p1, { lw: 1.744, nw: 2.232, rx: 1.576, rw: 0.578 });
  item(s, 1.591, 5.264, '1) first service', T.p1, { lw: 1.31, nw: 2.232, rx: 1.292, rw: 0.862 });
}

function slide30(s) {
  box(s, 0, 0, 13.333, 7.5, C.gray, 15);
  txt(s, 'THANK YOU !', 3.028, 3.144, 7.278, 1.212, S.mark, { size: 66, align: 'center' });
  sideTag(s, 12.796, 0, { h: 0.599, rule: C.white, color: C.faint, align: 'right' });
}

/* -------------------------------------------------------------- assemble */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'CONTRUZ_16x9', width: 13.333, height: 7.5 });
  pptx.layout = 'CONTRUZ_16x9';
  pptx.author = 'CONTRUZ';
  pptx.title = 'CONTRUZ presentation template';

  const builders = [
    slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
    slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
  ];
  builders.forEach(fn => fn(pptx.addSlide()));

  return pptx.writeFile({ fileName: path.join(__dirname, '09b7ed50-e230-4f0a-ac22-0d8173b155c1_grok_final.pptx') });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
