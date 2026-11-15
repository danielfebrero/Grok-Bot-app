#!/usr/bin/env node
/**
 * "Nursing" — medical presentation template (32 slides, 13.333 x 7.5 in),
 * rebuilt from scratch with pptxgenjs.
 *
 * Raster photos of the original deck are replaced by flat placeholder
 * rectangles; every other element is a native pptxgenjs shape / text / chart.
 *
 * Run:  node 129f2e8b-e9f8-446f-bb63-e24e03b634f7_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ================================================================== *
 * 1. Theme
 * ================================================================== */
const W = 13.333;                  // slide width  (in)
const H = 7.5;                     // slide height (in)

const HEAD = 'Montserrat';         // major (theme) font
const BODY = 'Lato';               // minor (theme) font

const C = {
  blue:      '0F6FC6',             // accent1
  sky:       '009DD9',             // accent2
  cyan:      '0BD0D9',             // accent3
  mint:      '10CF9B',             // accent4
  green:     '7CCA62',             // accent5
  olive:     'A5C249',             // accent6
  navy:      '17406D',             // dk2
  pale:      'DBEFF9',             // lt2
  white:     'FFFFFF',
  offWhite:  'F2F2F2',
  ink:       '404040',             // headline grey
  body:      '595959',             // body-copy grey
  grey:      'BFBFBF',
  lightGrey: 'D9D9D9',
  midGrey:   'A6A6A6',
  iceBlue:   '89DAF8',
  teal:      '10A092',             // "placeholder image" caption colour
  paleBlue:  'B3DDF2',             // lt2 @ 90 % luminance
  // accents at 60 % luminance + 40 % offset (icon chips)
  chipBlue:  '59AAF2',
  chipSky:   '4FCEFF',
  chipCyan:  '5EF0F7',
  chipMint:  '5FF3CA',
  // accents at 20 % luminance + 80 % offset (riser highlights)
  tintBlue:  'C8E3FB',
  tintSky:   'C4EFFF',
  tintCyan:  'C9FAFC',
  tintMint:  'CAFBED',
  // accents at 75 % luminance (icon tiles)
  deepBlue:  '0B5394',
  deepSky:   '0076A3',
  deepCyan:  '089CA3',
  deepMint:  '0C9B74',
};

/* Body copy of the template, verbatim. */
const T = {
  sub:    'Input here for subtitle text',
  one:    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor',
  short:  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed',
  half:   'Lorem ipsum dolor sit amet, consectetur adipiscing elit, ',
  sent:   'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt. ',
  labore: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore. ',
  adi:    'Lorem ipsum dolor sit amet, consectetur adi piscing elit, sed do eiusmod tempor',
  adi2:   'Lorem ipsum dolor sit amet, consectetur adi piscing elit, sed do eiusmod tempor incididunt. ',
  noDot:  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt',
  two:    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt. Lorem ipsum dolor sit amet, ',
  brief:  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt. Lorem ipsum',
  three:  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt. Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt. Lorem ipsum dolor sit amet',
  commodo:'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ',
};
T.four   = T.three + ' Lorem ipsum dolor sit amet, ';
T.five   = T.four + 'consectetur adipiscing elit, sed do eiusmod tempor incididunt. Lorem ipsum dolor sit amet';
T.eiusmod = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod';

const SHADOW = { type: 'outer', color: '000000', opacity: 0.13, blur: 18, offset: 5, angle: 90 };

/* ================================================================== *
 * 2. Drawing helpers
 * ================================================================== */

/** OOXML "adj" percentage -> pptxgenjs rectRadius, in inches. */
const rr = (adj, w, h) => (adj / 100000) * Math.min(w, h);

/** OOXML angles run clockwise from 3 o'clock; author them from 12 o'clock. */
const compass = deg => (deg + 270) % 360;

/** Blend two hex colours (t = 0 -> a, t = 1 -> b). */
function mix(a, b, t) {
  const ch = (s, i) => parseInt(s.substr(i * 2, 2), 16);
  return [0, 1, 2]
    .map(i => Math.round(ch(a, i) + (ch(b, i) - ch(a, i)) * t))
    .map(v => v.toString(16).padStart(2, '0').toUpperCase()).join('');
}

/* Borderless primitives — pptxgenjs draws a default outline unless `line`
 * is omitted entirely, so none of these pass a `line` property. */
function rect(s, x, y, w, h, fill, o) { s.addShape('rect', Object.assign({ x, y, w, h, fill }, o)); }
function card(s, x, y, w, h, color, adj, o) {
  s.addShape('roundRect', Object.assign({ x, y, w, h, fill: { color }, rectRadius: rr(adj, w, h) }, o));
}
function diagCard(s, x, y, w, h, color, adj, o) {
  s.addShape('round2DiagRect', Object.assign({ x, y, w, h, fill: { color }, rectRadius: rr(adj, w, h) }, o));
}
function circle(s, cx, cy, d, o) { s.addShape('ellipse', Object.assign({ x: cx - d / 2, y: cy - d / 2, w: d, h: d }, o)); }
function poly(s, x, y, w, h, color, pts, o) {
  s.addShape('custGeom', Object.assign({ x, y, w, h, fill: { color }, points: pts.concat([{ close: true }]) }, o));
}
function hline(s, x, y, w, color, width) { s.addShape('line', { x, y, w, h: 0, line: { color, width } }); }
function vline(s, x, y, h, color, width) { s.addShape('line', { x, y, w: 0, h, line: { color, width } }); }

/* ---- text ---- */
function txt(s, body, o) {
  s.addText(body, Object.assign({ fontFace: BODY, color: C.body, valign: 'top', margin: [4, 7, 4, 7] }, o));
}
/** Page headline — Montserrat 40 bold, dark grey. */
function title(s, x, y, w, h, body, o) {
  txt(s, body, Object.assign({ x, y, w, h, fontFace: HEAD, fontSize: 40, bold: true, color: C.ink, lineSpacingMultiple: 0.95 }, o));
}
/** Blue kicker line under the headline. */
function kicker(s, x, y, o) {
  txt(s, T.sub, Object.assign({ x, y, w: 3.325, h: 0.404, fontSize: 18, color: C.blue }, o));
}
/** 12 pt paragraph at the deck's 1.3 line spacing. */
function para(s, x, y, w, h, body, o) {
  txt(s, body, Object.assign({ x, y, w, h, fontSize: 12, lineSpacingMultiple: 1.3 }, o));
}
/** 18 pt bold sub-heading. */
function lead(s, x, y, w, h, body, o) {
  txt(s, body, Object.assign({ x, y, w, h, fontSize: 18, bold: true }, o));
}
/** Value + superscript "%" (67%, 75%, 80%). */
function percent(s, x, y, w, h, value, size, color, o) {
  s.addText(
    [{ text: value, options: { fontSize: size, bold: true } },
     { text: '%', options: { fontSize: size, bold: true, superscript: true } }],
    Object.assign({ x, y, w, h, fontFace: BODY, color, align: 'center', valign: 'middle' }, o));
}

/* ---- pictograms (stand-ins for the deck's icon artwork) ---- */
function iconGlyph(s, x, y, d, color) {
  circle(s, x + d / 2, y + d / 2, d * 0.82, { fill: { type: 'none' }, line: { color, width: 1.25 } });
  rect(s, x + d * 0.44, y + d * 0.22, d * 0.12, d * 0.56, { color });
  rect(s, x + d * 0.22, y + d * 0.44, d * 0.56, d * 0.12, { color });
}
function personIcon(s, x, y, w, h, color) {
  circle(s, x + w / 2, y + h * 0.17, h * 0.33, { fill: { color } });
  s.addShape('roundRect', { x, y: y + h * 0.38, w, h: h * 0.62, fill: { color }, rectRadius: w * 0.45 });
}
function checkMark(s, x, y, w, h, color) {
  poly(s, x, y, w, h, color, [
    { x: 0, y: h * 0.45 }, { x: w * 0.15, y: h * 0.27 }, { x: w * 0.38, y: h * 0.58 },
    { x: w * 0.85, y: 0 }, { x: w, y: h * 0.2 }, { x: w * 0.38, y: h },
  ]);
}

/** Stand-in for a raster photo. `caption` shows the deck's "image" prompt. */
function imageBox(s, x, y, w, h, fill, caption) {
  rect(s, x, y, w, h, fill);
  if (caption) imageCaption(s, x, y, w, h);
}
function imageCaption(s, x, y, w, h) {
  txt(s, '[image]', { x, y, w, h, align: 'center', valign: 'middle', fontSize: 16, color: C.teal });
}

/** Half-stadium band used by the SWOT quadrants. */
function pillPanel(s, x, y, w, h, color, roundedSide) {
  const r = h / 2, right = roundedSide === 'right';
  rect(s, right ? x : x + r, y, w - r, h, { color });
  circle(s, right ? x + w - r : x + r, y + r, h, { fill: { color } });
}

/** Icon tile + "Title Here" + copy, on a diagonal-corner panel (slides 3, 12). */
function featureRow(s, o) {
  diagCard(s, o.x, o.y, 5.17, 1.412, C.offWhite, 20000);
  diagCard(s, o.x + o.tileDx, o.y + 0.206, 1, 1, o.tile, 22222);
  iconGlyph(s, o.x + o.tileDx + 0.32, o.y + 0.53, 0.36, C.white);
  txt(s, 'Title Here', { x: o.x + o.textDx, y: o.y + o.titleDy, w: 1.3, h: 0.337, fontSize: 14, bold: true, color: o.label, valign: 'bottom', charSpacing: 1 });
  para(s, o.x + o.textDx, o.y + o.textDy, o.textW, 0.6, o.text, { fontSize: o.fontSize });
}

/** One 3-D riser of the staircase infographic (slide 17). */
function processStep(s, o) {
  const sw = 1.278, sh = 2.054, lw = 1.06, lh = 1.692, cw = 1.242, cap = 0.17;
  poly(s, o.x + 0.794, o.y - cap, lw, lh, o.light,
    [{ x: 0, y: cap }, { x: lw, y: 0 }, { x: lw, y: lh - cap }, { x: 0, y: lh }]);
  poly(s, o.x + 0.415, o.y - cap, cw, cap, o.cap,
    [{ x: 0.38, y: 0 }, { x: cw, y: 0 }, { x: cw - 0.38, y: cap }, { x: 0, y: cap }]);
  poly(s, o.x, o.y, sw, sh, o.main,
    [{ x: 0, y: cap }, { x: sw, y: 0 }, { x: sw, y: sh - cap }, { x: 0, y: sh }]);
  poly(s, o.x + 0.19, o.y + sh - 0.05, 0.38, 0.18, o.cap,
    [{ x: 0.1, y: 0 }, { x: 0.38, y: 0 }, { x: 0.28, y: 0.18 }, { x: 0, y: 0.18 }]);
  txt(s, o.num, { x: o.x + 0.21, y: o.y + 0.123, w: 1.064, h: 0.64, fontSize: 32, color: C.white });
}

/* ================================================================== *
 * 3. Slides
 * ================================================================== */

/* 1 — Nursing (cover) */
function slide01(s) {
  const strips = 64;                                  // accent1 -> 35 % accent1, left to right
  for (let i = 0; i < strips; i++) {
    const t = i / (strips - 1);
    const alpha = t <= 0.2 ? 1 : t <= 0.67 ? 1 - (t - 0.2) / 0.47 * 0.37 : 0.63 - (t - 0.67) / 0.33 * 0.28;
    rect(s, (i * W) / strips, 0, W / strips + 0.02, H, { color: mix(C.white, C.blue, alpha) });
  }
  for (let col = 0; col < 8; col++) {                 // honeycomb, right half
    for (let row = 0; row < 7; row++) {
      s.addShape('hexagon', {
        x: 6.016 + col * 0.8835, y: -0.013 + row * 1.0025 + (col % 2 ? 0.5 : 0),
        w: 1.132, h: 0.997, fill: { type: 'none' },
        line: { color: 'A4B3BB', width: 2, transparency: 65 },
      });
    }
  }
  txt(s, 'Nursing', { x: 1.757, y: 2.312, w: 7.521, h: 1.717, fontFace: HEAD, fontSize: 96, bold: true, color: C.white });
  txt(s, 'Medical Presentation Template', { x: 3.405, y: 4.109, w: 4.472, h: 0.438, fontSize: 20, color: C.white });
  hline(s, 2.021, 4.328, 1.217, C.pale, 4.75);
}

/* 2 — We Are Your Healthy */
function slide02(s) {
  title(s, 3.352, 0.852, 6.629, 0.774, 'We Are Your Healthy', { align: 'center' });
  kicker(s, 5.009, 1.564, { align: 'center' });
  para(s, 1.429, 2.114, 10.475, 0.859,
    T.commodo + 'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.',
    { align: 'center' });

  [
    { x: 0.89,  y: 4.254, color: C.blue, chip: C.chipBlue, chipX: 1.158, tX: 1.102, tY: 4.737, bY: 5.248 },
    { x: 4.831, y: 4.276, color: C.sky,  chip: C.chipSky,  chipX: 5.126, tX: 5.052, tY: 4.781, bY: 5.213 },
    { x: 8.771, y: 4.276, color: C.cyan, chip: C.chipCyan, chipX: 9.055, tX: 8.987, tY: 4.765, bY: 5.204 },
  ].forEach(c => {
    card(s, c.x, c.y, 3.675, 1.932, c.color, 4674);
    card(s, c.chipX, 3.94, 0.762, 0.676, c.chip, 16667);
    iconGlyph(s, c.chipX + 0.19, 4.09, 0.38, C.white);
    lead(s, c.tX, c.tY, 3.325, 0.404, T.sub, { color: C.white });
    para(s, c.tX, c.bY, 3.325, 0.596, T.one, { color: C.white });
  });
}

/* 3 — Medical Tool Development */
function slide03(s) {
  title(s, 1.295, 1.023, 9.063, 0.774, 'Medical Tool Development');
  kicker(s, 1.331, 1.856);
  para(s, 1.295, 2.339, 9.381, 0.596, T.commodo);

  [
    { x: 1.274, y: 3.544, tile: C.navy, label: '0D62AF', text: T.adi },
    { x: 6.828, y: 3.582, tile: C.blue, label: C.blue,   text: T.adi + '. ' },
    { x: 1.295, y: 5.164, tile: C.sky,  label: C.sky,    text: T.adi },
    { x: 6.849, y: 5.202, tile: C.cyan, label: C.cyan,   text: T.adi },
  ].forEach(f => featureRow(s, Object.assign(
    { tileDx: 0.274, textDx: 1.424, titleDy: 0.202, textDy: 0.542, textW: 3.5, fontSize: 12 }, f)));
}

/* 4 — Medical Development */
function slide04(s) {
  title(s, 0.635, 0.645, 5.74, 1.447, 'Medical Development');
  kicker(s, 0.655, 2.155);
  para(s, 0.667, 2.891, 5.74, 0.859,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris');

  [
    { x: 0.874, color: C.blue, iconX: 1.062, tX: 1.627, bX: 1.157 },
    { x: 5.516, color: C.sky,  iconX: 5.704, tX: 6.269, bX: 5.752 },
  ].forEach(c => {
    card(s, c.x, 4.418, 4.296, 1.932, c.color, 4674);
    iconGlyph(s, c.iconX, 4.623, 0.58, C.white);
    lead(s, c.tX, 4.713, 3.325, 0.404, T.sub, { color: C.white });
    para(s, c.bX, 5.332, 3.747, 0.596, T.one, { color: C.white });
  });
}

/* 5 — We Care About Your Health */
function slide05(s) {
  title(s, 2.081, 0.761, 9.171, 0.774, 'We Care About Your Health', { align: 'center' });
  kicker(s, 5.009, 1.564, { align: 'center' });

  [
    { x: 0.694, y: 2.283, color: C.blue, tX: 0.869, tY: 4.845, bY: 5.302, pill: 3.168, pillY: 6.505, text: T.labore },
    { x: 4.829, y: 2.307, color: C.sky,  tX: 5.004, tY: 4.869, bY: 5.328, pill: 7.102, pillY: 6.505, text: T.labore },
    { x: 8.996, y: 2.299, color: C.cyan, tX: 9.194, tY: 4.861, bY: 5.368, pill: 11.379, pillY: 6.521, text: T.labore.trim().slice(0, -1) },
  ].forEach(c => {
    card(s, c.x, c.y, 3.675, 4.733, c.color, 4674, { shadow: SHADOW });
    lead(s, c.tX, c.tY, 3.325, 0.404, T.sub, { color: C.white });
    para(s, c.tX, c.bY, 3.325, 0.859, c.text, { color: C.white });
    card(s, c.pill, c.pillY, 0.934, 0.306, C.offWhite, 50000);
    txt(s, 'Next', { x: c.pill, y: c.pillY, w: 0.934, h: 0.306, fontSize: 11, bold: true, align: 'center', valign: 'middle' });
  });
}

/* 6 — Ensure Treatment Services For People */
function slide06(s) {
  title(s, 6.667, 0.886, 5.486, 1.313, 'Ensure Treatment\nServices For People ', { fontSize: 36 });
  kicker(s, 6.632, 2.331, { align: 'center' });
  para(s, 6.735, 2.992, 5.486, 1.121, T.labore + T.labore);

  card(s, 1.365, 5.198, 4.503, 1.391, C.blue, 6715, { shadow: SHADOW });
  lead(s, 1.603, 5.376, 3.802, 0.404, 'Medical Tittle Text Here', { color: C.white });
  para(s, 1.603, 5.796, 4.0, 0.596, T.one, { color: C.offWhite });

  [
    { x: 6.792, y: 4.307, w: 2.859, color: C.sky,  num: '325+', nX: 7.007,  bX: 7.01, text: T.half },
    { x: 9.762, y: 4.282, w: 2.905, color: C.cyan, num: '425+', nX: 10.006, bX: 9.97, text: T.half.trim().slice(0, -1) },
  ].forEach(c => {
    card(s, c.x, c.y, c.w, 2.283, c.color, 4674);
    txt(s, c.num, { x: c.nX, y: 4.483, w: 1.55, h: 0.584, fontSize: 28, color: C.offWhite, lineSpacingMultiple: 1.13 });
    lead(s, c.nX, 5.15, 2.6, 0.428, 'Medical Text Here', { color: C.offWhite });
    para(s, c.bX, 5.618, 2.618, 0.596, c.text, { color: C.offWhite });
  });
}

/* 7 — Medical Team */
function slide07(s) {
  title(s, 8.714, 1.534, 3.746, 1.447, 'Medical\nTeam ');
  kicker(s, 8.645, 3.077, { align: 'center' });
  para(s, 8.778, 3.825, 4.111, 0.859,
    T.labore + 'Lorem ipsum dolor sit amet, consectetur');

  [
    { x: 1.49,  y: 3.318, color: C.blue, name: 'Dr. Marketa Brown', role: 'Theraphist' },
    { x: 5.115, y: 3.352, color: C.sky,  name: 'Dr. Emily Watson',  role: 'Dentist' },
  ].forEach(c => {
    card(s, c.x, c.y, 3.311, 3.474, c.color, 6517);
    for (let i = 0; i < 3; i++) {
      circle(s, c.x + 1.264 + i * 0.352, 4.758, 0.237, { fill: { color: C.white } });
      txt(s, ['f', 't', 'g'][i], { x: c.x + 1.145 + i * 0.352, y: 4.64, w: 0.237, h: 0.237, fontSize: 9, bold: true, color: c.color, align: 'center', valign: 'middle' });
    }
    txt(s, c.name, { x: c.x + 0.168, y: 4.983, w: 2.895, h: 0.438, fontSize: 20, bold: true, color: C.white, align: 'center' });
    txt(s, c.role, { x: c.x + 0.806, y: 5.331, w: 1.619, h: 0.337, fontSize: 14, color: C.white, align: 'center' });
    para(s, c.x + 0.06, 5.756, 3.056, 0.596, T.short, { color: C.white, align: 'center' });
  });
}

/* 8 — Nurse profile cards */
function slide08(s) {
  [
    { x: 1.032, color: C.blue, chip: C.chipBlue, name: 'Chase Wilson',    tX: 0.9,   rX: 1.675, bX: 1.146, bY: 5.35 },
    { x: 5.04,  color: C.sky,  chip: C.chipSky,  name: 'Hubbert jackson', tX: 4.908, rX: 5.683, bX: 5.138, bY: 5.393 },
    { x: 9.048, color: C.cyan, chip: C.chipCyan, name: 'Peggie Cannon',   tX: 8.916, rX: 9.691, bX: 9.124, bY: 5.407 },
  ].forEach(c => {
    card(s, c.x, 3.872, 3.253, 2.57, c.color, 3876, { shadow: SHADOW });
    circle(s, c.x + 1.626, 4.008, 0.928, { fill: { color: c.chip } });
    iconGlyph(s, c.x + 1.39, 3.772, 0.474, C.white);
    lead(s, c.tX, 4.621, 3.517, 0.404, c.name, { color: C.white, align: 'center' });
    txt(s, 'Lead Nurse', { x: c.rX, y: 4.989, w: 1.967, h: 0.337, fontSize: 14, color: C.white, align: 'center' });
    para(s, c.bX, c.bY, 3.056, 0.596, T.short + ' do', { color: C.white, align: 'center' });
  });
}

/* 9 — Doctor Christine */
function slide09(s) {
  card(s, 8.021, 4.082, 4.646, 2.722, C.blue, 6715, { shadow: SHADOW });
  title(s, 7.989, 1.15, 4.503, 1.447, 'Doctor\nChristine');
  lead(s, 8.021, 2.601, 3.802, 0.404, 'Expert Clinical Psychologist', { color: C.blue });
  para(s, 8.002, 3.229, 4.22, 0.596, T.brief);

  [
    { y: 1.385, bodyY: 1.772, head: '2018 – First Clinical' },
    { y: 3.18,  bodyY: 3.649, head: '2019 – Second Clinical' },
    { y: 4.931, bodyY: 5.363, head: '2020 – Third Clinical' },
  ].forEach(t => {
    lead(s, 0.637, t.y, 3.4, 0.404, t.head, { color: C.body });
    para(s, 0.62, t.bodyY, 3.028, 0.859, T.sent);
  });

  txt(s, 'Skill & Performance', { x: 8.478, y: 4.436, w: 4.014, h: 0.438, fontSize: 20, color: C.white });
  [
    { label: 'Parameter 1', value: '80%', y: 4.966, barY: 5.409, fill: 2.723, valX: 11.301 },
    { label: 'Parameter 2', value: '70%', y: 5.598, barY: 6.041, fill: 2.517, valX: 10.784 },
  ].forEach(b => {
    txt(s, b.label, { x: 8.501, y: b.y, w: 1.47, h: 0.364, fontSize: 12, color: C.white, lineSpacingMultiple: 1.5 });
    txt(s, b.value, { x: b.valX, y: b.y - 0.02, w: 0.851, h: 0.408, fontSize: 14, bold: true, color: C.white, align: 'center' });
    card(s, 8.579, b.barY, 3.468, 0.12, C.white, 50000, { fill: { color: C.white, transparency: 70 } });
    card(s, 8.579, b.barY, b.fill, 0.12, C.white, 50000);
  });
}

/* 10 — How Can We Help You? */
function slide10(s) {
  title(s, 1.96, 0.737, 9.413, 0.774, 'How Can We Help You?', { align: 'center' });
  kicker(s, 4.963, 1.53, { align: 'center' });
  para(s, 2.399, 2.058, 8.556, 0.596, T.three, { align: 'center' });

  [
    { x: 1.43,  y: 3.122, tileX: 0.68,  color: C.blue, tile: C.deepBlue, head: 'Cardiology',      hX: 2.187, bX: 2.177, bW: 3.76 },
    { x: 1.43,  y: 5.096, tileX: 0.68,  color: C.sky,  tile: C.deepSky,  head: 'Heart transplant', hX: 2.175, bX: 2.187, bW: 3.749 },
    { x: 7.781, y: 3.097, tileX: 7.032, color: C.cyan, tile: C.deepCyan, head: 'Outdoor Checkup',  hX: 8.479, bX: 8.516, bW: 3.917 },
    { x: 7.781, y: 5.07,  tileX: 7.032, color: C.mint, tile: C.deepMint, head: 'Eye surgery',      hX: 8.506, bX: 8.511, bW: 3.711 },
  ].forEach(c => {
    card(s, c.x, c.y, 4.885, 1.667, c.color, 6715, { shadow: SHADOW });
    card(s, c.tileX, c.y + 0.267, 1.17, 1.06, c.tile, 8267);
    iconGlyph(s, c.tileX + 0.33, c.y + 0.54, 0.5, C.offWhite);
    lead(s, c.hX, c.y + 0.194, 2.7, 0.404, c.head, { color: C.offWhite });
    para(s, c.bX, c.y + 0.756, c.bW, 0.596, T.short, { color: C.offWhite });
  });
}

/* 11 — High Quality Service */
function slide11(s) {
  title(s, 3.156, 0.661, 7.011, 0.774, 'High Quality Service', { align: 'center' });
  kicker(s, 5.004, 1.468, { align: 'center' });
  para(s, 2.389, 1.978, 8.556, 0.596, T.three, { align: 'center' });

  [
    { x: 1.353, y: 3.62,  h: 2.99,  adj: 3841, color: C.blue, head: 'Modern Function Lab', hColor: C.white,    hX: 1.531, tX: 1.522 },
    { x: 5.041, y: 3.629, h: 2.99,  adj: 4103, color: C.sky,  head: 'Cancer Treatment',    hColor: C.offWhite, hX: 5.219, tX: 5.238 },
    { x: 8.729, y: 3.638, h: 2.981, adj: 4191, color: C.cyan, head: 'Dental Care Service', hColor: C.offWhite, hX: 8.907, tX: 8.96 },
  ].forEach(c => {
    card(s, c.x, c.y, 3.251, c.h, c.color, c.adj);
    circle(s, c.x + 1.625, c.y + 0.818, 0.907, { fill: { color: C.offWhite } });
    iconGlyph(s, c.x + 1.44, c.y + 0.633, 0.375, c.color);
    lead(s, c.hX, c.y + 1.54, 2.92, 0.404, c.head, { color: c.hColor, align: 'center' });
    para(s, c.tX, c.y + 2.032, 2.876, 0.505, T.eiusmod + ' do', { color: C.white, align: 'center' });
  });
}

/* 12 — The Features */
function slide12(s) {
  rect(s, 0, 3.362, W, 4.138, { color: C.blue, transparency: 15 });
  title(s, 2.688, 0.71, 7.917, 0.644, 'The Features', { align: 'center' });
  kicker(s, 5.004, 1.294, { align: 'center' });
  para(s, 2.378, 1.737, 8.556, 0.596, T.three, { align: 'center' });

  const copy = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt. Lorem';
  [
    { x: 1.328, y: 2.644, tile: C.blue },
    { x: 6.809, y: 2.644, tile: C.sky },
    { x: 1.328, y: 4.36,  tile: C.cyan },
    { x: 6.809, y: 4.36,  tile: C.mint },
  ].forEach(f => featureRow(s, {
    x: f.x, y: f.y, tile: f.tile, label: f.tile, text: copy,
    tileDx: 0.222, textDx: 1.415, titleDy: 0.171, textDy: 0.595, textW: 3.833, fontSize: 11,
  }));

  imageBox(s, 0, 3.397, W, 4.103, { color: C.iceBlue, transparency: 80 }, true);
}

/* 13 — Online Consulting */
function slide13(s) {
  rect(s, 8.312, 0, 5.021, 7.5, { color: C.blue });
  card(s, 6.438, 0.484, 3.748, 6.531, 'FCFCFC', 9000, { shadow: SHADOW });      // phone mock-up
  rect(s, 6.79, 1.28, 3.05, 4.82, { color: 'E4EEF4' });
  circle(s, 8.312, 6.53, 0.66, { fill: { type: 'none' }, line: { color: C.lightGrey, width: 1 } });
  rect(s, 7.98, 0.86, 0.66, 0.05, { color: C.lightGrey });
  imageBox(s, 7.13, 1.62, 2.365, 4.238, { color: C.iceBlue, transparency: 80 }, true);

  title(s, 0.852, 1.603, 5.61, 0.644, 'Online Consulting');
  hline(s, 0.996, 2.493, 1.088, C.blue, 2.75);
  para(s, 0.865, 2.906, 5.056, 1.121, T.three);
  [{ y: 4.155, x: 0.979 }, { y: 4.808, x: 0.965 }, { y: 5.433, x: 0.965 }].forEach(b => {
    circle(s, b.x + 0.058, b.y + 0.187, 0.116, { fill: { color: C.blue } });
    para(s, b.x + 0.145, b.y, 5.056, 0.555, T.sent, { fontSize: 11 });
  });

  circle(s, 11.162, 2.314, 1.707, { fill: { type: 'none' }, line: { color: C.white, width: 4, transparency: 80 } });
  s.addShape('arc', { x: 10.308, y: 1.46, w: 1.707, h: 1.707, angleRange: [146.7, 355.9], line: { color: C.white, width: 5.25 } });
  percent(s, 10.248, 1.925, 1.828, 0.777, '67', 36, C.white);
  txt(s, '34.690', { x: 10.582, y: 3.632, w: 1.5, h: 0.438, fontSize: 20, color: C.white, charSpacing: 1 });
  para(s, 10.163, 4.07, 2.821, 1.909, T.three, { color: C.offWhite });
}

/* 14 — Desktop Mockup */
function slide14(s) {
  imageBox(s, 7.959, 1.621, 4.149, 2.807, { color: C.iceBlue, transparency: 80 }, false);
  rect(s, 8.36, 1.12, 4.5, 3.1, { color: '2E3238' });                           // iMac stand-in
  rect(s, 8.52, 1.28, 4.18, 2.68, { color: 'DCE7EC' });
  rect(s, 8.36, 4.22, 4.5, 0.42, { color: 'C9CDD2' });
  rect(s, 10.05, 4.64, 1.1, 1.0, { color: 'D9DDE1' });
  rect(s, 9.4, 5.64, 2.4, 0.2, { color: 'C9CDD2' });
  imageCaption(s, 7.959, 1.621, 4.149, 2.807);

  title(s, 0.774, 1.064, 5.764, 0.644, 'Desktop Mockup');
  kicker(s, 0.674, 1.871, { align: 'center' });
  para(s, 0.774, 2.565, 5.48, 0.859,
    T.sent + 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do');

  [
    { x: 0.916, color: C.blue, tile: C.deepBlue, tileX: 1.169, iconX: 1.347, bX: 1.142 },
    { x: 4.458, color: C.cyan, tile: C.deepCyan, tileX: 4.711, iconX: 4.876, bX: 4.684 },
  ].forEach(c => {
    card(s, c.x, 4.81, 3.369, 1.785, c.color, 6715);
    card(s, c.tileX, 4.445, 0.841, 0.73, c.tile, 16667);
    iconGlyph(s, c.iconX, 4.56, 0.5, C.white);
    para(s, c.bX, 5.363, 2.737, 0.859, T.sent, { color: C.offWhite });
  });
}

/* 15 — Healthcare Solution */
function slide15(s) {
  rect(s, 7.897, 0, 5.437, 7.5, { color: C.blue });
  [
    { x: 5.994, y: 1.095, w: 2.554, h: 5.309 },
    { x: 8.896, y: 1.421, w: 2.241, h: 4.659 },
    { x: 11.415, y: 1.421, w: 2.241, h: 4.659 },
  ].forEach(p => {
    card(s, p.x, p.y, p.w, p.h, 'EFEFEF', 13869, { shadow: SHADOW });
    rect(s, p.x - 0.036, p.y + 0.75, 0.036, 0.2, { color: '6D7381' });          // side buttons
    rect(s, p.x - 0.036, p.y + 1.24, 0.036, 0.4, { color: '6D7381' });
    rect(s, p.x - 0.036, p.y + 1.73, 0.036, 0.4, { color: '6D7381' });
    rect(s, p.x + p.w, p.y + 1.24, 0.036, 0.4, { color: '6D7381' });
  });

  title(s, 0.893, 0.858, 5.74, 1.447, 'Healthcare\nSolution ');
  para(s, 0.967, 2.48, 4.199, 0.596, T.sent);
  [
    { y: 3.528, color: C.blue, head: 'Medical Tittle One', bold: true,  tX: 1.345, tY: 4.169, tW: 3.851 },
    { y: 5.218, color: C.sky,  head: 'Medical Tittle Two', bold: false, tX: 1.312, tY: 5.856, tW: 3.757 },
  ].forEach(c => {
    card(s, 1.06, c.y, 4.375, 1.552, c.color, 8348, { shadow: SHADOW });
    txt(s, c.head, { x: 1.328, y: c.y + 0.212, w: 4.375, h: 0.404, fontSize: 18, bold: c.bold, color: C.offWhite });
    para(s, c.tX, c.tY, c.tW, 0.596, T.adi2, { color: C.offWhite });
  });
}

/* 16 — Taking a Break */
function slide16(s) {
  rect(s, 0, 0, W, H, { color: C.blue, transparency: 16 });
  txt(s, 'Taking a Break', {
    x: 1.638, y: 2.719, w: 10.057, h: 1.295, fontFace: HEAD, fontSize: 88, bold: true,
    color: C.white, align: 'center', valign: 'middle', lineSpacingMultiple: 0.8,
  });
  hline(s, 5.958, 4.147, 1.417, C.white, 1);
  txt(s, 'Break Slide Time To rest Lunch Time',
    { x: 4.037, y: 4.569, w: 5.259, h: 0.438, fontSize: 20, color: C.white, align: 'center' });
}

/* 17 — Process Infographic */
function slide17(s) {
  title(s, 1.086, 0.765, 5.74, 1.447, 'Process Infographic');
  para(s, 1.118, 2.522, 4.531, 0.859,
    T.sent + 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, Lorem ipsum dolor');
  txt(s, '325+', { x: 1.086, y: 4.441, w: 1.454, h: 0.584, fontSize: 28, color: C.body, lineSpacingMultiple: 1.13 });
  lead(s, 1.086, 5.033, 3.486, 0.428, 'Medical Text Here', { color: C.body });
  para(s, 1.065, 5.484, 3.2, 0.859, T.brief);

  s.addShape('upArrow', { x: 10.356, y: 0.769, w: 1.479, h: 2.486, fill: { color: C.lightGrey } });
  poly(s, 10.839, 3.255, 0.374, 0.183, C.grey,
    [{ x: 0.1, y: 0 }, { x: 0.374, y: 0 }, { x: 0.274, y: 0.183 }, { x: 0, y: 0.183 }]);
  poly(s, 4.571, 6.108, 1.242, 0.168, C.chipBlue,
    [{ x: 0.38, y: 0 }, { x: 1.242, y: 0 }, { x: 0.862, y: 0.168 }, { x: 0, y: 0.168 }]);

  [
    { x: 4.95,  y: 4.054, num: '01', main: C.blue, light: C.tintBlue, cap: C.chipBlue },
    { x: 6.302, y: 3.339, num: '02', main: C.sky,  light: C.tintSky,  cap: C.chipSky },
    { x: 7.653, y: 2.627, num: '03', main: C.cyan, light: C.tintCyan, cap: C.chipCyan },
    { x: 9.004, y: 1.914, num: '04', main: C.mint, light: C.tintMint, cap: C.chipMint },
  ].forEach(st => processStep(s, st));

  [[5.56, 4.972], [6.915, 4.296], [8.265, 3.528], [9.601, 2.847]].forEach(([x, y]) => {
    s.addShape('arc', { x, y, w: 1.427, h: 1.427, rotate: 92.7, flipH: true,
      angleRange: [145.4, 275.5], line: { color: C.grey, width: 1.5, dashType: 'dash' } });
  });

  [
    { x: 4.94,  y: 3.573, color: C.blue },
    { x: 6.278, y: 2.801, color: C.sky },
    { x: 7.628, y: 2.066, color: C.cyan },
    { x: 8.979, y: 1.361, color: C.mint },
  ].forEach(b => {
    circle(s, b.x + 0.33, b.y + 0.33, 0.659, { fill: { color: b.color }, line: { color: C.white, width: 3 } });
    iconGlyph(s, b.x + 0.16, b.y + 0.16, 0.34, C.white);
  });

  [
    { x: 6.494,  y: 6.151, color: C.blue },
    { x: 8.017,  y: 5.39,  color: C.sky },
    { x: 9.46,   y: 4.682, color: C.cyan },
    { x: 10.661, y: 3.874, color: C.mint },
  ].forEach(t => txt(s, 'Your Text Here', { x: t.x, y: t.y, w: 2.241, h: 0.355, fontSize: 18, color: t.color }));
}

/* 18 — Our Timeline Clinic */
function slide18(s) {
  title(s, 2.441, 0.757, 8.451, 0.774, 'Our Timeline Clinic', { align: 'center' });
  hline(s, 1.264, 4.164, 10.974, C.grey, 4);

  [
    { x: 2.786, up: true,  color: C.blue, year: '2016', headX: 0.6,   headY: 1.957, textX: 0.6,   textY: 2.327 },
    { x: 5.599, up: false, color: C.sky,  year: '2017', headX: 3.332, headY: 5.147, textX: 3.361, textY: 5.513 },
    { x: 8.427, up: true,  color: C.cyan, year: '2018', headX: 6.361, headY: 1.957, textX: 6.19,  textY: 2.316 },
    { x: 11.2,  up: false, color: C.mint, year: '2019', headX: 8.586, headY: 5.147, textX: 8.952, textY: 5.46 },
  ].forEach(n => {
    const ringY = n.up ? 3.4 : 3.392;
    const stem = n.x + (n.up ? 0.775 : 0.762);
    if (n.up) {
      vline(s, stem, 2.157, 1.262, C.grey, 4);
      circle(s, stem, 2.157, 0.14, { fill: { color: C.grey } });
      s.addShape('arc', { x: n.x, y: ringY, w: 1.539, h: 1.539, angleRange: [181.4, 269.1], line: { color: C.grey, width: 4 } });
    } else {
      vline(s, stem, 4.911, 1.262, C.grey, 4);
      circle(s, stem, 6.173, 0.14, { fill: { color: C.grey } });
      s.addShape('arc', { x: n.x, y: ringY, w: 1.539, h: 1.539, flipV: true, angleRange: [181.4, 269.1], line: { color: C.grey, width: 4 } });
    }
    const cx = n.x + 0.769, cy = ringY + 0.752;
    circle(s, cx, cy, 1.504, { fill: { color: 'F8F8F8' }, shadow: SHADOW });
    circle(s, cx, cy, 1.234, { fill: { color: n.color } });
    circle(s, cx, cy, 0.998, { fill: { color: C.white } });
    txt(s, n.year, { x: cx - 0.412, y: cy - 0.202, w: 0.825, h: 0.404, fontSize: 18, color: n.color, align: 'center' });

    lead(s, n.headX, n.headY, 3.06, 0.404, 'Medical Text Here', { color: n.color, align: 'right' });
    para(s, n.textX, n.textY, 2.762, 0.596, T.half, { align: 'right' });
  });
}

/* 19 — The Agenda */
function slide19(s) {
  title(s, 3.917, 1.174, 5.5, 0.644, 'The Agenda', { align: 'center' });
  para(s, 2.347, 1.962, 8.556, 0.596, T.three, { align: 'center' });
  hline(s, 2.99, 3.371, 7.155, '808080', 1);

  [
    { x: 1.295, y: 3.605, color: C.mint,  dot: 2.99,   date: '10 July' },
    { x: 4.981, y: 3.605, color: C.green, dot: 6.568,  date: '16 July' },
    { x: 8.667, y: 3.634, color: C.olive, dot: 10.145, date: '29 July' },
  ].forEach(c => {
    circle(s, c.dot + 0.073, 3.371, 0.146, { fill: { color: c.color } });
    card(s, c.x, c.y, 3.371, 2.713, c.color, 6715);
    txt(s, c.date, { x: c.x + 0.289, y: c.y + 0.297, w: 1.65, h: 0.424, fontSize: 24, color: C.offWhite });
    lead(s, c.x + 0.285, c.y + 0.816, 3.0, 0.428, 'Medical Text Here', { color: C.offWhite });
    para(s, c.x + 0.286, c.y + 1.412, 3.0, 0.859, T.noDot, { color: C.offWhite });
  });
}

/* 20 — The Three Images */
function slide20(s) {
  title(s, 0.989, 1.215, 4.264, 1.182, 'The Three Images');
  lead(s, 4.973, 1.326, 2.2, 0.37, 'Your Text Here', { fontSize: 16, color: C.body });
  para(s, 4.98, 1.762, 6.322, 0.859, T.three);

  imageBox(s, 1.11,  3.333, 3.704, 1.592, { color: C.iceBlue, transparency: 80 }, true);
  imageBox(s, 4.809, 4.916, 3.704, 1.587, { color: C.iceBlue, transparency: 80 }, true);
  imageBox(s, 8.519, 3.333, 3.693, 1.587, { color: C.iceBlue, transparency: 80 }, true);

  [
    { x: 1.116, y: 4.906, w: 3.704, label: 'Image 1', tX: 1.33 },
    { x: 4.809, y: 3.34,  w: 3.709, label: 'Image 2', tX: 5.034 },
    { x: 8.507, y: 4.916, w: 3.716, label: 'Image 1', tX: 8.749 },
  ].forEach(c => {
    card(s, c.x, c.y, c.w, 1.587, C.blue, 5724);
    txt(s, c.label, { x: c.tX, y: c.y + 0.19, w: 1.181, h: 0.37, fontSize: 16, bold: true, color: C.white });
    para(s, c.tX, c.y + 0.645, 3.361, 0.596, T.eiusmod, { color: C.offWhite });
  });
}

/* 21 — Gallery Slide */
function slide21(s) {
  title(s, 4.203, 0.998, 4.928, 0.707, 'Gallery Slide', { align: 'center' });
  para(s, 2.134, 1.892, 8.983, 0.596, T.three, { align: 'center' });
  [-1.268, 2.155, 5.578, 9.001, 12.424].forEach(x =>
    imageBox(s, x, 3.868, 3.249, 2.023, { color: C.offWhite }, false));
}

/* 22 — We Care About Your Health (section card) */
function slide22(s) {
  card(s, 0.552, 0.69, 12.155, 6.155, C.blue, 3138);
  // oversized "?" watermark plus the three sparkles around it
  txt(s, '?', { x: 9.0, y: 2.75, w: 3.2, h: 3.0, fontFace: HEAD, fontSize: 190, bold: true, color: C.white, transparency: 88, align: 'center', valign: 'middle' });
  [[8.65, 4.1, 0.9], [9.32, 4.9, 0.64], [10.66, 4.9, 0.64]].forEach(([x, y, d]) =>
    poly(s, x, y, d, d, C.white, [
      { x: d / 2, y: 0 }, { x: d * 0.6, y: d * 0.4 }, { x: d, y: d / 2 }, { x: d * 0.6, y: d * 0.6 },
      { x: d / 2, y: d }, { x: d * 0.4, y: d * 0.6 }, { x: 0, y: d / 2 }, { x: d * 0.4, y: d * 0.4 },
    ], { fill: { color: C.white, transparency: 85 } }));

  txt(s, 'We Care\nAbout Your Health', {
    x: 1.545, y: 1.794, w: 3.553, h: 1.919, fontFace: HEAD, fontSize: 40, bold: true,
    color: C.white, valign: 'bottom', lineSpacingMultiple: 0.9,
  });
  txt(s, 'Put your subtitle here', { x: 1.562, y: 3.796, w: 4.928, h: 0.404, fontSize: 20, color: C.white });
  para(s, 1.537, 4.685, 3.361, 0.859, T.sent, { color: C.offWhite });
}

/* 23 — Table Slide Book */
function slide23(s) {
  title(s, 2.629, 0.769, 8.075, 0.774, 'Table Slide Book', { align: 'center' });
  card(s, 1.383, 1.933, 10.567, 4.658, C.blue, 3202);

  const cols = [
    { x: 1.897, w: 2.902 }, { x: 4.059, w: 2.902 }, { x: 6.381, w: 1.704 },
    { x: 8.085, w: 2.902 }, { x: 10.108, w: 1.517 },
  ];
  ['Name Of Service', 'Doctor Specialist', 'Time Schedule', 'Catergory', 'PRICE'].forEach((h, i) =>
    txt(s, h, { x: cols[i].x, y: 2.145, w: 2.7, h: 0.337, fontSize: 14, bold: true, color: C.offWhite }));

  [
    { y: 2.66,  doctor: 'Dr. Indra Alamsah',  fill: C.white, color: C.body },
    { y: 3.396, doctor: 'Dr. Hendra Wijaya',  fill: C.sky,   color: C.white },
    { y: 4.15,  doctor: 'Dr. Arnesia Reski',  fill: C.white, color: C.body },
    { y: 4.894, doctor: 'Dr. Deddy Hadid',    fill: C.white, color: C.body },
    { y: 5.639, doctor: 'Dr. Andi Pamungkas', fill: C.white, color: C.body },
  ].forEach(r => {
    card(s, 1.667, r.y, 10.0, 0.661, r.fill, 17410, { shadow: SHADOW });
    ['The creative project', r.doctor, '08.00-10.00', 'The creative project', '$7526385.00'].forEach((v, i) =>
      txt(s, v, { x: cols[i].x, y: r.y + 0.179, w: cols[i].w, h: 0.303, fontSize: 12, color: r.color }));
  });
}

/* 24 — SWOT Infographic */
function slide24(s) {
  title(s, 2.666, 0.891, 7.98, 0.774, 'SWOT Infographic', { align: 'center' });
  para(s, 1.96, 1.945, 9.33, 0.596, T.three, { align: 'center' });

  [
    { x: 0,     y: 3.055, w: 5.958, side: 'right', color: C.blue, label: 'Strengths',     lX: 0.196, tX: 0.185, tY: 3.958, text: T.two },
    { x: 7.374, y: 3.055, w: 5.959, side: 'left',  color: C.sky,  label: 'Weakness',      lX: 8.138, tX: 8.148, tY: 3.889, text: T.two },
    { x: 0,     y: 4.987, w: 5.987, side: 'right', color: C.cyan, label: 'Opportunities', lX: 0.195, tX: 0.173, tY: 5.793, text: T.two },
    { x: 7.373, y: 4.988, w: 5.96,  side: 'left',  color: C.mint, label: 'Threats',       lX: 8.167, tX: 8.138, tY: 5.751, text: T.two.replace('amet, ', 'amet,. ') },
  ].forEach(b => {
    pillPanel(s, b.x, b.y, b.w, 1.875, b.color, b.side);
    txt(s, b.label, { x: b.lX, y: b.y + 0.277, w: 4.375, h: 0.505, fontSize: 24, bold: true, color: C.white });
    para(s, b.tX, b.tY, 5.017, 0.596, b.text, { color: C.offWhite });
  });

  [
    { x: 5.391, y: 3.561, color: C.blue, letter: 'S' },
    { x: 7.006, y: 3.526, color: C.sky,  letter: 'W' },
    { x: 5.391, y: 5.521, color: C.cyan, letter: 'O' },
    { x: 7.006, y: 5.46,  color: C.mint, letter: 'T' },
  ].forEach(b => {
    circle(s, b.x + 0.466, b.y + 0.466, 0.931, { fill: { color: b.color }, line: { color: C.white, width: 3 } });
    txt(s, b.letter, { x: b.x - 0.186, y: b.y + 0.146, w: 1.303, h: 0.64, fontSize: 32, color: C.white, align: 'center' });
  });
}

/* 25 — Healthcare Chart Infographic */
function slide25(s, pptx) {
  rect(s, 0, -0.012, 7.317, 4.785, { color: C.paleBlue });

  const cats = ['Data 1', 'Data 2', 'Data 3 ', 'Data 4'];
  s.addChart(pptx.ChartType.bar, [
    { name: 'Series 1', labels: cats, values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels: cats, values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels: cats, values: [2.0, 2.0, 3.0, 5.0] },
  ], {
    x: 0.934, y: 1.247, w: 5.93, h: 3.012,
    barDir: 'col', barGapWidthPct: 219, barOverlapPct: -27,
    chartColors: [C.blue, C.cyan, C.lightGrey],
    showLegend: true, legendPos: 't', legendFontSize: 16, legendColor: '7F7F7F', legendFontFace: BODY,
    catAxisLabelFontSize: 16, catAxisLabelColor: C.body, catAxisLabelFontFace: BODY,
    catAxisLineColor: C.lightGrey, valAxisHidden: true, valGridLine: { style: 'none' }, showValue: false,
  });

  s.addShape('wedgeRectCallout', { x: 0.694, y: 0.739, w: 2.084, h: 0.666, fill: { color: C.blue } });
  txt(s, 'Value Average', { x: 0.82, y: 0.884, w: 1.066, h: 0.43, fontSize: 12, color: C.offWhite, lineSpacingMultiple: 0.8 });
  txt(s, '81%', { x: 1.504, y: 0.768, w: 1.39, h: 0.536, fontSize: 24, color: C.offWhite, align: 'center', valign: 'middle' });

  title(s, 7.799, 0.785, 5.328, 2.121, 'Healthcare \nChart Infographic');
  para(s, 7.799, 2.932, 5.075, 1.121, T.four);

  card(s, 1.097, 4.495, 5.443, 2.044, C.blue, 5000, { shadow: SHADOW });
  txt(s, '135+', { x: 1.342, y: 4.664, w: 1.454, h: 0.584, fontSize: 28, color: C.white, lineSpacingMultiple: 1.13 });
  lead(s, 1.342, 5.219, 2.718, 0.428, 'Medical Text Here', { color: C.white });
  para(s, 1.342, 5.631, 4.789, 0.596, T.two, { color: C.offWhite });

  card(s, 6.715, 4.495, 5.443, 2.044, C.blue, 5000, { shadow: SHADOW });
  circle(s, 7.756, 5.505, 1.409, { fill: { type: 'none' }, line: { color: '000000', width: 6.25, transparency: 90 } });
  s.addShape('arc', { x: 7.051, y: 4.8, w: 1.409, h: 1.409, angleRange: [270, 181.4], line: { color: C.cyan, width: 8.25 } });
  percent(s, 7.051, 4.8, 1.409, 1.409, '75', 28, C.offWhite);
  lead(s, 8.621, 4.808, 3.4, 0.404, 'Medical Tittle Text Here', { color: C.offWhite });
  para(s, 8.602, 5.32, 3.287, 0.859, T.two, { color: C.offWhite });
}

/* 26 — Chart Infographic (doughnuts + pictogram bars) */
function slide26(s, pptx) {
  rect(s, 6.012, 0, 7.321, 7.5, { color: C.blue });
  card(s, 6.012, 1.496, 6.504, 3.909, C.offWhite, 2263, { fill: { color: C.offWhite, transparency: 10 } });
  card(s, 6.012, 4.98, 6.504, 1.803, C.offWhite, 1406);

  [
    { x: 6.366, values: [70, 30], color: C.blue, label: '$ 2.500.000',        labelX: 6.586, iconX: 7.578 },
    { x: 9.264, values: [70, 60], color: C.sky,  label: 'Medical-Healthcare', labelX: 9.484, iconX: 10.443 },
  ].forEach(d => {
    s.addChart(pptx.ChartType.doughnut,
      [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: d.values }], {
        x: d.x, y: 1.456, w: 2.898, h: 3.129,
        holeSize: 50, firstSliceAng: 0, chartColors: [d.color, C.grey],
        showLegend: false, showPercent: true,
        dataLabelColor: C.body, dataLabelFontSize: 16, dataLabelFontFace: BODY,
      });
    circle(s, d.iconX + 0.24, 3.021, 0.62, { fill: { color: C.white } });
    iconGlyph(s, d.iconX + 0.06, 2.84, 0.36, d.color);
    txt(s, d.label, { x: d.labelX, y: 4.34, w: 2.417, h: 0.37, fontSize: 16, color: C.ink, align: 'center' });
  });

  [
    { y: 5.26,  color: C.blue, grey: 2, pct: '76%', pctColor: C.blue },
    { y: 5.985, color: C.mint, grey: 1, pct: '84%', pctColor: C.sky },
  ].forEach(row => {
    for (let i = 0; i < 12; i++) {
      personIcon(s, 7.9 + i * 0.264, row.y, 0.185, 0.55, i < 12 - row.grey ? row.color : C.midGrey);
    }
    txt(s, row.pct, { x: 11.146, y: row.y + 0.05, w: 0.892, h: 0.446, fontSize: 20, bold: true, color: row.pctColor });
  });

  title(s, 0.917, 1.123, 4.964, 1.447, 'Chart Infographic');
  para(s, 0.959, 2.867, 4.567, 1.646, T.five);
  card(s, 1.048, 5.313, 6.222, 1.398, C.sky, 6125);
  iconGlyph(s, 1.34, 5.68, 0.62, C.white);
  lead(s, 2.135, 5.479, 2.6, 0.412, 'Medical Text Here', { color: C.white });
  para(s, 2.143, 5.914, 4.949, 0.596,
    T.sent + 'Lorem ipsum dolor sit', { color: C.offWhite });
}

/* 27 — Pricing Table */
function slide27(s) {
  rect(s, 0, 0, W, 4.531, { color: C.blue });
  card(s, 1.47, 2.419, 10.312, 3.922, C.offWhite, 6715);
  title(s, 3.917, 1.236, 5.5, 0.644, 'Pricing Table', { align: 'center', color: C.white });
  vline(s, 4.948, 3.013, 2.75, C.lightGrey, 1);
  vline(s, 8.385, 3.013, 2.75, C.lightGrey, 1);

  [
    { x: 2.016, price: '299,00 $', tier: 'SMALL',  tierSize: 16, color: C.blue, btnX: 2.693, tX: 1.788, text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt. Lorem ipsum dolor sit amet' },
    { x: 5.453, price: '499,00 $', tier: 'MEDIUM', tierSize: 18, color: C.sky,  btnX: 6.13,  tX: 5.185, text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt. Lorem ipsum dolor sit amet' },
    { x: 8.891, price: '699,00 $', tier: 'LARGE',  tierSize: 18, color: C.cyan, btnX: 9.568, tX: 8.574, text: T.two },
  ].forEach(c => {
    txt(s, c.price, { x: c.x + 0.165, y: 2.956, w: 2.097, h: 0.55, fontSize: 32, color: C.body, align: 'center', valign: 'middle' });
    txt(s, c.tier, { x: c.x, y: 3.609, w: 2.427, h: 0.427, fontSize: c.tierSize, bold: true, charSpacing: 1, color: C.body, align: 'center', valign: 'bottom' });
    para(s, c.tX, 4.105, 2.97, 1.121, c.text, { align: 'center' });
    s.addShape('round2DiagRect', { x: c.btnX, y: 5.486, w: 1.073, h: 0.333, fill: { color: c.color } });
    txt(s, 'Choose', { x: c.btnX, y: 5.486, w: 1.073, h: 0.333, fontSize: 12, color: C.white, align: 'center', valign: 'middle' });
  });
}

/* 28 — Circle Infographic */
function slide28(s) {
  title(s, 3.917, 1.171, 5.5, 0.644, 'Circle Infographic', { align: 'center' });

  const ring = { x: 4.784, y: 2.436, d: 3.77 };
  [                                                   // angles measured from 12 o'clock
    { from: 331, to: 29,  color: C.blue },
    { from: 31,  to: 89,  color: C.olive },
    { from: 91,  to: 149, color: C.green },
    { from: 151, to: 209, color: C.mint },
    { from: 211, to: 269, color: C.cyan },
    { from: 271, to: 329, color: C.sky },
  ].forEach(g => {
    s.addShape('blockArc', {
      x: ring.x, y: ring.y, w: ring.d, h: ring.d, fill: { color: g.color },
      angleRange: [compass(g.from), compass(g.to)], arcThicknessRatio: 0.52,
    });
    const span = g.to < g.from ? g.to + 360 : g.to;
    const mid = ((g.from + span) / 2) * Math.PI / 180;
    const r = ring.d * 0.365;
    iconGlyph(s, ring.x + ring.d / 2 + Math.sin(mid) * r - 0.19,
                 ring.y + ring.d / 2 - Math.cos(mid) * r - 0.19, 0.38, C.white);
  });

  [
    { align: 'right', y: 2.914, x: 0.794, tick: 4.423, tickY: 3.012, color: C.blue },
    { align: 'right', y: 4.052, x: 0.573, tick: 4.107, tickY: 4.162, color: C.sky },
    { align: 'right', y: 5.079, x: 0.658, tick: 4.26,  tickY: 5.242, color: C.cyan },
    { align: 'left',  y: 2.891, x: 9.0,   tick: 8.705, tickY: 3.064, color: C.olive },
    { align: 'left',  y: 4.029, x: 9.264, tick: 9.027, tickY: 4.213, color: C.green },
    { align: 'left',  y: 5.13,  x: 9.009, tick: 8.756, tickY: 5.361, color: C.mint },
  ].forEach(r => {
    checkMark(s, r.tick, r.tickY, 0.216, 0.167, r.color);
    para(s, r.x, r.y, 3.538, 0.859, T.sent, { align: r.align });
  });
}

/* 29 — United states Map Infographic */
function slide29(s) {
  const m = { x: 6.121, y: 2.034, w: 6.212, h: 4.145 };     // simplified mainland silhouette
  poly(s, m.x, m.y, m.w, m.h, C.lightGrey, [
    [0.02, 0.24], [0.07, 0.14], [0.19, 0.10], [0.34, 0.06], [0.52, 0.03], [0.62, 0.00],
    [0.70, 0.05], [0.83, 0.06], [0.92, 0.13], [0.97, 0.24], [1.00, 0.36], [0.95, 0.44],
    [0.93, 0.55], [0.90, 0.66], [0.84, 0.62], [0.76, 0.58], [0.70, 0.66], [0.63, 0.60],
    [0.55, 0.74], [0.49, 0.86], [0.44, 0.72], [0.36, 0.58], [0.24, 0.50], [0.12, 0.42],
    [0.05, 0.34],
  ].map(([px, py]) => ({ x: px * m.w, y: py * m.h })));
  poly(s, m.x - 0.05, m.y + 2.55, 1.45, 0.9, C.lightGrey, [   // Alaska
    { x: 0, y: 0.5 }, { x: 0.32, y: 0.16 }, { x: 0.78, y: 0.28 }, { x: 1.15, y: 0.02 },
    { x: 1.45, y: 0.32 }, { x: 1.05, y: 0.68 }, { x: 0.5, y: 0.9 },
  ]);
  poly(s, m.x + 1.9, m.y + 3.5, 0.75, 0.35, C.lightGrey, [    // Hawaii
    { x: 0, y: 0.1 }, { x: 0.3, y: 0 }, { x: 0.55, y: 0.2 }, { x: 0.75, y: 0.35 }, { x: 0.35, y: 0.3 },
  ]);
  poly(s, 8.252, 3.946, 1.549, 1.472, C.blue, [               // Texas
    { x: 0.15, y: 0 }, { x: 1.35, y: 0.05 }, { x: 1.549, y: 0.5 }, { x: 1.1, y: 0.72 },
    { x: 0.9, y: 1.472 }, { x: 0.42, y: 1.05 }, { x: 0, y: 0.62 },
  ]);
  rect(s, 8.033, 3.296, 0.786, 0.585, { color: C.sky });      // Colorado
  poly(s, 9.941, 3.063, 0.448, 0.764, C.cyan, [               // Illinois
    { x: 0.05, y: 0 }, { x: 0.4, y: 0.08 }, { x: 0.448, y: 0.5 }, { x: 0.26, y: 0.764 }, { x: 0, y: 0.4 },
  ]);

  txt(s, 'United states\nMap Infographic', { x: 1.243, y: 1.358, w: 5.708, h: 1.313, fontFace: HEAD, fontSize: 40, bold: true, color: C.ink, lineSpacingMultiple: 0.9 });
  txt(s, 'Infographic Map Section', { x: 1.243, y: 2.735, w: 4.928, h: 0.374, fontSize: 18, color: C.body });
  [
    { n: '01', y: 3.596, tY: 3.588, color: C.blue },
    { n: '02', y: 4.498, tY: 4.497, color: C.sky },
    { n: '03', y: 5.499, tY: 5.462, color: C.cyan },
  ].forEach(r => {
    txt(s, r.n, { x: 1.098, y: r.y, w: 0.913, h: 0.505, fontSize: 24, bold: true, color: r.color, align: 'center', valign: 'middle' });
    para(s, 1.844, r.tY, 4.035, 0.596, T.sent);
  });
}

/* 30 — Denmark Map Infographic */
function slide30(s) {
  poly(s, 7.0, 1.53, 2.2, 2.1, C.lightGrey, [                 // Jutland
    { x: 0.9, y: 0 }, { x: 1.5, y: 0.3 }, { x: 1.6, y: 1.0 }, { x: 2.2, y: 1.25 },
    { x: 1.8, y: 1.75 }, { x: 1.0, y: 2.1 }, { x: 0.3, y: 1.6 }, { x: 0, y: 0.8 },
  ]);
  poly(s, 9.35, 2.9, 1.5, 1.35, C.lightGrey, [                // Zealand
    { x: 0.5, y: 0 }, { x: 1.5, y: 0.35 }, { x: 1.25, y: 1.1 }, { x: 0.4, y: 1.35 }, { x: 0, y: 0.6 },
  ]);
  poly(s, 8.6, 4.2, 1.1, 1.0, C.lightGrey, [{ x: 0.4, y: 0 }, { x: 1.1, y: 0.4 }, { x: 0.7, y: 1.0 }, { x: 0, y: 0.55 }]);
  rect(s, 11.55, 3.55, 0.55, 0.45, { color: C.lightGrey });
  poly(s, 6.563, 2.826, 2.767, 1.644, C.blue, [               // highlighted region
    { x: 0.6, y: 0 }, { x: 2.2, y: 0.15 }, { x: 2.767, y: 0.55 }, { x: 2.3, y: 1.15 },
    { x: 1.4, y: 1.644 }, { x: 0.5, y: 1.2 }, { x: 0, y: 0.5 },
  ]);

  txt(s, 'Denmark\nMap Infographic', { x: 1.14, y: 1.937, w: 5.403, h: 1.313, fontFace: HEAD, fontSize: 40, bold: true, color: C.ink, lineSpacingMultiple: 0.9 });
  para(s, 1.186, 3.572, 5.091, 1.121, T.labore + T.sent);
  card(s, 1.29, 5.191, 1.889, 0.414, C.blue, 50000);
  txt(s, 'Explore More', { x: 1.29, y: 5.191, w: 1.889, h: 0.414, fontSize: 11, bold: true, color: C.white, align: 'center', valign: 'middle' });

  rect(s, 8.78, 2.129, 2.858, 1.539, { color: C.blue });
  poly(s, 8.45, 2.55, 0.34, 0.55, C.blue, [{ x: 0.34, y: 0 }, { x: 0.34, y: 0.55 }, { x: 0, y: 0.24 }]);
  percent(s, 9.856, 2.317, 1.242, 0.572, '80', 28, C.white);
  txt(s, 'A wonderful serenity has taken possession',
    { x: 9.597, y: 2.895, w: 1.767, h: 0.516, fontSize: 10, color: C.offWhite, align: 'center', lineSpacingMultiple: 1.3 });
}

/* 31 — Contact Us */
function slide31(s) {
  rect(s, 0, 0, W, H, { color: C.blue });
  imageBox(s, 0, 0, W, H, { color: '0098C8', transparency: 90 }, true);

  txt(s, 'Contact Us', { x: 1.446, y: 1.346, w: 4.11, h: 0.752, fontFace: HEAD, fontSize: 48, bold: true, color: C.white });
  lead(s, 1.446, 2.346, 1.6, 0.438, 'Address', { fontSize: 20, color: C.offWhite });
  para(s, 1.446, 3.091, 3.969, 0.859, T.labore.replace('dolore. ', 'dolore. ') + T.half, { color: C.offWhite });

  txt(s, 'Phone', { x: 1.446, y: 4.838, w: 1.3, h: 0.438, fontSize: 20, color: C.offWhite });
  [
    { y: 5.32,  text: '081333xxxxxx / 082233xxxxx' },
    { y: 5.827, text: 'Lorem ipsum dolor sit amet' },
  ].forEach(r => {
    iconGlyph(s, 1.567, r.y - 0.01, 0.27, C.offWhite);
    txt(s, r.text, { x: 1.96, y: r.y, w: 4.269, h: 0.319, fontSize: 12, charSpacing: 1, color: C.white });
  });

  card(s, 9.002, 4.317, 3.218, 1.174, C.blue, 5702);
  s.addShape('triangle', { x: 9.431, y: 4.158, w: 0.307, h: 0.164, fill: { color: C.blue } });
  txt(s, 'Address', { x: 9.222, y: 4.47, w: 1.297, h: 0.286, fontSize: 11, charSpacing: 1, color: C.offWhite, valign: 'bottom' });
  para(s, 9.193, 4.751, 2.885, 0.555, T.half, { fontSize: 10.5, color: C.offWhite });
}

/* 32 — Thank You */
function slide32(s) {
  s.background = { color: C.blue };
  s.addText(
    [{ text: 'Thank You ', options: { fontSize: 80, bold: true } },
     { text: '\n', options: { fontSize: 80 } },
     { text: 'For The Attention', options: { fontSize: 32 } }],
    { x: 5.994, y: 2.634, w: 7.72, h: 1.666, fontFace: HEAD, color: C.white, valign: 'middle', lineSpacingMultiple: 0.8 });

  txt(s, 'A wonderful serenity has taken possession',
    { x: 5.994, y: 4.419, w: 5.116, h: 0.372, fontSize: 14, bold: true, color: C.white });
  para(s, 5.994, 4.938, 4.494, 0.596, T.sent, { color: C.offWhite });

  card(s, 1.39, 6.736, 1.099, 0.284, C.cyan, 50000);
  ['f', 't', 'g'].forEach((g, i) =>
    txt(s, g, { x: 1.47 + i * 0.29, y: 6.736, w: 0.28, h: 0.284, fontSize: 11, bold: true, color: C.white, align: 'center', valign: 'middle' }));
}

/* ================================================================== *
 * 4. Build
 * ================================================================== */
const SLIDES = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
  slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
  slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24,
  slide25, slide26, slide27, slide28, slide29, slide30, slide31, slide32,
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'CUSTOM_16x9', width: W, height: H });
  pptx.layout = 'CUSTOM_16x9';
  pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
  pptx.title = 'Nursing – Medical Presentation Template';

  SLIDES.forEach(fn => fn(pptx.addSlide(), pptx));

  return pptx.writeFile({
    fileName: path.join(__dirname, '129f2e8b-e9f8-446f-bb63-e24e03b634f7_grok_final.pptx'),
  });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
