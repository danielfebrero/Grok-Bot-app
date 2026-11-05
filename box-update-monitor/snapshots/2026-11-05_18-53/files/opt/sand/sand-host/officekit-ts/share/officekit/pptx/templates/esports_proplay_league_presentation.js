/**
 * "E-SPORT CHAMPIONS" — 20-slide gaming deck rebuilt with pptxgenjs.
 * Run: node 0653328f-bd8f-438e-9fb7-b114bf96c71d_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Design tokens
 * ------------------------------------------------------------------ */

const SLIDE_W = 13.3333;
const SLIDE_H = 7.5;

const C = {
  ink: '200B2C',        // theme accent1
  darkBg: '311E3C',     // accent1 at 92% over white — the dark slides' panel
  purple: '7030A0',     // theme accent2 — brand purple
  purpleLight: 'A049D2',
  glassEdge: 'A094A7',  // hairline around the translucent bars on dark slides
  white: 'FFFFFF',
  heading: '404040',    // tx1 @ 75% luminance
  body: '404040',
  muted: '808080',      // tx2 @ 50% luminance
};

// Dark slides cover the master chrome with a 92%-opaque plum panel, leaving
// the header/footer barely legible. These are the resulting blended tones.
const DIM = { text: '230F2D', pill: '260D35', pillText: '311E3C' };

const HEAD = 'Roboto Condensed';  // theme major font
const BODY = 'Karla';             // theme minor font

// Diagonal-corner-snip fraction used by every "snip2DiagRect" in the deck.
const SNIP = { card: 0.20174, pill: 0.29334, wide: 0.43738, tile: 0.25946, team: 0.23495, half: 0.5 };

// Soft drop shadow under the purple/white pill buttons.
const PILL_SHADOW = { type: 'outer', blur: 40, offset: 0, angle: 0, color: '000000', opacity: 0.07 };
// The photo cards sit on a hard purple silhouette offset down-left (16pt @ 134°).
const CARD_OFFSET = { x: -0.154, y: 0.160 };

/* ------------------------------------------------------------------ *
 * Geometry helpers
 * ------------------------------------------------------------------ */

// snip2DiagRect: top-right and bottom-left corners cut off.
function snipPoints(w, h, frac) {
  const c = Math.min(w, h) * frac;
  return [{ x: 0, y: 0 }, { x: w - c, y: 0 }, { x: w, y: c },
          { x: w, y: h }, { x: c, y: h }, { x: 0, y: h - c }, { close: true }];
}

function snip(slide, o) {
  slide.addShape('custGeom', Object.assign({
    points: snipPoints(o.w, o.h, o.frac === undefined ? SNIP.pill : o.frac),
  }, o));
}

// Icon outlines, stored as unit-square paths ([m|l]=move/line, c=cubic, z=close).
const ICONS = {
  console: [['m',0.583,0.35],['l',0.208,0.35],['l',0.208,0.25],['l',0.583,0.25],['z'],['m',1,0.15],['l',1,0.25],['l',0.958,0.25],['l',0.958,0.6],['l',0.792,0.6],['l',0.792,1],['l',0,1],['l',0,0.15],['c',0,0.067,0.056,0,0.125,0],['l',0.667,0],['c',0.736,0,0.792,0.067,0.792,0.15],['l',0.792,0.5],['l',0.875,0.5],['l',0.875,0.15],['z'],['m',0.333,0.65],['l',0.25,0.65],['l',0.25,0.55],['l',0.167,0.55],['l',0.167,0.65],['l',0.083,0.65],['l',0.083,0.75],['l',0.167,0.75],['l',0.167,0.85],['l',0.25,0.85],['l',0.25,0.75],['l',0.333,0.75],['z'],['m',0.708,0.7],['c',0.708,0.659,0.68,0.625,0.646,0.625],['c',0.611,0.625,0.583,0.659,0.583,0.7],['c',0.583,0.741,0.611,0.775,0.646,0.775],['c',0.68,0.775,0.708,0.741,0.708,0.7],['z'],['m',0.542,0.7],['c',0.542,0.659,0.514,0.625,0.479,0.625],['c',0.445,0.625,0.417,0.659,0.417,0.7],['c',0.417,0.741,0.445,0.775,0.479,0.775],['c',0.514,0.775,0.542,0.741,0.542,0.7],['z'],['m',0.667,0.15],['l',0.125,0.15],['l',0.125,0.45],['l',0.667,0.45],['z']],
  monitorTop: [['m',1,1],['l',1,0.25],['c',1,0.112,0.944,0,0.875,0],['l',0.125,0],['c',0.056,0,0,0.112,0,0.25],['l',0,1],['z']],
  monitorBase: [['m',1,0.5],['l',1,0],['l',0,0],['l',0,0.5],['l',0.458,0.5],['l',0.458,0.75],['l',0.25,0.75],['l',0.25,1],['l',0.75,1],['l',0.75,0.75],['l',0.542,0.75],['l',0.542,0.5],['l',1,0.5],['z']],
  gamepad: [['m',0.888,0.152],['c',0.869,0.091,0.82,0.05,0.766,0.05],['l',0.75,0.05],['l',0.75,0],['l',0.667,0],['l',0.667,0.05],['l',0.333,0.05],['l',0.333,0],['l',0.25,0],['l',0.25,0.05],['l',0.234,0.05],['c',0.18,0.05,0.131,0.091,0.111,0.152],['c',0.019,0.447,0,0.773,0,0.844],['c',0,0.93,0.058,1,0.13,1],['c',0.242,1,0.349,0.766,0.385,0.65],['l',0.616,0.65],['c',0.651,0.766,0.758,1,0.87,1],['c',0.942,1,1,0.93,1,0.844],['c',1,0.773,0.981,0.447,0.889,0.152],['z'],['m',0.312,0.45],['c',0.278,0.45,0.25,0.416,0.25,0.375],['c',0.25,0.334,0.278,0.3,0.312,0.3],['c',0.347,0.3,0.375,0.334,0.375,0.375],['c',0.375,0.416,0.347,0.45,0.312,0.45],['z'],['m',0.688,0.45],['c',0.653,0.45,0.625,0.416,0.625,0.375],['c',0.625,0.334,0.653,0.3,0.688,0.3],['c',0.722,0.3,0.75,0.334,0.75,0.375],['c',0.75,0.416,0.722,0.45,0.688,0.45],['z']],
  headset: [['m',0.875,0.521],['l',0.875,0.45],['c',0.875,0.202,0.707,0,0.5,0],['c',0.293,0,0.125,0.202,0.125,0.45],['l',0.125,0.521],['c',0.02,0.576,-0.029,0.724,0.018,0.85],['c',0.051,0.941,0.126,1,0.208,1],['l',0.292,1],['l',0.292,0.5],['l',0.208,0.5],['l',0.208,0.45],['c',0.208,0.257,0.339,0.1,0.5,0.1],['c',0.661,0.1,0.792,0.257,0.792,0.45],['l',0.792,0.5],['l',0.708,0.5],['l',0.708,1],['l',0.792,1],['c',0.907,1,1,0.888,1,0.75],['c',1,0.651,0.951,0.561,0.875,0.521],['z']],
  chess: [['m',0.455,0.458],['l',0.045,0.458],['l',0.045,0.39],['c',0.046,0.333,0.085,0.282,0.144,0.263],['c',0.278,0.218,0.369,0.164,0.405,0.032],['l',0.413,0],['l',0.449,0],['c',0.949,0,0.955,0.412,0.955,0.417],['l',0.955,0.833],['l',0.193,0.833],['c',0.29,0.752,0.418,0.619,0.455,0.458],['z'],['m',0,0.917],['l',0,1],['l',1,1],['l',1,0.917],['z']],
  phone: [['m',0,0.301],['c',0,0.242,0.024,0.186,0.066,0.143],['c',0.162,0.048,0.331,0,0.5,0],['c',0.669,0,0.838,0.048,0.934,0.143],['c',0.977,0.186,1,0.242,1,0.301],['c',1,0.358,0.969,0.407,0.923,0.434],['l',0.816,0.281],['c',0.791,0.247,0.758,0.22,0.718,0.205],['c',0.667,0.186,0.59,0.166,0.5,0.166],['c',0.41,0.166,0.333,0.186,0.282,0.205],['c',0.242,0.22,0.208,0.247,0.184,0.281],['l',0.076,0.434],['c',0.031,0.407,0,0.358,0,0.301],['z'],['m',0.625,0.625],['c',0.625,0.694,0.569,0.75,0.5,0.75],['c',0.431,0.75,0.375,0.694,0.375,0.625],['c',0.375,0.555,0.431,0.499,0.5,0.499],['c',0.569,0.499,0.625,0.555,0.625,0.625],['z'],['m',1,0.841],['c',0.998,0.93,0.923,1,0.834,1],['l',0.166,1],['c',0.079,1,0.002,0.929,0,0.841],['c',-0.001,0.771,0.014,0.701,0.044,0.637],['c',0.05,0.624,0.058,0.611,0.067,0.598],['l',0.238,0.344],['c',0.259,0.314,0.288,0.291,0.322,0.279],['c',0.365,0.265,0.428,0.249,0.5,0.249],['c',0.572,0.249,0.634,0.265,0.678,0.28],['c',0.712,0.291,0.742,0.314,0.762,0.344],['l',0.946,0.617],['c',0.983,0.686,1.001,0.763,1,0.841],['z'],['m',0.708,0.625],['c',0.708,0.509,0.615,0.416,0.5,0.416],['c',0.385,0.416,0.292,0.509,0.292,0.625],['c',0.292,0.74,0.385,0.833,0.5,0.833],['c',0.615,0.833,0.708,0.74,0.708,0.625],['z']],
  pin: [['m',0.5,0],['c',0.224,0,0,0.186,0,0.416],['c',0,0.523,0.1,0.691,0.296,0.914],['c',0.378,1.008,0.536,1.028,0.649,0.96],['c',0.67,0.947,0.688,0.931,0.704,0.914],['c',0.9,0.691,1,0.523,1,0.416],['c',1,0.186,0.776,0,0.5,0],['z'],['m',0.5,0.582],['c',0.389,0.582,0.3,0.507,0.3,0.415],['c',0.3,0.323,0.389,0.248,0.5,0.248],['c',0.611,0.248,0.7,0.323,0.7,0.415],['c',0.7,0.507,0.611,0.582,0.5,0.582],['z']],
  mailBody: [['m',0.998,0],['l',0.647,0.482],['c',0.566,0.594,0.434,0.594,0.353,0.482],['l',0.002,0],['c',0.001,0.009,0,0.017,0,0.026],['l',0,0.714],['c',0,0.872,0.093,1,0.208,1],['l',0.792,1],['c',0.907,1,1,0.872,1,0.714],['l',1,0.026],['c',1,0.017,0.999,0.009,0.998,0],['z']],
  mailFlap: [['m',0.594,0.929],['l',1,0.194],['c',0.96,0.074,0.888,0,0.811,0],['l',0.189,0],['c',0.112,0,0.04,0.074,0,0.194],['l',0.406,0.929],['c',0.458,1.024,0.542,1.024,0.594,0.929],['z']],
  // Outlined display letters used on the title / thank-you slides.
  letterE: [['m',0,0],['l',0.242,0],['l',0.351,0],['l',0.998,0],['l',0.998,0.167],['l',0.351,0.167],['l',0.351,0.405],['l',0.898,0.405],['l',0.898,0.567],['l',0.351,0.567],['l',0.351,0.834],['l',1,0.834],['l',1,1],['l',0.351,1],['l',0.242,1],['l',0,1],['z']],
  letterT: [['m',0,0],['l',0.352,0],['l',0.641,0],['l',1,0],['l',1,0.167],['l',0.641,0.167],['l',0.641,1],['l',0.352,1],['l',0.352,0.167],['l',0,0.167],['z']],
};

function iconPoints(name, w, h) {
  return ICONS[name].map(function (p) {
    if (p[0] === 'z') return { close: true };
    if (p[0] === 'm') return { x: p[1] * w, y: p[2] * h, moveTo: true };
    if (p[0] === 'l') return { x: p[1] * w, y: p[2] * h };
    return { x: p[5] * w, y: p[6] * h,
             curve: { type: 'cubic', x1: p[1] * w, y1: p[2] * h, x2: p[3] * w, y2: p[4] * h } };
  });
}

function icon(slide, name, x, y, w, h, color) {
  slide.addShape('custGeom', { x: x, y: y, w: w, h: h, points: iconPoints(name, w, h), fill: { color: color } });
}

/* ------------------------------------------------------------------ *
 * Text helpers — every text box in the deck is top-anchored, no bullet
 * ------------------------------------------------------------------ */

function txt(slide, text, o) {
  slide.addText(text, Object.assign({
    fontFace: BODY, fontSize: 12, color: C.body, valign: 'top',
    margin: [7.2, 7.2, 3.6, 3.6], // PowerPoint default 0.1" / 0.05" insets
  }, o));
}

// Two- or three-tone headline: array of [text, color] pairs, rendered as one paragraph.
function heading(slide, runs, o) {
  txt(slide, runs.map(function (r) { return { text: r[0], options: { color: r[1] } }; }),
      Object.assign({ fontFace: HEAD, fontSize: 48, bold: true, lineSpacingMultiple: 0.99 }, o));
}

// Body paragraph — 12pt Karla at 1.5 line spacing.
function para(slide, text, o) {
  txt(slide, text, Object.assign({ lineSpacingMultiple: 1.5 }, o));
}

// White plate + purple pill + centred label ("COMPETITION", "LEGENDS NEVER STOP PLAYING", ...).
function button(slide, o) {
  snip(slide, { x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: C.white }, shadow: PILL_SHADOW });
  snip(slide, { x: o.x + 0.071, y: o.y + 0.078, w: o.w - 0.143, h: o.h - 0.156,
                fill: { color: C.purple }, shadow: PILL_SHADOW });
  txt(slide, o.text, { x: o.x + 0.228, y: o.y + 0.191, w: o.w - 0.456, h: 0.4376,
                       fontFace: HEAD, fontSize: o.fontSize || 20, bold: true,
                       color: C.white, align: 'center' });
}

// Bare purple pill with a centred label ("READ MORE", "JOIN US").
function pill(slide, o) {
  snip(slide, { x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: C.purple }, shadow: PILL_SHADOW });
  txt(slide, o.text, { x: o.labelX, y: o.y + 0.086, w: 2.2775, h: 0.4039, align: 'center',
                       fontFace: HEAD, fontSize: 18, bold: true, color: C.white });
}

// Small square badge holding a white gamepad glyph.
function gamepadBadge(slide, x, y) {
  snip(slide, { x: x, y: y, w: 1.1084, h: 0.8194, fill: { color: C.white }, shadow: PILL_SHADOW });
  snip(slide, { x: x + 0.0865, y: y + 0.0778, w: 0.9353, h: 0.6638, fill: { color: C.purple }, shadow: PILL_SHADOW });
  icon(slide, 'gamepad', x + 0.3758, y + 0.2611, 0.3568, 0.2973, C.white);
}

// "2.500+ / Global Championship" stat pair.
function stat(slide, x, y, value, label, o) {
  o = o || {};
  txt(slide, value, { x: x, y: y, w: 2.0139, h: o.valueH || 0.4376,
                      fontFace: o.valueFace || BODY, fontSize: o.valueSize || 20,
                      bold: true, color: o.valueColor || C.body, align: o.align });
  txt(slide, label, { x: x, y: y + (o.gap || 0.443), w: 2.3333, h: 0.3029,
                      fontSize: o.labelSize || 11, color: o.labelColor || C.muted, align: o.align });
}

// 85%-style progress bar: dark track, purple fill, purple end caps.
function progressBar(slide, x, y) {
  slide.addShape('rect', { x: x + 0.008, y: y + 0.034, w: 4.704, h: 0.1508, fill: { color: C.ink }, shadow: PILL_SHADOW });
  slide.addShape('rect', { x: x + 0.008, y: y + 0.034, w: 3.9133, h: 0.1508, fill: { color: C.purple } });
  slide.addShape('rect', { x: x, y: y, w: 0.0527, h: 0.2179, fill: { color: C.purple } });
  slide.addShape('rect', { x: x + 3.910, y: y + 0.017, w: 0.0516, h: 0.1832, fill: { color: C.purple } });
}

/* ------------------------------------------------------------------ *
 * Image placeholders — every photo in the source deck becomes a
 * flat rectangle in its original footprint, tagged "[image]".
 * ------------------------------------------------------------------ */

const PLACEHOLDER_FILL = 'D9D3DE';
const PLACEHOLDER_TEXT = '6E5F78';

function imagePlaceholder(slide, o) {
  const frac = o.frac === undefined ? SNIP.card : o.frac;
  if (o.cardShadow !== false) {
    slide.addShape('custGeom', {
      x: o.x + CARD_OFFSET.x, y: o.y + CARD_OFFSET.y, w: o.w, h: o.h,
      points: snipPoints(o.w, o.h, frac), fill: { color: C.purple },
    });
  }
  slide.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h, points: snipPoints(o.w, o.h, frac),
    fill: { color: o.fill || PLACEHOLDER_FILL },
  });
  txt(slide, '[image]', { x: o.x, y: o.y + o.h / 2 - 0.2, w: o.w, h: 0.4,
                          align: 'center', fontSize: 11, color: o.labelColor || PLACEHOLDER_TEXT });
}

/* ------------------------------------------------------------------ *
 * Chrome shared by every slide (from the slide master)
 * ------------------------------------------------------------------ */

function chrome(slide, pageNo, opts) {
  opts = opts || {};
  const label = opts.onDark ? DIM.text : C.heading;
  txt(slide, 'PROPLAY LEAGUE', { x: 0.446, y: 0.3267, w: 2.7911, h: 0.2524,
                                 fontFace: HEAD, fontSize: 9, bold: true, color: label });
  txt(slide, 'E-SPORTS & GAMING', { x: 8.8588, y: 0.3267, w: 2.7911, h: 0.2524, align: 'right',
                                    fontFace: HEAD, fontSize: 9, bold: true, color: label });
  snip(slide, { x: 11.9165, y: 0.3038, w: 0.9756, h: 0.2882,
                fill: { color: opts.onDark ? DIM.pill : C.purple } });
  txt(slide, 'JOIN US', { x: 11.8189, y: 0.3267, w: 1.1707, h: 0.2524, align: 'center',
                          fontFace: HEAD, fontSize: 9, bold: true,
                          color: opts.onDark ? DIM.pillText : C.white });
  txt(slide, 'PAGE    ' + pageNo, { x: 0.3999, y: 6.9115, w: 3.0, h: 0.3993, valign: 'middle',
                                    fontFace: HEAD, fontSize: 10.5, bold: true, color: label });
}

/* ------------------------------------------------------------------ *
 * Repeated copy
 * ------------------------------------------------------------------ */

const LEAD_ESPORTS = 'eSports is more than a sport—it’s a lifestyle. It unites fans, creators, and players around shared passion and skill.';
const LEAD_DIVE = 'Get ready to dive into the fast-paced world of gaming sports. This is more than just a game—it’s a global movement powered by technology, skill, and community.';
const LEAD_FEMALE = 'Female gamers are leading a wave of change in a traditionally male-dominated arena.';
const LEAD_BRANDS = 'Brands engage fans through tournaments and influencers.';
const LEAD_GEAR = 'Elite gaming needs elite gear.';

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

// 1 — Title: full-bleed plum panel, translucent angled bars, outlined "E".
function slide01(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.darkBg };
  darkDecor(s, [
    { x: -0.5, y: 5.3983, w: 5.9632, h: 0.7609 },
    { x: 9.4775, y: 1.2205, w: 4.4604, h: 0.7609 },
    { x: 11.4769, y: 2.3162, w: 2.4537, h: 0.7609 },
  ], [
    { x: -0.3722, y: 1.1737, w: 1.8222, h: 0.296 },
    { x: 10.9792, y: 6.5308, w: 2.5671, h: 0.296 },
  ]);
  chrome(s, 1, { onDark: true });

  txt(s, '  -SPORT', { x: 1.3177, y: 2.0707, w: 9.8039, h: 2.4234,
                       fontFace: HEAD, fontSize: 138, bold: true, color: C.white });
  txt(s, 'CHAMPIONS', { x: 5.9606, y: 3.9818, w: 7.5581, h: 1.4473,
                        fontFace: HEAD, fontSize: 80, bold: true, color: C.white });
  para(s, 'GAMING SPORT PRESENTATION TEMPLATE',
       { x: 1.4527, y: 4.2361, w: 5.4851, h: 0.3717, color: C.white });
  outlineGlyph(s, 'letterE', 1.5281, 2.5726, 0.7815, 1.3626);
}

// Translucent decoration bars on the two dark slides.
function darkDecor(s, purpleBars, glassBars) {
  purpleBars.forEach(function (b) {
    snip(s, { x: b.x, y: b.y, w: b.w, h: b.h, frac: SNIP.wide, fill: { color: C.purple, transparency: 57 } });
  });
  glassBars.forEach(function (b) {
    snip(s, { x: b.x, y: b.y, w: b.w, h: b.h, frac: SNIP.wide,
              fill: { color: C.white, transparency: 80 },
              line: { color: C.glassEdge, width: 2.5 } });
  });
}

function outlineGlyph(s, name, x, y, w, h) {
  s.addShape('custGeom', { x: x, y: y, w: w, h: h, points: iconPoints(name, w, h),
                           line: { color: C.purpleLight, width: 2.5 } });
}

// 2 — Evolution of competitive gaming.
function slide02(pptx) {
  const s = pptx.addSlide();
  chrome(s, 2);
  imagePlaceholder(s, { x: 1.2346, y: 1.1513, w: 4.1787, h: 5.1975 });

  heading(s, [['EVOLUTION OF COMPETITIVE ', C.heading], ['GAMING', C.purple]],
          { x: 6.1176, y: 1.825, w: 7.3963, h: 1.7166 });
  para(s, LEAD_DIVE, { x: 6.1254, y: 3.5963, w: 5.6524, h: 0.9776 });
  button(s, { x: 4.5722, y: 5.0741, w: 2.9278, h: 0.8194, text: 'COMPETITION' });
  txt(s, '$156.8', { x: 8.1001, y: 5.2075, w: 2.0139, h: 0.5722,
                     fontFace: HEAD, fontSize: 28, bold: true, color: C.heading });
  para(s, 'Global Championship', { x: 10.2638, y: 5.1563, w: 1.75, h: 0.6747, color: C.muted });
}


// 3 — Where gamers build community.
function slide03(pptx) {
  const s = pptx.addSlide();
  chrome(s, 3);
  imagePlaceholder(s, { x: 6.9861, y: 1.039, w: 6.5161, h: 3.037 });

  heading(s, [['WHERE\n', C.heading], ['GAMERS BUILD ', C.heading], ['COMMUNITY', C.purple]],
          { x: 0.8189, y: 1.6163, w: 5.7395, h: 2.5244 });
  para(s, LEAD_ESPORTS, { x: 0.8267, y: 4.2196, w: 5.1594, h: 0.6747 });
  pill(s, { x: 0.9418, y: 5.3084, w: 2.113, h: 0.5753, text: 'READ MORE', labelX: 0.8596 });

  [{ x: 6.6667, icon: 'console', label: 'ONLINE COMPETITION' },
   { x: 10.0579, icon: 'monitor', label: 'GLOBAL CHAMPIONSHIP' }].forEach(function (col) {
    if (col.icon === 'console') {
      icon(s, 'console', col.x + 1.159, 4.9453, 0.3568, 0.2973, C.purple);
    } else {
      icon(s, 'monitorTop', col.x + 1.168, 4.9305, 0.3568, 0.1784, C.purple);
      icon(s, 'monitorBase', col.x + 1.168, 5.1386, 0.3568, 0.1189, C.purple);
    }
    txt(s, col.label, { x: col.x + 0.068, y: 5.5029, w: 2.5544, h: 0.3366, align: 'center',
                        fontFace: HEAD, fontSize: 14, bold: true, color: C.heading });
    para(s, LEAD_GEAR, { x: col.x, y: 5.8799, w: 2.6897, h: 0.3717, align: 'center' });
  });
}

// 4 — Monetizing victory (this slide hides the master chrome except its own copies).
function slide04(pptx) {
  const s = pptx.addSlide();
  imagePlaceholder(s, { x: 1.1915, y: -0.9167, w: 4.2553, h: 6.8954 });

  heading(s, [['MONETIZING VICTORY IN ', C.heading], ['ONLINE ARENAS', C.purple]],
          { x: 6.6667, y: 1.4384, w: 5.658, h: 2.5244 });
  para(s, 'From League of Legends to Call of Duty, explore the titles that dominate global tournaments and fan followings.',
       { x: 6.6745, y: 5.4328, w: 5.4851, h: 0.6747 });
  txt(s, 'PAGE    4', { x: 0.3999, y: 6.9115, w: 3.0, h: 0.3993, valign: 'middle',
                        fontFace: HEAD, fontSize: 10.5, bold: true, color: C.heading });
  snip(s, { x: 11.9165, y: 0.3038, w: 0.9756, h: 0.2882, fill: { color: C.purple } });
  txt(s, 'JOIN US', { x: 11.8189, y: 0.3267, w: 1.1707, h: 0.2524, align: 'center',
                      fontFace: HEAD, fontSize: 9, bold: true, color: C.white });
  txt(s, 'E-SPORTS & GAMING', { x: 8.8588, y: 0.3267, w: 2.7911, h: 0.2524, align: 'right',
                                fontFace: HEAD, fontSize: 9, bold: true, color: C.heading });

  button(s, { x: 4.3301, y: 4.2519, w: 5.3458, h: 0.8194, text: 'ARENA-FAVORITE GAME LEGENDS' });
  txt(s, '2.500+', { x: 10.1418, y: 4.33, w: 2.0139, h: 0.5722,
                     fontFace: HEAD, fontSize: 28, bold: true, color: C.heading });
}

// 5 — Exciting rewards: four staggered prize bars.
function slide05(pptx) {
  const s = pptx.addSlide();
  chrome(s, 5);

  heading(s, [['EXCITING \n', C.heading], ['REWARDS FOR ', C.heading], ['PARTICIPATION', C.purple]],
          { x: 0.8189, y: 1.6163, w: 5.7395, h: 2.5244 });
  para(s, LEAD_ESPORTS, { x: 0.8267, y: 4.2327, w: 5.4851, h: 0.6747 });
  pill(s, { x: 0.9382, y: 5.3084, w: 1.4753, h: 0.5753, text: 'JOIN US', labelX: 0.5371 });
  para(s, 'Online Competition', { x: 2.9233, y: 5.2466, w: 1.3113, h: 0.6747 });
  txt(s, 'TOURNAMENT', { x: 4.3891, y: 5.4317, w: 2.2775, h: 0.4376,
                         fontFace: HEAD, fontSize: 20, bold: true, color: C.heading });

  const prizes = [
    { x: 6.37, w: 7.9003, no: '01', amount: '250.000 IN TOTAL', note: 'Competitive Gaming League', dark: false },
    { x: 7.1502, w: 6.9918, no: '02', amount: '200.000 IN TOTAL', note: 'Global Championship', dark: true },
    { x: 8.0603, w: 6.0832, no: '03', amount: '150.000 IN TOTAL', note: 'Competition', dark: false },
    { x: 9.2172, w: 5.1342, no: '04', amount: '100.000 IN TOTAL', note: '', dark: false },
  ];
  prizes.forEach(function (p, i) {
    const y = 1.5709 + i * 1.3056;
    const fg = p.dark ? C.white : C.heading;
    snip(s, { x: p.x, y: y, w: p.w, h: 0.8874, frac: SNIP.wide,
              fill: { color: p.dark ? C.purple : C.white },
              shadow: p.dark ? undefined : PILL_SHADOW });
    txt(s, p.no, { x: p.x + 0.4814, y: y + 0.0567, w: 1.2042, h: 0.7742,
                   fontFace: HEAD, fontSize: 40, bold: true, color: fg });
    txt(s, [{ text: '$' }, { text: p.amount }], { x: p.x + 1.4453, y: y + 0.2418, w: 3.2373, h: 0.4039,
                                                  valign: 'middle', fontFace: HEAD, fontSize: 18, bold: true, color: fg });
    if (p.note) {
      txt(s, p.note, { x: p.x + 3.7914, y: y + 0.2923, w: 3.0882, h: 0.3029, color: fg });
    }
  });
}

// 6 — How gaming became big business.
function slide06(pptx) {
  const s = pptx.addSlide();
  chrome(s, 6);
  [{ x: 2.1153, y: 0.9829 }, { x: -0.6667, y: 3.9578 }].forEach(function (c) {
    imagePlaceholder(s, { x: c.x, y: c.y, w: 4.8194, h: 2.3678 });
  });

  heading(s, [['HOW GAMING BECAME BIG ', C.heading], ['BUSINESS', C.purple]],
          { x: 7.9737, y: 1.0968, w: 5.7395, h: 2.5244 });
  para(s, LEAD_DIVE, { x: 5.3455, y: 4.4283, w: 6.96, h: 0.6747 });
  gamepadBadge(s, 3.7111, 4.3805);

  [[4.7967, '2.500+', 'Global Championship'],
   [7.5004, '$2.24K', 'Competitive Gaming'],
   [10.2017, '205+', 'Game Development']].forEach(function (st) {
    stat(s, st[0], 5.5813, st[1], st[2], { labelSize: 12 });
  });
}

// 7 — Pocket sized battles: progress bar layout.
function slide07(pptx) {
  const s = pptx.addSlide();
  chrome(s, 7);
  imagePlaceholder(s, { x: 7.6104, y: 0.974, w: 4.5251, h: 5.5519 });

  heading(s, [['POCKET\n', C.heading], ['SIZED BATTLES GO ', C.heading], ['WORLDWIDE', C.purple]],
          { x: 0.8189, y: 1.3328, w: 6.22, h: 2.5244 });
  para(s, LEAD_ESPORTS, { x: 0.8267, y: 3.9362, w: 5.1038, h: 0.6747 });
  para(s, 'ONLINE COMPETITION', { x: 0.8262, y: 4.7675, w: 5.5711, h: 0.46,
                                  fontFace: HEAD, fontSize: 16, bold: true, color: C.heading });
  progressBar(s, 0.9675, 5.3641);
  txt(s, '[ 85% ]       VIRTUAL LEAGUE', { x: 2.5335, y: 5.8306, w: 3.2567, h: 0.3366,
                                           align: 'right', fontFace: HEAD, fontSize: 14, color: C.heading });
  gamepadBadge(s, 11.5722, 2.7972);
}

// 8 — The grind behind every match: four icon tiles.
function slide08(pptx) {
  const s = pptx.addSlide();
  chrome(s, 8);
  heading(s, [['THE GRIND BEHIND ', C.heading], ['EVERY MATCH', C.purple]],
          { x: 0.5974, y: 0.9939, w: 12.0587, h: 0.9088, align: 'center' });
  para(s, LEAD_FEMALE, { x: 2.6424, y: 1.902, w: 7.9687, h: 0.3717, align: 'center' });

  const tiles = [
    { x: 1.1209, glyph: 'console', tag: 'E-SPORTS', title: 'ONLINE COMPETITION', dark: false },
    { x: 4.2222, glyph: 'chess', tag: 'STRATEGY', title: 'GLOBAL CHAMPIONSHIP', dark: true },
    { x: 7.3234, glyph: 'headset', tag: 'STREAMERS', title: 'VICTORY MOMENT', dark: false },
    { x: 10.4247, glyph: 'monitor', tag: 'GAMEPLAY', title: 'GAME DEVELOPMENT', dark: false },
  ];
  tiles.forEach(function (t) {
    const fg = t.dark ? C.white : C.purple;
    snip(s, { x: t.x, y: 3.1389, w: 1.7877, h: 1.5689, frac: SNIP.tile,
              fill: { color: t.dark ? C.purple : C.white },
              shadow: t.dark ? undefined : PILL_SHADOW });
    if (t.glyph === 'monitor') {
      icon(s, 'monitorTop', t.x + 0.6994, 3.5433, 0.3889, 0.1945, fg);
      icon(s, 'monitorBase', t.x + 0.6994, 3.7702, 0.3889, 0.1296, fg);
    } else if (t.glyph === 'chess') {
      icon(s, 'chess', t.x + 0.7156, 3.5271, 0.3565, 0.3889, fg);
    } else {
      icon(s, t.glyph, t.x + 0.6994, 3.5595, 0.3889, 0.3241, fg);
    }
    para(s, t.tag, { x: t.x + 0.095, y: 3.9255, w: 1.5977, h: 0.3714, align: 'center',
                     fontFace: HEAD, bold: true, color: t.dark ? C.white : C.heading });
    para(s, t.title, { x: t.x - 0.5036, y: 5.0758, w: 2.7949, h: 0.4151, align: 'center', valign: 'middle',
                       fontFace: HEAD, fontSize: 14, bold: true, color: C.heading });
    para(s, LEAD_BRANDS, { x: t.x - 0.4109, y: 5.5194, w: 2.6193, h: 0.6747, align: 'center' });
  });
}

// 9 — How great games are built.
function slide09(pptx) {
  const s = pptx.addSlide();
  chrome(s, 9);
  imagePlaceholder(s, { x: -0.6494, y: 1.1322, w: 7.2761, h: 2.3678 });
  imagePlaceholder(s, { x: 7.0573, y: 4.1389, w: 2.9149, h: 2.3678 });

  heading(s, [['HOW ', C.heading], ['GREAT GAMES ', C.purple], ['ARE BUILT', C.heading]],
          { x: 7.2569, y: 1.5467, w: 5.5353, h: 1.7166 });
  para(s, 'DIGITAL CHAMPIONS', { x: 0.7483, y: 4.6505, w: 5.5711, h: 0.46,
                                 fontFace: HEAD, fontSize: 16, bold: true, color: C.heading });
  para(s, 'Get ready to dive into the fast-paced world of gaming sports.\nThis is more than just a game—it’s a global movement powered by technology, skill, and community.',
       { x: 0.7483, y: 5.1088, w: 5.5017, h: 0.9776 });
  gamepadBadge(s, 9.5028, 5.0333);
  stat(s, 11.0, 5.0813, '$2.24K', 'Virtual Champions');
}

// 10 — The icons of esports history: three portrait cards.
function slide10(pptx) {
  const s = pptx.addSlide();
  chrome(s, 10);
  heading(s, [['THE ICONS OF ESPORTS ', C.heading], ['HISTORY', C.purple]],
          { x: 1.809, y: 0.9939, w: 9.6354, h: 0.9088, align: 'center' });
  para(s, LEAD_FEMALE, { x: 3.0293, y: 1.954, w: 7.1948, h: 0.3717, align: 'center' });

  [{ x: 1.2121, name: 'CHRISTOPHER' },
   { x: 5.4025, name: 'THEODORETTE' },
   { x: 9.5929, name: 'MARCELLINDA' }].forEach(function (card) {
    imagePlaceholder(s, { x: card.x, y: 2.987, w: 2.5284, h: 3.2526, frac: SNIP.team });
    const nx = card.x + 1.3971;
    snip(s, { x: nx, y: 5.4019, w: 2.0912, h: 0.5828, fill: { color: C.white }, shadow: PILL_SHADOW });
    snip(s, { x: nx + 0.0642, y: 5.461, w: 1.9529, h: 0.464, fill: { color: C.purple }, shadow: PILL_SHADOW });
    txt(s, card.name, { x: nx + 0.1522, y: 5.5415, w: 1.9691, h: 0.3029,
                        fontFace: HEAD, fontSize: 12, bold: true, color: C.white });
    txt(s, '- Master', { x: nx + 1.1398, y: 5.5751, w: 0.8102, h: 0.2356,
                         align: 'right', fontSize: 8, color: C.white });
  });
}

// 11 — Driven by purpose: big purple portrait panel.
function slide11(pptx) {
  const s = pptx.addSlide();
  chrome(s, 11);
  imagePlaceholder(s, { x: 0.8675, y: 1.2361, w: 5.9588, h: 5.4306, frac: SNIP.half,
                        fill: C.purple, labelColor: 'C9A8DC', cardShadow: false });
  txt(s, 'NATHANIEL ALEXANDER', { x: 1.1721, y: 2.4602, w: 3.2612, h: 0.37,
                                  fontFace: HEAD, fontSize: 16, bold: true, color: C.white });
  txt(s, 'PROFESSIONAL GAMER', { x: 1.1721, y: 2.7818, w: 3.2612, h: 0.2861, fontSize: 11, color: C.white });
  gamepadBadge(s, 1.6022, 5.0127);

  heading(s, [['DRIVEN BY PURPOSE AND ', C.heading], ['PROGRESSION', C.purple]],
          { x: 7.5938, y: 1.5167, w: 5.7395, h: 2.5244 });
  para(s, LEAD_ESPORTS, { x: 7.6016, y: 4.1331, w: 5.2855, h: 0.6747 });
  progressBar(s, 7.7145, 5.1801);
  txt(s, '[ 85% ]          EXPERIENCE', { x: 9.2805, y: 5.6467, w: 3.2567, h: 0.3366,
                                          align: 'right', fontFace: HEAD, fontSize: 14, color: C.heading });
}

// 12 — Streaming sparks global connection.
function slide12(pptx) {
  const s = pptx.addSlide();
  chrome(s, 12);
  imagePlaceholder(s, { x: -0.6234, y: 4.0022, w: 7.2501, h: 2.3678 });

  heading(s, [['STREAMING\n', C.heading], ['SPARKS GLOBAL ', C.heading], ['CONNECTION', C.purple]],
          { x: 0.8056, y: 1.0521, w: 5.9607, h: 2.5244 });
  icon(s, 'gamepad', 7.7735, 1.765, 0.5442, 0.4535, C.purple);
  txt(s, 'COMPETITIVE GAMING', { x: 8.5357, y: 1.7971, w: 2.9612, h: 0.37,
                                 fontFace: HEAD, fontSize: 16, bold: true, color: C.heading });
  para(s, LEAD_ESPORTS, { x: 7.6571, y: 2.359, w: 4.7556, h: 0.6747 });
  button(s, { x: 5.8366, y: 4.4336, w: 2.9442, h: 0.8194, text: 'VIRTUAL LEAGUE' });
  txt(s, '2.500+', { x: 9.3496, y: 4.5117, w: 2.0139, h: 0.5722,
                     fontFace: HEAD, fontSize: 28, bold: true, color: C.heading });
  para(s, LEAD_ESPORTS, { x: 7.2545, y: 5.4629, w: 5.2455, h: 0.6747 });
}

// 13 — Break slide (dark interlude).
function slide13(pptx) {
  const s = pptx.addSlide();
  // This slide paints an opaque accent1 rectangle over everything, so no chrome.
  s.background = { color: C.ink };
  darkDecor(s, [
    { x: 10.5208, y: 1.5671, w: 3.3305, h: 0.7609 },
    { x: -1.0213, y: 3.9612, w: 4.8792, h: 0.9675 },
  ], [
    { x: -0.4861, y: 5.8152, w: 2.0219, h: 0.296 },
    { x: 12.5417, y: 6.414, w: 2.0219, h: 0.296 },
    { x: 9.0463, y: 1.0599, w: 4.6284, h: 0.296 },
  ]);
  txt(s, 'BREAK', { x: 0.8438, y: 1.4153, w: 7.4479, h: 2.5412,
                    fontFace: HEAD, fontSize: 145, bold: true, color: C.white });
  txt(s, 'SLIDE', { x: 7.5938, y: 3.2923, w: 5.7395, h: 2.5412,
                    fontFace: HEAD, fontSize: 145, bold: true, color: C.white });
  txt(s, '30 MINUTES', { x: 1.0728, y: 4.1723, w: 2.9895, h: 0.5722,
                         fontFace: HEAD, fontSize: 28, bold: true, color: C.white });
}

// 14 — Future trends driving gaming's revolution.
function slide14(pptx) {
  const s = pptx.addSlide();
  chrome(s, 14);
  imagePlaceholder(s, { x: 7.3351, y: 3.0519, w: 6.6926, h: 2.9957 });

  heading(s, [['FUTURE TRENDS DRIVING GAMING’S ', C.heading], ['REVOLUTION', C.purple]],
          { x: 0.7583, y: 1.366, w: 10.5794, h: 1.7166 });
  para(s, LEAD_ESPORTS, { x: 0.7661, y: 3.2454, w: 5.4851, h: 0.6747 });
  txt(s, 'GAME DEVELOPMENT', { x: 0.7825, y: 4.1716, w: 2.9612, h: 0.37,
                               fontFace: HEAD, fontSize: 16, bold: true, color: C.heading });

  icon(s, 'console', 0.8539, 4.8974, 0.3568, 0.2973, C.purple);
  txt(s, 'CONSOLE ', { x: 1.5693, y: 4.8778, w: 1.4802, h: 0.3366, valign: 'middle',
                       fontFace: HEAD, fontSize: 14, bold: true, color: C.heading });
  para(s, LEAD_GEAR, { x: 3.0887, y: 4.8602, w: 3.2258, h: 0.3717 });

  icon(s, 'monitorTop', 0.8539, 5.7846, 0.3568, 0.1784, C.purple);
  icon(s, 'monitorBase', 0.8539, 5.9927, 0.3568, 0.1189, C.purple);
  txt(s, 'COMPUTER', { x: 1.5693, y: 5.7798, w: 1.4802, h: 0.3366, valign: 'middle',
                       fontFace: HEAD, fontSize: 14, bold: true, color: C.heading });
  para(s, LEAD_GEAR, { x: 3.0887, y: 5.7622, w: 3.2258, h: 0.3717 });

  gamepadBadge(s, 11.5246, 2.566);
}

// 15 — From concept to competitive play.
function slide15(pptx) {
  const s = pptx.addSlide();
  chrome(s, 15);
  imagePlaceholder(s, { x: 0.922, y: 1.125, w: 2.9929, h: 4.6111 });
  imagePlaceholder(s, { x: 4.8369, y: 4.4785, w: 2.3121, h: 2.1039 });

  heading(s, [['FROM CONCEPT TO ', C.heading], ['COMPETITIVE PLAY', C.purple]],
          { x: 4.7364, y: 1.3458, w: 6.8806, h: 1.7166 });
  para(s, LEAD_ESPORTS, { x: 4.7407, y: 3.1617, w: 5.4851, h: 0.6747 });
  gamepadBadge(s, 6.5802, 4.8841);
  txt(s, 'ONLINE COMPETITION', { x: 8.1297, y: 5.0888, w: 2.9612, h: 0.4039,
                                 fontFace: HEAD, fontSize: 18, bold: true, color: C.heading });
  stat(s, 8.1297, 5.7618, '$2.24K', 'Game Development');
  stat(s, 11.0, 5.7618, '205+', 'Competitive Gaming');
}

// 16 — Gaming spirit fully unleashed (illustration → placeholder block).
function slide16(pptx) {
  const s = pptx.addSlide();
  chrome(s, 16);
  heading(s, [['GAMING SPIRIT FULLY ', C.heading], ['UNLEASHED', C.purple]],
          { x: 0.8568, y: 1.2075, w: 7.5587, h: 1.7166 });
  para(s, 'As the industry grows, so do its challenges—player health, cheating, and ethical play must be addressed.',
       { x: 0.9095, y: 3.0753, w: 4.6877, h: 0.6747 });

  imagePlaceholder(s, { x: 5.9558, y: 2.8344, w: 4.8883, h: 3.8416, frac: 0, cardShadow: false });
  gamepadBadge(s, 9.5211, 2.4159);
  button(s, { x: 3.6409, y: 4.7749, w: 2.9278, h: 0.8194, text: 'COMPETITION' });

  txt(s, '$2.24K', { x: 1.1557, y: 4.8985, w: 2.0139, h: 0.5722, align: 'right',
                     fontSize: 28, bold: true, color: C.heading });
  para(s, LEAD_BRANDS, { x: 0.7963, y: 5.7321, w: 5.0738, h: 0.3717, align: 'right' });
  stat(s, 11.0, 3.647, '205+', 'Champion’s Arena',
       { valueSize: 28, valueH: 0.5722, gap: 0.547, labelSize: 14 });
}

// 17 — Moments captured in competition (laptop photo).
function slide17(pptx) {
  const s = pptx.addSlide();
  chrome(s, 17);
  imagePlaceholder(s, { x: 1.354, y: 1.966, w: 4.415, h: 5.534, frac: 0, cardShadow: false });
  txt(s, 'GLOBAL CHAMPIONSHIP', { x: 1.5535, y: 1.3128, w: 5.5711, h: 0.46,
                                  fontFace: HEAD, fontSize: 16, bold: true, color: C.heading });
  gamepadBadge(s, 1.0359, 2.773);

  heading(s, [['MOMENTS CAPTURED IN ', C.heading], ['COMPETITION', C.purple]],
          { x: 6.6667, y: 1.4384, w: 5.658, h: 2.5244 });
  para(s, LEAD_ESPORTS, { x: 6.6745, y: 4.1828, w: 5.4851, h: 0.6747 });
  button(s, { x: 4.3301, y: 5.3352, w: 5.3458, h: 0.8194, text: 'LEGENDS NEVER STOP PLAYING' });
  txt(s, '2.500+', { x: 10.1418, y: 5.4134, w: 2.0139, h: 0.5722,
                     fontFace: HEAD, fontSize: 28, bold: true, color: C.heading });
}

// 18 — Gamers in full action mode (tablet photo).
function slide18(pptx) {
  const s = pptx.addSlide();
  chrome(s, 18);
  heading(s, [['GAMERS IN FULL ', C.heading], ['ACTION MODE', C.purple]],
          { x: 1.809, y: 0.9939, w: 9.6354, h: 0.9088, align: 'center' });
  para(s, LEAD_FEMALE, { x: 3.0293, y: 1.902, w: 7.1948, h: 0.3717, align: 'center' });

  imagePlaceholder(s, { x: 3.453, y: 3.021, w: 6.133, h: 4.479, frac: 0, cardShadow: false });
  gamepadBadge(s, 8.4387, 3.3403);
  stat(s, 10.0833, 3.3883, '$2.24K', 'Competitive Gaming');
  para(s, 'Winning isn’t just about speed—it’s about composure under pressure. ',
       { x: 0.2988, y: 3.75, w: 3.0621, h: 0.6747, align: 'right' });
  para(s, 'From reflex drills to coaching sessions, performance is everything.',
       { x: 10.1041, y: 4.8986, w: 2.5348, h: 0.9776 });
  button(s, { x: 1.9526, y: 5.0741, w: 2.9278, h: 0.8194, text: 'COMPETITIVE' });
}

// 19 — Let's stay connect with us.
function slide19(pptx) {
  const s = pptx.addSlide();
  chrome(s, 19);
  imagePlaceholder(s, { x: 7.4815, y: 4.0533, w: 7.4213, h: 2.7007, frac: SNIP.half,
                        fill: C.purple, labelColor: 'C9A8DC', cardShadow: false });

  heading(s, [['L', C.heading], ['ET’S STAY CONNECT ', C.heading], ['WITH US', C.purple]],
          { x: 0.9184, y: 1.391, w: 6.8316, h: 1.9186, fontSize: 54 });
  para(s, LEAD_ESPORTS, { x: 0.9184, y: 3.4145, w: 5.3148, h: 0.6746 });

  const rows = [
    { y: 4.365, label: 'TELEPHONE', value: '123-456-7890' },
    { y: 5.067, label: 'LOCATION', value: '123 Anywhere St., Any City' },
    { y: 5.769, label: 'E-MAIL', value: 'yourmail@gmail.com' },
  ];
  rows.forEach(function (r, i) {
    if (i === 0) icon(s, 'phone', 1.0755, r.y + 0.048, 0.2442, 0.2438, C.purple);
    if (i === 1) icon(s, 'pin', 1.0904, r.y + 0.042, 0.2143, 0.2577, C.purple);
    if (i === 2) {
      icon(s, 'mailBody', 1.0755, r.y + 0.06, 0.2442, 0.1776, C.purple);
      icon(s, 'mailFlap', 1.083, r.y + 0.058, 0.2291, 0.1264, C.purple);
    }
    txt(s, r.label, { x: 1.618, y: r.y + 0.002, w: 1.4802, h: 0.3366, valign: 'middle',
                      fontFace: HEAD, fontSize: 14, bold: true, color: C.heading });
    txt(s, r.value, { x: 3.1827, y: r.y, w: 2.6156, h: 0.34, lineSpacingMultiple: 1.3 });
  });
}

// 20 — Thank you (dark closer).
function slide20(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.darkBg };
  darkDecor(s, [
    { x: -0.4778, y: 1.2044, w: 2.2641, h: 0.7609 },
    { x: 10.0502, y: 3.2249, w: 3.7322, h: 0.7609 },
    { x: 12.486, y: 4.4771, w: 1.2779, h: 0.7609 },
  ], [
    { x: 11.808, y: 1.485, w: 1.7383, h: 0.296 },
    { x: -0.7153, y: 5.8266, w: 4.8208, h: 0.296 },
  ]);
  chrome(s, 20, { onDark: true });

  txt(s, '   HANK', { x: 1.9616, y: 2.1444, w: 7.8935, h: 2.5412,
                      fontFace: HEAD, fontSize: 145, bold: true, color: C.white });
  txt(s, 'YOU!', { x: 6.6489, y: 4.0784, w: 3.9297, h: 2.1373,
                   fontFace: HEAD, fontSize: 121, bold: true, color: C.white });
  para(s, 'FOR JOINING THIS PRESENTATION', { x: 2.3222, y: 4.3612, w: 5.4851, h: 0.3717, color: C.white });
  outlineGlyph(s, 'letterT', 2.4164, 2.6788, 0.9895, 1.4293);
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'GAMING_16x9', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'GAMING_16x9';
  pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
  pptx.author = 'Ella Afrilina';
  pptx.title = 'E-SPORT CHAMPIONS';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
   slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach(function (fn) { fn(pptx); });

  return pptx.writeFile({
    fileName: path.join(__dirname, '0653328f-bd8f-438e-9fb7-b114bf96c71d_grok_final.pptx'),
  });
}

build().then(function (f) { console.log('wrote ' + f); }, function (e) { console.error(e); process.exit(1); });
