/**
 * BIKEO — "World Bicycle Day" deck (30 slides, 13.333in x 7.5in widescreen)
 * Standalone pptxgenjs recreation of the reference presentation.
 *
 * Run:  node 12ea8975-f5cc-4bd0-920d-80638da77d86_grok_final.js
 * Out:  12ea8975-f5cc-4bd0-920d-80638da77d86_grok_final.pptx  (next to this file)
 *
 * Photographs in the original are replaced by flat "PLACE YOUR IMAGE"
 * rectangles at the same position/size (see `photo()`).
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette
const NAVY  = '25365C'; // theme tx1 — headings & body copy
const RED   = 'EE5538'; // theme accent5 — highlight colour
const GOLD  = 'FBC352'; // theme accent3 — the translucent "sticky tape" strips
const WHITE = 'F3F3F3'; // theme bg1 — page background
const GREY  = '797979'; // bg1 @ 50% luminance — small caption text
const PHOTO = '4B5A6D'; // flat fill standing in for the reference photographs

const BODY_FONT = 'Raleway';
const HEAD_FONT = 'Raleway ExtraBold';

/** Every slide is a plain page on the off-white background. */
function newSlide(pptx) {
  const slide = pptx.addSlide();
  slide.background = { color: WHITE };
  return slide;
}

// ------------------------------------------------------- text run helpers
// T(text) / T(text, colour)  → regular run;  B(...) → bold run.
const T = (text, color) => ({ text, options: { color: color || null } });
const B = (text, color) => ({ text, options: { color: color || null, bold: true } });

// A paragraph list is either  'plain string'  |  [run, run, ...]
// |  [[run, ...], [run, ...]]  (one inner array per paragraph).
// `bullet` has to ride on every run: pptxgenjs re-emits <a:pPr> per run and the
// last one wins, so a shape-level bullet would be cancelled by later runs.
function paragraphs(content, fallbackColor, bullet) {
  const lines = Array.isArray(content) ? content : [[T(content)]];
  const rows = Array.isArray(lines[0]) ? lines : [lines];
  const out = [];
  rows.forEach((row, r) => {
    row.forEach((run, i) => {
      const opts = Object.assign({}, run.options);
      if (!opts.color) opts.color = fallbackColor;
      if (bullet) opts.bullet = bullet;
      if (i === row.length - 1 && r < rows.length - 1) opts.breakLine = true;
      out.push({ text: run.text, options: opts });
    });
  });
  return out;
}

// Every text style used in the deck. `box` is always [x, y, w, h] in inches.
const STYLES = {
  title:   { fontFace: HEAD_FONT, fontSize: 40, color: NAVY, wrap: false },
  name:    { fontFace: HEAD_FONT, fontSize: 32, color: RED,  wrap: false },
  bignum:  { fontFace: HEAD_FONT, fontSize: 60, color: NAVY, wrap: false },
  tag:     { fontFace: HEAD_FONT, fontSize: 20, color: NAVY, italic: true, wrap: false },
  h2:      { fontFace: BODY_FONT, fontSize: 20, color: NAVY, bold: true, wrap: false },
  body:    { fontFace: BODY_FONT, fontSize: 14, color: NAVY, lineSpacingMultiple: 1.5 },
  small:   { fontFace: BODY_FONT, fontSize: 11, color: NAVY, lineSpacingMultiple: 1.5 },
  bullets: { fontFace: BODY_FONT, fontSize: 11, color: NAVY, lineSpacingMultiple: 1.5 },
  note:    { fontFace: BODY_FONT, fontSize: 9,  color: GREY, bold: true, lineSpacingMultiple: 1.5 },
};

function addText(slide, kind, box, content, extra) {
  const st = STYLES[kind];
  const o = Object.assign({
    x: box[0], y: box[1], w: box[2], h: box[3],
    valign: 'top', align: 'left', wrap: true, fit: 'resize',
  }, st, extra || {});
  let bullet = null;
  if (kind === 'bullets') {
    bullet = { characterCode: '2022', indent: o.indent || 22.5 };
  }
  delete o.indent;
  slide.addText(paragraphs(content, st.color, bullet), o);
}

const title  = (s, box, c, o) => addText(s, 'title',   box, c, o);
const name   = (s, box, c, o) => addText(s, 'name',    box, c, o);
const bignum = (s, box, c, o) => addText(s, 'bignum',  box, c, o);
const tag    = (s, box, c, o) => addText(s, 'tag',     box, c, o);
const h2     = (s, box, c, o) => addText(s, 'h2',      box, c, o);
const body   = (s, box, c, o) => addText(s, 'body',    box, c, o);
const small  = (s, box, c, o) => addText(s, 'small',   box, c, o);
const bullets= (s, box, c, o) => addText(s, 'bullets', box, c, o);
const note   = (s, box, c, o) => addText(s, 'note',    box, c, o);

// ------------------------------------------------------------ shape helpers
const NO_LINE = { type: 'none' };

/** Flat rectangle — used for the bar-chart columns on slides 26 & 28. */
function bar(slide, box, color) {
  slide.addShape('rect', { x: box[0], y: box[1], w: box[2], h: box[3], fill: { color }, line: NO_LINE });
}

/** Small filled circle — the bullet dots on slide 15. */
function dot(slide, box, color) {
  slide.addShape('ellipse', { x: box[0], y: box[1], w: box[2], h: box[3], fill: { color }, line: NO_LINE });
}

/** The recurring translucent gold "sticky tape" strip. */
function tape(slide, box, rotate) {
  slide.addShape('rect', {
    x: box[0], y: box[1], w: box[2], h: box[3], rotate,
    fill: { color: GOLD, transparency: 50 }, line: NO_LINE,
  });
}

/** Stand-in for a reference photograph: flat rectangle + "PLACE YOUR IMAGE". */
function photo(slide, box, fontSize) {
  slide.addShape('rect', { x: box[0], y: box[1], w: box[2], h: box[3], fill: { color: PHOTO }, line: NO_LINE });
  slide.addText(
    [{ text: 'PLACE', options: { breakLine: true } }, { text: 'YOUR IMAGE' }],
    {
      x: box[0], y: box[1], w: box[2], h: box[3],
      fontFace: BODY_FONT, fontSize, bold: true, color: 'FEFEFE', charSpacing: fontSize * 0.26,
      align: 'center', valign: 'middle', lineSpacing: fontSize * 1.21, wrap: false,
    }
  );
}

/**
 * Pictogram person (slide 27), traced as a custom path in a 0.781 x 1.81in box:
 * rounded shoulders, a body split by a leg gap, and a separate circular head.
 */
const PERSON = [
  { x: 0.156, y: 0.442, moveTo: true },
  { x: 0.625, y: 0.442 },
  { x: 0.781, y: 0.598, curve: { type: 'cubic', x1: 0.711, y1: 0.442, x2: 0.781, y2: 0.511 } },
  { x: 0.781, y: 1.020 }, { x: 0.642, y: 1.020 }, { x: 0.642, y: 1.756 },
  { x: 0.587, y: 1.810, curve: { type: 'cubic', x1: 0.642, y1: 1.786, x2: 0.617, y2: 1.810 } },
  { x: 0.461, y: 1.810 },
  { x: 0.407, y: 1.756, curve: { type: 'cubic', x1: 0.431, y1: 1.810, x2: 0.407, y2: 1.785 } },
  { x: 0.407, y: 1.304 },
  { x: 0.392, y: 1.288, curve: { type: 'cubic', x1: 0.407, y1: 1.296, x2: 0.400, y2: 1.288 } },
  { x: 0.376, y: 1.304, curve: { type: 'cubic', x1: 0.383, y1: 1.288, x2: 0.376, y2: 1.296 } },
  { x: 0.376, y: 1.756 },
  { x: 0.322, y: 1.810, curve: { type: 'cubic', x1: 0.376, y1: 1.786, x2: 0.351, y2: 1.810 } },
  { x: 0.196, y: 1.810 },
  { x: 0.142, y: 1.756, curve: { type: 'cubic', x1: 0.166, y1: 1.810, x2: 0.142, y2: 1.785 } },
  { x: 0.142, y: 1.020 }, { x: 0.000, y: 1.020 }, { x: 0.000, y: 0.598 },
  { x: 0.156, y: 0.442, curve: { type: 'cubic', x1: 0.000, y1: 0.511, x2: 0.070, y2: 0.442 } },
  { close: true },
  { x: 0.388, y: 0.000, moveTo: true },
  { x: 0.582, y: 0.194, curve: { type: 'cubic', x1: 0.495, y1: 0.000, x2: 0.582, y2: 0.087 } },
  { x: 0.388, y: 0.387, curve: { type: 'cubic', x1: 0.582, y1: 0.301, x2: 0.495, y2: 0.387 } },
  { x: 0.195, y: 0.194, curve: { type: 'cubic', x1: 0.281, y1: 0.387, x2: 0.195, y2: 0.301 } },
  { x: 0.388, y: 0.000, curve: { type: 'cubic', x1: 0.195, y1: 0.087, x2: 0.281, y2: 0.000 } },
  { close: true },
];

function person(slide, x, y, color) {
  slide.addShape('custGeom', { x, y, w: 0.781, h: 1.81, points: PERSON, fill: { color }, line: NO_LINE });
}

// White glyphs sitting inside the slide-29 diamonds (traced vector outlines).
const GLYPHS = {
  news: { w: 0.489, h: 0.49, points: [
    {x:0.473,y:0.327,moveTo:1}, {x:0.441,y:0.327}, {x:0.441,y:0.032},
    {x:0.409,y:0,curve:{type:'cubic',x1:0.441,y1:0.015,x2:0.426,y2:0}}, {x:0.097,y:0},
    {x:0.065,y:0.032,curve:{type:'cubic',x1:0.079,y1:0,x2:0.065,y2:0.015}}, {x:0.065,y:0.065},
    {x:0.032,y:0.065}, {x:0,y:0.097,curve:{type:'cubic',x1:0.015,y1:0.065,x2:0,y2:0.079}}, {x:0,y:0.458},
    {x:0.032,y:0.49,curve:{type:'cubic',x1:0,y1:0.475,x2:0.015,y2:0.49}}, {x:0.342,y:0.49},
    {x:0.374,y:0.458,curve:{type:'cubic',x1:0.36,y1:0.49,x2:0.374,y2:0.475}}, {x:0.374,y:0.426},
    {x:0.423,y:0.426}, {x:0.489,y:0.361,curve:{type:'cubic',x1:0.46,y1:0.426,x2:0.489,y2:0.397}},
    {x:0.489,y:0.344}, {x:0.473,y:0.327,curve:{type:'cubic',x1:0.489,y1:0.335,x2:0.482,y2:0.327}},
    {close:1}, {x:0.31,y:0.081,moveTo:1}, {x:0.359,y:0.081},
    {x:0.375,y:0.098,curve:{type:'cubic',x1:0.368,y1:0.081,x2:0.375,y2:0.089}},
    {x:0.359,y:0.115,curve:{type:'cubic',x1:0.375,y1:0.107,x2:0.368,y2:0.115}}, {x:0.31,y:0.115},
    {x:0.293,y:0.098,curve:{type:'cubic',x1:0.3,y1:0.115,x2:0.293,y2:0.107}},
    {x:0.31,y:0.081,curve:{type:'cubic',x1:0.293,y1:0.09,x2:0.3,y2:0.081}}, {close:1},
    {x:0.31,y:0.164,moveTo:1}, {x:0.359,y:0.164},
    {x:0.375,y:0.18,curve:{type:'cubic',x1:0.368,y1:0.164,x2:0.375,y2:0.171}},
    {x:0.359,y:0.197,curve:{type:'cubic',x1:0.375,y1:0.19,x2:0.368,y2:0.197}}, {x:0.31,y:0.197},
    {x:0.293,y:0.18,curve:{type:'cubic',x1:0.3,y1:0.197,x2:0.293,y2:0.19}},
    {x:0.31,y:0.164,curve:{type:'cubic',x1:0.293,y1:0.171,x2:0.3,y2:0.164}}, {close:1},
    {x:0.129,y:0.098,moveTo:1}, {x:0.146,y:0.081,curve:{type:'cubic',x1:0.129,y1:0.089,x2:0.137,y2:0.081}},
    {x:0.244,y:0.081}, {x:0.261,y:0.098,curve:{type:'cubic',x1:0.253,y1:0.081,x2:0.261,y2:0.089}},
    {x:0.261,y:0.18}, {x:0.244,y:0.197,curve:{type:'cubic',x1:0.261,y1:0.19,x2:0.253,y2:0.197}},
    {x:0.146,y:0.197}, {x:0.129,y:0.18,curve:{type:'cubic',x1:0.137,y1:0.197,x2:0.129,y2:0.19}},
    {x:0.129,y:0.098}, {close:1}, {x:0.145,y:0.246,moveTo:1}, {x:0.358,y:0.246},
    {x:0.374,y:0.263,curve:{type:'cubic',x1:0.367,y1:0.246,x2:0.374,y2:0.253}},
    {x:0.358,y:0.279,curve:{type:'cubic',x1:0.374,y1:0.272,x2:0.367,y2:0.279}}, {x:0.145,y:0.279},
    {x:0.128,y:0.263,curve:{type:'cubic',x1:0.136,y1:0.279,x2:0.128,y2:0.272}},
    {x:0.145,y:0.246,curve:{type:'cubic',x1:0.129,y1:0.253,x2:0.137,y2:0.246}}, {close:1},
    {x:0.342,y:0.459,moveTo:1}, {x:0.03,y:0.459}, {x:0.03,y:0.098}, {x:0.063,y:0.098}, {x:0.063,y:0.358},
    {x:0.125,y:0.425,curve:{type:'cubic',x1:0.063,y1:0.394,x2:0.09,y2:0.424}},
    {x:0.342,y:0.425,curve:{type:'cubic',x1:0.128,y1:0.425,x2:0.17,y2:0.425}}, {x:0.342,y:0.459}, {close:1},
    {x:0.424,y:0.393,moveTo:1}, {x:0.186,y:0.393},
    {x:0.194,y:0.361,curve:{type:'cubic',x1:0.191,y1:0.384,x2:0.194,y2:0.372}}, {x:0.457,y:0.361},
    {x:0.424,y:0.393,curve:{type:'cubic',x1:0.457,y1:0.378,x2:0.442,y2:0.393}}, {close:1}
  ] },
  folder: { w: 0.491, h: 0.393, points: [
    {x:0.491,y:0.119,moveTo:1}, {x:0.425,y:0.38},
    {x:0.409,y:0.393,curve:{type:'cubic',x1:0.423,y1:0.388,x2:0.417,y2:0.393}}, {x:0.017,y:0.393},
    {x:0.001,y:0.372,curve:{type:'cubic',x1:0.006,y1:0.393,x2:-0.002,y2:0.383}}, {x:0.067,y:0.11},
    {x:0.082,y:0.098,curve:{type:'cubic',x1:0.069,y1:0.103,x2:0.075,y2:0.098}}, {x:0.475,y:0.098},
    {x:0.491,y:0.119,curve:{type:'cubic',x1:0.485,y1:0.098,x2:0.493,y2:0.108}}, {close:1},
    {x:0.034,y:0.103,moveTo:1}, {x:0.082,y:0.066,curve:{type:'cubic',x1:0.04,y1:0.081,x2:0.059,y2:0.066}},
    {x:0.423,y:0.066}, {x:0.377,y:0.033,curve:{type:'cubic',x1:0.416,y1:0.047,x2:0.398,y2:0.033}},
    {x:0.154,y:0.033}, {x:0.126,y:0.005},
    {x:0.115,y:0,curve:{type:'cubic',x1:0.123,y1:0.002,x2:0.119,y2:0}}, {x:0.049,y:0},
    {x:0,y:0.049,curve:{type:'cubic',x1:0.022,y1:0,x2:0,y2:0.022}}, {x:0,y:0.24}, {x:0.034,y:0.103},
    {close:1}
  ] },
  bag: { w: 0.492, h: 0.458, points: [
    {x:0.475,y:0.458,moveTo:1}, {x:0.016,y:0.458},
    {x:0.001,y:0.437,curve:{type:'cubic',x1:0.006,y1:0.458,x2:-0.002,y2:0.448}},
    {x:0.055,y:0.241,curve:{type:'cubic',x1:0.01,y1:0.403,x2:0.039,y2:0.338}},
    {x:0.197,y:0.326,curve:{type:'cubic',x1:0.094,y1:0.283,x2:0.137,y2:0.32}}, {x:0.197,y:0.344},
    {x:0.213,y:0.36,curve:{type:'cubic',x1:0.197,y1:0.353,x2:0.204,y2:0.36}}, {x:0.279,y:0.36},
    {x:0.296,y:0.344,curve:{type:'cubic',x1:0.288,y1:0.36,x2:0.296,y2:0.353}}, {x:0.296,y:0.326},
    {x:0.437,y:0.241,curve:{type:'cubic',x1:0.355,y1:0.32,x2:0.398,y2:0.282}},
    {x:0.492,y:0.437,curve:{type:'cubic',x1:0.454,y1:0.34,x2:0.482,y2:0.401}},
    {x:0.475,y:0.458,curve:{type:'cubic',x1:0.494,y1:0.448,x2:0.485,y2:0.458}}, {close:1},
    {x:0.06,y:0.204,moveTo:1}, {x:0.065,y:0.115,curve:{type:'cubic',x1:0.063,y1:0.176,x2:0.065,y2:0.147}},
    {x:0.082,y:0.098,curve:{type:'cubic',x1:0.065,y1:0.105,x2:0.073,y2:0.098}}, {x:0.148,y:0.098},
    {x:0.246,y:0,curve:{type:'cubic',x1:0.148,y1:0.044,x2:0.191,y2:0}},
    {x:0.344,y:0.098,curve:{type:'cubic',x1:0.3,y1:0,x2:0.344,y2:0.044}}, {x:0.409,y:0.098},
    {x:0.426,y:0.115,curve:{type:'cubic',x1:0.419,y1:0.098,x2:0.426,y2:0.105}},
    {x:0.431,y:0.204,curve:{type:'cubic',x1:0.426,y1:0.147,x2:0.428,y2:0.176}},
    {x:0.295,y:0.294,curve:{type:'cubic',x1:0.413,y1:0.208,x2:0.371,y2:0.283}}, {x:0.295,y:0.278},
    {x:0.278,y:0.261,curve:{type:'cubic',x1:0.295,y1:0.269,x2:0.287,y2:0.261}}, {x:0.212,y:0.261},
    {x:0.196,y:0.278,curve:{type:'cubic',x1:0.203,y1:0.261,x2:0.196,y2:0.269}}, {x:0.196,y:0.294},
    {x:0.06,y:0.204,curve:{type:'cubic',x1:0.121,y1:0.283,x2:0.078,y2:0.208}}, {close:1},
    {x:0.18,y:0.098,moveTo:1}, {x:0.311,y:0.098},
    {x:0.246,y:0.032,curve:{type:'cubic',x1:0.311,y1:0.061,x2:0.282,y2:0.032}},
    {x:0.18,y:0.098,curve:{type:'cubic',x1:0.209,y1:0.032,x2:0.18,y2:0.061}}, {close:1}
  ] },
  basket: { w: 0.491, h: 0.491, points: [
    {x:0.465,y:0.232,moveTo:1}, {x:0.459,y:0.238},
    {x:0.432,y:0.165,curve:{type:'cubic',x1:0.457,y1:0.211,x2:0.447,y2:0.186}},
    {x:0.468,y:0.137,curve:{type:'cubic',x1:0.446,y1:0.158,x2:0.459,y2:0.149}},
    {x:0.489,y:0.065,curve:{type:'cubic',x1:0.485,y1:0.116,x2:0.492,y2:0.091}},
    {x:0.471,y:0.05,curve:{type:'cubic',x1:0.488,y1:0.055,x2:0.48,y2:0.049}},
    {x:0.386,y:0.128,curve:{type:'cubic',x1:0.427,y1:0.054,x2:0.394,y2:0.088}},
    {x:0.328,y:0.115,curve:{type:'cubic',x1:0.368,y1:0.12,x2:0.349,y2:0.115}},
    {x:0.295,y:0.119,curve:{type:'cubic',x1:0.317,y1:0.115,x2:0.305,y2:0.117}},
    {x:0.165,y:0,curve:{type:'cubic',x1:0.289,y1:0.052,x2:0.232,y2:0}},
    {x:0.033,y:0.131,curve:{type:'cubic',x1:0.093,y1:0,x2:0.033,y2:0.058}}, {x:0.033,y:0.238},
    {x:0.027,y:0.232}, {x:0,y:0.245,curve:{type:'cubic',x1:0.017,y1:0.224,x2:0,y2:0.231}}, {x:0,y:0.442},
    {x:0.049,y:0.491,curve:{type:'cubic',x1:0,y1:0.469,x2:0.022,y2:0.491}}, {x:0.442,y:0.491},
    {x:0.491,y:0.442,curve:{type:'cubic',x1:0.469,y1:0.491,x2:0.491,y2:0.469}}, {x:0.491,y:0.245},
    {x:0.465,y:0.232,curve:{type:'cubic',x1:0.491,y1:0.231,x2:0.475,y2:0.223}}, {close:1},
    {x:0.262,y:0.276,moveTo:1}, {x:0.245,y:0.29}, {x:0.173,y:0.232},
    {x:0.151,y:0.234,curve:{type:'cubic',x1:0.167,y1:0.227,x2:0.157,y2:0.228}}, {x:0.096,y:0.289},
    {x:0.065,y:0.264}, {x:0.065,y:0.13},
    {x:0.163,y:0.032,curve:{type:'cubic',x1:0.065,y1:0.076,x2:0.108,y2:0.032}},
    {x:0.261,y:0.13,curve:{type:'cubic',x1:0.217,y1:0.032,x2:0.261,y2:0.076}}, {x:0.261,y:0.276}, {close:1},
    {x:0.423,y:0.265,moveTo:1}, {x:0.394,y:0.288}, {x:0.339,y:0.232},
    {x:0.317,y:0.231,curve:{type:'cubic',x1:0.333,y1:0.226,x2:0.323,y2:0.226}}, {x:0.294,y:0.249},
    {x:0.294,y:0.151}, {x:0.326,y:0.145,curve:{type:'cubic',x1:0.304,y1:0.147,x2:0.316,y2:0.145}},
    {x:0.424,y:0.243,curve:{type:'cubic',x1:0.38,y1:0.145,x2:0.424,y2:0.189}},
    {x:0.423,y:0.265,curve:{type:'cubic',x1:0.425,y1:0.251,x2:0.425,y2:0.259}}, {close:1},
    {x:0.196,y:0.097,moveTo:1}, {x:0.179,y:0.114,curve:{type:'cubic',x1:0.196,y1:0.106,x2:0.189,y2:0.114}},
    {x:0.147,y:0.114}, {x:0.13,y:0.097,curve:{type:'cubic',x1:0.138,y1:0.114,x2:0.13,y2:0.106}},
    {x:0.147,y:0.08,curve:{type:'cubic',x1:0.13,y1:0.088,x2:0.138,y2:0.08}}, {x:0.179,y:0.08},
    {x:0.196,y:0.097,curve:{type:'cubic',x1:0.189,y1:0.081,x2:0.196,y2:0.089}}, {close:1},
    {x:0.196,y:0.179,moveTo:1}, {x:0.179,y:0.196,curve:{type:'cubic',x1:0.196,y1:0.189,x2:0.189,y2:0.196}},
    {x:0.147,y:0.196}, {x:0.13,y:0.179,curve:{type:'cubic',x1:0.138,y1:0.196,x2:0.13,y2:0.189}},
    {x:0.147,y:0.163,curve:{type:'cubic',x1:0.13,y1:0.17,x2:0.138,y2:0.163}}, {x:0.179,y:0.163},
    {x:0.196,y:0.179,curve:{type:'cubic',x1:0.189,y1:0.163,x2:0.196,y2:0.17}}, {close:1}
  ] },
};

/** Draws one white glyph, centred on (cx, cy). */
function glyphIcon(slide, kind, cx, cy) {
  const g = GLYPHS[kind];
  slide.addShape('custGeom', {
    x: cx - g.w / 2, y: cy - g.h / 2, w: g.w, h: g.h,
    points: g.points, fill: { color: WHITE }, line: NO_LINE,
  });
}

/** Pale-gold band behind a slide-29 step. */
function slab(slide, box) {
  slide.addShape('rect', {
    x: box[0], y: box[1], w: box[2], h: box[3],
    fill: { color: GOLD, transparency: 50 }, line: NO_LINE,
  });
}

/** Square rotated 45° — the slide-29 step marker. */
function diamond(slide, box, color) {
  slide.addShape('rect', {
    x: box[0], y: box[1], w: box[2], h: box[3], rotate: 315,
    fill: { color }, line: NO_LINE,
  });
}

// Slide 1 — Cover — BIKEO / World Bicycle Day
function slide01(pptx) {
  const s = newSlide(pptx);
  photo(s, [6.667, 3.25, 5.031, 3.615], 12.1);
  tape(s, [9.872, 3.111, 2.253, 0.611], 28.57);
  title(s, [1.635, 1.326, 1.897, 0.774], 'BIKEO');
  body(s, [1.635, 3.978, 4.246, 2.887], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. '),
    B('Vestibulum', RED),
    B(' feugiat vitae nulla id efficitur. Nullam molestie lectus eget fringilla interdum. '),
    T('Maecenas volutpat nisi vel augue consequat gravida. '),
    B('Integer gravida risus diam, eget malesuada.'),
  ]);
  title(s, [1.635, 2.073, 5.361, 0.774], [
    T('World Bicycle Day', RED),
    T('.'),
  ]);
  tag(s, [0.063, 5.815, 1.664, 0.438], '#BikeoDay', { rotate: 270 });
  note(s, [10.717, 2.082, 1.962, 0.756], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum'),
  ]);
}

// Slide 2 — Introduction
function slide02(pptx) {
  const s = newSlide(pptx);
  photo(s, [1.635, 2.558, 3.606, 3.615], 12.1);
  tape(s, [1.101, 5.757, 2.03, 0.611], 43.19);
  title(s, [1.569, 1.326, 3.74, 0.774], 'Introduction.');
  body(s, [5.968, 2.558, 5.73, 2.18], [
    B('Lorem ipsum dolor sit amet', RED),
    B(', consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur. Nullam molestie lectus eget fringilla interdum. Maecenas volutpat nisi vel augue consequat gravida. Integer gravida risus diam, eget malesuada.'),
  ]);
  note(s, [5.968, 5.645, 2.698, 0.529], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum'),
  ]);
}

// Slide 3 — World Bicycle Day — banner photo + two bullets
function slide03(pptx) {
  const s = newSlide(pptx);
  photo(s, [1.439, 0.767, 10.456, 1.986], 21.1);
  tape(s, [10.508, 2.293, 2.03, 0.611], 327.22);
  title(s, [1.439, 3.363, 5.361, 0.774], [
    T('World '),
    T('Bicycle Day', RED),
    T('.'),
  ]);
  body(s, [1.439, 4.553, 3.922, 2.18], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur. Nullam molestie lectus eget fringilla interdum.'),
  ]);
  bullets(s, [7.403, 4.553, 4.492, 0.901], [
    T('Lorem ipsum dolor sit amet', RED),
    T(', consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur.'),
  ], { indent: 22.5 });
  bullets(s, [7.403, 5.831, 4.492, 0.901], [
    T('Lorem ipsum dolor sit amet', RED),
    T(', consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur.'),
  ], { indent: 22.5 });
  note(s, [7.403, 3.372, 1.962, 0.756], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum'),
  ]);
}

// Slide 4 — For Your Health — two stacked photos
function slide04(pptx) {
  const s = newSlide(pptx);
  photo(s, [0.841, 0.818, 3.978, 2.75], 9.2);
  photo(s, [2.147, 3.932, 3.978, 2.75], 9.2);
  tape(s, [3.467, 0.806, 2.03, 0.611], 37.15);
  title(s, [6.971, 2.793, 4.618, 0.774], [
    T('For Your '),
    T('Health', RED),
    T('.'),
  ]);
  body(s, [6.971, 3.93, 5.194, 1.473], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. '),
    B('Vestibulum feugiat vitae nulla id efficitur. Nullam molestie lectus eget fringilla interdum.', RED),
  ]);
  tag(s, [0.555, 4.546, 1.664, 0.438], '#BikeoDay', { rotate: 270 });
  note(s, [6.971, 5.926, 1.962, 0.756], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum'),
  ]);
}

// Slide 5 — For The Environment — full-height photo right
function slide05(pptx) {
  const s = newSlide(pptx);
  photo(s, [10.281, 0, 3.052, 7.5], 15);
  tape(s, [9.266, 2.096, 2.03, 0.611], 69.81);
  title(s, [1.16, 2.917, 3.82, 1.447], [
    [T('For The')],
    [T('Environment', RED), T('.')],
  ]);
  body(s, [1.16, 4.732, 3.909, 1.827], [
    T('Lorem ipsum dolor sit amet, consectetur adipiscing elit. '),
    B('Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur. Nullam molestie lectus.'),
  ]);
  bullets(s, [5.85, 4.732, 3.351, 0.901], [
    T('Lorem ipsum dolor sit amet', RED),
    T(', consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim.'),
  ], { indent: 13.5 });
  bullets(s, [5.85, 5.657, 3.351, 0.901], [
    T('Lorem ipsum dolor sit amet', RED),
    T(', consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim.'),
  ], { indent: 13.5 });
  tag(s, [1.16, 2.407, 1.664, 0.438], '#BikeoDay');
  note(s, [1.16, 1.344, 1.962, 0.529], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit.'),
  ]);
}

// Slide 6 — To Relive Stress — tall photo left
function slide06(pptx) {
  const s = newSlide(pptx);
  photo(s, [0.736, 0.736, 3.394, 6.028], 12.1);
  tape(s, [3.124, 6.024, 1.484, 0.611], 324.65);
  title(s, [4.762, 1.629, 2.03, 2.121], [
    [T('To')],
    [T('Relive', RED)],
    [T('Stress.')],
  ]);
  note(s, [4.762, 4.386, 1.501, 0.983], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum'),
  ]);
  body(s, [7.299, 1.57, 5.132, 2.18], [
    B('Lorem ipsum ', RED),
    B('dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur. Nullam molestie lectus eget fringilla interdum. '),
    T('Maecenas volutpat nisi vel augue consequat gravida. ', RED),
    B('Integer gravida risus diam, eget malesuada.'),
  ]);
  bullets(s, [7.299, 4.386, 5.132, 0.901], [
    T('Lorem ipsum dolor sit amet', RED),
    T(', consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur.'),
  ], { indent: 22.5 });
  bullets(s, [7.299, 5.429, 5.132, 0.901], [
    T('Lorem ipsum dolor sit amet', RED),
    T(', consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur.'),
  ], { indent: 22.5 });
}

// Slide 7 — To Reduce Traffic
function slide07(pptx) {
  const s = newSlide(pptx);
  photo(s, [8.944, 1.458, 3.394, 3.278], 11.4);
  tape(s, [11.19, 1.496, 1.69, 0.611], 30.17);
  title(s, [5.11, 1.458, 3.114, 1.447], [
    [T('To '), T('Reduce', RED)],
    [T('Traffic.')],
  ]);
  body(s, [0.995, 3.263, 7.229, 1.473], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. ', RED),
    B('Vestibulum feugiat vitae nulla id efficitur. Nullam molestie lectus eget fringilla interdum. Maecenas volutpat nisi vel augue consequat gravida. Integer gravida risus diam.'),
  ]);
  note(s, [0.995, 1.804, 1.904, 0.756], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum'),
  ]);
  small(s, [0.995, 5.439, 4.115, 1.179], 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur. Nullam molestie lectus eget fringilla interdum.');
}

// Slide 8 — To Save Money — half-bleed photo left
function slide08(pptx) {
  const s = newSlide(pptx);
  photo(s, [0, 0, 6, 7.5], 20.2);
  title(s, [6.667, 2.319, 4.43, 0.774], [
    T('To '),
    T('Save', RED),
    T(' Money.'),
  ]);
  body(s, [6.667, 3.694, 5.472, 1.827], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. '),
    T('Praesent bibendum mauris id turpis consectetur dignissim.'),
    B(' Vestibulum feugiat vitae nulla id efficitur. '),
    B('Nullam molestie lectus eget fringilla interdum', RED),
    B('. Maecenas volutpat nisi vel augue consequat gravida.'),
  ]);
  note(s, [6.667, 6.122, 1.904, 0.756], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum'),
  ]);
  tape(s, [5.012, 1.413, 2.253, 0.611], 289.03);
}

// Slide 9 — Slows Global Warming
function slide09(pptx) {
  const s = newSlide(pptx);
  photo(s, [0, 1.611, 6.667, 2.121], 13.5);
  title(s, [7.125, 1.611, 2.86, 2.121], [
    [T('Slows')],
    [T('Global')],
    [T('Warming', RED), T('.')],
  ]);
  body(s, [7.125, 4.194, 4.694, 2.18], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur. '),
    B('Nullam molestie lectus eget fringilla interdum. ', RED),
    B('Maecenas volutpat nisi vel augue consequat gravida.'),
  ]);
  small(s, [2.208, 4.194, 4.458, 1.179], 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur. Nullam molestie lectus eget fringilla interdum.', { align: 'right' });
  tape(s, [5.453, 1.157, 1.511, 0.611], 20.33);
}

// Slide 10 — Zero Emissions
function slide10(pptx) {
  const s = newSlide(pptx);
  photo(s, [7.136, 0.75, 5.447, 2.736], 11);
  title(s, [1.727, 2.781, 4.411, 0.774], [
    T('Zero', RED),
    T(' Emissions.'),
  ]);
  body(s, [1.727, 4.332, 4.795, 2.534], [
    B('Lorem ipsum dolor sit amet, ', RED),
    B('consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur. Nullam molestie lectus eget fringilla interdum. '),
    T('Maecenas volutpat nisi vel augue consequat gravida. '),
    B('Integer gravida risus diam, eget malesuada.'),
  ]);
  body(s, [7.136, 4.332, 5.447, 1.473], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur. Nullam molestie lectus eget fringilla interdum.'),
  ]);
  tag(s, [0.063, 5.815, 1.664, 0.438], '#BikeoDay', { rotate: 270 });
  small(s, [7.136, 6.242, 5.447, 0.624], 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim.');
  tape(s, [6.052, 1.069, 2.253, 0.611], 289.03);
}

// Slide 11 — Faster and Easier — wide photo bottom
function slide11(pptx) {
  const s = newSlide(pptx);
  photo(s, [0.859, 4.521, 8.183, 2.121], 16.5);
  title(s, [9.633, 4.521, 1.995, 2.121], [
    [T('Faster')],
    [T('and')],
    [T('Easier', RED), T('.')],
  ]);
  body(s, [0.865, 0.859, 4.086, 2.887], [
    T('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. '),
    B('Vestibulum feugiat vitae nulla id efficitur. Nullam molestie lectus eget fringilla interdum. '),
    B('Maecenas volutpat nisi vel augue consequat gravida. ', RED),
    B('Integer gravida risus diam, eget malesuada.'),
  ]);
  small(s, [5.817, 0.859, 5.811, 0.901], [
    T('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur. Nullam molestie lectus eget fringilla interdum.', RED),
  ]);
  tag(s, [5.204, 2.695, 1.664, 0.438], '#BikeoDay', { rotate: 270 });
  note(s, [6.667, 2.999, 1.904, 0.756], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum'),
  ]);
  tape(s, [0.345, 4.45, 1.028, 0.611], 289.03);
}

// Slide 12 — Riding a Bicycle — two photos
function slide12(pptx) {
  const s = newSlide(pptx);
  photo(s, [0, 0, 4.647, 7.5], 15.6);
  photo(s, [10.057, 2.458, 2.679, 2.583], 9);
  title(s, [5.077, 3.363, 4.683, 0.774], [
    T('Riding a '),
    T('Bicycle', RED),
    T('.'),
  ]);
  body(s, [0.645, 2.483, 3.358, 2.534], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. '),
    B('Vestibulum feugiat vitae nulla id efficitur', WHITE),
    B('. Nullam molestie lectus eget fringilla interdum.', WHITE),
  ]);
  note(s, [5.077, 4.513, 2.679, 0.529], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum'),
  ]);
  tape(s, [3.899, 1.057, 1.707, 0.611], 69.57);
}

// Slide 13 — Why Should You Bike?
function slide13(pptx) {
  const s = newSlide(pptx);
  photo(s, [0, 4.222, 6.667, 2.639], 13.5);
  title(s, [3.15, 2.303, 3.517, 1.447], [
    [T('Why Should')],
    [T('You '), T('Bike', RED), T('?')],
  ]);
  tag(s, [0.063, 1.252, 1.664, 0.438], '#BikeoDay', { rotate: 270 });
  note(s, [0.676, 2.869, 1.935, 0.756], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum'),
  ]);
  body(s, [8.083, 2.307, 4.574, 2.534], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. '),
    T('Praesent bibendum mauris id turpis consectetur dignissim. '),
    B('Vestibulum feugiat vitae nulla id efficitur. '),
    B('Nullam molestie lectus eget fringilla interdum.', RED),
    B(' Maecenas volutpat nisi vel augue consequat gravida. Integer gravida risus diam, eget malesuada.'),
  ]);
  small(s, [8.083, 5.96, 4.574, 0.901], 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur.');
  tape(s, [5.948, 6.24, 1.028, 0.611], 289.03);
}

// Slide 14 — Benefits of Cycling — bullet column + two photos
function slide14(pptx) {
  const s = newSlide(pptx);
  photo(s, [4.632, 4.106, 4.069, 2.403], 8.2);
  photo(s, [9.264, 4.106, 4.069, 2.403], 8.2);
  title(s, [4.632, 2.077, 3.08, 1.447], [
    [T('Benefits')],
    [T('of '), T('Cycling', RED), T('.')],
  ]);
  body(s, [8.701, 0.991, 3.66, 2.534], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. '),
    B('Vestibulum feugiat vitae nulla id efficitur. ', RED),
    B('Nullam molestie lectus eget fringilla interdum. Maecenas volutpat nisi vel'),
  ]);
  bullets(s, [0.868, 0.991, 2.947, 1.457], 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur.', { indent: 13.5 });
  bullets(s, [0.868, 3.022, 2.947, 1.457], [
    T('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur.', RED),
  ], { indent: 13.5 });
  bullets(s, [0.868, 5.052, 2.947, 1.457], 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur.', { indent: 13.5 });
  note(s, [4.632, 0.991, 2.947, 0.529], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum'),
  ]);
  tape(s, [4.45, 5.995, 1.133, 0.611], 35.41);
}

// Slide 15 — Health Benefits of Cycling — two dot-marked columns
function slide15(pptx) {
  const s = newSlide(pptx);
  photo(s, [0.597, 0.597, 5.736, 2.444], 11.6);
  title(s, [7.189, 1.664, 4.455, 1.447], [
    [T('Health Benefits')],
    [T('of '), T('Cycling', RED), T('.')],
  ]);
  body(s, [2.708, 3.75, 3.625, 2.887], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. '),
    B('Praesent bibendum mauris id turpis consectetur dignissim. ', RED),
    T('Vestibulum feugiat vitae nulla id efficitur. '),
    B('Nullam molestie lectus eget fringilla interdum. Maecenas volutpat nisi vel augue consequat gravida.'),
  ]);
  dot(s, [1.802, 3.905, 0.55, 0.55], RED);
  note(s, [0.859, 5.666, 1.493, 0.983], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum'),
  ], { align: 'right' });
  body(s, [8.762, 3.75, 3.625, 2.887], [
    B('Lorem ipsum ', RED),
    B('dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur. Nullam molestie lectus eget fringilla interdum. '),
    T('Maecenas volutpat nisi vel augue consequat gravida.'),
  ]);
  dot(s, [7.856, 3.905, 0.55, 0.55], RED);
  note(s, [6.913, 5.666, 1.493, 0.983], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum'),
  ], { align: 'right' });
  tape(s, [5.605, 0.724, 1.457, 0.611], 69.41);
}

// Slide 16 — The Solution — two photos + two bullets
function slide16(pptx) {
  const s = newSlide(pptx);
  photo(s, [0, 0.912, 2.472, 4.31], 8.6);
  photo(s, [3.056, 3.347, 4.417, 1.875], 8.9);
  title(s, [3.685, 5.813, 3.787, 0.774], [
    T('The '),
    T('Solution', RED),
    T('.'),
  ]);
  body(s, [8.056, 3.347, 3.764, 3.24], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. '),
    B('Praesent bibendum mauris id turpis consectetur dignissim. ', RED),
    B('Vestibulum feugiat vitae nulla id efficitur. Nullam molestie lectus eget fringilla interdum. Maecenas volutpat nisi vel augue consequat gravida. '),
    T('Integer gravida risus diam, eget malesuada.'),
  ]);
  note(s, [0.979, 5.604, 1.493, 0.983], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum'),
  ], { align: 'right' });
  bullets(s, [3.056, 1.855, 4.417, 0.901], 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur.', { indent: 13.5 });
  bullets(s, [8.056, 1.855, 4.417, 0.901], 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur.', { indent: 13.5 });
  tape(s, [1.972, 0.811, 1.001, 0.611], 59.52);
}

// Slide 17 — Ride with Me — Jason Drips
function slide17(pptx) {
  const s = newSlide(pptx);
  photo(s, [7.222, 0.674, 4.417, 6.153], 14.8);
  title(s, [2.795, 1.684, 3.847, 0.774], [
    T('Ride with '),
    T('Me', RED),
    T('.'),
  ]);
  tag(s, [11.606, 5.815, 1.664, 0.438], '#BikeoDay', { rotate: 270 });
  body(s, [1.196, 4.685, 5.446, 2.18], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. '),
    B('Vestibulum feugiat vitae nulla id efficitur. ', RED),
    B('Nullam molestie lectus eget fringilla interdum. Maecenas volutpat nisi vel augue consequat gravida. '),
    T('Integer gravida risus diam, eget malesuada.'),
  ], { align: 'right' });
  name(s, [3.945, 3.75, 2.697, 0.64], [
    T('Jason Drips', RED),
  ]);
  note(s, [4.75, 0.633, 1.892, 0.756], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum'),
  ], { align: 'right' });
  tape(s, [10.136, 0.786, 2.253, 0.611], 50.1);
}

// Slide 18 — Enjoy The Ride
function slide18(pptx) {
  const s = newSlide(pptx);
  photo(s, [1.361, 3.125, 6.943, 2.715], 14);
  title(s, [3.958, 1.893, 4.346, 0.774], [
    T('Enjoy The '),
    T('Ride', RED),
    T('.'),
  ]);
  tag(s, [6.64, 6.344, 1.664, 0.438], '#BikeoDay');
  body(s, [8.877, 1.893, 3.096, 3.947], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. ', RED),
    B('Vestibulum feugiat vitae nulla id efficitur. Nullam molestie lectus eget fringilla interdum. '),
    T('Maecenas volutpat nisi vel augue consequat gravida. '),
    B('Integer gravida risus diam, eget malesuada.'),
  ]);
  note(s, [1.361, 1.911, 1.892, 0.756], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum'),
  ]);
  tape(s, [0.773, 5.007, 1.601, 0.611], 44.51);
}

// Slide 19 — I Feel Good
function slide19(pptx) {
  const s = newSlide(pptx);
  photo(s, [5.028, 0.611, 7.694, 2.715], 15.6);
  title(s, [1.244, 3.778, 3.41, 0.774], [
    T('I Feel '),
    T('Good', RED),
    T('.'),
  ]);
  body(s, [5.028, 3.778, 5.569, 2.18], [
    B('Lorem ipsum dolor sit amet, ', RED),
    B('consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur. Nullam molestie lectus eget fringilla interdum. Maecenas volutpat nisi vel augue consequat gravida. '),
    T('Integer gravida risus diam, eget malesuada.'),
  ]);
  note(s, [1.244, 2.57, 1.892, 0.756], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum'),
  ]);
  small(s, [1.244, 5.056, 3.41, 0.901], 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim.');
  tape(s, [4.651, 0.602, 1.409, 0.611], 315.26);
}

// Slide 20 — Freedom Machine — full-width photo band
function slide20(pptx) {
  const s = newSlide(pptx);
  photo(s, [0, 0, 13.333, 3.458], 27);
  title(s, [6.667, 4.042, 5.226, 0.774], [
    T('Freedom '),
    T('Machine', RED),
    T('.'),
  ]);
  body(s, [1.139, 4.042, 4.764, 2.534], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. '),
    B('Praesent bibendum mauris id turpis consectetur dignissim. ', RED),
    B('Vestibulum feugiat vitae nulla id efficitur. Nullam molestie lectus eget fringilla interdum. Maecenas volutpat nisi vel augue consequat gravida. '),
    T('Integer gravida risus diam, eget malesuada.'),
  ], { align: 'right' });
  small(s, [6.667, 5.674, 3.41, 0.901], 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim.');
  tape(s, [10.484, 3.1, 2.253, 0.611], 17.85);
}

// Slide 21 — Why Ride?
function slide21(pptx) {
  const s = newSlide(pptx);
  photo(s, [6.667, 0.819, 5.847, 3.01], 11.8);
  title(s, [3.109, 3.056, 3.086, 0.774], [
    T('Why '),
    T('Ride', RED),
    T('?'),
  ]);
  body(s, [1.431, 4.264, 4.764, 2.534], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. '),
    B('Vestibulum feugiat vitae nulla id efficitur. ', RED),
    B('Nullam molestie lectus eget fringilla interdum. Maecenas volutpat nisi vel augue consequat gravida. Integer gravida risus diam, '),
    T('eget malesuada.'),
  ], { align: 'right' });
  body(s, [6.667, 4.264, 5.847, 1.827], [
    B('Lorem ipsum dolor sit amet, ', RED),
    B('consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur. Nullam molestie lectus eget fringilla interdum. '),
    T('Maecenas volutpat nisi vel augue consequat gravida.'),
  ]);
  note(s, [4.303, 1.947, 1.892, 0.756], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum'),
  ], { align: 'right' });
  tape(s, [11.6, 0.764, 1.358, 0.611], 42.53);
}

// Slide 22 — Improves Mental Health
function slide22(pptx) {
  const s = newSlide(pptx);
  photo(s, [0.708, 0.681, 3.292, 2.444], 8.2);
  photo(s, [0.708, 3.745, 3.292, 3.074], 11.1);
  title(s, [4.624, 4.181, 4.218, 1.447], [
    [T('Improves')],
    [T('Mental Health', RED), T('.')],
  ]);
  body(s, [4.624, 2.272, 7.654, 1.473], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. '),
    T('Praesent bibendum mauris id turpis consectetur dignissim.'),
    B(' Vestibulum feugiat vitae nulla id efficitur. '),
    B('Nullam molestie lectus eget fringilla interdum. ', RED),
    B('Maecenas volutpat nisi vel augue consequat gravida. Integer gravida risus diam, eget malesuada.'),
  ]);
  note(s, [4.624, 6.064, 1.892, 0.756], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum'),
  ]);
  tape(s, [3.245, 0.552, 1.215, 0.611], 37.66);
}

// Slide 23 — Faster Than Walking
function slide23(pptx) {
  const s = newSlide(pptx);
  photo(s, [9.292, 0.73, 3.292, 4.08], 11.1);
  title(s, [5.98, 2.69, 2.618, 2.121], [
    [T('Faster')],
    [T('Than')],
    [T('Walking', RED), T('.')],
  ]);
  body(s, [0.874, 1.923, 4.279, 2.887], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. '),
    B('Praesent bibendum mauris id turpis consectetur dignissim.', RED),
    B(' Vestibulum feugiat vitae nulla id efficitur. Nullam molestie lectus eget fringilla interdum. Maecenas volutpat nisi vel augue consequat gravida. '),
    T('Integer gravida risus diam, eget malesuada.'),
  ]);
  note(s, [5.98, 1.923, 2.484, 0.529], [
    B('Lorem ipsum dolor amet, consectetur adipiscing elit. Praesent bibendum'),
  ]);
  small(s, [0.874, 5.479, 4.279, 1.179], 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur. Nullam molestie lectus eget fringilla interdum.');
  small(s, [5.98, 5.479, 4.279, 1.179], 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. Vestibulum feugiat vitae nulla id efficitur. Nullam molestie lectus eget fringilla interdum.');
  tape(s, [8.595, 0.722, 1.393, 0.611], 289.03);
}

// Slide 24 — Reasons for Cycling — full-width photo band
function slide24(pptx) {
  const s = newSlide(pptx);
  photo(s, [0, 4.292, 13.333, 3.208], 27);
  title(s, [6.667, 3.008, 5.726, 0.774], [
    T('Reasons for '),
    T('Cycling', RED),
    T('.'),
  ]);
  body(s, [1.653, 1.248, 4.542, 2.534], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. '),
    B('Vestibulum feugiat vitae nulla id efficitur. ', RED),
    B('Nullam molestie lectus eget fringilla interdum. Maecenas volutpat nisi vel augue consequat gravida. Integer gravida risus diam, '),
    T('eget malesuada.'),
  ]);
  tag(s, [0.063, 2.731, 1.664, 0.438], '#BikeoDay', { rotate: 270 });
  note(s, [6.667, 1.248, 1.892, 0.756], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum'),
  ]);
  tape(s, [11.974, 3.957, 1.089, 0.611], 13.64);
}

// Slide 25 — Go Out!
function slide25(pptx) {
  const s = newSlide(pptx);
  photo(s, [5.333, 0, 8, 7.5], 26.9);
  title(s, [2.57, 2.976, 2.278, 0.774], [
    T('Go '),
    T('Out', RED),
    T('!'),
  ]);
  body(s, [1.014, 4.11, 3.833, 2.18], [
    B('Lorem ipsum dolor sit amet, ', RED),
    B('consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. '),
    T('Vestibulum feugiat vitae nulla id efficitur. '),
    B('Nullam molestie lectus eget fringilla interdum.'),
  ], { align: 'right' });
  note(s, [2.57, 2.088, 2.278, 0.529], [
    B('Lorem ipsum dolor sit amet, adipiscing elit. Praesent bibendum'),
  ], { align: 'right' });
  tape(s, [4.705, 0.718, 1.279, 0.611], 289.03);
}

// Slide 26 — Stacked-bar infographic (4 columns)
// Each column stacks a RED block (upper share) on top of a NAVY block (lower share).
function slide26(pptx) {
  const s = newSlide(pptx);
  h2(s, [1.157, 1.604, 3.962, 0.438], [B('Lorem ipsum dolor sit amet.')]);
  h2(s, [1.157, 2.042, 3.878, 0.438], [B('Consectetur adipiscing elit.', RED)]);
  small(s, [1.157, 3.884, 3.878, 2.012], 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Morbi in suscipit elit. Sed sagittis quis nunc ut molestie. Duis urna eros, pellentesque id nibh quis, interdum pharetra arcu. Vivamus pellentesque quis ipsum eu egestas. Praesent luctus sollicitudin felis, non ornare velit rhoncus sed. Morbi ligula quam, consectetur in neque eget.');

  const COL_W = 1.375;
  const CAPTIONS = ['Lorem ipsum dolor amet, consectetur', 'Lorem ipsum dolor amet, consectetur', 'Lorem ipsum dolor', 'Lorem ipsum dolor'];
  // [x, navyY, navyH, redY, redH, redLabel, labelX, labelY, navyLabel, navyLabelX, captionH]
  const COLUMNS = [
    [5.968, 3.306, 2.59, 2.292, 1.014, '20%', 6.28, 2.583, '60%', 6.273, 0.901],
    [7.579, 4.167, 1.729, 2.07, 2.097, '35%', 7.888, 3.531, '40%', 7.888, 0.901],
    [9.19, 3.75, 2.146, 2.903, 0.847, '15%', 9.519, 3.222, '50%', 9.502, 0.624],
    [10.802, 3.306, 2.59, 1.722, 1.583, '30%', 11.124, 2.767, '60%', 11.118, 0.624]
  ];
  COLUMNS.forEach(([x, ny, nh, ry, rh, redPct, lx, ly, navyPct, nlx, capH], i) => {
    bar(s, [x, ny, COL_W, nh], NAVY);
    bar(s, [x, ry, COL_W, rh], RED);
    h2(s, [lx, ly, 0.76, 0.438], [B(redPct)], { align: 'center' });
    h2(s, [nlx, 5.201, 0.76, 0.438], [B(navyPct, WHITE)], { align: 'center' });
    small(s, [x, 6.038, COL_W, capH], CAPTIONS[i]);
  });
}

// Slide 27 — Pictogram infographic (2 rows of 8 people)
// Two rows of eight figures; the trailing N of each row are highlighted RED.
function slide27(pptx) {
  const s = newSlide(pptx);
  h2(s, [1.073, 1.62, 1.973, 0.774], [[B('About Fact :')], [B('Lorem Ipsum', RED)]]);
  small(s, [1.073, 3.957, 1.857, 2.568], 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Morbi in suscipit elit. Sed sagittis quis nunc ut molestie. Duis urna eros, pellentesque id nibh quis, interdum pharetra arcu. ');

  const ROWS = [
    { y: 1.62, highlighted: 4, captionY: 3.536 },
    { y: 4.308, highlighted: 2, captionY: 6.224 },
  ];
  const FIRST_X = 3.604, STEP_X = 1.125, COUNT = 8;
  ROWS.forEach(({ y, highlighted, captionY }) => {
    for (let i = 0; i < COUNT; i++) {
      person(s, FIRST_X + i * STEP_X, y, i >= COUNT - highlighted ? RED : NAVY);
    }
    small(s, [3.601, captionY, 5.712, 0.346], 'Donec interdum felis et urna interdum, vel porta est condimentum. Donec.');
  });
}

// Slide 28 — Column-chart infographic + six figures
// Six full-height RED tracks with NAVY columns growing from the baseline,
// plus a 2x3 grid of value + caption pairs on the right.
function slide28(pptx) {
  const s = newSlide(pptx);
  h2(s, [1.371, 1.165, 3.962, 0.438], [B('Lorem ipsum dolor sit amet.')]);
  h2(s, [1.371, 1.603, 3.878, 0.438], [B('Consectetur adipiscing elit.', RED)]);
  small(s, [7.167, 1.191, 4.796, 0.901], 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Morbi in suscipit elit. Sed sagittis quis nunc ut molestie. Duis urna eros, pellentesque id nibh quis, interdum pharetra arcu. ');

  const TRACK_Y = 2.361, TRACK_H = 3.945, BAR_W = 0.638;
  // [x, navyY, navyH] — the NAVY column sits inside a full-height RED track.
  const BARS = [
    [1.371, 4.743, 1.564],
    [2.203, 4.281, 2.025],
    [3.034, 4.413, 1.893],
    [3.866, 3.923, 2.383],
    [4.697, 3.421, 2.886],
    [5.529, 3.046, 3.261]
  ];
  BARS.forEach(([x, y, h]) => {
    bar(s, [x, TRACK_Y, BAR_W, TRACK_H], RED);
    bar(s, [x, y, BAR_W, h], NAVY);
  });

  const STAT_CAPTION = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. ';
  // [value, x, y]
  const STATS = [
    ['23,5', 7.197, 2.506],
    ['30,2', 9.996, 2.506],
    ['28,8', 7.197, 3.897],
    ['35,5', 9.996, 3.897],
    ['40,4', 7.197, 5.289],
    ['45,8', 9.996, 5.289]
  ];
  STATS.forEach(([value, x, y]) => {
    h2(s, [x, y, 0.765, 0.438], [B(value, RED)]);
    small(s, [x, y + 0.462, 2.199, 0.624], STAT_CAPTION);
  });
}

// Slide 29 — Four-step diamond infographic
// Four pale-gold bands, each fronted by a rotated square (diamond) holding a
// white glyph, then the step number and the copy.
function slide29(pptx) {
  const s = newSlide(pptx);
  const BAND_H = 1.392, DIA = 1.292, STEP_TEXT = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam hendrerit ligula. Donec nec libero at justo gravida sollicitudin a';
  // [number, diamondColor, bandX, bandY, bandW, diaX, diaY, numX, numY, numW, textX, textY, glyph]
  const STEPS = [
    ['1', RED, 7.201, 1.451, 4.823, 6.555, 1.501, 8.328, 1.587, 0.638, 9.267, 1.551, 'news'],
    ['2', NAVY, 1.309, 2.52, 4.823, 5.487, 2.569, 1.559, 2.655, 0.695, 2.499, 2.62, 'folder'],
    ['3', RED, 7.201, 3.589, 4.823, 6.555, 3.638, 8.328, 3.733, 0.695, 9.267, 3.698, 'bag'],
    ['4', NAVY, 1.309, 4.657, 4.823, 5.486, 4.707, 1.559, 4.802, 0.695, 2.499, 4.766, 'basket']
  ];
  STEPS.forEach(([n, color, bx, by, bw, dx, dy, nx, ny, nw, tx, ty, glyph]) => {
    slab(s, [bx, by, bw, BAND_H]);
    diamond(s, [dx, dy, DIA, DIA], color);
    glyphIcon(s, glyph, dx + DIA / 2, dy + DIA / 2);
    bignum(s, [nx, ny, nw, 1.111], n);
    small(s, [tx, ty, 2.597, 1.182], STEP_TEXT);
  });
}

// Slide 30 — Thanks
function slide30(pptx) {
  const s = newSlide(pptx);
  title(s, [1.635, 1.326, 2.3, 0.774], 'Thanks.');
  body(s, [1.635, 3.993, 5.361, 2.18], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum mauris id turpis consectetur dignissim. '),
    B('Vestibulum', RED),
    B(' feugiat vitae nulla id efficitur. Nullam molestie lectus eget fringilla interdum. '),
    T('Maecenas volutpat nisi vel augue consequat gravida. '),
    B('Integer gravida risus diam, eget malesuada.'),
  ]);
  title(s, [1.635, 2.073, 5.361, 0.774], [
    T('World Bicycle Day', RED),
    T('.'),
  ]);
  tag(s, [0.063, 5.123, 1.664, 0.438], '#BikeoDay', { rotate: 270 });
  note(s, [7.518, 5.418, 1.962, 0.756], [
    B('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent bibendum'),
  ]);
}

// ------------------------------------------------------------------- build
const SLIDES = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
  slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
  slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24,
  slide25, slide26, slide27, slide28, slide29, slide30,
];

function build() {
  const pptx = new PptxGenJS();
  pptx.title = 'BIKEO — World Bicycle Day';
  pptx.layout = 'LAYOUT_WIDE'; // 13.333in x 7.5in
  SLIDES.forEach((fn) => fn(pptx));
  return pptx.writeFile({
    fileName: path.join(__dirname, '12ea8975-f5cc-4bd0-920d-80638da77d86_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
