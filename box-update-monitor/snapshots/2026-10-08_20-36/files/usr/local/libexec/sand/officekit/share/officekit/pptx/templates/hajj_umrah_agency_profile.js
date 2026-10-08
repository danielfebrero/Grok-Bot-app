/**
 * Rihlah — Hajj & Umrah Agency deck (30 slides, 13.333 x 7.5 in).
 * Standalone pptxgenjs recreation of the reference presentation.
 *
 * Photographs in the original are replaced by flat grey placeholder boxes
 * labelled "[image]", per the conversion brief.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */
const TEAL = '0B9491';   // accent2 - primary brand colour
const TEAL_LT = '67B9B1';   // accent3
const SAND = 'C8AA83';   // accent1 - gold ornament colour
const SAND_LT = 'CEB38F';   // accent4
const DARK = '3C3C3C';   // background
const SLATE = '313C41';   // dark text on light cards
const WHITE = 'FFFFFF';
const GREY = 'D9D9D9';
const IMG_GREY = '9D9D9D';   // fill used for image placeholders

const HEAD = 'DM Serif Display';
const BODY = 'roboto';

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

/* -------------------------------------------------------------- utilities */
function newSlide(pptx) {
  const s = pptx.addSlide();
  s.background = { color: DARK };
  return s;
}

/** addText with the deck's zero-ish insets and top anchoring. */
function txt(s, text, o) {
  s.addText(text, Object.assign({
    valign: 'top',
    margin: [7.2, 7.2, 3.6, 3.6],   // 0.1" / 0.05" as points
    isTextBox: true,
  }, o));
}

/** Grey box standing in for a photo, captioned "[image]". */
function imageBox(s, box, geo) {
  const shape = geo === 'ellipse' ? 'ellipse' : geo === 'roundRect' ? 'roundRect' : 'rect';
  s.addShape(shape, Object.assign({ fill: { color: IMG_GREY } }, box));
  s.addText('[image]', Object.assign({}, box, {
    align: 'center', valign: 'middle', fontFace: BODY, fontSize: 11, color: WHITE,
  }));
}

/* ------------------------------------------------------- ornament: 8-point star */
/** Points of an 8-pointed (Rub el Hizb) star inscribed in the given box. */
function starPoints(box, scale) {
  const cx = box.w / 2, cy = box.h / 2;
  const rOuter = Math.min(box.w, box.h) / 2 * scale;
  const rInner = rOuter * 0.765;
  const pts = [];
  for (let i = 0; i < 16; i++) {
    const a = Math.PI / 8 * i - Math.PI / 2;
    const r = i % 2 === 0 ? rOuter : rInner;
    pts.push({ x: cx + r * Math.cos(a), y: cy + r * Math.sin(a), moveTo: i === 0 });
  }
  pts.push({ close: true });
  return pts;
}

/** Solid 8-point star badge (used behind the section numbers). */
function star8(s, box, color) {
  s.addShape('custGeom', Object.assign({}, box, {
    fill: { color: color }, points: starPoints(box, 1),
  }));
}

/**
 * The gold corner ornament: four concentric stars in one path. Even-odd
 * filling turns the alternating rings into two hairline star outlines.
 */
function arabesque(s, box, color, transparency) {
  const RINGS = [1.0, 0.906, 0.827, 0.749];
  let pts = [];
  RINGS.forEach(function (k) { pts = pts.concat(starPoints(box, k)); });
  s.addShape('custGeom', Object.assign({}, box, {
    fill: { color: color, transparency: transparency || 0 }, points: pts,
  }));
}

/* ------------------------------------------------------------ brand mark */
/** The little gold cube + "Rihlah" wordmark repeated at the top-left. */
function logoMark(s, box, color) {
  const w = box.w, h = box.h;
  s.addShape('custGeom', {
    x: box.x, y: box.y, w: w, h: h, rotate: 90, fill: { color: color },
    points: [
      { x: 0.381 * w, y: 1.0 * h }, { x: 0.584 * w, y: 0.5 * h }, { x: 0.381 * w, y: 0 },
      { x: 0.797 * w, y: 0 }, { x: 1.0 * w, y: 0.5 * h }, { x: 0.797 * w, y: 1.0 * h },
      { close: true },
      { x: 0, y: 0.5 * h, moveTo: true }, { x: 0.194 * w, y: 0 }, { x: 0.294 * w, y: 0 },
      { x: 0.488 * w, y: 0.5 * h }, { x: 0.294 * w, y: 1.0 * h }, { x: 0.194 * w, y: 1.0 * h },
      { close: true },
    ],
  });
}

function brandMark(s) {
  logoMark(s, { x: 0.716, y: 0.408, w: 0.268, h: 0.217 }, SAND);
  txt(s, 'Rihlah', { x: 0.924, y: 0.339, w: 0.807, h: 0.353, fontFace: HEAD, fontSize: 15, color: TEAL, wrap: false });
}

/* ----------------------------------------------------------------- icons */
/**
 * Flat gold pictograms built from native shapes. `kind` picks the glyph;
 * every glyph is drawn inside the given box.
 */
function icon(s, box, color, kind) {
  const x = box.x, y = box.y, w = box.w, h = box.h;
  const fill = { color: color };
  const line = { color: color, width: Math.max(1, Math.round(w * 6)) };
  const put = function (shape, o) { s.addShape(shape, Object.assign({ fill: fill }, o)); };
  const ring = function (o) { s.addShape('ellipse', Object.assign({ fill: { type: 'none' }, line: line }, o)); };

  switch (kind) {
    case 'arrow':                                     // → in package headers
      put('rightArrow', { x: x, y: y + h * 0.3, w: w, h: h * 0.4 });
      break;
    case 'starBadge':                                 // star outline + crescent
    case 'moonBadge':
      arabesque(s, { x: x, y: y, w: w, h: h }, color);
      s.addShape('moon', { x: x + w * 0.34, y: y + h * 0.27, w: w * 0.3, h: h * 0.46, fill: fill, rotate: 180 });
      if (kind === 'starBadge') put('star5', { x: x + w * 0.5, y: y + h * 0.3, w: w * 0.2, h: h * 0.2 });
      break;
    case 'mosque':                                    // dome, arch and minarets
      put('chord', { x: x + w * 0.22, y: y + h * 0.08, w: w * 0.56, h: h * 0.5, rotate: 180 });
      put('rect', { x: x + w * 0.16, y: y + h * 0.5, w: w * 0.68, h: h * 0.42 });
      put('rect', { x: x, y: y + h * 0.3, w: w * 0.1, h: h * 0.62 });
      put('rect', { x: x + w * 0.9, y: y + h * 0.3, w: w * 0.1, h: h * 0.62 });
      s.addShape('chord', { x: x + w * 0.38, y: y + h * 0.6, w: w * 0.24, h: h * 0.32, fill: { color: DARK }, rotate: 180 });
      break;
    case 'person': case 'man': case 'woman': case 'avatar':
      put('ellipse', { x: x + w * 0.3, y: y, w: w * 0.4, h: h * 0.4 });
      put('chord', { x: x + w * 0.1, y: y + h * 0.5, w: w * 0.8, h: h * 0.9, rotate: 180 });
      break;
    case 'globe':
      ring({ x: x, y: y, w: w, h: h });
      ring({ x: x + w * 0.3, y: y, w: w * 0.4, h: h });
      s.addShape('line', { x: x, y: y + h / 2, w: w, h: 0, line: line });
      break;
    case 'phone':
      put('roundRect', { x: x + w * 0.05, y: y, w: w * 0.9, h: h, rectRadius: 0.25 });
      s.addShape('rect', { x: x + w * 0.2, y: y + h * 0.15, w: w * 0.6, h: h * 0.6, fill: { color: DARK } });
      break;
    case 'send':                                      // paper plane
      s.addShape('custGeom', {
        x: x, y: y, w: w, h: h, fill: fill,
        points: [{ x: 0, y: h * 0.45 }, { x: w, y: 0 }, { x: w * 0.45, y: h }, { x: w * 0.36, y: h * 0.62 }, { close: true }],
      });
      break;
    case 'plane':                                     // departure
      s.addShape('custGeom', {
        x: x, y: y, w: w, h: h * 0.7, fill: fill,
        points: [{ x: 0, y: h * 0.34 }, { x: w * 0.28, y: h * 0.16 }, { x: w * 0.55, y: h * 0.3 },
        { x: w, y: 0 }, { x: w * 0.82, y: h * 0.4 }, { x: w * 0.2, y: h * 0.56 }, { close: true }],
      });
      put('rect', { x: x, y: y + h * 0.82, w: w * 0.8, h: h * 0.12 });
      break;
    case 'bed':
      s.addShape('line', { x: x, y: y + h * 0.15, w: 0, h: h * 0.85, line: line });
      put('roundRect', { x: x + w * 0.12, y: y + h * 0.15, w: w * 0.35, h: h * 0.35, rectRadius: 0.4 });
      put('rect', { x: x + w * 0.5, y: y + h * 0.12, w: w * 0.5, h: h * 0.38 });
      put('rect', { x: x, y: y + h * 0.55, w: w, h: h * 0.18 });
      break;
    case 'cloche':                                    // meal service
      put('chord', { x: x + w * 0.05, y: y + h * 0.12, w: w * 0.9, h: h * 1.1, rotate: 180 });
      put('ellipse', { x: x + w * 0.42, y: y, w: w * 0.16, h: h * 0.16 });
      put('rect', { x: x, y: y + h * 0.72, w: w, h: h * 0.1 });
      break;
    case 'bell':
      put('chord', { x: x + w * 0.12, y: y + h * 0.15, w: w * 0.76, h: h * 1.1, rotate: 180 });
      put('ellipse', { x: x + w * 0.44, y: y, w: w * 0.12, h: h * 0.12 });
      put('rect', { x: x + w * 0.05, y: y + h * 0.7, w: w * 0.9, h: h * 0.09 });
      put('rect', { x: x + w * 0.42, y: y + h * 0.82, w: w * 0.16, h: h * 0.1 });
      break;
    case 'pin':                                       // map marker
      put('teardrop', { x: x, y: y, w: w, h: h * 0.85, rotate: 225 });
      s.addShape('ellipse', { x: x + w * 0.3, y: y + h * 0.22, w: w * 0.4, h: h * 0.32, fill: { color: DARK } });
      break;
    case 'calendar':
      put('rect', { x: x, y: y + h * 0.12, w: w, h: h * 0.88 });
      s.addShape('rect', { x: x + w * 0.1, y: y + h * 0.38, w: w * 0.8, h: h * 0.48, fill: { color: DARK } });
      put('rect', { x: x + w * 0.2, y: y, w: w * 0.1, h: h * 0.22 });
      put('rect', { x: x + w * 0.7, y: y, w: w * 0.1, h: h * 0.22 });
      break;
    case 'bus':
      put('roundRect', { x: x, y: y, w: w, h: h * 0.78, rectRadius: 0.2 });
      s.addShape('rect', { x: x + w * 0.12, y: y + h * 0.14, w: w * 0.76, h: h * 0.28, fill: { color: DARK } });
      put('ellipse', { x: x + w * 0.12, y: y + h * 0.7, w: w * 0.22, h: h * 0.3 });
      put('ellipse', { x: x + w * 0.66, y: y + h * 0.7, w: w * 0.22, h: h * 0.3 });
      break;
    case 'shield':
      s.addShape('custGeom', {
        x: x, y: y, w: w, h: h, fill: fill,
        points: [{ x: w / 2, y: 0 }, { x: w, y: h * 0.2 }, { x: w, y: h * 0.6 }, { x: w / 2, y: h },
        { x: 0, y: h * 0.6 }, { x: 0, y: h * 0.2 }, { close: true }],
      });
      s.addShape('mathPlus', { x: x + w * 0.24, y: y + h * 0.24, w: w * 0.52, h: h * 0.42, fill: { color: DARK } });
      break;
    default:
      put('ellipse', { x: x, y: y, w: w, h: h });
  }
}

/* ------------------------------------------------------------- world map */
/** Simplified continent outlines (unit coordinates) for the "Global Pilgrims" slide. */
function worldMap(s, box) {
  WORLD.forEach(function (flat) {
    const pts = [];
    for (let i = 0; i < flat.length; i += 2) {
      pts.push({ x: box.x + flat[i] * box.w, y: box.y + flat[i + 1] * box.h });
    }
    pts.push({ close: true });
    s.addShape('custGeom', { x: box.x, y: box.y, w: box.w, h: box.h, fill: { color: GREY }, points: pts });
  });
}

/** Teal/gold speech bubble used to label a region on the map. */
function mapTag(s, box, color, headline, caption) {
  const w = box.w, h = box.h;
  s.addShape('custGeom', {
    x: box.x, y: box.y, w: w, h: h, fill: { color: color },
    points: [{ x: 0, y: 0 }, { x: w, y: 0 }, { x: w, y: 0.737 * h }, { x: 0.63 * w, y: 0.737 * h },
    { x: 0.5 * w, y: h }, { x: 0.37 * w, y: 0.737 * h }, { x: 0, y: 0.737 * h }, { close: true }],
  });
  txt(s, headline, { x: box.x, y: box.y, w: w, h: 0.337, align: 'center', fontFace: HEAD, fontSize: 14, color: WHITE, wrap: false });
  txt(s, caption, { x: box.x + 0.041, y: box.y + 0.152, w: w - 0.082, h: 0.37, align: 'center', fontFace: BODY, fontSize: 12, color: WHITE, lineSpacingMultiple: 1.5 });
}

const WORLD = [
  [0.996,0.206,0.761,0.138,0.783,0.102,0.759,0.082,0.681,0.144,0.69,0.191,0.672,0.205,0.668,0.13,0.662,0.185,0.593,0.181,0.575,0.226,0.56,0.196,0.583,0.189,0.539,0.156,0.488,0.27,0.515,0.289,0.533,0.205,0.531,0.252,0.553,0.259,0.526,0.306,0.497,0.284,0.447,0.429,0.496,0.378,0.515,0.419,0.508,0.368,0.532,0.431,0.555,0.362,0.587,0.401,0.544,0.409,0.572,0.435,0.563,0.47,0.456,0.44,0.425,0.581,0.435,0.509,0.566,0.509,0.559,0.474,0.61,0.509,0.607,0.473,0.673,0.522,0.685,0.605,0.721,0.52,0.76,0.602,0.77,0.568,0.743,0.527,0.754,0.403,0.828,0.403,0.863,0.267,0.926,0.236,0.905,0.331,0.924,0.259],
  [0.314,0.324,0.302,0.314,0.309,0.304,0.29,0.258,0.279,0.275,0.266,0.241,0.253,0.241,0.258,0.292,0.249,0.334,0.24,0.301,0.212,0.285,0.205,0.257,0.216,0.231,0.208,0.222,0.224,0.221,0.242,0.179,0.232,0.17,0.226,0.195,0.206,0.148,0.204,0.195,0.033,0.155,0.007,0.181,0.022,0.205,0.003,0.21,0.022,0.217,0.008,0.248,0.031,0.27,0.024,0.295,0.048,0.248,0.049,0.264,0.06,0.251,0.099,0.277,0.13,0.347,0.126,0.414,0.148,0.474,0.149,0.395,0.224,0.395,0.233,0.367,0.214,0.364,0.224,0.346,0.247,0.375,0.241,0.392,0.233,0.375,0.23,0.395,0.271,0.395,0.299,0.373,0.289,0.347,0.275,0.357],
  [0.27,0.584,0.257,0.608,0.25,0.602,0.255,0.634,0.246,0.656,0.244,0.684,0.26,0.741,0.275,0.761,0.271,0.855,0.264,0.902,0.268,0.908,0.267,0.935,0.261,0.943,0.265,0.952,0.265,0.981,0.269,0.98,0.271,0.996,0.285,0.954,0.281,0.938,0.289,0.918,0.287,0.902,0.295,0.903,0.296,0.888,0.308,0.885,0.311,0.875,0.307,0.859,0.319,0.861,0.333,0.823,0.334,0.804,0.353,0.785,0.361,0.757,0.361,0.727,0.373,0.695,0.359,0.674,0.332,0.655,0.33,0.66,0.334,0.669,0.325,0.663,0.33,0.649,0.327,0.634,0.324,0.638,0.325,0.628,0.311,0.622,0.294,0.593,0.274,0.589,0.27,0.603],
  [0.436,0.029,0.405,0.041,0.408,0.024,0.396,0.033,0.399,0.025,0.379,0.025,0.409,0.014,0.374,0,0.363,0.012,0.349,0.007,0.356,0.015,0.346,0.012,0.351,0.017,0.346,0.027,0.333,0.017,0.331,0.029,0.301,0.026,0.301,0.036,0.283,0.046,0.29,0.048,0.288,0.062,0.268,0.074,0.285,0.078,0.273,0.087,0.279,0.097,0.299,0.096,0.313,0.112,0.316,0.149,0.326,0.146,0.33,0.161,0.321,0.159,0.331,0.165,0.332,0.191,0.322,0.191,0.327,0.218,0.332,0.217,0.328,0.226,0.349,0.257,0.36,0.207,0.373,0.201,0.381,0.18,0.407,0.166,0.391,0.161,0.41,0.154,0.394,0.132,0.414,0.127,0.417,0.102,0.408,0.089,0.42,0.086,0.41,0.079,0.419,0.051,0.413,0.05],
  [0.591,0.646,0.552,0.646,0.529,0.692,0.504,0.648,0.498,0.646,0.497,0.66,0.505,0.681,0.508,0.68,0.506,0.691,0.51,0.715,0.505,0.742,0.505,0.758,0.512,0.784,0.515,0.816,0.523,0.823,0.523,0.826,0.517,0.824,0.524,0.859,0.546,0.849,0.558,0.82,0.559,0.804,0.566,0.794,0.564,0.767,0.581,0.739,0.577,0.679],
  [0.893,0.806,0.88,0.775,0.874,0.767,0.871,0.743,0.866,0.737,0.864,0.722,0.861,0.752,0.858,0.761,0.843,0.742,0.847,0.727,0.837,0.723,0.83,0.731,0.828,0.744,0.823,0.746,0.819,0.738,0.811,0.757,0.808,0.756,0.803,0.77,0.783,0.788,0.783,0.813,0.789,0.84,0.788,0.861,0.795,0.864,0.8,0.856,0.809,0.857,0.817,0.846,0.833,0.841,0.84,0.846,0.845,0.862,0.851,0.852,0.857,0.881,0.865,0.89,0.872,0.885,0.874,0.891,0.878,0.883,0.883,0.882,0.892,0.842],
  [0.299,0.019,0.274,0.009,0.257,0.009,0.258,0.017,0.249,0.011,0.241,0.017,0.241,0.024,0.233,0.019,0.232,0.026,0.217,0.026,0.226,0.04,0.232,0.031,0.23,0.041,0.244,0.04,0.253,0.031,0.252,0.045,0.241,0.047,0.243,0.056,0.23,0.047,0.23,0.052,0.243,0.066,0.23,0.067,0.227,0.072,0.234,0.069,0.234,0.079,0.238,0.082,0.234,0.086,0.225,0.078,0.229,0.087,0.221,0.092,0.25,0.092,0.242,0.082,0.25,0.082,0.252,0.074,0.259,0.071,0.26,0.066,0.252,0.066,0.252,0.057,0.263,0.056,0.262,0.05,0.271,0.051,0.287,0.032,0.276,0.036,0.279,0.026],
  [0.221,0.135,0.22,0.15,0.227,0.154,0.223,0.159,0.249,0.165,0.248,0.16,0.25,0.159,0.259,0.171,0.258,0.177,0.261,0.175,0.267,0.182,0.269,0.195,0.264,0.205,0.265,0.211,0.255,0.212,0.253,0.22,0.265,0.217,0.272,0.231,0.284,0.241,0.278,0.226,0.288,0.233,0.289,0.226,0.279,0.205,0.284,0.199,0.292,0.211,0.298,0.199,0.28,0.182,0.279,0.174,0.282,0.172,0.278,0.169,0.281,0.164,0.278,0.167,0.278,0.16,0.271,0.163,0.267,0.158,0.269,0.15,0.263,0.155,0.262,0.145,0.257,0.145,0.259,0.14,0.253,0.138,0.252,0.146,0.25,0.148,0.249,0.143,0.245,0.144,0.244,0.133,0.24,0.127,0.232,0.136,0.235,0.139,0.232,0.145,0.235,0.155,0.232,0.156,0.228,0.14,0.232,0.125],
  [0.177,0.531,0.178,0.54,0.189,0.552,0.201,0.562,0.208,0.558,0.214,0.566,0.215,0.571,0.22,0.572,0.222,0.576,0.226,0.576,0.232,0.589,0.232,0.594,0.235,0.597,0.238,0.604,0.242,0.605,0.246,0.61,0.245,0.607,0.249,0.6,0.245,0.604,0.238,0.6,0.235,0.592,0.237,0.569,0.235,0.565,0.222,0.565,0.224,0.545,0.219,0.533,0.218,0.542,0.215,0.547,0.207,0.55,0.202,0.545,0.198,0.531],
  [0.793,0.618,0.791,0.618,0.789,0.624,0.788,0.632,0.786,0.633,0.785,0.632,0.781,0.638,0.778,0.64,0.776,0.644,0.776,0.648,0.771,0.646,0.771,0.654,0.773,0.659,0.774,0.669,0.778,0.671,0.779,0.674,0.782,0.671,0.783,0.672,0.786,0.672,0.786,0.676,0.789,0.675,0.79,0.664,0.792,0.659,0.793,0.651,0.796,0.648,0.792,0.636,0.793,0.63,0.794,0.629,0.796,0.629,0.796,0.625,0.797,0.624,0.795,0.624,0.794,0.622],
  [0.837,0.672,0.837,0.676,0.838,0.675,0.843,0.676,0.843,0.679,0.852,0.686,0.854,0.692,0.855,0.694,0.855,0.7,0.86,0.703,0.861,0.707,0.864,0.706,0.864,0.703,0.866,0.7,0.869,0.697,0.874,0.7,0.878,0.712,0.884,0.715,0.875,0.697,0.875,0.694,0.877,0.692,0.874,0.69,0.871,0.682,0.866,0.676,0.854,0.669,0.852,0.666,0.85,0.666,0.845,0.676,0.842,0.676,0.84,0.671],
  [0.606,0.726,0.604,0.733,0.603,0.734,0.603,0.739,0.598,0.744,0.598,0.747,0.596,0.747,0.594,0.749,0.593,0.749,0.592,0.757,0.592,0.758,0.593,0.77,0.59,0.782,0.59,0.787,0.591,0.792,0.591,0.798,0.592,0.8,0.596,0.801,0.598,0.8,0.6,0.795,0.601,0.788,0.601,0.784,0.606,0.758,0.606,0.744,0.608,0.743,0.608,0.737],
  [0.457,0.273,0.457,0.277,0.456,0.278,0.456,0.285,0.457,0.287,0.457,0.292,0.459,0.293,0.459,0.299,0.458,0.3,0.462,0.299,0.463,0.308,0.463,0.315,0.463,0.316,0.46,0.316,0.46,0.323,0.458,0.325,0.464,0.328,0.464,0.33,0.463,0.333,0.459,0.333,0.459,0.335,0.461,0.335,0.462,0.333,0.473,0.331,0.472,0.33,0.472,0.328,0.475,0.32,0.474,0.319,0.471,0.319,0.471,0.314,0.469,0.313,0.47,0.309,0.467,0.304,0.463,0.294,0.461,0.293,0.461,0.289,0.465,0.284,0.465,0.28,0.463,0.279,0.459,0.28,0.459,0.277,0.461,0.273],
  [0.733,0.625,0.734,0.629,0.739,0.638,0.739,0.64,0.74,0.641,0.742,0.644,0.743,0.651,0.747,0.661,0.748,0.669,0.757,0.686,0.759,0.686,0.76,0.687,0.76,0.686,0.76,0.675,0.761,0.674,0.76,0.67,0.762,0.669,0.761,0.666,0.76,0.67,0.758,0.67,0.757,0.664,0.754,0.659,0.754,0.655,0.754,0.654,0.753,0.654,0.75,0.648,0.748,0.645,0.746,0.644,0.744,0.639,0.737,0.627,0.734,0.627],
];

function slide01(pptx) {
  const s = newSlide(pptx);
  imageBox(s, { x: 0, y: 0, w: 13.333, h: 7.5 });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: TEAL, transparency: 30 } });
  logoMark(s, { x: 3.601, y: 2.714, w: 0.944, h: 0.766 }, SAND);
  txt(s, 'Rihlah', { x: 4.606, y: 2.103, w: 5.037, h: 2.121, fontSize: 120, fontFace: HEAD, color: WHITE, wrap: false });
  txt(s, 'HAJJ & UMRAH AGENCY', { x: 3.892, y: 4.03, w: 5.549, h: 0.522, align: 'center', fontSize: 25, fontFace: HEAD, color: WHITE, charSpacing: 6, wrap: false });
  arabesque(s, { x: -1.473, y: -0.893, w: 2.946, h: 2.946 }, SAND);
  arabesque(s, { x: 11.86, y: -0.893, w: 2.946, h: 2.946 }, SAND);
  arabesque(s, { x: -1.473, y: 5.447, w: 2.946, h: 2.946 }, SAND);
  arabesque(s, { x: 11.86, y: 5.447, w: 2.946, h: 2.946 }, SAND);
  txt(s, 'www.yourwebsite.com`', { x: 4.973, y: 6.743, w: 3.387, h: 0.353, align: 'center', fontSize: 15, color: WHITE, bold: true, fontFace: BODY });
}

function slide02(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  s.addShape(pptx.ShapeType.rect, { x: 8.659, y: 0, w: 4.675, h: 7.5, fill: { color: TEAL } });
  txt(s, [{ text: 'Welcome to Our ', options: { breakLine: true } }, { text: 'Hajj & Umrah Agency', options: { color: TEAL } }], { x: 1.021, y: 1.208, w: 6.424, h: 1.479, fontSize: 45, fontFace: HEAD, color: WHITE });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence in this spot, which was created for the bliss of souls like mine. I am so happy, my dear friend, so absorbed in the exquisite sense of mere.', { x: 1.021, y: 2.844, w: 6.229, h: 1.279, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  s.addShape(pptx.ShapeType.rect, { x: 1.11, y: 4.505, w: 1.207, h: 0.38, fill: { color: TEAL } });
  txt(s, 'Read More', { x: 1.214, y: 4.543, w: 0.998, h: 0.303, align: 'center', fontSize: 12, fontFace: HEAD, color: WHITE, wrap: false });
  s.addShape(pptx.ShapeType.line, { x: 1.11, y: 5.416, w: 6.229, h: 0, line: { color: SAND, width: 1.5 } });
  txt(s, 'Hasan Nasser', { x: 1.021, y: 5.534, w: 1.873, h: 0.438, fontSize: 20, fontFace: HEAD, color: TEAL, wrap: false });
  txt(s, 'Founder & CEO', { x: 1.021, y: 5.82, w: 1.924, h: 0.37, align: 'justify', fontSize: 12, bold: true, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
}

function slide03(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  txt(s, 'List of Content', { x: 1.021, y: 1.208, w: 7.451, h: 0.797, fontSize: 45, fontFace: HEAD, color: WHITE });
  txt(s, 'Suitable for all categories.', { x: 1.888, y: 2.606, w: 2.2, h: 0.37, fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'About Agency', { x: 1.888, y: 2.279, w: 1.771, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  star8(s, { x: 1.021, y: 2.252, w: 0.806, h: 0.806 }, TEAL);
  txt(s, '01', { x: 1.082, y: 2.368, w: 0.683, h: 0.572, align: 'center', fontSize: 28, fontFace: HEAD, color: WHITE });
  txt(s, 'Suitable for all categories.', { x: 1.888, y: 3.684, w: 2.2, h: 0.37, fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Vision & Mission', { x: 1.888, y: 3.357, w: 2.029, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  star8(s, { x: 1.021, y: 3.33, w: 0.806, h: 0.806 }, TEAL);
  txt(s, '02', { x: 1.082, y: 3.446, w: 0.683, h: 0.572, align: 'center', fontSize: 28, fontFace: HEAD, color: WHITE });
  txt(s, 'Suitable for all categories.', { x: 1.888, y: 5.84, w: 2.2, h: 0.37, fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Umrah Services', { x: 1.888, y: 5.514, w: 1.967, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  star8(s, { x: 1.021, y: 5.486, w: 0.806, h: 0.806 }, TEAL);
  txt(s, '04', { x: 1.082, y: 5.602, w: 0.683, h: 0.572, align: 'center', fontSize: 28, fontFace: HEAD, color: WHITE });
  txt(s, 'Suitable for all categories.', { x: 1.888, y: 4.762, w: 2.2, h: 0.37, fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Hajj Services', { x: 1.888, y: 4.435, w: 1.641, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  star8(s, { x: 1.021, y: 4.408, w: 0.806, h: 0.806 }, TEAL);
  txt(s, '03', { x: 1.082, y: 4.524, w: 0.683, h: 0.572, align: 'center', fontSize: 28, fontFace: HEAD, color: WHITE });
  txt(s, 'Suitable for all categories.', { x: 6.099, y: 2.606, w: 2.2, h: 0.37, fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Achievement', { x: 6.099, y: 2.279, w: 1.682, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  star8(s, { x: 5.232, y: 2.252, w: 0.806, h: 0.806 }, TEAL);
  txt(s, '05', { x: 5.293, y: 2.368, w: 0.683, h: 0.572, align: 'center', fontSize: 28, fontFace: HEAD, color: WHITE });
  txt(s, 'Suitable for all categories.', { x: 6.099, y: 3.684, w: 2.2, h: 0.37, fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'The Team', { x: 6.099, y: 3.357, w: 1.326, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  star8(s, { x: 5.232, y: 3.33, w: 0.806, h: 0.806 }, TEAL);
  txt(s, '06', { x: 5.293, y: 3.446, w: 0.683, h: 0.572, align: 'center', fontSize: 28, fontFace: HEAD, color: WHITE });
  txt(s, 'Suitable for all categories.', { x: 6.099, y: 5.84, w: 2.2, h: 0.37, fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Contact', { x: 6.099, y: 5.514, w: 1.075, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  star8(s, { x: 5.232, y: 5.486, w: 0.806, h: 0.806 }, TEAL);
  txt(s, '08', { x: 5.293, y: 5.602, w: 0.683, h: 0.572, align: 'center', fontSize: 28, fontFace: HEAD, color: WHITE });
  txt(s, 'Suitable for all categories.', { x: 6.099, y: 4.762, w: 2.2, h: 0.37, fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Infographics', { x: 6.099, y: 4.435, w: 1.62, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  star8(s, { x: 5.232, y: 4.408, w: 0.806, h: 0.806 }, TEAL);
  txt(s, '07', { x: 5.293, y: 4.524, w: 0.683, h: 0.572, align: 'center', fontSize: 28, fontFace: HEAD, color: WHITE });
  imageBox(s, { x: 9.014, y: 0, w: 4.319, h: 7.5 });
}

function slide04(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  txt(s, [{ text: 'Agency ' }, { text: 'Overview', options: { color: TEAL } }], { x: 1.021, y: 1.209, w: 6.424, h: 0.797, fontSize: 45, fontFace: HEAD, color: WHITE });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart.', { x: 1.021, y: 2.106, w: 5.382, h: 0.673, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  star8(s, { x: 1.021, y: 5.028, w: 0.396, h: 0.396 }, TEAL);
  txt(s, 'Mornings of spring which I enjoy with my whole heart.', { x: 1.021, y: 5.808, w: 2.2, h: 0.673, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 1.021, y: 5.481, w: 1.499, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  star8(s, { x: 4.051, y: 5.028, w: 0.396, h: 0.396 }, TEAL);
  txt(s, 'Mornings of spring which I enjoy with my whole heart.', { x: 4.051, y: 5.808, w: 2.2, h: 0.673, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 4.051, y: 5.481, w: 1.499, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  star8(s, { x: 7.082, y: 5.028, w: 0.396, h: 0.396 }, TEAL);
  txt(s, 'Mornings of spring which I enjoy with my whole heart.', { x: 7.082, y: 5.808, w: 2.2, h: 0.673, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 7.082, y: 5.481, w: 1.499, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  star8(s, { x: 10.112, y: 5.028, w: 0.396, h: 0.396 }, TEAL);
  txt(s, 'Mornings of spring which I enjoy with my whole heart.', { x: 10.112, y: 5.808, w: 2.2, h: 0.673, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 10.112, y: 5.481, w: 1.499, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  s.addShape(pptx.ShapeType.line, { x: 1.572, y: 5.226, w: 2.324, h: 0, line: { color: SAND, width: 1.5 } });
  s.addShape(pptx.ShapeType.line, { x: 4.603, y: 5.226, w: 2.324, h: 0, line: { color: SAND, width: 1.5 } });
  s.addShape(pptx.ShapeType.line, { x: 7.633, y: 5.226, w: 2.324, h: 0, line: { color: SAND, width: 1.5 } });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence in this spot.', { x: 1.021, y: 3.403, w: 5.382, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 1.021, y: 3.076, w: 1.499, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  imageBox(s, { x: 7.181, y: 0, w: 5.132, h: 4.379 });
}

function slide05(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  txt(s, [{ text: 'Vision & ' }, { text: 'Mission', options: { color: TEAL } }], { x: 6.569, y: 1.209, w: 5.743, h: 0.797, fontSize: 45, fontFace: HEAD, color: WHITE });
  txt(s, 'Sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence in this spot, which was created for the bliss of souls like mine. I am so happy, my dear friend, so absorbed in the exquisite sense of mere.', { x: 6.569, y: 2.18, w: 5.549, h: 1.279, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  s.addShape(pptx.ShapeType.rect, { x: 5.722, y: 4.152, w: 6.733, h: 2.612, fill: { color: TEAL } });
  icon(s, { x: 6.218, y: 4.508, w: 0.551, h: 0.551 }, SAND, 'starBadge');
  icon(s, { x: 9.672, y: 4.508, w: 0.551, h: 0.551 }, SAND, 'moonBadge');
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings.', { x: 6.107, y: 5.431, w: 2.511, h: 0.976, align: 'justify', fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 6.107, y: 5.105, w: 1.499, h: 0.404, fontFace: HEAD, color: WHITE, fontSize: 18, wrap: false });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings.', { x: 9.56, y: 5.431, w: 2.511, h: 0.976, align: 'justify', fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 9.561, y: 5.105, w: 1.499, h: 0.404, fontFace: HEAD, color: WHITE, fontSize: 18, wrap: false });
  imageBox(s, { x: 1.021, y: 1.209, w: 5.202, h: 2.612 });
  imageBox(s, { x: 1.021, y: 4.152, w: 4.452, h: 3.348 });
}

function slide06(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  txt(s, [{ text: 'Why Choose', options: { breakLine: true } }, { text: 'Our Agency?', options: { color: TEAL } }], { x: 1.021, y: 1.209, w: 4.045, h: 1.509, fontSize: 45, fontFace: HEAD, color: WHITE });
  txt(s, 'I am so happy, my dear friend, so absorbed in the exquisite sense of mere tranquil existence, that I neglect my talents. I should be incapable.', { x: 1.021, y: 2.774, w: 3.868, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  s.addShape(pptx.ShapeType.rect, { x: 1.021, y: 4.185, w: 3.868, h: 3.315, fill: { color: TEAL } });
  txt(s, 'Possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel.', { x: 1.286, y: 5.463, w: 3.338, h: 0.976, align: 'justify', fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 1.286, y: 5.136, w: 1.499, h: 0.404, fontFace: HEAD, color: WHITE, fontSize: 18, wrap: false });
  icon(s, { x: 1.397, y: 4.494, w: 0.661, h: 0.551 }, SAND, 'mosque');
  txt(s, 'Possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart.', { x: 5.749, y: 5.463, w: 2.749, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 5.749, y: 5.136, w: 1.499, h: 0.404, fontFace: HEAD, fontSize: 18, color: WHITE, wrap: false });
  icon(s, { x: 5.86, y: 4.494, w: 0.661, h: 0.551 }, SAND, 'mosque');
  txt(s, 'Possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart.', { x: 9.317, y: 5.463, w: 2.749, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 9.317, y: 5.136, w: 1.499, h: 0.404, fontFace: HEAD, fontSize: 18, color: WHITE, wrap: false });
  icon(s, { x: 9.428, y: 4.494, w: 0.661, h: 0.551 }, SAND, 'mosque');
  imageBox(s, { x: 5.502, y: 0, w: 3.244, h: 3.978 });
  imageBox(s, { x: 9.069, y: 0, w: 3.244, h: 3.978 });
}

function slide07(pptx) {
  const s = newSlide(pptx);
  imageBox(s, { x: 0, y: 0, w: 13.333, h: 7.5 });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: TEAL, transparency: 30 } });
  icon(s, { x: 5.69, y: 2.371, w: 1.727, h: 1.727 }, SAND, 'starBadge');
  txt(s, [{ text: 'Hajj & Umrah', options: { breakLine: true } }, { text: 'Services' }], { x: 7.635, y: 2.275, w: 4.735, h: 1.919, fontSize: 54, fontFace: HEAD, color: WHITE, wrap: false });
  txt(s, 'Sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence in this spot, which was created for the bliss of souls like mine..', { x: 7.635, y: 4.249, w: 4.735, h: 0.976, align: 'justify', fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  arabesque(s, { x: -1.473, y: -0.893, w: 2.946, h: 2.946 }, SAND);
  arabesque(s, { x: 11.86, y: -0.893, w: 2.946, h: 2.946 }, SAND);
  arabesque(s, { x: -1.473, y: 5.447, w: 2.946, h: 2.946 }, SAND);
  arabesque(s, { x: 11.86, y: 5.447, w: 2.946, h: 2.946 }, SAND);
  imageBox(s, { x: 0, y: 2.371, w: 5.472, h: 2.854 });
}

function slide08(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  txt(s, [{ text: 'Definition of Hajj ' }, { text: '& Its Procedures', options: { color: TEAL } }], { x: 4.15, y: 1.209, w: 5.278, h: 1.509, fontSize: 45, fontFace: HEAD, color: WHITE });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence .', { x: 4.151, y: 2.774, w: 4.849, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Possession of my entire soul, like these sweet mornings of spring which I enjoy.', { x: 4.151, y: 5.316, w: 2.43, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 4.151, y: 4.989, w: 1.499, h: 0.404, fontFace: HEAD, fontSize: 18, color: WHITE, wrap: false });
  txt(s, 'Possession of my entire soul, like these sweet mornings of spring which I enjoy.', { x: 7.11, y: 5.316, w: 2.43, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 7.11, y: 4.989, w: 1.499, h: 0.404, fontFace: HEAD, fontSize: 18, color: WHITE, wrap: false });
  icon(s, { x: 7.221, y: 4.493, w: 0.485, h: 0.404 }, SAND, 'mosque');
  txt(s, 'Possession of my entire soul, like these sweet mornings of spring which I enjoy.', { x: 10.068, y: 5.316, w: 2.43, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 10.068, y: 4.989, w: 1.499, h: 0.404, fontFace: HEAD, fontSize: 18, color: WHITE, wrap: false });
  icon(s, { x: 4.262, y: 4.493, w: 0.364, h: 0.404 }, SAND, 'calendar');
  icon(s, { x: 10.179, y: 4.493, w: 0.404, h: 0.404 }, SAND, 'avatar');
  imageBox(s, { x: 0, y: 1.209, w: 3.523, h: 5.263 });
  imageBox(s, { x: 9.628, y: 0, w: 2.87, h: 3.75 });
}

function slide09(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  txt(s, [{ text: 'Our ' }, { text: 'Hajj Packages', options: { color: TEAL } }], { x: 3.455, y: 1.209, w: 6.424, h: 0.797, fontSize: 45, fontFace: HEAD, color: WHITE });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings.', { x: 0.835, y: 5.761, w: 3.586, h: 0.673, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  s.addShape(pptx.ShapeType.rect, { x: 0.835, y: 5.2, w: 3.585, h: 0.5, fill: { color: TEAL } });
  txt(s, 'Economy', { x: 0.946, y: 5.248, w: 1.24, h: 0.404, fontFace: HEAD, color: WHITE, fontSize: 18, wrap: false });
  icon(s, { x: 3.975, y: 5.325, w: 0.335, h: 0.251 }, WHITE, 'arrow');
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings.', { x: 4.874, y: 5.761, w: 3.586, h: 0.673, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  s.addShape(pptx.ShapeType.rect, { x: 4.874, y: 5.2, w: 3.585, h: 0.5, fill: { color: TEAL } });
  txt(s, 'Standard', { x: 4.984, y: 5.248, w: 1.236, h: 0.404, fontFace: HEAD, color: WHITE, fontSize: 18, wrap: false });
  icon(s, { x: 8.014, y: 5.325, w: 0.335, h: 0.251 }, WHITE, 'arrow');
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings.', { x: 8.913, y: 5.761, w: 3.586, h: 0.673, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  s.addShape(pptx.ShapeType.rect, { x: 8.913, y: 5.2, w: 3.585, h: 0.5, fill: { color: TEAL } });
  txt(s, 'VIP', { x: 9.023, y: 5.248, w: 0.579, h: 0.404, fontFace: HEAD, color: WHITE, fontSize: 18, wrap: false });
  icon(s, { x: 12.053, y: 5.325, w: 0.335, h: 0.251 }, WHITE, 'arrow');
  imageBox(s, { x: 0.835, y: 2.349, w: 3.586, h: 2.665 });
  imageBox(s, { x: 4.874, y: 2.349, w: 3.586, h: 2.665 });
  imageBox(s, { x: 8.913, y: 2.349, w: 3.586, h: 2.665 });
}

function slide10(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  txt(s, [{ text: 'Package ' }, { text: 'Inclusions', options: { color: TEAL } }], { x: 1.016, y: 1.208, w: 5.806, h: 0.797, fontSize: 45, fontFace: HEAD, color: WHITE });
  icon(s, { x: 10.058, y: 1.366, w: 0.621, h: 0.424 }, SAND, 'plane');
  icon(s, { x: 7.269, y: 3.691, w: 0.565, h: 0.424 }, SAND, 'bed');
  icon(s, { x: 10.058, y: 3.602, w: 0.513, h: 0.513 }, SAND, 'person');
  txt(s, 'Strikes the upper surface of the impenetrable foliage of my trees, and but a few stray', { x: 7.158, y: 2.235, w: 2.334, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 7.158, y: 1.908, w: 1.499, h: 0.404, fontFace: HEAD, color: WHITE, fontSize: 18, wrap: false });
  icon(s, { x: 7.269, y: 1.277, w: 0.584, h: 0.513 }, SAND, 'cloche');
  txt(s, 'When, while the lovely valley teems with vapour around me, and the meridian sun strikes the upper surface of the impenetrable foliage of my trees.', { x: 2.097, y: 2.207, w: 4.569, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  icon(s, { x: 1.015, y: 2.246, w: 0.899, h: 0.899 }, SAND, 'starBadge');
  txt(s, 'Strikes the upper surface of the impenetrable foliage of my trees, and but a few stray', { x: 9.984, y: 2.235, w: 2.334, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 9.984, y: 1.908, w: 1.499, h: 0.404, fontFace: HEAD, color: WHITE, fontSize: 18, wrap: false });
  txt(s, 'Strikes the upper surface of the impenetrable foliage of my trees, and but a few stray', { x: 7.158, y: 4.559, w: 2.334, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 7.158, y: 4.233, w: 1.499, h: 0.404, fontFace: HEAD, color: WHITE, fontSize: 18, wrap: false });
  txt(s, 'Strikes the upper surface of the impenetrable foliage of my trees, and but a few stray', { x: 9.984, y: 4.559, w: 2.334, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 9.984, y: 4.233, w: 1.499, h: 0.404, fontFace: HEAD, color: WHITE, fontSize: 18, wrap: false });
  s.addShape(pptx.ShapeType.rect, { x: 7.158, y: 5.985, w: 5.16, h: 0.899, fill: { color: TEAL } });
  txt(s, 'Possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone.', { x: 7.368, y: 6.098, w: 4.74, h: 0.673, align: 'justify', fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  imageBox(s, { x: 0.861, y: 3.602, w: 5.805, h: 3.898 });
}

function slide11(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  star8(s, { x: 9.846, y: 2.36, w: 0.396, h: 0.396 }, TEAL);
  txt(s, 'Day 04', { x: 10.241, y: 2.38, w: 0.947, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  s.addShape(pptx.ShapeType.line, { x: 10.044, y: 3.034, w: 0, h: 4.5, line: { color: SAND, width: 1 } });
  txt(s, 'Suitable for all categories', { x: 10.112, y: 3.174, w: 2.2, h: 0.37, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 10.112, y: 2.902, w: 1.215, h: 0.337, fontSize: 14, bold: true, fontFace: BODY, color: WHITE, wrap: false });
  s.addShape(pptx.ShapeType.ellipse, { x: 9.973, y: 3, w: 0.141, h: 0.141, fill: { color: SAND } });
  txt(s, 'Suitable for all categories', { x: 10.112, y: 4.565, w: 2.2, h: 0.37, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 10.112, y: 4.293, w: 1.215, h: 0.337, fontSize: 14, bold: true, fontFace: BODY, color: WHITE, wrap: false });
  s.addShape(pptx.ShapeType.ellipse, { x: 9.973, y: 4.391, w: 0.141, h: 0.141, fill: { color: SAND } });
  txt(s, 'Suitable for all categories', { x: 10.112, y: 5.955, w: 2.2, h: 0.37, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 10.112, y: 5.683, w: 1.215, h: 0.337, fontSize: 14, bold: true, fontFace: BODY, color: WHITE, wrap: false });
  s.addShape(pptx.ShapeType.ellipse, { x: 9.973, y: 5.781, w: 0.141, h: 0.141, fill: { color: SAND } });
  star8(s, { x: 6.904, y: 2.325, w: 0.396, h: 0.396 }, TEAL);
  txt(s, 'Day 03', { x: 7.3, y: 2.346, w: 0.947, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  s.addShape(pptx.ShapeType.line, { x: 7.102, y: 3, w: 0, h: 4.5, line: { color: SAND, width: 1 } });
  txt(s, 'Suitable for all categories', { x: 7.17, y: 3.14, w: 2.2, h: 0.37, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 7.17, y: 2.868, w: 1.215, h: 0.337, fontSize: 14, bold: true, fontFace: BODY, color: WHITE, wrap: false });
  s.addShape(pptx.ShapeType.ellipse, { x: 7.031, y: 2.966, w: 0.141, h: 0.141, fill: { color: SAND } });
  txt(s, 'Suitable for all categories', { x: 7.17, y: 4.53, w: 2.2, h: 0.37, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 7.17, y: 4.259, w: 1.215, h: 0.337, fontSize: 14, bold: true, fontFace: BODY, color: WHITE, wrap: false });
  s.addShape(pptx.ShapeType.ellipse, { x: 7.031, y: 4.356, w: 0.141, h: 0.141, fill: { color: SAND } });
  txt(s, 'Suitable for all categories', { x: 7.17, y: 5.921, w: 2.2, h: 0.37, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 7.17, y: 5.649, w: 1.215, h: 0.337, fontSize: 14, bold: true, fontFace: BODY, color: WHITE, wrap: false });
  s.addShape(pptx.ShapeType.ellipse, { x: 7.031, y: 5.747, w: 0.141, h: 0.141, fill: { color: SAND } });
  star8(s, { x: 3.962, y: 2.346, w: 0.396, h: 0.396 }, TEAL);
  txt(s, 'Day 02', { x: 4.358, y: 2.366, w: 0.947, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  s.addShape(pptx.ShapeType.line, { x: 4.16, y: 3.02, w: 0, h: 4.5, line: { color: SAND, width: 1 } });
  txt(s, 'Suitable for all categories', { x: 4.228, y: 3.16, w: 2.2, h: 0.37, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 4.229, y: 2.888, w: 1.215, h: 0.337, fontSize: 14, bold: true, fontFace: BODY, color: WHITE, wrap: false });
  s.addShape(pptx.ShapeType.ellipse, { x: 4.09, y: 2.986, w: 0.141, h: 0.141, fill: { color: SAND } });
  txt(s, 'Suitable for all categories', { x: 4.228, y: 4.551, w: 2.2, h: 0.37, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 4.229, y: 4.279, w: 1.215, h: 0.337, fontSize: 14, bold: true, fontFace: BODY, color: WHITE, wrap: false });
  s.addShape(pptx.ShapeType.ellipse, { x: 4.09, y: 4.376, w: 0.141, h: 0.141, fill: { color: SAND } });
  txt(s, 'Suitable for all categories', { x: 4.228, y: 5.941, w: 2.2, h: 0.37, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 4.229, y: 5.669, w: 1.215, h: 0.337, fontSize: 14, bold: true, fontFace: BODY, color: WHITE, wrap: false });
  s.addShape(pptx.ShapeType.ellipse, { x: 4.09, y: 5.767, w: 0.141, h: 0.141, fill: { color: SAND } });
  txt(s, [{ text: 'Itinerary ' }, { text: 'Example', options: { color: TEAL } }], { x: 1.021, y: 1.209, w: 7.452, h: 0.797, fontSize: 45, fontFace: HEAD, color: WHITE });
  star8(s, { x: 1.021, y: 2.325, w: 0.396, h: 0.396 }, TEAL);
  txt(s, 'Day 01', { x: 1.417, y: 2.346, w: 0.9, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  s.addShape(pptx.ShapeType.line, { x: 1.219, y: 3, w: 0, h: 4.5, line: { color: SAND, width: 1 } });
  txt(s, 'Suitable for all categories', { x: 1.287, y: 3.14, w: 2.2, h: 0.37, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 1.287, y: 2.868, w: 1.215, h: 0.337, fontSize: 14, bold: true, fontFace: BODY, color: WHITE, wrap: false });
  s.addShape(pptx.ShapeType.ellipse, { x: 1.148, y: 2.966, w: 0.141, h: 0.141, fill: { color: SAND } });
  txt(s, 'Suitable for all categories', { x: 1.287, y: 4.53, w: 2.2, h: 0.37, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 1.287, y: 4.259, w: 1.215, h: 0.337, fontSize: 14, bold: true, fontFace: BODY, color: WHITE, wrap: false });
  s.addShape(pptx.ShapeType.ellipse, { x: 1.148, y: 4.356, w: 0.141, h: 0.141, fill: { color: SAND } });
  txt(s, 'Suitable for all categories', { x: 1.287, y: 5.921, w: 2.2, h: 0.37, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 1.287, y: 5.649, w: 1.215, h: 0.337, fontSize: 14, bold: true, fontFace: BODY, color: WHITE, wrap: false });
  s.addShape(pptx.ShapeType.ellipse, { x: 1.148, y: 5.747, w: 0.141, h: 0.141, fill: { color: SAND } });
}

function slide12(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  txt(s, [{ text: 'Facilities & ' }, { text: 'Support', options: { color: TEAL } }], { x: 6.62, y: 1.208, w: 5.693, h: 0.797, fontSize: 45, fontFace: HEAD, color: WHITE });
  icon(s, { x: 3.894, y: 4.406, w: 0.403, h: 0.478 }, SAND, 'bus');
  icon(s, { x: 6.973, y: 4.38, w: 0.403, h: 0.503 }, SAND, 'shield');
  txt(s, 'Strikes the upper surface of the impenetrable foliage of my trees, and but a few stray', { x: 3.82, y: 5.288, w: 2.334, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 3.821, y: 4.961, w: 1.499, h: 0.404, fontFace: HEAD, color: WHITE, fontSize: 18, wrap: false });
  txt(s, 'Strikes the upper surface of the impenetrable foliage of my trees, and but a few stray', { x: 6.899, y: 5.288, w: 2.334, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 6.9, y: 4.961, w: 1.499, h: 0.404, fontFace: HEAD, color: WHITE, fontSize: 18, wrap: false });
  icon(s, { x: 10.052, y: 4.38, w: 0.503, h: 0.503 }, SAND, 'person');
  txt(s, 'Strikes the upper surface of the impenetrable foliage of my trees, and but a few stray', { x: 9.978, y: 5.288, w: 2.334, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 9.979, y: 4.961, w: 1.499, h: 0.404, fontFace: HEAD, color: WHITE, fontSize: 18, wrap: false });
  s.addShape(pptx.ShapeType.rect, { x: 6.667, y: 2.22, w: 5.646, h: 1.53, fill: { color: TEAL } });
  txt(s, 'I am so happy, my dear friend, so absorbed in the exquisite sense of mere tranquil existence, that I neglect my talents. I should be incapable of drawing a single stroke at the present moment.', { x: 6.974, y: 2.487, w: 5.031, h: 0.976, align: 'justify', fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  imageBox(s, { x: 1.021, y: 1.209, w: 2.527, h: 2.527 });
  imageBox(s, { x: 3.82, y: 1.209, w: 2.527, h: 2.527 });
  imageBox(s, { x: 1.021, y: 4.059, w: 2.527, h: 2.527 });
}

function slide13(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  s.addShape(pptx.ShapeType.rect, { x: 1.021, y: 4.279, w: 5.418, h: 2.291, fill: { color: WHITE } });
  s.addShape(pptx.ShapeType.rect, { x: 6.894, y: 4.279, w: 5.418, h: 2.291, fill: { color: WHITE } });
  txt(s, [{ text: 'Past Experience & ' }, { text: 'Success Stories', options: { color: TEAL } }], { x: 1.021, y: 1.209, w: 6.424, h: 1.479, fontSize: 45, fontFace: HEAD, color: WHITE });
  txt(s, '“When, while the lovely valley teems with vapour around me, and the meridian sun strikes the upper surface of the”', { x: 2.84, y: 4.601, w: 3.241, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: SLATE, italic: true, lineSpacingMultiple: 1.5 });
  txt(s, 'Ahmed Murad', { x: 2.84, y: 5.844, w: 1.794, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  txt(s, '“When, while the lovely valley teems with vapour around me, and the meridian sun strikes the upper surface of the”', { x: 8.775, y: 4.601, w: 3.241, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: SLATE, italic: true, lineSpacingMultiple: 1.5 });
  txt(s, 'Miria Sasha', { x: 8.775, y: 5.844, w: 1.492, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  txt(s, 'I am so happy, my dear friend, so absorbed in the exquisite sense of mere tranquil existence, that I neglect my talents. I should be incapable of drawing a single stroke at the present moment.', { x: 1.021, y: 2.731, w: 5.031, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  imageBox(s, { x: 1.221, y: 4.546, w: 1.384, h: 1.756 });
  imageBox(s, { x: 7.095, y: 4.546, w: 1.384, h: 1.756 });
  imageBox(s, { x: 6.894, y: 0, w: 2.571, h: 3.75 });
  imageBox(s, { x: 9.741, y: 0, w: 2.571, h: 3.75 });
}

function slide14(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  txt(s, [{ text: 'What is ' }, { text: 'Umrah?', options: { color: TEAL } }], { x: 6.667, y: 1.209, w: 5.646, h: 0.797, fontSize: 45, fontFace: HEAD, color: WHITE });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence in this spot.', { x: 6.667, y: 2.522, w: 5.646, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Definition, difference from Hajj', { x: 6.667, y: 2.195, w: 3.701, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  s.addShape(pptx.ShapeType.rect, { x: 6.773, y: 3.75, w: 1.207, h: 0.38, fill: { color: TEAL } });
  txt(s, 'Read More', { x: 6.877, y: 3.789, w: 0.998, h: 0.303, align: 'center', fontSize: 12, fontFace: HEAD, color: WHITE, wrap: false });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 4.965, w: 6.337, h: 1.53, fill: { color: TEAL } });
  txt(s, 'I am so happy, my dear friend, so absorbed in the exquisite sense of mere tranquil existence, that I neglect my talents. I should be incapable of drawing a single stroke at the present moment.', { x: 0.653, y: 5.242, w: 5.031, h: 0.976, align: 'justify', fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  imageBox(s, { x: 0, y: 1.033, w: 3.367, h: 3.49 });
  imageBox(s, { x: 3.696, y: 1.064, w: 2.641, h: 3.49 });
  imageBox(s, { x: 6.666, y: 4.965, w: 5.646, h: 2.535 });
}

function slide15(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  txt(s, [{ text: 'Year-Round ', options: { breakLine: true } }, { text: 'Umrah Packages', options: { color: TEAL } }], { x: 1.021, y: 1.209, w: 5.045, h: 1.479, fontSize: 45, fontFace: HEAD, color: WHITE });
  s.addShape(pptx.ShapeType.rect, { x: 9.453, y: 1.209, w: 2.86, h: 4.285, fill: { color: TEAL } });
  s.addShape(pptx.ShapeType.rect, { x: 6.267, y: 1.209, w: 2.86, h: 4.285, fill: { color: WHITE } });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence .', { x: 1.021, y: 2.774, w: 4.779, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'I am so happy, my dear friend, so absorbed in the exquisite sense of mere tranquil existence, that I neglect my talents. I should be incapable of drawing .', { x: 6.267, y: 5.846, w: 6.045, h: 0.673, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  star8(s, { x: 6.497, y: 1.452, w: 0.396, h: 0.396 }, TEAL);
  txt(s, 'STANDARD', { x: 6.892, y: 1.47, w: 1.469, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  txt(s, 'Suitable for all categories.', { x: 6.713, y: 2.894, w: 2.125, h: 0.371, fontSize: 12, color: SLATE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  s.addShape(pptx.ShapeType.ellipse, { x: 6.557, y: 3.031, w: 0.156, h: 0.156, fill: { color: SAND } });
  txt(s, 'Mornings of spring which I enjoy with my whole heart.', { x: 6.557, y: 2.011, w: 2.281, h: 0.673, align: 'justify', fontSize: 12, color: SLATE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  txt(s, 'Suitable for all categories.', { x: 6.713, y: 3.348, w: 2.125, h: 0.371, fontSize: 12, color: SLATE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  s.addShape(pptx.ShapeType.ellipse, { x: 6.557, y: 3.485, w: 0.156, h: 0.156, fill: { color: SAND } });
  txt(s, 'Suitable for all categories.', { x: 6.713, y: 3.802, w: 2.125, h: 0.371, fontSize: 12, color: SLATE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  s.addShape(pptx.ShapeType.ellipse, { x: 6.557, y: 3.939, w: 0.156, h: 0.156, fill: { color: SAND } });
  s.addShape(pptx.ShapeType.rect, { x: 6.557, y: 4.871, w: 2.281, h: 0.38, fill: { color: TEAL } });
  txt(s, 'See More', { x: 7.256, y: 4.909, w: 0.882, h: 0.303, align: 'center', fontSize: 12, fontFace: HEAD, color: WHITE, wrap: false });
  star8(s, { x: 9.712, y: 1.452, w: 0.396, h: 0.396 }, SAND);
  txt(s, 'VIP', { x: 10.108, y: 1.47, w: 0.579, h: 0.404, fontFace: HEAD, color: WHITE, fontSize: 18, wrap: false });
  txt(s, 'Suitable for all categories.', { x: 9.928, y: 2.894, w: 2.125, h: 0.371, fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  s.addShape(pptx.ShapeType.ellipse, { x: 9.772, y: 3.031, w: 0.156, h: 0.156, fill: { color: SAND } });
  txt(s, 'Mornings of spring which I enjoy with my whole heart.', { x: 9.772, y: 2.011, w: 2.281, h: 0.673, align: 'justify', fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  txt(s, 'Suitable for all categories.', { x: 9.928, y: 3.333, w: 2.125, h: 0.371, fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  s.addShape(pptx.ShapeType.ellipse, { x: 9.772, y: 3.47, w: 0.156, h: 0.156, fill: { color: SAND } });
  txt(s, 'Suitable for all categories.', { x: 9.928, y: 3.773, w: 2.125, h: 0.371, fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  s.addShape(pptx.ShapeType.ellipse, { x: 9.772, y: 3.91, w: 0.156, h: 0.156, fill: { color: SAND } });
  s.addShape(pptx.ShapeType.rect, { x: 9.772, y: 4.871, w: 2.281, h: 0.38, fill: { color: WHITE } });
  txt(s, 'See More', { x: 10.472, y: 4.909, w: 0.882, h: 0.303, align: 'center', fontSize: 12, fontFace: HEAD, color: TEAL, wrap: false });
  txt(s, 'Suitable for all categories.', { x: 9.928, y: 4.212, w: 2.125, h: 0.371, fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  s.addShape(pptx.ShapeType.ellipse, { x: 9.772, y: 4.349, w: 0.156, h: 0.156, fill: { color: SAND } });
  imageBox(s, { x: 1.021, y: 4.067, w: 4.779, h: 2.611 });
}

function slide16(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  txt(s, [{ text: 'Umrah ', options: { breakLine: true } }, { text: 'Plus Tours', options: { color: TEAL } }], { x: 8.36, y: 1.209, w: 4.014, h: 1.479, fontSize: 45, fontFace: HEAD, color: WHITE });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone.', { x: 8.36, y: 2.656, w: 4.014, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  s.addShape(pptx.ShapeType.rect, { x: 8.447, y: 3.913, w: 1.207, h: 0.38, fill: { color: TEAL } });
  txt(s, 'Read More', { x: 8.551, y: 3.951, w: 0.998, h: 0.303, align: 'center', fontSize: 12, fontFace: HEAD, color: WHITE, wrap: false });
  txt(s, 'Mornings of spring which I enjoy with my whole heart.', { x: 1.021, y: 5.808, w: 2.2, h: 0.673, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 1.021, y: 5.481, w: 1.499, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  icon(s, { x: 1.125, y: 5.028, w: 0.475, h: 0.396 }, SAND, 'mosque');
  txt(s, 'Mornings of spring which I enjoy with my whole heart.', { x: 4.072, y: 5.808, w: 2.2, h: 0.673, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 4.072, y: 5.481, w: 1.499, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  icon(s, { x: 4.176, y: 5.028, w: 0.475, h: 0.396 }, SAND, 'mosque');
  txt(s, 'Mornings of spring which I enjoy with my whole heart.', { x: 7.123, y: 5.808, w: 2.2, h: 0.673, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 7.123, y: 5.481, w: 1.499, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  icon(s, { x: 7.227, y: 5.028, w: 0.475, h: 0.396 }, SAND, 'mosque');
  txt(s, 'Mornings of spring which I enjoy with my whole heart.', { x: 10.174, y: 5.808, w: 2.2, h: 0.673, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 10.174, y: 5.481, w: 1.499, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  icon(s, { x: 10.278, y: 5.028, w: 0.475, h: 0.396 }, SAND, 'mosque');
  imageBox(s, { x: 0, y: 1.033, w: 2.439, h: 3.49 });
  imageBox(s, { x: 2.787, y: 1.064, w: 2.439, h: 3.49 });
  imageBox(s, { x: 5.574, y: 1.064, w: 2.439, h: 3.49 });
}

function slide17(pptx) {
  const s = newSlide(pptx);
  brandMark(s);
  s.addShape(pptx.ShapeType.rect, { x: 6.947, y: 1.209, w: 5.366, h: 6.444, fill: { color: TEAL } });
  s.addShape(pptx.ShapeType.rect, { x: 1.021, y: 1.209, w: 5.366, h: 6.444, fill: { color: WHITE } });
  star8(s, { x: 1.326, y: 1.521, w: 0.396, h: 0.396 }, TEAL);
  txt(s, 'GROUP PACKAGE', { x: 1.722, y: 1.54, w: 2.146, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  star8(s, { x: 7.251, y: 1.521, w: 0.396, h: 0.396 }, SAND);
  txt(s, 'FAMILY PACKAGE', { x: 7.647, y: 1.54, w: 2.211, h: 0.404, fontFace: HEAD, color: WHITE, fontSize: 18, wrap: false });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence .', { x: 1.229, y: 4.982, w: 4.757, h: 0.976, align: 'justify', fontSize: 12, color: SLATE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  s.addShape(pptx.ShapeType.rect, { x: 1.325, y: 6.291, w: 2.281, h: 0.38, fill: { color: TEAL } });
  txt(s, 'See More', { x: 2.024, y: 6.33, w: 0.882, h: 0.303, align: 'center', fontSize: 12, fontFace: HEAD, color: WHITE, wrap: false });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence .', { x: 7.251, y: 4.982, w: 4.757, h: 0.976, align: 'justify', fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  s.addShape(pptx.ShapeType.rect, { x: 7.347, y: 6.458, w: 2.281, h: 0.38, fill: { color: WHITE } });
  txt(s, 'See More', { x: 8.046, y: 6.496, w: 0.882, h: 0.303, align: 'center', fontSize: 12, fontFace: HEAD, color: TEAL, wrap: false });
  imageBox(s, { x: 1.325, y: 2.147, w: 4.757, h: 2.464 });
  imageBox(s, { x: 7.251, y: 2.147, w: 4.757, h: 2.464 });
}

function slide18(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  txt(s, [{ text: 'Facilities & ' }, { text: 'Support', options: { color: TEAL } }], { x: 1.021, y: 1.209, w: 6.424, h: 0.797, fontSize: 45, fontFace: HEAD, color: WHITE });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart..', { x: 3.764, y: 3.137, w: 2.535, h: 1.582, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, [{ text: 'Accommodation ', options: { breakLine: true } }, { text: 'Options' }], { x: 3.764, y: 2.278, w: 2.102, h: 0.858, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false, lineSpacingMultiple: 1.5 });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart..', { x: 9.778, y: 3.137, w: 2.535, h: 1.582, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, [{ text: 'Visa Assistance ', options: { breakLine: true } }, { text: '& Support' }], { x: 9.778, y: 2.278, w: 1.969, h: 0.707, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  s.addShape(pptx.ShapeType.rect, { x: 7.034, y: 5.175, w: 6.299, h: 1.669, fill: { color: TEAL } });
  s.addShape(pptx.ShapeType.ellipse, { x: 1.021, y: 5.635, w: 0.148, h: 0.148, fill: { color: SAND } });
  txt(s, 'Mornings of spring which I enjoy with my whole heart.', { x: 1.168, y: 5.837, w: 2.2, h: 0.673, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 1.169, y: 5.51, w: 1.499, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  s.addShape(pptx.ShapeType.ellipse, { x: 3.764, y: 5.635, w: 0.148, h: 0.148, fill: { color: SAND } });
  txt(s, 'Mornings of spring which I enjoy with my whole heart.', { x: 3.911, y: 5.837, w: 2.2, h: 0.673, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 3.912, y: 5.51, w: 1.499, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  s.addShape(pptx.ShapeType.ellipse, { x: 7.356, y: 5.635, w: 0.148, h: 0.148, fill: { color: SAND } });
  txt(s, 'Mornings of spring which I enjoy with my whole heart.', { x: 7.503, y: 5.837, w: 2.2, h: 0.673, align: 'justify', fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 7.504, y: 5.51, w: 1.499, h: 0.404, fontFace: HEAD, color: WHITE, fontSize: 18, wrap: false });
  s.addShape(pptx.ShapeType.ellipse, { x: 10.099, y: 5.635, w: 0.148, h: 0.148, fill: { color: SAND } });
  txt(s, 'Mornings of spring which I enjoy with my whole heart.', { x: 10.246, y: 5.837, w: 2.2, h: 0.673, align: 'justify', fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 10.247, y: 5.51, w: 1.499, h: 0.404, fontFace: HEAD, color: WHITE, fontSize: 18, wrap: false });
  imageBox(s, { x: 1.021, y: 2.278, w: 2.535, h: 2.638 });
  imageBox(s, { x: 7.034, y: 2.278, w: 2.535, h: 2.638 });
}

function slide19(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  s.addShape(pptx.ShapeType.rect, { x: 4.012, y: 2.511, w: 2.913, h: 3.744, fill: { color: WHITE } });
  s.addShape(pptx.ShapeType.rect, { x: 7.216, y: 2.511, w: 2.913, h: 3.744, fill: { color: WHITE } });
  s.addShape(pptx.ShapeType.rect, { x: 10.421, y: 2.511, w: 2.913, h: 3.744, fill: { color: WHITE } });
  txt(s, [{ text: 'Client ' }, { text: 'Reviews', options: { color: TEAL } }], { x: 3.455, y: 1.209, w: 6.424, h: 0.797, fontSize: 45, fontFace: HEAD, color: WHITE });
  txt(s, 'Client Reviews & Experiences', { x: 0.813, y: 3.037, w: 2.767, h: 1.717, fontSize: 32, fontFace: HEAD, color: TEAL });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring.', { x: 0.813, y: 4.754, w: 2.679, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  logoMark(s, { x: 5.253, y: 2.878, w: 0.431, h: 0.204 }, SAND);
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart.', { x: 4.33, y: 3.148, w: 2.276, h: 1.582, align: 'center', fontSize: 12, color: SLATE, italic: true, fontFace: BODY, lineSpacingMultiple: 1.5 });
  logoMark(s, { x: 8.457, y: 2.878, w: 0.431, h: 0.204 }, SAND);
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart.', { x: 7.535, y: 3.148, w: 2.276, h: 1.582, align: 'center', fontSize: 12, color: SLATE, italic: true, fontFace: BODY, lineSpacingMultiple: 1.5 });
  logoMark(s, { x: 11.661, y: 2.878, w: 0.431, h: 0.204 }, SAND);
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart.', { x: 10.739, y: 3.148, w: 2.276, h: 1.582, align: 'center', fontSize: 12, color: SLATE, italic: true, fontFace: BODY, lineSpacingMultiple: 1.5 });
  txt(s, 'Ahmer Mahmud', { x: 4.662, y: 4.951, w: 1.611, h: 0.337, align: 'center', fontSize: 14, fontFace: HEAD, color: TEAL, wrap: false });
  txt(s, 'Profession', { x: 4.557, y: 5.153, w: 1.822, h: 0.37, align: 'center', fontSize: 12, color: SLATE, bold: true, fontFace: BODY, lineSpacingMultiple: 1.5 });
  txt(s, 'Gail Alzair', { x: 8.122, y: 4.951, w: 1.101, h: 0.337, align: 'center', fontSize: 14, fontFace: HEAD, color: TEAL, wrap: false });
  txt(s, 'Profession', { x: 7.762, y: 5.153, w: 1.822, h: 0.37, align: 'center', fontSize: 12, color: SLATE, bold: true, fontFace: BODY, lineSpacingMultiple: 1.5 });
  txt(s, 'Anissa Alves', { x: 11.233, y: 4.951, w: 1.287, h: 0.337, align: 'center', fontSize: 14, fontFace: HEAD, color: TEAL, wrap: false });
  txt(s, 'Profession', { x: 10.966, y: 5.153, w: 1.822, h: 0.37, align: 'center', fontSize: 12, color: SLATE, bold: true, fontFace: BODY, lineSpacingMultiple: 1.5 });
  imageBox(s, { x: 4.968, y: 5.73, w: 1, h: 1 }, 'ellipse');
  imageBox(s, { x: 8.173, y: 5.73, w: 1, h: 1 }, 'ellipse');
  imageBox(s, { x: 11.377, y: 5.73, w: 1, h: 1 }, 'ellipse');
}

function slide20(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  txt(s, [{ text: 'Our ' }, { text: 'Achievement', options: { color: TEAL } }], { x: 6.808, y: 1.209, w: 5.646, h: 0.797, fontSize: 45, fontFace: HEAD, color: WHITE });
  txt(s, 'Best Travel Agency 2023', { x: 8.007, y: 2.247, w: 4.049, h: 1.178, fontSize: 32, fontFace: HEAD, color: TEAL });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone.', { x: 8.007, y: 3.376, w: 4.049, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  arabesque(s, { x: 6.807, y: 2.297, w: 1.078, h: 1.078 }, SAND);
  txt(s, '01', { x: 6.977, y: 2.517, w: 0.739, h: 0.64, align: 'center', fontSize: 32, fontFace: HEAD, color: TEAL });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 4.674, w: 13.333, h: 2.015, fill: { color: TEAL } });
  txt(s, 'Possession of my entire soul, like these sweet mornings of spring which.', { x: 1.021, y: 5.682, w: 3.068, h: 0.673, align: 'justify', fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  txt(s, 'Years experience', { x: 1.021, y: 5.355, w: 1.71, h: 0.337, fontSize: 14, fontFace: HEAD, color: WHITE });
  txt(s, '10', { x: 1.021, y: 4.874, w: 0.586, h: 0.64, fontSize: 32, fontFace: HEAD, color: WHITE });
  txt(s, 'Possession of my entire soul, like these sweet mornings of spring which.', { x: 5.004, y: 5.682, w: 3.068, h: 0.673, align: 'justify', fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  txt(s, 'Departures', { x: 5.004, y: 5.355, w: 1.71, h: 0.337, fontSize: 14, fontFace: HEAD, color: WHITE });
  txt(s, '+700K', { x: 5.004, y: 4.874, w: 1.803, h: 0.64, fontSize: 32, fontFace: HEAD, color: WHITE });
  txt(s, 'Possession of my entire soul, like these sweet mornings of spring which.', { x: 8.988, y: 5.682, w: 3.068, h: 0.673, align: 'justify', fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  txt(s, 'Best Team', { x: 8.988, y: 5.355, w: 1.71, h: 0.337, fontSize: 14, fontFace: HEAD, color: WHITE });
  txt(s, '+50', { x: 8.988, y: 4.874, w: 1.217, h: 0.64, fontSize: 32, fontFace: HEAD, color: WHITE });
  imageBox(s, { x: 0, y: 1.209, w: 6.384, h: 3.172 });
}

function slide21(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  arabesque(s, { x: 11.86, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  txt(s, [{ text: 'Our ' }, { text: 'Best Team', options: { color: TEAL } }], { x: 3.455, y: 1.209, w: 6.424, h: 0.797, fontSize: 45, fontFace: HEAD, color: WHITE });
  txt(s, 'Mornings of spring which I enjoy with my whole heart.', { x: 1.763, y: 5.695, w: 2.2, h: 0.673, align: 'center', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Alan Doe', { x: 2.209, y: 5.336, w: 1.31, h: 0.438, align: 'center', fontSize: 20, fontFace: HEAD, color: TEAL, wrap: false });
  txt(s, 'Mornings of spring which I enjoy with my whole heart.', { x: 5.566, y: 5.695, w: 2.2, h: 0.673, align: 'center', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Aisyah Kareen', { x: 5.664, y: 5.336, w: 2.006, h: 0.438, align: 'center', fontSize: 20, fontFace: HEAD, color: TEAL, wrap: false });
  txt(s, 'Mornings of spring which I enjoy with my whole heart.', { x: 9.369, y: 5.695, w: 2.2, h: 0.673, align: 'center', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Nur Linda', { x: 9.756, y: 5.336, w: 1.427, h: 0.438, align: 'center', fontSize: 20, fontFace: HEAD, color: TEAL, wrap: false });
  imageBox(s, { x: 1.622, y: 2.609, w: 2.483, h: 2.483 }, 'ellipse');
  imageBox(s, { x: 5.425, y: 2.609, w: 2.483, h: 2.483 }, 'ellipse');
  imageBox(s, { x: 9.228, y: 2.609, w: 2.483, h: 2.483 }, 'ellipse');
}

function slide22(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  s.addShape(pptx.ShapeType.rect, { x: 9, y: 0, w: 4.333, h: 7.5, fill: { color: TEAL } });
  arabesque(s, { x: 11.86, y: 5.645, w: 2.946, h: 2.946 }, SAND, 30);
  txt(s, [{ text: 'Umrah Guide ' }, { text: '(Mutawwif)', options: { color: TEAL } }], { x: 1.021, y: 1.208, w: 7.451, h: 0.797, fontSize: 45, fontFace: HEAD, color: WHITE });
  txt(s, 'Abdullah Hasan', { x: 1.021, y: 1.866, w: 4.618, h: 0.858, fontSize: 45, fontFace: HEAD, color: TEAL, wrap: false });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart.', { x: 1.021, y: 2.851, w: 2.951, h: 1.279, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'I am alone, and feel the charm of existence in this spot, which was created for the bliss of souls like mine. I am so happy, my dear friend.', { x: 4.334, y: 2.851, w: 2.951, h: 1.279, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Language Mastered', { x: 1.021, y: 4.602, w: 2.386, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  s.addShape(pptx.ShapeType.rect, { x: 1.06, y: 5.484, w: 2.912, h: 0.161, fill: { color: GREY } });
  txt(s, 'First Language', { x: 1.021, y: 5.143, w: 1.327, h: 0.303, fontSize: 12, bold: true, fontFace: BODY, color: WHITE, wrap: false });
  s.addShape(pptx.ShapeType.rect, { x: 1.06, y: 5.484, w: 2.697, h: 0.161, fill: { color: SAND } });
  s.addShape(pptx.ShapeType.rect, { x: 1.06, y: 6.238, w: 2.912, h: 0.161, fill: { color: GREY } });
  txt(s, 'Second Language', { x: 1.021, y: 5.898, w: 1.554, h: 0.303, fontSize: 12, bold: true, fontFace: BODY, color: WHITE, wrap: false });
  s.addShape(pptx.ShapeType.rect, { x: 1.06, y: 6.238, w: 2.531, h: 0.161, fill: { color: TEAL } });
  s.addShape(pptx.ShapeType.rect, { x: 4.373, y: 5.484, w: 2.912, h: 0.161, fill: { color: GREY } });
  txt(s, 'Third Language', { x: 4.334, y: 5.143, w: 1.385, h: 0.303, fontSize: 12, bold: true, fontFace: BODY, color: WHITE, wrap: false });
  s.addShape(pptx.ShapeType.rect, { x: 4.373, y: 5.484, w: 2.384, h: 0.161, fill: { color: TEAL_LT } });
  s.addShape(pptx.ShapeType.rect, { x: 4.373, y: 6.238, w: 2.912, h: 0.161, fill: { color: GREY } });
  txt(s, 'Fourth Language', { x: 4.334, y: 5.898, w: 1.482, h: 0.303, fontSize: 12, bold: true, fontFace: BODY, color: WHITE, wrap: false });
  s.addShape(pptx.ShapeType.rect, { x: 4.373, y: 6.238, w: 2.384, h: 0.161, fill: { color: SAND_LT } });
}

function slide23(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  imageBox(s, { x: 1.056, y: 1.387, w: 3.236, h: 6.918 }, 'roundRect');
  imageBox(s, { x: 0.875, y: 1.236, w: 3.597, h: 7.194 });
  txt(s, [{ text: 'Get in ' }, { text: 'Touch With Us', options: { color: TEAL } }], { x: 5.423, y: 1.209, w: 6.424, h: 0.797, fontSize: 45, fontFace: HEAD, color: WHITE });
  s.addShape(pptx.ShapeType.rect, { x: 3.906, y: 3.519, w: 2.538, h: 1.561, fill: { color: WHITE } });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence in this spot, which was created for the bliss of souls like mine.', { x: 5.423, y: 2.006, w: 6.591, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  s.addShape(pptx.ShapeType.rect, { x: 6.823, y: 3.519, w: 2.538, h: 1.561, fill: { color: WHITE } });
  s.addShape(pptx.ShapeType.rect, { x: 9.74, y: 3.519, w: 2.538, h: 1.561, fill: { color: WHITE } });
  icon(s, { x: 10.867, y: 3.73, w: 0.323, h: 0.404 }, SAND, 'bell');
  icon(s, { x: 7.968, y: 3.73, w: 0.249, h: 0.404 }, SAND, 'pin');
  icon(s, { x: 4.994, y: 3.73, w: 0.364, h: 0.404 }, SAND, 'calendar');
  txt(s, 'Suitable for all categories.', { x: 4.075, y: 4.5, w: 2.2, h: 0.37, fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 4.426, y: 4.173, w: 1.499, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  txt(s, 'Suitable for all categories.', { x: 6.992, y: 4.5, w: 2.2, h: 0.37, fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 7.342, y: 4.173, w: 1.499, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  txt(s, 'Suitable for all categories.', { x: 9.908, y: 4.5, w: 2.2, h: 0.37, fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 10.259, y: 4.173, w: 1.499, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  s.addShape(pptx.ShapeType.rect, { x: 3.906, y: 5.618, w: 8.371, h: 1.146, fill: { color: TEAL } });
  icon(s, { x: 4.197, y: 5.916, w: 0.551, h: 0.551 }, SAND, 'starBadge');
  txt(s, 'Sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence in this spot, which was created for the bliss of souls like mine. ', { x: 4.874, y: 5.854, w: 6.974, h: 0.673, align: 'justify', fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
}

function slide24(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  txt(s, [{ text: 'Registration ' }, { text: 'Process', options: { color: TEAL } }], { x: 1.021, y: 1.209, w: 7.452, h: 0.797, fontSize: 45, fontFace: HEAD, color: WHITE });
  s.addShape(pptx.ShapeType.rect, { x: 1.071, y: 2.245, w: 2.424, h: 2.591, fill: { color: WHITE } });
  txt(s, 'Step 01', { x: 1.5, y: 2.737, w: 1.764, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18 });
  txt(s, 'Has taken possession of my entire soul, like these sweet mornings of spring which I enjoy', { x: 1.406, y: 3.196, w: 1.856, h: 1.279, align: 'justify', fontSize: 12, color: SLATE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  s.addShape(pptx.ShapeType.rect, { x: 4.105, y: 2.245, w: 2.424, h: 2.591, fill: { color: WHITE } });
  txt(s, 'Step 02', { x: 4.534, y: 2.737, w: 1.764, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18 });
  txt(s, 'Has taken possession of my entire soul, like these sweet mornings of spring which I enjoy', { x: 4.44, y: 3.196, w: 1.856, h: 1.28, align: 'justify', fontSize: 12, color: SLATE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  s.addShape(pptx.ShapeType.rect, { x: 7.139, y: 2.245, w: 2.424, h: 2.591, fill: { color: WHITE } });
  txt(s, 'Step 03', { x: 7.569, y: 2.737, w: 1.764, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18 });
  txt(s, 'Has taken possession of my entire soul, like these sweet mornings of spring which I enjoy', { x: 7.474, y: 3.196, w: 1.856, h: 1.28, align: 'justify', fontSize: 12, color: SLATE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  s.addShape(pptx.ShapeType.rect, { x: 10.173, y: 2.245, w: 2.424, h: 2.591, fill: { color: WHITE } });
  txt(s, 'Step 04', { x: 10.603, y: 2.737, w: 1.764, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18 });
  txt(s, 'Has taken possession of my entire soul, like these sweet mornings of spring which I enjoy', { x: 10.508, y: 3.196, w: 1.856, h: 1.28, align: 'justify', fontSize: 12, color: SLATE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  star8(s, { x: 0.736, y: 2.604, w: 0.669, h: 0.669 }, TEAL);
  txt(s, '01', { x: 0.736, y: 2.737, w: 0.669, h: 0.404, align: 'center', fontFace: HEAD, color: WHITE, fontSize: 18 });
  star8(s, { x: 3.77, y: 2.604, w: 0.669, h: 0.669 }, TEAL);
  txt(s, '02', { x: 3.77, y: 2.737, w: 0.669, h: 0.404, align: 'center', fontFace: HEAD, color: WHITE, fontSize: 18 });
  star8(s, { x: 6.804, y: 2.604, w: 0.669, h: 0.669 }, TEAL);
  txt(s, '03', { x: 6.804, y: 2.737, w: 0.669, h: 0.404, align: 'center', fontFace: HEAD, color: WHITE, fontSize: 18 });
  star8(s, { x: 9.839, y: 2.604, w: 0.669, h: 0.669 }, TEAL);
  txt(s, '04', { x: 9.839, y: 2.737, w: 0.669, h: 0.404, align: 'center', fontFace: HEAD, color: WHITE, fontSize: 18 });
  s.addShape(pptx.ShapeType.rect, { x: 1.065, y: 5.328, w: 11.532, h: 1.146, fill: { color: TEAL } });
  icon(s, { x: 1.355, y: 5.626, w: 0.551, h: 0.551 }, SAND, 'starBadge');
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence in this spot, which was created for the bliss of souls like mine. I am so happy, my dear friend.', { x: 2.021, y: 5.565, w: 10.248, h: 0.673, align: 'justify', fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
}

function slide25(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  txt(s, [{ text: 'Demographic ' }, { text: 'Hajj Chart', options: { color: TEAL } }], { x: 1.021, y: 1.209, w: 7.452, h: 0.797, fontSize: 45, fontFace: HEAD, color: WHITE });
  icon(s, { x: 8.672, y: 5.542, w: 0.667, h: 0.776 }, TEAL, 'woman');
  icon(s, { x: 8.672, y: 4.265, w: 0.667, h: 0.74 }, SAND, 'man');
  txt(s, 'Created for the bliss of souls like mine. I am so happy, my dear friend.', { x: 9.371, y: 4.444, w: 2.942, h: 0.673, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Male Pilgrims', { x: 9.371, y: 4.117, w: 1.717, h: 0.404, fontFace: HEAD, fontSize: 18, color: WHITE, wrap: false });
  txt(s, 'Created for the bliss of souls like mine. I am so happy, my dear friend.', { x: 9.371, y: 5.722, w: 2.942, h: 0.673, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Female Pilgrims', { x: 9.371, y: 5.395, w: 1.999, h: 0.404, fontFace: HEAD, fontSize: 18, color: WHITE, wrap: false });
  txt(s, '+500K Pilgrims', { x: 8.672, y: 2.242, w: 3.184, h: 0.64, fontSize: 32, fontFace: HEAD, color: TEAL, wrap: false });
  txt(s, 'I am so happy, my dear friend, so absorbed in the exquisite sense of mere tranquil existence, ', { x: 8.672, y: 2.774, w: 3.64, h: 0.673, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  imageBox(s, { x: 0.778, y: 2.562, w: 7.255, h: 3.833 });
}

function slide26(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  txt(s, [{ text: 'Data ' }, { text: 'Comparison', options: { color: TEAL } }], { x: 1.021, y: 1.209, w: 7.452, h: 0.797, fontSize: 45, fontFace: HEAD, color: WHITE });
  txt(s, '2023', { x: 5.324, y: 5.313, w: 2.185, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18 });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like', { x: 5.324, y: 5.618, w: 2.686, h: 0.673, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, '2024', { x: 9.361, y: 5.313, w: 2.185, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18 });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like', { x: 9.361, y: 5.618, w: 2.686, h: 0.673, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, '2022', { x: 1.286, y: 5.313, w: 2.185, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18 });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like', { x: 1.286, y: 5.618, w: 2.686, h: 0.673, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  imageBox(s, { x: 1.286, y: 2.34, w: 2.535, h: 2.638 });
  imageBox(s, { x: 5.324, y: 2.34, w: 2.535, h: 2.638 });
  imageBox(s, { x: 9.186, y: 2.34, w: 2.535, h: 2.638 });
}

function slide27(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  txt(s, [{ text: 'Global ' }, { text: 'Pilgrims Data', options: { color: TEAL } }], { x: 1.021, y: 1.209, w: 7.452, h: 0.797, fontSize: 45, fontFace: HEAD, color: WHITE });
  txt(s, 'for all categories', { x: 6.888, y: 4.341, w: 1.463, h: 0.37, align: 'center', fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  txt(s, '78% - Asia', { x: 7.076, y: 4.189, w: 1.087, h: 0.337, align: 'center', fontSize: 14, fontFace: HEAD, color: WHITE, wrap: false });
  txt(s, 'for all categories', { x: 1.888, y: 4.306, w: 1.463, h: 0.37, align: 'center', fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  txt(s, '30% - America', { x: 1.888, y: 4.155, w: 1.464, h: 0.337, align: 'center', fontSize: 14, fontFace: HEAD, color: WHITE, wrap: false });
  txt(s, 'for all categories', { x: 4.249, y: 4.771, w: 1.463, h: 0.37, align: 'center', fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  txt(s, '50% - Africa', { x: 4.351, y: 4.62, w: 1.259, h: 0.337, align: 'center', fontSize: 14, fontFace: HEAD, color: WHITE, wrap: false });
  txt(s, 'Sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence in this spot.', { x: 9.133, y: 2.522, w: 3.178, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 9.134, y: 2.195, w: 1.499, h: 0.404, fontFace: HEAD, color: TEAL, fontSize: 18, wrap: false });
  imageBox(s, { x: 9.216, y: 3.834, w: 3.013, h: 2.638 });
  worldMap(s, { x: 0.615, y: 2.561, w: 8.263, h: 4.031 });
}

function slide28(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  txt(s, [{ text: 'What to Prepare ', options: { breakLine: true } }, { text: 'Before Traveling', options: { color: TEAL } }], { x: 1.021, y: 1.209, w: 7.452, h: 1.479, fontSize: 45, fontFace: HEAD, color: WHITE });
  s.addShape(pptx.ShapeType.rect, { x: 6.91, y: 3.031, w: 5.403, h: 1.393, fill: { color: WHITE } });
  s.addShape(pptx.ShapeType.rect, { x: 1.021, y: 3.031, w: 5.403, h: 1.393, fill: { color: TEAL } });
  s.addShape(pptx.ShapeType.rect, { x: 1.021, y: 5.028, w: 5.403, h: 1.393, fill: { color: WHITE } });
  s.addShape(pptx.ShapeType.rect, { x: 6.91, y: 5.028, w: 5.403, h: 1.393, fill: { color: TEAL } });
  icon(s, { x: 1.332, y: 3.391, w: 0.673, h: 0.673 }, SAND, 'moonBadge');
  txt(s, 'I am alone, and feel the charm of existence in this spot, which was created for the bliss of souls like.', { x: 2.165, y: 3.554, w: 3.946, h: 0.673, align: 'justify', fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 2.165, y: 3.227, w: 1.499, h: 0.404, fontFace: HEAD, color: WHITE, fontSize: 18, wrap: false });
  icon(s, { x: 7.222, y: 5.388, w: 0.673, h: 0.673 }, SAND, 'moonBadge');
  txt(s, 'I am alone, and feel the charm of existence in this spot, which was created for the bliss of souls like.', { x: 8.055, y: 5.551, w: 3.946, h: 0.673, align: 'justify', fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 8.055, y: 5.225, w: 1.499, h: 0.404, fontFace: HEAD, color: WHITE, fontSize: 18, wrap: false });
  icon(s, { x: 7.222, y: 3.391, w: 0.673, h: 0.673 }, SAND, 'moonBadge');
  txt(s, 'I am alone, and feel the charm of existence in this spot, which was created for the bliss of souls like.', { x: 8.055, y: 3.554, w: 3.946, h: 0.673, align: 'justify', fontSize: 12, color: SLATE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 8.055, y: 3.227, w: 1.499, h: 0.404, fontFace: HEAD, color: SLATE, fontSize: 18, wrap: false });
  icon(s, { x: 1.332, y: 5.388, w: 0.673, h: 0.673 }, SAND, 'moonBadge');
  txt(s, 'I am alone, and feel the charm of existence in this spot, which was created for the bliss of souls like.', { x: 2.165, y: 5.551, w: 3.946, h: 0.673, align: 'justify', fontSize: 12, color: SLATE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  txt(s, 'Description', { x: 2.165, y: 5.225, w: 1.499, h: 0.404, fontFace: HEAD, color: SLATE, fontSize: 18, wrap: false });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence in this spot,', { x: 6.91, y: 1.46, w: 5.403, h: 0.976, align: 'justify', fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
}

function slide29(pptx) {
  const s = newSlide(pptx);
  arabesque(s, { x: -1.473, y: 5.645, w: 2.946, h: 2.946 }, SAND);
  brandMark(s);
  imageBox(s, { x: 6.867, y: 1.209, w: 5.387, h: 5.436 });
  txt(s, [{ text: 'Our ' }, { text: 'Contact', options: { color: TEAL } }], { x: 1.021, y: 1.208, w: 3.675, h: 0.797, fontSize: 45, fontFace: HEAD, color: WHITE });
  txt(s, '123-456-7890', { x: 1.393, y: 2.344, w: 2.011, h: 0.37, fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Phone', { x: 1.393, y: 2.16, w: 1.205, h: 0.337, fontSize: 14, fontFace: HEAD, color: TEAL });
  icon(s, { x: 1.072, y: 2.296, w: 0.238, h: 0.282 }, SAND, 'phone');
  txt(s, 'youremail@domain.com', { x: 1.393, y: 3.251, w: 2.011, h: 0.37, fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Email', { x: 1.393, y: 3.067, w: 1.205, h: 0.337, fontSize: 14, fontFace: HEAD, color: TEAL });
  icon(s, { x: 1.021, y: 3.203, w: 0.305, h: 0.282 }, SAND, 'send');
  txt(s, 'www.yourwebsite.com', { x: 4.294, y: 2.344, w: 2.011, h: 0.37, fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Website', { x: 4.294, y: 2.16, w: 1.205, h: 0.337, fontSize: 14, fontFace: HEAD, color: TEAL });
  icon(s, { x: 3.939, y: 2.296, w: 0.282, h: 0.282 }, SAND, 'globe');
  txt(s, '@youraccountname', { x: 4.294, y: 3.251, w: 2.011, h: 0.37, fontSize: 12, fontFace: BODY, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Follow Us', { x: 4.294, y: 3.067, w: 1.205, h: 0.337, fontSize: 14, fontFace: HEAD, color: TEAL });
  icon(s, { x: 3.939, y: 3.203, w: 0.282, h: 0.282 }, SAND, 'globe');
  s.addShape(pptx.ShapeType.rect, { x: 6.877, y: 5.159, w: 3.994, h: 1.146, fill: { color: TEAL } });
  icon(s, { x: 7.022, y: 5.457, w: 0.551, h: 0.551 }, SAND, 'starBadge');
  txt(s, 'Sweet mornings of spring which I enjoy with my whole heart. I am alone,.', { x: 7.659, y: 5.488, w: 3.055, h: 0.673, align: 'justify', fontSize: 12, color: WHITE, fontFace: BODY, lineSpacingMultiple: 1.5 });
  txt(s, 'Location', { x: 7.639, y: 5.303, w: 1.205, h: 0.337, fontSize: 14, fontFace: HEAD, color: WHITE });
  imageBox(s, { x: 3.938, y: 4.117, w: 2.528, h: 2.528 });
  imageBox(s, { x: 1.045, y: 4.117, w: 2.528, h: 2.528 });
}

function slide30(pptx) {
  const s = newSlide(pptx);
  imageBox(s, { x: 0, y: 0, w: 13.333, h: 7.5 });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: TEAL, transparency: 30 } });
  arabesque(s, { x: -1.473, y: -0.893, w: 2.946, h: 2.946 }, SAND);
  arabesque(s, { x: 11.86, y: -0.893, w: 2.946, h: 2.946 }, SAND);
  arabesque(s, { x: -1.473, y: 5.447, w: 2.946, h: 2.946 }, SAND);
  arabesque(s, { x: 11.86, y: 5.447, w: 2.946, h: 2.946 }, SAND);
  txt(s, 'Thank You', { x: 2.523, y: 1.644, w: 8.287, h: 2.121, align: 'center', fontSize: 120, fontFace: HEAD, color: WHITE, wrap: false });
  txt(s, 'For Trusting Us', { x: 3.967, y: 3.429, w: 5.4, h: 1.01, align: 'center', fontSize: 54, fontFace: HEAD, color: WHITE, wrap: false });
  txt(s, 'Rihlah 2025', { x: 4.973, y: 6.743, w: 3.387, h: 0.353, align: 'center', fontSize: 15, color: WHITE, bold: true, fontFace: BODY });
}

/* --------------------------------------------------------------- assembly */
function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'RIHLAH', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'RIHLAH';
  pptx.author = 'Rihlah';
  pptx.title = 'Rihlah — Hajj & Umrah Agency';

  const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
    slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30];
  SLIDES.forEach(function (fn) { fn(pptx); });

  return pptx.writeFile({
    fileName: path.join(__dirname, '0d332b04-2443-4c14-9659-c807d3cbef75_grok_final.pptx'),
  });
}

build().then(function (f) { console.log('wrote', f); }).catch(function (e) { console.error(e); process.exit(1); });
