/**
 * "Fabulous / Ourslide" lookbook deck - 30 slides, 13.335 x 7.5 in.
 *
 * Rebuilt with pptxgenjs only. Photographs in the source deck are represented
 * by flat colour blocks labelled "[image]"; every other element (blocks,
 * freeform silhouettes, rules, text) is drawn natively.
 *
 *   node 0c02ebdd-c7f5-4a50-94c5-3941e5da61eb_grok_final.js
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const TEAL = '5B8F8F'; // accent2
const PEACH = 'F5B09B'; // accent1
const MIST = 'F2F2F2'; // bg1 lumMod 95%
const GRAY = '808080'; // body copy
const SLATE = '595959'; // sub-headings
const INK = '404040'; // dark headings
const SILVER = 'BFBFBF'; // footer
const WHITE = 'FFFFFF';
const PHOTO = '474747'; // stand-in colour for the deck's photography
const PHOTO_CAP = 'A6A6A6';
const SKIN_24 = 'D5B09D'; // hands in the slide 24 mock-up photo
const SKIN_25 = 'CDA184'; // hands in the slide 25 mock-up photo

const HEAD = 'Montserrat SemiBold'; // theme major latin
const SANS = 'Roboto'; // theme minor latin

const W = 13.335;
const H = 7.5;

/* ------------------------------------------------------------------- copy */

const INTRO =
  'PLACEHOLDER' +
  'PLACEHOLDER';

const LOREM =
  'PLACEHOLDER' +
  'PLACEHOLDER';
const LOREM_TO_ELEMENTUM =
  'PLACEHOLDER' +
  'turpis antes elementum';
const LOREM_TO_DAPIBUS =
  'PLACEHOLDER' +
  'turpis antes elementum ipsum volutpat tincidunt dapibus';
const LOREM_TO_DA =
  'PLACEHOLDER' +
  'PLACEHOLDER';
const LOREM_TO_VEL =
  'PLACEHOLDER' +
  'PLACEHOLDER';
const LOREM_TO_MAGNA =
  'PLACEHOLDER' +
  'PLACEHOLDER';

const IPSUM =
  'PLACEHOLDER' +
  'PLACEHOLDER' +
  'PLACEHOLDER';
const IPSUM_TO_LACUS =
  'PLACEHOLDER' +
  'PLACEHOLDER' +
  'eleifend nulla dictum pellentesque lacus';
const IPSUM_TO_TINCIDUNT =
  'PLACEHOLDER' +
  'turpis antes dolor elementum ipsum volutpat tincidunt';
const IPSUM_FROM_ELEMENTUM =
  'PLACEHOLDER' +
  'PLACEHOLDER';
const IPSUM_TO_MOLLIS =
  'PLACEHOLDER' +
  'PLACEHOLDER';

const MAP_COPY =
  'PLACEHOLDER' +
  'PLACEHOLDER' +
  'eleifend nulla dictum pellentesque lacus id dapibus quam ';

/* ---------------------------------------------------------------- helpers */

/** Flat colour block. */
function block(s, x, y, w, h, color, o) {
  s.addShape('rect', Object.assign({ x, y, w, h, fill: { color }, line: { type: 'none' } }, o));
}

/** Stand-in for one of the deck's photographs. */
function photo(s, x, y, w, h, o) {
  const shape = (o && o.round) ? 'roundRect' : 'rect';
  const geom = Object.assign({ x, y, w, h }, o && o.rotate !== undefined ? { rotate: o.rotate } : null);
  s.addShape(shape, Object.assign({ fill: { color: PHOTO }, line: { type: 'none' }, rectRadius: 0.12 }, geom));
  s.addText('[image]', Object.assign(
    { align: 'center', valign: 'middle', margin: 0, fontFace: SANS, fontSize: 11, color: PHOTO_CAP },
    geom
  ));
}

/** Dark device shell used behind the mock-up screens. */
function deviceFrame(s, x, y, w, h, o) {
  s.addShape('roundRect', Object.assign(
    { x, y, w, h, rectRadius: 0.16, fill: { color: '2E2E33' }, line: { type: 'none' } }, o));
}

/** Montserrat display heading; `lines` may be a string or an array of lines. */
function heading(s, x, y, w, h, lines, o) {
  o = o || {};
  const runs = [].concat(lines).map((t, i, a) => ({ text: t, options: { breakLine: i < a.length - 1 } }));
  s.addText(runs, {
    x, y, w, h, margin: 0, wrap: false, valign: 'top', fit: 'resize',
    fontFace: o.fontFace || HEAD, fontSize: o.size || 28, color: o.color || TEAL,
    bold: o.bold || false, align: o.align || 'left', rotate: o.rotate,
  });
}

/** Justified body copy at 150% leading. */
function body(s, x, y, w, h, text, o) {
  o = o || {};
  s.addText(text, {
    x, y, w, h, margin: 0, valign: 'top', fit: 'resize',
    fontFace: SANS, fontSize: o.size || 12, color: o.color || GRAY,
    align: o.align || 'justify', lineSpacingMultiple: o.lead === undefined ? 1.5 : o.lead,
    bold: o.bold || false, italic: o.italic || false, wrap: o.wrap,
  });
}

/** Small bold sub-heading. */
function label(s, x, y, w, h, text, o) {
  o = o || {};
  s.addText(text, {
    x, y, w, h, margin: 0, valign: 'top', wrap: false, fit: 'resize',
    fontFace: SANS, fontSize: o.size || 14, bold: true, color: o.color || SLATE,
    lineSpacingMultiple: o.lead,
  });
}

/** Closed polygon from normalised (0..1) coordinate pairs. */
function polygon(s, x, y, w, h, norm, color, o) {
  const points = [];
  for (let i = 0; i < norm.length; i += 2) {
    points.push({ x: +(norm[i] * w).toFixed(4), y: +(norm[i + 1] * h).toFixed(4), moveTo: i === 0 });
  }
  points.push({ close: true });
  s.addShape('custGeom', Object.assign(
    { x, y, w, h, points, fill: { color }, line: { type: 'none' } }, o));
}

/** Closed polygon from absolute inch coordinate pairs. */
function blob(s, abs, color) {
  const xs = abs.filter((_, i) => i % 2 === 0);
  const ys = abs.filter((_, i) => i % 2 === 1);
  const x = Math.min(...xs), y = Math.min(...ys);
  const w = Math.max(...xs) - x, h = Math.max(...ys) - y;
  const norm = abs.map((v, i) => (i % 2 === 0 ? (v - x) / w : (v - y) / h));
  polygon(s, x, y, w, h, norm, color);
}

/** The rotated copyright strip that runs down the right edge of every slide. */
function footer(s, accent) {
  s.addText([
    { text: '\u00A9 2023  ', options: { color: SILVER } },
    { text: 'Ourslide ~ Template  ', options: { color: accent || PEACH, bold: true } },
    { text: 'all rights reserved', options: { color: SILVER } },
  ], {
    x: 11.609, y: 1.999, w: 2.987, h: 0.168, fit: 'resize',
    margin: 0, wrap: false, valign: 'top', rotate: 90, fontFace: SANS, fontSize: 10,
  });
}

/* -------------------------------------------- silhouettes (normalised 0..1) */

const CHINA = [
  0.83,0.435, 0.86,0.407, 0.886,0.399, 0.907,0.375, 0.925,0.354, 0.929,0.361, 0.93,0.358,
  0.942,0.343, 0.937,0.291, 0.972,0.278, 0.99,0.219, 0.998,0.179, 0.969,0.188, 0.934,0.193,
  0.916,0.147, 0.882,0.132, 0.869,0.08, 0.851,0.025, 0.823,0.006, 0.78,0.007, 0.763,0.034,
  0.754,0.078, 0.738,0.124, 0.707,0.134, 0.705,0.133, 0.704,0.132, 0.703,0.131, 0.702,0.129,
  0.695,0.148, 0.693,0.195, 0.729,0.191, 0.74,0.23, 0.686,0.273, 0.632,0.284, 0.622,0.329,
  0.566,0.368, 0.518,0.386, 0.476,0.381, 0.42,0.363, 0.362,0.345, 0.336,0.302, 0.304,0.287,
  0.279,0.256, 0.27,0.208, 0.24,0.18, 0.227,0.155, 0.225,0.155, 0.216,0.15, 0.215,0.152,
  0.199,0.176, 0.185,0.225, 0.147,0.237, 0.135,0.283, 0.101,0.289, 0.111,0.352, 0.104,0.362,
  0.104,0.365, 0.104,0.369, 0.104,0.375, 0.104,0.378, 0.087,0.39, 0.032,0.433, 0.004,0.442,
  0.003,0.443, 0.002,0.445, 0.001,0.448, 0,0.451, 0,0.454, 0,0.468, 0.007,0.484, 0.019,0.513,
  0.019,0.516, 0.019,0.52, 0.014,0.531, 0.036,0.551, 0.058,0.578, 0.069,0.577, 0.096,0.567,
  0.1,0.605, 0.083,0.632, 0.088,0.665, 0.079,0.678, 0.1,0.71, 0.115,0.728, 0.124,0.727,
  0.146,0.739, 0.18,0.773, 0.21,0.792, 0.23,0.792, 0.237,0.792, 0.238,0.792, 0.238,0.791,
  0.245,0.804, 0.256,0.79, 0.275,0.79, 0.293,0.799, 0.316,0.778, 0.346,0.76, 0.365,0.752,
  0.369,0.78, 0.385,0.782, 0.398,0.794, 0.407,0.834, 0.39,0.88, 0.398,0.898, 0.409,0.913,
  0.417,0.944, 0.433,0.969, 0.447,0.96, 0.448,0.965, 0.458,0.969, 0.455,0.947, 0.463,0.942,
  0.466,0.941, 0.481,0.939, 0.496,0.934, 0.519,0.922, 0.539,0.939, 0.557,0.965, 0.579,0.969,
  0.594,0.996, 0.618,0.968, 0.653,0.948, 0.669,0.94, 0.702,0.919, 0.73,0.886, 0.751,0.84,
  0.769,0.799, 0.788,0.754, 0.777,0.726, 0.785,0.699, 0.774,0.662, 0.744,0.593, 0.779,0.54,
  0.786,0.518, 0.741,0.523, 0.717,0.487, 0.767,0.428, 0.782,0.445, 0.791,0.466,
];

const INDIA = [
  1,0.289, 0.988,0.289, 0.969,0.276, 0.964,0.253, 0.946,0.25, 0.919,0.26, 0.881,0.27,
  0.838,0.304, 0.808,0.31, 0.812,0.32, 0.808,0.342, 0.765,0.344, 0.741,0.348, 0.707,0.324,
  0.707,0.317, 0.704,0.296, 0.693,0.301, 0.692,0.301, 0.691,0.301, 0.688,0.301, 0.683,0.315,
  0.684,0.343, 0.664,0.357, 0.616,0.351, 0.579,0.34, 0.549,0.319, 0.506,0.317, 0.473,0.305,
  0.441,0.286, 0.41,0.265, 0.421,0.236, 0.436,0.218, 0.427,0.211, 0.38,0.18, 0.353,0.141,
  0.367,0.141, 0.387,0.12, 0.367,0.091, 0.376,0.074, 0.415,0.035, 0.393,0.006, 0.349,0.01,
  0.338,0.019, 0.314,0.03, 0.256,0.057, 0.202,0.047, 0.204,0.103, 0.24,0.137, 0.219,0.164,
  0.218,0.189, 0.184,0.224, 0.148,0.267, 0.099,0.308, 0.076,0.299, 0.044,0.343, 0.066,0.369,
  0.089,0.398, 0.091,0.435, 0.068,0.434, 0.018,0.443, 0.01,0.469, 0.061,0.496, 0.053,0.535,
  0.13,0.546, 0.152,0.515, 0.155,0.615, 0.218,0.793, 0.276,0.919, 0.332,0.992, 0.375,0.946,
  0.401,0.888, 0.42,0.826, 0.416,0.743, 0.481,0.709, 0.541,0.648, 0.619,0.591, 0.656,0.532,
  0.696,0.531, 0.711,0.502, 0.704,0.464, 0.688,0.422, 0.7,0.404, 0.698,0.39, 0.682,0.375,
  0.693,0.356, 0.698,0.357, 0.7,0.357, 0.701,0.358, 0.702,0.358, 0.704,0.358, 0.726,0.371,
  0.741,0.396, 0.793,0.402, 0.83,0.41, 0.817,0.431, 0.791,0.446, 0.796,0.477, 0.816,0.456,
  0.836,0.496, 0.838,0.512, 0.844,0.515, 0.857,0.498, 0.863,0.472, 0.867,0.441, 0.897,0.433,
  0.91,0.405, 0.914,0.39, 0.923,0.362, 0.968,0.325, 0.985,0.317,
];

const SUMATRA = [
  0.976,0.976, 0.989,0.932, 0.985,0.796, 1,0.762, 0.939,0.692, 0.864,0.716, 0.886,0.684,
  0.853,0.636, 0.842,0.598, 0.768,0.568, 0.786,0.486, 0.759,0.46, 0.672,0.486, 0.713,0.454,
  0.718,0.408, 0.656,0.418, 0.632,0.382, 0.6,0.328, 0.573,0.336, 0.523,0.31, 0.512,0.328,
  0.468,0.286, 0.411,0.232, 0.341,0.192, 0.263,0.098, 0.182,0.066, 0.114,0.062, 0.013,0.056,
  0.042,0.112, 0.138,0.184, 0.208,0.252, 0.236,0.28, 0.319,0.346, 0.339,0.386, 0.37,0.472,
  0.464,0.56, 0.49,0.602, 0.527,0.674, 0.543,0.712, 0.615,0.78, 0.683,0.852, 0.753,0.896,
  0.867,0.99, 0.888,0.974, 0.956,0.976,
];

const JAVA = [
  0.98,0.635, 0.944,0.595, 0.824,0.635, 0.793,0.508, 0.76,0.373, 0.688,0.31, 0.614,0.262,
  0.575,0.238, 0.532,0.365, 0.422,0.357, 0.391,0.317, 0.348,0.27, 0.317,0.159, 0.276,0.143,
  0.235,0.127, 0.21,0.04, 0.179,0.056, 0.159,0.119, 0.123,0.063, 0.061,0.119, 0.046,0.214,
  0.018,0.333, 0.051,0.365, 0.079,0.357, 0.123,0.421, 0.13,0.548, 0.235,0.579, 0.327,0.667,
  0.389,0.619, 0.486,0.675, 0.565,0.786, 0.645,0.802, 0.742,0.865, 0.777,0.905, 0.808,0.897,
  0.829,0.849, 0.882,0.897, 0.926,0.937, 0.964,0.968, 0.974,0.937, 0.987,0.857, 0.977,0.762,
  0.997,0.683,
];

const BORNEO = [
  0.814,0.634, 0.851,0.584, 0.883,0.427, 0.953,0.419, 0.97,0.372, 0.886,0.279, 0.895,0.226,
  0.828,0.123, 0.867,0.098, 0.86,0.056, 0.821,0.006, 0.746,0.006, 0.669,0.045, 0.629,0.165,
  0.604,0.237, 0.569,0.327, 0.51,0.338, 0.424,0.344, 0.338,0.349, 0.256,0.397, 0.193,0.408,
  0.175,0.416, 0.11,0.355, 0.072,0.265, 0.037,0.321, 0.005,0.427, 0.014,0.48, 0.042,0.581,
  0.086,0.654, 0.124,0.718, 0.138,0.804, 0.149,0.858, 0.231,0.863, 0.294,0.858, 0.31,0.93,
  0.352,0.899, 0.38,0.905, 0.429,0.874, 0.483,0.916, 0.529,0.897, 0.559,0.955, 0.576,0.989,
  0.639,0.969, 0.699,0.947, 0.727,0.975, 0.744,0.905, 0.711,0.88, 0.744,0.824, 0.741,0.743,
  0.779,0.665,
];

const SULAWESI = [
  0.638,0.968, 0.685,0.927, 0.678,0.861, 0.656,0.826, 0.627,0.794, 0.688,0.794, 0.659,0.759,
  0.558,0.731, 0.565,0.658, 0.46,0.532, 0.569,0.424, 0.656,0.38, 0.641,0.449, 0.696,0.446,
  0.75,0.44, 0.707,0.373, 0.696,0.304, 0.576,0.326, 0.467,0.354, 0.38,0.396, 0.293,0.386,
  0.217,0.307, 0.283,0.161, 0.377,0.177, 0.467,0.168, 0.598,0.174, 0.75,0.199, 0.884,0.171,
  0.993,0.054, 0.899,0.038, 0.736,0.104, 0.63,0.114, 0.525,0.079, 0.395,0.066, 0.304,0.098,
  0.188,0.177, 0.159,0.275, 0.094,0.475, 0.043,0.573, 0.029,0.658, 0.094,0.687, 0.134,0.772,
  0.123,0.87, 0.156,0.981, 0.257,0.956, 0.25,0.899, 0.264,0.804, 0.261,0.747, 0.268,0.665,
  0.268,0.601, 0.362,0.614, 0.355,0.722, 0.446,0.807, 0.518,0.848, 0.572,0.883,
];

const PAPUA = [
  0.989,0.821, 0.987,0.744, 0.989,0.649, 0.987,0.277, 0.934,0.237, 0.833,0.216, 0.711,0.142,
  0.654,0.169, 0.57,0.232, 0.509,0.309, 0.419,0.325, 0.406,0.272, 0.368,0.15, 0.371,0.106,
  0.274,0.047, 0.134,0.055, 0.101,0.077, 0.055,0.066, 0.018,0.066, 0.05,0.103, 0.072,0.132,
  0.16,0.182, 0.18,0.23, 0.263,0.237, 0.292,0.222, 0.322,0.253, 0.274,0.256, 0.208,0.282,
  0.169,0.288, 0.228,0.346, 0.25,0.435, 0.305,0.404, 0.362,0.412, 0.393,0.433, 0.485,0.483,
  0.561,0.501, 0.717,0.586, 0.77,0.691, 0.77,0.797, 0.708,0.831, 0.715,0.918, 0.794,0.908,
  0.851,0.892, 0.904,0.892, 0.991,1,
];

const NUSA = [
  0.944,0.048, 0.778,0.067, 0.642,0.106, 0.457,0.183, 0.315,0.394, 0.179,0.481, 0.031,0.596,
  0.012,0.808, 0.074,0.962, 0.179,0.846, 0.346,0.712, 0.512,0.49, 0.673,0.404, 0.815,0.298,
  0.963,0.183,
];

const MALUKU = [
  0.955,0.588, 0.833,0.255, 0.644,0.157, 0.492,0.039, 0.341,0.118, 0.235,0.118, 0.091,0.176,
  0.03,0.627, 0.076,0.843, 0.152,0.804, 0.091,0.647, 0.136,0.451, 0.182,0.627, 0.295,0.667,
  0.333,0.588, 0.394,0.627, 0.492,0.706, 0.545,0.647, 0.621,0.588, 0.697,0.706, 0.788,0.804,
  0.856,0.882, 0.939,1, 0.97,0.804,
];

const HOKKAIDO = [
  0.975,0.568, 0.95,0.493, 0.954,0.326, 0.929,0.374, 0.792,0.401, 0.725,0.374, 0.483,0.185,
  0.312,0.119, 0.35,0.233, 0.312,0.401, 0.3,0.458, 0.254,0.59, 0.192,0.577, 0.104,0.568,
  0.104,0.617, 0.025,0.714, 0.033,0.846, 0.058,0.89, 0.087,0.978, 0.133,0.938, 0.221,0.894,
  0.175,0.846, 0.113,0.767, 0.237,0.767, 0.3,0.736, 0.446,0.797, 0.546,0.837, 0.675,0.731,
  0.775,0.652, 0.979,0.608,
];

const HONSHU = [
  0.951,0.128, 0.934,0.075, 0.888,0.025, 0.911,0.048, 0.85,0.035, 0.82,0.093, 0.795,0.113,
  0.801,0.185, 0.801,0.306, 0.78,0.346, 0.698,0.466, 0.662,0.556, 0.581,0.564, 0.505,0.622,
  0.476,0.649, 0.452,0.727, 0.397,0.732, 0.381,0.704, 0.273,0.737, 0.159,0.752, 0.07,0.842,
  0.006,0.88, 0.017,0.917, 0.091,0.92, 0.129,0.92, 0.127,0.892, 0.169,0.897, 0.188,0.88,
  0.256,0.86, 0.334,0.83, 0.389,0.84, 0.37,0.917, 0.442,0.975, 0.495,0.897, 0.518,0.852,
  0.516,0.81, 0.605,0.84, 0.717,0.807, 0.74,0.777, 0.797,0.777, 0.854,0.767, 0.869,0.729,
  0.869,0.637, 0.879,0.584, 0.896,0.454, 0.945,0.378, 0.96,0.316, 0.994,0.216,
];

const SHIKOKU = [
  0.763,0.578, 0.851,0.433, 0.912,0.311, 0.886,0.122, 0.711,0.056, 0.561,0.178, 0.482,0.278,
  0.377,0.344, 0.272,0.233, 0.167,0.422, 0.035,0.533, 0.079,0.6, 0.132,0.689, 0.193,0.867,
  0.298,0.867, 0.404,0.656, 0.535,0.533, 0.64,0.578,
];

const KYUSHU = [
  0.913,0.245, 0.846,0.165, 0.663,0.122, 0.567,0.05, 0.212,0.151, 0.048,0.223, 0.154,0.338,
  0.26,0.338, 0.058,0.424, 0.154,0.46, 0.308,0.374, 0.346,0.345, 0.365,0.468, 0.212,0.496,
  0.26,0.583, 0.317,0.518, 0.327,0.612, 0.231,0.734, 0.25,0.914, 0.394,0.928, 0.462,0.978,
  0.596,0.95, 0.788,0.813, 0.808,0.662, 0.942,0.439, 0.971,0.324,
];

// Torso + legs of the little "people" pictograms that sit above each country name.
const FIG_M = [
  1,0.143, 0.713,0, 0.425,0, 0.402,0, 0.287,0, 0,0.143, 0,0.143, 0,0.435, 0.092,0.482,
  0.184,0.435, 0.184,0.262, 0.184,0.167, 0.23,0.167, 0.23,0.274, 0.23,0.458, 0.23,0.482,
  0.23,0.935, 0.356,1, 0.471,0.935, 0.471,0.482, 0.529,0.482, 0.529,0.935, 0.644,1, 0.77,0.935,
  0.77,0.482, 0.77,0.458, 0.77,0.274, 0.77,0.167, 0.816,0.167, 0.816,0.262, 0.816,0.435,
  0.908,0.482, 1,0.435,
];
const FIG_F = [
  0.989,0.383, 0.842,0.114, 0.832,0.108, 0.589,0, 0.516,0, 0.484,0, 0.411,0, 0.168,0.108,
  0.158,0.114, 0.011,0.383, 0.053,0.437, 0.158,0.413, 0.295,0.15, 0.337,0.15, 0.084,0.599,
  0.284,0.599, 0.284,0.946, 0.379,1, 0.463,0.946, 0.463,0.599, 0.537,0.599, 0.537,0.946,
  0.621,1, 0.716,0.946, 0.716,0.599, 0.916,0.599, 0.663,0.15, 0.705,0.15, 0.842,0.413,
  0.947,0.437,
];

/** Male + female pictogram pair; `x` is the left edge of the male figure. */
function people(s, x, y, color) {
  s.addShape('ellipse', { x: x + 0.061, y, w: 0.089, h: 0.088, fill: { color }, line: { type: 'none' } });
  polygon(s, x, y + 0.097, 0.21, 0.4, FIG_M, color);
  s.addShape('ellipse', { x: x + 0.483, y, w: 0.089, h: 0.088, fill: { color }, line: { type: 'none' } });
  polygon(s, x + 0.413, y + 0.095, 0.229, 0.398, FIG_F, color);
}

/* ---------------------------------------------------------------- slides */

/** 1 - cover: mist panel, tall photo, rotated captions, oversized wordmark. */
function slide01(s) {
  block(s, 0, 0, 4.905, 7.5, MIST);
  photo(s, 2.762, 1.095, 3.556, 5.31);
  // three caption pairs reading bottom-to-top down the left margin
  [
    { y: 5.415, subY: 6.009, subX: 1.364, subW: 0.545, sub: '30 slide' },
    { y: 3.492, subY: 3.985, subX: 1.293, subW: 0.687, sub: 'Lookbook' },
    { y: 1.569, subY: 1.929, subX: 1.161, subW: 0.952, sub: 'Studio design' },
  ].forEach(r => {
    s.addText('dapibus integer dolor', {
      x: 1.083, y: r.y, w: 1.695, h: 0.224, margin: 0, valign: 'top', rotate: 270, fit: 'resize',
      fontFace: SANS, fontSize: 10, color: GRAY,
    });
    s.addText(r.sub, {
      x: r.subX, y: r.subY, w: r.subW, h: 0.247, margin: 0, valign: 'top', wrap: false, rotate: 270,
      fit: 'resize', fontFace: SANS, fontSize: 11, bold: true, color: SLATE, lineSpacingMultiple: 1.5,
    });
  });
  block(s, 5.907, 2.056, 2.177, 3.389, PEACH);
  body(s, 8.827, 4.992, 3.708, 1.08, INTRO, { size: 11 });
  heading(s, 6.032, 3.077, 6.502, 1.346, 'FABULOUS', { size: 80, align: 'center' });
  footer(s);
}

/** 2 - statement page: wordmark left, full-bleed photo right. */
function slide02(s) {
  block(s, 11.0, 0, 2.331, 4.819, MIST);
  block(s, 0, -0.067, 2.177, 3.389, PEACH);
  heading(s, 1.571, 2.031, 4.879, 1.01, 'FABULOUS', { size: 60 });
  label(s, 1.571, 4.008, 1.704, 0.236, 'Creative Slide Here');
  body(s, 1.571, 4.459, 5.016, 1.481, IPSUM);
  photo(s, 7.947, 0.965, 4.401, 6.535);
  footer(s);
}

/** 3 - photo left, headline + stat card right. */
function slide03(s) {
  block(s, 0.732, 1.812, 4.542, 5.688, MIST);
  photo(s, 0, 0.948, 3.563, 5.344);
  heading(s, 4.825, 1.977, 2.43, 0.942, ['CREATIVE', 'LOOKBOOK']);
  body(s, 4.825, 3.193, 5.869, 0.875, LOREM);
  block(s, 10.325, 4.767, 3.01, 2.25, PEACH);
  heading(s, 10.88, 4.969, 1.278, 0.74, '2.1K', { size: 44, bold: true, fontFace: SANS });
  label(s, 10.881, 5.752, 1.013, 0.314, 'Description', { color: TEAL });
  body(s, 10.88, 6.092, 2.455, 0.572, 'asuscipit eros etus auctor dapibus quam aliquam',
    { color: TEAL, align: 'left' });
  label(s, 5.679, 4.767, 0.87, 0.236, 'Main Idea');
  body(s, 5.679, 5.218, 3.932, 0.875, IPSUM_TO_TINCIDUNT);
  footer(s);
}

/** 4 - "you need to know about us": tinted photo + peach corner. */
function slide04(s) {
  photo(s, 6.576, 2.908, 3.043, 4.11);
  s.addShape('rect', {
    x: 6.576, y: 2.907, w: 3.043, h: 4.11,
    fill: { color: TEAL, transparency: 20 }, line: { type: 'none' },
  });
  block(s, 9.125, 0, 4.21, 2.312, PEACH);
  photo(s, 9.826, 1.21, 3.509, 5.08);
  heading(s, 1.108, 2.072, 3.629, 0.942, ['YOU NEED TO ', 'KNOW ABOUT US'], { color: INK });
  label(s, 1.108, 3.495, 1.013, 0.236, 'Description');
  body(s, 1.108, 3.946, 4.656, 1.481, IPSUM);
  footer(s);
}

/** 5 - "proudly introduction": highlighted lead paragraph. */
function slide05(s) {
  block(s, -0.815, 2.347, 5.968, 4.338, PEACH, { rotate: 270 });
  photo(s, 0.708, 0.832, 4.208, 5.916);
  heading(s, 6.2, 2.13, 3.227, 0.942, ['PROUDLY', 'INTRODUCTION']);
  s.addText([
    { text: 'PLACEHOLDER', options: { color: PEACH, bold: true } },
    { text: 'PLACEHOLDER', options: { color: TEAL, bold: true } },
    { text: 'dapibus integer mollis da vel magna hendrerit eleifend nulla', options: { color: GRAY } },
  ], {
    x: 6.2, y: 3.346, w: 5.869, h: 0.875, margin: 0, valign: 'top', fit: 'resize',
    fontFace: SANS, fontSize: 12, align: 'justify', lineSpacingMultiple: 1.5,
  });
  body(s, 6.2, 4.495, 5.869, 0.875, LOREM);
  footer(s);
}

/** 6 - agenda list over a mist band. */
function slide06(s) {
  block(s, 0, 0.809, 11.486, 5.882, MIST);
  heading(s, 1.21, 1.722, 1.797, 0.942, ['LIST', 'AGENDA']);
  const agenda = ['01. Our profile', '02. Awesome team', '03. Album gallery', '04. Infographic'];
  s.addText(agenda.map((t, i) => ({ text: t, options: { breakLine: i < agenda.length - 1 } })), {
    x: 1.21, y: 3.569, w: 2.377, h: 1.963, margin: 0, valign: 'top', wrap: false, fit: 'resize',
    fontFace: SANS, fontSize: 20, bold: true, italic: true, color: SLATE, lineSpacingMultiple: 1.5,
  });
  body(s, 5.208, 1.604, 5.917, 1.178, IPSUM_TO_LACUS);
  photo(s, 5.208, 3.569, 8.127, 3.931);
  footer(s);
}

/** 7 - welcome speech: teal panel, portrait, signature rule. */
function slide07(s) {
  block(s, 11.741, 5.403, 1.594, 2.097, MIST);
  block(s, 0, 0, 4.542, 5.688, TEAL);
  photo(s, 1.708, 1.321, 3.563, 4.858);
  heading(s, 6.979, 1.851, 2.311, 0.942, ['WELCOME ', 'SPEECH']);
  label(s, 6.979, 3.37, 1.267, 0.236, 'David Bengald');
  s.addShape('line', { x: 6.979, y: 3.755, w: 1.869, h: 0, line: { color: PEACH, width: 1.25 } });
  body(s, 6.979, 3.821, 3.336, 0.247, 'asuscipit eros iste metus auctor idas',
    { size: 11, align: 'left', italic: true });
  body(s, 6.979, 4.471, 4.465, 1.178, LOREM);
  footer(s);
}

/** 8 / 16 / 21 / 26 - section dividers. */
function sectionSlide(s, title, titleW, sub, subW) {
  block(s, 0.839, 0, 1.575, 3.278, PEACH);
  heading(s, 0.839, 3.761, titleW, 0.471, title);
  label(s, 0.839, 4.47, subW, 0.269, sub, { size: 16 });
  body(s, 0.839, 4.977, 4.708, 1.178, LOREM);
  block(s, 9.57, 0, 3.902, 7.5, MIST);
  photo(s, 7.87, 1.188, 4.972, 5.83);
  footer(s);
}

/** 9 - about us: three stacked paragraphs between two colour blocks. */
function slide09(s) {
  block(s, 0, 1.812, 3.264, 5.688, MIST);
  block(s, 9.097, 3.181, 2.707, 4.319, PEACH);
  photo(s, 0, 0, 2.903, 6.535);
  heading(s, 10.085, 3.636, 2.442, 0.539, 'ABOUT US', { size: 32 });
  body(s, 3.852, 1.951, 4.656, 0.875, IPSUM_FROM_ELEMENTUM, { bold: true });
  body(s, 3.852, 3.165, 4.656, 1.481, IPSUM);
  body(s, 3.852, 4.984, 4.656, 0.875, IPSUM_TO_MOLLIS, { bold: true });
  footer(s);
}

/** 10 - company origins, photo columns either side. */
function slide10(s) {
  block(s, 0, 0, 2.554, 4.806, MIST);
  photo(s, 0.486, 1.792, 2.804, 3.917);
  photo(s, 10.531, 1.792, 2.804, 5.708);
  heading(s, 4.546, 1.925, 2.149, 0.942, ['COMPANY', 'ORIGINS']);
  label(s, 4.546, 3.3, 1.511, 0.236, '2022 Description');
  body(s, 4.546, 3.641, 4.708, 1.178, LOREM);
  body(s, 4.546, 5.003, 4.708, 0.572, LOREM_TO_ELEMENTUM);
  footer(s);
}

/** 11 - a few words about us: teal year tile overlapped by a photo. */
function slide11(s) {
  block(s, 10.193, 0.497, 3.639, 2.645, MIST, { rotate: 270 });
  block(s, 7.334, 2.347, 3.356, 3.292, TEAL);
  photo(s, 9.989, 1.292, 3.346, 6.208);
  block(s, -0.893, 0.893, 2.569, 0.783, PEACH, { rotate: 270 });
  heading(s, 7.889, 3.522, 0.969, 0.942, ['2023', 'Here'], { color: WHITE, align: 'center' });
  heading(s, 1.504, 1.925, 3.113, 0.942, ['A FEW WORDS', 'ABOUT US']);
  label(s, 1.504, 3.3, 1.013, 0.236, 'Description');
  body(s, 1.504, 3.641, 4.708, 1.178, LOREM);
  body(s, 1.504, 5.003, 4.708, 0.572, LOREM_TO_ELEMENTUM);
  footer(s);
}

/** 12 - introduce branding: full-height peach column. */
function slide12(s) {
  block(s, 9.744, 3.909, 4.444, 2.738, MIST, { rotate: 270 });
  block(s, 0, 0, 4.383, 7.5, PEACH);
  photo(s, 0.92, 1.647, 2.543, 4.206);
  photo(s, 10.792, 3.294, 2.543, 4.206);
  heading(s, 7.706, 1.537, 4.856, 0.471, 'INTRODUCE BRANDING');
  label(s, 5.136, 3.3, 1.013, 0.236, 'Description');
  body(s, 5.136, 3.641, 4.708, 1.178, LOREM);
  body(s, 5.136, 5.003, 4.708, 0.572, LOREM_TO_ELEMENTUM, { bold: true, italic: true });
  footer(s);
}

/** 13 - new style branded. */
function slide13(s) {
  block(s, -1.507, 3.424, 5.583, 2.569, MIST, { rotate: 270 });
  block(s, 9.614, 0, 3.722, 4.338, TEAL);
  photo(s, 7.977, 1.121, 3.308, 5.258);
  heading(s, 1.365, 1.925, 2.379, 0.942, ['NEW STYLE', 'BRANDED']);
  label(s, 1.365, 3.3, 1.013, 0.236, 'Description');
  body(s, 1.365, 3.641, 4.708, 1.178, LOREM);
  body(s, 1.365, 5.003, 4.708, 0.572, LOREM_TO_ELEMENTUM);
  footer(s);
}

/** 14 - see amazing stylish: vertical headline on the peach edge. */
function slide14(s) {
  block(s, 9.664, 0, 3.671, 7.5, PEACH);
  photo(s, 0, 1.682, 2.0, 4.136);
  photo(s, 7.92, 1.682, 3.218, 4.136);
  heading(s, 9.818, 3.514, 4.628, 0.471, 'SEE AMAZING STYLISH', { rotate: 90 });
  body(s, 2.606, 2.078, 4.708, 0.572, LOREM_TO_ELEMENTUM, { bold: true, italic: true });
  label(s, 2.606, 3.148, 1.013, 0.236, 'Description');
  body(s, 2.606, 3.488, 4.708, 1.178, LOREM);
  body(s, 2.606, 4.851, 4.708, 0.572, LOREM_TO_ELEMENTUM);
  footer(s, TEAL);
}

/** 15 - a new target. */
function slide15(s) {
  block(s, 0, 0, 2.542, 7.5, TEAL);
  photo(s, 1.655, 1.063, 3.547, 5.373);
  heading(s, 6.644, 2.102, 3.163, 0.471, 'A NEW TARGET');
  body(s, 6.644, 2.859, 4.494, 1.178, LOREM);
  body(s, 6.644, 4.22, 4.494, 1.178, LOREM);
  block(s, 12.58, 4.583, 0.755, 2.917, PEACH);
  footer(s);
}

/** 17 - influencer team: square portrait over a peach panel. */
function slide17(s) {
  block(s, 6.999, 3.667, 6.336, 3.833, PEACH);
  block(s, -0.005, 0, 3.154, 5.677, MIST);
  photo(s, 4.978, 3.675, 2.704, 2.704);
  photo(s, 1.078, 1.121, 3.308, 5.258);
  heading(s, 4.978, 1.963, 2.654, 0.942, ['INFLUENCER', 'TEAM']);
  label(s, 8.274, 4.225, 0.966, 0.236, 'Name here');
  body(s, 8.274, 4.566, 3.989, 1.178, LOREM);
  body(s, 8.274, 5.953, 3.989, 0.875, LOREM_TO_DAPIBUS);
  footer(s);
}

/** 18 - the designer: stacked portraits, rotated title, 90% stat. */
function slide18(s) {
  block(s, 2.32, 0, 3.917, 7.5, PEACH);
  photo(s, 0.486, 1.083, 2.882, 2.619);
  photo(s, 0.486, 4.196, 2.882, 2.619);
  heading(s, 3.317, 3.514, 3.056, 0.471, 'THE DESIGNER', { rotate: 270 });
  label(s, 7.481, 2.076, 0.847, 0.539, '90%', { size: 32 });
  label(s, 8.451, 2.146, 1.012, 0.314, 'Team Work', { lead: 1.5 });
  body(s, 7.477, 2.885, 4.494, 1.178, LOREM);
  body(s, 7.477, 4.246, 4.494, 1.178, LOREM);
  footer(s);
}

/** 19 - team expert: three portraits in a row. */
function slide19(s) {
  block(s, 0, 0, 3.792, 7.5, MIST);
  block(s, 11.036, 3.097, 2.299, 4.403, TEAL);
  photo(s, 2.867, 1.244, 2.519, 3.363);
  photo(s, 6.025, 1.244, 2.519, 3.363);
  photo(s, 9.183, 1.244, 3.499, 5.011);
  heading(s, 0.732, 5.396, 2.886, 0.471, 'TEAM EXPERT');
  body(s, 4.05, 5.418, 4.494, 0.875, LOREM_TO_DA);
  footer(s);
}

/** 20 - the leader: mist background, rotated title on the peach column. */
function slide20(s) {
  s.background = { color: MIST };
  block(s, 0, 0, 3.671, 7.5, PEACH);
  heading(s, 0.116, 3.481, 2.964, 0.539, 'THE LEADER', { size: 32, rotate: 270 });
  photo(s, 2.854, 0.81, 3.218, 5.881);
  label(s, 7.298, 2.386, 1.101, 0.269, 'Name here', { size: 16 });
  body(s, 7.298, 2.893, 4.708, 1.178, LOREM);
  body(s, 7.298, 4.239, 4.708, 0.875, LOREM_TO_MAGNA);
  footer(s);
}

/** 22 - album branded. */
function slide22(s) {
  block(s, 11.801, 2.972, 1.575, 4.528, PEACH);
  block(s, 0, 0, 3.902, 7.5, MIST);
  photo(s, 1.534, 0.81, 3.218, 5.881);
  photo(s, 9.392, 3.364, 3.218, 4.136);
  heading(s, 8.879, 1.537, 3.731, 0.471, 'ALBUM BRANDED');
  label(s, 5.296, 3.3, 1.013, 0.236, 'Description');
  body(s, 5.296, 3.641, 3.552, 1.178, LOREM_TO_VEL);
  body(s, 5.296, 5.003, 3.552, 0.875, LOREM_TO_ELEMENTUM);
  footer(s);
}

/** 23 - a new mockup: tablet + phone device mock-ups. */
function slide23(s) {
  block(s, 0, 0.528, 8.583, 6.444, MIST);
  deviceFrame(s, 7.05, 1.767, 2.95, 3.86);             // landscape tablet shell
  photo(s, 7.166, 1.894, 2.697, 3.615);
  s.addShape('rect', {
    x: 7.166, y: 1.894, w: 2.697, h: 3.615,
    fill: { color: TEAL, transparency: 20 }, line: { type: 'none' },
  });
  deviceFrame(s, 10.2, 1.85, 1.68, 3.77);              // phone shell
  photo(s, 10.271, 1.918, 1.547, 3.576, { round: true });
  heading(s, 1.259, 2.102, 3.415, 0.471, 'A NEW MOCKUP');
  body(s, 1.259, 2.859, 4.494, 1.178, LOREM);
  body(s, 1.259, 4.22, 4.494, 1.178, LOREM);
  block(s, 12.074, 1.894, 1.261, 3.615, PEACH);
  footer(s);
}

/** 24 - mockup layout: tilted tablet cradled by two hands. */
function slide24(s) {
  block(s, 0, 3.847, 3.671, 3.653, TEAL);
  // upper hand with the stylus, sweeping in from the top and out at the left edge
  blob(s, [
    1.83, 1.56, 2.60, 1.72, 2.90, 2.10, 2.75, 2.70, 2.35, 3.05, 1.55, 3.35,
    0.85, 3.75, 0.40, 4.30, 0.20, 4.90, 0.00, 5.20, 0.00, 3.10, 0.45, 2.65,
    0.95, 2.25, 1.40, 1.85,
  ], SKIN_24);
  s.addShape('line', { x: 1.78, y: 1.78, w: 1.97, h: 2.17, flipV: true, line: { color: 'F2F0F0', width: 5 } });
  // tilted tablet: black bezel with the screen placeholder inset
  polygon(s, 1.567, 2.06, 3.6, 4.27, [
    0.216, 0, 0.985, 0.113, 0.807, 0.985, 0.017, 0.878,
  ], '141414');
  polygon(s, 1.667, 2.105, 3.405, 4.166, [
    0.217, 0, 0.98, 0.114, 0.811, 0.984, 0.02, 0.886,
  ], PHOTO);
  s.addText('[image]', {
    x: 1.9, y: 3.6, w: 2.9, h: 0.4, margin: 0,
    align: 'center', valign: 'middle', fontFace: SANS, fontSize: 11, color: PHOTO_CAP,
  });
  // lower hand cupping the bottom-right corner of the tablet
  blob(s, [
    4.30, 5.65, 5.05, 5.80, 5.25, 6.40, 5.05, 7.10, 4.55, 7.50,
    3.35, 7.50, 3.40, 6.55, 3.95, 6.20, 4.10, 5.95,
  ], SKIN_24);
  heading(s, 6.842, 2.102, 3.65, 0.471, 'MOCKUP LAYOUT');
  body(s, 6.842, 2.859, 4.494, 1.178, LOREM);
  body(s, 6.842, 4.22, 4.494, 1.178, LOREM);
  footer(s);
  block(s, 7.644, 6.403, 5.691, 1.097, MIST);
}

/** 25 - new mockup branded style. */
function slide25(s) {
  block(s, 9.664, 3.847, 3.671, 3.653, PEACH);
  // two arms drop in from the top edge and hold the tablet between them
  blob(s, [
    6.62, -0.10, 7.32, -0.10, 7.36, 1.90, 7.30, 2.55, 7.12, 3.15, 6.95, 3.45,
    6.80, 3.05, 6.66, 2.30,
  ], SKIN_25);
  blob(s, [
    10.50, -0.10, 11.32, -0.10, 11.32, 2.60, 11.15, 3.60, 10.90, 4.55, 10.62, 5.15,
    10.48, 5.00, 10.42, 4.20, 10.36, 3.20, 10.38, 1.60,
  ], SKIN_25);
  deviceFrame(s, 7.27, 1.7, 3.35, 4.37);
  photo(s, 7.408, 1.821, 3.07, 4.103, { round: true });
  heading(s, 1.365, 1.925, 3.459, 0.942, ['NEW MOCKUP ', 'BRANDED STYLE']);
  label(s, 1.365, 3.3, 1.013, 0.236, 'Description');
  body(s, 1.365, 3.641, 4.708, 1.178, LOREM);
  body(s, 1.365, 5.003, 4.708, 0.572, LOREM_TO_ELEMENTUM);
  footer(s);
  block(s, 0, 0, 5.691, 1.097, MIST);
}

/** 27 / 28 - one-country infographic; the map sits on the opposite side. */
function countrySlide(s, opt) {
  heading(s, 5.206, 0.581, 2.922, 0.471, 'INFOGRAPHIC', { align: 'center' });
  footer(s);
  people(s, opt.textX - 0.044, 2.76, opt.color);
  heading(s, opt.textX, 3.605, opt.nameW, 0.337, opt.name, { size: 20, color: INK, fontFace: HEAD });
  body(s, opt.textX, 4.075, 4.367, 1.481, MAP_COPY);
  heading(s, opt.textX, 5.654, opt.pointW, 0.337, opt.point, { size: 20, color: INK, fontFace: HEAD });
  opt.map(s);
}

function chinaMap(s) {
  polygon(s, 0.876, 3.008, 4.952, 3.173, CHINA, TEAL);
  s.addShape('ellipse', { x: 4.628, y: 5.75, w: 0.154, h: 0.304, fill: { color: TEAL }, line: { type: 'none' } });
  s.addShape('ellipse', { x: 3.7, y: 6.195, w: 0.207, h: 0.16, fill: { color: TEAL }, line: { type: 'none' } });
}

/** 29 - three-country comparison with arrows between the percentages. */
function slide29(s) {
  heading(s, 5.206, 0.581, 2.922, 0.471, 'INFOGRAPHIC', { align: 'center' });
  footer(s);

  // Indonesia
  polygon(s, 2.336, 3.231, 0.729, 0.798, SUMATRA, TEAL);
  polygon(s, 3.016, 4.019, 0.623, 0.201, JAVA, TEAL);
  polygon(s, 3.244, 3.335, 0.685, 0.574, BORNEO, TEAL);
  polygon(s, 3.901, 3.511, 0.44, 0.504, SULAWESI, TEAL);
  polygon(s, 4.225, 4.168, 0.257, 0.167, NUSA, TEAL);
  polygon(s, 4.507, 3.808, 0.211, 0.081, MALUKU, TEAL);
  polygon(s, 4.669, 3.639, 0.727, 0.606, PAPUA, TEAL);
  // India
  polygon(s, 6.08, 2.861, 1.676, 1.711, INDIA, PEACH);
  // Japan
  polygon(s, 10.179, 2.863, 0.54, 0.512, HOKKAIDO, TEAL);
  polygon(s, 9.339, 3.359, 1.066, 0.896, HONSHU, TEAL);
  polygon(s, 9.465, 4.127, 0.257, 0.201, SHIKOKU, TEAL);
  polygon(s, 9.22, 4.181, 0.233, 0.312, KYUSHU, TEAL);

  [4.691, 7.739].forEach(x => {
    s.addShape('line', {
      x, y: 5.134, w: 1.37, h: 0,
      line: { color: GRAY, width: 1, endArrowType: 'triangle' },
    });
  });

  [
    { x: 3.619, w: 0.547, pct: '90%', capX: 2.72 },
    { x: 6.672, w: 0.547, pct: '80%', capX: 5.73 },
    { x: 9.736, w: 0.512, pct: '79%', capX: 8.741 },
  ].forEach(c => {
    s.addText(c.pct, {
      x: c.x, y: 4.829, w: c.w, h: 0.448, margin: 0, valign: 'top', wrap: false, fit: 'resize',
      fontFace: HEAD, fontSize: 20, color: INK, align: 'center', lineSpacingMultiple: 1.5,
    });
    body(s, c.capX, 5.417, 2.259, 0.678, 'adipiscing elit maecenas porttitor   volupat', { align: 'center' });
  });
}

/** 30 - thank you. */
function slide30(s) {
  block(s, 8.43, 0, 4.905, 7.5, MIST);
  body(s, 0.801, 4.7, 3.708, 1.08, INTRO, { size: 11 });
  block(s, 5.252, 2.056, 2.177, 3.389, PEACH);
  heading(s, 0.801, 3.077, 5.804, 1.111, 'THANK YOU', { size: 66 });
  footer(s);
  photo(s, 7.428, 1.095, 3.556, 5.31);
}

/* ------------------------------------------------------------------ build */

const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07,
  s => sectionSlide(s, 'SECTION ONE', 2.826, 'Company Profile', 1.685),
  slide09, slide10, slide11, slide12, slide13, slide14, slide15,
  s => sectionSlide(s, 'SECTION TWO', 2.931, 'Meet Our Team', 1.564),
  slide17, slide18, slide19, slide20,
  s => sectionSlide(s, 'SECTION THREE', 3.282, 'Album Gallery', 1.422),
  slide22, slide23, slide24, slide25,
  s => sectionSlide(s, 'SECTION FOUR', 3.094, 'Our Infographic', 1.578),
  s => countrySlide(s, {
    textX: 7.526, name: 'China country', nameW: 2.042, point: '89 Point', pointW: 1.211,
    color: TEAL, map: chinaMap,
  }),
  s => countrySlide(s, {
    textX: 1.816, name: 'India Country', nameW: 1.969, point: '88 Point', pointW: 1.218,
    color: PEACH, map: sl => polygon(sl, 7.897, 2.603, 3.665, 3.741, INDIA, PEACH),
  }),
  slide29, slide30,
];

const pptx = new PptxGenJS();
pptx.author = 'Ourslide';
pptx.title = 'Fabulous Lookbook';
pptx.defineLayout({ name: 'CUSTOM', width: W, height: H });
pptx.layout = 'CUSTOM';

BUILDERS.forEach(build => build(pptx.addSlide()));

pptx.writeFile({
  fileName: path.join(__dirname, '0c02ebdd-c7f5-4a50-94c5-3941e5da61eb_grok_final.pptx'),
}).then(f => console.log('wrote', f));
