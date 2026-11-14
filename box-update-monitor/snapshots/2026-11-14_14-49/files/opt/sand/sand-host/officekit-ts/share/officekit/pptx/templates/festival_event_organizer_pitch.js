// Recreation of the "Festival Event" template deck (30 slides, 16:9 / 10 x 5.625 in)
// using pptxgenjs only. The raster art in the original file is a set of solid
// #FF002A fills and a handful of tiny monochrome icons; both are redrawn here as
// plain shapes, so this script embeds no binary data and needs no asset files.
//
//   node <this file>   ->  writes the .pptx next to itself

const path = require('path');
const P = require('pptxgenjs');

const OUTPUT = '0dfebdd2-b271-4132-9704-3f23d7b88afa_grok_final.pptx';

const pptx = new P();
const SH = pptx.ShapeType;

// ---------------------------------------------------------------- theme palette
// "Festival Event" colour scheme: accent1 red, accent2 yellow, greys derived from
// the 7F7F7F theme colour via the luminance modifiers used in the original XML.
const RED   = 'D71B00';
const YEL   = 'FCC30B';
const W     = 'FFFFFF';
const BLACK = '000000';
const DARK  = '3F3F3F';
const GREY  = '7F7F7F';
const GREY7 = '5F5F5F';
const GREY5 = 'B2B2B2';
const GREY3 = 'CCCCCC';
const GREY1 = 'E5E5E5';
const REDD  = '6C0D00';
const PHOTO = 'FF002A';

const HEAD = 'Work Sans Black';   // theme major font
const BODY = 'Open Sans';         // theme minor font

// ---------------------------------------------------------------- shadows
// pptxgenjs mutates the shadow object it is given, so hand out a fresh copy.
const GLOW = function () { return { type: 'outer', blur: 18, offset: 0, angle: 0, color: BLACK, opacity: 0.44 }; }
const CARD = function () { return { type: 'outer', blur: 16, offset: 3, angle: 225, color: BLACK, opacity: 0.32 }; }
const DROP = function () { return { type: 'outer', blur: 16, offset: 3, angle: 135, color: BLACK, opacity: 0.36 }; }
const GLOW2 = function () { return { type: 'outer', blur: 16, offset: 0, angle: 0, color: BLACK, opacity: 0.36 }; }
const DROP3 = function () { return { type: 'outer', blur: 16, offset: 3, angle: 135, color: BLACK, opacity: 0.54 }; }
const SHDW6 = function () { return { type: 'outer', blur: 13, offset: 3, angle: 135, color: BLACK, opacity: 0.31 }; }
const HALO = function () { return { type: 'outer', blur: 21, offset: 0, angle: 0, color: BLACK, opacity: 0.22 }; }
const WIDE = function () { return { type: 'outer', blur: 53, offset: 1, angle: 0, color: BLACK, opacity: 0.35 }; }

// ---------------------------------------------------------------- outlines
// Every free-form silhouette in the deck as a normalised (0..1) path:
//   'M',x,y | 'L',x,y | 'C',x1,y1,x2,y2,x,y | 'Z'
// SPLASH  - the yellow corner blob repeated on almost every slide
// SHAPE*  - decorative red/yellow blobs
// FRAME*  - outlines of the picture frames (drawn as flat photo placeholders)
const SPLASH = [
  'M',0.011,0.845, 'C',0.003,0.821,-0.001,0.793,0,0.761, 'C',0.012,0.503,0.252,0.856,0.537,0.224,
  'C',0.626,0.027,0.712,-0.024,0.788,0.009, 'L',0.804,0.02, 'L',1,0.797, 'L',0.195,1, 'L',0.152,0.981,
  'C',0.083,0.947,0.032,0.907,0.011,0.845, 'Z'
];
const FRAME = ['M',0,0, 'L',1,0, 'L',1,1, 'L',0,1, 'Z'];
const SHAPE = [
  'M',0.877,0, 'L',1,0.819, 'L',0,1, 'L',0,0.97, 'C',0.015,0.628,0.33,1.096,0.703,0.259,
  'C',0.762,0.128,0.819,0.046,0.874,0.002, 'Z'
];
const SHAPE1 = [
  'M',0,0.616, 'C',0.011,0.407,0.235,0.694,0.5,0.182, 'C',0.766,-0.33,1,0.376,1,0.616,
  'C',1,0.856,0.916,1.102,0.605,0.957, 'C',0.294,0.812,-0.01,0.825,0,0.616, 'Z'
];
const FRAME1 = [
  'M',0.5,0, 'C',0.776,0,1,0.224,1,0.5, 'C',1,0.776,0.776,1,0.5,1, 'C',0.224,1,0,0.776,0,0.5,
  'C',0,0.224,0.224,0,0.5,0, 'Z'
];
const SHAPE2 = [
  'M',0,0.61, 'C',0.011,0.398,0.254,0.512,0.5,0.169, 'C',0.846,-0.313,1,0.366,1,0.61,
  'C',1,0.854,0.916,1.103,0.605,0.956, 'C',0.294,0.809,-0.01,0.823,0,0.61, 'Z'
];
const FRAME2 = [
  'M',0.057,0, 'L',0.943,0, 'C',0.975,0,1,0.043,1,0.096, 'L',1,0.904, 'C',1,0.957,0.975,1,0.943,1, 'L',0.057,1,
  'C',0.025,1,0,0.957,0,0.904, 'L',0,0.096, 'C',0,0.043,0.025,0,0.057,0, 'Z'
];
const FRAME3 = [
  'M',0.085,0, 'L',0.915,0, 'C',0.962,0,1,0.043,1,0.096, 'L',1,0.904, 'C',1,0.957,0.962,1,0.915,1, 'L',0.085,1,
  'C',0.038,1,0,0.957,0,0.904, 'L',0,0.096, 'C',0,0.043,0.038,0,0.085,0, 'Z'
];
const SHAPE3 = [
  'M',0,0.301, 'L',0.101,1, 'L',0.985,0.926, 'L',0.993,0.869, 'C',0.998,0.828,1,0.785,1,0.743,
  'C',1,0.404,0.573,-0.287,0.193,0.13, 'C',0.134,0.195,0.073,0.248,0.014,0.291, 'Z'
];
const FRAME4 = [
  'M',0.706,0, 'C',0.81,0,0.883,0.058,0.873,0.213, 'C',0.846,0.626,1.099,0.721,0.956,0.917,
  'C',0.814,1.112,0.582,0.911,0.231,0.848, 'C',-0.12,0.786,-0.01,0.512,0.163,0.295,
  'C',0.271,0.159,0.532,0,0.706,0, 'Z'
];
const SHAPE4 = [
  'M',0,0.579, 'C',0.011,0.349,0.265,0.426,0.5,0.101, 'C',0.735,-0.224,1,0.315,1,0.579,
  'C',1,0.842,0.916,1.112,0.605,0.952, 'C',0.294,0.793,-0.01,0.808,0,0.579, 'Z'
];
const FRAME5 = [
  'M',0.036,0, 'L',0.964,0, 'C',0.984,0,1,0.016,1,0.036, 'L',1,0.964, 'C',1,0.984,0.984,1,0.964,1, 'L',0.036,1,
  'C',0.016,1,0,0.984,0,0.964, 'L',0,0.036, 'C',0,0.016,0.016,0,0.036,0, 'Z'
];
const FRAME6 = [
  'M',0.096,0, 'L',0.904,0, 'C',0.957,0,1,0.036,1,0.08, 'L',1,0.92, 'C',1,0.964,0.957,1,0.904,1, 'L',0.096,1,
  'C',0.043,1,0,0.964,0,0.92, 'L',0,0.08, 'C',0,0.036,0.043,0,0.096,0, 'Z'
];
const FRAME7 = [
  'M',0.09,0, 'L',1,0, 'L',1,1, 'L',0.09,1, 'C',0.04,1,0,0.957,0,0.904, 'L',0,0.096, 'C',0,0.043,0.04,0,0.09,0,
  'Z'
];
const SHAPE5 = [
  'M',0.005,0.636, 'C',0.003,0.65,0.001,0.666,0,0.683, 'C',-0.003,0.767,0.024,0.824,0.073,0.867,
  'L',0.073,0.867, 'L',0.959,1, 'L',0.971,0.961, 'C',0.992,0.878,1,0.78,1,0.683,
  'C',1,0.372,0.735,-0.264,0.5,0.119, 'C',0.28,0.479,0.043,0.421,0.005,0.636, 'Z'
];
const FRAME8 = [
  'M',0.036,0, 'L',0.964,0, 'C',0.984,0,1,0.013,1,0.029, 'L',1,0.971, 'C',1,0.987,0.984,1,0.964,1, 'L',0.036,1,
  'C',0.016,1,0,0.987,0,0.971, 'L',0,0.029, 'C',0,0.013,0.016,0,0.036,0, 'Z'
];
const FRAME9 = [
  'M',0.036,0, 'L',0.964,0, 'C',0.984,0,1,0.013,1,0.028, 'L',1,0.972, 'C',1,0.987,0.984,1,0.964,1, 'L',0.036,1,
  'C',0.016,1,0,0.987,0,0.972, 'L',0,0.028, 'C',0,0.013,0.016,0,0.036,0, 'Z'
];
const SHAPE6 = [
  'M',0.955,0.002, 'C',0.964,0.003,0.974,0.005,0.983,0.008, 'L',1,0.014, 'L',0.915,1, 'L',0.885,0.994,
  'C',0.861,0.987,0.837,0.98,0.811,0.971, 'C',0.394,0.824,-0.014,0.838,0,0.625,
  'C',0.015,0.413,0.315,0.704,0.67,0.184, 'C',0.77,0.038,0.867,-0.01,0.955,0.002, 'Z'
];
const SHAPE7 = [
  'M',1,0.338, 'L',0.727,1, 'L',0,0.703, 'L',0.001,0.702, 'C',0.095,0.549,0.321,0.536,0.534,0.153,
  'C',0.696,-0.139,0.87,0.03,0.984,0.298, 'Z'
];
const FRAME10 = [
  'M',0.01,0, 'L',0.99,0, 'C',0.996,0,1,0.016,1,0.036, 'L',1,0.964, 'C',1,0.984,0.996,1,0.99,1, 'L',0.01,1,
  'C',0.004,1,0,0.984,0,0.964, 'L',0,0.036, 'C',0,0.016,0.004,0,0.01,0, 'Z'
];
const FRAME11 = ['M',0.032,0, 'L',1,0, 'L',1,1, 'L',0,1, 'L',0,0.041, 'C',0,0.018,0.014,0,0.032,0, 'Z'];
const SHAPE8 = ['M',0,1, 'L',1,0, 'L',1,0.541, 'C',1,0.795,0.837,1,0.636,1, 'Z'];
const SHAPE9 = [
  'M',0.086,0.642, 'C',-0.1,0.615,0.038,0.357,0.292,0.292, 'C',0.546,0.228,0.189,0.051,0.514,0.004,
  'C',0.838,-0.043,0.754,0.353,0.944,0.502, 'C',1.14,0.656,0.78,1.009,0.514,1,
  'C',0.248,0.99,0.272,0.669,0.086,0.642, 'Z'
];
const FRAME12 = [
  'M',0.539,0, 'C',0.77,-0.002,0.749,0.339,0.933,0.451, 'C',1.143,0.579,0.818,0.974,0.552,0.999,
  'C',0.287,1.023,0.281,0.702,0.094,0.699, 'C',-0.094,0.696,0.019,0.422,0.265,0.326,
  'C',0.512,0.23,0.141,0.1,0.459,0.012, 'C',0.489,0.004,0.515,0,0.539,0, 'Z'
];
const FRAME13 = [
  'M',0.024,0, 'L',0.976,0, 'C',0.989,0,1,0.027,1,0.06, 'L',1,0.94, 'C',1,0.973,0.989,1,0.976,1, 'L',0.024,1,
  'C',0.011,1,0,0.973,0,0.94, 'L',0,0.06, 'C',0,0.027,0.011,0,0.024,0, 'Z'
];
const SHAPE10 = [
  'M',0.854,1, 'L',0,0.462, 'L',0.008,0.458, 'C',0.092,0.407,0.184,0.325,0.279,0.191,
  'C',0.662,-0.347,1,0.395,1,0.648, 'C',1,0.774,0.97,0.901,0.883,0.978, 'Z'
];
const SHAPE11 = [
  'M',1,0.457, 'L',0.44,1, 'L',0.391,0.983, 'C',0.167,0.907,-0.008,0.872,0,0.695,
  'C',0.011,0.459,0.243,0.782,0.517,0.205, 'C',0.723,-0.228,0.911,0.112,0.992,0.423, 'Z'
];
const FRAME14 = [
  'M',0.209,0, 'L',1,0, 'L',0.999,0.032, 'C',0.959,0.505,1.076,0.882,0.862,0.985,
  'C',0.648,1.087,0.86,0.645,0.255,0.595, 'C',-0.123,0.564,-0.019,0.283,0.172,0.044, 'Z'
];
const FRAME15 = [
  'M',0.466,0.731, 'C',0.485,0.731,0.504,0.735,0.523,0.742, 'C',0.598,0.771,0.632,0.85,0.6,0.919,
  'C',0.568,0.987,0.482,1.018,0.407,0.989, 'C',0.333,0.96,0.298,0.881,0.33,0.813,
  'C',0.354,0.762,0.409,0.731,0.466,0.731, 'Z', 'M',0.689,0, 'C',0.802,-0.002,0.891,0.057,0.907,0.231,
  'C',0.935,0.54,1.1,0.767,0.91,0.86, 'C',0.719,0.953,0.86,0.642,0.273,0.686,
  'C',-0.313,0.729,0.194,0.208,0.407,0.091, 'C',0.5,0.041,0.602,0.002,0.689,0, 'Z'
];
const FRAME16 = [
  'M',0.407,0, 'C',0.56,0,0.788,0.102,0.973,0.212, 'L',1,0.229, 'L',1,0.91, 'L',0.975,0.912,
  'C',0.52,0.954,0.199,1.088,0.036,0.91, 'C',-0.138,0.72,0.387,0.86,0.248,0.274, 'C',0.2,0.072,0.277,0,0.407,0,
  'Z'
];
const SHAPE12 = ['M',0,0, 'L',0.973,0, 'C',0.988,0,1,0.02,1,0.044, 'L',1,1, 'L',0,1, 'Z'];
const FRAME17 = [
  'M',0.032,0, 'L',0.968,0, 'C',0.986,0,1,0.016,1,0.036, 'L',1,0.964, 'C',1,0.984,0.986,1,0.968,1, 'L',0.032,1,
  'C',0.014,1,0,0.984,0,0.964, 'L',0,0.036, 'C',0,0.016,0.014,0,0.032,0, 'Z'
];
const SHAPE13 = [
  'M',0.046,0, 'L',1,0, 'L',1,0.745, 'L',0.968,0.79, 'C',0.865,0.92,0.724,1,0.567,1,
  'C',0.254,1,0,0.679,0,0.283, 'C',0,0.184,0.016,0.089,0.045,0.004, 'Z'
];
const SHAPE14 = ['M',0.757,0, 'L',1,1, 'L',0.18,1, 'C',0.081,1,0,0.776,0,0.5, 'C',0,0.224,0.081,0,0.18,0, 'Z'];
const SHAPE15 = [
  'M',0.978,0.188, 'C',0.992,0.257,1,0.339,1,0.427, 'L',1,0.573, 'C',1,0.602,0.999,0.631,0.997,0.659,
  'L',0.99,0.735, 'L',0.87,1, 'L',0.869,1, 'L',0.131,1, 'C',0.059,1,0,0.809,0,0.573, 'L',0,0.427,
  'C',0,0.191,0.059,0,0.131,0, 'L',0.869,0, 'C',0.914,0,0.954,0.075,0.978,0.188, 'Z'
];
const FRAME18 = [
  'M',0.021,0, 'L',0.979,0, 'C',0.991,0,1,0.016,1,0.036, 'L',1,0.964, 'C',1,0.984,0.991,1,0.979,1, 'L',0.021,1,
  'C',0.009,1,0,0.984,0,0.964, 'L',0,0.036, 'C',0,0.016,0.009,0,0.021,0, 'Z'
];
const FRAME19 = [
  'M',0.036,0, 'L',0.964,0, 'C',0.984,0,1,0.016,1,0.035, 'L',1,0.965, 'C',1,0.984,0.984,1,0.964,1, 'L',0.036,1,
  'C',0.016,1,0,0.984,0,0.965, 'L',0,0.035, 'C',0,0.016,0.016,0,0.036,0, 'Z'
];
const SHAPE16 = [
  'M',0.376,1, 'C',0.343,1,0.314,0.986,0.292,0.961, 'C',0.033,0.627,0.033,0.627,0.033,0.627,
  'C',0,0.585,0.007,0.525,0.051,0.489, 'C',0.099,0.458,0.161,0.468,0.193,0.511,
  'C',0.365,0.729,0.365,0.729,0.365,0.729, 'C',0.799,0.06,0.799,0.06,0.799,0.06,
  'C',0.828,0.014,0.891,0,0.938,0.028, 'C',0.985,0.056,1,0.116,0.971,0.165,
  'C',0.46,0.954,0.46,0.954,0.46,0.954, 'C',0.442,0.982,0.412,0.996,0.38,1, 'C',0.38,1,0.376,1,0.376,1, 'Z'
];
const FRAME20 = [
  'M',0.033,0, 'L',0.967,0, 'C',0.985,0,1,0.016,1,0.036, 'L',1,0.964, 'C',1,0.984,0.985,1,0.967,1, 'L',0.033,1,
  'C',0.015,1,0,0.984,0,0.964, 'L',0,0.036, 'C',0,0.016,0.015,0,0.033,0, 'Z'
];
const SHAPE17 = [
  'M',0.994,0.809, 'C',0.986,0.877,0.969,0.941,0.939,0.995, 'L',0.936,1, 'L',0,0.947, 'L',0.053,0.485,
  'L',0.062,0.478, 'C',0.136,0.421,0.215,0.337,0.296,0.208, 'C',0.67,-0.378,1,0.432,1,0.707,
  'C',1,0.741,0.998,0.776,0.994,0.809, 'Z'
];
const FRAME21 = [
  'M',0,0, 'L',0.697,0, 'L',0.698,0, 'C',0.922,0.116,1.078,0.271,0.959,0.311,
  'C',0.749,0.383,0.809,0.922,0.442,0.987, 'C',0.293,1.014,0.144,0.997,0.015,0.948, 'L',0,0.942, 'Z'
];
const SHAPE18 = [
  'M',1,0.15, 'L',0.366,0, 'L',0,1, 'L',0.041,0.994, 'C',0.502,0.911,0.624,0.242,0.927,0.198,
  'C',0.957,0.193,0.978,0.183,0.991,0.167, 'Z'
];
const SHAPE19 = [
  'M',1,0.071, 'L',0.913,1, 'L',0.102,0.943, 'L',0.091,0.937, 'C',0.03,0.895,-0.004,0.842,0,0.761,
  'C',0.014,0.503,0.292,0.856,0.623,0.224, 'C',0.757,-0.032,0.885,-0.041,0.989,0.06, 'Z'
];
const SHAPE20 = [
  'M',1,0.464, 'L',0.733,1, 'L',0,0.671, 'L',0.002,0.659, 'C',0.038,0.497,0.257,0.748,0.515,0.206,
  'C',0.73,-0.246,0.925,0.142,1,0.463, 'Z'
];
const SHAPE21 = [
  'M',0.5,0, 'L',0.5,0, 'L',0.398,0.003, 'L',0.306,0.023, 'L',0.224,0.051, 'L',0.148,0.087, 'L',0.082,0.132,
  'L',0.036,0.19, 'L',0.01,0.251, 'L',0,0.315, 'L',0,0.315, 'L',0,0.373, 'L',0.02,0.431, 'L',0.046,0.495,
  'L',0.077,0.553, 'L',0.158,0.669, 'L',0.25,0.775, 'L',0.342,0.868, 'L',0.418,0.936, 'L',0.5,1, 'L',0.5,1,
  'L',0.571,0.936, 'L',0.658,0.868, 'L',0.75,0.775, 'L',0.842,0.669, 'L',0.918,0.553, 'L',0.954,0.495,
  'L',0.969,0.431, 'L',0.99,0.373, 'L',1,0.315, 'L',1,0.315, 'L',0.99,0.251, 'L',0.954,0.19, 'L',0.908,0.132,
  'L',0.852,0.087, 'L',0.776,0.051, 'L',0.694,0.023, 'L',0.602,0.003, 'L',0.5,0, 'L',0.5,0, 'Z', 'M',0.5,0.505,
  'L',0.5,0.505, 'L',0.434,0.502, 'L',0.378,0.489, 'L',0.327,0.469, 'L',0.276,0.447, 'L',0.24,0.418,
  'L',0.214,0.383, 'L',0.194,0.347, 'L',0.194,0.315, 'L',0.194,0.315, 'L',0.194,0.273, 'L',0.214,0.238,
  'L',0.24,0.203, 'L',0.276,0.174, 'L',0.327,0.151, 'L',0.378,0.132, 'L',0.434,0.122, 'L',0.5,0.116,
  'L',0.5,0.116, 'L',0.556,0.122, 'L',0.622,0.132, 'L',0.668,0.151, 'L',0.714,0.174, 'L',0.75,0.203,
  'L',0.776,0.238, 'L',0.796,0.273, 'L',0.806,0.315, 'L',0.806,0.315, 'L',0.796,0.347, 'L',0.776,0.383,
  'L',0.75,0.418, 'L',0.714,0.447, 'L',0.668,0.469, 'L',0.622,0.489, 'L',0.556,0.502, 'L',0.5,0.505,
  'L',0.5,0.505, 'Z', 'M',0.306,0.315, 'L',0.306,0.315, 'L',0.306,0.338, 'L',0.316,0.36, 'L',0.332,0.376,
  'L',0.362,0.395, 'L',0.388,0.412, 'L',0.423,0.424, 'L',0.454,0.431, 'L',0.5,0.431, 'L',0.5,0.431,
  'L',0.536,0.431, 'L',0.571,0.424, 'L',0.602,0.412, 'L',0.638,0.395, 'L',0.658,0.376, 'L',0.673,0.36,
  'L',0.684,0.338, 'L',0.694,0.315, 'L',0.694,0.315, 'L',0.684,0.283, 'L',0.673,0.26, 'L',0.658,0.244,
  'L',0.638,0.225, 'L',0.602,0.209, 'L',0.571,0.196, 'L',0.536,0.19, 'L',0.5,0.19, 'L',0.5,0.19,
  'L',0.454,0.19, 'L',0.423,0.196, 'L',0.388,0.209, 'L',0.362,0.225, 'L',0.332,0.244, 'L',0.316,0.26,
  'L',0.306,0.283, 'L',0.306,0.315, 'L',0.306,0.315, 'Z'
];
const SHAPE22 = [
  'M',0.579,0.58, 'L',0.579,0.58, 'C',0.5,0.659,0.399,0.738,0.36,0.7, 'C',0.3,0.639,0.259,0.598,0.14,0.7,
  'C',0,0.799,0.099,0.878,0.16,0.919, 'C',0.218,0.998,0.459,0.939,0.698,0.7,
  'C',0.937,0.46,0.998,0.219,0.937,0.138, 'C',0.878,0.079,0.818,0,0.718,0.12,
  'C',0.619,0.239,0.66,0.278,0.718,0.341, 'C',0.759,0.377,0.68,0.479,0.579,0.58
];
const SHAPE23 = [
  'M',0.767,0, 'L',0.767,0, 'C',0.918,0,0.957,0.054,0.957,0.106, 'C',0.957,0.161,0.845,0.231,0.69,0.231,
  'C',0.539,0.231,0.461,0.195,0.461,0.124, 'C',0.461,0.07,0.578,0,0.767,0, 'Z', 'M',0.31,0.998, 'L',0.31,0.998,
  'C',0.194,0.998,0.121,0.96,0.194,0.817, 'C',0.349,0.57,0.349,0.57,0.349,0.57,
  'C',0.349,0.534,0.349,0.516,0.349,0.516, 'C',0.31,0.516,0.159,0.552,0.082,0.57, 'C',0,0.534,0,0.534,0,0.534,
  'C',0.272,0.428,0.616,0.357,0.728,0.357, 'C',0.845,0.357,0.884,0.428,0.806,0.516,
  'C',0.651,0.783,0.651,0.783,0.651,0.783, 'C',0.651,0.835,0.651,0.853,0.69,0.853,
  'C',0.69,0.853,0.806,0.817,0.918,0.783, 'C',0.996,0.817,0.996,0.817,0.996,0.817,
  'C',0.728,0.96,0.422,0.998,0.31,0.998, 'Z'
];

// ---------------------------------------------------------------- helpers

// Expand a normalised outline into the point list pptxgenjs custGeom expects.
function pts(outline, w, h) {
  const out = [];
  for (let i = 0; i < outline.length; ) {
    const op = outline[i++];
    if (op === 'Z') { out.push({ close: true }); continue; }
    if (op === 'M') { out.push({ x: outline[i++] * w, y: outline[i++] * h, moveTo: true }); continue; }
    if (op === 'L') { out.push({ x: outline[i++] * w, y: outline[i++] * h }); continue; }
    const x1 = outline[i++] * w, y1 = outline[i++] * h;
    const x2 = outline[i++] * w, y2 = outline[i++] * h;
    out.push({ x: outline[i++] * w, y: outline[i++] * h, curve: { type: 'cubic', x1: x1, y1: y1, x2: x2, y2: y2 } });
  }
  return out;
}

// Free-form decorative shape.
function blob(s, outline, x, y, w, h, o) {
  s.addShape('custGeom', Object.assign({ x: x, y: y, w: w, h: h, points: pts(outline, w, h) }, o));
}

// Photo placeholder: original frame silhouette filled flat (source art is solid #FF002A).
function photo(s, outline, x, y, w, h, o) {
  blob(s, outline, x, y, w, h, Object.assign({ fill: { color: PHOTO } }, o));
}

// Tiny pictogram (chevron / map pin / phone / fan icons in the original artwork),
// stood in for by a small red disc so the composition still reads correctly.
function glyph(s, x, y, w, h) {
  s.addShape(SH.ellipse, { x: x + w * 0.22, y: y + h * 0.22, w: w * 0.56, h: h * 0.56, fill: { color: RED } });
}

// Linear gradient stand-in: pptxgenjs paints solid fills only, so a gradient is
// approximated by a run of opaque bands stepping between the colour stops.
// `stops` is [[position%, hexColour], ...] and `ang` is the OOXML angle in
// degrees, measured clockwise from the +x axis.
function fade(s, x, y, w, h, stops, ang) {
  const BANDS = 40;
  const rad = (ang * Math.PI) / 180, dx = Math.cos(rad), dy = Math.sin(rad);
  const along = Math.abs(w * dx) + Math.abs(h * dy);
  const across = Math.abs(w * dy) + Math.abs(h * dx);
  const cx = x + w / 2, cy = y + h / 2, step = along / BANDS;
  for (let i = 0; i < BANDS; i++) {
    const u = ((i + 0.5) / BANDS) * 100;
    s.addShape(SH.rect, {
      x: cx + dx * (u / 100 - 0.5) * along - step * 0.75,
      y: cy + dy * (u / 100 - 0.5) * along - across / 2,
      w: step * 1.5, h: across, rotate: ang,
      fill: { color: rampAt(stops, u) },
    });
  }
}

// Linear interpolation between the colour stops of a gradient.
function rampAt(stops, u) {
  let a = stops[0], b = stops[stops.length - 1];
  for (let k = 0; k < stops.length - 1; k++) {
    if (u <= stops[k + 1][0]) { a = stops[k]; b = stops[k + 1]; break; }
  }
  const f = b[0] === a[0] ? 0 : Math.max(0, Math.min(1, (u - a[0]) / (b[0] - a[0])));
  let out = '';
  for (let c = 0; c < 6; c += 2) {
    const p = parseInt(a[1].substr(c, 2), 16), q = parseInt(b[1].substr(c, 2), 16);
    out += ('0' + Math.round(p + (q - p) * f).toString(16)).slice(-2);
  }
  return out.toUpperCase();
}

function oval(s, x, y, w, h, o)  { s.addShape(SH.ellipse, Object.assign({ x: x, y: y, w: w, h: h }, o)); }
function box(s, x, y, w, h, o)   { s.addShape(SH.rect, Object.assign({ x: x, y: y, w: w, h: h }, o)); }
function panel(s, x, y, w, h, o) { s.addShape(SH.roundRect, Object.assign({ x: x, y: y, w: w, h: h }, o)); }
function rule(s, x, y, w, h, o)  { s.addShape(SH.line, { x: x, y: y, w: w, h: h, line: o }); }

// White content card with the deck's standard soft drop shadow.
function card(s, x, y, w, h, o) {
  panel(s, x, y, w, h, Object.assign({ fill: { color: W }, shadow: CARD() }, o));
}

// Plain text box; `o` overrides the deck default of 8.25pt grey Open Sans.
function txt(s, text, x, y, w, h, o) {
  s.addText(text, Object.assign({ x: x, y: y, w: w, h: h, fontFace: BODY, fontSize: 8.25,
    color: GREY, valign: 'top', bullet: false }, o));
}

// Justified body copy at 1.5 line spacing - the deck's most common paragraph.
function body(s, text, x, y, w, h, o) {
  txt(s, text, x, y, w, h, Object.assign({ align: 'justify', lineSpacingMultiple: 1.5 }, o));
}

// Section heading: 24pt Work Sans Black.
function head(s, text, x, y, w, h, o) {
  txt(s, text, x, y, w, h, Object.assign({ fontFace: HEAD, fontSize: 24, color: BLACK }, o));
}

// Fully rounded pill button with a centred label.
function btn(s, label, x, y, w, h, o) {
  s.addText(label, Object.assign({ x: x, y: y, w: w, h: h, shape: SH.roundRect,
    rectRadius: Math.min(w, h) / 2, fontFace: BODY, fontSize: 8.25, color: W, align: 'center',
    valign: 'middle', bullet: false, shadow: DROP() }, o));
}

// ---------------------------------------------------------------- page furniture

const BRAND = { x: 0.295, y: 0.333 };   // brand mark origin (moves on the dark slides)
function brandAt(x, y) { BRAND.x = x; BRAND.y = y; }

// Brand mark, "Follow Us" + social pill, copyright line and the yellow corner
// splash - the elements the original repeats through its slide layouts.
function chrome(s, o) {
  if (o.brand) {
    const split = Array.isArray(o.brand);
    s.addText(split ? [{ text: 'Festival', options: { color: o.brand[0] } },
                       { text: ' Event ', options: { color: o.brand[1] } }] : 'Festival Event ',
      { x: BRAND.x, y: BRAND.y, w: 1.198, h: 0.24, fontFace: HEAD, fontSize: 8.25,
        color: split ? W : o.brand, valign: 'top', bullet: false });
  }
  if (o.link) {
    txt(s, 'Follow Us', 7.691, 0.337, 1.03, 0.252, { fontSize: 9, color: o.link, align: 'right' });
    socialPill(s, 8.764, 0.329);
  }
  if (o.foot) {
    txt(s, '\u00a92021 Festival Event. All Rights Reserved', 5.649, 5.122, 4.083, 0.24,
      { color: o.foot, align: 'right' });
  }
  if (o.splash !== false) {
    const b = o.splash || [-0.106, 4.54, 1.189, 1.19, 104.1];
    const d = o.dot || [0.651, 4.709, 293.2];
    blob(s, SPLASH, b[0], b[1], b[2], b[3], { rotate: b[4], fill: { color: YEL }, shadow: GLOW() });
    oval(s, d[0], d[1], 0.284, 0.284, { rotate: d[2], fill: { color: YEL } });
  }
}

// Rounded yellow pill holding four white social badges
// (twitter / facebook / instagram / linkedin in the original artwork).
function socialPill(s, x, y) {
  panel(s, x, y, 0.896, 0.243, { rectRadius: 0.121, fill: { color: YEL } });
  [0.089, 0.282, 0.466, 0.658].forEach(function (dx) {
    oval(s, x + dx, y + 0.068, 0.107, 0.107, { fill: { color: W } });
  });
}

// ---------------------------------------------------------------- slides

// 1. Cover
function slide01(s) {
  photo(s, FRAME, 0, 0, 10, 5.625);
  fade(s, 0, 0, 10, 5.625, [[0, 'D71B00'], [99, 'F1091B']], 315);
  brandAt(0.41, 0.469);
  chrome(s, { brand: W, link: W, foot: W });
  blob(s, SHAPE, 6.218, 2.326, 3.998, 3.648, { rotate: 9.4, fill: { color: YEL }, shadow: GLOW() });
  oval(s, 3.417, 1.229, 3.167, 3.167, { line: { color: YEL, width: 2.25, transparency: 51, dashType: 'sysDot' } });
  head(s, 'ORGANIZE YOUR FESTIVAL WITH US', 2.031, 2.091, 5.938, 1.313, { fontSize: 36, color: W, align: 'center' });
  body(s, 'Festival Event Presentation', 2.969, 1.83, 4.062, 0.287, { color: W, align: 'center' });
  btn(s, 'Start Present', 4.319, 3.529, 1.363, 0.267, { fontSize: 9, color: RED, bold: true, rotate: 180, flipV: true, fill: { color: W }, shadow: CARD() });
}

// 2. Introduction
function slide02(s) {
  photo(s, FRAME, 0, 0, 10, 5.625);
  box(s, 0, 0, 10, 5.625, { fill: { color: RED, transparency: 34 } });
  oval(s, 5.904, 1.636, 2.701, 2.701, { fill: { color: W } });
  blob(s, SHAPE1, 6.853, 1.05, 1.955, 2.249, { rotate: 347, flipH: true, fill: { color: W } });
  photo(s, FRAME1, 5.989, 1.722, 2.53, 2.53);
  brandAt(0.41, 0.469);
  chrome(s, { brand: W, link: W, foot: W });
  oval(s, 8.997, 3.2, 0.233, 0.233, { fill: { color: W, transparency: 70 } });
  oval(s, 8.997, 2.855, 0.233, 0.233, { rotate: 270, fill: { color: W } });
  glyph(s, 9.028, 2.887, 0.171, 0.171);
  glyph(s, 9.028, 3.231, 0.171, 0.171);
  head(s, 'Welcome Our  Festival Event Organizer', 1.206, 1.494, 4.164, 0.909, { color: W });
  btn(s, 'Introduction', 1.206, 1.233, 0.941, 0.215, { fontSize: 7.88, fill: { color: W, transparency: 87 }, shadow: null });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipis cing elit, sed do eiusmod temp or incididunt ut labore et dolore magna ipsum dolor sit amet, con sectetur adipis cing elit, sed do eiusmod temp or incididunt ut labore et dolore magna ipsum dolor sit amet, consectetur ut labore et dolore magna ipsum dolor sit amet, con sectetur adipis cing elit, sed', 1.206, 2.524, 4.062, 1.12, { color: W });
  body(s, [{ text: 'Affan', options: { bold: true } }, { text: ' Xafier' }], 1.206, 3.807, 4.062, 0.33, { fontSize: 10.13, color: W });
  body(s, 'CEO & Founder', 1.206, 4.207, 4.062, 0.287, { color: W });
  rule(s, 1.154, 4.216, 1.64, 0, { color: W });
  blob(s, SHAPE1, 6.213, 3.792, 0.535, 0.615, { rotate: 276.1, fill: { color: YEL }, shadow: GLOW() });
}

// 3. What We Do
function slide03(s) {
  blob(s, SHAPE2, 1.374, 1.18, 3.008, 3.404, { rotate: 59.8, flipH: true, fill: { color: RED } });
  blob(s, SHAPE2, 2.289, 1.919, 2.718, 3.077, { rotate: 300.8, flipH: true, fill: { color: W, transparency: 91 } });
  chrome(s, { brand: [RED, GREY3], link: GREY5, foot: GREY5 });
  oval(s, 1.392, 1.1, 2.025, 2.025, { fill: { color: W } });
  oval(s, 2.808, 2.984, 1.785, 1.785, { fill: { color: W } });
  head(s, 'What We Do ?', 4.882, 1.331, 4.164, 0.505);
  oval(s, 3.945, 2.784, 0.102, 0.102, { rotate: 313.6, fill: { color: YEL } });
  oval(s, 4.126, 2.679, 0.102, 0.102, { rotate: 313.6, fill: { color: YEL } });
  oval(s, 3.897, 2.543, 0.188, 0.188, { rotate: 313.6, fill: { color: RED } });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna ac ipsum dolor ipsum dolor sit amet, consectetur adi euismod leo. Donec com modo et urna consectetur adi pis cing elit. Vivamus vel euismod leo. ', 4.935, 1.965, 3.864, 0.912);
  body(s, '90.12 %', 4.911, 2.993, 1.176, 0.507, { fontSize: 18, color: DARK, bold: true });
  s.addShape(SH.triangle, { x: 5.958, y: 3.117, w: 0.076, h: 0.065, fill: { color: RED } });
  body(s, 'Culture Festival Event Project', 4.9, 3.437, 2.295, 0.287);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et', 4.935, 3.722, 3.106, 0.495);
  photo(s, FRAME1, 1.484, 1.193, 1.84, 1.84);
  photo(s, FRAME1, 2.91, 3.066, 1.622, 1.622);
}

// 4. What Is Festival Event
function slide04(s) {
  blob(s, SHAPE2, 4.863, 1.389, 3.008, 3.404, { rotate: 338.9, flipH: true, fill: { color: RED } });
  photo(s, FRAME2, 5.782, 1.057, 2.976, 1.761);
  photo(s, FRAME3, 1.319, 2.835, 2.063, 1.829);
  photo(s, FRAME3, 3.498, 2.835, 2.063, 1.829);
  chrome(s, { brand: [RED, GREY3], link: GREY5, foot: GREY5 });
  card(s, 5.172, 2.008, 2.19, 2.345, { rectRadius: 0.159 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna', 1.188, 2.032, 3.162, 0.495);
  body(s, 'Show Case', 5.419, 2.305, 1.738, 0.372, { fontSize: 12, color: DARK, bold: true });
  body(s, 'Culture Festival ', 5.626, 2.689, 1.682, 0.304, { fontSize: 9 });
  body(s, 'Event Festival Conceptor ', 5.626, 3.006, 1.682, 0.304, { fontSize: 9 });
  oval(s, 5.535, 2.82, 0.065, 0.065, { fill: { color: RED } });
  body(s, 'Festival Organizer', 5.626, 3.322, 1.682, 0.304, { fontSize: 9 });
  oval(s, 5.535, 3.137, 0.065, 0.065, { fill: { color: RED } });
  body(s, 'Amazing Festival Planner', 5.626, 3.647, 1.682, 0.304, { fontSize: 9 });
  oval(s, 5.535, 3.453, 0.065, 0.065, { fill: { color: RED } });
  oval(s, 5.535, 3.778, 0.065, 0.065, { fill: { color: RED } });
  head(s, [{ text: 'What Is ', options: { breakLine: true } }, { text: 'Festival Event ?' }], 1.188, 1.144, 3.319, 0.909);
}

// 5. Key Milestones
function slide05(s) {
  blob(s, SHAPE3, 5.144, -0.867, 4.373, 5.755, { rotate: 263.7, flipH: true, fill: { color: RED } });
  card(s, 4.667, 2.534, 4.492, 2.019, { rectRadius: 0.074 });
  chrome(s, { brand: [RED, GREY3], link: GREY5, foot: GREY5 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna ac ipsum dolor ipsum dolor sit amet, consectetur adi euismod leo. Donec com modo et urna consectetur adi pis cing elit. Vivamus vel euismod leo. ', 4.943, 3.355, 3.864, 0.912);
  body(s, '2020', 4.945, 2.781, 1.647, 0.507, { fontSize: 18, color: DARK, bold: true });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et', 5.657, 2.802, 3.149, 0.495);
  blob(s, SHAPE1, 8.162, 4.22, 0.512, 0.589, { rotate: 54.4, fill: { color: YEL }, shadow: GLOW() });
  head(s, 'Key Milestones', 4.935, 1.699, 4.164, 0.505, { color: W });
  photo(s, FRAME4, 0.989, 1.108, 3.463, 3.496);
}

// 6. About Our Vision Statement
function slide06(s) {
  blob(s, SHAPE4, 1.282, 0.583, 3.436, 4.751, { rotate: 70.6, flipH: true, fill: { color: RED } });
  chrome(s, { brand: [RED, GREY3], link: GREY5, foot: GREY5 });
  oval(s, 4.691, 3.56, 0.434, 0.434, { rotate: 270, fill: { color: W }, shadow: DROP() });
  glyph(s, 4.748, 3.62, 0.319, 0.319);
  head(s, 'About Our Vision Statement', 5.964, 1.294, 3.112, 0.909);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna ac ipsum dolor ipsum dolor sit amet, consectetur adi euismod leo. Donec ipsum dolor sit amet, ', 5.964, 2.272, 3.123, 0.912);
  btn(s, 'See More', 5.964, 3.936, 1.041, 0.278, { fill: { color: RED }, shadow: CARD() });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et', 5.964, 3.214, 3.123, 0.495);
  photo(s, FRAME5, 0.913, 1.869, 1.483, 1.486);
  photo(s, FRAME5, 2.532, 1.869, 1.483, 1.486);
  photo(s, FRAME5, 4.166, 1.869, 1.483, 1.486);
}

// 7. Mission Statement
function slide07(s) {
  blob(s, SHAPE2, 5.759, 0.516, 3.609, 3.717, { rotate: 302.8, flipH: true, fill: { color: RED } });
  chrome(s, { brand: [RED, GREY3], link: GREY5, foot: GREY5 });
  head(s, 'Mission Statement', 1.188, 1.349, 4.164, 0.505);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna ac ipsum dolor ipsum dolor sit amet, consectetur adi euismod leo. Donec', 1.188, 1.906, 3.675, 0.704);
  body(s, 'Amazing And Memorable', 1.188, 2.701, 3.504, 0.287, { bullet: { characterCode: '2022', indent: 10 } });
  body(s, 'Happiness On Festival Event', 1.188, 2.988, 3.504, 0.287, { bullet: { characterCode: '2022', indent: 10 } });
  body(s, 'More Than Other Festival Organizer', 1.188, 3.276, 3.504, 0.287, { bullet: { characterCode: '2022', indent: 10 } });
  btn(s, 'More Detail', 1.211, 4.021, 0.941, 0.215, { fill: { color: RED }, shadow: null });
  body(s, 'Creative And Innovative Project For Festival', 1.188, 3.563, 3.504, 0.287, { bullet: { characterCode: '2022', indent: 10 } });
  photo(s, FRAME6, 5, 2.411, 1.732, 2.063);
  photo(s, FRAME7, 6.88, 1.152, 3.12, 2.93);
}

// 8. Meet Our Professional
function slide08(s) {
  panel(s, 3.626, 1.366, 2.748, 3.716, { rectRadius: 0.1, fill: { color: RED }, shadow: GLOW2() });
  chrome(s, { brand: [RED, GREY3], link: GREY5, foot: GREY5 });
  head(s, 'Meet Our Professional', 2.237, 0.622, 5.526, 0.505, { align: 'center' });
  btn(s, 'MARTIN FRANCISO', 1.444, 3.696, 1.585, 0.302, { fontSize: 9, color: RED, bold: true, fill: { color: W }, shadow: CARD() });
  body(s, 'Lorem ipsum dolor sit amet, consec tetur adi pis cing elit. ', 1.101, 4.228, 2.271, 0.495, { align: 'center' });
  body(s, 'Position Here', 1.635, 4.007, 1.203, 0.287, { color: DARK, italic: true, align: 'center' });
  btn(s, 'MARCOS FIRGUSO', 4.208, 3.696, 1.585, 0.302, { fontSize: 9, bold: true, fill: { color: YEL }, shadow: CARD() });
  body(s, 'Lorem ipsum dolor sit amet, consec tetur adi pis cing elit. ', 3.865, 4.228, 2.271, 0.495, { color: W, align: 'center' });
  body(s, 'Position Here', 4.398, 4.007, 1.203, 0.287, { color: W, italic: true, align: 'center' });
  btn(s, 'MARTINA LETTY', 7.012, 3.696, 1.585, 0.302, { fontSize: 9, color: RED, bold: true, fill: { color: W }, shadow: CARD() });
  body(s, 'Lorem ipsum dolor sit amet, consec tetur adi pis cing elit. ', 6.669, 4.228, 2.271, 0.495, { align: 'center' });
  body(s, 'Position Here', 7.203, 4.007, 1.203, 0.287, { color: DARK, italic: true, align: 'center' });
  photo(s, FRAME1, 1.292, 1.617, 1.89, 1.89);
  photo(s, FRAME1, 4.055, 1.617, 1.89, 1.89);
  photo(s, FRAME1, 6.86, 1.617, 1.89, 1.89);
}

// 9. Organizing Your Festival Concept
function slide09(s) {
  blob(s, SHAPE5, 6.274, 0.756, 3.964, 4.092, { rotate: 278.8, flipH: true, fill: { color: RED } });
  chrome(s, { brand: [RED, GREY3], link: GREY5, foot: GREY5 });
  head(s, 'Organizing Your Festival Concept', 1.192, 1.314, 4.164, 0.909);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna ac ipsum dolor ipsum dolor sit amet, ', 1.211, 2.327, 3.258, 0.704);
  body(s, 'Lorem ipsum dolor sit a met, consectetur ad', 1.192, 3.432, 1.431, 0.495);
  body(s, 'Lorem ipsum dolor sit a met, consectetur adi', 2.829, 3.432, 1.431, 0.495);
  body(s, '02', 2.846, 3.136, 1.431, 0.372, { fontSize: 12, color: DARK, bold: true });
  body(s, '01', 1.174, 3.136, 1.431, 0.372, { fontSize: 12, color: RED, bold: true });
  btn(s, 'See More', 1.236, 4.076, 0.858, 0.236, { fill: { color: RED }, shadow: CARD() });
  btn(s, 'See More', 2.864, 4.076, 0.858, 0.236, { color: GREY, fill: { color: W }, shadow: CARD() });
  photo(s, FRAME8, 4.891, 0.794, 2.251, 2.786);
  photo(s, FRAME9, 7.281, 2.095, 1.483, 1.911);
}

// 10. About Our Service
function slide10(s) {
  photo(s, FRAME, 0, 0, 5.767, 5.625);
  fade(s, 0, 0, 5.975, 5.625, [[0, 'FF4D6A'], [68, 'FFFFFF']], 0);
  blob(s, SHAPE1, 5.736, 2.122, 3.274, 2.854, { rotate: 167.6, fill: { color: YEL, transparency: 28 }, shadow: GLOW() });
  oval(s, 6.853, 3.123, 1.33, 1.294, { rotate: 345.1, fill: { color: RED } });
  chrome(s, { brand: [RED, GREY3], link: GREY5, foot: GREY5 });
  blob(s, SHAPE1, 3.966, 0.892, 5.104, 4.062, { rotate: 184.8, fill: { color: RED, transparency: 28 } });
  head(s, 'About Our Service', 4.53, 1.681, 4.164, 0.505, { color: W });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna ac ipsum dolor ipsum dolor sit amet, consectetur adi euismod leo. Donec com modo et urna consectetur adi pis cing elit. Vivamus vel', 4.53, 2.217, 3.649, 0.912, { color: W });
  blob(s, SHAPE6, 8.582, 2.757, 1.482, 1.706, { rotate: 355.7, fill: { color: RED, transparency: 26 }, shadow: GLOW() });
  btn(s, 'More Detail >', 7.075, 3.703, 0.941, 0.215, { line: { color: W }, shadow: null });
}

// 11. Our Service Types
function slide11(s) {
  card(s, 1.2, 2.663, 7.979, 2.097, { rectRadius: 0.076 });
  blob(s, SHAPE7, 6.039, -0.457, 4.51, 4.528, { rotate: 247.7, fill: { color: RED } });
  photo(s, FRAME10, 1.2, 2.663, 7.979, 2.097);
  chrome(s, { brand: [RED, GREY3], link: GREY5, foot: GREY5 });
  head(s, 'Our Service Types', 1.211, 1.142, 4.164, 0.505);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna ac ipsum dolor ipsum dolor sit amet, sit amet, consectetur adi pis cing elit. ', 1.211, 1.704, 4.164, 0.704);
  card(s, 6.173, 1.515, 2.289, 1.88, { rectRadius: 0.068 });
  oval(s, 6.544, 2.316, 0.065, 0.065, { fill: { color: RED } });
  oval(s, 6.544, 2.633, 0.065, 0.065, { fill: { color: RED } });
  oval(s, 6.544, 2.957, 0.065, 0.065, { fill: { color: RED } });
  body(s, 'Event Festival Conceptor ', 6.635, 2.185, 1.682, 0.304, { fontSize: 9 });
  body(s, 'Festival Organizer', 6.635, 2.502, 1.682, 0.304, { fontSize: 9 });
  body(s, 'Amazing Festival Planner', 6.635, 2.826, 1.682, 0.304, { fontSize: 9 });
  body(s, 'Options Detail:', 6.446, 1.839, 2.755, 0.372, { fontSize: 12, color: DARK });
}

// 12. About Our Curants Festival Project
function slide12(s) {
  photo(s, FRAME11, 4.167, 1.065, 5.833, 4.56);
  fade(s, 4.167, 1.052, 5.833, 4.573, [[0, 'FF405F'], [60, 'FFFFFF']], 225);
  chrome(s, { brand: [RED, GREY3], link: GREY5, foot: GREY7, dot: [0.673, 4.444, 293.2] });
  card(s, 1.108, 1.397, 4.164, 1.512, { rectRadius: 0.323 });
  txt(s, '01', 0.849, 1.627, 0.583, 0.583, { fontSize: 10.5, color: W, bold: true, align: 'center', valign: 'middle', shape: SH.ellipse, fill: { color: RED } });
  panel(s, 1.692, 1.991, 2.501, 0.1, { rectRadius: 0.05, fill: { color: GREY1 } });
  panel(s, 1.692, 1.991, 1.864, 0.1, { rectRadius: 0.05, fill: { color: RED } });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna', 1.626, 2.161, 3.253, 0.495);
  body(s, 'About Aspect Ratio', 1.626, 1.641, 3.253, 0.287, { color: DARK, bold: true });
  body(s, '+11.381', 4.272, 1.879, 0.79, 0.287);
  card(s, 1.108, 3.127, 4.164, 1.512, { rectRadius: 0.323 });
  txt(s, '02', 0.849, 3.357, 0.583, 0.583, { fontSize: 10.5, color: W, bold: true, align: 'center', valign: 'middle', shape: SH.ellipse, fill: { color: RED } });
  panel(s, 1.692, 3.721, 2.501, 0.1, { rectRadius: 0.05, fill: { color: GREY1 } });
  panel(s, 1.692, 3.721, 1.864, 0.1, { rectRadius: 0.05, fill: { color: RED } });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna', 1.626, 3.891, 3.253, 0.495);
  body(s, 'About Aspect Ratio ', 1.626, 3.371, 3.253, 0.287, { color: DARK, bold: true });
  body(s, '+13.381', 4.272, 3.609, 0.79, 0.287);
  head(s, 'About Our Curants Festival Project', 5.731, 2.812, 4.164, 0.909);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna amet, consectetur adi pis cing elit. Vivamus vel euismod', 5.727, 3.799, 3.253, 0.704);
  blob(s, SHAPE8, 2.588, -1.167, 2.939, 2.333, { rotate: 38.4, fill: { color: RED } });
}

// 13. Spectacular Show
function slide13(s) {
  blob(s, SHAPE9, 1.834, 1.327, 3.803, 3.271, { rotate: 90.8, fill: { color: RED } });
  photo(s, FRAME12, 1.635, 1.224, 3.945, 3.402);
  chrome(s, { brand: [RED, GREY3], link: GREY5, foot: GREY5 });
  head(s, 'Spectacular Show ', 5.336, 1.257, 4.164, 0.505);
  panel(s, 0.757, 1.574, 2.549, 1.344, { rectRadius: 0.049, fill: { color: RED }, shadow: CARD() });
  card(s, 5.27, 2.794, 3.734, 1.389, { rectRadius: 0.051 });
  body(s, 'Lorem ipsum dolor sit amet, cons ectetur adi pis cing elit. ', 0.976, 2.068, 2.003, 0.495, { color: W });
  body(s, 'Desciption Here', 0.976, 1.694, 2.003, 0.372, { fontSize: 12, color: W, bold: true });
  btn(s, 'See More', 2.324, 2.538, 0.858, 0.236, { fill: { color: YEL }, shadow: CARD() });
  btn(s, 'See More', 7.932, 3.829, 0.858, 0.236, { fill: { color: YEL }, shadow: CARD() });
  body(s, 'Desciption Here', 5.44, 2.918, 2.003, 0.372, { fontSize: 12, color: DARK, bold: true });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod ipsum dolor sit amet, ', 5.43, 3.24, 3.325, 0.495);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna ac ipsum dolor ipsum dolor sit amet, sit amet, consectetur adi pis cing elit. ', 5.336, 1.736, 3.684, 0.704);
}

// 14. Conceptual Festival Of Tradition
function slide14(s) {
  card(s, 0.891, 1.105, 8.375, 3.363, { rectRadius: 0.2 });
  photo(s, FRAME13, 0.891, 1.105, 8.375, 3.363);
  panel(s, 0.891, 1.105, 8.375, 3.363, { rectRadius: 0.2, fill: { color: RED, transparency: 30 } });
  photo(s, FRAME1, 1.897, 1.619, 2.191, 2.191);
  chrome(s, { brand: [RED, GREY3], link: GREY5, foot: GREY5 });
  head(s, 'Conceptual Festival Of Tradition', 4.356, 1.779, 3.845, 0.909, { color: W });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna ac ipsum dolor ipsum dolor sit amet, sit amet, consectetur adi pis cing elit. ', 4.356, 2.594, 3.911, 0.704, { color: W });
  blob(s, SHAPE1, 1.362, 0.624, 0.698, 0.803, { rotate: 29.1, fill: { color: YEL }, shadow: GLOW() });
  btn(s, 'Awesome Festival ', 4.404, 3.476, 1.607, 0.35, { fill: { color: W, transparency: 73 }, shadow: CARD() });
  oval(s, 8.788, 2.381, 0.187, 0.187, { rotate: 270, fill: { color: W }, shadow: DROP() });
  glyph(s, 8.813, 2.407, 0.138, 0.138);
  oval(s, 8.788, 2.939, 0.187, 0.187, { rotate: 90, fill: { color: W }, shadow: DROP() });
  glyph(s, 8.813, 2.965, 0.138, 0.138);
  blob(s, SHAPE1, 8.027, 4.076, 0.698, 0.803, { rotate: 224.4, fill: { color: RED }, shadow: GLOW() });
}

// 15. Implementing Our Procedural To Execution
function slide15(s) {
  blob(s, SHAPE10, 8.672, 1.563, 1.642, 2.569, { rotate: 225.4, fill: { color: RED } });
  blob(s, SHAPE11, 3.813, -0.455, 2.288, 2.394, { rotate: 225.4, fill: { color: RED } });
  chrome(s, { brand: [RED, GREY3], link: GREY5, foot: GREY5 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna ac ipsum dolor ipsum dolor sit amet, sit amet, consectetur adi pis cing elit. consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna ac ipsum dolor ipsum dolor sit amet, sit amet, consectetur adi pis', 1.211, 2.962, 4.873, 0.912);
  btn(s, 'See More', 1.286, 4.083, 1.652, 0.406, { fill: { color: RED }, shadow: CARD() });
  btn(s, 'See More', 3.004, 4.083, 1.652, 0.406, { color: RED, fill: { color: W }, shadow: CARD() });
  head(s, 'Implementing Our Procedural To Execution', 1.181, 1.557, 3.518, 1.313);
  photo(s, FRAME14, 3.26, 0, 2.504, 1.748);
  photo(s, FRAME15, 5.825, 0.737, 2.617, 2.864);
  photo(s, FRAME16, 8.352, 1.883, 1.648, 2.56);
}

// 16. Our Festival Concept Plan
function slide16(s) {
  blob(s, SHAPE12, 0, 1.708, 6.083, 3.917, { fill: { color: RED } });
  card(s, 1.199, 1.963, 2.262, 2.816, { rectRadius: 0.082 });
  card(s, 3.574, 1.963, 2.262, 2.816, { rectRadius: 0.082 });
  chrome(s, { brand: [RED, GREY3], link: GREY5, foot: GREY5, splash: [-0.206, 1.133, 1.233, 1.315, 137.6], dot: [0.737, 1.611, 326.7] });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna ac ipsum dolor ipsum dolor sit amet, sit amet, consectetur adi pis', 6.386, 2.359, 2.659, 0.912);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. ', 6.386, 3.986, 2.659, 0.495);
  glyph(s, 6.388, 1.903, 0.415, 0.415);
  body(s, 'Desciption Here', 6.842, 1.937, 2.003, 0.372, { fontSize: 12, color: DARK, bold: true });
  body(s, '+1.120', 6.386, 3.493, 2.003, 0.507, { fontSize: 18, color: RED, bold: true });
  s.addShape(SH.triangle, { x: 7.319, y: 3.666, w: 0.069, h: 0.059, fill: { color: YEL } });
  head(s, 'Our Festival Concept Plan', 2.366, 0.922, 5.268, 0.505, { align: 'center' });
  body(s, 'June, 21 2021', 1.394, 3.96, 1.832, 0.304, { fontSize: 9, color: YEL, bold: true });
  body(s, 'Lorem ipsum dolor sit amet, cons ectetur adi pis cing elit. ', 1.383, 4.239, 1.944, 0.495);
  body(s, 'Sept, 21 2021', 3.861, 3.942, 1.832, 0.304, { fontSize: 9, color: YEL, bold: true });
  body(s, 'Lorem ipsum dolor sit amet, cons ectetur adi pis cing elit. ', 3.85, 4.221, 1.944, 0.495);
  rule(s, 2.424, 4.137, 0.803, 0, { color: GREY5 });
  rule(s, 4.847, 4.096, 0.803, 0, { color: GREY5 });
  photo(s, FRAME17, 1.199, 1.963, 2.262, 1.959);
  photo(s, FRAME17, 3.574, 1.963, 2.262, 1.959);
}

// 17. Break Slide
function slide17(s) {
  blob(s, SHAPE13, 5.265, 0, 4.735, 3.743, { fill: { color: YEL } });
  panel(s, 3.502, 2.088, 3.774, 0.889, { rectRadius: 0.444, rotate: 304, fill: { color: RED } });
  blob(s, SHAPE14, 3.971, 0.373, 2.784, 1.002, { rotate: 304, fill: { color: YEL } });
  panel(s, 4.805, 1.888, 3.307, 0.496, { rectRadius: 0.248, rotate: 304, fill: { color: YEL } });
  panel(s, 5.193, 2.24, 3.176, 0.352, { rectRadius: 0.176, rotate: 304, fill: { color: RED } });
  panel(s, 5.891, 2.268, 2.6, 0.352, { rectRadius: 0.176, rotate: 304, fill: { color: YEL } });
  panel(s, 6.11, 2.998, 2.113, 0.67, { rectRadius: 0.286, rotate: 304, fill: { color: YEL } });
  panel(s, 7.215, 3.535, 0.788, 0.67, { rectRadius: 0.286, rotate: 304, fill: { color: RED } });
  panel(s, 7.967, 3.252, 0.788, 0.67, { rectRadius: 0.286, rotate: 304, fill: { color: YEL } });
  blob(s, SHAPE15, 8.182, 3.106, 2.18, 0.67, { rotate: 304, fill: { color: RED } });
  chrome(s, { brand: W, link: W, foot: W });
  body(s, 'MINUTES', 5.428, 2.68, 3.908, 1.11, { fontSize: 45, fontFace: HEAD, color: W, bold: true });
  head(s, '0', 5.754, 0.859, 1.343, 2.613, { fontSize: 149.25, color: W, align: 'center' });
  head(s, '2', 4.782, 0.535, 1.343, 2.613, { fontSize: 149.25, color: W, align: 'center' });
  oval(s, 6.798, 2.378, 0.31, 0.31, { fill: { color: RED }, shadow: DROP3() });
  head(s, 'Break Slide ', 1.044, 2.001, 4.71, 0.783, { fontSize: 40.5, color: W });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna ac ipsum dolor ipsum dolor sit amet, sit amet, consectetur', 1.044, 2.768, 3.433, 0.704, { color: W });
  btn(s, 'Start Again', 1.111, 3.657, 1.514, 0.363, { fontSize: 10.5, color: RED, fill: { color: W }, shadow: CARD() });
  oval(s, 5.631, 4.017, 0.123, 0.123, { fill: { color: W } });
  oval(s, 7.529, 3.765, 0.325, 0.325, { fill: { color: W } });
  oval(s, 5.794, 4.109, 0.123, 0.123, { fill: { color: W } });
  oval(s, 8.293, 3.28, 0.166, 0.166, { fill: { color: RED }, shadow: DROP3() });
}

// 18. Our Production
function slide18(s) {
  photo(s, FRAME18, 0.976, 1.054, 3.86, 2.242);
  panel(s, 4.279, 2.438, 2.04, 2.133, { rectRadius: 0.074, fill: { color: W }, shadow: GLOW2() });
  photo(s, FRAME19, 4.279, 2.424, 2.04, 2.133);
  chrome(s, { brand: [RED, GREY3], link: GREY5, foot: GREY5 });
  panel(s, 2.013, 2.47, 2.152, 2.42, { rectRadius: 0.078, fill: { color: RED }, shadow: CARD() });
  body(s, 'Lorem ipsum dolor sit amet, conse ctetur adi pis cing elit. Vivamus vel euismod', 2.28, 3.892, 1.819, 0.704, { color: W });
  glyph(s, 2.357, 3.113, 0.444, 0.444);
  body(s, 'Description Here', 2.28, 3.615, 2.152, 0.304, { fontSize: 9, color: W, bold: true });
  btn(s, 'See More', 6.584, 4.082, 0.902, 0.263, { fill: { color: RED }, shadow: CARD() });
  body(s, 'Desciption Here', 6.535, 2.6, 2.003, 0.372, { fontSize: 12, color: DARK, bold: true });
  body(s, 'Lorem ipsum dolor sit amet, conse ctetur adi pis cing elit. Vivamus vel euis mod ipsum dolor sit amet, ipsum dolor sit amet, conse ctetur adi pis cing ', 6.525, 3.005, 2.239, 1.12);
  body(s, '+5731', 7.602, 4.024, 0.838, 0.372, { fontSize: 12, color: YEL, bold: true });
  head(s, 'Our Production', 5.173, 1.068, 4.164, 0.505);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna ac ipsum dolor', 5.173, 1.649, 3.729, 0.495);
}

// 19. Approved Projects
function slide19(s) {
  blob(s, SHAPE9, 5.47, 0.925, 3.29, 3.521, { rotate: 294.4, fill: { color: RED } });
  panel(s, 5.489, 1.079, 1.499, 1.568, { rectRadius: 0.055, fill: { color: W }, shadow: GLOW2() });
  panel(s, 7.19, 1.612, 2.197, 1.972, { rectRadius: 0.072, fill: { color: W }, shadow: GLOW2() });
  panel(s, 5.435, 2.83, 1.563, 1.553, { rectRadius: 0.057, fill: { color: W }, shadow: GLOW2() });
  chrome(s, { brand: [RED, GREY3], link: GREY5, foot: GREY5 });
  body(s, 'Lorem ipsum dolor sit amet, conse ctetur adi pis cing elit. Vivamus vel euis mod ipsum dolor sit amet, ipsum dolor sit amet, conse ctetur adi pis cing dolor sit amet, conse amet, conse ctetur adi pis cing elit. Vivamus vel euis mod ipsum', 1.17, 1.978, 3.651, 0.912);
  body(s, [{ text: 'June, 12 2021- ', options: { color: RED } }, { text: 'Festival Description ', options: { color: DARK, breakLine: true } }], 1.493, 3.146, 3.065, 0.544, { fontSize: 10.5, bold: true, color: GREY });
  blob(s, SHAPE16, 1.285, 3.319, 0.057, 0.059, { fill: { color: RED } });
  s.addShape(SH.donut, { x: 1.251, y: 3.286, w: 0.125, h: 0.125, fill: { color: RED } });
  body(s, [{ text: 'August, 12 2021- ', options: { color: RED } }, { text: 'Festival Description ', options: { color: DARK } }], 1.493, 3.505, 3.219, 0.338, { fontSize: 10.5, bold: true, color: GREY });
  blob(s, SHAPE16, 1.285, 3.665, 0.057, 0.059, { fill: { color: RED } });
  s.addShape(SH.donut, { x: 1.251, y: 3.632, w: 0.125, h: 0.125, fill: { color: RED } });
  body(s, [{ text: 'Nov,12 2021- ', options: { color: RED } }, { text: 'Festival Description ', options: { color: DARK } }], 1.493, 3.903, 3.219, 0.338, { fontSize: 10.5, bold: true, color: GREY });
  blob(s, SHAPE16, 1.285, 4.064, 0.057, 0.059, { fill: { color: RED } });
  s.addShape(SH.donut, { x: 1.251, y: 4.031, w: 0.125, h: 0.125, fill: { color: RED } });
  head(s, 'Approved Projects', 1.17, 1.41, 4.164, 0.505);
  photo(s, FRAME19, 5.489, 1.079, 1.499, 1.568);
  photo(s, FRAME20, 7.19, 1.623, 2.197, 1.972);
  photo(s, FRAME5, 5.435, 2.83, 1.563, 1.553);
}

// 20. Provide The Best Concept
function slide20(s) {
  blob(s, SHAPE1, 0.927, 0.519, 4.365, 3.804, { rotate: 19.6, fill: { color: RED } });
  panel(s, 1.153, 1.486, 2.202, 2.779, { rectRadius: 0.08, fill: { color: W }, shadow: GLOW2() });
  panel(s, 0.954, 2.855, 2.202, 1.16, { rectRadius: 0.042, fill: { color: RED }, shadow: GLOW2() });
  chrome(s, { brand: [RED, GREY3], link: GREY5, foot: GREY5 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing', 1.411, 2.134, 1.685, 0.704);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing', 1.211, 3.332, 1.685, 0.704, { color: W });
  body(s, '01. Description', 1.41, 1.829, 1.685, 0.338, { fontSize: 10.5, color: RED, bold: true });
  body(s, '02. Decription ', 1.211, 3.032, 1.685, 0.338, { fontSize: 10.5, color: W, bold: true });
  head(s, 'Provide The Best Concept', 5.936, 1.589, 2.785, 0.909);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna ac ipsum dolor ipsum dolor sit amet, sit amet, consectetur adi pis sit amet, consectetur adi pis cing elit. ', 5.936, 2.546, 2.827, 1.12);
  oval(s, 7.691, 3.946, 0.209, 0.209, { fill: { color: RED } });
  body(s, 'Read More', 7.916, 3.877, 2.27, 0.304, { fontSize: 9, color: RED, bold: true });
  rule(s, 7.737, 4.051, 0.118, 0, { color: W });
  photo(s, FRAME8, 3.487, 1.486, 2.202, 2.779);
}

// 21. Fixed Festival Schedule
function slide21(s) {
  panel(s, 2.499, 2.646, 6.482, 1.153, { rectRadius: 0.039, fill: { color: W }, shadow: SHDW6() });
  chrome(s, { brand: [RED, GREY3], link: GREY5, foot: GREY5, splash: false });
  head(s, 'Fixed Festival  Schedule ', 2.237, 0.67, 5.526, 0.505, { align: 'center' });
  rule(s, 2.365, 2.023, 0.022, 3.602, { color: GREY3, width: 1 });
  panel(s, 2.328, 1.975, 0.095, 0.095, { rectRadius: 0.047, rotate: 45, fill: { color: RED }, line: { color: W, width: 2.25 } });
  panel(s, 2.328, 3.171, 0.095, 0.095, { rectRadius: 0.047, rotate: 45, fill: { color: YEL }, line: { color: W, width: 2.25 } });
  panel(s, 2.328, 4.424, 0.095, 0.095, { rectRadius: 0.047, rotate: 45, fill: { color: RED }, line: { color: W, width: 2.25 } });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com', 4.48, 1.855, 2.865, 0.495);
  body(s, 'First Title Festival', 4.48, 1.526, 2.003, 0.338, { fontSize: 10.5, color: DARK, bold: true });
  btn(s, 'Detail', 7.687, 1.855, 0.902, 0.263, { fill: { color: RED }, shadow: CARD() });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com', 4.523, 3.099, 2.865, 0.495);
  body(s, 'Second Title Festival', 4.523, 2.77, 2.003, 0.338, { fontSize: 10.5, color: DARK, bold: true });
  btn(s, 'Detail', 7.73, 3.1, 0.902, 0.263, { fill: { color: YEL }, shadow: CARD() });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com', 4.523, 4.313, 2.865, 0.495);
  body(s, 'Third Title Festival', 4.523, 3.984, 2.003, 0.338, { fontSize: 10.5, color: DARK, bold: true });
  btn(s, 'Detail', 7.73, 4.313, 0.902, 0.263, { fill: { color: RED }, shadow: CARD() });
  oval(s, 1.368, 1.652, 0.74, 0.74, { fill: { color: RED } });
  oval(s, 1.368, 2.848, 0.74, 0.74, { fill: { color: YEL } });
  oval(s, 1.368, 4.044, 0.74, 0.74, { fill: { color: RED } });
  body(s, '20', 1.454, 1.738, 0.57, 0.33, { fontSize: 10.13, color: W, bold: true, align: 'center' });
  body(s, 'Sept', 1.454, 1.963, 0.57, 0.304, { fontSize: 9, color: W, bold: true, align: 'center' });
  body(s, '30', 1.454, 2.926, 0.57, 0.33, { fontSize: 10.13, color: W, bold: true, align: 'center' });
  body(s, 'Sept', 1.454, 3.151, 0.57, 0.304, { fontSize: 9, color: W, bold: true, align: 'center' });
  body(s, '17 ', 1.454, 4.107, 0.57, 0.33, { fontSize: 10.13, color: W, bold: true, align: 'center' });
  body(s, 'Nov', 1.454, 4.332, 0.57, 0.304, { fontSize: 9, color: W, bold: true, align: 'center' });
  photo(s, FRAME18, 2.596, 1.537, 1.678, 0.971);
  photo(s, FRAME18, 2.596, 2.733, 1.678, 0.971);
  photo(s, FRAME18, 2.596, 3.929, 1.678, 0.971);
}

// 22. Festival Schedule (continued)
function slide22(s) {
  chrome(s, { brand: [RED, GREY3], link: GREY5, foot: GREY5 });
  rule(s, 2.39, -0.104, 0, 2.659, { color: GREY3, width: 1 });
  panel(s, 2.328, 1.409, 0.095, 0.095, { rectRadius: 0.047, rotate: 45, fill: { color: RED }, line: { color: W, width: 2.25 } });
  panel(s, 2.343, 2.597, 0.095, 0.095, { rectRadius: 0.047, rotate: 45, fill: { color: YEL }, line: { color: W, width: 2.25 } });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com', 4.48, 1.289, 2.865, 0.495);
  body(s, 'About Our Festival', 4.48, 0.96, 2.003, 0.338, { fontSize: 10.5, color: DARK, bold: true });
  btn(s, 'Detail', 7.687, 1.289, 0.902, 0.263, { fill: { color: RED }, shadow: CARD() });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com', 4.523, 2.533, 2.865, 0.495);
  body(s, 'Last Projects', 4.523, 2.204, 2.003, 0.338, { fontSize: 10.5, color: DARK, bold: true });
  btn(s, 'Detail', 7.73, 2.534, 0.902, 0.263, { fill: { color: YEL }, shadow: CARD() });
  body(s, 'Lorem ipsum dolor sit amet, conse ctetur adi pis cing elit. Vivamus vel euis mod ipsum dolor sit amet, ipsum dolor sit amet, conse ctetur adi pis cing dolor sit amet, conse amet, conse ctetur adi pis cing elit. Vivamus vel euis mod ipsum ipsum dolor sit amet, ', 1.279, 4.221, 7.442, 0.495, { align: 'center' });
  rule(s, 0, 3.521, 10, 0, { color: GREY5, width: 1.5, dashType: 'sysDash' });
  head(s, 'More Creative For Amazing Festival Concept ', 2.465, 3.783, 5.265, 0.37, { fontSize: 12, bold: true, align: 'center', lineSpacingMultiple: 1.5 });
  oval(s, 1.368, 1.087, 0.74, 0.74, { fill: { color: RED } });
  oval(s, 1.368, 2.282, 0.74, 0.74, { fill: { color: YEL } });
  body(s, '25', 1.454, 1.177, 0.57, 0.33, { fontSize: 10.13, color: W, bold: true, align: 'center' });
  body(s, 'Nov', 1.454, 1.402, 0.57, 0.304, { fontSize: 9, color: W, bold: true, align: 'center' });
  body(s, '30 ', 1.454, 2.362, 0.57, 0.33, { fontSize: 10.13, color: W, bold: true, align: 'center' });
  body(s, 'Dec', 1.454, 2.587, 0.57, 0.304, { fontSize: 9, color: W, bold: true, align: 'center' });
  photo(s, FRAME18, 2.596, 0.971, 1.678, 0.971);
  photo(s, FRAME18, 2.596, 2.167, 1.678, 0.971);
}

// 23. Last Festival Project
function slide23(s) {
  blob(s, SHAPE17, 6.008, 1.684, 3.414, 4.826, { rotate: 265.4, fill: { color: RED }, shadow: GLOW() });
  chrome(s, { brand: [RED, GREY3], link: GREY5, foot: GREY5 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna ac ipsum dolor ipsum dolor sit amet, sit amet, consectetur adi', 1.211, 1.851, 3.617, 0.704);
  txt(s, '01', 1.104, 2.699, 0.567, 0.505, { color: W, bold: true, align: 'center', valign: 'middle', shape: SH.ellipse, fill: { color: RED } });
  body(s, 'Simple Description', 1.722, 2.688, 2.27, 0.338, { fontSize: 10.5, color: DARK, bold: true });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com', 1.722, 2.967, 3.105, 0.495);
  txt(s, '02', 1.104, 3.607, 0.567, 0.505, { color: W, bold: true, align: 'center', valign: 'middle', shape: SH.ellipse, fill: { color: RED } });
  body(s, 'Simple Description', 1.722, 3.596, 2.27, 0.338, { fontSize: 10.5, color: DARK, bold: true });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com', 1.722, 3.875, 3.105, 0.495);
  head(s, 'Last Festival Project', 1.211, 1.213, 4.164, 0.505);
  photo(s, FRAME1, 5.446, 0.939, 3.321, 3.321);
}

// 24. Chart Project Analysis
function slide24(s) {
  photo(s, FRAME21, 0, 0, 4.592, 5.586);
  fade(s, 0, 0, 4.592, 5.585, [[0, 'FF4D6A'], [68, 'FFFFFF']], 348);
  blob(s, SHAPE18, 6.633, 2.68, 3.178, 3.957, { rotate: 106.4, flipH: true, fill: { color: RED } });
  chrome(s, { brand: [RED, GREY3, GREY7], link: GREY5, foot: W });
  head(s, 'Our Chart Project‘s Analysis System Presentation', 5.516, 1.352, 3.756, 1.313);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna ac ipsum dolor ipsum dolor sit amet, sit amet, consectetur adi consectetur adi pis cing elit. ', 5.516, 2.812, 3.232, 0.912);
  panel(s, 1.578, 2.129, 3.537, 2.295, { rectRadius: 0.21, fill: { color: W }, shadow: GLOW2() });
  btn(s, 'Detail', 5.516, 3.931, 0.902, 0.263, { fill: { color: RED }, shadow: CARD() });
  txt(s, '6', 1.796, 2.381, 0.201, 0.252, { fontSize: 9 });
  txt(s, '5', 1.796, 2.582, 0.201, 0.252, { fontSize: 9 });
  txt(s, '4', 1.796, 2.782, 0.201, 0.252, { fontSize: 9 });
  txt(s, '3', 1.796, 2.983, 0.201, 0.252, { fontSize: 9 });
  txt(s, '2', 1.796, 3.183, 0.201, 0.252, { fontSize: 9 });
  txt(s, '1', 1.796, 3.384, 0.201, 0.252, { fontSize: 9 });
  txt(s, '0', 1.796, 3.572, 0.201, 0.252, { fontSize: 9 });
  rule(s, 2.067, 2.495, 2.689, 0, { color: GREY5 });
  rule(s, 2.067, 2.695, 2.689, 0, { color: GREY5 });
  rule(s, 2.067, 3.301, 2.689, 0, { color: GREY5 });
  rule(s, 2.067, 3.492, 2.689, 0, { color: GREY5 });
  rule(s, 2.067, 3.697, 2.689, 0, { color: GREY5 });
  txt(s, 'Category 1', 2.037, 3.752, 0.775, 0.24);
  box(s, 2.398, 4.096, 0.075, 0.076, { fill: { color: RED } });
  txt(s, 'Series 1', 2.414, 4.02, 0.645, 0.404, { fontSize: 9, color: DARK });
  box(s, 2.35, 3.221, 0.12, 0.481, { fill: { color: YEL } });
  box(s, 2.499, 3.307, 0.12, 0.395, { fill: { color: REDD } });
  box(s, 3.173, 3.221, 0.12, 0.481, { fill: { color: REDD } });
  box(s, 2.88, 3.168, 0.12, 0.534, { fill: { color: RED } });
  box(s, 3.852, 3.092, 0.12, 0.61, { fill: { color: REDD } });
  box(s, 3.703, 3.168, 0.12, 0.534, { fill: { color: YEL } });
  box(s, 4.378, 3.139, 0.12, 0.563, { fill: { color: YEL } });
  rule(s, 2.067, 2.898, 2.689, 0, { color: GREY5 });
  rule(s, 2.067, 3.098, 2.689, 0, { color: GREY5 });
  box(s, 3.55, 3.007, 0.12, 0.695, { fill: { color: RED } });
  box(s, 4.221, 2.795, 0.12, 0.907, { fill: { color: RED } });
  box(s, 4.525, 2.728, 0.12, 0.974, { fill: { color: REDD } });
  box(s, 3.029, 2.825, 0.12, 0.877, { fill: { color: YEL } });
  box(s, 2.206, 2.862, 0.12, 0.84, { fill: { color: RED } });
  txt(s, 'Category 2', 2.721, 3.752, 0.775, 0.24);
  txt(s, 'Category 3', 3.405, 3.752, 0.775, 0.24);
  txt(s, 'Category 4', 4.089, 3.752, 0.775, 0.24);
  box(s, 3.145, 4.096, 0.075, 0.076, { fill: { color: YEL } });
  txt(s, 'Series 2', 3.162, 4.02, 0.645, 0.404, { fontSize: 9, color: DARK });
  box(s, 3.894, 4.096, 0.075, 0.076, { fill: { color: REDD } });
  txt(s, 'Series 3', 3.911, 4.02, 0.645, 0.404, { fontSize: 9, color: DARK });
}

// 25. SWOT - Strength
function slide25(s) {
  blob(s, SHAPE19, 7.584, -0.281, 2.312, 2.682, { rotate: 265.4, fill: { color: RED }, shadow: GLOW() });
  chrome(s, { brand: [RED, GREY3], link: W, foot: GREY5 });
  head(s, 'SWOT Analysis System', 2.077, 0.914, 5.846, 0.505, { align: 'center' });
  oval(s, 3.548, 2.005, 1.35, 1.35, { fill: { color: W }, shadow: HALO() });
  oval(s, 5.105, 2.005, 1.35, 1.35, { fill: { color: W }, shadow: HALO() });
  oval(s, 6.694, 2.005, 1.35, 1.35, { fill: { color: W }, shadow: HALO() });
  oval(s, 1.683, 1.714, 1.853, 1.853, { fill: { color: RED }, shadow: WIDE() });
  txt(s, 'S', 1.739, 1.555, 1.736, 2.196, { fontSize: 124.5, color: W, bold: true, align: 'center' });
  txt(s, 'W', 3.341, 2.137, 1.736, 1.111, { fontSize: 60, color: RED, bold: true, align: 'center' });
  txt(s, 'O', 4.912, 2.137, 1.736, 1.111, { fontSize: 60, color: RED, bold: true, align: 'center' });
  txt(s, 'T', 6.501, 2.137, 1.736, 1.111, { fontSize: 60, color: RED, bold: true, align: 'center' });
  txt(s, 'STRENGHT', 1.684, 2.544, 1.848, 0.198, { fontSize: 12, color: W, align: 'center', valign: 'middle', shape: SH.rect, rotate: 30, fill: { color: YEL } });
  body(s, 'Lorem ipsum dolor sit amet, conse ctetur adi pis cing elit. Vivamus vel euis mod ipsum dolor sit amet, ipsum dolor sit amet, conse ctetur adi pis cing dolor sit amet, conse amet, conse ctetur adi pis cing elit. Vivamus vel euis mod ipsum ipsum dolor sit amet, ', 1.279, 4.316, 7.442, 0.495, { align: 'center' });
  body(s, '+1.120', 3.911, 3.658, 0.885, 0.439, { fontSize: 15, color: DARK, bold: true, align: 'center' });
  s.addShape(SH.triangle, { x: 4.727, y: 3.814, w: 0.069, h: 0.059, fill: { color: YEL } });
  body(s, '90.23%', 5.175, 3.67, 1.049, 0.439, { fontSize: 15, color: DARK, bold: true, align: 'center' });
  s.addShape(SH.triangle, { x: 6.139, y: 3.813, w: 0.069, h: 0.059, fill: { color: YEL } });
  oval(s, 8.631, 1.385, 0.238, 0.238, { rotate: 293.2, fill: { color: RED } });
  body(s, 'Projects Apporval', 3.547, 3.967, 1.627, 0.287, { align: 'center' });
  body(s, 'Our Projects Ratio', 4.885, 3.987, 1.627, 0.287, { align: 'center' });
}

// 26. SWOT - Weakness
function slide26(s) {
  blob(s, SHAPE19, 7.584, -0.281, 2.312, 2.682, { rotate: 265.4, fill: { color: RED }, shadow: GLOW() });
  chrome(s, { brand: [RED, GREY3], link: W, foot: GREY5 });
  head(s, 'SWOT Analysis System', 2.077, 0.914, 5.846, 0.505, { align: 'center' });
  body(s, 'Lorem ipsum dolor sit amet, conse ctetur adi pis cing elit. Vivamus vel euis mod ipsum dolor sit amet, ipsum dolor sit amet, conse ctetur adi pis cing dolor sit amet, conse amet, ', 1.911, 4.144, 3.534, 0.704);
  body(s, '1K', 5.928, 4.069, 0.816, 0.439, { fontSize: 15, color: DARK, bold: true, align: 'center' });
  s.addShape(SH.triangle, { x: 6.628, y: 4.225, w: 0.069, h: 0.059, fill: { color: YEL } });
  body(s, '0.23%', 6.994, 4.08, 1.049, 0.439, { fontSize: 15, color: DARK, bold: true, align: 'center' });
  s.addShape(SH.triangle, { x: 7.882, y: 4.223, w: 0.069, h: 0.059, rotate: 180, fill: { color: RED } });
  oval(s, 8.631, 1.385, 0.238, 0.238, { rotate: 293.2, fill: { color: RED } });
  body(s, 'First Aspect ', 5.564, 4.377, 1.627, 0.287, { align: 'center' });
  body(s, 'Weakness Ratio', 6.705, 4.398, 1.627, 0.287, { align: 'center' });
  oval(s, 1.966, 2.005, 1.35, 1.35, { fill: { color: W }, shadow: HALO() });
  txt(s, 'S', 1.76, 2.137, 1.736, 1.111, { fontSize: 60, color: RED, bold: true, align: 'center' });
  oval(s, 5.108, 2.005, 1.35, 1.35, { fill: { color: W }, shadow: HALO() });
  oval(s, 6.697, 2.005, 1.35, 1.35, { fill: { color: W }, shadow: HALO() });
  txt(s, 'O', 4.915, 2.137, 1.736, 1.111, { fontSize: 60, color: RED, bold: true, align: 'center' });
  txt(s, 'T', 6.504, 2.137, 1.736, 1.111, { fontSize: 60, color: RED, bold: true, align: 'center' });
  oval(s, 3.32, 1.714, 1.853, 1.853, { fill: { color: RED }, shadow: WIDE() });
  txt(s, 'W', 3.344, 1.893, 1.736, 1.553, { fontSize: 86.25, color: W, bold: true, align: 'center' });
  txt(s, 'WEAKNESS', 3.321, 2.544, 1.848, 0.198, { fontSize: 12, color: W, align: 'center', valign: 'middle', shape: SH.rect, rotate: 30, fill: { color: YEL } });
  body(s, 'Weakness Solving ', 1.919, 3.886, 3.534, 0.287, { color: DARK, bold: true });
}

// 27. SWOT - Opportunity
function slide27(s) {
  blob(s, SHAPE19, 7.584, -0.281, 2.312, 2.682, { rotate: 265.4, fill: { color: RED }, shadow: GLOW() });
  chrome(s, { brand: [RED, GREY3], link: W, foot: GREY5 });
  head(s, 'SWOT Analysis System', 2.077, 0.914, 5.846, 0.505, { align: 'center' });
  body(s, 'Lorem ipsum dolor sit amet, conse ctetur adi pis cing elit. Vivamus vel euis mod ipsum', 1.958, 4.172, 1.853, 0.704);
  oval(s, 8.631, 1.385, 0.238, 0.238, { rotate: 293.2, fill: { color: RED } });
  oval(s, 1.966, 2.005, 1.35, 1.35, { fill: { color: W }, shadow: HALO() });
  txt(s, 'S', 1.76, 2.137, 1.736, 1.111, { fontSize: 60, color: RED, bold: true, align: 'center' });
  oval(s, 6.697, 2.005, 1.35, 1.35, { fill: { color: W }, shadow: HALO() });
  oval(s, 5.108, 2.005, 1.35, 1.35, { fill: { color: W }, shadow: HALO() });
  txt(s, 'O', 4.915, 2.137, 1.736, 1.111, { fontSize: 60, color: RED, bold: true, align: 'center' });
  txt(s, 'T', 6.504, 2.137, 1.736, 1.111, { fontSize: 60, color: RED, bold: true, align: 'center' });
  oval(s, 3.536, 2.005, 1.35, 1.35, { fill: { color: W }, shadow: HALO() });
  txt(s, 'W', 3.343, 2.137, 1.736, 1.111, { fontSize: 60, color: RED, bold: true, align: 'center' });
  oval(s, 4.874, 1.714, 1.853, 1.853, { fill: { color: RED }, shadow: WIDE() });
  txt(s, 'O', 4.898, 1.893, 1.736, 1.553, { fontSize: 86.25, color: W, bold: true, align: 'center' });
  txt(s, 'OPPORTUNITY', 4.875, 2.544, 1.848, 0.198, { fontSize: 12, color: W, align: 'center', valign: 'middle', shape: SH.rect, rotate: 30, fill: { color: YEL } });
  body(s, 'Lorem ipsum dolor sit amet, conse ctetur adi pis cing elit. Vivamus vel euis mod ipsum', 4.103, 4.134, 1.853, 0.704);
  body(s, 'Lorem ipsum dolor sit amet, conse ctetur adi pis cing elit. Vivamus vel euis mod ipsum', 6.248, 4.101, 1.853, 0.704);
  body(s, '01. About Relations', 1.958, 3.883, 1.853, 0.287, { color: DARK, bold: true });
  body(s, '02. Marketplace', 4.074, 3.883, 1.853, 0.287, { color: DARK, bold: true });
  body(s, '03. More Opportunity', 6.219, 3.883, 1.853, 0.287, { color: DARK, bold: true });
}

// 28. SWOT - Threats
function slide28(s) {
  blob(s, SHAPE19, 7.584, -0.281, 2.312, 2.682, { rotate: 265.4, fill: { color: RED }, shadow: GLOW() });
  chrome(s, { brand: [RED, GREY3], link: W, foot: GREY5 });
  head(s, 'SWOT Analysis System', 2.077, 0.914, 5.846, 0.505, { align: 'center' });
  body(s, 'Lorem ipsum dolor sit amet, conse ctetur adi pis cing elit. Vivamus vel euis mod ipsum dolor sit amet, ipsum dolor sit am et, conse ctetur adi pis cing dolor sit amet, ', 4.625, 3.953, 3.542, 0.704);
  body(s, '10 %', 2.029, 3.997, 0.816, 0.33, { fontSize: 10.13, color: DARK, bold: true, align: 'center' });
  s.addShape(SH.triangle, { x: 2.749, y: 4.114, w: 0.069, h: 0.059, fill: { color: YEL } });
  body(s, '12.23%', 3.21, 3.997, 1.049, 0.33, { fontSize: 10.13, color: DARK, bold: true, align: 'center' });
  s.addShape(SH.triangle, { x: 4.097, y: 4.114, w: 0.069, h: 0.059, rotate: 180, fill: { color: RED } });
  oval(s, 8.631, 1.385, 0.238, 0.238, { rotate: 293.2, fill: { color: RED } });
  body(s, 'Projects Treats ', 1.665, 4.278, 1.627, 0.287, { align: 'center' });
  body(s, 'Threats Ratio', 2.921, 4.278, 1.627, 0.287, { align: 'center' });
  oval(s, 1.962, 2.005, 1.35, 1.35, { fill: { color: W }, shadow: HALO() });
  txt(s, 'S', 1.756, 2.137, 1.736, 1.111, { fontSize: 60, color: RED, bold: true, align: 'center' });
  oval(s, 5.103, 2.005, 1.35, 1.35, { fill: { color: W }, shadow: HALO() });
  txt(s, 'O', 4.911, 2.137, 1.736, 1.111, { fontSize: 60, color: RED, bold: true, align: 'center' });
  oval(s, 3.532, 2.005, 1.35, 1.35, { fill: { color: W }, shadow: HALO() });
  txt(s, 'W', 3.339, 2.137, 1.736, 1.111, { fontSize: 60, color: RED, bold: true, align: 'center' });
  oval(s, 6.391, 1.714, 1.853, 1.853, { fill: { color: RED }, shadow: WIDE() });
  txt(s, 'T', 6.416, 1.893, 1.736, 1.553, { fontSize: 86.25, color: W, bold: true, align: 'center' });
  txt(s, 'THREATS', 6.393, 2.544, 1.848, 0.198, { fontSize: 12, color: W, align: 'center', valign: 'middle', shape: SH.rect, rotate: 30, fill: { color: YEL } });
}

// 29. Get In Touch
function slide29(s) {
  photo(s, FRAME, 0, 0, 10, 5.625);
  fade(s, 0, 0, 10, 5.625, [[0, 'FF4D6A'], [79, 'FFFFFF']], 135);
  blob(s, SHAPE20, 5.052, -0.516, 5.643, 5.95, { rotate: 244.7, fill: { color: RED } });
  chrome(s, { brand: [RED, GREY3, W], link: W, foot: '404040', splash: [-0.092, 4.561, 1.167, 1.19, 104.1], dot: [0.651, 4.741, 293.2] });
  card(s, 1.055, 3.651, 2.498, 0.794, { rectRadius: 0.169 });
  head(s, 'Get In Touch ', 5.224, 1.519, 3.337, 0.505, { color: W });
  body(s, 'Lorem ipsum dolor sit amet, conse ctetur adi pis cing elit. Vivamus vel euis mod ipsum dolor sit amet, ipsum dolor sit am et, conse ctetur adi pis cing dolor sit amet, ', 5.224, 2.05, 3.542, 0.704, { color: W });
  card(s, 3.712, 3.651, 1.902, 0.794, { rectRadius: 0.169 });
  card(s, 5.762, 3.651, 2.176, 0.794, { rectRadius: 0.169 });
  oval(s, 1.192, 3.828, 0.442, 0.442, { fill: { color: RED } });
  blob(s, SHAPE21, 1.351, 3.951, 0.123, 0.196, { fill: { color: W } });
  oval(s, 3.883, 3.828, 0.442, 0.442, { fill: { color: RED } });
  blob(s, SHAPE22, 4.014, 3.96, 0.179, 0.178, { fill: { color: W } });
  oval(s, 5.933, 3.828, 0.442, 0.442, { fill: { color: RED } });
  blob(s, SHAPE23, 6.107, 3.949, 0.094, 0.199, { fill: { color: W } });
  body(s, 'Our Sites', 6.406, 3.782, 1.722, 0.338, { fontSize: 10.5, color: DARK, bold: true, align: null });
  txt(s, 'www.jobinterview.com', 6.406, 4.047, 1.766, 0.24);
  body(s, 'Contact 	', 4.426, 3.781, 1.722, 0.338, { fontSize: 10.5, color: DARK, bold: true, align: null });
  body(s, '+10 1234 333 33 ', 4.426, 4.001, 1.389, 0.287, { align: null });
  body(s, 'Office', 1.668, 3.787, 1.722, 0.338, { fontSize: 10.5, color: DARK, bold: true, align: null });
  body(s, '21, California St, 44F-23 USA ', 1.668, 4.008, 2.187, 0.287, { align: null });
}

// 30. Thank You
function slide30(s) {
  photo(s, FRAME, 0, 0, 10, 5.625);
  fade(s, 0, 0, 10, 5.625, [[0, 'E00025'], [62, '000000']], 0);
  chrome(s, { brand: W, link: GREY5, foot: GREY5, splash: [-0.092, 4.561, 1.167, 1.19, 104.1], dot: [0.651, 4.741, 293.2] });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivam us vel euismod leo. Donec com modo et urna ac ipsum dolor ipsum dolor sit amet, sit amet, consectetur adi', 4.924, 3.463, 3.617, 0.704, { color: W });
  head(s, [{ text: 'THANK YOU ', options: { color: RED } }, { text: 'FOR WATCHING US', options: { color: W } }], 4.882, 1.529, 4.164, 1.111, { fontSize: 30, color: BLACK });
  body(s, 'Festival Event - Presentation', 4.924, 3.175, 3.617, 0.304, { fontSize: 9, color: W, bold: true });
  rule(s, 5, 2.806, 5, 0.007, { color: BLACK });
}

const SLIDES = [
  slide01, slide02, slide03, slide04, slide05, slide06,
  slide07, slide08, slide09, slide10, slide11, slide12,
  slide13, slide14, slide15, slide16, slide17, slide18,
  slide19, slide20, slide21, slide22, slide23, slide24,
  slide25, slide26, slide27, slide28, slide29, slide30,
];

const BACKGROUNDS = {
  1: W,
  17: RED,
};

SLIDES.forEach(function (build, i) {
  brandAt(0.295, 0.333);
  const s = pptx.addSlide();
  const bg = BACKGROUNDS[i + 1];
  if (bg) s.background = { color: bg };
  build(s);
});

pptx.writeFile({ fileName: path.join(__dirname, OUTPUT) })
  .then(function (f) { console.log('wrote ' + f); });
