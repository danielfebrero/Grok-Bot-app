/**
 * Lycoss "Brand Guidelines" deck — rebuilt with pptxgenjs.
 *
 * 24 slides, 13.333in x 7.5in (16:9). Two slide archetypes:
 *   - divider  : blue field, pink ribbon, huge Dela Gothic One wordmark
 *   - content  : white field, blue heading, lorem body, section caption, art on the right
 *
 * Raster art in the source deck (logo lockups, the submark, the seamless
 * pattern, device mockups, photos) is redrawn here from native pptxgenjs
 * shapes, or stands in as a flat placeholder rectangle.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Design tokens
 * ------------------------------------------------------------------ */

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

const C = {
  blue: '3E09FF',
  black: '181818',
  pink: 'F442B9',
  orange: 'EB5C2A',
  green: '0A3E29',
  ribbon: 'F779CD',
  white: 'FFFFFF',
  ink: '000000',      // body copy (theme dk1)
  cream: 'FDF5F0',    // letterhead paper
  photo: 'C3C3C3',    // stand-in for a photograph
  device: 'BDBDBD',   // stand-in for a device mockup body
  screen: 'D8D8D8',
};

const F = {
  display: 'Dela Gothic One',
  body: 'Montserrat',
  semi: 'Montserrat SemiBold',
  xbold: 'Montserrat ExtraBold',
};

// Common geometry shared by every slide.
const MARGIN_X = 1.168;          // left text rail
const LOGO = { x: 1.168, y: 0.604, w: 1.446, h: 0.472 };
const CAPTION_Y = 6.727;         // "logo section" etc.
const HEAD_Y = 2.419;            // content-slide heading
const BODY = { x: MARGIN_X, y: 3.607, w: 4.88, h: 1.458 };

const LOREM =
  'Sit amet consectetur adipiscing elit duies. Cursus eget nunc scelerisque viverra mauris in ' +
  'aliquam sem fringilla. Lorem donecy massa sapien faucibus eti. Lacinia quis vel eros donec ac. ' +
  'Risus quis varius quam quisque id. Vel orci porta noona pulvinar neque laoreet suspendisse interdum. ';

// Text boxes in the source deck carry Google-Slides insets, not PowerPoint defaults.
const INSET = [0.1, 0.1, 0.05, 0.05]; // [left, right, bottom, top] in inches

/* ------------------------------------------------------------------ *
 * Vector art, normalised to the unit square
 *
 * Each entry below is a list of closed polygons in 0..1 coordinates that get
 * scaled into whatever box the caller asks for. `ribbon` is a bezier outline
 * (moveTo / cubicBezTo / close); the rest are polygon outlines plus holes.
 * ------------------------------------------------------------------ */

// The pink swoosh behind every divider slide. Rows of 2 numbers are moveTo,
// rows of 6 are a cubic bezier (two controls + endpoint), empty rows close.
const RIBBON = [
  [0.0087,0.8259],
  [0.0151,0.8217,0.0209,0.831,0.023,0.8449],
  [0.0257,0.8631,0.032,0.8822,0.0384,0.8939],
  [0.0473,0.9102,0.0566,0.9196,0.0668,0.927],
  [0.0901,0.944,0.1136,0.9464,0.1385,0.9451],
  [0.1918,0.9423,0.2441,0.9107,0.2913,0.853],
  [0.3154,0.8236,0.3378,0.7876,0.3582,0.7457],
  [0.3802,0.7003,0.4001,0.6497,0.42,0.5993],
  [0.4578,0.5038,0.4975,0.4052,0.5487,0.3496],
  [0.5722,0.3241,0.5982,0.3087,0.6241,0.3074],
  [0.6328,0.2407,0.6488,0.1788,0.6699,0.1307],
  [0.7081,0.0435,0.762,-0.0006,0.8148,0],
  [0.8692,0.0006,0.9225,0.0492,0.9656,0.1253],
  [0.9765,0.1443,0.9868,0.1649,0.9965,0.187],
  [1.0073,0.2115,0.9908,0.25,0.98,0.2256],
  [0.9436,0.1426,0.8978,0.0818,0.8475,0.0611],
  [0.8003,0.0418,0.7499,0.0637,0.7097,0.1254],
  [0.6813,0.1691,0.6593,0.2356,0.6481,0.3105],
  [0.6653,0.3156,0.6819,0.3283,0.6971,0.3477],
  [0.7204,0.3778,0.7406,0.428,0.7481,0.4888],
  [0.7556,0.5497,0.7476,0.6193,0.7246,0.6556],
  [0.7007,0.6933,0.6681,0.6734,0.6486,0.6289],
  [0.6291,0.5845,0.6203,0.5193,0.6176,0.4573],
  [0.6162,0.4255,0.6166,0.3934,0.6188,0.3618],
  [0.5946,0.3654,0.5704,0.3832,0.5489,0.4101],
  [0.5018,0.4693,0.4658,0.5639,0.4299,0.6545],
  [0.4108,0.7028,0.3915,0.7508,0.37,0.7936],
  [0.3499,0.834,0.3279,0.8691,0.3045,0.8983],
  [0.2578,0.956,0.2065,0.99,0.1538,0.9981],
  [0.1288,1.0019,0.1031,1.0005,0.0785,0.9898],
  [0.0568,0.9803,0.0348,0.9612,0.0185,0.9253],
  [0.0103,0.907,0.0043,0.8845,0.0005,0.8594],
  [-0.0016,0.8455,0.0029,0.8296,0.0087,0.8259],
  [],
  [0.6521,0.5484],
  [0.662,0.5928,0.682,0.6322,0.7038,0.6183],
  [0.7246,0.6052,0.7312,0.5483,0.7257,0.5033],
  [0.7193,0.4514,0.7008,0.4125,0.681,0.3894],
  [0.6688,0.3754,0.6557,0.3665,0.6423,0.3631],
  [0.6416,0.3726,0.641,0.3821,0.6407,0.3917],
  [0.6387,0.445,0.6414,0.5003,0.6521,0.5484],
  [],
];

// "Lycoss / creative agency" horizontal lockup (aspect 3.059:1).
const LOGO_ART = [
  [0.998,0.83,0.988,0.83,0.973,0.904,0.968,0.83,0.958,0.83,0.966,0.94,0.961,0.955,0.954,0.949,0.949,
   0.972,0.956,0.98,0.968,0.972],
  [0.944,0.83,0.93,0.835,0.92,0.861,0.918,0.898,0.924,0.932,0.938,0.941,0.948,0.932,0.953,0.917,0.946,
   0.901,0.939,0.915,0.933,0.915,0.928,0.895,0.929,0.872,0.935,0.856,0.941,0.855,0.947,0.869,0.955,
   0.856,0.951,0.84],
  [0.68,0.83,0.67,0.83,0.655,0.903,0.649,0.832,0.639,0.832,0.647,0.94,0.658,0.938],
  [0.634,0.83,0.625,0.83,0.618,0.94,0.628,0.938],
  [0.909,0.841,0.901,0.829,0.888,0.841,0.887,0.83,0.878,0.83,0.871,0.94,0.881,0.938,0.885,0.875,0.889,
   0.861,0.895,0.855,0.9,0.86,0.902,0.873,0.897,0.938,0.907,0.94,0.912,0.873],
  [0.856,0.83,0.843,0.833,0.834,0.855,0.83,0.898,0.832,0.917,0.838,0.934,0.849,0.941,0.863,0.927,0.859,
   0.907,0.847,0.917,0.84,0.901,0.841,0.894,0.867,0.892,0.866,0.853],
  [0.826,0.83,0.817,0.83,0.815,0.847,0.805,0.829,0.791,0.841,0.784,0.875,0.786,0.912,0.796,0.934,0.805,
   0.932,0.81,0.923,0.81,0.931,0.803,0.955,0.792,0.954,0.786,0.944,0.781,0.966,0.793,0.98,0.805,0.978,
   0.812,0.968,0.819,0.935],
  [0.745,0.841,0.738,0.881,0.738,0.909,0.742,0.927,0.749,0.94,0.757,0.94,0.763,0.929,0.763,0.94,0.772,
   0.94,0.78,0.832,0.77,0.83,0.768,0.844,0.763,0.832,0.755,0.829],
  [0.705,0.832,0.692,0.83,0.68,0.856,0.677,0.901,0.684,0.932,0.696,0.941,0.709,0.927,0.706,0.907,0.694,
   0.917,0.69,0.912,0.687,0.897,0.714,0.892,0.714,0.858],
  [0.55,0.843,0.543,0.881,0.546,0.921,0.555,0.94,0.563,0.94,0.569,0.929,0.57,0.94,0.578,0.94,0.586,
   0.832,0.576,0.83,0.574,0.844,0.569,0.832,0.562,0.829],
  [0.528,0.83,0.514,0.833,0.505,0.858,0.502,0.906,0.509,0.934,0.52,0.941,0.534,0.927,0.531,0.907,0.518,
   0.917,0.512,0.9,0.512,0.894,0.539,0.894,0.537,0.85],
  [0.502,0.829,0.489,0.843,0.489,0.83,0.48,0.832,0.473,0.94,0.483,0.938,0.488,0.872,0.493,0.858,0.5,
   0.856],
  [0.467,0.838,0.454,0.829,0.44,0.847,0.434,0.88,0.435,0.909,0.439,0.929,0.453,0.941,0.462,0.935,0.468,
   0.92,0.462,0.901,0.457,0.914,0.45,0.915,0.445,0.898,0.446,0.873,0.45,0.858,0.458,0.855,0.464,0.869,
   0.472,0.856],
  [0.598,0.807,0.597,0.832,0.591,0.833,0.59,0.853,0.595,0.858,0.592,0.92,0.596,0.937,0.608,0.94,0.612,
   0.935,0.61,0.914,0.604,0.917,0.602,0.909,0.605,0.86,0.614,0.856,0.616,0.833,0.607,0.832,0.608,0.807],
  [0.628,0.784,0.626,0.793,0.628,0.812,0.633,0.815,0.638,0.804,0.636,0.782],
  [0.4,0.16,0.36,0.162,0.319,0.457,0.28,0.162,0.237,0.16,0.299,0.602,0.293,0.633,0.281,0.659,0.265,
   0.657,0.251,0.634,0.236,0.721,0.259,0.756,0.285,0.761,0.305,0.742,0.32,0.704,0.332,0.648],
  [0.919,0.159,0.898,0.181,0.88,0.225,0.874,0.275,0.876,0.341,0.881,0.369,0.892,0.397,0.908,0.415,0.953,
   0.441,0.959,0.455,0.96,0.48,0.95,0.505,0.927,0.509,0.901,0.492,0.884,0.466,0.87,0.557,0.891,0.586,
   0.916,0.602,0.943,0.603,0.965,0.591,0.981,0.566,0.995,0.522,0.999,0.475,0.996,0.412,0.986,0.377,
   0.977,0.36,0.924,0.324,0.915,0.31,0.913,0.285,0.92,0.259,0.935,0.248,0.959,0.255,0.981,0.281,0.994,
   0.193,0.982,0.174,0.962,0.159],
  [0.777,0.159,0.756,0.181,0.739,0.225,0.733,0.275,0.735,0.341,0.74,0.369,0.755,0.403,0.814,0.444,0.819,
   0.465,0.818,0.483,0.808,0.505,0.786,0.509,0.764,0.497,0.742,0.468,0.729,0.557,0.75,0.586,0.774,0.602,
   0.802,0.603,0.82,0.594,0.837,0.573,0.852,0.529,0.858,0.475,0.856,0.418,0.851,0.395,0.84,0.366,0.822,
   0.346,0.781,0.323,0.773,0.309,0.771,0.285,0.777,0.261,0.794,0.248,0.818,0.255,0.84,0.281,0.853,0.191,
   0.821,0.159],
  [0.616,0.16,0.6,0.176,0.586,0.201,0.567,0.269,0.558,0.349,0.558,0.407,0.562,0.458,0.577,0.531,0.588,
   0.56,0.604,0.586,0.628,0.603,0.649,0.602,0.669,0.586,0.692,0.543,0.708,0.474,0.714,0.406,0.714,0.35,
   0.711,0.302,0.697,0.23,0.686,0.201,0.672,0.176,0.642,0.154],
  [0.465,0.157,0.452,0.168,0.434,0.194,0.42,0.231,0.413,0.261,0.406,0.302,0.403,0.349,0.403,0.409,0.406,
   0.451,0.42,0.528,0.431,0.557,0.446,0.583,0.472,0.603,0.492,0.603,0.509,0.593,0.53,0.557,0.546,0.494,
   0.514,0.441,0.504,0.48,0.491,0.498,0.475,0.5,0.462,0.486,0.453,0.463,0.447,0.429,0.444,0.394,0.447,
   0.33,0.459,0.279,0.473,0.259,0.487,0.258,0.496,0.265,0.507,0.287,0.514,0.318,0.546,0.264,0.531,0.204,
   0.515,0.173,0.49,0.154],
  [0.002,0.688,0.001,0.773,0.006,0.821,0.013,0.864,0.032,0.923,0.047,0.949,0.062,0.963,0.087,0.966,
   0.114,0.944,0.138,0.901,0.157,0.846,0.186,0.912,0.211,0.952,0.243,0.985,0.274,0.998,0.307,0.995,
   0.348,0.968,0.377,0.932,0.384,0.894,0.384,0.869,0.381,0.843,0.368,0.818,0.356,0.821,0.328,0.855,
   0.303,0.872,0.277,0.873,0.251,0.86,0.214,0.812,0.18,0.73,0.188,0.664,0.195,0.568,0.203,0.329,0.202,
   0.173,0.198,0.049,0.192,0.015,0.181,0,0.165,0.012,0.159,0.032,0.157,0.054,0.161,0.208,0.16,0.364,
   0.155,0.515,0.147,0.633,0.124,0.585,0.107,0.56,0.086,0.543,0.062,0.542,0.041,0.557,0.023,0.588,0.011,
   0.63],
];
const LOGO_HOLES = [
  [0.841,0.873,0.848,0.852,0.855,0.855,0.859,0.869,0.858,0.875],
  [0.807,0.855,0.813,0.872,0.811,0.894,0.804,0.909,0.797,0.906,0.794,0.89,0.795,0.872,0.8,0.856],
  [0.756,0.855,0.763,0.858,0.767,0.875,0.765,0.901,0.759,0.915,0.752,0.914,0.748,0.898,0.749,0.87],
  [0.688,0.872,0.695,0.852,0.7,0.852,0.705,0.866,0.705,0.875],
  [0.562,0.855,0.569,0.858,0.573,0.873,0.571,0.901,0.565,0.915,0.557,0.912,0.553,0.897,0.555,0.87],
  [0.513,0.87,0.517,0.855,0.524,0.852,0.53,0.873],
  [0.631,0.258,0.646,0.261,0.66,0.282,0.669,0.315,0.673,0.353,0.673,0.404,0.667,0.451,0.655,0.486,0.642,
   0.5,0.626,0.497,0.613,0.477,0.604,0.443,0.599,0.395,0.6,0.344,0.607,0.301,0.618,0.272],
  [0.042,0.727,0.045,0.705,0.051,0.684,0.064,0.67,0.077,0.668,0.103,0.696,0.127,0.755,0.113,0.799,0.097,
   0.827,0.072,0.838,0.059,0.824,0.051,0.804,0.044,0.769],
];

// The standalone "d"-like submark (aspect 1.183:1).
const MARK_ART = [
  [0.0027,0.7067,0,0.7676,0.0041,0.7949,0.0312,0.867,0.0637,0.9103,0.0894,0.9327,0.1192,0.9503,0.1585,
   0.9631,0.2141,0.9647,0.2846,0.9423,0.3564,0.8894,0.4038,0.8285,0.4743,0.8974,0.5501,0.9487,0.622,
   0.9792,0.6924,0.9952,0.7913,0.9968,0.8726,0.9808,0.9878,0.9311,0.9986,0.9071,0.9932,0.8686,0.9824,
   0.8558,0.9661,0.8494,0.9363,0.8542,0.8713,0.8862,0.8062,0.9038,0.7276,0.9071,0.6653,0.8958,0.6003,
   0.8702,0.5474,0.8381,0.5,0.7997,0.4458,0.7436,0.4688,0.6715,0.4892,0.5689,0.5108,0.3221,0.5108,
   0.1971,0.5,0.0433,0.4919,0.0192,0.481,0.0064,0.4648,0,0.4377,0.008,0.4201,0.0385,0.4309,0.1923,
   0.4295,0.3638,0.4133,0.5385,0.3848,0.6699,0.3157,0.6106,0.2642,0.5801,0.2127,0.5641,0.1491,0.5625,
   0.0949,0.5785,0.0556,0.6058,0.0244,0.6458],
];
const MARK_HOLES = [
  [0.0799,0.734,0.0894,0.7019,0.1138,0.6731,0.1504,0.6571,0.1924,0.6554,0.2344,0.6683,0.2656,0.6859,
   0.3035,0.7163,0.3455,0.7628,0.3035,0.8189,0.2751,0.8429,0.2195,0.8686,0.1734,0.8702,0.1314,0.8526,
   0.1003,0.8189,0.0827,0.7772],
];

// One tile of the seamless brand pattern (aspect 0.976:1).
const TILE_ART = [
  [0.7043,0.506,0.7043,0.5387,0.7317,0.5625,0.9085,0.5714,0.997,0.6101,0.9939,0.5089,0.8537,0.4762,
   0.7439,0.4762],
  [0.0213,0.3988,0,0.4673,0.003,0.5417,0.0305,0.6071,0.0884,0.6667,0.1738,0.6935,0.2226,0.6964,0.3049,
   0.6815,0.378,0.6488,0.4604,0.7827,0.6128,0.9167,0.7287,0.9732,0.811,0.9911,0.936,0.994,0.9817,0.9762,
   0.9909,0.9494,0.9726,0.9137,0.7835,0.8988,0.7287,0.881,0.6433,0.8304,0.5335,0.7292,0.4482,0.5952,
   0.5457,0.4643,0.6433,0.2738,0.7134,0.0655,0.7165,0.0327,0.7043,0.003,0.6555,0,0.6372,0.0119,0.5427,
   0.2857,0.4787,0.4077,0.4085,0.506,0.3323,0.3988,0.2957,0.369,0.2317,0.3423,0.1829,0.3333,0.1067,
   0.3393],
];
const TILE_HOLES = [
  [0.0915,0.4613,0.128,0.4226,0.2043,0.4226,0.2896,0.4851,0.3384,0.5685,0.25,0.6071,0.1585,0.6012,
   0.1311,0.5863,0.0915,0.5387],
];

/* ------------------------------------------------------------------ *
 * Drawing helpers
 * ------------------------------------------------------------------ */

const rect = (s, o) => s.addShape('rect', o);

/** Text box using the source deck's insets and "resize shape to fit" behaviour. */
function textBox(slide, text, opts) {
  slide.addText(text, Object.assign({ margin: INSET, fit: 'resize', valign: 'top' }, opts));
}

/** Draw a unit-square polygon set (outlines first, then holes in `holeColor`). */
function drawPolys(slide, polys, box, color, holes, holeColor) {
  const put = (list, fill) => {
    list.forEach((flat) => {
      const pts = [];
      for (let i = 0; i < flat.length; i += 2) {
        pts.push({ x: box.w * flat[i], y: box.h * flat[i + 1] });
      }
      slide.addShape('custGeom', {
        x: box.x, y: box.y, w: box.w, h: box.h,
        points: pts.concat([{ close: true }]),
        fill: { color: fill }, line: { width: 0 },
      });
    });
  };
  put(polys, color);
  if (holes && holes.length) put(holes, holeColor);
}

/** The full "Lycoss creative agency" lockup, sized by width. */
function logoLockup(slide, x, y, w, color, bg) {
  drawPolys(slide, LOGO_ART, { x, y, w, h: w / 3.059 }, color, LOGO_HOLES, bg);
}

/** Header logo that appears in the same spot on all 24 slides. */
function headerLogo(slide, color, bg) {
  logoLockup(slide, LOGO.x, LOGO.y, LOGO.w, color, bg);
}

/** The pink ribbon that sweeps across every divider slide. */
function ribbon(slide) {
  const box = { x: -0.43, y: 0.501, w: 14.168, h: 6.053 };
  const pts = [];
  RIBBON.forEach((row) => {
    if (row.length === 0) { pts.push({ close: true }); return; }
    if (row.length === 2) { pts.push({ x: box.w * row[0], y: box.h * row[1], moveTo: true }); return; }
    pts.push({
      x: box.w * row[4], y: box.h * row[5],
      curve: {
        type: 'cubic',
        x1: box.w * row[0], y1: box.h * row[1],
        x2: box.w * row[2], y2: box.h * row[3],
      },
    });
  });
  slide.addShape('custGeom', {
    x: box.x, y: box.y, w: box.w, h: box.h,
    points: pts, fill: { color: C.ribbon }, line: { width: 0 },
  });
}

/** Flat stand-in for a photograph from the source deck. */
function photoPlaceholder(slide, x, y, w, h, label) {
  slide.addText(label || '[image]', {
    x, y, w, h, margin: 0, align: 'center', valign: 'middle',
    fill: { color: C.photo }, line: { width: 0 },
    fontFace: F.body, fontSize: 11, color: '8A8A8A',
  });
}

/* ------------------------------------------------------------------ *
 * Slide archetypes
 * ------------------------------------------------------------------ */

/**
 * Divider slide: blue field, ribbon, section number and a giant wordmark.
 * `title` is split into `pieces` so the ribbon can weave between the letters,
 * exactly as the source deck stacks duplicate text boxes over the artwork.
 */
function dividerSlide(pptx, spec) {
  const s = pptx.addSlide();
  s.background = { color: C.blue };
  const size = spec.fontSize;
  const wordY = spec.wordY;
  const wordH = spec.wordH;

  // Layer 1: full word, painted under the ribbon.
  spec.lines.forEach((ln) => {
    textBox(s, ln.text, {
      x: ln.x, y: ln.y, w: ln.w, h: wordH,
      fontFace: F.display, fontSize: size, color: C.white,
    });
  });

  ribbon(s);
  headerLogo(s, C.white, C.blue);

  // Layer 2: the leading letters, repainted on top of the ribbon.
  (spec.overlays || []).forEach((ov) => {
    textBox(s, ov.text, {
      x: ov.x, y: ov.y, w: ov.w, h: wordH,
      fontFace: F.display, fontSize: size, color: ov.color || C.white,
    });
  });
  (spec.dots || []).forEach((d) => {
    s.addShape('ellipse', { x: d.x, y: d.y, w: d.w, h: d.h, fill: { color: C.white }, line: { width: 0 } });
  });

  if (spec.number) {
    textBox(s, spec.number, {
      x: MARGIN_X, y: wordY - 1.279, w: 2.3, h: 1.279,
      fontFace: F.display, fontSize: 70, color: spec.numberColor || C.white,
    });
  }
  return s;
}

/**
 * Content slide: white field, blue heading, lorem paragraph, section caption.
 * `art` paints whatever belongs in the right-hand half.
 */
function contentSlide(pptx, heading, caption, art, opts) {
  const s = pptx.addSlide();
  headerLogo(s, C.blue, C.white);

  if (heading) {
    textBox(s, heading, {
      x: MARGIN_X, y: HEAD_Y, w: 6, h: 0.64,
      fontFace: F.display, fontSize: 32, color: C.blue,
    });
  }
  if (!opts || !opts.noBody) {
    textBox(s, LOREM, Object.assign({}, BODY, {
      fontFace: F.body, fontSize: 11, color: C.ink, lineSpacingMultiple: 1.5,
    }));
  }
  textBox(s, caption, {
    x: MARGIN_X, y: CAPTION_Y, w: 3, h: 0.303,
    fontFace: F.display, fontSize: 12, color: C.black,
  });
  if (art) art(s);
  return s;
}

/* ------------------------------------------------------------------ *
 * Per-slide content
 * ------------------------------------------------------------------ */

// 1 — title
function slideTitle(pptx) {
  dividerSlide(pptx, {
    fontSize: 110, wordY: 4.837, wordH: 1.952,
    lines: [
      { text: 'Brand', x: MARGIN_X, y: 2.885, w: 4.688 },
      { text: 'Guidelines', x: MARGIN_X, y: 4.837, w: 9.789 },
    ],
    overlays: [
      { text: 'd', x: 5.612, y: 2.885, w: 1.334 },
      { text: 'G', x: MARGIN_X, y: 4.828, w: 1.62 },
      { text: 'i', x: 3.633, y: 4.828, w: 0.738 },
    ],
  });
}

// 2 — contents
const CONTENTS = [
  ['Logo', 0], ['Colors', 1], ['Typography', 2],
  ['Stationary', 0], ['Digital', 1], ['Photography', 2],
];
function slideContents(pptx) {
  const s = pptx.addSlide();
  headerLogo(s, C.blue, C.white);
  textBox(s, 'Contents', {
    x: MARGIN_X, y: 2.161, w: 2.995, h: 0.707,
    fontFace: F.display, fontSize: 36, color: C.blue,
  });
  const colX = [MARGIN_X, 4.569, 7.971];
  CONTENTS.forEach(([label, col], i) => {
    const rowY = i < 3 ? 3.397 : 5.09;
    textBox(s, label, {
      x: colX[col], y: rowY, w: 3, h: 0.572,
      fontFace: F.body, fontSize: 28, bold: true, color: C.black,
    });
    textBox(s, 'Sit amet consectetur adipiscing duies sceleris', {
      x: colX[col], y: rowY + 0.572, w: 2.413, h: 0.625,
      fontFace: F.body, fontSize: 11, color: C.ink, lineSpacingMultiple: 1.5,
    });
  });
}

// 3, 8, 11, 15, 18, 21 — section dividers
const SECTIONS = [
  { n: '01.', word: 'Logo', wordW: 4.863, head: 'L', headW: 1.443,
    tail: { text: 'g', x: 3.534, w: 1.35 }, size: 110, y: 4.837, h: 1.952 },
  { n: '02.', word: 'Colors', wordW: 6.401, head: 'Co', headW: 2.793,
    size: 110, y: 4.837, h: 1.952, headColor: C.white, numberColor: C.white },
  { n: '03.', word: 'Typography', wordW: 10.433, head: 'T', headW: 1.518,
    tail: { text: 'p', x: 3.448, w: 1.231 }, size: 100, y: 4.974, h: 1.784 },
  { n: '04.', word: 'Stationary', wordW: 9.297, head: 'S', headW: 1.497,
    tail: { text: 'a', x: 3.169, w: 1.242 }, size: 100, y: 4.983, h: 1.784 },
  { n: '05.', word: 'Digital', wordW: 5.673, head: 'D', headW: 1.454,
    tail: { text: 'g', x: 2.879, w: 1.245 }, size: 100, y: 4.97, h: 1.784,
    dot: { x: 4.075, y: 5.185, w: 0.354, h: 0.315 } },
  { n: '06.', word: 'Photography', wordW: 11.081, head: 'Ph', headW: 2.423,
    size: 100, y: 4.95, h: 1.784 },
];
function sectionDivider(pptx, i) {
  const sec = SECTIONS[i];
  const overlays = [{ text: sec.head, x: MARGIN_X, y: sec.y, w: sec.headW, color: sec.headColor }];
  if (sec.tail) overlays.push({ text: sec.tail.text, x: sec.tail.x, y: sec.y, w: sec.tail.w });
  dividerSlide(pptx, {
    fontSize: sec.size, wordY: sec.y, wordH: sec.h,
    lines: [{ text: sec.word, x: MARGIN_X, y: sec.y, w: sec.wordW }],
    overlays,
    dots: sec.dot ? [sec.dot] : [],
    number: sec.n, numberColor: sec.numberColor,
  });
}

// 4 — primary logo
function slidePrimaryLogo(pptx) {
  contentSlide(pptx, 'Primary Logo', 'logo section', (s) => {
    logoLockup(s, 7.156, 2.932, 5.01, C.blue, C.white);
  });
}

// 5 — submark logo
function slideSubmark(pptx) {
  contentSlide(pptx, 'Submark Logo', 'logo section', (s) => {
    drawPolys(s, MARK_ART, { x: 7.961, y: 2.272, w: 3.507, h: 2.957 }, C.blue, MARK_HOLES, C.white);
  });
}

// 6 — logo variation swatches
const VARIATIONS = [
  { x: 7.629, y: 1.615, w: 4.536, h: 1.481, fill: C.blue, logoW: 1.446 },
  { x: 7.629, y: 3.174, w: 2.228, h: 1.317, fill: C.black, logoW: 1.103 },
  { x: 9.937, y: 3.174, w: 2.228, h: 1.317, fill: C.pink, logoW: 1.103 },
  { x: 7.629, y: 4.568, w: 2.228, h: 1.317, fill: C.orange, logoW: 1.103 },
  { x: 9.937, y: 4.568, w: 2.228, h: 1.317, fill: C.green, logoW: 1.103 },
];
function slideLogoVariation(pptx) {
  contentSlide(pptx, 'Logo Variation', 'logo section', (s) => {
    VARIATIONS.forEach((v) => {
      rect(s, { x: v.x, y: v.y, w: v.w, h: v.h, fill: { color: v.fill }, line: { width: 0 } });
      logoLockup(s, v.x + (v.w - v.logoW) / 2, v.y + (v.h - v.logoW / 3.059) / 2, v.logoW, C.cream, v.fill);
    });
  });
}

// 7 — seamless pattern
function slidePattern(pptx) {
  contentSlide(pptx, 'Pattern', 'logo section', (s) => {
    rect(s, { x: 7.123, y: 0, w: 6.21, h: 7.5, fill: { color: C.blue }, line: { width: 0 } });
    // 5 columns x 8 rows of the tile, offset every other row, clipped at the slide edge.
    const tileW = 1.487, tileH = tileW / 0.976, x0 = 7.161, y0 = -0.35;
    for (let r = 0; r < 8; r += 1) {
      for (let c = 0; c < 5; c += 1) {
        const x = x0 + c * tileW + (r % 2 ? tileW * 0.47 : 0);
        const y = y0 + r * tileH * 0.585;
        if (x > SLIDE_W - 0.4 || y > SLIDE_H) continue;
        drawPolys(s, TILE_ART, { x, y, w: tileW, h: tileH }, C.cream, TILE_HOLES, C.blue);
      }
    }
  });
}

// 9 — colour palette
const PALETTE = [
  { name: 'Blue', hex: '#3E09FF', rgb: 'RGB: 62, 9, 255', fill: C.blue,
    x: 7.629, y: 0.999, w: 4.536, nameSize: 20, valueSize: 16, split: true },
  { name: 'Black', hex: '#181818', rgb: 'RGB: 24, 24, 24', fill: C.black,
    x: 7.629, y: 2.897, w: 2.228, nameSize: 18, valueSize: 14 },
  { name: 'Pink', hex: '#F442B9', rgb: 'RGB: 244, 66, 185', fill: C.pink,
    x: 9.937, y: 2.897, w: 2.228, nameSize: 18, valueSize: 14 },
  { name: 'Orange', hex: '#EB5C2A', rgb: 'RGB: 235, 92, 42', fill: C.orange,
    x: 7.629, y: 4.977, w: 2.228, nameSize: 18, valueSize: 14 },
  { name: 'Green', hex: '#0A3E29', rgb: 'RGB: 10, 62, 41', fill: C.green,
    x: 9.937, y: 4.977, w: 2.228, nameSize: 18, valueSize: 14 },
];
function slideColorPalette(pptx) {
  contentSlide(pptx, 'Color Palette', 'color section', (s) => {
    PALETTE.forEach((p) => {
      rect(s, { x: p.x, y: p.y, w: p.w, h: 0.354, fill: { color: p.fill }, line: { width: 0 } });
      textBox(s, p.name, {
        x: p.x, y: p.y + 0.494, w: 2, h: 0.438,
        fontFace: F.body, fontSize: p.nameSize, bold: true, color: C.black,
      });
      const style = { fontFace: F.body, fontSize: p.valueSize, bold: true, color: C.black };
      if (p.split) {
        // The blue swatch prints hex and rgb side by side under the wide bar.
        textBox(s, 'Hex: ' + p.hex, Object.assign({ x: p.x, y: p.y + 0.967, w: 1.81, h: 0.37 }, style));
        textBox(s, p.rgb, Object.assign({ x: 9.937, y: p.y + 0.967, w: 1.901, h: 0.37 }, style));
      } else {
        textBox(s, [{ text: 'Hex: ' + p.hex }, { text: p.rgb, options: { breakLine: false } }].map(
          (t, i) => ({ text: t.text, options: { breakLine: i === 0 } })
        ), Object.assign({ x: p.x, y: p.y + 0.946, w: 2.112, h: 0.572 }, style));
      }
    });
  });
}

// 10 — colour shades: 5 hues x 5 tints
const SHADES = [
  ['3E09FF', '5121FF', '5F33FF', '6B43FF', '7B57FF'],
  ['181818', '212121', '2A2A2A', '343434', '414141'],
  ['F442B9', 'F65CC3', 'F775CC', 'F888D3', 'F99DDA'],
  ['EB5C2A', 'EC6738', 'EE7348', 'F0845E', 'F29676'],
  ['0A3E29', '0C4C32', '0F6140', '12744C', '158558'],
];
function slideColorShades(pptx) {
  contentSlide(pptx, 'Color Shades', 'color section', (s) => {
    const x0 = 7.638, y0 = 0.774, cw = 0.709, ch = 1.102, gap = 1.2125;
    SHADES.forEach((row, r) => {
      row.forEach((hex, c) => {
        rect(s, {
          x: x0 + c * cw, y: y0 + r * gap, w: cw, h: ch,
          fill: { color: hex }, line: { width: 0 },
        });
      });
    });
  });
}

// 12 — typography overview
function slideTypography(pptx) {
  contentSlide(pptx, 'Typography', 'typography section', (s) => {
    textBox(s, 'Aa', {
      x: 7.285, y: 0.91, w: 2.434, h: 2.036,
      fontFace: F.body, fontSize: 115, bold: true, color: C.black,
    });
    textBox(s, 'Montserrat Bold', {
      x: 7.285, y: 2.947, w: 2.113, h: 0.37,
      fontFace: F.body, fontSize: 16, bold: true, color: C.black,
    });
    textBox(s, 'Aa', {
      x: 7.285, y: 4.096, w: 2.516, h: 2.036,
      fontFace: F.semi, fontSize: 115, italic: true, color: C.black,
    });
    textBox(s, 'Montserrat SemiBold Italic', {
      x: 7.285, y: 6.132, w: 3.345, h: 0.37,
      fontFace: F.semi, fontSize: 16, italic: true, color: C.black,
    });
  });
}

// 13, 14 — character sets
const GLYPHS = [
  { text: ['ABCDEFGHIJKLMNOPQRSTUV', 'WXYZ'], y: 1.912, w: 7.335, h: 1.178 },
  { text: ['abcdefghijklmnopqrstuvwxyz'], y: 3.283, w: 7.249, h: 0.64 },
  { text: ['1234567890'], y: 4.116, w: 2.935, h: 0.64 },
  { text: ['!@#$%^&*()_-=+[\\]{}|:;”<>’,./?`~'], y: 4.949, w: 7.335, h: 0.64 },
];
function specimenSlide(pptx, spec) {
  const s = contentSlide(pptx, null, 'typography section', null, { noBody: true });
  textBox(s, 'Aa', {
    x: MARGIN_X, y: 2.751, w: 2.134, h: 1.717,
    fontFace: spec.face, fontSize: 96, bold: spec.bold, italic: spec.italic, color: C.blue,
  });
  textBox(s, spec.label, {
    x: MARGIN_X, y: 4.683, w: 2.702, h: 0.64,
    fontFace: F.body, fontSize: 16, bold: true, italic: spec.italic, color: C.black,
  });
  GLYPHS.forEach((g) => {
    textBox(s, g.text.map((t, i) => ({ text: t, options: { breakLine: i < g.text.length - 1 } })), {
      x: 5.176, y: g.y, w: g.w, h: g.h,
      fontFace: spec.face, fontSize: 32, bold: spec.bold, italic: spec.italic, color: C.black,
    });
  });
}

// 16 — business card
function slideBusinessCard(pptx) {
  contentSlide(pptx, 'Business Card', 'stationary section', (s) => {
    rect(s, { x: 7.097, y: 2.278, w: 5.069, h: 2.944, fill: { color: C.blue }, line: { width: 0 } });
    logoLockup(s, 7.311, 2.516, 1.044, C.white, C.blue);
    textBox(s, 'Matthew Ryder', {
      x: 7.311, y: 4.207, w: 2.654, h: 0.438,
      fontFace: F.body, fontSize: 20, bold: true, color: C.white,
    });
    textBox(s, 'Creative Director', {
      x: 7.311, y: 4.645, w: 1.71, h: 0.303,
      fontFace: F.body, fontSize: 12, bold: true, color: C.white,
    });
    [['+123 456 7890', 4.056, 1.322], ['matthew@mail.com', 4.359, 1.881], ['hellomatthew.com', 4.662, 1.666]]
      .forEach(([t, y, w]) => {
        textBox(s, t, {
          x: 10.143, y, w, h: 0.28,
          fontFace: F.body, fontSize: 10.5, bold: true, color: C.white,
        });
      });
  });
}

// 17 — letterhead
const LETTER_BODY = [
  'Sit amet consectetur adipiscing elit duies. Cursus eget nunc scelerisque viverra mauris in ' +
    'aliquam sem fringilla. Lorem donecy massa sapien faucibus eti. Lacinia quis vel eros donec ac. ' +
    'Risus quis varius quam quisque id. ',
  '',
  'Vel orci porta noona pulvinar neque laoreet suspendisse interdum. Cursus eget nunc scelerisque ' +
    'viverra mauris in aliquam sem Fringilla. Lacinia quis vel eros donec ac.',
];
function slideLetterhead(pptx) {
  contentSlide(pptx, 'Letterhead', 'stationary section', (s) => {
    rect(s, { x: 7.839, y: 1.033, w: 4.134, h: 5.433, fill: { color: C.cream }, line: { width: 0 } });
    logoLockup(s, 8.065, 1.267, 0.866, C.blue, C.cream);
    const sig = (name, role, y) => {
      textBox(s, name, {
        x: 8.058, y, w: 1.75, h: 0.32,
        fontFace: F.xbold, fontSize: 13, color: C.blue,
      });
      textBox(s, role, {
        x: 8.058, y: y + 0.288, w: 1.5, h: 0.269,
        fontFace: F.semi, fontSize: 10, color: C.black,
      });
    };
    sig('Jonathan Ryder', 'CEO of Lycoss', 2.003);
    const meta = { fontFace: F.semi, fontSize: 10, color: C.black };
    textBox(s, '+123 456 7890', Object.assign({ x: 8.058, y: 2.802, w: 1.203, h: 0.269 }, meta));
    textBox(s, 'www.lycoss.com', Object.assign({ x: 8.058, y: 3.014, w: 1.471, h: 0.278 }, meta));
    textBox(s, '03.092023', Object.assign({ x: 10.786, y: 3.014, w: 0.938, h: 0.269, align: 'right' }, meta));
    textBox(s, LETTER_BODY.map((t, i) => ({ text: t, options: { breakLine: i < LETTER_BODY.length - 1 } })), {
      x: 8.058, y: 3.649, w: 3.666, h: 1.694,
      fontFace: F.body, fontSize: 8, color: C.ink, align: 'justify', lineSpacingMultiple: 1.5,
    });
    sig('Matthew Ryder', 'Creative Director', 5.651);
  });
}

// 19 — laptop mockup
function slideWebsite(pptx) {
  contentSlide(pptx, 'Website Used', 'digital section', (s) => {
    // Lid, bezel, screen, then the base wedge.
    rect(s, { x: 7.244, y: 1.952, w: 5.056, h: 3.42, fill: { color: '2B2B2B' }, line: { width: 0 } });
    photoPlaceholder(s, 7.365, 2.104, 4.816, 3.035);
    rect(s, { x: 6.667, y: 5.372, w: 6.211, h: 0.175, fill: { color: C.device }, line: { width: 0 } });
    s.addShape('roundRect', {
      x: 8.9, y: 5.372, w: 0.75, h: 0.1, rectRadius: 0.05,
      fill: { color: '9C9C9C' }, line: { width: 0 },
    });
  });
}

// 20 — phone mockup
function slideSocmed(pptx) {
  contentSlide(pptx, 'Socmed Used', 'digital section', (s) => {
    s.addShape('roundRect', {
      x: 8.275, y: 0.74, w: 3.056, h: 6.02, rectRadius: 0.33,
      fill: { color: '2B2B2B' }, line: { width: 0 },
    });
    s.addShape('roundRect', {
      x: 8.445, y: 0.849, w: 2.717, h: 5.733, rectRadius: 0.26,
      fill: { color: C.photo }, line: { width: 0 },
    });
    textBox(s, '[image]', {
      x: 8.445, y: 3.5, w: 2.717, h: 0.4, margin: 0,
      align: 'center', fontFace: F.body, fontSize: 11, color: '8A8A8A',
    });
  });
}

// 22, 23 — imagery grids
function slideImageryA(pptx) {
  contentSlide(pptx, 'Imagery', 'photography section', (s) => {
    photoPlaceholder(s, 7.298, 2.402, 4.867, 2.384);
    photoPlaceholder(s, 7.298, 4.867, 2.559, 1.736);
    photoPlaceholder(s, 9.921, 4.867, 2.244, 1.736);
  });
}
function slideImageryB(pptx) {
  contentSlide(pptx, 'Imagery', 'photography section', (s) => {
    photoPlaceholder(s, 7.298, 2.402, 2.556, 4.326);
    photoPlaceholder(s, 9.936, 2.402, 2.293, 2.265);
    textBox(s, 'Lorem donecy massa sapien faucibus eti. Lacinia quis vel eros donec ac. ', {
      x: 10.145, y: 5.825, w: 2.293, h: 0.903,
      fontFace: F.body, fontSize: 11, color: C.ink, lineSpacingMultiple: 1.5,
    });
  });
}

// 24 — thank you
function slideThankYou(pptx) {
  dividerSlide(pptx, {
    fontSize: 110, wordY: 4.837, wordH: 1.952,
    lines: [
      { text: 'Than', x: MARGIN_X, y: 2.885, w: 4.939 },
      { text: 'You', x: 1.163, y: 4.837, w: 4.5 },
    ],
    overlays: [
      { text: 'k', x: 5.823, y: 2.885, w: 1.352 },
      { text: 'Y', x: 1.163, y: 4.837, w: 1.657 },
      { text: 'o', x: 2.613, y: 4.837, w: 1.338 },
      { text: 'u', x: 3.735, y: 4.837, w: 1.275 },
    ],
  });
}

/* ------------------------------------------------------------------ *
 * Deck assembly
 * ------------------------------------------------------------------ */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'LYCOSS', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'LYCOSS';
  pptx.title = 'Brand Guidelines';

  slideTitle(pptx);              //  1
  slideContents(pptx);           //  2
  sectionDivider(pptx, 0);       //  3  Logo
  slidePrimaryLogo(pptx);        //  4
  slideSubmark(pptx);            //  5
  slideLogoVariation(pptx);      //  6
  slidePattern(pptx);            //  7
  sectionDivider(pptx, 1);       //  8  Colors
  slideColorPalette(pptx);       //  9
  slideColorShades(pptx);        // 10
  sectionDivider(pptx, 2);       // 11  Typography
  slideTypography(pptx);         // 12
  specimenSlide(pptx, { face: F.body, bold: true, label: 'Montserrat Bold' });          // 13
  specimenSlide(pptx, { face: F.semi, italic: true, label: 'Montserrat SemiBold\nItalic' }); // 14
  sectionDivider(pptx, 3);       // 15  Stationary
  slideBusinessCard(pptx);       // 16
  slideLetterhead(pptx);         // 17
  sectionDivider(pptx, 4);       // 18  Digital
  slideWebsite(pptx);            // 19
  slideSocmed(pptx);             // 20
  sectionDivider(pptx, 5);       // 21  Photography
  slideImageryA(pptx);           // 22
  slideImageryB(pptx);           // 23
  slideThankYou(pptx);           // 24

  return pptx;
}

build()
  .writeFile({ fileName: path.join(__dirname, '102666ea-6082-45c7-bfd0-0801f826bab6_grok_final.pptx') })
  .then((f) => console.log('wrote', f))
  .catch((e) => { console.error(e); process.exit(1); });
