/**
 * Pratomina fashion deck - 30 slides, 13.333 x 7.5 in - rebuilt with pptxgenjs.
 *
 * Run: node 040501e0-5135-4dd9-86fa-9c73bc346433_grok_final.js   ->   writes 040501e0-5135-4dd9-86fa-9c73bc346433_grok_final.pptx next to this file.
 *
 * Raster artwork from the original deck (the dark "REPLACE YOUR IMAGE HERE"
 * photo tiles and the device mock-up PNGs) is redrawn as flat placeholder
 * shapes; everything else - the organic blobs, torn-paper cards, washi tape,
 * leaf branch and line icons - is a real custom-geometry outline.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */
const C = {
  bg:       'EEECE1', // page background  (theme lt2)
  ink:      '404040', // headings         (tx1 @ 75%)
  body:     '808080', // body copy        (tx1 @ 50%)
  accent:   'E84C22', // theme accent1 - orange red
  peach:    'FF8427', // theme accent4 - used at 50% alpha for the blobs
  amber:    'FFBD47', // theme accent2 - price / contact icons
  white:    'FFFFFF',
  photo:    '474747', // average colour of the replaced photo tiles
  photoEdge:'3A3A3A',
  device:   '20232A', // device mock-up bezels
  grey:     'BFBFBF', // search bar outline
  greyIcon: 'D9D9D9',
};
const HEAD_FONT = 'Poppins';   // theme major latin
const BODY_FONT = 'Roboto';    // theme minor latin
const SCRIPT_FONT = 'Alex Brush';

const CARD_SHADOW = { type: 'outer', blur: 5, offset: 0, angle: 0, color: '000000', opacity: 0.2 };
const NO_LINE = { type: 'none' };

/* -------------------------------------------------------------- geometry
 * Outlines traced from the original deck, normalised to the unit square and
 * stored as flat [x0,y0, x1,y1, ...] runs - one array per sub-path.
 */
const PATHS = {
  leaf:       [[0.165,0.64,0.144,0.48,0.003,0.344,0.03,0.512,0.149,0.67,0.067,1,0.088,0.988,0.139,0.773,0.696,0.553,0.407,0.604,0.145,0.743,0.228,0.572,0.36,0.505,0.798,0.377,0.475,0.431,0.259,0.528,0.293,0.48,0.532,0.317,0.837,0.257,1,0.196,0.78,0.216,0.631,0.267,0.778,0.161,0.846,0.001,0.674,0.101,0.54,0.3,0.31,0.448,0.444,0.275,0.456,0.186,0.433,0.121,0.287,0.31,0.268,0.371,0.283,0.468,0.165,0.64],[0.582,0.264,0.691,0.116,0.794,0.034,0.582,0.264],[0.778,0.038,0.69,0.113,0.583,0.257,0.654,0.13,0.778,0.038],[0.833,0.015,0.635,0.229,0.833,0.015],[0.839,0.016,0.711,0.182,0.574,0.285,0.839,0.016],[0.811,0.09,0.747,0.173,0.611,0.27,0.745,0.153,0.841,0.019,0.811,0.09],[0.637,0.259,0.798,0.116,0.73,0.201,0.637,0.259],[0.552,0.298,0.815,0.021,0.552,0.298],[0.902,0.203,0.672,0.256,0.902,0.203],[0.947,0.201,0.623,0.278,0.947,0.201],[0.978,0.199,0.639,0.275,0.978,0.199],[0.935,0.215,0.639,0.285,0.935,0.215],[0.619,0.291,0.921,0.222,0.619,0.291],[0.629,0.286,0.58,0.296,0.951,0.207,0.629,0.286],[0.391,0.399,0.466,0.349,0.391,0.399],[0.31,0.294,0.287,0.42,0.28,0.357,0.31,0.294],[0.426,0.146,0.293,0.431,0.315,0.297,0.401,0.186,0.332,0.302,0.426,0.146],[0.42,0.239,0.322,0.409,0.42,0.239],[0.407,0.254,0.377,0.293,0.43,0.184,0.407,0.254],[0.338,0.347,0.31,0.395,0.428,0.163,0.338,0.347],[0.432,0.172,0.434,0.133,0.432,0.172],[0.346,0.404,0.42,0.297,0.45,0.204,0.429,0.298,0.346,0.404],[0.318,0.425,0.42,0.252,0.444,0.158,0.428,0.274,0.35,0.395,0.428,0.248,0.318,0.425],[0.295,0.456,0.334,0.361,0.295,0.456],[0.328,0.444,0.354,0.423,0.328,0.444],[0.305,0.464,0.32,0.452,0.305,0.464],[0.418,0.476,0.346,0.499,0.733,0.394,0.418,0.476],[0.312,0.514,0.283,0.528,0.312,0.514],[0.749,0.383,0.441,0.449,0.244,0.544,0.442,0.444,0.749,0.383],[0.747,0.389,0.456,0.454,0.285,0.523,0.464,0.444,0.764,0.383,0.747,0.389],[0.318,0.515,0.76,0.388,0.318,0.515],[0.224,0.564,0.248,0.545,0.224,0.564],[0.212,0.583,0.232,0.562,0.212,0.583],[0.124,0.814,0.116,0.851,0.124,0.814],[0.367,0.624,0.212,0.702,0.367,0.624],[0.283,0.71,0.658,0.566,0.283,0.71],[0.143,0.762,0.678,0.556,0.143,0.762],[0.153,0.751,0.428,0.624,0.652,0.56,0.139,0.761,0.153,0.751],[0.181,0.733,0.467,0.601,0.644,0.56,0.481,0.602,0.181,0.733],[0.167,0.738,0.395,0.614,0.58,0.567,0.167,0.738],[0.143,0.627,0.024,0.373,0.077,0.446,0.143,0.627],[0.143,0.599,0.086,0.452,0.143,0.599],[0.033,0.492,0.059,0.537,0.033,0.492],[0.024,0.457,0.033,0.488,0.024,0.457],[0.02,0.419,0.134,0.62,0.055,0.523,0.02,0.419],[0.012,0.362,0.124,0.592,0.012,0.362],[0.159,0.577,0.088,0.422,0.144,0.491,0.159,0.577],[0.155,0.634,0.113,0.49,0.02,0.362,0.131,0.501,0.155,0.634]], // 212 pts / 48 sub
  blobRight:  [[1,0,1,1,0.228,1,0.057,0.85,0.027,0.812,0.007,0.773,0,0.734,0.009,0.697,0.038,0.662,0.09,0.63,0.156,0.604,0.47,0.52,0.57,0.477,0.648,0.426,0.709,0.369,0.755,0.308,0.79,0.244,0.83,0.15,0.865,0.091,0.92,0.037,1,0]], // 21 pts / 1 sub
  blobTop:    [[0,0,1,0,0.992,0.287,0.982,0.393,0.963,0.494,0.931,0.588,0.882,0.673,0.833,0.727,0.777,0.771,0.716,0.805,0.651,0.83,0.517,0.862,0.314,0.888,0.181,0.916,0.095,0.948,0,1,0,0]], // 17 pts / 1 sub
  tape:       [[0.979,0.177,0.983,0.25,0.968,0.287,0.965,0.334,0.978,0.477,0.972,0.591,0.983,0.665,0.97,0.72,0.991,0.933,0.985,0.951,0.996,1,0.008,1,0.006,0.97,0.023,0.915,0.013,0.902,0.021,0.823,0.009,0.794,0.007,0.749,0.019,0.72,0.016,0.678,0.028,0.61,0.009,0.512,0.023,0.433,0.009,0.306,0.019,0.165,0.002,0,1,0,0.979,0.177]], // 28 pts / 1 sub
  blobHill:   [[0.38,0,0.454,0.009,0.522,0.037,0.583,0.086,0.635,0.162,0.672,0.253,0.747,0.564,0.797,0.719,0.877,0.864,0.943,0.945,1,1,0,1,0,0.127,0.152,0.054,0.227,0.026,0.304,0.007,0.38,0]], // 17 pts / 1 sub
  torn:       [[0,0,1,0,1,0.954,0.979,0.973,0.961,0.977,0.932,0.996,0.902,0.994,0.886,1,0.862,0.991,0.813,0.988,0.742,0.998,0.699,0.981,0.632,0.983,0.608,0.968,0.528,0.954,0.515,0.94,0.484,0.93,0.427,0.892,0.39,0.883,0.362,0.889,0.354,0.883,0.331,0.882,0.263,0.829,0.225,0.837,0.223,0.83,0.192,0.819,0.062,0.814,0.004,0.821,0,0]], // 29 pts / 1 sub
  tornTall:   [[0,0,1,0,1,0.964,0.979,0.979,0.932,0.997,0.886,1,0.813,0.991,0.742,0.998,0.699,0.985,0.632,0.987,0.608,0.975,0.528,0.964,0.427,0.916,0.331,0.908,0.263,0.866,0.225,0.873,0.223,0.867,0.192,0.859,0.062,0.855,0.004,0.861,0,0]], // 21 pts / 1 sub
  quarterA:   [[0.288,0,0.478,0.032,0.648,0.124,0.792,0.265,0.903,0.449,0.975,0.665,1,0.906,0.996,1,0,1,0,0.079,0.068,0.044,0.154,0.016,0.288,0]], // 13 pts / 1 sub
  quarterB:   [[0.236,0,0.439,0.032,0.621,0.121,0.776,0.26,0.896,0.439,0.973,0.651,1,0.887,0.993,1,0,1,0,0.045,0.106,0.013,0.236,0]], // 12 pts / 1 sub
  quarterC:   [[0.236,0,0.428,0.033,0.601,0.128,0.751,0.276,0.871,0.468,0.956,0.696,0.998,0.952,1,1,0,1,0,0.053,0.107,0.015,0.236,0]], // 12 pts / 1 sub
  blobCorner: [[0,0,0.306,0.01,0.411,0.02,0.512,0.04,0.606,0.073,0.69,0.123,0.744,0.173,0.788,0.23,0.821,0.293,0.847,0.359,0.879,0.497,0.905,0.706,0.932,0.842,0.959,0.92,1,1,0,1,0,0]], // 17 pts / 1 sub
  blobSide:   [[0,0,1,0,1,1,0.726,0.973,0.599,0.952,0.48,0.923,0.374,0.886,0.284,0.84,0.218,0.783,0.177,0.714,0.167,0.643,0.175,0.569,0.211,0.42,0.218,0.346,0.211,0.289,0.194,0.234,0.166,0.18,0.129,0.128,0.084,0.076,0,0]], // 20 pts / 1 sub
  blobWave:   [[0,0,1,0,1,1,0.919,0.828,0.866,0.74,0.808,0.66,0.744,0.59,0.672,0.532,0.592,0.486,0.503,0.453,0.317,0.403,0.228,0.372,0.148,0.325,0.079,0.255,0.043,0.192,0.01,0.084,0,0]], // 17 pts / 1 sub
  blobRight2: [[1,0,1,1,0.227,1,0.057,0.852,0.027,0.813,0.007,0.774,0,0.736,0.009,0.698,0.038,0.663,0.09,0.631,0.156,0.605,0.47,0.521,0.57,0.478,0.648,0.427,0.709,0.37,0.755,0.309,0.79,0.245,0.83,0.151,0.865,0.091,0.92,0.037,1,0]], // 21 pts / 1 sub
  blobWide:   [[0,0,1,0,0.958,0.109,0.926,0.246,0.839,0.768,0.813,0.868,0.782,0.945,0.744,0.991,0.712,1,0.678,0.988,0.643,0.959,0.608,0.915,0.538,0.798,0.402,0.533,0.333,0.429,0.285,0.383,0.142,0.283,0.095,0.228,0.042,0.13,0,0]], // 20 pts / 1 sub
  blobTop2:   [[0.005,0,1,0,1,1,0.906,0.945,0.814,0.908,0.704,0.881,0.478,0.849,0.369,0.824,0.266,0.781,0.174,0.71,0.105,0.62,0.055,0.513,0.023,0.393,0.005,0.264,0,0.132,0.005,0]], // 16 pts / 1 sub
  seed:       [[0.034,0.741,0.097,0.783,0.177,0.8,0.447,0.803,0.538,0.833,0.672,0.907,0.759,0.93,0.836,0.913,0.897,0.867,0.939,0.805,0.956,0.736,0.944,0.673,0.861,0.568,0.846,0.499,0.862,0.438,0.932,0.335,0.943,0.284,0.941,0.23,0.903,0.128,0.828,0.053,0.779,0.033,0.675,0.031,0.582,0.075,0.503,0.182,0.467,0.206,0.411,0.214,0.219,0.204,0.119,0.243,0.093,0.275,0.088,0.329,0.137,0.468,0.139,0.517,0.114,0.554,0.051,0.594,0.02,0.638,0.014,0.692,0.034,0.741]], // 37 pts / 1 sub
  magnifier:  [[0.438,0,0.554,0.015,0.659,0.059,0.748,0.127,0.817,0.215,0.861,0.32,0.877,0.438,0.86,0.562,0.811,0.673,0.988,0.861,0.997,0.931,0.953,0.988,0.907,1,0.861,0.988,0.673,0.805,0.482,0.874,0.323,0.861,0.218,0.817,0.129,0.748,0.06,0.659,0.016,0.554,0,0.438,0.016,0.32,0.06,0.215,0.129,0.127,0.218,0.059,0.323,0.015,0.438,0],[0.435,0.12,0.351,0.131,0.212,0.212,0.131,0.351,0.119,0.435,0.131,0.519,0.212,0.658,0.351,0.739,0.435,0.75,0.519,0.739,0.659,0.658,0.74,0.519,0.751,0.435,0.74,0.351,0.659,0.212,0.519,0.131,0.435,0.12]], // 45 pts / 2 sub
  arrowDisc:  [[0.499,0,0.632,0.018,0.752,0.068,0.853,0.146,0.932,0.248,0.982,0.367,1,0.5,0.982,0.633,0.932,0.752,0.853,0.854,0.752,0.932,0.632,0.982,0.499,1,0.382,0.986,0.275,0.947,0.18,0.886,0.102,0.806,0.044,0.709,0.001,0.529,0.568,0.529,0.404,0.693,0.438,0.727,0.66,0.505,0.438,0.283,0.404,0.317,0.568,0.481,0,0.481,0.008,0.399,0.044,0.291,0.102,0.194,0.18,0.114,0.275,0.053,0.382,0.014,0.499,0]], // 34 pts / 1 sub
  icoPhone:   [[0.51,1,0.16,0.96,0.07,0.89,0.02,0.78,0.02,0.22,0.09,0.09,0.22,0.02,0.78,0.02,0.91,0.09,0.98,0.22,1,0.62,0.96,0.66,0.92,0.63,0.91,0.24,0.82,0.12,0.72,0.08,0.28,0.08,0.19,0.11,0.12,0.18,0.08,0.31,0.08,0.69,0.11,0.8,0.18,0.88,0.28,0.92,0.72,0.92,0.82,0.88,0.92,0.78,0.97,0.8,0.94,0.88,0.84,0.96,0.51,1],[0.74,0.5,0.67,0.33,0.5,0.26,0.33,0.33,0.25,0.5,0.33,0.67,0.5,0.74,0.67,0.67,0.74,0.5],[0.66,0.5,0.62,0.62,0.5,0.67,0.38,0.62,0.33,0.5,0.38,0.38,0.5,0.33,0.62,0.38,0.66,0.5],[0.77,0.18,0.71,0.24,0.77,0.29,0.83,0.24,0.77,0.18]], // 54 pts / 4 sub
  icoInsta:   [[0.5,1,0.28,0.95,0.03,0.98,0.06,0.74,0,0.5,0.04,0.31,0.15,0.15,0.31,0.04,0.5,0,0.69,0.04,0.85,0.15,0.96,0.31,1,0.5,0.96,0.69,0.91,0.79,0.87,0.79,0.84,0.75,0.91,0.58,0.92,0.45,0.87,0.29,0.8,0.2,0.66,0.11,0.5,0.08,0.39,0.09,0.2,0.2,0.09,0.39,0.08,0.57,0.14,0.73,0.1,0.9,0.28,0.87,0.5,0.92,0.73,0.86,0.77,0.87,0.77,0.91,0.72,0.95,0.5,1],[0.39,0.3,0.32,0.29,0.27,0.36,0.28,0.44,0.35,0.55,0.51,0.68,0.63,0.71,0.71,0.68,0.73,0.6,0.62,0.55,0.55,0.61,0.43,0.52,0.39,0.46,0.43,0.4,0.39,0.3]], // 51 pts / 2 sub
  icoPin:     [[0.5,0.82,0.33,0.66,0.26,0.44,0.33,0.26,0.5,0.19,0.68,0.26,0.75,0.44,0.67,0.66,0.5,0.82],[0.5,0.27,0.38,0.32,0.33,0.44,0.39,0.59,0.5,0.73,0.62,0.6,0.67,0.44,0.62,0.32,0.5,0.27],[0.5,0.35,0.49,0.41,0.45,0.42,0.42,0.39,0.41,0.45,0.44,0.51,0.5,0.53,0.56,0.51,0.59,0.45,0.57,0.38,0.5,0.35],[0.76,0.93,0.78,0.88,0.75,0.86,0.5,0.92,0.39,0.91,0.2,0.8,0.09,0.61,0.09,0.39,0.2,0.2,0.39,0.09,0.5,0.08,0.71,0.14,0.86,0.29,0.91,0.39,0.91,0.58,0.84,0.75,0.89,0.79,0.96,0.69,1,0.55,0.98,0.37,0.93,0.25,0.8,0.1,0.63,0.02,0.43,0,0.25,0.07,0.1,0.2,0.02,0.37,0,0.57,0.07,0.75,0.2,0.9,0.37,0.98,0.59,0.99,0.76,0.93]], // 62 pts / 4 sub
  icoBank:    [[0.16,0.86,0.11,0.92,0.89,0.92,0.84,0.86,0.16,0.86],[0.58,0.38,0.53,0.43,0.53,0.79,0.63,0.79,0.63,0.43,0.58,0.38],[0.41,0.38,0.36,0.43,0.36,0.79,0.46,0.79,0.46,0.43,0.41,0.38],[0.23,0.38,0.18,0.43,0.18,0.79,0.28,0.79,0.28,0.43,0.23,0.38],[0.5,0.17,0.54,0.21,0.5,0.25,0.46,0.21,0.5,0.17],[0.5,0.08,0.08,0.26,0.08,0.3,0.92,0.3,0.92,0.26,0.5,0.08],[0.51,0,1,0.22,0.99,0.36,0.88,0.38,0.87,0.66,0.81,0.64,0.8,0.41,0.74,0.38,0.71,0.43,0.71,0.79,0.89,0.79,1,0.98,0.96,1,0.02,0.99,0,0.94,0.1,0.8,0.1,0.38,0.01,0.36,0,0.22,0.51,0]], // 54 pts / 7 sub
  icoBox:     [[0.907,0.026,0.666,0.003,0.554,0.026,0.479,0.074,0.461,0.121,0.461,0.485,0.334,0.456,0.205,0.456,0.093,0.481,0.018,0.529,0,0.576,0,0.879,0.035,0.943,0.205,0.997,0.364,0.993,0.5,0.946,0.607,0.988,0.73,1,0.882,0.982,0.982,0.926,1,0.879,1,0.121,0.982,0.074,0.907,0.026],[0.593,0.096,0.78,0.08,0.92,0.121,0.827,0.156,0.681,0.162,0.541,0.121,0.593,0.096],[0.132,0.55,0.319,0.533,0.459,0.574,0.366,0.611,0.22,0.617,0.08,0.576,0.132,0.55],[0.407,0.904,0.22,0.92,0.086,0.883,0.078,0.817,0.31,0.847,0.367,0.836,0.383,0.8,0.35,0.766,0.22,0.768,0.103,0.742,0.078,0.726,0.078,0.665,0.27,0.698,0.461,0.664,0.461,0.877,0.407,0.904],[0.868,0.904,0.681,0.92,0.539,0.877,0.539,0.817,0.771,0.848,0.828,0.837,0.843,0.801,0.81,0.767,0.657,0.767,0.539,0.726,0.539,0.665,0.771,0.696,0.828,0.685,0.843,0.649,0.81,0.615,0.681,0.617,0.564,0.59,0.539,0.574,0.539,0.511,0.771,0.542,0.828,0.531,0.843,0.495,0.81,0.461,0.681,0.463,0.564,0.436,0.539,0.42,0.539,0.36,0.771,0.391,0.828,0.38,0.843,0.344,0.81,0.311,0.681,0.312,0.564,0.286,0.539,0.269,0.539,0.21,0.73,0.242,0.922,0.21,0.922,0.877,0.868,0.904],[0.076,0.161,0.076,0.105,0.161,0.02,0.208,0.001,0.28,0.02,0.377,0.133,0.365,0.161,0.328,0.171,0.26,0.109,0.258,0.346,0.221,0.375,0.183,0.346,0.182,0.109,0.118,0.169,0.076,0.161]], // 108 pts / 6 sub
  icoCoins:   [[0.68,0.729,0.723,0.768,0.68,0.807,0.637,0.768,0.68,0.729],[0.347,0.68,0.354,0.708,0.387,0.7,0.347,0.68],[0.754,0.481,0.754,0.523,0.781,0.502,0.754,0.481],[0.136,0.475,0.179,0.514,0.136,0.553,0.094,0.514,0.136,0.475],[0.499,0.337,0.33,0.42,0.326,0.5,0.331,0.578,0.497,0.662,0.665,0.581,0.669,0.5,0.665,0.42,0.499,0.337],[0.352,0.288,0.344,0.321,0.388,0.299,0.352,0.288],[0.672,0.189,0.714,0.229,0.672,0.268,0.629,0.229,0.672,0.189],[0.499,0,0.592,0.032,0.645,0.085,0.649,0.124,0.602,0.146,0.529,0.086,0.48,0.082,0.422,0.13,0.376,0.213,0.615,0.302,0.842,0.447,0.914,0.322,0.89,0.288,0.787,0.258,0.79,0.203,0.869,0.199,0.936,0.223,0.985,0.272,1,0.334,0.975,0.408,0.902,0.503,0.991,0.653,0.985,0.726,0.939,0.778,0.849,0.808,0.807,0.784,0.819,0.739,0.907,0.692,0.891,0.624,0.841,0.557,0.611,0.699,0.376,0.784,0.425,0.873,0.484,0.92,0.529,0.914,0.602,0.854,0.649,0.876,0.645,0.915,0.592,0.968,0.499,1,0.382,0.949,0.292,0.8,0.084,0.787,0.014,0.729,0,0.674,0.018,0.602,0.057,0.586,0.092,0.609,0.101,0.706,0.172,0.727,0.27,0.724,0.241,0.5,0.268,0.271,0.158,0.27,0.1,0.292,0.088,0.329,0.11,0.397,0.077,0.433,0.045,0.428,0.021,0.396,0.006,0.291,0.041,0.236,0.11,0.2,0.293,0.195,0.393,0.041,0.499,0]], // 102 pts / 8 sub
  icoAtom:    [[0.371,0.276,0.4,0.325,0.421,0.444,0.481,0.503,0.52,0.503,0.592,0.43,0.603,0.327,0.67,0.208,0.678,0.155,0.655,0.074,0.551,0.005,0.455,0.004,0.379,0.038,0.337,0.089,0.322,0.155,0.371,0.276],[0.521,0.355,0.521,0.395,0.482,0.395,0.482,0.355,0.521,0.355],[0.502,0.078,0.562,0.091,0.594,0.126,0.595,0.185,0.54,0.277,0.464,0.277,0.402,0.169,0.416,0.11,0.502,0.078],[0.884,0.498,0.995,0.359,0.996,0.322,0.753,0.238,0.715,0.266,0.723,0.304,0.894,0.361,0.804,0.473,0.5,0.559,0.196,0.473,0.106,0.361,0.277,0.304,0.281,0.257,0.237,0.24,0.004,0.322,0.005,0.359,0.116,0.499,0,0.671,0.021,0.7,0.117,0.731,0.132,0.849,0.199,0.921,0.45,0.995,0.554,0.995,0.831,0.902,0.881,0.818,0.885,0.73,0.993,0.688,0.996,0.647,0.884,0.498],[0.177,0.548,0.425,0.62,0.348,0.724,0.104,0.645,0.177,0.548],[0.252,0.859,0.203,0.816,0.195,0.756,0.369,0.808,0.461,0.703,0.461,0.918,0.252,0.859],[0.807,0.784,0.766,0.852,0.539,0.918,0.539,0.704,0.623,0.807,0.807,0.756,0.807,0.784],[0.65,0.724,0.575,0.62,0.825,0.549,0.897,0.645,0.65,0.724]], // 84 pts / 8 sub
  icoUser:    [[0.474,0.922,0.111,0.921,0.078,0.879,0.11,0.78,0.225,0.631,0.393,0.544,0.571,0.522,0.634,0.495,0.729,0.4,0.766,0.266,0.729,0.132,0.634,0.036,0.5,0,0.366,0.036,0.271,0.132,0.244,0.195,0.238,0.308,0.263,0.386,0.341,0.478,0.202,0.549,0.098,0.651,0.038,0.748,0,0.875,0.027,0.956,0.101,0.999,0.484,0.999,0.511,0.951,0.474,0.922],[0.313,0.266,0.367,0.133,0.5,0.078,0.633,0.133,0.688,0.266,0.634,0.397,0.504,0.453,0.443,0.444,0.364,0.395,0.313,0.266],[0.859,0.764,0.831,0.801,0.736,0.791,0.736,0.656,0.791,0.656,0.803,0.725,0.84,0.73,0.859,0.764],[0.764,0.527,0.701,0.536,0.597,0.597,0.536,0.701,0.527,0.764,0.536,0.826,0.597,0.931,0.701,0.992,0.764,1,0.826,0.992,0.931,0.931,0.992,0.826,1,0.764,0.992,0.701,0.931,0.597,0.826,0.536,0.764,0.527],[0.764,0.922,0.652,0.875,0.605,0.764,0.652,0.652,0.764,0.605,0.875,0.652,0.922,0.764,0.875,0.875,0.764,0.922]], // 72 pts / 5 sub
};

/** Turn a PATHS entry into pptxgenjs `points`, scaled into a w x h box. */
function poly(name, w, h) {
  const pts = [];
  for (const flatPts of PATHS[name]) {
    for (let i = 0; i < flatPts.length; i += 2) {
      pts.push({ x: +(flatPts[i] * w).toFixed(4), y: +(flatPts[i + 1] * h).toFixed(4), moveTo: i === 0 });
    }
    pts.push({ close: true });
  }
  return pts;
}

/** Custom-geometry shape helper: `shape(slide, 'leaf', {x,y,w,h,...})`. */
function shape(slide, name, o) {
  slide.addShape('custGeom', Object.assign({ points: poly(name, o.w, o.h), line: NO_LINE }, o));
}

/** Solid peach blob at 50% transparency - by far the most common decoration. */
function blob(slide, name, o) {
  shape(slide, name, Object.assign({ fill: { color: C.peach, transparency: 50 } }, o));
}

/** Same outline drawn as a thin accent outline instead of a fill. */
function blobOutline(slide, name, o) {
  shape(slide, name, Object.assign({ fill: { type: 'none' }, line: { color: C.accent, width: 1 } }, o));
}

/** Torn-paper note card (white, soft shadow) with its washi tape strip. */
function tornCard(slide, o) {
  shape(slide, o.tall ? 'tornTall' : 'torn',
    { x: o.x, y: o.y, w: o.w, h: o.h, rotate: o.rotate, fill: { color: C.white }, shadow: CARD_SHADOW });
  if (o.tape) shape(slide, 'tape', Object.assign({ w: 1.405, h: 0.44, fill: { color: C.accent, transparency: 20 } }, o.tape));
}

/** Flat stand-in for one of the deck's dark photo tiles. */
function photo(slide, x, y, w, h) {
  slide.addShape('rect', { x, y, w, h, fill: { color: C.photo }, line: { color: C.photoEdge, width: 0.75 } });
  slide.addText('[image]', {
    x, y, w, h, align: 'center', valign: 'middle', margin: 0,
    fontFace: BODY_FONT, fontSize: Math.max(8, Math.min(13, Math.round(Math.min(w, h) * 4.5))),
    color: 'BFBFBF',
  });
}

/** The search bar + arrow button that the slide master paints top-left. */
function searchBar(slide) {
  slide.addShape('roundRect', { x: 0.245, y: 0.197, w: 1.691, h: 0.24, rectRadius: 0.12,
    fill: { type: 'none' }, line: { color: C.grey, width: 1 } });
  shape(slide, 'magnifier', { x: 0.353, y: 0.244, w: 0.143, h: 0.143, fill: { color: C.greyIcon } });
  slide.addText('Suggestion Search . . .', { x: 0.578, y: 0.253, w: 1.113, h: 0.135,
    fontFace: BODY_FONT, fontSize: 8, color: C.grey, margin: 0, valign: 'top', wrap: false });
  slide.addShape('ellipse', { x: 2.017, y: 0.197, w: 0.24, h: 0.24, fill: { color: C.accent }, line: NO_LINE });
  shape(slide, 'arrowDisc', { x: 2.047, y: 0.197, w: 0.24, h: 0.24, fill: { color: C.white } });
}

/** Small accent bullet dot used beside "DESCRIPTION" blocks. */
function dot(slide, x, y) {
  slide.addShape('ellipse', { x, y, w: 0.136, h: 0.136, fill: { color: C.accent }, line: NO_LINE });
}

/** "Learn more" (filled) + "improve" (outline) pill pair. */
function pills(slide, x, y) {
  slide.addShape('roundRect', { x, y, w: 1.241, h: 0.281, rectRadius: 0.14, fill: { color: C.accent }, line: NO_LINE });
  slide.addText('Learn more', { x, y: y + 0.006, w: 1.241, h: 0.269, align: 'center', valign: 'middle',
    fontFace: BODY_FONT, fontSize: 10, color: C.white, margin: 0 });
  slide.addShape('roundRect', { x: x + 1.581, y, w: 1.241, h: 0.281, rectRadius: 0.14,
    fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
  slide.addText('improve', { x: x + 1.581, y: y + 0.006, w: 1.241, h: 0.269, align: 'center', valign: 'middle',
    fontFace: BODY_FONT, fontSize: 10, color: C.accent, margin: 0 });
}

/* ------------------------------------------------ device mock-up stand-ins */
function laptopMockup(slide, o) {
  slide.addShape('roundRect', { x: o.x, y: o.y, w: o.w, h: o.h * 0.94, rectRadius: 0.1, fill: { color: C.device }, line: NO_LINE });
  photo(slide, o.x + o.w * 0.103, o.y + o.h * 0.039, o.w * 0.792, o.h * 0.825);
  slide.addShape('trapezoid', { x: o.x - o.w * 0.077, y: o.y + o.h * 0.935, w: o.w * 1.154, h: o.h * 0.065,
    fill: { color: C.device }, line: NO_LINE, flipV: true });
}

function phoneMockup(slide, o) {
  slide.addShape('roundRect', { x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.28, fill: { color: C.device }, line: NO_LINE });
  photo(slide, o.x + o.w * 0.048, o.y + o.h * 0.039, o.w * 0.904, o.h * 0.922);
}

function tabletMockup(slide, o) {
  slide.addShape('roundRect', { x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.16, fill: { color: C.device }, line: NO_LINE });
  photo(slide, o.x + o.w * 0.033, o.y + o.h * 0.043, o.w * 0.908, o.h * 0.885);
}

/* --------------------------------------------------------- text shorthands */
const TITLE = { fontFace: HEAD_FONT, fontSize: 28, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] };
const TITLE24 = Object.assign({}, TITLE, { fontSize: 24 });
const LEAD = { fontFace: BODY_FONT, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] };
const SMALL = { fontFace: BODY_FONT, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] };
const LABEL = { fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] };
const HILITE = { fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.accent, lineSpacingMultiple: 1.5, valign: 'top', margin: 0 };
const SCRIPT = { fontFace: SCRIPT_FONT, fontSize: 40, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] };

/* ------------------------------------------------------------ boilerplate */
const LOREM = 'PLACEHOLDER' +
  'PLACEHOLDER' +
  'PLACEHOLDER' +
  'PLACEHOLDER';
const LOREM_SI = 'PLACEHOLDER' +
  'PLACEHOLDER' +
  'PLACEHOLDER';
const ASUS = 'PLACEHOLDER' +
  'turpis antes dolor elementum';

/** First `count` words of one of the lorem strings. */
function words(src, count) { return src.split(' ').slice(0, count).join(' '); }


/** Category labels for the stock chart on slide 28. */
const STOCK_DATES = ['5/1/2002', '6/1/2002', '7/1/2002', '8/1/2002', '9/1/2002'];

/* ============================== the 30 slides ============================= */
const SLIDES = [];

/* ---  1. Cover - PRATOMINA / Fashion Collection --- */
SLIDES.push(function slide01(s, pres) {
  shape(s, 'blobRight', { x: 10.979, y: 2.197, w: 2.354, h: 5.303, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  photo(s, 6.667, 0, 5.342, 7.5);
  shape(s, 'quarterB', { x: 0, y: 5.899, w: 2.229, h: 1.601, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  s.addText('PRATOMINA', { x: 0.912, y: 1.536, w: 2.593, h: 0.572, fontFace: HEAD_FONT, fontSize: 28, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'leaf', { x: 12.368, y: 1.422, w: 1.274, h: 2.816, rotate: 323.83, fill: { color: C.accent }, line: NO_LINE });
  shape(s, 'torn', { x: 3.182, y: 4.675, w: 4.168, h: 2.366, rotate: 354.36, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  shape(s, 'tape', { x: 4.407, y: 4.431, w: 1.405, h: 0.44, rotate: 352.64, fill: { color: C.accent, transparency: 20 }, line: NO_LINE });
  s.addText('PLACEHOLDER',
    { x: 0.912, y: 2.187, w: 4.597, h: 1.279, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'quarterA', { x: -0.188, y: 5.696, w: 2.726, h: 1.929, fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
  searchBar(s);
  s.addText('Fashion', { x: 3.587, y: 5.112, w: 2.016, h: 0.909, rotate: 353.4, fontFace: SCRIPT_FONT, fontSize: 48, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('Collection', { x: 4.79, y: 5.618, w: 2.388, h: 0.909, rotate: 353.4, fontFace: SCRIPT_FONT, fontSize: 48, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
});

/* ---  2. We are fashion design --- */
SLIDES.push(function slide02(s, pres) {
  photo(s, 4.471, 0.896, 2.625, 2.625);
  shape(s, 'blobCorner', { x: 0.003, y: 5.045, w: 2.101, h: 2.565, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'blobCorner', { x: -0.128, y: 4.884, w: 2.423, h: 2.762, fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
  shape(s, 'blobRight', { x: 10.979, y: 2.197, w: 2.354, h: 5.303, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  s.addText('WE ARE\nFASHION DESIGN', { x: 8.091, y: 1.223, w: 3.019, h: 0.909, fontFace: HEAD_FONT, fontSize: 24, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(LOREM,
    { x: 4.843, y: 5.045, w: 7.882, h: 1.279, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(ASUS,
    { x: 8.168, y: 2.277, w: 3.924, h: 0.88, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.accent, lineSpacingMultiple: 1.5, valign: 'top', margin: 0 });
  shape(s, 'leaf', { x: 12.368, y: 0.219, w: 1.274, h: 2.816, rotate: 323.83, fill: { color: C.accent }, line: NO_LINE });
  s.addText('Pratomina Clothes', { x: 4.843, y: 4.301, w: 4.053, h: 0.774, fontFace: SCRIPT_FONT, fontSize: 40, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  photo(s, 0.609, 0.896, 3.6, 6.021);
});

/* ---  3. Creat your style --- */
SLIDES.push(function slide03(s, pres) {
  photo(s, 10.548, 4.252, 2.202, 2.202);
  shape(s, 'blobTop', { x: 9.438, y: 0, w: 3.896, h: 3.604, rotate: 180, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'blobTop', { x: 0, y: 5.353, w: 3.062, h: 2.13, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  s.addText('CREAT YOUR STYLE', { x: 5.335, y: 1.361, w: 3.896, h: 0.572, fontFace: HEAD_FONT, fontSize: 28, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 41),
    { x: 5.335, y: 1.992, w: 5.878, h: 1.279, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('CLASSIC STYLE', { x: 5.335, y: 3.76, w: 1.441, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 11),
    { x: 5.335, y: 4.063, w: 3.136, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('MODERN STYLE', { x: 5.335, y: 5.05, w: 1.461, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 11),
    { x: 5.335, y: 5.353, w: 3.136, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'blobTop', { x: 11.214, y: -0.212, w: 2.326, h: 1.932, rotate: 180, flipV: true, fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
  shape(s, 'leaf', { x: 9.231, y: 5.046, w: 1.274, h: 2.816, rotate: 323.83, fill: { color: C.accent }, line: NO_LINE });
  photo(s, 0.583, 0.862, 3.938, 6.078);
});

/* ---  4. Today agenda --- */
SLIDES.push(function slide04(s, pres) {
  photo(s, 7.989, 4.485, 2.175, 2.175);
  photo(s, 5.533, 4.485, 2.175, 2.175);
  shape(s, 'blobHill', { x: 3.689, y: 6.375, w: 1.654, h: 1.125, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  s.addText('TODAY AGENDA', { x: 5.566, y: 1.105, w: 3.333, h: 0.572, fontFace: HEAD_FONT, fontSize: 28, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 41),
    { x: 5.566, y: 1.735, w: 5.818, h: 1.279, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addShape('rect', { x: 0, y: 0, w: 4.323, h: 7.5, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  s.addText('ABOUT US', { x: 1.107, y: 1.392, w: 1.045, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 8), { x: 1.107, y: 1.695, w: 2.109, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('01', { x: 0.712, y: 1.392, w: 0.395, h: 0.303, fontSize: 12, bold: true, color: C.ink, align: 'right', valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('MEET OUR TEAM', { x: 1.107, y: 2.857, w: 1.54, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 8), { x: 1.107, y: 3.16, w: 2.109, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('02', { x: 0.707, y: 2.857, w: 0.4, h: 0.303, fontSize: 12, bold: true, color: C.ink, align: 'right', valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('COLLECTION GALLERY', { x: 1.107, y: 4.321, w: 2.011, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 8), { x: 1.107, y: 4.624, w: 2.109, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('03', { x: 0.707, y: 4.321, w: 0.4, h: 0.303, fontSize: 12, bold: true, color: C.ink, align: 'right', valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('INFOGRAPHIC DATA', { x: 1.107, y: 5.786, w: 1.876, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 8), { x: 1.107, y: 6.089, w: 2.109, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('04', { x: 0.705, y: 5.786, w: 0.402, h: 0.303, fontSize: 12, bold: true, color: C.ink, align: 'right', valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  searchBar(s);
  shape(s, 'blobRight', { x: 10.979, y: 0, w: 2.354, h: 5.303, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'leaf', { x: 12.306, y: 4.831, w: 1.274, h: 2.816, rotate: 314.92, fill: { color: C.accent }, line: NO_LINE });
  pills(s, 5.654, 3.409);
});

/* ---  5. About Pratomina (section 01) --- */
SLIDES.push(function slide05(s, pres) {
  photo(s, 8.486, 0.583, 4.264, 6.333);
  photo(s, 4.583, 4.336, 2.218, 2.221);
  photo(s, 2.07, 4.336, 2.218, 2.221);
  shape(s, 'blobHill', { x: 0, y: 6.03, w: 2.204, h: 1.47, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'torn', { x: 7.003, y: 1.174, w: 2.968, h: 1.684, rotate: 354.36, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  shape(s, 'tape', { x: 7.627, y: 0.996, w: 1.405, h: 0.44, rotate: 352.64, fill: { color: C.accent, transparency: 20 }, line: NO_LINE });
  s.addText('ABOUT\nPRATOMINA', { x: 0.779, y: 1.238, w: 2.593, h: 1.043, fontFace: HEAD_FONT, fontSize: 28, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 39),
    { x: 0.779, y: 2.336, w: 5.57, h: 1.279, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('About Us', { x: 7.333, y: 1.594, w: 2.307, h: 0.774, rotate: 353.4, fontFace: SCRIPT_FONT, fontSize: 40, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('01', { x: 9.439, y: 2.262, w: 0.461, h: 0.438, rotate: 353.4, fontFace: SCRIPT_FONT, fontSize: 20, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'leaf', { x: 0.315, y: 4.646, w: 1.274, h: 2.816, rotate: 323.83, fill: { color: C.accent }, line: NO_LINE });
});

/* ---  6. What you need to know about us --- */
SLIDES.push(function slide06(s, pres) {
  shape(s, 'leaf', { x: 12.306, y: 1.308, w: 1.274, h: 2.816, rotate: 314.92, fill: { color: C.accent }, line: NO_LINE });
  photo(s, 8.625, 3.75, 4.702, 3.75);
  photo(s, 2.804, 0.953, 2.606, 2.606);
  photo(s, 0.006, 0.953, 2.6, 2.606);
  shape(s, 'blobTop', { x: 0, y: 5.353, w: 3.062, h: 2.13, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  s.addText('WHAT YOU NEED\nTO KNOW ABOUT US', { x: 6.191, y: 1.181, w: 3.601, h: 0.909, fontFace: HEAD_FONT, fontSize: 24, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(ASUS,
    { x: 6.311, y: 2.256, w: 5.543, h: 0.577, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.accent, lineSpacingMultiple: 1.5, valign: 'top', margin: 0 });
  s.addText(words(LOREM, 31),
    { x: 4.859, y: 4.433, w: 2.904, h: 2.188, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('DESCRIPTION', { x: 0.939, y: 4.469, w: 1.319, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 11),
    { x: 0.939, y: 4.772, w: 3.136, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('DESCRIPTION', { x: 0.939, y: 5.759, w: 1.319, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 11),
    { x: 0.939, y: 6.062, w: 3.136, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  dot(s, 0.753, 4.559);
  dot(s, 0.753, 5.843);
  shape(s, 'blobTop', { x: 11.208, y: -0.085, w: 2.221, h: 1.21, rotate: 180, flipV: true, fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
});

/* ---  7. The origins - timeline --- */
SLIDES.push(function slide07(s, pres) {
  shape(s, 'blobRight', { x: 10.979, y: 0, w: 2.354, h: 5.303, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'leaf', { x: 7.858, y: 1.084, w: 1.01, h: 2.233, rotate: 314.92, fill: { color: C.accent }, line: NO_LINE });
  shape(s, 'blobCorner', { x: -0.128, y: 4.884, w: 2.423, h: 2.762, fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
  shape(s, 'blobCorner', { x: 0.003, y: 5.045, w: 2.101, h: 2.565, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  s.addText('THE ORIGINS', { x: 0.779, y: 1.238, w: 2.647, h: 0.572, fontFace: HEAD_FONT, fontSize: 28, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 39),
    { x: 0.779, y: 1.878, w: 4.634, h: 1.582, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addShape('rect', { x: 4.007, y: 4.299, w: 0.862, h: 0.862, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  s.addText('APRIL\n2020', { x: 4.025, y: 4.411, w: 0.826, h: 0.64, fontFace: HEAD_FONT, fontSize: 16, bold: true, color: C.accent, align: 'center', valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('DESCRIPTION', { x: 5.107, y: 4.266, w: 1.319, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 11),
    { x: 5.107, y: 4.569, w: 3.136, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addShape('rect', { x: 4.007, y: 5.807, w: 0.862, h: 0.862, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  s.addText('APRIL\n2020', { x: 4.025, y: 5.919, w: 0.826, h: 0.64, fontFace: HEAD_FONT, fontSize: 16, bold: true, color: C.accent, align: 'center', valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('DESCRIPTION', { x: 5.107, y: 5.774, w: 1.319, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 11),
    { x: 5.107, y: 6.077, w: 3.136, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'tornTall', { x: 9.668, y: 4.569, w: 3.082, h: 2.245, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  shape(s, 'tape', { x: 8.964, y: 4.349, w: 1.405, h: 0.44, rotate: 323.74, fill: { color: C.accent, transparency: 20 }, line: NO_LINE });
  s.addText(words(ASUS, 14),
    { x: 9.983, y: 4.96, w: 2.452, h: 1.183, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.accent, align: 'center', lineSpacingMultiple: 1.5, valign: 'top', margin: 0 });
  photo(s, 0.583, 4.074, 2.843, 2.843);
  photo(s, 8.766, 0, 3.984, 3.801);
});

/* ---  8. All about our vision --- */
SLIDES.push(function slide08(s, pres) {
  shape(s, 'quarterB', { x: 0, y: 5.899, w: 2.229, h: 1.601, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'quarterA', { x: -0.188, y: 5.696, w: 2.726, h: 1.929, fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
  shape(s, 'blobRight', { x: 10.979, y: 0, w: 2.354, h: 5.303, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'leaf', { x: 12.368, y: 4.463, w: 1.274, h: 2.816, rotate: 323.83, fill: { color: C.accent }, line: NO_LINE });
  s.addText('ALL ABOUT\nOUR VISION', { x: 4.465, y: 1.707, w: 2.164, h: 0.909, fontFace: HEAD_FONT, fontSize: 24, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 37),
    { x: 4.465, y: 2.728, w: 3.503, h: 1.885, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(ASUS,
    { x: 4.558, y: 4.995, w: 2.782, h: 1.183, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.accent, lineSpacingMultiple: 1.5, valign: 'top', margin: 0 });
  s.addShape('rect', { x: 9.107, y: 1.851, w: 3.063, h: 4.455, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  s.addText('VISION ONE', { x: 9.565, y: 2.333, w: 1.168, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 8), { x: 9.565, y: 2.635, w: 2.147, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('VISION TWO', { x: 9.565, y: 3.68, w: 1.227, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 8), { x: 9.565, y: 3.983, w: 2.147, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('VISION THREE', { x: 9.565, y: 5.028, w: 1.334, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 8), { x: 9.565, y: 5.331, w: 2.147, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'tape', { x: 9.937, y: 1.579, w: 1.405, h: 0.44, fill: { color: C.accent, transparency: 20 }, line: NO_LINE });
  photo(s, 0.006, 0.968, 3.497, 5.948);
});

/* ---  9. All about our mission --- */
SLIDES.push(function slide09(s, pres) {
  shape(s, 'blobTop', { x: 10.296, y: 0, w: 3.038, h: 2.304, rotate: 180, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'blobTop', { x: 10.391, y: -0.085, w: 3.038, h: 2.159, rotate: 180, flipV: true, fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
  s.addText('ALL ABOUT\nOUR MISSION', { x: 0.989, y: 1.084, w: 2.434, h: 0.909, fontFace: HEAD_FONT, fontSize: 24, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 37),
    { x: 0.989, y: 2.104, w: 5.106, h: 1.279, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'blobRight', { x: 3.409, y: 2.576, w: 1.514, h: 8.333, rotate: 270, flipH: true, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'leaf', { x: 1.724, y: 4.858, w: 1.274, h: 2.816, rotate: 34.62, flipH: true, fill: { color: C.accent }, line: NO_LINE });
  s.addText('MISSION ONE', { x: 7.83, y: 4.266, w: 1.303, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, align: 'right', valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 11),
    { x: 5.996, y: 4.569, w: 3.136, h: 0.625, fontSize: 11, color: C.body, align: 'right', lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('MISSION TWO', { x: 7.77, y: 5.698, w: 1.362, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, align: 'right', valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 11),
    { x: 5.996, y: 6.001, w: 3.136, h: 0.625, fontSize: 11, color: C.body, align: 'right', lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('The Mission', { x: -0.12, y: 4.978, w: 2.716, h: 0.774, rotate: 90, fontFace: SCRIPT_FONT, fontSize: 40, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  photo(s, 7.639, 0.643, 2.304, 2.304);
  photo(s, 10.207, 1.795, 3.121, 5.705);
});

/* --- 10. What we can do? --- */
SLIDES.push(function slide10(s, pres) {
  shape(s, 'quarterA', { x: -0.188, y: 6.745, w: 2.266, h: 0.88, fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
  shape(s, 'blobRight', { x: 7.375, y: 1.542, w: 3.583, h: 8.333, rotate: 90, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  s.addText('WHAT WE CAN DO ?', { x: 0.696, y: 1.3, w: 3.521, h: 0.505, fontFace: HEAD_FONT, fontSize: 24, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 38),
    { x: 0.696, y: 1.878, w: 5.284, h: 1.279, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'leaf', { x: -0.275, y: 4.858, w: 1.274, h: 2.816, rotate: 34.62, flipH: true, fill: { color: C.accent }, line: NO_LINE });
  s.addText(ASUS,
    { x: 2.063, y: 5.007, w: 3.771, h: 0.88, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.accent, lineSpacingMultiple: 1.5, valign: 'top', margin: 0 });
  pills(s, 0.838, 3.682);
  shape(s, 'blobWide', { x: 2.419, y: 0, w: 5.531, h: 0.766, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  photo(s, 8.979, 0.601, 3.771, 6.316);
  s.addShape('rect', { x: 6.667, y: 3.35, w: 3.063, h: 3.039, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  s.addText(words(LOREM, 8), { x: 7.214, y: 3.667, w: 2.147, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  dot(s, 7.028, 3.801);
  s.addText(words(LOREM, 8), { x: 7.214, y: 4.557, w: 2.147, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  dot(s, 7.028, 4.691);
  s.addText(words(LOREM, 8), { x: 7.214, y: 5.447, w: 2.147, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  dot(s, 7.028, 5.581);
});

/* --- 11. Some offer from us --- */
SLIDES.push(function slide11(s, pres) {
  photo(s, 4.111, 0.853, 2.352, 2.352);
  shape(s, 'quarterB', { x: 0, y: 5.899, w: 2.229, h: 1.601, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'quarterA', { x: -0.188, y: 5.696, w: 2.726, h: 1.929, fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
  s.addText('SOME OFFER\nFROM US', { x: 7.061, y: 1.028, w: 2.271, h: 0.909, fontFace: HEAD_FONT, fontSize: 24, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 27),
    { x: 7.061, y: 2.05, w: 5.284, h: 0.976, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('DESCRIPTION', { x: 4.715, y: 4.089, w: 1.319, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 11),
    { x: 4.715, y: 4.392, w: 3.136, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  dot(s, 4.53, 4.179);
  s.addText('DESCRIPTION', { x: 4.715, y: 5.462, w: 1.319, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 11),
    { x: 4.715, y: 5.765, w: 3.136, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  dot(s, 4.53, 5.551);
  shape(s, 'blobRight', { x: 10.979, y: 2.197, w: 2.354, h: 5.303, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'leaf', { x: 12.368, y: 0.219, w: 1.274, h: 2.816, rotate: 323.83, fill: { color: C.accent }, line: NO_LINE });
  s.addText(ASUS,
    { x: 9.161, y: 4.114, w: 2.782, h: 1.183, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.accent, lineSpacingMultiple: 1.5, valign: 'top', margin: 0 });
  photo(s, 0.583, 1.125, 3.321, 5.792);
});

/* --- 12. Get the newest design --- */
SLIDES.push(function slide12(s, pres) {
  shape(s, 'blobTop', { x: 9.438, y: 0, w: 3.896, h: 4.302, rotate: 180, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  s.addText('GET THE\nNEWEST DESIGN', { x: 0.779, y: 1.238, w: 2.893, h: 0.909, fontFace: HEAD_FONT, fontSize: 24, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 39),
    { x: 0.779, y: 2.274, w: 4.634, h: 1.582, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'blobHill', { x: 0, y: 6.03, w: 2.204, h: 1.47, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  s.addText('Pratomina\nCollection', { x: 3.124, y: 4.814, w: 2.512, h: 1.447, fontFace: SCRIPT_FONT, fontSize: 40, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(ASUS, 14),
    { x: 10.634, y: 4.643, w: 1.718, h: 1.789, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.accent, lineSpacingMultiple: 1.5, valign: 'top', margin: 0 });
  shape(s, 'blobCorner', { x: 12.352, y: 6.389, w: 1.07, h: 1.225, flipH: true, fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
  shape(s, 'leaf', { x: -0.18, y: 4.858, w: 1.274, h: 2.816, rotate: 34.62, flipH: true, fill: { color: C.accent }, line: NO_LINE });
  photo(s, 6.323, 0.601, 3.709, 6.316);
  photo(s, 10.236, 1.236, 2.514, 2.514);
});

/* --- 13. Our special offer - price card --- */
SLIDES.push(function slide13(s, pres) {
  shape(s, 'blobTop', { x: 7.637, y: -0.085, w: 2.221, h: 1.21, rotate: 180, flipV: true, fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
  photo(s, 9.416, 0, 3.912, 7.5);
  shape(s, 'blobRight', { x: 3.07, y: 2.237, w: 2.193, h: 8.333, rotate: 270, flipH: true, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  s.addShape('rect', { x: 5.174, y: 4.008, w: 5.9, h: 2.668, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  s.addText('CLASSIC DESIGN', { x: 8.081, y: 4.337, w: 1.575, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 8), { x: 8.081, y: 4.64, w: 2.109, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('$ 29,00', { x: 8.081, y: 5.307, w: 1.147, h: 0.438, fontSize: 20, bold: true, color: C.amber, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 8), { x: 8.081, y: 5.69, w: 2.109, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('OUR SPECIAL\nOFFER', { x: 6.568, y: 1.332, w: 2.36, h: 0.909, fontFace: HEAD_FONT, fontSize: 24, bold: true, color: C.ink, align: 'right', valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(LOREM_SI,
    { x: 0.779, y: 1.424, w: 2.427, h: 3.097, fontSize: 12, color: C.body, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'leaf', { x: 0.003, y: 4.858, w: 1.274, h: 2.816, rotate: 34.62, flipH: true, fill: { color: C.accent }, line: NO_LINE });
  s.addText('Collection Today', { x: 5.61, y: 2.404, w: 3.37, h: 0.774, fontFace: SCRIPT_FONT, fontSize: 40, color: C.ink, align: 'right', valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  photo(s, 5.289, 4.123, 2.437, 2.438);
});

/* --- 14. Get your best design --- */
SLIDES.push(function slide14(s, pres) {
  shape(s, 'blobHill', { x: 3.527, y: 6.317, w: 2.204, h: 1.183, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'leaf', { x: 3.655, y: 5.7, w: 0.947, h: 2.093, rotate: 47.28, flipH: true, fill: { color: C.accent }, line: NO_LINE });
  shape(s, 'blobTop', { x: 9.438, y: 0, w: 3.896, h: 3.604, rotate: 180, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'blobTop', { x: 11.214, y: -0.212, w: 2.326, h: 1.932, rotate: 180, flipV: true, fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
  s.addText('GET YOUR BEST DESIGN', { x: 4.434, y: 1.549, w: 4.053, h: 0.505, fontFace: HEAD_FONT, fontSize: 24, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(LOREM_SI,
    { x: 4.434, y: 2.472, w: 2.427, h: 3.097, fontSize: 12, color: C.body, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('DESCRIPTION', { x: 7.729, y: 2.583, w: 1.319, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 8), { x: 7.729, y: 2.886, w: 2.245, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('DESCRIPTION', { x: 10.523, y: 2.583, w: 1.319, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 8), { x: 10.523, y: 2.886, w: 2.245, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  photo(s, 0.006, 0.984, 3.661, 6.516);
  photo(s, 7.603, 4.154, 2.497, 2.497);
  photo(s, 10.397, 4.154, 2.497, 2.497);
});

/* --- 15. Meet our team (section 02) --- */
SLIDES.push(function slide15(s, pres) {
  photo(s, 6.19, 0.601, 3.56, 6.316);
  shape(s, 'blobTop', { x: 10.296, y: 0, w: 3.038, h: 2.304, rotate: 180, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'blobTop', { x: 10.391, y: -0.085, w: 3.038, h: 2.159, rotate: 180, flipV: true, fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
  shape(s, 'torn', { x: 8.33, y: 4.224, w: 2.968, h: 1.684, rotate: 354.36, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  shape(s, 'tape', { x: 8.954, y: 4.045, w: 1.405, h: 0.44, rotate: 352.64, fill: { color: C.accent, transparency: 20 }, line: NO_LINE });
  s.addText('Our Team', { x: 8.655, y: 4.644, w: 2.316, h: 0.774, rotate: 353.4, fontFace: SCRIPT_FONT, fontSize: 40, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('02', { x: 10.724, y: 5.312, w: 0.546, h: 0.438, rotate: 353.4, fontFace: SCRIPT_FONT, fontSize: 20, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('MEET OUR TEAM', { x: 0.779, y: 1.238, w: 3.319, h: 0.572, fontFace: HEAD_FONT, fontSize: 28, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 39),
    { x: 0.779, y: 1.972, w: 4.634, h: 1.582, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  pills(s, 0.88, 4.036);
  shape(s, 'blobRight', { x: 3.318, y: 2.484, w: 1.698, h: 8.333, rotate: 270, flipH: true, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'leaf', { x: 0.003, y: 4.858, w: 1.274, h: 2.816, rotate: 34.62, flipH: true, fill: { color: C.accent }, line: NO_LINE });
  photo(s, 10.065, 0.601, 2.685, 2.685);
});

/* --- 16. Meet our creative team --- */
SLIDES.push(function slide16(s, pres) {
  photo(s, 10.417, 1.125, 2.333, 2.333);
  photo(s, 7.749, 1.125, 2.333, 2.333);
  photo(s, 0.605, 3.876, 3.064, 3.064);
  shape(s, 'leaf', { x: 8.401, y: 4.858, w: 1.274, h: 2.816, rotate: 34.62, flipH: true, fill: { color: C.accent }, line: NO_LINE });
  shape(s, 'blobRight', { x: 0, y: 0, w: 3.064, h: 4.578, flipH: true, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'blobTop', { x: 10.417, y: 4.313, w: 2.917, h: 3.221, rotate: 180, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  s.addText('MEET OUR\nCREATIVE TEAM', { x: 0.779, y: 1.238, w: 2.844, h: 0.909, fontFace: HEAD_FONT, fontSize: 24, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 38),
    { x: 0.779, y: 2.184, w: 5.396, h: 1.279, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('FAST RESPOND', { x: 4.713, y: 4.275, w: 1.448, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 11),
    { x: 4.713, y: 4.578, w: 3.136, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  dot(s, 4.527, 4.365);
  s.addText('CREATIVE', { x: 4.713, y: 5.507, w: 1.019, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 11),
    { x: 4.713, y: 5.809, w: 3.136, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  dot(s, 4.527, 5.596);
  shape(s, 'tornTall', { x: 9.668, y: 4.261, w: 3.082, h: 2.245, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  shape(s, 'tape', { x: 8.964, y: 4.041, w: 1.405, h: 0.44, rotate: 323.74, fill: { color: C.accent, transparency: 20 }, line: NO_LINE });
  s.addText(words(ASUS, 14),
    { x: 9.983, y: 4.651, w: 2.452, h: 1.183, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.accent, align: 'center', lineSpacingMultiple: 1.5, valign: 'top', margin: 0 });
  searchBar(s);
});

/* --- 17. Our best employee - photo grid --- */
SLIDES.push(function slide17(s, pres) {
  shape(s, 'seed', { x: 4.657, y: 0.053, w: 2.162, h: 2.204, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  photo(s, 6.299, 4.674, 2.243, 2.243);
  photo(s, 6.299, 2.224, 2.243, 2.243);
  photo(s, 3.878, 3.453, 2.243, 2.243);
  photo(s, 3.878, 1.012, 2.243, 2.243);
  shape(s, 'blobRight', { x: 10.979, y: 0, w: 2.354, h: 5.303, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'blobWave', { x: 10.128, y: -0.113, w: 3.372, h: 3.147, fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
  shape(s, 'leaf', { x: 12.368, y: 4.571, w: 1.274, h: 2.816, rotate: 323.83, fill: { color: C.accent }, line: NO_LINE });
  s.addText('OUR BEST\nEMPLOYEE', { x: 0.628, y: 1.393, w: 1.92, h: 0.909, fontFace: HEAD_FONT, fontSize: 24, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(LOREM_SI,
    { x: 0.628, y: 2.416, w: 2.427, h: 3.097, fontSize: 12, color: C.body, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addShape('rect', { x: 9.312, y: 0.841, w: 3.063, h: 6.165, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  shape(s, 'tape', { x: 10.141, y: 0.641, w: 1.405, h: 0.44, fill: { color: C.accent, transparency: 20 }, line: NO_LINE });
  s.addText('VISION ONE', { x: 9.77, y: 1.465, w: 1.168, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 8), { x: 9.77, y: 1.768, w: 2.147, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('VISION TWO', { x: 9.77, y: 2.848, w: 1.227, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 8), { x: 9.77, y: 3.151, w: 2.147, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('VISION THREE', { x: 9.77, y: 4.23, w: 1.334, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 8), { x: 9.77, y: 4.533, w: 2.147, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('VISION THREE', { x: 9.77, y: 5.613, w: 1.334, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 8), { x: 9.77, y: 5.916, w: 2.147, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'quarterC', { x: 0, y: 6.161, w: 2.224, h: 1.339, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
});

/* --- 18. Our designer - contact card --- */
SLIDES.push(function slide18(s, pres) {
  photo(s, 0.583, 1.151, 2.599, 2.599);
  shape(s, 'blobRight', { x: 10.979, y: 0, w: 2.354, h: 5.303, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'leaf', { x: 4.487, y: 5.006, w: 1.274, h: 2.816, rotate: 323.83, fill: { color: C.accent }, line: NO_LINE });
  s.addShape('rect', { x: 5.626, y: 4.41, w: 3.739, h: 1.938, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  s.addText('OUR DESIGNER', { x: 3.969, y: 1.457, w: 3.044, h: 0.572, fontFace: HEAD_FONT, fontSize: 28, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 32),
    { x: 3.969, y: 2.115, w: 4.672, h: 1.279, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('DESCRIPTION', { x: 1.016, y: 5.191, w: 1.319, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 11),
    { x: 1.016, y: 5.493, w: 3.136, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'icoPhone', { x: 6.128, y: 5.098, w: 0.211, h: 0.211, fill: { color: C.amber }, line: NO_LINE });
  s.addText('564 Street name ,city name ,\nstatename/country name 1458', { x: 6.483, y: 5.383, w: 2.325, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'icoInsta', { x: 6.128, y: 4.726, w: 0.211, h: 0.211, fill: { color: C.amber }, line: NO_LINE });
  shape(s, 'icoPin', { x: 6.128, y: 5.463, w: 0.211, h: 0.211, fill: { color: C.amber }, line: NO_LINE });
  s.addText('@mariajovanka', { x: 6.483, y: 5.058, w: 1.266, h: 0.286, fontSize: 11, color: C.body, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('+7 4562 7895 1238', { x: 6.483, y: 4.695, w: 1.522, h: 0.286, fontSize: 11, color: C.body, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'icoBank', { x: 0.616, y: 5.232, w: 0.263, h: 0.264, fill: { color: C.accent }, line: NO_LINE });
  s.addText('Maria Jovanka', { x: 0.492, y: 4.314, w: 3.457, h: 0.774, fontFace: SCRIPT_FONT, fontSize: 40, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'blobHill', { x: 0.006, y: 6.598, w: 1.84, h: 0.902, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  photo(s, 9.364, 0.8, 3.405, 6.139);
});

/* --- 19. It is time to break --- */
SLIDES.push(function slide19(s, pres) {
  photo(s, 6.5, 1.167, 5.359, 5.772);
  shape(s, 'blobRight', { x: 10.269, y: 0, w: 3.064, h: 4.578, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'blobWave', { x: 9.94, y: -0.113, w: 3.56, h: 4.676, fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
  s.addText('IT’S TIME\nTO BREAK', { x: 0.852, y: 1.167, w: 2.115, h: 1.043, fontFace: HEAD_FONT, fontSize: 28, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 32),
    { x: 0.852, y: 2.347, w: 4.672, h: 1.279, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'torn', { x: 8.925, y: 0.6, w: 4.072, h: 1.595, rotate: 12.57, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  s.addText('Break Time', { x: 9.265, y: 0.918, w: 3.422, h: 1.01, rotate: 12.57, fontFace: SCRIPT_FONT, fontSize: 54, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'tape', { x: 10.423, y: 0.401, w: 1.405, h: 0.44, rotate: 12.57, fill: { color: C.accent, transparency: 20 }, line: NO_LINE });
  s.addText('11:00-13:00', { x: 2.387, y: 5.191, w: 1.128, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 11),
    { x: 2.387, y: 5.493, w: 3.136, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'blobHill', { x: 0, y: 6.03, w: 2.649, h: 1.47, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'leaf', { x: -0.18, y: 4.858, w: 1.274, h: 2.816, rotate: 34.62, flipH: true, fill: { color: C.accent }, line: NO_LINE });
  s.addShape('ellipse', { x: 2.495, y: 4.698, w: 0.349, h: 0.349, fill: { type: 'none' }, line: { color: 'B06A57', width: 2 } });
});

/* --- 20. Collection gallery (section 03) --- */
SLIDES.push(function slide20(s, pres) {
  photo(s, 7.595, 4.397, 2.267, 2.267);
  shape(s, 'blobRight', { x: 3.318, y: 2.484, w: 1.698, h: 8.333, rotate: 270, flipH: true, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  photo(s, 0.908, 0.985, 3.877, 5.955);
  shape(s, 'leaf', { x: 12.323, y: 4.861, w: 1.274, h: 2.816, rotate: 325.38, fill: { color: C.accent }, line: NO_LINE });
  shape(s, 'blobTop2', { x: 10.296, y: 0, w: 3.037, h: 2.589, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'torn', { x: 3.334, y: 4.399, w: 2.968, h: 1.684, rotate: 354.36, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  shape(s, 'tape', { x: 3.959, y: 4.22, w: 1.405, h: 0.44, rotate: 352.64, fill: { color: C.accent, transparency: 20 }, line: NO_LINE });
  s.addText('Collection', { x: 3.806, y: 4.819, w: 2.025, h: 0.774, rotate: 353.4, fontFace: SCRIPT_FONT, fontSize: 40, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('03', { x: 5.719, y: 5.487, w: 0.565, h: 0.438, rotate: 353.4, fontFace: SCRIPT_FONT, fontSize: 20, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('COLLECTION GALLERY', { x: 5.676, y: 1.237, w: 4.562, h: 0.572, fontFace: HEAD_FONT, fontSize: 28, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 38),
    { x: 5.676, y: 1.964, w: 5.396, h: 1.279, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  photo(s, 10.127, 4.397, 2.267, 2.267);
});

/* --- 21. Pratomina collection gallery --- */
SLIDES.push(function slide21(s, pres) {
  photo(s, 4.415, 0.699, 4.264, 5.255);
  shape(s, 'blobRight', { x: 3.318, y: 2.484, w: 1.698, h: 8.333, rotate: 270, flipH: true, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  s.addText('PRATOMINA', { x: 0.668, y: 1.214, w: 2.593, h: 0.572, fontFace: HEAD_FONT, fontSize: 28, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 31),
    { x: 9.221, y: 4.076, w: 3.363, h: 1.885, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'leaf', { x: 0.003, y: 4.858, w: 1.274, h: 2.816, rotate: 34.62, flipH: true, fill: { color: C.accent }, line: NO_LINE });
  s.addText('Collection Gallery', { x: 0.668, y: 1.803, w: 3.329, h: 0.707, fontFace: SCRIPT_FONT, fontSize: 36, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'blobTop', { x: 10.296, y: 0, w: 3.038, h: 2.304, rotate: 180, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'blobTop', { x: 10.391, y: -0.085, w: 3.038, h: 2.159, rotate: 180, flipV: true, fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
  s.addShape('rect', { x: 7.633, y: 1.088, w: 2.356, h: 2.356, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  photo(s, 1.137, 3.771, 3.03, 3.168);
  photo(s, 7.71, 1.165, 2.202, 2.202);
});

/* --- 22. Best sales on this week --- */
SLIDES.push(function slide22(s, pres) {
  photo(s, 6.661, 3.962, 2.387, 2.387);
  photo(s, 6.661, 1.346, 2.387, 2.387);
  shape(s, 'blobRight', { x: 9.67, y: 3.659, w: 3.064, h: 4.578, rotate: 90, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'blobWave', { x: 9.341, y: 3.546, w: 3.56, h: 4.676, rotate: 90, fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
  s.addShape('rect', { x: 0, y: 4.424, w: 6.375, h: 1.726, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  s.addText('BEST SALES\nON THIS WEEK', { x: 0.801, y: 1.151, w: 2.567, h: 0.909, fontFace: HEAD_FONT, fontSize: 24, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 38),
    { x: 0.801, y: 2.128, w: 5.396, h: 1.279, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('125 SALES', { x: 3.085, y: 4.824, w: 1.008, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 11),
    { x: 3.085, y: 5.127, w: 3.136, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('$30', { x: 1.097, y: 5.052, w: 1.438, h: 0.909, fontFace: SCRIPT_FONT, fontSize: 48, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('PRICE', { x: 1.239, y: 4.779, w: 0.66, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, color: C.accent, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'leaf', { x: 0.162, y: 4.6, w: 0.683, h: 1.51, rotate: 34.62, flipH: true, fill: { color: C.accent }, line: NO_LINE });
  photo(s, 9.261, 0, 3.508, 6.939);
});

/* --- 23. Best product --- */
SLIDES.push(function slide23(s, pres) {
  photo(s, 5.142, 3.865, 2.574, 2.574);
  shape(s, 'blobSide', { x: 10.847, y: 0, w: 2.487, h: 4.681, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'blobSide', { x: 11.013, y: -0.25, w: 2.487, h: 5.098, fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
  s.addText('BEST PRODUCT', { x: 0.655, y: 1.227, w: 2.714, h: 0.505, fontFace: HEAD_FONT, fontSize: 24, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'leaf', { x: 2.699, y: 5.031, w: 1.274, h: 2.816, rotate: 34.62, flipH: true, fill: { color: C.accent }, line: NO_LINE });
  s.addText(LOREM_SI,
    { x: 0.655, y: 1.921, w: 2.427, h: 3.097, fontSize: 12, color: C.body, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'blobHill', { x: 0, y: 6.439, w: 4.524, h: 1.061, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  photo(s, 8.333, 0, 4.436, 6.939);
  shape(s, 'tornTall', { x: 5.84, y: 1.12, w: 3.082, h: 2.245, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  shape(s, 'tape', { x: 5.136, y: 0.901, w: 1.405, h: 0.44, rotate: 323.74, fill: { color: C.accent, transparency: 20 }, line: NO_LINE });
  s.addText(words(ASUS, 14),
    { x: 6.155, y: 1.511, w: 2.452, h: 1.183, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.accent, align: 'center', lineSpacingMultiple: 1.5, valign: 'top', margin: 0 });
});

/* --- 24. Device mockup - laptop --- */
SLIDES.push(function slide24(s, pres) {
  shape(s, 'blobSide', { x: 10.847, y: 0, w: 2.487, h: 4.681, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'blobSide', { x: 11.013, y: -0.25, w: 2.487, h: 5.098, fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
  s.addText('DEVICE MOCKUP', { x: 0.76, y: 1.95, w: 3.463, h: 0.572, fontFace: HEAD_FONT, fontSize: 28, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 38),
    { x: 0.76, y: 2.677, w: 5.396, h: 1.279, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'leaf', { x: 12.496, y: 0.256, w: 1.274, h: 2.816, rotate: 323.83, fill: { color: C.accent }, line: NO_LINE });
  s.addText(ASUS,
    { x: 7.554, y: 5.726, w: 3.756, h: 0.88, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.accent, align: 'center', lineSpacingMultiple: 1.5, valign: 'top', margin: 0 });
  pills(s, 0.864, 4.313);
  shape(s, 'blobRight', { x: 1.823, y: 4.783, w: 0.894, h: 4.54, rotate: 270, flipH: true, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  photo(s, 7.326, 1.481, 4.212, 3.155);
  laptopMockup(s, { x: 6.777, y: 1.334, w: 5.313, h: 3.816 });
});

/* --- 25. Newest design - phones --- */
SLIDES.push(function slide25(s, pres) {
  shape(s, 'blobRight', { x: 10.269, y: 0, w: 3.064, h: 4.578, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'blobRight', { x: 1.823, y: 4.783, w: 0.894, h: 4.54, rotate: 270, flipH: true, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'leaf', { x: 7.45, y: 4.814, w: 1.274, h: 2.816, rotate: 323.83, fill: { color: C.accent }, line: NO_LINE });
  shape(s, 'seed', { x: 4.999, y: 0.344, w: 2.231, h: 2.274, fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
  s.addShape('rect', { x: 8.439, y: 0.714, w: 4.1, h: 6.165, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  shape(s, 'tape', { x: 9.786, y: 0.514, w: 1.405, h: 0.44, fill: { color: C.accent, transparency: 20 }, line: NO_LINE });
  s.addText('NEWEST DESIGN', { x: 9.202, y: 1.506, w: 2.221, h: 0.404, fontFace: HEAD_FONT, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(LOREM_SI,
    { x: 9.202, y: 1.978, w: 2.427, h: 3.097, fontSize: 12, color: C.body, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('$30', { x: 9.14, y: 5.704, w: 1.438, h: 0.909, fontFace: SCRIPT_FONT, fontSize: 48, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('PRICE', { x: 9.282, y: 5.43, w: 0.66, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, color: C.accent, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(ASUS, 4), { x: 10.834, y: 5.393, w: 1.133, h: 0.88, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.accent, lineSpacingMultiple: 1.5, valign: 'top', margin: 0 });
  photo(s, 3.921, 1.264, 2.447, 5.344);
  photo(s, 0.679, 1.264, 2.447, 5.344);
  phoneMockup(s, { x: 0.429, y: 1.089, w: 3.277, h: 5.824 });
  phoneMockup(s, { x: 3.671, y: 1.089, w: 3.277, h: 5.824 });
});

/* --- 26. Best collection - tablet --- */
SLIDES.push(function slide26(s, pres) {
  photo(s, 7.079, 2.845, 4.997, 3.727);
  tabletMockup(s, { x: 6.9, y: 2.659, w: 5.49, h: 4.207 });
  shape(s, 'leaf', { x: 12.353, y: 3.574, w: 1.274, h: 2.816, rotate: 323.83, fill: { color: C.accent }, line: NO_LINE });
  shape(s, 'blobRight', { x: 8.705, y: -2.55, w: 2.069, h: 7.187, rotate: 90, flipH: true, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'blobHill', { x: 0, y: 6.03, w: 2.649, h: 1.47, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  s.addText('BEST COLLECTION', { x: 0.76, y: 1.506, w: 3.675, h: 0.572, fontFace: HEAD_FONT, fontSize: 28, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 38),
    { x: 0.76, y: 2.233, w: 5.396, h: 1.279, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'tornTall', { x: 7.187, y: 0.788, w: 4.69, h: 1.507, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  shape(s, 'tape', { x: 6.525, y: 0.568, w: 1.405, h: 0.44, rotate: 323.74, fill: { color: C.accent, transparency: 20 }, line: NO_LINE });
  s.addText('Simple Design', { x: 7.608, y: 1.019, w: 3.715, h: 0.909, fontFace: SCRIPT_FONT, fontSize: 48, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('DESCRIPTION', { x: 1.113, y: 3.988, w: 1.319, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 11),
    { x: 1.113, y: 4.291, w: 3.136, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('DESCRIPTION', { x: 1.113, y: 5.278, w: 1.319, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 11),
    { x: 1.113, y: 5.581, w: 3.136, h: 0.625, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  dot(s, 0.927, 4.078);
  dot(s, 0.927, 5.362);
  shape(s, 'blobHill', { x: -0.083, y: 6.51, w: 3.136, h: 1.122, fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
});

/* --- 27. Infographic data (section 04) --- */
SLIDES.push(function slide27(s, pres) {
  shape(s, 'blobRight2', { x: 8.965, y: 3.089, w: 3, h: 5.762, rotate: 90, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  photo(s, 6.706, 0.746, 4.222, 6.153);
  shape(s, 'leaf', { x: 11.994, y: 4.713, w: 1.274, h: 2.816, rotate: 323.83, fill: { color: C.accent }, line: NO_LINE });
  shape(s, 'blobRight', { x: 1.535, y: -1.535, w: 1.151, h: 4.222, rotate: 90, flipH: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  s.addText('INFOGRAPHIC DATA', { x: 0.744, y: 1.966, w: 4.108, h: 0.572, fontFace: HEAD_FONT, fontSize: 28, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 38),
    { x: 0.744, y: 2.693, w: 5.396, h: 1.279, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  searchBar(s);
  pills(s, 0.832, 4.409);
  shape(s, 'torn', { x: 9.726, y: 1.275, w: 2.968, h: 1.684, rotate: 354.36, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  shape(s, 'tape', { x: 10.351, y: 1.097, w: 1.405, h: 0.44, rotate: 352.64, fill: { color: C.accent, transparency: 20 }, line: NO_LINE });
  s.addText('Info Data', { x: 10.057, y: 1.695, w: 2.306, h: 0.774, rotate: 353.4, fontFace: SCRIPT_FONT, fontSize: 40, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('04', { x: 12.132, y: 2.363, w: 0.523, h: 0.438, rotate: 353.4, fontFace: SCRIPT_FONT, fontSize: 20, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
});

/* --- 28. Infographic data - stock chart --- */
SLIDES.push(function slide28(s, pres) {
  shape(s, 'blobRight2', { x: 9.692, y: 3.816, w: 1.248, h: 6.06, rotate: 90, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'leaf', { x: 12.033, y: 4.814, w: 1.274, h: 2.816, rotate: 323.83, fill: { color: C.accent }, line: NO_LINE });
  s.addText('INFOGRAPHIC DATA', { x: 4.891, y: 0.562, w: 3.552, h: 0.505, fontFace: HEAD_FONT, fontSize: 24, bold: true, color: C.ink, align: 'center', valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addChart(pres.ChartType.line, [
    { name: 'Open',  labels: STOCK_DATES, values: [44, 25, 38, 50, 34] },
    { name: 'High',  labels: STOCK_DATES, values: [55, 57, 57, 58, 36] },
    { name: 'Low',   labels: STOCK_DATES, values: [11, 12, 13, 11,  5] },
    { name: 'Close', labels: STOCK_DATES, values: [25, 38, 50, 34, 18] },
  ], {
    x: 0.591, y: 1.691, w: 6.363, h: 3.263,
    chartColors: [C.ink, C.accent, C.amber, C.peach],
    lineSize: 2, lineDataSymbol: 'none',
    showLegend: true, legendPos: 'b', legendFontFace: BODY_FONT, legendFontSize: 12, legendColor: C.body,
    catAxisLabelFontFace: BODY_FONT, catAxisLabelFontSize: 12, catAxisLabelColor: C.body,
    valAxisLabelFontFace: BODY_FONT, valAxisLabelFontSize: 12, valAxisLabelColor: C.body,
    valGridLine: { style: 'solid', color: 'D6D3C7', size: 1 },
    catGridLine: { style: 'none' },
    catAxisLineShow: false, valAxisLineShow: false,
    plotArea: { fill: { color: C.bg } }, chartArea: { fill: { color: C.bg } },
  });
  s.addText('DESCRIPTION', { x: 0.602, y: 5.54, w: 1.319, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('50%', { x: 0.602, y: 5.823, w: 0.944, h: 0.572, fontSize: 28, bold: true, color: C.accent, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 38),
    { x: 7.534, y: 2.029, w: 4.531, h: 1.582, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('DESCRIPTION', { x: 7.534, y: 1.726, w: 1.319, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(ASUS,
    { x: 7.635, y: 3.892, w: 3.756, h: 0.88, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.accent, lineSpacingMultiple: 1.5, valign: 'top', margin: 0 });
  s.addText(words(LOREM, 11),
    { x: 2.008, y: 5.479, w: 2.155, h: 0.903, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('DESCRIPTION', { x: 4.602, y: 5.54, w: 1.319, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('50%', { x: 4.602, y: 5.823, w: 0.944, h: 0.572, fontSize: 28, bold: true, color: C.accent, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 11),
    { x: 6.008, y: 5.479, w: 2.155, h: 0.903, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
});

/* --- 29. Infographic data - icon tiles --- */
SLIDES.push(function slide29(s, pres) {
  shape(s, 'blobRight2', { x: 11.771, y: 0, w: 1.575, h: 5.02, flipV: true, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  s.addText('INFOGRAPHIC DATA', { x: 4.891, y: 0.562, w: 3.552, h: 0.505, fontFace: HEAD_FONT, fontSize: 24, bold: true, color: C.ink, align: 'center', valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addShape('rect', { x: 7.86, y: 1.891, w: 1.867, h: 1.867, fill: { color: C.accent }, line: NO_LINE });
  s.addShape('rect', { x: 10.045, y: 1.891, w: 1.867, h: 1.867, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  s.addShape('rect', { x: 7.86, y: 4.108, w: 1.867, h: 1.867, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  s.addShape('rect', { x: 10.045, y: 4.108, w: 1.867, h: 1.867, fill: { color: C.accent }, line: NO_LINE });
  shape(s, 'icoBox', { x: 10.602, y: 2.448, w: 0.754, h: 0.754, fill: { color: C.accent }, line: NO_LINE });
  shape(s, 'icoCoins', { x: 8.448, y: 4.665, w: 0.691, h: 0.754, fill: { color: C.accent }, line: NO_LINE });
  shape(s, 'icoAtom', { x: 8.417, y: 2.449, w: 0.752, h: 0.753, fill: { color: C.white }, line: NO_LINE });
  shape(s, 'icoUser', { x: 10.601, y: 4.665, w: 0.755, h: 0.754, fill: { color: C.white }, line: NO_LINE });
  s.addText(words(LOREM, 11),
    { x: 1.244, y: 2.367, w: 2.371, h: 0.903, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('PRODUCTION', { x: 1.25, y: 1.997, w: 1.326, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 11),
    { x: 1.244, y: 3.933, w: 2.371, h: 0.903, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('GLOBAL MARKETING', { x: 1.25, y: 3.563, w: 1.894, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 11),
    { x: 4.723, y: 2.367, w: 2.371, h: 0.903, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('PROFIT', { x: 4.728, y: 1.997, w: 0.791, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(words(LOREM, 11),
    { x: 4.723, y: 3.933, w: 2.371, h: 0.903, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText('CLIENT', { x: 4.728, y: 3.563, w: 0.777, h: 0.303, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  s.addText(ASUS,
    { x: 1.352, y: 5.211, w: 5.966, h: 0.577, fontFace: HEAD_FONT, fontSize: 12, bold: true, color: C.accent, lineSpacingMultiple: 1.5, valign: 'top', margin: 0 });
  shape(s, 'blobHill', { x: 0, y: 6.03, w: 2.649, h: 1.47, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'leaf', { x: 0.26, y: 5.131, w: 1.274, h: 2.816, rotate: 323.83, fill: { color: C.accent }, line: NO_LINE });
});

/* --- 30. Thank you --- */
SLIDES.push(function slide30(s, pres) {
  shape(s, 'quarterB', { x: 0, y: 5.899, w: 2.229, h: 1.601, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  shape(s, 'quarterA', { x: -0.188, y: 5.696, w: 2.726, h: 1.929, fill: { type: 'none' }, line: { color: C.accent, width: 1 } });
  photo(s, 0.564, 1.824, 8.267, 5.115);
  shape(s, 'blobRight', { x: 10.979, y: 2.197, w: 2.354, h: 5.303, fill: { color: C.peach, transparency: 50 }, line: NO_LINE });
  s.addText('THANK YOU', { x: 0.564, y: 1.336, w: 1.687, h: 0.404, fontFace: HEAD_FONT, bold: true, color: C.accent, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'leaf', { x: 12.368, y: 1.422, w: 1.274, h: 2.816, rotate: 323.83, fill: { color: C.accent }, line: NO_LINE });
  s.addText(words(LOREM_SI, 30),
    { x: 9.476, y: 4.128, w: 2.427, h: 2.188, fontSize: 12, color: C.body, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
  shape(s, 'torn', { x: 7.137, y: 1.434, w: 4.739, h: 1.645, fill: { color: C.white }, line: NO_LINE, shadow: CARD_SHADOW });
  shape(s, 'tape', { x: 8.683, y: 1.181, w: 1.405, h: 0.44, rotate: 358.28, fill: { color: C.accent, transparency: 20 }, line: NO_LINE });
  s.addText('Thanks To Come', { x: 7.473, y: 1.777, w: 4.139, h: 0.909, rotate: 359.04, fontFace: SCRIPT_FONT, fontSize: 48, bold: true, color: C.ink, valign: 'top', wrap: false, margin: [7.2, 7.2, 3.6, 3.6] });
});

/* ================================= build ================================== */
function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'PRATOMINA', width: 13.333, height: 7.5 });
  pres.layout = 'PRATOMINA';
  pres.author = 'Pratomina';
  pres.title = 'Pratomina Fashion Collection';

  SLIDES.forEach(function (buildSlide) {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    buildSlide(s, pres);
  });

  return pres.writeFile({ fileName: path.join(__dirname, '040501e0-5135-4dd9-86fa-9c73bc346433_grok_final.pptx') });
}

build().then(function (f) { console.log('wrote ' + f); }, function (e) { console.error(e); process.exit(1); });
