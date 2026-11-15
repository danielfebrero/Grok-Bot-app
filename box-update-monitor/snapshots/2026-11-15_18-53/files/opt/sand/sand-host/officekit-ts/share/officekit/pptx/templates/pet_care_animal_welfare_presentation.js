/**
 * WeCare - Pet Animal Presentation Template  (40 slides, 13.333 x 7.5 in)
 * Standalone pptxgenjs re-creation of the reference deck.
 *
 * Run:  node 16a2dd8b-3d57-47f5-966f-3a138f95bccb_grok_final.js
 * Out:  16a2dd8b-3d57-47f5-966f-3a138f95bccb_grok_final.pptx  (next to this file)
 *
 * Raster artwork of the original (the bowl logo and the desktop mockup photo) is
 * redrawn with native shapes; empty picture placeholders keep the tinted panels
 * the template itself draws behind them.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const C = {
  ORANGE: 'FD9340', // brand accent
  RED: 'FF3F3F', // secondary accent
  SAND: 'F3BA8D', // tertiary accent
  BLUSH: 'FFD7D7', // pale red
  PEACH: 'FFE7D7', // pale orange
  DARK: '404040', // headings          (tx1 lumMod 75%)
  MUTED: '595959', // sub headings      (tx1 lumMod 65%)
  BODY: '808080', // paragraph copy    (tx1 lumMod 50%)
  WHITE: 'FFFFFF',
  BG: 'FFFDFD', // slide background
  GREY: 'BFBFBF', // hairlines / big numerals
  SILVER: 'E9E7E7',
  // Flat stand-ins for the deck's orange -> red diagonal gradients
  GRAD_BOWL: 'FDBD9D',
  GRAD_CARD: 'FCA87A',
  GRAD_DISC: 'FCC9A9',
  GRAD_TILE: 'FC9F6D',
  GRAD_BAND: 'FC9C6C',
};

const F = {
  TITLE: 'Mulish Bold', // 40/44 pt slide titles
  DISPLAY: 'Mulish Black', // 80 pt cover words
  BODY: 'Open Sans',
  SEMI: 'Open Sans SemiBold',
  BOLD: 'Open Sans bold',
  MONO: 'Montserrat SemiBold', // corner chrome
  MB: 'Montserrat Bold', // small caps labels
  MBLACK: 'Montserrat Black', // big numerals
};

/* -------------------------------------------------------------- vector art */
/* Normalised (0..1) outlines lifted from the template's freeform shapes. */

const BLOB = // organic paint splash used in the page corners and as a backdrop
  "M 0.125,0.93 C 0.146,0.93 0.163,0.946 0.163,0.965 C 0.163,0.984 0.146,1 0.125,1 C 0.104,1 0.087,0.984 0.087,0.965 C 0.087,0.946 0.104,0.93 0.125,0.93 Z M 0.139,0.771 C 0.174,0.771 0.202,0.798 0.202,0.83 C 0.202,0.862 0.174,0.888 0.139,0.888 C 0.104,0.888 0.075,0.862 0.075,0.83 C 0.075,0.798 0.104,0.771 0.139,0.771 Z M 0.828,0.571 C 0.849,0.571 0.866,0.586 0.866,0.605 C 0.866,0.625 0.849,0.64 0.828,0.64 C 0.807,0.64 0.79,0.625 0.79,0.605 C 0.79,0.586 0.807,0.571 0.828,0.571 Z M 0.897,0.253 C 0.932,0.253 0.96,0.279 0.96,0.311 C 0.96,0.344 0.932,0.37 0.897,0.37 C 0.861,0.37 0.833,0.344 0.833,0.311 C 0.833,0.279 0.861,0.253 0.897,0.253 Z M 0.443,0.019 C 0.464,0.019 0.481,0.034 0.481,0.053 C 0.481,0.073 0.464,0.088 0.443,0.088 C 0.423,0.088 0.406,0.073 0.406,0.053 C 0.406,0.034 0.423,0.019 0.443,0.019 Z M 0.27,0.001 C 0.282,0.002 0.293,0.004 0.303,0.008 C 0.384,0.035 0.378,0.129 0.439,0.139 C 0.496,0.148 0.529,0.072 0.584,0.085 C 0.637,0.097 0.639,0.174 0.69,0.178 C 0.712,0.179 0.73,0.165 0.753,0.147 C 0.803,0.109 0.797,0.074 0.839,0.052 C 0.887,0.027 0.969,0.041 0.992,0.088 C 1.011,0.129 0.996,0.18 0.954,0.203 C 0.924,0.219 0.892,0.219 0.841,0.234 C 0.815,0.242 0.729,0.273 0.731,0.314 C 0.733,0.339 0.77,0.357 0.796,0.369 C 0.865,0.402 0.884,0.412 0.892,0.435 C 0.9,0.458 0.887,0.49 0.864,0.505 C 0.827,0.528 0.791,0.492 0.733,0.512 C 0.725,0.515 0.677,0.532 0.67,0.566 C 0.661,0.61 0.721,0.629 0.721,0.682 C 0.72,0.719 0.69,0.76 0.652,0.776 C 0.591,0.8 0.55,0.745 0.476,0.767 C 0.425,0.783 0.397,0.826 0.395,0.829 C 0.364,0.878 0.394,0.919 0.356,0.952 C 0.346,0.96 0.313,0.988 0.273,0.98 C 0.238,0.973 0.21,0.941 0.209,0.911 C 0.207,0.87 0.255,0.861 0.266,0.808 C 0.271,0.78 0.267,0.736 0.237,0.714 C 0.188,0.677 0.117,0.739 0.053,0.702 C 0.018,0.683 -0.004,0.64 0,0.604 C 0.007,0.55 0.071,0.54 0.071,0.492 C 0.072,0.453 0.02,0.431 0.006,0.376 C -0.004,0.335 0.003,0.297 0.033,0.27 C 0.063,0.242 0.108,0.249 0.138,0.214 C 0.19,0.155 0.124,0.078 0.172,0.029 C 0.195,0.006 0.234,-0.003 0.27,0.001 Z";

const PAW = // four toes + pad
  "M 0.512,0.393 C 0.551,0.396 0.589,0.414 0.616,0.445 C 0.637,0.468 0.652,0.495 0.659,0.526 C 0.669,0.563 0.687,0.597 0.712,0.626 L 0.729,0.646 C 0.743,0.662 0.759,0.676 0.776,0.689 C 0.801,0.707 0.82,0.73 0.834,0.757 C 0.855,0.798 0.86,0.845 0.848,0.89 C 0.845,0.898 0.843,0.906 0.839,0.914 C 0.805,0.987 0.714,1.02 0.637,0.987 C 0.627,0.983 0.618,0.979 0.609,0.974 C 0.58,0.955 0.545,0.945 0.51,0.945 L 0.481,0.946 C 0.448,0.948 0.415,0.958 0.387,0.975 C 0.375,0.982 0.361,0.988 0.346,0.992 C 0.262,1.016 0.174,0.97 0.15,0.89 C 0.137,0.845 0.143,0.798 0.164,0.757 C 0.177,0.73 0.197,0.707 0.221,0.689 C 0.239,0.676 0.255,0.662 0.269,0.646 L 0.285,0.626 C 0.311,0.597 0.329,0.563 0.339,0.526 C 0.346,0.496 0.361,0.468 0.382,0.445 C 0.388,0.438 0.394,0.432 0.401,0.426 C 0.433,0.401 0.473,0.39 0.512,0.393 Z M 0.119,0.299 C 0.169,0.303 0.219,0.34 0.244,0.397 C 0.277,0.472 0.252,0.553 0.189,0.578 C 0.125,0.603 0.047,0.563 0.015,0.487 C -0.018,0.412 0.006,0.331 0.07,0.306 C 0.085,0.299 0.102,0.297 0.119,0.299 Z M 0.881,0.299 C 0.898,0.297 0.915,0.299 0.93,0.306 C 0.994,0.331 1.018,0.412 0.985,0.487 C 0.953,0.563 0.875,0.603 0.811,0.578 C 0.748,0.553 0.723,0.472 0.756,0.397 C 0.781,0.34 0.831,0.303 0.881,0.299 Z M 0.669,0.001 C 0.677,0 0.686,0 0.694,0.001 C 0.762,0.011 0.803,0.098 0.787,0.195 C 0.77,0.292 0.702,0.362 0.635,0.351 C 0.567,0.341 0.525,0.254 0.542,0.157 C 0.556,0.073 0.61,0.009 0.669,0.001 Z M 0.332,0.001 C 0.39,0.009 0.444,0.073 0.459,0.157 C 0.475,0.254 0.433,0.341 0.366,0.351 C 0.298,0.361 0.23,0.292 0.214,0.195 C 0.197,0.098 0.239,0.011 0.306,0.001 C 0.315,0 0.323,0 0.332,0.001 Z";

const CHECK = // filled disc with a tick
  "M 0.5,0 C 0.776,0 1,0.224 1,0.5 C 1,0.776 0.776,1 0.5,1 C 0.224,1 0,0.776 0,0.5 C 0,0.224 0.224,0 0.5,0 Z M 0.668,0.265 C 0.656,0.268 0.645,0.275 0.638,0.286 L 0.439,0.606 L 0.359,0.502 C 0.344,0.481 0.314,0.477 0.293,0.492 C 0.273,0.508 0.269,0.538 0.284,0.558 L 0.405,0.718 C 0.414,0.729 0.428,0.736 0.442,0.736 C 0.443,0.736 0.444,0.736 0.445,0.736 C 0.46,0.736 0.474,0.727 0.482,0.714 L 0.718,0.336 C 0.732,0.314 0.725,0.285 0.703,0.271 C 0.692,0.264 0.679,0.262 0.668,0.265 Z";

const DISH = "M 0,0 L 1,0 L 0.999,0.097 C 0.993,0.604 0.94,1 0.876,1 L 0.124,1 C 0.064,1 0.014,0.655 0.002,0.197 L 0,0 Z";
const DOME = "M 0.5,0 C 0.776,0 1,0.441 1,0.985 L 0.999,1 L 0.001,1 L 0,0.985 C 0,0.441 0.224,0 0.5,0 Z";
const ARCH = "M 0.5,0 C 0.692,0 0.865,0.356 0.984,0.922 L 1,1 L 0,1 L 0.016,0.922 C 0.135,0.356 0.308,0 0.5,0 Z";
const WAVE = "M 0,1 L 0,0.181 C 0.333,-0.446 0.667,0.808 1,0.181 L 1,1 Z";
const CARD = "M 0.03,0 L 0.97,0 C 0.986,0 1,0.06 1,0.135 L 1,0.865 C 1,0.94 0.986,1 0.97,1 L 0.03,1 C 0.014,1 0,0.94 0,0.865 L 0,0.135 C 0,0.06 0.014,0 0.03,0 Z";
const RSQUARE = "M 0.268,0 L 0.732,0 C 0.88,0 1,0.125 1,0.279 L 1,0.721 C 1,0.875 0.88,1 0.732,1 L 0.268,1 C 0.12,1 0,0.875 0,0.721 L 0,0.279 C 0,0.125 0.12,0 0.268,0 Z";
const BOWL = "M 0.09,0.19 L 0.91,0.19 L 1,0.8 L 0,0.8 Z";

const CAT =
  "M 0.99,0.21 L 0.85,0.06 L 0.85,0.01 C 0.85,0 0.83,0 0.82,0 L 0.6,0.16 C 0.57,0.18 0.56,0.21 0.56,0.24 C 0.56,0.35 0.18,0.39 0.18,0.73 C 0.18,0.77 0.18,0.81 0.19,0.84 C 0.19,0.84 0.19,0.84 0.18,0.83 C 0.14,0.81 0.1,0.75 0.1,0.62 C 0.1,0.57 0.13,0.52 0.16,0.48 C 0.2,0.41 0.26,0.32 0.17,0.25 C 0.15,0.24 0.11,0.24 0.1,0.25 C 0.08,0.27 0.08,0.29 0.1,0.31 C 0.13,0.33 0.12,0.37 0.07,0.44 C 0.04,0.49 0,0.55 0,0.62 C 0,0.74 0.04,0.84 0.11,0.89 C 0.15,0.92 0.2,0.93 0.23,0.94 C 0.26,0.99 0.3,1 0.31,1 L 0.51,1 C 0.54,1 0.56,0.98 0.56,0.96 C 0.56,0.94 0.54,0.93 0.51,0.93 L 0.46,0.93 C 0.52,0.9 0.57,0.85 0.57,0.75 C 0.57,0.67 0.55,0.64 0.48,0.61 C 0.47,0.6 0.46,0.59 0.47,0.58 C 0.48,0.57 0.49,0.57 0.5,0.57 C 0.6,0.61 0.62,0.67 0.62,0.75 C 0.62,0.8 0.61,0.84 0.59,0.87 L 0.59,0.96 C 0.59,0.98 0.61,1 0.64,1 C 0.67,1 0.69,0.98 0.69,0.96 L 0.69,0.85 C 0.7,0.84 0.71,0.83 0.72,0.82 L 0.72,0.96 C 0.72,0.98 0.74,1 0.77,1 C 0.8,1 0.82,0.98 0.82,0.96 L 0.82,0.63 C 0.83,0.6 0.84,0.56 0.84,0.54 C 0.84,0.44 0.8,0.43 0.8,0.38 C 0.8,0.35 0.83,0.33 0.87,0.33 C 0.94,0.33 1,0.29 1,0.23 C 1,0.22 1,0.21 0.99,0.21 Z";

const HOUSE =
  "M 0.5,0.66 C 0.45,0.66 0.41,0.7 0.41,0.75 L 0.41,0.95 L 0.59,0.95 L 0.59,0.75 C 0.59,0.7 0.55,0.66 0.5,0.66 Z M 0.41,0.44 C 0.39,0.44 0.38,0.45 0.38,0.46 C 0.38,0.47 0.38,0.48 0.39,0.49 C 0.38,0.51 0.38,0.53 0.39,0.54 C 0.39,0.54 0.4,0.55 0.41,0.55 C 0.42,0.55 0.43,0.53 0.44,0.52 L 0.43,0.51 L 0.57,0.51 L 0.56,0.52 C 0.57,0.53 0.58,0.55 0.59,0.55 C 0.6,0.55 0.62,0.53 0.62,0.52 C 0.62,0.51 0.62,0.5 0.61,0.49 C 0.62,0.48 0.62,0.46 0.61,0.45 C 0.6,0.44 0.59,0.44 0.58,0.44 C 0.57,0.45 0.56,0.46 0.56,0.47 L 0.44,0.47 C 0.44,0.46 0.43,0.45 0.42,0.44 C 0.42,0.44 0.41,0.44 0.41,0.44 Z M 0.5,0.19 L 0.81,0.48 L 0.81,0.95 L 1,0.95 L 1,1 L 0,1 L 0,0.95 L 0.19,0.95 L 0.19,0.48 Z M 0.5,0 L 0.97,0.44 L 0.91,0.49 L 0.5,0.11 L 0.09,0.5 L 0.03,0.45 Z";

const DOG =
  "M 0.72,0.75 L 0.74,0.75 C 0.77,0.77 0.8,0.78 0.83,0.79 L 0.83,0.8 C 0.83,0.82 0.81,0.84 0.79,0.84 L 0.72,0.84 C 0.69,0.84 0.67,0.82 0.67,0.8 C 0.67,0.77 0.69,0.75 0.72,0.75 Z M 0.35,0.45 L 0.25,0.63 L 0.28,0.67 L 0.39,0.47 C 0.39,0.47 0.38,0.47 0.38,0.46 Z M 0.82,0 C 0.92,0 1,0.09 1,0.2 C 1,0.26 0.98,0.31 0.95,0.35 C 0.96,0.4 0.96,0.45 0.94,0.5 L 0.92,0.54 C 0.91,0.56 0.9,0.6 0.89,0.63 L 0.89,0.66 L 0.94,0.71 C 0.95,0.72 0.95,0.73 0.95,0.74 L 0.95,0.89 C 0.95,0.92 0.94,0.94 0.91,0.94 L 0.84,0.94 C 0.82,0.94 0.8,0.92 0.8,0.89 C 0.8,0.86 0.82,0.84 0.84,0.84 L 0.87,0.84 L 0.87,0.76 C 0.81,0.76 0.75,0.73 0.71,0.68 L 0.68,0.73 C 0.64,0.79 0.59,0.82 0.53,0.83 C 0.51,0.83 0.49,0.85 0.48,0.87 L 0.47,0.9 C 0.47,0.92 0.45,0.93 0.44,0.94 L 0.19,1 C 0.17,1 0.15,0.99 0.14,0.96 C 0.14,0.94 0.15,0.91 0.17,0.91 L 0.32,0.87 L 0.3,0.83 L 0.12,0.86 C 0.1,0.86 0.08,0.84 0.07,0.82 C 0.07,0.79 0.08,0.77 0.11,0.77 L 0.28,0.74 L 0.21,0.65 L 0.14,0.68 C 0.09,0.71 0.02,0.68 0,0.63 L 0,0.62 L 0.05,0.56 C 0.04,0.49 0.07,0.42 0.13,0.39 L 0.18,0.37 C 0.19,0.37 0.2,0.36 0.21,0.36 L 0.22,0.36 L 0.25,0.43 C 0.26,0.44 0.26,0.45 0.26,0.46 C 0.26,0.47 0.25,0.49 0.24,0.49 C 0.23,0.49 0.23,0.5 0.23,0.5 C 0.22,0.5 0.21,0.5 0.2,0.49 C 0.19,0.49 0.19,0.48 0.18,0.47 C 0.18,0.46 0.17,0.46 0.17,0.46 C 0.16,0.47 0.16,0.48 0.16,0.48 C 0.16,0.5 0.17,0.51 0.18,0.52 C 0.2,0.52 0.21,0.53 0.22,0.53 C 0.24,0.53 0.25,0.52 0.25,0.52 C 0.27,0.51 0.28,0.49 0.28,0.46 C 0.29,0.45 0.28,0.43 0.28,0.42 L 0.26,0.37 L 0.27,0.38 L 0.4,0.42 C 0.44,0.44 0.48,0.43 0.51,0.39 L 0.61,0.28 C 0.69,0.2 0.81,0.2 0.9,0.27 C 0.93,0.23 0.92,0.16 0.88,0.12 C 0.87,0.1 0.84,0.09 0.82,0.09 C 0.8,0.09 0.78,0.07 0.78,0.05 C 0.78,0.02 0.8,0 0.82,0 Z";

/* -------------------------------------------------------------- boilerplate */

const LOREM = {
  full: 'lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo libero vivamus porttitor dolor, conubia mollit. sapien nam suspendisse, tincidunt eget ante tincidunt, eros in auctor fringilla',
  long: 'lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo libero vivamus porttitor dolor, conubia mollit. sapien nam suspendisse, tincidunt eget ante tincidunt',
  med: 'lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo libero vivamus porttitor dolor, conubia mollit',
  medNulla: 'lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo libero vivamus porttitor dolor, conubia mollit lacus nulla',
  half: 'lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo libero vivamus',
  card: 'PLACEHOLDER',
  tiny: 'lorem ipsum dolor sit amet, lacus nulla ac netus',
  brief: 'lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula',
};

const SED = {
  full: 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam ',
  totam: 'Sed ut perspiciatis unde omnis iste natus error totam',
  error: 'Sed ut perspiciatis unde omnis iste natus error',
  omnis: 'Sed ut perspiciatis unde omnis iste',
};

const IPSUM = 'Lorem ipsum dolor sit amet consectetur adipiscing elit';
const IPSUM_TEMPOR = IPSUM + ', sed eiusmod tempor';
const VOLUPTATE = 'PLACEHOLDER';
const VOLUPTATE_S = 'PLACEHOLDER';

/* ------------------------------------------------------------------ helpers */

/** Brand orange at `pct` % opacity. */
const tint = (pct, color) => ({ color: color || C.ORANGE, transparency: 100 - pct });

/** Turn a normalised outline into pptxgenjs custGeom points sized to w x h. */
function outline(d, w, h) {
  const tok = d.split(' ');
  const pts = [];
  let i = 0;
  const next = () => {
    const p = tok[i++].split(',');
    return { x: +p[0] * w, y: +p[1] * h };
  };
  while (i < tok.length) {
    const cmd = tok[i++];
    if (cmd === 'M') pts.push(Object.assign(next(), { moveTo: true }));
    else if (cmd === 'L') pts.push(next());
    else if (cmd === 'Z') pts.push({ close: true });
    else if (cmd === 'C') {
      const a = next(), b = next(), p = next();
      pts.push({ x: p.x, y: p.y, curve: { type: 'cubic', x1: a.x, y1: a.y, x2: b.x, y2: b.y } });
    }
  }
  return pts;
}

/** Freeform shape from one of the outlines above. */
function art(s, d, x, y, w, h, fill, opt) {
  s.addShape('custGeom', Object.assign({ x, y, w, h, fill, points: outline(d, w, h) }, opt));
}

const rect = (s, x, y, w, h, fill, opt) => s.addShape('rect', Object.assign({ x, y, w, h, fill }, opt));
const oval = (s, x, y, w, h, fill, opt) => s.addShape('ellipse', Object.assign({ x, y, w, h, fill }, opt));

/** Rounded rectangle; `adj` is the template's corner ratio (fraction of the short side). */
function round(s, x, y, w, h, fill, adj, opt) {
  s.addShape('roundRect', Object.assign(
    { x, y, w, h, fill, rectRadius: (adj === undefined ? 0.16667 : adj) * Math.min(w, h) }, opt));
}

/** Hollow circle (ring) drawn as disc + background-coloured hole. */
function ring(s, x, y, d, pct, thickness) {
  const t = thickness === undefined ? 0.18056 : thickness;
  oval(s, x, y, d, d, tint(pct));
  oval(s, x + d * t, y + d * t, d * (1 - 2 * t), d * (1 - 2 * t), { color: C.BG });
}

/** Halftone disc: grid of small dots clipped to a circle. */
function dotDisc(s, x, y, d, pct) {
  const n = 24, step = d / n, dot = step * 0.45;
  for (let r = 0; r <= n; r++) {
    for (let c = 0; c <= n; c++) {
      const u = c / n - 0.5, v = r / n - 0.5;
      if (u * u + v * v > 0.25) continue;
      oval(s, x + c * step, y + r * step, dot, dot, tint(pct));
    }
  }
}

/** Diamond patch of diagonal hatch strokes. */
function hatch(s, x, y, w, h, pct) {
  const step = 0.185, len = 0.125;
  for (let r = 0; r * step <= h; r++) {
    for (let c = 0; c * step <= w; c++) {
      const u = (c * step) / w - 0.5, v = (r * step) / h - 0.5;
      if (Math.abs(u) + Math.abs(v) > 0.5) continue;
      s.addShape('line', {
        x: x + c * step, y: y + r * step, w: len, h: len, flipV: true,
        line: { color: C.ORANGE, transparency: 100 - pct, width: 1.5 },
      });
    }
  }
}

/** Body copy: 12 pt Open Sans, 1.5 line spacing, grey. */
function body(s, text, x, y, w, h, opt) {
  s.addText(text, Object.assign({
    x, y, w, h, valign: 'top', fontFace: F.BODY, fontSize: 12, color: C.BODY, lineSpacingMultiple: 1.5,
  }, opt));
}

/** Small bold caption above a paragraph. */
function label(s, text, x, y, w, opt) {
  s.addText(text, Object.assign({
    x, y, w, h: 0.34, valign: 'top', fontFace: F.BOLD, fontSize: 14, color: C.MUTED,
  }, opt));
}

/** Letter-spaced upper-case tag. */
function tag(s, text, x, y, w, opt) {
  s.addText(text, Object.assign({
    x, y, w, h: 0.3, valign: 'top', fontFace: F.BOLD, fontSize: 12, color: C.MUTED, charSpacing: 2,
  }, opt));
}

/** Two-tone slide title: dark words then orange words. */
function title(s, dark, accent, x, y, w, h, opt) {
  s.addText([
    { text: dark, options: { color: C.DARK } },
    { text: accent, options: { color: C.ORANGE } },
  ], Object.assign({ x, y, w, h, valign: 'top', fontFace: F.TITLE, fontSize: 40 }, opt));
}

/** The WeCare dog-bowl mark (redrawn; the template ships it as a raster/svg logo). */
function logo(s) {
  const x = 0.372, y = 0.327, w = 0.459, h = 0.234;
  const at = (u, v) => [x + u * w, y + v * h];
  oval(s, x, y + 0.56 * h, w, 0.44 * h, { color: C.ORANGE }); // flared foot
  art(s, BOWL, x, y, w, h, { color: C.ORANGE }); // tapered body
  oval(s, x + 0.09 * w, y, 0.82 * w, 0.38 * h, { color: C.ORANGE }); // rim
  oval(s, x + 0.13 * w, y + 0.06 * h, 0.74 * w, 0.25 * h, { color: C.BG }); // rim opening
  const bone = { lobe: 0.05, bx: at(0.36, 0.64)[0], by: at(0, 0.64)[1] };
  rect(s, bone.bx, bone.by - 0.016, 0.28 * w, 0.032, { color: C.BG });
  [0.36, 0.64].forEach((u) => {
    [0.53, 0.75].forEach((v) => {
      const [cx, cy] = at(u, v);
      oval(s, cx - bone.lobe / 2, cy - bone.lobe / 2, bone.lobe, bone.lobe, { color: C.BG });
    });
  });
}

/** Corner blobs, logo and the two running heads that every content slide shares. */
function chrome(s) {
  art(s, BLOB, -0.941, 6.617, 2.59, 2.816, tint(54));
  art(s, BLOB, 12.306, 6.632, 2.59, 2.816, tint(54));
  logo(s);
  s.addText('WeCare', {
    x: 0.842, y: 0.28, w: 1.329, h: 0.337, valign: 'top', fontFace: F.MONO, fontSize: 14, color: C.BODY,
  });
  s.addText('AnimalCare', {
    x: 11.211, y: 0.28, w: 1.761, h: 0.337, valign: 'top', align: 'right',
    fontFace: F.MONO, fontSize: 14, color: C.BODY,
  });
}

/* ============================================================ slide builders */

/** 1 - cover */
function slide01(s) {
  oval(s, 6.286, 1.251, 3.479, 3.479, tint(20));
  oval(s, 6.647, 1.612, 2.758, 2.758, tint(24));
  oval(s, 7.116, 2.078, 1.824, 1.824, tint(34));
  oval(s, 6.276, 4.567, 6.736, 1.783, { color: C.ORANGE });
  art(s, PAW, 2.404, 3.038, 0.623, 0.61, { color: C.ORANGE });
  oval(s, 6.276, 4.458, 6.736, 1.783, { color: C.SILVER });
  s.addText([
    { text: 'We', options: { color: C.DARK } },
    { text: 'Care', options: { color: C.ORANGE } },
  ], { x: 1.027, y: 3.225, w: 5.358, h: 1.447, valign: 'top', fontFace: F.DISPLAY, fontSize: 80 });
  s.addText('PET’S ANIMAL PRESENTATION TEMPLATE', {
    x: 1.071, y: 4.566, w: 4.499, h: 0.337, valign: 'top',
    fontFace: F.SEMI, fontSize: 14, color: C.BODY, charSpacing: 1,
  });
  chrome(s);
}

/** 2 - welcome */
function slide02(s) {
  art(s, BLOB, 9.901, 2.942, 2.59, 2.816, tint(6));
  rect(s, 1.948, 4.057, 3.281, 2.326, tint(12));
  body(s, LOREM.full, 6.391, 3.784, 4.995, 1.28);
  s.addText([
    { text: 'Welcome to our ', options: { color: C.DARK } },
    { text: 'family', options: { color: C.ORANGE } },
  ], { x: 6.391, y: 1.599, w: 5.693, h: 1.582, valign: 'top', fontFace: F.TITLE, fontSize: 44 });
  round(s, 6.544, 5.922, 1.775, 0.473, { color: C.ORANGE }, 0.5);
  s.addText('WELCOME', {
    x: 6.574, y: 6.007, w: 1.715, h: 0.303, valign: 'top', align: 'center',
    fontFace: F.TITLE, fontSize: 12, color: C.WHITE, charSpacing: 2,
  });
  art(s, PAW, 4.196, 1.001, 1.221, 1.195, tint(16));
  chrome(s);
}

/** 3 - about (dog-bowl band) */
function slide03(s) {
  art(s, PAW, 8.764, 0.453, 1.778, 1.74, tint(8));
  s.addText('About WeCare Family', {
    x: 1.279, y: 5.852, w: 2.181, h: 0.303, valign: 'top', fontFace: F.BOLD, fontSize: 12, color: C.MUTED,
  });
  s.addText([
    { text: 'About WeCare ', options: { color: C.DARK } },
    { text: 'family', options: { color: C.ORANGE } },
  ], { x: 1.279, y: 1.139, w: 8.374, h: 0.841, valign: 'top', fontFace: F.TITLE, fontSize: 44 });
  body(s, LOREM.full, 3.46, 5.795, 8.583, 0.675);
  art(s, DISH, 1.279, 3.847, 10.763, 1.332, { color: C.GRAD_BOWL });
  chrome(s);
}

/** 4 - socialize */
function slide04(s) {
  rect(s, 0, 2.198, 10.291, 1.357, tint(10));
  s.addShape('teardrop', { x: 8.481, y: 3.706, w: 2.926, h: 2.926, fill: { color: C.ORANGE } });
  s.addShape('teardrop', { x: 10.437, y: 0.985, w: 2.57, h: 2.57, fill: { color: C.ORANGE }, rotate: 270, flipH: true });
  title(s, 'Socialize your pet with other ', 'animals', 1.242, 1.436, 6.336, 1.447);
  label(s, 'WeCare text 01', 1.242, 4.081, 2.59);
  body(s, LOREM.half, 1.242, 4.575, 3.13, 0.978);
  label(s, 'WeCare text 02', 4.778, 4.081, 2.59);
  body(s, LOREM.half, 4.778, 4.575, 3.13, 0.978);
  rect(s, 11.57, 3.706, 1.764, 1.425, tint(10));
  chrome(s);
}

/** 5 - training */
function slide05(s) {
  rect(s, 0, 0, 3.792, 7.5, tint(6));
  title(s, 'Train your pet to maintain ', 'health', 7.191, 1.629, 4.465, 2.121);
  body(s, LOREM.full, 7.285, 4.122, 3.926, 1.583);
  art(s, DOG, 10.38, 3.262, 1.712, 1.5, tint(8));
  chrome(s);
}

/** 6 - grooming */
function slide06(s) {
  art(s, BLOB, 9.074, 1.363, 3.513, 3.819, tint(13));
  title(s, "Brush your pet's coat ", 'regularly', 1.254, 2.019, 4.92, 1.447);
  body(s, LOREM.full, 1.254, 4.631, 4.92, 1.28);
  s.addText('Wecare Grooming', {
    x: 1.242, y: 4.198, w: 2.59, h: 0.337, valign: 'top', fontFace: F.SEMI, fontSize: 14, color: C.MUTED,
  });
  chrome(s);
}

/** 7 - behaviour */
function slide07(s) {
  rect(s, 10.366, 0, 2.968, 7.5, tint(7));
  dotDisc(s, 2.764, 3.936, 2.144, 60);
  ring(s, 6.463, 3.111, 2.286, 22);
  title(s, 'Every pet is ', 'unique', 5.943, 1.473, 6.362, 0.774, { align: 'right' });
  body(s, LOREM.long, 1.08, 2.087, 5.382, 0.978);
  s.addText('WeCare Behaviour', {
    x: 1.08, y: 1.543, w: 2.59, h: 0.337, valign: 'top', fontFace: F.SEMI, fontSize: 14, color: C.MUTED,
  });
  chrome(s);
}

/** 8 - nutrition */
function slide08(s) {
  hatch(s, 2.154, 4.606, 3.439, 2.293, 35);
  title(s, 'WeCare provide balanced ', 'nutrition', 1.649, 1.353, 6.509, 1.447);
  body(s, LOREM.long, 1.649, 3.456, 5.605, 0.978);
  const nutrients = [
    ['Vitamins and Minerals:', 1.65, 5.273, 2.59, true],
    ['Protein', 1.669, 5.777, 2.525, false],
    ['Healthy Fats', 4.769, 5.273, 2.816, false],
    ['Carbohydrates', 4.789, 5.777, 2.157, false],
  ];
  nutrients.forEach(([text, x, y, w, bold]) => {
    s.addText(text, {
      x, y, w, h: 0.303, valign: 'top', fontFace: F.BOLD, fontSize: 12, color: C.MUTED, bold,
    });
  });
  art(s, PAW, 8.158, 0.747, 1.238, 1.212, tint(8));
  ring(s, 10.541, 1.791, 2.286, 10);
  chrome(s);
}

/** 9 - care with love (two gradient cards) */
function slide09(s) {
  round(s, 8.118, 3.727, 4.137, 2.806, tint(15), 0.08966);
  [1.078, 4.598].forEach((x) => round(s, x, 4.306, 2.974, 2.227, { color: C.GRAD_CARD }, 0.08966));
  const cards = [
    { x: 1.078, head: 'Attention', hx: 1.711, hw: 1.715, bx: 1.365 },
    { x: 4.598, head: 'Physical emotion', hx: 4.985, hw: 2.207, bx: 4.885 },
  ];
  cards.forEach((c) => {
    s.addText(c.head, {
      x: c.hx, y: 4.72, w: c.hw, h: 0.37, valign: 'top', align: 'center',
      fontFace: F.SEMI, fontSize: 16, color: C.WHITE,
    });
    body(s, 'lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula',
      c.bx, 5.141, 2.408, 0.978, { align: 'center', color: C.WHITE });
  });
  title(s, 'We look after and care for your pets with ', 'love', 1.078, 1.212, 7.411, 1.447);
  body(s, LOREM.half + ' porttitor dolor', 1.078, 3.061, 4.92, 0.675);
  round(s, 8.118, 3.75, 4.137, 2.432, tint(15), 0.08966);
  round(s, 8.118, 3.75, 4.137, 2.019, tint(15), 0.08966);
  art(s, PAW, 9.805, 1.318, 2.034, 1.99, tint(8));
  chrome(s);
}

/** 10 - vision (everything flush right) */
function slide10(s) {
  ring(s, 0.397, 0.807, 2.286, 10);
  title(s, 'WeCare ', 'vision', 5.607, 1.492, 6.509, 0.774, { align: 'right' });
  body(s, LOREM.long, 7.81, 2.817, 4.306, 1.28, { align: 'right' });
  body(s, LOREM.med, 8.127, 5.687, 3.989, 0.978, { align: 'right' });
  label(s, 'Your text here', 9.501, 5.096, 2.59, { align: 'right' });
  body(s, LOREM.med + ' libero vivamus porttitor dolor, conubia', 1.218, 5.687, 5.449, 0.978, { align: 'right' });
  label(s, 'Your text here', 4.052, 5.096, 2.59, { align: 'right' });
  chrome(s);
}

/** 11 - mission */
function slide11(s) {
  rect(s, 9.354, 1.107, 3.979, 5.285, tint(7));
  title(s, 'WeCare ', 'mission', 1.0, 1.492, 6.509, 0.774);
  oval(s, 6.746, 1.107, 5.285, 5.285, tint(20));
  oval(s, 7.13, 1.492, 4.517, 4.517, { color: C.GRAD_DISC });
  body(s, LOREM.med, 1.0, 3.75, 3.989, 0.978);
  label(s, 'Your text here', 1.0, 2.907, 2.59);
  body(s, LOREM.long, 1.0, 4.909, 4.955, 0.978);
  chrome(s);
}

/** 12 - pet collection */
function slide12(s) {
  ring(s, 11.076, 1.333, 3.403, 10);
  body(s, LOREM.med, 5.744, 2.009, 6.59, 0.675);
  title(s, 'WeCare pet ', 'collection', 1.0, 1.269, 5.667, 1.447);
  s.addText('Our Collection', {
    x: 5.744, y: 1.526, w: 2.59, h: 0.37, valign: 'top', fontFace: F.BOLD, fontSize: 16, color: C.MUTED,
  });
  art(s, ARCH, 2.601, 0, 4.065, 0.945, tint(11), { rotate: 180 });
  [[1.0, 2.653], [4.965, 2.41], [8.937, 2.41]].forEach(([x, w], i) => {
    s.addText('Your text here', {
      x: [1.0, 4.965, 8.937][i], y: 5.884, w: 2.59, h: 0.337, valign: 'top',
      fontFace: F.SEMI, fontSize: 14, color: C.MUTED,
    });
    body(s, 'lorem ipsum dolor sit amet, lacus nulla ac netus ', [1.0, 4.959, 8.931][i], 6.222, w, 0.675);
  });
  chrome(s);
}

/** 13 - environment */
function slide13(s) {
  rect(s, 12.306, 0, 1.028, 7.5, tint(15));
  rect(s, 10.743, 0, 2.59, 7.5, tint(7));
  title(s, 'Environment is crucial for the well-being of your ', 'pets', 1.398, 1.229, 9.081, 1.447);
  body(s, LOREM.med, 1.398, 3.75, 4.118, 0.978);
  body(s, LOREM.long, 1.398, 5.103, 4.118, 1.28);
  rect(s, 11.55, 0, 1.783, 7.5, tint(15));
  chrome(s);
}

/** 14 - pets are important */
function slide14(s) {
  hatch(s, 10.138, -0.457, 3.439, 2.293, 45);
  art(s, CARD, 5.792, 5.131, 6.335, 1.429, tint(20));
  art(s, CARD, 5.792, 5.32, 6.335, 1.429, tint(20));
  title(s, 'Pets are becoming more and more important ', 'to us', 1.476, 1.036, 10.651, 1.447);
  body(s, LOREM.med + ' lacus nulla ac netus nibh aliquet, porttitor ligula justo libero',
    5.792, 3.672, 6.335, 0.978);
  s.addText('Your text here', {
    x: 5.792, y: 3.115, w: 2.59, h: 0.37, valign: 'top', fontFace: F.BOLD, fontSize: 16, color: C.MUTED,
  });
  chrome(s);
}

/** 15 - common household pets (two stat cards) */
function slide15(s) {
  [0.82, 8.71].forEach((x) => round(s, x, 5.046, 3.707, 1.764, { color: C.WHITE }, 0.11155,
    { shadow: { type: 'outer', color: '000000', opacity: 0.08, blur: 12, offset: 2, angle: 90 } }));
  s.addText([
    { text: 'The most common household ', options: { color: C.DARK } },
    { text: 'pets', options: { color: C.ORANGE } },
  ], { x: 1.341, y: 0.9, w: 10.651, h: 0.774, valign: 'top', align: 'center', fontFace: F.TITLE, fontSize: 40 });
  body(s, LOREM.medNulla, 0.799, 2.762, 2.59, 1.583, { align: 'right' });
  body(s, LOREM.medNulla, 9.944, 2.762, 2.59, 1.583);
  const stats = [
    { num: '105+', cap: 'WECARE DOGS', nx: 1.952, ny: 5.417, cx: 0.879, cy: 6.188, align: 'right' },
    { num: '150+', cap: 'WECARE BIRDS', nx: 9.864, ny: 5.432, cx: 9.864, cy: 6.204, align: 'left' },
  ];
  stats.forEach((st) => {
    s.addText(st.num, {
      x: st.nx, y: st.ny, w: 1.517, h: 0.707, valign: 'top', align: st.align,
      fontFace: F.MBLACK, fontSize: 36, color: C.GREY,
    });
    s.addText(st.cap, {
      x: st.cx, y: st.cy, w: 2.59, h: 0.303, valign: 'top', align: st.align,
      fontFace: F.MB, fontSize: 12, color: C.ORANGE,
    });
  });
  chrome(s);
}

/** 16 - services (split background) */
function slide16(s) {
  rect(s, 0, 0, 6.465, 7.5, tint(3));
  ring(s, 8.352, 0.464, 2.625, 20);
  title(s, 'WeCare ', 'services', 1.117, 1.389, 4.7, 0.774);
  body(s, LOREM.medNulla, 1.117, 2.64, 4.48, 0.978);
  [[1.117], [7.111]].forEach(([x]) => {
    body(s, LOREM.med + ' lacus nulla ac netus nibh aliquet', x, 5.18, 5.105, 0.978);
    label(s, 'Your text here', x, 4.623, 2.59);
  });
  chrome(s);
}

/** 17 - three services with round icons */
function slide17(s) {
  const services = [
    { x: 5.284, tag: 'GROOMING', tx: 4.858, bx: 4.538, icon: CAT, ix: 5.67, iy: 3.656, iw: 0.414, ih: 0.547,
      copy: 'PLACEHOLDER' },
    { x: 8.031, tag: 'DAYCARE', tx: 7.606, bx: 7.286, icon: HOUSE, ix: 8.374, iy: 3.678, iw: 0.502, ih: 0.502,
      copy: LOREM.card },
    { x: 10.779, tag: 'TRAINING', tx: 10.353, bx: 10.033, icon: DOG, ix: 11.111, iy: 3.7, iw: 0.521, ih: 0.456,
      copy: LOREM.card },
  ];
  services.forEach((sv) => {
    oval(s, sv.x, 3.335, 1.188, 1.188, tint(20));
    oval(s, sv.x + 0.148, 3.484, 0.89, 0.89, tint(20));
    art(s, DOME, sv.x + 0.148, 3.486, 0.89, 0.452, tint(40));
  });
  rect(s, 0, 5.288, 13.333, 2.212, tint(6));
  title(s, 'We provide care for your ', 'pets', 0.948, 1.097, 8.399, 0.774);
  body(s, LOREM.medNulla, 0.948, 2.029, 8.774, 0.675);
  services.forEach((sv) => {
    tag(s, sv.tag, sv.tx, 4.786, 2.038, { align: 'center' });
    body(s, sv.copy, sv.bx, 5.501, 2.678, 0.978, { align: 'center' });
    art(s, sv.icon, sv.ix, sv.iy, sv.iw, sv.ih, { color: C.ORANGE });
  });
  chrome(s);
}

/** 18 - services with love (three photo cards) */
function slide18(s) {
  title(s, 'WeCare services', ' with love', 2.874, 0.929, 7.586, 0.774, { align: 'center' });
  body(s, LOREM.medNulla + ' netus nibh aliquet, porttitor ligula justo libero vivamus porttitor dolor',
    2.061, 1.85, 9.211, 0.675, { align: 'center' });
  [0.983, 5.226, 9.468].forEach((x) => {
    round(s, x, 2.96, 2.882, 2.585, tint(15), 0.08966);
    round(s, x, 3.207, 2.882, 2.338, tint(15), 0.08966);
  });
  [1.129, 5.371, 9.614].forEach((x) => {
    body(s, LOREM.tiny, x, 6.139, 2.59, 0.675, { align: 'center' });
    label(s, 'Your text here', x, 5.774, 2.59, { align: 'center' });
  });
  chrome(s);
}

/** 19 - portfolio with big circle */
function slide19(s) {
  title(s, 'About WeCare ', 'portfolio', 1.187, 1.527, 4.574, 1.447);
  oval(s, 7.82, 1.693, 4.08, 4.08, { color: C.ORANGE });
  body(s, LOREM.medNulla, 1.187, 3.395, 5.626, 0.675);
  [['PORTFOLIO 01', 1.187], ['PORTFOLIO 02', 4.643]].forEach(([txt, x]) => {
    tag(s, txt, x, 4.977, 2.038);
    body(s, LOREM.card, x, 5.46, 2.991, 0.675);
  });
  oval(s, 8.659, 4.211, 2.406, 2.406, tint(27));
  chrome(s);
}

/** 20 - portfolio 01/02/03 tiles */
function slide20(s) {
  rect(s, 0, 0, 13.333, 4.695, tint(5));
  title(s, 'About WeCare ', 'portfolio', 3.287, 1.439, 6.759, 0.774, { align: 'center' });
  const tiles = [
    { x: 0.828, w: 2.858, num: '01', fill: tint(80),
      copy: 'lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo', cw: 2.191 },
    { x: 4.096, w: 5.142, num: '02', fill: { color: C.GRAD_TILE },
      copy: 'lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo libero vivamus porttitor dolor. nulla ac netus nibh aliquet, porttitor ligula justo libero ', cw: 4.265 },
    { x: 9.648, w: 2.858, num: '03', fill: tint(80),
      copy: 'lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo', cw: 2.191 },
  ];
  tiles.forEach((t) => {
    round(s, t.x, 3.868, t.w, 2.509, t.fill, 0.14137);
    body(s, t.copy, t.x + 0.365, 5.139, t.cw, 0.831, { fontSize: 10, color: C.WHITE, align: 'justify' });
    s.addText(t.num, {
      x: t.x + 0.398, y: 4.055, w: 1.423, h: 0.64, valign: 'top',
      fontFace: F.TITLE, fontSize: 32, color: C.WHITE,
    });
    tag(s, 'PORTFOLIO ' + t.num, t.x + 0.398, 4.785, 2.038, { color: C.WHITE });
  });
  body(s, LOREM.medNulla + ' dolor sit amet, lacus nulla ac netus nibh aliquet',
    2.122, 2.54, 9.089, 0.675, { align: 'center' });
  chrome(s);
}

/** 21 - portfolio with orange band */
function slide21(s) {
  rect(s, 0, 4.391, 7.635, 0.595, { color: C.ORANGE });
  title(s, 'About WeCare ', 'portfolio', 4.835, 1.14, 5.483, 1.447);
  body(s, LOREM.medNulla, 8.422, 5.211, 4.235, 0.978);
  body(s, LOREM.medNulla, 8.422, 3.031, 4.235, 0.978);
  s.addText([
    { text: 'WECARE ', options: { color: C.MUTED } },
    { text: 'PORTFOLIO', options: { color: C.ORANGE } },
  ], { x: 8.422, y: 4.612, w: 2.644, h: 0.303, valign: 'top', fontFace: F.BOLD, fontSize: 12, charSpacing: 2 });
  chrome(s);
}

/** 22 - break slide */
function slide22(s) {
  art(s, PAW, 4.417, 1.548, 4.499, 4.403, tint(8));
  s.addText([
    { text: 'Break', options: { color: C.DARK } },
    { text: 'Slide', options: { color: C.ORANGE } },
  ], { x: 2.972, y: 2.83, w: 7.39, h: 1.447, valign: 'top', align: 'center', fontFace: F.DISPLAY, fontSize: 80 });
  s.addText('LET’S TAKE A FIVE MINUTES BREAK', {
    x: 4.417, y: 4.334, w: 4.499, h: 0.337, valign: 'top', align: 'center',
    fontFace: F.SEMI, fontSize: 14, color: C.BODY, charSpacing: 1,
  });
  chrome(s);
}

/** 23 - team members */
function slide23(s) {
  title(s, 'WeCare team ', 'member', 3.016, 1.209, 7.302, 0.774, { align: 'center' });
  const team = [
    { x: 1.06, name: 'Dennish Wayn', role: 'Marketing', no: '#01' },
    { x: 5.151, name: 'David McCall', role: 'Designer', no: '#02' },
    { x: 9.242, name: 'Kevin Dean', role: 'Sales', no: '#03' },
  ];
  team.forEach((m) => {
    round(s, m.x, 4.111, 3.383, 2.591, { color: C.WHITE }, 0.1271,
      { shadow: { type: 'outer', color: '000000', opacity: 0.1, blur: 14, offset: 2, angle: 90 } });
    s.addText(m.name, {
      x: m.x + 0.516, y: 4.771, w: 2.591, h: 0.37, valign: 'top',
      fontFace: F.BOLD, fontSize: 16, color: C.MUTED,
    });
    body(s, IPSUM, m.x + 0.516, 5.616, 2.591, 0.675);
    rect(s, m.x + 2.629, 4.469, 0.753, 0.337, { color: C.ORANGE });
    s.addText(m.no, {
      x: m.x + 2.7, y: 4.485, w: 0.611, h: 0.286, valign: 'top', align: 'center',
      fontFace: F.MB, fontSize: 11, color: C.WHITE, charSpacing: 2,
    });
    s.addText(m.role, {
      x: m.x + 0.516, y: 5.209, w: 1.784, h: 0.337, valign: 'top',
      fontFace: F.MB, fontSize: 14, color: C.ORANGE,
    });
  });
  ring(s, 10.924, -1.201, 3.403, 10);
  chrome(s);
}

/** 24 - expert members */
function slide24(s) {
  rect(s, 0, 0, 13.333, 2.803, tint(4));
  art(s, RSQUARE, 5.418, 4.538, 2.508, 2.413, tint(23));
  title(s, 'WeCare expert ', 'member', 3.794, 1.003, 5.746, 1.447, { align: 'center' });
  const experts = [
    { x: 0.601, name: 'David McCall', role: 'Marketing', rx: 1.407, align: 'right', lx: 0.862 },
    { x: 10.141, name: 'Enco Cruif', role: 'Designer', rx: 10.141, align: 'left', lx: 10.259 },
  ];
  experts.forEach((e) => {
    s.addText(e.name, {
      x: e.x, y: 3.815, w: 2.591, h: 0.37, valign: 'top', align: e.align,
      fontFace: F.BOLD, fontSize: 16, color: C.MUTED,
    });
    body(s, IPSUM, e.x, 4.895, 2.591, 0.675, { align: e.align });
    s.addText(e.role, {
      x: e.rx, y: 4.22, w: 1.784, h: 0.337, valign: 'top', align: e.align,
      fontFace: F.MB, fontSize: 14, color: C.ORANGE,
    });
    s.addShape('line', { x: e.lx, y: 4.774, w: 2.174, h: 0, line: { color: C.GREY, width: 0.75 } });
  });
  chrome(s);
}

/** 25 - gallery */
function slide25(s) {
  title(s, 'WeCare ', 'gallery', 0.864, 1.128, 4.573, 0.774);
  round(s, 4.948, 5.241, 7.212, 1.734, { color: C.GRAD_BAND });
  body(s, LOREM.medNulla + ' dolor sit ame', 5.667, 1.175, 6.639, 0.675);
  s.addText('Where Pets Feel at Home', {
    x: 5.856, y: 5.721, w: 2.591, h: 0.774, valign: 'top', align: 'right',
    fontFace: F.BOLD, fontSize: 20, color: C.WHITE,
  });
  chrome(s);
}

/** SWOT bullet: check disc plus one line of copy. */
function swotRow(s, text, cx, cy, tx, ty, tw, align) {
  art(s, CHECK, cx, cy, 0.342, 0.342, { color: C.ORANGE });
  s.addText(text, {
    x: tx, y: ty, w: tw, h: 0.372, valign: 'top', align: align || 'left',
    fontFace: F.BOLD, fontSize: 12, color: C.BODY, lineSpacingMultiple: 1.5,
  });
}

/** 26 - strength analysis */
function slide26(s) {
  dotDisc(s, 9.698, 1.043, 2.608, 48);
  body(s, SED.full, 5.746, 2.69, 6.19, 0.675, { align: 'right' });
  title(s, 'Strength ', 'analysis', 6.597, 1.402, 5.282, 0.774, { align: 'right' });
  ring(s, -1.304, -1.304, 2.608, 10);
  s.addShape('line', { x: 6.94, y: 3.81, w: 4.938, h: 0, line: { color: C.GREY, width: 0.75 } });
  [[SED.totam, 4.352], [SED.error.replace('omnis iste', 'iste'), 5.042], [SED.error, 5.729]]
    .forEach(([t, y], i) => swotRow(s, t, 11.536, y + 0.035, 6.389, y, 4.996, 'right'));
  chrome(s);
}

/** 27 - weaknesses analysis */
function slide27(s) {
  art(s, PAW, 7.22, 1.448, 5.282, 5.169, tint(10));
  title(s, 'Weakenesses ', 'analysis', 1.169, 1.349, 5.523, 1.447);
  body(s, SED.full, 1.169, 3.21, 5.739, 0.675);
  [[SED.omnis, 4.352], [SED.error, 5.042]]
    .forEach(([t, y]) => swotRow(s, t, 1.169, y + 0.035, 1.662, y, 4.317));
  hatch(s, -0.214, 0.599, 3.439, 2.293, 35);
  chrome(s);
}

/** 28 - opportunities analysis */
function slide28(s) {
  art(s, WAVE, -1.51, 1.51, 7.5, 4.48, tint(10), { rotate: 270, flipV: true });
  body(s, SED.full, 5.942, 3.065, 6.739, 0.675);
  title(s, 'Opportunities ', 'analysis', 5.942, 1.171, 5.282, 1.447);
  [[5.942, 4.301], [5.942, 5.423], [9.456, 4.301], [9.456, 5.423]].forEach(([x, y]) => {
    swotRow(s, SED.totam, x, y, x + 0.493, y - 0.063, 2.732);
  });
  chrome(s);
}

/** 29 - threats analysis */
function slide29(s) {
  art(s, BLOB, 7.144, 0.795, 4.067, 4.421, tint(11));
  title(s, 'Threats ', 'analysis', 1.169, 1.349, 5.523, 0.774);
  body(s, SED.full, 1.169, 2.648, 6.053, 0.675);
  [[SED.omnis, 3.825], [SED.error, 4.515], [SED.omnis, 5.216], [SED.error, 5.906]]
    .forEach(([t, y]) => swotRow(s, t, 1.169, y, 1.662, y - 0.035, 4.317));
  chrome(s);
}

/** 30 - infographic: four rings on a dashed timeline */
function slide30(s) {
  title(s, 'WeCare ', 'infographic', 3.383, 1.086, 6.567, 0.774, { align: 'center' });
  const dash = { color: C.GREY, width: 1.75, dashType: 'lgDash' };
  s.addShape('line', { x: 2.275, y: 5.279, w: 8.783, h: 0, line: dash });
  s.addShape('line', { x: 11.038, y: 4.585, w: 0, h: 0.694, line: dash });
  s.addShape('line', { x: 5.193, y: 4.585, w: 0, h: 0.694, line: dash });
  const steps = [
    { n: '01', ring: 1.417, ry: 4.421, outer: C.ORANGE, inner: C.PEACH, nx: 1.987, nw: 0.575, ny: 5.027,
      cap: 1.045, cx: 0.81, cy: 2.943, by: 3.246 },
    { n: '02', ring: 4.344, ry: 2.955, outer: C.RED, inner: C.BLUSH, nx: 4.882, nw: 0.64, ny: 3.56,
      cap: 3.973, cx: 3.738, cy: 5.66, by: 5.963 },
    { n: '03', ring: 7.272, ry: 4.421, outer: C.ORANGE, inner: C.PEACH, nx: 7.809, nw: 0.642, ny: 5.027,
      cap: 6.9, cx: 6.666, cy: 2.943, by: 3.246 },
    { n: '04', ring: 10.2, ry: 2.955, outer: C.RED, inner: C.BLUSH, nx: 10.72, nw: 0.675, ny: 3.56,
      cap: 9.828, cx: 9.593, cy: 5.677, by: 5.979 },
  ];
  steps.forEach((st) => {
    tag(s, 'INFOGRAPHIC ' + st.n, st.cap, st.cy, 2.46, { align: 'center', fontFace: F.MB });
    body(s, VOLUPTATE, st.cx, st.by, 2.929, 0.978, { align: 'center' });
    oval(s, st.ring, st.ry, 1.716, 1.716, { color: st.outer });
    oval(s, st.ring + 0.342, st.ry + 0.343, 1.031, 1.031, { color: st.inner, line: { color: C.WHITE, width: 7.5 } });
    s.addText(st.n, {
      x: st.nx, y: st.ny, w: st.nw, h: 0.505, valign: 'top', align: 'center',
      fontFace: F.MBLACK, fontSize: 24, color: st.outer,
    });
  });
  oval(s, 5.106, 5.192, 0.174, 0.174, { color: C.GREY });
  oval(s, 10.951, 5.198, 0.174, 0.174, { color: C.GREY });
  chrome(s);
}

/** 31 - infographic: four teardrop petals */
function slide31(s) {
  title(s, 'WeCare ', 'infographic', 3.383, 1.086, 6.567, 0.774, { align: 'center' });
  const petals = [
    { x: 4.846, y: 2.817, fill: C.RED, rotate: 270, flipH: true, flipV: true },
    { x: 6.86, y: 2.817, fill: C.ORANGE, rotate: 90, flipV: true },
    { x: 4.846, y: 4.811, fill: C.ORANGE, rotate: 90, flipH: true },
    { x: 6.86, y: 4.811, fill: C.RED, rotate: 180, flipV: true },
  ];
  petals.forEach((p) => s.addShape('teardrop', {
    x: p.x, y: p.y, w: 1.712, h: 1.712, fill: { color: p.fill },
    rotate: p.rotate, flipH: p.flipH, flipV: p.flipV,
  }));
  const bubbles = [
    { n: '01', x: 5.216, y: 3.195, tx: 5.437, tw: 0.514, col: C.ORANGE },
    { n: '02', x: 7.23, y: 3.195, tx: 7.425, tw: 0.567, col: C.RED },
    { n: '03', x: 5.22, y: 5.176, tx: 5.414, tw: 0.568, col: C.ORANGE },
    { n: '04', x: 7.234, y: 5.176, tx: 7.414, tw: 0.596, col: C.RED },
  ];
  bubbles.forEach((b) => {
    oval(s, b.x, b.y, 0.956, 0.956, { color: C.WHITE });
    s.addText(b.n, {
      x: b.tx, y: b.y + 0.259, w: b.tw, h: 0.438, valign: 'top', align: 'center',
      fontFace: F.MBLACK, fontSize: 20, color: b.col,
    });
  });
  const notes = [
    { n: 'INFOGRAPHIC 01', tx: 2.201, tw: 2.063, bx: 0.833, y: 3.086, by: 3.578, align: 'right' },
    { n: 'INFOGRAPHIC 02', tx: 9.069, tw: 2.265, bx: 9.069, y: 3.086, by: 3.578, align: 'left' },
    { n: 'INFOGRAPHIC 03', tx: 1.935, tw: 2.33, bx: 0.833, y: 5.249, by: 5.74, align: 'right' },
    { n: 'INFOGRAPHIC 04', tx: 9.069, tw: 2.265, bx: 9.069, y: 5.249, by: 5.74, align: 'left' },
  ];
  notes.forEach((t) => {
    body(s, IPSUM_TEMPOR, t.bx, t.by, 3.431, 0.675, { align: t.align });
    tag(s, t.n, t.tx, t.y, t.tw, { align: t.align, fontFace: F.MB });
  });
  chrome(s);
}

/** Flat-screen monitor standing in for the mockup photo. */
function monitor(s, x, y, w, h) {
  rect(s, x, y + 0.01 * h, w, 0.71 * h, { color: '0A0A0A' }); // bezel
  rect(s, x + 0.042 * w, y + 0.065 * h, 0.916 * w, 0.61 * h, { color: C.WHITE }); // screen
  rect(s, x, y + 0.72 * h, w, 0.11 * h, { color: 'BCBDC1' }); // chin
  rect(s, x + 0.42 * w, y + 0.83 * h, 0.16 * w, 0.1 * h, { color: 'A8A9AD' }); // neck
  rect(s, x + 0.3 * w, y + 0.93 * h, 0.4 * w, 0.07 * h, { color: 'CACBCF' }); // foot
}

/** 32 - mockup slide */
function slide32(s) {
  dotDisc(s, 9.783, 3.878, 2.912, 48);
  ring(s, 6.453, 1.258, 2.286, 15);
  monitor(s, 7.406, 2.249, 4.788, 4.026);
  const stats = [['80', 1.348, 1.049], ['40', 4.376, 3.975]];
  stats.forEach(([num, nx, cx]) => {
    s.addText([
      { text: num, options: { color: C.GREY } },
      { text: '+', options: { color: C.ORANGE } },
    ], { x: nx, y: 5.07, w: 1.826, h: 0.64, valign: 'top', align: 'center', fontFace: F.MBLACK, fontSize: 32 });
    s.addText('YOUR TEXT HERE', {
      x: cx, y: 5.656, w: 2.228, h: 0.37, valign: 'top', align: 'center',
      fontFace: F.MB, fontSize: 12, color: C.MUTED, charSpacing: 1, lineSpacingMultiple: 1.5,
    });
  });
  body(s, SED.full + ', eaque ipsa quae ab illo inventore veritatis et quasi', 1.139, 3.602, 4.982, 0.978);
  title(s, 'WeCare mockup ', 'slide', 1.139, 1.796, 5.71, 1.447);
  chrome(s);
}

/** 33 - doughnut chart slide */
function slide33(s) {
  rect(s, 0, 0, 5.391, 7.5, tint(4));
  s.addChart('doughnut', [{
    name: 'Sales',
    labels: ['1text', '2text', '3text', '4text'],
    values: [8.2, 3.2, 1.4, 2.2],
  }], {
    x: -0.005, y: 2.978, w: 5.396, h: 3.597,
    chartColors: [C.ORANGE, C.RED, C.BLUSH, C.SAND],
    dataBorder: { pt: 1.5, color: C.WHITE },
    holeSize: 50, showLegend: true, legendPos: 'b', legendFontSize: 12, legendColor: C.MUTED,
    showValue: false, showTitle: false,
  });
  const processes = [
    { dot: C.ORANGE, dx: 6.144, dy: 2.8, tx: 6.624, ty: 2.827, by: 3.255 },
    { dot: C.SAND, dx: 9.519, dy: 2.8, tx: 9.999, ty: 2.827, by: 3.255 },
    { dot: C.RED, dx: 6.061, dy: 5.063, tx: 6.541, ty: 5.09, by: 5.517 },
    { dot: C.BLUSH, dx: 9.435, dy: 5.063, tx: 9.916, ty: 5.09, by: 5.517 },
  ];
  processes.forEach((p) => {
    body(s, VOLUPTATE_S, p.tx, p.by, 2.46, 0.978);
    tag(s, 'YOUR PROCESS', p.tx, p.ty, 2.46, { fontFace: F.MB, bold: true });
    oval(s, p.dx, p.dy, 0.286, 0.286, { color: p.dot });
  });
  title(s, 'WeCare slide ', 'chart', 0.611, 1.247, 4.169, 1.447, { align: 'center' });
  chrome(s);
}

/** 34 - clustered bar chart slide */
function slide34(s) {
  body(s, SED.full + ', eaque ipsa quae ab illo inventore veritatis et quasi architecto beatae vitae dicta sunt explicabo. Nemo',
    1.25, 1.931, 10.833, 0.675, { align: 'center' });
  const cats = ['Category 1', 'Category 2', 'Category 3', 'Category 4'];
  s.addChart('bar', [
    { name: 'Series 1', labels: cats, values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels: cats, values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels: cats, values: [2, 2, 3, 5] },
  ], {
    x: 1.35, y: 3.015, w: 10.633, h: 3.892,
    chartColors: [C.ORANGE, C.RED, C.SAND],
    barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 219, barOverlapPct: -27,
    showLegend: false, showTitle: false,
    catAxisLabelColor: C.MUTED, catAxisLabelFontSize: 12, catAxisLineColor: 'D9D9D9',
    valAxisLabelColor: C.MUTED, valAxisLabelFontSize: 12, valAxisLineShow: false,
    valGridLine: { color: 'D9D9D9', size: 0.75 }, catGridLine: { style: 'none' },
  });
  title(s, 'WeCare slide ', 'chart', 3.938, 0.96, 5.458, 0.774, { align: 'center' });
  chrome(s);
}

/** 35 - contact */
function slide35(s) {
  rect(s, 0, 0, 13.336, 3.889, tint(6));
  const columns = [
    { head: 'GET IN TOUCH', hx: 1.243, hy: 5.43, bx: 1.243, by: 5.84, lines: ['(+12) 3 4567 890', '(+12) 3 4567 325'], hw: 1.706 },
    { head: 'OUR ADDRESS', hx: 4.164, hy: 5.43, bx: 4.164, by: 5.84, lines: ['Urban Street Subway', 'Texas 1225. US'], hw: 1.706 },
    { head: 'OFFICE HOURS', hx: 7.284, hy: 5.45, bx: 7.284, by: 5.819, lines: ['Monday – Thursday', '08AM – 16PM'], hw: 1.764 },
    { head: 'FOLLOW US', hx: 10.204, hy: 5.45, bx: 10.204, by: 5.819, lines: ['www.company.com', 'info@company.com'], hw: 1.536 },
  ];
  columns.forEach((col) => {
    s.addText(col.head, {
      x: col.hx, y: col.hy, w: col.hw, h: 0.3, valign: 'top',
      fontFace: F.MB, fontSize: 11, color: C.MUTED, charSpacing: 2,
    });
    body(s, col.lines.map((t) => ({ text: t, options: { breakLine: true } })),
      col.bx, col.by, 2.085, 0.675);
  });
  body(s, SED.full + ', eaque ipsa quae ab illo inventore', 5.319, 1.668, 6.771, 0.675);
  title(s, 'Contact ', 'us', 1.243, 1.471, 3.861, 0.774);
  chrome(s);
}

/** 36 - closing break slide */
function slide36(s) {
  s.addText([
    { text: 'Break', options: { color: C.DARK } },
    { text: 'Slide', options: { color: C.ORANGE } },
  ], { x: 4.438, y: 2.83, w: 7.39, h: 1.447, valign: 'top', align: 'right', fontFace: F.DISPLAY, fontSize: 80 });
  s.addText('FOR WATCHING OUR SLIDES', {
    x: 7.328, y: 4.334, w: 4.499, h: 0.337, valign: 'top', align: 'right',
    fontFace: F.SEMI, fontSize: 14, color: C.BODY, charSpacing: 1,
  });
  chrome(s);
}

/* ------------------------------------------------------ 37-40: icon sheets */
/*
 * The last four pages of the template are plain reference sheets: 55 solid
 * black pictograms laid out 11 across and 5 down on a white page. The glyphs
 * are one-off freeform artwork, so each slot is redrawn here with the closest
 * native preset shape - same grid, same size, same weight.
 */

const ICON_COLS = 11;
const ICON_SIZE = 0.42;
const ICON_X0 = 1.1;
const ICON_DX = 1.113;
const ICON_Y0 = 1.4;
const ICON_DY = 1.175;

/* Each cell is a list of [preset, x, y, w, h, carve?] layers - positions and
   sizes are percentages of the 0.42in cell, and a trailing 1 means the layer is
   punched out in the page colour. The presets and boxes were chosen to match the
   silhouette of the corresponding glyph in the reference sheet. */
const ICONS_37 = [
  [['rightBrace',17,0,69,100],['frame',41,39,58,59]],
  [['snip2SameRect',0,11,100,89]],
  [['pieWedge',2,36,98,64],['diagStripe',2,36,98,64,1]],
  [['downArrowCallout',0,19,100,67],['mathMultiply',39,42,59,56,1]],
  [['mathMultiply',0,14,100,86],['rtTriangle',0,14,44,38,1]],
  [['pentagon',8,8,84,84],['blockArc',0,0,100,44,1]],
  [['round2SameRect',0,14,100,86],['lightningBolt',0,61,100,38,1]],
  [['leftArrow',0,14,91,86],['leftArrow',0,61,91,38]],
  [['octagon',12,17,64,83],['downArrow',0,17,91,83]],
  [['pentagon',12,14,64,86],['quadArrow',0,14,89,86]],
  [['flowChartDelay',0,28,87,59]],
  [['flowChartOnlineStorage',0,0,100,100],['bentUpArrow',39,39,59,59,1]],
  [['flowChartMultidocument',0,0,100,100],['flowChartExtract',0,0,100,44,1]],
  [['flowChartMultidocument',0,0,100,100],['flowChartExtract',0,0,100,44,1]],
  [['snip2SameRect',0,0,100,100],['donut',0,0,100,100,1]],
  [['flowChartDisplay',0,5,100,95],['curvedUpArrow',39,42,59,56,1]],
  [['mathEqual',0,12,100,75]],
  [['ribbon2',19,0,59,100]],
  [['pieWedge',0,0,100,100]],
  [['arc',0,2,95,98]],
  [['rightBracket',0,14,100,84],['star4',0,14,100,38,1]],
  [['pieWedge',0,0,86,100],['halfFrame',34,39,52,59,1]],
  [['flowChartPunchedCard',0,12,100,75]],
  [['flowChartMagneticDisk',0,0,100,100],['rightArrow',39,39,59,59,1]],
  [['pieWedge',6,0,94,100],['rightArrowCallout',6,0,56,59,1]],
  [['snip2SameRect',0,0,100,100]],
  [['bevel',0,0,100,100]],
  [['flowChartManualInput',0,0,100,100],['curvedRightArrow',0,0,59,59,1]],
  [['flowChartDocument',11,0,75,100],['rtTriangle',52,0,33,100,1]],
  [['bevel',22,0,50,100]],
  [['flowChartOffpageConnector',2,0,87,100]],
  [['frame',0,0,100,100],['swooshArrow',39,0,59,59]],
  [['flowChartDocument',0,0,97,100]],
  [['leftUpArrow',19,2,81,87],['leftCircularArrow',19,36,48,52]],
  [['flowChartPunchedTape',0,0,100,100],['bevel',39,0,59,59]],
  [['moon',0,0,100,97],['frame',55,0,44,97]],
  [['snip2SameRect',0,0,100,95]],
  [['mathEqual',0,0,100,100],['heart',0,55,100,44,1]],
  [['bevel',12,0,77,100]],
  [['upArrowCallout',0,0,100,100]],
  [['ellipseRibbon',0,0,100,100],['teardrop',55,0,44,100]],
  [['diagStripe',0,0,98,100],['bentArrow',0,0,58,59,1]],
  [['frame',0,0,100,95],['mathDivide',0,0,100,42]],
  [['frame',0,0,100,100],['cube',0,0,59,59,1]],
  [['plaque',8,0,92,92]],
  [['flowChartOffpageConnector',6,0,94,94]],
  [['flowChartOffpageConnector',3,0,97,94]],
  [['frame',0,0,100,100],['flowChartDelay',39,0,59,59,1]],
  [['donut',0,0,100,92]],
  [['plaque',0,12,100,59],['diagStripe',55,0,44,84,1]],
  [['snip2SameRect',0,0,100,78]],
  [['frame',0,0,100,100],['bentUpArrow',0,0,59,59]],
  [['bevel',2,0,87,100]],
  [['flowChartManualInput',0,0,94,100],['verticalScroll',0,0,94,44,1]],
  [['bevel',0,0,100,91]],
];
const ICONS_38 = [
  [['plaque',0,5,100,95]],
  [['bevel',0,16,100,84]],
  [['heptagon',0,5,100,95]],
  [['pieWedge',8,12,84,80],['leftUpArrow',39,42,59,56,1]],
  [['mathMultiply',0,16,100,84],['homePlate',0,16,44,84]],
  [['donut',8,12,84,80],['upDownArrowCallout',0,5,44,42]],
  [['frame',8,5,86,95],['verticalScroll',8,56,86,42,1]],
  [['frame',9,5,84,95],['verticalScroll',9,56,84,42,1]],
  [['horizontalScroll',5,5,95,95],['chevron',42,42,56,56,1]],
  [['frame',0,5,100,95],['donut',39,42,59,56]],
  [['plaque',0,5,100,95]],
  [['bevel',0,12,100,84],['rightArrow',39,45,59,50,1]],
  [['flowChartMultidocument',0,2,100,98],['rightBrace',0,2,44,44,1]],
  [['mathEqual',0,12,100,84],['flowChartExtract',0,12,100,38,1]],
  [['bracePair',0,6,100,94]],
  [['flowChartTerminator',0,17,100,75],['halfFrame',0,17,44,33,1]],
  [['flowChartMagneticTape',0,17,100,75],['snip2DiagRect',0,58,100,33]],
  [['bentArrow',8,2,86,98],['frame',8,2,86,98]],
  [['flowChartOffpageConnector',0,6,100,94]],
  [['plaque',0,2,100,98],['rtTriangle',39,2,59,58,1]],
  [['bevel',0,12,100,84]],
  [['snip2SameRect',0,12,100,84]],
  [['snip2SameRect',2,0,95,100],['star8',2,55,95,44,1]],
  [['bevel',6,0,86,100]],
  [['octagon',0,0,100,100]],
  [['ellipseRibbon2',0,0,100,100]],
  [['frame',0,2,100,97],['bentArrow',39,2,59,58]],
  [['frame',0,2,100,97],['bentArrow',39,2,59,58]],
  [['frame',0,2,100,97],['bentArrow',0,2,59,58]],
  [['bevel',0,0,100,100]],
  [['frame',5,2,95,97],['star4',5,2,95,97]],
  [['snip1Rect',9,0,86,100]],
  [['frame',9,0,86,100],['leftCircularArrow',9,39,52,59]],
  [['frame',6,0,84,100],['flowChartExtract',6,55,84,44]],
  [['flowChartDocument',6,0,86,100]],
  [['leftRightArrowCallout',19,8,61,84],['diagStripe',12,0,44,59,1]],
  [['gear6',20,0,58,100],['chevron',42,0,34,59]],
  [['flowChartMerge',0,0,100,100],['leftBrace',39,39,59,59,1]],
  [['flowChartManualOperation',19,0,64,98]],
  [['pieWedge',8,8,84,84],['leftBrace',39,39,59,59,1]],
  [['flowChartOffpageConnector',3,14,97,72],['pieWedge',41,0,58,59]],
  [['nonIsoscelesTrapezoid',14,0,72,100],['mathDivide',0,55,100,44]],
  [['snip1Rect',0,5,100,84]],
  [['flowChartManualOperation',27,0,53,100]],
  [['downArrowCallout',0,0,100,97],['arc',0,53,100,42,1]],
  [['flowChartMerge',0,0,100,97]],
  [['bevel',0,2,100,84]],
  [['plaque',0,0,100,97]],
  [['frame',0,2,100,84],['bentArrow',39,2,59,50]],
  [['heart',0,0,100,91],['flowChartTerminator',0,0,100,41]],
  [['snip2SameRect',0,0,100,97],['upArrow',0,0,100,42,1]],
  [['bevel',9,0,84,97]],
  [['can',14,0,75,97]],
  [['plaque',0,0,100,97],['moon',39,0,59,58,1]],
  [['snip2SameRect',0,0,100,97]],
];
const ICONS_39 = [
  [['irregularSeal2',0,0,100,98],['cloudCallout',0,39,59,58]],
  [['leftArrowCallout',14,9,72,69],['uturnArrow',0,36,59,41]],
  [['heart',6,0,87,98],['corner',53,0,39,98,1]],
  [['flowChartMultidocument',14,8,70,81],['heart',41,0,48,58]],
  [['donut',0,0,100,95],['diagStripe',0,0,44,42]],
  [['leftUpArrow',6,0,89,89],['swooshArrow',55,0,39,89]],
  [['flowChartMultidocument',9,0,86,94]],
  [['leftArrowCallout',2,14,98,58]],
  [['cloudCallout',2,0,98,89]],
  [['bracketPair',0,0,100,89]],
  [['round1Rect',25,0,58,100],['funnel',25,0,58,44,1]],
  [['star10',0,0,98,100]],
  [['dodecagon',0,0,100,98]],
  [['arc',19,0,63,100]],
  [['bentUpArrow',5,3,91,91],['leftArrow',53,3,41,91]],
  [['frame',17,0,67,100],['leftCircularArrow',17,55,67,44]],
  [['plus',16,0,72,100]],
  [['flowChartOffpageConnector',6,14,91,78]],
  [['bevel',12,19,80,58]],
  [['flowChartPunchedTape',5,14,95,69],['flowChartDelay',5,0,42,42,1]],
  [['ellipseRibbon',19,0,67,98],['leftCircularArrow',6,0,94,44]],
  [['doubleWave',0,20,100,55]],
  [['can',0,0,100,100],['curvedUpArrow',39,39,59,59,1]],
  [['bevel',0,16,100,69]],
  [['curvedDownArrow',5,6,89,87],['blockArc',5,53,89,39]],
  [['upArrowCallout',11,0,78,100],['downArrow',11,0,78,100]],
  [['upArrowCallout',11,0,78,100],['nonIsoscelesTrapezoid',11,0,78,100]],
  [['heart',23,5,56,89],['funnel',23,5,56,39,1]],
  [['flowChartMerge',14,0,72,100],['frame',0,55,100,44]],
  [['upArrowCallout',2,0,98,100],['corner',2,0,44,100]],
  [['flowChartPunchedTape',8,5,89,91],['homePlate',8,5,39,41,1]],
  [['leftCircularArrow',14,0,78,100],['halfFrame',14,0,34,44]],
  [['donut',0,0,100,100],['bentUpArrow',55,0,44,100]],
  [['flowChartMerge',0,9,97,89],['ellipseRibbon2',0,9,42,89]],
  [['star10',0,9,100,89],['upArrowCallout',0,9,100,39]],
  [['bracketPair',0,22,100,64],['notchedRightArrow',0,22,44,64,1]],
  [['plaque',0,0,100,100]],
  [['dodecagon',2,0,97,100]],
  [['plaque',0,0,100,100]],
  [['flowChartTerminator',0,11,100,86],['leftRightUpArrow',0,58,100,38,1]],
  [['downArrowCallout',0,19,100,73],['flowChartMerge',0,47,59,44,1]],
  [['horizontalScroll',0,14,100,78],['mathMultiply',0,14,100,34]],
  [['heart',8,3,91,97],['funnel',8,3,91,97]],
  [['upArrowCallout',2,2,98,98],['upArrowCallout',2,2,44,44,1]],
  [['ribbon',12,11,72,89],['rightBrace',41,45,42,53,1]],
  [['ribbon',12,11,72,89],['leftBrace',12,59,72,39,1]],
  [['plaque',0,5,100,95]],
  [['parallelogram',11,17,78,69],['leftArrowCallout',42,41,47,58,1]],
  [['bevel',0,8,100,92],['bentArrow',55,8,44,92,1]],
  [['noSmoking',0,27,100,63]],
  [['flowChartManualInput',0,11,100,86],['leftArrow',0,58,100,38,1]],
  [['plaque',0,5,100,95],['flowChartDecision',0,5,100,95,1]],
  [['frame',0,5,100,95],['donut',0,5,100,42]],
  [['plaque',0,5,100,95],['bentUpArrow',39,42,59,56,1]],
  [['plaque',0,5,100,95],['bentUpArrow',39,42,59,56,1]],
];
const ICONS_40 = [
  [['flowChartMagneticDisk',0,12,100,87]],
  [['ribbon',2,16,98,84]],
  [['flowChartExtract',23,12,61,87],['plaque',23,59,61,39]],
  [['snip2SameRect',3,20,97,80],['flowChartExtract',3,64,97,34,1]],
  [['flowChartMerge',11,16,89,84]],
  [['snip2SameRect',3,12,97,87]],
  [['snip2DiagRect',6,14,94,86]],
  [['flowChartExtract',3,16,97,84],['heart',3,61,97,38,1]],
  [['leftArrowCallout',17,25,75,69],['diagStripe',11,19,89,36,1]],
  [['flowChartOffpageConnector',20,9,78,91]],
  [['snip2SameRect',22,16,77,84],['curvedRightArrow',22,48,45,50,1]],
  [['donut',8,16,89,84],['flowChartDecision',8,16,89,84,1]],
  [['mathEqual',0,22,100,78],['stripedRightArrow',39,53,59,47,1]],
  [['plaque',2,22,98,78],['halfFrame',2,22,44,34,1]],
  [['snip2DiagRect',2,8,98,92],['bentUpArrow',41,44,58,55,1]],
  [['blockArc',0,11,100,89],['upArrowCallout',39,11,59,53]],
  [['teardrop',3,28,97,59],['homePlate',56,16,42,84]],
  [['leftArrowCallout',3,30,97,63],['lightningBolt',3,30,42,63]],
  [['teardrop',17,14,81,86],['flowChartMerge',17,14,48,52,1]],
  [['heart',9,17,91,83],['flowChartMultidocument',58,17,41,83]],
  [['nonIsoscelesTrapezoid',28,8,59,92],['leftRightUpArrow',16,8,38,92,1]],
  [['heart',12,20,87,80],['chevron',12,64,87,34]],
  [['donut',2,11,98,89],['leftRightRibbon',2,59,98,39,1]],
  [['heptagon',25,6,55,94]],
  [['flowChartManualOperation',19,5,67,95],['funnel',5,5,95,95]],
  [['noSmoking',0,2,100,98],['homePlate',39,41,59,58]],
  [['flowChartExtract',3,6,97,94],['wedgeEllipseCallout',3,58,97,42]],
  [['plaque',3,3,97,97],['flowChartMerge',3,56,97,42,1]],
  [['upArrowCallout',2,17,98,72],['leftRightCircularArrow',2,17,98,72]],
  [['upArrowCallout',14,6,86,94],['diagStripe',14,6,86,42,1]],
  [['downArrowCallout',14,6,86,94],['bentUpArrow',61,6,38,94,1]],
  [['plaque',16,2,84,98],['flowChartMerge',16,55,84,44,1]],
  [['star5',12,9,87,91],['flowChartMerge',47,45,52,53]],
  [['flowChartMagneticDisk',16,3,73,97]],
  [['plaque',0,0,100,100],['notchedRightArrow',0,55,100,44,1]],
  [['star5',6,6,94,94]],
  [['bracePair',5,3,95,97]],
  [['pieWedge',2,25,98,48]],
  [['snip2DiagRect',6,0,94,100]],
  [['downArrowCallout',17,8,69,87],['flowChartMagneticTape',3,8,97,39]],
  [['mathMultiply',3,3,97,97],['mathNotEqual',41,3,58,58]],
  [['mathEqual',9,8,91,87],['uturnArrow',9,42,53,52,1]],
  [['ellipse',16,3,84,97],['upArrowCallout',16,3,84,42]],
  [['moon',19,19,81,64],['flowChartCollate',62,6,36,91]],
  [['mathMultiply',27,20,52,56],['leftCircularArrow',27,20,52,25]],
  [['plaque',0,0,100,100],['diamond',0,0,100,100,1]],
  [['snipRoundRect',9,5,89,87],['upArrow',9,5,89,39,1]],
  [['pieWedge',19,2,67,94],['downArrowCallout',5,2,56,56,1]],
  [['plus',12,11,80,80],['flowChartPunchedTape',44,3,56,56]],
  [['heart',14,12,84,72],['homePlate',47,41,50,42,1]],
  [['donut',5,0,95,100]],
  [['bracketPair',3,6,97,84]],
  [['heart',8,0,92,100],['moon',44,0,55,59,1]],
  [['snip2SameRect',16,6,84,86]],
  [['bracketPair',5,5,95,87]],
];

/** Lay out one 11 x 5 sheet of black pictograms. */
function iconSheet(s, icons) {
  icons.forEach((layers, i) => {
    const ox = ICON_X0 + (i % ICON_COLS) * ICON_DX - ICON_SIZE / 2;
    const oy = ICON_Y0 + Math.floor(i / ICON_COLS) * ICON_DY - ICON_SIZE / 2;
    layers.forEach(([shape, x, y, w, h, carve]) => {
      s.addShape(shape, {
        x: ox + (x / 100) * ICON_SIZE, y: oy + (y / 100) * ICON_SIZE,
        w: (w / 100) * ICON_SIZE, h: (h / 100) * ICON_SIZE,
        fill: { color: carve ? C.BG : '000000' },
      });
    });
  });
}

/* ---------------------------------------------------------------- assembly */

const SLIDES = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
  slide31, slide32, slide33, slide34, slide35, slide36,
  (s) => iconSheet(s, ICONS_37),
  (s) => iconSheet(s, ICONS_38),
  (s) => iconSheet(s, ICONS_39),
  (s) => iconSheet(s, ICONS_40),
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WECARE', width: 13.333, height: 7.5 });
  pptx.layout = 'WECARE';
  pptx.title = 'WeCare - Pet Animal Presentation Template';
  pptx.theme = { headFontFace: F.TITLE, bodyFontFace: F.BODY };

  SLIDES.forEach((draw) => {
    const slide = pptx.addSlide();
    slide.background = { color: C.BG };
    draw(slide);
  });

  const out = path.join(__dirname, '16a2dd8b-3d57-47f5-966f-3a138f95bccb_grok_final.pptx');
  return pptx.writeFile({ fileName: out }).then(() => console.log('wrote ' + out));
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
