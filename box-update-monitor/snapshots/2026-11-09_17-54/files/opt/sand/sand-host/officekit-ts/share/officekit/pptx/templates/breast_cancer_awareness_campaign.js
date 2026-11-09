/**
 * "Awareness Can Saves Lives" — Restnaire breast-cancer awareness deck.
 * Standalone pptxgenjs recreation (15 slides, 13.333in x 7.5in).
 *
 * Photographs in the source deck are replaced by flat grey placeholder shapes
 * that keep the original silhouette (clouds, circles, rounded cards, ...).
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const C = {
  blush: 'F4D8DC',   // accent1 - page background
  pink: 'E18AA8',    // accent2 - primary brand pink
  cream: 'FCF5F6',   // accent3 - off-white cards
  orange: 'DC7340',  // accent4 - accent dots
  ink: '19131A',     // accent5 - headings
  white: 'FFFFFF',
  body: '595959',    // tx1 lum 65/35 - body copy
  dark: '0D0D0D',    // tx1 lum 95/5
  black: '000000',
  photo: 'CCCCCC',   // stand-in for the stock photos
  photoTx: 'DCDCDC',
};

const HEAD = 'Quicksand Bold';
const BODY = 'Roboto';

const RING = 6;      // pt - outline weight used by every decorative ring
const NONE = { type: 'none' };

/* ------------------------------------------------------- generic helpers */

/** Text box. Defaults mirror the deck: top aligned, Quicksand Bold, ink. */
function txt(slide, text, o) {
  slide.addText(text, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: o.face || HEAD,
    fontSize: o.size || 18,
    color: o.color || C.ink,
    align: o.align || 'left',
    valign: 'top',
    lineSpacingMultiple: o.ls,
    margin: [7.2, 7.2, 3.6, 3.6], // pptxgenjs order: left, right, bottom, top
    isTextBox: true,
  });
}

/** 11pt Roboto grey paragraph with the deck's 150% leading. */
function para(slide, text, o) {
  txt(slide, text, Object.assign({ face: BODY, size: 11, color: C.body, ls: 1.5 }, o));
}

/** 32pt Quicksand section heading. */
function heading(slide, text, o) {
  txt(slide, text, Object.assign({ size: 32, color: C.ink }, o));
}

function ellipse(slide, x, y, w, h, fill) {
  slide.addShape('ellipse', { x, y, w, h: h === undefined ? w : h, fill: { color: fill }, line: NONE });
}

function ring(slide, x, y, d, color, width) {
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: NONE, line: { color, width: width || RING } });
}

function roundRect(slide, o) {
  slide.addShape('roundRect', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: { color: o.fill }, line: NONE,
    rectRadius: o.r === undefined ? 0.0835 * Math.min(o.w, o.h) : o.r,
    rotate: o.rotate,
  });
}

function vline(slide, x, y, h, color) {
  slide.addShape('line', { x, y, w: 0, h, line: { color, width: 0.5 } });
}

/** Freeform path — see the PATHS table below. Coordinates are shape-local. */
function freeform(slide, pts, o) {
  slide.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: o.fill ? { color: o.fill } : NONE,
    line: o.line ? { color: o.line, width: o.lineWidth || RING } : NONE,
    rotate: o.rotate,
    points: toPoints(pts),
  });
}

/**
 * Compact segment lists -> pptxgenjs point objects.
 *   []            close the sub path
 *   [x,y]         moveTo (first of a sub path) / lineTo
 *   [x1,y1,x2,y2,x,y]  cubic bezier
 */
function toPoints(segs) {
  const out = [];
  let fresh = true;
  segs.forEach(s => {
    if (s.length === 0) { out.push({ close: true }); fresh = true; return; }
    if (s.length === 2) {
      out.push({ x: s[0], y: s[1], moveTo: fresh });
    } else {
      out.push({ x: s[4], y: s[5], curve: { type: 'cubic', x1: s[0], y1: s[1], x2: s[2], y2: s[3] } });
    }
    fresh = false;
  });
  return out;
}

/** Grey stand-in for a photo, clipped to the original outline. */
function photo(slide, o) {
  if (o.points) {
    freeform(slide, o.points, { x: o.x, y: o.y, w: o.w, h: o.h, fill: C.photo });
  } else if (o.shape === 'ellipse') {
    ellipse(slide, o.x, o.y, o.w, o.h, C.photo);
  } else if (o.r) {
    roundRect(slide, { x: o.x, y: o.y, w: o.w, h: o.h, fill: C.photo, r: o.r });
  } else {
    slide.addShape('rect', { x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: C.photo }, line: NONE });
  }
  // Full-bleed silhouettes get no caption: their bounding-box centre is outside the shape.
  if (o.label === false) return;
  slide.addText('[image]', {
    x: o.x, y: o.y + o.h / 2 - 0.2, w: o.w, h: 0.4,
    fontFace: BODY, fontSize: 12, color: C.photoTx, align: 'center', valign: 'middle',
  });
}

/** "Restnaire" wordmark + dotted sun-burst logo, present on every slide. */
function brand(slide, color, markColor, o) {
  o = o || {};
  const x = o.x === undefined ? 0.618 : o.x;
  const y = o.y === undefined ? 0.532 : o.y;
  const d = o.d === undefined ? 0.174 : o.d;
  const r = d / 2;
  for (let i = 0; i < 12; i++) {
    const a = (i * Math.PI) / 6;
    const big = i % 2 === 0 ? 0.052 * d / 0.174 * 0.174 : 0.03;
    ellipse(slide, x + r + Math.cos(a) * r * 0.78 - big / 2, y + r + Math.sin(a) * r * 0.78 - big / 2, big, big, markColor);
  }
  txt(slide, 'Restnaire', {
    x: o.tx === undefined ? 0.881 : o.tx, y: o.ty === undefined ? 0.484 : o.ty,
    w: 1.3, h: 0.269, size: 10, color,
  });
}

/* ---------------------------------------------------------------- shapes */
/* Freeform outlines lifted from the source deck (inches, local origin).    */

const P = {
  // s1: cloud silhouette running along the bottom of the title slide
  cloud1: [[13.33,0],[13.33,5.342],[0,5.342],[0,1.316],[0.087,1.292],[0.405,1.21,0.739,1.166,1.082,1.166],[1.268,1.171],[1.265,1.146],[1.265,0.795,1.55,0.511,1.9,0.511],[2.251,0.511,2.536,0.795,2.536,1.146],[2.536,1.234,2.518,1.317,2.486,1.393],[2.473,1.417],[2.632,1.479],[3.108,1.681,3.537,1.972,3.897,2.332],[4.121,2.579],[4.13,2.567],[4.376,2.27,4.747,2.08,5.162,2.08],[5.623,2.08,6.03,2.314,6.27,2.67],[6.354,2.824],[6.481,2.805],[6.541,2.799,6.603,2.796,6.665,2.796],[7.1,2.796,7.5,2.95,7.811,3.207],[7.855,3.247],[7.869,3.226],[8.403,2.435,9.308,1.915,10.334,1.915],[10.911,1.915,11.45,2.079,11.906,2.364],[12.021,2.445],[12.037,2.401],[12.152,2.128,12.384,1.916,12.67,1.827],[12.709,1.818],[12.674,1.761],[12.602,1.629,12.553,1.481,12.534,1.325],[12.528,1.245],[12.523,1.245],[12.275,1.245,12.074,1.044,12.074,0.795],[12.074,0.547,12.275,0.345,12.523,0.345],[12.586,0.345,12.645,0.358,12.699,0.381],[12.772,0.421],[12.808,0.372],[12.923,0.234,13.067,0.12,13.23,0.041],[]],
  // s2: big pink hill (layout), top-left corner wedge, bottom-right hump, photo cloud
  hill2:  [[2.765,0],[5.432,0,7.658,1.892,8.172,4.407],[8.203,4.58],[0,4.58],[0,0.743],[0.134,0.666],[0.916,0.241,1.812,0,2.765,0],[]],
  wedge2: [[0,0],[2.668,0],[2.61,0.16],[2.31,0.869,1.608,1.366,0.791,1.366],[0.518,1.366,0.259,1.311,0.022,1.211],[0,1.2],[]],
  hump2:  [[1.328,0],[2.018,0,2.585,0.524,2.653,1.195],[2.656,1.265],[0,1.265],[0.004,1.195],[0.072,0.524,0.639,0,1.328,0],[]],
  cloud2: [[4.928,0],[5.56,0,6.073,0.513,6.073,1.145],[6.073,1.778,5.56,2.291,4.928,2.291],[4.82,2.285],[4.95,2.459],[5.21,2.845,5.362,3.309,5.362,3.809],[5.362,3.892,5.358,3.975,5.35,4.056],[5.343,4.102],[5.459,4.114],[5.819,4.187,6.089,4.505,6.089,4.886],[6.089,5.321,5.736,5.674,5.301,5.674],[5.083,5.674,4.886,5.586,4.743,5.443],[4.733,5.43],[4.655,5.516],[4.218,5.953,3.615,6.223,2.948,6.223],[2.365,6.223,1.83,6.016,1.413,5.672],[1.344,5.609],[1.341,5.64],[1.278,5.948,1.004,6.18,0.677,6.18],[0.303,6.18,0,5.877,0,5.503],[0,5.129,0.303,4.826,0.677,4.826],[0.766,4.835],[0.724,4.749],[0.602,4.46,0.535,4.142,0.535,3.809],[0.535,2.476,1.615,1.395,2.948,1.395],[3.198,1.395,3.439,1.433,3.666,1.504],[3.867,1.577],[3.834,1.486],[3.8,1.378,3.782,1.264,3.782,1.145],[3.782,0.513,4.295,0,4.928,0],[]],
  // s3: scalloped pink band across the bottom
  wave3: [[1.924,0],[2.469,0,2.969,0.193,3.359,0.515],[3.454,0.602],[3.549,0.515],[3.939,0.193,4.438,0,4.984,0],[5.529,0,6.029,0.193,6.419,0.515],[6.557,0.641],[6.695,0.515],[7.085,0.193,7.585,0,8.13,0],[8.675,0,9.175,0.193,9.565,0.515],[9.66,0.602],[9.755,0.515],[10.145,0.193,10.645,0,11.19,0],[12.125,0,12.927,0.568,13.269,1.378],[13.33,1.545],[13.33,2.18],[0,2.18],[0,1.082],[0.053,0.995],[0.458,0.395,1.145,0,1.924,0],[]],
  cloud5: [[1.608,0],[2.226,0,2.807,0.156,3.315,0.432],[3.513,0.553],[3.523,0.521],[3.617,0.299,3.837,0.144,4.093,0.144],[4.434,0.144,4.71,0.42,4.71,0.761],[4.71,0.975,4.602,1.163,4.438,1.274],[4.375,1.308],[4.576,1.578],[4.769,1.864,4.922,2.179,5.027,2.515],[5.07,2.676],[0,2.676],[0,0.385],[0.215,0.281],[0.643,0.1,1.114,0,1.608,0],[]],
  hill6:  [[0.156,0],[2.09,0,3.657,1.568,3.657,3.501],[3.642,3.801],[3.653,3.789],[4.095,3.346,4.707,3.073,5.382,3.073],[6.31,3.073,7.117,3.59,7.531,4.352],[7.583,4.459],[0,4.459],[0,0.004],[]],
  cloud6: [[2.162,0.266],[2.331,0.266,2.468,0.403,2.468,0.573],[2.468,0.657,2.434,0.734,2.378,0.789],[2.285,0.852],[2.301,0.871],[2.407,1.027,2.468,1.215,2.468,1.417],[2.468,1.957,2.031,2.394,1.491,2.394],[0.952,2.394,0.515,1.957,0.515,1.417],[0.515,0.878,0.952,0.441,1.491,0.441],[1.559,0.441,1.625,0.448,1.688,0.461],[1.861,0.514],[1.862,0.511],[1.89,0.371,2.014,0.266,2.162,0.266],[],[4.375,0],[5.22,0,5.905,0.685,5.905,1.53],[5.905,2.058,5.637,2.524,5.23,2.799],[5.185,2.826],[5.279,2.951],[5.495,3.271,5.621,3.657,5.621,4.072],[5.621,5.178,4.724,6.075,3.617,6.075],[2.926,6.075,2.316,5.725,1.956,5.192],[1.923,5.138],[1.85,5.271],[1.669,5.54,1.361,5.717,1.012,5.717],[0.453,5.717,0,5.264,0,4.705],[0,4.147,0.453,3.694,1.012,3.694],[1.221,3.694,1.416,3.758,1.577,3.867],[1.622,3.904],[1.624,3.867],[1.699,3.126,2.178,2.504,2.837,2.226],[2.987,2.171],[2.965,2.126],[2.887,1.943,2.844,1.741,2.844,1.53],[2.844,0.685,3.529,0,4.375,0],[]],
  cloud7: [[6.439,0],[12.626,0],[12.626,4.188],[12.528,4.261],[11.892,4.691,11.126,4.942,10.301,4.942],[9.064,4.942,7.958,4.377,7.227,3.492],[7.127,3.358],[7.086,3.517],[6.87,4.209,6.225,4.712,5.461,4.712],[4.698,4.712,4.052,4.209,3.836,3.517],[3.828,3.486],[3.807,3.514],[3.495,3.892,3.024,4.133,2.496,4.133],[2.261,4.133,2.038,4.085,1.835,3.999],[1.77,3.968],[1.77,3.971],[1.725,4.418,1.347,4.768,0.887,4.768],[0.397,4.768,0,4.37,0,3.88],[0,3.39,0.397,2.993,0.887,2.993],[0.893,2.993],[0.873,2.939],[0.824,2.779,0.797,2.61,0.797,2.434],[0.797,1.495,1.558,0.735,2.496,0.735],[3.258,0.735,3.903,1.237,4.118,1.929],[4.126,1.959],[4.148,1.929],[4.46,1.551,4.933,1.31,5.461,1.31],[5.755,1.31,6.031,1.384,6.272,1.515],[6.37,1.575],[6.338,1.366],[6.325,1.232,6.318,1.096,6.318,0.959],[6.318,0.684,6.346,0.415,6.399,0.156],[]],
  wave8:  [[0.584,0],[0.769,0.048],[1.762,0.356,2.482,1.282,2.482,2.376],[2.482,3.217,2.056,3.959,1.407,4.397],[1.28,4.475],[1.311,4.476],[1.862,4.532,2.33,4.87,2.569,5.343],[2.605,5.423],[2.71,5.365],[2.8,5.327,2.899,5.306,3.003,5.306],[3.211,5.306,3.399,5.39,3.536,5.526],[3.554,5.549],[3.682,5.339],[4.239,4.515,5.182,3.973,6.251,3.973],[7.214,3.973,8.074,4.412,8.642,5.101],[8.659,5.123],[8.743,5.102],[8.854,5.079,8.969,5.067,9.087,5.067],[9.619,5.067,10.093,5.309,10.407,5.689],[10.488,5.798],[10.514,5.716],[10.7,5.275,11.137,4.966,11.645,4.966],[11.73,4.966,11.813,4.974,11.893,4.991],[11.983,5.014],[11.973,4.814],[11.973,3.833,12.619,3.003,13.509,2.726],[13.7,2.677],[13.286,6.443],[13.25,6.771],[12.65,6.705],[0,5.314],[0.013,5.193],[0.079,4.598],[]],
  hill9:  [[3.54,0],[3.56,0,3.58,0.002,3.599,0.006],[3.627,0.015],[3.627,4.335],[0,4.335],[0.047,4.249],[0.214,4.002,0.496,3.84,0.817,3.84],[0.872,3.842],[0.832,3.586],[0.822,3.486,0.817,3.384,0.817,3.282],[0.817,1.844,1.837,0.644,3.193,0.367],[3.255,0.357],[3.253,0.352],[3.249,0.333,3.247,0.313,3.247,0.293],[3.247,0.131,3.378,0,3.54,0],[]],
  arc10:  [[0.54,0],[1.193,0,1.722,0.529,1.722,1.182],[1.722,1.264,1.714,1.344,1.698,1.421],[1.657,1.553],[0,1.553],[0,0.131],[0.08,0.093],[0.221,0.033,0.377,0,0.54,0],[]],
  hill10: [[2.916,0],[2.956,0.004],[2.956,2.252],[0,2.252],[0.013,2.209],[0.092,2.022,0.278,1.891,0.494,1.891],[0.53,1.891,0.565,1.895,0.599,1.902],[0.641,1.915],[0.639,1.899],[0.633,1.849,0.631,1.798,0.631,1.747],[0.631,0.925,1.297,0.258,2.119,0.258],[2.274,0.258,2.422,0.282,2.562,0.325],[2.618,0.346],[2.613,0.303],[2.613,0.136,2.749,0,2.916,0],[]],
  top11:  [[0.103,0],[2.904,0],[2.939,0.096],[2.983,0.237,3.007,0.388,3.007,0.543],[3.007,1.374,2.334,2.047,1.503,2.047],[0.673,2.047,0,1.374,0,0.543],[0,0.388,0.024,0.237,0.068,0.096],[]],
  side11: [[3.051,0],[3.051,3.746],[0.063,3.746],[0.016,3.441],[0.005,3.336,0,3.229,0,3.121],[0,1.503,1.23,0.172,2.806,0.012],[]],
  // s11 photo tiles: rounded on the two corners that face away from the gutter
  tile11a: [[0,0],[3.799,0],[3.799,2.246],[3.799,2.45,3.634,2.615,3.43,2.615],[0.368,2.615],[0.165,2.615,0,2.45,0,2.246],[]],
  tile11b: [[0,0],[3.799,0],[3.799,4.066],[3.799,4.277,3.628,4.448,3.417,4.448],[0.382,4.448],[0.171,4.448,0,4.277,0,4.066],[]],
  tile11c: [[0.382,0],[3.417,0],[3.628,0,3.799,0.171,3.799,0.382],[3.799,4.365],[0,4.365],[0,0.382],[0,0.171,0.171,0,0.382,0],[]],
  tile11d: [[0.368,0],[3.43,0],[3.634,0,3.799,0.165,3.799,0.368],[3.799,2.531],[0,2.531],[0,0.368],[0,0.165,0.165,0,0.368,0],[]],
  arc12:  [[2.697,0],[3.687,0,4.489,0.803,4.489,1.792],[4.489,1.854,4.486,1.915,4.48,1.976],[4.48,1.976],[0,1.976],[0.023,1.902],[0.151,1.601,0.449,1.389,0.797,1.389],[0.826,1.389,0.854,1.391,0.883,1.394],[0.948,1.404],[0.985,1.259],[1.212,0.53,1.893,0,2.697,0],[]],
  hill12: [[2.335,0],[3.714,0],[3.714,7.5],[0.501,7.5],[0.487,7.481],[0.179,7.026,0,6.478,0,5.888],[0,4.412,1.121,3.199,2.558,3.053],[2.599,3.051],[2.591,3],[2.588,2.967,2.586,2.933,2.586,2.898],[2.586,2.484,2.838,2.13,3.196,1.978],[3.24,1.964],[3.172,1.923],[2.652,1.572,2.309,0.976,2.309,0.3],[2.309,0.233,2.313,0.166,2.319,0.1],[]],
  // s12 phone screen: rounded body with the notch cut out of the top edge
  screen12: [[0.381,0],[1.041,0],[1.041,0.115],[1.041,0.223,1.127,0.31,1.234,0.31],[2.334,0.31],[2.441,0.31,2.527,0.223,2.527,0.115],[2.527,0],[3.204,0],[3.414,0,3.585,0.17,3.585,0.381],[3.585,6.379],[0,6.379],[0,0.381],[0,0.17,0.17,0,0.381,0],[]],
  hill13: [[2.563,0],[4.061,0,5.382,0.759,6.162,1.913],[6.305,2.149],[6.325,2.134],[6.683,1.892,7.114,1.751,7.578,1.751],[8.738,1.751,9.692,2.633,9.807,3.763],[9.813,3.874],[0,3.874],[0,0.843],[0.137,0.741],[0.829,0.273,1.664,0,2.563,0],[]],
  corner13: [[0.932,0],[1.06,0,1.183,0.026,1.295,0.073],[1.43,0.147],[1.43,1.423],[0.143,1.423],[0.073,1.295],[0.026,1.183,0,1.06,0,0.932],[0,0.417,0.417,0,0.932,0],[]],
  top13:  [[0,0],[2.068,0],[2.068,1.026],[2.031,1.044],[1.853,1.119,1.657,1.161,1.451,1.161],[0.783,1.161,0.218,0.721,0.029,0.115],[]],
  hill14: [[7.416,5.579],[7.751,5.579,8.039,5.783,8.162,6.073],[8.2,6.195],[8.443,6.208],[9.313,6.296,10.076,6.748,10.577,7.41],[10.639,7.5],[5.617,7.5],[5.627,7.484],[5.841,7.183,6.108,6.923,6.414,6.716],[6.632,6.583],[6.623,6.551],[6.612,6.499,6.606,6.444,6.606,6.388],[6.606,5.941,6.969,5.579,7.416,5.579],[],[0,0],[4.881,0],[4.916,0.045],[5.347,0.635,5.657,1.32,5.808,2.062],[5.817,2.108],[5.819,2.059],[5.883,1.431,6.413,0.941,7.058,0.941],[7.746,0.941,8.304,1.498,8.304,2.186],[8.304,2.874,7.746,3.432,7.058,3.432],[6.542,3.432,6.1,3.118,5.911,2.671],[5.89,2.615],[5.908,2.844],[5.912,2.933,5.914,3.022,5.914,3.112],[5.914,4.911,5.003,6.496,3.617,7.433],[3.512,7.5],[0,7.5],[]],
  cloud14: [[4.507,0],[8.901,0],[8.92,0.121],[8.927,0.196,8.931,0.272,8.931,0.349],[8.931,1.579,7.934,2.576,6.704,2.576],[6.627,2.576,6.551,2.572,6.476,2.565],[6.439,2.559],[6.56,2.669],[7.16,3.269,7.531,4.098,7.531,5.014],[7.531,5.472,7.438,5.908,7.271,6.304],[7.192,6.468],[7.229,6.454],[7.352,6.416,7.483,6.395,7.619,6.395],[8.253,6.395,8.782,6.845,8.904,7.443],[8.913,7.5],[0,7.5],[0,5.216],[0.001,5.215],[0.12,5.191,0.243,5.178,0.369,5.178],[0.495,5.178,0.619,5.191,0.738,5.215],[0.91,5.259],[0.904,5.184],[0.901,5.128,0.9,5.071,0.9,5.014],[0.9,3.984,1.369,3.063,2.106,2.455],[2.332,2.287],[2.322,2.284],[2.084,2.183,1.917,1.947,1.917,1.672],[1.917,1.305,2.214,1.007,2.581,1.007],[2.948,1.007,3.245,1.305,3.245,1.672],[3.245,1.717,3.24,1.762,3.232,1.805],[3.217,1.851],[3.229,1.847],[3.541,1.75,3.872,1.698,4.215,1.698],[4.444,1.698,4.668,1.721,4.884,1.765],[5.016,1.799],[4.986,1.766],[4.668,1.381,4.477,0.887,4.477,0.349],[4.477,0.272,4.481,0.196,4.489,0.121],[]],
  top14:  [[0,0],[2.058,0],[2.058,0.796],[2.041,0.806],[1.838,0.917,1.605,0.979,1.358,0.979],[0.765,0.979,0.256,0.619,0.038,0.105],[]],
  phoneIcon: [[0.138,0.137],[0.119,0.156,0.095,0.175,0.086,0.166],[0.071,0.151,0.062,0.142,0.033,0.166],[0,0.189,0.024,0.208,0.038,0.218],[0.052,0.236,0.11,0.222,0.167,0.166],[0.224,0.109,0.238,0.052,0.224,0.033],[0.21,0.019,0.195,0,0.171,0.028],[0.148,0.057,0.157,0.066,0.171,0.081],[0.181,0.089,0.162,0.113,0.138,0.137],[]],
  mailIcon: [[0.01,0.015],[0.019,0.019,0.109,0.067,0.109,0.067],[0.114,0.072,0.119,0.072,0.124,0.072],[0.129,0.072,0.133,0.072,0.133,0.067],[0.138,0.067,0.229,0.019,0.234,0.015],[0.243,0.01,0.248,0,0.238,0],[0.01,0],[0,0,0.005,0.01,0.01,0.015],[],[0.238,0.043],[0.234,0.043,0.138,0.091,0.133,0.096],[0.129,0.096,0.124,0.096,0.119,0.096],[0.104,0.091,0.015,0.043,0.01,0.043],[0.005,0.039,0.005,0.043,0.005,0.043],[0.005,0.143],[0.005,0.148,0.01,0.153,0.019,0.153],[0.229,0.153],[0.238,0.153,0.243,0.148,0.243,0.143],[0.243,0.047,0.243,0.043,0.243,0.043],[0.243,0.043,0.243,0.039,0.238,0.043],[]],
  cloud15: [[3.782,0],[7.822,0],[7.822,7.5],[0,7.5],[0.04,7.344],[0.268,6.611,0.952,6.078,1.761,6.078],[2.01,6.078,2.247,6.129,2.462,6.22],[2.593,6.283],[2.638,5.944],[2.924,4.343,4.167,3.074,5.754,2.749],[5.836,2.737],[5.81,2.697],[5.689,2.495,5.591,2.278,5.519,2.048],[5.474,1.87],[5.415,1.886],[5.322,1.905,5.225,1.915,5.126,1.915],[4.83,1.915,4.555,1.825,4.327,1.67],[4.308,1.656],[4.306,1.66],[4.233,1.768,4.109,1.839,3.969,1.839],[3.745,1.839,3.563,1.657,3.563,1.432],[3.563,1.264,3.665,1.12,3.811,1.058],[3.816,1.056],[3.808,1.041],[3.736,0.87,3.696,0.682,3.696,0.484],[3.696,0.336,3.719,0.193,3.76,0.059],[]],
  hill15: [[0.052,0],[1.253,0,2.255,0.852,2.487,1.985],[2.508,2.125],[0,2.125],[0,0.003],[]],
  bump15: [[0,0],[2.224,0],[2.224,0.614,1.726,1.112,1.112,1.112],[0.498,1.112,0,0.614,0,0],[]],
  // rating widgets on s12
  halfStarL: [[0.066,0],[0.071,0.015],[0.071,0.104],[0.066,0.101],[0.025,0.132],[0.041,0.082],[0,0.05],[0.05,0.05],[0.066,0],[]],
  halfStarR: [[0,0],[0.011,0.035],[0.061,0.035],[0.021,0.067],[0.036,0.117],[0,0.089],[0,0],[]],
  star: [[0.066,0],[0.071,0.015],[0.082,0.05],[0.132,0.05],[0.091,0.082],[0.107,0.132],[0.071,0.104],[0.066,0.101],[0.025,0.132],[0.041,0.082],[0,0.05],[0.05,0.05],[]],
};

/* --------------------------------------------------------- reused copy */

const LOREM_LONG = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin et enim venenatis, mollis mauris sit amet, dictum quam. Praesent varius ipsum sed suscipit sodales. ';
const LOREM_MED = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin et enim venenatis, mollis mauris sit amet, dictum quam. ';
const LOREM_CUT = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin et enim venenatis, mollis mauris sit amet, dictum quam. Praesent varius ipsum sed suscipit';

/* ================================================================ slides */

function slide01(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.pink };
  photo(s, { x: 0.002, y: 2.158, w: 13.33, h: 5.342, points: P.cloud1, label: false });
  txt(s, 'Awareness Can', { x: 2.821, y: 1.42, w: 7.692, h: 1.313, size: 72, color: C.white, align: 'center' });
  txt(s, 'Saves Lives', { x: 2.821, y: 2.411, w: 7.692, h: 1.313, size: 72, color: C.white, align: 'center' });
  brand(s, C.white, C.cream);
  ring(s, 10.233, -1.241, 2.482, C.cream);
  ellipse(s, 11.886, 0.643, 0.706, 0.706, C.orange);
  ellipse(s, 11.474, 3.525, 0.457, 0.451, C.cream);
  roundRect(s, { x: 3.578, y: 6.167, w: 6.177, h: 0.615, fill: C.white, r: 0.3075 });
  txt(s, 'PLACEHOLDER',
    { x: 3.557, y: 6.258, w: 6.177, h: 0.369, size: 12, color: C.body, align: 'center', ls: 1.5 });
}

function slide02(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.blush };
  freeform(s, P.hill2, { x: 0.002, y: 2.92, w: 8.203, h: 4.58, fill: C.pink });
  photo(s, { x: 0.618, y: 0.706, w: 6.089, h: 6.223, points: P.cloud2 });
  freeform(s, P.wedge2, { x: 0.002, y: 0, w: 2.668, h: 1.366, fill: C.pink });
  brand(s, C.white, C.cream);
  heading(s, 'Why Awareness Changes Everything', { x: 7.608, y: 1.65, w: 4.888, h: 1.178 });
  para(s, LOREM_LONG, { x: 8.644, y: 3.115, w: 4.071, h: 0.903 });
  rocketIcon(s, 7.707, 3.297, 0.539);
  freeform(s, P.hump2, { x: 7.336, y: 6.235, w: 2.656, h: 1.265, fill: C.pink });
  ellipse(s, 8.363, 5.965, 0.539, 0.539, C.orange);
  ring(s, 9.786, 5.28, 0.995, C.white, 4.5);
  ellipse(s, 6.821, 6.427, 0.292, 0.292, C.cream);
}

function slide03(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.blush };
  freeform(s, P.wave3, { x: 0.002, y: 5.32, w: 13.33, h: 2.18, fill: C.pink });
  [0.658, 3.706, 6.796, 9.865].forEach(x =>
    photo(s, { x, y: 3.319, w: 2.717, h: 2.717, shape: 'ellipse' }));
  brand(s, C.ink, C.pink);
  heading(s, 'Knowing the Risks Saves Lives', { x: 0.881, y: 1.618, w: 4.503, h: 1.178 });
  para(s, LOREM_LONG + 'Nulla facilisi. Ut sit amet nulla massa. ', { x: 6.508, y: 1.851, w: 4.923, h: 0.903 });
  const labels = [
    ['Age and Gender', 0.544, 0.404],
    ['Genetics Role', 3.621, 0.404],
    ['Lifestyle Impact', 6.712, 0.707],
    ['key risk factor', 9.78, 0.404],
  ];
  labels.forEach(([text, x, h]) =>
    txt(s, text, { x, y: 6.309, w: 2.888, h, color: C.white, align: 'center' }));
  [3.432, 6.524, 9.731].forEach(x => vline(s, x, 6.309, 0.493, C.blush));
}

function slide04(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.blush };
  ellipse(s, 0.477, 6.284, 0.684, 0.684, C.orange);
  ring(s, 11.868, 6.023, 2.156, C.orange);
  // three "cancer type" cards; the leftmost one is the highlighted pink card
  const cards = [
    { x: 0.689, card: C.pink, dot: C.cream, text: C.white, rule: C.cream,
      title: 'Triple-Negative', sub: '(Lacks of Common Receptors)', tw: 2.888, th: 0.404 },
    { x: 4.84, card: C.cream, dot: C.pink, text: C.body, rule: C.pink,
      title: 'Invasive Lobular Carcinoma (ILC)', tw: 2.888, th: 0.707 },
    { x: 9.018, card: C.cream, dot: C.orange, text: C.body, rule: C.blush,
      title: 'Invasive Ductal Carcinoma (IDC)', tw: 2.888, th: 0.707 },
  ];
  cards.forEach(c => roundRect(s, { x: c.x, y: 2.762, w: 3.627, h: 3.973, fill: c.card }));
  brand(s, C.ink, C.pink);
  heading(s, 'The Most Common Cancer Types', { x: 1.344, y: 1.316, w: 10.646, h: 0.64, align: 'center' });
  cards.forEach(c => ellipse(s, c.x + 1.389, 2.337, 0.849, 0.849, c.dot));
  cards.forEach(c => {
    const titleColor = c.text === C.white ? C.white : C.dark;
    txt(s, c.title, { x: c.x + 0.382, y: 3.667 + (c.sub ? 0.021 : 0), w: c.tw, h: c.th, color: titleColor, align: 'center' });
    if (c.sub) para(s, c.sub, { x: c.x + 0.277, y: 4.008, w: 3.099, h: 0.348, color: c.text, align: 'center' });
    para(s, LOREM_MED, { x: c.x + 0.277, y: 4.494, w: 3.099, h: 0.903, color: c.text, align: 'center' });
    slide04Rule(s, c);
    para(s, 'Patient', { x: c.x + 0.42, y: 6.023, w: 1.208, h: 0.348, color: c.text });
    txt(s, '1.456.789 Infected', { x: c.x + 1.365, y: 6.057, w: 2.35, h: 0.337, size: 14, color: c.text });
  });
  leafIcon(s, 2.352, 2.596, 0.264, 0.369, C.orange);
  bulbIcon(s, 6.54, 2.572, 0.229, 0.383, C.cream);
  targetIcon(s, 10.651, 2.576, 0.361, 0.358, C.cream);
}

function slide04Rule(s, c) {
  s.addShape('line', { x: c.x - 0.048, y: 5.781, w: 3.627, h: 0, line: { color: c.rule, width: 0.5 } });
}

function slide05(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.blush };
  photo(s, { x: 0.002, y: 4.824, w: 5.07, h: 2.676, points: P.cloud5 });
  brand(s, C.ink, C.pink);
  heading(s, "Symptoms You Shouldn't Ignore", { x: 0.881, y: 1.704, w: 4.309, h: 1.178 });
  para(s, LOREM_CUT, { x: 0.881, y: 3.185, w: 4.246, h: 0.903 });
  // three horizontal cards, authored rotated 270 degrees in the source deck
  const rows = [
    { y: -1.199, card: C.pink, dot: C.cream, letter: 'A.', lc: C.pink, title: 'Nipple Discharge',
      tc: C.white, bc: C.white, bw: 4.487, ty: 1.19 },
    { y: 0.911, card: C.cream, dot: C.pink, letter: 'B.', lc: C.cream, title: 'Unusual Swelling',
      tc: C.ink, bc: C.body, bw: 4.846, ty: 3.316 },
    { y: 3.02, card: C.cream, dot: C.orange, letter: 'C.', lc: C.cream, title: 'Lump or Discharge',
      tc: C.ink, bc: C.body, bw: 4.487, ty: 5.429 },
  ];
  rows.forEach((r, i) => {
    roundRect(s, { x: 8.875, y: r.y, w: 1.572, h: 5.806, fill: r.card, r: 0.2504, rotate: 270 });
    ellipse(s, 6.262, 1.28 + i * 2.109, 0.849, 0.849, r.dot);
    txt(s, r.letter, { x: 6.225, y: 1.481 + i * 2.1, w: 0.952, h: 0.404, color: r.lc, align: 'center' });
    txt(s, r.title, { x: 7.607, y: r.ty, w: 2.555, h: 0.404, color: r.tc });
    para(s, LOREM_MED, { x: 7.607, y: r.ty + 0.356, w: r.bw, h: 0.625, color: r.bc });
  });
}

function slide06(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.blush };
  freeform(s, P.hill6, { x: 0.002, y: 3.041, w: 7.583, h: 4.459, fill: C.pink });
  ellipse(s, 6.607, 5.687, 0.427, 0.427, C.cream);
  ring(s, 6.315, 6.524, 1.605, C.orange);
  photo(s, { x: 0.762, y: 0.935, w: 5.905, h: 6.075, points: P.cloud6 });
  brand(s, C.ink, C.pink);
  heading(s, 'Free Screenings \nand Clinics Available', { x: 7.989, y: 1.618, w: 5.075, h: 1.178 });
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing \nelit. Proin et enim venenatis, mollis mauris sit amet, dictum quam. Praesent varius ipsum',
    { x: 7.989, y: 3.032, w: 4.376, h: 0.903 });
  txt(s, 'Annual Checkups Offered', { x: 8.302, y: 4.624, w: 3.963, h: 0.37, size: 16 });
  roundRect(s, { x: 7.519, y: 5.104, w: 1.135, h: 0.05, fill: C.orange, r: 0.025, rotate: 270 });
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing \nelit. Proin et enim venenatis, mollis mauris sit amet',
    { x: 8.302, y: 4.989, w: 4.022, h: 0.625 });
}

function slide07(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.blush };
  photo(s, { x: 0.705, y: 0, w: 12.626, h: 4.942, points: P.cloud7, label: false });
  brand(s, C.ink, C.pink);
  heading(s, 'We Offer Hope and Guidance', { x: 0.853, y: 5.587, w: 3.938, h: 1.178 });
  para(s, LOREM_CUT, { x: 4.9, y: 5.724, w: 4.246, h: 0.903 });
  txt(s, 'Community resources', { x: 9.683, y: 5.669, w: 3.055, h: 0.37, size: 16 });
  para(s, 'Survivor-led healing groups.', { x: 10.033, y: 6.038, w: 2.96, h: 0.348 });
  para(s, ' Informative talks at various places', { x: 10.033, y: 6.33, w: 2.96, h: 0.348 });
  ellipse(s, 9.798, 6.177, 0.126, 0.126, C.pink);
  ellipse(s, 9.798, 6.484, 0.126, 0.126, C.orange);
  ellipse(s, 3.448, 4.297, 0.792, 0.792, C.cream);
  ring(s, 5.293, -0.364, 1.322, C.pink);
  ellipse(s, 0.75, 2.363, 0.342, 0.342, C.orange);
}

function slide08(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.blush };
  freeform(s, P.wave8, { x: -0.25, y: 1.449, w: 13.7, h: 6.771, fill: C.pink, rotate: 353.7 });
  roundRect(s, { x: 0.618, y: 2.762, w: 6.197, h: 3.973, fill: C.cream });
  roundRect(s, { x: 7.388, y: 2.762, w: 5.327, h: 3.973, fill: C.cream });
  photo(s, { x: 0.881, y: 2.994, w: 2.692, h: 3.486, r: 0.19 });
  photo(s, { x: 7.651, y: 2.994, w: 2.198, h: 3.486, r: 0.19 });
  brand(s, C.ink, C.pink);
  heading(s, 'Treatment Paths You Should Know', { x: 3.921, y: 1.248, w: 5.492, h: 1.178, align: 'center' });
  const cols = [
    { badge: 6.043, dot: C.pink, tx: 3.953, title: 'Surgery and Chemotherapy', tw: 3.055,
      body: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin et enim venenatis, mollis', bx: 3.949 },
    { badge: 11.926, dot: C.orange, tx: 10.233, title: 'Hormonal Therapy', tw: 2.017,
      body: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. \nProin et enim venenatis', bx: 10.229 },
  ];
  cols.forEach(c => {
    ellipse(s, c.badge, 2.994, 0.527, 0.527, c.dot);
    txt(s, c.title, { x: c.tx, y: 4.663, w: c.tw, h: 0.64, size: 16 });
    para(s, c.body, { x: c.bx, y: 5.348, w: 2.438, h: 0.903 });
  });
  ring(s, 10.511, -1.155, 2.156, C.orange);
  ellipse(s, 11.926, 0.636, 0.527, 0.527, C.pink);
  docIcon(s, 6.227, 3.168, 0.196, 0.2, C.white);
  lockIcon(s, 12.107, 3.167, 0.164, 0.2, C.white);
}

function slide09(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.blush };
  ring(s, 5.026, 4.833, 1.426, C.orange);
  freeform(s, P.hill9, { x: 9.704, y: 3.165, w: 3.627, h: 4.335, fill: C.pink });
  roundRect(s, { x: 5.859, y: 0.822, w: 2.955, h: 5.951, fill: C.pink, r: 0.3074 });
  roundRect(s, { x: 9.352, y: 0.822, w: 2.955, h: 5.951, fill: C.cream, r: 0.3074 });
  photo(s, { x: 6.122, y: 1.054, w: 2.43, h: 3.265, r: 0.2 });
  photo(s, { x: 9.614, y: 1.054, w: 2.43, h: 3.265, r: 0.2 });
  brand(s, C.ink, C.pink);
  heading(s, 'Meet Our Advocacy Leaders', { x: 0.881, y: 1.686, w: 4.377, h: 1.178 });
  para(s, LOREM_CUT, { x: 0.881, y: 3.11, w: 4.298, h: 0.903, color: C.black });
  const people = [
    { x: 6.168, name: 'Fionateza Joe', role: 'Survivor of Breast Cancer', nc: C.white, c: C.white },
    { x: 9.66, name: 'Margareta Amber', role: 'Specialized Doctor', nc: C.ink, c: C.body },
  ];
  people.forEach(p => {
    txt(s, p.name, { x: p.x, y: 4.642, w: 2.555, h: 0.404, color: p.nc });
    para(s, p.role, { x: p.x + 0.018, y: 4.945, w: 2.438, h: 0.37, size: 12, color: p.c });
    para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. \nProin et enim venenatis',
      { x: p.x + 0.018, y: 5.479, w: 2.801, h: 0.903, color: p.c });
  });
  txt(s, '10+ Years', { x: 0.853, y: 5.046, w: 3.938, h: 0.774, size: 40, color: C.pink });
  txt(s, 'Experience in Field', { x: 0.891, y: 5.816, w: 2.555, h: 0.404 });
}

function slide10(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.blush };
  freeform(s, P.arc10, { x: -0.238, y: 6.169, w: 1.722, h: 1.553, line: C.orange });
  freeform(s, P.hill10, { x: 10.376, y: 5.248, w: 2.956, h: 2.252, fill: C.pink });
  roundRect(s, { x: 5.978, y: 2.344, w: 6.65, h: 4.512, fill: C.cream, r: 0.4694 });
  roundRect(s, { x: 0.705, y: 2.344, w: 4.711, h: 4.512, fill: C.pink, r: 0.4694 });
  photo(s, { x: 0.951, y: 2.594, w: 4.204, h: 3.25, r: 0.27 });
  photo(s, { x: 6.244, y: 2.604, w: 6.086, h: 3.25, r: 0.27 });
  brand(s, C.ink, C.pink);
  heading(s, 'Campaigns That Made Impact', { x: 2.722, y: 1.346, w: 7.89, h: 0.64, align: 'center' });
  txt(s, 'Walk for Awareness', { x: 1.146, y: 6.121, w: 3.889, h: 0.404, color: C.white });
  para(s, '27/10', { x: 4.391, y: 6.149, w: 1.019, h: 0.348, color: C.white });
  txt(s, 'Ribbon Art Installation', { x: 6.582, y: 6.121, w: 4.366, h: 0.404 });
  para(s, '31 December 2029', { x: 10.625, y: 6.149, w: 2.801, h: 0.348 });
}

function slide11(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.blush };
  freeform(s, P.top11, { x: 7.485, y: 0, w: 3.007, h: 2.047, fill: C.pink });
  freeform(s, P.side11, { x: 10.281, y: 3.754, w: 3.051, h: 3.746, fill: C.pink });
  photo(s, { x: 8.917, y: 4.969, w: 3.799, h: 2.531, points: P.tile11d });
  photo(s, { x: 4.729, y: 3.135, w: 3.799, h: 4.365, points: P.tile11c });
  photo(s, { x: 8.917, y: 0, w: 3.799, h: 4.448, points: P.tile11b });
  photo(s, { x: 4.729, y: 0, w: 3.799, h: 2.615, points: P.tile11a });
  ring(s, 3.626, -0.728, 1.718, C.orange);
  brand(s, C.ink, C.pink, { x: 0.618, y: 0.501, d: 0.205, ty: 0.458 });
  heading(s, 'Real Lives. \nReal Courage.', { x: 0.881, y: 1.644, w: 5.492, h: 1.178 });
  para(s, LOREM_MED, { x: 0.881, y: 3.187, w: 3.286, h: 0.903 });
  const pills = [
    { x: 4.939, y: 2.53, fill: C.pink, text: 'The Strong Survivor' },
    { x: 9.219, y: 4.353, fill: C.orange, text: 'Awareness Event' },
  ];
  pills.forEach(p => {
    roundRect(s, { x: p.x, y: p.y, w: 3.286, h: 0.691, fill: p.fill, r: 0.3455 });
    txt(s, p.text, { x: p.x + 0.232, y: p.y + 0.126, w: 2.822, h: 0.404, color: C.white, align: 'center' });
  });
  txt(s, '27K+', { x: 0.853, y: 5.046, w: 3.938, h: 0.774, size: 40, color: C.pink });
  txt(s, 'Breast Cancer Survivors', { x: 0.891, y: 5.816, w: 2.555, h: 0.707 });
  ring(s, -1.103, 6.775, 2.207, C.pink);
  ellipse(s, 3.214, -0.394, 0.599, 0.599, C.pink);
}

function slide12(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.blush };
  freeform(s, P.arc12, { x: 3.732, y: 5.78, w: 4.489, h: 1.976, line: C.pink });
  freeform(s, P.hill12, { x: 9.618, y: 0, w: 3.714, h: 7.5, fill: C.pink });
  // phone mock-up: light chassis, dark bezel, grey screen with a notch
  roundRect(s, { x: 7.53, y: 0.925, w: 4.003, h: 6.575, fill: 'C8CACB', r: 0.46 });
  roundRect(s, { x: 7.6, y: 0.99, w: 3.86, h: 6.51, fill: '15171A', r: 0.43 });
  photo(s, { x: 7.755, y: 1.121, w: 3.585, h: 6.379, points: P.screen12 });
  brand(s, C.ink, C.pink);
  heading(s, 'Survivor Testimonials ', { x: 0.881, y: 1.8, w: 4.244, h: 1.178 });
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin et enim venenatis, mollis',
    { x: 4.585, y: 1.946, w: 2.498, h: 0.903 });
  const quotes = [
    { x: 0.618, fill: C.pink, name: 'Mia Thompson', nc: C.white, star: C.cream, score: '4,9', half: true, tc: C.white },
    { x: 4.868, fill: C.cream, name: 'Alexander Davidson', nc: C.ink, star: C.pink, score: '5.0', half: false, tc: C.body },
  ];
  quotes.forEach(q => {
    roundRect(s, { x: q.x, y: 3.816, w: 3.861, h: 2.515, fill: q.fill });
    txt(s, q.name, { x: q.x + 0.431, y: 4.277, w: 3.055, h: 0.37, size: 16, color: q.nc });
    stars(s, q.x + 0.526, 4.675, q.star, q.half);
    para(s, q.score, { x: q.x + 1.375, y: 4.55, w: 0.618, h: 0.348, size: 10.5, color: q.nc });
    para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin et enim venenatis, mollis mauris sit amet, dictum',
      { x: q.x + 0.426, y: 4.945, w: 3.193, h: 0.903, color: q.tc });
  });
}

/** Five 0.132in stars; the last one is split in half when `half` is set. */
function stars(s, x, y, color, half) {
  for (let i = 0; i < 4; i++) {
    s.addShape('star5', { x: x + i * 0.1615, y, w: 0.132, h: 0.132, fill: { color }, line: NONE });
  }
  const lastX = x + 4 * 0.1615;
  if (half) {
    freeform(s, P.halfStarL, { x: lastX - 0.0015, y, w: 0.071, h: 0.132, fill: color });
    freeform(s, P.halfStarR, { x: lastX + 0.0695, y: y + 0.014, w: 0.061, h: 0.117, fill: 'BFBFBF' });
  } else {
    freeform(s, P.star, { x: lastX, y, w: 0.132, h: 0.132, fill: color });
  }
}

function slide13(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.blush };
  freeform(s, P.hill13, { x: 0.002, y: 3.626, w: 9.813, h: 3.874, fill: C.pink });
  // laptop mock-up: dark lid, grey screen, silver base with a trackpad notch
  roundRect(s, { x: -0.6, y: 1.333, w: 7.46, h: 5.06, fill: '060606', r: 0.15 });
  photo(s, { x: 0.002, y: 1.669, w: 6.618, h: 4.401 });
  roundRect(s, { x: -0.6, y: 6.39, w: 8.3, h: 0.2, fill: 'EFF2F5', r: 0.06 });
  roundRect(s, { x: 2.44, y: 6.4, w: 1.3, h: 0.055, fill: 'A6A7AB', r: 0.027 });
  roundRect(s, { x: -0.6, y: 6.55, w: 8.3, h: 0.06, fill: '8C8B90', r: 0.03 });
  brand(s, C.ink, C.pink);
  heading(s, 'Spread the Message Visually', { x: 8.139, y: 1.573, w: 4.309, h: 1.178 });
  para(s, LOREM_LONG, { x: 8.139, y: 3.021, w: 4.309, h: 0.903 });
  targetIcon(s, 8.08, 4.466, 0.409, 0.406, C.orange);
  vline(s, 8.767, 4.479, 0.484, C.pink);
  txt(s, 'Long-Term Vision', { x: 9.093, y: 4.413, w: 3.233, h: 0.337, size: 14, color: C.dark });
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Ut sed accumsan nunc. ',
    { x: 9.059, y: 4.749, w: 3.233, h: 0.625 });
  freeform(s, P.corner13, { x: 11.902, y: 6.077, w: 1.43, h: 1.423, fill: C.pink });
  ring(s, 8.87, 6.077, 1.194, C.cream);
  freeform(s, P.top13, { x: 11.264, y: 0, w: 2.068, h: 1.161, fill: C.pink });
  ellipse(s, 11.227, 0.415, 0.497, 0.497, C.pink);
}

function slide14(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.blush };
  freeform(s, P.hill14, { x: 0.002, y: 0, w: 10.639, h: 7.5, fill: C.pink });
  ring(s, 0.278, 2.333, 2.321, C.cream);
  ellipse(s, 0.742, 2.017, 0.633, 0.633, C.orange);
  photo(s, { x: 0.002, y: 0, w: 8.931, h: 7.5, points: P.cloud14, label: false });
  heading(s, 'Contact Us for Further Information', { x: 9.568, y: 1.881, w: 2.885, h: 1.717 });
  brand(s, C.white, C.cream);
  txt(s, 'Phone', { x: 9.695, y: 4.297, w: 1.755, h: 0.337, size: 14, color: C.black });
  const contacts = [
    { label: 'Customer Support', value: '+1 555-123-4567', lw: 2.065, vw: 1.656, y: 4.665, icon: P.phoneIcon, ix: 9.805, iy: 4.755, iw: 0.239, ih: 0.237 },
    { label: 'Company Email', value: 'name@company.mail', lw: 2.241, vw: 2.694, y: 5.28, icon: P.mailIcon, ix: 9.8, iy: 5.375, iw: 0.248, ih: 0.153 },
  ];
  contacts.forEach(c => {
    txt(s, c.label, { x: 10.196, y: c.y, w: c.lw, h: 0.278, face: BODY, size: 10.5, color: C.black });
    txt(s, c.value, { x: 10.196, y: c.y + 0.241, w: c.vw, h: 0.303, face: BODY, size: 12, color: C.black });
    freeform(s, c.icon, { x: c.ix, y: c.iy, w: c.iw, h: c.ih, fill: C.pink });
  });
  ring(s, 7.886, 4.637, 0.604, C.orange, 4.5);
  ellipse(s, 7.705, 2.877, 0.391, 0.391, C.cream);
  freeform(s, P.top14, { x: 11.273, y: 0, w: 2.058, h: 0.979, fill: C.pink });
}

function slide15(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.pink };
  photo(s, { x: 5.51, y: 0, w: 7.822, h: 7.5, points: P.cloud15, label: false });
  brand(s, C.white, C.cream);
  txt(s, 'Keep Talking.', { x: 0.881, y: 2.23, w: 7.692, h: 1.111, size: 60, color: C.white });
  txt(s, 'Keep Checking.', { x: 0.881, y: 3.247, w: 7.692, h: 1.111, size: 60, color: C.white });
  txt(s, 'Awareness isn\u2019t just a word\u2014it\u2019s a lifeline.',
    { x: 0.944, y: 4.461, w: 5.722, h: 0.37, size: 16, color: C.white });
  freeform(s, P.hill15, { x: 0.002, y: 5.375, w: 2.508, h: 2.125, fill: C.cream });
  freeform(s, P.bump15, { x: 5.213, y: 0, w: 2.224, h: 1.112, fill: C.cream });
  ellipse(s, 6.881, 0.33, 1.112, 1.112, C.orange);
  ring(s, 2.088, 6.89, 1.221, C.blush);
}

/* ----------------------------------------------------------------- icons */
/* Small pictograms drawn from primitives instead of embedding the SVGs.    */

/** Rocket flying to the upper right: hull, porthole, two fins and a flame. */
function rocketIcon(s, x, y, d) {
  const pink = { color: C.pink };
  s.addShape('teardrop', { x: x + 0.22 * d, y: y + 0.02 * d, w: 0.76 * d, h: 0.76 * d, fill: pink, line: NONE });
  ellipse(s, x + 0.53 * d, y + 0.2 * d, 0.2 * d, 0.2 * d, C.blush);
  s.addShape('rtTriangle', { x: x + 0.05 * d, y: y + 0.3 * d, w: 0.24 * d, h: 0.24 * d, fill: pink, line: NONE, rotate: 90 });
  s.addShape('rtTriangle', { x: x + 0.7 * d, y: y + 0.72 * d, w: 0.24 * d, h: 0.24 * d, fill: pink, line: NONE, rotate: 270 });
  s.addShape('pie', { x: x + 0.03 * d, y: y + 0.64 * d, w: 0.33 * d, h: 0.33 * d, fill: pink, line: NONE, angleRange: [90, 180] });
}

/** Leaf: teardrop tipped so the point faces up-right, with a pale midrib. */
function leafIcon(s, x, y, w, h, color) {
  s.addShape('teardrop', { x, y, w, h, fill: { color }, line: NONE, rotate: 335 });
  s.addShape('line', { x: x + 0.2 * w, y: y + 0.88 * h, w: 0.55 * w, h: -0.58 * h, line: { color: C.cream, width: 1.2 } });
}

/** Light bulb: round glass over a short screw base. */
function bulbIcon(s, x, y, w, h, color) {
  ellipse(s, x, y, w, w, color);
  s.addShape('rect', { x: x + 0.26 * w, y: y + 0.68 * w, w: 0.48 * w, h: h - 0.76 * w, fill: { color }, line: NONE });
  s.addShape('rect', { x: x + 0.2 * w, y: y + h - 0.09 * w, w: 0.6 * w, h: 0.09 * w, fill: { color }, line: NONE });
}

/** Dart hitting a ring-shaped target. */
function targetIcon(s, x, y, w, h, color) {
  s.addShape('donut', { x, y: y + 0.28 * h, w: w * 0.92, h: h * 0.72, fill: { color }, line: NONE });
  s.addShape('line', { x: x + 0.24 * w, y: y + 0.74 * h, w: 0.5 * w, h: -0.56 * h, line: { color, width: 2.4 } });
  s.addShape('rtTriangle', { x: x + 0.62 * w, y: y - 0.04 * h, w: 0.38 * w, h: 0.32 * h, fill: { color }, line: NONE, rotate: 200 });
}

/** Document with a magnifying glass. */
function docIcon(s, x, y, w, h, color) {
  s.addShape('rect', { x, y, w: w * 0.62, h: h * 0.72, fill: { color }, line: NONE });
  s.addShape('donut', { x: x + w * 0.4, y: y + h * 0.4, w: w * 0.5, h: h * 0.5, fill: { color }, line: NONE });
  s.addShape('line', { x: x + w * 0.82, y: y + h * 0.82, w: w * 0.18, h: h * 0.18, line: { color, width: 1.4 } });
}

/** Padlock: shackle arc over a solid body. */
function lockIcon(s, x, y, w, h, color) {
  s.addShape('blockArc', {
    x: x + 0.18 * w, y, w: 0.64 * w, h: 0.6 * h,
    fill: { color }, line: NONE, angleRange: [180, 360], arcThicknessRatio: 0.32,
  });
  s.addShape('rect', { x, y: y + 0.44 * h, w, h: 0.56 * h, fill: { color }, line: NONE });
  ellipse(s, x + 0.42 * w, y + 0.62 * h, 0.16 * w, 0.16 * w, C.orange);
}

/* ------------------------------------------------------------------ main */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK', width: 13.333, height: 7.5 });
  pptx.layout = 'DECK';
  pptx.title = 'Awareness Can Saves Lives';
  pptx.author = 'Restnaire';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
    slide09, slide10, slide11, slide12, slide13, slide14, slide15]
    .forEach(fn => fn(pptx));

  return pptx.writeFile({
    fileName: path.join(__dirname, '0efdb537-fe25-4eca-812f-980a1626a2cb_grok_final.pptx'),
  });
}

build().then(f => console.log('wrote ' + f)).catch(e => { console.error(e); process.exit(1); });
