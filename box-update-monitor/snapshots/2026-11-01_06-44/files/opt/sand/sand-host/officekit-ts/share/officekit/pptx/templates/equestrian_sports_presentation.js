/*
 * "Equestrian" — 25-slide deck rebuilt with pptxgenjs.
 *
 * Layout units are inches on a 13.333 x 7.5 stage. Every slide has a builder
 * function at the bottom of the file; shared decoration (plus grids, striped
 * bars, chevrons, arches, the horse silhouette) lives in the motif helpers.
 * Raster photos in the original are replaced by flat placeholder rectangles.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette --
const C = {
  white: 'FFFFFF',
  black: '000000',
  ink: '262626',            // tx1 lum 85%
  charcoal: '1C1911',       // near-black used for full-bleed panels
  brick: '9B2D1F',          // accent2 - primary red
  brickDark: '742217',      // accent2 @ 75%
  brickDeep: '4E1710',      // accent2 @ 50%
  orange: 'D34817',         // accent1
  orangeDark: '9E3611',     // accent1 shaded
  orangeDeep: '69240C',
  khaki: 'A28E6A',          // accent3
  khakiMid: '7C6B4D',       // accent3 @ 85%
  olive: '534733',          // accent3 @ 50%
  clay: '956251',           // accent4
  clayDark: '704A3D',
  cocoa: '4A3128',          // accent4 @ 50%
  rose: 'DF6C5D',
  grey: '918485',           // accent5
  grey40: '404040',
  grey35: '595959',
  greyMid: 'A6A6A6',
  greyLt: 'BFBFBF',
  greySoft: 'D9D9D9',
  greyPale: 'F2F2F2',
  gold: '78490C',
  taupe: '855D5D',
  photo: 'FCFBFA',        // stand-in tone for the deck's empty photo frames
};

const HEAD = 'Libre Baskerville';   // theme major font
const BODY = 'Raleway';             // theme minor font
const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

// ------------------------------------------------------------ boilerplate --
const LOREM = {
  short: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nunc ullamcorper sit amet orci et consequat. ',
  med: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nunc ullamcorper sit amet orci et consequat. Morbi semper eros vitae tincidunt porta. Mauris euismod. Cum sociis natoque penatibus.',
  tail: ' Lorem ipsum dolor sit amet, consectetur adipiscing elit, ullamco\u00a0',
  tail2: 'consectetur adipiscing elit. Integer vitae justo',
  mauris: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Mauris mattis libero id felis rhoncus mattis. Orci varius natoque penatibus et magnis.',
  maurisLong: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Mauris mattis libero id felis rhoncus mattis. Orci varius natoque penatibus et magnis dis parturient.',
  maurisMontes: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Mauris mattis libero id felis rhoncus mattis. Orci varius natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. ',
  qui: 'Lorem ipsum dolor sit amet. Qui sint neque a velit modi quo numquam. Non exercitationem reiciendis qui consequatur. Pellentesque habitant morbi tristique senectus et netus.',
  ullamco: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, ullamco\u00a0consectetur adipiscing elit. Integer vitae justo ullamcorper',
  tempor: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, veniam minim as tempor incididunt ut labore et dolore magna aliqua. ',
  aenean: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. ',
  aeneanShort: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque.',
  integer: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer. ',
  row: 'Lorem ipsum dolor sit amet, consectetuer adipiscing'
};
const QUOTE = 'The art of riding and training a horse in a manner that develops obedience, flexibility, and balance.';

// --------------------------------------------- vector motif path libraries --
// Each entry is a list of sub-paths; a sub-path is a flat [x0,y0,x1,y1,...]
// array of coordinates normalised to the shape's own 0..1 bounding box.
const HORSE = [[0.993,0.342,0.969,0.23,0.842,0.078,0.842,0.113,0.823,0.087,0.787,0.129,0.806,0.149,0.767,0.147,0.816,0.092,0.815,0.022,0.741,-0.019,0.722,0.06,0.61,0.055,0.545,0.099,0.542,0.173,0.624,0.266,0.608,0.305,0.543,0.272,0.505,0.301,0.405,0.287,0.281,0.399,0.159,0.377,0.137,0.399,0.181,0.383,0.152,0.405,0.213,0.414,0.162,0.423,0.152,0.442,0.18,0.43,0.144,0.459,0.132,0.418,0.142,0.485,0.111,0.502,0.101,0.455,0.106,0.503,0.069,0.543,0.014,0.518,0.054,0.555,0,0.531,0.03,0.57,0.002,0.573,0.011,0.592,0.049,0.585,0.01,0.61,0.123,0.627,0.183,0.573,0.123,0.643,0.195,0.598,0.15,0.645,0.307,0.537,0.262,0.59,0.295,0.589,0.316,0.528,0.317,0.581,0.327,0.539,0.404,0.613,0.39,0.764,0.49,0.806,0.546,0.927,0.563,0.906,0.62,0.985,0.656,0.968,0.531,0.799,0.509,0.612,0.57,0.666,0.623,0.667,0.722,0.823,0.666,0.868,0.629,0.837,0.588,0.857,0.66,0.919,0.747,0.826,0.719,0.679,0.84,0.744,0.891,1.011,0.878,0.749,0.789,0.623,0.852,0.366,0.946,0.365,0.952,0.392],[0.088,0.518,0.115,0.537],[0.227,0.573,0.205,0.593,0.248,0.525],[0.242,0.567,0.264,0.518],[0.264,0.566,0.275,0.514],[0.299,0.53,0.286,0.556,0.288,0.507],[0.462,0.745,0.455,0.692],[0.736,0.203,0.746,0.14,0.774,0.15,0.753,0.164,0.772,0.204,0.75,0.174,0.748,0.227],[0.671,0.219,0.711,0.186,0.716,0.256],[0.917,0.358,0.855,0.356,0.906,0.334]];
const LOTUS = [[0.496,0.089,0.496,0.714],[0.21,0.263,0.462,0.857],[0.784,0.263,0.532,0.857],[0.055,0.471,0.447,0.94,0.055,0.758],[0.94,0.471,0.94,0.758,0.548,0.94],[0.499,0,0.356,0.367,0.12,0.122,0.234,0.638,0.021,0.289,0,0.669,0.213,0.953,0.675,1,1,0.305,0.766,0.638,0.883,0.135,0.642,0.367]];
const PAW_PAD = [[0.5,0.125,0.689,0.125,0.909,0.62,0.778,0.875,0.091,0.735,0.311,0.125],[0.5,0,0.217,0.106,0,0.569,0.122,1,1,0.859,0.783,0.106]];
const PAW_TOE_L = [[0.5,0.197,0.75,0.715,0.361,0.803],[0.5,0,0,0.506,0.223,0.999,0.999,0.824,0.951,0.364]];
const PAW_TOE_R = [[0.5,0.197,0.75,0.715,0.362,0.803],[0.5,0,0,0.506,0.223,1,1,0.824,0.951,0.364]];
const PAW_TOE_OL = [[0.303,0.207,0.748,0.707,0.362,0.755],[0.304,0,0,0.224,0.065,0.7,0.687,1,1,0.506]];
const PAW_TOE_OR = [[0.727,0.208,0.45,0.784,0.313,0.529],[0.723,0,0,0.654,0.201,1,0.941,0.7,1,0.224]];
const HEART = [[0.736,0.066,0.941,0.496,0.5,0.923,0.059,0.496,0.059,0.256,0.346,0.066,0.512,0.293],[0.264,0,0,0.42,0.507,1,0.967,0.508,1,0.24,0.809,0,0.5,0.177]];
const ENVELOPE = [[0.178,0.311,0.178,0.459],[0.822,0.311,0.822,0.459],[0.764,0.059,0.764,0.503,0.5,0.7,0.236,0.503,0.236,0.059],[0.941,0.443,0.928,0.941,0.059,0.928,0.059,0.443,0.494,0.765],[0.207,0,0,0.381,0.039,1,1,0.961,0.793,0]];
const PIN = [[0.5,0.059,0.8,0.178,0.5,0.814,0.2,0.369],[0.65,0.715,0.933,0.844,0.067,0.844],[0.5,0,0.139,0.224,0.31,0.661,0.028,0.914,1,0.861,0.69,0.661,0.885,0.207]];
const PIN_DOT = [[0.519,0.2,0.83,0.502,0.359,0.799],[0.519,0,0,0.504,0.784,0.999,1,0.298]];
const PHONE = [[0.223,0.059,0.342,0.253],[0.788,0.646,0.907,0.841],[0.143,0.13,0.566,0.736,0.867,0.882,0.532,0.866,0.068,0.383],[0.222,0,0,0.395,0.552,0.969,0.864,0.969,1,0.772,0.809,0.587,0.595,0.683,0.322,0.399,0.435,0.184]];
const IG = [[0.5,0.344,0.344,0.414,0.414,0.656,0.656,0.586],[0.5,0.25,0.75,0.362,0.75,0.638,0.362,0.75,0.25,0.362],[0.312,0.094,0.094,0.192,0.192,0.906,0.906,0.808,0.906,0.192,0.688,0.094],[0.312,0,1,0.14,1,0.86,0.86,1,0,0.86,0,0.14]];
const TW = [[1,0.118,0.882,0.158,0.971,0.022,0.843,0.083,0.579,0,0.493,0.311,0.321,0.303,0.068,0.048,0.132,0.386,0.039,0.351,0.039,0.478,0.204,0.605,0.111,0.61,0.139,0.706,0.304,0.781,0,0.886,0.196,1,0.693,1]];
const YT = [[0.957,0.096,0.93,0.026,0.781,0.002,0.07,0.026,0,0.287,0,0.712,0.07,0.973,0.93,0.973,1,0.712,1,0.287],[0.375,0.773,0.375,0.227,0.688,0.5]];
const IC_BARS = [[0.324,0.442,0.101,0.933],[0.285,0.885,0.14,0.491],[0.631,0.248,0.408,0.933],[0.592,0.885,0.447,0.291],[0.939,0,0.715,0.933],[0.894,0.885,0.754,0.048],[0,0.952,1,1]];
const IC_TREND = [[0.798,0,0.634,0.366,0.393,0.128,0,0.384,0.388,0.183,0.634,0.433,0.956,0.079,1,0.226],[0.377,0.354,0.158,0.482,0.158,1,0.88,1,0.88,0.378,0.634,0.598],[0.842,0.957,0.202,0.957,0.202,0.506,0.634,0.652,0.842,0.476]];
const IC_GROWTH = [[0.841,0,0,0.074,0.074,1,1,0.926,1,0.074],[0.955,0.841,0.903,0.955,0.097,0.955,0.045,0.097,0.903,0.045],[0.614,0.273,0.722,0.273,0.142,0.739,0.432,0.443,0.545,0.557,0.75,0.307,0.795,0.409,0.795,0.227]];
const IC_BOARD = [[0.896,0,0,0.05,0.125,0.706,0.479,0.756,0.396,0.994,0.875,0.706],[0.833,0.644,0.193,0.711,0.167,0.222,0.833,0.222],[0.896,0.178,0.042,0.072],[0.479,0.333,0.25,0.378],[0.688,0.444,0.25,0.489],[0.688,0.556,0.25,0.6]];
const PERSON_M = [[0.474,0.206,0.83,0.26,0.991,0.416,0.985,0.526,0.834,0.546,0.734,0.341,0.755,0.954,0.71,0.993,0.53,0.962,0.514,0.625,0.465,0.962,0.245,0.971,0.254,0.35,0.188,0.538,0.004,0.511,0.098,0.3,0.443,0.207],[0.496,0,0.646,0.029,0.694,0.097,0.526,0.168,0.305,0.109,0.329,0.038,0.467,0.001]];
const PERSON_F = [[0.651,0.19,0.843,0.215,0.942,0.308,0.981,0.4,0.812,0.454,1,0.69,0.691,0.69,0.691,0.97,0.598,1,0.3,0.984,0.289,0.69,0,0.691,0.197,0.46,0.008,0.406,0.025,0.331,0.128,0.217,0.327,0.19],[0.507,0,0.669,0.029,0.72,0.097,0.507,0.168,0.293,0.097,0.325,0.039,0.475,0.001]];
// one ">" arrow head, outlined (stroke drawn as a filled ring, hence 2 loops)
const CHEV_UNIT = [[0,0,0,0.26,0.48,0.5,0,0.74,0,1,1,0.5],[0.02,0.746,0.513,0.5,0.02,0.253,0.02,0.026,0.966,0.499,0.02,0.973]];
const CHEV_SOLID = [[0,1,0,0.875,0.749,0.5,0,0.125,0,0,1,0.5]];
const SLASH = [[0.746,0,1,0,0.254,1,0,1]];

// ------------------------------------------------------------ tiny helpers --
const fill = (color, transparency) =>
  transparency ? { color, transparency } : { color };
const noLine = { type: 'none' };
const INSET = [7.2, 7.2, 3.6, 3.6];   // the deck's 0.1" / 0.05" text-box insets
// Every roundRect in the source carries adj=0, i.e. square corners. pptxgenjs
// omits the adjust value when it is falsy, so pass a hair above zero instead.
const SQUARE = 0.001;

/** addText carrying the deck's default text-box insets and top anchoring. */
function txt(slide, text, o) {
  slide.addText(text, Object.assign({ margin: INSET, valign: 'top' }, o));
}

/** Draw a custGeom shape from a normalised sub-path library. */
function drawPaths(slide, paths, o) {
  const pts = [];
  paths.forEach(sp => {
    for (let i = 0; i < sp.length; i += 2) {
      const p = { x: sp[i] * o.w, y: sp[i + 1] * o.h };
      if (i === 0) p.moveTo = true;
      pts.push(p);
    }
    pts.push({ close: true });
  });
  slide.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h, points: pts,
    fill: fill(o.color), line: noLine, flipH: o.flipH, rotate: o.rotate
  });
}

/** 4x4 grid of small "+" glyphs (the deck's most common corner ornament). */
function plusGrid(slide, o) {
  const cols = o.cols || 4, rows = o.rows || 4;
  const gx = o.w / (cols - 1 + 0.14), gy = o.h / (rows - 1 + 0.14);
  const arm = gx * 0.135, bar = Math.min(gx, gy) * 0.025;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const cx = o.x + c * gx + arm, cy = o.y + r * gy + arm;
      slide.addShape('rect', { x: cx - arm, y: cy - bar, w: arm * 2, h: bar * 2, fill: fill(o.color), line: noLine });
      slide.addShape('rect', { x: cx - bar, y: cy - arm, w: bar * 2, h: arm * 2, fill: fill(o.color), line: noLine });
    }
  }
}

/** Row of 19 vertical bars — the "barcode" strip along slide edges. */
function stripeBar(slide, o) {
  const n = 19, pitch = o.w / n;
  for (let i = 0; i < n; i++) {
    slide.addShape('rect', { x: o.x + i * pitch, y: o.y, w: pitch * 0.52, h: o.h, fill: fill(o.color), line: noLine });
  }
}

/** Three (or n) outlined chevrons pointing right. */
function chevrons(slide, o) {
  const n = o.count || 3, pitch = o.w / (n - 0.075 * (n - 1));
  for (let i = 0; i < n; i++) {
    drawPaths(slide, CHEV_UNIT, {
      x: o.x + i * pitch * 0.925, y: o.y, w: pitch, h: o.h, color: o.color, flipH: o.flipH
    });
  }
}

/**
 * Six nested "U" arches: straight legs from the top of the box down to
 * `straight`, closed by a half-ellipse. Each ring steps in 7.6% of the width
 * and its bottom rises 10.1%, which is how the source artwork is laid out.
 */
function archRings(slide, o) {
  for (let i = 0; i < 6; i++) {
    const left = i * 0.0759 * o.w;
    const w = o.w - 2 * left;
    const bottom = (1 - i * 0.10126) * o.h;
    const straight = (0.7062 - i * 0.0567) * o.h;
    slide.addShape('custGeom', {
      x: o.x + left, y: o.y, w, h: bottom, line: { color: o.color, width: 0.75 }, fill: { type: 'none' },
      points: [
        { x: 0, y: 0 },
        { x: 0, y: straight },
        { curve: { type: 'arc', hR: bottom - straight, wR: w / 2, stAng: 180, swAng: -180 }, x: w, y: straight },
        { x: w, y: 0 },
      ]
    });
  }
}

/** 5x5 grid of tiny plus-shaped dots. */
function dotGrid(slide, o) {
  const n = 5, gx = o.w / (n - 1 + 0.16), gy = o.h / (n - 1 + 0.16);
  const d = Math.min(gx, gy) * 0.16;
  for (let r = 0; r < n; r++) {
    for (let c = 0; c < n; c++) {
      slide.addShape('mathPlus', {
        x: o.x + c * gx, y: o.y + r * gy, w: d, h: d, fill: fill(o.color), line: noLine
      });
    }
  }
}

/**
 * Stand-in for the deck's photo frames. The source ships *empty* picture
 * placeholders — the .pptx contains no media parts at all — so these render as
 * bare page. A faint tinted block keeps the intended photo area readable.
 */
function imagePlaceholder(slide, o) {
  slide.addShape('rect', { x: o.x, y: o.y, w: o.w, h: o.h, fill: fill(C.photo), line: noLine });
}

/**
 * Translucent charcoal scrim. Slides 1/2/9/25 cover their photo frame with one,
 * so the frame itself is not drawn — the scrim over the white page is the look.
 */
function scrim(slide, o) {
  slide.addShape('rect', { x: o.x, y: o.y, w: o.w, h: o.h, fill: fill(C.charcoal, 15), line: noLine });
}

/** The recurring "bold caption + small paragraph" pair used on section slides. */
function captionBlock(slide, o) {
  txt(slide, o.title, {
    x: o.x, y: o.y, w: o.w, h: 0.404, fontFace: BODY, fontSize: 18, bold: true,
    color: o.titleColor || C.white
  });
  txt(slide, o.body, {
    x: o.x, y: o.y + 0.404, w: o.bodyW || o.w, h: o.bodyH || 0.863,
    fontFace: BODY, fontSize: 12, color: o.bodyColor || C.white,
    lineSpacingMultiple: 1.3
  });
}

/** Big serif slide title. */
function title(slide, text, o) {
  txt(slide, text, {
    x: o.x, y: o.y, w: o.w, h: o.h, fontFace: HEAD, fontSize: o.size || 40,
    color: o.color || C.brick, align: o.align || 'left',
    lineSpacingMultiple: o.ls
  });
}

/** Body copy in the theme's sans face. */
function body(slide, text, o) {
  txt(slide, text, {
    x: o.x, y: o.y, w: o.w, h: o.h, fontFace: BODY, fontSize: o.size || 12,
    color: o.color || C.black, align: o.align || 'left', valign: o.valign || 'top',
    lineSpacingMultiple: o.ls === undefined ? 1.3 : o.ls,
    bold: o.bold
  });
}

/** Body paragraph made of the three lorem fragments the deck reuses verbatim. */
function loremParagraph(slide, o) {
  const opt = {
    fontFace: BODY, fontSize: 12, color: o.color || C.black,
    lineSpacingMultiple: 1.3, align: 'left'
  };
  txt(slide, [
    { text: LOREM.med, options: opt },
    { text: LOREM.tail, options: opt },
    { text: LOREM.tail2, options: opt },
  ], { x: o.x, y: o.y, w: o.w, h: o.h });
}

// ------------------------------------------------------------ slide chrome --
// The master paints a small wordmark top-left and a page number bottom-right.
function chrome(slide, n) {
  txt(slide, 'EQUESTRIAN', {
    x: 0.304, y: 0.155, w: 1.79, h: 0.303, fontFace: HEAD, fontSize: 12,
    color: C.brick
  });
  txt(slide, String(n), {
    x: 12.223, y: 6.968, w: 0.57, h: 0.404, fontFace: HEAD, fontSize: 18,
    color: C.brickDeep, align: 'center'
  });
}

// ============================================================== slides 1-8 ==

// 1 — Cover: full-bleed dark photo panel, wordmark, horse silhouette.
function slide01(s) {
  scrim(s, { x: 0, y: 0, w: 13.388, h: 7.5 });
  drawPaths(s, HORSE, { x: 1.634, y: 1.473, w: 2.772, h: 1.757, color: C.white });
  txt(s, 'EQUESTRIAN', {
    x: 1.399, y: 3.628, w: 8.129, h: 1.192, fontFace: HEAD, fontSize: 72,
    color: C.white, lineSpacingMultiple: 0.9
  });
  body(s, LOREM.short, { x: 1.399, y: 4.866, w: 8.129, h: 0.338, color: C.white });
  txt(s, [
    { text: 'EQUESTRIAN SPORT', options: { breakLine: true } },
    { text: 'PRESENTATION TEMPLATE' },
  ], {
    x: 1.399, y: 5.406, w: 4.601, h: 0.774, fontFace: BODY, fontSize: 20,
    color: C.white
  });
  plusGrid(s, { x: 11.631, y: 0.669, w: 1.203, h: 1.14, color: C.white });
  stripeBar(s, { x: 9.875, y: 6.483, w: 3.513, h: 0.348, color: C.white });
}

// 2 — Section opener: text left, dark photo column right.
function slide02(s) {
  scrim(s, { x: 7.104, y: 0, w: 6.229, h: 7.5 });
  title(s, 'The World of Equestrian Sports', { x: 0.812, y: 1.984, w: 5.418, h: 1.447 });
  body(s, 'From Basics to Competitions ', { x: 0.812, y: 3.75, w: 5.126, h: 0.404, size: 18, bold: true, color: C.brick, ls: 1 });
  loremParagraph(s, { x: 0.812, y: 4.087, w: 5.418, h: 1.388 });
  body(s, LOREM.short, { x: 7.927, y: 5.395, w: 3.604, h: 0.863, color: C.white });
  plusGrid(s, { x: 11.797, y: 0.416, w: 1.203, h: 1.14, color: C.white });
  stripeBar(s, { x: 0, y: 6.397, w: 3.513, h: 0.348, color: C.olive });
}

// 3 — Three offset colour cards on the left, title block on the right.
function slide03(s) {
  const cards = [
    { x: 0, y: 0.941, w: 5.209, h: 1.851, color: C.brickDark, tx: 1.437, ty: 1.156, tw: 3.605, label: 'Dressage' },
    { x: 2.703, y: 2.94, w: 3.963, h: 1.926, color: C.olive, tx: 2.988, ty: 3.195, tw: 3.395, label: 'Show Jumping' },
    { x: 0, y: 5.015, w: 5.209, h: 1.851, color: C.charcoal, tx: 1.437, ty: 5.216, tw: 3.395, label: 'Eventing' },
  ];
  title(s, 'Types of Equestrian Sports', { x: 7.413, y: 1.337, w: 4.657, h: 2.121 });
  loremParagraph(s, { x: 7.413, y: 3.696, w: 4.428, h: 1.651 });
  cards.forEach(c => {
    s.addShape('rect', { x: c.x, y: c.y, w: c.w, h: c.h, fill: fill(c.color), line: noLine });
    captionBlock(s, { x: c.tx, y: c.ty, w: c.tw, body: LOREM.short });
    txt(s, c.label, {
      x: c.tx, y: c.ty, w: c.tw, h: 0.404, fontFace: BODY, fontSize: 18, bold: true,
      color: C.white
    });
  });
  dotGrid(s, { x: 0.772, y: 3.364, w: 1.157, h: 1.093, color: C.gold });
  chevrons(s, { x: 7.441, y: 5.725, w: 1.203, h: 0.653, color: C.clayDark, flipH: true });
  plusGrid(s, { x: 11.713, y: 0.396, w: 1.187, h: 1.124, color: C.gold });
}

// 4 — Pull quote over a photo, full-width red band along the bottom.
function slide04(s) {
  imagePlaceholder(s, { x: 0, y: 0, w: 5.905, h: 5.12 });
  title(s, QUOTE, { x: 6.667, y: 1.475, w: 5.333, h: 2.794, size: 32 });
  loremParagraph(s, { x: 5.718, y: 5.649, w: 5.556, h: 1.388 });
  s.addShape('rect', { x: 0, y: 5.12, w: SLIDE_W, h: 2.38, fill: fill(C.brickDark), line: noLine });
  captionBlock(s, {
    x: 0.971, y: 5.57, w: 3.605, bodyW: 8.251, title: 'Dressage',
    body: LOREM.short + 'natoque penatibus. amet orci et consequat. Morbi semper eros vitae tincidunt porta. Mauris euismod. Cum sociis natoque penatibus. Lorem ipsum dolor sit sit amet, consectetur adipiscing '
  });
  archRings(s, { x: 11.295, y: -0.05, w: 1.411, h: 1.076, color: C.olive });
  stripeBar(s, { x: 9.821, y: 6.401, w: 3.513, h: 0.348, color: C.white });
}

// 5 — Quote left, olive info column right, photo bottom-left.
function slide05(s) {
  imagePlaceholder(s, { x: 0, y: 3.75, w: 8.365, h: 3.75 });
  s.addShape('rect', { x: 8.365, y: 0, w: 4.968, h: 5.781, fill: fill(C.olive), line: noLine });
  title(s, QUOTE, { x: 1.265, y: 0.978, w: 6.713, h: 2.255, size: 32 });
  captionBlock(s, { x: 8.958, y: 0.727, w: 3.395, title: 'Show Jumping', body: LOREM.short });
  body(s, LOREM.med + ' Lorem ipsum dolor sit amet, consectetur adipiscing elit, ullamco\u00a0' + LOREM.tail2, {
    x: 8.958, y: 2.407, w: 3.395, h: 2.176, color: C.white
  });
  plusGrid(s, { x: 11.554, y: 4.954, w: 1.187, h: 1.124, color: C.white });
  chevrons(s, { x: 8.988, y: 6.119, w: 1.203, h: 0.653, color: C.clayDark, flipH: true });
}

// 6 — Quote left, orange caption card top-right, photo behind it.
function slide06(s) {
  imagePlaceholder(s, { x: 8.124, y: 1.875, w: 5.209, h: 5.625 });
  s.addShape('rect', { x: 8.124, y: 0, w: 5.209, h: 1.875, fill: fill(C.orangeDark), line: noLine });
  captionBlock(s, { x: 8.689, y: 0.354, w: 4.441, title: 'Eventing', body: LOREM.short });
  title(s, QUOTE, { x: 1.046, y: 1.406, w: 6.032, h: 2.794, size: 32 });
  body(s, 'From Basics to Competitions ', { x: 1.046, y: 4.675, w: 5.126, h: 0.404, size: 18, bold: true, color: C.brick, ls: 1 });
  loremParagraph(s, { x: 1.046, y: 5.012, w: 5.418, h: 1.388 });
  plusGrid(s, { x: 7.318, y: 5.215, w: 1.611, h: 1.527, color: C.cocoa });
}

// 7 — Khaki sidebar left, title + intro, wide photo bottom-right.
function slide07(s) {
  s.addShape('rect', { x: 0, y: 0, w: 3.312, h: SLIDE_H, fill: fill(C.khakiMid), line: noLine });
  imagePlaceholder(s, { x: 3.312, y: 3.047, w: 10.021, h: 4.453 });
  plusGrid(s, { x: 0.47, y: 0.657, w: 1.187, h: 1.124, color: C.white });
  title(s, 'Popular Horse Breeds', { x: 3.848, y: 0.723, w: 6.572, h: 0.774 });
  loremParagraph(s, { x: 3.958, y: 1.63, w: 8.558, h: 0.863 });
  body(s, LOREM.med + ' ', { x: 0.47, y: 4.57, w: 2.546, h: 1.913, color: C.white });
  archRings(s, { x: 11.295, y: -0.05, w: 1.411, h: 1.076, color: C.olive });
}

// 8 — Three numbered breed cards over an olive header block.
function slide08(s) {
  const cards = [
    { x: 2.127, color: C.brick, num: '01', name: 'Arabian', nx: 2.568, nw: 1.838, tx: 2.248, numX: 3.146, numY: 3.817 },
    { x: 5.307, color: C.charcoal, num: '02', name: 'Warmblood', nx: 5.748, nw: 1.838, tx: 5.428, numX: 6.238, numY: 3.817 },
    { x: 8.468, color: C.orangeDark, num: '03', name: 'Quarter Horse', nx: 8.658, nw: 2.34, tx: 8.59, numX: 9.399, numY: 3.793 },
  ];
  s.addShape('rect', { x: 0, y: 0, w: 8.837, h: 4.462, fill: fill(C.olive), line: noLine });
  body(s, LOREM.maurisLong, { x: 9.236, y: 0.918, w: 2.389, h: 1.651 });
  title(s, 'Popular Horse Breeds', { x: 1.333, y: 1.897, w: 6.572, h: 0.774, color: C.white });
  stripeBar(s, { x: -0.013, y: 0.651, w: 3.513, h: 0.348, color: C.white });
  cards.forEach(c => {
    s.addShape('roundRect', { x: c.x, y: 3.407, w: 2.72, h: 3.47, rectRadius: SQUARE, fill: fill(c.color), line: noLine });
    txt(s, c.num, {
      x: c.numX, y: c.numY, w: 0.857, h: 0.64, fontFace: HEAD, fontSize: 32, bold: true,
      color: C.white
    });
    txt(s, c.name, {
      x: c.nx, y: 4.457, w: c.nw, h: 0.499, fontFace: BODY, fontSize: 20, color: C.white,
      align: 'center', lineSpacingMultiple: 1.3, paraSpaceAfter: 6,
      margin: [7.2, 7.2, 3.6, 3.6]
    });
    body(s, LOREM.short.trim(), {
      x: c.tx, y: 4.924, w: 2.477, h: 1.277, color: C.white, align: 'center', ls: 1.5
    });
  });
}

// ============================================================= slides 9-16 ==

// 9 — Dark panel with big title left, three numbered notes right.
function slide09(s) {
  scrim(s, { x: 0, y: 0, w: 6.667, h: 7.5 });
  title(s, 'The Fundamental of Training and Riding Techniques', { x: 1.233, y: 1.592, w: 4.642, h: 3.467, color: C.white });
  [0, 1, 2].forEach(i => {
    const y = 1.058 + i * 1.828;
    txt(s, '0' + (i + 1), {
      x: 7.517, y, w: 0.857, h: 0.64, fontFace: HEAD, fontSize: 32, bold: true,
      color: C.charcoal
    });
    txt(s, 'Lorem Ipsum', {
      x: 8.375, y: y + 0.13, w: 3.819, h: 0.404, fontFace: HEAD, fontSize: 18,
      color: C.black
    });
    body(s, LOREM.qui, { x: 7.517, y: y + 0.657, w: 4.858, h: 0.865 });
  });
  for (let i = 0; i < 10; i++) {
    drawPaths(s, CHEV_SOLID, { x: 1.233 + i * 0.3152, y: 5.631, w: 0.293, h: 0.562, color: C.white });
  }
}

// 10 — Title block over a red band containing a heart icon + copy.
function slide10(s) {
  imagePlaceholder(s, { x: 8.413, y: 0, w: 4.921, h: 6.357 });
  s.addShape('roundRect', { x: 0, y: 4.027, w: 8.413, h: 3.473, rectRadius: SQUARE, fill: fill(C.brick), line: noLine });
  title(s, 'Basic Horse Care', { x: 1.229, y: 1.335, w: 5.448, h: 0.774 });
  body(s, LOREM.mauris, { x: 1.229, y: 2.244, w: 4.921, h: 0.863 });
  drawPaths(s, HEART, { x: 1.675, y: 4.56, w: 0.945, h: 0.835, color: C.white });
  txt(s, 'Health Care', {
    x: 1.229, y: 5.53, w: 1.838, h: 0.467, fontFace: BODY, fontSize: 18, bold: true,
    color: C.white, align: 'center', lineSpacingMultiple: 1.3,
    paraSpaceAfter: 6
  });
  body(s, LOREM.mauris, { x: 3.362, y: 4.671, w: 3.482, h: 1.126, color: C.white });
  plusGrid(s, { x: 7.755, y: 4.671, w: 1.399, h: 1.326, color: C.white });
  archRings(s, { x: 6.296, y: 0, w: 1.411, h: 1.076, color: C.olive });
}

// 11 — Khaki panel left; three icon tiles with headings on the right.
function slide11(s) {
  const rows = [
    { y: 1.159, color: C.orange, label: 'Daily Routines', labelColor: C.orange, w: 4.627, icon: 'paw' },
    { y: 2.991, color: C.brick, label: 'Health care', labelColor: C.brick, w: 4.627, icon: 'heart' },
    { y: 4.715, color: C.olive, label: 'Proper Feeding', labelColor: C.olive, w: 4.919, icon: 'lotus' },
  ];
  s.addShape('rect', { x: 0, y: 0, w: 6.667, h: SLIDE_H, fill: fill(C.khakiMid), line: noLine });
  title(s, 'Basic Horse Care', { x: 1.605, y: 2.075, w: 3.462, h: 1.447, color: C.white });
  body(s, LOREM.mauris, { x: 1.641, y: 3.75, w: 3.297, h: 1.817, size: 16, color: C.white });
  rows.forEach(r => {
    s.addShape('rect', { x: 5.777, y: r.y, w: 1.545, h: 1.464, fill: fill(r.color), line: noLine });
    txt(s, r.label, {
      x: 7.714, y: r.y + 0.133, w: 2.637, h: 0.37, fontFace: BODY, fontSize: 16, bold: true,
      color: r.labelColor
    });
    body(s, LOREM.ullamco, { x: 7.714, y: r.y + 0.469, w: r.w, h: 0.863 });
  });
  // paw print (row 1), heart (row 2), lotus (row 3)
  drawPaths(s, PAW_TOE_L, { x: 6.352, y: 1.539, w: 0.172, h: 0.218, color: C.white });
  drawPaths(s, PAW_TOE_OL, { x: 6.176, y: 1.719, w: 0.188, h: 0.207, color: C.white });
  drawPaths(s, PAW_TOE_R, { x: 6.566, y: 1.539, w: 0.172, h: 0.218, color: C.white });
  drawPaths(s, PAW_TOE_OR, { x: 6.706, y: 1.718, w: 0.206, h: 0.207, color: C.white });
  drawPaths(s, PAW_PAD, { x: 6.309, y: 1.8, w: 0.472, h: 0.343, color: C.white });
  drawPaths(s, HEART, { x: 6.231, y: 3.494, w: 0.591, h: 0.522, color: C.white });
  drawPaths(s, LOTUS, { x: 6.287, y: 5.157, w: 0.516, h: 0.514, color: C.white });
  for (let i = 0; i < 7; i++) {
    drawPaths(s, SLASH, { x: -0.221 + i * 0.3226, y: 1.017, w: 0.314, h: 0.431, color: C.white });
  }
  plusGrid(s, { x: 0.476, y: 5.972, w: 1.187, h: 1.124, color: C.white });
}

// 12 — Title right, olive stats bar bottom-left over a photo.
function slide12(s) {
  imagePlaceholder(s, { x: 0, y: 0, w: 8.078, h: 5.241 });
  title(s, 'Major Equestrian Competition', { x: 8.43, y: 1.9, w: 3.947, h: 2.121 });
  body(s, LOREM.med + ' veniam minim as tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam.', {
    x: 8.461, y: 4.257, w: 4.291, h: 2.219
  });
  s.addShape('roundRect', { x: 0, y: 5.241, w: 8.078, h: 2.259, rectRadius: SQUARE, fill: fill(C.olive), line: noLine });
  [['76%', 1.871], ['96%', 3.257]].forEach(([pct, x]) => {
    txt(s, pct, {
      x, y: 5.757, w: 1.385, h: 0.572, fontFace: HEAD, fontSize: 28,
      color: C.white
    });
    body(s, 'Lorem ipsum', { x, y: 6.289, w: 1.385, h: 0.34, color: C.white });
  });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, veniam minim magna aliqua. ', {
    x: 4.92, y: 5.757, w: 2.592, h: 0.863, color: C.white
  });
  stripeBar(s, { x: 6.514, y: 0.672, w: 3.513, h: 0.348, color: C.clayDark });
}

// 13 — Title, then a red panel with a progress bar and supporting copy.
function slide13(s) {
  imagePlaceholder(s, { x: 6.302, y: 0, w: 7.032, h: 4.522 });
  s.addShape('roundRect', { x: 0, y: 4.522, w: 7.378, h: 2.978, rectRadius: SQUARE, fill: fill(C.brickDark), line: noLine });
  title(s, 'Basic Training for Competition', { x: 1.003, y: 2.255, w: 5.231, h: 1.447 });
  txt(s, 'Training', {
    x: 1.32, y: 4.916, w: 4.128, h: 0.37, fontFace: HEAD, fontSize: 16, bold: true,
    color: C.white
  });
  txt(s, '96%', {
    x: 5.867, y: 4.916, w: 0.8, h: 0.37, fontFace: HEAD, fontSize: 16, bold: true,
    color: C.white, align: 'right'
  });
  s.addShape('rect', { x: 1.37, y: 5.354, w: 4.496, h: 0.259, fill: fill(C.khaki), line: noLine });
  s.addShape('rect', { x: 5.867, y: 5.354, w: 0.8, h: 0.259, fill: fill(C.greySoft), line: noLine });
  body(s, LOREM.tempor, { x: 1.32, y: 5.801, w: 5.346, h: 0.992, size: 14, color: C.white });
  body(s, LOREM.maurisLong, { x: 7.868, y: 4.916, w: 4.275, h: 1.297 });
  chevrons(s, { x: 7.987, y: 5.99, w: 1.319, h: 0.716, color: C.gold });
  archRings(s, { x: 5.596, y: 0, w: 1.411, h: 1.076, color: C.olive });
}

// 14 — Trainer profile: name card, years badge and social handles.
function slide14(s) {
  imagePlaceholder(s, { x: 0, y: 0, w: 5.619, h: SLIDE_H });
  title(s, 'Meet Our Professional Trainer', { x: 6.206, y: 1.219, w: 6.141, h: 1.447 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, veniam minim as tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco\u00a0', {
    x: 6.206, y: 2.898, w: 5.935, h: 0.99, size: 14
  });
  s.addShape('rect', { x: 1.233, y: 4.576, w: 5.312, h: 1.846, fill: fill(C.brick), line: noLine });
  txt(s, 'Paula Almodovar', {
    x: 1.573, y: 4.955, w: 4.632, h: 0.735, fontFace: HEAD, fontSize: 32, color: C.white,
    lineSpacingMultiple: 0.9
  });
  txt(s, 'Professional Trainer', {
    x: 1.62, y: 5.671, w: 3.371, h: 0.438, fontFace: BODY, fontSize: 20, color: C.white
   
  });
  s.addShape('rect', { x: 6.692, y: 4.606, w: 2.143, h: 1.817, fill: fill(C.orange), line: noLine });
  txt(s, '13+', {
    x: 6.985, y: 5.086, w: 1.385, h: 0.707, fontFace: HEAD, fontSize: 36, color: C.white
   
  });
  body(s, 'Years Eperiences', { x: 6.985, y: 5.618, w: 1.785, h: 0.337, color: C.white });
  s.addShape('rect', { x: 8.982, y: 4.606, w: 2.875, h: 1.817, fill: fill(C.olive), line: noLine });
  const handles = [
    { icon: IG, y: 4.812, iy: 4.856, iw: 0.257, ih: 0.257, text: '@almodovar' },
    { icon: TW, y: 5.308, iy: 5.39, iw: 0.262, ih: 0.214, text: '@almodovar' },
    { icon: YT, y: 5.805, iy: 5.904, iw: 0.246, ih: 0.169, text: 'Paula Almodovar' },
  ];
  handles.forEach(h => {
    drawPaths(s, h.icon, { x: 9.319, y: h.iy, w: h.iw, h: h.ih, color: C.white });
    txt(s, h.text, {
      x: 9.647, y: h.y, w: 2.045, h: 0.386, fontFace: HEAD, fontSize: 14, color: C.white,
      lineSpacingMultiple: 1.3, paraSpaceAfter: 6
    });
  });
  chevrons(s, { x: 11.143, y: 0.444, w: 1.426, h: 0.774, color: C.gold });
}

// 15 — Pricing: three tall cards with plan / price / CTA button.
function slide15(s) {
  const plans = [
    { x: 1.535, color: C.orange, name: 'Basic', price: '$45' },
    { x: 5.076, color: C.olive, name: 'Premium', price: '$55' },
    { x: 8.647, color: C.brick, name: 'Personal', price: '$75' },
  ];
  title(s, 'Get Your Training Class', { x: 1.877, y: 0.993, w: 9.579, h: 0.841, align: 'center' });
  plans.forEach(p => {
    const dx = p.x - 1.535;
    s.addShape('roundRect', { x: p.x, y: 2.254, w: 3.181, h: 4.253, rectRadius: SQUARE, fill: fill(p.color), line: noLine });
    txt(s, p.name, {
      x: 1.896 + dx, y: 2.866, w: 2.42, h: 0.438, fontFace: BODY, fontSize: 20, bold: true,
      color: C.white, align: 'center'
    });
    txt(s, p.price, {
      x: 2.035 + dx, y: 3.231, w: 2.141, h: 1.01, fontFace: BODY, fontSize: 54, color: C.white,
      align: 'center'
    });
    txt(s, 'Permonth', {
      x: 1.896 + dx, y: 4.196, w: 2.42, h: 0.37, fontFace: BODY, fontSize: 16, bold: true,
      color: C.white, align: 'center'
    });
    body(s, 'Ut wisi enim ad minim veniam, quis', {
      x: 1.926 + dx, y: 4.687, w: 2.359, h: 0.684, size: 14, color: C.white, align: 'center'
    });
    s.addShape('roundRect', { x: 2.087 + dx, y: 5.601, w: 2.037, h: 0.518, rectRadius: SQUARE, fill: fill(C.white), line: noLine });
    txt(s, 'Lorem Ipsum', {
      x: 2.087 + dx, y: 5.692, w: 2.037, h: 0.337, fontFace: BODY, fontSize: 14, bold: true,
      color: C.black, align: 'center'
    });
  });
  chevrons(s, { x: -0.153, y: 1.46, w: 1.426, h: 0.774, color: C.gold });
}

// 16 — World map (simplified land masses) with three highlight tiles.
// Each entry is [x, y, w, h, colourKey, sub-paths normalised inside the box].
const MAP_COLORS = { G: C.greySoft, O: C.orange, R: C.brick, K: C.cocoa };
const MAP = [
  [8.59,0.79,3.98,1.76,'O',[[1,0.47,0.71,0.28,0.5,0.29,0.49,0.13,0.31,0.47,0.31,0.28,0.1,0.52,0.06,0.38,0.1,0.93,0.17,1,0.16,0.82,0.29,0.72,0.4,0.85,0.61,0.77,0.66,0.97,0.71,0.64,0.85,0.56,0.8,0.82,1,0.47],[0.2,0.27,0.22,0.35,0.29,0.15,0.2,0.27]]],
  [4.82,0.86,2.08,1.6,'G',[[0.57,0.06,0.68,0.24,0.91,0.03,0.57,0.06],[0.91,0.93,0.86,0.64,0.69,0.84,0.52,0.67,0.57,0.3,0,0.42,0.21,0.88,0.91,0.93]]],
  [6.43,0.84,1.43,1.07,'G',[[0.95,0.09,0.8,0.16,0.84,0.08,0.66,0.1,0.85,0.05,0.77,0.01,0.49,0.01,0.41,0.05,0.47,0.11,0.21,0.09,0,0.26,0.27,0.45,0.41,0.97,0.49,1,0.52,0.82,0.84,0.63,0.76,0.64,0.85,0.61,0.78,0.51,0.89,0.46,0.83,0.36,0.91,0.34,0.83,0.31,1,0.12,0.95,0.09]]],
  [9.86,2.13,1.4,0.97,'G',[[0.99,0.18,0.93,0.19,0.85,0.02,0.76,0,0.67,0.17,0.75,0.22,0.49,0.37,0.36,0.34,0.22,0.13,0.1,0.27,0.11,0.35,0,0.41,0.04,0.53,0.11,0.55,0.08,0.65,0.24,0.76,0.35,0.7,0.42,0.91,0.5,0.86,0.59,0.94,0.79,0.69,0.73,0.56,0.8,0.5,0.71,0.44,0.94,0.34,0.99,0.18]]],
  [5.21,2.25,1.35,0.68,'G',[[0.99,0.11,0.73,0.36,0.7,0.16,0.63,0.31,0.71,0.15,0.64,0.08,0.57,0.13,0.61,0.07,0.51,0,0.03,0.03,0,0.3,0.07,0.63,0.13,0.71,0.32,0.74,0.47,0.96,0.53,0.81,0.62,0.85,0.64,0.77,0.7,0.79,0.76,1,0.75,0.76,0.86,0.57,0.84,0.42,0.84,0.53,1,0.2,0.99,0.11]]],
  [6.39,3.4,0.93,0.95,'R',[[0.53,0.99,0.64,0.77,0.84,0.69,0.88,0.45,1,0.32,0.88,0.2,0.74,0.21,0.67,0.14,0.63,0.19,0.65,0.13,0.57,0.14,0.58,0.03,0.38,0.1,0.36,0,0.23,0.03,0.27,0.08,0.11,0.09,0.11,0.23,0,0.32,0.08,0.41,0.22,0.37,0.4,0.54,0.4,0.69,0.51,0.81,0.42,0.9,0.53,0.99]]],
  [10.74,3.78,0.96,0.85,'K',[[0.99,0.5,0.72,0,0.67,0.2,0.55,0.12,0.59,0.04,0.48,0.01,0.42,0.13,0.29,0.11,0.18,0.27,0.03,0.31,0,0.39,0.07,0.71,0.46,0.6,0.55,0.72,0.62,0.66,0.58,0.72,0.61,0.69,0.68,0.82,0.82,0.85,0.91,0.8,0.99,0.5],[0.82,0.91,0.77,0.9,0.83,1,0.87,0.95,0.82,0.91]]],
  [8.25,1.02,0.6,0.95,'G',[[0.97,0.53,0.43,0.55,0.01,0.96,0.27,0.97,0.5,0.58,0.97,0.53],[0.28,0.07,0.44,0.1,0.32,0.15,0.44,0.22,0.56,0.1,0.76,0.15,0.28,0.07]]],
  [4.19,1.47,0.91,0.62,'G',[[0.98,0.94,0.7,0.69,0.7,0.12,0.29,0,0.04,0.17,0.2,0.34,0.01,0.37,0.19,0.44,0.06,0.68,0.29,0.76,0.13,0.99,0.46,0.62,0.41,0.75,0.52,0.64,0.73,0.7,0.98,0.94]]],
  [9.73,2.64,0.65,0.69,'G',[[0.99,0.3,0.49,0.3,0.37,0.17,0.38,0,0.21,0.05,0.26,0.14,0,0.45,0.07,0.56,0.15,0.51,0.33,1,0.43,0.73,0.72,0.52,0.71,0.35,0.83,0.4,0.85,0.52,0.99,0.3]]],
  [6.4,4.04,0.47,0.94,'G',[[0.99,0.11,0.75,0.16,0.39,0,0.26,0.07,0.01,0.84,0.26,0.9,0.39,0.74,0.29,0.71,0.52,0.58,0.42,0.54,0.8,0.46,0.79,0.24,0.99,0.11]]],
  [9.21,2.07,0.95,0.45,'G',[[0.99,0.41,0.58,0,0.37,0.1,0.37,0.34,0.12,0.26,0,0.48,0.17,0.62,0.16,0.96,0.31,0.64,0.54,1,0.62,0.85,0.84,0.9,0.99,0.41]]],
  [5.39,2.73,0.71,0.44,'G',[[0.83,0.94,0.97,0.82,0.99,0.62,0.89,0.65,0.84,0.82,0.73,0.82,0.54,0.18,0.46,0.22,0.35,0.05,0,0.01,0.24,0.57,0.07,0.05,0.14,0.09,0.39,0.76,0.66,0.97,0.83,0.94]]],
  [10.17,2.17,0.73,0.32,'G',[[0.96,0.42,0.9,0.21,0.64,0.28,0.33,0,0.27,0.24,0.13,0.13,0,0.29,0.25,0.91,0.51,1,0.96,0.42]]],
  [7.93,2.6,0.49,0.47,'G',[[0.35,0.19,0.36,0.29,0.01,0.48,0.01,0.57,0.58,1,1,0.76,0.77,0.18,0.83,0.01,0.45,0.04,0.35,0.19]]],
  [6.35,3.93,0.21,1.06,'G',[[0.45,0.89,0.24,0.82,0.99,0.13,0.8,0.02,0,0.78,1,0.98,0.45,0.89]]],
  [8.93,2.74,0.49,0.4,'G',[[0.6,0.21,0.25,0,0.01,0.16,0.38,1,0.96,0.79,1,0.67,0.82,0.57,0.6,0.21]]],
  [8.42,3.39,0.43,0.45,'G',[[0.91,0.66,0.99,0.1,0.7,0,0.36,0.03,0.19,0.54,0,0.56,0.49,0.67,0.51,0.88,0.89,1,0.91,0.66]]],
  [8.63,2.96,0.39,0.48,'G',[[0.8,0.02,0.19,0.07,0.14,0.38,0,0.47,0.09,0.72,0.39,0.96,0.83,0.94,0.66,0.77,1,0.27,0.8,0.02]]],
  [9.15,2.53,0.46,0.39,'G',[[0.97,0.87,0.89,0.24,0.64,0.12,0.35,0.25,0.22,0.01,0.01,0,0.08,0.43,0.36,0.81,0.86,1,0.97,0.87]]],
  [11.13,2.37,0.4,0.4,'G',[[0.63,0.34,0.09,0.82,0.62,0.74,0.63,0.34]]],
  [8.4,1.56,0.29,0.51,'G',[[0.99,0.24,0.71,0,0.25,0.25,0,0.76,0.15,1,0.6,0.7,0.48,0.51,0.99,0.24]]],
  [7.85,2.92,0.39,0.37,'G',[[0.95,0.41,0.34,0,0.43,0.63,0.04,0.63,0.07,0.89,0.19,0.83,0.37,0.99,0.49,0.76,0.96,0.65,0.95,0.41]]],
  [9.55,2.61,0.39,0.35,'G',[[0.95,0.1,0.71,0.03,0.51,0.42,0,0.57,0.15,0.78,0.04,0.88,0.6,0.95,0.53,0.74,0.95,0.1]]],
  [6.22,3.53,0.3,0.44,'G',[[0.91,0.95,0.98,0.65,0.58,0.4,0.9,0.13,0.48,0,0.19,0.27,0.09,0.17,0,0.24,0.4,0.77,0.91,0.95]]],
  [8.51,4.04,0.38,0.33,'G',[[0.95,0.37,0.82,0,0.32,0.38,0.22,0.22,0,0.5,0.12,0.93,0.63,0.89,0.95,0.37]]],
  [8.36,2.71,0.36,0.34,'G',[[0.11,0.06,0,0.23,0.05,0.63,0.98,0.97,0.98,0.07,0.74,0,0.59,0.21,0.28,0.01,0.11,0.06]]],
  [6.28,3.23,0.28,0.4,'G',[[0.15,0.73,0.78,0.99,0.99,0.37,0.54,0.34,0.65,0.01,0.34,0.08,0.09,0.32,0,0.66,0.15,0.73]]],
  [8.14,2.96,0.36,0.28,'G',[[0.96,0.14,0.77,0,0.26,0.39,0.22,0.7,0.01,0.73,0.06,0.91,0.82,0.89,0.98,0.57,0.96,0.14]]],
  [11.99,4.38,0.29,0.35,'G',[[0.61,0.48,0,0.87,0.25,0.97,0.61,0.48],[0.9,0.24,0.51,0,0.7,0.18,0.57,0.36,0.69,0.53,0.9,0.24]]],
  [8.65,2.16,0.41,0.24,'G',[[0.89,0.62,1,0.37,0.65,0,0.07,0.11,0,0.51,0.38,0.56,0.33,0.88,0.52,0.7,0.8,0.93,0.89,0.62]]],
  [9.43,2.37,0.41,0.24,'G',[[0.97,0.57,0.83,0.46,0.71,0.6,0.54,0.25,0.01,0.06,0,0.49,0.15,0.34,0.65,0.99,0.77,0.58,0.97,0.57]]],
  [8.82,3.77,0.25,0.4,'G',[[1,0.26,0.96,0,0.44,0.06,0.46,0.39,0,0.26,0.27,0.37,0.18,1,0.47,0.85,0.41,0.55,1,0.26]]],
  [8.02,2.2,0.34,0.29,'R',[[0.87,0.23,0.47,0,0,0.27,0.22,0.81,0.56,0.9,0.86,0.76,0.87,0.23]]],
  [8.62,1.51,0.25,0.39,'G',[[0.85,0.84,1,0.73,0.84,0.64,0.78,0.05,0.01,0.11,0.41,0.48,0,0.75,0.01,0.95,0.85,0.84]]],
  [8.89,3.17,0.35,0.27,'G',[[0.07,0.66,0.19,0.89,0.44,1,0.79,0.88,1,0.62,0.74,0.52,0.5,0.04,0.33,0,0.02,0.54,0.07,0.66]]],
  [6.5,3.75,0.29,0.32,'G',[[0.99,0.61,0.77,0.32,0.38,0.17,0.36,0,0.02,0.1,0.16,1,0.6,0.92,0.7,0.72,0.96,0.75,0.99,0.61]]],
  [8.45,2.96,0.23,0.39,'G',[[1,0.25,0.14,0.03,0.19,0.42,0,0.62,0.19,0.84,0.04,0.87,0.3,1,0.89,0.78,0.78,0.59,1,0.25]]],
  [8.41,3.92,0.3,0.3,'G',[[0.37,0.95,0.61,0.96,0.69,0.12,0.88,0.13,1,0.03,0.74,0.09,0,0.03,0.21,0.46,0.25,0.85,0.37,0.95]]],
  [7.74,2.86,0.28,0.31,'G',[[0.08,0.87,0.39,1,0.97,0.93,0.86,0.19,1,0.19,0.69,0,0.41,0.11,0.33,0.49,0,0.49,0.08,0.87]]],
  [6.41,3.24,0.32,0.27,'G',[[0.1,0.06,0.06,0.45,0.46,0.5,0.6,0.99,0.62,0.69,0.93,0.65,1,0.32,0.3,0,0.16,0.28,0.1,0.06]]],
  [10.26,2.83,0.2,0.44,'G',[[0.81,0.91,0.62,0.51,1,0.39,0.59,0.27,0.66,0.01,0,0.43,0.19,0.69,0.5,0.62,0.71,1,0.81,0.91]]],
  [7.94,1.88,0.24,0.36,'G',[[1,0.77,0.49,0.19,0.12,0.33,0.52,0.62,0.3,1,1,0.77]]],
  [7.94,1.88,0.24,0.36,'G',[[1,0.77,0.49,0.19,0.12,0.33,0.52,0.62,0.3,1,1,0.77]]],
  [9.54,2.57,0.33,0.24,'G',[[0.97,0.14,0.8,0.22,0.74,0,0.63,0.17,0.37,0.1,0.19,0.36,0.04,0.34,0.03,0.98,0.42,0.95,0.97,0.14]]],
  [8.41,3.66,0.27,0.29,'G',[[0.06,0.91,0.93,0.97,0.84,0.58,1,0.4,0.84,0.42,0.81,0.1,0.48,0.19,0.39,0,0.06,0.03,0.06,0.91]]],
  [8.29,2.32,0.27,0.29,'G',[[0.83,0.56,0.49,0,0,0.2,0.78,0.88,0.83,0.56]]],
  [9.07,3.23,0.24,0.33,'G',[[0.23,0.19,0.67,0.31,0.37,0.54,0.08,0.58,0.05,1,0.78,0.46,1,0,0.3,0.15,0.23,0.06,0.23,0.19]]],
  [8.73,2.47,0.44,0.17,'G',[[0.97,0.7,0.9,0.07,0.37,0,0.01,0.32,0.18,0.97,1,0.88,0.97,0.7]]],
  [9.35,2.45,0.33,0.21,'G',[[0.99,0.63,0.43,0,0.31,0.21,0.02,0.05,0.18,0.26,0.01,0.33,0.1,0.71,0.33,0.59,0.7,1,0.99,0.63]]],
  [10.84,3.08,0.23,0.31,'G',[[0.53,0.39,0.9,0.56,0.31,0.04,0.53,0.39]]],
  [10.38,3.03,0.19,0.36,'G',[[0.6,0.53,1,0.41,0.91,0.22,0.02,0.07,0.22,0.54,0.05,0.83,0.52,0.99,0.15,0.75,0.26,0.5,0.6,0.53]]],
  [8.81,3.54,0.26,0.26,'G',[[0.93,0.8,0.75,0.19,0.4,0.01,0.4,0.15,0.18,0.16,0.1,0.02,0.07,0.67,0.6,1,0.9,0.96,0.93,0.8]]],
  [8.7,2.74,0.26,0.26,'G',[[0.93,0.22,0.86,0.06,0.39,0.1,0.05,0,0.02,0.98,0.92,0.95,1,0.77,0.69,0.17,0.83,0.41,0.93,0.22]]],
  [8.2,3.19,0.28,0.23,'G',[[0.61,0.72,0.73,0.8,0.99,0.25,0.91,0.03,0.57,0.12,0.14,0.04,0.03,0.75,0.3,1,0.49,0.95,0.61,0.72]]],
  [7.92,2.42,0.3,0.21,'R',[[0.99,0.18,0,0.05,0.05,0.24,0.26,0.27,0.14,0.55,0.25,1,0.68,0.8,0.99,0.18]]],
  [10.43,3.37,0.45,0.14,'G',[[0.18,0.35,0,0.02,0.2,0.94,0.18,0.35],[0.96,0.17,0.85,0,0.54,1,0.77,0.92,0.96,0.17]]],
  [7.83,2.64,0.28,0.22,'G',[[0.37,1,0.37,0.84,0.99,0.45,0.92,0.1,0.63,0,0.49,0.27,0.35,0.3,0.3,0.69,0,0.96,0.37,1]]],
  [8.63,3.71,0.27,0.24,'G',[[0.77,0.04,0.58,0.01,0.67,0.49,0.17,0.26,0.01,0.48,0.1,0.96,0.4,1,0.95,0.58,1,0.21,0.77,0.04]]],
  [5.98,3.48,0.39,0.16,'G',[[0.97,0.25,0.77,0,0.64,0.42,0.77,1,0.97,0.25],[0.02,0.22,0.05,0.4,0.05,0.13,0.02,0.22]]],
  [10.48,2.96,0.17,0.36,'G',[[0.75,0.12,0,0.07,0.32,0.16,0.25,0.29,0.83,0.72,0.34,0.99,0.98,0.75,0.49,0.34,0.75,0.12]]],
  [9.12,3.79,0.17,0.34,'G',[[0.85,0.03,0.58,0.24,0.12,0.3,0.2,0.59,0,0.71,0.08,0.92,0.55,0.92,0.85,0.29,0.98,0.31,0.85,0.03]]],
  [8.47,3.25,0.29,0.2,'G',[[0.92,0.57,0.62,0,0.43,0.27,0.06,0.44,0.01,0.74,0.13,1,0.35,0.71,0.61,0.81,1,0.7,0.92,0.57]]],
  [11.4,3.59,0.3,0.19,'G',[[0.72,0.92,0,0,0,0.78,0.25,0.62,0.72,0.92]]],
  [9.03,2.6,0.23,0.22,'G',[[0.84,0.91,1,0.88,0.68,0.45,0.75,0.19,0.38,0.01,0.22,0.36,0,0.49,0.03,0.64,0.59,0.97,0.84,0.91]]],
  [8.28,2.08,0.21,0.24,'G',[[0.12,0.26,0,0.63,0.25,0.79,0.17,0.96,0.79,0.98,0.68,0.6,1,0.48,0.85,0.02,0.29,0,0.12,0.26]]],
  [9.12,3.07,0.28,0.17,'G',[[0.71,0.05,0.54,0.05,0.38,0.31,0.06,0.22,0.07,0.9,0.82,0.49,0.9,0.36,0.71,0.05]]],
  [8.59,3.94,0.21,0.23,'G',[[0.92,0.41,0.56,0,0.11,0.06,0.11,0.46,0,0.44,0.06,0.99,0.32,0.78,0.58,0.86,0.99,0.47,0.92,0.41]]],
  [8.34,3.21,0.17,0.27,'G',[[0.42,0.95,1,0.92,0.8,0.73,0.92,0.49,0.72,0.32,0.92,0.28,0.78,0,0.38,0.61,0,0.74,0.42,0.95]]],
];
function slide16(s) {
  const tiles = [
    { x: 5.907, color: C.orange, tag: 'POLO', val: '93,65 +', tx: 6.034 },
    { x: 8.034, color: C.brick, tag: 'EVENTING', val: '25,60 +', tx: 8.162 },
    { x: 10.158, color: C.cocoa, tag: 'DRESSAGE', val: '37,99 +', tx: 10.286 },
  ];
  MAP.forEach(([x, y, w, h, key, paths]) => drawPaths(s, paths, { x, y, w, h, color: MAP_COLORS[key] }));
  title(s, 'Popular Equestrian Sport', { x: 0.933, y: 2.922, w: 4.489, h: 2.121 });
  body(s, LOREM.maurisMontes, { x: 0.902, y: 5.267, w: 4.663, h: 1.277, ls: 1.5 });
  tiles.forEach(t => {
    s.addShape('roundRect', { x: t.x, y: 5.478, w: 1.883, h: 1.066, rectRadius: SQUARE, fill: fill(t.color), line: noLine });
    txt(s, t.tag, {
      x: t.tx, y: 5.641, w: 1.684, h: 0.303, fontFace: BODY, fontSize: 12, bold: true,
      color: C.white
    });
    txt(s, t.val, {
      x: t.tx, y: 5.943, w: 1.38, h: 0.438, fontFace: BODY, fontSize: 20, color: C.white
    });
  });
  stripeBar(s, { x: -0.292, y: 1.766, w: 3.513, h: 0.348, color: C.clayDark });
}

// ============================================================ slides 17-25 ==

// 17 — Four diamonds around a central "Main Idea" diamond.
function slide17(s) {
  const cells = [
    { x: 4.162, y: 1.798, color: C.cocoa, label: 'Idea One', lx: 4.742, lw: 1.284, icon: IC_BARS, ix: 5.192, iy: 2.504, iw: 0.387, ih: 0.356 },
    { x: 6.727, y: 1.798, color: C.brick, label: 'Idea Two', lx: 7.334, lw: 1.284, icon: IC_GROWTH, ix: 7.776, iy: 2.508, iw: 0.346, ih: 0.348 },
    { x: 4.161, y: 4.363, color: C.orange, label: 'Idea Four', lx: 4.742, lw: 1.284, icon: IC_BOARD, ix: 5.196, iy: 5.172, iw: 0.378, ih: 0.354 },
    { x: 6.727, y: 4.363, color: C.olive, label: 'Idea Three', lx: 7.211, lw: 1.475, icon: IC_TREND, ix: 7.752, iy: 5.146, iw: 0.395, ih: 0.354 },
  ];
  const side = [
    { x: 0.887, y: 2.517, align: 'right', label: 'Idea One' },
    { x: 0.887, y: 5.046, align: 'right', label: 'Idea Four' },
    { x: 9.445, y: 2.517, align: 'left', label: 'Idea Two' },
    { x: 9.445, y: 5.046, align: 'left', label: 'Idea Three' },
  ];
  title(s, 'Creative Infographic', { x: 2.388, y: 0.65, w: 8.558, h: 0.774, align: 'center' });
  cells.forEach(c => {
    s.addShape('diamond', { x: c.x, y: c.y, w: 2.445, h: 2.445, fill: fill(c.color), line: noLine });
    drawPaths(s, c.icon, { x: c.ix, y: c.iy, w: c.iw, h: c.ih, color: C.white });
    txt(s, c.label, {
      x: c.lx, y: c.y + 1.238, w: c.lw, h: 0.37, fontFace: HEAD, fontSize: 16, color: C.white,
      align: 'center'
    });
  });
  s.addShape('diamond', { x: 5.267, y: 2.904, w: 2.799, h: 2.799, fill: fill(C.greyPale), line: noLine });
  txt(s, 'Main Idea', {
    x: 6.064, y: 3.849, w: 1.205, h: 0.909, fontFace: HEAD, fontSize: 24, color: C.black,
    align: 'center'
  });
  side.forEach(t => {
    txt(s, t.label, {
      x: t.x, y: t.y, w: 3.001, h: 0.404, fontFace: HEAD, fontSize: 18, color: C.ink,
      align: t.align
    });
    body(s, LOREM.integer, { x: t.x, y: t.y + 0.345, w: 3.001, h: 0.673, color: C.ink, align: t.align, ls: 1.5 });
  });
  chevrons(s, { x: -0.153, y: 1.46, w: 1.426, h: 0.774, color: C.gold });
}

// 18 — Data table: 4 columns, 8 data rows plus a TOTAL row.
const TABLE_ROWS = [
  ['8', '423', '12'], ['8', '856', '12'], ['9', '343', '12'], ['7', '565', '34'],
  ['7', '0', '36'], ['0', '323', '80'], ['6', '55', '12'], ['3', '58', '12'],
];
function slide18(s) {
  const headFills = [C.khakiMid, C.orange, C.cocoa, C.brick];
  const cellFills = [C.white, C.orangeDark, C.clay, C.rose];
  const rowBorder = [{ type: 'none' }, { type: 'none' }, { color: C.grey35, pt: 0.5 }, { type: 'none' }];
  const rows = [
    ['Product Title', 'Value A', 'Value B', 'Value C'].map((t, i) => ({
      text: t, options: { fill: fill(headFills[i]), color: C.white, fontFace: HEAD,
        fontSize: i === 0 ? 14 : 12, align: 'center', valign: 'middle', border: [{ type: 'none' }] }
    })),
  ];
  TABLE_ROWS.forEach(vals => {
    rows.push([
      { text: LOREM.row, options: { fill: fill(C.white), color: C.ink, align: 'left', border: rowBorder } },
      ...vals.map((v, i) => ({
        text: v,
        options: { fill: fill(cellFills[i + 1], 70), color: C.ink, align: 'center', border: rowBorder }
      })),
    ]);
  });
  rows.push([
    { text: 'TOTAL', options: { fill: fill(C.khakiMid), color: C.white, bold: true, fontFace: HEAD, align: 'left', border: [{ type: 'none' }] } },
    ...['46', '2345', '457'].map(v => ({
      text: v, options: { fill: fill(C.olive), color: C.white, bold: true, align: 'center', border: [{ type: 'none' }] }
    })),
  ]);
  title(s, 'Table Slide', { x: 3.252, y: 0.8, w: 6.83, h: 0.774, align: 'center' });
  s.addTable(rows, {
    x: 1.697, y: 1.792, w: 9.94, colW: [5.487, 1.484, 1.484, 1.484],
    rowH: [0.697].concat(new Array(9).fill(0.44)),
    fontFace: BODY, fontSize: 12, color: C.ink, valign: 'middle'
  });
  chevrons(s, { x: -0.153, y: 1.46, w: 1.426, h: 0.774, color: C.gold });
}

// 19 — Two donut charts on the left, an area chart plus KPI on the right.
function slide19(s) {
  const donutOpts = { holeSize: 75, showLegend: false, showTitle: false, dataBorder: { pt: 0, color: C.white } };
  const donutData = [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: [30, 70] }];
  title(s, 'Data Visual Infographic', { x: 1.3, y: 1.274, w: 6.044, h: 1.447 });
  [{ x: 1.243, ring: C.orange, tx: 1.558, label: C.orange, pct: C.clay },
   { x: 4.108, ring: C.cocoa, tx: 4.424, label: C.cocoa, pct: C.grey40 }].forEach(d => {
    s.addChart('doughnut', donutData, Object.assign({
      x: d.x, y: 2.971, w: 2.148, h: 2.626, chartColors: ['EFEFEF', d.ring]
    }, donutOpts));
    txt(s, '70%', {
      x: d.tx, y: 4.013, w: 1.516, h: 0.572, fontFace: BODY, fontSize: 28, bold: true,
      color: d.pct, align: 'center', valign: 'middle'
    });
    txt(s, 'Lorem Ipsum', {
      x: d.x - 0.245, y: d.x < 3 ? 5.39 : 5.383, w: 2.637, h: 0.37, fontFace: BODY,
      fontSize: 16, bold: true, color: d.label, align: 'center'
    });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit', {
      x: d.x - 0.245, y: d.x < 3 ? 5.726 : 5.719, w: 2.637, h: 0.562, size: 11, align: 'center'
    });
  });
  txt(s, 'Lorem Ipsum', {
    x: 7.344, y: 1.308, w: 2.568, h: 0.37, fontFace: BODY, fontSize: 16, bold: true,
    color: C.black
  });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Mauris mattis libero id felis rhoncus mattis. ', {
    x: 7.344, y: 1.638, w: 4.503, h: 0.604
  });
  s.addShape('line', {
    x: 7.421, y: 2.58, w: 4.425, h: 0,
    line: { color: C.greyMid, width: 1, dashType: 'dash', transparency: 50 }
  });
  txt(s, '$340,214,000', {
    x: 7.344, y: 2.678, w: 3.538, h: 0.745, fontFace: HEAD, fontSize: 32, color: C.black,
    valign: 'bottom', paraSpaceAfter: 6
  });
  body(s, 'Lorem Ipsum', { x: 7.344, y: 3.342, w: 2.231, h: 0.32, size: 11 });
  s.addChart('area', [
    { name: 'Series 1', labels: ['Data 2', 'Data 3', 'Data 4', 'Data 5', 'Data 6'], values: [2.5, 3.5, 4.5, 3, 1] },
    { name: 'Series 2', labels: ['Data 2', 'Data 3', 'Data 4', 'Data 5', 'Data 6'], values: [4.4, 1.8, 2.8, 5, 4] },
  ], {
    x: 7.351, y: 3.858, w: 5.155, h: 2.626, chartColors: [C.cocoa, C.orange],
    showLegend: false, showTitle: false, catGridLine: { style: 'none' }, valGridLine: { style: 'none' },
    catAxisLineColor: C.greySoft, valAxisLineShow: false,
    catAxisLabelFontFace: BODY, catAxisLabelFontSize: 10, catAxisLabelColor: C.black,
    valAxisLabelFontFace: BODY, valAxisLabelFontSize: 10, valAxisLabelColor: C.black
  });
  chevrons(s, { x: 11.617, y: 0.443, w: 1.426, h: 0.774, color: C.gold });
}

// 20 — Exploded pie left, three mini gauge rings right.
function slide20(s) {
  // Exploded pie built from four `pie` wedges of decreasing radius.
  // Smallest first: each larger wedge is stacked on top of the previous one,
  // so only the leading slice of each remains visible.
  const wedges = [
    { x: 1.537, y: 1.716, d: 3.776, color: C.charcoal, ang: [92.0, 310.1] },
    { x: 1.349, y: 1.527, d: 4.154, color: C.brick, ang: [147.1, 310.1] },
    { x: 1.141, y: 1.32, d: 4.569, color: C.orange, ang: [192.5, 310.1] },
    { x: 0.912, y: 1.091, d: 5.026, color: C.olive, ang: [270.0, 120.3] },
  ];
  wedges.forEach(w => s.addShape('pie', {
    x: w.x, y: w.y, w: w.d, h: w.d, angleRange: w.ang, fill: fill(w.color), line: noLine
  }));
  drawPaths(s, PAW_TOE_L, { x: 4.275, y: 3.08, w: 0.172, h: 0.218, color: C.white });
  drawPaths(s, PAW_TOE_OL, { x: 4.1, y: 3.26, w: 0.188, h: 0.207, color: C.white });
  drawPaths(s, PAW_TOE_R, { x: 4.49, y: 3.08, w: 0.172, h: 0.218, color: C.white });
  drawPaths(s, PAW_TOE_OR, { x: 4.63, y: 3.259, w: 0.206, h: 0.207, color: C.white });
  drawPaths(s, PAW_PAD, { x: 4.232, y: 3.341, w: 0.472, h: 0.343, color: C.white });
  [['65%', 3.883, 3.912, 1.17, 32], ['20%', 1.993, 2.372, 1.142, 32], ['10%', 1.795, 3.703, 0.651, 16]]
    .forEach(([t, x, y, w, sz]) => txt(s, t, {
      x, y, w, h: sz === 32 ? 0.64 : 0.37, fontFace: HEAD, fontSize: sz, bold: true,
      color: C.white, align: 'center'
    }));
  title(s, 'Pie Chart Data', { x: 6.577, y: 1.271, w: 5.087, h: 0.774 });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient', {
    x: 6.587, y: 2.338, w: 5.834, h: 0.994, size: 14
  });
  const gauges = [
    { x: 6.587, color: C.olive, ang: [41.0, 270.0], pct: '65%', tw: 0.747, tx: 6.922, lx: 6.497 },
    { x: 8.551, color: C.orange, ang: [211.6, 287.1], pct: '20%', tw: 0.73, tx: 8.886, lx: 8.46 },
    { x: 10.514, color: C.brick, ang: [248.8, 296.8], pct: '10%', tw: 0.681, tx: 10.874, lx: 10.424 },
  ];
  gauges.forEach(g => {
    s.addShape('ellipse', { x: g.x, y: 4.067, w: 1.401, h: 1.401, fill: fill(C.greySoft, 50), line: noLine });
    s.addShape('pie', { x: g.x + 0.064, y: 4.13, w: 1.273, h: 1.273, angleRange: g.ang, fill: fill(g.color), line: noLine });
    s.addShape('ellipse', { x: g.x + 0.251, y: 4.317, w: 0.899, h: 0.899, fill: fill(C.white), line: noLine });
    txt(s, g.pct, {
      x: g.tx, y: 4.606, w: g.tw, h: 0.404, fontFace: HEAD, fontSize: 18, bold: true,
      color: C.black, align: 'center'
    });
    txt(s, 'Sample text', {
      x: g.lx, y: 5.549, w: 1.582, h: 0.386, fontFace: HEAD, fontSize: 14, color: C.black,
      align: 'center', lineSpacingMultiple: 1.3
    });
  });
  s.addShape('rect', { x: 7.739, y: 4.117, w: 0.365, h: 0.365, fill: fill(C.grey40), line: noLine });
  s.addShape('triangle', { x: 7.868, y: 4.256, w: 0.107, h: 0.076, fill: fill(C.white), line: noLine });
  stripeBar(s, { x: 9.875, y: 6.483, w: 3.513, h: 0.348, color: C.clayDark });
}

// 21 — Gantt: week header table above a 7-step schedule grid.
const GANTT = [
  ['Step One', [[1, 3, C.brick]]],
  ['Step Two', [[4, 6, C.brick], [17, 19, C.brick]]],
  ['Step Three', [[6, 14, C.khaki]]],
  ['Step Four', [[11, 20, C.orange]]],
  ['Step Five', [[21, 23, C.khaki]]],
  ['Step Six', [[13, 14, C.orange], [22, 26, C.brick]]],
  ['Step Seven', [[26, 28, C.grey]]],
];
function slide21(s) {
  const idle = fill('696464', 85);
  const noB = [{ type: 'none' }];
  const weekCell = t => ({ text: t, options: { colspan: 7, fill: fill(C.orangeDeep), color: C.white, bold: true, fontSize: 15, align: 'center', valign: 'middle', border: noB } });
  const dayHdr = [];
  for (let w = 0; w < 4; w++) {
    for (let d = 1; d <= 7; d++) {
      dayHdr.push({ text: String(d), options: { fill: fill(C.orange, 15), color: C.white, fontSize: 15, align: 'center', valign: 'middle', border: noB } });
    }
  }
  title(s, 'Agenda Schedule', { x: 2.667, y: 0.741, w: 8, h: 0.774, align: 'center' });
  s.addTable([
    [{ text: 'Month', options: { colspan: 28, fill: fill(C.orangeDark), color: C.white, bold: true, fontSize: 15, align: 'center', valign: 'middle', border: noB } }],
    [weekCell('1st Week'), weekCell('2nd Week'), weekCell('3rd Week'), weekCell('4th Week')],
    dayHdr,
  ], {
    x: 2.048, y: 1.983, w: 10.396, colW: new Array(28).fill(0.371), rowH: [0.371, 0.371, 0.366],
    fontFace: BODY
  });
  s.addTable(GANTT.map(([label, bars]) => {
    const row = [{ text: label, options: { fill: fill(C.clayDark, 35), color: C.white, fontSize: 12, align: 'left', valign: 'middle', border: noB } }];
    for (let c = 1; c <= 28; c++) {
      const hit = bars.find(b => c >= b[0] && c <= b[1]);
      row.push({ text: '', options: { fill: hit ? fill(hit[2]) : idle, border: noB } });
    }
    return row;
  }), {
    x: 0.772, y: 3.099, w: 11.672,
    colW: [1.265].concat(new Array(28).fill(0.3717)),
    rowH: new Array(7).fill(0.334), fontFace: BODY
  });
  [[C.orange, 3.529, 3.827], [C.khaki, 5.16, 5.458], [C.grey, 6.824, 7.122], [C.brick, 8.515, 8.812]]
    .forEach(([color, sx, tx]) => {
      s.addShape('rect', { x: sx, y: 5.947, w: 0.234, h: 0.234, fill: fill(color), line: noLine });
      txt(s, 'Description', {
        x: tx, y: 5.901, w: 1.15, h: 0.325, fontFace: BODY, fontSize: 13.3, color: C.black
        });
    });
  chevrons(s, { x: -0.153, y: 1.46, w: 1.426, h: 0.774, color: C.gold });
}

// 22 — A/B comparison panel with pictogram rows and headline percentages.
function slide22(s) {
  const sides = [
    { x: 1.273, bx: 3.707, color: C.brick, letter: 'A', lx: 2.938, tx: 1.783, fig: PERSON_M,
      figW: 0.324, rowW: 0.157, filled: 3, pct: '30%', px: 4.969 },
    { x: 6.667, bx: 9.102, color: C.orange, letter: 'B', lx: 8.332, tx: 7.177, fig: PERSON_F,
      figW: 0.29, rowW: 0.157, filled: 8, pct: '80%', px: 10.352 },
  ];
  title(s, 'Comparison Slide', { x: 2.938, y: 0.889, w: 7.554, h: 0.774, align: 'center' });
  body(s, LOREM.aenean, { x: 2.179, y: 1.8, w: 9.071, h: 0.599, align: 'center' });
  s.addShape('rect', { x: 1.273, y: 3.495, w: 10.789, h: 3.018, fill: fill(C.greyPale), line: noLine });
  sides.forEach(sd => {
    s.addShape('rect', { x: sd.x, y: 2.844, w: 5.394, h: 0.651, fill: fill(sd.color), line: noLine });
    s.addShape('flowChartMerge', { x: sd.bx, y: 3.495, w: 0.525, h: 0.212, fill: fill(sd.color), line: noLine });
    txt(s, sd.letter, {
      x: sd.lx, y: sd.x < 5 ? 2.917 : 2.91, w: 2.064, h: 0.46, fontFace: HEAD, fontSize: 18,
      bold: true, color: C.white, align: 'center', paraSpaceAfter: 6
    });
    body(s, LOREM.aeneanShort, { x: sd.tx, y: 4.005, w: 4.372, h: 0.863, align: 'center' });
    drawPaths(s, sd.fig, { x: sd.tx, y: 5.149, w: sd.figW, h: sd.figW === 0.324 ? 0.781 : 0.778, color: sd.color });
    for (let i = 0; i < 10; i++) {
      drawPaths(s, sd.fig, {
        x: sd.tx + 0.654 + i * (sd.x < 5 ? 0.159 : 0.155), y: 5.532, w: 0.157, h: 0.394,
        color: i < sd.filled ? sd.color : C.greySoft
      });
    }
    txt(s, sd.pct, {
      x: sd.px, y: sd.x < 5 ? 5.23 : 5.22, w: 1.381, h: 0.896, fontFace: HEAD, fontSize: 40,
      bold: true, color: sd.color, paraSpaceAfter: 6
    });
  });
  s.addShape('line', { x: 6.667, y: 3.962, w: 0, h: 1.7, line: { color: C.grey35, width: 1, dashType: 'lgDash' } });
  chevrons(s, { x: -0.153, y: 1.46, w: 1.426, h: 0.774, color: C.gold });
}

// 23 — Four upward arrows of different heights plus flanking stats.
function slide23(s) {
  const arrows = [
    { x: 4.219, y: 4.884, h: 2.616, ty: 3.887, color: C.brickDark, num: '01', ny: 4.441, lx: 3.9, lw: 1.847, ly: 6.085, label: 'Product One', th: 1.018 },
    { x: 5.401, y: 3.69, h: 3.81, ty: 2.686, color: C.olive, num: '02', ny: 3.273, lx: 5.096, lw: 1.847, ly: 6.085, label: 'Product Two', th: 1.018 },
    { x: 6.582, y: 4.438, h: 3.062, ty: 3.474, color: C.orange, num: '03', ny: 4.057, lx: 6.009, lw: 2.356, ly: 5.83, label: 'Product Three', th: 0.963 },
    { x: 7.764, y: 3.474, h: 4.026, ty: 2.484, color: C.clay, num: '04', ny: 3.09, lx: 7.432, lw: 1.847, ly: 6.085, label: 'Product Four', th: 1.018 },
  ];
  const stats = [
    { x: 0.854, y: 3.457, align: 'right', num: '3456', color: C.brickDark },
    { x: 0.854, y: 4.863, align: 'right', num: '11976', color: C.olive },
    { x: 9.774, y: 3.457, align: 'left', num: '9976', color: C.orange },
    { x: 9.774, y: 4.863, align: 'left', num: '12976', color: C.clay },
  ];
  title(s, 'Creative Infographic', { x: 2.388, y: 0.65, w: 8.558, h: 0.774, align: 'center' });
  body(s, LOREM.aenean, { x: 2.179, y: 1.616, w: 9.071, h: 0.599, align: 'center' });
  arrows.forEach(a => {
    s.addShape('rect', { x: a.x, y: a.y, w: 1.182, h: a.h, fill: fill(a.color), line: noLine });
    s.addShape('triangle', { x: a.x - 0.196, y: a.ty, w: 1.574, h: a.th, fill: fill(a.color), line: noLine });
    txt(s, a.num, {
      x: a.x + 0.222, y: a.ny, w: 0.738, h: 0.553, fontFace: HEAD, fontSize: 24, bold: true,
      color: C.white, align: 'center', lineSpacingMultiple: 1.2
    });
    txt(s, a.label, {
      x: a.lx, y: a.ly, w: a.lw, h: 0.365, rotate: 270, fontFace: BODY, fontSize: 14,
      color: C.white, lineSpacingMultiple: 1.2
    });
  });
  stats.forEach(st => {
    txt(s, st.num, {
      x: st.x, y: st.y, w: 2.637, h: 0.404, fontFace: BODY, fontSize: 18, bold: true,
      color: st.color, align: st.align
    });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit ', {
      x: st.x, y: st.y + 0.337, w: 2.637, h: 0.602, align: st.align
    });
  });
  chevrons(s, { x: -0.153, y: 1.46, w: 1.426, h: 0.774, color: C.gold });
}

// 24 — SWOT diamonds with leader lines to four labelled paragraphs.
function slide24(s) {
  const diamonds = [
    { x: 5.548, y: 2.205, color: C.clayDark, letter: 'S', lx: 6.209, lw: 0.882, sz: 80 },
    { x: 4.277, y: 3.411, color: C.brick, letter: 'W', lx: 4.75, lw: 1.257, sz: 72 },
    { x: 6.819, y: 3.411, color: C.orange, letter: 'T', lx: 7.495, lw: 0.886, sz: 80 },
    { x: 5.548, y: 4.616, color: C.olive, letter: 'O', lx: 6.159, lw: 1.045, sz: 80 },
  ];
  const notes = [
    { x: 1.378, y: 1.997, label: 'Strengths', color: C.clayDark },
    { x: 1.378, y: 4.291, label: 'Weakness', color: C.brick },
    { x: 9.071, y: 2.026, label: 'Threaths', color: C.orange },
    { x: 9.071, y: 4.316, label: 'Opportunities', color: C.olive },
  ];
  const leaders = [
    { x: 2.992, y: 5.648, w: 2.404, dot: 2.964, dy: 5.623 },
    { x: 4.277, y: 2.205, w: 2.39, dot: 4.249, dy: 2.177 },
    { x: 7.744, y: 5.955, w: 2.237, dot: 9.953, dy: 5.705 },
    { x: 7.938, y: 3.633, w: 2.237, dot: 10.147, dy: 3.383 },
  ];
  title(s, 'SWOT Infographic', { x: 2.388, y: 0.873, w: 8.558, h: 0.774, align: 'center' });
  leaders.forEach(l => {
    s.addShape('line', { x: l.x, y: l.y, w: l.w, h: 0, line: { color: C.greyMid, width: 1 } });
    s.addShape('ellipse', { x: l.dot, y: l.dy, w: 0.056, h: 0.056, fill: fill(C.orange), line: noLine });
  });
  diamonds.forEach(d => {
    s.addShape('diamond', { x: d.x, y: d.y, w: 2.237, h: 2.237, fill: fill(d.color), line: noLine });
    txt(s, d.letter, {
      x: d.lx, y: d.y + 0.395, w: d.lw, h: 1.447, fontFace: BODY, fontSize: d.sz, color: C.white,
      wrap: false
    });
  });
  notes.forEach(n => {
    txt(s, n.label, {
      x: n.x, y: n.y, w: 2.779, h: 0.37, fontFace: HEAD, fontSize: 16, bold: true,
      color: n.color, align: 'center'
    });
    body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ', {
      x: n.x, y: n.y + 0.413, w: 2.779, h: 0.692, size: 14, align: 'center'
    });
  });
  chevrons(s, { x: 11.617, y: 0.443, w: 1.426, h: 0.774, color: C.gold });
}

// 25 — Closing contact slide over the dark photo panel.
function slide25(s) {
  const contacts = [
    { y: 3.162, icon: ENVELOPE, ix: 1.856, iy: 3.324, iw: 0.372, ih: 0.372, ty: 3.137,
      lines: ['equestrian@mail.com', '2829-347'] },
    { y: 4.313, icon: PIN, ix: 1.879, iy: 4.494, iw: 0.327, ih: 0.372, ty: 4.292,
      lines: ['1066 Summit Park Avenue Southfield, MI 48034'] },
    { y: 5.456, icon: PHONE, ix: 1.849, iy: 5.637, iw: 0.387, ih: 0.372, ty: 5.434,
      lines: ['091 \u2013 900', '2003 \u2013 9485- 0293'] },
  ];
  scrim(s, { x: 0, y: 0, w: 13.388, h: 7.5 });
  txt(s, 'Get in Touch With Us!', {
    x: 1.656, y: 1.478, w: 10.076, h: 0.909, fontFace: HEAD, fontSize: 48, color: C.white,
    align: 'center'
  });
  contacts.forEach(c => {
    s.addShape('roundRect', { x: 1.656, y: c.y, w: 0.772, h: 0.734, rectRadius: SQUARE, fill: fill(C.orange), line: noLine });
    drawPaths(s, c.icon, { x: c.ix, y: c.iy, w: c.iw, h: c.ih, color: C.white });
    txt(s, c.lines.map((t, i) => ({ text: t, options: { breakLine: i < c.lines.length - 1 } })), {
      x: 2.582, y: c.ty, w: 3.289, h: 0.776, fontFace: HEAD, fontSize: 16, color: C.white,
      lineSpacingMultiple: 1.3
    });
  });
  drawPaths(s, ENVELOPE.slice(0, 0), { x: 0, y: 0, w: 0.01, h: 0.01, color: C.white });
  s.addShape('rect', { x: 1.966, y: 3.391, w: 0.153, h: 0.022, fill: fill(C.white), line: noLine });
  s.addShape('rect', { x: 1.966, y: 3.435, w: 0.153, h: 0.022, fill: fill(C.white), line: noLine });
  s.addShape('rect', { x: 1.966, y: 3.479, w: 0.087, h: 0.022, fill: fill(C.white), line: noLine });
  drawPaths(s, PIN_DOT, { x: 1.983, y: 4.56, w: 0.113, h: 0.109, color: C.orange });
  drawPaths(s, HORSE, { x: 8.103, y: 3.559, w: 2.772, h: 1.757, color: C.white });
  chevrons(s, { x: 11.013, y: 6.211, w: 1.572, h: 0.853, color: C.white });
}

// ==================================================================== main ==
const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
  slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18,
  slide19, slide20, slide21, slide22, slide23, slide24, slide25,
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'CUSTOM_16x9', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'CUSTOM_16x9';
  pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
  pptx.title = 'Equestrian Sport Presentation Template';

  BUILDERS.forEach((builder, i) => {
    const slide = pptx.addSlide();
    slide.background = { color: C.white };
    chrome(slide, i + 1);
    builder(slide);
  });

  return pptx.writeFile({
    fileName: path.join(__dirname, '0edf8f95-d3ca-4c16-8fd3-730b1a3067a2_grok_final.pptx')
  });
}

build().then(f => console.log('wrote ' + f)).catch(err => {
  console.error(err);
  process.exit(1);
});
