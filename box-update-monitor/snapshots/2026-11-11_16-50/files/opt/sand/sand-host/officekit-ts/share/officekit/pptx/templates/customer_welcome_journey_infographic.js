/*
 * Customer Welcome Journey Infographic - 20 slide deck rebuilt with pptxgenjs.
 *
 * Run:  node 0bca4cf4-18e3-49a2-952f-38972b23d61b_grok_final.js
 * Writes 0bca4cf4-18e3-49a2-952f-38972b23d61b_grok_final.pptx next to this file.
 *
 * Raster icons from the source deck are drawn as outlined placeholder boxes
 * tinted with the icon's dominant colour (see `icon()`).
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------------------
// Theme
// ---------------------------------------------------------------------------
const PINK = 'EC2B8C';   // accent1
const LILAC = '9477E2';  // accent2
const VIOLET = '6A37C7'; // accent3
const CYAN = '19C8EB';   // accent4
const PLUM = '892293';   // accent5
const MAGENTA = 'CB0FB0'; // accent6
const INK = '2D3847';    // dark 2
const GREY_TX = '404040';
const GREY_BG = 'F2F2F2';
const GREY_LN = 'D9D9D9';
const GREY_LN2 = 'BFBFBF';
const WHITE = 'FFFFFF';
const PAPER = 'FDFDFD';

const HEAD = 'Montserrat SemiBold'; // theme major font
const BODY = 'Montserrat Light';    // theme minor font

const ACCENTS = [PINK, LILAC, VIOLET, CYAN, PLUM, MAGENTA];

// Dominant colours of the icon artwork that was replaced by placeholders.
const IC_WHITE = WHITE;
const IC_BLUE = '3A86FF';
const IC_ROSE = 'FF006E';
const IC_AMBER = 'FFBE0B';
const IC_PURPLE = '8338EC';
const IC_ORANGE = 'FB5607';
const IC_GREEN = '57CC99';
const IC_GREY = GREY_TX;

const CARD_SHADOW = { type: 'outer', blur: 35, offset: 10, angle: 50, color: '000000', opacity: 0.05 };

// Day-1 checklist, reused on slides 15 and 20.
const DAY1_TASKS = ['Confirm their\nsign-up', 'Send a\nwelcome email', 'Help them get\nstarted'];

/** Arrow band of slide 8: notched tail, pointed head (prstGeom notchedRightArrow). */
const P_ARROWBAND = [
  ['M',0,0], ['L',0.906,0], ['L',1,0.5], ['L',0.906,1],
  ['L',0,1], ['L',0.094,0.5], ['Z'],
];

// ---------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------

/** Convert a compact ['M'|'L'|'C'|'Z', ...coords] path into pptxgenjs points. */
function geom(cmds, x, y, w, h) {
  const pts = [];
  for (const c of cmds) {
    const k = c[0];
    if (k === 'M') pts.push({ x: x + c[1] * w, y: y + c[2] * h, moveTo: true });
    else if (k === 'L') pts.push({ x: x + c[1] * w, y: y + c[2] * h });
    else if (k === 'C') pts.push({
      x: x + c[5] * w, y: y + c[6] * h,
      curve: { type: 'cubic', x1: x + c[1] * w, y1: y + c[2] * h, x2: x + c[3] * w, y2: y + c[4] * h },
    });
    else if (k === 'Z') pts.push({ close: true });
  }
  return pts;
}

/** Draw a custom-geometry shape from a normalised path at x/y/w/h. */
function freeform(slide, cmds, x, y, w, h, opts = {}) {
  slide.addShape('custGeom', Object.assign({ x, y, w, h, points: geom(cmds, 0, 0, w, h) }, opts));
}

function rect(slide, x, y, w, h, color, opts = {}) {
  slide.addShape('rect', Object.assign({ x, y, w, h, fill: { color } }, opts));
}

function roundRect(slide, x, y, w, h, color, radius, opts = {}) {
  slide.addShape('roundRect', Object.assign(
    { x, y, w, h, rectRadius: radius }, color ? { fill: { color } } : {}, opts));
}

function oval(slide, x, y, w, h, color, opts = {}) {
  slide.addShape('ellipse', Object.assign({ x, y, w, h }, color ? { fill: { color } } : {}, opts));
}

/** Circle centred on (cx, cy). */
function dot(slide, cx, cy, d, color, opts = {}) {
  oval(slide, cx - d / 2, cy - d / 2, d, d, color, opts);
}

function line(slide, x, y, w, h, color, width, opts = {}) {
  slide.addShape('line', { x, y, w, h, line: Object.assign({ color, width }, opts) });
}

/** White disc with the deck's soft drop shadow. */
function disc(slide, x, y, d) {
  oval(slide, x, y, d, d, WHITE, { shadow: CARD_SHADOW });
}

/** Placeholder standing in for a raster icon of the source deck. */
function icon(slide, x, y, w, h, color) {
  slide.addShape('roundRect', {
    x, y, w, h, rectRadius: 0.18,
    fill: { color: WHITE, transparency: 100 },
    line: { color: color || IC_GREY, width: 1 },
  });
}

/** Text box; defaults mirror the source deck (top anchored, Montserrat). */
function text(slide, body, o) {
  slide.addText(body, Object.assign({ fontFace: HEAD, valign: 'top', color: GREY_TX }, o));
}

/** Small caption paragraph (Montserrat Light, 11pt, 120% leading). */
function caption(slide, body, x, y, w, h, o = {}) {
  text(slide, body, Object.assign(
    { x, y, w, h, fontFace: BODY, fontSize: 11, color: GREY_TX, align: 'center', lineSpacingMultiple: 1.2 }, o));
}

/**
 * Start a slide: paper background, the two-tone headline every page carries,
 * and the master furniture (page number flanked by two pink dots).
 */
function newSlide(pres) {
  const slide = pres.addSlide();
  slide.background = { color: PAPER };
  slide.slideNumber = { x: 6.358, y: 6.908, w: 0.616, h: 0.37,
    fontFace: HEAD, fontSize: 16, color: PINK, align: 'center' };
  oval(slide, 6.301, 7.049, 0.089, 0.089, 'FBD5E8');
  oval(slide, 6.943, 7.049, 0.089, 0.089, 'FBD5E8');
  slide.addText(
    [{ text: 'Customer Welcome ', options: { color: PINK } },
     { text: 'Journey Infographic', options: { color: GREY_TX } }],
    { x: 2, y: 0.521, w: 9.333, h: 0.572, align: 'center', valign: 'top', fontFace: HEAD, fontSize: 28 });
  return slide;
}

/**
 * Bulleted list. Every run needs its own `bullet` (pptxgenjs only inherits the
 * shape-level one onto the first run); `\n` inside an item becomes a soft break
 * so wrapped lines stay inside one bullet.
 */
function bulletList(slide, items, x, y, w, h, o = {}) {
  const marker = Object.assign({ characterCode: '006F', indent: 13.5 }, o.marker || {});
  const runs = [];
  items.forEach(item => {
    item.split('\n').forEach((ln, i) =>
      runs.push({ text: ln, options: { bullet: marker, align: 'left', softBreakBefore: i > 0 } }));
    runs[runs.length - 1].options.breakLine = true;
  });
  text(slide, runs, Object.assign({ x, y, w, h, fontFace: BODY, fontSize: 9, color: GREY_TX,
    align: 'left', lineSpacingMultiple: 1.2 }, o, { marker: undefined }));
}

// ---------------------------------------------------------------------------
// Custom geometry, normalised to a 0..1 box (traced from the source deck)
// ---------------------------------------------------------------------------
/** Rounded tag with a circular cap on the left (slide 2). */
const P_PILL = [
  ['M',0.321,0], ['C',0.455,0,0.569,0.126,0.618,0.305], ['L',0.62,0.319], ['L',1,0.319], ['L',1,0.681],
  ['L',0.62,0.681], ['L',0.618,0.695], ['C',0.569,0.874,0.455,1,0.321,1], ['C',0.144,1,0,0.776,0,0.5], ['C',0,0.224,0.144,0,0.321,0],
  ['Z'],
];
/** Rosette petal / map pin (slide 3). */
const P_PETAL = [
  ['M',0.553,0.015], ['C',0.691,-0.032,0.839,0.036,0.926,0.192], ['L',0.823,0.22], ['L',0.823,0.22], ['L',0.843,0.214],
  ['L',0.926,0.192], ['L',0.951,0.242], ['L',0.951,0.242], ['C',1.051,0.475,0.992,0.773,0.821,0.908], ['C',0.649,1.042,0,0.99,0,0.99],
  ['L',0.38,0.692], ['L',0.41,0.667], ['L',0.41,0.667], ['L',0.401,0.674], ['L',0,0.99],
  ['C',0,0.99,0.097,0.724,0.213,0.475], ['L',0.263,0.371], ['L',0.263,0.371], ['C',0.331,0.232,0.403,0.112,0.461,0.066], ['C',0.49,0.042,0.521,0.025,0.553,0.015],
  ['Z'],
];
/** Journey-bar segment: notch on the left, round cap right (slide 5). */
const P_TAB = [
  ['M',0,0], ['L',0.229,0], ['L',0.635,0], ['L',0.738,0], ['C',0.883,0,1,0.224,1,0.5],
  ['C',1,0.776,0.883,1,0.738,1], ['L',0.635,1], ['L',0.074,1], ['L',0,1], ['L',0.025,0.975],
  ['C',0.104,0.872,0.157,0.698,0.157,0.5], ['C',0.157,0.302,0.104,0.128,0.025,0.025], ['Z'],
];
/** First journey-bar segment: rounded head, chevron tail (slide 5). */
const P_CAP = [
  ['M',0,0.914], ['L',0,0.259], ['C',0,0.116,0.224,0,0.5,0], ['C',0.776,0,1,0.116,1,0.259], ['L',1,0.914],
  ['L',0.987,0.947], ['C',0.97,0.968,0.938,0.985,0.898,0.993], ['L',0.833,1], ['L',0.167,1], ['L',0.102,0.993],
  ['C',0.062,0.985,0.03,0.968,0.013,0.947], ['Z'],
];
/** Funnel band 1 - lit face (slide 6). */
const P_FUNNEL_TOP = [
  ['M',0.5,0.272], ['C',0.294,0.272,0.112,0.165,0,0], ['L',0.094,0.696], ['C',0.153,0.873,0.313,1,0.5,1], ['C',0.687,1,0.847,0.873,0.906,0.696],
  ['L',1,0], ['C',0.887,0.165,0.706,0.272,0.5,0.272], ['Z'],
];
/** Funnel band 2 - lit face (slide 6). */
const P_FUNNEL_B1 = [
  ['M',0.907,0.612], ['L',1,0], ['C',0.926,0.194,0.731,0.333,0.5,0.333], ['C',0.27,0.333,0.074,0.194,0,0], ['L',0.092,0.612],
  ['L',0.093,0.612], ['L',0.093,0.617], ['L',0.108,0.719], ['C',0.157,0.88,0.314,1,0.5,1], ['C',0.686,1,0.843,0.882,0.891,0.719],
  ['L',0.907,0.617], ['L',0.907,0.612], ['Z'],
];
/** Funnel band 3 - lit face (slide 6). */
const P_FUNNEL_B2 = [
  ['M',0.5,0.329], ['C',0.262,0.329,0.062,0.191,0,0], ['L',0.128,0.775], ['C',0.185,0.906,0.33,1,0.5,1], ['C',0.67,1,0.815,0.908,0.872,0.775],
  ['L',1,0], ['C',0.938,0.191,0.738,0.329,0.5,0.329], ['Z'],
];
/** Funnel band 4 - lit face (slide 6). */
const P_FUNNEL_B3 = [
  ['M',0.5,0.252], ['C',0.272,0.252,0.076,0.149,0,0], ['L',0.165,0.84], ['C',0.224,0.934,0.354,1,0.505,1], ['C',0.645,1,0.768,0.943,0.831,0.86],
  ['L',1,0.003], ['C',0.924,0.149,0.728,0.252,0.5,0.252], ['Z'],
];
/** Funnel band 5 - lit face (slide 6). */
const P_FUNNEL_B4 = [
  ['M',0.286,0.949], ['C',0.337,0.98,0.416,1,0.506,1], ['C',0.591,1,0.666,0.983,0.717,0.955], ['L',1,0.02], ['C',0.905,0.101,0.721,0.157,0.51,0.157],
  ['C',0.283,0.157,0.088,0.093,0,0], ['L',0.286,0.949], ['Z'],
];
/** Funnel band 1 - shaded underside (slide 6). */
const P_FUNNEL_S1 = [
  ['M',0.896,0.407], ['L',0.097,1], ['L',0,0.619], ['L',1,0], ['Z'],
];
/** Funnel band 2 - shaded underside (slide 6). */
const P_FUNNEL_S2 = [
  ['M',0.879,0.423], ['L',0.113,1], ['L',0,0.609], ['L',1,0], ['Z'],
];
/** Funnel band 3 - shaded underside (slide 6). */
const P_FUNNEL_S3 = [
  ['M',0.817,0.491], ['L',0.119,1], ['L',0,0.628], ['L',1,0], ['Z'],
];
/** Funnel band 4 - shaded underside (slide 6). */
const P_FUNNEL_S4 = [
  ['M',0.798,0.399], ['L',0.229,1], ['L',0,0.549], ['L',1,0], ['Z'],
];
/** Small "next" chevron drawn inside the white pucks. */
const P_CHEVRON = [
  ['M',1,0.514], ['L',0.191,1], ['L',0.143,1], ['L',0.095,1], ['L',0.048,0.943],
  ['C',0,0.943,0,0.943,0,0.914], ['L',0.048,0.886], ['L',0.714,0.486], ['L',0.048,0.086], ['L',0,0.086],
  ['C',0,0.057,0,0.057,0.048,0.057], ['L',0.095,0], ['L',0.143,0], ['L',0.191,0], ['L',1,0.486],
  ['L',1,0.514], ['Z'],
];
/** Half loop of the S-curve ribbon (slide 9). */
const P_SCURVE = [
  ['M',0.147,0.854], ['C',0.242,0.948,0.367,1,0.5,1], ['L',0.5,0.765], ['C',0.43,0.765,0.363,0.737,0.313,0.686], ['C',0.264,0.637,0.236,0.571,0.236,0.499],
  ['C',0.236,0.429,0.264,0.363,0.313,0.312], ['C',0.363,0.263,0.43,0.235,0.5,0.235], ['C',0.57,0.235,0.637,0.263,0.687,0.312], ['C',0.736,0.363,0.764,0.429,0.764,0.499], ['L',1,0.499],
  ['C',1,0.366,0.948,0.241,0.853,0.146], ['C',0.76,0.052,0.633,0,0.5,0], ['C',0.367,0,0.242,0.052,0.147,0.146], ['C',0.052,0.241,0,0.366,0,0.499], ['C',0,0.634,0.052,0.759,0.147,0.854],
  ['Z'],
];
/** Grey hill silhouette across the foot of slide 15. */
const P_HILL = [
  ['M',0.514,0.001], ['C',0.666,0.015,0.817,0.26,0.952,0.735], ['L',0.977,0.827], ['L',1,0.917], ['L',1,1],
  ['L',0.844,1], ['L',0.835,0.974], ['C',0.623,0.386,0.376,0.385,0.163,0.978], ['L',0.156,1], ['L',0,1],
  ['L',0,0.918], ['L',0.034,0.787], ['L',0.072,0.655], ['C',0.209,0.204,0.362,-0.013,0.514,0.001], ['Z'],
];
/** Concentric ring 1 - outermost (slide 17). */
const P_RING01 = [
  ['M',0.87,0.265], ['C',0.741,0.061,0.47,0,0.265,0.13], ['C',0.061,0.259,0,0.53,0.13,0.735], ['C',0.259,0.939,0.53,1,0.735,0.87], ['C',0.939,0.741,1,0.47,0.87,0.265],
];
/** Concentric ring 2 (slide 17). */
const P_RING02 = [
  ['M',0.87,0.265], ['C',0.741,0.061,0.47,0,0.265,0.13], ['C',0.061,0.259,0,0.53,0.13,0.735], ['C',0.259,0.939,0.53,1,0.735,0.87], ['C',0.939,0.741,1,0.47,0.87,0.265],
];
/** Concentric ring 3 (slide 17). */
const P_RING03 = [
  ['M',0.87,0.265], ['C',0.741,0.06,0.47,0,0.265,0.13], ['C',0.061,0.259,0,0.53,0.13,0.735], ['C',0.259,0.939,0.53,1,0.735,0.87], ['C',0.939,0.741,1,0.47,0.87,0.265],
];
/** Concentric ring 4 - solid core (slide 17). */
const P_RING04 = [
  ['M',0.87,0.265], ['C',0.741,0.061,0.47,0,0.265,0.129], ['C',0.061,0.259,0,0.53,0.129,0.735], ['C',0.259,0.939,0.53,1,0.734,0.87], ['C',0.939,0.741,1,0.47,0.87,0.265],
];
/** Open ring finished with an arrow head (slide 18). */
const P_ARC_ARROW = [
  ['M',0.918,0], ['L',0.914,0.001], ['L',0.911,0.009], ['L',0.911,0.029], ['L',0.704,0.029],
  ['C',0.689,0.03,0.674,0.039,0.663,0.054], ['C',0.65,0.071,0.642,0.096,0.642,0.122], ['L',0.642,0.543], ['C',0.638,0.642,0.608,0.735,0.559,0.805], ['C',0.489,0.906,0.388,0.951,0.289,0.925],
  ['C',0.241,0.912,0.195,0.882,0.157,0.838], ['C',0.12,0.795,0.09,0.741,0.072,0.68], ['C',0.045,0.592,0.041,0.493,0.059,0.401], ['C',0.073,0.332,0.099,0.269,0.135,0.218], ['C',0.194,0.137,0.274,0.093,0.356,0.097],
  ['C',0.425,0.101,0.492,0.139,0.544,0.204], ['C',0.553,0.216,0.567,0.217,0.576,0.205], ['C',0.587,0.192,0.587,0.17,0.577,0.156], ['C',0.514,0.075,0.433,0.03,0.348,0.03], ['C',0.243,0.029,0.144,0.093,0.078,0.207],
  ['C',0.046,0.261,0.024,0.325,0.011,0.393], ['C',-0.004,0.472,-0.004,0.555,0.011,0.634], ['C',0.059,0.896,0.248,1.052,0.434,0.984], ['C',0.509,0.957,0.572,0.895,0.617,0.815], ['C',0.662,0.734,0.689,0.633,0.69,0.524],
  ['L',0.691,0.121], ['L',0.695,0.105], ['L',0.708,0.097], ['L',0.911,0.097], ['L',0.911,0.116],
  ['L',0.914,0.124], ['L',0.921,0.125], ['L',0.996,0.071], ['L',1,0.063], ['L',0.997,0.054],
  ['L',0.921,0.002], ['L',0.918,0], ['Z'], ['M',0.605,0.215], ['L',0.588,0.225],
  ['C',0.578,0.238,0.578,0.26,0.588,0.273], ['C',0.597,0.286,0.612,0.286,0.622,0.273], ['C',0.631,0.26,0.631,0.238,0.622,0.225], ['L',0.605,0.215], ['Z'],
];
/** Open ring without arrow head - last step (slide 18). */
const P_ARC_PLAIN = [
  ['M',0.504,0], ['C',0.353,-0.001,0.209,0.066,0.113,0.182], ['C',0.067,0.239,0.035,0.304,0.016,0.374], ['C',-0.005,0.456,-0.006,0.541,0.015,0.623], ['C',0.085,0.892,0.36,1.054,0.629,0.984],
  ['C',0.737,0.955,0.828,0.892,0.894,0.809], ['C',0.959,0.726,0.998,0.622,1,0.51], ['C',1.001,0.473,0.997,0.437,0.99,0.401], ['C',0.982,0.363,0.97,0.326,0.953,0.29], ['C',0.946,0.274,0.928,0.266,0.911,0.271],
  ['C',0.891,0.278,0.881,0.3,0.89,0.319], ['C',0.907,0.354,0.919,0.39,0.925,0.428], ['C',0.931,0.461,0.932,0.495,0.929,0.529], ['C',0.922,0.63,0.88,0.726,0.81,0.799], ['C',0.709,0.904,0.562,0.951,0.419,0.923],
  ['C',0.349,0.909,0.282,0.878,0.227,0.833], ['C',0.173,0.789,0.131,0.734,0.104,0.67], ['C',0.065,0.579,0.059,0.478,0.086,0.383], ['C',0.106,0.311,0.143,0.246,0.196,0.194], ['C',0.28,0.11,0.396,0.065,0.515,0.07],
  ['C',0.616,0.073,0.713,0.112,0.788,0.179], ['C',0.801,0.192,0.821,0.193,0.834,0.181], ['C',0.849,0.168,0.85,0.144,0.836,0.13], ['C',0.745,0.047,0.627,0.001,0.504,0], ['Z'],
  ['M',0.874,0.188], ['L',0.849,0.199], ['C',0.835,0.212,0.835,0.234,0.849,0.248], ['C',0.862,0.262,0.884,0.262,0.898,0.248], ['C',0.912,0.234,0.912,0.212,0.898,0.199],
  ['L',0.874,0.188], ['Z'],
];
/** Rounded number tile with a bite out of the top-left (slide 19). */
const P_NUMBOX = [
  ['M',0.793,0], ['L',0.341,0], ['C',0.341,0.188,0.188,0.341,0,0.341], ['L',0,0.793], ['C',0,0.907,0.093,1,0.207,1],
  ['L',0.793,1], ['C',0.907,1,1,0.907,1,0.793], ['L',1,0.207], ['C',1,0.092,0.907,0,0.793,0],
];
/** Thin bracket hooking a tile to the next icon (slide 19). */
const P_BRACKET = [
  ['M',0.998,0.822], ['L',0.918,0.655], ['L',0.913,0.655], ['L',0.913,0.667], ['L',0.987,0.819],
  ['L',0.864,0.819], ['C',0.825,0.819,0.794,0.755,0.794,0.674], ['L',0.794,0.227], ['C',0.794,0.102,0.745,0,0.685,0], ['L',0.109,0],
  ['C',0.049,0,0,0.102,0,0.227], ['L',0,0.828], ['L',0.004,0.835], ['L',0.007,0.828], ['L',0.007,0.227],
  ['C',0.007,0.11,0.053,0.016,0.109,0.016], ['L',0.685,0.016], ['C',0.741,0.016,0.786,0.11,0.786,0.227], ['L',0.786,0.674], ['C',0.786,0.763,0.821,0.835,0.864,0.835],
  ['L',0.987,0.835], ['L',0.914,0.986], ['L',0.914,0.997], ['L',0.917,1], ['L',0.919,0.997],
  ['L',0.998,0.833], ['L',0.998,0.822],
];
/** Card with a stack of chevrons cut into the left edge (slide 20). */
const P_CHEVCARD = [
  ['M',0.227,0.154], ['L',0.061,0.28], ['C',0.053,0.286,0.041,0.286,0.033,0.28], ['L',0.033,0.276], ['L',0.193,0.154],
  ['C',0.201,0.148,0.201,0.138,0.193,0.132], ['L',0.033,0.009], ['L',0.033,0.006], ['C',0.041,0,0.053,0,0.061,0.006], ['L',0.227,0.132],
  ['C',0.235,0.138,0.235,0.148,0.227,0.154], ['Z'], ['M',1,0.019], ['L',1,0.981], ['C',1,0.991,0.989,1,0.975,1],
  ['L',0.164,1], ['C',0.15,1,0.139,0.991,0.139,0.981], ['L',0.139,0.256], ['L',0.27,0.156], ['C',0.279,0.149,0.279,0.137,0.27,0.129],
  ['L',0.139,0.029], ['L',0.139,0.019], ['C',0.139,0.009,0.15,0,0.164,0], ['L',0.975,0], ['C',0.989,0,1,0.009,1,0.019],
  ['Z'], ['M',0.353,0.134], ['L',0.22,0.033], ['C',0.214,0.028,0.204,0.028,0.197,0.033], ['L',0.197,0.036],
  ['L',0.326,0.134], ['C',0.333,0.139,0.333,0.147,0.326,0.152], ['L',0.197,0.25], ['L',0.197,0.253], ['C',0.203,0.258,0.214,0.258,0.22,0.253],
  ['L',0.353,0.152], ['C',0.359,0.147,0.359,0.139,0.353,0.134], ['Z'], ['M',0.151,0.15], ['C',0.156,0.146,0.156,0.14,0.151,0.135],
  ['L',0.036,0.048], ['C',0.031,0.044,0.022,0.044,0.017,0.048], ['L',0.017,0.051], ['L',0.128,0.135], ['C',0.133,0.14,0.133,0.146,0.128,0.15],
  ['L',0.017,0.235], ['L',0.017,0.237], ['C',0.022,0.242,0.031,0.242,0.036,0.237], ['L',0.151,0.15], ['Z'],
  ['M',0.083,0.138], ['L',0.012,0.084], ['L',0,0.084], ['L',0,0.086], ['L',0.069,0.138],
  ['L',0.069,0.148], ['L',0,0.2], ['L',0,0.202], ['L',0.012,0.202], ['L',0.083,0.148],
  ['L',0.083,0.138], ['Z'],
];

// ---------------------------------------------------------------------------
// Slide 1 - zig-zag milestone timeline
// ---------------------------------------------------------------------------
function slide1(pres) {
  const s = newSlide(pres);

  const STEPS = [
    { x: 1.560, c: PINK,   up: false, n: '01', label: 'Sign Up \nForm' },
    { x: 3.817, c: LILAC,  up: true,  n: '02', label: 'First Log In' },
    { x: 6.075, c: VIOLET, up: false, n: '03', label: 'Welcome Email' },
    { x: 8.333, c: CYAN,   up: true,  n: '04', label: 'Product Walkthroughs' },
    { x: 10.590, c: PLUM,  up: false, n: '05', label: 'Customer Support' },
  ];
  // Grey elbows joining consecutive milestones (they degenerate to straight runs).
  const RUNS = [[2.158, 4.531, 4.416], [4.416, 3.313, 6.674], [6.674, 4.531, 8.932], [8.931, 3.313, 11.189]];
  RUNS.forEach(([x0, y, x1]) => line(s, x0, y, x1 - x0, 0, GREY_LN, 3));
  line(s, 11.182, 4.517, 0, 0.487, GREY_LN, 3);

  STEPS.forEach(st => {
    oval(s, st.x, 3.333, 1.183, 1.183, st.c, { line: { color: WHITE, width: 3 } });
    icon(s, st.x + 0.34, 3.62, 0.5, 0.5, IC_WHITE);
    const by = st.up ? 2.295 : 5.003;   // number badge above or below the row
    dot(s, st.x + 0.591, by + 0.28, 0.56, st.c);
    text(s, st.n, { x: st.x + 0.325, y: by + 0.111, w: 0.546, h: 0.337, fontSize: 14, color: WHITE, align: 'center' });
    text(s, st.label, { x: st.x + 0.591 - 0.971, y: st.up ? 1.471 : 5.736, w: 1.942, h: 0.64,
      fontSize: 16, color: st.c, align: 'center' });
  });
}

// ---------------------------------------------------------------------------
// Slide 2 - grid of pill-shaped tags
// ---------------------------------------------------------------------------
function slide2(pres) {
  const s = newSlide(pres);

  //          pill x,  row y, colour,  label,                  label w, icon colour, icon x/y/size
  const PILLS = [
    [2.939, 2.154, PINK,    'Welcome\nEmail',      1.406, IC_BLUE,   2.313, 2.340, 0.380],
    [6.044, 2.154, LILAC,   'Product\nTutorial',   1.406, IC_ROSE,   5.399, 2.343, 0.428],
    [9.149, 2.154, VIOLET,  'Signup\nForm',        1.406, IC_AMBER,  8.558, 2.350, 0.361],
    [1.387, 3.420, CYAN,    'Documentation',       1.644, IC_PURPLE, 0.791, 3.624, 0.344],
    [4.492, 3.420, PLUM,    'First Login',         1.406, IC_ORANGE, 3.863, 3.586, 0.421],
    [7.596, 3.420, MAGENTA, 'Data Import',         1.406, IC_GREEN,  6.975, 3.594, 0.405],
    [10.701, 3.420, PINK,   'Swag',                1.406, IC_BLUE,  10.032, 3.562, 0.469],
    [2.939, 4.686, LILAC,   'Notification',        1.406, IC_ROSE,   2.309, 4.860, 0.403],
    [6.044, 4.686, VIOLET,  'Check Up Call',       1.406, IC_AMBER,  5.436, 4.876, 0.371],
    [9.149, 4.686, CYAN,    'Educational\nEmails', 1.406, IC_PURPLE, 8.553, 4.864, 0.396],
  ];

  PILLS.forEach(([px, py, color, label, lw, ic, ix, iy, isz]) => {
    roundRect(s, px, py, 2.056, 0.752, color, 0.376);
    freeform(s, P_PILL, px - 0.810, py, 1.17, 0.752, { fill: { color: WHITE }, shadow: CARD_SHADOW });
    line(s, px + 0.360, py + 0.116, 0, 0.521, WHITE, 0.75);
    const twoLine = label.includes('\n');
    text(s, label, { x: px + 0.467, y: py + 0.116 + (twoLine ? 0 : 0.109), w: lw,
      h: twoLine ? 0.505 : 0.303, fontSize: 12, color: WHITE });
    icon(s, ix, iy, isz, isz, ic);
  });
}

// ---------------------------------------------------------------------------
// Slide 3 - five petal rosette
// ---------------------------------------------------------------------------
function slide3(pres) {
  const s = newSlide(pres);

  //        petal x, y,  rot, flipH, disc x,  disc y, colour, icon x/y/size
  const PETALS = [
    [5.374, 1.730, 300, false, 6.070, 1.683, PINK,   6.430, 2.043, 0.459],
    [6.987, 2.770,  15, false, 7.878, 3.082, LILAC,  8.171, 3.374, 0.594],
    [6.456, 4.587,  90, false, 7.050, 5.179, VIOLET, 7.317, 5.447, 0.644],
    [4.587, 4.587, 270, true,  5.112, 5.179, CYAN,   5.373, 5.441, 0.656],
    [4.047, 2.770, 345, true,  4.276, 3.082, PLUM,   4.564, 3.369, 0.604],
  ];
  PETALS.forEach(([px, py, rot, flipH, dx, dy, color, ix, iy, isz]) => {
    freeform(s, P_PETAL, px, py, 2.298, 1.697, { fill: { color }, rotate: rot, flipH });
    disc(s, dx, dy, 1.179);
    icon(s, ix, iy, isz, isz, color === PINK ? IC_BLUE : color === LILAC ? IC_ROSE
      : color === VIOLET ? IC_AMBER : color === CYAN ? IC_PURPLE : IC_ORANGE);
  });

  // Hub: grey disc split into five sectors by thin radial gaps.
  const HX = 6.6725, HY = 4.1585, HR = 0.9155;
  oval(s, HX - HR, HY - HR, 2 * HR, 2 * HR, GREY_BG);
  oval(s, HX - HR + 0.045, HY - HR + 0.045, 2 * HR - 0.09, 2 * HR - 0.09, WHITE);
  [-126, -54, 18, 90, 162].forEach(deg => {
    const r = deg * Math.PI / 180;
    line(s, HX, HY, Math.cos(r) * HR, Math.sin(r) * HR, GREY_BG, 2.5);
  });

  const HUB = [['01', 6.385, 3.496, PINK], ['02', 6.858, 3.884, LILAC], ['03', 6.649, 4.391, VIOLET],
               ['04', 6.132, 4.391, CYAN], ['05', 5.924, 3.884, PLUM]];
  HUB.forEach(([n, tx, ty, color]) =>
    text(s, n, { x: tx, y: ty, w: 0.546, h: 0.337, fontSize: 14, color, align: 'center' }));

  text(s, 'Sign \nUp Form', { x: 7.600, y: 1.953, w: 2.174, h: 0.64, fontSize: 16, color: PINK });
  text(s, 'First Login', { x: 9.411, y: 3.486, w: 2.174, h: 0.37, fontSize: 16, color: LILAC });
  text(s, 'Welcome \nEmail', { x: 8.597, y: 5.449, w: 2.174, h: 0.64, fontSize: 16, color: VIOLET });
  text(s, 'Customer Support', { x: 1.748, y: 3.352, w: 2.174, h: 0.64, fontSize: 16, color: PLUM, align: 'right' });
  text(s, 'Product Walkthrough', { x: 2.562, y: 5.449, w: 2.174, h: 0.64, fontSize: 16, color: CYAN, align: 'right' });
}

// ---------------------------------------------------------------------------
// Slide 4 - onboarding hub with satellite cards
// ---------------------------------------------------------------------------
function slide4(pres) {
  const s = newSlide(pres);
  const dash = { color: GREY_LN, width: 1.5, dashType: 'dash', endArrowType: 'triangle' };

  s.addShape('bentConnector3', { x: 4.210, y: 3.193, w: 1.294, h: 0.945, rotate: 180, flipV: true, line: dash });
  s.addShape('bentConnector3', { x: 7.830, y: 3.193, w: 1.294, h: 0.945, rotate: 180, flipH: true, flipV: true, line: dash });

  disc(s, 5.391, 1.398, 2.55);
  icon(s, 6.160, 1.678, 1.013, 1.013, IC_GREY);
  text(s, 'Customer Onboarding', { x: 5.580, y: 2.766, w: 2.174, h: 0.64, fontSize: 16, align: 'center' });

  //        card x,  card y, badge x, badge y, colour,  label,                          text y, text h
  const CARDS = [
    [1.167, 2.043, 1.777, 1.603, PINK,    'Contract',                        2.690, 0.337, 1.412, 1.610],
    [10.066, 2.043, 10.676, 1.603, MAGENTA, 'Check Up\nCall',                2.568, 0.572, 10.131, 1.969],
    [2.109, 3.918, 2.719, 3.477, LILAC,   'Client Onboarding\nQuestionnaire', 4.450, 0.572, 2.090, 2.139],
    [9.124, 3.918, 9.734, 3.477, PLUM,    'Welcome Package',                 4.450, 0.572, 9.190, 1.969],
    [4.444, 4.988, 5.054, 4.548, VIOLET,  'Launch Project',                  5.643, 0.337, 4.510, 1.969],
    [6.770, 4.988, 7.380, 4.548, CYAN,    'Client Kick-\noff Meeting',       5.530, 0.572, 6.835, 1.969],
  ];
  const order = [0, 1, 2, 3, 4, 5];
  order.forEach(i => {
    const [cx, cy, bx, by, color] = CARDS[i];
    roundRect(s, cx, cy, 2.101, 1.26, GREY_BG, 0.1227);
    oval(s, bx, by, 0.881, 0.881, color, { line: { color: WHITE, width: 3 }, shadow: CARD_SHADOW });
    icon(s, bx + 0.18, by + 0.18, 0.521, 0.521, IC_WHITE);
  });

  s.addShape('line', { x: 3.268, y: 2.673, w: 2.124, h: 0, flipH: true, flipV: true, line: dash });
  s.addShape('line', { x: 7.942, y: 2.673, w: 2.124, h: 0, flipV: true, line: dash });
  s.addShape('bentConnector3', { x: 5.781, y: 3.662, w: 0.599, h: 1.172, rotate: 90, line: dash });
  s.addShape('bentConnector3', { x: 6.944, y: 3.672, w: 0.599, h: 1.153, rotate: 270, flipH: true, line: dash });

  CARDS.forEach(([, , , , color, label, ty, th, tx, tw]) =>
    text(s, label, { x: tx, y: ty, w: tw, h: th, fontSize: 14, color, align: 'center' }));
}

// ---------------------------------------------------------------------------
// Slide 5 - horizontal journey bar with quote callouts
// ---------------------------------------------------------------------------
function slide5(pres) {
  const s = newSlide(pres);

  freeform(s, P_TAB, 2.177, 3.357, 1.973, 1.035, { fill: { color: LILAC } });
  freeform(s, P_CAP, 0.875, 2.875, 1.035, 2.0, { fill: { color: PINK }, rotate: 90 });
  [[3.935, VIOLET], [5.694, CYAN], [7.452, PLUM], [9.210, MAGENTA], [10.968, GREY_TX]]
    .forEach(([tx, color]) => freeform(s, P_TAB, tx, 3.357, 1.973, 1.035, { fill: { color } }));

  const BAR = [
    ['I need a solution\nto meet (business\nneed).', 0.549, 3.522, 1.748, 0.707],
    ['Research\nOptions', 2.765, 3.623, 1.007, 0.505],
    ['Buying Decision', 4.283, 3.724, 1.624, 0.303],
    ['Onboarding', 6.182, 3.724, 1.369, 0.303],
    ['Ongoing\nUsage', 8.047, 3.623, 1.369, 0.505],
    ['Support', 9.870, 3.724, 1.369, 0.303],
    ['Renewal', 11.540, 3.724, 1.369, 0.303],
  ];
  BAR.forEach(([t, tx, ty, tw, th]) => text(s, t, { x: tx, y: ty, w: tw, h: th, fontSize: 12, color: WHITE }));

  // Leader lines: a stub under/over the bar plus a bracket spanning two stops.
  line(s, 1.302, 5.097, 1.876, 0, GREY_LN, 3, { beginArrowType: 'oval' });
  line(s, 3.164, 4.393, 0, 0.705, GREY_LN, 3);
  [[3.164, 2.653], [6.653, 4.393]].forEach(([bx, by]) => {
    line(s, bx, by, 0, 0.705, GREY_LN, 3);
    line(s, bx + 0.005, by + 0.010, 3.511, 0, GREY_LN, 3);
    line(s, bx + 3.516, by, 0, 0.705, GREY_LN, 3);
  });
  line(s, 10.164, 2.653, 1.876, 0, GREY_LN, 3, { beginArrowType: 'oval' });
  line(s, 10.178, 2.653, 0, 0.705, GREY_LN, 3);

  const NODES = [[2.023, 4.881, PINK], [2.946, 2.453, LILAC], [4.702, 2.446, VIOLET], [6.458, 2.453, CYAN],
                 [6.458, 4.890, CYAN], [8.214, 4.882, PLUM], [9.971, 4.890, MAGENTA], [10.886, 2.453, INK]];
  NODES.forEach(([nx, ny, color]) => oval(s, nx, ny, 0.433, 0.433, color));

  caption(s, 'I need a solution to help\nme with (business need)', 1.017, 5.396, 2.436, 0.528);
  caption(s, 'It was difficult to\ncompare options.', 0.903, 2.411, 1.973, 0.528, { align: 'right' });
  caption(s, 'I feel confident with my choice after comparing to other solutions.', 3.290, 1.826, 3.256, 0.528);
  caption(s, 'How do my team\nand I use this?', 6.970, 2.411, 1.973, 0.528, { align: 'left' });
  caption(s, 'We didn\u2019t realize this\nwas a metered service.', 4.417, 4.828, 1.973, 0.528, { align: 'right' });
  caption(s, 'This is helping my team.', 6.810, 5.396, 3.256, 0.306);
  caption(s, 'I had a issue with the\nsoftware and needed support.', 10.475, 4.828, 2.567, 0.528, { align: 'left' });
  caption(s, 'Great solution, will be\nrecommending to friends!', 9.884, 1.824, 2.436, 0.528);
}

// ---------------------------------------------------------------------------
// Slide 6 - sales funnel next to onboarding stage cards
// ---------------------------------------------------------------------------
function slide6(pres) {
  const s = newSlide(pres);

  // Funnel: dark under-slices first, then the lit faces, then the pink rim.
  freeform(s, P_FUNNEL_S1, 1.251, 1.957, 3.461, 1.827, { fill: { color: '800B46' } });
  freeform(s, P_FUNNEL_S2, 1.588, 3.088, 2.764, 1.645, { fill: { color: '3D1E8E' } });
  freeform(s, P_FUNNEL_S3, 1.899, 4.089, 2.119, 1.405, { fill: { color: '351C64' } });
  freeform(s, P_FUNNEL_S4, 2.197, 4.971, 1.509, 1.592, { fill: { color: '0A6578' } });
  freeform(s, P_FUNNEL_B1, 1.249, 3.088, 3.103, 0.972, { fill: { color: LILAC } });
  freeform(s, P_FUNNEL_B2, 1.588, 4.087, 2.430, 0.830, { fill: { color: VIOLET } });
  freeform(s, P_FUNNEL_B3, 1.899, 4.971, 1.808, 0.739, { fill: { color: CYAN } });
  freeform(s, P_FUNNEL_B4, 2.197, 5.848, 1.204, 0.754, { fill: { color: PLUM } });
  freeform(s, P_FUNNEL_TOP, 0.893, 1.957, 3.818, 1.067, { fill: { color: PINK } });

  const STAGES = [['Awareness', 20, 1.742, 2.409, 2.120, 0.438], ['Discovery', 18, 1.967, 3.536, 1.671, 0.404],
                  ['Evaluation', 16, 1.967, 4.438, 1.671, 0.370], ['Purchase', 14, 1.967, 5.245, 1.671, 0.337],
                  ['Loyalty', 12, 1.967, 6.081, 1.671, 0.303]];
  STAGES.forEach(([t, sz, tx, ty, tw, th]) =>
    text(s, t, { x: tx, y: ty, w: tw, h: th, fontSize: sz, color: WHITE, align: 'center' }));

  //         card y, colour, connector x0, body text,                                        body y, body h
  const CARDS = [
    [1.957, PINK,   4.352, 'User exploring the product.', 2.232, 0.370],
    [2.888, LILAC,  4.165, 'User is achieving their goal\nusing the product.', 3.012, 0.673],
    [3.820, VIOLET, 3.863, 'A journey from being the user to\ncustomer begins at this stage.', 3.943, 0.673],
    [4.751, CYAN,   3.540, 'The user becomes the customer and\nutilize the full potential of the product.', 4.874, 0.673],
    [5.682, PLUM,   3.186, 'Stays with the product after continuous success and becomes advocate.', 5.745, 0.673],
  ];
  CARDS.forEach(([cy, color, lx, body, by, bh], i) => {
    roundRect(s, 5.555, cy, 3.818, 0.919, color, 0.0898);
    line(s, lx, cy + 0.46, 5.555 - lx, 0, color, 3);
    text(s, body, { x: i === 4 ? 5.555 : 5.760, y: by, w: i === 4 ? 3.818 : 3.408, h: bh,
      fontFace: BODY, fontSize: 12, color: WHITE, align: 'center', lineSpacingMultiple: 1.5 });
  });

  // Right-hand braces grouping the cards into two owning teams.
  const brace = { color: GREY_LN2, width: 1.5 };
  s.addShape('bentConnector3', { x: 9.373, y: 2.417, w: 0.744, h: 0.454, line: brace });
  s.addShape('bentConnector3', { x: 9.373, y: 2.871, w: 0.744, h: 0.478, flipV: true, line: brace });
  s.addShape('bentConnector3', { x: 9.373, y: 4.279, w: 0.744, h: 0.906, line: brace });
  s.addShape('bentConnector3', { x: 9.373, y: 5.185, w: 0.744, h: 0.896, flipV: true, line: brace });
  s.addShape('line', { x: 9.373, y: 5.185, w: 0.790, h: 0, flipH: true, flipV: true, line: brace });

  s.addText([{ text: 'Taken care by', options: { fontFace: BODY, breakLine: true } },
             { text: 'Product Team ', options: { fontFace: HEAD } },
             { text: 'and', options: { fontFace: BODY, breakLine: true } },
             { text: 'Marketing Team', options: { fontFace: BODY } }],
    { x: 10.163, y: 2.467, w: 2.277, h: 0.808, fontSize: 14, color: GREY_TX, align: 'center', valign: 'top' });
  s.addText([{ text: 'Taken care by', options: { fontFace: BODY, breakLine: true } },
             { text: 'Customer Success', options: { fontFace: HEAD, breakLine: true } },
             { text: 'Team ', options: { fontFace: HEAD } },
             { text: 'and', options: { fontFace: BODY, breakLine: true } },
             { text: 'Account Manager', options: { fontFace: BODY } }],
    { x: 10.163, y: 4.664, w: 2.277, h: 1.043, fontSize: 14, color: GREY_TX, align: 'center', valign: 'top' });

  text(s, 'Sales Funnel Stages', { x: 1.253, y: 1.390, w: 3.099, h: 0.37, fontSize: 16, align: 'center' });
  text(s, 'Onboarding Stage', { x: 5.915, y: 1.390, w: 3.099, h: 0.37, fontSize: 16, align: 'center' });
}

// ---------------------------------------------------------------------------
// Slide 7 - six dashed rings between two grey bands
// ---------------------------------------------------------------------------
function slide7(pres) {
  const s = newSlide(pres);

  const STEPS = [
    [1.020, PINK,    IC_WHITE, 1.550, 3.694, 0.581, 'Onboarding request', 5.499, 0.572],
    [2.950, LILAC,   IC_WHITE, 3.443, 3.657, 0.656, 'Data capture', 5.617, 0.337],
    [4.881, VIOLET,  IC_WHITE, 5.368, 3.651, 0.667, 'Duo-diligence', 5.617, 0.337],
    [6.812, CYAN,    IC_WHITE, 7.343, 3.696, 0.577, 'Legal\ndocumentation', 5.499, 0.572],
    [8.742, PLUM,    IC_WHITE, 9.245, 3.667, 0.635, 'Customer/\naccount set-up', 5.499, 0.572],
    [10.673, MAGENTA, IC_WHITE, 11.156, 3.648, 0.674, 'Communication/\ninteraction', 5.499, 0.572],
  ];
  STEPS.forEach(([cx, color]) =>
    oval(s, cx, 3.164, 1.641, 1.641, null, { fill: { color: WHITE, transparency: 100 },
      line: { color, width: 1.5, dashType: 'dash' } }));
  STEPS.forEach(([cx, color]) => oval(s, cx + 0.213, 3.377, 1.215, 1.215, color));

  // Small chevron badge tucked at the lower right of every ring.
  STEPS.forEach(([cx, color], i) => {
    const bx = 2.170 + i * 1.942;
    oval(s, bx, 4.473, 0.32, 0.32, WHITE, { rotate: 322.2, shadow: CARD_SHADOW });
    freeform(s, P_CHEVRON, bx + 0.128, 4.579, 0.064, 0.108, { fill: { color }, rotate: 322.2 });
  });
  STEPS.slice(0, 5).forEach(([cx, color]) =>
    line(s, cx + 1.641, 3.985, 0.29, 0, color, 1.5, { dashType: 'dash' }));

  roundRect(s, 1.020, 1.657, 11.294, 1.055, GREY_BG, 0.1291);
  const TOP = [['Relationship\nmanagement', 1.791, 1.898, 0.572], ['KYC / Credit', 3.722, 2.016, 0.337],
               ['Legal', 5.652, 2.016, 0.337], ['Operations', 7.583, 2.016, 0.337],
               ['Relationship\nmanagement', 9.513, 1.898, 0.572]];
  TOP.forEach(([t, tx, ty, th]) => text(s, t, { x: tx, y: ty, w: 2.029, h: th, fontSize: 14, align: 'center' }));

  roundRect(s, 0.407, 5.258, 12.520, 1.055, GREY_BG, 0.1291);
  STEPS.forEach(([, color, , , , , label, ly, lh], i) =>
    text(s, label, { x: 0.826 + i * 1.9305, y: ly, w: 2.029, h: lh, fontSize: 14, color, align: 'center' }));
  STEPS.forEach(([, , ic, ix, iy, isz]) => icon(s, ix, iy, isz, isz, ic));
}

// ---------------------------------------------------------------------------
// Slide 8 - arrow ribbon over four icon rings
// ---------------------------------------------------------------------------
function slide8(pres) {
  const s = newSlide(pres);

  const LOREM = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod';
  //        arrow x, dot x,   ring x, colour,  heading,                  head x, head w, body x, icon x/y/size
  const COLS = [
    [1.276, 2.584, 1.911, PINK,   'Onboarding Assistance', 1.618, 2.173, 1.491, 2.394, 3.738, 0.581, IC_BLUE],
    [3.918, 5.226, 4.568, LILAC,  'Proactive\nDetection',  4.259, 2.173, 4.132, 4.998, 3.685, 0.688, IC_ROSE],
    [6.559, 7.867, 7.214, VIOLET, 'Dedicated\nTouchpoint', 6.901, 2.173, 6.774, 7.640, 3.681, 0.695, IC_AMBER],
    [9.201, 10.509, 9.854, CYAN,  'Cyclic Checks',         9.857, 1.544, 9.416, 10.291, 3.691, 0.675, IC_PURPLE],
  ];
  const DOT_Y = [1.975, 1.919, 1.928, 1.958];

  COLS.forEach(([ax, dx, rx, color], i) => {
    freeform(s, P_ARROWBAND, ax, 1.702, 2.856, 0.717, { fill: { color } });
    oval(s, dx, DOT_Y[i], 0.24, 0.24, WHITE);
  });
  COLS.forEach(([, dx, rx, color], i) => {
    oval(s, rx, 3.255, 1.547, 1.547, color);
    s.addShape('line', { x: dx + 0.1, y: DOT_Y[i] + 0.24, w: 0.01, h: 3.255 - DOT_Y[i] - 0.24, flipH: true,
      line: { color: GREY_LN, width: 1.5, dashType: 'dash' } });
    disc(s, rx + 0.184, 3.439, 1.179);
  });
  COLS.forEach(([, , , color, head, hx, hw, bx, ix, iy, isz, ic]) => {
    icon(s, ix, iy, isz, isz, ic);
    text(s, head, { x: hx, y: 4.986, w: hw, h: 0.64, fontSize: 16, color, align: 'center' });
    caption(s, LOREM, bx, 5.625, 2.427, 0.75);
  });
}

// ---------------------------------------------------------------------------
// Slide 9 - ten stage S-curve ribbon
// ---------------------------------------------------------------------------
function slide9(pres) {
  const s = newSlide(pres);

  // Alternating half-loops; odd indexes sit high, even indexes low.
  for (let i = 0; i < 10; i++) {
    const low = i % 2 === 0;
    const cx = (low ? 1.452 : 2.405) + Math.floor(i / 2) * 1.916;
    const color = i < 5 ? PINK : LILAC;
    freeform(s, P_SCURVE, cx, low ? 3.601 : 2.606, 1.811, 1.803,
      { fill: { color }, rotate: low ? 135 : 315, flipH: true });
    const bx = (low ? 2.106 : 3.059) + Math.floor(i / 2) * 1.916;
    disc(s, bx, low ? 5.113 : 2.394, 0.504);
    text(s, String(i + 1), { x: bx - 0.012, y: low ? 5.194 : 2.494, w: 0.526, h: 0.303,
      fontSize: 12, color: '000000', align: 'center' });
  }

  const LOWER = [['Unaware', 1.271, PINK], ['Consideration', 3.186, PINK], ['Decision', 5.101, PINK],
                 ['Adoption', 7.016, LILAC], ['Loyalty', 8.931, LILAC]];
  LOWER.forEach(([t, tx, color]) =>
    text(s, t, { x: tx, y: 5.746, w: 2.173, h: 0.37, fontSize: 16, color, align: 'center' }));
  const UPPER = [['Aware', 2.229, 1.894, 0.37, PINK], ['Evaluation', 4.144, 1.894, 0.37, PINK],
                 ['On-board', 6.059, 1.894, 0.37, LILAC], ['Value\nRealization', 7.974, 1.625, 0.64, LILAC],
                 ['Advocacy', 9.889, 1.894, 0.37, LILAC]];
  UPPER.forEach(([t, tx, ty, th, color]) =>
    text(s, t, { x: tx, y: ty, w: 2.173, h: th, fontSize: 16, color, align: 'center' }));

  // Grey end caps.
  [[0.726, 0.919, 1.253, 'Customer Acquisition'], [10.987, 11.180, 11.514, 'Customer \nExpansion']]
    .forEach(([ox, gx, ix, label]) => {
      oval(s, ox, 3.195, 1.621, 1.621, GREY_BG);
      icon(s, ix, 3.449, 0.566, 0.566, IC_GREY);
      text(s, label, { x: gx, y: 4.029, w: 1.234, h: 0.505, fontFace: BODY, fontSize: 12, align: 'center' });
    });
}

// ---------------------------------------------------------------------------
// Slide 10 - automation flow chart
// ---------------------------------------------------------------------------
function slide10(pres) {
  const s = newSlide(pres);

  oval(s, 2.076, 4.211, 2.091, 2.091, PINK);
  s.addShape('roundRect', { x: 5.925, y: 4.515, w: 1.483, h: 1.483, rectRadius: 0.1355, rotate: 45, fill: { color: CYAN } });
  roundRect(s, 9.167, 4.211, 2.091, 2.091, MAGENTA, 0.191);
  oval(s, 9.167, 1.607, 2.091, 2.091, PLUM, { flipH: true });
  s.addShape('roundRect', { x: 5.925, y: 1.911, w: 1.483, h: 1.483, rectRadius: 0.1355, rotate: 315, flipH: true, fill: { color: VIOLET } });
  roundRect(s, 2.076, 1.607, 2.091, 2.091, LILAC, 0.191, { flipH: true });

  const arrow = { color: GREY_LN, width: 1.5, endArrowType: 'triangle' };
  s.addShape('line', { x: 3.121, y: 3.750, w: 0, h: 0.461, flipV: true, line: arrow });
  s.addShape('line', { x: 3.861, y: 3.177, w: 2.192, h: 1.341, flipV: true, line: arrow });
  s.addShape('line', { x: 4.167, y: 5.257, w: 1.451, h: 0, line: arrow });
  s.addShape('line', { x: 7.653, y: 5.257, w: 1.451, h: 0, line: arrow });
  s.addShape('line', { x: 7.653, y: 2.653, w: 1.451, h: 0, line: arrow });
  text(s, 'Yes', { x: 7.767, y: 2.691, w: 1.153, h: 0.337, fontSize: 14, align: 'center' });
  text(s, 'On Send', { x: 7.767, y: 5.302, w: 1.153, h: 0.337, fontSize: 14, align: 'center' });

  //       icon x/y/size, title,                  title x/y/w, sub,                sub x/y/w/h
  const NODES = [
    [2.864, 1.891, 0.514, 'Show Heads Up\nMessage', 2.114, 2.558, 2.014, 0.505, 'Paying Customer', 2.161, 3.023, 1.921, 0.306],
    [6.403, 2.101, 0.527, 'Check Field',            5.823, 2.760, 1.688, 0.303, 'MPR',              6.267, 3.023, 0.799, 0.306],
    [9.953, 1.920, 0.518, 'Schedule Activity',      9.205, 2.558, 2.014, 0.303, 'Current Owner \u2013\nWelcome', 9.252, 2.801, 1.921, 0.528],
    [2.849, 4.408, 0.545, 'Smart Segment Trigger',  2.114, 5.049, 2.014, 0.505, 'Paying Customer', 2.161, 5.513, 1.921, 0.306],
    [6.396, 4.613, 0.542, 'Send Email',             5.895, 5.250, 1.480, 0.303, 'Welcome!',         6.026, 5.513, 1.282, 0.306],
    [9.921, 4.475, 0.583, 'Add Delay',              9.472, 5.250, 1.480, 0.303, '2 Days',           9.602, 5.513, 1.282, 0.306],
  ];
  NODES.forEach(([ix, iy, isz, title, tx, ty, tw, th, sub, sx, sy, sw, sh]) => {
    icon(s, ix, iy, isz, isz, IC_WHITE);
    text(s, title, { x: tx, y: ty, w: tw, h: th, fontSize: 12, color: WHITE, align: 'center' });
    caption(s, sub, sx, sy, sw, sh, { color: WHITE });
  });
}

// ---------------------------------------------------------------------------
// Slide 11 - snake of labelled circles
// ---------------------------------------------------------------------------
function slide11(pres) {
  const s = newSlide(pres);

  //        circle x, y,   colour,  label,                     flipH, label y, label w, label h
  const NODES = [
    [2.218, 3.022, VIOLET,  'Product\nSetup',        false, 3.451, 1.239, 0.505],
    [4.730, 3.022, LILAC,   'Greeting\nMessage',     false, 3.451, 1.239, 0.505],
    [7.241, 3.022, LILAC,   'Mini\nCelebrations',    false, 3.451, 1.393, 0.505],
    [9.769, 3.022, CYAN,    'Routine\nCheck-Ins',    false, 3.451, 1.239, 0.505],
    [2.218, 5.034, CYAN,    'Empty\nStates',         true,  5.462, 1.239, 0.505],
    [4.730, 5.034, PLUM,    'Feature\nCallouts',     true,  5.509, 1.239, 0.505],
    [7.257, 5.034, MAGENTA, 'Interactive\nWalk-\nThrough', true, 5.361, 1.239, 0.707],
    [9.769, 5.034, VIOLET,  'Knowledge\nBase',       true,  5.462, 1.239, 0.505],
    [5.986, 1.454, PINK,    'Welcome \nEmail',       false, 1.883, 1.239, 0.505],
  ];
  NODES.forEach(([cx, cy, color, label, flipH, ly, lw, lh]) => {
    oval(s, cx, cy, 1.362, 1.362, color, flipH ? { flipH: true } : {});
    text(s, label, { x: cx + 0.681 - lw / 2, y: ly, w: lw, h: lh, fontSize: 12, color: WHITE, align: 'center' });
  });

  // Direction chevrons between the circles.
  const ARROWS = [
    [5.827, 2.711, 126.87, false, PINK], [3.961, 3.510, 0, true, LILAC], [2.705, 4.515, 270, true, VIOLET],
    [3.961, 5.521, 0, false, CYAN], [6.481, 5.521, 0, false, PLUM], [9.008, 5.521, 0, false, MAGENTA],
    [10.256, 4.515, 90, true, VIOLET], [8.992, 3.510, 0, true, CYAN],
  ];
  ARROWS.forEach(([ax, ay, rot, back, color]) => {
    oval(s, ax, ay, 0.387, 0.387, WHITE, { rotate: rot, shadow: CARD_SHADOW });
    freeform(s, P_CHEVRON, ax + 0.155, ay + 0.128, 0.078, 0.131,
      { fill: { color }, rotate: (rot + (back ? 180 : 0)) % 360 });
  });
}

// ---------------------------------------------------------------------------
// Slide 12 - customer journey matrix
// ---------------------------------------------------------------------------
function slide12(pres) {
  const s = newSlide(pres);

  const COL_X = [0.581, 2.624, 4.666, 6.709, 8.752, 10.794];
  const COL_W = 1.958;
  const HEADS = ['Stage', 'Awareness', 'Consideration', 'Decision', 'Service', 'Loyalty'];
  const HEAD_COLORS = [PINK, LILAC, VIOLET, CYAN, PLUM, MAGENTA];
  const ROW_Y = [1.821, 2.631, 3.433, 4.496, 5.298, 6.099];
  const ROW_H = [0.7, 0.7, 0.961, 0.7, 0.7, 0.7];

  COL_X.forEach((cx, i) => {
    s.addShape('round2SameRect', { x: cx, y: 1.304, w: COL_W, h: 0.424, fill: { color: HEAD_COLORS[i] } });
    text(s, HEADS[i], { x: cx + 0.140, y: 1.359, w: 1.678, h: 0.315, fontSize: 12, color: WHITE, align: 'center' });
  });
  ROW_Y.forEach((ry, r) => COL_X.forEach(cx => rect(s, cx, ry, COL_W, ROW_H[r], GREY_BG)));

  // Row 1 - customer actions
  const R1 = [['Customer Actions', 2.037, 0.268], ['View the\nadvertisement online', 1.946, 0.45],
              ['Online research\ncompetitor analysis', 1.946, 0.45], ['Made purchase', 2.037, 0.268],
              ['Receives post-safe customer service and product walkthrough', 1.855, 0.632],
              ['Make another purchase and refer product with others', 1.855, 0.632]];
  R1.forEach(([t, ty, th], i) => caption(s, t, COL_X[i] + 0.078, ty, 1.802, th, { fontSize: 9 }));

  // Row 2 - touchpoints
  caption(s, 'Touchpoints', COL_X[0] + 0.078, 2.847, 1.802, 0.268, { fontSize: 9 });
  const R2 = [['Word-of-mouth', 'Social media'], ['Website', 'Social media'], ['Website', 'Mobile app'],
              ['Phone', 'Chat', 'Email'], ['Word-of-mouth', 'Product reviews']];
  R2.forEach((items, i) => bulletList(s, items, COL_X[i + 1] + 0.078, i === 3 ? 2.666 : 2.757, 1.802, i === 3 ? 0.632 : 0.45));

  // Row 3 - experience sentiment line with emoji faces
  caption(s, 'Customer Experience', COL_X[0] + 0.078, 3.766, 1.802, 0.268, { fontSize: 9 });
  const SEG = [[2.617, 4.034, 1.965, 0.183, false, true], [4.660, 3.661, 1.965, 0.350, false, true],
               [6.709, 3.661, 1.958, 0.000, false, false], [8.752, 3.661, 1.965, 0.350, true, true],
               [10.788, 3.661, 1.965, 0.350, false, true]];
  SEG.forEach(([lx, ly, lw, lh, fh, fv]) =>
    s.addShape('line', { x: lx, y: ly, w: lw, h: lh, flipH: fh, flipV: fv, line: { color: '808080', width: 1 } }));

  const MOODS = [['Interested\nand uncertain', 2.986, 3.524, 0.404, 3.480, 4.011, LILAC, 'flat'],
                 ['Excited\nand inquisitive', 5.028, 3.941, 0.404, 5.519, 3.706, VIOLET, 'smile'],
                 ['Excited\nand happy', 7.071, 3.941, 0.404, 7.565, 3.539, CYAN, 'smile'],
                 ['Annoyed', 9.114, 4.092, 0.252, 9.608, 3.709, PLUM, 'frown'],
                 ['Satisfied\nand happy', 11.156, 3.941, 0.404, 11.647, 3.686, MAGENTA, 'smile']];
  MOODS.forEach(([label, tx, ty, th, fx, fy, color, mouth]) => {
    text(s, label, { x: tx, y: ty, w: 1.234, h: th, fontFace: BODY, fontSize: 9, align: 'center' });
    oval(s, fx, fy, 0.246, 0.246, color);
    oval(s, fx + 0.055, fy + 0.065, 0.05, 0.05, WHITE);
    oval(s, fx + 0.140, fy + 0.065, 0.05, 0.05, WHITE);
    if (mouth === 'flat') line(s, fx + 0.068, fy + 0.162, 0.109, 0, WHITE, 1);
    else s.addShape('arc', { x: fx + 0.036, y: fy + (mouth === 'smile' ? 0.011 : 0.160), w: 0.174, h: 0.174,
      rotate: mouth === 'smile' ? 135 : 225, flipV: mouth !== 'smile', line: { color: WHITE, width: 1 } });
  });

  // Row 4 - KPIs
  caption(s, 'KPIs', COL_X[0] + 0.078, 4.712, 1.802, 0.268, { fontSize: 9 });
  text(s, '-', { x: COL_X[1], y: 4.496, w: COL_W, h: 0.7, fontFace: 'Montserrat Medium', fontSize: 12, align: 'center', valign: 'middle' });
  bulletList(s, ['Total visitors', 'Average session duration'], COL_X[2] + 0.078, 4.530, 1.802, 0.632);
  bulletList(s, ['Leads generated', 'Conversion rate'], COL_X[3] + 0.078, 4.621, 1.802, 0.45);
  bulletList(s, ['Customer service\nsuccess rate', 'Net promoter score'], COL_X[4] + 0.078, 4.530, 1.880, 0.632);
  bulletList(s, ['Retention rate', 'Customer lifetime value', 'Customer satisfaction'], COL_X[5] + 0.078, 4.530, 2.259, 0.632);

  // Rows 5 and 6
  const R5 = [['Business Goals', 5.514, 0.268], ['Increase brand awareness and attract more customer ', 5.332, 0.632],
              ['Increase website visitors', 5.514, 0.268], ['Increase conversion rate\nand revenue', 5.423, 0.45],
              ['Improve customer\nsupport services', 5.423, 0.45], ['Increase retention and reduce chum rate', 5.423, 0.45]];
  R5.forEach(([t, ty, th], i) => caption(s, t, COL_X[i] + 0.078, ty, 1.802, th, { fontSize: 9 }));
  const R6 = [['Team(S) Involved', 6.315, 0.268], ['Marketing team', 6.315, 0.268],
              ['Communication and\nmarketing team', 6.225, 0.45], ['IT, sales and \nmarketing team', 6.225, 0.45],
              ['Customer service team', 6.315, 0.268], ['Marketing team', 6.315, 0.268]];
  R6.forEach(([t, ty, th], i) => caption(s, t, COL_X[i] + 0.078, ty, 1.802, th, { fontSize: 9 }));
}

// ---------------------------------------------------------------------------
// Slide 13 - dashed timeline with alternating cards
// ---------------------------------------------------------------------------
function slide13(pres) {
  const s = newSlide(pres);
  line(s, 0, 3.881, 13.333, 0, GREY_LN2, 1.5, { dashType: 'dash' });

  //       card x, above?, colour,  label,              label x, label w, label h, label y
  const CARDS = [
    [1.294, true,  PINK,    'Welcome Mail',      1.487, 1.735, 0.337, 2.287],
    [2.732, false, LILAC,   'Product Tutorial',  2.583, 2.417, 0.337, 5.138],
    [4.169, true,  VIOLET,  'Follow-Ups',        4.362, 1.735, 0.337, 2.287],
    [5.606, false, CYAN,    'Data Import',       5.799, 1.735, 0.337, 5.138],
    [7.043, true,  PLUM,    'App-Notifications', 6.851, 2.507, 0.337, 2.287],
    [8.481, false, MAGENTA, 'Knowledge\nBase',   8.674, 1.735, 0.572, 5.020],
    [9.918, true,  GREY_TX, 'Routine\nCheck-ins', 10.111, 1.735, 0.572, 2.169],
  ];
  CARDS.forEach(([cx, above, color, label, lx, lw, lh, ly]) => {
    line(s, cx + 1.061, above ? 2.881 : 3.881, 0, 1.0, GREY_LN2, 1.5, { dashType: 'dash' });
    s.addShape('roundRect', { x: cx, y: above ? 2.030 : 4.881, w: 2.121, h: 0.851, rectRadius: 0.0862,
      flipV: !above, fill: { color } });
    text(s, label, { x: lx, y: ly, w: lw, h: lh, fontSize: 14, color: WHITE, align: 'center' });
  });

  // Chevron pucks riding the dashed rail.
  [PINK, LILAC, VIOLET, CYAN, PLUM, MAGENTA].forEach((color, i) => {
    const bx = 2.880 + i * 1.4374;
    oval(s, bx, 3.687, 0.387, 0.387, WHITE, { shadow: CARD_SHADOW });
    freeform(s, P_CHEVRON, bx + 0.155, 3.815, 0.078, 0.131, { fill: { color } });
  });
}

// ---------------------------------------------------------------------------
// Slide 14 - three tall feature cards
// ---------------------------------------------------------------------------
function slide14(pres) {
  const s = newSlide(pres);
  line(s, 3.605, 3.686, 1.439, 0, GREY_LN2, 1.5, { dashType: 'dash' });
  line(s, 7.711, 3.686, 1.439, 0, GREY_LN2, 1.5, { dashType: 'dash' });

  const LOREM = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor';
  //        card x, colour,  number, label,               label y, label h, icon x/y/size
  const CARDS = [
    [1.705, PINK,   '01', 'Feature\nPromotion', 3.400, 0.572, 2.536, 2.314, 0.817],
    [5.810, LILAC,  '02', 'Customization',      3.518, 0.337, 6.607, 2.270, 0.886],
    [9.915, VIOLET, '03', 'Instructions',       3.518, 0.337, 10.765, 2.335, 0.779],
  ];
  CARDS.forEach(([cx, color, num, label, ly, lh, ix, iy, isz]) => {
    roundRect(s, cx, 2.133, 2.479, 3.642, color, 0.1315);
    icon(s, ix, iy, isz, isz, IC_WHITE);
    roundRect(s, cx - 0.766, 3.331, 2.667, 0.711, WHITE, 0.0853, { shadow: CARD_SHADOW });
    text(s, num, { x: cx - 0.716, y: 3.400, w: 0.784, h: 0.572, fontSize: 28, color, align: 'center' });
    text(s, label, { x: cx + 0.068, y: ly, w: 1.735, h: lh, fontSize: 14, color, align: 'center' });
    caption(s, LOREM, cx + 0.187, 4.216, 2.104, 0.972, { color: WHITE });
    disc(s, cx + 0.872, 5.408, 0.734);
    s.addShape('rightArrow', { x: cx + 1.082, y: 5.628, w: 0.315, h: 0.294, fill: { color } });
  });
}

// ---------------------------------------------------------------------------
// Slide 15 - staggered day cards riding a grey hill
// ---------------------------------------------------------------------------
function slide15(pres) {
  const s = newSlide(pres);
  freeform(s, P_HILL, 0, 5.212, 13.333, 2.288, { fill: { color: GREY_BG } });

  //       card x, card y, tab y,  pin y,   colour,  tab label,       tab h, tab ty
  const CARDS = [
    [5.661, 1.876, 1.517, 5.626, VIOLET, 'Day 3',        0.337, 1.708],
    [3.405, 2.179, 1.819, 5.929, LILAC,  'Day 2',        0.337, 2.010],
    [7.918, 2.179, 1.819, 5.929, CYAN,   'End of\nWeek 1', 0.572, 1.892],
    [1.148, 2.828, 2.468, 6.578, PINK,   'Day 1',        0.337, 2.659],
    [10.174, 2.828, 2.468, 6.578, PLUM,  'Week 2 &\nBeyond', 0.572, 2.542],
  ];
  CARDS.forEach(([cx, cy, ty, py, color, label, th, tty]) => {
    oval(s, cx + 1.005 - 0.146, py, 0.292, 0.292, color);
    line(s, cx + 1.005, cy + 2.625, 0, py - cy - 2.625, color, 1.5);
    roundRect(s, cx, cy, 2.010, 2.625, color, 0.1144);
    roundRect(s, cx + 0.172, ty, 1.667, 0.719, WHITE, 0.05, { shadow: CARD_SHADOW });
    text(s, label, { x: cx + 0.419, y: tty, w: 1.173, h: th, fontSize: 14, color, align: 'center' });
    oval(s, cx + 1.005 - 0.0885, py + 0.057, 0.177, 0.177, WHITE);
  });

  bulletList(s, DAY1_TASKS, 1.253, 3.372, 1.802, 1.536,
    { fontSize: 12, color: WHITE, marker: { characterCode: '2022' } });
  caption(s, 'Show them how\nto get human\nsupport', 3.509, 3.087, 1.802, 0.809, { fontSize: 12, color: WHITE });
  caption(s, 'Show them\nuseful features\n& resources', 5.766, 2.784, 1.802, 0.809, { fontSize: 12, color: WHITE });
  caption(s, 'Request\nfeedback', 8.022, 3.208, 1.802, 0.567, { fontSize: 12, color: WHITE });
  caption(s, 'Continue to deliver\nvalue relevant to the customer\u2019s goals', 10.279, 3.615, 1.802, 1.051,
    { fontSize: 12, color: WHITE });
}

// ---------------------------------------------------------------------------
// Slide 16 - teardrop stages in two rows
// ---------------------------------------------------------------------------
function slide16(pres) {
  const s = newSlide(pres);

  //        drop x, drop y, colour,  label,             label x, label w
  const DROPS = [
    [3.647, 2.525, LILAC,   'Research\nOptions',   3.797, 1.173],
    [2.506, 3.897, PINK,    'Business need',       2.656, 1.173],
    [4.789, 3.897, VIOLET,  'Buying decision',     4.939, 1.173],
    [5.931, 2.525, CYAN,    'Onboarding',          6.127, 1.079],
    [7.072, 3.897, PLUM,    'Ongoing\nUsage',      7.160, 1.296],
    [8.214, 2.525, MAGENTA, 'Support',             8.410, 1.079],
    [9.355, 3.897, INK,     'Renewal',             9.552, 1.079],
  ];
  DROPS.forEach(([dx, dy, color]) => {
    s.addShape('teardrop', { x: dx, y: dy, w: 1.472, h: 1.472, rotate: 225, flipH: true, flipV: true, fill: { color } });
    disc(s, dx + 0.156, dy + 0.156, 1.161);
  });
  DROPS.forEach(([dx, dy, color, label, lx, lw]) => {
    const oneLine = !label.includes('\n');
    text(s, label, { x: lx, y: dy + (oneLine && lw === 1.079 ? 0.568 : 0.450), w: lw,
      h: oneLine && lw === 1.079 ? 0.337 : 0.572, fontSize: 14, color, align: 'center' });
  });

  caption(s, 'I need a solution to help\nme with (business need)', 2.024, 5.550, 2.436, 0.528);
  caption(s, 'I feel confident with my choice after comparing to other solutions.', 4.509, 5.550, 2.020, 0.75);
  caption(s, 'This is helping my team.', 6.579, 5.550, 2.458, 0.306);
  caption(s, 'It was difficult to\ncompare options.', 3.373, 1.775, 2.020, 0.528);
  caption(s, 'How do my team\nand I use this?', 5.657, 1.775, 2.020, 0.528);
  caption(s, 'I had a issue with the\nsoftware and needed support.', 7.940, 1.553, 2.020, 0.75);
  caption(s, 'Great solution, will be\nrecommending to friends!', 8.849, 5.550, 2.458, 0.528);
}

// ---------------------------------------------------------------------------
// Slide 17 - concentric rings with a bulleted list
// ---------------------------------------------------------------------------
function slide17(pres) {
  const s = newSlide(pres);

  freeform(s, P_RING01, 7.665, 1.637, 4.840, 4.840, { fill: { color: PINK } });
  freeform(s, P_RING02, 8.200, 2.521, 3.771, 3.771, { fill: { color: LILAC } });
  freeform(s, P_RING03, 8.735, 3.406, 2.701, 2.698, { fill: { color: VIOLET } });
  freeform(s, P_RING04, 9.271, 4.290, 1.629, 1.629, { fill: { color: CYAN } });

  const dash = { color: GREY_LN, width: 1.5, dashType: 'dash', endArrowType: 'triangle' };
  const LOREM = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam in pretium risus. Fusce eget laoreet leo. Nulla';
  const ROWS = [[2.465, 2.201, 2.286, 0.356, IC_BLUE], [3.526, 3.264, 3.303, 0.422, IC_ROSE],
                [4.588, 4.327, 4.386, 0.426, IC_AMBER], [5.650, 5.391, 5.473, 0.414, IC_PURPLE]];
  ROWS.forEach(([ly, ty, iy, isz, ic], i) => {
    s.addShape('line', { x: 6.176, y: ly, w: 1.336, h: 0, flipH: true, line: dash });
    icon(s, i === 0 ? 0.863 : i === 1 ? 0.831 : i === 2 ? 0.828 : 0.835, iy, isz, isz, ic);
    caption(s, LOREM, 1.319, ty, 4.590, 0.528, { align: 'left' });
  });

  text(s, 'Cyclic Checks', { x: 9.467, y: 2.100, w: 1.236, h: 0.505, fontSize: 12, color: WHITE, align: 'center' });
  text(s, 'Dedicated\nTouchpoint', { x: 9.316, y: 2.936, w: 1.538, h: 0.505, fontSize: 12, color: WHITE, align: 'center' });
  text(s, 'Proactive\nDetection', { x: 9.316, y: 3.819, w: 1.538, h: 0.505, fontSize: 12, color: WHITE, align: 'center' });
  text(s, 'Onboarding Assistance', { x: 9.316, y: 4.868, w: 1.538, h: 0.505, fontSize: 12, color: WHITE, align: 'center' });
}

// ---------------------------------------------------------------------------
// Slide 18 - five open rings chained by arrows
// ---------------------------------------------------------------------------
function slide18(pres) {
  const s = newSlide(pres);

  //        ring x, colour, arrow ring?, label,               label x, label w, icon x/y/size
  const RINGS = [
    [0.650, PINK,   true,  'Sign \nUp Form',        1.088, 1.370, 1.544, 2.900, 0.459, IC_BLUE],
    [3.066, LILAC,  true,  'First Login',           3.448, 1.624, 3.963, 2.833, 0.594, IC_ROSE],
    [5.482, VIOLET, true,  'Welcome \nEmail',       5.954, 1.425, 6.344, 2.807, 0.644, IC_AMBER],
    [7.898, CYAN,   true,  'Product Walkthrough',   7.981, 2.174, 8.740, 2.802, 0.656, IC_PURPLE],
    [10.314, PLUM,  false, 'Customer Support',     10.411, 2.174, 11.196, 2.828, 0.604, IC_ORANGE],
  ];
  // Draw right-to-left so each arrow head overlaps the ring to its right.
  RINGS.slice().reverse().forEach(([rx, color, arrow]) => {
    if (arrow) freeform(s, P_ARC_ARROW, rx, 2.247, 3.431, 2.440, { fill: { color }, line: { color, width: 1 } });
    else freeform(s, P_ARC_PLAIN, rx, 2.284, 2.369, 2.368, { fill: { color }, line: { color, width: 1 } });
  });
  RINGS.forEach(([rx, color, , label, lx, lw, ix, iy, isz, ic], i) => {
    icon(s, ix, iy, isz, isz, ic);
    text(s, label, { x: lx, y: 3.495, w: lw, h: label.includes('\n') ? 0.64 : 0.37,
      fontSize: 16, color, align: 'center' });
    const dx = 1.322 + i * 2.43125;
    disc(s, dx, 4.959, 0.903);
    text(s, '0' + (i + 1), { x: dx + 0.077, y: 5.192, w: 0.749, h: 0.438, fontSize: 20, color, align: 'center' });
  });
}

// ---------------------------------------------------------------------------
// Slide 19 - numbered tiles with bracket arrows
// ---------------------------------------------------------------------------
function slide19(pres) {
  const s = newSlide(pres);

  //        tile x, bracket x, colour,  number, disc x, icon x/y/size, heading, head x/y/w/h, body x
  const COLS = [
    [1.978, 1.721, PINK,   '01', 4.231, 4.368, 3.495, 0.370, IC_BLUE,  'Feature\nPromotion', 1.799, 1.951, 1.735, 0.64, 1.369],
    [5.514, 5.258, LILAC,  '02', 7.768, 7.890, 3.480, 0.401, IC_ROSE,  'Customization', 5.027, 2.186, 2.354, 0.37, 4.906],
    [9.051, 8.795, VIOLET, '03', 11.319, 11.465, 3.504, 0.353, IC_AMBER, 'Instructions', 8.911, 2.186, 1.735, 0.37, 8.482],
  ];
  const LOREM = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Etiam in pretium risus. Fusce eget laoreet leo. Nulla';

  COLS.forEach(([tx, bx, color]) => {
    freeform(s, P_NUMBOX, tx, 2.992, 1.379, 1.379, { fill: { color } });
    freeform(s, P_BRACKET, bx, 2.733, 2.385, 1.145, { fill: { color } });
  });
  COLS.forEach(([tx, , color, num, dx, ix, iy, isz, ic, head, hx, hy, hw, hh, bodyX]) => {
    disc(s, dx, 3.358, 0.645);
    icon(s, ix, iy, isz, isz, ic);
    s.addText(num, { x: tx + 0.117, y: 3.371, w: 1.144, h: 0.638, fontFace: HEAD, fontSize: 36, bold: true,
      color: WHITE, align: 'center', valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] });
    text(s, head, { x: hx, y: hy, w: hw, h: hh, fontSize: 16, color, align: 'center' });
    caption(s, LOREM, bodyX, 4.744, 2.595, 0.972);
  });
}

// ---------------------------------------------------------------------------
// Slide 20 - chevron cards over a dashed baseline
// ---------------------------------------------------------------------------
function slide20(pres) {
  const s = newSlide(pres);

  //        card x, colour,  title,             title x, title y, title w, title h
  const CARDS = [
    [0.683, PINK,   'Day 1',            1.557, 2.206, 1.173, 0.37],
    [3.113, LILAC,  'Day 2',            4.008, 2.206, 1.173, 0.37],
    [5.544, VIOLET, 'Day 3',            6.423, 2.206, 1.173, 0.37],
    [7.974, CYAN,   'End of\nWeek 1',   8.865, 2.072, 1.173, 0.64],
    [10.404, PLUM,  'Week 2 &\nBeyond', 11.292, 2.072, 1.525, 0.64],
  ];
  CARDS.forEach(([cx, color]) =>
    freeform(s, P_CHEVCARD, cx, 1.973, 2.246, 2.951, { fill: { color }, line: { color, width: 1 } }));
  CARDS.forEach(([, , title, tx, ty, tw, th]) =>
    text(s, title, { x: tx, y: ty, w: tw, h: th, fontSize: 16, color: WHITE }));

  bulletList(s, DAY1_TASKS, 1.068, 2.906, 1.802, 1.536,
    { fontSize: 12, color: WHITE, marker: { characterCode: '2022' } });
  caption(s, 'Show them how\nto get human\nsupport', 3.499, 2.906, 1.802, 0.809, { fontSize: 12, color: WHITE, align: 'left' });
  caption(s, 'Show them\nuseful features\n& resources', 5.912, 2.906, 1.802, 0.809, { fontSize: 12, color: WHITE, align: 'left' });
  caption(s, 'Request\nfeedback', 8.342, 2.906, 1.802, 0.567, { fontSize: 12, color: WHITE, align: 'left' });
  caption(s, 'Continue to deliver\nvalue relevant to the customer\u2019s goals', 10.768, 2.906, 1.802, 1.051,
    { fontSize: 12, color: WHITE, align: 'left' });

  line(s, 0, 5.615, 13.333, 0, GREY_LN2, 1.5, { dashType: 'dash' });
  [[1.772, PINK], [4.237, LILAC], [6.701, VIOLET], [9.166, CYAN], [11.631, PLUM]].forEach(([px, color]) => {
    oval(s, px, 5.464, 0.292, 0.292, color);
    line(s, px + 0.146, 4.929, 0, 0.535, color, 1.5);
    oval(s, px + 0.058, 5.521, 0.177, 0.177, WHITE);
  });
}

// ---------------------------------------------------------------------------
// Assemble
// ---------------------------------------------------------------------------
function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'WIDE_13_33', width: 13.333, height: 7.5 });
  pres.layout = 'WIDE_13_33';
  pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };

  const BUILDERS = [slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10,
                    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20];
  BUILDERS.forEach(fn => fn(pres));

  return pres.writeFile({ fileName: path.join(__dirname, '0bca4cf4-18e3-49a2-952f-38972b23d61b_grok_final.pptx') });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
