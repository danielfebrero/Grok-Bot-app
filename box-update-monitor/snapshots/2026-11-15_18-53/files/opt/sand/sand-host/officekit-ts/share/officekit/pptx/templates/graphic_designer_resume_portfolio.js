/**
 * Standalone pptxgenjs re-creation of "033effef-5599-4a31-b0e6-e76577639573.pptx".
 * A 30-slide 10 x 5.625in resume/portfolio deck (navy + slate + mist palette).
 * Raster art from the source deck is redrawn as flat placeholder shapes.
 *
 *   node 033effef-5599-4a31-b0e6-e76577639573_grok_final.js
 */
'use strict';
const path_ = require('path');
const PptxGenJS = require('pptxgenjs');

const SLIDE_W = 10, SLIDE_H = 5.625;

// deck palette (theme accents + the greys used by the source artwork)
const C = {
  navy:   '253745',
  slate:  '4A5C6A',
  sage:   '9BABAB',
  mist:   'CCD0DF',
  lilac:  'D6D9ED',
  white:  'FFFFFF',
  black:  '000000',
  coal:   '0C0C0C',
  gray:   '7F7F7F',
  gray2:  '696969',
  gray3:  '6A6A6A',
  silver: 'BFBFBF',
  smoke:  'F2F2F2',
  photoB: 'C3C3C3',
};

const DM = 'DM Sans Medium', OS = 'Open Sans';
const MARGIN = [5.4, 5.4, 2.7, 2.7];           // lIns/rIns/bIns/tIns in points
// soft ambient glow used behind the deck's white cards
// (built fresh per call — pptxgenjs rewrites the object it is handed)
const shadow = () => ({ type: 'outer', color: C.black, opacity: 0.098, blur: 15, offset: 0.001, angle: 359 });

// text-style factory -> pptxgenjs run options
function st({ face, size, color, bold, ital, under }) {
  const o = { fontFace: face, fontSize: size, color: color };
  if (bold) o.bold = true;
  if (ital) o.italic = true;
  if (under) o.underline = { style: 'sng' };
  return o;
}

// `fill` args are either a colour string or [colour, transparency]
function F(v) {
  if (v == null) return { type: 'none' };
  return Array.isArray(v) ? { color: v[0], transparency: v[1] } : { color: v };
}

// pptxgenjs has no gradient fill, so the source deck's gradients collapse to
// the average of their stops
function mix(...hex) {
  const ch = i => Math.round(hex.reduce((a, h) => a + parseInt(h.substr(i, 2), 16), 0) / hex.length);
  return [ch(0), ch(2), ch(4)].map(v => v.toString(16).padStart(2, '0').toUpperCase()).join('');
}

// common shape options: { fill, line:{c,w,d}, rot, flipH, flipV, rad, shadow, arc, thick }
function opts(o = {}) {
  const r = { fill: F(o.fill) };
  if (o.line) r.line = { color: o.line.c, width: o.line.w == null ? 1 : o.line.w, dashType: o.line.d || 'solid' };
  if (o.rot) r.rotate = o.rot;
  if (o.flipH) r.flipH = true;
  if (o.flipV) r.flipV = true;
  if (o.rad != null) r.rectRadius = o.rad;
  if (o.shadow) r.shadow = shadow();
  if (o.arc) r.angleRange = o.arc;
  if (o.thick != null) r.arcThicknessRatio = o.thick;
  return r;
}

const box   = (s, kind, x, y, w, h, o) => s.addShape(kind, { x, y, w, h, ...opts(o) });
const rect  = (s, x, y, w, h, fill) => box(s, 'rect', x, y, w, h, { fill });
const ell   = (s, x, y, w, h, fill) => box(s, 'ellipse', x, y, w, h, { fill });
const line  = (s, x, y, w, h, o) => box(s, 'line', x, y, w, h, o);
const photo = (s, x, y, w, h, fill) => rect(s, x, y, w, h, fill);

// --- custom geometry --------------------------------------------------------
const M = (x, y) => ({ x, y, moveTo: true });                                     // move-to
const B = (x1, y1, x2, y2, x, y) => ({ x, y, curve: { type: 'cubic', x1, y1, x2, y2 } });
const Z = { close: true };
const K = 0.5523;                                                                 // circle bezier constant

function path(s, x, y, w, h, pts, o) {
  const points = pts.map(p => (Array.isArray(p) ? { x: p[0], y: p[1] } : p));
  s.addShape('custGeom', { x, y, w, h, points, ...opts(o) });
}

// rounded rectangle with independent corner radii [TL, TR, BR, BL]
function rrect(s, x, y, w, h, r, fill, o = {}) {
  const [a, b, c, d] = Array.isArray(r) ? r : [r, r, r, r];
  const p = [M(a, 0)];
  p.push([w - b, 0]); if (b) p.push(B(w - b * (1 - K), 0, w, b * (1 - K), w, b));
  p.push([w, h - c]); if (c) p.push(B(w, h - c * (1 - K), w - c * (1 - K), h, w - c, h));
  p.push([d, h]);     if (d) p.push(B(d * (1 - K), h, 0, h - d * (1 - K), 0, h - d));
  p.push([0, a]);     if (a) p.push(B(0, a * (1 - K), a * (1 - K), 0, a, 0));
  p.push(Z);
  path(s, x, y, w, h, p, { fill, ...o });
}

// annulus: outer + inner circle wound in opposite directions
function ring(s, x, y, d, fill, transparency) {
  const R = d / 2, t = d * 0.1138, r = R - t;
  const circle = (rad, cw) => {
    const q = cw ? [[0, -1], [1, 0], [0, 1], [-1, 0]] : [[0, -1], [-1, 0], [0, 1], [1, 0]];
    const out = [M(R, R - rad)];
    for (let i = 0; i < 4; i++) {
      const [ax, ay] = q[i], [bx, by] = q[(i + 1) % 4];
      out.push(B(R + rad * (ax + bx * K), R + rad * (ay + by * K),
                 R + rad * (bx + ax * K), R + rad * (by + ay * K),
                 R + rad * bx, R + rad * by));
    }
    out.push(Z);
    return out;
  };
  path(s, x, y, d, d, [...circle(r, false), ...circle(R, true)],
       { fill: transparency == null ? fill : [fill, transparency] });
}

// Every raster icon in the source deck becomes the same line-art placeholder:
// a rounded outline with a small centred dot, drawn in the icon's own colour.
function icon(s, x, y, w, h, color) {
  box(s, 'roundRect', x, y, w, h, { line: { c: color, w: 0.75 }, rad: w * 0.25 });
  ell(s, x + w * 0.375, y + h * 0.375, w * 0.25, h * 0.25, color);
}

// outlined circle with a letterform inside (slide 2's social links)
function badge(s, x, y, d, ch) {
  box(s, 'ellipse', x, y, d, d, { line: { c: C.slate } });
  s.addText(ch, { x, y, w: d, h: d, align: 'center', valign: 'middle', margin: 0,
                  fontFace: DM, fontSize: 8, bold: true, color: C.navy });
}

// the recurring "swoosh" corner motif (a crescent 1.929 x 0.558in)
function swoosh(s, x, y, o) {
  path(s, x, y, 1.929, 0.558,
       [M(0, 0), [0.005, 0], [0.078, 0.035],
        B(0.245, 0.106, 0.429, 0.145, 0.623, 0.145),
        B(0.816, 0.145, 1.000, 0.106, 1.168, 0.035),
        [1.241, 0], [1.929, 0], [1.905, 0.027],
        B(1.577, 0.355, 1.123, 0.558, 0.623, 0.558),
        B(0.435, 0.558, 0.254, 0.529, 0.083, 0.476),
        [0, 0.446], Z],
       { fill: [C.mist, 75.295], ...o });
}

// grid of dots; mask 0 = full block, 1 = lower-left triangle, 2 = upper-left triangle
function dots(s, x, y, cols, rows, gap, sz, color, mask = 0, transparency) {
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      if (mask === 1 && c < cols - 1 - r) continue;
      if (mask === 2 && c < r) continue;
      ell(s, x + c * gap, y + r * gap, sz, sz, transparency == null ? color : [color, transparency]);
    }
  }
}

// --- text -------------------------------------------------------------------
// `body` is a string, or an array of strings (one per paragraph)
function txt(s, body, x, y, w, h, style, o = {}) {
  const runs = (Array.isArray(body) ? body : [body]).map((t, i, all) => ({
    text: t, options: { ...style, breakLine: i < all.length - 1 },
  }));
  s.addText(runs, textOpts(x, y, w, h, o));
}

// mixed run styles inside one paragraph: [[text, style], ...]
function rich(s, segs, x, y, w, h, o = {}) {
  s.addText(segs.map(([t, style]) => ({ text: t, options: style })), textOpts(x, y, w, h, o));
}

function textOpts(x, y, w, h, o) {
  const r = { x, y, w, h, margin: MARGIN, valign: o.va || 'top', fit: 'resize', wrap: true,
              align: o.al || 'left', fill: F(o.fill), isTextBox: true };
  if (o.ls) r.lineSpacingMultiple = o.ls;
  if (o.line) r.line = { color: o.line.c, width: o.line.w == null ? 1 : o.line.w };
  if (o.shape) r.shape = o.shape;
  if (o.rot) r.rotate = o.rot;
  if (o.rad != null) r.rectRadius = o.rad;
  return r;
}

// --- slide 29 map -----------------------------------------------------------
function usaMap(s) {
  const paint = (polys, fill) => polys.forEach(flat => {
    const pts = [];
    for (let i = 0; i < flat.length; i += 2) pts.push([flat[i], flat[i + 1]]);
    const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
    const x = Math.min(...xs), y = Math.min(...ys);
    path(s, x, y, Math.max(...xs) - x, Math.max(...ys) - y,
         [M(pts[0][0] - x, pts[0][1] - y), ...pts.slice(1).map(p => [p[0] - x, p[1] - y]), Z],
         { fill, line: { c: C.white, w: 0.75 } });
  });
  paint(USA_MAP.main, C.lilac);
  paint(USA_MAP.wa, C.navy);
  paint(USA_MAP.fl, C.sage);
}

// run styles used throughout the deck: <family><size><b/i/u><colour>
const T = {
  o8Gray: st({ face:OS, size:8, color:C.gray }),
  o8Gray2: st({ face:OS, size:8, color:C.gray2 }),
  d11bSlate: st({ face:DM, size:11, color:C.slate, bold:1 }),
  d21bBlack: st({ face:DM, size:21, color:C.black, bold:1 }),
  o9bSlate: st({ face:OS, size:9, color:C.slate, bold:1 }),
  o9Gray3: st({ face:OS, size:9, color:C.gray3 }),
  o11bSlate: st({ face:OS, size:11, color:C.slate, bold:1 }),
  d24bBlack: st({ face:DM, size:24, color:C.black, bold:1 }),
  o11Black: st({ face:OS, size:11, color:C.black }),
  o11bBlack: st({ face:OS, size:11, color:C.black, bold:1 }),
  o8White: st({ face:OS, size:8, color:C.white }),
  d9Black: st({ face:DM, size:9, color:C.black }),
  d9Navy: st({ face:DM, size:9, color:C.navy }),
  o9Gray: st({ face:OS, size:9, color:C.gray }),
  d14bWhite: st({ face:DM, size:14, color:C.white, bold:1 }),
  d11Slate: st({ face:DM, size:11, color:C.slate }),
  d11bWhite: st({ face:DM, size:11, color:C.white, bold:1 }),
  o11bWhite: st({ face:OS, size:11, color:C.white, bold:1 }),
  o9bNavy: st({ face:OS, size:9, color:C.navy, bold:1 }),
  d12Coal: st({ face:DM, size:12, color:C.coal }),
  d12bNavy: st({ face:DM, size:12, color:C.navy, bold:1 }),
  d36bBlack: st({ face:DM, size:36, color:C.black, bold:1 }),
  o11bNavy: st({ face:OS, size:11, color:C.navy, bold:1 }),
  o14White: st({ face:OS, size:14, color:C.white }),
  o8iGray: st({ face:OS, size:8, color:C.gray, ital:1 }),
  o9Slate: st({ face:OS, size:9, color:C.slate }),
  d11White: st({ face:DM, size:11, color:C.white }),
  d11bNavy: st({ face:DM, size:11, color:C.navy, bold:1 }),
  d21bCoal: st({ face:DM, size:21, color:C.coal, bold:1 }),
  d24bWhite: st({ face:DM, size:24, color:C.white, bold:1 }),
  d54bBlack: st({ face:DM, size:54, color:C.black, bold:1 }),
  o8bNavy: st({ face:OS, size:8, color:C.navy, bold:1 }),
};

// body copy that appears on more than one slide
const L = [
  'Lorem ipsum dolor sit amet, consecte',
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore.',
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod.',
  'Lorem ipsum dolor sit amet',
  'Lorem ipsum dolor sit amet, consecte adipisc',
  'Lorem ipsum dolor sit amet, consectetu adipiscing elit, sed does eiusmod.',
  'Lorem ipsum dolor sit amet, consectetur adi',
  'Lorem ipsum dolor sit amet, consectetur adipi',
  'Lorem ipsum dolor sit amet, consectetur adipiscing e, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim ve. ',
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim ve. ',
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam.',
  'Lorem ipsum dolor sit amet, consectetur adipiscing.',
  'Bachelor’s Degree  	19/08/2024',
  'Impact : Lorem ipsum dolor sit amet, consectetur adipiscing elit. ',
  'Lorem ipsum dolor sit amet, consectetur adi elit, sed do eiusmod tempor incididunt. ',
  'Lorem ipsum dolor sit amet, consectetur adipie, sed do eiusmod tempor incididunt ut labore et dolore magna. ',
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna. ',
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore.',
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt. ',
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit.',
  'Your Certificate Organizations Names',
  'www.julianmaxwell.com',
];

// simplified outline of the slide-29 map, as flat [x0,y0, x1,y1, ...] rings
const USA_MAP = {
  main: [
    [9.053,1.255,9.101,1.27,9.138,1.4,9.226,1.472,9.021,1.64,9.003,1.797,9.051,1.842,9.078,1.835,9.076,1.805,9.093,1.855,8.933,1.922,8.933,1.945,8.963,1.942,8.816,2.032,8.823,2.117,8.783,2.205,8.733,2.187,8.778,2.26,8.736,2.41,8.733,2.332,8.678,2.3,8.698,2.412,8.738,2.437,8.813,2.592,8.788,2.607,8.811,2.59,8.788,2.542,8.733,2.602,8.731,2.627,8.761,2.635,8.741,2.68,8.688,2.682,8.638,2.78,8.561,2.812,8.536,2.89,8.411,3.035,8.396,3.155,8.351,3.155,8.346,3.202,8.093,3.202,8.081,3.177,7.856,3.202,7.858,3.277,7.696,3.295,7.718,3.315,7.683,3.375,7.726,3.372,7.743,3.405,7.711,3.435,7.683,3.397,7.571,3.435,7.528,3.387,7.346,3.372,7.261,3.395,7.021,3.582,6.988,3.652,7.013,3.767,6.993,3.797,6.831,3.742,6.791,3.62,6.618,3.392,6.506,3.392,6.488,3.445,6.438,3.465,6.336,3.395,6.308,3.297,6.181,3.17,6.023,3.157,6.018,3.205,5.778,3.185,5.451,3.022,5.451,3,5.253,2.99,5.236,2.905,5.163,2.85,5.156,2.815,4.986,2.727,4.996,2.67,4.921,2.542,4.936,2.475,4.893,2.432,4.896,2.367,4.861,2.345,4.863,2.295,4.826,2.24,4.836,2.152,4.806,2.072,4.853,1.982,4.846,1.832,4.901,1.737,4.963,1.507,5.038,1.532,5.043,1.582,5.073,1.597,5.476,1.625,5.516,1.332,6.366,1.42,7.093,1.42,7.113,1.382,7.143,1.445,7.258,1.447,7.366,1.495,7.411,1.475,7.508,1.487,7.426,1.53,7.341,1.632,7.418,1.595,7.451,1.597,7.441,1.63,7.466,1.632,7.643,1.525,7.608,1.58,7.666,1.587,7.688,1.617,7.851,1.57,7.863,1.597,7.926,1.585,7.946,1.632,7.983,1.632,7.923,1.64,7.921,1.67,8.013,1.722,8.018,1.795,7.981,1.862,8.006,1.87,8.066,1.82,8.111,1.915,8.053,2.055,8.136,2.077,8.363,1.922,8.383,1.895,8.363,1.845,8.516,1.815,8.566,1.777,8.548,1.72,8.631,1.612,8.886,1.547,8.948,1.455,8.988,1.262,9.021,1.282,9.053,1.255],
    [5.491,3.29,5.573,3.335,5.758,3.355,5.851,3.807,5.898,3.8,5.948,3.85,5.998,3.807,6.133,3.93,6.186,3.937,6.196,4.005,6.133,4.03,5.986,3.892,5.886,3.84,5.803,3.832,5.768,3.855,5.726,3.83,5.593,3.902,5.573,3.842,5.528,3.895,5.551,3.92,5.526,3.957,5.363,4.08,5.373,4.117,5.323,4.082,5.286,4.127,5.288,4.1,5.213,4.125,5.231,4.09,5.348,4.05,5.433,3.977,5.438,3.925,5.321,3.917,5.331,3.852,5.278,3.85,5.253,3.812,5.251,3.732,5.293,3.682,5.371,3.677,5.381,3.655,5.368,3.612,5.341,3.63,5.266,3.602,5.241,3.545,5.321,3.512,5.371,3.542,5.293,3.417,5.491,3.29],
    [6.896,4.217,6.976,4.25,7.018,4.3,6.918,4.37,6.873,4.275,6.896,4.217],
    [5.568,3.93,5.591,3.947,5.576,3.99,5.511,4.042,5.513,3.985,5.568,3.93],
    [6.803,4.127,6.883,4.16,6.803,4.19,6.828,4.175,6.803,4.127],
  ],
  wa: [
    [5.108,1.257,5.516,1.332,5.478,1.572,5.483,1.625,5.336,1.602,5.176,1.61,5.141,1.592,5.066,1.595,5.043,1.582,5.038,1.532,5.001,1.51,4.958,1.502,4.968,1.427,4.966,1.355,4.951,1.325,4.958,1.285,5.011,1.322,5.086,1.34,5.098,1.315,5.071,1.3,5.068,1.28,5.103,1.277,5.108,1.257],
  ],
  fl: [
    [8.348,3.152,8.398,3.157,8.438,3.267,8.518,3.37,8.518,3.41,8.593,3.532,8.601,3.687,8.558,3.75,8.546,3.717,8.518,3.717,8.496,3.665,8.456,3.655,8.431,3.6,8.401,3.602,8.388,3.567,8.316,3.467,8.326,3.39,8.308,3.342,8.281,3.34,8.211,3.267,8.158,3.257,8.151,3.28,8.136,3.28,8.118,3.305,8.088,3.317,8.061,3.312,8.053,3.285,7.996,3.257,7.938,3.255,7.863,3.275,7.871,3.235,7.848,3.217,7.851,3.205,8.078,3.175,8.093,3.202,8.313,3.185,8.346,3.197,8.348,3.152],
  ],
};

// ---------------------------------------------------------------- slide 1 — Hello I’am Julian Maxwell
function slide01(s) {
  ell(s,5.825,-0.899,1.99,1.99,[C.sage,80]);
  ring(s,4.855,3.555,3.606,C.mist,75.295);
  ring(s,-1.427,-2.469,3.957,C.mist,75.295);
  box(s,'round2SameRect',5.801,1.426,5.625,2.773,{fill:C.navy,rot:270,rad:0.523});
  box(s,'roundRect',5.463,0.918,3.533,3.533,{fill:C.white,rot:45,shadow:1,rad:0.589});
  box(s,'roundRect',0.712,3.373,3.546,0.536,{fill:C.slate,rad:0.154});
  txt(s,['Hello I’am','Julian Maxwell'],0.712,1.866,3.957,1.288,T.d36bBlack);
  txt(s,'- Graphic Designer',0.864,3.49,3.204,0.303,T.o14White);
  dots(s,0.223,0.173,4,4,0.158,0.06,C.navy);
  dots(s,6.035,4.653,4,4,0.158,0.06,C.navy);
  dots(s,8.492,4.964,3,1,0.414,0.157,C.white);
  dots(s,8.492,0.446,3,1,0.414,0.157,C.white);
  ell(s,-0.669,4.964,1.99,1.99,[C.sage,80]);
  path(s,5.1,0.555,4.259,4.259,[M(2.13,0),B(2.272,0,2.414,0.054,2.523,0.163),[4.096,1.736],B(4.313,1.953,4.313,2.306,4.096,2.523),[2.523,4.096],B(2.306,4.313,1.953,4.313,1.736,4.096),[0.163,2.523],B(-0.054,2.306,-0.054,1.953,0.163,1.736),[1.736,0.163],B(1.845,0.054,1.987,0,2.13,0),Z],{fill:C.gray});
}

// ---------------------------------------------------------------- slide 2 — Personal Information
function slide02(s) {
  box(s,'rect',0,0,3.711,5.625,{fill:C.white,shadow:1});
  box(s,'roundRect',0.516,0.766,4.484,4.094,{fill:C.navy,rad:0.366});
  txt(s,'Personal Information',6.06,0.881,2.865,0.883,T.d24bBlack);
  txt(s,'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt.',6.06,1.857,2.881,0.47,T.o8Gray,{ls:1.5});
  badge(s,6.143,4.514,0.28,'f');
  badge(s,6.556,4.514,0.28,'in');
  badge(s,6.969,4.52,0.28,'X');
  txt(s,'200 Arcadway Av Nevv Canberra WA 5024 West',7.216,2.938,1.446,0.47,T.o8Gray,{ls:1.5});
  txt(s,'Address',6.094,2.95,1.066,0.252,T.o11Black);
  txt(s,L[21],7.224,3.913,1.583,0.262,T.o8Gray,{ls:1.5});
  txt(s,'Email',6.094,3.936,1.066,0.252,T.o11Black);
  txt(s,'(+62) 8123 4567 790',7.231,3.494,1.339,0.262,T.o8Gray,{ls:1.5});
  txt(s,'Phone',6.094,3.503,1.066,0.252,T.o11Black);
  txt(s,':',7.076,2.95,0.194,0.252,T.o11bBlack,{al:'center'});
  txt(s,':',7.076,3.501,0.194,0.252,T.o11bBlack,{al:'center'});
  txt(s,':',7.078,3.932,0.194,0.252,T.o11bBlack,{al:'center'});
  dots(s,9.192,0.236,4,4,0.158,0.06,C.navy);
  ring(s,-1.709,-1.427,3.957,C.mist,75.295);
  txt(s,'19/09/1999',7.231,2.526,1.339,0.262,T.o8Gray,{ls:1.5});
  txt(s,'Born',6.094,2.535,1.066,0.252,T.o11Black);
  txt(s,':',7.076,2.533,0.194,0.252,T.o11bBlack,{al:'center'});
  rrect(s,0.644,0.883,4.228,3.859,0.345,C.gray);
}

// ---------------------------------------------------------------- slide 3 — Professional Summary
function slide03(s) {
  ell(s,8.655,-0.938,1.99,1.99,[C.sage,80]);
  box(s,'round2SameRect',5.932,0.724,3.987,4.148,{fill:C.white,rot:270,shadow:1,rad:0.417});
  dots(s,8.515,5.144,3,1,0.414,0.157,C.navy);
  ell(s,-0.777,4.658,1.99,1.99,[C.sage,80]);
  dots(s,0.5,0.324,3,1,0.414,0.157,C.navy);
  txt(s,'Professional Summary',1.151,0.674,3.278,0.883,T.d24bBlack);
  txt(s,'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et. ',1.151,1.695,3.489,0.47,T.o8Gray,{ls:1.5});
  txt(s,'Career Highlights :',1.151,2.352,3.278,0.252,T.d11bNavy,{al:'justify'});
  icon(s,1.279,2.856,0.3,0.3,C.slate);
  txt(s,'Your Company',1.742,2.977,1.559,0.262,T.o8Gray,{ls:1.5});
  txt(s,'Your Career One',1.742,2.802,1.559,0.227,T.o9bSlate);
  txt(s,L[4],2.05,3.309,2.56,0.215,T.o8Gray);
  txt(s,L[4],2.05,3.564,2.56,0.215,T.o8Gray);
  ell(s,1.894,3.377,0.079,0.079,mix(C.navy,C.slate));
  ell(s,1.894,3.632,0.079,0.079,mix(C.navy,C.slate));
  icon(s,1.279,4.028,0.3,0.3,C.slate);
  txt(s,'Your Company',1.742,4.15,1.559,0.262,T.o8Gray,{ls:1.5});
  txt(s,'Your Career Two',1.742,3.975,1.559,0.227,T.o9bSlate);
  txt(s,L[4],2.05,4.482,2.56,0.215,T.o8Gray);
  txt(s,L[4],2.05,4.736,2.56,0.215,T.o8Gray);
  ell(s,1.894,4.55,0.079,0.079,mix(C.navy,C.slate));
  ell(s,1.894,4.804,0.079,0.079,mix(C.navy,C.slate));
  rrect(s,5.984,0.93,4.016,3.741,[0.391,0,0,0.391],C.gray);
}

// ---------------------------------------------------------------- slide 4 — Career Objective
function slide04(s) {
  ring(s,6.374,2.962,3.626,C.mist,75.295);
  ring(s,0,2.962,3.626,C.mist,75.295);
  box(s,'round2SameRect',0.859,3.93,8.281,1.695,{fill:C.navy,rad:0.533});
  txt(s,'“My career goal is to become a senior graphic designer, leading innovative projects that leverage cutting-edge technology to solve complex problems and enhance user experiences. I aim to continuously develop my technical skills while also mentoring junior developers to foster a collaborative and growth-oriented environment”',1.797,4.223,6.406,1.108,T.d11White,{al:'center',ls:1.5});
  box(s,'roundRect',3.883,1.303,2.234,2.234,{fill:C.white,shadow:1,rad:0.372});
  txt(s,'Career Objective',3.117,0.431,3.766,0.48,T.d24bBlack,{al:'center'});
  dots(s,0.275,0.236,4,4,0.157,0.06,C.navy);
  dots(s,9.192,0.236,4,4,0.158,0.06,C.navy);
  rrect(s,4.003,1.423,1.994,1.994,0.332,C.gray);
}

// ---------------------------------------------------------------- slide 5 — Skills Overview
function slide05(s) {
  txt(s,'Leadership',5.119,2.409,1.104,0.252,T.o11bSlate);
  rect(s,6.805,2.485,2.199,0.102,C.silver);
  rect(s,6.353,2.485,2.361,0.102,C.slate);
  txt(s,'Team Work',5.119,2.925,1.104,0.252,T.o11bSlate);
  rect(s,6.805,3,2.199,0.102,C.silver);
  rect(s,6.353,3,1.804,0.102,C.slate);
  txt(s,'Skills Overview',5.115,0.902,3.113,0.429,T.d21bBlack);
  txt(s,L[1],5.119,1.488,3.116,0.47,T.o8Gray,{ls:1.5});
  txt(s,'Photography',5.119,3.44,1.104,0.252,T.o11bSlate);
  rect(s,6.805,3.515,2.199,0.102,C.silver);
  rect(s,6.353,3.515,2.361,0.102,C.slate);
  txt(s,'Typography',5.119,3.955,1.104,0.252,T.o11bSlate);
  rect(s,6.805,4.031,2.199,0.102,C.silver);
  rect(s,6.353,4.031,1.804,0.102,C.slate);
  txt(s,'Editing',5.119,4.471,1.104,0.252,T.o11bSlate);
  rect(s,6.805,4.546,2.199,0.102,C.silver);
  rect(s,6.353,4.546,1.284,0.102,C.slate);
  rect(s,0,0,4.12,5.625,C.navy);
  dots(s,4.283,0.115,4,4,0.158,0.06,C.navy);
  ell(s,8.77,-0.769,1.99,1.99,[C.sage,80]);
  dots(s,6.567,5.285,3,1,0.414,0.157,C.navy);
  rrect(s,0,0.258,3.871,5.367,[0,0.645,0,0],C.gray);
}

// ---------------------------------------------------------------- slide 6 — Work Experience Overview
function slide06(s) {
  ring(s,6.774,-0.062,5.75,C.mist,75.295);
  line(s,5.447,0,0,5.625,{line:{c:C.navy,w:1.5,d:'lgDash'}});
  ell(s,5.378,2.744,0.138,0.138,mix(C.navy,C.slate));
  ell(s,5.378,4.311,0.138,0.138,mix(C.navy,C.slate));
  ell(s,5.378,1.176,0.138,0.138,mix(C.navy,C.slate));
  line(s,5.516,1.245,1.008,0,{line:{c:C.navy,w:1.5,d:'lgDash'}});
  line(s,5.516,2.812,1.008,0,{line:{c:C.navy,w:1.5,d:'lgDash'}});
  line(s,5.516,4.38,1.008,0,{line:{c:C.navy,w:1.5,d:'lgDash'}});
  box(s,'roundRect',7.109,-0.318,1.25,3.125,{fill:mix(C.navy,C.slate),rot:90,rad:0.208});
  ell(s,6.369,0.807,0.875,0.875,C.white);
  icon(s,6.574,1.013,0.465,0.465,C.navy);
  txt(s,L[11],7.497,1.119,1.667,0.47,T.o8White,{ls:1.5});
  txt(s,'Timeline One',7.497,0.901,1.323,0.252,T.o11bWhite);
  box(s,'roundRect',7.109,1.25,1.25,3.125,{fill:mix(C.navy,C.slate),rot:90,rad:0.208});
  ell(s,6.369,2.375,0.875,0.875,C.white);
  icon(s,6.574,2.58,0.465,0.465,C.navy);
  txt(s,L[11],7.497,2.687,1.667,0.47,T.o8White,{ls:1.5});
  txt(s,'Timeline Two',7.497,2.468,1.323,0.252,T.o11bWhite);
  box(s,'roundRect',7.109,2.817,1.25,3.125,{fill:mix(C.navy,C.slate),rot:90,rad:0.208});
  ell(s,6.369,3.942,0.875,0.875,C.white);
  icon(s,6.574,4.147,0.465,0.465,C.navy);
  txt(s,L[11],7.497,4.254,1.667,0.47,T.o8White,{ls:1.5});
  txt(s,'Timeline Three',7.497,4.036,1.323,0.252,T.o11bWhite);
  txt(s,'Work Experience Overview',0.823,0.858,3.278,0.883,T.d24bBlack);
  txt(s,'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore. ',0.823,1.879,3.489,0.47,T.o8Gray,{ls:1.5});
  dots(s,0.1,5.138,3,3,0.158,0.06,C.navy);
  dots(s,0.1,0.099,3,3,0.158,0.06,C.navy);
  rrect(s,0.703,2.715,3.85,2.052,0.342,C.gray);
}

// ---------------------------------------------------------------- slide 7 — Work Experience 1
function slide07(s) {
  rrect(s,0.813,0.802,4.187,4.823,[0.364,0.364,0,0],C.gray);
  box(s,'round2SameRect',5.495,2.102,3.679,3.523,{fill:C.white,shadow:1,rad:0.389});
  ell(s,0.295,0.303,1.99,1.99,[C.sage,80]);
  txt(s,'Work Experience 1',5.495,0.778,3.278,0.48,T.d24bBlack);
  txt(s,L[16],5.495,1.402,3.489,0.47,T.o8Gray,{ls:1.5});
  dots(s,9.275,4.888,4,4,0.158,0.06,C.navy);
  ring(s,8.207,-2.318,3.626,C.mist,75.295);
  txt(s,'Your Company',6.379,2.926,1.559,0.262,T.o8Gray,{ls:1.5});
  txt(s,'2022 – 2024',6.382,2.522,1.395,0.252,T.d11bSlate);
  txt(s,'Your Experience One',6.379,2.751,1.559,0.227,T.o9bSlate);
  txt(s,L[0],6.686,3.258,2.188,0.215,T.o8Gray);
  txt(s,L[0],6.686,3.512,2.188,0.215,T.o8Gray);
  ell(s,6.531,3.326,0.079,0.079,mix(C.navy,C.slate));
  ell(s,6.531,3.58,0.079,0.079,mix(C.navy,C.slate));
  txt(s,'Your Company',6.382,4.454,1.559,0.262,T.o8Gray,{ls:1.5});
  txt(s,'2018 – 2020',6.382,4.05,1.395,0.252,T.d11bSlate);
  txt(s,'Your Experience Two',6.382,4.279,1.559,0.227,T.o9bSlate);
  txt(s,L[0],6.689,4.786,2.188,0.215,T.o8Gray);
  txt(s,L[0],6.689,5.04,2.188,0.215,T.o8Gray);
  ell(s,6.534,4.854,0.079,0.079,mix(C.navy,C.slate));
  ell(s,6.534,5.108,0.079,0.079,mix(C.navy,C.slate));
  icon(s,5.878,2.6,0.3,0.3,C.slate);
  icon(s,5.878,4.128,0.3,0.3,C.slate);
}

// ---------------------------------------------------------------- slide 8 — Work Experience 2
function slide08(s) {
  ring(s,6.242,0.263,5.099,C.mist,75.295);
  box(s,'roundRect',5.883,0.73,3.597,4.165,{fill:C.navy,rad:0.373});
  dots(s,0.119,0.115,4,2,0.158,0.06,C.navy);
  dots(s,0.119,5.292,4,2,0.158,0.06,C.navy);
  txt(s,'Work Experience 2',1.053,0.991,3.278,0.48,T.d24bBlack);
  txt(s,'Your Company',1.708,2.268,1.559,0.262,T.o8Gray,{ls:1.5});
  txt(s,'2022 – 2024',1.708,1.844,1.395,0.252,T.d11bSlate);
  txt(s,'Your Experience One',1.708,2.093,1.559,0.227,T.o9bSlate);
  txt(s,L[7],2.012,2.6,2.619,0.215,T.o8Gray);
  txt(s,L[7],2.012,2.855,2.619,0.215,T.o8Gray);
  ell(s,1.856,2.668,0.079,0.079,mix(C.navy,C.slate));
  ell(s,1.856,2.923,0.079,0.079,mix(C.navy,C.slate));
  txt(s,'Your Company',1.708,3.833,1.559,0.262,T.o8Gray,{ls:1.5});
  txt(s,'2018 – 2020',1.708,3.409,1.395,0.252,T.d11bSlate);
  txt(s,'Your Experience Two',1.708,3.658,1.559,0.227,T.o9bSlate);
  txt(s,L[7],2.015,4.165,2.616,0.215,T.o8Gray);
  txt(s,L[7],2.015,4.42,2.616,0.215,T.o8Gray);
  ell(s,1.859,4.233,0.079,0.079,mix(C.navy,C.slate));
  ell(s,1.859,4.488,0.079,0.079,mix(C.navy,C.slate));
  icon(s,1.21,1.932,0.3,0.3,C.slate);
  icon(s,1.21,3.497,0.3,0.3,C.slate);
  rrect(s,5.758,0.73,3.597,4.165,0.373,C.gray);
}

// ---------------------------------------------------------------- slide 9 — Work Experience 3
function slide09(s) {
  path(s,0,1.921,2.436,3.626,[M(0.623,0),B(1.624,0,2.436,0.812,2.436,1.813),B(2.436,2.814,1.624,3.626,0.623,3.626),B(0.435,3.626,0.254,3.598,0.083,3.545),[0,3.514],[0,3.066],[0.077,3.103],B(0.245,3.174,0.429,3.213,0.623,3.213),B(1.396,3.213,2.023,2.586,2.023,1.813),B(2.023,1.04,1.396,0.413,0.623,0.413),B(0.429,0.413,0.245,0.452,0.077,0.523),[0,0.56],[0,0.112],[0.083,0.082],B(0.254,0.029,0.435,0,0.623,0),Z],{fill:[C.mist,75.295]});
  box(s,'round2SameRect',-0.357,1.004,4.331,3.617,{fill:C.navy,rot:90,rad:0.432});
  dots(s,9.337,4.978,4,4,0.158,0.06,C.navy,1);
  dots(s,9.337,0.115,4,4,0.158,0.06,C.navy,2);
  swoosh(s,0,0,{});
  txt(s,'Work Experience 3',5.443,0.718,3.278,0.48,T.d24bBlack);
  txt(s,L[16],5.443,1.364,3.489,0.47,T.o8Gray,{ls:1.5});
  txt(s,'Your Company',6.029,2.502,1.559,0.262,T.o8Gray,{ls:1.5});
  txt(s,'2022 – 2024',6.029,2.086,1.395,0.252,T.d11bSlate);
  txt(s,'Your Experience One',6.029,2.328,1.559,0.227,T.o9bSlate);
  txt(s,L[6],6.333,2.835,2.517,0.215,T.o8Gray);
  txt(s,L[6],6.333,3.089,2.517,0.215,T.o8Gray);
  ell(s,6.177,2.903,0.079,0.079,mix(C.navy,C.slate));
  ell(s,6.177,3.157,0.079,0.079,mix(C.navy,C.slate));
  txt(s,'Your Company',6.029,4.069,1.559,0.262,T.o8Gray,{ls:1.5});
  txt(s,'2018 – 2020',6.029,3.652,1.395,0.252,T.d11bSlate);
  txt(s,'Your Experience Two',6.029,3.894,1.559,0.227,T.o9bSlate);
  txt(s,L[6],6.336,4.401,2.517,0.215,T.o8Gray);
  txt(s,L[6],6.336,4.655,2.517,0.215,T.o8Gray);
  ell(s,6.18,4.469,0.079,0.079,mix(C.navy,C.slate));
  ell(s,6.18,4.723,0.079,0.079,mix(C.navy,C.slate));
  icon(s,5.581,2.171,0.3,0.3,C.slate);
  icon(s,5.581,3.737,0.3,0.3,C.slate);
  rrect(s,0,0.825,4.375,3.975,[0,0.475,0.475,0],C.gray);
}

// ---------------------------------------------------------------- slide 10 — Education Overview
function slide10(s) {
  rect(s,5.768,0,2.917,5.625,C.navy);
  txt(s,'Education Overview',1.042,0.669,3.113,0.429,T.d21bBlack,{al:'center'});
  txt(s,L[1],1.04,1.202,3.116,0.47,T.o8Gray,{al:'center',ls:1.5});
  box(s,'roundRect',0.816,3.012,3.564,2.035,{fill:C.white,shadow:1,rad:0.3});
  box(s,'roundRect',0.816,1.977,3.564,1.535,{fill:C.white,shadow:1,rad:0.256});
  ell(s,1.211,2.208,0.75,0.75,C.navy);
  txt(s,'Institution',1.14,3.001,0.891,0.279,T.o9bSlate,{al:'center',ls:1.5});
  txt(s,'2018 - 2021',2.355,2.486,1.105,0.252,T.d11bSlate);
  txt(s,'Associate Degree',2.355,2.262,1.46,0.252,T.d11bSlate);
  txt(s,L[19],2.336,2.756,1.659,0.47,T.o8Gray,{ls:1.5});
  ell(s,1.211,3.727,0.75,0.75,C.navy);
  txt(s,'University',1.14,4.52,0.891,0.279,T.o9bSlate,{al:'center',ls:1.5});
  txt(s,'2021 - 2023',2.355,4.005,1.105,0.252,T.d11bSlate);
  txt(s,'Bachelor’s Degree',2.355,3.781,1.46,0.252,T.d11bSlate);
  txt(s,L[19],2.336,4.275,1.659,0.47,T.o8Gray,{ls:1.5});
  icon(s,1.361,2.358,0.45,0.45,C.white);
  icon(s,1.361,3.877,0.45,0.45,C.white);
  dots(s,6.734,5.257,3,1,0.414,0.157,C.white);
  dots(s,6.734,0.213,3,1,0.414,0.157,C.white);
  dots(s,9.369,2.544,4,4,0.158,0.06,C.navy);
  rrect(s,5.197,0.578,4.059,4.469,0.227,C.gray);
}

// ---------------------------------------------------------------- slide 11 — Education 1
function slide11(s) {
  ring(s,0.156,0,5.625,C.mist,80);
  rect(s,0,1.462,2.606,2.701,C.navy);
  dots(s,0.215,2.544,4,4,0.158,0.06,C.white);
  ell(s,8.655,4.63,1.99,1.99,[C.sage,80]);
  ell(s,8.655,-0.995,1.99,1.99,[C.sage,80]);
  txt(s,'Education 1',6.128,1.317,3.113,0.429,T.d21bBlack);
  txt(s,L[1],6.132,2.244,3.116,0.47,T.o8Gray,{ls:1.5});
  txt(s,'Institution Name',6.132,1.992,3.023,0.252,T.d11bSlate,{al:'justify'});
  txt(s,L[12],6.74,3.021,2.533,0.252,T.d11bSlate);
  txt(s,L[14],6.74,3.267,2.533,0.47,T.o8Gray,{ls:1.5});
  txt(s,L[0],7.014,3.839,2.188,0.215,T.o8Gray);
  txt(s,L[0],7.014,4.093,2.188,0.215,T.o8Gray);
  ell(s,6.859,3.907,0.079,0.079,mix(C.navy,C.slate));
  ell(s,6.859,4.161,0.079,0.079,mix(C.navy,C.slate));
  txt(s,L[0],7.014,4.358,2.188,0.215,T.o8Gray);
  ell(s,6.859,4.426,0.079,0.079,mix(C.navy,C.slate));
  ell(s,6.189,3.021,0.482,0.482,C.navy);
  icon(s,6.285,3.117,0.289,0.289,C.white);
  ell(s,0.938,0.781,4.062,4.062,C.gray);
}

// ---------------------------------------------------------------- slide 12 — Education 2
function slide12(s) {
  path(s,0,3.725,2.686,1.9,[M(0.873,0),B(1.874,0,2.686,0.812,2.686,1.813),[2.681,1.9],[2.268,1.9],[2.273,1.813],B(2.273,1.04,1.646,0.413,0.873,0.413),B(0.583,0.413,0.313,0.501,0.09,0.652),[0,0.719],[0,0.224],[0.008,0.219],B(0.265,0.079,0.56,0,0.873,0),Z],{fill:[C.mist,75.295]});
  path(s,8.28,0,1.72,1.307,[M(0,0),[0.435,0],[0.45,0.04],B(0.645,0.501,1.079,0.835,1.597,0.888),[1.72,0.894],[1.72,1.307],[1.555,1.299],B(0.824,1.224,0.221,0.715,0.009,0.034),Z],{fill:[C.mist,75.295]});
  box(s,'round2SameRect',5.304,0.191,4.15,5.242,{fill:C.navy,rot:270,rad:2.075});
  dots(s,0.394,0.4,3,1,0.414,0.157,C.navy);
  dots(s,9.275,5.068,4,3,0.158,0.06,C.navy);
  txt(s,'Education 2',0.94,1.194,3.113,0.429,T.d21bBlack);
  txt(s,L[1],0.944,2.106,3.116,0.47,T.o8Gray,{ls:1.5});
  txt(s,'Institution Name',0.944,1.854,3.023,0.252,T.d11bSlate,{al:'justify'});
  txt(s,L[12],1.556,2.879,2.533,0.252,T.d11bSlate);
  txt(s,L[14],1.556,3.125,2.533,0.47,T.o8Gray,{ls:1.5});
  txt(s,L[0],1.831,3.697,2.188,0.215,T.o8Gray);
  txt(s,L[0],1.831,3.952,2.188,0.215,T.o8Gray);
  ell(s,1.676,3.766,0.079,0.079,mix(C.navy,C.slate));
  ell(s,1.676,4.02,0.079,0.079,mix(C.navy,C.slate));
  txt(s,L[0],1.831,4.216,2.188,0.215,T.o8Gray);
  ell(s,1.676,4.284,0.079,0.079,mix(C.navy,C.slate));
  ell(s,0.944,2.879,0.482,0.482,C.navy);
  icon(s,1.04,2.976,0.289,0.289,C.white);
  path(s,5,0.96,5,3.704,[M(1.852,0),B(2.901,0,3.951,0,5,0),[5,3.704],[1.852,3.704],B(0.829,3.704,0,2.875,0,1.852),B(0,0.829,0.829,0,1.852,0),Z],{fill:C.gray});
}

// ---------------------------------------------------------------- slide 13 — Skills Proficiency
function slide13(s) {
  line(s,1.188,2.417,3.914,0,{line:{c:C.sage,w:0.75}});
  line(s,1.188,3.519,3.969,0,{line:{c:C.sage,w:0.75}});
  line(s,1.188,2.968,3.969,0.011,{line:{c:C.sage,w:0.75}});
  line(s,1.188,1.866,3.969,0,{line:{c:C.sage,w:0.75}});
  line(s,1.198,1.335,3.959,0,{line:{c:C.sage,w:0.75}});
  line(s,1.188,0.793,3.969,0,{line:{c:C.sage,w:0.75}});
  ell(s,0.208,5.058,0.042,0.042,C.navy);
  ell(s,0.319,5.17,0.042,0.042,C.navy);
  ell(s,0.208,5.17,0.042,0.042,C.navy);
  ell(s,0.43,5.281,0.042,0.042,C.navy);
  ell(s,0.319,5.281,0.042,0.042,C.navy);
  ell(s,0.208,5.281,0.042,0.042,C.navy);
  ell(s,0.541,5.393,0.042,0.042,C.navy);
  ell(s,0.43,5.393,0.042,0.042,C.navy);
  ell(s,0.319,5.393,0.042,0.042,C.navy);
  ell(s,0.208,5.393,0.042,0.042,C.navy);
  ell(s,0.541,0.195,0.042,0.042,C.navy);
  ell(s,0.43,0.195,0.042,0.042,C.navy);
  ell(s,0.319,0.195,0.042,0.042,C.navy);
  ell(s,0.208,0.195,0.042,0.042,C.navy);
  ell(s,0.43,0.306,0.042,0.042,C.navy);
  ell(s,0.319,0.306,0.042,0.042,C.navy);
  ell(s,0.208,0.306,0.042,0.042,C.navy);
  ell(s,0.319,0.418,0.042,0.042,C.navy);
  ell(s,0.208,0.418,0.042,0.042,C.navy);
  ell(s,0.208,0.529,0.042,0.042,C.navy);
  dots(s,9.417,5.058,4,4,0.111,0.042,C.navy,1);
  dots(s,9.417,0.195,4,4,0.111,0.042,C.navy,2);
  txt(s,'Skills Proficiency',6.001,0.71,3.113,0.429,T.d21bBlack);
  txt(s,L[10],6.001,1.304,3.116,0.678,T.o8Gray,{ls:1.5});
  rrect(s,5.998,2.272,3.119,2.643,0.284,C.gray);
  rect(s,1.259,1.734,0.234,2.336,C.navy);
  line(s,1.188,0.803,0,3.289,{line:{c:C.sage,w:0.75}});
  line(s,5.156,0.803,0,3.289,{line:{c:C.sage,w:0.75}});
  line(s,1.188,4.057,3.969,0.014,{line:{c:C.sage,w:0.75},rot:180,flipH:1});
  txt(s,'0',0.899,3.957,0.224,0.227,T.o9Gray3);
  txt(s,'1',0.899,3.406,0.224,0.227,T.o9Gray3);
  txt(s,'2',0.899,2.855,0.224,0.227,T.o9Gray3);
  txt(s,'3',0.899,2.304,0.224,0.227,T.o9Gray3);
  txt(s,'4',0.899,1.753,0.224,0.227,T.o9Gray3);
  txt(s,'5',0.899,1.213,0.224,0.227,T.o9Gray3);
  txt(s,'6',0.912,0.698,0.224,0.227,T.o9Gray3);
  rect(s,1.556,2.763,0.234,1.308,C.slate);
  rect(s,2.571,1.675,0.234,2.384,C.slate);
  rect(s,3.563,3.082,0.234,0.983,C.slate);
  rect(s,4.578,2.531,0.234,1.553,C.slate);
  rect(s,1.853,2.98,0.234,1.091,C.mist);
  rect(s,2.858,2.966,0.234,1.091,C.mist);
  rect(s,3.875,2.423,0.234,1.641,C.mist);
  rect(s,4.867,1.344,0.234,2.714,C.mist);
  txt(s,'Data 1',1.407,4.129,0.533,0.227,T.o9Gray3);
  txt(s,'Data 2',2.422,4.118,0.533,0.227,T.o9Gray3);
  txt(s,'Data 3',3.414,4.129,0.533,0.227,T.o9Gray3);
  txt(s,'Data 4',4.4,4.125,0.533,0.227,T.o9Gray3);
  txt(s,'Skill One',1.391,4.688,0.705,0.227,T.o9bNavy);
  txt(s,'Skill Two',2.859,4.631,0.704,0.227,T.o9bNavy);
  txt(s,'Skill Three',4.285,4.648,0.809,0.227,T.o9bNavy);
  box(s,'roundRect',2.51,4.66,0.213,0.213,{fill:C.navy,rad:0.036});
  box(s,'roundRect',3.947,4.655,0.213,0.213,{fill:C.mist,rad:0.036});
  box(s,'roundRect',1.073,4.698,0.213,0.213,{fill:C.navy,rad:0.036});
}

// ---------------------------------------------------------------- slide 14 — Certifications
function slide14(s) {
  ring(s,4.638,0.129,5.201,C.mist,80);
  box(s,'round2SameRect',5.327,0.786,3.824,4.839,{fill:C.navy,rad:1.912});
  dots(s,0.207,0.191,4,4,0.12,0.046,C.navy);
  swoosh(s,0,5.067,{rot:180,flipH:1});
  txt(s,'Certifications',0.924,1.398,3.113,0.429,T.d21bBlack);
  txt(s,L[1],0.924,1.893,3.116,0.47,T.o8Gray,{ls:1.5});
  txt(s,'Certificate Title',1.531,2.641,1.371,0.252,T.d11bSlate);
  txt(s,'(2019)',1.531,3.094,0.551,0.227,T.o9bSlate);
  txt(s,L[20],1.531,2.863,2.34,0.215,T.o8Gray);
  icon(s,0.924,2.641,0.525,0.525,C.slate);
  txt(s,'Certificate Title',1.528,3.547,1.371,0.252,T.d11bSlate);
  txt(s,'(2023)',1.528,4,0.551,0.227,T.o9bSlate);
  txt(s,L[20],1.528,3.769,2.34,0.215,T.o8Gray);
  icon(s,0.92,3.547,0.525,0.525,C.slate);
  path(s,5.446,0.911,3.586,4.714,[M(1.793,0),B(2.783,0,3.586,0.803,3.586,1.793),B(3.586,2.767,3.586,3.74,3.586,4.714),[0,4.714],[0,1.793],B(0,0.803,0.803,0,1.793,0),Z],{fill:C.gray});
}

// ---------------------------------------------------------------- slide 15 — Projects Overview
function slide15(s) {
  ring(s,0.204,0.212,5.201,C.mist,75.295);
  txt(s,'Projects Overview',6.065,0.827,3.113,0.429,T.d21bBlack);
  txt(s,L[10],6.065,1.437,3.116,0.678,T.o8Gray,{ls:1.5});
  box(s,'roundRect',6.065,2.477,3.158,2.341,{fill:C.white,shadow:1,rad:0.222});
  txt(s,L[0],6.727,2.859,2.188,0.215,T.o8Gray);
  txt(s,L[0],6.727,3.196,2.188,0.215,T.o8Gray);
  txt(s,L[0],6.727,4.222,2.188,0.215,T.o8Gray);
  txt(s,'01',6.373,2.853,0.385,0.227,T.d9Navy,{al:'center'});
  txt(s,'02',6.373,3.19,0.385,0.227,T.d9Navy,{al:'center'});
  txt(s,'05',6.373,4.216,0.385,0.227,T.d9Navy,{al:'center'});
  txt(s,L[0],6.727,3.534,2.188,0.215,T.o8Gray);
  txt(s,L[0],6.727,3.871,2.188,0.215,T.o8Gray);
  txt(s,'03',6.373,3.528,0.385,0.227,T.d9Navy,{al:'center'});
  txt(s,'04',6.373,3.865,0.385,0.227,T.d9Navy,{al:'center'});
  box(s,'roundRect',2.846,0.785,2.385,3.991,{fill:C.navy,rad:0.271});
  box(s,'roundRect',0.311,0.785,2.385,3.991,{fill:C.navy,rad:0.271});
  dots(s,8.674,0.245,3,1,0.414,0.157,C.navy);
  dots(s,0.341,5.163,3,1,0.414,0.157,C.navy);
  dots(s,9.417,5.058,4,4,0.111,0.042,C.navy,1);
  rrect(s,0.367,0.849,2.385,3.991,0.271,C.gray);
  rrect(s,2.903,0.849,2.385,3.991,0.271,C.photoB);
}

// ---------------------------------------------------------------- slide 16 — Project 1
function slide16(s) {
  box(s,'round2SameRect',6.081,0.752,3.716,4.121,{fill:C.navy,rot:270,rad:0.286});
  txt(s,'Project 1',0.895,1.203,3.113,0.429,T.d21bBlack);
  txt(s,L[1],0.898,2.129,3.116,0.47,T.o8Gray,{ls:1.5});
  txt(s,'Team Management',0.898,1.877,3.023,0.252,T.d11bSlate,{al:'justify'});
  txt(s,'Two Months',1.883,2.938,1.46,0.252,T.d11bSlate);
  txt(s,L[13],1.864,3.166,2.151,0.47,T.o8Gray,{ls:1.5});
  box(s,'roundRect',0.898,2.886,0.75,0.75,{fill:C.navy,rad:0.125});
  icon(s,1.048,3.036,0.45,0.45,C.white);
  txt(s,L[8],0.895,3.744,3.027,0.678,T.o8Gray,{ls:1.5});
  swoosh(s,0,0,{});
  swoosh(s,0,5.067,{rot:180,flipH:1});
  dots(s,9.275,0.204,4,4,0.158,0.06,C.navy);
  dots(s,9.275,4.888,4,4,0.158,0.06,C.navy);
  rrect(s,4.909,1.161,1.611,3.303,0.238,C.photoB);
  rrect(s,6.604,1.161,1.611,3.303,0.238,C.gray);
  rrect(s,8.298,1.161,1.611,3.303,0.238,C.photoB);
}

// ---------------------------------------------------------------- slide 17 — Project 2
function slide17(s) {
  box(s,'rect',0,0,5.152,5.625,{fill:[C.mist,75.295],flipH:1});
  txt(s,'Project 2',6.016,1.203,3.113,0.429,T.d21bBlack);
  txt(s,L[1],6.019,2.129,3.116,0.47,T.o8Gray,{ls:1.5});
  txt(s,'Team Leader',6.019,1.877,3.023,0.252,T.d11bSlate,{al:'justify'});
  txt(s,'Three Months',7.004,2.938,1.46,0.252,T.d11bSlate);
  txt(s,L[13],6.985,3.166,2.151,0.47,T.o8Gray,{ls:1.5});
  box(s,'roundRect',6.019,2.886,0.75,0.75,{fill:C.navy,rad:0.125});
  icon(s,6.169,3.036,0.45,0.45,C.white);
  txt(s,L[8],6.016,3.744,3.027,0.678,T.o8Gray,{ls:1.5});
  dots(s,4.872,2.39,1,3,0.355,0.135,C.navy);
  dots(s,9.417,5.058,4,4,0.111,0.042,C.navy,1);
  dots(s,9.417,0.195,4,4,0.111,0.042,C.navy,2);
  path(s,0.424,0.394,4.298,2.356,[M(0.393,0),[3.905,0],B(4.122,0,4.298,0.176,4.298,0.393),[4.298,1.963],B(4.298,2.18,4.122,2.356,3.905,2.356),[0.393,2.356],B(0.203,2.356,0.045,2.221,0.008,2.042),[0,1.963],[0,0.393],[0.008,0.314],B(0.045,0.135,0.203,0,0.393,0),Z],{fill:C.gray});
  path(s,0.424,2.871,4.298,2.345,[M(0.393,0),[3.905,0],B(4.122,0,4.298,0.176,4.298,0.393),[4.298,1.963],B(4.298,2.126,4.199,2.265,4.058,2.325),[3.993,2.345],[0.305,2.345],[0.24,2.325],B(0.122,2.275,0.034,2.17,0.008,2.042),[0,1.963],[0,0.393],[0.008,0.314],B(0.045,0.135,0.203,0,0.393,0),Z],{fill:C.photoB});
}

// ---------------------------------------------------------------- slide 18 — Achievements  & Awards
function slide18(s) {
  ring(s,-1.67,-2.177,3.957,C.mist,80);
  box(s,'rect',0,3.985,10,1.64,{fill:C.white,shadow:1});
  txt(s,['Achievements ','& Awards'],0.75,0.458,2.234,0.783,T.d21bBlack);
  box(s,'roundRect',5,1.725,4.5,3.341,{fill:C.white,shadow:1,rad:0.352});
  ell(s,5.548,2.992,0.106,0.106,mix(C.navy,C.slate));
  txt(s,'Awards Title',6.008,2.829,1.371,0.252,T.d11bSlate);
  txt(s,'Organizations',6.008,3,1.105,0.262,T.o8Gray,{ls:1.5});
  ell(s,5.548,3.641,0.106,0.106,mix(C.navy,C.slate));
  txt(s,'Awards Title',6.008,3.477,1.371,0.252,T.d11bSlate);
  txt(s,'Organizations',6.008,3.648,1.105,0.262,T.o8Gray,{ls:1.5});
  ell(s,5.548,4.297,0.106,0.106,mix(C.navy,C.slate));
  txt(s,'Awards Title',6.008,4.133,1.371,0.252,T.d11bSlate);
  txt(s,'Organizations',6.008,4.305,1.105,0.262,T.o8Gray,{ls:1.5});
  ell(s,5.548,2.389,0.106,0.106,mix(C.navy,C.slate));
  txt(s,'Awards Title',6.008,2.225,1.371,0.252,T.d11bSlate);
  txt(s,'Organizations',6.008,2.396,1.105,0.262,T.o8Gray,{ls:1.5});
  txt(s,'11/03/2022',7.339,4.133,1.669,0.227,T.o9bSlate);
  txt(s,L[3],7.339,4.305,1.669,0.262,T.o8Gray,{ls:1.5});
  txt(s,'02/04/2021',7.339,3.477,1.669,0.227,T.o9bSlate);
  txt(s,L[3],7.339,3.648,1.669,0.262,T.o8Gray,{ls:1.5});
  txt(s,'03/12/2020',7.339,2.829,1.669,0.227,T.o9bSlate);
  txt(s,L[3],7.339,3,1.669,0.262,T.o8Gray,{ls:1.5});
  txt(s,'12/06/2019',7.339,2.225,1.669,0.227,T.o9bSlate);
  txt(s,L[3],7.339,2.396,1.669,0.262,T.o8Gray,{ls:1.5});
  txt(s,L[8],3.357,0.51,3.027,0.678,T.o8Gray,{ls:1.5});
  txt(s,L[15],6.765,0.51,2.485,0.678,T.o8Gray,{ls:1.5});
  dots(s,9.075,1.58,4,4,0.157,0.06,C.navy);
  rrect(s,0.5,1.725,4.235,3.341,0.352,C.gray);
}

// ---------------------------------------------------------------- slide 19 — Portfolio Showcase
function slide19(s) {
  rect(s,0,2.281,10,2.141,[C.mist,75.295]);
  box(s,'roundRect',0.402,2.117,2.821,2.821,{fill:C.navy,rad:0.337});
  box(s,'roundRect',6.777,2.117,2.821,2.821,{fill:C.navy,rad:0.337});
  box(s,'roundRect',3.59,2.117,2.821,2.821,{fill:C.navy,rad:0.337});
  txt(s,'Portfolio Showcase',3.444,0.382,3.113,0.429,T.d21bBlack,{al:'center'});
  txt(s,L[10],2.484,0.882,5.031,0.47,T.o8Gray,{al:'center',ls:1.5});
  txt(s,L[2],3.83,4.286,2.34,0.47,T.o8White,{ls:1.5});
  txt(s,'Portofolio Two',3.83,4.033,2.34,0.252,T.d11bWhite,{al:'justify'});
  txt(s,L[2],7.018,4.286,2.34,0.47,T.o8White,{ls:1.5});
  txt(s,'Portofolio Three',7.018,4.033,2.34,0.252,T.d11bWhite,{al:'justify'});
  txt(s,L[2],0.642,4.286,2.34,0.47,T.o8White,{ls:1.5});
  txt(s,'Portofolio One',0.642,4.033,2.34,0.252,T.d11bWhite,{al:'justify'});
  dots(s,4.508,5.21,3,1,0.413,0.157,C.navy);
  ell(s,0.541,0.195,0.042,0.042,C.navy);
  ell(s,0.43,0.195,0.042,0.042,C.navy);
  ell(s,0.319,0.195,0.042,0.042,C.navy);
  ell(s,0.208,0.195,0.042,0.042,C.navy);
  ell(s,0.43,0.306,0.042,0.042,C.navy);
  ell(s,0.319,0.306,0.042,0.042,C.navy);
  ell(s,0.208,0.306,0.042,0.042,C.navy);
  ell(s,0.319,0.418,0.042,0.042,C.navy);
  ell(s,0.208,0.418,0.042,0.042,C.navy);
  ell(s,0.208,0.529,0.042,0.042,C.navy);
  dots(s,9.417,0.195,4,4,0.111,0.042,C.navy,2);
  rrect(s,0.402,1.766,2.821,2.125,0.309,C.photoB);
  rrect(s,3.59,1.766,2.821,2.125,0.309,C.gray);
  rrect(s,6.777,1.766,2.821,2.125,0.293,C.photoB);
}

// ---------------------------------------------------------------- slide 20 — Testimonials
function slide20(s) {
  box(s,'rect',0,0,4.174,5.625,{fill:[C.mist,75.295],flipH:1});
  box(s,'roundRect',0.508,0.703,5.008,2.013,{fill:C.white,shadow:1,rad:0.335});
  box(s,'roundRect',0.508,2.909,5.008,2.013,{fill:C.white,shadow:1,rad:0.335});
  txt(s,'Ferdinand Bridge',2.857,3.295,1.841,0.278,T.d12bNavy);
  txt(s,'Manager',2.857,3.549,1.796,0.227,T.o9Slate);
  txt(s,L[17],2.857,3.857,2.235,0.678,T.o8iGray,{ls:1.5});
  txt(s,'Brian Mendoza',2.857,1.09,1.841,0.278,T.d12bNavy);
  txt(s,'Lecturer',2.857,1.343,1.796,0.227,T.o9Slate);
  txt(s,L[17],2.857,1.651,2.235,0.678,T.o8iGray,{ls:1.5});
  dots(s,9.275,0.204,4,4,0.158,0.06,C.navy);
  dots(s,9.275,4.888,4,4,0.158,0.06,C.navy);
  txt(s,'Testimonials',6.316,1.923,2.853,0.429,T.d21bBlack);
  txt(s,L[9],6.316,2.478,2.884,0.678,T.o8Gray,{ls:1.5});
  txt(s,L[18],6.316,3.232,2.884,0.47,T.o8Gray,{ls:1.5});
  dots(s,1.026,0.289,3,1,0.378,0.143,C.navy);
  dots(s,1.026,5.21,3,1,0.378,0.143,C.navy);
  rrect(s,0.607,0.798,1.823,1.824,0.304,C.gray);
  rrect(s,0.607,3.003,1.823,1.824,0.304,C.photoB);
}

// ---------------------------------------------------------------- slide 21 — Language
function slide21(s) {
  rect(s,5.939,0,2.727,5.625,[C.mist,75.295]);
  txt(s,'English',1.054,2.966,1.104,0.252,T.o11bSlate);
  rect(s,2.609,3.042,1.561,0.102,C.silver);
  rect(s,2.288,3.042,1.675,0.102,C.slate);
  txt(s,'Spanish',1.054,3.482,1.104,0.252,T.o11bSlate);
  rect(s,2.609,3.557,1.561,0.102,C.silver);
  rect(s,2.288,3.557,1.16,0.102,C.slate);
  txt(s,'Language',1.05,1.375,3.113,0.429,T.d21bBlack);
  txt(s,L[9],1.054,1.961,3.116,0.678,T.o8Gray,{ls:1.5});
  txt(s,'Indonesian',1.054,3.997,1.104,0.252,T.o11bSlate);
  rect(s,2.609,4.073,1.561,0.102,C.silver);
  rect(s,2.288,4.073,1.675,0.102,C.slate);
  ell(s,5.22,0.729,4.167,4.167,C.navy);
  ring(s,-1.433,-3.05,3.957,C.mist,75.295);
  ell(s,0.129,4.978,0.06,0.06,C.navy);
  ell(s,0.287,5.137,0.06,0.06,C.navy);
  ell(s,0.129,5.137,0.06,0.06,C.navy);
  ell(s,0.445,5.296,0.06,0.06,C.navy);
  ell(s,0.287,5.296,0.06,0.06,C.navy);
  ell(s,0.129,5.296,0.06,0.06,C.navy);
  ell(s,0.603,5.455,0.06,0.06,C.navy);
  ell(s,0.445,5.455,0.06,0.06,C.navy);
  ell(s,0.287,5.455,0.06,0.06,C.navy);
  ell(s,0.129,5.455,0.06,0.06,C.navy);
  dots(s,6.811,0.289,3,1,0.413,0.157,C.navy);
  dots(s,6.811,5.21,3,1,0.413,0.157,C.navy);
  ell(s,5.372,0.881,3.864,3.864,C.gray);
}

// ---------------------------------------------------------------- slide 22 — Hobbies & Interests
function slide22(s) {
  box(s,'round2SameRect',-1.278,1.278,5.625,3.068,{fill:[C.mist,75.295],rot:90,rad:0.511});
  txt(s,'Hobbies & Interests',5.548,0.864,3.489,0.48,T.d24bBlack);
  txt(s,'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud. ',5.548,1.502,3.531,0.678,T.o8Gray,{ls:1.5});
  ell(s,5.617,2.573,0.943,0.943,C.navy);
  txt(s,L[2],6.752,2.936,2.399,0.47,T.o8Gray,{ls:1.5});
  txt(s,'Basketball',6.752,2.684,2.327,0.252,T.d11bSlate,{al:'justify'});
  icon(s,5.864,2.82,0.45,0.45,C.white);
  ell(s,5.617,3.797,0.943,0.943,C.navy);
  txt(s,L[2],6.752,4.16,2.399,0.47,T.o8Gray,{ls:1.5});
  txt(s,'Comic Collections',6.752,3.907,2.327,0.252,T.d11bSlate,{al:'justify'});
  icon(s,5.864,4.043,0.45,0.45,C.white);
  ell(s,8.754,4.739,1.99,1.99,[C.sage,80]);
  ell(s,8.754,-1.11,1.99,1.99,[C.sage,80]);
  dots(s,9.711,2.39,1,3,0.355,0.135,C.navy);
  rrect(s,0.53,0.79,4.169,4.045,0.371,C.gray);
}

// ---------------------------------------------------------------- slide 23 — References
function slide23(s) {
  ring(s,-0.157,0.055,3.11,C.mist,75.295);
  box(s,'round2SameRect',1.016,0.21,4.11,5.205,{fill:C.white,rot:270,shadow:1,rad:0.685});
  txt(s,'Lorem ipsum dolor sit amet, consectetur adipis elit, sed do eiusmod tempor incididunt. ',1.265,1.755,2.714,0.47,T.o8Gray,{ls:1.5});
  txt(s,'Contact Information',1.266,2.343,2.544,0.252,T.d11bSlate,{al:'justify'});
  txt(s,'Julian Maxwell',2.388,2.759,1.446,0.262,T.o8Gray,{ls:1.5});
  txt(s,'Name',1.266,2.771,1.066,0.252,T.o11Black);
  txt(s,'Rograde Comp.',2.396,3.458,1.413,0.262,T.o8Gray,{ls:1.5});
  txt(s,'Company',1.266,3.481,1.066,0.252,T.o11Black);
  txt(s,'Graphic Designer',2.403,3.109,1.339,0.262,T.o8Gray,{ls:1.5});
  txt(s,'Title',1.266,3.118,1.066,0.252,T.o11Black);
  txt(s,':',2.248,2.771,0.194,0.252,T.o11bBlack,{al:'center'});
  txt(s,':',2.248,3.116,0.194,0.252,T.o11bBlack,{al:'center'});
  txt(s,':',2.248,3.477,0.194,0.252,T.o11bBlack,{al:'center'});
  txt(s,'References',1.265,1.176,2.714,0.48,T.d24bBlack);
  txt(s,L[21],2.396,4.173,1.583,0.262,T.o8Gray,{ls:1.5});
  txt(s,'Email',1.266,4.197,1.066,0.252,T.o11Black);
  txt(s,'(+62) 8123 4567 790',2.403,3.823,1.339,0.262,T.o8Gray,{ls:1.5});
  txt(s,'Phone',1.266,3.832,1.066,0.252,T.o11Black);
  txt(s,':',2.248,3.83,0.194,0.252,T.o11bBlack,{al:'center'});
  txt(s,':',2.25,4.192,0.194,0.252,T.o11bBlack,{al:'center'});
  dots(s,9.417,5.058,4,4,0.111,0.042,C.navy,1);
  dots(s,9.417,0.195,4,4,0.111,0.042,C.navy,2);
  dots(s,0.704,4.335,1,3,0.356,0.135,C.navy);
  rrect(s,4.682,0.758,5.318,4.11,[0.435,0,0,0.435],C.gray);
}

// ---------------------------------------------------------------- slide 24 — Cover Letter
function slide24(s) {
  txt(s,'Dear Hiring Manager,',0.873,1.257,2.544,0.252,T.d11Slate,{al:'justify'});
  txt(s,'Body Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim ve, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.',0.873,1.697,8.254,0.678,T.o8Gray,{ls:1.5});
  txt(s,'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum.',0.873,2.408,8.254,0.47,T.o8Gray,{ls:1.5});
  txt(s,'Closing Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in repr.',0.873,2.939,8.254,0.47,T.o8Gray,{ls:1.5});
  txt(s,'Julian Maxwell',0.873,4.491,1.668,0.252,T.d11Slate);
  txt(s,'Sincerely,',0.873,3.777,1.668,0.252,T.d11Slate);
  txt(s,'Cover Letter',3.444,0.491,3.113,0.429,T.d21bBlack,{al:'center'});
  swoosh(s,8.071,0,{flipH:1});
  swoosh(s,8.071,5.067,{rot:180});
  swoosh(s,0,0,{});
  swoosh(s,0,5.067,{rot:180,flipH:1});
  dots(s,4.508,0.14,3,1,0.413,0.157,C.mist,0,75.295);
  dots(s,4.508,5.328,3,1,0.413,0.157,C.mist,0,75.295);
}

// ---------------------------------------------------------------- slide 25 — Visual  Timeline
function slide25(s) {
  box(s,'roundRect',0.702,1.901,1.891,3.102,{fill:C.white,shadow:1,rad:0.315});
  box(s,'roundRect',2.938,1.901,1.891,3.102,{fill:C.white,shadow:1,rad:0.315});
  box(s,'roundRect',5.177,1.901,1.891,3.102,{fill:C.white,shadow:1,rad:0.315});
  box(s,'roundRect',7.406,1.901,1.891,3.102,{fill:C.white,shadow:1,rad:0.315});
  line(s,0,2.351,10,0,{line:{c:C.navy,w:1.5,d:'lgDash'}});
  txt(s,'2024',7.666,2.155,1.37,0.391,T.d14bWhite,{al:'center',ls:1,va:'middle',fill:C.slate,shape:'roundRect',rad:0.195});
  txt(s,L[5],7.521,2.985,1.662,0.673,T.o8Gray,{al:'center',ls:1.5});
  txt(s,'Timeline Four',7.521,2.747,1.662,0.252,T.o11bSlate,{al:'center'});
  txt(s,L[5],5.286,2.985,1.662,0.673,T.o8Gray,{al:'center',ls:1.5});
  txt(s,'Timeline Three',5.286,2.747,1.662,0.252,T.o11bNavy,{al:'center'});
  txt(s,L[5],3.052,2.985,1.662,0.673,T.o8Gray,{al:'center',ls:1.5});
  txt(s,'Timeline Two',3.052,2.747,1.662,0.252,T.o11bSlate,{al:'center'});
  txt(s,L[5],0.818,2.985,1.662,0.673,T.o8Gray,{al:'center',ls:1.5});
  txt(s,'Timeline One',0.818,2.747,1.662,0.252,T.o11bNavy,{al:'center'});
  box(s,'ellipse',7.906,3.901,0.891,0.891,{fill:C.white,shadow:1});
  icon(s,8.119,4.114,0.465,0.465,C.slate);
  box(s,'ellipse',5.672,3.901,0.891,0.891,{fill:C.white,shadow:1});
  icon(s,5.885,4.114,0.465,0.465,C.navy);
  box(s,'ellipse',3.438,3.901,0.891,0.891,{fill:C.white,shadow:1});
  icon(s,3.65,4.114,0.465,0.465,C.slate);
  box(s,'ellipse',1.203,3.901,0.891,0.891,{fill:C.white,shadow:1});
  icon(s,1.416,4.114,0.465,0.465,C.navy);
  txt(s,['Visual ','Timeline'],0.75,0.622,2.234,0.783,T.d21bBlack);
  txt(s,L[8],3.357,0.674,3.027,0.678,T.o8Gray,{ls:1.5});
  txt(s,L[15],6.765,0.674,2.485,0.678,T.o8Gray,{ls:1.5});
  txt(s,'2022',5.437,2.155,1.37,0.391,T.d14bWhite,{al:'center',ls:1,va:'middle',fill:C.navy,shape:'roundRect',rad:0.195});
  txt(s,'2020',3.198,2.155,1.37,0.391,T.d14bWhite,{al:'center',ls:1,va:'middle',fill:C.slate,shape:'roundRect',rad:0.195});
  txt(s,'2019',0.963,2.155,1.37,0.391,T.d14bWhite,{al:'center',va:'middle',fill:C.navy,shape:'roundRect',rad:0.195});
  ell(s,0.208,5.058,0.042,0.042,C.navy);
  ell(s,0.319,5.17,0.042,0.042,C.navy);
  ell(s,0.208,5.17,0.042,0.042,C.navy);
  ell(s,0.43,5.281,0.042,0.042,C.navy);
  ell(s,0.319,5.281,0.042,0.042,C.navy);
  ell(s,0.208,5.281,0.042,0.042,C.navy);
  ell(s,0.541,5.393,0.042,0.042,C.navy);
  ell(s,0.43,5.393,0.042,0.042,C.navy);
  ell(s,0.319,5.393,0.042,0.042,C.navy);
  ell(s,0.208,5.393,0.042,0.042,C.navy);
  ell(s,0.541,0.195,0.042,0.042,C.navy);
  ell(s,0.43,0.195,0.042,0.042,C.navy);
  ell(s,0.319,0.195,0.042,0.042,C.navy);
  ell(s,0.208,0.195,0.042,0.042,C.navy);
  ell(s,0.43,0.306,0.042,0.042,C.navy);
  ell(s,0.319,0.306,0.042,0.042,C.navy);
  ell(s,0.208,0.306,0.042,0.042,C.navy);
  ell(s,0.319,0.418,0.042,0.042,C.navy);
  ell(s,0.208,0.418,0.042,0.042,C.navy);
  ell(s,0.208,0.529,0.042,0.042,C.navy);
  dots(s,9.417,5.058,4,4,0.111,0.042,C.navy,1);
  dots(s,9.417,0.195,4,4,0.111,0.042,C.navy,2);
}

// ---------------------------------------------------------------- slide 26 — Contact Me
function slide26(s) {
  ell(s,-0.934,4.739,1.99,1.99,[C.sage,80]);
  ell(s,-0.934,-1.11,1.99,1.99,[C.sage,80]);
  box(s,'round2SameRect',5.689,0.741,4.477,4.144,{fill:C.white,rot:270,shadow:1,rad:0.445});
  txt(s,'Body Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod incididunt ut labore et dolore magna aliqua.',1.113,1.176,3.645,0.47,T.o8Gray,{ls:1.5});
  txt(s,'Contact Me',1.113,0.678,2.234,0.429,T.d21bBlack);
  icon(s,6.931,1.9,0.45,0.45,C.slate);
  txt(s,'(+62)8123 4567790',7.595,2.114,1.612,0.227,T.o9Gray);
  txt(s,'Contact Number',7.584,1.876,1.329,0.251,T.d9Black,{shape:'roundRect',rad:0.042});
  icon(s,6.931,1.196,0.45,0.45,C.slate);
  txt(s,'info@julianmaxwell.com',7.604,1.419,1.76,0.227,T.o9Gray);
  txt(s,'Email Address',7.584,1.196,1.329,0.251,T.d9Black,{shape:'roundRect',rad:0.042});
  icon(s,6.931,3.315,0.375,0.375,C.slate);
  icon(s,6.931,2.615,0.375,0.375,C.slate);
  icon(s,6.931,3.972,0.45,0.45,C.slate);
  txt(s,'Julian Maxwell',7.604,2.808,1.612,0.227,T.o9Gray);
  txt(s,'Facebook',7.592,2.57,1.329,0.251,T.d9Black,{shape:'roundRect',rad:0.042});
  txt(s,'@Julian_Maxwell',7.615,3.507,1.612,0.227,T.o9Gray);
  txt(s,'X/Twitter',7.604,3.27,1.329,0.251,T.d9Black,{shape:'roundRect',rad:0.042});
  txt(s,'Julian Maxwell',7.615,4.202,1.612,0.227,T.o9Gray);
  txt(s,'Linkedin',7.604,3.964,1.329,0.251,T.d9Black,{shape:'roundRect',rad:0.042});
  dots(s,9.192,4.797,4,4,0.158,0.06,C.navy);
  rrect(s,0.818,1.901,5.455,3.15,[0.358,0,0,0.358],C.gray);
}

// ---------------------------------------------------------------- slide 27 — Customizable Icons
function slide27(s) {
  txt(s,'Customizable Icons',3.444,0.491,3.113,0.429,T.d21bBlack,{al:'center'});
  dots(s,4.508,0.14,3,1,0.413,0.157,C.mist,0,75.295);
  dots(s,4.508,5.328,3,1,0.413,0.157,C.mist,0,75.295);
  swoosh(s,-0.686,0.686,{rot:270,flipH:1});
  swoosh(s,-0.686,4.381,{rot:270});
  swoosh(s,8.757,0.686,{rot:90});
  swoosh(s,8.757,4.381,{rot:90,flipH:1});
  icon(s,4.787,1.55,0.45,0.45,C.navy);
  icon(s,3.577,2.802,0.45,0.45,C.navy);
  icon(s,4.787,2.802,0.45,0.45,C.navy);
  icon(s,5.995,1.55,0.45,0.45,C.navy);
  icon(s,3.577,1.536,0.45,0.45,C.navy);
  icon(s,2.368,1.55,0.45,0.45,C.navy);
  icon(s,2.346,2.802,0.45,0.45,C.navy);
  icon(s,7.204,1.55,0.45,0.45,C.navy);
  icon(s,5.999,2.802,0.45,0.45,C.navy);
  icon(s,7.204,2.802,0.45,0.45,C.navy);
  icon(s,3.577,4.054,0.45,0.45,C.navy);
  icon(s,2.367,4.054,0.45,0.45,C.navy);
  icon(s,4.787,4.053,0.45,0.45,C.navy);
  icon(s,5.995,4.053,0.45,0.45,C.navy);
  icon(s,7.204,4.053,0.45,0.45,C.navy);
}

// ---------------------------------------------------------------- slide 28 — Graphs & Charts
function slide28(s) {
  rect(s,0,0,3.633,5.625,C.navy);
  box(s,'round2SameRect',0.324,0.024,4.93,5.577,{fill:C.white,rot:90,shadow:1,rad:0.564});
  ell(s,8.754,4.739,1.99,1.99,[C.sage,80]);
  ell(s,8.754,-1.11,1.99,1.99,[C.sage,80]);
  dots(s,9.711,2.39,1,3,0.355,0.135,C.sage,0,80);
  txt(s,'Graphs & Charts',6.324,1.434,2.853,0.429,T.d21bBlack);
  txt(s,L[9],6.324,2.091,2.884,0.678,T.o8Gray,{ls:1.5});
  txt(s,'Skill One',6.396,3.913,0.981,0.278,T.d12Coal);
  txt(s,'Skill Two',8.156,3.913,0.981,0.278,T.d12Coal);
  line(s,6.396,3.736,0.946,0,{line:{c:C.navy,w:4}});
  line(s,8.156,3.736,0.946,0,{line:{c:C.sage,w:4}});
  txt(s,L[18],6.324,2.849,2.884,0.47,T.o8Gray,{ls:1.5});
  txt(s,'Jan',0.634,2.589,0.356,0.232,T.o8Gray2);
  txt(s,'Feb',1.03,2.589,0.377,0.232,T.o8Gray2);
  txt(s,'Mar',1.426,2.589,0.402,0.232,T.o8Gray2);
  txt(s,'Apr',1.821,2.589,0.372,0.232,T.o8Gray2);
  txt(s,'May',2.217,2.589,0.414,0.232,T.o8Gray2);
  txt(s,'0',0.343,2.435,0.252,0.232,T.o8Gray2);
  txt(s,'1',0.343,2.221,0.252,0.232,T.o8Gray2);
  txt(s,'2',0.343,2.019,0.252,0.232,T.o8Gray2);
  txt(s,'3',0.343,1.811,0.252,0.232,T.o8Gray2);
  txt(s,'4',0.343,1.597,0.252,0.232,T.o8Gray2);
  txt(s,'5',0.343,1.389,0.252,0.232,T.o8Gray2);
  txt(s,'6',0.343,1.197,0.252,0.232,T.o8Gray2);
  txt(s,'7',0.343,0.984,0.252,0.232,T.o8Gray2);
  txt(s,'8',0.343,0.77,0.252,0.232,T.o8Gray2);
  line(s,0.611,2.559,2.001,0,{line:{c:C.smoke,w:0.75}});
  line(s,0.609,2.346,2.001,0,{line:{c:C.smoke,w:0.75}});
  line(s,0.588,2.142,2.001,0,{line:{c:C.smoke,w:0.75}});
  line(s,0.602,1.939,2.001,0,{line:{c:C.smoke,w:0.75}});
  line(s,0.602,1.735,2.001,0,{line:{c:C.smoke,w:0.75}});
  line(s,0.585,1.526,2.001,0,{line:{c:C.smoke,w:0.75}});
  line(s,0.583,1.313,2.001,0,{line:{c:C.smoke,w:0.75}});
  line(s,0.602,1.109,2.001,0,{line:{c:C.smoke,w:0.75}});
  line(s,0.602,0.911,2.001,0,{line:{c:C.smoke,w:0.75}});
  line(s,0.609,0.911,0,1.649,{line:{c:C.smoke,w:0.75}});
  rect(s,1.135,2.019,0.129,0.532,C.navy);
  rect(s,1.502,1.819,0.129,0.73,C.navy);
  rect(s,1.888,1.139,0.129,1.426,C.navy);
  rect(s,2.289,1.546,0.129,1.005,C.navy);
  rect(s,0.742,1.663,0.129,0.888,C.navy);
  txt(s,'Jan',3.224,2.589,0.356,0.232,T.o8Gray2);
  txt(s,'Feb',3.62,2.589,0.377,0.232,T.o8Gray2);
  txt(s,'Mar',4.016,2.589,0.402,0.232,T.o8Gray2);
  txt(s,'Apr',4.412,2.589,0.372,0.232,T.o8Gray2);
  txt(s,'May',4.807,2.589,0.414,0.232,T.o8Gray2);
  txt(s,'0',2.933,2.435,0.252,0.232,T.o8Gray2);
  txt(s,'1',2.933,2.221,0.252,0.232,T.o8Gray2);
  txt(s,'2',2.933,2.019,0.252,0.232,T.o8Gray2);
  txt(s,'3',2.933,1.811,0.252,0.232,T.o8Gray2);
  txt(s,'4',2.933,1.597,0.252,0.232,T.o8Gray2);
  txt(s,'5',2.933,1.389,0.252,0.232,T.o8Gray2);
  txt(s,'6',2.933,1.197,0.252,0.232,T.o8Gray2);
  txt(s,'7',2.933,0.984,0.252,0.232,T.o8Gray2);
  txt(s,'8',2.933,0.77,0.252,0.232,T.o8Gray2);
  line(s,3.201,2.559,2.001,0,{line:{c:C.smoke,w:0.75}});
  line(s,3.199,2.346,2.001,0,{line:{c:C.smoke,w:0.75}});
  line(s,3.178,2.142,2.001,0,{line:{c:C.smoke,w:0.75}});
  line(s,3.192,1.939,2.001,0,{line:{c:C.smoke,w:0.75}});
  line(s,3.192,1.735,2.001,0,{line:{c:C.smoke,w:0.75}});
  line(s,3.175,1.526,2.001,0,{line:{c:C.smoke,w:0.75}});
  line(s,3.174,1.313,2.001,0,{line:{c:C.smoke,w:0.75}});
  line(s,3.192,1.109,2.001,0,{line:{c:C.smoke,w:0.75}});
  line(s,3.192,0.911,2.001,0,{line:{c:C.smoke,w:0.75}});
  line(s,3.199,0.911,0,1.649,{line:{c:C.smoke,w:0.75}});
  rect(s,3.725,2.019,0.129,0.532,C.sage);
  rect(s,4.092,1.819,0.129,0.73,C.sage);
  rect(s,4.478,1.139,0.129,1.426,C.sage);
  rect(s,4.879,1.546,0.129,1.005,C.sage);
  rect(s,3.332,1.663,0.129,0.888,C.sage);
  box(s,'blockArc',2.981,3.445,1.194,1.194,{fill:C.navy,line:{c:C.white,w:2.25},arc:[0,359.9],thick:0.219});
  box(s,'blockArc',2.984,3.448,1.194,1.194,{fill:C.mist,line:{c:C.white,w:2.25},rot:327.49,arc:[209.132,300.946],thick:0.234});
  box(s,'blockArc',2.97,3.455,1.194,1.194,{fill:C.sage,line:{c:C.white,w:2.25},rot:327.49,arc:[242.832,271.453],thick:0.249});
  box(s,'blockArc',2.977,3.454,1.194,1.194,{fill:C.sage,line:{c:C.white,w:2.25},rot:327.49,arc:[209.132,242.373],thick:0.252});
  txt(s,'2 nd Qtr',4.406,3.782,0.558,0.202,T.o8Gray2);
  txt(s,'4 th Qtr',4.405,4.191,0.53,0.202,T.o8Gray2);
  txt(s,'1 st Qtr',4.41,3.6,0.516,0.202,T.o8Gray2);
  txt(s,'3 rd Qtr',4.406,3.976,0.537,0.202,T.o8Gray2);
  ell(s,0.567,3.213,1.751,1.751,C.navy);
  box(s,'pie',0.567,3.204,1.751,1.751,{fill:C.mist,arc:[235.793,270]});
  box(s,'pie',0.56,3.22,1.751,1.751,{fill:C.sage,arc:[190.758,237.012]});
  box(s,'pie',0.567,3.221,1.751,1.751,{fill:C.slate,arc:[113.741,190.895]});
}

// ---------------------------------------------------------------- slide 29 — USA Map
function slide29(s) {
  rect(s,-0.032,0,3.818,5.625,C.navy);
  usaMap(s);
  txt(s,'USA Map',5.7,0.462,2.386,0.429,T.d21bCoal,{al:'center'});
  ell(s,5.283,4.971,0.102,0.102,C.navy);
  txt(s,'Work Location',5.452,4.915,1.159,0.227,T.d9Black);
  ell(s,7.059,4.971,0.102,0.102,C.sage);
  txt(s,'StudyLocation',7.228,4.915,1.275,0.227,T.d9Black);
  txt(s,'Editable Map',0.514,1.433,2.717,0.48,T.d24bWhite);
  txt(s,'Lorem ipsum dolor sit amet, consectetur adipisc elit sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim.',0.514,2.183,2.724,0.678,T.o8White,{ls:1.5});
  txt(s,'Find Now',0.514,3.874,1.043,0.318,T.o8bNavy,{al:'center',va:'middle',fill:C.white,shape:'roundRect',rad:0.159});
  txt(s,'Lorem ipsum dolor sit amet, consectetur adipisc elit sed do eiusmod tempor incididunt.',0.514,2.927,2.724,0.47,T.o8White,{ls:1.5});
  dots(s,9.417,5.058,4,4,0.111,0.042,C.navy,1);
  dots(s,9.417,0.195,4,4,0.111,0.042,C.navy,2);
  ell(s,0.208,5.058,0.042,0.042,C.white);
  ell(s,0.319,5.17,0.042,0.042,C.white);
  ell(s,0.208,5.17,0.042,0.042,C.white);
  ell(s,0.43,5.281,0.042,0.042,C.white);
  ell(s,0.319,5.281,0.042,0.042,C.white);
  ell(s,0.208,5.281,0.042,0.042,C.white);
  ell(s,0.541,5.393,0.042,0.042,C.white);
  ell(s,0.43,5.393,0.042,0.042,C.white);
  ell(s,0.319,5.393,0.042,0.042,C.white);
  ell(s,0.208,5.393,0.042,0.042,C.white);
  ell(s,0.541,0.195,0.042,0.042,C.white);
  ell(s,0.43,0.195,0.042,0.042,C.white);
  ell(s,0.319,0.195,0.042,0.042,C.white);
  ell(s,0.208,0.195,0.042,0.042,C.white);
  ell(s,0.43,0.306,0.042,0.042,C.white);
  ell(s,0.319,0.306,0.042,0.042,C.white);
  ell(s,0.208,0.306,0.042,0.042,C.white);
  ell(s,0.319,0.418,0.042,0.042,C.white);
  ell(s,0.208,0.418,0.042,0.042,C.white);
  ell(s,0.208,0.529,0.042,0.042,C.white);
}

// ---------------------------------------------------------------- slide 30 — Thank You
function slide30(s) {
  rect(s,0,0,10,5.625,C.gray);
  rect(s,0,0,10,5.625,[C.white,14.118]);
  txt(s,'Thank You',2.727,2.011,4.547,0.985,T.d54bBlack,{al:'center'});
  box(s,'roundRect',3.227,3.077,3.546,0.536,{fill:C.slate,flipH:1,rad:0.154});
  txt(s,'Contact Me, Check Portfolio',3.398,3.194,3.204,0.303,T.o14White,{al:'center'});
  dots(s,9.495,5.138,3,3,0.158,0.06,C.navy);
  dots(s,9.495,0.099,3,3,0.158,0.06,C.navy);
  path(s,0,0.687,1.711,4.33,[M(1.711,0),[1.669,0.01],B(0.709,0.257,0,1.128,0,2.165),B(0,3.202,0.709,4.073,1.669,4.32),[1.711,4.33],[1.711,3.804],[1.585,3.761],B(0.953,3.507,0.507,2.888,0.507,2.165),B(0.507,1.442,0.953,0.823,1.585,0.569),[1.711,0.525],Z],{fill:[C.mist,75.295],flipH:1});
  dots(s,0.129,5.138,3,3,0.158,0.06,C.navy);
  dots(s,0.129,0.099,3,3,0.158,0.06,C.navy);
  path(s,8.289,0.648,1.711,4.33,[M(1.711,0),[1.669,0.01],B(0.709,0.257,0,1.128,0,2.165),B(0,3.202,0.709,4.073,1.669,4.32),[1.711,4.33],[1.711,3.804],[1.585,3.761],B(0.953,3.507,0.507,2.888,0.507,2.165),B(0.507,1.442,0.953,0.823,1.585,0.569),[1.711,0.525],Z],{fill:[C.mist,75.295]});
}

const SLIDES = [
  slide01, slide02, slide03, slide04, slide05, slide06,
  slide07, slide08, slide09, slide10, slide11, slide12,
  slide13, slide14, slide15, slide16, slide17, slide18,
  slide19, slide20, slide21, slide22, slide23, slide24,
  slide25, slide26, slide27, slide28, slide29, slide30,
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'CUSTOM', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'CUSTOM';
  pptx.author = 'Julian Maxwell';
  pptx.title = 'Resume / Portfolio';
  SLIDES.forEach(fn => {
    const s = pptx.addSlide();
    s.background = { color: C.white };
    fn(s);
  });
  return pptx.writeFile({ fileName: path_.join(__dirname, '033effef-5599-4a31-b0e6-e76577639573_grok_final.pptx') });
}

build().then(f => console.log(`wrote ${f}`)).catch(e => { console.error(e); process.exit(1); });
