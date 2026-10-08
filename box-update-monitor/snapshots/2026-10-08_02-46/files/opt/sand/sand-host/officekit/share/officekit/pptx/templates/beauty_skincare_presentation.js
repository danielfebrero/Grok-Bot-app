/**
 * Beauty Presentation Template - 30 slide deck rebuilt with pptxgenjs.
 * Run: node 12e35bc5-051d-4585-ad18-ffb1c7c9d549_grok_final.js
 */
'use strict';
const pptxgen = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ theme */
const PINK   = 'FD81AD';  // accent1
const DEEP   = 'FC2370';  // accent1, 75% luminance
const SOFT   = 'FEB3CE';  // accent1, 60% lum / 40% offset
const DARK   = '262626';  // text, 85% lum
const INK    = '0D0D0D';
const GRAY   = '808080';
const SILVER = 'D9D9D9';
const SMOKE  = 'F2F2F2';
const WHITE  = 'FFFFFF';
const ORANGE = 'EF6522';

const HEAD = 'Roboto Medium';   // theme major font
const BODY = 'Open Sans';       // theme minor font

// pptxgenjs mutates the shadow object it is handed, so hand it a fresh copy.
const softShadow = () => ({ type: 'outer', blur: 30, offset: 8, angle: 45, color: '000000', opacity: 0.2 });

/* -------------------------------------------------------------- primitives */
// Every text box in the source deck is top anchored with the PowerPoint
// default insets, so that is what tx() reproduces.
function tx(s, text, o) {
  s.addText(text, Object.assign({ fontFace: BODY, valign: 'top', color: DARK }, o));
}

function shape(s, kind, o) {
  const opt = { x: o.x, y: o.y, w: o.w, h: o.h };
  opt.fill = o.fill ? { color: o.fill } : { type: 'none' };
  opt.line = o.line ? { color: o.line, width: o.lineW || 1 } : { type: 'none' };
  if (o.radius) opt.rectRadius = o.radius;
  if (o.rotate) opt.rotate = o.rotate;
  if (o.flipH) opt.flipH = true;
  if (o.flipV) opt.flipV = true;
  if (o.shadow) opt.shadow = softShadow();
  if (o.wedge) { kind = 'pie'; opt.angleRange = o.wedge; }
  s.addShape(kind, opt);
}
const rect       = (s, o) => shape(s, 'rect', o);
const roundRect  = (s, o) => shape(s, 'roundRect', o);
const round1Rect = (s, o) => shape(s, 'round1Rect', o);
const ellipse    = (s, o) => shape(s, 'ellipse', o);

function line(s, o) {
  const ln = { color: o.color, width: o.width };
  if (o.head) ln.beginArrowType = o.head;
  if (o.tail) ln.endArrowType = o.tail;
  s.addShape('line', { x: o.x, y: o.y, w: o.w, h: o.h, line: ln });
}

// Filled polygon inside the box [x,y,w,h]; pts are 0..1 fractions of the box.
function poly(s, pts, o) {
  s.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: { color: o.color }, line: { type: 'none' },
    points: pts.map(function (p, i) {
      return { x: +(p[0] * o.w).toFixed(4), y: +(p[1] * o.h).toFixed(4), moveTo: i === 0 };
    }).concat([{ close: true }])
  });
}

// Diagonal two stop gradient, painted as a stack of rotated bands because
// pptxgenjs has no gradient fill.
function gradient(s, from, to) {
  const bands = 60, span = 14.8, len = 20;
  for (let i = 0; i < bands; i++) {
    const f = (i + 0.5) / bands;
    const c = f * span;
    s.addShape('rect', {
      x: c * 0.7071 - len / 2, y: c * 0.7071 - (span / bands) * 0.9,
      w: len, h: (span / bands) * 1.8, rotate: -45,
      fill: { color: blend(from, to, f) }, line: { type: 'none' }
    });
  }
}

function blend(c1, c2, f) {
  const ch = (c, i) => parseInt(c.substr(i * 2, 2), 16);
  return [0, 1, 2].map(function (i) {
    return Math.round(ch(c1, i) + (ch(c2, i) - ch(c1, i)) * f).toString(16).padStart(2, '0');
  }).join('').toUpperCase();
}

/* ------------------------------------------------------------- deck chrome */
const NAV = [['Home', 0.354], ['About Us', 1.396], ['Team', 2.438], ['Services', 3.48]];

function chrome(s, o) {
  if (o.nav) {
    NAV.forEach(function (item) {
      tx(s, item[0], { x: item[1], y: 0.429, w: 1.081, h: 0.286,
                       fontSize: 11, color: o.nav, align: 'center' });
    });
  }
  if (o.brand) {
    tx(s, 'Beauty Presentation', { x: 10.322, y: 0.429, w: 1.754, h: 0.303,
                                   fontSize: 12, fontFace: HEAD, bold: true,
                                   color: o.brand, align: 'right' });
  }
  if (o.logo) logo(s, 12.217, 0.442, 0.278, o.logo);
}

// The "Graphic 214" mark: three concentric rings around a solid centre dot.
function logo(s, x, y, d, color) {
  ellipse(s, { x: x, y: y, w: d, h: d, line: color, lineW: 2.2 });
  ellipse(s, { x: x + d * 0.24, y: y + d * 0.24, w: d * 0.52, h: d * 0.52, line: color, lineW: 1.6 });
  ellipse(s, { x: x + d * 0.41, y: y + d * 0.41, w: d * 0.18, h: d * 0.18, fill: color });
}

/* -------------------------------------------------------------- icon table */
// The reference deck uses small monochrome PNG glyphs. Each is redrawn here
// from primitives laid out in a 0..1 unit square:
//   ['e', x, y, w, h]      filled ellipse     ['O', x, y, w, h]   ring outline
//   ['r', x, y, w, h, rot] filled rectangle   ['R', x, y, w, h]   rounded bar
//   ['B', x, y, w, h]      rectangle outline  ['S', kind, x, y, w, h] preset shape
//   ['W', cx, cy, r, a, b] pie wedge          ['p', pts]          filled polygon
//   ['h', ...]             same as the row above it but painted in the hole colour
const ICONS = {
  check:  [['p', [[0.02, 0.46], [0.16, 0.31], [0.38, 0.56], [0.84, 0.07], [0.99, 0.21], [0.38, 0.88]]]],
  // radiation trefoil: three blades around a small hub
  burst:  [['W', 0.50, 0.50, 0.50, 180, 240], ['W', 0.50, 0.50, 0.50, 300, 360],
           ['W', 0.50, 0.50, 0.50, 60, 120],
           ['e', 0.34, 0.34, 0.32, 0.32, 'hole'], ['e', 0.39, 0.39, 0.22, 0.22]],
  snow:   [['r', 0.45, 0.00, 0.10, 1.00], ['r', 0.45, 0.00, 0.10, 1.00, 60],
           ['r', 0.45, 0.00, 0.10, 1.00, 120], ['S', 'star6', 0.16, 0.16, 0.68, 0.68]],
  gear:   [['S', 'star8', 0.00, 0.00, 1.00, 1.00],
           ['e', 0.28, 0.28, 0.44, 0.44, 'hole'], ['O', 0.32, 0.32, 0.36, 0.36]],
  frame:  [['B', 0.16, 0.16, 0.68, 0.68], ['r', 0.00, 0.00, 0.18, 0.18], ['r', 0.82, 0.00, 0.18, 0.18],
           ['r', 0.00, 0.82, 0.18, 0.18], ['r', 0.82, 0.82, 0.18, 0.18]],
  note:   [['B', 0.02, 0.02, 0.96, 0.96], ['p', [[0.70, 0.70], [0.98, 0.70], [0.70, 0.98]]]],
  bars:   [['R', 0.34, 0.10, 0.56, 0.13], ['R', 0.34, 0.40, 0.56, 0.13], ['R', 0.06, 0.70, 0.84, 0.13]],
  book:   [['e', 0.36, 0.00, 0.28, 0.28],
           ['p', [[0.00, 0.42], [0.44, 0.52], [0.44, 1.00], [0.00, 0.90]]],
           ['p', [[1.00, 0.42], [0.56, 0.52], [0.56, 1.00], [1.00, 0.90]]]],
  tag:    [['e', 0.58, 0.00, 0.26, 0.26], ['R', 0.02, 0.28, 0.58, 0.72], ['R', 0.48, 0.02, 0.50, 0.42]],
  ring:   [['O', 0.00, 0.22, 1.00, 0.56], ['O', 0.20, 0.42, 0.60, 0.24]],
  pin:    [['e', 0.02, 0.28, 0.40, 0.40], ['p', [[0.22, 1.00], [0.04, 0.56], [0.40, 0.56]]],
           ['e', 0.62, 0.00, 0.34, 0.34], ['R', 0.42, 0.44, 0.54, 0.14]],
  phone:  [['p', [[0.10, 0.00], [0.38, 0.00], [0.46, 0.30], [0.28, 0.44],
                  [0.56, 0.72], [0.70, 0.54], [1.00, 0.62], [1.00, 0.90],
                  [0.80, 1.00], [0.34, 0.82], [0.06, 0.36]]]],
  folder: [['p', [[0.02, 0.16], [0.40, 0.16], [0.52, 0.34], [0.98, 0.34], [0.98, 0.92], [0.02, 0.92]]],
           ['h', [[0.14, 0.28], [0.34, 0.28], [0.46, 0.46], [0.86, 0.46], [0.86, 0.80], [0.14, 0.80]]]],
  map:    [['p', [[0.00, 0.14], [0.34, 0.00], [0.66, 0.14], [1.00, 0.00],
                  [1.00, 0.86], [0.66, 1.00], [0.34, 0.86], [0.00, 1.00]]],
           ['h', [[0.12, 0.26], [0.26, 0.20], [0.26, 0.74], [0.12, 0.80]]],
           ['h', [[0.40, 0.20], [0.60, 0.28], [0.60, 0.82], [0.40, 0.74]]],
           ['h', [[0.74, 0.20], [0.88, 0.14], [0.88, 0.68], [0.74, 0.74]]]],
  bell:   [['p', [[0.50, 0.00], [0.86, 0.26], [0.86, 0.70], [1.00, 0.84], [0.00, 0.84],
                  [0.14, 0.70], [0.14, 0.26]]],
           ['h', [[0.50, 0.16], [0.72, 0.34], [0.72, 0.74], [0.28, 0.74], [0.28, 0.34]]],
           ['e', 0.36, 0.84, 0.28, 0.16]]
};

function icon(s, name, o) {
  const color = o.color || WHITE;
  const hole = o.hole || WHITE;
  ICONS[name].forEach(function (it) {
    const k = it[0];
    if (k === 'p' || k === 'h') {
      poly(s, it[1], { x: o.x, y: o.y, w: o.w, h: o.h, color: k === 'h' ? hole : color });
      return;
    }
    if (k === 'W') {
      const r = it[3];
      ellipse(s, { x: o.x + (it[1] - r) * o.w, y: o.y + (it[2] - r) * o.h,
                   w: 2 * r * o.w, h: 2 * r * o.h, fill: color, wedge: [it[4], it[5]] });
      return;
    }
    const off = k === 'S' ? 1 : 0;
    const bx = { x: o.x + it[1 + off] * o.w, y: o.y + it[2 + off] * o.h,
                 w: it[3 + off] * o.w, h: it[4 + off] * o.h };
    if (k === 'S') shape(s, it[1], Object.assign(bx, { fill: color }));
    else if (k === 'e') ellipse(s, Object.assign(bx, { fill: it[5] === 'hole' ? hole : color }));
    else if (k === 'O') ellipse(s, Object.assign(bx, { line: color, lineW: 1.5 }));
    else if (k === 'B') shape(s, 'rect', Object.assign(bx, { line: color, lineW: 1.5 }));
    else if (k === 'R') roundRect(s, Object.assign(bx, { fill: color, radius: Math.min(bx.w, bx.h) / 2 }));
    else if (k === 'r') rect(s, Object.assign(bx, { fill: color, rotate: it[5] || 0 }));
  });
}

/* ---------------------------------------------------------- composite bits */
// Pink disc with a white chevron pointing right (Freeform 31 in the source).
function arrowBadge(s, o) {
  const d = o.size;
  ellipse(s, { x: o.x, y: o.y, w: d, h: d, fill: o.fill || PINK });
  chevron(s, { x: o.x + d * 0.30, y: o.y + d * 0.26, w: d * 0.34, h: d * 0.48, color: o.arrow || WHITE });
}

// Bare right pointing chevron (Freeform 32 in the source).
function chevron(s, o) {
  poly(s, [[0.00, 0.12], [0.18, 0.00], [1.00, 0.50], [0.18, 1.00], [0.00, 0.88], [0.62, 0.50]],
       { x: o.x, y: o.y, w: o.w, h: o.h, color: o.color });
}

// Pink disc with a white tick (Shape 12417 / 12420 in the source).
function checkBadge(s, o) {
  ellipse(s, { x: o.x, y: o.y, w: o.size, h: o.size, fill: PINK });
  icon(s, 'check', { x: o.x + o.size * 0.24, y: o.y + o.size * 0.28,
                     w: o.size * 0.52, h: o.size * 0.44, color: WHITE });
}

// Three dots used as a "more" affordance.
function dots(s, o) {
  for (let i = 0; i < 3; i++) {
    ellipse(s, { x: o.x + i * (o.w - o.d) / 2, y: o.y, w: o.d, h: o.d, fill: o.color });
  }
}

// Small bar chart inside a ring (the "Statistic" glyph).
function barChart(s, o) {
  const bw = o.w * 0.16;
  [0.38, 0.58, 0.80, 1.00].forEach(function (f, i) {
    rect(s, { x: o.x + i * o.w * 0.26, y: o.y + o.h * (1 - f), w: bw, h: o.h * f - o.h * 0.10, fill: o.color });
  });
  rect(s, { x: o.x, y: o.y + o.h * 0.92, w: o.w, h: o.h * 0.08, fill: o.color });
}

// Hourglass glyph.
function hourglass(s, o) {
  poly(s, [[0.00, 0.00], [1.00, 0.00], [0.50, 0.50]], { x: o.x, y: o.y, w: o.w, h: o.h, color: o.color });
  poly(s, [[0.50, 0.50], [1.00, 1.00], [0.00, 1.00]], { x: o.x, y: o.y, w: o.w, h: o.h, color: o.color });
}

// Stand-in for the photographic phone mock-ups on slide 29.
function photo(s, o) {
  roundRect(s, { x: o.x, y: o.y, w: o.w, h: o.h, fill: INK, radius: o.w * 0.12 });
  roundRect(s, { x: o.x + 0.05, y: o.y + 0.05, w: o.w - 0.1, h: o.h - 0.1, fill: SOFT, radius: o.w * 0.10 });
  tx(s, '[image]', { x: o.x, y: o.y + o.h / 2 - 0.15, w: o.w, h: 0.3,
                     fontSize: 11, color: WHITE, align: 'center' });
}

/* ------------------------------------------------------------ slide bodies */

function slide01(s) {
  gradient(s, PINK, DEEP);
  chrome(s, { nav:null, brand:WHITE, logo:WHITE });
  tx(s, 'BEAUTY', { x:7.58, y:1.839, w:4.496, h:1.447, fontSize:80, fontFace:HEAD, bold:true, color:WHITE });
  tx(s, 'Presentation Template', { x:7.58, y:3.19, w:3.884, h:0.303, fontSize:12, fontFace:HEAD, bold:true, color:WHITE, charSpacing:6 });
  tx(s, 'Lorem ipsum dolor sit, consectetur adipisicing elit, sed do eiusmod tempor incididunt', { x:7.58, y:3.957, w:4.078, h:0.675, fontSize:12, color:WHITE, lineSpacingMultiple:1.5 });
  tx(s, 'www.yourwebsite.com', { x:8.61, y:6.376, w:3.884, h:0.303, fontSize:12, color:SILVER, align:'right', charSpacing:6 });
  roundRect(s, { x:7.684, y:4.839, w:2.271, h:0.418, fill:WHITE, radius:0.033 });
  tx(s, 'Lets Start Presentation', { x:7.717, y:4.905, w:2.206, h:0.286, fontSize:11, bold:true, color:PINK, align:'center' });
  line(s, { x:7.684, y:3.753, w:3.689, h:0, color:WHITE, width:1.5, head:'oval', tail:'oval' });
}

function slide02(s) {
  rect(s, { x:7.248, y:0, w:2.546, h:7.5, fill:PINK });
  chrome(s, { nav:GRAY, brand:DARK, logo:PINK });
  tx(s, 'Secrets To Radiant Skin', { x:1.005, y:1.515, w:5.662, h:0.707, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna', { x:1.005, y:2.269, w:5.317, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Skincare Routine', { x:1.005, y:3.058, w:2.099, h:0.337, fontSize:14, bold:true, color:DARK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna', { x:1.005, y:3.427, w:5.317, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  roundRect(s, { x:1.109, y:5.498, w:1.197, h:0.418, fill:PINK, radius:0.033 });
  tx(s, 'Next Page', { x:1.174, y:5.564, w:1.067, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
  roundRect(s, { x:1.109, y:4.368, w:0.172, h:0.172, fill:PINK, flipH:true });
  tx(s, 'Exercise Regularly', { x:1.422, y:4.299, w:2.051, h:0.337, fontSize:14, bold:true, color:DARK });
  roundRect(s, { x:3.745, y:4.368, w:0.172, h:0.172, fill:PINK, flipH:true });
  tx(s, 'Gentle Cleansing', { x:4.058, y:4.299, w:2.015, h:0.337, fontSize:14, bold:true, color:DARK });
  roundRect(s, { x:1.109, y:4.906, w:0.172, h:0.172, fill:PINK, flipH:true });
  tx(s, 'Processed Foods', { x:1.422, y:4.837, w:2.051, h:0.337, fontSize:14, bold:true, color:DARK });
  roundRect(s, { x:3.745, y:4.906, w:0.172, h:0.172, fill:PINK, flipH:true });
  tx(s, 'Serums Treatments', { x:4.058, y:4.837, w:2.264, h:0.337, fontSize:14, bold:true, color:DARK });
}

function slide03(s) {
  rect(s, { x:0, y:0, w:5.071, h:5.213, fill:PINK });
  chrome(s, { nav:WHITE, brand:DARK, logo:PINK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit sed do eiusmod tempor incididunt ut labore dolore magna', { x:7.318, y:2.912, w:5.071, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  roundRect(s, { x:7.455, y:3.916, w:0.151, h:0.151, fill:PINK });
  tx(s, 'Cleansing', { x:7.802, y:3.81, w:1.219, h:0.337, fontSize:14, bold:true, color:DARK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore', { x:7.802, y:4.146, w:4.587, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  roundRect(s, { x:7.455, y:5.062, w:0.151, h:0.151, fill:PINK });
  tx(s, 'Moisturizer', { x:7.802, y:4.956, w:2.603, h:0.337, fontSize:14, bold:true, color:DARK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore', { x:7.802, y:5.293, w:4.587, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Skincare Guide Without Makeup', { x:7.484, y:1.699, w:4.238, h:1.313, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
}

function slide04(s) {
  rect(s, { x:7.569, y:0, w:2.667, h:4.556, fill:PINK });
  rect(s, { x:10.236, y:2.944, w:2.572, h:4.556, fill:PINK });
  chrome(s, { nav:GRAY, brand:DARK, logo:PINK });
  tx(s, 'Dramatic to Minimalistic', { x:0.959, y:1.629, w:5.883, h:0.707, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing eli, sed eiusmod tempor incididunt ut labore et dolore magna adipisicing', { x:0.959, y:2.86, w:5.549, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing', { x:0.965, y:4.633, w:2.458, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, '2873+', { x:0.965, y:3.745, w:1.516, h:0.505, fontSize:24, fontFace:HEAD, bold:true, color:DARK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing', { x:3.737, y:4.633, w:2.458, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, '9882+', { x:3.737, y:3.745, w:1.516, h:0.505, fontSize:24, fontFace:HEAD, bold:true, color:DARK });
  roundRect(s, { x:1.084, y:5.514, w:1.197, h:0.418, fill:PINK, radius:0.033 });
  tx(s, 'See more', { x:1.149, y:5.58, w:1.067, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
  tx(s, 'Limit Sugar And Processed Foods', { x:0.959, y:2.492, w:3.66, h:0.337, fontSize:14, bold:true, color:DARK });
  tx(s, 'Insert Title Here', { x:0.959, y:4.291, w:2.031, h:0.337, fontSize:14, bold:true, color:DARK });
  tx(s, 'Insert Title Here', { x:3.733, y:4.264, w:2.031, h:0.337, fontSize:14, bold:true, color:DARK });
}

function slide05(s) {
  chrome(s, { nav:null, brand:DARK, logo:PINK });
  tx(s, 'Everyday Makeup Tutorial For You', { x:7.318, y:1.429, w:4.071, h:1.313, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  ellipse(s, { x:7.52, y:3.136, w:0.152, h:0.152, fill:PINK });
  ellipse(s, { x:7.468, y:3.084, w:0.256, h:0.256, line:PINK, lineW:1.5 });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt', { x:7.91, y:2.916, w:4.325, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  ellipse(s, { x:7.512, y:4.077, w:0.152, h:0.152, fill:PINK });
  ellipse(s, { x:7.46, y:4.024, w:0.256, h:0.256, line:PINK, lineW:1.5 });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt', { x:7.903, y:3.856, w:4.325, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  ellipse(s, { x:7.505, y:5.017, w:0.152, h:0.152, fill:PINK });
  ellipse(s, { x:7.452, y:4.965, w:0.256, h:0.256, line:PINK, lineW:1.5 });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt', { x:7.895, y:4.797, w:4.325, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  roundRect(s, { x:7.468, y:5.798, w:1.197, h:0.418, fill:PINK, radius:0.038 });
  tx(s, 'Next Page', { x:7.533, y:5.864, w:1.067, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
}

function slide06(s) {
  chrome(s, { nav:GRAY, brand:null, logo:null });
  tx(s, 'Tips For Choosing The Right Skincare', { x:1.213, y:1.386, w:4.74, h:1.313, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing eli, sed eiusmod tempor incididunt ut labore et dolore magna', { x:1.213, y:3.265, w:5.105, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Know Your Skin Type', { x:1.213, y:2.881, w:2.436, h:0.337, fontSize:14, bold:true, color:DARK });
  roundRect(s, { x:9.922, y:4.196, w:2.619, h:2.667, fill:PINK, radius:0.145 });
  line(s, { x:1.296, y:4.288, w:0.782, h:0, color:PINK, width:3, tail:'triangle' });
  tx(s, 'Read Ingredients Lists', { x:2.24, y:4.133, w:2.649, h:0.337, fontSize:14, bold:true, color:DARK });
  line(s, { x:1.296, y:4.818, w:0.782, h:0, color:PINK, width:3, tail:'triangle' });
  tx(s, 'Patch Test New Products', { x:2.24, y:4.663, w:3.022, h:0.337, fontSize:14, bold:true, color:DARK });
  line(s, { x:1.296, y:5.348, w:0.782, h:0, color:PINK, width:3, tail:'triangle' });
  tx(s, 'Check For Key Ingredients', { x:2.24, y:5.193, w:2.843, h:0.337, fontSize:14, bold:true, color:DARK });
  icon(s, 'burst', { x:10.99, y:4.512, w:0.484, h:0.5 });
  tx(s, 'Start Simple', { x:10.346, y:5.205, w:1.771, h:0.337, fontSize:14, bold:true, color:WHITE, align:'center' });
  tx(s, 'Lorem ipsum dolor amet, consectetur adipisicing eli, sed eiusmod tempor', { x:10.054, y:5.6, w:2.356, h:0.978, fontSize:12, color:WHITE, align:'center', lineSpacingMultiple:1.5 });
  roundRect(s, { x:1.296, y:5.851, w:1.197, h:0.418, fill:PINK, radius:0.035 });
  tx(s, 'Learn More', { x:1.327, y:5.917, w:1.135, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
}

function slide07(s) {
  rect(s, { x:0, y:0, w:5.159, h:4.983, fill:PINK });
  chrome(s, { nav:WHITE, brand:DARK, logo:PINK });
  tx(s, 'Tackling Dry And Split Ends With Ease', { x:7.153, y:1.405, w:4.841, h:1.313, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing', { x:7.153, y:3.308, w:2.416, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Trim Regularly', { x:7.153, y:2.937, w:1.878, h:0.337, fontSize:14, bold:true, color:DARK });
  roundRect(s, { x:7.257, y:4.135, w:1.197, h:0.418, fill:PINK, radius:0.035 });
  tx(s, 'Learn More', { x:7.287, y:4.201, w:1.135, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing', { x:9.847, y:3.308, w:2.416, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Limit Heat Styling', { x:9.847, y:2.937, w:2.147, h:0.337, fontSize:14, bold:true, color:DARK });
  roundRect(s, { x:9.95, y:4.135, w:1.197, h:0.418, fill:PINK, radius:0.035 });
  tx(s, 'Learn More', { x:9.981, y:4.201, w:1.135, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing', { x:7.153, y:5.183, w:2.416, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Natural Oils', { x:7.153, y:4.813, w:1.878, h:0.337, fontSize:14, bold:true, color:DARK });
  roundRect(s, { x:7.257, y:6.011, w:1.197, h:0.418, fill:PINK, radius:0.035 });
  tx(s, 'Learn More', { x:7.287, y:6.077, w:1.135, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing', { x:9.847, y:5.183, w:2.416, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Use Cold Water', { x:9.847, y:4.813, w:1.878, h:0.337, fontSize:14, bold:true, color:DARK });
  roundRect(s, { x:9.95, y:6.011, w:1.197, h:0.418, fill:PINK, radius:0.035 });
  tx(s, 'Learn More', { x:9.981, y:6.077, w:1.135, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
}

function slide08(s) {
  rect(s, { x:7.073, y:0, w:3.117, h:4.812, fill:PINK });
  chrome(s, { nav:GRAY, brand:DARK, logo:PINK });
  tx(s, 'Overcoming Acne And Its Marks Quickly', { x:1.089, y:1.291, w:5.026, h:1.313, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna', { x:1.089, y:2.688, w:5.317, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  icon(s, 'snow', { x:3.9, y:3.744, w:0.292, h:0.333, color:'000000' });
  icon(s, 'gear', { x:1.191, y:3.744, w:0.333, h:0.333, color:'000000' });
  tx(s, 'Natural Oils', { x:1.089, y:4.285, w:1.649, h:0.337, fontSize:14, bold:true, color:DARK });
  tx(s, 'Salicylic Acid', { x:3.815, y:4.285, w:1.649, h:0.337, fontSize:14, bold:true, color:DARK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing', { x:1.089, y:4.649, w:2.464, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing', { x:3.815, y:4.649, w:2.464, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  roundRect(s, { x:1.191, y:5.619, w:1.197, h:0.418, fill:PINK, radius:0.035 });
  tx(s, 'Read More', { x:1.222, y:5.685, w:1.135, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
}

function slide09(s) {
  rect(s, { x:0, y:0, w:5.111, h:2.756, fill:PINK });
  chrome(s, { nav:WHITE, brand:DARK, logo:PINK });
  tx(s, 'DIY Hair Care With Natural Ingredients', { x:7.322, y:1.297, w:4.694, h:1.313, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing eli, sed eiusmod tempor incididunt ut labore et dolore magna', { x:7.322, y:3.251, w:5.234, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Aloe Vera Scalp Treatment', { x:7.322, y:2.866, w:2.942, h:0.337, fontSize:14, bold:true, color:DARK });
  roundRect(s, { x:7.417, y:4.108, w:1.197, h:0.418, fill:PINK, radius:0.035 });
  tx(s, 'Learn More', { x:7.448, y:4.174, w:1.135, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing eli, sed eiusmod tempor incididunt ut labore et dolore magna', { x:7.322, y:5.116, w:5.234, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Apple Cider Vinegar Rinse', { x:7.322, y:4.732, w:2.942, h:0.337, fontSize:14, bold:true, color:DARK });
  roundRect(s, { x:7.417, y:5.974, w:1.197, h:0.418, fill:PINK, radius:0.035 });
  tx(s, 'Read More', { x:7.448, y:6.04, w:1.135, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
  roundRect(s, { x:1.317, y:1.178, w:2.413, h:2.756, fill:WHITE, radius:0.164, shadow:true });
  roundRect(s, { x:3.889, y:4.074, w:2.413, h:2.756, fill:PINK, radius:0.164 });
  icon(s, 'frame', { x:4.916, y:4.465, w:0.359, h:0.319, color:WHITE });
  icon(s, 'note', { x:2.384, y:1.551, w:0.279, h:0.319, color:PINK });
  tx(s, 'Oil Hair Mask', { x:1.699, y:2.103, w:1.649, h:0.337, fontSize:14, bold:true, color:DARK, align:'center' });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit', { x:1.527, y:2.461, w:1.994, h:0.978, fontSize:12, color:GRAY, align:'center', lineSpacingMultiple:1.5 });
  tx(s, 'Honey Hair Mask', { x:4.098, y:5.007, w:1.994, h:0.337, fontSize:14, bold:true, color:WHITE, align:'center' });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit', { x:4.098, y:5.366, w:1.994, h:0.978, fontSize:12, color:WHITE, align:'center', lineSpacingMultiple:1.5 });
}

function slide10(s) {
  chrome(s, { nav:GRAY, brand:DARK, logo:PINK });
  roundRect(s, { x:6.802, y:1.125, w:5.443, h:2.438, fill:PINK, radius:0.168 });
  tx(s, 'Overcoming Acne And Its Marks Quickly', { x:1.089, y:1.125, w:5.026, h:1.313, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna', { x:1.089, y:2.967, w:5.317, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Preventing Acne', { x:1.089, y:2.593, w:1.992, h:0.337, fontSize:14, bold:true, color:DARK });
  line(s, { x:7.306, y:1.592, w:0.782, h:0, color:WHITE, width:3, tail:'triangle' });
  tx(s, 'Non-Comedogenic Products', { x:8.251, y:1.438, w:3.051, h:0.337, fontSize:14, bold:true, color:WHITE });
  line(s, { x:7.306, y:2.084, w:0.782, h:0, color:WHITE, width:3, tail:'triangle' });
  tx(s, 'Regular Cleansing And Exfoliation', { x:8.251, y:1.929, w:3.51, h:0.337, fontSize:14, bold:true, color:WHITE });
  line(s, { x:7.306, y:2.576, w:0.782, h:0, color:WHITE, width:3, tail:'triangle' });
  tx(s, 'Consult A Dermatologist', { x:8.251, y:2.421, w:2.649, h:0.337, fontSize:14, bold:true, color:WHITE });
  line(s, { x:7.306, y:3.067, w:0.782, h:0, color:WHITE, width:3, tail:'triangle' });
  tx(s, 'Stress Management', { x:8.251, y:2.913, w:2.649, h:0.337, fontSize:14, bold:true, color:WHITE });
}

function slide11(s) {
  round1Rect(s, { x:0, y:4.619, w:6.667, h:2.881, fill:PINK, radius:0.244 });
  chrome(s, { nav:GRAY, brand:DARK, logo:PINK });
  roundRect(s, { x:1.333, y:1.331, w:2.413, h:1.891, fill:PINK, radius:0.129 });
  tx(s, 'Benefits Of Essential Oils For Beauty', { x:7.306, y:1.331, w:5.012, h:1.313, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et', { x:8.083, y:2.84, w:4.552, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, '01.', { x:7.306, y:2.884, w:0.778, h:0.505, fontSize:24, fontFace:HEAD, bold:true, color:INK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et', { x:8.083, y:3.693, w:4.552, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, '02.', { x:7.306, y:3.737, w:0.778, h:0.505, fontSize:24, fontFace:HEAD, bold:true, color:INK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et', { x:8.083, y:4.546, w:4.552, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, '03.', { x:7.306, y:4.59, w:0.778, h:0.505, fontSize:24, fontFace:HEAD, bold:true, color:INK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna', { x:7.306, y:5.349, w:5.317, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, ', { x:1.544, y:1.95, w:1.991, h:0.978, fontSize:12, color:WHITE, align:'center', lineSpacingMultiple:1.5 });
  tx(s, 'Moisturization', { x:1.544, y:1.623, w:1.992, h:0.337, fontSize:14, bold:true, color:WHITE, align:'center' });
}

function slide12(s) {
  round1Rect(s, { x:0.476, y:4.286, w:12.857, h:3.214, fill:PINK, flipH:true, radius:0.272 });
  chrome(s, { nav:GRAY, brand:null, logo:null });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing', { x:1.267, y:5.737, w:2.942, h:0.675, fontSize:12, color:WHITE, align:'center', lineSpacingMultiple:1.5 });
  tx(s, 'Triangular Face Shape', { x:1.383, y:5.353, w:2.711, h:0.337, fontSize:14, bold:true, color:WHITE, align:'center' });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing', { x:4.91, y:5.737, w:2.942, h:0.675, fontSize:12, color:WHITE, align:'center', lineSpacingMultiple:1.5 });
  tx(s, 'Heart Face Shape', { x:5.241, y:5.353, w:2.28, h:0.337, fontSize:14, bold:true, color:WHITE, align:'center' });
  tx(s, 'Choosing Sunglasses That Suit Your Face Shape', { x:1.048, y:1.082, w:6.397, h:1.313, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
}

function slide13(s) {
  rect(s, { x:0, y:3.381, w:13.333, h:4.119, fill:PINK });
  chrome(s, { nav:GRAY, brand:DARK, logo:PINK });
  tx(s, 'Practical Hand And Nail Care With Cuticle Oil', { x:8.417, y:1.31, w:4.075, h:1.919, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing eli, sed eiusmod tempor incididunt ut labore', { x:8.417, y:4.134, w:4.444, h:0.675, fontSize:12, color:WHITE, lineSpacingMultiple:1.5 });
  tx(s, 'Gather Your Supplies', { x:8.417, y:3.845, w:2.408, h:0.337, fontSize:14, bold:true, color:WHITE });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing eli, sed eiusmod tempor incididunt ut labore', { x:8.417, y:5.329, w:4.444, h:0.675, fontSize:12, color:WHITE, lineSpacingMultiple:1.5 });
  tx(s, 'Clean Your Hands and Nails', { x:8.417, y:4.945, w:2.942, h:0.337, fontSize:14, bold:true, color:WHITE });
}

function slide14(s) {
  rect(s, { x:6.794, y:4.238, w:6.54, h:3.262, fill:PINK });
  chrome(s, { nav:GRAY, brand:DARK, logo:PINK });
  tx(s, 'Latest Tips For Achieving Thick', { x:1.159, y:1.331, w:4.27, h:1.313, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  tx(s, 'Lorem ipsum dolor si consectetur adipisicing elit, sed do', { x:9.581, y:1.804, w:2.824, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, '145K', { x:9.581, y:1.363, w:1.108, h:0.505, fontSize:24, fontFace:HEAD, bold:true, color:DARK });
  icon(s, 'check', { x:1.281, y:4.147, w:0.191, h:0.191, color:PINK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut', { x:1.159, y:3.167, w:4.429, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Nutrient-Rich Diet', { x:1.159, y:2.792, w:1.992, h:0.337, fontSize:14, bold:true, color:DARK });
  tx(s, 'Avoid Tight Hairstyles', { x:1.704, y:4.07, w:2.532, h:0.337, fontSize:14, bold:true, color:DARK });
  icon(s, 'check', { x:1.281, y:4.712, w:0.191, h:0.191, color:PINK });
  tx(s, 'Hair Supplements', { x:1.704, y:4.635, w:1.992, h:0.337, fontSize:14, bold:true, color:DARK });
  icon(s, 'check', { x:1.281, y:5.277, w:0.191, h:0.191, color:PINK });
  tx(s, 'Gentle Hair Care', { x:1.704, y:5.2, w:1.992, h:0.337, fontSize:14, bold:true, color:DARK });
  roundRect(s, { x:1.281, y:5.817, w:1.197, h:0.418, fill:PINK, radius:0.035 });
  tx(s, 'Next Page', { x:1.312, y:5.883, w:1.135, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
}

function slide15(s) {
  rect(s, { x:0, y:5.143, w:13.333, h:2.357, fill:PINK });
  chrome(s, { nav:GRAY, brand:DARK, logo:PINK });
  tx(s, 'Hair Styling Inspiration With Ribbons And Hairpins', { x:4.687, y:0.968, w:6.25, h:1.313, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  tx(s, [{ text:'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore ' }, { text:'magna consectetur adipisicing elit, sed do eiusmod tempor' }], { x:4.687, y:2.365, w:7.313, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
}

function slide16(s) {
  rect(s, { x:7.143, y:4, w:6.19, h:3.5, fill:PINK });
  chrome(s, { nav:GRAY, brand:DARK, logo:PINK });
  tx(s, 'Hollywood Starlet\'s Secrets To Gorgeous Hair', { x:1.111, y:1.331, w:4.381, h:1.919, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et', { x:1.111, y:3.766, w:4.603, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Quality Haircare Products', { x:1.111, y:3.391, w:2.762, h:0.337, fontSize:14, bold:true, color:DARK });
  roundRect(s, { x:1.249, y:4.684, w:1.197, h:0.418, fill:PINK, radius:0.035 });
  tx(s, 'Learn More', { x:1.28, y:4.75, w:1.135, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et', { x:1.111, y:5.346, w:4.603, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  roundRect(s, { x:8.043, y:4.893, w:3.756, h:1.084, fill:WHITE, radius:0.11, shadow:true });
  arrowBadge(s, { x:8.401, y:5.186, size:0.494 });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit', { x:9.086, y:5.051, w:2.622, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
}

function slide17(s) {
  rect(s, { x:0, y:0, w:5.007, h:7.5, fill:PINK });
  chrome(s, { nav:WHITE, brand:DARK, logo:PINK });
  tx(s, 'Vibrant Colors And Unique Designs', { x:7.26, y:1.145, w:4.536, h:1.313, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  roundRect(s, { x:7.322, y:3.931, w:5.007, h:1.093, fill:PINK, radius:0.141, shadow:true });
  ellipse(s, { x:11.928, y:4.13, w:0.696, h:0.696, fill:WHITE, shadow:true });
  chevron(s, { x:12.125, y:4.38, w:0.302, h:0.196, color:PINK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod', { x:8.282, y:4.096, w:3.41, h:0.675, fontSize:12, color:WHITE, lineSpacingMultiple:1.5 });
  roundRect(s, { x:7.322, y:5.372, w:5.007, h:1.093, fill:WHITE, radius:0.141, shadow:true });
  ellipse(s, { x:11.928, y:5.571, w:0.696, h:0.696, fill:PINK });
  chevron(s, { x:12.125, y:5.821, w:0.302, h:0.196, color:WHITE });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod', { x:8.282, y:5.538, w:3.41, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et', { x:7.26, y:2.993, w:4.918, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Accessorize Creatively', { x:7.26, y:2.604, w:2.762, h:0.337, fontSize:14, bold:true, color:DARK });
  icon(s, 'bars', { x:7.709, y:4.299, w:0.313, h:0.358, color:WHITE });
  icon(s, 'book', { x:7.687, y:5.74, w:0.358, h:0.358, color:PINK });
}

function slide18(s) {
  chrome(s, { nav:GRAY, brand:DARK, logo:PINK });
  tx(s, 'Combatting Dull Skin With Homemade Scrubs', { x:1.111, y:1.601, w:5.667, h:1.313, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  tx(s, [{ text:'Lorem ipsum dolor sit amet, consectetur adipisicing elit, seeiusmod tempor incididunt ut labore ' }, { text:'et adipisicing elit, sed do' }], { x:1.111, y:3.397, w:5.556, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing', { x:1.111, y:4.817, w:2.408, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, '2873+', { x:1.111, y:4.312, w:1.516, h:0.505, fontSize:24, fontFace:HEAD, bold:true, color:INK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing', { x:4.049, y:4.817, w:2.408, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, '9882+', { x:4.049, y:4.312, w:1.516, h:0.505, fontSize:24, fontFace:HEAD, bold:true, color:INK });
  tx(s, 'Honey And Sugar Scrub', { x:1.111, y:3.019, w:2.762, h:0.337, fontSize:14, bold:true, color:DARK });
  roundRect(s, { x:1.21, y:5.77, w:1.197, h:0.418, fill:PINK, radius:0.035 });
  tx(s, 'Learn More', { x:1.241, y:5.836, w:1.135, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
  roundRect(s, { x:6.479, y:5.243, w:5.183, h:1.384, fill:WHITE, radius:0.135, shadow:true });
  ellipse(s, { x:6.742, y:5.621, w:0.628, h:0.628, line:PINK, lineW:2.25 });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, seeiusmod tempor incididunt ut labore', { x:7.622, y:5.77, w:4.04, h:0.627, fontSize:11, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Statistic', { x:7.621, y:5.417, w:1.017, h:0.37, fontSize:16, fontFace:HEAD, bold:true, color:PINK });
  dots(s, { x:11.016, y:5.561, w:0.366, d:0.085, color:PINK });
  barChart(s, { x:6.904, y:5.798, w:0.305, h:0.274, color:PINK });
}

function slide19(s) {
  rect(s, { x:0, y:2.886, w:5.604, h:4.614, fill:PINK });
  chrome(s, { nav:GRAY, brand:DARK, logo:PINK });
  tx(s, 'Monochromatic Makeup Elegance', { x:7.607, y:1.574, w:4.536, h:1.313, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  line(s, { x:7.844, y:3.308, w:0, h:1.824, color:SOFT, width:1 });
  tx(s, 'Lorem ipsum dolor sit amet  consec tetur adlorem ipsum dolor sit amet  consec tetur conse', { x:8.209, y:3.091, w:4.175, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  ellipse(s, { x:7.729, y:3.241, w:0.233, h:0.233, fill:PINK });
  tx(s, 'Lorem ipsum dolor sit amet  consec tetur adlorem ipsum dolor sit amet  consec tetur consec', { x:8.209, y:3.945, w:4.175, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  ellipse(s, { x:7.729, y:4.095, w:0.233, h:0.233, fill:PINK });
  tx(s, 'Lorem ipsum dolor sit amet  consec tetur adlorem ipsum dolor sit amet  consec tetur consec', { x:8.209, y:4.798, w:4.175, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  ellipse(s, { x:7.729, y:4.948, w:0.233, h:0.233, fill:PINK });
  roundRect(s, { x:7.729, y:5.802, w:1.197, h:0.418, fill:PINK, radius:0.035 });
  tx(s, 'Learn More', { x:7.76, y:5.868, w:1.135, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
}

function slide20(s) {
  rect(s, { x:9.229, y:2.604, w:4.104, h:4.896, fill:PINK });
  chrome(s, { nav:GRAY, brand:DARK, logo:PINK });
  tx(s, 'Natural Beauty With Minimal Makeup', { x:1.098, y:1.395, w:4.892, h:1.313, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  tx(s, 'Lorem ipsum dolor sit amet  consec tetur adlorem ipsum dolor sit amet  consec tetur conse', { x:1.583, y:3.29, w:4.407, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, '3788K', { x:1.599, y:2.863, w:1.606, h:0.505, fontSize:24, fontFace:HEAD, bold:true, color:DARK });
  tx(s, 'Lorem ipsum dolor sit amet  consec tetur adlorem ipsum dolor sit amet  consec tetur conse', { x:1.583, y:5.209, w:4.407, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, '4839K', { x:1.583, y:4.782, w:1.606, h:0.505, fontSize:24, fontFace:HEAD, bold:true, color:DARK });
  ellipse(s, { x:1.205, y:3.081, w:0.087, h:0.087, fill:PINK });
  ellipse(s, { x:1.205, y:4.99, w:0.087, h:0.087, fill:PINK });
  line(s, { x:1.249, y:3.168, w:0, h:1.822, color:SOFT, width:1 });
  roundRect(s, { x:1.676, y:4.101, w:1.197, h:0.418, fill:PINK, radius:0.035 });
  tx(s, 'Learn More', { x:1.707, y:4.166, w:1.135, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
  roundRect(s, { x:1.704, y:6.02, w:1.197, h:0.418, fill:PINK, radius:0.035 });
  tx(s, 'Learn More', { x:1.735, y:6.085, w:1.135, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
}

function slide21(s) {
  rect(s, { x:0, y:3.304, w:5.794, h:4.196, fill:PINK });
  chrome(s, { nav:GRAY, brand:DARK, logo:PINK });
  line(s, { x:8.142, y:3.989, w:0, h:1.632, color:SOFT, width:1 });
  tx(s, 'Feminine Short Hairstyles That Remain Chic', { x:7.854, y:1.593, w:4.06, h:1.919, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  tx(s, 'Lorem ipsum dolor sit amet  consec tetur adlorem ipsum dolor sit ame consec', { x:8.54, y:3.682, w:3.676, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Lorem ipsum dolor sit amet  consec tetur adlorem ipsum dolor sit ame consec', { x:8.54, y:4.583, w:3.676, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Lorem ipsum dolor sit amet  consec tetur adlorem ipsum dolor sit ame consec', { x:8.54, y:5.458, w:3.676, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  ellipse(s, { x:8.029, y:3.84, w:0.226, h:0.226, fill:PINK });
  ellipse(s, { x:8.029, y:4.73, w:0.226, h:0.226, fill:DEEP });
  ellipse(s, { x:8.029, y:5.62, w:0.226, h:0.226, fill:PINK });
}

function slide22(s) {
  rect(s, { x:8.655, y:3.898, w:4.675, h:3.602, fill:PINK });
  chrome(s, { nav:GRAY, brand:DARK, logo:PINK });
  tx(s, 'The Marvels Of Mud Masks For Skin', { x:1.259, y:1.615, w:4.714, h:1.313, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  tx(s, [{ text:'Lorem ipsum dolor sit amet, consectetur adipisicing, seeiusmod tempor incididunt ut labore ' }, { text:'et adipisicing elit, sed do' }], { x:1.259, y:3.012, w:5.271, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing', { x:1.259, y:4.739, w:2.408, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, '982+', { x:1.259, y:4.234, w:1.516, h:0.505, fontSize:24, fontFace:HEAD, bold:true, color:INK });
  tx(s, 'Cleansing', { x:1.259, y:3.898, w:1.516, h:0.337, fontSize:14, bold:true });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing', { x:4.101, y:4.739, w:2.408, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, '231+', { x:4.101, y:4.234, w:1.516, h:0.505, fontSize:24, fontFace:HEAD, bold:true, color:INK });
  tx(s, 'Exfoliation', { x:4.101, y:3.898, w:1.516, h:0.337, fontSize:14, bold:true });
  roundRect(s, { x:1.388, y:5.644, w:1.197, h:0.418, fill:PINK, radius:0.035 });
  tx(s, 'Learn More', { x:1.419, y:5.71, w:1.135, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
}

function slide23(s) {
  round1Rect(s, { x:0, y:2.921, w:6.333, h:4.579, fill:PINK, radius:0.279 });
  chrome(s, { nav:GRAY, brand:DARK, logo:PINK });
  tx(s, 'Makeup Tips For All-day Wear', { x:7.832, y:1.749, w:3.876, h:1.313, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  tx(s, [{ text:'Lorem ipsum dolor sit amet, consectetur adipisicing, seeiusmod tempor incididunt ut labore ' }, { text:'et' }], { x:7.832, y:3.688, w:4.527, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  roundRect(s, { x:1.232, y:4.814, w:3.756, h:1.084, fill:WHITE, radius:0.11, shadow:true });
  arrowBadge(s, { x:1.59, y:5.107, size:0.494 });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit', { x:2.275, y:4.972, w:2.622, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Start With Skincare', { x:7.832, y:3.273, w:2.482, h:0.337, fontSize:14, bold:true });
  icon(s, 'check', { x:7.938, y:4.718, w:0.191, h:0.191, color:PINK });
  tx(s, 'Primer Is Key', { x:8.319, y:4.641, w:1.672, h:0.337, fontSize:14, bold:true, color:DARK });
  icon(s, 'check', { x:10.269, y:4.718, w:0.191, h:0.191, color:PINK });
  tx(s, 'Set Your Base', { x:10.651, y:4.641, w:1.672, h:0.337, fontSize:14, bold:true, color:DARK });
  icon(s, 'check', { x:7.94, y:5.286, w:0.191, h:0.191, color:PINK });
  tx(s, 'Eyebrow Gel', { x:8.322, y:5.209, w:1.672, h:0.337, fontSize:14, bold:true, color:DARK });
  icon(s, 'check', { x:10.272, y:5.286, w:0.191, h:0.191, color:PINK });
  tx(s, 'Setting Spray', { x:10.653, y:5.209, w:1.672, h:0.337, fontSize:14, bold:true, color:DARK });
  roundRect(s, { x:7.938, y:5.854, w:1.197, h:0.418, fill:PINK, radius:0.035 });
  tx(s, 'Read More', { x:7.968, y:5.92, w:1.135, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
}

function slide24(s) {
  chrome(s, { nav:GRAY, brand:DARK, logo:PINK });
  tx(s, 'Guide To Choosing An Effective Sunscreen', { x:1.275, y:1.387, w:5.408, h:1.313, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  roundRect(s, { x:6.667, y:5.339, w:3.756, h:1.084, fill:WHITE, radius:0.11, shadow:true });
  arrowBadge(s, { x:7.025, y:5.632, size:0.494 });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit', { x:7.71, y:5.496, w:2.622, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Broad-Spectrum Protection', { x:1.275, y:2.925, w:3.045, h:0.337, fontSize:14, bold:true });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing eli, sed eiusmod tempor incididunt ut labore et dolore magna', { x:1.275, y:3.279, w:5.058, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  roundRect(s, { x:1.393, y:4.117, w:1.197, h:0.418, fill:PINK, radius:0.035 });
  tx(s, 'Read More', { x:1.424, y:4.183, w:1.135, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
  tx(s, 'Water Resistance', { x:1.275, y:4.813, w:2.244, h:0.337, fontSize:14, bold:true });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing eli, sed eiusmod tempor incididunt ut labore et dolore magna', { x:1.275, y:5.167, w:5.058, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  roundRect(s, { x:1.393, y:6.005, w:1.197, h:0.418, fill:PINK, radius:0.035 });
  tx(s, 'Learn More', { x:1.424, y:6.071, w:1.135, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
}

function slide25(s) {
  round1Rect(s, { x:0, y:2.921, w:6.333, h:4.579, fill:PINK, radius:0.279 });
  chrome(s, { nav:GRAY, brand:DARK, logo:PINK });
  tx(s, 'Concealer Techniques For A Spotless Look', { x:7.784, y:1.704, w:3.899, h:1.919, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing eli, sed eiusmod tempor incididunt ut labore et dolore', { x:7.784, y:3.75, w:4.655, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  roundRect(s, { x:1.602, y:4.902, w:3.756, h:1.084, fill:WHITE, radius:0.11, shadow:true });
  arrowBadge(s, { x:1.96, y:5.195, size:0.494 });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit', { x:2.646, y:5.059, w:2.622, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  roundRect(s, { x:7.784, y:4.679, w:5.007, h:1.284, fill:WHITE, radius:0.134, shadow:true });
  ellipse(s, { x:12.314, y:4.895, w:0.849, h:0.849, fill:SMOKE });
  chevron(s, { x:12.554, y:5.2, w:0.368, h:0.239, color:PINK });
  tx(s, 'Lorem ipsum dolor sit, consectetur adipisicing eli, sed eiusmod tempor', { x:8.895, y:4.966, w:3.114, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  hourglass(s, { x:8.196, y:5.146, w:0.242, h:0.361, color:PINK });
}

function slide26(s) {
  chrome(s, { nav:GRAY, brand:DARK, logo:PINK });
  roundRect(s, { x:2.722, y:4.647, w:1.811, h:1.227, fill:PINK });
  tx(s, 'Aloe Vera\'s Marvelous Beauty Benefits', { x:1.275, y:1.824, w:5.408, h:1.313, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing eli, sed eiusmod tempor incididunt ut labore et dolore magna', { x:1.275, y:3.679, w:5.058, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, '44%', { x:1.275, y:4.768, w:1.516, h:0.505, fontSize:24, fontFace:HEAD, bold:true, color:INK });
  tx(s, 'Treatment', { x:1.275, y:5.322, w:1.35, h:0.337, fontSize:14, bold:true });
  tx(s, '224+', { x:2.907, y:4.768, w:1.445, h:0.505, fontSize:24, fontFace:HEAD, bold:true, color:WHITE });
  tx(s, 'Conditioning', { x:2.907, y:5.322, w:1.445, h:0.337, fontSize:14, bold:true, color:WHITE });
  tx(s, '544+', { x:4.72, y:4.768, w:1.516, h:0.505, fontSize:24, fontFace:HEAD, bold:true, color:INK });
  tx(s, 'Exfoliation', { x:4.72, y:5.322, w:1.516, h:0.337, fontSize:14, bold:true });
  tx(s, 'Soothing Sunburns', { x:1.275, y:3.281, w:2.244, h:0.337, fontSize:14, bold:true });
}

function slide27(s) {
  rect(s, { x:0, y:3.032, w:5.714, h:4.468, fill:PINK });
  chrome(s, { nav:GRAY, brand:DARK, logo:PINK });
  tx(s, 'Selecting Hair Dyes That Are Kind To The Skin', { x:7.707, y:1.76, w:4.029, h:1.919, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing eli, sed eiusmod tempor incididunt ut labore et dolore', { x:7.707, y:4.793, w:4.787, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  icon(s, 'check', { x:7.8, y:3.922, w:0.191, h:0.191, color:PINK });
  tx(s, 'Proper Aftercare', { x:8.082, y:3.845, w:1.862, h:0.337, fontSize:14, bold:true, color:DARK });
  icon(s, 'check', { x:10.132, y:3.922, w:0.191, h:0.191, color:PINK });
  tx(s, 'Hypoallergenic', { x:10.413, y:3.845, w:1.672, h:0.337, fontSize:14, bold:true, color:DARK });
  tx(s, 'Avoid Frequent Dyeing', { x:7.707, y:4.443, w:2.89, h:0.337, fontSize:14, bold:true, color:DARK });
  roundRect(s, { x:7.8, y:5.662, w:1.197, h:0.418, fill:PINK, radius:0.035 });
  tx(s, 'Learn More', { x:7.831, y:5.728, w:1.135, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
  roundRect(s, { x:1.363, y:5.242, w:2.747, h:1.178, fill:WHITE, radius:0.137, shadow:true });
  tx(s, '544+', { x:1.978, y:5.398, w:1.516, h:0.505, fontSize:24, fontFace:HEAD, bold:true, color:INK, align:'center' });
  tx(s, 'Patch Tests', { x:1.747, y:5.902, w:1.98, h:0.337, fontSize:14, bold:true, color:DARK, align:'center' });
}

function slide28(s) {
  chrome(s, { nav:GRAY, brand:DARK, logo:PINK });
  tx(s, 'Effective Postpartum Beauty Care For You', { x:1.1, y:1.729, w:5.408, h:1.313, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  roundRect(s, { x:1.19, y:3.296, w:8.382, h:2.912, fill:PINK, radius:0.119 });
  icon(s, 'tag', { x:1.726, y:3.708, w:0.226, h:0.259, color:WHITE });
  icon(s, 'burst', { x:1.765, y:4.855, w:0.25, h:0.259 });
  icon(s, 'ring', { x:5.639, y:3.708, w:0.259, h:0.259, color:WHITE });
  icon(s, 'pin', { x:5.646, y:4.85, w:0.259, h:0.259, color:WHITE });
  tx(s, 'Lorem ipsum dolor sit, consectetur adipisicing eli, sed eiusmod', { x:2.275, y:3.94, w:3.047, h:0.675, fontSize:12, color:WHITE, lineSpacingMultiple:1.5 });
  tx(s, 'Avoid Frequent Dyeing', { x:2.275, y:3.643, w:2.89, h:0.337, fontSize:14, bold:true, color:WHITE });
  tx(s, 'Lorem ipsum dolor sit, consectetur adipisicing eli, sed eiusmod', { x:2.275, y:5.123, w:3.047, h:0.675, fontSize:12, color:WHITE, lineSpacingMultiple:1.5 });
  tx(s, 'Avoid Frequent Dyeing', { x:2.275, y:4.827, w:2.89, h:0.337, fontSize:14, bold:true, color:WHITE });
  tx(s, 'Lorem ipsum dolor sit, consectetur adipisicing eli, sed eiusmod', { x:6.17, y:3.929, w:3.047, h:0.675, fontSize:12, color:WHITE, lineSpacingMultiple:1.5 });
  tx(s, 'Avoid Frequent Dyeing', { x:6.17, y:3.633, w:2.89, h:0.337, fontSize:14, bold:true, color:WHITE });
  tx(s, 'Lorem ipsum dolor sit, consectetur adipisicing eli, sed eiusmod', { x:6.17, y:5.113, w:3.047, h:0.675, fontSize:12, color:WHITE, lineSpacingMultiple:1.5 });
  tx(s, 'Avoid Frequent Dyeing', { x:6.17, y:4.817, w:2.89, h:0.337, fontSize:14, bold:true, color:WHITE });
}

function slide29(s) {
  rect(s, { x:0, y:3.491, w:5.587, h:4.009, fill:PINK });
  photo(s, { x:1.274, y:1.342, w:2.454, h:4.815 });
  photo(s, { x:4.005, y:2.089, w:2.454, h:4.815 });
  chrome(s, { nav:GRAY, brand:DARK, logo:PINK });
  tx(s, 'Green Tea\'s Extraordinary Beauty Benefits', { x:7.59, y:1.588, w:4.029, h:1.919, fontSize:36, fontFace:HEAD, bold:true, color:DARK });
  checkBadge(s, { x:7.719, y:3.81, size:0.347 });
  checkBadge(s, { x:7.719, y:4.618, size:0.347 });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing eli, sed eiusmod tempor incididunt ut labore et dolore magna', { x:8.273, y:3.672, w:4.221, h:0.676, fontSize:12, fontFace:'Calibri', color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing eli, sed eiusmod tempor incididunt ut labore et dolore magna', { x:8.273, y:4.47, w:4.221, h:0.676, fontSize:12, fontFace:'Calibri', color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing eli, sed eiusmod tempor incididunt ut labore et dolore', { x:7.59, y:5.284, w:4.787, h:0.675, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  roundRect(s, { x:7.719, y:6.183, w:1.197, h:0.418, fill:PINK, radius:0.035 });
  tx(s, 'Learn More', { x:7.75, y:6.248, w:1.135, h:0.286, fontSize:11, bold:true, color:WHITE, align:'center' });
}

function slide30(s) {
  chrome(s, { nav:GRAY, brand:DARK, logo:PINK });
  rect(s, { x:0, y:1.882, w:7.236, h:4.587, fill:WHITE, shadow:true });
  tx(s, [{ text:'THANK ' }, { text:'YOU', options:{ color:PINK } }], { x:1.079, y:2.581, w:5.722, h:1.212, fontSize:66, fontFace:HEAD, bold:true, color:DARK });
  tx(s, 'www.yourwebsite.com', { x:1.784, y:4.402, w:2.251, h:0.372, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Website', { x:1.784, y:4.029, w:1.05, h:0.337, fontSize:14, bold:true });
  tx(s, '0123 Street, City 0123', { x:1.784, y:5.521, w:1.959, h:0.372, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Address', { x:1.784, y:5.148, w:1.05, h:0.337, fontSize:14, bold:true });
  tx(s, '+012 000 000 000', { x:4.439, y:4.391, w:1.959, h:0.372, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Telephone', { x:4.439, y:4.017, w:1.456, h:0.337, fontSize:14, bold:true });
  tx(s, '@youraccountname', { x:4.439, y:5.521, w:1.959, h:0.372, fontSize:12, color:GRAY, lineSpacingMultiple:1.5 });
  tx(s, 'Social Media', { x:4.439, y:5.148, w:1.783, h:0.337, fontSize:14, bold:true });
  icon(s, 'phone', { x:4.005, y:4.123, w:0.18, h:0.18, color:ORANGE });
  icon(s, 'folder', { x:1.316, y:4.065, w:0.292, h:0.259, color:ORANGE });
  icon(s, 'map', { x:1.303, y:5.181, w:0.306, h:0.272, color:ORANGE });
  icon(s, 'bell', { x:4.066, y:5.184, w:0.238, h:0.272, color:ORANGE });
  tx(s, 'Beauty Presentation Template', { x:1.09, y:2.356, w:5.131, h:0.337, fontSize:14, fontFace:HEAD, bold:true, color:PINK, charSpacing:6 });
}

/* ---------------------------------------------------------------- assemble */
const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30];

function build() {
  const pptx = new pptxgen();
  pptx.defineLayout({ name: 'DECK', width: 13.333, height: 7.5 });
  pptx.layout = 'DECK';
  pptx.author = 'Beauty Presentation Template';
  pptx.title = 'Beauty Presentation Template';
  SLIDES.forEach(function (fn) {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    fn(s);
  });
  return pptx.writeFile({ fileName: path.join(__dirname, '12e35bc5-051d-4585-ad18-ffb1c7c9d549_grok_final.pptx') });
}

build().then(function (f) { console.log('wrote ' + f); }).catch(function (e) { console.error(e); process.exit(1); });
