/*
 * Recreation of the "Amor Templates" 30-slide deck with pptxgenjs.
 * Photographic content in the original is replaced with flat colour
 * placeholder rectangles (see `photo()`), everything else is native shapes.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * palette (theme "Blue" + the luminance-modified variants it uses)
 * ------------------------------------------------------------------ */
const C = {
  white: 'FFFFFF',
  blue: '0F6FC6',   // accent1
  blueDk: '0B5394', // accent1 lumMod 75%
  blueLt: '59AAF2', // accent1 lumMod 60% / lumOff 40%
  cyan: '009DD9',   // accent2
  teal: '0BD0D9',   // accent3
  tealDk: '089CA3', // accent3 lumMod 75%
  green: '10CF9B',  // accent4
  lime: '7CCA62',   // accent5
  olive: 'A5C249',  // accent6
  navy: '17406D',   // tx2
  dark: '404040',   // tx1 lum 25%
  dark15: '262626',
  dark35: '595959',
  gray50: '808080',
  gray60: '666666',
  gray40: '999999',
  gray65: 'A6A6A6', // bg1 lumMod 65%
  gray95: 'F2F2F2', // bg1 lumMod 95%
  grayTxt: '848484',
  photo: '0F6FC6'   // stand-in colour for replaced images
};

const F = { head: 'Montserrat', body: 'Raleway', script: 'Satisfy', alt: 'Lato' };

/* ------------------------------------------------------------------ *
 * repeated copy
 * ------------------------------------------------------------------ */
const T = {
  isum: 'Lorem Isum Aliquam varius adipiscing tempor. Vivamus id ipsum sit amet massa consectetur porta. ' +
        'Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos.',
  lacus: 'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo libero vivamus ' +
         'porttitor dolor, conubia mollit. Sapien nam suspendisse, tincidunt eget ante tincidunt, eros in auctor',
  soulA: 'A Amor Powerpoint has taken possession of my entire soul, like these sweet mornings of spring which I ' +
         'enjoy with my whole heart. I am alone, and feel the charm of existence in this spot, which.',
  soulBest: 'Best Powerpoint has taken possession of my entire soul, like these sweet mornings of spring which I ' +
            'enjoy with my whole heart. I am alone, and feel the charm of existence in this spot, which.',
  soulMoob: 'A Moob Powerpoint has taken possession of my entire soul, like these sweet mornings of spring which I ' +
            'enjoy with my whole heart. I am alone, and feel the charm of existence in this spot, which.',
  vivamus: 'Vivamus quam dolor, tempor ac gravida sit amet, porta fermentum.',
  perspic: 'Perspiciatis unde omnis ut iste voluptatem sed fringilla.',
  vestib: 'Vestibulum neque elit, Class aptent taciti sociosqu ad litora torquent per conubia nostra.',
  etiam: 'Etiam faucibus tortor a ipsum vehicula sed hendrerit eros suscipit. Vestibulum neque elit, Class aptent ' +
         'taciti sociosqu ad litora torquent.',
  nunc: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nunc bibendum eleifend tortor, non porta justo gravida posuere. ',
  eiusmod: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.'
};

/* ------------------------------------------------------------------ *
 * low level helpers
 * ------------------------------------------------------------------ */
const NOLINE = { type: 'none' };

function shape(s, kind, x, y, w, h, opt) {
  s.addShape(kind, Object.assign({ x: x, y: y, w: w, h: h, line: NOLINE }, opt || {}));
}
const NOFILL = { type: 'none' };
/** a null colour means "no fill" - handy for outline icons over coloured shapes */
function paint(color) { return color ? { color: color } : NOFILL; }
function box(s, x, y, w, h, color, opt) {
  shape(s, 'rect', x, y, w, h, Object.assign({ fill: paint(color) }, opt || {}));
}
function oval(s, x, y, w, h, color, opt) {
  shape(s, 'ellipse', x, y, w, h, Object.assign({ fill: paint(color) }, opt || {}));
}
function circle(s, cx, cy, r, color, opt) {
  oval(s, cx - r, cy - r, 2 * r, 2 * r, color, opt);
}
function ring(s, cx, cy, r, thick, color) {           // filled circle + hole punched with white
  circle(s, cx, cy, r, color);
  circle(s, cx, cy, r - thick, C.white);
}
/** jigsaw tab: a circle straddling a seam; the half inside its own piece is covered again
 *  so only the protruding half keeps the white outline. `dir` is where the tab points. */
function knob(s, cx, cy, r, color, dir) {
  circle(s, cx, cy, r, color, { line: { color: C.white, width: 4.25 } });
  var m = 0.05;                                       // overlap that hides the stroke
  var cover = { right: [cx - r - m, cy - r - m, r + m, 2 * (r + m)],
                left: [cx, cy - r - m, r + m, 2 * (r + m)],
                up: [cx - r - m, cy, 2 * (r + m), r + m],
                down: [cx - r - m, cy - r - m, 2 * (r + m), r + m] }[dir];
  box(s, cover[0], cover[1], cover[2], cover[3], color);
}
function poly(s, x, y, w, h, pts, color, opt) {       // pts are fractions of w/h
  shape(s, 'custGeom', x, y, w, h, Object.assign({
    fill: paint(color),
    points: pts.map(function (p) { return { x: p[0] * w, y: p[1] * h }; }).concat([{ close: true }])
  }, opt || {}));
}
function stroke(s, x1, y1, x2, y2, color, pt) {
  s.addShape('line', { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
    line: { color: color, width: pt }, flipH: x2 < x1, flipV: y2 < y1 });
}
function txt(s, str, x, y, w, h, o) {
  s.addText(str, Object.assign({ x: x, y: y, w: w, h: h, valign: 'top', fontFace: F.body }, o || {}));
}
/** placeholder standing in for a photograph in the source deck */
function photo(s, x, y, w, h, color) {
  box(s, x, y, w, h, color || C.photo);
}

/* ------------------------------------------------------------------ *
 * recurring composite blocks
 * ------------------------------------------------------------------ */
/** big two-line headline: bold line + light line */
function headline(s, x, y, w, h, l1, l2, c1, c2, o) {
  s.addText([
    { text: l1, options: { bold: true, color: c1, breakLine: true } },
    { text: l2, options: { bold: false, color: c2 } }
  ], Object.assign({ x: x, y: y, w: w, h: h, fontFace: F.head, fontSize: 36, charSpacing: -1.5, valign: 'top' }, o || {}));
}
/** small paragraph used all over the deck (11pt Raleway, 150% leading) */
function para(s, x, y, w, h, str, color, o) {
  txt(s, str, x, y, w, h, Object.assign({ fontSize: 11, color: color, lineSpacingMultiple: 1.5,
    margin: [4.8, 4.8, 2.4, 2.4] }, o || {}));
}
/** rounded "Check it Out" style pill */
function pill(s, x, y, w, h, str, wrap) {
  s.addText(str, { x: x, y: y, w: w, h: h, shape: 'roundRect', rectRadius: 0.08, fill: { color: C.blue },
    align: 'center', valign: 'top', color: C.white, fontFace: F.head, fontSize: 20, italic: true,
    wrap: wrap === true, margin: [7.2, 7.2, 3.6, 3.6] });
}
/** circular check-mark bullet + title + description */
function checkItem(s, x, y, title, dotColor, titleY, descY, tw) {
  circle(s, x + 0.191, y + 0.191, 0.191, dotColor);
  checkMark(s, x + 0.191, y + 0.191, 0.115, C.white);
  txt(s, title, x + 0.479, titleY, tw || 1.823, 0.353, { fontSize: 12, color: C.dark, lineSpacingMultiple: 1.2 });
  txt(s, T.vivamus, x + 0.479, descY, 2.567, 0.464, { fontSize: 9, color: C.gray65, lineSpacingMultiple: 1.2 });
}
function checkMark(s, cx, cy, r, color) {
  poly(s, cx - r, cy - r * 0.8, 2 * r, 1.6 * r,
    [[0.04, 0.46], [0.22, 0.28], [0.42, 0.5], [0.8, 0.1], [0.98, 0.28], [0.42, 0.86]], color);
}
/** teardrop map-pin (rotated so the tip points down-left) */
function pin(s, x, y, size, color) {
  shape(s, 'teardrop', x, y, size, size, { rotate: 44.45, flipV: true, fill: { color: color } });
}
/** "AMOR INFOGRAPHIC" banner shared by slides 18-29 */
function infoTitle(s) {
  s.addText([
    { text: 'AMOR', options: { color: C.blue } },
    { text: ' INFOGRAPHIC', options: { color: C.dark } }
  ], { x: 2.924, y: 0.531, w: 7.486, h: 0.64, align: 'center', valign: 'top',
       fontFace: F.head, fontSize: 32, bold: true, italic: true, charSpacing: -1.5 });
}
/** dashed progress ring: 12 ticks starting at 12 o'clock, `on` of them coloured */
function dashRing(s, cx, cy, r, thick, on, color, offColor) {
  for (var k = 0; k < 12; k++) {
    var a0 = -87 + k * 30;
    s.addShape('blockArc', { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, line: NOLINE,
      angleRange: [(a0 + 360) % 360, (a0 + 24 + 360) % 360], arcThicknessRatio: thick / r,
      fill: { color: k < on ? color : offColor } });
  }
}

/* ------------------------------------------------------------------ *
 * line-art icons - each draws inside the box (cx, cy) with radius r
 * ------------------------------------------------------------------ */
const LW = 1;   // outline weight in points

function iOutlineCircle(s, cx, cy, r, c) {
  shape(s, 'ellipse', cx - r, cy - r, 2 * r, 2 * r, { fill: NOFILL, line: { color: c, width: LW } });
}
const ICON = {
  plane: function (s, cx, cy, r, c) {
    poly(s, cx - r, cy - r, 2 * r, 2 * r, [[1, 0], [0, 0.52], [0.36, 0.63], [0.47, 1], [0.6, 0.56]], null,
      { line: { color: c, width: LW * 1.5 } });
  },
  clock: function (s, cx, cy, r, c) {
    iOutlineCircle(s, cx, cy, r, c);
    iOutlineCircle(s, cx, cy, r * 0.8, c);
    stroke(s, cx, cy - r * 0.5, cx, cy, c, LW);
    stroke(s, cx, cy, cx + r * 0.36, cy + r * 0.18, c, LW);
    circle(s, cx, cy, r * 0.08, c);
  },
  gear: function (s, cx, cy, r, c) {
    shape(s, 'gear9', cx - r, cy - r, 2 * r, 2 * r, { fill: NOFILL, line: { color: c, width: LW } });
    iOutlineCircle(s, cx, cy, r * 0.3, c);
  },
  gearSolid: function (s, cx, cy, r, c, hole) {
    shape(s, 'gear6', cx - r, cy - r, 2 * r, 2 * r, { fill: { color: c } });
    circle(s, cx, cy, r * 0.24, hole || C.blue);
  },
  life: function (s, cx, cy, r, c) {
    iOutlineCircle(s, cx, cy, r, c);
    iOutlineCircle(s, cx, cy, r * 0.42, c);
    [45, 135, 225, 315].forEach(function (a) {
      var t = a * Math.PI / 180;
      stroke(s, cx + Math.cos(t) * r * 0.42, cy + Math.sin(t) * r * 0.42,
        cx + Math.cos(t) * r, cy + Math.sin(t) * r, c, LW);
    });
  },
  wrench: function (s, cx, cy, r, c) {
    shape(s, 'roundRect', cx - r * 0.13, cy - r * 0.25, r * 0.26, r * 1.25,
      { rotate: -40, rectRadius: 0.45, fill: { type: 'none' }, line: { color: c, width: LW } });
    shape(s, 'blockArc', cx - r * 0.95, cy - r * 1.05, r * 1.1, r * 1.1,
      { angleRange: [50, 330], arcThicknessRatio: 0.3, fill: { type: 'none' }, line: { color: c, width: LW } });
  },
  search: function (s, cx, cy, r, c) {
    iOutlineCircle(s, cx - r * 0.22, cy - r * 0.22, r * 0.6, c);
    stroke(s, cx + r * 0.18, cy + r * 0.18, cx + r * 0.85, cy + r * 0.85, c, LW * 1.6);
  },
  home: function (s, cx, cy, r, c, outline) {
    poly(s, cx - r, cy - r * 0.9, 2 * r, r * 1.8,
      [[0.5, 0], [1, 0.45], [0.86, 0.45], [0.86, 1], [0.14, 1], [0.14, 0.45], [0, 0.45]],
      outline ? null : c, outline ? { line: { color: c, width: LW } } : {});
  },
  star: function (s, cx, cy, r, c, hole) {
    shape(s, 'star5', cx - r * 1.05, cy - r, r * 2.1, r * 2, { fill: { color: c } });
    shape(s, 'star5', cx - r * 0.66, cy - r * 0.63, r * 1.32, r * 1.26, { fill: { color: hole || C.blue } });
  },
  drop: function (s, cx, cy, r, c) {
    shape(s, 'teardrop', cx - r * 0.78, cy - r * 0.82, r * 1.56, r * 1.56, { rotate: 225, fill: { color: c } });
  },
  stack: function (s, cx, cy, r, c, outline) {
    var o = outline ? { fill: NOFILL, line: { color: c, width: LW } } : { fill: { color: c } };
    [-0.56, 0, 0.56].forEach(function (dy) {
      shape(s, 'can', cx - r * 0.78, cy + dy * r - r * 0.3, r * 1.56, r * 0.72, o);
    });
  },
  lock: function (s, cx, cy, r, c) {
    shape(s, 'blockArc', cx - r * 0.48, cy - r * 1.05, r * 0.96, r * 0.96,
      { angleRange: [180, 0], arcThicknessRatio: 0.26, fill: { color: c } });
    shape(s, 'roundRect', cx - r * 0.72, cy - r * 0.42, r * 1.44, r * 1.32,
      { rectRadius: 0.14, fill: { color: c } });
    circle(s, cx, cy + r * 0.16, r * 0.17, C.dark);
    box(s, cx - r * 0.08, cy + r * 0.16, r * 0.16, r * 0.4, C.dark);
  },
  wifi: function (s, cx, cy, r, c) {
    [1, 0.66, 0.34].forEach(function (k) {
      shape(s, 'blockArc', cx - r * k, cy - r * k + r * 0.55, 2 * r * k, 2 * r * k,
        { angleRange: [215, 325], arcThicknessRatio: 0.2, fill: { color: c } });
    });
    circle(s, cx, cy + r * 0.55, r * 0.14, c);
  },
  cloud: function (s, cx, cy, r, c) {
    shape(s, 'cloud', cx - r * 1.1, cy - r * 0.95, r * 2.2, r * 1.5, { fill: { color: c } });
    poly(s, cx - r * 0.38, cy - r * 0.2, r * 0.76, r * 1.0,
      [[0.5, 1], [0, 0.42], [0.3, 0.42], [0.3, 0], [0.7, 0], [0.7, 0.42], [1, 0.42]], C.dark);
  },
  music: function (s, cx, cy, r, c) {
    box(s, cx - r * 0.14, cy - r * 0.85, r * 0.14, r * 1.35, c);
    box(s, cx + r * 0.62, cy - r * 0.85, r * 0.14, r * 1.35, c);
    box(s, cx - r * 0.14, cy - r * 0.85, r * 0.9, r * 0.24, c);
    circle(s, cx - r * 0.33, cy + r * 0.5, r * 0.3, c);
    circle(s, cx + r * 0.41, cy + r * 0.5, r * 0.3, c);
  },
  clip: function (s, cx, cy, r, c) {
    shape(s, 'roundRect', cx - r * 0.62, cy - r * 0.8, r * 1.24, r * 1.6,
      { rectRadius: 0.1, fill: NOFILL, line: { color: c, width: LW * 1.4 } });
    box(s, cx - r * 0.3, cy - r * 0.95, r * 0.6, r * 0.28, null, { line: { color: c, width: LW * 1.4 } });
    [-0.3, -0.06, 0.18, 0.42].forEach(function (k) {
      stroke(s, cx - r * 0.36, cy + k * r, cx + r * 0.36, cy + k * r, c, LW * 1.2);
    });
  },
  tray: function (s, cx, cy, r, c) {
    [-0.95, 0.05].forEach(function (dy) {
      poly(s, cx - r, cy + dy * r, 2 * r, r * 0.8, [[0.18, 0], [0.82, 0], [1, 1], [0, 1]], null,
        { line: { color: c, width: LW } });
    });
  },
  wallet: function (s, cx, cy, r, c) {
    shape(s, 'roundRect', cx - r, cy - r * 0.65, 2 * r, r * 1.3,
      { rectRadius: 0.14, fill: NOFILL, line: { color: c, width: LW } });
    shape(s, 'roundRect', cx + r * 0.22, cy - r * 0.32, r * 0.76, r * 0.64,
      { rectRadius: 0.25, fill: NOFILL, line: { color: c, width: LW } });
    circle(s, cx + r * 0.6, cy, r * 0.11, c);
  },
  bin: function (s, cx, cy, r, c) {
    poly(s, cx - r * 0.8, cy - r * 0.45, r * 1.6, r * 1.35, [[0, 0], [1, 0], [0.88, 1], [0.12, 1]], null,
      { line: { color: c, width: LW } });
    box(s, cx - r, cy - r * 0.85, 2 * r, r * 0.36, null, { line: { color: c, width: LW } });
    box(s, cx - r * 0.26, cy - r * 1.05, r * 0.52, r * 0.2, null, { line: { color: c, width: LW } });
  },
  bank: function (s, cx, cy, r, c) {
    poly(s, cx - r, cy - r * 0.95, 2 * r, r * 0.5, [[0.5, 0], [1, 1], [0, 1]], null, { line: { color: c, width: LW } });
    [-0.62, -0.21, 0.21, 0.62].forEach(function (k) { box(s, cx + k * r - r * 0.08, cy - r * 0.4, r * 0.16, r * 1.05, c); });
    box(s, cx - r, cy + r * 0.68, 2 * r, r * 0.22, c);
  },
  cart: function (s, cx, cy, r, c) {
    poly(s, cx - r * 0.95, cy - r * 0.35, r * 1.9, r * 1.0, [[0, 0], [1, 0], [0.84, 1], [0.16, 1]], null,
      { line: { color: c, width: LW } });
    circle(s, cx - r * 0.45, cy + r * 0.88, r * 0.16, c);
    circle(s, cx + r * 0.45, cy + r * 0.88, r * 0.16, c);
    stroke(s, cx - r * 0.6, cy - r * 0.85, cx + r * 0.5, cy - r * 0.85, c, LW);
    poly(s, cx + r * 0.2, cy - r * 1.05, r * 0.5, r * 0.4, [[0, 0.5], [0.7, 0.5], [0.5, 0], [1, 0.5], [0.5, 1], [0.7, 0.5]], c);
  },
  monitor: function (s, cx, cy, r, c) {
    box(s, cx - r, cy - r * 0.7, r * 1.42, r * 1.1, null, { line: { color: c, width: LW * 1.6 } });
    box(s, cx - r * 0.5, cy + r * 0.4, r * 0.42, r * 0.3, c);
    box(s, cx - r * 0.8, cy + r * 0.7, r * 1.02, r * 0.16, c);
    box(s, cx + r * 0.56, cy - r * 0.7, r * 0.44, r * 1.56, null, { line: { color: c, width: LW * 1.6 } });
  },
  rss: function (s, cx, cy, r, c) {
    var ox = cx - r * 0.6, oy = cy + r * 0.62;          // waves radiate from this dot
    [1.6, 1.0].forEach(function (k) {
      shape(s, 'blockArc', ox - r * k, oy - r * k, 2 * r * k, 2 * r * k,
        { angleRange: [270, 0], arcThicknessRatio: 0.13, fill: { color: c } });
    });
    circle(s, ox, oy, r * 0.16, c);
  },
  people: function (s, cx, cy, r, c) {
    [[-0.62, 0.85], [0.62, 0.85], [0, 1.05]].forEach(function (p) {
      circle(s, cx + p[0] * r, cy - r * 0.4, r * 0.3 * p[1], c);
      shape(s, 'blockArc', cx + p[0] * r - r * 0.5 * p[1], cy - r * 0.3, r * p[1], r * 1.2 * p[1],
        { angleRange: [180, 0], arcThicknessRatio: 0.5, fill: { color: c } });
    });
  },
  person: function (s, cx, cy, r, c) {
    circle(s, cx, cy - r * 0.45, r * 0.34, c);
    shape(s, 'blockArc', cx - r * 0.62, cy - r * 0.2, r * 1.24, r * 1.4,
      { angleRange: [180, 0], arcThicknessRatio: 0.5, fill: { color: c } });
  },
  globe: function (s, cx, cy, r, c) {
    circle(s, cx, cy, r, c);
    shape(s, 'ellipse', cx - r * 0.42, cy - r, r * 0.84, 2 * r, { fill: NOFILL, line: { color: C.white, width: LW } });
    stroke(s, cx - r * 0.98, cy, cx + r * 0.98, cy, C.white, LW);
  },
  starGroup: function (s, cx, cy, r, c) {
    ICON.person(s, cx, cy - r * 0.25, r * 0.78, c);
    [-0.55, 0, 0.55].forEach(function (k) {
      shape(s, 'star5', cx + k * r * 0.62 - r * 0.22, cy + r * 0.5, r * 0.44, r * 0.44, { fill: { color: c } });
    });
  }
};

/* ------------------------------------------------------------------ *
 * slide builders
 * ------------------------------------------------------------------ */
function slideTitle(s, word) {                     // slides 1 and 30
  photo(s, 0, 0, 13.333, 7.5);
  s.addText(word, { x: 2.962, y: 2.691, w: 7.408, h: 1.313, align: 'center', valign: 'top',
    fontFace: F.script, fontSize: 72, bold: true, color: C.white });
  s.addText([{ text: 'Amor', options: { breakLine: true } }, { text: 'Templats' }],
    { x: 5.445, y: 4.186, w: 2.443, h: 0.484, shape: 'roundRect', rectRadius: 0.08, fill: { color: C.blue },
      align: 'center', valign: 'top', color: C.white, fontFace: F.head, fontSize: 20, italic: true,
      margin: [7.2, 7.2, 3.6, 3.6] });
  para(s, 4.313, 5.176, 4.707, 0.9, T.isum, C.white, { align: 'center' });
}

function slide01(s) { slideTitle(s, 'Amor'); }
function slide30(s) { slideTitle(s, 'Thanks'); }

function slide02(s) {
  photo(s, 0, 0, 6.611, 7.5);
  para(s, 7.503, 3.909, 4.707, 0.9, T.isum, C.dark);
  headline(s, 7.444, 2.206, 5.68, 1.313, 'Amor Templates', 'For Fashionista', C.dark, C.dark);
  pill(s, 8.286, 5.608, 1.998, 0.484, 'Check it Out');
}

function slide03(s) {
  photo(s, 6.722, 0, 6.611, 7.5);
  box(s, 6.722, 5.845, 6.611, 1.655, C.dark, { fill: { color: C.dark, transparency: 14 } });
  para(s, 1.355, 3.09, 4.707, 0.9, T.isum, C.dark, { align: 'right' });
  headline(s, 0.381, 1.247, 5.68, 1.313, 'Amor Templates', 'For Fashionista', C.blue, C.dark, { align: 'right' });
  para(s, 7.553, 6.222, 4.707, 0.9, T.isum, C.white, { align: 'justify' });
  [[4.522, C.dark15], [5.716, C.blue]].forEach(function (row, i) {
    pin(s, i ? 2.153 : 2.164, row[0], 0.653, row[1]);
    (i ? ICON.tray : ICON.stack)(s, (i ? 2.153 : 2.164) + 0.326, row[0] + 0.30, 0.15, C.white, true);
    txt(s, 'Simple', 3.341, row[0] + 0.005, 0.861, 0.337, { fontSize: 14, bold: true, color: C.navy });
    txt(s, 'Lorem ipsum dolor sit amet, adipiscing elit.', 3.341, row[0] + 0.316, 2.135, 0.561,
      { fontSize: 11, color: C.navy, lineSpacingMultiple: 1.2 });
  });
}

function slide04(s) {
  photo(s, 4.458, 0, 3.722, 7.5);
  box(s, 8.181, 0, 5.153, 7.5, C.dark);
  para(s, 8.798, 3.912, 3.947, 1.178, T.isum, C.white);
  headline(s, 8.742, 1.481, 4.062, 1.919, 'Amor Templates', 'For Fashionista', C.blue, C.white);
  var rows = [[2.135, 2.44, 0.794, 2.228, C.blue], [3.04, 3.345, 0.794, 3.141, C.blue],
              [3.935, 4.24, 0.806, 4.028, C.blue], [4.838, 5.143, 0.806, 4.939, C.navy]];
  rows.forEach(function (r) {
    var lx = r[0] < 3.5 ? 1.273 : 1.285;
    circle(s, r[2] + 0.191, r[3] + 0.191, 0.191, r[4]);
    checkMark(s, r[2] + 0.191, r[3] + 0.191, 0.115, C.white);
    txt(s, 'Your Title Here', lx, r[0], 1.823, 0.353, { fontSize: 12, color: C.dark, lineSpacingMultiple: 1.2 });
    txt(s, T.vivamus, lx, r[1], 2.567, 0.464, { fontSize: 9, color: C.gray65, lineSpacingMultiple: 1.2 });
  });
  txt(s, 'Your Name Here', 8.798, 5.826, 3.807, 0.572, { fontSize: 28, italic: true, charSpacing: -1.5,
    color: C.blue, fontFace: F.head });
}

function slide05(s) {
  photo(s, 0, 0, 3.722, 7.5);
  box(s, 3.722, 0, 4.376, 7.5, C.dark);
  txt(s, 'Amor Slides', 4.346, 1.587, 3.192, 0.707, { fontFace: F.script, fontSize: 36, bold: true,
    charSpacing: -1.5, color: C.blue });
  txt(s, T.soulA, 4.365, 2.392, 3.207, 1.212, { fontSize: 11, color: C.white, align: 'justify', lineSpacingMultiple: 1.2 });
  txt(s, 'About Amor', 4.379, 3.934, 3.192, 0.572, { fontFace: F.head, fontSize: 28, bold: true,
    charSpacing: -1.5, color: C.white });
  txt(s, T.soulBest, 4.365, 4.871, 3.207, 1.212, { fontSize: 11, color: C.white, align: 'justify', lineSpacingMultiple: 1.2 });
  para(s, 8.826, 3.005, 3.469, 1.456, T.isum, C.dark);
  headline(s, 8.62, 1.302, 4.589, 1.313, 'Amor Templates', 'For Fashionista', C.dark, C.dark);
  [[4.869, 5.174, 4.962, C.blue], [5.772, 6.076, 5.872, C.navy]].forEach(function (r) {
    circle(s, 8.826 + 0.191, r[2] + 0.191, 0.191, r[3]);
    checkMark(s, 8.826 + 0.191, r[2] + 0.191, 0.115, C.white);
    txt(s, 'Your Title Here', 9.305, r[0], 1.823, 0.353, { fontSize: 12, color: C.dark, lineSpacingMultiple: 1.2 });
    txt(s, T.vivamus, 9.305, r[1], 2.567, 0.464, { fontSize: 9, color: C.gray65, lineSpacingMultiple: 1.2 });
  });
}

function slide06(s) {
  photo(s, 0, 0, 8.833, 7.5);
  box(s, 0, 0, 8.833, 7.5, C.dark, { fill: { color: C.dark, transparency: 38 } });
  var items = [[1.773, C.blue, C.dark], [2.661, C.blue, C.blue], [3.548, C.dark, C.white],
               [4.435, C.dark, C.white], [5.445, C.dark, C.white]];
  items.forEach(function (it, i) {
    var y = it[0];
    shape(s, 'roundRect', 9.562, y, 0.607, 0.603, { rectRadius: 0.12, fill: { color: it[1] } });
    s.addText('0' + (i + 1), { x: 9.562, y: y + 0.1, w: 0.607, h: 0.4, align: 'center', valign: 'top',
      fontFace: F.alt, fontSize: 18, color: it[2], margin: 0 });
    txt(s, 'Text sample here', 10.35, y - 0.116, 2.426, 0.303, { fontSize: 12, bold: true, color: C.dark, fontFace: F.alt });
    txt(s, 'Lorem ipsum dolor sit dudu amet, consectetur adipiscing elit. ', 10.365, y + 0.171, 2.951, 0.599,
      { fontSize: 11, color: C.dark, fontFace: F.alt, lineSpacingMultiple: 1.1 });
  });
  txt(s, 'Amor', 0.554, 4.736, 3.488, 1.313, { fontFace: F.script, fontSize: 72, bold: true, color: C.blue });
  para(s, 0.399, 6.183, 4.707, 0.9, T.isum, C.white);
}

function slide07(s) {
  box(s, 8.8, 0, 4.533, 7.5, C.dark);
  photo(s, 9.439, 1.08, 3.181, 3.315);
  photo(s, 0.845, 3.285, 1.352, 1.292);
  photo(s, 0.845, 5.114, 1.352, 1.292);
  box(s, 5.366, 0, 3.475, 7.5, C.blue);
  headline(s, 0.777, 1.128, 4.307, 1.313, 'Amor Teams', 'For Fashionista', C.dark, C.dark);
  [[3.452, 'Amor One', 1.143], [5.226, 'Amor Two', 1.163]].forEach(function (r) {
    txt(s, r[1], 2.686, r[0], r[2], 0.337, { fontSize: 14, bold: true, color: C.dark });
    txt(s, T.perspic, 2.686, r[0] + 0.458, 3.18, 0.471, { fontSize: 11, color: C.dark });
  });
  txt(s, T.soulA, 5.833, 2.019, 2.766, 1.434, { fontSize: 11, color: C.dark, align: 'justify', lineSpacingMultiple: 1.2 });
  txt(s, 'Fashionista', 5.847, 3.56, 3.192, 0.572, { fontFace: F.head, fontSize: 28, bold: true, charSpacing: -1.5, color: C.dark });
  txt(s, T.soulBest, 5.833, 4.497, 2.766, 1.434, { fontSize: 11, color: C.dark, align: 'justify', lineSpacingMultiple: 1.2 });
  txt(s, 'Amor Slides', 9.439, 4.76, 3.192, 0.707, { fontFace: F.script, fontSize: 36, bold: true, charSpacing: -1.5, color: C.blue });
  txt(s, T.soulMoob, 9.458, 5.565, 3.207, 1.212, { fontSize: 11, color: C.white, align: 'justify', lineSpacingMultiple: 1.2 });
}

function slide08(s) {
  box(s, 0, 0, 4.533, 7.5, C.dark);
  photo(s, 0.639, 1.08, 3.181, 3.315);
  photo(s, 5.331, 1.08, 1.352, 1.292);
  photo(s, 5.331, 2.908, 1.352, 1.292);
  box(s, 9.858, -0.014, 3.475, 7.5, C.blue);
  [[1.269, 'Amor Three', 1.308], [3.043, 'Amor Four', 1.185]].forEach(function (r) {
    txt(s, r[1], 6.939, r[0], r[2], 0.337, { fontSize: 14, bold: true, color: C.dark });
    txt(s, T.perspic, 6.939, r[0] + 0.458, 3.18, 0.471, { fontSize: 11, color: C.dark });
  });
  txt(s, T.soulA, 10.273, 1.609, 2.766, 1.434, { fontSize: 11, color: C.dark, align: 'justify', lineSpacingMultiple: 1.2 });
  txt(s, 'Galery', 10.287, 3.151, 3.192, 0.572, { fontFace: F.head, fontSize: 28, bold: true, charSpacing: -1.5, color: C.dark });
  txt(s, T.soulBest, 10.273, 4.088, 2.766, 1.434, { fontSize: 11, color: C.dark, align: 'justify', lineSpacingMultiple: 1.2 });
  txt(s, 'Amor Slides', 5.538, 4.737, 3.192, 0.707, { fontFace: F.script, fontSize: 36, bold: true, charSpacing: -1.5, color: C.dark });
  txt(s, T.soulMoob, 5.557, 5.542, 3.207, 1.212, { fontSize: 11, color: C.dark, align: 'justify', lineSpacingMultiple: 1.2 });
  txt(s, 'YOUR TITLE HERE', 0.92, 5.074, 2.624, 0.37, { fontFace: F.head, fontSize: 16, bold: true, italic: true,
    align: 'center', color: C.blue });
  txt(s, T.lacus, 0.513, 5.629, 3.437, 1.01, { fontSize: 9, color: C.white, align: 'center', lineSpacingMultiple: 1.5 });
}

function slide09(s) {
  photo(s, 6.619, 0, 6.714, 3.753);
  photo(s, 0, 3.753, 6.619, 3.747);
  [[1.998, 1.054, 1.591, 1.889], [8.664, 4.761, 8.258, 5.596]].forEach(function (r) {
    txt(s, 'Amor Text', r[0], r[1], 2.624, 0.64, { fontFace: F.script, fontSize: 32, bold: true, italic: true,
      align: 'center', color: C.dark });
    txt(s, T.lacus, r[2], r[3], 3.437, 1.01, { fontSize: 9, color: C.dark, align: 'center', lineSpacingMultiple: 1.5 });
  });
}

function slide10(s) {
  photo(s, 4.111, 0.007, 4.111, 3.747);
  photo(s, 0, 3.753, 4.111, 3.747);
  box(s, 4.111, 0, 4.111, 3.753, C.dark, { fill: { color: C.dark, transparency: 38 } });
  box(s, 4.111, 3.747, 4.111, 3.753, C.gray95, { fill: { color: C.gray95, transparency: 38 } });
  para(s, 0.698, 1.037, 2.786, 1.733, T.isum, C.dark15, { align: 'justify' });
  txt(s, 'Amor', 5.021, 1.247, 3.488, 1.313, { fontFace: F.script, fontSize: 72, bold: true, color: C.blue });
  headline(s, 8.796, 1.6, 4.197, 1.919, 'Amor Templates', 'For Fashionista', C.dark, C.dark);
  para(s, 8.968, 4.05, 3.478, 1.456, T.isum, C.dark);
  pill(s, 9.103, 6.036, 1.998, 0.484, 'Check it Out');
  txt(s, 'Amor Text', 4.855, 4.866, 2.624, 0.64, { fontFace: F.script, fontSize: 32, bold: true, italic: true,
    align: 'center', color: C.dark });
  txt(s, T.lacus, 4.448, 5.701, 3.437, 1.01, { fontSize: 9, color: C.dark, align: 'center', lineSpacingMultiple: 1.5 });
}

function slide11(s) {
  photo(s, 1.063, 0.338, 2.993, 4.211);
  photo(s, 8.906, 2.871, 2.993, 4.211);
  box(s, 4.704, 0.338, 3.817, 6.744, C.dark);
  txt(s, 'Amor Text', 5.301, 2.897, 2.624, 0.64, { fontFace: F.script, fontSize: 32, bold: true, italic: true,
    align: 'center', color: C.blue });
  txt(s, T.lacus, 5.021, 3.779, 3.186, 1.237, { fontSize: 9, color: C.white, align: 'center', lineSpacingMultiple: 1.5 });
  txt(s, T.lacus.replace('auctor', 'auctorç'), 1.063, 5.205, 2.993, 1.237, { fontSize: 9, color: C.dark,
    align: 'justify', lineSpacingMultiple: 1.5 });
  txt(s, T.lacus, 8.906, 0.865, 2.993, 1.237, { fontSize: 9, color: C.dark, align: 'justify', lineSpacingMultiple: 1.5 });
  [[1.063, 4.213], [8.906, 2.871]].forEach(function (p) {
    s.addText('YOUR TITLE HERE', { x: p[0], y: p[1], w: 2.993, h: 0.37, fill: { color: C.blue }, align: 'center',
      valign: 'top', color: C.white, fontFace: F.head, fontSize: 16, bold: true, italic: true });
  });
}

function slide12(s) {
  box(s, 0.569, 0.583, 12.306, 6.208, C.dark);
  photo(s, 1.695, 1.927, 3.181, 3.315);
  photo(s, 8.486, 1.927, 3.181, 3.315);
  txt(s, 'Amor', 5.301, 2.897, 2.624, 0.774, { fontFace: F.script, fontSize: 40, bold: true, italic: true,
    align: 'center', color: C.blue });
  txt(s, T.lacus, 5.369, 4.005, 2.624, 1.237, { fontSize: 9, color: C.white, align: 'justify', lineSpacingMultiple: 1.5 });
  [1.695, 8.486].forEach(function (x) {
    s.addText('YOUR TITLE HERE', { x: x, y: 4.871, w: 3.181, h: 0.37, fill: { color: C.blue }, align: 'center',
      valign: 'top', color: C.white, fontFace: F.head, fontSize: 16, bold: true, italic: true });
  });
}

function slide13(s) {
  box(s, 0, 4.5, 13.333, 3, C.blue);
  box(s, 0.569, 0.583, 12.306, 6.208, C.dark);
  photo(s, 0.569, 0.583, 4.639, 6.208);
  headline(s, 6.31, 1.295, 4.197, 1.919, 'Amor Templates', 'For Fashionista', C.white, C.white);
  para(s, 6.482, 3.744, 3.478, 1.456, T.isum, C.white);
  pill(s, 6.617, 5.73, 1.998, 0.484, 'Check it Out', true);
  s.addText('Amor', { x: 7.905, y: 3.088, w: 7.408, h: 1.313, rotate: 90, align: 'center', valign: 'top',
    fontFace: F.script, fontSize: 72, bold: true, color: C.blue });
}

function slide14(s) {
  box(s, 0, 4.5, 13.333, 3, C.blue);
  box(s, 0.569, 0.583, 12.306, 6.208, C.dark);
  photo(s, 0.569, 0.583, 5.917, 6.208);
  photo(s, 6.958, 0.583, 5.917, 6.208);
  box(s, 6.958, 0.583, 5.917, 6.208, C.dark, { fill: { color: C.dark, transparency: 38 } });
  box(s, 0.569, 5.181, 5.917, 1.611, C.dark, { fill: { color: C.dark, transparency: 38 } });
  s.addText('YOUR TITLE HERE', { x: 3.305, y: 4.81, w: 3.181, h: 0.37, fill: { color: C.blue }, align: 'center',
    valign: 'top', color: C.white, fontFace: F.head, fontSize: 16, bold: true, italic: true });
  txt(s, 'Amor', 7.415, 4.056, 3.488, 1.313, { fontFace: F.script, fontSize: 72, bold: true, color: C.blue });
  para(s, 1.051, 5.502, 5.165, 0.9, T.isum, C.white);
  para(s, 7.26, 5.502, 4.707, 0.9, T.isum, C.white);
}

function slide15(s) {
  box(s, 0, 4.5, 13.333, 3, C.dark);
  [1.046, 3.334, 5.623, 7.911, 10.2].forEach(function (x) { photo(s, x, 1.802, 2.122, 3.301); });
  txt(s, 'Amor Great Model', 3.735, 0.76, 5.864, 0.656, { fontFace: F.head, fontSize: 33, bold: true, italic: true,
    charSpacing: -1.5, align: 'center', color: C.dark });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I ' +
        'enjoy with my whole heart. I am alone, and feel the charm of existence in this spot, which was created for ' +
        'the bliss of souls like mine. I am so happy, my dear friend.',
    1.046, 5.6, 5.357, 0.99, { fontSize: 11, color: C.white, align: 'justify', lineSpacingMultiple: 1.2 });
  txt(s, 'Amor Text', 6.701, 5.761, 2.347, 0.64, { fontFace: F.script, fontSize: 32, bold: true, italic: true,
    align: 'center', color: C.blue });
  txt(s, 'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo libero vivamus ' +
        'porttitor dolor, conubia mollit. Sapien nam', 9.347, 5.618, 3.227, 0.783,
    { fontSize: 9, color: C.white, align: 'justify', lineSpacingMultiple: 1.5 });
}

/** shared right-hand column of slides 16 and 17 */
function mockupColumn(s, tx, cx) {
  headline(s, tx, 1.277, 4.589, 1.313, 'Amor Mockup', 'Creatives ', C.dark, C.dark);
  para(s, tx, 2.903, 3.469, 1.456, T.isum, C.dark);
  [['Amor Mockup', 4.758, 5.062, 4.851, C.blue], ['Your Title Here', 5.66, 5.965, 5.761, C.navy]].forEach(function (r) {
    circle(s, cx + 0.191, r[3] + 0.191, 0.191, r[4]);
    checkMark(s, cx + 0.191, r[3] + 0.191, 0.115, C.white);
    txt(s, r[0], cx + 0.479, r[1], 1.823, 0.353, { fontSize: 12, color: C.dark, lineSpacingMultiple: 1.2 });
    txt(s, T.vivamus, cx + 0.479, r[2], 2.567, 0.464, { fontSize: 9, color: C.gray65, lineSpacingMultiple: 1.2 });
  });
}

function slide16(s) {                                  // tablet mock-up
  box(s, 0, 0, 4.225, 7.5, C.gray95);
  shape(s, 'roundRect', 1.03, 0.73, 4.4, 6.29, { rectRadius: 0.055, fill: { color: '070709' } });
  photo(s, 1.271, 1.277, 3.903, 5.167);
  circle(s, 3.23, 6.72, 0.145, '1E1E22', { line: { color: '4A4A50', width: 0.75 } });
  circle(s, 3.23, 1.0, 0.028, '2B2B30');
  mockupColumn(s, 7.397, 7.52);
  txt(s, 'Amor', 4.409, 6.052, 3.488, 1.313, { fontFace: F.script, fontSize: 72, bold: true, color: C.blue });
}

function slide17(s) {                                  // smart-watch mock-up
  box(s, 0, 0, 4.819, 7.5, C.gray95);
  shape(s, 'roundRect', 1.72, 1.953, 1.44, 1.05, { rectRadius: 0.3, fill: { color: C.white } });   // top strap
  shape(s, 'roundRect', 1.72, 4.55, 1.44, 1.02, { rectRadius: 0.3, fill: { color: C.white } });    // bottom strap
  shape(s, 'roundRect', 1.119, 2.627, 2.52, 2.36, { rectRadius: 0.22, fill: { color: 'C9C9C9' } }); // case
  box(s, 3.63, 3.28, 0.11, 0.36, 'B4B4B4');                                                        // crown
  shape(s, 'roundRect', 1.219, 2.72, 2.32, 2.17, { rectRadius: 0.2, fill: { color: '070707' } });   // bezel
  photo(s, 1.493, 2.908, 1.769, 1.661);
  mockupColumn(s, 6.995, 7.103);
  txt(s, 'Amor', 3.076, 3.772, 3.488, 1.313, { fontFace: F.script, fontSize: 72, bold: true, color: C.blue });
}

function slide18(s) {
  infoTitle(s);
  txt(s, 'Mauris quam dolor, cursus at porta et, luctus eget purus. Nunc tempor luctus interdum. Duis libero leo, ' +
        'consequat ut accumsan eu, viverra et erat. Nunc rhoncus tellus in ipsum molestie et gravida tortor dignissim. ',
    1.588, 1.981, 10.158, 0.505, { fontSize: 10, color: C.gray65, align: 'center', lineSpacingMultiple: 1.2 });
  var drops = [
    [1.15, C.navy, ICON.wifi], [2.923, C.blue, ICON.gearSolid], [4.695, C.dark, ICON.lock],
    [6.483, C.teal, ICON.clip], [8.263, C.dark, ICON.cloud], [10.018, C.blue, ICON.music]
  ];
  drops.forEach(function (d, i) {
    shape(s, 'teardrop', d[0], 2.937, 1.793, 1.793, { rotate: 45, fill: { color: d[1] } });
    d[2](s, d[0] + 0.897, 3.834, 0.28, C.white, d[3]);
    txt(s, 'Your Title Here', 1.309 + i * 1.7736, 5.066, 1.475, 0.332, { fontSize: 12.5, color: C.dark, lineSpacingMultiple: 1.2 });
    txt(s, 'Vivamus quam dolor, tempor ac gravida sit amet, porta fermentum', 1.309 + i * 1.7736, 5.371, 1.447, 0.909,
      { fontSize: 10, color: C.gray65, lineSpacingMultiple: 1.2 });
  });
}

function slide19(s) {
  infoTitle(s);
  // [ring x, bar x, accent, shade, coloured run of the bar, icon]
  var groups = [
    [1.608, 1.346, C.blue, C.blueDk, 0.72, ICON.clock],
    [4.412, 4.005, C.dark, C.dark15, 0.18, ICON.gear],
    [7.215, 6.677, C.teal, C.tealDk, 0.99, ICON.plane],
    [10.018, 9.342, C.dark, C.dark15, 0.62, ICON.life]
  ];
  groups.forEach(function (g) {
    var cx = g[0] + 0.852, cy = 3.402, r = 0.852, thick = 0.207;
    ring(s, cx, cy, r, thick, g[2]);
    box(s, g[1], 4.047, 2.665, thick, g[2]);
    box(s, g[1], 4.047, g[4], thick, g[3]);
    // little shaded wedge where the loop tucks behind the bar
    poly(s, g[0] + 0.87, 3.776, 0.747, 0.271, [[1, 0], [1, 1], [0, 1]], g[3]);
    g[5](s, cx, cy, 0.23, C.dark15);
    poly(s, g[1] + 0.957, 4.614, 0.317, 0.23, [[0, 0], [1, 0], [0.5, 1]], g[2]);
    txt(s, 'Write your title', g[1] + 0.258, 5.174, 1.623, 0.337,
      { fontSize: 14, align: 'center', charSpacing: 0.3, wrap: false });
    s.addText([{ text: 'Lorem ipsum dolor amet ', options: { breakLine: true } }, { text: 'consectetur elit adipisc' }],
      { x: g[1] + 0.067, y: 5.602, w: 2.004, h: 0.591, align: 'center', valign: 'top', fontFace: F.body,
        fontSize: 11, color: C.grayTxt, lineSpacingMultiple: 1.4, wrap: false });
  });
}

function slide20(s) {
  infoTitle(s);
  var steps = [
    [0.992, C.blue, ICON.tray], [3.187, C.dark, ICON.wallet], [5.382, C.teal, ICON.bin],
    [7.577, C.dark, ICON.bank], [9.772, C.blue, ICON.cart]
  ];
  steps.forEach(function (st, i) {
    shape(s, 'chevron', st[0], 2.57, 2.91, 1.848, { fill: { color: st[1] } });
    st[2](s, st[0] + 1.545, 3.37, 0.26, C.white);
    s.addText('Step 0' + (i + 1), { x: st[0], y: 3.7, w: 2.91, h: 0.5, align: 'center', valign: 'top',
      fontFace: F.body, fontSize: 18, color: 'FEFFFF' });
    s.addText([{ text: 'Text here', options: { fontSize: 16, breakLine: true } },
               { text: 'Nam eu ex ut leo euismod tempus. Nunc at lobortis urna.', options: { fontSize: 10.5 } }],
      { x: 0.927 + i * 2.2295, y: 4.724, w: 2.121, h: 1.069, valign: 'top', fontFace: F.body, paraSpaceBefore: 12 });
  });
}

function slide21(s) {
  infoTitle(s);
  var cols = [1.689, 5.696, 9.703], numX = [1.138, 5.145, 9.152], ovX = [0.964, 4.971, 8.978];
  var items = [
    ['Welcome Message', 'About Us', 'Our Mission'],
    ['Company Timeline', 'Our Team', 'Special Portfolio'],
    ['Infographic', 'Mockup Device', 'Map']
  ];
  var nums = [['1', '2', '3'], ['4', '5', '6'], ['7', '8', '9']];
  var cs = [[C.blue, C.cyan, C.teal], [C.cyan, C.blue, C.olive], [C.lime, C.cyan, C.blue]];
  var rowY = [2.329, 3.865, 5.401];
  items.forEach(function (row, r) {
    row.forEach(function (title, c) {
      txt(s, title, cols[c], rowY[r], 2.667, 0.236, { fontSize: 14, margin: 0, wrap: false });
      txt(s, T.eiusmod, cols[c], rowY[r] + 0.303, 2.667, 0.643, { fontSize: 10.5, color: C.gray60,
        lineSpacingMultiple: 1.25, margin: 0 });
      shape(s, 'ellipse', ovX[c], rowY[r] + 0.211, 0.512, 0.512, { fill: { type: 'none' },
        line: { color: cs[r][c], width: 1 } });
      txt(s, nums[r][c], numX[c], rowY[r] + 0.299, 0.2, 0.269, { fontSize: 16, color: cs[r][c], margin: 0, wrap: false });
    });
  });
}

function slide22(s) {
  infoTitle(s);
  var cards = [
    [1.221, 60, C.dark, C.cyan, ICON.home], [4.148, 50, C.blue, C.teal, ICON.star],
    [7.147, 90, C.dark, C.green, ICON.drop], [10.085, 40, C.blue, C.lime, ICON.stack]
  ];
  var labelX = [1.049, 3.991, 7.01, 9.953];
  cards.forEach(function (cd, i) {
    var cx = cd[0] + 1.02, cy = 3.549;
    dashRing(s, cx, cy, 1.02, 0.1, Math.round(cd[1] * 12 / 100), cd[3], tint(cd[3]));
    circle(s, cx, cy, 0.789, cd[2]);
    cd[4](s, cx, cy, 0.32, C.white);
    txt(s, cd[1] + '%', cd[0] + 0.513, 4.765, 1.0, 0.572, { fontSize: 28, align: 'center', wrap: false });
    txt(s, 'Your Title Here', labelX[i], 5.357, 2.329, 0.36, { fontSize: 14, align: 'center', lineSpacingMultiple: 1.2 });
    txt(s, T.vestib, labelX[i] + 0.021, 5.661, 2.286, 0.631, { fontSize: 9, color: C.gray65, align: 'center',
      lineSpacingMultiple: 1.2 });
  });
}
/** 20% tint of an accent colour, used for the "empty" ring segments */
function tint(hex) {
  return [0, 2, 4].map(function (i) {
    var v = parseInt(hex.substr(i, 2), 16);
    return ('0' + Math.round(v + (255 - v) * 0.8).toString(16)).slice(-2).toUpperCase();
  }).join('');
}

function slide23(s) {
  infoTitle(s);
  var tags = [[0.943, C.cyan], [3.873, C.blue], [6.803, C.dark], [9.733, C.dark]];
  var w = 2.658, notch = 0.394, top = 2.509;
  tags.forEach(function (t) {
    var x = t[0];
    shape(s, 'roundRect', x, top + notch, w, 3.817 - notch, { rectRadius: 0.05, fill: { color: t[1] } });
    poly(s, x + w * 0.35, top, w * 0.3, notch + 0.02, [[0.5, 0], [1, 1], [0, 1]], t[1]);
    ICON.plane(s, x + w / 2, 3.7, 0.32, C.white);
    txt(s, 'Your Title Here', x + 0.125, 4.233, 2.357, 0.337,
      { fontSize: 14, bold: true, color: C.white, align: 'center' });
    txt(s, T.nunc, x + 0.125, 4.489, 2.357, 1.111,
      { fontSize: 10, color: C.white, align: 'center', lineSpacingMultiple: 1.5 });
  });
}

function slide24(s) {
  infoTitle(s);
  // interlocking arrows: [x, y, w, h, colour, shape, label, labelX, labelY]
  var arrows = [
    [4.343, 2.43, 3.035, 1.49, C.blue, 'leftArrow', '01', 5.448, 2.889],
    [5.932, 3.186, 2.786, 1.497, C.cyan, 'rightArrow', '02', 6.912, 3.645],
    [4.756, 3.94, 2.704, 1.492, C.dark, 'leftArrow', '03', 5.695, 4.4],
    [5.601, 4.699, 3.473, 1.49, C.dark, 'rightArrow', '04', 6.924, 5.159]
  ];
  arrows.forEach(function (a) {
    shape(s, a[5], a[0], a[1], a[2], a[3], { fill: { color: a[4] } });
    txt(s, a[6], a[7], a[8], 0.826, 0.37, { fontSize: 16, color: C.white });
  });
  // captions: [textX, textY, titleX, align, icon, iconX, iconY]
  var labels = [
    [1.506, 2.688, 2.242, 'right', ICON.plane, 3.657, 3.0],
    [1.911, 4.199, 2.647, 'right', ICON.clock, 4.065, 4.51],
    [9.835, 4.958, 9.835, 'left', ICON.home, 9.387, 5.262],
    [9.48, 3.444, 9.48, 'left', ICON.gear, 9.031, 3.752]
  ];
  labels.forEach(function (l) {
    txt(s, 'Write your title', l[2], l[1], 1.326, 0.286, { fontSize: 11, charSpacing: 0.3, align: l[3], wrap: false });
    s.addText([{ text: 'Lorem ipsum dolor, consect ', options: { breakLine: true } },
               { text: 'adipiscing elit roin cursu.' }],
      { x: l[0], y: l[1] + 0.335, w: 2.062, h: 0.547, align: l[3], valign: 'top', fontFace: F.body,
        fontSize: 10, color: C.navy, lineSpacingMultiple: 1.4, wrap: false });
    l[4](s, l[5] + 0.183, l[6] + 0.183, 0.16, C.dark15, true);
  });
}

function slide25(s) {
  infoTitle(s);
  ring(s, 5.177, 4.272, 1.181, 0.55, C.blue);
  ring(s, 8.153, 4.272, 1.184, 0.55, C.dark);
  shape(s, 'blockArc', 5.476, 1.806, 2.375, 2.375, { angleRange: [136, 44], arcThicknessRatio: 0.46,
    fill: { color: C.teal } });
  stroke(s, 6.66, 3.496, 6.66, 5.151, C.gray40, 1);
  circle(s, 6.66, 5.151, 0.05, C.gray40);
  [[4.663, 3.762, ICON.starGroup, C.blue], [7.64, 3.762, ICON.globe, C.green], [6.147, 2.477, ICON.people, C.teal]]
    .forEach(function (o) {
      circle(s, o[0] + 0.513, o[1] + 0.51, 0.513, C.white, { line: { color: C.gray40, width: 1 } });
      o[2](s, o[0] + 0.513, o[1] + 0.51, 0.18, o[3]);
    });
  txt(s, 'Insert title here', 5.875, 5.763, 1.65, 0.337,
    { fontSize: 14, bold: true, color: C.gray50, align: 'center', wrap: false });
  txt(s, 'Sed ut perspiciatis unde omnis iste natus voluptatem', 5.294, 6.196, 2.813, 0.656,
    { fontSize: 11, color: C.gray50, align: 'center', lineSpacingMultiple: 1.5 });
  txt(s, 'Insert title here', 1.741, 3.009, 1.571, 0.337, { fontSize: 14, color: C.gray50, align: 'right', wrap: false });
  txt(s, 'Sed perspiciati unde omnis iste elit voluptatem.', 0.807, 3.469, 2.506, 0.656,
    { fontSize: 11, color: C.gray50, align: 'right', lineSpacingMultiple: 1.5 });
  txt(s, 'Insert title here', 10.021, 3.009, 1.571, 0.337, { fontSize: 14, color: C.gray50, wrap: false });
  txt(s, 'Sed perspiciati unde omnis iste elit voluptatem', 10.021, 3.404, 2.506, 0.656,
    { fontSize: 11, color: C.gray50, lineSpacingMultiple: 1.5 });
}

function slide26(s) {
  infoTitle(s);
  var els = [[1.378, 70, C.blue, 'Element 01', 1.391, 1.76], [4.299, 50, C.cyan, 'Element 02', 4.301, 1.782],
             [7.219, 65, C.teal, 'Elenment 03', 7.137, 1.95], [10.14, 40, C.green, 'Element 04', 10.14, 1.786]];
  els.forEach(function (e) {
    var x = e[0], d = 1.786, cx = x + d / 2, cy = 2.61 + d / 2, r = d / 2, thick = 0.1;
    shape(s, 'blockArc', x, 2.61, d, d, { angleRange: [0, 359.9], arcThicknessRatio: thick / r,
      fill: { color: C.gray95 } });
    shape(s, 'blockArc', x, 2.61, d, d, { angleRange: [270, (270 + e[1] * 3.6) % 360],
      arcThicknessRatio: thick / r, fill: { color: e[2] } });
    txt(s, e[1] + '%', cx - 0.45, cy - 0.26, 0.9, 0.505,
      { fontSize: 24, align: 'center', charSpacing: 0.3, wrap: false });
    txt(s, e[3], e[4], 4.639, e[5], 0.303, { fontSize: 12, align: 'center', charSpacing: 0.3 });
    txt(s, 'Lorem ipsum dolor amet consectetur adipiscing elit roin cursus ligula.', x - 0.372, 5.044, 2.529, 0.843,
      { fontSize: 10.5, color: C.navy, align: 'center', lineSpacingMultiple: 1.4 });
  });
}

function slide27(s) {
  infoTitle(s);
  // pale "shadow" tiles sit behind and to the right of each dark step
  [[3.798, 3.984], [5.325, 3.58], [6.85, 3.183]].forEach(function (p) {
    shape(s, 'parallelogram', p[0], p[1], 1.451, 1.264, { fill: { color: C.gray95 } });
  });
  var steps = [[2.934, 3.984, C.dark35, 3.466, 4.752], [4.459, 3.58, C.dark35, 5.102, 4.146],
               [5.975, 3.183, C.dark, 6.738, 3.539]];
  steps.forEach(function (st, i) {
    shape(s, 'parallelogram', st[0], st[1], 1.657, 1.665, { fill: { color: st[2] } });
    circle(s, st[3] + 0.208, st[4] + 0.208, 0.208, C.white);
    s.addText(String(i + 1), { x: st[3], y: st[4], w: 0.417, h: 0.417, align: 'center', valign: 'middle',
      fontFace: F.body, fontSize: 12, margin: 0 });
  });
  // step 4 is a rising arrow
  poly(s, 7.509, 1.953, 2.103, 2.489,
    [[0.16, 1], [0.55, 0.36], [0.33, 0.24], [0.86, 0], [0.9, 0.62], [0.71, 0.51], [0.44, 1]], C.blue);
  circle(s, 8.374 + 0.208, 2.932 + 0.208, 0.208, C.white);
  s.addText('4', { x: 8.374, y: 2.932, w: 0.417, h: 0.417, align: 'center', valign: 'middle', fontFace: F.body,
    fontSize: 12, margin: 0 });
  [[2.789, 5.848], [4.863, 5.356], [6.937, 4.688], [8.947, 3.857]].forEach(function (p) {
    txt(s, 'Write your title', p[0], p[1], 1.326, 0.286, { fontSize: 11, charSpacing: 0.3, wrap: false });
    s.addText([{ text: 'Lorem ipsum dolor amet ', options: { breakLine: true } }, { text: 'consectetur adipiscing.' }],
      { x: p[0], y: p[1] + 0.334, w: 1.92, h: 0.572, valign: 'top', fontFace: F.body, fontSize: 10, color: C.navy,
        lineSpacingMultiple: 1.4, wrap: false });
  });
}

function slide28(s) {
  infoTitle(s);
  // four half-rings chained into an S: bottom arc, top arc, bottom arc, top arc
  var cells = [[2.527, C.blue, 'bottom', ICON.search], [4.524, C.cyan, 'top', ICON.wrench],
               [6.522, C.dark, 'bottom', ICON.life], [8.519, C.gray50, 'top', ICON.gear]];
  var cy = 3.75, r = 1.216, thick = 0.42;
  cells.forEach(function (c) {                              // pale rings first
    var cx = c[0] + 1.264;
    circle(s, cx, cy, r, C.gray95);
    circle(s, cx, cy, r - thick, C.white);
  });
  cells.forEach(function (c) {
    var cx = c[0] + 1.264, low = c[2] === 'bottom';
    shape(s, 'blockArc', cx - r, cy - r, 2 * r, 2 * r,
      { angleRange: low ? [0, 180] : [180, 360], arcThicknessRatio: thick / r, fill: { color: c[1] } });
    // arrow head at the leading end of the arc
    var tipX = cx + r - thick / 2;
    poly(s, tipX, low ? cy - 0.44 : cy, thick, 0.44,
      low ? [[0, 1], [1, 1], [0.5, 0]] : [[0, 0], [1, 0], [0.5, 1]], c[1]);
    c[3](s, cx, cy, 0.28, C.dark15);
  });
  s.addText([{ text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Ut pulvinar nulla in gravida convallis. Suspendisse', options: { breakLine: true } },
             { text: 'molestie pulvinar purus, nec dictum erat pulvinar viverra. Duis consequat tortor.' }],
    { x: 2.165, y: 5.7, w: 9.002, h: 0.591, align: 'center', valign: 'top', fontFace: F.body, fontSize: 11,
      color: C.navy, lineSpacingMultiple: 1.4 });
}

function slide29(s) {
  infoTitle(s);
  var cx = 6.667, cy = 4.266, knob_r = 0.37;
  // [x, y, w, h, colour] - the four quadrants overlap slightly at the seams
  var quads = [
    [4.606, 2.251, 2.061, 2.015, C.dark],
    [6.667, 2.251, 2.06, 2.015, C.navy],
    [4.606, 4.266, 2.061, 2.018, C.blue],
    [6.667, 4.266, 2.06, 2.018, C.cyan]
  ];
  quads.forEach(function (q) {
    shape(s, 'roundRect', q[0], q[1], q[2], q[3], { rectRadius: 0.07, fill: { color: q[4] },
      line: { color: C.white, width: 4.25 } });
  });
  // interlocking tabs: each is owned by the piece whose colour it carries
  knob(s, cx, 3.257, knob_r, C.dark, 'right');   // top-left reaches into top-right
  knob(s, 5.63, cy, knob_r, C.blue, 'up');       // bottom-left reaches into top-left
  knob(s, 7.7, cy, knob_r, C.navy, 'down');      // top-right reaches into bottom-right
  knob(s, cx, 5.27, knob_r, C.cyan, 'left');     // bottom-right reaches into bottom-left
  [[5.45, 3.24, ICON.monitor, 0.26, false], [7.99, 3.2, ICON.clip, 0.25, false],
   [5.36, 5.26, ICON.gearSolid, 0.26, C.blue], [7.76, 5.4, ICON.rss, 0.26, false]]
    .forEach(function (o) { o[2](s, o[0], o[1], o[3], C.white, o[4]); });
  [[9.262, 2.231, 2.653, 'left', 2.99], [9.262, 4.972, 5.394, 'left', 2.99],
   [1.769, 2.231, 2.653, 'right', 2.987], [1.769, 4.972, 5.394, 'right', 2.987]].forEach(function (r) {
    txt(s, 'Your Title Here', r[0], r[1], 2.329, 0.343,
      { fontSize: 12, color: C.dark, align: r[3], lineSpacingMultiple: 1.2 });
    txt(s, T.etiam, r[3] === 'left' ? 9.262 : 1.111, r[2], r[4], 0.586,
      { fontSize: 8, color: C.gray65, align: r[3], lineSpacingMultiple: 1.2 });
  });
}

const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30];

/* ------------------------------------------------------------------ *
 * build
 * ------------------------------------------------------------------ */
const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'AMOR', width: 13.333333333333334, height: 7.5 });
pptx.layout = 'AMOR';
pptx.title = 'Amor Templates';
pptx.theme = { headFontFace: F.head, bodyFontFace: F.body };

BUILDERS.forEach(function (build) {
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  build(slide);
});

pptx.writeFile({ fileName: path.join(__dirname, '13447ace-abd3-4d7b-93db-14f6e7f6c914_grok_final.pptx') })
  .then(function (f) { console.log('wrote ' + f); });
