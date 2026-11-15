/*
 * Orion Express - Logistics & Shipping presentation (37 slides, 13.333in x 7.5in).
 * Rebuilt from scratch with pptxgenjs: every slide is a plain builder function
 * that places shapes/text at literal inch coordinates.
 *
 * Photographs in the original deck are replaced by flat grey placeholder blocks
 * (see `photo`), and raster icons are re-drawn as native line-art (see `ICONS`).
 *
 * Run:  node <thisfile>.js   ->  writes <thisfile>.pptx next to the script.
 */
'use strict';
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette
const NAVY = '000099';   // primary brand blue
const RED = 'F70103';    // primary brand red
const BLUE = '0000B0';   // chart / infographic blue
const WHITE = 'FFFFFF';
const INK = '262626';
const GREY = '595959';   // body copy
const DKGREY = '3F3F3F';
const MIDGREY = '7F7F7F';
const SILVER = 'F2F2F2';
const PHOTO_LIGHT = 'C3C3C3'; // stand-in colour for photographs
const PHOTO_DARK = '7F7F7F';

const F = {
  head: 'Montserrat',
  semi: 'Montserrat SemiBold',
  black: 'Montserrat ExtraBold',
  body: 'Open Sans',
  pop: 'Poppins',
  popSemi: 'Poppins SemiBold',
  mul: 'Mulish',
  mulBlack: 'Mulish ExtraBold',
};

const TEXT_MARGIN = [7.2, 7.2, 3.6, 3.6]; // pt: matches the deck's 0.1in / 0.05in insets

// ---------------------------------------------------------------- helpers
function geo(o) {
  const g = { x: o.x, y: o.y, w: o.w, h: o.h };
  if (o.rotate) g.rotate = o.rotate;
  if (o.flipH) g.flipH = true;
  if (o.flipV) g.flipV = true;
  return g;
}

function paint(o) {
  const p = {};
  if (o.fill) p.fill = o.transparency ? { color: o.fill, transparency: o.transparency } : { color: o.fill };
  if (o.line) p.line = { color: o.line, width: o.lineWidth || 1 };
  // pptxgenjs rescales shadow numbers in place, so hand it a fresh copy every time
  if (o.shadow) p.shadow = Object.assign({}, o.shadow);
  return p;
}

/** Draw a shape. `poly` turns it into a closed custom path; each entry is
 *  [x, y] for a line/move or [x, y, cx1, cy1, cx2, cy2] for a cubic segment.
 *  Coordinates are inches relative to the shape's own box. */
function shape(s, kind, o) {
  const opt = Object.assign(geo(o), paint(o));
  if (o.rectRadius !== undefined) opt.rectRadius = o.rectRadius;
  if (o.angleRange) opt.angleRange = o.angleRange;
  if (o.poly) {
    opt.points = o.poly.map((p, i) => (p.length === 6
      ? { x: p[0], y: p[1], curve: { type: 'cubic', x1: p[2], y1: p[3], x2: p[4], y2: p[5] } }
      : { x: p[0], y: p[1], moveTo: i === 0 }));
    opt.points.push({ close: true });
  }
  s.addShape(o.poly ? 'custGeom' : kind, opt);
}

/** Text box. Extra keys (fontFace/fontSize/color/bold/align/lineSpacingMultiple) pass through. */
function txt(s, body, o) {
  const opt = Object.assign(geo(o), paint(o), {
    margin: TEXT_MARGIN,
    valign: o.valign || 'top',
    fontFace: o.fontFace || F.body,
    fontSize: o.fontSize || 11,
    color: o.color || INK,
    align: o.align || 'left',
    isTextBox: true,
  });
  if (o.bold) opt.bold = true;
  if (o.italic) opt.italic = true;
  if (o.lineSpacingMultiple) opt.lineSpacingMultiple = o.lineSpacingMultiple;
  if (o.shape) opt.shape = o.shape;
  if (o.rectRadius !== undefined) opt.rectRadius = o.rectRadius;
  s.addText(body, opt);
}

/** Placeholder standing in for a photograph in the original deck. */
function photo(s, o) {
  const opt = Object.assign(geo(o), { fill: { color: o.dark ? PHOTO_DARK : PHOTO_LIGHT } });
  if (o.rectRadius !== undefined) opt.rectRadius = o.rectRadius;
  if (o.shadow) opt.shadow = Object.assign({}, o.shadow);
  s.addShape(o.shape || 'rect', opt);
}

/** pptxgenjs has no gradient fill, so paint the ramp as a grid of flat tiles.
 *  stops = [[hex, position 0-100, transparency 0-100], ...]
 *  `over`     - colour underneath, so translucent stops can be pre-blended (no seams)
 *  `diagonal` - ramp runs top-left -> bottom-right instead of left -> right */
function gradient(s, o) {
  const cols = 60;
  const rows = o.diagonal ? 34 : 1;
  for (let c = 0; c < cols; c++) {
    for (let r = 0; r < rows; r++) {
      const u = (c + 0.5) / cols, v = (r + 0.5) / rows;
      const t = (o.diagonal ? (u * o.w + v * o.h) / (o.w + o.h) : u) * 100;
      let [color, transparency] = mixStops(o.stops, t);
      if (o.over) { color = blend(o.over, color, 1 - transparency / 100); transparency = 0; }
      s.addShape('rect', {
        x: o.x + (o.w * c) / cols, y: o.y + (o.h * r) / rows,
        w: o.w / cols + (o.over ? 0.008 : 0), h: o.h / rows + (o.over ? 0.008 : 0),
        fill: { color, transparency },
      });
    }
  }
}

/** Composite `top` over `base` at the given opacity, returning a hex string. */
function blend(base, top, alpha) {
  const rgb = (h) => [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
  const [r1, g1, b1] = rgb(base), [r2, g2, b2] = rgb(top);
  return [r1 + (r2 - r1) * alpha, g1 + (g2 - g1) * alpha, b1 + (b2 - b1) * alpha]
    .map((v) => Math.round(v).toString(16).padStart(2, '0')).join('').toUpperCase();
}

function mixStops(stops, t) {
  let a = stops[0], b = stops[stops.length - 1];
  for (let i = 0; i < stops.length - 1; i++) {
    if (t >= stops[i][1] && t <= stops[i + 1][1]) { a = stops[i]; b = stops[i + 1]; break; }
  }
  const span = b[1] - a[1] || 1;
  const k = Math.min(1, Math.max(0, (t - a[1]) / span));
  const lerp = (p, q) => Math.round(p + (q - p) * k);
  const rgb = (h) => [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
  const [r1, g1, b1] = rgb(a[0]), [r2, g2, b2] = rgb(b[0]);
  const hex = [lerp(r1, r2), lerp(g1, g2), lerp(b1, b2)].map((v) => v.toString(16).padStart(2, '0')).join('');
  return [hex.toUpperCase(), Math.round(a[2] + (b[2] - a[2]) * k)];
}

// ---------------------------------------------------------------- logo
/** The Orion "target + drop" mark: navy ring, red teardrop, white highlight. */
function logoMark(s, x, y, size) {
  s.addShape('ellipse', { x, y, w: size, h: size, fill: { color: NAVY } });
  s.addShape('teardrop', {
    x: x + size * 0.168, y: y + size * 0.168, w: size * 0.663, h: size * 0.663,
    fill: { color: RED }, line: { color: WHITE, width: size * 5.6 },
  });
  s.addShape('teardrop', {
    x: x + size * 0.357, y: y + size * 0.357, w: size * 0.286, h: size * 0.286, fill: { color: WHITE },
  });
}

/** Corner lock-up: the mark used as the "O" of the word "rion". */
function logoTag(s, x, y, color) {
  logoMark(s, x, y + 0.069, 0.232);
  txt(s, 'rion', { x: x + 0.155, y, w: 0.732, h: 0.37, fontFace: F.black, fontSize: 16, bold: true, color });
}

// ---------------------------------------------------------------- line-art icons
// Each glyph is drawn inside a 0..1 unit box. Primitives:
//   ['p', pts]        closed outline (append 'fill' to fill it instead of stroking)
//   ['l', pts]        open polyline
//   ['c', cx, cy, r]  circle
//   ['r', x, y, w, h] rectangle
// Points are [x, y] or [x, y, cx1, cy1, cx2, cy2] for a cubic segment.
const ICONS = {
  // freight & shipping
  box: [['p', [[.5, 0], [1, .27], [1, .73], [.5, 1], [0, .73], [0, .27]]], ['l', [[0, .27], [.5, .53], [1, .27]]], ['l', [[.5, .53], [.5, 1]]]],
  openbox: [['p', [[.5, .44], [1, .24], [1, .72], [.5, 1], [0, .72], [0, .24]]], ['l', [[.5, .44], [.5, 1]]], ['p', [[0, .24], [.26, 0], [.74, .2], [.5, .44]]], ['p', [[1, .24], [.74, 0], [.26, .2], [.5, .44]]]],
  globe: [['c', .5, .5, .49], ['l', [[.02, .5], [.98, .5]]], ['l', [[.08, .25], [.92, .25]]], ['l', [[.08, .75], [.92, .75]]], ['p', [[.5, .01], [.72, .18], [.79, .5], [.72, .82], [.5, .99], [.28, .82], [.21, .5], [.28, .18]]]],
  bag: [['p', [[.06, .3], [.74, .3], [.8, 1], [0, 1]]], ['l', [[.2, .42], [.2, .18]]], ['l', [[.6, .42], [.6, .18]]], ['l', [[.2, .18], [.4, .02], [.6, .18]]], ['l', [[.82, .28], [1, .28], [.94, .92], [.84, .92]]]],
  plane: [['p', [[.5, 0], [.57, .1], [.57, .36], [1, .64], [1, .76], [.57, .62], [.57, .85], [.71, .95], [.71, 1], [.5, .93], [.29, 1], [.29, .95], [.43, .85], [.43, .62], [0, .76], [0, .64], [.43, .36], [.43, .1]]]],
  ship: [['p', [[0, .66], [1, .62], [.84, 1], [.16, 1]]], ['l', [[.1, .66], [.1, .5], [.86, .5], [.86, .62]]], ['r', .16, .5, .7, .16], ['r', .34, .26, .28, .24], ['l', [[.4, .26], [.4, .1]]], ['c', .38, .05, .06], ['c', .3, .58, .05], ['c', .46, .58, .05], ['c', .62, .58, .05]],
  truck: [['r', .02, .2, .5, .44], ['p', [[.52, .3], [.78, .3], [.98, .5], [.98, .64], [.52, .64]]], ['c', .21, .74, .11], ['c', .79, .74, .11]],
  train: [['p', [[.2, .04], [.8, .04], [.8, .58], [.2, .58]]], ['r', .28, .12, .44, .28], ['l', [[.36, 0], [.64, 0]]], ['l', [[.5, .58], [.5, 1]]], ['l', [[.14, .68], [.86, .68]]], ['l', [[.1, .78], [.9, .78]]], ['l', [[.06, .88], [.94, .88]]], ['l', [[.02, .98], [.98, .98]]], ['c', .32, .5, .04], ['c', .68, .5, .04]],
  doc: [['p', [[0, 0], [.68, 0], [1, .32], [1, 1], [0, 1]]], ['l', [[.68, 0], [.68, .32], [1, .32]]], ['l', [[.14, .44], [.66, .44]]], ['l', [[.14, .56], [.86, .56]]], ['l', [[.14, .68], [.86, .68]]], ['l', [[.14, .8], [.86, .8]]]],
  stopwatch: [['c', .5, .6, .4], ['c', .5, .6, .3], ['l', [[.5, .6], [.5, .38]]], ['l', [[.5, .6], [.68, .72]]], ['r', .36, .04, .28, .1], ['l', [[.5, .14], [.5, .2]]], ['l', [[.78, .2], [.9, .1]]]],
  // ui / infographic
  check: [['c', .46, .54, .44], ['l', [[.24, .56], [.42, .76], [.9, .16]]]],
  store: [['p', [[.08, .3], [.92, .3], [.92, 1], [.08, 1]]], ['l', [[0, .3], [.16, .04], [.84, .04], [1, .3]]], ['l', [[.3, .3], [.34, .04]]], ['l', [[.7, .3], [.66, .04]]], ['l', [[.5, .04], [.5, .3]]], ['r', .34, .58, .32, .42]],
  search: [['c', .58, .4, .36], ['l', [[.32, .66], [.06, .94]]], ['l', [[.44, .18], [.28, .32]]]],
  person: [['c', .5, .3, .28], ['p', [[.5, .6], [.9, .88], [.98, 1], [.02, 1], [.1, .88]]]],
  home: [['l', [[.02, .5], [.5, .06], [.98, .5]]], ['p', [[.14, .46], [.86, .46], [.86, .98], [.14, .98]]], ['r', .4, .68, .2, .3]],
  clock: [['c', .5, .5, .48], ['l', [[.5, .5], [.5, .24]]], ['l', [[.5, .5], [.7, .62]]]],
  gear: [['c', .5, .5, .42], ['c', .5, .5, .17], ['l', [[.5, .02], [.5, .14]]], ['l', [[.5, .86], [.5, .98]]], ['l', [[.02, .5], [.14, .5]]], ['l', [[.86, .5], [.98, .5]]], ['l', [[.16, .16], [.25, .25]]], ['l', [[.75, .75], [.84, .84]]], ['l', [[.84, .16], [.75, .25]]], ['l', [[.25, .75], [.16, .84]]]],
  bolt: [['p', [[.62, 0], [.06, .54], [.4, .54], [.34, 1], [.94, .44], [.56, .44]]]],
  phone: [['p', [[.06, 0], [.94, 0], [.94, 1], [.06, 1]]], ['r', .16, .1, .68, .78]],
  pin: [['p', [[.5, 1], [.16, .62, .3, .92, .2, .74], [.5, 0, .12, .48, .16, .19], [.84, .62, .84, .19, .88, .48], [.5, 1, .8, .74, .7, .92]]], ['c', .5, .38, .17]],
  play: [['p', [[.04, .02], [.98, .5], [.04, .98]]]],
  apple: [['p', [[.5, .26], [.86, .34], [1, .64], [.76, 1], [.5, .92], [.24, 1], [0, .64], [.14, .34]], 'fill'], ['p', [[.5, .24], [.44, 0], [.72, .08]], 'fill']],
  paperplane: [['p', [[1, 0], [0, .46], [.42, .6], [.54, 1], [.66, .52]]], ['l', [[.42, .6], [1, 0]]], ['l', [[.06, .78], [.2, .66]]], ['l', [[.1, .96], [.3, .8]]]],
};

function icon(s, name, o) {
  const color = o.color || RED;
  const px = (u) => u * o.w;   // custGeom points are relative to the shape box
  const py = (u) => u * o.h;
  const stroke = { color, width: Math.max(0.75, Math.min(o.w, o.h) * 3.2) };
  for (const prim of ICONS[name] || []) {
    if (prim[0] === 'c') {
      s.addShape('ellipse', { x: o.x + (prim[1] - prim[3]) * o.w, y: o.y + (prim[2] - prim[3]) * o.h, w: o.w * prim[3] * 2, h: o.h * prim[3] * 2, line: stroke });
    } else if (prim[0] === 'r') {
      s.addShape('rect', { x: o.x + prim[1] * o.w, y: o.y + prim[2] * o.h, w: o.w * prim[3], h: o.h * prim[4], line: stroke });
    } else {
      const filled = prim[2] === 'fill';
      const pts = prim[1].map((p, i) => (p.length === 6
        ? { x: px(p[0]), y: py(p[1]), curve: { type: 'cubic', x1: px(p[2]), y1: py(p[3]), x2: px(p[4]), y2: py(p[5]) } }
        : { x: px(p[0]), y: py(p[1]), moveTo: i === 0 }));
      if (prim[0] === 'p') pts.push({ close: true });
      s.addShape('custGeom', Object.assign({ x: o.x, y: o.y, w: o.w, h: o.h, points: pts },
        filled ? { fill: { color } } : { line: stroke }));
    }
  }
}

/** Phone / desktop mockup placeholder (the original deck used photo renders). */
function device(s, kind, o) {
  if (kind === 'phone') {
    // rotation is about the shape centre, so both layers share the same box centre
    shape(s, 'roundRect', { x: o.x, y: o.y, w: o.w, h: o.h, rotate: o.rotate, fill: '3F3F3F', rectRadius: o.w * 0.1 });
    shape(s, 'roundRect', { x: o.x + o.w * 0.018, y: o.y + o.h * 0.009, w: o.w * 0.964, h: o.h * 0.982, rotate: o.rotate, fill: INK, rectRadius: o.w * 0.095 });
    shape(s, 'roundRect', { x: o.x + o.w * 0.055, y: o.y + o.h * 0.028, w: o.w * 0.89, h: o.h * 0.944, rotate: o.rotate, fill: PHOTO_LIGHT, rectRadius: o.w * 0.075 });
  } else {
    const f = (fx, fy, fw, fh, fill) => shape(s, 'rect', { x: o.x + o.w * fx, y: o.y + o.h * fy, w: o.w * fw, h: o.h * fh, fill });
    f(0.019, 0.020, 0.962, 0.679, INK);          // bezel
    f(0.040, 0.053, 0.910, 0.623, PHOTO_LIGHT);  // screen
    f(0.019, 0.699, 0.962, 0.104, 'D9D9D9');     // chin
    shape(s, 'trapezoid', { x: o.x + o.w * 0.34, y: o.y + o.h * 0.803, w: o.w * 0.32, h: o.h * 0.13, fill: 'D9D9D9' }); // neck
    f(0.31, 0.928, 0.38, 0.036, 'BFBFBF');       // base
  }
}
const L1 = 'Turpis egestas pretium aenean pharetra magna ac placerat. Cras sed felis eget velit aliquet sagittis id consectetur purus. Habitant morbi tristique senectus et netus et. ';
const L2 = 'Turpis egestas pretium aenean pharetra magna ac placerat. Cras sed felis eget velit aliquet sagittis id consectetur purus. Habitant morbi tristique senectus et netus et. Blandit turpis cursus in hac habitasse. Eu lobortis elementum nibh tellus molestie nunc non blandit massa bibendum arcu vitae elementum.';
const L3 = 'Turpis egestas pretium aenean pharetra magna ac placerat. Cras sed felis eget velit aliquet sagittis id consectetur purus habitant morbi tristique.';
const L4 = 'Turpis egestas pretium aenean pharetra magna ac placerat. Cras sed felis eget velit aliquet sagittis id consectetur purus. Habitant morbi tristique senectus et netus et. Blandit turpis cursus in hac habitasse. Eu lobortis elementum nibh tellus molestie.';
const L5 = 'Bibendum arcu vitae elementum curabitur vitae nunc sed id diam maecenas ultricies .';
const L6 = 'Lorem dolor sed viverra ipsum nunc aliquet bibendum enim.';
const L7 = 'Blandit turpis cursus in hac habitasse. Eu lobortis elementum nibh tellus molestie nunc non blandit massa bibendum arcu vitae elementum.';
const L8 = 'Turpis egestas pretium aenean pharetra magna ac placerat cras sed felis eget velit aliquet sagittis.';
const L9 = 'Tortor id aliquet lectus proin nibh nisl. Euismod lacinia at quis risus sed vulputate odio. Ut enim blandit volutpat maecenas volutpat. Amet aliquam id diam maecenas ultricies mi. ';
const L10 = 'Turpis egestas pretium aenean pharetra magna. Blandit turpis cursus in hac habitasse. ';
const L11 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit sed do eiusmod tempor.';
const L12 = 'Habitant morbi tristique senectus et netus et. Blandit turpis cursus in hac habitasse. Eu lobortis elementum nibh tellus molestie nunc non blandit massa.';
const L13 = 'Turpis egestas pretium aenean pharetra magna ac placerat. Cras sed felis eget velit aliquet sagittis.';
const L14 = 'Habitant morbi tristique senectus et netus et. Blandit turpis cursus in hac habitasse. Eu lobortis elementum nibh tellus molestie.';
const L15 = 'Turpis egestas pretium aenean pharetra magna. Blandit turpis cursus in hac habitasse. Eu lobortis elementum nibh tellus molestie.';
const L16 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.';
const L17 = 'Bibendum arcu vitae elementum curabitur vitae nunc sed id diam.';
const L18 = 'PLACEHOLDER';
const L19 = 'LOGISTIC & SHIPPING SERVICE  PRESENTATION TEMPLATE ';
const L20 = 'Making the world a better place one delivery at a time.';

// Reusable text styles (font face / size / colour / alignment / line spacing)
const T = {
  body10WhiteJL15: { fontFace: F.body, fontSize: 10, color: WHITE, align: 'justify', lineSpacingMultiple: 1.5 },
  body11GreyCL15: { fontFace: F.body, fontSize: 11, color: GREY, align: 'center', lineSpacingMultiple: 1.5 },
  body11GreyJL15: { fontFace: F.body, fontSize: 11, color: GREY, align: 'justify', lineSpacingMultiple: 1.5 },
  body11GreyL15: { fontFace: F.body, fontSize: 11, color: GREY, lineSpacingMultiple: 1.5 },
  body11WhiteCL15: { fontFace: F.body, fontSize: 11, color: WHITE, align: 'center', lineSpacingMultiple: 1.5 },
  body11WhiteJL15: { fontFace: F.body, fontSize: 11, color: WHITE, align: 'justify', lineSpacingMultiple: 1.5 },
  head10_5WhiteBC: { fontFace: F.head, fontSize: 10.5, color: WHITE, bold: true, align: 'center' },
  head12NavyB: { fontFace: F.head, fontSize: 12, color: NAVY, bold: true },
  head12NavyBC: { fontFace: F.head, fontSize: 12, color: NAVY, bold: true, align: 'center' },
  head12WhiteB: { fontFace: F.head, fontSize: 12, color: WHITE, bold: true },
  head12WhiteBC: { fontFace: F.head, fontSize: 12, color: WHITE, bold: true, align: 'center' },
  head12WhiteC: { fontFace: F.head, fontSize: 12, color: WHITE, align: 'center' },
  head14InkB: { fontFace: F.head, fontSize: 14, color: INK, bold: true },
  head14InkBC: { fontFace: F.head, fontSize: 14, color: INK, bold: true, align: 'center' },
  head14WhiteB: { fontFace: F.head, fontSize: 14, color: WHITE, bold: true },
  head14WhiteBC: { fontFace: F.head, fontSize: 14, color: WHITE, bold: true, align: 'center' },
  head22WhiteBC: { fontFace: F.head, fontSize: 22, color: WHITE, bold: true, align: 'center' },
  head24NavyB: { fontFace: F.head, fontSize: 24, color: NAVY, bold: true },
  head24WhiteBC: { fontFace: F.head, fontSize: 24, color: WHITE, bold: true, align: 'center' },
  head28NavyBC: { fontFace: F.head, fontSize: 28, color: NAVY, bold: true, align: 'center' },
  head9InkC: { fontFace: F.head, fontSize: 9, color: INK, align: 'center' },
  pop12White: { fontFace: F.pop, fontSize: 12, color: WHITE },
  popSemi18WhiteBC: { fontFace: F.popSemi, fontSize: 18, color: WHITE, bold: true, align: 'center' },
  semi11Red: { fontFace: F.semi, fontSize: 11, color: RED },
  semi11White: { fontFace: F.semi, fontSize: 11, color: WHITE },
  semi12WhiteBC: { fontFace: F.semi, fontSize: 12, color: WHITE, bold: true, align: 'center' },
  semi16NavyBC: { fontFace: F.semi, fontSize: 16, color: NAVY, bold: true, align: 'center' },
};

// Reusable drop shadows
const SH = {
  sh1: { type: 'outer', color: '000000', blur: 25, offset: 3, angle: 90, opacity: 0.098 },
  sh2: { type: 'outer', color: INK, blur: 25, offset: 3, angle: 90, opacity: 0.098 },
  sh3: { type: 'outer', color: '000000', blur: 9, offset: 4, angle: 90, opacity: 0.067 },
  sh4: { type: 'outer', color: '222A35', blur: 25, offset: 3, angle: 90, opacity: 0.098 },
  sh5: { type: 'outer', color: INK, blur: 20, offset: 3, angle: 90, opacity: 0.118 },
  sh6: { type: 'outer', color: '000000', blur: 17, offset: 0, angle: 0, opacity: 0.047 },
  sh7: { type: 'outer', color: '000000', blur: 28, offset: 0, angle: 0, opacity: 0.118 },
};

function slide01(s) {
  photo(s, { x: 0, y: -0.007, w: 13.333, h: 7.513 });
  shape(s, 'rect', { x: 0, y: -0.007, w: 13.333, h: 7.507, fill: WHITE, transparency: 12.2 });
  logoMark(s, 6.133, 2.303, 1.068);
  txt(s, L19, { x: 3.249, y: 4.894, w: 6.835, h: 0.303, fontFace: F.semi, fontSize: 12, color: DKGREY, bold: true, align: 'center' });
  txt(s, [{ text: 'Orion', options: { color: NAVY } }, { text: ' ', options: { color: INK } }, { text: 'Express', options: { color: RED } }], { x: 2.004, y: 3.582, w: 9.313, h: 1.313, fontFace: F.black, fontSize: 72, bold: true, align: 'center' });
}

function slide02(s) {
  s.background = { color: RED };
  shape(s, 'custGeom', { x: 8.909, y: 0, w: 4.42, h: 5.258, poly: [[0.291, 0], [4.42, 0], [4.42, 5.215], [4.239, 5.238], [3.845, 5.258, 4.109, 5.251, 3.978, 5.258], [0, 1.447, 1.722, 5.258, 0, 3.552], [0.233, 0.137, 0, 0.986, 0.082, 0.545]], fill: NAVY });
  photo(s, { x: 6.667, y: 0, w: 6.667, h: 7.513 });
  logoMark(s, 0.924, 5.263, 0.863);
  txt(s, L19, { x: 1.778, y: 6.199, w: 5.391, h: 0.303, fontFace: F.semi, fontSize: 12, color: WHITE, bold: true });
  txt(s, 'rion Express', { x: 1.778, y: 5.19, w: 5.781, h: 1.027, fontFace: F.black, fontSize: 55, color: WHITE, bold: true });
  txt(s, L20, { x: 0.833, y: 1.321, w: 3.647, h: 2.794, fontFace: F.head, fontSize: 32, color: WHITE, bold: true });
  shape(s, 'line', { x: 0.94, y: 1.027, w: 0.702, h: 0, line: WHITE, lineWidth: 0.75 });
}

function slide03(s) {
  shape(s, 'rect', { x: -2.815, y: 2.815, w: 6.333, h: 0.703, rotate: -90, fill: RED });
  shape(s, 'rect', { x: -0.232, y: 6.565, w: 1.167, h: 0.703, rotate: -90, fill: NAVY });
  logoTag(s, 12.266, 0.181, NAVY);
  txt(s, 'WELCOME MESSAGE', { x: 7.425, y: 1.522, w: 4.854, h: 0.286, ...T.semi11Red });
  txt(s, L2, { x: 7.425, y: 3.521, w: 4.854, h: 1.46, ...T.body11GreyJL15 });
  txt(s, 'We are the largest shipping service company in America.', { x: 7.425, y: 1.916, w: 4.854, h: 1.313, ...T.head24NavyB });
  txt(s, L1, { x: 7.425, y: 5.061, w: 4.854, h: 0.904, ...T.body11GreyJL15 });
  txt(s, 'GLOBAL SHIPPING SERVICE', { x: -2.075, y: 2.506, w: 4.854, h: 0.303, rotate: -90, fontFace: F.head, fontSize: 12, color: WHITE, bold: true, align: 'right' });
  photo(s, { x: 0.703, y: 0, w: 5.667, h: 7.5 });
}

function slide04(s) {
  shape(s, 'rect', { x: 8.987, y: 0.707, w: 4.346, h: 6.793, fill: NAVY });
  photo(s, { x: 7.796, y: 0.707, w: 4.346, h: 6.086 });
  shape(s, 'rect', { x: 5.106, y: 2.897, w: 1.438, h: 6.353, rotate: 90, fill: RED, shadow: SH.sh1 });
  txt(s, 'EMPLOYEE', { x: 2.98, y: 6.183, w: 1.909, h: 0.303, ...T.head12WhiteBC });
  txt(s, '5000', { x: 3.261, y: 5.65, w: 1.348, h: 0.505, ...T.head24WhiteBC });
  txt(s, 'CUSTOMER', { x: 4.864, y: 6.183, w: 1.909, h: 0.303, ...T.head12WhiteBC });
  txt(s, '5M+', { x: 5.313, y: 5.65, w: 1.011, h: 0.505, ...T.head24WhiteBC });
  txt(s, 'AWARD', { x: 6.76, y: 6.183, w: 1.909, h: 0.303, ...T.head12WhiteBC });
  txt(s, '115', { x: 7.196, y: 5.65, w: 1.011, h: 0.505, ...T.head24WhiteBC });
  shape(s, 'roundRect', { x: 4.852, y: 5.744, w: 0.05, h: 0.632, fill: SILVER, rectRadius: 0.025 });
  shape(s, 'roundRect', { x: 6.718, y: 5.744, w: 0.05, h: 0.632, fill: SILVER, rectRadius: 0.025 });
  txt(s, 'ABOUT US', { x: 1.008, y: 1.285, w: 4.854, h: 0.286, ...T.semi11Red });
  txt(s, 'Turpis egestas pretium aenean pharetra magna ac placerat. Cras sed felis eget velit aliquet sagittis id consectetur purus. Habitant morbi tristique senectus et netus et. Blandit turpis cursus in hac habitasse. Eu lobortis elementum nibh tellus molestie nunc non blandit massa. Bibendum arcu vitae elementum curabitur vitae nunc sed. Id diam maecenas ultricies mi eget mauris. ', { x: 1.008, y: 3.248, w: 5.78, h: 1.46, ...T.body11GreyJL15 });
  txt(s, 'We are committed to providing quality in each and every delivery.', { x: 1.008, y: 1.68, w: 5.475, h: 1.313, ...T.head24NavyB });
  logoTag(s, 0.21, 0.196, NAVY);
}

function slide05(s) {
  shape(s, 'rect', { x: 0, y: 5.125, w: 13.333, h: 2.375, fill: RED });
  logoMark(s, 1.64, 0.672, 0.776);
  txt(s, [{ text: 'Orion', options: { color: NAVY } }, { text: ' ', options: { color: INK } }, { text: 'Express', options: { color: RED } }], { x: 0.422, y: 1.595, w: 3.212, h: 0.404, fontFace: F.black, fontSize: 18, bold: true, align: 'center' });
  txt(s, L12, { x: 4.247, y: 6.1, w: 3.833, h: 0.83, ...T.body10WhiteJL15 });
  txt(s, 'OUR VISION', { x: 4.247, y: 5.667, w: 1.909, h: 0.337, ...T.head14WhiteB });
  txt(s, L12, { x: 8.886, y: 6.1, w: 3.833, h: 0.83, ...T.body10WhiteJL15 });
  txt(s, 'OUR MISSION', { x: 8.886, y: 5.667, w: 1.909, h: 0.337, ...T.head14WhiteB });
  txt(s, 'ABOUT US', { x: 7.855, y: 1.059, w: 4.854, h: 0.286, ...T.semi11Red });
  txt(s, L2, { x: 7.855, y: 2.64, w: 4.854, h: 1.46, ...T.body11GreyJL15 });
  txt(s, 'We take your business wherever it goes.', { x: 7.855, y: 1.453, w: 4.662, h: 0.942, fontFace: F.head, fontSize: 25, color: NAVY, bold: true });
  photo(s, { x: 0.422, y: 2.375, w: 3.212, h: 5.125 });
  photo(s, { x: 4.019, y: 0, w: 3.212, h: 5.125 });
}

function slide06(s) {
  photo(s, { x: 0, y: -0.013, w: 7.688, h: 7.513 });
  shape(s, 'rect', { x: 0, y: -0.013, w: 3.844, h: 7.513, fill: RED, transparency: 17.3 });
  txt(s, 'ABOUT US', { x: 8.134, y: 1.448, w: 4.493, h: 0.286, ...T.semi11Red });
  txt(s, L1, { x: 8.134, y: 5.147, w: 4.753, h: 0.904, ...T.body11GreyJL15 });
  txt(s, 'We are expanding to various countries around the world.', { x: 8.134, y: 1.843, w: 4.753, h: 1.313, ...T.head24NavyB });
  txt(s, 'Turpis egestas pretium aenean pharetra magna ac placerat. Cras sed felis eget velit aliquet sagittis id consectetur purus. Habitant morbi tristique senectus et netus et. Blandit turpis cursus in hac habitasse. Eu lobortis elementum nibh tellus molestie nunc non blandit massa.', { x: 8.134, y: 3.509, w: 4.753, h: 1.46, ...T.body11GreyJL15 });
  txt(s, L3, { x: 0.32, y: 2.239, w: 3.203, h: 1.182, ...T.body11WhiteCL15 });
  txt(s, '15M DELIVERED PACKAGE ', { x: 0.32, y: 1.875, w: 3.203, h: 0.32, fontFace: F.head, fontSize: 13, color: WHITE, bold: true, align: 'center' });
  txt(s, L3, { x: 0.32, y: 5.225, w: 3.203, h: 1.182, ...T.body11WhiteCL15 });
  txt(s, '120 COUNTRIES', { x: 0.32, y: 4.86, w: 3.203, h: 0.32, fontFace: F.head, fontSize: 13, color: WHITE, bold: true, align: 'center' });
  icon(s, 'globe', { x: 1.642, y: 4.138, w: 0.56, h: 0.56, color: WHITE });
  icon(s, 'box', { x: 1.642, y: 1.059, w: 0.56, h: 0.653, color: WHITE });
  logoTag(s, 12.266, 0.181, NAVY);
}

function slide07(s) {
  shape(s, 'rect', { x: 11.794, y: 0, w: 1.54, h: 0.703, fill: RED });
  shape(s, 'rect', { x: 0, y: 0, w: 11.794, h: 0.703, fill: NAVY });
  photo(s, { x: 0, y: 0.703, w: 13.333, h: 4.214 });
  gradient(s, { x: 0, y: 0.703, w: 7.14, h: 4.214, over: 'C3C3C3', stops: [[WHITE, 0, 8.2], [WHITE, 37, 8.2], [WHITE, 67, 22.4], [WHITE, 100, 100]] });
  txt(s, 'DOMESTIC SHIPPING ', { x: 1.824, y: 5.65, w: 2.271, h: 0.303, ...T.head12NavyB });
  txt(s, L5, { x: 1.824, y: 5.969, w: 2.703, h: 0.904, ...T.body11GreyJL15 });
  txt(s, 'INTERNATIONAL SHIPPING ', { x: 5.686, y: 5.65, w: 2.703, h: 0.303, ...T.head12NavyB });
  txt(s, 'ECOMMERCE SHIPPING', { x: 9.549, y: 5.65, w: 2.271, h: 0.303, ...T.head12NavyB });
  icon(s, 'box', { x: 1.237, y: 5.71, w: 0.433, h: 0.505, color: RED });
  txt(s, L5, { x: 5.686, y: 5.969, w: 2.703, h: 0.904, ...T.body11GreyJL15 });
  txt(s, L5, { x: 9.549, y: 5.969, w: 2.703, h: 0.904, ...T.body11GreyJL15 });
  txt(s, 'A delivery service you can depend on.', { x: 0.732, y: 2.308, w: 5.155, h: 1.043, fontFace: F.head, fontSize: 28, color: NAVY, bold: true });
  txt(s, L1, { x: 0.732, y: 3.53, w: 5.155, h: 0.904, ...T.body11GreyJL15 });
  txt(s, 'ABOUT US', { x: 0.732, y: 1.963, w: 2.408, h: 0.286, ...T.semi11Red });
  icon(s, 'bag', { x: 8.905, y: 5.71, w: 0.504, h: 0.505, color: RED });
  icon(s, 'globe', { x: 5.095, y: 5.701, w: 0.505, h: 0.505, color: RED });
  logoTag(s, 12.168, 0.171, WHITE);
}

function slide08(s) {
  shape(s, 'rect', { x: 8.875, y: 0, w: 4.458, h: 7.5, fill: RED });
  txt(s, 'ABOUT US', { x: 1.261, y: 1.485, w: 3.667, h: 0.286, ...T.semi11Red });
  txt(s, L2, { x: 1.261, y: 3.067, w: 4.854, h: 1.46, ...T.body11GreyJL15 });
  txt(s, 'We deliver in rain and shine, on land, sea and air.', { x: 1.261, y: 1.88, w: 4.854, h: 0.909, ...T.head24NavyB });
  txt(s, 'ORION CARGO', { x: 1.261, y: 4.807, w: 2.703, h: 0.303, ...T.head12NavyB });
  txt(s, L1, { x: 1.261, y: 5.11, w: 4.854, h: 0.904, ...T.body11GreyJL15 });
  logoTag(s, 0.269, 0.271, NAVY);
  photo(s, { x: 7.375, y: 0, w: 5.094, h: 6.812 });
}

function slide09(s) {
  shape(s, 'rect', { x: 8.688, y: 4.227, w: 4.646, h: 3.273, fill: RED });
  txt(s, L3, { x: 9.37, y: 5.843, w: 3.28, h: 1.182, ...T.body11WhiteCL15 });
  txt(s, 'PACKAGE ON TIME', { x: 9.591, y: 5.397, w: 2.84, h: 0.303, ...T.head12WhiteBC });
  txt(s, '90%', { x: 10.056, y: 4.702, w: 1.909, h: 0.64, fontFace: F.head, fontSize: 32, color: WHITE, bold: true, align: 'center' });
  txt(s, 'No other courier company makes deliveries as easy as we do.', { x: 0.97, y: 5.004, w: 6.747, h: 0.909, ...T.head24NavyB });
  txt(s, 'Turpis egestas pretium aenean pharetra magna ac placerat. Cras sed felis eget velit aliquet sagittis id consectetur purus. Habitant morbi tristique senectus et netus et. Blandit turpis cursus in hac habitasse eu lobortis elementum nibh tellus molestie nunc.', { x: 0.97, y: 6.164, w: 6.747, h: 0.904, ...T.body11GreyJL15 });
  txt(s, 'ABOUT US', { x: 0.97, y: 4.659, w: 2.408, h: 0.286, ...T.semi11Red });
  txt(s, L3, { x: 9.398, y: 2.308, w: 3.28, h: 1.182, ...T.body11GreyCL15 });
  txt(s, 'SAFE & RELIABLE LOGISTICS', { x: 9.618, y: 1.862, w: 2.84, h: 0.303, ...T.head12NavyBC });
  icon(s, 'openbox', { x: 10.546, y: 0.919, w: 0.998, h: 0.706, color: RED });
  photo(s, { x: 0, y: 0, w: 8.688, h: 4.227 });
}

function slide10(s) {
  shape(s, 'rect', { x: 0, y: 0, w: 5.211, h: 7.5, fill: NAVY });
  txt(s, 'www.orion.com', { x: 0.483, y: 5.738, w: 2.267, h: 0.377, fontFace: F.head, fontSize: 16, color: WHITE, align: 'center' });
  txt(s, 'ABOUT US', { x: 7.729, y: 0.888, w: 3.437, h: 0.286, ...T.semi11Red });
  txt(s, L1, { x: 7.729, y: 2.417, w: 4.854, h: 0.904, ...T.body11GreyJL15 });
  txt(s, 'Why you should choose Orion Express?', { x: 7.729, y: 1.283, w: 4.854, h: 0.909, ...T.head24NavyB });
  icon(s, 'check', { x: 7.839, y: 3.568, w: 0.25, h: 0.247, color: RED });
  txt(s, 'SAFE & RELIABLE', { x: 8.105, y: 3.539, w: 2.271, h: 0.303, ...T.head12NavyB });
  txt(s, L5, { x: 8.105, y: 3.858, w: 3.709, h: 0.627, ...T.body11GreyJL15 });
  icon(s, 'check', { x: 7.839, y: 4.631, w: 0.25, h: 0.247, color: RED });
  txt(s, '100% SATISFACTION GUARANTEE', { x: 8.105, y: 4.603, w: 3.334, h: 0.303, ...T.head12NavyB });
  txt(s, L5, { x: 8.105, y: 4.922, w: 3.709, h: 0.627, ...T.body11GreyJL15 });
  icon(s, 'check', { x: 7.823, y: 5.695, w: 0.25, h: 0.247, color: RED });
  txt(s, 'FAST AND EFFICIENT DELIVERY', { x: 8.089, y: 5.666, w: 3.334, h: 0.303, ...T.head12NavyB });
  txt(s, L5, { x: 8.089, y: 5.985, w: 3.709, h: 0.627, ...T.body11GreyJL15 });
  logoTag(s, 12.266, 0.181, NAVY);
  photo(s, { x: 0.276, y: 1.993, w: 2.682, h: 3.514 });
  photo(s, { x: 3.269, y: 0.902, w: 3.709, h: 5.697 });
}

function slide11(s) {
  shape(s, 'rect', { x: 4.524, y: 3.069, w: 4.285, h: 2.92, rotate: 90, fill: WHITE, shadow: SH.sh2 });
  shape(s, 'rect', { x: 7.894, y: 3.069, w: 4.285, h: 2.92, rotate: 90, fill: WHITE, shadow: SH.sh2 });
  shape(s, 'rect', { x: 1.155, y: 3.069, w: 4.285, h: 2.92, rotate: 90, fill: RED, shadow: { type: 'outer', color: 'FF2525', blur: 25, offset: 3, angle: 90, opacity: 0.098 } });
  txt(s, 'OUR TEAM', { x: 4.948, y: 0.707, w: 3.437, h: 0.286, fontFace: F.semi, fontSize: 11, color: RED, align: 'center' });
  txt(s, 'Meet Our Leadership', { x: 4.24, y: 1.064, w: 4.854, h: 0.572, ...T.head28NavyBC });
  txt(s, 'Head of Production', { x: 5.689, y: 4.947, w: 1.955, h: 0.286, fontFace: F.head, fontSize: 11, color: GREY, align: 'center' });
  txt(s, 'NOEL P. SULLIVAN', { x: 5.504, y: 4.61, w: 2.326, h: 0.337, fontFace: F.head, fontSize: 14, color: NAVY, bold: true, align: 'center' });
  txt(s, 'Head of Marketing', { x: 2.32, y: 4.949, w: 1.955, h: 0.286, fontFace: F.head, fontSize: 11, color: WHITE, align: 'center' });
  txt(s, 'ALFEN ALEXANDER', { x: 2.134, y: 4.612, w: 2.326, h: 0.337, ...T.head14WhiteBC });
  txt(s, 'Head of Management', { x: 8.966, y: 4.947, w: 2.14, h: 0.286, fontFace: F.head, fontSize: 11, color: GREY, align: 'center' });
  txt(s, 'DERINA C. DERIN', { x: 8.873, y: 4.61, w: 2.326, h: 0.337, fontFace: F.head, fontSize: 14, color: NAVY, bold: true, align: 'center' });
  txt(s, L16, { x: 2.012, y: 5.405, w: 2.57, h: 0.904, ...T.body11WhiteCL15 });
  txt(s, L16, { x: 5.381, y: 5.405, w: 2.57, h: 0.904, ...T.body11GreyCL15 });
  txt(s, L16, { x: 8.751, y: 5.405, w: 2.57, h: 0.904, ...T.body11GreyCL15 });
  logoTag(s, 12.266, 0.181, NAVY);
  photo(s, { x: 2.626, y: 2.8, w: 1.344, h: 1.344, shape: 'ellipse' });
  photo(s, { x: 5.995, y: 2.8, w: 1.344, h: 1.344, shape: 'ellipse' });
  photo(s, { x: 9.364, y: 2.8, w: 1.344, h: 1.344, shape: 'ellipse' });
}

function slide12(s) {
  shape(s, 'rect', { x: 0, y: 0, w: 5.208, h: 6.633, fill: RED });
  txt(s, 'Ruddy C. Paulsen', { x: 6.729, y: 1.687, w: 4.707, h: 0.64, fontFace: F.head, fontSize: 32, color: NAVY, bold: true });
  txt(s, 'OUR TEAM', { x: 6.729, y: 1.338, w: 1.862, h: 0.286, ...T.semi11Red });
  txt(s, 'Head of Marketing', { x: 6.729, y: 2.391, w: 2.429, h: 0.337, fontFace: F.head, fontSize: 14, color: GREY });
  txt(s, L2, { x: 6.729, y: 3.123, w: 5.083, h: 1.46, ...T.body11GreyJL15 });
  txt(s, 'Your Title Here', { x: 6.729, y: 4.899, w: 2.079, h: 0.286, fontFace: F.semi, fontSize: 11, color: GREY });
  txt(s, '80%', { x: 11.274, y: 4.899, w: 0.539, h: 0.286, fontFace: F.semi, fontSize: 11, color: GREY, align: 'right' });
  txt(s, [{ text: 'Your Title Here', options: { fontFace: F.semi, fontSize: 11, color: GREY, align: 'left', breakLine: true } }, { text: '', options: { align: 'left', breakLine: true } }], { x: 6.729, y: 5.66, w: 2.079, h: 0.471 });
  txt(s, '90%', { x: 11.277, y: 5.66, w: 0.535, h: 0.286, fontFace: F.semi, fontSize: 11, color: GREY, align: 'right' });
  txt(s, 'www.orion.com', { x: -0.619, y: 4.997, w: 2.267, h: 0.377, rotate: -90, fontFace: F.head, fontSize: 16, color: WHITE, align: 'center' });
  logoTag(s, 12.266, 0.181, NAVY);
  photo(s, { x: 1.1, y: 0.867, w: 4.109, h: 5.767 });
}

function slide13(s) {
  shape(s, 'rect', { x: 0, y: 0, w: 13.333, h: 4.447, flipH: true, fill: RED });
  photo(s, { x: 0.693, y: 2.374, w: 2.858, h: 3.458 });
  photo(s, { x: 3.723, y: 2.374, w: 2.858, h: 3.458 });
  photo(s, { x: 6.752, y: 2.374, w: 2.858, h: 3.458 });
  photo(s, { x: 9.782, y: 2.374, w: 2.858, h: 3.458 });
  shape(s, 'rect', { x: 0.693, y: 5.113, w: 2.858, h: 0.719, fill: NAVY, shadow: SH.sh3 });
  shape(s, 'rect', { x: 3.723, y: 5.113, w: 2.858, h: 0.719, fill: NAVY, shadow: SH.sh3 });
  shape(s, 'rect', { x: 6.752, y: 5.113, w: 2.858, h: 0.719, fill: NAVY, shadow: SH.sh3 });
  shape(s, 'rect', { x: 9.782, y: 5.113, w: 2.858, h: 0.719, fill: NAVY, shadow: SH.sh3 });
  txt(s, 'LEONY ANGELISTA ', { x: 3.966, y: 5.321, w: 2.372, h: 0.303, ...T.head12WhiteBC });
  txt(s, 'NIKO L. DARMIAN', { x: 6.995, y: 5.321, w: 2.372, h: 0.303, ...T.head12WhiteBC });
  txt(s, 'DERINA C. DERIN', { x: 10.025, y: 5.321, w: 2.372, h: 0.303, ...T.head12WhiteBC });
  txt(s, 'MICHAEL N. SALOSA', { x: 0.936, y: 5.321, w: 2.372, h: 0.303, ...T.head12WhiteBC });
  txt(s, 'OUR TEAM', { x: 4.948, y: 0.707, w: 3.437, h: 0.286, fontFace: F.semi, fontSize: 11, color: WHITE, align: 'center' });
  txt(s, 'Meet Our Best Team', { x: 4.24, y: 1.064, w: 4.854, h: 0.572, fontFace: F.head, fontSize: 28, color: WHITE, bold: true, align: 'center' });
  txt(s, L1, { x: 2.436, y: 6.29, w: 8.462, h: 0.627, ...T.body11GreyCL15 });
}

function slide14(s) {
  photo(s, { x: 0, y: 0.703, w: 13.333, h: 6.094, dark: true });
  gradient(s, { x: 0, y: 0.703, w: 13.333, h: 6.094, diagonal: true, over: MIDGREY, stops: [[WHITE, 0, 17.3], ['8080CC', 51, 17.3], [NAVY, 96, 17.3], [NAVY, 100, 17.3]] });
  photo(s, { x: 0.797, y: 0.703, w: 5.406, h: 6.094 });
  txt(s, 'WE ARE PROUD TO DELIVER YOUR HAPPINESS TO YOUR LOVED ONES', { x: 5.857, y: 2.056, w: 6.679, h: 1.717, fontFace: F.head, fontSize: 32, color: WHITE, bold: true, align: 'right' });
  txt(s, 'Ruddy C. Paulsen', { x: 9.891, y: 4.766, w: 2.645, h: 0.337, fontFace: F.head, fontSize: 14, color: WHITE, bold: true, align: 'right' });
  txt(s, 'Founder', { x: 9.891, y: 5.158, w: 2.645, h: 0.286, fontFace: F.head, fontSize: 11, color: WHITE, align: 'right' });
  shape(s, 'line', { x: 11.583, y: 4.571, w: 0.827, h: 0, line: RED, lineWidth: 0.75 });
  logoTag(s, 12.266, 0.181, NAVY);
}

function slide15(s) {
  s.background = { color: RED };
  shape(s, 'rect', { x: 0.879, y: 0.628, w: 4.802, h: 4.394, fill: WHITE });
  shape(s, 'rect', { x: 5.911, y: 4.233, w: 1.511, h: 3.767, rotate: 90, fill: WHITE, shadow: SH.sh4 });
  shape(s, 'rect', { x: 9.815, y: 4.233, w: 1.511, h: 3.767, rotate: 90, fill: WHITE, shadow: SH.sh4 });
  shape(s, 'rect', { x: 2.007, y: 4.233, w: 1.511, h: 3.767, rotate: 90, fill: WHITE, shadow: SH.sh4 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Nullam ac tortor vitae purus faucibus ornare suspendisse sed nisi. ', { x: 1.113, y: 3.227, w: 4.333, h: 1.182, ...T.body11GreyCL15 });
  txt(s, 'The only service you need for every type of delivery', { x: 1.228, y: 1.593, w: 3.986, h: 1.363, fontFace: F.head, fontSize: 24, color: NAVY, bold: true, align: 'center' });
  txt(s, 'OUR SERVICE', { x: 2.092, y: 1.241, w: 2.375, h: 0.286, fontFace: F.semi, fontSize: 11, color: RED, align: 'center' });
  txt(s, 'ON TIME DELIVERED', { x: 1.781, y: 5.644, w: 2.271, h: 0.303, ...T.head12NavyB });
  txt(s, L17, { x: 1.781, y: 5.963, w: 2.549, h: 0.627, ...T.body11GreyJL15 });
  txt(s, 'EASY TRACKING', { x: 5.686, y: 5.644, w: 2.549, h: 0.303, ...T.head12NavyB });
  txt(s, L17, { x: 5.686, y: 5.963, w: 2.549, h: 0.627, ...T.body11GreyJL15 });
  icon(s, 'box', { x: 5.098, y: 5.704, w: 0.433, h: 0.505, color: RED });
  txt(s, '100% GUARANTEE', { x: 9.59, y: 5.644, w: 2.271, h: 0.303, ...T.head12NavyB });
  txt(s, L17, { x: 9.59, y: 5.963, w: 2.549, h: 0.627, ...T.body11GreyJL15 });
  icon(s, 'doc', { x: 8.946, y: 5.644, w: 0.572, h: 0.572 });
  icon(s, 'stopwatch', { x: 1.123, y: 5.649, w: 0.627, h: 0.627 });
  photo(s, { x: 5.681, y: 0.628, w: 6.774, h: 4.394 });
}

function slide16(s) {
  shape(s, 'rect', { x: 9.64, y: 3.896, w: 2.616, h: 2.605, rotate: 90, fill: WHITE, shadow: SH.sh5 });
  shape(s, 'rect', { x: 6.739, y: 3.923, w: 2.616, h: 2.595, rotate: 90, fill: WHITE, shadow: SH.sh5 });
  shape(s, 'rect', { x: 9.64, y: 0.977, w: 2.616, h: 2.605, rotate: 90, fill: WHITE, shadow: SH.sh5 });
  shape(s, 'rect', { x: 6.739, y: 1.004, w: 2.616, h: 2.595, rotate: 90, fill: WHITE, shadow: SH.sh5 });
  photo(s, { x: 0, y: 0, w: 5.667, h: 7.5 });
  shape(s, 'rect', { x: 0, y: 0, w: 5.667, h: 7.5, fill: RED, transparency: 15.3 });
  txt(s, 'OUR SERVICE', { x: 0.582, y: 1.621, w: 3.713, h: 0.286, ...T.semi11White });
  txt(s, 'Turpis egestas pretium aenean pharetra magna ac placerat. Cras sed felis eget velit aliquet sagittis id consectetur purus. Habitant morbi tristique senectus et netus et. Blandit turpis cursus in hac habitasse. Eu lobortis elementum nibh tellus molestie nunc non blandit massa. Bibendum arcu vitae elementum curabitur vitae nunc sed id diam.', { x: 0.582, y: 3.147, w: 4.504, h: 1.737, ...T.body11WhiteJL15 });
  txt(s, 'Orion Express Shipping Service', { x: 0.582, y: 2.021, w: 4.504, h: 0.909, fontFace: F.head, fontSize: 24, color: WHITE, bold: true });
  txt(s, 'www.orion.com', { x: 0.661, y: 5.331, w: 1.777, h: 0.549, ...T.head12WhiteC, valign: 'middle', fill: NAVY, transparency: 15.3, shadow: SH.sh6 });
  txt(s, 'RAIL FREIGHT', { x: 9.884, y: 5.047, w: 2.128, h: 0.303, ...T.head12NavyBC });
  txt(s, 'ROAD FREIGHT', { x: 6.983, y: 5.047, w: 2.128, h: 0.303, ...T.head12NavyBC });
  txt(s, L11, { x: 9.773, y: 5.416, w: 2.349, h: 0.904, ...T.body11GreyCL15 });
  txt(s, 'OCEAN FREIGHT', { x: 9.884, y: 2.128, w: 2.128, h: 0.303, ...T.head12NavyBC });
  txt(s, 'AIR FREIGHT', { x: 6.983, y: 2.128, w: 2.128, h: 0.303, ...T.head12NavyBC });
  txt(s, L11, { x: 6.873, y: 5.416, w: 2.349, h: 0.904, ...T.body11GreyCL15 });
  txt(s, L11, { x: 6.873, y: 2.496, w: 2.349, h: 0.904, ...T.body11GreyCL15 });
  txt(s, L11, { x: 9.773, y: 2.496, w: 2.349, h: 0.904, ...T.body11GreyCL15 });
  icon(s, 'truck', { x: 7.655, y: 4.21, w: 0.784, h: 0.784 });
  icon(s, 'train', { x: 10.615, y: 4.22, w: 0.664, h: 0.664 });
  icon(s, 'plane', { x: 7.655, y: 1.228, w: 0.784, h: 0.784, rotate: 48.735 });
  icon(s, 'ship', { x: 10.615, y: 1.372, w: 0.647, h: 0.647 });
}

function slide17(s) {
  shape(s, 'rect', { x: 10.281, y: 0, w: 3.052, h: 7.5, flipH: true, fill: RED });
  txt(s, 'Our major priority is providing excellent service', { x: 1.369, y: 1.766, w: 5.434, h: 0.942, fontFace: F.head, fontSize: 25, color: NAVY, bold: true });
  txt(s, 'OUR SERVICE', { x: 1.369, y: 1.422, w: 2.408, h: 0.286, ...T.semi11Red });
  icon(s, 'check', { x: 1.497, y: 4.947, w: 0.25, h: 0.247, color: RED });
  txt(s, 'FAST AND EFFICIENT DELIVERY', { x: 1.82, y: 4.925, w: 3.413, h: 0.303, ...T.head12NavyB });
  icon(s, 'check', { x: 1.497, y: 5.373, w: 0.25, h: 0.247, color: RED });
  txt(s, 'REAL TIME TRACKING', { x: 1.82, y: 5.35, w: 3.413, h: 0.303, ...T.head12NavyB });
  icon(s, 'check', { x: 1.497, y: 4.522, w: 0.25, h: 0.247, color: RED });
  txt(s, 'SAFE & RELIABLE LOGISTIC', { x: 1.82, y: 4.499, w: 3.413, h: 0.303, ...T.head12NavyB });
  icon(s, 'check', { x: 1.497, y: 5.798, w: 0.25, h: 0.247, color: RED });
  txt(s, '100% SATISFACTION GUARANTEE', { x: 1.82, y: 5.775, w: 3.413, h: 0.303, ...T.head12NavyB });
  txt(s, 'Blandit turpis cursus in hac habitasse. Eu lobortis elementum nibh tellus molestie nunc non blandit massa. Bibendum arcu vitae elementum curabitur vitae nunc sed id diam maecenas ultricies. Turpis egestas pretium aenean pharetra magna ac placerat. ', { x: 1.369, y: 3.012, w: 5.434, h: 1.182, ...T.body11GreyJL15 });
  logoTag(s, 0.21, 0.196, NAVY);
  photo(s, { x: 8.172, y: 0, w: 4.234, h: 4.318 });
  photo(s, { x: 8.172, y: 4.621, w: 4.234, h: 2.879 });
}

function slide18(s) {
  photo(s, { x: 9.602, y: -0.01, w: 3.731, h: 7.51 });
  shape(s, 'rect', { x: 9.603, y: -0.01, w: 3.73, h: 7.51, fill: NAVY, transparency: 17.3 });
  txt(s, 'OUR SERVICE', { x: 1.204, y: 1.384, w: 4.854, h: 0.286, ...T.semi11Red });
  txt(s, 'Turpis egestas pretium aenean pharetra magna ac placerat. Cras sed felis eget velit aliquet sagittis id consectetur purus. Habitant morbi tristique senectus et netus et. Blandit turpis cursus in hac habitasse. Eu lobortis elementum nibh tellus molestie nunc non blandit massa. Bibendum arcu vitae elementum curabitur vitae nunc sed id diam maecenas ultricies.', { x: 1.204, y: 3.383, w: 4.854, h: 1.737, ...T.body11GreyJL15 });
  txt(s, 'Time is gold. We deliver it to you without wasting a second.', { x: 1.204, y: 1.778, w: 4.854, h: 1.313, ...T.head24NavyB });
  txt(s, L1, { x: 1.204, y: 5.193, w: 4.854, h: 0.904, ...T.body11GreyJL15 });
  shape(s, 'rect', { x: 8.831, y: 3.516, w: 1.544, h: 4.502, rotate: 90, fill: WHITE, shadow: SH.sh5 });
  shape(s, 'rect', { x: 8.843, y: 1.491, w: 1.524, h: 4.498, rotate: 90, fill: WHITE, shadow: SH.sh5 });
  shape(s, 'rect', { x: 8.841, y: -0.528, w: 1.524, h: 4.502, rotate: 90, fill: RED, shadow: SH.sh5 });
  txt(s, 'REGULAR DELIVERY', { x: 7.712, y: 1.25, w: 2.825, h: 0.303, ...T.head12WhiteB });
  txt(s, L13, { x: 7.712, y: 1.569, w: 3.782, h: 0.627, ...T.body11WhiteJL15 });
  txt(s, 'SAME DAY DELIVERY', { x: 7.712, y: 3.268, w: 2.825, h: 0.303, ...T.head12NavyB });
  txt(s, L13, { x: 7.712, y: 3.586, w: 3.782, h: 0.627, ...T.body11GreyJL15 });
  txt(s, 'LIGHTNING DELIVERY', { x: 7.712, y: 5.294, w: 2.825, h: 0.303, ...T.head12NavyB });
  txt(s, L13, { x: 7.712, y: 5.613, w: 3.782, h: 0.627, ...T.body11GreyJL15 });
  logoTag(s, 0.21, 0.196, NAVY);
}

function slide19(s) {
  shape(s, 'rect', { x: 0, y: 0.703, w: 13.333, h: 3.543, fill: RED });
  shape(s, 'rect', { x: 0, y: 0, w: 13.333, h: 0.703, fill: NAVY });
  txt(s, 'Turpis egestas pretium aenean pharetra magna ac placerat. Cras sed felis eget velit aliquet sagittis id consectetur purus. Habitant morbi tristique senectus et netus et. Blandit turpis cursus in hac habitasse. Eu lobortis elementum nibh tellus molestie nunc non blandit massa. ', { x: 0.509, y: 5.369, w: 4.562, h: 1.46, ...T.body11GreyJL15 });
  txt(s, L14, { x: 5.611, y: 5.999, w: 3.368, h: 0.83, fontFace: F.body, fontSize: 10, color: GREY, align: 'justify', lineSpacingMultiple: 1.5 });
  txt(s, 'WAREHOUSE', { x: 5.611, y: 5.617, w: 1.909, h: 0.303, ...T.head12NavyB });
  txt(s, L14, { x: 9.32, y: 5.999, w: 3.368, h: 0.83, fontFace: F.body, fontSize: 10, color: GREY, align: 'justify', lineSpacingMultiple: 1.5 });
  txt(s, 'TRANSPORTATION', { x: 9.32, y: 5.617, w: 2.236, h: 0.303, ...T.head12NavyB });
  txt(s, 'We also package your goods with care.', { x: 0.509, y: 2.177, w: 4.562, h: 1.01, fontFace: F.head, fontSize: 27, color: WHITE, bold: true });
  txt(s, 'OUR SERVICE', { x: 0.509, y: 1.821, w: 3.713, h: 0.286, ...T.semi11White });
  txt(s, 'BEST SERVICE FOR LOYAL CUSTOMERS', { x: 0.509, y: 4.959, w: 4.562, h: 0.303, ...T.head12NavyB });
  txt(s, 'Orion', { x: 12.014, y: 0.168, w: 1.069, h: 0.37, fontFace: F.black, fontSize: 16, color: WHITE, bold: true, align: 'center' });
  photo(s, { x: 5.611, y: 1.415, w: 3.368, h: 3.543 });
  photo(s, { x: 9.32, y: 1.415, w: 3.368, h: 3.543 });
}

function slide20(s) {
  photo(s, { x: 0.688, y: 0.531, w: 11.958, h: 6.438 });
  shape(s, 'rect', { x: 0.688, y: 0.531, w: 5.979, h: 6.438, fill: NAVY, transparency: 17.3 });
  shape(s, 'rect', { x: 6.667, y: 0.531, w: 5.979, h: 6.438, fill: RED, transparency: 17.3 });
  shape(s, 'rect', { x: 2.021, y: 2.302, w: 9.208, h: 2.896, line: WHITE, lineWidth: 6 });
  shape(s, 'rect', { x: 2.279, y: 2.502, w: 8.775, h: 2.495, fill: WHITE });
  txt(s, [{ text: 'Break', options: { color: NAVY } }, { text: ' ', options: { color: WHITE } }, { text: 'Slide.', options: { color: RED } }], { x: 2.629, y: 3.026, w: 8.075, h: 1.447, fontFace: F.black, fontSize: 80, bold: true, align: 'center' });
}

function slide21(s) {
  shape(s, 'rect', { x: 0, y: 0, w: 13.333, h: 0.703, fill: NAVY });
  photo(s, { x: 0, y: 0.703, w: 13.333, h: 4.065 });
  txt(s, 'The only service you need for every type of delivery.', { x: 1.166, y: 5.6, w: 4.647, h: 1.414, fontFace: F.head, fontSize: 26, color: NAVY, bold: true });
  txt(s, 'OUR PORTFOLIO', { x: 1.166, y: 5.255, w: 2.408, h: 0.286, ...T.semi11Red });
  txt(s, 'Turpis egestas pretium aenean pharetra magna ac placerat. Cras sed felis eget velit aliquet sagittis id consectetur purus. Habitant morbi tristique senectus et netus et blandit turpis cursus.', { x: 6.98, y: 6.03, w: 5.292, h: 0.904, ...T.body11GreyJL15 });
  shape(s, 'rect', { x: 9.461, y: 1.591, w: 1.392, h: 6.353, rotate: 90, fill: RED, shadow: SH.sh1 });
  txt(s, 'DELIVERED PACKAGE', { x: 7.335, y: 4.883, w: 1.909, h: 0.278, ...T.head10_5WhiteBC });
  txt(s, '5M+', { x: 7.615, y: 4.35, w: 1.348, h: 0.471, ...T.head22WhiteBC });
  txt(s, 'TOTAL INCOME', { x: 9.218, y: 4.883, w: 1.909, h: 0.278, ...T.head10_5WhiteBC });
  txt(s, '$10B', { x: 9.667, y: 4.35, w: 1.011, h: 0.471, ...T.head22WhiteBC });
  txt(s, '150', { x: 11.551, y: 4.35, w: 1.011, h: 0.471, ...T.head22WhiteBC });
  txt(s, 'COMPANY PARTNER', { x: 11.101, y: 4.877, w: 1.909, h: 0.278, ...T.head10_5WhiteBC });
  txt(s, 'Orion', { x: 12.014, y: 0.168, w: 1.069, h: 0.37, fontFace: F.black, fontSize: 16, color: WHITE, bold: true, align: 'center' });
}

function slide22(s) {
  shape(s, 'rect', { x: 0, y: 0, w: 13.333, h: 3.75, fill: RED });
  txt(s, 'The reliable choice for delivering your high expectations.', { x: 7.254, y: 1.758, w: 4.951, h: 1.414, fontFace: F.head, fontSize: 26, color: WHITE, bold: true });
  txt(s, 'OUR PORTFOLIO', { x: 7.254, y: 1.413, w: 2.408, h: 0.286, ...T.semi11White });
  txt(s, 'Turpis egestas pretium aenean pharetra magna ac placerat. Cras sed felis eget velit aliquet sagittis id consectetur purus. Habitant morbi tristique senectus et netus et. Blandit turpis cursus in hac habitasse lobortis elementum nibh tellus molestie.', { x: 7.254, y: 4.905, w: 4.951, h: 1.182, ...T.body11GreyJL15 });
  txt(s, 'COMPANY PARTNER', { x: 7.254, y: 4.549, w: 4.562, h: 0.303, ...T.head12NavyB });
  logoTag(s, 12.266, 0.181, WHITE);
  photo(s, { x: 0.875, y: 0.774, w: 5.25, h: 6.726 });
}

function slide23(s) {
  photo(s, { x: 0, y: 0, w: 7.686, h: 7.5 });
  shape(s, 'rect', { x: 0, y: 0, w: 7.692, h: 7.5, fill: NAVY, transparency: 15.3 });
  shape(s, 'rect', { x: 0.523, y: 4.111, w: 1.936, h: 2.285, fill: WHITE, shadow: SH.sh7 });
  txt(s, 'Your Title Here', { x: 0.674, y: 5.758, w: 1.634, h: 0.303, ...T.head12NavyBC });
  shape(s, 'ellipse', { x: 0.942, y: 4.446, w: 1.098, h: 1.093, fill: SILVER });
  shape(s, 'pie', { x: 0.942, y: 4.446, w: 1.093, h: 1.093, flipH: true, fill: RED, angleRange: [11.319, 270] });
  shape(s, 'ellipse', { x: 0.998, y: 4.502, w: 0.981, h: 0.981, fill: WHITE });
  txt(s, '70%', { x: 1.147, y: 4.808, w: 0.682, h: 0.37, ...T.semi16NavyBC });
  shape(s, 'rect', { x: 2.884, y: 4.111, w: 1.936, h: 2.285, fill: WHITE, shadow: SH.sh7 });
  txt(s, 'Your Title Here', { x: 3.035, y: 5.758, w: 1.634, h: 0.303, ...T.head12NavyBC });
  shape(s, 'ellipse', { x: 3.303, y: 4.446, w: 1.098, h: 1.093, fill: SILVER });
  shape(s, 'pie', { x: 3.303, y: 4.446, w: 1.093, h: 1.093, flipH: true, fill: RED, angleRange: [352.341, 270] });
  shape(s, 'ellipse', { x: 3.359, y: 4.502, w: 0.981, h: 0.981, fill: WHITE });
  txt(s, '80%', { x: 3.504, y: 4.808, w: 0.691, h: 0.37, ...T.semi16NavyBC });
  shape(s, 'rect', { x: 5.238, y: 4.114, w: 1.936, h: 2.285, fill: WHITE, shadow: SH.sh7 });
  txt(s, 'Your Title Here', { x: 5.389, y: 5.76, w: 1.634, h: 0.303, ...T.head12NavyBC });
  shape(s, 'ellipse', { x: 5.657, y: 4.449, w: 1.098, h: 1.093, fill: SILVER });
  shape(s, 'pie', { x: 5.657, y: 4.449, w: 1.093, h: 1.093, flipH: true, fill: RED, angleRange: [318.877, 270] });
  shape(s, 'ellipse', { x: 5.713, y: 4.505, w: 0.981, h: 0.981, fill: WHITE });
  txt(s, '90%', { x: 5.861, y: 4.81, w: 0.686, h: 0.37, ...T.semi16NavyBC });
  txt(s, L20, { x: 0.517, y: 1.446, w: 5.795, h: 0.976, fontFace: F.head, fontSize: 26, color: WHITE, bold: true });
  txt(s, 'OUR PORTFOLIO', { x: 0.517, y: 1.102, w: 2.408, h: 0.286, ...T.semi11White });
  txt(s, 'Turpis egestas pretium aenean pharetra magna ac placerat. Cras sed felis eget velit aliquet sagittis id consectetur purus. Habitant morbi tristique senectus et netus et. Blandit turpis cursus in hac habitasse lobortis elementum nibh tellus.', { x: 0.517, y: 2.595, w: 6.657, h: 0.904, ...T.body11WhiteJL15 });
  icon(s, 'check', { x: 8.499, y: 1.369, w: 0.25, h: 0.247, color: RED });
  txt(s, 'YOUR TITLE HERE', { x: 8.765, y: 1.34, w: 3.334, h: 0.303, ...T.head12NavyB });
  txt(s, L7, { x: 8.765, y: 1.659, w: 3.766, h: 0.903, ...T.body11GreyJL15 });
  icon(s, 'check', { x: 8.499, y: 3.138, w: 0.25, h: 0.247, color: RED });
  txt(s, 'YOUR TITLE HERE', { x: 8.765, y: 3.11, w: 3.334, h: 0.303, ...T.head12NavyB });
  txt(s, L7, { x: 8.765, y: 3.429, w: 3.766, h: 0.903, ...T.body11GreyJL15 });
  icon(s, 'check', { x: 8.499, y: 4.965, w: 0.25, h: 0.247, color: RED });
  txt(s, 'YOUR TITLE HERE', { x: 8.765, y: 4.936, w: 3.334, h: 0.303, ...T.head12NavyB });
  txt(s, L7, { x: 8.765, y: 5.255, w: 3.766, h: 0.904, ...T.body11GreyJL15 });
  logoTag(s, 12.266, 0.181, NAVY);
}

function slide24(s) {
  shape(s, 'rect', { x: 0, y: 5.125, w: 9.901, h: 2.375, fill: RED });
  txt(s, 'OUR PORTFOLIO', { x: 0.796, y: 1.134, w: 3.667, h: 0.286, ...T.semi11Red });
  txt(s, L4, { x: 0.796, y: 3.093, w: 4.854, h: 1.182, ...T.body11GreyJL15 });
  txt(s, 'Guaranteed global logistics services to every corner of the globe.', { x: 0.796, y: 1.529, w: 4.854, h: 1.313, ...T.head24NavyB });
  txt(s, L15, { x: 0.796, y: 6.087, w: 3.368, h: 0.83, ...T.body10WhiteJL15 });
  txt(s, 'YOUR TITLE HERE', { x: 0.796, y: 5.704, w: 1.909, h: 0.303, ...T.head12WhiteB });
  txt(s, L15, { x: 5.328, y: 6.087, w: 3.368, h: 0.83, ...T.body10WhiteJL15 });
  txt(s, 'YOUR TITLE HERE', { x: 5.328, y: 5.704, w: 2.236, h: 0.303, ...T.head12WhiteB });
  logoTag(s, 0.21, 0.196, NAVY);
  photo(s, { x: 6.446, y: 0, w: 3.455, h: 5.125 });
  photo(s, { x: 9.901, y: 0, w: 3.455, h: 7.5, dark: true });
}

function slide25(s) {
  shape(s, 'rect', { x: 1.097, y: 1.245, w: 4.838, h: 5.01, fill: RED });
  txt(s, 'Orion Express Gallery', { x: 1.611, y: 2.228, w: 3.397, h: 1.043, fontFace: F.head, fontSize: 28, color: WHITE, bold: true });
  txt(s, 'OUR GALLERY', { x: 1.611, y: 1.872, w: 2.881, h: 0.286, ...T.semi11White });
  txt(s, 'Blandit turpis cursus in hac habitasse. Eu lobortis elementum nibh tellus molestie nunc non blandit massa. Bibendum arcu vitae elementum curabitur vitae nunc sed id diam maecenas ultricies.', { x: 1.611, y: 3.558, w: 3.785, h: 1.182, ...T.body11WhiteJL15 });
  txt(s, 'www.orion.com', { x: 1.728, y: 5.173, w: 1.581, h: 0.455, ...T.head12WhiteC, valign: 'middle', fill: NAVY, transparency: 15.3, shadow: SH.sh6 });
  photo(s, { x: 8.713, y: 1.245, w: 3.523, h: 2.505 });
  photo(s, { x: 6.054, y: 1.245, w: 2.54, h: 2.505 });
  photo(s, { x: 6.054, y: 3.862, w: 6.183, h: 2.393, dark: true });
}

function slide26(s) {
  s.background = { color: NAVY };
  shape(s, 'rect', { x: 0.565, y: 0.521, w: 4.521, h: 6.458, line: WHITE, lineWidth: 6 });
  txt(s, 'TIME IS GOLD. WE DELIVER IT TO YOU WITHOUT WASTING A SECOND.', { x: 1.254, y: 1.733, w: 3.143, h: 4.342, fontFace: F.head, fontSize: 36, color: WHITE, bold: true, align: 'center' });
  shape(s, 'line', { x: 2.45, y: 1.425, w: 0.752, h: 0, line: WHITE, lineWidth: 1 });
  photo(s, { x: 5.651, y: 0, w: 7.683, h: 7.5 });
}

function slide27(s) {
  s.background = { color: RED };
  shape(s, 'custGeom', { x: 8.325, y: 0, w: 5.008, h: 5.671, poly: [[0.853, 0], [5.008, 0], [5.008, 5.277], [4.748, 5.402], [3.417, 5.671, 4.339, 5.575, 3.889, 5.671], [0, 2.253, 1.53, 5.671, 0, 4.141], [0.78, 0.079, 0, 1.428, 0.293, 0.67]], fill: NAVY });
  device(s, 'phone', { x: 8.734, y: 3.336, w: 3.541, h: 7.137, rotate: -37.774 });
  device(s, 'phone', { x: 11.005, y: -0.406, w: 3.541, h: 7.137, rotate: -37.774 });
  txt(s, 'ALWAYS YOUR RELIABLE PARTNER IN LOGISTICS', { x: 0.676, y: 2.877, w: 5.738, h: 1.919, fontFace: F.head, fontSize: 36, color: WHITE, bold: true });
  txt(s, 'MOCKUP DEVICE', { x: 0.676, y: 2.562, w: 1.685, h: 0.286, fontFace: F.mulBlack, fontSize: 11, color: WHITE });
  txt(s, 'Turpis egestas pretium aenean pharetra magna ac placerat. Cras sed felis eget velit aliquet sagittis id consectetur purus. Habitant morbi tristique senectus et netus et. Blandit turpis cursus in hac habitasse. ', { x: 0.676, y: 5.296, w: 5.738, h: 0.904, ...T.body11WhiteJL15 });
  logoTag(s, 0.21, 0.196, WHITE);
  photo(s, { x: 11.089, y: -0.201, w: 3.204, h: 6.652, rotate: -37.315, shape: 'roundRect', rectRadius: 0.337 });
  photo(s, { x: 8.902, y: 3.558, w: 3.204, h: 6.595, rotate: -37.315, shape: 'roundRect', rectRadius: 0.337 });
}

function slide28(s) {
  s.background = { color: RED };
  shape(s, 'ellipse', { x: 2.258, y: -0.659, w: 8.818, h: 8.818, fill: WHITE, transparency: 90.2 });
  shape(s, 'ellipse', { x: 3.599, y: 0.682, w: 6.135, h: 6.135, fill: WHITE, transparency: 90.2 });
  device(s, 'phone', { x: 5.125, y: 0.688, w: 3.084, h: 6.125 });
  txt(s, 'OR1134A56', { x: 8.832, y: 3.462, w: 2.422, h: 0.543, flipH: true, fontFace: F.head, fontSize: 14, color: MIDGREY, bold: true, align: 'center', valign: 'middle', fill: WHITE, shadow: SH.sh5 });
  txt(s, 'TRACKING', { x: 11.365, y: 3.462, w: 1.417, h: 0.543, ...T.head12WhiteBC, valign: 'middle', fill: NAVY, shadow: SH.sh5 });
  txt(s, 'Easy Tracking Shipment ', { x: 8.721, y: 2.787, w: 3.381, h: 0.326, ...T.head14WhiteB });
  txt(s, 'Cras sed felis eget velit aliquet sagittis id consectetur purus. Habitant morbi tristique senectus et netus et. Blandit turpis cursus in hac habitasse. ', { x: 8.721, y: 4.353, w: 4.061, h: 0.903, ...T.body11WhiteJL15 });
  txt(s, 'Download the Orion Express app now', { x: 0.532, y: 2.589, w: 4.061, h: 1.504, fontFace: F.head, fontSize: 28, color: WHITE, bold: true });
  txt(s, 'MOCKUP DEVICE', { x: 0.532, y: 2.255, w: 3.718, h: 0.286, fontFace: F.mulBlack, fontSize: 11, color: WHITE });
  shape(s, 'roundRect', { x: 0.59, y: 4.65, w: 1.876, h: 0.595, fill: '0C0C0C', rectRadius: 0.047 });
  icon(s, 'play', { x: 0.699, y: 4.773, w: 0.364, h: 0.364, color: 'F6B94C' });
  txt(s, 'GET IN ON ', { x: 1.091, y: 4.687, w: 0.689, h: 0.219, fontFace: F.pop, fontSize: 7, color: WHITE });
  txt(s, 'Google Play ', { x: 1.081, y: 4.872, w: 1.559, h: 0.337, fontFace: F.popSemi, fontSize: 14, color: WHITE, bold: true });
  shape(s, 'roundRect', { x: 2.693, y: 4.65, w: 1.877, h: 0.595, fill: '0C0C0C', rectRadius: 0.047 });
  txt(s, 'Download on the  ', { x: 3.304, y: 4.687, w: 1.246, h: 0.219, fontFace: F.pop, fontSize: 7, color: WHITE });
  txt(s, 'App Store ', { x: 3.294, y: 4.872, w: 1.367, h: 0.337, fontFace: F.popSemi, fontSize: 14, color: WHITE, bold: true });
  icon(s, 'apple', { x: 2.825, y: 4.731, w: 0.372, h: 0.433, color: WHITE });
  logoTag(s, 12.266, 0.181, WHITE);
  photo(s, { x: 5.307, y: 0.872, w: 2.718, h: 5.755, shape: 'roundRect', rectRadius: 0.288 });
}

function slide29(s) {
  shape(s, 'rect', { x: 0, y: 0, w: 4.015, h: 7.5, fill: NAVY });
  device(s, 'monitor', { x: 1.369, y: 1.375, w: 5.297, h: 4.75 });
  txt(s, 'MOCKUP DEVICE', { x: 7.555, y: 1.912, w: 3.667, h: 0.286, fontFace: F.mulBlack, fontSize: 11, color: RED });
  txt(s, L4, { x: 7.555, y: 3.477, w: 4.854, h: 1.182, ...T.body11GreyJL15 });
  txt(s, 'Download apps orion express get free shipping', { x: 7.555, y: 2.307, w: 4.854, h: 0.909, ...T.head24NavyB });
  txt(s, 'www.orion.com', { x: 7.666, y: 5.132, w: 1.581, h: 0.455, ...T.head12WhiteC, valign: 'middle', fill: RED, transparency: 15.3, shadow: SH.sh6 });
  logoTag(s, 12.266, 0.181, NAVY);
  photo(s, { x: 1.574, y: 1.632, w: 4.853, h: 2.941 });
}

function slide30(s) {
  shape(s, 'rect', { x: 0, y: 0, w: 5.667, h: 7.5, fill: RED });
  shape(s, 'rect', { x: 0.718, y: 1.007, w: 4.232, h: 5.487, flipH: true, fill: WHITE });
  shape(s, 'pie', { x: 1.126, y: 2.023, w: 3.454, h: 3.454, fill: '113CFC', angleRange: [287.924, 32.223] });
  shape(s, 'pie', { x: 1.126, y: 2.023, w: 3.454, h: 3.454, fill: '1597E5', angleRange: [252.684, 288.2] });
  shape(s, 'pie', { x: 1.126, y: 2.023, w: 3.454, h: 3.454, fill: '69DADB', angleRange: [170.587, 252.803] });
  shape(s, 'pie', { x: 1.126, y: 2.023, w: 3.454, h: 3.454, fill: NAVY, angleRange: [11.632, 171.858] });
  shape(s, 'ellipse', { x: 2.529, y: 3.455, w: 0.596, h: 0.596, fill: WHITE });
  txt(s, '60%', { x: 2.472, y: 4.545, w: 0.709, h: 0.404, ...T.popSemi18WhiteBC });
  txt(s, '40%', { x: 3.494, y: 3.168, w: 0.716, h: 0.404, ...T.popSemi18WhiteBC });
  txt(s, '20%', { x: 2.586, y: 2.402, w: 0.584, h: 0.337, fontFace: F.popSemi, fontSize: 14, color: WHITE, bold: true, align: 'center' });
  txt(s, '40%', { x: 1.469, y: 3.168, w: 0.716, h: 0.404, ...T.popSemi18WhiteBC });
  txt(s, 'Diagram Style', { x: 1.583, y: 1.35, w: 2.487, h: 0.337, ...T.head14InkBC });
  txt(s, 'Lorem dolor sed viverra ipsum nunc aliquet bibendum enim facilisis.. ', { x: 1.14, y: 5.629, w: 3.387, h: 0.627, ...T.body11GreyCL15 });
  txt(s, L8, { x: 7.943, y: 3, w: 3.747, h: 0.625, ...T.body11GreyJL15 });
  txt(s, 'YOUR TITLE HERE', { x: 7.941, y: 2.539, w: 2.806, h: 0.337, ...T.head14InkB });
  txt(s, L8, { x: 7.943, y: 1.665, w: 3.747, h: 0.625, ...T.body11GreyJL15 });
  txt(s, 'YOUR TITLE HERE', { x: 7.941, y: 1.203, w: 2.704, h: 0.337, ...T.head14InkB });
  txt(s, L8, { x: 7.941, y: 4.336, w: 3.748, h: 0.625, ...T.body11GreyJL15 });
  txt(s, 'YOUR TITLE HERE', { x: 7.94, y: 3.874, w: 2.539, h: 0.415, fontFace: F.head, fontSize: 14, color: INK, bold: true, lineSpacingMultiple: 1.5 });
  txt(s, L8, { x: 7.941, y: 5.672, w: 3.748, h: 0.625, ...T.body11GreyJL15 });
  txt(s, 'YOUR TITLE HERE', { x: 7.94, y: 5.21, w: 2.539, h: 0.415, fontFace: F.head, fontSize: 14, color: INK, bold: true, lineSpacingMultiple: 1.5 });
  shape(s, 'ellipse', { x: 7.311, y: 1.346, w: 0.51, h: 0.51, fill: NAVY });
  txt(s, '01', { x: 7.378, y: 1.449, w: 0.365, h: 0.303, fontFace: F.popSemi, fontSize: 12, color: WHITE });
  shape(s, 'ellipse', { x: 7.311, y: 2.681, w: 0.51, h: 0.51, fill: '113CFC' });
  txt(s, '02', { x: 7.36, y: 2.785, w: 0.411, h: 0.303, ...T.pop12White });
  shape(s, 'ellipse', { x: 7.311, y: 4.017, w: 0.51, h: 0.51, fill: '1597E5' });
  txt(s, '03', { x: 7.361, y: 4.121, w: 0.409, h: 0.303, ...T.pop12White });
  shape(s, 'ellipse', { x: 7.311, y: 5.353, w: 0.51, h: 0.51, fill: '69DADB' });
  txt(s, '04', { x: 7.361, y: 5.456, w: 0.425, h: 0.303, ...T.pop12White });
  logoTag(s, 12.266, 0.181, NAVY);
}

function slide31(s) {
  shape(s, 'rect', { x: 7.674, y: -0.012, w: 5.667, h: 7.5, fill: RED });
  shape(s, 'rect', { x: 8.583, y: 1.321, w: 3.747, h: 4.859, flipH: true, fill: WHITE });
  txt(s, 'Chart Style', { x: 8.842, y: 1.665, w: 2.462, h: 0.505, fontFace: F.head, fontSize: 24, color: INK, bold: true });
  txt(s, 'Tortor id aliquet lectus proin nibh nisl.', { x: 8.829, y: 2.24, w: 2.931, h: 0.346, fontFace: F.mul, fontSize: 11, color: GREY, align: 'justify', lineSpacingMultiple: 1.5 });
  shape(s, 'round2SameRect', { x: 8.97, y: 4.794, w: 0.312, h: 0.764, fill: BLUE });
  shape(s, 'round2SameRect', { x: 9.417, y: 4.492, w: 0.294, h: 1.066, fill: '1597E5' });
  shape(s, 'round2SameRect', { x: 9.865, y: 4.237, w: 0.31, h: 1.321, fill: BLUE });
  shape(s, 'round2SameRect', { x: 10.313, y: 3.907, w: 0.305, h: 1.651, fill: '1597E5' });
  shape(s, 'round2SameRect', { x: 10.76, y: 3.633, w: 0.325, h: 1.925, fill: BLUE });
  shape(s, 'round2SameRect', { x: 11.208, y: 3.232, w: 0.301, h: 2.326, fill: '1597E5' });
  shape(s, 'round2SameRect', { x: 11.655, y: 2.94, w: 0.301, h: 2.618, fill: BLUE });
  txt(s, 'Mon', { x: 8.842, y: 5.558, w: 0.56, h: 0.252, ...T.head9InkC });
  txt(s, 'Tue', { x: 9.287, y: 5.558, w: 0.56, h: 0.252, ...T.head9InkC });
  txt(s, 'Wed', { x: 9.732, y: 5.558, w: 0.56, h: 0.252, ...T.head9InkC });
  txt(s, 'Thu', { x: 10.177, y: 5.558, w: 0.56, h: 0.252, ...T.head9InkC });
  txt(s, 'Fri', { x: 10.622, y: 5.558, w: 0.56, h: 0.252, ...T.head9InkC });
  txt(s, 'Sat', { x: 11.067, y: 5.558, w: 0.56, h: 0.252, ...T.head9InkC });
  txt(s, 'Sun', { x: 11.512, y: 5.558, w: 0.56, h: 0.252, ...T.head9InkC });
  txt(s, L1, { x: 1.406, y: 2.509, w: 4.832, h: 0.904, ...T.body11GreyJL15 });
  txt(s, 'The value of an idea lies in the using of it', { x: 1.406, y: 1.208, w: 4.832, h: 0.909, fontFace: F.head, fontSize: 24, color: BLUE, bold: true });
  txt(s, 'Your Title Here', { x: 2.619, y: 3.738, w: 1.634, h: 0.303, fontFace: F.head, fontSize: 12, color: BLUE, bold: true });
  shape(s, 'ellipse', { x: 1.502, y: 3.742, w: 0.908, h: 0.904, fill: SILVER });
  shape(s, 'pie', { x: 1.502, y: 3.742, w: 0.904, h: 0.904, flipH: true, fill: RED, angleRange: [11.319, 270] });
  shape(s, 'ellipse', { x: 1.549, y: 3.788, w: 0.812, h: 0.812, fill: WHITE });
  txt(s, '70%', { x: 1.669, y: 4.041, w: 0.691, h: 0.37, fontFace: F.head, fontSize: 16, color: BLUE, bold: true, align: 'center' });
  txt(s, L18, { x: 2.618, y: 4.097, w: 3.39, h: 0.627, ...T.body11GreyJL15 });
  txt(s, 'Your Title Here', { x: 2.619, y: 5.195, w: 1.634, h: 0.303, fontFace: F.head, fontSize: 12, color: BLUE, bold: true });
  shape(s, 'ellipse', { x: 1.502, y: 5.199, w: 0.908, h: 0.904, fill: SILVER });
  shape(s, 'pie', { x: 1.502, y: 5.199, w: 0.904, h: 0.904, flipH: true, fill: RED, angleRange: [312.057, 270] });
  shape(s, 'ellipse', { x: 1.548, y: 5.245, w: 0.812, h: 0.812, fill: WHITE });
  txt(s, '90%', { x: 1.667, y: 5.498, w: 0.693, h: 0.37, fontFace: F.head, fontSize: 16, color: BLUE, bold: true, align: 'center' });
  txt(s, L18, { x: 2.618, y: 5.554, w: 3.39, h: 0.627, ...T.body11GreyJL15 });
  logoTag(s, 0.21, 0.196, NAVY);
}

function slide32(s) {
  photo(s, { x: 0.02, y: -0.013, w: 13.333, h: 7.513 });
  shape(s, 'rect', { x: 0, y: -0.007, w: 13.333, h: 7.507, fill: RED, transparency: 17.3 });
  shape(s, 'rect', { x: 8.007, y: 2.611, w: 2.922, h: 3.913, fill: WHITE });
  txt(s, '$10', { x: 8.719, y: 3.582, w: 1.618, h: 0.572, fontFace: F.popSemi, fontSize: 28, color: INK, align: 'center' });
  txt(s, [{ text: 'Manage Service', options: { breakLine: true } }, { text: 'Problem Solution', options: { breakLine: true } }, { text: 'Digital Marketing', options: { breakLine: true } }, { text: '24/7 Customer Support', options: { breakLine: true } }], { x: 8.37, y: 4.214, w: 2.343, h: 1.526, fontFace: F.body, fontSize: 11, color: GREY, align: 'center', lineSpacingMultiple: 2 });
  txt(s, 'Get Started', { x: 8.738, y: 6.289, w: 1.618, h: 0.469, fontFace: F.head, fontSize: 10, color: WHITE, bold: true, align: 'center', valign: 'middle', fill: BLUE });
  txt(s, 'SAMEDAY', { x: 8.544, y: 3.163, w: 1.969, h: 0.337, ...T.head14InkBC });
  shape(s, 'rect', { x: 2.405, y: 2.611, w: 2.922, h: 3.913, fill: WHITE });
  txt(s, '$5', { x: 2.996, y: 3.582, w: 1.618, h: 0.572, fontFace: F.popSemi, fontSize: 28, color: INK, align: 'center' });
  txt(s, [{ text: 'Advance Security', options: { breakLine: true } }, { text: 'Data Visualization', options: { breakLine: true } }, { text: 'Digital Marketing', options: { breakLine: true } }, { text: '24/7 Customer Support', options: { breakLine: true } }], { x: 2.633, y: 4.214, w: 2.343, h: 1.526, fontFace: F.body, fontSize: 11, color: GREY, align: 'center', lineSpacingMultiple: 2 });
  txt(s, 'Get Started', { x: 3.032, y: 6.289, w: 1.618, h: 0.469, fontFace: F.head, fontSize: 10, color: WHITE, bold: true, align: 'center', valign: 'middle', fill: BLUE });
  txt(s, 'REGULAR', { x: 3.011, y: 3.163, w: 1.595, h: 0.337, ...T.head14InkBC });
  shape(s, 'rect', { x: 5.206, y: 2.473, w: 2.922, h: 4.345, fill: BLUE, shadow: { type: 'outer', color: '571FD4', blur: 24, offset: 0, angle: 0, opacity: 0.118 } });
  txt(s, '$15', { x: 5.858, y: 3.385, w: 1.618, h: 0.572, fontFace: F.popSemi, fontSize: 28, color: WHITE, align: 'center' });
  txt(s, [{ text: 'Free Consultation', options: { breakLine: true } }, { text: 'Integration Manage Service', options: { breakLine: true } }, { text: 'Digital Marketing', options: { breakLine: true } }, { text: '24/7 Customer Support', options: { breakLine: true } }], { x: 5.495, y: 4.011, w: 2.343, h: 1.526, fontFace: F.body, fontSize: 11, color: WHITE, align: 'center', lineSpacingMultiple: 2 });
  txt(s, 'Get Started', { x: 5.878, y: 5.951, w: 1.618, h: 0.469, fontFace: F.head, fontSize: 10, color: INK, bold: true, align: 'center', valign: 'middle', fill: WHITE });
  txt(s, 'LIGHTNING ', { x: 5.869, y: 2.966, w: 1.595, h: 0.337, ...T.head14WhiteBC });
  txt(s, 'Our Pricing Table ', { x: 3.616, y: 0.586, w: 6.101, h: 0.572, fontFace: F.head, fontSize: 28, color: WHITE, bold: true, align: 'center' });
  txt(s, L9, { x: 2.718, y: 1.308, w: 7.897, h: 0.627, ...T.body11WhiteCL15 });
}

function slide33(s) {
  shape(s, 'rect', { x: 1.173, y: 2.697, w: 2.326, h: 2.326, fill: RED });
  shape(s, 'rect', { x: 4.06, y: 2.697, w: 2.326, h: 2.326, fill: NAVY });
  shape(s, 'rect', { x: 6.947, y: 2.697, w: 2.326, h: 2.326, fill: RED });
  shape(s, 'rect', { x: 9.835, y: 2.697, w: 2.326, h: 2.326, fill: NAVY });
  txt(s, 'Our Services', { x: 1.598, y: 4.123, w: 1.497, h: 0.337, ...T.head14WhiteBC });
  txt(s, 'Our Services', { x: 4.474, y: 4.123, w: 1.497, h: 0.337, ...T.head14WhiteBC });
  txt(s, 'Our Services', { x: 7.358, y: 4.123, w: 1.497, h: 0.337, ...T.head14WhiteBC });
  txt(s, 'Our Services', { x: 10.249, y: 4.123, w: 1.497, h: 0.337, ...T.head14WhiteBC });
  txt(s, L10, { x: 1.033, y: 5.566, w: 2.636, h: 0.904, ...T.body11GreyCL15 });
  txt(s, L10, { x: 3.937, y: 5.566, w: 2.636, h: 0.934, ...T.body11GreyCL15 });
  txt(s, L10, { x: 6.76, y: 5.566, w: 2.636, h: 0.934, ...T.body11GreyCL15 });
  txt(s, L10, { x: 9.666, y: 5.566, w: 2.636, h: 0.934, ...T.body11GreyCL15 });
  shape(s, 'line', { x: 1.173, y: 5.289, w: 2.326, h: 0, line: RED, lineWidth: 4.5 });
  shape(s, 'line', { x: 4.06, y: 5.289, w: 2.326, h: 0, line: NAVY, lineWidth: 4.5 });
  shape(s, 'line', { x: 6.947, y: 5.308, w: 2.326, h: 0, line: RED, lineWidth: 4.5 });
  shape(s, 'line', { x: 9.835, y: 5.289, w: 2.326, h: 0, line: NAVY, lineWidth: 4.5 });
  txt(s, 'Our Infographic', { x: 3.616, y: 0.586, w: 6.101, h: 0.572, ...T.head28NavyBC });
  txt(s, L9, { x: 2.718, y: 1.308, w: 7.897, h: 0.627, ...T.body11GreyCL15 });
  icon(s, 'store', { x: 2.043, y: 3.273, w: 0.585, h: 0.585, color: WHITE });
  icon(s, 'paperplane', { x: 4.931, y: 3.278, w: 0.585, h: 0.585, color: WHITE });
  icon(s, 'search', { x: 7.814, y: 3.282, w: 0.587, h: 0.585, color: WHITE });
  icon(s, 'person', { x: 10.704, y: 3.284, w: 0.587, h: 0.584, color: WHITE });
}

function slide34(s) {
  txt(s, 'Write your title', { x: 2.063, y: 3.337, w: 1.844, h: 0.337, fontFace: F.head, fontSize: 14, color: INK, bold: true, align: 'right' });
  txt(s, L6, { x: 1.528, y: 3.668, w: 2.379, h: 0.627, fontFace: F.body, fontSize: 11, color: GREY, align: 'right', lineSpacingMultiple: 1.5 });
  txt(s, '1', { x: 3.449, y: 2.688, w: 0.458, h: 0.458, ...T.head12WhiteC, valign: 'middle', shape: 'ellipse', fill: NAVY });
  txt(s, 'Write your title', { x: 2.063, y: 5.479, w: 1.844, h: 0.337, fontFace: F.head, fontSize: 14, color: INK, bold: true, align: 'right' });
  txt(s, L6, { x: 1.139, y: 5.813, w: 2.768, h: 0.627, fontFace: F.body, fontSize: 11, color: GREY, align: 'right', lineSpacingMultiple: 1.5 });
  txt(s, '2', { x: 3.449, y: 4.83, w: 0.458, h: 0.458, ...T.head12WhiteC, valign: 'middle', shape: 'ellipse', fill: RED });
  txt(s, 'Write your title', { x: 9.427, y: 3.337, w: 1.844, h: 0.337, ...T.head14InkB });
  txt(s, L6, { x: 9.427, y: 3.671, w: 2.379, h: 0.627, ...T.body11GreyL15 });
  txt(s, '4', { x: 9.526, y: 2.688, w: 0.458, h: 0.458, ...T.head12WhiteC, valign: 'middle', shape: 'ellipse', fill: RED });
  txt(s, 'Write your title', { x: 9.427, y: 5.479, w: 1.844, h: 0.337, ...T.head14InkB });
  txt(s, L6, { x: 9.427, y: 5.813, w: 2.379, h: 0.627, ...T.body11GreyL15 });
  txt(s, '3', { x: 9.526, y: 4.83, w: 0.458, h: 0.458, ...T.head12WhiteC, valign: 'middle', shape: 'ellipse', fill: NAVY });
  shape(s, 'teardrop', { x: 6.43, y: 2.281, w: 2.179, h: 2.179, rotate: -180, fill: SILVER });
  shape(s, 'teardrop', { x: 6.552, y: 2.403, w: 1.934, h: 1.934, rotate: -180, fill: RED });
  icon(s, 'home', { x: 7.288, y: 3.139, w: 0.463, h: 0.463, color: WHITE });
  shape(s, 'teardrop', { x: 6.43, y: 4.509, w: 2, h: 2, rotate: -90, fill: SILVER });
  shape(s, 'teardrop', { x: 6.542, y: 4.621, w: 1.776, h: 1.776, rotate: -90, fill: NAVY });
  icon(s, 'clock', { x: 7.198, y: 5.277, w: 0.463, h: 0.463, color: WHITE });
  shape(s, 'teardrop', { x: 4.98, y: 3.066, w: 1.394, h: 1.394, rotate: 90, fill: SILVER });
  shape(s, 'teardrop', { x: 5.058, y: 3.144, w: 1.238, h: 1.238, rotate: 90, fill: NAVY });
  icon(s, 'gear', { x: 5.516, y: 3.601, w: 0.323, h: 0.323, color: WHITE });
  shape(s, 'teardrop', { x: 4.724, y: 4.509, w: 1.65, h: 1.65, fill: SILVER });
  shape(s, 'teardrop', { x: 4.817, y: 4.601, w: 1.465, h: 1.465, fill: RED });
  icon(s, 'bolt', { x: 5.46, y: 5.17, w: 0.178, h: 0.326, color: WHITE });
  txt(s, 'Our Infographic', { x: 3.616, y: 0.815, w: 6.101, h: 0.572, ...T.head28NavyBC });
}

function slide35(s) {
  txt(s, 'Your Step One', { x: 1.18, y: 3.337, w: 1.745, h: 0.337, ...T.head14InkB });
  txt(s, L6, { x: 1.18, y: 3.671, w: 2.471, h: 0.627, ...T.body11GreyL15 });
  txt(s, '1', { x: 1.28, y: 2.736, w: 0.458, h: 0.458, ...T.semi12WhiteBC, valign: 'middle', shape: 'ellipse', fill: NAVY });
  txt(s, L6, { x: 1.18, y: 5.813, w: 2.471, h: 0.627, ...T.body11GreyL15 });
  txt(s, '3', { x: 1.28, y: 4.878, w: 0.458, h: 0.458, ...T.semi12WhiteBC, valign: 'middle', shape: 'ellipse', fill: NAVY });
  txt(s, L6, { x: 9.909, y: 3.678, w: 2.615, h: 0.627, ...T.body11GreyL15 });
  txt(s, '2', { x: 10.009, y: 2.742, w: 0.458, h: 0.458, ...T.semi12WhiteBC, valign: 'middle', shape: 'ellipse', fill: RED });
  txt(s, L6, { x: 9.909, y: 5.82, w: 2.504, h: 0.627, ...T.body11GreyL15 });
  txt(s, '4', { x: 10.009, y: 4.884, w: 0.458, h: 0.458, ...T.semi12WhiteBC, valign: 'middle', shape: 'ellipse', fill: RED });
  shape(s, 'custGeom', { x: 5.522, y: 2.878, w: 1.827, h: 1.619, poly: [[1.449, 0], [0.349, 0.97], [0, 1.619], [1.827, 0.582], [1.449, 0]], fill: SILVER });
  shape(s, 'custGeom', { x: 4.968, y: 3.838, w: 2.942, h: 1.685, poly: [[2.593, 0], [0.349, 1.037], [0, 1.685], [2.942, 0.649], [2.593, 0, 2.942, 0.649, 2.593, 0]], fill: SILVER });
  shape(s, 'custGeom', { x: 4.414, y: 4.878, w: 4.056, h: 1.685, poly: [[3.708, 0], [0.349, 1.037], [0, 1.685], [4.056, 0.649], [3.708, 0, 4.056, 0.649, 3.708, 0]], fill: SILVER });
  shape(s, 'custGeom', { x: 4.414, y: 5.919, w: 4.614, h: 0.649, poly: [[4.614, 0.649], [4.265, 0], [0.349, 0], [0, 0.649], [4.614, 0.649, 0, 0.649, 4.614, 0.649]], fill: RED });
  icon(s, 'home', { x: 6.518, y: 6.058, w: 0.406, h: 0.406, color: WHITE });
  shape(s, 'custGeom', { x: 6.076, y: 2.43, w: 1.274, h: 1.031, poly: [[0.637, 0], [1.274, 1.031], [0, 1.031]], fill: NAVY });
  icon(s, 'paperplane', { x: 6.55, y: 2.964, w: 0.341, h: 0.341, color: WHITE });
  shape(s, 'custGeom', { x: 5.522, y: 3.838, w: 2.384, h: 0.649, poly: [[0.349, 0], [0, 0.649], [2.384, 0.649], [2.035, 0], [0.349, 0, 2.035, 0, 0.349, 0]], fill: RED });
  icon(s, 'clock', { x: 6.518, y: 3.959, w: 0.406, h: 0.406, color: WHITE });
  shape(s, 'custGeom', { x: 4.968, y: 4.878, w: 3.499, h: 0.649, poly: [[0.349, 0], [0, 0.649], [3.499, 0.649], [3.15, 0], [0.349, 0, 3.15, 0, 0.349, 0]], fill: NAVY });
  icon(s, 'gear', { x: 6.518, y: 5, w: 0.406, h: 0.406, color: WHITE });
  txt(s, 'Your Step Two', { x: 9.907, y: 3.337, w: 1.78, h: 0.337, ...T.head14InkB });
  txt(s, 'Your Step Four', { x: 9.906, y: 5.483, w: 2.018, h: 0.337, ...T.head14InkB });
  txt(s, 'Your Step Three', { x: 1.18, y: 5.477, w: 1.916, h: 0.337, ...T.head14InkB });
  txt(s, 'Our Infographic', { x: 4.098, y: 0.634, w: 5.167, h: 0.64, fontFace: F.head, fontSize: 32, color: NAVY, bold: true, align: 'center' });
}

function slide36(s) {
  s.background = { color: RED };
  shape(s, 'rect', { x: 0, y: 4.531, w: 3.854, h: 2.969, fill: NAVY });
  shape(s, 'rect', { x: 1.349, y: 4.531, w: 5.081, h: 2.438, fill: WHITE });
  txt(s, 'Let’s Get in Touch', { x: 1.937, y: 5.143, w: 3.905, h: 0.539, fontFace: F.head, fontSize: 26, color: NAVY, bold: true });
  txt(s, 'CONTACT US', { x: 1.937, y: 4.809, w: 3.718, h: 0.286, ...T.semi11Red });
  txt(s, 'Turpis egestas pretium aenean pharetra magna ac placerat. Cras sed felis eget velit aliquet sagittis id consectetur purus. ', { x: 1.937, y: 5.787, w: 3.887, h: 0.904, ...T.body11GreyJL15 });
  txt(s, '13th Street, New York, NY 10011, USA.', { x: 8.369, y: 5.431, w: 3.363, h: 0.349, fontFace: F.body, fontSize: 11, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, '+62 205 7516  / orionglobal@mail.com', { x: 8.369, y: 6.32, w: 3.363, h: 0.349, fontFace: F.body, fontSize: 11, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Our Address', { x: 8.369, y: 5.107, w: 2.125, h: 0.324, fontFace: F.head, fontSize: 12, color: WHITE, bold: true, lineSpacingMultiple: 1.2 });
  txt(s, 'Phone & Email', { x: 8.375, y: 5.996, w: 2.125, h: 0.324, fontFace: F.head, fontSize: 12, color: WHITE, bold: true, lineSpacingMultiple: 1.2 });
  icon(s, 'phone', { x: 7.975, y: 6.101, w: 0.24, h: 0.349, color: WHITE });
  icon(s, 'pin', { x: 7.964, y: 5.194, w: 0.262, h: 0.349, color: WHITE });
  photo(s, { x: 0, y: 0, w: 13.333, h: 4.531 });
}

function slide37(s) {
  logoMark(s, 6.133, 2.232, 1.068);
  txt(s, L19, { x: 3.249, y: 4.958, w: 6.835, h: 0.303, fontFace: F.semi, fontSize: 12, color: DKGREY, bold: true, align: 'center' });
  txt(s, [{ text: 'Thank', options: { color: NAVY } }, { text: ' ', options: { color: INK } }, { text: 'You', options: { color: RED } }], { x: 2.004, y: 3.511, w: 9.313, h: 1.447, fontFace: F.black, fontSize: 80, bold: true, align: 'center' });
}

// ---------------------------------------------------------------- build
const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
  slide31, slide32, slide33, slide34, slide35, slide36, slide37];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'WIDE';
pptx.author = 'Orion Express';
pptx.title = 'Orion Express - Logistic & Shipping Service Presentation';

SLIDES.forEach((build) => build(pptx.addSlide()));

const outFile = path.join(__dirname, path.basename(__filename, '.js') + '.pptx');
pptx.writeFile({ fileName: outFile }).then(() => console.log('wrote ' + outFile));
