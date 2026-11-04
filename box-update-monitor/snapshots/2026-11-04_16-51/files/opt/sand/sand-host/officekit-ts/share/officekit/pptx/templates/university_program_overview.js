/**
 * "University Program" template deck - recreated with pptxgenjs.
 *
 * Layout notes
 *  - slide size 13.333 x 7.5 in (16:9); all coordinates below are inches.
 *  - the source deck leans on 50-degree linear gradients, but pptxgenjs only emits
 *    solid fills, so those are painted as grids/rings of solid cells (see diagFill,
 *    bands and shadeOval) sampled from the ramps in the palette block.
 *  - every photo / device mock-up is redrawn from native shapes; nothing is embedded.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const C = {
  orange: 'FF8825', // theme accent1
  orange2: 'FE7A36', // theme accent6 - icon badges
  blue: '3468C0', // theme accent4 - cards
  blueLt: '86A7FC', // theme accent3
  blueDk: '274E90',
  white: 'FFFFFF',
  black: '000000',
  ink: '404040',
  grey: 'E5E5E5',
  greyArch: 'E6E6E6',
  imgFill: 'DEDEDE',
  imgText: '9A9A9A',
  peachPanel: 'FF9740' // accent1 -> FFB87C flattened (used behind the banded fills)
};

// Direction cosines of the deck's 50-degree gradient axis.
const GX = Math.cos((50 * Math.PI) / 180);
const GY = Math.sin((50 * Math.PI) / 180);

// The deck's recurring orange gradients, as ramp stop lists along that axis.
const PEACH_RAMP = [[0.377, C.orange], [1, 'FFB87C']]; // panels and cards
const SOFT_RAMP = [[0.45, 'FFB87C'], [1, C.white]]; // peach blobs and the cover disc
const WARM_RAMP = [[0, C.orange], [1, 'FFB87C']]; // closing-slide disc

const F = { head: 'Poppins SemiBold', body: 'Open Sans' };

const SHADOW = {
  card: { type: 'outer', color: '000000', opacity: 0.43, blur: 20, offset: 4, angle: 90 },
  cardHi: { type: 'outer', color: '000000', opacity: 0.56, blur: 21, offset: 4, angle: 90 },
  glow: { type: 'outer', color: '000000', opacity: 0.14, blur: 60, offset: 0, angle: 0 },
  bloom: { type: 'outer', color: '000000', opacity: 0.2, blur: 60, offset: 0, angle: 0 },
  table: { type: 'outer', color: '7F7F7F', opacity: 0.1, blur: 20, offset: 3, angle: 270 }
};

/* ---------------------------------------------------------------- gradients */

/** Sample a stop list [[pos 0..1, hex], ...] at u, clamping outside the ends. */
function ramp(stops, u) {
  let a = stops[0], b = stops[stops.length - 1];
  for (let k = 0; k < stops.length - 1; k++) {
    if (u >= stops[k][0] && u <= stops[k + 1][0]) { a = stops[k]; b = stops[k + 1]; break; }
  }
  const t = b[0] === a[0] ? 0 : Math.max(0, Math.min(1, (u - a[0]) / (b[0] - a[0])));
  let hex = '';
  for (let ch = 0; ch < 3; ch++) {
    const av = parseInt(a[1].substr(ch * 2, 2), 16);
    const bv = parseInt(b[1].substr(ch * 2, 2), 16);
    hex += ('0' + Math.round(av + (bv - av) * t).toString(16)).slice(-2).toUpperCase();
  }
  return hex;
}

/**
 * Paint a left-to-right gradient as vertical solid bands.
 * `span` = {x, w} lets several boxes borrow colours from one wider ramp so they line up.
 */
function bands(slide, x, y, w, h, stops, span) {
  const sp = span || { x: x, w: w };
  const n = Math.max(8, Math.round(w * 14));
  for (let i = 0; i < n; i++) {
    const bx = x + (w * i) / n;
    slide.addShape('rect', {
      x: bx, y: y, w: w / n + 0.02, h: h,
      fill: { color: ramp(stops, (bx + w / (2 * n) - sp.x) / sp.w) }, line: { type: 'none' }
    });
  }
}

/**
 * Paint the deck's 50-degree orange gradient over a rectangle as a grid of solid cells
 * (pptxgenjs emits solid fills only).
 *   from  - corner holding the first stop: 'tl' | 'tr' | 'bl' | 'br'
 *   ang   - gradient axis in degrees, measured down from the +x axis
 *   clips - sub-rectangles to paint; defaults to the whole box
 */
function diagFill(slide, x, y, w, h, stops, from, ang, clips) {
  const gx = ang === undefined ? GX : Math.cos((ang * Math.PI) / 180);
  const gy = ang === undefined ? GY : Math.sin((ang * Math.PI) / 180);
  const total = w * gx + h * gy;
  const fromRight = (from || 'tl').indexOf('r') >= 0;
  const fromBottom = (from || 'tl').indexOf('b') >= 0;
  (clips || [{ x: x, y: y, w: w, h: h }]).forEach(function (box) {
    const nx = Math.max(2, Math.round(Math.sqrt((300 * box.w) / box.h)));
    const ny = Math.max(2, Math.round(300 / nx));
    for (let iy = 0; iy < ny; iy++) {
      for (let ix = 0; ix < nx; ix++) {
        const cx = box.x + (box.w * (ix + 0.5)) / nx;
        const cy = box.y + (box.h * (iy + 0.5)) / ny;
        const u = ((fromRight ? x + w - cx : cx - x) * gx + (fromBottom ? y + h - cy : cy - y) * gy) / total;
        slide.addShape('rect', {
          x: box.x + (box.w * ix) / nx, y: box.y + (box.h * iy) / ny,
          w: box.w / nx + 0.02, h: box.h / ny + 0.02,
          fill: { color: ramp(stops, u) }, line: { type: 'none' }
        });
      }
    }
  });
}

/** Everything inside a rounded rect except the four corner squares of radius r. */
function roundClips(x, y, w, h, r) {
  return [{ x: x, y: y + r, w: w, h: h - 2 * r }, { x: x + r, y: y, w: w - 2 * r, h: h }];
}

/**
 * Approximate the 50-degree gradient inside a circle: draw the disc in the dark tone,
 * then stack shrinking discs tangent to the light end, each one tinted by the ramp
 * position of its trailing edge.
 */
function shadeOval(slide, x, y, d, stops) {
  const half = 1 / (2 * (GX + GY)); // half of the circle's span along the gradient axis
  const steps = 30;
  for (let i = 0; i <= steps; i++) {
    const k = 1 - (i / steps) * 0.97;
    slide.addShape('ellipse', {
      x: x + ((1 - k) * d * (1 + GX)) / 2, y: y + ((1 - k) * d * (1 + GY)) / 2,
      w: d * k, h: d * k,
      fill: { color: ramp(stops, 0.5 + half * (1 - 2 * k)) }, line: { type: 'none' }
    });
  }
}

/* ------------------------------------------------------------------ helpers */

function text(slide, body, opts) {
  slide.addText(body, Object.assign({ fontFace: F.body, fontSize: 12, color: C.black }, opts));
}

/** Turn a list of strings into hard-broken paragraphs of one text object. */
function lines(arr, opts) {
  return arr.map(function (t, i) {
    return { text: t, options: Object.assign({ breakLine: i < arr.length - 1 }, opts || {}) };
  });
}

/** Rounded rectangle; `adj` is the OOXML corner adjust value / 100000. */
function panel(slide, x, y, w, h, fill, adj, extra) {
  slide.addShape('roundRect', Object.assign({
    x: x, y: y, w: w, h: h,
    fill: { color: fill }, line: { type: 'none' },
    rectRadius: (adj === undefined ? 0.16667 : adj) * Math.min(w, h)
  }, extra || {}));
}

/**
 * "Rectangle: single corner rounded" accent panel (flip = 'H' | 'V' | 'HV').
 * round1Rect rounds the top-right corner; the flips move it to the other three.
 * The gradient's first (dark) stop sits in the corner opposite the rounded one.
 */
function cornerPanel(slide, x, y, w, h, adj, flip, grad) {
  const r = adj * Math.min(w, h);
  const flipH = !!(flip && flip.indexOf('H') >= 0);
  const flipV = !!(flip && flip.indexOf('V') >= 0);
  const g = grad || {};
  slide.addShape('round1Rect', {
    x: x, y: y, w: w, h: h,
    fill: { color: ramp(PEACH_RAMP, 0.5) }, line: { type: 'none' },
    rectRadius: r, flipH: flipH, flipV: flipV
  });
  const from = g.from || (flipV ? 'b' : 't') + (flipH ? 'r' : 'l');
  diagFill(slide, x, y, w, h, PEACH_RAMP, from, g.ang, [
    { x: x, y: flipV ? y : y + r, w: w, h: h - r },
    { x: flipH ? x + r : x, y: flipV ? y + h - r : y, w: w - r, h: r }
  ]);
}

/** Soft peach decoration bubble. */
function bubble(slide, x, y, d) {
  shadeOval(slide, x, y, d, SOFT_RAMP);
}

/** Short orange rule used as a footer flourish. */
function rule(slide, x, y) {
  slide.addShape('line', { x: x, y: y, w: 0.301, h: 0, line: { color: C.orange, width: 3.25 } });
}

/* -------------------------------------------------------------------- icons */

/**
 * Line-art marks standing in for the deck's SVG glyphs. Each part is
 * [shape, x, y, w, h, flag] with coordinates in 0..1 of the icon box;
 * flag 'fill' paints the part solid, 'V' flips it vertically.
 */
const ICONS = {
  school: [['line', 0.50, 0.00, 0, 0.30], ['rect', 0.50, 0.02, 0.20, 0.13, 'fill'],
    ['triangle', 0.22, 0.26, 0.56, 0.28], ['rect', 0.24, 0.50, 0.52, 0.50],
    ['rect', 0.00, 0.62, 0.24, 0.38], ['rect', 0.76, 0.62, 0.24, 0.38],
    ['roundRect', 0.40, 0.72, 0.20, 0.28]],
  searchBook: [['rect', 0.02, 0.22, 0.96, 0.70], ['line', 0.50, 0.22, 0, 0.70],
    ['ellipse', 0.30, 0.06, 0.40, 0.40], ['line', 0.62, 0.40, 0.22, 0.26]],
  backpack: [['roundRect', 0.14, 0.16, 0.72, 0.80], ['roundRect', 0.36, 0.02, 0.28, 0.22],
    ['roundRect', 0.30, 0.56, 0.40, 0.30], ['rect', 0.00, 0.46, 0.16, 0.30],
    ['rect', 0.84, 0.46, 0.16, 0.30]],
  book: [['rect', 0.02, 0.14, 0.96, 0.74], ['line', 0.50, 0.14, 0, 0.74],
    ['line', 0.10, 0.24, 0.32, 0.06], ['line', 0.58, 0.30, 0.32, -0.06]],
  brain: [['ellipse', 0.02, 0.06, 0.96, 0.76], ['line', 0.34, 0.78, 0.06, 0.20],
    ['line', 0.52, 0.78, -0.06, 0.20], ['line', 0.30, 0.40, 0.40, 0]],
  printer: [['roundRect', 0.24, 0.00, 0.52, 0.22], ['roundRect', 0.02, 0.22, 0.96, 0.46],
    ['rect', 0.24, 0.62, 0.52, 0.38], ['line', 0.34, 0.76, 0.32, 0], ['line', 0.34, 0.88, 0.32, 0]],
  clip: [['roundRect', 0.30, 0.00, 0.40, 1.00], ['roundRect', 0.42, 0.14, 0.16, 0.62]],
  doc: [['rect', 0.16, 0.00, 0.68, 1.00], ['line', 0.30, 0.24, 0.40, 0],
    ['line', 0.30, 0.40, 0.40, 0], ['line', 0.30, 0.56, 0.40, 0], ['line', 0.30, 0.72, 0.24, 0]],
  wifi: [['arc', 0.00, 0.06, 1.00, 0.92], ['arc', 0.20, 0.34, 0.60, 0.56],
    ['ellipse', 0.38, 0.70, 0.24, 0.24]],
  ruler: [['triangle', 0.00, 0.00, 1.00, 1.00, 'HV'], ['line', 0.10, 0.30, 0.16, 0],
    ['line', 0.10, 0.52, 0.16, 0], ['line', 0.36, 0.14, 0, 0.16], ['line', 0.58, 0.14, 0, 0.16]],
  person: [['ellipse', 0.34, 0.00, 0.32, 0.28], ['triangle', 0.20, 0.28, 0.60, 0.42],
    ['rect', 0.34, 0.62, 0.10, 0.38], ['rect', 0.56, 0.62, 0.10, 0.38]],
  phone: [['roundRect', 0.28, 0.00, 0.44, 1.00], ['line', 0.42, 0.86, 0.16, 0]],
  mail: [['rect', 0.00, 0.18, 1.00, 0.64], ['line', 0.02, 0.20, 0.48, 0.34],
    ['line', 0.98, 0.20, -0.48, 0.34]],
  home: [['triangle', 0.00, 0.00, 1.00, 0.52], ['rect', 0.16, 0.50, 0.68, 0.50]]
};

function checkMark(slide, x, y, s, color) {
  const t = s * 0.17;
  slide.addShape('rect', { x: x + s * 0.1, y: y + s * 0.48, w: s * 0.4, h: t, fill: { color: color }, line: { type: 'none' }, rotate: 45 });
  slide.addShape('rect', { x: x + s * 0.32, y: y + s * 0.3, w: s * 0.62, h: t, fill: { color: color }, line: { type: 'none' }, rotate: -45 });
}

function icon(slide, kind, x, y, s, color, filled) {
  if (kind === 'check') { checkMark(slide, x, y, s, color); return; }
  const w = Math.max(0.6, s * 3.4);
  ICONS[kind].forEach(function (p) {
    const o = { x: x + p[1] * s, y: y + p[2] * s, w: p[3] * s, h: p[4] * s, line: { color: color, width: w } };
    o.fill = filled || p[5] === 'fill' ? { color: color } : { type: 'none' };
    if (p[5] && p[5].indexOf('H') >= 0) o.flipH = true;
    if (p[5] && p[5].indexOf('V') >= 0) o.flipV = true;
    if (p[0] === 'arc') o.angleRange = [180, 360];
    slide.addShape(p[0], o);
  });
}

/** Filled disc with a centred icon; `ring` adds the white outline badge style. */
function iconDisc(slide, x, y, d, discColor, kind, iconColor, ring) {
  slide.addShape('ellipse', {
    x: x, y: y, w: d, h: d,
    fill: { color: discColor },
    line: ring ? { color: C.white, width: 4.75 } : { type: 'none' }
  });
  icon(slide, kind, x + d * 0.2, y + d * 0.2, d * 0.6, iconColor);
}

/** The orange/white badge that floats over most slides. */
function badge(slide, x, y, d) {
  iconDisc(slide, x, y, d === undefined ? 0.852 : d, C.orange2, 'school', C.white, true);
}

/* --------------------------------------------------------- shared furniture */

function brandMark(slide, x, y, light) {
  icon(slide, 'searchBook', x, y, 0.295, light ? C.white : C.orange);
  text(slide, [
    { text: 'University ', options: { color: light ? C.white : C.orange, italic: true } },
    { text: 'Template', options: { color: light ? C.white : C.black } }
  ], { x: x + 0.233, y: y + 0.021, w: 1.749, h: 0.268, fontSize: 11, bold: true, lineSpacingMultiple: 0.9 });
}

/**
 * Kicker line + big two-tone title.
 * title = array of paragraphs, each an array of [text, color?] runs.
 */
function heading(slide, o) {
  const base = o.light ? C.white : C.black;
  text(slide, o.kicker, {
    x: o.kx === undefined ? o.x + 0.03 : o.kx, y: o.ky, w: o.kw || 3.812, h: 0.337,
    fontSize: 14, color: base, align: o.kalign || o.align || 'left'
  });
  const runs = [];
  o.title.forEach(function (para, i) {
    para.forEach(function (run, j) {
      runs.push({
        text: run[0],
        options: { color: run[1] || base, breakLine: j === para.length - 1 && i < o.title.length - 1 }
      });
    });
  });
  text(slide, runs, {
    x: o.x, y: o.y, w: o.w, h: o.h, fontSize: o.size || 44, bold: true,
    fontFace: F.head, lineSpacingMultiple: 0.9, align: o.align || 'left'
  });
}

/** Orange disc + "Text Here" caption + paragraph, the recurring left column. */
function noteBlock(slide, o) {
  iconDisc(slide, o.x, o.y, 0.673, C.orange, o.icon || 'backpack', C.white);
  let ty = o.by;
  if (ty === undefined) {
    ty = o.title ? o.y + 0.417 : o.y + 0.777;
  }
  if (o.title) text(slide, o.title, { x: o.tx, y: o.ty === undefined ? o.y : o.ty, w: o.tw || 2.993, h: 0.337, fontSize: 14, bold: true, color: C.orange });
  text(slide, o.body, { x: o.bx === undefined ? o.tx : o.bx, y: ty, w: o.bw || 2.993, h: o.bh || 1.28, fontSize: 12, lineSpacingMultiple: 1.5 });
}

/** Blue card: rounded panel + white icon disc + bold heading + body copy. */
function iconCard(slide, c) {
  panel(slide, c.x, c.y, c.w, c.h, C.blue, 0.14713, { shadow: SHADOW.card });
  iconDisc(slide, c.dx, c.dy, 0.683, C.white, c.icon, C.blue);
  text(slide, c.title, { x: c.tx, y: c.ty, w: c.tw, h: 0.364, fontSize: 14, bold: true, color: C.white, lineSpacingMultiple: 1.2 });
  text(slide, c.body, { x: c.tx, y: c.ty + (c.gap || 0.374), w: c.bw, h: c.bh || 0.675, fontSize: 12, color: C.white, lineSpacingMultiple: 1.5 });
}

/** Blue "stat" card: big percentage + italic caption (+ optional body copy). */
function statCard(slide, c) {
  panel(slide, c.x, c.y, c.w, c.h, C.blue, 0.10912, { shadow: SHADOW.card });
  const al = c.align || 'left';
  text(slide, c.value, { x: c.vx, y: c.vy, w: c.vw, h: 0.59, fontSize: c.vsize || 32, bold: true, fontFace: F.head, color: C.white, align: al, lineSpacingMultiple: 0.9 });
  text(slide, c.caption, { x: c.cx, y: c.cy, w: c.cw, h: 0.367, fontSize: c.csize || 14, italic: true, color: C.white, align: al, lineSpacingMultiple: 1.2 });
  if (c.body) text(slide, c.body, { x: c.bx, y: c.by, w: c.bw, h: c.bh || 0.978, fontSize: 12, color: C.white, lineSpacingMultiple: 1.5 });
}

/** "01." + paragraph list rows. */
function numberedItem(slide, o) {
  text(slide, o.num, { x: o.x, y: o.y, w: o.nw || 0.775, h: 0.529, fontSize: 28, bold: true, fontFace: F.head, color: C.orange, lineSpacingMultiple: 0.9 });
  text(slide, o.body, { x: o.bx, y: o.y, w: o.bw, h: 0.978, fontSize: 12, lineSpacingMultiple: 1.5 });
}

/** Grey stand-in for a photo / device mock-up from the original deck. */
function imagePlaceholder(slide, x, y, w, h, opts) {
  const o = opts || {};
  slide.addShape(o.shape || 'rect', Object.assign({
    x: x, y: y, w: w, h: h,
    fill: { color: o.fill || C.imgFill }, line: { type: 'none' }
  }, o.rectRadius ? { rectRadius: o.rectRadius } : {}));
  if (o.label !== false) {
    text(slide, '[image]', {
      x: x, y: y + h / 2 - 0.18, w: w, h: 0.36,
      fontSize: 12, color: o.labelColor || C.imgText, align: 'center'
    });
  }
}

/* ------------------------------------------------------------- slide bodies */

// 1 - title / cover
function slide01(s) {
  shadeOval(s, 6.055, 0.46, 6.547, SOFT_RAMP);
  brandMark(s, 0.823, 0.46);
  heading(s, {
    kicker: 'Specializing in Different Areas', kx: 0.732, ky: 1.553,
    x: 0.702, y: 1.923, w: 4.124, h: 1.744, size: 54,
    title: [[['University ']], [['Program', C.orange]]]
  });
  noteBlock(s, { x: 0.835, y: 4.306, tx: 1.727, bx: 1.727, by: 4.306, bw: 3.439, body: 'PLACEHOLDER' });
  text(s, 'WWW. YOURWEBSITE. COM', { x: 0.731, y: 6.082, w: 3.29, h: 0.372, fontSize: 12, color: C.orange, charSpacing: 3, lineSpacingMultiple: 1.5 });
  rule(s, 0.845, 6.809);
  badge(s, 6.156, 1.589);
  s.addShape('donut', { x: 11.489, y: 5.199, w: 0.938, h: 0.938, fill: { color: C.orange }, line: { type: 'none' } });
  statCard(s, {
    x: 10.803, y: 2.015, w: 2.059, h: 1.058,
    value: '+11,8%', vx: 11.43, vy: 2.196, vw: 1.299, vsize: 24,
    caption: 'Text Tittle', cx: 11.439, cy: 2.616, cw: 1.29, csize: 12
  });
  icon(s, 'printer', 11.035, 2.163, 0.413, C.white);
  panel(s, 5.463, 4.964, 2.059, 1.058, C.blue, 0.14309, { shadow: SHADOW.card });
  icon(s, 'brain', 5.657, 5.2, 0.413, C.white);
  text(s, 'At vero eos et accusamus ', { x: 6.063, y: 5.2, w: 1.327, h: 0.627, fontSize: 11, color: C.white, lineSpacingMultiple: 1.5 });
}

// 2 - table of content
function slide02(s) {
  const cards = [
    { n: '01', fill: C.peachPanel, shadow: null },
    { n: '02', fill: C.blue, shadow: SHADOW.cardHi },
    { n: '03', fill: C.peachPanel, shadow: null },
    { n: '04', fill: C.peachPanel, shadow: null }
  ];
  cards.forEach(function (c, i) {
    const x = 0.839 + i * 2.933;
    panel(s, x, 3.015, 2.856, 3.075, c.fill, 0.11043, c.shadow ? { shadow: c.shadow } : {});
    if (c.fill === C.peachPanel) {
      diagFill(s, x, 3.015, 2.856, 3.075, PEACH_RAMP, 'tl', 50, roundClips(x, 3.015, 2.856, 3.075, 0.34));
    }
    text(s, c.n, { x: x + 0.344, y: 3.468, w: 0.912, h: 0.64, fontSize: 32, bold: true, color: C.white });
    text(s, 'Heading text ', { x: x + 0.344, y: 4.07, w: 1.941, h: 0.337, fontSize: 14, bold: true, italic: true, color: C.white });
    text(s, 'At vero eos et accusamus et iusto odio dignissimos', { x: x + 0.344, y: 4.849, w: 2.28, h: 0.678, fontSize: 12, color: C.white, lineSpacingMultiple: 1.5 });
  });
  brandMark(s, 0.823, 0.46);
  badge(s, 11.213, 2.548);
  heading(s, {
    kicker: 'Discover Your Path to Success', kx: 4.761, ky: 1.162, align: 'center',
    x: 1.599, y: 1.64, w: 10.136, h: 0.774,
    title: [[['Table Of '], ['Content', C.orange]]]
  });
}

// 3 - about university program
function slide03(s) {
  cornerPanel(s, 5.225, 0, 2.622, 4.756, 0.27523, 'HV');
  brandMark(s, 0.823, 0.46);
  heading(s, {
    kicker: 'Specializing in Different Areas', kx: 0.732, ky: 1.609,
    x: 0.702, y: 1.979, w: 4.124, h: 2.106,
    title: [[['About University '], ['Program', C.orange]]]
  });
  noteBlock(s, { x: 0.835, y: 4.594, tx: 1.675, title: 'Text Here', bh: 1.583, body: 'PLACEHOLDER' });
  badge(s, 10.769, 1.467);
  bubble(s, 12.194, 4.377, 0.601);
  rule(s, 12.194, 6.727);
  statCard(s, {
    x: 5.84, y: 4.052, w: 2.993, h: 2.809,
    value: '+46,8%', vx: 6.239, vy: 4.497, vw: 2.231,
    caption: 'Text Here', cx: 6.249, cy: 5.094, cw: 2.508,
    body: 'PLACEHOLDER', bx: 6.249, by: 5.438, bw: 2.231
  });
}

// 4 - service university program
function slide04(s) {
  cornerPanel(s, 3.344, 2.263, 2.888, 4.463, 0.27523, 'V');
  brandMark(s, 0.823, 0.46);
  heading(s, {
    kicker: 'Discover Your Path to Success', kx: 8.38, ky: 1.08,
    x: 8.35, y: 1.45, w: 4.124, h: 2.106,
    title: [[['Service University '], ['Program', C.orange]]]
  });
  numberedItem(s, { num: '01.', x: 8.361, y: 4.186, bx: 9.099, bw: 3.092, body: 'PLACEHOLDER' });
  numberedItem(s, { num: '02.', x: 8.361, y: 5.569, nw: 0.871, bx: 9.099, bw: 2.911, body: 'PLACEHOLDER' });
  bubble(s, 12.186, 0.947, 0.601);
  rule(s, 12.194, 6.727);
  badge(s, 2.98, 1.006);
  iconCard(s, {
    x: 3.096, y: 4.567, w: 3.728, h: 1.827,
    dx: 3.472, dy: 4.841, icon: 'clip',
    tx: 4.317, ty: 4.841, tw: 2.508, gap: 0.343,
    title: 'Text Here', body: 'PLACEHOLDER', bw: 2.231, bh: 0.978
  });
}

// 5 - teachers grid
function slide05(s) {
  // rotated 90 degrees in the source, so its gradient axis swaps x/y and keeps the dark corner top-left
  cornerPanel(s, 5.837, 0.003, 7.5, 7.493, 0.15312, 'HV', { from: 'tl', ang: 40 });
  brandMark(s, 0.823, 0.46);
  heading(s, {
    kicker: 'Discover Your Path to Success', kx: 0.734, ky: 1.469,
    x: 0.704, y: 1.839, w: 4.124, h: 2.106,
    title: [[['Teachers University '], ['Program', C.orange]]]
  });
  noteBlock(s, { x: 0.835, y: 4.702, tx: 1.675, ty: 4.841, by: 5.23, title: 'Text Here', icon: 'book', body: 'PLACEHOLDER' });
  for (let r = 0; r < 2; r++) {
    for (let c = 0; c < 3; c++) {
      const y = 1.092 + r * 3.007;
      badge(s, 6.699 + c * 2.139, y - 0.008, 0.553);
      text(s, 'Name Here', { x: 6.699 + c * 2.098, y: y + 1.702, w: 1.58, h: 0.337, fontSize: 14, bold: true, color: C.white, align: 'center' });
      text(s, 'Text Tittle Here', { x: 6.807 + c * 2.098, y: y + 1.96, w: 1.364, h: 0.349, fontSize: 10.5, italic: true, color: C.white, align: 'center', lineSpacingMultiple: 1.5 });
    }
  }
}

// 6 - benefits of higher education
function slide06(s) {
  cornerPanel(s, 9.816, 0.521, 2.913, 4.463, 0.27523, 'V');
  brandMark(s, 0.823, 0.46);
  heading(s, {
    kicker: 'Discover Your Path to Success', kx: 0.734, ky: 1.187,
    x: 0.704, y: 1.558, w: 4.124, h: 2.106,
    title: [[['Benefits']], [['of Higher '], ['Education', C.orange]]]
  });
  noteBlock(s, { x: 0.835, y: 4.171, tx: 0.76, by: 4.948, bw: 3.439, body: 'PLACEHOLDER' });
  rule(s, 0.845, 6.809);
  badge(s, 10.762, 1.723);
  [{ y: 2.119, icon: 'clip' }, { y: 4.276, icon: 'brain' }].forEach(function (c) {
    iconCard(s, {
      x: 4.853, y: c.y, w: 3.624, h: 1.827,
      dx: 5.229, dy: c.y + 0.273, icon: c.icon,
      tx: 6.074, ty: c.y + 0.273, tw: 2.508, gap: 0.343,
      title: 'Text Here', body: 'PLACEHOLDER', bw: 2.231, bh: 0.978
    });
  });
  statCard(s, {
    x: 10.488, y: 4.37, w: 1.656, h: 1.611, align: 'center',
    value: '+46%', vx: 10.555, vy: 4.744, vw: 1.521,
    caption: 'Text Here', cx: 10.66, cy: 5.281, cw: 1.312, csize: 12
  });
  bubble(s, 6.771, 0.853, 0.601);
}

// 7 - why choose a university program
function slide07(s) {
  cornerPanel(s, 0, 0, 5.208, 3.588, 0.16703, 'V');
  brandMark(s, 0.823, 0.46, true);
  heading(s, {
    kicker: 'Discover Your Path to Success', kx: 7.744, ky: 1.177,
    x: 7.715, y: 1.547, w: 5.0, h: 2.106,
    title: [[['Why Choose']], [['a '], ['University', C.orange], [' Program?']]]
  });
  bubble(s, 12.178, 0.908, 0.537);
  // small blue "check" stat, lower-left
  iconDisc(s, 0.863, 5.182, 0.395, C.blue, 'check', C.white);
  text(s, '+21,5%', { x: 1.414, y: 5.132, w: 1.607, h: 0.468, fontSize: 24, bold: true, fontFace: F.head, color: C.blue, lineSpacingMultiple: 0.9 });
  text(s, 'Text Here', { x: 1.407, y: 5.555, w: 1.428, h: 0.326, fontSize: 12, italic: true, lineSpacingMultiple: 1.2 });
  text(s, lines(['At vero eos et accusamus et iusto', 'odio dignissimos']), { x: 1.407, y: 5.841, w: 2.08, h: 0.978, fontSize: 12, lineSpacingMultiple: 1.5 });
  numberedItem(s, { num: '01.', x: 7.726, y: 4.046, bx: 8.464, bw: 3.714, body: lines(['PLACEHOLDER', 'voluptatum deleniti atque corrupti quos'], { breakLine: true }) });
  numberedItem(s, { num: '02.', x: 7.726, y: 5.429, nw: 0.871, bx: 8.464, bw: 3.731, body: lines(['PLACEHOLDER', 'voluptatum deleniti atque corrupti quos'], { breakLine: true }) });
  rule(s, 12.194, 6.727);
  panel(s, 3.375, 3.001, 1.656, 1.611, C.blue, 0.10912, { shadow: SHADOW.card });
  icon(s, 'searchBook', 3.89, 3.344, 0.627, C.white);
  text(s, 'Text Here', { x: 3.547, y: 3.943, w: 1.312, h: 0.326, fontSize: 12, italic: true, color: C.white, align: 'center', lineSpacingMultiple: 1.2 });
  badge(s, 5.692, 5.691);
}

// 8 - a guide to university programs
function slide08(s) {
  cornerPanel(s, 0, 3.762, 3.832, 3.738, 0.16804, 'H');
  brandMark(s, 0.823, 0.46);
  // white-on-white heading kept from the source deck (sits under the visible title)
  text(s, 'Encouraging Nutritional Education', { x: 8.358, y: 1.43, w: 4.009, h: 0.37, fontSize: 16, color: C.white });
  text(s, lines(['Public Health', 'Awareness']), { x: 8.35, y: 1.834, w: 4.017, h: 1.447, fontSize: 40, color: C.white, fontFace: F.head, lineSpacingMultiple: 1.0 });
  heading(s, {
    kicker: 'Discover Your Path to Success', kx: 8.377, ky: 1.279,
    x: 8.348, y: 1.713, w: 4.124, h: 2.106,
    title: [[['A Guide to '], ['University', C.orange], [' Programs']]]
  });
  noteBlock(s, { x: 8.489, y: 4.267, tx: 8.38, by: 5.172, bw: 3.798, bh: 1.583, body: 'Choosing a university program is a significant step in shaping your future. This presentation will provide you with a comprehensive overview of different university programs, helping you make an informed decision.' });
  bubble(s, 12.178, 0.908, 0.537);
  rule(s, 12.194, 6.727);
  iconDisc(s, 0.692, 1.716, 0.673, C.orange, 'searchBook', C.white);
  text(s, 'Text Here', { x: 1.532, y: 1.716, w: 2.147, h: 0.337, fontSize: 14, bold: true });
  text(s, 'At vero eos et accusamus et iusto odio dignissimos ducimus qui', { x: 1.532, y: 2.077, w: 2.202, h: 0.978, fontSize: 12, lineSpacingMultiple: 1.5 });
  panel(s, 3.489, 3.543, 3.993, 1.342, C.blue);
  [['86%', 3.903, 3.877], ['35k', 5.638, 5.612]].forEach(function (v) {
    text(s, v[0], { x: v[1], y: 3.758, w: 1.431, h: 0.64, fontSize: 32, bold: true, color: C.white, align: 'center' });
    text(s, 'Text here ', { x: v[2], y: 4.333, w: 1.483, h: 0.337, fontSize: 14, italic: true, color: C.white, align: 'center' });
  });
  bubble(s, 6.664, 5.799, 0.537);
}

// 9 - why choose a university program (3 cards)
function slide09(s) {
  bands(s, 0, 0, 13.333, 7.5, [[0, C.white], [1, C.orange]]);
  brandMark(s, 0.823, 0.46);
  heading(s, {
    kicker: 'Benefits of Higher Education', kx: 0.734, ky: 1.431,
    x: 0.704, y: 1.866, w: 4.124, h: 1.924, size: 40,
    title: [[['Why Choose a '], ['University', C.orange], [' Program?']]]
  });
  text(s, 'PLACEHOLDER', { x: 0.748, y: 4.41, w: 3.074, h: 1.583, fontSize: 12, lineSpacingMultiple: 1.5 });
  rule(s, 0.845, 6.809);
  [
    { y: 1.184, icon: 'clip', title: 'Expand your knowledge', body: 'Delve deeper into your chosen field of study.', bw: 2.195 },
    { y: 3.001, icon: 'printer', title: 'Develop critical skills', body: 'Learn to analyze information and solve problems', bw: 2.518 },
    { y: 4.818, icon: 'searchBook', title: 'Enhance career prospects', body: 'Gain the qualifications needed for your desired career.', bw: 2.615 }
  ].forEach(function (c) {
    iconCard(s, {
      x: 8.514, y: c.y, w: 3.977, h: 1.519,
      dx: 8.824, dy: c.y + 0.243, icon: c.icon,
      tx: 9.642, ty: c.y + 0.243, tw: 2.749,
      title: c.title, body: c.body, bw: c.bw
    });
  });
  badge(s, 4.488, 3.391);
}

// 10 - types of university programs
function slide10(s) {
  brandMark(s, 0.823, 0.46);
  heading(s, {
    kicker: 'Find the Right Fit for You', kx: 0.734, ky: 1.431,
    x: 0.704, y: 1.866, w: 4.124, h: 1.924, size: 40,
    title: [[['Types of '], ['University', C.orange], [' Programs']]]
  });
  [
    { y: 0.819, icon: 'clip', title: 'Undergraduate programs', body: "Bachelor's degrees (e.g., BA, BSc)" },
    { y: 2.34, icon: 'doc', title: 'Graduate programs', body: "Master's and doctoral degrees" },
    { y: 3.86, icon: 'wifi', title: 'Professional programs', body: 'Law, medicine, engineering' },
    { y: 5.381, icon: 'book', title: 'Vocational programs', body: 'Certificates and diplomas for  skills', bw: 2.972 }
  ].forEach(function (c) {
    iconCard(s, {
      x: 4.511, y: c.y, w: 4.314, h: 1.295,
      dx: 4.821, dy: c.y + 0.307, icon: c.icon,
      tx: 5.639, ty: c.y + 0.275, tw: 2.749,
      title: c.title, body: c.body, bw: c.bw || 2.872, gap: 0.374, bh: 0.372
    });
  });
  text(s, 'PLACEHOLDER', { x: 0.748, y: 4.41, w: 3.074, h: 1.583, fontSize: 12, lineSpacingMultiple: 1.5 });
  rule(s, 0.845, 6.809);
  badge(s, 11.955, 0.906);
}

// 11 - pricing table
function slide11(s) {
  panel(s, 0.856, 0.797, 6.323, 5.906, C.white, 0.05107, { shadow: SHADOW.table });
  const money = ['$1,500', '$2,500', '$1,200', '$500', '$800', '$300'];
  const rowH = [0.563, 0.563, 0.686, 0.722, 0.722, 0.722, 0.722, 0.563];
  const none = { type: 'none' };
  const rows = [];
  rows.push(['Text Here', 'Text Here'].map(function (t) {
    return { text: t, options: { fill: { color: C.blue }, color: C.white, border: [{ type: 'solid', color: '7F7F7F', pt: 0.5 }, none, { type: 'solid', color: C.white, pt: 1 }, none] } };
  }));
  money.forEach(function (m, i) {
    const top = i === 0 ? C.white : C.blue;
    const bottom = i === money.length - 1 ? 'FFDD95' : C.blue;
    const cell = { fill: { color: C.white }, color: C.black, border: [{ type: 'solid', color: top, pt: 1 }, none, { type: 'solid', color: bottom, pt: 1 }, none] };
    rows.push([{ text: 'Text Here', options: cell }, { text: m, options: cell }]);
  });
  rows.push(['Total', '$6,300'].map(function (t) {
    return { text: t, options: { fill: { color: C.orange }, color: C.white, bold: true, border: [{ type: 'solid', color: 'FFDD95', pt: 1 }, none, none, none] } };
  }));
  s.addTable(rows, {
    x: 1.193, y: 1.118, w: 5.67, colW: [2.939, 2.731], rowH: rowH,
    fontFace: F.body, fontSize: 12, align: 'center', valign: 'middle', margin: 4
  });
  brandMark(s, 10.658, 0.521);
  heading(s, {
    kicker: 'Benefits of Higher Education', kx: 7.748, ky: 1.25,
    x: 7.718, y: 1.684, w: 4.124, h: 1.924, size: 40,
    title: [[['Why Choose a '], ['University', C.orange], [' Program?']]]
  });
  text(s, 'PLACEHOLDER', { x: 7.764, y: 5.228, w: 4.061, h: 1.28, fontSize: 12, lineSpacingMultiple: 1.5 });
  [['+21,5%', 7.848, 8.168, 8.175], ['+41,8%', 10.261, 10.581, 10.588]].forEach(function (v) {
    iconDisc(s, v[1], 4.273, 0.292, C.blue, 'check', C.white);
    text(s, v[0], { x: v[3], y: 4.185, w: 1.607, h: 0.468, fontSize: 24, bold: true, fontFace: F.head, color: C.blue, lineSpacingMultiple: 0.9 });
    text(s, 'Text Here', { x: v[2], y: 4.608, w: 1.428, h: 0.326, fontSize: 12, italic: true, lineSpacingMultiple: 1.2 });
  });
  bubble(s, 12.208, 1.684, 0.601);
  rule(s, 12.194, 6.727);
}

// 12 - two doughnut charts on a blue banner
function slide12(s, pptx) {
  brandMark(s, 0.823, 0.46);
  heading(s, {
    kicker: 'Discover Your Path to Success', kx: 4.761, ky: 0.932, align: 'center',
    x: 1.599, y: 1.37, w: 10.136, h: 0.774,
    title: [[['Benefits of Higher '], ['Education', C.orange]]]
  });
  [[1.84, 2.782], [7.162, 8.062]].forEach(function (g) {
    icon(s, 'person', g[0] + 0.18, 2.485, 0.36, C.blue, true);
    text(s, 'PLACEHOLDER', { x: g[1], y: 2.502, w: 3.435, h: 0.678, fontSize: 12, lineSpacingMultiple: 1.5 });
  });
  panel(s, 0.681, 3.75, 11.656, 3.131, C.blue, 0.11012);
  [
    { x: 1.386, tx: 2.06, pct: 25, label: '25%', gx: 3.832, vx: 4.158, cx: 4.151, bx: 4.153, val: '+21,5%' },
    { x: 6.975, tx: 7.649, pct: 45, label: '45%', gx: 9.42, vx: 9.747, cx: 9.74, bx: 9.742, val: '+41,5%' }
  ].forEach(function (d) {
    s.addChart(pptx.ChartType.doughnut, [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: [d.pct, 100 - d.pct] }], {
      x: d.x, y: 4.065, w: 2.317, h: 2.474,
      holeSize: 82, showLegend: false, showTitle: false, showValue: false,
      chartColors: [C.orange, C.white], dataBorder: { pt: 0, color: C.blue }, firstSliceAng: 0
    });
    text(s, d.label, { x: d.tx, y: 5.1, w: 0.969, h: 0.404, fontSize: 20, bold: true, color: C.white, align: 'center', lineSpacingMultiple: 0.9 });
    iconDisc(s, d.gx, 4.526, 0.292, C.white, 'check', C.blue);
    text(s, d.val, { x: d.vx, y: 4.438, w: 1.607, h: 0.468, fontSize: 24, bold: true, fontFace: F.head, color: C.white, lineSpacingMultiple: 0.9 });
    text(s, 'Text Here', { x: d.cx, y: 4.861, w: 1.428, h: 0.326, fontSize: 12, italic: true, color: C.white, lineSpacingMultiple: 1.2 });
    text(s, 'PLACEHOLDER', { x: d.bx, y: 5.188, w: 2.205, h: 0.978, fontSize: 12, color: C.white, lineSpacingMultiple: 1.5 });
  });
  bubble(s, 11.96, 0.299, 0.601);
}

// 13 - the application process
function slide13(s) {
  brandMark(s, 0.823, 0.46);
  heading(s, {
    kicker: 'Steps to Take', kx: 9.047, ky: 1.348,
    x: 9.017, y: 1.782, w: 3.61, h: 1.924, size: 40,
    title: [[['The Application '], ['Process', C.orange]]]
  });
  [
    { y: 1.442, w: 4.994, icon: 'clip', title: 'Prepare application materials', tw: 3.25, body: 'Write essays, gather transcripts, and request letters of recommendation.', bw: 3.696 },
    { y: 3.258, w: 4.64, icon: 'printer', title: 'Submit your application', tw: 3.12, body: 'Follow the specific guidelines for each university', bw: 3.414 },
    { y: 5.074, w: 4.01, icon: 'ruler', title: 'Research universities', tw: 2.749, body: 'Explore different institutions and their programs.', bw: 2.749 }
  ].forEach(function (c) {
    iconCard(s, {
      x: 0.846, y: c.y, w: c.w, h: 1.519,
      dx: 1.156, dy: c.y + 0.243, icon: c.icon,
      tx: 1.974, ty: c.y + 0.243, tw: c.tw,
      title: c.title, body: c.body, bw: c.bw
    });
  });
  noteBlock(s, { x: 9.134, y: 4.105, tx: 9.059, by: 4.882, bw: 3.439, body: 'PLACEHOLDER' });
  rule(s, 9.059, 6.632);
  bubble(s, 12.178, 0.908, 0.537);
  badge(s, 7.576, 1.09);
}

// 14 - start your journey today (three doughnuts + stat cards)
function slide14(s, pptx) {
  bands(s, 0, 0, 13.333, 7.5, [[0, C.orange], [0.88, 'FEF2E9'], [1, C.white]]);
  [
    { ox: 4.875, oy: 1.181, od: 3.121, cx: 4.407, cy: 1.388, cw: 4.056, ch: 2.704, pct: 23, lx: 5.76, ly: 2.489, lh: 0.505, ls: 24 },
    { ox: 3.408, oy: 3.943, od: 2.275, cx: 2.99, cy: 4.046, cw: 3.112, ch: 2.074, pct: 57, lx: 3.92, ly: 4.862, lh: 0.438, ls: 20 },
    { ox: 5.992, oy: 4.626, od: 1.791, cx: 5.707, cy: 4.734, cw: 2.361, ch: 1.574, pct: 20, lx: 6.216, ly: 5.336, lh: 0.37, ls: 16 }
  ].forEach(function (d) {
    s.addShape('ellipse', { x: d.ox, y: d.oy, w: d.od, h: d.od, fill: { color: C.white }, line: { type: 'none' }, shadow: SHADOW.glow });
    s.addChart(pptx.ChartType.doughnut, [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: [d.pct, 100 - d.pct] }], {
      x: d.cx, y: d.cy, w: d.cw, h: d.ch,
      holeSize: 63, showLegend: false, showTitle: false, showValue: false,
      chartColors: [C.blue, C.orange], dataBorder: { pt: 0, color: C.white }, firstSliceAng: 0
    });
    text(s, d.pct + '%', { x: d.lx, y: d.ly, w: 1.351, h: d.lh, fontSize: d.ls, bold: true, color: C.ink, align: 'center', valign: 'middle' });
  });
  [
    { x: 1.332, y: 1.75, w: 3.0, h: 1.673, adj: 0.13337, vx: 1.86, vy: 2.131, vw: 1.549, num: '335+', bx: 1.86, by: 2.568, bw: 2.224, bh: 0.579, body: 'At vero eos et accusamus et iusto odio dignissimos' },
    { x: 0.998, y: 4.092, w: 2.101, h: 2.028, adj: 0.10156, vx: 1.382, vy: 4.507, vw: 1.323, num: '565+', bx: 1.382, by: 4.923, bw: 1.475, bh: 0.831, body: 'At vero eos et accusamus et iusto odio dignissimos' },
    { x: 8.334, y: 4.417, w: 3.334, h: 2.0, adj: 0.14785, vx: 8.88, vy: 4.824, vw: 1.606, num: '2335+', bx: 8.88, by: 5.262, bw: 2.454, bh: 0.834, body: 'PLACEHOLDER' }
  ].forEach(function (c) {
    panel(s, c.x, c.y, c.w, c.h, C.blue, c.adj, { shadow: SHADOW.bloom });
    text(s, c.num, { x: c.vx, y: c.vy, w: c.vw, h: 0.457, fontSize: 28, fontFace: F.head, color: C.white, lineSpacingMultiple: 0.7 });
    text(s, c.body, { x: c.bx, y: c.by, w: c.bw, h: c.bh, fontSize: 10, color: C.white, lineSpacingMultiple: 1.5 });
  });
  heading(s, {
    kicker: 'Steps to Take', kx: 8.741, ky: 1.348,
    x: 8.712, y: 1.782, w: 3.61, h: 2.106,
    title: [[['Start Your Journey Today']]]
  });
  brandMark(s, 0.823, 0.46, true);
  bubble(s, 12.268, 0.879, 0.537);
  badge(s, 6.887, 1.21);
}

// 15 - four-quadrant teardrop diagram
function slide15(s) {
  heading(s, {
    kicker: 'Steps to Take', kx: 4.761, ky: 1.063, align: 'center',
    x: 1.333, y: 1.497, w: 10.667, h: 0.774,
    title: [[['Discover Your Path to Success']]]
  });
  brandMark(s, 0.823, 0.46);
  s.addShape('ellipse', { x: 5.619, y: 3.646, w: 2.095, h: 2.095, fill: { color: C.grey }, line: { type: 'none' } });
  const quads = [
    { x: 4.685, y: 2.803, flip: 'H', color: C.orange, icon: 'book', ix: 4.897, iy: 3.015, side: 'l', ty: 3.029, chipX: 2.904, chipY: 4.048, chipTx: 3.026, chipTy: 4.064, chipW: 1.205 },
    { x: 6.844, y: 2.803, flip: '', color: C.blueLt, icon: 'printer', ix: 7.056, iy: 3.015, side: 'r', ty: 3.029, chipX: 9.225, chipY: 4.051, chipTx: 9.347, chipTy: 4.068, chipW: 1.205 },
    { x: 4.685, y: 4.713, flip: 'HV', color: C.blueDk, icon: 'brain', ix: 4.897, iy: 5.023, side: 'l', ty: 5.102, chipX: 2.904, chipY: 6.117, chipTx: 3.032, chipTy: 6.134, chipW: 1.216 },
    { x: 6.844, y: 4.713, flip: 'V', color: C.blue, icon: 'school', ix: 7.056, iy: 5.023, side: 'r', ty: 5.102, chipX: 9.213, chipY: 6.108, chipTx: 9.342, chipTy: 6.125, chipW: 1.216 }
  ];
  quads.forEach(function (q) {
    s.addShape('teardrop', {
      x: q.x, y: q.y, w: 1.804, h: 1.804, fill: { color: q.color }, line: { type: 'none' },
      flipH: q.flip.indexOf('H') >= 0, flipV: q.flip.indexOf('V') >= 0
    });
    iconDisc(s, q.ix, q.iy, 1.38, C.white, q.icon, '231F20');
    const tx = q.side === 'l' ? 0.734 : 9.102;
    const cap = q.side === 'l' ? 2.464 : 9.102;
    const al = q.side === 'l' ? 'right' : 'left';
    text(s, 'Text Here', { x: cap, y: q.ty, w: 1.748, h: 0.337, fontSize: 14, bold: true, fontFace: F.head, align: al });
    text(s, 'PLACEHOLDER', { x: tx, y: q.ty + 0.228, w: 3.477, h: 0.675, fontSize: 12, align: al, lineSpacingMultiple: 1.5 });
    panel(s, q.chipX, q.chipY, q.chipW, 0.337, q.color);
    text(s, 'Text Here', { x: q.chipTx, y: q.chipTy, w: 0.96, h: 0.303, fontSize: 12, color: C.white, align: 'center' });
  });
}

// 16 - find the right fit for you (laptop mock-up)
function slide16(s) {
  cornerPanel(s, 0.839, 1.877, 3.266, 3.186, 0.16804, 'H');
  // laptop mock-up: dark bezel, white screen, tapered base bar
  s.addShape('rect', { x: 1.76, y: 2.65, w: 5.52, h: 3.75, fill: { color: '1B1B1C' }, line: { type: 'none' } });
  s.addShape('rect', { x: 1.94, y: 2.89, w: 5.16, h: 3.23, fill: { color: C.white }, line: { type: 'none' } });
  panel(s, 1.14, 6.4, 6.78, 0.15, 'C9C9CB', 0.5);
  text(s, '[image]', { x: 1.94, y: 4.33, w: 5.16, h: 0.36, fontSize: 12, color: C.imgText, align: 'center' });
  brandMark(s, 0.823, 0.46);
  heading(s, {
    kicker: 'The Application Process', kx: 8.743, ky: 1.39,
    x: 8.714, y: 1.824, w: 3.841, h: 1.924, size: 40,
    title: [[['Find the '], ['Right', C.orange], [' Fit for You']]]
  });
  bubble(s, 11.825, 0.723, 0.537);
  text(s, 'PLACEHOLDER', { x: 8.735, y: 3.789, w: 3.942, h: 1.28, fontSize: 12, lineSpacingMultiple: 1.5 });
  iconDisc(s, 8.832, 5.4, 0.292, C.blue, 'check', C.white);
  text(s, '+21,5%', { x: 9.158, y: 5.312, w: 1.607, h: 0.468, fontSize: 24, bold: true, fontFace: F.head, color: C.blue, lineSpacingMultiple: 0.9 });
  text(s, 'At vero eos et accusamus et iusto odio dignissimos ducimus qui', { x: 9.151, y: 5.735, w: 2.942, h: 0.675, fontSize: 12, lineSpacingMultiple: 1.5 });
  iconCard(s, {
    x: 4.505, y: 1.38, w: 3.624, h: 1.827,
    dx: 4.881, dy: 1.653, icon: 'brain',
    tx: 5.725, ty: 1.653, tw: 2.508, gap: 0.343,
    title: 'Text Here', body: 'PLACEHOLDER', bw: 2.231, bh: 0.978
  });
}

// 17 - start your journey today (orange freeform panel)
function slide17(s) {
  // curved orange panel: straight right edge, bulging left edge
  s.addShape('custGeom', {
    x: 5.865, y: 0, w: 7.469, h: 7.5,
    fill: { color: C.orange }, line: { type: 'none' },
    points: [
      { x: 1.389, y: 0, moveTo: true },
      { x: 7.469, y: 0 },
      { x: 7.469, y: 7.5 },
      { x: 1.322, y: 7.5 },
      { x: 1.166, y: 7.301 },
      { x: 0, y: 3.79, curve: { type: 'cubic', x1: 0.434, y1: 6.322, x2: 0, y2: 5.106 } },
      { x: 1.34, y: 0.057, curve: { type: 'cubic', x1: 0, y1: 2.371, x2: 0.503, y2: 1.071 } },
      { close: true }
    ]
  });
  brandMark(s, 0.823, 0.46);
  // dome-topped "image place holder" (round2SameRect with 50% top corners)
  s.addShape('ellipse', { x: 0.854, y: 1.13, w: 4.632, h: 4.632, fill: { color: C.greyArch }, line: { type: 'none' } });
  s.addShape('rect', { x: 0.854, y: 3.446, w: 4.632, h: 2.463, fill: { color: C.greyArch }, line: { type: 'none' } });
  text(s, 'image place holder', { x: 0.854, y: 3.65, w: 4.632, h: 0.36, fontSize: 18, align: 'center' });
  heading(s, {
    kicker: 'Take the First Step', kx: 8.39, ky: 1.406, light: true,
    x: 8.361, y: 1.84, w: 3.841, h: 2.106,
    title: [[['Start Your Journey Today']]]
  });
  icon(s, 'backpack', 8.469, 4.459, 0.413, C.white);
  text(s, 'Explore your options, create a shortlist for a better research, if possible talk to advisor and visits university to make good comparison.', { x: 8.361, y: 5.148, w: 3.841, h: 0.978, fontSize: 12, color: C.white, lineSpacingMultiple: 1.5 });
  badge(s, 5.942, 0.798);
  rule(s, 12.194, 6.727);
  statCard(s, {
    x: 1.859, y: 3.93, w: 2.993, h: 2.809,
    value: '+46,8%', vx: 2.258, vy: 4.375, vw: 2.231,
    caption: 'Text Here', cx: 2.267, cy: 4.972, cw: 2.508,
    body: 'PLACEHOLDER', bx: 2.267, by: 5.316, bw: 2.231
  });
  bubble(s, 1.222, 1.442, 0.601);
}

// 18 - service university program (phone mock-up)
function slide18(s) {
  cornerPanel(s, 3.462, 1.567, 3.393, 3.31, 0.16804, '');
  // phone mock-up: dark bezel outline, the panel behind shows through the screen
  s.addShape('roundRect', {
    x: 2.56, y: 2.62, w: 3.11, h: 4.97, rectRadius: 0.44,
    fill: { type: 'none' }, line: { color: '1A1A1A', width: 13 }
  });
  panel(s, 3.3, 2.71, 1.63, 0.26, '1A1A1A', 0.35); // notch
  panel(s, 4.853, 0.788, 1.656, 1.611, C.blue, 0.10912, { shadow: SHADOW.card });
  icon(s, 'searchBook', 5.368, 1.131, 0.627, C.white);
  text(s, 'Text Here', { x: 5.026, y: 1.73, w: 1.312, h: 0.326, fontSize: 12, italic: true, color: C.white, align: 'center', lineSpacingMultiple: 1.2 });
  brandMark(s, 0.823, 0.46);
  heading(s, {
    kicker: 'Discover Your Path to Success', kx: 7.747, ky: 1.094,
    x: 7.718, y: 1.465, w: 4.124, h: 2.106,
    title: [[['Service University '], ['Program', C.orange]]]
  });
  badge(s, 1.426, 1.729);
  bubble(s, 12.186, 0.947, 0.601);
  numberedItem(s, { num: '01.', x: 7.726, y: 4.046, bx: 8.464, bw: 3.714, body: lines(['PLACEHOLDER', 'voluptatum deleniti atque corrupti quos'], { breakLine: true }) });
  numberedItem(s, { num: '02.', x: 7.726, y: 5.429, nw: 0.871, bx: 8.464, bw: 3.731, body: lines(['PLACEHOLDER', 'voluptatum deleniti atque corrupti quos'], { breakLine: true }) });
  rule(s, 12.194, 6.727);
  statCard(s, {
    x: 0.838, y: 3.778, w: 2.993, h: 2.45,
    value: '+46,8%', vx: 1.237, vy: 4.196, vw: 2.231,
    caption: 'Text Here', cx: 1.246, cy: 4.793, cw: 2.508,
    body: 'At vero eos et accusamus et iusto odio dignissimos', bx: 1.246, by: 5.137, bw: 2.231, bh: 0.675
  });
}

// 19 - contact
function slide19(s) {
  panel(s, 1.52, 4.807, 9.321, 2.17, C.blue, 0.5);
  heading(s, {
    kicker: "Let's Create a Better Future Together", kx: 4.761, ky: 0.924, kalign: 'left',
    x: 1.333, y: 1.358, w: 10.667, h: 0.774, align: 'center',
    title: [[['Contact us For More '], ['Information', C.orange]]]
  });
  brandMark(s, 0.823, 0.46);
  [
    { icon: 'home', ix: 2.476, iy: 5.22, iw: 0.249, tx: 2.355, label: 'Address', bx: 2.38, bw: 2.535, body: ['1234, Name Street, Building, Your City, Your Country.'] },
    { icon: 'phone', ix: 5.315, iy: 5.167, iw: 0.363, tx: 5.248, label: 'Phone', bx: 5.273, bw: 1.72, body: ['+00 123 4567 890', '+12 345 6789'] },
    { icon: 'mail', ix: 7.393, iy: 5.254, iw: 0.249, tx: 7.272, label: 'Mail', bx: 7.297, bw: 1.941, body: ['yourinfo@email.com', 'yourmail@mail.co '] }
  ].forEach(function (c) {
    icon(s, c.icon, c.ix, c.iy, c.iw, C.white);
    text(s, c.label, { x: c.tx, y: 5.647, w: 1.583, h: 0.37, fontSize: 16, bold: true, fontFace: F.head, color: C.white });
    text(s, lines(c.body), { x: c.bx, y: 5.938, w: c.bw, h: 0.678, fontSize: 12, italic: true, color: C.white, lineSpacingMultiple: 1.5 });
  });
  badge(s, 1.091, 5.465);
  bubble(s, 11.805, 0.473, 0.301);
  rule(s, 12.194, 6.727);
  panel(s, 9.313, 4.006, 3.334, 2.0, C.orange, 0.14785, { shadow: SHADOW.bloom });
  text(s, '2335/', { x: 9.859, y: 4.414, w: 1.606, h: 0.457, fontSize: 28, fontFace: F.head, color: C.white, lineSpacingMultiple: 0.7 });
  text(s, 'Text Here', { x: 11.044, y: 4.409, w: 1.012, h: 0.367, fontSize: 14, italic: true, color: C.white, lineSpacingMultiple: 1.2 });
  text(s, 'PLACEHOLDER', { x: 9.859, y: 4.851, w: 2.454, h: 0.834, fontSize: 10, color: C.white, lineSpacingMultiple: 1.5 });
}

// 20 - thank you
function slide20(s) {
  bands(s, 0, 0, 13.333, 7.5, [[0, C.white], [0.5, C.orange], [1, C.orange]]);
  shadeOval(s, 1.351, 0.756, 5.956, WARM_RAMP);
  brandMark(s, 0.823, 0.46);
  heading(s, {
    kicker: "Let's Create a Better Future Together", kx: 7.761, ky: 1.219, light: true,
    x: 7.731, y: 1.603, w: 3.123, h: 1.927, size: 60,
    title: [[['Thank You']]]
  });
  text(s, 'At vero eos et accusamus et iusto odio dignissimos ducimus qui blanditiis praesentium voluptatum deleniti atque corrupti quos dolores et quas molestias excepturi sint occaecati cupiditate non provident, similique sunt in culpa qui officia', { x: 7.772, y: 3.849, w: 4.482, h: 1.583, fontSize: 12, color: C.white, lineSpacingMultiple: 1.5 });
  text(s, 'WWW. YOURWEBSITE. COM', { x: 7.77, y: 6.312, w: 3.29, h: 0.372, fontSize: 12, color: C.white, charSpacing: 3, lineSpacingMultiple: 1.5 });
  rule(s, 0.845, 6.809);
  badge(s, 1.458, 1.812);
  s.addShape('donut', { x: 6.028, y: 5.286, w: 1.287, h: 1.287, fill: { color: C.blue }, line: { type: 'none' } });
}

/* ------------------------------------------------------------------- build */

const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07,
  slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15,
  slide16, slide17, slide18, slide19, slide20];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE_13x7_5', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE_13x7_5';
  pptx.theme = { headFontFace: F.head, bodyFontFace: F.body };
  pptx.title = 'University Program';

  BUILDERS.forEach(function (fn) {
    const slide = pptx.addSlide();
    slide.background = { color: C.white };
    fn(slide, pptx);
  });

  return pptx.writeFile({
    fileName: path.join(__dirname, '01ff275e-1a44-496b-9d34-0b563460a7fa_grok_final.pptx')
  });
}

build().then(function (f) { console.log('wrote', f); }, function (e) { console.error(e); process.exit(1); });
