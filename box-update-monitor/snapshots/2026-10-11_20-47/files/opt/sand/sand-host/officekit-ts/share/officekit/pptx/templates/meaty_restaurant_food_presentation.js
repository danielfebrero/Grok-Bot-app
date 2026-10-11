/* ---------------------------------------------------------------------------
 * Meaty - Food Presentation Template
 * Recreated with pptxgenjs. Run: node <this file>  ->  writes the .pptx here.
 * ------------------------------------------------------------------------ */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ----- palette ---------------------------------------------------------- */
const RED = 'BC0517';
const DARK = '282B34';
const W = 'FFFFFF';
const MIST = 'F7F7F8';
const GREY = 'A6A6A6';
const GOLD = 'FFC000';
const PHOTO = 'D6D6D6';   // stand-in for the light food photos
const PHOTO2 = '999999';  // stand-in for the darker food photos

/* ----- type ------------------------------------------------------------- */
const SORA = 'Sora SemiBold';
const DM = 'DM Sans';
const DMM = 'DM Sans Medium';
const MONT = 'Montserrat SemiBold';
const POP = 'Poppins SemiBold';

/* ----- reusable text styles --------------------------------------------- */
const HEAD = { sz: 60, f: SORA, v: 'm', ls: 1, m: 0 };            // section headline
const HEADC = { sz: 60, f: SORA, al: 'c', v: 'm', ls: 1, m: 0 };  // centred headline
const EYEBROW = { f: DMM, c: RED, sp: 2 };                        // small red kicker
const EYEBROWC = { f: DMM, c: RED, al: 'c', sp: 2 };
const LABEL = { f: DMM, c: RED };
const LABELR = { f: DMM, c: RED, al: 'r' };
const BODY = { sz: 16, c: GREY, al: 'j', ls: 1.5 };               // grey body copy
const BODYL = { sz: 16, c: GREY, ls: 1.5 };
const BODYR = { sz: 16, c: GREY, al: 'r', ls: 1.5 };
const BODYC = { sz: 16, c: GREY, al: 'c', ls: 1.5 };
const WBODY = { sz: 16, c: W, al: 'j', ls: 1.5 };                 // body copy on cards
const WBODYC = { sz: 16, c: W, al: 'c', ls: 1.5 };
const WTITLE = { f: DMM, c: W, sp: 2 };                           // card heading
const WNAME = { f: DMM, c: W, al: 'c', sp: 1 };                   // person name
const WROLE = { sz: 16, c: W, al: 'c', sp: 2 };                   // person role
const BTN = { f: SORA, c: W, al: 'c', v: 'm' };                   // "Read More"
const BULLET = { sz: 16, c: W, al: 'c', ls: 1.5, sp: 1, bu: true };

function st(base, extra) { return Object.assign({}, base, extra); }

/* ----- primitives ------------------------------------------------------- */
const ALIGN = { c: 'center', r: 'right', j: 'justify', l: 'left' };

/** Text. `spec` toggles between the base colour and the accent at every "|",
 *  and starts a new paragraph at every "\n". */
function tx(s, x, y, w, h, spec, o) {
  o = o || {};
  const base = o.c || DARK;
  const accent = o.a || RED;
  const lines = spec.split('\n');
  const runs = [];
  lines.forEach(function (line, li) {
    line.split('|').forEach(function (piece, i) {
      if (!piece) return;
      const on = i % 2 === 1;
      runs.push({ text: piece, options: { color: on ? accent : base, fontSize: on ? o.az : undefined } });
    });
    if (li < lines.length - 1 && runs.length) {
      runs[runs.length - 1].options.breakLine = true;
    }
  });
  s.addText(runs, {
    x: x, y: y, w: w, h: h,
    fontFace: o.f || DM,
    fontSize: o.sz || 18,
    align: ALIGN[o.al] || 'left',
    valign: o.v === 'm' ? 'middle' : 'top',
    lineSpacingMultiple: o.ls,
    charSpacing: o.sp,
    bullet: o.bu ? { code: '2022' } : false,
    margin: o.m === 0 ? 0 : undefined,
    rotate: o.rot,
    wrap: true
  });
}

/** Plain filled rectangle. */
function box(s, x, y, w, h, color) {
  s.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: color } });
}

/** Filled rounded rectangle. */
function rbox(s, x, y, w, h, r, color) {
  s.addShape('roundRect', { x: x, y: y, w: w, h: h, rectRadius: r, fill: { color: color } });
}

/** Convert a compact path table into pptxgenjs `points`.
 *  [x, y] draws a line; [x, y, cx1, cy1, cx2, cy2] draws a cubic bezier.
 *  The first entry is the moveTo and the path is always closed. */
function pts(table) {
  const out = table.map(function (p, i) {
    if (i === 0) return { x: p[0], y: p[1], moveTo: true };
    if (p.length === 6) return { x: p[0], y: p[1], curve: { type: 'cubic', x1: p[2], y1: p[3], x2: p[4], y2: p[5] } };
    if (p.length === 4) return { x: p[0], y: p[1], curve: { type: 'quadratic', x1: p[2], y1: p[3] } };
    return { x: p[0], y: p[1] };
  });
  out.push({ close: true });
  return out;
}

/** Filled free-form polygon built from a compact path table. */
function poly(s, x, y, w, h, color, table) {
  s.addShape('custGeom', { x: x, y: y, w: w, h: h, fill: { color: color }, points: pts(table) });
}

/** Photo placeholder: a flat colour block where the template had a picture. */
function ph(s, x, y, w, h, color, radius) {
  if (radius) {
    s.addShape('roundRect', { x: x, y: y, w: w, h: h, rectRadius: radius, fill: { color: color } });
  } else {
    s.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: color } });
  }
}

function oval(s, x, y, d, color, h) {
  s.addShape('ellipse', { x: x, y: y, w: d, h: h || d, fill: { color: color } });
}

/* ----- recurring decorations -------------------------------------------- */

/** Cloche (serving dome) logo mark: knob, dome and two plate bars. */
function cloche(s, x, y, d, color) {
  const f = { color: color };
  s.addShape('ellipse', { x: x + 0.435 * d, y: y + 0.12 * d, w: 0.13 * d, h: 0.13 * d, fill: f });
  s.addShape('pie', { x: x + 0.06 * d, y: y + 0.19 * d, w: 0.88 * d, h: 0.88 * d, angleRange: [180, 360], fill: f });
  s.addShape('roundRect', { x: x, y: y + 0.6 * d, w: d, h: 0.085 * d, rectRadius: 0.042 * d, fill: f });
  s.addShape('roundRect', { x: x + 0.07 * d, y: y + 0.73 * d, w: 0.86 * d, h: 0.075 * d, rectRadius: 0.037 * d, fill: f });
}

/** The logo / page-number rail that runs down the left or right edge.
 *  `tone` is DARK on light slides and W on the dark/red ones. */
function rail(s, pageNo, side, tone) {
  const left = side === 'L';
  const rule = tone === W ? W : RED;
  const lx = left ? 0.634 : 19.375;
  cloche(s, left ? 0.236 : 18.977, 0.589, 0.795, tone);
  s.addShape('line', { x: lx, y: 1.667, w: 0, h: 1.844, line: { color: rule, width: 3 } });
  s.addShape('line', { x: lx, y: 7.906, w: 0, h: 1.969, line: { color: rule, width: 3 } });
  tx(s, left ? -1.273 : 17.469, 5.5, 3.813, 0.438, 'Meaty Reastaurant',
    { sz: 20, f: DMM, c: tone, al: 'c', sp: 2, rot: left ? 270 : 90 });
  tx(s, left ? 0.236 : 19.047, 10.156, 0.795, 0.505, pageNo, { sz: 24, f: DMM, c: tone, al: 'c' });
}

/** Full-height panel whose inner edge is an S-curve (cover / closing slides). */
function wave(s, x, y, w, h, color, mirror) {
  const b = 0.123 * w;          // how far the curve bulges
  const e = mirror ? w - b : b; // straight edge sits opposite the curve
  const pts = mirror
    ? [{ x: 0, y: 0, moveTo: true }, { x: e, y: 0 },
       { x: e + 0.005 * w, y: h / 2, curve: { type: 'cubic', x1: e - 0.002 * w, y1: 0.16 * h, x2: e - 0.21 * w, y2: 0.28 * h } },
       { x: e, y: h, curve: { type: 'cubic', x1: e + 0.21 * w, y1: 0.67 * h, x2: e - 0.002 * w, y2: 0.83 * h } },
       { x: 0, y: h }, { close: true }]
    : [{ x: e, y: 0, moveTo: true }, { x: w, y: 0 }, { x: w, y: h }, { x: e, y: h },
       { x: e - 0.003 * w, y: h / 2, curve: { type: 'cubic', x1: e - 0.001 * w, y1: 0.82 * h, x2: e - 0.27 * w, y2: 0.73 * h } },
       { x: e, y: 0, curve: { type: 'cubic', x1: e + 0.27 * w, y1: 0.21 * h, x2: e - 0.001 * w, y2: 0.16 * h } },
       { close: true }];
  s.addShape('custGeom', { x: x, y: y, w: w, h: h, fill: { color: color }, points: pts });
}

/** Circled tick used on the "reasons" list. */
function tick(s, x, y, d) {
  s.addShape('ellipse', { x: x, y: y, w: d, h: d, line: { color: W, width: 3 } });
  s.addShape('custGeom', {
    x: x + 0.22 * d, y: y + 0.3 * d, w: 0.56 * d, h: 0.4 * d, fill: { color: W },
    points: [{ x: 0, y: 0.2 * d, moveTo: true }, { x: 0.09 * d, y: 0.11 * d },
             { x: 0.21 * d, y: 0.23 * d }, { x: 0.47 * d, y: 0 },
             { x: 0.56 * d, y: 0.09 * d }, { x: 0.21 * d, y: 0.4 * d }, { close: true }]
  });
}

/** Row of three social glyphs (twitter / facebook / instagram). */
function social(s, x, y) {
  const d = 0.26;
  s.addShape('custGeom', {
    x: x, y: y, w: 0.32, h: d, fill: { color: W },
    points: [{ x: 0, y: 0.05, moveTo: true }, { x: 0.32, y: 0 }, { x: 0.19, y: 0.12 },
             { x: 0.29, y: d }, { x: 0.12, y: 0.17 }, { x: 0, y: d }, { close: true }]
  });
  s.addShape('custGeom', {
    x: x + 0.67, y: y, w: 0.16, h: d, fill: { color: W },
    points: [{ x: 0.16, y: 0, moveTo: true }, { x: 0.16, y: 0.06 }, { x: 0.09, y: 0.06 },
             { x: 0.09, y: 0.1 }, { x: 0.16, y: 0.1 }, { x: 0.14, y: 0.16 },
             { x: 0.09, y: 0.16 }, { x: 0.09, y: d }, { x: 0.03, y: d },
             { x: 0.03, y: 0.16 }, { x: 0, y: 0.16 }, { x: 0, y: 0.1 },
             { x: 0.03, y: 0.1 }, { x: 0.03, y: 0.04 }, { close: true }]
  });
  s.addShape('roundRect', { x: x + 1.27, y: y, w: d, h: d, rectRadius: 0.07, line: { color: W, width: 1.6 } });
  s.addShape('ellipse', { x: x + 1.34, y: y + 0.07, w: 0.12, h: 0.12, line: { color: W, width: 1.6 } });
}

/** Rounded tile carrying a social glyph (Instagram / Facebook / LinkedIn). */
function socialTile(s, kind, x, y, d, color) {
  s.addShape('roundRect', { x: x, y: y, w: d, h: d, rectRadius: 0.22 * d, fill: { color: color } });
  const m = 0.26 * d;
  const i = d - 2 * m;
  if (kind === 'ig') {
    s.addShape('roundRect', { x: x + m, y: y + m, w: i, h: i, rectRadius: 0.3 * i, line: { color: W, width: 2.2 } });
    s.addShape('ellipse', { x: x + m + 0.3 * i, y: y + m + 0.3 * i, w: 0.4 * i, h: 0.4 * i, line: { color: W, width: 2.2 } });
  } else if (kind === 'fb') {
    s.addShape('custGeom', {
      x: x + 0.36 * d, y: y + m, w: 0.28 * d, h: i, fill: { color: W },
      points: [{ x: 0.28 * d, y: 0, moveTo: true }, { x: 0.28 * d, y: 0.22 * i },
               { x: 0.14 * d, y: 0.22 * i }, { x: 0.14 * d, y: 0.36 * i },
               { x: 0.27 * d, y: 0.36 * i }, { x: 0.23 * d, y: 0.56 * i },
               { x: 0.14 * d, y: 0.56 * i }, { x: 0.14 * d, y: i },
               { x: 0.03 * d, y: i }, { x: 0.03 * d, y: 0.56 * i },
               { x: 0, y: 0.56 * i }, { x: 0, y: 0.36 * i },
               { x: 0.03 * d, y: 0.36 * i }, { x: 0.03 * d, y: 0.14 * i }, { close: true }]
    });
  } else {
    s.addShape('rect', { x: x + 0.3 * d, y: y + 0.3 * d, w: 0.09 * d, h: 0.09 * d, fill: { color: W } });
    s.addShape('rect', { x: x + 0.3 * d, y: y + 0.44 * d, w: 0.09 * d, h: 0.26 * d, fill: { color: W } });
    s.addShape('custGeom', {
      x: x + 0.45 * d, y: y + 0.44 * d, w: 0.25 * d, h: 0.26 * d, fill: { color: W },
      points: [{ x: 0, y: 0, moveTo: true }, { x: 0.09 * d, y: 0 }, { x: 0.09 * d, y: 0.04 * d },
               { x: 0.25 * d, y: 0.1 * d }, { x: 0.25 * d, y: 0.26 * d },
               { x: 0.16 * d, y: 0.26 * d }, { x: 0.16 * d, y: 0.13 * d },
               { x: 0.09 * d, y: 0.13 * d }, { x: 0.09 * d, y: 0.26 * d },
               { x: 0, y: 0.26 * d }, { close: true }]
    });
  }
}

/** Ink-splat mark at the centre of the infographic wheels. */
function splat(s, x, y, d) {
  const c = { color: '8C0410' };
  s.addShape('ellipse', { x: x + 0.22 * d, y: y + 0.22 * d, w: 0.5 * d, h: 0.48 * d, fill: c });
  s.addShape('ellipse', { x: x + 0.58 * d, y: y + 0.12 * d, w: 0.2 * d, h: 0.2 * d, fill: c });
  s.addShape('ellipse', { x: x + 0.06 * d, y: y + 0.5 * d, w: 0.24 * d, h: 0.22 * d, fill: c });
  s.addShape('ellipse', { x: x + 0.62 * d, y: y + 0.56 * d, w: 0.26 * d, h: 0.24 * d, fill: c });
  s.addShape('ellipse', { x: x + 0.38 * d, y: y + 0.72 * d, w: 0.16 * d, h: 0.16 * d, fill: c });
  s.addShape('ellipse', { x: x + 0.84 * d, y: y + 0.36 * d, w: 0.09 * d, h: 0.09 * d, fill: c });
  s.addShape('ellipse', { x: x + 0.02 * d, y: y + 0.2 * d, w: 0.07 * d, h: 0.07 * d, fill: c });
}

/** Simplified line-art pictograms used across the infographic slides. */
function glyph(s, kind, x, y, d, color) {
  const f = { color: color };
  const ln = { color: color, width: Math.max(1, 2.2 * d) };
  if (kind === 'chart') {
    s.addShape('rect', { x: x + 0.16 * d, y: y + 0.2 * d, w: 0.68 * d, h: 0.6 * d, line: ln });
    [0.3, 0.45, 0.6].forEach(function (bx, i) {
      s.addShape('rect', { x: x + bx * d, y: y + (0.68 - i * 0.13) * d, w: 0.08 * d, h: (0.06 + i * 0.13) * d, fill: f });
    });
  } else if (kind === 'medal') {
    s.addShape('ellipse', { x: x + 0.24 * d, y: y + 0.12 * d, w: 0.52 * d, h: 0.52 * d, line: ln });
    s.addShape('ellipse', { x: x + 0.4 * d, y: y + 0.28 * d, w: 0.2 * d, h: 0.2 * d, fill: f });
    s.addShape('rect', { x: x + 0.36 * d, y: y + 0.62 * d, w: 0.08 * d, h: 0.26 * d, fill: f, rotate: 12 });
    s.addShape('rect', { x: x + 0.56 * d, y: y + 0.62 * d, w: 0.08 * d, h: 0.26 * d, fill: f, rotate: -12 });
  } else if (kind === 'print' || kind === 'target') {
    s.addShape('ellipse', { x: x + 0.16 * d, y: y + 0.16 * d, w: 0.68 * d, h: 0.68 * d, line: ln });
    s.addShape('ellipse', { x: x + 0.32 * d, y: y + 0.32 * d, w: 0.36 * d, h: 0.36 * d, line: ln });
    s.addShape('ellipse', { x: x + 0.44 * d, y: y + 0.44 * d, w: 0.12 * d, h: 0.12 * d, fill: f });
  } else if (kind === 'rocket') {
    s.addShape('custGeom', {
      x: x + 0.32 * d, y: y + 0.12 * d, w: 0.36 * d, h: 0.52 * d, fill: f,
      points: [{ x: 0.18 * d, y: 0, moveTo: true }, { x: 0.36 * d, y: 0.3 * d },
               { x: 0.36 * d, y: 0.52 * d }, { x: 0, y: 0.52 * d },
               { x: 0, y: 0.3 * d }, { close: true }]
    });
    s.addShape('custGeom', {
      x: x + 0.24 * d, y: y + 0.5 * d, w: 0.52 * d, h: 0.24 * d, fill: f,
      points: [{ x: 0, y: 0.24 * d, moveTo: true }, { x: 0.1 * d, y: 0 },
               { x: 0.42 * d, y: 0 }, { x: 0.52 * d, y: 0.24 * d }, { close: true }]
    });
    s.addShape('ellipse', { x: x + 0.44 * d, y: y + 0.74 * d, w: 0.12 * d, h: 0.14 * d, fill: f });
  } else if (kind === 'bulb') {
    s.addShape('ellipse', { x: x + 0.26 * d, y: y + 0.12 * d, w: 0.48 * d, h: 0.48 * d, fill: f });
    s.addShape('rect', { x: x + 0.4 * d, y: y + 0.56 * d, w: 0.2 * d, h: 0.2 * d, fill: f });
    s.addShape('rect', { x: x + 0.38 * d, y: y + 0.8 * d, w: 0.24 * d, h: 0.07 * d, fill: f });
  } else if (kind === 'gear') {
    s.addShape('star8', { x: x + 0.1 * d, y: y + 0.1 * d, w: 0.62 * d, h: 0.62 * d, fill: f });
    s.addShape('ellipse', { x: x + 0.29 * d, y: y + 0.29 * d, w: 0.24 * d, h: 0.24 * d, fill: { color: RED } });
    s.addShape('star8', { x: x + 0.44 * d, y: y + 0.44 * d, w: 0.46 * d, h: 0.46 * d, fill: f });
  } else if (kind === 'eye') {
    s.addShape('ellipse', { x: x + 0.1 * d, y: y + 0.3 * d, w: 0.8 * d, h: 0.4 * d, line: ln });
    s.addShape('ellipse', { x: x + 0.39 * d, y: y + 0.39 * d, w: 0.22 * d, h: 0.22 * d, fill: f });
  } else if (kind === 'stack') {
    [0.2, 0.44, 0.68].forEach(function (sy) {
      s.addShape('ellipse', { x: x + 0.16 * d, y: y + sy * d, w: 0.68 * d, h: 0.2 * d, line: ln });
    });
  } else if (kind === 'board') {
    s.addShape('rect', { x: x + 0.14 * d, y: y + 0.16 * d, w: 0.72 * d, h: 0.52 * d, line: ln });
    s.addShape('rect', { x: x + 0.47 * d, y: y + 0.68 * d, w: 0.06 * d, h: 0.18 * d, fill: f });
    [0.3, 0.45, 0.6].forEach(function (bx, i) {
      s.addShape('rect', { x: x + bx * d, y: y + (0.58 - i * 0.11) * d, w: 0.07 * d, h: (0.04 + i * 0.11) * d, fill: f });
    });
  } else if (kind === 'phone') {
    s.addShape('custGeom', {
      x: x + 0.15 * d, y: y + 0.15 * d, w: 0.7 * d, h: 0.7 * d, fill: f,
      points: [{ x: 0.08 * d, y: 0, moveTo: true }, { x: 0.26 * d, y: 0 },
               { x: 0.32 * d, y: 0.2 * d }, { x: 0.2 * d, y: 0.3 * d },
               { x: 0.4 * d, y: 0.5 * d }, { x: 0.5 * d, y: 0.38 * d },
               { x: 0.7 * d, y: 0.44 * d }, { x: 0.7 * d, y: 0.62 * d },
               { x: 0.56 * d, y: 0.7 * d }, { x: 0.24 * d, y: 0.56 * d },
               { x: 0, y: 0.2 * d }, { close: true }]
    });
  } else if (kind === 'pin') {
    s.addShape('rect', { x: x + 0.18 * d, y: y + 0.2 * d, w: 0.36 * d, h: 0.62 * d, fill: f });
    s.addShape('rect', { x: x + 0.56 * d, y: y + 0.4 * d, w: 0.26 * d, h: 0.42 * d, fill: f });
    s.addShape('rect', { x: x + 0.26 * d, y: y + 0.3 * d, w: 0.08 * d, h: 0.08 * d, fill: { color: RED } });
    s.addShape('rect', { x: x + 0.4 * d, y: y + 0.3 * d, w: 0.08 * d, h: 0.08 * d, fill: { color: RED } });
  } else if (kind === 'mail') {
    s.addShape('rect', { x: x + 0.1 * d, y: y + 0.26 * d, w: 0.8 * d, h: 0.5 * d, fill: f });
    s.addShape('custGeom', {
      x: x + 0.1 * d, y: y + 0.26 * d, w: 0.8 * d, h: 0.32 * d, fill: { color: RED },
      points: [{ x: 0.04 * d, y: 0.03 * d, moveTo: true }, { x: 0.4 * d, y: 0.3 * d },
               { x: 0.76 * d, y: 0.03 * d }, { x: 0.76 * d, y: 0 },
               { x: 0.04 * d, y: 0 }, { close: true }]
    });
  }
}

/** Four-lobed clover behind the SWOT circles: two crossing stadium bars. */
function quatrefoil(s, x, y, d, color) {
  const arm = 0.35 * d;   // bar thickness
  const c = d / 2;
  s.addShape('roundRect', { x: x, y: y + c - arm / 2, w: d, h: arm, rectRadius: arm / 2, fill: { color: color } });
  s.addShape('roundRect', { x: x + c - arm / 2, y: y, w: arm, h: d, rectRadius: arm / 2, fill: { color: color } });
  s.addShape('ellipse', { x: x + c - 0.29 * d, y: y + c - 0.29 * d, w: 0.58 * d, h: 0.58 * d, fill: { color: color } });
}

/** Rectangle with two diagonally-opposite corners rounded (SWOT petals). */
function petal(s, x, y, w, h, r, corners, color) {
  const tlbr = corners === 'tl-br';
  const pts = tlbr
    ? [{ x: r, y: 0, moveTo: true },
       { x: 0, y: r, curve: { type: 'cubic', x1: 0.45 * r, y1: 0, x2: 0, y2: 0.45 * r } },
       { x: 0, y: h }, { x: w - r, y: h },
       { x: w, y: h - r, curve: { type: 'cubic', x1: w - 0.45 * r, y1: h, x2: w, y2: h - 0.45 * r } },
       { x: w, y: 0 }, { close: true }]
    : [{ x: 0, y: 0, moveTo: true }, { x: 0, y: h - r },
       { x: r, y: h, curve: { type: 'cubic', x1: 0, y1: h - 0.45 * r, x2: 0.45 * r, y2: h } },
       { x: w, y: h }, { x: w, y: r },
       { x: w - r, y: 0, curve: { type: 'cubic', x1: w, y1: 0.45 * r, x2: w - 0.45 * r, y2: 0 } },
       { close: true }];
  s.addShape('custGeom', { x: x, y: y, w: w, h: h, fill: { color: color }, points: pts });
}

/** Chevron arrow motif on the SWOT arrow slide. */
function swotArrow(s, x, y, w, h, color, mirror) {
  const p = [[0.604, 0], [0.604, 0.162], [0.234, 0.162], [0, 0.484], [0, 1],
             [0.234, 0.679], [0.604, 0.679], [0.604, 0.855], [1, 0.429], [0.604, 0]];
  const pts = p.map(function (q, i) {
    const px = mirror ? (1 - q[0]) * w : q[0] * w;
    return i === 0 ? { x: px, y: q[1] * h, moveTo: true } : { x: px, y: q[1] * h };
  });
  pts.push({ close: true });
  s.addShape('custGeom', { x: x, y: y, w: w, h: h, fill: { color: color }, points: pts });
}

/** Rounded corner "tail" that tucks under the arrow motif. */
function swotTail(s, x, y, w, h, color, mirror) {
  const p = mirror
    ? [[0.62, 0], [1, 0], [1, 1], [0.49, 1], [0, 0.55], [0, 0]]
    : [[0.38, 0], [0, 0], [0, 1], [0.51, 1], [1, 0.55], [1, 0]];
  const pts = p.map(function (q, i) {
    return i === 0 ? { x: q[0] * w, y: q[1] * h, moveTo: true } : { x: q[0] * w, y: q[1] * h };
  });
  pts.push({ close: true });
  s.addShape('custGeom', { x: x, y: y, w: w, h: h, fill: { color: color }, points: pts });
}

/** Donut segment (used for the SWOT wheel). */
function segment(s, cx, cy, r, from, to, thick, color) {
  s.addShape('blockArc', {
    x: cx - r, y: cy - r, w: 2 * r, h: 2 * r,
    angleRange: [from, to], arcThicknessRatio: thick, fill: { color: color }
  });
}

/** Laptop mock-up: dark bezel, grey screen and a light base. */
function laptop(s, x, y, w, h) {
  s.addShape('roundRect', { x: x + 0.06 * w, y: y, w: 0.88 * w, h: 0.83 * h, rectRadius: 0.02 * w, fill: { color: '181818' } });
  s.addShape('rect', { x: x + 0.09 * w, y: y + 0.04 * h, w: 0.82 * w, h: 0.72 * h, fill: { color: PHOTO2 } });
  s.addShape('roundRect', { x: x, y: y + 0.83 * h, w: w, h: 0.09 * h, rectRadius: 0.02 * h, fill: { color: 'D2D3D5' } });
  s.addShape('roundRect', { x: x + 0.4 * w, y: y + 0.83 * h, w: 0.2 * w, h: 0.03 * h, rectRadius: 0.01 * h, fill: { color: 'A8A9AA' } });
}

/** Desktop monitor mock-up: bezel, screen, neck and foot. */
function monitor(s, x, y, w, h) {
  s.addShape('roundRect', { x: x, y: y, w: w, h: 0.78 * h, rectRadius: 0.012 * w, fill: { color: '181818' } });
  s.addShape('rect', { x: x + 0.025 * w, y: y + 0.04 * h, w: 0.95 * w, h: 0.66 * h, fill: { color: PHOTO2 } });
  s.addShape('rect', { x: x + 0.42 * w, y: y + 0.78 * h, w: 0.16 * w, h: 0.12 * h, fill: { color: 'D2D3D5' } });
  s.addShape('roundRect', { x: x + 0.22 * w, y: y + 0.9 * h, w: 0.56 * w, h: 0.07 * h, rectRadius: 0.02 * h, fill: { color: 'D2D3D5' } });
}

/** Smartphone mock-up: rounded white body with a grey screen. */
function phone(s, x, y, w, h) {
  s.addShape('roundRect', { x: x, y: y, w: w, h: h, rectRadius: 0.09 * w, fill: { color: W }, line: { color: 'D2D3D5', width: 1.5 } });
  s.addShape('rect', { x: x + 0.05 * w, y: y + 0.09 * h, w: 0.9 * w, h: 0.8 * h, fill: { color: PHOTO } });
  s.addShape('roundRect', { x: x + 0.36 * w, y: y + 0.045 * h, w: 0.28 * w, h: 0.012 * h, rectRadius: 0.006 * h, fill: { color: 'A8A9AA' } });
  s.addShape('ellipse', { x: x + 0.43 * w, y: y + 0.915 * h, w: 0.14 * w, h: 0.14 * w, line: { color: 'D2D3D5', width: 1.5 } });
}

/* A stand-in "icon library": each cell is a thin-stroke pictogram built from
 * one or two preset shapes, cycled so the sheet reads like the original grid. */
const ICON_CELLS = [
  ['roundRect'], ['ellipse'], ['triangle'], ['diamond'], ['mathPlus'], ['hexagon'],
  ['star5'], ['chevron'], ['heart'], ['pentagon'], ['cloud'], ['moon'],
  ['sun'], ['teardrop'], ['donut'], ['arc'], ['can'], ['cube'],
  ['bevel'], ['flowChartConnector'], ['flowChartDocument'], ['homePlate'], ['plaque'], ['smileyFace']
];

/** Grid of outlined icon placeholders (the two icon-library slides). */
function iconGrid(s, cols, rows) {
  const d = 0.52;
  rows.forEach(function (cy, r) {
    cols.forEach(function (cx, c) {
      const kind = ICON_CELLS[(r * cols.length + c) % ICON_CELLS.length][0];
      s.addShape(kind, {
        x: cx - d / 2, y: cy - d / 2, w: d, h: d,
        line: { color: W, width: 1.25 }
      });
    });
  });
}

/* ----- shared body copy -------------------------------------------------- */
const L1 =
  'PLACEHOLDER';
const L2 =
  'PLACEHOLDER';
const L3 =
  'PLACEHOLDER';
const L4 =
  'PLACEHOLDER';
const L5 =
  'PLACEHOLDER';
const L6 =
  'PLACEHOLDER';
const L7 =
  'PLACEHOLDER';
const L8 =
  'In eu mi bibendum neque egestas congue quisque';
const L9 =
  'PLACEHOLDER';
const L10 =
  'PLACEHOLDER';
const L11 =
  'PLACEHOLDER';
const L12 =
  'PLACEHOLDER';
const L13 =
  'PLACEHOLDER';
const L14 =
  'PLACEHOLDER';
const L15 =
  'PLACEHOLDER';
const L16 =
  'Lorem ipsum and a the dolor sitet and consectetur';

/* ----- slides ----------------------------------------------------------- */
function slide01(p) {
  const s = p.addSlide();
  rail(s, '01', 'L', DARK);
  wave(s, 14.146, 0, 5.854, 11.25, DARK);
  tx(s, 1.267, 3.9, 8.998, 3.046, 'Me|aty|', { sz: 175, f: SORA, al: 'c' });
  tx(s, 1.267, 6.946, 8.998, 0.404, 'Food Presentation| Template|', { f: DMM, al: 'c' });
  oval(s, 11.161, 1.786, 7.677, RED);
  oval(s, 11.466, 2.091, 7.067, PHOTO2);
}

function slide02(p) {
  const s = p.addSlide();
  rail(s, '02', 'R', DARK);
  rbox(s, 0.896, 0.896, 6.556, 9.453, 0.512, DARK);
  tx(s, 10.124, 2.717, 7.918, 0.404, 'About Restaurant', EYEBROW);
  tx(s, 10.124, 3.121, 7.918, 1.96, 'Welcome to |Meaty| Restaurant', HEAD);
  tx(s, 10.124, 5.581, 7.918, 1.68, 'PLACEHOLDER', BODY);
  rbox(s, 10.124, 7.671, 2.327, 0.863, 0.138, RED);
  tx(s, 10.124, 7.671, 2.327, 0.863, 'Read More', BTN);
  ph(s, 1.195, 1.216, 5.958, 8.812, PHOTO, 0.465);
  rbox(s, 3.782, 2.161, 5.359, 7.419, 0.366, RED);
  ph(s, 3.987, 2.38, 4.95, 6.982, PHOTO2, 0.338);
}

function slide03(p) {
  const s = p.addSlide();
  rail(s, '03', 'R', DARK);
  ph(s, 5.909, 1.372, 4.536, 8.503, PHOTO2, 0.371);
  tx(s, 11.159, 2.567, 6.98, 1.96, 'Meaty| Best Food|', { sz: 60, f: SORA, c: RED, a: DARK, v: 'm', ls: 1, m: 0 });
  tx(s, 11.159, 4.986, 6.98, 1.276, 'PLACEHOLDER', BODY);
  rbox(s, 6.659, 6.979, 5.573, 2.231, 0.357, RED);
  rbox(s, 12.566, 6.979, 5.573, 2.231, 0.357, RED);
  tx(s, 6.9, 7.457, 5.03, 1.276, L9, WBODY);
  tx(s, 12.808, 7.457, 5.03, 1.276, L9, WBODY);
  rbox(s, 0, -1.889, 5.583, 7.331, 0.456, RED);
  rbox(s, 0, 5.738, 5.583, 7.179, 0.456, DARK);
  tx(s, 11.159, 2.163, 6.98, 0.404, 'About Restaurant', EYEBROW);
  ph(s, 0.15, -1.704, 5.283, 6.961, PHOTO2, 0.432);
  ph(s, 0.15, 5.91, 5.283, 6.835, PHOTO, 0.432);
}

function slide04(p) {
  const s = p.addSlide();
  rail(s, '04', 'L', DARK);
  poly(s, 12.387, 0.296, 7.613, 10.658, RED, [[0.406, 0], [7.613, 0], [7.613, 10.658], [0.406, 10.658], [0, 10.23, 0.182, 10.658, 0, 10.466], [0, 0.428], [0.406, 0, 0, 0.192, 0.182, 0]]);
  tx(s, 2.041, 2.553, 8.224, 0.404, 'About Restaurant', EYEBROW);
  tx(s, 2.041, 2.96, 8.224, 1.96, 'Favorite| Food in Our Restaurant|', { sz: 60, f: SORA, c: RED, a: DARK, v: 'm', ls: 1, m: 0 });
  tx(s, 3.425, 5.833, 6.84, 0.872, L10, BODY);
  tx(s, 3.427, 5.426, 6.839, 0.404, 'Steak', { f: DMM, sp: 1 });
  tx(s, 3.425, 7.602, 6.84, 0.872, L10, BODY);
  tx(s, 3.427, 7.199, 6.839, 0.404, 'Beef', { f: DMM, sp: 1 });
  rbox(s, 2.041, 5.511, 1.109, 1.109, 0.177, RED);
  tx(s, 2.041, 5.511, 1.109, 1.109, '01', { sz: 24, f: DMM, c: W, al: 'c', v: 'm' });
  rbox(s, 2.041, 7.282, 1.109, 1.109, 0.177, RED);
  tx(s, 2.041, 7.282, 1.109, 1.109, '02', { sz: 24, f: DMM, c: W, al: 'c', v: 'm' });
  ph(s, 11.161, 0.896, 8.224, 9.453, PHOTO, 0.486);
}

function slide05(p) {
  const s = p.addSlide();
  rail(s, '05', 'R', DARK);
  poly(s, 0, 7.439, 18.323, 3.811, PHOTO, [[0, 0], [14.526, 0], [18.323, 3.808, 16.623, 0, 18.323, 1.705], [18.323, 3.811], [0, 3.811]]);
  tx(s, 10.281, 0.896, 7.488, 0.404, 'About Restaurant', EYEBROW);
  tx(s, 10.281, 1.3, 7.488, 1.96, 'Food Menu at Our |Restaurant|', HEAD);
  tx(s, 10.281, 3.761, 7.488, 1.68, 'PLACEHOLDER', BODY);
  rbox(s, 10.281, 5.941, 2.327, 0.863, 0.138, RED);
  tx(s, 10.281, 5.941, 2.327, 0.863, 'Read More', BTN);
  rbox(s, 0.896, 0.896, 8.437, 9.453, 0.664, RED);
  ph(s, 1.165, 1.197, 7.899, 8.851, PHOTO2, 0.622);
}

function slide06(p) {
  const s = p.addSlide();
  rail(s, '06', 'R', DARK);
  poly(s, 0, 0, 6.396, 6.522, RED, [[0, 0], [5.952, 0], [6.036, 0.174], [6.396, 1.953, 6.268, 0.721, 6.396, 1.322], [1.814, 6.522, 6.396, 4.476, 4.344, 6.522], [0.03, 6.163, 1.181, 6.522, 0.578, 6.394], [0, 6.149]]);
  ph(s, 0.896, 0.896, 7.947, 11.943, PHOTO, 0.55);
  tx(s, 9.738, 2.551, 8.201, 1.96, 'Fast and Safe Food |Delivery|', HEAD);
  tx(s, 9.738, 5.008, 8.201, 1.276, 'PLACEHOLDER', BODY);
  tx(s, 9.738, 2.147, 8.201, 0.404, 'About Restaurant', EYEBROW);
  rbox(s, 7.789, 6.872, 4.827, 2.231, 0.357, RED);
  rbox(s, 13.112, 6.872, 4.827, 2.231, 0.357, RED);
  rbox(s, 2.465, 6.872, 4.827, 2.231, 0.357, RED);
  tx(s, 2.719, 7.553, 4.318, 1.276, L5, WBODY);
  tx(s, 2.721, 7.146, 4.317, 0.404, 'Fast Delivery', { f: DMM, c: W, sp: 1 });
  tx(s, 8.043, 7.553, 4.318, 1.276, L5, WBODY);
  tx(s, 8.044, 7.146, 4.317, 0.404, 'Safe Delivery', { f: DMM, c: W, sp: 1 });
  tx(s, 13.367, 7.553, 4.318, 1.276, L5, WBODY);
  tx(s, 13.367, 7.146, 4.317, 0.404, 'Happy Delivery', { f: DMM, c: W, sp: 1 });
}

function slide07(p) {
  const s = p.addSlide();
  rail(s, '07', 'L', DARK);
  oval(s, 13.935, -0.944, 8.252, DARK);
  oval(s, 13.935, 4.965, 8.252, RED);
  tx(s, 2.185, 6.437, 8.088, 0.872, L11, BODY);
  tx(s, 2.185, 6.033, 8.088, 0.404, 'First Reason', { f: DMM, sp: 2 });
  tx(s, 2.185, 8.219, 8.088, 0.872, L11, BODY);
  tx(s, 2.185, 7.815, 8.088, 0.404, 'Second Reason', { f: DMM, sp: 2 });
  tx(s, 2.178, 2.159, 8.088, 0.404, 'About Restaurant', EYEBROW);
  tx(s, 2.178, 2.565, 8.088, 2.97, 'Reasons why |Meaty| is the Best Restaurant', HEAD);
  oval(s, 10.989, 1.499, 8.252, PHOTO);
}

function slide08(p) {
  const s = p.addSlide();
  rail(s, '08', 'R', DARK);
  poly(s, 0, 9.434, 3.581, 1.816, RED, [[0, 0], [1.496, 0.72, 0.499, 0.24, 0.242, 0.79], [3.581, 1.816, 2.751, 0.65, 2.886, 1.451], [0, 1.816], [0, 0]]);
  ph(s, 5.164, 4.084, 3.81, 6.265, PHOTO2, 0.243);
  ph(s, 0.88, 0.896, 4.047, 7.292, PHOTO2, 0.258);
  tx(s, 9.853, 2.088, 8.086, 2.97, 'Meaty| Has Been Around For Over 20 Years|', { sz: 60, f: SORA, c: RED, a: DARK, v: 'm', ls: 1, m: 0 });
  tx(s, 9.853, 5.564, 8.086, 1.276, 'PLACEHOLDER', BODY);
  tx(s, 9.853, 1.681, 8.086, 0.404, 'About Restaurant', EYEBROW);
  rbox(s, 7.72, 7.338, 4.827, 2.231, 0.357, RED);
  rbox(s, 13.112, 7.338, 4.827, 2.231, 0.357, RED);
  tx(s, 7.975, 8.019, 4.318, 1.276, L5, WBODY);
  tx(s, 7.975, 7.612, 4.317, 0.404, 'Business License', { f: DMM, c: W, sp: 1 });
  tx(s, 13.367, 8.019, 4.318, 1.276, L5, WBODY);
  tx(s, 13.367, 7.612, 4.317, 0.404, 'Food Safety Letter', { f: DMM, c: W, sp: 1 });
}

function slide09(p) {
  const s = p.addSlide();
  rail(s, '09', 'L', DARK);
  box(s, 14.774, 0, 5.226, 11.247, DARK);
  rbox(s, 2.144, 5.585, 3.874, 0.929, 0.149, RED);
  rbox(s, 2.144, 6.838, 3.874, 0.929, 0.149, RED);
  rbox(s, 6.391, 5.585, 3.874, 0.929, 0.149, RED);
  rbox(s, 6.391, 6.838, 3.874, 0.929, 0.149, RED);
  tx(s, 2.144, 2.111, 8.122, 2.97, 'The Reason We \nare the Best |Restaurant|', HEAD);
  tx(s, 2.144, 1.707, 8.122, 0.404, 'About Restaurant', EYEBROW);
  tx(s, 2.144, 8.268, 8.122, 1.275, 'PLACEHOLDER', BODY);
  tx(s, 7.417, 7.101, 2.68, 0.404, 'Best Dessert', WTITLE);
  tx(s, 7.417, 5.847, 2.68, 0.404, 'Hygienic Food', WTITLE);
  tx(s, 3.15, 7.101, 2.68, 0.404, 'Delicious food', WTITLE);
  tx(s, 3.15, 5.847, 2.68, 0.404, 'Best Food', WTITLE);
  ph(s, 12.335, 5.747, 6.769, 4.764, PHOTO2, 0.439);
  ph(s, 11.161, 0.74, 6.33, 6.394, PHOTO, 0.442);
  [5.736, 6.99].forEach(function (y) { [2.331, 6.56].forEach(function (x) { tick(s, x, y, 0.632); }); });
}

function slide10(p) {
  const s = p.addSlide();
  rail(s, '10', 'R', DARK);
  poly(s, 0, 4.418, 17.914, 5.684, DARK, [[0, 0], [15.072, 0], [17.914, 2.842, 16.642, 0, 17.914, 1.272], [15.072, 5.684, 17.914, 4.412, 16.642, 5.684], [0, 5.684]]);
  oval(s, 13.157, 5.055, 3.859, PHOTO);
  oval(s, 9.006, 5.055, 3.859, PHOTO2);
  oval(s, 4.881, 5.055, 3.859, PHOTO);
  oval(s, 0.744, 5.055, 3.859, PHOTO2);
  tx(s, 3.926, 1.296, 10.958, 0.95, 'Introducing| Our Team|', { sz: 60, f: SORA, c: RED, a: DARK, al: 'c', v: 'm', ls: 1, m: 0 });
  tx(s, 3.926, 0.896, 10.958, 0.404, 'Team', EYEBROWC);
  tx(s, 3.926, 2.538, 10.958, 0.871, L12, BODYC);
  rbox(s, 0.736, 7.944, 3.875, 1.521, 0.637, RED);
  rbox(s, 4.873, 7.944, 3.875, 1.521, 0.637, RED);
  rbox(s, 8.998, 7.944, 3.875, 1.521, 0.637, RED);
  rbox(s, 13.149, 7.944, 3.875, 1.521, 0.637, RED);
  tx(s, 0.876, 8.317, 3.594, 0.404, 'Jagger Jacson', WNAME);
  tx(s, 0.876, 8.721, 3.594, 0.37, 'Executive Chef', WROLE);
  tx(s, 5.013, 8.317, 3.594, 0.404, 'Jay Kiana', WNAME);
  tx(s, 5.013, 8.721, 3.594, 0.37, 'Kitchen Manager', WROLE);
  tx(s, 9.139, 8.317, 3.594, 0.404, 'Junior Kelvin', WNAME);
  tx(s, 9.139, 8.721, 3.594, 0.37, 'Sous-chef', WROLE);
  tx(s, 13.289, 8.317, 3.594, 0.404, 'Kiara Kimber', WNAME);
  tx(s, 13.289, 8.721, 3.594, 0.37, 'Assistant Cooks', WROLE);
}

function slide11(p) {
  const s = p.addSlide();
  rail(s, '11', 'L', DARK);
  tx(s, 5.095, 1.296, 10.958, 0.95, 'Introducing Our |Team|', HEADC);
  tx(s, 5.095, 0.896, 10.958, 0.404, 'Team', EYEBROWC);
  tx(s, 5.095, 2.538, 10.958, 0.871, L12, BODYC);
  rbox(s, 8.2, 5.885, 4.748, 4.388, 0.79, RED);
  oval(s, 8.769, 4.607, 3.61, W);
  tx(s, 8.777, 8.456, 3.594, 0.404, 'Lady Laura', WNAME);
  tx(s, 8.777, 8.859, 3.594, 0.37, 'Head Waiter', WROLE);
  rbox(s, 2.286, 5.885, 4.748, 4.388, 0.79, DARK);
  oval(s, 2.855, 4.607, 3.61, W);
  tx(s, 2.863, 8.456, 3.594, 0.404, 'Layton Leivent', WNAME);
  tx(s, 2.863, 8.859, 3.594, 0.37, 'Receptionist', WROLE);
  rbox(s, 14.104, 5.885, 4.748, 4.388, 0.79, DARK);
  oval(s, 14.673, 4.607, 3.61, W);
  tx(s, 14.681, 8.456, 3.594, 0.404, 'Melden Markues', WNAME);
  tx(s, 14.681, 8.859, 3.594, 0.37, 'Waiters', WROLE);
  oval(s, 14.87, 4.804, 3.216, PHOTO2);
  oval(s, 8.966, 4.804, 3.216, PHOTO);
  oval(s, 3.052, 4.804, 3.216, PHOTO2);
  [3.895, 9.621, 15.713].forEach(function (x) { social(s, x, 9.423); });
}

function slide12(p) {
  const s = p.addSlide();
  rail(s, '12', 'R', DARK);
  rbox(s, 5.235, 3.488, 4.433, 5.764, 0.268, DARK);
  rbox(s, 0.438, 0.6, 4.433, 5.764, 0.268, DARK);
  ph(s, 5.429, 3.685, 4.045, 5.371, PHOTO, 0.245);
  ph(s, 0.632, 0.797, 4.045, 5.371, PHOTO2, 0.245);
  tx(s, 10.84, 2.671, 7.119, 0.404, 'Team', EYEBROW);
  tx(s, 10.84, 3.075, 7.119, 1.96, 'The Best |Team| We Have', HEAD);
  tx(s, 10.84, 5.536, 7.119, 1.68, 'PLACEHOLDER', BODY);
  rbox(s, 10.84, 7.716, 2.327, 0.863, 0.138, RED);
  tx(s, 10.84, 7.716, 2.327, 0.863, 'Read More', BTN);
  oval(s, 0.947, 4.224, 3.416, RED);
  oval(s, 5.743, 6.947, 3.416, RED);
  tx(s, 0.857, 5.297, 3.594, 0.404, 'Noren Nevaen', WNAME);
  tx(s, 0.857, 5.701, 3.594, 0.37, 'Cleaning Team', WROLE);
  tx(s, 5.677, 8.02, 3.594, 0.404, 'Poppy Olivia', WNAME);
  tx(s, 5.677, 8.423, 3.594, 0.37, 'Sommeliers', WROLE);
  social(s, 1.889, 6.337);
  social(s, 6.686, 9.055);
}

function slide13(p) {
  const s = p.addSlide();
  rail(s, '13', 'R', DARK);
  poly(s, 3.896, 0.896, 5.031, 10.354, DARK, [[0, 0], [4.118, 0], [5.031, 0.913, 4.622, 0, 5.031, 0.409], [5.031, 10.354], [0, 10.354]]);
  box(s, 0, 0.896, 4.25, 10.354, RED);
  tx(s, 9.837, 2.149, 8.017, 2.121, 'Xiandra |Zeidan| William', { sz: 60, f: MONT });
  tx(s, 9.837, 1.743, 8.017, 0.404, 'Team', EYEBROW);
  tx(s, 9.837, 4.772, 8.017, 0.404, 'Meaty Owner', st(LABEL, { sp: 1 }));
  tx(s, 9.837, 5.175, 8.017, 1.276, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore et dolore magna aliqua Ut enim ad minim veniam, quis consectetur adipiscing elit sed consectetur adipiscing elit seni', BODY);
  s.addShape('roundRect', { x: 9.838, y: 7.571, w: 7.979, h: 0.376, rectRadius: 0.188, fill: { color: W }, line: { color: 'F2F2F2', width: 1 } });
  tx(s, 9.838, 7.571, 7.979, 0.376, '90%', { sz: 16, f: DMM, al: 'r', v: 'm' });
  rbox(s, 10.003, 7.692, 6.985, 0.134, 0.067, RED);
  tx(s, 9.846, 7.045, 7.971, 0.404, 'Finance', st(LABEL, { sp: 1 }));
  s.addShape('roundRect', { x: 9.846, y: 9.131, w: 7.971, h: 0.376, rectRadius: 0.188, fill: { color: W }, line: { color: 'F2F2F2', width: 1 } });
  tx(s, 9.846, 9.131, 7.971, 0.376, '80%', { sz: 16, f: DMM, al: 'r', v: 'm' });
  rbox(s, 10.003, 9.252, 6.352, 0.134, 0.067, RED);
  tx(s, 9.846, 8.605, 7.971, 0.404, 'Invest', st(LABEL, { sp: 1 }));
  ph(s, 0.724, 1.797, 7.479, 8.552, PHOTO2, 0.163);
}

function slide14(p) {
  const s = p.addSlide();
  rail(s, '14', 'R', DARK);
  ph(s, 0.478, -0.677, 7.876, 11.448, PHOTO2, 0.359);
  rbox(s, 2.516, 0.762, 6.984, 2.229, 0.357, RED);
  rbox(s, 2.516, 3.261, 6.984, 2.229, 0.357, DARK);
  rbox(s, 2.516, 5.76, 6.984, 2.229, 0.357, RED);
  rbox(s, 2.516, 8.258, 6.984, 2.229, 0.357, DARK);
  tx(s, 10.396, 2.208, 7.458, 0.404, 'Menu', EYEBROW);
  tx(s, 10.396, 2.614, 7.458, 2.97, 'Some of the |Menus| We Provide', HEAD);
  tx(s, 10.396, 6.089, 7.458, 1.68, 'PLACEHOLDER', BODY);
  rbox(s, 10.396, 8.179, 2.327, 0.863, 0.138, RED);
  tx(s, 10.396, 8.179, 2.327, 0.863, 'Read More', BTN);
  tx(s, 4.799, 1.647, 4.397, 0.872, L2, WBODY);
  tx(s, 4.799, 1.235, 4.397, 0.404, 'Fried Noodles', WTITLE);
  tx(s, 4.799, 4.146, 4.397, 0.872, L2, WBODY);
  tx(s, 4.799, 3.734, 4.397, 0.404, 'Grilled Chicken', WTITLE);
  tx(s, 4.799, 6.645, 4.397, 0.872, L2, WBODY);
  tx(s, 4.799, 6.232, 4.397, 0.404, 'Fried Rice', WTITLE);
  tx(s, 4.799, 9.143, 4.397, 0.872, L2, WBODY);
  tx(s, 4.799, 8.731, 4.397, 0.404, 'Salad', WTITLE);
  oval(s, 2.821, 0.976, 1.802, PHOTO);
  oval(s, 2.821, 3.475, 1.802, PHOTO2);
  oval(s, 2.821, 5.973, 1.802, PHOTO);
  oval(s, 2.821, 8.472, 1.802, PHOTO2);
}

function slide15(p) {
  const s = p.addSlide();
  rail(s, '15', 'L', DARK);
  rbox(s, 11.073, 5.661, 8.031, 4.674, 0.166, RED);
  rbox(s, 2.146, 1.775, 5.354, 2.99, 0.274, DARK);
  rbox(s, 7.948, 1.775, 5.354, 2.99, 0.274, DARK);
  rbox(s, 13.75, 1.775, 5.354, 2.99, 0.274, DARK);
  oval(s, 2.365, 0.916, 1.781, RED);
  oval(s, 8.163, 0.916, 1.781, RED);
  oval(s, 13.962, 0.916, 1.781, RED);
  tx(s, 2.625, 3.392, 4.397, 0.872, L2, WBODY);
  tx(s, 2.625, 2.964, 4.397, 0.404, 'Signature Steak', WTITLE);
  tx(s, 8.427, 3.392, 4.397, 0.872, L2, WBODY);
  tx(s, 8.427, 2.964, 4.397, 0.404, 'Sarlion Wagyu', WTITLE);
  tx(s, 14.229, 3.392, 4.397, 0.872, L2, WBODY);
  tx(s, 14.229, 2.964, 4.397, 0.404, 'Grilled Tuna', WTITLE);
  tx(s, 2.142, 5.927, 8.035, 0.404, 'Menu', EYEBROW);
  tx(s, 2.142, 6.331, 8.035, 1.96, 'Menus| at Our Restaurant|', { sz: 60, f: SORA, c: RED, a: DARK, v: 'm', ls: 1, m: 0 });
  tx(s, 2.142, 8.792, 8.035, 1.276, L6, BODY);
  oval(s, 2.477, 1.028, 1.556, PHOTO2);
  oval(s, 8.276, 1.028, 1.556, PHOTO2);
  oval(s, 14.074, 1.028, 1.556, PHOTO);
  ph(s, 11.25, 5.817, 7.677, 4.361, PHOTO, 0.155);
}

function slide16(p) {
  const s = p.addSlide();
  rail(s, '16', 'R', DARK);
  tx(s, 3.896, 1.927, 10.958, 0.95, 'Our |Favorite| Menu', HEADC);
  tx(s, 3.896, 1.527, 10.958, 0.404, 'Menu', EYEBROWC);
  tx(s, 3.896, 3.169, 10.958, 1.276, L6, BODYC);
  rbox(s, 1.01, 5.222, 7.804, 4.502, 0.345, RED);
  rbox(s, 9.937, 5.222, 7.804, 4.502, 0.345, RED);
  tx(s, 4.586, 7.048, 3.716, 1.276, L13, WBODY);
  tx(s, 4.586, 6.621, 3.716, 0.404, 'Spaghetti Bolognese', WTITLE);
  tx(s, 13.56, 7.048, 3.716, 1.276, L13, WBODY);
  tx(s, 13.56, 6.621, 3.716, 0.404, 'Italian Beef Soup & Rice', WTITLE);
  ph(s, 1.27, 5.491, 2.922, 3.963, PHOTO, 0.224);
  ph(s, 10.256, 5.491, 2.922, 3.963, PHOTO2, 0.224);
}

function slide17(p) {
  const s = p.addSlide();
  rail(s, '17', 'L', DARK);
  box(s, 16.98, 0, 3.02, 11.25, RED);
  tx(s, 2.146, 2.671, 8.031, 0.404, 'Menu', EYEBROW);
  tx(s, 2.146, 3.075, 8.031, 1.96, 'Various Kinds of |Desserts|', HEAD);
  tx(s, 2.146, 5.536, 8.031, 1.68, L14, BODY);
  rbox(s, 2.146, 7.716, 2.327, 0.863, 0.138, RED);
  tx(s, 2.146, 7.716, 2.327, 0.863, 'Read More', BTN);
  box(s, 15.321, 0.896, 3.952, 9.458, DARK);
  box(s, 11.769, 0.896, 3.552, 9.458, RED);
  ph(s, 12.169, 1.31, 6.704, 8.63, PHOTO2, 0);
}

function slide18(p) {
  const s = p.addSlide();
  rail(s, '18', 'R', W);
  ph(s, 0, 0.005, 20, 11.245, PHOTO);
  s.addShape('rect', { x: 0, y: 0, w: 20, h: 11.245, fill: { color: DARK, transparency: 10 } });
  oval(s, 0.896, 2.927, 5.396, RED);
  tx(s, 6.851, 4.253, 11.003, 2.339, 'Break Slide', { sz: 133, f: SORA, c: W });
  tx(s, 6.851, 6.593, 11.003, 0.404, 'Please Take A Break and Come Back After 30 Minutes', { f: DMM, c: W, al: 'c' });
  cloche(s, 1.771, 3.6, 3.646, W);
}

function slide19(p) {
  const s = p.addSlide();
  rail(s, '19', 'L', DARK);
  rbox(s, 11.073, 4.148, 8.031, 6.201, 0.544, RED);
  tx(s, 2.146, 2.671, 8.031, 0.404, 'Portfolio', EYEBROW);
  tx(s, 2.146, 3.075, 8.031, 1.96, 'Meaty |Restaurant Portfolio|', { sz: 60, f: SORA, c: RED, a: DARK, v: 'm', ls: 1, m: 0 });
  tx(s, 2.146, 5.536, 8.031, 1.68, L14, BODY);
  rbox(s, 2.146, 7.716, 2.327, 0.863, 0.138, RED);
  tx(s, 2.146, 7.716, 2.327, 0.863, 'Read More', BTN);
  poly(s, 11.333, 0.896, 7.51, 9.133, PHOTO2, [[0.763, 0], [6.747, 0], [7.269, 0.544, 7.036, 0, 7.269, 0.243], [7.269, 3.642], [7.295, 3.656], [7.51, 4.06, 7.425, 3.743, 7.51, 3.892], [7.51, 8.645], [7.022, 9.133, 7.51, 8.915, 7.292, 9.133], [0.488, 9.133], [0, 8.645, 0.218, 9.133, 0, 8.915], [0, 4.06], [0.215, 3.656, 0, 3.892, 0.085, 3.743], [0.241, 3.642], [0.241, 0.544], [0.763, 0, 0.241, 0.243, 0.474, 0]]);
}

function slide20(p) {
  const s = p.addSlide();
  rail(s, '20', 'R', DARK);
  poly(s, 0, 0.437, 8.927, 10.813, RED, [[0, 0], [8.004, 0], [8.927, 0.923, 8.514, 0, 8.927, 0.413], [8.927, 10.813], [0, 10.813]]);
  tx(s, 9.823, 1.216, 8.031, 0.404, 'Portfolio', EYEBROW);
  tx(s, 9.823, 1.62, 8.031, 1.96, 'Portfolio that We Have |Owned|', HEAD);
  tx(s, 9.823, 4.081, 8.031, 0.872, L7, BODY);
  ph(s, 5.281, 1.056, 3.163, 4.056, PHOTO2, 0.191);
  ph(s, 0.903, 1.056, 3.923, 4.056, PHOTO2, 0.236);
  ph(s, 0.896, 5.532, 7.549, 4.662, PHOTO, 0.281);
  ph(s, 9.34, 5.532, 8.514, 4.662, PHOTO2, 0.281);
}

function slide21(p) {
  const s = p.addSlide();
  rail(s, '21', 'L', DARK);
  rbox(s, 2.146, 0.896, 8.031, 4.729, 0.369, DARK);
  rbox(s, 11.073, 6.065, 8.031, 4.287, 0.334, RED);
  tx(s, 2.146, 6.33, 8.031, 0.404, 'Portfolio', EYEBROW);
  tx(s, 2.146, 6.734, 8.031, 1.96, 'The Best |Portfolio| We\'ve Ever Had', HEAD);
  tx(s, 2.146, 9.195, 8.031, 0.872, L7, BODY);
  ph(s, 2.391, 1.135, 7.542, 4.25, PHOTO, 0.332);
  ph(s, 11.073, 0.914, 3.781, 4.217, PHOTO2, 0.295);
  ph(s, 15.573, 0.914, 3.531, 4.217, PHOTO, 0.275);
  ph(s, 11.251, 6.26, 7.675, 3.896, PHOTO2, 0.304);
}

function slide22(p) {
  const s = p.addSlide();
  rail(s, '22', 'R', DARK);
  rbox(s, 9.427, 5.877, 8.427, 4.666, 0.364, RED);
  tx(s, 9.82, 9.954, 7.602, 0.404, 'Gallery 3', { f: DMM, c: W, al: 'c', sp: 2 });
  rbox(s, 0.896, 5.877, 8.031, 4.666, 0.364, DARK);
  tx(s, 1.289, 9.954, 7.244, 0.404, 'Gallery 2', { f: DMM, c: W, al: 'c', sp: 2 });
  rbox(s, 0.896, 0.705, 8.031, 4.666, 0.364, RED);
  tx(s, 9.823, 1.058, 8.031, 0.404, 'Portfolio', EYEBROW);
  tx(s, 9.823, 1.462, 8.031, 1.96, 'Gallery| that We Have Owned|', { sz: 60, f: SORA, c: RED, a: DARK, v: 'm', ls: 1, m: 0 });
  tx(s, 9.823, 3.922, 8.031, 0.872, L7, BODY);
  tx(s, 1.289, 4.782, 7.244, 0.404, 'Gallery 1', { f: DMM, c: W, al: 'c', sp: 2 });
  ph(s, 1.166, 0.975, 7.49, 3.545, PHOTO2, 0.277);
  ph(s, 1.166, 6.148, 7.49, 3.545, PHOTO, 0.277);
  ph(s, 9.701, 6.148, 7.86, 3.545, PHOTO2, 0.277);
}

function slide23(p) {
  const s = p.addSlide();
  rail(s, '23', 'L', DARK);
  tx(s, 5.146, 1.294, 10.958, 0.95, 'Our |Favorite| Menu', HEADC);
  tx(s, 5.146, 0.895, 10.958, 0.404, 'Portfolio', EYEBROWC);
  tx(s, 5.146, 2.536, 10.958, 1.276, L6, BODYC);
  rbox(s, 2.146, 5.608, 7.491, 4.748, 0.555, RED);
  rbox(s, 11.073, 5.608, 7.491, 4.748, 0.555, RED);
  tx(s, 2.269, 9.797, 7.244, 0.404, 'Activity 1', { f: DMM, c: W, al: 'c', sp: 2 });
  tx(s, 11.196, 9.797, 7.244, 0.404, 'Activity 2', { f: DMM, c: W, al: 'c', sp: 2 });
  ph(s, 2.333, 5.781, 7.117, 3.791, PHOTO, 0.443);
  ph(s, 11.26, 5.781, 7.117, 3.791, PHOTO, 0.443);
  rbox(s, 4.213, 4.696, 5.964, 3.688, 0.44, RED);
  rbox(s, 13.154, 4.696, 5.964, 3.688, 0.534, RED);
  ph(s, 4.364, 4.848, 5.662, 3.382, PHOTO2, 0.396);
  ph(s, 13.305, 4.848, 5.662, 3.382, PHOTO2, 0.396);
}

function slide24(p) {
  const s = p.addSlide();
  rail(s, '24', 'R', DARK);
  poly(s, 0, 0, 8.927, 5.615, '3A3A3A', [[0, 0], [8.927, 0], [8.927, 5.252], [8.563, 5.615, 8.927, 5.452, 8.764, 5.615], [0, 5.615]]);
  rbox(s, 9.823, 0.896, 8.365, 4.62, 0.207, RED);
  tx(s, 0.896, 6.511, 8.031, 0.404, 'Portfolio', EYEBROW);
  tx(s, 0.896, 6.915, 8.031, 1.96, 'The Best |Gallery| We Have', HEAD);
  tx(s, 0.896, 9.375, 8.031, 0.872, L7, BODY);
  poly(s, -0.006, 0, 8.724, 5.407, PHOTO2, [[0, 0], [8.724, 0], [8.724, 5.025], [8.357, 5.407, 8.724, 5.236, 8.56, 5.407], [0, 5.407]]);
  ph(s, 10.031, 1.1, 7.948, 4.207, PHOTO, 0.189);
  poly(s, 9.823, 6.417, 8.365, 4.833, PHOTO2, [[0.294, 0], [8.071, 0], [8.365, 0.282, 8.233, 0, 8.365, 0.126], [8.365, 4.833], [0, 4.833], [0, 0.282], [0.294, 0, 0, 0.126, 0.132, 0]]);
}

function slide25(p) {
  const s = p.addSlide();
  rail(s, '25', 'L', DARK);
  laptop(s, 2.192, 5.29, 6.626, 4.51);
  rbox(s, 9.341, 5.267, 4.513, 4.533, 0.277, RED);
  rbox(s, 14.507, 5.267, 4.513, 4.533, 0.277, DARK);
  tx(s, 5.126, 1.85, 10.958, 0.95, 'Restaurant |Collaboration|', HEADC);
  tx(s, 5.126, 1.45, 10.958, 0.404, 'Mockup Device', EYEBROWC);
  tx(s, 5.126, 3.092, 10.958, 1.276, L6, BODYC);
  tx(s, 9.671, 5.826, 3.853, 1.111, '+180', { sz: 60, f: MONT, c: W, al: 'c' });
  tx(s, 9.671, 7.157, 3.853, 0.404, 'Restaurant Branch', WNAME);
  tx(s, 9.671, 7.561, 3.853, 1.68, L15, WBODYC);
  tx(s, 14.836, 5.826, 3.853, 1.111, '+500', { sz: 60, f: MONT, c: W, al: 'c' });
  tx(s, 14.836, 7.157, 3.853, 0.404, 'Investor', WNAME);
  tx(s, 14.836, 7.561, 3.853, 1.68, L15, WBODYC);
}

function slide26(p) {
  const s = p.addSlide();
  rail(s, '26', 'R', DARK);
  monitor(s, 1.385, 2.749, 8.112, 6.608);
  tx(s, 10.886, 2.536, 6.966, 0.404, 'Mokup Device', EYEBROW);
  tx(s, 10.886, 2.94, 6.966, 1.96, 'Customer |Ratings|', HEAD);
  tx(s, 10.886, 5.401, 6.966, 1.68, 'PLACEHOLDER', BODY);
  rbox(s, 10.886, 7.582, 2.327, 0.863, 0.138, RED);
  tx(s, 10.886, 7.582, 2.327, 0.863, 'Read More', BTN);
  rbox(s, 6.29, 1.667, 3.696, 3.062, 0.469, RED);
  rbox(s, 0.898, 6.52, 3.696, 3.062, 0.469, RED);
  tx(s, 1.349, 7.043, 2.796, 0.404, 'Our Food Ratings', { f: DMM, c: W, al: 'c' });
  s.addShape('star5', { x: 1.704, y: 7.551, w: 0.31, h: 0.31, fill: { color: GOLD } });
  s.addShape('star5', { x: 2.148, y: 7.551, w: 0.31, h: 0.31, fill: { color: GOLD } });
  s.addShape('star5', { x: 2.592, y: 7.551, w: 0.31, h: 0.31, fill: { color: GOLD } });
  s.addShape('star5', { x: 3.036, y: 7.551, w: 0.31, h: 0.31, fill: { color: GOLD } });
  s.addShape('star5', { x: 3.48, y: 7.551, w: 0.31, h: 0.31, fill: { color: GOLD } });
  tx(s, 1.227, 8.187, 3.039, 0.872, L16, WBODYC);
  tx(s, 6.74, 2.191, 2.796, 0.404, 'Rate of Service', { f: DMM, c: W, al: 'c' });
  s.addShape('star5', { x: 7.096, y: 2.698, w: 0.31, h: 0.31, fill: { color: GOLD } });
  s.addShape('star5', { x: 7.54, y: 2.698, w: 0.31, h: 0.31, fill: { color: GOLD } });
  s.addShape('star5', { x: 7.984, y: 2.698, w: 0.31, h: 0.31, fill: { color: GOLD } });
  s.addShape('star5', { x: 8.428, y: 2.698, w: 0.31, h: 0.31, fill: { color: GOLD } });
  s.addShape('star5', { x: 8.872, y: 2.698, w: 0.31, h: 0.31, fill: { color: GOLD } });
  tx(s, 6.619, 3.335, 3.039, 0.872, L16, WBODYC);
}

function slide27(p) {
  const s = p.addSlide();
  rail(s, '27', 'L', DARK);
  phone(s, 14.93, 0.577, 4.174, 8.546);
  phone(s, 10.625, 2.127, 4.174, 8.546);
  oval(s, 15.349, -2.872, 8.406, RED);
  tx(s, 2.146, 3.151, 7.518, 0.404, 'Mockup Device', EYEBROW);
  tx(s, 2.146, 3.555, 7.518, 1.96, 'Meaty| Social Media|', { sz: 60, f: SORA, c: RED, a: DARK, v: 'm', ls: 1, m: 0 });
  tx(s, 2.146, 6.016, 7.518, 2.084, L14 + ' amet consectetur adipiscing elit sed dol', BODY);
  rbox(s, 14.132, 6.088, 1.362, 1.362, 0.218, DARK);
  rbox(s, 14.805, 4.034, 1.15, 1.15, 0.184, RED);
  rbox(s, 13.198, 2.037, 1.652, 1.652, 0.264, DARK);
  socialTile(s, 'ig', 13.198, 2.037, 1.652, DARK);
  socialTile(s, 'fb', 14.805, 4.034, 1.15, RED);
  socialTile(s, 'in', 14.132, 6.088, 1.362, DARK);
}

function slide28(p) {
  const s = p.addSlide();
  rail(s, '28', 'R', DARK);
  rbox(s, 0.927, 1.222, 8.488, 8.806, 0.548, MIST);
  tx(s, 1.677, 8.359, 6.988, 0.471, 'Business', st(LABEL, { v: 'm', sp: 1.5, m: 0 }));
  tx(s, 1.677, 2.823, 6.988, 2.121, 'SWOT |Analysis| Infographics', { sz: 60, f: SORA });
  tx(s, 1.677, 5.442, 6.988, 1.68, 'PLACEHOLDER', BODY);
  tx(s, 1.677, 2.419, 6.988, 0.404, 'Infographics', EYEBROW);
  s.addShape('roundRect', { x: 1.677, y: 7.762, w: 6.988, h: 0.455, rectRadius: 0.228, fill: { color: W }, line: { color: 'F2F2F2', width: 1 } });
  tx(s, 1.677, 7.762, 6.988, 0.455, '90%', { f: DMM, al: 'r', v: 'm' });
  rbox(s, 1.796, 7.894, 5.727, 0.191, 0.096, RED);
  poly(s, 14.484, 2.532, 3.36, 3.007, DARK, [[0.899, 0], [0, 0.899, 0.402, 0, 0, 0.403], [0, 3.007], [2.461, 3.007], [3.36, 2.108, 2.958, 3.007, 3.36, 2.604], [3.36, 0]]);
  poly(s, 14.484, 5.711, 3.36, 3.007, RED, [[0, 0], [0, 2.108], [0.899, 3.007, 0, 2.604, 0.402, 3.007], [3.36, 3.007], [3.36, 0.899], [2.461, 0, 3.36, 0.403, 2.958, 0]]);
  poly(s, 10.982, 2.532, 3.36, 3.007, RED, [[0, 0], [0, 2.108], [0.899, 3.007, 0, 2.604, 0.402, 3.007], [3.36, 3.007], [3.36, 0.899], [2.461, 0, 3.36, 0.403, 2.958, 0]]);
  poly(s, 10.982, 5.711, 3.36, 3.007, DARK, [[0.899, 0], [0, 0.899, 0.402, 0, 0, 0.403], [0, 3.007], [2.461, 3.007], [3.36, 2.108, 2.958, 3.007, 3.36, 2.604], [3.36, 0]]);
  oval(s, 15.522, 1.891, 1.284, W);
  oval(s, 15.679, 2.047, 0.97, RED);
  tx(s, 15.679, 2.047, 0.97, 0.97, 'W', { sz: 36, f: SORA, c: W, al: 'c', v: 'm' });
  oval(s, 15.522, 8.076, 1.284, W);
  oval(s, 15.679, 8.232, 0.97, DARK);
  tx(s, 15.679, 8.232, 0.97, 0.97, 'T', { sz: 36, f: SORA, c: W, al: 'c', v: 'm' });
  oval(s, 12.021, 1.891, 1.284, W);
  oval(s, 12.177, 2.047, 0.97, DARK);
  tx(s, 12.177, 2.047, 0.97, 0.97, 'S', { sz: 36, f: SORA, c: W, al: 'c', v: 'm' });
  oval(s, 12.021, 8.076, 1.284, W);
  oval(s, 12.177, 8.232, 0.97, RED);
  tx(s, 12.177, 8.232, 0.97, 0.97, 'O', { sz: 36, f: SORA, c: W, al: 'c', v: 'm' });
  oval(s, 12.867, 4.079, 3.092, W);
  poly(s, 13.102, 4.314, 2.621, 2.621, MIST, [[2.621, 1.31], [1.311, 2.621, 2.621, 2.034, 2.034, 2.621], [0, 1.31, 0.587, 2.621, 0, 2.034], [1.311, 0, 0, 0.587, 0.587, 0], [2.621, 1.31, 2.034, 0, 2.621, 0.587]]);
  splat(s, 13.763, 5.002, 1.3);
  [['chart', 12.005, 3.581], ['medal', 15.509, 3.567], ['print', 11.65, 6.697],
   ['rocket', 15.509, 6.697]].forEach(function (g) { glyph(s, g[0], g[1], g[2], 0.96, W); });
}

function slide29(p) {
  const s = p.addSlide();
  rail(s, '29', 'L', DARK);
  poly(s, 10.693, 6.784, 2.507, 2.507, RED, [[0.953, 0], [0, 0.954, 0.925, 0.514, 0.514, 0.926], [0, 2.507], [1.754, 1.759, 0.684, 2.493, 1.303, 2.209], [1.697, 1.506, 1.718, 1.682, 1.697, 1.596], [2.284, 0.918, 1.697, 1.181, 1.96, 0.918], [2.316, 0.919, 2.295, 0.918, 2.305, 0.919], [2.507, 0, 2.432, 0.635, 2.5, 0.325]]);
  tx(s, 10.693, 6.784, 2.507, 2.507, 'T', { sz: 54, f: SORA, c: W, al: 'c', v: 'm' });
  oval(s, 12.504, 7.816, 0.947, DARK);
  rbox(s, 13.858, 7.146, 5.057, 2.686, 0.206, MIST);
  poly(s, 10.693, 4.165, 2.507, 2.507, DARK, [[0, 0], [0, 1.553], [0.953, 2.507, 0.514, 1.581, 0.925, 1.993], [2.507, 2.507], [2.316, 1.588, 2.5, 2.182, 2.432, 1.872], [2.284, 1.589, 2.305, 1.588, 2.295, 1.589], [1.697, 1.002, 1.96, 1.589, 1.697, 1.326], [1.754, 0.749, 1.697, 0.911, 1.718, 0.826], [0, 0, 1.303, 0.298, 0.684, 0.015]]);
  tx(s, 10.693, 4.165, 2.507, 2.507, 'W', { sz: 54, f: SORA, c: W, al: 'c', v: 'm' });
  oval(s, 12.504, 4.694, 0.947, RED);
  rbox(s, 13.858, 3.624, 5.057, 2.686, 0.206, MIST);
  poly(s, 8.074, 6.784, 2.507, 2.507, DARK, [[0, 0], [0.192, 0.919, 0.007, 0.325, 0.074, 0.635], [0.223, 0.918, 0.202, 0.919, 0.212, 0.918], [0.809, 1.506, 0.547, 0.918, 0.809, 1.181], [0.752, 1.759, 0.809, 1.596, 0.789, 1.682], [2.507, 2.507, 1.204, 2.209, 1.823, 2.493], [2.507, 0.954], [1.553, 0, 1.993, 0.926, 1.581, 0.514]]);
  tx(s, 8.074, 6.784, 2.507, 2.507, 'O', { sz: 54, f: SORA, c: W, al: 'c', v: 'm' });
  oval(s, 7.824, 7.816, 0.947, RED);
  rbox(s, 2.335, 7.146, 5.057, 2.686, 0.206, MIST);
  poly(s, 8.074, 4.165, 2.507, 2.507, RED, [[2.507, 0], [0.752, 0.749, 1.822, 0.015, 1.204, 0.298], [0.809, 1.002, 0.789, 0.826, 0.809, 0.911], [0.223, 1.589, 0.809, 1.326, 0.547, 1.589], [0.192, 1.588, 0.212, 1.589, 0.202, 1.588], [0, 2.507, 0.074, 1.872, 0.007, 2.182], [1.553, 2.507], [2.507, 1.553, 1.581, 1.993, 1.993, 1.581], [2.507, 0]]);
  tx(s, 8.074, 4.165, 2.507, 2.507, 'S', { sz: 54, f: SORA, c: W, al: 'c', v: 'm' });
  oval(s, 7.824, 4.694, 0.947, DARK);
  rbox(s, 2.335, 3.624, 5.057, 2.686, 0.206, MIST);
  tx(s, 4.149, 1.815, 12.976, 1.111, 'SWOT| Analysis Infographics|', { sz: 60, f: SORA, c: RED, a: DARK, al: 'c' });
  tx(s, 4.149, 1.418, 12.976, 0.404, 'Infographics', EYEBROWC);
  tx(s, 14.551, 4.532, 3.67, 1.275, L1, BODYL);
  tx(s, 14.551, 4.128, 3.67, 0.404, 'Weaknesses', LABEL);
  tx(s, 14.551, 8.054, 3.67, 1.275, L1, BODYL);
  tx(s, 14.551, 7.65, 3.67, 0.404, 'Threats', LABEL);
  tx(s, 3.029, 4.532, 3.67, 1.276, L1, BODYR);
  tx(s, 3.029, 4.128, 3.67, 0.404, 'Strenghts', st(LABEL, { al: 'r' }));
  tx(s, 3.029, 8.054, 3.67, 1.275, L1, BODYR);
  tx(s, 3.029, 7.65, 3.67, 0.404, 'Opportunities', st(LABEL, { al: 'r' }));
  [['chart', 7.997, 4.867], ['medal', 12.678, 4.868], ['print', 7.998, 7.99],
   ['rocket', 12.678, 7.99]].forEach(function (g) { glyph(s, g[0], g[1], g[2], 0.6, W); });
}

function slide30(p) {
  const s = p.addSlide();
  rail(s, '30', 'R', DARK);
  quatrefoil(s, 6.538, 4.004, 5.674, DARK);
  s.addShape('custGeom', { x: 10.243, y: 4.504, w: 3.35, h: 0.503, points: pts([[0, 0], [0.561, 0], [0.561, 0.503], [3.35, 0.503]]), line: { color: 'BFBFBF', width: 0.75 }, flipH: true });
  s.addShape('custGeom', { x: 12.212, y: 6.794, w: 1.381, h: 0.94, points: pts([[0, 0], [0.814, 0], [0.814, 0.94], [1.381, 0.94]]), line: { color: 'BFBFBF', width: 0.75 } });
  s.addShape('custGeom', { x: 5.166, y: 7.747, w: 3.35, h: 1.199, points: pts([[0, 0], [0.561, 0], [0.561, 1.199], [3.35, 1.199]]), line: { color: 'BFBFBF', width: 0.75 } });
  s.addShape('custGeom', { x: 5.157, y: 4.504, w: 1.381, h: 2.304, points: pts([[0, 0], [0.814, 0], [0.814, 2.304], [1.381, 2.304]]), line: { color: 'BFBFBF', width: 0.75 }, rotate: 180 });
  oval(s, 8.436, 4.053, 1.886, W);
  oval(s, 8.772, 4.389, 1.214, 'EEEEEE');
  oval(s, 8.837, 4.455, 1.083, RED);
  tx(s, 8.837, 4.455, 1.083, 1.083, 'W', { sz: 50, f: SORA, c: W, al: 'c', v: 'm' });
  oval(s, 10.277, 5.894, 1.886, W);
  oval(s, 10.613, 6.23, 1.214, 'EEEEEE');
  oval(s, 10.678, 6.295, 1.084, RED);
  tx(s, 10.678, 6.295, 1.084, 1.084, 'T', { sz: 50, f: SORA, c: W, al: 'c', v: 'm' });
  oval(s, 8.436, 7.753, 1.886, W);
  oval(s, 8.772, 8.089, 1.214, 'EEEEEE');
  oval(s, 8.837, 8.154, 1.083, RED);
  tx(s, 8.837, 8.154, 1.083, 1.083, 'O', { sz: 50, f: SORA, c: W, al: 'c', v: 'm' });
  oval(s, 6.588, 5.894, 1.885, W);
  oval(s, 6.924, 6.23, 1.214, 'EEEEEE');
  oval(s, 6.989, 6.295, 1.084, RED);
  tx(s, 6.989, 6.295, 1.084, 1.084, 'S', { sz: 50, f: SORA, c: W, al: 'c', v: 'm' });
  tx(s, 2.628, 1.964, 13.518, 1.111, 'SWOT| Analysis Infographics|', { sz: 60, f: SORA, c: RED, a: DARK, al: 'c' });
  tx(s, 2.628, 1.562, 13.518, 0.404, 'Infographics', EYEBROWC);
  tx(s, 13.858, 4.757, 3.67, 1.275, L1, BODYL);
  tx(s, 13.858, 4.353, 3.67, 0.404, 'Weaknesses', LABEL);
  tx(s, 13.798, 7.963, 3.67, 1.275, L1, BODYL);
  tx(s, 13.798, 7.559, 3.67, 0.404, 'Threats', LABEL);
  tx(s, 1.221, 4.757, 3.67, 1.275, L1, BODYR);
  tx(s, 1.221, 4.353, 3.67, 0.404, 'Strenghts', st(LABEL, { al: 'r' }));
  tx(s, 1.221, 7.963, 3.67, 1.275, L1, BODYR);
  tx(s, 1.221, 7.559, 3.67, 0.404, 'Opportunities', st(LABEL, { al: 'r' }));
}

function slide31(p) {
  const s = p.addSlide();
  rail(s, '31', 'L', DARK);
  rbox(s, 10.906, 3.59, 6.423, 2.734, 0.209, MIST);
  rbox(s, 10.906, 6.951, 6.423, 2.734, 0.209, MIST);
  rbox(s, 3.886, 3.59, 6.423, 2.734, 0.209, MIST);
  rbox(s, 3.886, 6.951, 6.423, 2.734, 0.209, MIST);
  poly(s, 17.584, 8.528, 0.987, 0.904, DARK, [[0, 0], [0, 0.405], [0.499, 0.904, 0, 0.681, 0.224, 0.904], [0.987, 0.904], [0.987, 0.614], [0.374, 0, 0.987, 0.274, 0.712, 0]]);
  poly(s, 15.951, 7.238, 2.62, 1.91, RED, [[1.037, 0], [0.002, 0.814], [0.002, 0.819, 0, 0.815, 0, 0.818], [1.037, 1.633], [1.037, 1.297], [2.007, 1.297], [2.62, 1.91, 2.345, 1.297, 2.62, 1.571], [2.62, 0.924], [2.007, 0.31, 2.62, 0.585, 2.345, 0.31], [1.037, 0.31], [1.037, 0]]);
  poly(s, 17.584, 5.14, 0.987, 0.904, RED, [[0, 0], [0, 0.404], [0.499, 0.904, 0, 0.681, 0.224, 0.904], [0.987, 0.904], [0.987, 0.614], [0.374, 0, 0.987, 0.275, 0.712, 0]]);
  poly(s, 15.951, 3.85, 2.62, 1.91, DARK, [[1.037, 0], [0.002, 0.814], [0.002, 0.819, 0, 0.815, 0, 0.818], [1.037, 1.633], [1.037, 1.297], [2.007, 1.297], [2.62, 1.91, 2.345, 1.297, 2.62, 1.571], [2.62, 0.924], [2.007, 0.311, 2.62, 0.585, 2.345, 0.311], [1.037, 0.311], [1.037, 0]]);
  poly(s, 2.678, 5.14, 0.987, 0.904, DARK, [[0.613, 0], [0, 0.614, 0.275, 0, 0, 0.275], [0, 0.904], [0.488, 0.904], [0.987, 0.404, 0.763, 0.904, 0.987, 0.681], [0.987, 0]]);
  poly(s, 2.678, 3.85, 2.621, 1.91, RED, [[1.584, 0], [1.584, 0.311], [0.613, 0.311], [0, 0.924, 0.275, 0.311, 0, 0.585], [0, 1.91], [0.613, 1.297, 0, 1.571, 0.275, 1.297], [1.584, 1.297], [1.584, 1.633], [2.619, 0.819], [2.619, 0.814, 2.621, 0.818, 2.621, 0.815], [1.584, 0]]);
  poly(s, 2.678, 8.528, 0.987, 0.904, RED, [[0.613, 0], [0, 0.614, 0.275, 0, 0, 0.274], [0, 0.904], [0.488, 0.904], [0.987, 0.405, 0.763, 0.904, 0.987, 0.681], [0.987, 0]]);
  poly(s, 2.678, 7.238, 2.621, 1.91, DARK, [[1.584, 0], [1.584, 0.31], [0.613, 0.31], [0, 0.924, 0.275, 0.31, 0, 0.585], [0, 1.91], [0.613, 1.297, 0, 1.571, 0.275, 1.297], [1.584, 1.297], [1.584, 1.633], [2.619, 0.819], [2.619, 0.814, 2.621, 0.818, 2.621, 0.815], [1.584, 0]]);
  tx(s, 3.886, 1.971, 13.478, 1.111, 'SWOT| Analysis Infographics|', { sz: 60, f: SORA, c: RED, a: DARK, al: 'c' });
  tx(s, 3.886, 1.565, 13.478, 0.404, 'Infographics', EYEBROWC);
  tx(s, 11.624, 4.522, 3.67, 1.275, L1, BODYR);
  tx(s, 11.624, 4.118, 3.67, 0.404, 'Weaknesses', st(LABEL, { al: 'r' }));
  tx(s, 11.563, 7.883, 3.67, 1.275, L1, BODYR);
  tx(s, 11.563, 7.479, 3.67, 0.404, 'Threats', st(LABEL, { al: 'r' }));
  tx(s, 5.969, 4.522, 3.67, 1.275, L1, BODYL);
  tx(s, 5.969, 4.118, 3.67, 0.404, 'Strenghts', LABEL);
  tx(s, 5.969, 7.883, 3.67, 1.275, L1, BODYL);
  tx(s, 5.969, 7.479, 3.67, 0.404, 'Opportunities', LABEL);
}

function slide32(p) {
  const s = p.addSlide();
  rail(s, '32', 'R', DARK);
  box(s, 3.387, 3.204, 8.735, 1.633, DARK);
  poly(s, 3.889, 3.377, 0.811, 1.287, W, [[0, 0], [0, 1.287], [0.44, 1.287], [0.811, 0.644], [0.44, 0]]);
  poly(s, 2.262, 2.989, 1.791, 2.063, RED, [[0.602, 0], [0.59, 0.007, 0.597, 0, 0.593, 0.003], [0.003, 1.025], [0.003, 1.039, 0, 1.029, 0, 1.034], [0.59, 2.056], [0.602, 2.063, 0.593, 2.06, 0.597, 2.063], [1.777, 2.063], [1.791, 2.049, 1.784, 2.063, 1.791, 2.057], [1.777, 2.035, 1.791, 2.041, 1.784, 2.035], [0.61, 2.035], [0.031, 1.032], [0.61, 0.028], [1.777, 0.028], [1.791, 0.014, 1.784, 0.028, 1.791, 0.022], [1.777, 0, 1.791, 0.007, 1.784, 0]]);
  poly(s, 2.508, 3.204, 1.886, 1.633, MIST, [[0.471, 0], [0, 0.817], [0.471, 1.633], [1.415, 1.633], [1.886, 0.817], [1.415, 0]]);
  tx(s, 2.508, 3.204, 1.886, 1.633, 'S', { sz: 70, f: POP, c: RED, al: 'c', v: 'm' });
  box(s, 3.387, 6.686, 8.735, 1.633, DARK);
  poly(s, 3.889, 6.859, 0.811, 1.288, W, [[0, 0], [0, 1.288], [0.44, 1.288], [0.811, 0.644], [0.44, 0]]);
  poly(s, 2.262, 6.472, 1.791, 2.063, RED, [[0.602, 0], [0.59, 0.007, 0.597, 0, 0.593, 0.003], [0.003, 1.024], [0.003, 1.038, 0, 1.029, 0, 1.034], [0.59, 2.056], [0.602, 2.063, 0.593, 2.06, 0.597, 2.063], [1.777, 2.063], [1.791, 2.049, 1.784, 2.063, 1.791, 2.057], [1.777, 2.035, 1.791, 2.042, 1.784, 2.035], [0.61, 2.035], [0.031, 1.032], [0.61, 0.028], [1.777, 0.028], [1.791, 0.014, 1.784, 0.028, 1.791, 0.022], [1.777, 0, 1.791, 0.006, 1.784, 0]]);
  poly(s, 2.508, 6.686, 1.886, 1.633, MIST, [[0.471, 0], [0, 0.817], [0.471, 1.633], [1.415, 1.633], [1.886, 0.817], [1.415, 0]]);
  tx(s, 2.508, 6.686, 1.886, 1.633, 'O', { sz: 70, f: POP, c: RED, al: 'c', v: 'm' });
  box(s, 6.628, 4.945, 8.735, 1.633, RED);
  poly(s, 14.049, 5.118, 0.811, 1.287, W, [[0.371, 0], [0, 0.643], [0.371, 1.287], [0.811, 1.287], [0.811, 0]]);
  poly(s, 14.697, 4.73, 1.791, 2.063, DARK, [[0.014, 0], [0, 0.014, 0.007, 0, 0, 0.006], [0.014, 0.028, 0, 0.022, 0.007, 0.028], [1.181, 0.028], [1.76, 1.031], [1.181, 2.035], [0.014, 2.035], [0, 2.049, 0.007, 2.035, 0, 2.041], [0.014, 2.063, 0, 2.056, 0.007, 2.063], [1.189, 2.063], [1.201, 2.056, 1.194, 2.063, 1.199, 2.06], [1.788, 1.038], [1.788, 1.024, 1.791, 1.034, 1.791, 1.029], [1.201, 0.007], [1.189, 0, 1.199, 0.003, 1.194, 0]]);
  poly(s, 14.356, 4.945, 1.886, 1.633, MIST, [[0.471, 0], [0, 0.817], [0.471, 1.633], [1.415, 1.633], [1.886, 0.817], [1.415, 0]]);
  tx(s, 14.356, 4.945, 1.886, 1.633, 'W', { sz: 70, f: POP, al: 'c', v: 'm' });
  box(s, 6.628, 8.428, 8.735, 1.633, RED);
  poly(s, 14.049, 8.601, 0.811, 1.287, W, [[0.371, 0], [0, 0.644], [0.371, 1.287], [0.811, 1.287], [0.811, 0]]);
  poly(s, 14.697, 8.213, 1.791, 2.063, DARK, [[0.014, 0], [0, 0.014, 0.007, 0, 0, 0.007], [0.014, 0.028, 0, 0.022, 0.007, 0.028], [1.181, 0.028], [1.76, 1.032], [1.181, 2.035], [0.014, 2.035], [0, 2.049, 0.007, 2.035, 0, 2.041], [0.014, 2.063, 0, 2.057, 0.007, 2.063], [1.189, 2.063], [1.201, 2.056, 1.194, 2.063, 1.199, 2.06], [1.788, 1.039], [1.788, 1.025, 1.791, 1.034, 1.791, 1.029], [1.201, 0.007], [1.189, 0, 1.199, 0.003, 1.194, 0]]);
  poly(s, 14.356, 8.428, 1.886, 1.633, MIST, [[0.471, 0], [0, 0.817], [0.471, 1.633], [1.415, 1.633], [1.886, 0.817], [1.415, 0]]);
  tx(s, 14.356, 8.428, 1.886, 1.633, 'T', { sz: 70, f: POP, al: 'c', v: 'm' });
  tx(s, 2.896, 1.375, 12.958, 1.111, 'SWOT| Analysis Infographics|', { sz: 60, f: SORA, c: RED, a: DARK, al: 'c' });
  tx(s, 2.896, 0.974, 12.958, 0.404, 'Infographics', EYEBROWC);
  tx(s, 7.045, 5.73, 6.593, 0.467, L8, { sz: 16, c: W, al: 'r', ls: 1.5, sp: 1 });
  tx(s, 7.045, 5.326, 6.593, 0.404, 'Weaknesses', { f: DMM, c: W, al: 'r' });
  tx(s, 7.043, 9.213, 6.593, 0.467, L8, { sz: 16, c: W, al: 'r', ls: 1.5, sp: 1 });
  tx(s, 7.043, 8.809, 6.593, 0.404, 'Threats', { f: DMM, c: W, al: 'r' });
  tx(s, 5.115, 4.01, 6.594, 0.467, L8, { sz: 16, c: W, ls: 1.5, sp: 1 });
  tx(s, 5.115, 3.606, 6.594, 0.404, 'Strenghts', { f: DMM, c: W });
  tx(s, 5.114, 7.443, 6.593, 0.467, L8, { sz: 16, c: W, ls: 1.5, sp: 1 });
  tx(s, 5.114, 7.039, 6.593, 0.404, 'Opportunities', { f: DMM, c: W });
}

function slide33(p) {
  const s = p.addSlide();
  rail(s, '33', 'L', DARK);
  s.addShape('ellipse', { x: 8.236, y: 3.807, w: 4.778, h: 4.778, fill: { color: MIST }, flipH: true });
  s.addShape('pie', { x: 7.496, y: 3.067, w: 6.259, h: 6.259, angleRange: [202.46, 238.076], fill: { color: DARK }, flipH: true });
  s.addShape('pie', { x: 7.496, y: 3.067, w: 6.259, h: 6.259, angleRange: [103.324, 170.884], fill: { color: DARK }, flipH: true });
  s.addShape('pie', { x: 7.496, y: 3.067, w: 6.259, h: 6.259, angleRange: [305.283, 351.319], fill: { color: RED }, flipH: true });
  s.addShape('pie', { x: 6.924, y: 2.495, w: 7.403, h: 7.403, angleRange: [160.504, 192.004], fill: { color: RED }, flipH: true });
  s.addShape('pie', { x: 7.058, y: 2.63, w: 7.134, h: 7.134, angleRange: [339.663, 33.396], fill: { color: DARK }, flipH: true });
  s.addShape('pie', { x: 7.868, y: 3.44, w: 5.513, h: 5.513, angleRange: [27.538, 69.034], fill: { color: RED }, flipH: true });
  s.addShape('ellipse', { x: 9.35, y: 4.921, w: 2.55, h: 2.55, fill: { color: W }, flipH: true });
  tx(s, 15.108, 4.158, 3.67, 1.275, L1, BODYL);
  tx(s, 15.108, 3.754, 3.67, 0.404, 'Your Title', LABEL);
  tx(s, 15.048, 7.364, 3.67, 1.275, L1, BODYL);
  tx(s, 15.048, 6.96, 3.67, 0.404, 'Your Title', LABEL);
  tx(s, 2.471, 4.158, 3.67, 1.275, L1, BODYR);
  tx(s, 2.471, 3.754, 3.67, 0.404, 'Your Title', st(LABEL, { al: 'r' }));
  tx(s, 2.471, 7.364, 3.67, 1.275, L1, BODYR);
  tx(s, 2.471, 6.96, 3.67, 0.404, 'Your Title', st(LABEL, { al: 'r' }));
  tx(s, 4.593, 1.761, 12.065, 1.111, 'Our |Infographics|', { sz: 60, f: SORA, al: 'c' });
  tx(s, 4.593, 1.352, 12.065, 0.404, 'Infographics', EYEBROWC);
  splat(s, 9.975, 5.573, 1.3);
}

function slide34(p) {
  const s = p.addSlide();
  rail(s, '34', 'R', DARK);
  poly(s, 6.39, 6.286, 1.459, 0.156, 'E8E8E9', [[0.005, 0], [0.005, 0.072], [1.459, 0.072], [1.459, 0.039], [1.459, 0], [0, 0.111], [0, 0.156], [0.005, 0.156], [0, 0.111, 0, 0.139, 0, 0.128]]);
  poly(s, 10.911, 6.286, 1.459, 0.072, 'E8E8E9', [[0, 0], [0.005, 0.039, 0, 0.011, 0.005, 0.027], [0, 0.072, 0.005, 0.05, 0, 0.061], [1.459, 0.072], [1.453, 0, 1.459, 0.05, 1.453, 0.027]]);
  poly(s, 6.39, 4.231, 2.906, 2.055, RED, [[0.852, 0], [0, 2.055, 0.334, 0.535, 0.011, 1.258], [2.906, 2.055], [0.852, 0]]);
  poly(s, 6.39, 6.358, 2.906, 2.055, DARK, [[0, 0], [0.852, 2.055, 0.011, 0.802, 0.334, 1.526], [2.906, 0]]);
  poly(s, 9.469, 6.358, 2.901, 2.055, RED, [[0, 0], [2.049, 2.055], [2.901, 0, 2.572, 1.526, 2.895, 0.802]]);
  poly(s, 7.292, 6.408, 2.054, 2.906, RED, [[2.054, 0], [0, 2.054], [2.054, 2.906, 0.529, 2.572, 1.252, 2.895], [2.054, 0]]);
  poly(s, 9.469, 4.231, 2.901, 2.055, DARK, [[2.049, 0], [0, 2.055], [2.901, 2.055], [2.049, 0, 2.895, 1.258, 2.572, 0.535]]);
  poly(s, 9.419, 6.408, 2.049, 2.906, DARK, [[0, 0], [0, 2.906], [2.049, 2.054, 0.796, 2.895, 1.52, 2.572], [0, 0]]);
  oval(s, 7.849, 4.788, 3.068, W);
  oval(s, 8.032, 4.977, 2.695, MIST);
  oval(s, 8.194, 5.133, 2.377, RED);
  oval(s, 8.367, 5.316, 2.033, W);
  s.addShape('line', { x: 11.566, y: 4.647, w: 1.587, h: 0, line: { color: DARK, width: 1.5 } });
  s.addShape('line', { x: 11.709, y: 6.805, w: 1.444, h: 0, line: { color: RED, width: 1.5 } });
  s.addShape('line', { x: 10.529, y: 8.96, w: 2.625, h: 0, line: { color: DARK, width: 1.5 } });
  s.addShape('line', { x: 5.597, y: 4.647, w: 1.7, h: 0, line: { color: RED, width: 1.5 } });
  s.addShape('line', { x: 5.597, y: 6.805, w: 1.509, h: 0, line: { color: DARK, width: 1.5 } });
  s.addShape('line', { x: 5.597, y: 8.96, w: 2.604, h: 0, line: { color: RED, width: 1.5 } });
  tx(s, 13.858, 4.414, 3.67, 0.871, L3, BODYR);
  tx(s, 13.858, 4.01, 3.67, 0.404, 'Your Title', LABEL);
  tx(s, 1.221, 4.414, 3.67, 0.872, L3, BODYR);
  tx(s, 1.221, 4.01, 3.67, 0.404, 'Your Title', st(LABEL, { al: 'r' }));
  tx(s, 13.858, 6.572, 3.67, 0.871, L3, BODYR);
  tx(s, 13.858, 6.168, 3.67, 0.404, 'Your Title', LABEL);
  tx(s, 1.221, 6.572, 3.67, 0.871, L3, BODYR);
  tx(s, 1.221, 6.168, 3.67, 0.404, 'Your Title', st(LABEL, { al: 'r' }));
  tx(s, 13.858, 8.727, 3.67, 0.871, L3, BODYR);
  tx(s, 13.858, 8.323, 3.67, 0.404, 'Your Title', LABEL);
  tx(s, 1.221, 8.727, 3.67, 0.871, L3, BODYR);
  tx(s, 1.221, 8.323, 3.67, 0.404, 'Your Title', st(LABEL, { al: 'r' }));
  tx(s, 4.208, 2.021, 10.333, 1.111, 'Our |Infographics|', { sz: 60, f: SORA, al: 'c' });
  tx(s, 4.208, 1.652, 10.333, 0.404, 'Infographics', EYEBROWC);
  splat(s, 8.861, 5.833, 1.043);
  [['bulb', 6.925, 5.084], ['target', 6.926, 6.77], ['gear', 8.297, 8.106],
   ['eye', 9.843, 8.037], ['stack', 11.321, 6.927], ['board', 11.19, 5.113]
  ].forEach(function (g) { glyph(s, g[0], g[1], g[2], 0.72, W); });
}

function slide35(p) {
  const s = p.addSlide();
  rail(s, '35', 'L', DARK);
  box(s, 3.664, 3.795, 5.872, 1.676, MIST);
  box(s, 3.021, 3.536, 1.292, 1.292, RED);
  tx(s, 3.021, 3.536, 1.292, 1.292, '01', { sz: 24, f: SORA, c: W, al: 'c', v: 'm' });
  s.addShape('rtTriangle', { x: 3.021, y: 4.828, w: 0.643, h: 0.643, fill: { color: DARK }, rotate: 180 });
  box(s, 3.664, 6.023, 5.872, 1.676, MIST);
  box(s, 3.021, 5.764, 1.292, 1.292, DARK);
  tx(s, 3.021, 5.764, 1.292, 1.292, '02', { sz: 24, f: SORA, c: W, al: 'c', v: 'm' });
  s.addShape('rtTriangle', { x: 3.021, y: 7.056, w: 0.643, h: 0.643, fill: { color: RED }, rotate: 180 });
  box(s, 3.664, 8.25, 5.872, 1.676, MIST);
  box(s, 3.021, 7.992, 1.292, 1.292, RED);
  tx(s, 3.021, 7.992, 1.292, 1.292, '03', { sz: 24, f: SORA, c: W, al: 'c', v: 'm' });
  s.addShape('rtTriangle', { x: 3.021, y: 9.284, w: 0.643, h: 0.643, fill: { color: DARK }, rotate: 180 });
  box(s, 12.357, 3.795, 5.872, 1.676, MIST);
  box(s, 11.714, 3.536, 1.292, 1.292, DARK);
  tx(s, 11.714, 3.536, 1.292, 1.292, '04', { sz: 24, f: SORA, c: W, al: 'c', v: 'm' });
  s.addShape('rtTriangle', { x: 11.714, y: 4.828, w: 0.643, h: 0.643, fill: { color: RED }, rotate: 180 });
  box(s, 12.357, 6.023, 5.872, 1.676, MIST);
  box(s, 11.714, 5.764, 1.292, 1.292, RED);
  tx(s, 11.714, 5.764, 1.292, 1.292, '05', { sz: 24, f: SORA, c: W, al: 'c', v: 'm' });
  s.addShape('rtTriangle', { x: 11.714, y: 7.056, w: 0.643, h: 0.643, fill: { color: DARK }, rotate: 180 });
  box(s, 12.357, 8.25, 5.872, 1.676, MIST);
  box(s, 11.714, 7.992, 1.292, 1.292, DARK);
  tx(s, 11.714, 7.992, 1.292, 1.292, '06', { sz: 24, f: SORA, c: W, al: 'c', v: 'm' });
  s.addShape('rtTriangle', { x: 11.714, y: 9.284, w: 0.643, h: 0.643, fill: { color: RED }, rotate: 180 });
  tx(s, 4.758, 3.996, 4.333, 1.276, L4, BODYL);
  tx(s, 4.758, 6.224, 4.333, 1.275, L4, BODYL);
  tx(s, 4.758, 8.451, 4.333, 1.275, L4, BODYL);
  tx(s, 13.451, 3.996, 4.333, 1.275, L4, BODYL);
  tx(s, 13.451, 6.224, 4.333, 1.275, L4, BODYL);
  tx(s, 13.451, 8.451, 4.333, 1.275, L4, BODYL);
  tx(s, 4.177, 1.734, 12.896, 1.111, 'Our |Infographics|', { sz: 60, f: SORA, al: 'c' });
  tx(s, 4.177, 1.323, 12.896, 0.404, 'Infographics', EYEBROWC);
}

function slide36(p) {
  const s = p.addSlide();
  rail(s, '36', 'R', DARK);
  rbox(s, 0.94, 3.735, 5.275, 6.317, 0.228, RED);
  rbox(s, 12.535, 3.735, 5.275, 6.317, 0.228, RED);
  s.addShape('parallelogram', { x: 1.265, y: 3.504, w: 4.694, h: 0.896, fill: { color: DARK } });
  s.addShape('parallelogram', { x: 12.859, y: 3.504, w: 4.694, h: 0.896, fill: { color: DARK } });
  tx(s, 3.896, 1.597, 10.958, 0.95, 'Restaurant |Pricing |Table', HEADC);
  tx(s, 3.896, 1.197, 10.958, 0.404, 'Pricing Table', EYEBROWC);
  rbox(s, 6.737, 3.735, 5.275, 6.317, 0.228, DARK);
  tx(s, 6.89, 4.963, 4.97, 1.212, '$50 |/month|', { sz: 66, f: DMM, c: W, a: W, az: 20, al: 'c' });
  tx(s, 6.89, 6.518, 4.97, 0.468, 'Not Facing Crowded Dining Places', st(WBODYC, { sp: 1, bu: true }));
  tx(s, 6.89, 7.29, 4.97, 0.468, 'Get a Beautiful View', st(WBODYC, { sp: 1, bu: true }));
  tx(s, 6.89, 8.063, 4.97, 0.468, 'Get a Private Table', st(WBODYC, { sp: 1, bu: true }));
  poly(s, 6.908, 3.506, 4.934, 0.896, RED, [[0, 0], [4.934, 0], [4.934, 0.016], [4.611, 0.447], [4.934, 0.879], [4.934, 0.896], [0, 0.896], [0, 0.886], [0.328, 0.447], [0, 0.008]]);
  tx(s, 7.503, 3.752, 3.744, 0.404, 'Normal ', WNAME);
  tx(s, 1.093, 4.963, 4.97, 1.212, '$30 |/month|', { sz: 66, f: DMM, c: W, a: W, az: 20, al: 'c' });
  tx(s, 1.093, 6.518, 4.97, 0.468, 'Not Facing Crowded Dining Places', st(WBODYC, { sp: 1, bu: true }));
  tx(s, 1.093, 7.29, 4.97, 0.468, 'Get a Beautiful View', st(WBODYC, { sp: 1, bu: true }));
  tx(s, 1.093, 8.063, 4.97, 0.468, 'Get a Private Table', st(WBODYC, { sp: 1, bu: true }));
  tx(s, 1.706, 3.752, 3.744, 0.404, 'Basic', WNAME);
  tx(s, 12.687, 4.963, 4.97, 1.212, '$90 |/month|', { sz: 66, f: DMM, c: W, a: W, az: 20, al: 'c' });
  tx(s, 12.687, 6.518, 4.97, 0.468, 'Not Facing Crowded Dining Places', st(WBODYC, { sp: 1, bu: true }));
  tx(s, 12.687, 7.29, 4.97, 0.468, 'Get a Beautiful View', st(WBODYC, { sp: 1, bu: true }));
  tx(s, 12.687, 8.063, 4.97, 0.468, 'Get a Private Table', st(WBODYC, { sp: 1, bu: true }));
  tx(s, 13.3, 3.752, 3.744, 0.404, 'Premium ', WNAME);
  tx(s, 6.89, 8.835, 4.97, 0.468, 'Free Dessert', st(WBODYC, { sp: 1, bu: true }));
  tx(s, 1.093, 8.835, 4.97, 0.468, 'Free Dessert', st(WBODYC, { sp: 1, bu: true }));
  tx(s, 12.687, 8.835, 4.97, 0.468, 'Free Dessert', st(WBODYC, { sp: 1, bu: true }));
}

function slide37(p) {
  const s = p.addSlide();
  rail(s, '37', 'L', DARK);
  rbox(s, 1.883, 0.899, 11.697, 9.258, 0.311, RED);
  poly(s, 12.016, 0, 7.984, 10.902, DARK, [[0, 0], [7.984, 0], [7.984, 10.902], [0.821, 10.902], [0, 9.992, 0.368, 10.902, 0, 10.495]]);
  tx(s, 2.762, 2.215, 8.29, 0.404, 'Contact Us', WTITLE);
  tx(s, 2.762, 2.624, 8.283, 1.96, 'Get In Touch With Meaty Company', { sz: 60, f: SORA, c: W, v: 'm', ls: 1, m: 0 });
  tx(s, 8.473, 5.806, 2.572, 0.872, 'Meaty Street 12,\nAnycity', { sz: 16, c: W, ls: 1.5, sp: 1 });
  tx(s, 8.473, 5.433, 2.572, 0.404, 'Address', { f: DMM, c: W, sp: 1 });
  tx(s, 3.695, 5.806, 3.144, 0.872, '+1234 456 678 890\n(+1234) 56789', { sz: 16, c: W, ls: 1.5, sp: 1 });
  tx(s, 3.695, 5.433, 3.144, 0.404, 'Phone', { f: DMM, c: W, sp: 1 });
  tx(s, 3.695, 7.967, 3.144, 0.872, 'mail@Meaty.com Info@Meaty.com ', { sz: 16, c: W, ls: 1.5, sp: 1 });
  tx(s, 3.695, 7.553, 3.144, 0.404, 'Email', { f: DMM, c: W, sp: 1 });
  tx(s, 8.473, 7.968, 2.572, 0.872, 'Monday - Saturday\n08.00 – 20.00', { sz: 16, c: W, ls: 1.5, sp: 1 });
  tx(s, 8.473, 7.553, 2.572, 0.404, 'Office', { f: DMM, c: W, sp: 1 });
  poly(s, 12.392, 0, 7.608, 10.526, PHOTO, [[0, 0], [7.608, 0], [7.608, 10.526], [0.821, 10.526], [0, 9.608, 0.368, 10.526, 0, 10.115]]);
  [['phone', 2.762, 5.443], ['pin', 7.518, 5.374], ['mail', 2.762, 7.553],
   ['medal', 7.518, 7.553]].forEach(function (g) { glyph(s, g[0], g[1], g[2], 0.714, W); });
}

function slide38(p) {
  const s = p.addSlide();
  rail(s, '40', 'R', DARK);
  wave(s, 0, 0, 5, 11.25, RED, true);
  tx(s, 7.775, 4.253, 10.079, 2.339, 'Thank| You|', { sz: 133, f: SORA });
  tx(s, 7.775, 6.593, 10.079, 0.404, 'Thank You For Your Time |and Attention|', { f: DMM, al: 'c' });
  oval(s, 0.954, 2.674, 6.502, DARK);
  oval(s, 1.245, 2.964, 5.921, PHOTO2);
}

function slide39(p) {
  const s = p.addSlide();
  s.background = { color: RED };
  rail(s, '39', 'L', W);
  iconGrid(s, [4.56, 5.62, 6.76, 7.83, 8.95, 10.03, 11.12, 12.22, 13.34, 14.413, 15.57, 16.66],
    [2.87, 3.958, 5.041, 6.15, 7.19, 8.32]);
}

function slide40(p) {
  const s = p.addSlide();
  s.background = { color: DARK };
  rail(s, '40', 'R', W);
  iconGrid(s, [3.25, 4.35, 5.45, 6.56, 7.69, 8.8, 9.93, 10.99, 12.1, 13.2, 14.3, 15.41],
    [2.83, 3.92, 5.093, 6.18, 7.36, 8.47]);
}

/* ----- deck ------------------------------------------------------------- */
function build() {
  const p = new PptxGenJS();
  p.defineLayout({ name: 'MEATY', width: 20, height: 11.25 });
  p.layout = 'MEATY';
  p.author = 'Meaty';
  p.title = 'Meaty - Food Presentation Template';

  const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
    slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
    slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24,
    slide25, slide26, slide27, slide28, slide29, slide30, slide31, slide32,
    slide33, slide34, slide35, slide36, slide37, slide38, slide39, slide40];
  SLIDES.forEach(function (fn) { fn(p); });

  return p.writeFile({ fileName: path.join(__dirname, '0aaa1958-023c-4c22-af0a-76bd566f0735_grok_final.pptx') });
}

build().then(function (f) { console.log('wrote', f); })
  .catch(function (e) { console.error(e); process.exit(1); });
