#!/usr/bin/env node
/**
 * CarListic — Cargo & Logistic deck (36 slides, 13.333in x 7.5in widescreen)
 * Standalone pptxgenjs recreation.  Run: node <this file>
 *
 * Raster photographs from the original deck are replaced by flat placeholder
 * rectangles (`ph`) and its SVG pictogram set by simple vector `icon` glyphs.
 */
'use strict';
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette ---
const OR  = 'FF5E16';   // brand orange
const OR2 = 'ED7D31';   // secondary orange (logo bars)
const WH  = 'FFFFFF';
const BK  = '000000';
const GY  = '808080';   // body grey
const LG  = 'BFBFBF';   // light grey
const WS  = 'F2F2F2';   // whitesmoke rule
const SL  = '333F50';   // slate panel
const DG  = '595959';   // dark grey disc
const NB  = '262626';   // quote-mark grey
const MG  = 'A6A6A6';
const LN  = 'D9D9D9';
const NV  = '242F44';
const FB  = '7093D2';   // facebook chip
const TW  = '00B0F0';   // twitter chip
const LI  = '19C5D7';   // linkedin chip
const YL  = 'FFFF00';   // star yellow

// ------------------------------------------------------------------ fonts ---
const PS = 'Poppins SemiBold';
const PM = 'Poppins Medium';
const PP = 'Poppins';
const LT = 'Lato';
const WK = 'Work Sans';

// ------------------------------------------------- recurring body copy ------
const VOC = 'Vocibus mediocrem ex vis. Novum errem pertinaci ethereum persequeris ea quote, '
  + 'any nam mnesarch voluptatibus sileuni definitionem, has in agam meliore delicata. '
  + 'Sed erosa aliquidy Inani convenire delenit. Vim eu ponderum incorrupte, graecie Dolore '
  + 'nostrum has in agam meliore delicata. Sed eros aliquid Inai convenire delenit. Vim eu '
  + 'ponderum maldives incorrupte est. Dolore nostrum lorem ipsum dolor sit amet';
const SED = 'Sed ut perspiciatis unde omnis iste natus voluptatea santium doloremque laudantium, '
  + 'totam aperiam, eaque ipsaab ilnventore veritatis et quasrchitecto beatae vitae dicta .';
const LIP = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nulla rhoncus non sapien '
  + 'a mollis. Phasellus pretium egestas posuere. Pellentesque porttitor fringilla justo in '
  + 'tincidunt. Cras metus metus, congue eget imperdiet eu, auctor vitae ';
const PEL = 'Pellentesque porttitor fringilla justo in memo tincidunt.';
const LID = 'Lorem Ipsum Dumet';
const WWW = 'www.carlistic.com';

// ------------------------------------------------------------- primitives ---
const AL = { l: 'left', c: 'center', r: 'right', j: 'justify' };

/** Preset/custom-geometry shape. o: {f fill, a transparency, l line, lw, ld, rot, fh, fv, r radius, pts, sd} */
function sh(s, kind, x, y, w, h, o) {
  o = o || {};
  const p = { x, y, w, h };
  if (o.f) p.fill = o.a ? { color: o.f, transparency: o.a } : { color: o.f };
  if (o.l) p.line = { color: o.l, width: o.lw || 1, dashType: o.ld || 'solid' };
  if (o.rot) p.rotate = o.rot;
  if (o.fh) p.flipH = true;
  if (o.fv) p.flipV = true;
  if (o.r !== undefined) p.rectRadius = o.r;
  if (o.pts) p.points = o.pts.map(([px, py]) => ({ x: px * w, y: py * h }));
  if (o.sd) p.shadow = o.sd;
  s.addShape(kind, p);
}

/** Text block. runs = [[text, {sz,f,c,cs,b,cap,br}], ...]; o = block-level defaults. */
function tx(s, x, y, w, h, runs, o) {
  o = o || {};
  const base = {
    fontSize: o.sz || 18, fontFace: o.f || PP, color: o.c || BK,
    charSpacing: o.cs || 0, bold: !!o.b,
    align: AL[o.al] || 'left', valign: o.va === 'm' ? 'middle' : o.va === 'b' ? 'bottom' : 'top',
  };
  const p = Object.assign({ x, y, w, h, margin: 0, wrap: !o.nw, isTextBox: true }, base);
  if (o.ls) p.lineSpacingMultiple = o.ls;
  if (o.lp) p.lineSpacing = o.lp;
  if (o.vt) p.vert = o.vt;
  if (o.rot) p.rotate = o.rot;
  if (o.bg) p.fill = { color: o.bg };
  if (o.ol) { base.outline = { color: o.ol, size: o.olw || 3 }; base.transparency = 100; }
  s.addText(runs.map(([t, r]) => {
    r = r || {};
    return {
      text: t, options: {
        fontSize: r.sz || base.fontSize, fontFace: r.f || base.fontFace,
        color: r.c || base.color, charSpacing: r.cs !== undefined ? r.cs : base.charSpacing,
        bold: r.b !== undefined ? !!r.b : base.bold,
        align: base.align, breakLine: !!r.br,
        outline: base.outline, transparency: base.transparency,
      },
    };
  }), p);
}

/** Straight connector. */
function ln(s, x, y, w, h, color, o) {
  o = o || {};
  const p = { x, y, w, h, line: { color, width: o.lw || 1, dashType: o.ld || 'solid' } };
  if (o.rot) p.rotate = o.rot;
  if (o.fh) p.flipH = true;
  s.addShape('line', p);
}

/** Blend hex colour `c` at `t` (0..1) opacity over backdrop `bg`. */
function mix(bg, c, t) {
  let o = '';
  for (let i = 0; i < 6; i += 2) {
    const a = parseInt(bg.substr(i, 2), 16), b = parseInt(c.substr(i, 2), 16);
    o += Math.round(a + (b - a) * t).toString(16).padStart(2, '0');
  }
  return o.toUpperCase();
}

/** Two-stop linear gradient, approximated by overlapping opaque bands pre-blended
 *  against `bg` (pptxgenjs shape fills are solid-only).  a/b = [color, transparency, stop%]. */
function grad(s, x, y, w, h, dir, a, b, bg) {
  const N = 48, pad = 0.02;   // bands overlap slightly so no seams show when rendered
  const span = (b[2] - a[2]) || 100;
  for (let i = 0; i < N; i++) {
    const k = Math.min(1, Math.max(0, ((i + 0.5) / N * 100 - a[2]) / span));
    const c = mix(bg, k < 0.5 ? a[0] : b[0], 1 - (a[1] + (b[1] - a[1]) * k) / 100);
    if (dir === 'v') sh(s, 'rect', x, y + h * i / N, w, h / N + pad, { f: c });
    else sh(s, 'rect', x + w * i / N, y, w / N + pad, h, { f: c });
  }
}

/** Placeholder standing in for a photograph in the source deck. */
function ph(s, x, y, w, h) {
  sh(s, 'rect', x, y, w, h, { f: LG, a: 45 });
  if (w > 1.2 && h > 0.5) {
    tx(s, x, y + h / 2 - 0.16, w, 0.32, [['[image]']],
       { sz: 11, f: LT, c: GY, al: 'c', va: 'm' });
  }
}

// --------------------------------------------------------- icon pictograms --
// Flat vector stand-ins for the deck's SVG icon set; `k` selects the glyph.
const ICONS = {
  truck: (s, x, y, u, c) => {
    sh(s, 'rect', x + .04 * u, y + .30 * u, .54 * u, .32 * u, { f: c });
    sh(s, 'rect', x + .58 * u, y + .42 * u, .30 * u, .20 * u, { f: c });
    sh(s, 'rect', x + .04 * u, y + .62 * u, .84 * u, .05 * u, { f: c });
    sh(s, 'ellipse', x + .13 * u, y + .63 * u, .17 * u, .17 * u, { f: c });
    sh(s, 'ellipse', x + .64 * u, y + .63 * u, .17 * u, .17 * u, { f: c });
  },
  plane: (s, x, y, u, c) => {
    sh(s, 'rect', x + .44 * u, y + .10 * u, .12 * u, .62 * u, { f: c, r: .05 * u });
    sh(s, 'triangle', x + .04 * u, y + .30 * u, .92 * u, .26 * u, { f: c });
    sh(s, 'triangle', x + .30 * u, y + .64 * u, .40 * u, .20 * u, { f: c });
  },
  ship: (s, x, y, u, c) => {
    sh(s, 'trapezoid', x + .02 * u, y + .56 * u, .96 * u, .24 * u, { f: c, rot: 180 });
    sh(s, 'rect', x + .12 * u, y + .38 * u, .70 * u, .16 * u, { f: c });
    sh(s, 'rect', x + .22 * u, y + .24 * u, .44 * u, .12 * u, { f: c });
    sh(s, 'rect', x + .34 * u, y + .16 * u, .16 * u, .06 * u, { f: c });
  },
  train: (s, x, y, u, c) => {
    sh(s, 'roundRect', x + .20 * u, y + .14 * u, .60 * u, .68 * u, { f: c, r: .12 * u });
    sh(s, 'rect', x + .47 * u, y + .02 * u, .06 * u, .13 * u, { f: c });
    sh(s, 'rect', x + .30 * u, y + .06 * u, .40 * u, .05 * u, { f: c });
    sh(s, 'roundRect', x + .28 * u, y + .24 * u, .44 * u, .26 * u, { f: WH, r: .04 * u });
    sh(s, 'rect', x + .26 * u, y + .82 * u, .06 * u, .12 * u, { f: c });
    sh(s, 'rect', x + .68 * u, y + .82 * u, .06 * u, .12 * u, { f: c });
  },
  check: (s, x, y, u, c) => {
    sh(s, 'rect', x + .06 * u, y + .46 * u, .34 * u, .13 * u, { f: c, rot: 45 });
    sh(s, 'rect', x + .26 * u, y + .18 * u, .66 * u, .13 * u, { f: c, rot: -38 });
  },
  users: (s, x, y, u, c) => {
    [[.10, .46], [.38, .34], [.66, .46]].forEach(([dx, dy]) => {
      sh(s, 'ellipse', x + dx * u, y + dy * u, .24 * u, .24 * u, { f: c });
      sh(s, 'roundRect', x + (dx - .04) * u, y + (dy + .24) * u, .32 * u, .22 * u, { f: c, r: .08 * u });
    });
  },
  direction: (s, x, y, u, c) => {
    sh(s, 'triangle', x + .10 * u, y + .10 * u, .80 * u, .80 * u, { f: c, rot: 135 });
  },
  box: (s, x, y, u, c) => {
    sh(s, 'cube', x + .10 * u, y + .18 * u, .76 * u, .68 * u, { f: c, l: c === WH ? BK : WH, lw: 1 });
  },
  gem: (s, x, y, u, c) => {
    sh(s, 'diamond', x + .12 * u, y + .30 * u, .76 * u, .62 * u, { f: c });
    sh(s, 'trapezoid', x + .12 * u, y + .10 * u, .76 * u, .24 * u, { f: c });
  },
  marker: (s, x, y, u, c) => {
    sh(s, 'ellipse', x + .20 * u, y + .06 * u, .60 * u, .60 * u, { f: c });
    sh(s, 'triangle', x + .30 * u, y + .48 * u, .40 * u, .46 * u, { f: c, rot: 180 });
    sh(s, 'ellipse', x + .39 * u, y + .24 * u, .22 * u, .22 * u, { f: WH });
  },
  world: (s, x, y, u, c) => {
    sh(s, 'ellipse', x + .08 * u, y + .08 * u, .84 * u, .84 * u, { l: c, lw: 2 });
    sh(s, 'ellipse', x + .32 * u, y + .08 * u, .36 * u, .84 * u, { l: c, lw: 1.5 });
    ln(s, x + .08 * u, y + .50 * u, .84 * u, 0, c, { lw: 1.5 });
  },
  phone: (s, x, y, u, c) => {
    sh(s, 'roundRect', x + .10 * u, y + .50 * u, .48 * u, .24 * u, { f: c, rot: -40, r: .07 * u });
    sh(s, 'blockArc', x + .34 * u, y + .06 * u, .60 * u, .60 * u, { f: c, rot: -45 });
  },
  envelope: (s, x, y, u, c) => {
    sh(s, 'rect', x + .26 * u, y + .10 * u, .48 * u, .34 * u, { f: c });
    sh(s, 'rect', x + .32 * u, y + .16 * u, .36 * u, .04 * u, { f: WH });
    sh(s, 'rect', x + .32 * u, y + .24 * u, .36 * u, .04 * u, { f: WH });
    sh(s, 'trapezoid', x + .04 * u, y + .38 * u, .92 * u, .46 * u, { f: c, rot: 180 });
  },
  cycle: (s, x, y, u, c) => {
    ln(s, x + .22 * u, y + .30 * u, .28 * u, .34 * u, c, { lw: 3.5, ld: 'dash' });
    ln(s, x + .50 * u, y + .30 * u, .28 * u, .34 * u, c, { lw: 3.5, ld: 'dash', fh: 1 });
    ln(s, x + .30 * u, y + .74 * u, .40 * u, 0, c, { lw: 3.5, ld: 'dash' });
    [[.34, .04], [.04, .58], [.64, .58]].forEach(([dx, dy]) =>
      sh(s, 'ellipse', x + dx * u, y + dy * u, .32 * u, .32 * u, { f: c }));
  },
  play: (s, x, y, u, c) => {
    sh(s, 'triangle', x + .16 * u, y + .10 * u, .68 * u, .80 * u, { f: c, rot: 90 });
  },
  star: (s, x, y, u, c) => {
    sh(s, 'star5', x, y, u, u, { f: c });
  },
};

/** Draw an icon glyph of size u (inches) at x,y in colour c. */
function icon(s, x, y, u, c, k) {
  (ICONS[k] || ICONS.truck)(s, x, y, u, c);
}


// The four oversized SWOT initials are outlined freeform letters in the source deck.
// Each entry is a list of sub-paths; a 2-number point is a line-to, 6 numbers a cubic
// Bezier (c1x,c1y,c2x,c2y,x,y).  Coordinates are fractions of the shape box.
const LETTERS = {
  S: [[[0.464,0],[0.611,0,0.729,0.03,0.818,0.09],[0.907,0.15,0.963,0.229,0.985,0.329],[0.696,0.329],[0.684,0.282,0.658,0.245,0.617,0.217],[0.576,0.188,0.523,0.174,0.457,0.174],[0.405,0.174,0.364,0.184,0.335,0.204],[0.306,0.224,0.291,0.251,0.291,0.285],[0.291,0.312,0.303,0.334,0.326,0.35],[0.349,0.367,0.379,0.38,0.414,0.389],[0.449,0.398,0.499,0.41,0.565,0.423],[0.656,0.439,0.731,0.456,0.789,0.474],[0.846,0.493,0.896,0.521,0.938,0.561],[0.979,0.6,1,0.653,1,0.72],[1,0.803,0.96,0.871,0.879,0.923],[0.798,0.974,0.689,1,0.551,1],[0.394,1,0.268,0.972,0.173,0.915],[0.078,0.858,0.02,0.776,0,0.671],[0.295,0.671],[0.303,0.719,0.329,0.757,0.374,0.784],[0.419,0.812,0.478,0.826,0.551,0.826],[0.604,0.826,0.644,0.815,0.672,0.795],[0.699,0.774,0.713,0.748,0.713,0.715],[0.713,0.687,0.701,0.664,0.677,0.646],[0.653,0.629,0.623,0.615,0.588,0.606],[0.552,0.596,0.501,0.585,0.435,0.572],[0.345,0.556,0.272,0.539,0.216,0.522],[0.159,0.505,0.111,0.478,0.071,0.44],[0.031,0.403,0.011,0.351,0.011,0.285],[0.011,0.201,0.051,0.132,0.131,0.079],[0.212,0.026,0.322,0,0.464,0]]],
  W: [[[0,0],[0.166,0],[0.272,0.773],[0.416,0],[0.592,0],[0.738,0.773],[0.845,0],[1,0],[0.825,1],[0.641,1],[0.501,0.273],[0.358,1],[0.174,1]]],
  O: [[[0.5,0.18],[0.441,0.18,0.39,0.193,0.346,0.219],[0.301,0.244,0.267,0.281,0.242,0.33],[0.216,0.378,0.204,0.435,0.204,0.499],[0.204,0.564,0.216,0.62,0.242,0.669],[0.267,0.717,0.301,0.754,0.346,0.78],[0.39,0.806,0.441,0.819,0.5,0.819],[0.559,0.819,0.61,0.806,0.654,0.78],[0.699,0.754,0.733,0.717,0.758,0.669],[0.784,0.62,0.796,0.564,0.796,0.499],[0.796,0.435,0.784,0.378,0.758,0.33],[0.733,0.281,0.699,0.244,0.654,0.219],[0.61,0.193,0.559,0.18,0.5,0.18]],[[0.5,0],[0.593,0,0.678,0.021,0.754,0.064],[0.83,0.106,0.89,0.165,0.934,0.241],[0.978,0.317,1,0.403,1,0.499],[1,0.595,0.978,0.681,0.934,0.758],[0.89,0.834,0.83,0.894,0.754,0.936],[0.678,0.979,0.593,1,0.5,1],[0.407,1,0.322,0.979,0.246,0.936],[0.17,0.894,0.11,0.834,0.066,0.758],[0.022,0.681,0,0.595,0,0.499],[0,0.403,0.022,0.317,0.066,0.241],[0.11,0.165,0.17,0.106,0.246,0.064],[0.322,0.021,0.407,0,0.5,0]]],
  T: [[[0,0],[1,0],[1,0.159],[0.632,0.159],[0.632,1],[0.366,1],[0.366,0.159],[0,0.159]]],
};

/** Draw one of the outlined SWOT initials. */
function letter(s, k, x, y, w, h, c, lw) {
  const pts = [];
  LETTERS[k].forEach(sub => {
    sub.forEach((p, i) => {
      if (i === 0) pts.push({ moveTo: true, x: p[0] * w, y: p[1] * h });
      else if (p.length === 2) pts.push({ x: p[0] * w, y: p[1] * h });
      else pts.push({ x: p[4] * w, y: p[5] * h, curve: { type: 'cubic',
        x1: p[0] * w, y1: p[1] * h, x2: p[2] * w, y2: p[3] * h } });
    });
    pts.push({ close: true });
  });
  s.addShape('custGeom', { x, y, w, h, points: pts, fill: { type: 'none' },
                           line: { color: c, width: lw } });
}

// ---------------------------------------------------- repeated decorations --
/** 6x6 triangular dot matrix (21 dots). `right` mirrors the triangle. */
function dots(s, x, y, dx, dy, d, c, right) {
  for (let r = 0; r < 6; r++) {
    for (let i = 0; i < 6 - r; i++) {
      const cx = right ? x + (5 - i) * dx : x + i * dx;
      sh(s, 'ellipse', cx, y + r * dy, d, d * 0.9, { f: c });
    }
  }
}

/** Four outlined social circles (facebook / twitter / google+ / youtube). */
function socRing(s, x, y, d, step, c, horiz) {
  ['f', 't', 'g+', 'yt'].forEach((g, i) => {
    const cx = horiz ? x + i * step : x;
    const cy = horiz ? y : y + i * step;
    sh(s, 'ellipse', cx, cy, d, d, { l: c, lw: 1.25 });
    tx(s, cx, cy, d, d, [[g]], { sz: d * 26, f: PP, c, al: 'c', va: 'm' });
  });
}

/** Three small solid social glyphs, no ring. */
function socGlyph(s, x, y, c) {
  ['t', 'f', 'in'].forEach((g, i) =>
    tx(s, x + i * 0.38, y - 0.06, 0.3, 0.3, [[g]], { sz: 11, f: PS, c, al: 'c', va: 'm' }));
}

/** Three coloured social discs (facebook / twitter / linkedin). */
function socDisc(s, x, y) {
  [[FB, 'f'], [TW, 't'], [LI, 'in']].forEach(([c, g], i) => {
    sh(s, 'ellipse', x + i * 0.38, y, 0.28, 0.28, { f: c });
    tx(s, x + i * 0.38, y, 0.28, 0.28, [[g]], { sz: 9, f: PS, c: WH, al: 'c', va: 'm' });
  });
}

/** White disc with a centred icon glyph. */
function iconDisc(s, x, y, d, ring, c, k) {
  sh(s, 'ellipse', x, y, d, d, { f: ring });
  icon(s, x + d * 0.26, y + d * 0.26, d * 0.48, c, k);
}


// ------------------------------------------------------------ slides ---

/** Slide 1 — Cover — CarListic (layout: Custom Layout) */
function slide01(s) {
  sh(s,'rect',4.45,1.98,7.35,3.62,{l:OR,lw:2.75});
  sh(s,'rect',4.64,2.15,7.06,3.29,{f:WH});
  tx(s,5.24,2.66,5.98,1.68,[['Car'],['Listic',{c:BK}]],{sz:96,f:PS,c:OR,al:'c',nw:1});
  tx(s,5.51,4.19,5.4,0.493,[['Cargo & logistic Presentation']],{sz:24,f:PP,c:GY,cs:1.2,al:'c',nw:1});
  sh(s,'rtTriangle',10.97,1.89,1.15,1.15,{f:OR,rot:180});
  ln(s,9.54,5.88,0,2.03,WH,{lw:1.25,rot:90});
  tx(s,6.77,5.69,0.488,2.4,[['www.'],['carlistic',{cs:1.2}],['.com']],{sz:17,f:PP,c:WH,al:'r',nw:1,vt:'vert270',rot:90});
  socRing(s,10.8,6.71,0.37,0.56,WH,true);
  sh(s,'rect',12.37,0.69,0.483,0.05,{f:OR2});
  sh(s,'rect',12.26,0.58,0.59,0.05,{f:OR2});
  sh(s,'rect',12.11,0.467,0.74,0.05,{f:OR2});
}

/** Slide 2 — About us (layout: 3_Custom Layout) */
function slide02(s) {
  sh(s,'rect',0,3.23,13.33,4.27,{f:OR});
  sh(s,'rect',0.95,1.1,5.29,2.58,{f:BK,a:31});
  tx(s,1.3,1.52,3.83,1.85,[['Welcome To '],['Carlistic',{sz:32,f:PP,c:OR,b:1}],[' Cargo And Logistic']],{sz:28,f:PM,c:WH,ls:1.2});
  tx(s,2.36,4.71,4.87,2.18,[[VOC]],{sz:12,f:LT,c:WH,ls:1.5});
  tx(s,9.37,4.94,3.19,0.303,[['LOREM IPSUM DUMET']],{sz:12,f:PS,c:WH,cs:0.5});
  tx(s,9.37,5.25,3.19,1.04,[['andiz oblivion dfinitionehas shine in service  beauty alins  maldies inru superior features women snoruma note oneom ipsum dolor amet consecteur everdeen ']],{sz:11,f:LT,c:WH,ls:1.3});
  sh(s,'rect',0.71,0.326,11.92,4.13,{l:OR,lw:2.75});
  sh(s,'rtTriangle',11.37,0.438,1.15,1.15,{f:OR,rot:180});
  tx(s,1.01,4.51,0.67,2.58,[['ABOUT US']],{sz:28,f:PS,c:WH,cs:1,al:'c',vt:'vert',rot:180});
  ln(s,1.32,6.39,0.75,0,WH,{lw:2.25,rot:270});
  ln(s,1.32,5.64,0.75,0,BK,{lw:2.25,rot:270});
  socGlyph(s,9.44,6.57,WH);
}

/** Slide 3 — About Carlistic (layout: Custom Layout) */
function slide03(s) {
  grad(s,0,0,13.33,7.5,'h',[BK,20,0],[BK,79,100],WH);
  tx(s,1.36,2.42,5.68,1.58,[[VOC.slice(0,327)]],{sz:12,f:LT,c:WH,ls:1.5});
  tx(s,1.36,1.1,5.09,1.2,[['About ',{br:1}],['Carlistic',{f:PP,b:1}],['  ']],{sz:28,f:PM,c:WH,cs:1.2,ls:1.2});
  ln(s,0.79,2.42,0.75,0,WH,{lw:2.25,rot:270});
  ln(s,0.79,1.67,0.75,0,OR,{lw:2.25,rot:270});
  icon(s,1.48,4.23,0.74,OR,'truck');
  tx(s,1.48,5.22,1.96,0.86,[[PEL]],{sz:10.5,f:LT,c:WH,ls:1.5});
  tx(s,1.48,4.9,2.21,0.375,[[LID]],{sz:12,f:PM,c:WH,cs:0.2,b:1,ls:1.5});
  icon(s,4.09,4.23,0.74,OR,'plane');
  tx(s,4.09,5.22,1.96,0.86,[[PEL]],{sz:10.5,f:LT,c:WH,ls:1.5});
  tx(s,4.09,4.9,2.21,0.375,[[LID]],{sz:12,f:PM,c:WH,cs:0.2,b:1,ls:1.5});
  icon(s,6.7,4.23,0.74,OR,'ship');
  tx(s,6.7,5.22,1.96,0.86,[[PEL]],{sz:10.5,f:LT,c:WH,ls:1.5});
  tx(s,6.7,4.9,2.21,0.375,[[LID]],{sz:12,f:PM,c:WH,cs:0.2,b:1,ls:1.5});
  ln(s,9.54,5.88,0,2.03,WH,{lw:1.25,rot:90});
  tx(s,6.77,5.76,0.488,2.25,[[WWW]],{sz:17,f:PP,c:WH,al:'r',nw:1,vt:'vert270',rot:90});
  socRing(s,10.8,6.71,0.37,0.56,WH,true);
}

/** Slide 4 — Vision & Mission (layout: 7_Custom Layout) */
function slide04(s) {
  sh(s,'rect',0.69,2.19,3.15,4.18,{f:OR});
  sh(s,'rect',6,2.03,6.72,2.4,{f:OR,a:48});
  tx(s,0.7,2.51,2.85,2.97,[['Our',{br:1}],['Vision and Mision']],{sz:36,f:PS,c:WH,al:'c',ls:1.2});
  tx(s,6.67,5.49,5.68,1.58,[[VOC.slice(0,327)]],{sz:12,f:LT,c:GY,ls:1.5});
  tx(s,6.6,2.25,1.8,0.447,[['Our Vision']],{f:PS,c:WH,ls:1.2});
  tx(s,6.67,5.04,1.8,0.447,[['Our mission']],{f:PS,ls:1.2});
  sh(s,'rect',9.11,0.351,2.95,0.67,{f:WH});
  tx(s,10.34,-0.56,0.488,2.54,[[WWW]],{sz:17,f:PP,cs:1.2,al:'r',nw:1,vt:'vert270',rot:90});
  sh(s,'rtTriangle',11.64,0.6,0.57,0.57,{f:OR,fh:1});
  sh(s,'rtTriangle',8.96,0.198,0.57,0.57,{f:OR,rot:90});
  tx(s,6.67,2.74,5.68,1.58,[[VOC.slice(0,327)]],{sz:12,f:LT,c:WH,ls:1.5});
  sh(s,'rtTriangle',2.83,5.54,1.15,1.15,{f:WH,rot:270});
  sh(s,'rtTriangle',2.97,5.55,1.15,1.15,{f:OR,rot:270});
  sh(s,'rect',0.54,2.03,3.44,4.5,{l:OR,lw:2.25});
}

/** Slide 5 — A Different Kind Of Company (layout: 4_Custom Layout) */
function slide05(s) {
  sh(s,'rtTriangle',0,0,2.76,2.76,{f:OR,rot:180,fh:1});
  sh(s,'rect',0.38,0.337,12.25,6.89,{l:OR,lw:2.75});
  tx(s,1.6,1.27,5.04,1.77,[['A Different Kind Of Company.'],[' A Different Kind Of Cargo.',{c:OR}]],{sz:28,f:PS,ls:1.2});
  tx(s,1.32,3.69,4.53,1.88,[[VOC.slice(0,303)+'.']],{sz:12,f:LT,c:GY,ls:1.5});
  socRing(s,1.67,6.31,0.308,0.468,BK,true);
  dots(s,0.52,0.47,0.25,0.22,0.066,WH);
  ln(s,2.54,3.24,0.75,0,OR,{lw:3,rot:180});
  ln(s,1.79,3.24,0.75,0,BK,{lw:3,rot:180});
  dots(s,9.84,0.55,0.25,0.22,0.066,OR);
  sh(s,'rect',9.9,6.27,2.95,0.67,{f:WH});
  tx(s,11.13,5.36,0.488,2.54,[[WWW]],{sz:17,f:PP,cs:1.2,al:'r',nw:1,vt:'vert270',rot:90});
  sh(s,'rtTriangle',12.42,6.52,0.57,0.57,{f:OR,fh:1});
  sh(s,'rtTriangle',9.75,6.12,0.57,0.57,{f:OR,rot:90});
  sh(s,'rtTriangle',5.27,6,1.15,1.15,{l:OR,lw:4.5,fh:1});
  sh(s,'rtTriangle',10.9,0.4,1.15,1.15,{f:OR,fh:1});
  sh(s,'custGeom',4.9,4.75,2.49,2.1,{f:OR,rot:180,fh:1,pts:[[0.87,1],[1,1],[0.12,0],[0,0],[0,0.01]]});
  sh(s,'custGeom',7.94,1.81,2.49,2.1,{f:OR,rot:180,fh:1,pts:[[0.87,1],[1,1],[0.12,0],[0,0],[0,0.01]]});
}

/** Slide 6 — Different Kind Of Cargo (layout: 5_Custom Layout) */
function slide06(s) {
  sh(s,'rect',0,3.23,13.33,4.27,{f:OR});
  tx(s,2.66,4.9,6.14,1.88,[[VOC]],{sz:12,f:LT,c:WH,ls:1.5});
  tx(s,7.66,0.56,5.05,1.77,[['A Different Kind Of Company.  A Different Kind Of '],['Cargo.',{c:OR}]],{sz:28,f:PS,ls:1.2});
  ln(s,8.51,2.52,2.64,0,OR,{lw:2.25});
  sh(s,'rtTriangle',11.69,0.337,1.15,1.15,{f:OR,rot:180});
  sh(s,'rtTriangle',0.73,5.55,1.15,1.15,{f:WH});
  dots(s,9.63,4,0.25,0.22,0.066,WH);
  sh(s,'rtTriangle',3.36,3.84,0.75,0.75,{f:OR,l:WH,lw:3,fh:1});
}

/** Slide 7 — Stats & checklist (layout: 6_Custom Layout) */
function slide07(s) {
  sh(s,'rect',0,3.75,6.67,3.75,{f:SL});
  sh(s,'rect',6.67,5.9,4.03,1.6,{f:OR});
  tx(s,1.74,4.52,3.89,0.68,[['Lorem ipsum dolor sit amet, anasanu consectetun amaimi eu antan onecail eget']],{sz:12,f:LT,c:WH,ls:1.5});
  icon(s,1.04,4.66,0.384,OR,'check');
  tx(s,1.74,5.34,3.89,0.68,[['Lorem ipsum dolor sit amet, anasanu consectetun amaimi eu antan onecail eget']],{sz:12,f:LT,c:WH,ls:1.5});
  icon(s,1.04,5.49,0.384,OR,'check');
  tx(s,1.74,6.06,3.89,0.68,[['Lorem ipsum dolor sit amet, anasanu consectetun amaimi eu antan onecail eget']],{sz:12,f:LT,c:WH,ls:1.5});
  icon(s,1.04,6.21,0.384,OR,'check');
  tx(s,7.48,0.89,5.04,1.77,[['A Different Kind Of Company.'],[' A Different Kind Of Cargo.',{c:OR}]],{sz:28,f:PS,ls:1.2});
  ln(s,8.41,2.86,0.75,0,OR,{lw:3,rot:180});
  ln(s,7.67,2.86,0.75,0,BK,{lw:3,rot:180});
  tx(s,6.95,6.73,1.56,0.364,[['Happy Client']],{sz:12,f:LT,c:WH,al:'c',ls:1.5});
  tx(s,6.95,6.3,1.56,0.56,[['1000+']],{sz:20,f:PS,c:WH,al:'c',ls:1.5});
  tx(s,8.85,6.73,1.56,0.364,[['Paket Delivered']],{sz:12,f:LT,c:WH,al:'c',ls:1.5});
  tx(s,8.85,6.3,1.56,0.56,[['2500+']],{sz:20,f:PS,c:WH,al:'c',ls:1.5});
  ln(s,8.66,6.2,0,0.99,WH);
  dots(s,11.84,0.3,0.22,0.25,0.066,OR,1);
}

/** Slide 8 — What we do — service cards (layout: 8_Custom Layout) */
function slide08(s) {
  sh(s,'rect',0,0,5.43,7.49,{f:BK,a:44,sd:{type:"outer",blur:35,offset:0,angle:25,color:"000000",opacity:0.4}});
  sh(s,'roundRect',4.68,0.325,8.47,6.85,{f:WH,r:0.155});
  sh(s,'rect',6.47,3.87,2.98,3.02,{f:BK,a:44,sd:{type:"outer",blur:35,offset:0,angle:25,color:"000000",opacity:0.4}});
  sh(s,'rect',9.95,4.07,2.58,2.62,{f:WH,sd:{type:"outer",blur:35,offset:0,angle:25,color:"000000",opacity:0.4}});
  tx(s,7.06,4.97,2.21,0.375,[[LID]],{sz:12,f:PS,c:WH,cs:0.2,ls:1.5});
  tx(s,10.27,5.47,1.96,0.97,[[PEL]],{sz:12,f:LT,c:GY,ls:1.5});
  tx(s,10.27,4.97,2.21,0.375,[[LID]],{sz:12,f:PS,cs:0.2,ls:1.5});
  sh(s,'rect',9.95,0.86,2.58,2.62,{f:WH,sd:{type:"outer",blur:35,offset:0,angle:25,color:"000000",opacity:0.4}});
  sh(s,'rect',6.67,0.86,2.58,2.62,{f:WH,sd:{type:"outer",blur:35,offset:0,angle:25,color:"000000",opacity:0.4}});
  icon(s,7.06,1,0.74,OR,'truck');
  tx(s,7.06,2.26,1.96,0.97,[[PEL]],{sz:12,f:LT,c:GY,ls:1.5});
  tx(s,7.06,1.76,2.21,0.375,[[LID]],{sz:12,f:PS,cs:0.2,ls:1.5});
  icon(s,10.27,0.94,0.74,OR,'plane');
  icon(s,7.05,4.27,0.74,WH,'ship');
  icon(s,10.21,4.26,0.77,OR,'train');
  tx(s,10.27,2.26,1.96,0.97,[[PEL]],{sz:12,f:LT,c:GY,ls:1.5});
  tx(s,10.27,1.76,2.21,0.375,[[LID]],{sz:12,f:PS,cs:0.2,ls:1.5});
  tx(s,7.06,5.47,1.96,0.97,[[PEL]],{sz:12,f:LT,c:WH,ls:1.5});
  sh(s,'rect',6.67,4.05,2.6,2.59,{l:WH});
  socGlyph(s,0.75,6.8,WH);
  tx(s,0.65,1.73,3.86,1.77,[['We Focuse And Deliver  Packet In Your Distance']],{sz:28,f:PS,c:WH,ls:1.2});
  ln(s,1.48,3.69,0.75,0,OR,{lw:3,rot:180});
  ln(s,0.74,3.69,0.75,0,WH,{lw:3,rot:180});
  tx(s,0.65,1.28,2.86,0.406,[['What we do']],{sz:16,f:PS,c:OR,ls:1.2});
  tx(s,0.65,4.07,3.46,1.27,[[VOC.slice(0,161)+' ']],{sz:12,f:LT,c:WH,ls:1.5});
  sh(s,'roundRect',0.65,5.73,1.58,0.479,{f:OR,r:0.071});
  tx(s,0.75,5.75,1.03,0.364,[['Read More']],{sz:12,f:LT,c:WH,al:'c',ls:1.5});
  icon(s,1.74,5.88,0.199,WH,'play');
  tx(s,3.5,3.19,3.87,1.13,[['ABOUT US']],{sz:54,f:PS,c:WS,ls:1.2,rot:270});
  dots(s,3.36,0.38,0.18,0.2,0.052,OR,1);
}

/** Slide 9 — How we work — 4 steps (layout: 9_Custom Layout) */
function slide09(s) {
  sh(s,'rect',0,0,13.33,7.49,{f:BK,a:34,sd:{type:"outer",blur:35,offset:0,angle:25,color:"000000",opacity:0.4}});
  tx(s,4.72,1.11,5.85,1.2,[['We Focuse  Deliver  ',{br:1}],['Packet In Your Destination']],{sz:28,f:PS,c:WH,al:'c',ls:1.2});
  ln(s,7.65,2.56,0.75,0,OR,{lw:3,rot:180});
  ln(s,6.9,2.56,0.75,0,WH,{lw:3,rot:180});
  tx(s,6.22,0.69,2.86,0.406,[['How We Work']],{sz:16,f:PS,c:OR,al:'c',ls:1.2});
  ln(s,4.96,4.04,5.62,0.011,WH,{lw:1.5,ld:'lgDashDot'});
  sh(s,'ellipse',3.88,3.5,1.08,1.08,{f:OR,l:WH,lw:1.5});
  tx(s,3.24,4.81,2.36,0.65,[['Creat',{br:1}],['  Your Account']],{sz:14,f:PS,c:WH,cs:1,al:'c',ls:1.2});
  icon(s,4.08,3.76,0.67,WH,'users');
  sh(s,'ellipse',3.75,3.44,0.426,0.426,{f:BK,l:WH,lw:1.5});
  tx(s,5.91,4.81,2.21,0.65,[['Select',{br:1}],[' your  Destination']],{sz:14,f:PS,c:WH,cs:1,al:'c',ls:1.2});
  sh(s,'ellipse',6.34,3.5,1.08,1.08,{f:OR,l:WH,lw:1.5});
  icon(s,6.54,3.74,0.67,WH,'direction');
  sh(s,'ellipse',6.28,3.44,0.426,0.426,{f:BK,l:WH,lw:1.5});
  sh(s,'ellipse',8.53,3.5,1.08,1.08,{f:OR,l:WH,lw:1.5});
  tx(s,8.13,4.81,1.87,0.65,[['We ',{br:1}],['Collect It']],{sz:14,f:PS,c:WH,cs:1,al:'c',ls:1.2});
  icon(s,8.73,3.76,0.67,WH,'box');
  sh(s,'ellipse',8.4,3.44,0.426,0.426,{f:BK,l:WH,lw:1.5});
  sh(s,'ellipse',10.58,3.51,1.08,1.08,{f:OR,l:WH,lw:1.5});
  tx(s,10.18,4.8,1.87,0.65,[['We ',{br:1}],['Deliveret']],{sz:14,f:PS,c:WH,cs:1,al:'c',ls:1.2});
  icon(s,10.75,3.78,0.67,WH,'truck');
  sh(s,'ellipse',10.54,3.44,0.426,0.426,{f:BK,l:WH,lw:1.5});
  tx(s,3.88,3.47,0.198,0.368,[['1']],{sz:14,f:PS,c:WH,cs:1,al:'c',ls:1.2});
  tx(s,6.41,3.47,0.198,0.368,[['2']],{sz:14,f:PS,c:WH,cs:1,al:'c',ls:1.2});
  tx(s,8.53,3.47,0.198,0.368,[['3']],{sz:14,f:PS,c:WH,cs:1,al:'c',ls:1.2});
  tx(s,10.65,3.47,0.198,0.368,[['4']],{sz:14,f:PS,c:WH,cs:1,al:'c',ls:1.2});
  tx(s,3.84,5.84,7.61,0.97,[[VOC.slice(0,231)+' ']],{sz:12,f:LT,c:WH,al:'c',ls:1.5});
}

/** Slide 10 — Our team (layout: 10_Custom Layout) */
function slide10(s) {
  sh(s,'rect',0.75,0,12.58,7.5,{f:OR});
  sh(s,'rect',0,0,0.75,7.5,{f:BK});
  tx(s,10.72,2.98,2.1,0.337,[['Laura Doel']],{sz:14,f:PS,c:WH,cs:1,al:'c'});
  tx(s,10.89,3.29,1.76,0.471,[['Manager Oprasional']],{sz:11,f:LT,c:WH,cs:1,b:1,al:'c'});
  tx(s,10.72,6.15,2.1,0.337,[['Andre Smith']],{sz:14,f:PS,c:WH,cs:1,al:'c'});
  tx(s,10.89,6.46,1.76,0.286,[['Head Manager']],{sz:11,f:LT,c:WH,cs:1,b:1,al:'c'});
  tx(s,5.97,2.98,2.1,0.337,[['Andi Smith']],{sz:14,f:PS,c:WH,cs:1,al:'c'});
  tx(s,6.14,3.29,1.76,0.286,[['Product MANAGER']],{sz:11,f:LT,c:WH,cs:1,b:1,al:'c'});
  tx(s,5.97,6.15,2.1,0.337,[['Jaka Smith']],{sz:14,f:PS,c:WH,cs:1,al:'c'});
  tx(s,6.14,6.46,1.76,0.286,[['Vice President']],{sz:11,f:LT,c:WH,cs:1,b:1,al:'c'});
  tx(s,8.34,4.81,2.1,0.337,[['Vincent Dhoe']],{sz:14,f:PS,c:WH,cs:1,al:'c'});
  tx(s,8.51,5.12,1.76,0.286,[['Ceo & Founder']],{sz:11,f:LT,c:WH,cs:1,b:1,al:'c'});
  socRing(s,0.223,0.7,0.321,0.487,WH,false);
  ln(s,0.383,2.7,0,1.75,WH,{lw:1.25});
  tx(s,0.172,4.72,0.423,2.08,[['www.'],['carlistic',{cs:1.2}],['.com']],{sz:17,f:PP,c:WH,al:'r',nw:1,vt:'vert270'});
  dots(s,1.07,0.41,0.2,0.18,0.052,BK);
  tx(s,1.43,1.44,1.33,0.37,[['Our Team']],{sz:16,f:PS});
  tx(s,1.43,4.04,3.96,1.27,[[SED]],{sz:12,f:LT,c:WH,ls:1.5});
  tx(s,1.43,1.79,4.12,1.51,[['We Focuse And '],['Deliver  Packet',{c:BK,b:1,br:1}],['In Your Distance']],{sz:28,f:PP,c:WH});
  ln(s,2.27,3.58,0.75,0,WH,{lw:3,rot:180});
  ln(s,1.52,3.58,0.75,0,BK,{lw:3,rot:180});
  socGlyph(s,1.54,6.71,WH);
  tx(s,1.43,5.44,3.96,0.67,[[SED.slice(0,84)]],{sz:12,f:LT,c:WH,ls:1.5});
}

/** Slide 11 — Meet our team leaders (layout: 11_Custom Layout) */
function slide11(s) {
  sh(s,'rect',0,2.36,13.33,4.94,{f:OR});
  sh(s,'rect',0,2.48,13.33,4.66,{f:BK});
  sh(s,'rect',0,4.19,13.33,1.74,{f:OR});
  grad(s,6.81,2.96,2.17,3.81,'v',[WH,100,38],[OR,26,74],WH);
  tx(s,6.84,5.59,1.99,0.337,[['SAMPAER DHOE']],{sz:14,f:PS,c:WH,cs:0.5,al:'c'});
  tx(s,7.16,5.9,1.35,0.286,[['Vice President']],{sz:11,f:LT,c:WH,al:'c'});
  dots(s,1.04,0.42,0.2,0.18,0.052,LG);
  tx(s,1.54,0.65,3.68,1.12,[['Meet Our ',{br:1}],['Team '],['Leaders.',{c:OR}]],{sz:28,f:PM,cs:2,b:1,ls:1.1});
  ln(s,2.43,1.89,0.75,0,BK,{lw:3,rot:180});
  ln(s,1.69,1.89,0.75,0,OR,{lw:3,rot:180});
  tx(s,6.58,0.8,5.71,0.97,[[SED]],{sz:12,f:LT,c:GY,ls:1.5});
  socGlyph(s,7.37,6.25,WH);
}

/** Slide 12 — Meet the team (layout: 12_Custom Layout) */
function slide12(s) {
  sh(s,'rtTriangle',0,2.71,4.79,4.79,{f:OR,rot:90,fh:1});
  sh(s,'rtTriangle',7.78,0,5.56,4.79,{f:OR,rot:180});
  sh(s,'ellipse',6.54,4.33,0.97,0.97,{f:DG});
  tx(s,7.9,4.25,4.77,0.97,[[SED]],{sz:12,f:LT,c:GY,ls:1.5});
  tx(s,7.9,5.45,4.38,0.67,[[SED.slice(0,84)]],{sz:12,f:LT,c:GY,ls:1.5});
  icon(s,6.79,4.58,0.472,WH,'gem');
  dots(s,0.52,0.66,0.18,0.19,0.052,LN);
  ln(s,10.09,6.07,0,1.55,GY,{lw:1.25,rot:90});
  tx(s,7.79,5.95,0.404,1.78,[['www.'],['carlistic',{cs:1.2}],['.com']],{sz:12,f:PP,c:GY,al:'r',nw:1,vt:'vert270',rot:90});
  socRing(s,11.33,6.7,0.283,0.431,GY,true);
  tx(s,0.85,1.34,3.05,1.41,[['Meet ',{br:1}],['The Team. ',{c:OR}]],{sz:36,f:PM,cs:2,b:1,ls:1.1});
  ln(s,1.05,2.89,0.75,0,BK,{lw:3,rot:180,fh:1});
  ln(s,1.8,2.89,0.75,0,OR,{lw:3,rot:180,fh:1});
}

/** Slide 13 — Personal data — Sampaer (layout: 14_Custom Layout) */
function slide13(s) {
  sh(s,'rtTriangle',0,2.71,4.79,4.79,{f:OR,rot:90,fh:1});
  sh(s,'rtTriangle',11.35,0,1.98,1.71,{f:OR,rot:180});
  dots(s,0.52,0.66,0.24,0.27,0.071,OR);
  ln(s,5.09,2.34,0.75,0,BK,{lw:3,rot:270});
  ln(s,5.09,1.6,0.75,0,OR,{lw:3,rot:270});
  tx(s,5.36,3.03,3.39,1.58,[[SED]],{sz:12,f:LT,c:GY,ls:1.5});
  tx(s,5.36,4.77,3.11,0.97,[[SED.slice(0,84)]],{sz:12,f:LT,c:GY,ls:1.5});
  tx(s,5.83,1.16,4.06,1.26,[['PERSONAL DATA ',{c:BK}],['SAMPAER',{sz:36}],['  ']],{sz:28,f:PM,c:OR,cs:2,b:1,ls:1.1});
  tx(s,5.9,2.41,2.04,0.37,[['Vice President']],{sz:16,f:LT});
  socDisc(s,5.53,6.11);
  sh(s,'roundRect',9.14,2.85,3.22,3.27,{f:WH,r:0.096,sd:{type:"outer",blur:25,offset:3,angle:45,color:"000000",opacity:0.4}});
  icon(s,9.46,3.14,0.482,OR,'marker');
  icon(s,9.45,5.32,0.482,OR,'world');
  icon(s,9.45,3.9,0.482,OR,'phone');
  icon(s,9.51,4.69,0.362,OR,'envelope');
  tx(s,10.11,3.98,1.98,0.342,[['(0295) 1234 567']],{sz:10.5,f:PS,ls:1.5});
  tx(s,10.11,3.14,2.7,0.62,[['1234 WASHINGTON DC',{br:1}],['CA 12345, UNITED STATES']],{sz:10.5,f:PS,ls:1.5});
  tx(s,10.11,4.7,2.42,0.341,[['OFFICE@CARLISTIC.COM']],{sz:10.5,f:PS,ls:1.5});
  tx(s,10.11,5.4,2.17,0.341,[['WWW. CARLISTIC.COM']],{sz:10.5,f:PS,ls:1.5});
  ln(s,10.18,6.19,0,1.55,GY,{lw:1.25,rot:90});
  tx(s,7.88,6.08,0.404,1.78,[['www.'],['carlistic',{cs:1.2}],['.com']],{sz:12,f:PP,c:GY,al:'r',nw:1,vt:'vert270',rot:90});
  socRing(s,11.42,6.83,0.283,0.431,GY,true);
  dots(s,12.02,0.27,0.18,0.17,0.052,WH,1);
}

/** Slide 14 — We have the best team (layout: 13_Custom Layout) */
function slide14(s) {
  sh(s,'rect',0,0,6.06,7.5,{f:OR});
  ln(s,9.96,6.04,0,1.55,GY,{lw:1.25,rot:90});
  tx(s,7.65,5.92,0.404,1.78,[['www.'],['carlistic',{cs:1.2}],['.com']],{sz:12,f:PP,c:GY,al:'r',nw:1,vt:'vert270',rot:90});
  socRing(s,11.19,6.67,0.283,0.431,GY,true);
  dots(s,0.52,0.66,0.18,0.19,0.052,LN);
  tx(s,0.91,1.34,3.69,2.15,[['We Have',{br:1}],['The Best Team',{c:BK}],[' Delivered Your Package']],{sz:28,f:PS,c:WH,cs:2,b:1,ls:1.1});
  ln(s,1.05,3.95,0.75,0,BK,{lw:3,rot:180,fh:1});
  ln(s,1.8,3.95,0.75,0,WH,{lw:3,rot:180,fh:1});
  tx(s,0.94,4.4,3.39,1.58,[[SED]],{sz:12,f:LT,c:WH,ls:1.5});
  sh(s,'rect',10.43,3.75,2.29,2.31,{f:OR});
  icon(s,10.35,3.69,2.42,BK,'cycle');
}

/** Slide 15 — Our services — 4 icons (layout: 16_Custom Layout) */
function slide15(s) {
  dots(s,0.52,0.51,0.24,0.27,0.071,LG);
  tx(s,1.04,0.89,6.12,1.12,[['For Over 12 Years In The Bussiness Cargo & Logistic']],{sz:28,f:PS,cs:1.2,ls:1.1});
  tx(s,1.04,0.416,2.15,0.418,[['Our Services']],{f:PM,c:OR,cs:2,ls:1.1});
  tx(s,7.89,0.79,4.78,0.97,[[SED]],{sz:12,f:LT,c:GY,ls:1.5});
  tx(s,0.55,6.09,2.21,0.375,[[LID]],{sz:12,f:PS,cs:0.2,al:'c',ls:1.5});
  tx(s,7.08,6.52,2.5,0.67,[[PEL]],{sz:12,f:LT,c:GY,al:'c',ls:1.5});
  tx(s,7.23,6.09,2.21,0.375,[[LID]],{sz:12,f:PS,cs:0.2,al:'c',ls:1.5});
  icon(s,4.63,2.54,0.74,OR,'truck');
  tx(s,3.76,3.88,2.48,0.67,[[PEL]],{sz:12,f:LT,c:GY,al:'c',ls:1.5});
  tx(s,3.9,3.39,2.21,0.375,[[LID]],{sz:12,f:PS,cs:0.2,al:'c',ls:1.5});
  icon(s,11.39,2.71,0.55,OR,'plane');
  icon(s,1.29,5.29,0.74,OR,'ship');
  icon(s,7.95,5.28,0.77,OR,'train');
  tx(s,10.36,3.88,2.61,0.67,[[PEL]],{sz:12,f:LT,c:GY,al:'c',ls:1.5});
  tx(s,10.56,3.39,2.21,0.375,[[LID]],{sz:12,f:PS,cs:0.2,al:'c',ls:1.5});
  tx(s,0.219,6.52,2.88,0.67,[[PEL]],{sz:12,f:LT,c:GY,al:'c',ls:1.5});
  sh(s,'rtTriangle',0,3.81,1.15,1.15,{f:OR,a:22});
  sh(s,'rtTriangle',5.52,6.35,1.15,1.15,{f:OR,a:22,fh:1});
  sh(s,'rtTriangle',6.67,3.81,1.15,1.15,{f:OR,a:22});
  sh(s,'rtTriangle',12.2,6.35,1.15,1.15,{f:OR,a:22,fh:1});
}

/** Slide 16 — 12 Years in business (layout: 17_Custom Layout) */
function slide16(s) {
  sh(s,'roundRect',4.15,2.53,4.48,4.46,{l:OR,lw:3,r:0.033});
  sh(s,'rect',12.58,0,0.75,7.5,{f:BK});
  tx(s,2.45,0.89,8.43,0.6,[[' 12 Years In  Bussiness Cargo & Logistic']],{sz:28,f:PS,cs:1,al:'c',ls:1.1});
  tx(s,5.59,0.416,2.15,0.418,[['Our Services']],{f:PM,c:OR,cs:2,al:'c',ls:1.1});
  tx(s,9.2,2.94,2.98,1.27,[[SED.slice(0,112)],['.'],['.']],{sz:12,f:LT,c:GY,ls:1.5});
  tx(s,0.68,4.34,3.11,0.9,[[SED.slice(0,84)]],{sz:11,f:LT,c:GY,al:'r',ls:1.5});
  tx(s,1.58,3.9,2.21,0.375,[[LID]],{sz:12,f:PS,cs:0.2,al:'r',ls:1.5});
  tx(s,0.68,5.82,3.11,0.9,[[SED.slice(0,84)]],{sz:11,f:LT,c:GY,al:'r',ls:1.5});
  tx(s,1.58,5.37,2.21,0.375,[[LID]],{sz:12,f:PS,cs:0.2,al:'r',ls:1.5});
  tx(s,0.68,2.72,3.11,0.9,[[SED.slice(0,84)]],{sz:11,f:LT,c:GY,al:'r',ls:1.5});
  tx(s,1.58,2.37,2.21,0.375,[[LID]],{sz:12,f:PS,cs:0.2,al:'r',ls:1.5});
  ln(s,12.94,2.7,0,1.75,WH,{lw:1.25});
  tx(s,12.73,4.72,0.423,2.08,[['www.'],['carlistic',{cs:1.2}],['.com']],{sz:17,f:PP,c:WH,al:'r',nw:1,vt:'vert270'});
  socRing(s,12.78,0.7,0.321,0.487,WH,false);
  dots(s,1.03,0.55,0.18,0.2,0.052,LG);
}

/** Slide 17 — Best services for our client (layout: 18_Custom Layout) */
function slide17(s) {
  sh(s,'rect',0,1.64,13.33,2.21,{f:OR});
  sh(s,'roundRect',1.06,0.56,11.11,4.37,{f:BK,a:36,r:0.128});
  sh(s,'roundRect',1.54,4.38,3.01,2.32,{f:WH,l:LN,r:0.109,sd:{type:"outer",blur:35,offset:0,angle:0,color:"A6A6A6",opacity:0.4}});
  tx(s,2.03,4.76,2.03,0.379,[[LID]],{sz:12,f:PS,cs:0.2,al:'c',ls:1.5});
  tx(s,1.86,5.18,2.54,1.19,[[SED.slice(0,112)],['.'],['.']],{sz:11,f:LT,c:GY,al:'c',ls:1.5});
  sh(s,'roundRect',5.16,4.38,3.01,2.32,{f:WH,l:LN,r:0.109,sd:{type:"outer",blur:35,offset:0,angle:0,color:"A6A6A6",opacity:0.4}});
  tx(s,5.65,4.76,2.03,0.379,[[LID]],{sz:12,f:PS,cs:0.2,al:'c',ls:1.5});
  tx(s,5.48,5.18,2.54,1.19,[[SED.slice(0,112)],['.'],['.']],{sz:11,f:LT,c:GY,al:'c',ls:1.5});
  sh(s,'roundRect',8.78,4.38,3.01,2.32,{f:WH,l:LN,r:0.109,sd:{type:"outer",blur:35,offset:0,angle:0,color:"A6A6A6",opacity:0.4}});
  tx(s,9.27,4.76,2.03,0.379,[[LID]],{sz:12,f:PS,cs:0.2,al:'c',ls:1.5});
  tx(s,9.1,5.18,2.54,1.19,[[SED.slice(0,112)],['.'],['.']],{sz:11,f:LT,c:GY,al:'c',ls:1.5});
  tx(s,6.98,2.26,4.78,0.97,[[SED]],{sz:12,f:LT,c:WH,ls:1.5});
  tx(s,1.86,2.58,3.86,1.04,[['Best Services For Our Client']],{sz:28,f:PS,c:WH,cs:1.2,b:1});
  tx(s,1.86,1.88,2.37,0.386,[['Our  '],['Services']],{sz:16,f:PS,c:WH,cs:2,b:1,ls:1.1});
  ln(s,1.95,2.37,0.75,0,OR,{lw:3,rot:180,fh:1});
  ln(s,2.7,2.37,0.75,0,WH,{lw:3,rot:180,fh:1});
}

/** Slide 18 — Services — 3 outlined cards (layout: 19_Custom Layout) */
function slide18(s) {
  sh(s,'rect',0,0,13.33,7.5,{f:BK,a:36});
  tx(s,2.45,1.64,8.43,0.6,[[' 12 Years In  Bussiness Cargo & Logistic']],{sz:28,f:PS,c:WH,al:'c',ls:1.1});
  tx(s,5.59,1.17,2.15,0.418,[['Our Services']],{f:PM,c:WH,cs:2,al:'c',ls:1.1});
  tx(s,2.66,2.29,7.86,0.67,[[VOC.slice(0,161)],[' ']],{sz:12,f:LT,c:WH,al:'c',ls:1.5});
  sh(s,'roundRect',0.98,3.64,3.37,2.69,{l:WH,r:0.126});
  tx(s,1.65,4.57,2.03,0.379,[['LOREM IPSUM DUMET']],{sz:12,f:PS,c:WH,cs:0.2,al:'c',ls:1.5});
  tx(s,1.21,5.1,2.9,0.9,[[SED.slice(0,97)+' ']],{sz:11,f:LT,c:WH,al:'c',ls:1.5});
  icon(s,2.29,3.73,0.74,WH,'ship');
  sh(s,'roundRect',4.98,3.64,3.37,2.69,{l:WH,r:0.126});
  tx(s,5.65,4.57,2.03,0.379,[['LOREM IPSUM DUMET']],{sz:12,f:PS,c:WH,cs:0.2,al:'c',ls:1.5});
  tx(s,5.22,5.1,2.9,0.9,[[SED.slice(0,97)+' ']],{sz:11,f:LT,c:WH,al:'c',ls:1.5});
  icon(s,6.3,3.83,0.74,WH,'truck');
  sh(s,'roundRect',8.98,3.64,3.37,2.69,{l:WH,r:0.126});
  tx(s,9.65,4.57,2.03,0.379,[['LOREM IPSUM DUMET']],{sz:12,f:PS,c:WH,cs:0.2,al:'c',ls:1.5});
  tx(s,9.22,5.1,2.9,0.9,[[SED.slice(0,97)+' ']],{sz:11,f:LT,c:WH,al:'c',ls:1.5});
  icon(s,10.32,3.74,0.7,WH,'plane');
}

/** Slide 19 — Break slide (layout: 1_Custom Layout) */
function slide19(s) {
  sh(s,'rect',0,0,13.33,7.5,{f:BK});
  sh(s,'rect',1.03,1.01,11.27,5.6,{f:BK,a:42});
  sh(s,'rect',0.8,0.69,11.73,6.12,{l:OR,lw:2.75});
  sh(s,'rtTriangle',10.95,0.427,1.83,1.94,{f:OR,rot:180});
  socDisc(s,1.46,1.42);
  sh(s,'ellipse',4.78,2.55,0.197,0.197,{f:OR,fh:1});
  sh(s,'ellipse',5.1,2.55,0.197,0.197,{f:OR,fh:1});
  sh(s,'ellipse',4.45,2.55,0.197,0.197,{f:OR,fh:1});
  sh(s,'ellipse',4.1,2.55,0.197,0.197,{f:OR,fh:1});
  tx(s,2.48,2.37,8.37,1.72,[['Break'],['slide',{c:OR}]],{sz:96,f:PS,c:WH,al:'c'});
  tx(s,6.48,4.09,0.379,1.71,[[WWW]],{sz:10.5,f:PP,c:WH,cs:1.2,al:'r',nw:1,vt:'vert270',rot:90});
  tx(s,3.2,3.88,6.93,0.67,[[SED]],{sz:12,f:LT,c:WH,al:'c',ls:1.5});
}

/** Slide 20 — Best portfolio (layout: 20_Custom Layout) */
function slide20(s) {
  sh(s,'rect',-0.214,0,0.75,7.5,{f:OR});
  sh(s,'rect',0.54,0,12.79,4.31,{f:BK});
  tx(s,8.46,2.43,4.05,1.64,[['Best Portfolio '],['Delivered Your Package',{c:OR}]],{sz:28,f:PS,c:WH,cs:2,b:1,ls:1.1});
  tx(s,8.53,4.86,3.87,1.13,[[VOC.slice(0,161)+' ']],{sz:10.5,f:LT,c:GY,ls:1.5});
  tx(s,0.77,1.19,1.93,0.82,[[SED.slice(0,72)+' ']],{sz:10,f:LT,c:WH,al:'r',ls:1.5});
  tx(s,0.324,0.76,2.38,0.398,[[LID]],{sz:13,f:PS,c:WH,cs:0.2,al:'r',ls:1.5});
  tx(s,3.43,5.83,1.93,0.82,[[SED.slice(0,72)+' , ']],{sz:10,f:LT,c:GY,al:'r',ls:1.5});
  tx(s,2.98,5.39,2.38,0.398,[[LID]],{sz:13,f:PS,cs:0.2,al:'r',ls:1.5});
  tx(s,5.6,1.24,1.93,0.82,[[SED.slice(0,72)+' ']],{sz:10,f:LT,c:WH,ls:1.5});
  tx(s,5.6,0.8,2.38,0.398,[[LID]],{sz:13,f:PS,c:WH,cs:0.2,ls:1.5});
  socDisc(s,8.57,6.91);
  ln(s,10.04,7.06,1.11,0,GY,{lw:1.25,fh:1});
  tx(s,11.96,6.27,0.379,1.58,[['www.'],['carlistic',{cs:1.2}],['.com']],{sz:10.5,f:PP,c:GY,al:'r',nw:1,vt:'vert270',rot:90});
  dots(s,11.77,0.33,0.18,0.18,0.052,WH,1);
}

/** Slide 21 — Our portfolio — ocean freight (layout: 21_Custom Layout) */
function slide21(s) {
  sh(s,'rect',0,0,12.58,3.75,{f:BK,fh:1});
  sh(s,'rect',10.46,0,2.88,7.5,{f:OR,fh:1});
  tx(s,1.01,1.67,4.09,1.64,[['Best Portfolio ',{br:1}],['Delivered Your Package']],{sz:28,f:PS,c:WH,cs:2,b:1,ls:1.1});
  tx(s,1.01,4.33,3.87,1.88,[[VOC.slice(0,256)+'.']],{sz:12,f:LT,c:GY,ls:1.5});
  tx(s,1.05,1.2,2,0.35,[['Our Portfolio']],{sz:14,f:PS,c:OR,ls:1.1});
  sh(s,'rect',5.96,3.65,3.18,1.23,{f:OR,a:42});
  tx(s,6.28,4.19,2.72,0.422,[['OCEAN FREIGHT']],{f:PS,c:WH,cs:2,b:1,ls:1.1});
  socDisc(s,1.09,6.91);
  ln(s,2.56,7.06,1.11,0,GY,{lw:1.25,fh:1});
  tx(s,4.48,6.27,0.379,1.58,[['www.'],['carlistic',{cs:1.2}],['.com']],{sz:10.5,f:PP,c:GY,al:'r',nw:1,vt:'vert270',rot:90});
  dots(s,10.65,0.35,0.34,0.34,0.102,WH,1);
}

/** Slide 22 — Exclusive gallery (layout: 22_Custom Layout) */
function slide22(s) {
  sh(s,'custGeom',5.87,-1,8.78,10.4,{f:OR,a:34,rot:33.99,pts:[[0,0.3],[0.52,0],[1,0.6],[0.29,1]]});
  tx(s,10.51,1.62,2.19,0.62,[['Lorem ipsum dolor sit amet consecteturil adipiscing elit. ']],{sz:11,f:LT,c:WH,ls:1.5});
  tx(s,10.51,1.2,2.19,0.32,[['OCEAN FREIGHT']],{sz:13,f:PS,c:WH,cs:3});
  tx(s,10.05,3.69,2.19,0.62,[['Lorem ipsum dolor sit amet consecteturil adipiscing elit. ']],{sz:11,f:LT,c:WH,ls:1.5});
  tx(s,10.05,3.27,1.85,0.32,[['AIR FREIGHT']],{sz:13,f:PS,c:WH,cs:3});
  tx(s,0.85,1.65,4.05,1.04,[['Exclusive ',{br:1}],['Carlistic  Gallery']],{sz:28,f:PS,c:WH,cs:1.2,b:1});
  tx(s,0.85,1.16,2,0.35,[['Our Portfolio']],{sz:14,f:PS,c:OR,ls:1.1});
  tx(s,0.88,2.87,3.31,1.88,[[VOC.slice(0,233)]],{sz:12,f:LT,c:WH,ls:1.5});
  sh(s,'ellipse',8.6,5.48,0.69,0.69,{f:WH});
  tx(s,9.59,5.72,2.19,0.62,[['Lorem ipsum dolor sit amet consecteturil adipiscing elit. ']],{sz:11,f:LT,c:WH,ls:1.5});
  tx(s,9.59,5.3,2.32,0.303,[['RAILWAY FREIGHT']],{sz:12,f:PP,c:WH,cs:3,b:1});
  icon(s,8.73,5.62,0.41,OR,'train');
  sh(s,'ellipse',9.1,3.44,0.69,0.69,{f:WH});
  sh(s,'ellipse',9.69,1.4,0.69,0.69,{f:WH});
  icon(s,9.85,1.54,0.389,OR,'ship');
  icon(s,9.18,3.52,0.5,OR,'plane');
  socDisc(s,0.99,5.2);
}

/** Slide 23 — Freight list (layout: 23_Custom Layout) */
function slide23(s) {
  sh(s,'rect',0,0,13.33,7.5,{f:BK});
  sh(s,'rtTriangle',0.343,-0.343,4.29,4.98,{f:OR,rot:270,fh:1,fv:1});
  sh(s,'rtTriangle',11.35,5.79,1.98,1.71,{f:OR,rot:180,fv:1});
  tx(s,5.21,4.63,2.19,0.62,[['Lorem ipsum dolor sit amet consecteturil adipiscing elit. ']],{sz:11,f:LT,c:WH,ls:1.5});
  tx(s,5.21,4.21,1.85,0.32,[['AIR FREIGHT']],{sz:13,f:PS,c:WH,cs:3});
  tx(s,1.03,5.99,2.19,0.62,[['Lorem ipsum dolor sit amet consecteturil adipiscing elit. ']],{sz:11,f:LT,c:WH,ls:1.5});
  tx(s,1.03,5.57,2.32,0.303,[['RAILWAY FREIGHT']],{sz:12,f:PP,c:WH,cs:3,b:1});
  tx(s,9.86,1.19,2.19,0.62,[['Lorem ipsum dolor sit amet consecteturil adipiscing elit. ']],{sz:11,f:LT,c:WH,ls:1.5});
  tx(s,9.86,0.76,2.19,0.32,[['OCEAN FREIGHT']],{sz:13,f:PS,c:WH,cs:3});
  tx(s,6.85,1.18,2.19,0.62,[['Lorem ipsum dolor sit amet consecteturil adipiscing elit. ']],{sz:11,f:LT,c:WH,ls:1.5});
  tx(s,6.85,0.87,2.32,0.303,[['ROAD FREIGHT']],{sz:12,f:PP,c:WH,cs:3,b:1});
}

/** Slide 24 — Portfolio (light) (layout: 24_Custom Layout) */
function slide24(s) {
  sh(s,'rect',10.46,0,2.88,7.5,{f:BK,fh:1});
  sh(s,'rect',0.8,1.53,4.68,1.11,{f:OR});
  tx(s,0.8,1.53,4.68,1.11,[['PORTFOLIO.']],{sz:44,f:PS,c:WH,cs:6,al:'c',ls:1.5,bg:OR});
  sh(s,'rect',0.8,-3.74,1,1,{f:OR});
  tx(s,0.8,3.14,2.34,1.73,[[VOC.slice(0,161)+' , ']],{sz:11,f:LT,c:GY,al:'r',ls:1.5});
  socDisc(s,1.93,6.37);
  sh(s,'rect',9.9,6.27,2.95,0.67,{f:WH});
  tx(s,11.13,5.36,0.488,2.54,[[WWW]],{sz:17,f:PP,cs:1.2,al:'r',nw:1,vt:'vert270',rot:90});
  sh(s,'rtTriangle',12.42,6.52,0.57,0.57,{f:OR,fh:1});
  sh(s,'rtTriangle',9.75,6.12,0.57,0.57,{f:OR,rot:90});
}

/** Slide 25 — Portfolio (dark) (layout: 25_Custom Layout) */
function slide25(s) {
  sh(s,'rect',0,0,13.33,7.5,{f:BK});
  tx(s,0.68,3.89,4.03,0.92,[['PORTFOLIO.']],{sz:36,f:PS,c:OR,cs:6,ls:1.5});
  tx(s,1.01,5.09,3.15,1.58,[[VOC.slice(0,161)+' , ']],{sz:12,f:LT,c:WH,al:'r',ls:1.5});
  sh(s,'rect',0.56,1.79,1.2,0.466,{f:OR});
  sh(s,'rect',9.94,6.67,1.2,0.466,{f:OR});
}

/** Slide 26 — SWOT — Strength (layout: 26_Custom Layout) */
function slide26(s) {
  sh(s,'roundRect',2.44,-1.97,6.07,10.96,{f:BK,a:30,rot:90,r:0.487});
  tx(s,1.09,1.84,4.1,1.2,[['Strength',{c:OR}],['  Our of',{br:1}],['Company.']],{sz:28,f:PS,c:WH,cs:1.2,b:1,ls:1.2});
  tx(s,1.09,1.19,1.22,0.406,[['SWOT']],{sz:16,f:PS,c:OR,cs:1.2,ls:1.2});
  letter(s,'S',5.26,1.4,3.46,4.3,WH,4);
  ln(s,1.18,1.72,0.79,0,WS,{lw:3});
  tx(s,1.95,5.14,3.11,0.9,[[SED.slice(0,84)]],{sz:11,f:LT,c:WH,ls:1.5});
  tx(s,1.95,4.69,2.21,0.375,[[LID]],{sz:12,f:PS,c:WH,cs:0.2,ls:1.5});
  tx(s,1.95,3.55,3.11,0.9,[[SED.slice(0,84)]],{sz:11,f:LT,c:WH,ls:1.5});
  tx(s,1.95,3.2,2.21,0.375,[[LID]],{sz:12,f:PS,c:WH,cs:0.2,ls:1.5});
  iconDisc(s,0.99,3.48,0.69,WH,OR,'ship');
  iconDisc(s,0.99,5.02,0.69,WH,OR,'plane');
  ln(s,4.17,6.01,0,2.03,LG,{lw:1.5,rot:90});
  tx(s,1.4,5.82,0.488,2.4,[['www.'],['carlistic',{cs:1.2}],['.com']],{sz:17,f:PP,c:LG,al:'r',nw:1,vt:'vert270',rot:90});
  socRing(s,5.44,6.84,0.37,0.56,LG,true);
  sh(s,'rect',0,2.01,0.226,3.22,{f:OR});
  dots(s,11.59,0.62,0.25,0.28,0.073,OR,1);
}

/** Slide 27 — SWOT — Weakness (layout: 27_Custom Layout) */
function slide27(s) {
  sh(s,'roundRect',4.83,-1.38,6.06,10.96,{f:BK,a:30,rot:270,fh:1,r:0.487});
  letter(s,'W',3.07,2.6,4.79,3.54,WH,4);
  tx(s,10.65,1.59,1.76,0.406,[['SWOT']],{sz:16,f:PS,c:OR,cs:1.2,al:'r',ls:1.2});
  tx(s,8.39,5.71,3.11,0.9,[[SED.slice(0,84)]],{sz:11,f:LT,c:WH,al:'r',ls:1.5});
  tx(s,9.29,5.26,2.21,0.375,[[LID]],{sz:12,f:PS,c:WH,cs:0.2,al:'r',ls:1.5});
  tx(s,8.39,4.12,3.11,0.9,[[SED.slice(0,84)]],{sz:11,f:LT,c:WH,al:'r',ls:1.5});
  tx(s,9.29,3.77,2.21,0.375,[[LID]],{sz:12,f:PS,c:WH,cs:0.2,al:'r',ls:1.5});
  iconDisc(s,11.77,4.05,0.69,WH,OR,'ship');
  iconDisc(s,11.77,5.59,0.69,WH,OR,'plane');
  tx(s,8.24,2.3,4.22,1.2,[['Weaknes'],['  Carlistic Analysis Slide',{c:WH}]],{sz:28,f:PS,c:OR,cs:1.2,b:1,al:'r',ls:1.2});
  ln(s,11.54,2.21,0.79,0,WS,{lw:3});
  ln(s,8.06,-0.466,0,2.03,LG,{lw:1.5,rot:90});
  tx(s,5.29,-0.66,0.488,2.4,[['www.'],['carlistic',{cs:1.2}],['.com']],{sz:17,f:PP,c:LG,al:'r',nw:1,vt:'vert270',rot:90});
  socRing(s,9.32,0.362,0.37,0.56,LG,true);
  sh(s,'rect',13.11,2.6,0.226,3.22,{f:OR});
  dots(s,0.74,0.47,0.25,0.28,0.073,OR);
}

/** Slide 28 — SWOT — Opportunity (layout: 28_Custom Layout) */
function slide28(s) {
  sh(s,'roundRect',2.44,-1.97,6.07,10.96,{f:BK,a:30,rot:90,r:0.487});
  letter(s,'O',5.55,1.6,4.15,4.25,WH,4);
  tx(s,0.9,1.7,5.28,1.2,[['Oppurtinity',{f:PP}],['  '],['Carlistic',{c:OR,br:1}],['Analyis Slide']],{sz:28,f:PS,c:WH,cs:1.2,b:1,ls:1.2});
  tx(s,0.9,1.05,1.22,0.406,[['SWOT']],{sz:16,f:PS,c:OR,cs:1.2,ls:1.2});
  ln(s,4.17,6.01,0,2.03,LG,{lw:1.5,rot:90});
  tx(s,1.4,5.82,0.488,2.4,[['www.'],['carlistic',{cs:1.2}],['.com']],{sz:17,f:PP,c:LG,al:'r',nw:1,vt:'vert270',rot:90});
  socRing(s,5.44,6.84,0.37,0.56,LG,true);
  ln(s,0.99,1.58,0.79,0,WS,{lw:3});
  tx(s,1.95,5,3.11,0.9,[[SED.slice(0,84)]],{sz:11,f:LT,c:WH,ls:1.5});
  tx(s,1.95,4.55,2.21,0.375,[[LID]],{sz:12,f:PS,c:WH,cs:0.2,ls:1.5});
  tx(s,1.95,3.41,3.11,0.9,[[SED.slice(0,84)]],{sz:11,f:LT,c:WH,ls:1.5});
  tx(s,1.95,3.06,2.21,0.375,[[LID]],{sz:12,f:PS,c:WH,cs:0.2,ls:1.5});
  iconDisc(s,0.99,3.34,0.69,WH,OR,'ship');
  iconDisc(s,0.99,4.88,0.69,WH,OR,'plane');
  sh(s,'rect',0,2.01,0.226,3.22,{f:OR});
  dots(s,11.51,0.62,0.25,0.28,0.073,OR,1);
}

/** Slide 29 — SWOT — Threat (layout: 29_Custom Layout) */
function slide29(s) {
  sh(s,'rect',0,0,13.33,7.5,{f:WH,fh:1});
  sh(s,'roundRect',4.83,-1.38,6.06,10.96,{f:BK,a:30,rot:270,fh:1,r:0.487});
  tx(s,7.94,2.28,4.37,1.77,[['Threat',{c:OR}],[' cleaning analysis slide',{br:1}],[' ']],{sz:28,f:PS,c:WH,cs:1.2,al:'r',ls:1.2});
  tx(s,10.55,1.59,1.76,0.406,[['SWOT']],{sz:16,f:PS,c:OR,cs:1.2,al:'r',ls:1.2});
  tx(s,8.29,5.71,3.11,0.9,[[SED.slice(0,84)]],{sz:11,f:LT,c:WH,al:'r',ls:1.5});
  tx(s,9.19,5.26,2.21,0.375,[[LID]],{sz:12,f:PS,c:WH,cs:0.2,al:'r',ls:1.5});
  tx(s,8.29,4.12,3.11,0.9,[[SED.slice(0,84)]],{sz:11,f:LT,c:WH,al:'r',ls:1.5});
  tx(s,9.19,3.77,2.21,0.375,[[LID]],{sz:12,f:PS,c:WH,cs:0.2,al:'r',ls:1.5});
  iconDisc(s,11.67,4.05,0.69,WH,OR,'ship');
  iconDisc(s,11.67,5.59,0.69,WH,OR,'plane');
  ln(s,11.44,2.21,0.79,0,WS,{lw:3});
  ln(s,8.06,-0.466,0,2.03,LG,{lw:1.5,rot:90});
  tx(s,5.29,-0.66,0.488,2.4,[['www.'],['carlistic',{cs:1.2}],['.com']],{sz:17,f:PP,c:LG,al:'r',nw:1,vt:'vert270',rot:90});
  socRing(s,9.32,0.362,0.37,0.56,LG,true);
  sh(s,'rect',13.11,2.6,0.226,3.22,{f:OR});
  dots(s,0.74,0.47,0.25,0.28,0.073,OR);
  letter(s,'T',4.5,1.79,3.44,4.58,WH,4);
}

/** Slide 30 — Testimonial (layout: 30_Custom Layout) */
function slide30(s) {
  tx(s,0.46,1.83,5.46,0.6,[['Revie From Our Client']],{sz:28,f:PS,c:WH,cs:1.2,al:'c',ls:1.1});
  tx(s,2.12,1.35,2.15,0.418,[['Testimonial']],{f:PM,c:WH,cs:2,al:'c',ls:1.1});
  tx(s,0.81,3.62,4.55,1.11,[['“ Vocibus mediocrem ex vis. Novum errem pertinaci ethereum persequeris ea quote, any nam mnesarch voluptatibus “']],{sz:14,f:LT,c:WH,al:'c',ls:1.5});
  icon(s,2.33,3.07,0.376,YL,'star');
  icon(s,2.71,3.07,0.376,YL,'star');
  icon(s,3.08,3.07,0.376,YL,'star');
  icon(s,3.46,3.07,0.376,YL,'star');
  icon(s,3.83,3.07,0.376,YL,'star');
  tx(s,2.42,4.97,1.99,0.56,[['Jhon Smith  ']],{sz:20,f:PM,c:WH,cs:1,ls:1.5});
  tx(s,2.52,5.47,1.78,0.364,[['Product Designer']],{sz:12,f:LT,c:WH,cs:1,ls:1.5});
  ln(s,2.42,2.59,0.79,0,OR,{lw:3});
  ln(s,3.17,2.59,0.79,0,WS,{lw:3});
}

/** Slide 31 — Our apps — iPhone mockup (layout: 31_Custom Layout) */
function slide31(s) {
  sh(s,'rect',0,0,13.33,7.5,{f:BK});
  sh(s,'rtTriangle',0.235,-0.235,7.5,7.97,{f:OR,rot:270,fh:1,fv:1});
  sh(s,'rtTriangle',11.35,5.79,1.98,1.71,{f:OR,rot:180,fv:1});
  grad(s,3.73,3.48,2.7,2.73,'v',[WH,100,0],[OR,0,86],LG);
  tx(s,7.75,1.6,3.91,1.92,[['Our Apps We’re Always Ready To Help']],{sz:36,f:PS,c:WH});
  tx(s,7.75,1.27,2.56,0.303,[['Device Mockup Iphone']],{sz:12,f:PP,c:WH,cs:1.2});
  tx(s,7.75,3.83,3.91,1.45,[[LIP],['felis.']],{sz:11,f:LT,c:WH,ls:1.5});
  sh(s,'roundRect',7.91,5.81,1.72,0.482,{f:OR,l:NV,r:0.08});
  tx(s,7.91,5.88,1.72,0.352,[['REGISTER NOW']],{sz:11,f:PP,c:WH,cs:1.2,al:'c',ls:1.5});
  ph(s,4.2,4.72,1.72,1.06);
  ln(s,6.76,6.23,0,1.65,LG,{lw:1.5,rot:90});
  tx(s,4.58,6.03,0.397,1.96,[['www.'],['carlistic',{cs:1.2}],['.com']],{sz:17,f:PP,c:LG,al:'r',nw:1,vt:'vert270',rot:90});
  socRing(s,7.86,6.85,0.301,0.458,LG,true);
  dots(s,0.74,0.47,0.25,0.28,0.073,'222A35');
  dots(s,11.05,0.63,0.25,0.28,0.073,GY,1);
  ph(s,3.6,1.27,2.96,5.06);
  ph(s,0.96,1.73,2.43,4.14);
}

/** Slide 32 — Our apps — tablet mockup (layout: 32_Custom Layout) */
function slide32(s) {
  sh(s,'rect',0,0,3.72,7.5,{f:OR});
  sh(s,'roundRect',7.95,3.21,4.69,1.59,{f:WH,l:WS,r:0.127,sd:{type:"outer",blur:35,offset:0,angle:0,color:"A6A6A6",opacity:0.4}});
  tx(s,5.61,1.37,5.77,1.31,[['Our Apps '],['We’re Always Ready To Help',{c:BK}]],{sz:36,f:PS,c:OR});
  tx(s,5.61,0.94,2.56,0.51,[['Device Mockup Tablet']],{sz:12,f:PP,cs:1.2});
  tx(s,9.38,5.77,3.11,0.9,[[SED.slice(0,84)]],{sz:11,f:LT,c:GY,ls:1.5});
  tx(s,9.38,5.32,2.21,0.375,[[LID]],{sz:12,f:PS,cs:0.2,ls:1.5});
  tx(s,9.38,3.68,3.11,0.9,[[SED.slice(0,84)]],{sz:11,f:LT,c:GY,ls:1.5});
  tx(s,9.38,3.33,2.21,0.375,[[LID]],{sz:12,f:PS,cs:0.2,ls:1.5});
  iconDisc(s,8.36,3.66,0.69,BK,OR,'ship');
  iconDisc(s,8.36,5.74,0.69,BK,OR,'plane');
  dots(s,11.09,0.63,0.25,0.28,0.073,LN,1);
  ln(s,5.71,2.81,2.34,0,OR,{lw:3});
  ph(s,1.13,0.99,3.91,5.77);
}

/** Slide 33 — iMac mockup (layout: 34_Custom Layout) */
function slide33(s) {
  sh(s,'rect',0,0,13.33,7.5,{f:BK,fh:1,fv:1});
  sh(s,'rtTriangle',5.6,-0.235,7.5,7.97,{f:OR,rot:270});
  sh(s,'rtTriangle',0,0,1.98,1.71,{f:OR,rot:180,fh:1});
  dots(s,11.05,0.63,0.25,0.28,0.073,GY,1);
  ph(s,6.67,1.13,5.81,5.29);
  tx(s,2.04,1.46,4.45,1.64,[['For Over 12 Years '],[' '],['In The Bussiness'],[' '],['Cargo & Logistic']],{sz:28,f:PS,c:WH,cs:1.2,ls:1.1});
  tx(s,2.56,3.48,1.55,0.421,[['Lorem Ipsum']],{sz:14,f:PS,c:OR,cs:0.2,ls:1.5});
  tx(s,2.16,4,3.06,0.97,[[SED.slice(0,84)]],{sz:12,f:LT,c:WH,ls:1.5});
  icon(s,2.19,3.6,0.216,OR,'check');
  sh(s,'rect',2.08,1.03,2.28,0.315,{f:OR});
  tx(s,2.08,1.03,2.28,0.315,[['Device Mockup Imac']],{sz:12,f:PP,c:WH,cs:1.2,ls:1.1,bg:OR});
  tx(s,2.56,5.16,1.55,0.421,[['Lorem Ipsum']],{sz:14,f:PS,c:OR,cs:0.2,ls:1.5});
  tx(s,2.16,5.68,3.33,0.97,[[SED.slice(0,84)]],{sz:12,f:LT,c:WH,ls:1.5});
  icon(s,2.19,5.29,0.216,OR,'check');
  ln(s,0.54,5.58,0,1.16,WH);
  socRing(s,0.382,3.09,0.321,0.54,WH,false);
  tx(s,9.15,5.59,0.488,2.64,[[WWW],['\t ']],{sz:17,f:PP,c:WH,cs:1.2,al:'r',nw:1,vt:'vert270',rot:90});
}

/** Slide 34 — Quote — Jhon Dhoe Smith (layout: 15_Custom Layout) */
function slide34(s) {
  sh(s,'rect',0,0,13.33,7.5,{f:BK,fv:1});
  sh(s,'rtTriangle',-0.168,0.168,7.5,7.16,{f:OR,rot:90,fh:1});
  tx(s,7.54,1.64,5.18,2.79,[['Sweet As The '],['Moment When ',{c:OR}],['The Cargo Went '],['\'Pop\'',{c:OR}]],{sz:40,f:PS,c:WH});
  tx(s,5.69,-0.078,1.71,2.36,[['“']],{sz:200,f:PS,c:NB,al:'c',va:'m'});
  tx(s,7.54,4.66,2.57,0.404,[['Jhon Dhoe Smith']],{f:PS,c:WH,cs:1,nw:1});
  tx(s,11.02,4.68,1.71,2.36,[['“']],{sz:200,f:PS,c:NB,al:'c',va:'m',rot:180});
  tx(s,7.54,5.12,1.68,0.337,[['Ceo & Founder']],{sz:14,f:LT,c:OR,cs:1,nw:1});
}

/** Slide 35 — Keep in touch (layout: 2_Custom Layout) */
function slide35(s) {
  sh(s,'rect',5.58,0,7.75,7.5,{f:OR});
  tx(s,8.15,2.78,4.4,0.96,[['1234 Melborn CA 1234, ',{br:1}],['AUSTRALIA']],{sz:16,f:LT,c:WH,lp:33});
  tx(s,8.15,3.85,4.4,0.496,[['0123 456 78']],{sz:16,f:LT,c:WH,lp:33});
  icon(s,7.45,3.92,0.458,WH,'phone');
  icon(s,7.46,2.89,0.57,WH,'marker');
  tx(s,7.34,1.16,3.76,1.45,[['Keep In Touch With Us']],{sz:28,f:PS,c:WH,cs:1.5,b:1,ls:1.5});
  tx(s,8.3,4.81,2.8,0.496,[['jhondhoe@carlistic.com']],{sz:16,f:LT,c:WH,lp:33});
  tx(s,8.33,5.72,2.4,0.485,[[WWW]],{sz:16,f:LT,c:WH,lp:32});
  icon(s,7.62,4.89,0.44,WH,'envelope');
  icon(s,7.62,5.71,0.493,WH,'world');
  sh(s,'rect',4.25,0.59,8.3,6.61,{l:BK,lw:2.75});
  sh(s,'rtTriangle',11.06,0.305,1.72,2.09,{f:BK,l:BK,rot:180});
  dots(s,11.54,0.58,0.18,0.2,0.052,'FF9D71',1);
}

/** Slide 36 — Thanks (layout: Custom Layout) */
function slide36(s) {
  grad(s,0,0,13.35,7.5,'v',[BK,69,0],[BK,36,100],WH);
  sh(s,'rect',2.84,1.93,7.35,3.62,{l:OR,lw:2.75});
  sh(s,'rect',3.02,2.1,7.06,3.29,{f:WH});
  tx(s,3.89,2.97,5.43,1.72,[['Thank’s']],{sz:96,f:PS,c:OR,al:'c',nw:1});
  sh(s,'rtTriangle',9.35,1.84,1.15,1.15,{f:OR,rot:180});
  ln(s,6.83,5.86,0,2.03,WH,{lw:1.25,rot:90});
  tx(s,4.07,5.73,0.488,2.27,[[WWW]],{sz:17,f:WK,c:WH,al:'r',nw:1,vt:'vert270',rot:90});
  socRing(s,8.1,6.68,0.37,0.56,WH,true);
}

// ------------------------------------------------------------ build ---

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';   // 13.333in x 7.5in
pptx.author = 'CarListic';
pptx.title  = 'CarListic — Cargo & Logistic Presentation';

const SLIDES = [
  slide01, slide02, slide03, slide04, slide05, slide06,
  slide07, slide08, slide09, slide10, slide11, slide12,
  slide13, slide14, slide15, slide16, slide17, slide18,
  slide19, slide20, slide21, slide22, slide23, slide24,
  slide25, slide26, slide27, slide28, slide29, slide30,
  slide31, slide32, slide33, slide34, slide35, slide36,
];

SLIDES.forEach(build => build(pptx.addSlide()));

pptx.writeFile({ fileName: path.join(__dirname, '15201362-80f2-4342-b241-d545e6c87145_grok_final.pptx') })
  .then(f => console.log('wrote ' + f))
  .catch(e => { console.error(e); process.exit(1); });
