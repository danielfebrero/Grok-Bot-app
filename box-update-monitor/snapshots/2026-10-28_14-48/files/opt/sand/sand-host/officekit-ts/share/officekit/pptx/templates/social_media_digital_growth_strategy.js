/*
 * "Social Media Presentation" — 25 slides, 13.333 x 7.5 in, rebuilt with pptxgenjs.
 *
 * Layout of this file:
 *   palette & fonts  ->  drawing helpers (gradients, sparkles, dot fields,
 *   zig-zags, rings) ->  one builder function per slide  ->  build & save.
 *
 * Run:  node 14eb7976-4e75-4313-a08a-e15c1c78e8ab_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette ---
const C = {
  purple: 'A45BE1',   // accent1
  blue:   '219AE3',   // accent2
  sand:   'FFDCA0',   // accent3
  red:    'F66C6C',   // accent4
  indigo: '5B63FB',   // accent5
  mint:   '45E8AD',   // accent6
  black:  '000000',
  white:  'FFFFFF',
  body:   '808080',   // grey paragraph copy
  nav:    '595959',   // grey nav labels
  peach:  'FCC8AA',   // faded arrow heads
  panel:  'F2F2F2'    // pale grey panels
};
const FT = { head: 'Sora Medium', body: 'Kanit' };

// ------------------------------------------------------------ lorem ipsum ---
const L = {};
L.labore  = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore';
L.magna   = L.labore + ' et dolore magna';
L.aliqua  = L.magna + ' aliqua';
L.minim   = L.aliqua + '. Ut enim ad minim';
L.veniam  = L.minim + ' veniam';
L.quis    = L.veniam + ', quis';
L.ullamco = L.quis + ' nostrud exercitation ullamco';
L.aliquip = L.ullamco + ' laboris nisi ut aliquip';
L.consequat = L.aliquip + ' ex ea commodo consequat';
L.tempo   = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempo';
L.elit    = 'Lorem ipsum dolor sit amet, consectetur elit adipiscing sed do eiusmod tempor';
L.eiusmod = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod';

// --------------------------------------------------------- colour helpers ---
const hex2rgb = (h) => [0, 2, 4].map((i) => parseInt(h.substr(i, 2), 16));
const rgb2hex = (v) => v.map((n) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0').toUpperCase()).join('');

/** linear blend of two hex colours, t in [0,1] */
function mix(a, b, t) {
  const A = hex2rgb(a), B = hex2rgb(b);
  return rgb2hex([0, 1, 2].map((i) => A[i] + (B[i] - A[i]) * t));
}

/** colour at t of a ramp whose stops sit at p0..p1 percent (defaults 0..100) */
function gstop(from, to, t, p0, p1) {
  const a = (p0 || 0) / 100, b = (p1 === undefined ? 100 : p1) / 100;
  return mix(from, to, Math.max(0, Math.min(1, (t - a) / (b - a || 1))));
}

// ----------------------------------------------------------- shape helpers ---
const NOLINE = { type: 'none' };
const NOFILL = { type: 'none' };
const SEAM = 0.02;   // extra inch of band overlap that hides anti-aliasing seams

function rect(s, x, y, w, h, color, transparency) {
  s.addShape('rect', { x, y, w, h, fill: { color, transparency }, line: NOLINE });
}

/**
 * Linear-gradient rectangle painted as a stack of thin bands, since pptxgenjs
 * only emits solid fills.  dir: 'h' L->R, 'hRev' R->L, 'v' top->bottom,
 * 'vUp' bottom->top.  `r` insets the band ends so the stack fits a rounded box.
 */
function gradRect(s, o) {
  const horiz = o.dir === 'h' || o.dir === 'hRev';
  const rev = o.dir === 'hRev' || o.dir === 'vUp';
  const span = horiz ? o.w : o.h;
  const n = Math.max(24, Math.min(96, Math.round(span * 20)));
  const step = span / n;
  for (let i = 0; i < n; i++) {
    const t = (i + 0.5) / n;
    const c = gstop(o.from, o.to, rev ? 1 - t : t, o.p0, o.p1);
    const off = i * step;
    let ins = 0;
    if (o.r) {
      const d = Math.min(off, span - off - step);
      if (d < o.r) ins = o.r - Math.sqrt(Math.max(0, o.r * o.r - (o.r - d) * (o.r - d)));
    }
    const insA = o.rSide === 'bottom' ? 0 : ins;   // inset at the low-coordinate edge
    const insB = o.rSide === 'top' ? 0 : ins;      // inset at the high-coordinate edge
    if (horiz) rect(s, o.x + off, o.y + insA, step + SEAM, o.h - insA - insB, c);
    else rect(s, o.x + insA, o.y + off, o.w - insA - insB, step + SEAM, c);
  }
}

/** rounded rectangle whose fill fades between two colours */
function gradRoundRect(s, o) {
  s.addShape('roundRect', { x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: o.r,
    fill: { color: mix(o.from, o.to, 0.5) }, line: NOLINE });
  gradRect(s, Object.assign({}, o, { r: o.r }));
}

/**
 * Photo card of slide 14: a horizontal colour ramp darkened towards the
 * bottom by a black scrim.  Painted as an opaque cell grid so the two ramps
 * composite cleanly instead of stacking translucent bands.
 */
function scrimCard(s, o) {
  const cols = 26, rows = 24, cw = o.w / cols, ch = o.h / rows;
  for (let r = 0; r < rows; r++) {
    const dark = 0.12 + 0.63 * ((r + 0.5) / rows);        // scrim strength, top -> bottom
    const d = Math.min(r * ch, o.h - (r + 1) * ch);
    const ins = d < o.r ? o.r - Math.sqrt(Math.max(0, o.r * o.r - (o.r - d) * (o.r - d))) : 0;
    for (let c = 0; c < cols; c++) {
      const x0 = Math.max(o.x + c * cw, o.x + ins), x1 = Math.min(o.x + (c + 1) * cw + SEAM, o.x + o.w - ins);
      if (x1 <= x0) continue;
      const base = mix(o.from, o.to, (c + 0.5) / cols);
      rect(s, x0, o.y + r * ch, x1 - x0, ch + SEAM, mix(base, C.black, dark));
    }
  }
}

/** ellipse with a vertical colour ramp, built from clipped horizontal bands */
function gradEllipseV(s, o) {
  const n = Math.max(56, Math.round(o.h * 26)), rx = o.w / 2, cx = o.x + rx;
  for (let i = 0; i < n; i++) {
    const yy = Math.abs((i + 0.5) / n - 0.5) * 2;
    const hw = rx * Math.sqrt(Math.max(0, 1 - yy * yy));
    rect(s, cx - hw, o.y + (i / n) * o.h, hw * 2, o.h / n + SEAM, mix(o.from, o.to, (i + 0.5) / n));
  }
}

/**
 * Thin outlined circle.  The source strokes these with a gradient, but a
 * gradient line renders as its first stop, so a solid stroke is a match.
 */
function ring(s, x, y, d, color, width) {
  s.addShape('ellipse', { x, y, w: d, h: d, fill: NOFILL, line: { color, width: width || 1.25 } });
}

/** thin horizontal rule */
function hRule(s, x, y, w, color, width) {
  const th = (width || 2) / 72;
  rect(s, x, y - th / 2, w, th, color);
}

/** full-height vertical guide (slides 7 / 16 / 21) */
function vRule(s, x, color, width) {
  const th = (width || 2) / 72;
  rect(s, x - th / 2, 0, th, 7.5, color);
}

/** four-point sparkle — the deck's signature "shine" motif */
function sparkle(s, x, y, d, color) {
  s.addShape('star4', { x, y, w: d, h: d, adjustment: 8, fill: { color }, line: NOLINE });
}
/** row of sparkles sharing a size, stepped by `gap` */
function sparkleRow(s, x, y, d, gap, colors) {
  colors.forEach((c, i) => sparkle(s, x + i * gap, y, d, c));
}
/** vertical stack of sparkles */
function sparkleCol(s, x, y, d, gap, colors) {
  colors.forEach((c, i) => sparkle(s, x, y + i * gap, d, c));
}
/** three sparkles fading between two colours */
function sparkleTrio(s, x, y, d, gap, from, to) {
  sparkleRow(s, x, y, d, gap, [from, mix(from, to, 0.5), to]);
}

/** zig-zag rule with 4 teeth */
function zigzag(s, x, y, w, h, color, width) {
  const pts = [];
  for (let i = 0; i <= 7; i++) pts.push({ x: (w * i) / 7, y: i % 2 ? 0 : h });
  s.addShape('custGeom', { x, y, w, h, points: pts, fill: NOFILL, line: { color, width: width || 1.75 } });
}
/** the recurring stacked pair of zig-zags */
function zigzagPair(s, x, y, w, top, bottom, width) {
  const h = w * 0.1077, gap = w * 0.1756;
  zigzag(s, x, y, w, h, top, width);
  zigzag(s, x, y + gap, w, h, bottom, width);
}

/**
 * rows x cols dot matrix with a colour ramp.
 * axis 'x' ramps left->right, 'y' ramps top->bottom.
 * whiteCols makes the trailing columns solid white (dots over a gradient panel).
 */
function dotGrid(s, o) {
  const axis = o.axis || 'x';
  const n = ((axis === 'x' ? o.cols : o.rows) - 1) || 1;
  for (let r = 0; r < o.rows; r++) {
    for (let c = 0; c < o.cols; c++) {
      const t = (axis === 'x' ? c : r) / n;
      const white = o.whiteCols !== undefined && c >= o.cols - o.whiteCols;
      s.addShape('ellipse', { x: o.x + c * o.dx, y: o.y + r * o.dy, w: o.d, h: o.d,
        fill: { color: white ? C.white : (o.to ? mix(o.from, o.to, t) : o.from) }, line: NOLINE });
    }
  }
}

/**
 * Six-dot staircase tucked into a slide corner; `corner` (tl|tr|bl|br) picks
 * which half of the 3x3 grid is kept.
 */
function dotStair(s, o) {
  for (let r = 0; r < 3; r++) {
    for (let c = 0; c < 3; c++) {
      const keep = { tl: c + r <= 2, tr: c >= r, bl: c <= r, br: c + r >= 2 }[o.corner];
      if (!keep) continue;
      s.addShape('ellipse', { x: o.x + c * 0.2925, y: o.y + r * 0.2925, w: 0.05, h: 0.05,
        fill: { color: mix(o.from, o.to, (c + r) / 4) }, line: NOLINE });
    }
  }
}

/** small solid triangle pointer; dir is one of r,l,u,d */
function tri(s, x, y, w, h, color, dir) {
  s.addShape('triangle', { x, y, w, h, rotate: { r: 90, l: 270, u: 0, d: 180 }[dir || 'r'],
    fill: { color }, line: NOLINE });
}

/** bullet marker: pale halo ring + solid core */
function haloDot(s, x, y, halo, core) {
  s.addShape('ellipse', { x, y, w: 0.21, h: 0.207,
    fill: { color: halo, transparency: halo === C.white ? 0 : 50 }, line: NOLINE });
  s.addShape('ellipse', { x: x + 0.045, y: y + 0.044, w: 0.119, h: 0.118,
    fill: { color: core }, line: NOLINE });
}

/** stand-in for the one raster image the original deck embedded */
function imageBox(s, x, y, w, h, label) {
  rect(s, x, y, w, h, 'DCDEE5');
  s.addText(label || '[image]', { x, y: y + h / 2 - 0.2, w, h: 0.4, align: 'center',
    fontSize: 12, fontFace: FT.body, color: '8A8F9C' });
}

// ------------------------------------------------------------ text helper ---
/**
 * txt(slide, text, x, y, w, h, opts)
 * opts: size, color, font ('head'|'body'), bold, align, lh (line-spacing
 * multiple), bullet, nowrap (single-line labels that must not break)
 */
function txt(s, text, x, y, w, h, o) {
  o = o || {};
  const props = {
    x, y, w, h,
    fontSize: o.size || 11,
    fontFace: o.font === 'head' ? FT.head : FT.body,
    color: o.color || C.black,
    align: o.align || 'left',
    valign: 'top',
    isTextBox: true,
    wrap: o.nowrap !== true
  };
  if (o.bold) props.bold = true;
  if (o.lh) props.lineSpacingMultiple = o.lh;
  if (o.bullet) props.bullet = o.bullet;
  s.addText(text, props);
}

/** grey paragraph copy — 11 pt, 1.5 line spacing, on every slide */
function body(s, text, x, y, w, h, align) {
  txt(s, text, x, y, w, h, { size: 11, color: C.body, lh: 1.5, align });
}

/** section heading in the display face */
function head(s, text, x, y, w, h, size, color, align) {
  txt(s, text, x, y, w, h, { size: size || 36, font: 'head', color: color || C.black, align });
}

/**
 * Bulleted list, one paragraph per line.  The bullet has to be repeated on
 * every item — pptxgenjs only honours a top-level `bullet` on the first one.
 */
function bulletList(s, items, x, y, w, h, o) {
  const per = { breakLine: true, bullet: o.bullet, lineSpacingMultiple: o.lh };
  txt(s, items.map((t) => ({ text: t, options: per })), x, y, w, h, o);
}

/** browser-chrome navigation bar of the cover / closing slides */
function navBar(s, items) {
  items.forEach((it) => {
    txt(s, it.t, it.x, it.b ? 0.349 : 0.365, it.w, it.b ? 0.303 : 0.269,
      { size: it.b ? 12 : 10, bold: it.b, color: it.c || C.nav, align: 'center' });
  });
}

// =============================================================== slides ====
const slides = [];

// 1 — cover ------------------------------------------------------------------
slides.push((s) => {
  gradRect(s, { x: 4.286, y: 6.5, w: 9.048, h: 1.0, from: C.red, to: C.sand, dir: 'h' });
  ring(s, 10.722, -0.891, 1.891, C.mint);
  navBar(s, [
    { t: 'Home', x: 1.304, w: 0.679, b: 1, c: C.purple },
    { t: 'About us', x: 3.286, w: 0.999 },
    { t: 'Service', x: 5.353, w: 0.663 },
    { t: 'Portfolio', x: 7.320, w: 0.999 },
    { t: 'Others', x: 9.355, w: 0.619 },
    { t: 'Search', x: 11.277, w: 0.752, b: 1, c: C.indigo }
  ]);
  hRule(s, 6.574, 4.133, 5.937, C.indigo, 2);
  txt(s, [{ text: 'Social ', options: { color: C.black } },
          { text: 'Media', options: { color: C.purple } },
          { text: ' Presentation ', options: { color: C.black } }],
    6.459, 1.786, 6.289, 2.322, { size: 66, font: 'head' });
  txt(s, 'Creative Strategy for Digital Growth', 6.459, 4.356, 4.719, 0.37, { size: 16, color: C.mint });
  body(s, L.magna, 6.459, 4.726, 4.473, 0.622);
  sparkle(s, 4.959, 1.972, 0.827, C.purple);
  dotGrid(s, { x: 4.871, y: 3.143, cols: 5, rows: 10, dx: 0.242, dy: 0.245, d: 0.033, from: C.indigo, to: C.mint });
  dotGrid(s, { x: 7.618, y: 6.741, cols: 10, rows: 3, dx: 0.244, dy: 0.242, d: 0.033, from: C.white });
  tri(s, 0.203, 0.439, 0.298, 0.121, C.purple, 'r');
  tri(s, 0.324, 0.439, 0.298, 0.121, C.blue, 'r');
  tri(s, 5.134, 6.931, 0.339, 0.138, C.white, 'r');
  tri(s, 5.272, 6.931, 0.339, 0.138, C.white, 'r');
  zigzagPair(s, 11.594, 6.793, 1.462, C.white, C.white, 2);
});

// 2 — agenda -----------------------------------------------------------------
const AGENDA = [
  { n: '01', c: C.purple, t: 'Introduction',           x: 4.766, y: 2.263, tw: 1.446 },
  { n: '04', c: C.indigo, t: 'Strategy Framework',     x: 8.950, y: 2.263, tw: 2.222 },
  { n: '02', c: C.indigo, t: 'Social Media Landscape', x: 4.766, y: 3.754, tw: 2.660 },
  { n: '05', c: C.purple, t: 'Performance Metrics',    x: 8.950, y: 3.754, tw: 2.274 },
  { n: '03', c: C.purple, t: 'Key Trends',             x: 4.766, y: 5.245, tw: 1.311 },
  { n: '06', c: C.indigo, t: 'Conclusion',             x: 8.950, y: 5.245, tw: 1.352 }
];
slides.push((s) => {
  s.addShape('ellipse', { x: -4.505, y: 1.984, w: 8.854, h: 3.531, rotate: 90,
    fill: { color: 'F4F4F4' }, line: NOLINE });
  dotGrid(s, { x: 10.723, y: 0.890, cols: 5, rows: 2, dx: 0.412, dy: 0.413, d: 0.070, axis: 'y', from: C.sand, to: C.red });
  ring(s, -1.404, 1.146, 5.412, C.sand, 1.5);
  tri(s, 1.099, 1.063, 0.406, 0.166, C.red, 'r');
  tri(s, 1.099, 6.474, 0.406, 0.166, C.sand, 'l');
  head(s, 'Agenda', 5.536, 0.778, 2.262, 0.707);
  AGENDA.forEach((a) => {
    txt(s, a.n, a.x, a.y, 0.756, 0.572, { size: 28, font: 'head', color: a.c });
    txt(s, a.t, a.x + 0.817, a.y + 0.101, a.tw, 0.37, { size: 16, nowrap: true });
    body(s, L.labore, a.x, a.y + 0.536, 3.764, 0.622);
  });
  tri(s, 3.805, 3.769, 0.406, 0.166, C.peach, 'd');
  sparkle(s, 3.661, 0.805, 0.653, C.blue);
  sparkle(s, 8.934, 0.805, 0.653, C.purple);
  zigzagPair(s, 2.495, 6.806, 1.413, C.purple, C.blue, 2);
  ring(s, 11.809, 6.726, 1.891, C.mint);
});

// 3 — introduction -----------------------------------------------------------
slides.push((s) => {
  head(s, 'Introduction', 1.098, 1.602, 3.567, 0.707);
  txt(s, 'Social media is the core of digital communication. This presentation showcases strategies, insights, and tools to build a strong online presence.',
    1.098, 2.689, 5.134, 0.909, { size: 16, color: C.purple });
  body(s, L.consequat + '. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum',
    1.098, 3.965, 5.667, 1.178);
  body(s, L.veniam, 1.098, 5.276, 5.667, 0.622);
  dotGrid(s, { x: 7.526, y: 2.890, cols: 1, rows: 5, dx: 0, dy: 0.412, d: 0.070, axis: 'y', from: C.sand, to: C.red });
  dotGrid(s, { x: 12.996, y: 2.890, cols: 1, rows: 5, dx: 0, dy: 0.412, d: 0.070, axis: 'y', from: C.red, to: C.sand });
  gradRect(s, { x: 8.441, y: 0, w: 3.711, h: 7.5, from: C.indigo, to: C.mint, dir: 'vUp' });
  dotGrid(s, { x: 9.180, y: 0.400, cols: 10, rows: 3, dx: 0.245, dy: 0.242, d: 0.033, from: C.white });
  dotGrid(s, { x: 9.180, y: 6.583, cols: 10, rows: 3, dx: 0.245, dy: 0.242, d: 0.033, from: C.white });
  sparkle(s, 0.468, 6.499, 0.653, C.sand);
  zigzagPair(s, 6.169, 6.641, 1.413, C.purple, C.blue, 2);
  ring(s, -1.084, -1.084, 2.168, C.sand);
  tri(s, 6.590, 0.576, 0.406, 0.166, C.purple, 'r');
  tri(s, 6.755, 0.576, 0.406, 0.166, C.blue, 'r');
});

// 4 — social media growth (bar chart) ----------------------------------------
slides.push((s) => {
  head(s, 'Social Media Growth', 7.383, 1.503, 3.567, 1.313);
  gradRoundRect(s, { x: 7.419, y: 4.464, w: 4.456, h: 0.633, r: 0.316, from: C.indigo, to: C.mint, dir: 'h' });
  txt(s, 'Over 5 billion active users worldwide', 7.653, 4.595, 3.951, 0.37, { size: 16, color: C.white });
  body(s, L.ullamco.replace('ullamco', 'ullamc'), 7.383, 3.286, 4.492, 0.9);
  body(s, L.magna, 7.383, 5.375, 4.492, 0.622);
  s.addShape('roundRect', { x: 1.497, y: 1.503, w: 4.494, h: 4.494, rectRadius: 0.293,
    fill: { color: C.white }, line: { color: 'F2F2F2', width: 0.75 } });
  s.addChart('bar', [{ name: 'Series 1', labels: ['2021', '2022', '2023', '2024', '2025'],
    values: [4.3, 2.5, 3.5, 7, 4.9] }], {
    x: 1.984, y: 2.070, w: 3.521, h: 3.360,
    barDir: 'col', barGapWidthPct: 182,
    chartColors: [C.purple, C.blue, C.sand, C.red, C.mint],
    showLegend: false, showValue: false,
    valAxisMaxVal: 8, valAxisMinVal: 0, valAxisMajorUnit: 1,
    catAxisLabelColor: C.nav, valAxisLabelColor: C.nav,
    catAxisLabelFontSize: 10, valAxisLabelFontSize: 10,
    catAxisLabelFontFace: FT.body, valAxisLabelFontFace: FT.body,
    valGridLine: { style: 'solid', size: 0.5, color: 'D9D9D9' },
    catGridLine: { style: 'none' }
  });
  tri(s, 12.551, 0.528, 0.406, 0.166, C.blue, 'r');
  dotGrid(s, { x: 12.363, y: 5.997, cols: 3, rows: 5, dx: 0.271, dy: 0.271, d: 0.046, axis: 'y', from: C.sand, to: C.red });
  zigzagPair(s, 5.960, 0.611, 1.413, C.purple, C.blue, 2);
  ring(s, 4.282, 6.555, 1.891, C.blue);
  sparkle(s, 4.902, 7.174, 0.653, C.purple);
});

// 5 — why social media matters ------------------------------------------------
slides.push((s) => {
  gradRect(s, { x: 6.667, y: 0, w: 3.111, h: 2.773, from: C.white, to: C.panel, dir: 'vUp', p1: 85 });
  gradRect(s, { x: 10.0, y: 4.727, w: 3.111, h: 2.773, from: C.white, to: C.panel, dir: 'v', p1: 85 });
  head(s, 'Why Social Media Matters', 1.349, 1.425, 3.935, 1.313);
  bulletList(s, ['Brand Visibility', 'Customer Engagement', 'Lead Generation', 'Real-Time Feedback'],
    1.349, 3.123, 3.058, 1.667,
    { size: 16, color: C.purple, lh: 1.5, bullet: { characterCode: '2022', indent: 22.5 } });
  body(s, L.quis, 1.349, 5.175, 3.969, 0.9);
  tri(s, 4.558, 1.695, 0.406, 0.166, C.purple, 'r');
  tri(s, 4.724, 1.695, 0.406, 0.166, C.blue, 'r');
  dotGrid(s, { x: 10.695, y: 0.487, cols: 5, rows: 3, dx: 0.413, dy: 0.412, d: 0.070, axis: 'y', from: C.indigo, to: C.mint });
  dotGrid(s, { x: 7.362, y: 6.118, cols: 5, rows: 3, dx: 0.413, dy: 0.413, d: 0.070, axis: 'y', from: C.sand, to: C.red });
  sparkle(s, 0.208, 6.768, 0.552, C.red);
  sparkle(s, 0.760, 6.768, 0.552, C.sand);
  zigzagPair(s, 0.208, 0.222, 1.105, C.indigo, C.mint, 2);
});

// 6 — popular platforms --------------------------------------------------------
const PLATFORMS = [
  { name: 'Facebook',    x: 5.687, y: 1.766, ix: 4.856, iy: 1.964, tw: 1.208 },
  { name: 'LinkedIn',    x: 9.984, y: 1.766, ix: 9.153, iy: 2.007, tw: 1.087 },
  { name: 'Instagram',   x: 5.687, y: 3.254, ix: 4.826, iy: 3.450, tw: 1.208 },
  { name: 'X (Twitter)', x: 9.984, y: 3.254, ix: 9.153, iy: 3.450, tw: 1.230 },
  { name: 'TikTok',      x: 5.687, y: 4.741, ix: 4.864, iy: 4.937, tw: 0.881 },
  { name: 'YouTube',     x: 9.984, y: 4.741, ix: 9.153, iy: 5.026, tw: 1.100 }
];
slides.push((s) => {
  rect(s, 0, 1.266, 4.975, 4.968, '666666');
  s.addShape('round2SameRect', { x: 6.286, y: -0.813, w: 4.968, h: 9.127, rotate: 270,
    rectRadius: 0.526, fill: { color: C.white }, line: NOLINE });
  head(s, 'Popular Platforms', 0.768, 3.094, 2.671, 1.313, 36, C.white, 'center');
  PLATFORMS.forEach((p) => {
    // stylised platform badge: purple/blue disc with a white glyph cut-out
    s.addShape('ellipse', { x: p.ix, y: p.iy, w: 0.6, h: 0.6,
      fill: { color: mix(C.purple, C.blue, 0.45) }, line: NOLINE });
    s.addShape('ellipse', { x: p.ix + 0.19, y: p.iy + 0.19, w: 0.22, h: 0.22,
      fill: { color: C.white }, line: NOLINE });
    txt(s, p.name, p.x, p.y, p.tw, 0.37, { size: 16, nowrap: true });
    body(s, L.tempo, p.x, p.y + 0.37, 2.962, 0.622);
  });
  dotGrid(s, { x: 0.138, y: 6.364, cols: 10, rows: 5, dx: 0.244, dy: 0.242, d: 0.033, from: C.indigo, to: C.mint });
  dotGrid(s, { x: 10.963, y: 0.132, cols: 10, rows: 5, dx: 0.244, dy: 0.242, d: 0.033, from: C.red, to: C.sand });
  sparkleTrio(s, 0.421, 0.357, 0.552, 0.557, C.red, C.sand);
  sparkleTrio(s, 11.246, 6.588, 0.552, 0.557, C.indigo, C.mint);
  zigzagPair(s, 6.114, 0.476, 1.105, C.indigo, C.mint);
  zigzagPair(s, 6.114, 6.708, 1.105, C.red, C.sand);
});

// 7 — key trends 2025 -----------------------------------------------------------
slides.push((s) => {
  head(s, 'Key Trends 2025', 1.287, 1.373, 4.093, 1.447, 40);
  body(s, L.veniam, 1.287, 3.142, 4.093, 0.903);
  ['Short-form Video Domination', 'AI-powered Personalization', 'Social Commerce Growth', 'AR/VR Experiences']
    .forEach((t, i) => {
      const y = 4.503 + i * 0.4295;
      haloDot(s, 1.562, y + 0.065, C.indigo, C.indigo);
      txt(s, t, 1.883, y, 3.497, 0.337, { size: 14, color: C.purple });
    });
  vRule(s, 12.688, C.sand, 2);
  tri(s, 12.485, 1.429, 0.406, 0.166, C.red, 'u');
  tri(s, 12.485, 3.667, 0.406, 0.166, C.peach, 'd');
  tri(s, 12.485, 5.906, 0.406, 0.166, C.sand, 'u');
  dotStair(s, { x: 0.447, y: 0.447, corner: 'tl', from: C.blue, to: C.purple });
  dotStair(s, { x: 5.585, y: 6.418, corner: 'br', from: C.indigo, to: C.mint });
  sparkle(s, 0.447, 6.736, 0.317, C.purple);
  sparkle(s, 0.764, 6.736, 0.317, C.blue);
  sparkle(s, 5.585, 0.447, 0.317, C.indigo);
  sparkle(s, 5.902, 0.447, 0.317, C.mint);
});

// 8 — target audience (doughnut chart) -------------------------------------------
const PERSONA = [
  { label: 'Age',          value: '20+', y: 1.316, lw: 1.764 },
  { label: 'Gender',       value: 'Any', y: 3.247, lw: 1.720 },
  { label: 'Income Level', value: '7K$', y: 5.174, lw: 1.745 }
];
slides.push((s) => {
  gradRect(s, { x: 0, y: 0, w: 2.667, h: 7.5, from: C.purple, to: C.red, dir: 'hRev' });
  s.addShape('roundRect', { x: 5.682, y: 0.856, w: 1.969, h: 5.789, rectRadius: 0.384,
    fill: { color: C.white }, line: { color: 'F0F0F2', width: 0.75 } });
  PERSONA.forEach((p, i) => {
    txt(s, p.label, 5.886 - (p.lw - 1.562) / 2, p.y, p.lw, 0.337, { size: 14, color: C.body, align: 'center' });
    txt(s, p.value, 5.886, p.y + 0.37, 1.562, 0.64, { size: 32, font: 'head', align: 'center' });
    if (i < 2) hRule(s, 5.682, 2.786 + i * 1.928, 1.969, i ? C.purple : C.indigo, 1.75);
  });
  s.addShape('ellipse', { x: 8.574, y: 4.279, w: 2.421, h: 2.421, fill: { color: C.white }, line: NOLINE });
  s.addChart('doughnut', [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'],
    values: [0.25, 0.1, 0.25, 0.4] }], {
    x: 8.431, y: 4.084, w: 2.702, h: 2.816,
    holeSize: 50, firstSliceAng: 222,
    chartColors: ['D9D9D9', 'BFBFBF', '7F7F7F', mix(C.sand, C.red, 0.4)],
    showLegend: false, showValue: true, dataLabelFormatCode: '0%',
    dataLabelColor: C.white, dataLabelFontSize: 12, dataLabelFontBold: true, dataLabelFontFace: FT.head,
    showTitle: false
  });
  txt(s, 'Chart', 8.574, 5.375, 2.421, 0.28, { size: 11, font: 'head', color: C.nav, align: 'center' });
  head(s, 'Target Audience', 8.565, 0.856, 3.713, 1.447, 40);
  body(s, L.veniam, 8.565, 2.491, 3.991, 0.903);
  txt(s, 'Psychographic:', 8.564, 3.707, 1.847, 0.37, { size: 16, color: C.purple });
  dotGrid(s, { x: 11.660, y: 4.632, cols: 3, rows: 5, dx: 0.413, dy: 0.412, d: 0.070, axis: 'y', from: C.indigo, to: C.mint });
  tri(s, 12.104, 1.496, 0.406, 0.166, C.purple, 'r');
  tri(s, 12.270, 1.496, 0.406, 0.166, C.blue, 'r');
  sparkle(s, 6.340, 0.150, 0.555, C.sand);
  sparkle(s, 6.340, 6.794, 0.555, C.red);
});

// 9 — buyer persona ---------------------------------------------------------------
slides.push((s) => {
  gradRect(s, { x: 7.802, y: 0, w: 3.388, h: 7.5, from: C.red, to: C.sand, dir: 'vUp' });
  head(s, 'Buyer Persona Example', 1.349, 1.425, 3.935, 1.313);
  bulletList(s, ['Name: Emma, 25, Lifestyle Enthusiast', 'Platforms: Instagram, TikTok', 'Interests: Fashion, Travel, Food'],
    1.349, 3.140, 4.309, 1.634,
    { size: 16, color: C.purple, lh: 2, bullet: { characterCode: '27A2', indent: 22.5 } });
  body(s, L.quis, 1.349, 5.175, 4.309, 0.9);
  ring(s, 6.752, 1.006, 5.487, C.indigo, 1.5);
  tri(s, 0.509, 0.630, 0.406, 0.166, C.purple, 'r');
  tri(s, 12.418, 6.705, 0.406, 0.166, C.red, 'l');
  dotStair(s, { x: 12.069, y: 0.447, corner: 'tr', from: C.sand, to: C.red });
  dotStair(s, { x: 0.630, y: 6.418, corner: 'bl', from: C.purple, to: C.blue });
  zigzagPair(s, 6.114, 0.476, 1.105, C.indigo, C.mint, 2);
  zigzagPair(s, 6.114, 6.708, 1.105, C.mint, C.indigo, 2);
});

// 10 — content strategy framework ---------------------------------------------------
const FRAMEWORK = [
  { t: 'Awareness',  x: 1.616, y: 2.529, tw: 1.294 },
  { t: 'Engagement', x: 6.844, y: 2.529, tw: 1.503 },
  { t: 'Conversion', x: 1.616, y: 4.257, tw: 1.354 },
  { t: 'Retention',  x: 6.844, y: 4.257, tw: 1.174 }
];
slides.push((s) => {
  hRule(s, 0, 3.226, 13.333, C.sand, 1.75);
  hRule(s, 0, 4.953, 13.333, C.purple, 1.75);
  FRAMEWORK.forEach((f) => {
    s.addShape('roundRect', { x: f.x, y: f.y, w: 4.873, h: 1.392, rectRadius: 0.219,
      fill: { color: C.white }, line: { color: 'F2F2F2', width: 0.75 } });
    txt(s, f.t, f.x + 1.649, f.y + 0.2, f.tw, 0.37, { size: 16, nowrap: true });
    body(s, L.tempo, f.x + 1.649, f.y + 0.571, 2.962, 0.622);
  });
  head(s, 'Content Strategy Framework', 4.297, 0.603, 4.733, 1.313, 36, C.black, 'center');
  body(s, L.consequat, 2.344, 6.264, 8.644, 0.622, 'center');
  dotGrid(s, { x: 1.032, y: 0.758, cols: 10, rows: 5, dx: 0.245, dy: 0.243, d: 0.033, from: C.indigo, to: C.mint });
  dotGrid(s, { x: 10.066, y: 0.758, cols: 10, rows: 5, dx: 0.244, dy: 0.243, d: 0.033, from: C.indigo, to: C.mint });
  sparkle(s, 0.530, 2.948, 0.555, C.sand);
  sparkle(s, 0.530, 4.676, 0.555, C.purple);
  sparkle(s, 12.248, 2.948, 0.555, C.red);
  sparkle(s, 12.248, 4.676, 0.555, C.blue);
  zigzagPair(s, 0.620, 6.418, 1.105, C.indigo, C.mint);
  zigzagPair(s, 11.609, 6.418, 1.105, C.indigo, C.mint);
});

// 11 — content types -----------------------------------------------------------------
const CTYPES = [
  { t: 'Educational',   x: 7.890, y: 3.750, dx: 7.569, c: C.purple },
  { t: 'Entertaining',  x: 7.890, y: 4.245, dx: 7.569, c: C.purple },
  { t: 'Inspirational', x: 10.236, y: 3.750, dx: 9.915, c: C.blue },
  { t: 'Promotional',   x: 10.236, y: 4.245, dx: 9.915, c: C.blue }
];
slides.push((s) => {
  s.addShape('round2SameRect', { x: 0.278, y: 0, w: 6.0, h: 6.291, flipV: true,
    rectRadius: 0.456, fill: { color: mix(C.sand, C.red, 0.5) }, line: NOLINE });
  gradRect(s, { x: 0.278, y: 0, w: 6.0, h: 6.291, from: C.sand, to: C.red, dir: 'h', p0: 23, r: 0.456, rSide: 'bottom' });
  gradRect(s, { x: 0, y: 0, w: 2.880, h: 7.5, from: C.white, to: C.panel, dir: 'h', p1: 85 });
  gradRoundRect(s, { x: 7.253, y: 3.451, w: 4.814, h: 1.429, r: 0.2, from: C.purple, to: C.blue, dir: 'h' });
  head(s, 'Content Types', 7.253, 1.499, 4.586, 0.774, 40);
  body(s, L.minim, 7.253, 2.573, 5.105, 0.622);
  CTYPES.forEach((c) => {
    haloDot(s, c.dx, c.y + 0.065, C.white, c.c);
    txt(s, c.t, c.x, c.y, 1.515, 0.337, { size: 14, font: 'head', color: C.white });
  });
  body(s, L.ullamco, 7.253, 5.100, 5.105, 0.9);
  dotGrid(s, { x: 0.785, y: 6.585, cols: 5, rows: 3, dx: 0.287, dy: 0.287, d: 0.049, axis: 'y', from: C.indigo, to: C.mint });
  tri(s, 6.381, 6.814, 0.406, 0.166, C.mint, 'r');
  tri(s, 6.547, 6.814, 0.406, 0.166, C.indigo, 'r');
  sparkleTrio(s, 11.399, 0.273, 0.552, 0.558, C.red, C.sand);
  zigzagPair(s, 11.962, 6.896, 1.105, C.indigo, C.mint);
});

// 12 — content calendar ----------------------------------------------------------------
const CAL_FILL = {                                  // date -> highlight colour
  3: C.sand, 4: C.sand, 5: C.sand,
  11: C.red, 12: C.red, 13: C.red,
  19: C.indigo, 20: C.indigo, 21: C.indigo,
  24: C.mint, 25: C.mint, 26: C.mint
};
slides.push((s) => {
  gradRoundRect(s, { x: 1.092, y: 3.905, w: 4.144, h: 0.633, r: 0.316, from: C.indigo, to: C.mint, dir: 'h' });
  txt(s, 'Weekly posting schedule mockup', 1.327, 4.037, 3.675, 0.37, { size: 16, color: C.white, align: 'center' });
  body(s, L.aliquip, 1.092, 2.695, 5.097, 0.9);
  head(s, 'Content Calendar Example', 1.092, 1.072, 4.764, 1.313);
  dotGrid(s, { x: 3.913, y: 1.997, cols: 5, rows: 1, dx: 0.412, dy: 0, d: 0.070, from: C.red });
  // month banner + month-number badge
  s.addShape('roundRect', { x: 7.281, y: 0.856, w: 4.306, h: 0.690, rectRadius: 0.143,
    fill: { color: C.purple }, line: NOLINE });
  txt(s, 'November 2025', 7.281, 0.986, 4.306, 0.44, { size: 18, color: C.white, align: 'center' });
  s.addShape('roundRect', { x: 11.673, y: 0.856, w: 0.804, h: 0.690, rectRadius: 0.143,
    fill: { color: C.purple }, line: NOLINE });
  txt(s, '11', 11.673, 0.986, 0.804, 0.44, { size: 18, color: C.white, align: 'center' });
  ['S', 'M', 'T', 'W', 'T', 'F', 'S'].forEach((d, i) => {
    const x = 7.281 + i * 0.7555;
    s.addShape('roundRect', { x, y: 1.748, w: 0.664, h: 0.372, rectRadius: 0.1,
      fill: { color: C.white }, line: { color: 'F0F0F0', width: 0.5 } });
    txt(s, d, x, 1.792, 0.664, 0.3, { size: 16, align: 'center', color: i === 0 ? C.red : C.black });
  });
  for (let i = 0; i < 42; i++) {                    // Nov 2025 starts on a Saturday
    const col = i % 7, row = Math.floor(i / 7), day = i - 5;
    const x = 7.281 + col * 0.7555, y = 2.210 + row * 0.7543;
    const fill = CAL_FILL[day] || C.white;
    s.addShape('roundRect', { x, y, w: 0.664, h: 0.664, rectRadius: 0.14,
      fill: { color: fill }, line: { color: fill === C.white ? 'F0F0F0' : fill, width: 0.5 } });
    if (day >= 1 && day <= 30) {
      txt(s, String(day), x, y + 0.18, 0.664, 0.35,
        { size: 16, align: 'center', color: fill !== C.white ? C.white : (col === 0 ? C.red : C.black) });
    }
  }
  ring(s, -1.084, -1.084, 2.168, C.sand);
  tri(s, 0.260, 6.856, 0.406, 0.166, C.mint, 'r');
  tri(s, 0.426, 6.856, 0.406, 0.166, C.indigo, 'r');
  sparkleTrio(s, 9.046, 0.149, 0.552, 0.557, C.purple, C.blue);
  zigzagPair(s, 9.327, 6.923, 1.105, C.purple, C.blue);
});

// 13 — engagement tactics -----------------------------------------------------------------
slides.push((s) => {
  s.addShape('roundRect', { x: 1.238, y: 3.689, w: 4.127, h: 2.811, rectRadius: 0.2,
    fill: { color: C.white }, line: { color: 'F2F2F2', width: 0.75 } });
  bulletList(s, ['Polls & Quizzes', 'User-Generated Content', 'Giveaways & Challenges', 'Interactive Stories'],
    1.762, 4.008, 3.080, 2.172,
    { size: 16, color: C.purple, lh: 2, bullet: { characterCode: '2713', indent: 22.5 } });
  head(s, 'Engagement Tactics', 6.352, 1.058, 5.397, 0.707, 36, C.black, 'center');
  body(s, L.aliquip.replace('tempor', 'tem'), 6.352, 5.542, 5.397, 0.9, 'center');
  tri(s, 0.394, 0.630, 0.406, 0.166, C.blue, 'r');
  tri(s, 12.533, 6.705, 0.406, 0.166, C.red, 'l');
  dotStair(s, { x: 12.276, y: 0.447, corner: 'tr', from: C.sand, to: C.red });
  dotStair(s, { x: 0.421, y: 6.418, corner: 'bl', from: C.purple, to: C.blue });
  zigzagPair(s, 6.114, 0.447, 1.105, C.indigo, C.mint);
  zigzagPair(s, 6.114, 6.740, 1.105, C.indigo, C.mint);
  sparkleCol(s, 0.321, 2.916, 0.552, 0.558, [C.purple, mix(C.purple, C.blue, 0.5), C.blue]);
  sparkleCol(s, 12.460, 2.916, 0.552, 0.558, [C.red, mix(C.red, C.sand, 0.5), C.sand]);
});

// 14 — paid advertising --------------------------------------------------------------------
const ADS = [
  { t: 'Facebook Ads',  x: 0.626, from: C.sand,   to: C.red,  tx: 1.727, tw: 1.643 },
  { t: 'Instagram Ads', x: 4.745, from: C.purple, to: C.blue, tx: 5.844, tw: 1.645 },
  { t: 'TikTok Ads',    x: 8.864, from: C.indigo, to: C.mint, tx: 10.128, tw: 1.315 }
];
slides.push((s) => {
  gradRect(s, { x: 0, y: 5.403, w: 13.333, h: 2.097, from: C.white, to: C.panel, dir: 'vUp', p1: 85 });
  ADS.forEach((a) => {
    gradRoundRect(s, { x: a.x, y: 3.368, w: 3.843, h: 3.542, r: 0.289, from: a.from, to: a.to, dir: 'h' });
    scrimCard(s, { x: a.x, y: 3.202, w: 3.843, h: 2.823, r: 0.31, from: a.from, to: a.to });
    txt(s, a.t, a.tx, 5.517, a.tw, 0.37, { size: 16, color: C.white, align: 'center', nowrap: true });
    body(s, L.eiusmod, a.x + 0.441, 6.156, 2.962, 0.622, 'center');
  });
  head(s, 'Paid Advertising', 4.439, 0.603, 4.450, 0.707, 36, C.black, 'center');
  txt(s, 'Maximizing reach with targeted ads', 4.100, 1.661, 5.134, 0.37,
    { size: 16, color: C.purple, align: 'center' });
  body(s, L.consequat, 2.344, 2.031, 8.644, 0.622, 'center');
  dotGrid(s, { x: 0.626, y: 0.588, cols: 5, rows: 3, dx: 0.311, dy: 0.310, d: 0.053, axis: 'y', from: C.sand, to: C.red });
  dotGrid(s, { x: 11.414, y: 0.588, cols: 5, rows: 3, dx: 0.310, dy: 0.310, d: 0.053, axis: 'y', from: C.indigo, to: C.mint });
  sparkle(s, 0.721, 1.955, 0.552, C.red);
  sparkle(s, 1.273, 1.955, 0.552, C.sand);
  sparkle(s, 11.508, 1.955, 0.552, C.indigo);
  sparkle(s, 12.060, 1.955, 0.552, C.mint);
  zigzagPair(s, 2.627, 0.768, 1.105, C.purple, C.blue);
  zigzagPair(s, 9.599, 0.768, 1.105, C.purple, C.blue);
});

// 15 — influencer marketing ------------------------------------------------------------------
slides.push((s) => {
  gradRect(s, { x: 0, y: 4.721, w: 1.0, h: 1.585, from: C.indigo, to: C.mint, dir: 'vUp', p1: 73 });
  gradRect(s, { x: 12.387, y: 1.278, w: 0.946, h: 2.345, from: C.sand, to: C.red, dir: 'vUp', p1: 73 });
  head(s, 'Influencer Marketing', 1.159, 1.113, 2.891, 1.313);
  txt(s, 'Leveraging creators to amplify your brand', 1.159, 2.757, 5.117, 0.37, { size: 16, color: C.purple });
  body(s, L.aliqua + '. Ut enim adm', 1.159, 3.165, 5.117, 0.622);
  txt(s, 'Micro vs Macro Influencers', 7.057, 4.807, 5.117, 0.37, { size: 16, color: C.purple });
  body(s, L.aliquip, 7.057, 5.215, 5.117, 0.9);
  dotGrid(s, { x: 4.801, y: 1.433, cols: 5, rows: 3, dx: 0.310, dy: 0.310, d: 0.053, axis: 'y', from: C.purple, to: C.red });
  ring(s, -0.349, -0.349, 1.433, C.sand);
  sparkle(s, -0.009, -0.009, 0.753, C.sand);
  ring(s, 12.238, 6.414, 1.433, C.mint);
  sparkle(s, 12.578, 6.754, 0.753, C.mint);
  [1.962, 2.368, 2.773].forEach((y) => tri(s, 12.657, y, 0.406, 0.166, C.white, 'l'));
  [5.229, 5.632].forEach((y) => tri(s, 0.270, y, 0.406, 0.166, C.white, 'r'));
  zigzagPair(s, 3.165, 6.899, 1.105, C.mint, C.indigo);
  zigzagPair(s, 9.063, 0.301, 1.105, C.red, C.sand);
});

// 16 — KPIs & metrics --------------------------------------------------------------------------
slides.push((s) => {
  vRule(s, 1.700, C.sand, 2);
  vRule(s, 4.769, C.sand, 2);
  tri(s, 5.000, 0.489, 0.406, 0.166, C.red, 'u');
  tri(s, 1.931, 6.845, 0.406, 0.166, C.sand, 'u');
  head(s, 'KPIs & Metrics', 7.825, 1.425, 3.979, 0.707);
  bulletList(s, ['Engagement Rate', 'Reach & Impressions', 'Click-through Rate', 'Conversion Rate'],
    7.825, 2.567, 3.979, 2.172,
    { size: 16, color: C.purple, lh: 2, bullet: { characterCode: '2022', indent: 22.5 } });
  body(s, L.quis, 7.825, 5.175, 3.979, 0.9);
  dotGrid(s, { x: 0.300, y: 2.890, cols: 1, rows: 5, dx: 0, dy: 0.412, d: 0.070, axis: 'y', from: C.sand, to: C.red });
  dotGrid(s, { x: 12.963, y: 2.890, cols: 1, rows: 5, dx: 0, dy: 0.412, d: 0.070, axis: 'y', from: C.indigo, to: C.mint });
  zigzagPair(s, 11.928, 0.301, 1.105, C.mint, C.indigo);
  zigzagPair(s, 11.928, 6.885, 1.105, C.mint, C.indigo);
  sparkle(s, 7.006, 0.360, 0.490, C.purple);
  sparkle(s, 7.496, 0.360, 0.490, C.blue);
  sparkle(s, 7.006, 6.650, 0.490, C.purple);
  sparkle(s, 7.496, 6.650, 0.490, C.blue);
});

// 17 — analytics dashboard ------------------------------------------------------------------------
slides.push((s) => {
  gradRect(s, { x: 11.014, y: 0, w: 2.319, h: 7.5, from: C.indigo, to: C.mint, dir: 'vUp' });
  imageBox(s, 6.453, 1.466, 5.493, 4.567, '[image] desktop mockup');
  tri(s, 0.314, 6.705, 0.406, 0.166, C.sand, 'l');
  tri(s, 12.614, 0.630, 0.406, 0.166, C.white, 'r');
  dotStair(s, { x: 0.434, y: 0.447, corner: 'tl', from: C.indigo, to: C.mint });
  dotStair(s, { x: 12.265, y: 6.418, corner: 'br', from: C.white, to: C.white });
  head(s, 'Analytics Dashboard Example', 1.387, 1.486, 3.184, 1.717, 32);
  body(s, L.elit + ' incididu ut labore et dolore magna aliqua', 1.387, 3.494, 3.184, 0.9);
  txt(s, 'Mockup of social media analytics report', 1.387, 4.700, 3.184, 0.64, { size: 16, color: C.purple });
  body(s, L.elit, 1.387, 5.392, 3.184, 0.622);
  s.addShape('roundRect', { x: 5.267, y: 3.869, w: 2.768, h: 2.828, rectRadius: 0.343,
    fill: { color: C.white }, line: { color: 'F2F2F2', width: 0.75 } });
  // little bar-chart glyph on the "Analysis" card
  [[0.00, 0.18, 0.16, 0.42, C.purple], [0.19, 0.00, 0.16, 0.60, C.purple],
   [0.38, 0.28, 0.16, 0.32, C.blue],   [0.38, 0.10, 0.16, 0.16, C.blue]]
    .forEach((b) => rect(s, 6.351 + b[0] * 0.6, 4.337 + b[1] * 0.6, b[2] * 0.6, b[3] * 0.6, b[4]));
  txt(s, 'Analysis', 5.716, 5.336, 1.871, 0.37, { size: 16, font: 'head', align: 'center' });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.', 5.459, 5.759, 2.384, 0.62, 'center');
  sparkleTrio(s, 3.600, 0.488, 0.552, 0.558, C.purple, C.blue);
  sparkleTrio(s, 8.366, 6.438, 0.552, 0.557, C.red, C.sand);
  zigzagPair(s, 8.647, 0.608, 1.105, C.mint, C.indigo);
  zigzagPair(s, 2.427, 6.579, 1.105, C.mint, C.indigo);
});

// 18 — case study -----------------------------------------------------------------------------------
slides.push((s) => {
  head(s, 'Case Study', 7.995, 0.769, 3.567, 0.707, 36, C.black, 'center');
  gradRoundRect(s, { x: 7.533, y: 1.937, w: 4.492, h: 0.932, r: 0.466, from: C.indigo, to: C.mint, dir: 'h' });
  txt(s, 'Brand X increased engagement by 120% through video marketing',
    7.978, 2.083, 3.601, 0.64, { size: 16, color: C.white, align: 'center' });
  body(s, L.veniam + ', quis nostrud exercitati', 7.533, 3.224, 4.492, 0.9, 'center');
  body(s, L.ullamco + ' laboris nisi ut', 1.257, 5.831, 5.300, 0.9, 'center');
  dotGrid(s, { x: 0.594, y: 1.880, cols: 1, rows: 5, dx: 0, dy: 0.413, d: 0.070, axis: 'y', from: C.sand, to: C.red });
  dotGrid(s, { x: 12.670, y: 5.129, cols: 1, rows: 5, dx: 0, dy: 0.413, d: 0.070, axis: 'y', from: C.purple, to: C.blue });
  sparkleCol(s, 0.353, 5.448, 0.552, 0.557, [C.red, mix(C.red, C.sand, 0.5), C.sand]);
  sparkleCol(s, 12.428, 0.331, 0.552, 0.557, [C.purple, mix(C.purple, C.blue, 0.5), C.blue]);
  // the two zig-zags here are rotated a quarter-turn into vertical ribbons
  s.addShape('custGeom', { x: 6.986, y: 0.569, w: 1.105, h: 0.119, rotate: 90,
    points: zigPoints(1.105, 0.119), fill: NOFILL, line: { color: C.mint, width: 1.75 } });
  s.addShape('custGeom', { x: 6.986, y: 5.710, w: 1.105, h: 0.119, rotate: 90,
    points: zigPoints(1.105, 0.119), fill: NOFILL, line: { color: C.indigo, width: 1.75 } });
});

/** point list for a 4-tooth zig-zag, shared by the rotated variants */
function zigPoints(w, h) {
  const pts = [];
  for (let i = 0; i <= 7; i++) pts.push({ x: (w * i) / 7, y: i % 2 ? 0 : h });
  return pts;
}

// 19 — do's & don'ts ---------------------------------------------------------------------------------
slides.push((s) => {
  gradRect(s, { x: 10.013, y: 0, w: 3.320, h: 7.5, from: C.purple, to: C.sand, dir: 'h' });
  head(s, 'Social Media Do\u2019s & Don\u2019ts', 0.981, 1.249, 3.700, 1.313);
  txt(s, 'Do: Be authentic, consistent, responsive', 0.981, 3.095, 5.446, 0.37, { size: 16, color: C.purple });
  body(s, L.aliquip, 0.981, 3.491, 5.712, 0.9);
  txt(s, 'Don\u2019t: Ignore feedback, spam content, be insensitive', 0.981, 4.930, 5.446, 0.37, { size: 16, color: C.purple });
  body(s, L.aliquip, 0.981, 5.326, 5.712, 0.9);
  // dot pairs fading out over the gradient panel: the right 3 columns go white
  dotGrid(s, { x: 9.153, y: 0.389, cols: 5, rows: 2, dx: 0.413, dy: 0.413, d: 0.070, axis: 'y',
    from: C.sand, to: C.red, whiteCols: 3 });
  dotGrid(s, { x: 9.153, y: 6.628, cols: 5, rows: 2, dx: 0.413, dy: 0.413, d: 0.070, axis: 'y',
    from: C.indigo, to: C.mint, whiteCols: 3 });
  tri(s, 0.453, 0.548, 0.406, 0.166, C.red, 'l');
  tri(s, 0.453, 6.787, 0.406, 0.166, C.red, 'l');
  sparkleTrio(s, 5.986, 0.405, 0.451, 0.455, C.purple, C.blue);
  sparkleTrio(s, 5.986, 6.644, 0.451, 0.455, C.purple, C.blue);
});

// 20 — crisis management ------------------------------------------------------------------------------
slides.push((s) => {
  gradRect(s, { x: 0, y: 1.143, w: 4.180, h: 5.214, from: C.indigo, to: C.red, dir: 'h', p1: 85 });
  head(s, 'Crisis Management', 8.359, 1.197, 3.700, 1.313);
  txt(s, [{ text: 'Handling negative comments & ', options: { breakLine: true } },
          { text: 'PR issues effectively' }],
    8.359, 4.486, 3.967, 0.64, { size: 16, color: C.purple });
  body(s, L.ullamco + ' laboris nisi ut aliquip ex ea', 8.359, 5.126, 3.967, 1.178);
  dotGrid(s, { x: 8.575, y: 3.050, cols: 5, rows: 3, dx: 0.412, dy: 0.413, d: 0.070, axis: 'y', from: C.indigo, to: C.mint });
  ring(s, 1.210, 0.781, 5.939, C.sand, 1.5);
  tri(s, 3.977, 0.698, 0.406, 0.166, C.red, 'r');
  tri(s, 3.977, 6.637, 0.406, 0.166, C.sand, 'l');
  tri(s, 6.946, 3.769, 0.406, 0.166, C.peach, 'd');
  tri(s, 1.007, 3.769, 0.406, 0.166, C.peach, 'd');
  ring(s, 12.238, -0.349, 1.433, C.mint);
  sparkle(s, 12.578, -0.009, 0.753, C.mint);
  ring(s, 12.238, 6.414, 1.433, C.blue);
  sparkle(s, 12.578, 6.754, 0.753, C.purple);
  zigzagPair(s, 0.394, 0.415, 1.105, C.mint, C.indigo);
  zigzagPair(s, 0.394, 6.772, 1.105, C.blue, C.purple);
});

// 21 — social media tools ------------------------------------------------------------------------------
slides.push((s) => {
  head(s, 'Social Media Tools', 1.330, 1.373, 3.969, 1.447, 40);
  body(s, L.veniam, 1.330, 3.142, 3.969, 0.903);
  ['Buffer', 'Hootsuite', 'Canva', 'Meta Business Suite'].forEach((t, i) => {
    const y = 4.503 + i * 0.4295;
    haloDot(s, 1.605, y + 0.065, C.indigo, C.indigo);
    txt(s, t, 1.926, y, 3.497, 0.337, { size: 14, font: 'head', color: C.purple });
  });
  vRule(s, 7.663, C.indigo, 2);
  vRule(s, 11.321, C.sand, 2);
  dotStair(s, { x: 0.447, y: 0.447, corner: 'tl', from: C.sand, to: C.red });
  dotStair(s, { x: 0.447, y: 6.418, corner: 'bl', from: C.indigo, to: C.mint });
  zigzag(s, 9.041, 0.576, 1.105, 0.119, C.mint, 1.75);
  zigzag(s, 9.041, 6.804, 1.105, 0.119, C.indigo, 1.75);
  sparkle(s, 6.005, 0.319, 0.634, C.mint);
  sparkle(s, 12.450, 0.319, 0.634, C.red);
  sparkle(s, 6.005, 6.547, 0.634, C.indigo);
  sparkle(s, 12.450, 6.547, 0.634, C.sand);
});

// 22 — future of social media ---------------------------------------------------------------------------
slides.push((s) => {
  gradEllipseV(s, { x: -1.840, y: -0.678, w: 3.531, h: 8.854, from: C.purple, to: C.mint });
  gradEllipseV(s, { x: 11.643, y: -0.678, w: 3.531, h: 8.854, from: C.purple, to: C.mint });
  head(s, 'Future of Social Media', 4.817, 1.730, 3.700, 1.313, 36, C.black, 'center');
  txt(s, 'AI, Metaverse, and beyond', 4.776, 3.449, 3.782, 0.37, { size: 16, color: C.purple, align: 'center' });
  body(s, L.quis, 4.684, 3.845, 3.966, 0.9, 'center');
  body(s, L.veniam, 4.684, 4.870, 3.966, 0.9, 'center');
  dotGrid(s, { x: 5.806, y: 0.623, cols: 5, rows: 2, dx: 0.413, dy: 0.413, d: 0.070, axis: 'y', from: C.sand, to: C.red });
  dotGrid(s, { x: 5.806, y: 6.394, cols: 5, rows: 2, dx: 0.413, dy: 0.413, d: 0.070, axis: 'y', from: C.indigo, to: C.mint });
  [2.741, 9.612].forEach((x) => {
    sparkle(s, x, 0.400, 0.490, C.sand);
    sparkle(s, x + 0.490, 0.400, 0.490, C.red);
    sparkle(s, x, 6.609, 0.490, C.indigo);
    sparkle(s, x + 0.490, 6.609, 0.490, C.mint);
  });
});

// 23 — conclusion ----------------------------------------------------------------------------------------
slides.push((s) => {
  hRule(s, 3.483, 0.447, 3.360, C.red, 2);
  hRule(s, 3.483, 7.053, 3.360, C.mint, 2);
  head(s, 'Conclusion', 7.158, 1.786, 3.147, 0.707);
  txt(s, 'Social media is not just about presence, it\u2019s about connection and value creation',
    7.158, 2.873, 4.375, 0.64, { size: 16, color: C.purple });
  body(s, L.consequat + '. Duis aute irure dolor in reprehenderit', 7.158, 3.781, 5.115, 1.178);
  body(s, L.aliqua + ' ut enim', 7.158, 5.092, 5.115, 0.622);
  tri(s, 6.723, 2.056, 0.406, 0.166, C.purple, 'r');
  dotStair(s, { x: 12.265, y: 0.447, corner: 'tr', from: C.sand, to: C.red });
  dotStair(s, { x: 12.265, y: 6.418, corner: 'br', from: C.indigo, to: C.mint });
  sparkleTrio(s, 8.873, 0.447, 0.451, 0.455, C.red, C.sand);
  sparkleTrio(s, 8.873, 6.601, 0.451, 0.455, C.indigo, C.mint);
  zigzag(s, 10.745, 2.080, 1.105, 0.119, C.blue, 1.75);
});

// 24 — call to action -------------------------------------------------------------------------------------
slides.push((s) => {
  gradRect(s, { x: 9.622, y: 0, w: 3.711, h: 7.5, from: C.red, to: C.sand, dir: 'hRev' });
  head(s, 'Call to Action', 1.489, 1.714, 1.828, 1.178, 32);
  body(s, L.elit + ' incididu ut labore et dolore magna aliqua', 1.489, 3.267, 3.184, 0.9);
  txt(s, 'Let\u2019s shape the future of digital together!', 1.489, 4.473, 3.184, 0.64, { size: 16, color: C.purple });
  body(s, L.elit, 1.489, 5.164, 3.184, 0.622);
  dotGrid(s, { x: 3.780, y: 1.850, cols: 3, rows: 3, dx: 0.413, dy: 0.413, d: 0.070, axis: 'y', from: C.indigo, to: C.mint });
  tri(s, 0.453, 0.548, 0.406, 0.166, C.blue, 'l');
  tri(s, 0.453, 6.787, 0.406, 0.166, C.indigo, 'l');
  zigzagPair(s, 5.610, 6.713, 1.105, C.mint, C.indigo);
  sparkle(s, 4.638, 0.386, 0.490, C.indigo);
  sparkle(s, 5.128, 0.386, 0.490, C.mint);
});

// 25 — thank you --------------------------------------------------------------------------------------------
const CONTACT = [
  { icon: 'pin',   ix: 1.419, iy: 4.248, tx: 1.944, ty: 4.212, tw: 2.752, th: 0.471,
    lines: ['200 Broadway Av New Canberra', 'WA 5024 West Australia'] },
  { icon: 'phone', ix: 1.419, iy: 4.887, tx: 1.944, ty: 4.894, tw: 2.427, th: 0.286,
    lines: ['(+62) 8123 4567 790'] },
  { icon: 'mail',  ix: 5.218, iy: 4.295, tx: 5.793, ty: 4.305, tw: 1.886, th: 0.286,
    lines: ['socialmedia@email.com'] },
  { icon: 'globe', ix: 5.218, iy: 4.887, tx: 5.793, ty: 4.894, tw: 1.886, th: 0.286,
    lines: ['www.socialmedia.com'] }
];
slides.push((s) => {
  gradRect(s, { x: 0, y: 6.5, w: 9.048, h: 1.0, from: C.sand, to: C.red, dir: 'h' });
  navBar(s, [
    { t: 'Search', x: 1.304, w: 0.752, b: 1, c: C.indigo },
    { t: 'About us', x: 3.286, w: 0.857 },
    { t: 'Service', x: 5.353, w: 0.663 },
    { t: 'Portfolio', x: 7.193, w: 0.857 },
    { t: 'Others', x: 9.355, w: 0.619 },
    { t: 'Home', x: 11.277, w: 0.679, b: 1, c: C.purple }
  ]);
  txt(s, [{ text: 'Thank ', options: { color: C.black } },
          { text: 'You', options: { color: C.purple } }],
    1.304, 1.915, 6.289, 1.212, { size: 66, font: 'head' });
  hRule(s, 1.419, 3.304, 5.937, C.indigo, 2);
  txt(s, 'Contact Info ', 1.351, 3.612, 4.719, 0.37, { size: 16, color: C.indigo });
  const ICON = mix(C.mint, C.indigo, 0.5);
  CONTACT.forEach((c) => {
    if (c.icon === 'pin') {
      s.addShape('ellipse', { x: c.ix, y: c.iy, w: 0.30, h: 0.30, fill: NOFILL, line: { color: ICON, width: 2 } });
      s.addShape('ellipse', { x: c.ix + 0.10, y: c.iy + 0.10, w: 0.10, h: 0.10, fill: { color: ICON }, line: NOLINE });
      tri(s, c.ix + 0.075, c.iy + 0.26, 0.15, 0.12, ICON, 'd');
    } else if (c.icon === 'phone') {
      s.addShape('ellipse', { x: c.ix, y: c.iy, w: 0.30, h: 0.30, fill: NOFILL, line: { color: ICON, width: 2 } });
      s.addShape('ellipse', { x: c.ix + 0.09, y: c.iy + 0.09, w: 0.12, h: 0.12, fill: { color: ICON }, line: NOLINE });
    } else if (c.icon === 'mail') {
      rect(s, c.ix, c.iy, 0.40, 0.28, ICON);
      s.addShape('triangle', { x: c.ix + 0.05, y: c.iy + 0.03, w: 0.30, h: 0.17, rotate: 180,
        fill: { color: C.white }, line: NOLINE });
    } else {
      s.addShape('ellipse', { x: c.ix, y: c.iy, w: 0.30, h: 0.30, fill: { color: ICON }, line: NOLINE });
      s.addShape('ellipse', { x: c.ix + 0.09, y: c.iy, w: 0.12, h: 0.30, fill: NOFILL, line: { color: C.white, width: 0.75 } });
      rect(s, c.ix, c.iy + 0.135, 0.30, 0.03, C.white);
    }
    txt(s, c.lines.map((t) => ({ text: t, options: { breakLine: true } })),
      c.tx, c.ty, c.tw, c.th, { size: 11, color: C.body });
  });
  ring(s, 10.855, -0.292, 1.583, C.purple);
  tri(s, 0.442, 0.439, 0.298, 0.121, C.purple, 'r');
  tri(s, 0.564, 0.439, 0.298, 0.121, C.blue, 'r');
  tri(s, 7.723, 6.931, 0.339, 0.138, C.white, 'r');
  tri(s, 7.861, 6.931, 0.339, 0.138, C.white, 'r');
  sparkle(s, 7.576, 2.094, 0.827, C.blue);
  zigzagPair(s, 0.278, 6.793, 1.462, C.white, C.white, 2);
  dotGrid(s, { x: -0.527, y: 2.634, cols: 5, rows: 10, dx: 0.242, dy: 0.244, d: 0.033, from: C.indigo, to: C.mint });
  dotGrid(s, { x: 3.484, y: 6.741, cols: 10, rows: 3, dx: 0.244, dy: 0.242, d: 0.033, from: C.white });
});

// =============================================================== build ====
const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'WIDE';
pptx.theme = { headFontFace: FT.head, bodyFontFace: FT.body };
pptx.title = 'Social Media Presentation';

slides.forEach((build) => {
  const s = pptx.addSlide();
  s.background = { color: C.white };
  build(s);
});

pptx.writeFile({ fileName: path.join(__dirname, '14eb7976-4e75-4313-a08a-e15c1c78e8ab_grok_final.pptx') })
  .then((f) => console.log('wrote', f));
