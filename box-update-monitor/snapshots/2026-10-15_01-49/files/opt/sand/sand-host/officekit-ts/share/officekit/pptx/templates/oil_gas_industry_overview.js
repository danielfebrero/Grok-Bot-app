/**
 * "Oil and Gas" presentation template - rebuilt with pptxgenjs.
 *
 * Slide size 10 x 5.625 in (16:9). Dark deck, purple accent, Manrope headings
 * and DM Sans body copy. Photographs in the original are replaced by flat
 * "[image]" placeholder rectangles of the same position and size.
 *
 * Run:  node 0d519d00-5b38-4091-ad98-ffb9e4af02a7_grok_final.js
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const C = {
  purple: '7619FF',   // accent 1
  violet: '6600F9',   // icon badges / calendar highlight
  lilac: 'AC75FF',    // eyebrow text + footer tick
  white: 'FFFFFF',
  black: '000000',
  ink: '0C0C0C',      // plan cards
  grey: 'BFBFBF',     // body copy
  greyLt: 'D8D8D8',   // body copy on cards
  greyMid: 'A5A5A5',  // table rows
  rule: '262626',     // table-of-contents rules
  divider: '7F7F7F',  // faint vertical dividers
  axis: '595959',
  photo: 'F2F2F2',    // image placeholder fill
  photoInk: '2B2B2B', // "[image]" caption
  calDay: '6D7381',
  ok: '92D050',
  no: 'C00000',
};

const HEAD = 'Manrope';
const BODY = 'DM Sans';
const INSET = [5.4, 5.4, 2.7, 2.7];   // l, r, b, t in points (matches the source deck)
const SLIDE_W = 10;
const SLIDE_H = 5.625;

/* ------------------------------------------------------------ small helpers */

const hex = (c) => [parseInt(c.slice(0, 2), 16), parseInt(c.slice(2, 4), 16), parseInt(c.slice(4, 6), 16)];
const str = (rgb) => rgb.map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0').toUpperCase()).join('');
const mix = (a, b, t) => str(hex(a).map((v, i) => v + (hex(b)[i] - v) * t));
const dim = (a, k) => str(hex(a).map((v) => v * k));

/**
 * Colour of a DrawingML "accent -> transparent" gradient at position t (0..1),
 * flattened onto the backdrop it sits on. `hold` keeps the first slice at full
 * strength, `a0` is the opacity of the first stop.
 */
function fade(t, opt) {
  const o = Object.assign({ from: C.purple, bg: C.black, hold: 0, a0: 1 }, opt);
  const u = o.hold ? Math.max(0, (t - o.hold) / (1 - o.hold)) : t;
  return mix(dim(o.from, 1 - u), o.bg, 1 - o.a0 * (1 - u));
}

/** Paints a linear gradient as a run of thin solid strips. */
function ramp(s, o) {
  const across = o.dir === 'right' || o.dir === 'left';
  const n = o.steps || Math.max(12, Math.min(90, Math.round((across ? o.w : o.h) * 34)));
  const eps = 0.006;
  for (let i = 0; i < n; i++) {
    let t = (i + 0.5) / n;
    if (o.dir === 'left' || o.dir === 'up') t = 1 - t;
    const x = across ? o.x + (o.w * i) / n : o.x;
    const y = across ? o.y : o.y + (o.h * i) / n;
    const w = across ? o.w / n + eps : o.w;
    const h = across ? o.h : o.h / n + eps;
    s.addShape('rect', { x, y, w, h, fill: { color: o.color(t, x + w / 2, y + h / 2) }, line: { type: 'none' } });
  }
}

/** Solid rectangle (optionally translucent). */
function box(s, x, y, w, h, fill, opt) {
  const o = opt || {};
  s.addShape(o.shape || 'rect', {
    x, y, w, h,
    fill: fill === null ? { type: 'none' } : { color: fill, transparency: o.alpha },
    line: o.line || { type: 'none' },
    rotate: o.rotate, flipH: o.flipH, rectRadius: o.rectRadius,
  });
}

/** Text box. `text` may be a string or an array of paragraphs. */
function T(s, o) {
  const paras = Array.isArray(o.text) ? o.text : [o.text];
  const runs = paras.map((t, i) => ({
    text: t,
    options: {
      breakLine: i < paras.length - 1,
      bullet: o.bullet ? { indent: 10 } : false,
      paraSpaceBefore: i > 0 ? o.before || 0 : 0,
    },
  }));
  s.addText(runs, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    margin: INSET, valign: o.valign || 'top', align: o.align || 'left',
    fontFace: o.font || BODY, fontSize: o.size || 9, color: o.color || C.grey,
    bold: o.bold, italic: o.italic, lineSpacingMultiple: o.ls, rotate: o.rotate,
    charSpacing: o.cs, wrap: o.wrap !== false,
  });
}

/** Flat stand-in for a photograph. */
function photo(s, x, y, w, h, opt) {
  const o = opt || {};
  box(s, x, y, w, h, o.color || C.photo, { rotate: o.rotate, shape: o.round ? 'ellipse' : 'rect' });
  if (o.label !== false && !o.rotate && w > 0.62 && h > 0.3) {
    T(s, {
      x, y: y + h / 2 - 0.13, w, h: 0.26, text: '[image]', align: 'center', valign: 'middle',
      size: Math.max(6, Math.min(9, w * 3.5)), color: C.photoInk, font: BODY,
    });
  }
}

/** Purple disc used behind the small pictogram icons. */
function iconDot(s, x, y, d, opt) {
  const o = opt || {};
  box(s, x, y, d, d, o.color || C.violet, { shape: 'ellipse', alpha: o.alpha });
  T(s, { x, y: y + d / 2 - 0.11, w: d, h: 0.22, text: '\u25C6', align: 'center', valign: 'middle',
    size: d * 22, color: C.white, font: BODY });
}

/* ----------------------------------------------------- repeated deck chrome */

/** Footer + page number drawn by the slide master on every slide but the cover. */
function chrome(s, n) {
  T(s, { x: 0.546, y: 5.196, w: 1.297, h: 0.189, text: 'www.yoursite.com', size: 7, color: C.grey, font: HEAD });
  s.addShape('line', { x: 0.383, y: 5.29, w: 0.141, h: 0, line: { color: C.lilac, width: 1.5 } });
  T(s, { x: 8.88, y: 0.24, w: 0.737, h: 0.189, text: 'Page ' + n, size: 7, color: C.white, font: HEAD, align: 'right' });
}

/** Small violet "FUELING THE FUTURE!" kicker above every section title. */
function eyebrow(s, x, y, opt) {
  const o = opt || {};
  T(s, { x, y, w: o.w || 2.272, h: 0.227, text: 'FUELING THE FUTURE!', size: 9, font: HEAD,
    color: o.color || C.lilac, align: o.align, rotate: o.rotate });
}

/** 41pt white section title. */
function title(s, x, y, w, h, text, opt) {
  const o = opt || {};
  T(s, { x, y, w, h, text, size: o.size || 41, font: HEAD, color: C.white, align: o.align, rotate: o.rotate });
}

/** Three-square pinwheel that decorates the top-right corner. */
function pinwheel(s, x, y, d) {
  const g = (dir) => ({ dir, color: (t) => fade(t, { hold: 0.28, bg: C.black }) });
  ramp(s, Object.assign({ x: x + d, y, w: d, h: d }, g('right')));
  ramp(s, Object.assign({ x, y: y + d, w: d, h: d }, g('down')));
  ramp(s, Object.assign({ x: x + d, y: y + 2 * d, w: d, h: d }, g('left')));
}

/**
 * The staircase of purple blocks: every step is a narrow block stacked on a
 * wide one, each fading from the accent colour down to the backdrop.
 */
function stairs(s, o) {
  for (let i = 0; i < o.n; i++) {
    const x = o.x + i * o.dx;
    const y = o.y + i * o.dy;
    const col = (t) => fade(t, { bg: o.bg || C.black });
    ramp(s, { x, y, w: o.w, h: o.narrow, dir: 'down', color: col });
    ramp(s, { x, y: y + o.narrow, w: o.w, h: o.wide, dir: 'down', color: col });
  }
}

/** Rounded pill with the chip gradient (used for dates / "More Text Here"). */
function chip(s, x, y, w, h, label, opt) {
  const o = opt || {};
  ramp(s, { x, y, w, h, dir: 'right', color: (t) => fade(t, { hold: 0.28, bg: o.bg || C.black }) });
  T(s, { x, y: y + (h - 0.202) / 2 + 0.015, w: o.labelW || w * 0.79, h: 0.202, text: label,
    size: 8, font: HEAD, color: C.white, align: 'center' });
}

/** Free-form polygon in slide coordinates. */
function poly(s, pts, fill, alpha) {
  s.addShape('custGeom', {
    x: 0, y: 0, w: SLIDE_W, h: SLIDE_H,
    points: pts.map(([x, y]) => ({ x, y })).concat([{ close: true }]),
    fill: { color: fill, transparency: alpha }, line: { type: 'none' },
  });
}

/**
 * Half disc with a horizontal gradient: the band is painted as vertical strips,
 * then everything outside the arc is masked back out with the backdrop colour.
 * `up` puts the dome above its flat edge.
 */
function halfDisc(s, o) {
  const b = o.h / 2;                     // vertical semi-axis
  const flat = o.y + b;                  // straight side of the half disc
  const top = o.up ? o.y : flat;
  ramp(s, { x: o.x, y: top, w: o.w, h: b, dir: o.purpleLeft ? 'right' : 'left',
    color: (t) => fade(t, { hold: 0.28, bg: C.black }) });
  const arc = [];
  for (let i = 0; i <= 60; i++) {
    const u = i / 60;
    const dy = b * Math.sqrt(Math.max(0, 1 - Math.pow(2 * u - 1, 2)));
    arc.push([o.x + o.w * u, o.up ? flat - dy : flat + dy]);
  }
  const far = (o.up ? o.y : o.y + o.h) + (o.up ? -0.02 : 0.02);   // overshoot hides seams
  poly(s, [[o.x - 0.02, flat]].concat(arc, [[o.x + o.w + 0.02, flat],
    [o.x + o.w + 0.02, far], [o.x - 0.02, far]]), C.black);
}

/** Two thin rotated bars forming a tick or a cross. */
function mark(s, cx, cy, color, cross) {
  const bar = (len, ang, dx, dy) => s.addShape('roundRect', {
    x: cx + dx - len / 2, y: cy + dy - 0.0125, w: len, h: 0.025, rotate: ang,
    rectRadius: 0.012, fill: { color }, line: { type: 'none' },
  });
  if (cross) { bar(0.125, 45, 0, 0); bar(0.125, -45, 0, 0); }
  else { bar(0.058, 40, -0.035, 0.019); bar(0.105, -44, 0.025, 0); }
}

/* ------------------------------------------------------------------- slides */

// 1 - Cover
function slide01(s) {
  photo(s, 0, 0, SLIDE_W, SLIDE_H, { label: false });
  box(s, 0, 0, SLIDE_W, SLIDE_H, C.black, { alpha: 6 });
  const bg = '0D0D0D';
  stairs(s, { x: -0.079, y: 1.331, w: 0.849, narrow: 0.511, wide: 1.175, dx: 0.692, dy: -0.187, n: 5, bg });
  stairs(s, { x: 6.462, y: 3.089, w: 0.849, narrow: 0.511, wide: 1.175, dx: 0.692, dy: -0.187, n: 5, bg });
  T(s, { x: 0.562, y: 0.45, w: 1.421, h: 0.227, text: 'John Doe Presents', color: C.white });
  T(s, { x: 8.064, y: 4.797, w: 1.374, h: 0.379, text: ['New York City', '2025'], color: C.white, align: 'right' });
  title(s, 0.562, 1.974, 8.875, 1.287, 'Oil and Gas', { size: 72, align: 'center' });
  T(s, { x: 1.021, y: 3.285, w: 7.959, h: 0.278, text: 'PRESENTATION TEMPLATE', size: 12, font: HEAD,
    color: C.white, align: 'center' });
}

// 2 - Table of content
function slide02(s, n) {
  for (let r = 0; r < 5; r++) {
    const x = -0.742 + r * 0.248;
    const y = r * 1.125;
    const col = (t) => fade(t, { bg: C.black });
    ramp(s, { x, y, w: 0.831, h: 1.125, dir: 'down', color: col });
    ramp(s, { x: x + 0.831, y, w: 1.911, h: 1.125, dir: 'down', color: col });
  }
  chrome(s, n);
  eyebrow(s, 2.978, 0.822);
  title(s, 2.991, 1.048, 2.619, 1.439, 'Table of Content');
  T(s, { x: 2.981, y: 3.983, w: 2.363, h: 0.695, valign: 'bottom', ls: 1.4,
    text: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean' });

  const items = ['Introduction', 'Formation and Exploration', 'Refining and Processing',
    'Transportation and Distribution', 'Major Players in the Oil and Gas', 'Economic and Political Influence'];
  s.addShape('line', { x: 6.106, y: 0.973, w: 3.331, h: 0, line: { color: C.rule, width: 0.75 } });
  items.forEach((label, i) => {
    const y = 1.09 + i * 0.6134;
    T(s, { x: 6.0, y, w: 0.506, h: 0.379, text: String(i + 1).padStart(2, '0'), size: 18, font: HEAD,
      color: C.white, align: 'right' });
    T(s, { x: 6.615, y: y + 0.051, w: 2.822, h: 0.278, text: label, size: 12, font: HEAD, color: C.white });
    s.addShape('line', { x: 6.106, y: y + 0.496, w: 3.331, h: 0, line: { color: C.rule, width: 0.75 } });
  });
}

// 3 - Definition of Oil and Gas (four bullet blocks)
function slide03(s, n) {
  photo(s, 0, 0, SLIDE_W, SLIDE_H, { label: false });
  box(s, 0, 0, SLIDE_W, SLIDE_H, C.black, { alpha: 12 });
  pinwheel(s, 8.88, 0, 0.559);
  chrome(s, n);
  eyebrow(s, 0.562, 0.525);
  title(s, 0.576, 0.751, 7.799, 0.757, 'Definition of Oil and Gas');

  const cardBg = '1D1D1D';
  const blocks = [
    [0.562, 1.86, 'Importance in Global Economy ', 9],
    [5.475, 1.86, 'History of Oil and Gas Industry', 8],
    [0.562, 3.697, 'Geological Formation of Oil and Gas ', 8],
    [5.475, 3.697, 'Major Oil and Gas Reserves Around the World ', 8],
  ];
  const bullets = ['Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ',
    'Aenean commodo ligula eget dolor. Aenean massa. Cum sociis. ',
    'Donec quam felids, ultricies nec, pellentesque eu, pretium quis'];
  blocks.forEach(([x, y, head, size]) => {
    T(s, { x, y, w: 3.962, h: 0.278, text: head, size: 12, font: HEAD, color: C.white });
    T(s, { x, y: y + 0.327, w: 3.962, h: size === 9 ? 0.695 : 0.643, text: bullets, size, ls: 1.4, bullet: true });
    chip(s, x, y + 1.132, 1.635, 0.234, 'More Text Here', { bg: cardBg, labelW: 1.048 });
  });
  s.addShape('line', { x: 5.0, y: 1.86, w: 0, h: 1.14, line: { color: C.divider, width: 0.75, transparency: 33 } });
  s.addShape('line', { x: 5.0, y: 3.697, w: 0, h: 1.14, line: { color: C.divider, width: 0.75, transparency: 33 } });
}

// 4 - Drilling and Extraction
function slide04(s, n) {
  chrome(s, n);
  ramp(s, { x: 9.063, y: 0, w: 0.937, h: 0.937, dir: 'right', color: (t) => dim(C.purple, 1 - t) });
  ramp(s, { x: 9.062, y: 0.938, w: 0.937, h: 0.937, dir: 'left', color: (t) => dim(C.purple, 1 - t) });
  eyebrow(s, 5.287, 0.989);
  title(s, 5.301, 1.215, 3.396, 1.439, 'Drilling and Extraction');
  [3.074, 3.968].forEach((y, i) => {
    box(s, 6.012, y, 3.425, 0.668, C.white, { alpha: 95 });
    T(s, { x: 6.4, y: y + 0.092, w: 2.698, h: 0.483, valign: 'bottom', ls: 1.4, color: C.white,
      text: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo.' });
    iconDot(s, 5.794, y + 0.135, 0.398);
  });
  T(s, { x: 1.406, y: 4.71, w: 2.462, h: 0.227, text: 'www.yoursite.com/oilandgas', color: C.white, align: 'center', ls: 1 });
  photo(s, 0.563, 2.812, 4.15, 2.25);
  photo(s, 0.563, 0.562, 4.15, 2.25);
}

// 5 - Extraction Techniques (three staggered purple columns)
function slide05(s, n) {
  chrome(s, n);
  const cards = [
    [4.585, 0.978, 'Seismic Surveys', 2.844],
    [6.391, 0.613, 'Geological Map', 2.477],
    [8.195, 0.245, 'Well Logging', 2.111],
  ];
  cards.forEach(([x, y]) => {
    const col = (t) => fade(t, { bg: C.black });
    ramp(s, { x, y, w: 1.805, h: 1.334, dir: 'down', color: col });
    ramp(s, { x, y: y + 1.334, w: 1.805, h: 3.067, dir: 'down', color: col });
  });
  cards.forEach(([x, y, label, dotY], i) => {
    iconDot(s, x + 0.16, dotY, 0.4, { color: C.black, alpha: 73 });
    T(s, { x: x + 0.121, y: dotY + 0.672, w: 1.563, h: 0.278, text: label, size: 12, font: HEAD, color: C.white });
    T(s, { x: x + 0.121, y: dotY + 1.002, w: 1.563, h: 0.695, ls: 1.4, color: C.greyLt,
      text: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean' });
  });
  eyebrow(s, 0.562, 1.98);
  title(s, 0.576, 2.206, 3.396, 1.439, 'Extraction Techniques');
}

// 6 - Geological Formation (photo mosaic)
function slide06(s, n) {
  chrome(s, n);
  const tiles = [[4.89, 0.562, 1.437, 2.456], [6.328, 0.562, 1.21, 2.456], [7.537, 0.562, 1.9, 2.456],
    [4.89, 3.019, 2.162, 2.044], [7.052, 3.019, 2.385, 2.044]];
  tiles.forEach((t) => photo(s, t[0], t[1], t[2], t[3]));
  const wash = (t) => fade(t, { bg: C.photo, a0: 0.878 });
  ramp(s, { x: 4.89, y: 0.562, w: 1.437, h: 2.456, dir: 'up', color: wash });
  ramp(s, { x: 7.052, y: 3.019, w: 2.385, h: 2.044, dir: 'up', color: wash });
  T(s, { x: 4.944, y: 2.691, w: 1.343, h: 0.227, text: 'Onshore', color: C.white, align: 'center', ls: 1 });
  T(s, { x: 7.573, y: 4.73, w: 1.343, h: 0.227, text: 'Directional', color: C.white, align: 'center', ls: 1 });
  eyebrow(s, 0.562, 0.581);
  title(s, 0.576, 0.806, 3.811, 2.121, 'Geological Formation of Oil and Gas ');
  T(s, { x: 0.562, y: 3.831, w: 3.681, h: 1.002, ls: 1.4, before: 7, text: [
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes,. ',
    'Donec quam felids, ultricies nec, pellentesque eu, pretium'] });
  s.addShape('line', { x: 4.594, y: 0.562, w: 0, h: 4.5, line: { color: C.divider, width: 0.75, transparency: 33 } });
}

// 7 - Who We're
function slide07(s, n) {
  chrome(s, n);
  const col = (t) => fade(t, { bg: C.black });
  ramp(s, { x: 3.802, y: -0.539, w: 2.396, h: 2.632, dir: 'down', color: col });
  ramp(s, { x: 3.802, y: 0.671, w: 2.396, h: 6.054, dir: 'down', color: col });
  photo(s, 0, 1.203, 3.219, 3.219);
  photo(s, 6.781, 1.203, 3.219, 3.219);
  eyebrow(s, 3.864, 1.577, { color: C.white, align: 'center' });
  title(s, 3.302, 1.803, 3.396, 0.757, 'Who We\u2019re', { align: 'center' });
  T(s, { x: 3.578, y: 2.834, w: 2.844, h: 1.214, align: 'center', ls: 1.4, before: 7, color: C.white, text: [
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, ',
    'Donec pede justo, fringilla vel, aliquet nec, '] });
}

// 8 - Types of Drilling
function slide08(s, n) {
  photo(s, 0, 0, SLIDE_W, 2.523, { label: false });
  chrome(s, n);
  ramp(s, { x: 0.499, y: -0.953, w: 3.127, h: 2.217, dir: 'down', color: (t) => fade(t, { bg: C.photo }) });
  ramp(s, { x: 0.499, y: 1.418, w: 3.127, h: 5.098, dir: 'down',
    color: (t, cx, cy) => fade(t, { bg: cy < 2.523 ? C.photo : C.black }) });
  eyebrow(s, 0.77, 0.629, { color: C.white });
  title(s, 0.783, 0.855, 2.736, 1.439, 'Types of Drilling ');
  const cols = [[0.844, 'Onshore', 0.922, true], [3.922, 'Offshore', 4.0, false], [7.0, 'Directional', 7.078, false]];
  cols.forEach(([x, label, dotX, dark]) => {
    iconDot(s, dotX, 2.877, 0.4, dark ? { color: C.black, alpha: 73 } : {});
    T(s, { x, y: 3.392, w: 2.367, h: 0.278, text: label, size: 12, font: HEAD, color: C.white });
    s.addShape('line', { x: x + 0.093, y: 3.838, w: 0, h: 0.855, line: { color: C.divider, width: 0.75, transparency: 33 } });
    T(s, { x: x + 0.258, y: 3.813, w: 2.18, h: 0.907, ls: 1.4,
      text: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis' });
  });
}

// 9 - All the Best Team
function slide09(s, n) {
  chrome(s, n);
  photo(s, 0.563, 1.901, 2.648, 3.161);
  ramp(s, { x: 0.563, y: 1.901, w: 2.648, h: 3.161, dir: 'up',
    color: (t) => fade(t, { bg: C.photo, a0: 0.878 }) });
  eyebrow(s, 0.562, 0.581);
  title(s, 0.576, 0.806, 8.861, 0.757, 'All the Best Team');
  T(s, { x: 0.809, y: 2.106, w: 2.156, h: 0.379, text: 'Our Team Group', size: 18, font: HEAD, color: C.white });

  const cards = [
    [3.676, 'Oilfield Services', 'March 19, 2025', 'Maintenance',
      '\u201CCompanies that provide specialized services such as drilling, well mainte, etc.'],
    [6.789, 'Downstream ', 'April 27, 2025', 'Petrochemicals',
      '\u201CRefining, marketing, and distribution of petroleum products\u201D '],
  ];
  cards.forEach(([x, head, date, sub, quote]) => {
    box(s, x, 1.901, 2.648, 3.161, C.white, { alpha: 95 });
    const tx = x + 0.246;
    T(s, { x: tx, y: 2.106, w: 2.156, h: 0.379, text: head, size: 18, font: HEAD, color: C.white });
    chip(s, tx, 2.697, 1.165, 0.234, date, { labelW: 0.919 });
    T(s, { x: tx, y: 3.077, w: 2.156, h: 0.252, text: sub, size: 11, font: HEAD, color: C.white });
    T(s, { x: tx, y: 3.345, w: 2.156, h: 0.483, ls: 1.4,
      text: 'Lorem ipsum dolor sit amet, adipiscing elit. Aenean commodo' });
    for (let i = 0; i < 3; i++) photo(s, tx + i * 0.4375, 3.994, 0.355, 0.355, { round: true });
    T(s, { x: tx, y: 4.443, w: 2.156, h: 0.415, text: quote, size: 8, ls: 1.4, italic: true });
  });
  pinwheel(s, 8.742, 0, 0.629);
}

// 10 - Refining and Processing (three ring gauges)
function slide10(s, n) {
  chrome(s, n);
  const col = (t) => fade(t, { bg: C.black });
  ramp(s, { x: 2.724, y: -0.427, w: 3.127, h: 1.166, dir: 'down', color: col });
  ramp(s, { x: 2.724, y: 2.627, w: 3.127, h: 2.682, dir: 'down', color: col });
  const gauges = [[0.665, '94%', 238.78], [2.297, '26%', 359.59], [3.93, '74%', 184.93]];
  gauges.forEach(([y, label, end]) => {
    s.addShape('ellipse', { x: 3.891, y, w: 0.793, h: 0.793, fill: { type: 'none' },
      line: { color: C.black, width: 3, transparency: 77 } });
    s.addShape('arc', { x: 3.891, y, w: 0.793, h: 0.793, angleRange: [270, end],
      fill: { type: 'none' }, line: { color: C.white, width: 3 } });
    T(s, { x: 4.004, y: y + 0.257, w: 0.565, h: 0.278, text: label, size: 12, font: HEAD, color: C.white, align: 'center' });
    T(s, { x: 3.734, y: y + 0.878, w: 1.106, h: 0.227, text: 'Your Text Here', color: C.white, align: 'center' });
  });
  eyebrow(s, 5.753, 1.223);
  title(s, 5.766, 1.449, 3.811, 1.439, 'Refining and Processing');
  T(s, { x: 5.753, y: 3.401, w: 3.681, h: 1.002, ls: 1.4, before: 7, text: [
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes,. ',
    'Donec quam felids, ultricies nec, pellentesque eu, pretium'] });
  photo(s, 0, 0, 3.5, SLIDE_H);
}

// 11 - Storage and Distribution
function slide11(s, n) {
  chrome(s, n);
  eyebrow(s, 0.562, 0.581);
  title(s, 0.576, 0.806, 6.043, 1.439, 'Storage and Distribution  ');
  T(s, { x: 0.562, y: 2.457, w: 5.008, h: 0.695, ls: 1.4,
    text: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes,. Donec quam felids, ultricies nec, pellentesque eu, pretium' });
  [['Plan A', 0.562], ['Plan B', 4.406]].forEach(([label, x]) => {
    box(s, x, 3.797, 3.523, 1.239, C.ink);
    chip(s, x + 0.131, 3.681, 0.787, 0.234, label, { labelW: 0.621 });
    photo(s, x + 0.131, 4.068, 0.742, 0.817);
    T(s, { x: x + 0.91, y: 4.067, w: 2.156, h: 0.278, text: 'Your Text Here #1', size: 12, font: HEAD, color: C.white });
    T(s, { x: x + 0.91, y: 4.403, w: 2.531, h: 0.483, ls: 1.4,
      text: 'Lorem ipsum dolor sit amet, adipiscing elit. Aenean commodo ligula eget dolor. ' });
  });
  // giant "8" built from two rings
  box(s, 5.988, 0, 1.069, 1.069, C.violet, { shape: 'donut' });
  box(s, 5.988, 0.807, 1.069, 1.069, C.violet, { shape: 'donut' });
  box(s, 6.176, 0.807, 0.693, 0.261, C.violet);
  photo(s, 7.055, 0, 2.945, SLIDE_H);
}

// 12 - Gallery
function slide12(s, n) {
  chrome(s, n);
  const col = (t) => dim(C.purple, 1 - t);
  ramp(s, { x: 8.671, y: -0.443, w: 0.886, h: 0.886, dir: 'down', color: col });
  ramp(s, { x: 8.671, y: 0.443, w: 0.886, h: 0.886, dir: 'down', color: col });
  eyebrow(s, 3.861, 0.585, { align: 'center' });
  title(s, 3.091, 0.811, 3.811, 0.757, 'Gallery Slide', { align: 'center' });
  T(s, { x: 2.138, y: 1.692, w: 5.725, h: 0.483, align: 'center', ls: 1.4,
    text: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes,.  Donec' });
  [[0.562, 'left', 'Your Text Here #1', 0.562], [6.903, 'right', 'Your Text Here #2', 8.508]].forEach(
    ([x, align, label, chipX]) => {
      chip(s, chipX, 3.618, 1.165, 0.234, 'March 19, 2025', { labelW: 0.919 });
      T(s, { x: align === 'left' ? x : 7.278, y: 3.999, w: 2.156, h: 0.278, text: label, size: 12,
        font: HEAD, color: C.white, align });
      T(s, { x, y: 4.351, w: 2.531, h: 0.483, align, ls: 1.4,
        text: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean' });
    });
  const tiles = [[2.231, 2.305], [3.339, 2.305], [4.446, 2.305], [5.554, 2.305], [6.661, 2.305],
    [3.339, 3.41], [4.446, 3.41], [5.554, 3.411], [4.446, 4.518]];
  tiles.forEach(([x, y]) => photo(s, x, y, 1.107, 1.107));
}

// 13 - Problem Statement
function slide13(s, n) {
  chrome(s, n);
  eyebrow(s, 0.526, 1.254);
  title(s, 0.54, 1.48, 4.424, 1.439, 'Problem Statement');
  T(s, { x: 0.562, y: 3.464, w: 3.729, h: 0.907, ls: 1.4,
    text: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes,. Donec quam felids, ultricies nec, pellentesque eu, pretium' });
  const col = (t) => fade(t, { bg: C.black });
  ramp(s, { x: 4.938, y: 0.562, w: 2.25, h: 2.25, dir: 'up', color: col });
  iconDot(s, 5.863, 0.887, 0.4, { color: C.black, alpha: 59 });
  T(s, { x: 5.281, y: 1.464, w: 1.563, h: 0.278, text: 'Problem A', size: 12, font: HEAD, color: C.white, align: 'center' });
  T(s, { x: 5.156, y: 1.793, w: 1.812, h: 0.695, align: 'center', ls: 1.4,
    text: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula' });
  ramp(s, { x: 7.188, y: 2.812, w: 2.25, h: 2.25, dir: 'down', color: col });
  iconDot(s, 8.113, 3.137, 0.4, { color: C.black, alpha: 59 });
  T(s, { x: 7.531, y: 3.714, w: 1.563, h: 0.278, text: 'Problem B', size: 12, font: HEAD, color: C.white, align: 'center' });
  T(s, { x: 7.406, y: 4.043, w: 1.812, h: 0.695, align: 'center', ls: 1.4,
    text: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula' });
  photo(s, 7.188, 0.562, 2.25, 2.25);
  photo(s, 4.938, 2.812, 2.25, 2.25);
}

// 14 - Oil Trading and Pricing Factors (two tilted phone mock-ups)
function slide14(s, n) {
  chrome(s, n);
  [[1.185, 2.377], [0.013, -0.99]].forEach(([x, y]) => {
    box(s, x, y, 2.488, 4.843, '111111', { rotate: 30, rectRadius: 0.28 });
    photo(s, x + 0.24, y + 0.24, 2.008, 4.363, { rotate: 30 });
  });
  eyebrow(s, 3.803, 0.891);
  title(s, 3.817, 1.117, 5.621, 1.439, 'Oil Trading and Pricing Factors ');
  T(s, { x: 4.939, y: 3.428, w: 2.644, h: 0.379, text: '64%', size: 18, font: HEAD, color: C.white });
  T(s, { x: 4.939, y: 3.851, w: 4.202, h: 0.907, ls: 1.4,
    text: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes,. Donec quam felids, ultricies nec, pellentesque eu, pretium' });
  pinwheel(s, 8.742, 0, 0.629);
}

// 15 - Infographic (S-curve of half discs)
function slide15(s, n) {
  chrome(s, n);
  const arcs = [[1.203, true], [2.722, false], [4.241, true], [5.759, false], [7.278, true]];
  arcs.forEach(([x, up]) => halfDisc(s, { x, y: 2.817, w: 1.519, h: 1.364, up, purpleLeft: up }));
  const labels = [[0.978, 1.891, 1.181], [2.497, 4.294, 2.7], [4.016, 1.891, 4.219], [5.534, 4.294, 5.738], [7.053, 1.891, 7.256]];
  labels.forEach(([bx, hy, hx]) => {
    T(s, { x: hx, y: hy, w: 1.563, h: 0.278, text: 'Your Text Here', size: 12, font: HEAD, color: C.white, align: 'center' });
    T(s, { x: bx, y: hy + 0.33, w: 1.969, h: 0.483, align: 'center', ls: 1.4, color: C.greyLt,
      text: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ' });
  });
  [1.764, 3.282, 4.801, 6.32, 7.839].forEach((x) => iconDot(s, x, 3.3, 0.398));
  eyebrow(s, 3.864, 0.578, { align: 'center' });
  title(s, 0.563, 0.804, 8.875, 0.757, 'Infographic', { align: 'center' });
}

// 16 - Target Agenda (June 2025 calendar)
function slide16(s, n) {
  chrome(s, n);
  box(s, 3.672, 1.791, 2.656, 3.121, C.white, { alpha: 95 });
  ramp(s, { x: 3.672, y: 1.791, w: 2.656, h: 0.372, dir: 'left',
    color: (t) => dim(C.violet, 1 - 0.9 * t) });
  T(s, { x: 3.672, y: 1.86, w: 2.656, h: 0.234, text: 'June 2025', size: 11, color: C.white, align: 'center' });

  const colX = (c) => 3.691 + c * 0.3726;
  const rowY = (r) => 2.223 + r * 0.3724;
  ['M', 'T', 'W', 'T', 'F', 'S', 'S'].forEach((d, c) => {
    T(s, { x: colX(c), y: rowY(0) + 0.06, w: 0.373, h: 0.25, text: d, size: 7, font: HEAD,
      color: C.calDay, bold: true, align: 'center' });
  });
  const weeks = [
    [27, 28, 29, 30, 31, 1, 2],
    [3, 4, 5, 6, 7, 8, 9],
    [10, 11, 12, 13, 14, 15, 16],
    [17, 18, 19, 20, 21, 22, 23],
    [24, 25, 26, 27, 28, 29, 30],
    [1, 2, 3, 4, 5, 6, 7],
  ];
  const marked = { '0-5': 1, '2-0': 1, '3-3': 1 };
  weeks.forEach((week, r) => week.forEach((day, c) => {
    const x = colX(c);
    const y = rowY(r + 1) + (r === 5 ? 0.012 : 0);
    const on = marked[r + '-' + c];
    if (on) box(s, x, y, 0.373, 0.372, C.violet, { shape: 'ellipse' });
    T(s, { x, y: y + 0.06, w: 0.373, h: 0.25, text: String(day), size: 7, font: HEAD,
      color: on ? C.white : C.calDay, align: 'center' });
  }));

  [['Target Reached', 1.025, 1.44, C.ok, false], ['Not Target', 7.013, 7.429, C.no, true]].forEach(
    ([head, cx, tx, tint, cross]) => {
      T(s, { x: cx, y: 2.313, w: 1.87, h: 0.278, text: head, size: 12, font: HEAD, color: C.white });
      for (let i = 0; i < 4; i++) {
        const y = 2.812 + i * 0.4295;
        box(s, cx, y, 0.289, 0.289, '3E3E3E', { shape: 'ellipse', alpha: 73 });
        mark(s, cx + 0.1445, y + 0.1445, tint, cross);
        T(s, { x: tx, y: y + 0.031, w: 1.75, h: 0.227, text: 'Insert Your Target Here', color: C.white });
      }
    });
  eyebrow(s, 3.864, 0.578, { align: 'center' });
  title(s, 0.563, 0.804, 8.875, 0.757, 'Target Agenda', { align: 'center' });
}

// 17 - Matrix Diagram
function slide17(s, n) {
  chrome(s, n);
  s.addShape('line', { x: 0.911, y: 1.812, w: 0, h: 2.928, rotate: 180,
    line: { color: C.axis, width: 1, endArrowType: 'triangle' } });
  s.addShape('line', { x: 0.911, y: 4.74, w: 8.177, h: 0,
    line: { color: C.axis, width: 1, endArrowType: 'triangle' } });
  const ax = { size: 7, font: HEAD, color: C.greyLt, align: 'center' };
  T(s, Object.assign({ x: 0.512, y: 4.82, w: 0.633, h: 0.189, text: 'Low' }, ax));
  T(s, Object.assign({ x: 8.856, y: 4.82, w: 0.633, h: 0.189, text: 'High' }, ax));
  T(s, Object.assign({ x: 0.305, y: 2.033, w: 0.633, h: 0.169, text: 'High', rotate: -90 }, ax));
  T(s, Object.assign({ x: 3.946, y: 4.82, w: 2.571, h: 0.189, text: 'Your text goes here' }, ax));
  T(s, Object.assign({ x: -0.215, y: 3.373, w: 1.673, h: 0.169, text: 'Your text goes here', rotate: -90 }, ax));

  const cells = [[1.155, 2.063, 1], [3.847, 2.063, 0], [6.539, 2.063, 1],
    [1.155, 3.405, 0], [3.847, 3.405, 1], [6.539, 3.405, 0]];
  cells.forEach(([x, y, hot]) => {
    if (hot) ramp(s, { x, y, w: 2.307, h: 1.088, dir: 'right',
      color: (t) => fade(t, { hold: 0.28, bg: C.black }) });
    else box(s, x, y, 2.307, 1.088, C.white, { alpha: 95 });
    T(s, { x: x + 0.169, y: y + 0.137, w: 1.563, h: 0.278, text: 'Your Text Here', size: 12, font: HEAD, color: C.white });
    T(s, { x: x + 0.169, y: y + 0.467, w: 1.969, h: 0.483, ls: 1.4, color: C.greyLt,
      text: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ' });
  });
  eyebrow(s, 3.864, 0.578, { align: 'center' });
  title(s, 0.563, 0.804, 8.875, 0.757, 'Matrix Diagram', { align: 'center' });
}

/** Simplified silhouette of the United States, in 1/1000 inch, relative to its bbox. */
const USA_MAP = { x: 1.06, y: 0.68, shapes: [
  '135,25,140,240,0,715,160,1290,250,1330,170,1345,310,1370,305,1435,340,1405,395,1520,565,1530,845,1690,1210,1680,1435,1940,1510,1870,1590,1875,1780,2180,1915,2230,1955,2030,2150,1880,2555,1910,2595,1785,2745,1755,2880,1810,2990,1765,3080,1835,3160,2060,3285,2170,3190,2230,3310,2180,3330,2005,3165,1710,3165,1570,3300,1365,3520,1175,3455,1010,3530,680,3675,560,3775,550,3690,470,3690,380,3885,185,3835,55,3660,140,3410,60,3245,140,3060,105,2830,45,2685,110,2560,55,2405,90,1170,10,1130,0,410,105,325,50,285,105',
  '270,2320,320,2340,410,2470,565,2500,600,2540,435,2610,480,2645,435,2700,320,2660,290,2620,250,2665,145,2685,150,2645,80,2670,55,2705,0,2725,240,2565,275,2530,215,2510,235,2470,150,2455,105,2415,150,2385,140,2350,240,2400,220,2360',
  '755,2680,760,2700,715,2720,700,2735,720,2755,780,2735,800,2700,790,2670',
] };

// 18 - USA Map
function slide18(s, n) {
  chrome(s, n);
  USA_MAP.shapes.forEach((outline) => {
    const nums = outline.split(',').map(Number);
    const pts = [];
    for (let i = 0; i < nums.length; i += 2) pts.push([USA_MAP.x + nums[i] / 1000, USA_MAP.y + nums[i + 1] / 1000]);
    poly(s, pts, C.rule, 19);
  });
  eyebrow(s, 5.622, 1.113);
  title(s, 5.622, 1.338, 3.396, 0.757, 'USA Map');
  T(s, { x: 5.622, y: 2.288, w: 3.701, h: 0.695, ls: 1.4,
    text: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes,. ' });
  [[1.102, '$361.000', 'up'], [3.7, '$517.000', 'down'], [6.302, '$831.000', 'up']].forEach(([x, value, dir]) => {
    ramp(s, { x, y: 3.814, w: 2.604, h: 1.002, dir, color: (t) => fade(t, { bg: C.black }) });
    T(s, { x: x + 0.888, y: 3.954, w: 1.413, h: 0.242, text: 'Your Value Here', size: 8, font: HEAD,
      color: C.white, ls: 1.3 });
    T(s, { x: x + 0.888, y: 4.248, w: 1.413, h: 0.429, text: value, size: 21, font: HEAD, color: C.white });
    iconDot(s, x + 0.303, 4.116, 0.4, { color: C.black, alpha: 73 });
  });
}

// 19 - Oil Report (data table)
function slide19(s, n) {
  chrome(s, n);
  const colX = [2.581, 4.774, 6.274, 7.844];
  const colW = [1.865, 1.079, 1.079, 1.079];
  const heads = ['Subtitle text one here', 'Value content', 'Value content', 'Value content'];
  heads.forEach((h, i) => {
    T(s, { x: colX[i], y: 1.006, w: i === 0 ? 1.865 : 1.173, h: 0.245, text: h, size: 9, font: HEAD,
      color: C.white, ls: 1.2 });
    box(s, [2.452, 4.645, 6.145, 7.716][i], 1.097, 0.11, 0.066, C.white);
  });
  const rows = [
    ['Startup event organizer', 'Value one', 'Mar 16, 2025', 'Category 01'],
    ['Conference event', 'Value one', 'Mar 16, 2025', 'Category 02'],
    ['Conference event', 'Value two', 'Mar 16, 2025', 'Category 03'],
    ['Startup event organizer', 'Value three', 'Mar 16, 2025', 'Category 04'],
    ['Conference event', 'Value three', 'Mar 16, 2025', 'Category 05'],
    ['Startup event organizer', 'Value two', 'Mar 16, 2025', 'Category 06'],
    ['Startup event organizer', 'Value two', 'Mar 16, 2025', 'Category 07'],
  ];
  rows.forEach((cells, r) => {
    const y = 1.412 + r * 0.506;
    const hot = r === 2;
    if (hot) ramp(s, { x: 2.185, y, w: 7.257, h: 0.444, dir: 'right',
      color: (t) => fade(t, { hold: 0.28, bg: C.black }) });
    else box(s, 2.185, y, 7.257, 0.444, C.white, { alpha: 95 });
    T(s, { x: 2.417, y: y + 0.11, w: 0.15, h: 0.15, text: '\u2192', size: 8, color: C.white });
    cells.forEach((t, i) => T(s, { x: colX[i], y: y + 0.101, w: colW[i], h: 0.245, text: t, size: 9,
      font: HEAD, color: hot ? C.white : C.greyMid, ls: 1.2 }));
    box(s, 9.061, y + 0.196, 0.212, 0.051, C.greyMid, { alpha: 75 });
  });
  chip(s, 2.185, 0.562, 1.229, 0.234, 'Conference', { labelW: 0.788 });
  chip(s, 3.274, 0.562, 1.865, 0.234, 'Program Introduction', { labelW: 1.195 });
  eyebrow(s, -0.648, 2.488, { rotate: -90 });
  title(s, -0.635, 2.714, 3.811, 0.757, 'Oil Report', { rotate: -90 });
}

// 20 - Thank You
function slide20(s, n) {
  chrome(s, n);
  stairs(s, { x: 1.76, y: 1.804, w: 1.295, narrow: 0.957, wide: 2.201, dx: 1.296, dy: -0.286, n: 5 });
  title(s, 0.562, 1.974, 8.875, 1.287, 'Thank You', { size: 72, align: 'center' });
  T(s, { x: 1.021, y: 3.285, w: 7.959, h: 0.278, text: 'SEE YOU NEXT PRESENTATION!', size: 12, font: HEAD,
    color: C.white, align: 'center' });
}

/* -------------------------------------------------------------------- build */

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'DECK', width: SLIDE_W, height: SLIDE_H });
pptx.layout = 'DECK';
pptx.title = 'Oil and Gas';
pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };

[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
 slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
  .forEach((build, i) => {
    const s = pptx.addSlide();
    s.background = { color: C.black };
    build(s, i + 1);
  });

pptx.writeFile({ fileName: path.join(__dirname, '0d519d00-5b38-4091-ad98-ffb9e4af02a7_grok_final.pptx') })
  .then((f) => console.log('wrote ' + f));
