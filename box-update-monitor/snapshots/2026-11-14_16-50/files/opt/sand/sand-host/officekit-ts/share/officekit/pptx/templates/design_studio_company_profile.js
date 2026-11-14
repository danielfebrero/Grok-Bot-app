/**
 * "FINA" business presentation template — 39 slides, 13.333 x 7.5 in.
 * Rebuilt with pptxgenjs only. Photographs in the original deck are replaced
 * by flat colour rectangles that keep the original position / size / tone.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const C = {
  orange: 'FF9800',
  ink: '0D0D0D', // tx1 95%
  dark: '262626', // tx1 85%
  g40: '404040',
  g59: '595959',
  g80: '808080',
  gA6: 'A6A6A6',
  gBF: 'BFBFBF',
  gCC: 'CCCCCC',
  gD0: 'D0CECE',
  gD9: 'D9D9D9',
  gF2: 'F2F2F2',
  white: 'FFFFFF',
  black: '000000',
  plum: '4B486C',
  teal: '11998E',
  mint: '29CE84',
  lime: '38EF7D',
};
/* dominant tone of each photograph these placeholders replace */
const PIC = {
  gray: '888888',
  light: 'A8ACAD',
  deep: '676767',
  sage: '768C8A',
  slate: '2B4041',
  mist: 'A7BAB8',
};

/* -------------------------------------------------------------------- fonts */
const F = {
  sb: 'Poppins SemiBold',
  md: 'Poppins Medium',
  rg: 'Poppins',
  ws: 'Work Sans Medium',
  os: 'Open Sans SemiBold',
};

/* ------------------------------------------------------- recurring text bits */
const L = {
  a: 'Lorep  ipsum duis aute irure dolor in kauselih oilue epree',
  a2: 'Lorep  ipsum duis aute irure dolor in kauselih oilue eprehe',
  b: 'deriti vols esse cill inure dolorlaboru sit amet. Duis autelo irusitakus reprehenderi Voluptate lorem kuisais.',
  c: 'Lorep  ipsum duis aute irure dolor in kauselih oilusioisduili reprehenderitisi voluptates esse cill inure dolorlaborusita amet. Duis aute irusitaseiad dolorin repreheno dui lroeml',
  d: 'Lorep  ipsum duis aute irure dolor in kauselih oilusioi reprehenderitilores voluptates esse',
  e: 'Lorep  ipsum duis aute irure dolor in kauselih oilusioi repiehenderiti volui ptates esse cill inure dolrlasue boru sit amet. Duis aute',
  f: 'Lorep  ipsum duis aute irure dolor in kauselih oilue eprehe deriti vols esse cill inure dolorlaboru sit amet. Duis autelo irusitakus reprelikasi henderi Voluptate lorem kuisais.',
  g: 'Lorep  ipsum duis aute irure dolor in kauselih oilue reprehenderiti vols',
  h: 'esse cill inure dolorlaboru sit amet. Duis aute irusitakus reprehenderi',
  i: 'Voluptate lorem kuisais.',
  j: 'Lorep  ipsum duis aute irure dolor in kauselih oilusioi reprehenderiti voluptates esse cill inure dolorlaboru',
  quote: 'Lorep  ipsum duis aute irure dolor in kauselih oilusioi reprehenderiti voluptates esse cill inure dolorlaboru',
  tt: 'TEXT TITTLE HERE',
  li: 'Lorem Ipsum',
  duist: 'Lorem Ipsum  duist',
  aure: 'Aure inure dolor',
};

/* ------------------------------------------------------------------ helpers */
const S = (s, shape, o) => s.addShape(shape, o);

/** plain filled rectangle */
function rect(s, x, y, w, h, color, extra) {
  S(s, 'rect', Object.assign({ x, y, w, h, fill: { color }, line: { type: 'none' } }, extra));
}

/** flat colour stand-in for a photograph */
function photo(s, x, y, w, h, color) {
  rect(s, x, y, w, h, color || PIC.gray);
}

/** straight rule; dx/dy give the length (dy 0 => horizontal) */
function rule(s, x, y, dx, dy, color, opts) {
  const ln = Object.assign({ color, width: 0.5 }, (opts || {}).line);
  S(s, 'line', { x, y, w: dx, h: dy, line: ln });
}

/** small arrow used all over the deck ("<-" when dir = -1) */
function arrow(s, x, y, len, dir, color) {
  const ln = { color: color || C.white, width: 0.5 };
  if (dir < 0) ln.beginArrowType = 'triangle';
  else ln.endArrowType = 'triangle';
  S(s, 'line', { x, y, w: len, h: 0, line: ln });
}

/**
 * text block.
 *   box   [x, y, w, h]
 *   lines array of paragraphs; a paragraph is a string or an array of
 *         [text, overrides] run tuples
 *   base  style shared by every run / paragraph
 */
function text(s, box, lines, base) {
  const st = base || {};
  const runs = [];
  lines.forEach(line => {
    const parts = typeof line === 'string' ? [[line]] : line;
    parts.forEach((p, j) => {
      runs.push({
        text: p[0],
        options: Object.assign({}, st, p[1], { breakLine: j === parts.length - 1 }),
      });
    });
  });
  /* PowerPoint text boxes are top-anchored with 0.1"/0.05" insets */
  s.addText(runs, Object.assign({
    x: box[0], y: box[1], w: box[2], h: box[3],
    valign: 'top', margin: [7.2, 7.2, 3.6, 3.6],
  }, st));
}

/* text presets */
const body = c => ({ fontFace: F.rg, fontSize: 8, color: c || C.g59, lineSpacingMultiple: 1.5 });
const head = { fontFace: F.md, fontSize: 40, color: C.orange };
const kicker = c => ({ fontFace: F.sb, fontSize: 10.5, color: c || C.ink, charSpacing: 0.5 });

/* ------------------------------------------------------- deck-wide furniture */
function mixStop(t) {
  const stops = [[0, 0x11, 0x99, 0x8e], [0.63, 0x29, 0xce, 0x84], [1, 0x38, 0xef, 0x7d]];
  let i = t <= stops[1][0] ? 0 : 1;
  const [p0, r0, g0, b0] = stops[i];
  const [p1, r1, g1, b1] = stops[i + 1];
  const k = (t - p0) / (p1 - p0);
  const h = v => Math.round(v).toString(16).padStart(2, '0').toUpperCase();
  return h(r0 + (r1 - r0) * k) + h(g0 + (g1 - g0) * k) + h(b0 + (b1 - b0) * k);
}

/** the teal -> green gradient strip down the left edge (stacked slices) */
function gradientBar(s) {
  const n = 20;
  for (let i = 0; i < n; i++) rect(s, 0, (i * 7.5) / n, 0.08, 7.5 / n + 0.02, mixStop(i / (n - 1)));
}

/** hamburger mark, top right */
function burger(s, color) {
  [0.185, 0.2787, 0.3725].forEach(y => rect(s, 12.911, y, 0.207, 0.0436, color));
}

/**
 * page furniture: "Modern style" / burger / "Art company" / page number
 * o = { num, top, foot, page, burger, noFoot }
 */
function chrome(s, o) {
  o = o || {};
  const base = { fontFace: F.sb, fontSize: 12 };
  text(s, [0.448, 0.149, 1.306, 0.303], ['Modern style'], Object.assign({ color: o.top || C.ink }, base));
  burger(s, o.burger || C.plum);
  if (!o.noFoot) {
    text(s, [o.footX === undefined ? 0.448 : o.footX, o.footY || 7.049, 1.324, 0.303], ['Art company'],
      Object.assign({ color: o.foot || C.ink }, base));
  }
  if (o.num) {
    text(s, [12.806, 7.049, 0.42, 0.303], [o.num], Object.assign({ color: o.page || C.ink }, base));
  }
}

/** rotated "01/ 02" pagination badge (group origin of the original file) */
function pager(s, gx, gy, rot, tColor, lColor) {
  const cx = gx + 0.8625, cy = gy + 0.408;
  const st = { fontFace: F.ws, fontSize: 12, color: tColor || C.g40, align: 'center', lineSpacingMultiple: 2.5 };
  if (rot) {
    text(s, [cx - 0.369, cy + 0.0855, 0.738, 0.816], ['01/ 02'], Object.assign({ rotate: 90 }, st));
    S(s, 'line', { x: cx + 0.037, y: cy - 0.8625, w: 0, h: 0.887, line: { color: lColor || C.ink, width: 0.5 } });
  } else {
    text(s, [gx + 0.987, gy, 0.738, 0.816], ['01/ 02'], st);
    rule(s, gx, gy + 0.371, 0.887, 0, lColor || C.ink);
  }
}

/** orange + dark "previous / next" button pair */
function navPair(s, x, y) {
  rect(s, x, y, 1.039, 0.665, C.orange);
  rect(s, x + 1.04, y, 1.039, 0.665, C.dark);
  arrow(s, x + 0.396, y + 0.333, 0.195, -1);
  arrow(s, x + 1.464, y + 0.333, 0.195, 1);
}

/** dark/orange square holding a single white arrow */
function arrowTile(s, x, y, color, dir) {
  rect(s, x, y, 0.778, 0.78, color);
  arrow(s, x + 0.293, y + 0.404, 0.246, dir === undefined ? -1 : dir);
}

/** stand-in for the line pictograms of the original: outlined tile + dot */
function glyph(s, x, y, w, h, color) {
  S(s, 'roundRect', {
    x, y, w, h, rectRadius: Math.min(w, h) * 0.22,
    fill: { type: 'none' }, line: { color, width: Math.max(1, Math.min(w, h) * 3) },
  });
  S(s, 'ellipse', {
    x: x + w * 0.38, y: y + h * 0.38, w: w * 0.24, h: h * 0.24,
    fill: { color }, line: { type: 'none' },
  });
}

/** device mock-up: dark shell + light screen */
function device(s, shell, screen, radius, shellColor) {
  S(s, 'roundRect', {
    x: shell[0], y: shell[1], w: shell[2], h: shell[3], rectRadius: radius || 0.05,
    fill: { color: shellColor || '1B1B1B' }, line: { type: 'none' },
  });
  photo(s, screen[0], screen[1], screen[2], screen[3], screen[4] || PIC.gray);
}

/** "TEXT TITTLE HERE" + rule + small paragraph, the deck's standard caption */
function caption(s, x, y, w, label, para, opts) {
  const o = opts || {};
  text(s, [x + (o.right ? w - 1.593 : 0), y, 1.593, 0.286], [label || L.tt],
    Object.assign(kicker(o.color), { align: o.right ? 'right' : 'left' }));
  rule(s, x + (o.right ? w - 2.109 : 0.1), y + 0.435, 2.109, 0, o.ruleColor || C.g80);
  text(s, [x, y + (o.gap || 0.586), w, 0.688], [para || L.d],
    Object.assign(body(o.bodyColor || C.g59), { align: o.right ? 'right' : 'left' }));
}

/** the two stacked lorem paragraphs used on most editorial slides */
function loremPair(s, x, y, color) {
  text(s, [x, y, 3.446, 0.688], [L.a, L.b], body(color));
  text(s, [x, y + 0.787, 3.561, 0.688], [L.c], body(color));
}

/* ============================================================ cover slides */
/** slides 1 / 18 / 39 share the orange cover layout */
function coverSlide(s, o) {
  rect(s, 0, 0, 0.08, 7.5, C.white);
  rect(s, 0.08, 0, 7.611, 7.5, C.orange);
  photo(s, 7.691, 0, 5.642, 7.5, o.picture);
  chrome(s, { num: o.num, top: C.white, foot: C.white, page: C.white, burger: C.white });
  text(s, [0.846, o.titleY, 4.7, o.titleH], [o.title],
    { fontFace: F.sb, fontSize: o.titleSize, color: C.white });
  text(s, [0.982, 3.024, 1.362, 0.269], ['Premium Design'], { fontFace: F.md, fontSize: 10, color: C.white });
  rule(s, 1.048, 3.641, 5.91, 0, C.white);
  text(s, [0.982, 4.058, 1.995, 1.111], ['Business', 'Presentation', 'template.'],
    { fontFace: F.md, fontSize: 20, color: C.white });
  text(s, [0.985, 5.355, 0.942, 0.286], ['Template'], { fontFace: F.md, fontSize: 11, color: C.white });
  text(s, [2.869, 4.961, 3.212, 0.688], [[
    ['Lorep  ipsum duis '], ['aute irure ', { bold: true }],
    ['dolor in kauselih oilue reprehend esse cill inure dolorlaboru sit amet. Duis aute'],
    [' irusitakus ', { bold: true }], ['reprei. Voluptate lorem kuisais.'],
  ]], body(C.white));
  S(s, 'line', { x: 4.174, y: 6.144, w: 0, h: 0.488, line: { color: C.white, width: 0.5, endArrowType: 'triangle' } });
  text(s, [5.538, 2.557, 1.648, 0.808], ['Dui lorem duis ', 'Sit amet ', 'Doloris in ciu.'],
    { fontFace: F.md, fontSize: 14, color: C.white });
  text(s, [o.stampX, o.stampY, 0.709, 0.286],
    [[['01', { fontSize: 11 }], ['/', { fontSize: 10.5 }], ['03', { fontSize: 8 }]]],
    { fontFace: F.md, fontSize: 11, color: C.white });
}

const slide01 = s => coverSlide(s, { num: '01', title: 'FINA', titleSize: 96, titleY: 1.134, titleH: 1.717, stampX: 5.45, stampY: 2.1 });
const slide18 = s => coverSlide(s, { num: '18', title: 'BreakSlide', titleSize: 60, titleY: 1.197, titleH: 1.111, stampX: 5.004, stampY: 3.077, picture: PIC.slate });
const slide39 = s => coverSlide(s, { num: '39', title: 'Thank You', titleSize: 60, titleY: 1.094, titleH: 1.111, stampX: 4.576, stampY: 3.077 });

/* ---------------------------------------------------------------- slide 02 */
function slide02(s) {
  gradientBar(s);
  rect(s, 0.072, 5.969, 6.728, 1.531, C.orange);
  photo(s, 6.8, 1.853, 5.852, 5.647);
  chrome(s, { num: '02' , noFoot: true });
  rect(s, 6.8, 5.969, 1.031, 1.531, C.dark);
  text(s, [1.642, 2.005, 3.62, 1.447], ['Welcome to', 'Fina studio'], head);
  loremPair(s, 1.675, 3.587, C.ink);
  arrow(s, 7.205, 6.775, 0.233, 1);
  /* progress strip */
  rect(s, 1.067, 1.305, 11.015, 0.05, C.gD9);
  rect(s, 1.067, 1.305, 0.893, 0.056, C.dark);
  text(s, [0.977, 0.977, 1.371, 0.236], ['Lorem Ipsum'], { fontFace: F.md, fontSize: 8, color: C.ink });
  text(s, [3.455, 0.977, 1.233, 0.269], ['Text Tittle Here'], { fontFace: F.md, fontSize: 10, color: C.ink });
  text(s, [6.242, 0.691, 1.026, 0.555], ['Lorem ipsum', 'Dolor', 'Sit amet.'], { fontFace: F.md, fontSize: 9, color: C.ink });
  text(s, [8.234, 0.97, 0.935, 0.236], ['Modern Style'], { fontFace: F.md, fontSize: 8, align: 'right' });
  text(s, [10.137, 0.97, 0.958, 0.236], ['Art Company'], { fontFace: F.md, fontSize: 8, align: 'right' });
  text(s, [11.695, 0.659, 0.725, 0.64], ['01.'], { fontFace: F.md, fontSize: 32, color: C.ink });
  /* orange footer band */
  text(s, [0.826, 6.539, 0.935, 0.236], ['Modern Style'], { fontFace: F.md, fontSize: 8, color: C.white, align: 'right' });
  text(s, [2.007, 6.539, 0.958, 0.236], ['Art Company'], { fontFace: F.md, fontSize: 8, color: C.white, align: 'right' });
  text(s, [3.714, 6.568, 1.417, 0.269], ['Lorem duils'], { fontFace: F.md, fontSize: 10, color: C.white });
  text(s, [3.714, 6.849, 1.699, 0.286], ['Lorep  ipsum duis aute.'], body(C.white));
  S(s, 'mathMultiply', { x: 5.923, y: 6.616, w: 0.318, h: 0.318, fill: { color: C.white }, line: { type: 'none' } });
}

/* ---------------------------------------------------------------- slide 03 */
function slide03(s) {
  gradientBar(s);
  rect(s, 5.046, 0, 4.203, 0.989, C.orange);
  photo(s, 0.08, 1.735, 5.472, 5.765);
  chrome(s, { num: '03', footX: 0.312 });
  text(s, [7.161, 2.266, 3.287, 1.447], ['History of', 'Fina studio'], head);
  arrowTile(s, 4.775, 6.72, C.dark);
  pager(s, 11.847, 4.204, true);
  text(s, [5.753, 0.322, 0.914, 0.252], ['Text Tittle '], { fontFace: F.md, fontSize: 9, color: C.white });
  rule(s, 6.974, 0.448, 1.723, 0, C.white);
  text(s, [7.212, 1.947, 1.593, 0.269], [L.tt], { fontFace: F.sb, fontSize: 10, charSpacing: 0.5 });
  loremPair(s, 7.212, 3.89, C.ink);
  text(s, [7.212, 5.498, 3.446, 0.871],
    ['Lorep  ipsum duis aute irure dolor in koloris kauseliha oilusioi repiehenderiti volui ptate esse cilleisu dolrlasue boru sit amet. '],
    { fontFace: F.md, fontSize: 10.5, color: C.ink, lineSpacingMultiple: 1.5 });
}

/* ---------------------------------------------------------------- slide 04 */
function slide04(s) {
  gradientBar(s);
  rect(s, 7.179, 0.889, 4.897, 0.989, C.orange);
  rect(s, 2.367, 6.15, 1.984, 0.426, C.orange);
  chrome(s, {});
  text(s, [2.221, 1.283, 3.287, 1.447], ['Who start', 'Fina studio'], head);
  text(s, [2.678, 6.245, 1.514, 0.252], ['TEXT TITTLE '],
    { fontFace: F.sb, fontSize: 9, charSpacing: 3, color: C.white, align: 'center' });
  text(s, [7.55, 1.242, 1.454, 0.252], ['Text Tittle Here'], { fontFace: F.md, fontSize: 9, color: C.white });
  text(s, [9.086, 1.242, 1.454, 0.252], ['Text Tittle Here'], { fontFace: F.md, fontSize: 9, color: C.gF2 });
  rule(s, 10.662, 1.368, 1.015, 0, C.white);
  loremPair(s, 2.293, 3.012, C.g59);
  text(s, [2.293, 4.829, 2.188, 0.303], ['Text Tittle  Here'], { fontFace: F.md, fontSize: 12, color: C.ink });
  text(s, [2.293, 5.168, 1.983, 0.329],
    [[['California, ', { fontFace: F.sb }], ['03/14/1990', { fontFace: F.os }]]],
    { fontSize: 10, italic: true, color: C.g59, fontFace: F.sb, lineSpacingMultiple: 1.5 });
  text(s, [2.293, 5.398, 1.746, 0.284], ['CEO & Owner'], body(C.ink));
  photo(s, 7.179, 1.878, 6.154, 4.734);
}

/* ---------------------------------------------------------------- slide 05 */
function slide05(s) {
  gradientBar(s);
  photo(s, 7.646, 1.235, 5.031, 6.265);
  photo(s, 0.656, 1.235, 2.003, 2.304, PIC.light);
  chrome(s, { num: '05' });
  rect(s, 7.646, 0.74, 3.744, 0.989, C.dark);
  text(s, [3.132, 1.945, 4.041, 1.447], ['Fina', 'collaboration'], head);
  text(s, [8.024, 1.129, 2.125, 0.278], [L.tt],
    { fontFace: F.os, fontSize: 10, charSpacing: 3, color: C.white });
  arrow(s, 10.545, 1.234, 0.411, -1);
  text(s, [3.586, 1.566, 2.503, 0.286], ['Lorem Ipsum dolor sit amet .'], { fontFace: F.md, fontSize: 11 });
  S(s, 'rtTriangle', { x: 3.192, y: 1.645, w: 0.128, h: 0.128, rotate: 225, fill: { color: C.dark }, line: { type: 'none' } });
  rule(s, 0.77, 3.997, 6.403, 0, C.ink);
  text(s, [0.636, 4.414, 1.593, 0.269], [L.tt], { fontFace: F.sb, fontSize: 10, charSpacing: 0.5 });
  loremPair(s, 0.656, 4.803, C.g59);
  text(s, [4.344, 4.749, 3.06, 1.185], [L.e],
    { fontFace: F.md, fontSize: 11, color: C.ink, lineSpacingMultiple: 1.5 });
}

/* ---------------------------------------------------------------- slide 06 */
function slide06(s) {
  gradientBar(s);
  rect(s, 0.08, 6.578, 4.856, 0.941, C.dark);
  rect(s, 8.785, 5.25, 1.246, 2.269, C.orange);
  photo(s, 0.08, 0.882, 3.862, 5.677, PIC.light);
  photo(s, 10.185, 0.882, 3.149, 3.618);
  chrome(s, { num: '06', foot: C.white, footY: 6.964 });
  text(s, [4.805, 1.699, 4.888, 1.447], ['How start', '           Fina studio'], head);
  rect(s, 4.911, 2.755, 1.047, 0.081, C.mint);
  text(s, [4.805, 0.718, 1.763, 0.486], [L.duist, L.aure], body(C.g59));
  loremPair(s, 4.788, 3.317, C.g59);
  [4.785, 6.552].forEach(x => text(s, [x, 5.403, 1.651, 0.341], [L.li],
    { fontFace: F.md, fontSize: 10.5, color: C.ink, underline: { style: 'sng' }, lineSpacingMultiple: 1.5 }));
  /* rotated caption on the orange tab + arrow tile over the right photo */
  text(s, [8.936, 6.745, 0.943, 0.269], ['Text tittle'],
    { fontFace: F.os, fontSize: 10, charSpacing: 0.6, color: C.white, rotate: 270 });
  arrow(s, 9.273, 5.801, 0.208, -1);
  arrowTile(s, 10.184, 3.714, C.dark);
  text(s, [2.068, 6.964, 0.849, 0.269], ['Text tittle'], { fontFace: F.md, fontSize: 10, color: C.white });
  rule(s, 3.169, 7.116, 1.455, 0, C.white);
  [10.449, 11.935].forEach(x => text(s, [x, 5.549, 1.454, 0.252], ['Text Tittle Here'], { fontFace: F.md, fontSize: 9, color: C.dark }));
  [9.871, 11.355].forEach(x => text(s, [x, 6.335, 1.763, 0.486], [L.duist, L.aure],
    Object.assign(body(C.g59), { align: 'right' })));
}

/* ---------------------------------------------------------------- slide 07 */
function slide07(s) {
  gradientBar(s);
  rect(s, 4.187, 6.521, 5.825, 1.002, C.orange);
  photo(s, 1.04, 1.296, 4.683, 5.225);
  chrome(s, { num: '07' });
  rect(s, 4.446, 3.594, 2.71, 0.923, C.ink, { rotate: 270 });
  text(s, [4.738, 3.917, 2.125, 0.278], [L.tt],
    { fontFace: F.os, fontSize: 10, charSpacing: 3, color: C.white, rotate: 270 });
  text(s, [7.724, 1.296, 3.296, 1.447], ['When start', 'Fina studio'], head);
  text(s, [7.779, 2.979, 1.593, 0.269], [L.tt], { fontFace: F.sb, fontSize: 10, charSpacing: 0.5 });
  loremPair(s, 7.8, 3.368, C.g59);
  [7.872, 10.012].forEach(x => text(s, [x, 5.478, 1.651, 0.341], [L.li],
    { fontFace: F.md, fontSize: 10.5, color: C.ink, underline: { style: 'sng' }, lineSpacingMultiple: 1.5 }));
  pager(s, 5.819, 0.148, false);
  text(s, [4.517, 6.729, 1.372, 0.64], ['01/10'], { fontFace: F.md, fontSize: 32, color: C.white });
  rule(s, 7.555, 7.049, 1.758, 0, C.white);
  text(s, [6.001, 6.806, 1.397, 0.486], [L.duist, L.aure], body(C.white));
}

/* ---------------------------------------------------------------- slide 08 */
function slide08(s) {
  gradientBar(s);
  rect(s, 8.957, 6.747, 0.778, 0.78, C.dark);
  rect(s, 6.44, 4.298, 1.984, 0.426, C.orange);
  photo(s, 0.596, 0.753, 5.193, 3.025);
  photo(s, 6.44, 0.753, 2.518, 3.025, PIC.light);
  chrome(s, { num: '08' });
  text(s, [9.573, 1.217, 3.012, 1.447], ['About our', 'vision'], head);
  text(s, [9.448, 2.97, 3.446, 0.688], [L.a2, L.b], body(C.g59));
  text(s, [9.448, 3.77, 3.561, 0.688], [L.c], body(C.g59));
  arrow(s, 9.23, 7.151, 0.246, -1);
  text(s, [9.685, 3.984, 1.514, 0.252], ['TEXT TITTLE '],
    { fontFace: F.md, fontSize: 9, charSpacing: 3, color: C.white, align: 'center' });
  /* four mini captions */
  [[0.725, 4.304], [3.66, 4.304], [0.725, 5.709], [3.66, 5.709]].forEach(([x, y]) => {
    text(s, [x, y, 1.593, 0.286], [L.tt], kicker());
    text(s, [x, y + 0.354, 2.411, 0.688], [L.d], body(C.g59));
  });
  pager(s, 12.063, 1.386, true);
  ['5.288|' + C.gBF, '5.942|' + C.ink, '6.58|' + C.gBF].forEach(row => {
    const [y, col] = row.split('|');
    text(s, [10.957, +y, 1.651, 0.341], [L.li], { fontFace: F.md, fontSize: 10.5, color: col, lineSpacingMultiple: 1.5 });
  });
  rule(s, 10.454, 6.144, 0.426, 0, C.plum);
  text(s, [6.17, 5.709, 3.06, 1.185], [L.e], { fontFace: F.md, fontSize: 11, lineSpacingMultiple: 1.5 });
}

/* ---------------------------------------------------------------- slide 09 */
function slide09(s) {
  rect(s, 10.965, 0, 2.368, 7.5, C.orange);
  gradientBar(s);
  photo(s, 0.944, 0.596, 1.976, 1.939);
  photo(s, 7.736, 2.868, 4.413, 3.238, PIC.light);
  chrome(s, { num: '09', page: C.white, burger: C.white });
  text(s, [4.49, 0.948, 2.368, 0.64], ['Lorem ipsum sit ', 'amet doloris'],
    { fontFace: F.md, fontSize: 16, charSpacing: 0.5, color: C.ink });
  text(s, [4.531, 1.783, 3.474, 0.688],
    ['Lorep  ipsum duis aute irure dolor in kauselih oirehe',
      'deriti vols esse cill inure dolorlaboru sit amet. Duilo irusitakus reprehenderi Volupem kuisais.'], body(C.g59));
  text(s, [1.023, 2.896, 3.012, 1.447], ['About our', 'mission'], head);
  navPair(s, 4.561, 3.726);
  [[1.023, 4.681], [4.035, 4.674], [1.023, 5.996], [4.035, 5.996]].forEach(([x, y]) => {
    text(s, [x, y, 1.593, 0.286], [L.tt], kicker());
    text(s, [x, y + 0.305, 2.233, 0.486],
      ['Lorep  ipsum duis aute irure  kause', ' oilusioi reprehenderitilores volupta'], body(C.g59));
  });
  pager(s, 2.596, 1.303, true);
  [3.643, 5.371].forEach(y => text(s, [12.099, y, 1.454, 0.252], ['Text Tittle Here'],
    { fontFace: F.md, fontSize: 9, color: C.white, rotate: 90 }));
  text(s, [9.105, 1.878, 1.763, 0.486], [L.duist, L.aure], body(C.g59));
}

/* ---------------------------------------------------------------- slide 10 */
function slide10(s) {
  gradientBar(s);
  chrome(s, { num: '10' });
  rect(s, 1.632, 0.971, 0.89, 0.088, C.plum);
  text(s, [1.544, 1.104, 2.368, 0.284], ['Lorep  ipsum duis aute irure dolor'], body(C.g59));
  text(s, [1.544, 1.44, 3.0, 0.774], ['Our Team'], head);
  const team = [[1.596, 'Bagus Adi', 1.732, PIC.light], [4.219, 'Ragile Soebekti', 4.408, PIC.deep],
    [6.842, 'Bagas Dimas', 7.083, PIC.gray], [9.465, 'Sapto Bayu', 9.758, PIC.light]];
  team.forEach(([px, name, tx, tone], i) => {
    photo(s, px, 2.498, 2.333, 3.333, tone);
    rect(s, px + 0.013, 5.452, 2.334, 0.377, C.dark);
    text(s, [tx, 6.109 - i * 0.005, 1.581, 0.269], [name], { fontFace: F.sb, fontSize: 10, color: C.ink });
    text(s, [tx, 6.269 - i * 0.005, 1.983, 0.261],
      [[['California, ', { fontFace: F.sb }], ['03/14/1990', { fontFace: F.os }]]],
      { fontSize: 7, italic: true, color: C.g59, fontFace: F.sb, lineSpacingMultiple: 1.5 });
  });
}

/* ---------------------------------------------------------------- slide 11 */
function slide11(s) {
  gradientBar(s);
  photo(s, 3.649, 2.162, 3.763, 2.842);
  photo(s, 7.702, 2.162, 3.763, 2.842, PIC.deep);
  photo(s, 11.754, 2.162, 1.579, 2.842, PIC.light);
  chrome(s, { num: '11' });
  text(s, [0.672, 0.789, 2.909, 0.774], ['Our team'], head);
  text(s, [4.256, 0.896, 3.446, 0.688],
    ['Lorep  ipsum duis aute irure dolor in kauselih  eprehe',
      'deriti vols esse cill inure dolorlaboru sit amuis autelo irusitakus reprehenderi Voluptate loisais.'], body(C.g80));
  [[4.807, 5.09, 6.863, 3.815], [8.86, 9.141, 10.915, 8.284]].forEach(([bx, tx, px, cx]) => {
    rect(s, bx, 4.237, 2.605, 0.768, C.orange);
    text(s, [tx, 4.491, 1.651, 0.284], ['LOREM IPSUM'],
      { fontFace: F.md, fontSize: 8, charSpacing: 3, color: C.white, lineSpacingMultiple: 1.5 });
    S(s, 'mathPlus', { x: px, y: 4.45, w: 0.368, h: 0.368, fill: { color: C.white }, line: { type: 'none' } });
    text(s, [cx, 5.3, 1.581, 0.269], ['Bagus Adi'], { fontFace: F.sb, fontSize: 10, color: C.ink });
    text(s, [cx, 5.46, 1.983, 0.261],
      [[['California, ', { fontFace: F.sb }], ['03/14/1990', { fontFace: F.os }]]],
      { fontSize: 7, italic: true, color: C.g59, fontFace: F.sb, lineSpacingMultiple: 1.5 });
    rule(s, cx + 0.08, 5.86, 0.992, 0, C.plum);
    text(s, [cx, 5.971, 3.446, 0.688],
      ['Lorep  ipsum duis aute irure dolor in kauselih  eprehe',
        'deriti vols esse cill inure dolorlaboru sit amuis autelo irusitakus reprehenderi Voluptate loisais.'], body(C.g80));
  });
  [[3.272, C.gBF], [3.927, C.ink], [4.564, C.gBF]].forEach(([y, col]) =>
    text(s, [1.051, y, 1.651, 0.341], [L.li], { fontFace: F.md, fontSize: 10.5, color: col, lineSpacingMultiple: 1.5 }));
  rule(s, 0.548, 4.128, 0.426, 0, C.plum);
}

/* ---------------------------------------------------------------- slide 12 */
function slide12(s) {
  rect(s, 6.698, 2.262, 6.636, 5.238, C.orange);
  gradientBar(s);
  chrome(s, { num: '12', page: C.white });
  text(s, [1.335, 1.569, 2.705, 1.447], ['Meet our', 'team'], head);
  text(s, [1.335, 3.15, 3.446, 0.688], [L.a2, L.b], body(C.g59));
  text(s, [1.335, 3.95, 3.561, 0.688], [L.c], body(C.g59));
  navPair(s, 1.458, 5.342);
  pager(s, 4.965, 6.081, true);
  /* photo mosaic on the orange field */
  [[7.62, 0.514, 1.764, PIC.deep], [6.702, 2.26, 1.752, PIC.light], [8.49, 2.26, 1.764, PIC.light],
    [10.256, 2.26, 1.764, PIC.sage], [8.478, 4.006, 1.764, PIC.gray], [6.702, 4.008, 1.764, PIC.slate],
    [6.683, 5.754, 1.764, PIC.slate]].forEach(([x, y, w, tone]) => photo(s, x, y, w, 1.746, tone));
  [[10.536, 4.638, 5.077], [8.895, 6.075, 6.513]].forEach(([x, ty, py]) => {
    text(s, [x, ty, 1.651, 0.341], ['Text Tittle Here'], { fontFace: F.md, fontSize: 10.5, color: C.white, lineSpacingMultiple: 1.5 });
    text(s, [x, py, 1.763, 0.486], [L.duist, L.aure], body(C.white));
  });
}

/* ---------------------------------------------------------------- slide 13 */
function slide13(s) {
  gradientBar(s);
  chrome(s, { num: '13' });
  photo(s, 1.138, 1.285, 5.111, 4.93, PIC.slate);
  text(s, [7.772, 1.285, 3.254, 0.774], ['Fina Expert'], head);
  text(s, [7.85, 2.194, 2.188, 0.337], ['Jadaey Ryan'], { fontFace: F.sb, fontSize: 14, color: C.g40 });
  text(s, [7.85, 2.528, 1.983, 0.352],
    [[['California, ', { fontFace: F.sb }], ['03/14/1990', { fontFace: F.os }]]],
    { fontSize: 10.5, italic: true, color: C.g40, fontFace: F.sb, lineSpacingMultiple: 1.5 });
  text(s, [7.85, 2.758, 1.746, 0.307], ['Illustrator'], Object.assign(body(C.g40), { fontSize: 9 }));
  text(s, [7.802, 3.151, 0.495, 1.01], ['\u201C'], { fontFace: F.sb, fontSize: 54, color: C.g40 });
  text(s, [7.85, 3.453, 4.709, 0.63], [L.quote],
    { fontFace: F.md, fontSize: 11, italic: true, charSpacing: 0.5, color: C.g40, lineSpacingMultiple: 1.5 });
  text(s, [7.833, 4.189, 4.52, 0.688],
    ['Lorep  ipsum duis aute irure dolor in kauselih oilusioi repreheni voluptates esse cill inure dolorlaboru sit amet. Duis aute irusita dolorin reprehenderit insa voluptate.'],
    body(C.g40));
  /* skill bars */
  [[5.131, 2.142, 9.595, 1.13, 'PSD ', 11.137, 5.021, 0.778],
    [5.593, 2.54, 9.595, 1.13, 'AI', 11.228, 5.476, 0.597],
    [6.054, 2.142, 9.137, 1.588, 'CDR', 11.074, 5.937, 0.597]].forEach(r => {
    rect(s, 7.895, r[0], r[1], 0.05, C.plum);
    rect(s, r[2], r[0], r[3], 0.05, C.gD9);
    text(s, [r[5], r[6], r[7], 0.278], [r[4]], { fontFace: F.ws, fontSize: 10.5, color: C.g40 });
  });
}

/* ---------------------------------------------------------------- slide 14 */
function slide14(s) {
  gradientBar(s);
  chrome(s, { num: '14' });
  photo(s, 7.667, 2.881, 3.92, 3.208, PIC.slate);
  text(s, [7.56, 0.961, 3.443, 0.774], ['Our service'], head);
  text(s, [7.56, 1.965, 4.027, 0.688], [L.f], body(C.g80));
  rect(s, 7.667, 5.575, 2.113, 0.515, C.orange);
  text(s, [7.976, 5.69, 1.651, 0.284], ['LOREM IPSUM'],
    { fontFace: F.md, fontSize: 8, charSpacing: 3, color: C.white, lineSpacingMultiple: 1.5 });
  const rows = [[1.735, 1.965, 2.213, 1.75, 1.78, 0.34, 0.4], [3.426, 3.655, 2.213, 1.66, 3.42, 0.4, 0.44],
    [5.112, 5.402, 2.156, 1.633, 5.112, 0.463, 0.463]];
  rows.forEach(([ty, py, x, gx, gy, gw, gh]) => {
    glyph(s, gx, gy, gw, gh, C.ink);
    text(s, [x, ty, 2.173, 0.269], [L.tt], { fontFace: F.md, fontSize: 10, charSpacing: 0.5, color: C.ink });
    text(s, [x, py, 4.671, 0.688],
      ['Lorep  ipsum duis aute irure dolor in kauselih oilue reprehenderiti vols esse cill inure dolorlaboru sit at. aute irusitakus reprehenderiVoluptate lores kausiesuli',
        'Lorem dolor sit amet.'], body(C.g80));
  });
}

/* ---------------------------------------------------------------- slide 15 */
function slide15(s) {
  gradientBar(s);
  S(s, 'ellipse', { x: 5.456, y: 3.622, w: 1.254, h: 1.254, fill: { color: C.orange }, line: { type: 'none' } });
  chrome(s, { num: '15', burger: C.white });
  photo(s, 6.395, 0, 6.938, 3.37, PIC.slate);
  text(s, [1.306, 1.044, 3.443, 0.774], ['Our service'], head);
  text(s, [1.306, 2.03, 4.027, 0.688], [L.f], body(C.g80));
  text(s, [1.306, 2.885, 4.027, 0.486],
    ['Lorep  ipsum duis aute irure dolor in kauselih oilue eprehe deriti vols esse cill inure dolorlaboru sit amet. Duis autelo irusitakus reprelikasi.'], body(C.g80));
  [1.306, 5.456, 9.606].forEach(x => {
    text(s, [x, 5.219, 1.593, 0.286], [L.tt], kicker());
    rule(s, x + 0.1, 5.654, 2.203, 0, C.g80);
    text(s, [x, 5.805, 2.303, 0.688], [L.d], body(C.g80));
  });
  glyph(s, 2.03, 4.06, 0.41, 0.41, C.ink);
  glyph(s, 5.79, 4.05, 0.56, 0.45, C.white);
  glyph(s, 10.3, 4.07, 0.45, 0.45, C.ink);
  [[3.609, 4.359], [8.199, 4.341]].forEach(([x, y]) => S(s, 'line',
    { x, y, w: 0.645, h: 0, line: { color: C.plum, width: 0.5, beginArrowType: 'triangle', endArrowType: 'triangle' } }));
}

/* ---------------------------------------------------------------- slide 16 */
function slide16(s) {
  gradientBar(s);
  [[9.748, 4.314], [6.23, 4.295], [9.748, 1.859], [6.23, 1.814]].forEach(([x, y]) =>
    S(s, 'ellipse', { x, y, w: 0.976, h: 0.976, fill: { color: C.orange }, line: { type: 'none' } }));
  chrome(s, { num: '16' });
  text(s, [1.306, 1.044, 3.443, 0.774], ['Our service'], head);
  text(s, [1.306, 2.03, 4.027, 0.688], [L.f], body(C.g80));
  photo(s, 1.412, 2.93, 3.814, 3.526, PIC.slate);
  [[6.015, 3.013], [9.676, 3.013], [6.015, 5.499], [9.676, 5.499]].forEach(([x, y]) => {
    text(s, [x, y, 1.593, 0.286], [L.tt], { fontFace: F.md, fontSize: 10.5, charSpacing: 0.5, color: C.ink });
    text(s, [x, y + 0.286, 2.799, 0.688],
      ['Lorep  ipsum duis aute irure dolor in kauselih oilusioi reprehenderiti volui ptates esse cill inure dolorlasue boru sit amet. Duis aute'], body(C.g80));
  });
  glyph(s, 6.407, 2.038, 0.619, 0.619, C.white);
  glyph(s, 9.943, 2.051, 0.619, 0.619, C.white);
  glyph(s, 6.44, 4.49, 0.55, 0.54, C.white);
  glyph(s, 9.96, 4.64, 0.41, 0.38, C.white);
}

/* ---------------------------------------------------------------- slide 17 */
function slide17(s) {
  gradientBar(s);
  chrome(s, { num: '17' });
  photo(s, 1.823, 2.947, 3.847, 3.509, PIC.slate);
  text(s, [1.755, 1.044, 3.443, 0.774], ['Our service'], head);
  text(s, [1.755, 2.03, 4.027, 0.688], [L.f], body(C.g80));
  text(s, [6.415, 1.651, 0.495, 1.01], ['\u201C'], { fontFace: F.sb, fontSize: 54, color: C.g40 });
  text(s, [6.462, 1.954, 4.709, 0.63], [L.quote],
    { fontFace: F.md, fontSize: 11, italic: true, charSpacing: 0.5, color: C.g40, lineSpacingMultiple: 1.5 });
  [[6.462, 5.186], [9.321, 5.182]].forEach(([x, y]) => {
    text(s, [x, y, 1.593, 0.286], [L.tt], kicker());
    rule(s, x + 0.1, y + 0.435, 2.203, 0, C.g80);
    text(s, [x, y + 0.586, 2.303, 0.688], [L.d], body(C.g80));
  });
  glyph(s, 6.53, 3.402, 1.397, 1.397, C.ink);
  glyph(s, 9.751, 3.677, 0.847, 1.024, C.ink);
  arrowTile(s, 4.9, 5.676, C.orange);
}

/* ---------------------------------------------------------------- slide 19 */
function slide19(s) {
  gradientBar(s);
  rect(s, 7.87, 6.63, 5.463, 0.87, C.orange);
  rect(s, 6.487, 5.492, 2.068, 2.008, C.ink);
  photo(s, 0.08, 1.441, 4.339, 6.059, PIC.mist);
  photo(s, 4.419, 5.492, 2.068, 2.008, PIC.slate);
  photo(s, 9.679, 0.735, 3.655, 2.008, PIC.sage);
  chrome(s, { num: '19', foot: C.white, page: C.white });
  text(s, [5.296, 0.481, 3.885, 1.447], ['Fina ', 'Photography'], head);
  text(s, [5.296, 2.089, 3.446, 0.688], [L.a2, L.b], body(C.g59));
  text(s, [5.296, 2.889, 3.561, 0.688], [L.c], body(C.g59));
  text(s, [5.623, 3.92, 2.503, 0.286], ['Lorem Ipsum dolor sit amet .'], { fontFace: F.md, fontSize: 11 });
  S(s, 'rtTriangle', { x: 5.229, y: 3.999, w: 0.128, h: 0.128, rotate: 225, fill: { color: C.dark }, line: { type: 'none' } });
  text(s, [9.694, 4.643, 1.807, 0.269], ['Lorem ipsum dolor sit'], { fontFace: F.md, fontSize: 10, color: C.dark });
  text(s, [9.694, 5.124, 2.43, 0.236], ['Lorem ipsum dolor sit amet'], { fontFace: F.md, fontSize: 8 });
  text(s, [9.89, 7.036, 2.43, 0.236], ['Lorem ipsum dolor sit amet'], { fontFace: F.md, fontSize: 8, color: C.white });
  text(s, [6.69, 5.895, 0.738, 0.739], ['/01'],
    { fontFace: F.md, fontSize: 11, color: C.white, align: 'center', lineSpacingMultiple: 2.5 });
  rule(s, 6.912, 6.413, 1.331, 0, C.gF2);
  text(s, [6.817, 6.534, 1.454, 0.236], ['Text Tittle Here'], { fontFace: F.md, fontSize: 8, color: C.white });
  text(s, [6.817, 6.721, 1.763, 0.341], [L.duist, L.aure], Object.assign(body(C.gD9), { fontSize: 5 }));
  pager(s, 11.458, 4.165, true);
}

/* ---------------------------------------------------------------- slide 20 */
function slide20(s) {
  rect(s, 0.08, 5.12, 7.014, 2.38, C.dark);
  rect(s, 10.201, 5.12, 3.132, 2.38, C.orange);
  gradientBar(s);
  photo(s, 9.016, 0, 4.317, 5.131, PIC.mist);
  photo(s, 0.523, 4.213, 1.862, 1.891, PIC.slate);
  photo(s, 2.692, 4.213, 1.862, 1.891, PIC.slate);
  photo(s, 4.845, 4.213, 1.862, 1.891, PIC.sage);
  photo(s, 7.094, 5.12, 3.132, 2.38, PIC.mist);
  chrome(s, { num: '20', page: C.white, noFoot: true, burger: C.white });
  text(s, [0.735, 0.803, 3.692, 1.447], ['The best art ', 'from maker'], head);
  text(s, [0.8, 2.387, 3.446, 0.688], [L.a, L.b], body(C.g59));
  text(s, [0.8, 3.153, 3.561, 0.688], [L.c], body(C.g59));
  rule(s, 4.659, 3.075, 2.257, 0, C.black);
  text(s, [4.61, 3.18, 2.751, 0.678], ['Lorem Ipsum dolor sit amet duis kaesius.'],
    { fontFace: F.md, fontSize: 12, color: C.ink, lineSpacingMultiple: 1.5 });
  text(s, [10.691, 5.85, 2.125, 0.278], [L.tt], { fontFace: F.os, fontSize: 10, charSpacing: 3, color: C.white });
  text(s, [10.691, 6.286, 2.152, 0.685],
    ['Lorep  ipsum duis aute irure dolor in kauselih oilue esse cill inure dolorlaboru sit amet.'], body(C.gF2));
  text(s, [0.931, 6.251, 1.08, 0.236], ['LOREM IPSUM'], { fontFace: F.rg, fontSize: 8, charSpacing: 1.1, color: C.white });
  text(s, [0.931, 6.56, 2.503, 0.614],
    ['Lorep  ipsum duis aute irure dolor in kaselih oilue reprehendesse cill inure dolorlaoru sit amet. Duis aute irusitakus.'],
    Object.assign(body(C.gD9), { fontSize: 7, italic: true }));
  text(s, [4.845, 6.939, 1.026, 0.236], ['Text Tittle Here'], { fontFace: F.md, fontSize: 8, color: C.white });
  arrow(s, 3.524, 7.049, 0.681, 1);
}

/* ---------------------------------------------------------------- slide 21 */
function slide21(s) {
  gradientBar(s);
  rect(s, 9.171, 3.962, 2.439, 2.802, C.orange);
  photo(s, 0.08, 0, 4.994, 7.5, PIC.slate);
  photo(s, 6.732, 3.962, 2.439, 2.802, PIC.mist);
  chrome(s, { num: '21', top: C.white, foot: C.white });
  text(s, [7.024, 0.498, 3.654, 0.774], ['Fiana Works'], head);
  text(s, [7.077, 1.377, 1.937, 0.512], ['Lorep  ipsum '],
    { fontFace: F.md, fontSize: 18, underline: { style: 'sng' }, lineSpacingMultiple: 1.5 });
  text(s, [7.077, 2.057, 3.446, 0.688], [L.a2, L.b], body(C.g80));
  text(s, [7.062, 2.846, 3.561, 0.688], [L.c], body(C.g80));
  navPair(s, 4.403, 6.686);
  pager(s, 4.949, 4.583, true);
  [[3.962, C.gBF], [4.617, C.ink], [5.254, C.gBF]].forEach(([y, col]) =>
    text(s, [12.085, y, 1.651, 0.341], [L.li], { fontFace: F.md, fontSize: 10, color: col, lineSpacingMultiple: 1.5 }));
  rule(s, 11.836, 4.818, 0.172, 0, C.plum);
  text(s, [9.36, 4.048, 0.738, 0.739], ['/01'],
    { fontFace: F.md, fontSize: 11, color: C.white, align: 'center', lineSpacingMultiple: 2.5 });
  rule(s, 9.582, 4.566, 1.331, 0, C.gF2);
  text(s, [9.494, 4.645, 1.723, 0.421], ['Text Tittle  Here'], { fontFace: F.md, fontSize: 14, color: C.white, lineSpacingMultiple: 1.5 });
  text(s, [9.494, 5.254, 1.884, 1.145],
    ['Lorep  ipsum duis aute irureui dolor in kauselih oilue eprehe',
      'deriti vols esse cill inure loem dolorlaboru sit amet. Duis sira autelo irusitakus reprehenderi Voluptate lorem kuisais.'],
    Object.assign(body(C.gF2), { fontSize: 7 }));
}

/* ---------------------------------------------------------------- slide 22 */
function slide22(s) {
  gradientBar(s);
  chrome(s, { num: '22' });
  text(s, [4.444, 0.231, 4.18, 0.774], ['Gallery layout'], Object.assign({ align: 'center' }, head));
  const tone = [[PIC.gray, PIC.light, PIC.mist], [PIC.mist, PIC.slate, PIC.sage]];
  [1.428, 4.223].forEach((py, r) => [0.976, 4.988, 9.0].forEach((px, c) => photo(s, px, py, 3.357, 1.714, tone[r][c])));
  [[3.334, 3.393, 3.67], [6.141, 6.2, 6.477]].forEach(([ly, ty, py]) => {
    [1.225, 5.216, 9.206].forEach(x => {
      rule(s, x + 0.099, ly, 0.387, 0, C.ink);
      text(s, [x, ty, 1.499, 0.278], [L.tt], { fontFace: F.sb, fontSize: 10, charSpacing: 0.5, color: C.ink });
      text(s, [x, py, 2.378, 0.39], [[
        ['Lorep  ipsum duis aute irure dolor in kaelih oilue eprehe '],
        ['deriti vols escill inure do.'],
      ]], Object.assign(body(C.g59), { fontSize: 6 }));
    });
  });
}

/* ---------------------------------------------------------------- slide 23 */
function slide23(s) {
  gradientBar(s);
  chrome(s, { num: '24', foot: C.white });
  text(s, [1.318, 0.697, 4.18, 0.774], ['Gallery layout'], Object.assign({ align: 'center' }, head));
  text(s, [1.479, 1.557, 3.981, 0.981],
    ['Lorep  ipsum duis aute irure dolor in kaseliha oilue epreh deriti vols cill inure dolorlru sitam amet. Duis autelo irusitakus.'],
    Object.assign(body(C.g80), { fontSize: 12 }));
  text(s, [6.422, 0.761, 2.828, 0.71], ['Lorem ipsum dolor', 'Sit amet'],
    { fontFace: F.md, fontSize: 20, color: C.ink, lineSpacingMultiple: 0.9 });
  text(s, [6.422, 1.618, 2.828, 0.89],
    ['Lorep  ipsum duis aute irure dolor in kaselih oilue reprehendesse cill inure dolorlaoru sit amet. Duis aute irusitakus reprei Voluptate lorem kuisais.'], body(C.g59));
  text(s, [9.377, 1.656, 1.026, 0.555], ['Lorem ipsum', 'Dolor', 'Sit amet.'], { fontFace: F.md, fontSize: 9, color: C.ink });
  [[0.077, 2.912, 3.403, 2.353, PIC.sage], [3.48, 2.912, 6.45, 2.353, PIC.mist], [9.93, 2.912, 3.403, 2.353, PIC.mist],
    [0.077, 5.265, 3.403, 2.235, PIC.slate], [3.48, 5.265, 6.45, 2.235, PIC.gray], [9.93, 5.265, 3.403, 2.235, PIC.light]]
    .forEach(p => photo(s, p[0], p[1], p[2], p[3], p[4]));
}

/* ------------------------------------------------- slides 24-27: mock-ups */
/** shared left column of the four device slides */
function deviceSlide(s, num, title) {
  gradientBar(s);
  navPair(s, 1.499, 5.145);
  chrome(s, { num });
  text(s, [1.316, 1.447, 4.6, 1.447], title, head);
  text(s, [1.386, 3.135, 3.446, 0.688], [L.a2, L.b], body(C.g80));
  text(s, [1.371, 3.923, 3.561, 0.688], [L.c], body(C.g80));
}

function slide24(s) {
  deviceSlide(s, '25', ['Dekstop', 'media mockup']);
  device(s, [7.15, 1.94, 4.79, 3.05], [7.31, 2.1, 4.47, 2.73], 0.06);
  rect(s, 7.15, 4.86, 4.79, 0.3, 'B8BBBD');
  rect(s, 9.06, 5.16, 0.97, 0.6, 'C8CBCD');
  S(s, 'ellipse', { x: 8.72, y: 5.66, w: 1.65, h: 0.36, fill: { color: 'D5D8DA' }, line: { type: 'none' } });
}

function slide25(s) {
  deviceSlide(s, '26', ['Leptop device', 'slide']);
  device(s, [7.44, 2.19, 4.65, 3.0], [7.57, 2.32, 4.38, 2.74], 0.05);
  S(s, 'trapezoid', { x: 6.955, y: 5.19, w: 5.601, h: 0.24, fill: { color: 'C8CBCD' }, line: { type: 'none' }, flipV: true });
}

function slide26(s) {
  deviceSlide(s, '27', ['Leptop device', 'slide']);
  device(s, [6.879, 1.494, 2.44, 4.739], [7.0, 1.636, 2.168, 4.43, PIC.gray], 0.22);
  device(s, [10.134, 1.494, 2.44, 4.739], [10.251, 1.636, 2.168, 4.43, PIC.light], 0.22);
}

function slide27(s) {
  deviceSlide(s, '28', ['Tablet device', 'slide']);
  device(s, [7.806, 1.346, 3.466, 5.303], [7.976, 1.894, 3.095, 4.274], 0.12, C.dark);
  S(s, 'ellipse', { x: 8.428, y: 6.24, w: 0.26, h: 0.26, fill: { color: '3D3D3D' }, line: { type: 'none' } });
}

/* ---------------------------------------------------------------- slide 28 */
function slide28(s) {
  gradientBar(s);
  chrome(s, { num: '29' });
  text(s, [0.968, 0.981, 3.522, 0.774], ['Infographic'], head);
  /* alternating semicircular ribbons */
  const arc = (x, color, rot) => S(s, 'blockArc', {
    x, y: 2.742, w: 2.086, h: 2.086, rotate: rot,
    angleRange: [180, 359.97], arcThicknessRatio: 0.315,
    fill: { color }, line: { type: 'none' },
  });
  [9.431, 2.407, 5.918].forEach(x => arc(x, C.gCC, 180));
  [4.164, 7.675].forEach(x => arc(x, C.gCC, 0));
  [2.409, 5.92, 9.431].forEach(x => arc(x, C.teal, 0));
  [4.166, 7.677].forEach(x => arc(x, C.dark, 180));
  /* icon discs and their numbered badges */
  const discs = [[2.856, C.orange, '1', 3.219, 2.499, C.orange], [4.613, C.gA6, '2', 4.974, 4.585, C.teal],
    [6.369, C.orange, '3', 6.73, 2.499, C.orange], [8.125, C.gA6, '4', 8.485, 4.585, C.teal],
    [9.882, C.orange, '5', 10.241, 2.499, C.orange]];
  discs.forEach(([x, col, n, bx, by, bcol]) => {
    S(s, 'ellipse', { x, y: 3.188, w: 1.194, h: 1.194, fill: { color: col }, line: { color: C.white, width: 1 } });
    glyph(s, x + 0.3, 3.488, 0.594, 0.594, C.white);
    S(s, 'ellipse', { x: bx, y: by, w: 0.466, h: 0.466, fill: { color: bcol }, line: { color: C.white, width: 3 } });
    text(s, [bx, by + 0.06, 0.466, 0.35], [n], { fontFace: 'Open Sans', fontSize: 14, bold: true, color: C.white, align: 'center' });
  });
  [[2.692, 5.323], [4.908, 1.149], [6.132, 5.313], [8.385, 1.142], [10.128, 5.265]].forEach(([x, y]) =>
    caption(s, x, y, 2.466, null, null, { bodyColor: C.g40, gap: 0.508 }));
}

/* ---------------------------------------------------------------- slide 29 */
function slide29(s) {
  gradientBar(s);
  chrome(s, { num: '29', noFoot: true });
  /* nested sectors of a disc anchored at the bottom-left corner */
  const disc = (r, from, to, color) => S(s, 'pie', {
    x: 4.197 - r, y: 7.5 - r, w: r * 2, h: r * 2, angleRange: [from, to],
    fill: { color }, line: { type: 'none' },
  });
  disc(4.068, 180, 360, C.dark);
  disc(4.068, 182, 240, C.orange);
  disc(2.903, 182, 330, C.orange);
  disc(1.743, 182, 358, C.orange);
  S(s, 'ellipse', { x: 3.897, y: 7.2, w: 0.6, h: 0.6, fill: { color: C.orange }, line: { type: 'none' } });
  glyph(s, 1.802, 4.53, 0.573, 0.4, C.white);
  glyph(s, 4.365, 5.076, 0.432, 0.432, C.white);
  glyph(s, 4.976, 6.772, 0.518, 0.518, C.white);
  [[1.599, 3.002, 2.599, 4.498], [4.197, 3.979, 2.004, 3.521], [4.75, 7.44, 2.903, 0.001]].forEach(([x, y, w, h]) =>
    S(s, 'line', { x, y, w, h, line: { color: C.orange, width: 6, endArrowType: 'triangle' }, flipV: true }));
  text(s, [8.985, 1.737, 3.522, 0.774], ['Infographic'], Object.assign({ align: 'center' }, head));
  text(s, [9.082, 1.392, 2.471, 0.303], ['BUSINESS IDEA'], { fontFace: F.md, fontSize: 12, charSpacing: 3, color: C.ink });
  text(s, [9.082, 2.653, 2.608, 0.89],
    ['Lorep  ipsum duis aute irure dolor in kausel  oilue eprehe deriti vols esse cill inure dolori laboru sit amet. Duis autelo irusitakus repre henderi Voluptate lorem kuisais.'], body(C.g80));
  text(s, [9.082, 3.681, 2.608, 0.89],
    ['Lorep  ipsum duis aute irure dolor in kausel  oilue eprehe deriti vols esse cill inure dolori laboru sit amet. Duis autelo irusitakus repre henderi Voluptate lorem kuisais. Sit ametai'], body(C.g80));
  [[0.659, 1.121], [5.354, 1.695], [9.066, 5.402]].forEach(([x, y]) => {
    text(s, [x, y, 1.026, 0.555], ['Lorem ipsum', 'Dolor', 'Sit amet.'], { fontFace: F.md, fontSize: 9, color: C.dark });
    rule(s, x + 0.076, y + 0.796, 2.072, 0, C.dark);
    text(s, [x, y + 1.0, 2.503, 0.791], [[
      ['Lorep  ipsum duis '], ['aute irure ', { bold: true }],
      ['dolor in kaselih oilue reprehendesse cill inure dolorlaoru sit amet. Duis aute'],
      [' irusitakus ', { bold: true }], ['reprei Voluptate lorem kuisais.'],
    ]], Object.assign(body(C.g80), { fontSize: 7 }));
  });
}

/* ---------------------------------------------------------------- slide 30 */
function slide30(s) {
  gradientBar(s);
  chrome(s, { num: '30' });
  text(s, [4.861, 0.665, 3.522, 0.774], ['Infographic'], head);
  /* pinwheel of four sectors around a grey hub */
  const R = 2.59, cx = 6.667, cy = 4.303;
  [[200, 277, C.orange], [290, 367, C.dark], [20, 97, C.orange], [110, 187, C.dark]].forEach(([a, b, col]) =>
    S(s, 'pie', { x: cx - R, y: cy - R, w: R * 2, h: R * 2, angleRange: [a, b], fill: { color: col }, line: { type: 'none' } }));
  S(s, 'ellipse', { x: 5.883, y: 3.519, w: 1.568, h: 1.572, fill: { color: C.gCC }, line: { color: C.white, width: 6 } });
  [[6.12, 1.913, C.dark], [7.861, 3.706, C.orange], [6.065, 5.52, C.dark], [4.284, 3.666, C.orange]].forEach(([x, y, col]) => {
    S(s, 'ellipse', { x, y, w: 1.194, h: 1.194, fill: { color: col }, line: { color: C.white, width: 3 } });
    glyph(s, x + 0.3, y + 0.3, 0.594, 0.594, C.white);
  });
  [['01', 5.24, 2.419, 0.619], ['02', 7.806, 2.864, 0.745], ['03', 7.529, 5.472, 0.763], ['04', 4.845, 5.152, 0.766]]
    .forEach(([n, x, y, w]) => text(s, [x, y, w, 0.64], [n],
      { fontFace: F.md, fontSize: 32, color: C.white, align: 'center', wrap: false }));
  caption(s, 1.337, 2.358, 2.579);
  caption(s, 1.337, 4.838, 2.579);
  caption(s, 9.508, 2.456, 2.579, null, null, { right: true });
  caption(s, 9.479, 4.934, 2.579, null, null, { right: true });
}

/* ---------------------------------------------------------------- slide 31 */
function slide31(s) {
  gradientBar(s);
  chrome(s, { num: '31' });
  text(s, [5.657, 0.948, 2.02, 0.337], [[['Business ', {}], ['System', { bold: true }]]], { fontFace: F.rg, fontSize: 14 });
  text(s, [5.07, 1.407, 3.193, 0.651], ['Infographic'],
    { fontFace: F.md, fontSize: 36, color: C.orange, align: 'center', lineSpacingMultiple: 0.9 });
  /* looping track behind the chevrons */
  S(s, 'roundRect', { x: 2.8, y: 4.458, w: 8.108, h: 2.663, rectRadius: 1.2, fill: { type: 'none' }, line: { color: C.gD0, width: 6 } });
  arrow(s, 2.686, 5.9, 0.001, -1, C.dark);
  S(s, 'line', { x: 4.335, y: 6.98, w: 2.28, h: 0, line: { color: C.orange, width: 4, beginArrowType: 'triangle' } });
  S(s, 'line', { x: 6.578, y: 6.98, w: 2.277, h: 0, line: { color: C.dark, width: 4, beginArrowType: 'triangle' } });
  /* chevron train */
  [[3.714, C.orange], [5.151, C.dark], [6.589, C.orange], [8.026, C.dark]].forEach(([x, col]) =>
    S(s, 'chevron', { x, y: 5.332, w: 1.844, h: 0.806, fill: { color: col }, line: { type: 'none' } }));
  S(s, 'chevron', { x: 9.651, y: 5.332, w: 0.603, h: 0.806, fill: { color: 'CACACA' }, line: { type: 'none' } });
  glyph(s, 4.339, 5.545, 0.382, 0.382, C.white);
  glyph(s, 5.842, 5.539, 0.459, 0.459, C.white);
  glyph(s, 7.281, 5.509, 0.459, 0.459, C.white);
  glyph(s, 8.72, 5.545, 0.507, 0.382, C.white);
  /* numbered rings with dashed feeds to the trophy card */
  [['01', 4.167, C.dark], ['02', 5.612, C.orange], ['03', 7.062, C.dark], ['04', 8.479, C.orange]].forEach(([n, x, col]) => {
    S(s, 'ellipse', { x, y: 4.172, w: 0.64, h: 0.639, fill: { color: col }, line: { type: 'none' } });
    S(s, 'ellipse', { x: x + 0.065, y: 4.237, w: 0.51, h: 0.51, fill: { color: C.white }, line: { type: 'none' } });
    text(s, [x, 4.28, 0.64, 0.4], [n], { fontFace: F.md, fontSize: 16, align: 'center' });
    S(s, 'line', { x: x + 0.32, y: 4.15, w: 0, h: -0.28, line: { color: col, width: 1.5, endArrowType: 'triangle' } });
  });
  S(s, 'roundRect', { x: 5.78, y: 2.785, w: 1.739, h: 0.797, rectRadius: 0.1, fill: { color: C.ink }, line: { type: 'none' } });
  glyph(s, 6.484, 2.99, 0.331, 0.373, C.white);
  [[4.531, 1.187], [7.59, 1.193]].forEach(([x, w]) => S(s, 'line',
    { x, y: 3.853, w, h: -0.729, line: { color: C.black, width: 1, dashType: 'dash', endArrowType: 'triangle' } }));
  caption(s, 0.833, 1.889, 2.579, null, null, { bodyColor: C.black });
  caption(s, 0.819, 3.642, 2.579, null, null, { bodyColor: C.black });
  caption(s, 10.02, 1.838, 2.579, null, null, { right: true, bodyColor: C.black });
  caption(s, 9.99, 3.725, 2.579, null, null, { right: true, bodyColor: C.black });
}

/* ---------------------------------------------------------------- slide 32 */
function slide32(s) {
  gradientBar(s);
  chrome(s, { num: '32' });
  text(s, [5.275, 1.447, 3.193, 0.651], ['Infographic'],
    { fontFace: F.md, fontSize: 36, color: C.orange, align: 'center', lineSpacingMultiple: 0.9 });
  const cx = 6.777, cy = 4.315, R = 1.74;
  [[182, 268, C.orange], [272, 358, C.dark], [2, 88, C.orange], [92, 178, C.dark]].forEach(([a, b, col]) =>
    S(s, 'blockArc', {
      x: cx - R, y: cy - R, w: R * 2, h: R * 2, angleRange: [a, b], arcThicknessRatio: 0.42,
      fill: { color: col }, line: { type: 'none' },
    }));
  [[5.118, 2.667, C.dark], [7.615, 2.734, C.orange], [5.112, 5.045, C.orange], [7.558, 5.045, C.dark]].forEach(([x, y, col]) =>
    S(s, 'ellipse', { x, y, w: 0.91, h: 0.91, fill: { color: col }, line: { type: 'none' } }));
  glyph(s, 5.343, 2.937, 0.403, 0.403, C.white);
  glyph(s, 7.912, 3.007, 0.336, 0.336, C.white);
  glyph(s, 5.316, 5.355, 0.445, 0.253, C.white);
  glyph(s, 7.807, 5.321, 0.403, 0.403, C.white);
  caption(s, 1.496, 2.727, 2.579, null, null, { bodyColor: C.black });
  caption(s, 1.496, 4.66, 2.579, null, null, { bodyColor: C.black });
  caption(s, 9.269, 2.727, 2.579, null, null, { right: true, bodyColor: C.black });
  caption(s, 9.212, 4.662, 2.579, null, null, { right: true, bodyColor: C.black });
}

/* ---------------------------------------------------------------- slide 33 */
function slide33(s) {
  gradientBar(s);
  chrome(s, { num: '33' });
  text(s, [0.646, 1.924, 1.98, 0.303], ['BUSINESS MODEL'],
    { fontFace: F.md, fontSize: 12, charSpacing: 2, align: 'center' });
  text(s, [0.582, 2.464, 3.193, 0.651], ['Infographic'],
    { fontFace: F.md, fontSize: 36, color: C.orange, align: 'center', lineSpacingMultiple: 0.9 });
  /* the black S-curve threading the four letter discs */
  [[4.106, 180], [5.813, 0], [7.515, 180], [9.221, 0]].forEach(([x, from]) =>
    S(s, 'blockArc', {
      x: x - 0.905, y: 4.381 - 0.905, w: 1.81, h: 1.81,
      angleRange: [from, from + 180], arcThicknessRatio: 0.1,
      fill: { color: C.ink }, line: { type: 'none' },
    }));
  [['S', 3.599, 3.866], ['W', 5.31, 3.865], ['O', 7.016, 3.866], ['T', 8.691, 3.947]].forEach(([ch, x, y]) => {
    S(s, 'ellipse', { x, y, w: 1.03, h: 1.03, fill: { color: C.orange }, line: { color: C.white, width: 2 } });
    text(s, [x, y + 0.16, 1.03, 0.707], [ch], { fontFace: F.md, fontSize: 36, color: C.white, align: 'center' });
  });
  caption(s, 3.838, 5.596, 2.632, null, null, { bodyColor: C.black });
  caption(s, 5.645, 1.841, 2.632, null, null, { bodyColor: C.black });
  caption(s, 7.278, 5.586, 2.632, null, null, { bodyColor: C.black });
  caption(s, 9.121, 1.834, 2.632, null, null, { bodyColor: C.black });
}

/* ---------------------------------------------------------------- slide 34 */
function slide34(s) {
  rect(s, 0.08, 0, 3.28, 5.2, C.orange);
  rect(s, 3.074, 6.482, 3.257, 1.033, C.dark);
  rect(s, 6.331, 6.482, 1.031, 1.033, C.orange);
  gradientBar(s);
  photo(s, 0.867, 1.018, 5.464, 5.464);
  chrome(s, { num: '34', top: C.white, burger: C.black });
  text(s, [8.562, 1.27, 4.057, 0.774], ['Riyan Rugiani'], head);
  rect(s, 7.398, 1.656, 1.006, 0.05, C.plum);
  text(s, [7.301, 2.12, 0.419, 0.774], ['\u201C'], { fontFace: F.sb, fontSize: 40, color: C.g40, align: 'center' });
  text(s, [7.301, 2.485, 4.674, 0.688], [L.g, L.h, L.i], body(C.g59));
  text(s, [7.284, 3.666, 2.188, 0.337], ['Text Tittle  Here'], { fontFace: F.sb, fontSize: 14, color: C.ink });
  text(s, [7.284, 4.087, 1.983, 0.352],
    [[['California, ', { fontFace: F.sb }], ['03/14/1990', { fontFace: F.os }]]],
    { fontSize: 10.5, italic: true, color: C.g59, fontFace: F.sb, lineSpacingMultiple: 1.5 });
  text(s, [7.284, 4.317, 1.746, 0.307], ['CEO & Owner'], Object.assign(body(C.ink), { fontSize: 9 }));
  [7.273, 9.383].forEach(x => text(s, [x, 5.199, 2.188, 0.286], [L.li],
    { fontFace: F.md, fontSize: 11, color: C.ink, underline: { style: 'sng' } }));
  text(s, [3.78, 6.914, 2.058, 0.269], [L.tt], { fontFace: F.os, fontSize: 10, charSpacing: 3, color: C.white });
  arrow(s, 6.581, 7.049, 0.385, 1);
  [7.468, 9.374, 11.355].forEach(x => text(s, [x, 6.818, 1.763, 0.486], [L.duist, L.aure], body(C.g59)));
  text(s, [12.212, 4.254, 0.738, 0.816], ['01/ 02'],
    { fontFace: F.ws, fontSize: 12, align: 'center', rotate: 90, lineSpacingMultiple: 2.5 });
  S(s, 'line', { x: 12.6185, y: 3.3055, w: 0, h: 0.887, line: { color: C.white, width: 0.5 } });
}

/* ---------------------------------------------------------------- slide 35 */
function slide35(s) {
  gradientBar(s);
  chrome(s, { num: '35' });
  photo(s, 7.815, 1.807, 5.519, 4.056);
  text(s, [1.388, 1.68, 3.338, 0.774], ['Contact Us'], Object.assign({ align: 'center' }, head));
  text(s, [1.453, 2.591, 4.674, 0.688], [L.g, L.h, L.i], body(C.g80));
  const label = (x, y, w, t, right) => text(s, [x, y, w, 0.286], [t],
    { fontFace: F.sb, fontSize: 11, charSpacing: 1.1, color: C.dark, align: right ? 'right' : 'left' });
  label(1.435, 3.821, 1.8, 'OFFICE HOURS');
  label(3.777, 3.811, 1.7, 'GET IN TOUCH', true);
  label(1.467, 5.002, 1.5, 'FOLLOW US');
  label(3.777, 4.977, 1.7, 'OUR ADDRESS', true);
  const line2 = (x, y, w, a, b) => text(s, [x, y, w, 0.486], [a, b], Object.assign(body(C.g80), { wrap: false }));
  line2(1.446, 4.114, 1.8, 'Monday - Saturday', '09.00 \u2013 18.00');
  line2(4.072, 4.114, 1.8, '(+62) 085 6554 3435', '(+62) 089 6398 3432');
  line2(1.435, 5.305, 1.8, 'www.fina.com', 'example@gmail.com');
  line2(4.111, 5.305, 1.8, '123, Rev Avenue, ', 'Kolabagan LA , 267');
  rect(s, 7.823, 5.791, 2.382, 0.716, C.orange);
  text(s, [8.105, 6.031, 1.954, 0.236], [L.tt],
    { fontFace: F.md, fontSize: 8, charSpacing: 3, color: C.white, align: 'center' });
}

/* ------------------------------------------------- slides 36-38: icon sets */
const ICON_SHEETS = [
  {
    size: 0.38,
    cols: [1.42, 2.1, 2.76, 3.38, 4.04, 4.68, 5.35, 6.0, 6.66, 7.3, 7.96, 8.61, 9.26, 9.92, 10.57, 11.22, 11.89],
    rows: [0.91, 1.66, 2.45, 3.22, 3.99, 4.79, 5.57, 6.35],
    grid: ['xxxx..xxxxxxxx.xx', 'xxxxxxxxxxxxxxxxx', 'xxxxxxxxxxxxxxxxx', 'xxxxx.xxxxxxxxxxx',
      'xxxx.xx.xxxx.xxx.', 'xxxxxxxxx.xxxxxxx', 'xxxxxx.xxxxxx.xxx', 'x.xxxxxxxxxxxxxxx'],
  },
  {
    size: 0.32,
    cols: [1.6, 2.24, 2.87, 3.52, 4.17, 4.81, 5.45, 6.09, 6.75, 7.39, 8.04, 8.69, 9.34, 9.98, 10.63, 11.28, 11.92],
    rows: [1.01, 1.77, 2.54, 3.3, 4.06, 4.85, 5.64, 6.42],
    grid: ['xxxxxxxxxxxxxxxxx', 'xxxxxxxxxxxxxxxxx', 'xxxxxxxxxxxxx.xxx', 'xxxxxx..xxxx..xxx',
      'xxxxxxxxxxxxxxxxx', 'xxxxxxxxxxxxxxxxx', 'xxxxxxxxxxxxxxxxx', 'xxxxxxxxxxxxxxxxx'],
  },
  {
    size: 0.3,
    cols: [1.43, 2.11, 2.78, 3.42, 4.08, 4.74, 5.4, 6.06, 6.73, 7.39, 8.07, 8.72, 9.39, 10.05, 10.73, 11.39, 12.05],
    rows: [0.98, 1.78, 2.54, 3.33, 4.17, 4.96, 5.74, 6.5],
    grid: ['xxxxxxxxxxxxxxxxx', 'xxxxxxxxxxxxxxxxx', 'xxxxxxxxxxxxxxxxx', 'xxxxxxxxxxxxxxxxx',
      'xxxxxxxxxxxxxxxxx', 'xxxxxxxxxxxxxxxxx', 'xxxxxxxxxxxxxxxxx', 'xxx............xx'],
  },
];

/* the solid silhouettes cycled through the three pictogram library pages */
const ICON_SHAPES = ['roundRect', 'ellipse', 'rect', 'triangle', 'hexagon', 'diamond', 'octagon', 'regularPentagon'];

/** icon library page: a grid of pictogram stand-ins */
function iconSheet(s, sheet) {
  burger(s, C.black);
  sheet.grid.forEach((row, r) => row.split('').forEach((cell, c) => {
    if (cell !== 'x') return;
    S(s, ICON_SHAPES[(r * 5 + c * 3) % ICON_SHAPES.length], {
      x: sheet.cols[c] - sheet.size / 2, y: sheet.rows[r] - sheet.size / 2, w: sheet.size, h: sheet.size,
      rectRadius: 0.06, fill: { color: C.ink }, line: { type: 'none' },
    });
  }));
}

/* ------------------------------------------------------------------ render */
const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
  slide31, slide32, slide33, slide34, slide35,
  s => iconSheet(s, ICON_SHEETS[0]), s => iconSheet(s, ICON_SHEETS[1]), s => iconSheet(s, ICON_SHEETS[2]),
  slide39];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'FINA', width: 13.333, height: 7.5 });
pptx.layout = 'FINA';
pptx.author = 'FINA';
pptx.title = 'Fina — Business Presentation Template';

BUILDERS.forEach(build => build(pptx.addSlide()));

pptx.writeFile({ fileName: path.join(__dirname, '15daa284-964a-4125-a4d6-9d7923577fdf_grok_final.pptx') })
  .then(f => console.log('wrote', f));
