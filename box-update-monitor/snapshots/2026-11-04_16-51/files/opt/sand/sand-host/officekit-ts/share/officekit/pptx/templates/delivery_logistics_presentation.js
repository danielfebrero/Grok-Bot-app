/**
 * "Delivery - Logistic Presentation Template" rebuilt with pptxgenjs.
 * 32 slides, 13.333 x 7.5 in.
 *
 * The source deck's photo frames are unfilled picture placeholders, so they
 * render blank; the only real bitmaps are the device mock-ups on slides 17-19,
 * which are redrawn here from native shapes plus a pale "picture" swatch.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const C = {
  navy: '2D3859',       // accent1
  navyDark: '222A43',   // accent1, lumMod 75%
  orange: 'F26D3D',     // accent2
  sand: 'D9A78B',       // accent3
  red: 'F2441D',        // accent4
  maroon: '731007',     // accent5
  taupe: '937863',      // accent6
  ink: '262626',
  body: '404040',
  slate: '303C4E',
  white: 'FFFFFF',
  paper: 'F2F2F2',
  silver: 'D9D9D9',
  grey: 'A6A6A6',
  picture: 'DCE6F1',    // swatch standing in for a photo
};

const F = { head: 'Hind SemiBold', body: 'Open Sans', heebo: 'Heebo Medium', heeboBlack: 'Heebo Black', inter: 'Inter', sym: 'DejaVu Sans' };

// glyphs standing in for the deck's vector icon set
const G = {
  truck: '\u2794', search: '\u25CE', bus: '\u25A4', box: '\u25A3', cart: '\u229E',
  phone: '\u260E', mail: '\u2709', plane: '\u2708', flag: '\u2691', star: '\u2605',
  play: '\u25B6', target: '\u25C9', bars: '\u2630', home: '\u2302', pie: '\u25D4',
  smile: '\u263A', map: '\u25A6', train: '\u25EB', check: '\u2714', plus: '\u2295',
  diamond: '\u2756',
};

const NBSP = '\u00a0';

/* ---------------------------------------------------------------- helpers */

const noLine = { type: 'none' };

/** Fresh shadow spec each call (pptxgenjs mutates the object it is handed). */
function shadow(blur, opacity) {
  return { shadow: { type: 'outer', blur: blur || 22, offset: 2, angle: 90, color: '000000', opacity: opacity || 0.11 } };
}

/** Solid-filled shape with no outline; `extra.transparency` fades the fill. */
function shape(s, kind, x, y, w, h, color, extra) {
  const o = Object.assign({}, extra || {});
  const fill = { color };
  if (o.transparency !== undefined) { fill.transparency = o.transparency; delete o.transparency; }
  return s.addShape(kind, Object.assign({ x, y, w, h, fill, line: noLine }, o));
}

/** Rounded rectangle; `r` is the corner radius in inches. */
function card(s, x, y, w, h, color, r, extra) {
  return shape(s, 'roundRect', x, y, w, h, color, Object.assign({ rectRadius: r }, extra || {}));
}

/** Blend two hex colours; `t` runs 0 (from) to 1 (to). */
function mix(from, to, t) {
  let out = '';
  for (let i = 0; i < 6; i += 2) {
    const a = parseInt(from.substr(i, 2), 16), b = parseInt(to.substr(i, 2), 16);
    out += ('0' + Math.round(a + (b - a) * t).toString(16)).slice(-2);
  }
  return out.toUpperCase();
}

/** Linear gradient painted as opaque stripes running left-to-right or top-to-bottom. */
function ramp(s, x, y, w, h, from, to, steps, dir) {
  for (let i = 0; i < steps; i++) {
    const f = i / (steps - 1), c = mix(from, to, f);
    if (dir === 'y') shape(s, 'rect', x, y + h * i / steps, w, h / steps + 0.01, c);
    else shape(s, 'rect', x + w * i / steps, y, w / steps + 0.01, h, c);
  }
}

/** 45-degree gradient: stripes rotated to run perpendicular to the light direction. */
function diagonalRamp(s, cx, cy, span, len, from, to, steps) {
  const u = Math.SQRT1_2, t = span / steps + 0.14;
  for (let i = 0; i < steps; i++) {
    const f = (i + 0.5) / steps, d = (f - 0.5) * span;
    shape(s, 'rect', cx + d * u - len / 2, cy + d * u - t / 2, len, t, mix(from, to, f), { rotate: 315 });
  }
}

/** Text block. `runs` is a string or an array of pptxgenjs text objects. */
function text(s, runs, o) {
  return s.addText(runs, Object.assign(
    { fontFace: F.body, fontSize: 18, color: C.body, valign: 'top', margin: [0.1, 0.1, 0.05, 0.05], wrap: true }, o));
}

/** Two/three-tone heading, e.g. "Gallery |Image| Slide". */
function heading(s, parts, o) {
  const base = Object.assign({ fontFace: F.head, fontSize: 44, color: C.ink, valign: 'middle' }, o || {});
  return text(s, parts.map(function (p) {
    return { text: p[0], options: { color: p[1] || base.color, bold: base.bold, fontFace: base.fontFace, fontSize: base.fontSize } };
  }), base);
}

/** Body copy used all over the deck (12 pt grey, 130% leading). */
function body12(s, str, x, y, w, h, o) {
  return text(s, str, Object.assign({ x, y, w, h, fontSize: 12, color: C.body, lineSpacingMultiple: 1.3 }, o || {}));
}

/** Bold "Lorem Ipsum" lead-in followed by regular copy. */
function leadIn(s, tail, x, y, w, h, o) {
  return text(s, [
    { text: 'Lorem Ipsum' + NBSP, options: { bold: true, fontSize: 14, color: C.body } },
    { text: tail, options: { fontSize: 12, color: C.body } },
  ], Object.assign({ x, y, w, h, fontSize: 12, lineSpacingMultiple: 1.3 }, o || {}));
}

/** Pale swatch standing in for a photo thumbnail inside the device mock-ups. */
function pictureSwatch(s, x, y, w, h) {
  s.addShape('rect', { x, y, w, h, fill: { color: C.picture }, line: { color: C.white, width: 1 } });
}

/** Circular progress ring: light track + coloured arc + knob. */
function ring(s, cx, cy, d, color, sweepEnd, label, o) {
  const opt = o || {};
  const x = cx - d / 2, y = cy - d / 2;
  s.addShape('arc', { x, y, w: d, h: d, line: { color: 'D6D8DC', width: 3 }, angleRange: [0, 359.9] });
  s.addShape('arc', { x, y, w: d, h: d, line: { color, width: 4 }, angleRange: [270, sweepEnd] });
  const a = (sweepEnd - 90) * Math.PI / 180, kd = 0.15;
  shape(s, 'ellipse', cx + Math.cos(a) * d / 2 - kd / 2, cy + Math.sin(a) * d / 2 - kd / 2, kd, kd, color);
  shape(s, 'ellipse', cx + Math.cos(a) * d / 2 - 0.028, cy + Math.sin(a) * d / 2 - 0.028, 0.056, 0.056, C.white);
  if (label) {
    text(s, label, { x: cx - 0.4, y: cy - 0.19, w: 0.8, h: 0.38, align: 'center', valign: 'middle', fontSize: 16, color: C.body, bold: !!opt.bold });
  }
}

/** Small round check-mark bullet. */
function checkDot(s, x, y, d, ringColor, tickColor) {
  shape(s, 'ellipse', x, y, d, d, ringColor);
  text(s, G.check, { x, y, w: d, h: d, align: 'center', valign: 'middle', fontSize: Math.round(d * 42), color: tickColor, fontFace: F.sym });
}

/** Square icon tile with a glyph, used for feature lists. */
function iconTile(s, x, y, w, h, fill, glyph, glyphColor, o) {
  const opt = o || {};
  card(s, x, y, w, h, fill, opt.rectRadius === undefined ? Math.min(w, h) * 0.22 : opt.rectRadius);
  text(s, glyph, { x, y, w, h, align: 'center', valign: 'middle', fontSize: opt.fontSize || Math.round(h * 26), color: glyphColor || C.white, fontFace: F.sym });
}

/** Round icon badge with a glyph. */
function iconDot(s, x, y, d, fill, glyph, glyphColor, fontSize) {
  shape(s, 'ellipse', x, y, d, d, fill);
  text(s, glyph, { x, y, w: d, h: d, align: 'center', valign: 'middle', fontSize: fontSize || Math.round(d * 30), color: glyphColor || C.white, fontFace: F.sym });
}

/* ============================================================== 1. cover */

function slide01(s) {
  // wash from full navy at the top-left down to the shaded navy bottom-right
  diagonalRamp(s, 6.667, 3.75, 15.0, 22, C.navy, C.navyDark, 60);
  text(s, 'Delivery', { x: 1.45, y: 2.3, w: 10.25, h: 2.45, fontFace: F.head, fontSize: 138, bold: true, color: C.orange, valign: 'middle' });
  text(s, 'Logistic Presentation Template', { x: 1.545, y: 4.36, w: 8.71, h: 0.45, fontSize: 20, color: C.white });
}

/* ================================================ 2. Best Delivery Logistic */

function slide02(s) {
  heading(s, [['Best ', C.slate], ['Delivery', C.orange], [' Logistic', C.slate]],
    { x: 0.87, y: 2.79, w: 6.22, h: 2.31, fontSize: 72, bold: true, lineSpacing: 65, valign: 'top' });
  s.addShape('line', { x: 7.6, y: 2.91, w: 0, h: 1.68, line: { color: C.navy, width: 2 } });
  text(s, 'Delivery Logistic Solution', { x: 8.11, y: 3.2, w: 3.3, h: 0.4, fontFace: F.head, fontSize: 18, bold: true, color: C.navy });
  body12(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut', 8.11, 3.69, 4.1, 0.6);
}

/* ================================== 3. Simple Logistic Best Logistic (banner) */

function slide03(s) {
  heading(s, [['Simple ', C.ink], ['Logistic Best ', C.orange], ['Logistic', C.ink]],
    { x: 2.24, y: 1.03, w: 8.78, h: 0.91, fontSize: 48, align: 'center' });
  body12(s, "Lorem Ipsum" + NBSP + "is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever",
    2.58, 2.19, 8.21, 0.69, { fontSize: 14, align: 'center' });

  card(s, 0.93, 3.33, 11.6, 1.89, C.navy, 0.194);
  text(s, '\u201C', { x: 1.37, y: 3.44, w: 0.53, h: 1.11, fontSize: 60, bold: true, color: C.white });
  body12(s, "Lorem Ipsum" + NBSP + "is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever",
    2.03, 3.78, 4.8, 0.98, { fontSize: 14, color: C.white });

  // two KPI columns, each preceded by a coloured divider
  [[7.03, C.orange, '365k', 7.4], [9.18, C.sand, '275k', 9.42]].forEach(function (col) {
    s.addShape('line', { x: col[0], y: 3.68, w: 0, h: 1.13, line: { color: col[1], width: 1.75 } });
    text(s, col[2], { x: col[3], y: 3.75, w: 1.7, h: 0.64, fontSize: 32, color: C.white });
    body12(s, 'Text here', col[3], 4.31, 1.49, 0.33, { color: C.white });
  });

  // stacked chevrons at the right end of the banner
  [[3.51, C.sand], [3.96, C.orange], [4.42, C.red]].forEach(function (ch) {
    shape(s, 'chevron', 11.02, ch[0] - 0.11, 0.86, 0.86, ch[1], { rotate: 270 });
  });

  body12(s, "Lorem Ipsum" + NBSP + "is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever",
    2.56, 5.67, 8.21, 0.6, { align: 'center', italic: true });
}

/* ================================= 4. Delivery Provide Logistics Solution */

function slide04(s) {
  heading(s, [['Delivery ', C.ink], ['Provide Logistics ', C.orange], ['Solution', C.ink]],
    { x: 1.35, y: 1.1, w: 8.6, h: 1.58, valign: 'top' });
  text(s, 'Lorem Ipsum' + NBSP + 'is simply dummy text of the printing and typesetting industry. ',
    { x: 1.35, y: 2.71, w: 5.57, h: 0.98, fontFace: F.head, fontSize: 20, bold: true, color: C.navy, lineSpacingMultiple: 1.3 });
  s.addShape('line', { x: 1.46, y: 4.2, w: 4.59, h: 0, line: { color: C.navy, width: 4.5 } });
  body12(s, "Lorem Ipsum" + NBSP + "is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever",
    1.39, 4.67, 5.1, 0.86);

  // two white cards on the right, each with a ringed icon
  [[3.62, C.navy, G.bus], [5.12, C.red, G.cart]].forEach(function (row, i) {
    card(s, 7.12 + i * 0.05, row[0], 5.57, 1.34, C.white, 0.216, shadow(20, 0.12));
    s.addShape('ellipse', { x: 7.5, y: row[0] + 0.32, w: 0.673, h: 0.673, fill: { color: C.white }, line: { color: row[1], width: 2 } });
    text(s, row[2], { x: 7.5, y: row[0] + 0.32, w: 0.673, h: 0.673, align: 'center', valign: 'middle', fontSize: 17, color: row[1], fontFace: F.sym });
    body12(s, 'Lorem Ipsum' + NBSP + 'is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been ',
      8.29, row[0] + 0.36, 4.16, 0.6);
  });
}

/* ======================================== 5. What We Provide For You (cards) */

function slide05(s) {
  heading(s, [['What ', C.ink], ['We Provide For ', C.orange], ['You', C.ink]],
    { x: 2.48, y: 1.07, w: 8.37, h: 0.91, fontSize: 48, bold: true, align: 'center' });
  body12(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ull',
    2.31, 2.1, 8.74, 0.61, { align: 'center' });

  const cards = [
    { x: 1.35, fill: C.navy, tile: '697CB3', glyph: G.truck },
    { x: 5.26, fill: C.orange, tile: 'F7A78B', glyph: G.search },
    { x: 9.14, fill: C.sand, tile: 'E8CAB9', glyph: G.bus },
  ];
  cards.forEach(function (c) {
    card(s, c.x, 3.39, 2.84, 3.14, c.fill, 0.159);
    iconTile(s, c.x + 1.11, 3.74, 0.625, 0.589, c.tile, c.glyph, C.white, { fontSize: 16, rectRadius: 0.069 });
    text(s, 'Delivery Logistic Featured #01',
      { x: c.x + 0.26, y: 4.5, w: 2.33, h: 0.71, align: 'center', fontFace: F.head, fontSize: 18, bold: true, color: C.white });
    body12(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor',
      c.x + 0.2, 5.24, 2.45, 0.87, { align: 'center', color: C.white });
  });
}

/* ==================================== 6. Proud To Deliver Excellence (photo) */

function slide06(s) {
  heading(s, [['Proud To Deliver Excellence', C.navy]],
    { x: 6.75, y: 1.96, w: 6.23, h: 1.57, fontSize: 48, bold: true, lineSpacing: 48, valign: 'top' });
  body12(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin consectetur egestas lacus ac porta. Quisque ultricies nibh eu tortor porttitor, ',
    6.75, 3.5, 4.82, 0.87, { align: 'justify' });
  card(s, 5.77, 4.72, 5.47, 0.87, C.navy, 0.102);
  text(s, G.truck, { x: 5.94, y: 4.91, w: 0.49, h: 0.49, align: 'center', valign: 'middle', fontSize: 16, color: C.white, fontFace: F.sym });
  body12(s, '\u201C Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin consectetur egestas lacus ac porta. \u201C ',
    6.66, 4.86, 4.08, 0.51, { color: C.white, italic: true, align: 'justify', lineSpacingMultiple: 1.15 });
}

/* ================================== 7. We Deliver Your Package To Your Door */

function slide07(s) {
  // photo band: near-black top-left grading to pale grey bottom-right
  diagonalRamp(s, 6.667, 2.7, 14.0, 22, '242C38', 'DCDEE3', 64);
  shape(s, 'rect', 0, 5.397, 13.334, 2.11, C.white);
  text(s, 'We Deliver Your Package To Your Door',
    { x: 1.15, y: 1.12, w: 7.32, h: 1.45, fontFace: F.head, fontSize: 40, color: C.white, valign: 'top' });
  card(s, 1.26, 3.07, 1.582, 0.384, C.white, 0.19);
  text(s, 'Learn More', { x: 1.26, y: 3.07, w: 1.582, h: 0.384, align: 'center', valign: 'middle', fontSize: 11, color: C.ink });
  s.addShape('roundRect', { x: 3.04, y: 3.07, w: 1.868, h: 0.384, fill: { type: 'none' }, line: { color: C.silver, width: 1 }, rectRadius: 0.19 });
  text(s, 'Support Us', { x: 3.04, y: 3.07, w: 1.868, h: 0.384, align: 'center', valign: 'middle', fontSize: 11, color: C.paper });

  card(s, 1.0, 4.34, 11.333, 2.046, C.white, 0.19, shadow(24, 0.1));
  ['01', '02', '03', '04'].forEach(function (n, i) {
    const x = 1.4 + i * 2.68;
    text(s, n, { x, y: 4.54, w: 1.56, h: 0.51, fontSize: 24, bold: true, color: C.navy });
    text(s, 'Global logistics  ', { x, y: 5.08, w: 2.79, h: 0.34, fontSize: 14, bold: true, color: C.body });
    body12(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.', x, 5.42, 2.42, 0.61);
  });
}

/* ============================ 8. Being A Global Logistics (doughnut chart) */

function slide08(s) {
  heading(s, [['Being A ', C.ink], ['Global Logistics ', C.orange], ['Service Provider', C.ink]],
    { x: 6.36, y: 1.34, w: 6.92, h: 1.33, fontSize: 40, lineSpacing: 38, valign: 'top' });
  text(s, 'Lorem Ipsum' + NBSP + 'is simply dummy text',
    { x: 6.4, y: 2.58, w: 5.58, h: 0.38, fontSize: 14, bold: true, color: C.navy, lineSpacingMultiple: 1.3 });
  body12(s, 'Lorem Ipsum' + NBSP + 'is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been ',
    6.4, 3.22, 5.6, 0.6);

  card(s, 5.44, 4.26, 6.64, 2.25, C.navy, 0.238);
  s.addChart('doughnut', [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: [80, 70] }], {
    x: 5.6, y: 4.5, w: 2.238, h: 1.77,
    chartColors: [C.red, '454E68'], holeSize: 80, showLegend: false, showTitle: false,
    chartArea: { fill: { color: C.navy } }, plotArea: { fill: { color: C.navy } }, dataBorder: { pt: 0, color: C.navy },
  });
  text(s, '50%', { x: 6.09, y: 5.14, w: 1.26, h: 0.44, align: 'center', fontSize: 18, bold: true, color: C.white });
  text(s, 'Fast & Good Response', { x: 8.01, y: 4.73, w: 3.46, h: 0.45, fontSize: 18, color: C.white, valign: 'middle' });
  s.addShape('line', { x: 8.11, y: 5.24, w: 2.654, h: 0, line: { color: C.red, width: 1.75 } });
  body12(s, 'Lorem Ipsum' + NBSP + 'is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been ',
    8.01, 5.3, 3.46, 0.86, { color: C.white });
}

/* ============================================== 9. Gallery Image Slide */

function slide09(s) {
  heading(s, [['Gallery ', C.ink], ['Image', C.orange], [' Slide', C.ink]], { x: 0.58, y: 1.09, w: 7.49, h: 0.91, fontSize: 48 });
  text(s, [
    { text: 'Lorem ipsum dolor sit amet, ', options: { bold: true, color: C.body, fontSize: 12 } },
    { text: 'consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua', options: { color: C.body, fontSize: 12 } },
  ], { x: 0.62, y: 2.12, w: 6.0, h: 0.6, fontSize: 12, lineSpacingMultiple: 1.3 });
}

/* =============================================== 10. Gallery Great Event */

function slide10(s) {
  heading(s, [['Gallery ', C.ink], ['Great ', C.orange], ['Event', C.ink]],
    { x: 3.69, y: 1.42, w: 5.96, h: 0.91, fontSize: 48, bold: true });
  body12(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin consectetur egestas lacus ac porta. Quisque',
    3.91, 2.56, 5.52, 0.69, { fontSize: 14, align: 'center' });
}

/* ========================================= 11. Alex Fernandes (CV layout) */

function slide11(s) {
  text(s, 'Alex Fernandes', { x: 1.28, y: 1.13, w: 4.99, h: 0.91, fontFace: F.head, fontSize: 48, color: C.ink, valign: 'middle' });
  text(s, 'Skill & Performance', { x: 1.28, y: 2.34, w: 4.03, h: 0.51, fontSize: 24, color: C.body, valign: 'middle' });

  [['2022', C.navy, 3.24], ['2023', C.orange, 4.59]].forEach(function (row) {
    text(s, row[0], { x: 1.34, y: row[2], w: 1.38, h: 0.56, fontSize: 20, color: row[1], valign: 'middle', lineSpacingMultiple: 1.5 });
    text(s, 'Work Experience', { x: 2.29, y: row[2] + 0.06, w: 2.48, h: 0.42, fontSize: 14, color: row[1], valign: 'middle', lineSpacingMultiple: 1.5 });
    body12(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.', 2.32, row[2] + 0.54, 3.87, 0.61);
  });

  card(s, 6.98, 5.1, 5.397, 1.879, C.white, 0.177, shadow(24, 0.12));
  ring(s, 7.92, 6.04, 1.13, C.navy, 234, '65%');
  text(s, 'Project Success', { x: 8.91, y: 5.43, w: 2.48, h: 0.4, fontSize: 18, bold: true, color: C.body });
  body12(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. ', 8.91, 5.95, 3.04, 0.71, { fontSize: 14, align: 'justify' });
}

/* ================================================ 12. Mark Martinez profile */

function slide12(s) {
  shape(s, 'rect', 0, 0, 6.667, 7.5, C.paper, { line: noLine, transparency: 40 });
  shape(s, 'round2SameRect', 1.093, 5.739, 5.441, 1.141, C.navy, { rotate: 180, rectRadius: 0.13 });
  text(s, 'Mark Martinez', { x: 1.46, y: 5.9, w: 4.41, h: 0.5, align: 'center', fontSize: 20, color: C.white });
  body12(s, 'Job Tittle Here', 2.32, 6.28, 2.68, 0.43, { align: 'center', color: C.white, lineSpacingMultiple: 1.5 });

  card(s, 6.964, 0.931, 5.208, 5.499, C.white, 0.403, shadow(30, 0.1));
  text(s, 'Profile', { x: 7.49, y: 1.02, w: 2.74, h: 1.49, fontSize: 60, color: C.body, valign: 'middle' });
  s.addShape('line', { x: 7.65, y: 2.61, w: 2.478, h: 0, line: { color: C.paper, width: 1 } });
  body12(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Lorem ipsum dolor sit', 7.54, 2.6, 4.32, 0.6);

  [['Instagram', C.navy, 3.98, G.target], ['Twitter', C.orange, 4.69, G.mail], ['Youtube', C.sand, 5.36, G.play]]
    .forEach(function (row) {
      iconDot(s, 7.59, row[2] + 0.08, 0.295, row[1], row[3], C.white, 9);
      text(s, row[0], { x: 8.03, y: row[2] - 0.02, w: 2.08, h: 0.34, fontSize: 10.5, color: C.body, lineSpacingMultiple: 1.5 });
      text(s, '@Mark Martinez97', { x: 8.03, y: row[2] + 0.22, w: 2.08, h: 0.38, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5 });
    });
}

/* ==================================================== 13. Team Profile */

function slide13(s) {
  heading(s, [['Team ', C.ink], ['Profile', C.orange]], { x: 2.52, y: 0.88, w: 8.3, h: 0.84, align: 'center' });

  const cols = [
    { x: 1.052, y: 4.184, name: 'Mark Martinez', color: C.navy },
    { x: 4.888, y: 4.225, name: 'Jhosua ', color: C.orange },
    { x: 8.72, y: 4.225, name: 'Robert khan', color: C.sand },
  ];
  cols.forEach(function (c) {
    card(s, c.x, 2.083, 3.545, 4.667, C.white, 0.35, shadow(26, 0.12));
    shape(s, 'round2SameRect', c.x, c.y, 3.549, 0.562, c.color, { rotate: 180, rectRadius: 0.064 });
    text(s, c.name, { x: c.x + 0.3, y: c.y + 0.06, w: 2.906, h: 0.4, align: 'center', fontSize: 18, color: C.white });
    text(s, 'Work Experience', { x: c.x + 0.3, y: c.y + 0.84, w: 2.06, h: 0.37, fontSize: 16, color: C.body });
    [0, 1].forEach(function (i) {
      checkDot(s, c.x + 0.42, c.y + 1.47 + i * 0.378, 0.171, c.color, C.white);
      body12(s, 'Lorem ipsum dolor sit amet, ', c.x + 0.63, c.y + 1.4 + i * 0.378, 2.42, 0.34);
    });
  });
}

/* ==================================================== 14. Our Best Team */

function slide14(s) {
  heading(s, [['Our ', C.ink], ['Best', C.orange], [' Team', C.ink]],
    { x: 3.1, y: 1.03, w: 7.13, h: 0.84, bold: true, align: 'center' });

  const members = [
    { x: 0.931, fill: C.navy, name: 'Donovan Southern', role: 'Engineer' },
    { x: 4.882, fill: C.orange, name: 'Emilia Laura', role: 'Manager' },
    { x: 8.833, fill: C.sand, name: 'Faisal Garner', role: 'Staff' },
  ];
  members.forEach(function (m) {
    card(s, m.x, 2.476, 3.57, 4.109, m.fill, 0.24);
    text(s, m.name, { x: m.x + 0.1, y: 4.68, w: 3.361, h: 0.467, align: 'center', fontSize: 18, bold: true, color: C.white });
    text(s, m.role, { x: m.x + 0.1, y: 5.13, w: 3.361, h: 0.422, align: 'center', fontSize: 16, color: C.white });
    [G.mail, G.flag, G.star].forEach(function (g, i) {
      iconDot(s, m.x + 1.036 + i * 0.55, 5.885, 0.393, i === 0 ? C.white : C.paper, g, m.fill, 11);
    });
  });
}

/* ================================ 15. Explore The Logistic Service (3 cards) */

function slide15(s) {
  heading(s, [['Explore The Logistic Service', C.ink]], { x: 1.29, y: 1.18, w: 8.04, h: 0.79, lineSpacing: 42, valign: 'top' });
  body12(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin consectetur Lorem ipsum ',
    1.36, 2.12, 7.2, 0.69, { fontSize: 14, align: 'justify' });

  const tiles = [
    { x: 1.33, color: C.navy, glyph: G.truck },
    { x: 5.01, color: C.orange, glyph: G.train },
    { x: 8.69, color: C.sand, glyph: G.search },
  ];
  const bullets = ['Lorem ipsum dolor sit amet,', 'Lorem ipsum dolor sit amet', 'Lorem ipsum dolor sit amet '];
  tiles.forEach(function (t) {
    card(s, t.x, 3.85, 3.34, 2.36, C.white, 0.122, shadow(22, 0.11));
    iconTile(s, t.x + 0.42, 3.6, 0.74, 0.6, t.color, t.glyph, C.white, { fontSize: 19, rectRadius: 0.09 });
    text(s, 'Option One', { x: t.x + 0.32, y: 4.36, w: 2.51, h: 0.37, fontSize: 16, bold: true, color: C.body });
    text(s, bullets.map(function (b) { return { text: G.diamond + '  ' + b, options: { breakLine: true } }; }),
      { x: t.x + 0.3, y: 4.9, w: 2.76, h: 0.87, fontSize: 12, color: C.body, lineSpacingMultiple: 1.3, fontFace: F.body });
  });
}

/* ============================= 16. Explore The Logistic Service (3 bands) */

function slide16(s) {
  heading(s, [['Explore ', C.ink], ['The Logistic ', C.orange], ['Service', C.ink]],
    { x: 0.9, y: 2.44, w: 4.65, h: 1.45, lineSpacing: 42, valign: 'top' });
  body12(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin consectetur Lorem ipsum ',
    0.9, 3.92, 4.1, 0.99, { fontSize: 14, align: 'justify' });

  const bands = [
    { y: 0.74, fill: C.navy, dot: '697CB3', n: '01' },
    { y: 2.77, fill: C.orange, dot: 'F7A78B', n: '02' },
    { y: 4.79, fill: C.sand, dot: 'E8CAB9', n: '03' },
  ];
  bands.forEach(function (b) {
    card(s, 6.61, b.y, 6.08, 1.99, b.fill, 0.193);
    shape(s, 'ellipse', 7.01, b.y + 0.52, 0.89, 0.89, b.dot);
    text(s, b.n, { x: 7.01, y: b.y + 0.52, w: 0.89, h: 0.89, align: 'center', valign: 'middle', fontSize: 24, bold: true, color: C.white });
    text(s, 'Your Text in Here', { x: 8.18, y: b.y + 0.38, w: 2.83, h: 0.4, fontFace: F.head, fontSize: 18, color: C.white });
    body12(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin consectetur',
      8.19, b.y + 0.78, 3.9, 0.69, { fontSize: 14, color: C.white, align: 'justify' });
  });
}

/* ==================================================== 17. Laptop Mockup */

function slide17(s) {
  // laptop drawn from plain shapes: lid, screen, base, notch
  shape(s, 'roundRect', 7.14, 2.18, 5.55, 3.63, C.silver, { rectRadius: 0.08 });
  shape(s, 'rect', 7.27, 2.4, 5.3, 3.22, C.paper);
  shape(s, 'ellipse', 9.9, 2.26, 0.08, 0.07, C.white);
  shape(s, 'roundRect', 6.67, 5.81, 6.47, 0.21, 'E0E0E0', { rectRadius: 0.06 });
  shape(s, 'roundRect', 9.41, 5.83, 0.97, 0.06, C.white, { rectRadius: 0.03 });
  text(s, 'Image Placeholder', { x: 7.27, y: 3.28, w: 5.3, h: 0.4, align: 'center', fontSize: 16, color: C.grey });
  pictureSwatch(s, 9.47, 3.75, 0.85, 0.68);

  heading(s, [['Laptop ', C.ink], ['Mockup', C.orange]], { x: 1.23, y: 2.09, w: 5.1, h: 0.91, fontSize: 48, bold: true });
  text(s, 'Your Text In Here', { x: 1.3, y: 3.19, w: 2.48, h: 0.4, fontFace: F.head, fontSize: 18, bold: true, color: C.navy });
  body12(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin consectetur Lorem ipsum ',
    1.3, 3.65, 4.2, 0.99, { fontSize: 14, align: 'justify' });
  card(s, 4.74, 5.12, 5.1, 1.13, C.navy, 0.093);
  body12(s, '\u201C Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin consectetur egestas lacus ac porta. \u201C ',
    5.09, 5.3, 4.39, 0.6, { color: C.white, italic: true, align: 'justify' });
}

/* ============================== 18. Smartwatch To Control Your packet */

function slide18(s) {
  heading(s, [['Smartwatch ', C.ink], ['To Control Your ', C.orange], ['packet', C.ink]],
    { x: 1.59, y: 0.84, w: 10.24, h: 0.84 });

  // watch: strap, case, screen, crown
  shape(s, 'roundRect', 5.86, 1.72, 1.6, 5.72, '2E2E2E', { rectRadius: 0.3 });
  shape(s, 'roundRect', 5.11, 2.94, 3.11, 3.26, 'AFB3B9', { rectRadius: 0.66 });
  shape(s, 'roundRect', 5.28, 3.1, 2.77, 2.94, C.paper, { rectRadius: 0.55 });
  shape(s, 'roundRect', 8.16, 3.52, 0.14, 0.4, '8E9298', { rectRadius: 0.06 });
  text(s, 'Image Placeholder', { x: 5.28, y: 4.32, w: 2.77, h: 0.32, align: 'center', fontSize: 11, color: C.grey });
  pictureSwatch(s, 6.3, 4.64, 0.73, 0.58);

  const feats = [
    { x: 1.49, tx: 2.25, y: 3.4, color: C.navy, glyph: G.box },
    { x: 1.49, tx: 2.33, y: 4.96, color: C.orange, glyph: G.target },
    { x: 9.06, tx: 9.86, y: 3.35, color: C.sand, glyph: G.smile },
    { x: 9.06, tx: 9.9, y: 4.92, color: C.red, glyph: G.pie },
  ];
  feats.forEach(function (f) {
    iconTile(s, f.x, f.y + 0.09, 0.63, 0.63, f.color, f.glyph, C.white, { fontSize: 17, rectRadius: 0.1 });
    text(s, 'Text here', { x: f.tx, y: f.y, w: 1.8, h: 0.37, fontSize: 16, bold: true, color: f.color });
    body12(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing', f.tx, f.y + 0.34, 2.56, 0.61);
  });
}

/* ================================================= 19. Smartphone Mockup */

function slide19(s) {
  heading(s, [['Smartphone ', C.ink], ['Mockup', C.orange]], { x: 1.33, y: 1.26, w: 5.3, h: 1.72, fontSize: 48, bold: true, valign: 'top' });
  body12(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin consectetur egestas lacus ac porta. Quisque',
    1.41, 2.94, 4.0, 0.99, { fontSize: 14, align: 'justify' });

  [[1.45, C.navy, '01. OPTION', 0.237], [4.96, C.orange, '02. OPTION', 0.189]].forEach(function (o) {
    card(s, o[0], 4.29, 3.22, 1.99, o[1], o[3]);
    text(s, o[2], { x: o[0] + 0.4, y: 4.56, w: 1.8, h: 0.4, fontFace: F.head, fontSize: 18, color: C.white });
    body12(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin consectetur',
      o[0] + 0.4, 5.11, 2.48, 0.83, { color: C.white });
  });

  // two phone bodies drawn last so they overlap the cards, as in the original
  [[7.33, 0.82, 2.67, 5.3], [10.07, 1.03, 2.57, 5.42]].forEach(function (p) {
    shape(s, 'roundRect', p[0], p[1], p[2], p[3], '141414', { rectRadius: 0.22 });
    shape(s, 'roundRect', p[0] + 0.09, p[1] + 0.09, p[2] - 0.18, p[3] - 0.18, C.white, { rectRadius: 0.16 });
    text(s, 'Picture', { x: p[0], y: p[1] + 0.4, w: p[2], h: 0.4, align: 'center', fontSize: 16, color: C.ink });
    pictureSwatch(s, p[0] + p[2] / 2 - 0.46, p[1] + 1.85, 0.92, 0.74);
  });
}

/* ==================================================== 20. Process Delivery */

function slide20(s) {
  heading(s, [['Process ', C.ink], ['Delivery ', C.orange]],
    { x: 1.33, y: 0.93, w: 10.67, h: 0.79, align: 'center', lineSpacing: 42, valign: 'top' });
  leadIn(s, 'is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been Lorem Ipsum' + NBSP + 'is simply dummy text of the printing and typesetting industry. ',
    2.59, 1.86, 8.16, 0.87, { align: 'center' });

  const steps = [
    { x: 1.35, fill: C.navy, glyph: G.phone, label: 'Online Quote', tx: 0.56 },
    { x: 4.32, fill: C.orange, glyph: G.cart, label: 'Packing Process', tx: 3.52 },
    { x: 7.38, fill: C.sand, glyph: G.box, label: 'Inspection Process', tx: 6.58 },
    { x: 10.48, fill: C.red, glyph: G.truck, label: 'Delivery Process', tx: 9.69 },
  ];
  const arrowX = [3.21, 6.19, 9.30], arrowFill = [C.navy, C.red, C.sand];
  steps.forEach(function (st, i) {
    shape(s, 'ellipse', st.x - 0.14, 3.05, 1.78, 1.78, C.white);
    iconTile(s, st.x, 3.19, 1.5, 1.5, st.fill, st.glyph, C.white, { fontSize: 34, rectRadius: 0.28 });
    text(s, st.label, { x: st.tx, y: 5.05, w: 3.09, h: 0.4, align: 'center', fontSize: 18, bold: true, color: C.body });
    text(s, [
      { text: 'Lorem Ipsum' + NBSP, options: { bold: true, fontSize: 14, color: C.body } },
      { text: 'is simply dummy text of the printing and', options: { fontSize: 12, color: C.body } },
    ], { x: st.tx + 0.61, y: 5.54, w: 1.87, h: 0.9, align: 'center', lineSpacingMultiple: 1.3, fontSize: 12, fontFace: F.body });
    if (i < 3) {
      shape(s, 'chevron', arrowX[i], 3.65, 0.375, 0.59, arrowFill[i]);
      shape(s, 'chevron', arrowX[i] + 0.375, 3.65, 0.375, 0.59, arrowFill[i]);
    }
  });
}

/* ============================================ 21. Timeline Logistic Delivery */

function slide21(s) {
  heading(s, [['Timeline ', C.body], ['Logistic', C.navy], [' Delivery', C.body]],
    { x: 1.02, y: 1.25, w: 7.34, h: 0.77, fontFace: F.heebo, fontSize: 40 });
  s.addShape('line', { x: 2.89, y: 3.45, w: 9.4, h: 0, line: { color: C.grey, width: 1.5, dashType: 'dash' } });

  const marks = [
    { x: 2.3, year: '2018', color: C.navy, cardX: 1.66, cardY: 4.7, r: 0.207 },
    { x: 6.18, year: '2019', color: C.orange, cardX: 5.7, cardY: 4.69, r: 0.228 },
    { x: 10.04, year: '2020', color: C.sand, cardX: 9.73, cardY: 4.67, r: 0.211 },
  ];
  marks.forEach(function (m) {
    card(s, m.x, 2.65, 1.19, 0.47, m.color, 0.059);
    text(s, m.year, { x: m.x + 0.12, y: 2.65, w: 0.94, h: 0.47, align: 'center', valign: 'middle', fontSize: 16, color: C.white });
    shape(s, 'ellipse', m.x + 0.49, 3.34, 0.22, 0.22, m.color, { transparency: 65 });
    shape(s, 'ellipse', m.x + 0.54, 3.39, 0.12, 0.12, m.color);
    card(s, m.cardX, m.cardY, 2.79, 2.06, C.white, m.r, shadow(22, 0.11));
    text(s, 'Storage', { x: m.cardX + 0.5, y: m.cardY + 0.66, w: 1.75, h: 0.47, align: 'center', fontSize: 16, bold: true, color: C.body, lineSpacingMultiple: 1.5 });
    body12(s, 'Lorem ipsum dolor sit amet, adipiscing', m.cardX + 0.34, m.cardY + 1.16, 2.05, 0.61, { align: 'center' });
  });
}

/* ==================================================== 22. Table Logistic */

function slide22(s) {
  heading(s, [['Table ', C.ink], ['Logistic', C.orange]], { x: 1.12, y: 0.87, w: 11.08, h: 0.84, align: 'center' });
  shape(s, 'rect', 1.22, 1.76, 10.9, 4.28, C.white, shadow(26, 0.12));
  shape(s, 'rect', 1.22, 5.86, 10.9, 0.72, C.orange);

  const head = ['Request a Quote', 'Freight Type', 'Delivery City', 'Price', 'Weight', 'Height'];
  const rows = [
    ['Your Text Here #1', 'Your Text Here ', 'Your Text Here ', '$ 360', '0.5-2kg', '5cm'],
    ['Your Text Here #2', 'Your Text Here ', 'Your Text Here ', '$ 310', '0.5-2kg', '8cm'],
    ['Your Text Here #3', 'Your Text Here ', 'Your Text Here ', '$ 460', '0.5-3kg', '12cm'],
    ['Your Text Here #4', 'Your Text Here ', 'Your Text Here ', '$ 550', '3-5kg', '25cm'],
    ['Your Text Here #5', 'Your Text Here ', 'Your Text Here ', '$ 370', '4-7kg', '15cm'],
  ];
  const tbl = [head.map(function (h) {
    return { text: h, options: { fill: { color: C.navy }, color: C.white, fontSize: 14, align: 'center' } };
  })];
  rows.forEach(function (r, i) {
    const bg = i % 2 ? C.paper : C.white;
    tbl.push(r.map(function (cell, j) {
      return { text: cell, options: { fill: { color: bg }, color: C.body, fontSize: 12, align: j === 0 ? 'left' : 'center' } };
    }));
  });
  s.addTable(tbl, {
    x: 1.505, y: 2.009, w: 10.323, colW: [3.856, 1.486, 1.472, 1.143, 1.183, 1.183], rowH: 0.645,
    valign: 'middle', fontFace: F.body, border: { type: 'solid', color: C.silver, pt: 0.5 }, margin: [0, 0.09, 0, 0.09],
  });
  body12(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua',
    1.94, 6.11, 9.85, 0.35, { color: C.white });
}

/* ======================================================= 23. SWOT diagram */

function slide23(s) {
  // four arrows meeting in the centre, each capped with a light disc
  const arrows = [
    { x: 5.74, y: 1.65, w: 1.86, h: 2.4, dir: 'downArrow', fill: C.orange, cx: 6.67, cy: 2.31 },
    { x: 4.33, y: 3.05, w: 2.4, h: 1.86, dir: 'rightArrow', fill: C.navy, cx: 5.0, cy: 3.98 },
    { x: 6.6, y: 3.05, w: 2.4, h: 1.86, dir: 'leftArrow', fill: C.sand, cx: 8.34, cy: 3.98 },
    { x: 5.74, y: 3.91, w: 1.86, h: 2.4, dir: 'upArrow', fill: C.red, cx: 6.67, cy: 5.65 },
  ];
  arrows.forEach(function (a) {
    shape(s, a.dir, a.x, a.y, a.w, a.h, a.fill);
    shape(s, 'ellipse', a.cx - 0.545, a.cy - 0.545, 1.09, 1.09, 'EEF3F7');
  });
  [['W', 6.67, 2.31], ['S', 5.0, 3.98], ['O', 8.34, 3.98], ['T', 6.67, 5.65]].forEach(function (l) {
    text(s, l[0], { x: l[1] - 0.6, y: l[2] - 0.6, w: 1.2, h: 1.2, align: 'center', valign: 'middle', fontFace: F.head, fontSize: 66, color: C.body });
  });

  const quads = [
    { title: 'Strengths', color: C.navy, tx: 0.61, ty: 3.42, by: 3.93, ly: 4.65, bar: 0.71, barY: 4.95, font: F.head },
    { title: 'Weaknesses', color: C.orange, tx: 2.64, ty: 0.68, by: 1.18, ly: 1.98, bar: 2.74, barY: 2.38, font: F.inter },
    { title: 'Opportunities', color: C.sand, tx: 9.33, ty: 2.68, by: 3.19, ly: 3.93, bar: 9.43, barY: 4.23, font: F.head },
    { title: 'Threats', color: C.red, tx: 7.5, ty: 5.66, by: 6.12, ly: 6.83, bar: 7.6, barY: 7.14, font: F.inter },
  ];
  quads.forEach(function (q) {
    text(s, q.title, { x: q.tx, y: q.ty, w: 3.69, h: 0.6, fontFace: q.font, fontSize: 32, bold: true, color: q.color, valign: 'middle' });
    leadIn(s, 'is simply dummy text of the printing and typesetting industry', q.tx, q.by, 3.62, 0.65);
    text(s, 'Education skill #1', { x: q.tx, y: q.ly, w: 1.6, h: 0.28, fontSize: 10, color: C.body });
    card(s, q.bar, q.barY, 2.88, 0.05, C.grey, 0.025, { transparency: 75 });
    card(s, q.bar, q.barY, 2.27, 0.05, q.color, 0.025);
    shape(s, 'ellipse', q.bar + 2.16, q.barY - 0.08, 0.21, 0.21, q.color, { transparency: 70 });
    shape(s, 'ellipse', q.bar + 2.21, q.barY - 0.03, 0.11, 0.11, q.color);
  });
}

/* ============================================= 24. Logistic Office Chart */

function slide24(s) {
  heading(s, [['Logistic ', C.ink], ['Office', C.orange], [' Chart', C.ink]],
    { x: 2.52, y: 0.88, w: 8.3, h: 0.84, align: 'center' });

  const cats = ['5/1/2023', '6/1/2023', '7/1/2023', '8/1/2023', '9/1/2023'];
  s.addChart('line', [
    { name: 'Data 1', labels: cats, values: [30, 25, 25, 15, 10] },
    { name: 'Data 2', labels: cats, values: [15, 20, 15, 25, 20] },
  ], {
    x: 1.3, y: 2.4, w: 10.74, h: 4.35,
    chartColors: [C.navy, C.sand], lineSize: 2.5, lineDataSymbolSize: 7,
    showLegend: false, showTitle: false,
    catAxisLabelFontSize: 10, valAxisLabelFontSize: 10,
    catAxisLabelColor: C.body, valAxisLabelColor: C.body,
    valGridLine: { color: 'F2F2F2', size: 1 }, catGridLine: { style: 'none' },
    valAxisMaxVal: 35, valAxisMajorUnit: 5,
    catAxisLineColor: C.silver, valAxisLineShow: false,
  });

  // two speech-bubble callouts over the plot
  [{ x: 2.84, y: 2.01, fill: C.navy, amount: '$352,100.00', tipX: 4.68 },
   { x: 8.33, y: 2.24, fill: C.orange, amount: '$217,100.00', tipX: 8.78 }].forEach(function (cb) {
    card(s, cb.x, cb.y, 2.84, 1.02, cb.fill, 0.1);
    shape(s, 'triangle', cb.tipX, cb.y + 1.0, 0.34, 0.26, cb.fill, { rotate: 180 });
    text(s, cb.amount, { x: cb.x + 0.25, y: cb.y + 0.1, w: 2.48, h: 0.47, fontFace: F.heeboBlack, fontSize: 16, bold: true, color: C.white, valign: 'middle' });
    body12(s, 'Lorem Ipsum' + NBSP + 'is dummy', cb.x + 0.24, cb.y + 0.53, 2.67, 0.34, { color: C.white });
  });
}

/* ============================================ 25. Bar Chart Infographic */

function slide25(s) {
  heading(s, [['Bar ', C.ink], ['Chart', C.orange], [' Infographic', C.ink]],
    { x: 1.12, y: 0.87, w: 11.08, h: 0.84, align: 'center' });

  const cats = ['Category 1', 'Category 2', 'Category 3', 'Category 4'];
  const panels = [
    { x: 1.1, chartX: 1.39, chartY: 2.76, chartH: 2.62, tagX: 4.34, tag: 'June', tagFill: C.navy,
      colors: [C.navy, C.sand, C.orange], vals: [[5, 3, 4, 5], [4, 3, 5, 4], [3, 4, 5, 4]] },
    { x: 6.86, chartX: 7.15, chartY: 2.64, chartH: 2.76, tagX: 10.35, tag: 'July', tagFill: C.orange,
      colors: [C.navy, C.orange, C.sand], vals: [[5, 4, 2, 3], [2, 4, 3, 5], [3, 4, 3, 4]] },
  ];
  panels.forEach(function (p) {
    card(s, p.x, 2.35, 5.37, 3.32, C.white, 0.137, shadow(24, 0.12));
    s.addChart('bar', p.vals.map(function (v, i) { return { name: 'Series ' + (i + 1), labels: cats, values: v }; }), {
      x: p.chartX, y: p.chartY, w: 4.79, h: p.chartH,
      chartColors: p.colors, barGapWidthPct: 219, barOverlapPct: -27,
      showLegend: false, showTitle: true, title: 'Chart Title', titleFontSize: 13, titleColor: C.body,
      catAxisLabelFontSize: 9, valAxisLabelFontSize: 9,
      catAxisLabelColor: C.body, valAxisLabelColor: C.body,
      valGridLine: { color: 'F2F2F2', size: 1 }, valAxisMaxVal: 6, valAxisMajorUnit: 1,
      catAxisLineColor: C.silver, valAxisLineShow: false,
    });
    card(s, p.tagX, 2.12, 1.59, 0.43, p.tagFill, 0.08);
    text(s, p.tag, { x: p.tagX + 0.12, y: 2.12, w: 1.35, h: 0.43, align: 'center', valign: 'middle', fontSize: 16, color: C.white });
  });

  leadIn(s, 'is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been Lorem Ipsum' + NBSP + 'is simply dummy text of the printing and typesetting industry. ',
    1.56, 5.94, 10.25, 0.9);
}

/* ================================= 26. Pricing Table Delivery Logistic */

function slide26(s) {
  heading(s, [['Pricing ', C.ink], ['Table Delivery ', C.orange], ['Logistic', C.ink]],
    { x: 2.52, y: 0.94, w: 8.3, h: 0.84, bold: true, align: 'center' });

  const plans = [
    { x: 1.23, y: 2.15, w: 3.39, h: 4.16, bg: C.white, fg: C.body, dot: C.navy, tick: C.white, r: 0.182,
      name: 'Basic', price: '$34', dx: 0.44, wide: false,
      items: ['Troelly for moving', 'Box for Packing', 'Safe and Secure Delivery', 'Tape for fragile Item'] },
    { x: 4.84, y: 1.95, w: 3.67, h: 4.98, bg: C.navy, fg: C.white, dot: C.paper, tick: C.navy, r: 0.197,
      name: 'Medium', price: '$34', dx: 0.41, wide: true,
      items: ['Box for Packing', 'Troelly for moving', 'Safe and Secure Delivery', 'Tape for fragile Item'] },
    { x: 8.72, y: 2.1, w: 3.25, h: 4.21, bg: C.white, fg: C.body, dot: C.sand, tick: C.white, r: 0.175,
      name: 'Advance', price: '$54', dx: 0.41, wide: false,
      items: ['Box for Packing', 'Troelly for moving', 'Tape for fragile Item', 'Safe and Secure Delivery'] },
  ];
  plans.forEach(function (p) {
    card(s, p.x, p.y, p.w, p.h, p.bg, p.r, p.bg === C.white ? shadow(24, 0.12) : {});
    text(s, p.name, { x: p.x + 0.31, y: p.y + 0.27, w: 2.5, h: 0.6, fontSize: 20, bold: true, color: p.fg, valign: 'middle', lineSpacingMultiple: 1.5 });
    text(s, p.price, { x: p.x + 0.31, y: p.y + 0.68, w: 2.0, h: p.wide ? 1.23 : 1.03, fontSize: 40, bold: true, color: p.fg, valign: 'middle', lineSpacingMultiple: 1.5 });
    text(s, '/month', { x: p.x + 1.42, y: p.y + 1.09, w: 1.8, h: 0.5, fontSize: 16, color: p.fg, valign: 'middle', lineSpacingMultiple: 1.5 });
    const step = p.wide ? 0.585 : 0.475, d = p.wide ? 0.25 : 0.21;
    p.items.forEach(function (item, i) {
      const y = p.y + (p.wide ? 2.27 : 1.89) + i * step;
      checkDot(s, p.x + p.dx, y, d, p.dot, p.tick);
      text(s, item, { x: p.x + p.dx + 0.42, y: y - 0.1, w: 2.6, h: 0.42, fontSize: 12, color: p.fg, valign: 'middle' });
    });
  });
}

/* ============================================= 27. Funnel Infographic */

function slide27(s) {
  heading(s, [['Funnel ', C.ink], ['Infographic', C.orange]], { x: 5.58, y: 1.16, w: 6.77, h: 0.84 });
  leadIn(s, 'is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been Lorem Ipsum' + NBSP + 'is simply dummy text of the printing and typesetting industry. ',
    5.58, 2.12, 7.09, 1.17);

  // three stacked funnel tiers: dark shell, coloured band, elliptical rims
  const tiers = [
    { x: 1.25, w: 3.68, y: 2.26, h: 1.08, fill: C.navy, label: 'Mission One', size: 28 },
    { x: 1.61, w: 2.97, y: 3.54, h: 0.87, fill: C.orange, label: 'Mission Two', size: 24 },
    { x: 1.87, w: 2.46, y: 4.71, h: 0.7, fill: C.sand, label: 'Mission Three', size: 20 },
  ];
  tiers.forEach(function (t) {
    shape(s, 'ellipse', t.x, t.y - 0.13, t.w, 0.26, '595959');           // outer rim
    shape(s, 'trapezoid', t.x, t.y, t.w, t.h, '404040');                  // dark bowl
    shape(s, 'ellipse', t.x + 0.06, t.y - 0.11, t.w - 0.12, 0.22, '404040'); // rim inner face
    shape(s, 'trapezoid', t.x + 0.05, t.y + 0.12, t.w - 0.1, t.h * 0.62, t.fill);
    text(s, t.label, { x: t.x, y: t.y + 0.12, w: t.w, h: t.h * 0.62, align: 'center', valign: 'middle', fontSize: t.size, bold: true, color: C.white });
  });
  shape(s, 'ellipse', 2.13, 5.79, 1.95, 0.15, 'CFCFCF');

  const stats = [
    { x: 5.54, w: 2.27, color: C.navy, label: 'Mission One' },
    { x: 7.91, w: 2.3, color: C.orange, label: 'Mission Two' },
    { x: 10.29, w: 2.21, color: C.sand, label: 'Mission Three' },
  ];
  stats.forEach(function (st) {
    card(s, st.x, 3.31, st.w, 2.67, C.white, 0.122, shadow(22, 0.11));
    ring(s, st.x + st.w / 2, 4.04, 0.8, st.color, 250, '78+');
    text(s, st.label, { x: st.x, y: 4.53, w: st.w, h: 0.45, align: 'center', fontSize: 16, bold: true, color: C.body, lineSpacingMultiple: 1.5 });
    body12(s, 'Lorem Ipsum' + NBSP + 'is simply dummy text of the', st.x - 0.02, 5.02, st.w + 0.06, 0.6, { align: 'center' });
  });
}

/* ======================================= 28. Puzzle Infographic (6 wedges) */

function slide28(s) {
  heading(s, [['Puzzle ', C.ink], ['Infographic', C.orange]],
    { x: 3.1, y: 0.98, w: 7.13, h: 0.84, bold: true, align: 'center' });

  // circle carved into six blocks: rounded quarter wedges top and bottom,
  // straight slabs across the middle. A pieWedge's square corner sits bottom
  // -left, so the flips place each quarter against the centre of the circle.
  const wedges = [
    { x: 4.73, y: 2.33, w: 1.76, h: 1.47, fill: C.navy, kind: 'pieWedge', flipH: false, flipV: true, tx: 4.96, ty: 3.02, px: 6.4, py: 2.75 },
    { x: 6.86, y: 2.33, w: 1.78, h: 1.47, fill: C.orange, kind: 'pieWedge', flipH: true, flipV: true, tx: 6.9, ty: 3.02, px: 6.98, py: 2.75 },
    { x: 4.55, y: 3.94, w: 2.17, h: 1.24, fill: C.red, kind: 'rect', tx: 4.72, ty: 4.38, px: 6.55, py: 4.86 },
    { x: 6.86, y: 3.94, w: 1.94, h: 1.24, fill: C.sand, kind: 'rect', tx: 7.17, ty: 4.38, px: 8.6, py: 4.86 },
    { x: 4.73, y: 5.28, w: 1.76, h: 1.47, fill: C.maroon, kind: 'pieWedge', flipH: false, flipV: false, tx: 4.96, ty: 5.74, px: 6.4, py: 5.4 },
    { x: 6.86, y: 5.28, w: 1.78, h: 1.47, fill: C.taupe, kind: 'pieWedge', flipH: true, flipV: false, tx: 6.87, ty: 5.74, px: 6.98, py: 5.4 },
  ];
  wedges.forEach(function (w) {
    shape(s, w.kind, w.x, w.y, w.w, w.h, w.fill, { flipH: w.flipH, flipV: w.flipV });
    text(s, 'Your text', { x: w.tx, y: w.ty, w: 1.47, h: 0.36, align: 'center', fontSize: 15, color: C.white });
    // small plus badge marking each interlocking tab
    text(s, G.plus, { x: w.px - 0.14, y: w.py - 0.14, w: 0.28, h: 0.28, align: 'center', valign: 'middle', fontSize: 11, color: C.white, fontFace: F.sym });
  });
  // knobs that make the quarters read as interlocking puzzle pieces
  [[5.42, 3.55, 0.48, 0.34, C.navy], [7.43, 3.55, 0.48, 0.34, C.orange],
   [6.47, 4.35, 0.34, 0.48, C.red], [7.43, 5.05, 0.48, 0.34, C.sand],
   [5.42, 5.22, 0.48, 0.34, C.maroon]].forEach(function (k) {
    shape(s, 'ellipse', k[0], k[1], k[2], k[3], k[4]);
  });

  const labels = [
    { t: 'Option One', c: C.navy, x: 1.29, y: 2.29, bx: 1.78, bw: 2.47, align: 'right' },
    { t: 'Option Four', c: C.red, x: 0.92, y: 3.99, bx: 1.49, bw: 2.29, align: 'right' },
    { t: 'Option Five', c: C.maroon, x: 1.29, y: 5.52, bx: 1.78, bw: 2.47, align: 'right' },
    { t: 'Option Two', c: C.orange, x: 9.02, y: 2.29, bx: 9.02, bw: 2.21, align: 'left' },
    { t: 'Option Three', c: C.sand, x: 9.39, y: 3.91, bx: 9.39, bw: 2.59, align: 'left' },
    { t: 'Option Six', c: C.taupe, x: 9.02, y: 5.52, bx: 9.02, bw: 2.5, align: 'left' },
  ];
  labels.forEach(function (l) {
    text(s, l.t, { x: l.x, y: l.y, w: 2.96, h: 0.4, fontSize: 18, bold: true, color: l.c, align: l.align });
    body12(s, 'Lorem ipsum dolor sit amet, consectetur', l.bx, l.y + 0.49, l.bw, 0.65, { fontSize: 14, align: l.align });
  });
}

/* ================================= 29. Puzzle Infographic (isometric tiles) */

function slide29(s) {
  heading(s, [['Puzzle ', C.ink], ['Infographic', C.orange]], { x: 3.37, y: 0.99, w: 6.53, h: 0.84, align: 'center' });

  // five diamond tiles zig-zagging across the middle, each on a darker plinth
  const tiles = [
    { x: 2.78, y: 3.68, fill: C.navy, side: '222A43', glyph: G.home },
    { x: 4.0, y: 4.38, fill: C.orange, side: 'BB6328', glyph: G.box },
    { x: 5.2, y: 3.68, fill: C.sand, side: 'C37448', glyph: G.cart },
    { x: 6.42, y: 4.38, fill: C.red, side: 'C02C0B', glyph: G.bus },
    { x: 7.62, y: 3.68, fill: C.maroon, side: '560C05', glyph: G.map },
  ];
  tiles.forEach(function (t) {
    shape(s, 'diamond', t.x, t.y + 0.1, 2.42, 1.4, t.side);
    shape(s, 'diamond', t.x, t.y, 2.42, 1.4, t.fill);
    text(s, t.glyph, { x: t.x + 0.86, y: t.y + 0.45, w: 0.7, h: 0.5, align: 'center', valign: 'middle', fontSize: 20, color: C.white, fontFace: F.sym });
  });

  const captions = [
    { t: 'Mission One', x: 2.99, y: 2.41, bx: 2.98, by: 2.83 },
    { t: 'Mission Three', x: 5.49, y: 2.37, bx: 5.49, by: 2.83 },
    { t: 'Mission Five', x: 8.08, y: 2.37, bx: 8.11, by: 2.85 },
    { t: 'Mission Two', x: 2.94, y: 5.75, bx: 3.06, by: 6.09 },
    { t: 'Mission Four', x: 8.24, y: 5.75, bx: 8.25, by: 6.21 },
  ];
  captions.forEach(function (c) {
    text(s, c.t, { x: c.x, y: c.y, w: 2.48, h: 0.46, fontSize: 16, bold: true, color: C.body, valign: 'middle', lineSpacingMultiple: 1.5 });
    body12(s, 'Lorem Ipsum' + NBSP + 'is simply dummy text of the', c.bx, c.by, 2.33, 0.6);
  });
}

/* ==================================================== 30. MAP Delivery */

function slide30(s) {
  // stylised South-America silhouette: a pale bulge over a pale tail, with two
  // picked-out countries in the brand colours
  shape(s, 'ellipse', 6.1, 0.58, 3.3, 2.8, C.paper);
  shape(s, 'triangle', 6.25, 2.5, 1.7, 4.4, C.paper, { flipV: true });
  shape(s, 'pentagon', 6.66, 1.28, 3.3, 3.05, C.navy, { rotate: 12 });
  shape(s, 'triangle', 6.5, 3.3, 1.3, 3.3, C.orange, { flipV: true });

  heading(s, [['MAP ', C.ink], ['Delivery', C.orange]], { x: 1.02, y: 2.55, w: 5.32, h: 0.84 });
  leadIn(s, 'is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been Lorem Ipsum' + NBSP + 'is simply dummy text of the printing and typesetting industry. ',
    0.92, 3.83, 4.92, 1.17);

  [{ x: 9.4, y: 2.24, w: 3.42, color: C.navy, val: '78+', tail: 9.09, ty: 2.4, rx: 10.05, ry: 2.9, tx: 10.6, bx: 10.64, by: 2.84, body: 'Lorem Ipsum' + NBSP + 'is simply dummy text of thez' },
   { x: 7.46, y: 4.68, w: 3.52, color: C.orange, val: '65+', tail: 7.13, ty: 4.82, rx: 8.1, ry: 5.33, tx: 8.65, bx: 8.64, by: 5.26, body: 'Lorem Ipsum' + NBSP + 'is simply dummy text of the' }]
    .forEach(function (cb) {
      card(s, cb.x, cb.y, cb.w, 1.36, C.white, 0.073, shadow(22, 0.12));
      shape(s, 'triangle', cb.tail, cb.ty, 0.35, 0.28, C.white, { rotate: 270 });
      ring(s, cb.rx, cb.ry, 0.8, cb.color, 250, cb.val, { bold: true });
      text(s, 'Client Worldwide', { x: cb.tx, y: cb.y + 0.13, w: 2.48, h: 0.47, fontSize: 16, color: C.body, valign: 'middle', lineSpacingMultiple: 1.5 });
      body12(s, cb.body, cb.bx, cb.by, 2.33, 0.6);
    });
}

/* ==================================================== 31. Contact Us */

function slide31(s) {
  shape(s, 'round2SameRect', 0.67, 3.75, 12.0, 3.0, C.navy, { flipV: true, rectRadius: 0.5 });
  text(s, 'Contact Us', { x: 1.62, y: 4.58, w: 3.57, h: 0.91, fontFace: F.head, fontSize: 48, bold: true, color: C.white, valign: 'middle' });
  body12(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Lorem ipsum dolor ',
    1.62, 5.4, 4.01, 0.68, { fontSize: 14, color: C.white });
  s.addShape('line', { x: 6.67, y: 4.68, w: 0, h: 1.64, line: { color: C.white, width: 1, dashType: 'sysDot' } });
  text(s, 'Metropolitan City of Venice, Italy',
    { x: 7.31, y: 4.24, w: 4.4, h: 0.49, fontFace: F.head, fontSize: 18, bold: true, color: C.white, valign: 'middle' });
  [[G.mail, 'logistic@account.com', 4.82], [G.plane, '@logisticaccount', 5.25], [G.target, '@logistictravel', 5.71]]
    .forEach(function (row) {
      text(s, row[0], { x: 7.32, y: row[2], w: 0.4, h: 0.34, align: 'center', valign: 'middle', fontSize: 14, color: C.white, fontFace: F.sym });
      body12(s, row[1], 7.87, row[2], 4.16, 0.34, { color: C.white });
    });
}

/* ==================================================== 32. Thank You */

function slide32(s) {
  // vertical gradient, pale slate at the top down to deep navy
  ramp(s, 0, 0, 13.334, 7.5, '9197A7', '222944', 48, 'y');
  text(s, 'Thank You', { x: 1.72, y: 2.45, w: 9.88, h: 2.42, align: 'center', valign: 'middle', fontFace: F.head, fontSize: 138, color: C.white });
  text(s, 'for your attention', { x: 2.3, y: 4.47, w: 8.74, h: 0.71, align: 'center', fontSize: 36, color: C.white });
}

/* ------------------------------------------------------------------ build */

const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
  slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
  slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24,
  slide25, slide26, slide27, slide28, slide29, slide30, slide31, slide32];

function build() {
  const pptx = new PptxGenJS();
  // 13 1/3 x 7.5 in (16:9); the extra decimals land exactly on 12192000 EMU
  pptx.defineLayout({ name: 'DECK', width: 40 / 3, height: 7.5 });
  pptx.layout = 'DECK';
  pptx.theme = { headFontFace: F.head, bodyFontFace: F.body };
  pptx.title = 'Delivery - Logistic Presentation Template';

  pptx.defineSlideMaster({
    title: 'BASE',
    background: { color: C.white },
    slideNumber: { x: 0.195, y: 6.965, w: 0.61, h: 0.37, align: 'center', fontFace: F.head, fontSize: 16, bold: true, color: '969DB4' },
  });

  BUILDERS.forEach(function (builder) {
    builder(pptx.addSlide({ masterName: 'BASE' }));
  });

  return pptx.writeFile({ fileName: path.join(__dirname, '0bfb8585-a2bc-45d3-8f11-f42c0a238e7b_grok_final.pptx') });
}

build().then(function (f) { console.log('wrote ' + f); }, function (e) { console.error(e); process.exit(1); });
