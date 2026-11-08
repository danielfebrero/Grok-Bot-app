/**
 * "Mineral Exploration" deck — 30 slides, 10 x 5.625 in.
 * Rebuilt with pptxgenjs only. Photographs are replaced by flat colour blocks.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

const OUT = path.join(__dirname, '0c6025a2-4832-4fa4-a2e7-a88be599eafb_grok_final.pptx');

/* ------------------------------------------------------------------ theme */

const C = {
  yellow: 'F8B908',
  blue: '03417F',
  navy: '02315F', // accent4 @ 75% luminance — title-slide background
  white: 'FFFFFF',
  ink: '0D0D0D',
  black: '000000',
  gray: '808080',
  silver: 'BFBFBF',
  charcoal: '262626',
  ring: 'D9D9D9',
  track: 'F2F2F2',
  photo: 'FF0028', // stand-in colour for the deck's photographs
};

const F = {
  head: 'Roboto Slab',
  body: 'Open Sans',
  semi: 'Open Sans SemiBold',
  alt: 'Montserrat',
  thanks: 'Heebo ExtraBold',
};

const SZ = { title: 24, body: 8.25, micro: 7.88, small: 9, lead: 10.5, big: 18 };

/* --------------------------------------------------------------- shadows */

const SH_CARD = (blur, opacity) => ({ type: 'outer', blur, offset: 0, angle: 90, color: C.black, opacity });
const SH_DROP = { type: 'outer', blur: 20, offset: 10, angle: 45, color: C.black, opacity: 0.2 };
const SH_ORB = { type: 'outer', blur: 10, offset: 5, angle: 45, color: C.black, opacity: 0.2 };

/* ------------------------------------------------------------- primitives */

let S; // pptxgenjs ShapeType lookup, assigned in build()

/** Flat colour block standing in for a photograph. */
function photo(s, x, y, w, h) {
  s.addShape(S.rect, { x, y, w, h, fill: { color: C.photo }, line: { type: 'none' }, objectName: '[image]' });
}

function box(s, x, y, w, h, fill, opts) {
  s.addShape(S.rect, Object.assign({ x, y, w, h, fill: { color: fill }, line: { type: 'none' } }, opts));
}

/** Text block: top-anchored like every text frame in the source deck. */
function text(s, runs, o) {
  s.addText(runs, Object.assign({ valign: 'top', fontFace: F.body, fontSize: SZ.body, color: C.gray }, o));
}

/** Grey 8.25pt paragraph at 150% leading — the deck's default body copy. */
function copy(s, str, x, y, w, h, o) {
  text(s, str, Object.assign({ x, y, w, h, lineSpacingMultiple: 1.5 }, o));
}

/** Two-line Roboto Slab headline; `blue` part is coloured, trailing period is ink. */
function headline(s, lines, x, y, w, h, o) {
  const runs = [];
  lines.forEach((ln, i) => {
    const dark = ln.dark || '';
    const accent = ln.accent || '';
    const last = i === lines.length - 1;
    if (dark) runs.push({ text: dark, options: { color: (o && o.color) || C.ink, breakLine: !accent && !last } });
    if (accent) runs.push({ text: accent, options: { color: (o && o.accent) || C.blue, breakLine: !ln.dot && !last } });
    if (ln.dot) runs.push({ text: '.', options: { color: (o && o.color) || C.ink, breakLine: !last } });
  });
  text(s, runs, Object.assign({ x, y, w, h, fontFace: F.head, fontSize: SZ.title, bold: true, color: C.ink }, o));
}

/** "MINERAL / EXPLORATION" corner lockup. variant: 'light' | 'dark'. */
function brand(s, o) {
  const opt = o || {};
  const dark = opt.variant === 'dark';
  text(s, [
    { text: dark ? 'MINERAL ' : 'MINERAL', options: { color: dark ? C.white : C.silver, breakLine: dark } },
    ...(dark ? [] : [{ text: ' ', options: { color: C.charcoal, breakLine: true } }]),
    { text: 'EXPLORATION', options: { bold: true, color: dark ? C.white : C.yellow } },
  ], {
    x: opt.x !== undefined ? opt.x : 8.159,
    y: opt.y !== undefined ? opt.y : 0.31,
    w: opt.w !== undefined ? opt.w : 1.642,
    h: 0.366,
    align: opt.align || 'right',
    fontSize: SZ.micro,
  });
}

/** White (or coloured) info card: rounded-square panel with soft drop shadow. */
function card(s, x, y, w, h, o) {
  const opt = o || {};
  box(s, x, y, w, h, opt.fill || C.white, { shadow: SH_CARD(opt.blur || 45, opt.opacity || 0.29) });
}

/** "More info" pill button. */
function moreInfo(s, x, y, color) {
  s.addText('More info', {
    shape: S.roundRect, x, y, w: 0.889, h: 0.263,
    fill: { type: 'none' }, line: { color, width: 0.75 },
    align: 'center', valign: 'middle', fontFace: F.body, fontSize: 7.5, bold: true, color,
  });
}

/**
 * Card caption pair used all over the deck:
 * small coloured "Description Here" kicker above a bold title.
 */
function caption(s, x, y, o) {
  const opt = o || {};
  text(s, opt.kicker || 'Description Here', {
    x, y, w: 1.549, h: 0.234, fontFace: F.semi, fontSize: SZ.micro, bold: true,
    color: opt.kickerColor || C.blue,
  });
  text(s, opt.title, {
    x: x + 0.031, y: y + 0.183, w: opt.titleW || 1.549, h: opt.titleH || 0.278,
    fontFace: F.semi, fontSize: opt.titleSize || SZ.lead, bold: true, color: opt.titleColor || C.ink,
  });
}

/** Big number over a "Description" label over a short grey blurb (stat panel). */
function statPanel(s, x, y, o) {
  const opt = o || {};
  text(s, opt.value, {
    x: x - 0.019, y, w: 2.412, h: 0.404, align: 'center',
    fontSize: SZ.big, bold: true, color: opt.valueColor || C.blue,
  });
  text(s, 'Description', {
    x, y: y + 0.383, w: 2.412, h: 0.252, align: 'center',
    fontSize: SZ.small, bold: true, color: opt.labelColor || C.ink,
  });
  copy(s, 'Lorem Ipsum\u00a0is simply dummy text of the printing and', x + 0.519, y + 0.649, 1.375, 0.704, {
    align: 'center', color: opt.bodyColor || C.gray,
  });
}

/* ------------------------------------------------------------ progress ring */

// `track` is the donut's rectRadius that yields the source deck's very thin grey ring.
const RING = { size: 0.968, plate: 0.871, band: 0.763, dot: 0.099, thickness: 0.1373, track: 0.0483 };

/**
 * Circular percentage gauge: white disc, grey track, coloured arc, end dot, label.
 * `from`/`to` are DrawingML angles (0 = east, clockwise).
 */
function gauge(s, x, y, o) {
  const cx = x + RING.size / 2;
  const cy = y + RING.size / 2;
  const dotA = (o.to * Math.PI) / 180;
  s.addShape(S.ellipse, { x, y, w: RING.size, h: RING.size, fill: { color: C.white }, shadow: SH_ORB });
  s.addShape(S.ellipse, {
    x: cx - RING.plate / 2, y: cy - RING.plate / 2, w: RING.plate, h: RING.plate,
    fill: { type: 'none' }, line: { color: C.ring, width: 0.75 },
  });
  s.addShape(S.donut, {
    x: cx - RING.band / 2, y: cy - RING.band / 2, w: RING.band, h: RING.band,
    fill: { color: C.ring }, line: { type: 'none' }, rectRadius: RING.track,
  });
  s.addShape(S.blockArc, {
    x: cx - RING.band / 2, y: cy - RING.band / 2, w: RING.band, h: RING.band,
    fill: { color: o.arc }, line: { type: 'none' },
    angleRange: [o.from, o.to], arcThicknessRatio: RING.thickness,
  });
  s.addShape(S.ellipse, {
    x: cx + 0.356 * Math.cos(dotA) - RING.dot / 2, y: cy + 0.356 * Math.sin(dotA) - RING.dot / 2,
    w: RING.dot, h: RING.dot, fill: { color: o.dot }, line: { type: 'none' },
  });
  text(s, o.label, {
    x: cx - 0.3165, y: cy - 0.128, w: 0.633, h: 0.303,
    align: 'center', fontSize: 12, bold: true, color: C.black,
  });
}

/* ------------------------------------------------------------------- icons */

/** Cog wheel with a hollow centre, optionally re-filled with a solid core. */
function cog(s, x, y, d, color, hole, holeD, core) {
  s.addShape(S.gear6, { x, y, w: d, h: d, fill: { color }, line: { type: 'none' }, rotate: 15 });
  const disc = (dia, fill) => s.addShape(S.ellipse, {
    x: x + d * (0.5 - dia / 2), y: y + d * (0.5 - dia / 2), w: d * dia, h: d * dia,
    fill: { color: fill }, line: { type: 'none' },
  });
  if (hole) disc(holeD, hole);
  if (core) disc(holeD * 0.5, color);
}

/** Big cog holding a person silhouette, with a smaller cog tucked below-left. */
function iconGearPerson(s, x, y, d, color, bg) {
  cog(s, x + d * 0.26, y, d * 0.74, color, bg, 0.56);
  s.addShape(S.ellipse, { x: x + d * 0.565, y: y + d * 0.245, w: d * 0.13, h: d * 0.13, fill: { color }, line: { type: 'none' } });
  s.addShape(S.round2SameRect, {
    x: x + d * 0.52, y: y + d * 0.4, w: d * 0.22, h: d * 0.115,
    fill: { color }, line: { type: 'none' }, rectRadius: d * 0.055,
  });
  cog(s, x, y + d * 0.52, d * 0.44, color, bg, 0.44);
}

/** Globe crossed by meridians with a magnifier sitting on its lower-right. */
function iconGlobeSearch(s, x, y, d, color, bg) {
  const g = d * 0.78;
  s.addShape(S.ellipse, { x, y, w: g, h: g, fill: { color }, line: { type: 'none' } });
  s.addShape(S.ellipse, {
    x: x + g * 0.3, y, w: g * 0.4, h: g,
    fill: { type: 'none' }, line: { color: bg, width: 0.75 },
  });
  s.addShape(S.line, { x, y: y + g / 2, w: g, h: 0, line: { color: bg, width: 0.75 } });
  s.addShape(S.ellipse, {
    x: x + d * 0.42, y: y + d * 0.42, w: d * 0.42, h: d * 0.42,
    fill: { color: bg }, line: { color, width: 1.25 },
  });
  s.addShape(S.rect, {
    x: x + d * 0.78, y: y + d * 0.78, w: d * 0.22, h: d * 0.09,
    fill: { color }, line: { type: 'none' }, rotate: 45,
  });
}

/** Analogue clock face with twelve hour ticks. */
function iconClock(s, x, y, d, color) {
  s.addShape(S.ellipse, { x, y, w: d, h: d, fill: { type: 'none' }, line: { color, width: 2 } });
  s.addShape(S.ellipse, {
    x: x + d * 0.12, y: y + d * 0.12, w: d * 0.76, h: d * 0.76,
    fill: { type: 'none' }, line: { color, width: 1 },
  });
  for (let i = 0; i < 12; i++) {
    const a = (i * Math.PI) / 6;
    const c = Math.cos(a);
    const sn = Math.sin(a);
    s.addShape(S.line, {
      x: x + d * (0.5 + 0.3 * c), y: y + d * (0.5 + 0.3 * sn),
      w: d * 0.07 * c, h: d * 0.07 * sn, line: { color, width: 1 },
      flipH: c < 0, flipV: sn < 0,
    });
  }
  s.addShape(S.line, { x: x + d * 0.5, y: y + d * 0.24, w: 0, h: d * 0.28, line: { color, width: 1.5 } });
  s.addShape(S.line, { x: x + d * 0.5, y: y + d * 0.5, w: d * 0.16, h: d * 0.13, line: { color, width: 1.5 } });
}

/** Simple house glyph. */
function iconHouse(s, x, y, d, color) {
  s.addShape(S.triangle, { x, y, w: d, h: d * 0.55, fill: { color }, line: { type: 'none' } });
  s.addShape(S.rect, { x: x + d * 0.16, y: y + d * 0.5, w: d * 0.68, h: d * 0.5, fill: { color }, line: { type: 'none' } });
  box(s, x + d * 0.28, y + d * 0.62, d * 0.16, d * 0.16, C.white);
  box(s, x + d * 0.52, y + d * 0.62, d * 0.18, d * 0.38, C.white);
}

/** Award rosette with a thumbs-up inside and two ribbon tails. */
function iconBadge(s, x, y, d, color) {
  const r = d * 0.8;
  s.addShape(S.triangle, { x, y: y + d * 0.55, w: r * 0.4, h: d * 0.45, fill: { color }, line: { type: 'none' }, rotate: 180 });
  s.addShape(S.triangle, { x: x + r * 0.6, y: y + d * 0.55, w: r * 0.4, h: d * 0.45, fill: { color }, line: { type: 'none' }, rotate: 180 });
  s.addShape(S.star16, { x, y, w: r, h: r, fill: { color }, line: { type: 'none' } });
  s.addShape(S.ellipse, { x: x + r * 0.2, y: y + r * 0.2, w: r * 0.6, h: r * 0.6, fill: { color: C.white }, line: { type: 'none' } });
  s.addShape(S.rect, { x: x + r * 0.3, y: y + r * 0.45, w: r * 0.4, h: r * 0.22, fill: { color }, line: { type: 'none' } });
  s.addShape(S.rect, { x: x + r * 0.38, y: y + r * 0.3, w: r * 0.1, h: r * 0.18, fill: { color }, line: { type: 'none' } });
}

/** Person / head-and-shoulders silhouette. */
function iconPerson(s, x, y, d, color) {
  s.addShape(S.ellipse, { x: x + d * 0.28, y, w: d * 0.44, h: d * 0.44, fill: { color }, line: { type: 'none' } });
  s.addShape(S.round2SameRect, {
    x: x + d * 0.1, y: y + d * 0.52, w: d * 0.8, h: d * 0.42,
    fill: { color }, line: { type: 'none' }, rectRadius: d * 0.24,
  });
}

/** Crosshair / target. */
function iconTarget(s, x, y, d, color) {
  [[0.4, 0, 0.2, 0.3], [0.4, 0.7, 0.2, 0.3], [0, 0.4, 0.3, 0.2], [0.7, 0.4, 0.3, 0.2]].forEach(function (t) {
    box(s, x + d * t[0], y + d * t[1], d * t[2], d * t[3], color);
  });
  s.addShape(S.ellipse, { x: x + d * 0.18, y: y + d * 0.18, w: d * 0.64, h: d * 0.64, fill: { color: C.white }, line: { color, width: 1.75 } });
  s.addShape(S.ellipse, { x: x + d * 0.38, y: y + d * 0.38, w: d * 0.24, h: d * 0.24, fill: { color }, line: { type: 'none' } });
}

/** Trio of interlocking cogs. */
function iconCogs(s, x, y, d, color) {
  cog(s, x, y + d * 0.28, d * 0.6, color, C.white, 0.34);
  cog(s, x + d * 0.52, y, d * 0.42, color, C.white, 0.34);
  cog(s, x + d * 0.58, y + d * 0.5, d * 0.36, color, C.white, 0.34);
}

/** Hamburger menu bars (top-left of the closing slide). */
function iconMenu(s, x, y, w, h, color) {
  [1, 0.72, 1, 0.62].forEach(function (frac, i) {
    s.addShape(S.roundRect, {
      x, y: y + (h / 3.4) * i, w: w * frac, h: h * 0.14,
      fill: { color }, line: { type: 'none' },
    });
  });
}

/** twitter / facebook / instagram marks. */
function iconSocial(s, x, y, d, color) {
  // bird: crescent body swooping right, triangular wing lifted above it
  s.addShape(S.moon, {
    x: x + d * 0.05, y: y + d * 0.05, w: d * 0.7, h: d * 0.85,
    fill: { color }, line: { type: 'none' }, rotate: 190,
  });
  s.addShape(S.triangle, { x: x + d * 0.45, y: y - d * 0.05, w: d * 0.45, h: d * 0.42, fill: { color }, line: { type: 'none' }, rotate: 140 });
  s.addText('f', {
    x: x + d * 2.0, y: y - d * 0.3, w: d * 0.9, h: d * 1.6,
    align: 'center', valign: 'middle', fontFace: 'Arial', fontSize: 13, bold: true, color,
  });
  // instagram: solid tile with the lens knocked out in the tile's backdrop colour
  s.addShape(S.roundRect, {
    x: x + d * 4.05, y: y - d * 0.05, w: d, h: d,
    fill: { color }, line: { type: 'none' }, rectRadius: 0.025,
  });
  s.addShape(S.ellipse, {
    x: x + d * 4.3, y: y + d * 0.2, w: d * 0.5, h: d * 0.5,
    fill: { color: C.blue }, line: { type: 'none' },
  });
  s.addShape(S.ellipse, {
    x: x + d * 4.38, y: y + d * 0.28, w: d * 0.34, h: d * 0.34,
    fill: { color }, line: { type: 'none' },
  });
}

/* ------------------------------------------------------------ shared copy */

const L = {
  short: 'Lorem Ipsum\u00a0is simply dummy text of the printing and typesetting industry. ',
  mid: "Lorem Ipsum\u00a0is simply dummy text of the printing and typesetting industry. Lorem Ipsum  has been the industry's standard dummy text of the printing and typesetting industry. ",
  tiny: 'Lorem Ipsum\u00a0is simply dummy text of the printing and \u00a0is simply dummy text of the printing and',
};
const LOREM = {
  a: L.mid + "Lorem Ipsum has been been the standard dummy text of the printing and the industry's standard dummy text of the",
  b: L.mid + "Lorem Ipsum has been been the standard dummy text of the printing and the industry's standard dummy text of. text of the printing ",
  c: L.mid + "Lorem Ipsum has been been the standard dummy text of the printing and the industry's text of. the printing and the industry's standard dummy text of. standard dummy text ",
  d: L.mid + 'Lorem Ipsum has',
  e: L.mid + 'Lorem Ipsum',
  f: L.mid + 'Lorem Ipsum standard.',
  g: L.mid + 'Lorem Ipsum dummy text of the printing and typesetting industry. Lorem Ipsum dummy text of the printing and typesetting industry. ',
  h: L.short + "industry's standard dummy text of the printing and typesetting industry. Lorem Ipsum has been been the standard dummy text of the printing and the industry's standard dummy text of the",
};

/* --------------------------------------------------------------- builders */

const slides = [];
const slide = (fn) => slides.push(fn);

// 1 — cover
slide((s) => {
  s.background = { color: C.navy };
  photo(s, 4.399, 0, 5.601, 5.625);
  brand(s, { x: 0.217, y: 0.327, w: 0.986, align: 'left', variant: 'dark' });
  text(s, [
    { text: 'MINERAL', options: { fontFace: F.head, fontSize: 49.5, bold: true, color: C.yellow, breakLine: true } },
    { text: 'EXPLORATION', options: { fontFace: F.alt, fontSize: 27, color: C.white } },
  ], { x: 0.554, y: 1.672, w: 4.924, h: 1.388 });
  box(s, 0.684, 3.379, 4.316, 0.766, C.yellow, { fill: { color: C.yellow, transparency: 20 }, shadow: SH_DROP });
  copy(s, 'Lorem Ipsum\u00a0is simply dummy text of the printing \u00a0is simply dummy text of the printing and typesetting industry. ',
    0.935, 3.585, 3.814, 0.379, { color: C.white, lineSpacingMultiple: 1 });
  copy(s, 'Start Now   >', 0.684, 5.036, 1.531, 0.288, { color: C.white });
});

// 2 — understanding mineral exploration
slide((s) => {
  photo(s, 0, 0, 5, 5.625);
  brand(s);
  headline(s, [{ dark: 'Understanding Mineral' }, { accent: 'Exploration', dot: true }], 5.531, 1.028, 4.469, 0.909);
  box(s, 1.584, 3.424, 1.391, 1.392, C.blue, { fill: { color: C.blue, transparency: 20 }, shadow: SH_DROP });
  gauge(s, 1.795, 3.661, { from: 123.95, to: 345.49, arc: C.yellow, dot: C.blue, label: '67%' });
  moreInfo(s, 5.611, 4.442, C.yellow);
  copy(s, LOREM.a, 5.531, 3.109, 3.705, 1.121);
  box(s, 0, 3.419, 1.391, 1.392, C.yellow, { fill: { color: C.yellow, transparency: 20 }, shadow: SH_DROP });
  gauge(s, 0.211, 3.656, { from: 82.14, to: 345.49, arc: C.blue, dot: C.yellow, label: '80%' });
  text(s, 'Description Here', { x: 5.531, y: 2.264, w: 1.937, h: 0.234, fontFace: F.semi, fontSize: SZ.micro, bold: true, color: C.ink });
  text(s, '4000+', { x: 5.531, y: 2.569, w: 1.937, h: 0.404, fontFace: F.semi, fontSize: SZ.big, bold: true, color: C.ink });
});

// 3 — exploring about our geology
slide((s) => {
  box(s, 7.925, 1.041, 2.075, 3.022, C.blue);
  photo(s, 4.444, 2.194, 4.713, 3.431);
  headline(s, [{ dark: 'Exploring About Our' }, { accent: 'Geology', dot: true }], 0.793, 1.005, 4.027, 0.909);
  copy(s, L.mid + "Lorem Ipsum has been been the standard dummy text of the printing and the industry's standard dummy text of. ",
    0.793, 2.219, 3.037, 1.329);
  card(s, 0.886, 4.124, 5.424, 1.036, { blur: 30, opacity: 0.2 });
  text(s, '15%', { x: 1.194, y: 4.453, w: 0.768, h: 0.404, align: 'center', fontSize: SZ.big, bold: true, color: C.black });
  copy(s, 'Lorem Ipsum\u00a0is simply dummy text of the printing and is', 2.067, 4.303, 1.531, 0.704);
  text(s, '45%', { x: 3.588, y: 4.476, w: 0.768, h: 0.404, align: 'center', fontSize: SZ.big, bold: true, color: C.blue });
  copy(s, 'Lorem Ipsum\u00a0is simply dummy text of the printing and is', 4.587, 4.277, 1.531, 0.704);
  brand(s);
});

// 4 — about business visionary
slide((s) => {
  photo(s, 0, 0, 7.393, 2.812);
  card(s, 6.714, 1.817, 1.806, 1.905, { fill: C.yellow, blur: 30, opacity: 0.2 });
  statPanel(s, 6.43, 2.03, { value: '$1670,000', valueColor: C.white, labelColor: C.white, bodyColor: C.white });
  brand(s);
  headline(s, [{ dark: 'About Business' }, { accent: 'Visionary', dot: true }], 0.861, 3.959, 4.027, 0.909);
  copy(s, LOREM.h, 4.993, 3.962, 4.257, 0.912);
});

// 5 — about business missionary
slide((s) => {
  box(s, 1.233, 4.591, 2.801, 1.034, C.blue);
  brand(s);
  headline(s, [{ dark: 'About Business' }, { accent: 'Missionary', dot: true }], 6.146, 1.346, 3.854, 0.909);
  copy(s, LOREM.d, 6.146, 2.56, 3.037, 0.912);
  copy(s, LOREM.d, 6.146, 3.776, 3.037, 0.912);
  photo(s, 0, 0, 2.648, 4.273);
  photo(s, 2.756, 1.034, 2.648, 4.159);
});

// 6 — commercially extracted minerals
slide((s) => {
  photo(s, 0, 1.03, 4.399, 3.565);
  brand(s);
  headline(s, [{ dark: 'Commercially Extracted' }, { accent: 'Minerals', dot: true }], 4.953, 0.936, 3.894, 1.313);
  card(s, 3.437, 3.72, 1.806, 1.905, { blur: 30, opacity: 0.2 });
  statPanel(s, 3.154, 3.976, { value: '$9670,000', valueColor: C.yellow });
  card(s, 5.44, 3.72, 1.704, 1.905, { fill: C.blue, blur: 30, opacity: 0.2 });
  copy(s, 'Product Market', 5.349, 4.96, 1.887, 0.338, { align: 'center', fontSize: SZ.lead, bold: true, color: C.white });
  copy(s, LOREM.b, 4.953, 2.193, 3.784, 1.121);
  gauge(s, 5.781, 3.9, { from: 123.95, to: 345.49, arc: C.blue, dot: C.blue, label: '67%' });
  card(s, 7.341, 3.72, 1.806, 1.905, { blur: 30, opacity: 0.2 });
  statPanel(s, 7.057, 3.976, { value: '$1670,000', valueColor: C.yellow });
});

// 7 — mining and the resources
slide((s) => {
  box(s, 0, 0, 3.235, 2.205, C.yellow);
  photo(s, 0.818, 1.023, 8.33, 2.398);
  brand(s);
  headline(s, [{ dark: 'Mining And The' }, { accent: 'Resources', dot: true }], 0.861, 3.959, 4.027, 0.909);
  copy(s, LOREM.h, 4.993, 3.962, 4.257, 0.912);
});

// 8 — exploration for mineral involve
slide((s) => {
  box(s, 2.591, 0, 7.409, 4.482, C.blue);
  photo(s, 0.857, 0.984, 4.143, 4.641);
  brand(s, { variant: 'dark' });
  headline(s, [{ dark: 'Exploration For' }, { dark: 'Mineral Involve.' }], 5.487, 1.016, 4.027, 0.909, { color: C.white });
  copy(s, LOREM.b, 5.487, 2.349, 3.784, 1.121, { color: C.white });
  card(s, 3.786, 3.959, 2.197, 0.975);
  caption(s, 4.545, 4.191, { title: 'Biggest Project' });
  card(s, 6.252, 3.974, 2.197, 0.975);
  caption(s, 7.023, 4.206, { title: 'Our Project' });
  iconGlobeSearch(s, 6.565, 4.304, 0.266, C.blue, C.white);
  iconGearPerson(s, 4.069, 4.272, 0.346, C.blue, C.white);
});

// 9 — testing sample and surface or air
slide((s) => {
  box(s, 7.332, 0, 2.668, 5.625, C.yellow);
  brand(s, { variant: 'dark' });
  headline(s, [{ dark: 'Testing Sample And' }, { accent: 'Surface Or Air', dot: true }], 0.791, 0.651, 4.027, 0.909);
  copy(s, L.mid + "been the standard dummy text of the printing and the industry's standard", 0.791, 1.681, 5.55, 0.704);
  copy(s, L.tiny, 7.807, 1.585, 1.84, 0.912, { color: C.white });
  text(s, 'Description', { x: 7.807, y: 1.307, w: 1.994, h: 0.252, fontSize: SZ.small, bold: true, color: C.white });
  copy(s, L.tiny, 7.807, 3.744, 1.84, 0.912, { color: C.white });
  text(s, 'Description', { x: 7.807, y: 3.466, w: 1.994, h: 0.252, fontSize: SZ.small, bold: true, color: C.white });
  photo(s, 0, 2.812, 7.332, 2.812);
});

// 10 — surveying the ground
slide((s) => {
  box(s, 0, 0, 2.06, 2.228, C.yellow);
  photo(s, 0.857, 0.994, 2.792, 3.652);
  photo(s, 3.724, 0.994, 1.276, 3.652);
  brand(s);
  headline(s, [{ dark: 'Surveying The' }, { accent: 'Ground', dot: true }], 6.126, 1.013, 3.784, 0.909);
  copy(s, LOREM.c, 6.126, 2.284, 3.136, 1.537);
  moreInfo(s, 6.126, 4.349, C.blue);
});

// 11 — about value and benefit
slide((s) => {
  photo(s, 5.123, 0.998, 2.398, 4.239);
  photo(s, 7.602, 0.998, 2.398, 4.239);
  headline(s, [{ dark: 'About Value And' }, { accent: 'Benefit', dot: true }], 0.793, 0.984, 4.027, 0.909);
  copy(s, L.mid + 'Lorem Ipsum has been been the', 0.793, 2.154, 3.48, 0.912);
  card(s, 2.048, 3.362, 3.577, 1.371, { blur: 50, opacity: 0.2 });
  text(s, 'Skill', { x: 2.306, y: 3.881, w: 1.376, h: 0.252, fontSize: SZ.small, color: C.ink });
  text(s, 'Fasting', { x: 2.306, y: 4.173, w: 1.376, h: 0.252, fontSize: SZ.small, color: C.ink });
  box(s, 2.994, 3.973, 1.785, 0.127, C.track);
  box(s, 2.994, 3.976, 1.429, 0.124, C.blue);
  box(s, 3.03, 4.188, 1.785, 0.101, C.track);
  box(s, 3.03, 4.188, 1.513, 0.101, C.yellow);
  box(s, 4.423, 3.969, 0.032, 0.143, C.blue);
  box(s, 4.527, 4.17, 0.032, 0.143, C.yellow);
  text(s, '80%', { x: 5.014, y: 3.881, w: 0.418, h: 0.404, fontSize: SZ.small, color: C.ink });
  text(s, '90%', { x: 5.014, y: 4.173, w: 0.418, h: 0.404, fontSize: SZ.small, color: C.ink });
  text(s, 'Description :', { x: 2.314, y: 3.47, w: 1.407, h: 0.298, fontSize: SZ.lead, bold: true, color: C.ink, lineSpacingMultiple: 1.2 });
  brand(s);
});

// 12 — energy asset such as wind turbines
slide((s) => {
  photo(s, 0, 0.432, 3.773, 2.381);
  photo(s, 3.898, 0.432, 3.773, 2.381);
  brand(s);
  card(s, 6.968, 1.027, 2.197, 0.975, { blur: 19, opacity: 0.25 });
  caption(s, 7.727, 1.259, { title: 'Biggest Project' });
  iconGearPerson(s, 7.251, 1.34, 0.346, C.blue, C.white);
  headline(s, [{ dark: 'Energy Asset Such As' }, { accent: 'Wind Turbines', dot: true }], 0.724, 3.701, 4.027, 0.909);
  copy(s, LOREM.h, 5.019, 3.718, 4.257, 0.912);
});

// 13 — impact manage environment
slide((s) => {
  photo(s, 2.182, 0, 2.818, 4.568);
  brand(s);
  box(s, 0, 0, 2.091, 4.568, C.blue);
  copy(s, L.tiny, 0.244, 1.608, 1.597, 1.121, { color: C.white });
  text(s, 'Description', { x: 0.188, y: 0.989, w: 1.994, h: 0.252, fontSize: SZ.small, bold: true, color: C.white });
  copy(s, L.tiny, 0.244, 2.888, 1.597, 1.121, { color: C.white });
  card(s, 4.017, 3.311, 1.806, 1.905, { blur: 30, opacity: 0.2 });
  statPanel(s, 3.733, 3.567, { value: '$9670,000', valueColor: C.yellow });
  headline(s, [{ dark: 'Impact Manage ', accent: 'Environment', dot: true }], 6.126, 1.013, 3.784, 0.909);
  copy(s, LOREM.c, 6.126, 2.284, 3.136, 1.537);
  moreInfo(s, 6.164, 4.051, C.blue);
});

// 14 — consult with landholder or resident
slide((s) => {
  headline(s, [{ dark: 'Consult With Landholder Or' }, { accent: 'Resident', dot: true }], 0.793, 0.984, 4.027, 1.313);
  brand(s);
  copy(s, L.mid + "Lorem Ipsum has been been the text of the printing and typesetting industry. Lorem Ipsum  has been the industry's standard dummy text of the printing and",
    0.793, 2.663, 3.48, 1.329);
  moreInfo(s, 0.861, 4.566, C.blue);
  photo(s, 5, 1.031, 5, 4.594);
});

// 15 — consider public safety
slide((s) => {
  box(s, 3.818, 4.032, 2.975, 1.593, C.blue);
  photo(s, 0.845, 1.03, 4.752, 3.565);
  card(s, 6.218, 4.34, 2.816, 0.975, { blur: 19, opacity: 0.25 });
  caption(s, 7.058, 4.583, { title: 'Management Project', titleW: 2.026 });
  iconGearPerson(s, 6.581, 4.664, 0.346, C.blue, C.white);
  brand(s);
  headline(s, [{ dark: 'Consider Public' }, { accent: 'Safety', dot: true }], 6.126, 1.013, 3.784, 0.909);
  copy(s, L.mid + 'Lorem Ipsum standard dummy text of the printing and typesetting', 6.133, 2.465, 3.136, 1.121);
  card(s, 3.237, 4.346, 2.816, 0.975, { blur: 19, opacity: 0.25 });
  caption(s, 4.076, 4.59, { title: 'Management Project', titleW: 2.026 });
  iconGearPerson(s, 3.6, 4.671, 0.346, C.yellow, C.white);
});

// 16 — what we doing
slide((s) => {
  photo(s, 6.201, 0, 3.799, 5.625);
  const rows = [
    { y: 1.741, fill: C.yellow, label: 'Professional Worker', w: 1.867, h: 0.505, ty: 2.035, icon: iconPerson, ic: C.yellow },
    // the middle (blue) bar casts a tight straight-down shadow, unlike its neighbours
    { y: 2.722, fill: C.blue, label: 'Full Target Mining', w: 1.723, h: 0.303, ty: 3.015, icon: iconTarget, ic: C.blue,
      shadow: { type: 'outer', blur: 4, offset: 3, angle: 90, color: C.black, opacity: 0.4 } },
    { y: 3.702, fill: C.yellow, label: 'Best Mining Plan', w: 1.867, h: 0.303, ty: 3.996, icon: iconCogs, ic: C.yellow },
  ];
  rows.forEach((r) => {
    box(s, 4.421, r.y, 3.012, 0.865, r.fill, { shadow: r.shadow || SH_DROP });
    box(s, 4.63, r.y + 0.199, 0.466, 0.466, C.white, { shadow: SH_DROP });
    text(s, r.label, { x: 5.335, y: r.ty, w: r.w, h: r.h, fontSize: 12, bold: true, color: C.white });
    r.icon(s, 4.755, r.y + 0.31, 0.216, r.ic);
  });
  headline(s, [{ dark: 'What We ', accent: 'Doing', dot: true }], 0.793, 1.302, 4.027, 0.505);
  copy(s, LOREM.e, 0.793, 2.631, 3.28, 0.912);
  copy(s, L.mid + 'Lorem Ipsum ', 0.793, 3.74, 3.28, 0.912);
  brand(s, { x: 0.193, y: 0.326, align: 'left' });
  text(s, 'Description Here :', { x: 0.793, y: 2.227, w: 1.897, h: 0.27, fontSize: SZ.small, bold: true, color: C.ink, lineSpacingMultiple: 1.2 });
});

// 17 — negotiate access and compensation
slide((s) => {
  photo(s, 0, 1.019, 3.795, 2.899);
  photo(s, 3.965, 3.022, 4.554, 2.603);
  card(s, 2.003, 4.283, 2.816, 0.975, { blur: 19, opacity: 0.25 });
  caption(s, 2.842, 4.526, { title: 'Management Project', titleW: 2.026 });
  iconGearPerson(s, 2.365, 4.608, 0.346, C.blue, C.white);
  brand(s);
  headline(s, [{ dark: 'Negotiate Access And' }, { accent: 'Compensation', dot: true }], 4.313, 0.98, 3.784, 0.909);
  copy(s, L.mid + 'Lorem Ipsum standard dummy text of', 4.341, 1.935, 4.875, 0.704);
});

// 18 — professional team
slide((s) => {
  const team = [
    { x: 0.932, name: 'Eleanor Fitzgerald', role: 'Mine Engineer', roleSize: 7.5 },
    { x: 3.897, name: 'Francisco Andrade', role: 'Logistic', roleSize: 7.88 },
    { x: 6.857, name: 'Hannah Morales', role: 'Supervisor', roleSize: 7.88 },
  ];
  team.forEach((t) => box(s, t.x, 1.625, 2.196, 3.205, C.blue));
  team.forEach((t) => photo(s, t.x + 0.181, 1.784, 1.839, 1.839));
  team.forEach((t) => {
    copy(s, t.name, t.x, 3.698, 2.196, 0.372, { align: 'center', fontSize: 12, bold: true, color: C.white });
    text(s, t.role, {
      x: t.x + 0.006, y: 4.042, w: 2.196, h: 0.234, align: 'center',
      fontSize: t.roleSize, charSpacing: 2.25, color: C.white,
    });
    s.addText('MORE PROFILE INFO', {
      shape: S.rect, x: t.x + 0.398, y: 4.361, w: 1.412, h: 0.204,
      fill: { color: C.yellow }, line: { type: 'none' }, shadow: SH_DROP,
      align: 'center', valign: 'middle', fontFace: F.body, fontSize: 6.75, bold: true, color: C.white,
    });
  });
  brand(s);
  headline(s, [{ dark: 'Professional ', accent: 'Team', dot: true }], 1.727, 0.762, 6.552, 0.505, { align: 'center' });
});

// 19 — agreement in place landholders
slide((s) => {
  box(s, 6.191, 4.59, 2.975, 1.035, C.blue);
  photo(s, 0.918, 2.261, 4.312, 2.933);
  photo(s, 5.364, 2.261, 2.662, 2.933);
  headline(s, [{ dark: 'Agreement In Place' }, { accent: 'Landholders', dot: true }], 0.793, 0.984, 4.027, 0.909);
  brand(s);
  card(s, 7.439, 3.781, 2.32, 0.975);
  caption(s, 8.244, 4.013, { title: 'Our Project' });
  iconGlobeSearch(s, 7.786, 4.111, 0.266, C.blue, C.white);
  copy(s, LOREM.f, 4.981, 1.046, 4.311, 0.704);
});

// 20 — any activities authorised
slide((s) => {
  photo(s, 0, 0, 3.235, 5.625);
  photo(s, 5, 1.041, 5, 2.373);
  brand(s);
  card(s, 3.847, 0.436, 1.806, 1.905, { blur: 30, opacity: 0.2 });
  statPanel(s, 3.564, 0.692, { value: '$9670,000', valueColor: C.blue });
  headline(s, [{ dark: 'Any Activities ', accent: 'Authorised', dot: true }], 3.774, 3.859, 5.981, 0.505);
  copy(s, LOREM.f + ' printing and typesetting industry. Lorem Ipsum  has', 3.774, 4.51, 5.416, 0.704);
});

// 21 — why choose us
slide((s) => {
  box(s, 2.643, 0, 7.357, 5.625, C.blue);
  const menu = [
    { y: 0.56, title: 'Our Project', size: 10.13, icon: (a, b) => iconGearPerson(s, a, b, 0.4, C.yellow, C.white), ix: 1.186, iy: 0.876 },
    { y: 1.772, title: '24 Hour', size: 10.13, icon: (a, b) => iconClock(s, a, b, 0.537, C.blue), ix: 1.147, iy: 1.981 },
    { y: 2.967, title: 'Project', size: 10.13, icon: (a, b) => iconHouse(s, a, b, 0.432, C.yellow), ix: 1.186, iy: 3.237 },
    { y: 4.152, title: 'Our System', size: 12, icon: (a, b) => cog(s, a, b, 0.432, C.blue, C.white, 0.56, true), ix: 1.158, iy: 4.408 },
  ];
  menu.forEach((m) => {
    card(s, 0.864, m.y, 2.448, 0.975);
    caption(s, 1.795, m.y + 0.209, { title: m.title, titleSize: m.size, kickerColor: C.ink });
    m.icon(m.ix, m.iy);
  });
  text(s, 'Why Choose Us.', { x: 3.793, y: 0.871, w: 3.525, h: 0.505, fontFace: F.head, fontSize: SZ.title, bold: true, color: C.white });
  brand(s, { variant: 'dark' });
  copy(s, L.short + "Lorem Ipsum  has been the industry's standard dummy text of the printing and dummy text of the printing and typesetting industry. Lorem Ipsum  has been the industry's standard",
    3.833, 1.566, 5.358, 0.704, { color: C.white });
  photo(s, 3.622, 2.812, 6.378, 2.414);
});

// 22 — reasonable the opportunity
slide((s) => {
  photo(s, 0, 1.063, 5, 3.493);
  card(s, 3.982, 2.776, 5.189, 2.849, { blur: 30, opacity: 0.2 });
  brand(s);
  headline(s, [{ dark: 'Reasonable The ', accent: 'Opportunity', dot: true }], 5.529, 0.965, 3.784, 0.909);
  copy(s, "Lorem Ipsum\u00a0is simply dummy text of the printing and typesetting industry. Lorem Ipsum  has been the industry's", 5.529, 2.042, 3.784, 0.496);
  card(s, 2.104, 4.237, 2.816, 0.975, { fill: C.yellow, blur: 19, opacity: 0.25 });
  caption(s, 2.943, 4.481, { title: 'Management Project', titleW: 2.026, kickerColor: C.white, titleColor: C.white });
  iconGearPerson(s, 2.467, 4.562, 0.346, C.white, C.yellow);
  copy(s, LOREM.g, 5.409, 3.888, 3.352, 1.329);
  text(s, 'Description Here :', { x: 5.409, y: 3.302, w: 1.897, h: 0.27, fontSize: SZ.small, bold: true, color: C.ink, lineSpacingMultiple: 1.2 });
  text(s, '15%', { x: 4.354, y: 3.25, w: 0.91, h: 0.353, fontSize: 15, bold: true, color: C.black });
});

// 23 — minimize our impacts
slide((s) => {
  brand(s);
  headline(s, [{ dark: 'Minimize Our ', accent: 'Impacts', dot: true }], 1.727, 0.762, 6.552, 0.505, { align: 'center' });
  copy(s, LOREM.f + " printing and typesetting industry. Lorem Ipsum  has Ipsum  has been the industry's standard dummy text of the printing and typesetting industry. Lorem Ipsum standard. printing and typesetting industry",
    0.821, 4.449, 8.262, 0.704, { align: 'center' });
  photo(s, 5.131, 1.667, 4.048, 2.357);
  photo(s, 0.821, 1.667, 4.048, 2.357);
});

// 24 — protect the our environment
slide((s) => {
  box(s, 3.592, 2.203, 4.538, 1.559, C.blue);
  photo(s, 0.81, 0, 4.213, 5.179);
  gauge(s, 5.452, 2.351, { from: 123.95, to: 326.14, arc: C.blue, dot: C.blue, label: '61%' });
  gauge(s, 6.71, 2.351, { from: 123.95, to: 35.48, arc: C.blue, dot: C.blue, label: '86%' });
  copy(s, 'Production 01', 5.404, 3.373, 1.094, 0.271, { align: 'center', fontSize: 7.5, color: C.white });
  copy(s, 'Production 02', 6.662, 3.368, 1.094, 0.271, { align: 'center', fontSize: 7.5, color: C.white });
  brand(s);
  headline(s, [{ dark: 'Protect The Our' }, { accent: 'Environment', dot: true }], 5.502, 0.978, 3.784, 0.909);
  copy(s, LOREM.f + " been the industry's standard dummy ", 5.552, 4.303, 3.734, 0.912);
});

// 25 — the licensee holder reasonable
slide((s) => {
  box(s, 5, 0, 5, 3.381, C.blue);
  photo(s, 0.862, 2.228, 5.955, 2.916);
  brand(s, { variant: 'dark' });
  headline(s, [{ dark: 'The Licensee Holder ' }, { accent: 'Reasonable', dot: true }], 0.795, 0.922, 3.784, 0.909);
  copy(s, L.mid, 5.471, 1.024, 3.734, 0.704, { color: C.white });
  card(s, 6.242, 3.626, 2.816, 0.975, { blur: 31, opacity: 0.26 });
  caption(s, 7.081, 3.87, { title: 'Management Project', titleW: 2.026 });
  iconGearPerson(s, 6.605, 3.951, 0.346, C.blue, C.white);
});

// 26 — exploration rules company
slide((s) => {
  photo(s, 2.037, 0, 7.963, 4.532);
  card(s, 0, 2.449, 5.205, 2.849, { blur: 30, opacity: 0.2 });
  copy(s, LOREM.e, 0.637, 3.775, 3.916, 0.704, { align: 'center' });
  headline(s, [{ dark: 'Exploration Rules' }, { accent: 'Company', dot: true }], 0.703, 2.725, 3.784, 0.909, { align: 'center' });
  moreInfo(s, 2.15, 4.74, C.blue);
  brand(s, { x: 0.193, y: 0.326, align: 'left' });
  card(s, 5.281, 4.028, 2.32, 0.975);
  caption(s, 6.086, 4.26, { title: 'Our Project' });
  iconGlobeSearch(s, 5.628, 4.358, 0.266, C.yellow, C.white);
});

// 27 — about our suitable development
slide((s) => {
  photo(s, 2.609, 0.476, 3.525, 2.178);
  photo(s, 0, 2.81, 4.444, 2.339);
  card(s, 0.872, 1.062, 2.32, 0.975);
  caption(s, 1.677, 1.295, { title: 'Our Project' });
  iconGlobeSearch(s, 1.219, 1.393, 0.266, C.blue, C.white);
  headline(s, [{ dark: 'About Our Suitable' }, { accent: 'Development', dot: true }], 4.896, 2.964, 3.525, 0.909);
  brand(s);
  copy(s, LOREM.f + " been the industry's standard dummy Ipsum standard. been ", 4.896, 4.262, 4.298, 0.912);
  gauge(s, 6.753, 1.019, { from: 123.95, to: 35.48, arc: C.yellow, dot: C.blue, label: '86%' });
  copy(s, 'Production 02', 6.705, 2.094, 1.094, 0.271, { align: 'center', fontSize: 7.5, bold: true, color: C.ink });
  gauge(s, 7.943, 1.019, { from: 123.95, to: 35.48, arc: C.blue, dot: C.yellow, label: '86%' });
  copy(s, 'Production 02', 7.895, 2.094, 1.094, 0.271, { align: 'center', fontSize: 7.5, bold: true, color: C.ink });
});

// 28 — biodiversity and water cathment
slide((s) => {
  photo(s, 5, 1.086, 4.123, 4.539);
  brand(s);
  text(s, [
    { text: 'Biodiversity And', options: { color: C.ink, breakLine: true } },
    { text: 'Water', options: { color: C.ink } },
    { text: ' Cathment', options: { color: C.blue } },
    { text: '.', options: { color: C.ink } },
  ], { x: 0.768, y: 0.954, w: 3.784, h: 0.909, fontFace: F.head, fontSize: SZ.title, bold: true });
  copy(s, LOREM.g, 0.813, 3.265, 3.31, 1.329);
  caption(s, 0.877, 2.503, { title: 'Our Project Management', titleW: 2.819 });
  card(s, 8.194, 2.402, 1.806, 1.905, { blur: 31, opacity: 0.24 });
  statPanel(s, 7.91, 2.658, { value: '$1670,000', valueColor: C.blue });
});

// 29 — welcome to our company
slide((s) => {
  photo(s, 3.828, 2.812, 6.172, 2.812);
  card(s, 3.684, 2.45, 2.816, 0.975, { blur: 19, opacity: 0.25 });
  caption(s, 4.524, 2.694, { title: 'Management Project', titleW: 2.026 });
  card(s, 1.43, 3.75, 2.816, 0.975, { fill: C.yellow, blur: 19, opacity: 0.25 });
  caption(s, 2.269, 3.993, { title: 'Management Project', titleW: 2.026, kickerColor: C.white, titleColor: C.white });
  card(s, 0.672, 2.459, 2.816, 0.975, { fill: C.blue, blur: 19, opacity: 0.25 });
  caption(s, 1.511, 2.703, { title: 'Management Project', titleW: 2.026, kickerColor: C.white, titleColor: C.white });
  iconGearPerson(s, 1.034, 2.784, 0.346, C.white, C.blue);
  headline(s, [{ dark: 'Welcome To Our' }, { accent: 'Company', dot: true }], 0.592, 1.016, 4.027, 0.909);
  brand(s);
  copy(s, LOREM.f, 4.981, 1.046, 4.311, 0.704);
  iconBadge(s, 4.019, 2.731, 0.378, C.blue);
  iconGlobeSearch(s, 1.696, 4.045, 0.334, C.white, C.yellow);
});

// 30 — thank you
slide((s) => {
  box(s, 0, 0, 7.903, 5.625, C.blue);
  photo(s, 0, 0.985, 7.051, 4.64);
  brand(s);
  card(s, 5, 2.596, 5, 1.01);
  text(s, [
    { text: 'THANK ', options: { color: C.ink } },
    { text: 'YOU', options: { color: C.yellow } },
    { text: '.', options: { color: C.ink } },
  ], { x: 5.256, y: 2.697, w: 4.489, h: 0.934, align: 'center', fontFace: F.thanks, fontSize: 49.5 });
  iconMenu(s, 0.299, 0.452, 0.235, 0.199, C.white);
  iconSocial(s, 6.345, 0.466, 0.145, C.white);
  text(s, 'Join Us!', { x: 5.571, y: 0.406, w: 0.984, h: 0.27, fontSize: SZ.small, color: C.white, lineSpacingMultiple: 1.2 });
  text(s, 'About Us!', { x: 2.763, y: 0.404, w: 0.984, h: 0.27, fontSize: SZ.small, color: C.white, lineSpacingMultiple: 1.2 });
  moreInfo(s, 8.855, 4.831, C.blue);
});

/* ------------------------------------------------------------------ build */

function build() {
  const pres = new PptxGenJS();
  S = pres.ShapeType;
  pres.defineLayout({ name: 'DECK', width: 10, height: 5.625 });
  pres.layout = 'DECK';
  pres.author = 'Mineral Exploration';
  pres.title = 'Mineral Exploration';
  slides.forEach((fn) => fn(pres.addSlide()));
  return pres.writeFile({ fileName: OUT });
}

build().then(() => console.log('wrote ' + OUT));
