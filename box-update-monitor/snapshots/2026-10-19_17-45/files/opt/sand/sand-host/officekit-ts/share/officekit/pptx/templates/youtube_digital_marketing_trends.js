/**
 * "Tootube Digital Marketing Trends" - 30 slide deck, 13.333 x 7.5 in.
 * Rebuilt from scratch with pptxgenjs. Every photo in the original is a
 * "REPLACE YOUR IMAGE HERE" stock placeholder, so each one is redrawn here as a
 * flat grey shape labelled [image].
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ─────────────────────────── design tokens ─────────────────────────── */

const C = {
  yellow: 'FFC000', // theme accent1
  orange: 'FF9900', // theme accent2
  red: 'FA3000', // theme accent3
  green: '00B050', // theme accent4
  black: '000000',
  white: 'FFFFFF',
  grey: '7F7F7F', // body copy (tx1 at 50% luminance)
  img: 'CCCCCC', // image placeholder body
  shell: 'E4E4E4', // device mock-up shell
  bezel: '141414' // dark tablet bezel
};

const FONT = {
  head: 'Montserrat ExtraBold',
  semi: 'Montserrat SemiBold',
  body: 'Open Sans',
  num: 'Roboto Medium'
};

const NONE = { type: 'none' }; // no line / no fill

/* Filler copy reused across the deck. */
const LOREM = {
  a: 'lorem ipsum, quia\u00a0dolor amet, consectetur, adipiscivelit, sed\u00a0quia non numquam\u00a0eius modi\u00a0tempora.',
  b: 'lorem ipsum, quia\u00a0dolor amet, consectetur, adipiscivelit, sed\u00a0quia non numquam\u00a0eius modi\u00a0tempora. lorem ipsum, quia\u00a0dolor sit, consectetur, adipiscivelit.',
  c: 'lorem ipsum, quia\u00a0dolor amet, consectetur, adipiscivelit, sed\u00a0quia non numquam\u00a0eius modi\u00a0tempora. lorem ipsum, quia\u00a0dolor amet, consectetur, adipiscivelit.',
  d: 'lorem ipsum, quia\u00a0dolor amet, consectetur, adipiscivelit, sed\u00a0quia non numquam\u00a0eius modi\u00a0tempora. lorem ipsum.',
  card: 'lorem ipsum quia\u00a0dolor sit consectetur adipiscivelit, sit quia nonnumquam eius modi\u00a0tempora.',
  icon: 'Lorem ipsum dolor sit amet, et dolore magna aliqlo rem ipsum. Ipsum dolor sit.',
  amet: '( lorem ipsum, Amet )'
};

/* ──────────────────────────── draw helpers ─────────────────────────── */

const rect = (s, x, y, w, h, color, o) =>
  s.addShape('rect', Object.assign({ x, y, w, h, fill: { color }, line: NONE }, o));

const oval = (s, x, y, d, color, o) =>
  s.addShape('ellipse', Object.assign({ x, y, w: d, h: d, fill: { color }, line: NONE }, o));

/** Fully rounded "pill" button. */
const pill = (s, x, y, w, h, color, o) =>
  s.addShape('roundRect', Object.assign(
    { x, y, w, h, fill: { color }, line: NONE, rectRadius: Math.min(w, h) / 2 }, o));

/** Grey stand-in for a photo. `shape` may be rect / ellipse / roundRect. */
function photo(s, x, y, w, h, o) {
  o = o || {};
  s.addShape(o.shape || 'rect', {
    x, y, w, h,
    fill: { color: o.color || C.img },
    line: o.line || NONE,
    rectRadius: o.rectRadius,
    rotate: o.rotate
  });
  if (w > 0.9 && h > 0.5 && o.label !== false) {
    s.addText('[image]', {
      x, y, w, h, align: 'center', valign: 'middle',
      fontFace: FONT.body, fontSize: Math.min(12, Math.max(8, w * 2.2)), color: C.white
    });
  }
}

/**
 * The deck's signature decoration: two translucent haloes around a solid dot.
 * x / y are the top-left of the outermost circle.
 */
const HALO = {
  red: { color: C.red, sizes: [0.8, 0.618, 0.445] },
  orange: { color: C.orange, sizes: [0.666, 0.515, 0.371] },
  yellow: { color: C.yellow, sizes: [0.361, 0.279, 0.201] },
  green: { color: C.green, sizes: [0.835, 0.56] }
};

function halo(s, kind, x, y) {
  const spec = HALO[kind];
  const cx = x + spec.sizes[0] / 2;
  const cy = y + spec.sizes[0] / 2;
  spec.sizes.forEach((d, i) => {
    const solid = i === spec.sizes.length - 1;
    s.addShape('ellipse', {
      x: cx - d / 2, y: cy - d / 2, w: d, h: d,
      fill: { color: spec.color, transparency: solid ? 0 : 85 },
      line: NONE
    });
  });
}

const haloes = (s, list) => list.forEach(h => halo(s, h[0], h[1], h[2]));

/**
 * Section title: mostly black with green highlights.
 * `parts` = array of [text, colorOrNull]; a trailing "\n" breaks the line.
 *
 * Each title in the source highlights two phrases in green - one filled and
 * one drawn as hollow outlined type. Both read as green here.
 */
function title(s, x, y, w, parts, o) {
  o = o || {};
  const runs = parts.map(p => {
    const brk = String(p[0]).endsWith('\n');
    return {
      text: brk ? p[0].slice(0, -1) : p[0],
      options: { color: p[1] || C.black, breakLine: brk }
    };
  });
  s.addText(runs, {
    x, y, w, h: o.h || 1.178, align: o.align || 'left', valign: 'top',
    fontFace: FONT.head, fontSize: 32, bold: true, lineSpacingMultiple: 1.0
  });
}

const G = C.green; // shorthand for the green words inside titles

/** Justified grey filler paragraph at 150% leading. */
function body(s, x, y, w, h, text, o) {
  o = o || {};
  s.addText(text, {
    x, y, w, h, valign: 'top',
    align: o.align || 'justify',
    fontFace: FONT.body,
    fontSize: o.size || 12,
    color: o.color || C.grey,
    lineSpacingMultiple: o.lead === undefined ? 1.5 : o.lead
  });
}

/** Small Montserrat SemiBold caption in an accent colour. */
function caption(s, x, y, w, h, text, color, o) {
  o = o || {};
  s.addText(text, {
    x, y, w, h, valign: 'top',
    align: o.align || 'left',
    fontFace: o.face || FONT.semi,
    fontSize: o.size || 16,
    color,
    bold: o.bold,
    italic: o.italic,
    lineSpacingMultiple: o.lead
  });
}

/** Solid colour chip with white label (used for tags and numbers). */
function chip(s, x, y, w, h, color, text, o) {
  o = o || {};
  rect(s, x, y, w, h, color, { rectRadius: o.rectRadius });
  s.addText(text, {
    x, y: o.ty === undefined ? y : o.ty, w, h: o.th || h,
    align: 'center', valign: o.th ? 'top' : 'middle',
    fontFace: o.face || FONT.semi, fontSize: o.size || 14, color: C.white, bold: o.bold
  });
}

/** Outlined capsule button (slides 2 and 4). */
function ghostButton(s, x, y, w, h, text) {
  s.addShape('roundRect', { x, y, w, h, fill: NONE, line: { color: C.red, width: 1 } });
  s.addText(text, {
    x, y: y + 0.069, w, h: 0.286, align: 'center', valign: 'top',
    fontFace: FONT.body, fontSize: 11, bold: true, color: C.yellow
  });
}

/** Filled capsule button with white label (slides 1, 7, 30). */
function solidButton(s, x, y, w, h, color, text, o) {
  o = o || {};
  pill(s, x, y, w, h, color);
  s.addText(text, {
    x, y: y + 0.073, w: o.tw || w, h: 0.286, align: 'center', valign: 'top',
    fontFace: FONT.body, fontSize: o.size || 11, bold: true,
    color: C.white, transparency: o.transparency
  });
}

/**
 * pptxgenjs has no gradient fill, so linear washes are painted as a row of
 * abutting slices whose transparency follows `alphaAt(t)`, t running 0 -> 1
 * left to right. Slices must butt exactly: any overlap shows as seams.
 */
function washBand(s, x, y, w, h, color, alphaAt) {
  const N = 80;
  for (let i = 0; i < N; i++) {
    const x0 = x + (w * i) / N;
    const x1 = x + (w * (i + 1)) / N;
    s.addShape('rect', {
      x: x0, y, w: x1 - x0, h,
      fill: { color, transparency: alphaAt((i + 0.5) / N) },
      line: NONE
    });
  }
}

/** Band fading solid -> transparent (`dir: 'right'`) or the reverse. */
const fadeBand = (s, x, y, w, h, color, dir) =>
  washBand(s, x, y, w, h, color, t => Math.round(100 * (dir === 'right' ? t : 1 - t)));

/**
 * Band that is most opaque at its two ends and lightest mid-span - the source
 * paints these with a radial "circle" gradient running 80% -> 50% alpha.
 * `cap` rounds the ends into a pill.
 */
function bulgeBand(s, x, y, w, h, color, cap) {
  const r = cap ? h / 2 : 0;
  washBand(s, x + r, y, w - 2 * r, h, color, t => Math.round(50 - 30 * Math.abs(t - 0.5) * 2));
  if (cap) {
    s.addShape('pie', { x, y, w: h, h, fill: { color, transparency: 20 }, line: NONE, angleRange: [90, 270] });
    s.addShape('pie', { x: x + w - h, y, w: h, h, fill: { color, transparency: 20 }, line: NONE, angleRange: [270, 90] });
  }
}

/** Half disc with a flat top - the little A/B/C tabs on slide 15. */
function tab(s, x, y, w, h, color, letter) {
  const flat = 0.109;
  rect(s, x, y, w, flat, color);
  s.addShape('pie', {
    x, y: y + flat - w / 2, w, h: w,
    fill: { color }, line: NONE, angleRange: [0, 180]
  });
  s.addText(letter, {
    x, y: y + 0.028, w, h: 0.404, align: 'center', valign: 'top',
    fontFace: FONT.semi, fontSize: 18, color: C.white
  });
}

/* Simple pictograms drawn from primitives (the original uses vector icons). */
function iconPerson(s, cx, cy, size, color) {
  oval(s, cx - size * 0.19, cy - size * 0.42, size * 0.38, color);
  s.addShape('chord', {
    x: cx - size * 0.36, y: cy - size * 0.06, w: size * 0.72, h: size * 0.62,
    fill: { color }, line: NONE, angleRange: [180, 0]
  });
}

function iconMail(s, cx, cy, size, color) {
  rect(s, cx - size * 0.45, cy - size * 0.3, size * 0.9, size * 0.6, color);
  s.addShape('triangle', {
    x: cx - size * 0.45, y: cy - size * 0.3, w: size * 0.9, h: size * 0.42,
    fill: { color: C.green }, line: { color, width: 0.75 }, flipV: true
  });
}

function iconHome(s, cx, cy, size, color) {
  s.addShape('triangle', { x: cx - size * 0.5, y: cy - size * 0.42, w: size, h: size * 0.45, fill: { color }, line: NONE });
  rect(s, cx - size * 0.32, cy - size * 0.02, size * 0.64, size * 0.44, color);
}

function iconGlobe(s, cx, cy, size, color) {
  s.addShape('ellipse', { x: cx - size / 2, y: cy - size / 2, w: size, h: size, fill: NONE, line: { color, width: 1.25 } });
  s.addShape('ellipse', { x: cx - size * 0.22, y: cy - size / 2, w: size * 0.44, h: size, fill: NONE, line: { color, width: 1 } });
  s.addShape('line', { x: cx - size / 2, y: cy, w: size, h: 0, line: { color, width: 1 } });
}

function iconTarget(s, cx, cy, size, color) {
  s.addShape('ellipse', { x: cx - size / 2, y: cy - size / 2, w: size, h: size, fill: NONE, line: { color, width: 1.25 } });
  s.addShape('ellipse', { x: cx - size * 0.28, y: cy - size * 0.28, w: size * 0.56, h: size * 0.56, fill: NONE, line: { color, width: 1 } });
  oval(s, cx - size * 0.1, cy - size * 0.1, size * 0.2, color);
}

/** White tick mark inscribed in a `size` box whose top-left is x / y. */
function iconCheck(s, x, y, size) {
  const ln = { color: C.white, width: 2.5 };
  s.addShape('line', { x: x + size * 0.08, y: y + size * 0.5, w: size * 0.3, h: size * 0.34, line: ln });
  s.addShape('line', { x: x + size * 0.38, y: y + size * 0.16, w: size * 0.54, h: size * 0.68, line: ln, flipV: true });
}

/* ───────────────────── master background (all slides) ──────────────── */

/**
 * Pin-stripe paper: 61 white bars, each casting a very soft shadow, plus the
 * green quarter-disc tucked into the top-right corner.
 */
function background(s) {
  const barW = 0.196;
  const gap = 0.2227;
  for (let i = 0; i < 61; i++) {
    s.addShape('rect', {
      x: i * gap, y: 0, w: barW, h: 7.5,
      fill: { color: C.white }, line: NONE,
      shadow: { type: 'outer', color: C.black, opacity: 0.05, blur: 52, offset: 4, angle: 45 }
    });
  }
  s.addShape('pie', {
    x: 12.579, y: -0.754, w: 1.508, h: 1.508,
    fill: { color: C.green }, line: NONE, angleRange: [90, 180]
  });
}

/* ───────── layout artwork (photo frames that live on the layouts) ───── */

/** Tablet mock-up: dark bezel, grey screen, home button (slide 21). */
function tabletDark(s, x, y, w, h) {
  s.addShape('roundRect', { x, y, w, h, fill: { color: C.bezel }, line: NONE, rectRadius: 0.1 });
  photo(s, 1.9, 1.32, 3.58, 4.79, { label: false });
  oval(s, 3.57, 6.36, 0.24, '3A3A3A');
  oval(s, x + w / 2 - 0.025, y + 0.16, 0.05, '3A3A3A');
}

/** Two light tablets, front and back (slide 26). */
function tabletsLight(s) {
  s.addShape('roundRect', { x: 9.99, y: 1.53, w: 1.93, h: 4.43, fill: { color: C.shell }, line: { color: 'BDBDBD', width: 1 }, rectRadius: 0.16 });
  s.addShape('roundRect', { x: 6.4, y: 1.56, w: 0.34, h: 4.47, fill: { color: 'E8E8E8' }, line: { color: 'BDBDBD', width: 1 }, rectRadius: 0.16, rotate: 3 });
  s.addShape('roundRect', { x: 7.0, y: 1.5, w: 3.14, h: 4.55, fill: { color: C.white }, line: { color: 'CFCFCF', width: 1 }, rectRadius: 0.16 });
  photo(s, 7.265, 1.933, 2.746, 3.655, { label: false });
  oval(s, 8.46, 5.68, 0.22, 'F4F4F4', { line: { color: 'C6C6C6', width: 0.75 } });
  oval(s, 8.54, 1.66, 0.06, 'C6C6C6');
}

/**
 * Gold-and-white smart watch (slide 28), tilted like the original photo.
 * `cx`/`cy` is the centre of the watch face and `f` its width in inches.
 */
function watch(s, cx, cy, f, label) {
  const spin = { rotate: -20 };
  const at = (w, h) => ({ x: cx - w / 2, y: cy - h / 2, w, h });
  s.addShape('roundRect', Object.assign(at(f * 0.58, f * 2.35), { fill: { color: 'F2F2F2' }, line: NONE, rectRadius: 0.16 }, spin));
  s.addShape('roundRect', Object.assign(at(f, f * 1.12), { fill: { color: 'D8A93C' }, line: NONE, rectRadius: 0.14 }, spin));
  s.addShape('roundRect', Object.assign(at(f * 0.9, f * 1.0), { fill: { color: C.bezel }, line: NONE, rectRadius: 0.11 }, spin));
  s.addShape('rect', Object.assign(at(f * 0.68, f * 0.74), { fill: { color: C.img }, line: NONE }, spin));
  if (label) {
    s.addText('[image]', Object.assign(at(f * 0.68, f * 0.74), {
      align: 'center', valign: 'middle', fontFace: FONT.body, fontSize: 9, color: C.white
    }, spin));
  }
}

/**
 * Head-and-shoulders cut-out on the cover, traced as a polygon: each row is
 * [y, leftEdge, rightEdge] in inches, measured off the original photo.
 */
const SILHOUETTE = [
  [0.55, 8.55, 9.50], [0.75, 8.42, 9.65], [1.13, 8.17, 9.84], [1.50, 8.01, 9.98],
  [1.88, 7.80, 10.06], [2.25, 7.80, 10.01], [2.63, 7.93, 10.03], [3.00, 8.13, 10.34],
  [3.75, 7.58, 10.94], [4.50, 7.79, 11.47], [5.25, 8.13, 11.87], [6.00, 8.38, 11.65],
  [6.75, 8.31, 10.67], [7.50, 8.06, 10.68]
];

function silhouette(s) {
  const x0 = 7.5;
  const pts = [];
  SILHOUETTE.forEach(r => pts.push({ x: r[1] - x0, y: r[0] }));
  SILHOUETTE.slice().reverse().forEach(r => pts.push({ x: r[2] - x0, y: r[0] }));
  pts.push({ close: true });
  s.addShape('custGeom', { x: x0, y: 0, w: 4.5, h: 7.5, points: pts, fill: { color: C.img }, line: NONE });
}

/**
 * Photo furniture supplied by each slide layout, keyed by slide number.
 * Drawn straight after the pin-stripe background and before slide content.
 */
const LAYOUT_ART = {
  2: s => photo(s, 6.667, 0.759, 5.254, 5.981),
  3: s => photo(s, 0.672, 0.756, 5.25, 5.987),
  4: s => photo(s, 7.425, 1.503, 4.481, 4.487),
  5: s => {
    oval(s, -1.567, -1.581, 7.653, C.red);
    photo(s, -1.737, -1.751, 7.654, 7.654, { shape: 'ellipse', label: false });
  },
  6: s => [1.422, 5.171, 8.919].forEach(x => photo(s, x, 2.617, 2.992, 1.503)),
  8: s => photo(s, 7.391, -0.494, 8.487, 8.487, { shape: 'ellipse', label: false }),
  9: s => [1.422, 5.171, 8.919].forEach(x =>
    photo(s, x, 2.55, 2.992, 1.663, { shape: 'roundRect', rectRadius: 0.155 })),
  10: s => photo(s, 7.417, 0, 5.917, 7.5),
  11: s => {
    oval(s, -0.689, -1.527, 6.599, C.yellow);
    photo(s, -0.771, -1.625, 6.599, 6.599, { shape: 'ellipse', label: false });
    photo(s, 2.61, 5.424, 3.758, 3.758, { shape: 'ellipse', label: false });
  },
  12: s => [1.415, 5.071, 8.726].forEach(x => photo(s, x, 2.443, 3.192, 3.557)),
  13: s => {
    rect(s, 8.93, 0, 4.403, 7.5, C.green);
    oval(s, 7.369, 1.447, 4.607, C.white);
    photo(s, 7.438, 1.516, 4.468, 4.468, { shape: 'ellipse' });
  },
  14: s => {
    rect(s, 1.593, 0.167, 4.496, 6.0, C.green);
    photo(s, 1.427, 0, 4.496, 6.0);
  },
  15: s => [1.435, 5.234, 9.033].forEach(x =>
    photo(s, x, 2.646, 2.858, 1.82, { shape: 'roundRect', rectRadius: 0.91 })),
  16: s => {
    rect(s, 0, 0, 3.667, 7.5, C.red);
    photo(s, 1.408, 1.511, 4.514, 4.489);
  },
  17: s => photo(s, 7.422, 0, 3.765, 6.0),
  18: s => [[1.433, C.yellow], [4.145, C.green], [6.857, C.orange], [9.569, C.red]].forEach(c => {
    oval(s, c[0], 2.443, 2.238, c[1]);
    photo(s, c[0] + 0.106, 2.495, 2.238, 2.238, { shape: 'ellipse', label: false });
  }),
  19: s => photo(s, 7.424, 0, 5.909, 7.5),
  20: s => photo(s, 1.664, 2.25, 10.004, 3.0, { shape: 'roundRect', rectRadius: 1.5 }),
  21: s => {
    rect(s, 0, 0, 3.678, 7.5, C.green);
    tabletDark(s, 1.46, 0.86, 4.44, 5.79);
  },
  22: s => [[1.432, 3.243, C.yellow], [5.084, 3.259, C.red], [8.747, 3.259, C.orange]].forEach(c => {
    rect(s, c[0], c[1], 3.154, 2.756, c[2]);
    photo(s, c[0] + 0.085, c[1] + 0.075, 2.984, 2.607);
  }),
  23: s => {
    oval(s, 8.87, -0.467, 8.617, C.red);
    photo(s, 7.422, 1.5, 4.493, 4.5);
  },
  24: s => photo(s, 0, 2.533, 13.333, 2.492),
  25: s => {
    photo(s, 0, 0, 5.904, 7.5);
    photo(s, 7.403, 1.511, 4.514, 2.656);
  },
  26: s => {
    oval(s, 8.87, -0.467, 8.617, C.orange);
    tabletsLight(s);
  },
  27: s => [[1.417, C.green], [4.157, C.yellow], [7.06, C.orange], [9.8, C.red]].forEach(c => {
    s.addShape('roundRect', { x: c[0], y: 2.444, w: 2.117, h: 2.383, fill: { color: c[1] }, line: NONE, rectRadius: 0.268 });
    photo(s, c[0] + 0.039, 2.489, 2.038, 2.294, { shape: 'roundRect', rectRadius: 0.229 });
  }),
  28: s => {
    oval(s, -0.689, -1.527, 6.599, C.red);
    watch(s, 4.43, 4.0, 2.55, true);
    watch(s, 2.26, 5.0, 1.42, false);
  },
  30: s => {
    oval(s, 2.218, 1.833, 3.754, C.yellow);
    photo(s, 2.176, 1.873, 3.754, 3.754, { shape: 'ellipse' });
  }
};

/* ───────────────────────── per-slide builders ──────────────────────── */

const SLIDES = [];

/* 1 - cover */
SLIDES.push(s => {
  s.addShape('pie', { x: 4.844, y: 4.032, w: 12.121, h: 12.121, fill: { color: C.red }, line: NONE, angleRange: [180, 0] });
  silhouette(s);
  haloes(s, [['red', 0.323, 0.378], ['orange', 3.455, 6.376], ['yellow', 1.032, 1.734], ['green', 1.926, 2.165]]);
  title(s, 2.076, 2.237, 4.726, [['Tootube ', null], ['Digital\n', G], ['Marketing', G], [' Trends', null]]);
  body(s, 2.076, 3.75, 4.7, 0.675, LOREM.a);
  solidButton(s, 2.162, 4.964, 2.051, 0.448, C.green, 'Sign in . . .', { size: 14, tw: 1.546, transparency: 40 });
  s.addShape('line', { x: 3.729, y: 5.037, w: 0, h: 0.303, line: { color: C.white, width: 1, transparency: 50 } });
  iconPerson(s, 3.967, 5.188, 0.226, C.white);
});

/* 2 - welcome */
SLIDES.push(s => {
  fadeBand(s, 6.667, 4.506, 6.667, 1.485, C.yellow, 'right');
  haloes(s, [['orange', 5.196, 6.407], ['red', 6.267, 0.411], ['green', 1.16, 1.328], ['yellow', 0.43, 0.398]]);
  title(s, 1.31, 1.4, 4.294, [['Welcome ', null], ['To\n', G], ['Our', G], [' Presentation', null]]);
  body(s, 1.31, 3.022, 4.7, 0.675, LOREM.a);
  body(s, 1.31, 3.987, 4.7, 0.978, LOREM.b);
  ghostButton(s, 1.427, 5.56, 1.203, 0.43, 'More Info');
  ghostButton(s, 3.188, 5.56, 1.203, 0.43, 'Learn More');
  s.addText('\u201c Content Marketing Is The Key For Success \u201d', {
    x: 6.912, y: 4.844, w: 4.763, h: 0.808, align: 'center', valign: 'top',
    fontFace: FONT.head, fontSize: 21, bold: true, color: C.white
  });
});

/* 3 - summary message */
SLIDES.push(s => {
  fadeBand(s, 0.672, 3.972, 11.242, 1.037, C.red, 'left');
  haloes(s, [['green', 7.171, 1.328], ['red', 0.357, 0.357], ['yellow', 6.472, 2.729], ['orange', 12.038, 6.41]]);
  title(s, 7.322, 1.4, 4.864, [['Our ', null], ['Tootuber\n', G], ['Summary', G], [' Message', null]]);
  body(s, 7.322, 2.936, 4.7, 0.675, LOREM.a);
  body(s, 7.322, 5.371, 4.7, 0.675, LOREM.a);
  s.addText('\u201c Makes Your Marketing So Useful People Would Pay For It \u201d', {
    x: 7.411, y: 4.192, w: 4.503, h: 0.572, align: 'center', valign: 'top',
    fontFace: FONT.head, fontSize: 14, color: C.white
  });
});

/* 4 - about digital marketing */
SLIDES.push(s => {
  haloes(s, [['green', 1.16, 1.328], ['red', 7.022, 4.403], ['orange', 7.729, 5.223], ['yellow', 7.241, 5.81]]);
  title(s, 1.31, 1.4, 4.598, [['About ', null], ['Tootube\n', G], ['Digital', G], [' Marketing', null]]);
  body(s, 1.31, 3.091, 4.7, 0.675, 'lorem ipsum, quia\u00a0dolor amet, consectetur, adipiscivelit, sed\u00a0quia non numquam\u00a0eius tempora.');
  body(s, 1.31, 5.068, 4.7, 0.978, LOREM.b);
  ghostButton(s, 1.427, 4.202, 1.203, 0.43, 'Streaming');
  ghostButton(s, 3.167, 4.202, 1.203, 0.43, 'Live Videos');
});

/* 5 - best chanel vision */
SLIDES.push(s => {
  haloes(s, [['red', 5.612, 4.939], ['orange', 4.607, 5.996], ['yellow', 3.619, 6.549], ['green', 7.171, 1.328]]);
  title(s, 7.322, 1.4, 4.89, [['About ', null], ['Tootube\n', G], ['Best Chanel', G], [' Vision', null]]);
  oval(s, 7.42, 3.003, 0.609, C.orange);
  iconGlobe(s, 7.725, 3.308, 0.29, C.white);
  body(s, 8.3, 2.955, 3.722, 0.627, LOREM.icon, { size: 11 });
  oval(s, 7.42, 4.083, 0.609, C.yellow);
  iconTarget(s, 7.725, 4.388, 0.29, C.white);
  body(s, 8.3, 4.035, 3.722, 0.627, LOREM.icon, { size: 11 });
  body(s, 7.322, 5.059, 4.7, 0.978, LOREM.b);
});

/* 6 - three-card mission */
SLIDES.push(s => {
  haloes(s, [['red', -0.4, 0.436], ['orange', 12.226, 6.451], ['green', 6.249, -0.427], ['yellow', 0.847, 1.72]]);
  title(s, 2.789, 1.4, 7.755, [['Our ', null], ['Tootube ', G], ['Chanel', G], [' Mission', null]], { h: 0.64, align: 'center' });
  const cards = [
    { x: 1.422, color: C.red, num: '01', tag: '( Be The Best )', tx: 1.876, bx: 1.208 },
    { x: 5.171, color: C.yellow, num: '02', tag: '( Daily Upload )', tx: 5.635, bx: 4.967 },
    { x: 8.919, color: C.orange, num: '03', tag: '( Consistent )', tx: 9.393, bx: 8.725 }
  ];
  cards.forEach(c => {
    chip(s, c.x, 2.97, 0.747, 0.75, c.color, c.num, { ty: 3.092, th: 0.505, size: 24 });
    caption(s, c.tx, 4.631, 2.064, 0.337, c.tag, c.color, { align: 'center' });
    body(s, c.bx, 5.091, 3.4, 0.978, LOREM.card, { align: 'center' });
  });
});

/* 7 - gaming chanels analystic */
SLIDES.push(s => {
  oval(s, 1.34, 1.419, 4.654, C.white);
  photo(s, 1.42, 1.503, 4.493, 4.493, { shape: 'ellipse' });
  haloes(s, [['red', 5.05, 4.836], ['orange', 1.269, 6.066], ['yellow', 4.054, 5.816], ['green', 6.423, 1.328]]);
  title(s, 6.573, 1.4, 4.89, [['About ', null], ['Gaming\n', G], ['Chanels ', G], ['Analystic', null]]);
  caption(s, 6.561, 2.999, 1.889, 0.337, 'Future Vision', C.yellow);
  body(s, 6.573, 3.524, 5.45, 0.675, LOREM.d);
  solidButton(s, 6.667, 4.555, 1.239, 0.448, C.yellow, 'More Info');
  solidButton(s, 8.281, 4.555, 1.239, 0.448, C.yellow, 'Learn More');
  body(s, 6.573, 5.36, 5.45, 0.675, LOREM.d);
});

/* 8 - media supports (bulleted) */
SLIDES.push(s => {
  haloes(s, [['yellow', 7.359, 4.651], ['red', 8.178, 6.202], ['orange', 7.042, 5.712], ['green', 1.16, 1.328]]);
  title(s, 1.31, 1.4, 4.294, [['Our ', null], ['Tootube\n', G], ['Media', G], [' Supports', null]]);
  oval(s, 1.418, 3.166, 0.175, C.red);
  body(s, 1.84, 3.027, 4.17, 0.627, 'lorem ipsum, quia\u00a0dolor sit, amet, consectetur, adipiscivelit, sed\u00a0quia non numquam\u00a0eius.', { size: 11 });
  oval(s, 1.418, 4.164, 0.175, C.yellow);
  body(s, 1.84, 4.025, 4.17, 0.627, 'lorem ipsum, quia\u00a0dolor sit, amet, consectetur, adipiscivelit, non numquam\u00a0eius.', { size: 11 });
  body(s, 1.31, 5.1, 4.7, 0.978, LOREM.c);
});

/* 9 - chanel types (three tags) */
SLIDES.push(s => {
  haloes(s, [['green', 6.249, -0.427], ['yellow', 1.061, 1.763], ['red', 11.511, 2.997], ['orange', 4.862, 3.797]]);
  title(s, 2.789, 1.4, 7.755, [['About ', null], ['Tootube ', G], ['Chanel ', G], ['Type', null]], { h: 0.64, align: 'center' });
  const cols = [
    { tag: 'Gaming', tx: 2.37, tw: 1.077, ty: 4.87, color: C.green, bx: 1.295 },
    { tag: 'Tutorial', tx: 6.1, tw: 1.133, ty: 4.87, color: C.red, bx: 5.054 },
    { tag: 'Vloger', tx: 9.886, tw: 1.077, ty: 4.832, color: C.yellow, bx: 8.812 }
  ];
  cols.forEach(c => {
    chip(s, c.tx, c.ty, c.tw, 0.337, c.color, c.tag);
    body(s, c.bx, 5.368, 3.23, 0.675,
      'lorem ipsum quia\u00a0dolor amet eius sit consectetur adipiscivelit.', { align: 'center' });
  });
});

/* 10 - most wanted chanel */
SLIDES.push(s => {
  haloes(s, [['orange', 0.28, 0.316], ['red', 7.071, 4.931], ['yellow', 6.8, 6.254], ['green', 1.16, 1.328]]);
  title(s, 1.31, 1.4, 4.294, [['Most ', null], ['Wanted\n', G], ['Chanel', G], [' This Year', null]]);
  body(s, 1.31, 3.06, 4.7, 0.675, LOREM.a);
  s.addText([
    { text: 'You Can\u2019t ', options: { color: C.black } },
    { text: 'Sell Anything', options: { color: C.yellow } },
    { text: ' If ', options: { color: C.black } },
    { text: 'You Can\u2019t ', options: { color: C.yellow } },
    { text: 'Tell Anything', options: { color: C.black } }
  ], { x: 1.308, y: 4.16, w: 4.79, h: 0.808, valign: 'top', fontFace: FONT.head, fontSize: 21 });
  body(s, 1.31, 5.393, 4.7, 0.675, LOREM.a);
});

/* 11 - musics premium */
SLIDES.push(s => {
  haloes(s, [['yellow', 12.525, 4.568], ['orange', 9.338, 7.167], ['red', 5.568, 4.248], ['green', 7.171, 1.328]]);
  title(s, 7.322, 1.4, 4.294, [['About ', null], ['Tootube\n', G], ['Musics', G], [' Premium', null]]);
  body(s, 7.322, 2.956, 4.7, 0.675, LOREM.a);
  s.addShape('roundRect', { x: 7.447, y: 3.979, w: 3.439, h: 0.805, fill: { color: C.white }, line: { color: C.orange, width: 1 }, rectRadius: 0.09 });
  caption(s, 7.447, 4.113, 3.439, 0.539, 'The Best Tootube Is How The Marketing Managed', C.red, { align: 'center', size: 13 });
  body(s, 7.322, 5.059, 4.7, 0.978, LOREM.b);
});

/* 12 - photo galery */
SLIDES.push(s => {
  haloes(s, [['red', 10.055, 6.404], ['orange', 1.132, 2.11], ['yellow', 12.201, 4.721], ['green', 6.249, -0.427]]);
  title(s, 2.789, 1.4, 7.755, [['Tootube ', null], ['Chanel ', G], ['Photo', G], [' Galery', null]], { h: 0.64, align: 'center' });
  [[1.415, C.yellow, 'Gaming Chanel'], [5.071, C.red, 'Filming Chanel'], [8.726, C.orange, 'Vloging Chanel']]
    .forEach(c => chip(s, c[0], 4.86, 2.257, 0.666, c[1], c[2], { ty: 5.002, th: 0.337 }));
});

/* 13 - best media analiystic */
SLIDES.push(s => {
  haloes(s, [['red', 7.824, 4.984], ['orange', 7.306, 4.248], ['yellow', 0.581, 0.595], ['green', 1.16, 1.328]]);
  title(s, 1.31, 1.4, 4.294, [['Best ', null], ['Tootube\n', G], ['Media', G], [' analiystic', null]]);
  body(s, 1.31, 2.954, 4.7, 0.675, LOREM.a);
  chip(s, 1.427, 3.961, 1.203, 0.725, C.red, '462+', { ty: 4.06, th: 0.505, size: 24 });
  s.addText([
    { text: 'About ', options: { color: C.black } },
    { text: 'Digital', options: { color: C.red, breakLine: true } },
    { text: 'Marketing', options: { color: C.black } }
  ], { x: 2.842, y: 4.004, w: 2.483, h: 0.64, valign: 'top', fontFace: FONT.head, fontSize: 16, italic: true, charSpacing: 3 });
  body(s, 1.31, 5.071, 4.7, 0.978, LOREM.b);
});

/* 14 - chanel media analiystic */
SLIDES.push(s => {
  haloes(s, [['red', 9.672, 6.917], ['orange', 11.355, 6.583], ['yellow', 12.541, 5.805], ['green', 7.171, 1.328]]);
  title(s, 7.322, 1.4, 4.294, [['Tootube ', null], ['Chanel\n', G], ['Media', G], [' analiystic', null]]);
  chip(s, 7.411, 3.087, 1.889, 0.337, C.orange, 'The Logarithm');
  body(s, 7.322, 3.655, 4.7, 0.675, LOREM.a);
  chip(s, 7.411, 4.805, 1.889, 0.337, C.yellow, 'Support System');
  body(s, 7.322, 5.374, 4.7, 0.675, LOREM.a);
});

/* 15 - chanel types A / B / C */
SLIDES.push(s => {
  haloes(s, [['green', 6.249, -0.427], ['yellow', 1.616, 0.668], ['red', -0.4, 0.321], ['orange', 0.547, 1.272]]);
  title(s, 2.789, 1.4, 7.755, [['Tootube ', null], ['Chanel ', G], ['Types', null]], { h: 0.64, align: 'center' });
  const cols = [
    { tab: 2.465, letter: 'A', color: C.yellow, label: 'Black Types', lx: 1.796, bx: 1.213 },
    { tab: 6.26, letter: 'B', color: C.orange, label: 'Yellow Types', lx: 5.595, bx: 5.012 },
    { tab: 10.055, letter: 'C', color: C.red, label: 'Red Types', lx: 9.394, bx: 8.811 }
  ];
  cols.forEach(c => {
    tab(s, c.tab, 2.646, 0.813, 0.515, c.color, c.letter);
    caption(s, c.lx, 4.9, 2.142, 0.459, c.label, c.color, { align: 'center' });
    body(s, c.bx, 5.402, 3.31, 0.627,
      'lorem ipsum, quia amet, sit consectetur, sed\u00a0quia non numquam.', { size: 11, align: 'center' });
  });
});

/* 16 - videos premium */
SLIDES.push(s => {
  haloes(s, [['yellow', 12.525, 4.568], ['orange', 11.616, 6.511], ['red', 5.417, 5.551], ['green', 7.171, 1.328]]);
  title(s, 7.322, 1.4, 4.294, [['About ', null], ['Tootube\n', G], ['Videos', G], [' Premium', null]]);
  caption(s, 7.311, 2.974, 2.489, 0.37, 'Subscribe Soon', C.red);
  body(s, 7.322, 3.481, 4.7, 0.675, LOREM.a);
  caption(s, 7.311, 4.552, 2.81, 0.37, 'Subscribe Option', C.yellow);
  body(s, 7.322, 5.059, 4.7, 0.978, LOREM.b);
});

/* 17 - highest paid chanel */
SLIDES.push(s => {
  haloes(s, [['yellow', 6.843, 4.933], ['red', 8.279, 6.185], ['orange', 7.102, 5.667], ['green', 1.16, 1.328]]);
  title(s, 1.31, 1.4, 4.294, [['Highest ', null], ['Paid\n', G], ['Tootube', G], [' Chanel', null]]);
  body(s, 1.31, 3.057, 4.7, 0.675, LOREM.a);
  chip(s, 1.421, 4.217, 1.279, 0.733, C.red, '$17', { ty: 4.298, th: 0.572, size: 28, bold: true });
  chip(s, 3.358, 4.217, 1.279, 0.733, C.orange, '$61', { ty: 4.298, th: 0.572, size: 28, bold: true });
  body(s, 1.31, 5.36, 4.7, 0.675, LOREM.a);
  s.addShape('roundRect', { x: 8.92, y: 3.762, w: 2.995, h: 1.49, fill: { color: C.yellow }, line: NONE, rectRadius: 0.204 });
  caption(s, 9.155, 3.934, 1.872, 0.572, '1201++', C.white, { size: 28 });
  body(s, 9.176, 4.403, 2.483, 0.579,
    'lorem ipsum, quia\u00a0dolor sit, amet, consectetur, adipiscivelit, ', { size: 10, color: C.white });
});

/* 18 - income types (four price circles) */
SLIDES.push(s => {
  haloes(s, [['green', 6.249, -0.427], ['yellow', 0.741, 0.666], ['red', -0.4, 1.32], ['orange', 12.34, 6.516]]);
  title(s, 2.789, 1.4, 7.755, [['Tootube ', null], ['Income ', G], ['Types', G]], { h: 0.64, align: 'center' });
  const tiers = [
    { price: '$10 - $15', color: C.yellow, lx: 1.933, bx: 1.296, by: 5.675 },
    { price: '$15 - $20', color: C.green, lx: 4.648, bx: 4.01, by: 5.452 },
    { price: '$20 \u2013 30$', color: C.orange, lx: 7.362, bx: 6.724, by: 5.452 },
    { price: '$30 \u2013 $50', color: C.red, lx: 10.076, bx: 9.439, by: 5.452 }
  ];
  tiers.forEach(t => {
    caption(s, t.lx, 5.026, 1.34, 0.46, t.price, t.color, { align: 'center', face: FONT.num });
    body(s, t.bx, t.by, 2.616, 0.627,
      'lorem ipsum, sit consectetur sed Lorem ipsum dolor.', { size: 11, align: 'center' });
  });
});

/* 19 - logarithm pattern */
SLIDES.push(s => {
  fadeBand(s, 0, 3.964, 13.323, 1.204, C.yellow, 'right');
  haloes(s, [['green', 1.16, 1.328], ['red', 4.59, 7.102], ['yellow', 6.539, 1.989], ['orange', 5.576, 0.616]]);
  title(s, 1.31, 1.4, 4.7, [['Tootube ', null], ['Chanel\n', G], ['Logarithm', G], [' Pattern', null]]);
  body(s, 1.316, 2.949, 4.7, 0.675, LOREM.a);
  caption(s, 1.314, 4.178, 4.702, 0.774, 'The Best Marketing Doesn\u2019t Feel Like Marketing', C.white,
    { align: 'center', size: 20 });
  body(s, 1.316, 5.365, 4.7, 0.675,
    'lorem ipsum, quia\u00a0dolor amet, consectetur, adipiscivelit, sed\u00a0quia non numquam\u00a0eius modi.');
});

/* 20 - coffee break */
SLIDES.push(s => {
  bulgeBand(s, 1.664, 2.25, 10.004, 3.0, C.yellow, true);
  haloes(s, [['green', 6.249, -0.427], ['yellow', 10.333, 6.551], ['red', 1.054, 0.949], ['orange', 11.539, 5.567]]);
  s.addText('- Coffee Break Here -', {
    x: 4.055, y: 3.04, w: 5.224, h: 0.64, align: 'center', valign: 'top',
    fontFace: FONT.head, fontSize: 32, bold: true, color: C.white
  });
  body(s, 4.317, 3.786, 4.7, 0.675,
    'lorem ipsum, quia\u00a0dolor amet, consectetur, adipiscivelit, sed\u00a0quia non numquam\u00a0eius.',
    { align: 'center', color: C.white });
});

/* 21 - trending videos */
SLIDES.push(s => {
  haloes(s, [['green', 7.171, 1.328], ['yellow', 12.525, 4.568], ['orange', 8.895, 7.141], ['red', 5.574, 3.356]]);
  title(s, 7.322, 1.4, 4.7, [['About ', null], ['Tootube\n', G], ['Trending', G], [' Videos', null]]);
  caption(s, 7.311, 2.974, 2.489, 0.37, 'A. Top Trending', C.orange);
  body(s, 7.322, 3.481, 4.7, 0.675, LOREM.a);
  caption(s, 7.311, 4.552, 2.81, 0.37, 'B. Most Wacthed', C.yellow);
  body(s, 7.322, 5.059, 4.7, 0.978, LOREM.b);
});

/* 22 - team */
SLIDES.push(s => {
  haloes(s, [['green', 6.249, -0.427], ['yellow', 0.741, 0.666], ['red', -0.4, 1.32], ['orange', 12.34, 6.516]]);
  title(s, 2.789, 1.4, 7.755, [['Our ', null], ['Tootube ', G], ['Chanel', G], [' Teams', null]], { h: 0.64, align: 'center' });
  [['Mr. Brandy', 1.504, 3.176, C.yellow], ['Mrs. Claura', 5.094, 3.146, C.red], ['Mr. Kandya', 9.06, 2.527, C.orange]]
    .forEach(m => caption(s, m[1], 2.464, m[2], 0.37, m[0], m[3], { align: 'center' }));
});

/* 23 - chanel trends */
SLIDES.push(s => {
  haloes(s, [['yellow', 7.023, 0.626], ['orange', 6.663, 1.429], ['red', 7.059, 2.391], ['green', 1.16, 1.328]]);
  title(s, 1.31, 1.4, 4.294, [['Tootube ', null], ['Chanel\n', G], ['Trends', G], [' This Year', null]]);
  chip(s, 1.418, 3.191, 1.224, 0.337, C.yellow, 'Gaming');
  chip(s, 3.212, 3.191, 1.224, 0.337, C.orange, 'Vlogging');
  oval(s, 1.418, 4.172, 0.175, C.green);
  body(s, 1.84, 4.033, 4.17, 0.627,
    'lorem ipsum, quia\u00a0dolor sit, amet, consectetur, adipiscivelit, sed\u00a0quia non numquam\u00a0eius.', { size: 11 });
  body(s, 1.31, 5.09, 4.7, 0.978, LOREM.c);
});

/* 24 - top trending banner */
SLIDES.push(s => {
  bulgeBand(s, 0, 2.533, 13.333, 2.492, C.yellow, false);
  haloes(s, [['green', 6.249, -0.427], ['yellow', 0.741, 0.408], ['red', -0.4, 1.063], ['orange', 12.175, 6.25]]);
  title(s, 2.789, 1.4, 7.755, [['Tootube ', null], ['Chanel', G], [' Trend', null]], { h: 0.64, align: 'center' });
  s.addText([
    { text: '2021', options: { breakLine: true } },
    { text: 'TOP TRENDING' }
  ], {
    x: 4.591, y: 3.258, w: 4.152, h: 1.043, align: 'center', valign: 'top',
    fontFace: FONT.semi, fontSize: 28, color: C.white
  });
  body(s, 2.257, 5.439, 8.819, 0.675,
    'lorem ipsum, quia\u00a0dolor sit, amet, consectetur, adipiscivelit, sed\u00a0quia non numquam\u00a0eius modi\u00a0tempora. lorem ipsum, quia\u00a0dolor sit, amet, consectetur, adipiscivelit. Lorem ipsum.',
    { align: 'center' });
});

/* 25 - four square icons */
SLIDES.push(s => {
  haloes(s, [['yellow', 7.828, 0.537], ['red', 12.149, 6.341], ['orange', 8.693, 7.141]]);
  const cells = [
    { x: 7.429, y: 4.457, color: C.yellow, mark: 'check', label: 'Sponsorship', lc: C.yellow, tx: 8.189, tw: 1.865 },
    { x: 9.803, y: 4.457, color: C.orange, mark: 'x', label: 'Film Maker', lc: C.orange, tx: 10.563, tw: 1.475 },
    { x: 7.429, y: 5.368, color: C.red, mark: 'x', label: 'Tootubers', lc: C.red, tx: 8.189, tw: 1.865 },
    { x: 9.803, y: 5.368, color: C.green, mark: 'check', label: 'Analysis', lc: C.green, tx: 10.563, tw: 1.475 }
  ];
  cells.forEach(c => {
    rect(s, c.x, c.y, 0.649, 0.62, c.color);
    if (c.mark === 'check') iconCheck(s, c.x + 0.14, c.y + 0.13, 0.37);
    else s.addShape('mathMultiply', { x: c.x + 0.17, y: c.y + 0.16, w: 0.3, h: 0.3, fill: { color: C.white }, line: NONE });
    caption(s, c.tx, c.y + 0.037, 1.259, 0.286, c.label, c.lc, { size: 11 });
    body(s, c.tx, c.y + 0.323, c.tw, 0.252, LOREM.amet, { size: 9, lead: 1 });
  });
});

/* 26 - trending chance analysis */
SLIDES.push(s => {
  haloes(s, [['yellow', 7.816, 0.656], ['orange', 0.38, 0.352], ['red', 5.356, 7.1], ['green', 1.16, 1.328]]);
  title(s, 1.31, 1.4, 4.694, [['Trending ', null], ['Chanel\n', G], ['Chances', G], [' Analysis', null]]);
  caption(s, 1.299, 3.013, 2.967, 0.37, 'Trending Chance', C.red);
  body(s, 1.299, 3.445, 4.705, 0.675, LOREM.a);
  chip(s, 1.422, 4.554, 0.705, 0.389, C.yellow, '50%', { ty: 4.564, th: 0.37, size: 16, face: FONT.head });
  chip(s, 2.942, 4.554, 0.705, 0.389, C.yellow, '90%', { ty: 4.564, th: 0.37, size: 16, face: FONT.head });
  rect(s, 2.414, 4.705, 0.241, 0.056, C.yellow);
  body(s, 1.299, 5.437, 4.705, 0.675, LOREM.a);
});

/* 27 - best chanel types (four numbered cards) */
SLIDES.push(s => {
  haloes(s, [['green', 6.249, -0.427], ['yellow', 0.741, 0.666], ['red', -0.4, 1.32], ['orange', 12.34, 6.516]]);
  title(s, 2.789, 1.4, 7.755, [['Our ', null], ['Best ', G], ['Chanel', G], [' Types', null]], { h: 0.64, align: 'center' });
  const cards = [
    { x: 1.443, color: C.green, num: '01', label: 'Daily Vlog', tx: 2.203, tw: 1.865 },
    { x: 4.157, color: C.yellow, num: '02', label: 'Game Online', tx: 4.917, tw: 1.475 },
    { x: 7.084, color: C.orange, num: '03', label: 'Film Maker', tx: 7.849, tw: 1.865 },
    { x: 9.798, color: C.red, num: '04', label: 'Tutorials', tx: 10.563, tw: 1.475 }
  ];
  cards.forEach(c => {
    chip(s, c.x, 5.368, 0.649, 0.62, c.color, c.num, { ty: 5.457, th: 0.438, size: 20, rectRadius: 0.05 });
    caption(s, c.tx, 5.405, 1.259, 0.286, c.label, c.color, { size: 11 });
    body(s, c.tx, 5.691, c.tw, 0.252, LOREM.amet, { size: 9, lead: 1 });
  });
});

/* 28 - media connection */
SLIDES.push(s => {
  haloes(s, [['yellow', 12.469, 6.534], ['orange', 9.338, 7.167], ['red', 5.62, 6.243], ['green', 7.171, 1.328]]);
  title(s, 7.322, 1.4, 4.59, [['Behind ', null], ['Tootube\n', G], ['Media', G], [' Connection', null]]);
  [[7.43, C.yellow, '01', 'Photos', 7.938, 7.331], [9.833, C.orange, '02', 'Videos', 10.34, 9.734]].forEach(c => {
    oval(s, c[0], 3.248, 0.436, c[1]);
    s.addText(c[2], {
      x: c[0], y: 3.315, w: 0.436, h: 0.303, align: 'center', valign: 'top',
      fontFace: FONT.semi, fontSize: 12, color: C.white
    });
    caption(s, c[4], 3.281, 1.274, 0.37, c[3], c[1]);
    body(s, c[5], 3.82, 2.284, 0.627,
      'lorem ipsum, quia\u00a0dolor sit, amet, adipiscivelit.', { size: 11 });
  });
  body(s, 7.322, 5.072, 4.7, 0.978, LOREM.c);
});

/* 29 - contact */
SLIDES.push(s => {
  haloes(s, [['yellow', 11.087, 6.813], ['red', 1.054, 0.949], ['orange', 12.336, 6.147], ['green', 6.249, -0.427]]);
  title(s, 3.522, 1.4, 6.289, [['Our ', null], ['Contact ', G], ['Chanel', G], [' Here', null]], { h: 0.64, align: 'center' });
  caption(s, 5.707, 2.803, 2.082, 0.303, '( Find Issues )', C.orange, { align: 'center', size: 12 });
  body(s, 1.178, 3.254, 10.978, 0.627,
    'lorem ipsum, quia\u00a0dolor sit, amet, consectetur, adipiscivelit, sed\u00a0quia non numquam\u00a0eius modi\u00a0tempora. lorem ipsum, quia\u00a0dolor amet, consectetur adipiscivelit. lorem ipsum, quia\u00a0dolor sit. lorem ipsum, quia\u00a0dolor sit, amet, consectetur.',
    { size: 11, align: 'center' });
  const cols = [
    { cx: 2.961, color: C.yellow, icon: iconPerson, label: 'Social Media', lx: 2.088, lw: 1.746, sub: '@BestChanel', sx: 2.246, sw: 1.43 },
    { cx: 6.748, color: C.red, icon: iconMail, label: 'Email', lx: 6.155, lw: 1.185, sub: 'Bestchanel@trang.com', sx: 5.707, sw: 2.082 },
    { cx: 10.453, color: C.green, icon: iconHome, label: 'Location', lx: 9.82, lw: 1.267, sub: 'Jalious- Hanhuicho', sx: 9.662, sw: 1.584 }
  ];
  cols.forEach(c => {
    oval(s, c.cx - 0.2315, 4.689, 0.463, c.color);
    c.icon(s, c.cx, 4.92, 0.2, C.white);
    caption(s, c.lx, 5.435, c.lw, 0.37, c.label, c.color, { align: 'center' });
    body(s, c.sx, 5.776, c.sw, 0.286, c.sub, { size: 11, color: c.color, align: 'center', lead: 1 });
  });
});

/* 30 - thank you */
SLIDES.push(s => {
  haloes(s, [['green', 6.387, 2.067], ['yellow', 3.787, 5.446], ['red', 1.795, 3.156], ['orange', 2.261, 4.671],
    ['yellow', 0.758, 6.155], ['red', -0.4, 1.32], ['orange', 12.34, 6.516]]);
  title(s, 6.537, 2.139, 4.726, [['Thank ', null], ['You\n', G], ['SO', G], [' Much !', null]]);
  body(s, 6.563, 3.618, 4.7, 0.675, LOREM.a);
  solidButton(s, 6.669, 4.79, 1.239, 0.448, C.red, 'Finished');
});

/* ─────────────────────────────── build ─────────────────────────────── */

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'TOOTUBE', width: 13.333, height: 7.5 });
pptx.layout = 'TOOTUBE';
pptx.theme = { headFontFace: FONT.head, bodyFontFace: FONT.body };
pptx.title = 'Tootube Digital Marketing Trends';

SLIDES.forEach((build, i) => {
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  background(slide);
  const art = LAYOUT_ART[i + 1];
  if (art) art(slide);
  build(slide);
});

pptx.writeFile({
  fileName: path.join(__dirname, '0bb6bd07-c28b-4cf6-b405-8a720ac292ea_grok_final.pptx')
}).then(f => console.log('wrote', f));
