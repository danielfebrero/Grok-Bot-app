/**
 * "Financial Planner" deck - recreated with pptxgenjs.
 *
 * Run:  node 0b218cca-34ae-4a3c-b3cd-34f8fba993aa_grok_final.js
 * Out:  0b218cca-34ae-4a3c-b3cd-34f8fba993aa_grok_final.pptx (next to this file)
 *
 * Raster artwork in the source deck (device mock-ups, store badges, world maps)
 * is replaced by flat native-shape placeholders; the small vector glyph icons
 * are redrawn from primitives (see ICONS below).
 */

'use strict';

const pptxgen = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ theme */

const GREEN = '39B951';
const LIME = '92D050';
const DARK = '404040';
const GRAY = '595959';
const MUTED = 'A6A6A6';
const SILVER = 'BFBFBF';
const SHADE = 'F2F2F2';
const BAND = 'D9D9D9';
const WHITE = 'FFFFFF';
const BLACK = '000000';
const SCREEN = 'E8ECEB';

const F_SEMI = 'Saira SemiBold';
const F_MED = 'Saira Medium';
const F_LIGHT = 'Saira Light';
const F_BODY = 'Karla';

const SLIDE_W = 13.3333333;
const SLIDE_H = 7.5;

/** the deck's single drop-shadow recipe (blurRad 8pt / dist 3pt / 45deg / 35%);
 *  a factory because pptxgenjs rewrites the object it is handed */
function shadow() {
  return { type: 'outer', color: BLACK, opacity: 0.35, blur: 8, offset: 3, angle: 45 };
}

const LOREM_LONG =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ' +
  'ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ' +
  'ullamco laboris nisi ut aliquip ex ea commodo consequat.';
const LOREM_LONG2 =
  'Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do eiusmod tempor incididunt ' +
  'ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ' +
  'ullamco laboris nisi ut aliquip ex ea commodo consequat.';
const LOREM_MED =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ' +
  'ut labore et dolore magna aliqua.';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet consectetur adipiscing elit sed do euismod.';

/* ---------------------------------------------------------------- helpers */

/** filled shape, no outline unless asked for */
function shape(s, kind, o) {
  s.addShape(kind, Object.assign({ line: { type: 'none' } }, o));
}

/** solid rectangle */
function box(s, x, y, w, h, color, extra) {
  shape(s, 'rect', Object.assign({ x, y, w, h, fill: { color } }, extra));
}

/** text box - top anchored / Karla, like every text frame in the source deck */
function text(s, str, o) {
  s.addText(str, Object.assign({ valign: 'top', fontFace: F_BODY }, o));
}

/** 11pt body copy on 200% leading */
function body(s, str, o) {
  text(s, str, Object.assign({ fontSize: 11, color: GRAY, lineSpacingMultiple: 2 }, o));
}

/** 18pt "Saira Medium" sub-heading on 150% leading */
function subhead(s, str, o) {
  text(s, str, Object.assign({ fontSize: 18, fontFace: F_MED, color: DARK, lineSpacingMultiple: 1.5 }, o));
}

/** the recurring "Financial Plan" eyebrow + 40pt Saira SemiBold headline */
function heading(s, o) {
  text(s, 'Financial Plan', {
    x: o.ex, y: o.ey, w: 1.482, h: 0.303, wrap: false,
    fontSize: 12, charSpacing: 1, fontFace: F_LIGHT,
    color: o.eyeColor || GRAY, align: o.align,
  });
  text(s, o.title, {
    x: o.x, y: o.y, w: o.w, h: o.h || 0.774,
    fontSize: 40, fontFace: F_SEMI, color: o.color || DARK, align: o.align,
  });
}

/** bottom-left copyright line, present on every slide but 31/32 */
function footer(s, color) {
  text(s, '\u00A9 2019. All rights Reserved.',
    { x: 0.505, y: 6.727, w: 2.58, h: 0.288, fontSize: 11, color: color || MUTED });
}

/** big "82%" style stat: white triangle bullet + number */
function stat(s, x, y, value, color) {
  shape(s, 'triangle', { x, y: y + 0.442, w: 0.261, h: 0.357, fill: { color: color || WHITE } });
  text(s, value, { x: x + 0.309, y, w: 2.176, h: 1.003, fontSize: 40, color: color || WHITE, lineSpacingMultiple: 1.5 });
}

/**
 * Vertical-strip stand-in for the deck's "white -> green" gradient band.
 * The source is a linear gradFill whose white stop is fully transparent, so on
 * the white page the visible ramp is quadratic: white + f^2 * (green - white),
 * reaching full colour 65% of the way across.
 */
function gradientBand(s, x, y, w, h, color, steps) {
  const n = steps || 100;
  const rgb = [parseInt(color.slice(0, 2), 16), parseInt(color.slice(2, 4), 16), parseInt(color.slice(4, 6), 16)];
  for (let i = 0; i < n; i++) {
    const f = Math.min(1, ((i + 0.5) / n) / 0.65);
    const hex = rgb.map(function (c) {
      return ('0' + Math.round(255 + (c - 255) * f * f).toString(16)).slice(-2);
    }).join('').toUpperCase();
    box(s, x + (i * w) / n, y, w / n + 0.02, h, hex);
  }
}

/** flat stand-in for a device bitmap: rounded body + inset "screen" */
function devicePlaceholder(s, o) {
  shape(s, 'roundRect', { x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: o.r, fill: { color: o.body } });
  const padY = o.padY === undefined ? o.pad : o.padY;
  shape(s, 'roundRect', {
    x: o.x + o.pad, y: o.y + padY, w: o.w - 2 * o.pad, h: o.h - 2 * padY,
    rectRadius: 0.02, fill: { color: o.screen },
  });
}

/* ------------------------------------------------------------------ icons */
/* Small line-art glyphs from the source deck, rebuilt out of native shapes.  */

const ICONS = {
  /** wall safe: outlined box holding a dial + latch bars, on two little feet */
  safe(s, x, y, w, h, c) {
    const lw = Math.max(1, h * 5);
    shape(s, 'roundRect', { x, y, w, h: h * 0.84, rectRadius: 0.06, fill: { type: 'none' }, line: { color: c, width: lw } });
    shape(s, 'roundRect', {
      x: x + 0.13 * w, y: y + 0.14 * h, w: 0.74 * w, h: 0.55 * h,
      rectRadius: 0.05, fill: { type: 'none' }, line: { color: c, width: lw * 0.75 },
    });
    shape(s, 'ellipse', { x: x + 0.22 * w, y: y + 0.28 * h, w: 0.26 * w, h: 0.26 * h, fill: { color: c } });
    box(s, x + 0.6 * w, y + 0.29 * h, 0.16 * w, 0.055 * h, c);
    box(s, x + 0.6 * w, y + 0.45 * h, 0.16 * w, 0.055 * h, c);
    box(s, x + 0.14 * w, y + 0.84 * h, 0.13 * w, 0.16 * h, c);
    box(s, x + 0.73 * w, y + 0.84 * h, 0.13 * w, 0.16 * h, c);
  },

  /** wallet / card holder */
  wallet(s, x, y, w, h, c) {
    shape(s, 'roundRect', { x, y: y + 0.18 * h, w, h: h * 0.82, rectRadius: 0.03, fill: { color: c } });
    box(s, x + 0.06 * w, y, 0.66 * w, 0.26 * h, c);
    shape(s, 'roundRect', {
      x: x + 0.7 * w, y: y + 0.44 * h, w: 0.36 * w, h: 0.3 * h,
      rectRadius: 0.02, fill: { color: WHITE }, line: { color: c, width: 1 },
    });
  },

  /** piggy bank: body + snout at the left, coin-slot ear on top, four trotters */
  piggy(s, x, y, w, h, c) {
    shape(s, 'ellipse', { x: x + 0.42 * w, y, w: 0.26 * w, h: 0.26 * h, fill: { color: c } });
    shape(s, 'ellipse', { x: x + 0.16 * w, y: y + 0.2 * h, w: 0.84 * w, h: 0.6 * h, fill: { color: c } });
    shape(s, 'ellipse', { x, y: y + 0.4 * h, w: 0.26 * w, h: 0.24 * h, fill: { color: c } });
    [0.24, 0.4, 0.62, 0.78].forEach(function (f) {
      box(s, x + f * w, y + 0.68 * h, 0.1 * w, 0.22 * h, c);
    });
  },

  /** rising bars behind a marker line that ends in an arrow head */
  chart(s, x, y, w, h, c) {
    [0.35, 0.58, 0.47, 0.64].forEach(function (f, i) {
      box(s, x + i * 0.27 * w, y + (1 - f) * h, 0.16 * w, f * h, c);
    });
    const pts = [[0.04, 0.55], [0.3, 0.18], [0.56, 0.36], [0.9, 0.02]];
    for (let i = 1; i < pts.length; i++) {
      s.addShape('line', {
        x: x + Math.min(pts[i - 1][0], pts[i][0]) * w,
        y: y + Math.min(pts[i - 1][1], pts[i][1]) * h,
        w: Math.abs(pts[i][0] - pts[i - 1][0]) * w,
        h: Math.abs(pts[i][1] - pts[i - 1][1]) * h,
        flipV: pts[i][1] < pts[i - 1][1],
        line: { color: c, width: Math.max(1, h * 3) },
      });
    }
    pts.slice(0, 3).forEach(function (p) {
      shape(s, 'ellipse', { x: x + p[0] * w - 0.06 * w, y: y + p[1] * h - 0.08 * h, w: 0.13 * w, h: 0.16 * h, fill: { color: c } });
    });
    shape(s, 'triangle', { x: x + 0.79 * w, y, w: 0.21 * w, h: 0.2 * h, fill: { color: c }, rotate: 45 });
  },

  /** desktop monitor showing a chart */
  monitor(s, x, y, w, h, c) {
    shape(s, 'roundRect', { x, y, w, h: h * 0.74, rectRadius: 0.03, fill: { type: 'none' }, line: { color: c, width: 2 } });
    box(s, x + 0.44 * w, y + 0.74 * h, 0.12 * w, 0.16 * h, c);
    box(s, x + 0.26 * w, y + 0.9 * h, 0.48 * w, 0.1 * h, c);
    [0.24, 0.4, 0.16].forEach((f, i) => box(s, x + (0.5 + i * 0.16) * w, y + (0.56 - f) * h, 0.09 * w, f * h, c));
    shape(s, 'ellipse', { x: x + 0.16 * w, y: y + 0.2 * h, w: 0.2 * w, h: 0.25 * h, fill: { color: c } });
  },

  /** cog with a dollar sign */
  gear(s, x, y, w, h, c, bg) {
    shape(s, 'gear6', { x, y, w, h, fill: { color: c } });
    shape(s, 'ellipse', { x: x + 0.26 * w, y: y + 0.26 * h, w: 0.48 * w, h: 0.48 * h, fill: { color: bg } });
    text(s, '$', {
      x: x + 0.2 * w, y: y + 0.24 * h, w: 0.6 * w, h: 0.52 * h,
      fontSize: Math.max(6, Math.round(h * 34)), bold: true, color: c, align: 'center', valign: 'middle', margin: 0,
    });
  },

  /** map pin with a dollar sign */
  pin(s, x, y, w, h, c, bg) {
    shape(s, 'ellipse', { x, y, w, h: h * 0.78, fill: { color: c } });
    shape(s, 'triangle', { x: x + 0.24 * w, y: y + 0.52 * h, w: 0.52 * w, h: 0.48 * h, fill: { color: c }, flipV: true });
    text(s, '$', {
      x: x + 0.15 * w, y: y + 0.1 * h, w: 0.7 * w, h: 0.5 * h,
      fontSize: Math.max(6, Math.round(h * 26)), bold: true, color: bg, align: 'center', valign: 'middle', margin: 0,
    });
  },

  /** circular "refresh" ring around a dollar sign */
  refresh(s, x, y, w, h, c) {
    shape(s, 'donut', { x, y, w, h, fill: { color: c } });
    text(s, '$', {
      x: x + 0.2 * w, y: y + 0.22 * h, w: 0.6 * w, h: 0.56 * h,
      fontSize: Math.max(8, Math.round(h * 22)), bold: true, color: c, align: 'center', valign: 'middle', margin: 0,
    });
  },

  /** facebook / twitter / instagram / linkedin row */
  social(s, x, y, w, h, c) {
    const g = w / 3.75;
    text(s, 'f', { x, y, w: g * 0.5, h, fontSize: 13, bold: true, color: c, align: 'center', valign: 'middle', margin: 0 });
    shape(s, 'triangle', { x: x + g * 0.95, y: y + 0.2 * h, w: g * 0.62, h: h * 0.6, fill: { color: c }, rotate: 110 });
    shape(s, 'roundRect', {
      x: x + g * 1.95, y, w: h, h, rectRadius: 0.05,
      fill: { type: 'none' }, line: { color: c, width: 1.5 },
    });
    shape(s, 'ellipse', { x: x + g * 1.95 + h * 0.3, y: y + h * 0.3, w: h * 0.4, h: h * 0.4, fill: { color: c } });
    text(s, 'in', { x: x + g * 2.9, y, w: g * 0.85, h, fontSize: 11, bold: true, color: c, align: 'center', valign: 'middle', margin: 0 });
  },
};

/* --------------------------------------------------------------- geometry */

/** slide 2 - full-height panel whose right edge bulges out in a circular arc */
const BULGE_PANEL = [
  { x: 0, y: 0, moveTo: true },
  { x: 9.904, y: 0 },
  { x: 9.960, y: 0.088 },
  { x: 10.983, y: 3.750, curve: { type: 'cubic', x1: 10.609, y1: 1.155, x2: 10.983, y2: 2.409 } },
  { x: 9.960, y: 7.412, curve: { type: 'cubic', x1: 10.983, y1: 5.091, x2: 10.609, y2: 6.344 } },
  { x: 9.904, y: 7.5 },
  { x: 0, y: 7.5 },
  { close: true },
];

/** slide 28 - the two halves of the exploded pie (path units already inches) */
const PIE_LEFT = [
  { x: 1.912, y: 0, moveTo: true },
  { x: 1.912, y: 2.372 },
  { x: 1.042, y: 3.878 },
  { x: 0.919, y: 3.803 },
  { x: 0.000, y: 2.075, curve: { type: 'cubic', x1: 0.364, y1: 3.428, x2: 0.000, y2: 2.794 } },
  { x: 1.871, y: 0.002, curve: { type: 'cubic', x1: 0.000, y1: 0.996, x2: 0.820, y2: 0.109 } },
  { x: 1.912, y: 0 },
  { close: true },
];
const PIE_RIGHT = [
  { x: 0.901, y: 0, moveTo: true },
  { x: 0.941, y: 0.002 },
  { x: 2.812, y: 2.075, curve: { type: 'cubic', x1: 1.992, y1: 0.109, x2: 2.812, y2: 0.996 } },
  { x: 0.729, y: 4.159, curve: { type: 'cubic', x1: 2.812, y1: 3.226, x2: 1.879, y2: 4.159 } },
  { x: 0.109, y: 4.065, curve: { type: 'cubic', x1: 0.513, y1: 4.159, x2: 0.305, y2: 4.126 } },
  { x: 0.000, y: 4.025 },
  { x: 0.901, y: 2.465 },
  { x: 0.901, y: 0 },
  { close: true },
];

/* ---------------------------------------------------------------- builders */

const SLIDES = [];

/* 1 - title -------------------------------------------------------------- */
SLIDES.push(function (s) {
  box(s, 0, 0, SLIDE_W, SLIDE_H, BLACK, { fill: { color: BLACK, transparency: 60 } });
  text(s, [
    { text: 'FINANCIAL ', options: { fontSize: 72, fontFace: F_SEMI, breakLine: true } },
    { text: 'PLANNER', options: { fontSize: 48, fontFace: F_MED } },
  ], { x: 2.473, y: 2.211, w: 8.387, h: 2.121, align: 'center', color: WHITE, charSpacing: 8 });
  text(s, LOREM_MED, { x: 3.544, y: 4.751, w: 6.245, h: 0.627, fontSize: 11, color: WHITE, align: 'center', lineSpacingMultiple: 1.5 });
  footer(s, WHITE);
  text(s, '225 Beeghley Street, Glen Ellyn\nIllnois, 60138',
    { x: 10.248, y: 6.332, w: 2.58, h: 0.789, fontSize: 11, color: WHITE, align: 'right', lineSpacingMultiple: 2 });
});

/* 2 - head office -------------------------------------------------------- */
SLIDES.push(function (s) {
  box(s, 4.97, 0, 8.413, 7.544, GREEN);
  text(s, 'v', { x: 4.97, y: 0, w: 8.413, h: 7.544, fontSize: 18, color: WHITE, align: 'center', valign: 'middle' });
  shape(s, 'custGeom', { x: 0, y: 0, w: 10.983, h: 7.5, points: BULGE_PANEL, fill: { color: BLACK, transparency: 50 } });
  text(s, 'Financial Plan', { x: 1.649, y: 1.329, w: 1.482, h: 0.303, wrap: false, fontSize: 12, charSpacing: 1, fontFace: F_LIGHT, color: WHITE });
  text(s, 'Head Office', { x: 1.61, y: 1.663, w: 4.636, h: 0.841, fontSize: 44, fontFace: F_SEMI, color: WHITE });
  body(s, LOREM_LONG, { x: 1.633, y: 2.595, w: 5.203, h: 1.529, color: WHITE });
  text(s, 'Alex Johnson', { x: 1.616, y: 4.444, w: 2.325, h: 0.656, fontSize: 24, bold: true, fontFace: F_SEMI, color: WHITE, lineSpacingMultiple: 1.5 });
  text(s, 'Financial Planner', { x: 1.633, y: 5.036, w: 1.919, h: 0.349, fontSize: 11, color: WHITE, lineSpacingMultiple: 1.5 });
  footer(s, WHITE);
});

/* 3 - our journey -------------------------------------------------------- */
SLIDES.push(function (s) {
  box(s, 0, 4.92, SLIDE_W, 2.58, GREEN);
  body(s, LOREM_LONG, { x: 1.526, y: 1.053, w: 5.878, h: 1.159 });
  heading(s, { ex: 8.145, ey: 1.146, x: 8.09, y: 1.481, w: 4.636, title: 'Our Journey' });
  subhead(s, 'Since 1992', { x: 8.098, y: 2.636, w: 1.919, h: 0.518 });
  body(s, LOREM_MED, { x: 8.127, y: 3.133, w: 4.27, h: 1.159 });
  ICONS.refresh(s, 8.145, 5.368, 0.773, 0.773, WHITE);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna.',
    { x: 9.214, y: 5.201, w: 3.183, h: 1.159, color: WHITE });
  footer(s, WHITE);
});

/* 4 - about us ----------------------------------------------------------- */
SLIDES.push(function (s) {
  box(s, 8.333, 0, 5.0, 7.5, GREEN);
  heading(s, { ex: 0.986, ey: 1.329, x: 0.947, y: 1.663, w: 4.636, title: 'About Us' });
  body(s, LOREM_LONG, { x: 0.986, y: 2.595, w: 4.982, h: 1.529 });
  subhead(s, 'Description Here', { x: 0.947, y: 4.451, w: 2.265, h: 0.518 });
  body(s, LOREM_MED, { x: 0.947, y: 4.99, w: 4.982, h: 0.789 });
  footer(s);
});

/* 5 - our story ---------------------------------------------------------- */
SLIDES.push(function (s) {
  box(s, 0, 0, SLIDE_W, 6.03, GREEN);
  box(s, 0.635, 0.65, 2.829, 4.689, WHITE, { shadow: shadow() });
  ICONS.safe(s, 1.678, 1.032, 0.796, 0.717, GREEN);
  text(s, 'Keyword Here', { x: 1.106, y: 1.912, w: 1.939, h: 0.518, fontSize: 18, fontFace: F_MED, color: DARK, align: 'center', lineSpacingMultiple: 1.5 });
  body(s, LOREM_MED, { x: 0.926, y: 2.456, w: 2.247, h: 1.9, align: 'center' });
  shape(s, 'roundRect', { x: 1.512, y: 4.682, w: 1.129, h: 0.331, rectRadius: 0.061, fill: { color: GREEN } });
  text(s, 'Read More', { x: 1.512, y: 4.682, w: 1.129, h: 0.331, fontSize: 12, color: WHITE, align: 'center', valign: 'middle', margin: 0 });
  heading(s, { ex: 7.831, ey: 0.958, x: 7.776, y: 1.292, w: 3.539, title: 'Our Story', color: WHITE, eyeColor: WHITE });
  stat(s, 7.97, 2.009, '82%');
  body(s, LOREM_LONG, { x: 7.858, y: 3.066, w: 4.428, h: 1.529, color: WHITE });
  footer(s);
});

/* 6 - what we do --------------------------------------------------------- */
SLIDES.push(function (s) {
  box(s, 7.705, 0, 5.628, 7.5, GREEN);
  heading(s, { ex: 0.986, ey: 0.961, x: 0.931, y: 1.296, w: 4.636, title: 'What We Do' });
  body(s, LOREM_LONG, { x: 0.986, y: 2.228, w: 4.597, h: 1.529 });
  ICONS.safe(s, 1.032, 4.359, 0.439, 0.395, GREEN);
  subhead(s, 'Keyword Here', { x: 1.675, y: 4.236, w: 1.992, h: 0.518 });
  body(s, LOREM_MED, { x: 0.986, y: 4.785, w: 4.597, h: 0.789 });
  ICONS.wallet(s, 9.539, 1.334, 0.375, 0.375, WHITE);
  body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit.', { x: 10.079, y: 1.113, w: 2.27, h: 0.789, color: WHITE });
  footer(s);
});

/* 7 - our work ----------------------------------------------------------- */
SLIDES.push(function (s) {
  gradientBand(s, 0, 0, SLIDE_W, 6.302, GREEN);
  box(s, 7.581, 3.868, 4.836, 1.785, WHITE, { shadow: shadow() });
  heading(s, { ex: 7.581, ey: 0.687, x: 7.511, y: 1.021, w: 3.576, title: 'Our Work', color: WHITE, eyeColor: WHITE });
  body(s, LOREM_LONG, { x: 7.581, y: 1.931, w: 4.836, h: 1.529, color: WHITE });
  ICONS.piggy(s, 7.941, 4.208, 0.727, 0.688, GREEN);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna.',
    { x: 8.941, y: 4.083, w: 3.329, h: 1.159 });
  footer(s);
});

/* 8 - what we offer ------------------------------------------------------ */
SLIDES.push(function (s) {
  box(s, 7.194, 3.5, 2.493, 4.0, GREEN, { shadow: shadow() });
  heading(s, { ex: 0.907, ey: 1.031, x: 0.822, y: 1.366, w: 3.576, h: 1.447, title: 'What We Offer' });
  subhead(s, 'The Best in State', { x: 6.667, y: 0.875, w: 4.464, h: 0.518 });
  body(s, LOREM_LONG, { x: 6.714, y: 1.394, w: 5.946, h: 1.159 });
  ICONS.chart(s, 8.097, 3.831, 0.687, 0.651, WHITE);
  text(s, 'Keyword Here', { x: 7.429, y: 4.581, w: 1.995, h: 0.507, fontSize: 18, bold: true, color: WHITE, align: 'center', lineSpacingMultiple: 1.5 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et.',
    { x: 7.384, y: 5.173, w: 2.113, h: 1.9, color: WHITE, align: 'center' });
  footer(s);
});

/* 9 - financial tips ----------------------------------------------------- */
SLIDES.push(function (s) {
  box(s, 0, 0, SLIDE_W, 4.391, GREEN);
  subhead(s, 'Description Here', { x: 0.59, y: 0.582, w: 2.265, h: 0.518, color: WHITE });
  body(s, LOREM_LONG2 + ' Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor.',
    { x: 0.606, y: 1.12, w: 4.139, h: 2.27, color: WHITE });
  box(s, 5.088, 0.712, 3.701, 2.849, WHITE, { shadow: shadow() });
  ICONS.monitor(s, 5.563, 1.099, 0.511, 0.418, GREEN);
  subhead(s, 'Keyword Here', { x: 6.154, y: 1.0, w: 2.088, h: 0.518 });
  body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do eiusmod tempor incididunt. Ut labore et dolore magna aliqua ut enim ad minim veniam.',
    { x: 5.478, y: 1.592, w: 3.041, h: 1.529 });
  heading(s, { ex: 0.632, ey: 5.1, x: 0.59, y: 5.434, w: 4.267, title: 'Financial Tips' });
  body(s, LOREM_LONG, { x: 6.5, y: 4.912, w: 6.077, h: 1.159 });
  footer(s);
});

/* 10 - retirement savings ------------------------------------------------ */
SLIDES.push(function (s) {
  box(s, 7.09, 1.994, 5.819, 4.99, GREEN, { shadow: shadow() });
  subhead(s, '401k Plan', { x: 7.511, y: 2.333, w: 2.265, h: 0.518, color: WHITE });
  body(s, LOREM_LONG2, { x: 7.527, y: 2.872, w: 4.886, h: 1.529, color: WHITE });
  subhead(s, 'Another Plan', { x: 7.495, y: 4.751, w: 2.265, h: 0.518, color: WHITE });
  body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam quis nostrud exercitation ullamco.',
    { x: 7.511, y: 5.29, w: 4.886, h: 1.159, color: WHITE });
  heading(s, { ex: 5.926, ey: 0.524, x: 3.907, y: 0.858, w: 5.519, title: 'Retirement Savings', align: 'center', eyeColor: DARK });
  footer(s);
});

/* 11 - credit cards ------------------------------------------------------ */
SLIDES.push(function (s) {
  box(s, 8.881, 4.284, 4.453, 2.015, GREEN, { shadow: shadow() });
  heading(s, { ex: 1.053, ey: 0.899, x: 1.011, y: 1.233, w: 4.267, title: 'Credit Cards' });
  body(s, LOREM_LONG, { x: 1.053, y: 2.071, w: 4.453, h: 1.529 });
  [['Percentage 1', '76%', 1.025, 1.235], ['Percentage 2', '87%', 2.444, 2.654]].forEach(function (r) {
    text(s, r[0], { x: 6.69, y: r[2], w: 1.553, h: 0.425, fontSize: 14, fontFace: F_MED, color: DARK, lineSpacingMultiple: 1.5 });
    text(s, r[1], { x: 6.627, y: r[3], w: 1.399, h: 1.027, fontSize: 40, color: DARK, align: 'center', lineSpacingMultiple: 1.5 });
  });
  ICONS.gear(s, 9.158, 4.569, 0.533, 0.533, WHITE, GREEN);
  subhead(s, 'Keyword Here', { x: 9.756, y: 4.569, w: 2.015, h: 0.518, color: WHITE });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore.',
    { x: 9.138, y: 5.173, w: 4.048, h: 0.789, color: WHITE });
  footer(s);
});

/* 12 - our accountants --------------------------------------------------- */
SLIDES.push(function (s) {
  box(s, 0, 0, SLIDE_W, 6.175, GREEN);
  heading(s, { ex: 5.926, ey: 0.524, x: 4.241, y: 0.858, w: 4.852, title: 'Our Accountants', align: 'center', color: WHITE, eyeColor: WHITE });
  [
    { name: 'Maximus Kord', x: 3.886, nx: 3.915, y: 2.188, sy: 2.663, by: 3.174, icons: [4.02, 4.826] },
    { name: 'Lupita Kongo', x: 9.969, nx: 9.998, y: 2.213, sy: 2.688, by: 3.199, icons: [10.104, 4.851] },
  ].forEach(function (p) {
    subhead(s, p.name, { x: p.x, y: p.y, w: 2.222, h: 0.518, color: WHITE });
    text(s, 'Financial Planner', { x: p.nx, y: p.sy, w: 1.919, h: 0.372, fontSize: 12, italic: true, color: WHITE, lineSpacingMultiple: 1.5 });
    body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do eiusmod tempor incididunt ut.',
      { x: p.nx, y: p.by, w: 2.704, h: 1.159, color: WHITE });
    ICONS.social(s, p.icons[0], p.icons[1], 1.571, 0.231, WHITE);
  });
  footer(s);
});

/* 13 - saving for college ------------------------------------------------ */
SLIDES.push(function (s) {
  gradientBand(s, 0, 0, SLIDE_W, 4.211, GREEN);
  text(s, 'k', { x: 0, y: 0, w: SLIDE_W, h: 4.211, fontSize: 18, color: WHITE, align: 'center', valign: 'middle' });
  ICONS.pin(s, 7.986, 0.829, 0.431, 0.539, WHITE, GREEN);
  subhead(s, 'Keyword Here', { x: 8.691, y: 0.743, w: 2.151, h: 0.518, color: WHITE });
  body(s, LOREM_LONG, { x: 8.691, y: 1.367, w: 3.687, h: 1.9, color: WHITE });
  text(s, 'Saving for College', { x: 1.554, y: 4.761, w: 4.267, h: 1.447, fontSize: 40, fontFace: F_SEMI, color: DARK });
  body(s, LOREM_LONG, { x: 6.464, y: 4.761, w: 6.057, h: 1.159 });
  footer(s);
});

/* 14 - financial strategy ------------------------------------------------ */
SLIDES.push(function (s) {
  const CARD_TEXT =
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ' +
    'ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut.';
  box(s, 0, 0.609, 5.171, 2.504, GREEN, { shadow: shadow() });
  box(s, 0, 3.537, 5.171, 2.504, GREEN, { shadow: shadow() });
  ICONS.gear(s, 0.461, 0.978, 0.807, 0.807, WHITE, GREEN);
  ICONS.chart(s, 0.483, 4.018, 0.687, 0.651, WHITE);
  body(s, CARD_TEXT, { x: 1.599, y: 0.835, w: 3.274, h: 1.9, color: WHITE });
  body(s, CARD_TEXT, { x: 1.591, y: 3.757, w: 3.274, h: 1.9, color: WHITE });
  heading(s, { ex: 10.049, ey: 0.978, x: 10.007, y: 1.313, w: 3.215, h: 1.447, title: 'Financial Strategy' });
  body(s, LOREM_LONG, { x: 10.049, y: 2.947, w: 2.822, h: 2.64 });
  footer(s);
});

/* 15 - managing the debts ------------------------------------------------ */
SLIDES.push(function (s) {
  box(s, 6.678, 0, 6.655, 7.5, GREEN);
  heading(s, { ex: 0.878, ey: 0.787, x: 0.837, y: 1.122, w: 4.267, h: 1.447, title: 'Managing The Debts' });
  body(s, LOREM_LONG, { x: 0.862, y: 2.649, w: 4.852, h: 1.529 });
  ICONS.gear(s, 0.878, 4.54, 0.462, 0.462, GREEN, WHITE);
  subhead(s, 'Description Here', { x: 1.492, y: 4.484, w: 2.265, h: 0.518 });
  body(s, LOREM_MED, { x: 1.508, y: 5.022, w: 4.222, h: 1.159 });
  footer(s);
});

/* 16 - college loans ----------------------------------------------------- */
SLIDES.push(function (s) {
  box(s, 0, 0, SLIDE_W, SLIDE_H, BLACK, { fill: { color: BLACK, transparency: 50 } });
  box(s, 9.855, 3.877, 3.479, 2.361, GREEN, { shadow: shadow() });
  subhead(s, 'Keyword Here', { x: 10.368, y: 4.418, w: 2.16, h: 0.518, color: WHITE });
  stat(s, 10.484, 4.694, '82%');
  heading(s, { ex: 5.926, ey: 0.524, x: 4.241, y: 0.858, w: 4.852, title: 'College Loans', align: 'center', color: WHITE, eyeColor: WHITE });
  body(s, LOREM_LONG, { x: 1.57, y: 1.807, w: 9.93, h: 0.789, color: WHITE, align: 'center' });
  footer(s, WHITE);
});

/* 17 - in-house accountants ---------------------------------------------- */
SLIDES.push(function (s) {
  box(s, 0, 0, 5.659, 7.5, GREEN);
  heading(s, { ex: 7.674, ey: 0.697, x: 7.632, y: 1.032, w: 4.852, h: 1.447, title: 'In-House Accountants' });
  body(s, LOREM_LONG, { x: 7.658, y: 2.559, w: 4.923, h: 1.529 });
  [7.634, 10.316].forEach(function (x, i) {
    subhead(s, 'Description Here', { x, y: 4.431, w: 2.265, h: 0.518 });
    body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do eiusmod tempor incididunt.',
      { x: x + 0.016, y: 4.97, w: 2.58, h: 1.159 });
  });
  footer(s, WHITE);
});

/* 18 - before retirement ------------------------------------------------- */
SLIDES.push(function (s) {
  box(s, 8.947, 0, 4.387, 7.5, GREEN);
  heading(s, { ex: 0.943, ey: 0.897, x: 0.902, y: 1.232, w: 4.852, h: 1.447, title: 'Before Retirement' });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam quis nostrud exercitation ullamco.',
    { x: 0.927, y: 2.759, w: 4.923, h: 1.159 });
  ICONS.piggy(s, 0.957, 4.678, 0.727, 0.688, GREEN);
  body(s, LOREM_MED, { x: 2.005, y: 4.547, w: 2.734, h: 1.529 });
  ICONS.pin(s, 5.102, 4.714, 0.475, 0.593, GREEN, WHITE);
  body(s, LOREM_MED, { x: 5.894, y: 4.569, w: 2.734, h: 1.529 });
  footer(s);
});

/* 19 - track your expenses ----------------------------------------------- */
SLIDES.push(function (s) {
  box(s, 0, 0, 3.429, 7.5, GREEN);
  heading(s, { ex: 6.613, ey: 0.937, x: 6.571, y: 1.271, w: 4.852, h: 1.447, title: 'Track Your Expenses' });
  body(s, LOREM_LONG, { x: 6.597, y: 2.798, w: 5.863, h: 1.159 });
  ICONS.piggy(s, 6.686, 4.494, 0.573, 0.543, GREEN);
  body(s, LOREM_MED, { x: 7.443, y: 4.319, w: 5.003, h: 0.789 });
  ICONS.chart(s, 6.725, 5.611, 0.534, 0.505, GREEN);
  body(s, LOREM_MED, { x: 7.443, y: 5.41, w: 5.003, h: 0.789 });
  footer(s, WHITE);
});

/* 20 - credits and debts ------------------------------------------------- */
SLIDES.push(function (s) {
  box(s, 1.222, 0, 12.111, 6.127, GREEN);
  heading(s, { ex: 1.975, ey: 0.801, x: 1.934, y: 1.135, w: 5.586, title: 'Credits and Debts', color: WHITE, eyeColor: WHITE });
  body(s, LOREM_LONG, { x: 1.975, y: 2.085, w: 6.452, h: 1.159, color: WHITE });
  body(s, LOREM_LONG, { x: 1.975, y: 3.356, w: 6.452, h: 1.159, color: WHITE });
  footer(s);
});

/* 21 - monthly budgeting ------------------------------------------------- */
SLIDES.push(function (s) {
  box(s, 9.08, 3.016, 4.253, 3.217, GREEN, { shadow: shadow() });
  heading(s, { ex: 0.661, ey: 0.651, x: 0.619, y: 0.985, w: 4.081, h: 1.447, title: 'Monthly Budgeting' });
  subhead(s, 'The Best Strategy', { x: 5.982, y: 0.667, w: 4.464, h: 0.518 });
  body(s, LOREM_LONG, { x: 5.996, y: 1.185, w: 6.019, h: 1.159 });
  ICONS.gear(s, 9.507, 3.434, 0.518, 0.518, WHITE, GREEN);
  subhead(s, 'Description Here', { x: 10.285, y: 3.306, w: 2.421, h: 0.518, color: WHITE });
  body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam.',
    { x: 10.299, y: 3.825, w: 2.775, h: 1.9, color: WHITE });
  footer(s);
});

/* 22 - control cashflow -------------------------------------------------- */
SLIDES.push(function (s) {
  box(s, 0, 0, 4.701, 7.5, GREEN);
  stat(s, 0.584, 4.095, '82%');
  body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua..',
    { x: 0.471, y: 5.099, w: 3.957, h: 1.159, color: WHITE });
  body(s, LOREM_LONG2, { x: 5.386, y: 4.457, w: 3.388, h: 2.27 });
  heading(s, { ex: 9.733, ey: 4.979, x: 9.659, y: 5.314, w: 3.388, h: 1.447, title: 'Control Cashflow' });
  footer(s, WHITE);
});

/* 23 - income report ----------------------------------------------------- */
SLIDES.push(function (s) {
  box(s, 8.632, 0, 4.701, 7.5, GREEN);
  // laptop mock-up placeholder: lid + screen, sitting on the keyboard deck
  devicePlaceholder(s, { x: 7.074, y: 1.817, w: 6.243, h: 4.24, r: 0.05, body: '1D1E22', screen: SCREEN, pad: 0.214, padY: 0.307 });
  shape(s, 'trapezoid', { x: 6.359, y: 6.057, w: 6.958, h: 0.143, flipV: true, fill: { color: '4D4C51' } });
  box(s, 6.359, 6.057, 6.958, 0.028, 'CFCFD2');
  heading(s, { ex: 0.819, ey: 0.915, x: 0.778, y: 1.249, w: 4.852, title: 'Income Report' });
  body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam quis nostrud exercitation ullamco laboris.',
    { x: 0.819, y: 2.204, w: 5.085, h: 1.159 });
  const head = ['Month', 'Lawyer', 'Pilot', 'Designer'];
  const rows = [['January', '$23k', '$10k', '$8k'], ['February', '$25k', '$15k', '$20k'], ['March', '$30k', '$20k', '$50k']];
  const border = [{ type: 'solid', color: WHITE, pt: 1.5 }, { type: 'solid', color: WHITE, pt: 1.5 },
    { type: 'solid', color: WHITE, pt: 1.5 }, { type: 'solid', color: WHITE, pt: 1.5 }];
  const table = [head.map(function (t) {
    return { text: t, options: { fill: { color: GREEN }, color: WHITE, fontSize: 14, bold: false, border } };
  })];
  rows.forEach(function (r, i) {
    const bg = i === 1 ? BAND : SHADE;
    table.push(r.map(function (t, c) {
      return { text: t, options: { fill: { color: bg }, color: GRAY, fontSize: c === 0 ? 11 : 12, border } };
    }));
  });
  s.addTable(table, {
    x: 0.848, y: 3.754, w: 4.958, colW: [1.24, 1.24, 1.24, 1.24], rowH: [0.459, 0.406, 0.406, 0.406],
    align: 'center', valign: 'middle', fontFace: F_BODY, margin: 0,
  });
  footer(s);
});

/* 24 - download our app -------------------------------------------------- */
SLIDES.push(function (s) {
  box(s, 0, 0, 5.238, 7.5, GREEN);
  // phone + tablet mock-up placeholders
  devicePlaceholder(s, { x: 1.042, y: 2.915, w: 1.648, h: 3.304, r: 0.18, body: '2B2B2D', screen: GREEN, pad: 0.07 });
  devicePlaceholder(s, { x: 3.379, y: 1.083, w: 3.596, h: 5.134, r: 0.09, body: '2B2B2D', screen: 'F1F1F1', pad: 0.216, padY: 0.484 });
  shape(s, 'ellipse', { x: 5.0, y: 5.79, w: 0.34, h: 0.34, fill: { type: 'none' }, line: { color: '6A6A6E', width: 1 } });
  heading(s, { ex: 7.696, ey: 0.915, x: 7.655, y: 1.249, w: 3.788, h: 1.447, title: 'Download Our App' });
  body(s, LOREM_LONG2, { x: 7.696, y: 2.961, w: 4.493, h: 1.529 });
  [
    { x: 7.796, ix: 8.025, tx: 8.476, tw: 1.311, label: 'Download on GooglePlay', badge: 'play' },
    { x: 10.227, ix: 10.414, tx: 10.899, tw: 1.243, label: 'Download on AppStore', badge: 'apple' },
  ].forEach(function (b) {
    shape(s, 'roundRect', { x: b.x, y: 5.015, w: 1.991, h: 0.76, rectRadius: 0.101, fill: { color: DARK } });
    if (b.badge === 'play') {
      shape(s, 'triangle', { x: b.ix, y: 5.175, w: 0.386, h: 0.44, fill: { color: GREEN }, rotate: 90 });
    } else {
      shape(s, 'ellipse', { x: b.ix + 0.03, y: 5.19, w: 0.33, h: 0.36, fill: { color: WHITE } });
      box(s, b.ix + 0.2, 5.149, 0.05, 0.09, WHITE);
    }
    text(s, b.label, { x: b.tx, y: 5.112, w: b.tw, h: 0.53, fontSize: 11, color: WHITE, lineSpacingMultiple: 1.2 });
  });
  footer(s, WHITE);
});

/* 25 - scheduling (calendar) --------------------------------------------- */
const CAL_COLS = [8.145, 8.696, 9.248, 9.786, 10.324, 10.861, 11.399];
const CAL_ROWS = [2.422, 2.96, 3.498, 4.036, 4.574, 5.129];
/** per cell: [label, fill, colour]  -  null = no cell drawn */
const CAL_CELLS = [
  [['2', WHITE, WHITE], ['3', WHITE, WHITE], ['', WHITE, WHITE], ['', WHITE, WHITE], ['', WHITE, WHITE], ['', WHITE, WHITE], ['1', SHADE, LIME]],
  [['2', SHADE, GRAY], ['3', SHADE, GRAY], ['4', SHADE, GRAY], ['5', SHADE, GRAY], ['6', SHADE, GRAY], ['7', SHADE, GRAY], ['8', SHADE, LIME]],
  [['9', SHADE, GRAY], ['10', SHADE, GRAY], ['11', GREEN, WHITE], ['12', SHADE, GRAY], ['13', SHADE, GRAY], ['14', SHADE, GRAY], ['15', SHADE, LIME]],
  [['16', SHADE, GRAY], ['17', SHADE, GRAY], ['18', SHADE, GRAY], ['19', SHADE, GRAY], ['20', SHADE, GRAY], ['21', SHADE, GRAY], ['22', SHADE, LIME]],
  [['23', SHADE, GRAY], ['25', SHADE, GRAY], ['25', SHADE, GRAY], ['26', SHADE, GRAY], ['27', GREEN, WHITE], ['28', SHADE, GRAY], ['29', SHADE, LIME]],
  [['30', SHADE, GRAY], ['31', SHADE, GRAY], ['1', WHITE, SILVER], ['2', WHITE, SILVER], ['3', WHITE, SILVER], ['4', WHITE, SILVER], ['5', WHITE, SILVER]],
];
SLIDES.push(function (s) {
  shape(s, 'roundRect', { x: 7.822, y: 1.194, w: 4.305, h: 4.839, rectRadius: 0.053, fill: { color: WHITE }, shadow: shadow() });
  box(s, 7.822, 1.194, 4.305, 0.496, GREEN);
  text(s, 'September, 2019', { x: 7.822, y: 1.194, w: 4.305, h: 0.496, fontSize: 20, color: WHITE, align: 'center' });
  ['M', 'T', 'W', 'T', 'F', 'S', 'S'].forEach(function (d, c) {
    text(s, d, { x: CAL_COLS[c], y: 1.884, w: 0.496, h: 0.496, fontSize: 10.5, bold: true, color: DARK, align: 'center', valign: 'middle', margin: 0 });
  });
  CAL_CELLS.forEach(function (row, r) {
    row.forEach(function (cell, c) {
      shape(s, 'roundRect', { x: CAL_COLS[c], y: CAL_ROWS[r], w: 0.496, h: 0.496, rectRadius: 0.083, fill: { color: cell[1] } });
      if (cell[0]) {
        text(s, cell[0], {
          x: CAL_COLS[c], y: CAL_ROWS[r], w: 0.496, h: 0.496, margin: 0,
          fontSize: 11, color: cell[2], align: 'center', valign: 'middle',
        });
      }
    });
  });
  heading(s, { ex: 0.958, ey: 1.043, x: 0.916, y: 1.378, w: 3.398, title: 'Scheduling' });
  body(s, LOREM_LONG2, { x: 0.958, y: 2.302, w: 5.412, h: 1.529 });
  body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.',
    { x: 0.958, y: 3.981, w: 5.412, h: 0.789 });
  shape(s, 'roundRect', { x: 1.041, y: 5.211, w: 1.211, h: 0.331, rectRadius: 0.052, fill: { color: GREEN } });
  text(s, 'Learn More', { x: 1.041, y: 5.211, w: 1.211, h: 0.331, fontSize: 12, color: WHITE, align: 'center', margin: 0 });
  footer(s);
});

/* 26 - infographic (doughnut + bars) ------------------------------------- */
SLIDES.push(function (s) {
  [[3.485, 3.009, GREEN], [4.009, 3.947, DARK], [2.952, 4.904, GREEN]].forEach(function (b) {
    box(s, 3.832, b[1], b[0], 0.593, b[2]);
  });
  box(s, 1.121, 2.495, 2.711, 3.497, WHITE, { shadow: shadow() });
  s.addChart('doughnut', [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: [8.2, 3.2] }], {
    x: 1.145, y: 2.729, w: 2.664, h: 1.81,
    chartColors: [GREEN, SHADE], holeSize: 75, showLegend: false, showValue: false,
    dataBorder: { pt: 0, color: WHITE }, firstSliceAng: 0,
  });
  text(s, '69%', { x: 1.741, y: 3.381, w: 1.553, h: 0.505, fontSize: 24, bold: true, color: GRAY, align: 'center' });
  body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do eiusmod tempor.',
    { x: 1.33, y: 4.548, w: 2.294, h: 1.159, align: 'center' });
  [[2.902, 3.207, GREEN, 2.822], [3.934, 4.228, DARK, 3.844], [4.975, 5.279, GREEN, 4.894]].forEach(function (k) {
    box(s, k[2] === DARK ? 9.465 : (k[0] === 4.975 ? 9.458 : 9.465), k[0], 0.214, 0.214, k[2]);
    text(s, 'Keyword', { x: 9.817, y: k[3], w: 1.301, h: 0.404, fontSize: 18, fontFace: F_MED, color: DARK });
    body(s, 'Lorem ipsum dolor sit amet consectetur.', { x: 9.385, y: k[1], w: 3.064, h: 0.419 });
  });
  heading(s, { ex: 0.661, ey: 0.651, x: 0.619, y: 0.985, w: 4.852, title: 'Infographic' });
  body(s, LOREM_LONG, { x: 6.014, y: 0.571, w: 6.019, h: 1.159 });
  footer(s);
});

/* 27 - infographic timeline ---------------------------------------------- */
SLIDES.push(function (s) {
  box(s, 0, 0, SLIDE_W, 3.509, GREEN);
  heading(s, { ex: 5.926, ey: 0.524, x: 4.241, y: 0.858, w: 4.852, title: 'Infographic Slide', align: 'center', color: WHITE, eyeColor: WHITE });
  body(s, LOREM_LONG, { x: 1.701, y: 1.905, w: 9.93, h: 0.789, color: WHITE, align: 'center' });
  // four drops: [stem x, stem y, stem h, dot x, dot y, colour, title x/y, copy x/y/w]
  const MARKS = [
    { sx: 2.297, sy: 2.581, sh: 1.807, dx: 2.226, dy: 4.352, c: GREEN, tx: 1.252, ty: 4.784, cx: 0.738, cy: 5.291, cw: 3.116 },
    { sx: 5.219, sy: 2.372, sh: 1.427, dx: 5.152, dy: 3.703, c: DARK, tx: 4.147, ty: 3.996, cx: 3.600, cy: 4.503, cw: 3.279 },
    { sx: 8.186, sy: 2.581, sh: 1.807, dx: 8.116, dy: 4.352, c: GREEN, tx: 7.141, ty: 4.784, cx: 6.628, cy: 5.291, cw: 3.116 },
    { sx: 11.109, sy: 2.372, sh: 1.427, dx: 11.041, dy: 3.703, c: DARK, tx: 10.036, ty: 3.996, cx: 9.489, cy: 4.503, cw: 3.279 },
  ];
  MARKS.forEach(function (m) {
    box(s, m.sx, m.sy, 0.057, m.sh, m.c);
    shape(s, 'ellipse', { x: m.dx, y: m.dy, w: 0.192, h: 0.192, fill: { color: m.c } });
    text(s, 'Subtitle Here', { x: m.tx, y: m.ty, w: 2.127, h: 0.518, fontSize: 18, fontFace: F_MED, color: DARK, align: 'center', lineSpacingMultiple: 1.5 });
    body(s, LOREM_SHORT, { x: m.cx, y: m.cy, w: m.cw, h: 0.789, align: 'center' });
  });
  footer(s);
});

/* 28 - infographic pie --------------------------------------------------- */
SLIDES.push(function (s) {
  heading(s, { ex: 5.926, ey: 0.524, x: 4.241, y: 0.858, w: 4.852, title: 'Infographic Slide', align: 'center', eyeColor: DARK });
  shape(s, 'custGeom', { x: 4.583, y: 2.185, w: 1.912, h: 3.878, points: PIE_LEFT, fill: { color: DARK } });
  shape(s, 'custGeom', { x: 5.938, y: 2.185, w: 2.812, h: 4.159, points: PIE_RIGHT, fill: { color: GREEN } });
  body(s, LOREM_LONG2, { x: 0.805, y: 2.397, w: 3.092, h: 2.27 });
  shape(s, 'roundRect', { x: 0.888, y: 5.332, w: 1.129, h: 0.331, rectRadius: 0.061, fill: { color: GREEN } });
  text(s, 'Read More', { x: 0.888, y: 5.332, w: 1.129, h: 0.331, fontSize: 12, color: WHITE, align: 'center', valign: 'middle', margin: 0 });
  [['57%', GREEN, 1.717, 2.805, 9.524, 9.554], ['43%', DARK, 4.015, 5.103, 9.532, 9.562]].forEach(function (p) {
    text(s, p[0], { x: p[4], y: p[2], w: 2.176, h: 1.184, fontSize: 48, color: p[1], lineSpacingMultiple: 1.5 });
    body(s, LOREM_SHORT, { x: p[5], y: p[3], w: 3.007, h: 0.789 });
  });
  footer(s);
});

/* 29 - infographic cards ------------------------------------------------- */
SLIDES.push(function (s) {
  box(s, 0, 3.224, SLIDE_W, 4.276, GREEN);
  heading(s, { ex: 5.926, ey: 0.524, x: 4.241, y: 0.858, w: 4.852, title: 'Infographic Slide', align: 'center', eyeColor: DARK });
  [
    { x: 2.075, num: '01', pct: '82%', cx: 2.204, ox: 2.764, nx: 2.796, tx: 2.022, bx: 1.728, oy: 2.571, ny: 2.607 },
    { x: 5.560, num: '02', pct: '73%', cx: 5.689, ox: 6.248, nx: 6.280, tx: 5.506, bx: 5.213, oy: 2.571, ny: 2.609 },
    { x: 9.008, num: '03', pct: '64%', cx: 9.137, ox: 9.697, nx: 9.729, tx: 8.955, bx: 8.726, oy: 2.571, ny: 2.609 },
  ].forEach(function (c, i) {
    box(s, c.x, 2.29, 2.02, 2.02, DARK, { shadow: shadow() });
    shape(s, 'ellipse', { x: c.ox, y: c.oy, w: 0.643, h: 0.643, fill: { color: WHITE } });
    text(s, c.num, { x: c.nx, y: c.ny, w: 0.579, h: 0.462, fontSize: 16, color: DARK, align: 'center', lineSpacingMultiple: 1.5 });
    text(s, c.pct, { x: c.cx, y: 3.081, w: 1.891, h: 1.003, fontSize: 40, color: WHITE, align: 'center', lineSpacingMultiple: 1.5 });
    text(s, 'Subtitle Here', { x: c.tx, y: i === 2 ? 4.574 : 4.566, w: 2.127, h: 0.518, fontSize: 18, fontFace: F_MED, color: WHITE, align: 'center', lineSpacingMultiple: 1.5 });
    body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit sed do euismod tempor.',
      { x: c.bx, y: i === 2 ? 5.081 : 5.073, w: 2.714, h: 1.159, color: WHITE, align: 'center' });
  });
  footer(s, WHITE);
});

/* 30 - closing ----------------------------------------------------------- */
SLIDES.push(function (s) {
  box(s, 0, 0, SLIDE_W, SLIDE_H, BLACK, { fill: { color: BLACK, transparency: 60 } });
  text(s, [
    { text: 'CLOSING', options: { fontSize: 72, fontFace: F_SEMI, breakLine: true } },
    { text: 'PRESENTATION', options: { fontSize: 48, fontFace: F_MED } },
  ], { x: 2.473, y: 2.211, w: 8.387, h: 2.121, align: 'center', color: WHITE, charSpacing: 8 });
  text(s, LOREM_MED, { x: 3.544, y: 4.751, w: 6.245, h: 0.627, fontSize: 11, color: WHITE, align: 'center', lineSpacingMultiple: 1.5 });
  footer(s, WHITE);
  text(s, '225 Beeghley Street, Glen Ellyn\nIllnois, 60138',
    { x: 10.248, y: 6.332, w: 2.58, h: 0.789, fontSize: 11, color: WHITE, align: 'right', lineSpacingMultiple: 2 });
});

/* 31 / 32 - map libraries ------------------------------------------------ */
/* The source slides are galleries of grey vector maps; each is stood in for
   by a soft grey blob at the same position and size.                        */
const MAPS_31 = [
  [1.458, 1.049, 2.216, 2.492], [4.564, 1.193, 3.502, 2.205], [9.181, 1.224, 3.028, 2.143],
  [1.568, 3.891, 1.997, 2.789], [4.936, 4.222, 2.758, 2.127], [9.338, 4.098, 2.712, 2.374],
];
const MAPS_32 = [
  [0.964, 1.394, 3.088, 2.057], [4.452, 1.116, 1.914, 2.614], [6.864, 0.847, 2.172, 3.152], [9.521, 1.283, 2.972, 2.280],
  [0.928, 4.501, 3.160, 1.658], [4.513, 4.304, 1.793, 2.051], [6.722, 3.979, 2.455, 2.700], [9.438, 4.098, 3.139, 2.462],
];
function mapSlide(list) {
  return function (s) {
    list.forEach(function (m) {
      shape(s, 'roundRect', { x: m[0], y: m[1], w: m[2], h: m[3], rectRadius: 0.08, fill: { color: BAND } });
      text(s, '[image]', {
        x: m[0], y: m[1], w: m[2], h: m[3], margin: 0,
        fontSize: 11, color: WHITE, align: 'center', valign: 'middle',
      });
    });
  };
}
SLIDES.push(mapSlide(MAPS_31));
SLIDES.push(mapSlide(MAPS_32));

/* ------------------------------------------------------------------- main */

function build() {
  const pptx = new pptxgen();
  pptx.defineLayout({ name: 'CUSTOM_16x9', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'CUSTOM_16x9';
  pptx.theme = { headFontFace: 'Montserrat', bodyFontFace: F_BODY };
  pptx.title = 'Financial Planner';

  SLIDES.forEach(function (builder) {
    const slide = pptx.addSlide();
    slide.background = { color: WHITE };
    builder(slide);
  });

  return pptx.writeFile({ fileName: path.join(__dirname, '0b218cca-34ae-4a3c-b3cd-34f8fba993aa_grok_final.pptx') });
}

build().then(function (f) { console.log('wrote', f); }, function (e) { console.error(e); process.exit(1); });
