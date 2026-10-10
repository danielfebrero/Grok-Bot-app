/**
 * "Medical Healthcare Monitoring" deck rebuilt with pptxgenjs.
 *
 * Run:  node 05770343-e0f0-4949-a85c-39a8841415a1_grok_final.js
 * Out:  05770343-e0f0-4949-a85c-39a8841415a1_grok_final.pptx (next to this file)
 *
 * Photographs in the source deck are redrawn as flat grey placeholder shapes.
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Design tokens
 * ------------------------------------------------------------------ */

const NAVY = '00529B';
const BLUE = '4F9BCE';
const GREY = 'A5A5A5';
const GREY_MID = '7B7B7B';
const GREY_PALE = 'EDEDED';
const WHITE = 'FFFFFF';
const PAPER = 'F2F2F2';
const PHOTO = '7F7F7F'; // stand-in for every bitmap in the original deck
const DEVICE = '1A1A1A';
const DEVICE_BASE = 'C9CCCF';

const SANS = 'Lato';
const SANS_LIGHT = 'Lato Light';

/** text-frame insets of the source deck, in points: [left, right, bottom, top] */
const INSET = [7.2, 7.2, 3.6, 3.6];

/**
 * The deck's stock drop shadow (blur 40pt / offset 18pt / 45 deg).
 * pptxgenjs rewrites the shadow object it is handed, so hand out a fresh one every call.
 */
const shadow = (opacity = 0.16) => ({
  type: 'outer', color: '000000', opacity, blur: 40, offset: 18, angle: 45,
});
/** softer straight-down shadow used by the pricing cards */
const dropShadow = (opacity, blur, offset) => ({
  type: 'outer', color: '000000', opacity, blur, offset, angle: 90,
});

/* ------------------------------------------------------------------ *
 * Small drawing helpers
 * ------------------------------------------------------------------ */

/** Text box, top anchored by default, with the deck's insets. */
function text(slide, body, opts) {
  slide.addText(body, Object.assign({ fontFace: SANS, valign: 'top', margin: INSET }, opts));
}

/** 12pt light grey kicker that sits above every section title. */
function kicker(slide, x, y, w, opts) {
  text(slide, 'Open research issues in', Object.assign(
    { x, y, w, h: 0.303, fontFace: SANS_LIGHT, fontSize: 12, color: GREY }, opts));
}

/** 40pt light section title. */
function title(slide, body, x, y, w, h, opts) {
  text(slide, body, Object.assign(
    { x, y, w, h, fontFace: SANS_LIGHT, fontSize: 40, color: NAVY }, opts));
}

/** 14pt grey running copy at 150% leading. */
function copy(slide, body, x, y, w, h, opts) {
  text(slide, body, Object.assign(
    { x, y, w, h, fontSize: 14, color: GREY, lineSpacingMultiple: 1.5 }, opts));
}

/** Grey block standing in for a photograph. */
function photo(slide, x, y, w, h, radius) {
  slide.addShape(radius ? 'roundRect' : 'rect', {
    x, y, w, h, fill: { color: PHOTO }, rectRadius: radius || undefined,
  });
}

/** Bar whose two bottom corners are rounded (round2SameRect, flipped). */
function bottomRoundedBar(slide, x, y, w, h, radius, color) {
  slide.addShape('roundRect', { x, y, w, h, rectRadius: radius, fill: { color } });
  slide.addShape('rect', { x, y, w, h: radius, fill: { color } });
}

/** Greek-cross watermark (a "mathPlus" with the deck's slim 15% arms). */
function crossMark(slide, x, y, size, fill) {
  const t = size * 0.15064;
  const a = (size - t) / 2;
  const b = a + t;
  slide.addShape('custGeom', {
    x, y, w: size, h: size, fill,
    points: [
      { x: a, y: 0 }, { x: b, y: 0 }, { x: b, y: a }, { x: size, y: a },
      { x: size, y: b }, { x: b, y: b }, { x: b, y: size }, { x: a, y: size },
      { x: a, y: b }, { x: 0, y: b }, { x: 0, y: a }, { x: a, y: a }, { close: true },
    ],
  });
}

/** Bulleted paragraphs (round bullet by default). */
function bulletList(items, opts, bullet) {
  return items.map((t) => ({
    text: t,
    options: Object.assign({ bullet: bullet || { characterCode: '2022', indent: 22.5 } }, opts),
  }));
}

/* ------------------------------------------------------------------ *
 * Pictogram library - each icon is a stack of native shapes placed in
 * a 0..1 box; tone 0 paints the icon colour, tone 1 punches the
 * background colour back out.
 * ------------------------------------------------------------------ */

const ICONS = {
  // doctor's case with a cross
  bag: [
    ['rect', 0.32, 0.00, 0.36, 0.12, 0],
    ['rect', 0.39, 0.05, 0.22, 0.09, 1],
    ['roundRect', 0.00, 0.20, 1.00, 0.80, 0, { rectRadius: 0.05 }],
    ['mathPlus', 0.30, 0.38, 0.40, 0.42, 1],
  ],
  // clipboard with ruled lines and a capsule
  clipboard: [
    ['roundRect', 0.02, 0.08, 0.80, 0.92, 0, { rectRadius: 0.05 }],
    ['roundRect', 0.10, 0.16, 0.64, 0.76, 1, { rectRadius: 0.03 }],
    ['ellipse', 0.34, 0.00, 0.16, 0.16, 0],
    ['rect', 0.26, 0.06, 0.32, 0.12, 0],
    ['rect', 0.18, 0.28, 0.48, 0.07, 0],
    ['rect', 0.18, 0.42, 0.48, 0.07, 0],
    ['rect', 0.18, 0.56, 0.30, 0.07, 0],
    ['roundRect', 0.52, 0.56, 0.48, 0.30, 0, { rectRadius: 0.15, rotate: -45 }],
    ['roundRect', 0.58, 0.62, 0.36, 0.18, 1, { rectRadius: 0.09, rotate: -45 }],
  ],
  // heart with a pulse notch
  heart: [
    ['heart', 0.00, 0.05, 1.00, 0.90, 0],
    ['rect', 0.12, 0.42, 0.76, 0.07, 1],
    ['rect', 0.36, 0.28, 0.08, 0.20, 1],
    ['rect', 0.58, 0.42, 0.08, 0.22, 1],
  ],
  // surgical mask with pleats and ear loops (drawn wide, 0..1 in x is landscape)
  mask: [
    ['donut', 0.00, 0.10, 0.22, 0.80, 0],
    ['donut', 0.78, 0.10, 0.22, 0.80, 0],
    ['roundRect', 0.14, 0.00, 0.72, 1.00, 0, { rectRadius: 0.16 }],
    ['roundRect', 0.24, 0.10, 0.52, 0.80, 1, { rectRadius: 0.10 }],
    ['roundRect', 0.30, 0.20, 0.40, 0.15, 0, { rectRadius: 0.075 }],
    ['roundRect', 0.30, 0.43, 0.40, 0.15, 0, { rectRadius: 0.075 }],
    ['roundRect', 0.30, 0.66, 0.40, 0.15, 0, { rectRadius: 0.075 }],
  ],
  // molar
  tooth: [
    ['ellipse', 0.00, 0.00, 1.00, 0.52, 0],
    ['roundRect', 0.04, 0.18, 0.42, 0.82, 0, { rectRadius: 0.14 }],
    ['roundRect', 0.54, 0.18, 0.42, 0.82, 0, { rectRadius: 0.14 }],
    ['rect', 0.42, 0.46, 0.16, 0.54, 1],
  ],
  // syringe next to a medicine bottle
  bottle: [
    ['rect', 0.06, 0.00, 0.34, 0.07, 0],
    ['rect', 0.17, 0.05, 0.12, 0.14, 0],
    ['roundRect', 0.08, 0.18, 0.30, 0.82, 0, { rectRadius: 0.05 }],
    ['rect', 0.14, 0.34, 0.18, 0.05, 1],
    ['rect', 0.14, 0.48, 0.18, 0.05, 1],
    ['rect', 0.14, 0.62, 0.18, 0.05, 1],
    ['rect', 0.56, 0.00, 0.20, 0.10, 0],
    ['roundRect', 0.48, 0.10, 0.38, 0.90, 0, { rectRadius: 0.09 }],
    ['ellipse', 0.56, 0.36, 0.22, 0.22, 1],
  ],
  // hand catching a droplet
  handDrop: [
    ['teardrop', 0.02, 0.02, 0.30, 0.36, 0, { rotate: 225 }],
    ['roundRect', 0.36, 0.34, 0.56, 0.66, 0, { rectRadius: 0.14 }],
    ['roundRect', 0.42, 0.16, 0.11, 0.36, 0, { rectRadius: 0.055 }],
    ['roundRect', 0.58, 0.12, 0.11, 0.40, 0, { rectRadius: 0.055 }],
    ['roundRect', 0.74, 0.18, 0.11, 0.34, 0, { rectRadius: 0.055 }],
    ['roundRect', 0.24, 0.52, 0.22, 0.40, 0, { rectRadius: 0.10, rotate: 35 }],
    ['rect', 0.55, 0.54, 0.05, 0.46, 1],
    ['rect', 0.71, 0.54, 0.05, 0.46, 1],
  ],
  // open palm with separated fingers
  hand: [
    ['roundRect', 0.10, 0.46, 0.76, 0.54, 0, { rectRadius: 0.16 }],
    ['roundRect', 0.14, 0.10, 0.15, 0.36, 0, { rectRadius: 0.075 }],
    ['roundRect', 0.33, 0.02, 0.15, 0.44, 0, { rectRadius: 0.075 }],
    ['roundRect', 0.52, 0.00, 0.15, 0.46, 0, { rectRadius: 0.075 }],
    ['roundRect', 0.71, 0.06, 0.15, 0.40, 0, { rectRadius: 0.075 }],
    ['roundRect', 0.00, 0.34, 0.14, 0.34, 0, { rectRadius: 0.07, rotate: 25 }],
    ['rect', 0.28, 0.74, 0.09, 0.08, 1],
    ['rect', 0.44, 0.74, 0.09, 0.08, 1],
    ['rect', 0.60, 0.74, 0.09, 0.08, 1],
  ],
  // knee joint with motion marks
  joint: [
    ['roundRect', 0.04, 0.00, 0.13, 0.42, 0, { rectRadius: 0.06 }],
    ['roundRect', 0.34, 0.00, 0.13, 0.42, 0, { rectRadius: 0.06 }],
    ['roundRect', 0.04, 0.58, 0.13, 0.42, 0, { rectRadius: 0.06 }],
    ['roundRect', 0.34, 0.58, 0.13, 0.42, 0, { rectRadius: 0.06 }],
    ['roundRect', 0.02, 0.34, 0.47, 0.32, 0, { rectRadius: 0.14 }],
    ['roundRect', 0.13, 0.42, 0.25, 0.16, 1, { rectRadius: 0.07 }],
    ['rect', 0.62, 0.16, 0.34, 0.10, 0, { rotate: -20 }],
    ['rect', 0.64, 0.45, 0.34, 0.10, 0],
    ['rect', 0.62, 0.74, 0.34, 0.10, 0, { rotate: 20 }],
  ],
  // lungs either side of a trachea
  lungs: [
    ['rect', 0.45, 0.00, 0.10, 0.44, 0],
    ['ellipse', 0.00, 0.22, 0.48, 0.78, 0],
    ['ellipse', 0.52, 0.22, 0.48, 0.78, 0],
    ['triangle', 0.30, 0.30, 0.40, 0.70, 1],
    ['rect', 0.12, 0.40, 0.76, 0.05, 1],
    ['rect', 0.08, 0.58, 0.84, 0.05, 1],
    ['rect', 0.12, 0.76, 0.76, 0.05, 1],
    ['rect', 0.45, 0.34, 0.10, 0.66, 1],
    ['rect', 0.45, 0.34, 0.10, 0.16, 0],
  ],
  // heart monitor trace inside a rounded frame
  ecg: [
    ['roundRect', 0.00, 0.00, 1.00, 1.00, 0, { rectRadius: 0.12 }],
    ['roundRect', 0.10, 0.14, 0.80, 0.72, 1, { rectRadius: 0.06 }],
    ['rect', 0.00, 0.44, 0.34, 0.12, 0],
    ['rect', 0.66, 0.44, 0.34, 0.12, 0],
    ['rect', 0.32, 0.30, 0.10, 0.40, 0, { rotate: 18 }],
    ['rect', 0.47, 0.24, 0.10, 0.52, 0, { rotate: -16 }],
    ['rect', 0.60, 0.36, 0.10, 0.28, 0, { rotate: 18 }],
  ],
  // globe with a signal badge
  globe: [
    ['donut', 0.00, 0.10, 0.90, 0.90, 0, { rectRadius: 0.10 }],
    ['rect', 0.06, 0.32, 0.78, 0.10, 0],
    ['rect', 0.06, 0.66, 0.78, 0.10, 0],
    ['ellipse', 0.28, 0.10, 0.34, 0.90, 0],
    ['ellipse', 0.34, 0.16, 0.22, 0.78, 1],
    ['rect', 0.72, 0.00, 0.28, 0.24, 0],
  ],
  // envelope
  mail: [
    ['rect', 0.00, 0.00, 1.00, 1.00, 0],
    ['triangle', 0.02, 0.02, 0.96, 0.66, 1, { flipV: true }],
    ['triangle', 0.10, 0.10, 0.80, 0.50, 0, { flipV: true }],
  ],
  // telephone handset
  phone: [
    ['blockArc', -0.28, -0.28, 1.56, 1.56, 0, { angleRange: [135, 30], arcThicknessRatio: 0.42 }],
  ],
  // house with a heart
  home: [
    ['triangle', 0.00, 0.00, 1.00, 0.54, 0],
    ['rect', 0.10, 0.46, 0.80, 0.54, 0],
    ['rect', 0.20, 0.56, 0.60, 0.44, 1],
    ['heart', 0.32, 0.60, 0.36, 0.34, 0],
  ],
  // tick
  check: [
    ['rect', 0.02, 0.42, 0.42, 0.16, 0, { rotate: 45 }],
    ['rect', 0.30, 0.30, 0.72, 0.16, 0, { rotate: -50 }],
  ],
};

function icon(slide, name, x, y, w, h, fg, bg) {
  ICONS[name].forEach(([shape, rx, ry, rw, rh, tone, extra]) => {
    slide.addShape(shape, Object.assign({
      x: x + rx * w, y: y + ry * h, w: rw * w, h: rh * h,
      fill: { color: tone ? bg : fg },
    }, extra || {}));
  });
}

/* ------------------------------------------------------------------ *
 * Backgrounds borrowed from the slide layouts
 * ------------------------------------------------------------------ */

/** Layout 1 / 15: navy field with two faint cross watermarks. */
function navyCrossBackdrop(slide, crosses) {
  slide.background = { color: NAVY };
  crosses.forEach(([x, y]) => crossMark(slide, x, y, 5.551, { color: BLUE, transparency: 94 }));
}

/** Layout 2: pale ribbon "X" bleeding off the left edge. */
function ribbonBackdrop(slide) {
  slide.addShape('custGeom', {
    x: -0.026, y: 0, w: 5.06, h: 7.5, fill: { color: BLUE, transparency: 96 },
    points: [
      { x: 4.006, y: 0.000 }, { x: 4.976, y: 0.000 }, { x: 2.486, y: 3.600 },
      { x: 5.060, y: 7.500 }, { x: 4.054, y: 7.500 }, { x: 3.766, y: 7.306 },
      { x: 1.752, y: 4.161 }, { x: 0.000, y: 6.854 }, { x: 0.000, y: 5.229 },
      { x: 1.061, y: 3.648 }, { x: 0.026, y: 2.125 }, { x: 0.026, y: 0.411 },
      { x: 1.810, y: 3.145 }, { x: 3.771, y: 0.162 }, { close: true },
    ],
  });
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

function slide01(pptx) {
  const s = pptx.addSlide();
  navyCrossBackdrop(s, [[-0.748, -0.782], [8.551, 2.767]]);

  s.addShape('ellipse', { x: 6.214, y: 1.777, w: 0.906, h: 0.906, fill: { color: BLUE } });
  icon(s, 'bag', 6.444, 2.007, 0.446, 0.446, WHITE, BLUE);

  title(s, [
    { text: 'Medical', options: { fontFace: SANS } },
    { text: ' Healthcare Monitoring with Wearable and Implantable ', options: { fontFace: SANS_LIGHT } },
    { text: 'Sensors', options: { fontFace: SANS } },
  ], 2.102, 3.037, 9.129, 1.447, { color: WHITE, align: 'center' });

  text(s, 'Open research issues in this area are predominantly related', {
    x: 2.575, y: 4.651, w: 8.184, h: 0.337,
    fontFace: SANS_LIGHT, fontSize: 14, color: PAPER, align: 'center',
  });
}

function slide02(pptx) {
  const s = pptx.addSlide();
  ribbonBackdrop(s);
  photo(s, 8.312, 0, 5.022, 6.071);

  kicker(s, 1.031, 1.009, 4.806);
  title(s, 'Medical Healthcare Monitoring with Wearable and Implantable', 1.031, 1.37, 6.963, 2.121);
  copy(s, 'PLACEHOLDER'
    + 'concentrate more intensively in order to contribute to the necessary improvement of adequate.',
    1.031, 3.72, 6.224, 1.124);
  text(s, bulletList([
    'Approaches should concentrate more intensively in order',
    'to contribute to the necessary improvement of adequate.',
  ], { fontSize: 14, color: GREY, lineSpacingMultiple: 1.5 },
  { type: 'number', style: 'arabicPeriod', indent: 27 }),
  { x: 1.031, y: 5.034, w: 5.888, h: 0.761 });
}

function slide03(pptx) {
  const s = pptx.addSlide();

  text(s, 'Open research issues in this area are predominantly', {
    x: 3.664, y: 0.957, w: 6.004, h: 0.303,
    fontFace: SANS_LIGHT, fontSize: 12, color: GREY, align: 'center',
  });
  title(s, 'Intrathecal Gene Therapy for Giant Axonal Neuropathy',
    2.605, 1.305, 8.122, 1.447, { align: 'center' });
  copy(s, 'PLACEHOLDER'
    + 'more intensively in order to contribute to the necessary improvement of adequate. Approaches '
    + 'PLACEHOLDER',
    2.151, 2.882, 9.031, 1.115, { align: 'center' });

  const quote = '\u201CIt is concluded that epidemiological and anthropological approaches '
    + 'should concentrate more intensively';
  [[0.333, GREY, 0.717, 5.057], [4.672, GREY_MID, 5.055, 5.031], [9.010, NAVY, 9.394, 5.031]]
    .forEach(([x, fill, tx, ty]) => {
      s.addShape('round1Rect', {
        x, y: 4.774, w: 3.99, h: 1.681, rectRadius: 0.28, fill: { color: fill },
      });
      copy(s, quote, tx, ty, 3.402, 1.115, { color: PAPER });
    });
}

function slide04(pptx) {
  const s = pptx.addSlide();
  photo(s, 0, 0, 3.808, 7.5);

  kicker(s, 5.265, 1.016, 4.806);
  title(s, 'Mental Health Services into Primary Care', 5.224, 1.377, 6.612, 1.447);
  copy(s, 'PLACEHOLDER'
    + 'concentrate more intensively in order to contribute', 5.255, 2.882, 6.224, 0.761);

  s.addShape('roundRect', {
    x: 5.378, y: 4.053, w: 6.612, h: 2.336, rectRadius: 0.2485, fill: { color: NAVY },
  });
  s.addShape('ellipse', {
    x: 6.087, y: 4.494, w: 0.322, h: 0.322, fill: { color: BLUE },
    shadow: { type: 'outer', color: '000000', opacity: 0.12, blur: 29, offset: 3, angle: 90 },
  });
  icon(s, 'check', 6.163, 4.588, 0.17, 0.134, WHITE, BLUE);

  text(s, 'Primary Care Health', {
    x: 6.648, y: 4.416, w: 2.73, h: 0.438, fontSize: 20, color: WHITE,
  });
  copy(s, [
    { text: 'Approaches should ', options: { bold: true } },
    { text: 'PLACEHOLDER'
      + 'epidemiological and anthropological approaches' },
  ], 6.648, 4.853, 4.79, 1.115, { color: PAPER });
}

function slide05(pptx) {
  const s = pptx.addSlide();
  s.background = { color: NAVY };
  photo(s, 0, 0, 13.333, 7.5);

  text(s, 'Open research issues in', {
    x: 1.551, y: 1.434, w: 4.806, h: 0.303, fontFace: SANS_LIGHT, fontSize: 12, color: PAPER,
  });
  title(s, 'Telepsychiatry in Mental Healthcare Delivery',
    1.551, 1.794, 6.612, 1.447, { color: WHITE });

  s.addShape('roundRect', {
    x: 1.643, y: 3.819, w: 3.857, h: 2.233, rectRadius: 0.2383,
    fill: { color: WHITE }, shadow: shadow(0.05),
  });
  text(s, '6,000+', {
    x: 2.573, y: 4.399, w: 1.997, h: 0.707, fontSize: 36, color: NAVY, align: 'center',
  });
  text(s, 'Health Presentations', {
    x: 2.21, y: 5.067, w: 2.723, h: 0.37, fontSize: 16, color: GREY, align: 'center',
  });
  copy(s, 'PLACEHOLDER'
    + 'more intensively in order to contribute to the necessary improvement of adequate. Approaches '
    + 'PLACEHOLDER',
    6.538, 3.923, 5.757, 1.822, { color: PAPER });
}

function slide06(pptx) {
  const s = pptx.addSlide();

  // photograph bleeding off the right edge behind a big circular cut
  s.addShape('custGeom', {
    x: 7.437, y: 0, w: 5.897, h: 7.5, fill: { color: PHOTO },
    points: [
      { x: 1.952, y: 0.000 }, { x: 5.897, y: 0.000 }, { x: 5.897, y: 7.500 },
      { x: 1.510, y: 7.500 }, { x: 1.451, y: 7.444 },
      { x: 0.006, y: 4.195, curve: { type: 'cubic', x1: 0.611, y1: 6.603, x2: 0.071, y2: 5.462 } },
      { x: 0.000, y: 3.940 }, { x: 0.008, y: 3.653 },
      { x: 1.803, y: 0.117, curve: { type: 'cubic', x1: 0.089, y1: 2.231, x2: 0.771, y2: 0.969 } },
      { close: true },
    ],
  });

  kicker(s, 1.071, 0.993, 4.806);
  title(s, 'Health Information Technology on Clinical Outcomes', 1.031, 1.354, 6.612, 2.121);

  const stats = [
    { x: 1.912, fill: BLUE, glyph: 'bag', ix: 2.153, iy: 4.267, iw: 0.497, ih: 0.497,
      value: '200+', vx: 1.829, cx: 1.814 },
    { x: 4.403, fill: NAVY, glyph: 'clipboard', ix: 4.644, iy: 4.242, iw: 0.497, ih: 0.547,
      value: '300+', vx: 4.32, cx: 4.32 },
  ];
  stats.forEach((st) => {
    s.addShape('ellipse', {
      x: st.x, y: 4.026, w: 0.979, h: 0.979, fill: { color: st.fill }, shadow: shadow(),
    });
    icon(s, st.glyph, st.ix, st.iy, st.iw, st.ih, WHITE, st.fill);
    text(s, st.value, {
      x: st.vx, y: 5.201, w: 1.146, h: 0.572, fontSize: 28, color: NAVY, align: 'center',
    });
    copy(s, 'epidemiological and anthropological', st.cx, 5.766, 2.347, 0.761);
  });
}

function slide07(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 4.51, w: 13.333, h: 2.99, fill: { color: NAVY } });

  kicker(s, 4.264, 0.832, 4.806, { align: 'center' });
  title(s, 'Medical Health Clinical Team', 3.087, 1.193, 7.16, 0.774, { align: 'center' });

  const team = [
    { px: 1.100, nx: 1.224, rx: 1.152, name: 'Genevieve Greta', role: 'Antibiotik nurse' },
    { px: 3.969, nx: 4.041, rx: 3.969, name: 'Hermioner', role: 'Specialist health' },
    { px: 6.838, nx: 7.002, rx: 6.930, name: 'Lucas Eveira', role: 'Antibiotik nurse' },
    { px: 9.706, nx: 9.818, rx: 9.746, name: 'Maximilian', role: 'Specialist health' },
  ];
  team.forEach((m) => {
    photo(s, m.px, 2.511, 2.527, 2.966, 0.1718);
    text(s, m.name, {
      x: m.nx, y: 5.839, w: 2.279, h: 0.438, fontSize: 20, color: WHITE, align: 'center',
    });
    text(s, m.role, {
      x: m.rx, y: 6.277, w: 2.423, h: 0.337,
      fontFace: SANS_LIGHT, fontSize: 14, color: PAPER, align: 'center',
    });
  });
}

function slide08(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: 12.092, h: 4.434, fill: { color: NAVY } });
  photo(s, 9.172, 0, 2.92, 4.434);

  text(s, 'Open research issues in', {
    x: 1.49, y: 0.774, w: 4.806, h: 0.303, fontFace: SANS_LIGHT, fontSize: 12, color: PAPER,
  });
  title(s, 'Manager Clinic & Directur', 1.449, 1.134, 6.612, 0.774, { color: WHITE });
  copy(s, 'PLACEHOLDER'
    + 'concentrate more intensively in order to contribute.', 1.474, 1.922, 6.224, 0.761,
  { color: PAPER });
  copy(s, 'PLACEHOLDER',
    1.49, 2.909, 5.102, 0.761, { color: PAPER, bold: true });
  title(s, '\u201CMedical Healthcare Monitoring with Wearable and Implantable\u201D',
    1.98, 5.108, 9.224, 1.447);
}

function slide09(pptx) {
  const s = pptx.addSlide();

  kicker(s, 1.227, 0.717, 4.461);
  title(s, 'Health Information Technology on', 1.218, 1.078, 5.398, 1.447);

  s.addShape('roundRect', {
    x: 1.327, y: 2.753, w: 2.216, h: 0.529, rectRadius: 0.2645,
    fill: { color: NAVY }, shadow: shadow(),
  });
  text(s, 'Health Outcomes', {
    x: 1.327, y: 2.753, w: 2.216, h: 0.529,
    fontSize: 16, color: WHITE, align: 'center', valign: 'middle',
  });

  [[6.985, 0.589, 5.501, 0.2065], [6.985, 3.867, 5.501, 0.2065],
    [3.891, 3.867, 2.778, 0.1889], [0.850, 3.867, 2.724, 0.1852]]
    .forEach(([x, y, w, r]) => photo(s, x, y, w, 3.037, r));
}

function slide10(pptx) {
  const s = pptx.addSlide();

  [0, 4.47, 8.939].forEach((x) => s.addShape('rect', {
    x, y: 4.686, w: 4.392, h: 2.426, fill: { color: WHITE }, shadow: shadow(0.15),
  }));

  kicker(s, 4.264, 0.52, 4.806, { align: 'center' });
  title(s, 'Gallery Medical Health', 3.087, 0.88, 7.16, 0.774, { align: 'center' });

  [[0, 0.793, 0.984, 0.476], [4.471, 5.262, 5.454, 4.946], [8.942, 9.731, 9.923, 9.415]]
    .forEach(([px, hx, sx, bx]) => {
      photo(s, px, 2.048, 4.392, 2.638);
      text(s, 'Maternal Healthcare', {
        x: hx, y: 5.08, w: 2.806, h: 0.438, fontSize: 20, color: NAVY, align: 'center',
      });
      text(s, 'Literacy and Health', {
        x: sx, y: 5.51, w: 2.423, h: 0.337,
        fontFace: SANS_LIGHT, fontSize: 14, color: GREY, align: 'center',
      });
      copy(s, 'PLACEHOLDER',
        bx, 5.904, 3.439, 0.761, { align: 'center' });
    });
}

function slide11(pptx) {
  const s = pptx.addSlide();

  kicker(s, 4.264, 0.652, 4.806, { align: 'center' });
  title(s, 'Mental Health Infographic', 2.432, 1.013, 8.469, 0.774, { align: 'center' });

  [{ x: 1.932, hx: 2.883, hw: 2.73, tx: 2.406, label: 'Primary Care Health' },
    { x: 6.769, hx: 7.494, hw: 3.182, tx: 7.242, label: 'Secondary Care Health' }]
    .forEach((card) => {
      s.addShape('rect', {
        x: card.x, y: 2.316, w: 4.633, h: 3.204, fill: { color: GREY_PALE },
      });
      bottomRoundedBar(s, card.x, 2.316, 4.633, 0.781, 0.13, NAVY);
      text(s, card.label, {
        x: card.hx, y: 2.487, w: card.hw, h: 0.438, fontSize: 20, color: WHITE,
      });
      text(s, bulletList(['PLACEHOLDER'
        + 'should concentrate more intensively in order to contribute.'],
      { fontSize: 14, color: GREY, lineSpacingMultiple: 1.5 }),
      { x: card.tx, y: 3.513, w: 3.781, h: 1.468 });
    });

  copy(s, 'PLACEHOLDER'
    + 'more intensively in order to contribute to the necessary improvement of adequate.',
    2.577, 5.957, 8.179, 0.761, { align: 'center' });
}

function slide12(pptx) {
  const s = pptx.addSlide();

  kicker(s, 4.264, 0.673, 4.806, { align: 'center' });
  title(s, 'Mental Health Infographic', 2.432, 1.033, 8.469, 0.774, { align: 'center' });

  s.addShape('donut', {
    x: 1.894, y: 2.606, w: 3.814, h: 3.814, rectRadius: 0.072, fill: { color: GREY },
  });

  const ring = [
    { cx: 1.763, cy: 3.292, glyph: 'lungs', ix: 1.938, iy: 3.426, iw: 0.482, ih: 0.482 },
    { cx: 3.386, cy: 2.260, glyph: 'hand', ix: 3.572, iy: 2.421, iw: 0.459, ih: 0.510 },
    { cx: 4.877, cy: 3.292, glyph: 'joint', ix: 5.080, iy: 3.466, iw: 0.465, ih: 0.482 },
    { cx: 4.943, cy: 5.012, glyph: 'bottle', ix: 5.116, iy: 5.167, iw: 0.485, ih: 0.523 },
    { cx: 3.320, cy: 5.935, glyph: 'tooth', ix: 3.526, iy: 6.083, iw: 0.418, ih: 0.523 },
    { cx: 1.829, cy: 5.012, glyph: 'ecg', ix: 2.002, iy: 5.234, iw: 0.485, ih: 0.388 },
  ];
  ring.forEach((node) => {
    s.addShape('ellipse', {
      x: node.cx, y: node.cy, w: 0.831, h: 0.831, fill: { color: NAVY }, shadow: shadow(),
    });
    icon(s, node.glyph, node.ix, node.iy, node.iw, node.ih, WHITE, NAVY);
  });

  text(s, '5,000+', {
    x: 2.554, y: 3.994, w: 2.494, h: 0.707, fontSize: 36, color: NAVY, align: 'center',
  });
  copy(s, 'anthropological', 2.789, 4.585, 2.012, 0.408, { align: 'center' });

  text(s, bulletList([
    'PLACEHOLDER',
    'more intensively in order to contribute to the necessary improvement of adequate. Approaches',
    'PLACEHOLDER',
  ], { fontSize: 14, color: GREY, lineSpacingMultiple: 1.5 }),
  { x: 7.151, y: 3.395, w: 4.827, h: 2.175 });
}

function slide13(pptx) {
  const s = pptx.addSlide();

  kicker(s, 4.264, 0.708, 4.806, { align: 'center' });
  title(s, 'Health Information Infographic', 2.784, 1.027, 7.765, 0.774, { align: 'center' });

  const pills = [
    { x: 1.721, y: 2.469, fill: NAVY, glyph: 'mask', ix: 2.339, iy: 2.831, iw: 0.685, ih: 0.436 },
    { x: 6.948, y: 2.469, fill: GREY, glyph: 'tooth', ix: 7.689, iy: 2.831, iw: 0.435, ih: 0.436 },
    { x: 1.721, y: 3.968, fill: GREY, glyph: 'clipboard', ix: 2.462, iy: 4.305, iw: 0.435, ih: 0.484 },
    { x: 6.948, y: 3.968, fill: NAVY, glyph: 'handDrop', ix: 7.595, iy: 4.305, iw: 0.512, ih: 0.484 },
    { x: 1.721, y: 5.466, fill: NAVY, glyph: 'bottle', ix: 2.521, iy: 5.727, iw: 0.318, ih: 0.636 },
    { x: 6.948, y: 5.466, fill: NAVY, glyph: 'heart', ix: 7.662, iy: 5.827, iw: 0.484, ih: 0.436 },
  ];
  pills.forEach((p) => {
    s.addShape('roundRect', {
      x: p.x, y: p.y, w: 4.665, h: 1.158, rectRadius: 0.579,
      fill: { color: p.fill }, shadow: shadow(),
    });
    icon(s, p.glyph, p.ix, p.iy, p.iw, p.ih, WHITE, p.fill);
    text(s, 'Dental Service Health Infographic', {
      x: p.x + 1.466, y: p.y + 0.192, w: 2.635, h: 0.774, fontSize: 20, color: WHITE,
    });
  });
}

function slide14(pptx) {
  const s = pptx.addSlide();

  kicker(s, 4.264, 0.796, 4.806, { align: 'center' });
  title(s, 'Medical SWOT Infographic', 3.361, 1.156, 6.612, 0.774, { align: 'center' });

  const columns = [
    { x: 0.269, fill: NAVY, letter: 'S', label: 'Health Strengths',
      body: 'PLACEHOLDER' },
    { x: 3.501, fill: BLUE, letter: 'W', label: 'Health Weakness',
      body: 'contribute to the necessary improvement of adequate. Approaches should' },
    { x: 6.739, fill: NAVY, letter: 'O', label: 'Health Opportunity',
      body: 'PLACEHOLDER' },
    { x: 9.976, fill: NAVY, letter: 'T', label: 'Health Threats',
      body: 'contribute to the necessary improvement of adequate. Approaches should' },
  ];
  columns.forEach((c) => {
    s.addShape('roundRect', {
      x: c.x, y: 2.581, w: 3.088, h: 4.123, rectRadius: 0.295, fill: { color: c.fill },
    });
    text(s, c.letter, {
      x: c.x + 0.179, y: 2.638, w: 2.73, h: 2.036,
      fontSize: 115, color: WHITE, align: 'center',
    });
    text(s, c.label, {
      x: c.x + 0.226, y: 4.678, w: 2.635, h: 0.438, fontSize: 20, color: WHITE, align: 'center',
    });
    copy(s, c.body, c.x + 0.163, 5.208, 2.762, 1.115, { color: PAPER, align: 'center' });
  });
}

function slide15(pptx) {
  const s = pptx.addSlide();
  photo(s, 8.0, 1.368, 4.02, 6.132);

  kicker(s, 1.111, 1.368, 4.806);
  title(s, 'Telepsychiatry in Mental Healthcare Delivery', 1.07, 1.729, 6.612, 1.447);
  text(s, bulletList([
    'PLACEHOLDER',
    'more intensively in order to contribute to the necessary improvement of adequate. Approaches',
    'PLACEHOLDER',
  ], { fontSize: 14, color: GREY, lineSpacingMultiple: 1.5 }),
  { x: 1.07, y: 3.919, w: 5.837, h: 2.175 });
}

function slide16(pptx) {
  const s = pptx.addSlide();
  s.addShape('ellipse', { x: 0.98, y: 0.96, w: 5.193, h: 5.193, fill: { color: GREY_PALE } });

  // phone mock-up: white shell, dark bezel, grey screen
  s.addShape('roundRect', {
    x: 2.147, y: 1.867, w: 2.86, h: 5.734, rectRadius: 0.34,
    fill: { color: WHITE }, line: { color: '3C3C3C', width: 0.75 },
  });
  s.addShape('roundRect', {
    x: 2.213, y: 1.925, w: 2.734, h: 5.62, rectRadius: 0.30, fill: { color: DEVICE },
  });
  s.addShape('roundRect', {
    x: 2.294, y: 2.01, w: 2.574, h: 5.455, rectRadius: 0.24, fill: { color: PHOTO },
  });

  kicker(s, 7.371, 1.68, 4.806);
  title(s, 'Telepsychiatry in Mental Healthcare Delivery', 7.33, 2.04, 5.639, 2.121);
  copy(s, [
    { text: 'More intensively in order ', options: { bold: true } },
    { text: 'to contribute to the necessary improvement of adequate. Approaches',
      options: { breakLine: true } },
    { text: 'PLACEHOLDER' },
  ], 7.371, 4.326, 5.009, 1.468);

  s.addShape('round1Rect', {
    x: 0, y: 4.541, w: 6.357, h: 1.681, rectRadius: 0.28, fill: { color: NAVY },
  });
  copy(s, '\u201CIt is concluded that epidemiological and anthropological approaches should '
    + 'concentrate more intensively', 0.796, 5.021, 5.193, 0.761, { color: PAPER });
}

function slide17(pptx) {
  const s = pptx.addSlide();

  // navy blob spilling off the top-right corner
  s.addShape('custGeom', {
    x: 7.68, y: -0.003, w: 5.653, h: 3.718, fill: { color: NAVY },
    points: [
      { x: 0.011, y: 0.000 }, { x: 5.653, y: 0.000 }, { x: 5.653, y: 3.309 },
      { x: 5.340, y: 3.445 },
      { x: 3.844, y: 3.718, curve: { type: 'cubic', x1: 4.880, y1: 3.620, x2: 4.375, y2: 3.718 } },
      { x: 0.000, y: 0.248, curve: { type: 'cubic', x1: 1.721, y1: 3.718, x2: 0.000, y2: 2.164 } },
      { x: 0.005, y: 0.069, curve: { type: 'cubic', x1: 0.000, y1: 0.188, x2: 0.002, y2: 0.128 } },
      { x: 0.011, y: 0.000 }, { close: true },
    ],
  });

  // laptop mock-up: dark lid, grey screen, silver base with a hinge lip
  s.addShape('roundRect', {
    x: 7.093, y: 1.873, w: 5.3, h: 3.49, rectRadius: 0.1, fill: { color: DEVICE },
  });
  s.addShape('rect', { x: 7.232, y: 2.056, w: 5.022, h: 3.139, fill: { color: PHOTO } });
  s.addShape('roundRect', {
    x: 6.473, y: 5.39, w: 6.535, h: 0.19, rectRadius: 0.08, fill: { color: DEVICE_BASE },
  });
  s.addShape('roundRect', {
    x: 9.28, y: 5.39, w: 0.78, h: 0.07, rectRadius: 0.03, fill: { color: 'D3D6D9' },
  });
  s.addShape('rect', { x: 7.093, y: 5.555, w: 5.3, h: 0.05, fill: { color: '65696D' } });

  kicker(s, 0.908, 0.882, 4.806);
  title(s, 'Telepsychiatry in Mental Healthcare Delivery', 0.908, 1.242, 5.639, 2.121);

  [{ y: 3.739, cy: 3.821, n: '1' }, { y: 4.662, cy: 4.745, n: '2' }].forEach((row) => {
    copy(s, [
      { text: 'More intensively in order ', options: { bold: true } },
      { text: 'to contribute to the necessary improvement of' },
    ], 1.771, row.y, 3.728, 0.761);
    s.addShape('ellipse', {
      x: 1.017, y: row.cy, w: 0.511, h: 0.511, fill: { color: NAVY }, shadow: shadow(),
    });
    text(s, row.n, {
      x: 1.017, y: row.cy, w: 0.511, h: 0.511,
      fontSize: 16, color: WHITE, align: 'center', valign: 'middle',
    });
  });

  copy(s, 'PLACEHOLDER'
    + 'more intensively in order to contribute to the necessary improvement of adequate. '
    + 'Approaches should', 2.145, 6.064, 9.192, 0.761, { align: 'center' });
}

function slide18(pptx) {
  const s = pptx.addSlide();

  kicker(s, 4.264, 0.708, 4.806, { align: 'center' });
  title(s, 'Infographic Table Medical', 2.784, 1.027, 7.765, 0.774, { align: 'center' });

  const tiers = [
    { card: 1.787, hx: 2.201, label: 'Silver', px: 2.406, bx: 3.741, tx: 3.701, price: '4',
      shadow: dropShadow(0.12, 36, 6), list: 2.139 },
    { card: 5.200, hx: 5.603, label: 'Gold', px: 5.817, bx: 7.154, tx: 7.114, price: '6',
      shadow: dropShadow(0.15, 66, 30), list: 5.506 },
    { card: 8.613, hx: 9.027, label: 'Diamond', px: 9.233, bx: 10.567, tx: 10.527, price: '11',
      shadow: dropShadow(0.12, 36, 6), list: 8.930 },
  ];
  tiers.forEach((t) => {
    s.addShape('roundRect', {
      x: t.card, y: 2.707, w: 2.933, h: 3.991, rectRadius: 0.2046,
      fill: { color: WHITE }, shadow: t.shadow,
    });
    bottomRoundedBar(s, t.hx, 2.707, 2.105, 0.513, 0.2565, NAVY);
    text(s, t.label, {
      x: t.hx, y: 2.745, w: 2.105, h: 0.438, fontSize: 20, color: WHITE, align: 'center',
    });
    s.addShape('ellipse', { x: t.px, y: 3.451, w: 1.694, h: 1.694, fill: { color: PHOTO } });
    s.addShape('roundRect', {
      x: t.bx, y: 3.453, w: 0.554, h: 0.552, rectRadius: 0.138, fill: { color: NAVY },
    });
    text(s, [
      { text: '$', options: { fontSize: 14 } },
      { text: t.price, options: { fontSize: 18 } },
    ], { x: t.tx, y: 3.521, w: 0.613, h: 0.404, color: WHITE, align: 'center' });
    text(s, bulletList(
      ['It is concluded that', 'epidemiological and', 'anthropological'],
      { fontSize: 14, color: GREY, lineSpacingMultiple: 1.5 },
      { characterCode: '2714', indent: 22.5 },
    ), { x: t.list, y: 5.396, w: 2.382, h: 1.044 });
  });
}

function slide19(pptx) {
  const s = pptx.addSlide();

  kicker(s, 4.264, 0.567, 4.806, { align: 'center' });
  title(s, 'Contact Us Healthcare & Medical', 2.604, 0.928, 8.124, 0.774, { align: 'center' });
  copy(s, [
    { text: 'More intensively ', options: { bold: true } },
    { text: 'PLACEHOLDER'
      + 'PLACEHOLDER' },
    { text: 'necessary', options: { bold: true } },
  ], 2.484, 1.814, 8.365, 0.761, { align: 'center' });

  const rows = [
    { y: 3.094, fill: NAVY, glyph: 'globe', ix: 4.022, iy: 3.212, iw: 0.439, ih: 0.439, bold: true },
    { y: 4.124, fill: BLUE, glyph: 'mail', ix: 4.022, iy: 4.331, iw: 0.463, ih: 0.287 },
    { y: 5.151, fill: NAVY, glyph: 'phone', ix: 4.041, iy: 5.291, iw: 0.425, ih: 0.421 },
    { y: 6.178, fill: BLUE, glyph: 'home', ix: 4.004, iy: 6.296, iw: 0.499, ih: 0.408 },
  ];
  rows.forEach((row) => {
    s.addShape('roundRect', {
      x: 4.454, y: row.y, w: 4.976, h: 0.701, rectRadius: 0.1168,
      fill: { color: WHITE }, shadow: shadow(),
    });
    s.addShape('roundRect', {
      x: 3.903, y: row.y, w: 0.701, h: 0.701, rectRadius: 0.1168,
      fill: { color: row.fill }, shadow: shadow(),
    });
    icon(s, row.glyph, row.ix, row.iy, row.iw, row.ih, WHITE, row.fill);
    copy(s, row.bold
      ? [{ text: 'More intensively in order ', options: { bold: true } }, { text: 'to contribute' }]
      : 'More intensively in order to contribute',
    5.186, row.y + 0.202, 3.738, 0.408);
  });
}

function slide20(pptx) {
  const s = pptx.addSlide();
  navyCrossBackdrop(s, [[8.534, -0.801], [-0.765, 3.281]]);

  title(s, [
    { text: 'Blockchain Technology in ', options: { fontFace: SANS_LIGHT } },
    { text: 'Healthcare Data', options: { fontFace: SANS } },
    { text: ' Management', options: { fontFace: SANS_LIGHT } },
  ], 2.102, 2.557, 9.129, 1.447, { color: WHITE, align: 'center' });

  text(s, 'Open research issues in this area are predominantly', {
    x: 2.575, y: 4.172, w: 8.184, h: 0.337,
    fontFace: SANS_LIGHT, fontSize: 14, color: PAPER, align: 'center',
  });
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE';
  pptx.author = 'pptxgenjs';
  pptx.title = 'PLACEHOLDER';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach((fn) => fn(pptx));

  return pptx.writeFile({
    fileName: path.join(__dirname, '05770343-e0f0-4949-a85c-39a8841415a1_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((err) => {
  console.error(err);
  process.exit(1);
});
