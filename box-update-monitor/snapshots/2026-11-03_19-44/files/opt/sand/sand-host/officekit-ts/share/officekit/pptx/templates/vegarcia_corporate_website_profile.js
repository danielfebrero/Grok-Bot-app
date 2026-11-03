/*
 * "Vegarcia" — 36-slide business presentation template, rebuilt with pptxgenjs.
 *
 * Slide size: 13.333 x 7.5 in (16:9).  All coordinates are inches, all font
 * sizes points, all colours 6-digit hex — exactly as in the source deck.
 *
 * The original is a photo template: every picture slot is an *empty* picture
 * placeholder that shows nothing until an image is dropped in.  Those slots are
 * reproduced here by imagePlaceholder(), which records the geometry and draws
 * nothing (pass { show: true } to see a labelled grey stand-in instead).
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette
const BEIGE   = 'EBCFB2';   // brand accent
const WHITE   = 'FFFFFF';
const CARD    = 'F2F2F2';   // card face (bg1, lum -5 %)
const INK_15  = '262626';   // tx1 lum 85/15
const INK_25  = '404040';   // tx1 lum 75/25
const INK_35  = '595959';   // tx1 lum 65/35
const INK_50  = '808080';   // tx1 lum 50/50
const GREY_7F = '7F7F7F';
const GREY_A6 = 'A6A6A6';
const GREY_BF = 'BFBFBF';
const GREY_D9 = 'D9D9D9';
const GREY_E7 = 'E7E6E6';
const MAP_GREY  = 'D9D9D9'; // countries on slide 33
const CONT_GREY = 'F2F2F2'; // continents on slide 32
const PH_FILL = 'DCDCDC';   // visible photo stand-in (opt-in)
const PH_LINE = 'C4C4C4';

const SANS   = 'Montserrat';
const SCRIPT = 'Yellowtail';

// Text-box insets, in points, ordered [left, right, bottom, top] per pptxgenjs.
const PAD_SHAPE = [7.2, 7.2, 3.6, 3.6];  // PowerPoint default for auto-shapes
const PAD_NONE  = [0, 0, 0, 0];          // text boxes authored with zero inset

const CARD_SHADOW = { type: 'outer', color: '000000', opacity: 0.36, blur: 46, offset: 14, angle: 120 };
const ICON_SHADOW = { type: 'outer', color: '000000', opacity: 0.40, blur: 18, offset: 3, angle: 90 };

// ---------------------------------------------------------------- helpers

/** Filled rectangle, no outline. */
function box(s, x, y, w, h, color, extra) {
  s.addShape('rect', Object.assign({ x, y, w, h, fill: { color }, line: { type: 'none' } }, extra));
}

/** Filled ellipse, no outline. */
function oval(s, x, y, w, h, color, extra) {
  s.addShape('ellipse', Object.assign({ x, y, w, h, fill: { color }, line: { type: 'none' } }, extra));
}

/** Empty photo frame (see file header). */
function imagePlaceholder(s, x, y, w, h, opts) {
  const o = opts || {};
  s.addShape(o.round ? 'ellipse' : 'rect', {
    x, y, w, h,
    fill: o.show ? { color: PH_FILL } : { type: 'none' },
    line: o.show ? { color: PH_LINE, width: 0.75 } : { type: 'none' }
  });
  if (o.show && w >= 0.9 && h >= 0.5) {
    s.addText('[image]', {
      x, y, w, h, align: 'center', valign: 'middle',
      fontFace: SANS, fontSize: 10, color: '9A9A9A'
    });
  }
}

/** Text box: top-anchored, Montserrat, zero inset unless told otherwise. */
function text(s, bodyText, o) {
  s.addText(bodyText, Object.assign({
    fontFace: SANS, color: INK_35, valign: 'top', margin: PAD_NONE
  }, o));
}

/** 16 pt bold section heading — appears on nearly every content slide. */
function subtitle(s, x, y, w, color, str) {
  text(s, str || 'Subtitle Here', {
    x, y, w, h: 0.37, fontSize: 16, bold: true, color: color || INK_35, margin: PAD_SHAPE
  });
}

/** 9 pt body copy at 200 % leading (the deck's standard paragraph). */
function body(s, x, y, w, h, str, o) {
  const p = o || {};
  text(s, str, {
    x, y, w, h,
    fontSize: p.size || 9, color: p.color || INK_35, bold: p.bold || false,
    align: p.align || 'left', lineSpacingMultiple: p.line || 2,
    margin: p.margin || PAD_SHAPE
  });
}

/** Script wordmark "Vegarcia" (Yellowtail 96 pt). */
function wordmark(s, x, y, color, rotate, w) {
  text(s, 'Vegarcia', {
    x, y, w: w || 7.491, h: 1.885, fontSize: 96, fontFace: SCRIPT, color,
    align: 'center', rotate: rotate || 0, margin: [19.2, 19.2, 9.6, 9.6]
  });
}

/** Beige headline percentage: value at 54 pt, "%" at 24 pt. */
function bigPercent(s, x, y, w, value) {
  s.addText([
    { text: String(value), options: { fontSize: 54 } },
    { text: '%', options: { fontSize: 24 } }
  ], {
    x, y, w, h: 0.828, bold: true, fontFace: SANS, color: BEIGE,
    valign: 'top', lineSpacingMultiple: 0.8, margin: PAD_NONE
  });
}

/**
 * Vertical beige-to-transparent wash (slides 17 & 18).  pptxgenjs has no
 * gradient fill, so it is banded: opaque at the bottom, clear at the top.
 */
function beigeWash(s, x, w) {
  const bands = 16, h = 7.5 / bands;
  for (let i = 0; i < bands; i++) {
    const opacity = Math.max(0, 1 - i / (bands - 1)); // 1 at the bottom row
    s.addShape('rect', {
      x, y: 7.5 - (i + 1) * h, w, h: h + 0.01,
      fill: { color: BEIGE, transparency: Math.round((1 - opacity) * 100) },
      line: { type: 'none' }
    });
  }
}

// ------------------------------------------------- small pictograms
// The source deck uses hand-drawn vector icons; these rebuild the same motifs
// from native pptxgenjs shapes so the artwork stays readable in code.
const MINI = {
  cart(s, x, y, d, c) {
    s.addShape('trapezoid', { x: x + d * 0.15, y: y + d * 0.2, w: d * 0.7, h: d * 0.42, flipV: true, fill: { color: c }, line: { type: 'none' } });
    oval(s, x + d * 0.24, y + d * 0.68, d * 0.15, d * 0.15, c);
    oval(s, x + d * 0.58, y + d * 0.68, d * 0.15, d * 0.15, c);
  },
  gear(s, x, y, d, c) {
    s.addShape('gear6', { x: x + d * 0.06, y: y + d * 0.12, w: d * 0.88, h: d * 0.76, fill: { color: c }, line: { type: 'none' } });
  },
  download(s, x, y, d, c) {
    s.addShape('downArrow', { x: x + d * 0.28, y: y + d * 0.08, w: d * 0.44, h: d * 0.6, fill: { color: c }, line: { type: 'none' } });
    box(s, x + d * 0.12, y + d * 0.76, d * 0.76, d * 0.12, c);
  },
  print(s, x, y, d, c) {
    box(s, x + d * 0.26, y + d * 0.08, d * 0.48, d * 0.2, c);
    s.addShape('roundRect', { x: x + d * 0.08, y: y + d * 0.3, w: d * 0.84, h: d * 0.38, fill: { color: c }, line: { type: 'none' }, rectRadius: 0.02 });
    box(s, x + d * 0.26, y + d * 0.72, d * 0.48, d * 0.2, c);
  },
  bars(s, x, y, d, c) {
    [[0.14, 0.5], [0.42, 0.72], [0.7, 0.34]].forEach(([bx, bh]) => {
      box(s, x + d * bx, y + d * (0.86 - bh), d * 0.17, d * bh, c);
    });
  },
  cube(s, x, y, d, c) {
    s.addShape('cube', { x: x + d * 0.15, y: y + d * 0.15, w: d * 0.7, h: d * 0.7, fill: { color: c }, line: { type: 'none' } });
  },
  people(s, x, y, d, c) {
    oval(s, x + d * 0.36, y + d * 0.14, d * 0.28, d * 0.28, c);
    oval(s, x + d * 0.1, y + d * 0.3, d * 0.24, d * 0.24, c);
    oval(s, x + d * 0.66, y + d * 0.3, d * 0.24, d * 0.24, c);
    s.addShape('arc', { x: x + d * 0.12, y: y + d * 0.42, w: d * 0.76, h: d * 0.7, angleRange: [180, 360], fill: { color: c }, line: { type: 'none' } });
  },
  globe(s, x, y, d, c) {
    s.addShape('donut', { x: x + d * 0.12, y: y + d * 0.12, w: d * 0.76, h: d * 0.76, fill: { color: c }, line: { type: 'none' } });
    box(s, x + d * 0.12, y + d * 0.46, d * 0.76, d * 0.08, c);
  },
  heart(s, x, y, d, c) {
    s.addShape('heart', { x: x + d * 0.18, y: y + d * 0.2, w: d * 0.64, h: d * 0.6, fill: { color: c }, line: { type: 'none' } });
  },
  key(s, x, y, d, c) {
    s.addShape('donut', { x: x + d * 0.1, y: y + d * 0.28, w: d * 0.42, h: d * 0.42, fill: { color: c }, line: { type: 'none' } });
    box(s, x + d * 0.48, y + d * 0.44, d * 0.42, d * 0.1, c);
    box(s, x + d * 0.78, y + d * 0.54, d * 0.08, d * 0.14, c);
  },
  diamond(s, x, y, d, c) {
    s.addShape('diamond', { x: x + d * 0.16, y: y + d * 0.22, w: d * 0.68, h: d * 0.56, fill: { color: c }, line: { type: 'none' } });
  },
  lock(s, x, y, d, c) {
    s.addShape('blockArc', { x: x + d * 0.26, y: y + d * 0.1, w: d * 0.48, h: d * 0.48, angleRange: [180, 360], fill: { color: c }, line: { type: 'none' } });
    s.addShape('roundRect', { x: x + d * 0.18, y: y + d * 0.42, w: d * 0.64, h: d * 0.46, fill: { color: c }, line: { type: 'none' }, rectRadius: 0.02 });
  },
  doc(s, x, y, d, c) {
    s.addShape('flowChartDocument', { x: x + d * 0.16, y: y + d * 0.14, w: d * 0.68, h: d * 0.72, fill: { color: c }, line: { type: 'none' } });
  },
  sliders(s, x, y, d, c) {
    box(s, x + d * 0.06, y + d * 0.3, d * 0.88, d * 0.07, c);
    box(s, x + d * 0.06, y + d * 0.62, d * 0.88, d * 0.07, c);
    oval(s, x + d * 0.6, y + d * 0.22, d * 0.22, d * 0.22, c);
    oval(s, x + d * 0.2, y + d * 0.54, d * 0.22, d * 0.22, c);
  }
};

/** White circular badge with a beige pictogram inside (slides 23, 24). */
function iconBadge(s, x, y, icon, d) {
  const size = d || 0.706;
  s.addShape('ellipse', { x, y, w: size, h: size, fill: { color: WHITE }, line: { type: 'none' }, shadow: ICON_SHADOW });
  MINI[icon](s, x + size * 0.27, y + size * 0.27, size * 0.46, BEIGE);
}

// ------------------------------------------------- shared lorem strings
const L_COMPANY = 'A company is an association or collection of individuals, whether natural persons, legal persons, or a mixture of both. ';
const L_MEMBERS = 'PLACEHOLDER';
const L_HALF    = L_COMPANY + L_MEMBERS;
const L_LONG    = L_HALF + 'A company is an association or collection of individuals, whether natural persons, legal persons, or a mixture of both. ' + L_MEMBERS;
const L_SHORT   = L_COMPANY + 'Company members share';
const L_FOCUS   = L_COMPANY + 'Company members share a common purpose and unite in order to focus their various. ' + L_COMPANY;
const L_DOT     = L_COMPANY + 'Company members share a common purpose and unite in order to focus their various.';
const L_IPSUM   = 'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula .';
const L_CONSEC  = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed';
const L_CTETUR  = 'Lorem ipsum dolor sit amet, conse ctetur';
const L_SERVICE = 'A company is an association or collection of individuals, whether natural persons, legal persons';
const L_STRATEGY = 'Lorem ipsum dolor sit amet, consectur adipiscing elit, sed Lorem ipsum dolor sit amet, consectetur';

// ---------------------------------------------------------------- slides

// 1 — full-bleed hero photo behind a white nav bar and the script wordmark.
function slide01(s) {
  imagePlaceholder(s, 0, 0, 13.333, 7.5);
  box(s, 0.706, 0.615, 11.922, 0.529, WHITE);
  const NAV = [
    ['Sign Up',  0.866, 1.003, BEIGE],
    ['About',    7.913, 0.837, INK_15],
    ['Trands',   9.095, 0.903, INK_15],
    ['Contact', 10.342, 1.017, INK_15],
    ['Login',   11.726, 0.780, INK_15]
  ];
  NAV.forEach(([label, x, w, color]) => {
    text(s, label, { x, y: 0.692, w, h: 0.337, fontSize: 14, color, align: 'center', margin: PAD_SHAPE });
  });
  wordmark(s, 2.921, 2.935, WHITE);
}

/** Beige field with the round "puzzle tab" bump (slides 2 and 3). */
function tabbedPanel(s, panelX, panelW, tabX) {
  box(s, panelX, 0, panelW, 7.5, BEIGE);
  oval(s, tabX, 3.148, 1.201, 1.203, BEIGE);
}

/** Light drop-shadow card carrying a subtitle and one paragraph. */
function card(s, x, y, tx, ty, tw) {
  s.addShape('rect', { x, y, w: 4.562, h: 1.991, fill: { color: CARD }, line: { type: 'none' }, shadow: CARD_SHADOW });
  subtitle(s, tx, ty, tw, INK_50);
  body(s, tx, ty + 0.339, 3.583, 0.965, L_COMPANY, { align: 'justify', color: INK_50 });
}

// 2 — beige left panel, one photo, one card, wordmark rotated 270 deg.
function slide02(s) {
  tabbedPanel(s, 0, 8.317, 7.713);
  imagePlaceholder(s, 1.095, 1.333, 5.095, 3.294);
  card(s, 2.834, 3.632, 3.281, 3.945, 2.909);
  wordmark(s, 7.253, 2.812, BEIGE, 270);
}

// 3 — mirrored panel with two stacked cards and two square photos.
function slide03(s) {
  tabbedPanel(s, 3.508, 9.825, 2.892);
  card(s, 5.310, 0.890, 5.758, 1.203, 2.528);
  card(s, 5.326, 3.921, 5.773, 4.234, 2.341);
  imagePlaceholder(s, 9.461, 0.890, 2.84, 2.778);
  imagePlaceholder(s, 9.461, 3.921, 2.84, 2.778);
}

// 4 — beige field, statistics column, tall portrait photo.
function slide04(s) {
  box(s, 0, 0, 9.444, 7.5, BEIGE);
  subtitle(s, 0.981, 1.479, 2.533, INK_25);
  body(s, 0.981, 2.041, 4.943, 1.616,
    L_COMPANY + L_MEMBERS + 'A company is an association or collection of individuals, whether natural persons, legal persons, ',
    { color: INK_25 });
  [['968+', 'PROFESSIONAL', 3.877, 4.382], ['29M', 'HIGH QUALITY', 5.015, 5.520]]
    .forEach(([num, label, ny, ly]) => {
      text(s, num, { x: 0.981, y: ny, w: 1.5, h: 0.505, fontSize: 30, bold: true, color: INK_25 });
      text(s, label, { x: 0.981, y: ly, w: 1.6, h: 0.202, fontSize: 12, color: INK_25 });
    });
  [['Colorful Poster MockUp PSD', 'Jun 01, 2018', 4.079, 4.367],
   ['Realistic T-Shirt Clothing MockUp', 'Mar 01, 2018', 5.298, 5.586]]
    .forEach(([title, date, ty, dy]) => {
      text(s, title, { x: 2.635, y: ty, w: 2.4, h: 0.151, fontSize: 9, color: INK_25 });
      text(s, date, { x: 2.635, y: dy, w: 1.2, h: 0.151, fontSize: 9, color: INK_25 });
    });
  imagePlaceholder(s, 6.492, 0.790, 4.349, 5.92);
}

// 5 — beige right panel with a photo strip and one card.
function slide05(s) {
  box(s, 6.476, 0, 6.857, 7.5, BEIGE);
  subtitle(s, 0.759, 3.114, 2.527, INK_35);
  body(s, 0.759, 3.676, 4.67, 2.524, L_LONG);
  imagePlaceholder(s, 6.476, 1.206, 5.571, 3.398);
  s.addShape('rect', { x: 7.486, y: 4.603, w: 4.562, h: 1.991, fill: { color: CARD }, line: { type: 'none' }, shadow: CARD_SHADOW });
  subtitle(s, 7.933, 4.916, 2.324, INK_50);
  body(s, 7.933, 5.256, 3.583, 1.268, L_SHORT + ' a common', { align: 'justify', color: INK_50 });
}

// 6 — beige disc behind a round photo, wordmark tilted -8 deg.
function slide06(s) {
  oval(s, 1.286, 1.849, 3.952, 3.952, BEIGE);
  imagePlaceholder(s, 1.747, 2.079, 3.492, 3.492, { round: true });
  wordmark(s, 6.421, 0.414, BEIGE, 351.87);
  subtitle(s, 6.743, 3.162, 2.414, INK_35);
  body(s, 6.743, 3.724, 4.67, 2.524, L_LONG);
}

// 7 — two beige discs with round photos plus a 54 % callout.
function slide07(s) {
  oval(s, 4.191, 3.294, 3.952, 3.952, BEIGE);
  oval(s, 7.985, 0.452, 3.952, 3.952, BEIGE);
  imagePlaceholder(s, 4.191, 3.524, 3.492, 3.492, { round: true });
  imagePlaceholder(s, 7.985, 0.683, 3.492, 3.492, { round: true });
  subtitle(s, 1.192, 1.222, 2.380, INK_35);
  body(s, 1.192, 1.613, 4.67, 1.313, L_HALF);
  text(s, '54%', { x: 9.960, y: 4.994, w: 1.657, h: 0.909, fontSize: 48, bold: true, color: INK_35, align: 'center' });
  body(s, 9.627, 5.885, 2.312, 0.768, L_CTETUR, { size: 14, align: 'center', line: 1.5 });
}

// 8 — full-bleed photo, beige right panel, dark tilted wordmark.
function slide08(s) {
  imagePlaceholder(s, 0, 0, 13.333, 7.5);
  box(s, 6.476, 0, 6.857, 7.5, BEIGE);
  wordmark(s, 6.159, 0.851, INK_25, 351.87);
  subtitle(s, 7.368, 3.256, 2.261, INK_25);
  body(s, 7.368, 3.818, 4.67, 2.524, L_LONG, { color: INK_25 });
}

// 9 — one wide photo above a caption block.
function slide09(s) {
  imagePlaceholder(s, 1.255, 1.307, 6.666, 3.526);
  subtitle(s, 4.097, 4.506, 2.689, INK_35);
  body(s, 4.097, 5.068, 8.174, 1.571, L_LONG);
}

// 10 — copy on the left, oversized photo bleeding off the right edge.
function slide10(s) {
  subtitle(s, 1.035, 1.423, 2.951, INK_35);
  body(s, 1.035, 1.984, 3.986, 2.783, L_LONG);
  imagePlaceholder(s, 5.512, 2.913, 7.150, 3.972);
}

// 11 — hero photo above three round team avatars.
function slide11(s) {
  imagePlaceholder(s, 1.815, 1.195, 9.212, 3.840);
  const AVATARS = [[6.233, 6.129, 6.018], [8.437, 8.437, 8.326], [10.640, 10.634, 10.524]];
  AVATARS.forEach(([ax, nx, cx]) => {
    imagePlaceholder(s, ax, 4.628, 1.561, 1.561, { round: true });
    text(s, 'Your Name', { x: nx, y: 6.346, w: 1.74, h: 0.202, fontSize: 12, bold: true, color: INK_35, align: 'center', valign: 'middle' });
    body(s, cx, 6.548, 1.961, 0.404, L_IPSUM, { size: 6, align: 'center', line: 1.5 });
  });
}

// 12 — beige block top right, portrait photo, 85 % statistic.
function slide12(s) {
  box(s, 7.792, 0, 5.542, 4.854, BEIGE);
  imagePlaceholder(s, 6.938, 1.771, 3.938, 4.500);
  subtitle(s, 1.098, 1.969, 2.487, INK_35);
  body(s, 1.146, 2.427, 4.738, 1.313, L_HALF);
  bigPercent(s, 1.098, 4.021, 3.224, 85);
  body(s, 1.165, 4.694, 3.498, 0.965, L_SHORT);
}

// 13 — two offset beige bands, each with copy and a photo.
function slide13(s) {
  const ROWS = [
    { bandX: 1.124, y: 0.957, textX: 1.560, subW: 3.197, copyY: 1.863, photoX: 6.666, copy: L_DOT },
    { bandX: 3.088, y: 4.054, textX: 3.410, subW: 2.505, copyY: 4.960, photoX: 8.630,
      copy: L_COMPANY + 'Company members share a common purpose and unite in order to focus their various talents.' }
  ];
  ROWS.forEach(r => {
    box(s, r.bandX, r.y, 5.542, 2.674, BEIGE);
    subtitle(s, r.textX, r.y + 0.344, r.subW, INK_25);
    body(s, r.textX, r.copyY, 4.67, 1.313, r.copy, { color: INK_25 });
    imagePlaceholder(s, r.photoX, r.y, 3.928, 2.674);
  });
}

// 14 — three stacked percentage statistics beside a wide photo.
function slide14(s) {
  box(s, 0, 3.529, 6.353, 4.006, BEIGE);
  imagePlaceholder(s, 1.464, 2.681, 7.115, 3.493);
  [[85, 9.141, 0.985, 1.658, 3.498, 1.010],
   [76, 9.187, 2.976, 3.684, 3.380, 1.313],
   [94, 9.187, 4.959, 5.709, 3.498, 0.965]].forEach(([v, px, py, ty, tw, th]) => {
    bigPercent(s, px, py, 3.224, v);
    body(s, 9.208, ty, tw, th, L_SHORT);
  });
}

// 15 — vertical wordmark, two beige bands with photos and copy.
function slide15(s) {
  wordmark(s, -0.75, 2.808, BEIGE, 270, 6.184);
  [[1.072, 1.553, 1.923, 1.313, 2.558], [4.087, 4.448, 4.818, 1.268, 2.443]]
    .forEach(([by, sy, py, ph, sw]) => {
      box(s, 7.450, by, 5.883, 2.333, BEIGE);
      imagePlaceholder(s, 4.174, by, 3.276, 2.333);
      subtitle(s, 8.057, sy, sw, INK_25);
      body(s, 8.057, py, 4.67, ph, L_DOT, { color: INK_25 });
    });
}

// 16 — "Our Team": beige field, ghosted script title, five avatars.
function slide16(s) {
  s.background = { color: BEIGE };
  text(s, 'Our Team', {
    x: 0.562, y: 1.069, w: 4.647, h: 1.481, fontSize: 88, bold: true, fontFace: SCRIPT,
    color: WHITE, transparency: 75, charSpacing: 3
  });
  text(s, 'Our Team', { x: 7.320, y: 1.836, w: 2.607, h: 0.606, fontSize: 36, bold: true, color: INK_25 });
  body(s, 7.321, 2.499, 2.725, 0.336, 'Director at Grammo for 15 years ', { size: 10.5, line: 1.5, bold: true, color: INK_25 });
  body(s, 7.320, 2.865, 5.180, 0.965,
    L_COMPANY + 'Company members share a common purpose and unite in order to focus.', { bold: true, color: INK_25 });
  const TEAM = [[1.504, 1.410, 1.299], [3.708, 3.640, 3.529], [5.911, 5.870, 5.759],
                [8.115, 8.006, 7.895], [10.319, 10.263, 10.153]];
  TEAM.forEach(([ax, nx, cx]) => {
    imagePlaceholder(s, ax, 4.441, 1.561, 1.561, { round: true });
    text(s, 'Your Name', { x: nx, y: 6.242, w: 1.74, h: 0.202, fontSize: 12, bold: true, color: INK_25, align: 'center', valign: 'middle', charSpacing: 3 });
    body(s, cx, 6.444, 1.961, 0.387, L_IPSUM, { size: 6, align: 'center', line: 1.5, color: INK_25 });
  });
}

/** Skill meter: grey track, beige fill, centred label and percentage. */
function skillBar(s, trackX, y, fillW, label, labelX, labelW, pct, pctX) {
  s.addShape('roundRect', { x: trackX + 0.016, y, w: 3.319, h: 0.155, fill: { color: GREY_E7 }, line: { type: 'none' }, rectRadius: 0.077 });
  s.addShape('roundRect', { x: trackX, y, w: fillW, h: 0.155, fill: { color: BEIGE }, line: { type: 'none' }, rectRadius: 0.077 });
  text(s, label, { x: labelX, y: y - 0.333, w: labelW, h: 0.236, fontSize: 14, bold: true, color: INK_35, align: 'center' });
  text(s, pct, { x: pctX, y: y - 0.416, w: 0.56, h: 0.303, fontSize: 18, bold: true, color: INK_35 });
}

/** Team-member profile (slides 17 and 18) — washed photo panel on one side. */
function profileSlide(s, panelX, colX, name) {
  beigeWash(s, panelX, 5.478);
  imagePlaceholder(s, panelX, 0, 5.478, 7.5);
  text(s, name, { x: colX, y: 1.362, w: 5.479, h: 0.993, fontSize: 28, bold: true, color: INK_35, valign: 'middle' });
  body(s, colX + 0.017, 2.088, 2.725, 0.366, 'Director at Grammo for 15 years ', { size: 10.5, line: 1.5 });
  body(s, colX, 2.514, 5.182, 1.616, L_FOCUS);
  skillBar(s, colX + 0.017, 4.714, 1.994, 'Design Grafis',   colX - 0.029, 1.373, '60%', colX + 2.851);
  skillBar(s, colX + 0.017, 5.813, 2.469, 'Adobe Ilustrator', colX - 0.044, 1.683, '70%', colX + 2.851);
}

function slide17(s) { profileSlide(s, 0,     6.814, 'Ms. Sarah Azhari H.'); }
function slide18(s) { profileSlide(s, 7.855, 1.520, 'Mr. Alexander Patos'); }

// 19 — two stacked photos, tilted wordmark, copy on the right.
function slide19(s) {
  imagePlaceholder(s, 1.079, 0.902, 4.879, 3.189);
  imagePlaceholder(s, 1.079, 4.312, 4.879, 2.165);
  wordmark(s, 5.616, 0.766, BEIGE, 351.87);
  subtitle(s, 7.027, 3.391, 2.530, INK_35);
  body(s, 7.027, 3.953, 4.67, 2.524, L_LONG);
}

// 20 — full-width intro copy above a row of three photos.
function slide20(s) {
  subtitle(s, 0.688, 1.214, 2.398, INK_35);
  body(s, 0.688, 1.776, 11.021, 0.965, L_LONG);
  [0.688, 4.417, 8.146].forEach(x => imagePlaceholder(s, x, 3.792, 3.249, 2.958));
}

/** One tall photo with copy on the opposite side (slides 21 and 22). */
function photoAndCopy(s, photoX, copyX) {
  imagePlaceholder(s, photoX, 1.146, 4.327, 5.208);
  subtitle(s, copyX, 2.078, 2.841, INK_35);
  body(s, copyX, 2.640, 4.67, 2.524, L_LONG);
}
function slide21(s) { photoAndCopy(s, 1.333, 6.930); }
function slide22(s) { photoAndCopy(s, 7.812, 1.588); }

// 23 — "Our Services": beige field with four icon rows.
function slide23(s) {
  s.background = { color: BEIGE };
  s.addText([
    { text: 'Our', options: { breakLine: true } },
    { text: 'Services' }
  ], { x: 1.645, y: 2.034, w: 3.4, h: 1.548, fontSize: 46, bold: true, fontFace: SANS, color: INK_25, valign: 'top', margin: PAD_NONE });
  body(s, 1.645, 3.944, 3.597, 1.818, L_FOCUS, { color: INK_25, margin: PAD_NONE });
  const ROWS = [[1.031, 'cart'], [2.587, 'gear'], [4.157, 'download'], [5.713, 'print']];
  ROWS.forEach(([y, icon]) => {
    iconBadge(s, 7.394, y - 0.024, icon);
    text(s, 'Sample Text', { x: 8.385, y, w: 2.2, h: 0.329, fontSize: 18, color: INK_25, margin: [2, 2, 2, 2] });
    body(s, 8.385, y + 0.326, 3.851, 0.617, L_SERVICE, { color: INK_25, margin: [2, 2, 2, 2] });
  });
}

// 24 — "Our Company Marketing Strategy": 2 x 2 icon-and-copy grid.
function slide24(s) {
  s.background = { color: BEIGE };
  text(s, 'Our Company Marketing Strategy', {
    x: 3.732, y: 0.969, w: 5.516, h: 1.292, fontSize: 28, bold: true, color: INK_25,
    align: 'center', valign: 'middle'
  });
  const CELLS = [
    [1.865, 2.879, 3.081, 2.905, 2.754, 'download'],
    [1.865, 5.129, 3.081, 5.154, 2.754, 'print'],
    [7.913, 2.872, 9.130, 2.898, 2.765, 'cart'],
    [7.913, 5.122, 9.130, 5.148, 2.765, 'gear']
  ];
  CELLS.forEach(([ix, iy, tx, ty, tw, icon]) => {
    iconBadge(s, ix, iy, icon);
    text(s, 'Your Text Here', { x: tx, y: ty, w: 1.9, h: 0.269, fontSize: 16, bold: true, color: INK_25 });
    body(s, tx, ty + 0.259, tw, 0.909, L_STRATEGY, { color: INK_25, margin: PAD_NONE });
  });
}

// 25 — three portrait photos, each stamped with a large white numeral.
function slide25(s) {
  [[1.159, '01'], [4.984, '02'], [8.810, '03']].forEach(([x, n]) => {
    imagePlaceholder(s, x, 1.043, 3.365, 5.413);
    text(s, n, { x: x - 0.017, y: 0.85, w: 1.7, h: 1.4, fontSize: 72, bold: true, color: WHITE });
  });
}

/** Beige statistic tile: 48 pt percentage over a two-line caption. */
function statTile(s, x, y, w, value, pctX, pctW, capX) {
  box(s, x, y, w, 3.351, BEIGE);
  text(s, value + '%', { x: pctX, y: y + 0.744, w: pctW, h: 0.909, fontSize: 48, bold: true, color: WHITE, align: 'center' });
  body(s, capX, y + 1.635, 2.312, 0.771, L_CTETUR, { size: 14, align: 'center', line: 1.5, color: WHITE });
}

// 26 — pinwheel of three photos and three beige percentage tiles.
function slide26(s) {
  imagePlaceholder(s, 0.488, 0.420, 4.116, 3.351);
  imagePlaceholder(s, 4.609, 3.771, 4.116, 3.351);
  imagePlaceholder(s, 8.719, 0.420, 4.116, 3.351);
  statTile(s, 8.719, 3.771, 4.116, '76', 9.972, 1.639, 9.631);
  statTile(s, 4.592, 0.420, 4.142, '88', 5.831, 1.680, 5.510);
  statTile(s, 0.476, 3.748, 4.116, '25', 1.755, 1.590, 1.389);
}

// 27 — 2 x 3 photo mosaic with a single beige 70 % tile.
function slide27(s) {
  [[0.376, 0.396], [0.375, 3.863], [4.629, 0.396], [8.883, 0.396], [8.882, 3.863]]
    .forEach(([x, y]) => imagePlaceholder(s, x, y, 4.055, 3.268));
  box(s, 4.628, 3.863, 4.055, 3.267, BEIGE);
  text(s, '70%', { x: 5.871, y: 4.636, w: 1.668, h: 0.909, fontSize: 48, bold: true, color: WHITE, align: 'center' });
  body(s, 5.544, 5.528, 2.312, 0.771, L_CTETUR, { size: 14, align: 'center', line: 1.5, color: WHITE });
}

/**
 * Flat-vector laptop built from rectangles.  (x, y) is the top-left of the
 * base bar's bounding box, `w` its full width; every part is proportional.
 */
function laptop(s, x, y, w, screenColor) {
  const H = w * 0.5716;              // total height incl. base
  const lidX = x + w * 0.1013, lidW = w * 0.805;
  box(s, lidX - w * 0.0027, y, lidW + w * 0.0051, H * 0.9709, 'D2D3D5');   // lid rim
  box(s, lidX, y + H * 0.005, lidW, H * 0.9611, '181818');                 // lid
  box(s, lidX + lidW * 0.0331, y + H * 0.0653, lidW * 0.9331, H * 0.8318, screenColor || 'EDF1F4');
  oval(s, x + w * 0.4989, y + H * 0.0295, w * 0.0085, H * 0.0148, '2C99B4'); // camera
  box(s, x, y + H * 0.9494, w, H * 0.032, 'D2D6D7');                       // base slab
  box(s, x, y + H * 0.9654, w * 0.5037, H * 0.0389, 'B3B4B5');
  box(s, x + w * 0.4963, y + H * 0.9654, w * 0.5037, H * 0.0389, 'B3B4B5');
  box(s, x + w * 0.4278, y + H * 0.9494, w * 0.1435, H * 0.018, 'B3B4B5'); // thumb notch
}

// 28 — "One Smart Laptop Mockup" with an 85 % stat and two icon rows.
function slide28(s) {
  laptop(s, 6.704, 1.423, 7.657);
  s.addText([
    { text: 'One Smart Laptop', options: { breakLine: true } },
    { text: 'Mockup' }
  ], { x: 0.641, y: 1.261, w: 5.0, h: 1.212, fontSize: 28, bold: true, fontFace: SANS, color: INK_35, valign: 'top', margin: [19.2, 19.2, 9.6, 9.6] });
  bigPercent(s, 0.771, 2.944, 4.012, 85);
  body(s, 0.854, 3.905, 4.354, 0.965, L_SHORT);
  MINI.doc(s, 1.028, 5.778, 0.44, INK_15);
  MINI.sliders(s, 4.103, 5.806, 0.44, GREY_BF);
  text(s, 'Main title', { x: 1.533, y: 5.884, w: 1.4, h: 0.269, fontSize: 16, bold: true, color: INK_35 });
  text(s, 'Main title', { x: 4.693, y: 5.884, w: 1.4, h: 0.269, fontSize: 16, bold: true, color: INK_35 });
}

// 29 — laptop with four beige percentage bubbles and captions.
function slide29(s) {
  laptop(s, 3.877, 1.407, 5.508, 'C6C6C6');
  const BUBBLES = [
    { ox: 8.642, oy: 1.132, pct: '65%', px: 8.611, py: 1.248, pw: 0.735, cx: 9.327, cy: 1.128, align: 'left'  },
    { ox: 3.872, oy: 1.841, pct: '80%', px: 3.849, py: 1.965, pw: 0.761, cx: 1.373, cy: 1.844, align: 'right' },
    { ox: 8.614, oy: 2.824, pct: '32%', px: 8.590, py: 2.940, pw: 0.721, cx: 9.300, cy: 2.820, align: 'left'  },
    { ox: 4.006, oy: 3.838, pct: '45%', px: 3.990, py: 3.961, pw: 0.747, cx: 1.507, cy: 3.841, align: 'right' }
  ];
  BUBBLES.forEach(b => {
    oval(s, b.ox, b.oy, 0.695, 0.695, BEIGE);
    text(s, b.pct, { x: b.px, y: b.py, w: b.pw, h: 0.404, fontSize: 18, color: WHITE, align: 'center' });
    body(s, b.cx, b.cy, 2.489, 0.555, L_CONSEC, { align: b.align, line: 1.5, margin: PAD_NONE });
  });
  body(s, 1.951, 5.670, 9.356, 0.909, L_FOCUS, { align: 'center', margin: PAD_NONE });
}

// 30 — "Pentagon Process": five up/down chevron pairs across a soft band.
function slide30(s) {
  // x/y/w/h are the *unrotated* frames; the 90/270 deg spin makes them point
  // down / up, exactly as authored in the source deck.
  const STEPS = [
    { up: [0.871, 3.400, 1.352], dn: [1.209, 5.493, 0.676], color: GREY_D9, top: '45', bot: '-15', numX: 1.317, numY: [3.399, 5.767], label: 'Identify',   labelColor: INK_15,  textX: 0.426, textW: 2.246 },
    { up: [2.468, 3.073, 2.007], dn: [2.875, 5.751, 1.192], color: GREY_BF, top: '60', bot: '-30', numX: 3.241, numY: [2.787, 6.256], label: 'Brainstorm', labelColor: GREY_BF, textX: 2.238, textW: 2.495 },
    { up: [4.637, 3.318, 1.516], dn: [4.929, 5.622, 0.933], color: GREY_D9, top: '50', bot: '-25', numX: 5.165, numY: [3.168, 5.985], label: 'Design',     labelColor: INK_50,  textX: 4.125, textW: 2.541 },
    { up: [6.164, 2.921, 2.311], dn: [6.607, 5.868, 1.425], color: BEIGE,   top: '75', bot: '-50', numX: 7.089, numY: [2.484, 6.487], label: 'Build',      labelColor: INK_50,  textX: 6.094, textW: 2.451 },
    { up: [8.484, 3.318, 1.516], dn: [8.776, 5.622, 0.933], color: GREY_BF, top: '52', bot: '-40', numX: 9.012, numY: [3.245, 5.998], label: 'Test',       labelColor: INK_35,  textX: 8.062, textW: 2.361 }
  ];
  STEPS.forEach(st => {
    s.addShape('homePlate', { x: st.up[0], y: st.up[1], w: st.up[2], h: 1.009, rotate: 270, fill: { color: st.color }, line: { type: 'none' } });
    s.addShape('homePlate', { x: st.dn[0], y: st.dn[1], w: st.dn[2], h: 1.009, rotate: 90,  fill: { color: st.color }, line: { type: 'none' } });
    [st.top, st.bot].forEach((n, i) => {
      text(s, n, { x: st.numX, y: st.numY[i], w: 0.461, h: 0.461, fontSize: 16, bold: true, color: WHITE, align: 'center', valign: 'middle' });
    });
  });
  // White band across the middle (rotated 270 deg, with a large soft shadow).
  s.addShape('rect', {
    x: 4.871, y: -0.278, w: 1.067, h: 10.810, rotate: 270,
    fill: { color: WHITE }, line: { type: 'none' },
    shadow: { type: 'outer', color: '000000', opacity: 0.3, blur: 60, offset: 0, angle: 0 }
  });
  STEPS.forEach(st => {
    s.addText([
      { text: st.label, options: { fontSize: 12, color: st.labelColor, breakLine: true } },
      { text: 'Imperdiet nec, imperdiet iaculis.', options: { fontSize: 11, color: GREY_7F } }
    ], { x: st.textX, y: 4.692, w: st.textW, h: 0.839, bold: true, fontFace: SANS, align: 'center', valign: 'top', margin: [19, 19, 9.5, 9.5] });
  });
  text(s, 'Pentagon Process', { x: 0.779, y: 1.140, w: 5.121, h: 0.789, fontSize: 28, bold: true, color: INK_35 });
  text(s, 'Insert Your Creative Idea ', { x: 0.809, y: 1.722, w: 3.757, h: 0.211, fontSize: 12, color: INK_35, valign: 'middle', margin: PAD_SHAPE });
  body(s, 10.269, 4.332, 2.572, 1.919,
    L_COMPANY + 'PLACEHOLDER', { bold: true, margin: PAD_NONE });
}

// 31 — "Hexagon Steps": four hexagons climbing a grey diagonal.
function slide31(s) {
  s.addShape('line', { x: 2.305, y: 2.792, w: 11.028, h: 4.708, flipV: true, line: { color: GREY_D9, width: 6 } });
  const STEPS = [
    { hexX: 4.997, hexY: 5.323, color: GREY_A6, icon: 'heart',   label: 'Step 1', labelX: 5.087,  labelY: 6.704, textX: 4.935,  textY: 3.457, stemY: 4.562 },
    { hexX: 6.984, hexY: 4.445, color: GREY_BF, icon: 'key',     label: 'Step 2', labelX: 7.075,  labelY: 5.830, textX: 6.910,  textY: 2.728, stemY: 3.732 },
    { hexX: 8.971, hexY: 3.605, color: BEIGE,   icon: 'diamond', label: 'Step 3', labelX: 9.062,  labelY: 4.955, textX: 8.877,  textY: 1.849, stemY: 2.893 },
    { hexX: 10.959, hexY: 2.724, color: GREY_BF, icon: 'lock',   label: 'Step 4', labelX: 11.021, labelY: 4.074, textX: 11.021, textY: 1.016, stemY: 2.063 }
  ];
  STEPS.forEach(st => {
    s.addShape('hexagon', { x: st.hexX, y: st.hexY, w: 1.202, h: 1.350, fill: { color: st.color }, line: { type: 'none' } });
    MINI[st.icon](s, st.hexX + 0.371, st.hexY + 0.445, 0.46, WHITE);
    text(s, st.label, { x: st.labelX, y: st.labelY, w: 1.02, h: 0.37, fontSize: 16, bold: true, color: INK_35, align: 'center' });
    text(s, 'Your Text', { x: st.textX, y: st.textY, w: 1.967, h: 0.343, fontSize: 16, bold: true, color: INK_35, lineSpacingMultiple: 0.9 });
    body(s, st.textX, st.textY + 0.287, 1.967, 0.783, 'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet.', { line: 1.5 });
    s.addShape('line', { x: st.textX + 0.662, y: st.stemY, w: 0, h: 0.8, line: { color: INK_35, width: 1, dashType: 'dash' } });
  });
  text(s, 'Hexagon Steps', { x: 0.820, y: 1.347, w: 3.605, h: 0.525, fontSize: 28, bold: true, color: INK_35, lineSpacingMultiple: 0.9 });
  text(s, 'Insert Your Creative Idea ', { x: 0.820, y: 1.926, w: 3.757, h: 0.211, fontSize: 12, color: INK_35, valign: 'middle', margin: PAD_SHAPE });
  body(s, 0.820, 2.532, 3.321, 1.919, L_HALF, { align: 'justify' });
}

// World-map outline traced from the source deck, normalised to a 1000 x 1000
// box: each entry is one closed polygon as a flat [x0,y0, x1,y1, ...] list.
const WORLD = [
  [995,122,965,100,951,100,951,111,943,104,924,111,918,93,873,86,877,79,858,93,845,86,839,97,829,82,827,89,823,72,769,75,788,57,759,43,752,54,753,61,734,54,715,64,714,79,698,72,707,86,704,100,693,82,690,93,677,86,684,104,674,129,664,122,674,111,671,89,674,79,658,89,666,104,661,115,647,100,641,100,641,111,614,111,603,122,596,122,595,111,593,132,589,125,574,147,558,118,585,118,554,104,558,150,547,186,552,183,549,197,557,201,570,240,584,247,576,283,600,301,603,301,600,283,606,272,600,258,603,236,611,229,638,236,642,211,664,201,679,218,687,208,695,233,704,233,715,251,728,236,745,244,748,226,759,240,801,247,808,218,823,218,831,247,842,261,851,258,846,283,840,287,846,297,864,254,867,229,850,208,867,179,905,175,903,168,913,154,926,161,934,147,930,168,910,197,911,236,916,218,927,208,926,190,932,168,949,168,973,150,972,140,962,140,964,132,1000,122,995,122],
  [223,118,206,104,198,129,196,111,193,118,171,111,166,115,166,133,163,118,144,118,144,108,117,100,114,108,108,97,90,111,73,104,73,165,100,197,101,219,106,219,109,237,117,237,124,251,201,247,219,258,226,247,241,276,238,298,268,280,272,265,268,269,283,240,309,233,307,219,298,222,302,212,293,204,293,187,287,194,290,187,287,172,279,194,271,187,274,183,269,169,272,169,261,158,252,158,249,183,255,201,247,212,249,244,238,233,236,212,204,197,207,190,200,169,203,151,211,147,201,140,220,140,223,118],
  [823,312,839,294,837,283,848,262,839,265,823,226,812,222,810,237,797,255,805,258,809,273,785,287,780,305,761,312,741,301,734,287,726,287,725,269,715,255,710,273,704,269,696,308,679,326,685,348,696,348,693,380,698,391,728,409,742,394,750,419,747,434,753,455,767,441,780,459,797,445,812,405,805,394,812,387,810,369,804,362,809,344,801,326,812,308,813,323,823,312],
  [277,283,274,269,271,283,258,287,237,315,231,280,223,301,222,283,226,276,214,280,207,272,214,265,207,258,125,258,124,269,119,262,122,269,119,315,128,333,130,358,139,373,158,384,168,376,177,398,184,394,190,419,204,391,225,387,233,391,242,427,239,380,252,358,250,330,268,308,268,298,277,283],
  [323,595,301,609,297,584,290,591,293,598,286,613,275,605,272,616,275,649,266,652,263,670,266,688,271,674,271,695,285,684,286,702,305,735,307,778,313,781,320,814,310,839,320,857,331,828,331,806,353,778,358,709,370,677,370,663,356,641,345,638,340,652,342,631,326,620,323,595],
  [900,824,900,806,881,760,870,702,864,753,850,727,854,709,843,706,836,720,836,734,824,724,818,738,818,763,815,745,790,796,796,874,842,849,851,871,859,857,864,892,873,903,891,896,900,824],
  [400,100,388,97,384,75,396,83,394,72,413,57,405,47,413,29,410,22,413,14,400,22,403,11,377,11,396,0,343,4,343,14,301,11,290,18,288,29,267,36,283,43,275,50,304,50,318,83,313,86,326,90,327,104,318,122,323,136,324,129,329,136,327,147,334,165,346,169,346,151,353,144,348,136,353,129,369,126,377,108,400,100],
  [603,254,609,269,620,265,614,290,619,301,627,308,627,279,641,272,644,290,655,290,663,311,671,294,694,297,693,283,702,265,710,265,713,251,704,240,694,243,685,218,679,226,664,208,644,218,642,243,612,236,608,251,603,243,603,254],
  [307,889,302,867,304,842,312,824,305,824,302,817,305,806,291,785,282,785,271,849,274,871,269,889,267,975,263,986,266,1000,274,1000,274,989,266,986,277,986,278,975,271,975,282,968,277,957,286,935,285,921,293,917,285,892,302,903,307,889],
  [718,445,717,423,694,409,694,391,690,387,693,351,679,362,683,380,672,412,666,416,671,438,663,441,672,459,677,430,676,481,687,552,693,545,694,502,718,445],
  [67,101,27,97,29,90,14,97,3,108,14,115,17,129,0,133,16,129,16,147,3,158,13,165,14,179,33,165,25,194,36,176,33,165,40,161,65,158,62,169,70,169,67,101],
  [446,412,480,473,502,440,494,423,495,390,489,369,492,344,467,355,467,380,446,412],
  [566,577,562,552,576,484,573,455,541,459,533,513,536,548,547,581,554,577,557,588,566,577],
  [552,666,549,634,555,591,546,580,533,588,524,580,516,645,506,656,517,656,519,673,533,666,533,695,549,706,547,684,552,666],
  [747,247,726,244,718,255,728,269,726,280,761,305,780,298,785,280,805,265,794,262,794,247,777,255,748,233,747,247],
  [592,484,601,491,623,466,625,452,617,452,606,409,576,384,568,405,592,484],
  [645,416,639,394,641,348,630,337,614,348,603,326,601,333,595,326,601,351,600,366,615,409,630,412,636,427,645,416],
  [192,427,182,398,176,405,168,384,150,384,173,452,174,473,207,502,212,495,211,480,217,473,218,480,223,459,215,462,212,480,201,484,192,452,192,427],
  [538,398,533,376,524,394,503,376,497,394,497,423,505,441,514,437,538,466,538,398],
  [264,918,274,799,279,789,271,746,274,710,260,678,261,656,269,638,258,624,249,656,242,652,242,660,255,721,272,764,261,921,266,925,264,918],
  [559,824,557,785,552,785,541,810,527,821,527,832,517,832,522,871,549,857,559,824],
  [513,548,518,555,530,537,535,498,535,473,514,444,514,495,508,512,513,519,513,548],
  [505,742,533,746,530,710,536,702,530,699,530,674,517,681,514,663,505,663,509,695,505,742],
  [489,512,495,519,506,512,513,487,511,452,503,444,483,477,481,502,472,509,476,523,489,512],
  [446,530,450,537,462,505,480,498,480,480,454,434,454,505,437,505,438,523,443,519,446,530],
  [589,538,584,516,573,513,565,556,571,581,581,588,597,577,601,559,592,552,589,538],
  [283,699,282,692,276,702,274,745,279,781,291,778,295,760,306,753,298,738,298,720,283,699],
  [250,563,248,606,272,634,271,602,277,598,277,573,263,559,263,534,250,563],
  [562,394,566,405,565,387,541,387,541,452,568,444,559,401,560,387,562,394],
  [424,466,424,491,435,502,453,498,451,427,446,419,437,427,434,462,424,466],
  [593,326,590,312,544,308,547,344,593,337,593,326],
  [577,663,574,642,557,627,554,663,568,688,568,699,581,696,577,663],
  [506,201,509,204,520,172,516,158,528,136,527,122,535,126,533,111,517,115,509,147,505,147,506,201],
  [497,563,502,563,510,527,484,516,478,567,492,581,497,563],
  [663,433,668,433,665,412,679,387,677,355,685,355,680,344,672,347,666,383,658,387,655,401,646,401,649,419,644,426,657,423,663,433],
  [269,541,266,559,280,570,282,606,288,599,286,581,298,581,294,573,298,556,291,541,275,534,269,541],
  [517,821,524,825,527,753,506,746,503,749,511,781,513,817,517,821],
  [250,86,236,89,241,96,249,89,263,107,266,114,263,136,272,143,271,150,286,147,274,129,279,118,293,125,272,104,272,93,267,100,256,82,250,86],
  [568,244,565,233,555,240,541,233,533,258,549,255,555,269,566,273,568,283,566,269,581,258,579,247,568,244],
  [467,294,488,294,484,273,491,255,475,240,478,247,457,258,467,273,467,294],
  [671,355,669,333,666,344,655,340,647,358,642,355,641,372,646,387,642,394,655,394,671,355],
  [582,724,582,699,568,706,571,720,568,745,563,724,555,727,563,738,560,806,568,792,566,760,582,734,582,724],
  [549,111,549,100,536,111,543,136,530,154,533,169,557,151,549,111],
  [529,807,530,814,541,803,549,782,540,753,530,756,529,807],
  [456,340,465,337,476,304,446,294,453,308,450,337,456,340],
  [592,541,606,559,597,584,589,588,587,624,606,577,614,538,612,530,592,541],
  [744,477,752,462,744,441,745,409,731,459,737,480,736,495,741,495,744,477],
  [516,591,525,573,532,584,543,573,533,538,511,570,516,591],
  [609,724,609,713,601,738,595,738,593,799,597,803,601,799,609,724],
  [584,623,586,591,576,595,567,584,570,602,567,623,579,645,584,623],
  [491,230,488,240,492,262,506,254,503,251,503,240,510,237,506,215,495,212,491,230],
  [141,89,155,89,152,100,142,100,150,107,182,96,174,93,169,75,165,89,142,75,138,79,141,89],
  [778,623,780,638,796,641,799,588,789,613,777,609,778,623],
  [624,330,624,337,636,333,644,351,654,337,638,304,620,319,624,330],
  [511,226,522,247,537,240,533,215,518,212,511,226],
  [597,351,587,347,581,376,600,398,605,390,597,372,597,351],
  [286,11,291,7,255,4,255,18,242,22,239,40,253,36,249,25,286,11],
  [761,505,766,509,763,484,755,491,753,473,747,473,750,523,753,512,758,519,761,505],
  [552,720,540,702,538,717,533,717,533,738,546,746,554,731,552,720],
  [446,398,464,376,462,358,453,355,443,376,438,409,446,398],
  [635,283,628,287,628,308,638,297,660,333,661,319,654,297,638,294,635,283],
  [763,627,755,602,739,580,764,659,767,641,763,627],
  [867,667,867,685,881,674,891,692,878,649,867,642,867,667],
  [617,487,612,480,601,498,592,491,592,520,616,498,617,487],
  [560,760,560,742,554,734,543,753,552,778,560,774,560,760],
  [495,581,499,599,514,599,508,545,495,581],
  [304,777,304,763,296,763,294,785,309,803,305,817,312,817,315,799,304,777],
  [544,262,535,262,528,276,535,290,549,290,544,262],
  [557,215,555,204,548,204,538,229,555,232,557,215],
  [847,638,845,645,859,656,859,678,866,678,866,638,858,631,851,645,847,638]
];

/** Paint WORLD scaled into the rectangle (x, y, w, h). */
function worldMap(s, x, y, w, h, color) {
  WORLD.forEach(flat => {
    const pts = [];
    for (let i = 0; i < flat.length; i += 2) {
      pts.push({ x: (flat[i] / 1000) * w, y: (flat[i + 1] / 1000) * h });
    }
    pts.push({ close: true });
    s.addShape('custGeom', { x, y, w, h, points: pts, fill: { color }, line: { type: 'none' } });
  });
}

/** Dashed flight path: a single cubic Bezier through four control points. */
function dashedArc(s, x, y, w, h, c1, c2, end) {
  s.addShape('custGeom', {
    x, y, w, h,
    points: [
      { x: 0, y: 0 },
      { x: end[0] * w, y: end[1] * h, curve: { type: 'cubic', x1: c1[0] * w, y1: c1[1] * h, x2: c2[0] * w, y2: c2[1] * h } }
    ],
    fill: { type: 'none' },
    line: { color: GREY_BF, width: 1.5, dashType: 'dash' }
  });
}

// 32 — world map with five journey pins joined by dashed flight paths.
function slide32(s) {
  worldMap(s, 0.657, 1.038, 11.549, 6.062, CONT_GREY);
  dashedArc(s, -0.008, 1.928, 10.462, 3.273, [0.045, 0.423], [0.306, 0.965], [1, 0.2]);
  dashedArc(s, 10.475, 2.547, -10.481, 3.673, [0.157, 0.404], [0.481, 0.809], [1, 0.789]);
  const PINS = [
    { x: 0.670, y: 2.991, color: BEIGE,   icon: 'bars',   textX: 1.836, textY: 2.730, textW: 2.455, align: 'left'  },
    { x: 3.194, y: 4.460, color: BEIGE,   icon: 'cube',   textX: 4.273, textY: 3.823, textW: 2.543, align: 'left'  },
    { x: 5.280, y: 5.358, color: GREY_D9, icon: 'people', textX: 6.440, textY: 5.730, textW: 2.544, align: 'left'  },
    { x: 7.365, y: 3.916, color: GREY_BF, icon: 'print',  textX: 8.484, textY: 4.289, textW: 2.521, align: 'left'  },
    { x: 9.578, y: 2.275, color: GREY_A6, icon: 'globe',  textX: 7.031, textY: 2.140, textW: 2.402, align: 'right' }
  ];
  PINS.forEach(p => {
    oval(s, p.x, p.y, 1.061, 1.061, p.color);
    MINI[p.icon](s, p.x + 0.32, p.y + 0.32, 0.42, WHITE);
    text(s, 'Subtitle Here', { x: p.textX, y: p.textY, w: p.textW, h: 0.37, fontSize: 16, bold: true, color: INK_35, align: p.align, margin: PAD_SHAPE });
    body(s, p.align === 'right' ? p.textX + 0.104 : p.textX, p.textY + 0.292, 2.298, 0.555, L_CONSEC, { align: p.align, line: 1.5 });
  });
  // Rocket accent flying off the north-east pin.
  s.addShape('teardrop', { x: 10.745, y: 1.605, w: 0.733, h: 0.732, rotate: 315, fill: { color: BEIGE }, line: { type: 'none' } });
  s.addShape('triangle', { x: 10.639, y: 2.181, w: 0.262, h: 0.262, rotate: 200, fill: { color: BEIGE }, line: { type: 'none' } });
  s.addShape('triangle', { x: 10.994, y: 2.232, w: 0.211, h: 0.207, rotate: 110, fill: { color: BEIGE }, line: { type: 'none' } });
}

// 33 — "How To Run A Successful Worldwide Map": title, dash rule, full map.
function slide33(s) {
  text(s, 'How To Run A Successful Worldwide Map', {
    x: 3.858, y: 0.552, w: 5.699, h: 1.043, fontSize: 28, bold: true, color: INK_35, align: 'center'
  });
  [5.733, 6.124, 6.515, 6.906, 7.297].forEach(x => box(s, x, 1.699, 0.303, 0.072, BEIGE));
  worldMap(s, 0.933, 2.412, 11.442, 5.049, MAP_GREY);
}

// 34 — closing card: "Thank You / For Watching" on a full beige field.
function slide34(s) {
  box(s, 0, 0, 13.333, 7.5, BEIGE);
  s.addShape('line', { x: 4.482, y: 2.915, w: 0, h: 1.662, flipV: true, line: { color: WHITE, width: 2.25 } });
  s.addText([
    { text: 'Thank You', options: { breakLine: true } },
    { text: 'For Watching' }
  ], { x: 4.707, y: 2.955, w: 4.722, h: 1.582, fontSize: 44, fontFace: SCRIPT, color: WHITE, charSpacing: 3, valign: 'top', margin: PAD_SHAPE });
}

// 35 / 36 — icon-library sheets: a 19 x 7 grid of pictograms.  The originals
// are ~130 bespoke vector glyphs each; here they cycle through native shapes.
const ICON_SET_A = [
  'cloud', 'teardrop', 'donut', 'ellipse', 'plus', 'mathMinus', 'noSmoking', 'downArrow', 'upArrow',
  'leftArrow', 'rightArrow', 'chartPlus', 'chartX', 'chartStar', 'flowChartDisplay', 'flowChartMagneticDisk',
  'rect', 'frame', 'plaque'
];
const ICON_SET_B = [
  'heart', 'star5', 'sun', 'moon', 'lightningBolt', 'gear6', 'cube', 'can', 'diamond', 'hexagon',
  'octagon', 'triangle', 'bevel', 'smileyFace', 'funnel', 'flowChartDocument', 'folderCorner',
  'flowChartMultidocument', 'flowChartTerminator'
];

function iconSheet(s, palette, x0, y0, dx, dy, cell, skipLast) {
  const COLS = 19, ROWS = 7;
  for (let r = 0; r < ROWS; r++) {
    for (let c = 0; c < COLS; c++) {
      if (skipLast && r === ROWS - 1 && c === COLS - 1) continue;
      s.addShape(palette[(r * 5 + c) % palette.length], {
        x: x0 + c * dx + (dx - cell) / 2,
        y: y0 + r * dy + (dy - cell) / 2,
        w: cell, h: cell,
        fill: { color: INK_25 }, line: { type: 'none' }
      });
    }
  }
}

function slide35(s) { iconSheet(s, ICON_SET_A, 1.667, 1.859, 0.5515, 0.5535, 0.30, true); }
function slide36(s) { iconSheet(s, ICON_SET_B, 1.127, 1.779, 0.5949, 0.5963, 0.29, false); }

// ---------------------------------------------------------------- assembly
const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
  slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18,
  slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27,
  slide28, slide29, slide30, slide31, slide32, slide33, slide34, slide35, slide36
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'VEGARCIA_16x9', width: 13.333, height: 7.5 });
  pptx.layout = 'VEGARCIA_16x9';
  pptx.title = 'Vegarcia';
  pptx.author = 'Vegarcia';

  BUILDERS.forEach(buildSlide => {
    const slide = pptx.addSlide();
    slide.background = { color: WHITE };
    buildSlide(slide);
  });

  return pptx.writeFile({
    fileName: path.join(__dirname, '058b8072-4de4-4fb4-806b-9a1336260366_grok_final.pptx')
  });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
