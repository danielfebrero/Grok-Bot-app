/**
 * Steale Beauty — skincare pitch deck (45 slides, 13.333" x 7.5").
 * Recreated with pptxgenjs. Photographic content in the original template is
 * replaced by flat vector placeholders (see `photo()` / `product()`).
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Palette / typography (from the deck theme "Custom 458")
 * ------------------------------------------------------------------ */
const C = {
  blush: 'E9B3A3',       // accent1
  blushLight: 'FEE3D8',  // accent2
  cream: 'EEEAD8',       // accent3
  green: '4D6F51',       // accent5
  forest: '245849',      // accent6
  forestDark: '122C24',  // accent6, 50% lum
  white: 'FFFFFF',
  body: '3A3A3A',        // bg2 -75%
  bodyGrey: '747373',
  gold: 'FFC000',
  greyStar: 'D9D9D9',
  track: 'E1DDDE',       // grey progress track
  sage: '94A997',
  pageTop: 'FBF0ED',     // page gradient start
  pageBottom: 'F6E1DA',  // page gradient end
};

const SERIF = 'Yeseva One';
const SANS = 'Epilogue';

const W = 13.333;        // slide width
const H = 7.5;           // slide height

/** Lorem-style copy blocks reused all over the deck. */
const T = {
  brand:
    'A brand is\u00a0an intangible marketing or business concept that helps people identify a ' +
    'company, product, or individual. People often confuse brands with things like logos, slogans, ' +
    'or other recognizable marks, which are marketing tools that help promote goods and services.',
  brandShort:
    'A brand is\u00a0an intangible marketing or business concept that helps people identify a ' +
    'company, product, or individual. ',
  confuse:
    'People often confuse brands with things like logos, slogans, or other recognizable marks, ' +
    'which are marketing tools that help promote goods and services.',
  confuseShort: 'People often confuse brands with things like logos, slogans, or other recognizable marks.',
  confuseTiny: 'People often confuse brands with things like logos, slogans.',
  confuseNoDot: 'People often confuse brands with things like logos, slogans, or other recognizable marks',
  confuseTools: 'People often confuse brands with things like logos, slogans, or other recognizable marks, which are marketing tools.',
  confusePromote: 'People often confuse brands with things like logos, slogans, or other recognizable marks, which are marketing tools that help promote',
  quoteTitle: ['\u201cBest Branding In ', 'Quite Ways Are The Most Beautiful\u201d'],
  tagline: 'MINIMALIST\tDESIGN GRADIENT\tTOUCH\tVINTAGE\tSTYLE BRAND',
};

/* ------------------------------------------------------------------ *
 * Small helpers
 * ------------------------------------------------------------------ */
const NOLINE = { type: 'none' };

/** Plain text box. Reference boxes are top-anchored with 0.1"/0.05" insets. */
function txt(slide, text, o) {
  slide.addText(text, Object.assign(
    { fontFace: SANS, fontSize: 12, color: C.forest, valign: 'top', margin: [3.6, 7.2, 3.6, 7.2] },
    o
  ));
}

/** Filled shape with no outline. */
function shape(slide, type, o) {
  slide.addShape(type, Object.assign({ line: NOLINE }, o));
}

/** Outlined shape, no fill. */
function outline(slide, type, o) {
  slide.addShape(type, Object.assign({ fill: { type: 'none' }, line: { color: C.green, width: 2.25 } }, o));
}

/** Page background: soft blush vertical gradient, faked with stacked bands. */
function background(slide) {
  const bands = 30;
  const from = [0xfb, 0xf0, 0xed];
  const to = [0xf6, 0xe1, 0xda];
  for (let i = 0; i < bands; i++) {
    const t = i / (bands - 1);
    const hex = from.map((v, k) => Math.round(v + (to[k] - v) * t))
      .map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
    shape(slide, 'rect', { x: 0, y: (H / bands) * i, w: W, h: H / bands + 0.02, fill: { color: hex } });
  }
}

/** Footer band repeated on every slide by the master. */
function footer(slide, num) {
  txt(slide, [
    { text: 'Soft Opening ', options: { bold: true, color: C.forest } },
    { text: 'Beautycare Shop', options: { color: C.forestDark } },
  ], { x: 0.33, y: 6.815, w: 1.962, h: 0.572, fontSize: 14, align: 'left' });
  txt(slide, String(num), { x: 6.167, y: 7.018, w: 0.999, h: 0.438, fontSize: 20, align: 'center' });
  txt(slide, [
    { text: 'Organic ', options: { color: C.forest } },
    { text: 'Skincare', options: { color: C.forestDark } },
  ], { x: 10.685, y: 7.018, w: 2.318, h: 0.37, fontSize: 16, align: 'right' });
}

/**
 * Decorative leafy sprig used in the template corners.
 * Drawn as a stem plus alternating leaf ellipses, so it stays vector-only.
 * (x, y) is where the stem starts, `a` its direction in degrees.
 */
function sprig(slide, x, y, len, a, color) {
  const col = color || C.green;
  const rad = (a * Math.PI) / 180;
  const dx = Math.cos(rad), dy = Math.sin(rad);
  const w = Math.abs(len * dx), h = Math.abs(len * dy);
  slide.addShape('line', {
    x: dx < 0 ? x - w : x, y: dy < 0 ? y - h : y, w, h,
    flipH: dx < 0, flipV: dy < 0,
    line: { color: col, width: 0.75 },
  });
  const leaf = Math.max(0.14, len * 0.17);
  const n = Math.max(3, Math.round(len / (leaf * 0.95)));
  for (let i = 1; i <= n; i++) {
    const t = (i / (n + 0.4)) * len;
    const px = x + dx * t, py = y + dy * t;
    for (const side of [-1, 1]) {
      const ang = a + side * 48;
      const ar = (ang * Math.PI) / 180;
      shape(slide, 'ellipse', {
        x: px + Math.cos(ar) * leaf * 0.45 - leaf * 0.5,
        y: py + Math.sin(ar) * leaf * 0.45 - leaf * 0.19,
        w: leaf, h: leaf * 0.38,
        rotate: ang,
        fill: { type: 'none' }, line: { color: col, width: 0.6 },
      });
    }
  }
}

/** Sprig positions per slide: [x, y, length, angle]. */
const SPRIGS = {
  1: [[13.33, 6.58, 2.15, -143], [0.0, 6.26, 2.15, -37], [13.33, 0.69, 2.15, 143], [0.0, 0.36, 2.15, 37]],
  2: [[0.69, 5.74, 0.89, -44], [13.27, 1.47, 1.2, 136], [0.08, 0.13, 1.12, 44]],
  3: [[0.4, 5.26, 1.68, -46], [2.44, 0.63, 1.68, 46]],
  4: [[8.39, 0.77, 2.3, 141], [13.33, 6.4, 1.66, -136]],
  5: [[13.33, 0.0, 2.15, 143], [4.72, 6.63, 1.68, -46]],
  6: [[0.0, 0.92, 2.32, 49], [4.84, 6.42, 2.32, -49]],
  7: [[13.27, 1.47, 1.2, 136], [0.08, 0.13, 1.12, 44]],
  8: [[1.55, 0.56, 2.3, 39], [13.33, 6.61, 1.66, -136]],
  9: [[13.33, 0.0, 2.15, 143], [0.94, 1.07, 1.68, 46]],
  10: [[-0.05, 2.87, 2.32, 49], [13.37, 5.01, 2.32, -131]],
  11: [[13.33, 6.4, 1.66, -136]],
  12: [[-0.1, 0.16, 2.3, 39]],
  13: [[13.27, 1.47, 1.2, 136], [0.08, 0.13, 1.12, 44]],
  14: [[13.33, 0.0, 2.15, 143], [-0.1, 6.19, 1.68, -46]],
  15: [[13.12, -0.16, 2.32, 131], [12.33, 6.51, 0.89, -136], [8.71, 0.93, 1.66, 136], [0.16, -0.16, 2.32, 49]],
  16: [[13.27, 0.41, 1.2, 136], [0.08, 0.13, 1.12, 44], [0.08, 6.2, 1.12, -44]],
  17: [[0.06, 1.51, 2.3, 39], [13.33, 6.61, 1.66, -136]],
  18: [[13.33, 0.0, 2.15, 143], [0.0, 0.25, 2.15, 37]],
  19: [[13.27, 2.16, 1.2, 136], [0.08, 0.13, 1.12, 44], [0.06, 6.5, 1.2, -44], [13.37, 6.88, 1.12, -136]],
  20: [[11.44, 6.33, 1.68, -134], [0.79, 1.78, 1.68, 46]],
  21: [[13.27, 5.04, 1.2, -136], [0.08, 0.13, 1.12, 44], [0.06, 6.5, 1.2, -44], [13.37, 6.88, 1.12, -136]],
  22: [[1.1, 0.88, 1.66, 44], [0.99, 5.43, 2.32, -49], [13.33, 6.58, 0.89, -136]],
  23: [[-0.1, 0.16, 2.3, 39], [13.12, -0.16, 2.32, 131]],
  25: [[9.9, 0.94, 2.3, 141], [10.12, 7.4, 1.66, -136]],
  26: [[13.27, 2.16, 1.2, 136], [0.08, 0.13, 1.12, 44], [0.06, 6.5, 1.2, -44], [13.37, 6.88, 1.12, -136]],
  27: [[0.08, 0.13, 1.12, 44], [13.37, 6.88, 1.12, -136]],
  28: [[13.27, 2.16, 1.2, 136], [0.08, 0.13, 1.12, 44], [0.06, 6.79, 1.2, -44], [13.37, 6.88, 1.12, -136]],
  29: [[13.33, 0.0, 2.15, 143], [0.0, 0.25, 2.15, 37]],
  30: [[8.21, 0.86, 1.66, 136], [13.12, -0.16, 2.32, 131], [12.33, 6.51, 0.89, -136], [0.16, -0.16, 2.32, 49]],
  31: [[0.08, 0.13, 1.12, 44], [0.06, 6.79, 1.2, -44]],
  32: [[2.04, 0.09, 1.68, 46], [13.33, 0.0, 2.15, 143]],
  33: [[-0.06, 1.46, 2.3, 39], [13.33, 1.66, 1.66, 136]],
  34: [[13.33, 0.43, 2.15, 143], [0.0, 0.68, 2.15, 37]],
  35: [[0.11, 0.13, 2.32, 49], [13.23, 6.99, 2.32, -131], [3.04, 7.4, 0.89, -44], [13.03, 0.93, 1.66, 136]],
  36: [[0.2, 5.13, 1.66, -44], [13.12, -0.16, 2.32, 131], [12.33, 6.51, 0.89, -136], [0.16, -0.16, 2.32, 49]],
  37: [[13.33, 0.0, 2.15, 143], [0.0, 0.25, 2.15, 37]],
  38: [[0.14, 1.02, 2.3, 39], [5.29, 1.37, 1.66, 44]],
  39: [[0.11, 0.13, 2.32, 49], [13.23, 6.99, 2.32, -131], [3.04, 7.4, 0.89, -44], [12.83, 0.5, 1.66, 136]],
  41: [[13.33, 0.0, 2.15, 143], [-0.35, 6.25, 2.3, -39]],
  42: [[13.77, 0.62, 2.3, 141], [0.0, 5.19, 1.66, -44]],
  43: [[13.33, 6.58, 2.15, -143], [0.0, 6.26, 2.15, -37], [13.33, 0.69, 2.15, 143], [0.0, 0.36, 2.15, 37]],
  44: [[13.33, 0.43, 2.15, 143], [0.0, 0.22, 2.15, 37]],
  45: [[13.33, 5.97, 2.15, -143], [0.0, 0.9, 2.15, 37], [13.27, 2.16, 1.2, 136],
       [0.08, 0.13, 1.12, 44], [0.06, 6.5, 1.2, -44], [13.37, 6.88, 1.12, -136]],
};

/**
 * Placeholder standing in for one of the deck's raster device mockups
 * (watch, phones, laptop, tablet): a light body with a screen panel and an
 * `[image]` caption. Empty picture placeholders in the original template
 * carry no artwork and are therefore not drawn at all.
 */
function photo(slide, x, y, w, h, opts) {
  const o = opts || {};
  const rot = o.rotate || 0;
  shape(slide, 'roundRect', {
    x, y, w, h, rectRadius: o.rectRadius === undefined ? 0.12 : o.rectRadius, rotate: rot,
    fill: { color: o.body || 'EFEFEF' }, line: { color: 'C9C9C9', width: 1 },
  });
  const pad = o.pad === undefined ? 0.12 : o.pad;
  const sx = x + pad, sy = y + (o.padY === undefined ? pad : o.padY);
  const sw = w - pad * 2, sh = h - (o.padY === undefined ? pad : o.padY) * 2;
  shape(slide, 'roundRect', {
    x: sx, y: sy, w: sw, h: sh, rectRadius: o.screenRadius || 0.06, rotate: rot,
    fill: { color: o.screen || '3B3F46' },
  });
  txt(slide, '[image]', {
    x: sx, y: sy + sh / 2 - 0.16, w: sw, h: 0.32, rotate: rot,
    align: 'center', fontSize: 11, color: o.label || 'D5D5D5',
  });
}

/**
 * Flat product illustrations that live in the template layouts, rebuilt from
 * rounded rectangles / ellipses. `parts` are fractions of the bounding box:
 * [fx, fy, fw, fh, colour, radius].
 */
function product(slide, x, y, w, h, parts) {
  parts.forEach(([fx, fy, fw, fh, col, r]) => {
    shape(slide, r === 'o' ? 'ellipse' : 'roundRect', {
      x: x + fx * w, y: y + fy * h, w: fw * w, h: fh * h,
      rectRadius: typeof r === 'number' ? r : 0,
      fill: { color: col },
    });
  });
}

const ART = {
  // Cream cosmetic jar with a swirl of cream and a tilted lid leaning on it (slide 2).
  jar: [[0.53, 0.33, 0.45, 0.62, 'E9B3A3', 'o'], [0.06, 0.19, 0.58, 0.22, 'FEE3D8', 'o'],
        [0.03, 0.28, 0.64, 0.16, 'D6D3C2', 0.06], [0.03, 0.36, 0.64, 0.58, 'EEEAD8', 0.12],
        [0.03, 0.44, 0.64, 0.05, 'E9B3A3', 0], [0.03, 0.52, 0.64, 0.42, 'E9B3A3', 0.1],
        [0.08, 0.58, 0.05, 0.3, 'FFFFFF', 0.02]],
  // Two stacked skincare tubes (slide 8).
  tubeA: [[0.0, 0.0, 1.0, 0.08, 'E5CCC2', 0.02], [0.0, 0.08, 1.0, 0.82, 'FEE3D8', 0.06],
          [0.0, 0.28, 0.9, 0.35, 'E9B3A3', 0], [0.15, 0.9, 0.7, 0.1, 'BEAAA2', 0.02]],
  tubeB: [[0.0, 0.05, 1.0, 0.78, 'D2A193', 0.18], [0.08, 0.1, 0.72, 0.68, 'E9B3A3', 0.16],
          [0.22, 0.3, 0.5, 0.28, 'D2A193', 'o'], [0.1, 0.83, 0.8, 0.17, '745A52', 0.03]],
  // Serum dropper bottle (slides 9 and 12).
  dropper: [[0.32, 0.0, 0.34, 0.22, '7A6979', 0.03], [0.1, 0.2, 0.8, 0.16, 'C7BAC7', 0.04],
            [0.0, 0.36, 1.0, 0.64, 'FEE3D8', 0.08], [0.0, 0.55, 1.0, 0.22, 'E9B3A3', 0]],
  // Pump bottle (slide 11).
  pump: [[0.4, 0.0, 0.28, 0.13, 'FEE3D8', 0.04], [0.28, 0.1, 0.15, 0.2, 'FEE3D8', 0.03],
         [0.0, 0.31, 0.93, 0.69, 'E5CCC2', 0.16], [0.0, 0.58, 0.93, 0.28, 'D2A193', 0]],
  // Small vial (slide 9, right of the dropper).
  vial: [[0.16, 0.0, 0.68, 0.1, 'AF867A', 0.02], [0.0, 0.1, 1.0, 0.9, 'F2F2F2', 0.05],
         [0.04, 0.55, 0.92, 0.45, 'A28CA1', 0.04]],
  // Slim cleanser bottle (slide 12, left of the dropper).
  bottle: [[0.24, 0.0, 0.52, 0.15, 'E9B3A3', 0.03], [0.0, 0.14, 1.0, 0.86, 'E5CCC2', 0.1],
           [0.0, 0.35, 1.0, 0.45, 'E9B3A3', 0], [0.15, 0.9, 0.3, 0.04, 'ECE8EC', 0.02]],
};

/** Pill button — rounded rect with a caption, used across the deck. */
function pill(slide, x, y, w, h, runs, o) {
  const opt = o || {};
  shape(slide, 'roundRect', {
    x, y, w, h, rectRadius: h / 2,
    fill: { color: opt.fill || C.forest },
  });
  txt(slide, runs, {
    x, y, w, h, align: 'center', valign: 'middle', wrap: false, margin: 0,
    fontSize: opt.fontSize || 12, color: opt.color || C.white,
  });
}

/** Pill with a circular arrow badge on its right side ("Register Here"). */
function arrowPill(slide, x, y) {
  shape(slide, 'roundRect', { x, y, w: 2.164, h: 0.622, rectRadius: 0.311, fill: { color: C.green } });
  txt(slide, 'Register Here', {
    x: x + 0.188, y, w: 1.4, h: 0.622, valign: 'middle', wrap: false, margin: 0,
    fontSize: 14, color: C.cream,
  });
  shape(slide, 'ellipse', { x: x + 1.669, y: y + 0.132, w: 0.359, h: 0.359, fill: { color: C.white } });
  shape(slide, 'rightArrow', { x: x + 1.752, y: y + 0.245, w: 0.19, h: 0.135, fill: { color: C.green } });
}

/** Five-star rating row: 4 gold + 1 grey, as in the template cards. */
function stars(slide, x, y, size) {
  const s = size || 0.163;
  for (let i = 0; i < 5; i++) {
    shape(slide, 'star5', {
      x: x + i * s * 1.28, y, w: s, h: s,
      fill: { color: i < 4 ? C.gold : C.greyStar },
    });
  }
}

/** Small cream/forest diamond holding a two-digit index. */
function diamond(slide, x, y, label, fill, color) {
  shape(slide, 'diamond', { x, y, w: 0.776, h: 0.78, fill: { color: fill } });
  txt(slide, label, {
    x, y, w: 0.776, h: 0.78, align: 'center', valign: 'middle', fontSize: 11, color,
  });
}

/** Circular icon badge with a glyph (contact icons, chevrons). */
function iconCircle(slide, x, y, d, glyph, opts) {
  const o = opts || {};
  shape(slide, 'ellipse', { x, y, w: d, h: d, fill: { color: o.fill || C.green } });
  txt(slide, glyph, {
    x, y, w: d, h: d, align: 'center', valign: 'middle',
    fontSize: o.fontSize || 12, color: o.color || C.white,
  });
}

/** Checkbox-and-caption row ("Insert text here"). */
function insertRow(slide, x, y, color) {
  shape(slide, 'parallelogram', { x, y: y + 0.057, w: 0.322, h: 0.297, fill: { color: color === C.white ? C.cream : C.green } });
  shape(slide, 'rect', { x: x + 0.12, y: y + 0.09, w: 0.24, h: 0.19, rotate: -12, fill: { color: C.blush } });
  txt(slide, [
    { text: 'Insert ', options: { bold: true } },
    { text: 'text here' },
  ], { x: x + 0.536, y, w: 1.818, h: 0.337, fontSize: 14, color });
}

/** Title in the display serif. */
function title(slide, text, o) {
  txt(slide, text, Object.assign({ fontFace: SERIF, fontSize: 44, color: C.forest, valign: 'top' }, o));
}

/** Justified 10pt body paragraph at 150% leading (the deck's default). */
function body(slide, text, o) {
  txt(slide, text, Object.assign(
    { fontSize: 10, color: C.body, align: 'justify', lineSpacingMultiple: 1.5 }, o
  ));
}

/** "EST 1980" eyebrow. */
function est(slide, x, y, w, o) {
  txt(slide, [
    { text: 'EST ', options: { bold: true } },
    { text: '1980' },
  ], Object.assign({ x, y, w, h: 0.337, fontSize: 14, color: C.forest }, o || {}));
}

/** Ordered list 01/02/03/04 with mixed emphasis, used on slides 3 and 18. */
function numberList(slide, x, y, color) {
  const styles = [{}, { italic: true }, { underline: { style: 'sng' } }, { bold: true }];
  ['01', '02', '03', '04'].forEach((n, i) => {
    txt(slide, [{ text: n, options: styles[i] }], {
      x, y: y + i * 0.348, w: 0.409, h: 0.303, align: 'right', fontSize: 12, color,
    });
  });
}

/* ================================================================== *
 * Slides
 * ================================================================== */

/** Slides 1, 20 and 45 share a big brushed oval with centred lockup. */
function heroSlide(slide, n, headline, opts) {
  const o = opts || {};
  shape(slide, 'ellipse', { x: o.ovalX, y: o.ovalY, w: o.ovalW, h: o.ovalH, fill: { color: C.green } });
  est(slide, o.estX, o.estY, 1.536, { align: 'center', color: C.white });
  title(slide, headline, {
    x: o.titleX, y: o.titleY, w: o.titleW, h: 1.582, fontSize: 88, color: C.white, align: 'center',
  });
  title(slide, 'The Newest Skincare Products', {
    x: 3.776, y: o.subY, w: 5.781, h: 0.639, fontSize: 24, color: C.white,
    align: 'center', lineSpacingMultiple: 1.5,
  });
}

function slide1(s) {
  heroSlide(s, 1, [{ text: 'Steale' }, { text: ' Beauty' }], {
    ovalX: 0.856, ovalY: 1.37, ovalW: 11.535, ovalH: 4.572,
    estX: 3.419, estY: 2.718, titleX: 2.296, titleY: 2.663, titleW: 8.936, subY: 4.049,
  });
}

function slide2(s) {
  outline(s, 'round2DiagRect', { x: 1.31, y: 2.245, w: 9.169, h: 3.912 });
  shape(s, 'ellipse', { x: 8.519, y: 1.67, w: 3.623, h: 3.896, fill: { color: C.green } });
  product(s, 8.813, 2.628, 3.681, 3.066, ART.jar);
  title(s, 'Introduction New Products', { x: 1.751, y: 1.499, w: 4.799, h: 1.447, fontSize: 40 });
  est(s, 1.751, 3.041, 1.375);
  txt(s, T.brand, {
    x: 1.751, y: 3.388, w: 6.265, h: 1.283, fontSize: 12, align: 'justify', lineSpacingMultiple: 1.5,
  });
  arrowPill(s, 1.751, 4.929);
}

function slide3(s) {
  outline(s, 'ellipse', { x: 1.538, y: 1.724, w: 3.38, h: 4.642 });
  title(s, [
    { text: T.quoteTitle[0], options: { bold: true, italic: true } },
    { text: T.quoteTitle[1], options: { italic: true } },
  ], { x: 5.61, y: 1.893, w: 6.575, h: 2.121, fontSize: 40, align: 'center' });
  pill(s, 8.005, 4.218, 1.786, 0.464, '\u201cQuote\u201d');
  s.addShape('line', {
    x: 4.629, y: 5.31, w: 0.497, h: 0,
    line: { color: C.forest, width: 1.5, endArrowType: 'triangle' },
  });
  numberList(s, 5.105, 5.159, C.forest);
  body(s, 'People Often Confuse Brands With Things Like Logos, Slogans, Or Other Recognizable Marks, ' +
    'Which Are Marketing Tools That Help Promote Goods And Services.',
    { x: 5.623, y: 5.09, w: 3.353, h: 1.397, lineSpacingMultiple: 2 });
  // Sparkles scattered above the oval.
  [[4.89, 1.13, 0.33], [4.59, 0.78, 0.23], [5.46, 1.05, 0.23], [12.59, 0.83, 0.33], [12.13, 0.65, 0.23]]
    .forEach(([x, y, d]) => shape(s, 'star4', { x, y, w: d, h: d * 1.5, fill: { color: C.green } }));
}

function slide4(s) {
  shape(s, 'rect', { x: 9.543, y: 0, w: 3.79, h: 3.889, fill: { color: C.green } });
  title(s, 'Relaxing Skincare', { x: 1.115, y: 1.446, w: 2.973, h: 1.447, fontSize: 40 });
  txt(s, 'Collage No. 12', {
    x: 1.115, y: 3.124, w: 1.036, h: 0.773, fontSize: 14, bold: true, lineSpacingMultiple: 1.5,
  });
  txt(s, T.tagline, { x: 2.976, y: 3.259, w: 3.283, h: 0.656, fontSize: 11, color: C.green });
  txt(s, T.brand, {
    x: 1.115, y: 4.282, w: 5.064, h: 1.58, fontSize: 12, align: 'justify', lineSpacingMultiple: 1.5,
  });
  pill(s, 8.642, 4.966, 1.786, 0.464, 'View More');
}

function slide5(s) {
  outline(s, 'round2DiagRect', { x: 5.878, y: 1.101, w: 3.257, h: 5.528, flipV: true });
  title(s, [
    { text: 'We Cannot Solve Our Problems With The ', options: { bold: true } },
    { text: 'Same Thinking We Used When We Created Them' },
  ], { x: 0.857, y: 0.989, w: 4.507, h: 4.813, fontSize: 40 });
  est(s, 7.543, 1.46, 1.375, { align: 'right' });
  txt(s, T.brand, {
    x: 6.124, y: 2.156, w: 2.838, h: 3.101, fontSize: 12, align: 'center', lineSpacingMultiple: 1.5,
  });
  s.addShape('line', {
    x: 6.667, y: 5.877, w: 0.858, h: 0,
    line: { color: C.green, width: 7.25, endArrowType: 'stealth' },
  });
  pill(s, 10.312, 5.803, 1.786, 0.464, [{ text: 'Learn ' }, { text: 'More', options: { bold: true } }]);
}

/** Slides 6 and 28 — "Product Reviews" two-column article layout. */
function reviewsSlide(s, o) {
  title(s, 'Product Reviews', { x: o.tx, y: o.ty, w: 5.686, h: 0.841 });
  txt(s, [
    { text: '\u201cLaunching ', options: { bold: true, italic: true } },
    { text: 'Brands\u201d', options: { italic: true } },
  ], { x: o.qx, y: o.qy, w: 3.768, h: 1.043, fontSize: 28, align: 'center' });
  const cols = [['Promotion ', o.cx], ['Searching ', o.cx + 3.202]];
  cols.forEach(([label, x]) => {
    txt(s, [{ text: label, options: { bold: true } }, { text: 'Branding' }],
      { x, y: o.cy, w: 2.235, h: 0.337, fontSize: 14 });
    body(s, T.confuse, { x, y: o.cy + 0.507, w: 2.623, h: 1.334 });
    if (o.secondRow) body(s, T.confuse, { x, y: o.cy + 1.97, w: 2.623, h: 1.334 });
  });
}

function slide6(s) {
  reviewsSlide(s, { tx: 6.4, ty: 1.619, qx: 1.122, qy: 0.72, cx: 6.4, cy: 2.703, secondRow: true });
}

function slide7(s) {
  outline(s, 'round2DiagRect', { x: 1.259, y: 1.519, w: 10.815, h: 4.463 });
  title(s, 'The Customer Target Skincare', { x: 1.879, y: 1.615, w: 5.184, h: 1.582 });
  const cards = [
    { x: 7.435, y: 1.953, dx: 7.631, dy: 2.225, n: '01', label: 'Promotions Brand' },
    { x: 10.128, y: 2.728, dx: 10.324, dy: 3.0, n: '02', label: 'Launching Products' },
  ];
  cards.forEach((c) => {
    shape(s, 'round2DiagRect', { x: c.x, y: c.y, w: 2.522, h: 3.595, flipH: true, fill: { color: C.green } });
    diamond(s, c.dx, c.dy, c.n, C.cream, C.green);
    txt(s, c.label, { x: c.dx - 0.025, y: c.dy + 0.94, w: 2.262, h: 0.337, fontSize: 14, bold: true, color: C.white });
    body(s, T.confuse, { x: c.dx - 0.025, y: c.dy + 1.328, w: 2.262, h: 1.586, align: 'left', color: C.white });
  });
}

/** Slides 8 and 30 — "Promotion Brands" with four insert rows. */
function promotionSlide(s, o) {
  title(s, 'Promotion Brands', { x: o.tx, y: o.ty, w: 6.096, h: 0.841, color: o.titleColor });
  for (let i = 0; i < 4; i++) {
    insertRow(s, o.rx + (i % 2) * 2.505, o.ry + Math.floor(i / 2) * 0.646, o.textColor);
  }
  txt(s, 'Special Products', { x: o.sx + 0.02, y: o.sy, w: 2.361, h: 0.337, fontSize: 14, color: o.accent });
  txt(s, [{ text: 'The ' }, { text: 'Beautycare' }], {
    x: o.sx, y: o.sy + 0.336, w: 2.352, h: 0.37, fontSize: 16, bold: true, color: o.accent,
  });
  iconCircle(s, o.cx, o.cy, 0.572, '\u276f', { fontSize: 16 });
  body(s, T.confuse, { x: o.bx, y: o.sy, w: 3.166, h: 1.086 });
  pill(s, o.px, o.py, 1.786, 0.464, [{ text: 'Learn ' }, { text: 'More', options: { bold: true } }]);
}

function slide8(s) {
  shape(s, 'rect', { x: 3.033, y: 0, w: 10.301, h: 3.531, fill: { color: C.green } });
  product(s, 0.925, 2.343, 1.47, 3.528, ART.tubeA);
  product(s, 2.426, 1.977, 2.047, 3.894, ART.tubeB);
  promotionSlide(s, {
    tx: 5.085, ty: 0.835, titleColor: C.white, rx: 5.28, ry: 1.954, textColor: C.white,
    sx: 5.28, sy: 3.748, accent: C.forest, cx: 7.543, cy: 3.798, bx: 8.431, px: 5.34, py: 4.647,
  });
}

function slide9(s) {
  title(s, [
    { text: T.quoteTitle[0], options: { bold: true, italic: true } },
    { text: T.quoteTitle[1], options: { italic: true } },
  ], { x: 1.792, y: 0.735, w: 9.748, h: 1.582, fontSize: 44 });
  product(s, 9.638, 1.855, 1.454, 3.79, ART.dropper);
  product(s, 11.238, 3.718, 0.838, 1.926, ART.vial);
  [2.766, 3.545].forEach((y) => {
    txt(s, '\u2714', { x: 5.09, y: y - 0.09, w: 0.3, h: 0.3, fontSize: 11, color: C.green });
    body(s, T.confuseShort, { x: 5.414, y: y - 0.135, w: 3.321, h: 0.576 });
  });
  txt(s, 'The Best Brands Price', {
    x: 5.039, y: 4.369, w: 1.511, h: 0.572, fontSize: 14, bold: true, color: C.green,
  });
  txt(s, '$ 1.970,12', {
    x: 5.02, y: 5.008, w: 2.079, h: 0.505, fontSize: 24, color: C.green, charSpacing: 1,
  });
  txt(s, T.confuseNoDot, {
    x: 6.864, y: 4.365, w: 2.079, h: 1.081, fontSize: 10, color: C.bodyGrey, lineSpacingMultiple: 1.5,
  });
}

function slide10(s) {
  title(s, 'Best Of Treatment For Your Skin', { x: 3.546, y: 0.387, w: 6.242, h: 1.582, align: 'center' });
  txt(s, T.confuse, {
    x: 2.921, y: 2.103, w: 7.492, h: 0.677, fontSize: 12, align: 'center', lineSpacingMultiple: 1.5,
  });
  const cards = [
    { fx: 0.724, px: 2.281, bx: 0.821, gx: 1.02, tw: 2.38, label: 'Skin And Facial Healths' },
    { fx: 4.87, px: 6.425, bx: 4.903, gx: 5.177, tw: 2.667, label: 'Best Serum Facial Treatment' },
    { fx: 8.946, px: 10.51, bx: 8.919, gx: 9.192, tw: 2.667, label: 'Ac Cure And Meso Treatment' },
  ];
  cards.forEach((c) => {
    outline(s, 'round2DiagRect', { x: c.fx, y: 3.155, w: 3.664, h: 1.53, flipV: true });
    pill(s, c.gx, 3.407, 3.03, 0.464, c.label);
    body(s, T.confuseTiny, { x: c.bx, y: 4.804, w: 1.592, h: 1.086, align: 'left' });
  });
}

/** Slides 11 and 31 — numbered "Start Today" columns. */
const START_TODAY = ['Ac Cure Skin', 'Facial Health', 'Treatment Skin', 'Serum Skincare'];

function startTodayColumn(s, x, y, index, label) {
  txt(s, String(index).padStart(2, '0') + '.', {
    x: x + 0.143, y, w: 1.432, h: 1.01, fontSize: 54, bold: true, color: C.cream,
  });
  txt(s, label, { x: x + 0.053, y: y + 1.01, w: 2.051, h: 0.337, fontSize: 14, color: C.cream });
  txt(s, 'Start Today 10.AM', {
    x: x + 0.034, y: y + 1.347, w: 2.352, h: 0.37, fontSize: 16, bold: true, color: C.cream,
  });
  body(s, T.confuseShort, {
    x, y: y + 1.83, w: 2.218, h: 1.086, align: 'left', color: C.cream,
  });
}

function slide11(s) {
  product(s, 10.19, 1.065, 2.909, 4.95, ART.pump);
  [0.366, 3.014, 5.715, 8.416].forEach((x) => {
    shape(s, 'roundRect', { x, y: 3.068, w: 2.427, h: 3.344, rectRadius: 0.28, fill: { color: C.green } });
  });
  title(s, 'Why Choose Us?', { x: 0.745, y: 1.142, w: 5.58, h: 0.841 });
  txt(s, T.brandShort, {
    x: 0.745, y: 2.094, w: 6.265, h: 0.677, fontSize: 12, align: 'justify', lineSpacingMultiple: 1.5,
  });
  pill(s, 7.608, 2.068, 1.786, 0.464, [{ text: 'Learn ' }, { text: 'More', options: { bold: true } }]);
  [0.452, 3.1, 5.801, 8.501].forEach((x, i) => startTodayColumn(s, x, 3.303, i + 1, START_TODAY[i]));
}

/** Slides 12, 23 and 32 — green cards with a check bullet and a star rating. */
function ratingCard(s, x, y, w, label, opts) {
  const o = opts || {};
  txt(s, label, { x: x + 0.383, y, w: 2.437, h: 0.37, fontSize: 16, color: C.white });
  iconCircle(s, x, y + 0.785, 0.26, '\u2714', { fill: C.cream, color: C.green, fontSize: 8 });
  body(s, T.confuseShort, {
    x: x + 0.383, y: y + 0.382, w: w, h: o.bodyH || 0.576, color: C.white, align: o.align || 'justify',
  });
  stars(s, x + 2.689, y + 1.076);
}

function slide12(s) {
  product(s, 0.937, 1.885, 1.484, 3.382, ART.bottle);
  product(s, 2.519, 1.374, 1.022, 3.893, ART.dropper);
  [[3.532, 2.86], [3.532, 4.722], [8.057, 2.91], [8.057, 4.772]].forEach(([x, y]) => {
    shape(s, 'round2DiagRect', { x, y, w: 4.338, h: 1.714, flipH: true, fill: { color: C.green } });
  });
  est(s, 3.778, 0.809, 1.375);
  title(s, 'All Of Skincare Treatment Plans', { x: 3.778, y: 1.146, w: 5.297, h: 1.447, fontSize: 40 });
  const cards = [
    [3.792, 3.023, 'Pedicure Skin'], [3.792, 4.884, 'Skin Facial Health'],
    [8.317, 3.072, 'Manicure Skin'], [8.317, 4.934, 'Skincare Serum'],
  ];
  cards.forEach(([x, y, label]) => ratingCard(s, x, y, 3.35, label));
}

function slide13(s) {
  title(s, 'The Expert Team Of Steale', { x: 0.741, y: 1.313, w: 3.908, h: 2.322 });
  txt(s, T.brand, {
    x: 0.741, y: 3.786, w: 3.908, h: 2.192, fontSize: 12, align: 'justify', lineSpacingMultiple: 1.5,
  });
  const names = [
    [10.333, 5.407, 'Juvilia', 'Yulae'], [10.333, 2.475, 'Velicia', 'Uita'],
    [6.344, 5.407, 'Fujjia', 'Uille'], [6.344, 2.475, 'Petricia', 'Brandy'],
  ];
  names.forEach(([x, y, first, last]) => {
    shape(s, 'rect', { x, y, w: 2.437, h: 0.7, fill: { color: C.green } });
    txt(s, [
      { text: first + ' ', options: { bold: true } },
      { text: last, options: { italic: true } },
    ], { x: x + 0.16, y, w: 2.099, h: 0.7, valign: 'middle', fontSize: 16, color: C.white });
  });
}

function slide14(s) {
  [[0.757, 2.772, 2.696, 1.303], [2.617, 4.608, 4.556, 3.139]].forEach(([bx, by, px, py]) => {
    outline(s, 'round2DiagRect', { x: bx, y: by, w: 3.442, h: 1.314, flipV: true });
  });
  title(s, 'The Expert Team Of Steale', { x: 7.34, y: 1.297, w: 4.95, h: 1.582 });
  txt(s, 'Collage No. 12', {
    x: 7.34, y: 3.101, w: 1.036, h: 0.773, fontSize: 14, bold: true, lineSpacingMultiple: 1.5,
  });
  txt(s, T.tagline, { x: 9.201, y: 3.237, w: 3.283, h: 0.656, fontSize: 11, color: C.green });
  txt(s, T.brand, {
    x: 7.34, y: 4.259, w: 5.064, h: 1.58, fontSize: 12, align: 'justify', lineSpacingMultiple: 1.5,
  });
  [[0.986, 2.879, 'Velicia Uita', 0.966, 3.216, 1.043, 3.646],
   [2.846, 4.715, 'Julieci Yuavita', 2.826, 5.051, 2.903, 5.482]].forEach(
    ([nx, ny, name, tx2, ty2, sx, sy]) => {
      txt(s, name, { x: nx, y: ny, w: 2.051, h: 0.337, fontSize: 14 });
      txt(s, 'Skincare Facial Team', { x: tx2, y: ty2, w: 2.623, h: 0.37, fontSize: 16, bold: true });
      stars(s, sx, sy);
    });
}

function slide15(s) {
  title(s, 'Angelie Youvae', { x: 1.098, y: 1.65, w: 5.486, h: 0.841 });
  const skills = [['Communication', 2.749, 1.213], ['Leadership', 3.351, 1.933],
                  ['Managements', 3.867, 0.804], ['Experience', 4.468, 1.501]];
  skills.forEach(([label, y, fill]) => {
    txt(s, label, { x: 1.257, y, w: 1.739, h: 0.286, fontSize: 11 });
    s.addShape('line', { x: 1.257, y: y + 0.329, w: 2.2, h: 0, line: { color: C.greyStar, width: 5.25 } });
    s.addShape('line', { x: 1.257, y: y + 0.329, w: fill, h: 0, line: { color: C.forest, width: 5.25 } });
  });
  body(s, T.confuse, { x: 1.098, y: 5.104, w: 3.062, h: 1.081 });
  body(s, T.confuse, { x: 4.617, y: 2.892, w: 3.062, h: 1.081 });
  ['Launching ', 'Promotions ', 'Strategies ', 'Planning '].forEach((label, i) => {
    const y = 4.195 + i * 0.4565;
    iconCircle(s, 4.616, y + 0.061, 0.26, '\u2714', { fill: C.forest, color: C.white, fontSize: 8 });
    txt(s, [
      { text: label, options: { bold: true } },
      { text: 'Brand', options: { italic: true } },
    ], { x: 4.999, y, w: 2.437, h: 0.37, fontSize: 16 });
  });
  pill(s, 9.361, 5.561, 1.786, 0.464, [{ text: 'Detail ' }, { text: 'More', options: { bold: true } }]);
}

function slide16(s) {
  shape(s, 'round2DiagRect', { x: 0.721, y: 1.156, w: 11.892, h: 4.648, flipH: true, fill: { color: C.green } });
  const items = [
    ['01', 1.298, 1.478, 3.067, 1.776, 'Promotions Brand'],
    ['02', 1.298, 3.502, 3.067, 3.799, 'Launching Products'],
    ['03', 6.767, 1.478, 8.536, 1.776, 'Selling The Brand'],
    ['04', 6.767, 3.502, 8.536, 3.799, 'Produce More'],
  ];
  items.forEach(([n, nx, ny, tx2, ty2, label]) => {
    txt(s, n, {
      x: nx, y: ny, w: 1.535, h: 1.212, fontSize: 66, bold: true,
      align: 'right', color: C.white, charSpacing: 1,
    });
    txt(s, label, { x: tx2, y: ty2, w: 2.808, h: 0.37, fontSize: 16, bold: true, color: C.white });
    body(s, T.confuse, { x: tx2, y: ty2 + 0.388, w: 3.135, h: 1.081, color: C.white });
  });
}

function slide17(s) {
  title(s, 'The Expert Service Of Steale Skincare', {
    x: 3.307, y: 0.412, w: 6.72, h: 1.582, align: 'center',
  });
  const cards = [
    [0.776, 2.355, '01.', 'Pedicure Skin'], [0.776, 4.352, '02.', 'Manicure Skin'],
    [4.618, 2.749, '03.', 'Skin Facial Health'], [4.618, 4.746, '04.', 'Skincare Serum'],
    [8.461, 2.355, '05.', 'AC Cure Treatment'], [8.461, 4.352, '06.', 'MesoTreatment'],
  ];
  cards.forEach(([x, y, n, label]) => {
    shape(s, 'round2DiagRect', { x, y, w: 3.579, h: 1.755, flipH: true, fill: { color: C.green } });
    txt(s, n, { x: x + 0.215, y, w: 1.224, h: 0.909, fontSize: 48, bold: true, color: C.cream });
    txt(s, label, { x: x + 1.262, y: y + 0.412, w: 2.043, h: 0.337, fontSize: 14, color: C.cream });
    body(s, T.confuseShort, {
      x: x + 0.263, y: y + 0.859, w: 3.316, h: 0.576, align: 'left', color: C.cream,
    });
  });
}

function slide18(s) {
  shape(s, 'round2DiagRect', { x: 9.445, y: 1.726, w: 3.085, h: 3.756, flipH: true, fill: { color: C.green } });
  s.addShape('line', {
    x: 9.471, y: 2.232, w: 0.497, h: 0,
    line: { color: C.white, width: 1.5, endArrowType: 'triangle' },
  });
  numberList(s, 9.947, 2.081, C.white);
  ['Pedicure Skin', 'Manicure Skin', 'Skin Facial Health', 'Skincare Serum'].forEach((label, i) => {
    txt(s, label, { x: 10.618, y: 2.068 + i * 0.353, w: 1.911, h: 0.303, fontSize: 12, color: C.white });
  });
  body(s, T.confuse, { x: 9.924, y: 3.597, w: 2.39, h: 1.591, color: C.white });
  title(s, 'The Expert Service Of Steale Skincare', { x: 0.741, y: 3.75, w: 6.72, h: 1.582 });
  pill(s, 7.124, 4.309, 1.786, 0.464, [{ text: 'Detail ' }, { text: 'More', options: { bold: true } }]);
}

function slide19(s) {
  shape(s, 'rect', { x: 0.432, y: 0.778, w: 12.469, h: 1.26, fill: { color: C.green } });
  [[0.929, '01'], [4.974, '02'], [8.817, '03']].forEach(([x, n], i) => {
    diamond(s, x, 0.901, n, C.blushLight, C.green);
    body(s, T.confuseNoDot, {
      x: 1.907 + i * 3.944, y: 1.003, w: 2.663, h: 0.829, align: 'left', color: C.cream,
    });
  });
  title(s, 'Daily Skincare Routine For Glowing And Brighten Skin', {
    x: 0.802, y: 2.451, w: 9.617, h: 1.582,
  });
  txt(s, T.brand, {
    x: 0.929, y: 4.222, w: 8.384, h: 0.98, fontSize: 12, align: 'justify', lineSpacingMultiple: 1.5,
  });
}

function slide20(s) {
  heroSlide(s, 20, 'Breaking Time', {
    ovalX: 1.278, ovalY: 1.842, ovalW: 10.777, ovalH: 3.816,
    estX: 3.419, estY: 2.755, titleX: 1.766, titleY: 2.7, titleW: 9.996, subY: 4.085,
  });
  [[0.73, 0.48, 0.33], [1.29, 0.3, 0.23], [12.26, 5.88, 0.33], [12.67, 5.52, 0.23], [11.8, 5.8, 0.23]]
    .forEach(([x, y, d]) => shape(s, 'star4', { x, y, w: d, h: d * 1.5, fill: { color: C.green } }));
}

function slide21(s) {
  shape(s, 'ellipse', { x: 9.4, y: -0.35, w: 4.6, h: 4.6, fill: { color: C.green } });
  title(s, 'New Products', { x: 1.302, y: 1.579, w: 4.765, h: 0.841 });
  est(s, 1.314, 2.558, 1.375);
  txt(s, T.brand, {
    x: 1.314, y: 3.057, w: 4.471, h: 1.883, fontSize: 12, align: 'justify', lineSpacingMultiple: 1.5,
  });
  arrowPill(s, 6.773, 4.629);
}

function slide22(s) {
  title(s, 'The Expert Galery Of Steale Skincare', { x: 6.868, y: 0.293, w: 5.25, h: 2.322 });
  const rows = [
    [2.889, '01', 3.128, 'Glowing Serum'], [3.873, '02', 4.111, 'Brighten Skincare'],
    [4.87, '03', 5.109, 'Moisturizing Cream'], [5.803, '04', 6.042, 'Anti Aging Cream'],
  ];
  rows.forEach(([dy, n, ty, label]) => {
    diamond(s, 6.987, dy, n, C.green, C.white);
    txt(s, label, { x: 7.813, y: ty, w: 1.942, h: 0.303, fontSize: 12, bold: true, color: C.green });
  });
}

function slide23(s) {
  title(s, 'The Galery Of Steale Skincare', { x: 1.621, y: 0.96, w: 10.225, h: 0.841, align: 'center' });
  [[3.545, 2.171, 3.711, 2.238, 'Pedicure Skin'], [3.545, 3.872, 3.69, 3.955, 'Manicure Skin']]
    .forEach(([bx, by, cx, cy, label]) => {
      shape(s, 'round2DiagRect', { x: bx, y: by, w: 4.048, h: 1.457, flipH: true, fill: { color: C.green } });
      ratingCard(s, cx, cy, 3.35, label);
    });
}

function slide24(s) {
  title(s, 'The Popular Products', { x: 0.98, y: 0.999, w: 4.163, h: 1.582 });
  [[6.051, 6.38, 'Brighten Skincare', 6.241, 7.815],
   [9.597, 9.926, 'Moisturizing Cream', 9.941, 11.36]].forEach(([bx, px, name, tx2, sx]) => {
    shape(s, 'round2SameRect', { x: bx, y: 1.568, w: 3.217, h: 4.614, fill: { color: C.green } });
    txt(s, name, { x: tx2 + 0.019, y: 4.774, w: 2.051, h: 0.337, fontSize: 14, color: C.white });
    txt(s, 'Skincare Facial Event', { x: tx2, y: 5.11, w: 2.658, h: 0.37, fontSize: 16, bold: true, color: C.white });
    stars(s, sx, 5.654);
  });
  txt(s, 'Glowing Serum', { x: 1.013, y: 2.914, w: 2.051, h: 0.337, fontSize: 14, color: C.green });
  txt(s, 'Skincare Facial Event', { x: 0.994, y: 3.251, w: 2.812, h: 0.37, fontSize: 16, bold: true });
  stars(s, 2.568, 3.794);
  txt(s, T.brand, {
    x: 1.014, y: 4.169, w: 4.471, h: 1.883, fontSize: 12, align: 'justify', lineSpacingMultiple: 1.5,
  });
}

function slide25(s) {
  title(s, 'Relaxing Skincare', { x: 2.083, y: 1.077, w: 5.921, h: 0.841 });
  [[2.154, 2.135, 'Ac Cure Skin', 2.432, 2.768, 3.513],
   [4.439, 4.42, 'Facial Health', 4.702, 5.039, 5.855]].forEach(
    ([by, py, label, ny, sy, cy]) => {
      outline(s, 'round2DiagRect', { x: 2.176, y: by, w: 3.442, h: 1.314, flipV: true });
      txt(s, label, { x: 2.616, y: ny, w: 2.051, h: 0.337, fontSize: 14, color: C.green });
      txt(s, 'Start Today 10.AM', { x: 2.597, y: sy, w: 2.352, h: 0.37, fontSize: 16, bold: true, color: C.green });
      body(s, T.confusePromote, { x: 2.206, y: cy, w: 3.381, h: 0.834 });
    });
}

function slide26(s) {
  shape(s, 'rect', { x: 3.486, y: -0.013, w: 2.494, h: 7.513, fill: { color: C.green } });
  [[0.815, 0.992, 1.23, 1.863, '01', 'Glowing Serum'],
   [3.762, 3.963, 4.201, 4.797, '02', 'Brighten Skincare']].forEach(
    ([py, dy, ty, by, n, label]) => {
      outline(s, 'round2DiagRect', { x: 6.667, y: py, w: 3.442, h: 2.672, flipV: true });
      diamond(s, 7.074, dy, n, C.green, C.white);
      txt(s, label, { x: 7.9, y: ty, w: 1.815, h: 0.303, fontSize: 12, bold: true, color: C.green });
      body(s, T.confuse, { x: 7.093, y: by, w: 2.623, h: 1.334 });
    });
  title(s, 'The Expert Event Now', { x: -0.036, y: 2.431, w: 3.984, h: 1.582, rotate: 270 });
}

function slide27(s) {
  shape(s, 'ellipse', { x: 1.074, y: 0.489, w: 3.638, h: 3.912, fill: { color: C.green } });
  shape(s, 'roundRect', { x: 1.68, y: 0.45, w: 1.95, h: 6.4, rectRadius: 0.4, fill: { color: 'F4F4F4' } });
  photo(s, 1.16, 1.72, 3.0, 3.6, { rectRadius: 0.55, pad: 0.24, body: 'BDBDBD', screenRadius: 0.4 });
  title(s, 'Watch Mockup', { x: 5.49, y: 1.343, w: 6.233, h: 1.01, fontSize: 54 });
  txt(s, 'Collage No. 12', {
    x: 5.49, y: 2.489, w: 1.036, h: 0.773, fontSize: 14, bold: true, lineSpacingMultiple: 1.5,
  });
  txt(s, T.tagline, { x: 7.911, y: 2.625, w: 3.283, h: 0.656, fontSize: 11, color: C.green });
  txt(s, T.brand, {
    x: 5.49, y: 3.647, w: 5.704, h: 1.58, fontSize: 12, align: 'justify', lineSpacingMultiple: 1.5,
  });
}

function slide28(s) {
  shape(s, 'rect', { x: 0, y: 4.506, w: W, h: 1.26, fill: { color: C.green } });
  photo(s, 7.34, 1.9, 2.32, 4.15, { rectRadius: 0.2, pad: 0.16, padY: 0.5, body: 'F7F7F7', rotate: 353 });
  photo(s, 9.72, 1.05, 2.66, 4.76, { rectRadius: 0.2, pad: 0.18, padY: 0.55, body: 'F7F7F7', rotate: 6 });
  reviewsSlide(s, { tx: 1.297, ty: 1.231, qx: 7.406, qy: 0.71, cx: 1.297, cy: 2.315 });
}

function slide29(s) {
  shape(s, 'round2DiagRect', { x: 0.506, y: 1.726, w: 7.667, h: 3.756, flipH: true, fill: { color: C.green } });
  shape(s, 'roundRect', { x: 6.1, y: 1.5, w: 1.7, h: 4.1, rectRadius: 0.35, rotate: 28, fill: { color: 'D3D9DC' } });
  photo(s, 5.42, 2.0, 2.6, 2.6, { rectRadius: 0.5, pad: 0.2, body: 'A9AEB2', screenRadius: 0.36, rotate: 28 });
  title(s, 'Best Of Treatment For Your Skin', { x: 0.97, y: 2.207, w: 5.571, h: 2.322, color: C.white });
  [[9.226, 2.237, 'Skin And Facial Healths', 9.389, 2.792],
   [9.286, 3.695, 'Best Serum Facial Treatment', 9.489, 4.25]].forEach(([x, y, label, bx, by]) => {
    pill(s, x, y, 3.03, 0.464, label);
    body(s, T.confuseTiny, { x: bx, y: by, w: 2.623, h: 0.576, align: 'center' });
  });
}

function slide30(s) {
  photo(s, 7.55, 1.6, 2.35, 4.1, {
    rectRadius: 0.2, pad: 0.14, padY: 0.42, body: 'F4EDE9', screen: 'FDFDFD', label: 'B9B9B9', rotate: 342,
  });
  photo(s, 9.9, 0.7, 2.35, 4.6, {
    rectRadius: 0.2, pad: 0.14, padY: 0.45, body: 'F4EDE9', screen: 'FDFDFD', label: 'B9B9B9', rotate: 17,
  });
  promotionSlide(s, {
    tx: 1.251, ty: 1.39, titleColor: C.green, rx: 1.446, ry: 2.509, textColor: C.green,
    sx: 1.446, sy: 4.303, accent: C.green, cx: 3.709, cy: 4.353, bx: 4.597, px: 1.506, py: 5.202,
  });
}

function slide31(s) {
  shape(s, 'rect', { x: 4.395, y: 0, w: 8.938, h: 5.778, fill: { color: C.green } });
  photo(s, 0.6, 1.9, 4.65, 3.45, {
    rectRadius: 0.12, pad: 0.12, body: 'E4E4E4', screen: 'FCFCFC', label: 'B9B9B9',
  });
  shape(s, 'roundRect', { x: 0.29, y: 5.35, w: 5.3, h: 0.42, rectRadius: 0.08, fill: { color: 'DCDCDC' } });
  title(s, 'The Expert Mockup Of Steale Skincare', { x: 5.34, y: 0.391, w: 6.4, h: 1.582, color: C.white });
  [5.264, 7.913, 10.613].forEach((x, i) => startTodayColumn(s, x, 2.316, i + 1, START_TODAY[i]));
}

function slide32(s) {
  shape(s, 'roundRect', { x: 0.2, y: 1.6, w: 4.6, h: 5.9, rectRadius: 0.7, fill: { color: 'F2CFC2' } });
  photo(s, 1.15, 0.9, 3.2, 5.5, {
    rectRadius: 0.28, pad: 0.16, padY: 0.55, body: 'FAFAFA', screen: 'F6E1DA', label: 'C4A79C',
  });
  [5.288, 9.041].forEach((x) => {
    shape(s, 'round2DiagRect', { x, y: 3.295, w: 3.23, h: 2.718, flipH: true, fill: { color: C.green } });
  });
  title(s, 'The Popular Products', { x: 5.211, y: 0.737, w: 7.132, h: 0.841 });
  txt(s, T.brand, {
    x: 5.211, y: 1.724, w: 7.301, h: 1.283, fontSize: 12, align: 'justify', lineSpacingMultiple: 1.5,
  });
  [[5.425, 3.7, 'Pedicure Skin'], [9.293, 3.75, 'Manicure Skin']].forEach(([x, y, label]) => {
    ratingCard(s, x, y, 2.437, label, { bodyH: 0.834, align: 'left' });
  });
}

function slide33(s) {
  shape(s, 'rect', { x: 0, y: 2.691, w: W, h: 3.889, fill: { color: C.green } });
  photo(s, 4.9, 3.2, 8.2, 4.1, {
    rectRadius: 0.2, pad: 0.3, body: 'F1F1F1', screen: 'FAFAFA', label: 'B9B9B9', rotate: 341,
  });
  title(s, 'The Expert Mockup Of Steale Skincare', { x: 0.775, y: 2.959, w: 7.023, h: 1.582, color: C.white });
  txt(s, T.brand, {
    x: 0.775, y: 4.737, w: 7.301, h: 1.283, fontSize: 12, color: C.white,
    align: 'justify', lineSpacingMultiple: 1.5,
  });
}

function slide34(s) {
  title(s, 'The Facial Pricing Plans', { x: 2.863, y: 1.074, w: 7.922, h: 0.841, align: 'center' });
  /** Price card: [cardX, cardY, cardW, cardH, tier, price, bodyX, bodyY, bodyH, tall] */
  const plans = [
    [1.238, 2.268, 2.606, 3.711, '\u201cBasic ', '9', 1.376, 4.005, 1.591, true],
    [9.621, 2.268, 2.606, 3.711, '\u201cExpert ', '24', 9.759, 4.005, 1.591, true],
    [4.243, 2.268, 4.98, 1.763, '\u201cMedium ', '10', 6.797, 2.478, 1.086, false],
    [4.243, 4.24, 4.98, 1.763, '\u201cPremium ', '20', 6.797, 4.45, 1.086, false],
  ];
  plans.forEach(([x, y, w, h, tier, price, bx, by, bh, tall]) => {
    shape(s, 'roundRect', { x, y, w, h, rectRadius: 0.18, fill: { color: C.green } });
    const labelX = tall ? x + 0.137 : x - 0.053;
    const labelW = tall ? 2.325 : 2.614;
    txt(s, [
      { text: tier, options: { bold: true } },
      { text: 'Facial\u201d' },
    ], { x: labelX, y: y + 0.308, w: labelW, h: 0.404, fontSize: 18, align: 'center', color: C.cream });
    txt(s, [
      { text: '$' + price, options: { fontSize: 40, bold: true } },
      { text: '/', options: { fontSize: 36, bold: true } },
      { text: 'Day', options: { fontSize: 18 } },
    ], {
      x: x + 0.19, y: y + 0.723, w: 2.083, h: 0.774,
      align: 'center', wrap: false, margin: 0, color: C.cream,
    });
    body(s, tall ? T.confuse : T.confuseTools, {
      x: bx, y: by, w: 2.323, h: bh, color: C.cream, align: tall ? 'center' : 'left',
    });
  });
}

function slide35(s) {
  const cards = [
    [1.688, 0.987, 'Basic ', '8', 1.928, 1.277, 4.1, 2.55, 4.258, 1.741],
    [6.085, 0.987, 'Medium ', '10', 6.325, 1.277, 8.497, 2.55, 8.655, 1.741],
    [1.688, 3.75, 'Premium ', '16', 1.928, 4.04, 4.1, 5.314, 4.258, 4.504],
    [6.085, 3.75, 'V.I.P ', '24', 6.325, 4.04, 8.497, 5.314, 8.655, 4.504],
  ];
  const perks = ['Pedicure Skin', 'Skin Facial Health', 'Manicure Skin', 'Skincare Serum'];
  cards.forEach(([cx, cy, tier, price, tx2, ty2, bx, by, px, py]) => {
    shape(s, 'roundRect', { x: cx, y: cy, w: 4.194, h: 2.585, rectRadius: 0.2, fill: { color: C.green } });
    txt(s, [
      { text: tier, options: { bold: true } },
      { text: 'Facial' },
    ], { x: tx2, y: ty2, w: 2.437, h: 0.37, fontSize: 16, color: C.cream });
    perks.forEach((p, i) => {
      txt(s, p, {
        x: tx2, y: ty2 + 0.491 + i * 0.372, w: 2.172, h: 0.303, fontSize: 12, color: C.cream,
        bullet: { characterCode: '2713', indent: 18 },
      });
    });
    txt(s, '$' + price, {
      x: px, y: py, w: 1.274, h: 0.774, fontSize: 40, bold: true,
      align: 'center', wrap: false, margin: 0, color: C.cream,
    });
    pill(s, bx, by, 1.567, 0.464, 'Book Now', { fill: C.cream, color: C.green });
  });
  title(s, 'The Facial Pricing Plans', {
    x: 9.043, y: 2.959, w: 4.434, h: 1.582, align: 'center', rotate: 90,
  });
}

function slide36(s) {
  title(s, 'Strength Skincare Facial', { x: 5.255, y: 1.415, w: 5.007, h: 1.582 });
  txt(s, T.brand, {
    x: 5.255, y: 3.319, w: 6.27, h: 1.283, fontSize: 12, align: 'justify', lineSpacingMultiple: 1.5,
  });
  txt(s, T.brandShort, {
    x: 5.255, y: 4.893, w: 6.27, h: 0.773, fontSize: 14, align: 'justify', lineSpacingMultiple: 1.5,
  });
}

function slide37(s) {
  shape(s, 'round2DiagRect', { x: 0.506, y: 2.675, w: 7.667, h: 3.756, flipH: true, fill: { color: C.green } });
  title(s, 'Weakness Brand', { x: 0.568, y: 1.755, w: 6.298, h: 0.841 });
  txt(s, T.brand, {
    x: 0.988, y: 2.98, w: 5.338, h: 1.586, fontSize: 12, color: C.white,
    align: 'justify', lineSpacingMultiple: 1.5,
  });
  for (let i = 0; i < 4; i++) {
    insertRow(s, 1.13 + (i % 2) * 2.505, 4.901 + Math.floor(i / 2) * 0.646, C.white);
  }
}

function slide38(s) {
  title(s, 'Opportunity Brand', { x: 6.15, y: 2.131, w: 6.225, h: 0.841 });
  txt(s, T.brand, {
    x: 6.15, y: 3.182, w: 6.27, h: 1.283, fontSize: 12, align: 'justify', lineSpacingMultiple: 1.5,
  });
  txt(s, T.brandShort, {
    x: 6.15, y: 4.756, w: 6.27, h: 0.773, fontSize: 14, align: 'justify', lineSpacingMultiple: 1.5,
  });
}

function slide39(s) {
  title(s, 'Threats Skincare Facial', { x: 0.955, y: 2.776, w: 7.77, h: 0.841 });
  [[1.15, 4.048], [4.927, 4.048], [1.15, 5.068], [4.927, 5.068]].forEach(([x, y]) => {
    txt(s, '\u2714', { x: x - 0.06, y: y - 0.09, w: 0.3, h: 0.3, fontSize: 11, color: C.green });
    body(s, T.confuseShort, { x: x + 0.264, y: y - 0.135, w: 3.321, h: 0.576 });
  });
}

/** Dotted progress ring used for the KPI donuts on slide 40. */
function dottedRing(slide, cx, cy, r, pct, onColor, offColor) {
  const dots = 24;
  const d = 0.155;
  for (let i = 0; i < dots; i++) {
    const a = (-90 + (360 / dots) * i) * (Math.PI / 180);
    shape(slide, 'ellipse', {
      x: cx + Math.cos(a) * r - d / 2, y: cy + Math.sin(a) * r - d / 2,
      w: d, h: d * 1.25,
      rotate: (360 / dots) * i,
      fill: { color: i / dots < pct ? onColor : offColor },
    });
  }
}

function slide40(s) {
  shape(s, 'roundRect', { x: 0.415, y: 0.333, w: 12.503, h: 6.414, rectRadius: 0.35, fill: { color: C.green } });
  title(s, 'Charts Brand', { x: 3.197, y: 0.713, w: 4.699, h: 0.841, color: C.cream });
  const rings = [
    { x: 1.328, y: 1.467, pct: 0.86, label: '86%', data: 'Data 01', on: C.blush, off: 'A98D84', tx: 1.155, ty: 3.915, bx: 0.978, by: 4.191, bw: 2.621 },
    { x: 4.272, y: 2.078, pct: 0.65, label: '65%', data: 'Data 02', on: C.blushLight, off: '8E9E8F', tx: 4.099, ty: 4.526, bx: 3.971, by: 4.803, bw: 2.552 },
    { x: 7.019, y: 1.519, pct: 0.96, label: '96%', data: 'Data 01', on: C.blush, off: 'A98D84', tx: 6.806, ty: 3.968, bx: 6.71, by: 4.244, bw: 2.552 },
    { x: 9.944, y: 1.988, pct: 0.35, label: '35%', data: 'Data 02', on: C.blushLight, off: '8E9E8F', tx: 9.731, ty: 4.436, bx: 9.635, by: 4.803, bw: 2.552 },
  ];
  rings.forEach((r) => {
    dottedRing(s, r.x + 0.967, r.y + 0.967, 0.89, r.pct, r.on, r.off);
    txt(s, r.label, {
      x: r.x + 0.18, y: r.y + 0.58, w: 1.575, h: 0.64, fontSize: 32, bold: true,
      align: 'center', color: C.cream,
    });
    txt(s, r.data, {
      x: r.x + 0.18, y: r.y + 1.1, w: 1.575, h: 0.303, fontSize: 12, align: 'center', color: C.cream,
    });
    txt(s, 'Insert Text Here', {
      x: r.tx, y: r.ty, w: 2.361, h: 0.303, fontSize: 12, bold: true, align: 'center', color: C.cream,
    });
    body(s, T.confuse, { x: r.bx, y: r.by, w: r.bw, h: 1.334, align: 'center', color: C.cream });
  });
}

const CHART_PCTS = [65, 63, 97, 91, 24];

function slide41(s) {
  title(s, 'Charts Brand', { x: 1.554, y: 0.66, w: 5.571, h: 0.841, color: C.green });
  const trackW = 4.136, barH = 0.429;
  CHART_PCTS.forEach((p, i) => {
    const y = 1.96 + i * 0.689;
    shape(s, 'roundRect', {
      x: 1.553, y, w: trackW, h: barH, rectRadius: barH / 2, fill: { color: C.track },
    });
    shape(s, 'roundRect', {
      x: 1.553, y, w: trackW * (p / 100), h: barH, rectRadius: barH / 2,
      fill: { color: i % 2 ? C.sage : C.green },
    });
    txt(s, p + '%', {
      x: 5.58, y: y + 0.008, w: 0.974, h: 0.329, fontSize: 12, align: 'center', color: C.green,
    });
  });
  const rows = [
    [1.902, 2.101, 2.394, 'Glowing Serum'], [2.911, 3.189, 3.506, 'Brighten Skincare'],
    [3.942, 4.186, 4.511, 'Moisturizing Cream'], [5.028, 5.274, 5.601, 'Anti Aging Cream'],
  ];
  rows.forEach(([dy, ty, by, label], i) => {
    diamond(s, 6.965, dy, '0' + (i + 1), C.green, C.white);
    txt(s, label, { x: 7.759, y: ty, w: 1.942, h: 0.303, fontSize: 12, bold: true, color: C.green });
    body(s, T.confuseShort, { x: 7.759, y: by, w: 3.753, h: 0.576 });
  });
  arrowPill(s, 4.089, 5.496);
}

function slide42(s) {
  title(s, 'Charts Brand', { x: 1.589, y: 1.139, w: 4.699, h: 0.841, color: C.green });
  const trackTop = 2.869, trackBot = 5.816, barW = 0.431;
  CHART_PCTS.forEach((p, i) => {
    const x = 7.009 + i * 1.1225;
    shape(s, 'roundRect', {
      x, y: trackTop, w: barW, h: trackBot - trackTop, rectRadius: barW / 2, fill: { color: C.track },
    });
    const h = (trackBot - trackTop) * (p / 100);
    shape(s, 'roundRect', {
      x, y: trackBot - h, w: barW, h, rectRadius: barW / 2, fill: { color: i % 2 ? C.sage : C.green },
    });
    txt(s, p + '%', {
      x: x - 0.131, y: 2.392, w: 0.693, h: 0.33, fontSize: 12, align: 'center', color: C.forest,
    });
  });
  [1.977, 3.341, 4.767].forEach((y, i) => {
    txt(s, '0' + (i + 1), {
      x: 1.261, y, w: 1.535, h: 1.212, fontSize: 66, bold: true, align: 'right',
      color: C.forest, charSpacing: 1,
    });
    txt(s, 'Skincare Facial Chart', { x: 3.03, y: y + 0.298, w: 2.835, h: 0.37, fontSize: 16, bold: true });
    body(s, T.confuseShort, { x: 3.03, y: y + 0.686, w: 3.626, h: 0.576 });
  });
}

function slide43(s) {
  title(s, 'Charts Brand', { x: 3.387, y: 0.746, w: 6.233, h: 0.841, align: 'center' });
  // Segmented ring built from four pie quadrants plus two blush side wedges.
  s.addChart('doughnut', [{
    name: 'Segments', labels: ['a', 'b', 'c', 'd', 'e', 'f'], values: [1, 1, 1, 1, 1, 1],
  }], {
    x: 4.318, y: 1.91, w: 4.33, h: 4.33, holeSize: 45,
    showLegend: false, showValue: false, showTitle: false,
    chartColors: ['4D6F51', 'FDC2AA', '3A533D', '4D6F51', 'FDC2AA', '3A533D'],
    dataBorder: { pt: 3, color: C.pageBottom },
  });
  shape(s, 'ellipse', { x: 5.673, y: 3.265, w: 1.62, h: 1.62, fill: { color: 'BFBFBF' } });
  const items = [
    [0.752, 2.053, 1.568, '01'], [0.752, 4.093, 1.568, '02'],
    [8.926, 2.053, 9.742, '03'], [8.926, 4.093, 9.742, '04'],
  ];
  items.forEach(([cx, cy, tx2, n]) => {
    iconCircle(s, cx, cy, 0.754, n, { fill: C.forest, color: C.cream, fontSize: 16 });
    txt(s, [{ text: 'Treatment ' }, { text: 'Ellita' }], {
      x: tx2 + 0.02, y: cy + 0.076, w: 2.051, h: 0.337, fontSize: 14, color: C.body,
    });
    txt(s, 'Skincare Facial Chart', { x: tx2, y: cy + 0.413, w: 2.839, h: 0.37, fontSize: 16, bold: true });
    body(s, T.confuseShort, { x: tx2 + 0.02, y: cy + 0.915, w: 2.659, h: 0.829 });
  });
}

function slide44(s) {
  shape(s, 'round2DiagRect', { x: 0.835, y: 1.221, w: 8.943, h: 1.429, fill: { color: C.green } });
  title(s, 'Contact Skincare', {
    x: 1.761, y: 1.47, w: 6.786, h: 0.841, bold: true, color: C.cream,
  });
  const blocks = [
    { icon: '\u27a4', ix: 1.88, iy: 3.093, hx: 2.515, hy: 3.045, hw: 2.454, head: 'Address Company',
      lines: ['76 United Street Ca,', 'California, CA 62 71593'], lx: 2.515, ly: 3.28, lw: 2.479 },
    { icon: '\u2709', ix: 5.249, iy: 3.064, hx: 5.879, hy: 3.045, hw: 2.145, head: 'Email Contact ',
      lines: ['@office.official.com', '@steale.office.com'], lx: 5.845, ly: 3.288, lw: 2.389 },
    { icon: '\u2706', ix: 1.88, iy: 4.327, hx: 2.506, hy: 4.357, hw: 2.088, head: 'Phone Contact ',
      lines: ['+09 235 854 787', '+65 235 854 787'], lx: 2.505, ly: 4.589, lw: 1.939 },
  ];
  blocks.forEach((b) => {
    iconCircle(s, b.ix, b.iy, 0.473, b.icon, { fontSize: 13 });
    txt(s, b.head, { x: b.hx, y: b.hy, w: b.hw, h: 0.37, fontSize: 16, bold: true });
    txt(s, b.lines.map((line) => ({ text: line, options: { breakLine: true } })), {
      x: b.lx, y: b.ly, w: b.lw, h: 0.42 + (b.lines.length - 1) * 0.305,
      fontSize: 14, color: C.body, lineSpacingMultiple: 1.5,
    });
  });
  body(s, T.confuse, { x: 5.367, y: 4.354, w: 2.995, h: 1.339 });
}

function slide45(s) {
  heroSlide(s, 45, 'Thank You', {
    ovalX: 1.073, ovalY: 0.901, ovalW: 10.878, ovalH: 4.54,
    estX: 3.257, estY: 2.07, titleX: 2.804, titleY: 2.168, titleW: 7.921, subY: 3.554,
  });
  txt(s, 'Connect with us:', {
    x: 1.656, y: 5.339, w: 2.186, h: 0.464, fontSize: 16, lineSpacingMultiple: 1.5,
  });
  [[1.706, '\uf09a'], [5.164, '\u2691'], [8.48, '\u25ce']].forEach(([x, glyph], i) => {
    shape(s, 'ellipse', {
      x, y: 5.883, w: 0.278, h: 0.278, fill: { type: 'none' }, line: { color: C.forest, width: 1 },
    });
    txt(s, ['f', 't', 'o'][i], {
      x, y: 5.883, w: 0.278, h: 0.278, align: 'center', valign: 'middle', fontSize: 9,
    });
    txt(s, [{ text: '@' }, { text: 'Account.Name' }], {
      x: x + 0.368, y: 5.858, w: 2.94, h: 0.303, fontSize: 12, charSpacing: 6, color: C.body,
    });
  });
}

const BUILDERS = [
  slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
  slide31, slide32, slide33, slide34, slide35, slide36, slide37, slide38, slide39, slide40,
  slide41, slide42, slide43, slide44, slide45,
];

/* ------------------------------------------------------------------ */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: W, height: H });
  pptx.layout = 'WIDE';
  pptx.author = 'Steale Beauty';
  pptx.title = 'Steale Beauty — The Newest Skincare Products';

  BUILDERS.forEach((builder, i) => {
    const slide = pptx.addSlide();
    background(slide);
    (SPRIGS[i + 1] || []).forEach(([x, y, len, a]) => sprig(slide, x, y, len, a));
    builder(slide);
    footer(slide, i + 1);
  });

  return pptx.writeFile({
    fileName: path.join(__dirname, '015792b6-6646-4119-86e5-1b3d2449d779_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => {
  console.error(e);
  process.exit(1);
});
