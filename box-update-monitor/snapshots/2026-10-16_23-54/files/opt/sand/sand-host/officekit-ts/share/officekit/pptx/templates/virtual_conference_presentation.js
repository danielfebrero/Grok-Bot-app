/**
 * "Conference Presentation Template" — 25 slides, 13.333 x 7.5 in.
 * Rebuilt from scratch with pptxgenjs only. Run: node <thisfile>.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Theme
 * ------------------------------------------------------------------ */

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

const INK = '404040'; // body / heading text (tx1 lumMod 75%)
const TEAL = '0DA7A7'; // accent1 lumMod 50%  – footer text
const DEEP_TEAL = '1391AF'; // accent2 lumMod 50%  – numerals / icons
const ROSE = 'CB77A1'; // accent4 lumMod 75%  – numerals / icons
const INDIGO = '837AD9'; // accent3 lumMod 75%  – numerals
const MUTED = '808080'; // image-placeholder caption
const HAIRLINE = 'D9D9D9';

const HEAD_FONT = 'Barlow'; // theme major latin
const BODY_FONT = 'Darker Grotesque'; // theme minor latin

// Big flat background circles.
const MIST = 'E3FDFD'; // accent1 tint, outer halo
const HAZE = 'C8FBFB'; // accent1 tint, inner halo

/** Linear gradients used all over the deck (pptxgenjs fills are solid, so we blend). */
const GRADIENT = {
  cyan: ['90E0F3', '75F4F4'], // accent2 -> accent1
  pink: ['D5D2F2', 'E9C5D7'], // accent3 -> accent4
  violet: ['B4AFE8', 'D5D2F2'], // accent3 lumMod90 -> accent3
  paper: ['F2F2F2', 'FFFFFF'], // bg1 lumMod95 -> bg1
};

/** Hairline stroke that goes with each gradient. */
const GRADIENT_EDGE = { cyan: 'E3FDFD', pink: 'F2F2F2', violet: 'F2F2F2', paper: 'FFFFFF' };

function blend(hexA, hexB) {
  let out = '';
  for (let i = 0; i < 6; i += 2) {
    const v = Math.round((parseInt(hexA.substr(i, 2), 16) + parseInt(hexB.substr(i, 2), 16)) / 2);
    out += v.toString(16).toUpperCase().padStart(2, '0');
  }
  return out;
}

/** fill+line pair for one of the four deck gradients. */
function skin(tone) {
  const [a, b] = GRADIENT[tone];
  return { fill: { color: blend(a, b) }, line: { color: GRADIENT_EDGE[tone], width: 1 } };
}

/** The one soft drop shadow this template uses everywhere. */
const SOFT_SHADOW = { type: 'outer', angle: 56, blur: 40, offset: 27, color: '404040', opacity: 0.1 };

/** roundRect adjust value (fraction of the short side) -> corner radius in inches. */
function radius(w, h, adj) {
  return Math.min(w, h) * adj;
}

/* ------------------------------------------------------------------ *
 * Text helpers
 * ------------------------------------------------------------------ */

const BODY_SIZE = 16;

/** 36 pt Barlow section heading (68 pt on the cover / closing slides). */
function heading(slide, x, y, w, h, text, opts = {}) {
  slide.addText(text, {
    x, y, w, h, valign: 'top', align: opts.align || 'left',
    fontFace: HEAD_FONT, fontSize: opts.fontSize || 36, bold: true, color: INK,
    charSpacing: -1.5, lineSpacingMultiple: 0.9,
  });
}

/** 16 pt Darker Grotesque paragraph (1.3 line spacing, 6 pt after). */
function paragraph(slide, x, y, w, h, text, opts = {}) {
  slide.addText(text, {
    x, y, w, h, valign: 'top', align: opts.align || 'left',
    fontFace: BODY_FONT, fontSize: opts.fontSize || BODY_SIZE, bold: !!opts.bold,
    color: opts.color || INK, lineSpacingMultiple: opts.lineSpacing || 1.3, paraSpaceAfter: 6,
  });
}

/** Bold 16 pt kicker that sits above most paragraphs. */
function kicker(slide, x, y, w, text, opts = {}) {
  paragraph(slide, x, y, w, 0.422, text, Object.assign({ bold: true }, opts));
}

/** Wingdings-style "↘" bulleted list. */
function bulletList(slide, x, y, w, h, lines) {
  const bullet = { characterCode: '2198', indent: 22.5 };
  slide.addText(lines.map((t) => ({ text: t, options: { bullet, breakLine: true } })), {
    x, y, w, h, valign: 'top', bullet,
    fontFace: BODY_FONT, fontSize: BODY_SIZE, color: INK,
    lineSpacingMultiple: 1.3, paraSpaceAfter: 6,
  });
}

/** Oversized stat number ("500+", "$400.000", "01."). */
function stat(slide, x, y, w, text, opts = {}) {
  slide.addText(text, {
    x, y, w: w, h: 0.824, valign: 'top', align: opts.align || 'left',
    fontFace: opts.fontFace || BODY_FONT, fontSize: opts.fontSize || 36, bold: true,
    color: opts.color || INK, lineSpacingMultiple: 1.3, paraSpaceAfter: 6,
  });
}

/* ------------------------------------------------------------------ *
 * Shape helpers
 * ------------------------------------------------------------------ */

/** Decorative gradient bubble. */
function bubble(slide, x, y, d, tone) {
  slide.addShape('ellipse', Object.assign({ x, y, w: d, h: d, shadow: SOFT_SHADOW }, skin(tone)));
}

/** Flat halo circle (no shadow) — the pale background discs. */
function halo(slide, x, y, d, color, outlined) {
  slide.addShape('ellipse', {
    x, y, w: d, h: d, fill: { color },
    line: outlined ? { color: MIST, width: 1 } : { type: 'none' },
  });
}

/** Gradient rounded-rectangle card. */
function card(slide, x, y, w, h, adj, tone, extra = {}) {
  slide.addShape('roundRect', Object.assign({
    x, y, w, h, rectRadius: radius(w, h, adj), shadow: SOFT_SHADOW,
  }, skin(tone), extra));
}

/** 1.125 x 1.077 rounded tile that hosts an icon. */
function iconTile(slide, x, y, tone) {
  card(slide, x, y, 1.125, 1.077, 0.31655, tone);
}

/**
 * Stand-in for one of the deck's line-art icons (the originals are embedded
 * PNG/SVG art, which we replace with a same-sized outlined glyph box).
 */
function icon(slide, x, y, size, color) {
  slide.addShape('roundRect', {
    x, y, w: size, h: size, rectRadius: size * 0.18,
    fill: { type: 'none' }, line: { color, width: 1.25 },
  });
  slide.addShape('ellipse', {
    x: x + size * 0.31, y: y + size * 0.31, w: size * 0.38, h: size * 0.38,
    fill: { type: 'none' }, line: { color, width: 1.25 },
  });
}

/**
 * Stand-in for one of the deck's picture placeholders: the same "Image
 * Placeholder" caption the layout carries, plus a small photo glyph centred in
 * the frame (the reference art itself is a raster image we do not embed).
 */
function imageFrame(slide, x, y, w, h) {
  slide.addText('Image Placeholder', {
    x, y: y + 0.01, w, h: 0.2, align: 'center', valign: 'top',
    fontFace: BODY_FONT, fontSize: 12, color: MUTED,
  });
  const gw = 0.883;
  const gh = 0.700;
  const gx = x + (w - gw) / 2;
  const gy = y + (h - gh) / 2;
  slide.addShape('rect', {
    x: gx, y: gy, w: gw, h: gh, fill: { color: 'FAFAFA' }, line: { color: '3A3A38', width: 1 },
  });
  slide.addShape('rect', {
    x: gx + 0.06, y: gy + 0.05, w: gw - 0.12, h: gh - 0.10,
    fill: { color: 'FFFFFF' }, line: { color: 'B9B8B7', width: 0.75 },
  });
  slide.addShape('ellipse', {
    x: gx + 0.15, y: gy + 0.13, w: 0.15, h: 0.15,
    fill: { color: 'F8DB8F' }, line: { color: 'F0A93B', width: 1 },
  });
  slide.addShape('triangle', {
    x: gx + 0.12, y: gy + 0.38, w: 0.34, h: 0.22, fill: { color: '83BEEC' },
  });
  slide.addShape('triangle', {
    x: gx + 0.31, y: gy + 0.25, w: 0.51, h: 0.35, fill: { color: '83BEEC' },
  });
}

/** White arrow-head connector used inside pills and small tiles. */
function tinyArrow(slide, x, y, w, color) {
  slide.addShape('line', {
    x, y, w, h: 0,
    line: { color: color || 'FFFFFF', width: 1, endArrowType: 'triangle' },
  });
}

/** Freeform path helper: `pts` are normalised 0..1 coordinates. */
function freeform(slide, x, y, w, h, pts, opts = {}) {
  const at = (p) => ({ x: p[0] * w, y: p[1] * h });
  const points = pts.map((p) => {
    if (p.length === 6) {
      return {
        x: p[4] * w, y: p[5] * h,
        curve: { type: 'cubic', x1: p[0] * w, y1: p[1] * h, x2: p[2] * w, y2: p[3] * h },
      };
    }
    return at(p);
  });
  points.push({ close: true });
  slide.addShape('custGeom', Object.assign({ x, y, w, h, points }, opts));
}

// Right-pointing banner arrow (slide 23).
const BANNER_ARROW = [
  [0.811, 0.000], [1.000, 0.500], [0.811, 1.000], [0.811, 0.798],
  [0.236, 0.798], [0.000, 0.797], [0.227, 0.203], [0.236, 0.203],
  [0.236, 0.202], [0.811, 0.202],
];

// Swoosh arrow that links the circles on slide 22.
const SWOOSH_ARROW = [
  [0.0449, 0.8164], [0.0942, 1.0000], [0.1278, 0.8810],
  [0.4147, 1.0646, 0.7557, 0.8980, 0.9147, 0.4796],
  [0.9663, 0.3470, 0.9933, 0.2041, 0.9978, 0.0612],
  [0.9985, 0.0408, 0.9992, 0.0204, 1.0000, 0.0000],
  [0.9112, 0.0881], [0.8296, 0.0000], [0.8296, 0.0374],
  [0.8251, 0.1429, 0.8071, 0.2517, 0.7691, 0.3504],
  [0.6503, 0.6564, 0.4036, 0.7789, 0.1928, 0.6428],
  [0.2040, 0.6031, 0.2153, 0.5635, 0.2264, 0.5238],
  [0.1188, 0.5850], [0.0000, 0.6530],
  [0.0150, 0.7074, 0.0299, 0.7620, 0.0449, 0.8164],
];

// Wide panel with a bulging right edge (slide 7 overlays).
const BULGE_PANEL = [
  [0.186, 0.000], [0.809, 0.000], [0.811, 0.002],
  [0.929, 0.142, 1.000, 0.323, 1.000, 0.520],
  [1.000, 0.689, 0.948, 0.845, 0.858, 0.976],
  [0.840, 1.000], [0.186, 1.000],
  [0.083, 1.000, 0.000, 0.918, 0.000, 0.817],
  [0.000, 0.183], [0.000, 0.082, 0.083, 0.000, 0.186, 0.000],
];

/* ------------------------------------------------------------------ *
 * Slide chrome (the shapes the slide master paints on slides 2-24)
 * ------------------------------------------------------------------ */

function chrome(slide, num) {
  slide.addShape('ellipse', {
    x: 1.4957, y: 6.9792, w: 0.3029, h: 0.3029,
    fill: { color: '75F4F4' }, line: { color: MIST, width: 1 }, shadow: SOFT_SHADOW,
  });
  slide.addShape('roundRect', {
    x: 0.2530, y: 6.9792, w: 1.2001, h: 0.3029, rectRadius: 0.1515,
    fill: { color: '75F4F4' }, line: { color: MIST, width: 1 }, shadow: SOFT_SHADOW,
  });
  slide.addShape('ellipse', {
    x: 12.5547, y: 6.8192, w: 0.5599, h: 0.5599,
    fill: { color: '75F4F4' }, line: { color: MIST, width: 1 }, shadow: SOFT_SHADOW,
  });
  slide.addText(String(num), {
    x: 12.1675, y: 6.8972, w: 0.8359, h: 0.3702, align: 'right', valign: 'top',
    fontFace: HEAD_FONT, fontSize: 16, bold: true, color: TEAL,
  });
  slide.addText('Conference', {
    x: 0.3333, y: 6.9583, w: 1.1198, h: 0.3029, valign: 'top',
    fontFace: HEAD_FONT, fontSize: 12, color: TEAL,
  });
  tinyArrow(slide, 1.5533, 7.1347, 0.1876);
}

/* ------------------------------------------------------------------ *
 * Repeated content blocks
 * ------------------------------------------------------------------ */

const LOREM = {
  short: 'Lorem ipsum dolor sit amet, consectetuer',
  line: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget',
  s1: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. ',
  s2: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. ',
  s3: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque',
  s4: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, ',
  s5: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. ',
  s6: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Donec quam felis, ultricies nec, pellentesque eu, pretium quis, sem. Nulla consequat massa',
  s7: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes',
};

/** kicker + paragraph, the deck's default text block. */
function textBlock(slide, x, y, w, h, body, opts = {}) {
  kicker(slide, opts.kickerX === undefined ? x : opts.kickerX, y, opts.kickerW || 1.556,
    opts.kicker || 'Your Text Here', { align: opts.align });
  paragraph(slide, x, y + 0.423, w, h, body, { align: opts.align });
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

function slide01(pptx) {
  const s = pptx.addSlide();
  halo(s, 5.596, -0.319, 8.139, MIST);
  halo(s, 6.285, 0.346, 6.761, HAZE);
  heading(s, 0.653, 1.540, 5.647, 3.191, 'Conference Presentation Template', { fontSize: 68 });
  s.addText([
    { text: 'Your Text Goes Here ', options: { bold: true } },
    { text: '/ 15 Dec 2024' },
  ], {
    x: 1.619, y: 4.902, w: 4.680, h: 0.503, valign: 'top',
    fontFace: BODY_FONT, fontSize: 20, color: INK, lineSpacingMultiple: 1.3, paraSpaceAfter: 6,
  });
  bubble(s, 0.762, 4.861, 0.704, 'pink');
  tinyArrow(s, 0.994, 5.215, 0.241);
  bubble(s, 7.034, 0.698, 0.825, 'cyan');
  bubble(s, 7.859, 5.635, 0.648, 'pink');
  bubble(s, 11.952, 5.748, 0.423, 'cyan');
  imageFrame(s, 7.026, 1.110, 5.280, 5.280);
}

function slide02(pptx) {
  const s = pptx.addSlide();
  chrome(s, 2);
  s.addShape('ellipse', Object.assign({ x: -1.039, y: -0.678, w: 4.203, h: 4.203 },
    skin('cyan'), { line: { type: 'none' } }));
  heading(s, 6.667, 2.201, 5.500, 2.100, 'The Power of Connection', { fontSize: 66 });
  paragraph(s, 6.667, 4.301, 5.167, 0.940, LOREM.s1, { fontSize: 20 });
  card(s, 3.344, 5.635, 3.049, 0.776, 0.5, 'pink');
  paragraph(s, 3.672, 5.635, 2.493, 0.663, 'Your Text Here',
    { align: 'center', bold: true, fontSize: 28 });
  bubble(s, 8.518, 6.518, 1.463, 'cyan');
  bubble(s, 2.592, 0.157, 0.648, 'pink');
  imageFrame(s, 1.062, 1.201, 4.875, 4.875);
}

function slide03(pptx) {
  const s = pptx.addSlide();
  chrome(s, 3);
  heading(s, 0.667, 2.965, 4.000, 0.646, 'Redefining Events');
  paragraph(s, 0.667, 3.545, 4.389, 2.523,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Donec quam felis, ultricies nec, pellentesque eu, pretium quis, sem. Nulla consequat massa quis enim. Donec pede justo, fringilla vel, aliquet nec, vulputate eget, ');
  s.addText('Introduction: ', {
    x: 0.667, y: 2.561, w: 1.479, h: 0.404, valign: 'top',
    fontFace: HEAD_FONT, fontSize: 18, bold: true, color: INK, charSpacing: -1.5,
  });
  card(s, 0.667, 1.185, 1.226, 1.173, 0.31655, 'cyan');
  icon(s, 0.951, 1.443, 0.658, DEEP_TEAL);
  bubble(s, 10.820, -0.755, 1.463, 'cyan');
  bubble(s, 6.478, 5.616, 0.648, 'pink');
  imageFrame(s, 5.582, 0.999, 3.931, 3.931);
  imageFrame(s, 7.653, 1.944, 4.875, 4.875);
}

function slide04(pptx) {
  const s = pptx.addSlide();
  chrome(s, 4);
  heading(s, 0.667, 0.542, 5.528, 0.646, 'Why Online Conferences?');
  card(s, 1.347, 3.894, 10.069, 2.958, 0.22951, 'cyan');
  [['01.', 1.604], ['02.', 4.601]].forEach(([num, y]) => {
    textBlock(s, 2.354, y, 5.528, 1.122, LOREM.s5);
    heading(s, 1.611, y + 0.020, 1.000, 0.646, num);
  });
  bubble(s, 0.522, 2.925, 0.825, 'cyan');
  bubble(s, 12.199, 0.788, 0.648, 'pink');
  imageFrame(s, 8.340, 1.112, 2.833, 2.464);
  imageFrame(s, 8.340, 4.141, 2.833, 2.464);
}

function slide05(pptx) {
  const s = pptx.addSlide();
  chrome(s, 5);
  halo(s, 6.088, 0.200, 7.136, HAZE, true);
  heading(s, 0.667, 0.573, 4.819, 1.192, 'From Physical Events to Virtual Gatherings');
  paragraph(s, 0.667, 1.764, 4.389, 0.772,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean');
  [2.959, 4.497].forEach((y, i) => {
    textBlock(s, 1.298, y, 3.940, 0.772, LOREM.line);
    icon(s, 0.814, y + 0.140, 0.480, INK);
  });
  bubble(s, 6.133, 1.711, 0.825, 'cyan');
  bubble(s, 8.158, 6.057, 0.648, 'pink');
  bubble(s, 12.828, 4.539, 0.423, 'cyan');
  imageFrame(s, 6.708, 0.759, 5.895, 5.895);
}

function slide06(pptx) {
  const s = pptx.addSlide();
  chrome(s, 6);
  halo(s, -0.521, -0.389, 8.326, MIST);
  halo(s, 0.184, 0.292, 6.917, HAZE);
  s.addShape('ellipse', Object.assign({ x: 0.990, y: 1.097, w: 5.306, h: 5.306 }, skin('cyan')));
  heading(s, 1.247, 3.022, 4.792, 1.192, 'Demystifying Online Conferences', { align: 'center' });
  paragraph(s, 7.936, 1.179, 4.493, 1.472, LOREM.s5);
  stat(s, 7.906, 0.528, 1.556, '2025');
  card(s, 6.674, 0.813, 1.125, 1.077, 0.31655, 'pink');
  s.addShape('roundRect', {
    x: 3.298, y: 4.276, w: 0.661, h: 0.302, rectRadius: 0.151,
    fill: { type: 'none' }, line: { color: MIST, width: 1 },
  });
  tinyArrow(s, 3.508, 4.427, 0.241);
  icon(s, 6.884, 1.000, 0.693, ROSE);
  bubble(s, 0.254, 1.827, 0.825, 'cyan');
  bubble(s, 4.383, 6.718, 0.648, 'pink');
  imageFrame(s, 6.667, 3.094, 5.783, 3.568);
}

function slide07(pptx) {
  const s = pptx.addSlide();
  chrome(s, 7);
  halo(s, 0.849, 1.188, 5.874, MIST);
  halo(s, 1.346, 1.685, 4.880, HAZE);
  card(s, 3.181, 2.251, 9.357, 3.604, 0.18326, 'cyan');
  heading(s, 2.750, 0.542, 7.833, 0.646, 'Engaging Your Audience from Afar', { align: 'center' });
  freeform(s, 3.181, 2.251, 3.543, 3.604, BULGE_PANEL,
    { fill: { color: MIST, transparency: 50 }, line: { type: 'none' } });
  freeform(s, 3.181, 2.260, 2.949, 3.604, BULGE_PANEL,
    { fill: { color: HAZE, transparency: 50 }, line: { type: 'none' } });
  textBlock(s, 7.210, 2.843, 4.898, 1.822, LOREM.s6, { kickerX: 7.181 });
  bubble(s, 2.402, 1.629, 0.641, 'cyan');
  bubble(s, 2.540, 5.492, 0.386, 'pink');
  imageFrame(s, 1.876, 2.215, 3.820, 3.820);
}

function slide08(pptx) {
  const s = pptx.addSlide();
  chrome(s, 8);
  heading(s, 1.421, 0.542, 10.492, 0.646, 'Choosing the Right Online Conference Platform',
    { align: 'center' });
  const columns = [
    { x: 0.926, tone: 'paper', iconX: 1.412, iconSize: 0.472, iconY: 2.458, iconColor: INK },
    { x: 5.017, tone: 'cyan', iconX: 5.534, iconSize: 0.550, iconY: 2.396, iconColor: DEEP_TEAL },
    { x: 9.108, tone: 'pink', iconX: 9.658, iconSize: 0.472, iconY: 2.474, iconColor: ROSE },
  ];
  columns.forEach((col) => {
    card(s, col.x, 1.957, 3.300, 4.313, 0.18326, col.tone);
    kicker(s, col.x + 1.021, 2.296, 1.556, 'Your Text Here');
    paragraph(s, col.x + 1.021, 2.575, 2.144, 0.422, 'Lorem ipsum dolor sit');
    icon(s, col.iconX, col.iconY, col.iconSize, col.iconColor);
    imageFrame(s, col.x + 0.320, 3.336, 2.660, 2.660);
  });
  bubble(s, 12.510, 0.250, 1.646, 'cyan');
  bubble(s, 0.280, 0.672, 0.386, 'pink');
  bubble(s, 4.534, 6.556, 0.483, 'cyan');
}

function slide09(pptx) {
  const s = pptx.addSlide();
  chrome(s, 9);
  halo(s, 6.262, -2.278, 9.237, HAZE, true);
  heading(s, 1.215, 0.542, 5.048, 0.646, 'Pre-Conference Hype');
  bubble(s, 6.416, 0.250, 0.538, 'pink');
  bubble(s, 9.500, 6.566, 0.483, 'cyan');
  imageFrame(s, 1.246, 1.546, 5.667, 4.747);
  [[7.246, 1.546], [9.833, 1.546], [9.833, 4.088], [7.246, 4.088]]
    .forEach(([x, y]) => imageFrame(s, x, y, 2.254, 2.204));
}

function slide10(pptx) {
  const s = pptx.addSlide();
  chrome(s, 10);
  heading(s, 3.153, 0.542, 7.028, 0.646, 'Integrating Interactive Elements');
  const cells = [
    { x: 1.003, y: 1.912, tone: 'cyan', iconColor: DEEP_TEAL, iconSize: 0.719, iconAt: [1.206, 2.091] },
    { x: 7.295, y: 1.912, tone: 'pink', iconColor: ROSE, iconSize: 0.660, iconAt: [7.527, 2.123] },
    { x: 1.003, y: 4.268, tone: 'pink', iconColor: ROSE, iconSize: 0.563, iconAt: [1.284, 4.525] },
    { x: 7.295, y: 4.268, tone: 'cyan', iconColor: DEEP_TEAL, iconSize: 0.625, iconAt: [7.548, 4.494] },
  ];
  cells.forEach((c) => {
    iconTile(s, c.x, c.y, c.tone);
    icon(s, c.iconAt[0], c.iconAt[1], c.iconSize, c.iconColor);
    kicker(s, c.x + 1.257, c.y, 1.556, 'Your Text Here');
    paragraph(s, c.x + 1.287, c.y + 0.422, 3.804, 1.472, LOREM.s4);
  });
  bubble(s, -0.465, -0.484, 1.468, 'pink');
  bubble(s, 7.666, 6.958, 0.806, 'cyan');
}

function slide11(pptx) {
  const s = pptx.addSlide();
  chrome(s, 11);
  halo(s, -4.618, -2.758, 9.237, HAZE, true);
  card(s, 5.676, 1.726, 6.531, 2.024, 0.20539, 'cyan');
  heading(s, 0.635, 0.542, 8.626, 0.646, 'Balancing Content and Engagement');
  paragraph(s, 6.419, 2.600, 5.684, 0.772, LOREM.s3);
  stat(s, 6.390, 1.857, 2.431, '$400.000');
  [['01.', 5.775], ['02.', 7.956], ['03.', 10.180]].forEach(([num, x]) => {
    stat(s, x, 4.300, 0.864, num);
    kicker(s, x + 0.043, 4.968, 1.556, 'Your Text Here');
    paragraph(s, x + 0.043, 5.366, 2.137, 0.772, LOREM.short);
  });
  bubble(s, 0.583, 6.138, 0.433, 'pink');
  bubble(s, 10.742, -0.681, 1.361, 'cyan');
  imageFrame(s, 1.230, 1.726, 3.602, 4.690);
}

function slide12(pptx) {
  const s = pptx.addSlide();
  chrome(s, 12);
  halo(s, 4.841, -4.092, 12.214, MIST);
  halo(s, 5.797, -3.136, 10.303, HAZE, true);
  heading(s, 0.686, 0.561, 4.278, 1.192, 'Navigation and Technical Support');
  kicker(s, 0.808, 2.129, 3.775, LOREM.short);
  paragraph(s, 0.808, 2.522, 4.033, 1.472, LOREM.s4);
  [['01.', 0.866], ['02.', 3.524]].forEach(([num, x]) => {
    stat(s, x, 4.418, 0.864, num);
    kicker(s, x + 0.043, 5.086, 1.556, 'Your Text Here');
    paragraph(s, x + 0.043, 5.484, 2.137, 0.772, LOREM.short);
  });
  card(s, 7.310, 5.099, 4.857, 1.644, 0.20539, 'cyan');
  icon(s, 7.694, 5.509, 0.574, DEEP_TEAL);
  kicker(s, 8.349, 5.300, 1.556, 'Your Text Here');
  paragraph(s, 8.379, 5.636, 3.624, 0.772,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula');
  bubble(s, 5.707, 0.947, 0.419, 'pink');
  bubble(s, 7.464, 6.958, 0.309, 'cyan');
  imageFrame(s, 7.310, 0.770, 4.857, 4.153);
}

function slide13(pptx) {
  const s = pptx.addSlide();
  chrome(s, 13);
  halo(s, 6.986, -2.706, 9.237, HAZE, true);
  heading(s, 0.667, 0.542, 6.000, 0.646, 'Optimizing Delivery Formats');
  iconTile(s, 0.927, 1.912, 'pink');
  icon(s, 1.210, 2.177, 0.556, ROSE);
  kicker(s, 2.138, 1.831, 1.556, 'Your Text Here');
  paragraph(s, 2.168, 2.253, 3.804, 1.472, LOREM.s4);
  bulletList(s, 2.138, 4.029, 3.804, 1.991, [
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean',
    'Commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et',
    'Magnis dis parturient montes, ',
  ]);
  card(s, 7.362, 1.831, 4.689, 4.189, 0.13196, 'cyan');
  stat(s, 7.882, 2.111, 2.431, '500+');
  paragraph(s, 7.911, 2.855, 3.820, 1.122, LOREM.s3);
  stat(s, 7.882, 4.242, 2.431, '200+');
  paragraph(s, 7.911, 4.985, 3.820, 0.422, LOREM.short);
  bubble(s, 6.777, 1.415, 0.419, 'pink');
  bubble(s, 10.312, 6.186, 0.461, 'cyan');
}

function slide14(pptx) {
  const s = pptx.addSlide();
  chrome(s, 14);
  heading(s, 3.778, 0.542, 5.778, 0.646, 'Spark the Conversation', { align: 'center' });
  paragraph(s, 4.356, 1.082, 4.622, 0.422,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing', { align: 'center' });
  [
    { x: 1.105, tone: 'pink', tileX: 2.948, iconX: 3.232, iconY: 2.218, iconSize: 0.557, iconColor: ROSE, textX: 1.608, kickerX: 2.732 },
    { x: 7.419, tone: 'cyan', tileX: 9.261, iconX: 9.550, iconY: 2.245, iconSize: 0.546, iconColor: DEEP_TEAL, textX: 7.922, kickerX: 9.046 },
  ].forEach((c) => {
    card(s, c.x, 2.496, 4.810, 3.418, 0.14522, c.tone);
    iconTile(s, c.tileX, 1.980, c.tone);
    icon(s, c.iconX, c.iconY, c.iconSize, c.iconColor);
    kicker(s, c.kickerX, 3.480, 1.556, 'Your Text Here', { align: 'center' });
    paragraph(s, c.textX, 3.902, 3.804, 1.472, LOREM.s4, { align: 'center' });
  });
  bubble(s, 12.667, 0.250, 1.539, 'cyan');
  bubble(s, 5.202, 6.487, 0.419, 'pink');
}

function slide15(pptx) {
  const s = pptx.addSlide();
  chrome(s, 15);
  heading(s, 0.667, 1.092, 3.903, 1.192, 'Mastering\nthe Virtual Stage');
  stat(s, 0.681, 2.855, 2.431, '300+');
  paragraph(s, 0.681, 3.606, 4.905, 1.822, LOREM.s6);
  [
    { y: 0.773, tone: 'cyan', iconColor: DEEP_TEAL, iconAt: [7.353, 1.017], iconSize: 0.548 },
    { y: 2.398, tone: 'pink', iconColor: ROSE, iconAt: [7.349, 2.674], iconSize: 0.525 },
    { y: 4.023, tone: 'cyan', iconColor: DEEP_TEAL, iconAt: [7.375, 4.325], iconSize: 0.472 },
    { y: 5.648, tone: 'pink', iconColor: ROSE, iconAt: [7.327, 5.902], iconSize: 0.569 },
  ].forEach((row) => {
    iconTile(s, 7.049, row.y, row.tone);
    icon(s, row.iconAt[0], row.iconAt[1], row.iconSize, row.iconColor);
    kicker(s, 8.306, row.y - 0.114, 1.556, 'Your Text Here');
    paragraph(s, 8.336, row.y + 0.308, 3.804, 0.772, LOREM.line);
  });
  bubble(s, 4.099, 6.731, 1.539, 'cyan');
  bubble(s, 12.456, 0.274, 0.422, 'pink');
}

function slide16(pptx) {
  const s = pptx.addSlide();
  chrome(s, 16);
  heading(s, 0.667, 0.485, 5.810, 0.646, 'Building Connections Online');
  kicker(s, 0.667, 1.422, 1.608, 'Your Text Here');
  paragraph(s, 0.702, 1.802, 11.964, 1.472,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Donec quam felis, ultricies nec, pellentesque eu, pretium quis, sem. Nulla consequat massa quis enim. Donec pede justo, fringilla vel, aliquet nec, vulputate eget, arcu. In enim justo, rhoncus ut, imperdiet a, venenatis vitae, justo. Nullam dictum felis eu pede mollis pretium. Integer tincidunt. Cras dapibus.');
  [
    { tile: 0.667, tone: 'cyan', iconX: 0.927, iconColor: DEEP_TEAL, textX: 1.936, kickerX: 1.901 },
    { tile: 6.667, tone: 'pink', iconX: 6.928, iconColor: ROSE, textX: 7.936, kickerX: 7.901 },
  ].forEach((c) => {
    iconTile(s, c.tile, 4.190, c.tone);
    icon(s, c.iconX, 4.422, 0.603, c.iconColor);
    kicker(s, c.kickerX, 4.062, 1.608, 'Your Text Here');
    paragraph(s, c.textX, 4.443, 3.735, 1.472, LOREM.s4);
  });
  bubble(s, 10.759, -0.939, 1.825, 'cyan');
  bubble(s, 6.054, 7.289, 0.422, 'pink');
}

function slide17(pptx) {
  const s = pptx.addSlide();
  chrome(s, 17);
  heading(s, 0.667, 0.485, 4.587, 0.646, 'Building Community');
  s.addChart(pptx.ChartType.bar, [
    { name: 'Series 1', labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'], values: [4.3, 2.5, 3.5, 2.1] },
    { name: 'Series 2', labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'], values: [2.4, 4.4, 3, 2.8] },
    { name: 'Series 3', labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'], values: [2, 2.5, 3, 5] },
  ], {
    x: 1.131, y: 1.630, w: 4.944, h: 4.977,
    barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 150,
    chartColors: [blend(...GRADIENT.cyan), blend(...GRADIENT.pink), blend(...GRADIENT.violet)],
    showLegend: false, showValue: false,
    valAxisMaxVal: 12, valAxisMajorUnit: 2, valAxisLineShow: false,
    valGridLine: { color: 'F2F2F2', size: 0.75 }, catGridLine: { style: 'none' },
    catAxisLineColor: HAIRLINE,
    catAxisLabelColor: INK, valAxisLabelColor: INK,
    catAxisLabelFontFace: BODY_FONT, valAxisLabelFontFace: BODY_FONT,
    catAxisLabelFontSize: 12, valAxisLabelFontSize: 12,
  });
  // Three legend chips to the right of the chart.
  [
    { x: 6.884, y: 1.778, tone: 'cyan', textX: 7.374 },
    { x: 6.884, y: 2.690, tone: 'pink', textX: 7.374 },
    { x: 9.857, y: 1.778, tone: 'violet', textX: 10.346 },
  ].forEach((chip) => {
    card(s, chip.x, chip.y, 0.422, 0.422, 0.30234, chip.tone);
    tinyArrow(s, chip.x + 0.118, chip.y + 0.208, 0.183);
    slideChipText(s, chip.textX, chip.y - 0.148);
  });
  kicker(s, 6.849, 3.750, 1.608, 'Your Text Here');
  bulletList(s, 6.884, 4.131, 5.456, 2.341, [
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa.',
    'Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus.',
    'Donec quam felis, ultricies nec, pellentesque eu, pretium quis, sem. Nulla consequat massa quis enim. Donec pede justo, ',
  ]);
  bubble(s, 6.884, -0.716, 1.432, 'cyan');
  bubble(s, -0.381, 5.317, 0.762, 'pink');
}

function slideChipText(s, x, y) {
  s.addText(LOREM.short, {
    x, y, w: 2.168, h: 0.640, valign: 'top',
    fontFace: BODY_FONT, fontSize: BODY_SIZE, bold: true, color: INK, paraSpaceAfter: 6,
  });
}

function slide18(pptx) {
  const s = pptx.addSlide();
  chrome(s, 18);
  heading(s, 0.667, 0.485, 7.117, 0.646, 'Utilizing On-Demand Recordings');
  s.addChart(pptx.ChartType.doughnut, [{
    name: 'Sales',
    labels: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'],
    values: [3, 2.5, 4, 0.5],
  }], {
    x: 4.360, y: 1.299, w: 5.104, h: 4.993,
    holeSize: 47, showLegend: false, showValue: true, firstSliceAng: 0,
    chartColors: ['3FEFEF', 'A3E5F5', 'C4C0ED', 'E2B2CA'],
    dataLabelColor: DEEP_TEAL, dataLabelFontFace: HEAD_FONT,
    dataLabelFontSize: 18, dataLabelFontBold: true,
  });
  // Callout labels, two per side, connected by dotted leader lines.
  [
    { x: 9.681, y: 1.566, align: 'left', textX: 9.710 },
    { x: 9.681, y: 4.482, align: 'left', textX: 9.710 },
    { x: 2.347, y: 1.566, align: 'right', textX: 1.238 },
    { x: 2.347, y: 4.482, align: 'right', textX: 1.238 },
  ].forEach((c) => {
    kicker(s, c.x, c.y, 1.556, 'Your Text Here', { align: c.align });
    paragraph(s, c.textX, c.y + 0.422, 2.665, 1.122, LOREM.line, { align: c.align });
  });
  [
    { x: 3.975, y: 1.834, w: 2.421, color: 'E9C5D7' },
    { x: 3.975, y: 4.757, w: 1.150, color: 'D5D2F2' },
    { x: 7.484, y: 1.834, w: 2.072, color: '75F4F4' },
    { x: 8.711, y: 4.757, w: 0.845, color: '90E0F3' },
  ].forEach((l) => {
    s.addShape('line', { x: l.x, y: l.y, w: l.w, h: 0, line: { color: l.color, width: 0.75, dashType: 'dash' } });
  });
  bubble(s, -0.445, 4.037, 0.891, 'pink');
  bubble(s, 11.903, 0.620, 0.472, 'cyan');
}

function slide19(pptx) {
  const s = pptx.addSlide();
  chrome(s, 19);
  bubble(s, 0.375, 4.825, 0.959, 'pink');
  card(s, 0.851, 2.122, 11.631, 4.414, 0.09578, 'paper');
  heading(s, 0.667, 0.485, 4.461, 0.646, 'Data-Driven Insights');
  const days = ['Day 1', 'Day 2', 'Day 3', 'Day 4', 'Day 5', 'Day 6', 'Day 7'];
  s.addChart(pptx.ChartType.line, [
    { name: 'Series 1', labels: days, values: [3.3, 5.3, 5, 4.5, 4, 4.7, 5] },
    { name: 'Series 2', labels: days, values: [1.2, 3, 2.5, 3.6, 3.6, 3, 4] },
    { name: 'Series 3', labels: days, values: [2.5, 2.8, 3, 4.6, 4.5, 3.5, 3] },
  ], {
    x: 1.293, y: 2.471, w: 10.747, h: 3.777,
    chartColors: ['75F4F4', '90E0F3', 'D5D2F2'],
    lineSize: 2.25, lineDataSymbol: 'circle', lineDataSymbolSize: 5,
    showLegend: false, showValue: false,
    valAxisMaxVal: 6, valAxisMajorUnit: 1, valAxisLineShow: false,
    valGridLine: { color: HAIRLINE, size: 0.75 }, catGridLine: { style: 'none' },
    catAxisLineColor: HAIRLINE,
    catAxisLabelColor: INK, valAxisLabelColor: INK,
    catAxisLabelFontFace: BODY_FONT, valAxisLabelFontFace: BODY_FONT,
    catAxisLabelFontSize: 12, valAxisLabelFontSize: 12,
  });
  card(s, 7.010, 0.777, 4.825, 1.625, 0.20202, 'cyan');
  paragraph(s, 0.702, 1.073, 5.143, 0.772, LOREM.s2);
  s.addShape('wedgeRoundRectCallout', Object.assign(
    { x: 3.514, y: 2.292, w: 1.053, h: 0.549, shadow: SOFT_SHADOW }, skin('cyan')));
  s.addText('$560', {
    x: 3.514, y: 2.276, w: 1.053, h: 0.503, align: 'center', valign: 'top',
    fontFace: HEAD_FONT, fontSize: 20, bold: true, color: DEEP_TEAL,
    lineSpacingMultiple: 1.3, paraSpaceAfter: 6,
  });
  kicker(s, 7.274, 0.941, 1.608, 'Your Text Here');
  paragraph(s, 7.310, 1.322, 4.308, 0.772, LOREM.s1);
  bubble(s, 12.011, 0.412, 0.472, 'cyan');
}

function slide20(pptx) {
  const s = pptx.addSlide();
  chrome(s, 20);
  halo(s, 6.667, 1.332, 5.626, MIST);
  heading(s, 0.667, 0.485, 7.188, 0.646, 'The Future of Online Conferences');
  halo(s, 7.438, 2.147, 4.084, HAZE, true);
  // Three percentage bubbles.
  [
    { x: 6.636, y: 3.850, d: 2.351, tone: 'cyan', color: DEEP_TEAL,
      label: '60%', size: 60, box: [6.874, 4.365, 1.953, 1.111], cap: [6.811, 5.337, 1.854, 14] },
    { x: 10.648, y: 3.465, d: 1.585, tone: 'violet', color: INDIGO,
      label: '10%', size: 40, box: [10.795, 3.759, 1.366, 0.774], cap: [10.740, 4.431, 1.345, 12] },
    { x: 8.587, y: 1.345, d: 1.873, tone: 'pink', color: ROSE,
      label: '30%', size: 48, box: [8.771, 1.649, 1.689, 0.909], cap: [8.764, 2.457, 1.541, 14] },
  ].forEach((b) => {
    bubble(s, b.x, b.y, b.d, b.tone);
    s.addText(b.label, {
      x: b.box[0], y: b.box[1], w: b.box[2], h: b.box[3], align: 'center', valign: 'middle',
      fontFace: HEAD_FONT, fontSize: b.size, color: b.color,
    });
    s.addText('Your Text Here', {
      x: b.cap[0], y: b.cap[1], w: b.cap[2], h: 0.313, align: 'center', valign: 'top',
      fontFace: BODY_FONT, fontSize: b.cap[3], color: b.color,
    });
  });
  bubble(s, 8.453, 6.693, 0.418, 'pink');
  bubble(s, 7.638, 1.489, 0.614, 'cyan');
  bubble(s, 10.169, 5.876, 0.292, 'violet');
  kicker(s, 0.771, 1.513, 1.608, 'Your Text Here');
  paragraph(s, 0.786, 1.894, 5.019, 2.873,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Donec quam felis, ultricies nec, pellentesque eu, pretium quis, sem. Nulla consequat massa quis enim. Donec pede justo, fringilla vel, aliquet nec, vulputate eget, arcu. In enim justo, rhoncus ut, imperdiet a, venenatis vitae, justo.');
  [4.911, 5.687].forEach((y) => {
    s.addText(LOREM.line, {
      x: 1.378, y, w: 3.824, h: 0.640, valign: 'top',
      fontFace: BODY_FONT, fontSize: BODY_SIZE, color: INK, paraSpaceAfter: 6,
    });
    card(s, 0.889, y + 0.148, 0.422, 0.422, 0.30234, 'paper');
    tinyArrow(s, 1.006, y + 0.356, 0.183, HAIRLINE);
  });
}

function slide21(pptx) {
  const s = pptx.addSlide();
  chrome(s, 21);
  halo(s, 7.639, 1.139, 5.222, MIST);
  halo(s, 8.261, 1.803, 3.979, HAZE, true);
  heading(s, 0.667, 0.485, 7.188, 0.646, 'Overcoming Challenges');
  [
    { x: 0.794, tone: 'cyan', label: 'Step 1', color: DEEP_TEAL },
    { x: 3.018, tone: 'pink', label: 'Step 2', color: ROSE },
    { x: 5.241, tone: 'violet', label: 'Step 3', color: INDIGO },
  ].forEach((step) => {
    s.addShape('chevron', Object.assign({
      x: step.x, y: 3.272, w: 2.021, h: 0.956, rectRadius: 0.956 * 0.35074, shadow: SOFT_SHADOW,
    }, skin(step.tone)));
    s.addText(step.label, {
      x: step.x + 0.595, y: 3.482, w: 1.056, h: 0.460, align: 'center', valign: 'top',
      fontFace: HEAD_FONT, fontSize: 20, bold: true, color: step.color,
      lineSpacingMultiple: 1.2, paraSpaceBefore: 6, paraSpaceAfter: 8,
    });
  });
  paragraph(s, 2.696, 1.408, 2.665, 1.122, LOREM.line, { align: 'center' });
  paragraph(s, 0.472, 4.970, 2.665, 1.122, LOREM.line, { align: 'center' });
  paragraph(s, 4.895, 4.970, 2.665, 1.122, LOREM.line, { align: 'center' });
  [
    { x: 4.028, y: 2.679, color: 'E9C5D7' },
    { x: 1.805, y: 4.259, color: '90E0F3' },
    { x: 6.227, y: 4.259, color: 'D5D2F2' },
  ].forEach((l) => {
    s.addShape('line', { x: l.x, y: l.y, w: 0, h: 0.593, line: { color: l.color, width: 0.75, dashType: 'dash' } });
  });
  s.addText('+560', {
    x: 9.273, y: 2.595, w: 1.953, h: 1.111, align: 'center', valign: 'middle',
    fontFace: HEAD_FONT, fontSize: 60, color: DEEP_TEAL,
  });
  paragraph(s, 8.883, 3.623, 2.734, 1.122,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula', { align: 'center' });
  bubble(s, 10.961, 5.310, 0.418, 'pink');
  bubble(s, 9.825, 0.792, 0.614, 'cyan');
}

function slide22(pptx) {
  const s = pptx.addSlide();
  chrome(s, 22);
  heading(s, 3.073, 0.485, 7.188, 0.646, 'The Impact of Online Conferences', { align: 'center' });
  // Alternating swoosh arrows + numbered discs; odd steps caption above, even below.
  const steps = [
    { arrowX: 1.450, arrowY: 3.706, up: true, discX: 2.041, discY: 3.348, tone: 'cyan', num: '01', color: DEEP_TEAL, numY: 3.591, textX: 1.373, textY: 1.722, kickerX: 1.907, kickerY: 1.341 },
    { arrowX: 4.086, arrowY: 2.823, up: false, discX: 4.677, discY: 3.697, tone: 'pink', num: '02', color: ROSE, numY: 3.937, textX: 3.937, textY: 5.256, kickerX: 4.471, kickerY: 4.876 },
    { arrowX: 6.756, arrowY: 3.706, up: true, discX: 7.347, discY: 3.348, tone: 'cyan', num: '03', color: DEEP_TEAL, numY: 3.591, textX: 6.632, textY: 1.722, kickerX: 7.166, kickerY: 1.341 },
    { arrowX: 9.392, arrowY: 2.823, up: false, discX: 9.983, discY: 3.697, tone: 'pink', num: '04', color: ROSE, numY: 3.937, textX: 9.280, textY: 5.256, kickerX: 9.814, kickerY: 4.876 },
  ];
  steps.forEach((st) => {
    freeform(s, st.arrowX, st.arrowY, 2.492, 1.643, SWOOSH_ARROW, Object.assign({
      rotate: st.up ? 333 : 27, flipH: true, flipV: !st.up, shadow: SOFT_SHADOW,
    }, skin(st.tone)));
    bubble(s, st.discX, st.discY, 1.127, st.tone);
    s.addText(st.num, {
      x: st.discX + 0.035, y: st.numY, w: 1.056, h: 0.604, align: 'center', valign: 'top',
      fontFace: HEAD_FONT, fontSize: 28, bold: true, color: st.color,
      lineSpacingMultiple: 1.2, paraSpaceBefore: 6, paraSpaceAfter: 8,
    });
    kicker(s, st.kickerX, st.kickerY, 1.501, 'Your Text Here', { align: 'center' });
    paragraph(s, st.textX, st.textY, 2.569, 1.472, LOREM.s2, { align: 'center' });
  });
  bubble(s, 0.312, 5.644, 0.418, 'pink');
  bubble(s, 12.827, 0.835, 1.014, 'cyan');
}

function slide23(pptx) {
  const s = pptx.addSlide();
  chrome(s, 23);
  heading(s, 2.667, 0.485, 8.000, 0.646, 'Embracing the Power of Online Events', { align: 'center' });
  // Five stacked banner arrows, alternating direction.
  const arrows = [
    { x: 4.994, y: 1.564, tone: 'cyan', flip: true },
    { x: 5.588, y: 2.555, tone: 'pink', flip: false },
    { x: 4.994, y: 3.546, tone: 'cyan', flip: true },
    { x: 5.588, y: 4.536, tone: 'pink', flip: false },
    { x: 4.994, y: 5.527, tone: 'cyan', flip: true },
  ];
  arrows.forEach((a) => {
    freeform(s, a.x, a.y, 2.752, 1.304, BANNER_ARROW,
      Object.assign({ flipH: a.flip, shadow: SOFT_SHADOW }, skin(a.tone)));
  });
  const icons = [
    [6.021, 1.970, 0.456, DEEP_TEAL], [6.890, 2.998, 0.418, ROSE],
    [6.051, 4.000, 0.396, DEEP_TEAL], [6.921, 5.010, 0.357, ROSE],
    [6.067, 5.997, 0.365, DEEP_TEAL],
  ];
  icons.forEach(([x, y, sz, color]) => icon(s, x, y, sz, color));
  // Captions: left column reads right-aligned, right column left-aligned.
  [1.477, 3.648, 5.584].forEach((y) => {
    kicker(s, 3.155, y, 1.501, 'Your Text Here', { align: 'right' });
    paragraph(s, 0.905, y + 0.380, 3.752, 0.772, LOREM.line, { align: 'right' });
  });
  [2.630, 4.576].forEach((y) => {
    kicker(s, 8.677, y, 1.501, 'Your Text Here');
    paragraph(s, 8.677, y + 0.380, 3.752, 0.772, LOREM.line);
  });
  bubble(s, 8.471, 7.124, 0.751, 'pink');
  bubble(s, 11.262, -1.288, 2.333, 'cyan');
}

function slide24(pptx) {
  const s = pptx.addSlide();
  chrome(s, 24);
  heading(s, 0.667, 0.485, 8.000, 0.646, 'Your Questions Answered');
  // Three interlocking rings (0.9" band) with a triangle in the middle.
  [
    { x: 5.427, y: 1.750, d: 2.467, tone: 'pink' },
    { x: 6.255, y: 3.200, d: 2.467, tone: 'violet' },
    { x: 4.591, y: 3.200, d: 2.467, tone: 'cyan' },
  ].forEach((ring) => {
    s.addShape('donut', Object.assign({
      x: ring.x, y: ring.y, w: ring.d, h: ring.d, rectRadius: 0.90, shadow: SOFT_SHADOW,
    }, skin(ring.tone)));
  });
  s.addShape('triangle', Object.assign({
    x: 5.915, y: 3.079, w: 1.503, h: 1.324, shadow: SOFT_SHADOW,
  }, skin('paper')));
  icon(s, 6.421, 3.688, 0.490, INK);
  [
    { text: '01', x: 6.236, y: 1.868, color: ROSE },
    { text: '02', x: 7.699, y: 4.442, color: INDIGO },
    { text: '03', x: 4.698, y: 4.442, color: DEEP_TEAL },
  ].forEach((n) => {
    s.addText(n.text, {
      x: n.x, y: n.y, w: 0.864, h: 0.604, align: 'center', valign: 'top',
      fontFace: HEAD_FONT, fontSize: 28, bold: true, color: n.color,
      lineSpacingMultiple: 1.2, paraSpaceBefore: 6, paraSpaceAfter: 8,
    });
  });
  paragraph(s, 7.979, 1.142, 4.180, 1.472, LOREM.s7);
  paragraph(s, 8.845, 4.442, 3.318, 1.822, LOREM.s7);
  paragraph(s, 1.102, 4.442, 3.318, 1.822, LOREM.s7, { align: 'right' });
  bubble(s, 7.101, 6.181, 0.305, 'pink');
  bubble(s, 3.970, 2.701, 0.478, 'cyan');
  bubble(s, 7.531, -0.417, 0.810, 'cyan');
  bubble(s, 8.619, 3.640, 0.221, 'pink');
}

function slide25(pptx) {
  const s = pptx.addSlide();
  halo(s, 1.219, -1.697, 10.894, MIST);
  halo(s, 2.142, -0.807, 9.050, HAZE);
  s.addShape('ellipse', Object.assign({ x: 3.196, y: 0.247, w: 6.942, h: 6.942 }, skin('cyan')));
  heading(s, 4.308, 2.944, 4.718, 1.131, 'Thank You', { fontSize: 68, align: 'center' });
  paragraph(s, 5.048, 3.723, 3.238, 0.744, 'For the attention',
    { align: 'center', bold: true, fontSize: 32 });
  bubble(s, 10.246, 5.905, 0.418, 'pink');
  bubble(s, 10.717, 0.365, 0.949, 'cyan');
  bubble(s, 1.417, 1.546, 0.292, 'violet');
  bubble(s, 1.815, 4.071, 0.648, 'pink');
  bubble(s, 4.308, 6.114, 0.498, 'cyan');
  bubble(s, 7.244, 0.109, 0.451, 'violet');
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'DECK';
  pptx.theme = { headFontFace: HEAD_FONT, bodyFontFace: BODY_FONT };
  pptx.title = 'Conference Presentation Template';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
    slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18,
    slide19, slide20, slide21, slide22, slide23, slide24, slide25].forEach((fn) => fn(pptx));

  const out = path.join(__dirname, '09e7c35d-54f5-4d11-8b10-f1335b1b77e4_grok_final.pptx');
  return pptx.writeFile({ fileName: out }).then(() => console.log('wrote', out));
}

build().catch((err) => { console.error(err); process.exit(1); });
