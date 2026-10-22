/**
 * "MediCos" medical company deck — 20 slides, 13.333in x 7.5in.
 * Rebuilt from scratch with pptxgenjs. Raster photos in the original are
 * replaced by grey rounded "[image]" placeholders of the same box.
 *
 *   node 16a31921-05e0-4f5f-9921-83642c2721fd_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const C = {
  purple: '5F38D3', // accent2 - deck background
  mint: '32FDCD', // accent1 - highlight
  deep: '452999', // accent3 - dark text on light panels
  white: 'FFFFFF',
  img: 'D4D4D4', // image-placeholder fill
  imgText: 'A0A0A0',
};

const HEAD = 'Poppins'; // titles / wordmark
const BODY = 'Open Sans'; // everything else

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

/* Soft drop shadows used by the original deck. */
const SHADOW_TIGHT = { type: 'outer', color: '000000', opacity: 0.4, blur: 4, offset: 3, angle: 45 };
const SHADOW_SOFT = { type: 'outer', color: '000000', opacity: 0.2, blur: 50, offset: 20, angle: 45 };
const SHADOW_TIGHT_TR = { type: 'outer', color: '000000', opacity: 0.4, blur: 4, offset: 3, angle: 135 };

/* Lorem blocks reused across the deck. */
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor';
const LOREM_MED =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
  'Fusce posuere congue massa. Fusce posuere, ';
const LOREM_LONG = LOREM_MED + 'magna sed pulvinar osuere, ';
const LOREM_XL = LOREM_LONG + 'posuere Fusce posuere congue massa. Fusce posuere, congue massa. ';

/* ---------------------------------------------------------------- helpers */

/** Body copy: 10.5pt Open Sans, justified, 1.5 line spacing. */
function body(slide, text, x, y, w, h, o) {
  o = o || {};
  slide.addText(text, {
    x: x, y: y, w: w, h: h,
    fontFace: BODY, fontSize: o.fontSize || 10.5, color: o.color || C.white,
    bold: o.bold || false,
    align: o.align || 'justify',
    lineSpacingMultiple: o.lineSpacingMultiple === undefined ? 1.5 : o.lineSpacingMultiple,
    shape: o.shape, fit: 'resize',
  });
}

/** Small bold label, e.g. "Company Vision". */
function label(slide, text, x, y, w, h, o) {
  o = o || {};
  slide.addText(text, {
    x: x, y: y, w: w, h: h,
    fontFace: BODY, fontSize: o.fontSize || 12, bold: true, color: o.color || C.deep,
    align: o.align || 'left',
    lineSpacingMultiple: o.lineSpacingMultiple,
    shape: o.shape, fit: 'resize',
  });
}

/** Two-tone 40pt Poppins heading. `lines` = [[whitePart, mintPart], ...]. */
function heading(slide, lines, x, y, w, h, o) {
  o = o || {};
  const runs = [];
  lines.forEach(function (parts, i) {
    parts.forEach(function (part, j) {
      runs.push({
        text: part[0],
        options: { color: part[1], breakLine: j === parts.length - 1 && i < lines.length - 1 },
      });
    });
  });
  slide.addText(runs, {
    x: x, y: y, w: w, h: h,
    fontFace: HEAD, fontSize: o.fontSize || 40, bold: true,
    align: o.align || 'left', fit: 'resize',
  });
}

/** Rounded card. */
function card(slide, x, y, w, h, fill, o) {
  o = o || {};
  const opts = { x: x, y: y, w: w, h: h, fill: { color: fill } };
  if (o.transparency) opts.fill.transparency = o.transparency;
  if (o.radius) opts.rectRadius = o.radius;
  if (o.shadow) opts.shadow = o.shadow;
  if (o.flipH) opts.flipH = true;
  if (o.flipV) opts.flipV = true;
  slide.addShape(o.shape || 'roundRect', opts);
}

/** Circle / rounded badge carrying a centred number. */
function badge(slide, shape, x, y, size, fill, text, o) {
  o = o || {};
  const opts = { x: x, y: y, w: size, h: size, fill: { color: fill }, valign: 'middle', align: 'center' };
  if (o.transparency) opts.fill.transparency = o.transparency;
  if (o.shadow) opts.shadow = o.shadow;
  slide.addShape(shape, opts);
  slide.addText(text, {
    x: x, y: y, w: size, h: size, shape: shape,
    fontFace: BODY, fontSize: o.fontSize || 12, bold: true, color: o.color || C.deep,
    align: 'center', valign: 'middle', fill: { type: 'none' }, line: { type: 'none' },
  });
}

/** Grey stand-in for a photograph. */
function imageBox(slide, x, y, w, h, o) {
  o = o || {};
  const radius = o.radius === undefined ? Math.min(w, h) / 6 : o.radius;
  slide.addShape(o.shape || 'roundRect', {
    x: x, y: y, w: w, h: h, rectRadius: radius, fill: { color: C.img },
  });
  // Keep the caption inside the visible part of boxes that bleed off-slide.
  const cx = Math.max(0, Math.min(x, SLIDE_W - w));
  const cy = Math.max(0, Math.min(y, SLIDE_H - h));
  slide.addText('[image]', {
    x: cx, y: cy + Math.min(h, SLIDE_H - cy) / 2 - 0.17, w: Math.min(w, SLIDE_W - cx), h: 0.34,
    fontFace: BODY, fontSize: 11, color: C.imgText, align: 'center', valign: 'middle',
  });
}

/**
 * Quarter-disc that sits in a slide corner (the deck's recurring "blob").
 * Local path: square outline closed by a 90-degree arc centred on the
 * top-right corner; flips move it to the other corners.
 */
function cornerDisc(slide, corner, color) {
  const w = 1.0776;
  const h = 1.0793;
  const pts = [
    { x: 0, y: 0, moveTo: true },
    { x: w, y: 0 },
    { x: w, y: h },
    { x: 0, y: 0, curve: { type: 'cubic', x1: 0.4346 * w, y1: h, x2: 0, y2: 0.5595 * h } },
    { close: true },
  ];
  const box = corner === 'bl'
    ? { x: 0, y: SLIDE_H - h, flipH: true, flipV: true }
    : { x: SLIDE_W - w - 0.0075, y: -0.0176 };
  slide.addShape('custGeom', {
    x: box.x, y: box.y, w: w, h: h, points: pts, fill: { color: color },
    flipH: !!box.flipH, flipV: !!box.flipV,
  });
}

/** "MediCos" wordmark in the top-left corner. */
function wordmark(slide, mediColor) {
  slide.addText(
    [{ text: 'Medi', options: { color: mediColor || C.white } }, { text: 'Cos', options: { color: C.mint } }],
    { x: 0.304, y: 0.365, w: 9.154, h: 0.303, fontFace: HEAD, fontSize: 12, bold: true, valign: 'top' }
  );
}

/** Thin "next" arrow bottom-right (present on every slide but the last). */
function nextArrow(slide, color) {
  slide.addShape('line', {
    x: 11.771, y: 6.803, w: 0.3943, h: 0,
    line: { color: color || C.mint, width: 2.5, endArrowType: 'triangle' },
  });
}

/**
 * Background + arrow + corner blob, shared by every slide.
 * `markOnTop` defers the wordmark to the builder so it can sit above a
 * full-bleed image (slides 4 and 19).
 */
function chrome(pptx, o) {
  o = o || {};
  const slide = pptx.addSlide();
  slide.background = { color: o.bg || C.purple };
  if (!o.markOnTop) cornerDisc(slide, o.corner || 'tr', o.cornerColor || C.white);
  if (!o.noArrow) nextArrow(slide, o.arrowColor);
  if (!o.markOnTop) wordmark(slide, o.wordmark);
  return slide;
}

/** Section title used by slides 6 and 9-18: centred 40pt two-tone. */
function centeredTitle(slide, white, mint, y) {
  heading(slide, [[[white, C.white], [mint, C.mint]]], 2.09, y === undefined ? 1.397 : y, 9.154, 0.774,
    { align: 'center' });
}

/* ------------------------------------------------------------------ icons */
/* The original uses small vector glyphs; these rebuild them from primitives. */

function iconBriefcase(slide, x, y, w, h, color) {
  slide.addShape('roundRect', {
    x: x + w * 0.3, y: y, w: w * 0.4, h: h * 0.3, rectRadius: h * 0.06,
    fill: { type: 'none' }, line: { color: color, width: 2 },
  });
  slide.addShape('roundRect', { x: x, y: y + h * 0.24, w: w, h: h * 0.76, rectRadius: h * 0.08, fill: { color: color } });
  slide.addShape('rect', { x: x, y: y + h * 0.52, w: w, h: h * 0.07, fill: { color: C.white } });
  slide.addShape('rect', { x: x + w * 0.42, y: y + h * 0.46, w: w * 0.16, h: h * 0.19, fill: { color: C.white } });
}

function iconPeople(slide, x, y, w, h, color) {
  [[0.0, 0.62], [0.62, 0.62], [0.28, 0.78]].forEach(function (p, i) {
    const s = p[1] * w * 0.42;
    const cx = x + p[0] * w;
    const cy = y + (i === 2 ? 0.0 : h * 0.14);
    slide.addShape('ellipse', { x: cx, y: cy, w: s, h: s, fill: { color: color } });
    slide.addShape('roundRect', {
      x: cx - s * 0.32, y: cy + s * 1.1, w: s * 1.64, h: h * (i === 2 ? 0.62 : 0.5),
      rectRadius: s * 0.5, fill: { color: color },
    });
  });
}

function iconBell(slide, x, y, w, h, color) {
  slide.addShape('roundRect', { x: x + w * 0.42, y: y, w: w * 0.16, h: h * 0.14, rectRadius: 0.03, fill: { color: color } });
  slide.addShape('blockArc', {
    x: x, y: y + h * 0.08, w: w, h: h * 1.3, fill: { color: color },
    angleRange: [180, 360], arcThicknessRatio: 1,
  });
  slide.addShape('roundRect', { x: x, y: y + h * 0.62, w: w, h: h * 0.14, rectRadius: 0.03, fill: { color: color } });
  slide.addShape('ellipse', { x: x + w * 0.36, y: y + h * 0.78, w: w * 0.28, h: h * 0.22, fill: { color: color } });
}

function iconChat(slide, x, y, w, h, color) {
  slide.addShape('ellipse', { x: x, y: y, w: w, h: h * 0.8, fill: { color: color } });
  slide.addShape('triangle', {
    x: x + w * 0.12, y: y + h * 0.6, w: w * 0.3, h: h * 0.4, rotate: 200, fill: { color: color },
  });
}

function iconBank(slide, x, y, w, h, color) {
  slide.addShape('triangle', { x: x, y: y, w: w, h: h * 0.3, fill: { color: color } });
  slide.addShape('rect', { x: x + w * 0.02, y: y + h * 0.32, w: w * 0.96, h: h * 0.1, fill: { color: color } });
  [0.12, 0.38, 0.64].forEach(function (fx) {
    slide.addShape('rect', { x: x + fx * w, y: y + h * 0.45, w: w * 0.14, h: h * 0.35, fill: { color: color } });
  });
  slide.addShape('rect', { x: x, y: y + h * 0.84, w: w, h: h * 0.14, fill: { color: color } });
}

/* Handset: two rounded "ear" blocks on a diagonal, joined by a slim bar. */
function iconPhone(slide, x, y, w, h, color) {
  slide.addShape('roundRect', {
    x: x + w * 0.04, y: y + h * 0.04, w: w * 0.42, h: h * 0.34, rectRadius: h * 0.1,
    rotate: 315, fill: { color: color },
  });
  slide.addShape('roundRect', {
    x: x + w * 0.54, y: y + h * 0.58, w: w * 0.42, h: h * 0.34, rectRadius: h * 0.1,
    rotate: 315, fill: { color: color },
  });
  slide.addShape('roundRect', {
    x: x + w * 0.25, y: y + h * 0.3, w: w * 0.5, h: h * 0.4, rectRadius: h * 0.06,
    rotate: 315, fill: { color: color },
  });
}

function iconDoc(slide, x, y, w, h, color) {
  slide.addShape('roundRect', { x: x, y: y, w: w, h: h, rectRadius: 0.03, fill: { color: color } });
  [0.28, 0.48, 0.68].forEach(function (fy, i) {
    slide.addShape('rect', {
      x: x + w * 0.2, y: y + h * fy, w: w * (i === 2 ? 0.4 : 0.6), h: h * 0.08, fill: { color: C.white },
    });
  });
}

function iconGlobe(slide, x, y, w, h, color) {
  slide.addShape('ellipse', { x: x, y: y, w: w, h: h, fill: { color: color } });
  slide.addShape('ellipse', {
    x: x + w * 0.32, y: y, w: w * 0.36, h: h, fill: { type: 'none' }, line: { color: C.white, width: 1 },
  });
  slide.addShape('line', { x: x + w * 0.06, y: y + h * 0.5, w: w * 0.88, h: 0, line: { color: C.white, width: 1 } });
}

/* ------------------------------------------------------------------ slides */

/* 1 — cover */
function slide01(pptx) {
  const s = chrome(pptx, { corner: 'bl' });
  // Big mint "capsule" rotated behind the cover photo.
  const w = 6.0535, h = 5.5219;
  s.addShape('custGeom', {
    x: 8.4227, y: -0.3797, w: w, h: h, rotate: 133.2, fill: { color: C.mint },
    points: [
      { x: 2.3504, y: 5.5219, moveTo: true },
      { x: 0, y: 3.0191 },
      { x: 3.2148, y: 0 },
      { x: 4.1818, y: 0 },
      { x: w, y: 2.7608, curve: { type: 'cubic', x1: 5.2155, y1: 0, x2: w, y2: 1.2361 } },
      { x: 4.1818, y: h, curve: { type: 'cubic', x1: w, y1: 4.2856, x2: 5.2155, y2: h } },
      { close: true },
    ],
  });
  imageBox(s, 8.274, 1.272, 5.865, 4.839, { radius: 0.806 }); // right edge bleeds off-slide
  heading(s, [[['Medi', C.white], ['Cos', C.mint]]], 1.117, 1.943, 5.328, 1.313, { fontSize: 72 });
  body(s, LOREM_MED, 1.17, 3.553, 5.497, 0.603);
  card(s, 1.257, 4.553, 2.062, 0.628, C.mint);
  s.addText('Start Slide', {
    x: 1.388, y: 4.715, w: 1.799, h: 0.303,
    fontFace: HEAD, fontSize: 12, color: C.purple, align: 'center', charSpacing: 3, fit: 'resize',
  });
}

/* 2 — company introduction */
function slide02(pptx) {
  const s = chrome(pptx);
  card(s, -1.146, 4.003, 4.67, 3.521, C.mint);
  imageBox(s, 1.19, 1.396, 4.67, 3.922, { radius: 0.6537 });
  card(s, 2.399, 4.809, 2.25, 1.017, C.white, { shadow: SHADOW_TIGHT });
  s.addText(
    [
      { text: 'Learn More About', options: { breakLine: true } },
      { text: 'Our Company', options: { lineSpacingMultiple: 1.5 } },
    ],
    {
      x: 2.687, y: 5.058, w: 1.854, h: 0.577,
      fontFace: HEAD, fontSize: 12, bold: true, color: C.purple, align: 'justify', fit: 'resize',
    }
  );
  heading(s, [[['Company', C.white]], [['Introduction', C.mint]]], 7.391, 1.272, 5.6, 1.447);
  body(s, LOREM_XL, 7.391, 3.036, 4.752, 1.133);
  body(s, LOREM_XL, 7.391, 4.747, 4.774, 1.133);
}

/* 3 — vision & mission */
function slide03(pptx) {
  const s = chrome(pptx, { corner: 'bl' });
  card(s, 10.593, -0.002, 2.74, 2.177, C.mint, { shape: 'round1Rect', radius: 0.444, flipH: true, flipV: true });
  imageBox(s, 8.159, 1.272, 4.037, 4.749, { radius: 0.6728 });
  heading(s, [[['Vision & ', C.white], ['Mission', C.mint]]], 1.137, 1.272, 9.154, 0.774);
  body(s, LOREM_LONG + 'posuere Fusce posuere congue massa. ', 1.137, 2.481, 5.489, 0.868);
  [
    { x: 1.137, fill: C.mint, title: 'Company Vision' },
    { x: 5.175, fill: C.white, title: 'Company Mission' },
  ].forEach(function (c) {
    card(s, c.x, 3.75, 3.825, 1.534, c.fill, { shadow: SHADOW_TIGHT });
    label(s, c.title, c.x + 0.309, 4.008, 1.854, 0.303, { align: 'justify' });
    body(s, LOREM_SHORT, c.x + 0.309, 4.363, 3.189, 0.603, { color: C.purple });
  });
}

/* 4 — medical facilities */
function slide04(pptx) {
  const s = chrome(pptx, { markOnTop: true });
  card(s, -0.6458, 1.402, 3.7518, 4.626, C.mint, { radius: 0.6458 });
  imageBox(s, -0.82, 0, 3.582, 7.5, { radius: 0.82 });
  cornerDisc(s, 'tr', C.white);
  heading(s, [[['Our Medical', C.white]], [['Facilities', C.mint]]], 4.239, 1.444, 5.6, 1.447);
  body(s, LOREM_LONG, 4.239, 3.334, 3.65, 1.133);
  body(s, LOREM_LONG, 4.239, 4.848, 3.65, 1.133);
  const NUM_COPY = 'Lorem ipsum dolor sit amet, cectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce';
  [1.527, 3.334, 4.892].forEach(function (y, i) {
    label(s, ['01.', '02.', '03.'][i], 8.983, y + 0.02, 0.935, 0.572,
      { fontSize: 28, color: C.mint, align: 'justify' });
    body(s, NUM_COPY, 9.944, y, 2.221, 1.136);
  });
  wordmark(s);
}

/* 5 — latest product (white background) */
function slide05(pptx) {
  const s = chrome(pptx, { bg: C.white, corner: 'bl', cornerColor: C.purple, wordmark: C.purple });
  card(s, 7.239, 0, 7.041, 7.5, C.purple, { radius: 0.934 }); // right edge bleeds off-slide
  heading(s, [[['Our Latest', C.white]], [['Product', C.mint]]], 8.665, 1.276, 3.65, 1.447);
  body(s, LOREM_LONG, 8.665, 3.165, 3.65, 1.133);
  body(s, LOREM_LONG, 8.665, 4.68, 3.65, 1.133);
  [
    { img: [1.073, 1.42], txt: 3.675, y: 1.271, name: 'Product 1' },
    { img: [3.675, 3.011], txt: 1.073, y: 2.862, name: 'Product 2' },
    { img: [1.073, 4.829], txt: 3.675, y: 4.681, name: 'Product 3' },
  ].forEach(function (p) {
    imageBox(s, p.img[0], p.img[1], 2.181, 1.19, { radius: 0.1984 });
    label(s, p.name, p.txt, p.y + 0.12, 1.854, 0.303, { align: 'justify' });
    body(s, LOREM_SHORT, p.txt, p.y + 0.471, 2.181, 0.868, { color: C.deep });
  });
}

/* 6 — team */
function slide06(pptx) {
  const s = chrome(pptx);
  centeredTitle(s, 'Meet Our ', 'Team', 1.272);
  [
    { x: 1.911, y: 2.654, name: 'Juliana Silva' },
    { x: 1.907, y: 4.518, name: 'Chad Gibbons' },
    { x: 6.976, y: 2.65, name: 'Daniel Gallego' },
    { x: 6.972, y: 4.515, name: 'Isabel Mercado' },
  ].forEach(function (m) {
    card(s, m.x + 0.232, m.y, 3.825, 1.534, C.white, { shadow: SHADOW_TIGHT });
    card(s, m.x, m.y, 1.433, 1.534, C.mint, { shadow: SHADOW_TIGHT });
    imageBox(s, m.x + 0.1, m.y + 0.108, 1.232, 1.319, { radius: 0.2054 });
    label(s, m.name, m.x + 1.615, m.y + 0.328, 2.835, 0.303, { color: C.purple });
    body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing', m.x + 1.605, m.y + 0.656, 2.181, 0.603,
      { color: C.purple });
  });
}

/* 7 — reliable service */
function slide07(pptx) {
  const s = chrome(pptx, { corner: 'bl' });
  imageBox(s, 1.197, 1.425, 4.657, 4.7, { radius: 0.7758 });
  heading(s, [[['Company', C.white]], [['Reliable Service', C.mint]]], 7.546, 1.272, 5.6, 1.447);
  [
    { x: 5.122, y: 3.364, fill: C.mint, icon: iconBriefcase, ix: 5.742, iy: 3.738, iw: 0.413, ih: 0.354 },
    { x: 7.546, y: 3.323, fill: C.white, icon: iconPeople, ix: 7.954, iy: 3.686, iw: 0.444, ih: 0.416 },
    { x: 9.97, y: 3.323, fill: C.mint, icon: iconBell, ix: 10.521, iy: 3.686, iw: 0.382, ih: 0.413 },
  ].forEach(function (c) {
    card(s, c.x, c.y, 2.166, 2.163, c.fill, { shadow: SHADOW_TIGHT });
    c.icon(s, c.ix, c.iy, c.iw, c.ih, C.deep);
    label(s, 'Service 01', c.x + 0.483, c.y + 0.818, 1.527, 0.375, { align: 'justify', lineSpacingMultiple: 1.5 });
    body(s, 'Lorem ipsum dolor consect etuer.', c.x + 0.483, c.y + 1.168, 1.642, 0.606, { color: C.deep });
  });
}

/* 8 — portfolio */
function slide08(pptx) {
  const s = chrome(pptx);
  imageBox(s, 6.982, 1.397, 2.06, 4.706, { radius: 0.3433 });
  imageBox(s, 9.458, 1.397, 2.707, 2.889, { radius: 0.4511 });
  imageBox(s, 1.112, 2.8, 5.453, 3.303, { radius: 0.5504 });
  heading(s, [[['Medical ', C.white], ['Portfolio', C.mint]]], 1.192, 1.397, 9.154, 0.774);
  label(s, 'Medical Portfolio Detail', 9.458, 4.777, 2.707, 0.303, { color: C.mint, align: 'justify' });
  body(s, LOREM_SHORT, 9.458, 5.235, 2.707, 0.868);
}

/* 9 — price list */
function slide09(pptx) {
  const s = chrome(pptx);
  centeredTitle(s, 'Our Service ', 'Price List');
  // Side plans.
  [
    { x: 1.354, name: 'Reguler', price: '$50.00' },
    { x: 9.038, name: 'Platinum', price: '$100.00' },
  ].forEach(function (p, i) {
    const dy = i === 0 ? 0 : -0.02;
    card(s, p.x, 2.992 + dy, 2.941, 2.941, C.mint);
    s.addText(p.name, {
      x: p.x + 0.187, y: 3.182 + dy, w: 2.568, h: 0.462,
      fontFace: BODY, fontSize: 16, bold: true, color: C.purple, align: 'center',
      lineSpacingMultiple: 1.5, fit: 'resize',
    });
    s.addText(p.price, {
      x: p.x + 0.518, y: 3.477 + dy, w: 1.736, h: 0.642,
      fontFace: BODY, fontSize: 24, bold: true, color: C.purple, align: 'center',
      lineSpacingMultiple: 1.5, fit: 'resize',
    });
    s.addShape('line', { x: p.x + 0.381, y: 4.224 + dy, w: 2.16, h: 0, line: { color: C.purple, width: 1 } });
    body(s, 'Lorem ipsum dolor sit am et, consec tetuer', p.x + 0.381, 4.394 + dy, 2.16, 0.603,
      { color: C.purple, align: 'center' });
    card(s, p.x + 0.381, 5.18 + dy, 2.16, 0.418, C.purple, { radius: 0.209 });
    s.addText('ORDER NOW', {
      x: p.x + 0.381, y: 5.18 + dy, w: 2.16, h: 0.418,
      fontFace: BODY, fontSize: 10.5, bold: true, color: C.white, align: 'center', valign: 'middle',
      charSpacing: 3,
    });
  });
  // Featured plan.
  card(s, 4.625, 2.421, 4.084, 4.084, C.white);
  s.addText('Premium', {
    x: 4.999, y: 2.745, w: 3.347, h: 0.642,
    fontFace: BODY, fontSize: 24, bold: true, color: C.purple, align: 'center',
    lineSpacingMultiple: 1.5, fit: 'resize',
  });
  s.addText('$75.00', {
    x: 5.014, y: 2.972, w: 3.347, h: 1.319,
    fontFace: BODY, fontSize: 54, bold: true, color: C.purple, align: 'center',
    lineSpacingMultiple: 1.5, fit: 'resize',
  });
  s.addShape('line', { x: 5.145, y: 4.442, w: 3.044, h: 0, line: { color: C.purple, width: 1 } });
  body(s, 'Lorem ipsum dolor sit amet, consec tetuer adipiscing elit. ', 5.145, 4.69, 3.044, 0.603,
    { color: C.purple, align: 'center' });
  card(s, 5.145, 5.516, 3.044, 0.552, C.purple, { radius: 0.276 });
  s.addText('ORDER NOW', {
    x: 5.145, y: 5.516, w: 3.044, h: 0.552,
    fontFace: BODY, fontSize: 12, bold: true, color: C.white, align: 'center', valign: 'middle',
    charSpacing: 3,
  });
}

/* 10 — timeline cards */
function slide10(pptx) {
  const s = chrome(pptx);
  centeredTitle(s, 'Infographic ', 'Slide');
  [
    { x: 1.164, y: 2.655, fill: C.mint, year: '2021' },
    { x: 4.959, y: 2.655, fill: C.white, year: '2022' },
    { x: 8.745, y: 2.655, fill: C.mint, year: '2023' },
    { x: 1.173, y: 4.539, fill: C.white, year: '2024' },
    { x: 4.959, y: 4.539, fill: C.mint, year: '2025' },
    { x: 8.75, y: 4.539, fill: C.white, year: '2026' },
  ].forEach(function (c) {
    card(s, c.x, c.y, 3.415, 1.501, c.fill);
    s.addText(c.year, {
      x: c.x + 0.32, y: c.y + 0.325, w: 1.09, h: 0.559,
      fontFace: BODY, fontSize: 24, bold: true, color: C.purple, fit: 'resize',
    });
    body(s, 'Lorem ipsum do lor sit amet, con tetuer massa', c.x + 1.45, c.y + 0.292, 1.591, 0.96,
      { color: C.purple });
  });
}

/* 11 — donut cards */
function slide11(pptx) {
  const s = chrome(pptx);
  centeredTitle(s, 'Infographic ', 'Slide');
  [
    { x: 1.178, fill: C.mint, n: '01' },
    { x: 4.119, fill: C.white, n: '02' },
    { x: 7.06, fill: C.mint, n: '03' },
    { x: 10.001, fill: C.white, n: '04' },
  ].forEach(function (c) {
    // Top tile: the original punches a hole so the purple background shows
    // through as a ring; recreate the ring with a purple disc.
    card(s, c.x, 2.659, 2.18, 1.73, c.fill, { radius: 0.174 });
    s.addShape('ellipse', { x: c.x + 0.4256, y: 2.8576, w: 1.3338, h: 1.3338, fill: { color: C.purple } });
    badge(s, 'ellipse', c.x + 0.658, 3.09, 0.869, C.white, c.n, { fontSize: 18, color: C.purple, shadow: SHADOW_SOFT });
    // Bottom tile.
    card(s, c.x, 4.547, 2.18, 1.593, c.fill, { radius: 0.16 });
    s.addText('Title Here', {
      x: c.x, y: 4.889, w: 2.18, h: 0.337,
      fontFace: BODY, fontSize: 14, bold: true, color: C.purple, align: 'center', fit: 'resize',
    });
    body(s, 'Lorem ipsum dolor sitam, consec', c.x + 0.216, 5.181, 1.748, 0.603,
      { color: C.purple, align: 'center' });
  });
}

/* 12 — percentage cards */
function slide12(pptx) {
  const s = chrome(pptx);
  centeredTitle(s, 'Infographic ', 'Slide');
  [
    { x: 1.181, y: 2.648, fill: C.white, dash: C.mint, pct: '87%' },
    { x: 5.007, y: 2.648, fill: C.mint, dash: C.white, pct: '77%' },
    { x: 8.833, y: 2.648, fill: C.white, dash: C.mint, pct: '37%' },
    { x: 1.181, y: 4.51, fill: C.mint, dash: C.white, pct: '64%' },
    { x: 5.007, y: 4.51, fill: C.white, dash: C.mint, pct: '92%' },
    { x: 8.833, y: 4.51, fill: C.mint, dash: C.white, pct: '46%' },
  ].forEach(function (c) {
    card(s, c.x, c.y, 3.347, 1.583, c.fill, { radius: 0.13, shadow: SHADOW_SOFT });
    s.addShape('roundRect', {
      x: c.x + 0.389, y: c.y + 0.278, w: 0.662, h: 0.078, rectRadius: 0.039, fill: { color: c.dash },
    });
    s.addText(c.pct, {
      x: c.x + 0.278, y: c.y + 0.426, w: 1.03, h: 0.539,
      fontFace: BODY, fontSize: 26, bold: true, color: C.purple, fit: 'resize',
    });
    body(s, 'Lorem ipsum dolor sit amet, consect ip sum dolor ', c.x + 1.308, c.y + 0.39, 1.678, 0.896,
      { color: C.purple });
  });
}

/* 13 — numbered tiles + list */
function slide13(pptx) {
  const s = chrome(pptx);
  centeredTitle(s, 'Infographic ', 'Slide');
  [
    { x: 1.173, y: 2.667, fill: C.mint, n: '01' },
    { x: 4.956, y: 2.667, fill: C.white, n: '02' },
    { x: 8.735, y: 2.648, fill: C.mint, n: '03' },
  ].forEach(function (c) {
    card(s, c.x, c.y, 3.425, 1.555, c.fill);
    badge(s, 'roundRect', c.x + 0.322, c.y + 0.342, 0.854, C.purple, c.n,
      { transparency: 85, fontSize: 20, color: C.purple });
    label(s, 'Title Here', c.x + 1.388, c.y + 0.319, 1.777, 0.303, { color: C.purple });
    body(s, 'Lorem ipsum dolor sit amet, adip', c.x + 1.388, c.y + 0.619, 1.577, 0.603, { color: C.purple });
  });
  [
    { x: 1.179, y: 4.61, fill: C.mint, n: '01' },
    { x: 4.962, y: 4.61, fill: C.white, n: '02' },
    { x: 8.745, y: 4.61, fill: C.mint, n: '03' },
    { x: 1.173, y: 5.543, fill: C.white, n: '04' },
    { x: 4.956, y: 5.543, fill: C.mint, n: '05' },
    { x: 8.739, y: 5.543, fill: C.white, n: '06' },
  ].forEach(function (c) {
    badge(s, 'roundRect', c.x, c.y, 0.571, c.fill, c.n, { fontSize: 14, color: C.purple });
    body(s, 'Lorem ipsum dolor sit amet, con sect etuer adipiscing elit. ', c.x + 0.754, c.y - 0.016, 2.665, 0.603);
  });
}

/* 14 — document rows */
function slide14(pptx) {
  const s = chrome(pptx);
  centeredTitle(s, 'Infographic ', 'Slide');
  [
    { x: 1.645, y: 2.663, fill: C.mint },
    { x: 1.645, y: 3.869, fill: C.white },
    { x: 1.645, y: 5.074, fill: C.mint },
    { x: 7.499, y: 2.663, fill: C.white },
    { x: 7.499, y: 3.869, fill: C.mint },
    { x: 7.499, y: 5.074, fill: C.white },
  ].forEach(function (r) {
    const dot = r.fill === C.mint ? C.white : C.mint;
    card(s, r.x, r.y, 4.688, 1.041, r.fill);
    s.addShape('ellipse', {
      x: r.x - 0.4, y: r.y + 0.17, w: 0.703, h: 0.703, fill: { color: dot }, shadow: SHADOW_SOFT,
    });
    iconDoc(s, r.x - 0.174, r.y + 0.372, 0.252, 0.298, C.purple);
    s.addText('Document', {
      x: r.x + 0.638, y: r.y + 0.336, w: 1.369, h: 0.37,
      fontFace: BODY, fontSize: 16, bold: true, color: C.purple, wrap: false, fit: 'resize',
    });
    body(s, 'Lorem ipsum dolor sit am et, adipiscing', r.x + 2.211, r.y + 0.132, 2.124, 0.627,
      { color: C.purple, fontSize: 11, align: 'left' });
  });
}

/* 15 — globe tiles */
function slide15(pptx) {
  const s = chrome(pptx);
  centeredTitle(s, 'Infographic ', 'Slide');
  [
    { x: 1.173, fill: C.mint, n: '1', dy: 0 },
    { x: 3.995, fill: C.white, n: '2', dy: 0.006 },
    { x: 6.816, fill: C.mint, n: '3', dy: 0.013 },
    { x: 9.638, fill: C.white, n: '4', dy: 0.019 },
  ].forEach(function (c) {
    card(s, c.x, 2.648, 2.544, 2.544, c.fill, { shape: 'round2SameRect' });
    iconGlobe(s, c.x + 1.07, 3.187, 0.404, 0.404, C.purple);
    s.addText('Title Here', {
      x: c.x + 0.384, y: 3.771, w: 1.777, h: 0.303,
      fontFace: BODY, fontSize: 12, bold: true, color: C.purple, align: 'center', fit: 'resize',
    });
    body(s, 'Lorem ipsum dolor sit amet, dolor', c.x + 0.384, 4.041, 1.777, 0.627,
      { color: C.purple, fontSize: 11, align: 'center' });
    badge(s, 'ellipse', c.x, 5.58 + c.dy, 0.515, c.fill, c.n, { color: C.purple });
    body(s, 'Lorem ipsum dolor sitam consect adipiscing', c.x + 0.619, 5.508 + c.dy, 1.925, 0.603);
  });
}

/* 16 — KPI panels */
function slide16(pptx) {
  const s = chrome(pptx);
  centeredTitle(s, 'Infographic ', 'Slide');
  [
    { x: 1.244, fill: C.mint, stats: [['250K', 'Total Like'], ['6.90', 'Total Point'], ['98%', 'Percentage']] },
    { x: 6.778, fill: C.white, stats: [['560K', 'Total Like'], ['8.55', 'Total Point'], ['76%', 'Percentage']] },
  ].forEach(function (p) {
    card(s, p.x, 2.675, 5.318, 1.633, p.fill, { shape: 'round2SameRect' });
    [0.317, 2.142, 3.872].forEach(function (dx, i) {
      s.addText(p.stats[i][0], {
        x: p.x + dx, y: 3.078, w: i === 0 ? 1.468 : 1.173, h: 0.572,
        fontFace: BODY, fontSize: 28, bold: true, color: C.purple, align: 'justify', fit: 'resize',
      });
      s.addText(p.stats[i][1], {
        x: p.x + dx, y: 3.62, w: i === 0 ? 1.468 : 1.304, h: 0.286,
        fontFace: BODY, fontSize: 11, color: C.purple, align: 'justify', fit: 'resize',
      });
    });
    [1.849, 3.579].forEach(function (dx) {
      s.addShape('line', { x: p.x + dx, y: 3.119, w: 0, h: 0.739, line: { color: C.white, width: 1 } });
    });
  });
  [
    { x: 1.243, fill: C.mint, n: '01' },
    { x: 6.778, fill: C.white, n: '02' },
  ].forEach(function (b) {
    badge(s, 'ellipse', b.x, 4.865, 0.585, b.fill, b.n, { fontSize: 11, color: C.purple });
    label(s, 'Title Here', b.x + 0.744, 4.824, 1.355, 0.303, { color: C.white, align: 'justify' });
    body(s,
      'Lorem ipsum dolor sit amet ipsum dolor sit amet ipsum dolor sit amet ipsum dolor sit amet,. ' +
      'Maecenas porttmassa. ipsum dolor sit ipsum dolor',
      b.x + 0.744, 5.147, 4.301, 0.904, { fontSize: 11 });
  });
}

/* 17 — bar-chart tiles */
function slide17(pptx) {
  const s = chrome(pptx);
  centeredTitle(s, 'Infographic ', 'Slide');
  const BAR_TRACK = 1.326;
  const BAR_VALUES = [1.104, 0.594, 1.104, 0.83];
  [
    { x: 1.167, fill: C.mint, n: '01', dy: 0 },
    { x: 3.98, fill: C.white, n: '02', dy: 0.016 },
    { x: 6.792, fill: C.mint, n: '03', dy: 0.032 },
    { x: 9.604, fill: C.white, n: '04', dy: 0.048 },
  ].forEach(function (c) {
    card(s, c.x, 2.648, 2.569, 2.375, c.fill);
    [0.529, 0.958, 1.387, 1.811].forEach(function (dx, i) {
      s.addShape('roundRect', {
        x: c.x + dx, y: 2.984, w: 0.227, h: BAR_TRACK, rectRadius: 0.1135,
        fill: { color: C.purple, transparency: 85 },
      });
      s.addShape('roundRect', {
        x: c.x + dx, y: 2.984 + BAR_TRACK - BAR_VALUES[i], w: 0.227, h: BAR_VALUES[i], rectRadius: 0.1135,
        fill: { color: C.purple },
      });
    });
    body(s, 'Statistic Here', c.x + 0.161, 4.408, 2.248, 0.349, { color: C.purple, fontSize: 11, align: 'center' });
    badge(s, 'roundRect', c.x, 5.395 + c.dy, 0.571, c.fill, c.n, { fontSize: 14, color: C.purple });
    body(s, 'Lorem ipsum dolor sit amet, consect', c.x + 0.754, 5.379 + c.dy, 1.654, 0.603);
  });
}

/* 18 — single-corner-rounded panels */
function slide18(pptx) {
  const s = chrome(pptx);
  centeredTitle(s, 'Infographic ', 'Slide');
  [
    { x: 1.213, y: 2.673, fill: C.mint, n: '1', flipH: true },
    { x: 6.754, y: 2.673, fill: C.white, n: '2' },
    { x: 1.213, y: 4.474, fill: C.white, n: '3', flipH: true, flipV: true },
    { x: 6.754, y: 4.474, fill: C.mint, n: '4', flipV: true },
  ].forEach(function (c) {
    card(s, c.x, c.y, 5.366, 1.651, c.fill, { shape: 'round1Rect', flipH: c.flipH, flipV: c.flipV });
    badge(s, 'ellipse', c.x + 0.426, c.y + 0.342, 0.552, C.purple, c.n, { transparency: 80, color: C.purple });
    label(s, 'Title Here', c.x + 1.276, c.y + 0.298, 2.072, 0.372, { color: C.purple, align: 'justify' });
    body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing dolor sit amet, consectetuer',
      c.x + 1.276, c.y + 0.657, 3.629, 0.627, { color: C.purple, fontSize: 11 });
  });
}

/* 19 — contact (white background) */
function slide19(pptx) {
  const s = chrome(pptx, {
    bg: C.white, corner: 'bl', cornerColor: C.purple, arrowColor: C.purple, markOnTop: true,
  });
  imageBox(s, -0.934, 0, 7.026, 7.5, { radius: 0.934 });
  cornerDisc(s, 'bl', C.purple); // sits on top of the photo
  card(s, 10.593, -0.002, 2.74, 2.177, C.mint, { shape: 'round1Rect', radius: 0.444, flipH: true, flipV: true });
  card(s, 3.472, 1.137, 5.009, 1.025, C.purple, { shadow: SHADOW_TIGHT_TR });
  heading(s, [[['Contact', C.white], [' ', C.purple], ['Us', C.mint]]], 4.159, 1.262, 3.968, 0.774,
    { align: 'justify' });
  card(s, 3.472, 3.026, 8.693, 2.975, C.purple, { shadow: SHADOW_TIGHT_TR });
  [
    { x: 5.044, y: 3.361, title: 'Company Address', lines: ['Lorem ipsum dolor sit amet, consectetuer'],
      icon: iconBank, ix: 4.193, iy: 3.623, iw: 0.541, ih: 0.506 },
    { x: 9.511, y: 3.361, title: 'Chat With Us', lines: ['www.reallygreatsite.com', 'reallygreat@site.com'],
      icon: iconChat, ix: 8.688, iy: 3.648, iw: 0.54, ih: 0.47 },
    { x: 5.044, y: 4.707, title: 'Phone Number', lines: ['Phone 1 (+123) 456 7890', 'Phone 2 (+123) 456 7890'],
      icon: iconPhone, ix: 4.193, iy: 4.889, iw: 0.541, ih: 0.536 },
    { x: 9.511, y: 4.704, title: 'Office Hour', lines: ['Monday - Saturday', '08.00 AM \u2013 03.00 PM'],
      icon: iconBriefcase, ix: 8.688, iy: 4.926, iw: 0.541, ih: 0.463 },
  ].forEach(function (g) {
    g.icon(s, g.ix, g.iy, g.iw, g.ih, C.white);
    label(s, g.title, g.x, g.y, 2.325, 0.303, { color: C.white });
    const runs = g.lines.map(function (t, i) {
      return { text: t, options: { breakLine: i < g.lines.length - 1 } };
    });
    s.addText(runs, {
      x: g.x, y: g.y + 0.303, w: 2.655, h: 0.601,
      fontFace: BODY, fontSize: 10.5, color: C.white, align: 'justify',
      lineSpacingMultiple: 1.5, fit: 'resize',
    });
  });
  wordmark(s, C.purple);
}

/* 20 — thank you */
function slide20(pptx) {
  const s = chrome(pptx, { noArrow: true });
  imageBox(s, 1.117, 5.328, 6.343, 2.858, { radius: 0.686 }); // bottoms bleed off-slide
  imageBox(s, 9.021, 3.253, 4.313, 4.968, { radius: 0.7214 });
  heading(s, [[['Thank You', C.white]], [['For Your Attention', C.mint]]], 1.117, 1.315, 9.868, 1.447);
  body(s,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
    'Fusce posuere congue massa. Fusce posuere, magna sed consectet',
    1.17, 3.253, 6.29, 0.627);
}

/* -------------------------------------------------------------------- main */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'MEDICOS', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'MEDICOS';
  pptx.author = 'MediCos';
  pptx.title = 'MediCos Medical Company Presentation';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach(function (fn) { fn(pptx); });

  return pptx;
}

build()
  .writeFile({ fileName: path.join(__dirname, '16a31921-05e0-4f5f-9921-83642c2721fd_grok_final.pptx') })
  .then(function (f) { console.log('wrote ' + f); })
  .catch(function (e) { console.error(e); process.exit(1); });
