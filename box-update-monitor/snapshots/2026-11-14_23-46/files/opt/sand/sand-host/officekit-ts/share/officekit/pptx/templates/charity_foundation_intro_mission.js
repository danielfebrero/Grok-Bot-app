/*
 * Atlica — "Charity Presentation Template" (30 slides, 13.333in x 7.5in)
 * Rebuilt from scratch with pptxgenjs. Photographs in the original deck are
 * replaced by native-shape stand-ins (device frames / labelled rectangles).
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const C = {
  orange: 'F08214', // accent1
  burnt: 'BC640C', // accent2
  peach: 'F4A75A', // accent3
  brown: 'A1560B', // accent4
  white: 'FFFFFF', // accent5 / bg1
  gray: '7F7F7F', // accent6  - body copy
  ink: '191919', // tx1      - headings
  slate: '595959', // tx2      - sub headings
  mute: 'B2B2B2', // accent6 60% - captions
  track: 'E5E5E5', // accent6 20% - progress-bar track
  ghost: 'FCE6D0', // accent1 20% - oversized "01" numerals
  cream: 'FDEDDE', // accent3 20% - oversized quote mark
  silver: 'D9D9D9',
};

const MAJOR = 'Raleway Black'; // theme major font (headlines)
const MINOR = 'Roboto'; // theme minor font (everything else)

// pptxgenjs rewrites shadow options in place, so hand each shape its own copy.
const cardShadow = () => ({ type: 'outer', blur: 20, offset: 0, angle: 90, color: '000000', opacity: 0.1 });
const softShadow = () => ({ type: 'outer', blur: 25, offset: 0, angle: 90, color: '000000', opacity: 0.17 });

/* --------------------------------------------------------- reused strings */

const L = {
  // The template repeats a handful of lorem-ipsum blocks verbatim.
  full: "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially.",
  scrambled: "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled.",
  scrambledPrinting: "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled. printing and typesetting industry. ",
  scrambledSimply: "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled. printing and typesetting industry. simply dummy text of the printing. ",
  eversince: "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since",
  since1500: "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, ",
  standard: "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the",
  short: 'Lorem Ipsum is simply dummy text of the printing and typesetting industry. ',
  typese: 'Lorem Ipsum is simply dummy text of the printing and typese tting industry. Lorem Ipsum',
  stanArd: "Lorem Ipsum is simply dummy text of the printing and types etting industry. Lorem Ipsum has been the industry's stan ard dummy text ever since the 1500s, when an unknown.",
  unknownPrinting: "Lorem Ipsum is simply dummy text of the printing and types etting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printing and.",
  standardDummy: "Lorem Ipsum is simply dummy text of the printing and types etting industry. Lorem Ipsum has been the industry's standard dummy",
  printiNg: "Lorem Ipsum is simply dummy text of the printi ng and typesetting industry. Lorem Ipsum has been the industry's standard dummy ",
  ipSum: "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ip sum has been the industry's standa.",
  eyebrow: 'FOUNDATION FOR WORLD BETTER',
};

/* ---------------------------------------------------------------- helpers */

// Body copy: 11pt grey, 150% leading, justified (the deck's default paragraph).
function body(s, text, o) {
  s.addText(text, Object.assign(
    { fontFace: MINOR, fontSize: 11, color: C.gray, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top' }, o));
}

// Headline in the major font.
function heading(s, text, o) {
  s.addText(text, Object.assign({ fontFace: MAJOR, fontSize: 32, color: C.ink, valign: 'top' }, o));
}

// Bold sub-heading above a paragraph.
function subhead(s, text, o) {
  s.addText(text, Object.assign(
    { fontFace: MINOR, fontSize: 14, bold: true, color: C.slate, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top' }, o));
}

// Pill button (rounded rect + centred white caption).
function button(s, text, x, y, w, h, o) {
  const opt = o || {};
  s.addShape(opt.square ? 'rect' : 'roundRect', {
    x, y, w, h, rectRadius: opt.radius, fill: { color: opt.fill || C.orange },
  });
  s.addText(text, {
    x, y, w, h, align: 'center', valign: 'middle',
    fontFace: MINOR, fontSize: opt.fontSize || 12, color: opt.color || C.white,
  });
}

// Four brand dots + wordmark, top-left corner of most slides.
function brandMark(s) {
  [C.orange, C.burnt, C.peach, C.brown].forEach((col, i) => {
    s.addShape('ellipse', { x: 0.261 + i * 0.11, y: 0.329, w: 0.144, h: 0.144, fill: { color: col } });
  });
  s.addText('Charity For World Better', {
    x: 0.803, y: 0.279, w: 1.541, h: 0.252, fontFace: MINOR, fontSize: 9, color: C.mute, valign: 'top',
  });
}

// Peach accent bar hugging the lower-left edge.
function sideBar(s) {
  s.addShape('rect', { x: 0, y: 5.206, w: 0.05, h: 2.0, fill: { color: C.peach } });
}

// Peach accent bar on the top-right edge.
function topBar(s) {
  s.addShape('rect', { x: 9.698, y: 0, w: 3.0, h: 0.05, fill: { color: C.peach } });
}

// "Presentation Page N" footer with its hairline rule.
function pageNote(s, n) {
  s.addText('Presentation Page ' + n, {
    x: 9.9, y: 7.033, w: 3.42, h: 0.252, align: 'right', fontFace: MINOR, fontSize: 9, color: C.mute, valign: 'top',
  });
  s.addShape('line', { x: 10.425, y: 7.206, w: 1.349, h: 0, line: { color: C.peach, width: 0.75 } });
}

// Three graduated dots that trail the eyebrow label.
function dots(s, x, y, d) {
  const size = d || 0.057;
  const gap = size * 1.833;
  [C.burnt, C.peach, C.brown].forEach((col, i) => {
    s.addShape('ellipse', { x: x + i * gap, y, w: size, h: size, fill: { color: col } });
  });
}

// "FOUNDATION FOR WORLD BETTER" eyebrow. Left aligned variants trail three dots.
function eyebrow(s, x, y, o) {
  const opt = o || {};
  s.addText(L.eyebrow, {
    x, y, w: opt.w || 2.637, h: 0.286, fontFace: MINOR, fontSize: opt.fontSize || 11,
    bold: true, color: C.orange, valign: 'top', align: opt.align || 'left',
  });
  if (!opt.noDots) dots(s, x + (opt.dotDx || 2.691), y + 0.115, opt.dotSize);
}

const K = 0.4477; // cubic control-point ratio for a quarter circle

// Rectangle with only the named corners rounded ('tl','tr','br','bl').
function cornerRect(s, x, y, w, h, r, corners, fill) {
  const on = (c) => corners.indexOf(c) >= 0;
  const p = [];
  p.push({ x: on('tl') ? r : 0, y: 0 });
  p.push({ x: on('tr') ? w - r : w, y: 0 });
  if (on('tr')) p.push({ x: w, y: r, curve: { type: 'cubic', x1: w - r + r * K, y1: 0, x2: w, y2: r - r * K } });
  p.push({ x: w, y: on('br') ? h - r : h });
  if (on('br')) p.push({ x: w - r, y: h, curve: { type: 'cubic', x1: w, y1: h - r + r * K, x2: w - r + r * K, y2: h } });
  p.push({ x: on('bl') ? r : 0, y: h });
  if (on('bl')) p.push({ x: 0, y: h - r, curve: { type: 'cubic', x1: r - r * K, y1: h, x2: 0, y2: h - r + r * K } });
  p.push({ x: 0, y: on('tl') ? r : 0 });
  if (on('tl')) p.push({ x: r, y: 0, curve: { type: 'cubic', x1: 0, y1: r - r * K, x2: r - r * K, y2: 0 } });
  p.push({ close: true });
  s.addShape('custGeom', { x, y, w, h, fill: { color: fill }, points: p });
}

// Horizontal skill bar: light track + coloured fill + percentage label.
function skillBar(s, x, y, trackW, fillW, pct, pctColor) {
  s.addShape('roundRect', { x, y, w: trackW, h: 0.173, rectRadius: 0.0865, fill: { color: C.track } });
  s.addShape('roundRect', { x, y, w: fillW, h: 0.173, rectRadius: 0.0865, fill: { color: C.orange } });
  s.addText(pct, {
    x: x + trackW + 0.09, y: y - 0.057, w: 0.532, h: 0.286,
    fontFace: MINOR, fontSize: 11, bold: true, color: pctColor || C.orange, valign: 'top',
  });
}

// Bezel stand-in for the tablet / phone / monitor photographs. The screen is
// left unfilled so whatever sits behind (white page or orange panel) shows
// through, exactly as the transparent mockup PNGs do in the original.
function deviceFrame(s, x, y, w, h, radius, bezelPt, bezelColor) {
  const i = bezelPt / 144; // half the stroke, so the bezel sits inside x/y/w/h
  s.addShape('roundRect', {
    x: x + i, y: y + i, w: w - 2 * i, h: h - 2 * i, rectRadius: radius,
    line: { color: bezelColor || '1C1C1C', width: bezelPt },
  });
}

// Small pictogram stand-in for the deck's freeform icon art.
function icon(s, shape, x, y, w, h, color) {
  s.addShape(shape, { x, y, w, h, fill: { color } });
}

/* ------------------------------------------------------------ slide 1, 30 */

function titleSlide(s, opts) {
  cornerRect(s, 6.667, 3.757, 5.071, 3.743, 0.294, ['tl', 'tr'], C.orange); // layout panel
  s.addShape('ellipse', { x: 4.892, y: 1.465, w: 1.774, h: 1.774, fill: { color: C.peach } });
  brandMark(s);
  sideBar(s);
  s.addShape('donut', { x: 9.28, y: 5.052, w: 1.397, h: 1.397, rectRadius: 0.0499, fill: { color: C.peach } });
  s.addText(opts.title, { x: 0.575, y: opts.titleY, w: 6.509, h: 2.4, fontFace: MAJOR, fontSize: opts.titleSize, color: C.ink, valign: 'top' });
  s.addText(opts.kicker, {
    x: 0.57, y: opts.kickerY, w: 5.2, h: 0.5, fontFace: MINOR, fontSize: opts.kickerSize, bold: true, color: C.orange, valign: 'top',
  });
  body(s, 'Lorem Ipsum is simply dummy text of the printing and typesetting industry. standard dummy text ever since the 1500s',
    { x: 2.577, y: 5.052, w: 3.114, h: 0.903, align: 'left' });
  s.addShape('line', { x: 2.316, y: 4.793, w: 0, h: 1.768, line: { color: C.orange, width: 2.25 } });
}

function slide01(s) {
  titleSlide(s, { title: 'Atlica.', titleSize: 138, titleY: 1.56, kicker: 'CHARITY PRESENTATION TEMPLATE', kickerSize: 20, kickerY: 3.733 });
}

function slide30(s) {
  titleSlide(s, { title: 'Thanks', titleSize: 115, titleY: 1.727, kicker: 'AND SEE YOU NEXT TIME', kickerSize: 24, kickerY: 3.692 });
}

/* ---------------------------------------------------------------- slide 2 */

function slide02(s) {
  cornerRect(s, 6.667, 0, 4.021, 4.771, 0.3785, ['bl', 'br'], C.orange);
  brandMark(s);
  sideBar(s);
  heading(s, [{ text: 'About ' }, { text: 'Atlica' }, { text: ' Charity Foundation' }], { x: 0.611, y: 1.583, w: 4.479, h: 1.178 });
  eyebrow(s, 0.591, 2.797);
  s.addText('Welcome!', { x: 0.585, y: 3.664, w: 1.621, h: 0.404, fontFace: MINOR, fontSize: 18, bold: true, color: C.slate, valign: 'top' });
  s.addText('Let’s Us Introduce Our Company', {
    x: 1.797, y: 3.735, w: 3.251, h: 0.286, fontFace: MINOR, fontSize: 11, bold: true, color: C.burnt, valign: 'top',
  });
  body(s, L.full, { x: 0.611, y: 4.233, w: 5.35, h: 1.459 });
  s.addText('More Information', {
    x: 3.729, y: 5.917, w: 2.233, h: 0.348, align: 'right', fontFace: MINOR, fontSize: 11,
    bold: true, underline: { style: 'sng' }, color: C.burnt, lineSpacingMultiple: 1.5, valign: 'top',
  });
}

/* ---------------------------------------------------------------- slide 3 */

function slide03(s) {
  cornerRect(s, 1.646, 0, 4.083, 5.755, 0.4048, ['bl', 'br'], C.orange);
  topBar(s);
  pageNote(s, 3);
  heading(s, 'Who We Are ?', { x: 7.269, y: 1.44, w: 4.479, h: 0.64 });
  eyebrow(s, 7.274, 2.101);
  body(s, L.scrambled, { x: 7.269, y: 2.526, w: 5.471, h: 0.903 });
  [[7.269, 3.801], [10.098, 3.801], [7.269, 5.052], [10.098, 5.052]].forEach(([x, y]) => {
    body(s, L.typese, { x, y, w: 2.55, h: 0.903, bullet: { characterCode: '25AA', indent: 13.5 } });
  });
}

/* ---------------------------------------------------------------- slide 4 */

function slide04(s) {
  cornerRect(s, 7.646, 3.75, 5.687, 3.031, 0.2645, ['tl', 'bl'], C.orange);
  brandMark(s);
  sideBar(s);
  heading(s, 'Key Milestones ', { x: 0.539, y: 1.353, w: 4.479, h: 0.64 });
  eyebrow(s, 0.544, 2.041);
  body(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printe. is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard ",
    { x: 0.552, y: 2.672, w: 7.094, h: 0.903 });
  s.addText('1982-', {
    x: 3.171, y: 3.995, w: 3.671, h: 0.55, fontFace: MINOR, fontSize: 20, bold: true, color: C.orange, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top',
  });
  body(s, [
    { text: 'Lorem Ipsum is simply ', options: { bold: true, color: C.slate } },
    { text: "dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a" },
  ], { x: 3.208, y: 4.507, w: 3.671, h: 1.181 });
  button(s, 'More Detail', 3.306, 5.857, 1.647, 0.425, { radius: 0.0523, fontSize: 14 });
}

/* ---------------------------------------------------------------- slide 5 */

function slide05(s) {
  s.addShape('roundRect', { x: 3.688, y: 2.271, w: 2.958, h: 4.021, rectRadius: 0.1779, fill: { color: C.orange } });
  topBar(s);
  pageNote(s, 5);
  icon(s, 'noSmoking', 1.049, 4.706, 0.419, 0.419, C.orange);
  s.addText('Useful Product', { x: 0.92, y: 5.237, w: 2.476, h: 0.37, fontFace: MINOR, fontSize: 16, bold: true, color: C.slate, valign: 'top' });
  body(s, 'Charity Foundation Present', { x: 0.934, y: 5.501, w: 2.392, h: 0.348, italic: true, color: C.mute });
  icon(s, 'star5', 4.162, 4.682, 0.436, 0.448, C.white);
  s.addText('Trusted Foundation ', { x: 4.094, y: 5.23, w: 3.336, h: 0.37, fontFace: MINOR, fontSize: 16, bold: true, color: C.white, valign: 'top' });
  body(s, 'Charity Foundation Present', { x: 4.094, y: 5.481, w: 2.392, h: 0.348, italic: true, color: C.white });

  heading(s, [{ text: 'About ' }, { text: 'Atlica’s' }, { text: ' Vision Statement' }], { x: 7.071, y: 1.464, w: 4.479, h: 1.178 });
  eyebrow(s, 7.076, 2.678);
  body(s, L.eversince, { x: 7.103, y: 3.06, w: 5.661, h: 0.625 });
  subhead(s, 'What We Want To Do ?', { x: 7.103, y: 4.078, w: 5.595, h: 0.415 });
  body(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged",
    { x: 7.103, y: 4.514, w: 5.595, h: 1.459 });
}

/* ---------------------------------------------------------------- slide 6 */

function slide06(s) {
  s.addShape('roundRect', { x: 6.708, y: 3.271, w: 2.958, h: 3.479, rectRadius: 0.1779, fill: { color: C.orange } });
  brandMark(s);
  sideBar(s);
  heading(s, [{ text: 'About ' }, { text: 'Atlica’s' }, { text: ' Charity Foundation Mission ' }], { x: 0.588, y: 1.443, w: 5.423, h: 1.178 });
  eyebrow(s, 0.593, 2.657);
  icon(s, 'heart', 0.688, 3.457, 0.495, 0.495, C.orange);
  subhead(s, 'Simple Description 01', { x: 1.387, y: 3.244, w: 4.624, h: 0.46, fontSize: 16 });
  body(s, L.stanArd, { x: 1.387, y: 3.736, w: 4.624, h: 0.903 });
  s.addText('+892', { x: 0.553, y: 4.879, w: 0.981, h: 0.505, fontFace: MINOR, fontSize: 18, bold: true, color: C.orange, valign: 'top' });
  subhead(s, 'Simple Description 02', { x: 1.387, y: 4.887, w: 4.624, h: 0.46, fontSize: 16 });
  body(s, L.stanArd, { x: 1.387, y: 5.351, w: 4.624, h: 0.903 });
  body(s, [
    { text: '“Always Give Without Remembering And Always Receive Without Forgetting.” - ', options: { italic: true } },
    { text: 'Brian Tracy', options: { italic: true, bold: true } },
  ], { x: 6.862, y: 5.351, w: 2.65, h: 0.903, align: 'center', color: C.white });
}

/* ---------------------------------------------------------------- slide 7 */

function slide07(s) {
  s.addShape('roundRect', { x: 7.708, y: 1.729, w: 4.967, h: 5.104, rectRadius: 0.2987, fill: { color: C.orange } });
  heading(s, [{ text: 'About ' }, { text: 'Atlica’s' }, { text: ' Strength & Value' }], { x: 0.561, y: 4.09, w: 4.479, h: 1.178 });
  eyebrow(s, 0.566, 5.304);
  body(s, L.scrambledPrinting, { x: 0.561, y: 5.59, w: 6.72, h: 0.903 });

  icon(s, 'moon', 8.098, 2.643, 0.486, 0.384, C.white);
  subhead(s, 'Dontate Now To Brighten The Future ', { x: 8.689, y: 2.432, w: 3.689, h: 0.415, color: C.white });
  body(s, L.eversince, { x: 8.689, y: 2.847, w: 3.597, h: 0.903, color: C.white });
  s.addShape('line', { x: 8.177, y: 4.306, w: 4.109, h: 0, line: { color: C.white, width: 2.25, dashType: 'dash' } });
  icon(s, 'cube', 8.206, 4.935, 0.357, 0.427, C.white);
  subhead(s, 'Simple Movement To The Future', { x: 8.914, y: 4.788, w: 3.689, h: 0.415, color: C.white });
  body(s, 'Lorem Ipsum is simply dummy text of the prin ting and type setting industry. Lorem Ipsu typ esetting industry. Lorem Ipsum has been',
    { x: 8.914, y: 5.174, w: 3.372, h: 0.903, color: C.white });
}

/* ---------------------------------------------------------------- slide 8 */

function slide08(s) {
  cornerRect(s, 2.679, 2.771, 3.988, 4.729, 0.3939, ['tl', 'tr'], C.orange);
  topBar(s);
  heading(s, [{ text: 'About ' }, { text: 'Atlica' }, { text: ' Focus & Foundation Plan' }], { x: 7.309, y: 1.306, w: 4.479, h: 1.178 });
  eyebrow(s, 7.314, 2.52);
  body(s, L.full, { x: 7.309, y: 3.331, w: 5.508, h: 1.459 });
  s.addText('89,55%', { x: 7.309, y: 5.112, w: 1.513, h: 0.505, fontFace: MINOR, fontSize: 24, bold: true, color: C.burnt, valign: 'top' });
  s.addText('Dontation’s Report', { x: 7.309, y: 5.503, w: 2.0, h: 0.286, fontFace: MINOR, fontSize: 11, italic: true, color: C.gray, valign: 'top' });
  body(s, L.eversince, { x: 8.903, y: 5.044, w: 3.801, h: 0.903 });
}

/* ---------------------------------------------------------------- slide 9 */

function slide09(s) {
  s.addShape('roundRect', { x: 0.646, y: 3.167, w: 4.848, h: 3.021, rectRadius: 0.3067, fill: { color: C.orange } });
  brandMark(s);
  sideBar(s);
  topBar(s);
  heading(s, [{ text: 'About ' }, { text: 'Atlica’s' }, { text: ' For Education Here ' }], { x: 0.646, y: 1.128, w: 4.479, h: 1.178 });
  eyebrow(s, 0.651, 2.342);
  [['01/', 5.523, 5.558], ['02/', 9.315, 9.295]].forEach(([num, nx, tx]) => {
    body(s, [
      { text: num, options: { bold: true, color: C.orange, fontSize: 18 } },
      { text: '02', options: { color: C.gray, fontSize: 11 } },
    ], { x: nx, y: 1.04, w: 3.403, h: 0.505 });
    body(s, L.printiNg, { x: tx, y: 1.485, w: 3.403, h: 0.903 });
  });
  icon(s, 'star4', 1.286, 3.708, 0.511, 0.552, C.white);
  s.addText('Give And Make A Little Happiness To The Futures Child', {
    x: 1.989, y: 3.647, w: 2.82, h: 0.673, fontFace: MINOR, fontSize: 12, bold: true, color: C.white, lineSpacingMultiple: 1.5, valign: 'top',
  });
  body(s, "Lorem Ipsum is simply dummy text of the printi ng and typesetting industry. Lorem Ipsum has been the industry's standard d dummy text of the printi ng and typesetting industry. Lorem ummy text ever",
    { x: 1.158, y: 4.414, w: 3.934, h: 1.181, color: C.white });
}

/* --------------------------------------------------------------- slide 10 */

function slide10(s) {
  s.addShape('rect', { x: 0, y: 0, w: 5.167, h: 7.5, fill: { color: C.orange } });
  topBar(s);
  heading(s, [{ text: 'About ' }, { text: 'Atlica’s' }, { text: ' Charity Programs' }], { x: 7.243, y: 1.212, w: 4.479, h: 1.178 });
  eyebrow(s, 7.248, 2.426);
  const chunk = "Lorem Ipsum is simply dummy text of the printing and types etting industry. Lorem Ipsum has been the industry's stan ard dummy text ever since the 1500s, when an unknown. printing and types etting industry. ";
  [['01', 2.66, 3.142, 3.634], ['02', 4.45, 4.921, 5.385]].forEach(([num, numY, headY, textY]) => {
    body(s, num, { x: 6.985, y: numY, w: 1.677, h: 1.178, fontSize: 48, bold: true, color: C.ghost });
    subhead(s, 'Simple Description', { x: 7.243, y: headY, w: 4.624, h: 0.46, fontSize: 16 });
    body(s, chunk, { x: 7.243, y: textY, w: 5.455, h: 0.903 });
  });
}

/* --------------------------------------------------------------- slide 11 */

function slide11(s) {
  brandMark(s);
  sideBar(s);
  heading(s, 'About Our Domestic Charity Program ', { x: 0.712, y: 1.501, w: 5.445, h: 1.178 });
  eyebrow(s, 0.717, 2.715);
  s.addShape('ellipse', { x: 0.814, y: 3.704, w: 0.892, h: 0.892, fill: { color: C.orange }, shadow: softShadow() });
  icon(s, 'irregularSeal1', 1.104, 4.01, 0.311, 0.309, C.white);
  subhead(s, 'Trusted Charity Foundation ', { x: 1.797, y: 3.542, w: 4.36, h: 0.415 });
  body(s, "Lorem Ipsum is simply dummy text of the printing and typese tting industry. Lorem Ipsum has been the industry's", { x: 1.797, y: 3.93, w: 4.36, h: 0.625 });
  body(s, L.scrambledSimply, { x: 0.712, y: 4.735, w: 5.512, h: 1.181 });

  s.addShape('roundRect', { x: 8.692, y: 0.729, w: 4.01, h: 5.042, rectRadius: 0.306, fill: { color: C.orange } });
  icon(s, 'flowChartConnector', 9.194, 1.313, 0.614, 0.614, C.white);
  body(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ips um has been the industry's standard", { x: 9.145, y: 1.941, w: 3.195, h: 0.903, color: C.white });
  s.addShape('line', { x: 9.27, y: 3.178, w: 0.775, h: 0, line: { color: C.white, width: 1.5, endArrowType: 'triangle' } });
}

/* --------------------------------------------------------------- slide 12 */

function slide12(s) {
  cornerRect(s, 1.691, 2.771, 3.956, 4.729, 0.3685, ['tl', 'tr'], C.orange);
  topBar(s);
  heading(s, 'International Program', { x: 6.206, y: 1.074, w: 5.168, h: 0.64 });
  eyebrow(s, 6.206, 1.701, { dotDx: 2.61 });
  body(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived",
    { x: 6.185, y: 2.083, w: 6.626, h: 0.903 });
  [['70 %', C.orange, 3.31, 3.422, 3.791], ['40 %', C.peach, 4.878, 5.048, 5.417]].forEach(([pct, col, pctY, headY, textY]) => {
    s.addText(pct, { x: 7.271, y: pctY, w: 1.238, h: 0.725, fontFace: MAJOR, fontSize: 28, bold: true, color: col, lineSpacingMultiple: 1.5, valign: 'top' });
    s.addText(' Simple Description', { x: 8.57, y: headY, w: 2.676, h: 0.369, fontFace: MINOR, fontSize: 14, bold: true, color: C.slate, lineSpacing: 20.8, valign: 'top', margin: [0, 7.2, 3.6, 3.6] });
    body(s, L.since1500, { x: 8.53, y: textY, w: 4.123, h: 0.903 });
  });
}

/* --------------------------------------------------------------- slide 13 */

function slide13(s) {
  s.addShape('ellipse', { x: 5.125, y: 1.979, w: 3.083, h: 3.083, fill: { color: C.orange } });
  brandMark(s);
  sideBar(s);
  topBar(s);
  heading(s, 'Governing Board', { x: 4.427, y: 0.729, w: 4.479, h: 0.64, align: 'center' });
  eyebrow(s, 5.348, 1.351, { noDots: true, align: 'center' });
  const bio = "Lorem Ipsum is simply dummy text of the printing and types etting industry. Lorem Ipsum has been the industry's stan ard dummy text ever since the 1500s, when an";
  [['Christian Rasford', 0.405, 0.98, 5.206, 5.669], ['James William', 4.355, 4.93, 5.21, 5.673], ['Keylor Martens', 8.304, 8.88, 5.21, 5.673]]
    .forEach(([name, nx, bx, ny, by]) => {
      s.addText(name, { x: nx, y: ny, w: 4.624, h: 0.46, align: 'center', fontFace: MINOR, fontSize: 16, bold: true, color: C.slate, lineSpacingMultiple: 1.5, valign: 'top' });
      body(s, bio, { x: bx, y: by, w: 3.473, h: 1.181, align: 'center' });
    });
  s.addShape('halfFrame', { x: 6.552, y: 6.853, w: 0.23, h: 0.23, rotate: 225, fill: { color: C.orange } });
}

/* --------------------------------------------------------------- slide 14 */

function slide14(s) {
  s.addShape('roundRect', { x: 2.601, y: 4.292, w: 4.0, h: 2.104, rectRadius: 0.2136, fill: { color: C.orange } });
  s.addShape('ellipse', { x: 2.971, y: 4.823, w: 1.042, h: 1.042, fill: { color: C.white, transparency: 75 } });
  icon(s, 'flowChartConnector', 3.297, 5.136, 0.389, 0.415, C.white);
  topBar(s);
  heading(s, 'Meet Charity’s CEO & Founder Here', { x: 5.189, y: 1.28, w: 4.479, h: 1.178 });
  eyebrow(s, 5.189, 2.494);
  body(s, L.scrambledSimply, { x: 5.181, y: 2.883, w: 7.465, h: 0.903 });
  s.addText([{ text: 'Christian ' }, { text: 'Rasford' }], {
    x: 4.119, y: 4.721, w: 2.123, h: 0.415, fontFace: MINOR, fontSize: 14, bold: true, color: C.white, lineSpacingMultiple: 1.5, valign: 'top',
  });
  body(s, 'Lorem Ipsum is simply dum my text of the printing and type setting industry. ', { x: 4.127, y: 5.085, w: 2.123, h: 0.903, color: C.white });
  body(s, 'Profile Description:', { x: 6.927, y: 4.191, w: 2.23, h: 0.348 });
  subhead(s, 'About Skill Here', { x: 6.971, y: 4.65, w: 2.186, h: 0.415 });
  skillBar(s, 6.977, 5.05, 5.045, 4.189, '92%', C.burnt);
  subhead(s, 'About Experience', { x: 6.986, y: 5.534, w: 2.186, h: 0.415 });
  skillBar(s, 6.998, 5.933, 5.045, 4.318, '93%', C.burnt);
}

/* --------------------------------------------------------------- slide 15 */

function slide15(s) {
  s.addShape('roundRect', { x: 3.611, y: 0.729, w: 3.056, h: 5.042, rectRadius: 0.3102, fill: { color: C.orange } });
  topBar(s);
  s.addText('James William', { x: 4.119, y: 1.269, w: 2.123, h: 0.415, fontFace: MINOR, fontSize: 14, bold: true, color: C.white, lineSpacingMultiple: 1.5, valign: 'top' });
  body(s, 'Lorem Ipsum is simply dum my text of the printing.', { x: 4.127, y: 1.633, w: 2.123, h: 0.625, color: C.white });
  heading(s, [{ text: 'Our Board Member ', options: { breakLine: true } }, { text: '& Treasurer' }], { x: 7.175, y: 1.348, w: 5.004, h: 1.178 });
  eyebrow(s, 7.189, 2.561);
  body(s, L.scrambledSimply, { x: 7.181, y: 2.95, w: 5.461, h: 1.181 });
  body(s, 'Profile Description:', { x: 7.189, y: 4.348, w: 2.23, h: 0.348, italic: true });
  subhead(s, 'About Skill Here', { x: 7.245, y: 4.82, w: 2.186, h: 0.415 });
  skillBar(s, 7.252, 5.219, 4.873, 4.251, '92%');
  subhead(s, 'About Experience', { x: 7.26, y: 5.703, w: 2.186, h: 0.415 });
  skillBar(s, 7.272, 6.103, 4.873, 3.917, '93%');
}

/* --------------------------------------------------------------- slide 16 */

function slide16(s) {
  s.addShape('rect', { x: 9.667, y: 0, w: 3.667, h: 7.5, fill: { color: C.orange } });
  s.addShape('roundRect', { x: 4.199, y: 4.05, w: 4.754, h: 2.214, rectRadius: 0.2248, fill: { color: C.white }, shadow: softShadow() });
  brandMark(s);
  sideBar(s);
  heading(s, [{ text: 'About ' }, { text: 'Atlica’s' }, { text: ' Board & Secretary' }], { x: 0.758, y: 1.236, w: 4.479, h: 1.178 });
  eyebrow(s, 0.809, 2.449);
  body(s, L.scrambledPrinting, { x: 0.758, y: 2.848, w: 6.286, h: 0.903 });
  s.addText([{ text: 'Keylor' }, { text: ' Martens' }], {
    x: 0.767, y: 3.99, w: 2.9, h: 0.415, fontFace: MINOR, fontSize: 14, bold: true, color: C.slate, lineSpacingMultiple: 1.5, valign: 'top',
  });
  body(s, L.eversince, { x: 0.785, y: 4.405, w: 2.9, h: 1.181 });
  button(s, 'More Info', 0.868, 5.803, 1.383, 0.395, { radius: 0.0401 });
  subhead(s, 'About Skill Here', { x: 4.621, y: 4.299, w: 2.186, h: 0.415 });
  skillBar(s, 4.628, 4.698, 3.411, 2.975, '92%');
  subhead(s, 'About Experience', { x: 4.637, y: 5.183, w: 2.186, h: 0.415 });
  skillBar(s, 4.642, 5.582, 3.411, 2.742, '93%');
}

/* --------------------------------------------------------------- slide 17 */

function slide17(s) {
  brandMark(s);
  sideBar(s);
  topBar(s);
  pageNote(s, 17);
  heading(s, 'Pricing Tables', { x: 2.611, y: 0.729, w: 8.111, h: 0.707, fontSize: 36, align: 'center' });
  eyebrow(s, 5.348, 1.351, { noDots: true, align: 'center' });
  s.addText('We Bring The Best Solution', {
    x: 5.553, y: 1.756, w: 2.227, h: 0.286, align: 'center', fontFace: MINOR, fontSize: 11, color: C.burnt, valign: 'top',
  });
  const plans = [
    { name: 'The Pricing One', price: '$30', per: '/month', cardX: 1.956, nameX: 2.273, nameW: 2.157, colX: 2.189, priceX: 2.25, btnX: 2.589 },
    { name: 'The Pricing Two', price: '$39', per: '/ month', cardX: 5.272, nameX: 5.575, nameW: 2.183, colX: 5.505, priceX: 5.566, btnX: 5.905 },
    { name: 'The Pricing Three', price: '$40', per: '/ month', cardX: 8.588, nameX: 8.792, nameW: 2.381, colX: 8.82, priceX: 8.882, btnX: 9.221 },
  ];
  plans.forEach((p) => {
    s.addShape('rect', { x: p.cardX, y: 2.451, w: 2.789, h: 3.983, fill: { color: C.white }, shadow: cardShadow() });
    s.addText(p.name, { x: p.nameX, y: 2.76, w: p.nameW, h: 0.404, align: 'center', fontFace: MAJOR, fontSize: 18, color: C.orange, valign: 'top' });
    s.addText('Detail Pricing Here', {
      x: p.btnX, y: 3.085, w: 1.524, h: 0.307, align: 'center', fontFace: MINOR, fontSize: 9, italic: true, color: C.slate, lineSpacingMultiple: 1.5, valign: 'top',
    });
    s.addText([
      { text: p.price, options: { fontSize: 44 } },
      { text: p.per, options: { fontSize: 11 } },
    ], { x: p.priceX, y: 3.354, w: 2.201, h: 1.107, align: 'center', fontFace: MAJOR, color: C.orange, lineSpacingMultiple: 1.5, valign: 'top' });
    [['Enim nec dui nunc', 4.493, 1.924], ['Magna sit amet risus pro', 4.792, 2.359], ['Risus pretium quam', 5.091, 2.201]].forEach(([t, y, w]) => {
      s.addText(t, {
        x: p.colX, y, w, h: 0.349, fontFace: MINOR, fontSize: 11, color: C.slate,
        lineSpacingMultiple: 1.5, valign: 'top', bullet: { characterCode: '2713', indent: 13.5 },
      });
    });
    button(s, 'Get It Now!', p.btnX, 5.73, 1.524, 0.419, { square: true, fontSize: 10 });
  });
}

/* --------------------------------------------------------------- slide 18 */

function slide18(s) {
  s.addShape('rect', { x: 7.458, y: 0, w: 5.875, h: 7.5, fill: { color: C.orange } });
  brandMark(s);
  sideBar(s);
  heading(s, 'Our Goal Is To Raise Their Comprehension In Educational', { x: 1.083, y: 1.542, w: 4.957, h: 1.717 });
  eyebrow(s, 1.097, 3.463);
  body(s, L.full, { x: 1.079, y: 3.994, w: 5.587, h: 1.459 });
  body(s, 'Lorem Ipsum is simply dummy text of the printing and typesetting into electronic typesetting, remaining essentially.',
    { x: 1.079, y: 5.686, w: 4.857, h: 0.625, italic: true, color: C.orange });
  s.addText('Our Focus & Charity Plan', {
    x: 8.233, y: 4.386, w: 4.325, h: 0.46, align: 'center', fontFace: MINOR, fontSize: 16, bold: true, color: C.white, lineSpacingMultiple: 1.5, valign: 'top',
  });
  body(s, L.since1500, { x: 8.233, y: 4.905, w: 4.325, h: 0.903, align: 'center', color: C.white });
  s.addShape('roundRect', { x: 9.54, y: 6.058, w: 1.943, h: 0.507, rectRadius: 0.0515, fill: { color: C.white, transparency: 68 } });
  s.addText('More Information ', { x: 9.54, y: 6.058, w: 1.943, h: 0.507, align: 'center', valign: 'middle', fontFace: MINOR, fontSize: 12, color: C.white });
}

/* --------------------------------------------------------------- slide 19 */

function slide19(s) {
  s.addShape('rect', { x: 2.688, y: 0, w: 3.979, h: 7.5, fill: { color: C.orange } });
  topBar(s);
  heading(s, 'About Community Education Program', { x: 7.348, y: 1.388, w: 4.479, h: 1.178 });
  eyebrow(s, 7.399, 2.602);
  body(s, "Lorem Ipsum is simply dummy text of the printing and typesetting ind ustry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley.",
    { x: 7.399, y: 2.966, w: 5.247, h: 0.903 });
  s.addText('Options Here :', {
    x: 7.399, y: 4.061, w: 5.247, h: 0.37, fontFace: MINOR, fontSize: 12, bold: true, color: C.slate, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top',
  });
  body(s, "Lorem Ipsum is simply dummy text of the printing and typesetting ind ustry. Lorem Ipsum has been the industry's standard dummy.",
    { x: 7.399, y: 4.59, w: 5.043, h: 0.625, bullet: { characterCode: '2713', indent: 13.5 } });
  body(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a",
    { x: 7.399, y: 5.374, w: 5.043, h: 0.903, bullet: { characterCode: '2713', indent: 13.5 } });
}

/* --------------------------------------------------------------- slide 20 */

function slide20(s) {
  s.addShape('rect', { x: 7.646, y: 0, w: 5.688, h: 7.5, fill: { color: C.orange } });
  brandMark(s);
  sideBar(s);
  [['Community Education', 0.848, 1.349, 2.401, 2.959, 9.762], ['International Program', 4.016, 4.517, 5.57, 6.128, 9.733]]
    .forEach(([title, ty, py, sy, by, x]) => {
      s.addText(title, { x, y: ty, w: 3.151, h: 0.46, fontFace: MINOR, fontSize: 16, bold: true, color: C.white, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top' });
      body(s, L.ipSum, { x, y: py, w: 3.151, h: 0.903, color: C.white });
      body(s, '1500s, - Lorem Ipsum is simply dummy', { x, y: sy, w: 3.151, h: 0.348, color: C.white });
      s.addShape('roundRect', { x: x + 0.111, y: by, w: 1.566, h: 0.453, rectRadius: 0.046, fill: { color: C.white, transparency: 68 } });
      s.addText('Detail Programs', { x: x + 0.111, y: by, w: 1.566, h: 0.453, align: 'center', valign: 'middle', fontFace: MINOR, fontSize: 12, color: C.white });
    });
  heading(s, [{ text: 'About ' }, { text: 'Atlica’s' }, { text: ' ', options: { breakLine: true } }, { text: 'Main Programs' }], { x: 0.822, y: 1.397, w: 4.479, h: 1.178 });
  eyebrow(s, 0.856, 2.632);
  body(s, L.scrambledPrinting, { x: 0.856, y: 3.283, w: 5.212, h: 1.181 });
  [4.658, 5.477].forEach((y) => {
    body(s, "Lorem Ipsum is simply dummy text of the printing Lorem Ipsum has and typesetting industry. been the industry's standard",
      { x: 0.856, y, w: 4.82, h: 0.625, bullet: { characterCode: '25AA', indent: 13.5 } });
  });
}

/* --------------------------------------------------------------- slide 21 */

function slide21(s) {
  heading(s, 'Supporting & Partnership', { x: 2.611, y: 0.729, w: 8.111, h: 0.64, align: 'center' });
  eyebrow(s, 5.348, 1.351, { noDots: true, align: 'center' });
  const cards = [
    { x: 4.759, y: 2.098, tx: 5.091, big: '1,3 Million', bigY: 2.402, label: 'Aspects 02', labelY: 2.863, textY: 3.308 },
    { x: 0.615, y: 4.563, tx: 0.946, label: 'Aspects 01', labelY: 5.353, textY: 5.799, iconX: 2.287, iconY: 4.917 },
    { x: 8.904, y: 4.563, tx: 9.235, big: '89,52%', bigY: 4.865, label: 'Aspects 03', labelY: 5.353, textY: 5.799 },
  ];
  cards.forEach((c) => {
    s.addShape('roundRect', { x: c.x, y: c.y, w: 3.815, h: 2.25, rectRadius: 0.1253, fill: { color: C.orange } });
    if (c.big) {
      s.addText(c.big, { x: c.tx, y: c.bigY, w: 3.151, h: 0.55, align: 'center', fontFace: MINOR, fontSize: 20, bold: true, color: C.white, lineSpacingMultiple: 1.5, valign: 'top' });
    }
    if (c.iconX) icon(s, 'star5', c.iconX, c.iconY, 0.471, 0.443, C.white);
    s.addText(c.label, { x: c.tx, y: c.labelY, w: 3.151, h: 0.46, align: 'center', fontFace: MINOR, fontSize: 16, color: C.white, lineSpacingMultiple: 1.5, valign: 'top' });
    body(s, L.short, { x: c.tx, y: c.textY, w: 3.151, h: 0.625, align: 'center', color: C.white });
  });
}

/* --------------------------------------------------------------- slide 22 */

function slide22(s) {
  s.addShape('roundRect', { x: 2.537, y: 3.396, w: 2.692, h: 3.312, rectRadius: 0.229, fill: { color: C.orange } });
  topBar(s);
  heading(s, [{ text: 'About ' }, { text: 'Atlica’s' }, { text: ' Volunteer Member' }], { x: 7.35, y: 1.571, w: 4.479, h: 1.178 });
  eyebrow(s, 7.384, 2.806);
  body(s, L.scrambledPrinting, { x: 7.384, y: 3.457, w: 5.212, h: 1.181 });
  body(s, [
    { text: '832,1- ', options: { bold: true, color: C.slate, fontSize: 12 } },
    { text: "dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of" },
  ], { x: 7.35, y: 4.881, w: 5.212, h: 0.928 });
  s.addText('Simple Text Here', { x: 2.68, y: 5.383, w: 2.406, h: 0.415, align: 'center', fontFace: MINOR, fontSize: 14, bold: true, color: C.white, lineSpacingMultiple: 1.5, valign: 'top' });
  body(s, 'Lorem Ipsum is simply dummy text of the printing and', { x: 2.68, y: 5.81, w: 2.406, h: 0.625, align: 'center', color: C.white });
}

/* --------------------------------------------------------------- slide 23 */

function slide23(s) {
  s.addShape('roundRect', { x: 3.699, y: 3.208, w: 2.822, h: 3.479, rectRadius: 0.2163, fill: { color: C.orange } });
  brandMark(s);
  sideBar(s);
  topBar(s);
  heading(s, [{ text: 'About ' }, { text: 'Atlica’s' }, { text: ' ', options: { breakLine: true } }, { text: 'Charity Event Present' }], { x: 0.822, y: 1.251, w: 5.01, h: 1.178 });
  eyebrow(s, 0.856, 2.486);
  subhead(s, 'About Charity Event Description ', { x: 6.572, y: 1.268, w: 6.536, h: 0.415 });
  body(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled. printing a.",
    { x: 6.572, y: 1.754, w: 5.905, h: 0.903 });
  s.addText('June, 29 2022', { x: 3.948, y: 3.751, w: 2.324, h: 0.55, fontFace: MINOR, fontSize: 20, bold: true, color: C.white, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top' });
  body(s, 'Lorem Ipsum is simply dummy text of the printing and types et ting industry. Lorem Ipsum', { x: 3.948, y: 4.442, w: 2.324, h: 0.903, color: C.white });
  body(s, 'Lorem Ipsum is simply dummy text of the printing and types', { x: 3.948, y: 5.433, w: 2.324, h: 0.625, color: C.white });
}

/* --------------------------------------------------------------- slide 24 */

function slide24(s) {
  cornerRect(s, 9.625, 0, 3.708, 4.688, 0.32, ['bl'], C.orange);
  brandMark(s);
  sideBar(s);
  heading(s, 'BREAK SLIDE PRESENTATION', { x: 1.197, y: 1.661, w: 5.201, h: 1.582, fontSize: 44 });
  eyebrow(s, 1.221, 3.304, { fontSize: 14, w: 3.417, dotDx: 3.471, dotSize: 0.117 });
  s.addText('15 Minutes ', { x: 10.726, y: 0.725, w: 2.276, h: 0.55, fontFace: MINOR, fontSize: 20, bold: true, color: C.white, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top' });
  s.addShape('line', { x: 1.294, y: 4.548, w: 1.387, h: 0, line: { color: C.orange, width: 2.25 } });
  subhead(s, 'About Charity Event Description ', { x: 1.197, y: 4.622, w: 3.417, h: 0.415 });
  body(s, 'Lorem Ipsum is simply dummy text of the printing and typesetting printer took a galley of type and scrambled. printing and typesetting industry',
    { x: 1.197, y: 5.019, w: 5.201, h: 0.625, align: 'left' });
}

/* --------------------------------------------------------------- slide 25 */

function slide25(s) {
  s.addShape('rect', { x: 7.667, y: 0, w: 5.667, h: 7.5, fill: { color: C.orange } });
  // tablet photo stand-in: bezel over the page, with a stylus alongside
  deviceFrame(s, 5.708, 1.458, 3.611, 5.0, 0.22, 11);
  s.addShape('roundRect', { x: 5.597, y: 1.556, w: 0.139, h: 4.777, rectRadius: 0.07, fill: { color: 'EEEFF1' } });
  brandMark(s);
  sideBar(s);
  heading(s, [{ text: 'About ' }, { text: 'Atlica’s' }, { text: ' Tablet Mockup ' }], { x: 0.604, y: 1.625, w: 4.479, h: 1.178 });
  eyebrow(s, 0.638, 2.86);
  body(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled. printing",
    { x: 0.604, y: 3.38, w: 4.479, h: 1.181 });
  body(s, [
    { text: '519M ', options: { bold: true, fontSize: 14 } },
    { text: "Ipsum is simply dummy text of the printing and types etting industry. Lorem Ipsum has been the industry's" },
  ], { x: 0.604, y: 4.743, w: 4.416, h: 0.701 });
  button(s, 'Mockup Information', 0.682, 5.787, 1.965, 0.416, { radius: 0.0422 });
  [['Simple Description 01', 1.508, 1.957], ['Simple Description 02', 3.365, 3.813], ['Simple Description 03', 5.136, 5.584]]
    .forEach(([t, ty, by]) => {
      subhead(s, t, { x: 9.777, y: ty, w: 2.821, h: 0.415, color: C.white });
      body(s, L.standard, { x: 9.777, y: by, w: 2.821, h: 0.903, color: C.white });
    });
}

/* --------------------------------------------------------------- slide 26 */

function slide26(s) {
  s.addShape('rect', { x: 7.033, y: 0, w: 4.946, h: 7.5, fill: { color: C.orange } });
  // desktop-computer photo stand-in: screen bezel, chin, neck and base
  deviceFrame(s, 7.722, 1.375, 5.153, 3.083, 0.02, 13);
  s.addShape('rect', { x: 7.722, y: 4.458, w: 5.167, h: 0.486, fill: { color: 'DFE0E2' } });
  s.addShape('trapezoid', { x: 9.78, y: 4.944, w: 1.05, h: 0.473, flipV: true, fill: { color: 'D2D3D6' } });
  s.addShape('roundRect', { x: 9.528, y: 5.472, w: 1.555, h: 0.181, rectRadius: 0.09, fill: { color: 'C6C8CB' } });
  s.addShape('rect', { x: 0.685, y: 5.419, w: 3.778, h: 0.197, fill: { color: C.track } });
  brandMark(s);
  sideBar(s);
  s.addShape('rect', { x: 6.085, y: 3.841, w: 2.979, h: 2.617, fill: { color: C.white }, shadow: cardShadow() });
  heading(s, [{ text: 'About ' }, { text: 'Atlica’s' }, { text: ' Computer Mockup' }], { x: 0.604, y: 1.245, w: 4.479, h: 1.178 });
  eyebrow(s, 0.638, 2.48);
  body(s, L.unknownPrinting, { x: 0.604, y: 2.938, w: 5.073, h: 0.903 });
  icon(s, 'star5', 0.673, 4.257, 0.592, 0.592, C.orange);
  icon(s, 'cloud', 2.949, 4.189, 0.592, 0.592, C.orange);
  s.addText('1.642+ ', { x: 1.321, y: 4.075, w: 1.389, h: 0.64, fontFace: MINOR, fontSize: 24, bold: true, color: C.slate, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top' });
  s.addText('2 Million', { x: 3.719, y: 4.075, w: 1.803, h: 0.64, fontFace: MINOR, fontSize: 24, bold: true, color: C.slate, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top' });
  body(s, 'Unity Community', { x: 1.321, y: 4.547, w: 1.945, h: 0.348, italic: true });
  body(s, 'Unity Community', { x: 3.733, y: 4.547, w: 1.945, h: 0.348, italic: true });
  s.addShape('rect', { x: 0.685, y: 5.419, w: 2.992, h: 0.197, fill: { color: C.burnt } });
  body(s, '89%', { x: 4.487, y: 5.332, w: 0.537, h: 0.37, fontSize: 12, bold: true, color: C.burnt });
  body(s, L.standardDummy, { x: 0.553, y: 5.709, w: 4.846, h: 0.625 });
  body(s, 'Total Funds ', { x: 6.319, y: 4.236, w: 2.579, h: 0.348, align: 'center', italic: true });
  s.addText('$ 1.283.122', { x: 6.157, y: 4.447, w: 2.812, h: 0.64, align: 'center', fontFace: MINOR, fontSize: 24, bold: true, color: C.orange, lineSpacingMultiple: 1.5, valign: 'top' });
  body(s, 'Lorem Ipsum is simply dummy text of the printing and types etting industry. Lorem', { x: 6.322, y: 5.118, w: 2.579, h: 0.903, align: 'center' });
}

/* --------------------------------------------------------------- slide 27 */

function slide27(s) {
  s.addShape('rect', { x: 7.667, y: 0, w: 5.667, h: 7.5, fill: { color: C.orange } });
  // two-phone photo stand-in: silver back shell behind a bezelled front phone
  s.addShape('roundRect', { x: 6.319, y: 1.417, w: 1.96, h: 6.1, rectRadius: 0.42, fill: { color: 'EDEEF0' } });
  s.addShape('roundRect', { x: 6.46, y: 1.6, w: 0.62, h: 1.28, rectRadius: 0.18, fill: { color: 'F6F7F8' } });
  [1.7, 2.34].forEach((cy) => s.addShape('ellipse', { x: 6.63, y: cy, w: 0.28, h: 0.28, fill: { color: '1E2740' } }));
  s.addShape('roundRect', {
    x: 7.903, y: 1.444, w: 4.708, h: 6.06, rectRadius: 0.42,
    fill: { color: C.orange }, line: { color: '3A3C40', width: 10 },
  });
  s.addShape('roundRect', { x: 9.25, y: 1.5, w: 2.02, h: 0.44, rectRadius: 0.14, fill: { color: '111111' } }); // notch
  brandMark(s);
  sideBar(s);
  topBar(s);
  heading(s, [{ text: 'About ' }, { text: 'Atlica’s' }, { text: ' Phone Mockup' }], { x: 0.706, y: 1.48, w: 4.479, h: 1.178 });
  eyebrow(s, 0.74, 2.715);
  body(s, L.unknownPrinting, { x: 0.706, y: 3.172, w: 5.073, h: 0.903 });
  s.addText('Affects Analysis  :', {
    x: 0.684, y: 4.201, w: 5.073, h: 0.415, fontFace: MINOR, fontSize: 14, bold: true, italic: true, color: C.slate, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top',
  });
  [[0.647, 4.712], [0.625, 5.468]].forEach(([x, y]) => {
    body(s, L.standardDummy, { x, y, w: 5.073, h: 0.625, bullet: { characterCode: '2713', indent: 13.5 } });
  });
}

/* --------------------------------------------------------------- slide 28 */

function slide28(s) {
  s.addShape('roundRect', { x: 2.646, y: 2.812, w: 3.042, h: 3.938, rectRadius: 0.3104, fill: { color: C.orange } });
  topBar(s);
  pageNote(s, 28);
  s.addText('“', { x: 5.945, y: 1.694, w: 4.479, h: 3.45, fontFace: MAJOR, fontSize: 199, color: C.cream, valign: 'top' });
  body(s, 'Love is not patronizing and charity isn’t about pity, it is about love. Charity and love are the same—with charity you give love, so don’t just give money but reach out your hand instead.',
    { x: 6.346, y: 2.608, w: 5.905, h: 1.868, fontSize: 18, bold: true, italic: true, color: C.slate });
  s.addText('-Mother Teresa', { x: 9.301, y: 4.729, w: 2.95, h: 0.438, align: 'right', fontFace: MAJOR, fontSize: 20, color: C.orange, valign: 'top' });
}

/* --------------------------------------------------------------- slide 29 */

function slide29(s) {
  s.addShape('roundRect', { x: 6.667, y: 2.812, w: 4.021, h: 3.938, rectRadius: 0.2826, fill: { color: C.orange } });
  brandMark(s);
  sideBar(s);
  heading(s, 'Check Out Our Profiles Programs and Donate Today!', { x: 0.706, y: 1.397, w: 5.042, h: 1.717 });
  eyebrow(s, 0.74, 3.238);
  const contacts = [
    { label: 'E-mail & Media', lx: 0.735, ly: 3.995, tx: 0.741, ty: 4.457, tw: 2.74, lines: ['Atlica.foundation@domain.com', '@atlica_foundation'] },
    { label: 'Office Hours', lx: 3.399, ly: 4.0, tx: 3.406, ty: 4.461, tw: 1.947, lines: ['Monday – Friday', '09.00 AM – 00.00 PM'] },
    { label: 'Get In Touch', lx: 0.757, ly: 5.364, tx: 0.761, ty: 5.81, tw: 2.245, lines: ['(+11) 1111 4313 51 ', '(0003) 0040 454 1'] },
    { label: 'Address', lx: 3.381, ly: 5.368, tx: 3.377, ty: 5.828, tw: 1.977, lines: ['8779 Windsor St. Fuquay Varina, NC 27526'] },
  ];
  contacts.forEach((c) => {
    s.addText(c.label, { x: c.lx, y: c.ly, w: 1.947, h: 0.46, fontFace: MINOR, fontSize: 16, bold: true, color: C.slate, lineSpacingMultiple: 1.5, valign: 'top' });
    body(s, c.lines.map((t, i) => ({ text: t, options: { breakLine: i < c.lines.length - 1 } })), { x: c.tx, y: c.ty, w: c.tw, h: 0.654 });
  });
}

/* ------------------------------------------------------------------- main */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'ATLICA', width: 13.333, height: 7.5 });
  pptx.layout = 'ATLICA';
  pptx.theme = { headFontFace: MAJOR, bodyFontFace: MINOR };
  pptx.title = 'Atlica — Charity Presentation Template';

  const builders = [
    slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
    slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
  ];
  builders.forEach((fn) => fn(pptx.addSlide()));

  return pptx.writeFile({ fileName: path.join(__dirname, '0f614da5-15fa-44da-8dcd-03d9b798cc52_grok_final.pptx') });
}

build().then((f) => console.log('wrote ' + f)).catch((e) => { console.error(e); process.exit(1); });
