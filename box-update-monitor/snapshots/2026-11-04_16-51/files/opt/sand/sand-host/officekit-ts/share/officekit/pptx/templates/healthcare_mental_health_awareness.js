/**
 * "Mental Health" presentation template - rebuilt with pptxgenjs.
 * Run: node 0afe3f6e-bf89-4710-aeef-29f38b96e7b8_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Design tokens
 * ------------------------------------------------------------------ */
const GREEN = '68D4B2'; // accent / decorative blocks
const CREAM = 'F1FFF8'; // deck background
const INK = '0C1712'; // headings + labels
const WHITE = 'FFFFFF';

const F_TITLE = 'Inter Tight Medium';
const F_LABEL = 'Inter Tight';
const F_BODY = 'Inter Light';

const SLIDE_W = 20.0;
const SLIDE_H = 11.25;

// Text-frame insets used throughout the source deck: 0.1in / 0.05in, in points.
const INSET = [7.2, 7.2, 3.6, 3.6]; // [left, right, bottom, top]

// Soft drop shadow shared by every floating card.
const CARD_SHADOW = { type: 'outer', color: '595959', blur: 60, offset: 37.35, angle: 135, opacity: 0.1 };

/* ------------------------------------------------------------------ *
 * Repeated body copy (the template ships with lorem-ipsum filler)
 * ------------------------------------------------------------------ */
const LOREM = {
  card: 'Duis diam nullam tincidunt ornare tincidunt. Sapiena tristique mattis taciti placerat curabit ipsum.',
  twoLine:
    'Duis diam nullam tincidunt ornare tincidunt. Sapiena tristique mattis tacit Duis diam nullam disan ' +
    'tincidun ornare tincidunt sapien tristique mattis taciti.',
  short: 'Duis diam nullam tincidunt ornare tincidu Sapiena tristique mattis tacit diam nullam disan tincidun ornare tincidunt here.',
  numbered:
    'PLACEHOLDER' +
    'metus dolor ridicu Quam litora dui id auctor adipiscing vestibulum proin rhoncus.',
  data:
    'Duis diam nullam dis tincidunt ornare tincidunt. Sapien tristique mattis taciti placerat curabitur ipsum ' +
    'metus dolor ridiculus. Quam litora dui id auctor adipiscing vestibulum proin rhoncus netus. Aliquam ' +
    'fermentum ornare dapibus ornare data convalis egestas blandit; adipiscing sollicitudin laoreet egestas sodales.',
  tail:
    'Duis diam nullam dis tincidunt ornare tincidunt. Sapien tristique mattis taciti placerat curabitur ipsum ' +
    'metus dolor ridiculus. Quam litora dui id auctor adipiscing vestibulum proin rhoncus netus. ',
  priority:
    'PLACEHOLDER' +
    'dolor ridiculus. Quam litora dui id auctor adipiscing.',
  impact:
    'Duis diam nullam dis tincidunt ornare tincidunt. Sapien tristique mattis taciti placerat curabitur ipsum ' +
    'metus dolor ridiculus. Quam litora dui id auctor adipiscing vestibulum proin rhoncus netus. Aliquam fermentum ornare.',
};

const WRITE_ANYTHING = 'Write Anything Here';
const TOWARDS = 'Towards Mental Health Awareness';

/* ------------------------------------------------------------------ *
 * Building blocks
 * ------------------------------------------------------------------ */

/** Heading text (Inter Tight Medium, dark ink). */
function heading(s, x, y, w, h, text, fontSize) {
  s.addText(text, {
    x, y, w, h, margin: INSET, fontFace: F_TITLE, fontSize: fontSize || 66, color: INK, valign: 'top',
  });
}

/** Body copy (Inter Light, 20pt, 150% leading). */
function body(s, x, y, w, h, text) {
  s.addText(text, {
    x, y, w, h, margin: INSET, fontFace: F_BODY, fontSize: 20, color: '000000',
    lineSpacingMultiple: 1.5, valign: 'top',
  });
}

/** Small caption / kicker (Inter Tight Medium, 24pt by default). */
function label(s, x, y, w, h, text, opts) {
  s.addText(text, Object.assign({
    x, y, w, h, margin: INSET, fontFace: F_TITLE, fontSize: 24, color: INK, valign: 'top',
  }, opts || {}));
}

/** New slide on the cream deck background. */
function newSlide(pres) {
  const s = pres.addSlide();
  s.background = { color: CREAM };
  return s;
}

/** Full-bleed decorative green block coming from the slide layout. */
function greenBlock(s, x, y, w, h) {
  s.addShape('rect', { x, y, w, h, fill: { color: GREEN }, line: { type: 'none' } });
}

/** Numbered circle badge ("01", "02", ...) - oval plus centred number. */
function badge(s, x, y, text) {
  s.addShape('ellipse', { x, y, w: 0.861, h: 0.861, fill: { color: GREEN }, line: { type: 'none' } });
  s.addText(text, {
    x: x + 0.121, y: y + 0.178, w: 0.621, h: 0.505, margin: INSET, wrap: false,
    fontFace: F_TITLE, fontSize: 24, color: INK, align: 'center', valign: 'top',
  });
}

/** Fully rounded green "Discover Now" button. */
function pill(s, x, y, text, opts) {
  const o = opts || {};
  const w = 3.047;
  s.addShape('roundRect', {
    x, y, w, h: 0.988, rectRadius: 0.494, fill: { color: GREEN }, line: { type: 'none' },
  });
  s.addText(text, {
    x, y: y + (o.textDy || 0.242), w, h: o.textH || 0.505, margin: INSET,
    fontFace: F_TITLE, fontSize: o.fontSize || 24, color: INK, align: 'center', valign: 'top',
  });
}

/** Green slab with left-aligned label ("01. Write Anything Here"). */
function slab(s, x, y, text, radius) {
  s.addShape('roundRect', {
    x, y, w: 5.515, h: 1.287, rectRadius: radius === undefined ? 0.128 : radius,
    fill: { color: GREEN }, line: { type: 'none' },
  });
  label(s, x + 0.553, y + 0.391, 4.408, 0.505, text);
}

/** White card with a green outline and a soft shadow. */
function card(s, x, y, w, h, opts) {
  const o = opts || {};
  s.addShape('roundRect', {
    x, y, w, h, rectRadius: Math.min(w, h) * 0.09944,
    fill: { color: o.fill || WHITE },
    line: o.fill ? { type: 'none' } : { color: GREEN, width: 2.25 },
    shadow: CARD_SHADOW,
  });
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

/**
 * Slides 1 & 20 share a big arrow-shaped green graphic on the right edge.
 * The source freeform is mirrored horizontally, hence `flipH`.
 */
function arrowGraphic(s) {
  s.addShape('custGeom', {
    x: 12.224, y: 0, w: 7.776, h: 11.25, flipH: true,
    fill: { color: GREEN }, line: { type: 'none' },
    points: [
      { x: 0.0, y: 0.0, moveTo: true },
      { x: 0.0, y: 3.4735 },
      { x: 0.0, y: 11.25 },
      { x: 7.7765, y: 11.25 },
      { x: 1.0087, y: 4.4823 },
      { x: 4.4823, y: 4.4823 },
      { close: true },
    ],
  });
}

/** 1 - Cover: "Mental / Health". */
function slide01(pres) {
  const s = newSlide(pres);
  arrowGraphic(s);
  s.addText('Mental', { x: 2.188, y: 2.715, w: 10.532, h: 3.45, margin: INSET, fontFace: F_TITLE, fontSize: 199, color: INK, valign: 'top' });
  s.addText('Health', { x: 5.408, y: 4.908, w: 9.895, h: 3.45, margin: INSET, fontFace: F_TITLE, fontSize: 199, color: GREEN, valign: 'top' });
  s.addText('Presentation Template', {
    x: 4.555, y: 2.763, w: 3.566, h: 0.505, margin: INSET, wrap: false,
    fontFace: F_LABEL, fontSize: 24, color: INK, valign: 'top',
  });
}

/** 2 - Intro with call-to-action button. */
function slide02(pres) {
  const s = newSlide(pres);
  heading(s, 2.009, 1.939, 8.527, 2.322, TOWARDS);
  body(s, 2.009, 4.811, 8.527, 3.071,
    'Duis diam nullam dis tincidunt ornare tincidunt. Sapien tristique mattis taciti placerat curabitur ipsum ' +
    'metus dolor ridiculus. Quam litora dui id auctor adipiscing vestibulum proin rhoncus netus. Aliquam ' +
    'fermentum ornare dapibus ornare convallis egestas blandit; adipiscing nisi. Sollicitudin laoreet egestas ' +
    'sodales habitasse amet metus condimen lobortis.');
  pill(s, 2.009, 8.322, 'Discover Now');
}

/** Two stacked outline cards on the right - used by slides 3 and 9. */
function rightCardStack(s) {
  [
    { y: 1.833, x: 12.344, badgeX: 11.915, badgeY: 3.169, n: '01' },
    { y: 5.884, x: 12.348, badgeX: 11.918, badgeY: 7.220, n: '02' },
  ].forEach(c => {
    card(s, c.x, c.y, 6.253, 3.533);
    label(s, c.x + 0.949, c.y + 0.692, 4.642, 0.505, WRITE_ANYTHING);
    body(s, c.x + 0.949, c.y + 1.284, 4.8, 1.557, LOREM.card);
    badge(s, c.badgeX, c.badgeY, c.n);
  });
}

/** 3 - "Why Mental Health Matters". */
function slide03(pres) {
  const s = newSlide(pres);
  greenBlock(s, 13.14, 0, 6.86, 11.25);
  heading(s, 1.749, 2.721, 7.357, 2.322, 'Why Mental Health Matters');
  body(s, 1.749, 5.458, 6.574, 3.071,
    'PLACEHOLDER' +
    'metus dolor ridiculus. Quam litora dui id auctor adipiscing vestibulum proin oncus netus Aliquam ' +
    'fermentum ornare dapibus ornare the convallis egestas blandit adipiscing.');
  rightCardStack(s);
}

/** 4 - "Research Problem": two numbered paragraphs. */
function slide04(pres) {
  const s = newSlide(pres);
  greenBlock(s, 19.028, 0, 0.972, 11.25);
  heading(s, 2.081, 2.038, 9.938, 1.212, 'Research Problem');
  [
    { y: 4.072, badgeY: 4.147, n: '01' },
    { y: 7.063, badgeY: 7.137, n: '02' },
  ].forEach(r => {
    label(s, 3.637, r.y, 7.547, 0.505, WRITE_ANYTHING);
    body(s, 3.637, r.y + 0.593, 7.547, 1.557, LOREM.twoLine);
    badge(s, 2.081, r.badgeY, r.n);
  });
}

/** 5 - "What We Aim to Achieve": five green slabs. */
function slide05(pres) {
  const s = newSlide(pres);
  ['01', '02', '03', '04', '05'].forEach((n, i) => {
    slab(s, 9.66, 1.743 + i * 1.619, n + '. ' + WRITE_ANYTHING);
  });
  heading(s, 2.009, 2.696, 6.649, 2.322, 'What We Aim to Achieve');
  body(s, 2.009, 5.483, 6.649, 3.071,
    'PLACEHOLDER' +
    'metus dolor ridiculus. Quam litora dui id auctor adipiscing vestibulum rhoncus netusan aliquam fermentum ' +
    'ornare dapibus ornare gun convallis egestas blandit adipiscing');
}

/** 6 - "Scope of Study" (right-hand column of copy). */
function slide06(pres) {
  const s = newSlide(pres);
  heading(s, 8.42, 1.816, 9.603, 1.212, 'Scope of Study');
  body(s, 8.42, 3.502, 9.603, 2.566,
    'Duis diam nullam dis tincidunt ornare tincidunt. Sapien tristique mattis taciti placerat curabitur ipsum ' +
    'metus dolor ridiculus. Quam litora dui id auctor adipiscing vestibulum proin rhoncus netus. Aliquam ' +
    'fermentum ornare dapibus ornare convallis egestas blandit; adipiscing sollicitudin laoreet egestas sodales ' +
    'habitasse amet metus condimen');
  body(s, 8.42, 6.344, 9.603, 1.557, LOREM.tail);
  pill(s, 8.42, 8.446, 'Discover Now');
}

/** 7 - "The Science of Mental Health". */
function slide07(pres) {
  const s = newSlide(pres);
  heading(s, 2.009, 2.252, 8.918, 2.322, 'The Science of Mental Health');
  body(s, 2.009, 5.009, 6.649, 2.566,
    'PLACEHOLDER' +
    'metus dolor ridiculus. Quam litora dui id auctor adipiscing vestibulum rhoncus netusan aliquam fermentum ' +
    'ornare dapibus.');
  pill(s, 2.009, 8.009, 'Discover Now', { textDy: 0.276, textH: 0.438, fontSize: 20 });
}

/** 8 - "How We Conducted Our Research": 2x2 numbered grid. */
function slide08(pres) {
  const s = newSlide(pres);
  heading(s, 2.009, 1.432, 10.57, 2.524, 'How We Conducted Our Research', 72);
  [
    { badgeX: 2.009, textX: 3.354, y: 4.805, badgeY: 4.879, n: '01' },
    { badgeX: 2.009, textX: 3.354, y: 7.669, badgeY: 7.743, n: '02' },
    { badgeX: 11.043, textX: 12.388, y: 4.805, badgeY: 4.879, n: '03' },
    { badgeX: 11.043, textX: 12.388, y: 7.669, badgeY: 7.743, n: '04' },
  ].forEach(c => {
    label(s, c.textX, c.y, 6.119, 0.505, WRITE_ANYTHING);
    body(s, c.textX, c.y + 0.593, 6.119, 1.557, LOREM.short);
    badge(s, c.badgeX, c.badgeY, c.n);
  });
}

/** 9 - "Data Analysis" on a green page with a cream panel. */
function slide09(pres) {
  const s = newSlide(pres);
  s.background = { color: GREEN };
  // Cream panel: only its right-hand corners are rounded, so the rect is
  // extended past the left slide edge and the left corners fall out of view.
  s.addShape('roundRect', {
    x: -1.0, y: 0.961, w: 16.605, h: 9.327, rectRadius: 0.691,
    fill: { color: CREAM }, line: { type: 'none' },
  });
  heading(s, 1.851, 1.995, 8.291, 1.212, 'Data Analysis');
  body(s, 1.851, 3.625, 8.666, 2.566, LOREM.data);
  [
    '01. ' + TOWARDS,
    '02. The Science of Mental Health',
    '03. Mental health physical health',
  ].forEach((t, i) => label(s, 1.851, 6.796 + i * 0.9095, 9.301, 0.64, t, { fontSize: 32 }));
  rightCardStack(s);
}

/** 10 - "Current State of Mental Health". */
function slide10(pres) {
  const s = newSlide(pres);
  greenBlock(s, 14.842, 0, 5.158, 11.25);
  heading(s, 2.009, 1.555, 8.918, 2.322, 'Current State of Mental Health');
  [
    { y: 4.555, badgeY: 4.630, n: '01' },
    { y: 7.546, badgeY: 7.620, n: '02' },
  ].forEach(r => {
    label(s, 3.565, r.y, 7.547, 0.505, WRITE_ANYTHING);
    body(s, 3.565, r.y + 0.593, 7.547, 1.557, LOREM.twoLine);
    badge(s, 2.009, r.badgeY, r.n);
  });
}

/** 11 - "Key Findings": intro paragraph over two wide cards. */
function slide11(pres) {
  const s = newSlide(pres);
  greenBlock(s, 0, 7.599, 20, 3.651);
  heading(s, 1.91, 1.465, 12.033, 1.212, 'Key Findings');
  body(s, 1.91, 3.139, 16.179, 2.061,
    'Duis diam nullam dis tincidunt ornare tincidunt. Sapien tristique mattis taciti placerat curabitur ipsum ' +
    'metus dolor ridiculus. Quam litora dui id auctor adipiscing vestibulum proin rhoncus netus. Aliquam ' +
    'fermentum ornare dapibus ornare convallis egestas blandit; adipiscing sollicitudin laoreet egestas sodales ' +
    'habitasse amet metus condiment. Sapien tristique mattis taciti placerat curabitur ipsum metus dolor ' +
    'ridiculus. Quam litora dui id auctor adipiscing vestibulum proin.');
  [
    { x: 1.91, badgeX: 7.939, n: '01' },
    { x: 10.175, badgeX: 16.204, n: '02' },
  ].forEach(c => {
    card(s, c.x, 6.45, 7.915, 3.533);
    label(s, c.x + 0.949, 7.13, 6.866, 0.505, WRITE_ANYTHING);
    body(s, c.x + 0.949, 7.747, 6.366, 1.557,
      'PLACEHOLDER' +
      'metus dolor ridiculus suam litora.');
    badge(s, c.badgeX, 5.962, c.n);
  });
}

/** Three stacked "NN. Towards ..." blocks - slides 12 and 18. */
function numberedList(s, x, w) {
  [0, 1, 2].forEach(i => {
    const y = 1.486 + i * 2.9885;
    label(s, x, y, w, 0.507, '0' + (i + 1) + '. ' + TOWARDS);
    body(s, x, y + 0.744, w, 1.557, LOREM.numbered);
  });
}

/** 12 - "Unpacking the Insights". */
function slide12(pres) {
  const s = newSlide(pres);
  numberedList(s, 9.657, 8.666);
  heading(s, 1.78, 2.31, 6.044, 2.322, 'Unpacking the Insights');
  body(s, 1.78, 4.922, 6.313, 2.566,
    'PLACEHOLDER' +
    'dolor ridiculus. Quam litora du auctor adipiscing vestibulum proin rhoncus netus aliquam fermentu ornare.');
  pill(s, 1.78, 7.952, 'Discover Now', { textDy: 0.241 });
}

/** 13 - "Mental health is as important as physical health". */
function slide13(pres) {
  const s = newSlide(pres);
  heading(s, 1.985, 1.766, 12.454, 2.322, 'Mental health is as important as physical health');
  body(s, 1.985, 4.667, 8.666, 2.566, LOREM.data);
  slab(s, 1.985, 7.888, '01. ' + WRITE_ANYTHING, 0.256);
  slab(s, 7.893, 7.888, '02. ' + WRITE_ANYTHING, 0.256);
}

/** 14 - "Investing in your mental health" with a big stat panel. */
function slide14(pres) {
  const s = newSlide(pres);
  heading(s, 1.78, 5.625, 8.015, 2.322, 'Investing in your mental health');
  body(s, 1.78, 8.242, 7.784, 1.557,
    'Duis diam nullam dis tincidunt ornare tincidunt. Sapienus tristique mattis taciti placerat curabitur ipsum ' +
    'metu dolor ridiculus. Quam litora dui id auctor adipiscing.');
  // Mirrored so the single rounded corner lands on the top-left.
  s.addShape('round1Rect', {
    x: 10.923, y: 2.923, w: 9.077, h: 8.327, flipH: true,
    fill: { color: GREEN }, line: { type: 'none' },
  });
  heading(s, 12.391, 4.378, 4.378, 1.212, '5489+');
  label(s, 12.391, 5.753, 5.404, 1.178, TOWARDS, { fontSize: 32 });
  body(s, 12.391, 7.229, 6.917, 2.566,
    'PLACEHOLDER' +
    'metus dolor ridiculus. Quam litora dui id auctor adipiscing vestibulum proin oncus netu aliquam fermentum ' +
    'ornare dapibus.');
}

/** 15 - "Making Mental Health a Priority" with an 85% circle. */
function slide15(pres) {
  const s = newSlide(pres);
  heading(s, 1.985, 1.706, 8.015, 2.322, 'Making Mental Health a Priority');
  [
    { y: 4.59, w: 7.727, n: '01' },
    { y: 7.347, w: 7.482, n: '02' },
  ].forEach(r => {
    label(s, 1.985, r.y, 6.759, 0.505, r.n + '. ' + WRITE_ANYTHING);
    body(s, 1.985, r.y + 0.641, r.w, 1.557, LOREM.priority);
  });
  s.addShape('ellipse', { x: 10.288, y: 6.509, w: 3.538, h: 3.538, fill: { color: GREEN }, line: { type: 'none' } });
  s.addText('85%', {
    x: 10.288, y: 7.672, w: 3.538, h: 1.212, margin: INSET,
    fontFace: F_TITLE, fontSize: 66, color: INK, align: 'center', valign: 'top',
  });
}

/** 16 - "Impact Summary": two badged paragraphs. */
function slide16(pres) {
  const s = newSlide(pres);
  heading(s, 8.473, 2.326, 9.908, 1.212, 'Impact Summary');
  [
    { y: 4.233, badgeY: 4.463, n: '01' },
    { y: 6.863, badgeY: 7.093, n: '02' },
  ].forEach(r => {
    body(s, 9.81, r.y, 7.929, 2.061, LOREM.impact);
    badge(s, 8.473, r.badgeY, r.n);
  });
}

/** 17 - "Case Study". */
function slide17(pres) {
  const s = newSlide(pres);
  greenBlock(s, 11.19, 4.238, 8.81, 7.012);
  heading(s, 1.675, 2.221, 7.729, 1.212, 'Case Study');
  body(s, 1.675, 3.789, 7.729, 2.566,
    'Duis diam nullam dis tincidunt ornare tincidun. Sapien tristique mattis taciti cerat curabit ipsum metus ' +
    'dolor ridiculus. Quam litora dui id auctor piscing vestibulum proin rhoncus netus. Aliquam ferment ornare ' +
    'dapibus ornare convallis egestas blanda adipiscing.');
  label(s, 1.675, 6.88, 7.547, 0.505, WRITE_ANYTHING);
  body(s, 1.675, 7.473, 7.547, 1.557,
    'Duis diam nullam dis tincidunt ornare tincidun. Sapien tristique mattis taciti cerat curabit ipsum metus ' +
    'dolor ridiculus. Quam litora dui id auctor piscing.');
}

/** 18 - "Limitations": green side panel plus the numbered list. */
function slide18(pres) {
  const s = newSlide(pres);
  greenBlock(s, 0, 0, 8.143, 11.25);
  heading(s, 1.485, 6.237, 5.396, 1.212, 'Limitations');
  body(s, 1.485, 7.699, 5.396, 2.061,
    'Duis diam nullam dis tincidunt ornare tincidun. Sapien tristique mattis taciti ceratus curabit ipsum metus ' +
    'ridiculus Quam litora dui idunas auctor.');
  numberedList(s, 10.038, 8.666);
}

/** 19 - "Future Research Directions" with a filled stat card. */
function slide19(pres) {
  const s = newSlide(pres);
  heading(s, 1.682, 1.914, 8.918, 2.322, 'Future Research Directions');
  body(s, 1.682, 4.734, 8.918, 2.566,
    'Duis diam nullam dis tincidunt ornare tincidunt. Sapien tristique mattis taciti placerat curabitur ipsum ' +
    'metus dolor ridiculus. Quam litora dui id auctor adipiscing vestibulum proin rhoncus netus. Aliquam ' +
    'fermentum ornare dapibus ornare convallis egestas ture blandit; adipiscing sollicitudin laoreet egestas sodales.');
  body(s, 1.682, 7.779, 8.918, 1.557, LOREM.tail);
  label(s, 12.065, 1.948, 5.404, 1.178, TOWARDS, { fontSize: 32 });
  body(s, 12.065, 3.297, 6.626, 2.061,
    'Duis diam nullam dis tincidunt ornare tincidunt. Sapien tristique mattis taciti placerat curabitur ipsum ' +
    'metus dolor ridiculus. Quam litora dui id auctor adipiscing vestibulum rhoncus');
  card(s, 12.065, 6.005, 6.253, 3.533, { fill: GREEN });
  label(s, 13.014, 6.572, 4.642, 0.841, '6457+', { fontSize: 44 });
  body(s, 13.014, 7.414, 4.8, 1.557, LOREM.card);
}

/** 20 - Closing: "Thanks / Attention". */
function slide20(pres) {
  const s = newSlide(pres);
  arrowGraphic(s);
  s.addText('Thanks', { x: 2.009, y: 2.675, w: 11.247, h: 3.45, margin: INSET, fontFace: F_TITLE, fontSize: 199, color: INK, valign: 'top' });
  s.addText('Attention', { x: 3.643, y: 4.77, w: 12.868, h: 3.45, margin: INSET, fontFace: F_TITLE, fontSize: 199, color: GREEN, valign: 'top' });
  s.addText('Presentation Template', {
    x: 4.664, y: 2.965, w: 4.057, h: 0.505, margin: INSET,
    fontFace: F_LABEL, fontSize: 24, color: INK, valign: 'top',
  });
  s.addShape('triangle', {
    x: 12.545, y: 6.576, w: 0.349, h: 0.301, rotate: 90, fill: { color: GREEN }, line: { type: 'none' },
  });
}

/* ------------------------------------------------------------------ *
 * Assemble & save
 * ------------------------------------------------------------------ */
function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'DECK', width: SLIDE_W, height: SLIDE_H });
  pres.layout = 'DECK';
  pres.defineSlideMaster({ title: 'BASE', background: { color: CREAM } });
  pres.title = 'Mental Health';

  [
    slide01, slide02, slide03, slide04, slide05,
    slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15,
    slide16, slide17, slide18, slide19, slide20,
  ].forEach(fn => fn(pres));

  return pres.writeFile({
    fileName: path.join(__dirname, '0afe3f6e-bf89-4710-aeef-29f38b96e7b8_grok_final.pptx'),
  });
}

build().then(f => console.log('wrote', f)).catch(err => {
  console.error(err);
  process.exit(1);
});
