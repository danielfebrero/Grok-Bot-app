/**
 * "Viol — Vase & Ceramics" template deck, rebuilt with pptxgenjs.
 *
 * Widescreen 13.333 x 7.5in. Raster/vector artwork from the original file is
 * replaced by labelled placeholder shapes of the same footprint.
 *
 *   node 1a31b2d9-304d-4202-b5b1-bc6e9b58c26d_grok_final.js
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const C = {
  blue: '6C7EAE', // accent2 - headings, rules, logo
  muted: 'A7B2CE', // accent2 @ 60% lum - body copy
  hair: 'E1E5EE', // accent1 - hairlines
  card: 'F2F2F4', // accent3 - card fill
  slot: 'F4F6FA', // image-placeholder fill
  gray: '595959',
  navy: '44546A',
  white: 'FFFFFF',
};

const FONT = { head: 'Quicksand', body: 'Lato', light: 'Lato Light' };

/* Text presets. Every block in the deck is one of these six. */
const TITLE = { fontFace: FONT.head, fontSize: 44, bold: true, color: C.blue, align: 'left', valign: 'top', margin: 0, lineSpacingMultiple: 0.9 };
const HERO = { ...TITLE, fontSize: 72 };
const LEAD = { fontFace: FONT.light, fontSize: 10, italic: true, color: C.muted, align: 'left', valign: 'top', margin: 0, lineSpacingMultiple: 1.5 };
const LABEL = { fontFace: FONT.body, fontSize: 16, bold: true, color: C.blue, align: 'left', valign: 'top', margin: 0, lineSpacingMultiple: 1.5 };
const SMALL = { fontFace: FONT.body, fontSize: 10, bold: true, color: C.blue, align: 'left', valign: 'top', margin: 0, lineSpacingMultiple: 1.5 };
const NUMBER = { fontFace: FONT.head, fontSize: 28, bold: true, color: C.blue, align: 'left', valign: 'top', margin: 0, lineSpacingMultiple: 0.9 };

/* Filler copy, verbatim from the template. */
const LOREM = {
  vest: 'Vestibulum eget lacus orci. In in quam bibendum, mattis ex aliquet, porttitor lorem. Maecenas eget erat ac risus porttitor viverra. Etiam quis purus at orci. ',
  short: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse tincidunt.',
  welcome:
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse varius tincidunt ligula id posuere. Vestibulum tempor fermentum urna, eget commodo est sollicitudin eu. Nulla sagittis dignissim turpis, interdum scelerisque arcu pretium in. Fusce quam nibh, cursus eget pellentesque eu, vestibulum vitae neque.',
  phasellus:
    'Phasellus dapibus quam quis eros ultrices sodales. Vivamus quis ex eget lectus elementum placerat. Nullam accumsan commodo nibh, mattis porta felis volutpat sit amet. Vestibulum ante ipsum primis in faucibus orci luctus.',
  mission1:
    'Vestibulum eget lacus orci. In in quam bibendum, mattis ex aliquet, porttitor lorem. Maecenas eget erat ac risus porttitor viverra. Etiam quis purus at orci molestie vehicula. Pellentesque non elementum est, sed pharetra erat. Sed bibendum sit amet nibh quis facilisis. Aenean tempus nibh sed tellus convallis dictum id cursus enim.',
  luctus: 'In luctus sed ligula consectetur congue. Donec enim velit, maximus eget euismod porttitor, sagittis quis risus.',
  semper:
    'Sed semper ligula id eros auctor dignissim. Nullam at lacus viverra, fermentum orci eget, aliquam lorem. Nunc erat ante, ullamcorper in urna at, tincidunt sagittis dui.',
  morbi: 'Morbi a ligula urna. Praesent nec commodo elit, sed tincidunt tellus.',
  swot: 'Vestibulum dignissim erat est, sed ultricies ligula congue non. Nulla et est non lorem molestie aliquam et sit amet orci.',
  quote: '\u201C Vestibulum eget lacus orci. In in quam bibendum, mattis ex aliquet, porttitor lorem. Maecenas eget erat ac risus porttitor viverra. Etiam quis purus at orci.  \u201C',
};

/* --------------------------------------------------------------- helpers */

/** Outer rounded rule that frames every slide (7.58% corner radius). */
function frame(s, y, h) {
  s.addShape('roundRect', {
    x: 0.2954, y, w: 12.7426, h,
    line: { color: C.blue, width: 4.5 },
    rectRadius: 0.07584 * Math.min(12.7426, h),
  });
}

/** Right-edge bookmark: a pale 2.31in rule with a solid 0.67in cap. */
function edgeMark(s, y) {
  s.addShape('line', { x: 12.1587, y, w: 0, h: 2.3056, line: { color: C.hair, width: 2 } });
  s.addShape('line', { x: 12.1587, y, w: 0, h: 0.6695, line: { color: C.blue, width: 2 } });
}

/** "Viol / Ceramics" lockup: overlapping squares + circle, then two words. */
function logo(s, x, y) {
  const mark = { line: { color: C.blue, width: 1.5 } };
  s.addShape('roundRect', { x: x + 0.0225, y: y + 0.0349, w: 0.2085, h: 0.2085, rectRadius: 0.045, ...mark });
  s.addShape('roundRect', { x: x + 0.1155, y: y + 0.0928, w: 0.2327, h: 0.2027, rectRadius: 0.045, ...mark });
  s.addShape('ellipse', { x: x, y: y + 0.1701, w: 0.1943, h: 0.1943, ...mark });
  s.addText('Viol', { ...SMALL, x: x + 0.4804, y, w: 0.5913, h: 0.1891 });
  s.addText('Ceramics', { ...SMALL, x: x + 0.4804, y: y + 0.1158, w: 0.5913, h: 0.1891 });
}

/** "Template Presentation" strap-line, flush left or right. */
function footer(s, x, align, y = 6.7801) {
  s.addText('Template Presentation', { ...SMALL, x, y, w: 2.4679, h: 0.2551, align });
}

/** Two-line slide title. `size` is 44 (section) or 72 (cover). */
function heading(s, x, y, w, line1, line2, size = 44) {
  const style = size === 72 ? HERO : TITLE;
  const h = size === 72 ? 0.8865 : 0.6083;
  s.addText(line1, { ...style, x, y, w, h });
  s.addText(line2, { ...style, x, y: y + h, w, h });
}

/** Placeholder standing in for a photo in the original deck. */
function photo(s, x, y, w, h, opts = {}) {
  const box = { x, y, w, h };
  s.addShape(opts.shape || 'rect', {
    ...box,
    fill: { color: opts.fill || C.slot },
    line: { color: C.muted, width: 0.75, dashType: 'dash' },
    ...(opts.rectRadius ? { rectRadius: opts.rectRadius } : {}),
    ...(opts.flipV ? { flipV: true } : {}),
  });
  s.addText('[image]', { ...box, fontFace: FONT.body, fontSize: 9, color: C.muted, align: 'center', valign: 'middle', margin: 0 });
}

/** Bulleted list of paragraphs, matching the deck's 0.25in hanging indent. */
function bulletList(s, items, opts) {
  s.addText(items.map(text => ({ text, options: { bullet: { indent: 18 }, breakLine: true } })),
    { ...LEAD, ...opts, paraSpaceBefore: 10 });
}

/** Grey stat tile: big number over a caption. Tile is 2.418 x 2.424in. */
function statCard(s, x, y, value) {
  s.addShape('roundRect', { x, y, w: 2.418, h: 2.424, fill: { color: C.card }, rectRadius: 0.403 });
  s.addText(value, { ...TITLE, x: x + 0.367, y: y + 0.4683, w: 1.684, h: 0.6083, align: 'center' });
  s.addText(LOREM.short, { ...LEAD, x: x + 0.4461, y: y + 1.141, w: 1.526, h: 0.8151, align: 'center' });
}

/** Outlined capsule with a centred caption (services / SWOT headers). */
function pill(s, x, y, text) {
  s.addShape('roundRect', { x, y, w: 1.532, h: 0.4686, line: { color: C.blue, width: 1 }, rectRadius: 0.2343 });
  s.addText(text, { ...SMALL, x, y: y + 0.0981, w: 1.532, h: 0.2551, align: 'center' });
}

/** One rung of the zig-zag timeline (slides 7 and 8). */
function timelineItem(s, side, tickY, title, body) {
  const right = side === 'right';
  s.addShape('line', { x: right ? 8.0443 : 6.7587, y: tickY, w: 1.2856, h: 0, line: { color: C.hair, width: 1.5 } });
  s.addText(title, { ...LABEL, x: right ? 8.3161 : 5.3038, y: tickY + 0.1442, w: 2.4679, h: 0.3405, align: right ? 'left' : 'right' });
  s.addText(body, { ...LEAD, x: right ? 8.3161 : 4.8557, y: tickY + 0.6262, w: right ? 2.9631 : 2.9159, h: 0.7509, align: right ? 'left' : 'right' });
}

/** Numbered feature paragraph used on the three mockup slides. */
function numberedNote(s, numX, textX, y, num, title, body, numDy = 0.0782) {
  s.addText(num, { ...NUMBER, x: numX, y: y + numDy, w: 0.5733, h: 0.4392 });
  s.addText(title, { ...LABEL, x: textX, y, w: 3.217, h: 0.3405 });
  if (body) s.addText(body, { ...LEAD, x: textX, y: y + 0.4823, w: 3.217, h: 0.7509 });
}

/* ------------------------------------------------------------ the slides */

// 1 — cover
function slide01(p) {
  const s = p.addSlide();
  frame(s, 0.2658, 8.4304);
  heading(s, 1.7765, 1.2274, 8.6413, 'Viol', 'Vase & Ceramics', 72);
  s.addText('A Comprehensive Review of the History, Use, and Beauty of Vases and Ceramics in Modern Culture and Design',
    { ...LEAD, x: 1.7765, y: 3.6483, w: 5.5359, h: 0.6937, fontSize: 14 });
  s.addText([{ text: 'Created By:  ', options: { bold: true } }, { text: 'Pedro Fernandes' }],
    { ...SMALL, bold: false, x: 1.7765, y: 4.685, w: 2.4679, h: 0.2551 });
  s.addText([{ text: 'Speaker:  ', options: { bold: true } }, { text: 'Anna Katrina Marchesi' }],
    { ...SMALL, bold: false, x: 4.6546, y: 4.685, w: 2.4679, h: 0.2551 });
  edgeMark(s, 1.4444);
  logo(s, 10.2638, 4.6352);
  photo(s, 0.952, 5.684, 11.4288, 1.8165, { shape: 'round2SameRect' });
}

// 2 — welcome message
function slide02(p) {
  const s = p.addSlide();
  frame(s, -3.9677, 12.6639);
  heading(s, 1.7765, 1.2274, 3.626, 'Welcome', 'Message');
  edgeMark(s, 1.2317);
  s.addText(LOREM.welcome, { ...LEAD, x: 1.7765, y: 3.1602, w: 5.1121, h: 0.9911 });
  statCard(s, 8.4716, 1.2317, '+1k');
  statCard(s, 8.4716, 3.7836, '+14k');
  logo(s, 1.7765, 4.6915);
  footer(s, 9.6908, 'right');
  photo(s, 1.7765, 5.684, 5.2144, 1.8165, { shape: 'round2SameRect' });
}

// 3 — table of content
function slide03(p) {
  const s = p.addSlide();
  frame(s, -3.9677, 12.6639);
  heading(s, 1.7765, 1.2274, 3.626, 'Table Of', 'Content');
  edgeMark(s, 1.2317);
  const toc = [
    ['01', 'About Our Viol', 2.3867],
    ['02', 'Vision & Mission', 3.042],
    ['03', 'Viol Timeline', 3.6928],
    ['04', 'Viol Services', 4.3436],
    ['05', 'Meet The Team', 4.9899],
  ];
  toc.forEach(([num, text, y]) => {
    s.addText(text, { ...LABEL, x: 6.4897, y, w: 2.4679, h: 0.3586 });
    s.addText(num, { ...LABEL, x: 8.6039, y, w: 2.4679, h: 0.3586, align: 'right' });
    s.addShape('line', { x: 6.4897, y: y + 0.4795, w: 4.5819, h: 0, line: { color: C.hair, width: 0.5 } });
  });
  logo(s, 6.4897, 1.2315);
  footer(s, 9.6908, 'right');
  photo(s, 1.7765, 3.4379, 3.6867, 4.0621, { shape: 'round2SameRect' });
}

// 4 — about our viol
function slide04(p) {
  const s = p.addSlide();
  frame(s, -3.9677, 12.6639);
  heading(s, 5.0923, 1.2274, 3.626, 'About Our', 'Viol');
  edgeMark(s, 1.2317);
  s.addText(LOREM.welcome, { ...LEAD, x: 5.0923, y: 3.0683, w: 4.5981, h: 1.0 });
  s.addText(LOREM.phasellus, { ...LEAD, x: 5.0923, y: 4.3946, w: 4.5981, h: 0.7635 });
  statCard(s, 1.7765, 1.2317, '+1k');
  statCard(s, 1.7765, 3.7836, '+14k');
  logo(s, 11.087, 4.7942);
  footer(s, 9.6908, 'right');
  photo(s, 5.0923, 6.0417, 4.5981, 1.4583, { shape: 'round2SameRect' });
}

// 5 — simple mission
function slide05(p) {
  const s = p.addSlide();
  frame(s, -3.9677, 12.6639);
  heading(s, 1.7765, 1.2274, 3.626, 'Simple', 'Mission');
  edgeMark(s, 1.2317);
  s.addText('Mission 1', { ...LABEL, x: 1.7765, y: 3.2678, w: 2.4679, h: 0.3405 });
  s.addText(LOREM.mission1, { ...LEAD, x: 1.7765, y: 3.75, w: 3.3249, h: 1.4649 });
  s.addText('Mission 2', { ...LABEL, x: 6.1059, y: 3.2678, w: 2.4679, h: 0.3405 });
  bulletList(s, [LOREM.luctus, LOREM.semper, LOREM.morbi], { x: 6.1059, y: 3.75, w: 5.3798, h: 1.4649 });
  logo(s, 1.7765, 6.6707);
  footer(s, 9.6908, 'right');
  photo(s, 6.1059, 0, 5.3798, 2.4441, { shape: 'round2SameRect', flipV: true });
}

// 6 — simple vision
function slide06(p) {
  const s = p.addSlide();
  frame(s, -3.9677, 12.6639);
  heading(s, 0.8969, 1.2274, 3.626, 'Simple', 'Vision');
  edgeMark(s, 1.2317);
  s.addText('Vision 1', { ...LABEL, x: 7.7873, y: 1.2286, w: 2.4679, h: 0.3405 });
  s.addText('Vestibulum eget lacus orci. In in quam bibendum, mattis ex aliquet, porttitor lorem. Maecenas eget erat ac risus porttitor viverra. Etiam quis purus at orci molestie vehicula. Pellen tesque non elementum est, sed pharetra erat. Sed bibendum sit amet nibh quis',
    { ...LEAD, x: 7.7873, y: 1.7107, w: 3.3249, h: 1.4649 });
  s.addText('Vision 2', { ...LABEL, x: 7.7873, y: 3.5567, w: 2.4679, h: 0.3405 });
  bulletList(s, [LOREM.luctus, 'Sed semper ligula id eros auctor dignissim. Nullam at lacus viverra, fermentum orci eget, aliquam lorem. Nunc erat ante, ullamcorper in urna at'],
    { x: 7.7873, y: 4.0388, w: 3.3249, h: 1.4649 });
  s.addText('Vestibulum eget lacus orci. In in quam bibendum, mattis ex aliquet, porttitor lorem. Maecenas eget erat ac risus porttitor viverra.',
    { ...LEAD, x: 0.8969, y: 4.3078, w: 2.0824, h: 1.0546 });
  logo(s, 1.7765, 6.6707);
  footer(s, 9.6908, 'right');
  photo(s, 3.5878, 1.2274, 3.5878, 6.2726, { shape: 'round2SameRect' });
}

// 7 / 8 — timeline, first and second half
function timelineSlide(p, opts) {
  const s = p.addSlide();
  frame(s, -3.9677, 12.6639);
  edgeMark(s, 1.2317);
  if (opts.title) heading(s, 1.7765, 1.2274, 3.626, opts.title[0], opts.title[1]);
  s.addShape('line', { x: 8.0443, y: opts.railY, w: 0, h: 6.2679, line: { color: C.hair, width: 2 } });
  s.addText(LOREM.vest, { ...LEAD, x: 1.7698, y: opts.introY, w: opts.introW, h: 1.4529 });
  opts.items.forEach(([side, tickY, title]) => timelineItem(s, side, tickY, title, LOREM.vest));
  logo(s, 1.7765, 5.6304);
  footer(s, 1.7765, 'left');
  return s;
}

function slide07(p) {
  timelineSlide(p, {
    title: ['Simple', 'Timeline'], railY: 1.2317, introY: 3.1136, introW: 1.8,
    items: [['right', 1.5047, 'First Time'], ['left', 2.4904, 'Second Time'], ['right', 3.8229, 'Third Time'], ['left', 4.8078, 'Fourth Time']],
  });
}

function slide08(p) {
  timelineSlide(p, {
    title: null, railY: 0, introY: 1.1417, introW: 2.4679,
    items: [['right', 1.1417, 'Fifth Time'], ['left', 2.1269, 'Sixth Time'], ['right', 3.4595, 'Seventh Time'], ['left', 4.4445, 'Eighth Time']],
  });
}

/* Slides 9 and 13 share a four-column card strip. */
const COL_X = [1.7765, 4.2617, 6.7469, 9.2321];
const PILL_X = [2.1733, 4.6585, 7.1437, 9.6289];

// 9 — viol services
function slide09(p) {
  const s = p.addSlide();
  frame(s, -3.9677, 12.6639);
  heading(s, 1.7765, 1.2274, 3.626, 'Viol', 'Services');
  edgeMark(s, 1.2317);
  COL_X.forEach(x => s.addShape('roundRect', { x, y: 3.0523, w: 2.3245, h: 3.2202, fill: { color: C.card }, rectRadius: 0.3874 }));
  ['Custom Design', 'Mass Production', 'Art Consulting', 'Global Shipping'].forEach((t, i) => pill(s, PILL_X[i], 5.4062, t));
  PILL_X.forEach(x => photo(s, x, 3.4503, 1.532, 1.532, { shape: 'ellipse' }));
  s.addText(LOREM.phasellus, { ...LEAD, x: 6.7469, y: 1.6008, w: 4.8104, h: 0.7635 });
  footer(s, 1.7765, 'left');
}

// 10 — meet the team
function slide10(p) {
  const s = p.addSlide();
  frame(s, -3.9677, 12.6639);
  heading(s, 1.7765, 1.2274, 3.0011, 'Meet The', 'Team');
  edgeMark(s, 1.2317);
  const team = [['Aaron Loeb', 1.7765], ['Adeline Palmerston', 4.2607], ['Adora Montminy', 6.7454], ['Alexander Aronowitz', 9.2295]];
  team.forEach(([name, x]) => {
    photo(s, x, 2.9793, 2.0499, 2.0499, { shape: 'ellipse' });
    s.addShape('roundRect', { x, y: 5.2812, w: 2.0499, h: 0.4158, fill: { color: C.card }, rectRadius: 0.2079 });
    s.addText(name, { ...SMALL, x, y: 5.3387, w: 2.0499, h: 0.2793, align: 'center' });
  });
  logo(s, 10.2078, 1.2317);
  footer(s, 9.6908, 'right');
}

// 11 — meet the leader
function slide11(p) {
  const s = p.addSlide();
  frame(s, -3.9677, 12.6639);
  photo(s, 4.9086, 1.2327, 4.0499, 4.0499, { shape: 'ellipse' });
  heading(s, 1.7765, 1.2274, 3.0011, 'Meet The', 'Leader');
  edgeMark(s, 1.2317);
  s.addShape('roundRect', { x: 8.0833, y: 3.2757, w: 2.9989, h: 0.6082, fill: { color: C.card }, rectRadius: 0.3041 });
  s.addText('Adeline Palmerston', { ...LABEL, x: 8.3789, y: 3.3521, w: 2.4064, h: 0.3253, align: 'center' });
  s.addText(LOREM.vest, { ...LEAD, x: 8.4707, y: 5.1884, w: 2.9631, h: 0.7509 });
  s.addText(LOREM.vest, { ...LEAD, x: 1.7765, y: 4.8542, w: 2.4536, h: 1.0844 });
  logo(s, 10.2078, 1.2317);
  footer(s, 9.6908, 'right');
}

// 12 — break slide
function slide12(p) {
  const s = p.addSlide();
  frame(s, -3.9677, 12.6639);
  photo(s, 7.03, 1.2124, 2.6603, 6.2726, { shape: 'round2SameRect' });
  heading(s, 1.7765, 1.2274, 4.3187, 'Break', 'Slide', 72);
  s.addText('\u201CCeramics is an art that speaks an eternal language, displaying the infinite beauty of the past to the present. \u201C',
    { ...LEAD, x: 1.7765, y: 3.6483, w: 4.4772, h: 1.3987, fontSize: 18 });
  edgeMark(s, 1.4444);
  statCard(s, 8.8106, 1.2317, '+1k');
  statCard(s, 8.8106, 3.7836, '+14k');
  logo(s, 1.7765, 5.6304);
  footer(s, 1.7765, 'left');
}

// 13 — SWOT analysis
function slide13(p) {
  const s = p.addSlide();
  frame(s, -3.9677, 12.6639);
  heading(s, 1.7765, 1.2274, 3.0011, 'SWOT', 'Analysis');
  edgeMark(s, 1.2317);
  COL_X.forEach(x => s.addShape('roundRect', { x, y: 3.0523, w: 2.3245, h: 3.2202, fill: { color: C.card }, rectRadius: 0.3874 }));
  const swot = [['Strengths', 'S'], ['Weaknesses', 'W'], ['Opportunities', 'O'], ['Threats', 'T']];
  const letterX = [1.7765, 4.2617, 6.7415, 9.2264];
  swot.forEach(([name, letter], i) => {
    pill(s, PILL_X[i], 3.4141, name);
    s.addText(LOREM.swot, { ...LEAD, x: PILL_X[i], y: 4.2447, w: 1.526, h: 1.1976, align: 'center' });
    s.addText(letter, { ...TITLE, x: letterX[i], y: 5.1771, w: 1.9285, h: 1.033, fontSize: 60, color: C.white, lineSpacingMultiple: 1.5 });
  });
  photo(s, 6.6699, 0, 4.8161, 2.4441, { shape: 'round2SameRect', flipV: true });
  footer(s, 9.6908, 'right');
}

// 14 — gallery
function slide14(p) {
  const s = p.addSlide();
  frame(s, -3.9677, 12.6639);
  heading(s, 1.7765, 1.2274, 3.0011, 'Gallery', 'Viol');
  edgeMark(s, 1.2317);
  const tiles = [
    [5.3695, -1.131], [7.8107, -1.131],
    [5.3695, 1.2119], [7.8107, 1.2119],
    [5.3695, 3.5822], [7.8107, 3.5822], [10.2519, 3.5822],
    [7.8107, 5.9528], [10.2519, 5.9528],
  ];
  tiles.forEach(([x, y]) => photo(s, x, y, 2.3411, 2.3167, { shape: 'ellipse' }));
  s.addText(LOREM.vest, { ...LEAD, x: 1.7698, y: 3.5822, w: 2.7202, h: 0.9847 });
  logo(s, 5.3695, 6.6707);
  footer(s, 1.7698, 'left');
}

// 15 — pricing list
function slide15(p) {
  const s = p.addSlide();
  const plans = [
    { name: 'Basic', price: '$99.9', cardX: 4.4317, textX: 4.9227, imgX: 5.0899, btnX: 5.0049, priceX: 4.8336, priceW: 1.5115, sessX: 5.1434, sessW: 1.2017 },
    { name: 'Advanced', price: '$199.9', cardX: 6.9377, textX: 7.4287, imgX: 7.5961, btnX: 7.5115, priceX: 7.2807, priceW: 1.6311, sessX: 7.6144, sessW: 1.2966 },
    { name: 'Premium', price: '$399.9', cardX: 9.4437, textX: 9.9347, imgX: 10.1023, btnX: 10.0181, priceX: 9.7947, priceW: 1.6136, sessX: 10.1263, sessW: 1.2843 },
  ];
  plans.forEach(pl => s.addShape('roundRect', { x: pl.cardX, y: 1.2274, w: 2.3167, h: 4.8628, fill: { color: C.card }, rectRadius: 1.1584 }));
  heading(s, 1.7765, 1.2274, 3.0011, 'Pricing', 'List');
  frame(s, -3.9677, 12.6639);
  edgeMark(s, 1.2317);
  footer(s, 1.7698, 'left');
  plans.forEach(pl => {
    photo(s, pl.imgX, 1.7215, 1.0, 1.0, { shape: 'ellipse' });
    s.addText(pl.name, { ...LABEL, x: pl.textX, y: 3.0995, w: 1.3342, h: 0.2911, fontSize: 18, color: '000000', align: 'center', lineSpacingMultiple: 1.0 });
    s.addText('Ut aliquam nibh vitae ligula pretium. ', { ...LEAD, italic: false, fontFace: FONT.body, color: C.gray, x: pl.textX, y: 3.6062, w: 1.3342, h: 0.6293, align: 'center', lineSpacingMultiple: 1.7 });
    s.addText(pl.price, { ...LABEL, x: pl.priceX, y: 4.4367, w: pl.priceW, h: 0.4059, fontSize: 20, color: C.navy, lineSpacingMultiple: 1.2 });
    s.addText('/ Session', { ...LABEL, x: pl.sessX, y: 4.4367, w: pl.sessW, h: 0.4059, fontSize: 14, color: C.navy, align: 'right', valign: 'middle', lineSpacingMultiple: 1.2 });
    s.addShape('roundRect', { x: pl.btnX, y: 5.0791, w: 1.1691, h: 0.3639, fill: { color: C.white, transparency: 15 }, rectRadius: 0.182 });
    s.addText('Start Now', { ...LABEL, x: pl.textX, y: 4.9978, w: 1.3342, h: 0.4849, fontSize: 12, color: '000000', align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  });
  s.addText(LOREM.vest, { ...LEAD, x: 1.7698, y: 3.6062, w: 1.9884, h: 1.2371 });
  logo(s, 1.7698, 5.7258);
}

// 16 — laptop mockup. Stands in for the laptop photo: bezel, screen, base.
function slide16(p) {
  const s = p.addSlide();
  frame(s, -3.9677, 12.6639);
  heading(s, 2.2896, 1.2274, 3.0011, 'Unique', 'Mockup');
  edgeMark(s, 1.2317);
  s.addShape('rect', { x: 2.3333, y: 3.1667, w: 4.2361, h: 2.8194, fill: { color: '2B2B2B' } });
  s.addShape('rect', { x: 2.4167, y: 3.2361, w: 4.0694, h: 2.6806, fill: { color: C.white } });
  s.addShape('roundRect', { x: 1.7765, y: 5.9861, w: 5.3436, h: 0.2083, fill: { color: 'C9CDD6' }, rectRadius: 0.06 });
  s.addShape('roundRect', { x: 3.9583, y: 5.9861, w: 0.9722, h: 0.0694, fill: { color: '9DA3B0' }, rectRadius: 0.035 });
  photo(s, 2.4503, 3.3283, 4.0208, 2.3);
  numberedNote(s, 7.5629, 8.2126, 2.6779, '01', 'Website', LOREM.vest);
  numberedNote(s, 7.5629, 8.2126, 4.4058, '02', 'Information', LOREM.vest, 0.0436);
  logo(s, 10.3578, 1.2317);
  footer(s, 9.6908, 'right');
}

// 17 — phone mockup. Stands in for the smartphone photo: body, screen, button.
function slide17(p) {
  const s = p.addSlide();
  frame(s, -3.9677, 12.6639);
  s.addShape('roundRect', { x: 6.7361, y: 1.0556, w: 2.8472, h: 5.4167, fill: { color: '1C1C1E' }, rectRadius: 0.42 });
  s.addShape('rect', { x: 6.9028, y: 1.7778, w: 2.5139, h: 3.9167, fill: { color: C.white } });
  s.addShape('ellipse', { x: 8.0694, y: 1.3056, w: 0.1111, h: 0.1111, fill: { color: '3A3A3C' } });
  s.addShape('roundRect', { x: 7.9167, y: 5.8472, w: 0.4722, h: 0.4028, fill: { color: '2C2C2E' }, line: { color: '505055', width: 1 }, rectRadius: 0.1 });
  photo(s, 6.9283, 1.8311, 2.4425, 3.8374);
  heading(s, 2.2896, 1.2274, 3.0011, 'Unique', 'Mockup');
  edgeMark(s, 1.2317);
  numberedNote(s, 2.2938, 2.9428, 3.3121, '01', 'Mobile', LOREM.vest);
  numberedNote(s, 2.2938, 2.9428, 5.0389, '02', 'Information', LOREM.vest, 0.0441);
  logo(s, 10.3578, 1.2317);
  footer(s, 9.6908, 'right');
}

// 18 — map mockup. The original artwork is a 24-region vector map of France;
// the two silhouettes below (mainland + Corsica) are simplified outlines of it,
// given as inch offsets from each shape's own top-left corner.
const MAP_MAINLAND = {
  x: 3.9513, y: 1.1328, w: 5.0042, h: 5.0175,
  pts: [
    [2.799, 0.000], [2.859, 0.007], [2.872, 0.133], [2.965, 0.207], [3.105, 0.193], [3.138, 0.320],
    [3.245, 0.333], [3.298, 0.446], [3.458, 0.460], [3.492, 0.666], [3.678, 0.646], [3.711, 0.560],
    [3.738, 0.713], [3.985, 0.893], [4.318, 0.906], [4.471, 1.073], [4.724, 1.073], [5.004, 1.166],
    [4.864, 1.359], [4.784, 1.653], [4.798, 1.952], [4.744, 2.019], [4.584, 2.026], [4.598, 2.106],
    [4.238, 2.585], [4.211, 2.779], [4.218, 2.825], [4.311, 2.812], [4.358, 2.625], [4.571, 2.599],
    [4.518, 2.639], [4.531, 2.805], [4.631, 2.939], [4.558, 3.052], [4.691, 3.265], [4.604, 3.352],
    [4.511, 3.365], [4.484, 3.432], [4.651, 3.592], [4.591, 3.718], [4.618, 3.865], [4.778, 3.958],
    [4.938, 3.951], [4.891, 4.118], [4.791, 4.151], [4.551, 4.378], [4.564, 4.464], [4.358, 4.538],
    [4.031, 4.464], [4.025, 4.391], [3.898, 4.398], [3.978, 4.371], [3.931, 4.291], [3.878, 4.305],
    [3.885, 4.351], [3.811, 4.344], [3.791, 4.404], [3.438, 4.291], [3.265, 4.444], [3.185, 4.451],
    [3.065, 4.604], [3.032, 4.718], [3.132, 4.958], [2.992, 4.944], [2.885, 5.017], [2.765, 4.964],
    [2.619, 4.991], [2.479, 4.858], [2.139, 4.738], [2.046, 4.738], [2.012, 4.824], [1.759, 4.798],
    [1.166, 4.571], [1.199, 4.471], [1.053, 4.404], [1.186, 4.285], [1.253, 4.051], [1.379, 3.138],
    [1.319, 3.038], [1.399, 2.952], [1.399, 2.832], [1.346, 2.779], [1.379, 2.699], [1.099, 2.559],
    [0.980, 2.365], [1.039, 2.279], [0.973, 2.199], [0.993, 2.119], [0.846, 2.106], [0.886, 2.046],
    [0.853, 1.959], [0.733, 1.972], [0.786, 1.946], [0.753, 1.892], [0.613, 1.926], [0.267, 1.706],
    [0.160, 1.766], [0.153, 1.692], [0.020, 1.626], [0.180, 1.599], [0.187, 1.539], [0.080, 1.526],
    [0.067, 1.439], [0.000, 1.439], [0.013, 1.346], [0.113, 1.293], [0.486, 1.286], [0.506, 1.213],
    [0.606, 1.193], [0.800, 1.386], [0.953, 1.326], [0.986, 1.373], [1.133, 1.326], [1.186, 1.393],
    [1.299, 1.379], [1.273, 1.013], [1.159, 0.740], [1.379, 0.760], [1.426, 0.960], [1.839, 1.019],
    [1.966, 0.940], [1.932, 0.840], [2.012, 0.740], [2.439, 0.593], [2.525, 0.480], [2.539, 0.100],
  ],
};
const MAP_CORSICA = {
  x: 9.1287, y: 5.6172, w: 0.4065, h: 0.9195,
  pts: [
    [0.313, 0.000], [0.346, 0.107], [0.333, 0.233], [0.366, 0.240], [0.406, 0.453], [0.346, 0.600],
    [0.366, 0.706], [0.340, 0.813], [0.280, 0.920], [0.127, 0.833], [0.153, 0.746], [0.080, 0.733],
    [0.113, 0.686], [0.100, 0.633], [0.033, 0.633], [0.073, 0.526], [0.013, 0.500], [0.033, 0.420],
    [0.000, 0.400], [0.087, 0.260], [0.180, 0.227], [0.207, 0.180], [0.307, 0.173],
  ],
};

function mapShape(s, region) {
  s.addShape('custGeom', {
    x: region.x, y: region.y, w: region.w, h: region.h,
    fill: { color: C.blue }, line: { color: C.white, width: 1.5 },
    points: [...region.pts.map(([x, y]) => ({ x, y })), { close: true }],
  });
}

function slide18(p) {
  const s = p.addSlide();
  frame(s, -3.9677, 12.6639);
  heading(s, 1.1908, 1.6618, 3.0012, 'Unique', 'Mockup');
  edgeMark(s, 1.2317);
  mapShape(s, MAP_MAINLAND);
  mapShape(s, MAP_CORSICA);
  numberedNote(s, 1.1964, 1.8461, 4.0234, '01', 'Map Information', '', 0.0432);
  s.addText(LOREM.vest, { ...LEAD, x: 1.8461, y: 4.5058, w: 1.9552, h: 1.3934 });
  s.addText('Vestibulum eget lacus orci. In in quam bibendum, mattis ex aliquet, porttitor lorem. Maecenas eget erat orci. ',
    { ...LEAD, x: 9.4107, y: 2.0203, w: 2.2425, h: 0.9338 });
  logo(s, 10.3578, 1.2317);
  footer(s, 9.6908, 'right');
}

// 19 — client testimonial
function slide19(p) {
  const s = p.addSlide();
  frame(s, -3.9677, 12.6639);
  edgeMark(s, 1.2317);
  heading(s, 1.7765, 1.2274, 3.5251, 'Client', 'Testimonial');
  const cards = [{ cardX: 1.7765, x: 2.2786, nameX: 3.6277, name: 'Eleanor Fitzgerald' },
    { cardX: 6.6667, x: 7.1687, nameX: 8.5178, name: 'Anna Katrina Marchesi' }];
  cards.forEach(c => {
    s.addShape('roundRect', { x: c.cardX, y: 3.1905, w: 4.6667, h: 3.0821, fill: { color: C.card }, rectRadius: 0.5326 });
    photo(s, c.x, 3.6118, 0.9973, 0.9973, { shape: 'ellipse' });
    s.addText(c.name, { ...LABEL, x: c.nameX, y: 3.7986, w: 1.9143, h: 0.6234, fontSize: 18, lineSpacingMultiple: 1.0 });
    s.addText(LOREM.quote, { ...LEAD, x: c.x, y: 4.7247, w: 3.6629, h: 0.7268 });
    [0, 0.2643, 0.5329, 0.8015, 1.0701].forEach((dx, i) => {
      s.addShape('star5', { x: c.x + dx, y: 5.666, w: 0.1857, h: 0.1857, fill: { color: i < 4 ? C.blue : C.white } });
    });
  });
  logo(s, 10.2616, 1.2317);
  footer(s, 9.6908, 'right');
}

// 20 — thank you
function slide20(p) {
  const s = p.addSlide();
  photo(s, 7.9028, 1.2274, 3.588, 6.2726, { shape: 'round2SameRect' });
  frame(s, -1.2148, 8.4304);
  heading(s, 1.7765, 1.2274, 4.3187, 'Thank', 'You', 72);
  edgeMark(s, 1.4444);
  s.addShape('roundRect', { x: 6.4019, y: 3.2116, w: 4.4487, h: 2.7821, fill: { color: C.card }, rectRadius: 0.3911 });
  s.addText('Get In Touch', { ...LABEL, x: 7.0699, y: 3.6402, w: 2.6825, h: 0.4468, fontSize: 18, color: '000000' });
  const contact = [
    ['Phone:', '+123-456-7890', 7.0738, 0.7669, 7.0699, 1.4219],
    ['Website:', 'www.yourwebsite.com', 8.8463, 0.7229, 8.8423, 1.3404],
  ];
  contact.forEach(([label, value, lx, lw, vx, vw]) => {
    const y = 4.2655;
    s.addText(label, { ...SMALL, x: lx, y, w: lw, h: 0.1994, color: '000000', lineSpacingMultiple: 1.0 });
    s.addText(value, { ...LEAD, fontSize: 10, color: C.gray, x: vx, y: y + 0.1459, w: vw, h: 0.2585 });
  });
  s.addText('Address', { ...SMALL, x: 7.0738, y: 4.8173, w: 0.7669, h: 0.1994, color: '000000', lineSpacingMultiple: 1.0 });
  s.addText('+ABC Solutions Inc. 123 Main Street Suite 456 Cityville, State 78901 United States',
    { ...LEAD, color: C.gray, x: 7.0699, y: 4.9632, w: 3.1128, h: 0.5884 });
  s.addText('Vestibulum eget lacus orci. In in quam bibendum, mattis ex aliquet, porttitor lorem. Maecenas eget erat ac risus porttitor viverra. Etiam quis purus at orci molestie vehicula. Pellentesque non elementum est, sed pharetra erat. ',
    { ...LEAD, x: 1.7765, y: 3.75, w: 3.7919, h: 0.9927 });
  logo(s, 1.7698, 5.6292);
  footer(s, 1.7698, 'left', 6.5772);
}

/* ------------------------------------------------------------------ build */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'VIOL_16x9', width: 13.3333, height: 7.5 });
  pptx.layout = 'VIOL_16x9';
  pptx.author = 'Pedro Fernandes';
  pptx.title = 'Viol Vase & Ceramics';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach(fn => fn(pptx));

  return pptx.writeFile({ fileName: path.join(__dirname, '1a31b2d9-304d-4202-b5b1-bc6e9b58c26d_grok_final.pptx') });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
