/**
 * "Cut Barbershop" deck - rebuilt with pptxgenjs.
 * Run: node 04f63833-9bb8-4ef5-8872-297683679a02_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette --
const INK = '232423';      // accent1 - near black
const GOLD = 'DEA862';     // accent2 - brass
const CREAM = 'F8F3F0';    // accent3 - page background
const BODY = '595959';     // tx1 @ 65% - body copy
const SILVER = 'D9D9D9';   // muted captions / map land
const MAP_EDGE = 'F2F2F2'; // borders between map countries
const WHITE = 'FFFFFF';

const HEAD = 'Poppins';    // major latin
const SANS = 'Lato';       // minor latin
const LIGHT = 'Lato Light';

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

// ------------------------------------------------------------- primitives --
function rect(slide, x, y, w, h, color) {
  slide.addShape('rect', { x, y, w, h, fill: { color }, line: { type: 'none' } });
}

function text(slide, str, o) {
  slide.addText(str, Object.assign({
    margin: 0, valign: 'top', isTextBox: true, fontFace: SANS, color: INK,
  }, o));
}

/** 10pt italic Lato Light paragraph copy, 150% leading. */
function body(slide, str, x, y, w, h, o) {
  text(slide, str, Object.assign({
    x, y, w, h, fontSize: 10, italic: true, fontFace: LIGHT, color: BODY,
    lineSpacingMultiple: 1.5,
  }, o));
}

/** Small bold caption (Lato). */
function label(slide, str, x, y, w, h, o) {
  text(slide, str, Object.assign({
    x, y, w, h, fontSize: 10, bold: true, lineSpacingMultiple: 1.0,
  }, o));
}

/** Two stacked 48pt Poppins words - the standard slide title. */
function heading(slide, x, y, line1, line2, o) {
  const opt = Object.assign({ w: 3.694, color: INK }, o);
  [line1, line2].forEach((word, i) => {
    text(slide, word, {
      x, y: y + i * 0.573, w: opt.w, h: 0.512, fontSize: 48, bold: true,
      fontFace: HEAD, color: opt.color, align: opt.align || 'left',
    });
  });
}

/** 80pt Poppins hero title (cover / break / closing slides). */
function hero(slide, x, y, w, line1, line2, align, color) {
  [line1, line2].forEach((word, i) => {
    text(slide, word, {
      x, y: y + i * 0.899, w, h: 0.899, fontSize: 80, bold: true,
      fontFace: HEAD, color: color || INK, align: align || 'center',
    });
  });
}

/** The three-bar brand rule; anchor is its left edge / top edge. */
function ruleMark(slide, x, y, color) {
  rect(slide, x + 0.493, y, 1.0, 0.129, color);
  rect(slide, x + 1.493, y + 0.039, 0.493, 0.05, color);
  rect(slide, x, y + 0.039, 0.493, 0.05, color);
}

/** Bottom-left brand rule + "Est. 1999". */
function footerMark(slide) {
  ruleMark(slide, 0.361, 7.095, INK);
  label(slide, 'Est. 1999', 2.622, 7.068, 0.759, 0.182);
}

/** Top-right brand rule + "Est. 1999". */
function headerMark(slide) {
  ruleMark(slide, 11.056, 0.288, INK);
  label(slide, 'Est. 1999', 10.05, 0.262, 0.759, 0.182, { align: 'right' });
}

const CORNER = { tl: [0, 0], tr: [13.207, 0], bl: [0, 6.665], br: [13.207, 6.665] };
function cornerTabs(slide) {
  for (let i = 0; i < arguments.length - 1; i++) {
    const [x, y] = CORNER[arguments[i + 1]];
    rect(slide, x, y, 0.127, 0.835, INK);
  }
}

/**
 * Stand-in for a picture placeholder: the little framed-photo glyph plus the
 * "Image Here" prompt, centred in the frame the photo would have filled.
 */
function imagePlaceholder(slide, x, y, w, h) {
  const cx = x + w / 2;
  const cy = y + h / 2;
  const ix = cx - 0.423;
  const iy = cy - 0.326;
  slide.addShape('rect', { x: ix, y: iy, w: 0.846, h: 0.652, fill: { color: 'FAFAFA' }, line: { color: '9A9A99', width: 1 } });
  slide.addShape('rect', { x: ix + 0.062, y: iy + 0.05, w: 0.722, h: 0.552, fill: { color: WHITE }, line: { color: 'BFBFBE', width: 0.75 } });
  slide.addShape('ellipse', { x: ix + 0.125, y: iy + 0.097, w: 0.153, h: 0.153, fill: { color: 'F8DB8F' }, line: { color: 'EFA93F', width: 1 } });
  slide.addShape('custGeom', {
    x: ix, y: iy, w: 0.846, h: 0.652, fill: { color: '83BEEC' }, line: { type: 'none' },
    points: [
      { x: 0.084, y: 0.611 }, { x: 0.300, y: 0.402 }, { x: 0.404, y: 0.505 },
      { x: 0.545, y: 0.250 }, { x: 0.764, y: 0.611 }, { close: true },
    ],
  });
  text(slide, 'Image Here', {
    x, y, w, h, fontSize: 18, color: '000000', align: 'center', valign: 'middle',
  });
}

/** Concentric "target" glyph used beside the Goals column. */
function targetIcon(slide, x, y, size) {
  const rings = [[0.360, INK], [0.306, CREAM], [0.222, INK], [0.166, CREAM], [0.070, INK]];
  const cx = x + size / 2;
  const cy = y + size / 2;
  rings.forEach(([d, color]) => {
    const s = (d / 0.36) * size;
    slide.addShape('ellipse', { x: cx - s / 2, y: cy - s / 2, w: s, h: s, fill: { color }, line: { type: 'none' } });
  });
}

/** Eye glyph used beside the Vision / Mission column. */
function eyeIcon(slide, x, y) {
  const w = 0.372;
  const h = 0.216;
  slide.addShape('custGeom', {
    x, y, w, h, fill: { color: INK }, line: { type: 'none' },
    points: [
      { x: 0, y: h / 2 },
      { curve: { type: 'quadratic', x1: w / 2, y1: -h * 0.34, x: w, y: h / 2 } },
      { curve: { type: 'quadratic', x1: w / 2, y1: h * 1.34, x: 0, y: h / 2 } },
      { close: true },
    ],
  });
  slide.addShape('ellipse', { x: x + 0.125, y: y + 0.039, w: 0.139, h: 0.139, fill: { color: CREAM }, line: { type: 'none' } });
  slide.addShape('ellipse', { x: x + 0.166, y: y + 0.080, w: 0.056, h: 0.056, fill: { color: INK }, line: { type: 'none' } });
}

/** "+100 / Cust/day / note" statistic stack. */
function statBlock(slide, x, y, value, caption, note) {
  text(slide, value, { x, y, w: 1.516, h: 0.368, fontSize: 32, bold: true, fontFace: HEAD, color: GOLD });
  label(slide, caption, x, y + 0.516, 0.759, 0.182);
  body(slide, note, x, y + 0.698, 1.516, 0.448);
}

const LOREM_LONG = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Donec a dignissim risus. ' +
  'Mauris pharetra vitae felis sed finibus. Donec mollis arcu sed leo sagittis, eget dignissim justo rutrum. ' +
  'Ut elementum sem dui, a accumsan massa malesuada at. Ut rhoncus id urna molestie fringilla. ' +
  'In eget lacus a augue facilisis consectetur.';
const LOREM_MED = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Donec a dignissim risus. ' +
  'Mauris pharetra vitae felis sed finibus. Donec mollis arcu sed leo sagittis, eget dignissim justo rutrum. ' +
  'Ut elementum sem dui, a accumsan massa malesuada at.';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Donec a dignissim risus. ' +
  'Mauris pharetra vitae felis sed finibus. Donec mollis arcu sed leo sagittis.';
const LOREM_TINY = 'Ut rhoncus id urna molestie fringilla. In eget lacus a.';
const LOREM_NOTE = 'Donec mollis arcu sed leo sagittis, eget dignissim justo rutrum. ' +
  'Ut elementum sem dui, a accumsan massa malesuada at.';
const PRICE_COPY = 'Aenean sagittis inter bibendum. Pellentesque habitant morbi tristique senectus ' +
  'et netus et malesuada fames ac turpis egestas. Donec sed purus elementum.';
const QUOTE = '\u201C In vel nunc vestibulum diam ultricies mattis. Cras massa massa, commodo vestibulum ' +
  'sit amet, efficitur at magna. Nam a dapibus lacus. Vestibulum tincidunt quam ut luctus suscipit. ' +
  'Aenean a dapibus nulla. \u201C';
const MOCK_COPY_A = 'Suspendisse egestas ipsum in mi dapibus tempor. Interdum et malesuada fames ac ante ' +
  'ipsum primis in faucibus. Fusce metus ante.';
const MOCK_COPY_B = 'In id tortor enim. Lorem ipsum dolor si, consectetur adipiscing elit. ' +
  'Mauris vehicula tellus tellus, vel volutpat elit elementum sed.';

// ------------------------------------------------------------ slide 1 -----
function coverSlide(slide) {
  imagePlaceholder(slide, 0, 0, SLIDE_W, SLIDE_H);
  body(slide, 'PLACEHOLDER' +
    'Building a Competitive Community', 4.075, 3.826, 5.184, 0.648, { fontSize: 12, color: CREAM, align: 'center' });
  hero(slide, 3.209, 1.368, 6.916, 'Cut', 'Barbershop', 'center', GOLD);
  text(slide, [
    { text: 'Created By: ', options: { bold: true } },
    { text: ' Rosa Maria Aguado' },
  ], { x: 4.439, y: 5.28, w: 2.115, h: 0.142, fontSize: 10, color: CREAM, lineSpacingMultiple: 1.0 });
  text(slide, [
    { text: 'Speaker: ', options: { bold: true } },
    { text: ' Morgan Maxwell' },
  ], { x: 6.756, y: 5.28, w: 2.138, h: 0.142, fontSize: 10, color: CREAM, align: 'right', lineSpacingMultiple: 1.0 });
  ruleMark(slide, 5.674, 6.043, GOLD);
  label(slide, 'EST', 4.851, 6.016, 0.759, 0.182, { color: CREAM });
  label(slide, '2018', 7.724, 6.016, 0.759, 0.182, { color: CREAM, align: 'right' });
}

// ------------------------------------------------------------ slide 2 -----
function welcomeSlide(slide) {
  heading(slide, 1.653, 1.227, 'Welcome', 'Message');
  body(slide, LOREM_LONG, 1.653, 3.153, 5.486, 0.988);
  footerMark(slide);
  statBlock(slide, 1.653, 4.667, '+100', 'Cust/day', LOREM_TINY);
  statBlock(slide, 3.763, 4.667, '+1.2K', 'Cust/week', LOREM_TINY);
  cornerTabs(slide, 'tl');
  imagePlaceholder(slide, 8.986, 3.153, 4.347, 4.347);
  imagePlaceholder(slide, 10.18, 0, 3.153, 3.153);
}

// ------------------------------------------------------------ slide 3 -----
const TOC_ITEMS = ['About Our Cut', 'Vision & Mission', 'Cut Timeline', 'Cut Services', 'Meet The Team'];

function tableOfContentsSlide(slide) {
  heading(slide, 1.653, 1.227, 'Table Of', 'Content');
  footerMark(slide);
  cornerTabs(slide, 'tl', 'br');
  TOC_ITEMS.forEach((item, i) => {
    const y = 2.286 + i * 0.7203;
    text(slide, item, {
      x: 8.033, y, w: 1.341, h: 0.269, fontSize: 10, bold: true, lineSpacingMultiple: 1.5,
      bullet: { characterCode: '2751', indent: 13.5 },
    });
    text(slide, '0' + (i + 1), {
      x: 10.675, y, w: 1.075, h: 0.269, fontSize: 10, bold: true, fontFace: HEAD,
      align: 'right', lineSpacingMultiple: 1.5,
    });
    slide.addShape('line', {
      x: 8.033, y: 2.783 + i * 0.7360, w: 3.717, h: 0, line: { color: INK, width: 1 },
    });
  });
  imagePlaceholder(slide, 0, 3.134, 4.014, 3.247);
  imagePlaceholder(slide, 4.014, 3.134, 2.653, 4.366);
}

// ------------------------------------------------------------ slide 4 -----
function aboutSlide(slide) {
  rect(slide, 10.553, 4.633, 2.78, 2.867, GOLD);
  heading(slide, 1.653, 1.227, 'About Our', 'Cut');
  footerMark(slide);
  cornerTabs(slide, 'tl', 'br');
  body(slide, LOREM_LONG, 1.653, 2.867, 4.739, 0.988);
  statBlock(slide, 7.57, 2.867, '+100', 'Cust/day', LOREM_TINY);
  statBlock(slide, 7.57, 4.377, '+1.2K', 'Cust/week', LOREM_TINY);
  body(slide, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Donec a dignissim risus. ' +
    'Mauris pharetra vitae felis sed finibus. Donec mollis arcu sed leo sagittis, eget dignissim justo rutrum.',
    10.83, 4.805, 1.995, 1.445, { color: CREAM });
  imagePlaceholder(slide, 0, 4.377, 6.392, 2.205);
  imagePlaceholder(slide, 10.553, 0, 2.78, 4.633);
}

// -------------------------------------------------------- slides 5 & 6 ----
function visionSlide(slide) {
  heading(slide, 5.417, 1.227, 'Our Great', 'Vision');
  footerMark(slide);
  cornerTabs(slide, 'tl', 'br');
  body(slide, LOREM_MED, 5.417, 4.112, 2.823, 1.202);
  label(slide, 'Vision', 5.417, 3.804, 0.759, 0.182, { fontSize: 14 });
  targetIcon(slide, 9.111, 3.193, 0.37);
  eyeIcon(slide, 5.417, 3.27);
  body(slide, [
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit.',
    'Donec a dignissim risus. Mauris pharetra vitae felis sed finibus.',
    'Donec mollis arcu sed leo sagittis, eget dignissim justo rutrum. Ut elementum sem dui, a accumsan massa malesuada at.',
  ].join('\n'), 9.111, 4.112, 2.823, 1.977, { bullet: { characterCode: '2022', indent: 18 } });
  label(slide, 'Goals', 9.111, 3.804, 0.759, 0.182, { fontSize: 14 });
  imagePlaceholder(slide, 0.361, 0, 3.861, 6.089);
}

function missionSlide(slide) {
  heading(slide, 1.691, 1.615, 'Our Great', 'Mission');
  footerMark(slide);
  cornerTabs(slide, 'tl', 'br');
  body(slide, LOREM_MED, 1.691, 4.583, 4.018, 1.092);
  label(slide, 'Mission', 1.691, 4.191, 0.759, 0.182, { fontSize: 14 });
  targetIcon(slide, 7.353, 3.58, 0.37);
  eyeIcon(slide, 1.691, 3.657);
  label(slide, 'Goals', 7.353, 4.191, 0.759, 0.182, { fontSize: 14 });
  body(slide, LOREM_MED, 7.353, 4.583, 4.018, 1.092);
}

// -------------------------------------------------------- slides 7 & 8 ----
const TIMELINE_A = [
  ['1999 (Jan)', 'Conceptualization & Planning'],
  ['1999 (July)', 'Grand Opening'],
  ['2000', 'Partnership Development'],
];
const TIMELINE_B = [
  ['2001', 'Partnership Development'],
  ['2002', 'Diversification of Services'],
  ['2003', 'Technology and Innovation'],
];

function timelineSlide(slide, entries, ruleX) {
  heading(slide, 1.653, 1.227, 'Cut', 'Timeline');
  footerMark(slide);
  cornerTabs(slide, 'tl', 'br');
  slide.addShape('line', { x: ruleX, y: 3.575, w: 11.68, h: 0, line: { color: INK, width: 1 } });
  entries.forEach(([year, title], i) => {
    const x = 1.653 + i * 3.602;
    label(slide, year, x, 3.307, 1.099, 0.12, { fontFace: HEAD });
    rect(slide, x, 3.75, 0.201, 0.201, GOLD);
    label(slide, title, x, 4.137, 2.823, 0.182, { fontSize: 14 });
    body(slide, LOREM_MED, x, 4.444, 2.823, 1.202);
  });
  imagePlaceholder(slide, 6.908, 0, 6.425, 0.835);
  imagePlaceholder(slide, 5.255, 6.665, 6.425, 0.835);
}

// ------------------------------------------------------------ slide 9 -----
const SERVICES = [
  ["Men's Haircut + Face treatment", 1.827, 2.104, 1.517],
  ['Beard Shaving and Facial Grooming', 5.344, 5.52, 1.721],
  ['Hair Care and Style', 8.861, 9.139, 1.517],
];

function servicesSlide(slide) {
  rect(slide, 0, 0.417, SLIDE_W, 3.333, GOLD);
  [1.653, 5.171, 8.688].forEach((x) => imagePlaceholder(slide, x, 2.962, 2.992, 2.738));
  heading(slide, 1.653, 1.227, 'Cut', 'Services', { color: CREAM });
  footerMark(slide);
  cornerTabs(slide, 'tl', 'br');
  SERVICES.forEach(([name, boxX, textX, textW], i) => {
    rect(slide, boxX, 5.318, 2.072, 0.76, GOLD);
    text(slide, name, {
      x: textX, y: 5.442, w: textW, h: 0.512, fontSize: 14, bold: true,
      color: CREAM, align: 'center', lineSpacingMultiple: 1.0,
    });
    text(slide, '0' + (i + 1), {
      x: 3.519 + i * 3.517, y: 2.962, w: 1.127, h: 0.512, fontSize: 48, bold: true,
      fontFace: HEAD, color: CREAM, align: 'right',
    });
  });
  [5.52, 9.037].forEach((x) => body(slide,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit.', x, 5.442, 1.721, 0.555,
    { color: CREAM, align: 'center' }));
  body(slide, LOREM_NOTE, 8.688, 1.667, 2.992, 0.693, { color: CREAM });
}

// ------------------------------------------------------ slides 10 & 11 ----
const TEAM = [
  ['Jacqueline Thompson', 'Master Barber'],
  ['Jamie Chastain', 'Grooming Expert'],
  ['Jonathan Patterson', 'Senior Barber'],
  ['Bartholomew Henderson', 'Customer Relations'],
];

function teamSlide(slide) {
  heading(slide, 1.653, 1.227, 'Meet The', 'Team');
  headerMark(slide);
  cornerTabs(slide, 'tl');
  TEAM.forEach((member, i) => rect(slide, i * 3.333, 3.194, 3.333, 4.306, i % 2 ? GOLD : INK));
  body(slide, LOREM_NOTE, 8.688, 1.667, 2.992, 0.693);
  TEAM.forEach(([name, role], i) => {
    const x = i * 3.333 + (i === 3 ? 0.408 : 0.576);
    text(slide, name, {
      x, y: 6.796, w: i === 3 ? 2.517 : 2.181, h: 0.201, fontSize: 14, bold: true,
      color: CREAM, align: 'center', lineSpacingMultiple: 1.0,
    });
    body(slide, role, i * 3.333 + 0.576, 7.058, 2.181, 0.201,
      { color: SILVER, align: 'center', lineSpacingMultiple: 1.0 });
  });
  TEAM.forEach((_, i) => imagePlaceholder(slide, 0.194 + i * 3.333, 3.389, 2.944, 3.167));
}

function leaderSlide(slide) {
  rect(slide, 10.0, 3.194, 3.333, 4.306, GOLD);
  rect(slide, 8.105, 1.949, 3.333, 4.306, INK);
  text(slide, 'Jacqueline Thompson', {
    x: 8.682, y: 5.551, w: 2.181, h: 0.201, fontSize: 14, bold: true, color: CREAM,
    align: 'center', lineSpacingMultiple: 1.0,
  });
  body(slide, 'Founder', 8.682, 5.812, 2.181, 0.201, { color: SILVER, align: 'center', lineSpacingMultiple: 1.0 });
  cornerTabs(slide, 'tl');
  heading(slide, 1.347, 2.485, 'Meet The', 'Leader');
  headerMark(slide);
  body(slide, LOREM_LONG, 1.347, 4.925, 4.739, 0.988);
  imagePlaceholder(slide, 8.3, 2.167, 2.944, 3.167);
}

// ----------------------------------------------------------- slide 12 -----
function breakSlide(slide) {
  // square brackets framing the title, drawn as three connector lines each
  [[3.49, 1], [9.843, -1]].forEach(([x, dir]) => {
    slide.addShape('line', { x, y: 1.007, w: 0, h: 5.486, line: { color: INK, width: 2 } });
    [1.007, 6.493].forEach((y) => slide.addShape('line', {
      x: dir > 0 ? x : x - 0.797, y, w: 0.797, h: 0, line: { color: INK, width: 2 },
    }));
  });
  body(slide, '"The right haircut not only creates style, but also defines personality."',
    4.867, 3.901, 3.599, 0.648, { fontSize: 12, align: 'center' });
  hero(slide, 4.333, 1.527, 4.667, 'Break', 'Slide', 'center');
  rect(slide, 2.692, 0, 0.218, SLIDE_H, INK);
  rect(slide, 10.423, 0, 0.218, SLIDE_H, INK);
  imagePlaceholder(slide, 0, 0, 2.692, SLIDE_H);
  imagePlaceholder(slide, 10.641, 0, 2.692, SLIDE_H);
  imagePlaceholder(slide, 2.91, 5.069, 7.513, 2.431);
}

// ----------------------------------------------------------- slide 13 -----
const SWOT = [
  ['S', 'Strenghts', 1.653, 3.014],
  ['O', 'Opportunities', 1.653, 4.71],
  ['W', 'Weaknesses', 6.667, 3.014],
  ['T', 'Threats', 6.667, 4.71],
];

function swotSlide(slide) {
  imagePlaceholder(slide, 9.139, 0, 4.194, 5.557);
  heading(slide, 1.653, 1.227, 'SWOT', 'Analysis');
  cornerTabs(slide, 'tl', 'br');
  SWOT.forEach(([letter, title, x, y]) => {
    slide.addShape('rect', {
      x, y, w: 1.361, h: 1.562, fill: { color: INK }, line: { type: 'none' },
      align: 'center', valign: 'middle', fontSize: 60, bold: true, fontFace: HEAD,
      color: WHITE,
    }, letter);
    rect(slide, x + 1.361, y, 3.652, 1.562, GOLD);
    label(slide, title, x + 1.683, y + 0.238, 2.823, 0.199, { fontSize: 14, color: CREAM });
    body(slide, LOREM_SHORT, x + 1.683, y + 0.546, 3.01, 0.778, { color: CREAM });
  });
  footerMark(slide);
}

// ----------------------------------------------------------- slide 14 -----
function gallerySlide(slide) {
  heading(slide, 3.248, 1.227, 'Gallery', 'Cut', { w: 2.992 });
  cornerTabs(slide, 'tl', 'br');
  footerMark(slide);
  ['2001', '2002'].forEach((year, i) => {
    const y = 3.526 + i * 1.634;
    body(slide, LOREM_MED, 3.248, y + 0.427, 4.798, 0.686);
    rect(slide, 3.248, y + 0.116, 0.201, 0.201, GOLD);
    text(slide, year, {
      x: 3.645, y, w: 1.099, h: 0.317, fontSize: 24, bold: true, fontFace: HEAD,
      color: GOLD, lineSpacingMultiple: 1.0,
    });
  });
  imagePlaceholder(slide, 0.361, 0, 1.986, 6.38);
  imagePlaceholder(slide, 8.947, 3.475, 4.025, 4.025);
  imagePlaceholder(slide, 9.66, 0, 3.312, 3.312);
}

// ----------------------------------------------------------- slide 15 -----
const PRICING = ['Basic', 'Advanced', 'Premium', 'VIP'];

function pricingSlide(slide) {
  PRICING.forEach((_, i) => rect(slide, 0.278 + i * 3.231, 0, 3.085, 5.662, INK));
  PRICING.forEach((_, i) => imagePlaceholder(slide, 0.278 + i * 3.231, 0, 3.085, 3.085));
  text(slide, 'Pricing List', {
    x: 4.032, y: 6.152, w: 5.27, h: 0.512, fontSize: 48, bold: true, fontFace: HEAD, align: 'center',
  });
  cornerTabs(slide, 'tl', 'br');
  footerMark(slide);
  PRICING.forEach((tier, i) => {
    const x = i * 3.231;
    rect(slide, 0.751 + x, 2.479, 2.138, 0.606, GOLD);
    text(slide, tier, {
      x: 1.053 + x, y: 2.617, w: 1.535, h: 0.317, fontSize: 18, bold: true, fontFace: HEAD,
      color: CREAM, align: 'center', lineSpacingMultiple: 1.0,
    });
    body(slide, PRICE_COPY, 0.566 + x, 3.335, 2.508, 0.936, { color: CREAM });
    body(slide, 'Start From', 0.566 + x, 4.704, 0.668, 0.191, { color: CREAM });
    text(slide, '$' + [99, 199, 299, 399][i] + '.99', {
      x: 0.566 + x, y: 4.901, w: 2.176, h: 0.477, fontSize: 36, bold: true, fontFace: HEAD,
      color: GOLD, lineSpacingMultiple: 1.0,
    });
  });
}

// ------------------------------------------------------ slides 16 & 17 ----
/** Laptop mock-up drawn from native shapes (stands in for the raster mockup). */
function laptopMock(slide) {
  slide.addShape('roundRect', { x: 5.780, y: 1.242, w: 5.919, h: 3.740, rectRadius: 0.09, fill: { color: '111111' }, line: { type: 'none' } });
  rect(slide, 5.986, 1.494, 5.521, 3.271, WHITE);
  slide.addShape('roundRect', { x: 5.118, y: 4.982, w: 7.236, h: 0.24, rectRadius: 0.05, fill: { color: 'E2E2E2' }, line: { type: 'none' } });
  slide.addShape('custGeom', {
    x: 5.118, y: 5.140, w: 7.236, h: 0.32, fill: { color: '9A9C9E' }, line: { type: 'none' },
    points: [{ x: 0.05, y: 0 }, { x: 7.186, y: 0 }, { x: 6.900, y: 0.32 }, { x: 0.336, y: 0.32 }, { close: true }],
  });
  slide.addShape('roundRect', { x: 8.05, y: 5.140, w: 1.36, h: 0.075, rectRadius: 0.03, fill: { color: '6E7072' }, line: { type: 'none' } });
}

/** Phone mock-up drawn from native shapes. */
function phoneMock(slide) {
  slide.addShape('roundRect', { x: 8.195, y: 0.790, w: 3.095, h: 5.919, rectRadius: 0.46, fill: { color: '1A1A1A' }, line: { type: 'none' } });
  slide.addShape('roundRect', { x: 8.396, y: 1.213, w: 2.709, h: 5.334, rectRadius: 0.20, fill: { color: 'FAFAFA' }, line: { type: 'none' } });
  slide.addShape('roundRect', { x: 9.093, y: 0.960, w: 1.300, h: 0.290, rectRadius: 0.13, fill: { color: '1A1A1A' }, line: { type: 'none' } });
  rect(slide, 11.270, 1.90, 0.055, 0.75, '2E2E2E');
}

function websiteMockSlide(slide) {
  rect(slide, 7.797, 0.43, 5.551, 3.886, INK);
  heading(slide, 1.653, 1.227, 'Unique', 'Mockup');
  cornerTabs(slide, 'tl', 'br');
  footerMark(slide);
  laptopMock(slide);
  [['Website', MOCK_COPY_A, 3.137], ['Information', MOCK_COPY_B, 4.609]].forEach(([title, copy, y]) => {
    body(slide, copy, 1.653, y + 0.306, 2.795, 0.707);
    label(slide, title, 1.653, y, 2.382, 0.205, { fontSize: 14, color: GOLD });
  });
  imagePlaceholder(slide, 5.987, 1.505, 5.521, 3.457);
}

function mobileMockSlide(slide) {
  rect(slide, 6.667, 1.597, 6.667, 4.306, GOLD);
  imagePlaceholder(slide, 8.367, 0.925, 2.746, 5.651);
  phoneMock(slide);
  heading(slide, 1.653, 1.602, 'Unique', 'Mockup');
  cornerTabs(slide, 'tl', 'br');
  footerMark(slide);
  [['Mobile Apps', MOCK_COPY_A, 3.852], [null, MOCK_COPY_B, 5.196]].forEach(([title, copy, y]) => {
    body(slide, copy, 1.653, title ? y + 0.307 : y, 3.347, 0.707);
    if (title) label(slide, title, 1.653, y, 2.382, 0.205, { fontSize: 14, color: GOLD });
  });
}

// ----------------------------------------------------------- slide 18 -----
// Europe silhouette: one polygon per country, coordinates in slide inches.
// 'A' = highlighted (brass), 'G' = land (grey).
const EUROPE = [
  ['A', '10.907 3.912,10.798 3.801,10.679 3.882,10.622 3.862,10.581 3.764,10.667 3.704,10.474 3.587,10.458 3.492,10.249 3.484,10.146 3.323,10.116 3.045,10.253 2.989,10.124 2.964,10.12 2.92,10.067 2.957,10.196 2.597,10.079 2.527,9.91 2.161,9.929 2.066,9.82 1.978,9.895 1.787,9.936 1.749,9.989 1.765,9.965 1.807,10.108 1.779,10.448 1.865,10.503 1.946,10.431 2.096,10.053 2.093,10.248 2.193,10.327 2.366,10.478 2.386,10.496 2.333,10.384 2.294,10.378 2.231,10.598 2.232,10.51 2.116,10.562 1.966,10.681 1.973,10.551 1.725,10.485 1.698,10.612 1.662,10.682 1.708,10.638 1.809,10.758 1.821,10.76 1.657,10.881 1.508,10.907 1.382,13.333 1.382,13.333 3.081,13.231 3.104,13.333 3.167,13.252 3.231,13.255 3.314,13.333 3.348,13.333 3.457,13.005 3.464,12.907 3.536,12.8 3.506,12.782 3.568,12.674 3.476,12.416 3.475,12.238 3.635,12.295 3.7,12.15 3.658,12.112 3.744,12.141 3.912,12.332 3.978,12.464 4.097,12.381 4.415,12.493 4.516,12.515 4.629,12.664 4.729,12.627 4.789,12.441 4.729,12.423 4.775,12.498 4.817,12.441 4.812,12.257 4.683,12.143 4.732,11.457 4.642,11.481 4.533,11.312 4.552,11.324 4.425,11.4 4.376,11.196 4.422,11.463 4.162,11.396 3.949,11.212 3.977,11.141 3.938,11.017 3.995'],
  ['G', '12.298 5.061,12.367 5.232,12.457 5.288,12.425 5.336,12.291 5.336,11.986 5.534,11.715 5.605,11.696 5.737,11.644 5.608,11.36 5.807,11.182 5.749,11.058 5.86,10.965 5.807,10.872 5.843,10.917 5.797,10.806 5.817,10.846 5.784,10.748 5.73,10.786 5.701,10.67 5.68,10.678 5.627,10.731 5.656,10.729 5.585,10.696 5.511,10.615 5.539,10.624 5.433,10.946 5.285,10.729 5.31,10.601 5.466,10.651 5.379,10.57 5.393,10.619 5.295,10.57 5.248,10.644 5.188,10.834 5.253,11.033 5.242,11.197 5.08,11.341 5.036,11.509 5.093,11.688 5.078,12.072 4.878,12.17 4.916,12.191 5.02,12.282 5.01'],
  ['G', '12.472 4.099,12.628 3.993,12.795 3.977,12.883 4.143,13.026 4.161,12.94 4.256,12.946 4.164,12.742 4.178,12.715 4.231,12.765 4.287,12.662 4.307,12.812 4.434,12.964 4.441,12.981 4.549,13.156 4.486,13.329 4.572,13.181 4.154,13.333 4.088,13.333 3.457,13.005 3.464,12.908 3.536,12.8 3.506,12.782 3.568,12.675 3.476,12.421 3.474,12.238 3.635,12.296 3.7,12.15 3.658,12.141 3.912,12.334 3.979'],
  ['G', '10.343 4.5,10.463 4.407,10.622 4.445,10.753 4.597,10.655 4.609,10.636 4.773,10.755 4.738,10.803 4.559,10.922 4.519,10.877 4.557,10.914 4.581,11.048 4.551,10.975 4.66,11.124 4.731,11.313 4.536,11.203 4.593,11.046 4.526,11.136 4.507,11.36 4.301,11.386 4.189,11.464 4.161,11.396 3.949,11.212 3.976,11.141 3.937,11.017 3.995,10.939 3.9,10.878 3.924,10.794 3.8,10.593 3.898,10.6 4.009,10.112 4.052,10.048 4.101,10.117 4.241,10.029 4.359,10.056 4.425,10.008 4.521,10.248 4.546'],
  ['G', '8.32 5.339,8.317 5.269,8.388 5.209,8.63 5.274,8.761 5.175,8.675 5.032,8.722 4.981,8.694 4.861,8.615 4.886,8.711 4.722,8.77 4.713,8.822 4.541,8.599 4.464,8.546 4.38,8.492 4.404,8.375 4.251,8.297 4.263,8.278 4.351,8.119 4.455,7.971 4.369,7.987 4.522,7.864 4.468,7.704 4.512,7.726 4.586,7.897 4.651,7.983 4.823,7.942 4.815,7.999 4.923,7.883 5.163,7.91 5.204'],
  ['G', '8.332 5.361,8.292 5.427,8.069 5.503,7.935 5.67,7.98 5.768,7.694 5.987,7.47 5.945,7.315 6.008,7.225 5.844,7.151 5.821,7.253 5.602,7.215 5.511,7.27 5.516,7.319 5.355,7.406 5.302,7.36 5.238,7.161 5.204,7.188 5.032,7.313 4.994'],
  ['G', '9.637 2.183,9.62 2.102,9.463 1.982,9.428 2.074,9.342 2.065,9.184 2.395,9.169 2.653,9.08 2.695,9.13 3.007,9.03 3.308,9.17 3.583,9.17 3.742,9.294 3.73,9.274 3.661,9.394 3.623,9.411 3.318,9.551 3.149,9.411 3.054,9.392 2.846,9.558 2.628,9.587 2.414,9.699 2.384'],
  ['G', '9.294 4.552,9.211 4.64,9.24 4.721,9.165 4.693,9.001 4.754,8.759 4.698,8.823 4.541,8.67 4.476,8.651 4.178,8.738 4.139,8.747 3.956,8.83 3.991,8.916 3.938,8.863 3.804,9.027 3.883,9.055 3.852,9.035 3.924,9.214 3.827,9.215 3.885,9.277 3.906,9.375 4.273,9.144 4.37'],
  ['G', '9.484 1.955,9.663 2.008,9.683 1.83,9.813 1.825,9.837 1.911,9.896 1.822,9.783 1.773,9.875 1.731,9.831 1.686,9.691 1.664,9.629 1.811,9.617 1.687,9.508 1.812,9.534 1.855,9.482 1.82,9.447 1.836,9.484 1.876,9.427 1.842,9.403 1.892,9.386 1.843,9.28 1.972,9.293 2.047,9.24 1.981,9.195 2.024,9.196 2.099,9.095 2.175,9.168 2.128,9.265 2.151,9.194 2.178,9.229 2.215,9.139 2.287,9.149 2.389,8.935 2.691,9.002 2.714,8.864 2.704,8.85 2.793,8.784 2.783,8.639 2.918,8.622 3.153,8.713 3.402,8.83 3.406,9.005 3.26,9.063 3.303,9.129 3.01,9.079 2.696,9.172 2.648,9.184 2.396,9.277 2.162,9.341 2.065,9.427 2.075,9.425 1.995'],
  ['G', '9.494 2.028,9.468 1.952,9.654 2.01,9.683 1.829,9.813 1.825,9.82 1.977,9.929 2.065,9.91 2.16,10.079 2.527,10.196 2.596,10.068 2.956,9.803 3.107,9.666 3.044,9.606 2.697,9.665 2.695,9.79 2.44,9.663 2.31,9.62 2.102'],
  ['G', '9.653 3.822,9.965 3.782,10.045 3.957,10.002 4.029,10.119 4.191,10.029 4.36,10.055 4.429,9.941 4.395,9.725 4.437,9.606 4.335,9.521 4.375,9.491 4.307,9.366 4.294,9.285 3.92,9.556 3.773'],
  ['G', '9.596 5.853,9.588 5.899,9.629 5.891,9.727 5.754,9.651 5.661,9.684 5.572,9.833 5.634,9.839 5.581,9.576 5.453,9.591 5.394,9.408 5.348,9.327 5.2,9.215 5.133,9.175 4.992,9.308 4.928,9.308 4.833,9.094 4.782,8.989 4.835,8.992 4.884,8.919 4.848,8.894 4.937,8.839 4.856,8.785 4.925,8.689 4.937,8.744 5.186,8.876 5.11,8.996 5.17,9.095 5.363,9.563 5.634,9.629 5.767'],
  ['G', '10.056 4.596,9.98 4.796,9.892 4.827,10.031 4.971,10.159 4.977,10.191 5.064,10.449 5.038,10.567 4.949,10.712 4.958,10.755 4.739,10.621 4.765,10.555 4.577,10.42 4.457'],
  ['G', '7.799 4.185,7.68 4.268,7.718 4.298,7.912 4.24,8.18 4.273,8.267 4.23,8.201 4.187,8.318 4.077,8.285 4.016,8.185 4.009,8.221 3.968,8.18 3.908,8.213 3.92,8.118 3.665,7.982 3.587,8.063 3.572,8.018 3.55,8.147 3.439,7.966 3.388,8.078 3.272,7.947 3.266,7.881 3.411,7.847 3.349,7.808 3.369,7.88 3.43,7.775 3.592,7.856 3.542,7.815 3.643,7.86 3.568,7.904 3.634,7.859 3.74,8.001 3.723,7.951 3.774,8.004 3.823,7.959 3.926,7.844 3.901,7.864 4.02,7.761 4.065,7.889 4.152,7.989 4.12'],
  ['G', '10.175 3.63,10.157 3.56,10.249 3.48,10.458 3.492,10.475 3.587,10.667 3.704,10.577 3.78,10.656 3.88,10.594 3.898,10.601 4.009,10.168 4.036,10.046 4.101,9.982 3.817,10.137 3.757,10.125 3.667'],
  ['G', '12.3 4.849,12.498 4.817,12.422 4.773,12.441 4.729,12.616 4.798,12.684 4.676,12.838 4.822,12.95 4.852,12.85 4.9,12.879 5.018,12.832 5.178,12.735 5.134,12.754 5.037,12.701 5.006,12.559 5.166,12.512 5.002,12.399 4.98,12.428 4.935'],
  ['G', '10.158 5.032,10.442 5.04,10.651 4.944,10.718 4.983,10.653 5.131,10.722 5.179,10.589 5.223,10.546 5.315,10.263 5.368,10.185 5.255,10.22 5.138,10.145 5.08'],
  ['G', '6.953 2.69,6.844 2.675,6.95 2.765,6.85 2.828,7.049 2.975,7.316 2.93,7.401 2.859,7.353 2.766,7.385 2.69,7.342 2.707,7.327 2.65,7.236 2.688,7.201 2.651,7.187 2.704,7.178 2.636,7.115 2.677,7.094 2.619,7.01 2.7,7.01 2.526,6.984 2.598,6.95 2.532,6.922 2.602,6.857 2.593,6.963 2.632'],
  ['G', '9.589 4.646,9.744 4.632,9.909 4.51,10.091 4.549,9.979 4.796,9.755 4.89,9.644 4.879,9.522 4.775,9.532 4.671'],
  ['G', '10.498 5.33,10.591 5.26,10.605 5.357,10.354 5.43,10.434 5.49,10.372 5.47,10.392 5.524,10.246 5.488,10.35 5.627,10.284 5.687,10.441 5.68,10.51 5.753,10.407 5.724,10.463 5.809,10.308 5.747,10.086 5.754,10.137 5.692,10.027 5.64,10.079 5.463,10.393 5.313'],
  ['G', '7.158 5.208,7.23 5.187,7.228 5.232,7.36 5.239,7.406 5.302,7.318 5.357,7.269 5.518,7.216 5.512,7.23 5.708,7.094 5.828,6.982 5.8,7.049 5.647,6.992 5.576,7.147 5.327'],
  ['G', '9.374 4.57,9.396 4.517,9.56 4.538,9.59 4.662,9.531 4.671,9.51 4.803,9.358 4.842,9.189 4.819,9.169 4.771,9.066 4.812,8.932 4.777,8.955 4.712,8.999 4.753,9.165 4.692,9.24 4.72,9.21 4.639,9.294 4.552'],
  ['G', '9.803 4.856,9.891 4.826,10.027 4.969,10.16 4.978,10.144 5.08,10.22 5.138,10.184 5.258,10.109 5.28,10.12 5.228,10.016 5.17,9.979 5.231,9.855 5.171,9.88 5.093,9.827 5.065'],
  ['G', '9.246 4.33,9.384 4.273,9.492 4.309,9.521 4.375,9.606 4.335,9.715 4.425,9.57 4.554,9.396 4.517,9.32 4.571,9.143 4.37'],
  ['G', '7.677 4.004,7.402 4.03,7.428 4.002,7.37 4.005,7.421 3.984,7.366 3.966,7.508 3.901,7.43 3.896,7.525 3.836,7.439 3.785,7.499 3.744,7.475 3.69,7.587 3.725,7.627 3.679,7.593 3.653,7.71 3.604,7.723 3.651,7.625 3.706,7.758 3.787'],
  ['G', '12.393 4.842,12.443 4.812,12.312 4.692,12.143 4.732,12.018 4.683,11.754 4.703,11.91 4.773,11.96 4.923'],
  ['G', '9.78 3.4,9.82 3.371,9.916 3.458,9.979 3.296,10.156 3.321,10.248 3.483,10.156 3.561,10.013 3.496,9.815 3.54,9.753 3.605'],
  ['G', '9.793 3.557,10.013 3.497,10.158 3.562,10.137 3.757,9.982 3.82,9.884 3.695,9.781 3.679,9.752 3.595'],
  ['G', '9.813 5.003,9.841 4.95,9.777 4.874,9.674 4.893,9.558 4.814,9.486 4.86,9.469 4.956,9.301 4.968,9.342 5.052,9.376 4.983,9.424 5.01,9.502 5.11,9.422 5.059,9.472 5.154,9.41 5.1,9.481 5.167,9.738 5.276,9.509 4.991'],
  ['G', '9.82 4.987,9.879 5.151,9.827 5.165,9.808 5.288,9.508 5.041,9.53 4.98'],
  ['G', '9.624 4.519,9.757 4.404,9.796 4.45,9.912 4.397,10.027 4.422,10.004 4.526,9.877 4.519,9.736 4.636,9.588 4.639,9.569 4.545'],
  ['G', '8.689 4.858,8.718 4.93,8.839 4.856,8.894 4.937,8.913 4.853,8.992 4.883,9.023 4.791,8.93 4.777,8.949 4.727,8.851 4.685,8.711 4.722,8.615 4.886'],
  ['G', '9.841 3.203,9.959 3.097,10.122 3.083,10.106 3.178,10.165 3.284,10.086 3.34,9.981 3.296,9.934 3.33,9.934 3.268,9.88 3.282'],
  ['G', '12.141 4.905,12.264 4.833,12.34 4.852,12.429 4.937,12.398 4.979,12.512 5.002,12.536 5.087,12.436 5.025,12.191 5.02'],
  ['G', '8.616 4.027,8.668 3.986,8.751 4.016,8.738 4.138,8.651 4.177,8.649 4.314,8.574 4.217,8.444 4.214,8.559 4.177,8.502 4.143,8.566 4.026,8.581 4.108,8.648 4.1'],
  ['G', '10.62 4.757,10.656 4.609,10.753 4.597,10.608 4.434,10.403 4.438,10.556 4.577'],
  ['G', '8.372 4.254,8.552 4.213,8.639 4.261,8.62 4.312,8.677 4.356,8.626 4.458,8.546 4.38,8.491 4.404'],
  ['G', '10.011 5.63,9.911 5.54,9.92 5.26,10.003 5.302,10.022 5.429,10.087 5.482'],
  ['G', '10.003 5.625,10.087 5.482,10.02 5.427,10.011 5.313,9.92 5.26,9.913 5.542'],
  ['G', '8.91 3.527,8.996 3.45,8.98 3.58,9.027 3.612,8.932 3.707,8.939 3.816,8.858 3.802,8.82 3.718,8.831 3.583,8.865 3.608,8.879 3.551,8.913 3.601,8.944 3.536,8.832 3.568'],
  ['G', '10.003 5.625,10.086 5.49,10.001 5.304,9.92 5.26,9.913 5.54'],
  ['G', '9.567 5.869,9.546 6.063,9.26 5.948,9.286 5.89'],
  ['G', '10.06 5.464,10.01 5.397,10.036 5.318,10.169 5.257,10.234 5.29,10.243 5.394'],
  ['G', '8.808 5.567,8.946 5.528,8.951 5.786,8.83 5.798'],
  ['G', '10.189 5.791,10.155 5.828,10.241 5.964,10.279 5.925,10.331 5.995,10.342 5.942,10.407 5.976,10.334 5.851,10.378 5.881,10.4 5.823'],
  ['G', '9.304 4.882,9.308 4.833,9.53 4.771,9.56 4.815,9.455 4.964,9.308 4.97'],
  ['G', '9.894 5.359,9.91 5.269,9.979 5.23,9.834 5.159,9.801 5.295'],
  ['G', '7.727 3.642,7.783 3.644,7.818 3.732,7.757 3.79,7.625 3.707'],
  ['G', '9.793 3.688,9.884 3.695,9.927 3.775,9.696 3.809,9.7 3.74,9.776 3.739'],
  ['G', '10.11 5.279,10.02 5.347,9.943 5.246,10.017 5.17,10.12 5.228'],
  ['G', '10.484 6.086,10.767 6.07,10.616 6.132'],
  ['G', '8.877 5.452,8.908 5.494,8.928 5.455,8.929 5.284,8.851 5.362'],
  ['G', '12.324 5.097,12.434 5.187,12.519 5.182,12.446 5.092'],
  ['G', '11.37 5.906,11.589 5.819,11.541 5.904'],
  ['G', '9.123 3.785,9.134 3.647,9.03 3.7'],
  ['G', '13.333 3.194,13.252 3.231,13.255 3.314,13.333 3.349'],
];

function mapSlide(slide) {
  EUROPE.forEach(([kind, coords]) => {
    const pts = coords.split(',').map((pair) => {
      const [px, py] = pair.split(' ');
      return { x: +px, y: +py };
    });
    pts.push({ close: true });
    slide.addShape('custGeom', {
      x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, points: pts,
      fill: { color: kind === 'A' ? GOLD : SILVER },
      line: { color: MAP_EDGE, width: 2 },
    });
  });
  heading(slide, 1.653, 1.377, 'Unique', 'Europe Map', { w: 4.476 });
  body(slide, LOREM_LONG, 1.653, 3.153, 5.103, 0.988);
  footerMark(slide);
  statBlock(slide, 1.653, 4.667, '+100', 'Cust/day', LOREM_TINY);
  statBlock(slide, 3.763, 4.667, '+1.2K', 'Cust/week', LOREM_TINY);
  cornerTabs(slide, 'tl');
}

// ----------------------------------------------------------- slide 19 -----
function testimonialSlide(slide) {
  heading(slide, 1.653, 1.227, 'Client', 'Testimonial', { w: 4.291 });
  cornerTabs(slide, 'tl', 'br', 'bl');
  headerMark(slide);
  rect(slide, 1.653, 2.778, 5.555, 4.722, INK);
  rect(slide, 7.208, 2.778, 5.555, 4.722, GOLD);
  [['Francisco Andrade', 2.056], ['Francois Mercer', 7.611]].forEach(([name, x], i) => {
    body(slide, QUOTE, x, 6.02, 4.749, 0.707, { color: CREAM });
    label(slide, name, x, 5.651, 2.382, 0.205, { fontSize: 14, color: CREAM });
    for (let s = 0; s < 5; s++) {
      slide.addShape('star5', {
        x: x + s * 0.2548, y: 6.914, w: 0.18, h: 0.18,
        fill: { color: CREAM }, line: { type: 'none' },
      });
    }
    imagePlaceholder(slide, [2.056, 7.612][i], 2.778, 2.527, 2.527);
  });
}

// ----------------------------------------------------------- slide 20 -----
function thankYouSlide(slide) {
  rect(slide, 0, 2.778, 5.555, 4.722, INK);
  imagePlaceholder(slide, 1.048, 1.737, 5.555, 4.722);
  hero(slide, 6.958, 1.527, 4.667, 'Thank', 'You', 'right');
  cornerTabs(slide, 'tl', 'br');
  headerMark(slide);
  rect(slide, 1.406, 3.525, 3.326, 2.624, CREAM);
  contactIcons(slide);
  [['www.yourwebsite.com', 4.746, 1.759], ['123 Anywhere St., Any City, ST 12345', 5.14, 2.155],
    ['123-456-7890', 5.498, 2.155]].forEach(([line, y, w]) =>
    body(slide, line, 2.156, y, w, 0.231, { color: INK }));
  body(slide, LOREM_LONG, 8.19, 4.221, 3.358, 1.424, { align: 'right' });
}

/** Small brass glyphs on the contact card: globe, map pin, handset. */
function contactIcons(slide) {
  const ring = (x, y, w, h) => slide.addShape('ellipse', {
    x, y, w, h, fill: { type: 'none' }, line: { color: GOLD, width: 1.25 },
  });
  ring(1.840, 4.812, 0.181, 0.181);
  ring(1.900, 4.812, 0.061, 0.181);
  slide.addShape('line', { x: 1.840, y: 4.902, w: 0.181, h: 0, line: { color: GOLD, width: 1 } });
  slide.addShape('teardrop', {
    x: 1.840, y: 5.173, w: 0.165, h: 0.165, rotate: 225,
    fill: { color: GOLD }, line: { type: 'none' },
  });
  slide.addShape('ellipse', { x: 1.884, y: 5.216, w: 0.062, h: 0.062, fill: { color: CREAM }, line: { type: 'none' } });
  ring(1.828, 5.555, 0.202, 0.181);
  slide.addShape('ellipse', { x: 1.874, y: 5.601, w: 0.110, h: 0.090, fill: { color: CREAM }, line: { type: 'none' } });
}

// ------------------------------------------------------------------ main --
const BUILDERS = [
  coverSlide, welcomeSlide, tableOfContentsSlide, aboutSlide, visionSlide,
  missionSlide,
  (s) => timelineSlide(s, TIMELINE_A, 1.653),
  (s) => timelineSlide(s, TIMELINE_B, 0),
  servicesSlide, teamSlide, leaderSlide, breakSlide, swotSlide, gallerySlide,
  pricingSlide, websiteMockSlide, mobileMockSlide, mapSlide, testimonialSlide,
  thankYouSlide,
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'CUT', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'CUT';
  pptx.theme = { headFontFace: HEAD, bodyFontFace: SANS };
  pptx.title = 'Cut Barbershop';

  BUILDERS.forEach((builder) => {
    const slide = pptx.addSlide();
    slide.background = { color: CREAM };
    builder(slide);
  });

  return pptx.writeFile({
    fileName: path.join(__dirname, '04f63833-9bb8-4ef5-8872-297683679a02_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f));
