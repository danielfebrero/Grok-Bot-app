/*
 * "Kunua" travel presentation template - 42 slides, 13.333 x 7.5 in (16:9).
 *
 * The original deck's picture placeholders are all empty (they render as blank
 * areas), so they are not reproduced. The two real photographs (slides 24, 25)
 * become grey "[image]" boxes.
 *
 * Run: node 0cb31026-d3cd-4ffb-a13f-18a3573e4bb5_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette --
const RED = 'C75233';       // theme dk1  - brand terracotta
const PEACH = 'FEBF9C';     // theme accent2
const TEAL = '70B0A7';      // theme accent4
const INK = '262626';
const PANEL = '6D6D6D';     // slide 26's 60%-opaque scrim, flattened
const GREY = 'D9D9D9';
const BLACK = '000000';
const WHITE = 'FFFFFF';

const SERIF = 'Playfair Display';   // headings
const SANS = 'Gudea';               // body copy

// Strings the template reuses on many slides.
const URL = 'W   W   W   .   K   U   N   U   A   .   C   O   M';
const TAG = ['Take Your Business', 'To The Next Level'];
const LOREM_A = 'Proactively envisioned multimedia based expertise and cross-media growth strategies. Seamlessly visualize quality intellectual capital without superior collaboration. Leverage agile frameworks to provide a robust.';
const LOREM_B = 'Proactively envisioned multimedia based expertise and cross-media growth strategies quality Make a type specimen book unknown printer took type and good scrambled.';
const LOREM_C = 'Proactively envisioned multimedia based expertise and cross-media growth strategies. Seamlessly visualize quality intellectual capital without superior.';
const LOREM_D = 'Bring the table win-win survival strategies ensure proactive dominan. At the end of the day, going forward, for new normal .';
const LOREM_E = 'Predominate extensible testing procedures for reliable good ';
const LOREM_F = 'Capitalize on low hanging fruit to identify a ballpark value content in provide a robust strategies. ';
const LOREM_G = 'Leverage agile frameworks to provide a robust synopsis for high level views new normal views new normal that ensure coordinate..';
const LOREM_H = 'Interactively coordinates make commerce make from great process outside pursue thinking scalable.';

// ---------------------------------------------------------------- helpers --

// Text box. `r` is [x, y, w, h] in inches; `body` a string or array of lines.
function txt(s, r, body, o) {
  o = o || {};
  const lines = Array.isArray(body) ? body : [body];
  const run = {
    fontFace: o.font || SANS,
    fontSize: o.size || 10,
    bold: !!o.bold,
    italic: !!o.italic,
    color: o.color || BLACK,
    align: o.align || 'left',
    lineSpacing: o.lhPt,
    lineSpacingMultiple: o.lh,
  };
  const parts = lines.map((t, i) => ({
    text: t === '' ? ' ' : t,
    options: Object.assign({ breakLine: i < lines.length - 1 }, run),
  }));
  s.addText(parts, Object.assign({
    x: r[0], y: r[1], w: r[2], h: r[3],
    margin: 0, wrap: true, valign: o.mid ? 'middle' : 'top',
    rotate: o.rot,
    fill: o.fill ? { color: o.fill } : undefined,
  }, run));
}

// Filled rectangle.
function box(s, r, color, o) {
  o = o || {};
  s.addShape('rect', {
    x: r[0], y: r[1], w: r[2], h: r[3],
    fill: color === null ? { type: 'none' } : { color: color },
    line: o.line ? { color: o.line, width: o.lw || 1 } : { type: 'none' },
  });
}

// Outlined rectangle (no fill).
function frame(s, r, color, lw) {
  box(s, r, null, { line: color, lw: lw });
}

// Straight rule. Height 0 -> horizontal, width 0 -> vertical.
function line(s, r, color, pt) {
  s.addShape('line', {
    x: r[0], y: r[1], w: r[2], h: r[3],
    line: { color: color, width: pt || 1 },
  });
}

// Polygon / free-form drawn from fractional (0..1) coordinates of its box.
function poly(s, r, pts, color) {
  s.addShape('custGeom', {
    x: r[0], y: r[1], w: r[2], h: r[3],
    fill: { color: color },
    line: { type: 'none' },
    points: pts.map((p, i) => ({ x: p[0] * r[2], y: p[1] * r[3], moveTo: i === 0 }))
      .concat([{ close: true }]),
  });
}

function ellipse(s, r, color) {
  s.addShape('ellipse', {
    x: r[0], y: r[1], w: r[2], h: r[3],
    fill: { color: color }, line: { type: 'none' },
  });
}

// Wedge of a (squashed) circle: angles in degrees, 0 = 3 o'clock, clockwise.
function wedge(s, r, a0, a1, color) {
  s.addShape('pie', {
    x: r[0], y: r[1], w: r[2], h: r[3],
    angleRange: [a0, a1],
    fill: { color: color }, line: { type: 'none' },
  });
}

// Stand-in for a bitmap in the source deck.
function photo(s, r) {
  box(s, r, 'E8E4E1');
  txt(s, [r[0], r[1] + r[3] / 2 - 0.15, r[2], 0.3], '[image]',
    { align: 'center', color: '8A8A8A', size: 11 });
}

// The spaced-out "W W W . K U N U A . C O M" footer.
function www(s, r, o) {
  txt(s, r, URL, Object.assign({ size: 9 }, o));
}

// The two-line "Take Your Business / To The Next Level" lockup.
function tagline(s, r, o) {
  txt(s, r, TAG, Object.assign({ font: SERIF, size: 14, bold: true }, o));
}

// -------------------------------------------------------------- pictograms --
// The template's icons are complex free-form art; they are re-drawn here as
// simple flat pictograms of the same size, colour and position.
// `cx`/`cy` are the icon centre, `d` its nominal size, in inches.
const GLYPHS = {
  pin: function (s, cx, cy, d, c) {           // map pin
    ellipse(s, [cx - d * 0.34, cy - d * 0.46, d * 0.68, d * 0.68], c);
    poly(s, [cx - d * 0.3, cy - d * 0.06, d * 0.6, d * 0.56], [[0.5, 1], [0, 0], [1, 0]], c);
  },
  target: function (s, cx, cy, d, c, bg) {    // concentric rings + centre dot
    ring(s, cx, cy, d * 0.95, d * 0.11, c, bg);
    ring(s, cx, cy, d * 0.55, d * 0.1, c, bg);
    ellipse(s, [cx - d * 0.11, cy - d * 0.11, d * 0.22, d * 0.22], c);
  },
  person: function (s, cx, cy, d, c) {        // head + shoulders
    ellipse(s, [cx - d * 0.17, cy - d * 0.5, d * 0.34, d * 0.34], c);
    box(s, [cx - d * 0.45, cy - d * 0.1, d * 0.9, d * 0.14], c);
    poly(s, [cx - d * 0.34, cy + d * 0.02, d * 0.68, d * 0.48], [[0.15, 0], [0.85, 0], [1, 1], [0.62, 1], [0.5, 0.45], [0.38, 1], [0, 1]], c);
  },
  flag: function (s, cx, cy, d, c) {          // pennant on a pole
    box(s, [cx - d * 0.4, cy - d * 0.5, d * 0.12, d], c);
    poly(s, [cx - d * 0.28, cy - d * 0.46, d * 0.72, d * 0.5], [[0, 0], [1, 0], [1, 1], [0, 1]], c);
  },
  aperture: function (s, cx, cy, d, c) {      // six-bladed camera iris
    for (let i = 0; i < 6; i++) {
      wedge(s, [cx - d / 2, cy - d / 2, d, d], i * 60 + 5, i * 60 + 55, c);
    }
  },
  swap: function (s, cx, cy, d, c, bg) {      // up + down arrows inside a disc
    ellipse(s, [cx - d / 2, cy - d / 2, d, d], c);
    poly(s, [cx - d * 0.32, cy - d * 0.3, d * 0.28, d * 0.26], [[0.5, 0], [1, 1], [0, 1]], bg);
    box(s, [cx - d * 0.27, cy - d * 0.06, d * 0.18, d * 0.36], bg);
    poly(s, [cx + d * 0.04, cy + d * 0.04, d * 0.28, d * 0.26], [[0.5, 1], [0, 0], [1, 0]], bg);
    box(s, [cx + d * 0.09, cy - d * 0.3, d * 0.18, d * 0.36], bg);
  },
  clipboard: function (s, cx, cy, d, c, bg) { // clipboard with lines
    box(s, [cx - d * 0.12, cy - d * 0.52, d * 0.24, d * 0.16], c);
    box(s, [cx - d * 0.42, cy - d * 0.42, d * 0.84, d * 0.94], c);
    [0.06, 0.26, 0.46].forEach(function (t) {
      box(s, [cx - d * 0.26, cy - d * 0.28 + d * t, d * 0.52, d * 0.1], bg);
    });
  },
  check: function (s, cx, cy, d, c, bg) {     // clipboard with a tick
    box(s, [cx - d * 0.12, cy - d * 0.52, d * 0.24, d * 0.16], c);
    box(s, [cx - d * 0.42, cy - d * 0.42, d * 0.84, d * 0.94], c);
    poly(s, [cx - d * 0.26, cy - d * 0.18, d * 0.52, d * 0.52], [[0.08, 0.45], [0.22, 0.3], [0.42, 0.55], [0.8, 0.05], [0.94, 0.2], [0.42, 0.9]], bg);
  },
  cart: function (s, cx, cy, d, c) {          // shopping trolley
    poly(s, [cx - d * 0.45, cy - d * 0.4, d * 0.9, d * 0.5], [[0, 0], [0.18, 0], [0.3, 0.25], [1, 0.25], [0.82, 1], [0.36, 1], [0.3, 0.25], [0.18, 0.18], [0, 0.18]], c);
    ellipse(s, [cx - d * 0.14, cy + d * 0.2, d * 0.16, d * 0.16], c);
    ellipse(s, [cx + d * 0.16, cy + d * 0.2, d * 0.16, d * 0.16], c);
  },
  truck: function (s, cx, cy, d, c) {         // delivery van
    box(s, [cx - d * 0.48, cy - d * 0.3, d * 0.58, d * 0.5], c);
    poly(s, [cx + d * 0.08, cy - d * 0.12, d * 0.42, d * 0.32], [[0, 0], [0.6, 0], [1, 0.5], [1, 1], [0, 1]], c);
    ellipse(s, [cx - d * 0.34, cy + d * 0.14, d * 0.18, d * 0.18], c);
    ellipse(s, [cx + d * 0.16, cy + d * 0.14, d * 0.18, d * 0.18], c);
  },
};

// Ring = filled disc with a disc of the background colour punched out.
function ring(s, cx, cy, d, t, c, bg) {
  ellipse(s, [cx - d / 2, cy - d / 2, d, d], c);
  ellipse(s, [cx - d / 2 + t, cy - d / 2 + t, d - 2 * t, d - 2 * t], bg || WHITE);
}

function glyph(s, cx, cy, d, kind, color, bg) {
  GLYPHS[kind](s, cx, cy, d, color, bg || WHITE);
}

// ------------------------------------------------------------ the slides --

function slide01(s) {                                   // cover
  box(s, [5.971, 0, 7.362, 7.5], RED);
  tagline(s, [11.069, 6.103, 1.497, 0.404], { size: 12, color: WHITE, align: 'right' });
  www(s, [10.998, 2.426, 3, 0.135], { size: 8, color: WHITE, align: 'right', rot: 270 });
  txt(s, [0.938, 5.622, 2.913, 0.884], 'Interactively proactive commerce process centric outside the box into thinking pursue scalable into customers for into based services through its star shape envisioned for a multimedia.', { size: 9, align: 'justify', lh: 1.5 });
  txt(s, [0.938, 2.668, 4.265, 1.616], 'Kunua', { font: SERIF, size: 96, bold: true });
  txt(s, [0.938, 0.994, 4.094, 0.337], 'Travel Presentation Template', { font: SERIF, size: 20, bold: true });
}

function slide02(s) {                                   // Let's Start Your Travel
  box(s, [0, 2.646, 13.333, 4.854], PEACH);
  txt(s, [1.216, 1.2, 2.919, 0.852], ['Let\u2019s Start', 'Your Travel'], { font: SERIF, size: 28, bold: true, lhPt: 30, mid: true });
  txt(s, [1.216, 3.284, 2.919, 1.988], ['Interactively procrastinate high-payoff content without backward compatible data. Quickly cultivate optimal processes and tactical.', '', 'Quickly disseminate superior deliverables whereas web-enabled applications. Quickly drive clicks and mortar catalysts for change. Continually reintermediate integrated.'], { lh: 1.5 });
  txt(s, [6.505, 6.117, 5.628, 0.539], '\u201c To be fully seen by somebody, then, and be loved anyhow this is a human offering that can border on miraculous \u201c', { size: 16, italic: true, align: 'right' });
  www(s, [1.216, 6.521, 1.971, 0.135], { size: 8 });
}

function slide03(s) {                                   // Background
  txt(s, [6.667, 1.21, 2.637, 0.657], 'Interactively coordinate proactive e-commerce via process-centric "outside the box" thinking pursue scalable customer service.', { size: 9, align: 'justify', lh: 1.5 });
  txt(s, [9.997, 1.21, 2.637, 0.657], 'Collaboratively administrate powered markets via  networks. Dynamic procrastinate B2C users after installed base benefits dramatic visualize.', { size: 9, align: 'justify', lh: 1.5 });
  txt(s, [0.7, 1.236, 3.203, 0.606], 'Background', { font: SERIF, size: 36, bold: true, mid: true });
  txt(s, [0.7, 3.304, 3.127, 0.884], 'Proactively envisioned for multimedia based expertise and cross media growth strategies. Seamlessly visualize quality intellectual capital without superior agile frameworks to provide a robust visualize special good.', { size: 9, align: 'justify', lh: 1.5 });
  www(s, [0.7, 6.665, 2.377, 0.135], { size: 8 });
}

function slide04(s) {                                   // Our History
  txt(s, [7.389, 3.031, 5.354, 1.344], ['Interactively proactive commerce process centric outside making box thinking pursue scalable into making the services quality more intellectual capital services.', '', 'Collaboratively administrate empowered markets via plug and play networks. Dynamic procrastinate B2C users. After installed base benefits dramatic. Interactively coordinate proactive e-commerce via process centric outside proactively good for all customers.'], { size: 9, align: 'justify', lh: 1.5 });
  www(s, [10.366, 6.775, 2.377, 0.135], { size: 8, align: 'right' });
  txt(s, [0.591, 1.111, 5.354, 0.808], 'Our History', { font: SERIF, size: 48, bold: true, mid: true });
  tagline(s, [8.627, 1.28, 4.116, 0.471], { align: 'right' });
}

function slide05(s) {                                   // Target Audience - 2x2 numbered blurbs
  const BLURB = 'Proactively envisioned based expertise and cross-media to growth strategies seamlessly visualize quality.';
  txt(s, [0.73, 1.001, 5.886, 0.673], 'Target Audience', { font: SERIF, size: 40, bold: true, mid: true });
  [['01', 7.525, 2.676], ['02', 10.293, 2.676], ['03', 7.525, 4.639], ['04', 10.293, 4.639]].forEach(function (c) {
    const n = c[0], x = c[1], y = c[2];
    txt(s, [x, y, 0.616, 0.567], n, { size: 36, bold: true, lhPt: 40 });
    txt(s, [x + 0.721, y, 1.59, 0.236], 'Your Title Here', { size: 14, bold: true });
    txt(s, [x + 0.721, y + 0.29, 1.59, 0.879], BLURB, { size: 9, lh: 1.5 });
  });
  www(s, [0.73, 6.586, 2.377, 0.135], { size: 8 });
  tagline(s, [10.677, 1.102, 1.926, 0.471], { align: 'right' });
}

function slide06(s) {                                   // Business Model Overview
  txt(s, [1.239, 1.012, 6.605, 0.539], 'Business Model Overview', { font: SERIF, size: 32, bold: true, color: INK, mid: true });
  txt(s, [7.928, 2.568, 4.167, 0.539], '2026 Financial serenity has taken possession of my entire soul, like these', { size: 16, bold: true, mid: true });
  txt(s, [7.928, 3.331, 4.167, 0.657], 'Collaboration agile frameworks to provide a robust. Seamlessly visualize quality intellectual capital without superior collaboration and idea sharing coordinate proactive for good customers experience.', { size: 9, lh: 1.5 });
  www(s, [9.304, 1.206, 2.79, 0.151], { align: 'right' });
}

function slide07(s) {                                   // Mission and Vision Statements
  txt(s, [8.399, 3.301, 3.746, 0.841], 'Mission and Vision Statements', { font: SERIF, size: 28, bold: true, lhPt: 30, mid: true });
  txt(s, [8.399, 4.812, 3.746, 1.988], ['Interactively procrastinate high-payoff content with backward compatible data. Quickly cultivate optimal processes and tactical architectures completely iterate incubate. ', '', 'Quickly disseminate superior deliverables whereas web-enabled applications. Quickly drive clicks-and-mortar catalysts change.. Continually reintermediate integrated processes through technically sound intellectual capital.'], { align: 'justify', lh: 1.5 });
  [['01', 1.186, RED, WHITE], ['02', 5.207, PEACH, BLACK], ['03', 9.227, TEAL, WHITE]].forEach(function (c) {
    const n = c[0], x = c[1];
    box(s, [x, 1.313, 0.611, 0.611], c[2]);
    txt(s, [x + 0.106, 1.484, 0.4, 0.269], n, { size: 16, bold: true, color: c[3], align: 'center' });
    txt(s, [x + 0.792, 1.262, 2.128, 0.236], 'Your Content Name', { size: 14, bold: true });
    txt(s, [x + 0.792, 1.565, 2.128, 0.474], 'Make a type book known printer took a galley of type good user.', { lh: 1.5 });
  });
}

function slide08(s) {                                   // Additional Services - 2x2 icon tiles
  txt(s, [0.818, 0.814, 5.933, 0.421], 'Additional Services', { font: SERIF, size: 28, bold: true, lhPt: 30, mid: true });
  txt(s, [0.818, 2.393, 5.933, 0.731], 'Proactively envisioned multimedia based expertise and cross-media growth strategies. Seamlessly visualize quality intellectual capital without superior collaboration. Globally incubate standards compliant channels before scalable benefits extensible testing fruit to identify.', { lh: 1.5 });
  const TILE = 'Interactively coordinate proactive e-commerce via process';
  [[0.811, 3.579, TEAL, 'pin'], [3.72, 3.579, RED, 'target'],
   [0.811, 4.661, RED, 'person'], [3.724, 4.661, TEAL, 'flag']].forEach(function (c) {
    const x = c[0], y = c[1];
    box(s, [x, y, 0.7, 0.7], c[2]);
    glyph(s, x + 0.35, y + 0.35, 0.37, c[3], WHITE, c[2]);
    txt(s, [x + 0.92, y - 0.015, 1.416, 0.73], TILE, { lh: 1.5 });
  });
  www(s, [0.818, 6.535, 2.7, 0.151]);
}

function slide09(s) {                                   // Connecting Destinations
  box(s, [0, 0, 13.333, 4.594], PEACH);
  txt(s, [8.618, 0.866, 3.924, 0.983], 'Interactively procrastinate high-payoff content without backward compatible data. Quickly cultivate optimal processes tactical quickly disseminate superior deliverables whereas web-enabled applications. Quickly drive clicks-and-mortar catalysts.', { lh: 1.5 });
  txt(s, [4.183, 5.792, 3.643, 0.841], 'Connecting Destinations', { font: SERIF, size: 28, bold: true, align: 'right', lhPt: 30, mid: true });
  www(s, [0.792, 6.137, 2.6, 0.151]);
}

// Slides 10 and 11 share a timeline: a dotted spine with numbered squares and
// a labelled progress bar for each step.
const TIMELINE_STEPS = [
  ['01', PEACH, LOREM_D, 2.881, '95%'],
  ['02', RED, 'Collaboratively administrate empowered markets via plug-and-play networks. Dynamic procrastinate B2C users after installed.', 3.074, '100%'],
  ['03', TEAL, 'Expertise and cross-media growth strategies. Seamlessly visualize quality intellectual capital without superior.', 2.688, '85%'],
  ['04', TEAL, LOREM_D, 2.881, '95%'],
  ['05', PEACH, 'Collaboratively administrate empowered markets via plug-and-play networks. Dynamic procrastinate B2C users after installed.', 3.074, '100%'],
  ['06', RED, 'Expertise and cross-media growth strategies. Seamlessly visualize quality intellectual capital without superior.', 2.688, '85%'],
];

function timelineRow(s, step, y) {
  frame(s, [3.254, y, 0.68, 0.68], step[1], 3);
  txt(s, [3.344, y + 0.205, 0.5, 0.269], step[0], { size: 16, bold: true, align: 'center' });
  txt(s, [1.019, y + 0.205, 1.8, 0.269], 'Title Goes Here', { size: 16, bold: true, align: 'right' });
  txt(s, [4.369, y, 4.218, 0.474], step[2], { lh: 1.5 });
  txt(s, [4.369, y + 0.6, 1.848, 0.185], 'Team Work', { size: 11 });
  line(s, [4.391, y + 0.889, step[3], 0], step[1], 10);
  txt(s, [4.391 + step[3] + 0.193, y + 0.805, 0.406, 0.177], step[4], { size: 10.5 });
}

function slide10(s) {                                   // Our Travel Planning (steps 1-3)
  txt(s, [1.174, 1.103, 4.859, 0.431], 'Our Travel Planning', { font: SERIF, size: 28, bold: true, align: 'center', lhPt: 30, mid: true });
  [[0, 0.899], [1.739, 0.88], [3.299, 0.79], [4.769, 0.79], [6.239, 1.261]].forEach(function (seg) {
    line(s, [3.594, seg[0], 0, seg[1]], GREY, 3);
  });
  [2.619, 4.089, 5.559].forEach(function (y, i) { timelineRow(s, TIMELINE_STEPS[i], y); });
}

function slide11(s) {                                   // End Schedule (steps 4-6)
  [[-0.014, 0.88], [1.545, 0.79], [3.015, 0.79], [4.485, 1.261], [6.601, 0.899]].forEach(function (seg) {
    line(s, [3.594, seg[0], 0, seg[1]], GREY, 3);
  });
  [0.88, 2.35, 3.82].forEach(function (y, i) { timelineRow(s, TIMELINE_STEPS[i + 3], y); });
  txt(s, [1.425, 5.966, 4.358, 0.431], 'End Schedule', { font: SERIF, size: 28, bold: true, align: 'center', lhPt: 30, mid: true });
}

function slide12(s) {                                   // Discover The Difference
  www(s, [4.583, 0.874, 2.6, 0.135], { size: 8, align: 'center' });
  txt(s, [3.95, 5.604, 3.423, 1.022], ['\u201cAnd I knew exactly how old Walt Disney\u2019s Cinderella felt when she found her prince.\u201d ', '\u2014Elizabeth Young'], { size: 14, italic: true, lh: 1.5 });
  txt(s, [7.918, 6.155, 4.574, 0.471], 'Discover The Difference', { font: SERIF, size: 28, bold: true, align: 'right', mid: true });
}

function slide13(s) {                                   // Journey Beyond Boundaries
  txt(s, [7.081, 1.256, 5.253, 0.421], 'Journey Beyond Boundaries', { font: SERIF, size: 28, bold: true, lhPt: 30, mid: true });
  txt(s, [9.899, 3.242, 2.438, 2.497], ['Interactively procrastinate for high-payoff content without backward compatible data. Quickly cultivate optimal processes and tactical architectures.', '', 'Quickly disseminate superior deliverables whereas web-enabled applications. Quickly drive clicks and mortar catalysts for change. Continually reintermediate integrated processes.'], { align: 'justify', lh: 1.5 });
  www(s, [1, 6.098, 2.6, 0.151]);
}

function slide14(s) {                                   // Eco-Friendly Travel Solutions
  www(s, [9.422, 1.07, 2.6, 0.151], { align: 'right' });
  txt(s, [1.074, 5.739, 3.499, 0.942], 'Eco-Friendly Travel Solutions', { font: SERIF, size: 28, bold: true, mid: true });
  txt(s, [5.16, 5.649, 3, 0.983], 'Interactively coordinate proactive e-commerce via process-centric outside the box  thinking pursue scalable customer service through. Collaboratively administrate empowered.', { lh: 1.5 });
  txt(s, [9.022, 5.649, 3, 0.983], 'Proactively envisioned multimedia based expertise and cross-media growth strategies. Seamlessly visualize quality intellectual capital without superior collaboration agile provide.', { lh: 1.5 });
}

function slide15(s) {                                   // Event Execution And Management
  const BLURB = 'Proactively visioned multimedia base expertise and cross media growth strategies visualize for  intellectual capital superior good collaboration for customers.';
  txt(s, [9.257, 3.561, 3.212, 2.423], ['Event Execution', 'And Management'], { font: SERIF, size: 36, bold: true, mid: true });
  [5.172, 9.257].forEach(function (x) {
    txt(s, [x, 1.071, 3.014, 0.236], 'Your Title Goes Here', { size: 14, bold: true });
    txt(s, [x, 1.376, 3.014, 0.657], BLURB, { size: 9, align: 'justify', lh: 1.5 });
  });
  www(s, [1.071, 6.338, 2.5, 0.151]);
}

function slide16(s) {                                   // Break Section
  box(s, [0.624, 1.058, 3.206, 4.002], PEACH);
  box(s, [9.812, 4.656, 2.617, 2.494], PEACH);
  txt(s, [5.997, 1.32, 6.432, 1.111], 'Break Section', { font: 'Bitter', size: 66, mid: true });
  txt(s, [1.029, 6.065, 6.072, 0.43], 'Proactively envisioned multimedia based expertise and cross-media growth strategies. Seamlessly visualize quality intellectual capital without superior collaboration. Leverage agile frameworks to provide a robust synopsis.', { size: 9, lh: 1.5 });
  www(s, [5.997, 4.521, 1.89, 0.135], { size: 8 });
}

function slide17(s) {                                   // Our Team Profile - skill bars
  txt(s, [0.844, 1.839, 1.929, 2.221], ['Our', 'Team', 'Profile'], { font: SERIF, size: 44, bold: true, mid: true });
  txt(s, [8.011, 1.94, 3.937, 0.404], 'CLARRA SIMMONS', { font: SERIF, size: 24, bold: true });
  txt(s, [8.011, 2.371, 3.937, 0.269], 'Your Jobdesk Here', { size: 16 });
  txt(s, [8.011, 3.056, 4.51, 0.611], 'Interactively coordinate proactive e-commerce via process-centric for outside the box thinking. Completely pursue scalable administrate empowered markets via plug-and-play networks. Standards in web-readiness energistically scale.', { size: 9, align: 'justify', lhPt: 15 });
  [['Branding', 3.939, 3.782, PEACH, '100%', 11.961],
   ['Public Relations', 4.359, 3.148, RED, '80%', 11.327],
   ['Graphic Design', 4.778, 2.303, PEACH, '60%', 10.482],
   ['Social Marketing', 5.206, 3.614, RED, '95%', 11.793]].forEach(function (b) {
    txt(s, [8.011, b[1], 2.222, 0.185], b[0], { size: 11 });
    line(s, [8.011, b[1] + 0.271, b[2], 0], b[3], 5);
    txt(s, [b[5], b[1] + 0.178, 0.383, 0.185], b[4], { size: 11 });
  });
  www(s, [0.844, 5.425, 1.89, 0.135], { size: 8 });
}

function slide18(s) {                                   // Our Team - three colour blocks
  [[3.4, 0, RED, 'Wilbert Alvarez', 'General Manager', WHITE],
   [0, 2.5, PEACH, 'Tricia Rodriguez', 'Accounting', BLACK],
   [3.4, 5, TEAL, 'Andrew Newton', 'Techinal', WHITE]].forEach(function (c) {
    const x = c[0], y = c[1], fg = c[5];
    box(s, [x, y, 3.4, 2.5], c[2]);
    txt(s, [x + 0.521, y + 0.588, 2.36, 0.269], c[3], { size: 16, bold: true, color: fg, align: 'center' });
    txt(s, [x + 0.519, y + 0.857, 2.36, 0.202], c[4], { size: 12, color: fg, align: 'center' });
    txt(s, [x + 0.521, y + 1.186, 2.358, 0.726], LOREM_F, { color: fg, align: 'center', lh: 1.5 });
  });
  txt(s, [8.208, 1.776, 3.722, 0.431], 'Our Team', { font: SERIF, size: 28, bold: true, align: 'right', lhPt: 30, mid: true });
  txt(s, [8.208, 2.643, 3.722, 2.493], ['Synergistically evolve 2.0 technologies rather than just in time initiatives. Quickly deploy strategic networks with compelling  e-business. Credibly pontificate highly efficient manufactured products and enabled data.', '', 'Objectively integrate emerging core competencies before process-centric communities. Dramatically evisculate holistic innovation rather than client-centric data. Progressively maintain extensive objectively pursue diverse catalysts for change content in provide a robust strategies. '], { align: 'right', lh: 1.5 });
  www(s, [9.229, 5.572, 2.7, 0.151], { align: 'right' });
}

function slide19(s) {                                   // Pricing List Table Plan
  txt(s, [2.956, 0.772, 7.421, 0.431], 'Pricing List Table Plan', { font: SERIF, size: 28, bold: true, align: 'center', lhPt: 30, mid: true });
  txt(s, [2.956, 1.294, 7.421, 0.478], LOREM_A, { align: 'center', lh: 1.5 });
  [[1.481, PEACH, '$180.00', 'Bronze Package', BLACK],
   [5.111, RED, '$240.00', 'Gold Package', WHITE],
   [8.741, TEAL, '$125.00', 'Silver Package', WHITE]].forEach(function (c) {
    const x = c[0], fg = c[4];
    box(s, [x, 2.543, 3.111, 4], c[1]);
    box(s, [x, 3.623, 3.111, 0.05], WHITE);
    txt(s, [x + 0.475, 2.814, 2.16, 0.539], c[2], { size: 32, bold: true, color: fg, align: 'center', mid: true });
    txt(s, [x + 0.475, 4.141, 2.16, 1.231], LOREM_C, { color: fg, align: 'center', lh: 1.5 });
    txt(s, [x + 0.475, 5.84, 2.16, 0.236], c[3], { size: 14, bold: true, color: fg, align: 'center', mid: true });
  });
  www(s, [5.367, 6.946, 2.6, 0.151], { align: 'center' });
}

// S.W.O.T funnel: each tier is a flat hexagon with two bevelled sides and a
// front band, stacked with a small overlap.
const HEX_TOP = [[0.75, 0], [1, 0.5], [0.75, 1], [0.25, 1], [0, 0.5], [0.25, 0]];
const HEX_LEFT = [[0, 0], [0, 0.404], [1, 1], [1, 0.596]];
const HEX_RIGHT = [[1, 0], [1, 0.404], [0, 1], [0, 0.596]];
const TIER_LIGHT = { top: 'FFE5D7', left: 'FD7D37', right: 'FED9C4', front: PEACH };
const TIER_DARK = { top: 'E9BAAD', left: '953E26', right: 'DD9785', front: RED };

function funnelTier(s, y, label, tone) {
  poly(s, [1.684, y, 2.7, 1.023], HEX_TOP, tone.top);
  poly(s, [3.709, y + 0.511, 0.675, 0.858], HEX_RIGHT, tone.right);
  poly(s, [1.684, y + 0.511, 0.675, 0.858], HEX_LEFT, tone.left);
  box(s, [2.36, y + 1.021, 1.349, 0.347], tone.front);
  txt(s, [2.289, y + 0.627, 1.491, 0.269], label, { size: 16, bold: true, align: 'center' });
}

function slide20(s) {                                   // S.W.O.T Analysis Slide
  txt(s, [1.684, 0.833, 6.203, 0.433], 'S.W.O.T Analysis Slide', { font: SERIF, size: 28, bold: true, lhPt: 30, mid: true });
  txt(s, [1.684, 1.409, 6.203, 0.478], 'Bring to the table win-win survival strategies to ensure proactive domination. At the end of the day, going forward, a new normal that has evolved from generation X is on the runway heading.', { lh: 1.5 });
  [['Threats', 5.281, TIER_DARK], ['Opportunities', 4.43, TIER_LIGHT],
   ['Weaknesses', 3.578, TIER_DARK], ['Strengths', 2.727, TIER_LIGHT]].forEach(function (t) {
    funnelTier(s, t[1], t[0], t[2]);
  });
  [['S', 'Strengths', PEACH, 2.839], ['W', 'Weaknesses', RED, 3.844],
   ['O', 'Opportunities', PEACH, 4.849], ['T', 'Threats', RED, 5.855]].forEach(function (r) {
    const y = r[3];
    txt(s, [5.824, y, 0.375, 0.536], r[0], { size: 32, bold: true, color: r[2] });
    txt(s, [6.565, y, 5.084, 0.269], r[1] + ' Analysis', { size: 16, bold: true });
    txt(s, [6.565, y + 0.319, 5.084, 0.478], LOREM_B, { lh: 1.5 });
    line(s, [4.75, y + 0.399, 0.709, 0], r[2], 1);
  });
  www(s, [8.489, 1.288, 3.16, 0.151], { align: 'right' });
}

function slide21(s) {                                   // Growth With Additional Funding
  txt(s, [1.108, 1.33, 6.478, 0.432], 'Growth With Additional Funding', { font: SERIF, size: 28, bold: true, lhPt: 30, mid: true });
  // Waffle chart: 3 year-groups x 3 series, 10 cells per bar filled from bottom.
  const FILLED = [[7, 3, 9], [3, 10, 6], [9, 7, 5]];
  const SERIES = [RED, PEACH, TEAL];
  ['2026', '2027', '2028'].forEach(function (year, g) {
    const gx = 1.107 + g * 2.2465;
    FILLED[g].forEach(function (n, c) {
      for (let row = 0; row < 10; row++) {
        box(s, [gx + c * 0.6885, 2.195 + row * 0.3465, 0.6, 0.26],
          row >= 10 - n ? SERIES[c] : GREY);
      }
    });
    txt(s, [gx + 0.41, 5.766, 1.158, 0.404], ['Data In', year], { size: 12, bold: true, align: 'center' });
  });
  txt(s, [8.693, 1.958, 3.539, 0.236], 'Your Best Statement Here', { size: 14, bold: true });
  txt(s, [8.693, 2.436, 3.539, 0.983], 'Collaboration provide a robust visualize quality intellectual capital without superior  and idea sharing. Proactively envisioned multimedia based expertise and cross media growth strategies seamlessly visualize quality.', { lh: 1.5 });
  [[TEAL, 3.871, 'Proactively envisioned multimedia based expertise and cross media growth quality.'],
   [RED, 4.499, 'Coordinate proactive via outside the box pursue scalable service intellectual capital.'],
   [PEACH, 5.126, 'Leverage agile frameworks to provide robust synopsis for high level views new.']].forEach(function (k) {
    box(s, [8.693, k[1], 0.237, 0.237], k[0]);
    txt(s, [9.2, k[1] - 0.062, 2.551, 0.478], k[2], { lh: 1.5 });
  });
}

// Exploded 3-D pie. Angles are degrees clockwise from 3 o'clock. A copy of the
// slice in the darker "side" colour, offset downwards by the extrusion depth,
// shows below the lighter top face and reads as the cylinder wall.
const PIE_RX = 1.966, PIE_RY = 0.955, PIE_DEPTH = 0.596;

function pieSlice(s, cx, cy, a0, a1, top, side) {
  const pts = [[0, 0]];                       // apex, then the arc
  for (let a = a0; a <= a1 + 0.01; a += 6) {
    const t = a * Math.PI / 180;
    pts.push([Math.cos(t) * PIE_RX, Math.sin(t) * PIE_RY]);
  }
  const xs = pts.map(function (p) { return p[0]; }), ys = pts.map(function (p) { return p[1]; });
  const x0 = Math.min.apply(null, xs), y0 = Math.min.apply(null, ys);
  const w = Math.max.apply(null, xs) - x0, h = Math.max.apply(null, ys) - y0;
  const frac = pts.map(function (p) { return [(p[0] - x0) / w, (p[1] - y0) / h]; });
  poly(s, [cx + x0, cy + y0 + PIE_DEPTH, w, h], frac, side);
  poly(s, [cx + x0, cy + y0, w, h], frac, top);
}

function slide22(s) {                                   // Infographic Slide
  txt(s, [2.421, 0.823, 8.49, 0.421], 'Infographic Slide', { font: SERIF, size: 28, bold: true, align: 'center', lhPt: 30, mid: true });
  txt(s, [2.421, 1.385, 8.49, 0.478], 'Bring the table win-win survival strategies ensure proactive dominan. At the end of the day, going forward, for new normal that has evolved from runway heading to our solution administrate. Seamlessly visualize quality intellectual capital without superior frameworks.', { align: 'center', lh: 1.5 });
  pieSlice(s, 9.387, 3.641, 130, 432, 'DD9785', RED);    // 302 deg major slice
  pieSlice(s, 8.899, 4.496, 72, 130, 'A9D0CA', TEAL);    // exploded 58 deg slice
  [[TEAL, 3.021, 'aperture', 3.351, '2026 Income', '+ $415.000', 3, 3.669],
   [RED, 4.787, 'swap', 5.117, '2027 Income', '+ $735.000', 4.766, 5.434]].forEach(function (k) {
    ellipse(s, [0.845, k[1], 1.003, 1.003], k[0]);
    glyph(s, 1.347, k[1] + 0.501, 0.446, k[2], WHITE, k[0]);
    txt(s, [2.4, k[6], 4.195, 0.269], 'Your Title Goes Here', { size: 16, bold: true });
    txt(s, [2.4, k[3], 1.8, 0.236], k[4], { size: 14, bold: true });
    txt(s, [4.794, k[3], 1.8, 0.236], k[5], { size: 14, bold: true, align: 'right' });
    txt(s, [2.4, k[7], 4.195, 0.478], LOREM_G, { lh: 1.5 });
  });
  [[RED, 5.078, 4.988], [TEAL, 6.002, 5.912]].forEach(function (k) {   // pie legend
    box(s, [9.962, k[1], 0.242, 0.242], k[0]);
    txt(s, [10.429, k[2], 2.059, 0.657], 'Proactively envisioned multimedia based expertise and cross media growth quality intellectual.', { size: 9, lh: 1.5 });
  });
}

function slide23(s) {                                   // Resource Allocation - three column charts
  txt(s, [0.952, 0.809, 5.626, 0.561], 'Resource Allocation', { font: SERIF, size: 36, bold: true, lhPt: 40, mid: true });
  txt(s, [1.101, 2.178, 11.163, 0.658], 'Synergistically evolve on technologies rather than just the into makes times pontificate manufactured product objectively integrate competency making the process-centric communities. Credibly manufactured integrated emerging process-centric the communities. Dramatically client-centric making company pontificate pursue. Bring the table win-win survival strategies to ensure the proactive more them dominant. At the end of the day, going onto forward, for new. After installed based to benefits for dramatic company galley robust media frameworks making telemedicine.', { size: 9, align: 'justify', lh: 1.5 });
  // Bar heights are taken from the source art (0 .. 1000 axis, 1.663" tall).
  const BARS = [
    [RED, 0, [[610, 0.996], [430, 0.736], [780, 1.202], [270, 0.477], [940, 1.472]]],
    [PEACH, 4.164, [[780, 1.202], [430, 0.736], [610, 0.996], [940, 1.472], [270, 0.477]]],
    [TEAL, 8.359, [[610, 0.996], [430, 0.736], [780, 1.202], [270, 0.477], [940, 1.472]]],
  ];
  BARS.forEach(function (chart) {
    const dx = chart[1];
    txt(s, [1.101 + dx, 3.645, 2.835, 0.236], 'Text Title Goes Here', { size: 14 });
    line(s, [1.45 + dx, 4.209, 0, 1.663], BLACK, 1);          // y axis
    line(s, [1.4 + dx, 5.833, 2.535, 0], BLACK, 1);           // x axis
    ['1000', '800', '600', '400', '200', '0'].forEach(function (t, i) {
      txt(s, [1.07 + dx, 4.209 + i * 0.314, 0.262, 0.118], t, { size: 7, align: 'right' });
    });
    chart[2].forEach(function (b, i) {
      const x = 1.76 + dx + i * 0.3955;
      box(s, [x, 5.833 - b[1], 0.328, b[1]], chart[0]);
      txt(s, [x + 0.07, 5.833 - b[1] - 0.167, 0.189, 0.118], String(b[0]), { size: 7, align: 'center' });
    });
    txt(s, [1.101 + dx, 6.261, 2.835, 0.43], LOREM_H, { size: 9, align: 'justify', lh: 1.5 });
  });
  www(s, [10.119, 1.022, 2.145, 0.135], { size: 8, align: 'right' });
}

function slide24(s) {                                   // Product Mockup Slide (phone photo)
  photo(s, [0.957, -0.971, 4.753, 6.667]);
  txt(s, [7.772, 2.248, 4.455, 0.673], '68%', { font: SERIF, size: 40, bold: true });
  txt(s, [7.772, 3.044, 4.455, 0.539], 'Smartphone user who download branded apps buy 67,9% of the branded products', { size: 16, bold: true });
  txt(s, [7.772, 3.913, 4.455, 1.339], ['Proactively envisioned multimedia based expertise and cross-media growth strategies. Seamlessly visualize quality intellectual capital.', '', 'Interactively coordinate proactive e-commerce via process-centric "outside the box" thinking. Completely pursue scalable customer service through sustainable potentialities administrate turnkey channels whereas virtual.'], { size: 9, lh: 1.5 });
  txt(s, [1.106, 5.91, 4.532, 0.433], 'Product Mockup Slide', { font: SERIF, size: 28, bold: true, lhPt: 30, mid: true });
  txt(s, [8.028, 6.051, 4.2, 0.151], 'W   W   W   .   U   W   O   L   A   .   C   O   M', { size: 9, align: 'right' });
}

function slide25(s) {                                   // Product Mockup Slide (monitor photo)
  photo(s, [6.8, 0.785, 5.38, 4.346]);
  tagline(s, [1.153, 0.785, 1.918, 0.471]);
  txt(s, [1.153, 2.116, 4.466, 0.421], 'Product Mockup Slide', { font: SERIF, size: 28, bold: true, lhPt: 30, mid: true });
  txt(s, [1.153, 2.898, 4.466, 1.339], ['Synergistically evolve 2.0 technologies rather than just in time initiatives. Quickly deploy strategic networks with compelling. Credibly pontificate highly efficient.', '', 'Objectively integrate emerging core competencies before process-centric dramatic evisculate holistic innovation rather than client-centric data. Progressively maintain extensive objectively pursue diverse special items for you.'], { size: 9, align: 'justify', lh: 1.5 });
  [[1.153, 'cart', PEACH, 1.96], [4.943, 'truck', RED, 5.751]].forEach(function (k) {
    glyph(s, k[0] + 0.24, 6.24, 0.46, k[1], k[2]);
    txt(s, [k[3], 5.894, 2.601, 0.236], 'Your Subtitle Gies Here', { size: 14, bold: true });
    txt(s, [k[3], 6.171, 2.601, 0.43], 'Leverage agile frameworks to provide a robust synopsis for high level views diverse catalys.', { size: 9, lh: 1.5 });
  });
  www(s, [9.452, 6.172, 2.728, 0.151], { align: 'right' });
}

// Germany map, simplified from the original artwork: 12 state outlines, two of
// them highlighted. Each entry is [colour, "x,y x,y ..."] in slide inches.
const MGREY = 'A6A6A6';   // unhighlighted states
const MAP = [
  [PEACH, '4.59,4.33 4.67,4.37 4.68,4.51 4.97,4.37 4.94,4.15 5.11,4.03 5,3.71 5.27,3.61 4.99,3.44 4.72,3.4 4.53,3.17 4.32,3.17 4.29,3.44 4.3,3.34 4.22,3.3 4.19,3.41 4.11,3.25 3.84,3.31 3.83,3.44 3.93,3.47 3.83,3.81 3.71,3.9 3.8,3.91 3.82,4.02 4.03,3.91 4.12,3.98 4.12,4.13 4.28,4.06 4.22,3.91 4.48,3.89 4.45,4.08 4.6,4.22 4.57,4.33 4.45,3.58 4.35,3.6 4.28,3.51 4.45,3.58'],
  [RED, '4.51,5.36 4.56,5.26 4.7,5.29 4.92,5.69 4.93,5.8 4.77,5.95 4.81,6.29 4.63,6.35 4.75,6.34 4.82,6.47 4.9,6.33 5.16,6.41 5.48,6.25 5.73,6.37 5.64,6.08 5.85,5.95 5.87,5.85 5.95,5.88 5.98,5.76 5.61,5.45 5.52,5.29 5.57,5.21 5.38,4.96 5.2,4.91 5.17,5.03 5.02,4.97 5,5.06 4.8,4.89 4.61,5.11 4.46,5.12 4.51,5.36'],
  [MGREY, '4.18,4.81 4.35,4.61 4.34,4.49 4.58,4.34 4.59,4.22 4.45,4.08 4.48,3.9 4.23,3.92 4.29,4.07 4.12,4.13 4.12,3.99 4,3.91 3.71,4.12 3.73,4.2 3.47,4.25 3.56,4.47 3.45,4.65 3.62,5 3.74,4.98 4.06,4.69 4.18,4.81'],
  [MGREY, '6.17,3.48 6.08,3.49 6.12,3.4 5.98,3.34 5.73,3.54 5.44,3.44 5.18,3.57 5.48,3.7 5.48,4.06 5.8,4.22 5.79,4.46 6.03,4.48 6.28,4.35 6.24,3.85 6.08,3.71 6.18,3.49 5.94,3.97 5.75,3.94 5.77,3.84 5.88,3.8 5.97,3.92'],
  [MGREY, '4.78,4.91 4.75,4.7 4.84,4.57 4.73,4.46 4.62,4.48 4.66,4.37 4.6,4.33 4.32,4.52 4.37,4.58 4.18,4.77 4.17,5.03 4.06,5.14 4.21,5.14 4.26,5.37 4.35,5.35 4.4,5.46 4.48,5.4 4.45,5.11 4.61,5.11 4.78,4.91'],
  [MGREY, '4.69,3.3 4.83,3.23 4.83,3.34 4.96,3.42 5.07,3.32 5,3.06 5.12,2.88 5,2.93 4.87,2.85 4.81,2.92 4.83,2.84 4.71,2.83 4.75,2.67 4.34,2.6 4.46,2.81 4.33,2.91 4.44,2.92 4.4,3.07 4.69,3.3'],
  [PEACH, '5,3.43 5.21,3.56 5.44,3.43 5.72,3.53 5.97,3.33 6.13,3.4 6.08,3.48 6.17,3.45 6.12,3.23 5.99,3.17 5.98,3.04 5.87,3.04 5.73,2.86 5.52,2.93 5.45,3.04 5.31,3.01 5.23,3.15 5.08,3.1 5,3.43'],
  [RED, '4.64,6.32 4.81,6.28 4.76,5.95 4.92,5.8 4.81,5.41 4.6,5.27 4.4,5.46 4.34,5.35 4.27,5.37 4.29,5.51 4.07,5.84 3.98,6.27 4.02,6.34 4.31,6.32 4.25,6.28 4.32,6.22 4.64,6.32'],
  [MGREY, '5.37,4.95 5.61,4.71 5.56,4.63 5.47,4.7 5.22,4.62 5.2,4.48 5.02,4.34 4.75,4.48 4.85,4.57 4.73,4.84 4.97,5.05 5.03,4.96 5.16,5.02 5.2,4.9 5.37,4.95'],
  [MGREY, '5.27,3.62 5.01,3.71 5.12,4.03 4.95,4.15 4.96,4.26 5.08,4.45 5.21,4.48 5.23,4.61 5.47,4.69 5.47,4.39 5.74,4.34 5.79,4.23 5.52,4.13 5.47,3.7 5.27,3.62'],
  [MGREY, '6.25,4.39 5.97,4.49 5.78,4.47 5.69,4.32 5.48,4.38 5.46,4.57 5.62,4.71 5.38,4.95 5.5,5.08 5.57,4.96 6.15,4.72 6.14,4.64 6.29,4.75 6.37,4.52 6.25,4.39'],
  [MGREY, '4.25,5.34 4.2,5.15 4.05,5.14 4.18,4.87 4.08,4.7 3.75,4.99 3.54,5.06 3.54,5.18 3.65,5.26 3.61,5.39 3.9,5.37 3.91,5.55 4.01,5.63 4.21,5.66 4.25,5.34'],];

function slide26(s) {                                   // Our Location - map + dark panel
  box(s, [7.474, 0.399, 6.188, 7.5], PANEL);   // 60%-opaque dark scrim over white
  www(s, [9.639, 1.422, 2.7, 0.151], { color: WHITE, align: 'right' });
  txt(s, [8.14, 3.008, 4.198, 0.269], 'Our Income Is Based On Area', { size: 16, bold: true, color: WHITE });
  txt(s, [8.14, 3.429, 4.198, 0.726], 'Proactively envisioned multimedia based expertise and cross-media growth strategies. Seamlessly visualize quality intellectual capital without superior collaboration globally incubate standards.', { color: WHITE, lh: 1.5 });
  [[8.14, RED, 'clipboard', '75% Sales', 4.958, 5.254], [10.241, PEACH, 'check', '82% Sales', 4.963, 5.26]].forEach(function (k) {
    glyph(s, k[0] + 0.2, 4.679, 0.44, k[2], k[1], PANEL);
    txt(s, [k[0], k[4], 1.923, 0.236], k[3], { size: 14, bold: true, color: WHITE });
    txt(s, [k[0], k[5], 1.923, 0.726], LOREM_E + 'uniquely matrix sound.', { color: WHITE, lh: 1.5 });
  });
  txt(s, [0.779, 0.825, 5.588, 0.457], 'Our Location', { font: SERIF, size: 28, bold: true, lhPt: 35, mid: true });
  txt(s, [0.779, 1.39, 5.588, 0.474], 'Bring the table win-win survival strategies ensure proactive dominan. At the end of the day, going forward, for new normal that has evolved from runway heading.', { lh: 1.5 });
  MAP.forEach(function (region) {
    const pts = region[1].split(' ').map(function (p) { return p.split(',').map(Number); });
    const xs = pts.map(function (p) { return p[0]; }), ys = pts.map(function (p) { return p[1]; });
    const x0 = Math.min.apply(null, xs), y0 = Math.min.apply(null, ys);
    const w = Math.max.apply(null, xs) - x0, h = Math.max.apply(null, ys) - y0;
    poly(s, [x0, y0, w, h], pts.map(function (p) { return [(p[0] - x0) / w, (p[1] - y0) / h]; }), region[0]);
  });
  [3.247, 5.441].forEach(function (y) {
    txt(s, [0.779, y, 2.063, 0.236], 'Your Text Here', { size: 14, bold: true, align: 'right' });
    txt(s, [0.779, y + 0.351, 2.063, 0.474], 'Make a type specimen book amet unknown printer took.', { align: 'right', lh: 1.5 });
  });
}

function slide27(s) {                                   // Our Location - two address cards
  [[1.033, TEAL], [4.338, RED]].forEach(function (c) {
    const x = c[0];
    box(s, [x, 3.174, 2.929, 3.483], c[1]);
    txt(s, [x + 0.31, 5.051, 2.309, 0.236], 'Travel Name Goes Here', { size: 14, bold: true, color: WHITE });
    txt(s, [x + 0.31, 5.439, 2.309, 0.404], ['At W Shunk St, Wtown,', 'PA 19145, USA'], { size: 12, color: WHITE });
    txt(s, [x + 0.31, 5.945, 2.309, 0.478], 'Credibly pontificate highly efficient manual structured products.', { color: WHITE, lh: 1.5 });
  });
  txt(s, [8.61, 5.109, 3.379, 0.431], 'Our Location', { font: SERIF, size: 28, bold: true, lhPt: 30, mid: true });
  txt(s, [8.61, 5.672, 3.38, 0.979], LOREM_A, { lh: 1.5 });
}

function slide28(s) {                                   // Travel Gallery Image
  txt(s, [2.956, 0.577, 7.421, 0.431], 'Travel Gallery Image', { font: SERIF, size: 28, bold: true, align: 'center', lhPt: 30, mid: true });
  txt(s, [2.956, 1.109, 7.421, 0.474], LOREM_A, { align: 'center', lh: 1.5 });
}

function slide29(s) {                                   // What Our Client Says
  txt(s, [0.707, 0.647, 7.739, 0.432], 'What Our Client Says', { font: SERIF, size: 28, bold: true, lhPt: 30, mid: true });
  txt(s, [0.707, 1.266, 7.739, 0.474], 'Interactively coordinate proactive e-commerce via process-centric outside the box  thinking pursue scalable customer service through. Collaboratively administrate empowered markets via plug-and-play networks dynamic procrastinate.', { lh: 1.5 });
  [[0.707, RED, 'Nadera Gileradin', WHITE], [4.767, TEAL, 'Fernando Jolion', WHITE],
   [8.826, PEACH, 'jennie Marker', BLACK]].forEach(function (c) {
    const x = c[0], fg = c[3];
    box(s, [x, 5.304, 3.8, 1.486], c[1]);
    txt(s, [x + 0.497, 5.661, 2.806, 0.269], c[2], { size: 16, bold: true, color: fg });
    txt(s, [x + 0.497, 5.959, 2.806, 0.474], LOREM_E + 'views new normal that.', { color: fg, lh: 1.5 });
  });
}

function slide30(s) {                                   // Thanks For Watching
  box(s, [0, 0, 5.802, 7.5], PEACH);
  www(s, [-0.09, 5.14, 2.6, 0.151], { rot: 270 });
  txt(s, [1.026, 0.984, 4.022, 1.346], ['Thanks', 'For Watching'], { font: SERIF, size: 44, bold: true, lhPt: 48, mid: true });
  txt(s, [6.781, 5.533, 3.219, 0.983], 'Interactively coordinate proactive e-commerce process centric "outside the box" thinking. Completely pursue scalable customer through sustainable potentialities collaboratively great administrate.', { align: 'justify', lh: 1.5 });
  tagline(s, [10.852, 6.112, 1.497, 0.404], { size: 12, align: 'right' });
}

// ------------------------------------------------------- icon sheets 31-42 --
// Twelve reference sheets, each a 10-column grid of small dark pictograms.
// The originals are hundreds of unique free-form glyphs; here each cell picks
// one of eight simple marks whose visual weight matches the source. Letters:
//   o ring   s square outline   r rounded box   t triangle
//   p pin    b filled block     d disc          h heavy square
const SHEETS = [
  ['tpddrpdrbt', 'ptdprrbbpr', 'dtbrrttptt', 'tdpdpttbrr', 'rbtbsrtpdr', 'dtpttbdddd', 'rtdtdhdhdr'],
  ['rrpssrrtrr', 'rrrrrtrrrr', 'ddtttpttbd', 'rrtprbbrbt', 'rpddbdhrrt', 'brrrtrpdpp', 'dhbprsbdrb'],
  ['trbrrbbpdr', 'btbddhpsbr', 'trtttdrttt', 'stdsstrttt', 'pprrtddpbt', 'tbbbprrtpr', 'rppbrdbttb'],
  ['rbtbtbbpts', 'pttpppbtdr', 'rbprrrsbdr', 'bdddhbrtpt', 'ttrprrbtpt', 'dpttttttbd', 'rsbrrrrdtd'],
  ['ttrtrrrrpr', 'drbttptdhd', 'ttrrrtttrt', 'brrrtttttt', 'pptrttprrd', 'pbtbtrrbrt', 'drrptttpbs'],
  ['btthppptrd', 'bdbtdrhprs', 'prrrrrrbtt', 'ptphtpdtrr', 'tbdrrttttt', 'btdpbtbbpr', 'rbbpbdpbtp'],
  ['bddtbbbdbb', 'phhbttstrb', 'dbtprttdhd', 'dbtrrbrtph', 'dtdpptbrbr', 'bbprtttdhp', 'rdbptrpttr'],
  ['pppprphdrp', 'bpbdbrtptt', 'ttpbrbdhtt', 'pdttppbprb', 'pptrrtprpb', 'tttrtttrtr', 'ddtbdrtdbt'],
  ['trrrrbppbt', 'rbrpttttdr', 'pbttbppbbt', 'tttptbbbpr', 'rrrrrpdrtd', 'rrrdtrdrrr', 'pttrtsrrrb', 'ddsrbbprpd'],
  ['dttrrrttpr', 'rprdtrrppt', 'tdprprptrr', 'ptrptprspt', 'prrtppdhtp', 'rbpbttrrtd', 'rpppptrrsd'],
  ['hprrptbpbb', 'bdbsstbtbd', 'dtttrprrtt', 'tbhrdrtttt', 'prrrttrtpt', 'ppbbttpprp', 'rrhttrrrrt'],
  ['dsrsbrrrrt', 'rdhddhpbtd', 'hdhhddpbhd', 'pbttpttppt', 'trrrdptdtt', 'pdrtpbpdpp', 'trsrr...bb'],
];

// Marks are listed lightest to heaviest; the sheet strings pick the one whose
// ink weight matches the glyph it replaces.
const MARKS = {
  o: function (s, x, y, d) { frame(s, [x - d * 0.42, y - d * 0.42, d * 0.84, d * 0.84], INK, 2.5); },
  s: function (s, x, y, d) { ring(s, x, y, d * 0.92, d * 0.11, INK, WHITE); },
  r: function (s, x, y, d) {
    s.addShape('roundRect', { x: x - d / 2, y: y - d * 0.4, w: d, h: d * 0.8,
      rectRadius: d * 0.18, fill: { color: INK }, line: { type: 'none' } });
    box(s, [x - d * 0.34, y - d * 0.24, d * 0.68, d * 0.48], WHITE);
  },
  t: function (s, x, y, d) { poly(s, [x - d / 2, y - d * 0.44, d, d * 0.88], [[0.5, 0], [1, 1], [0, 1]], INK); },
  p: function (s, x, y, d) { GLYPHS.pin(s, x, y, d, INK, WHITE); },
  b: function (s, x, y, d) { ellipse(s, [x - d * 0.44, y - d * 0.44, d * 0.88, d * 0.88], INK); },
  d: function (s, x, y, d) {
    s.addShape('roundRect', { x: x - d / 2, y: y - d * 0.42, w: d, h: d * 0.84,
      rectRadius: d * 0.2, fill: { color: INK }, line: { type: 'none' } });
    box(s, [x - d * 0.16, y - d * 0.26, d * 0.32, d * 0.52], WHITE);
  },
  h: function (s, x, y, d) { box(s, [x - d * 0.46, y - d * 0.44, d * 0.92, d * 0.88], INK); },
};

function iconSheet(s, rows) {
  const x0 = 1.372, dx = 1.1765, d = 0.43;
  const y0 = rows.length === 8 ? 0.879 : 0.893;
  const dy = rows.length === 8 ? 0.814 : 0.952;
  rows.forEach(function (row, r) {
    row.split('').forEach(function (ch, c) {
      if (MARKS[ch]) MARKS[ch](s, x0 + c * dx, y0 + r * dy, d);
    });
  });
}

// -------------------------------------------------------------- assembly --

const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07,
  slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15,
  slide16, slide17, slide18, slide19, slide20, slide21, slide22, slide23,
  slide24, slide25, slide26, slide27, slide28, slide29, slide30];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'KUNUA', width: 13.333, height: 7.5 });
  pptx.layout = 'KUNUA';
  pptx.title = 'Kunua Travel Presentation Template';

  BUILDERS.forEach(function (fn) {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    fn(s);
  });
  SHEETS.forEach(function (rows) {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    iconSheet(s, rows);
  });

  return pptx.writeFile({
    fileName: path.join(__dirname, '0cb31026-d3cd-4ffb-a13f-18a3573e4bb5_grok_final.pptx'),
  });
}

build().then(function (f) { console.log('wrote', f); },
  function (e) { console.error(e); process.exit(1); });
