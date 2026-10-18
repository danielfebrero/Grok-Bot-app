/**
 * "Logistica" — Logistics & Transport presentation template (18 slides, 16:9 / 13.333x7.5in).
 * Recreated with pptxgenjs only. The source deck ships with empty picture
 * placeholders (no embedded raster art), so those regions are left blank here too.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const ORANGE = 'FE4905';
const INK = '262626';
const GRAY = '808080';
const WHITE = 'FFFFFF';
const PILL_GRAY = 'F2F2F2';
const TRACK = 'D9D9D9';

const TITLE_FONT = 'Poppins';
const BODY_FONT = 'Open Sans';

/**
 * Soft grey drop shadow used by every "card" panel in the deck.
 * Returns a fresh object each time: pptxgenjs rewrites shadow options in place.
 */
const cardShadow = () => ({ type: 'outer', color: 'BFBFBF', blur: 14, offset: 0.01, angle: 90, opacity: 0.55 });

/* ------------------------------------------------------- low level helpers */

/**
 * Body copy / headline text box. Every text box in the source is top-anchored
 * and set to "resize shape to fit text", so those are the defaults here.
 */
function text(slide, runs, o) {
  slide.addText(runs, Object.assign({ fontFace: BODY_FONT, fontSize: 11, color: GRAY, valign: 'top', fit: 'resize' }, o));
}

function rect(slide, x, y, w, h, fill) {
  slide.addShape('rect', { x, y, w, h, fill: { color: fill }, line: { type: 'none' } });
}

function roundRect(slide, x, y, w, h, fill, radius) {
  slide.addShape('roundRect', {
    x, y, w, h, rectRadius: radius, fill: { color: fill }, line: { type: 'none' },
  });
}

/** Two dots + a stubby bar: the accent mark that sits above every section title. */
function accentMark(slide, x, y, d) {
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color: INK }, line: { type: 'none' } });
  slide.addShape('ellipse', { x: x + 1.534 * d, y, w: d, h: d, fill: { color: INK }, line: { type: 'none' } });
  roundRect(slide, x + 3.07 * d, y, 7 * d, d, ORANGE, d / 2);
}

/** Section title: orange lead-in + ink remainder, with the accent mark above it. */
function sectionTitle(slide, o) {
  const size = o.fontSize || 32;
  accentMark(slide, o.markX || o.x + 0.116, o.y - 0.2, 0.082);
  text(slide, [
    { text: o.lead, options: { color: ORANGE } },
    { text: o.rest, options: { color: INK } },
  ], {
    x: o.x, y: o.y, w: o.w, h: o.h || 0.64,
    fontFace: TITLE_FONT, fontSize: size, bold: true, align: o.align || 'left',
  });
}

/** Card sub-heading (14pt Poppins bold). */
function cardHeading(slide, x, y, w, label, color) {
  text(slide, label, { x, y, w, h: 0.337, fontFace: TITLE_FONT, fontSize: 14, bold: true, color: color || INK });
}

/** Big orange percentage / index number (24pt Poppins bold). */
function bigNumber(slide, x, y, w, label, color, align) {
  text(slide, label, {
    x, y, w, h: 0.505, fontFace: TITLE_FONT, fontSize: 24, bold: true,
    color: color || ORANGE, wrap: false, align: align || 'left',
  });
}

/** Orange call-to-action button ("Read More" / "Learn More"). */
function button(slide, x, y, w, h, label) {
  roundRect(slide, x, y, w, h, ORANGE, 0.264 * h);
  text(slide, label, {
    x, y: y + (h - 0.286) / 2, w, h: 0.286,
    color: WHITE, italic: true, align: 'center', charSpacing: 3,
  });
}

/** Labelled progress bar used on the two team-member slides. */
function skillBar(slide, x, y, label, pct) {
  const full = 2.313;
  text(slide, label, { x: x - 0.12, y: y - 0.32, w: 0.823, h: 0.286, color: INK, italic: true, wrap: false });
  text(slide, pct + '%', { x: x + 1.898, y: y - 0.32, w: 0.505, h: 0.286, color: INK, italic: true, align: 'right', wrap: false });
  roundRect(slide, x, y, full, 0.216, TRACK, 0.108);
  roundRect(slide, x, y, full * (pct / 100), 0.216, ORANGE, 0.108);
}

/* ------------------------------------------------------------- freeform art */

/**
 * Build a pptxgenjs custGeom point list from compact path commands.
 * Commands use coordinates expressed as a fraction of the shape box:
 *   ['M',x,y] move, ['L',x,y] line, ['C',x1,y1,x2,y2,x,y] cubic bezier, ['Z'] close.
 * custGeom points are local to the shape box, so no origin is added.
 */
function freeform(slide, cmds, o) {
  const px = (v) => v * o.w;
  const py = (v) => v * o.h;
  const points = cmds.map((c) => {
    if (c[0] === 'M') return { x: px(c[1]), y: py(c[2]), moveTo: true };
    if (c[0] === 'L') return { x: px(c[1]), y: py(c[2]) };
    if (c[0] === 'Z') return { close: true };
    return { x: px(c[5]), y: py(c[6]), curve: { type: 'cubic', x1: px(c[1]), y1: py(c[2]), x2: px(c[3]), y2: py(c[4]) } };
  });
  slide.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h, points,
    fill: { color: o.fill }, line: { type: 'none' },
    shadow: o.shadow ? cardShadow() : undefined,
  });
}

/** Rectangle with individually rounded corners, as a freeform. */
function panel(slide, o) {
  const { w, h } = o;
  const c = Object.assign({ tl: 0, tr: 0, br: 0, bl: 0 }, o.corners);
  const k = 0.4477; // bezier circle constant, as a fraction of the radius
  const pts = [{ x: c.tl, y: 0, moveTo: true }];
  const arc = (x1, y1, x2, y2, ex, ey) => pts.push({ x: ex, y: ey, curve: { type: 'cubic', x1, y1, x2, y2 } });
  pts.push({ x: w - c.tr, y: 0 });
  if (c.tr) arc(w - c.tr * k, 0, w, c.tr * k, w, c.tr);
  pts.push({ x: w, y: h - c.br });
  if (c.br) arc(w, h - c.br * k, w - c.br * k, h, w - c.br, h);
  pts.push({ x: c.bl, y: h });
  if (c.bl) arc(c.bl * k, h, 0, h - c.bl * k, 0, h - c.bl);
  pts.push({ x: 0, y: c.tl });
  if (c.tl) arc(0, c.tl * k, c.tl * k, 0, c.tl, 0);
  pts.push({ close: true });
  slide.addShape('custGeom', {
    x: o.x, y: o.y, w, h, points: pts,
    fill: { color: o.fill }, line: { type: 'none' },
    shadow: o.shadow ? cardShadow() : undefined,
  });
}

/** The large orange bracket behind the cover / closing photo (slides 1 and 18). */
const HERO_BRACKET = [
  ['M', 0.9574, 0.0], ['L', 0.9589, 0.0017],
  ['C', 0.9843, 0.0362, 1.0, 0.084, 1.0, 0.1368],
  ['L', 1.0, 0.8089],
  ['C', 1.0, 0.9144, 0.9371, 1.0, 0.8596, 1.0],
  ['L', 0.5634, 1.0], ['L', 0.5607, 0.9996], ['L', 0.558, 1.0],
  ['L', 0.0, 1.0], ['L', 0.0, 0.9732], ['L', 0.8489, 0.9732],
  ['C', 0.9202, 0.9732, 0.978, 0.8945, 0.978, 0.7974],
  ['L', 0.978, 0.0946],
  ['C', 0.978, 0.0703, 0.9744, 0.0472, 0.9678, 0.0262],
  ['Z'],
];

/** Notched corner block, bottom-right of slide 2. */
const CORNER_BLOCK = [
  ['M', 0.5901, 0.0], ['L', 1.0, 0.0], ['L', 1.0, 1.0],
  ['L', 0.0, 1.0], ['L', 0.0, 0.7409], ['L', 0.5901, 0.7409], ['Z'],
];

/** Slim orange "parenthesis" outline (slide 3). */
const PAREN = [
  ['M', 0.0, 0.0], ['L', 0.2164, 0.0],
  ['C', 0.6492, 0.0, 1.0, 0.0746, 1.0, 0.1667],
  ['L', 1.0, 0.8333],
  ['C', 1.0, 0.9254, 0.6492, 1.0, 0.2164, 1.0],
  ['L', 0.0, 1.0],
  ['C', 0.4328, 1.0, 0.7836, 0.9254, 0.7836, 0.8333],
  ['L', 0.7836, 0.1667],
  ['C', 0.7836, 0.0746, 0.4328, 0.0, 0.0, 0.0],
  ['Z'],
];

/** Bottom-left orange step block with a rounded inner corner (slide 5). */
const STEP_BLOCK_L = [
  ['M', 0.0, 0.0], ['L', 0.2919, 0.0], ['L', 0.2919, 0.3896],
  ['C', 0.2919, 0.5073, 0.35, 0.6027, 0.4217, 0.6027],
  ['L', 1.0, 0.6027], ['L', 1.0, 1.0], ['L', 0.0, 1.0], ['Z'],
];

/** Right-hand orange panel with a scooped-out rounded bay (slide 7). */
const BAY_PANEL = [
  ['M', 0.0, 0.0], ['L', 1.0, 0.0], ['L', 1.0, 1.0], ['L', 0.0, 1.0],
  ['L', 0.0, 0.8498], ['L', 0.4782, 0.8498],
  ['C', 0.6244, 0.8498, 0.7429, 0.794, 0.7429, 0.7251],
  ['L', 0.7429, 0.31],
  ['C', 0.7429, 0.2412, 0.6244, 0.1854, 0.4782, 0.1854],
  ['L', 0.0, 0.1854], ['Z'],
];

/** Top-left orange step block with a rounded inner corner (slide 8). */
const STEP_BLOCK_T = [
  ['M', 0.0, 0.0], ['L', 1.0, 0.0], ['L', 1.0, 0.2773], ['L', 0.605, 0.2773],
  ['C', 0.5298, 0.2773, 0.4689, 0.3235, 0.4689, 0.3805],
  ['L', 0.4689, 1.0], ['L', 0.0, 1.0], ['Z'],
];

/** Open-sided orange frame around the portrait on slide 13. */
const OPEN_FRAME = [
  ['M', 0.0, 0.0], ['L', 1.0, 0.0], ['L', 1.0, 1.0], ['L', 0.0, 1.0],
  ['L', 0.0, 0.9766], ['L', 0.9297, 0.9766], ['L', 0.9297, 0.0234], ['L', 0.0, 0.0234], ['Z'],
];

/** Full-width orange band with four scalloped bays for the team portraits (slide 11). */
const TEAM_BAYS = [1.0878, 3.9363, 6.7848, 9.6333];
function teamBand(slide) {
  const y = 5.254, w = 13.333, h = 2.246, r = 0.4353, depth = 1.3924, bay = 2.6123;
  const k = 0.4477;
  const pts = [{ x: 0, y: 0, moveTo: true }];
  TEAM_BAYS.forEach((L) => {
    const R = L + bay;
    pts.push({ x: L, y: 0 });
    pts.push({ x: L, y: depth - r });
    pts.push({ x: L + r, y: depth, curve: { type: 'cubic', x1: L, y1: depth - r * k, x2: L + r * k, y2: depth } });
    pts.push({ x: R - r, y: depth });
    pts.push({ x: R, y: depth - r, curve: { type: 'cubic', x1: R - r * k, y1: depth, x2: R, y2: depth - r * k } });
    pts.push({ x: R, y: 0 });
  });
  pts.push({ x: w, y: 0 }, { x: w, y: h }, { x: 0, y: h }, { close: true });
  slide.addShape('custGeom', {
    x: 0, y, w, h, points: pts, fill: { color: ORANGE }, line: { type: 'none' },
  });
}

/* ------------------------------------------------------------ page furniture */

/** Hamburger glyph in the top-right corner. */
function hamburger(slide, color) {
  [0, 0.0646, 0.1291].forEach((dy) => roundRect(slide, 12.736, 0.332 + dy, 0.178, 0.0419, color, 0.021));
}

/** 3x3 square dot grid used as a decorative corner motif. */
function dotGrid(slide, x, y) {
  for (let r = 0; r < 3; r++) {
    for (let c = 0; c < 3; c++) rect(slide, x + c * 0.2634, y + r * 0.2634, 0.0707, 0.0707, ORANGE);
  }
}
const GRID_TOP_LEFT = [0.416, 0.332];
const GRID_BOTTOM_RIGHT = [12.317, 6.476];

/** Two offset squares tucked into the bottom-right corner. */
function cornerSquares(slide, color) {
  rect(slide, 12.9, 7.054, 0.264, 0.257, color);
  rect(slide, 12.722, 6.899, 0.178, 0.174, color);
}

/** The grey/white "COMPANY WEBSITE HERE" pill along the bottom edge. */
function websitePill(slide, fill, color, y) {
  const top = y || 6.761;
  roundRect(slide, 4.935, top, 3.464, 0.483, fill, 0.1275);
  text(slide, 'COMPANY WEBSITE HERE', {
    x: 5.089, y: top + 0.066, w: 3.143, h: 0.327,
    fontSize: 10, color, align: 'center', charSpacing: 3, lineSpacingMultiple: 1.5,
  });
}

/** Autofitted width of each page-number box, taken from the source deck. */
const PAGE_NUM_WIDTH = [
  0.376, 0.409, 0.414, 0.426, 0.421, 0.419, 0.402, 0.421, 0.416,
  0.376, 0.328, 0.361, 0.367, 0.379, 0.374, 0.372, 0.354, 0.374,
];

/**
 * Chrome shared by every slide: hamburger + "Company Name", page number,
 * and the optional decorative motifs.
 */
function chrome(slide, n, o) {
  const opts = o || {};
  const headerColor = opts.header || GRAY;
  hamburger(slide, opts.menu || ORANGE);
  text(slide, 'Company Name', {
    x: 11.28, y: 0.274, w: 1.355, h: 0.286,
    italic: true, color: headerColor, align: 'right', wrap: false,
  });
  (opts.grids || []).forEach((g) => dotGrid(slide, g[0], g[1]));
  if (opts.corners) cornerSquares(slide, opts.corners);
  if (opts.website) websitePill(slide, opts.website[0], opts.website[1], opts.website[2]);
  text(slide, String(n).padStart(2, '0'), {
    x: 0.209, y: opts.pageY || 7.002, w: PAGE_NUM_WIDTH[n - 1], h: 0.303,
    fontFace: TITLE_FONT, fontSize: 12, bold: true, color: opts.page || INK, wrap: false,
  });
}

/** Thin orange rule. */
function rule(slide, x, y, w) {
  slide.addShape('line', { x, y, w, h: 0, line: { color: ORANGE, width: 0.75 } });
}

/* ------------------------------------------------------------------- lorem */
const L_LONG = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco';
const L_MED = 'Lorem ipsum dolor amet, consectetur adipiscing elit, sed do eiusmod tempor ';
const L_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipiscing';

/* ------------------------------------------------------------ slide builders */

function slide01(s) {
  freeform(s, HERO_BRACKET, { x: 0.121, y: 1.949, w: 5.497, h: 4.038, fill: ORANGE });
  rule(s, 0.905, 0.778, 12.429);
  accentMark(s, 7.014, 2.735, 0.104);
  text(s, [
    { text: 'Logist', options: { color: ORANGE } },
    { text: 'ica', options: { color: INK } },
  ], { x: 6.857, y: 2.839, w: 4.877, h: 1.313, fontFace: TITLE_FONT, fontSize: 72, bold: true, wrap: false });
  text(s, 'Logistics & Transport Presentation Template', { x: 6.906, y: 4.151, w: 3.677, h: 0.303, fontSize: 12, wrap: false });
  button(s, 7.014, 4.718, 1.628, 0.483, 'Read More');
  chrome(s, 1, { grids: [GRID_BOTTOM_RIGHT] });
}

function slide02(s) {
  freeform(s, CORNER_BLOCK, { x: 10.474, y: 3.75, w: 2.86, h: 3.747, fill: ORANGE });
  s.addShape('rect', { x: 0, y: 1.111, w: 8.721, h: 5.279, fill: { color: WHITE }, line: { type: 'none' }, shadow: cardShadow() });
  sectionTitle(s, { x: 1.554, y: 2.338, w: 3.268, h: 1.178, lead: 'Welcome ', rest: 'To Logistica' });
  bigNumber(s, 1.559, 3.945, 0.905, '28%');
  text(s, 'Lorem ipsum dolor sit consectetur adipiscing sed eiusmod', { x: 1.554, y: 4.457, w: 1.82, h: 0.905, lineSpacingMultiple: 1.5 });
  text(s, L_LONG, { x: 4.024, y: 3.902, w: 3.143, h: 1.461, lineSpacingMultiple: 1.5 });
  chrome(s, 2, { grids: [[7.635, 1.54]], corners: WHITE, website: [PILL_GRAY, GRAY] });
}

function slide03(s) {
  panel(s, { x: 1.467, y: 1.111, w: 11.866, h: 5.279, fill: WHITE, shadow: true, corners: { tl: 0.461, bl: 0.461 } });
  freeform(s, PAREN, { x: 6.154, y: 1.659, w: 0.889, h: 4.182, fill: ORANGE });
  sectionTitle(s, { x: 8.232, y: 2.258, w: 3.268, h: 1.178, lead: 'About ', rest: 'Our Company' });
  text(s, L_LONG, { x: 8.232, y: 3.467, w: 3.912, h: 1.183, lineSpacingMultiple: 1.5 });
  button(s, 8.287, 4.959, 1.628, 0.483, 'Read More');
  chrome(s, 3, { grids: [GRID_TOP_LEFT], corners: ORANGE, website: [PILL_GRAY, GRAY] });
}

function slide04(s) {
  rect(s, 10.641, 0, 2.693, 6.844, ORANGE);
  sectionTitle(s, { x: 1.554, y: 1.657, w: 5.446, h: 1.717, lead: 'We Are The Best ', rest: 'Logisticts & Transport In Town' });
  text(s, L_LONG + ' laboris nisi ut aliquip ex ea commodo consequat. ', { x: 1.554, y: 3.555, w: 5.446, h: 1.183, lineSpacingMultiple: 1.5 });
  bigNumber(s, 1.559, 4.901, 0.905, '28%');
  text(s, L_SHORT, { x: 1.554, y: 5.413, w: 2.381, h: 0.627, lineSpacingMultiple: 1.5 });
  bigNumber(s, 4.624, 4.901, 0.942, '45%');
  text(s, L_SHORT, { x: 4.619, y: 5.413, w: 2.381, h: 0.627, lineSpacingMultiple: 1.5 });
  chrome(s, 4, { grids: [GRID_TOP_LEFT], menu: WHITE, header: WHITE });
}

function slide05(s) {
  panel(s, { x: 6.795, y: 4.068, w: 6.538, h: 2.004, fill: WHITE, shadow: true, corners: { tl: 0.175, bl: 0.175 } });
  freeform(s, STEP_BLOCK_L, { x: 0, y: 5.078, w: 3.974, h: 2.422, fill: ORANGE });
  sectionTitle(s, { x: 1.381, y: 1.274, w: 3.537, h: 1.178, lead: 'Our Company', rest: ' History' });
  text(s, L_LONG + ' laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum',
    { x: 6.078, y: 1.274, w: 5.874, h: 1.183, align: 'justify', lineSpacingMultiple: 1.5 });
  [['2010 - 2017', 7.521], ['2017 - 2023', 10.226]].forEach(([label, x]) => {
    cardHeading(s, x + 0.005, 4.588, 1.4, label, ORANGE);
    text(s, L_SHORT, { x, y: 4.925, w: 2.381, h: 0.627, align: 'justify', lineSpacingMultiple: 1.5 });
  });
  rule(s, 5.032, 6.571, 8.302);
  chrome(s, 5, { grids: [GRID_TOP_LEFT], page: WHITE, corners: ORANGE, website: [PILL_GRAY, GRAY] });
}

function slide06(s) {
  rect(s, 10.116, 5.811, 3.217, 1.689, ORANGE);
  panel(s, { x: 3.395, y: 5.014, w: 8.925, h: 1.594, fill: WHITE, shadow: true, corners: { tl: 0.2, tr: 0.2, br: 0.2, bl: 0.2 } });
  sectionTitle(s, { x: 6.394, y: 1.371, w: 4.306, h: 1.178, lead: 'Our Company', rest: ' Vision & Mission' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris  ',
    { x: 6.394, y: 2.66, w: 5.084, h: 0.905, lineSpacingMultiple: 1.5 });
  button(s, 6.465, 3.868, 1.628, 0.483, 'Read More');
  [['Company Vision', 4.02], ['Company Mission', 8.093]].forEach(([label, x]) => {
    cardHeading(s, x + 0.005, 5.467, 2.1, label, ORANGE);
    text(s, 'Lorem ipsum dolor amet, consectetur adipiscing', { x, y: 5.804, w: 3.695, h: 0.35, align: 'justify', lineSpacingMultiple: 1.5 });
  });
  chrome(s, 6, { grids: [GRID_TOP_LEFT], corners: WHITE });
}

function slide07(s) {
  freeform(s, BAY_PANEL, { x: 9.802, y: 0, w: 3.531, h: 7.5, fill: ORANGE });
  sectionTitle(s, { x: 1.381, y: 1.393, w: 3.537, h: 1.178, lead: 'Our Excellent ', rest: 'Service' });
  const items = [
    ['01', 'Logistic Service', 'Lorem ipsum dolor amet, consectetur adipiscing elit, sed do eiusmod tempor ', 2.992, 3.695, 'justify', 0.624],
    ['02', 'Package Delivery', 'Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, ', 4.101, 3.695, 'justify', 0.614],
    ['03', 'International Logistics', 'Ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit  ', 5.216, 3.945, 'left', 0.624],
  ];
  items.forEach(([num, head, body, y, bw, algn, nw]) => {
    bigNumber(s, 1.381, y + 0.134, nw, num);
    cardHeading(s, 2.327, y, 2.6, head);
    text(s, body, { x: 2.322, y: y + 0.241, w: bw, h: 0.627, align: algn, lineSpacingMultiple: 1.5 });
  });
  chrome(s, 7, { grids: [GRID_TOP_LEFT], menu: WHITE, header: WHITE, corners: WHITE, website: [PILL_GRAY, GRAY] });
}

/**
 * Numbered card used on slides 8 and 9: rounded white panel, orange number
 * bubble at [bx, by], heading at [tx, hy] and body copy at [tx, ty].
 */
function bubbleCard(s, c) {
  panel(s, { x: c.x, y: c.y, w: c.w, h: 1.847, fill: WHITE, shadow: true, corners: { tl: 0.23, tr: 0.23, br: 0.23, bl: 0.23 } });
  s.addShape('ellipse', { x: c.bx, y: c.by, w: 0.878, h: 0.878, fill: { color: ORANGE }, line: { type: 'none' } });
  bigNumber(s, c.bx, c.by + 0.215, 0.878, c.num, WHITE, 'center');
  cardHeading(s, c.tx, c.hy, 2.8, c.head);
  text(s, c.body, { x: c.tx - 0.005, y: c.ty, w: c.tw, h: 0.627, align: c.align || 'left', lineSpacingMultiple: 1.5 });
}

function slide08(s) {
  freeform(s, STEP_BLOCK_T, { x: 0, y: 0, w: 2.842, h: 3.75, fill: ORANGE });
  bubbleCard(s, { x: 3.215, y: 1.631, w: 5.374, bx: 2.776, by: 2.072, tx: 4.162, hy: 2.072, ty: 2.409, tw: 3.695, align: 'justify', num: '01', head: 'Logistic Service', body: L_MED });
  bubbleCard(s, { x: 5.285, y: 4.023, w: 5.374, bx: 4.849, by: 4.508, tx: 6.242, hy: 4.468, ty: 4.805, tw: 3.695, num: '02', head: 'International Logistics', body: 'Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea  ' });
  chrome(s, 8, { grids: [GRID_BOTTOM_RIGHT], website: [PILL_GRAY, GRAY] });
}

function slide09(s) {
  bubbleCard(s, { x: 1.171, y: 0.901, w: 4.95, bx: 1.492, by: 1.343, tx: 2.637, hy: 1.284, ty: 1.621, tw: 3.458, num: '01', head: 'Logistic Services', body: L_MED });
  bubbleCard(s, { x: 1.171, y: 4.598, w: 4.95, bx: 1.492, by: 5.083, tx: 2.632, hy: 5.101, ty: 5.376, tw: 3.136, num: '02', head: 'Package Delivery', body: 'Quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo  ' });
  sectionTitle(s, { x: 7.706, y: 2.062, w: 4.135, h: 1.178, lead: 'The Service ', rest: 'We Provide' });
  text(s, L_LONG + ' laboris nisi ut aliquip ex ea commodo consequat. ', { x: 7.706, y: 3.412, w: 3.937, h: 1.461, lineSpacingMultiple: 1.5 });
  button(s, 7.822, 5.311, 1.84, 0.541, 'Learn More');
  chrome(s, 9, { grids: [GRID_BOTTOM_RIGHT], website: [PILL_GRAY, GRAY] });
}

function slide10(s) {
  rect(s, 0, 5.288, 9.638, 2.212, ORANGE);
  s.addShape('rect', { x: 1.534, y: 4.286, w: 9.812, h: 2.004, fill: { color: WHITE }, line: { type: 'none' }, shadow: cardShadow() });
  sectionTitle(s, { x: 1.534, y: 2.069, w: 3.455, lead: 'Our ', rest: 'Services' });
  text(s, L_LONG + ' laboris nisi ut aliquip ex ea commodo consequat. ', { x: 1.534, y: 2.749, w: 6.773, h: 0.905, lineSpacingMultiple: 1.5 });
  [['International Logistics', 2.425], ['Logistic Service', 6.77]].forEach(([label, x]) => {
    cardHeading(s, x, 4.806, 2.7, label, ORANGE);
    text(s, L_MED, { x: x - 0.005, y: 5.143, w: 3.695, h: 0.627, lineSpacingMultiple: 1.5 });
  });
  chrome(s, 10, { grids: [GRID_TOP_LEFT], page: WHITE, website: [WHITE, ORANGE] });
}

function slide11(s) {
  teamBand(s);
  sectionTitle(s, { x: 4.627, y: 1.47, w: 4.079, markX: 4.949, lead: 'Meet ', rest: 'Our Team', align: 'center' });
  text(s, L_LONG + ' laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ',
    { x: 1.841, y: 2.288, w: 9.652, h: 0.905, align: 'center', lineSpacingMultiple: 1.5 });
  chrome(s, 11, { grids: [GRID_TOP_LEFT], page: WHITE, corners: WHITE, website: [WHITE, ORANGE, 6.819] });
}

function slide12(s) {
  rect(s, 0, 0, 2.071, 7.5, ORANGE);
  s.addShape('rect', { x: 4.418, y: 2.087, w: 7.408, h: 3.957, fill: { color: WHITE }, line: { type: 'none' }, shadow: cardShadow() });
  rect(s, 11.639, 1.891, 0.388, 0.399, ORANGE);
  rect(s, 12.013, 1.631, 0.262, 0.269, ORANGE);
  sectionTitle(s, { x: 5.308, y: 3.088, w: 4.349, lead: 'Graham ', rest: 'Braxton' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore',
    { x: 5.308, y: 4.048, w: 2.234, h: 1.183, lineSpacingMultiple: 1.5 });
  skillBar(s, 8.444, 4.186, 'Skill One', 80);
  skillBar(s, 8.444, 4.96, 'Skill Two', 90);
  chrome(s, 12, { grids: [GRID_BOTTOM_RIGHT], page: WHITE, website: [PILL_GRAY, GRAY] });
}

function slide13(s) {
  freeform(s, OPEN_FRAME, { x: 8.217, y: 0.74, w: 2.071, h: 6.02, fill: ORANGE });
  s.addShape('rect', { x: 9.474, y: 1.588, w: 3.859, h: 4.324, fill: { color: WHITE }, line: { type: 'none' }, shadow: cardShadow() });
  sectionTitle(s, { x: 1.318, y: 2.742, w: 4.094, h: 0.555, lead: 'Joseph ', rest: 'Fitzgerald', fontSize: 27 });
  text(s, L_LONG, { x: 1.335, y: 3.362, w: 3.851, h: 1.183, lineSpacingMultiple: 1.5 });
  button(s, 1.335, 4.868, 1.84, 0.541, 'Learn More');
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt  ',
    { x: 10.157, y: 2.392, w: 2.523, h: 0.905, lineSpacingMultiple: 1.5 });
  skillBar(s, 10.262, 4.139, 'Skill One', 80);
  skillBar(s, 10.262, 4.913, 'Skill Two', 90);
  chrome(s, 13, { grids: [GRID_TOP_LEFT, GRID_BOTTOM_RIGHT], website: [PILL_GRAY, GRAY] });
}

function slide14(s) {
  s.addShape('rect', { x: 0, y: 0, w: 10.714, h: 5.279, fill: { color: WHITE }, line: { type: 'none' }, shadow: cardShadow() });
  sectionTitle(s, { x: 1.409, y: 1.545, w: 3.64, h: 1.178, lead: 'Our Company', rest: ' Portfolio' });
  text(s, 'Lorem ipsum dolor sit consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco  ',
    { x: 1.409, y: 2.791, w: 3.924, h: 1.183, lineSpacingMultiple: 1.5 });
  bigNumber(s, 5.333, 6.058, 0.905, '28%');
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore ',
    { x: 6.486, y: 5.935, w: 3.853, h: 0.627, lineSpacingMultiple: 1.5 });
  panel(s, { x: 11.392, y: 6.071, w: 1.941, h: 0.633, fill: ORANGE, corners: { tl: 0.3165, bl: 0.3165 } });
  text(s, 'Learn More', { x: 11.636, y: 6.245, w: 1.454, h: 0.286, color: WHITE, italic: true, align: 'center', charSpacing: 3, wrap: false });
  chrome(s, 14, { pageY: 4.73 });
}

function slide15(s) {
  s.addShape('rect', { x: 3.593, y: 0.778, w: 9.74, h: 3.756, fill: { color: WHITE }, line: { type: 'none' }, shadow: cardShadow() });
  rect(s, 8.484, 6.017, 4.849, 1.483, ORANGE);
  sectionTitle(s, { x: 1.519, y: 5.689, w: 3.761, lead: 'Our ', rest: 'Best Work' });
  text(s, L_LONG, { x: 5.28, y: 1.648, w: 2.344, h: 2.016, lineSpacingMultiple: 1.5 });
  [['The Best Work', 1.648], ['Logistic Services', 2.693]].forEach(([label, y]) => {
    cardHeading(s, 8.437, y, 2.4, label, ORANGE);
    text(s, L_MED, { x: 8.432, y: y + 0.337, w: 3.695, h: 0.627, align: 'justify', lineSpacingMultiple: 1.5 });
  });
  chrome(s, 15, { grids: [GRID_TOP_LEFT] });
}

function slide16(s) {
  rect(s, 0, 5.618, 9.291, 1.882, ORANGE);
  s.addShape('rect', { x: 8.869, y: 1.397, w: 4.464, h: 2.169, fill: { color: WHITE }, line: { type: 'none' }, shadow: cardShadow() });
  sectionTitle(s, { x: 1.487, y: 1.514, w: 3.761, lead: 'Our ', rest: 'Portfolio' });
  text(s, L_LONG + ' laboris nisi ut aliquip ex ea commodo consequat. ', { x: 1.487, y: 2.186, w: 6.773, h: 0.905, lineSpacingMultiple: 1.5 });
  bigNumber(s, 9.479, 1.684, 0.905, '28%');
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore',
    { x: 9.479, y: 2.141, w: 3.245, h: 0.905, lineSpacingMultiple: 1.5 });
  chrome(s, 16, { grids: [GRID_TOP_LEFT], page: WHITE, corners: ORANGE });
}

function slide17(s) {
  rect(s, 0, 5.207, 3.573, 2.293, ORANGE);
  s.addShape('rect', { x: 0.847, y: 3.599, w: 5.452, h: 3.21, fill: { color: WHITE }, line: { type: 'none' }, shadow: cardShadow() });
  sectionTitle(s, { x: 1.487, y: 1.752, w: 3.092, lead: 'Contact ', rest: 'Us' });
  text(s, L_LONG + ' laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ',
    { x: 5.523, y: 1.208, w: 6.324, h: 1.183, align: 'justify', lineSpacingMultiple: 1.5 });
  [['Our Mail Address', 4.147], ['Visit Our Office', 5.297]].forEach(([label, y]) => {
    cardHeading(s, 1.731, y, 2.9, label, ORANGE);
    text(s, L_MED, { x: 1.726, y: y + 0.337, w: 3.695, h: 0.627, lineSpacingMultiple: 1.5 });
  });
  chrome(s, 17, { grids: [GRID_TOP_LEFT], page: WHITE });
}

function slide18(s) {
  freeform(s, HERO_BRACKET, { x: 0.121, y: 1.949, w: 5.497, h: 4.038, fill: ORANGE });
  rule(s, 0.905, 0.778, 12.429);
  accentMark(s, 7.025, 2.514, 0.104);
  text(s, [
    { text: 'Thank ', options: { color: ORANGE } },
    { text: 'You', options: { color: INK } },
  ], { x: 6.917, y: 2.936, w: 4.997, h: 1.111, fontFace: TITLE_FONT, fontSize: 60, bold: true, wrap: false });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ',
    { x: 6.924, y: 4.051, w: 5.682, h: 0.627, lineSpacingMultiple: 1.5 });
  button(s, 7.025, 4.956, 1.628, 0.483, 'Read More');
  chrome(s, 18, { grids: [GRID_BOTTOM_RIGHT], website: [PILL_GRAY, GRAY] });
}

/* --------------------------------------------------------------------- main */

const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
  slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18,
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'LOGISTICA', width: 13.333, height: 7.5 });
  pptx.layout = 'LOGISTICA';
  pptx.author = 'Logistica';
  pptx.title = 'Logistics & Transport Presentation Template';
  BUILDERS.forEach((fn) => {
    const slide = pptx.addSlide();
    slide.background = { color: WHITE };
    fn(slide);
  });
  return pptx;
}

build().writeFile({ fileName: path.join(__dirname, '194be181-91b0-4329-a23a-6f1b81084b86_grok_final.pptx') })
  .then((f) => console.log('wrote', f))
  .catch((e) => { console.error(e); process.exit(1); });
