/**
 * "Fintech — Financial Presentation" (36 slides), rebuilt with pptxgenjs.
 *
 * Coordinates are inches on a 13.333 x 7.5 stage and mirror the source deck.
 * Bitmap artwork (icon graphics, photo frames) is replaced by flat placeholder
 * shapes; charts, tables and every other object are rebuilt natively.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

const W = 13.333;
const H = 7.5;

/* ------------------------------------------------------------------ palette */

const C = {
  black: '000000',
  white: 'FFFFFF',
  smoke: 'F2F2F2',      // bg1 @95% – default body copy on the dark slides
  silver: 'D9D9D9',
  ash: 'BFBFBF',
  grey: '808080',
  slate: '595959',
  charcoal: '404040',
  ink: '0D0D0D',
  photo: '000000',      // an empty photo frame reads as bare background
  panel: '36424A',      // barely-there card wash used on dark slides
  blue: '0000CC',       // accent1
  blueBright: '3333FF', // accent2
  blueDeep: '3333CC',   // accent3
  blueSky: '3366FF',    // accent4
  violet: '6600FF',     // accent5
  periwinkle: '6666FF', // accent6
  navy: '000099',
  indigo: '262699',
  cobalt: '2626BF',
  royal: '264CBF',
  grape: '4C00BF',
  mapGrey: '6C6C6C',
};

const FONT = { head: 'Montserrat Medium', body: 'Lato' };
const NOLINE = { type: 'none' };

/* ------------------------------------------------------------------ helpers */

const fill = (color, transparency) =>
  (transparency ? { color, transparency } : { color });

/** Body copy: Lato, light-on-dark, top anchored. */
function text(slide, str, o = {}) {
  slide.addText(str, Object.assign(
    { fontFace: FONT.body, fontSize: 18, color: C.smoke, valign: 'top', margin: 0 },
    o
  ));
}

/** Montserrat Medium headline. */
function heading(slide, str, o = {}) {
  text(slide, str, Object.assign({ fontFace: FONT.head, fontSize: 40 }, o));
}

function rect(slide, x, y, w, h, o = {}) {
  slide.addShape('rect', Object.assign({ x, y, w, h, line: NOLINE }, o));
}

function roundRect(slide, x, y, w, h, radius, o = {}) {
  slide.addShape('roundRect', Object.assign({ x, y, w, h, rectRadius: radius, line: NOLINE }, o));
}

function ellipse(slide, x, y, w, h, o = {}) {
  slide.addShape('ellipse', Object.assign({ x, y, w, h, line: NOLINE }, o));
}

function pie(slide, x, y, d, angleRange, o = {}) {
  slide.addShape('pie', Object.assign({ x, y, w: d, h: d, angleRange, line: NOLINE }, o));
}

function line(slide, x, y, w, h, o = {}) {
  slide.addShape('line', Object.assign({ x, y, w, h }, o));
}

/** Linear blend between two hex colors, t in [0,1]. */
function mix(a, b, t) {
  let out = '';
  for (let i = 0; i < 3; i++) {
    const ca = parseInt(a.substr(i * 2, 2), 16);
    const cb = parseInt(b.substr(i * 2, 2), 16);
    out += Math.round(ca + (cb - ca) * t).toString(16).padStart(2, '0');
  }
  return out.toUpperCase();
}

/**
 * pptxgenjs has no gradient fill, so gradients are painted as a grid of solid
 * tiles. `dir` is 'x', 'y' or 'xy' (diagonal, top-left -> bottom-right).
 */
function gradient(slide, x, y, w, h, from, to, o = {}) {
  const cols = o.dir === 'y' ? 1 : (o.cols || 18);
  const rows = o.dir === 'x' ? 1 : (o.rows || 10);
  const bleed = 0.06;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const u = (c + 0.5) / cols;
      const v = (r + 0.5) / rows;
      const t = o.dir === 'x' ? u : o.dir === 'y' ? v : (u + v) / 2;
      rect(slide, x + (w * c) / cols, y + (h * r) / rows,
        w / cols + bleed, h / rows + bleed, { fill: fill(mix(from, to, t)) });
    }
  }
}

/** Empty photo frame from the source deck, shown as a flat block. */
function photo(slide, x, y, w, h, o = {}) {
  rect(slide, x, y, w, h, { fill: fill(o.fill || C.photo, o.transparency) });
  if (o.label) {
    text(slide, o.label, {
      x, y: y + h / 2 - 0.2, w, h: 0.4,
      align: 'center', fontSize: 16, color: o.labelColor || C.grey,
    });
  }
}

/**
 * Stand-in for a small line-art icon graphic: an outlined rounded square,
 * which keeps the visual weight of the original stroked glyphs.
 */
function icon(slide, x, y, size, color) {
  slide.addShape('roundRect', {
    x, y, w: size, h: size, rectRadius: size * 0.22,
    fill: { type: 'none' }, line: { color, width: Math.max(0.75, size * 3) },
  });
}

/** Ring gauge: track + coloured sweep + hub + "NN%" label. */
function gauge(slide, x, y, d, o) {
  ellipse(slide, x, y, d, d, { fill: fill(o.track, o.trackTransparency) });
  pie(slide, x, y, d, o.angleRange, { fill: fill(o.sweep) });
  ellipse(slide, x + d * 0.09, y + d * 0.09, d * 0.82, d * 0.82, { fill: fill(o.hub) });
  text(slide, [
    { text: o.value, options: { fontSize: 24, color: o.labelColor } },
    { text: '%', options: { fontSize: 10.5, color: o.labelColor } },
  ], { x: x + d * 0.06, y: y + d * 0.24, w: d * 0.78, h: 0.5, align: 'center' });
}

/**
 * Paint an ASCII map. Every non-'.' character is a region keyed in `colors`;
 * horizontal runs of the same region collapse into a single rectangle.
 */
function mapMask(slide, box, mask, colors) {
  const rows = mask.length;
  const cw = box.w / mask[0].length;
  const ch = box.h / rows;
  for (let r = 0; r < rows; r++) {
    let c = 0;
    while (c < mask[r].length) {
      const key = mask[r][c];
      let end = c;
      while (end + 1 < mask[r].length && mask[r][end + 1] === key) end++;
      if (key !== '.') {
        rect(slide, box.x + c * cw, box.y + r * ch,
          (end - c + 1) * cw + 0.008, ch + 0.008, { fill: fill(colors[key]) });
      }
      c = end + 1;
    }
  }
}

/** Paint an ASCII map as a halftone dot field. */
function mapDots(slide, box, mask, color) {
  const cw = box.w / mask[0].length;
  const ch = box.h / mask.length;
  mask.forEach((row, r) => {
    for (let c = 0; c < row.length; c++) {
      if (row[c] !== '.') {
        ellipse(slide, box.x + c * cw, box.y + r * ch, cw * 0.55, ch * 0.55, { fill: fill(color) });
      }
    }
  });
}

/* ------------------------------------------------------------------- assets */

const ASIA_MAP = [
  '...................###......................',
  '...............########.....................',
  '..............############..................',
  '..........##########################........',
  '..........################################..',
  '........##################################..',
  '........#################################...',
  '........###############################.....',
  '........#######################....##.......',
  '........######################.....##.......',
  '........######################..............',
  '......#########################.............',
  '......########################..............',
  '......#######################...............',
  '......#####################.................',
  '..###..##################...................',
  '..###..##################...................',
  '.########################...................',
  '#########################...................',
  '.#######################....................',
  '..######..##############....................',
  '..######...####..####.......................',
  '...###......##....###.......................',
  '............##.....###......................',
  '....................#####...................',
  '.....................####...................',
  '......................##....................',
];

/** Australia by state: W=Western, G=grey states, Q=Queensland, S=South. */
const AUSTRALIA_MAP = [
  '.................GGG................',
  '................GGGGG...............',
  '............WW.GGGGGG.....QQ........',
  '...........WWWWGGGGGG.....QQQ.......',
  '..........WWWWWGGGGGGG....QQQ.......',
  '.........WWWWWWGGGGGGGG..QQQQQ......',
  '........WWWWWWWGGGGGGGGQQQQQQQQ.....',
  '....W..WWWWWWWWGGGGGGGGQQQQQQQQQ....',
  '...WWWWWWWWWWWGGGGGGGGGQQQQQQQQ.....',
  '..WWWWWWWWWWWWGGGGGGGGGQQQQQQQ......',
  '.WWWWWWWWWWWWWGGGGGGGGQQQQQQQQQ.....',
  '.WWWWWWWWWWWWWGGGGGGGGQQQQQQQQQQQQ..',
  '.WWWWWWWWWWWWWGGGGGGGGQQQQQQQQQQQQQ.',
  '.WWWWWWWWWWWWWGGGGGGGGQQQQQQQQQQQQQQ',
  '.WWWWWWWWWWWWWSSSSSSSSSSQQQQQQQQQQQQ',
  '.WWWWWWWWWWWWWSSSSSSSSSSSQQQQQQQQQQQ',
  '.WWWWWWWWWWWWWSSSSSSSSSSSQQQQQQQQQQG',
  '..WWWWWWWWWWWWSSSSSSSSSSSGGGGGGGGGGG',
  '..WWWWWWWWWWWWSSSSSSSSSSSGGGGGGGGGG.',
  '..WWWWWWWWWWWWSSSSSSSSSSGGGGGGGGGGG.',
  '..WWWWWWWWWWW..SSSSSSSSSGGGGGGGGGGG.',
  '..WWWWWWWW......S..SSSSSGGGGGGGGGG..',
  '...WWW..............SSSSGGGGGGGGG...',
  '....................SSSSGGGGGGGG....',
  '......................SSGGGGGGGG....',
  '........................GGGGGGGG....',
  '.........................GGGG.......',
];

const WORLD_MAP = [
  '.........................############.####.####.####....................',
  '....................##############..########.############..............',
  '..................################..#####################.#............',
  '..............##########.####.##..############################.........',
  '.............##########..##......###############################.......',
  '...........#############.........##################################....',
  '.........################........###################################...',
  '.....####################........###############################.####..',
  '...###################...........#####.#########################.###...',
  '..#####.#############............#############################.####....',
  '..###....##########.####..###..####################################....',
  '..##.....##########.###..##....#################..############.######..',
  '###......#####..######.....#...################....##..#############...',
  '#.........##########......###..###############.....###..###.####.#####.',
  '...........###########....###..##############...........###..###...####',
  '...........###########....##...#############.............##..####.####.',
  '............##########.........############..............#..####.###...',
  '............#########..........###########...................####.##...',
  '.............########..........##########....................#####.....',
  '..............#######..........#########......................####.....',
  '..............######...........########.......................###......',
  '...............#####...........#######........................###......',
  '................####...........######.........................####.....',
  '.................###...........#####.....................###..#####....',
  '.................##.............###......................###....###....',
  '.................#..............##.......................##............',
];

const LOREM_SHORT = 'Lorem ipsum dolor sit amet, conseert ctetur adipiscing elit lorem';
const SERENITY = 'A wonderful serenity has taken possession of my entire soul, like these ' +
  'sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the ' +
  'charm of existence.';
const UTWISI = 'Ut wisi enim ad minim veniam, quis nostrud exerci tation';

/* -------------------------------------------------------------- slide 1..36 */

function slide01(s) {
  gradient(s, 0, 0, W, H, '000084', '3B3BB4', { cols: 16, rows: 9 });
  heading(s, 'Fintech', {
    x: 1.958, y: 2.282, w: 9.417, h: 1.717,
    fontSize: 96, align: 'center', lineSpacingMultiple: 1.0,
  });
  text(s, 'Financial Presentation ', {
    x: 4.265, y: 3.999, w: 4.804, h: 0.572, fontSize: 28, align: 'center', color: C.white,
  });
}

function slide02(s) {
  roundRect(s, 7.792, 3.375, 5.903, 2.875, 0.33, { fill: fill(C.panel, 97) });
  roundRect(s, 8.042, 5.103, 5.903, 0.822, 0.19, { fill: fill(C.white) });
  ellipse(s, 8.435, 5.26, 0.5, 0.5, { fill: fill(C.blue) });
  icon(s, 8.566, 5.39, 0.244, C.white);
  text(s, 'Earn with Fintech',
    { x: 9.252, y: 5.295, w: 2.817, h: 0.438, fontSize: 20, bold: true, color: C.slate });

  [['Content one for financial', 3.747], ['Content two for financial', 4.431]].forEach(([msg, y]) => {
    icon(s, 8.502, y + 0.022, 0.339, C.ash);
    text(s, msg, { x: 9.297, y, w: 2.883, h: 0.382, fontSize: 14, lineSpacingMultiple: 1.3 });
  });

  heading(s, 'Easy Payment With Us', { x: 1.255, y: 2.445, w: 7.311, h: 0.774 });
  text(s, 'PLACEHOLDER' +
    'overviews. Iterative approaches to corporate strategy foster collaborative thinking ' +
    'to further the overall value proposition. Organically grow the holistic world view ' +
    'of disruptive.',
    { x: 1.252, y: 3.76, w: 6.012, h: 1.291, fontSize: 14, lineSpacingMultiple: 1.3 });
  text(s, '- Leverage agile frameworks to provide a robust synopsis',
    { x: 1.252, y: 5.51, w: 6.012, h: 0.382, fontSize: 14, bold: true });
}

function slide03(s) {
  const cards = [
    { rule: 1.193, tx: 1.531, num: '01.', title: "App's appearance", accent: C.blue },
    { rule: 5.056, tx: 5.394, num: '02.', title: 'Custom', accent: C.blueBright, card: true },
    { rule: 8.919, tx: 9.256, num: '03.', title: 'Cash learn', accent: C.blueDeep },
  ];
  cards.forEach((c) => {
    if (c.card) roundRect(s, 4.556, 3.625, 4.056, 2.958, 0.25, { fill: fill(C.panel, 97) });
    line(s, c.rule, 3.916, 0, 2.376, { line: { color: C.smoke, width: 1 } });
    line(s, c.rule, 5.222, 0, 0.731, { line: { color: c.accent, width: 3 } });
    text(s, c.num, { x: c.tx, y: 4.186, w: 0.443, h: 0.315, fontSize: 10.5, bold: true });
    text(s, c.title, { x: c.tx, y: 4.531, w: 2.837, h: 0.423, fontSize: 16, bold: true });
    text(s, 'Leverage agile frameworks to provide a robust synopsis for high level overviews. ',
      { x: c.tx, y: 5.143, w: 2.837, h: 0.995, fontSize: 14, lineSpacingMultiple: 1.3 });
  });
  text(s, 'Leverage agile frameworks to provide a robust synopsis for high level overviews. ' +
    'Iterative approaches proposition. Organically grow the holistic world view of disruptive.',
    { x: 2.188, y: 2.413, w: 8.958, h: 0.689, fontSize: 14, align: 'center', lineSpacingMultiple: 1.3 });
  heading(s, 'Find The Best Financial', { x: 1.615, y: 0.944, w: 10.103, h: 0.774, align: 'center' });
}

function slide04(s) {
  heading(s, 'Get Every Single Update',
    { x: 3.292, y: 1.127, w: 6.75, h: 0.64, fontSize: 32, align: 'center' });
  text(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet ' +
    'mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm.',
    { x: 2.286, y: 2.004, w: 8.762, h: 0.598, fontSize: 12, align: 'center', lineSpacingMultiple: 1.3 });

  photo(s, 0, 3.016, 6.625, 4.484);
  rect(s, 6.771, 3.016, 6.562, 4.484, { fill: fill(C.blue) });

  text(s, 'Your Text Goes Here',
    { x: 7.473, y: 3.638, w: 4.213, h: 0.404, fontSize: 18, bold: true, color: C.white });
  text(s, [
    { text: 'Ut wisi enim ad minim veniam', options: { bold: true } },
    { text: ', quis nostrud exerci tation ullamcorper suscipit lobortis nisl ut aliquip ' +
        'ex ea commodo consequat. Lorem ipsum dolor sit amet, consectetuer adipiscing' },
  ], { x: 7.473, y: 4.189, w: 5.296, h: 0.805, fontSize: 12, color: C.white, lineSpacingMultiple: 1.2 });
  text(s, '\u201cUt wisi enim ad minim veniam, quisnostrud exerci\ntation ullamcorper',
    { x: 7.473, y: 5.265, w: 4.447, h: 0.624, fontSize: 14, italic: true, color: C.white, lineSpacingMultiple: 1.0 });
  text(s, '21',
    { x: 11.916, y: 5.782, w: 0.854, h: 0.774, fontSize: 40, bold: true, color: C.white, align: 'right' });
  text(s, 'Monday, Apr 2023',
    { x: 10.706, y: 6.491, w: 2.064, h: 0.37, fontSize: 16, color: C.white, align: 'right' });
}

function slide05(s) {
  photo(s, 7.227, 1.295, 2.457, 4.91);
  // Half-disc bulging right, flat edge at x = 9.822.
  pie(s, 7.365, 1.295, 4.91, [270, 90], { fill: fill(C.blue) });
  text(s, '86%', { x: 10.068, y: 2.687, w: 1.792, h: 0.829, fontSize: 36, color: C.white, valign: 'bottom' });
  text(s, UTWISI,
    { x: 10.068, y: 3.516, w: 1.977, h: 0.985, fontSize: 14, color: C.white, lineSpacingMultiple: 1.3 });

  heading(s, 'Start Your Idea From Here', { x: 1.938, y: 1.493, w: 4.891, h: 1.447 });
  [{ n: '01', y: 3.71, color: C.blue }, { n: '02', y: 5.137, color: C.blueBright }].forEach((row) => {
    ellipse(s, 2.021, row.y, 0.601, 0.601, { fill: fill(row.color) });
    text(s, row.n, { x: 2.021, y: row.y + 0.146, w: 0.601, h: 0.37, fontSize: 20, color: C.white, align: 'center' });
    text(s, 'Your Text Goes Here',
      { x: 2.795, y: row.y + 0.023, w: 2.941, h: 0.316, fontSize: 16, bold: true, lineSpacingMultiple: 0.8 });
    text(s, UTWISI + ' ullamcorper',
      { x: 2.795, y: row.y + 0.322, w: 3.706, h: 0.684, fontSize: 14, lineSpacingMultiple: 1.3 });
  });
}

function slide06(s) {
  photo(s, 10.021, 1.375, 2.507, 3.167);
  photo(s, 7.306, 3.681, 2.507, 2.928);
  heading(s, 'Product Overview', { x: 1.94, y: 1.693, w: 5.151, h: 1.313, lineSpacingMultiple: 0.9 });
  [
    { letter: 'A', label: 'Good', y: 3.741, color: C.blue },
    { letter: 'B', label: 'Excellent', y: 5.28, color: C.blueBright },
  ].forEach((row) => {
    ellipse(s, 2.083, row.y, 0.629, 0.629, { fill: fill(row.color) });
    text(s, row.letter,
      { x: 2.083, y: row.y + 0.115, w: 0.629, h: 0.478, fontSize: 28, color: C.white, align: 'center' });
    text(s, row.label,
      { x: 2.863, y: row.y + 0.059, w: 2.698, h: 0.343, fontSize: 18, bold: true, lineSpacingMultiple: 0.8 });
    text(s, UTWISI, { x: 2.863, y: row.y + 0.375, w: 3.165, h: 0.684, fontSize: 14, lineSpacingMultiple: 1.3 });
  });
}

function slide07(s) {
  photo(s, 1.048, 1.212, 5.238, 5.238);
  ['60%', '40%', '30%'].forEach((pct, i) => {
    const y = 1.542 + i * 1.4645;
    text(s, pct, { x: 6.288, y, w: 2.297, h: 0.841, fontSize: 44, align: 'right' });
    text(s, 'Your Text Goes Here', { x: 8.694, y: y + 0.084, w: 3.423, h: 0.37, fontSize: 16, bold: true });
    text(s, UTWISI + ' ullamcorper',
      { x: 8.694, y: y + 0.455, w: 3.704, h: 0.643, fontSize: 14, lineSpacingMultiple: 1.2 });
  });
}

function slide08(s) {
  photo(s, 6.912, 1.124, 5.563, 5.513);
  heading(s, 'Competitive Advantage Thought People',
    { x: 1.554, y: 1.621, w: 5.858, h: 1.919, lineSpacingMultiple: 0.9 });
  text(s, 'Ut wisi enim ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis.',
    { x: 1.554, y: 3.873, w: 4.448, h: 0.684, fontSize: 14, lineSpacingMultiple: 1.3 });
  [{ x: 1.575, ix: 1.666 }, { x: 3.775, ix: 3.869 }].forEach((col) => {
    icon(s, col.ix, 4.988, 0.41, C.smoke);
    text(s, '20.000', { x: col.x, y: 5.505, w: 1.814, h: 0.404, fontSize: 18 });
    text(s, 'Ut wisi enim ad', { x: col.x, y: 5.844, w: 1.814, h: 0.36, fontSize: 14, lineSpacingMultiple: 1.2 });
  });
}

function slide09(s) {
  heading(s, 'Extraordinary Team Fintech',
    { x: 0.667, y: 2.242, w: 4.255, h: 1.313, lineSpacingMultiple: 0.9 });
  text(s, 'Awesome Team',
    { x: 0.667, y: 3.824, w: 2.534, h: 0.343, fontSize: 16, bold: true, lineSpacingMultiple: 0.9 });
  text(s, SERENITY, { x: 0.667, y: 4.281, w: 3.852, h: 0.978, fontSize: 12, lineSpacingMultiple: 1.1 });

  [
    { x: 5.125, name: 'Maja Osvald', role: 'Head Manager Area', handle: '@majaosvald', color: C.blueSky },
    { x: 7.999, name: 'Simon Piper', role: 'Head Marketing', handle: '@simonpiper', color: C.blue },
    { x: 10.872, name: 'Adelina Munro', role: 'Assist. Manager', handle: '@adelinamunro', color: C.periwinkle },
  ].forEach((m) => {
    photo(s, m.x - 0.005, 1.607, 2.68, 2.56);
    rect(s, m.x, 4.167, 2.675, 1.723, { fill: fill(m.color) });
    text(s, m.name, { x: m.x + 0.219, y: 4.322, w: 2.237, h: 0.374, color: C.white, lineSpacingMultiple: 0.9 });
    text(s, m.role,
      { x: m.x + 0.219, y: 4.63, w: 2.217, h: 0.283, fontSize: 12, italic: true, color: C.white, lineSpacingMultiple: 0.9 });
    icon(s, m.x + 0.323, 5.031, 0.134, C.white);
    icon(s, m.x + 0.525, 5.033, 0.13, C.white);
    text(s, m.handle,
      { x: m.x + 0.258, y: 5.252, w: 1.2, h: 0.252, fontSize: 10, color: C.white, lineSpacingMultiple: 0.9 });
    icon(s, m.x + 2.19, 5.505, 0.245, C.white);
  });
}

function slide10(s) {
  heading(s, 'Never Leave Your Team',
    { x: 0.667, y: 2.242, w: 3.852, h: 1.313, lineSpacingMultiple: 0.9 });
  text(s, 'Stay at Home',
    { x: 0.667, y: 3.824, w: 3.852, h: 0.343, fontSize: 16, bold: true, lineSpacingMultiple: 0.9 });
  text(s, SERENITY, { x: 0.667, y: 4.281, w: 3.852, h: 0.978, fontSize: 12, lineSpacingMultiple: 1.1 });

  [
    { x: 5.379, y: 4.164, photoY: 0.003, name: 'Edison Brandt', role: 'Speakers Presenter', color: C.blueBright },
    { x: 9.107, y: 5.258, photoY: 1.095, name: 'Daniele Banks', role: 'Commentator', color: C.periwinkle },
  ].forEach((m) => {
    photo(s, m.x, m.photoY, 2.536, 5.778);
    rect(s, m.x, m.y, 2.536, 1.617, { fill: fill(m.color) });
    icon(s, m.x + 0.199, m.y + 0.369, 0.285, C.white);
    text(s, m.name,
      { x: m.x + 0.084, y: m.y + 0.843, w: 1.802, h: 0.343, fontSize: 16, color: C.white, lineSpacingMultiple: 0.9 });
    text(s, m.role,
      { x: m.x + 0.084, y: m.y + 1.186, w: 1.802, h: 0.283, fontSize: 12, italic: true, color: C.white, lineSpacingMultiple: 0.9 });
  });
}

function slide11(s) {
  rect(s, 0, 1.542, W, 2.136, { fill: fill(C.blueBright) });
  heading(s, 'Quidel Moacir',
    { x: 8.566, y: 1.788, w: 3.337, h: 1.313, color: C.white, lineSpacingMultiple: 0.9 });
  text(s, 'He Story',
    { x: 8.692, y: 3.29, w: 2.534, h: 0.343, fontSize: 16, color: C.white, lineSpacingMultiple: 0.9 });
  text(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet ' +
    'mornings of spring which I enjoy with my whole heart. I am alone.',
    { x: 8.692, y: 3.766, w: 3.234, h: 0.978, fontSize: 12, lineSpacingMultiple: 1.1 });

  line(s, 1.736, 5.395, 11.597, 0, { line: { color: C.periwinkle, width: 1.25 } });
  ['2020', '2021', '2022', '2023'].forEach((year, i) => {
    const x = 1.462 + i * 2.803;
    ellipse(s, x + 0.06, 5.268, 0.254, 0.254, { fill: fill(C.blue, 79) });
    roundRect(s, x + 0.099, 5.307, 0.175, 0.176, 0.088, { fill: fill(C.blue) });
    roundRect(s, x, 5.73, 0.845, 0.404, 0.2, { fill: fill(C.periwinkle) });
    text(s, year,
      { x, y: 5.73, w: 0.845, h: 0.404, fontSize: 14, color: C.white, align: 'center', valign: 'middle' });
    text(s, 'A wonderful serenity has taken possession.',
      { x, y: 6.247, w: 2.055, h: 0.533, fontSize: 12, lineSpacingMultiple: 1.1 });
  });
}

function slide12(s) {
  // A wide translucent band across the lower half (a rotated rectangle in the source).
  rect(s, 0, 3.55, 12.667, 3.325, { fill: fill(C.smoke, 65) });
  heading(s, 'Featured Or Points Slide In Fintech',
    { x: 4.589, y: 1.292, w: 5.077, h: 1.5, fontSize: 36, lineSpacingMultiple: 1.2 });
  ['01.', '02.', '03.'].forEach((num, i) => {
    const x = 4.589 + i * 3.0015;
    heading(s, num, { x, y: 3.947, w: 1.503, h: 0.659, fontSize: 28, color: C.blue, lineSpacingMultiple: 1.3 });
    text(s, 'Add Your Project Title Here.',
      { x, y: 4.729, w: 1.55, h: 0.602, fontSize: 12, bold: true, lineSpacingMultiple: 1.3 });
    text(s, LOREM_SHORT, { x, y: 5.405, w: 2.075, h: 0.865, fontSize: 12, lineSpacingMultiple: 1.3 });
  });
}

function slide13(s) {
  heading(s, 'Second Featured Or Points Slide In Fintech',
    { x: 0.667, y: 0.857, w: 6.829, h: 1.653, lineSpacingMultiple: 1.2 });
  text(s, 'Add Your Project Title Here.',
    { x: 0.667, y: 2.849, w: 2.895, h: 0.372, fontSize: 14, bold: true, lineSpacingMultiple: 1.3 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
    'incididunt ut labore et dolore magna aliqua. Ut enim ad.',
    { x: 0.667, y: 3.332, w: 2.403, h: 1.39, fontSize: 12, lineSpacingMultiple: 1.3 });

  ['01.', '02.', '03.'].forEach((num, i) => {
    const x = 3.667 + i * 3;
    rect(s, x, 2.849, 2.171, 4.026, { fill: fill(C.blue) });
    icon(s, x + 0.194, 3.724, 0.5, C.white);
    text(s, num,
      { x: x + 1.476, y: 2.959, w: 0.589, h: 0.42, fontSize: 16, color: C.white, align: 'right', lineSpacingMultiple: 1.3 });
    text(s, 'Add Your Project Title Here.',
      { x: x + 0.092, y: 4.647, w: 1.655, h: 0.602, fontSize: 12, bold: true, color: C.white, lineSpacingMultiple: 1.3 });
    text(s, LOREM_SHORT,
      { x: x + 0.092, y: 5.323, w: 1.987, h: 0.865, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3 });
  });
}

function slide14(s) {
  rect(s, 0, 0, 3.688, H, { fill: fill(C.blue) });
  // Smartphone mockup (from the slide layout).
  roundRect(s, 1.938, 1.389, 3.482, 6.111, 0.42, { fill: fill(C.white) });
  roundRect(s, 2.153, 1.605, 3.051, 5.68, 0.3, { fill: fill('F2F4F3') });
  rect(s, 5.385, 3.032, 0.096, 0.825, { fill: fill(C.silver) });
  rect(s, 1.876, 2.838, 0.131, 0.542, { fill: fill(C.silver) });
  rect(s, 1.872, 3.484, 0.131, 0.542, { fill: fill(C.silver) });
  rect(s, 1.886, 2.319, 0.103, 0.31, { fill: fill(C.silver) });
  photo(s, 2.153, 1.605, 3.051, 5.68, { fill: 'A6A6A6', transparency: 88, label: 'Drag and Drop Image Here' });

  heading(s, 'Mockup Smartphone', { x: 6.667, y: 1.684, w: 4.379, h: 1.447, lineSpacingMultiple: 1.0 });
  ellipse(s, 6.667, 3.297, 0.643, 0.643, { fill: fill(C.blue) });
  text(s, '01',
    { x: 6.667, y: 3.297, w: 0.643, h: 0.643, fontSize: 16, color: C.white, align: 'center', valign: 'middle' });
  text(s, 'Text in Here', { x: 7.516, y: 3.401, w: 3.146, h: 0.438, fontSize: 20 });
  text(s, 'User experience buyer bandwidth A/B testing ecosystem non-disclosure agreement ' +
    'hypo theses from  network effects series.\n\nA financing deployment agile development. ' +
    'Seed round creative holy grail other.',
    { x: 6.667, y: 4.179, w: 4.031, h: 1.796, fontSize: 14, lineSpacingMultiple: 1.32 });
}

function slide15(s) {
  rect(s, 0, 0, W, 3.75, { fill: fill(C.blue) });
  // Browser mockup (from the slide layout).
  rect(s, 4.797, 1.27, 6.995, 0.379, { fill: fill(C.charcoal) });
  [C.blue, C.blueBright, C.blueDeep].forEach((c, i) =>
    ellipse(s, 4.985 + i * 0.121, 1.41, 0.077, 0.077, { fill: fill(c) }));
  roundRect(s, 5.528, 1.362, 2.491, 0.172, 0.086, { fill: fill(C.white) });
  text(s, 'www.website.com',
    { x: 5.528, y: 1.362, w: 2.491, h: 0.172, fontSize: 8, color: C.ash, align: 'center', valign: 'middle' });
  photo(s, 4.797, 1.649, 6.995, 3.935, { fill: 'A6A6A6', transparency: 88, label: 'Drag and Drop Image Here' });

  rect(s, 1.542, 2.149, 4.236, 4.051, { fill: fill(C.white), line: { color: C.silver, width: 1 } });
  icon(s, 1.955, 2.624, 0.872, C.ash);
  text(s, '01', { x: 5.064, y: 2.355, w: 0.43, h: 0.337, fontSize: 14, color: C.ink, align: 'right' });
  text(s, 'Web Mockup',
    { x: 1.955, y: 3.83, w: 3.822, h: 0.774, fontSize: 40, color: C.charcoal, lineSpacingMultiple: 1.0 });
  text(s, 'User experience buyer bandwidth testing ecosystem non-disclosure agreement hypotheses.',
    { x: 2.065, y: 4.764, w: 3.131, h: 0.87, fontSize: 12, color: C.ink, lineSpacingMultiple: 1.32 });

  rect(s, 8.573, 4.983, 3.219, 0.6, { fill: fill(C.blueBright) });
  text(s, 'Research progress',
    { x: 8.573, y: 4.983, w: 3.219, h: 0.6, fontSize: 20, color: C.white, align: 'center', valign: 'middle' });
}

function slide16(s) {
  rect(s, 0, 0, 6.667, H, { fill: fill(C.blue) });
  // Tablet mockup (from the slide layout).
  roundRect(s, 1.03, 1.116, 7.764, 5.26, 0.3, { fill: fill(C.white), line: { color: C.ash, width: 1 } });
  roundRect(s, 1.622, 1.305, 6.563, 4.881, 0.04, { fill: fill(C.white), line: { color: C.ash, width: 1 } });
  ellipse(s, 1.313, 3.697, 0.097, 0.097, { fill: fill(C.white), line: { color: C.ash, width: 1 } });
  ellipse(s, 8.284, 3.573, 0.344, 0.344, { fill: fill(C.white), line: { color: C.ash, width: 1 } });
  photo(s, 1.622, 1.305, 6.546, 4.873, { fill: C.silver, transparency: 90 });

  [
    { n: '01.', label: 'Mobile\ninterface', y: 1.321 },
    { n: '02.', label: 'Web &\nmockup', y: 3.131 },
    { n: '03.', label: 'Other\ninterface', y: 5.005 },
  ].forEach((row) => {
    ellipse(s, 9.763, row.y + 0.295, 0.633, 0.633, { fill: fill(C.blueBright) });
    icon(s, 9.93, row.y + 0.478, 0.3, C.white);
    text(s, row.n, { x: 10.601, y: row.y, w: 0.669, h: 0.349, fontSize: 11 });
    text(s, row.label, { x: 10.601, y: row.y + 0.349, w: 1.8, h: 0.707, fontSize: 18, lineSpacingMultiple: 1.0 });
  });
  line(s, 9.533, 2.816, 2.617, 0, { line: { color: C.smoke, width: 0.75 } });
  line(s, 9.533, 4.684, 2.617, 0, { line: { color: C.smoke, width: 0.75 } });
}

function slide17(s) {
  text(s, '02', {
    x: 9.181, y: 4.587, w: 3.667, h: 2.849,
    fontSize: 138, color: C.silver, transparency: 90, align: 'right', lineSpacingMultiple: 1.3,
  });
  text(s, 'Session', {
    x: 10.912, y: 4.892, w: 1.87, h: 0.459,
    fontSize: 18, color: C.silver, transparency: 90, align: 'right', lineSpacingMultiple: 1.3,
  });
  rect(s, 0, 2.744, 6.667, 4.756, { fill: fill(C.blue) });
  text(s, 'Break Time\nUntil 02:00 AM',
    { x: 0.667, y: 0.625, w: 2.403, h: 0.809, fontSize: 16, lineSpacingMultiple: 1.3 });
  text(s, '2023',
    { x: 11.288, y: 0.625, w: 1.378, h: 0.379, fontSize: 14, color: C.slate, align: 'right', lineSpacingMultiple: 1.3 });
  text(s, 'Session 02',
    { x: 4.796, y: 1.606, w: 1.87, h: 0.489, fontSize: 20, align: 'right', lineSpacingMultiple: 1.3 });
  text(s, 'Javier Kandiewald',
    { x: 4.796, y: 2.065, w: 1.87, h: 0.372, fontSize: 14, align: 'right', lineSpacingMultiple: 1.3 });
  heading(s, 'Break Time Section',
    { x: 6.819, y: 2.388, w: 5.847, h: 2.432, fontSize: 60, lineSpacingMultiple: 1.2 });
}

function slide18(s) {
  heading(s, 'Process', { x: 1.117, y: 0.956, w: 3.7, h: 0.774 });

  // The white ribbon threads behind the chevrons; each chevron's notch lets it
  // read as an arrowhead.
  [
    [4.889, 1.874, 7.698, 0.568], [2.326, 3.328, 2.876, 0.568], [0.018, 4.778, 2.876, 0.568],
    [4.880, 1.879, 0.568, 2.018], [2.326, 3.328, 0.568, 2.018],
  ].forEach(([x, y, w, h]) => rect(s, x, y, w, h, { fill: fill(C.smoke) }));

  // Two-tone chevrons: accent body with a darker leading wedge.
  const steps = [
    { x: 6.133, y: 1.313, dark: 5.933, body: C.blueDeep, edge: C.indigo, title: 'Your Tittle #3', tx: 5.799, ty: 3.463 },
    { x: 8.908, y: 1.313, dark: 8.708, body: C.blueSky, edge: C.royal, title: 'Your Tittle #4', tx: 8.331, ty: 3.463 },
    { x: 11.59, y: 1.313, dark: 11.39, body: C.violet, edge: C.grape, title: 'Your Tittle #5', tx: 11.008, ty: 3.463 },
    { x: 3.251, y: 2.765, dark: 3.049, body: C.blueBright, edge: C.cobalt, title: 'Your Tittle 2', tx: 2.94, ty: 4.778 },
    { x: 0.697, y: 4.217, dark: 0.492, body: C.blue, edge: C.navy, title: 'Your Tittle #1', tx: 0.056, ty: 2.585 },
  ];
  steps.forEach((st) => {
    s.addShape('chevron', { x: st.x, y: st.y, w: 1.753, h: 1.693, fill: fill(st.body), line: NOLINE });
    s.addShape('chevron', { x: st.dark, y: st.y, w: 1.055, h: 1.693, fill: fill(st.edge), line: NOLINE });
  });

  steps.forEach((st) => {
    icon(s, st.x + 0.65, st.y + 0.63, 0.44, C.white);
    text(s, st.title, { x: st.tx, y: st.ty, w: 2.063, h: 0.337, fontSize: 14, bold: true, color: st.body });
    text(s, 'Ut wisi enim ad minim veniam, quis nostrud exerct.',
      { x: st.tx + 0.23, y: st.ty + 0.335, w: 1.603, h: 0.9, fontSize: 12, align: 'center', lineSpacingMultiple: 1.2 });
  });
}

function slide19(s) {
  heading(s, 'Our Agenda Schedule', { x: 1.542, y: 1.717, w: 9.948, h: 0.774, align: 'center' });
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula ' +
    'eget dolor. Aenean massa.',
    { x: 3.542, y: 2.815, w: 6.25, h: 0.689, fontSize: 14, align: 'center', lineSpacingMultiple: 1.3 });
  roundRect(s, 1.284, 4.017, 10.765, 2.396, 0.21, { fill: fill(C.blue) });

  [
    { x: 1.542, time: '01.00 pm', body: [{ text: 'First we start the meeting and begin discuss about it' }], w: 2.689 },
    { x: 5.094, time: '05.00 pm', body: [
      { text: 'Second' },
      { text: ' bring the problem and try to select one of it', options: { bold: true } }], w: 2.689 },
    { x: 8.647, time: '10.00 pm', body: [{ text: 'Third make sure what we needs from the problem we get' }], w: 3.055 },
  ].forEach((it) => {
    line(s, it.x, 4.618, 1.958, 0, { line: { color: C.smoke, width: 1, transparency: 27 } });
    text(s, it.time,
      { x: it.x + 1.799, y: 4.437, w: 1.237, h: 0.365, fontSize: 14, align: 'right', lineSpacingMultiple: 1.2 });
    text(s, it.body, { x: it.x + 0.347, y: 4.983, w: it.w, h: 0.572, fontSize: 14, lineSpacingMultiple: 1.0 });
  });
}

function slide20(s) {
  [
    { x: 1.542, y: 0.958, label: 'Project started', color: C.blue },
    { x: 5.058, y: 0.958, label: 'Meeting', color: C.blueBright },
    { x: 8.575, y: 0.958, label: 'Design make', color: C.blue },
    { x: 1.542, y: 3.901, label: 'More research', color: C.blueBright },
    { x: 5.058, y: 3.901, label: 'Fixing and finishing', color: C.blue },
    { x: 8.575, y: 3.901, label: 'Market analytics', color: C.blueBright },
  ].forEach((t) => {
    const h = t.y > 2 ? 2.599 : 2.641;
    photo(s, t.x, t.y, 3.216, h, { fill: 'A6A6A6', transparency: 88, label: 'Drag and Drop Image Here' });
    const barY = t.y + h - 0.6;
    rect(s, t.x, barY, 3.217, 0.6, { fill: fill(t.color) });
    text(s, t.label,
      { x: t.x, y: barY, w: 3.217, h: 0.6, fontSize: 20, color: C.white, align: 'center', valign: 'middle' });
  });
}

function slide21(s) {
  heading(s, 'Gallery Image Slide', { x: 1.02, y: 1.503, w: 5.333, h: 1.447 });
  [
    [0.923, 3.233, 3.907, 3.359], [5.005, 2.878, 2.843, 3.119],
    [7.984, 1.807, 2.148, 2.239], [7.984, 4.204, 2.148, 2.239],
    [10.263, 2.85, 2.148, 2.239],
  ].forEach(([x, y, w, h]) => photo(s, x, y, w, h));
}

function slide22(s) {
  photo(s, 0, 1.448, 3.389, 2.701);
  photo(s, 7.183, 0, 6.15, 4.149);
  gradient(s, 3.389, 4.149, 3.794, 3.351, '1F3D99', '0F1F4D', { cols: 14, rows: 12 });

  text(s, 'Our Journey \nFor Your Story', { x: 4.044, y: 1.875, w: 2.722, h: 0.646, lineSpacingMultiple: 0.9 });
  text(s, 'A wonderful serenity has taken possession of my entire soul, these sweet mornings',
    { x: 4.044, y: 2.595, w: 2.722, h: 0.755, fontSize: 12, lineSpacingMultiple: 1.1 });
  roundRect(s, 0.73, 5.004, 1.498, 0.336, 0.168, { fill: fill(C.blue) });
  text(s, 'Keep Healthy', {
    x: 0.73, y: 5.004, w: 1.498, h: 0.336,
    fontSize: 10, color: C.white, align: 'center', valign: 'middle', lineSpacingMultiple: 0.9,
  });
  text(s, 'A wonderful serenity taken possession of my entire soul, like these sweet',
    { x: 0.667, y: 5.462, w: 2.333, h: 0.755, fontSize: 12, lineSpacingMultiple: 1.1 });
  icon(s, 4.474, 5.227, 0.467, C.white);
  text(s, 'Always Connected', { x: 4.39, y: 5.776, w: 1.793, h: 0.646, color: C.white, lineSpacingMultiple: 0.9 });
  heading(s, 'Fintech Gallery', { x: 8.361, y: 5.194, w: 4.466, h: 0.707, lineSpacingMultiple: 0.9 });
  text(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet',
    { x: 8.361, y: 5.922, w: 3.794, h: 0.533, fontSize: 12, lineSpacingMultiple: 1.1 });
}

function slide23(s) {
  roundRect(s, 0.757, 2.417, 11.819, 4.139, 0.27, { fill: fill(C.white) });
  heading(s, 'Table Economic', { x: 0.757, y: 1.314, w: 5.927, h: 0.774 });

  const rows = [['Name', 'Change (%)', 'Ico date', 'Ico price', 'Curr. Price'].map((t) => ({
    text: t,
    options: { fill: fill(C.blue), color: C.white, bold: true, fontSize: 12, fontFace: FONT.head },
  }))];
  ['Bitcoin', 'Ethereum', 'Ripple', 'Zcash', 'Monero-XMR', 'Dash'].forEach((name, i) => {
    const band = i % 2 ? 'F2F2F2' : C.white;
    rows.push([name, '+3244.77%', '11/05/2019', '$0.001', '$0.004'].map((t) => ({
      text: t, options: { fill: fill(band), color: C.black, fontSize: 11 },
    })));
  });
  s.addTable(rows, {
    x: 1.169, y: 2.797, colW: [1.651, 1.651, 1.651, 1.651, 1.651],
    rowH: [0.373, 0.427, 0.427, 0.427, 0.427, 0.427, 0.427],
    align: 'center', valign: 'middle', fontFace: FONT.body,
    border: [{ type: 'solid', color: C.ash, pt: 0.25 }],
  });

  [
    { t: 'BIND', x: 5.687, bg: C.blue, color: C.white, size: 12 },
    { t: 'BITF', x: 6.448, bg: C.blueBright, color: C.white, size: 12 },
    { t: 'BITS', x: 7.259, bg: 'F7F7F8', color: C.slate, size: 11 },
    { t: 'BLEU', x: 8.020, bg: 'F7F7F8', color: C.slate, size: 11 },
    { t: 'BMEX', x: 8.731, bg: 'F7F7F8', color: C.slate, size: 11 },
  ].forEach((tag) => {
    roundRect(s, tag.x, 5.847, 0.677, 0.339, 0.08, { fill: fill(tag.bg) });
    text(s, tag.t, {
      x: tag.x, y: 5.847, w: 0.677, h: 0.339,
      fontSize: tag.size, color: tag.color, align: 'center', valign: 'middle',
    });
  });

  text(s, 'Change (%)', { x: 9.721, y: 4.204, w: 2.258, h: 0.404, fontSize: 18, bold: true, color: C.slate });
  text(s, 'The European languages are. Their separate existence is a myth. ',
    { x: 9.721, y: 4.739, w: 2.682, h: 0.995, fontSize: 14, color: C.slate, lineSpacingMultiple: 1.3 });
}

function slide24(s) {
  gradient(s, 8.869, 0, 4.464, H, '3D3D99', '1F1F4D', { cols: 10, rows: 16 });
  const flowing = 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings';
  const listed = ['A wonderful serenity', 'Taken possession of ', 'Entire soul, like these', 'Sweet mornings'];
  [
    { x: 0.667, y: 1.957, letter: 'S', color: C.blue, title: 'Strength', bullets: false },
    { x: 4.768, y: 1.957, letter: 'W', color: C.blueDeep, title: 'Weaknesses ', bullets: true },
    { x: 0.667, y: 4.015, letter: 'O', color: C.blueSky, title: 'Opportunities ', bullets: true },
    { x: 4.768, y: 4.015, letter: 'T', color: C.periwinkle, title: 'Threats', bullets: false },
  ].forEach((q) => {
    rect(s, q.x, q.y, 1.528, 1.528, { fill: fill(q.color) });
    text(s, q.letter, {
      x: q.x, y: q.y, w: 1.528, h: 1.528,
      fontSize: 54, color: C.white, align: 'center', valign: 'middle',
    });
    text(s, q.title,
      { x: q.x + 1.766, y: q.y + 0.099, w: 2.097, h: 0.343, fontSize: 16, bold: true, lineSpacingMultiple: 0.9 });
    const body = q.bullets ? listed.map((t) => ({ text: t, options: { bullet: true } })) : flowing;
    text(s, body, { x: q.x + 1.75, y: q.y + 0.43, w: 2.097, h: 1.0, fontSize: 12, lineSpacingMultiple: 1.1 });
  });
}

function slide25(s) {
  rect(s, 6.667, 0.911, 6.667, 5.678, { fill: fill(C.blue) });
  heading(s, 'This Sector Has Worldclass',
    { x: 1.208, y: 1.265, w: 5.118, h: 1.344, fontSize: 32, lineSpacingMultiple: 1.2 });
  text(s, 'Add Your Project Title Here.',
    { x: 1.208, y: 2.898, w: 2.683, h: 0.372, fontSize: 14, bold: true, lineSpacingMultiple: 1.3 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do\neiusmod tempor ' +
    'incididunt ut labore et dolore magna aliqua. ',
    { x: 1.208, y: 3.382, w: 5.141, h: 0.602, fontSize: 12, lineSpacingMultiple: 1.3 });
  [{ x: 1.208, pct: '73', label: 'Project A' }, { x: 3.891, pct: '53', label: 'Project B' }].forEach((k) => {
    text(s, [
      { text: k.pct, options: { fontSize: 36 } },
      { text: '%', options: { fontSize: 20 } },
    ], { x: k.x, y: 4.34, w: 1.396, h: 0.816, lineSpacingMultiple: 1.3 });
    text(s, k.label, { x: k.x, y: 5.225, w: 2.384, h: 0.372, fontSize: 14, bold: true, lineSpacingMultiple: 1.3 });
    text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.',
      { x: k.x, y: 5.633, w: 2.458, h: 0.602, fontSize: 12, lineSpacingMultiple: 1.3 });
  });

  [
    { gx: 7.385, gy: 1.457, tx: 8.438, ty: 1.469, value: '70', angleRange: [117, 320] },
    { gx: 7.385, gy: 3.204, tx: 8.438, ty: 3.215, value: '40', angleRange: [117, 244] },
    { gx: 7.324, gy: 4.95, tx: 8.377, ty: 4.961, value: '90', angleRange: [117, 46] },
  ].forEach((g) => {
    gauge(s, g.gx, g.gy, 0.965, {
      track: C.smoke, trackTransparency: 70, sweep: C.white, hub: C.blueSky,
      angleRange: g.angleRange, value: g.value, labelColor: C.white,
    });
    text(s, 'Add Your Project.',
      { x: g.tx, y: g.ty, w: 4.177, h: 0.34, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3 });
    text(s, 'Lorem ipsum dolor sit amet, lorem ip conselor ctetur adipiscing elit, Lorem dolo ' +
      'ipsum dolor sit lorem ips.',
      { x: g.tx, y: g.ty + 0.34, w: 4.177, h: 0.602, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3 });
  });
}

function slide26(s) {
  heading(s, 'The Tax And Accounting Department',
    { x: 2.167, y: 0.924, w: 8.722, h: 1.653, align: 'center', lineSpacingMultiple: 1.2 });

  [
    { x: 0.667, label: 'Project A', bars: C.blue, rule: C.blue, values: [4.3, 2.5, 3.5, 4.5] },
    { x: 5.167, label: 'Project B', bars: C.violet, rule: C.violet, values: [4.3, 3.0, 2.0, 5.0] },
    { x: 9.667, label: 'Project C', bars: C.blue, rule: C.blueDeep, values: [2.0, 4.0, 3.0, 4.5] },
  ].forEach((col) => {
    s.addChart('bar', [{
      name: 'Series 1',
      labels: ['Pro A', 'Pro B', 'Pro C', 'Pro D'],
      values: col.values,
    }], {
      x: col.x, y: 2.612, w: 3.0, h: 2.504,
      barDir: 'col', barGapWidthPct: 219, chartColors: [col.bars],
      showLegend: false, showValue: false,
      valAxisHidden: true, valGridLine: { color: '3A3A3A', size: 0.75 },
      catAxisLineColor: '262626',
      catAxisLabelColor: C.smoke, catAxisLabelFontFace: FONT.body, catAxisLabelFontSize: 12,
    });
    rect(s, col.x, 5.379, 3.0, 0.141, { fill: fill(col.rule) });
    text(s, col.label, { x: col.x, y: 5.623, w: 2.458, h: 0.34, fontSize: 12, lineSpacingMultiple: 1.3 });
    text(s, 'Lorem ipsum dolor sit amet, lorem ip consectetur adipiscing elit, sed do',
      { x: col.x, y: 6.031, w: 3.0, h: 0.602, fontSize: 12, lineSpacingMultiple: 1.3 });
  });
}

function slide27(s) {
  heading(s, 'Office Chart In Fintech',
    { x: 3.673, y: 1.042, w: 5.987, h: 0.698, fontSize: 32, align: 'center', lineSpacingMultiple: 1.2 });
  s.addChart('pie', [{
    name: 'Sales',
    labels: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'],
    values: [0.8, 0.2, 0.45, 0.32],
  }], {
    x: 3.8, y: 2.028, w: 5.732, h: 4.847,
    chartColors: [C.blue, C.violet, C.blueDeep, C.blueSky],
    dataBorder: { pt: 2.5, color: C.white },
    showLegend: false, showValue: true, dataLabelFormatCode: '0%',
    dataLabelColor: C.white, dataLabelFontFace: FONT.body, dataLabelFontSize: 20,
  });

  [
    { x: 0.672, y: 3.223, align: 'right', color: C.blueSky },
    { x: 1.348, y: 4.737, align: 'right', color: C.blueDeep },
    { x: 10.343, y: 3.223, align: 'left', color: C.blue },
    { x: 9.667, y: 4.737, align: 'left', color: C.violet },
  ].forEach((n) => {
    text(s, 'Add Your Project.',
      { x: n.x, y: n.y, w: 2.324, h: 0.34, fontSize: 12, color: n.color, align: n.align, lineSpacingMultiple: 1.3 });
    text(s, 'Lorem ipsum dolor sit amet, lorem ip conselor lorem ip',
      { x: n.x, y: n.y + 0.34, w: 2.324, h: 0.602, fontSize: 12, align: n.align, lineSpacingMultiple: 1.3 });
  });

  [
    { x: 4.952, y: 2.343, color: C.blueSky }, { x: 8.456, y: 3.223, color: C.blue },
    { x: 4.181, y: 4.802, color: C.blueDeep }, { x: 6.274, y: 6.284, color: C.violet },
  ].forEach((b) => {
    rect(s, b.x, b.y, 0.591, 0.591, { fill: fill(C.white) });
    icon(s, b.x + 0.073, b.y + 0.073, 0.444, b.color);
  });
}

function slide28(s) {
  heading(s, 'Pricing Table', { x: 2.515, y: 0.964, w: 8.303, h: 0.774, align: 'center' });
  photo(s, 3.067, 1.993, 3.25, 4.501);

  const perks = ['Box for Packing', 'Troelly for moving', 'Safe and Secure Delivery', 'Tape for fragile Item'];
  [
    { x: 0.745, y: 1.983, w: 2.786, h: 4.51, radius: 0.12, card: C.blue, ink: C.white, dot: C.white, tick: C.blue, name: 'Econom', price: '$34' },
    { x: 6.474, y: 1.984, w: 2.954, h: 4.56, radius: 0.16, card: C.white, ink: C.slate, accent: C.blueBright, dot: C.blueBright, tick: C.white, name: 'Premium', price: '$64' },
    { x: 9.582, y: 1.993, w: 2.898, h: 4.56, radius: 0.16, card: C.white, ink: C.slate, accent: C.blueDeep, dot: C.blueDeep, tick: C.white, name: 'VIP', price: '$84' },
  ].forEach((p) => {
    roundRect(s, p.x, p.y, p.w, p.h, p.radius, { fill: fill(p.card) });
    const cx = p.x + 0.279;
    text(s, p.name, { x: cx, y: p.y + 0.302, w: 2.379, h: 0.627, fontSize: 24, color: p.ink, lineSpacingMultiple: 1.5 });
    text(s, 'Subtitle goes here',
      { x: cx + 0.03, y: p.y + 0.801, w: 2.067, h: 0.352, fontSize: 11, color: p.ink, lineSpacingMultiple: 1.5 });
    text(s, p.price,
      { x: cx, y: p.y + 1.173, w: 1.892, h: 0.978, fontSize: 40, color: p.accent || p.ink, lineSpacingMultiple: 1.5 });
    text(s, '/month',
      { x: cx + 1.109, y: p.y + 1.581, w: 1.684, h: 0.452, fontSize: 16, color: p.accent || p.ink, lineSpacingMultiple: 1.5 });
    perks.forEach((perk, i) => {
      const y = p.y + 2.394 + i * 0.4835;
      ellipse(s, p.x + 0.407, y, 0.21, 0.21, { fill: fill(p.dot) });
      text(s, '\u2713', { x: p.x + 0.407, y, w: 0.21, h: 0.21, fontSize: 9, color: p.tick, align: 'center', valign: 'middle' });
      text(s, perk, { x: p.x + 0.7, y: y - 0.085, w: 2.1, h: 0.364, fontSize: 12, color: p.ink, lineSpacingMultiple: 1.5 });
    });
  });
}

function slide29(s) {
  // Six petals around a hub — each petal is a 52-degree pie slice.
  const cx = 6.61, cy = 4.04, r = 1.83;
  [C.blue, C.blueBright, C.blueDeep, C.blueSky, C.violet, C.periwinkle].forEach((color, i) => {
    const start = 214 + i * 60;
    pie(s, cx - r, cy - r, r * 2, [start % 360, (start + 52) % 360], { fill: fill(color) });
    const mid = ((start + 26) % 360) * Math.PI / 180;
    icon(s, cx + Math.cos(mid) * r * 0.72 - 0.19, cy + Math.sin(mid) * r * 0.72 - 0.19, 0.38, C.white);
  });
  ellipse(s, cx - 0.97, cy - 0.97, 1.94, 1.94, { fill: fill(C.black) });
  ellipse(s, cx - 0.877, cy - 0.875, 1.753, 1.75, { fill: fill(C.blue) });
  icon(s, cx - 0.27, cy - 0.62, 0.54, C.white);
  text(s, 'Target',
    { x: cx - 0.765, y: cy - 0.023, w: 1.547, h: 0.419, fontSize: 16, color: C.white, align: 'center', lineSpacingMultiple: 1.3 });

  [
    { x: 1.555, y: 1.551, align: 'right' }, { x: 8.418, y: 1.551, align: 'left' },
    { x: 0.752, y: 3.683, align: 'right' }, { x: 9.326, y: 3.683, align: 'left' },
    { x: 1.708, y: 5.826, align: 'right' }, { x: 8.155, y: 5.826, align: 'left' },
  ].forEach((n) => {
    text(s, UTWISI, { x: n.x, y: n.y, w: 3.256, h: 0.596, fontSize: 12, align: n.align, lineSpacingMultiple: 1.3 });
  });
}

/** One quarter-ellipse pinwheel blade rooted at (hx, hy). */
function slide30Blade(s, hx, hy, rx, ry, range, color) {
  s.addShape('pie', {
    x: hx - rx, y: hy - ry, w: rx * 2, h: ry * 2,
    angleRange: range, fill: fill(color), line: NOLINE,
  });
}

function slide30(s) {
  heading(s, 'We Become The Driving Wind', { x: 1.319, y: 1.213, w: 10.695, h: 0.774, align: 'center' });

  // Pinwheel on a grey mast. Each blade is a quarter ellipse springing from the
  // hub, with a darker quarter nested at its root.
  rect(s, 6.599, 4.602, 0.136, 2.893, { fill: fill(C.silver) });
  const hx = 6.681, hy = 4.588;
  [
    { rx: 0.946, ry: 1.974, range: [180, 270], color: C.blue, root: C.navy, rootQ: [180, 270] },
    { rx: 1.974, ry: 0.955, range: [270, 360], color: C.blueBright, root: C.cobalt, rootQ: [270, 360] },
    { rx: 0.950, ry: 1.967, range: [0, 90], color: C.blueDeep, root: C.indigo, rootQ: [0, 90] },
    { rx: 1.974, ry: 0.946, range: [90, 180], color: C.blueSky, root: C.royal, rootQ: [90, 180] },
  ].forEach((b) => {
    slide30Blade(s, hx, hy, b.rx, b.ry, b.range, b.color);
    slide30Blade(s, hx, hy, 0.95, 0.95, b.rootQ, b.root);
  });
  ellipse(s, hx - 0.18, hy - 0.18, 0.361, 0.361, { fill: fill(C.smoke) });
  ellipse(s, hx - 0.085, hy - 0.082, 0.17, 0.164, { fill: fill('6D7381') });

  [
    { year: '2020', color: C.blue, badge: [4.388, 2.816], head: [2.518, 3.243], body: [2.282, 3.649], align: 'right' },
    { year: '2021', color: C.blueSky, badge: [3.77, 5.292], head: [1.89, 5.719], body: [1.656, 6.073], align: 'right' },
    { year: '2022', color: C.blueBright, badge: [9.023, 3.058], head: [8.997, 3.485], body: [8.997, 3.839], align: 'left' },
    { year: '2023', color: C.blueDeep, badge: [8.179, 5.218], head: [8.153, 5.646], body: [8.153, 5.999], align: 'left' },
  ].forEach((m) => {
    roundRect(s, m.badge[0], m.badge[1], 0.95, 0.291, 0.145, { fill: fill(m.color) });
    text(s, m.year, {
      x: m.badge[0], y: m.badge[1], w: 0.95, h: 0.291,
      fontSize: 14, color: C.white, align: 'center', valign: 'middle',
    });
    text(s, 'Ut wisi enim ad minim',
      { x: m.head[0], y: m.head[1], w: 2.82, h: 0.337, fontSize: 14, align: m.align, lineSpacingMultiple: 1.1 });
    text(s, 'Ut wisi enim ad minim veniam, quis nostrud exerci',
      { x: m.body[0], y: m.body[1], w: 3.03, h: 0.596, fontSize: 12, align: m.align, lineSpacingMultiple: 1.3 });
  });
}

function slide31(s) {
  heading(s, 'Make Powerful Market To Project', { x: 1.829, y: 1.317, w: 5.83, h: 1.447 });
  text(s, 'Ut wisi enim ad minim veniam, quis nostrud exerci tation ullamcorper',
    { x: 1.887, y: 3.071, w: 4.253, h: 0.643, fontSize: 14, lineSpacingMultiple: 1.2 });

  // Four nested quarter discs, largest first.
  [
    { x: 7.575, y: 1.404, d: 4.693, range: [269.7, 360], color: C.blueSky },
    { x: 7.763, y: 1.591, d: 4.317, range: [0.1, 89.9], color: C.blueDeep },
    { x: 7.95, y: 1.779, d: 3.942, range: [89.6, 179.6], color: C.blueBright },
    { x: 8.185, y: 2.014, d: 3.473, range: [179.9, 270], color: C.blue },
  ].forEach((ring) => {
    s.addShape('pie', {
      x: ring.x, y: ring.y, w: ring.d, h: ring.d, angleRange: ring.range,
      fill: fill(ring.color), line: { color: C.white, width: 2 },
    });
  });
  [
    { big: '12.5B', small: '2026', x: 9.968, y: 2.389, w: 1.924, size: 32, sub: 14 },
    { big: '285M', small: '2020', x: 8.276, y: 2.765, w: 1.693, size: 20, sub: 10.5 },
    { big: '519M', small: '2022', x: 8.185, y: 4.214, w: 1.693, size: 24, sub: 11 },
    { big: '830M', small: '2024', x: 10.059, y: 4.234, w: 1.693, size: 28, sub: 12 },
  ].forEach((k) => {
    text(s, k.big, { x: k.x, y: k.y, w: k.w, h: k.size / 45, fontSize: k.size, color: C.white, align: 'center' });
    text(s, k.small, {
      x: k.x, y: k.y + k.size / 45 + 0.03, w: k.w, h: 0.3,
      fontSize: k.sub, color: C.white, align: 'center',
    });
  });
  ellipse(s, 9.429, 3.257, 0.985, 0.985, { fill: fill(C.white) });
  icon(s, 9.715, 3.543, 0.413, C.slate);

  [
    { x: 1.887, y: 4.25, dot: C.blue, label: 'UI Design', pct: '38%', lw: 1.641 },
    { x: 4.169, y: 4.25, dot: C.blueBright, label: 'Test User', pct: '6%', lw: 1.885 },
    { x: 1.887, y: 5.235, dot: C.blueDeep, label: 'Web Design', pct: '24%', lw: 1.823 },
    { x: 4.169, y: 5.235, dot: C.blueSky, label: 'UI/UX Interface', pct: '11%', lw: 2.277 },
  ].forEach((l) => {
    ellipse(s, l.x, l.y + 0.131, 0.169, 0.168, { fill: fill(l.dot) });
    text(s, l.label, { x: l.x + 0.184, y: l.y, w: l.lw, h: 0.401, fontSize: 16, lineSpacingMultiple: 1.2 });
    text(s, l.pct, { x: l.x + 0.184, y: l.y + 0.297, w: 1.6, h: 0.549, fontSize: 24, bold: true, lineSpacingMultiple: 1.2 });
  });
}

function slide32(s) {
  mapMask(s, { x: 6.30, y: 0.99, w: 6.40, h: 4.60 }, ASIA_MAP, { '#': C.mapGrey });

  heading(s, 'Asian Mapping Business Marketing', { x: 1.118, y: 1.948, w: 4.948, h: 2.121 });
  ellipse(s, 1.172, 4.286, 1.623, 1.623, { fill: fill(C.blue) });
  s.addShape('arc', {
    x: 1.412, y: 4.526, w: 1.143, h: 1.143, angleRange: [270, 260],
    fill: { type: 'none' }, line: { color: C.black, width: 8, transparency: 90 },
  });
  s.addShape('arc', {
    x: 1.412, y: 4.526, w: 1.143, h: 1.143, angleRange: [270, 175],
    fill: { type: 'none' }, line: { color: C.white, width: 8 },
  });
  text(s, [
    { text: '75', options: { fontSize: 24 } },
    { text: '%', options: { fontSize: 10.5 } },
  ], { x: 1.172, y: 4.873, w: 1.623, h: 0.505, color: C.white, align: 'center' });
  ellipse(s, 2.48, 4.347, 0.357, 0.357, { fill: fill(C.blueBright) });
  icon(s, 2.548, 4.415, 0.22, C.white);
  text(s, 'A wonderful serenity has in possession of my entire soul, like these sweet mornings.',
    { x: 3.238, y: 4.47, w: 3.021, h: 0.97, fontSize: 12, lineSpacingMultiple: 1.5 });

  [
    { n: '01', x: 6.588, y: 3.532 }, { n: '02', x: 7.838, y: 2.04 },
    { n: '03', x: 9.091, y: 3.547 }, { n: '04', x: 9.407, y: 4.995 },
    { n: '05', x: 9.388, y: 1.603 }, { n: '06', x: 10.693, y: 1.547 },
  ].forEach((p) => {
    s.addShape('teardrop', { x: p.x, y: p.y, w: 0.502, h: 0.502, rotate: 135, fill: fill(C.blue), line: NOLINE });
    ellipse(s, p.x + 0.074, p.y + 0.074, 0.36, 0.36, { fill: fill(C.white) });
    text(s, p.n, { x: p.x + 0.038, y: p.y + 0.09, w: 0.431, h: 0.341, fontSize: 9, color: C.slate, align: 'center' });
  });
}

function slide33(s) {
  mapMask(s, { x: 0.667, y: 1.158, w: 5.381, h: 5.184 }, AUSTRALIA_MAP,
    { W: C.blue, G: C.silver, Q: C.blueSky, S: C.violet });

  [
    { x: 0.964, y: 1.854, value: '70', sweep: C.blue, ring: 'CCCCF5', angleRange: [117, 320] },
    { x: 4.901, y: 2.337, value: '80', sweep: C.blueSky, ring: 'D6E0FF', angleRange: [117, 3] },
    { x: 2.789, y: 4.288, value: '50', sweep: C.violet, ring: 'E0CCFF', angleRange: [89, 269] },
  ].forEach((g) => {
    ellipse(s, g.x, g.y, 1.147, 1.147, { fill: fill(C.white) });
    gauge(s, g.x + 0.091, g.y + 0.091, 0.965, {
      track: g.ring, trackTransparency: 70, sweep: g.sweep, hub: C.white,
      angleRange: g.angleRange, value: g.value, labelColor: g.sweep,
    });
  });

  heading(s, 'Australia Market And Maps On Fintech',
    { x: 6.674, y: 0.804, w: 5.451, h: 1.344, fontSize: 32, lineSpacingMultiple: 1.2 });
  [
    { n: '01.', color: C.blue, x: 6.674, y: 2.552 },
    { n: '02.', color: C.violet, x: 9.667, y: 2.552 },
    { n: '03.', color: C.blueSky, x: 6.674, y: 4.454 },
  ].forEach((item) => {
    text(s, item.n,
      { x: item.x, y: item.y, w: 1.449, h: 0.659, fontSize: 28, color: item.color, lineSpacingMultiple: 1.3 });
    text(s, 'Add Your Title Project Here.',
      { x: item.x + 0.807, y: item.y + 0.056, w: 1.572, h: 0.602, fontSize: 12, lineSpacingMultiple: 1.3 });
    text(s, 'Lorem ipsum dolor sit amet, lorem ip conselor ctetur adipiscing.',
      { x: item.x, y: item.y + 0.701, w: 2.715, h: 0.602, fontSize: 12, lineSpacingMultiple: 1.3 });
  });
}

function slide34(s) {
  heading(s, 'World Maps Balance', { x: 2.31, y: 0.486, w: 8.714, h: 0.774, align: 'center' });
  mapDots(s, { x: 1.489, y: 2.399, w: 10.355, h: 4.111 }, WORLD_MAP, C.silver);

  [1.856, 5.333, 8.625].forEach((x, i) => {
    const y = i === 1 ? 4.564 : 4.898;
    roundRect(s, x, y, 2.667, 1.027, 0.07, { fill: fill(C.white) });
    text(s, 'Accounting balance',
      { x: x + 0.223, y: y + 0.118, w: 1.916, h: 0.337, fontSize: 14, bold: true, color: C.slate });
    text(s, '\u2611 12 146,78 PLN',
      { x: x + 0.223, y: y + 0.534, w: 2.399, h: 0.37, fontSize: 20, color: C.slate, lineSpacingMultiple: 0.8 });
  });
}

function slide35(s) {
  rect(s, 0, 0, 9.028, H, { fill: fill(C.blue) });
  rect(s, 6.209, 0.807, 6.218, 5.885, { fill: fill(C.white) });
  photo(s, 6.408, 1.049, 5.822, 5.402, { fill: C.white });

  heading(s, 'Contact Us \nOn Our Website',
    { x: 1.161, y: 1.562, w: 4.889, h: 1.313, color: C.white, lineSpacingMultiple: 0.9 });
  text(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet ' +
    'mornings of spring which I enjoy with my whole heart. I am alone.',
    { x: 1.161, y: 3.258, w: 4.504, h: 0.755, fontSize: 12, color: C.white, lineSpacingMultiple: 1.1 });

  ['124 Green View Av., Canada', '+33 1234 5678 XXXX', 'mail@company.com', 'www.companywebsite.com']
    .forEach((row, i) => {
      const y = 4.403 + i * 0.3645;
      icon(s, 1.263, y + 0.05, 0.15, C.white);
      text(s, row, { x: 1.53, y, w: 2.737, h: 0.327, fontSize: 12, color: C.white, lineSpacingMultiple: 1.2 });
    });

  s.addShape('wedgeRoundRectCallout', {
    x: 5.665, y: 4.275, w: 3.622, h: 1.504, fill: fill(C.blueSky), line: NOLINE,
  });
  icon(s, 6.05, 4.784, 0.485, C.white);
  text(s, '124 Green View Av, Canada',
    { x: 6.733, y: 4.664, w: 2.135, h: 0.726, fontSize: 16, color: C.white, lineSpacingMultiple: 1.2 });
  text(s, '35', { x: 0.67, y: 6.875, w: 0.718, h: 0.342, fontSize: 12, fontFace: FONT.head, color: C.white });
}

function slide36(s) {
  text(s, 'Fintech.',
    { x: 2.568, y: 1.859, w: 2.677, h: 0.399, fontSize: 16, bold: true, align: 'right', lineSpacingMultiple: 1.2 });
  line(s, 5.339, 2.099, 2.481, 0, { line: { color: C.smoke, width: 1 } });
  heading(s, 'Thank You',
    { x: 1.667, y: 2.515, w: 10.0, h: 2.036, fontSize: 115, align: 'center', charSpacing: -3 });
  text(s, 'For watching', { x: 5.56, y: 4.67, w: 2.213, h: 0.404, fontSize: 18, charSpacing: 3 });
}

const SLIDES = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
  slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18,
  slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27,
  slide28, slide29, slide30, slide31, slide32, slide33, slide34, slide35, slide36,
];

/* ---------------------------------------------------------------------- main */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'FINTECH', width: W, height: H });
  pptx.layout = 'FINTECH';
  pptx.theme = { headFontFace: FONT.head, bodyFontFace: FONT.body };
  pptx.title = 'Fintech \u2014 Financial Presentation';

  SLIDES.forEach((builder) => {
    const slide = pptx.addSlide();
    slide.background = { color: C.black };
    builder(slide);
  });

  return pptx.writeFile({
    fileName: path.join(__dirname, '034fb550-734d-4f66-b3a9-bb3d9110e642_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
