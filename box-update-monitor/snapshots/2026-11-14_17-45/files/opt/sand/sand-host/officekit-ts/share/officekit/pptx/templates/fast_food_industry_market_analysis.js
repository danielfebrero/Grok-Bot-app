/**
 * "Frenzy" fast-food business deck - 22 slides, 13.333 x 7.5 in.
 * Rebuilt with pptxgenjs only. Photographs in the original are replaced by
 * flat "[image]" placeholder shapes that keep the original geometry.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const RED = 'D62021';
const RED_MID = 'B31115';
const RED_DARK = '9F0004';
const AMBER = 'FDA600';
const GOLD = 'FFC000';
const WHITE = 'FFFFFF';
const INK = '262626'; // headline dark
const INK_SOFT = '0C0C0C'; // small caps labels
const CHARCOAL = '2C2D32';
const GREY = '7F7F7F'; // body copy
const GREY_LT = 'BFBFBF';
const PINK = 'F6C6C6'; // body copy on red
const PINK_LT = 'F7CDCD';
const BG_TOP = 'EFEEED';
const BG_BOT = 'E8E5E4';
const PH_FILL = 'FBFCFE'; // image placeholder body
const PH_LINE = 'C8D4EC'; // image placeholder hatching
const PH_TEXT = '9FB0CC'; // image placeholder caption

const RED_GRAD = [[0, RED], [0.81, RED_MID], [1, RED_DARK]];
const PAGE_GRAD = [[0, BG_TOP], [1, BG_BOT]];

/* -------------------------------------------------------------------- fonts */
const XB = 'Merriweather Sans ExtraBold'; // headlines
const MD = 'Merriweather Sans Medium'; // small caps labels
const LT = 'Merriweather Sans Light'; // "LOGO"
const BODY = 'Lato'; // paragraphs & numbers
const UI = 'Poppins'; // url, marker digits

/* ------------------------------------------------------------ tiny helpers */
const mix = (a, b, t) => {
  const p = (h, i) => parseInt(h.substr(i, 2), 16);
  const c = i => Math.round(p(a, i) + (p(b, i) - p(a, i)) * t).toString(16).padStart(2, '0');
  return (c(0) + c(2) + c(4)).toUpperCase();
};

/** colour of a multi-stop linear gradient at position t (0..1) */
function gradAt(stops, t) {
  for (let i = 1; i < stops.length; i++) {
    if (t <= stops[i][0] || i === stops.length - 1) {
      const [p0, c0] = stops[i - 1];
      const [p1, c1] = stops[i];
      return mix(c0, c1, p1 === p0 ? 0 : (t - p0) / (p1 - p0));
    }
  }
  return stops[0][1];
}

/** pptxgenjs has no gradient fill, so paint one as a stack of solid bands */
function gradRect(s, x, y, w, h, stops, dir = 'v', steps = 26) {
  const span = (dir === 'v' ? h : w) / steps;
  for (let i = 0; i < steps; i++) {
    const c = gradAt(stops, (i + 0.5) / steps);
    const o = dir === 'v'
      ? { x, y: y + i * span, w, h: span + 0.012 }
      : { x: x + i * span, y, w: span + 0.012, h };
    s.addShape('rect', Object.assign(o, { fill: { color: c }, line: { type: 'none' } }));
  }
}

const page = s => gradRect(s, 0, 0, 13.333, 7.5, PAGE_GRAD, 'v', 8);
const redPanel = (s, x, y, w, h, dir = 'v') => gradRect(s, x, y, w, h, RED_GRAD, dir);

/**
 * Stand-in for a photo, matching the original picture frame: a pale panel
 * ruled with faint 45-degree lines, captioned "[image]" along the top edge.
 */
function pic(s, x, y, w, h) {
  s.addShape('rect', { x, y, w, h, fill: { color: PH_FILL }, line: { type: 'none' } });
  // rules run bottom-left to top-right along u + v = k, clipped to the panel
  const step = 0.15;
  for (let k = step; k < w + h; k += step) {
    const u0 = Math.max(0, k - h), u1 = Math.min(w, k);
    const v0 = Math.max(0, k - w), v1 = Math.min(h, k);
    if (u1 - u0 > 0.02) {
      s.addShape('line', {
        x: x + u0, y: y + v0, w: u1 - u0, h: v1 - v0, flipV: true,
        line: { color: PH_LINE, width: 0.75 },
      });
    }
  }
  s.addText('[image]', {
    x, y: y + 0.03, w, h: 0.28, align: 'center', valign: 'top',
    fontFace: UI, fontSize: 11, color: PH_TEXT,
  });
}

/** thin rotated bar - used for the cutlery glyph strokes */
function bar(s, cx, cy, len, thick, deg, color) {
  s.addShape('rect', {
    x: cx - len / 2, y: cy - thick / 2, w: len, h: thick, rotate: deg,
    fill: { color }, line: { type: 'none' },
  });
}

/** crossed fork + spoon mark used inside every logo badge */
function cutlery(s, x, y, d, color) {
  bar(s, x + d / 2, y + d / 2, d * 1.15, d * 0.12, 315, color);
  bar(s, x + d / 2, y + d / 2, d * 1.15, d * 0.12, 45, color);
  s.addShape('ellipse', {
    x: x + d * 0.58, y: y + d * 0.02, w: d * 0.4, h: d * 0.46, rotate: 45,
    fill: { color }, line: { type: 'none' },
  });
  bar(s, x + d * 0.2, y + d * 0.22, d * 0.42, d * 0.3, 45, color);
}

/**
 * Circular "frenzy" badge. (x,y) is the top-left of the big ring, d its
 * diameter; a small satellite ring sits above-right. The "LOGO" wordmark
 * defaults to the right of the badge; pass `at` to place it elsewhere.
 */
function logo(s, x, y, d, color, labelPt, at) {
  const ring = o => s.addShape('ellipse', Object.assign(o, { fill: { type: 'none' }, line: { color, width: 1.5 } }));
  ring({ x, y, w: d, h: d });
  ring({ x: x + d * 0.863, y: y - d * 0.27, w: d * 0.27, h: d * 0.27 });
  cutlery(s, x + d * 0.27, y + d * 0.29, d * 0.46, color);
  s.addText('LOGO', Object.assign({
    x: at ? at[0] : x + d * 1.215, y: at ? at[1] : y + d * 0.228,
    w: 0.99, h: 0.37, fontFace: LT, fontSize: labelPt, color, charSpacing: 3,
  }, TXT));
}

/* ------------------------------------------------------------ text helpers */
const TXT = { valign: 'top', wrap: false, isTextBox: true };

/** two-tone headline; `lines` is an array of arrays of [text, colour] runs */
function title(s, x, y, w, h, lines, size, align = 'left') {
  const runs = [];
  lines.forEach((line, i) => {
    line.forEach((run, j) => {
      runs.push({
        text: run[0],
        options: {
          color: run[1],
          breakLine: i < lines.length - 1 && j === line.length - 1,
        },
      });
    });
  });
  s.addText(runs, Object.assign({ x, y, w, h, align, fontFace: XB, fontSize: size }, TXT));
}

/** small bold letter-spaced caps label */
function label(s, x, y, w, text, color = INK_SOFT, align = 'left') {
  s.addText(text, Object.assign({ x, y, w, h: 0.286, align, fontFace: MD, fontSize: 11, bold: true, color, charSpacing: 3 }, TXT));
}

/** justified / left grey paragraph copy */
function body(s, x, y, w, h, text, opt = {}) {
  s.addText(text, Object.assign({
    x, y, w, h, valign: 'top', isTextBox: true, wrap: true,
    fontFace: BODY, fontSize: opt.size || 12, color: opt.color || GREY,
    align: opt.align || 'left', lineSpacingMultiple: 1.5, italic: opt.italic || false,
  }));
}

const url = (s, x, y, color) => s.addText('www.frenzy.com', Object.assign(
  { x, y, w: 2.044, h: 0.286, fontFace: UI, fontSize: 11, color, charSpacing: 3 }, TXT));

/** little rounded "tab" accent, 0.63 x 0.2 in */
const tab = (s, x, y, color) => s.addShape('roundRect', {
  x, y, w: 0.629, h: 0.203, rectRadius: 0.1, fill: { color }, line: { type: 'none' },
});

/** thin rounded rule, 2.37 x 0.05 in */
const rule = (s, x, y, color) => s.addShape('roundRect', {
  x, y, w: 2.37, h: 0.05, rectRadius: 0.025, fill: { color }, line: { type: 'none' },
});

const SHADOW = { type: 'outer', angle: 135, blur: 4, offset: 2, color: '000000', opacity: 0.3 };

/** numbered disc: filled circle with a white keyline and a centred numeral */
function disc(s, x, y, d, fill, text, opt = {}) {
  s.addShape('ellipse', {
    x, y, w: d, h: d, fill: { color: fill },
    line: opt.ring === false ? { type: 'none' } : { color: WHITE, width: 3 },
    shadow: SHADOW,
  });
  if (text) {
    s.addText(text, Object.assign({
      x, y: y + d / 2 - 0.21, w: d, h: 0.42, align: 'center',
      fontFace: BODY, fontSize: opt.size || 18, bold: true, color: opt.color || WHITE,
    }, TXT));
  }
}

/** map pin: rounded caption plate, numbered disc, drop line and end dot */
function marker(s, x, y, num, discFill, plateFill, textColor) {
  s.addShape('roundRect', {
    x, y, w: 1.144, h: 0.361, rectRadius: 0.18,
    fill: { color: plateFill }, line: { color: WHITE, width: 1 },
  });
  s.addShape('line', { x: x + 0.057, y: y + 0.408, w: 0, h: 0.312, line: { color: WHITE, width: 1 } });
  s.addShape('ellipse', { x: x + 0.02, y: y + 0.708, w: 0.082, h: 0.082, fill: { color: plateFill }, line: { color: WHITE, width: 1 } });
  s.addShape('ellipse', {
    x: x - 0.18, y: y - 0.066, w: 0.474, h: 0.474,
    fill: { color: discFill }, line: { color: WHITE, width: 1.5 }, shadow: SHADOW,
  });
  s.addText(num, Object.assign({ x: x - 0.18, y: y - 0.03, w: 0.474, h: 0.4, align: 'center', fontFace: UI, fontSize: 20, bold: true, color: WHITE, charSpacing: 3 }, TXT));
  s.addText('Market', { x: x + 0.321, y: y - 0.04, w: 0.797, h: 0.364, valign: 'top', fontFace: BODY, fontSize: 12, color: textColor, lineSpacingMultiple: 1.5 });
}

/* -------------------------------------------------------------- food icons */
function icon(s, kind, x, y, w, h, color) {
  const f = { fill: { color }, line: { type: 'none' } };
  if (kind === 'burger') {
    s.addShape('roundRect', Object.assign({ x, y, w, h: h * 0.3, rectRadius: h * 0.15 }, f));
    s.addShape('rect', Object.assign({ x, y: y + h * 0.38, w, h: h * 0.14 }, f));
    s.addShape('wave', Object.assign({ x, y: y + h * 0.55, w, h: h * 0.18 }, f));
    s.addShape('roundRect', Object.assign({ x, y: y + h * 0.78, w, h: h * 0.22, rectRadius: h * 0.09 }, f));
  } else if (kind === 'bottle') {
    s.addShape('roundRect', Object.assign({ x, y: y + h * 0.18, w: w * 0.42, h: h * 0.82, rectRadius: w * 0.1 }, f));
    s.addShape('rect', Object.assign({ x: x + w * 0.12, y, w: w * 0.16, h: h * 0.24 }, f));
    s.addShape('ellipse', { x: x + w * 0.52, y: y + h * 0.4, w: w * 0.48, h: h * 0.48, fill: { type: 'none' }, line: { color, width: 2 } });
  } else if (kind === 'cone') {
    s.addShape('ellipse', Object.assign({ x: x + w * 0.1, y, w: w * 0.8, h: h * 0.5 }, f));
    s.addShape('triangle', Object.assign({ x: x + w * 0.18, y: y + h * 0.42, w: w * 0.64, h: h * 0.58, rotate: 180 }, f));
  } else {
    cutlery(s, x, y, Math.min(w, h), color);
  }
}

/** ASCII silhouette painted as merged rectangles (one per horizontal run) */
function mapArt(s, rows, x, y, w, h, stops) {
  const cw = w / rows[0].length;
  const ch = h / rows.length;
  rows.forEach((row, r) => {
    const color = gradAt(stops, (r + 0.5) / rows.length);
    let c = 0;
    while (c < row.length) {
      if (row[c] === '#') {
        let e = c;
        while (e < row.length && row[e] === '#') e++;
        s.addShape('rect', {
          x: x + c * cw, y: y + r * ch, w: (e - c) * cw + 0.006, h: ch + 0.006,
          fill: { color }, line: { type: 'none' },
        });
        c = e;
      } else c++;
    }
  });
}

const WORLD_MAP = [
  '.........................###############......................#.................................',
  '...................#...#####.###########......#................####.............................',
  '.................####.#####....########................#.....###########..#.....................',
  '.......#####.##########.####...########.........###.......##########################.##.........',
  '.....####################.###..#####...........########################################.........',
  '...###########.#########.###...###....##.....##########################################.........',
  '..###########...................#............###.######################################.........',
  '..##.....###...............#..............#..###.###############################....##..........',
  '.........###...............##............###..############################......................',
  '.........####..............##............################################.......................',
  '.........#####..##########.##.............###############################.......................',
  '.........##################................##############################.......................',
  '........################.................###########...##.################....###...............',
  '........###############..................###...#####....##.######################...#...........',
  '........##############...................###......#......#.####################.##..#...........',
  '........#############.....................#####........#.######################..#.##...........',
  '........############.....................######..............###################..#.............',
  '.........########.#.....................#######...............###################...............',
  '.........#####....##...................########..............###################................',
  '...........###.........................########.............##..#################...............',
  '...........###..#..#..................#################.######...######.######..................',
  '...........######.....................#######################.....####...####...................',
  '.............#####....................##################.###......###.....####...#..............',
  '................##....................####################.........##.....####...##.............',
  '.................#..#..................#####################.......##.....#.#.....#.............',
  '...................######...............###################................#......#.............',
  '...................#########.............##..##############...............##...##...............',
  '...................#########..................############................##..###...............',
  '..................###########.................###########..................##.####...#..........',
  '..................##############...............#########....................#..#.#...####.......',
  '..................###############..............#########.....................##........###......',
  '...................##############...............########................................#.#.....',
  '...................#############................########.............................#..#.......',
  '....................############...............#########..#.......................####..#.......',
  '.....................###########...............#########.##......................########.......',
  '......................#########.................#######..##.....................##########......',
  '......................#########.................#######..##...................############......',
  '......................#######...................######........................#############.....',
  '......................#######...................######........................############......',
  '......................######.....................####.........................############......',
  '......................######.....................##...........................##....#####.......',
  '......................#####..........................................................###.......#',
  '......................####....................................................................#.',
  '......................###............................................................#......##..',
  '......................###..................................................................#....',
  '.......................##.......................................................................',
  '.......................##.......................................................................',
  '........................##......................................................................',
];

const US_MAP = [
  '........##..............................................................',
  '......########..........................................................',
  '......#############.....................................................',
  '.....######################.........................................##..',
  '.....##################################.............................###.',
  '.....#######################################........................###.',
  '....########################################.......................#####',
  '....#######################################..###.##..............######.',
  '...#################################################..........########..',
  '...#########################################..................#######...',
  '...#########################################.................########...',
  '...#########################################...............##########...',
  '..#######################################.##...............##########...',
  '..#########################################...............#########.....',
  '..#################...........##################..####..###########.....',
  '..##################...........###################################......',
  '...##############.##............##################################......',
  '...##############.##...........##################################.......',
  '...################...........###################################.......',
  '...############################################################.#.......',
  '...#############################################################........',
  '....############################################################........',
  '....############################################################........',
  '.....########################################################...........',
  '......########################################################..........',
  '.......###################################################..##..........',
  '....######.........##########################################...........',
  '....######...........#######################################............',
  '....##.###...........#######################################............',
  '....###.##...........######################################.............',
  '....#####............######################################.............',
  '.....###...............####################################.............',
  '...########.............###################################.............',
  '...########..............#########################..#######.............',
  '..#.########.............######################.........####............',
  '..##########..............#...#########......#...........###............',
  '....########...................#######...................####...........',
  '..##########...................#####......................####..........',
  '..##########....................###........................###..........',
  '..##########....................###........................###..........',
  '...######....###..................##....................................',
  '.....##.......###...........##..........................................',
  '....#.#.........##..........##..........................................',
  '..##........................#...........................................',
];

/* --------------------------------------------------------------- copy text */
const LOREM = 'PLACEHOLDER';
const L_SHORT = 'PLACEHOLDER';
const L_TINY = 'Duis aute irure dolor in';
const L_MINI = 'Duis aute irure dolor in reprehenderit in';
const L_LINE = 'PLACEHOLDER';

/* =========================================================== slide builders */

// 1 - cover
function s01(s) {
  page(s);
  redPanel(s, 6.667, 0, 6.667, 7.5);
  pic(s, 4.361, 1.458, 4.611, 4.584);
  title(s, 7.675, 2.231, 4.788, 1.717, [[['Fre', WHITE], ['nzy', AMBER]]], 96, 'right');
  s.addText('Presentation', Object.assign({ x: 9.564, y: 4.016, w: 2.663, h: 0.37, align: 'right', fontFace: LT, fontSize: 16, color: WHITE, charSpacing: 6 }, TXT));
  tab(s, 11.678, 6.238, WHITE);
  disc(s, 4.479, 1.17, 0.549, WHITE, null, { ring: false });
  logo(s, 0.636, 0.995, 0.71, RED, 16);
  s.addText('Owned By', Object.assign({ x: 0.644, y: 5.933, w: 1.117, h: 0.337, fontFace: BODY, fontSize: 14, bold: true, color: INK_SOFT }, TXT));
  s.addText('Lanterby_Studio', Object.assign({ x: 0.636, y: 6.229, w: 1.585, h: 0.303, fontFace: UI, fontSize: 12, color: GREY }, TXT));
  rule(s, 0.691, 3.658, WHITE);
  url(s, 10.386, 1.172, WHITE);
}

// 2 - what is fast food
function s02(s) {
  page(s);
  redPanel(s, 1.895, 0.9, 11.439, 2.949, 'h');
  pic(s, 0.561, 0.9, 2.967, 2.949);
  pic(s, 3.791, 0.9, 2.967, 2.949);
  logo(s, 10.562, 2.048, 0.71, WHITE, 16);
  rule(s, 6.496, 2.395, WHITE);
  title(s, 1.525, 4.77, 3.631, 1.582, [[['What Is ', INK]], [['Fast Food ?', RED]]], 44);
  label(s, 6.496, 4.77, 2.942, 'INDUSTRY SIZE & SCALE');
  body(s, 6.496, 5.024, 5.544, 1.576, LOREM + ' conseuate', { align: 'justify' });
  url(s, 10.534, 1.301, WHITE);
  disc(s, 0.561, 0.48, 0.549, WHITE, null, { ring: false });
}

// 3 - market trends
function s03(s) {
  page(s);
  pic(s, 0, 0, 4.444, 7.5);
  redPanel(s, 8.685, 0, 4.648, 7.5);
  pic(s, 7.199, 2.273, 2.972, 2.954);
  title(s, 9.767, 0.586, 2.614, 1.582, [[['Market ', WHITE]], [['Trends', AMBER]]], 44, 'right');
  tab(s, 11.148, 3.61, WHITE);
  logo(s, 10.534, 5.984, 0.71, WHITE, 16);
  [[0.997, 0.743, 'MARKET 01'], [5.793, 5.539, 'MARKET 02']].forEach(([by, ly, cap]) => {
    body(s, 5.281, by, 2.772, 0.97, L_SHORT);
    label(s, 5.281, ly, 1.46, cap);
    disc(s, 4.247, ly, 0.394, RED, null);
  });
  s.addText('Owned By', Object.assign({ x: 5.289, y: 3.285, w: 1.117, h: 0.337, fontFace: BODY, fontSize: 14, bold: true, color: INK_SOFT }, TXT));
  s.addText('Lanterby_Studio', Object.assign({ x: 5.281, y: 3.581, w: 1.585, h: 0.303, fontFace: UI, fontSize: 12, color: GREY }, TXT));
  url(s, 1.067, 6.289, WHITE);
}

// 4 - industry snapshot (bar chart)
function s04(s) {
  page(s);
  redPanel(s, 8.685, 0, 4.648, 7.5);
  pic(s, 9.523, 2.273, 2.972, 2.954);
  pic(s, 5.789, 0, 2.896, 5.227);
  const bars = [
    { x: 1.229, pct: '50%', year: '2021', fillTop: 3.922, color: AMBER, yearColor: AMBER },
    { x: 2.030, pct: '80%', year: '2022', fillTop: 3.531, color: RED, yearColor: RED },
    { x: 2.801, pct: '60%', year: '2023', fillTop: 3.755, color: AMBER, yearColor: AMBER },
    { x: 3.602, pct: '26%', year: '2024', fillTop: 4.503, color: RED, yearColor: RED },
  ];
  bars.forEach(b => {
    s.addShape('rect', { x: b.x, y: 3.096, w: 0.188, h: 1.914, fill: { color: INK }, line: { type: 'none' } });
    s.addShape('rect', { x: b.x, y: b.fillTop, w: 0.188, h: 5.01 - b.fillTop, fill: { color: b.color }, line: { type: 'none' } });
    disc(s, b.x - 0.18, 2.556, 0.549, WHITE, null, { ring: false });
    s.addText(b.pct, Object.assign({ x: b.x - 0.143, y: 2.696, w: 0.475, h: 0.269, align: 'center', fontFace: BODY, fontSize: 10, bold: true, color: CHARCOAL }, TXT));
    s.addText(b.year, Object.assign({ x: b.x - 0.2, y: 5.056, w: 0.59, h: 0.286, align: 'center', fontFace: MD, fontSize: 11, color: b.yearColor, charSpacing: 3 }, TXT));
  });
  title(s, 1.031, 0.537, 3.166, 1.582, [[['Industry ', INK]], [['Snapshot', AMBER]]], 44);
  tab(s, 10.695, 6.336, WHITE);
  url(s, 9.987, 1.02, WHITE);
  label(s, 1.019, 5.802, 2.134, 'GLOBAL REVENUE');
  logo(s, 6.386, 0.761, 0.71, WHITE, 16);
  body(s, 1.031, 6.083, 6.571, 0.667, 'PLACEHOLDER', { align: 'justify' });
  disc(s, 8.411, 5.01, 0.549, RED, null, { ring: false });
}

// 5 - target market (world map)
function s05(s) {
  page(s);
  pic(s, 0, 0, 2.896, 7.5);
  mapArt(s, WORLD_MAP, 3.689, 2.481, 8.466, 4.179, [[0, RED], [1, RED_DARK]]);
  title(s, 3.822, 0.772, 4.718, 0.841, [[['Target ', INK], ['Market', RED]]], 44);
  [[4.938, 3.015, '1'], [5.645, 4.272, '2'], [7.979, 3.835, '3'], [10.311, 3.203, '4']]
    .forEach(([x, y, n]) => marker(s, x, y, n, AMBER, AMBER, WHITE));
  tab(s, 7.807, 6.77, WHITE);
  logo(s, 0.451, 0.77, 0.71, WHITE, 16);
  ['01', '02', '03', '04'].forEach((n, i) => {
    s.addText([
      { text: n + '  ', options: { fontFace: MD, bold: true, color: INK_SOFT } },
      { text: L_TINY, options: { fontFace: BODY, color: GREY } },
    ], { x: 10.21, y: 0.462 + i * 0.371, w: 2.772, h: 0.348, valign: 'top', fontSize: 11, lineSpacingMultiple: 1.5, charSpacing: 0 });
  });
  url(s, 10.205, 6.759, RED);
}

// 6 - target market (US map)
function s06(s) {
  page(s);
  redPanel(s, 6.537, 0, 6.796, 7.5);
  pic(s, 0.671, 0.931, 3.119, 3.1);
  pic(s, 2.879, 3.75, 3.119, 3.1);
  mapArt(s, US_MAP, 7.463, 3.105, 5.319, 3.478, [[0, AMBER], [1, AMBER]]);
  title(s, 3.231, 0.646, 2.504, 1.582, [[['Target ', INK]], [['Market', RED]]], 44);
  [[8.173, 2.654, '1'], [8.653, 4.228, '2'], [10.412, 3.839, '3'], [11.736, 4.902, '4'], [7.903, 5.185, '5']]
    .forEach(([x, y, n]) => marker(s, x, y, n, AMBER, WHITE, RED));
  tab(s, 6.205, 0.931, WHITE);
  logo(s, 0.906, 5.553, 0.71, RED, 16, [0.816, 6.448]);
  url(s, 9.259, 6.818, WHITE);
  [['01', 7.892, 0.769], ['03', 7.892, 1.096], ['02', 10.281, 0.769], ['04', 10.281, 1.096], ['05', 10.301, 1.415]]
    .forEach(([n, x, y]) => s.addText([
      { text: n + '    ', options: { fontFace: MD, bold: true, color: WHITE } },
      { text: L_TINY, options: { fontFace: BODY, color: PINK } },
    ], { x, y, w: 2.457, h: 0.348, valign: 'top', fontSize: 11, lineSpacingMultiple: 1.5 }));
}

// 7 - consumer behavior
function s07(s) {
  page(s);
  redPanel(s, 0, 0, 13.333, 2.949, 'h');
  [4.15, 7.115, 10.08].forEach(x => pic(s, x, 1.475, 2.801, 2.785));
  title(s, 0.664, 0.545, 3.485, 1.582, [[['Consumer ', WHITE]], [['Behavior', AMBER]]], 44);
  tab(s, 0.837, 6.394, WHITE);
  [[4.446, 5.12, AMBER, '01', WHITE], [7.395, 8.069, RED, '02', WHITE], [10.436, 11.11, WHITE, '03', CHARCOAL]]
    .forEach(([tx, cx, fill, num, numColor]) => {
      body(s, tx, 5.475, 1.979, 1.273, 'PLACEHOLDER', { align: 'center' });
      label(s, tx + 0.247, 5.189, 1.483, 'CONSUMER ', INK_SOFT, 'center');
      disc(s, cx, 3.937, 0.788, fill, num, { ring: false, color: numColor });
    });
  url(s, 7.395, 0.795, WHITE);
  logo(s, 10.684, 0.603, 0.526, WHITE, 14);
  body(s, 0.746, 3.9, 2.266, 1.576, '\u201c Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore europes vui officia deserut mollit anim id est \u201d"', { italic: true });
}

// 8 - business innovation
function s08(s) {
  page(s);
  redPanel(s, 0, 0, 4.547, 3.863, 'h');
  pic(s, 0, 3.863, 4.547, 3.637);
  pic(s, 4.547, 0, 4.547, 3.863);
  disc(s, 10.0, 0.943, 0.788, RED, null, { ring: false });
  disc(s, 9.987, 4.473, 0.788, AMBER, null, { ring: false });
  disc(s, 6.149, 4.474, 0.788, WHITE, null, { ring: false });
  icon(s, 'burger', 10.185, 1.141, 0.418, 0.376, WHITE);
  icon(s, 'bottle', 6.366, 4.671, 0.393, 0.393, RED);
  icon(s, 'cone', 10.185, 4.628, 0.392, 0.479, WHITE);
  title(s, 0.633, 1.141, 3.645, 1.582, [[['Business', WHITE]], [['Innovation', AMBER]]], 44);
  tab(s, 0.773, 6.107, WHITE);
  [[6.153, 5.514, 'INNOVATION 02'], [9.996, 2.019, 'INNOVATION 01'], [9.996, 5.514, 'INNOVATION 03']]
    .forEach(([x, y, cap]) => {
      label(s, x, y, 1.98, cap);
      body(s, x, y + 0.255, 2.504, 0.97, L_SHORT);
    });
  url(s, 0.654, 0.775, WHITE);
  logo(s, 6.127, 1.268, 0.526, WHITE, 14);
  pic(s, 3.246, 2.57, 2.602, 2.587);
}

// 9 - challenges & opportunities
function s09(s) {
  page(s);
  redPanel(s, 8.347, 0, 4.986, 7.5);
  pic(s, 9.44, 2.358, 2.801, 2.785);
  pic(s, 5.452, 0, 2.896, 7.5);
  title(s, 6.311, 0.632, 4.671, 1.582, [[['Challenges ', WHITE]], [[' opportunities', WHITE]]], 44);
  tab(s, 9.32, 6.319, WHITE);
  url(s, 6.521, 6.279, WHITE);
  logo(s, 10.981, 6.128, 0.526, WHITE, 14);
  [[1.558, 'INFLATION', 1.403, RED], [4.415, 'COST PRESSURES', 2.085, AMBER]].forEach(([y, cap, w, dotColor]) => {
    label(s, 0.927, y, w, cap);
    body(s, 0.927, y + 0.254, 3.273, 1.273, 'PLACEHOLDER', { align: 'justify' });
    disc(s, 5.254, y, 0.394, dotColor, null);
  });
}

// 10 - break slide
function s10(s) {
  page(s);
  redPanel(s, 0, 0, 6.667, 7.5);
  pic(s, 4.361, 1.458, 4.611, 4.584);
  title(s, 8.207, 2.231, 4.257, 1.717, [[['Break', RED]]], 96, 'right');
  title(s, 8.186, 3.466, 4.278, 1.717, [[['Slides', GOLD]]], 96, 'right');
  tab(s, 11.678, 6.315, WHITE);
  disc(s, 8.423, 6.143, 0.549, RED, null, { ring: false });
  logo(s, 10.549, 0.995, 0.71, RED, 16);
  s.addText('Owned By', Object.assign({ x: 0.644, y: 5.933, w: 1.117, h: 0.337, fontFace: BODY, fontSize: 14, bold: true, color: WHITE }, TXT));
  s.addText('Lanterby_Studio', Object.assign({ x: 0.636, y: 6.229, w: 1.585, h: 0.303, fontFace: UI, fontSize: 12, color: WHITE }, TXT));
  rule(s, 0.691, 3.658, WHITE);
  url(s, 0.546, 1.172, WHITE);
}

// 11 - value menu strategy
function s11(s) {
  page(s);
  redPanel(s, 8.862, 0, 4.471, 7.5);
  [0.278, 2.626, 4.974].forEach(y => pic(s, 7.732, y, 2.261, 2.247));
  title(s, 1.016, 0.723, 3.975, 1.582, [[['Value Menu ', INK]], [['Strategy', RED]]], 44);
  [['$15', 'MENU 01', 0.548, 0.526, AMBER], ['$17', 'MENU 02', 2.895, 2.895, RED], ['$18', 'MENU 03', 5.243, 5.243, AMBER]]
    .forEach(([price, cap, y, dy, dotColor]) => {
      s.addText(price, Object.assign({ x: 10.555, y, w: 0.886, h: 0.572, fontFace: BODY, fontSize: 28, bold: true, color: WHITE }, TXT));
      label(s, 10.615, y + 0.617, 1.19, cap, WHITE);
      body(s, 10.615, y + 0.854, 1.982, 0.667, L_MINI, { color: PINK });
      disc(s, 7.732, dy, 0.394, dotColor, null);
    });
  tab(s, 5.806, 1.064, WHITE);
  label(s, 1.016, 3.306, 1.54, 'LUXE BOXES');
  body(s, 1.016, 3.561, 4.669, 1.879, LOREM, { align: 'justify' });
  url(s, 1.016, 6.622, GREY_LT);
  logo(s, 4.137, 6.429, 0.526, RED, 14);
}

// 12 - behavioral pricing
function s12(s) {
  page(s);
  redPanel(s, 0, 0, 4.471, 3.874);
  pic(s, 4.471, 0, 3.718, 7.5);
  pic(s, 1.038, 2.655, 2.451, 2.437);
  logo(s, 1.546, 1.128, 0.526, WHITE, 14);
  tab(s, 1.93, 6.415, RED);
  title(s, 8.824, 0.863, 3.719, 1.582, [[['Behavioral ', INK]], [['Pricing', RED]]], 44);
  [[8.824, 3.62, 3.566, 'KEEP BELOW $10'], [8.876, 5.716, 5.716, 'KEEP BELOW $12']].forEach(([x, y, dy, cap]) => {
    label(s, x, y, 2.053, cap);
    body(s, x, y + 0.254, 3.529, 0.667, L_SHORT + ' ');
    disc(s, 7.992, dy, 0.394, RED, null);
  });
}

// 13 - bundles & upsell
function s13(s) {
  page(s);
  redPanel(s, 0, 0, 2.284, 5.444);
  // red pill band behind the cards: rounded right cap, then the straight body
  s.addShape('ellipse', { x: 8.42, y: 2.057, w: 3.387, h: 3.387, fill: { color: RED_MID }, line: { type: 'none' } });
  gradRect(s, 2.284, 2.057, 7.829, 3.387, RED_GRAD);
  const cards = [
    { x: 1.142, price: '$15', pc: RED, dot: RED, n: '01' },
    { x: 3.740, price: '$9', pc: AMBER, dot: AMBER, n: '02' },
    { x: 6.339, price: '$11', pc: RED, dot: RED, n: '03' },
    { x: 8.937, price: '$10', pc: AMBER, dot: AMBER, n: '04' },
  ];
  cards.forEach(c => pic(s, c.x, 2.532, 2.451, 2.437));
  title(s, 2.845, 0.676, 8.441, 0.841, [[['Bundles & ', INK], ['Upsell', RED], [' ', INK], ['Strategy', RED]]], 44);
  logo(s, 0.879, 0.67, 0.526, WHITE, 14, [0.688, 1.328]);
  tab(s, 12.08, 0.993, RED);
  cards.forEach(c => {
    s.addText(c.price, Object.assign({ x: c.x + 0.556, y: 5.787, w: 0.886, h: 0.572, fontFace: BODY, fontSize: 28, bold: true, color: c.pc }, TXT));
    label(s, c.x + 0.616, 6.405, 1.166, 'MENU 01', INK);
    body(s, c.x + 0.616, 6.642, 1.982, 0.364, L_TINY);
    disc(s, c.x - 0.135, 2.532, 0.682, c.dot, c.n);
  });
  s.addText('www.frenzy.com', Object.assign({ x: -0.474, y: 3.803, w: 2.044, h: 0.286, rotate: 270, fontFace: UI, fontSize: 11, color: WHITE, charSpacing: 3 }, TXT));
}

// 14 - promotional campaigns
function s14(s) {
  page(s);
  redPanel(s, 0, 0, 7.899, 7.5);
  pic(s, 7.018, 1.185, 2.483, 4.991);
  pic(s, 9.713, 1.185, 2.483, 4.991);
  title(s, 1.352, 1.589, 4.253, 1.582, [[['Promotional ', WHITE]], [['Campaigns ', WHITE]]], 44);
  disc(s, 8.886, 1.253, 0.682, AMBER, '01');
  disc(s, 11.62, 1.253, 0.682, RED, '02');
  tab(s, 10.64, 6.817, RED);
  tab(s, 10.64, 0.48, AMBER);
  [[2.096, 3.852, 1.432, 3.775], [2.077, 5.255, 1.413, 5.178]].forEach(([tx, ty, cx, cy]) => {
    label(s, tx, ty, 2.044, 'ADVERTISEMENT', WHITE);
    body(s, tx, ty + 0.254, 3.529, 0.667, L_SHORT + ' ', { color: PINK_LT });
    s.addShape('ellipse', { x: cx, y: cy, w: 0.394, h: 0.394, fill: { color: RED }, line: { color: WHITE, width: 3 } });
  });
  logo(s, 8.388, 6.618, 0.526, RED, 14);
  url(s, 1.352, 0.965, WHITE);
}

// 15 - location & concept
function s15(s) {
  page(s);
  redPanel(s, 9.746, 0, 3.547, 7.5);
  pic(s, 10.398, 1.509, 2.229, 4.481);
  pic(s, 0, 0, 5.0, 3.75);
  pic(s, 0, 3.75, 5.0, 3.75);
  title(s, 5.639, 0.595, 3.345, 1.582, [[['Location', INK]], [['& ', INK], ['Concept', RED]]], 44);
  logo(s, 10.739, 3.558, 0.526, WHITE, 14);
  tab(s, 11.216, 6.639, WHITE);
  [[3.175, 'LOCATION', 1.326], [4.71, 'CONCEPT', 1.213]].forEach(([y, cap, w]) => {
    label(s, 5.673, y, w, cap);
    body(s, 5.673, y + 0.254, 3.529, 0.667, L_SHORT + ' ');
    disc(s, 4.807, y - 0.054, 0.394, RED, null);
  });
  url(s, 5.673, 6.455, RED);
}

// 16 - supply chain & inventory
function s16(s) {
  page(s);
  redPanel(s, 0, 0, 5.074, 7.5);
  pic(s, 1.105, 2.326, 2.865, 2.848);
  // four ringed icon nodes on a shared axis; arcs alternate above / below
  const nodes = [
    { x: 5.588, color: RED, icon: 'burger', num: '01', numX: 5.908, top: true },
    { x: 7.359, color: AMBER, icon: 'bottle', num: '02', numX: 7.704, top: false },
    { x: 9.131, color: RED, icon: 'cone', num: '03', numX: 9.475, top: true },
    { x: 10.902, color: AMBER, icon: 'cutlery', num: '04', numX: 11.256, top: false },
  ];
  nodes.forEach(n => {
    s.addShape('arc', {
      x: n.x, y: 3.519, w: 1.771, h: 1.771, angleRange: [180, 0], rotate: n.top ? 0 : 180,
      line: { color: n.color, width: 6 },
    });
    s.addShape('ellipse', {
      x: n.x + 0.213, y: 3.682, w: 1.346, h: 1.346,
      fill: { color: n.color }, line: { color: WHITE, width: 3 }, shadow: SHADOW,
    });
    icon(s, n.icon, n.x + 0.676, 4.158, 0.42, 0.42, WHITE);
    s.addText(n.num, Object.assign({ x: n.numX, y: n.top ? 1.922 : 5.855, w: 1.082, h: 1.01, align: 'center', fontFace: BODY, fontSize: 54, bold: true, color: n.color }, TXT));
  });
  // captions sit beside the numeral: right of it on the top row, left below
  [[7.359, 1.841], [10.902, 1.841], [5.588, 5.677], [9.131, 5.677]].forEach(([x, y]) => {
    label(s, x + 0.32, y, 1.028, 'SUPPLY');
    body(s, x, y + 0.231, 1.771, 0.97, 'Duis aute irure dolor in reprehenderit in voluptate velit esse', { align: 'center' });
  });
  logo(s, 1.646, 0.719, 0.526, WHITE, 14);
  tab(s, 2.286, 5.982, AMBER);
  title(s, 5.653, 0.673, 5.184, 0.572, [[['Supply Chain & Inventory', INK]]], 28);
  url(s, 1.636, 6.575, WHITE);
}

// 17 - enhancing productivity
function s17(s) {
  page(s);
  redPanel(s, 6.667, 0, 3.547, 7.5);
  pic(s, 5.157, 0.958, 3.019, 5.435);
  pic(s, 8.704, 0.958, 3.019, 5.435);
  title(s, 0.704, 1.356, 3.759, 1.447, [[['Enhancing ', INK]], [['Productivity', RED]]], 40);
  logo(s, 11.92, 5.803, 0.71, RED, 16, [11.831, 6.698]);
  tab(s, 12.001, 0.784, AMBER);
  url(s, 7.413, 6.74, WHITE);
  label(s, 0.704, 3.867, 1.54, 'LUXE BOXES');
  body(s, 0.704, 4.121, 3.925, 2.182, LOREM, { align: 'justify' });
  disc(s, 9.872, 0.601, 0.682, RED, null);
  disc(s, 6.326, 0.601, 0.682, AMBER, null);
}

// 18 - building a strong brand
function s18(s) {
  page(s);
  redPanel(s, 0, 0, 3.547, 3.463);
  pic(s, 0, 3.463, 3.547, 4.037);
  pic(s, 3.547, 0, 3.547, 3.463);
  logo(s, 0.999, 1.469, 0.526, WHITE, 14);
  title(s, 7.985, 0.548, 4.443, 1.447, [[['Building ', INK]], [['a Strong ', INK], ['Brand', RED]]], 40);
  [[4.553, RED, '01'], [7.595, AMBER, '02'], [10.637, RED, '03']].forEach(([x, color, n]) => {
    body(s, x, 5.688, 2.28, 0.97, L_LINE);
    label(s, x, 5.433, 0.958, 'BRAND');
    disc(s, x + 0.013, 4.428, 0.682, color, n);
  });
  tab(s, 11.595, 0.775, RED);
  url(s, 0.751, 6.172, WHITE);
  body(s, 8.024, 2.493, 4.364, 0.97, 'PLACEHOLDER', { align: 'justify' });
}

// 19 - promotional events
function s19(s) {
  page(s);
  redPanel(s, 7.565, 0, 3.547, 7.5);
  pic(s, 11.112, 0, 2.275, 7.5);
  [0.721, 2.786, 4.852].forEach((y, i) => pic(s, [8.313, 8.315, 8.369][i], y, 1.939, 1.927));
  title(s, 1.511, 1.178, 3.882, 1.447, [[['Promotional ', INK]], [['Events', RED]]], 40);
  tab(s, 1.584, 3.155, RED);
  label(s, 1.511, 4.041, 1.701, 'MAIN EVENTS');
  body(s, 1.511, 4.295, 4.499, 2.182, LOREM + ' aute irure dolor in reprehenderit ', { align: 'justify' });
  logo(s, 8.564, 6.548, 0.526, WHITE, 14);
  url(s, 8.316, 0.31, WHITE);
}

// 20 - food challenges
function s20(s) {
  page(s);
  redPanel(s, 0, 0, 5.971, 7.512);
  pic(s, 0, 0, 2.986, 3.762);
  pic(s, 2.986, 3.75, 2.986, 3.762);
  pic(s, 2.016, 2.938, 1.939, 1.927);
  title(s, 7.353, 2.304, 3.328, 1.447, [[['Food', INK]], [['Challenges', RED]]], 40);
  label(s, 7.353, 4.332, 3.044, 'INTERACTIVE CAMPAIGNS');
  body(s, 7.353, 4.618, 4.675, 1.879, LOREM, { align: 'justify' });
  [[3.34, 1.779, 3.928, 1.525, 4.034, 0.589, AMBER, '02'], [0.349, 5.73, 0.938, 5.476, 1.051, 4.54, RED, '01']]
    .forEach(([bx, by, lx, ly, cx, cy, color, n]) => {
      body(s, bx, by, 2.28, 0.97, L_LINE, { color: PINK_LT, align: 'center' });
      label(s, lx, ly, 0.958, 'BRAND', WHITE, 'center');
      disc(s, cx, cy, 0.682, color, n);
    });
  tab(s, 11.399, 1.403, RED);
  logo(s, 7.467, 1.25, 0.526, RED, 14);
}

// 21 - contact us
function s21(s) {
  page(s);
  redPanel(s, 6.667, 0, 6.667, 7.5);
  title(s, 1.316, 0.627, 3.177, 0.774, [[['Contact ', INK], ['us', RED]]], 40);
  logo(s, 11.24, 6.42, 0.526, WHITE, 14);
  url(s, 7.131, 6.59, WHITE);
  tab(s, 12.02, 0.707, WHITE);
  [['WEBSITE', 1.163, 2.424, RED, '01'], ['EMAIL', 0.889, 4.041, AMBER, '02'], ['SOCIAL MEDIA', 1.785, 5.659, RED, '03']]
    .forEach(([cap, w, y, color, n]) => {
      label(s, 2.16, y, w, cap);
      body(s, 2.16, y + 0.286, 3.177, 0.667, 'PLACEHOLDER');
      disc(s, 1.382, y - 0.092, 0.682, color, n);
    });
  rule(s, 7.211, 0.829, WHITE);
  pic(s, 7.738, 1.508, 4.511, 4.484);
}

// 22 - thank you
function s22(s) {
  page(s);
  redPanel(s, 6.667, 0, 6.667, 7.5);
  pic(s, 4.361, 1.458, 4.611, 4.584);
  title(s, 8.001, 2.351, 4.462, 1.717, [[['Thank', WHITE]]], 96, 'right');
  title(s, 9.753, 3.497, 2.711, 1.717, [[['you', AMBER]]], 96, 'right');
  tab(s, 11.678, 6.238, WHITE);
  disc(s, 4.479, 1.17, 0.549, RED, null, { ring: false });
  logo(s, 0.636, 0.995, 0.71, RED, 16);
  s.addText('Owned By', Object.assign({ x: 0.644, y: 5.933, w: 1.117, h: 0.337, fontFace: BODY, fontSize: 14, bold: true, color: INK_SOFT }, TXT));
  s.addText('Lanterby_Studio', Object.assign({ x: 0.636, y: 6.229, w: 1.585, h: 0.303, fontFace: UI, fontSize: 12, color: GREY }, TXT));
  rule(s, 0.691, 3.658, WHITE);
  url(s, 10.386, 1.172, WHITE);
}

/* ==================================================================== build */
const SLIDES = [s01, s02, s03, s04, s05, s06, s07, s08, s09, s10, s11,
  s12, s13, s14, s15, s16, s17, s18, s19, s20, s21, s22];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'FRENZY', width: 13.333, height: 7.5 });
pptx.layout = 'FRENZY';
pptx.author = 'Lanterby_Studio';
pptx.title = 'Frenzy Presentation';

SLIDES.forEach(build => build(pptx.addSlide()));

pptx.writeFile({ fileName: path.join(__dirname, '00b813de-d42f-4a38-bd4e-f96c70a5ec45_grok_final.pptx') })
  .then(f => console.log('wrote', f));
