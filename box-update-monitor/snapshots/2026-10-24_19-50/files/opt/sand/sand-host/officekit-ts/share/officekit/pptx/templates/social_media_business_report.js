/**
 * Hakira "Social Media Business" deck - rebuilt with pptxgenjs.
 * 24 slides, 13.333 x 7.5 in, dark background, Bricolage Grotesque / Syne / Playfair.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const C = {
  bg:       '0E0606',
  white:    'FFFFFF',
  black:    '000000',
  red:      'E83202',
  redSoft:  'FE4F31',
  redCard:  'DA3930',
  blue:     '4937FB',
  blueBtn:  '4838FA',
  star:     'EAFD1F',
  panel:    'F5F6FA',
  gray:     'D8D8D8',
  grayMid:  '7F7F7F',
  grayLt:   'BFBFBF',
  grayDk:   '3F3F3F',
  grayIcon: '535353',
  silver:   'A5A5A5',
  rule:     '595959',
  ruleLt:   'F2F2F2',
  ruleGray: 'BFBFBF'
};

const F = {
  light:    'Bricolage Grotesque Light',
  reg:      'Bricolage Grotesque',
  med:      'Bricolage Grotesque Medium',
  syne:     'Syne',
  playfair: 'Playfair Display'
};

/* ------------------------------------------------------------------ helpers */

/** Text box: reference boxes are top anchored with zero internal padding. */
function txt(s, text, o) {
  s.addText(text, Object.assign({
    fontFace: F.light, color: C.white, align: 'left', valign: 'top',
    margin: 0, isTextBox: true
  }, o));
}

/** roundRect radius (inches) from the OOXML "adj" value of the reference shape. */
function rad(adj, w, h) { return (adj / 100000) * Math.min(w, h); }

function roundRect(s, x, y, w, h, fill, adj, extra) {
  s.addShape('roundRect', Object.assign({
    x: x, y: y, w: w, h: h, fill: { color: fill }, rectRadius: rad(adj, w, h)
  }, extra));
}

function rect(s, x, y, w, h, fill, extra) {
  s.addShape('rect', Object.assign({ x: x, y: y, w: w, h: h, fill: { color: fill } }, extra));
}

function ellipse(s, x, y, w, h, fill, transparency) {
  s.addShape('ellipse', {
    x: x, y: y, w: w, h: h,
    fill: transparency ? { color: fill, transparency: transparency } : { color: fill }
  });
}

function line(s, x, y, w, h, color, width, extra) {
  s.addShape('line', Object.assign({
    x: x, y: y, w: w, h: h, line: { color: color, width: width }
  }, extra));
}

/** blockArc / arc built from the reference adj values (60000ths of a degree). */
function blockArc(s, x, y, size, rotate, adj1, adj2, adj3, fill) {
  s.addShape('blockArc', {
    x: x, y: y, w: size, h: size, rotate: rotate,
    angleRange: [adj1 / 60000, adj2 / 60000],
    arcThicknessRatio: adj3 / 50000,
    fill: { color: fill }
  });
}

/** Hairline arc (a blockArc of zero thickness in the reference deck). */
function arcLine(s, x, y, w, h, rotate, adj1, adj2, color) {
  s.addShape('blockArc', {
    x: x, y: y, w: w, h: h, rotate: rotate,
    angleRange: [adj1 / 60000, adj2 / 60000], arcThicknessRatio: Number.MIN_VALUE,
    line: { color: color, width: 0.75 }
  });
}

/** Full ring; `adj` is the OOXML donut thickness value. */
function donut(s, x, y, size, adj, fill) {
  s.addShape('donut', {
    x: x, y: y, w: size, h: size,
    rectRadius: (adj / 100000) * size, fill: { color: fill }
  });
}

/* --------------------------------------------------------- shared furniture */

const NAV_LINKS = [
  ['Home', 3.969, 0.460],
  ['Service', 5.499, 0.575],
  ['Product', 7.144, 0.575],
  ['Contact', 8.789, 0.575]
];

/** Top navigation bar - present on every slide. */
function navBar(s) {
  txt(s, 'Hakira', { x: 0.627, y: 0.267, w: 0.908, h: 0.236, fontFace: F.med, fontSize: 14 });
  roundRect(s, 1.562, 0.313, 0.285, 0.145, C.blue, 50000);
  ellipse(s, 1.722, 0.336, 0.099, 0.099, C.black);
  NAV_LINKS.forEach(function (n) {
    txt(s, n[0], { x: n[1], y: 0.301, w: n[2], h: 0.168, fontFace: F.med, fontSize: 10, align: 'center' });
  });
  txt(s, 'Explore Here', { x: 11.591, y: 0.309, w: 1.115, h: 0.151, fontFace: F.med, fontSize: 9, align: 'right' });
}

/** Pill button "Explore Here" + translucent "S+" bubble. */
function exploreBtn(s, x, y, sw) {
  roundRect(s, x, y, 1.928, 0.456, C.blueBtn, 50000);
  txt(s, 'Explore Here', { x: x + 0.222, y: y + 0.127, w: 0.997, h: 0.202, fontSize: 12, align: 'center' });
  ellipse(s, x + 1.459, y - 0.010, 0.469, 0.469, C.white, 72.16);
  txt(s, 'S+', { x: x + 1.583, y: y + 0.107, w: sw || 0.221, h: 0.236, fontSize: 14, align: 'center' });
}

/** "4.9/5 - Rate for Company - *****" review block. */
function ratingBlock(s, x, y, noCaption) {
  txt(s, '4.9/5', { x: x, y: y, w: 2.091, h: 0.471, fontSize: 28, color: C.red });
  if (!noCaption) txt(s, 'Rate for Company', { x: x + 1.256, y: y - 0.004, w: 1.611, h: 0.185, fontSize: 11 });
  for (var i = 0; i < 5; i++) {
    s.addShape('star5', { x: x + 1.256 + i * 0.1804, y: y + 0.295, w: 0.156, h: 0.168, fill: { color: C.star } });
  }
  txt(s, '5.0', { x: x + 2.208, y: y + 0.307, w: 0.262, h: 0.168, fontSize: 10, align: 'right' });
  txt(s, 'Suspendisse interdum conse.', {
    x: x, y: y + 0.605, w: 2.38, h: 0.273, fontSize: 12, lineSpacingMultiple: 1.5
  });
}

/** White "5.0K" statistics card with the A++ badge and the 95% progress ring. */
function statCard(s, x, y) {
  roundRect(s, x, y, 2.208, 2.660, C.white, 2584);
  roundRect(s, x + 1.721, y + 0.171, 0.343, 0.352, C.black, 10113);
  s.addShape('line', {
    x: x + 1.806, y: y + 0.259, w: 0.171, h: 0.176, flipV: true,
    line: { color: C.white, width: 1.25, endArrowType: 'stealth' }
  });
  txt(s, '5.0K', { x: x + 0.465, y: y + 0.383, w: 1.277, h: 0.404, fontSize: 24, color: C.black, align: 'center' });
  txt(s, 'Suspendisse interdum consectetur libe.', {
    x: x, y: y + 0.933, w: 2.208, h: 0.532, fontSize: 11, color: C.black,
    align: 'center', lineSpacingMultiple: 1.5
  });
  roundRect(s, x - 0.537, y + 1.546, 0.873, 0.773, C.red, 4768);
  txt(s, 'A++', { x: x - 0.723, y: y + 1.672, w: 1.277, h: 0.236, fontSize: 14, align: 'center' });
  txt(s, 'Grade', { x: x - 0.497, y: y + 1.880, w: 0.794, h: 0.255, fontSize: 11, align: 'center', lineSpacingMultiple: 1.5 });
  blockArc(s, x + 0.747, y + 1.668, 0.713, 99.08, 10031575, 5158064, 6478, C.black);
  blockArc(s, x + 0.747, y + 1.668, 0.713, -150, 19438533, 5158064, 6478, C.blue);
  txt(s, '95%', { x: x + 0.759, y: y + 1.949, w: 0.689, h: 0.151, fontSize: 9, color: C.black, align: 'center' });
  txt(s, 'Suspendisse ', {
    x: x + 0.848, y: y + 2.098, w: 0.510, h: 0.093, fontSize: 4, color: C.black,
    align: 'center', lineSpacingMultiple: 1.5
  });
}

const SYNE_BODY = 'Suspendisse interdum ce libero id fauus nisl tinu Au risus quvarius qua.';

/** Small Syne headline card (blue / red / white variants). */
function syneCard(s, o) {
  var h = o.h || 2.217;
  var color = o.color || C.white;
  roundRect(s, o.x, o.y, 2.450, h, o.fill, 5977);
  txt(s, o.title, {
    x: o.x + 0.253, y: o.y + 0.336, w: o.titleW || 1.628, h: 0.606,
    fontFace: F.syne, fontSize: 18, color: color
  });
  txt(s, o.body || SYNE_BODY, {
    x: o.x + 0.271, y: o.y + 1.083 + (h - 2.217) * 0.56, w: 2.106, h: 0.878,
    fontFace: F.syne, fontSize: 12, color: color, lineSpacingMultiple: 1.5
  });
}

const TRIO_BODY = 'Suspendisse interdum consectetur libero id faucibus nisl tinu Arcu risus quvarius qua interdum consectetur libero .';

/** Row of three wide cards (blue / red / white) used on the platform slides. */
function trioCards(s, y, titles) {
  var cols = [
    { x: 0.642, fill: C.blue, color: C.white, titleW: 3.354, bodyDx: 0.374 },
    { x: 4.761, fill: C.red, color: C.white, titleW: 3.234, bodyDx: 0.365 },
    { x: 8.888, fill: C.white, color: C.black, titleW: 2.931, bodyDx: 0.347 }
  ];
  cols.forEach(function (col, i) {
    roundRect(s, col.x, y, 3.888, 2.155, col.fill, 13665);
    txt(s, titles[i], {
      x: col.x + 0.378, y: y + 0.358, w: col.titleW, h: 0.337, fontSize: 20, color: col.color
    });
    txt(s, TRIO_BODY, {
      x: col.x + col.bodyDx, y: y + 0.971, w: 3.302, h: 0.878, fontSize: 12,
      color: col.color, lineSpacingMultiple: 1.5
    });
  });
}

/**
 * Continent silhouette of the "Followers Growth" slide, one row of the bitmap
 * per line; horizontal runs of '#' become a single rectangle.
 */
const WORLD_MAP = [
  '.........................###......#######.........................................................',
  '......................######################......................................................',
  '.....................######################.........##................##..........................',
  '..................#..#####..##############........###..................###........................',
  '...............#....#..###.###############.........#.....................#........................',
  '.............####.##.####...##############..................###......#######.......##.............',
  '............##.#.....####.......##########.................##.......########......................',
  '............######.##.####......#########.................##...##################...##............',
  '..##..........####..#.######.....#########...........##....#...##########################.........',
  '.##################.##.##.###....########..........####.....###################################...',
  '.########################..###..######............################################################',
  '########################..####...####....###......#############################################.#.',
  '.#####################.#...##....###.............##############################################...',
  '.####################....###......##............########################################.####.....',
  '..##....##############...#####..................#.##################################.....#........',
  '..#......##############...####...............#...##################################.....##........',
  '..........######################............###..###################################....#.........',
  '...........#####################............###.#####################################.............',
  '............##################.#..............######################################..............',
  '............##################................#####################################...............',
  '............################.#................###########.#########################.#.............',
  '............################................###...########.######################.................',
  '.............#############..................###....#######.###################.#...#..............',
  '.............#############...................#####....########################...###..............',
  '..............###########...................#######.##########################....................',
  '...............##########...................##################################....................',
  '...............#####....#.................################.###################....................',
  '.................###......................###################..##############.....................',
  '..................##..#..#................##################....####..####........................',
  '...................####...................##################....###....###........................',
  '.....................###..................#################......##.....###.......................',
  '.......................#...##.............###################....##......#........................',
  '........................#.######...........##################.....#.....#......#..................',
  '..........................########..........#...############...........##...#.....................',
  '.........................#########...............##########.............#..###....................',
  '.........................###########.............#########..............##.####..###..............',
  '.........................#############............########.........................###............',
  '..........................############............########..........................#.#...........',
  '..........................############............########.......................##.#.............',
  '..........................###########.............########.##..................####.#.............',
  '............................#########.............#######..##.................########............',
  '............................#########.............#######..#................###########...........',
  '............................#######...............######...#................############..........',
  '............................#######................#####....................############..........',
  '............................######.................####......................###########..........',
  '............................#####...................##.......................##...#####...........',
  '...........................#####....................................................###...........',
  '...........................####..............................................................#....',
  '...........................###.......................................................#......#.....',
  '...........................###.............................................................##.....',
  '...........................##.....................................................................',
  '...........................##.....................................................................',
  '...........................##.....................................................................',
  '............................#.....................................................................'
];

function worldMap(s, x0, y0, w, h) {
  const cw = w / WORLD_MAP[0].length;
  const ch = h / WORLD_MAP.length;
  WORLD_MAP.forEach(function (row, r) {
    var c = 0;
    while (c < row.length) {
      if (row[c] === '#') {
        var end = c;
        while (row[end + 1] === '#') end++;
        rect(s, x0 + c * cw, y0 + r * ch, (end - c + 1) * cw, ch * 1.1, C.gray);
        c = end + 1;
      } else {
        c++;
      }
    }
  });
}

/** Small "chart" glyph standing in for the raster icon of the reference deck. */
function chartIconPlaceholder(s, x, y, size) {
  var u = size / 6;
  rect(s, x + u * 0.4, y + u * 3.2, u, u * 2.4, C.black);
  rect(s, x + u * 2.0, y + u * 2.0, u, u * 3.6, C.black);
  rect(s, x + u * 3.6, y + u * 0.8, u, u * 4.8, C.black);
}

/* ---------------------------------------------------------------- the deck */

const build = [];

/* 1 - cover ---------------------------------------------------------------- */
build.push(function (s) {
  txt(s, 'Social', { x: 0.627, y: 1.204, w: 4.752, h: 1.346, fontSize: 80 });
  txt(s, 'Media', { x: 0.627, y: 2.289, w: 3.545, h: 1.346, fontSize: 80 });
  txt(s, 'Business', { x: 0.627, y: 3.336, w: 5.414, h: 1.346, fontSize: 80 });
  txt(s, 'Suspendisse interdum c libero id faucibus nisl tincidu Arcu risus quvarius qua.', {
    x: 4.281, y: 2.553, w: 2.112, h: 0.878, fontFace: F.reg, fontSize: 12, lineSpacingMultiple: 1.5
  });

  // red highlight card
  roundRect(s, 7.057, 3.517, 2.208, 1.596, C.redCard, 11539, { rotate: 0.49 });
  roundRect(s, 8.679, 3.723, 0.343, 0.352, C.white, 20292, { rotate: 0.49 });
  chartIconPlaceholder(s, 8.758, 3.805, 0.186);
  txt(s, '5.3K', { x: 7.423, y: 3.756, w: 1.277, h: 0.404, fontSize: 24, align: 'center', rotate: 0.49 });
  txt(s, 'Suspendisse interdum consectetur libe.', {
    x: 7.079, y: 4.350, w: 2.208, h: 0.528, fontSize: 11, align: 'center', rotate: 0.49, lineSpacingMultiple: 1.5
  });

  exploreBtn(s, 0.627, 5.763, 0.285);
  ratingBlock(s, 3.530, 5.496);

  // footer counters
  var stats = [['9213+', 'Best Company', 8.242, 8.195, 1.191],
               ['98%', 'Value Company', 9.947, 9.858, 1.303],
               ['8650+', 'Best Client ', 11.652, 11.605, 1.191]];
  stats.forEach(function (st) {
    txt(s, st[0], { x: st[2], y: 6.373, w: 1.097, h: 0.438, fontFace: F.syne, fontSize: 20, align: 'center', margin: [7.2, 7.2, 3.6, 3.6] });
    txt(s, st[1], { x: st[3], y: 7.023, w: st[4], h: 0.185, fontFace: F.syne, fontSize: 11, align: 'center' });
  });
  line(s, 9.643, 6.291, 0, 0.999, C.ruleGray, 0.75);
  line(s, 11.348, 6.291, 0, 0.999, C.ruleGray, 0.75);
  line(s, 0.708, 7.208, 5.366, 0, C.white, 0.75);

  // decorative wordmark parked off-canvas in the reference deck
  txt(s, 'media', { x: 11.109, y: -1.820, w: 3.196, h: 1.212, fontFace: F.playfair, fontSize: 72, italic: true });
  txt(s, 'media', { x: 11.109, y: -2.867, w: 3.196, h: 1.212, fontFace: F.playfair, fontSize: 72, italic: true });
});

/* 2 - executive summary ---------------------------------------------------- */
build.push(function (s) {
  txt(s, 'Executive', { x: 0.627, y: 1.420, w: 6.040, h: 1.346, fontSize: 80 });
  txt(s, 'Summary', { x: 0.627, y: 2.505, w: 5.962, h: 1.346, fontSize: 80 });

  roundRect(s, 7.436, 1.611, 5.340, 2.617, C.blueBtn, 7051);
  txt(s, 'Top Performing Platform', { x: 7.950, y: 1.997, w: 2.776, h: 0.673, fontSize: 20 });
  txt(s, 'Suspendisse interdum consectetur libero id faucis nisl tinu Arcu risus quvarius qua quiue inte consec interdum consectetur interdum consectetur.', {
    x: 7.948, y: 2.963, w: 4.315, h: 0.879, fontSize: 12, lineSpacingMultiple: 1.5
  });

  roundRect(s, 0.627, 4.506, 5.336, 2.617, C.red, 7051);
  txt(s, 'Major Campaign Achievements', { x: 1.129, y: 4.893, w: 2.724, h: 0.673, fontSize: 20 });
  txt(s, 'Suspendisse interdum consectetur libero id faucis nisl tinu Arcu risus quvarius qua quiue inte consec interdum consectetur interdum consectetur libero.', {
    x: 1.127, y: 5.858, w: 4.336, h: 0.879, fontSize: 12, lineSpacingMultiple: 1.5
  });

  txt(s, 'Suspendisse interdum consectetur libero id faucibus nisl tincidu Arcu risus quvarius qua quisque id diam vel. Volutpat Suspendie interdum consectetur libero id faucibus nisl tincidu Arcu .', {
    x: 7.436, y: 4.684, w: 5.271, h: 0.884, fontFace: F.reg, fontSize: 12, lineSpacingMultiple: 1.5
  });
  exploreBtn(s, 7.436, 6.514, 0.285);
  ratingBlock(s, 10.339, 6.247);
});

/* 3 - table of content ----------------------------------------------------- */
build.push(function (s) {
  txt(s, 'Table of', { x: 0.627, y: 1.420, w: 5.151, h: 1.346, fontSize: 80 });
  txt(s, 'Content', { x: 0.627, y: 2.505, w: 5.962, h: 1.346, fontSize: 80 });

  var toc = [['01', 'Content One', 0.627, 5.150], ['02', 'Content Two', 0.627, 6.532],
             ['03', 'Content Three', 7.280, 5.150], ['04', 'Content Four', 7.280, 6.532]];
  toc.forEach(function (t) {
    var x = t[2], y = t[3];
    txt(s, t[0], { x: x, y: y, w: 0.886, h: 0.471, fontFace: F.reg, fontSize: 28, color: C.red });
    txt(s, t[1], { x: x + 0.737, y: y + 0.101, w: 1.584, h: 0.269, fontFace: F.reg, fontSize: 16 });
    txt(s, 'Suspendisse interdum cocr libero id faucibus nisl.', {
      x: x + 3.041, y: y - 0.053, w: 2.541, h: 0.576, fontFace: F.reg, fontSize: 12, lineSpacingMultiple: 1.5
    });
  });
  line(s, 0.627, 6.102, 5.447, 0, C.ruleLt, 0.75);
  line(s, 7.280, 6.102, 5.447, 0, C.ruleLt, 0.75);

  ratingBlock(s, 6.409, 2.267);
  exploreBtn(s, 6.409, 3.661, 0.285);
});

/* 4 - introduction --------------------------------------------------------- */
build.push(function (s) {
  txt(s, 'Introduction', { x: 0.627, y: 1.205, w: 7.092, h: 1.111, fontSize: 66 });
  txt(s, 'To Social Media', { x: 0.627, y: 2.142, w: 7.092, h: 1.111, fontSize: 66 });
  txt(s, 'Best Social Media', { x: 5.717, y: 4.084, w: 3.760, h: 0.337, fontFace: F.reg, fontSize: 20 });
  txt(s, 'Suspendisse interdum consectetur libero id faucibus nisl tinu Arcu risus quvarius qua quisque id diam vel. Volutpat Suspene interdum consectetur libero id faucibus nisl tincidu Arcu tinu Arcu risus quvarius qua quisque id diam vel. Volutpat Suspene interdum consectetur.', {
    x: 5.717, y: 4.831, w: 6.989, h: 0.878, fontFace: F.reg, fontSize: 12, lineSpacingMultiple: 1.5
  });
  exploreBtn(s, 5.717, 6.505, 0.262);
  ratingBlock(s, 8.620, 6.238);
  statCard(s, 10.505, 1.507);
});

/* 5 - social media landscape ----------------------------------------------- */
build.push(function (s) {
  txt(s, 'Social Media', { x: 6.830, y: 1.128, w: 6.207, h: 1.111, fontSize: 66 });
  txt(s, 'Landscape', { x: 6.830, y: 2.066, w: 6.207, h: 1.111, fontSize: 66 });
  txt(s, 'Suspendisse interdum consectetur libero id faucibus nisl tincidu Arcu risus quvarius qua quisque id diam vel. Volutpat Suspendie interdum consectetur libero id faucibus nisl tincidu Arcu aucibus nisl.', {
    x: 0.675, y: 1.819, w: 5.156, h: 0.884, fontFace: F.reg, fontSize: 12, lineSpacingMultiple: 1.5
  });
  ratingBlock(s, 0.675, 3.414);
  exploreBtn(s, 6.830, 3.602, 0.285);
  syneCard(s, { x: 0.675, y: 4.848, fill: C.blueBtn, title: 'Active Use Count' });
  syneCard(s, { x: 3.380, y: 4.848, fill: C.red, title: 'Emerging Trends' });
});

/* 6 - objectives and goals ------------------------------------------------- */
build.push(function (s) {
  var body = 'Suspendisse interdum consectetur libero id faucibus nisl tinu Arcu risus quvarius qua quisque id diam vel. Volutpat Suspene.';
  txt(s, '57%', { x: 0.627, y: 1.149, w: 1.564, h: 1.010, fontFace: F.playfair, fontSize: 60, color: C.red });
  txt(s, 'Driving Website Traffic', { x: 2.799, y: 1.348, w: 2.361, h: 0.673, fontFace: F.playfair, fontSize: 20 });
  txt(s, body, { x: 2.799, y: 2.489, w: 3.390, h: 0.878, fontFace: F.reg, fontSize: 12, lineSpacingMultiple: 1.5 });

  txt(s, '76%', { x: 7.144, y: 1.149, w: 1.564, h: 1.010, fontFace: F.playfair, fontSize: 60, color: C.red });
  txt(s, [{ text: 'Increase Brand', options: { breakLine: true } }, { text: 'Awareness' }],
    { x: 9.316, y: 1.348, w: 3.119, h: 0.673, fontFace: F.playfair, fontSize: 20, color: C.white, align: 'left', valign: 'top', margin: 0 });
  txt(s, body, { x: 9.316, y: 2.489, w: 3.390, h: 0.878, fontFace: F.reg, fontSize: 12, lineSpacingMultiple: 1.5 });

  txt(s, 'About', { x: 7.144, y: 3.755, w: 2.966, h: 1.010, fontSize: 60 });
  txt(s, 'Objectives', { x: 7.144, y: 4.557, w: 5.122, h: 1.010, fontSize: 60 });
  txt(s, 'And Goals', { x: 7.144, y: 5.402, w: 5.122, h: 1.010, fontSize: 60 });
  line(s, 7.144, 6.916, 5.084, 0, C.white, 0.75);

  exploreBtn(s, 0.627, 6.306, 0.262);
  ratingBlock(s, 3.530, 6.039);
});

/* 7 - audience demographic -------------------------------------------------- */
build.push(function (s) {
  txt(s, 'Audience', { x: 0.627, y: 1.534, w: 6.207, h: 0.841, fontSize: 50 });
  txt(s, 'Demographic', { x: 0.627, y: 2.271, w: 4.872, h: 0.841, fontSize: 50 });

  // donut breakdown
  blockArc(s, 0.615, 3.393, 3.689, -81.11, 15088885, 10333615, 25224, C.white);
  blockArc(s, 0.615, 3.393, 3.689, -81.11, 6958843, 16038290, 25648, C.red);
  blockArc(s, 0.615, 3.393, 3.689, -81.11, 20039252, 7871769, 25514, C.blue);
  txt(s, '+2720', { x: 1.931, y: 4.874, w: 1.090, h: 0.337, fontFace: F.reg, fontSize: 20, bold: true, align: 'center' });
  txt(s, 'Suspendisse inter consecteturid.', {
    fontFace: F.reg, x: 1.853, y: 5.309, w: 1.247, h: 0.429, fontSize: 9, align: 'center', lineSpacingMultiple: 1.5
  });

  var cols = [
    { x: 5.933, head: 'Gender Distribution', big: '+16.70', note: 'Suspendisse interdum consectetur libero id faucibus nisl tincid.', hw: 1.975 },
    { x: 9.794, head: 'Age Distribution', big: '+4500', note: 'Suspendisse interdum consectetur libero id faucibus nisl tincidu.', hw: 1.839 }
  ];
  cols.forEach(function (c) {
    txt(s, c.head, { fontFace: F.reg, x: c.x, y: 2.177, w: c.hw, h: 0.318, fontSize: 14, lineSpacingMultiple: 1.5 });
    txt(s, c.big, { fontFace: F.reg, x: c.x, y: 2.620, w: 1.839, h: 0.638, fontSize: 28, lineSpacingMultiple: 1.5 });
    txt(s, c.note, { fontFace: F.reg, x: c.x, y: 3.407, w: 2.940, h: 0.524, fontSize: 11, lineSpacingMultiple: 1.5 });
  });

  txt(s, 'Detailed Chart', { fontFace: F.reg, x: 5.933, y: 4.314, w: 3.979, h: 0.456, fontSize: 20, lineSpacingMultiple: 1.5 });

  // two progress bars with callout bubbles
  var bars = [
    { label: 'Gender', ly: 5.281, y: 5.708, redX: 10.110, redW: 2.557, blueW: 4.399, cx: 9.911, cy: 4.972, cLab: '+16.70' },
    { label: 'Age', ly: 6.414, y: 6.841, redX: 8.540, redW: 4.127, blueW: 3.319, cx: 8.873, cy: 6.105, cLab: '+4500' }
  ];
  bars.forEach(function (b) {
    txt(s, b.label, { fontFace: F.reg, x: 5.933, y: b.ly, w: 1.154, h: 0.236, fontSize: 14 });
    roundRect(s, b.redX, b.y, b.redW, 0.14, C.red, 50000);
    roundRect(s, 5.933, b.y, b.blueW, 0.14, C.blue, 50000);
    s.addShape('wedgeRectCallout', { x: b.cx, y: b.cy, w: 1.184, h: 0.384, fill: { color: C.white } });
    txt(s, b.cLab, { fontFace: F.reg, x: b.cx + 0.015, y: b.cy + 0.098, w: 1.154, h: 0.185, fontSize: 11, color: C.black, align: 'center' });
  });
});

/* 8 - engagement matrix ----------------------------------------------------- */
build.push(function (s) {
  // horizontal bar chart drawn from plain rectangles
  [1.376, 2.361, 3.346, 4.331].forEach(function (gx) { line(s, gx, 1.970, 0, 1.696, C.gray, 1.25); });
  var bars = [[2.054, 1.379, C.red], [2.172, 2.219, C.blue], [2.490, 0.890, C.red], [2.609, 1.730, C.blue],
              [2.913, 2.170, C.red], [3.031, 1.250, C.blue], [3.341, 1.181, C.red], [3.459, 2.125, C.blue]];
  bars.forEach(function (b) { rect(s, 1.381, b[0], b[1], 0.115, b[2]); });

  var qtrLabels = [['4st', 2.085], ['3st', 2.513], ['2st', 2.935]];
  qtrLabels.forEach(function (q) {
    txt(s, [{ text: q[0], options: { superscript: true } }, { text: ' Qtr' }],
      { x: 0.667, y: q[1], w: 0.550, h: 0.185, fontSize: 11, color: C.white, valign: 'top', margin: 0 });
  });
  txt(s, [{ text: '1' }, { text: 'st', options: { superscript: true } }, { text: ' Qtr' }],
    { x: 0.667, y: 3.363, w: 0.550, h: 0.185, fontSize: 11, color: C.white, valign: 'top', margin: 0 });
  [['0', 1.101], ['2', 2.086], ['4', 3.071], ['6', 4.056]].forEach(function (a) {
    txt(s, a[0], { x: a[1], y: 3.749, w: 0.550, h: 0.185, fontSize: 11, align: 'center' });
  });

  txt(s, 'Engagement', { x: 0.667, y: 4.935, w: 5.507, h: 0.942, fontSize: 56 });
  txt(s, 'Matrix', { x: 0.667, y: 5.799, w: 3.479, h: 0.942, fontSize: 56 });

  var note = 'Suspendisse interdum consectr libero id faucibs nisl tinu .';
  var quarters = [['1', 'st', 6.420, 2.138, 6.451, 2.627], ['2', 'nd', 6.420, 4.079, 6.414, 4.512],
                  ['3', 'rd', 6.414, 5.891, 6.451, 6.381], ['4', 'th', 10.093, 5.872, 10.093, 6.381]];
  quarters.forEach(function (q) {
    txt(s, [{ text: q[0] }, { text: q[1], options: { superscript: true } }, { text: ' Quarter' }],
      { x: q[2], y: q[3], w: 1.808, h: 0.337, fontSize: 20, color: C.white, valign: 'top', margin: 0 });
    txt(s, note, { x: q[4], y: q[5], w: 2.742, h: 0.576, fontSize: 12, lineSpacingMultiple: 1.5 });
  });
});

/* 9 - followers growth ------------------------------------------------------ */
build.push(function (s) {
  txt(s, 'Followers', { x: 0.627, y: 1.109, w: 6.207, h: 1.111, fontSize: 66 });
  txt(s, 'Growth', { x: 0.627, y: 1.978, w: 4.872, h: 1.111, fontSize: 66 });
  worldMap(s, 0.602, 3.670, 5.910, 3.665);

  [[1.400, 4.590, C.redSoft, '$2M'], [4.441, 4.422, C.blue, '$5M']].forEach(function (p) {
    ellipse(s, p[0], p[1], 0.802, 0.802, p[2], 41.18);
    ellipse(s, p[0] + 0.117, p[1] + 0.117, 0.570, 0.570, p[2]);
    txt(s, p[3], { x: p[0] + 0.126, y: p[1] + 0.291, w: 0.551, h: 0.252, fontSize: 9, align: 'center' });
  });

  var panels = [
    { y: 2.630, ringY: 3.135, ringX: 7.893, pct: '64%', name: 'ASIA', rot: 141.74, a1: 6577277, color: C.redSoft },
    { y: 5.046, ringY: 5.593, ringX: 7.950, pct: '35%', name: 'EUROPA', rot: 30.01, a1: 14926721, color: C.blue }
  ];
  panels.forEach(function (p) {
    roundRect(s, 7.822, p.y, 4.947, 2.026, C.panel, 5664);
    donut(s, p.ringX, p.ringY, 1.149, 12303, C.gray);
    blockArc(s, p.ringX, p.ringY, 1.149, p.rot, p.a1, 20509, 12314, p.color);
    txt(s, p.pct, { x: 8.175, y: p.y + 0.971, w: 0.752, h: 0.303, fontSize: 12, color: C.black, align: 'center', margin: [7.2, 7.2, 3.6, 3.6] });
    txt(s, p.name, { x: 9.683, y: p.y + 0.412, w: 2.374, h: 0.269, fontSize: 16, color: C.black });
    txt(s, 'Suspendisse interdum consectetur libero id faucibus nisl tinu Arcu risus quvarius qua quisque id diam vel. ', {
      x: 9.654, y: p.y + 0.813, w: 2.816, h: 0.877, fontSize: 12, color: C.black, lineSpacingMultiple: 1.5
    });
  });
});

/* 10 - content performance -------------------------------------------------- */
build.push(function (s) {
  txt(s, 'Content Performance', { x: 0.627, y: 1.144, w: 9.910, h: 0.808, fontSize: 48 });

  // connecting snake
  line(s, 1.454, 2.622, 10.340, 0, C.white, 0.75);
  line(s, 1.524, 4.164, 10.268, 0.003, C.white, 0.75);
  line(s, 1.454, 5.699, 10.340, 0, C.white, 0.75);
  line(s, 5.934, 7.240, 5.861, 0, C.white, 0.75);
  arcLine(s, 10.979, 2.650, 1.542, 1.485, 90, 10946882, 21336871, C.white);
  arcLine(s, 10.979, 5.727, 1.542, 1.485, 90, 10946882, 21336871, C.white);
  arcLine(s, 0.812, 4.193, 1.542, 1.485, -90, 11222027, 21336871, C.white);
  ellipse(s, 1.283, 2.544, 0.171, 0.156, C.white);
  ellipse(s, 5.798, 7.163, 0.171, 0.156, C.white);

  var steps = [
    { n: '01', x: 1.909, y: 3.024, fill: C.red, num: C.white, tx: 2.954, ty: 3.026, by: 3.376, title: 'Consultation' },
    { n: '02', x: 6.756, y: 3.024, fill: C.blue, num: C.white, tx: 7.759, ty: 2.977, by: 3.411, title: 'Planning', tw: 1.947 },
    { n: '03', x: 4.294, y: 4.612, fill: C.grayMid, num: C.white, tx: 5.298, ty: 4.565, by: 4.999, title: 'System Design' },
    { n: '04', x: 9.140, y: 4.612, fill: C.grayLt, num: C.black, tx: 10.144, ty: 4.565, by: 4.999, title: 'Installation' },
    { n: '05', x: 1.950, y: 6.152, fill: C.gray, num: C.black, tx: 2.954, ty: 6.105, by: 6.539, title: 'Description Here' },
    { n: '06', x: 6.797, y: 6.152, fill: C.white, num: C.black, tx: 7.800, ty: 6.105, by: 6.539, title: 'Description Here' }
  ];
  steps.forEach(function (t) {
    roundRect(s, t.x, t.y, 0.757, 0.757, t.fill, 16667);
    roundRect(s, t.x + 0.489, t.y - 0.008, 0.268, 0.275, C.grayIcon, 20292);
    s.addShape('line', {
      x: t.x + 0.551, y: t.y + 0.055, w: 0.145, h: 0.149, flipV: true,
      line: { color: C.white, width: 1.25, endArrowType: 'stealth' }
    });
    txt(s, t.n, { x: t.x + 0.089, y: t.y + 0.260, w: 0.517, h: 0.404, fontSize: 24, color: t.num, align: 'center' });
    txt(s, t.title, { x: t.tx, y: t.ty, w: t.tw || 1.486, h: 0.202, fontSize: 12 });
    txt(s, 'Suspendisse interdum consectetur libero id.', {
      x: t.tx, y: t.by, w: 2.162, h: 0.508, fontSize: 10.5, lineSpacingMultiple: 1.5
    });
  });
});

/* 11 - research and impression ---------------------------------------------- */
build.push(function (s) {
  txt(s, 'Research And', { x: 0.627, y: 1.200, w: 5.666, h: 1.010, fontSize: 60 });
  txt(s, 'Impression', { x: 0.627, y: 2.027, w: 5.666, h: 1.010, fontSize: 60 });
  txt(s, 'Suspendisse interdum consectetur libero id faucibus nisl tincidu Arcu risus quvarius qua quisque id diam vel. Volutpat Suspendie interdum consectetur.', {
    x: 0.663, y: 3.333, w: 5.666, h: 0.575, fontFace: F.reg, fontSize: 12, lineSpacingMultiple: 1.5
  });
  statCard(s, 1.163, 4.452);
  syneCard(s, { x: 3.777, y: 4.712, fill: C.blueBtn, title: 'Social Media Post' });
  syneCard(s, { x: 10.256, y: 1.551, fill: C.red, title: 'Campaign' });
  ratingBlock(s, 7.229, 2.267, true);
  exploreBtn(s, 7.229, 3.661);
});

/* 12 - hashtag analysis ------------------------------------------------------ */
build.push(function (s) {
  trioCards(s, 1.340, ['Impact', 'Effectiveness', 'Description Here']);
  txt(s, 'Hashtag', { x: 0.627, y: 4.001, w: 4.447, h: 1.212, fontSize: 72 });
  txt(s, 'Analysis', { x: 0.627, y: 4.961, w: 4.276, h: 1.212, fontSize: 72 });
  txt(s, 'Suspendisse interdum consectetur libero id faucibus nisl tincidu Arcu risus quvarius qua quisque id diam vel. Volutpat Suspendie interdum.', {
    x: 0.627, y: 6.430, w: 5.003, h: 0.575, fontFace: F.reg, fontSize: 12, lineSpacingMultiple: 1.5
  });
  exploreBtn(s, 7.006, 6.306);
  ratingBlock(s, 9.909, 6.039);
});

/* 13 - social media platform ------------------------------------------------- */
build.push(function (s) {
  txt(s, 'Social Media', { x: 0.627, y: 1.503, w: 6.040, h: 1.010, fontSize: 60 });
  txt(s, 'Platform', { x: 0.627, y: 2.357, w: 4.276, h: 1.010, fontSize: 60 });
  txt(s, 'Suspendisse interdum consectetur libero id faucibus nisl tincidu Arcu risus quvarius qua quisque id diam vel. Volutpat Suspendie interdum.', {
    x: 0.627, y: 3.595, w: 5.003, h: 0.575, fontFace: F.reg, fontSize: 12, lineSpacingMultiple: 1.5
  });
  txt(s, 'Suspendisse interdum consectetur libero id faucibus nisl tincidu Arcu risus quvarius qua quisque id diam vel. Volutpat Suspendie interdum consectetur libero id faucibus nisl tincidu Arcu aucibus nisl libero id faucibus nisl.', {
    x: 7.006, y: 1.819, w: 5.770, h: 0.884, fontFace: F.reg, fontSize: 12, lineSpacingMultiple: 1.5
  });
  exploreBtn(s, 7.006, 3.545);
  ratingBlock(s, 9.909, 3.278);
  trioCards(s, 4.729, ['Platform One', 'Platform Two', 'Platform Three']);
});

/* 14 - platform one analysis -------------------------------------------------- */
build.push(function (s) {
  txt(s, 'Platform One', { x: 6.905, y: 1.387, w: 6.040, h: 1.010, fontSize: 60 });
  txt(s, 'Analysis', { x: 6.905, y: 2.241, w: 4.276, h: 1.010, fontSize: 60 });
  txt(s, '8765.90', { x: 0.627, y: 1.879, w: 2.202, h: 0.404, fontFace: F.playfair, fontSize: 24 });
  txt(s, 'Suspendisse interdum consectetur libero id faucibs nisl tinu Arcu risus quvas qua quisque id diam vel. Volutpat Suspene interdum consectetur libero id consectetur libero id faucibs nisl tinu Arcu risus quvas .', {
    x: 0.627, y: 2.664, w: 5.225, h: 0.878, fontFace: F.reg, fontSize: 12, lineSpacingMultiple: 1.5
  });
  txt(s, 'Suspendisse interdum consectetur libero id faucibus nisl tincidu Arcu risus quvarius qua quisque id diam vel. Volutpat Suspendie interdum.', {
    x: 6.905, y: 3.591, w: 5.003, h: 0.575, fontFace: F.reg, fontSize: 12, lineSpacingMultiple: 1.5
  });
  syneCard(s, { x: 6.953, y: 4.848, fill: C.blueBtn, title: 'Best Feature' });
  syneCard(s, { x: 9.658, y: 4.848, fill: C.red, title: 'Coreaction Feature' });
});

/* 15 - platform two analysis -------------------------------------------------- */
build.push(function (s) {
  txt(s, 'Platform Two', { x: 0.627, y: 1.534, w: 6.040, h: 1.010, fontSize: 60 });
  txt(s, 'Analysis', { x: 0.627, y: 2.388, w: 4.276, h: 1.010, fontSize: 60 });
  syneCard(s, { x: 7.551, y: 2.003, fill: C.blueBtn, title: 'Best Feature' });
  syneCard(s, { x: 10.256, y: 2.003, fill: C.red, title: 'Coreaction Feature' });
  txt(s, 'Suspendisse interdum consectetur libero id faucibus nisl tinu Arcu risus quvarius qua quisque id diam vel. Volutpat Suspene interdum consectetur libero id faucibus nisl tincidu Arcu tinu Arcu risus quvarius qua quisque.', {
    x: 0.627, y: 4.220, w: 5.614, h: 0.878, fontFace: F.reg, fontSize: 12, lineSpacingMultiple: 1.5
  });
  exploreBtn(s, 0.627, 6.084);
  ratingBlock(s, 3.530, 5.817);
});

/* 16 - platform three analysis ------------------------------------------------ */
build.push(function (s) {
  txt(s, 'Platform Three', { x: 6.937, y: 1.287, w: 6.040, h: 1.010, fontSize: 60 });
  txt(s, 'Analysis', { x: 6.937, y: 2.140, w: 4.276, h: 1.010, fontSize: 60 });
  var cardBody = 'Suspendisse interdum consectetur libero id faucis nisl tinu Arcu risus quvarius qua quiue inte consec interdum consectetur interdum consectetur libero.';

  roundRect(s, 0.627, 1.458, 5.336, 2.617, C.red, 7051);
  txt(s, 'Best Feature', { x: 1.129, y: 2.010, w: 2.724, h: 0.337, fontSize: 20 });
  txt(s, cardBody, { x: 1.127, y: 2.682, w: 4.336, h: 0.879, fontSize: 12, lineSpacingMultiple: 1.5 });

  roundRect(s, 6.937, 4.506, 5.336, 2.617, C.blue, 7051);
  txt(s, 'Corection Feature', { x: 7.439, y: 5.059, w: 2.724, h: 0.337, fontSize: 20 });
  txt(s, cardBody, { x: 7.437, y: 5.731, w: 4.336, h: 0.879, fontSize: 12, lineSpacingMultiple: 1.5 });

  txt(s, 'Suspendisse interdum consectetur libero id faucibus nisl tincidu Arcu risus quvarius qua quisque id diam vel. Volutpat Suspendie interdum consectetur libero id faucibus nisl tincidu Arcu .', {
    x: 0.627, y: 4.684, w: 5.271, h: 0.884, fontFace: F.reg, fontSize: 12, lineSpacingMultiple: 1.5
  });
  txt(s, 'Suspendisse interdum consectetur libero id faucibus nisl tincidu Arcu risus quvarius qua quisque id diam vel. Volutpat Suspendie interdum consectetur.', {
    x: 6.937, y: 3.463, w: 5.770, h: 0.575, fontFace: F.reg, fontSize: 12, lineSpacingMultiple: 1.5
  });
  exploreBtn(s, 0.627, 6.514);
  ratingBlock(s, 3.530, 6.247);
});

/* 17 - influencer marketing ---------------------------------------------------- */
build.push(function (s) {
  txt(s, 'Influencer', { x: 0.640, y: 1.342, w: 5.155, h: 1.212, fontSize: 72 });
  txt(s, 'Marketing', { x: 0.640, y: 2.325, w: 4.736, h: 1.212, fontSize: 72 });
  txt(s, 'Aaron Loeb', { x: 0.640, y: 4.006, w: 4.218, h: 0.539, fontFace: F.playfair, fontSize: 32 });
  txt(s, 'Suspendisse interdum consectetur libero id faucibus nisl tincidu Arcu risus quvari qua quisque id diam vel. Volutpat interdum consectetur libero id faucibu.', {
    x: 0.640, y: 4.895, w: 5.564, h: 0.575, fontFace: F.reg, fontSize: 12, lineSpacingMultiple: 1.5
  });
  exploreBtn(s, 0.627, 6.514);
  ratingBlock(s, 3.530, 6.247);

  txt(s, 'Performance', { x: 7.287, y: 5.259, w: 2.011, h: 0.438, fontFace: F.syne, fontSize: 20, margin: [7.2, 7.2, 3.6, 3.6] });
  var meters = [
    { label: 'Popularitas', ly: 6.061, lw: 2.026, x: 7.265, y: 6.382, full: 5.441, done: 4.061 },
    { label: 'Social Media ', ly: 6.738, lw: 1.367, x: 7.275, y: 7.043, full: 5.431, done: 2.980 }
  ];
  meters.forEach(function (m) {
    txt(s, m.label, { x: 7.287, y: m.ly, w: m.lw, h: 0.202, fontFace: F.syne, fontSize: 12 });
    line(s, m.x, m.y, m.full, 0, C.red, 7);
    line(s, m.x, m.y, m.done, 0, C.blue, 7);
  });
});

/* 18 - competitor analysis ------------------------------------------------------ */
build.push(function (s) {
  txt(s, 'Competitor Analysis', { x: 0.548, y: 1.074, w: 8.093, h: 0.909, fontSize: 54 });

  var quads = [
    { x: 1.011, y: 2.490, fill: C.white, adj: 5731, label: 'PRODUCT', lx: 2.051, ly: 2.636 },
    { x: 5.173, y: 2.490, fill: C.gray, adj: 4913, label: 'PRICE', lx: 6.213, ly: 2.632 },
    { x: 1.011, y: 4.758, fill: C.gray, adj: 5731, label: 'PROMOTION', lx: 2.051, ly: 4.922 },
    { x: 5.173, y: 4.758, fill: C.white, adj: 5732, label: 'PLACE', lx: 6.213, ly: 4.918 }
  ];
  quads.forEach(function (q) {
    roundRect(s, q.x, q.y, 4.004, 2.106, q.fill, q.adj);
    txt(s, q.label, { x: q.lx, y: q.ly, w: 1.924, h: 0.286, fontSize: 11, color: C.black, align: 'center', margin: [7.2, 7.2, 3.6, 3.6] });
  });

  // axes
  s.addShape('line', { x: 0.605, y: 2.490, w: 0, h: 1.357, flipV: true, line: { color: C.white, width: 0.75, endArrowType: 'triangle' } });
  line(s, 0.605, 5.277, 0, 1.862, C.white, 0.75);
  line(s, 0.602, 7.139, 3.733, 0, C.white, 0.75);
  s.addShape('line', { x: 5.587, y: 7.139, w: 3.590, h: 0, line: { color: C.white, width: 0.75, endArrowType: 'triangle' } });
  txt(s, 'MARKET GROWTH', { x: -0.469, y: 4.427, w: 2.147, h: 0.269, fontSize: 10, align: 'center', rotate: -90, margin: [7.2, 7.2, 3.6, 3.6] });
  txt(s, 'MARKET SHARE', { x: 3.888, y: 7.004, w: 2.147, h: 0.269, fontSize: 10, align: 'center', margin: [7.2, 7.2, 3.6, 3.6] });

  var bubbles = [
    { x: 1.667, y: 3.341, halo: C.blue, core: C.blue, n: '02', t: C.white, ha: 41.18 },
    { x: 4.729, y: 3.051, halo: C.red, core: C.redSoft, n: '01', t: C.white, ha: 41.18 },
    { x: 2.362, y: 5.569, halo: C.white, core: C.white, n: '05', t: C.black, ha: 41.18 },
    { x: 3.412, y: 5.569, halo: C.grayDk, core: C.grayDk, n: '06', t: C.white, ha: 45.10 },
    { x: 5.869, y: 5.569, halo: C.grayMid, core: C.grayMid, n: '03', t: C.white, ha: 41.18 },
    { x: 7.421, y: 5.942, halo: C.grayLt, core: C.grayLt, n: '04', t: C.black, ha: 41.18 }
  ];
  bubbles.forEach(function (b) {
    ellipse(s, b.x, b.y, 0.802, 0.802, b.halo, b.ha);
    ellipse(s, b.x + 0.116, b.y + 0.117, 0.570, 0.570, b.core);
    txt(s, b.n, { x: b.x + 0.126, y: b.y + 0.291, w: 0.551, h: 0.252, fontSize: 9, color: b.t, align: 'center', margin: [7.2, 7.2, 3.6, 3.6] });
  });

  var legend = [
    { fill: C.red, adj: 17816, label: '1. OUR BUSINESS', color: C.white, ly: 2.660 },
    { fill: C.blue, adj: 15271, label: '2. COMPETITOR ONE', color: C.white, ly: 3.392 },
    { fill: C.grayMid, adj: 12726, label: '3. COMPETITOR TWO', color: C.black, ly: 4.153 },
    { fill: C.grayLt, adj: 15271, label: '4. COMPETITOR THREE', color: C.black, ly: 4.874 },
    { fill: C.white, adj: 12726, label: '5. COMPETITOR FOUR', color: C.black, ly: 5.635 },
    { fill: C.grayDk, adj: 12726, label: '6. COMPETITOR FIVE', color: C.white, ly: 6.311 }
  ];
  legend.forEach(function (l, i) {
    var y = 2.490 + i * 0.7318;
    roundRect(s, 9.478, y, 3.308, 0.677, l.fill, l.adj);
    var wide = l.label.indexOf('THREE') > -1;
    txt(s, l.label, {
      x: wide ? 10.016 : 10.170, y: l.ly, w: wide ? 2.233 : 1.924, h: 0.303,
      fontSize: 12, color: l.color, align: 'center', margin: [7.2, 7.2, 3.6, 3.6]
    });
  });
});

/* 19 - campaign recommendations --------------------------------------------------- */
build.push(function (s) {
  txt(s, 'Campaign', { x: 6.667, y: 1.638, w: 3.908, h: 0.808, fontSize: 48 });
  txt(s, 'Recommendations', { x: 6.667, y: 2.435, w: 6.033, h: 0.808, fontSize: 48 });
  syneCard(s, { x: 0.627, y: 1.555, h: 2.413, fill: C.blueBtn, title: 'Set Clear Goals' });
  syneCard(s, { x: 3.332, y: 1.555, h: 2.413, fill: C.red, title: 'Know Your Target Audience', titleW: 2.0 });

  roundRect(s, 0.627, 4.380, 5.156, 2.217, C.white, 5977);
  txt(s, 'Choose The Right Platform', { x: 1.036, y: 4.660, w: 1.628, h: 0.606, fontFace: F.syne, fontSize: 18, color: C.black });
  txt(s, 'Suspendisse interdum consectetur libero id faucis nisl tinu Arcu risus quvarius qua quiue inte consec interdum consectetur interdum consectetur libero.', {
    x: 1.036, y: 5.462, w: 4.336, h: 0.879, fontSize: 12, color: C.black, lineSpacingMultiple: 1.5
  });

  txt(s, 'Suspendisse interdum consectetur libero id faucibus nisl tincidu Arcu risus quvarius qua quisque id diam vel. Volutpat Suspendie interdum consectetur libero id faucibus nisl tincidu Arcu aucibus nisl libero id faucibus nisl.', {
    x: 6.665, y: 4.216, w: 5.770, h: 0.884, fontFace: F.reg, fontSize: 12, lineSpacingMultiple: 1.5
  });
  exploreBtn(s, 6.665, 6.140);
  ratingBlock(s, 9.568, 5.873);
});

/* 20 - investment corporate --------------------------------------------------------- */
build.push(function (s) {
  txt(s, 'Investment', { x: 0.627, y: 0.992, w: 5.757, h: 1.212, fontSize: 72 });
  txt(s, 'Corporate', { x: 0.627, y: 1.975, w: 5.594, h: 1.212, fontSize: 72 });
  txt(s, 'Our Client Location', { x: 0.682, y: 3.737, w: 3.961, h: 0.404, fontFace: F.reg, fontSize: 24, bold: true });

  blockArc(s, 0.765, 4.527, 1.899, -81.11, 15088885, 13623628, 23363, C.silver);
  blockArc(s, 0.765, 4.527, 1.899, 108.48, 1052751, 5200203, 22966, C.white);
  blockArc(s, 0.765, 4.527, 1.899, -81.11, 6958843, 12456348, 23495, C.blue);
  blockArc(s, 0.765, 4.527, 1.899, -81.11, 20039252, 7808909, 23886, C.red);

  var legend = [[C.red, '+667', C.white], [C.blue, '+890', C.white], [C.silver, '+450', C.white], [C.white, '+530', C.black]];
  legend.forEach(function (l, i) {
    var y = 4.629 + i * 0.5238;
    roundRect(s, 5.994, y, 1.102, 0.452, l[0], 50000);
    txt(s, 'Description Here', { x: 3.341, y: y + 0.109, w: 1.820, h: 0.236, fontFace: F.reg, fontSize: 14 });
    txt(s, l[1], { x: 6.232, y: y + 0.109, w: 0.627, h: 0.236, fontFace: F.reg, fontSize: 14, bold: true, color: l[2], align: 'center' });
  });

  [[7.874, 3.305, C.red, '34,689.90', 'Investment 01', 3.070, 2.035], [7.890, 5.282, C.blue, '20,697.35', 'Investment 02', 3.070, 2.019]].forEach(function (b) {
    roundRect(s, b[0], b[1], 4.777, 1.777, b[2], 11149);
    txt(s, b[3], { x: b[0] + 0.948, y: b[1] + 0.346, w: b[5], h: 0.740, fontFace: F.reg, fontSize: 44, bold: true });
    txt(s, b[4], { x: b[0] + 0.948, y: b[1] + 1.094, w: b[6], h: 0.337, fontFace: F.reg, fontSize: 20 });
  });
});

/* 21 - budget allocation ------------------------------------------------------------- */
build.push(function (s) {
  txt(s, 'About Budget', { x: 0.627, y: 1.055, w: 6.852, h: 0.841, fontSize: 50 });
  txt(s, 'Allocation', { x: 0.627, y: 1.827, w: 5.594, h: 0.841, fontSize: 50 });

  // header band: two capsule ends rotated into place plus a plain middle cell
  s.addShape('round2SameRect', { x: 5.380, y: 1.426, w: 0.777, h: 2.803, rotate: -90, fill: { color: C.blue } });
  rect(s, 7.170, 2.438, 2.803, 0.777, C.white);
  s.addShape('round2SameRect', { x: 10.986, y: 1.426, w: 0.777, h: 2.803, rotate: 90, fill: { color: C.red } });
  txt(s, 'Qty', { x: 4.879, y: 2.675, w: 1.778, h: 0.337, fontSize: 20, align: 'center' });
  txt(s, 'Price', { x: 7.630, y: 2.675, w: 1.778, h: 0.337, fontSize: 20, color: C.black, align: 'center' });
  txt(s, 'Amount', { x: 10.495, y: 2.680, w: 1.778, h: 0.337, fontSize: 20, align: 'center' });

  var rows = [
    { label: 'Content Creator', y: 3.617, lw: 2.277, qx: 4.963, qw: 1.778, vw: 1.778, price: '$1.200.00', amount: '$350.000' },
    { label: 'Research & Planning', y: 4.420, lw: 2.549, qx: 4.577, qw: 2.549, vw: 2.549, price: '$ 1.200.00', amount: '$ 350.000' },
    { label: 'Influencer Partnerships', y: 5.337, lw: 2.912, qx: 4.713, qw: 2.277, vw: 2.277, price: '$ 1.200.00', amount: '$ 350.000' },
    { label: 'Advertising', y: 6.140, lw: 2.549, qx: 4.577, qw: 2.549, vw: 2.549, price: '$ 1.200.00', amount: '$ 350.000' }
  ];
  rows.forEach(function (r) {
    txt(s, r.label, { x: 0.675, y: r.y, w: r.lw, h: 0.303, fontSize: 18 });
    txt(s, '70', { x: r.qx, y: r.y + 0.005, w: r.qw, h: 0.269, fontSize: 16, align: 'center' });
    txt(s, r.price, { x: 7.800, y: r.y, w: r.vw, h: 0.269, fontSize: 16 });
    txt(s, r.amount, { x: 10.582, y: r.y, w: r.vw, h: 0.269, fontSize: 16 });
  });
  [4.143, 5.004, 5.865, 6.726].forEach(function (y) { line(s, 0.675, y, 12.162, 0, C.rule, 0.75); });

  txt(s, 'TOTAL', { x: 4.622, y: 7.055, w: 2.549, h: 0.337, fontSize: 20, color: C.red, align: 'center' });
  txt(s, '$5.350.000.00', { x: 10.578, y: 7.055, w: 2.549, h: 0.337, fontSize: 20 });
});

/* 22 - social media tools -------------------------------------------------------------- */
build.push(function (s) {
  txt(s, 'Social Media', { x: 0.627, y: 1.295, w: 6.852, h: 1.111, fontSize: 66 });
  txt(s, 'Tools', { x: 0.647, y: 2.234, w: 3.504, h: 1.111, fontSize: 66 });
  txt(s, 'Suspendisse interdum consectetur libero id faucibus nisl tincidu Arcu risus quvas qua quisque id diam vel. Volutpat Suspendie interdum consectetur.', {
    x: 0.627, y: 3.503, w: 5.634, h: 0.575, fontFace: F.reg, fontSize: 12, lineSpacingMultiple: 1.5
  });
  syneCard(s, { x: 7.551, y: 1.555, h: 2.413, fill: C.gray, color: C.black, title: 'Management Platform' });
  syneCard(s, { x: 10.256, y: 1.555, h: 2.413, fill: C.white, color: C.black, title: 'Analytical Reporting', titleW: 2.0 });
  syneCard(s, { x: 0.625, y: 4.512, h: 2.413, fill: C.blueBtn, title: 'Management Platform' });
  syneCard(s, { x: 3.500, y: 4.512, h: 2.413, fill: C.red, title: 'Analytical Reporting', titleW: 2.0 });
  txt(s, 'Suspendisse interdum consectetur libero id faucibus nisl tincidu Arcu risus quvarius qua quisque id diam vel. Volutpat Suspendie interdum consectetur libero id faucibus nisl tincidu Arcu aucibus nisl libero id faucibus nisl.', {
    x: 7.551, y: 4.546, w: 5.283, h: 0.884, fontFace: F.reg, fontSize: 12, lineSpacingMultiple: 1.5
  });
  exploreBtn(s, 7.551, 6.469);
  ratingBlock(s, 10.230, 6.202);
});

/* 23 - contact info ---------------------------------------------------------------------- */
build.push(function (s) {
  txt(s, 'Contact', { x: 5.050, y: 1.192, w: 4.892, h: 1.481, fontSize: 88 });
  txt(s, 'Info', { x: 5.050, y: 2.329, w: 3.504, h: 1.481, fontSize: 88 });

  txt(s, 'Location                  :', { x: 0.627, y: 2.515, w: 2.493, h: 0.337, fontSize: 20, bold: true, valign: 'middle' });
  txt(s, [{ text: '28 Alma Vale Rd, Clifton, Bristol', options: { breakLine: true } }, { text: ' BS8 2HY, United Kingdom' }],
    { x: 0.627, y: 3.042, w: 2.693, h: 0.572, fontFace: F.light, fontSize: 12, color: C.white, valign: 'middle', margin: 0, lineSpacingMultiple: 1.5 });

  txt(s, 'Telephone              :', { x: 5.050, y: 5.621, w: 2.437, h: 0.337, fontSize: 20, bold: true, valign: 'middle' });
  txt(s, '+4411-7973-4300', { x: 5.050, y: 6.157, w: 1.840, h: 0.202, fontSize: 12, valign: 'middle' });
  txt(s, '+4411-8733-4310', { x: 5.050, y: 6.509, w: 1.429, h: 0.202, fontSize: 12, valign: 'middle' });

  txt(s, 'Email Address       :', { x: 9.414, y: 5.647, w: 2.488, h: 0.337, fontSize: 20, bold: true, valign: 'middle' });
  txt(s, 'hakirainfo@gmail.com', { x: 9.414, y: 6.209, w: 2.181, h: 0.202, fontSize: 12, valign: 'middle' });
  txt(s, 'hakira@gmail.com', { x: 9.414, y: 6.509, w: 2.181, h: 0.202, fontSize: 12, valign: 'middle' });

  exploreBtn(s, 5.050, 4.277);
  statCard(s, 10.505, 1.981);
});

/* 24 - thank you ------------------------------------------------------------------------- */
build.push(function (s) {
  txt(s, 'Thank You', { x: 6.357, y: 1.444, w: 5.511, h: 1.212, fontSize: 72 });
  txt(s, 'For Attention', { x: 6.357, y: 2.541, w: 6.784, h: 1.212, fontSize: 72 });
  txt(s, 'Suspendisse interdum consectetur libero id faucibus nisl tincidu Arcu risus quvarius qua quisque id diam vel. Volutpat Suspendie interdum consectetur libero id faucibus nisl tincidu Arcu aucibus nisl libero id faucibus nisl.', {
    x: 6.357, y: 4.383, w: 6.143, h: 0.884, fontFace: F.reg, fontSize: 12, lineSpacingMultiple: 1.5
  });
  exploreBtn(s, 6.357, 6.443);
  ratingBlock(s, 9.260, 6.176);

  var stats = [['9213+', 'Best Company', 0.821, 0.774, 1.191],
               ['98%', 'Value Company', 2.526, 2.437, 1.303],
               ['8650+', 'Best Client ', 4.231, 4.184, 1.191]];
  stats.forEach(function (st) {
    txt(s, st[0], { x: st[2], y: 6.131, w: 1.097, h: 0.438, fontFace: F.syne, fontSize: 20, align: 'center', margin: [7.2, 7.2, 3.6, 3.6] });
    txt(s, st[1], { x: st[3], y: 6.780, w: st[4], h: 0.185, fontFace: F.syne, fontSize: 11, align: 'center' });
  });
  line(s, 2.222, 6.049, 0, 0.999, C.ruleGray, 0.75);
  line(s, 3.927, 6.049, 0, 0.999, C.ruleGray, 0.75);
});

/* ------------------------------------------------------------------------ main */

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'HAKIRA', width: 13.333, height: 7.5 });
pptx.layout = 'HAKIRA';
pptx.author = 'Hakira';
pptx.title = 'Social Media Business';

build.forEach(function (fn) {
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  fn(slide);
  navBar(slide);
});

pptx.writeFile({ fileName: path.join(__dirname, '0edaa218-bb06-487d-bc65-653e37478722_grok_final.pptx') })
  .then(function (f) { console.log('wrote', f); });
