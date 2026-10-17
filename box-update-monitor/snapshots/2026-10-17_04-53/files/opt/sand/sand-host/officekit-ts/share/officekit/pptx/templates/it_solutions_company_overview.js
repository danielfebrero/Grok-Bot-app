// Recreation of "IT Solutions & Technology" deck (40 slides, 13.333 x 7.5 in)
// Every slide is built from plain pptxgenjs calls; images are replaced by flat placeholders.
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---- palette ---------------------------------------------------------------
const C = {
  PURPLE: '411867',   // accent3  - brand purple
  ORANGE: 'FF8B00',   // accent2  - brand orange
  BLUE:   '3455D1',   // accent1
  INK:    '200C33',   // dark pill background
  DEEP:   '270E3D',
  WHITE:  'FFFFFF',
  GREY:   '737373',   // body copy on light backgrounds
  LIGHT:  'D8D8D8',   // body copy on purple backgrounds
  MIST:   'F2F2F2',
  PANEL:  'F7F7F7',   // washed panel fill (stands in for the page gradient)
  SILVER: 'BFBFBF',
  HAIRLINE: 'E7E6E6',
  AMBER:  'E57D00',
  BLACK:  '000000'
};

// ---- typography ------------------------------------------------------------
const F = {
  REG:    'Poppins',
  SEMI:   'Poppins SemiBold',
  LIGHT:  'Poppins Light',
  MED:    'Poppins Medium',
  MINCHO: 'Sawarabi Mincho'
};

// ---- shadows ---------------------------------------------------------------
const SH = {
  drop: { type: 'outer', color: C.BLACK, opacity: 0.24, blur: 56, offset: 30, angle: 90 },
  deep: { type: 'outer', color: C.BLACK, opacity: 0.32, blur: 65, offset: 27, angle: 90 },
  card: { type: 'outer', color: C.BLACK, opacity: 0.20, blur: 23, offset: 8,  angle: 45 },
  soft: { type: 'outer', color: C.BLACK, opacity: 0.20, blur: 9,  offset: 3,  angle: 45 },
  hint: { type: 'outer', color: C.BLACK, opacity: 0.10, blur: 19, offset: 3,  angle: 90 },
  glow: { type: 'outer', color: C.ORANGE, opacity: 0.22, blur: 12, offset: 3, angle: 90 },
  ring: { type: 'outer', color: '1F337D', opacity: 0.40, blur: 5,  offset: 0,  angle: 0 }
};

// Filler copy used throughout the original deck; `lorem(n)` returns the first n words.
const LOREM = ('Re paragone creatura acerbita ai guardava lasciami vi. Entro tue forza miele mazzo per pur oltre ' +
  'sul. Lo ti il gabbie quanto lancio. Fato mare arme tu anch vi mine riso. Poi affannata ami cresciuto ' +
  'melagrani una abbandona brillanti. Far aspettando nel voluttuosa sei turbamento tra. Grappoli tuo inquieta ' +
  'cio orribile dissolve scoperto. Alzeremo voi parlando pei qualcuno serbatoi mio bellezza. Impregnato voi ' +
  'san esaltavano dal dal sfaldavano. ').split(' ');
function lorem(words, trailingSpace) {
  return LOREM.slice(0, words).join(' ') + (trailingSpace ? ' ' : '');
}

// ---- low level drawing helpers --------------------------------------------
function fillOf(o) {
  if (!o || !o.fill) return { type: 'none' };
  return o.alpha === undefined ? { color: o.fill } : { color: o.fill, transparency: 100 - o.alpha };
}

function frame(o) {
  const p = {};
  if (!o) return p;
  if (o.line) p.line = { color: o.line, width: o.lw === undefined ? 1 : o.lw };
  if (o.rot) p.rotate = o.rot;
  if (o.flipH) p.flipH = true;
  if (o.flipV) p.flipV = true;
  if (o.r !== undefined) p.rectRadius = o.r;
  if (o.ang) p.angleRange = o.ang;
  if (o.shadow) p.shadow = o.shadow;
  return p;
}

// Generic auto-shape.
function sh(s, kind, x, y, w, h, o) {
  s.addShape(kind, Object.assign({ x: x, y: y, w: w, h: h, fill: fillOf(o) }, frame(o)));
}

// Partial circle (progress ring segment).
function pie(s, x, y, w, h, o) {
  s.addShape('pie', { x: x, y: y, w: w, h: h, fill: fillOf(o), angleRange: o.ang });
}

// Horizontal rule / progress bar.
function hline(s, x, y, w, o) {
  s.addShape('line', { x: x, y: y, w: w, h: 0, line: { color: o.line, width: o.lw } });
}

// Single-run text block.
function tx(s, text, x, y, w, h, o) {
  o = o || {};
  s.addText(text, Object.assign({
    x: x, y: y, w: w, h: h,
    fontFace: o.f || F.REG,
    fontSize: o.sz || 12,
    bold: !!o.b,
    color: o.c || C.GREY,
    align: o.al || 'left',
    valign: o.va || 'top',
    margin: 0,
    lineSpacingMultiple: o.ls,
    shape: o.sh || 'rect',
    fill: o.fill ? fillOf(o) : { type: 'none' },
    bullet: o.bul ? { code: '2022', indent: 13.5 } : false
  }, frame(o)));
}

// Multi-run / multi-paragraph text block: runs are [text, overrides] pairs,
// a run of ['\n'] starts a new paragraph.
function rt(s, runs, x, y, w, h, o) {
  o = o || {};
  const body = [];
  runs.forEach(function (r) {
    if (r[0] === '\n') { if (body.length) body[body.length - 1].options.breakLine = true; return; }
    const ov = r[1] || {};
    body.push({
      text: r[0],
      options: {
        fontFace: ov.f || o.f || F.REG,
        fontSize: ov.sz || o.sz || 12,
        bold: ov.b === undefined ? !!o.b : ov.b === 1,
        color: ov.c || o.c || C.GREY,
        breakLine: false
      }
    });
  });
  s.addText(body, Object.assign({
    x: x, y: y, w: w, h: h,
    align: o.al || 'left',
    valign: o.va || 'top',
    margin: 0,
    lineSpacingMultiple: o.ls,
    shape: o.sh || 'rect',
    fill: o.fill ? fillOf(o) : { type: 'none' }
  }, frame(o)));
}

// ---- brand furniture -------------------------------------------------------
// Dark pill lockup: "IT Solutions & Technology" plus the 2x2 tile mark.
function pill(s, x, y) {
  s.addText([{ text: 'IT Solutions', options: { bold: true } }, { text: ' & Technology' }], {
    x: x, y: y, w: 2.844, h: 0.524, shape: 'roundRect', rectRadius: 0.262,
    fill: { color: C.INK }, color: C.WHITE, fontFace: F.REG, fontSize: 12,
    valign: 'middle', margin: 0, inset: 0
  });
  tiles(s, x + 2.387, y + 0.115, 0.134, [C.PURPLE, C.ORANGE, C.WHITE, C.BLUE]);
}

// 2x2 mark of diagonally-rounded tiles (top-right, top-left, bottom-right, bottom-left).
function tiles(s, x, y, d, cols) {
  const g = d * 1.265;
  [[x + g, y, cols[0], {}], [x, y, cols[1], { flipH: 1 }],
   [x + g, y + g * 1.17, cols[2], { flipV: 1 }], [x, y + g * 1.17, cols[3], { flipH: 1, flipV: 1 }]
  ].forEach(function (t) {
    sh(s, 'round2DiagRect', t[0], t[1], d, d, Object.assign({ fill: t[2], r: d * 0.5 }, t[3]));
  });
}

// Header wordmark: small tile mark + "IT Solutions & Technology ®2021".
function wordmark(s, x, y, color, mark, markTop) {
  if (mark !== null) tiles(s, x + 0.057, y + 0.086, 0.114, [markTop || C.DEEP, C.ORANGE, C.WHITE, C.BLUE]);
  rt(s, [['IT Solutions & Technology ', { b: 1 }], ['®2021']],
     x, y, 4.068, 0.421, { sz: 14, c: color, al: 'center', ls: 1.5 });
}

// Small eyebrow line above every section title.
function eyebrow(s, x, y, o) {
  rt(s, [['IT Solutions & Technology ', { b: 1 }], ['®2021']],
     x, y, 3.623, 0.375, Object.assign({ sz: 12, c: C.ORANGE, ls: 1.5 }, o || {}));
}

// Orange check badge (halo ring + disc + tick).
function check(s, x, y, d) {
  sh(s, 'ellipse', x, y, d, d, { fill: C.ORANGE, alpha: 40 });
  sh(s, 'ellipse', x + d * 0.11, y + d * 0.11, d * 0.78, d * 0.78, { fill: C.ORANGE });
  tickMark(s, x + d * 0.245, y + d * 0.31, d * 0.5, d * 0.34, C.WHITE);
}

// Standalone orange tick glyph used in bullet lists.
function tick(s, x, y) {
  sh(s, 'ellipse', x, y, 0.26, 0.26, { fill: C.ORANGE });
  tickMark(s, x + 0.062, y + 0.09, 0.14, 0.09, C.WHITE);
}

function tickMark(s, x, y, w, h, color) {
  s.addShape('custGeom', {
    x: x, y: y, w: w, h: h, fill: { color: color },
    points: [{ x: 0, y: h * 0.55 }, { x: w * 0.13, y: h * 0.38 }, { x: w * 0.37, y: h * 0.62 },
             { x: w * 0.87, y: 0 }, { x: w, y: h * 0.19 }, { x: w * 0.38, y: h }, { close: true }]
  });
}

// Circular ">" button (halo disc + inner disc + chevron).
function arrowDot(s, x, y, d, inner) {
  sh(s, 'ellipse', x, y, d, d, { fill: C.ORANGE });
  s.addText('>', {
    x: x + d * 0.096, y: y + d * 0.096, w: d * 0.81, h: d * 0.81, shape: 'ellipse',
    fill: { color: inner }, color: C.WHITE, fontFace: F.MINCHO, fontSize: d * 40,
    bold: true, align: 'center', valign: 'middle', margin: 0
  });
}

// ---- background furniture --------------------------------------------------
// The deck's page background is a very subtle white -> F2F2F2 vertical wash.
function panel(s, x, y, w, h, o) {
  s.addShape('rect', Object.assign({ x: x, y: y, w: w, h: h, fill: { color: C.PANEL } }, frame(o)));
}

// Half-round "dome" block that hangs off a page edge (used as a section mask).
function dome(s, x, y, w, h, color, flip) {
  s.addShape('custGeom', {
    x: x, y: y, w: w, h: h, fill: { color: color }, flipH: !!flip,
    points: [
      { x: w * 0.666, y: 0 }, { x: 0, y: 0 }, { x: 0, y: h }, { x: w * 0.666, y: h },
      { x: w * 0.674, y: h * 0.995 },
      { x: w, y: h * 0.5, curve: { type: 'cubic', x1: w * 0.872, y1: h * 0.883, x2: w, y2: h * 0.703 } },
      { x: w * 0.674, y: h * 0.005, curve: { type: 'cubic', x1: w, y1: h * 0.297, x2: w * 0.872, y2: h * 0.117 } },
      { close: true }
    ]
  });
}

// Browser chrome used on the website mock-up slide.
function browserFrame(s, x, y, w, h) {
  sh(s, 'round2SameRect', x, y, w, 0.4, { fill: C.SILVER, r: 0.06 });
  [[C.WHITE, 'FF0000'], [C.WHITE, 'FFC000'], [C.WHITE, '92D050']].forEach(function (c, i) {
    sh(s, 'ellipse', x + 0.219 + i * 0.141, y + 0.163, 0.089, 0.089, { fill: c[1] });
  });
  sh(s, 'roundRect', x + 0.851, y + 0.108, 2.898, 0.2, { fill: C.WHITE, r: 0.033 });
  sh(s, 'rect', x, y + 0.4, w, h - 0.4, { fill: C.WHITE });
}

// Raster art in the original deck is replaced by a flat block of its average colour.
function photo(s, x, y, w, h, o) {
  o = o || {};
  s.addShape(o.shape || 'rect', Object.assign({
    x: x, y: y, w: w, h: h, fill: { color: o.c || C.MIST }
  }, frame(o)));
}

// ---- bespoke artwork -------------------------------------------------------
// Big orange quotation glyph on the testimonial slide.
function quoteIcon(s, x, y, w, h) {
  sh(s, 'rect', x, y, w, h, { line: C.ORANGE, lw: 3 });
  s.addText('\u201C', {
    x: x, y: y + h * 0.06, w: w, h: h, color: C.ORANGE, fontFace: F.REG, fontSize: 44,
    bold: true, align: 'center', valign: 'middle', margin: 0
  });
}

// Ring built from independent arc segments: segs = [[startDeg, endDeg, color], ...].
function segRing(s, cx, cy, rOuter, thickness, segs) {
  segs.forEach(function (g) {
    s.addShape('blockArc', {
      x: cx - rOuter, y: cy - rOuter, w: rOuter * 2, h: rOuter * 2,
      fill: { color: g[2] },
      angleRange: [(g[0] + 360) % 360, (g[1] + 360) % 360],
      arcThicknessRatio: thickness / rOuter
    });
  });
}

// Thin dashed progress ring used on the statistics cards.
function gaugeRing(s, cx, cy, r, colors) {
  const step = 360 / colors.length;
  colors.forEach(function (c, i) {
    s.addShape('arc', {
      x: cx - r, y: cy - r, w: r * 2, h: r * 2,
      angleRange: [(i * step + 130) % 360, (i * step + step - 6 + 130) % 360],
      line: { color: c, width: 1.5 }
    });
  });
}

// Small orange square with a white plus sign (card corner action).
function plusBadge(s, x, y, d) {
  sh(s, 'rect', x, y, d, d, { fill: C.ORANGE });
  sh(s, 'plus', x + d * 0.29, y + d * 0.29, d * 0.42, d * 0.42, { fill: C.WHITE });
}

// Four-blade pinwheel diagram.
function pinwheel(s, cx, cy, r, colors) {
  colors.forEach(function (c, i) {
    s.addShape('pie', {
      x: cx - r, y: cy - r, w: r * 2, h: r * 2,
      fill: { color: c }, angleRange: [i * 90 + 45, i * 90 + 135]
    });
  });
}

// ---- contact icons ---------------------------------------------------------
function planeIcon(s, x, y, w, h) {
  s.addShape('custGeom', {
    x: x, y: y, w: w, h: h, fill: { color: C.WHITE },
    points: [{ x: 0, y: h * 0.45 }, { x: w, y: 0 }, { x: w * 0.62, y: h },
             { x: w * 0.44, y: h * 0.66 }, { close: true }]
  });
}

function phoneIcon(s, x, y, w, h) {
  s.addShape('custGeom', {
    x: x, y: y, w: w, h: h, fill: { color: C.WHITE },
    points: [{ x: 0, y: h * 0.12 }, { x: w * 0.28, y: 0 }, { x: w * 0.46, y: h * 0.3 },
             { x: w * 0.3, y: h * 0.46 }, { x: w * 0.56, y: h * 0.72 }, { x: w * 0.72, y: h * 0.56 },
             { x: w, y: h * 0.74 }, { x: w * 0.86, y: h }, { x: w * 0.5, y: h * 0.94 },
             { x: w * 0.12, y: h * 0.56 }, { close: true }]
  });
}

function birdIcon(s, x, y, w, h) {
  s.addShape('custGeom', {
    x: x, y: y, w: w, h: h, fill: { color: C.ORANGE },
    points: [{ x: w, y: h * 0.12 }, { x: w * 0.86, y: h * 0.2 }, { x: w * 0.94, y: h * 0.05 },
             { x: w * 0.76, y: h * 0.1 }, { x: w * 0.5, y: h * 0.3 }, { x: w * 0.1, y: h * 0.06 },
             { x: w * 0.16, y: h * 0.42 }, { x: w * 0.05, y: h * 0.36 }, { x: w * 0.24, y: h * 0.66 },
             { x: w * 0.12, y: h * 0.64 }, { x: w * 0.32, y: h * 0.84 }, { x: w * 0.02, y: h * 0.92 },
             { x: w * 0.36, y: h }, { x: w * 0.9, y: h * 0.42 }, { close: true }]
  });
}

function fbIcon(s, x, y, w, h) {
  s.addShape('roundRect', { x: x, y: y, w: w, h: h, fill: { color: C.ORANGE }, rectRadius: w * 0.12 });
  s.addText('f', {
    x: x, y: y, w: w, h: h, color: C.WHITE, fontFace: F.SEMI, fontSize: h * 62,
    bold: true, align: 'center', valign: 'middle', margin: 0
  });
}

function mailIcon(s, x, y, w, h) {
  sh(s, 'rect', x, y, w, h, { fill: C.ORANGE });
  s.addShape('custGeom', {
    x: x, y: y, w: w, h: h, fill: { color: C.WHITE },
    points: [{ x: w * 0.06, y: h * 0.14 }, { x: w * 0.5, y: h * 0.6 }, { x: w * 0.94, y: h * 0.14 },
             { x: w * 0.94, y: h * 0.06 }, { x: w * 0.5, y: h * 0.5 }, { x: w * 0.06, y: h * 0.06 },
             { close: true }]
  });
}

// ---- charts ----------------------------------------------------------------
function donutChart(s, x, y, w, h) {
  s.addChart('doughnut', [{
    name: 'Sales',
    labels: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'],
    values: [8.2, 3.2, 1.4, 1.2]
  }], {
    x: x, y: y, w: w, h: h,
    chartColors: [C.ORANGE, '55208C', C.BLUE, '7238B4'],
    dataBorder: { pt: 1.5, color: C.WHITE },
    holeSize: 75,
    showTitle: true, title: 'Sales', titleColor: '262626', titleFontFace: F.REG, titleFontSize: 18.6,
    showLegend: true, legendPos: 'b', legendColor: '262626', legendFontFace: F.REG, legendFontSize: 12,
    showValue: false, border: { pt: 0, color: C.WHITE }, fill: C.WHITE
  });
}


// Slide 1 - cover
function slide01(s) {
  sh(s, 'rect', 0, 0, 13.333, 6.111, {fill:C.PURPLE});
  panel(s, 0, 6.691, 13.333, 0.809);
  sh(s, 'rect', 0, 6.111, 13.333, 0.348, {fill:C.ORANGE});
  [[11.206, 0, 2.127, 1.923], [0, 4.188, 2.127, 1.923],
   [2.127, 3.273, 1.056, 0.954], [10.151, 1.923, 1.056, 0.954]].forEach(function (b) {
    sh(s, 'rect', b[0], b[1], b[2], b[3], {fill:C.WHITE, alpha:10});
  });
  pill(s, 1.147, 2.336);
  wordmark(s, 0.266, 0.233, C.WHITE);
  tx(s, 'New IT Solutions & Technology Intelligence', 0.469, 6.567, 2.557, 0.678, {c:C.PURPLE, ls:1.5});
  rt(s, [['Sahala'], [' IT.', {c:C.ORANGE}], ['\n'], ['IT Solutions & Technology', {sz:28}]], 1.097, 2.909, 5.533, 1.683, {sz:66, b:1, c:C.WHITE});
}

// Slide 2 - welcome
function slide02(s) {
  panel(s, 0, 0, 13.333, 7.5);
  sh(s, 'rect', 0.305, 0.3, 12.724, 5.543, {fill:C.PURPLE});
  sh(s, 'rect', 0.305, 5.843, 12.724, 0.316, {fill:C.ORANGE});
  arrowDot(s, 12.369, 0.519, 0.488, C.BLUE);
  tx(s, lorem(68, 1), 5.665, 2.02, 6.06, 2.193, {c:C.LIGHT, al:'justify', ls:1.5});
  tx(s, 'Welcome, to Sahala', 5.665, 1.346, 5.128, 0.572, {sz:32, b:1, c:C.WHITE});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 5.645, 0.984, 3.623, 0.375, {b:1, c:C.ORANGE, ls:1.5});
  tx(s, lorem(35), 5.665, 4.252, 6.06, 0.981, {c:C.LIGHT, al:'justify', ls:1.5});
  tx(s, 'New IT Solutions & Technology Intelligence', 5.672, 6.553, 2.557, 0.678, {c:C.PURPLE, ls:1.5});
}

// Slide 3 - quote
function slide03(s) {
  sh(s, 'rect', 0, 0, 13.333, 7.5, {fill:C.PURPLE});
  dome(s, 7.271, 0, 6.062, 7.5, C.ORANGE, true);
  tx(s, 'The advance of technology is based on making it fit in so that you don\'t really even notice it, so it\'s part of everyday life.', 0.302, 2.564, 6.81, 2.457, {sz:28, b:1, c:C.WHITE, al:'center'});
  tx(s, 'Florene Messina', 1.919, 5.021, 3.623, 0.421, {sz:14, c:C.WHITE, f:F.SEMI, al:'center', ls:1.5});
  quoteIcon(s, 0.808, 1.533, 0.798, 0.795);
  pill(s, 9.851, 0.455);
}

// Slide 4 - about (purple)
function slide04(s) {
  sh(s, 'rect', 0, 0, 13.333, 7.5, {fill:C.PURPLE});
  sh(s, 'rect', 0, 7.152, 13.333, 0.348, {fill:C.ORANGE});
  tx(s, 'Insert Text Here', 7.598, 4.956, 2.361, 0.303, {c:C.WHITE, f:F.SEMI});
  check(s, 7.308, 4.962, 0.29);
  tx(s, lorem(17), 7.621, 5.904, 5.029, 0.678, {c:C.LIGHT, ls:1.5, bul:1});
  tx(s, lorem(17), 7.598, 5.221, 5.029, 0.678, {c:C.LIGHT, ls:1.5, bul:1});
  tx(s, lorem(21), 7.328, 2.031, 5.299, 0.678, {c:C.LIGHT, al:'justify', ls:1.5});
  tx(s, lorem(51), 7.328, 2.793, 5.283, 1.89, {c:C.LIGHT, al:'justify', ls:1.5});
  tx(s, 'About Sahala', 7.308, 1.315, 3.835, 0.64, {sz:32, b:1, c:C.WHITE});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 7.288, 0.953, 3.623, 0.375, {b:1, c:C.ORANGE, ls:1.5});
}

// Slide 5 - about (light)
function slide05(s) {
  panel(s, 0, 0, 13.333, 7.5);
  sh(s, 'rect', 9.168, 0, 2.816, 7.5, {fill:C.PURPLE});
  tx(s, 'Join us now', 1.458, 5.438, 2.018, 0.426, {c:C.PURPLE, f:F.SEMI});
  arrowDot(s, 2.897, 5.478, 0.347, C.PURPLE);
  tx(s, lorem(21), 1.457, 2.443, 5.299, 0.678, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, lorem(51), 1.457, 3.205, 5.283, 1.89, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, 'About Sahala', 1.437, 1.727, 3.835, 0.64, {sz:32, b:1, c:C.PURPLE});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 1.417, 1.365, 3.623, 0.375, {b:1, c:C.ORANGE, ls:1.5});
  tx(s, 'New IT Solutions & Technology Intelligence', 0.4, 6.796, 4.673, 0.375, {c:C.PURPLE, ls:1.5});
  wordmark(s, 0.266, 0.169, C.PURPLE, C.PURPLE);
}

// Slide 6 - about + rounded photo
function slide06(s) {
  panel(s, 0, 0, 13.333, 7.5);
  sh(s, 'rect', 5.571, 7.152, 7.762, 0.348, {fill:C.ORANGE});
  tick(s, 6.3, 3.911);
  tx(s, lorem(21), 6.709, 3.831, 5.247, 0.678, {c:C.GREY, al:'justify', ls:1.5});
  tick(s, 6.3, 4.789);
  tx(s, lorem(21), 6.709, 4.709, 5.247, 0.678, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, lorem(32), 6.305, 2.661, 5.652, 0.981, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, 'About Sahala', 6.3, 1.956, 3.835, 0.64, {sz:32, b:1, c:C.PURPLE});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 6.28, 1.594, 3.623, 0.375, {b:1, c:C.ORANGE, ls:1.5});
  pill(s, 9.851, 6.353);
}

// Slide 7 - about + progress rings
function slide07(s) {
  panel(s, 0, 0, 13.333, 5.086);
  sh(s, 'round1Rect', 2.079, 1.314, 3.692, 6.186, {fill:C.PURPLE, r:1.63, flipH:1});
  sh(s, 'rect', 0.951, 2.349, 3.054, 3.137, {fill:C.WHITE, flipH:1, shadow:SH.drop});
  sh(s, 'ellipse', 1.813, 2.516, 1.331, 1.331, {fill:C.WHITE, alpha:25, line:C.WHITE});
  sh(s, 'ellipse', 1.919, 2.615, 1.119, 1.119, {fill:C.MIST});
  pie(s, 1.919, 2.615, 1.119, 1.119, {fill:C.INK, ang:[88.437, 269.872]});
  sh(s, 'ellipse', 1.983, 2.679, 0.991, 0.991, {fill:C.PURPLE});
  tx(s, '50%', 1.909, 3.018, 1.139, 0.37, {sz:16, b:1, c:C.ORANGE, al:'center'});
  tx(s, 'Insert Text Here', 1.298, 3.984, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, lorem(15), 0.951, 4.221, 3.054, 0.981, {c:C.GREY, al:'center', ls:1.5});
  sh(s, 'rect', 6.233, 3.936, 3.054, 3.137, {fill:C.WHITE, flipH:1, shadow:SH.drop});
  sh(s, 'ellipse', 7.095, 4.24, 1.331, 1.331, {fill:C.WHITE, alpha:25, line:C.WHITE});
  sh(s, 'ellipse', 7.201, 4.339, 1.119, 1.119, {fill:C.MIST});
  pie(s, 7.201, 4.339, 1.119, 1.119, {fill:C.INK, ang:[39.546, 269.872]});
  sh(s, 'ellipse', 7.265, 4.403, 0.991, 0.991, {fill:C.PURPLE});
  tx(s, '70%', 7.191, 4.742, 1.139, 0.37, {sz:16, b:1, c:C.ORANGE, al:'center'});
  tx(s, 'Insert Text Here', 6.58, 5.571, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, lorem(15), 6.233, 5.808, 3.054, 0.981, {c:C.GREY, al:'center', ls:1.5});
  tx(s, lorem(32), 7.904, 1.919, 4.27, 1.284, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, 'About Sahala', 7.899, 1.214, 3.835, 0.64, {sz:32, b:1, c:C.PURPLE});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 7.879, 0.852, 3.623, 0.375, {b:1, c:C.ORANGE, ls:1.5});
  pill(s, 9.932, 6.353);
}

// Slide 8 - new tech sahala
function slide08(s) {
  panel(s, 0, 0, 6.329, 7.5);
  sh(s, 'rect', 0, 0, 6.329, 0.348, {fill:C.ORANGE});
  tx(s, 'Read More', 0.51, 5.854, 1.729, 0.426, {sh:'roundRect', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, al:'center', r:0.213, shadow:SH.card});
  sh(s, 'rect', 6.329, 5.333, 3.077, 2.167, {fill:C.PURPLE, flipH:1, shadow:SH.drop});
  rt(s, [['New High '], ['Technology', {c:C.ORANGE}], [' With Best '], ['IT Solution', {c:C.ORANGE}], [' For This Year In 2021']], 6.643, 5.577, 2.449, 1.678, {sz:16, c:C.WHITE, f:F.SEMI, al:'center', ls:1.5});
  check(s, 9.091, 5.139, 0.614);
  tx(s, lorem(17, 1), 0.485, 2.925, 5.299, 0.678, {b:1, c:C.GREY, al:'justify', ls:1.5});
  tx(s, lorem(51), 0.485, 3.687, 5.283, 1.89, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, 'New Tech Sahala', 0.465, 2.209, 4.576, 0.64, {sz:32, b:1, c:C.PURPLE});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 0.445, 1.847, 3.623, 0.375, {b:1, c:C.ORANGE, ls:1.5});
}

// Slide 9 - new tech + progress bars
function slide09(s) {
  panel(s, 0, 0, 13.333, 7.5);
  sh(s, 'rect', 4.921, 1.413, 1.19, 4.675, {fill:C.PURPLE});
  tx(s, '01', 6.479, 3.662, 0.857, 0.788, {sz:36, b:1, c:C.PURPLE, ls:1.2});
  hline(s, 7.484, 6.381, 2.436, {line:C.SILVER, lw:6});
  hline(s, 7.484, 6.381, 2.23, {line:C.ORANGE, lw:6});
  tx(s, 'Insert Text Here', 7.336, 3.662, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(22), 7.336, 3.89, 5.365, 0.678, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, '02', 6.479, 5.131, 0.857, 0.788, {sz:36, b:1, c:C.PURPLE, ls:1.2});
  tx(s, 'Insert Text Here', 7.336, 5.131, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(22), 7.336, 5.359, 5.365, 0.678, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, '88.00%', 10.65, 6.182, 1.733, 0.505, {sz:24, b:1, c:C.PURPLE});
  hline(s, 7.484, 4.823, 2.436, {line:C.SILVER, lw:6});
  hline(s, 7.484, 4.823, 0.59, {line:C.ORANGE, lw:6});
  tx(s, '75%', 10.65, 4.624, 1.733, 0.505, {sz:24, b:1, c:C.PURPLE});
  rt(s, [['New High '], ['Technology', {c:C.ORANGE}], [' With Best '], ['IT Solution', {c:C.ORANGE}], [' For This Year In 2021']], 3.497, 3.315, 3.975, 0.871, {sz:16, c:C.WHITE, f:F.SEMI, al:'center', ls:1.5, rot:270});
  tx(s, lorem(34), 6.519, 2.474, 6.183, 0.981, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, 'New Tech Sahala', 6.499, 1.772, 4.576, 0.64, {sz:32, b:1, c:C.PURPLE});
  pill(s, 6.519, 1.165);
  tx(s, 'New IT Solutions & Technology Intelligence', 0.469, 6.567, 2.557, 0.678, {c:C.PURPLE, ls:1.5});
}

// Slide 10 - numbered cards
function slide10(s) {
  panel(s, 0, 0, 13.333, 7.5);
  dome(s, 0, 0, 4.968, 7.5, C.PURPLE);
  sh(s, 'rect', 4.314, 1.627, 6.394, 1.4, {fill:C.WHITE, flipH:1, shadow:SH.drop});
  tx(s, '01', 4.543, 1.982, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, va:'middle'});
  tx(s, 'Insert Text Here', 5.146, 1.834, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(22), 5.146, 2.061, 5.381, 0.678, {c:C.GREY, al:'justify', ls:1.5});
  sh(s, 'rect', 4.314, 3.255, 6.394, 1.4, {fill:C.WHITE, flipH:1, shadow:SH.drop});
  tx(s, '02', 4.543, 3.61, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, va:'middle'});
  tx(s, 'Insert Text Here', 5.146, 3.461, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(22), 5.146, 3.689, 5.381, 0.678, {c:C.GREY, al:'justify', ls:1.5});
  sh(s, 'rect', 4.314, 4.882, 6.394, 1.4, {fill:C.WHITE, flipH:1, shadow:SH.drop});
  tx(s, '03', 4.543, 5.238, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, va:'middle'});
  tx(s, 'Insert Text Here', 5.146, 5.089, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(22), 5.146, 5.317, 5.381, 0.678, {c:C.GREY, al:'justify', ls:1.5});
  pill(s, 10.014, 0.408);
  tx(s, 'New IT Solutions & Technology Intelligence', 10.301, 6.567, 2.557, 0.678, {c:C.PURPLE, al:'right', ls:1.5});
}

// Slide 11 - about card + read more
function slide11(s) {
  panel(s, 0, 0, 13.333, 7.5);
  sh(s, 'rect', 6.071, 0, 0.44, 4.4, {fill:C.PURPLE});
  sh(s, 'rect', 1.075, 2.11, 8.664, 4.58, {fill:C.WHITE, flipH:1, shadow:SH.drop});
  tx(s, '01', 1.388, 5.136, 0.857, 0.788, {sz:36, b:1, c:C.ORANGE, ls:1.2});
  tx(s, 'Insert Text Here', 2.245, 5.136, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(16), 2.245, 5.364, 3.104, 0.981, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, '02', 5.538, 5.136, 0.857, 0.788, {sz:36, b:1, c:C.ORANGE, ls:1.2});
  tx(s, 'Insert Text Here', 6.395, 5.136, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(16), 6.395, 5.364, 3.104, 0.981, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, lorem(51), 1.457, 3.45, 8.042, 1.284, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, 'About Sahala', 1.437, 2.79, 3.835, 0.64, {sz:32, b:1, c:C.PURPLE});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 1.417, 2.428, 3.623, 0.375, {b:1, c:C.ORANGE, ls:1.5});
  wordmark(s, 0.266, 0.169, C.PURPLE, null);
  tx(s, 'Read More', 8.972, 4.467, 1.729, 0.426, {sh:'roundRect', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, al:'center', r:0.213, rot:90, shadow:SH.card});
  tx(s, 'New IT Solutions & Technology Intelligence', 10.301, 6.567, 2.557, 0.678, {c:C.PURPLE, al:'right', ls:1.5});
}

// Slide 12 - checklist + name plate
function slide12(s) {
  panel(s, 0, 0, 13.333, 7.5);
  sh(s, 'rect', 0, 0, 1.683, 7.5, {fill:C.PURPLE});
  sh(s, 'roundRect', 2.723, 5.457, 5.333, 1.367, {fill:C.WHITE, r:0.683, flipH:1, shadow:SH.drop});
  sh(s, 'roundRect', 2.172, 5.275, 5.333, 0.987, {fill:C.WHITE, r:0.493, flipH:1, shadow:SH.drop});
  rt(s, [['New High '], ['Technology', {c:C.ORANGE}], [' With Best '], ['IT Solution', {c:C.ORANGE}]], 2.937, 6.313, 4.905, 0.375, {c:C.PURPLE, f:F.SEMI, al:'center', ls:1.5});
  tx(s, 'New Tech Sahala', 2.526, 5.422, 4.576, 0.64, {sz:32, b:1, c:C.PURPLE});
  pill(s, 4.838, 4.872);
  tick(s, 5.908, 0.776);
  tx(s, lorem(21), 6.318, 0.697, 5.247, 0.678, {c:C.GREY, al:'justify', ls:1.5});
  tick(s, 5.908, 1.672);
  tx(s, lorem(21), 6.318, 1.593, 5.247, 0.678, {c:C.GREY, al:'justify', ls:1.5});
}

// Slide 13 - expert staff
function slide13(s) {
  panel(s, 0, 0, 13.333, 7.5);
  sh(s, 'rect', 0, 0, 13.333, 0.348, {fill:C.ORANGE});
  sh(s, 'roundRect', 0.726, 5.531, 3.408, 0.731, {fill:C.WHITE, r:0.365, flipH:1, shadow:SH.drop});
  tx(s, 'Lissette Stull', 1.009, 5.573, 2.55, 0.37, {sz:16, c:C.PURPLE, f:F.SEMI});
  tx(s, 'Job position ', 1.009, 5.938, 1.868, 0.286, {sz:11, c:C.ORANGE, f:F.LIGHT});
  check(s, 3.674, 5.333, 0.515);
  sh(s, 'roundRect', 5.014, 5.531, 3.408, 0.731, {fill:C.WHITE, r:0.365, flipH:1, shadow:SH.drop});
  tx(s, 'Cathleen Crain', 5.297, 5.573, 2.55, 0.37, {sz:16, c:C.PURPLE, f:F.SEMI});
  tx(s, 'Job position ', 5.297, 5.938, 1.868, 0.286, {sz:11, c:C.ORANGE, f:F.LIGHT});
  check(s, 7.962, 5.333, 0.515);
  sh(s, 'roundRect', 8.964, 5.531, 3.408, 0.731, {fill:C.WHITE, r:0.365, flipH:1, shadow:SH.drop});
  tx(s, 'Hortense Gilson', 9.247, 5.573, 2.55, 0.37, {sz:16, c:C.PURPLE, f:F.SEMI});
  tx(s, 'Job position ', 9.247, 5.938, 1.868, 0.286, {sz:11, c:C.ORANGE, f:F.LIGHT});
  check(s, 11.912, 5.333, 0.515);
  tx(s, lorem(31, 1), 2.482, 1.65, 8.042, 0.678, {c:C.GREY, al:'center', ls:1.5});
  tx(s, 'Expert Staff Sahala', 3.354, 0.99, 6.299, 0.64, {sz:32, b:1, c:C.PURPLE, al:'center'});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 4.692, 0.628, 3.623, 0.375, {b:1, c:C.ORANGE, al:'center', ls:1.5});
  tx(s, 'New IT Solutions & Technology Intelligence', 8.651, 6.706, 4.207, 0.375, {c:C.PURPLE, al:'right', ls:1.5});
}

// Slide 14 - meet our manager
function slide14(s) {
  sh(s, 'rect', 0, 0, 13.333, 7.5, {fill:C.PURPLE});
  sh(s, 'ellipse', 0.77, 0.881, 5.73, 5.73, {fill:C.PANEL});
  sh(s, 'roundRect', 1.949, 5.494, 3.339, 1.078, {fill:C.WHITE, r:0.087, shadow:SH.drop});
  tx(s, 'Morton Mccormack', 2.066, 5.67, 3.106, 0.438, {sz:20, c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, 'Manager of Sahala Tech', 2.344, 6.062, 2.55, 0.286, {sz:11, f:F.MED, al:'center'});
  sh(s, 'parallelogram', 10.165, 3.41, 0.748, 0.122, {fill:C.PURPLE, line:C.ORANGE, lw:0.5});
  sh(s, 'parallelogram', 9.421, 3.41, 0.748, 0.122, {fill:C.PURPLE, line:C.ORANGE, lw:0.5});
  sh(s, 'parallelogram', 8.677, 3.41, 0.748, 0.122, {fill:C.PURPLE, line:C.ORANGE, lw:0.5});
  sh(s, 'parallelogram', 7.933, 3.41, 0.748, 0.122, {fill:C.ORANGE, line:C.ORANGE, lw:0.5});
  sh(s, 'parallelogram', 7.189, 3.41, 0.748, 0.122, {fill:C.ORANGE, line:C.ORANGE, lw:0.5});
  sh(s, 'parallelogram', 10.165, 4.26, 0.748, 0.122, {fill:C.PURPLE, line:C.ORANGE, lw:0.5});
  sh(s, 'parallelogram', 9.421, 4.26, 0.748, 0.122, {fill:C.PURPLE, line:C.ORANGE, lw:0.5});
  sh(s, 'parallelogram', 8.677, 4.26, 0.748, 0.122, {fill:C.ORANGE, line:C.ORANGE, lw:0.5});
  sh(s, 'parallelogram', 7.933, 4.26, 0.748, 0.122, {fill:C.ORANGE, line:C.ORANGE, lw:0.5});
  sh(s, 'parallelogram', 7.189, 4.26, 0.748, 0.122, {fill:C.ORANGE, line:C.ORANGE, lw:0.5});
  sh(s, 'parallelogram', 10.165, 3.833, 0.748, 0.122, {fill:C.PURPLE, line:C.ORANGE, lw:0.5});
  sh(s, 'parallelogram', 9.421, 3.833, 0.748, 0.122, {fill:C.ORANGE, line:C.ORANGE, lw:0.5});
  sh(s, 'parallelogram', 8.677, 3.833, 0.748, 0.122, {fill:C.ORANGE, line:C.ORANGE, lw:0.5});
  sh(s, 'parallelogram', 7.933, 3.833, 0.748, 0.122, {fill:C.ORANGE, line:C.ORANGE, lw:0.5});
  sh(s, 'parallelogram', 7.189, 3.833, 0.748, 0.122, {fill:C.ORANGE, line:C.ORANGE, lw:0.5});
  tx(s, lorem(21), 7.154, 2.347, 5.299, 0.678, {c:C.LIGHT, al:'justify', ls:1.5});
  tx(s, lorem(46, 1), 7.154, 4.621, 5.283, 1.587, {c:C.LIGHT, al:'justify', ls:1.5});
  tx(s, 'Meet Our Manager', 7.134, 1.631, 4.982, 0.64, {sz:32, b:1, c:C.WHITE});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 7.114, 1.269, 3.623, 0.375, {b:1, c:C.ORANGE, ls:1.5});
  wordmark(s, 0.266, 0.233, C.ORANGE);
}

// Slide 15 - staff of sahala
function slide15(s) {
  panel(s, 0, 0, 13.333, 7.5);
  [[0.829, 0.914], [0.829, 3.886], [4.199, 0.914], [4.199, 3.886]].forEach(function (b) {
    sh(s, 'roundRect', b[0], b[1], 2.8, 2.629, {fill:C.WHITE, r:0.252, shadow:SH.card});
    sh(s, 'rect', b[0], b[1], 0.12, 2.629, {fill:C.PURPLE});
  });
  tx(s, 'Dewayne Grey', 1.009, 5.849, 2.55, 0.337, {sz:14, c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, 'Job position ', 1.35, 6.157, 1.868, 0.286, {sz:10.5, c:C.ORANGE, f:F.LIGHT, al:'center'});
  tx(s, 'Evelin Messer', 4.399, 5.849, 2.55, 0.337, {sz:14, c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, 'Job position ', 4.74, 6.157, 1.868, 0.286, {sz:10.5, c:C.ORANGE, f:F.LIGHT, al:'center'});
  tx(s, 'Sandy Prescott', 1.009, 2.843, 2.55, 0.337, {sz:14, c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, 'Job position ', 1.35, 3.151, 1.868, 0.286, {sz:10.5, c:C.ORANGE, f:F.LIGHT, al:'center'});
  tx(s, 'Seymour Hutchens', 4.399, 2.843, 2.55, 0.337, {sz:14, c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, 'Job position ', 4.74, 3.151, 1.868, 0.286, {sz:10.5, c:C.ORANGE, f:F.LIGHT, al:'center'});
  tx(s, 'Staff of Sahala', 7.486, 2.402, 4.576, 0.64, {sz:32, b:1, c:C.PURPLE});
  pill(s, 7.506, 1.795);
  tx(s, lorem(51), 7.506, 3.122, 5.001, 1.89, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, lorem(20), 7.506, 5.093, 5.001, 0.678, {c:C.GREY, al:'justify', ls:1.5});
}

// Slide 16 - service of sahala
function slide16(s) {
  panel(s, 0, 0, 7.4, 7.5);
  sh(s, 'roundRect', 0.391, 5.076, 6.5, 1.371, {fill:C.WHITE, r:0.131, shadow:SH.card});
  sh(s, 'rect', 7.4, 0, 5.933, 7.5, {fill:C.PURPLE, alpha:60});
  sh(s, 'roundRect', 5.824, 0.985, 5.333, 0.987, {fill:C.WHITE, r:0.493, flipH:1, shadow:SH.drop});
  tx(s, 'Service of Sahala', 6.178, 1.132, 4.576, 0.64, {sz:32, b:1, c:C.PURPLE});
  pill(s, 8.49, 0.582);
  tx(s, 'Insert Text Here', 8.49, 2.637, 2.361, 0.303, {c:C.WHITE, f:F.SEMI});
  check(s, 8.2, 2.644, 0.29);
  tx(s, lorem(11), 8.49, 2.902, 3.431, 0.678, {c:C.LIGHT, al:'justify', ls:1.5, bul:1});
  tx(s, 'Insert Text Here', 8.49, 3.744, 2.361, 0.303, {c:C.WHITE, f:F.SEMI});
  check(s, 8.2, 3.751, 0.29);
  tx(s, lorem(11), 8.49, 4.009, 3.431, 0.678, {c:C.LIGHT, al:'justify', ls:1.5, bul:1});
  tx(s, 'Insert Text Here', 8.49, 4.881, 2.361, 0.303, {c:C.WHITE, f:F.SEMI});
  check(s, 8.2, 4.887, 0.29);
  tx(s, lorem(11), 8.49, 5.145, 3.431, 0.678, {c:C.LIGHT, al:'justify', ls:1.5, bul:1});
  tx(s, lorem(51), 0.765, 3.131, 5.695, 1.587, {c:C.GREY, al:'justify', ls:1.5});
  rt(s, [['We Cav Provide High '], ['Tech', {c:C.ORANGE}], [' & '], ['IT Solutions', {c:C.ORANGE}]], 0.745, 2.243, 4.936, 0.909, {sz:24, b:1, c:C.PURPLE});
  tx(s, '01', 0.765, 5.449, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, al:'center', va:'middle'});
  tx(s, 'Insert Text Here', 1.368, 5.3, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(22), 1.368, 5.528, 5.381, 0.678, {c:C.GREY, al:'justify', ls:1.5});
}

// Slide 17 - best way to coding
function slide17(s) {
  panel(s, 7.314, 0, 6.019, 7.5);
  sh(s, 'rect', 12.894, 3.557, 0.44, 3.943, {fill:C.PURPLE});
  sh(s, 'roundRect', 0.444, 2.708, 3.219, 2.085, {fill:C.WHITE, r:0, flipH:1, shadow:SH.drop});
  tx(s, '01.', 0.614, 2.843, 0.683, 0.505, {sz:24, b:1, c:C.PURPLE, al:'center'});
  tx(s, 'Insert Text Here', 0.614, 3.361, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(15), 0.614, 3.589, 2.877, 0.981, {c:C.GREY, al:'justify', ls:1.5});
  sh(s, 'roundRect', 4.081, 2.708, 3.219, 2.085, {fill:C.WHITE, r:0, flipH:1, shadow:SH.drop});
  tx(s, '02.', 4.252, 2.843, 0.683, 0.505, {sz:24, b:1, c:C.PURPLE, al:'center'});
  tx(s, 'Insert Text Here', 4.252, 3.361, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(15), 4.252, 3.589, 2.877, 0.981, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, lorem(51), 7.616, 4.792, 5.051, 1.89, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, 'The Best Way To Coding New IT Solutions', 0.464, 5.504, 6.149, 1.178, {sz:32, b:1, c:C.PURPLE});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 0.444, 5.142, 3.623, 0.375, {b:1, c:C.ORANGE, ls:1.5});
  tx(s, 'Read More', 7.719, 4.222, 1.729, 0.426, {sh:'roundRect', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, al:'center', r:0.213, shadow:SH.card});
}

// Slide 18 - we can provide high tech
function slide18(s) {
  panel(s, 0, 0, 6.971, 7.5);
  sh(s, 'rect', 5.801, 0.521, 3.348, 6.457, {fill:C.WHITE});
  tx(s, '01', 7.171, 1.125, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, al:'center', va:'middle'});
  tx(s, 'Insert Text Here', 6.284, 1.894, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, lorem(22), 6.016, 2.122, 2.897, 1.284, {c:C.GREY, al:'center', ls:1.5});
  tx(s, '02', 7.171, 3.96, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, al:'center', va:'middle'});
  tx(s, 'Insert Text Here', 6.284, 4.73, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, lorem(22), 6.016, 4.957, 2.897, 1.284, {c:C.GREY, al:'center', ls:1.5});
  tx(s, lorem(51), 0.591, 3.131, 4.641, 1.89, {c:C.GREY, al:'justify', ls:1.5});
  rt(s, [['We Cav Provide High '], ['Tech', {c:C.ORANGE}], [' & '], ['IT Solutions', {c:C.ORANGE}]], 0.571, 2.243, 4.936, 0.909, {sz:24, b:1, c:C.PURPLE});
  tx(s, lorem(27), 0.591, 5.096, 4.641, 0.981, {c:C.GREY, al:'justify', ls:1.5});
  pill(s, 0.58, 1.563);
}

// Slide 19 - best way to coding (card)
function slide19(s) {
  sh(s, 'rect', 0, 0, 0.957, 7.5, {fill:C.PURPLE});
  panel(s, 5.386, 0, 7.948, 7.5);
  sh(s, 'roundRect', 5.313, 2.708, 7.02, 3.974, {fill:C.WHITE, r:0, flipH:1, shadow:SH.drop});
  tx(s, 'The Best Way To Coding New IT Solutions', 5.58, 3.336, 6.149, 1.178, {sz:32, b:1, c:C.PURPLE});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 5.56, 2.974, 3.623, 0.375, {b:1, c:C.ORANGE, ls:1.5});
  tick(s, 5.727, 4.818);
  tx(s, 'Database Analysis', 6.059, 4.795, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tick(s, 5.727, 5.318);
  tx(s, 'Happy Customers', 6.059, 5.295, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tick(s, 5.727, 5.819);
  tx(s, 'IT Consultancy', 6.059, 5.796, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tick(s, 8.087, 4.818);
  tx(s, 'UI/UX Designs', 8.419, 4.795, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tick(s, 8.087, 5.318);
  tx(s, 'Web Development', 8.419, 5.295, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, 'Let’s Tour', 9.985, 5.948, 1.729, 0.426, {sh:'roundRect', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, al:'center', r:0.213, shadow:SH.card});
  tx(s, 'Re paragone creatura acerbita ai guardava lasciami vi. Entro tue forza miele mazzo per pur oltre sul. Lo ti il gabbie quanto lancio. Fato mare affannata ami cresciuto melagrani nel voluttuosa sei turbamento tra. Grappoli tuo inquieta cio orribile', 1.423, 4.818, 3.289, 2.193, {c:C.GREY, al:'justify', ls:1.5});
}

// Slide 20 - breaktime
function slide20(s) {
  sh(s, 'rect', 0, 0, 13.333, 7.5, {fill:C.PURPLE});
  sh(s, 'rect', 0, 0, 0.957, 7.5, {fill:C.ORANGE});
  rt(s, [['Breaktime'], ['\n'], ['45 ', {sz:66, b:1}], ['Minute', {sz:66, b:1}]], 0.999, 2.34, 5.888, 1.683, {sz:28, c:C.WHITE, al:'right'});
  rt(s, [['12.'], ['45', {b:0}]], 2.823, 3.75, 4.064, 2.036, {sz:115, b:1, c:C.WHITE, al:'right'});
  pill(s, 9.336, 6.266);
  wordmark(s, 1.426, 0.233, C.WHITE);
  tx(s, 'New IT Solutions & Technology Intelligence', 1.426, 6.567, 2.557, 0.678, {c:C.WHITE, ls:1.5});
}

// Slide 21 - study case it consultancy
function slide21(s) {
  panel(s, 6.4, 3.614, 6.933, 3.886);
  sh(s, 'rect', 0.821, 1.963, 3.077, 2.167, {fill:C.PURPLE, flipH:1, shadow:SH.drop});
  rt(s, [['New High '], ['Technology', {c:C.ORANGE}], [' With Best '], ['IT Solution', {c:C.ORANGE}], [' For This Year In 2021']], 1.136, 2.207, 2.449, 1.678, {sz:16, c:C.WHITE, f:F.SEMI, al:'center', ls:1.5});
  check(s, 2.053, 1.593, 0.614);
  tx(s, 'Study Case IT Consultancy', 6.739, 4.256, 3.623, 1.178, {sz:32, b:1, c:C.PURPLE});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 6.719, 3.894, 3.623, 0.375, {b:1, c:C.ORANGE, ls:1.5});
  tx(s, lorem(51), 6.739, 5.614, 6.072, 1.587, {c:C.GREY, al:'justify', ls:1.5});
}

// Slide 22 - high technology
function slide22(s) {
  panel(s, 0, 0, 13.333, 7.5);
  sh(s, 'rect', 0, 4.3, 6.829, 3.2, {fill:C.PURPLE});
  [0.693, 4.799].forEach(function (x) { sh(s, 'rect', x, 2.907, 3.736, 3.964, {fill:C.WHITE, shadow:SH.card}); });
  tx(s, 'High Technology', 0.907, 0.937, 4.385, 0.64, {sz:32, b:1, c:C.PURPLE});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 0.887, 0.575, 3.623, 0.375, {b:1, c:C.ORANGE, ls:1.5});
  tx(s, lorem(41), 0.907, 1.719, 7.649, 0.981, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, 'Insert Text Here', 9.099, 3.75, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(20), 9.099, 3.978, 3.348, 0.981, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, 'Insert Text Here', 9.099, 5.131, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(20), 9.099, 5.359, 3.348, 0.981, {c:C.GREY, al:'justify', ls:1.5});
  pill(s, 10.037, 0.312);
  hline(s, 1.358, 6.563, 2.436, {line:C.SILVER, lw:6});
  hline(s, 1.358, 6.563, 2.23, {line:C.ORANGE, lw:6});
  tx(s, 'Insert Text Here', 1.358, 6.142, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  hline(s, 5.486, 6.563, 2.436, {line:C.SILVER, lw:6});
  hline(s, 5.486, 6.563, 1.069, {line:C.ORANGE, lw:6});
  tx(s, 'Insert Text Here', 5.486, 6.142, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
}

// Slide 23 - it solutions
function slide23(s) {
  panel(s, 0, 0, 9.443, 7.5);
  tx(s, 'IT Solutions', 0.865, 1.921, 3.603, 0.64, {sz:32, b:1, c:C.PURPLE});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 0.845, 1.559, 3.623, 0.375, {b:1, c:C.ORANGE, ls:1.5});
  tx(s, lorem(23, 1), 0.865, 2.704, 4.204, 0.981, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, lorem(41), 0.865, 3.829, 4.204, 1.89, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, 'Let’s Go', 0.927, 5.941, 1.729, 0.426, {sh:'roundRect', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, al:'center', r:0.213, shadow:SH.card});
}

// Slide 24 - circles list
function slide24(s) {
  sh(s, 'rect', 0, 0, 5.657, 7.5, {fill:C.PURPLE});
  sh(s, 'rect', 0.95, 0.807, 3.736, 5.007, {fill:C.WHITE, shadow:SH.card});
  [0.707, 2.821, 4.936].forEach(function (y) { sh(s, 'ellipse', 6.667, y, 1.757, 1.757, {fill:C.WHITE}); });
  rt(s, [['High '], ['Technology', {c:C.ORANGE}], [' With Best '], ['IT Solution', {c:C.ORANGE}]], 1.144, 3.938, 3.348, 0.375, {c:C.PURPLE, f:F.SEMI, ls:1.5});
  tx(s, lorem(20), 1.144, 4.355, 3.348, 0.871, {sz:10, c:C.GREY, al:'justify', ls:1.5});
  tx(s, 'Let’s Go', 2.636, 5.226, 1.729, 0.426, {sh:'roundRect', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, al:'center', r:0.213, shadow:SH.card});
  tx(s, '01', 6.48, 0.807, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, al:'center', va:'middle'});
  tx(s, 'Insert Text Here', 8.938, 0.707, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(22), 8.938, 0.935, 3.243, 1.284, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, '02', 6.48, 2.916, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, al:'center', va:'middle'});
  tx(s, '03', 6.48, 4.964, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, al:'center', va:'middle'});
  tx(s, 'Insert Text Here', 8.938, 3.025, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(22), 8.938, 3.253, 3.243, 1.284, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, 'Insert Text Here', 8.938, 5.042, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(22), 8.938, 5.27, 3.243, 1.284, {c:C.GREY, al:'justify', ls:1.5});
}

// Slide 25 - three photo cards
function slide25(s) {
  panel(s, 0, 3.614, 13.333, 3.886);
  [1.187, 4.882, 8.576].forEach(function (x) {
    sh(s, 'roundRect', x, 2.182, 3.348, 3.552, {fill:C.WHITE, r:0.282, shadow:SH.card});
  });
  tx(s, 'The Best Way To Coding New IT Solutions', 1.294, 0.808, 6.149, 1.178, {sz:32, b:1, c:C.PURPLE});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 1.274, 0.446, 3.623, 0.375, {b:1, c:C.ORANGE, ls:1.5});
  tx(s, 'Insert Text Here', 1.361, 4.671, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(10), 1.361, 4.899, 3, 0.678, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, 'Insert Text Here', 5.056, 4.671, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(10), 5.056, 4.899, 3, 0.678, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, 'Insert Text Here', 8.75, 4.671, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(10), 8.75, 4.899, 3, 0.678, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, lorem(41), 1.294, 6.206, 7.649, 0.981, {c:C.GREY, al:'justify', ls:1.5});
  pill(s, 10.037, 0.312);
}

// Slide 26 - study case + panel
function slide26(s) {
  panel(s, 0, 3.614, 13.333, 3.886);
  sh(s, 'roundRect', 1.847, 4.775, 6.753, 2.086, {fill:C.PURPLE, r:0.09});
  tx(s, 'High Technology With Best IT Solution', 2.269, 4.908, 4.856, 0.467, {sz:16, c:C.WHITE, f:F.SEMI, ls:1.5});
  tx(s, 'Insert Text Here', 2.559, 5.476, 2.361, 0.303, {c:C.WHITE, f:F.SEMI});
  check(s, 2.269, 5.482, 0.29);
  tx(s, lorem(8, 1), 2.559, 5.741, 2.674, 0.981, {c:C.LIGHT, ls:1.5, bul:1});
  tx(s, 'Insert Text Here', 5.571, 5.476, 2.361, 0.303, {c:C.WHITE, f:F.SEMI});
  check(s, 5.281, 5.482, 0.29);
  tx(s, lorem(8, 1), 5.571, 5.741, 2.674, 0.981, {c:C.LIGHT, ls:1.5, bul:1});
  tx(s, 'Study Case IT Consultancy', 1.781, 1.001, 3.623, 1.178, {sz:32, b:1, c:C.PURPLE});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 1.761, 0.639, 3.623, 0.375, {b:1, c:C.ORANGE, ls:1.5});
  sh(s, 'roundRect', 9.866, 1.447, 2.662, 2.257, {fill:C.WHITE, r:0, flipH:1, shadow:SH.drop});
  tx(s, '01', 10.919, 1.584, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, al:'center', va:'middle'});
  tx(s, 'Insert Text Here', 10.032, 2.354, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, lorem(11), 9.866, 2.582, 2.693, 0.981, {c:C.GREY, al:'center', ls:1.5});
}

// Slide 27 - mockup web developments
function slide27(s) {
  panel(s, 0, 3.614, 13.333, 3.886);
  photo(s, 7.408, 1.798, 5.279, 5.259, {c:'E6E6E6'});      // [image] desktop mock-up
  sh(s, 'rect', 0.564, 0.486, 6.103, 3.889, {fill:C.WHITE, flipH:1, shadow:SH.drop});
  tx(s, '01.', 0.865, 4.746, 0.683, 0.505, {sz:24, b:1, c:C.PURPLE, al:'center'});
  tx(s, 'Insert Text Here', 0.865, 5.264, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(15), 0.865, 5.492, 2.877, 0.981, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, 'Mockup Web Developments', 0.976, 1.22, 3.997, 1.178, {sz:32, b:1, c:C.PURPLE});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 0.956, 0.858, 3.623, 0.375, {b:1, c:C.ORANGE, ls:1.5});
  tx(s, lorem(41), 0.976, 2.411, 5.218, 1.587, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, '02.', 4.085, 4.746, 0.683, 0.505, {sz:24, b:1, c:C.PURPLE, al:'center'});
  tx(s, 'Insert Text Here', 4.085, 5.264, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(15), 4.085, 5.492, 2.877, 0.981, {c:C.GREY, al:'justify', ls:1.5});
  pill(s, 10.037, 0.312);
}

// Slide 28 - mockup web developments (laptops)
function slide28(s) {
  panel(s, 0, 3.614, 13.333, 3.886);
  sh(s, 'rect', 0, 6.611, 13.333, 0.889, {fill:C.PURPLE});
  photo(s, 7.653, 1.873, 5.396, 5.627, {c:'DDDDDD'});      // [image] laptop mock-ups
  photo(s, 1.05, 0.99, 3.21, 3.348, {c:'F4F4F4'});
  sh(s, 'rect', 1.114, 3.75, 3.389, 1.367, {fill:C.WHITE, flipH:1, shadow:SH.drop});
  tx(s, '01', 1.539, 4.117, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, al:'center', va:'middle'});
  tx(s, 'Insert Text Here', 2.142, 3.968, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(6), 2.142, 4.196, 2.361, 0.678, {c:C.GREY, ls:1.5});
  sh(s, 'rect', 1.114, 5.497, 3.389, 1.367, {fill:C.WHITE, flipH:1, shadow:SH.drop});
  tx(s, '02', 1.539, 5.864, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, al:'center', va:'middle'});
  tx(s, 'Insert Text Here', 2.142, 5.716, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(6), 2.142, 5.944, 2.361, 0.678, {c:C.GREY, ls:1.5});
  tx(s, 'Mockup Web Developments', 5.227, 1.86, 3.997, 1.178, {sz:32, b:1, c:C.PURPLE});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 5.207, 1.498, 3.623, 0.375, {b:1, c:C.ORANGE, ls:1.5});
  tx(s, lorem(14), 5.227, 3.051, 4.218, 0.678, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, lorem(32), 5.227, 3.833, 4.218, 1.284, {c:C.GREY, al:'justify', ls:1.5});
}

// Slide 29 - expert web developments
function slide29(s) {
  panel(s, 0, 1.056, 13.333, 6.444);
  photo(s, -1.318, 3.75, 8.874, 4.933, {c:'F2F2F2'});      // [image] tablet mock-up
  tick(s, 3.823, 2.717);
  tx(s, 'UI/UX Designs', 4.155, 2.694, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tick(s, 3.823, 3.218);
  tx(s, 'Web Development', 4.155, 3.195, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, 'Database Analysis', 1.795, 2.694, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tick(s, 1.435, 2.717);
  tx(s, 'Happy Customers', 1.795, 3.195, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tick(s, 1.435, 3.218);
  tx(s, 'Expert Web Developments', 1.283, 1.241, 3.997, 1.178, {sz:32, b:1, c:C.PURPLE});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 1.263, 0.879, 3.623, 0.375, {b:1, c:C.ORANGE, ls:1.5});
  tx(s, lorem(41), 6.955, 1.924, 5.218, 1.587, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, '01', 8.212, 3.761, 0.857, 0.788, {sz:36, b:1, c:C.ORANGE, ls:1.2});
  tx(s, 'Insert Text Here', 9.07, 3.761, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(16), 9.07, 3.989, 3.104, 0.981, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, '02', 8.212, 5.221, 0.857, 0.788, {sz:36, b:1, c:C.ORANGE, ls:1.2});
  tx(s, 'Insert Text Here', 9.07, 5.221, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(16), 9.07, 5.449, 3.104, 0.981, {c:C.GREY, al:'justify', ls:1.5});
}

// Slide 30 - expert web developments (phones)
function slide30(s) {
  panel(s, 0, 0, 13.333, 7.5);
  dome(s, 0, 0, 4.968, 7.5, C.PURPLE);
  photo(s, 7.724, 1.0, 2.776, 4.626, {rot:-17.7, c:'D2D1D1'}); // [image] phone mock-up
  photo(s, 9.723, 0.312, 2.776, 4.626, {c:'EDEDED', rot:-342.339});
  sh(s, 'rect', 0.882, 1.564, 6.394, 1.4, {fill:C.WHITE, flipH:1, shadow:SH.drop});
  tx(s, '01', 1.111, 1.92, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, va:'middle'});
  tx(s, 'Insert Text Here', 1.714, 1.771, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(22), 1.714, 1.999, 5.381, 0.678, {c:C.GREY, al:'justify', ls:1.5});
  sh(s, 'rect', 0.882, 3.192, 6.394, 1.4, {fill:C.WHITE, flipH:1, shadow:SH.drop});
  tx(s, '02', 1.111, 3.548, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, va:'middle'});
  tx(s, 'Insert Text Here', 1.714, 3.399, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(22), 1.714, 3.627, 5.381, 0.678, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, 'Expert Web Developments', 7.935, 6.052, 3.997, 1.178, {sz:32, b:1, c:C.PURPLE});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 7.915, 5.69, 3.623, 0.375, {b:1, c:C.ORANGE, ls:1.5});
  sh(s, 'rect', 0.882, 4.82, 6.394, 1.4, {fill:C.WHITE, flipH:1, shadow:SH.drop});
  tx(s, '03', 1.111, 5.176, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, va:'middle'});
  tx(s, 'Insert Text Here', 1.714, 5.027, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(22), 1.714, 5.255, 5.381, 0.678, {c:C.GREY, al:'justify', ls:1.5});
  pill(s, 10.037, 0.312);
}

// Slide 31 - mockup of sahala
function slide31(s) {
  panel(s, 0, 0, 13.333, 7.5);
  sh(s, 'roundRect', 8.832, 0.856, 4.007, 5.689, {fill:C.MIST, r:0.28, shadow:SH.card});
  sh(s, 'ellipse', 10.689, 6.141, 0.298, 0.303, {fill:C.SILVER});
  sh(s, 'roundRect', 4.821, 1.297, 3.361, 5.248, {fill:C.PURPLE, r:0.168});
  sh(s, 'roundRect', 0.892, 1.297, 3.361, 5.248, {fill:C.WHITE, r:0.168, shadow:SH.card});
  tx(s, 'Insert Text Here', 1.087, 1.81, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(15), 1.087, 2.037, 2.877, 0.981, {c:C.GREY, al:'justify', ls:1.5});
  hline(s, 1.218, 3.299, 2.436, {line:C.SILVER, lw:6});
  hline(s, 1.218, 3.299, 2.23, {line:C.ORANGE, lw:6});
  hline(s, 1.218, 3.813, 2.436, {line:C.SILVER, lw:6});
  hline(s, 1.218, 3.813, 1.379, {line:C.ORANGE, lw:6});
  tx(s, 'Insert Text Here', 5.086, 1.81, 2.361, 0.303, {c:C.WHITE, f:F.SEMI});
  tx(s, lorem(15), 5.086, 2.037, 2.877, 0.981, {c:C.LIGHT, al:'justify', ls:1.5});
  hline(s, 5.216, 3.299, 2.436, {line:C.SILVER, lw:6});
  hline(s, 5.216, 3.299, 2.23, {line:C.ORANGE, lw:6});
  hline(s, 5.216, 3.813, 2.436, {line:C.SILVER, lw:6});
  hline(s, 5.216, 3.813, 1.379, {line:C.ORANGE, lw:6});
  sh(s, 'roundRect', 5.824, 5.9, 5.333, 0.987, {fill:C.WHITE, r:0.493, flipH:1, shadow:SH.drop});
  tx(s, 'Mockup of Sahala', 6.178, 6.048, 4.576, 0.64, {sz:32, b:1, c:C.PURPLE});
  pill(s, 8.49, 5.498);
}

// Slide 32 - mockup of sahala (browser)
function slide32(s) {
  panel(s, 0, 0, 13.333, 7.5);
  browserFrame(s, 0.705, 2.202, 6.18, 4.018);
  sh(s, 'round1Rect', 8.586, 1.069, 4.748, 6.431, {fill:C.PURPLE, r:2.096, flipH:1});
  tx(s, 'Key Solutions', 9.872, 5.685, 2.361, 0.303, {c:C.WHITE, f:F.SEMI, al:'center'});
  tx(s, lorem(15), 9.476, 5.912, 3.152, 0.871, {sz:10.5, c:C.LIGHT, al:'center', ls:1.5});
  tx(s, 'High Technology With Best IT Solution', 9.623, 2.388, 3.005, 0.871, {sz:16, c:C.WHITE, f:F.SEMI, ls:1.5});
  sh(s, 'roundRect', 4.907, 1.281, 5.333, 0.987, {fill:C.WHITE, r:0.493, flipH:1, shadow:SH.drop});
  tx(s, 'Mockup of Sahala', 5.261, 1.428, 4.576, 0.64, {sz:32, b:1, c:C.PURPLE});
  pill(s, 7.574, 0.878);
  sh(s, 'roundRect', 5.2, 4.943, 2.662, 2.257, {fill:C.WHITE, r:0, flipH:1, shadow:SH.drop});
  tx(s, '02', 6.254, 5.081, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, al:'center', va:'middle'});
  tx(s, 'Insert Text Here', 5.367, 5.851, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, lorem(11), 5.2, 6.079, 2.693, 0.981, {c:C.GREY, al:'center', ls:1.5});
  sh(s, 'roundRect', 2.142, 4.943, 2.662, 2.257, {fill:C.WHITE, r:0, flipH:1, shadow:SH.drop});
  tx(s, '01', 3.195, 5.081, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, al:'center', va:'middle'});
  tx(s, 'Insert Text Here', 2.308, 5.851, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, lorem(11), 2.142, 6.079, 2.693, 0.981, {c:C.GREY, al:'center', ls:1.5});
}

// Slide 33 - pricing plans
function slide33(s) {
  panel(s, 0, 1.861, 13.333, 3.097);
  panel(s, 0, 3.542, 13.333, 3.958);
  sh(s, 'roundRect', 1.679, 2.174, 3.214, 4.736, {fill:C.PURPLE, line:C.WHITE, r:0.22, shadow:SH.deep});
  tx(s, lorem(7), 1.979, 4.98, 2.565, 0.582, {sz:10, c:C.LIGHT, al:'center', ls:1.5});
  rt(s, [['$'], ['123'], [',', {sz:16}], ['45 ', {sz:16}], ['/monthly', {sz:11}]], 1.884, 2.863, 2.755, 0.64, {sz:36, c:C.ORANGE, f:F.SEMI, al:'center'});
  tx(s, 'IT Consultancy', 1.933, 3.729, 1.89, 0.278, {sz:10.5, c:C.WHITE, f:F.SEMI, bul:1});
  tx(s, 'Digital User Interface', 1.933, 4.05, 2.706, 0.278, {sz:10.5, c:C.WHITE, f:F.SEMI, bul:1});
  tx(s, 'Database Center', 1.933, 4.351, 2.21, 0.278, {sz:10.5, c:C.WHITE, f:F.SEMI, bul:1});
  tx(s, 'Network Maintenance', 1.933, 4.666, 2.706, 0.278, {sz:10.5, c:C.WHITE, f:F.SEMI, bul:1});
  tx(s, '7/24 Full Supports', 2.118, 5.65, 2.336, 0.278, {sz:10.5, c:C.WHITE, f:F.SEMI, al:'center'});
  tx(s, 'Standard', 2.118, 2.527, 2.367, 0.337, {sz:14, b:1, c:C.WHITE, al:'center'});
  pill(s, 10.037, 0.312);
  tx(s, 'Purchase Now', 2.413, 6.155, 1.729, 0.426, {sh:'roundRect', fill:C.WHITE, c:C.PURPLE, f:F.SEMI, al:'center', r:0.213, shadow:SH.card});
  sh(s, 'roundRect', 8.434, 2.174, 3.214, 4.736, {fill:C.PURPLE, line:C.WHITE, r:0.22, shadow:SH.deep});
  tx(s, lorem(7), 8.734, 4.98, 2.565, 0.582, {sz:10, c:C.LIGHT, al:'center', ls:1.5});
  rt(s, [['$'], ['323'], [',', {sz:16}], ['45 ', {sz:16}], ['/monthly', {sz:11}]], 8.639, 2.863, 2.755, 0.707, {sz:36, c:C.ORANGE, f:F.SEMI, al:'center'});
  tx(s, 'IT Consultancy', 8.688, 3.729, 1.89, 0.278, {sz:10.5, c:C.WHITE, f:F.SEMI, bul:1});
  tx(s, 'Digital User Interface', 8.688, 4.05, 2.706, 0.278, {sz:10.5, c:C.WHITE, f:F.SEMI, bul:1});
  tx(s, 'Database Center', 8.688, 4.351, 2.21, 0.278, {sz:10.5, c:C.WHITE, f:F.SEMI, bul:1});
  tx(s, 'Network Maintenance', 8.688, 4.666, 2.706, 0.278, {sz:10.5, c:C.WHITE, f:F.SEMI, bul:1});
  tx(s, '7/24 Full Supports', 8.873, 5.65, 2.336, 0.278, {sz:10.5, c:C.WHITE, f:F.SEMI, al:'center'});
  tx(s, 'Popular', 8.873, 2.527, 2.367, 0.337, {sz:14, b:1, c:C.WHITE, al:'center'});
  tx(s, 'Purchase Now', 9.168, 6.155, 1.729, 0.426, {sh:'roundRect', fill:C.WHITE, c:C.PURPLE, f:F.SEMI, al:'center', r:0.213, shadow:SH.card});
  sh(s, 'roundRect', 5.06, 1.896, 3.214, 5.292, {fill:C.WHITE, line:C.WHITE, r:0.22, shadow:SH.deep});
  tx(s, lorem(7), 5.36, 5.031, 2.565, 0.582, {sz:10, c:C.PURPLE, al:'center', ls:1.5});
  rt(s, [['$'], ['523'], [',', {sz:16}], ['45 ', {sz:16}], ['/monthly', {sz:11}]], 5.265, 2.666, 2.755, 0.715, {sz:36, c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, 'IT Consultancy', 5.313, 3.634, 1.89, 0.278, {sz:10.5, c:C.PURPLE, f:F.SEMI, bul:1});
  tx(s, 'Digital User Interface', 5.313, 3.993, 2.706, 0.278, {sz:10.5, c:C.PURPLE, f:F.SEMI, bul:1});
  tx(s, 'Database Center', 5.313, 4.329, 2.21, 0.278, {sz:10.5, c:C.PURPLE, f:F.SEMI, bul:1});
  tx(s, 'Network Maintenance', 5.313, 4.68, 2.706, 0.278, {sz:10.5, c:C.PURPLE, f:F.SEMI, bul:1});
  tx(s, '7/24 Full Supports', 5.498, 5.78, 2.336, 0.278, {sz:10.5, c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, 'Premium', 5.498, 2.29, 2.367, 0.337, {sz:14, b:1, c:C.PURPLE, al:'center'});
  tx(s, 'Purchase Now', 5.794, 6.344, 1.729, 0.426, {sh:'roundRect', fill:C.PURPLE, c:C.ORANGE, f:F.SEMI, al:'center', r:0.213, shadow:SH.card});
  tx(s, 'Pricing Plans', 1.825, 0.919, 3.997, 0.64, {sz:32, b:1, c:C.PURPLE});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 1.805, 0.557, 3.623, 0.375, {b:1, c:C.ORANGE, ls:1.5});
}

// Slide 34 - pricing plans (two)
function slide34(s) {
  panel(s, 4.278, 1.542, 7.5, 4.417, {rot:270});
  panel(s, 7.375, 1.542, 7.5, 4.417, {rot:270});
  sh(s, 'roundRect', 9.624, 2.271, 3.214, 3.979, {fill:C.PURPLE, line:C.WHITE, r:0.22, shadow:SH.deep});
  tx(s, 'Get 50% Off For 1 Years Plans Technology Corp.', 10.111, 5.189, 2.239, 0.678, {c:C.WHITE, f:F.SEMI, al:'center', ls:1.5});
  rt(s, [['$'], ['123'], [',', {sz:16}], ['45 ', {sz:16}], ['/monthly', {sz:11}]], 9.853, 2.96, 2.755, 0.64, {sz:36, c:C.ORANGE, f:F.SEMI, al:'center'});
  tx(s, 'IT Consultancy', 10.285, 3.813, 1.89, 0.278, {sz:10.5, c:C.WHITE, f:F.SEMI, al:'center'});
  tx(s, 'Digital User Interface', 9.877, 4.134, 2.706, 0.278, {sz:10.5, c:C.WHITE, f:F.SEMI, al:'center'});
  tx(s, 'Database Center', 10.126, 4.435, 2.21, 0.278, {sz:10.5, c:C.WHITE, f:F.SEMI, al:'center'});
  tx(s, 'Network Maintenance', 9.877, 4.749, 2.706, 0.278, {sz:10.5, c:C.WHITE, f:F.SEMI, al:'center'});
  tx(s, 'Standard', 10.047, 2.624, 2.367, 0.337, {sz:14, b:1, c:C.WHITE, al:'center'});
  tx(s, 'Purchase Now', 8.683, 3.067, 1.729, 0.426, {sh:'roundRect', fill:C.WHITE, c:C.PURPLE, f:F.SEMI, al:'center', r:0.213, rot:270, shadow:SH.card});
  sh(s, 'roundRect', 5.867, 2.056, 3.214, 4.41, {fill:C.WHITE, line:C.WHITE, r:0.22, shadow:SH.deep});
  tx(s, 'Get 50% Off For 1 Years Plans Technology Corp.', 6.354, 5.29, 2.239, 0.678, {c:C.PURPLE, f:F.SEMI, al:'center', ls:1.5});
  rt(s, [['$'], ['423'], [',', {sz:16}], ['45 ', {sz:16}], ['/monthly', {sz:11}]], 6.097, 2.82, 2.755, 0.709, {sz:36, c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, 'IT Consultancy', 6.529, 3.764, 1.89, 0.278, {sz:10.5, c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, 'Digital User Interface', 6.121, 4.12, 2.706, 0.278, {sz:10.5, c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, 'Database Center', 6.369, 4.453, 2.21, 0.278, {sz:10.5, c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, 'Network Maintenance', 6.121, 4.802, 2.706, 0.278, {sz:10.5, c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, 'Premium', 6.29, 2.447, 2.367, 0.337, {sz:14, b:1, c:C.PURPLE, al:'center'});
  tx(s, 'Purchase Now', 4.926, 3.067, 1.729, 0.426, {sh:'roundRect', fill:C.WHITE, c:C.PURPLE, f:F.SEMI, al:'center', r:0.213, rot:270, shadow:SH.card});
  tx(s, 'Pricing Plans', 0.77, 2.777, 3.997, 0.64, {sz:32, b:1, c:C.PURPLE});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 0.75, 2.415, 3.623, 0.375, {b:1, c:C.ORANGE, ls:1.5});
  tx(s, lorem(41), 0.806, 3.457, 4.189, 1.89, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, lorem(14), 0.806, 5.386, 4.189, 0.678, {c:C.GREY, al:'justify', ls:1.5});
  pill(s, 10.037, 0.312);
}

// Slide 35 - expert charts - segmented ring
function slide35(s) {
  panel(s, 0, 0, 13.333, 7.5);
  sh(s, 'round1Rect', 0.375, 0, 12.958, 0.556, {fill:C.PURPLE, r:0.245});
  // six-part segmented ring, grey hub in the middle
  segRing(s, 6.646, 4.764, 2.21, 0.606,
    [[-28, 28, C.PURPLE], [32, 88, C.ORANGE], [92, 148, C.BLUE],
     [152, 208, C.PURPLE], [212, 268, C.ORANGE], [272, 328, C.BLUE]]);
  sh(s, 'ellipse', 5.818, 3.932, 1.64, 1.642, {fill:C.SILVER});
  tx(s, lorem(31, 1), 2.482, 1.65, 8.042, 0.678, {c:C.GREY, al:'center', ls:1.5});
  tx(s, 'Expert Charts Sahala', 3.354, 0.99, 6.299, 0.64, {sz:32, b:1, c:C.PURPLE, al:'center'});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 4.692, 0.628, 3.623, 0.375, {b:1, c:C.ORANGE, al:'center', ls:1.5});
  tx(s, '01', 1.912, 2.621, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, sz:11, c:C.WHITE, f:F.SEMI, al:'center', va:'middle'});
  tx(s, 'Insert Text Here', 1.025, 3.391, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, lorem(10), 0.506, 3.619, 3.398, 0.678, {c:C.GREY, al:'center', ls:1.5});
  tx(s, '02', 1.912, 4.525, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, sz:11, c:C.WHITE, f:F.SEMI, al:'center', va:'middle'});
  tx(s, 'Insert Text Here', 1.025, 5.295, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, lorem(10), 0.506, 5.523, 3.398, 0.678, {c:C.GREY, al:'center', ls:1.5});
  tx(s, '03', 10.715, 2.621, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, sz:11, c:C.WHITE, f:F.SEMI, al:'center', va:'middle'});
  tx(s, 'Insert Text Here', 9.828, 3.391, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, lorem(10), 9.309, 3.619, 3.398, 0.678, {c:C.GREY, al:'center', ls:1.5});
  tx(s, '04', 10.715, 4.525, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, sz:11, c:C.WHITE, f:F.SEMI, al:'center', va:'middle'});
  tx(s, 'Insert Text Here', 9.828, 5.295, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI, al:'center'});
  tx(s, lorem(10), 9.309, 5.523, 3.398, 0.678, {c:C.GREY, al:'center', ls:1.5});
}

// Slide 36 - expert charts - doughnut chart
function slide36(s) {
  panel(s, 0, 0, 13.333, 7.5);
  sh(s, 'round1Rect', 0.375, 0, 12.958, 0.556, {fill:C.PURPLE, r:0.245});
  donutChart(s, 3.338, 2.3, 6.657, 4.244);
  sh(s, 'round2SameRect', 10.091, 1.533, 1.552, 4.426, {fill:C.PURPLE, rot:90, shadow:SH.deep});
  sh(s, 'round2SameRect', 10.091, 3.22, 1.552, 4.426, {fill:C.PURPLE, rot:90, shadow:SH.deep});
  tx(s, '03', 8.927, 3.287, 0.774, 0.774, {sh:'ellipse', fill:C.ORANGE, sz:16, b:1, c:C.WHITE, al:'center', va:'middle'});
  tx(s, lorem(14), 9.869, 3.415, 2.971, 0.981, {c:C.LIGHT, ls:1.5});
  tx(s, 'Insert Text Here', 9.869, 3.18, 2.1, 0.303, {c:C.WHITE, f:F.SEMI});
  tx(s, '04', 8.927, 4.982, 0.774, 0.774, {sh:'ellipse', fill:C.ORANGE, sz:16, b:1, c:C.WHITE, al:'center', va:'middle'});
  tx(s, lorem(14), 9.869, 5.11, 2.971, 0.981, {c:C.LIGHT, ls:1.5});
  tx(s, 'Insert Text Here', 9.869, 4.874, 2.1, 0.303, {c:C.WHITE, f:F.SEMI});
  sh(s, 'round2SameRect', 1.788, 1.533, 1.552, 4.426, {fill:C.PURPLE, rot:270, flipH:1, shadow:SH.deep});
  sh(s, 'round2SameRect', 1.788, 3.22, 1.552, 4.426, {fill:C.PURPLE, rot:270, flipH:1, shadow:SH.deep});
  tx(s, '01', 3.729, 3.287, 0.774, 0.774, {sh:'ellipse', fill:C.ORANGE, sz:16, b:1, c:C.WHITE, al:'center', va:'middle', flipH:1});
  tx(s, lorem(14), 0.645, 3.415, 2.916, 0.981, {c:C.LIGHT, al:'right', ls:1.5});
  tx(s, 'Insert Text Here', 1.461, 3.18, 2.1, 0.303, {c:C.WHITE, f:F.SEMI, al:'right'});
  tx(s, '02', 3.729, 4.982, 0.774, 0.774, {sh:'ellipse', fill:C.ORANGE, sz:16, b:1, c:C.WHITE, al:'center', va:'middle', flipH:1});
  tx(s, lorem(14), 0.645, 5.11, 2.916, 0.981, {c:C.LIGHT, al:'right', ls:1.5});
  tx(s, 'Insert Text Here', 1.461, 4.874, 2.1, 0.303, {c:C.WHITE, f:F.SEMI, al:'right'});
  tx(s, lorem(31, 1), 2.482, 1.65, 8.042, 0.678, {c:C.GREY, al:'center', ls:1.5});
  tx(s, 'Expert Charts Sahala', 3.354, 0.99, 6.299, 0.64, {sz:32, b:1, c:C.PURPLE, al:'center'});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 4.692, 0.628, 3.623, 0.375, {b:1, c:C.ORANGE, al:'center', ls:1.5});
}

// Slide 37 - expert charts - interlocked rings
function slide37(s) {
  panel(s, 0, 0, 13.333, 7.5);
  sh(s, 'round1Rect', 0.375, 0, 12.958, 0.556, {fill:C.PURPLE, r:0.245});
  // three interlocked rings
  sh(s, 'donut', 4.853, 2.811, 1.693, 1.695, {fill:C.SILVER});
  sh(s, 'donut', 6.652, 2.811, 1.696, 1.695, {fill:C.PURPLE});
  sh(s, 'donut', 5.754, 1.912, 1.693, 1.693, {fill:C.ORANGE});
  sh(s, 'ellipse', 1.515, 2.274, 2.646, 2.392, {fill:C.WHITE});
  tx(s, 'Technology', 2.15, 3.402, 1.375, 0.286, {sz:11, c:C.ORANGE, f:F.SEMI, al:'center'});
  tx(s, '45.25', 2.15, 2.972, 1.375, 0.572, {sz:28, b:1, c:C.PURPLE, al:'center'});
  gaugeRing(s, 2.838, 3.414, 0.855,
    [C.ORANGE, C.ORANGE, C.ORANGE, C.MIST, C.MIST, C.MIST, C.MIST, C.MIST]);
  plusBadge(s, 3.688, 2.274, 0.381);
  sh(s, 'ellipse', 9.074, 2.274, 2.646, 2.392, {fill:C.WHITE});
  tx(s, 'IT Solutions', 9.71, 3.402, 1.375, 0.286, {sz:11, c:C.ORANGE, f:F.SEMI, al:'center'});
  tx(s, '95.10', 9.71, 2.972, 1.375, 0.572, {sz:28, b:1, c:C.PURPLE, al:'center'});
  gaugeRing(s, 10.397, 3.414, 0.855,
    [C.ORANGE, C.ORANGE, C.ORANGE, C.ORANGE, C.ORANGE, C.ORANGE, C.ORANGE, C.MIST]);
  plusBadge(s, 11.248, 2.274, 0.381);
  tx(s, '02', 4.698, 5.213, 0.857, 0.788, {sz:36, b:1, c:C.ORANGE, ls:1.2});
  tx(s, 'Insert Text Here', 5.556, 5.213, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(22), 5.556, 5.44, 3.104, 1.284, {c:C.BLACK, al:'justify', ls:1.5});
  tx(s, '03', 8.848, 5.213, 0.857, 0.788, {sz:36, b:1, c:C.ORANGE, ls:1.2});
  tx(s, 'Insert Text Here', 9.706, 5.213, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(22), 9.706, 5.44, 3.104, 1.284, {c:C.BLACK, al:'justify', ls:1.5});
  tx(s, '01', 0.586, 5.213, 0.857, 0.788, {sz:36, b:1, c:C.ORANGE, ls:1.2});
  tx(s, 'Insert Text Here', 1.443, 5.213, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(22), 1.443, 5.44, 3.104, 1.284, {c:C.BLACK, al:'justify', ls:1.5});
  tx(s, 'Expert Charts Sahala', 3.354, 0.99, 6.299, 0.64, {sz:32, b:1, c:C.PURPLE, al:'center'});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 4.692, 0.628, 3.623, 0.375, {b:1, c:C.ORANGE, al:'center', ls:1.5});
}

// Slide 38 - expert charts - pinwheel
function slide38(s) {
  panel(s, 0, 0, 13.333, 7.5);
  sh(s, 'round1Rect', 0.375, 0, 12.958, 0.556, {fill:C.PURPLE, r:0.245});
  // four-blade pinwheel
  pinwheel(s, 6.664, 4.671, 2.248, [C.PURPLE, C.ORANGE, C.PURPLE, C.ORANGE]);
  sh(s, 'ellipse', 6.188, 2.598, 1.038, 1.035, {fill:C.WHITE, line:C.HAIRLINE, lw:2.25});
  sh(s, 'ellipse', 6.14, 5.728, 1.037, 1.037, {fill:C.WHITE, line:C.HAIRLINE, lw:2.25});
  sh(s, 'ellipse', 4.595, 4.12, 1.038, 1.037, {fill:C.WHITE, line:C.HAIRLINE, lw:2.25});
  sh(s, 'ellipse', 7.7, 4.154, 1.038, 1.035, {fill:C.WHITE, line:C.HAIRLINE, lw:2.25});
  tx(s, '01', 6.417, 2.806, 0.605, 0.608, {sh:'ellipse', fill:C.ORANGE, c:C.WHITE, f:F.SEMI, al:'center', va:'middle'});
  tx(s, '02', 7.951, 4.361, 0.605, 0.608, {sh:'ellipse', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, va:'middle'});
  tx(s, '03', 6.356, 5.95, 0.605, 0.608, {sh:'ellipse', fill:C.ORANGE, c:C.WHITE, f:F.SEMI, al:'center', va:'middle'});
  tx(s, '04', 4.816, 4.353, 0.605, 0.608, {sh:'ellipse', fill:C.PURPLE, c:C.WHITE, f:F.SEMI, va:'middle'});
  tx(s, '01', 0.523, 2.642, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, sz:11, c:C.WHITE, f:F.SEMI, va:'middle'});
  tx(s, 'Insert Text Here', 1.127, 2.493, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(22), 1.127, 2.721, 2.94, 1.284, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, '02', 0.523, 4.654, 0.587, 0.589, {sh:'ellipse', fill:C.ORANGE, sz:11, c:C.WHITE, f:F.SEMI, va:'middle'});
  tx(s, 'Insert Text Here', 1.127, 4.506, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(22), 1.127, 4.734, 2.94, 1.284, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, '03', 9.186, 2.642, 0.587, 0.589, {sh:'ellipse', fill:C.ORANGE, sz:11, c:C.WHITE, f:F.SEMI, va:'middle'});
  tx(s, 'Insert Text Here', 9.79, 2.493, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(22), 9.79, 2.721, 2.94, 1.284, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, '04', 9.186, 4.654, 0.587, 0.589, {sh:'ellipse', fill:C.PURPLE, sz:11, c:C.WHITE, f:F.SEMI, va:'middle'});
  tx(s, 'Insert Text Here', 9.79, 4.506, 2.361, 0.303, {c:C.PURPLE, f:F.SEMI});
  tx(s, lorem(22), 9.79, 4.734, 2.94, 1.284, {c:C.GREY, al:'justify', ls:1.5});
  tx(s, lorem(31, 1), 2.482, 1.65, 8.042, 0.678, {c:C.GREY, al:'center', ls:1.5});
  tx(s, 'Expert Charts Sahala', 3.354, 0.99, 6.299, 0.64, {sz:32, b:1, c:C.PURPLE, al:'center'});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 4.692, 0.628, 3.623, 0.375, {b:1, c:C.ORANGE, al:'center', ls:1.5});
}

// Slide 39 - stay touch with us
function slide39(s) {
  panel(s, 0, 0, 13.333, 7.5);
  sh(s, 'rect', 0, 4.139, 6.472, 2.681, {fill:C.PURPLE});
  sh(s, 'round1Rect', 6.463, 4.125, 4.255, 2.694, {fill:C.WHITE, line:C.WHITE, shadow:SH.deep});
  tx(s, 'Address Company', 7.668, 4.577, 1.82, 0.303, {c:C.PURPLE, f:F.MED});
  rt(s, [['76 United Street Ca,'], ['\n'], ['California, CA 62 71593']], 7.668, 4.812, 2.186, 0.627, {sz:11, c:C.GREY, ls:1.5});
  sh(s, 'ellipse', 7.034, 4.625, 0.473, 0.472, {fill:C.PURPLE, shadow:SH.drop});
  planeIcon(s, 7.148, 4.755, 0.223, 0.206);
  tx(s, 'Phone Contact ', 7.668, 5.577, 1.745, 0.303, {c:C.PURPLE, f:F.MED});
  tx(s, '+09 235 854 787', 7.668, 5.82, 1.939, 0.349, {sz:11, c:C.GREY, ls:1.5});
  sh(s, 'ellipse', 7.034, 5.622, 0.473, 0.472, {fill:C.AMBER, shadow:SH.drop});
  phoneIcon(s, 7.156, 5.743, 0.228, 0.229);
  tx(s, '+65 235 854 787', 7.668, 6.119, 1.939, 0.349, {sz:11, c:C.GREY, ls:1.5});
  birdIcon(s, 4.384, 6.104, 0.274, 0.22);
  tx(s, 'Twi.account', 4.711, 6.078, 1.471, 0.321, {c:C.WHITE, f:F.SEMI, ls:1.14});
  fbIcon(s, 0.384, 6.054, 0.253, 0.255);
  tx(s, 'Fac.account', 0.704, 6.051, 1.515, 0.321, {c:C.WHITE, f:F.SEMI, ls:1.14});
  mailIcon(s, 2.39, 6.098, 0.277, 0.182);
  tx(s, 'Email.Contact', 2.718, 6.062, 1.998, 0.321, {c:C.WHITE, f:F.SEMI, ls:1.14});
  tx(s, 'Connect with us:', 0.279, 5.472, 2.186, 0.421, {sz:14, c:C.WHITE, f:F.SEMI, ls:1.5});
  tx(s, 'Stay Touch With Us', 0.384, 3.298, 5.394, 0.64, {sz:32, b:1, c:C.PURPLE});
  rt(s, [['IT Solutions & Technology '], ['®2021', {b:0}]], 0.384, 2.937, 3.623, 0.375, {b:1, c:C.ORANGE, ls:1.5});
  tx(s, lorem(22), 0.384, 4.557, 5.535, 0.678, {c:C.LIGHT, al:'justify', ls:1.5});
  pill(s, 10.037, 0.312);
}

// Slide 40 - thank you
function slide40(s) {
  sh(s, 'rect', 0.305, 0.611, 12.724, 6.278, {fill:C.PURPLE});
  pill(s, 6.18, 2.336);
  wordmark(s, 0.813, 1.01, C.WHITE);
  tx(s, 'New IT Solutions & Technology Intelligence', 0.985, 6.011, 2.557, 0.678, {c:C.WHITE, ls:1.5});
  rt(s, [['Thank '], ['You', {c:C.ORANGE}], ['\n'], ['IT Solutions & Technology', {sz:28}]], 6.13, 2.909, 5.533, 1.683, {sz:66, b:1, c:C.WHITE});
}

const SLIDES = [
  slide01, slide02, slide03, slide04, slide05, slide06,
  slide07, slide08, slide09, slide10, slide11, slide12,
  slide13, slide14, slide15, slide16, slide17, slide18,
  slide19, slide20, slide21, slide22, slide23, slide24,
  slide25, slide26, slide27, slide28, slide29, slide30,
  slide31, slide32, slide33, slide34, slide35, slide36,
  slide37, slide38, slide39, slide40
];


// ---------------------------------------------------------------------------
const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'WIDE';
pptx.author = 'Sahala IT';
pptx.title = 'IT Solutions & Technology';

SLIDES.forEach(function (build) {
  const s = pptx.addSlide();
  s.background = { color: C.WHITE };
  build(s);
});

pptx.writeFile({ fileName: path.join(__dirname, '15c6d364-1076-4cc8-86d5-854e187c1b8e_grok_final.pptx') })
  .then(function (f) { console.log('wrote ' + f); });
