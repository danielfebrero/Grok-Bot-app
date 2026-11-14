/**
 * "Solar Energy" deck — recreated with pptxgenjs.
 *
 * Slide size 26.667 x 15 in (24384000 x 13716000 EMU).
 * The source deck's picture placeholders are all empty, so nothing is drawn for
 * them; the one real photograph (slide 19) becomes a flat colour placeholder,
 * and the line-art icons become simple labelled tiles.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette ---
const C = {
  white: 'FFFFFF',
  sand: 'E3DFDE',        // accent1
  coral: 'E76F51',       // accent2
  coralDark: 'CE401C',   // accent2 @ 75% luminance
  navy: '264653',        // accent3
  teal: '24A19C',        // accent4
  tealDark: '1B7975',    // accent4 @ 75% luminance
  grey: '7F7F7F',        // tx2
  body: '5F5F5F',        // tx2 @ 75% luminance — every paragraph of body copy
  cloud: 'E7E6E6',       // bg2
  cloudDark: 'D0CECE',   // bg2 @ 90% luminance
  ash: 'D9D9D9',         // bg1 @ 85% luminance
  soot: '3B3838',        // bg2 @ 25% luminance
  grey20: 'E5E5E5',
  grey40: 'CCCCCC',
  grey60: 'B2B2B2',
  axis: '595959',
  black: '000000',
};

const FONT = 'Montserrat';
const FONT_HEAD = 'Montserrat SemiBold';

// pptxgenjs preset-geometry names, keyed by the OOXML `prst` used in the source.
const GEOM = {
  rect: 'rect',
  roundRect: 'roundRect',
  round1Rect: 'round1Rect',
  round2SameRect: 'round2SameRect',
  ellipse: 'ellipse',
  triangle: 'triangle',
  flowChartDelay: 'flowChartDelay',
};

// OOXML adjust values are a percentage of the shape's shorter side; pptxgenjs
// wants that corner radius expressed in inches.
const radius = (adj, w, h) => (adj / 100000) * Math.min(w, h);

// ------------------------------------------------------------- primitives ---
/** Filled shape. `o.adj` is the OOXML corner-radius percentage. */
function box(slide, x, y, w, h, fill, o) {
  o = o || {};
  const opts = {
    x, y, w, h,
    fill: { type: 'solid', color: fill },
    line: o.line || { type: 'none' },
    flipH: o.flipH, flipV: o.flipV,
  };
  if (o.adj !== undefined) opts.rectRadius = radius(o.adj, w, h);
  slide.addShape(GEOM[o.shape] || GEOM.rect, opts);
}

/**
 * "Top corners one rounded, one snipped" with a zero snip — i.e. exactly one
 * rounded top corner, which pptxgenjs expresses as a mirrored round1Rect.
 */
function oneCorner(slide, x, y, w, h, fill, adj, flipH) {
  box(slide, x, y, w, h, fill, { shape: 'round1Rect', adj, flipH: !flipH });
}

/** Body copy: no insets, 1.5 line spacing, grey. */
function body(slide, text, o) {
  slide.addText(text, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: o.head ? FONT_HEAD : FONT, fontSize: o.size || 28,
    color: o.color || C.body, bold: o.bold,
    align: o.align || 'left', valign: 'top',
    lineSpacingMultiple: o.line === undefined ? 1.5 : o.line,
    margin: o.margin === undefined ? 0 : o.margin, wrap: true,
  });
}

/** Centred label inside a shape. */
function label(slide, text, x, y, w, h, o) {
  o = o || {};
  slide.addText(text, {
    x, y, w, h, fontFace: o.head ? FONT_HEAD : FONT, fontSize: o.size || 28,
    color: o.color || C.white, bold: o.bold, align: o.align || 'center',
    valign: o.valign || 'middle', margin: 0,
  });
}

/** Small uppercase eyebrow line that sits above every section title. */
function eyebrow(slide, text, x, y, w, align) {
  slide.addText(text, {
    x, y, w, h: 0.337, fontFace: FONT, fontSize: 20, color: C.black,
    align: align || 'left', valign: 'top', margin: 0,
  });
}

/** Two-tone 80pt headline; `br: true` ends a line. */
function headline(slide, parts, o) {
  slide.addText(parts.map(p => ({
    text: p.text,
    options: { color: p.color || C.black, breakLine: p.br, bold: true,
               fontFace: FONT_HEAD, fontSize: 80 },
  })), {
    x: o.x, y: o.y, w: o.w, h: o.h, align: o.align || 'left', valign: 'top',
    lineSpacingMultiple: 0.9, margin: [7.2, 7.2, 3.6, 3.6], wrap: true,
  });
}

/**
 * Numbered tab in the top-right corner. Slides 3-5 set the number at 32pt and
 * slightly higher; every later slide uses 28pt.
 */
function cornerBadge(slide, num, fill, big) {
  box(slide, 24.5, 0, 2.167, 2.167, fill, { shape: 'round1Rect', adj: 33013, flipH: true, flipV: true });
  slide.addText(num, {
    x: 24.5, y: big ? 0.548 : 0.815, w: 2.167, h: 0.8, fontFace: FONT,
    fontSize: big ? 32 : 28, color: C.white, align: 'center', valign: 'top', margin: 0,
  });
}

/** Rounded square holding a check mark (an SVG tick in the source). */
function checkTile(slide, x, y, size, fill, tick) {
  box(slide, x, y, size, size, fill, { shape: 'roundRect', adj: 16667 });
  label(slide, '\u2713', x, y, size, size, { size: Math.round(size * 34), color: tick });
}

/** Large decorative icon tile: one rounded top corner, "[icon]" stand-in. */
function iconPanel(slide, x, y, size, fill, flipH) {
  oneCorner(slide, x, y, size, size, fill, 16667, flipH);
  label(slide, '[icon]', x, y, size, size, { size: Math.round(size * 9) });
}

/**
 * Circle with one quadrant bitten out by a concave quarter-arc, so the four
 * petals of the slide-14 flower interlock like a pinwheel. `notch` names the
 * corner the bite curves into: 0 = TR, 1 = BR, 2 = BL, 3 = TL.
 */
function petal(slide, x, y, d, fill, notch) {
  const K = 0.5523;                                  // circle -> cubic constant
  const r = d / 2;
  const rim = [[r, 0], [d, r], [r, d], [0, r]];      // top, right, bottom, left
  const corner = [[d, 0], [d, d], [0, d], [0, 0]];   // TR, BR, BL, TL
  const points = [{ x: rim[0][0], y: rim[0][1] }];
  for (let q = 0; q < 4; q++) {
    const P = rim[q], Q = rim[(q + 1) % 4];
    const O = q === notch ? corner[q] : [r, r];      // concave arcs pivot on the corner
    points.push({ curve: { type: 'cubic',
      x1: P[0] + K * (Q[0] - O[0]), y1: P[1] + K * (Q[1] - O[1]),
      x2: Q[0] + K * (P[0] - O[0]), y2: Q[1] + K * (P[1] - O[1]) },
      x: Q[0], y: Q[1] });
  }
  points.push({ close: true });
  slide.addShape('custGeom', { x, y, w: d, h: d, points, fill: { type: 'solid', color: fill }, line: { type: 'none' } });
}

/** Rectangle standing in for a photograph in the source deck. */
function imagePlaceholder(slide, x, y, w, h, fill, o) {
  box(slide, x, y, w, h, fill, o);
  label(slide, '[image]', x, y, w, h, { size: 24 });
}

/** Short accent rule above a "Data Fact" caption. */
function rule(slide, x, y, w, color) {
  slide.addShape(GEOM.rect, { x, y, w, h: 0, line: { color, width: 3 } });
}

/** Pill button with a centred white label. */
function pill(slide, x, y, w, h, fill, text) {
  box(slide, x, y, w, h, fill, { shape: 'roundRect', adj: 24306 });
  label(slide, text, x, y, w, h);
}

// ---------------------------------------------------------------- slides ----
// Left-to-right navy -> pale teal wash on the title slide, sampled at quarter
// points. pptxgenjs has no gradient fill, so it is painted as thin bands.
const TITLE_WASH = ['264653', '447379', '629A9D', '81BCBC', 'A0D6D4'];

function slide01(pres) {
  const sl = pres.addSlide();
  const bands = 96;
  for (let i = 0; i < bands; i++) {
    sl.addShape(GEOM.rect, {
      x: (26.667 / bands) * i, y: 0, w: 26.667 / bands + 0.02, h: 15,
      fill: { type: 'solid', color: rampColor(TITLE_WASH, i / (bands - 1)) },
      line: { type: 'none' },
    });
  }

  box(sl, 22.061, 0, 4.605, 2.833, C.white, { shape: 'round1Rect', adj: 19569, flipH: true, flipV: true });
  sl.addShape(GEOM.ellipse, { x: 23.047, y: 0.896, w: 0.479, h: 0.479, fill: { type: 'none' }, line: { color: C.soot, width: 1.5 } });
  sl.addText([
    { text: 'www', options: { color: C.soot } },
    { text: '.solergy.com', options: { color: C.black } },
  ], { x: 23.047, y: 1.533, w: 3.05, h: 0.404, fontFace: FONT, fontSize: 24, margin: 0, valign: 'top' });

  rule(sl, 2.663, 2.833, 0.667, C.white);
  body(sl, 'Illuminating a Sustainable Tomorrow', { x: 2.622, y: 3.263, w: 10.508, h: 0.896, size: 40, color: C.white });
  sl.addText('Solar Energy', {
    x: 2.6, y: 5.298, w: 15.402, h: 2.794, fontFace: FONT_HEAD, fontSize: 166,
    color: C.white, margin: 0, valign: 'top',
  });
  sl.addShape(GEOM.rect, { x: 2.635, y: 9.437, w: 0.765, h: 0, line: { color: C.white, width: 3, endArrowType: 'triangle' } });
  body(sl, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere.',
    { x: 4.623, y: 8.77, w: 10.71, h: 1.334, color: C.white });
  box(sl, 2.663, 11.833, 2.002, 0.333, C.ash, { shape: 'roundRect', adj: 50000 });
  box(sl, 4.847, 11.833, 2.002, 0.333, C.coral, { shape: 'roundRect', adj: 50000 });
}

function slide02(pres) {
  const sl = pres.addSlide();
  box(sl, 18.663, 0, 6.67, 3.501, C.navy, { shape: 'round2SameRect', flipV: true });

  eyebrow(sl, 'UNVEILING SOLAR ENERGY', 1.298, 1.444, 4.002);
  headline(sl, [
    { text: 'Harnessing', br: true },
    { text: 'The Power - of The Sun', color: C.teal },
  ], { x: 1.123, y: 2.132, w: 7.027, h: 3.736 });

  iconPanel(sl, 9.998, 3.5, 2.0, C.teal, false);
  checkTile(sl, 1.329, 7.096, 0.63, C.coral, C.white);
  body(sl, "Solar energy is the utilization of sunlight to generate electricity or heat. It's a clean and renewable energy source revolutionizing the way we power our world.",
    { x: 2.799, y: 6.856, w: 7.335, h: 3.555, margin: [7.2, 7.2, 3.6, 3.6] });

  rule(sl, 2.896, 11.898, 0.708, C.teal);
  body(sl, 'Data Fact', { x: 2.868, y: 12.397, w: 1.967, h: 0.471, color: C.teal, bold: true, line: 1 });
  body(sl, 'The sun provides more energy to Earth in one hour than the entire world uses in a year.',
    { x: 5.886, y: 12.151, w: 10.558, h: 1.435, margin: [7.2, 7.2, 3.6, 3.6] });
}

function slide03(pres) {
  const sl = pres.addSlide();
  cornerBadge(sl, '02', C.navy, true);
  eyebrow(sl, 'UNVEILING SOLAR ENERGY', 11.332, 1.435, 4.002, 'center');
  headline(sl, [
    { text: 'Empowering', br: true },
    { text: 'Earth\u2019s ' }, { text: 'Applications', color: C.teal },
  ], { x: 7.467, y: 2.059, w: 11.732, h: 2.524, align: 'center' });

  oneCorner(sl, 1.33, 6.361, 5.335, 7.139, C.navy, 12235, false);
  iconPanel(sl, 2.788, 5.5, 2.0, C.coral, true);
  rule(sl, 2.794, 9.901, 0.708, C.coral);
  body(sl, [{ text: 'Lorem', options: { breakLine: true } }, { text: 'Ipsum dolor sit amet' }],
    { x: 2.756, y: 10.44, w: 2.483, h: 2.041, color: C.coral, bold: true });

  checkTile(sl, 17.102, 6.382, 0.63, C.coral, C.white);
  body(sl, [
    { text: 'Explore the various applications', options: { breakLine: true } },
    { text: 'of solar energy, from electricity generation to heating water and powering transportation.' },
  ], { x: 18.572, y: 6.097, w: 7.335, h: 2.849, margin: [7.2, 7.2, 3.6, 3.6] });

  box(sl, 18.661, 9.883, 0.781, 0.781, C.navy, { shape: 'ellipse' });
  label(sl, '\u2192', 18.661, 9.883, 0.781, 0.781, { size: 20 });
  body(sl, [
    { text: 'Solar energy ', options: { bold: true, color: C.coral, fontFace: FONT_HEAD } },
    { text: 'is used to power spacecraft, homes, and even desalinate water in remote areas.' },
  ], { x: 18.586, y: 11.54, w: 7.335, h: 2.142, margin: [7.2, 7.2, 3.6, 3.6] });
}

function slide04(pres) {
  const sl = pres.addSlide();
  cornerBadge(sl, '03', C.teal, true);
  eyebrow(sl, 'UNVEILING SOLAR ENERGY', 1.298, 1.444, 4.002);
  headline(sl, [
    { text: 'Navigating Environmental ' }, { text: 'Challenges', color: C.teal },
  ], { x: 1.123, y: 2.132, w: 9.173, h: 3.736 });
  iconPanel(sl, 14.667, 2.833, 2.0, C.coral, false);

  const cards = [
    { x: 1.329, y: 6.453, fill: C.navy, adj: 11472, flipH: false,
      tile: C.coral, tick: C.navy, ix: 4.394, iy: 7.787,
      text: 'Manufacturing Impacts', color: C.coral, tx: 2.353, tw: 4.621, ty: 8.883, th: 0.627 },
    { x: 7.998, y: 6.442, fill: C.coral, square: true,
      tile: C.navy, tick: C.coral, ix: 11.063, iy: 7.776,
      text: 'Land Use Concerns', color: C.navy, tx: 9.433, tw: 3.798, ty: 8.871, th: 0.627 },
    { x: 14.667, y: 6.442, fill: C.teal, adj: 13332, flipH: true,
      tile: C.navy, tick: C.teal, ix: 17.732, iy: 7.422,
      text: 'Electronic Waste From Decommissioned Panels', color: C.navy, tx: 15.431, tw: 5.139, ty: 8.518, th: 1.334 },
  ];
  cards.forEach(c => {
    if (c.square) box(sl, c.x, c.y, 6.668, 4.39, c.fill, { shape: 'rect' });
    else oneCorner(sl, c.x, c.y, 6.668, 4.39, c.fill, c.adj, c.flipH);
    checkTile(sl, c.ix, c.iy, 0.538, c.tile, c.tick);
    body(sl, c.text, { x: c.tx, y: c.ty, w: c.tw, h: c.th, color: c.color, align: 'center' });
  });

  rule(sl, 1.326, 11.716, 0.708, C.teal);
  body(sl, 'Data Fact', { x: 1.298, y: 12.397, w: 1.967, h: 0.471, bold: true, line: 1 });
  body(sl, 'The production of solar panels involves some toxic materials, emphasizing the importance of responsible disposal and recycling.',
    { x: 4.317, y: 12.151, w: 13.449, h: 1.435, margin: [7.2, 7.2, 3.6, 3.6] });
  pill(sl, 21.335, 12.167, 4.0, 1.333, C.navy, 'Data Fact');
}

function slide05(pres) {
  const sl = pres.addSlide();
  cornerBadge(sl, '04', C.navy, true);
  eyebrow(sl, 'UNVEILING SOLAR ENERGY', 1.32, 10.646, 4.002);
  headline(sl, [
    { text: 'Embracing Solar ' }, { text: 'Solutions', color: C.teal },
  ], { x: 1.144, y: 11.333, w: 8.876, h: 2.524 });

  iconPanel(sl, 10.878, 2.833, 2.0, C.coral, false);
  body(sl, 'Lorem ipsum', { x: 13.979, y: 2.582, w: 2.736, h: 0.627, color: C.coral, head: true });
  body(sl, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar.',
    { x: 13.86, y: 3.498, w: 8.653, h: 2.142, margin: [7.2, 7.2, 3.6, 3.6] });

  pill(sl, 18.678, 6.833, 4.0, 1.333, C.teal, 'Solar Solution');
  checkTile(sl, 17.335, 9.162, 0.63, C.coral, C.white);
  body(sl, 'Steps and initiatives needed for individuals and industries to transition towards solar energy.',
    { x: 18.527, y: 8.876, w: 7.335, h: 2.142, margin: [7.2, 7.2, 3.6, 3.6] });
  checkTile(sl, 17.322, 11.762, 0.63, C.coral, C.white);
  body(sl, 'The cost of solar panels has decreased by over 80% in the last decade, making it more accessible.',
    { x: 18.527, y: 11.476, w: 7.335, h: 2.142, margin: [7.2, 7.2, 3.6, 3.6] });
}

function slide06(pres) {
  const sl = pres.addSlide();
  cornerBadge(sl, '05', C.teal);

  eyebrow(sl, 'ADVANCEMENTS DRIVING CHANGE', 1.373, 1.444, 5.193);
  headline(sl, [
    { text: 'Innovations', br: true }, { text: 'in Solar Power', color: C.teal },
  ], { x: 1.198, y: 2.132, w: 8.869, h: 2.524 });

  pill(sl, 12.667, 1.5, 6.417, 1.333, C.teal, 'Solar Power Innovation');
  box(sl, 12.667, 6.833, 12.703, 6.667, C.coral, { shape: 'round1Rect', adj: 11212, flipH: true });
  iconPanel(sl, 22.003, 5.5, 2.0, C.navy, false);

  body(sl, 'Explore', { x: 14.957, y: 8.58, w: 1.761, h: 0.728, color: C.navy, bold: true, margin: [7.2, 7.2, 3.6, 3.6] });
  body(sl, 'Explore cutting-edge solar technologies, including photovoltaic cells, solar thermal systems, and emerging innovations.',
    { x: 14.957, y: 9.611, w: 8.379, h: 2.142, color: C.navy, margin: [7.2, 7.2, 3.6, 3.6] });

  body(sl, 'Lorem ipsum', { x: 1.325, y: 11.325, w: 2.736, h: 0.627, color: C.coral, head: true });
  body(sl, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor.',
    { x: 1.216, y: 12.241, w: 10.003, h: 1.435, margin: [7.2, 7.2, 3.6, 3.6] });
}

function slide07(pres) {
  const sl = pres.addSlide();
  cornerBadge(sl, '06', C.navy);
  eyebrow(sl, 'UNVEILING SOLAR ENERGY', 1.302, 1.431, 4.002);
  headline(sl, [
    { text: 'Financial', br: true }, { text: 'Overview', color: C.teal },
  ], { x: 1.144, y: 2.187, w: 5.828, h: 2.524 });

  checkTile(sl, 9.998, 1.5, 0.667, C.coral, C.white);
  body(sl, [
    { text: 'Lorem ipsum ', options: { color: C.coral } },
    { text: 'dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus' },
  ], { x: 11.958, y: 1.26, w: 11.337, h: 2.041 });

  // Column headers are separate rounded shapes sitting on top of the table.
  ['Region', 'Income (USD)', 'Expenditure (USD)', 'Net Balanced (USD)'].forEach((text, i) => {
    const x = 9.998 + 3.835 * i;
    box(sl, x, 4.167, 3.835, 2.0, i % 2 ? C.coralDark : C.coral, { shape: 'round2SameRect' });
    label(sl, text, x, 4.167, 3.835, 2.0, { size: 20 });
  });

  const rows = [
    ['North America', '$ 2.500.000', '$ 1.800.000', '$ 700.000'],
    ['Europe', '$1,800,000', '$1,200,000', '$600,000'],
    ['Asia \u2013 Pasific', '$3,200,000', '$2,500,000', '$700,000'],
    ['Middle East', '$1,500,000', '$1,000,000', '$500,000'],
    ['Africa', '$900,000', '$700,000', '$200,000'],
  ];
  sl.addTable(rows.map(r => r.map((text, i) => ({
    text,
    options: { fill: { color: i % 2 ? C.cloud : C.cloudDark }, align: 'center', valign: 'middle',
               fontFace: FONT, fontSize: 20, color: C.body },
  }))), {
    x: 9.998, y: 6.167, w: 15.339, colW: [3.835, 3.835, 3.835, 3.835], rowH: 1.467,
    border: { type: 'none' }, margin: 0,
  });
}

function slide08(pres) {
  const sl = pres.addSlide();
  cornerBadge(sl, '07', C.teal);
  box(sl, 0, 1.5, 6.425, 5.6, C.navy, { shape: 'round1Rect', adj: 11391 });

  eyebrow(sl, 'INTEGRATING SOLAR INTO POWER GRIDS', 16.332, 1.431, 6.039);
  headline(sl, [
    { text: 'Empowering the ' }, { text: 'Grid', color: C.teal },
  ], { x: 16.174, y: 2.187, w: 7.83, h: 2.524 });

  body(sl, 'Empowering', { x: 16.358, y: 5.244, w: 2.736, h: 0.627, color: C.coral, head: true });
  body(sl, 'Examine the integration of solar energy into existing power grids and its role in enhancing energy resilience.',
    { x: 16.24, y: 6.16, w: 9.182, h: 2.142, margin: [7.2, 7.2, 3.6, 3.6] });

  box(sl, 16.345, 9.365, 9.025, 4.135, C.coral, { shape: 'round1Rect', adj: 11212 });
  body(sl, 'Data Fact', { x: 17.681, y: 10.314, w: 1.967, h: 0.471, color: C.white, bold: true, line: 1 });
  body(sl, 'Solar power can contribute to grid stability through smart grid.',
    { x: 17.647, y: 11.301, w: 6.421, h: 1.334, color: C.white });
}

function slide09(pres) {
  const sl = pres.addSlide();
  cornerBadge(sl, '08', C.navy);
  eyebrow(sl, 'SOCIO-ECONOMIC IMPACT', 11.332, 1.435, 4.002, 'center');
  headline(sl, [
    { text: 'Empowering', br: true },
    { text: 'Local ' }, { text: 'Communities', color: C.teal },
  ], { x: 7.467, y: 2.059, w: 11.732, h: 2.524, align: 'center' });

  box(sl, 2.997, 5.292, 9.951, 4.0, C.navy, { shape: 'round1Rect' });
  box(sl, 13.733, 5.292, 9.937, 4.0, C.coral, { shape: 'round1Rect', flipH: true });
  checkTile(sl, 2.663, 5.748, 0.667, C.coral, C.white);
  checkTile(sl, 23.351, 5.748, 0.667, C.navy, C.white);

  body(sl, [
    { text: 'Solar energy ', options: { bold: true } },
    { text: 'projects can positively impact local communities, providing jobs and energy independence.' },
  ], { x: 4.288, y: 6.165, w: 7.365, h: 2.142, color: C.white, margin: [7.2, 7.2, 3.6, 3.6] });
  body(sl, [
    { text: 'Solar jobs ', options: { bold: true } },
    { text: 'have seen significant growth, employing millions globally.' },
  ], { x: 15.128, y: 6.574, w: 7.259, h: 1.435, color: C.white, margin: [7.2, 7.2, 3.6, 3.6] });
}

function slide10(pres) {
  const sl = pres.addSlide();
  cornerBadge(sl, '09', C.teal);
  eyebrow(sl, 'UNVEILING SOLAR ENERGY', 1.302, 1.431, 4.002);
  headline(sl, [
    { text: 'Paving the Way for Sustainable ' }, { text: 'Solutions', color: C.teal },
  ], { x: 1.144, y: 2.187, w: 12.856, h: 2.524 });

  body(sl, 'Strategies', { x: 1.324, y: 5.958, w: 2.736, h: 0.627, color: C.coral, head: true });
  body(sl, 'Strategies and innovations aimed at mitigating the environmental impact of solar energy, such as improved recycling methods and sustainable manufacturing practices.',
    { x: 1.206, y: 7.171, w: 9.645, h: 2.849, margin: [7.2, 7.2, 3.6, 3.6] });

  iconPanel(sl, 13.79, 6.166, 2.0, C.coral, false);
  checkTile(sl, 1.332, 11.819, 0.667, C.coral, C.white);
  body(sl, [
    { text: 'Lorem ipsum', options: { bold: true, color: C.coral } },
    { text: ' dolor sit amet, consectetuer adipiscing elit.', options: { breakLine: true } },
    { text: 'Maecenas porttitor congue massa. ' },
  ], { x: 2.939, y: 11.569, w: 6.786, h: 2.041 });
  pill(sl, 13.9, 12.165, 4.0, 1.333, C.teal, 'Strategies');
}

function slide11(pres) {
  const sl = pres.addSlide();
  eyebrow(sl, 'UNVEILING SOLAR ENERGY', 1.302, 10.096, 4.002);
  headline(sl, [
    { text: 'Regional Solar ' }, { text: 'Power Surge', color: C.teal },
  ], { x: 1.13, y: 11.083, w: 8.876, h: 2.524 });

  const kpis = [
    { x: 1.335, fill: C.navy, value: '260', fg: C.white, tile: C.white, tick: C.navy },
    { x: 6.297, fill: C.teal, value: '120', fg: C.white, tile: C.white, tick: C.teal },
    { x: 11.311, fill: C.tealDark, value: '250', fg: C.white, tile: C.white, tick: C.teal },
    { x: 16.324, fill: C.grey, value: '175', fg: C.white, tile: C.white, tick: C.grey },
    { x: 21.324, fill: C.cloud, value: '175', fg: C.body, tile: C.body, tick: C.white },
  ];
  kpis.forEach(k => {
    box(sl, k.x, 1.5, 4.0, 3.997, k.fill, { shape: 'roundRect', adj: 4854 });
    checkTile(sl, k.x + 1.725, 2.15, 0.539, k.tile, k.tick);
    label(sl, 'Your Text Here', k.x, 2.985, 4.0, 0.337, { size: 20, color: k.fg, valign: 'top' });
    label(sl, k.value, k.x, 3.504, 4.0, 1.346, { size: 80, color: k.fg, valign: 'top', head: true });
  });

  body(sl, [
    { text: 'Lorem ipsum', options: { bold: true, color: C.coral } },
    { text: ' dolor sit amet, consectetuer adipiscing elit.', options: { breakLine: true } },
    { text: 'Maecenas porttitor congue massa. ' },
  ], { x: 1.298, y: 6.701, w: 6.786, h: 2.041 });

  sl.addChart(pres.ChartType.bar, [{
    name: 'Series 1',
    labels: ['2018', '2019', '2020', '2021', '2022', '2023'],
    values: [175, 150, 260, 120, 230, 250],
  }], {
    x: 11.333, y: 6.833, w: 14.003, h: 6.917,
    barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 132,
    chartColors: [C.grey, C.sand, C.navy, C.teal, C.coral, C.tealDark],
    showLegend: false, showValue: false,
    catAxisLabelFontFace: FONT, catAxisLabelFontSize: 20, catAxisLabelColor: C.axis,
    valAxisLabelFontFace: FONT, valAxisLabelFontSize: 20, valAxisLabelColor: C.axis,
    catAxisLineColor: C.ash, valAxisLineShow: false,
    valGridLine: { color: C.ash, size: 1 }, catGridLine: { style: 'none' },
    catAxisMajorTickMark: 'none', valAxisMajorTickMark: 'none',
  });

  // marker floating over the tallest bar
  box(sl, 16.667, 7.5, 0.913, 0.913, C.coral, { shape: 'roundRect', adj: 9047 });
  box(sl, 17.019, 7.866, 0.209, 0.18, C.white, { shape: 'triangle' });

  cornerBadge(sl, '10', C.navy);  // overlaps the last KPI card
}

function slide12(pres) {
  const sl = pres.addSlide();
  cornerBadge(sl, '11', C.teal);
  eyebrow(sl, 'INTEGRATING SOLAR INTO POWER GRIDS', 1.309, 1.431, 6.039);
  headline(sl, [
    { text: 'A Solar \u2013', br: true },
    { text: 'Powered ' }, { text: 'Future', color: C.teal },
  ], { x: 1.222, y: 2.187, w: 12.252, h: 2.524 });

  iconPanel(sl, 1.33, 5.5, 1.333, C.teal, true);
  body(sl, [
    { text: 'Showcase', options: { fontFace: FONT_HEAD } },
    { text: ' global trends in solar energy adoption, highlighting regions leading the way in solar installations.' },
  ], { x: 1.298, y: 7.754, w: 6.033, h: 2.748 });

  box(sl, 10.667, 5.5, 10.002, 4.0, C.coral, { shape: 'round1Rect', adj: 11212 });
  body(sl, 'Data Fact', { x: 11.704, y: 6.237, w: 1.998, h: 0.627, color: C.white, head: true });
  body(sl, [
    { text: 'China', options: { fontFace: FONT_HEAD } },
    { text: ' and the ' },
    { text: 'United States', options: { fontFace: FONT_HEAD } },
    { text: ' are among the top countries in solar capacity.' },
  ], { x: 11.704, y: 7.262, w: 7.844, h: 1.334, color: C.white });

  checkTile(sl, 1.332, 11.754, 0.5, C.coral, C.white);
  body(sl, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue',
    { x: 2.832, y: 11.508, w: 6.5, h: 2.041 });
}

function slide13(pres) {
  const sl = pres.addSlide();
  cornerBadge(sl, '12', C.navy);
  eyebrow(sl, 'SOCIO-ECONOMIC IMPACT', 1.219, 1.435, 4.002, 'center');
  headline(sl, [
    { text: 'Empowering', br: true },
    { text: 'Local ' }, { text: 'Communities', color: C.teal },
  ], { x: 1.124, y: 2.059, w: 11.732, h: 2.524 });

  iconPanel(sl, 20.335, 2.833, 2.0, C.teal, false);

  const notes = [
    { x: 1.33, y: 5.279, ty: 6.926, w: 6.351, text: 'Positively impact local communities, providing jobs and energy independence.' },
    { x: 1.33, y: 9.96, ty: 11.613, w: 5.582, text: 'Solar jobs have seen significant growth, employing millions globally.' },
    { x: 7.848, y: 9.96, ty: 11.61, w: 6.248, text: [
        { text: 'Energy storage solutions', options: { breakLine: true } },
        { text: 'are advancing to address solar intermittency challenges.' }] },
  ];
  notes.forEach(n => {
    box(sl, n.x, n.y, 1.087, 1.087, C.navy, { shape: 'roundRect', adj: 16667 });
    label(sl, '[icon]', n.x, n.y, 1.087, 1.087, { size: 10 });
    body(sl, n.text, { x: n.x, y: n.ty, w: n.w, h: 2.041 });
  });
}

function slide14(pres) {
  const sl = pres.addSlide();
  cornerBadge(sl, '13', C.teal);
  eyebrow(sl, "SOLAR ENERGY'S IMPACT", 11.332, 1.435, 4.002, 'center');
  headline(sl, [
    { text: 'Illuminating', br: true }, { text: 'the World', color: C.teal },
  ], { x: 7.467, y: 2.059, w: 11.732, h: 2.524, align: 'center' });

  // four petals pinwheeling around the centre; each is notched by its neighbour
  [{ x: 11.554, y: 4.87, fill: C.teal, notch: 2 },
   { x: 9.712, y: 6.712, fill: C.tealDark, notch: 1 },
   { x: 11.554, y: 8.553, fill: C.coral, notch: 0 },
   { x: 13.395, y: 6.712, fill: C.coralDark, notch: 3 }]
    .forEach(p => petal(sl, p.x, p.y, 3.683, p.fill, p.notch));

  [{ n: '01', x: 11.414, y: 5.195 }, { n: '02', x: 9.336, y: 8.073 },
   { n: '03', x: 16.578, y: 8.073 }, { n: '04', x: 14.623, y: 11.072 }].forEach(b => {
    box(sl, b.x, b.y, 0.753, 0.753, C.navy, { shape: 'roundRect', adj: 11097 });
    label(sl, b.n, b.x, b.y, 0.753, 0.753, { size: 18 });
  });
  ['[icon]', '[icon]', '[icon]', '[icon]'].forEach((t, i) => {
    const p = [{ x: 12.927, y: 6.243 }, { x: 11.085, y: 8.085 },
               { x: 12.927, y: 9.926 }, { x: 14.768, y: 8.085 }][i];
    label(sl, t, p.x, p.y, 0.937, 0.937, { size: 9 });
  });

  // corner captions: line-art icons in the source, drawn here as outlined tiles
  [{ ix: 1.219, iy: 4.71, s: 1.575, tx: 1.295, ty: 6.584, color: C.teal },
   { ix: 1.248, iy: 10.25, s: 1.575, tx: 1.295, ty: 12.182, color: C.tealDark },
   { ix: 19.833, iy: 4.727, s: 1.575, tx: 19.876, ty: 6.584, color: C.coralDark },
   { ix: 19.784, iy: 10.25, s: 1.441, tx: 19.876, ty: 12.219, color: C.coral }].forEach(c => {
    sl.addShape(GEOM.roundRect, {
      x: c.ix, y: c.iy, w: c.s, h: c.s, fill: { type: 'none' },
      line: { color: c.color, width: 2 }, rectRadius: 0.12,
    });
    label(sl, '[icon]', c.ix, c.iy, c.s, c.s, { size: 12, color: c.color });
    body(sl, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ', { x: c.tx, y: c.ty, w: 5.698, h: 1.334 });
  });

  box(sl, 11.241, 13.273, 2.002, 0.227, C.navy, { shape: 'roundRect', adj: 50000 });
  box(sl, 13.424, 13.273, 2.002, 0.227, C.coral, { shape: 'roundRect', adj: 50000 });
}

function slide15(pres) {
  const sl = pres.addSlide();
  cornerBadge(sl, '14', C.navy);
  eyebrow(sl, 'JOB CREATION IN THE SOLAR INDUSTRY', 1.33, 1.435, 5.814);
  headline(sl, [
    { text: 'Empowering Careers, ' }, { text: 'Fuelling Growth', color: C.teal },
  ], { x: 1.124, y: 2.101, w: 12.695, h: 2.524 });
  iconPanel(sl, 1.334, 5.5, 2.0, C.navy, true);

  // timeline: rounded caps either side of four grey segments
  box(sl, 6.713, 7.518, 2.667, 2.667, C.coral, { shape: 'flowChartDelay', flipH: true });
  box(sl, 22.718, 7.518, 2.667, 2.667, C.teal, { shape: 'flowChartDelay' });
  label(sl, '[icon]', 7.498, 8.218, 1.266, 1.266, { size: 11 });
  label(sl, '[icon]', 23.345, 8.25, 1.268, 1.268, { size: 11 });
  [{ x: 9.38, w: 3.335, fill: C.grey20, year: '2021', fg: C.body },
   { x: 12.713, w: 3.337, fill: C.grey40, year: '2022', fg: C.body },
   { x: 16.048, w: 3.337, fill: C.grey60, year: '2023', fg: C.white },
   { x: 19.383, w: 3.335, fill: C.body, year: '2024', fg: C.white }].forEach(g => {
    box(sl, g.x, 7.518, g.w, 2.667, g.fill, { shape: 'rect' });
    label(sl, g.year, g.x, 7.518, g.w, 2.667, { size: 20, color: g.fg, head: true });
  });

  // value pins dangling above the timeline
  [{ x: 10.455, v: '75', fill: C.navy }, { x: 13.819, v: '48', fill: C.coral },
   { x: 17.184, v: '96', fill: C.navy }, { x: 20.548, v: '53', fill: C.coral }].forEach(p => {
    box(sl, p.x, 5.5, 1.184, 1.184, p.fill, { shape: 'ellipse' });
    label(sl, p.v, p.x, 5.5, 1.184, 1.184, { size: 18 });
    sl.addShape(GEOM.rect, { x: p.x + 0.592, y: 6.684, w: 0, h: 1.483, line: { color: p.fill, width: 3 } });
  });

  [1.295, 5.296, 9.296, 13.3].forEach((x, i) => {
    body(sl, String(2021 + i), { x, y: 11.294, w: 1.075, h: 0.627, head: true });
    body(sl, 'Lorem ipsum dolor sit amet. ', { x, y: 12.274, w: 2.834, h: 1.334 });
  });
  pill(sl, 17.339, 12.167, 7.998, 1.333, C.teal, 'Lorem ipsum dolor sit amet');
}

function slide16(pres) {
  const sl = pres.addSlide();
  cornerBadge(sl, '15', C.teal);
  box(sl, 1.33, 1.5, 14.67, 7.333, C.teal, { shape: 'round1Rect', adj: 8733, flipH: true });
  box(sl, 16.657, 3.5, 8.68, 5.333, C.coral, { shape: 'round1Rect', adj: 14015 });
  iconPanel(sl, 18.009, 2.173, 2.0, C.navy, false);

  body(sl, 'Discuss strategies', { x: 3.981, y: 3.103, w: 3.687, h: 0.627, color: C.white, head: true });
  body(sl, 'Discuss strategies and innovations aimed at mitigating the environmental impact of solar energy, such as improved recycling methods and sustainable manufacturing practices.',
    { x: 3.981, y: 4.381, w: 9.124, h: 2.849, color: C.white });
  body(sl, [
    { text: 'While solar ', options: { fontFace: FONT_HEAD } },
    { text: 'job growth is significant, a just transition is crucial to minimize negative impacts on existing industries.' },
  ], { x: 18.07, y: 4.987, w: 5.977, h: 2.748, color: C.white });

  eyebrow(sl, 'STRATEGIES FOR MITIGATION', 1.32, 10.323, 4.264);
  headline(sl, [
    { text: 'The Solutions', br: true }, { text: 'For Future', color: C.teal },
  ], { x: 1.144, y: 11.311, w: 8.876, h: 2.524 });

  checkTile(sl, 18.295, 11.74, 0.746, C.coral, C.white);
  body(sl, [
    { text: 'Lorem ipsum ', options: { color: C.coral } },
    { text: 'dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue' },
  ], { x: 19.823, y: 11.508, w: 6.5, h: 2.041 });
}

function slide17(pres) {
  const sl = pres.addSlide();
  cornerBadge(sl, '16', C.teal);
  eyebrow(sl, "SOLAR ENERGY'S IMPACT", 11.332, 1.435, 4.002, 'center');
  headline(sl, [
    { text: 'Investing in ', br: true }, { text: 'A Solar Future', color: C.teal },
  ], { x: 7.467, y: 2.293, w: 11.732, h: 2.524, align: 'center' });

  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Sun'];
  sl.addChart(pres.ChartType.bar, [
    { name: 'Project A', labels: months, values: [60, 25, 35, 70, 65, 35, 45] },
    { name: 'Project B', labels: months, values: [40, 75, 65, 30, 35, 65, 55] },
  ], {
    x: 1.394, y: 5.5, w: 16.258, h: 8.0,
    barDir: 'bar', barGrouping: 'stacked', barGapWidthPct: 80, barOverlapPct: 100,
    chartColors: [C.grey40, C.coral],
    showLegend: false, showValue: true, dataLabelPosition: 'ctr',
    dataLabelColor: C.white, dataLabelFontFace: FONT, dataLabelFontSize: 12,
    catAxisLabelFontFace: FONT, catAxisLabelFontSize: 20, catAxisLabelColor: C.axis,
    valAxisLabelFontFace: FONT, valAxisLabelFontSize: 12, valAxisLabelColor: C.axis,
    valAxisMaxVal: 100, catAxisLineColor: C.ash, valAxisLineShow: false,
    valGridLine: { color: C.ash, size: 1 }, catGridLine: { style: 'none' },
    catAxisMajorTickMark: 'none', valAxisMajorTickMark: 'none',
  });

  // pptxgenjs cannot recolour one point of a multi-series bar, so the teal
  // June segment is painted over the chart and its label redrawn.
  box(sl, 7.542, 6.958, 9.625, 0.584, C.teal, { shape: 'rect' });
  label(sl, '65', 7.542, 6.958, 9.625, 0.584, { size: 12 });

  [{ dot: 18.749, tx: 18.684, fill: C.coral }, { dot: 22.421, tx: 22.384, fill: C.teal }].forEach(l => {
    box(sl, l.dot, 5.774, 0.559, 0.559, l.fill, { shape: 'ellipse' });
    body(sl, [{ text: 'Lorem ipsum', options: { breakLine: true } }, { text: 'dolor sit amet.' }],
      { x: l.tx, y: 6.719, w: 3.234, h: 1.143, size: 24 });
  });

  sl.addText([
    { text: '65 ', options: { fontSize: 115, bold: true } },
    { text: '+', options: { fontSize: 66 } },
  ], { x: 18.668, y: 9.268, w: 3.384, h: 1.935, fontFace: FONT_HEAD, color: C.black, margin: 0, valign: 'top' });

  body(sl, [
    { text: 'Lorem ipsum ', options: { color: C.coral } },
    { text: 'dolor sit amet, consectetuer adipiscing elit.', options: { breakLine: true } },
    { text: 'Maecenas porttitor congue massa. ' },
  ], { x: 18.685, y: 11.459, w: 8.035, h: 2.041 });
}

function slide18(pres) {
  const sl = pres.addSlide();
  cornerBadge(sl, '17', C.teal);
  iconPanel(sl, 11.477, 5.833, 2.0, C.navy, false);

  eyebrow(sl, 'FUTURE HORIZONS', 15.327, 1.431, 6.039);
  headline(sl, [
    { text: 'Advancing Solar ' }, { text: 'Solutions', color: C.teal },
  ], { x: 15.168, y: 2.187, w: 9.502, h: 2.524 });

  body(sl, 'A Glimpse', { x: 15.308, y: 5.638, w: 2.736, h: 0.627, color: C.coral, head: true });
  body(sl, 'A glimpse into upcoming solar energy initiatives and innovations, highlighting the commitment to continuous growth and improvement.',
    { x: 15.189, y: 6.554, w: 9.758, h: 2.142, margin: [7.2, 7.2, 3.6, 3.6] });

  rule(sl, 15.34, 9.532, 0.708, C.teal);
  body(sl, 'Data Fact', { x: 15.312, y: 10.031, w: 1.967, h: 0.471, color: C.teal, bold: true, line: 1 });
  body(sl, [
    { text: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa.', options: { breakLine: true } },
    { text: 'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo.' },
  ], { x: 15.217, y: 10.823, w: 10.558, h: 2.849, margin: [7.2, 7.2, 3.6, 3.6] });
}

function slide19(pres) {
  const sl = pres.addSlide();
  cornerBadge(sl, '18', C.teal);
  box(sl, 14.905, 6.833, 10.432, 6.667, C.coral, { shape: 'round1Rect', adj: 10533 });

  eyebrow(sl, 'CONNECT WITH US', 1.419, 1.431, 6.039);
  headline(sl, [
    { text: 'Get in', br: true }, { text: 'Touch', color: C.teal },
  ], { x: 1.261, y: 2.187, w: 9.502, h: 2.524 });

  iconPanel(sl, 1.335, 5.5, 2.667, C.teal, true);
  body(sl, 'Lorem ipsum dolor sit amet, consectetuer.', { x: 1.305, y: 9.241, w: 4.668, h: 1.334, color: C.coral, bold: true });
  body(sl, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar.',
    { x: 1.33, y: 11.51, w: 8.668, h: 2.041 });

  // the hand-holding-a-phone photograph
  imagePlaceholder(sl, 9.382, 1.455, 7.49, 13.545, C.sand, { shape: 'rect' });

  box(sl, 19.599, 6.167, 4.668, 1.333, C.teal, { shape: 'roundRect', adj: 50000 });
  [20.415, 21.699, 23.008].forEach(x => sl.addShape(GEOM.ellipse, {
    x, y: 6.551, w: 0.547, h: 0.547, fill: { type: 'none' }, line: { color: C.white, width: 1.5 },
  }));

  [{ x: 16.581, y: 8.865, w: 4.002, head: 'Mail', text: 'yourmail@mail.com' },
   { x: 21.773, y: 8.865, w: 3.079, head: 'Phone', text: '+ 123 456 789 0' },
   { x: 16.581, y: 10.938, w: 7.802, head: 'Address', text: '123 Solar Street, Your City Name, Your Country' }]
    .forEach(c => body(sl, [
      { text: c.head, options: { fontFace: FONT_HEAD, breakLine: true } },
      { text: c.text },
    ], { x: c.x, y: c.y, w: c.w, h: 1.143, size: 24, color: C.white }));
}

function slide20(pres) {
  const sl = pres.addSlide();
  cornerBadge(sl, '19', C.teal);
  eyebrow(sl, 'SOCIO-ECONOMIC IMPACT', 1.219, 1.435, 4.002, 'center');
  headline(sl, [
    { text: 'Thank You', br: true },
    { text: 'For ' }, { text: 'Joining Us', color: C.teal },
  ], { x: 1.224, y: 2.453, w: 11.732, h: 2.524 });

  box(sl, 12.667, 1.5, 10.714, 7.426, C.ash, { shape: 'round1Rect', adj: 9164, flipH: true, line: { color: C.white, width: 14 } });
  iconPanel(sl, 1.335, 8.167, 2.667, C.coral, true);

  box(sl, 14.667, 8.188, 10.665, 5.333, C.teal, { shape: 'round1Rect' });
  body(sl, [
    { text: 'Thank you! ', options: { fontFace: FONT_HEAD, breakLine: true } },
    { text: 'For the opportunity to share ideas', options: { breakLine: true } },
    { text: 'and information with all of you. ', options: { breakLine: true } },
    { text: 'I hope it proves beneficial to us all.' },
  ], { x: 16.632, y: 9.48, w: 6.734, h: 2.748, color: C.white });
}

/** Colour at position `t` (0..1) along an evenly spaced list of hex stops. */
function rampColor(stops, t) {
  const pos = t * (stops.length - 1);
  const i = Math.min(stops.length - 2, Math.floor(pos));
  const f = pos - i;
  const ch = (h, k) => parseInt(h.substr(k * 2, 2), 16);
  return [0, 1, 2].map(k => Math.round(ch(stops[i], k) + (ch(stops[i + 1], k) - ch(stops[i], k)) * f)
    .toString(16).padStart(2, '0')).join('').toUpperCase();
}

// ----------------------------------------------------------------- build ----
function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'SOLAR', width: 26.6667, height: 15 });
  pres.layout = 'SOLAR';
  pres.theme = { headFontFace: FONT_HEAD, bodyFontFace: FONT };
  pres.title = 'Solar Energy';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
   slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach(fn => fn(pres));

  return pres.writeFile({
    fileName: path.join(__dirname, '056d2821-6550-4640-8b61-05ee13cee16e_grok_final.pptx'),
  });
}

build().then(f => console.log('wrote', f)).catch(e => { console.error(e); process.exit(1); });
