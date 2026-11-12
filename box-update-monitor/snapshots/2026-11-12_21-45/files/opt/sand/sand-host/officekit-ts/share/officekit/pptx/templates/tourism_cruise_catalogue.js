#!/usr/bin/env node
/**
 * "KUYLIN — Travel Catalogue Presentation" (21 slides, 13.333in x 7.5in).
 * Rebuilt from scratch with pptxgenjs; every position/colour/font below is a plain literal.
 *
 * The source deck contains empty picture placeholders (no image data is stored in the
 * file). They render as nothing, so they are not drawn here.
 */

const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ palette */
const YELLOW = 'FEDE3C'; // primary brand yellow
const YELLOW_SOFT = 'FEE35D'; // lighter yellow panels
const YELLOW_DEEP = 'FED50A'; // closing-slide squares
const GOLD = 'FFCC00'; // timeline circles / dashes
const CHARCOAL = '3D3D3D'; // dark banners + most body copy
const SLATE = '3F3F3F'; // headings on light panels
const INK = '141414'; // near-black accents + photo scrims
const GRAY = '595959'; // chart bars, rules, badges
const SILVER = 'A5A5A5'; // chart bars (2nd series)
const BLACK = '000000';
const WHITE = 'FFFFFF';

/* -------------------------------------------------------------------- fonts */
const HEAD = 'Work Sans'; // display headings
const HEAD_MED = 'Work Sans Medium'; // title-slide kicker
const BODY = 'Roboto'; // body copy
const ALT = 'Poppins'; // credits + chart labels

const NO_LINE = { type: 'none' };

/* ------------------------------------------------------------------ helpers */

/** Flat colour rectangle. `opts.transparency` (0-100) and `opts.rotate` are optional. */
function rect(slide, x, y, w, h, color, opts = {}) {
  const { transparency, ...rest } = opts;
  slide.addShape('rect', { x, y, w, h, fill: { color, transparency }, line: NO_LINE, ...rest });
}

/** Filled disc. */
function disc(slide, x, y, d, color) {
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color }, line: NO_LINE });
}

/** Unfilled ring. */
function ring(slide, x, y, d, color, width) {
  slide.addShape('ellipse', { x, y, w: d, h: d, line: { color, width } });
}

/** Horizontal rule (thin straight connector). */
function rule(slide, x, y, w, color, width, dashType) {
  slide.addShape('line', { x, y, w, h: 0, line: { color, width, dashType: dashType || 'solid' } });
}

/**
 * Turn a list of paragraphs into a pptxgenjs run array.
 * Each entry is a plain string, or `{ t: 'text', ...paragraphOptions }`.
 */
function paras(list) {
  return list.map((p, i) => {
    const o = typeof p === 'string' ? { t: p } : p;
    const { t, ...paraOpts } = o;
    return { text: t, options: { breakLine: i < list.length - 1, ...paraOpts } };
  });
}

/** Big display heading (Work Sans bold). */
function heading(slide, lines, o) {
  slide.addText(paras(lines), {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: o.font || HEAD, fontSize: o.size || 40, bold: true,
    color: o.color || CHARCOAL, align: o.align || 'left', valign: 'top',
    lineSpacingMultiple: o.lineSpacing, rotate: o.rotate,
  });
}

/** Small body copy (Roboto, 1.5 line spacing by default). */
function body(slide, lines, o) {
  slide.addText(paras(lines), {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: o.font || BODY, fontSize: o.size || 8, bold: o.bold || false,
    color: o.color || CHARCOAL, align: o.align || 'left', valign: o.valign || 'top',
    lineSpacingMultiple: o.lineSpacing === undefined ? 1.5 : o.lineSpacing,
    rotate: o.rotate,
  });
}

/** The recurring "LET'S / KUY / OUT." wordmark. */
function wordmark(slide, x, y, o = {}) {
  slide.addText(paras(o.lines || ['LET’S', 'KUY', 'OUT.']), {
    x, y, w: o.w || 1.392, h: 0.808,
    fontFace: HEAD, fontSize: 14, bold: true,
    color: o.color || SLATE, align: o.align || 'left', valign: 'top',
  });
}

/** The recurring "KuyLin Creative Design" credit line. */
function credit(slide, o) {
  slide.addText('KuyLin Creative Design', {
    x: o.x, y: o.y, w: o.w || 3.579, h: o.h || 0.306,
    fontFace: o.font || ALT, fontSize: o.size || 9, bold: o.bold !== false,
    color: o.color || CHARCOAL, align: o.align || 'right', valign: 'top',
    lineSpacingMultiple: 1.5, rotate: o.rotate,
  });
}

/**
 * The white "corner frame" motif used on the opening and closing slides:
 * an L-shaped freeform — a horizontal arm across the top plus a vertical arm
 * down one side.
 */
function cornerFrame(slide, o) {
  const { w, h, armW, armH } = o;
  const inner = o.side === 'right' ? w - armW : armW;
  const outer = o.side === 'right' ? w : 0;
  slide.addShape('custGeom', {
    x: o.x, y: o.y, w, h, fill: { color: WHITE }, line: NO_LINE,
    points: [
      { x: outer, y: 0 },
      { x: w - outer, y: 0 },
      { x: w - outer, y: armH },
      { x: inner, y: armH },
      { x: inner, y: h },
      { x: outer, y: h },
      { close: true },
    ],
  });
}

/* ------------------------------------------------------------------- slides */

// 1 — title slide: photo scrim, yellow strap, corner frame
function slide01(pptx) {
  const s = pptx.addSlide();
  rect(s, 0, 0, 11.343, 6.381, INK, { transparency: 39 }); // scrim over the (absent) hero photo
  rect(s, 5.958, 5.956, 7.375, 1.544, YELLOW);
  s.addText('TRAVEL CATALOGUE PRESENTATION', {
    x: 7.451, y: 6.706, w: 5.265, h: 0.438,
    fontFace: HEAD_MED, fontSize: 20, color: WHITE, align: 'center', valign: 'top',
  });
  heading(s, ['KUYLIN'], { x: 6.161, y: 4.703, w: 4.382, h: 1.313, size: 72, color: WHITE });
  cornerFrame(s, { x: 5.0065, y: 4.3082, w: 7.7674, h: 2.1033, armH: 0.1861, armW: 0.2129, side: 'left' });
}

// 2 — agenda: yellow header band + three vertical chapter bars
function slide02(pptx) {
  const s = pptx.addSlide();
  rect(s, 0, 0, 13.333, 2.775, YELLOW);
  heading(s, ['CONTENT'], { x: 3.478, y: 0.399, w: 6.378, h: 0.998, color: SLATE, align: 'center', lineSpacing: 1.5 });
  body(s, ["Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it"],
    { x: 3.478, y: 1.51, w: 6.378, h: 0.577, size: 10, color: SLATE, align: 'center' });

  const chapters = [
    { label: 'INTRODUCTION', barX: 2.284, textX: 2.788, textY: 5.419 },
    { label: 'MAIN SLIDE', barX: 6.492, textX: 7.034, textY: 5.482 },
    { label: 'THE LAST', barX: 10.686, textX: 11.229, textY: 5.482 },
  ];
  chapters.forEach(c => {
    rect(s, c.barX, 4.939, 3.478, 0.527, CHARCOAL, { rotate: -90 });
    s.addText(c.label, {
      x: c.textX, y: c.textY, w: 2.483, h: 0.438, rotate: -90,
      fontFace: BODY, fontSize: 20, bold: true, color: WHITE, align: 'left', valign: 'top',
    });
  });
}

// 3 — "This summer vibes"
function slide03(pptx) {
  const s = pptx.addSlide();
  heading(s, ['THIS SUMMER VIBES'], { x: 7.289, y: 2.018, w: 4.856, h: 1.447, color: INK, align: 'right' });
  rect(s, 1.808, 0.981, 4.856, 5.55, YELLOW_SOFT);
  body(s, [
    'Gulf emirates such as Dubai and Abu Dhabi are working hard to bring in world-class attractions, events, performances and cultural activities for 2021.',
    { t: 'With the Expo 2020 Dubai postponed for the winter of 2021, there is more reason to spend time living it up in the Gulf. ', paraSpaceBefore: 10 },
    { t: '.', paraSpaceBefore: 10 },
  ], { x: 8.354, y: 3.767, w: 3.791, h: 1.772, color: INK, align: 'right' });
  wordmark(s, 0.197, 0.394);
  credit(s, { x: 11.2, y: 5.387, rotate: 90 });
}

// 4 / 9 / 13 — section dividers: dark banner bottom-left
function sectionDivider(pptx, o) {
  const s = pptx.addSlide();
  rect(s, 0, 5.522, 7.958, 1.978, CHARCOAL);
  body(s, [o.copy], { x: o.x, y: o.copyY, w: o.copyW, h: o.copyH, size: o.copySize, color: WHITE });
  heading(s, [o.title], { x: o.x, y: o.titleY, w: o.titleW, h: 0.708, color: WHITE, lineSpacing: 0.9 });
  return s;
}

function slide04(pptx) {
  const s = sectionDivider(pptx, {
    title: 'INTRODUCTION', x: 0.322, titleY: 5.733, titleW: 5.26,
    copy: "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer.",
    copyY: 6.517, copyW: 5.067, copyH: 0.577, copySize: 10,
  });
  credit(s, { x: 0.188, y: 0.153, align: 'justify' });
}

function slide09(pptx) {
  sectionDivider(pptx, {
    title: 'MAIN SLIDE', x: 0.418, titleY: 5.806, titleW: 4.46,
    copy: "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book.",
    copyY: 6.556, copyW: 6.844, copyH: 0.482, copySize: 8,
  });
}

function slide13(pptx) {
  sectionDivider(pptx, {
    title: 'THE LAST', x: 0.557, titleY: 5.783, titleW: 4.46,
    copy: "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book.",
    copyY: 6.527, copyW: 5.679, copyH: 0.684, copySize: 8,
  });
}

// 5 — MSC Virtuosa, yellow panel right
function slide05(pptx) {
  const s = pptx.addSlide();
  rect(s, 7.931, 1.053, 4.625, 5.28, YELLOW);
  heading(s, [' MSC VIRTUOSA'], { x: 8.38, y: 1.715, w: 3.685, h: 1.447, color: SLATE, align: 'right', lineSpacing: 1.0 });
  body(s, ['The generated Lorem Ipsum is therefore.'],
    { x: 8.38, y: 3.507, w: 3.685, h: 0.325, size: 10, bold: true, color: SLATE, align: 'right' });
  body(s, ["Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book."],
    { x: 8.38, y: 4.015, w: 3.685, h: 0.886, color: SLATE, align: 'right' });
  credit(s, { x: 9.517, y: 6.996, h: 0.353 });
  wordmark(s, 0.22, 0.123, { color: WHITE });
}

// 6 — Cruise the Mediterranean, full yellow block
function slide06(pptx) {
  const s = pptx.addSlide();
  rect(s, 0, -0.006, 8.833, 5.256, YELLOW);
  heading(s, ['CRUISE THE', 'MEDITERRANEAN'], { x: 1.433, y: 0.966, w: 5.522, h: 1.447, color: SLATE, lineSpacing: 1.0 });
  body(s, ['MSC Cruises’ newest ships, launching in April 2021, will spend her first winter season in the region, homeporting in Dubai. From there, the ship will sail to the nearby modern metropolis of Abu Dhabi, on to Sir Bani Yas Island, where guests will have a chance to discover an exotic array of wildlife or bask in the warm sun on the immaculate beaches.'],
    { x: 1.433, y: 2.744, w: 5.747, h: 1.088, color: SLATE });
  wordmark(s, 0.349, 6.442);
  credit(s, { x: 12.005, y: 1.019, w: 2.145, rotate: 90, align: 'left' });
}

// 7 — Grandiosa
function slide07(pptx) {
  const s = pptx.addSlide();
  rect(s, 2.12, 1.477, 3.974, 2.554, YELLOW);
  rect(s, 8.937, 0, 4.397, 0.933, YELLOW);
  heading(s, ['GRANDIOSA'], { x: 6.93, y: 2.201, w: 3.715, h: 0.707, lineSpacing: 0.9 });
  body(s, ['PLACEHOLDER'],
    { x: 6.93, y: 3.198, w: 4.504, h: 0.578, size: 10, bold: true, color: INK });
  body(s, [
    'will call in some of Europe’s most popular destinations. With its numerous palaces and winding streets, Genoa will serve as a key embarkation port for guests.',
    { t: 'As well as the artistic and gastronomic wonderland of Barcelona and key to the South of France, Marseille. In addition to these gems, the ship will call in the beautiful Palermo.', paraSpaceBefore: 10 },
  ], { x: 6.93, y: 4.065, w: 4.07, h: 1.301, size: 7 });
  wordmark(s, 0.196, 0.176);
  credit(s, { x: 11.304, y: 5.412, rotate: 90 });
}

// 8 — Virginia, rotated display word on the left
function slide08(pptx) {
  const s = pptx.addSlide();
  heading(s, ['VIRGINIA'], { x: -0.41, y: 3.377, w: 2.993, h: 0.707, align: 'center', rotate: -90, lineSpacing: 0.9 });
  body(s, ['It is a long established fact that a reader will be distracted.'],
    { x: 8.128, y: 3.081, w: 4.0, h: 0.578, size: 10, bold: true, color: INK });
  body(s, [
    'Richard McClintock, a Latin professor at Hampden-Sydney College in Virginia, looked up one of the more obscure Latin words, consectetur.From a Lorem Ipsum passage, and going through the ',
    { t: 'cites of the word in classical literature discovered the undoubtable source. Lorem Ipsum comes from sections 1.10.32 and 1.10.33 of "de Finibus Bonorum et Malorum by Cicero.', paraSpaceBefore: 10 },
  ], { x: 8.128, y: 3.822, w: 4.073, h: 1.857 });
  rect(s, 8.128, 6.889, 5.206, 0.611, YELLOW);
  s.addText('SUBTITLE EXAMPLE', {
    x: 8.128, y: 2.481, w: 3.203, h: 0.438,
    fontFace: BODY, fontSize: 20, bold: true, color: CHARCOAL, align: 'left', valign: 'top',
  });
  wordmark(s, 7.544, 0.128);
  credit(s, { x: 9.632, y: 7.032 });
}

// 10 — Preziosa
function slide10(pptx) {
  const s = pptx.addSlide();
  rect(s, 0, 0, 2.667, 1.9, YELLOW);
  body(s, [
    'MSC Cruises is bringing in four modern ships with embarkation in Brazil and a fifth ship in the region with embarkations in Argentina. MSC Seaside.',
    { t: ' MSC Preziosa, MSC Splendida, MSC Sinfonia, and MSC Orchestra will cover Brazil, Uruguay and Argentina with itineraries ranging from 3 to 10 nights.', paraSpaceBefore: 10 },
    { t: '2021 is already being chalked out as an exceptional year, with substantial pent up demand for travel, exploration, experimentation and above all new, refreshing experiences', paraSpaceBefore: 10 },
  ], { x: 7.244, y: 2.803, w: 4.622, h: 2.199, align: 'right' });
  heading(s, ['PREZIOSA'], { x: 7.783, y: 1.405, w: 4.083, h: 0.774, align: 'right' });
  rect(s, 7.867, 7.022, 5.467, 0.49, YELLOW);
  wordmark(s, 0.243, 0.142);
  credit(s, { x: 9.556, y: 7.139 });
}

// 11 — The Simply Cruise
function slide11(pptx) {
  const s = pptx.addSlide();
  rect(s, 9.956, 0, 2.921, 2.585, YELLOW);
  heading(s, ['THE ', 'SIMPLY CRUISE'], { x: 10.0, y: 0.225, w: 2.714, h: 2.121, color: SLATE, align: 'right' });
  heading(s, ['MSC DIVINA'], { x: 1.252, y: 3.625, w: 3.296, h: 0.438, size: 20, align: 'right' });
  body(s, [
    'But also the leap into electronic typesetting, remaining essentially unchanged. It has survived not only five centuries.',
    '',
    'It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. ',
  ], { x: 1.489, y: 4.24, w: 3.059, h: 1.515, align: 'right' });
  wordmark(s, 8.672, 4.062);
  credit(s, { x: -1.493, y: 5.238, h: 0.353, rotate: 90 });
}

// 12 — Magnifica
function slide12(pptx) {
  const s = pptx.addSlide();
  rect(s, -0.011, 3.692, 3.74, 3.792, YELLOW);
  heading(s, ['MAGNIFICA'], { x: 9.367, y: 1.398, w: 3.568, h: 1.111, align: 'right', lineSpacing: 1.5 });
  body(s, ['It was popularised in the 1960s'],
    { x: 10.25, y: 2.64, w: 2.604, h: 0.353, size: 10, bold: true, color: INK });
  body(s, [
    "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type. ",
    { t: 'It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged.', paraSpaceBefore: 10 },
  ], { x: 9.917, y: 2.994, w: 2.983, h: 1.632, color: BLACK, align: 'right' });
  body(s, [
    'Contrary to popular belief, Lorem Ipsum is not simply random text. ',
    '',
    'It has roots in a piece of classical Latin literature from 45 BC, making it over 2000 years old. ',
    '',
  ], { x: 0.437, y: 5.204, w: 2.617, h: 1.691, size: 9, color: SLATE });
  wordmark(s, 0.132, 0.147);
  credit(s, { x: 9.361, y: 7.007, bold: false });
}

// 14 — "The years": three statistic columns on a yellow slab
function slide14(pptx) {
  const s = pptx.addSlide();
  rect(s, 1.415, 5.084, 7.089, 1.943, YELLOW);
  heading(s, ['THE YEARS'], { x: 1.415, y: 1.336, w: 3.561, h: 1.111, lineSpacing: 1.5 });
  heading(s, ['TYPE YOUR TEXT'], { x: 1.544, y: 2.616, w: 2.637, h: 0.55, size: 20, font: BODY, lineSpacing: 1.5 });
  body(s, [
    'Various versions have evolved over the years, sometimes by accident, sometimes on purpose.',
    '',
    'The generated Lorem',
    'Ipsum is therefore',
  ], { x: 1.544, y: 3.282, w: 3.016, h: 1.464, size: 9 });

  const stats = [
    { pct: '7%', pctX: 1.591, pctW: 0.598, ruleX: 1.678, capX: 1.591 },
    { pct: '30%', pctX: 3.962, pctW: 0.773, ruleX: 4.05, capX: 3.962 },
    { pct: '25%', pctX: 6.434, pctW: 0.773, ruleX: 6.439, capX: 6.352 },
  ];
  stats.forEach(st => {
    heading(s, [st.pct], { x: st.pctX, y: 5.359, w: st.pctW, h: 0.55, size: 20, font: BODY, color: SLATE, lineSpacing: 1.5 });
    rule(s, st.ruleX, 6.029, 1.498, GRAY, 1.5);
    body(s, ['The generated Lorem', 'Ipsum is therefore'],
      { x: st.capX, y: 6.108, w: 1.668, h: 0.53, size: 9, color: SLATE });
  });

  wordmark(s, 0.334, 0.106, { w: 1.345 });
  credit(s, { x: 9.121, y: 7.03, bold: false });
}

// 15 — Hampden Sydney / History
function slide15(pptx) {
  const s = pptx.addSlide();
  rect(s, 0.015, -0.016, 3.159, 7.516, YELLOW_SOFT);
  heading(s, ['HAMPDEN', 'SYDNEY'], { x: 9.608, y: 1.044, w: 3.354, h: 1.447, align: 'right' });
  body(s, [
    'Richard McClintock, a Latin professor at Hampden-Sydney College in Virginia, looked up one of the more obscure Latin words, consectetur.',
    '',
    'From a Lorem Ipsum passage, and going through the cites of the word in classical literature, discovered the undoubtable source. ',
  ], { x: 9.101, y: 3.079, w: 3.861, h: 2.121, size: 10, align: 'right' });
  wordmark(s, 0.257, 0.099);
  credit(s, { x: -0.994, y: 5.997, w: 2.501, rotate: 90, align: 'justify' });
  rect(s, 5.151, 5.801, 3.083, 0.939, YELLOW_SOFT);
  heading(s, ['HISTORY'], { x: 5.43, y: 5.951, w: 2.525, h: 0.64, size: 32, align: 'center' });
}

// 16 — Travel market: hand-built bar chart + yellow copy panel
function slide16(pptx) {
  const s = pptx.addSlide();
  rect(s, 7.856, 0, 5.478, 7.5, YELLOW);
  heading(s, ['TRAVEL', 'MARKET'], { x: 8.652, y: 0.378, w: 3.354, h: 2.008, color: SLATE, lineSpacing: 1.5 });

  const notes = [
    { letter: 'A', badgeY: 2.657, textY: 3.44, textW: 3.083, textH: 0.783 },
    { letter: 'B', badgeY: 4.4, textY: 5.068, textW: 2.348, textH: 1.01 },
  ];
  notes.forEach(n => {
    s.addText(n.letter, {
      shape: 'ellipse', x: 8.726, y: n.badgeY, w: 0.636, h: 0.636,
      fill: { color: GRAY },
      fontFace: ALT, fontSize: 14, color: WHITE, align: 'center', valign: 'middle',
    });
    body(s, ['He standard chunk of Lorem Ipsum used since the 1500s is reproduced below for those interested. '],
      { x: 8.652, y: n.textY, w: n.textW, h: n.textH, size: 9, font: ALT, color: SLATE, align: 'justify' });
  });

  // Grouped bar chart, drawn as rectangles (as in the source deck).
  // Per bar: [x, top, height]; series colour comes from the column index.
  const BAR_W = 0.214;
  const SERIES_COLORS = [GRAY, SILVER, YELLOW];
  const barGroups = [
    [[1.596, 4.462, 2.059], [1.864, 4.717, 1.804], [2.133, 4.207, 2.315]], // A
    [[2.805, 5.224, 1.305], [3.074, 5.773, 0.756], [3.343, 5.495, 1.034]], // B
    [[4.014, 4.978, 1.544], [4.289, 4.717, 1.804], [4.558, 5.765, 0.756]], // C
    [[5.232, 4.978, 1.544], [5.508, 3.953, 2.568], [5.776, 4.462, 2.059]], // D
  ];
  barGroups.forEach(group => {
    group.forEach(([x, y, h], i) => rect(s, x, y, BAR_W, h, SERIES_COLORS[i]));
  });

  const axisLabel = (text, x, y) => s.addText(text, {
    x, y, w: 0.819, h: 0.303,
    fontFace: ALT, fontSize: 12, color: GRAY, align: 'center', valign: 'top',
  });
  [['60', 3.294], ['50', 3.803], ['40', 4.312], ['30', 4.822], ['20', 5.331], ['10', 5.84], ['0', 6.349]]
    .forEach(([t, y]) => axisLabel(t, 0.695, y));
  [['A', 1.514], ['B', 2.744], ['C', 3.974], ['D', 5.205]]
    .forEach(([t, x]) => axisLabel(t, x, 6.652));

  wordmark(s, 0.408, 0.378, { w: 1.064 });
  credit(s, { x: 11.224, y: 3.931, h: 0.353, rotate: 90, align: 'center', font: BODY, size: 10, bold: false, color: INK });
}

/** Shared timeline geometry for slides 17 & 18. */
const TIMELINE_Y = 3.75;
function timelineStop(slide, o) {
  disc(slide, o.x, o.y, 1.542, GOLD);
  slide.addText(o.month, {
    x: o.labelX, y: o.labelY, w: o.labelW, h: 0.438,
    fontFace: BODY, fontSize: 20, color: o.labelColor, align: 'center', valign: 'top',
  });
  slide.addText('SUBTITTLE', {
    x: o.headX, y: 2.186, w: 1.932, h: 0.55,
    fontFace: BODY, fontSize: 20, bold: true, color: CHARCOAL,
    align: 'center', valign: 'top', lineSpacingMultiple: 1.5,
  });
  body(slide, ['The generated Lorem Ipsum is therefore always free from repetition.'],
    { x: o.capX, y: 4.764, w: 1.948, h: 0.684, align: 'center' });
}

// 17 — Schedule: START ring + JAN/FEB/MAR
function slide17(pptx) {
  const s = pptx.addSlide();
  rect(s, 0, 0.466, 4.126, 1.233, YELLOW);
  heading(s, ['SCHEDULE'], { x: 0.488, y: 0.696, w: 3.149, h: 0.774, font: BODY, color: SLATE, align: 'center' });

  [[4.012, 1.249], [6.804, 1.23], [9.587, 1.23], [12.353, 0.98]]
    .forEach(([x, w]) => rule(s, x, TIMELINE_Y, w, GOLD, 2.25, 'dash'));

  ring(s, 1.09, 2.289, 2.922, GOLD, 3);
  heading(s, ['START'], { x: 1.09, y: 3.363, w: 2.922, h: 0.774, font: BODY, align: 'center' });

  [
    { month: 'JAN', x: 5.262, y: 2.979, labelX: 5.324, labelY: 3.559, labelW: 1.386, headX: 5.051, capX: 5.113 },
    { month: 'FEB', x: 8.045, y: 2.968, labelX: 8.11, labelY: 3.534, labelW: 1.384, headX: 7.85, capX: 7.991 },
    { month: 'MAR', x: 10.811, y: 2.979, labelX: 10.807, labelY: 3.545, labelW: 1.551, headX: 10.616, capX: 10.6 },
  ].forEach(stop => timelineStop(s, { ...stop, labelColor: SLATE }));

  credit(s, { x: -1.449, y: 5.56, h: 0.353, rotate: 90, align: 'center', font: BODY, size: 10, bold: false, color: INK });
  wordmark(s, 11.756, 0.151, { align: 'right' });
}

// 18 — Schedule continued: APR/MAY/JUN + END ring
function slide18(pptx) {
  const s = pptx.addSlide();
  [[0, 1.281], [2.823, 1.133], [5.499, 1.121], [8.162, 1.201]]
    .forEach(([x, w]) => rule(s, x, TIMELINE_Y, w, GOLD, 2.25, 'dash'));

  ring(s, 9.364, 2.397, 2.705, GOLD, 3);
  heading(s, ['END'], { x: 9.602, y: 3.397, w: 2.228, h: 0.774, font: BODY, align: 'center' });

  [
    { month: 'APR', x: 1.281, y: 2.979, labelX: 1.344, labelY: 3.531, labelW: 1.386, headX: 1.15, capX: 1.078 },
    { month: 'MAY', x: 3.956, y: 2.979, labelX: 4.022, labelY: 3.531, labelW: 1.384, headX: 3.748, capX: 3.739 },
    { month: 'JUN', x: 6.62, y: 2.979, labelX: 6.616, labelY: 3.531, labelW: 1.551, headX: 6.425, capX: 6.525 },
  ].forEach(stop => timelineStop(s, { ...stop, labelColor: WHITE }));

  wordmark(s, 11.756, 0.151, { align: 'right' });
  credit(s, { x: 0.247, y: 7.037, h: 0.325, align: 'left', font: BODY, size: 10, bold: false, color: SLATE });
}

// 19 — Glamourous MSC Fantasia
function slide19(pptx) {
  const s = pptx.addSlide();
  rect(s, 0.213, 4.681, 12.568, 2.14, YELLOW);
  heading(s, ['GLAMOUROUS MSC FANTASIA '], { x: 6.937, y: 1.415, w: 5.47, h: 2.121, align: 'right', lineSpacing: 1.5 });
  body(s, ["Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book."],
    { x: 7.075, y: 5.322, w: 5.331, h: 0.858, size: 10, color: WHITE, align: 'right' });
  wordmark(s, 11.8, 0.162, { w: 1.214, align: 'right', lines: ['LET’S KUY', 'OUT.'] });
  credit(s, { x: -1.499, y: 2.313, h: 0.325, rotate: 90, align: 'left', font: BODY, size: 10, bold: false, color: SLATE });
}

// 20 — Last to all
function slide20(pptx) {
  const s = pptx.addSlide();
  rect(s, 9.148, 3.991, 3.173, 2.511, YELLOW_DEEP);
  rect(s, 7.478, 0.998, 3.173, 2.511, YELLOW_DEEP);
  rect(s, 0, 0, 6.667, 7.5, YELLOW);
  heading(s, ['LAST TO ALL'], { x: 1.013, y: 1.584, w: 6.21, h: 0.998, color: SLATE, lineSpacing: 1.5 });
  body(s, [
    'Contrary to popular belief, Lorem Ipsum is not simply random text. It has roots in a piece of classical Latin literature from 45 BC, making it over 2000 years old. Richard McClintock, a Latin professor at Hampden-Sydney College in Virginia,Looked ',
    { t: '(The Extremes of Good and Evil) by Cicero, written in 45 BC. This book is a treatise on the theory of ethics, very popular during the Renaissance.The first line of Lorem Ipsum, "Lorem ipsum dolor sit amet..", comes from a line in section 1.10.32.', paraSpaceBefore: 10 },
  ], { x: 1.013, y: 3.033, w: 4.499, h: 1.834, color: SLATE });
  wordmark(s, 0.199, 0.07);
  credit(s, { x: 11.318, y: 1.697, h: 0.325, rotate: 90, align: 'left', font: BODY, size: 10, bold: false, color: SLATE });
}

// 21 — Thank you
function slide21(pptx) {
  const s = pptx.addSlide();
  rect(s, 2.233, 0, 11.101, 6.212, INK, { transparency: 29 }); // scrim over the (absent) photo
  rect(s, 0, 5.956, 7.022, 1.544, YELLOW);
  heading(s, ['THANK’S'], { x: 2.295, y: 4.727, w: 5.488, h: 1.313, size: 72, color: WHITE, align: 'center' });
  heading(s, ['A BUNCH'], { x: 0.767, y: 6.293, w: 5.488, h: 1.01, size: 54, color: WHITE, align: 'center' });
  cornerFrame(s, { x: 2.2326, y: 4.1085, w: 6.0221, h: 2.1033, armH: 0.1861, armW: 0.1651, side: 'right' });
}

/* --------------------------------------------------------------------- main */
function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'KUYLIN', width: 13.3333333, height: 7.5 });
  pptx.layout = 'KUYLIN';
  pptx.title = 'KUYLIN — Travel Catalogue Presentation';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
    slide21].forEach(fn => fn(pptx));

  return pptx.writeFile({
    fileName: path.join(__dirname, '157fc674-6029-4250-9b7e-cc9a3d05c87b_grok_final.pptx'),
  });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
