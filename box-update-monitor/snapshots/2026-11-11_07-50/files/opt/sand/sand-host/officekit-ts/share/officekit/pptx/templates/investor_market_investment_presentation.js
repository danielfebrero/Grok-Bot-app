/**
 * Recreates "Marketing Investment" (20 slides, 13.333 x 7.5 in) with pptxgenjs.
 * Raster photos in the source deck are drawn here as labelled placeholder panels.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

const OUT = path.join(__dirname, '0a2b86b5-ac24-4f67-8690-7a04991facc4_grok_final.pptx');

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

const F = { head: 'Montserrat', body: 'Lato' };

const C = {
  black: '161616',   // accent1
  ink: '000000',     // tx1
  dark: '141414',
  head: '101010',    // eyebrow headings
  body: '404040',    // tx1 75%
  soft: '262626',
  blue: '2D5FFF',    // accent2
  blueDeep: '0036E1',// accent2 75%
  gray: 'A5A5A5',    // accent3 75%
  grayMid: '8A8A8A',
  grayDeep: '6E6E6E',
  silver: 'DCDCDC',  // accent3
  yellow: 'FFC000',  // accent4
  white: 'FFFFFF',
  imgFill: 'EDEDED',
  imgLine: 'CFCFCF',
};

// Card drop shadows. pptxgenjs rewrites the object it is given, so hand out a fresh one each time.
function shadow(blur, offset, angle, opacity) {
  return { type: 'outer', color: '000000', blur, offset, angle: angle || 0, opacity: opacity || 0.15 };
}
const SHADOW = () => shadow(20, 10);
const SHADOW_SM = () => shadow(20, 8);

// ---------------------------------------------------------------- text helpers
const EYEBROW = { fontFace: F.head, fontSize: 16, color: C.head };
const TITLE = { fontFace: F.head, fontSize: 40, color: C.black };
const BODY = { fontFace: F.body, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5 };

function text(slide, str, opts) {
  slide.addText(str, Object.assign({ fontFace: F.body, fontSize: 12, color: C.body }, opts));
}
function eyebrow(slide, x, y, w, str) {
  text(slide, str, Object.assign({ x, y, w, h: 0.37 }, EYEBROW));
}
function title(slide, x, y, w, h, str, fontSize) {
  text(slide, str, Object.assign({ x, y, w, h }, TITLE, { fontSize: fontSize || 40 }));
}
function body(slide, x, y, w, h, str, opts) {
  text(slide, str, Object.assign({ x, y, w, h }, BODY, opts));
}
// Paragraph list -> pptxgenjs rich-text runs.
function lines(arr, opts) {
  return arr.map((t, i) => ({ text: t, options: Object.assign({ breakLine: i < arr.length - 1 }, opts) }));
}

// ------------------------------------------------------------- shape helpers
function card(slide, x, y, w, h, fill, radius, shdw) {
  slide.addShape('roundRect', {
    x, y, w, h, fill: { color: fill }, rectRadius: radius,
    shadow: shdw === false ? undefined : (shdw || SHADOW()),
  });
}
function pill(slide, x, y, w, h, fill, shdw) {
  card(slide, x, y, w, h, fill, h / 2, shdw);
}
// The deck's line-art pictograms are stood in for by unicode glyphs at the same box.
const GLYPH = {
  money: '\u0024', chart: '\u2588', refresh: '\u21BB', people: '\u25CF', pin: '\u2302',
  lock: '\u2299', wand: '\u2726', bulb: '\u25C9', mute: '\u25C1', clock: '\u25F4',
  mega: '\u25B6', sliders: '\u2261', globe: '\u2295', seed: '\u2663', spark: '\u2726',
};
function icon(slide, x, y, size, color, glyph) {
  text(slide, glyph || GLYPH.money, {
    x, y, w: size, h: size, align: 'center', valign: 'middle',
    fontFace: F.body, fontSize: Math.max(5, Math.round(size * 62)), color,
  });
}
// Diamond badge (rounded square rotated 45 degrees) with an icon inside.
function diamond(slide, x, y, size, fill, iconColor, glyph) {
  slide.addShape('roundRect', {
    x, y, w: size, h: size, rotate: 43.26, fill: { color: fill },
    rectRadius: size * 0.167, shadow: SHADOW_SM(),
  });
  icon(slide, x + size * 0.28, y + size * 0.28, size * 0.44, iconColor, glyph);
}
// The three device photos in the source deck become labelled placeholder panels.
// (Its other picture placeholders were left empty and render as nothing.)
function imagePlaceholder(slide, x, y, w, h, label) {
  slide.addShape('roundRect', {
    x, y, w, h, fill: { color: C.imgFill }, line: { color: C.imgLine, width: 1 }, rectRadius: 0.15,
  });
  text(slide, label || '[image]', {
    x, y: y + h / 2 - 0.2, w, h: 0.4, align: 'center', fontSize: 12, color: C.grayMid,
  });
}

// Hero backdrop: black fading to white left-to-right. The source uses a gradient
// fill (black -> transparent white); composited it is grey = 255*(1-(1-t)^2).
// pptxgenjs has no gradient API, so it is painted as a stack of thin strips.
function gradientBackdrop(slide, startFrac) {
  const strips = 128;
  const w = SLIDE_W / strips;
  for (let i = 0; i < strips; i++) {
    const t = Math.max(0, ((i + 0.5) / strips - startFrac) / (1 - startFrac));
    const v = Math.round(255 * (1 - (1 - t) * (1 - t)));
    const hex = v.toString(16).toUpperCase().padStart(2, '0').repeat(3);
    slide.addShape('rect', { x: i * w, y: 0, w: w * 1.6, h: SLIDE_H, fill: { color: hex } });
  }
}

// "Investor ~" brand mark: four nodes around a central S.
function logo(slide, color) {
  const cx = 0.612, cy = 0.590;
  [[-0.115, -0.115], [0.115, -0.115], [-0.115, 0.115], [0.115, 0.115]].forEach(([dx, dy]) => {
    slide.addShape('ellipse', { x: cx + dx - 0.028, y: cy + dy - 0.028, w: 0.056, h: 0.056, line: { color, width: 0.75 } });
  });
  text(slide, 'S', {
    x: cx - 0.16, y: cy - 0.16, w: 0.32, h: 0.32, align: 'center', valign: 'middle',
    fontFace: F.head, fontSize: 11, bold: true, color,
  });
  text(slide, ' Investor ~', {
    x: 0.771, y: 0.447, w: 1.116, h: 0.286, fontFace: F.head, fontSize: 11, italic: true, color,
  });
}
function pageNumber(slide, n) {
  text(slide, String(n), {
    x: 12.212, y: 6.832, w: 0.783, h: 0.399, align: 'right', valign: 'middle',
    fontFace: F.head, fontSize: 12, color: C.body,
  });
}

// ------------------------------------------------------------- shared copy
const LIBERO = 'Libero, Sit Amet Commodo Magna Eros Quis Urna.Nunc Viverra Imperdiet Enim. Fusce Est. ';
const LOREM_SHORT = 'Lorem Ipsum Dolor Sit Amet, Consectetuer Adipiscing Elit. ';
const LOREM_MID = 'Lorem Ipsum Dolor Sit Amet, Consectetuer Adipiscing Elit. Maecenas Porttitor Congue Massa. Fusce Posuere';
const PURUS = 'Purus Lectus Malesuada Libero, Sit Amet Commodo Magna Eros Quis';
const VIVERRA = 'Viverra Imperdiet Enim. Fusce Est. Vivamus A Tellus.Pellentesque Habitant Morbi Tristique';
const ULTRICIES = 'Ultricies, Purus Lectus Malesuada Libero, Sit Amet Commodo Magna';
const NETUS = 'Netus Et Malesuada Fames Ac Turpis Egestas. Proin Pharetra Nonummy Pede. Mauris Et Orci.';
const URNA_SHORT = 'Urna.Nunc Viverra Imperdiet Enim. Fusce Est. ';
const ASPERNATUR = 'Aspernatur Aut Odit Aut Fugit, Sed Quia Consequuntur';

// ================================================================= slide 1
function slide01(pres) {
  const s = pres.addSlide();
  gradientBackdrop(s, 0);
  // White hero panel: rounded on its right edge only, left edge runs off-slide.
  card(s, -0.5, 2.293, 7.69, 3.327, C.white, 0.1);
  s.addShape('roundRect', { x: 6.953, y: 2.293, w: 0.237, h: 3.327, fill: { color: C.grayDeep }, rectRadius: 0.085 });
  logo(s, C.white);
  diamond(s, 0.971, 3.591, 0.694, '373737', C.white);
  text(s, [
    { text: 'Marketing ', options: { color: C.black, breakLine: true } },
    { text: 'Invesment', options: { color: C.blue } },
  ], { x: 2.174, y: 3.002, w: 5.462, h: 1.717, fontFace: F.head, fontSize: 48, bold: true });
  text(s, 'Marketing Presentation', { x: 2.186, y: 2.663, w: 4.214, h: 0.337, fontFace: F.head, fontSize: 14, color: C.body });
  s.addShape('roundRect', { x: 2.302, y: 4.965, w: 1.171, h: 0.283, rectRadius: 0.1415, line: { color: C.black, width: 0.5 } });
  text(s, 'Start Slide', { x: 2.025, y: 4.964, w: 1.725, h: 0.286, align: 'center', fontFace: F.head, fontSize: 11, color: C.ink });
}

// ================================================================= slide 2
function slide02(pres) {
  const s = pres.addSlide();
  logo(s, C.blueDeep);
  pageNumber(s, 2);
  // Two stat cards under a dark banner.
  [
    { x: 3.135, fill: C.gray, pct: '87,90%', px: 3.220, pw: 2.030, tx: 3.387 },
    { x: 5.510, fill: C.blue, pct: '77,89%', px: 5.554, pw: 2.113, tx: 5.762 },
  ].forEach((col) => {
    card(s, col.x, 3.087, 2.2, 3.284, col.fill, 0.2);
    text(s, col.pct, { x: col.px, y: 3.690, w: col.pw, h: 0.438, align: 'center', fontFace: F.head, fontSize: 20, bold: true, color: C.white });
    text(s, 'Financial Assets.', { x: col.px, y: col.fill === C.gray ? 4.225 : 4.236, w: col.pw, h: 0.303, align: 'center', fontFace: F.head, fontSize: 12, color: C.white });
    body(s, col.tx, 4.684, 1.696, 0.98, LIBERO, { align: 'center', fontSize: 9, color: C.white });
  });
  card(s, 3.135, 1.886, 4.578, 1.03, C.black, 0.135);
  icon(s, 3.658, 2.195, 0.412, C.white);
  body(s, 4.256, 2.114, 3.497, 0.573, LIBERO, { fontSize: 10, color: C.white });

  eyebrow(s, 8.532, 2.367, 2.656, 'Market Investment');
  title(s, 8.487, 2.849, 4.846, 0.774, 'Introduction');
  body(s, 8.532, 3.827, 3.7, 0.667, LIBERO);
  text(s, '+45', { x: 8.543, y: 5.048, w: 1.364, h: 0.841, fontFace: F.head, fontSize: 44, bold: true, color: C.blueDeep });
  body(s, 10.716, 5.206, 1.36, 0.478, 'Fusce Est. Vivamus Tellus.Pellentesque', { fontSize: 8 });
  s.addShape('line', { x: 10.057, y: 5.469, w: 0.48, h: 0, line: { color: C.body, width: 0.5 } });
  s.addShape('ellipse', { x: 10.537, y: 5.439, w: 0.059, h: 0.059, fill: { color: C.body } });
}

// ================================================================= slide 3
function slide03(pres) {
  const s = pres.addSlide();
  logo(s, C.blueDeep);
  pageNumber(s, 3);
  eyebrow(s, 1.058, 1.655, 4.323, 'Business Innovation');
  text(s, lines(['Importance of', 'Market Investment']), { x: 1.038, y: 2.177, w: 6.02, h: 1.178, fontFace: F.head, fontSize: 32, color: C.black });
  body(s, 1.089, 3.537, 5.211, 0.364, LOREM_SHORT);
  [
    { n: '01.', y: 4.449 },
    { n: '02.', y: 5.273 },
  ].forEach((row) => {
    text(s, row.n, { x: 1.056, y: row.y, w: 1.603, h: 0.572, fontFace: F.head, fontSize: 28, bold: true, color: C.head });
    body(s, 1.938, row.y, 3.045, 0.573, 'Pulvinar Ultricies, Purus Lectus Malesuada Libero, Sit Amet Commodo Magna Eros Quis', { fontSize: 10 });
  });
  [
    { x: 8.251, fill: C.blue, pct: '56,99%', px: 8.435, tx: 8.448 },
    { x: 10.352, fill: C.gray, pct: '45,89%', px: 10.536, tx: 10.548 },
  ].forEach((col) => {
    card(s, col.x, 3.75, 1.817, 2.183, col.fill, 0.195);
    text(s, col.pct, { x: col.px, y: 4.082, w: 1.64, h: 0.37, fontFace: F.head, fontSize: 16, bold: true, color: C.white });
    body(s, col.tx, 4.562, 1.635, 1.078, PURUS, { fontSize: 10, color: C.white });
  });
}

// ================================================================= slide 4
function slide04(pres) {
  const s = pres.addSlide();
  logo(s, C.blueDeep);
  pageNumber(s, 4);
  eyebrow(s, 0.817, 1.449, 4.759, 'Ownership In Companies.');
  title(s, 0.83, 2.062, 6.006, 1.313, 'Types of Market Investments', 36);
  card(s, 4.927, 3.072, 3.732, 0.879, C.black, 0.12);
  text(s, '$54', { x: 4.53, y: 3.102, w: 1.314, h: 0.505, align: 'center', fontFace: F.head, fontSize: 24, bold: true, color: C.white });
  body(s, 5.741, 3.068, 2.947, 0.573, lines(['Viverra Imperdiet Enim. Fusce Est.', 'Vivamus A Tellus.Pellentesque']), { fontSize: 10, color: C.white });
  [
    { y: 4.685, dy: 4.766, fill: C.blue },
    { y: 5.695, dy: 5.776, fill: C.gray },
  ].forEach((row) => {
    diamond(s, 7.334, row.dy, 0.458, row.fill, C.white);
    body(s, 8.266, row.y, 3.732, 0.62, VIVERRA, { fontSize: 11 });
  });
}

// ================================================================= slide 5
function slide05(pres) {
  const s = pres.addSlide();
  logo(s, C.blueDeep);
  pageNumber(s, 5);
  eyebrow(s, 1.064, 1.497, 3.912, 'Capital Growth');
  title(s, 1.011, 2.118, 7.036, 0.774, 'Investment Goals');
  body(s, 1.064, 3.176, 5.556, 0.667, 'Lorem Ipsum Dolor Sit Amet, Consectetuer Adipiscing Elit. Maecenas Porttitor Congue Massa. Fusce Posuere, Magna Sed Pulvinar');
  [
    { x: 3.508, tx: 3.692, hx: 3.709, hw: 2.456, tw: 3.155, fill: C.blue },
    { x: 6.834, tx: 7.018, hx: 7.035, hw: 2.631, tw: 3.17, fill: C.gray },
  ].forEach((col) => {
    card(s, col.x, 4.475, 3.158, 1.658, col.fill, 0.184);
    text(s, 'Text Here', { x: col.hx, y: 4.719, w: col.hw, h: 0.404, fontFace: F.head, fontSize: 18, color: C.white });
    body(s, col.tx, 5.222, col.tw, 0.667, ULTRICIES, { color: C.white });
  });
  // "47,32K" metric with arrow marker
  icon(s, 1.166, 4.979, 0.373, C.black);
  text(s, '47,32K', { x: 2.186, y: 4.98, w: 1.414, h: 0.37, fontFace: F.head, fontSize: 16, bold: true, color: C.black });
  s.addShape('line', { x: 1.768, y: 5.165, w: 0.258, h: 0, line: { color: C.ink, width: 0.75 } });
  s.addShape('ellipse', { x: 2.004, y: 5.14, w: 0.05, h: 0.05, fill: { color: C.black } });
  body(s, 1.064, 5.456, 2.121, 0.269, 'Fusce Posuere, Magna Pulvinar', { align: 'center', fontSize: 10, lineSpacingMultiple: 1 });
}

// ================================================================= slide 6
function slide06(pres) {
  const s = pres.addSlide();
  logo(s, C.blueDeep);
  pageNumber(s, 6);
  eyebrow(s, 4.165, 2.365, 3.883, 'Come With Higher ');
  title(s, 4.169, 2.972, 5.066, 1.447, 'Risk & Return Trade-Off');
  body(s, 4.19, 4.67, 4.166, 0.667, 'Eros Quis Urna.Nunc Viverra Imperdiet Enim. Fusce Est. Vivamus A Tellus.Pellentesque Habitant Morbi');
  [
    { y: 1.18, n: '01', fill: C.dark, nx: 8.778, nw: 1.138 },
    { y: 3.167, n: '02', fill: C.gray, nx: 8.717, nw: 1.261 },
    { y: 5.154, n: '03', fill: C.blue, nx: 8.717, nw: 1.261 },
  ].forEach((row) => {
    card(s, 8.762, row.y, 3.78, 1.165, row.fill, 0.181);
    s.addShape('ellipse', { x: 8.957, y: row.y + 0.193, w: 0.781, h: 0.781, fill: { color: C.white } });
    text(s, row.n, { x: row.nx, y: row.y + 0.381, w: row.nw, h: 0.404, align: 'center', fontFace: F.head, fontSize: 18, bold: true, color: row.fill });
    text(s, 'Text Your Here', { x: 9.933, y: row.y + 0.364, w: 2.997, h: 0.438, fontFace: F.head, fontSize: 20, color: C.white });
  });
}

// ================================================================= slide 7
function slide07(pres) {
  const s = pres.addSlide();
  logo(s, C.blueDeep);
  pageNumber(s, 7);
  eyebrow(s, 1.395, 1.415, 3.783, 'Company Performance.');
  title(s, 1.33, 1.9, 5.87, 0.774, 'Market Analysis');
  body(s, 1.395, 2.902, 4.849, 0.667, 'Tristique Senectus Et Netus Et Malesuada Fames Ac Turpis Egestas. Proin Pharetra Nonummy Pede. Mauris Et Orci.');
  eyebrow(s, 1.395, 3.933, 3.546, 'Higher Conversion.');
  body(s, 1.395, 4.416, 3.802, 0.667, NETUS);
  pill(s, 1.451, 5.657, 1.508, 0.428, C.grayMid, false);
  text(s, 'Read More', { x: 1.476, y: 5.719, w: 1.458, h: 0.303, align: 'center', fontSize: 12, color: C.white });
  // Two floating stat chips
  card(s, 9.644, 0.704, 2.28, 1.196, C.gray, 0.155);
  icon(s, 10.074, 0.931, 0.25, C.white);
  text(s, '44,66%', { x: 10.404, y: 0.887, w: 1.314, h: 0.337, fontFace: F.head, fontSize: 14, color: C.white });
  body(s, 9.965, 1.239, 2.2, 0.478, LOREM_SHORT, { fontSize: 8, color: C.white });
  card(s, 7.07, 5.564, 2.28, 1.196, C.blue, 0.185);
  icon(s, 7.499, 5.791, 0.25, C.white);
  text(s, '88,17%', { x: 7.83, y: 5.747, w: 1.252, h: 0.337, fontFace: F.head, fontSize: 14, color: C.white });
  body(s, 7.391, 6.099, 2.15, 0.478, LOREM_SHORT, { fontSize: 8, color: C.white });
  // Big numbers
  [
    { x: 7.344, w: 1.995, lx: 7.176, lw: 2.331, y: 0.802, val: '+65', color: C.gray },
    { x: 9.638, w: 2.135, lx: 9.494, lw: 2.424, y: 5.729, val: '+98', color: C.blue },
  ].forEach((m) => {
    text(s, m.val, { x: m.x, y: m.y, w: m.w, h: 0.64, align: 'center', fontFace: F.head, fontSize: 32, bold: true, color: m.color });
    body(s, m.lx, m.y + 0.614, m.lw, 0.269, 'Mauris Et Orci.', { align: 'center', fontSize: 10, lineSpacingMultiple: 1 });
  });
}

// ================================================================= slide 8
function slide08(pres) {
  const s = pres.addSlide();
  logo(s, C.blueDeep);
  pageNumber(s, 8);
  eyebrow(s, 6.945, 1.095, 5.102, 'Government Policies');
  title(s, 6.883, 1.553, 6.451, 1.447, 'Factors Affecting Investments');
  body(s, 6.957, 3.111, 4.967, 0.667, 'Lorem Ipsum Dolor Sit Amet, Consectetuer Adipiscing Elit. Maecenas Porttitor Congue Massa. Fusce Posuere, Magna Sed');
  icon(s, 8.913, 4.911, 0.471, C.blueDeep);
  text(s, 'Industry Trends.', { x: 9.588, y: 4.945, w: 3.746, h: 0.404, fontFace: F.head, fontSize: 18, color: C.body });
  body(s, 8.792, 5.579, 3.75, 0.573, 'Nunc Viverra Imperdiet Enim. Fusce . Vivamus Tellus.Pellentesque Habitant Morbi Tristique', { fontSize: 10 });
  card(s, 1.108, 5.235, 3.135, 0.842, C.gray, 0.14);
  icon(s, 1.465, 5.435, 0.441, C.white);
  body(s, 2.118, 5.37, 2.611, 0.573, 'Pulvinar Ultricies, Purus Lectus Malesuada Libero, ', { fontSize: 10, color: C.white });
}

// ================================================================= slide 9
function slide09(pres) {
  const s = pres.addSlide();
  logo(s, C.blueDeep);
  pageNumber(s, 9);
  eyebrow(s, 1.139, 1.401, 4.19, 'Eggs In One Basket.');
  title(s, 1.101, 1.935, 5.566, 1.447, 'Portfolio Diversification');
  body(s, 1.164, 3.565, 4.357, 0.667, 'Lorem Ipsum Dolor Sit Amet, Consectetuer Adipiscing Elit. Maecenas Porttitor Congue Massa. Fusce Posuere, ');
  [
    { x: 1.278, fill: C.blue, ix: 1.542, tx: 2.164, tw: 3.213 },
    { x: 5.372, fill: C.gray, ix: 5.636, tx: 6.257, tw: 3.495 },
  ].forEach((col) => {
    card(s, col.x, 4.812, 3.736, 0.933, col.fill, 0.155);
    icon(s, col.ix, 5.063, 0.431, C.white);
    body(s, col.tx, 4.992, col.tw, 0.573, lines(['Sed Pulvinar Ultricies, Purus Lectus', 'Malesuada Libero, Sit Amet Commodo']), { fontSize: 10, color: C.white });
  });
}

// ================================================================= slide 10
function slide10(pres) {
  const s = pres.addSlide();
  logo(s, C.blueDeep);
  pageNumber(s, 10);
  text(s, 'Investment Objectives.', { x: 3.727, y: 0.684, w: 5.94, h: 0.37, align: 'center', fontFace: F.head, fontSize: 16, color: C.head });
  text(s, 'Risk Management Investment', { x: 0.06, y: 1.084, w: 13.273, h: 0.64, align: 'center', fontFace: F.head, fontSize: 32, color: C.black });

  // Grey S-curve ribbon threaded behind the four badges: a chain of semicircle
  // arcs (outer dia 2.332", centres 2.148" apart) alternating below / above y=4.571.
  const ARC_D = 2.332;
  for (let i = 0; i < 4; i++) {
    const cx = 3.474 + i * 2.148;
    s.addShape('blockArc', {
      x: cx - ARC_D / 2, y: 4.571 - ARC_D / 2, w: ARC_D, h: ARC_D, fill: { color: 'EDEDED' },
      angleRange: i % 2 === 0 ? [0, 180] : [180, 360], arcThicknessRatio: 0.158,
    });
  }

  const badges = [
    { x: 2.886, fill: C.black, glyph: GLYPH.lock },
    { x: 5.057, fill: C.blue, glyph: GLYPH.wand },
    { x: 7.228, fill: C.gray, glyph: GLYPH.bulb },
    { x: 9.338, fill: C.yellow, glyph: GLYPH.pin },
  ];
  badges.forEach((b) => {
    s.addShape('ellipse', { x: b.x, y: 3.978, w: 1.176, h: 1.176, fill: { color: b.fill } });
    icon(s, b.x + 0.42, 4.397, 0.339, C.white, b.glyph);
  });

  // Callout stems: two above (right aligned copy), two below (left aligned copy).
  const stems = [
    { color: C.black, seg: [[4.371, 2.702, 0, 0.837], [4.05, 3.534, 0.323, 0.408, true]], dot: [4.336, 2.638] },
    { color: C.gray, seg: [[8.804, 2.692, 0, 0.837], [8.484, 3.525, 0.323, 0.408, true]], dot: [8.77, 2.629] },
    { color: C.blue, seg: [[4.755, 5.115, 0.004, 0.784], [4.751, 4.976, 0.186, 0.145, true]], dot: [4.727, 5.899] },
    { color: C.yellow, seg: [[9.039, 5.106, 0.004, 0.784], [9.035, 4.968, 0.186, 0.145, true]], dot: [9.015, 5.89] },
  ];
  stems.forEach((st) => {
    st.seg.forEach(([x, y, w, h, flipV]) => {
      s.addShape('line', { x, y, w, h, flipV: !!flipV, line: { color: st.color, width: 1 } });
    });
    s.addShape('ellipse', { x: st.dot[0], y: st.dot[1], w: 0.068, h: 0.068, fill: { color: st.color } });
  });

  const callouts = [
    { hx: 1.764, tx: 1.424, y: 2.373, by: 2.833, align: 'right', tw: 2.66 },
    { hx: 6.106, tx: 5.766, y: 2.373, by: 2.833, align: 'right', tw: 2.66 },
    { hx: 4.996, tx: 4.996, y: 5.69, by: 6.149, align: 'left', tw: 2.75 },
    { hx: 9.25, tx: 9.25, y: 5.69, by: 6.149, align: 'left', tw: 2.75 },
  ];
  callouts.forEach((c) => {
    text(s, 'Option Here', { x: c.hx, y: c.y, w: 2.32, h: 0.459, align: c.align, fontFace: F.head, fontSize: 16, color: C.soft, lineSpacingMultiple: 1.5 });
    body(s, c.tx, c.by, c.tw, 0.667, 'Eos Et Accusamus Et Iusto Odio Dignissimos Ducimus', { align: c.align });
  });
}

// ================================================================= slide 11
function slide11(pres) {
  const s = pres.addSlide();
  logo(s, C.blueDeep);
  pageNumber(s, 11);
  text(s, 'Compounding Returns.', { x: 5.027, y: 0.905, w: 3.343, h: 0.37, align: 'center', fontFace: F.head, fontSize: 16, color: C.head });
  text(s, 'Time Horizon in Investing', { x: 4.374, y: 1.38, w: 4.648, h: 1.178, align: 'center', fontFace: F.head, fontSize: 32, color: C.black });

  // Four jigsaw discs (dia 1.823") in a 2x2 block; each pushes a tab into the
  // neighbour it locks with. Tabs are drawn last so they sit over the next disc.
  const DISC = 1.823;
  const pieces = [
    { x: 4.868, y: 2.897, fill: C.black, cx: 5.471, cy: 3.515, glyph: GLYPH.mute,
      tab: { x: 6.60, y: 3.833, w: 0.39, h: 0.167 } },   // tab points right
    { x: 6.727, y: 2.897, fill: C.blue, cx: 7.123, cy: 3.515, glyph: GLYPH.clock,
      tab: { x: 7.45, y: 4.64, w: 0.139, h: 0.375 } },   // tab points down
    { x: 4.868, y: 4.772, fill: C.yellow, cx: 5.471, cy: 5.149, glyph: GLYPH.bulb,
      tab: { x: 5.78, y: 4.486, w: 0.20, h: 0.37 } },    // tab points up
    { x: 6.728, y: 4.772, fill: C.gray, cx: 7.123, cy: 5.149, glyph: GLYPH.mega,
      tab: { x: 6.42, y: 5.49, w: 0.37, h: 0.16 } },     // tab points left
  ];
  pieces.forEach((q) => s.addShape('ellipse', { x: q.x, y: q.y, w: DISC, h: DISC, fill: { color: q.fill } }));
  pieces.forEach((q) => {
    const g = 0.045;
    s.addShape('ellipse', { x: q.tab.x - g, y: q.tab.y - g, w: q.tab.w + 2 * g, h: q.tab.h + 2 * g, fill: { color: C.white } });
    s.addShape('ellipse', { x: q.tab.x, y: q.tab.y, w: q.tab.w, h: q.tab.h, fill: { color: q.fill } });
  });
  pieces.forEach((q) => {
    s.addShape('ellipse', { x: q.cx, y: q.cy, w: 0.84, h: 0.826, fill: { color: C.white } });
    icon(s, q.cx + 0.296, q.cy + 0.289, 0.248, q.fill, q.glyph);
  });

  // Leader lines out to the four labels
  const leaders = [
    { seg: [[4.009, 2.909, 0.708, 0, true, false], [4.711, 2.907, 0.145, 0.186, true, true]], dot: [3.957, 2.876], color: C.body },
    { seg: [[3.638, 5.313, 0.842, 0, true, false], [4.474, 5.311, 0.219, 0.197, true, true]], dot: [3.57, 5.279], color: C.yellow },
    { seg: [[8.69, 2.925, 0.708, 0, false, false], [8.55, 2.923, 0.145, 0.186, false, true]], dot: [9.382, 2.892], color: C.blue },
    { seg: [[8.909, 5.303, 0.842, 0, false, false], [8.696, 5.301, 0.219, 0.197, false, true]], dot: [9.752, 5.269], color: C.gray },
  ];
  leaders.forEach((l) => {
    l.seg.forEach(([x, y, w, h, flipH, flipV]) => {
      s.addShape('line', { x, y, w, h, flipH, flipV, line: { color: l.color, width: 1 } });
    });
    s.addShape('ellipse', { x: l.dot[0], y: l.dot[1], w: 0.068, h: 0.068, fill: { color: l.color } });
  });

  const labels = [
    { hx: 1.517, tx: 1.838, hy: 2.568, ty: 3.076, align: 'right' },
    { hx: 1.174, tx: 1.495, hy: 5.003, ty: 5.512, align: 'right' },
    { hx: 9.602, tx: 9.602, hy: 2.588, ty: 3.096, align: 'left' },
    { hx: 9.853, tx: 9.853, hy: 5.002, ty: 5.511, align: 'left' },
  ];
  labels.forEach((l) => {
    text(s, 'Title Here', { x: l.hx, y: l.hy, w: 2.307, h: 0.504, align: l.align, fontFace: F.head, fontSize: 18, color: C.ink, lineSpacingMultiple: 1.5 });
    body(s, l.tx, l.ty, 1.99, 0.897, 'In Voluptate Velit Esse Cillum Dolore Eu Fugiat Nulla Pariatur. ', { align: l.align, fontSize: 11 });
  });
}

// ================================================================= slide 12
function slide12(pres) {
  const s = pres.addSlide();
  logo(s, C.blueDeep);
  pageNumber(s, 12);
  text(s, 'Improve Accessibility.', { x: 4.167, y: 0.48, w: 5, h: 0.37, align: 'center', fontFace: F.head, fontSize: 16, color: C.head });
  text(s, 'Role of Technology', { x: 2.667, y: 0.883, w: 8, h: 0.707, align: 'center', fontFace: F.head, fontSize: 36, color: C.black });

  const plans = [
    {
      head: 'BASIC', headFill: C.black, price: '$15', featured: false,
      body: { x: 2.208, y: 3.284, w: 2.269, h: 3.365 },
      cap: { x: 2.165, y: 2.304, w: 2.349, h: 1.513 },
      tab: { x: 2.298, y: 3.261, w: 1.501, h: 0.467 },
      title: { x: 2.084, y: 2.628, w: 2.511, h: 0.438, size: 20 },
      priceBox: { x: 2.18, y: 3.272, w: 1.737, h: 0.438, size: 20, unit: 16 },
      bullets: [4.213, 4.784, 5.318, 5.907], bulletX: 2.33, textX: 2.516, bulletSize: 0.186, fontSize: 11,
    },
    {
      head: 'PREMIUM', headFill: C.blue, price: '$50', featured: true,
      body: { x: 5.299, y: 3.052, w: 2.739, h: 4.062 },
      cap: { x: 5.251, y: 1.868, w: 2.838, h: 1.827 },
      tab: { x: 5.405, y: 3.024, w: 1.812, h: 0.564 },
      title: { x: 5.154, y: 2.217, w: 3.032, h: 0.582, size: 28 },
      priceBox: { x: 5.26, y: 3.054, w: 2.102, h: 0.505, size: 24, unit: 18 },
      bullets: [4.129, 4.701, 5.234, 5.824, 6.355], bulletX: 5.656, textX: 5.846, bulletSize: 0.193, fontSize: 12,
    },
    {
      head: 'STANDART', headFill: C.gray, price: '$25', featured: false,
      body: { x: 8.861, y: 3.256, w: 2.269, h: 3.365 },
      cap: { x: 8.822, y: 2.276, w: 2.346, h: 1.513 },
      tab: { x: 8.951, y: 3.233, w: 1.498, h: 0.467 },
      title: { x: 8.739, y: 2.6, w: 2.511, h: 0.438, size: 20 },
      priceBox: { x: 8.832, y: 3.244, w: 1.737, h: 0.438, size: 20, unit: 16 },
      bullets: [4.185, 4.757, 5.29, 5.88], bulletX: 8.985, textX: 9.171, bulletSize: 0.186, fontSize: 11,
    },
  ];
  const FEATURES = ['Quis nostrud', 'Exercitation ullamco', 'Laboris nisi', 'Commodo consequat', 'Quis nostrud'];
  plans.forEach((p) => {
    card(s, p.cap.x, p.cap.y, p.cap.w, p.cap.h, p.headFill, 0.26, false);
    s.addShape('roundRect', {
      x: p.body.x, y: p.body.y, w: p.body.w, h: p.body.h, fill: { color: C.white }, rectRadius: 0.26,
      shadow: { type: 'outer', color: '000000', blur: 37, offset: 15, angle: 40, opacity: 0.12 },
    });
    text(s, p.head, { x: p.title.x, y: p.title.y, w: p.title.w, h: p.title.h, align: 'center', valign: 'middle', fontFace: F.head, fontSize: p.title.size, color: C.white });
    pill(s, p.tab.x, p.tab.y, p.tab.w, p.tab.h, C.white, { type: 'outer', color: '000000', blur: 37, offset: 5, angle: 40, opacity: 0.16 });
    text(s, [
      { text: p.price, options: { fontSize: p.priceBox.size } },
      { text: '/Month', options: { fontSize: p.priceBox.unit, baseline: 600 } },
    ], { x: p.priceBox.x, y: p.priceBox.y, w: p.priceBox.w, h: p.priceBox.h, align: 'center', valign: 'middle', fontFace: F.head, color: '1A1A1A' });
    p.bullets.forEach((by, i) => {
      const cross = i === 1;
      s.addShape('ellipse', { x: p.bulletX, y: by, w: p.bulletSize, h: p.bulletSize, fill: { color: p.headFill } });
      text(s, cross ? '\u2715' : '\u2713', {
        x: p.bulletX, y: by, w: p.bulletSize, h: p.bulletSize, align: 'center', valign: 'middle',
        fontFace: F.body, fontSize: 7, color: C.white,
      });
      text(s, FEATURES[i], { x: p.textX, y: by - 0.061, w: 1.834, h: 0.303, fontSize: p.fontSize, color: C.body });
    });
  });
}

// ================================================================= slide 13
function slide13(pres) {
  const s = pres.addSlide();
  logo(s, C.blueDeep);
  pageNumber(s, 13);
  eyebrow(s, 1.313, 1.603, 3.781, 'Market Fluctuations.');
  title(s, 1.268, 2.105, 5.132, 1.717, 'Behavioral Finance', 48);
  body(s, 1.335, 4.013, 3.817, 0.667, 'Urna.Nunc Viverra Imperdiet Enim. Fusce Est. Vivamus A Tellus.Pellentesque Habitant');
  text(s, '89%', { x: 1.363, y: 5.257, w: 1.825, h: 0.707, fontFace: F.head, fontSize: 36, bold: true, color: C.blue });
  body(s, 2.702, 5.347, 2.157, 0.525, lines(['Morbi Tristique Senectus Et', 'Netus Et Malesuada']), { fontSize: 9 });

  card(s, 6.637, 1.514, 2.192, 2.43, C.black, 0.169, false);
  text(s, '98,77%', { x: 6.751, y: 1.923, w: 1.999, h: 0.404, fontFace: F.head, fontSize: 18, bold: true, color: C.white });
  body(s, 6.768, 2.457, 2.232, 1.078, LOREM_MID, { fontSize: 10, color: C.white });

  [
    { x: 5.6, fill: C.gray, ix: 5.814, tx: 5.702, tc: C.white },
    { x: 7.093, fill: C.blue, ix: 7.306, tx: 7.195, tc: C.white },
    { x: 8.585, fill: C.white, ix: 8.802, tx: 8.686, tc: C.body },
  ].forEach((col) => {
    card(s, col.x, col.fill === C.white ? 4.326 : 4.332, 1.279, 1.654, col.fill, 0.172);
    icon(s, col.ix, 4.662, 0.25, col.tc === C.white ? C.white : C.body);
    body(s, col.tx, 4.976, 1.369, 0.68, lines(['Urna.Nunc Viverra Imperdiet Enim.', 'Fusce Est. ']), { fontSize: 8, color: col.tc });
  });
}

// ================================================================= slide 14
function slide14(pres) {
  const s = pres.addSlide();
  logo(s, C.blueDeep);
  pageNumber(s, 14);
  eyebrow(s, 0.886, 1.312, 4.302, 'Responsible Investments.');
  title(s, 0.886, 1.863, 7.279, 0.707, 'Sustainable Investing', 36);
  body(s, 0.903, 2.813, 6.203, 0.667, 'Lorem Ipsum Dolor Amet, Consectetuer Adipiscing Elit. Maecenas Porttitor Congue Massa. Fusce Posuere, Magna Sed Pulvinar Ultricies, Purus Lectus');
  [
    { x: 5.745, fill: C.black, pct: '88,34%', px: 5.874, pw: 1.217, tx: 5.847 },
    { x: 7.256, fill: C.blue, pct: '23,45%', px: 7.37, pw: 1.183, tx: 7.357 },
    { x: 8.766, fill: C.gray, pct: '97,07%', px: 8.881, pw: 1.203, tx: 8.868 },
  ].forEach((col) => {
    card(s, col.x, 5.022, 1.279, 1.654, col.fill, 0.172);
    text(s, col.pct, { x: col.px, y: 5.274, w: col.pw, h: 0.337, fontFace: F.head, fontSize: 14, bold: true, color: C.white });
    body(s, col.tx, 5.666, 1.076, 0.68, URNA_SHORT, { fontSize: 8, color: C.white });
  });
  text(s, '$77', { x: 10.475, y: 5.53, w: 1.442, h: 0.64, fontFace: F.head, fontSize: 32, bold: true, color: C.dark });
  body(s, 11.534, 5.473, 1.097, 0.753, 'Nunc Viverra Imperdiet Enim. Fusce Est. ', { fontSize: 9 });
}

// ================================================================= slide 15
function slide15(pres) {
  const s = pres.addSlide();
  logo(s, C.blueDeep);
  pageNumber(s, 15);
  [
    { x: 5.85, fill: C.blue, ix: 6.195, tx: 6.63, tw: 2.276, lines: ['Developed Markets ', 'Stability And Liquidity.'] },
    { x: 9.173, fill: C.gray, ix: 9.424, tx: 9.822, tw: 2.544, lines: ['Diversifying Globally', 'Reduces Local Economic'] },
  ].forEach((col) => {
    card(s, col.x, 2.056, 3.051, 0.846, col.fill, 0.117);
    icon(s, col.ix, 2.318, 0.32, C.white);
    text(s, lines(col.lines), { x: col.tx, y: 2.167, w: col.tw, h: 0.625, fontFace: F.head, fontSize: 11, color: C.white, lineSpacingMultiple: 1.5 });
  });
  eyebrow(s, 6.969, 3.932, 2.667, 'Emerging Markets');
  title(s, 6.969, 4.452, 6.365, 0.774, 'Global Investment');
  body(s, 6.969, 5.377, 5.396, 0.667, 'Lorem Ipsum Dolor Sit Amet, Consectetuer Adipiscing Elit. Maecenas Porttitor Congue Massa. Fusce Posuere, Magna Sed Pulvinar Ultricies, ');
}

// ================================================================= slide 16
function slide16(pres) {
  const s = pres.addSlide();
  logo(s, C.blueDeep);
  pageNumber(s, 16);
  eyebrow(s, 0.772, 1.791, 3.581, 'Investment Insights.');
  title(s, 0.718, 2.246, 5.949, 0.841, 'Future Trends', 44);
  body(s, 0.79, 3.299, 4.573, 0.667, 'Purus Lectus Malesuada Libero, Amet Commodo Magna Eros Quis Urna.Nunc Viverra Imperdiet Enim. Fusce Est. ');
  eyebrow(s, 0.78, 4.368, 3.067, 'Rising Interest');
  body(s, 0.79, 4.86, 4.308, 0.667, 'Pellentesque Habitant Morbi Tristique Senectus  Netus Et Malesuada Fames Ac Turpis Egestas. Proin Pharetra');
  text(s, '+12', { x: 10.221, y: 1.708, w: 1.085, h: 0.64, fontFace: F.head, fontSize: 32, bold: true, color: C.black });
  body(s, 11.174, 1.741, 0.814, 0.573, 'Fusce Est. Vivamus ', { fontSize: 10 });
  [
    { y: 4.601, fill: C.gray, ix: 6.999, isz: 0.359, iy: 4.794, tx: 7.51, tw: 3.137, ty: 4.806, label: 'Increased Use Of Digital' },
    { y: 5.701, fill: C.blue, ix: 6.972, isz: 0.34, iy: 5.904, tx: 7.458, tw: 3.228, ty: 5.906, label: 'Rising Interest Alternative' },
  ].forEach((row) => {
    card(s, 6.625, row.y, 3.831, 0.746, row.fill, 0.124);
    icon(s, row.ix, row.iy, row.isz, C.white);
    text(s, row.label, { x: row.tx, y: row.ty, w: row.tw, h: 0.337, fontFace: F.head, fontSize: 14, color: C.white });
  });
}

// ================================================================= slide 17
function slide17(pres) {
  const s = pres.addSlide();
  logo(s, C.blueDeep);
  pageNumber(s, 17);
  imagePlaceholder(s, 6.478, 0.648, 6.868, 6.102, '[image: monitor mockup]');
  title(s, 1.015, 1.33, 5.032, 1.582, 'Investment Mockup', 44);
  body(s, 1.083, 3.069, 3.577, 0.667, 'Tellus.Pellentesque Habitant Morbi Tristique Senectus Et Netus Et Malesuada Fames Turpis');

  // Column chart drawn with shapes: 5 gridlines, 2 categories x 3 series.
  const plot = { x: 1.47, y: 4.409, w: 2.202, h: 1.366 };
  for (let i = 0; i < 5; i++) {
    const gy = plot.y + (plot.h - 0.272) * (i / 4);
    s.addShape('line', { x: plot.x, y: gy, w: plot.w, h: 0, line: { color: 'F2F2F2', width: 1.1 } });
  }
  s.addShape('line', { x: 1.466, y: 5.771, w: 2.202, h: 0, line: { color: 'BFBFBF', width: 1.9 } });
  const SERIES = [
    { color: C.black, values: [1.177, 0.683] },
    { color: C.blue, values: [0.655, 1.205] },
    { color: C.silver, values: [0.55, 0.55] },
  ];
  [1.602, 2.703].forEach((groupX, cat) => {
    SERIES.forEach((ser, si) => {
      const h = ser.values[cat];
      s.addShape('rect', {
        x: groupX + si * 0.251, y: 5.774 - h, w: 0.335, h,
        fill: { color: ser.color, transparency: 30 },
      });
    });
  });
  ['1', '2', '3', '4', '5'].forEach((lab, i) => {
    text(s, lab, { x: 1.076, y: 5.595 - i * 0.3, w: 0.208, h: 0.198, fontSize: 10, color: '595959' });
  });
  ['Category 1', 'Category 2'].forEach((lab, i) => {
    text(s, lab, { x: 1.462 + i * 1.068, y: 5.862, w: 1.121, h: 0.269, align: 'center', fontSize: 10, color: '595959' });
  });
  body(s, 3.916, 4.901, 2.16, 0.606, 'Ipsum Quia Dolor Sit Ametadipisci Velit  Sed Quia', { fontSize: 10 });

  [
    { y: 1.695, x: 5.56, fill: C.gray, pct: '76,45%', px: 6.446, pw: 1.354, ix: 5.933, tx: 5.824, tw: 2.788, ty: 2.26 },
    { y: 3.23, x: 5.54, fill: C.blue, pct: '34,76%', px: 6.429, pw: 1.418, ix: 5.942, tx: 5.807, tw: 2.943, ty: 3.795 },
  ].forEach((cardOpts) => {
    card(s, cardOpts.x, cardOpts.y, 2.805, 1.3, cardOpts.fill, 0.132);
    icon(s, cardOpts.ix, cardOpts.y + 0.163, 0.342, C.white);
    text(s, cardOpts.pct, { x: cardOpts.px, y: cardOpts.y + 0.165, w: cardOpts.pw, h: 0.337, fontFace: F.head, fontSize: 14, bold: true, color: C.white });
    body(s, cardOpts.tx, cardOpts.ty, cardOpts.tw, 0.573, 'Morbi Tristique Senectus Et Netus Et Malesuada Fames Ac Turpis', { fontSize: 10, color: C.white });
  });
}

// ================================================================= slide 18
function slide18(pres) {
  const s = pres.addSlide();
  logo(s, C.blueDeep);
  pageNumber(s, 18);
  imagePlaceholder(s, 6.42, 0.528, 3.699, 6.335, '[image: phone mockup]');
  pill(s, 1.135, 2.005, 1.779, 0.4, C.white, { type: 'outer', color: '000000', blur: 16, offset: 9, angle: 90, opacity: 0.15 });
  text(s, 'Let\u2019s Join Us!', { x: 1.132, y: 2.054, w: 1.784, h: 0.303, align: 'center', fontSize: 12, bold: true, color: C.body });
  eyebrow(s, 1.009, 3.073, 3.66, 'Market Volatility');
  title(s, 1.009, 3.653, 5.658, 0.774, 'Key Challenges');
  body(s, 1.061, 4.72, 4.965, 0.667, LOREM_MID, { color: C.soft });
  text(s, '70%', { x: 10.174, y: 0.844, w: 1.454, h: 0.572, valign: 'middle', fontFace: F.head, fontSize: 28, bold: true, color: C.blue });
  text(s, 'Inventory and Order Management', { x: 10.174, y: 1.377, w: 1.973, h: 0.572, fontFace: F.head, fontSize: 14, color: C.ink });
  diamond(s, 6.257, 3.584, 0.797, C.white, C.blue);

  // Score gauge card
  card(s, 8.799, 3.885, 3.348, 2.218, C.white, 0.234);
  text(s, lines(['4 Product by', 'Spend']), { x: 9.27, y: 4.039, w: 2.406, h: 0.606, align: 'center', fontFace: F.head, fontSize: 18, bold: true, color: C.body });
  const G = { x: 9.17, y: 4.795, w: 2.607, h: 2.607 };
  [
    { color: C.yellow, sweep: 180 },
    { color: C.silver, sweep: 145.4 },
    { color: C.blue, sweep: 112.2 },
    { color: C.black, sweep: 71.5 },
  ].forEach((seg) => {
    s.addShape('pie', { x: G.x, y: G.y, w: G.w, h: G.h, fill: { color: seg.color }, angleRange: [0, seg.sweep], rotate: 180 });
  });
  s.addShape('pie', { x: 9.44, y: 5.065, w: 2.067, h: 2.067, fill: { color: C.white }, angleRange: [359.95, 180], rotate: 180 });
  pill(s, 9.969, 5.371, 1.007, 0.604, C.white, { type: 'outer', color: '000000', blur: 29, offset: 18, angle: 60, opacity: 0.05 });
  text(s, '5,892', { x: 9.825, y: 5.414, w: 1.297, h: 0.404, align: 'center', fontFace: F.head, fontSize: 18, bold: true, color: C.head });
  text(s, 'Total Score', { x: 9.825, y: 5.766, w: 1.297, h: 0.202, align: 'center', fontFace: F.head, fontSize: 6, color: C.black });
  pill(s, 11.204, 4.968, 0.62, 0.236, C.soft, false);
  text(s, [
    { text: '+', options: { color: C.black } },
    { text: '657%', options: { color: C.white } },
  ], { x: 11.127, y: 4.968, w: 0.774, h: 0.236, align: 'center', valign: 'middle', fontFace: F.head, fontSize: 8 });
}

// ================================================================= slide 19
function slide19(pres) {
  const s = pres.addSlide();
  logo(s, C.blueDeep);
  pageNumber(s, 19);
  imagePlaceholder(s, 5.555, 0.879, 7.111, 6.008, '[image: laptop mockup]');
  eyebrow(s, 1.058, 1.697, 3.365, 'Timeless Principles.');
  title(s, 1.014, 2.268, 4.816, 0.841, 'Conclusion', 44);
  text(s, 'Subtitle Here', { x: 1.036, y: 3.237, w: 2.712, h: 0.505, fontFace: F.head, fontSize: 16, color: C.blueDeep, lineSpacingMultiple: 1.5 });
  body(s, 1.047, 3.872, 4.069, 0.707, 'Voluptatem Quia Voluptas Sit Aspernatur Aut Odit Aut Fugit, Sed Quia Consequuntur Magni');
  text(s, '89%', { x: 1.065, y: 4.898, w: 0.854, h: 0.404, fontFace: F.head, fontSize: 12, color: C.blueDeep, lineSpacingMultiple: 1.5 });
  s.addShape('line', { x: 1.172, y: 5.289, w: 2.528, h: 0, line: { color: '858585', width: 2 } });
  s.addShape('line', { x: 1.172, y: 5.289, w: 1.995, h: 0, line: { color: C.blue, width: 2 } });
  body(s, 1.065, 5.438, 3.124, 0.404, 'Sed Quia Consequuntur Magni Eos');
  [
    { x: 5.641, fill: C.blue, ix: 5.851, iy: 2.916, px: 6.14, pw: 0.888, py: 2.89, pct: '43%', tx: 5.739, ty: 3.263 },
    { x: 7.19, fill: C.gray, ix: 7.383, iy: 2.919, px: 7.679, pw: 0.883, py: 2.892, pct: '76%', tx: 7.289, ty: 3.266 },
  ].forEach((col) => {
    card(s, col.x, col.fill === C.blue ? 2.641 : 2.644, 1.403, 1.551, col.fill, 0.126);
    icon(s, col.ix, col.iy, 0.25, C.white);
    text(s, col.pct, { x: col.px, y: col.py, w: col.pw, h: 0.303, fontFace: F.head, fontSize: 12, bold: true, color: C.white });
    body(s, col.tx, col.ty, 1.343, 0.68, ASPERNATUR, { fontSize: 8, color: C.white });
  });
}

// ================================================================= slide 20
function slide20(pres) {
  const s = pres.addSlide();
  gradientBackdrop(s, 0.26);
  logo(s, C.white);
  text(s, '20', { x: 12.212, y: 6.832, w: 0.783, h: 0.399, align: 'right', valign: 'middle', fontFace: F.head, fontSize: 12, color: C.body });
  text(s, 'Thank You', { x: 1.043, y: 1.664, w: 7.294, h: 1.212, fontFace: F.head, fontSize: 66, bold: true, color: C.white });
  pill(s, 1.193, 3.181, 3.768, 0.439, C.blue);
  text(s, 'Finish In The Presentation', { x: 1.342, y: 3.249, w: 3.472, h: 0.303, align: 'center', fontFace: F.head, fontSize: 12, color: C.white });
  body(s, 1.077, 4.159, 4.793, 0.573, 'Nunc Viverra Imperdiet Enim. Fusce Est. Vivamus A Tellus.Pellentesque Habitant Morbi Tristique Senectus Et Netus Et Malesuada.', { fontSize: 10, color: C.white });
  [
    { x: 1.225, fill: C.white, ic: C.ink },
    { x: 1.98, fill: C.grayDeep, ic: C.white },
    { x: 2.736, fill: C.blueDeep, ic: C.white },
  ].forEach((d) => diamond(s, d.x, 5.486, 0.486, d.fill, d.ic));
}

// ===================================================================== build
function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'WIDE', width: SLIDE_W, height: SLIDE_H });
  pres.layout = 'WIDE';
  pres.theme = { headFontFace: F.head, bodyFontFace: F.body };
  pres.title = 'Marketing Invesment';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach((fn) => fn(pres));

  return pres.writeFile({ fileName: OUT });
}

build().then(() => console.log('wrote', OUT));
