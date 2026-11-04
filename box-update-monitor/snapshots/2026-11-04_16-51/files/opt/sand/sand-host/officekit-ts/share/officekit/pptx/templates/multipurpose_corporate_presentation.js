/**
 * "Freshno" multipurpose deck (27 slides, 16:9 / 13.333 x 7.5 in) rebuilt with pptxgenjs.
 *
 * Run:  node 08d7236f-3965-4319-93c4-9f5a7f8b6677_grok_final.js
 * Out:  08d7236f-3965-4319-93c4-9f5a7f8b6677_grok_final.pptx  (next to this file)
 *
 * The source deck contains no raster media at all - its "photo" areas are empty
 * PowerPoint picture placeholders (pattern fill, no image), which render blank.
 * Vector icon art and the world map are re-drawn here with native shapes.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const C = {
  blue: '003586',   // brand blue
  ink: '262626',    // tx1 lumMod 85% / lumOff 15%
  white: 'FFFFFF',
  gray85: 'D9D9D9', // bg1 lumMod 85%
  gray95: 'F2F2F2'  // bg1 lumMod 95%
};

const F = {
  reg: 'Montserrat',
  med: 'Montserrat Medium',
  semi: 'Montserrat SemiBold',
  xbold: 'Montserrat ExtraBold'
};

/* -------------------------------------------------------------- text helpers */
// plain run / bold run; extra takes pptxgenjs run options.
// NOTE: options must be a fresh object per run - pptxgenjs writes inherited
// shape options (align, color, ...) back into them.
const t = (text, extra = {}) => ({ text, options: { ...extra } });
const b = (text, extra = {}) => ({ text, options: { bold: true, ...extra } });
const NL = { softBreakBefore: true };          // line break, same paragraph
const END = { breakLine: true };               // end of paragraph
const EMPTY = () => t('', END);                // blank paragraph

function text(slide, runs, opts) {
  slide.addText(runs, { margin: 0, valign: 'top', fontFace: F.reg, color: C.ink, ...opts });
}

// 11 pt body copy: 130% line spacing + 10 pt space-before, as in the source deck
function body(slide, runs, opts) {
  text(slide, runs, { fontSize: 11, lineSpacingMultiple: 1.3, paraSpaceBefore: 10, ...opts });
}

// 18 pt ExtraBold two-tone section heading, e.g. ["ABOUT ", ink] + ["US", blue]
function head(slide, box, parts) {
  text(slide, parts.map(([s, color]) => b(s, { color })),
    { ...box, h: box.h || 0.606, fontSize: 18, fontFace: F.xbold });
}

// 9 pt "Photo Text" caption: bold label, line break, grey detail line
function caption(slide, box, label, detail, color = C.ink) {
  body(slide, [b(label), t(detail, NL)], { ...box, fontSize: 9, color });
}

/* ------------------------------------------------------------- shape helpers */
// pptxgenjs writes <a:gd name="adj" fmla="val ..."> from rectRadius, so the
// original parallelogram `adj` values can be reproduced exactly.
function slant(slide, { x, y, w, h, adj, fill }) {
  slide.addShape('parallelogram', {
    x, y, w, h,
    rectRadius: (adj / 100000) * Math.min(w, h),
    fill: { color: fill },
    line: { type: 'none' }
  });
}

function band(slide, { x, y, w, h, fill }) {
  slide.addShape('rect', { x, y, w, h, fill: { color: fill }, line: { type: 'none' } });
}

/* ------------------------------------------------------------- page furniture */
// "Freshno . SLD 07" / page number / "www.freshnosld.com" strip
function chrome(slide, color = C.ink, num = null) {
  const opts = { fontSize: 11, color, lineSpacingMultiple: 1.3, paraSpaceBefore: 10 };
  text(slide, [b('Freshno'), b(' . SLD 07')], { x: 0.697, y: 6.845, w: 2.59, h: 0.219, ...opts });
  text(slide, [b('www.freshnosld.com')], { x: 10.851, y: 6.853, w: 2.59, h: 0.219, ...opts });
  if (num !== null) {
    text(slide, [b(String(num))], { x: 6.283, y: 6.886, w: 0.6, h: 0.185, fontSize: 11, color, wrap: false });
  }
}

/* ------------------------------------------------------------- icon stand-ins */
// Simple native-shape stand-ins for the deck's vector icon art.
function iconBottles(slide, x, y, s, color) {
  const line = { color, width: 1.25 };
  [[0.02, 0.30, 0.32, 0.70], [0.36, 0.24, 0.28, 0.76], [0.68, 0.34, 0.30, 0.66]].forEach(([bx, by, bw, bh]) => {
    slide.addShape('roundRect', { x: x + bx * s, y: y + by * s, w: bw * s, h: bh * s, rectRadius: 0.05 * s, fill: { type: 'none' }, line });
    slide.addShape('rect', { x: x + (bx + bw / 2 - 0.06) * s, y: y + (by - 0.2) * s, w: 0.12 * s, h: 0.2 * s, fill: { color }, line: { type: 'none' } });
  });
}

function iconBars(slide, x, y, s, color) {
  const line = { color, width: 1.25 };
  [[0.22, 0.35], [0.48, 0.78], [0.74, 0.6]].forEach(([bx, bh]) => {
    slide.addShape('rect', { x: x + bx * s, y: y + (0.88 - bh) * s, w: 0.2 * s, h: bh * s, fill: { type: 'none' }, line });
  });
  slide.addShape('line', { x: x + 0.12 * s, y: y + 0.02 * s, w: 0, h: 0.86 * s, line });
  slide.addShape('line', { x: x + 0.12 * s, y: y + 0.88 * s, w: 0.88 * s, h: 0, line });
  for (let i = 1; i <= 4; i++) {                       // axis ticks
    slide.addShape('line', { x: x + 0.04 * s, y: y + (0.02 + i * 0.17) * s, w: 0.08 * s, h: 0, line });
  }
}

function iconGears(slide, x, y, s, color) {
  const dash = { color, width: 1, dashType: 'dash' };
  // three gears in a triangle, joined by short dashed links between their rims
  slide.addShape('line', { x: x + 0.30 * s, y: y + 0.40 * s, w: 0.10 * s, h: 0.16 * s, line: dash });
  slide.addShape('line', { x: x + 0.60 * s, y: y + 0.40 * s, w: 0.10 * s, h: 0.16 * s, line: dash, flipH: true });
  slide.addShape('line', { x: x + 0.42 * s, y: y + 0.75 * s, w: 0.16 * s, h: 0, line: dash });
  [[0.3, 0.0], [0.0, 0.55], [0.6, 0.55]].forEach(([gx, gy]) => {
    slide.addShape('gear6', { x: x + gx * s, y: y + gy * s, w: 0.4 * s, h: 0.4 * s, fill: { type: 'none' }, line: { color, width: 1.25 } });
    slide.addShape('ellipse', { x: x + (gx + 0.14) * s, y: y + (gy + 0.14) * s, w: 0.12 * s, h: 0.12 * s, fill: { type: 'none' }, line: { color, width: 1 } });
  });
}

function iconBriefcase(slide, x, y, s, color) {
  const line = { color, width: 1.5 };
  slide.addShape('roundRect', { x, y: y + 0.22 * s, w: s, h: 0.78 * s, rectRadius: 0.08 * s, fill: { type: 'none' }, line });
  slide.addShape('roundRect', { x: x + 0.3 * s, y: y + 0.02 * s, w: 0.4 * s, h: 0.22 * s, rectRadius: 0.04 * s, fill: { type: 'none' }, line });
  slide.addShape('rect', { x: x + 0.36 * s, y: y + 0.46 * s, w: 0.28 * s, h: 0.3 * s, fill: { type: 'none' }, line });
  slide.addShape('line', { x, y: y + 0.42 * s, w: s, h: 0, line });
}

/* ------------------------------------------------------------ reusable layout */
// Slides 8 / 13 / 18 / 23 are the same "section cover" with different wording.
function sectionCover(slide, { kicker, title, titleBox, num }) {
  band(slide, { x: 0, y: 2.782, w: 13.333, h: 3.741, fill: C.blue });
  slant(slide, { x: 6.071, y: -0.21, w: 1.545, h: 1.056, adj: 24821, fill: C.ink });
  body(slide, [b(kicker), t('Easy Customizable and editable  design', { ...NL, fontFace: F.med })],
    { x: 0.704, y: 1.464, w: 2.095, h: 0.7 });
  text(slide, [b(title)], { x: 1.695, ...titleBox, fontSize: 38, fontFace: F.xbold, color: C.white });
  body(slide, [b('Established'), t('2019', { ...NL, fontFace: F.med })],
    { x: 0.704, y: 5.8, w: 1.005, h: 0.459, color: C.white });
  chrome(slide, C.ink, num);
}

// Slides 15 / 16 share the "01 / 02 / 03 feature column" on the right edge.
function featureColumn(slide, prefix, colors) {
  const rows = [
    { y: 1.348, x: 8.416, runs: [t('A peep at some distant orb has power to raise and purify our thoughts like. Our thoughts like a strain of sacred music, or a noble picture.', END)] },
    { y: 3.236, x: 7.84, runs: [t('A dazzling spot marked one of its poles, and the rest of its visible. '), t('A peep at some distant orb has power to raise and purify our thoughts like.', END)] },
    { y: 5.11, x: 7.264, runs: [t('Sacred music, or a noble picture, or a passage from the grander. Some distant orb has power to raise and purify our thoughts like. Sacred music.', END)] }
  ];
  rows.forEach((row, i) => {
    body(slide, [b(`${prefix} 0${i + 1}`, END), ...row.runs],
      { x: row.x, y: row.y, w: 4.235, h: 1.081, color: colors[i] });
  });
}

/* --------------------------------------------------------------- deck content */
const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_16x9';
pptx.defineLayout({ name: 'FRESHNO', width: 13.333, height: 7.5 });
pptx.layout = 'FRESHNO';
pptx.author = 'Freshno';
pptx.title = 'Freshno Multipurpose Presentation Template';

/* --- 1. Title slide ------------------------------------------------------- */
function slide01() {
  const s = pptx.addSlide();
  s.background = { color: C.blue };
  body(s, [b('Multipurpose'), t('Presentation Templates', { ...NL, fontFace: F.med })],
    { x: 2.399, y: 6.083, w: 2.072, h: 0.459, color: C.white });
  body(s, [b('Established'), t('2019', { ...NL, fontFace: F.med })],
    { x: 0.697, y: 6.083, w: 1.005, h: 0.459, color: C.white });
  slant(s, { x: 7.741, y: -0.017, w: 2.006, h: 7.517, adj: 94335, fill: C.white });
  text(s, [t('01 . 07 . 09', { fontFace: F.med })],
    { x: 2.399, y: 0.996, w: 2.156, h: 0.303, fontSize: 18, color: C.white });
  slant(s, { x: 2.096, y: 2.105, w: 5.32, h: 3.29, adj: 25163, fill: C.white });
  slant(s, { x: 6.576, y: 2.105, w: 4.406, h: 3.29, adj: 25163, fill: C.white });
  text(s, [b('FRESHNO')], { x: 3.754, y: 3.094, w: 5.826, h: 1.313, fontSize: 78, fontFace: F.xbold, color: C.blue });
  chrome(s, C.white);
}

/* --- 2. Welcome message (full-bleed blue) --------------------------------- */
function slide02() {
  const s = pptx.addSlide();
  s.background = { color: C.blue };
  body(s, [
    b('Based on master slides'),
    t('A dazzling spot marked one of its poles, and the rest of its visible surface was mottled with ruddy and greenish tints which faded into white at the rim', { ...NL, ...END }),
    EMPTY()
  ], { x: 4.125, y: 3.289, w: 5.084, h: 1.321, color: C.white, align: 'center' });
  text(s, [b('WELCOME MESSAGE')],
    { x: 3.697, y: 2.545, w: 5.939, h: 0.303, fontSize: 18, fontFace: F.xbold, color: C.white, align: 'center' });
  slant(s, { x: 9.827, y: 3.603, w: 0.218, h: 0.692, adj: 77806, fill: C.white });
  slant(s, { x: 3.289, y: 3.603, w: 0.218, h: 0.692, adj: 77806, fill: C.white });
  chrome(s, C.white, 2);
}

/* --- 3. Welcome message (blue band) --------------------------------------- */
function slide03() {
  const s = pptx.addSlide();
  band(s, { x: 0.731, y: 2.199, w: 11.872, h: 3.103, fill: C.blue });
  body(s, [
    b('Greeting Text'),
    t('White at the rim, ', NL),
    t('a peep at some distant orb has power to raise and purify our thoughts like a strain of sacred music, or a noble picture, or a passage from the grander poets. ')
  ], { x: 1.758, y: 3.159, w: 3.882, h: 1.181, color: C.white });
  body(s, [b('Moraso'), b(' '), b('Monono'), b(' CEO'), t('Detail you need to share', NL)],
    { x: 7.606, y: 5.622, w: 2.256, h: 0.459, align: 'right' });
  text(s, [b('WELCOME MESSAGE')],
    { x: 1.758, y: 2.665, w: 5.939, h: 0.303, fontSize: 18, fontFace: F.xbold, color: C.white });
  chrome(s, C.ink, 3);
}

/* --- 4. About us (blue slab) ---------------------------------------------- */
function slide04() {
  const s = pptx.addSlide();
  slant(s, { x: -1.586, y: -0.026, w: 8.743, h: 5.974, adj: 24821, fill: C.blue });
  body(s, [
    b('Subtittle', END),
    t('A dazzling spot marked one of its poles, and the rest of its visible surface was mottled with ruddy and greenish tints which faded into white at the rim', END),
    t('A peep at some distant orb has power to raise and purify our thoughts like a strain of sacred music, or a noble picture, or a passage from the grander poets. It always does one good.', END),
    EMPTY()
  ], { x: 0.721, y: 2.518, w: 3.779, h: 2.805, color: C.white });
  head(s, { x: 0.721, y: 1.264, w: 1.023 }, [['ABOUT US', C.white]]);
  slant(s, { x: 6.522, y: 4.867, w: 1.661, h: 1.135, adj: 24821, fill: C.ink });
  slant(s, { x: 7.897, y: 4.867, w: 1.947, h: 1.135, adj: 24821, fill: C.white });
  caption(s, { x: 8.346, y: 5.148, w: 1.183, h: 0.573 }, 'Photo Text', 'Rest of its visible surface was');
  text(s, [b('01')], { x: 7.05, y: 5.115, w: 0.694, h: 0.64, fontSize: 38, fontFace: F.xbold, color: C.white });
  chrome(s, C.ink, 4);
}

/* --- 5. About us (photo grid) --------------------------------------------- */
function slide05() {
  const s = pptx.addSlide();
  slant(s, { x: 10.41, y: -0.015, w: 0.817, h: 2.782, adj: 82661, fill: C.gray95 });
  body(s, [t('A dazzling spot marked one of its poles, and the rest of its visible surface was mottled with ruddy and greenish tints which faded into white at the rim')],
    { x: 6.998, y: 2.067, w: 4.786, h: 0.7 });
  body(s, [t('A peep at some distant orb has power to raise and purify our thoughts like a strain of sacred music, or a noble picture, or a passage from the grander poets. It always does one good. ')],
    { x: 0.62, y: 5.369, w: 4.616, h: 1.042, margin: [7.2, 3.6, 7.2, 3.6] });
  head(s, { x: 6.998, y: 1.264, w: 1.023 }, [['ABOUT ', C.ink], ['US', C.blue]]);
  caption(s, { x: 1.483, y: 3.562, w: 1.925, h: 0.376 }, 'Photo Text 01', 'Rest of its visible surface was');
  chrome(s, C.ink, 5);
}

/* --- 6. Three angled photo columns ---------------------------------------- */
function slide06() {
  const s = pptx.addSlide();
  caption(s, { x: 5.5, y: 0.781, w: 1.394, h: 0.573 }, 'Photo Text 02', 'Rest of its visible here surface was');
  slant(s, { x: 9.327, y: 4.769, w: 2.145, h: 1.466, adj: 24821, fill: C.ink });
  caption(s, { x: 9.934, y: 5.205, w: 1.365, h: 0.573 }, 'Photo Text', 'Rest of its visible surface was', C.white);
  caption(s, { x: 1.071, y: 5.778, w: 1.925, h: 0.376 }, 'Photo Text 01', 'Rest of its visible surface was');
  head(s, { x: 8.613, y: 5.569, w: 0.447, h: 0.303 }, [['03', C.ink]]);
  chrome(s, C.ink, 6);
}

/* --- 7. About our history (timeline band) --------------------------------- */
function slide07() {
  const s = pptx.addSlide();
  slant(s, { x: 3.231, y: -0.015, w: 0.817, h: 2.782, adj: 82661, fill: C.gray95 });
  body(s, [
    t('A dazzling spot marked one of its poles, and the rest of its visible surface was mottled with ruddy and '),
    t('greenis'), t('.')
  ], { x: 0.765, y: 2.067, w: 4.004, h: 0.7 });
  band(s, { x: 0.731, y: 4.321, w: 11.872, h: 1.846, fill: C.blue });

  const eras = [
    { x: 1.259, labelX: 1.259, labelW: 2.048, from: '2006 - ', to: '2009',
      copy: ['A dazzling spot marked one of its poles, and the rest of its visible surface was mottled with ruddy and greenish tints which faded into white at the rim', '.'] },
    { x: 5.016, labelX: 5.054, labelW: 1.869, from: '2012 - ', to: '2014',
      copy: ['PLACEHOLDER', '.'] },
    { x: 8.913, labelX: 8.887, labelW: 1.497, from: '2016 - ', to: '2019',
      copy: ['Distant orb has power to raise and purify our thoughts like a strain of sacred music, or a noble picture, or a ', 'greenish tints which faded into white at the rim', '.'] }
  ];
  eras.forEach(era => {
    body(s, [b('Subtittle', { color: C.blue, ...END }), EMPTY(), ...era.copy.map(p => t(p))],
      { x: era.x, y: 3.89, w: 2.882, h: 1.943, color: C.white });
    head(s, { x: era.labelX, y: 3.429, w: era.labelW, h: 0.303 }, [[era.from, C.ink], [era.to, C.blue]]);
  });
  head(s, { x: 0.765, y: 1.264, w: 1.869 }, [['ABOUT ', C.ink], ['OUR HISTORY', C.blue]]);
  chrome(s, C.ink, 7);
}

/* --- 8 / 13 / 18 / 23. Section covers ------------------------------------- */
const ONE_LINE_TITLE = { y: 4.193, w: 2.669, h: 0.64 };
const TWO_LINE_TITLE = { y: 3.763, w: 3.344, h: 1.279 };
function slide08() { sectionCover(pptx.addSlide(), { kicker: 'Timeline master slides', title: 'TIMELINE', titleBox: ONE_LINE_TITLE, num: 8 }); }
function slide13() { sectionCover(pptx.addSlide(), { kicker: 'Service master slides', title: 'OUR SERVICE', titleBox: { ...TWO_LINE_TITLE, w: 2.669 }, num: 13 }); }
function slide18() { sectionCover(pptx.addSlide(), { kicker: 'Portfolio master slides', title: 'OUR PORT FOLIO', titleBox: TWO_LINE_TITLE, num: 18 }); }
function slide23() { sectionCover(pptx.addSlide(), { kicker: 'Pricelist master slides', title: 'OUR PRICE LIST', titleBox: TWO_LINE_TITLE, num: 23 }); }

/* --- 9. Timeline detail ---------------------------------------------------- */
function slide09() {
  const s = pptx.addSlide();
  slant(s, { x: 1.542, y: 4.945, w: 2.145, h: 1.466, adj: 24821, fill: C.blue });
  caption(s, { x: 2.15, y: 5.381, w: 1.365, h: 0.573 }, 'Photo Text 02', 'Rest of its visible surface was', C.white);
  body(s, [
    b('Subtittle'), b(' Timeline Detail', END),
    t('A dazzling spot marked one of its poles, and the rest of its visible surface was mottled with ruddy and greenish tints which faded into white at the rim, '),
    t('peep at some distant orb has power to raise and purify our thoughts like a strain of sacred music.', END),
    EMPTY()
  ], { x: 7.494, y: 4.409, w: 4.468, h: 1.943 });
  slant(s, { x: 4.011, y: 0.917, w: 3.245, h: 2.747, adj: 24821, fill: C.blue });
  caption(s, { x: 4.608, y: 2.752, w: 1.365, h: 0.573 }, 'Photo Text 01', 'Rest of its visible surface was', C.white);
  slant(s, { x: 7.494, y: 7.097, w: 1.661, h: 1.135, adj: 24821, fill: C.ink });
  caption(s, { x: 11.019, y: 2.752, w: 1.365, h: 0.573 }, 'Photo Text 03', 'Rest of its visible surface was');
  chrome(s, C.ink, 9);
}

/* --- 10. Timeline 01 / 02 / 03 -------------------------------------------- */
function slide10() {
  const s = pptx.addSlide();
  slant(s, { x: 9.769, y: 1.34, w: 2.852, h: 1.949, adj: 24821, fill: C.gray95 });
  slant(s, { x: 1.281, y: 2.314, w: 4.101, h: 2.803, adj: 24821, fill: C.gray95 });
  slant(s, { x: 6.627, y: 4.36, w: 2.852, h: 1.949, adj: 24821, fill: C.gray95 });

  // the source deck stacks two identical copies of the 01 and 03 blocks
  const block01 = () => body(s, [
    b('01', { fontSize: 18, ...END }),
    t('A peep at some distant orb has power to '), t('rais'), t(' e and purify our thoughts like. '),
    t('A dazzling spot marked one of its poles, and the rest of its visible surface was mottled with ruddy and greenish tints which faded into white at the rim', END),
    EMPTY()
  ], { x: 0.712, y: 2.314, w: 3.827, h: 2.096 });
  const block03 = () => body(s, [
    b('03', { fontSize: 18, ...END }),
    t('Sacred music, or a noble picture, or a passage from the grander '),
    t('PLACEHOLDER')
  ], { x: 9.079, y: 1.264, w: 3.136, h: 1.715 });

  block01();
  block03();
  head(s, { x: 0.721, y: 1.264, w: 1.6, h: 0.303 }, [['TIME', C.ink], ['LINE', C.blue]]);
  slant(s, { x: 7.276, y: 1.34, w: 0.6, h: 1.949, adj: 81653, fill: C.blue });
  block01();
  body(s, [b('02', { fontSize: 18, ...END }), t('A dazzling spot marked one of its poles, and the rest of its visible.')],
    { x: 5.774, y: 4.776, w: 3.136, h: 0.993 });
  block03();
  chrome(s, C.ink, 10);
}

/* --- 11. Timeline (three slabs) ------------------------------------------- */
function slide11() {
  const s = pptx.addSlide();
  slant(s, { x: 9.486, y: 0.545, w: 5.43, h: 5.867, adj: 27182, fill: C.gray95 });
  slant(s, { x: -3.574, y: 0.545, w: 5.43, h: 5.867, adj: 27182, fill: C.gray95 });
  slant(s, { x: 3.951, y: 0.545, w: 5.43, h: 5.867, adj: 27182, fill: C.blue });
  head(s, { x: 0.721, y: 1.264, w: 1.6, h: 0.303 }, [['TIME', C.ink], ['LINE', C.blue]]);

  const cols = [
    { x: 0.712, h: 1.181, color: C.ink, num: '01',
      runs: [t('A peep at some distant orb has power to raise and purify our thoughts like. Our thoughts like a strain of sacred music, or a noble picture, or a passage from the grander poets. ')] },
    { x: 5.099, h: 0.941, color: C.white, num: '02',
      runs: [t('A dazzling spot marked one of its poles, and the rest of its visible. '), t('A peep at some distant orb has power to raise and purify our thoughts like.')] },
    { x: 9.486, h: 1.181, color: C.ink, num: '03',
      runs: [t('Sacred music, or a noble picture, or a passage from the grander. Some distant orb has power to raise and purify our thoughts like. Sacred music, or a noble picture, or a passage.')] }
  ];
  cols.forEach(col => {
    body(s, col.runs, { x: col.x, y: 3.621, w: 3.136, h: col.h, color: col.color });
    text(s, [b(col.num)], { x: col.x, y: 2.743, w: 0.694, h: 0.64, fontSize: 38, fontFace: F.xbold, color: col.color });
  });
  chrome(s, C.ink, 11);
}

/* --- 12. Three numbered photo tiles --------------------------------------- */
function slide12() {
  const s = pptx.addSlide();
  s.background = { color: C.white };
  [[1.444, 1.627, '01', 0.472, 2.609], [4.764, 4.914, '02', 0.523, 6.038], [8.139, 8.299, '03', 0.523, 9.329]]
    .forEach(([slantX, numX, num, numW, capX]) => {
      slant(s, { x: slantX, y: 4.769, w: 0.838, h: 0.573, adj: 24821, fill: C.blue });
      text(s, [b(num)], { x: numX, y: 4.854, w: numW, h: 0.404, fontSize: 18, color: C.white, margin: null, wrap: false });
      caption(s, { x: capX, y: 5.205, w: 1.365, h: 0.573 }, 'Photo Text', 'Rest of its visible surface was');
    });
  chrome(s, C.ink, 12);
}

/* --- 14. Our service (blue wedge) ----------------------------------------- */
function slide14() {
  const s = pptx.addSlide();
  slant(s, { x: -1.081, y: 1.003, w: 4.055, h: 2.771, adj: 24821, fill: C.blue });
  body(s, [
    b('Subtittle', END),
    t('A dazzling spot marked one of its poles, and the rest of its visible surface was mottled with ruddy and greenish tints which faded into white at the rim', END),
    t('PLACEHOLDER')
  ], { x: 0.682, y: 4.129, w: 3.499, h: 1.943 });
  body(s, [
    b('Service Focus', END),
    b('A peep at some distant ', { fontFace: F.semi, color: C.blue, ...END }),
    t('A peep at some distant orb has power to raise and purify our thoughts like. Our thoughts like a strain of sacred music, or a noble picture.')
  ], { x: 5.28, y: 4.129, w: 3.136, h: 1.702 });
  body(s, [
    b('Product Focus', END),
    b('A dazzling spot marked ', { fontFace: F.semi, color: C.blue, ...END }),
    t('A dazzling spot marked one of its poles, and the rest of its visible. '),
    t('A peep at some distant orb has power to raise and purify our thoughts like.')
  ], { x: 9.515, y: 4.129, w: 3.136, h: 1.702 });
  head(s, { x: 0.721, y: 1.96, w: 1.6 }, [['OUR SERVICE', C.white]]);
  slant(s, { x: 2.974, y: -0.053, w: 1.545, h: 1.056, adj: 24821, fill: C.ink });
  chrome(s, C.ink, 14);
}

/* --- 15. Our service (icon list) ------------------------------------------ */
function slide15() {
  const s = pptx.addSlide();
  slant(s, { x: 8.213, y: -0.906, w: 6.974, h: 7.535, adj: 27182, fill: C.gray95 });
  iconGears(s, 6.029, 5.125, 0.6, C.blue);
  iconBars(s, 6.504, 3.31, 0.6, C.blue);
  iconBottles(s, 7.043, 1.359, 0.6, C.blue);
  body(s, [
    b('Service Focus', END),
    b('A dazzling spot marked one of its poles,', { fontFace: F.semi, ...END }),
    t('PLACEHOLDER', END),
    t('A peep at some distant orb has power to raise and purify our thoughts like a strain of sacred music, or a noble picture.')
  ], { x: 0.682, y: 2.329, w: 4.235, h: 2.324 });
  featureColumn(s, 'Service', [C.ink, C.ink, C.ink]);
  body(s, [b('Note Service', END), t('Sacred music, or a noble picture, or a passage from the grander. Some distant orb has power to raise.')],
    { x: 0.682, y: 5.125, w: 4.235, h: 0.84 });
  head(s, { x: 0.682, y: 1.264, w: 1.6 }, [['OUR SERVICE', C.ink]]);
  chrome(s, C.ink, 15);
}

/* --- 16. Our product (blue ribbon) ---------------------------------------- */
function slide16() {
  const s = pptx.addSlide();
  head(s, { x: 0.682, y: 1.264, w: 1.6 }, [['OUR PRODUCT', C.white]]);
  slant(s, { x: 1.128, y: 2.593, w: 12.993, h: 2.39, adj: 24821, fill: C.blue });
  body(s, [
    b('Product Focus', END),
    t('A dazzling spot marked one of its poles, and the rest of its visible surface was mottled with ruddy and greenish tints which faded into white at the rim', END),
    t('.', END), EMPTY()
  ], { x: 2.081, y: 3.229, w: 4.235, h: 1.843, color: C.white });
  featureColumn(s, 'Product', [C.ink, C.white, C.ink]);
  chrome(s, C.ink, 16);
}

/* --- 17. Our teams -------------------------------------------------------- */
function slide17() {
  const s = pptx.addSlide();
  caption(s, { x: 10.763, y: 5.12, w: 1.365, h: 0.573 }, 'Photo Text', 'Rest of its visible surface was');
  caption(s, { x: 5.823, y: 5.795, w: 1.907, h: 0.376 }, 'Photo Text', 'Rest of its visible surface was');
  slant(s, { x: 2.172, y: 4.684, w: 3.447, h: 0.573, adj: 24821, fill: C.ink });
  caption(s, { x: 2.172, y: 5.471, w: 2.107, h: 0.376 }, 'Photo Text', 'Rest of its visible surface was');
  slant(s, { x: 5.624, y: 5.056, w: 3.794, h: 0.573, adj: 24821, fill: C.ink });
  slant(s, { x: 9.573, y: 4.684, w: 0.838, h: 0.573, adj: 24821, fill: C.ink });
  [[2.355, 4.768, '01', 0.472], [5.775, 5.14, '02', 0.523], [9.734, 4.768, '03', 0.523]].forEach(([x, y, num, w]) => {
    text(s, [b(num)], { x, y, w, h: 0.404, fontSize: 18, color: C.white, margin: null, wrap: false });
  });
  head(s, { x: 0.682, y: 1.264, w: 1.177 }, [['OUR TEAMS', C.ink]]);
  chrome(s, C.ink, 17);
}

/* --- 19. Portfolio grid ---------------------------------------------------- */
function slide19() {
  const s = pptx.addSlide();
  [
    { x: 0.74, y: 2.937, label: 'Photo Text 01', detail: 'A dazzling spot marked one of its' },
    { x: 4.926, y: 2.937, label: 'Photo Text 02', detail: 'Some distant orb has power to raise' },
    { x: 9.423, y: 2.937, label: 'Photo Text 03', detail: 'Rest of its visible surface was' },
    { x: 2.003, y: 5.283, label: 'Photo Text 04', detail: 'Rest of its visible surface was' },
    { x: 6.487, y: 5.283, label: 'Photo Text 05', detail: 'A dazzling spot marked one of its' },
    { x: 10.977, y: 5.283, label: 'Photo Text 06', detail: 'Some distant orb has power to raise' }
  ].forEach(c => caption(s, { x: c.x, y: c.y, w: 1.423, h: 0.573 }, c.label, c.detail));
  slant(s, { x: 1.39, y: 0.586, w: 1.545, h: 1.056, adj: 24821, fill: C.ink });
  slant(s, { x: 8.192, y: 6.096, w: 0.722, h: 0.494, adj: 24821, fill: C.ink });
  slant(s, { x: 11.094, y: 1.149, w: 0.722, h: 0.494, adj: 24821, fill: C.ink });
  chrome(s, C.ink, 19);
}

/* --- 20. The best of our work --------------------------------------------- */
function slide20() {
  const s = pptx.addSlide();
  slant(s, { x: 6.071, y: -0.21, w: 1.545, h: 1.056, adj: 24821, fill: C.ink });
  slant(s, { x: 5.226, y: 3.907, w: 1.339, h: 4.562, adj: 82661, fill: C.gray95 });
  body(s, [b('Photo Text 02'), t('White at the rim, ', NL), t('a peep at.')],
    { x: 9.689, y: 5.615, w: 1.184, h: 0.573, fontSize: 9 });
  body(s, [b('Photo Text 01 '), t('White at the rim, '), t('a peep at visible.')],
    { x: 1.91, y: 0.846, w: 1.184, h: 0.573, fontSize: 9, align: 'right' });
  body(s, [
    b('Subtittle'),
    t('A dazzling spot marked one of its poles ', NL),
    t('distant orb has power to raise and purify our thoughts like  peep at some distant orb has power to raise and purify our thoughts like a strain of sacred music.')
  ], { x: 0.704, y: 5.393, w: 4.373, h: 1.181 });
  text(s, [b('THE BEST OF', END), b('OUR WORK ', { color: C.blue })],
    { x: 0.682, y: 4.354, w: 1.82, h: 0.606, fontSize: 18, fontFace: F.xbold });
  chrome(s, C.ink, 20);
}

/* --- 21. Our region (map) -------------------------------------------------- */
// The map is redrawn as free-form polygons traced from the original artwork.
const MAP_LAND = [[11.24, 0.53], [11.42, 0.7], [11.24, 0.87], [11.29, 1.09], [11.33, 1.13], [11.57, 1.15], [11.52, 1.27], [11.65, 1.7],
   [11.56, 1.79], [11.5, 2.06], [11.54, 2.09], [11.62, 2.08], [11.59, 2.18], [11.62, 2.23], [11.72, 2.23], [11.86, 2.28],
   [11.9, 2.4], [12.0, 2.45], [11.99, 2.58], [12.07, 2.65], [12.18, 2.93], [12.14, 2.96], [11.94, 2.88], [11.59, 2.67],
   [11.5, 2.52], [11.49, 2.36], [11.36, 2.0], [11.31, 1.97], [11.27, 2.01], [11.27, 2.15], [11.14, 2.15], [11.08, 2.23],
   [11.14, 2.46], [11.23, 2.52], [11.23, 2.56], [11.15, 2.68], [11.0, 2.72], [10.95, 2.82], [10.69, 3.05], [10.68, 3.73],
   [10.75, 3.72], [10.8, 3.64], [10.89, 3.72], [10.97, 3.63], [11.15, 3.61], [11.17, 3.48], [11.67, 3.81], [11.63, 3.88],
   [11.67, 3.98], [11.88, 4.09], [11.8, 4.18], [11.39, 3.75], [11.28, 3.75], [11.3, 3.85], [11.48, 4.1], [11.54, 4.43],
   [11.51, 4.86], [11.44, 4.95], [11.31, 5.0], [11.27, 5.08], [11.14, 4.85], [11.25, 4.71], [11.16, 4.4], [11.09, 4.41],
   [10.98, 4.54], [10.91, 4.57], [10.78, 4.5], [10.53, 4.51], [10.17, 4.27], [9.98, 4.3], [9.84, 4.4], [9.82, 4.45],
   [9.88, 4.54], [9.84, 4.74], [8.49, 4.75], [8.45, 4.79], [8.44, 4.96], [8.21, 4.99], [8.16, 5.2], [7.97, 5.21],
   [7.86, 5.15], [7.71, 5.14], [7.33, 5.31], [7.15, 5.26], [7.01, 5.11], [6.73, 5.06], [6.52, 4.67], [6.3, 4.69],
   [6.15, 4.57], [6.12, 4.47], [6.01, 4.41], [5.41, 4.4], [5.39, 4.5], [5.33, 4.51], [5.26, 4.58], [5.27, 4.68],
   [5.22, 4.71], [4.92, 4.56], [4.78, 4.53], [4.69, 4.37], [4.6, 4.31], [4.38, 4.3], [4.28, 4.35], [4.18, 4.31],
   [4.04, 4.43], [4.08, 4.59], [4.06, 4.71], [3.76, 4.78], [3.68, 5.12], [3.62, 5.15], [3.56, 5.13], [3.46, 4.88],
   [3.16, 4.52], [3.09, 4.23], [3.23, 4.21], [3.41, 4.11], [3.53, 4.13], [3.66, 4.02], [3.51, 3.73], [3.55, 3.48],
   [3.45, 3.37], [3.49, 3.31], [3.6, 3.31], [3.62, 3.18], [3.7, 3.06], [3.66, 2.9], [3.78, 2.76], [5.15, 2.75],
   [5.19, 2.72], [5.21, 2.63], [5.3, 2.63], [5.43, 2.45], [5.5, 2.54], [5.39, 2.6], [5.42, 2.68], [5.82, 2.65],
   [5.9, 2.72], [6.13, 2.79], [6.17, 2.74], [6.16, 2.61], [6.48, 2.93], [6.52, 2.64], [6.75, 2.42], [6.79, 2.43],
   [6.84, 2.59], [6.92, 2.54], [6.96, 2.62], [7.12, 2.6], [7.17, 2.47], [7.35, 2.44], [7.41, 2.31], [7.66, 2.16],
   [7.81, 2.14], [7.97, 1.94], [8.07, 2.0], [8.29, 2.04], [8.34, 2.1], [8.35, 2.22], [8.43, 2.31], [8.74, 2.31],
   [8.82, 2.18], [8.96, 2.16], [9.13, 2.32], [9.22, 2.35], [9.24, 2.23], [9.33, 2.21], [9.45, 2.1], [9.41, 1.91],
   [9.3, 1.82], [9.4, 1.82], [9.5, 1.88], [9.72, 1.78], [9.96, 1.76], [10.13, 1.58], [10.32, 1.58], [10.44, 1.4],
   [10.48, 1.2], [10.61, 1.01], [11.07, 0.7], [11.14, 0.59], [11.24, 0.53]
];
const MAP_BLUE = [
  [[10.36, 1.55], [10.47, 1.62], [10.45, 1.83], [10.58, 1.87], [10.62, 1.94], [10.53, 2.04], [10.59, 2.28], [10.42, 2.54],
   [10.54, 2.77], [10.32, 2.86], [10.33, 3.08], [10.41, 3.2], [10.22, 3.44], [10.27, 3.67], [10.32, 3.74], [10.31, 3.84],
   [10.04, 4.0], [9.57, 4.04], [9.38, 3.83], [9.16, 3.77], [9.03, 3.97], [8.82, 4.02], [8.8, 3.82], [8.66, 3.59],
   [8.57, 3.48], [8.43, 3.46], [8.32, 2.91], [8.41, 2.83], [8.49, 2.64], [8.34, 2.34], [8.58, 2.29], [8.73, 2.31],
   [8.8, 2.27], [8.81, 2.18], [8.95, 2.15], [9.13, 2.32], [9.23, 2.35], [9.24, 2.23], [9.32, 2.22], [9.46, 2.11],
   [9.42, 1.99], [9.45, 1.92], [9.71, 1.78], [9.96, 1.77], [10.12, 1.58], [10.27, 1.6], [10.35, 1.55]],
  [[5.63, 2.83], [6.04, 3.01], [6.31, 2.93], [6.33, 3.01], [5.88, 3.23], [5.67, 3.52], [5.27, 3.43], [5.17, 3.49],
   [5.01, 3.49], [5.0, 3.4], [5.22, 3.18], [5.12, 3.14], [5.17, 2.94], [5.4, 3.01], [5.45, 2.85], [5.63, 2.83]],
  [[9.39, 3.96], [9.78, 4.16], [9.88, 4.33], [9.82, 4.44], [9.88, 4.52], [9.88, 4.77], [9.81, 4.85], [9.21, 4.84],
   [9.37, 4.67], [9.36, 4.53], [9.49, 4.38], [9.33, 4.19], [9.42, 4.11], [9.37, 4.02], [9.39, 3.96]],
  [[6.76, 3.95], [7.14, 3.98], [7.2, 4.09], [7.33, 4.11], [7.41, 4.19], [7.4, 4.29], [7.46, 4.38], [7.38, 4.46],
   [7.07, 4.54], [7.03, 4.47], [6.85, 4.45], [6.6, 4.33], [6.54, 4.18], [6.76, 3.95]],
  [[5.21, 2.62], [5.34, 2.78], [5.42, 2.83], [5.41, 2.97], [5.36, 3.0], [5.19, 2.94], [5.1, 3.14], [5.21, 3.18],
   [5.11, 3.29], [5.02, 3.28], [4.57, 2.99], [4.56, 2.92], [4.67, 2.76], [5.17, 2.75], [5.21, 2.62]],
  [[11.4, 4.22], [11.48, 4.27], [11.54, 4.39], [11.54, 4.79], [11.47, 4.94], [11.32, 4.99], [11.27, 5.08], [11.14, 4.84],
   [11.25, 4.74], [11.25, 4.5], [11.38, 4.45], [11.44, 4.34], [11.37, 4.27], [11.4, 4.22]],
  [[6.8, 1.96], [6.82, 1.99], [6.79, 2.03], [6.53, 2.08], [6.19, 2.28], [6.12, 2.36], [6.1, 2.52], [6.0, 2.46],
   [5.98, 2.3], [6.38, 2.03], [6.79, 1.96]]
];
const MAP_DARK = [
  [[8.58, 3.52], [8.78, 3.82], [8.82, 4.03], [9.04, 3.96], [9.11, 3.81], [9.19, 3.78], [9.36, 3.83], [9.43, 3.91],
   [9.36, 3.99], [9.41, 4.11], [9.28, 4.14], [9.21, 4.21], [9.04, 4.23], [8.87, 4.31], [8.92, 4.58], [8.87, 4.74],
   [8.48, 4.75], [8.42, 4.87], [8.01, 4.77], [7.98, 4.73], [8.03, 4.67], [8.08, 4.35], [8.27, 4.31], [8.38, 4.14],
   [8.53, 4.11], [8.52, 3.98], [8.45, 3.89], [8.52, 3.73], [8.49, 3.57], [8.57, 3.52]],
  [[11.54, 1.43], [11.66, 1.7], [11.55, 1.78], [11.5, 2.06], [11.62, 2.24], [11.71, 2.22], [11.86, 2.28], [11.89, 2.39],
   [12.0, 2.45], [11.98, 2.57], [12.07, 2.65], [12.08, 2.74], [12.16, 2.86], [12.16, 2.96], [11.6, 2.68], [11.49, 2.52],
   [11.5, 2.37], [11.35, 1.98], [11.21, 1.87], [11.19, 1.92], [11.26, 1.99], [11.22, 2.0], [10.96, 1.79], [11.13, 1.61],
   [11.29, 1.64], [11.4, 1.6], [11.47, 1.44], [11.54, 1.43]]
];

function polygon(slide, pts, fill) {
  const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
  const x = Math.min(...xs), y = Math.min(...ys);
  const w = Math.max(...xs) - x, h = Math.max(...ys) - y;
  slide.addShape('custGeom', {
    x, y, w, h,
    points: [...pts.map(([px, py]) => ({ x: px - x, y: py - y })), { close: true }],
    fill: { color: fill }, line: { type: 'none' }
  });
}

function slide21() {
  const s = pptx.addSlide();
  polygon(s, MAP_LAND, C.gray85);
  MAP_BLUE.forEach(poly => polygon(s, poly, C.blue));
  MAP_DARK.forEach(poly => polygon(s, poly, C.ink));
  text(s, [b('OUR ', { color: C.ink }), b('REGION', { color: C.blue })],
    { x: 0.682, y: 4.354, w: 1.528, h: 0.606, fontSize: 18, fontFace: F.xbold });
  body(s, [
    b('Subtittle'),
    t('A dazzling spot marked one of its poles ', NL),
    t('distant orb has power to raise and purify our thoughts like  peep at some distant orb has power to raise and purify our thoughts like a strain of sacred music.')
  ], { x: 0.704, y: 5.393, w: 4.373, h: 1.181 });
  slant(s, { x: 2.824, y: 1.142, w: 2.145, h: 1.466, adj: 24821, fill: C.blue });
  caption(s, { x: 3.432, y: 1.578, w: 1.365, h: 0.573 }, 'Map Text 01', 'Rest of its visible surface was', C.white);
  slant(s, { x: 8.579, y: 4.86, w: 2.145, h: 1.466, adj: 24821, fill: C.ink });
  caption(s, { x: 9.09, y: 5.296, w: 1.225, h: 0.573 }, 'Map Text 02', 'Some distant orb has power to raise', C.white);
  chrome(s, C.ink, 21);
}

/* --- 22. About our region (three cities) ---------------------------------- */
function slide22() {
  const s = pptx.addSlide();
  const cities = [
    { x: 0.731, fill: C.gray85, name: 'NOAH CITY', nameX: 1.551, nameW: 2.048, textX: 1.321, textW: 2.51, color: C.ink,
      copy: [t('A dazzling spot marked one of its poles, and the rest of its visible surface was mottled with ruddy and greenish tints which faded into white at the rim'), t('.')] },
    { x: 4.809, fill: C.blue, name: 'REINO CITY', nameX: 5.63, nameW: 2.048, textX: 5.36, textW: 2.613, color: C.white,
      copy: [t('Visible surface was mottled with ruddy and greenish tints which faded into white at the rim mottled with ruddy and greenish tints which faded into.')] },
    { x: 8.887, fill: C.ink, name: 'KARA CITY', nameX: 9.708, nameW: 2.048, textX: 9.462, textW: 2.613, color: C.white,
      copy: [t('Distant orb has power to raise and purify our thoughts like a strain of sacred music, or a noble picture, or a '), t('greenish tints which faded into white at the rim'), t('.')] }
  ];
  cities.forEach(city => band(s, { x: city.x, y: 3.128, w: 3.689, h: 3.282, fill: city.fill }));
  slant(s, { x: 3.405, y: -0.015, w: 0.817, h: 2.782, adj: 82661, fill: C.gray95 });
  cities.forEach(city => {
    body(s, [b('Subtittle', END), ...city.copy],
      { x: city.textX, y: 4.29, w: city.textW, h: 1.562, color: city.color });
    text(s, [b(city.name)],
      { x: city.nameX, y: 3.429, w: city.nameW, h: 0.303, fontSize: 18, fontFace: F.xbold, color: city.color, align: 'center' });
  });
  body(s, [
    t('A dazzling spot marked one of its poles, and the rest of its visible surface was mottled with ruddy and '),
    t('greenis'), t('.')
  ], { x: 0.765, y: 2.067, w: 4.004, h: 0.7 });
  head(s, { x: 0.765, y: 1.264, w: 1.869 }, [['ABOUT OUR', C.ink], [' REGION', C.blue]]);
  body(s, [b('A peep at some distant orb has power to raise and purify.', { fontFace: F.semi })],
    { x: 9.897, y: 2.067, w: 2.721, h: 0.459, align: 'right' });
  chrome(s, C.ink, 22);
}

/* --- 24. Pricelist --------------------------------------------------------- */
function slide24() {
  const s = pptx.addSlide();
  s.background = { color: C.blue };
  slant(s, { x: 9.699, y: -0.009, w: 2.366, h: 7.517, adj: 77806, fill: C.ink });

  const tiers = [
    { name: 'Basic', nameW: 0.549, price: 'USD 250,00', x: 0.704, blurbY: 5.01, blurbW: 3.335, iconX: 1.598, slantX: 2.305, numX: 2.488, numW: 0.472, num: '01',
      blurb: 'A dazzling spot marked one of its poles, and the rest of its visible surface was mottled with ruddy.' },
    { name: 'Business', nameW: 1.046, price: 'USD 280,00', x: 4.906, blurbY: 5.006, blurbW: 3.1, iconX: 5.766, slantX: 6.499, numX: 6.682, numW: 0.523, num: '02',
      blurb: 'A peep at some distant orb has power to raise and purify our thoughts like a strain of sacred music, or a noble picture, or a passage.' },
    { name: 'Profesional', nameW: 1.152, price: 'USD 320,00', x: 8.831, blurbY: 5.006, blurbW: 3.335, iconX: 9.689, slantX: 10.422, numX: 10.605, numW: 0.523, num: '03',
      blurb: 'A dazzling spot marked one of its poles, and the rest of its visible surface was mottled with ruddy.' }
  ];
  tiers.forEach(tier => {
    const blurbX = tier.name === 'Profesional' ? 8.81 : tier.x;
    body(s, [t(tier.blurb)], { x: blurbX, y: tier.blurbY, w: tier.blurbW, h: 0.636, fontSize: 10, color: C.white });
    body(s, [b(tier.name)], { x: tier.x, y: 3.89, w: tier.nameW, h: 0.278, fontSize: 14, color: C.white });
    text(s, [b(tier.price)], { x: tier.x, y: 4.51, w: 2.536, h: 0.303, fontSize: 18, fontFace: F.xbold, color: C.white });
    iconBriefcase(s, tier.iconX, 3.1, 0.58, C.white);
    slant(s, { x: tier.slantX, y: 3.538, w: 0.838, h: 0.573, adj: 24821, fill: C.white });
    text(s, [b(tier.num)], { x: tier.numX, y: 3.633, w: tier.numW, h: 0.404, fontSize: 18, color: C.ink, margin: null, wrap: false });
  });
  head(s, { x: 0.682, y: 1.264, w: 1.472 }, [['OUR PRICELIST', C.white]]);
  chrome(s, C.white, 24);
}

/* --- 25. Statistic progress (four doughnut charts) ------------------------- */
function donut(slide, { cx, cy, values, colors, startAng, label }) {
  slide.addChart('doughnut', [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values }], {
    x: cx - 1.238, y: cy - 0.9785, w: 2.476, h: 1.957,
    holeSize: 75, firstSliceAng: startAng, chartColors: colors,
    dataBorder: { pt: 0, color: C.white },
    showLegend: false, showTitle: false, showValue: false,
    chartArea: { fill: { color: C.white } }
  });
  slide.addText([{ text: label, options: { bold: true } }], {
    x: cx - 0.647, y: cy - 0.231, w: 1.274, h: 0.573,
    fontSize: 28, fontFace: F.reg, color: C.ink, align: 'center', valign: 'middle', margin: 0
  });
}

function slide25() {
  const s = pptx.addSlide();
  s.background = { color: C.white };
  const GRAY = C.gray85;
  donut(s, { cx: 1.950, cy: 3.519, values: [30, 70], colors: [C.blue, GRAY], startAng: 263, label: '22%' });
  donut(s, { cx: 4.927, cy: 3.546, values: [30, 70], colors: [GRAY, C.blue], startAng: 0, label: '68%' });
  donut(s, { cx: 7.905, cy: 3.546, values: [20, 80], colors: [C.blue, GRAY], startAng: 0, label: '12%' });
  donut(s, { cx: 10.882, cy: 3.546, values: [50, 50], colors: [GRAY, C.blue], startAng: 0, label: '50%' });

  [
    { x: 1.167, y: 4.989, w: 2.158, title: 'Statistic Client', copy: 'A peep at some distant orb has power to raise and purify our thoughts.' },
    { x: 4.295, y: 4.96, w: 1.99, title: 'Grow Company', copy: 'A dazzling spot marked one of its poles, and the rest of its visible. ' },
    { x: 7.256, y: 4.96, w: 2.119, title: 'Product Grow', copy: 'Sacred music, or a noble picture, or a passage from the grander.' },
    { x: 10.459, y: 4.96, w: 2.119, title: 'Service Statistic', copy: 'Power to raise or a noble picture, or a passage from the grander.' }
  ].forEach(col => body(s, [b(col.title, END), t(col.copy)], { x: col.x, y: col.y, w: col.w, h: 1.081 }));

  head(s, { x: 0.712, y: 1.264, w: 2.241 }, [['OUR STATISTIC ', C.ink], ['PROGRESS', C.blue]]);
  [3.169, 6.224, 9.119].forEach(x => slant(s, { x, y: 2.575, w: 0.6, h: 1.949, adj: 81653, fill: C.blue }));
  chrome(s, C.ink, 25);
}

/* --- 26. Closing statement ------------------------------------------------- */
function slide26() {
  const s = pptx.addSlide();
  slant(s, { x: 2.559, y: 0.501, w: 12.48, h: 5.867, adj: 27182, fill: C.gray85 });
  slant(s, { x: -2.61, y: 0.501, w: 5.43, h: 5.867, adj: 27182, fill: C.gray85 });
  body(s, [
    b('Closing Statement', { color: C.blue, ...END }),
    t('A dazzling spot marked one of its poles, and the rest of its visible surface was mottled with ruddy and greenish tints which faded into white at the rim', END),
    t('A peep at some distant orb has power to raise and purify our thoughts like a strain of sacred music, or a noble picture.', END),
    EMPTY()
  ], { x: 4.82, y: 2.42, w: 3.158, h: 2.565 });
  body(s, [
    t('A noble picture, or a passage from the grander poets. It always does one good. ', END),
    t('Peep at some distant orb has power to raise and purify our thoughts like. Our thoughts like a strain of sacred music, or a noble picture, or a passage from the grander poets. '),
    t('visible surface was mottled with ruddy and greenish.')
  ], { x: 8.594, y: 2.801, w: 4.057, h: 1.803 });
  text(s, [b('CLOSING STATEMENT')],
    { x: 0.682, y: 2.795, w: 3.523, h: 1.279, fontSize: 38, fontFace: F.xbold, color: C.ink });
  chrome(s, C.ink, 26);
}

/* --- 27. Thank you --------------------------------------------------------- */
function slide27() {
  const s = pptx.addSlide();
  s.background = { color: C.blue };
  text(s, [b('THANK YOU')],
    { x: 4.657, y: 3.43, w: 4.02, h: 0.64, fontSize: 38, fontFace: F.xbold, color: C.white, align: 'center' });
  chrome(s, C.white);
}

[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
 slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18,
 slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27]
  .forEach(build => build());

pptx.writeFile({ fileName: path.join(__dirname, '08d7236f-3965-4319-93c4-9f5a7f8b6677_grok_final.pptx') })
  .then(f => console.log('wrote', f));
