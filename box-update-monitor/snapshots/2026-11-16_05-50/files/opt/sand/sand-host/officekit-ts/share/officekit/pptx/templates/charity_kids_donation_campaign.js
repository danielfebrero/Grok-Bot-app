/*
 * "Social Activity & Charity Donations" — 22-slide charity presentation.
 *
 * Slide size: 26.667 x 15 in (16:9). Every coordinate below is in inches and
 * every colour is a plain hex literal, so the layout can be read straight off
 * the source. Photographs in the original deck live in empty picture
 * placeholders (they ship with no image), so nothing is drawn for them here.
 *
 * Run: node <this file>   ->   writes the .pptx next to the script.
 */
const pptxgen = require('pptxgenjs');
const path = require('path');

const SLIDE_W = 26.6667;
const SLIDE_H = 15;

/* ------------------------------------------------------------------ palette */
const GREEN  = '00B050'; // primary brand green
const DEEP   = '00863D'; // darker green, only used inside the corner motif
const FOREST = '19423B'; // near-black green for the dark panels
const ORANGE = 'F6951D'; // accent
const WHITE  = 'FFFFFF';
const W95    = 'F2F2F2'; // white shaded 5% — body copy on green
const W85    = 'D9D9D9';
const W75    = 'BFBFBF';
const GRAY   = '808080';
const INK    = '262626'; // headings
const INK5   = '0D0D0D';
const INK25  = '404040';
const BODY   = '595959'; // body copy on light backgrounds
const MUTED  = '808080';

/* -------------------------------------------------------------------- fonts */
const SB = 'Poppins SemiBold';
const MD = 'Poppins Medium';
const LT = 'Poppins Light';
const BD = 'Poppins Bold';
const RB = 'Roboto Medium';

/* ------------------------------------------------------------- text presets */
const TITLE   = { fontSize: 72, fontFace: SB };
const TITLE80 = { fontSize: 80, fontFace: SB };
const TITLE66 = { fontSize: 66, fontFace: SB };
const HEAD60  = { fontSize: 60, fontFace: SB };
const HEAD48  = { fontSize: 48, fontFace: SB };
const HEAD40  = { fontSize: 40, fontFace: SB };
const HEAD32  = { fontSize: 32, fontFace: MD };
const STAT32  = { fontSize: 32, fontFace: LT };
const PCT36   = { fontSize: 36, fontFace: MD, lineSpacingMultiple: 1.5 };
const LABEL   = { fontSize: 28, fontFace: MD };
const LABELB  = { fontSize: 28, fontFace: SB };
const LABEL15 = { fontSize: 28, fontFace: MD, lineSpacingMultiple: 1.5 };
const SMALL   = { fontSize: 24, fontFace: MD };
const SMALLB  = { fontSize: 24, fontFace: SB };
const TINY    = { fontSize: 20, fontFace: MD };
const PARA    = { fontSize: 24, fontFace: RB, lineSpacingMultiple: 1.5 };

/* ------------------------------------------------------------------ helpers */

// All text boxes in this deck are top-anchored, grow to fit, and use
// PowerPoint's default 0.1 x 0.05 in insets (expressed in points here).
function text(s, body, x, y, w, h, opts) {
  s.addText(body, Object.assign(
    { x, y, w, h, valign: 'top', fit: 'resize', margin: [7.2, 7.2, 3.6, 3.6] }, opts));
}

// Closed polygon from [x, y] pairs measured inside the shape's own box.
function poly(s, x, y, w, h, pts, opts) {
  s.addShape('custGeom', Object.assign({
    x, y, w, h,
    points: pts.map(([px, py], i) => ({ x: px, y: py, moveTo: i === 0 }))
      .concat([{ close: true }]),
  }, opts));
}

// Fully rounded ("stadium") bar.
function stadium(s, x, y, w, h, color) {
  s.addShape('roundRect', { x, y, w, h, fill: { color }, rectRadius: Math.min(w, h) / 2 });
}

// Single outlined circle — the deck's decorative motif.
function ring(s, x, y, d, color, lw) {
  s.addShape('ellipse', { x, y, w: d, h: d, line: { color, width: lw } });
}

// The pair of overlapping outlined circles used as a logo mark everywhere.
function rings(s, x1, y1, x2, y2, d, lw, c1, c2) {
  ring(s, x1, y1, d, c1, lw);
  ring(s, x2, y2, d, c2, lw);
}

// Green square with a darker triangle over its top-left half; sits in a corner.
function corner(s, x, y, w, h) {
  s.addShape('rect', { x, y, w, h, fill: { color: GREEN } });
  poly(s, x, y, w, h, [[0, 0], [w, 0], [0, h]], { fill: { color: DEEP } });
}

// Drop shadow used by the deck's floating cards.
function softShadow(blur, opacity) {
  return { type: 'outer', blur, offset: 4, angle: 90, color: '000000', opacity };
}

/* ------------------------------------------------------- slide 1: title page */
function slide01(pres) {
  const s = pres.addSlide();
  // Empty photo frame: 16.12 x 11.45 at (10.56, 0).
  // Four slanted panels build the diagonal split.
  poly(s, 0, 0, 16.062, 11.45, [[0, 0], [10.636, 0], [16.062, 11.45], [0, 11.45]],
    { fill: { color: FOREST } });
  poly(s, 0, 7.881, 12.095, 3.569, [[0, 0.032], [10.318, 0], [12.095, 3.569], [0, 3.569]],
    { fill: { color: GREEN } });
  poly(s, 16.062, 11.45, 10.605, 3.55, [[0, 0], [8.778, 0], [10.605, 3.55], [0, 3.55]],
    { fill: { color: FOREST }, flipH: true, flipV: true });
  poly(s, 6.433, 0, 17.949, 15, [[0, 0], [10.66, 0], [17.949, 15], [7.411, 15]],
    { fill: { color: GREEN, transparency: 78 } });

  text(s, 'Donate Now For Kids Future', 1.709, 1.721, 5.742, 0.572, { ...LABEL, color: ORANGE });
  text(s, [{ text: 'Social Activity & ' },
           { text: 'Charity Donations !', options: { color: ORANGE } }],
    1.709, 2.262, 11.853, 2.928, { ...TITLE, color: WHITE, lineSpacingMultiple: 1.2 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore,',
    1.709, 5.388, 9.518, 1.245, { ...PARA, color: W75 });

  // Three counters on the green band, separated by hairlines.
  [['17k', 'Donations', 1.709, 1.961, 1.992, 9.982],
   ['7M', 'Collecting', 4.658, 1.961, 2.541, 9.913],
   ['28k', 'Giving House', 7.713, 2.398, 2.523, 9.913],
  ].forEach(([value, caption, x, numW, capW, capY]) => {
    text(s, value, x, 8.874, numW, 1.111, { ...HEAD60, color: WHITE });
    text(s, caption, x, capY, capW, 0.505, { ...SMALL, color: WHITE });
  });
  [4.244, 7.189].forEach((x) => s.addShape('line',
    { x, y: 8.874, w: 0, h: 1.544, line: { color: WHITE, transparency: 8, width: 0.5 } }));

  // Footer strip.
  text(s, 'Save Life For Kids', 1.709, 12.887, 3.727, 0.574, { ...LABEL, color: ORANGE });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore,',
    5.725, 12.552, 9.518, 1.245, { ...PARA, color: BODY });
  text(s, 'LOGO ', 18.994, 12.72, 2.123, 0.909, { fontSize: 48, fontFace: BD, color: WHITE, bold: true });
  text(s, 'Company Name', 21.387, 12.909, 3.275, 0.505, { ...SMALL, color: ORANGE });

  rings(s, 9.422, 1.364, 9.748, 1.203, 0.613, 6, WHITE, ORANGE);
  rings(s, 24.361, 10.913, 24.072, 11.145, 0.824, 8.5, FOREST, ORANGE);
  ring(s, 0.989, 7.29, 1.064, ORANGE, 8.5);
  ring(s, 6.821, 10.81, 0.824, WHITE, 8.5);
  ring(s, 6.054, 11.136, 0.824, GREEN, 8.5);
  corner(s, 0, 14.38, 0.613, 0.62);
}

/* --------------------------------------------- slide 2: about + progress bars */
function slide02(pres) {
  const s = pres.addSlide();
  // Empty photo frame: 10.52 x 12.90 at (2.09, 2.10).
  s.addShape('rect', { x: 0, y: 0, w: 19.355, h: 8.355, fill: { color: GREEN, transparency: 90 } });
  s.addShape('snip1Rect', { x: 12.581, y: 8.355, w: 12.871, h: 6.644, fill: { color: FOREST } });

  text(s, '- About Us -', 13.962, 2.102, 2.88, 0.572, { ...LABEL, color: GREEN });
  text(s, [{ text: 'From Conation To ' }, { text: 'Concrete Help', options: { color: ORANGE } }],
    13.962, 2.768, 9.48, 2.524, { ...TITLE, color: INK });
  text(s, 'Lorem ipsum dolor sit a met, connecter adipescent elite, sed do eiusmod temper incident Ut labore et dolore magna aliquant. Ut denim ad minim venial, ',
    13.962, 5.386, 9.942, 1.851, { ...PARA, color: BODY });

  text(s, [{ text: 'Charity Needs', options: { underline: { style: 'sng' } } }],
    13.962, 9.826, 3.318, 0.572, { ...LABEL, color: ORANGE });
  // [caption, caption y, percent, percent y, bar y, bar width]
  [['3 Millions We Need', 10.873, '40%', 10.811, 11.684, 4.39],
   ['4 Thousand Homeless Kids', 12.53, '70%', 12.468, 13.317, 7.65],
  ].forEach(([caption, capY, pct, pctY, barY, barW]) => {
    text(s, caption, 13.962, capY, 5.779, 0.572, { ...LABEL, color: WHITE });
    text(s, pct, 22.356, pctY, 1.182, 0.64, { ...HEAD32, color: ORANGE });
    s.addShape('rect', { x: 13.962, y: barY, w: barW, h: 0.151, fill: { color: ORANGE } });
  });

  rings(s, 8.343, 1.349, 8.809, 1.561, 1.207, 8.5, GREEN, WHITE);
  rings(s, 20.418, 7.93, 20.884, 8.142, 0.824, 8.5, GREEN, ORANGE);
  corner(s, 25.677, 0, 0.989, 1);
  corner(s, 0, 14.314, 0.677, 0.685);
}

/* --------------------------------------------- slide 3: foodless people cases */
function slide03(pres) {
  const s = pres.addSlide();
  s.addShape('rect', { x: 0, y: 8.982, w: 10.911, h: 4.094, fill: { color: GREEN, transparency: 92 } });

  // Eight stadium bars: [bar y, bar width, percent, percent x, percent y].
  const BARS = [
    [3.406, 6.442, '40%', 18.868, 3.203], [4.484, 10.911, '90%', 23.129, 4.375],
    [5.562, 8.8, '80%', 20.987, 5.453], [6.641, 10.911, '85%', 23.129, 6.469],
    [7.719, 7.567, '60%', 20.07, 7.617], [8.797, 6.442, '30%', 18.868, 8.656],
    [9.875, 10.911, '90%', 23.129, 9.805], [10.953, 9.536, '80%', 21.813, 10.898],
  ];
  BARS.forEach(([y, w, pct, pctX, pctY]) => {
    stadium(s, 11.996, y, w, 0.906, GREEN);
    text(s, pct, pctX, pctY, 1.388, 0.924, { ...PCT36, color: INK25 });
  });

  text(s, '- Hold hands & Give New Life -', 2.137, 2.413, 7.733, 0.572, { ...LABEL, color: GREEN });
  text(s, [{ text: 'Our Foodless People ' }, { text: 'Cases', options: { color: ORANGE } }],
    2.137, 3.102, 7.508, 2.524, { ...TITLE, color: INK, bold: true });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim venom,',
    2.137, 5.738, 7.567, 2.457, { ...PARA, color: MUTED });
  text(s, 'Donate Now For Kids Future', 2.137, 9.805, 5.736, 0.569, { ...LABELB, color: GREEN });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua,',
    2.137, 10.45, 7.567, 1.851, { ...PARA, color: MUTED });

  rings(s, 8.992, 1.713, 9.457, 1.925, 0.824, 8.5, GREEN, ORANGE);
  rings(s, 22.875, 12.863, 23.341, 13.075, 0.824, 8.5, GREEN, ORANGE);
}

/* -------------------------------------------------- slide 4: about foundation */
function slide04(pres) {
  const s = pres.addSlide();
  // Empty photo frame: 15.38 x 9.18 at (11.29, 0.01).
  s.addShape('rect', { x: 0, y: 7.983, w: 13.806, h: 7.016, fill: { color: GREEN } });
  s.addShape('rect', { x: 8.456, y: 0, w: 9.48, h: 15.003, fill: { color: GREEN, transparency: 86 } });

  text(s, '- Fund Progress -', 2.292, 1.671, 4.006, 0.572, { ...LABEL, color: GREEN });
  text(s, [{ text: 'About Kindness', options: { breakLine: true } },
           { text: 'Foundation', options: { color: ORANGE } }],
    2.292, 2.356, 9.48, 2.524, { ...TITLE, color: INK5 });
  text(s, 'Lorem ipsum dolor sit a met, connecter adipescent elite, sed do eiusmod temper incident Ut labore et dolore magna aliquant.',
    2.292, 5.046, 7.869, 1.851, { ...PARA, color: BODY });

  text(s, 'Get Every General Answers ', 2.292, 9.517, 6.643, 0.651, { ...HEAD32, color: WHITE });
  text(s, 'Lorem ipsum dolor sit a met, connecter adipescent elite, sed do eiusmod temper incident Ut labore et dolore magna aliquant.',
    2.292, 10.433, 7.869, 1.851, { ...PARA, color: W85 });
  text(s, 'Read More', 2.292, 12.959, 2.38, 0.572, { ...LABEL, color: WHITE });
  s.addShape('triangle', { x: 4.999, y: 13.083, w: 0.376, h: 0.324, fill: { color: WHITE }, rotate: 90 });

  // Bulleted answer list on the right.
  [10.462, 11.345, 12.229, 13.09].forEach((y, i) => {
    s.addShape('ellipse',
      { x: 15.62, y: [10.729, 11.527, 12.549, 13.431][i], w: 0.169, h: 0.169, fill: { color: DEEP } });
    text(s, 'Lorem ipsum dolor sit a met, connecter adipescent elite,',
      16.238, y, 8.732, 0.64, { ...PARA, color: BODY });
  });

  rings(s, 1.354, 7.6, 1.796, 7.499, 0.902, 6, WHITE, ORANGE);
  corner(s, 25.989, 14.314, 0.677, 0.685);
}

/* -------------------------------- slides 5 & 12: reusable donation case cards */

// Wide card (slide 5): white header + tinted amount, green stats row, orange CTA.
function wideCaseCard(s, x, y) {
  s.addShape('rect', { x, y, w: 10.753, h: 2.01, fill: { color: WHITE }, shadow: softShadow(28, 0.26) });
  s.addShape('rect', { x: x + 7.025, y: y - 0.002, w: 3.752, h: 1.968, fill: { color: GREEN, transparency: 86 } });
  s.addShape('rect', { x, y: y + 1.971, w: 10.753, h: 2.01, fill: { color: GREEN } });
  s.addShape('rect', { x: x + 7.025, y: y + 1.971, w: 3.729, h: 2.007, fill: { color: ORANGE } });

  text(s, [{ text: 'Contribute For The Abled ' }, { text: 'Child Cancer', options: { color: ORANGE } }],
    x + 0.802, y + 0.495, 5.399, 1.043, { ...LABEL, color: INK });
  text(s, '$50k', x + 7.575, y + 0.63, 2.628, 0.774, { ...HEAD40, color: FOREST, align: 'center' });
  text(s, 'Achieved', x + 0.802, y + 2.316, 2.088, 0.572, { ...LABELB, color: WHITE });
  text(s, '0%', x + 4.665, y + 2.248, 0.997, 0.639, { ...STAT32, color: WHITE, align: 'right' });
  text(s, 'Target', x + 0.802, y + 2.955, 2.088, 0.572, { ...LABELB, color: WHITE });
  text(s, '$50.000', x + 3.574, y + 2.888, 2.088, 0.639, { ...STAT32, color: WHITE, align: 'right' });
  text(s, 'Donate', x + 7.575, y + 2.624, 2.38, 0.572, { ...LABELB, color: WHITE, align: 'center' });
}

// Narrow card (slide 12): orange CTA floats above a white header + green stats.
function tallCaseCard(s, x) {
  s.addShape('rect', { x: x + 3.296, y: 7.925, w: 3.729, h: 1.494, fill: { color: ORANGE } });
  text(s, 'Donate', x + 3.846, 8.385, 2.38, 0.572, { ...LABELB, color: WHITE, align: 'center' });
  s.addShape('rect', { x, y: 9.417, w: 7.025, h: 2.01, fill: { color: WHITE }, shadow: softShadow(28, 0.26) });
  text(s, [{ text: 'Contribute For The Abled ' }, { text: 'Child Cancer', options: { color: ORANGE } }],
    x + 0.801, 9.912, 5.399, 1.043, { ...LABEL, color: INK });
  s.addShape('rect', { x, y: 11.388, w: 7.025, h: 2.01, fill: { color: GREEN } });
  text(s, 'Achieved', x + 0.801, 11.732, 2.088, 0.572, { ...LABELB, color: WHITE });
  text(s, '0%', x + 4.665, 11.665, 0.997, 0.639, { ...STAT32, color: WHITE, align: 'right' });
  text(s, 'Target', x + 0.801, 12.372, 2.088, 0.572, { ...LABELB, color: WHITE });
  text(s, '$50.000', x + 3.574, 12.305, 2.088, 0.639, { ...STAT32, color: WHITE, align: 'right' });
}

/* ------------------------------------------------------ slide 5: cancer cases */
function slide05(pres) {
  const s = pres.addSlide();
  // Empty photo frames: 11.60 x 7.24 at (1.52, 4.26) and 11.55 x 7.24 at (13.58, 6.20).
  text(s, '- Hold hands & Give New Life -', 9.55, 1.534, 7.733, 0.572, { ...LABEL, color: GREEN, align: 'center' });
  text(s, [{ text: 'Our Cancer ' }, { text: 'Cases', options: { color: GREEN } }],
    8.046, 2.054, 10.741, 1.313, { ...TITLE, color: INK, bold: true, align: 'center' });

  wideCaseCard(s, 13.548, 4.231);
  wideCaseCard(s, 1.53, 9.532);

  rings(s, 7.089, 1.203, 7.555, 1.415, 0.824, 8.5, GREEN, ORANGE);
  rings(s, 21.836, 13.059, 22.302, 13.271, 0.824, 8.5, GREEN, ORANGE);
  corner(s, 25.677, 0, 0.989, 1);
  corner(s, 0, 14.315, 0.677, 0.685);
}

/* ------------------------------------------------------- slide 6: we need funds */

// Donut gauge: light disc, orange partial ring, shadowed inner disc, arrow + amount.
function donutStat(s, x, y, amount, arrowDown) {
  s.addShape('ellipse', { x, y, w: 2.909, h: 2.909, fill: { color: W95 } });
  s.addShape('pie', { x: x + 0.167, y: y + 0.167, w: 2.576, h: 2.576, fill: { color: ORANGE } });
  s.addShape('ellipse', { x: x + 0.446, y: y + 0.446, w: 2.018, h: 2.018,
    fill: { color: W95 }, shadow: softShadow(23, 0.28) });
  s.addShape('downArrow', Object.assign(
    { x: x + 1.241, y: y + 0.85, w: 0.426, h: 0.465, fill: { color: GREEN } },
    arrowDown ? {} : { rotate: 180 }));
  text(s, amount, x + 0.446, y + 1.439, 2.018, 0.572, { ...LABEL, color: INK, align: 'center' });
}

function slide06(pres) {
  const s = pres.addSlide();
  // Empty photo frame: 12.80 x 15.00 at (0, 0).
  poly(s, 0, 0, 25.381, 15,
    [[0, 0], [12.495, 0], [25.381, 15], [11.913, 15], [0, 1.247]],
    { fill: { color: DEEP, transparency: 96 } });

  text(s, '- Fund Progress -', 14.123, 2.287, 7.733, 0.572, { ...LABEL, color: GREEN });
  text(s, [{ text: 'We Need Funds', options: { breakLine: true } },
           { text: 'To ' }, { text: 'Help !', options: { color: ORANGE } }],
    14.123, 2.967, 9.48, 2.524, { ...TITLE, color: INK });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident,',
    14.123, 5.649, 9.942, 3.669, { ...PARA, color: BODY });

  donutStat(s, 14.123, 9.976, '$2,500', false);
  donutStat(s, 17.367, 9.976, '$6,500', true);
  donutStat(s, 20.626, 9.81, '$8,000', true);

  rings(s, 9.347, 2.848, 9.813, 3.06, 0.824, 8.5, GREEN, ORANGE);
  rings(s, 23.238, 12.719, 23.704, 12.931, 0.824, 8.5, GREEN, ORANGE);
  corner(s, 25.677, 0, 0.989, 1);
  corner(s, 12.806, 14.283, 0.71, 0.717);
}

/* ------------------------------------------------------ slide 7: three services */
function slide07(pres) {
  const s = pres.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: { color: GREEN, transparency: 96 } });

  text(s, '- Services -', 2.768, 1.746, 7.733, 0.572, { ...LABEL, color: GREEN });
  text(s, [{ text: 'The Purpose Behind Our ' }, { text: 'Passion', options: { color: GREEN } }],
    2.768, 2.426, 10.458, 2.524, { ...TITLE, color: INK });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ',
    14.346, 2.171, 9.942, 2.457, { ...PARA, color: BODY });

  // Three cards; the third is a solid green "highlight" card.
  const CARDS = [
    { x: 2.026, textX: 2.7679, title: 'Donate Houses',  titleY: 9.107, bodyY: 9.861, onGreen: false, iconY: 7.096, iconD: 1.747 },
    { x: 9.759, textX: 10.501, title: 'Donate Food',    titleY: 9.153, bodyY: 9.909, onGreen: false, iconY: 7.125, iconD: 1.688 },
    { x: 17.53, textX: 18.188, title: 'Donate Medical', titleY: 9.105, bodyY: 9.861, onGreen: true,  iconY: 7.137, iconD: 1.665 },
  ];
  CARDS.forEach((c) => {
    s.addShape('rect', Object.assign({ x: c.x, y: 5.814, w: 7.733, h: 7.057 },
      c.onGreen ? { fill: { color: GREEN } } : { fill: { color: WHITE }, shadow: softShadow(20, 0.34) }));
    // Placeholder in place of the source deck's flat pictogram art.
    s.addShape('roundRect', { x: c.textX, y: c.iconY, w: c.iconD, h: c.iconD,
      fill: { color: c.onGreen ? W95 : W85 }, rectRadius: 0.15 });
    text(s, '[image]', c.textX, c.iconY, c.iconD, c.iconD,
      { fontSize: 11, fontFace: MD, color: BODY, align: 'center', valign: 'middle' });
    text(s, c.title, c.textX, c.titleY, 4.199, 0.651, { ...HEAD32, color: c.onGreen ? WHITE : INK });
    text(s, 'Lorem ipsum dolor sit a met, connecter adipescent elite, sed do eiusmod temper incident Ut labore et dolore,',
      c.textX, c.bodyY, 6.532, 1.851, { ...PARA, color: c.onGreen ? W85 : BODY });
  });

  rings(s, 7.072, 12.498, 7.514, 12.396, 0.902, 6, GREEN, ORANGE);
  rings(s, 22.944, 5.303, 23.387, 5.202, 0.902, 6, GREEN, ORANGE);
  corner(s, 0, 0, 0.989, 1);
  corner(s, 26.064, 14.391, 0.602, 0.609);
}

/* -------------------------------------------- slide 8: full-bleed photo (empty) */
function slide08(pres) {
  pres.addSlide(); // one full-slide picture placeholder, no artwork in the source
}

/* ------------------------------------------------------------ slide 9: our team */
function slide09(pres) {
  const s = pres.addSlide();
  [2.151, 17.026].forEach((x) => s.addShape('rect',
    { x, y: 0, w: 7.494, h: 15, fill: { color: GREEN, transparency: 96 } }));
  // Empty portrait frames: 6.30 x 7.11 at x = 2.744 / 10.163 / 17.622, y = 4.507.

  text(s, 'Who Help To Charity', 9.467, 1.591, 7.733, 0.572, { ...LABEL, color: GREEN, align: 'center' });
  text(s, "Meet Our Valente's", 7.963, 2.231, 11.601, 1.447, { ...TITLE80, color: INK, align: 'center' });

  [['Scott William', 3.199], ['Tania Vandy', 10.509], ['Liam Irvines', 18.073]].forEach(([name, x]) => {
    text(s, name, x, 12.034, 5.399, 0.774, { ...HEAD40, color: INK, align: 'center' });
    text(s, 'Valentes', x, 12.836, 5.399, 0.572, { ...LABEL, color: GREEN, align: 'center' });
  });

  rings(s, 3.946, 3.868, 4.556, 3.934, 0.902, 6, GREEN, ORANGE);
  rings(s, 24.367, 12.358, 23.916, 12.812, 0.902, 6, GREEN, ORANGE);
  corner(s, 25.677, 0, 0.989, 1);
  corner(s, 0, 14.315, 0.677, 0.685);
}

/* ------------------------------------------------- slide 10: six process steps */

// Card with one semicircular end; `flipH` mirrors it for the left-hand column.
function stepPill(s, x, y, w, color, flipH) {
  const h = 3.531;
  const cap = 0.1424 * w, c1 = 0.0686 * w, c2 = 0.008 * w, tip = 0.0007 * w;
  s.addShape('custGeom', {
    x, y, w, h, fill: { color }, flipH: !!flipH,
    points: [
      { x: cap, y: 0, moveTo: true }, { x: w, y: 0 }, { x: w, y: h }, { x: cap, y: h },
      { x: tip, y: 0.551 * h, curve: { type: 'cubic', x1: c1, y1: h, x2: c2, y2: 0.803 * h } },
      { x: 0, y: 0.5 * h },
      { x: tip, y: 0.449 * h },
      { x: cap, y: 0, curve: { type: 'cubic', x1: c2, y1: 0.197 * h, x2: c1, y2: 0 } },
      { close: true },
    ],
  });
}

function slide10(pres) {
  const s = pres.addSlide();
  const LEAD = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore,';

  // Left column: pill mirrored so its round end faces inwards, disc + number
  // sit on the right of the bar and the copy runs down the left margin.
  const LEFT = [
    { n: '01', title: 'Give Clean Water', y: 2.953,  color: GREEN,  discY: 3.391,  numX: 10.507, numY: 4.196,  numW: 1.247, titleY: 3.836,  bodyY: 4.486 },
    { n: '02', title: 'Give Medical',     y: 6.818,  color: ORANGE, discY: 7.229,  numX: 10.384, numY: 7.978,  numW: 1.493, titleY: 7.697,  bodyY: 8.347 },
    { n: '03', title: 'Give Clothes',     y: 10.682, color: GREEN,  discY: 11.094, numX: 10.384, numY: 11.863, numW: 1.493, titleY: 11.558, bodyY: 12.208 },
  ];
  LEFT.forEach((c) => {
    stepPill(s, 0, c.y, 13, c.color, true);
    s.addShape('ellipse', { x: 9.755, y: c.discY, w: 2.75, h: 2.75, fill: { color: WHITE } });
    text(s, c.n, c.numX, c.numY, c.numW, 1.212, { ...TITLE66, color: c.color });
    text(s, c.title, 1.104, c.titleY, 4.621, 0.639, { ...HEAD32, color: WHITE });
    text(s, LEAD, 1.104, c.bodyY, 7.887, 1.245, { ...PARA, color: W95 });
  });

  const RIGHT = [
    { n: '04', title: 'Charity Donate', y: 0.938, color: ORANGE, discY: 1.385, numX: 15.02,  numY: 2.196, numW: 1.406, titleY: 1.786, bodyY: 2.436 },
    { n: '05', title: 'Donations',      y: 4.802, color: GREEN,  discY: 5.224, numX: 15.042, numY: 6.04,  numW: 1.493, titleY: 5.648, bodyY: 6.298 },
    { n: '06', title: 'Give Houses',    y: 8.667, color: ORANGE, discY: 9.088, numX: 15.042, numY: 9.925, numW: 1.493, titleY: 9.509, bodyY: 10.159 },
  ];
  RIGHT.forEach((c) => {
    stepPill(s, 13.906, c.y, 12.76, c.color, false);
    s.addShape('ellipse', { x: 14.406, y: c.discY, w: 2.75, h: 2.75, fill: { color: WHITE } });
    text(s, c.n, c.numX, c.numY, c.numW, 1.212, { ...TITLE66, color: c.color });
    text(s, c.title, 17.675, c.titleY, 4.621, 0.639, { ...HEAD32, color: WHITE });
    text(s, LEAD, 17.675, c.bodyY, 7.887, 1.245, { ...PARA, color: W95 });
  });
}

/* --------------------------------------------------- slide 11: donation form */
function slide11(pres) {
  const s = pres.addSlide();
  // Empty photo frame: 10.54 x 15.00 at (16.13, 0).
  s.addShape('rect', { x: 4.772, y: 0, w: 11.389, h: 15, fill: { color: GREEN, transparency: 96 } });

  text(s, '- Fund Progress -', 2.347, 3.322, 4.228, 0.572, { ...LABEL, color: GREEN });
  text(s, [{ text: 'Bringing Clean Water To Every Corner ' },
           { text: 'Of The World', options: { color: GREEN } }],
    2.18, 3.851, 9.149, 4.543, { ...TITLE66, color: INK, bold: true });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore,',
    2.18, 8.395, 8.766, 1.245, { ...PARA, color: BODY });

  // Row of overlapping supporter avatars (empty photo circles in the source).
  [2.204, 3.136, 4.121, 5.053].forEach((x) => s.addShape('ellipse',
    { x, y: 10.28, w: 1.29, h: 1.29, fill: { color: GRAY }, line: { color: WHITE, width: 1 } }));
  text(s, '2M +', 6.605, 10.608, 1.539, 0.707, { fontSize: 36, fontFace: SB, color: GREEN });

  // Floating donation card.
  s.addShape('rect', { x: 11.675, y: 2.132, w: 9.016, h: 10.735,
    fill: { color: WHITE }, shadow: softShadow(43, 0.21) });
  text(s, 'Hold The Hand', 14.489, 3.322, 3.389, 0.572, { ...LABEL, color: INK5, align: 'center' });
  text(s, 'Need Help', 13.64, 3.934, 5.085, 0.909, { ...HEAD48, color: DEEP, align: 'center' });
  text(s, 'Lorem ipsum dolor sit a met consenter.', 12.483, 4.915, 7.399, 0.639,
    { ...PARA, color: MUTED, align: 'center' });
  // [outlined field y, label, label y, label width]
  [[6.332, 'Name:', 6.581, 1.601], [7.5, 'Email:', 7.781, 1.601],
   [9.512, '$50', 9.765, 1.118]].forEach(([fieldY, label, labelY, labelW]) => {
    s.addShape('rect', { x: 12.933, y: fieldY, w: 6.5, h: 0.909, line: { color: W85, width: 1.75 } });
    text(s, label, 13.37, labelY, labelW, 0.438, { ...TINY, color: BODY });
  });
  text(s, 'Select The Donation Amount', 13.37, 8.811, 5.017, 0.438, { ...TINY, color: BODY });
  s.addShape('rect', { x: 11.675, y: 10.832, w: 9.016, h: 0.988, fill: { color: GREEN } });
  text(s, 'Donate Now', 14.753, 11.041, 2.861, 0.572, { ...LABEL, color: WHITE, align: 'center' });

  rings(s, 12.491, 1.692, 12.933, 1.591, 0.902, 6, GREEN, ORANGE);
  rings(s, 2.202, 12.545, 2.644, 12.444, 0.902, 6, GREEN, ORANGE);
  corner(s, 0, 14, 0.989, 1);
}

/* ---------------------------------------------------- slide 12: houseless cases */
function slide12(pres) {
  const s = pres.addSlide();
  // Empty photo frames: 7.02 x 5.28 at x = 2.384 / 9.799 / 17.179, y ≈ 4.13.
  text(s, '- Hold hands & Give New Life -', 9.215, 1.666, 7.733, 0.572, { ...LABEL, color: GREEN, align: 'center' });
  text(s, [{ text: 'Our Houseless ' }, { text: 'Cases', options: { color: GREEN } }],
    7.71, 2.186, 11.205, 1.313, { ...TITLE, color: INK, bold: true, align: 'center' });

  [2.384, 9.821, 17.174].forEach((x) => tallCaseCard(s, x));

  rings(s, 6.271, 1.571, 6.737, 1.783, 0.824, 8.5, GREEN, ORANGE);
  rings(s, 22.335, 12.92, 22.924, 12.982, 0.902, 6, GREEN, ORANGE);
  corner(s, 25.677, 0, 0.989, 1);
}

/* --------------------------------------------------- slide 13: different cases */
function slide13(pres) {
  const s = pres.addSlide();
  // Empty photo frames cluster on the right: 7.29 x 7.49 at (15.66, 1.06),
  // 8.09 x 8.63 at (18.57, 6.39) and 7.29 x 7.94 at (13.78, 4.67).
  s.addShape('rect', { x: 9.095, y: 0, w: 9.48, h: 15.003, fill: { color: GREEN, transparency: 91 } });

  text(s, '- Fund Progress -', 2.714, 2.478, 3.892, 0.572, { ...LABEL, color: GREEN });
  text(s, [{ text: 'We Have Different', options: { breakLine: true } },
           { text: 'Cases ' }, { text: 'To See!', options: { color: ORANGE } }],
    2.714, 3.173, 9.149, 2.322, { ...TITLE66, color: INK, bold: true });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore,',
    2.714, 5.585, 8.766, 1.245, { ...PARA, color: BODY });

  // Two identical "Charity Needs — 80%" progress rows, 1.597 in apart.
  [0, 1.597].forEach((dy) => {
    text(s, 'Charity Needs', 2.714, 7.823 + dy, 5.822, 0.572, { ...LABELB, color: INK });
    text(s, '80%', 11.108, 7.76 + dy, 1.182, 0.64, { ...HEAD32, color: INK });
    s.addShape('rect', { x: 2.714, y: 8.545 + dy, w: 7.45, h: 0.279, fill: { color: ORANGE } });
  });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliquant.',
    2.714, 11.116, 10.62, 1.245, { ...PARA, color: BODY });

  rings(s, 1.224, 1.193, 1.812, 1.256, 0.902, 6, GREEN, ORANGE);
  rings(s, 15.171, 12.088, 15.759, 12.15, 0.902, 6, GREEN, ORANGE);
  corner(s, 25.677, 0, 0.989, 1);
  corner(s, 0, 14.315, 0.677, 0.685);
}

/* --------------------------------------------------- slide 14: charity gallery */
function slide14(pres) {
  const s = pres.addSlide();
  s.addShape('rect', { x: 0, y: 10.399, w: SLIDE_W, h: 4.604, fill: { color: GREEN, transparency: 92 } });
  // Eight empty 4.60 in photo tiles in a 4 x 2 grid starting at (4.125, 4.125).

  text(s, '- Some Special Moments -', 10.466, 1.514, 5.736, 0.569, { ...LABELB, color: GREEN, align: 'center' });
  text(s, 'Charity Gallery', 9.431, 2.243, 7.805, 1.212, { ...TITLE66, color: INK, align: 'center' });

  rings(s, 8.14, 1.207, 8.606, 1.419, 0.824, 8.5, GREEN, ORANGE);
  rings(s, 19.128, 3.501, 19.594, 3.713, 0.824, 8.5, GREEN, ORANGE);
  corner(s, 25.677, 0, 0.989, 1);
  corner(s, 0, 14.315, 0.677, 0.685);
}

/* ------------------------------------------------------ slide 15: hunger stats */
function slide15(pres) {
  const s = pres.addSlide();
  // Empty photo frames: 6.77 x 15.00 at x = 0 and x = 6.787.
  s.addShape('rect', { x: 9.22, y: 0, w: 9.48, h: 15.003, fill: { color: GREEN, transparency: 92 } });

  text(s, '- We Need More -', 14.545, 2.968, 7.401, 0.572, { ...LABELB, color: GREEN });
  text(s, 'When Hope Is Hungry, Everything Feeds It.', 14.545, 3.557, 9.149, 3.433,
    { ...TITLE66, color: INK, bold: true });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore,',
    14.545, 7.215, 8.766, 1.245, { ...PARA, color: BODY });

  text(s, 'Special Kids And Cases OF Poor', 14.545, 10.086, 4.002, 1.043, { ...LABELB, color: INK });
  text(s, 'Read More', 14.545, 11.347, 2.38, 0.505, { ...SMALLB, color: GREEN });
  s.addShape('triangle', { x: 17.291, y: 11.438, w: 0.376, h: 0.324, fill: { color: GREEN }, rotate: 90 });

  // Two solid stat blocks.
  [{ x: 18.687, w: 2.824, fill: GREEN,  value: '45k', valueX: 19.263, label: 'Poor Kids', labelX: 18.909, labelW: 2.38 },
   { x: 21.43,  w: 2.722, fill: ORANGE, value: '85k', valueX: 21.954, label: 'Homeless',  labelX: 21.78,  labelW: 2.022 },
  ].forEach((b) => {
    s.addShape('rect', { x: b.x, y: 9.231, w: b.w, h: 3.239, fill: { color: b.fill } });
    text(s, b.value, b.valueX, 10.221, 1.672, 0.909, { ...HEAD48, color: WHITE, align: 'center' });
    text(s, b.label, b.labelX, 11.186, b.labelW, 0.505, { ...SMALLB, color: WHITE, align: 'center' });
  });

  rings(s, 22.748, 2.673, 23.214, 2.885, 0.824, 8.5, GREEN, ORANGE);
  rings(s, 16.229, 13.557, 16.695, 13.769, 0.824, 8.5, GREEN, ORANGE);
  corner(s, 25.677, 0, 0.989, 1);
}

/* ---------------------------------------------------- slide 16: charity analysis */
function slide16(pres) {
  const s = pres.addSlide();

  // Two mirrored pyramids of stadium bars; widths shrink towards the top and
  // the colours alternate, so the two columns are each other's negative.
  const ROWS = [
    { y: 4.5,    titleY: 5.37,    bodyY: 6.01,   leftX: 3.996, leftW: 7.061,  leftFill: ORANGE, rightX: 12.857, rightW: 11.983, rightFill: GREEN },
    { y: 7.438,  titleY: 8.2667,  bodyY: 8.906,  leftX: 2.906, leftW: 9.24,   leftFill: GREEN,  rightX: 14.229, rightW: 9.24,   rightFill: ORANGE },
    { y: 10.375, titleY: 11.1635, bodyY: 11.803, leftX: 1.534, leftW: 11.983, leftFill: ORANGE, rightX: 15.319, rightW: 7.061,  rightFill: GREEN },
  ];
  ROWS.forEach((r) => {
    stadium(s, r.leftX, r.y, r.leftW, 2.938, r.leftFill);
    stadium(s, r.rightX, r.y, r.rightW, 2.938, r.rightFill);
    [[5.216, 4.191], [16.542, 15.518]].forEach(([titleX, bodyX]) => {
      text(s, 'Charity Donate', titleX, r.titleY, 4.621, 0.639, { ...HEAD32, color: WHITE, align: 'center' });
      text(s, 'Lorem ipsum dolor sit met connecter.', bodyX, r.bodyY, 6.669, 0.639,
        { ...PARA, color: W95, align: 'center' });
    });
  });

  text(s, '- Charity Give Life -', 11.109, 1.714, 4.449, 0.572, { ...LABEL, color: GREEN, align: 'center' });
  text(s, 'Charity Analysis', 8.207, 2.343, 10.254, 1.447, { ...TITLE80, color: INK, bold: true, align: 'center' });

  rings(s, 8.207, 1.178, 8.672, 1.39, 0.824, 8.5, GREEN, ORANGE);
  rings(s, 20.32, 3.873, 20.786, 4.085, 0.824, 8.5, GREEN, ORANGE);
  corner(s, 25.677, 0, 0.989, 1);
  corner(s, 0, 14.315, 0.677, 0.685);
}

/* ----------------------------------------------------- slide 17: contribute copy */
function slide17(pres) {
  const s = pres.addSlide();
  // Three empty full-height photo columns: 5.84 x 15.00 at x = 9.140 / 14.974 / 20.825.
  text(s, '- Donate Now For Kids Future -', 1.446, 3.092, 6.46, 0.572, { ...LABELB, color: GREEN });
  text(s, [{ text: 'Contribute Make For The Somalians ' },
           { text: 'Happy And Foods', options: { color: GREEN } }],
    1.446, 3.821, 6.751, 5.655, { ...TITLE66, color: INK });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim venom,',
    1.446, 9.681, 6.751, 2.457, { ...PARA, color: BODY });

  rings(s, 1.748, 1.486, 2.214, 1.698, 0.824, 8.5, GREEN, ORANGE);
  rings(s, 6.303, 12.464, 6.769, 12.676, 0.824, 8.5, GREEN, ORANGE);
  corner(s, 8.174, 0, 0.989, 1);
  corner(s, 0, 14.315, 0.677, 0.685);
}

/* -------------------------------------------------------- slide 18: testimonial */
function slide18(pres) {
  const s = pres.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: SLIDE_W, h: 7.035, fill: { color: GREEN, transparency: 96 } });
  s.addShape('rect', { x: 4.01, y: 7.035, w: 18.688, h: 5.834, fill: { color: GREEN } });
  // Three empty circular portrait frames straddle the band around (9.1-15.4, 5.5).

  text(s, '- Hold hands & Give New Life -', 9.215, 1.947, 7.733, 0.572, { ...LABEL, color: GREEN, align: 'center' });
  text(s, [{ text: 'Testimonials From Our ' }, { text: 'Community', options: { color: ORANGE } }],
    7.71, 2.467, 11.205, 2.524, { ...TITLE, color: INK, bold: true, align: 'center' });
  text(s, 'Jhene Sons', 11.367, 9.245, 3.892, 0.639, { ...HEAD32, color: WHITE, align: 'center' });
  text(s, 'Donner', 11.387, 9.825, 3.892, 0.505, { ...SMALL, color: WHITE, align: 'center' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore,',
    6.967, 10.404, 12.693, 1.245, { ...PARA, color: W85, align: 'center' });

  ring(s, 8.243, 0.985, 0.916, GREEN, 8.5);
  ring(s, 18.654, 4.852, 0.916, ORANGE, 8.5);
  rings(s, 5.042, 6.463, 5.508, 6.676, 0.824, 8.5, GREEN, ORANGE);
  rings(s, 20.052, 12.469, 20.518, 12.681, 0.824, 8.5, GREEN, ORANGE);
  corner(s, 25.677, 0, 0.989, 1);
  corner(s, 0, 14.315, 0.677, 0.685);
}

/* ------------------------------------------- slide 19: full-bleed photo (empty) */
function slide19(pres) {
  pres.addSlide(); // one full-slide picture placeholder, no artwork in the source
}

/* --------------------------------------------------------- slide 20: price table */
function slide20(pres) {
  const s = pres.addSlide();

  // The middle "STANDARD" column is taller, green, and carries the Hot badge.
  const COLUMNS = [
    {
      x: 4.0866, y: 5.3809, h: 7.303, cardFill: WHITE, footerY: 10.8702, footerH: 1.821,
      title: 'BASIC', titleX: 5.1511, titleY: 6.3865, titleColor: INK,
      itemColor: INK25, dotColor: GREEN, dotX: 5.3218, priceX: 5.1511,
      items: [
        { t: '50% Discount',         x: 5.768, y: 7.423, w: 3.141, h: 0.767, dotY: 7.677 },
        { t: 'Give Houses',          x: 5.821, y: 8.267, w: 3.088, h: 0.741, dotY: 8.584 },
        { t: 'Give Water & Clothes', x: 5.768, y: 9.084, w: 3.088, h: 1.448, dotY: 9.344 },
      ],
    },
    {
      x: 10.4192, y: 4.6526, h: 8.016, cardFill: GREEN, footerY: 10.8803, footerH: 1.801,
      title: 'STANDARD', titleX: 11.3669, titleY: 5.5594, titleColor: WHITE,
      itemColor: WHITE, dotColor: ORANGE, dotX: 11.5197, priceX: 11.3669, hot: true,
      items: [
        { t: '50% Discount',         x: 11.999, y: 6.5,   w: 3.712, h: 0.741, dotY: 6.8125 },
        { t: 'Give Houses',          x: 12.052, y: 7.343, w: 3.712, h: 0.741, dotY: 7.7192 },
        { t: 'Give Water & Clothes', x: 11.999, y: 8.16,  w: 3.712, h: 1.448, dotY: 8.4789 },
        { t: 'Give Medical',         x: 11.999, y: 9.684, w: 3.023, h: 0.741, dotY: 9.9859 },
      ],
    },
    {
      x: 16.7138, y: 4.6526, h: 8.016, cardFill: WHITE, footerY: 10.8803, footerH: 1.801,
      title: 'Premium', titleX: 17.6614, titleY: 5.5594, titleColor: INK,
      itemColor: INK25, dotColor: GREEN, dotX: 17.7489, priceX: 17.6614,
      items: [
        { t: '50% Discount',         x: 18.262, y: 6.5,   w: 3.712, h: 0.741, dotY: 6.8229 },
        { t: 'Give Houses',          x: 18.316, y: 7.343, w: 3.712, h: 0.741, dotY: 7.7296 },
        { t: 'Give Water & Clothes', x: 18.262, y: 8.16,  w: 3.712, h: 1.448, dotY: 8.4893 },
        { t: 'Give Medical',         x: 18.262, y: 9.684, w: 3.023, h: 0.741, dotY: 9.9963 },
      ],
    },
  ];

  COLUMNS.forEach((c) => {
    s.addShape('snip1Rect', { x: c.x, y: c.y, w: 5.816, h: c.h,
      fill: { color: c.cardFill }, shadow: softShadow(21, 0.32) });
    s.addShape('rect', { x: c.x, y: c.footerY, w: 5.816, h: c.footerH, fill: { color: ORANGE } });
    text(s, c.title, c.titleX, c.titleY, 3.272, 0.774, { ...HEAD40, color: c.titleColor });
    c.items.forEach((it) => {
      s.addShape('ellipse', { x: c.dotX, y: it.dotY, w: 0.258, h: 0.258, fill: { color: c.dotColor } });
      text(s, it.t, it.x, it.y, it.w, it.h, { ...LABEL15, color: c.itemColor });
    });
    text(s, [{ text: '9.55' }, { text: ' ', options: { fontSize: 36 } },
             { text: '/ Book', options: { fontSize: 28 } }],
      c.priceX, 11.3264, 4.186, 0.909, { fontSize: 48, fontFace: MD, color: WHITE });
    if (c.hot) {
      s.addShape('rect', { x: 15.0311, y: 9.8637, w: 1.2088, h: 1.0143, fill: { color: WHITE } });
      text(s, 'Hot', 15.0624, 10.094, 1.1496, 0.572, { ...LABELB, color: GREEN, align: 'center' });
    }
  });

  text(s, '- Charity Make Life -', 9.818, 1.749, 7.033, 0.572,
    { ...LABEL, color: GREEN, charSpacing: 3, align: 'center' });
  text(s, 'Our Price Table', 8.796, 2.332, 9.075, 1.447, { ...TITLE80, color: INK, align: 'center' });

  rings(s, 7.599, 1.31, 8.065, 1.523, 0.824, 8.5, GREEN, ORANGE);
  rings(s, 19.647, 4.066, 20.113, 4.278, 0.824, 8.5, GREEN, ORANGE);
  corner(s, 25.677, 0, 0.989, 1);
  corner(s, 0, 14.315, 0.677, 0.685);
}

/* --------------------------------------------------- slide 21: donate now spread */

// White side panel with a circular bite taken out of its inner edge, which
// frames the central photograph. `mirrored` flips the bite to the left edge.
function bittenPanel(s, x, mirrored) {
  const w = 10.859, h = 15;
  const pts = mirrored
    ? [[0, 0], [w, 0], [w, h], [0, h], [0.244, 14.857]]
    : [[0, 0], [w, 0], [10.616, 0.143]];
  s.addShape('custGeom', {
    x, y: 0, w, h, fill: { color: WHITE },
    points: pts.map(([px, py], i) => ({ x: px, y: py, moveTo: i === 0 })).concat(mirrored ? [
      { x: 3.794, y: 7.5, curve: { type: 'cubic', x1: 2.344, y1: 13.539, x2: 3.794, y2: 10.74 } },
      { x: 0.244, y: 0.143, curve: { type: 'cubic', x1: 3.794, y1: 4.26, x2: 2.344, y2: 1.461 } },
      { x: 0, y: 0 }, { close: true },
    ] : [
      { x: 7.065, y: 7.5, curve: { type: 'cubic', x1: 8.515, y1: 1.461, x2: 7.065, y2: 4.26 } },
      { x: 10.616, y: 14.857, curve: { type: 'cubic', x1: 7.065, y1: 10.74, x2: 8.515, y2: 13.539 } },
      { x: w, y: h }, { x: 0, y: h }, { x: 0, y: 0 }, { close: true },
    ]),
  });
}

function slide21(pres) {
  const s = pres.addSlide();
  // Empty photo frame fills the middle: 12.77 x 15.00 at (6.95, 0).
  bittenPanel(s, 0, false);
  bittenPanel(s, 15.807, true);

  text(s, '- Now For Kids Future -', 1.545, 2.857, 6.46, 0.572, { ...LABELB, color: GREEN });
  text(s, [{ text: 'Donate', options: { breakLine: true } },
           { text: 'Now For', options: { breakLine: true } },
           { text: 'Poor ' }, { text: 'Kids', options: { color: GREEN, breakLine: true } },
           { text: 'Give Them', options: { color: GREEN, breakLine: true } },
           { text: 'New Life', options: { color: GREEN } }],
    1.545, 3.586, 5.167, 5.655, { ...TITLE66, color: INK });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim venom,',
    1.545, 9.454, 5.477, 3.063, { ...PARA, color: BODY });

  // Two identical donation call-outs down the right edge: [amount y, label y, copy y].
  [[1.727, 2.857, 3.503], [8.065, 9.209, 9.854]].forEach(([amountY, labelY, copyY]) => {
    text(s, '25K', 20.417, amountY, 1.679, 0.909, { ...HEAD48, color: GREEN });
    text(s, 'Donation now', 20.417, labelY, 3.099, 0.572, { ...LABELB, color: GREEN });
    text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua,',
      20.417, copyY, 4.792, 3.063, { ...PARA, color: MUTED });
  });

  rings(s, 6.247, 1.196, 6.712, 1.408, 0.824, 8.5, GREEN, ORANGE);
  rings(s, 18.892, 12.705, 19.358, 12.917, 0.824, 8.5, GREEN, ORANGE);
  corner(s, 25.677, 0, 0.989, 1);
  corner(s, 0, 14.315, 0.677, 0.685);
}

/* ------------------------------------------------------------ slide 22: thank you */
function slide22(pres) {
  const s = pres.addSlide();
  s.addShape('rect', { x: 0.002, y: 0, w: 26.665, h: 14.999, fill: { color: GREEN } });

  text(s, 'LOGO', 1.967, 1.787, 1.366, 1.717, { fontSize: 48, fontFace: BD, color: WHITE, bold: true });
  text(s, 'Charity Presentation Template', 8.755, 4.677, 9.156, 0.641,
    { ...HEAD32, color: WHITE, charSpacing: 3, align: 'center' });
  // Oversized wordmark; the second line deliberately overlaps the copy below it.
  text(s, 'THANK YOU', 6.335, 5.318, 13.996, 2.894,
    { fontSize: 165.98, fontFace: BD, color: WHITE, bold: true, align: 'center' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim venom,',
    6.159, 8.083, 14.35, 1.245, { ...PARA, color: W95, align: 'center' });

  text(s, 'Contact Now For Oder', 2.452, 11.239, 9.156, 0.641, { ...HEAD32, color: WHITE, charSpacing: 3 });
  s.addShape('rect', { x: 2.531, y: 12.323, w: 21.605, h: 0.05, fill: { color: WHITE, transparency: 88 } });
  // [label, label x, label width, value, value x, value width]
  [['Phone:', 2.452, 2.175, '+963 254 5994', 4.627, 4.742],
   ['Email:', 9.231, 2.314, null, 11.405, 5.046],
   ['Website:', 17.724, 2.666, 'www.yoursite.com', 20.39, 5.157],
  ].forEach(([label, lx, lw, value, vx, vw]) => {
    text(s, label, lx, 13.161, lw, 0.639, { ...HEAD32, color: WHITE, charSpacing: 3 });
    text(s, value || [{ text: 'Yourmail' }, { text: '@gmail.com', options: { charSpacing: 3 } }],
      vx, 13.193, vw, 0.572, { ...LABEL, color: WHITE });
  });

  s.addShape('ellipse', { x: 5.485, y: 9.018, w: 0.37, h: 0.37, fill: { color: WHITE } });
  s.addShape('ellipse', { x: 19.139, y: 3.506, w: 0.37, h: 0.37, fill: { color: WHITE } });
  ring(s, 6.664, 3.504, 0.844, WHITE, 8.5);
  ring(s, 20.198, 10.246, 0.844, WHITE, 8.5);
}

/* -------------------------------------------------------------------- build */
const pres = new pptxgen();
pres.defineLayout({ name: 'WIDE', width: SLIDE_W, height: SLIDE_H });
pres.layout = 'WIDE';
pres.title = 'Social Activity & Charity Donations';

[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
 slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
 slide17, slide18, slide19, slide20, slide21, slide22].forEach((build) => build(pres));

pres.writeFile({ fileName: path.join(__dirname, '02c3bc78-af50-4a0d-991c-4e3d06aba270_grok_final.pptx') })
  .then((f) => console.log('wrote', f));
