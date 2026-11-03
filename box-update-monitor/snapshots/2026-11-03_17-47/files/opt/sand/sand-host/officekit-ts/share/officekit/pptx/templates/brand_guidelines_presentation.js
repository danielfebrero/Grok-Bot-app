/**
 * "Brand Guideline" presentation template — rebuilt with pptxgenjs.
 *
 * 20 slides, 13.333in x 7.5in.
 * The deck's only embedded raster — a tablet mock-up on slide 13 — is redrawn
 * here from native shapes. Every other picture region is an *empty* picture
 * placeholder in the source (no artwork), so each is reproduced by
 * `photoFrame()` as a flat rectangle in the page tone: the reserved area stays
 * visible in the code and in the file without inventing imagery.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Theme
 * ------------------------------------------------------------------ */

const C = {
  ink: '000000',
  white: 'FFFFFF',
  cream: 'F6F2ED',   // accent2 @ 80% luminance — the warm page background
  paper: 'F6F6F6',   // accent1 — the cool page background
  sand: 'D4C0A7',    // accent2
  mist: 'CCCCC9',    // accent3
  clay: '959083',    // accent4
  moss: '60655B',    // accent5
  coal: '3E3E3E',    // accent6
  charcoal: '404040',
  linen: 'EEE6DC',
  silver: 'E7E6E6',
  gridline: 'D9D9D9',
  axisText: '595959',
};

const SERIF = 'Playfair Display';
const SANS = 'Open Sans';

// PowerPoint's default text-box insets, in points (pptxgenjs margin unit).
const INSET = [7.2, 7.2, 3.6, 3.6]; // [left, right, bottom, top]

const LOREM =
  'PLACEHOLDER' +
  'PLACEHOLDER' +
  'excepturi sint occaecati cupiditate non provident, similique sunt in culpa.';

/** First `words` words of LOREM, closed with a period. */
function lorem(words) {
  return LOREM.split(' ').slice(0, words).join(' ').replace(/[,.]$/, '') + '.';
}

/* ------------------------------------------------------------------ *
 * Element helpers
 * ------------------------------------------------------------------ */

/** Playfair display headline — 96pt at 80% leading unless overridden. */
function headline(slide, x, y, w, h, text, opts = {}) {
  slide.addText(text, {
    x, y, w, h,
    fontFace: SERIF, fontSize: opts.fontSize || 96, color: opts.color || C.ink,
    align: opts.align || 'left', valign: 'top', margin: INSET,
    lineSpacingMultiple: opts.lineSpacingMultiple === undefined ? 0.8 : opts.lineSpacingMultiple,
  });
}

/** 12pt Open Sans body copy at 150% leading. */
function body(slide, x, y, w, h, text, opts = {}) {
  slide.addText(text, {
    x, y, w, h,
    fontFace: SANS, fontSize: 12, color: C.ink,
    align: opts.align || 'left', valign: 'top', margin: INSET,
    lineSpacingMultiple: 1.5,
    bullet: opts.bullet ? { characterCode: '2022', indent: 13.5 } : false,
  });
}

/** 12pt Playfair running-head / footer label. */
function microLabel(slide, x, y, text, align = 'left') {
  slide.addText(text, {
    x, y, w: 2.121, h: 0.303,
    fontFace: SERIF, fontSize: 12, color: C.ink,
    align, valign: 'top', margin: INSET,
  });
}

/** Hairline horizontal rule. */
function rule(slide, x, y, w, width = 0.75) {
  slide.addShape('line', { x, y, w, h: 0, line: { color: C.ink, width } });
}

function rect(slide, x, y, w, h, fill) {
  slide.addShape('rect', { x, y, w, h, fill: { color: fill }, line: { type: 'none' } });
}

/**
 * Reserved picture area. The source placeholders are empty, so the frame is
 * filled with the page tone it sits on and left unlabelled.
 */
function photoFrame(slide, x, y, w, h, pageTone) {
  rect(slide, x, y, w, h, pageTone);
}

/** "Brand Guideline ———— Presentation Template" masthead. */
function masthead(slide, left, right) {
  rule(slide, 2.29, 0.522, 8.159);
  microLabel(slide, 0.726, 0.332, left, 'left');
  microLabel(slide, 10.45, 0.353, right, 'right');
}

/** Contents row: serif label left, counter right, optional rule beneath. */
function indexRow(slide, x, y, label, counter, ruleY) {
  const common = { fontFace: SERIF, fontSize: 14, color: C.ink, h: 0.337, valign: 'top', margin: INSET };
  slide.addText(label, { ...common, x, y, w: 2.398, align: 'left' });
  slide.addText(counter, { ...common, x: x + 2.991, y, w: 1.413, align: 'right' });
  if (ruleY !== undefined) rule(slide, x + 0.037, ruleY, 4.33);
}

/** "Logo." wordmark, optionally on a colour swatch. */
function logoSwatch(slide, x, y, w, h, fill, textColor) {
  if (fill) rect(slide, x, y, w, h, fill);
  headline(slide, x + 0.895, y + 0.778, 3.954, 1.393, 'Logo.', { align: 'center', color: textColor });
}

/** One 72pt glyph specimen cell. */
function glyph(slide, x, y, w, h, text, fontFace) {
  slide.addText(text, {
    x, y, w, h, fontFace, fontSize: 72, color: C.ink,
    valign: 'top', margin: INSET, lineSpacingMultiple: 1.2,
  });
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

// 1 — Cover
function slide01(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.cream };
  masthead(s, 'Brand Guideline', 'Presentation Template');
  headline(s, 6.341, 1.365, 6.57, 3.332, 'Brand Guideline.', { lineSpacingMultiple: null });
  const agenda = [
    { x: 6.494, y: 5.052, w: 2.629, text: 'Brand Overview' },
    { x: 8.477, y: 5.062, w: 1.263, text: 'Logo' },
    { x: 9.341, y: 5.062, w: 2.125, text: 'Typography' },
    { x: 6.494, y: 5.543, w: 2.034, text: 'Aplication' },
    { x: 7.847, y: 5.553, w: 1.747, text: 'Imagery' },
    { x: 9.109, y: 5.553, w: 2.125, text: 'Contact' },
  ];
  agenda.forEach(a => s.addText(a.text, {
    x: a.x, y: a.y, w: a.w, h: 0.462,
    fontFace: SANS, fontSize: 16, color: C.ink, valign: 'top', margin: INSET,
    lineSpacingMultiple: 1.5, bullet: { characterCode: '2022', indent: 13.5 },
  }));
  rule(s, 0.002, 6.971, 13.329);
  photoFrame(s, 0.839, 1.674, 3.331, 3.331, C.cream);
}

// 2 — Table of content
function slide02(pptx) {
  const s = pptx.addSlide();
  rect(s, 6.667, -0.007, 6.667, 7.507, C.cream);
  headline(s, 0.782, 0.649, 3.047, 1.447, 'Table Of Content', { fontSize: 40, lineSpacingMultiple: null });
  body(s, 7.435, 1.018, 4.133, 0.978, lorem(19));
  [
    { y: 3.350, ruleY: 4.023 },
    { y: 4.296, ruleY: 4.906 },
    { y: 5.179, ruleY: 5.789 },
    { y: 6.062 },
  ].forEach(r => indexRow(s, 0.839, r.y, 'Your Text Here', '001/004', r.ruleY));
  photoFrame(s, 7.493, 2.746, 5.002, 4.754, C.cream);
}

// 3 — Visual Identity divider
function slide03(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.paper };
  headline(s, 0.768, 1.799, 5.963, 2.686, 'Visual Identity');
  body(s, 6.73, 1.693, 5.0, 1.28, LOREM);
  body(s, 6.73, 3.845, 5.0, 1.28, LOREM);
  rule(s, 0.002, 6.796, 13.329);
  microLabel(s, 0.726, 6.979, 'Visual Identity System', 'left');
  microLabel(s, 10.486, 6.979, 'Brand Guideline', 'right');
}

// 4 — Brand Overview divider
function slide04(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.cream };
  headline(s, 6.532, 1.799, 5.963, 2.686, 'Brand Overview');
  body(s, 6.599, 4.691, 5.31, 0.978, lorem(25));
  microLabel(s, 10.45, 0.353, 'Presentation Template', 'right');
  rule(s, 6.667, 6.032, 1.181);
  s.addShape('ellipse', { x: 7.889, y: 5.994, w: 0.076, h: 0.076, fill: { color: C.ink }, line: { type: 'none' } });
  photoFrame(s, 0.839, 0.521, 3.663, 5.147, C.cream);
  photoFrame(s, 3.175, 3.083, 2.666, 3.896, C.cream);
}

// 5 — Value, split warm/cool bands
function slide05(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.paper };
  rect(s, 0, 0, 13.333, 3.284, C.cream);
  headline(s, 0.868, 1.064, 4.311, 1.393, 'Value');
  [
    { titleY: 1.064, bodyY: 1.594, h: 1.28, words: 25 },
    { titleY: 4.055, bodyY: 4.540, h: 0.978, words: 21 },
  ].forEach(b => {
    headline(s, 7.402, b.titleY, 2.414, 0.37, 'The Tittle Here', { fontSize: 20 });
    body(s, 7.402, b.bodyY, 4.311, b.h, lorem(b.words));
  });
  body(s, 7.402, 5.776, 4.109, 0.978, lorem(19));
  photoFrame(s, 0.839, 4.055, 5.002, 2.924, C.paper);
}

// 6 — Vision & Mission
function slide06(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.white };
  rect(s, 7.493, -0.007, 5.833, 3.757, C.cream);
  rect(s, 7.493, 3.743, 5.833, 3.757, C.paper);
  headline(s, 0.763, 1.064, 5.963, 2.686, 'Vision &\nMission');
  [{ y: 0.794, title: 'Vision.' }, { y: 4.544, title: 'Mission.' }].forEach(g => {
    headline(s, 8.257, g.y, 1.475, 0.424, g.title, { fontSize: 24 });
    body(s, 8.257, g.y + 0.874, 4.085, 1.28, lorem(25));
  });
  photoFrame(s, 0.839, 4.544, 5.0, 2.956, C.white);
}

// 7 — Value, 100% stacked bar chart
function slide07(pptx) {
  const s = pptx.addSlide();
  headline(s, 0.868, 0.81, 4.311, 1.393, 'Value');
  body(s, 5.751, 0.904, 3.424, 0.978, lorem(16));
  const labels = ['Category 1', 'Category 2', 'Category 3', 'Category 4'];
  s.addChart('bar', [
    { name: 'Series 1', labels, values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels, values: [2.4, 4.4, 1.8, 2.8] },
  ], {
    x: 0.839, y: 3.923, w: 7.993, h: 3.056,
    barDir: 'bar', barGrouping: 'percentStacked', barGapWidthPct: 150, barOverlapPct: 100,
    chartColors: [C.clay, C.sand],
    showLegend: false, showTitle: false, showValue: false,
    catAxisLineShow: true, catAxisLineColor: C.gridline, catAxisLineSize: 0.75,
    catAxisMajorTickMark: 'none', catAxisMinorTickMark: 'none', catAxisLabelPos: 'nextTo',
    catAxisLabelFontFace: SANS, catAxisLabelFontSize: 11.97, catAxisLabelColor: C.axisText,
    valAxisLineShow: false, valAxisMajorTickMark: 'none', valAxisMinorTickMark: 'none',
    valAxisLabelFormatCode: '0%', valAxisLabelPos: 'nextTo',
    valAxisLabelFontFace: SANS, valAxisLabelFontSize: 11.97, valAxisLabelColor: C.axisText,
    valGridLine: { color: C.gridline, size: 0.75 },
    catGridLine: { style: 'none' },
  });
  photoFrame(s, 9.502, 0, 3.832, 7.5, C.white);
}

// 8 — Logo divider
function slide08(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.cream };
  headline(s, 0.715, 4.988, 3.954, 1.393, 'Logo.');
  body(s, 7.493, 5.195, 5.002, 0.978, lorem(25));
  photoFrame(s, 0.839, 0.521, 11.656, 3.697, C.cream);
}

// 9 — Wordmark on light and dark
function slide09(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.paper };
  headline(s, 4.69, 0.884, 3.954, 1.393, 'Logo.', { align: 'center' });
  logoSwatch(s, 0.839, 2.671, 5.75, 2.921, C.linen, C.ink);
  logoSwatch(s, 6.745, 2.685, 5.75, 2.921, C.charcoal, C.white);
  body(s, 2.123, 6.002, 9.086, 0.678, LOREM, { align: 'center' });
  rule(s, 5.217, 6.979, 2.899);
}

// 10 — Wordmark colour grid
function slide10(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.paper };
  microLabel(s, 0.726, 0.413, 'Visual Identity System', 'left');
  microLabel(s, 10.486, 0.413, 'Brand Guideline', 'right');
  [
    { x: 0.839, y: 0.911, fill: C.white, text: C.ink },
    { x: 6.745, y: 0.911, fill: C.clay, text: C.white },
    { x: 0.839, y: 4.058, fill: C.moss, text: C.white },
    { x: 6.745, y: 4.058, fill: C.sand, text: C.ink },
  ].forEach(c => logoSwatch(s, c.x, c.y, 5.75, 2.921, c.fill, c.text));
}

// 11 — Aplication divider
function slide11(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.cream };
  headline(s, 5.411, 1.303, 7.084, 1.393, 'Aplication.');
  body(s, 5.394, 2.968, 6.267, 0.675, lorem(21));
  [
    { y: 4.517, ruleY: 5.190 },
    { y: 5.464, ruleY: 6.073 },
    { y: 6.347 },
  ].forEach((r, i) => indexRow(s, 5.486, r.y, `Aplication ${i + 1}`, `00${i + 1}/003`, r.ruleY));
  photoFrame(s, 0, 0, 3.832, 7.5, C.cream);
}

// 12 — Logo Aplication
function slide12(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.paper };
  microLabel(s, 0.726, 0.413, 'Aplication', 'left');
  headline(s, 0.708, 1.819, 6.526, 2.686, 'Logo Aplication');
  body(s, 0.73, 5.013, 4.008, 0.978, lorem(19));
  [0.521, 2.738, 4.956].forEach(y => photoFrame(s, 8.477, y, 4.008, 2.024, C.paper));
  headline(s, 9.614, 1.186, 2.058, 0.693, 'Logo.', { fontSize: 44, align: 'center', color: C.white });
  headline(s, 11.51, 2.977, 0.951, 0.37, 'Logo.', { fontSize: 20 });
  headline(s, 8.738, 6.405, 0.951, 0.37, 'Logo.', { fontSize: 20, color: C.white });
}

// 13 — Media Aplication; the tablet mock-up is drawn from native shapes
function slide13(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.paper };
  s.addShape('roundRect', {
    x: 4.597, y: 1.100, w: 0.333, h: 0.056, rectRadius: 0.02,
    fill: { color: C.ink }, line: { type: 'none' },
  }); // camera bar peeking above the top edge
  [1.597, 1.917].forEach(y => s.addShape('rect', {
    x: 5.306, y, w: 0.056, h: 0.292, fill: { color: C.ink }, line: { type: 'none' },
  })); // volume keys
  s.addShape('roundRect', {
    x: 0.931, y: 1.132, w: 4.382, h: 5.764, rectRadius: 0.33,
    fill: { type: 'none' }, line: { color: C.ink, width: 0.5 },
  }); // hairline ring tracing just outside the shell
  s.addShape('roundRect', {
    x: 0.951, y: 1.169, w: 4.347, h: 5.696, rectRadius: 0.30,
    fill: { color: C.ink }, line: { type: 'none' },
  });
  s.addShape('roundRect', {
    x: 1.118, y: 1.336, w: 4.014, h: 5.363, rectRadius: 0.21,
    fill: { color: C.paper }, line: { type: 'none' },
  });
  headline(s, 5.679, 2.935, 6.829, 2.686, 'Media Aplication.');
  body(s, 8.734, 0.902, 3.936, 0.978, lorem(16), { bullet: true });
  rule(s, 4.594, 5.925, 3.884);
}

// 14 — Colors divider
function slide14(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.cream };
  masthead(s, 'Brand Guideline', 'Presentation Template');
  headline(s, 0.739, 1.872, 4.844, 1.393, 'Colors.');
  body(s, 0.726, 5.425, 5.089, 0.978, lorem(25));
  [C.sand, C.mist, C.coal].forEach((color, i) => rect(s, 8.095 + i * 1.505, 5.588, 1.39, 1.39, color));
}

// 15 — Colour swatch specimens
function slide15(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.paper };
  rect(s, 0, 4.697, 13.333, 2.803, C.white);
  microLabel(s, 0.726, 0.47, 'Visual Identity System', 'left');
  body(s, 8.832, 0.425, 3.082, 0.675, lorem(11));
  [C.sand, C.mist, C.clay, C.coal].forEach((color, i) => {
    const x = 0.839 + i * 2.718;
    rect(s, x, 1.832, 2.518, 3.805, color);
    s.addText('#D4C0A7', {
      x: x - 0.082, y: 5.78, w: 1.511, h: 0.462,
      fontFace: SANS, fontSize: 16, italic: true, color: C.ink,
      valign: 'top', margin: INSET, lineSpacingMultiple: 1.5,
    });
    body(s, x - 0.082, 6.339, 2.153, 0.675, 'At vero eos et accusamus et iusto odio');
  });
}

// 16 — Typography divider
function slide16(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.cream };
  masthead(s, 'Brand Guideline', 'Presentation Template');
  headline(s, 0.726, 1.977, 7.883, 1.393, 'Typography.');
  body(s, 9.399, 2.197, 3.04, 0.978, lorem(14));
  photoFrame(s, 0, 5.045, 13.333, 2.476, C.cream);
}

// 17 — Playfair Display specimen
function slide17(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.white };
  rect(s, 0.843, 5.262, 11.647, 1.698, C.paper);
  rect(s, 0.848, 1.695, 11.647, 3.568, C.cream);
  s.addText('Playfair Display', {
    x: 0.839, y: 0.783, w: 3.906, h: 0.478,
    fontFace: SERIF, fontSize: 28, color: C.ink, valign: 'top', margin: INSET,
    lineSpacingMultiple: 0.8, bullet: { characterCode: '2022', indent: 36 },
  });
  rule(s, 4.856, 1.022, 8.477, 0.5);
  [
    { y: 1.972, cells: [[1.183, 1.415, 'Aa'], [2.980, 1.415, 'Bb'], [4.777, 1.415, 'Cc'],
                        [6.575, 1.615, 'Dd'], [8.572, 1.415, 'Ef'], [10.370, 1.615, 'Gg']] },
    { y: 3.435, cells: [[1.199, 1.615, 'Hh'], [2.996, 1.415, 'Ii'], [4.794, 1.415, 'Jj'],
                        [6.591, 1.615, 'Kk'], [8.589, 1.415, 'Ll'], [10.386, 1.975, 'Mm']] },
  ].forEach(r => r.cells.forEach(([x, w, t]) => glyph(s, x, r.y, w, 1.42, t, SERIF)));
  [
    { x: 1.216, w: 3.063, text: '0123456789' },
    { x: 8.845, w: 3.640, text: '!@#$%^&*()_+' },
  ].forEach(t => s.addText(t.text, {
    x: t.x, y: 5.742, w: t.w, h: 0.687,
    fontFace: SERIF, fontSize: 32, color: C.ink, valign: 'top', margin: INSET, lineSpacingMultiple: 1.2,
  }));
}

// 18 — Open Sans specimen
function slide18(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.paper };
  rect(s, 0, 1.284, 7.493, 6.216, C.cream);
  s.addText('Open Sans', {
    x: 8.882, y: 0.783, w: 2.993, h: 0.502,
    fontFace: SANS, fontSize: 28, color: C.ink, valign: 'top', margin: INSET,
    lineSpacingMultiple: 0.8, bullet: { characterCode: '2022', indent: 36 },
  });
  rule(s, 0, 0.989, 8.477, 0.5);
  [
    { y: 1.355, cells: [[0.789, 1.901, 'Aa'], [2.704, 1.901, 'Bb'], [4.502, 1.901, 'Cc']] },
    { y: 2.900, cells: [[0.768, 2.171, 'Dd'], [2.766, 1.901, 'Ef'], [4.401, 2.171, 'Gg']] },
    { y: 4.320, cells: [[0.716, 2.171, 'Hh'], [2.683, 1.901, 'Ii'], [4.480, 1.901, 'Jj']] },
    { y: 5.740, cells: [[0.768, 2.171, 'Kk'], [2.766, 1.901, 'Ll'], [4.563, 2.654, 'Mm']] },
  ].forEach(r => r.cells.forEach(([x, w, t]) => glyph(s, x, r.y, w, 1.452, t, SANS)));
  rect(s, 7.482, 4.32, 4.343, 3.18, C.silver);
  [
    { y: 4.914, w: 3.063, text: '0123456789' },
    { y: 6.051, w: 3.640, text: '!@#$%^&*()_+' },
  ].forEach(t => s.addText(t.text, {
    x: 7.916, y: t.y, w: t.w, h: 0.702,
    fontFace: SANS, fontSize: 32, color: C.ink, valign: 'top', margin: INSET, lineSpacingMultiple: 1.2,
  }));
  body(s, 9.455, 1.701, 3.04, 0.978, lorem(14));
}

// 19 — Image Direction moodboard
function slide19(pptx) {
  const s = pptx.addSlide();
  masthead(s, 'Image Direction', 'Presentation Template');
  [
    { x: 0.000, y: 1.028, w: 2.493, h: 4.206 },
    { x: 2.674, y: 1.028, w: 2.182, h: 2.165 },
    { x: 5.036, y: 1.028, w: 2.182, h: 2.165 },
    { x: 2.677, y: 3.347, w: 4.542, h: 1.887 },
    { x: 7.399, y: 1.042, w: 5.934, h: 4.193 },
    { x: 0.000, y: 5.389, w: 8.477, h: 2.111 },
    { x: 8.659, y: 5.389, w: 4.674, h: 2.111 },
  ].forEach(t => photoFrame(s, t.x, t.y, t.w, t.h, C.white));
}

// 20 — Thank you
function slide20(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.cream };
  masthead(s, 'Brand Guideline', 'Presentation Template');
  headline(s, 2.232, 1.617, 8.87, 1.717, 'Thank You!', { align: 'center', lineSpacingMultiple: null });
  body(s, 2.568, 3.771, 8.197, 0.675, lorem(25), { align: 'center' });
  photoFrame(s, 0.839, 5.407, 11.656, 2.114, C.cream);
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE_13_33x7_5', width: 13.3333333, height: 7.5 });
  pptx.layout = 'WIDE_13_33x7_5';
  pptx.title = 'Brand Guideline';
  pptx.theme = { headFontFace: SERIF, bodyFontFace: SANS };

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
   slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach(fn => fn(pptx));

  return pptx.writeFile({
    fileName: path.join(__dirname, '105d886d-a820-4891-b989-9c2deedd974f_grok_final.pptx'),
  });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
