/**
 * "Medico" presentation template — rebuilt with pptxgenjs.
 *
 * Slide size 10" x 5.625" (16:9), 21 slides.
 * Photographs in the original deck are replaced by flat placeholder rectangles;
 * the small vector pictograms are redrawn with the closest native preset shape.
 *
 *   node 0c7af98e-b230-42db-bb35-d09b9ca56787_grok_final.js
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
// Theme "Custom 87" of the source deck.
const C = {
  black: '000000',
  white: 'FFFFFF',
  teal: '40C8CB', // accent1
  blue: '0135C9', // accent2
  purple: '8330DE', // accent3
  orange: 'D35E1D', // accent4
  ink: '262626', // dark slide background
  mute: 'BFBFBF', // muted body copy / hairlines
  rule: 'D8D8D8', // light rules and outlines
  hair: 'A5A5A5', // thin vertical accents
  panel: 'F2F2F2', // light grey panels / skill-bar troughs
  ghost: 'D9D9D8', // faint hexagon outline
  violet: '7351A1', // hexagon outline on the cover
  axis: '595959' // chart axis + legend labels
};

/* -------------------------------------------------------------------- fonts */
const F = {
  display: 'Poppins SemiBold',
  displayLt: 'Poppins',
  body: 'Inter',
  semi: 'Inter SemiBold',
  medium: 'Inter Medium',
  mont: 'Montserrat'
};

/* ------------------------------------------------------------- text insets */
// The deck uses 68575 / 34275 EMU insets (5.4pt / 2.7pt) on most text boxes and
// 51425 / 25700 EMU (4.05pt / 2.02pt) on a few tight ones. [L, R, B, T] in points.
const PAD = [5.4, 5.4, 2.7, 2.7];
const PAD_SM = [4.05, 4.05, 2.02, 2.02];

/* ------------------------------------------------------ repeated copy blocks */
const NB = '\u00a0'; // the deck writes "Lorem Ipsum<nbsp>is simply..."
const L = {
  cover: `Lorem ipsum${NB}is simply dummy text of the printing and typesetting`,
  twoLine: [
    `Lorem Ipsum${NB}is simply dummy text of the printing typesetting industry. Lorem Ipsum has been type`,
    `survived  Lorem Ipsum${NB}is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been type and survived`
  ].join('\n'),
  oneLine: `Lorem Ipsum${NB}is simply dummy text of the printing typesetting industry. Lorem Ipsum has been type survived  Lorem Ipsum${NB}is typesetting industry. Lorem has been type and survived`,
  wide: `Lorem Ipsum${NB}is simply dummy text of the printing typesetting industry. Lorem Ipsum has been type survived  Lorem Ipsum${NB}is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been type and survived`,
  tiny: `Lorem Ipsum${NB}is simply dummy text of the printing and typesetting`,
  centuries: `Lorem Ipsum${NB}is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been type and survived not only five centuries,`,
  industry: `Lorem Ipsum${NB}is simply dummy text of the printing and typesetting industry`,
  ipsumHas: `Lorem Ipsum${NB}is simply dummy text of the printing and typesetting industry. Lorem ipsum has`,
  hasBeen: `Lorem Ipsum${NB}is simply dummy text of the printing and typesetting industry. Lorem Ipsum has`,
  stat: `Lorem Ipsum${NB}is simply dummy text of the\nprinting and typesetting industry. Lorem`,
  survivedSimply: `Lorem Ipsum${NB}is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been type and survived  Lorem Ipsum${NB}is simply`,
  simplyDummy: `Lorem Ipsum${NB}is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been type and survived  Lorem Ipsum${NB}is simply dummy`,
  andLorem: `Lorem Ipsum${NB}is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been type and survived  Lorem`,
  short: `Lorem Ipsum${NB}is simply dummy text of the printing typesetting industry.`,
  caption: `Lorem Ipsum${NB}is simply dummy text of the printing typesetting industry. Lorem Ipsum has been type survived  Lorem Ipsum${NB}is typesetting industry.`,
  variations: 'There are many variations of passages of Lorem Ipsum available',
  variationsLong: 'There are many variations of passages of Lorem Ipsum available, but the majority have',
  email: 'lorem@example.com'
};

/* ------------------------------------------------------------------ helpers */

/** Text box. Defaults mirror the deck: top anchored, left aligned, Inter 8pt black. */
function tx(slide, text, o) {
  slide.addText(text, Object.assign({ valign: 'top', margin: PAD, fontFace: F.body, color: C.black }, o));
}

/** Section headline: Poppins SemiBold 26pt sitting on the deck's title grid. */
function title(slide, text, o) {
  tx(slide, text, Object.assign({ x: 0.347, y: 0.852, h: 0.776, fontFace: F.display, fontSize: 26, lineSpacingMultiple: 0.8 }, o));
}

/** Recurring "Write Your / Headline Here" kicker. */
function kicker(slide, x, y) {
  tx(slide, [
    { text: 'Write Your', options: { fontFace: F.displayLt, breakLine: true } },
    { text: 'Headline Here', options: { fontFace: F.display } }
  ], { x, y, w: 2.325, h: 0.53, fontSize: 15, lineSpacingMultiple: 0.9 });
}

/** Small-caps style email line at the bottom of the left column. */
function emailLine(slide, y) {
  tx(slide, L.email, { x: 0.347, y, w: 2.325, h: 0.227, fontSize: 8, lineSpacingMultiple: 1.6 });
}

/** Filled shape shorthand. */
function fill(slide, shape, x, y, w, h, color, extra) {
  slide.addShape(shape, Object.assign({ x, y, w, h, fill: { color } }, extra));
}

/**
 * Pictogram stand-in. The deck's medical glyphs are thin line drawings, so the
 * nearest preset shape is stroked rather than filled to keep the same weight.
 * `g` is [presetShape, x, y, w, h].
 */
function icon(slide, g, color) {
  slide.addShape(g[0], { x: g[1], y: g[2], w: g[3], h: g[4], line: { color: color, width: 1 } });
}

/**
 * Vertices of a pointy-top hexagon, as offsets inside a shape box.
 * `custGeom` path coordinates are relative to the shape's own origin.
 */
function hexPath(ox, oy, w, h) {
  return [
    { x: ox, y: oy + h * 0.25, moveTo: true },
    { x: ox + w / 2, y: oy },
    { x: ox + w, y: oy + h * 0.25 },
    { x: ox + w, y: oy + h * 0.75 },
    { x: ox + w / 2, y: oy + h },
    { x: ox, y: oy + h * 0.75 },
    { close: true }
  ];
}

function hex(slide, x, y, w, h, o) {
  slide.addShape('custGeom', Object.assign({ x, y, w, h, points: hexPath(0, 0, w, h) }, o));
}

/**
 * Photo placeholder. The source deck leaves its picture frames empty, so the
 * marker is kept deliberately light: a hairline frame plus an "[image]" tag.
 */
function photo(slide, x, y, w, h) {
  slide.addShape('rect', { x, y, w, h, line: { color: C.panel, width: 0.75 } });
  tx(slide, '[image]', { x, y: y + h / 2 - 0.13, w, h: 0.26, align: 'center', fontSize: 9, color: C.rule });
}

/**
 * Hexagon-cropped photo slots. The deck paints a solid panel and punches
 * hexagonal windows through it, so the picture behind shows only inside the
 * hexagons — one custom-geometry shape, outer rectangle plus hexagon subpaths.
 */
function hexWindows(slide, panel, holes) {
  const pts = [
    { x: 0, y: 0, moveTo: true },
    { x: 0, y: panel.h },
    { x: panel.w, y: panel.h },
    { x: panel.w, y: 0 },
    { close: true }
  ];
  holes.forEach(function (b) {
    hexPath(b[0] - panel.x, b[1] - panel.y, b[2], b[3]).forEach(function (p) { pts.push(p); });
  });
  slide.addShape('custGeom', { x: panel.x, y: panel.y, w: panel.w, h: panel.h, fill: { color: panel.color }, points: pts });
  holes.forEach(function (b) {
    tx(slide, '[image]', { x: b[0], y: b[1] + b[3] / 2 - 0.13, w: b[2], h: 0.26, align: 'center', fontSize: 8, color: C.rule });
  });
}

// Normalised spine of the ECG trace used on the cover and the mission slide.
const ECG = [
  [0.00, 0.57], [0.13, 0.57], [0.21, 0.33], [0.25, 0.30], [0.29, 0.34], [0.345, 0.90],
  [0.40, 0.06], [0.44, 0.02], [0.48, 0.06], [0.55, 0.60], [0.59, 0.45], [0.63, 0.42],
  [0.67, 0.46], [0.73, 0.72], [0.78, 0.58], [0.83, 0.56], [1.00, 0.56]
];

function ecg(slide, x, y, w, h, width) {
  slide.addShape('custGeom', {
    x, y, w, h,
    line: { color: C.black, width: width },
    points: ECG.map(function (p, i) { return { x: p[0] * w, y: p[1] * h, moveTo: i === 0 }; })
  });
}

// Twitter / Facebook / Google+ marks: three tiny glyphs on a shared baseline.
const SOCIAL = [[0.000, 0.096, 0.078], [0.253, 0.071, 0.071], [0.476, 0.105, 0.067]];

function socialRow(slide, x, y, color) {
  SOCIAL.forEach(function (g) { fill(slide, 'ellipse', x + g[0], y, g[1], g[2], color); });
}

/* -------------------------------------------------------------------- deck */

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'MEDICO', width: 10, height: 5.625 });
pptx.layout = 'MEDICO';
pptx.author = 'Medico';
pptx.title = 'Medico — Presentation Template';

/* --- 1. Cover ------------------------------------------------------------ */
function slide01() {
  const s = pptx.addSlide();
  hexWindows(s, { x: 0.007, y: 0, w: 9.995, h: 5.625, color: C.white },
    [[4.664, 1.042, 1.749, 2.018], [5.583, 2.69, 1.749, 2.018], [7.496, 2.69, 1.749, 2.018]]);

  tx(s, 'Medico', { x: 0.421, y: 2.854, w: 4.697, h: 1.338, fontFace: F.display, fontSize: 75 });
  tx(s, 'PRESENTATION TEMPLATE', { x: 0.464, y: 4.016, w: 3.303, h: 0.29, fontFace: F.medium, fontSize: 13 });
  s.addShape('line', { x: 0.547, y: 0.427, w: 0, h: 0.727, line: { color: C.hair, width: 0.75 } });
  tx(s, 'Syble Vane', { x: 0.459, y: 1.581, w: 1.418, h: 0.24, fontFace: F.semi, fontSize: 10, bold: true });
  tx(s, L.cover, { x: 0.463, y: 1.8, w: 2.277, h: 0.372, fontSize: 8, lineSpacingMultiple: 1.2 });

  hex(s, 6.78, 0.346, 2.691, 2.711, { line: { color: C.violet, width: 0.75 } });
  hex(s, 6.558, 1.042, 1.75, 2.018, { fill: { color: C.teal } });
  ecg(s, 7.841, 1.862, 1.009, 0.34, 2.25);
  tx(s, '30\nJuly', {
    x: 6.904, y: 1.718, w: 0.855, h: 0.707, align: 'center',
    fontFace: F.semi, fontSize: 23, bold: true, lineSpacingMultiple: 1.0
  });
}

/* --- 2. Welcome ---------------------------------------------------------- */
function slide02() {
  const s = pptx.addSlide();
  hexWindows(s, { x: 0.003, y: 0.003, w: 9.995, h: 5.625, color: C.white }, [[4.475, 1.578, 2.853, 3.29]]);

  title(s, 'Welcome To Medico', { w: 2.966 });
  kicker(s, 0.347, 2.471);
  tx(s, L.twoLine, { x: 0.347, y: 3.2, w: 2.891, h: 0.911, fontSize: 8, lineSpacingMultiple: 1.3 });
  tx(s, `Lorem Ipsum${NB}is simply dummy text of the printing and typesetting industry. Lorem has been`,
    { x: 7.782, y: 2.676, w: 1.04, h: 1.18, fontSize: 8, lineSpacingMultiple: 1.2 });

  hex(s, 7.782, 4.192, 0.51, 0.589, { fill: { color: C.blue } });
  hex(s, 5.107, 0.438, 3.997, 4.617, { line: { color: C.ghost, width: 0.75 } });
}

/* --- 3. About (dark) ----------------------------------------------------- */
function slide03() {
  const s = pptx.addSlide();
  fill(s, 'rect', 0, 0, 10, 5.625, C.ink);
  photo(s, 0.442, 2.147, 3.951, 3.478);
  title(s, 'About To Medico', { w: 2.736, color: C.white });

  [
    { y: 1.412, glyph: ['heart', 5.217, 1.477, 0.466, 0.36] },      // heart-with-pulse
    { y: 3.411, glyph: ['donut', 5.26, 3.466, 0.38, 0.431] }        // stethoscope
  ].forEach(function (row) {
    tx(s, 'Elements Text', { x: 6.078, y: row.y, w: 1.689, h: 0.303, fontFace: F.semi, fontSize: 14, bold: true, color: C.white });
    tx(s, L.centuries, { x: 6.077, y: row.y + 0.62, w: 3.261, h: 0.56, fontSize: 8, color: C.mute, lineSpacingMultiple: 1.3 });
    s.addShape('line', { x: 6.16, y: row.y + 0.484, w: 0.385, h: 0, line: { color: C.rule, width: 1.5 } });
    icon(s, row.glyph, C.white);
  });
}

/* --- 4. Mission & vision ------------------------------------------------- */
function slide04() {
  const s = pptx.addSlide();
  title(s, 'Our Mission And Vision', { w: 2.966 });
  tx(s, 'Our Mission', { x: 0.347, y: 2.513, w: 2.325, h: 0.332, fontFace: F.display, fontSize: 15, lineSpacingMultiple: 1.25 });
  tx(s, L.twoLine, { x: 0.347, y: 3.067, w: 2.891, h: 0.911, fontSize: 8, lineSpacingMultiple: 1.3 });
  tx(s, 'Our Mission', { x: 4.406, y: 3.034, w: 2.325, h: 0.332, fontFace: F.display, fontSize: 12, lineSpacingMultiple: 1.5625 });
  tx(s, L.wide, { x: 4.406, y: 3.44, w: 4.867, h: 0.588, fontSize: 8, lineSpacingMultiple: 1.4 });
  tx(s, '$1899/ 35K', { x: 6.119, y: 4.625, w: 2.325, h: 0.332, fontFace: F.display, fontSize: 15, lineSpacingMultiple: 1.25 });

  ecg(s, 0.5, 4.667, 0.883, 0.248, 1.75);
  s.addShape('line', { x: 1.923, y: 4.804, w: 3.656, h: 0, line: { color: C.mute, width: 1 } });
  photo(s, 4.484, 0.852, 1.922, 1.726);
  photo(s, 7.25, 0.852, 1.922, 1.726);
}

/* --- 5. Why choose ------------------------------------------------------- */
function slide05() {
  const s = pptx.addSlide();
  photo(s, 7.202, 0, 2.798, 5.625);
  title(s, 'Why Choose Medico', { w: 2.966 });
  tx(s, 'Example Here', { x: 0.347, y: 2.513, w: 2.325, h: 0.332, fontFace: F.display, fontSize: 15, lineSpacingMultiple: 1.25 });
  tx(s, L.twoLine, { x: 0.347, y: 3.067, w: 2.891, h: 0.911, fontSize: 8, lineSpacingMultiple: 1.3 });

  [
    { head: 'Honesty', y: 1.524, w: 1.314 },
    { head: 'Hard working', y: 2.812, w: 1.522 },
    { head: 'Simplicity', y: 4.101, w: 1.251 }
  ].forEach(function (it) {
    tx(s, it.head, { x: 4.031, y: it.y, w: it.w, h: 0.303, fontFace: F.semi, fontSize: 14, bold: true });
    tx(s, L.tiny, { x: 4.031, y: it.y + 0.337, w: 2.55, h: 0.397, fontFace: F.mont, fontSize: 8, lineSpacingMultiple: 1.45 });
  });
}

/* --- 6. What we do ------------------------------------------------------- */
function slide06() {
  const s = pptx.addSlide();
  hexWindows(s, { x: 0.003, y: -0.003, w: 9.995, h: 5.625, color: C.white },
    [[5.18, 0.805, 1.75, 2.018], [7.574, 0.805, 1.75, 2.018]]);
  fill(s, 'rect', 0, 0, 4.479, 5.625, C.blue);

  tx(s, 'What We Do', { x: 0.347, y: 1.352, w: 2.966, h: 0.455, fontFace: F.display, fontSize: 26, color: C.white, lineSpacingMultiple: 1.0 });
  tx(s, [
    `Lorem Ipsum${NB}is simply dummy text of`,
    ' the printing typesetting industry. Lorem',
    ' Ipsum has been type',
    `survived  Lorem Ipsum${NB}is simply dummy`,
    ' text of the printing and typesetting',
    ' industry. Lorem Ipsum has been type',
    ' and survived'
  ].join('\n'), {
    x: 0.347, y: 2.6, w: 3.378, h: 2.007, fontSize: 11, color: C.white,
    lineSpacingMultiple: 1.571, bullet: { characterCode: '2714', indent: 17 }
  });

  [{ x: 5.308, xd: 5.248, wd: 1.613, v: '$2.80' }, { x: 7.703, xd: 7.703, wd: 1.493, v: '6.35K' }].forEach(function (st) {
    tx(s, st.v, { x: st.x, y: 3.432, w: 1.493, h: 0.425, align: 'center', fontFace: F.semi, fontSize: 19, bold: true, lineSpacingMultiple: 1.4 });
    tx(s, L.stat, { x: st.xd, y: 3.969, w: st.wd, h: 0.695, align: 'center', fontSize: 8, lineSpacingMultiple: 1.2 });
  });
}

/* --- 7. Always help you -------------------------------------------------- */
function slide07() {
  const s = pptx.addSlide();
  photo(s, 4.062, 0.674, 3.266, 4.277);
  fill(s, 'roundRect', 6.422, 3.042, 3.066, 1.202, C.orange, { rectRadius: 0.2 });
  fill(s, 'roundRect', 6.422, 1.38, 3.066, 1.202, C.purple, { rectRadius: 0.2 });

  title(s, 'Always Help You carefully', { w: 2.966 });
  tx(s, 'Example Here', { x: 0.347, y: 2.545, w: 2.325, h: 0.332, fontFace: F.display, fontSize: 15, lineSpacingMultiple: 1.25 });
  tx(s, L.twoLine, { x: 0.347, y: 3.067, w: 2.891, h: 0.911, fontSize: 8, lineSpacingMultiple: 1.3 });

  [1.675, 3.289].forEach(function (y) {
    tx(s, 'Example Service', { x: 6.831, y: y, w: 1.689, h: 0.252, fontFace: F.semi, fontSize: 11, bold: true, color: C.white });
    tx(s, L.industry, { x: 6.831, y: y + 0.266, w: 2.374, h: 0.399, fontSize: 8, color: C.white, lineSpacingMultiple: 1.4 });
  });
}

/* --- 8. Services list ---------------------------------------------------- */
function slide08() {
  const s = pptx.addSlide();
  title(s, 'Medico Services', { w: 2.523 });
  kicker(s, 0.347, 2.471);
  tx(s, L.oneLine, { x: 0.347, y: 3.2, w: 2.397, h: 0.911, fontSize: 8, lineSpacingMultiple: 1.3 });

  [
    { y: 1.726, head: 'Ambulance', ring: C.mute, glyph: ['round2SameRect', 6.766, 1.93, 0.208, 0.172] },
    { y: 2.881, head: 'Body Test', ring: C.blue, glyph: ['flowChartConnector', 6.788, 3.067, 0.186, 0.208] },
    { y: 4.037, head: 'Medicine', ring: C.mute, glyph: ['roundRect', 6.791, 4.223, 0.171, 0.208] }
  ].forEach(function (row) {
    s.addShape('ellipse', { x: 6.58, y: row.y, w: 0.581, h: 0.581, fill: { color: C.white }, line: { color: row.ring, width: 1 } });
    tx(s, row.head, { x: 7.435, y: row.y - 0.029, w: 1.689, h: 0.261, fontFace: F.semi, fontSize: 11, bold: true });
    tx(s, L.ipsumHas, { x: 7.435, y: row.y + 0.236, w: 2.123, h: 0.533, fontSize: 8, lineSpacingMultiple: 1.2 });
    icon(s, row.glyph, C.black);
  });

  photo(s, 3.302, 1.063, 2.788, 4.062);
}

/* --- 9. Service card ----------------------------------------------------- */
function slide09() {
  const s = pptx.addSlide();
  title(s, 'Medico Services', { w: 2.523 });
  fill(s, 'rect', 3.858, 0.952, 2.336, 1.804, C.purple);
  kicker(s, 0.347, 2.471);
  tx(s, L.oneLine, { x: 0.347, y: 3.2, w: 2.397, h: 0.911, fontSize: 8, lineSpacingMultiple: 1.3 });

  tx(s, 'Service Text', { x: 4.32, y: 1.648, w: 1.411, h: 0.278, align: 'center', fontFace: F.semi, fontSize: 12, bold: true, color: C.white });
  tx(s, L.variations, { x: 4.134, y: 1.887, w: 1.783, h: 0.533, align: 'center', fontSize: 8, color: C.white, lineSpacingMultiple: 1.2 });
  icon(s, ['moon', 4.92, 1.275, 0.212, 0.278], C.white); // stomach

  photo(s, 6.647, 0.952, 2.336, 1.804);
  photo(s, 3.858, 3.16, 2.336, 1.804);
  photo(s, 6.647, 3.16, 2.336, 1.804);
}

/* --- 10. Four services --------------------------------------------------- */
function slide10() {
  const s = pptx.addSlide();
  photo(s, 0, 0, 3.969, 5.625);

  [
    { n: '02', dot: [7.544, 1.074], color: C.blue, tx: 7.48, head: 'Services Two', hw: 1.522, dw: 1.707 },
    { n: '01', dot: [4.967, 1.074], color: C.teal, tx: 4.863, head: 'Services One', hw: 1.418, dw: 1.707 },
    { n: '03', dot: [4.928, 3.136], color: C.purple, tx: 4.863, head: 'Services Three', hw: 1.707, dw: 1.634 },
    { n: '03', dot: [7.609, 3.136], color: C.orange, tx: 7.544, head: 'Service Four', hw: 1.854, dw: 1.707 }
  ].forEach(function (it) {
    const dy = it.dot[1];
    fill(s, 'ellipse', it.dot[0], dy, 0.494, 0.494, it.color);
    tx(s, it.n, { x: it.dot[0], y: dy + 0.081, w: 0.494, h: 0.303, align: 'center', fontFace: F.semi, fontSize: 14, bold: true, color: C.white });
    tx(s, it.head, { x: it.tx, y: dy + 0.683, w: it.hw, h: 0.311, fontFace: F.semi, fontSize: 14, bold: true });
    tx(s, L.tiny, { x: it.tx, y: dy + 1.02, w: it.dw, h: 0.56, fontSize: 8, lineSpacingMultiple: 1.3 });
  });

  fill(s, 'rect', 3.429, 1.683, 0.95, 2.259, C.white);
  tx(s, 'Services', {
    x: 2.63, y: 2.585, w: 2.523, h: 0.455, rotate: -90, align: 'center',
    fontFace: F.display, fontSize: 26, lineSpacingMultiple: 1.0
  });
}

/* --- 11. Best quality clinic --------------------------------------------- */
function slide11() {
  const s = pptx.addSlide();
  fill(s, 'rect', 6.262, 0, 3.738, 5.625, C.orange);
  photo(s, 4.029, 0.852, 5.156, 4.112);
  title(s, 'Our Best Quality Clinic', { w: 3.046 });

  [{ head: 'Hard working', y: 2.51, w: 1.522 }, { head: 'Simplicity', y: 3.798, w: 1.251 }].forEach(function (it) {
    tx(s, it.head, { x: 0.353, y: it.y, w: it.w, h: 0.303, fontFace: F.semi, fontSize: 14, bold: true });
    tx(s, L.tiny, { x: 0.353, y: it.y + 0.336, w: 2.55, h: 0.385, fontSize: 8, lineSpacingMultiple: 1.3 });
  });
}

/* --- 12. Enovation ------------------------------------------------------- */
function slide12() {
  const s = pptx.addSlide();
  title(s, 'Medico Enovation', { x: 6.328, y: 0.986, w: 3.046 });
  tx(s, '$689', { x: 0.571, y: 2.788, w: 1.493, h: 0.444, fontFace: F.semi, fontSize: 23, bold: true, lineSpacingMultiple: 1.167 });
  tx(s, L.stat, { x: 0.591, y: 3.326, w: 1.493, h: 0.695, fontSize: 8, lineSpacingMultiple: 1.2 });

  [2.408, 3.753].forEach(function (y) {
    tx(s, 'Lorem 02', { x: 6.328, y: y, w: 1.493, h: 0.278, fontFace: F.semi, fontSize: 12, bold: true });
    tx(s, L.survivedSimply, { x: 6.332, y: y + 0.312, w: 3.042, h: 0.588, fontSize: 8, lineSpacingMultiple: 1.4 });
  });

  photo(s, 0.666, 0.548, 1.373, 1.425);
  photo(s, 2.491, 0.548, 3.217, 4.488);
}

/* --- 13. Nursing and health care ----------------------------------------- */
function slide13() {
  const s = pptx.addSlide();
  photo(s, 7.5, 1.359, 2.5, 3.748);
  title(s, 'Nursing And Health care', { w: 3.046 });
  fill(s, 'rect', 3.69, 1.359, 3.299, 3.748, C.panel);

  // Medical bag, nurse and doctor pictograms on coloured discs.
  [[4.164, C.teal, ['rect', 4.411, 2.39, 0.208, 0.184]],
   [4.993, C.blue, ['flowChartConnector', 5.255, 2.371, 0.18, 0.221]],
   [5.823, C.purple, ['flowChartConnector', 6.099, 2.38, 0.152, 0.204]]].forEach(function (d) {
    fill(s, 'ellipse', d[0], 2.13, 0.704, 0.704, d[1]);
    icon(s, d[2], C.white);
  });

  tx(s, 'Example Text', { x: 4.075, y: 3.256, w: 1.493, h: 0.278, fontFace: F.semi, fontSize: 12, bold: true });
  tx(s, L.simplyDummy, { x: 4.079, y: 3.534, w: 2.539, h: 0.736, fontSize: 8, lineSpacingMultiple: 1.3 });

  tx(s, [
    `Lorem Ipsum${NB}is simply dummy`,
    ' the printing typesetting industry.',
    ' Ipsum has been type',
    `survived  Lorem Ipsum${NB}is simply`,
    'industry. Lorem Ipsum has',
    ' and survived lore'
  ].join('\n'), {
    x: 0.347, y: 2.6, w: 3.184, h: 1.372, fontSize: 11,
    lineSpacingMultiple: 1.2, bullet: { characterCode: '2022', indent: 17 }
  });

  s.addShape('roundRect', { x: 0.424, y: 4.348, w: 1.066, h: 0.309, rectRadius: 0.1545, line: { color: C.rule, width: 1 } });
  tx(s, 'Medico', { x: 0.52, y: 4.364, w: 0.874, h: 0.252, align: 'center', fontFace: F.medium, fontSize: 11, color: C.blue });
}

/* --- 14. Advantages (teardrops) ------------------------------------------ */
function slide14() {
  const s = pptx.addSlide();
  const TILT = -46.104;
  title(s, 'Medico Advantage', { w: 4.161, h: 0.455, lineSpacingMultiple: 1.0 });
  s.addShape('teardrop', { x: 4.362, y: 1.745, w: 4.554, h: 4.554, rotate: TILT, line: { color: C.rule, width: 1 } });

  [
    { x: 4.256, y: 2.526, color: C.teal, label: [4.387, 2.992, 0.837], glyph: ['leftRightArrow', 4.691, 2.772, 0.215, 0.203] },
    { x: 5.246, y: 1.01, color: C.blue, label: [5.431, 1.487, 0.775], glyph: ['moon', 5.739, 1.204, 0.159, 0.208] },
    { x: 6.949, y: 1.299, color: C.purple, label: [7.121, 1.827, 0.775], glyph: ['cloud', 7.404, 1.559, 0.208, 0.206] },
    { x: 8.119, y: 2.526, color: C.orange, label: [8.281, 3.023, 0.82], glyph: ['flowChartDocument', 8.586, 2.721, 0.211, 0.231] }
  ].forEach(function (d) {
    fill(s, 'teardrop', d.x, d.y, 1.079, 1.079, d.color, { rotate: TILT });
    icon(s, d.glyph, C.white);
    tx(s, 'Some Text', { x: d.label[0], y: d.label[1], w: d.label[2], h: 0.202, align: 'center', fontSize: 8, color: C.white });
  });

  [
    { n: 'Advantage 01  ', y: 1.775, color: C.teal, w: 2.841 },
    { n: 'Advantage 02  ', y: 2.619, color: C.blue, w: 2.755 },
    { n: 'Advantage 03  ', y: 3.463, color: C.purple, w: 3.019 },
    { n: 'Advantage 04  ', y: 4.308, color: C.orange, w: 3.019 }
  ].forEach(function (it) {
    tx(s, it.n, { x: 0.4, y: it.y, w: 1.584, h: 0.278, fontFace: F.semi, fontSize: 12, bold: true, color: it.color });
    tx(s, L.hasBeen, { x: 0.4, y: it.y + 0.266, w: it.w, h: 0.372, fontSize: 8, lineSpacingMultiple: 1.2 });
  });
}

/* --- 15. Special doctor -------------------------------------------------- */
function slide15() {
  const s = pptx.addSlide();
  photo(s, 5.495, 0.852, 3.875, 4.2);
  title(s, 'Special Doctor', { w: 4.059, h: 0.455, lineSpacingMultiple: 1.0 });

  tx(s, 'Profession', { x: 2.313, y: 1.913, w: 0.912, h: 0.227, fontFace: F.semi, fontSize: 9, bold: true });
  tx(s, 'Henry Wotton', { x: 2.313, y: 2.067, w: 2.096, h: 0.309, margin: PAD_SM, fontSize: 15, bold: true });
  tx(s, L.andLorem, { x: 2.321, y: 2.414, w: 2.364, h: 0.38, margin: PAD_SM, fontSize: 6, lineSpacingMultiple: 1.1 });

  s.addShape('roundRect', { x: 2.38, y: 3.03, w: 0.901, h: 0.225, rectRadius: 0.1125, line: { color: C.rule, width: 1 } });
  socialRow(s, 2.54, 3.104, C.ink);

  [{ y: 3.731, bar: 4.049, color: C.teal }, { y: 4.273, bar: 4.539, color: C.blue }].forEach(function (sk) {
    tx(s, 'Skills Title', { x: 2.36, y: sk.y, w: 0.981, h: 0.202, fontFace: F.semi, fontSize: 8, bold: true });
    fill(s, 'rect', 2.426, sk.bar, 1.307, 0.078, sk.color);
    fill(s, 'rect', 3.732, sk.bar, 0.736, 0.078, C.panel);
  });
}

/* --- 16. Special team ---------------------------------------------------- */
function slide16() {
  const s = pptx.addSlide();
  hexWindows(s, { x: 3.264, y: 0.883, w: 6.733, h: 4.74, color: C.purple },
    [[3.987, 1.7, 1.322, 1.524], [5.899, 1.7, 1.322, 1.524], [7.81, 1.7, 1.322, 1.524]]);

  title(s, 'Special Team', { w: 2.58 });
  kicker(s, 0.347, 2.471);
  tx(s, L.oneLine, { x: 0.347, y: 3.2, w: 2.371, h: 0.911, fontSize: 8, lineSpacingMultiple: 1.3 });
  emailLine(s, 4.474);

  [{ x: 3.912, dx: 3.92, dw: 1.623, soc: 3.98 },
   { x: 5.837, dx: 5.844, dw: 1.623, soc: 5.904 },
   { x: 7.743, dx: 7.751, dw: 1.725, soc: 7.81 }].forEach(function (m) {
    tx(s, 'Profession', { x: m.x, y: 3.689, w: 1.167, h: 0.227, fontFace: F.semi, fontSize: 9, bold: true, color: C.white });
    tx(s, 'Robert William', { x: m.x, y: 3.842, w: 1.495, h: 0.259, margin: PAD_SM, fontFace: F.semi, fontSize: 12, bold: true, color: C.white });
    tx(s, L.tiny, { x: m.dx, y: 4.168, w: m.dw, h: 0.299, margin: PAD_SM, fontSize: 6, color: C.white, lineSpacingMultiple: 1.2 });
    socialRow(s, m.soc, 4.569, C.white);
  });
}

/* --- 17. Facilities list ------------------------------------------------- */
function slide17() {
  const s = pptx.addSlide();
  title(s, 'Our Facilities', { w: 4.059, h: 0.455, lineSpacingMultiple: 1.0 });
  tx(s, 'Our Facilities', { x: 0.355, y: 2.274, w: 1.352, h: 0.278, fontFace: F.semi, fontSize: 12, bold: true });

  [[2.8, 2.845], [3.302, 3.356], [3.803, 3.866], [4.305, 4.375]].forEach(function (row) {
    tx(s, L.tiny, { x: 0.687, y: row[0], w: 2.163, h: 0.36, margin: PAD_SM, fontSize: 8, lineSpacingMultiple: 1.2 });
    icon(s, ['rect', 0.434, row[1], 0.161, 0.141], C.black); // pencil-on-page bullet
  });

  photo(s, 3.621, 1.396, 3.011, 3.707);
  s.addShape('line', { x: 7.276, y: 1.762, w: 0, h: 0.727, line: { color: C.hair, width: 0.75 } });
  tx(s, '890K', { x: 7.171, y: 2.922, w: 1.942, h: 0.415, fontFace: F.semi, fontSize: 19, bold: true, lineSpacingMultiple: 1.4 });
  tx(s, L.oneLine, { x: 7.171, y: 3.498, w: 2.418, h: 0.911, fontSize: 8, lineSpacingMultiple: 1.3 });
}

/* --- 18. Facilities cards ------------------------------------------------ */
function slide18() {
  const s = pptx.addSlide();
  photo(s, 5.847, 0.884, 4.153, 4.741);
  title(s, 'Our Facilities', { w: 4.059, h: 0.455, lineSpacingMultiple: 1.0 });

  [
    { ring: 0.626, name: 0.54, body: 0.54, glyph: ['moon', 0.806, 2.279, 0.464, 0.424], color: C.black },
    { ring: 2.343, name: 2.267, body: 2.265, glyph: ['heart', 2.483, 2.319, 0.544, 0.345], color: C.blue },
    { ring: 4.029, name: 3.945, body: 3.944, glyph: ['funnel', 4.193, 2.288, 0.494, 0.408], color: C.black }
  ].forEach(function (card) {
    icon(s, card.glyph, card.color);
    s.addShape('roundRect', { x: card.ring, y: 2.08, w: 0.824, h: 0.824, rectRadius: 0.412, line: { color: C.rule, width: 2 } });
    tx(s, 'Example Name', { x: card.name, y: 3.244, w: 1.04, h: 0.429, fontFace: F.semi, fontSize: 11, bold: true });
    tx(s, L.short, { x: card.body, y: 3.748, w: 1.406, h: 0.736, fontSize: 8, lineSpacingMultiple: 1.3 });
  });
}

/* --- 19. Opportunities --------------------------------------------------- */
function slide19() {
  const s = pptx.addSlide();
  fill(s, 'rect', 6.188, 0.011, 3.812, 5.625, C.blue);
  title(s, 'Medico Opportunities', { w: 4.059 });
  kicker(s, 0.347, 2.471);
  tx(s, L.oneLine, { x: 0.347, y: 3.2, w: 2.404, h: 0.911, fontSize: 8, lineSpacingMultiple: 1.3 });
  emailLine(s, 4.474);
  photo(s, 3.294, 2.194, 2.404, 2.9);

  [
    { head: 'Hart 90%', y: 0.697, hw: 1.278, glyph: ['cloud', 7.233, 0.721, 0.225, 0.223] },       // lungs
    { head: 'Cardio.. 60%', y: 2.244, hw: 1.524, glyph: ['plus', 7.185, 2.274, 0.26, 0.26] },      // caduceus
    { head: 'Stomach 99%', y: 3.868, hw: 1.524, glyph: ['moon', 7.236, 3.929, 0.158, 0.207] }
  ].forEach(function (row) {
    tx(s, row.head, { x: 7.546, y: row.y, w: row.hw, h: 0.303, fontFace: F.medium, fontSize: 14, color: C.white });
    tx(s, L.variationsLong, { x: 7.129, y: row.y + 0.42, w: 2.141, h: 0.56, fontSize: 8, color: C.white, lineSpacingMultiple: 1.3 });
    icon(s, row.glyph, C.white);
  });
}

/* --- 20. Data chart ------------------------------------------------------ */
function slide20() {
  const s = pptx.addSlide();
  title(s, 'Data Chart', { w: 4.059, h: 0.455, lineSpacingMultiple: 1.0 });
  kicker(s, 0.347, 2.471);
  tx(s, L.oneLine, { x: 0.347, y: 3.2, w: 2.53, h: 0.911, fontSize: 8, lineSpacingMultiple: 1.3 });
  emailLine(s, 4.474);
  tx(s, L.caption, { x: 4.075, y: 4.368, w: 4.832, h: 0.385, fontSize: 8, lineSpacingMultiple: 1.3 });

  // Value axis: 6 .. 0, one label every 0.44"
  ['6', '5', '4', '3', '2', '1', '0'].forEach(function (v, i) {
    tx(s, v, {
      x: 3.858, y: 0.868 + i * 0.4395, w: 0.433, h: 0.227,
      align: 'right', fontSize: 8, color: C.axis, lineSpacingMultiple: 1.6
    });
  });

  // Columns grow upward from the zero line at y = 3.641
  [
    { name: 'Series 1', color: C.teal, bar: 4.669, top: 1.747, sw: 5.276, lab: 5.295, lw: 0.574 },
    { name: 'Series 2', color: C.blue, bar: 5.68, top: 1.438, sw: 5.856, lab: 5.876, lw: 0.623 },
    { name: 'Series 3', color: C.purple, bar: 6.697, top: 2.758, sw: 6.449, lab: 6.466, lw: 0.6 },
    { name: 'Series 4', color: C.orange, bar: 7.711, top: 2.316, sw: 7.049, lab: 7.065, lw: 0.623 }
  ].forEach(function (d) {
    fill(s, 'rect', d.bar, d.top, 0.503, 3.641 - d.top, d.color);
    fill(s, 'rect', d.sw, 3.825, 0.059, 0.061, d.color);
    tx(s, d.name, { x: d.lab, y: 3.747, w: d.lw, h: 0.21, fontSize: 8, color: C.axis, lineSpacingMultiple: 1.6 });
  });
}

/* --- 21. Contact --------------------------------------------------------- */
function slide21() {
  const s = pptx.addSlide();
  title(s, 'Contact Us', { w: 4.059, h: 0.455, lineSpacingMultiple: 1.0 });
  kicker(s, 0.347, 2.471);
  tx(s, L.oneLine, { x: 0.347, y: 3.2, w: 2.51, h: 0.911, fontSize: 8, lineSpacingMultiple: 1.3 });
  emailLine(s, 4.474);
  photo(s, 3.458, 1.37, 3.612, 4.255);

  [
    { y: 2.09, head: 'Phone', value: '5555 5555 5555' },
    { y: 2.925, head: 'Mail', value: 'lore@example.com' },
    { y: 3.616, head: 'Web', value: 'www.example.com' },
    { y: 4.451, head: 'Social Net', value: '@twitter.com' }
  ].forEach(function (row) {
    tx(s, [
      { text: row.head, options: { fontFace: F.semi, fontSize: 9, bold: true, lineSpacingMultiple: 1.333, breakLine: true } },
      { text: row.value, options: { fontFace: F.body, fontSize: 10, lineSpacingMultiple: 1.231 } }
    ], { x: 7.672, y: row.y, w: 1.575, h: 0.412 });
  });
}

[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
 slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
 slide21].forEach(function (build) { build(); });

pptx.writeFile({ fileName: path.join(__dirname, '0c7af98e-b230-42db-bb35-d09b9ca56787_grok_final.pptx') })
  .then(function (f) { console.log('wrote ' + f); })
  .catch(function (e) { console.error(e); process.exit(1); });
