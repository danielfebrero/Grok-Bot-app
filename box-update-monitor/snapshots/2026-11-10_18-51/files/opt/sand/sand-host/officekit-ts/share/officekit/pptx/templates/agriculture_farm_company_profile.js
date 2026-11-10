/*
 * "Agriculture" — 30-slide deck rebuilt with pptxgenjs.
 *
 *   node 18711064-428a-4765-ba7f-57993ed1fee1_grok_final.js
 *     -> 18711064-428a-4765-ba7f-57993ed1fee1_grok_final.pptx (next to this file)
 *
 * Everything is plain pptxgenjs: the geometry below mirrors the source deck
 * (13.333 x 7.5 in, EMU offsets converted to inches). Raster photographs and
 * icon PNGs from the original are NOT embedded; they are re-drawn as native
 * vector pictograms or as flat "[image]" placeholder rectangles occupying the
 * same box. Empty picture placeholders in the source render as blank white in
 * PowerPoint, so they are intentionally omitted here too.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const C = {
  dark:  '385723', // accent6 lumMod 50%  – dominant deep green
  mid:   '548235', // accent6 lumMod 75%  – leaf green
  pale:  'C5E0B4', // accent6 lum 40/60   – chart tint
  white: 'FFFFFF',
  offWhite: 'F2F2F2',
  black: '000000',
  ink:   '0D0D0D',
  ink2:  '262626',
  ink3:  '252726',
  gray:  '808080',
  gray2: '7F7F7F',
  gray3: 'BFBFBF',
  grid:  'D9D9D9',
  axis:  '595959',
  photo: '5A5F43', // mean colour of the replaced landscape photographs
  bezel: '1F1F1F', // device mock-up bezel
  metal: 'C6C6C6', // device mock-up chrome
};

const F = { serif: 'Playfair Display', sans: 'Poppins', med: 'Poppins Medium',
            body: 'Lato', quote: 'Merriweather' };

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

const merge = (base, over) => Object.assign({}, base, over || {});

/* ----------------------------------------------------------------- helpers */
// Playfair headline, 28pt bold unless overridden.
function head(s, x, y, w, h, str, o) {
  s.addText(str, merge({ x, y, w, h, fontFace: F.serif, fontSize: 28, bold: true,
    color: C.black, valign: 'top' }, o));
}
// Poppins tracking-out label — the deck's ubiquitous 11pt caps caption.
function cap(s, x, y, w, h, str, o) {
  s.addText(str, merge({ x, y, w, h, fontFace: F.sans, fontSize: 11, bold: true,
    charSpacing: 1, color: C.ink, valign: 'top' }, o));
}
// Lato body copy, justified at 150% leading.
function para(s, x, y, w, h, str, o) {
  s.addText(str, merge({ x, y, w, h, fontFace: F.body, fontSize: 11, color: C.gray,
    align: 'justify', lineSpacingMultiple: 1.5, valign: 'top' }, o));
}
// Big Poppins figure ($79, 01, $38.000, 1927 ...).
function num(s, x, y, w, h, str, size, o) {
  s.addText(str, merge({ x, y, w, h, fontFace: F.sans, fontSize: size, bold: true,
    color: C.dark, valign: 'top', wrap: false }, o));
}
function box(s, x, y, w, h, o) {
  s.addShape('rect', merge({ x, y, w, h }, o));
}
function fillBox(s, x, y, w, h, color) {
  s.addShape('rect', { x, y, w, h, fill: { color } });
}
// Rectangle stored rotated 90 deg in the source: give it the on-screen w/h and
// the pre-rotation centre stays put.
function rotBox(s, cx, cy, w, h, color) {
  fillBox(s, cx + h / 2 - w / 2, cy + w / 2 - h / 2, w, h, color);
}
function hline(s, x, y, w, color, pt) {
  s.addShape('line', { x, y, w, h: 0, line: { color, width: pt } });
}
function vline(s, x, y, h, color, pt) {
  s.addShape('line', { x, y, w: 0, h, line: { color, width: pt } });
}
// The 0.577" "view next slide" arrow.
function arrow(s, x, y, color) {
  s.addShape('line', { x, y, w: 0.577, h: 0, line: { color, width: 1.5, endArrowType: 'triangle' } });
}
// Short accent rule under most headlines (1.131" @ 2.25pt by default).
function rule(s, x, y, color, w, pt) {
  hline(s, x, y, w === undefined ? 1.131 : w, color || C.mid, pt === undefined ? 2.25 : pt);
}
// 1.552 x 0.562 button chip + its centred caption. fill === null => outlined.
function chip(s, x, y, fill, label, labelColor) {
  s.addShape('rect', merge({ x, y, w: 1.552, h: 0.562 }, fill === null
    ? { fill: { type: 'none' }, line: { color: C.white, width: 1 } }
    : { fill: { color: fill } }));
  cap(s, x + 0.153, y + 0.138, 1.245, 0.286, label, { color: labelColor || C.white });
}
// Tiny 0.083" hollow square, repeated three times as a divider ornament.
function squares(s, x, y) {
  [0, 0.1625, 0.325].forEach(dy => box(s, x, y + dy, 0.083, 0.083,
    { fill: { type: 'none' }, line: { color: C.dark, width: 1 } }));
}
// Three stacked diamonds, the deck's other divider ornament.
function diamonds(s, x, y, sz, gap, fill) {
  for (let i = 0; i < 3; i++) {
    s.addShape('diamond', { x, y: y + i * gap, w: sz, h: sz,
      fill: fill ? { color: fill } : { type: 'none' }, line: { color: C.dark, width: 0.75 } });
  }
}
// Dotted texture panel from the master layouts: a 6 x 6 lattice of 0.066"
// dots on an uneven row pitch, drawn at `scale` (1 = full size).
const DOT_COLS = [0, 0.222, 0.444, 0.666, 0.889, 1.111];
const DOT_ROWS = [0, 0.263, 0.542, 0.821, 1.100, 1.364];
function dotGrid(s, x, y, scale) {
  DOT_COLS.forEach(dx => DOT_ROWS.forEach(dy =>
    s.addShape('ellipse', { x: x + dx * scale, y: y + dy * scale,
      w: 0.066 * scale, h: 0.066 * scale, fill: { color: 'F0F0F0' } })));
}
/* Flat stand-in for a replaced photograph. */
function imagePlaceholder(s, x, y, w, h, color, labelColor) {
  fillBox(s, x, y, w, h, color);
  s.addText('[image]', { x, y: y + h / 2 - 0.16, w, h: 0.32, align: 'center', valign: 'middle',
    fontFace: F.sans, fontSize: 10, charSpacing: 1, color: labelColor });
}

/* Stand-ins for the replaced device mock-ups. The source PNGs are transparent
   product renders whose screens let the slide background show through, so
   these draw only the bezel / chin / stand outline with native shapes. */
function phonePlaceholder(s, x, y, w, h) {
  s.addShape('roundRect', { x, y, w, h, rectRadius: 0.18,
    fill: { type: 'none' }, line: { color: C.bezel, width: w * 0.030 * 72 } });
  s.addShape('roundRect', { x: x + 0.32 * w, y: y + 0.028 * h, w: 0.36 * w, h: 0.030 * h,
    rectRadius: 0.02, fill: { color: C.bezel } });
  s.addText('[image]', { x, y: y + h / 2 - 0.16, w, h: 0.32, align: 'center', valign: 'middle',
    fontFace: F.sans, fontSize: 10, charSpacing: 1, color: C.gray });
}
// All-in-one desktop: dark bezel frame, silver chin, neck and foot.
function monitorPlaceholder(s, x, y, w, h) {
  const fx = (a, b, cw, ch, color) => fillBox(s, x + a * w, y + b * h, cw * w, ch * h, color);
  box(s, x + 0.019 * w, y + 0.145 * h, 0.960 * w, 0.575 * h,
    { fill: { type: 'none' }, line: { color: C.bezel, width: w * 0.024 * 72 } });
  fx(0.019, 0.715, 0.960, 0.085, C.metal);
  fx(0.395, 0.800, 0.210, 0.120, C.metal);
  fx(0.335, 0.920, 0.330, 0.040, C.metal);
  s.addText('[image]', { x, y: y + 0.40 * h, w, h: 0.32, align: 'center', valign: 'middle',
    fontFace: F.sans, fontSize: 10, charSpacing: 1, color: C.gray });
}

/* ------------------------------------------------- native icon pictograms */
/* Each pictogram is drawn in a unit box scaled by `sz` so the same routine
   serves both the 0.62" logo mark and the 1.0" feature icons. */

// Barn on a hillside with a sun — the deck's logo mark.
function iconBarn(s, x, y, sz, color) {
  const f = { color }, u = v => v * sz;
  s.addShape('sun',       { x: x + u(0.02), y: y + u(0.02), w: u(0.20), h: u(0.20), fill: f });
  s.addShape('triangle',  { x: x + u(0.40), y: y + u(0.14), w: u(0.44), h: u(0.17), fill: f });
  s.addShape('rect',      { x: x + u(0.45), y: y + u(0.28), w: u(0.34), h: u(0.32), fill: f });
  s.addShape('rect',      { x: x + u(0.60), y: y + u(0.40), w: u(0.05), h: u(0.20), fill: { color: C.white } });
  s.addShape('rtTriangle',{ x: x + u(0.08), y: y + u(0.42), w: u(0.42), h: u(0.22), fill: f, flipH: true });
  s.addShape('rect',      { x: x + u(0.48), y: y + u(0.42), w: u(0.52), h: u(0.22), fill: f });
  s.addShape('rect',      { x: x + u(0.02), y: y + u(0.66), w: u(0.96), h: u(0.11), fill: f });
}
// Grain silo: domed cylinder with rungs.
function iconSilo(s, x, y, sz, color) {
  const f = { color }, u = v => v * sz;
  s.addShape('roundRect', { x: x + u(0.26), y: y + u(0.06), w: u(0.48), h: u(0.36), fill: f,
    rectRadius: u(0.20) });
  s.addShape('rect',      { x: x + u(0.26), y: y + u(0.20), w: u(0.48), h: u(0.74), fill: f });
  s.addShape('rect',      { x: x + u(0.50), y: y + u(0.32), w: u(0.04), h: u(0.62),
    fill: { color: C.white } });
  for (let i = 0; i < 4; i++) {
    s.addShape('rect', { x: x + u(0.31), y: y + u(0.38 + i * 0.14), w: u(0.15), h: u(0.05),
      fill: { color: C.white } });
  }
}
// Farmer in a wide-brim hat, holding a pitchfork.
function iconFarmer(s, x, y, sz, color) {
  const f = { color }, u = v => v * sz;
  s.addShape('trapezoid', { x: x + u(0.36), y: y + u(0.08), w: u(0.30), h: u(0.15), fill: f });
  s.addShape('rect',      { x: x + u(0.24), y: y + u(0.21), w: u(0.54), h: u(0.06), fill: f });
  s.addShape('ellipse',   { x: x + u(0.38), y: y + u(0.27), w: u(0.26), h: u(0.22), fill: f });
  s.addShape('roundRect', { x: x + u(0.26), y: y + u(0.52), w: u(0.50), h: u(0.42), fill: f,
    rectRadius: u(0.10) });
  s.addShape('rect',      { x: x + u(0.07), y: y + u(0.18), w: u(0.045), h: u(0.76), fill: f });
  s.addShape('rect',      { x: x + u(0.00), y: y + u(0.16), w: u(0.19), h: u(0.045), fill: f });
  [0.005, 0.072, 0.14].forEach(dx =>
    s.addShape('rect', { x: x + u(dx), y: y + u(0.05), w: u(0.035), h: u(0.13), fill: f }));
}
// Seed sack with a leaf motif.
function iconSeedBag(s, x, y, sz, color) {
  const f = { color }, u = v => v * sz;
  s.addShape('rect',      { x: x + u(0.28), y: y + u(0.08), w: u(0.44), h: u(0.05), fill: f });
  s.addShape('roundRect', { x: x + u(0.24), y: y + u(0.18), w: u(0.52), h: u(0.74), fill: f,
    rectRadius: u(0.12) });
  s.addShape('teardrop',  { x: x + u(0.36), y: y + u(0.36), w: u(0.28), h: u(0.28),
    fill: { color: C.white }, rotate: 225 });
}
/* Oversized translucent succulent leaves behind the dark title slides.
   The source art is one huge SVG bled off two corners; these teardrops
   approximate the four visible blades. */
function leafPattern(s) {
  const leaves = [
    { x: -1.20, y: -2.60, w: 5.60, h: 4.60, rotate: 45 },  // top-left mass
    { x: -2.10, y: 2.00, w: 3.20, h: 4.80, rotate: -20 },  // blade running down the left
    { x: 11.40, y: 0.20, w: 2.40, h: 4.40, rotate: -20 },  // top-right sliver
    { x: 9.50, y: 2.20, w: 4.60, h: 6.40, rotate: 22 },    // bottom-right mass
  ];
  leaves.forEach(l => s.addShape('teardrop', merge(l, { fill: { color: C.white, transparency: 95 } })));
}

/* ------------------------------------------------------- shared slide parts */
// Layout 1 chrome: deep-green field, foliage and a hairline frame.
function titleSlide(s, title) {
  s.background = { color: C.dark };
  leafPattern(s);
  box(s, 0.387, 0.387, 12.56, 6.726,
    { fill: { type: 'none' }, line: { color: C.white, width: 1, transparency: 75 } });
  iconBarn(s, 6.357, 2.543, 0.62, C.white);
  s.addText(title, { x: 2.127, y: 3.107, w: 9.079, h: 1.313, align: 'center',
    fontFace: F.serif, fontSize: 72, color: C.white, valign: 'top' });
  cap(s, 5.919, 4.338, 1.496, 0.269, 'BY SUB1 STUDIO',
    { fontSize: 10, color: C.offWhite, wrap: false });
}
// Slides 21-24 share one SWOT frame; only the letter and the word change.
function swotSlide(s, letter, letterX, letterW, word) {
  s.addText(letter, { x: letterX, y: 2.69, w: letterW, h: 2.121, fontFace: F.serif,
    fontSize: 120, bold: true, color: C.white, valign: 'top', wrap: false });
  box(s, 6.554, 0.401, 5.985, 6.698, { fill: { type: 'none' }, line: { color: C.dark, width: 1 } });
  diamonds(s, 6.490, 3.554, 0.127, 0.132, C.white);
  diamonds(s, 12.476, 3.554, 0.127, 0.132, C.white);
  cap(s, 7.653, 1.231, 3.788, 0.286, 'OUR SWOT',
    { align: 'center', bold: false, fontFace: F.med, color: C.dark });
  s.addText([
    { text: word, options: { color: C.dark } },
    { text: ' Analysis Design Slide', options: { color: C.ink2 } },
  ], { x: 7.653, y: 1.631, w: 3.788, h: 1.919, align: 'center', fontFace: F.serif,
    fontSize: 36, bold: true, valign: 'top' });
  hline(s, 9.056, 3.750, 0.982, C.dark, 1.5);
  cap(s, 7.853, 4.145, 3.448, 0.286, 'FARM COMPANY IN US', { align: 'center', color: C.dark });
  para(s, 7.853, 4.538, 3.448, 1.731,
    'PLACEHOLDER' +
    'PLACEHOLDER' +
    'laborum Lorem enimvas internno', { align: 'center' });
}

/* ------------------------------------------------------------ reused copy */
const T = {
  lorem: 'PLACEHOLDER' +
    'estas laborum Lorem enimvas internase nostrud consectur anime id est laborumer  in wepsum ' +
    'dolor sit in emase anime laborum ipsum dolor versacereser',
  loremShort: 'PLACEHOLDER' +
    'PLACEHOLDER',
  loremTiny: 'PLACEHOLDER',
  loremWide: 'PLACEHOLDER' +
    'PLACEHOLDER' +
    'sthe mawna aliqua utaradenisi',
  duis: 'PLACEHOLDER' +
    'PLACEHOLDER' +
    'nostrud consectetur anime id est laborum ',
  volu: 'PLACEHOLDER' +
    'dolore magna aliqua uta enimvas minimasa',
  card: 'internase nostrud consectur anime id est laborumer  in wepsum dolor sit in emase anime ' +
    'laborum ipsum dolor versacereser',
  europes: 'PLACEHOLDER' +
    'PLACEHOLDER' +
    'PLACEHOLDER' +
    'sed do eiusmod tempor incididunt aut laborevet',
  tail: 'PLACEHOLDER',
  team: 'Voluptate beliton essesa in cillum ipsum dolor  ronupan mawna aliqua uta consectetur ' +
    'PLACEHOLDER',
  price: 'PLACEHOLDER' +
    'PLACEHOLDER' +
    'ullamco laborisoni',
  loreme: 'PLACEHOLDER' +
    'officia deserut mollit anim id est laborum Loreme',
  member: 'Duis aute irure dolor in reprehenderit in',
  step: 'PLACEHOLDER',
  contact: 'PLACEHOLDER',
  dolorans: 'Voluptate beliton esesa in cillume ipsum dolorans',
};

/* ------------------------------------------------------------------ slides */
const build = [];

// 1 — cover
build.push(s => titleSlide(s, 'AGRICULTURE'));

// 2 — intro, two colour cards along the bottom
build.push(s => {
  fillBox(s, 0, 4.694, 3.833, 2.806, C.mid);
  fillBox(s, 3.833, 4.694, 3.976, 2.806, C.dark);
  dotGrid(s, 5.566, 2.217, 0.5);
  head(s, 0.472, 0.604, 4.387, 1.043, 'Introduce Natural & Healthy Farm');
  rule(s, 0.595, 1.821);
  cap(s, 0.5, 2.085, 2.903, 0.286, 'FARM COMPANY IN CA');
  para(s, 0.5, 2.419, 2.903, 1.731, T.lorem);
  s.addShape('ellipse', { x: 4.425, y: 2.227, w: 0.447, h: 0.447, fill: { color: C.mid } });
  s.addShape('ellipse', { x: 4.572, y: 2.374, w: 0.154, h: 0.154, fill: { color: C.white } });
  cap(s, 4.375, 3.864, 2.903, 0.286, 'VIEW NEXT SLIDE');
  arrow(s, 6.249, 4.0, C.mid);
  cap(s, 0.5, 5.384, 2.903, 0.286, 'TEXT TITTLE HERE', { color: C.white });
  para(s, 0.5, 5.721, 2.556, 1.175, T.card, { color: C.offWhite });
  cap(s, 4.375, 5.384, 2.903, 0.286, 'TEXT TITTLE HERE', { color: C.white });
  para(s, 4.375, 5.721, 2.556, 1.175, T.card, { color: C.offWhite });
});

// 3 — same intro mirrored; dark card bottom-right
build.push(s => {
  fillBox(s, 7.667, 4.694, 5.667, 2.806, C.dark);
  head(s, 8.333, 0.576, 4.387, 1.043, 'Introduce Natural & Healthy Farm');
  rule(s, 8.456, 1.793);
  cap(s, 8.361, 2.057, 2.903, 0.286, 'FARM COMPANY IN CA');
  para(s, 8.361, 2.391, 2.903, 1.731, T.lorem);
  vline(s, 12.507, 2.571, 1.466, C.dark, 0.75);
  diamonds(s, 12.348, 3.252, 0.106, 0.106, C.white);
  cap(s, 8.375, 5.115, 2.903, 0.286, 'TEXT TITTLE HERE', { color: C.white });
  para(s, 8.4, 5.452, 2.864, 0.897, 'internase nostrud consectur anime id est laborumer  in ' +
    'wepsum dolor sit in emase anime laborum ipsumei duri', { color: C.offWhite });
  cap(s, 8.428, 6.924, 2.903, 0.286, 'VIEW NEXT SLIDE', { color: C.white });
  arrow(s, 11.958, 7.05, C.white);
});

// 4 — agenda table
build.push(s => {
  fillBox(s, 5.472, 0, 7.861, 4.236, C.dark);
  head(s, 0.858, 0.698, 3.756, 1.717, 'Introduce Natural & Healthy Farm', { fontSize: 32 });
  rule(s, 1.014, 2.670);
  cap(s, 0.918, 2.970, 2.903, 0.286, 'FARM COMPANY IN CA');
  para(s, 0.918, 3.305, 2.903, 1.731, T.lorem);
  cap(s, 6.234, 0.807, 2.518, 0.471, 'INTRODUCE NATURAL HEALTHY FARM', { color: C.white });
  para(s, 8.641, 0.740, 3.930, 1.453,
    T.duis + 'Lorem ipsum dolor sit in emase versace in the flre dolor', { color: C.offWhite });
  cap(s, 6.234, 3.391, 2.903, 0.286, 'VIEW NEXT SLIDE', { color: C.white });
  arrow(s, 11.792, 3.517, C.white);
  [6.182, 6.484, 6.786].forEach(y => hline(s, 1.012, y, 2.928, C.gray, 1));
  [['NAME', 0.914, 5.907, 0.984], ['ABOUT US', 0.914, 6.230, 1.063],
   ['INTRODUCTION', 0.914, 6.547, 1.470], ['NO', 2.623, 5.849, 0.459], ['NO', 3.480, 5.849, 0.459],
   ['01', 2.623, 6.230, 0.459], ['02', 3.517, 6.230, 0.459],
   ['03', 2.620, 6.517, 0.459], ['04', 3.514, 6.517, 0.459],
  ].forEach(([t, x, y, w]) => cap(s, x, y, w, 0.255, t, { fontSize: 10, bold: false }));
});

// 5 — mission / vision under the (empty) hero band
build.push(s => {
  fillBox(s, 0, 4.063, 5.175, 3.437, C.dark);
  num(s, 2.615, 4.669, 1.392, 0.774, '1927', 40,
    { fontFace: F.med, bold: false, color: C.mid, transparency: 85 });
  cap(s, 0.852, 4.905, 3.690, 0.286, 'AGROCULTURE MISSION', { color: C.white });
  s.addText('“Our mission in life is not merely to survive, but to thrive and to do so with some ' +
    'passion, some work, some humor, and some style.”', { x: 0.852, y: 5.284, w: 3.690, h: 1.576,
    fontFace: F.quote, fontSize: 15, italic: true, color: C.offWhite,
    lineSpacingMultiple: 1.5, valign: 'top' });
  cap(s, 5.938, 4.905, 3.930, 0.286, 'AGROCULTURE VISION');
  para(s, 5.938, 5.239, 3.930, 1.731, T.duis +
    'PLACEHOLDER');
});

// 6 — product intro with two tab chips top-left
build.push(s => {
  fillBox(s, 5.766, 5.986, 7.568, 1.514, C.dark);
  fillBox(s, 2.044, 1.306, 1.863, 0.778, C.mid);
  fillBox(s, 3.907, 1.306, 1.863, 0.778, C.dark);
  cap(s, 2.367, 1.577, 1.245, 0.286, 'TITTLE HERE', { color: C.white });
  cap(s, 4.251, 1.577, 1.245, 0.286, 'TITTLE HERE', { color: C.white });
  iconBarn(s, 11.905, 0.315, 0.62, C.dark);
  cap(s, 6.667, 1.579, 2.903, 0.286, 'AGRICULTURE PRESENTATION');
  hline(s, 6.736, 2.083, 5.770, C.dark, 0.75);
  head(s, 6.667, 2.437, 4.387, 1.043, 'Product Natural & Healthy Farm');
  rule(s, 6.761, 3.638);
  cap(s, 6.667, 3.902, 2.903, 0.286, 'FARM COMPANY IN CA');
  para(s, 6.667, 4.236, 2.903, 1.175, T.loremShort);
  vline(s, 12.513, 2.669, 0.872, C.dark, 1);
  squares(s, 12.472, 3.789);
  vline(s, 12.518, 4.469, 0.872, C.dark, 1);
  cap(s, 6.761, 6.611, 2.903, 0.286, 'TEXT TITTLE HERE', { color: C.white });
  cap(s, 8.991, 6.606, 2.903, 0.286, 'VIEW NEXT SLIDE', { color: C.white });
  arrow(s, 11.894, 6.751, C.white);
});

// 7 — photo card behind a dark scrim, numbered lead-in
build.push(s => {
  fillBox(s, 0, 0, 6.667, 1.514, C.dark);
  cap(s, 0.597, 0.656, 2.903, 0.286, '01. TEXT TITTLE HERE', { color: C.white });
  cap(s, 3.365, 0.656, 2.903, 0.286, '02. TEXT TITTLE HERE', { color: C.white });
  imagePlaceholder(s, 0.597, 2.118, 2.595, 2.438, C.photo, C.offWhite);
  box(s, 0.597, 2.118, 2.595, 2.438, { fill: { color: C.black, transparency: 75 } });
  cap(s, 3.379, 3.101, 2.903, 0.471, [
    { text: 'INTRODUCING YOUR ', options: { breakLine: true } }, { text: 'TITLE HERE' }]);
  cap(s, 5.949, 3.138, 0.507, 0.404, '01', { fontSize: 18 });
  hline(s, 3.483, 3.673, 2.843, C.dark, 0.75);
  para(s, 3.379, 3.936, 2.947, 0.620,
    'PLACEHOLDER');
  head(s, 0.443, 4.915, 4.387, 1.043, 'Product Natural & Healthy Farm');
  para(s, 0.472, 6.128, 3.951, 0.897, T.loremShort);
  cap(s, 4.732, 6.725, 1.454, 0.286, 'NEXT');
  arrow(s, 5.660, 6.858, C.mid);
});

// 8 — about, wide dark band
build.push(s => {
  fillBox(s, 0, 2.41, 13.333, 4.236, C.dark);
  fillBox(s, 0, 1.858, 3.873, 0.552, C.mid);
  iconBarn(s, 0.517, 0.439, 0.62, C.dark);
  cap(s, 4.540, 1.978, 2.903, 0.286, 'VIEW NEXT SLIDE');
  arrow(s, 7.156, 2.086, C.mid);
  cap(s, 0.486, 3.292, 2.903, 0.286, 'ABOUT OUR VISION', { color: C.white });
  para(s, 0.486, 3.669, 3.388, 0.897, T.volu, { color: C.offWhite });
  cap(s, 4.541, 3.292, 2.903, 0.286, 'ABOUT OUR MISION', { color: C.white });
  para(s, 4.541, 3.669, 3.388, 0.897, T.volu, { color: C.offWhite });
  head(s, 0.486, 5.224, 4.387, 0.572, 'About Agriculture', { color: C.white });
  chip(s, 4.628, 5.225, null, 'TITTLE HERE');
  chip(s, 6.377, 5.225, C.white, 'TITTLE HERE', C.dark);
});

// 9 — three tabs over the photo band, two-column copy
build.push(s => {
  fillBox(s, 0, 2.167, 6.097, 1.194, C.dark);
  fillBox(s, 6.097, 6.306, 5.736, 1.194, C.dark);
  fillBox(s, 11.833, 6.306, 1.5, 1.194, C.mid);
  iconBarn(s, 11.905, 0.315, 0.62, C.dark);
  cap(s, 0.768, 1.239, 2.903, 0.286, 'FARM COMPANY IN CA');
  para(s, 0.768, 1.574, 2.903, 0.342, 'Voluptate beliton essesa in cillum ipsum');
  [0.761, 2.426, 4.091].forEach(x => cap(s, x, 2.636, 1.245, 0.286, 'TITTLE HERE', { color: C.white }));
  cap(s, 6.666, 1.630, 2.903, 0.286, 'TITTLE HERE');
  cap(s, 8.843, 1.630, 2.903, 0.286, 'TITTLE HERE');
  hline(s, 6.736, 2.168, 5.770, C.dark, 0.75);
  head(s, 6.667, 2.471, 4.387, 1.043, 'Product Natural & Healthy Farm');
  para(s, 6.667, 3.629, 5.150, 0.897, T.loremWide);
  para(s, 6.667, 4.676, 5.150, 0.897, T.loremWide);
  vline(s, 12.560, 3.750, 1.750, C.dark, 0.75);
  diamonds(s, 12.401, 4.573, 0.106, 0.106, C.white);
  cap(s, 6.736, 6.774, 2.903, 0.286, 'TEXT TITTLE HERE', { color: C.white });
  cap(s, 9.603, 6.759, 2.903, 0.286, 'VIEW NEXT SLIDE', { color: C.white });
  arrow(s, 12.309, 6.917, C.white);
});

// 10 — two packages beside vertical tabs
build.push(s => {
  rotBox(s, 4.869, 2.202, 1.318, 3.111, C.dark);
  rotBox(s, 4.869, 5.313, 1.318, 3.111, C.mid);
  [2.714, 5.806].forEach(y =>
    cap(s, 5.563, y, 1.725, 0.286, 'TEXT TITTLE HERE', { color: C.white, rotate: 270 }));
  iconBarn(s, 11.905, 0.315, 0.62, C.dark);
  head(s, 7.464, 1.988, 4.387, 1.043, 'Product Natural & Healthy Farm');
  rule(s, 7.559, 3.278);
  [['01', 7.464, 3.806, '01. PACKAGE ONE', 4.690], ['02', 10.336, 3.832, '02. PACKAGE TWO', 4.693]]
    .forEach(([n, x, y, label, ly]) => {
      num(s, x, y, 0.958, 0.774, n, 40, { charSpacing: 1, color: C.mid, transparency: 84, wrap: true });
      cap(s, x, ly, 2.744, 0.286, label, { charSpacing: 0, color: C.ink2 });
      para(s, x, 5.080, 2.536, 1.731, T.price);
    });
});

// 11 — four icon features around two vertical bars
build.push(s => {
  [[2.528, 1.583, C.dark], [2.528, 5.333, C.mid], [7.056, 1.583, C.dark], [7.056, 5.333, C.mid]]
    .forEach(([cx, cy, col]) => rotBox(s, cx, cy, 0.583, 3.75, col));
  dotGrid(s, 3.689, 3.035, 1);
  dotGrid(s, 8.467, 3.035, 1);
  [[iconSilo,    0.649, 1.145, 1.629, 0.857],
   [iconBarn,    0.649, 4.049, 1.556, 1.000],
   [iconFarmer,  9.870, 1.145, 10.813, 0.928],
   [iconSeedBag, 9.870, 4.139, 10.777, 1.000],
  ].forEach(([icon, x, y, ix, isz]) => {
    icon(s, ix, y, isz, C.dark);
    cap(s, x, y + 1.012, 2.815, 0.286, 'INTERACTION DESIGN', { align: 'center', color: C.black });
    para(s, x, y + 1.331, 2.815, 0.897, T.loremTiny, { align: 'center' });
  });
});

// 12 — pricing tiers beneath a dark hero panel
build.push(s => {
  rotBox(s, 7.840, -1.174, 6.667, 4.319, C.dark);
  cap(s, 8.360, 0.793, 2.518, 0.471, 'INTRODUCE NATURAL HEALTHY FARM', { color: C.white });
  para(s, 8.360, 1.391, 3.388, 0.897, T.volu, { color: C.offWhite });
  chip(s, 8.447, 2.947, null, 'TITTLE HERE');
  arrow(s, 11.079, 3.238, C.white);
  fillBox(s, 6.167, 1.649, 1.0, 1.0, C.white);
  iconBarn(s, 6.371, 1.854, 0.591, C.dark);
  [['$79', 1.586, 1.217], ['$89', 5.259, 1.280], ['$99', 8.933, 1.263]].forEach(([p, x, w]) => {
    num(s, x, 5.138, w, 0.774, p, 40);
    para(s, x, 5.809, 2.815, 0.897, T.loremTiny, { align: 'left' });
  });
  vline(s, 4.592, 5.267, 1.440, C.mid, 1);
  vline(s, 8.351, 5.267, 1.440, C.mid, 1);
});

// 13 — three numbered steps
build.push(s => {
  fillBox(s, 1.167, 2.944, 5.363, 0.349, C.dark);
  rotBox(s, 4.731, 4.743, 0.349, 3.249, C.mid);
  fillBox(s, 6.529, 0, 5.304, 1.194, C.dark);
  fillBox(s, 11.833, 0, 1.5, 1.194, C.mid);
  head(s, 1.110, 1.481, 4.387, 1.043, 'Product Natural & Healthy Farm');
  cap(s, 7.478, 1.971, 2.903, 0.286, 'VIEW NEXT SLIDE', { color: C.dark });
  arrow(s, 10.562, 2.080, C.mid);
  [['1', 7.577, 2.958, 7.621, 3.010, 7.711, 3.057, 0.244, 0.318, 8.341, 3.109, 8.348, 2.881],
   ['2', 7.621, 4.257, 7.664, 4.309, 7.754, 4.356, 0.360, 0.404, 8.385, 4.408, 8.392, 4.179],
   ['3', 7.618, 5.556, 7.661, 5.608, 7.751, 5.655, 0.368, 0.404, 8.382, 5.707, 8.389, 5.478],
  ].forEach(([n, bx, by, fx, fy, tx, ty, tw, th, px, py, cx, cy]) => {
    fillBox(s, bx, by, 0.487, 0.487, C.mid);
    box(s, fx, fy, 0.487, 0.487, { fill: { color: C.dark }, line: { color: C.white, width: 1 } });
    num(s, tx, ty, tw, th, n, 18, { charSpacing: 1, color: C.white });
    cap(s, cx, cy, 2.708, 0.286, 'AGRICULTURE YOUR TEXT', { color: C.ink3 });
    para(s, px, py, 2.999, 0.897, [
      { text: T.step, options: { breakLine: true } },
      { text: 'oratau ad don esse murik  dolore sepurani' }], { color: C.gray2 });
  });
  vline(s, 12.527, 2.958, 1.337, C.dark, 1);
  squares(s, 12.486, 4.544);
  vline(s, 12.532, 5.224, 1.309, C.dark, 1);
});

// 14 — team, three members
build.push(s => {
  fillBox(s, 0, 7.151, 13.333, 0.349, C.dark);
  head(s, 1.366, 1.055, 4.387, 1.043, [
    { text: 'Our Agriculture ', options: { breakLine: true } }, { text: 'Team' }]);
  para(s, 5.175, 1.128, 3.896, 0.897, T.team, { align: 'left' });
  chip(s, 10.250, 1.340, C.dark, 'TITTLE HERE');
  [[3.507, 2.523, 3.507, 3.692, 1.423, 1.440, 'CAESAR MALDIVO', 1.580],
   [7.297, 2.501, 7.293, 7.494, 5.175, 5.193, 'SRI VERONICA DENA', 1.745],
   [11.099, 2.501, 11.094, 11.299, 9.016, 9.033, 'JOHNSON RAMSEY', 1.643],
  ].forEach(([bx, by, sx, ix, tx, nx, name, nw]) => {
    fillBox(s, bx, by, 0.755, 0.787, C.mid);
    fillBox(s, sx, by, 0.118, 0.787, C.dark);
    iconFarmer(s, ix, 2.665, 0.503, C.white);
    cap(s, nx, 5.694, nw, 0.286, name, { charSpacing: 0, color: C.dark, wrap: false });
    para(s, tx, 5.967, 2.570, 0.620, T.member, { align: 'left', color: C.gray2 });
  });
});

// 15 — team variant with an index list and a progress bar
build.push(s => {
  fillBox(s, 0, 5.963, 1.794, 1.537, C.dark);
  fillBox(s, 1.794, 5.963, 1.794, 1.537, C.mid);
  head(s, 0.722, 1.055, 3.541, 1.043, [
    { text: 'Our Agriculture ', options: { breakLine: true } }, { text: 'Team' }]);
  para(s, 4.330, 1.128, 3.896, 0.897, T.team, { align: 'left' });
  cap(s, 9.763, 1.734, 2.903, 0.286, 'VIEW NEXT SLIDE', { color: C.dark });
  arrow(s, 11.952, 1.842, C.mid);
  [['01. AGRICULTURE YOUR TEXT', 0.831, 2.605, 2.708, C.gray3],
   ['02. AGRICULTURE', 0.834, 3.237, 2.708, C.gray3],
   ['03. AGRICULTURE TEXT', 0.831, 3.869, 2.708, C.gray3],
   ['04. AGRICULTURE YOUR TEXT', 0.831, 4.501, 2.915, C.dark],
   ['05. YOUR TEXT', 0.831, 5.129, 2.708, C.gray3],
  ].forEach(([t, x, y, w, col]) => cap(s, x, y, w, 0.286, t, { color: col }));
  [[6.328, 2.647, 6.513, 2.789, 4.274, 4.291, 6.135, 5.861],
   [9.090, 2.642, 9.275, 2.784, 7.083, 7.101, 6.135, 5.861],
   [11.795, 2.631, 11.979, 2.773, 9.845, 9.863, 6.131, 5.858],
  ].forEach(([bx, by, ix, iy, tx, nx, ty, ny]) => {
    fillBox(s, bx, by, 0.755, 0.787, C.mid);
    fillBox(s, bx, by, 0.118, 0.787, C.dark);
    iconFarmer(s, ix, iy, 0.503, C.white);
    cap(s, nx, ny, 1.580, 0.286, 'CAESAR MALDIVO', { charSpacing: 0, color: C.dark, wrap: false });
    para(s, tx, ty, 2.570, 0.620, T.member, { align: 'left', color: C.gray2 });
  });
  fillBox(s, 4.372, 5.197, 4.083, 0.207, C.dark);
  fillBox(s, 8.455, 5.197, 4.101, 0.207, C.mid);
});

// 16 — break slide
build.push(s => titleSlide(s, 'BREAK SLIDE'));

// 17 — long copy plus three chips
build.push(s => {
  fillBox(s, 5.766, 6.528, 7.568, 0.972, C.dark);
  head(s, 6.607, 0.945, 4.387, 1.043, 'Product Natural & Healthy Farm');
  para(s, 6.607, 2.336, 5.460, 1.453, T.europes + T.tail);
  para(s, 6.607, 3.930, 5.460, 0.620, 'lorem ipsum dolor siabes dolore magna aliqua uta enimvas ' +
    'PLACEHOLDER');
  cap(s, 6.649, 4.898, 2.903, 0.286, 'VIEW NEXT SLIDE', { color: C.dark });
  arrow(s, 8.479, 5.020, C.mid);
  chip(s, 6.732, 5.544, C.mid, 'TITTLE HERE');
  chip(s, 8.481, 5.525, C.dark, 'TITTLE HERE');
  chip(s, 10.229, 5.525, C.mid, 'TITTLE HERE');
  cap(s, 6.593, 6.893, 2.903, 0.286, 'TEXT YOUR TITTLE HERE', { color: C.white });
  cap(s, 9.669, 6.878, 2.903, 0.286, 'TEXT YOUR TITTLE HERE', { color: C.white });
});

// 18 — introduction / specialization
build.push(s => {
  fillBox(s, 7.641, 0, 5.693, 3.437, C.dark);
  dotGrid(s, 5.944, 3.313, 1);
  rotBox(s, 11.049, 5.923, 1.652, 0.705, C.mid);
  head(s, 0.681, 1.108, 4.387, 1.043, 'Product Natural & Healthy Farm');
  rule(s, 0.809, 2.454);
  cap(s, 0.709, 2.878, 1.552, 0.286, 'INTRODUCTION');
  para(s, 0.709, 3.212, 4.611, 1.453, T.europes);
  cap(s, 0.709, 4.930, 2.152, 0.286, 'OUR SPECIALIZATION');
  para(s, 0.709, 5.217, 4.611, 1.175, 'Duis aute irure dolor in reprehenderit in voluptate velit ' +
    'PLACEHOLDER' +
    'PLACEHOLDER');
  cap(s, 10.829, 6.620, 1.265, 0.286, 'NEXT SLIDE', { color: C.white });
});

// 19 — two floating price cards
build.push(s => {
  rotBox(s, 6.541, 4.969, 4.093, 0.972, C.mid);
  rotBox(s, 10.714, 4.889, 4.260, 0.979, C.dark);
  [['$38.000', 1.118, 1.927, 1.596, 'PRICE VALUE 75%', 2.431],
   ['$85.000', 4.912, 4.837, 1.611, 'PRICE VALUE 80%', 5.341]].forEach(([p, x, y, w, lbl, ly]) => {
    num(s, x, y, w, 0.505, p, 24);
    cap(s, x, ly, 2.025, 0.286, lbl, { charSpacing: 0.7, color: C.ink2 });
  });
  head(s, 8.946, 1.731, 4.387, 1.043, 'Product Natural & Healthy Farm');
  rule(s, 9.073, 3.040);
  cap(s, 8.946, 3.581, 2.815, 0.286, 'LETS FIND YOUR DREAM HOME', { charSpacing: 0.7, color: C.ink2 });
  para(s, 8.948, 3.919, 3.750, 1.731, T.duis +
    'PLACEHOLDER');
});

// 20 — three-point list beside the photo column
build.push(s => {
  rotBox(s, 10.830, 5.005, 4.028, 0.979, C.dark);
  rotBox(s, 9.473, 5.383, 1.313, 0.979, C.mid);
  vline(s, 0.677, 4.709, 2.791, C.dark, 1);
  diamonds(s, 0.518, 4.499, 0.106, 0.106, C.white);
  head(s, 1.153, 1.464, 4.387, 1.043, 'Product Natural & Healthy Farm');
  para(s, 1.195, 2.820, 3.482, 2.008, [
    { text: T.loreme, options: { breakLine: true } },
    { text: 'PLACEHOLDER' +
        'europes vui officia deserut mollit anim id est laborum Loreme' }]);
  cap(s, 1.195, 4.692, 3.534, 0.286, 'YOUR TEXT AGROCULTURE', { color: C.black });
  para(s, 1.195, 5.000, 3.482, 0.897, T.loreme);
  fillBox(s, 1.310, 6.519, 1.806, 0.328, C.dark);
  cap(s, 1.518, 6.540, 1.391, 0.286, 'YOUR TITTLE', { color: C.white });
  hline(s, 10.228, 1.502, 1.831, C.dark, 1);
  [['01. YOUR TEXT HERE', 1.941, 2.260], ['02. YOUR TEXT HERE', 3.089, 3.408],
   ['03. YOUR TEXT HERE', 4.260, 4.578]].forEach(([t, ly, py]) => {
    cap(s, 10.078, ly, 3.534, 0.286, t, { color: C.black });
    para(s, 10.078, py, 2.556, 0.620, T.dolorans, { align: 'left' });
  });
  iconBarn(s, 9.669, 5.726, 0.591, C.white);
  cap(s, 9.624, 6.907, 2.903, 0.286, 'VIEW NEXT SLIDE', { color: C.white });
  arrow(s, 11.950, 7.024, C.white);
});

// 21-24 — SWOT letters
build.push(s => swotSlide(s, 'S', 2.264, 1.185, 'Strength'));
build.push(s => swotSlide(s, 'W', 1.942, 1.831, 'Weakness'));
build.push(s => swotSlide(s, 'O', 2.089, 1.536, 'Opportunity'));
build.push(s => swotSlide(s, 'T', 2.206, 1.326, 'Threat'));

// 25 — phone mock-ups
build.push(s => {
  fillBox(s, 6.667, 0, 6.667, 7.5, C.dark);
  phonePlaceholder(s, 9.291, -0.410, 3.463, 5.710);
  phonePlaceholder(s, 5.596, 2.266, 3.463, 5.710);
  head(s, 0.698, 1.272, 4.387, 1.043, 'Monitoring Your Farm With Gadget');
  para(s, 0.698, 2.663, 4.387, 2.008, T.europes + T.tail);
  cap(s, 0.739, 5.019, 2.903, 0.286, 'VIEW NEXT SLIDE', { color: C.dark });
  arrow(s, 2.612, 5.142, C.mid);
  chip(s, 0.823, 5.666, C.mid, 'TITTLE HERE');
  chip(s, 2.571, 5.647, C.dark, 'TITTLE HERE');
});

// 26 — desktop mock-up
build.push(s => {
  fillBox(s, 0, 0, 4.587, 7.5, C.dark);
  fillBox(s, 0, 0, 0.764, 7.5, C.mid);
  monitorPlaceholder(s, -0.569, 1.988, 5.837, 4.853);
  fillBox(s, 6.370, 6.306, 5.823, 1.194, C.dark);
  fillBox(s, 12.193, 6.306, 1.140, 1.194, C.mid);
  iconBarn(s, 11.609, 0.397, 0.62, C.dark);
  cap(s, 6.299, 1.630, 2.903, 0.286, 'TITTLE HERE');
  cap(s, 8.477, 1.630, 2.903, 0.286, 'TITTLE HERE');
  hline(s, 6.370, 2.168, 5.770, C.dark, 0.75);
  head(s, 6.300, 2.471, 4.387, 1.043, 'Product Natural & Healthy Farm');
  para(s, 6.300, 3.629, 5.150, 0.897, T.loremWide);
  para(s, 6.300, 4.676, 5.150, 0.897, T.loremWide);
  vline(s, 12.193, 3.750, 1.750, C.dark, 0.75);
  diamonds(s, 12.035, 4.573, 0.106, 0.106, C.white);
  cap(s, 7.249, 6.782, 1.245, 0.286, 'TITTLE HERE', { color: C.white });
  cap(s, 9.742, 6.782, 2.903, 0.286, 'VIEW NEXT SLIDE', { color: C.white });
  arrow(s, 12.499, 6.912, C.white);
});

// 27 — square mock-up with a stats pill
build.push(s => {
  fillBox(s, 0, 5.963, 1.794, 1.537, C.dark);
  fillBox(s, 1.794, 5.963, 2.925, 1.537, C.mid);
  monitorPlaceholder(s, 1.433, 0.640, 4.820, 4.820);
  s.addShape('roundRect', { x: 7.194, y: 5.968, w: 6.773, h: 0.816, rectRadius: 0.408,
    fill: { type: 'none' }, line: { color: C.dark, width: 3 } });
  head(s, 7.081, 1.427, 4.387, 1.043, 'Product Natural & Healthy Farm');
  rule(s, 7.175, 2.628);
  cap(s, 7.081, 2.892, 2.903, 0.286, 'FARM COMPANY IN CA');
  para(s, 7.081, 3.227, 2.903, 1.175, T.loremShort);
  cap(s, 7.081, 4.927, 2.903, 0.286, 'VIEW NEXT SLIDE', { color: C.dark });
  arrow(s, 9.306, 5.049, C.mid);
  vline(s, 12.696, 1.714, 0.872, C.dark, 1);
  squares(s, 12.655, 2.835);
  vline(s, 12.701, 3.515, 0.872, C.dark, 1);
  cap(s, 2.612, 6.604, 1.245, 0.286, 'TITTLE HERE', { color: C.white });
  cap(s, 8.060, 6.266, 1.245, 0.286, 'TITTLE HERE', { color: C.black });
  ['190K', '290K'].forEach((v, i) => s.addText(v, { x: 10.402 + i * 1.443, y: 6.177, w: 1.066,
    h: 0.375, align: 'right', fontFace: F.sans, fontSize: 12, bold: true,
    lineSpacingMultiple: 1.5, valign: 'top' }));
});

// 28 — clustered column chart
build.push((s, pptx) => {
  fillBox(s, 0.634, 3.887, 5.437, 2.784, C.dark);
  head(s, 1.187, 1.002, 4.387, 1.043, 'Our Agriculture Pricing Table');
  cap(s, 1.201, 2.470, 2.903, 0.286, 'TITLE HERE');
  cap(s, 4.864, 2.409, 0.507, 0.404, '01', { fontSize: 18 });
  hline(s, 1.287, 2.971, 3.990, C.dark, 0.75);
  cap(s, 1.229, 3.301, 2.903, 0.286, 'VIEW NEXT SLIDE', { color: C.dark });
  para(s, 1.259, 4.238, 4.178, 0.897, T.volu, { color: C.offWhite });
  chip(s, 1.346, 5.612, null, 'TITTLE HERE');
  arrow(s, 4.672, 5.903, C.white);
  const labels = ['Graphic 1', 'Graphic 2', 'Graphic 3', 'Graphic 4'];
  s.addChart(pptx.ChartType.bar, [
    { name: 'Series 1', labels, values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels, values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels, values: [2, 2, 3, 3] },
  ], {
    x: 6.667, y: 1.828, w: 6.256, h: 5.167,
    barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 150,
    chartColors: [C.dark, C.pale, C.mid],
    showLegend: false, showTitle: false, showValue: false,
    catAxisLabelFontFace: 'Arial', catAxisLabelFontSize: 10, catAxisLabelColor: C.axis,
    valAxisLabelFontFace: 'Arial', valAxisLabelFontSize: 10, valAxisLabelColor: C.axis,
    catAxisLineColor: C.grid, valAxisLineShow: false,
    valGridLine: { color: C.grid, style: 'solid', size: 1 },
    catGridLine: { style: 'none' },
    valAxisMinVal: 0, valAxisMaxVal: 5, valAxisMajorUnit: 0.5,
    border: { pt: 0, color: C.white },
  });
});

// 29 — contact
build.push(s => {
  fillBox(s, 0, 4.958, 3.319, 2.542, C.mid);
  fillBox(s, 3.319, 4.958, 3.347, 2.542, C.dark);
  head(s, 0.626, 0.654, 2.377, 0.572, 'Contact Us');
  rule(s, 0.766, 1.441);
  [[0.612, 1.833, 2.529], [0.612, 2.731, 2.529], [3.515, 1.833, 2.541], [3.515, 2.731, 2.529]]
    .forEach(([x, y, w]) => para(s, x, y, w, 0.620, T.contact));
  cap(s, 0.643, 4.064, 2.903, 0.286, 'VIEW NEXT SLIDE', { color: C.dark });
  arrow(s, 5.465, 4.244, C.mid);
  hline(s, 0.729, 4.564, 5.327, C.dark, 0.75);
  cap(s, 0.751, 6.157, 2.903, 0.286, 'TEXT TITTLE HERE', { color: C.white });
  cap(s, 4.134, 6.175, 2.903, 0.286, 'TEXT TITTLE HERE', { color: C.white });
});

// 30 — closing
build.push(s => titleSlide(s, 'THANK YOU'));

/* -------------------------------------------------------------------- main */
function main() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'WIDE';
  pptx.title = 'Agriculture';
  pptx.author = 'SUB1 STUDIO';

  build.forEach(fn => {
    const slide = pptx.addSlide();
    slide.background = { color: C.white };
    fn(slide, pptx);
  });

  return pptx.writeFile({
    fileName: path.join(__dirname, '18711064-428a-4765-ba7f-57993ed1fee1_grok_final.pptx'),
  });
}

main().then(f => console.log('wrote ' + f)).catch(e => { console.error(e); process.exit(1); });
