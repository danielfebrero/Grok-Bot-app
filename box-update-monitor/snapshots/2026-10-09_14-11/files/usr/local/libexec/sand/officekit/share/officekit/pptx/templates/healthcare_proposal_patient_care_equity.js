/**
 * "Healthcare Proposal Deck" - 20 slide deck rebuilt with pptxgenjs.
 *
 * Run:  node 0390f017-0580-4d09-97b4-aef049866ad4_grok_final.js
 * Out:  0390f017-0580-4d09-97b4-aef049866ad4_grok_final.pptx (next to this file)
 *
 * Photographs in the source deck are redrawn as flat placeholder shapes and the
 * source's tiny SVG line-icons are approximated with native outline shapes.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
// Theme "Kustom 294" plus the lumMod/lumOff variants the deck uses.
const C = {
  a1: '024D60', // accent1  deep teal
  a2: '2CACAD', // accent2  teal
  a3: '75E2E0', // accent3  light teal
  a4: 'D9F5F0', // accent4  mint
  a5: '0F969C', // accent5  strong teal
  a2d: '218182', // accent2 lumMod 75%
  a2dd: '165656', // accent2 lumMod 50%
  a2l: 'D0F3F3', // accent2 lumMod 20% / lumOff 80%
  a3d: '2ED3D0', // accent3 lumMod 75%
  a3l: 'E3F9F9', // accent3 lumMod 20% / lumOff 80%
  white: 'FFFFFF',
  black: '000000',
  body: '404040', // tx1 lumMod 75% / lumOff 25%
  grey: 'E8E8E8', // bg2
  ring: 'DCEAF7', // tx2 lumMod 10% / lumOff 90%
  ringLine: 'A6CAEC', // tx2 lumMod 25% / lumOff 75%
  hairline: 'F2F2F2', // bg1 lumMod 95%
  // Flat stand-ins for the gradients pptxgenjs cannot express.
  nightTop: '2E6A69', // full-bleed accent1 -> accent2 title background, top ...
  nightBottom: '215F70', // ... and bottom of the ramp
  btn: '279797', // accent2 -> accent2 lumMod 75% button
  card: '26989A', // team card
  cardTeal: '1BA0A2', // accent5 -> accent2 stat card
  crossPale: 'E8FAF8', // faintest decorative cross
  crossMint: 'C7F3EF', // mint decorative cross
  crossGlow: '3C868C', // large translucent cross on the dark slides
  crossDeep: '18727D',
  crossDeep2: '187882',
};

const MAJOR = 'Montserrat'; // +mj-lt
const MINOR = 'Roboto'; // +mn-lt

/* ------------------------------------------------------------- shape geometry */
// Rounded medical-cross motif (normalised 0..1) - the deck's signature decoration.
const CROSS = [
  ['M', 0.872, 0.347],
  ['L', 0.655, 0.359],
  ['L', 0.667, 0.139],
  ['C', 0.672, 0.057, 0.658, 0.007, 0.571, 0.002],
  ['C', 0.537, 0.0, 0.531, 0.001, 0.511, 0.001],
  ['L', 0.457, 0.0],
  ['C', 0.376, 0.0, 0.35, 0.048, 0.35, 0.131],
  ['C', 0.35, 0.134, 0.35, 0.136, 0.35, 0.139],
  ['L', 0.362, 0.359],
  ['L', 0.145, 0.347],
  ['C', 0.065, 0.343, 0.007, 0.348, 0.003, 0.43],
  ['C', 0.002, 0.433, 0.0, 0.5, 0.0, 0.503],
  ['C', 0.002, 0.542, 0.003, 0.552, 0.005, 0.577],
  ['C', 0.005, 0.659, 0.067, 0.665, 0.09, 0.668],
  ['C', 0.118, 0.671, 0.143, 0.67, 0.145, 0.67],
  ['L', 0.362, 0.658],
  ['L', 0.35, 0.878],
  ['C', 0.346, 0.96, 0.384, 0.994, 0.439, 0.998],
  ['C', 0.472, 1.001, 0.485, 0.999, 0.508, 1.0],
  ['C', 0.566, 1.001, 0.542, 0.999, 0.567, 0.998],
  ['C', 0.663, 0.988, 0.668, 0.968, 0.668, 0.886],
  ['C', 0.668, 0.883, 0.668, 0.88, 0.668, 0.878],
  ['L', 0.656, 0.658],
  ['L', 0.872, 0.67],
  ['C', 0.953, 0.674, 0.992, 0.632, 0.999, 0.581],
  ['C', 0.999, 0.579, 1.0, 0.54, 1.0, 0.537],
  ['C', 1.0, 0.509, 1.0, 0.48, 1.0, 0.451],
  ['C', 1.0, 0.368, 0.962, 0.347, 0.881, 0.346],
  ['C', 0.878, 0.346, 0.875, 0.347, 0.872, 0.347],
  ['Z'],
];

// Flat-shouldered up arrow used on the bar-chart infographic (slide 12).
// `shaft` is where the head meets the shaft, as a fraction of the shape height.
function upArrow(shaft) {
  const tip = shaft * 0.728;
  const ctl = tip + (shaft - tip) * 0.367;
  return [
    ['M', 0.972, tip],
    ['L', 0.558, 0.02],
    ['C', 0.527, -0.007, 0.473, -0.007, 0.441, 0.02],
    ['L', 0.028, tip],
    ['C', -0.032, ctl, 0.011, shaft, 0.096, shaft],
    ['L', 0.216, shaft],
    ['L', 0.216, 1.0],
    ['L', 0.787, 1.0],
    ['L', 0.787, shaft],
    ['L', 0.906, shaft],
    ['C', 0.989, shaft, 1.032, ctl, 0.972, tip],
    ['Z'],
  ];
}

/** Scale a normalised path into pptxgenjs `points` for a w x h box. */
function scalePath(segments, w, h) {
  return segments.map((s) => {
    if (s[0] === 'Z') return { close: true };
    if (s[0] === 'M') return { x: s[1] * w, y: s[2] * h, moveTo: true };
    if (s[0] === 'L') return { x: s[1] * w, y: s[2] * h };
    return {
      x: s[5] * w,
      y: s[6] * h,
      curve: { type: 'cubic', x1: s[1] * w, y1: s[2] * h, x2: s[3] * w, y2: s[4] * h },
    };
  });
}

// pptxgenjs surfaces a preset shape's two adjust values only via `angleRange`,
// which it writes as adj1/adj2 in 60000ths - so scale raw OOXML adjusts by 1/60000.
const adj = (a1, a2) => [a1 / 60000, a2 / 60000];
// roundRect takes its corner radius in inches; OOXML stores a fraction of min(w,h).
const radius = (a, w, h) => (a / 100000) * Math.min(w, h);

/* ------------------------------------------------------------------- helpers */
/** Decorative rounded cross. */
function cross(s, x, y, w, h, color, rotate, flipH) {
  s.addShape('custGeom', {
    x, y, w, h,
    points: scalePath(CROSS, w, h),
    fill: { color },
    line: { type: 'none' },
    rotate: rotate || 0,
    flipH: !!flipH,
  });
}

/** "Rectangle: top corners rounded" - the arch / shield motif. */
function arch(s, x, y, w, h, color, o) {
  o = o || {};
  s.addShape('round2SameRect', {
    x, y, w, h,
    fill: { color },
    line: { type: 'none' },
    angleRange: adj(o.adj1 === undefined ? 10255 : o.adj1, o.adj2 === undefined ? 50000 : o.adj2),
    rotate: o.rotate || 0,
    flipH: o.flipH !== false,
  });
}

// Stand-ins for the deck's small white SVG line icons: [preset shape, rotation].
const ICON = {
  pulse: ['heart', 0], bag: ['flowChartInternalStorage', 0], syringe: ['stripedRightArrow', 315],
  stetho: ['donut', 0], pills: ['can', 0], bottle: ['can', 0], dna: ['chartX', 0],
  doc: ['flowChartDocument', 0], trophy: ['trapezoid', 180], star: ['star5', 0],
  cloud: ['cloud', 0], laptop: ['flowChartManualInput', 0], heart: ['heart', 0],
  key: ['donut', 0], gear: ['gear6', 0], bulb: ['teardrop', 135],
  pin: ['teardrop', 135], phone: ['moon', 45], globe: ['donut', 0], mail: ['flowChartPreparation', 0],
};
function icon(s, kind, x, y, w, h, color) {
  const [shape, rotate] = ICON[kind];
  s.addShape(shape, {
    x, y, w, h,
    fill: { type: 'none' },
    line: { color: color || C.white, width: 1 },
    rotate,
  });
}

// Soft drop shadow for the deck's floating white cards. pptxgenjs rewrites the
// object it is handed, so hand each shape a fresh one.
const cardShadow = () => ({ type: 'outer', color: '000000', opacity: 0.13, blur: 26, offset: 2, angle: 90 });

/** White/teal magnifier badge that sits in every slide's top-right corner. */
function searchBadge(s, onDark) {
  const disc = onDark ? C.white : C.a3;
  const ink = onDark ? C.a5 : C.white;
  s.addShape('ellipse', { x: 12.576, y: 0.353, w: 0.419, h: 0.419, fill: { color: disc }, line: { type: 'none' } });
  s.addShape('ellipse', { x: 12.7, y: 0.475, w: 0.15, h: 0.15, fill: { type: 'none' }, line: { color: ink, width: 1 } });
  s.addShape('line', { x: 12.83, y: 0.605, w: 0.06, h: 0.06, line: { color: ink, width: 1.25 } });
}

/** Thin arrow glyph used inside the small square "next" buttons. */
function arrowGlyph(s, x, y, w, color) {
  s.addShape('line', { x, y, w, h: 0, line: { color, width: 1.5, endArrowType: 'triangle' } });
}

/** Label + percentage + track/fill progress bar. */
function progressBar(s, o) {
  s.addText(o.label, { x: o.labelX, y: o.labelY, w: o.labelW, h: 0.269, fontFace: MINOR, fontSize: 10, bold: true, color: C.body, valign: 'top' });
  s.addText(o.pct, { x: o.pctX, y: o.labelY, w: o.pctW, h: 0.269, fontFace: MINOR, fontSize: 10, bold: true, color: C.body, align: 'right', valign: 'top' });
  s.addShape('line', { x: o.x, y: o.y, w: o.w, h: 0, line: { color: C.grey, width: 7 } });
  s.addShape('line', { x: o.x, y: o.y, w: o.fillW, h: 0, line: { color: o.color, width: 7 } });
}

/** Two-tone heading: a light accent phrase next to a bold black phrase. */
function heading(s, x, y, w, h, first, second, o) {
  o = o || {};
  const boldFirst = !!o.boldFirst;
  s.addText(
    [
      { text: first, options: { bold: boldFirst, color: o.firstColor || C.a5 } },
      { text: second, options: { bold: !boldFirst, color: o.secondColor || C.black } },
    ],
    {
      x, y, w, h,
      fontFace: MAJOR, fontSize: o.fontSize || 40, color: C.black,
      align: o.align || 'left', valign: 'top',
      lineSpacingMultiple: o.lineSpacing,
    }
  );
}

/** Body copy: 10.5pt Roboto, 150% leading. */
function body(s, text, x, y, w, h, o) {
  o = o || {};
  s.addText(text, {
    x, y, w, h,
    fontFace: MINOR, fontSize: 10.5,
    color: o.color || C.body,
    align: o.align || 'left',
    valign: 'top',
    lineSpacingMultiple: o.lineSpacing === undefined ? 1.5 : o.lineSpacing,
  });
}

/** Small teal eyebrow label, e.g. "Healthcare Solutions". */
function eyebrow(s, text, x, y, w, o) {
  o = o || {};
  s.addText(text, { x, y, w, h: 0.337, fontFace: MAJOR, fontSize: 14, color: o.color || C.a5, bold: !!o.bold, align: o.align || 'left', valign: 'top' });
}

/** Pill button with centred caption. */
function button(s, o) {
  s.addShape('roundRect', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: { color: o.fill || C.btn },
    line: { type: 'none' },
    rectRadius: radius(o.adj === undefined ? 14359 : o.adj, o.w, o.h),
  });
  if (o.text) {
    s.addText(o.text, {
      x: o.textX, y: o.textY, w: o.textW, h: o.textH || 0.37,
      fontFace: MAJOR, fontSize: o.fontSize || 16, bold: !!o.bold,
      color: o.color || C.white, align: 'center', valign: o.valign || 'middle',
    });
  }
}

/**
 * pptxgenjs shapes only take solid fills, so the deck's full-bleed gradient is
 * painted as a stack of horizontal bands interpolating between two colours.
 */
function gradientBackdrop(s, topHex, bottomHex, bands) {
  const rgb = (h) => [0, 2, 4].map((i) => parseInt(h.substr(i, 2), 16));
  const [t, b] = [rgb(topHex), rgb(bottomHex)];
  const h = 7.5 / bands;
  for (let i = 0; i < bands; i++) {
    const k = (i + 0.5) / bands;
    const hex = t.map((v, j) => Math.round(v + (b[j] - v) * k)).map((v) => `0${v.toString(16)}`.slice(-2)).join('').toUpperCase();
    s.addShape('rect', { x: 0, y: i * h, w: 13.333, h: h + 0.01, fill: { color: hex }, line: { type: 'none' } });
  }
}

/** Flat stand-in for a photograph. */
function photo(s, o) {
  s.addShape('rect', { x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: o.fill || C.hairline }, line: { type: 'none' } });
  s.addText('[image]', {
    x: o.x, y: o.y + o.h / 2 - 0.2, w: o.w, h: 0.4,
    fontFace: MINOR, fontSize: 11, color: o.labelColor || 'B0B0B0', align: 'center', valign: 'middle',
  });
}

/* ------------------------------------------------------------- shared copy */
const LOREM = {
  sit: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. Nunc viverra',
  full: 'Lorem ipsum dolor amet, consectetuer adipiscing  elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.',
  short: 'Lorem ipsum dolor amet, consectetuer adipiscing  elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus',
  tiny: 'Lorem ipsum dolor amet, consectetuer adipiscing  elit. Maecenas porttitor congue massa. Fusce posuere',
  volu: 'In voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident',
  voluShort: 'In voluptate velit esse cillum dolore eu fugiat nulla',
  pleasure: 'chooses to enjoy a pleasure that has no annoying consequences, ',
  painful: 'consequences that are extremely painful. Nor again is there anyone who',
};

/* ================================================================== slides */

/** Slides 1 & 20 share the dark gradient cover layout. */
function coverSlide(s, boldWord, lightWord, boldFace) {
  gradientBackdrop(s, C.nightTop, C.nightBottom, 24);
  cross(s, 8.21, 0.657, 3.692, 3.633, C.crossGlow, 90);
  cross(s, 10.219, 4.337, 2.165, 2.13, C.crossDeep, 90);
  cross(s, 6.906, 0.576, 1.363, 1.341, C.crossDeep2, 90);

  s.addText(boldWord, { x: 1.084, y: 2.187, w: 6.448, h: 1.313, fontFace: boldFace, fontSize: 72, bold: true, color: C.white, valign: 'top' });
  s.addText(lightWord, { x: 1.035, y: 3.294, w: 7.703, h: 1.313, fontFace: MAJOR, fontSize: 72, color: C.white, valign: 'top' });
  body(s, LOREM.sit, 1.145, 4.793, 7.703, 0.602, { color: C.white });

  button(s, { x: 1.187, y: 5.755, w: 2.361, h: 0.602, fill: C.a3d, text: 'More Detail', bold: true, fontSize: 18, textX: 1.487, textY: 5.854, textW: 1.76, textH: 0.404 });
  button(s, { x: 3.858, y: 5.755, w: 3.662, h: 0.602, fill: C.white, text: 'Start  Presentation ', fontSize: 18, color: C.a2, textX: 4.138, textY: 5.854, textW: 2.564, textH: 0.404 });
  s.addShape('roundRect', { x: 6.982, y: 5.875, w: 0.361, h: 0.361, fill: { color: C.a3d }, line: { type: 'none' }, rectRadius: radius(16667, 0.361, 0.361) });
  arrowGlyph(s, 7.068, 6.056, 0.189, C.white);
  searchBadge(s, true);
}

function slide1(s) {
  coverSlide(s, 'Healthcare', 'Proposal Deck', MAJOR);
}

function slide2(s) {
  cross(s, 4.431, 3.952, 3.156, 3.105, C.crossPale, 270, true);
  s.addText(
    [
      { text: 'Improving', options: { color: C.a5 } },
      { text: ' ', options: { color: C.a1 } },
      { text: 'Patient Care Quality', options: { bold: true, color: C.black } },
    ],
    { x: 8.095, y: 1.426, w: 3.836, h: 2.121, fontFace: MAJOR, fontSize: 40, color: C.black, valign: 'top' }
  );
  body(s, LOREM.full, 8.095, 3.995, 4.155, 1.132);
  button(s, { x: 8.22, y: 5.565, w: 2.173, h: 0.509, adj: 14036, text: 'More Detail', textX: 8.493, textY: 5.635, textW: 1.627 });
  searchBadge(s);
  cross(s, 0.457, 0.74, 1.61, 1.584, C.crossMint, 270, true);
}

function slide3(s) {
  cross(s, 6.074, 4.371, 2.944, 2.897, C.crossPale, 270, true);
  cross(s, 0.238, 0.124, 1.312, 1.291, C.crossMint, 270, true);
  heading(s, 1.282, 1.645, 4.865, 1.447, 'Accessible ', 'Health Services');
  eyebrow(s, 'Healthcare Solutions', 1.282, 3.531, 2.298);
  body(s, LOREM.full, 1.282, 3.914, 5.254, 0.867);
  button(s, { x: 1.407, y: 5.346, w: 2.173, h: 0.509, adj: 18831, text: 'More Detail', textX: 1.661, textY: 5.415, textW: 1.664 });
  searchBadge(s);
  arch(s, 7.551, 1.177, 1.119, 1.178, C.a3d);
  icon(s, 'pulse', 7.856, 1.511, 0.511, 0.511);
}

function slide4(s) {
  arch(s, 1.262, 1.558, 5.146, 5.415, C.a3, { rotate: 180 });
  heading(s, 7.916, 1.766, 4.448, 1.447, 'Enhancing ', 'Infrastructure');
  body(s, LOREM.full, 7.916, 3.654, 4.155, 1.132);
  button(s, { x: 8.041, y: 5.225, w: 2.173, h: 0.509, adj: 16434, text: 'More Detail', textX: 8.314, textY: 5.294, textW: 1.627 });
  cross(s, 11.484, 5.983, 1.312, 1.291, C.crossMint, 270, true);
  cross(s, 5.653, -0.487, 2.545, 2.504, C.crossPale, 270, true);
  searchBadge(s);
  cross(s, 0.78, 2.048, 1.255, 1.234, C.crossMint, 270, true);
}

function slide5(s) {
  cross(s, -0.98, 3.567, 4.564, 4.491, C.crossPale, 270, true);
  body(s, LOREM.tiny, 7.797, 0.981, 4.028, 0.602);
  heading(s, 1.208, 1.63, 5.938, 1.447, 'Health Equity For ', 'Communities');
  eyebrow(s, 'Preventive Health', 1.208, 3.633, 2.528);
  body(s, LOREM.short, 1.208, 3.969, 5.458, 0.602);
  eyebrow(s, 'Healthcare Solutions', 1.208, 4.932, 2.298);
  body(s, LOREM.short, 1.208, 5.268, 5.458, 0.602);
  cross(s, 6.05, -0.213, 1.447, 1.424, C.crossMint, 270, true);
  searchBadge(s);
  arch(s, 6.933, 5.252, 1.147, 1.207, C.a3d, { rotate: 180 });
  icon(s, 'bag', 7.279, 5.628, 0.455, 0.455);
}

function slide6(s) {
  cross(s, -0.684, 4.529, 3.549, 3.492, C.crossPale, 270, true);
  searchBadge(s);
  s.addText('Healthcare Proposal Deck presentation', { x: 6.656, y: 1.462, w: 4.325, h: 0.303, fontFace: MINOR, fontSize: 12, color: C.black, charSpacing: 1, valign: 'top' });
  heading(s, 6.602, 1.765, 5.19, 1.447, 'Health Equity For ', 'Communities');
  eyebrow(s, 'Healthcare Solutions', 3.048, 5.213, 2.298, { align: 'right' });
  body(s, LOREM.short, 1.282, 5.549, 4.063, 0.867, { align: 'right' });
  arch(s, 11.391, 3.369, 1.148, 1.209, C.a3d, { rotate: 180 });
  icon(s, 'syringe', 11.692, 3.773, 0.547, 0.547);
  cross(s, 5.76, -0.261, 1.147, 1.128, C.crossMint, 270, true);
}

function slide7(s) {
  cross(s, 0.224, 4.275, 3.345, 3.291, C.crossPale, 270, true);
  searchBadge(s);
  heading(s, 1.16, 1.56, 4.306, 1.447, 'Ensuring Food ', 'And Nutrition');
  button(s, { x: 1.206, y: 3.543, w: 2.633, h: 0.616, adj: 21602, text: 'More Detail', fontSize: 20, textX: 1.519, textY: 3.632, textW: 2.007, textH: 0.438 });
  s.addShape('roundRect', { x: 4.019, y: 3.543, w: 0.616, h: 0.616, fill: { color: C.white }, line: { type: 'none' }, rectRadius: radius(21602, 0.616, 0.616), rotate: 180, shadow: cardShadow() });
  arrowGlyph(s, 4.182, 3.851, 0.292, C.a3);
  body(s, LOREM.full, 1.16, 4.809, 4.306, 1.132);
  arch(s, 6.551, 1.798, 1.148, 1.209, C.a3d, { adj1: 18720 });
  icon(s, 'stetho', 6.876, 2.153, 0.5, 0.5);
  cross(s, 3.767, 6.605, 1.147, 1.128, C.crossMint, 270, true);
}

function slide8(s) {
  cross(s, 10.363, 4.982, 2.809, 2.764, C.crossPale, 270, true);
  searchBadge(s);
  progressBar(s, { x: 1.249, y: 5.421, w: 4.552, fillW: 3.106, color: C.a2, label: 'undertakes', labelX: 1.119, labelY: 5.016, labelW: 1.506, pct: '70%', pctX: 4.866, pctW: 0.995 });
  body(s, LOREM.short, 1.119, 5.762, 5.254, 0.602);
  heading(s, 7.473, 1.354, 4.865, 1.447, 'Strengthening ', 'Capabilities');
  body(s, 'Lorem ipsum dolor amet, consectetuer adipiscing  elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros', 7.473, 5.497, 4.759, 0.867);

  s.addShape('roundRect', { x: 7.466, y: 3.435, w: 4.61, h: 1.079, fill: { color: C.btn }, line: { type: 'none' }, rectRadius: radius(18831, 4.61, 1.079) });
  s.addText('78%', { x: 7.782, y: 3.621, w: 1.68, h: 0.707, fontFace: MINOR, fontSize: 36, bold: true, color: C.white, align: 'center', valign: 'middle' });
  s.addText('24+', { x: 10.087, y: 3.621, w: 1.68, h: 0.707, fontFace: MINOR, fontSize: 36, bold: true, color: C.white, align: 'center', valign: 'middle' });
  s.addShape('line', { x: 9.771, y: 3.575, w: 0, h: 0.799, line: { color: C.a3, width: 1 } });
  eyebrow(s, 'Healthcare Solutions', 7.473, 5.113, 2.298);

  cross(s, 7.932, -0.122, 1.147, 1.128, C.crossMint, 270, true);
  arch(s, 5.599, 1.354, 0.964, 1.014, C.a3d, { adj1: 18720 });
  icon(s, 'pills', 5.863, 1.609, 0.436, 0.436);
}

function slide9(s) {
  cross(s, 7.463, 5.803, 2.167, 2.132, C.crossPale, 270, true);
  cross(s, 12.149, 1.801, 1.517, 1.493, C.crossPale, 270, true);
  searchBadge(s);
  heading(s, 3.03, 0.55, 7.593, 0.774, 'Meet Our ', 'Best Team', { align: 'center' });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus.Pellentesque',
    1.394, 6.181, 10.545, 0.602, { color: C.black, align: 'center' });

  const team = [
    { cardX: 1.338, name: 'Henry Oliver', nameX: 1.535, nameW: 2.706, nameAlign: 'center', roleX: 2.047, roleY: 5.08 },
    { cardX: 5.117, name: 'Emily Victoria', nameX: 5.44, nameW: 2.454, nameAlign: 'left', roleX: 5.853, roleY: 5.081 },
    { cardX: 8.896, name: 'Julian Styles', nameX: 9.163, nameW: 2.565, nameAlign: 'left', roleX: 9.631, roleY: 5.112 },
  ];
  team.forEach((t) => {
    arch(s, t.cardX, 1.571, 3.1, 4.214, C.card, { rotate: 180 });
    s.addText(t.name, { x: t.nameX, y: 4.75, w: t.nameW, h: 0.438, fontFace: MAJOR, fontSize: 20, color: C.white, align: t.nameAlign, valign: 'top' });
    s.addText('the master-builder', { x: t.roleX, y: t.roleY, w: 1.628, h: 0.348, fontFace: MINOR, fontSize: 10.5, color: C.white, align: 'center', valign: 'top' });
  });
  cross(s, 0.961, 0.409, 1.147, 1.128, C.crossMint, 270, true);
}

function slide10(s) {
  gradientBackdrop(s, C.nightTop, C.nightBottom, 24);
  cross(s, 6.818, 1.724, 4.774, 4.698, C.crossGlow, 90);
  cross(s, 10.514, 5.948, 2.165, 2.13, C.crossDeep, 90);
  cross(s, 11.677, 1.48, 1.363, 1.341, C.crossDeep2, 90);

  s.addText('BREAK', { x: 1.389, y: 1.132, w: 6.345, h: 2.036, fontFace: MAJOR, fontSize: 115, color: C.white, valign: 'middle' });
  s.addText('SLIDE', { x: 1.483, y: 2.679, w: 6.345, h: 2.036, fontFace: MINOR, fontSize: 115, bold: true, color: C.white, valign: 'middle' });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna p', 1.57, 4.71, 4.749, 0.602, { color: C.white });

  button(s, { x: 1.621, y: 5.766, w: 2.361, h: 0.602, fill: C.a3d, text: 'Nexs Slide', bold: true, fontSize: 18, textX: 1.922, textY: 5.865, textW: 1.76, textH: 0.404 });
  s.addShape('roundRect', { x: 4.283, y: 5.766, w: 0.602, h: 0.602, fill: { color: C.white }, line: { type: 'none' }, rectRadius: radius(16090, 0.602, 0.602) });
  arrowGlyph(s, 4.426, 6.067, 0.314, C.a2);
  searchBadge(s, true);
}

function slide11(s) {
  cross(s, 9.785, 4.279, 3.236, 3.184, C.crossPale, 270, true);
  searchBadge(s);
  s.addShape('ellipse', { x: 4.382, y: 2.206, w: 4.57, h: 4.57, fill: { color: C.a2l }, line: { type: 'none' } });

  // Four petals of the circular infographic: [x, y, w, h, colour, rotation, icon]
  const petals = [
    [4.819, 2.634, 1.782, 1.782, C.a2, 270, 'trophy'],
    [6.733, 2.634, 1.782, 1.782, C.a3, 0, 'star'],
    [4.819, 4.519, 1.782, 1.828, C.a5, 180, 'cloud'],
    [6.71, 4.543, 1.828, 1.782, C.a2d, 90, 'laptop'],
  ];
  petals.forEach(([x, y, w, h, color, rot, kind]) => {
    arch(s, x, y, w, h, color, { rotate: rot });
    icon(s, kind, x + w / 2 - 0.21, y + h / 2 - 0.21, 0.42, 0.42);
  });

  // Four callouts around the wheel: [titleX, titleY, bodyX, bodyY, align]
  const callouts = [
    [1.591, 2.305, 1.035, 2.738, 'right'],
    [1.591, 4.956, 1.035, 5.379, 'right'],
    [9.436, 2.305, 9.436, 2.707, 'left'],
    [9.436, 4.956, 9.436, 5.379, 'left'],
  ];
  callouts.forEach(([tx, ty, bx, by, align]) => {
    s.addText('Title Here', { x: tx, y: ty, w: 2.307, h: 0.415, fontFace: MAJOR, fontSize: 14, bold: true, color: C.black, align, valign: 'top', lineSpacingMultiple: 1.5 });
    body(s, LOREM.volu, bx, by, 2.862, 0.867, { align });
  });

  heading(s, 2.834, 0.723, 7.667, 0.653, 'Inpographic ', 'Design', { align: 'center', lineSpacing: 0.9 });
  cross(s, -0.321, -0.212, 2.186, 2.151, C.crossMint, 270, true);
}

function slide12(s) {
  cross(s, 7.185, 2.963, 5.533, 5.444, C.crossPale, 270, true);
  searchBadge(s);

  // Column chart: [x, y, w, h, colour]
  const bars = [
    [1.188, 4.686, 1.397, 2.814, C.a2dd],
    [2.733, 2.311, 1.4, 5.189, C.a2],
    [4.281, 5.903, 1.4, 1.597, C.a3],
    [5.829, 3.468, 1.4, 4.032, C.a5],
  ];
  bars.forEach(([x, y, w, h, color]) => {
    s.addShape('round2SameRect', {
      x, y, w, h,
      fill: { color },
      line: { type: 'none' },
      angleRange: adj(Math.round((0.329 / Math.min(w, h)) * 100000), 0),
    });
  });

  // Pale arrows rising between the columns; the head is always 0.511" tall.
  [[2.233, 5.55, 0.852, 1.95], [3.781, 6.515, 0.852, 0.985], [5.329, 6.515, 0.852, 0.985]].forEach(([x, y, w, h]) => {
    s.addShape('custGeom', { x, y, w, h, points: scalePath(upArrow(0.511 / h), w, h), fill: { color: C.a3l }, line: { type: 'none' } });
  });

  [['heart', 1.68, 4.916], ['trophy', 3.227, 2.609], ['key', 6.323, 3.755], ['star', 4.775, 6.195]]
    .forEach(([kind, x, y]) => icon(s, kind, x, y, 0.413, 0.413));

  // Legend on the right: [title, titleY, bodyY]
  const legend = [
    ['Advancing Research', 2.513, 2.927],
    ['Emergency Medical', 3.604, 3.986],
    ['Creating Inclusive', 4.662, 5.077],
    ['Lifestyle Promotion', 5.754, 6.168],
  ];
  legend.forEach(([title, ty, by]) => {
    s.addText(title, { x: 8.518, y: ty, w: 3.1, h: 0.415, fontFace: MAJOR, fontSize: 14, color: C.a5, valign: 'top', lineSpacingMultiple: 1.5 });
    body(s, LOREM.voluShort, 8.518, by, 3.725, 0.337);
  });

  heading(s, 2.834, 0.723, 7.667, 0.653, 'Inpographic ', 'Design', { align: 'center', lineSpacing: 0.9 });
  cross(s, -0.321, -0.212, 2.186, 2.151, C.crossMint, 270, true);
}

function slide13(s) {
  cross(s, 9.785, 4.279, 3.236, 3.184, C.crossPale, 270, true);
  heading(s, 2.834, 0.723, 7.667, 0.653, 'Inpographic ', 'Design', { align: 'center', lineSpacing: 0.9 });

  // Pale rounded track linking the four stages.
  s.addShape('roundRect', { x: 4.524, y: 2.748, w: 4.309, h: 3.378, fill: { type: 'none' }, line: { color: C.ring, width: 1.5 }, rectRadius: 0.85 });
  arch(s, 6.078, 5.593, 1.184, 1.184, C.a3, { rotate: 180 });
  arch(s, 6.078, 2.181, 1.184, 1.184, C.a2d);
  s.addShape('line', { x: 5.567, y: 4.455, w: 2.205, h: 0, line: { color: C.ringLine, width: 1 } });

  // Numbered nodes: [x, y, disc colour, label, labelX, labelW]
  const nodes = [
    [5.102, 3.145, C.a1, '01', 5.226, 0.409],
    [7.592, 3.145, C.a3, '02', 7.696, 0.447],
    [7.592, 5.094, C.a4, '03', 7.696, 0.447],
    [5.102, 5.094, C.a2, '04', 5.197, 0.467],
  ];
  nodes.forEach(([x, y, color]) => s.addShape('ellipse', { x, y, w: 0.657, h: 0.657, fill: { color }, line: { type: 'none' } }));

  s.addShape('ellipse', { x: 5.968, y: 3.722, w: 1.429, h: 1.429, fill: { color: C.white }, line: { type: 'none' } });
  icon(s, 'trophy', 6.354, 4.108, 0.657, 0.657, C.ringLine);

  s.addShape('roundRect', { x: 3.514, y: 4.047, w: 2.053, h: 0.816, fill: { color: C.a2dd }, line: { type: 'none' }, rectRadius: radius(15153, 2.053, 0.816) });
  s.addShape('roundRect', { x: 7.772, y: 4.047, w: 2.053, h: 0.816, fill: { color: C.a2 }, line: { type: 'none' }, rectRadius: radius(19025, 2.053, 0.816) });
  [3.682, 7.941].forEach((x) => s.addText('Title Here', { x, y: 4.253, w: 1.717, h: 0.404, fontFace: MAJOR, fontSize: 18, color: C.grey, align: 'center', valign: 'top' }));

  // Four corner paragraphs (two paragraphs each, as in the source).
  [[9.408, 2.255, 'left'], [9.408, 5.77, 'left'], [0.772, 2.255, 'right'], [0.772, 5.77, 'right']].forEach(([x, y, align]) => {
    s.addText(
      [
        { text: 'In voluptate velit esse cillum dolore ', options: { breakLine: true } },
        { text: 'eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident' },
      ],
      { x, y, w: 3.177, h: 0.867, fontFace: MINOR, fontSize: 10.5, color: C.body, align, valign: 'top', lineSpacingMultiple: 1.5 }
    );
  });

  icon(s, 'gear', 6.432, 2.536, 0.475, 0.475);
  icon(s, 'bulb', 6.432, 5.948, 0.475, 0.475);
  nodes.forEach(([, , , label, lx, lw], i) => {
    const ly = i < 2 ? 3.305 : 5.254;
    s.addText(label, { x: lx, y: ly, w: lw, h: 0.337, fontFace: MAJOR, fontSize: 14, color: C.grey, align: 'center', valign: 'middle' });
  });
  searchBadge(s);
  cross(s, -0.321, -0.212, 2.186, 2.151, C.crossMint, 270, true);
}

function slide14(s) {
  // Tablet photograph -> flat placeholder.
  s.addShape('roundRect', { x: 6.757, y: 1.403, w: 5.869, h: 4.25, fill: { color: '111111' }, line: { type: 'none' }, rectRadius: 0.3 });
  s.addShape('rect', { x: 7.131, y: 1.62, w: 5.12, h: 3.82, fill: { color: C.white }, line: { type: 'none' } });
  s.addText('[image]', { x: 7.131, y: 3.33, w: 5.12, h: 0.4, fontFace: MINOR, fontSize: 11, color: 'C8C8C8', align: 'center', valign: 'middle' });

  cross(s, 8.699, 4.613, 4.087, 4.022, C.crossMint, 270, true);
  searchBadge(s);
  heading(s, 1.065, 1.389, 4.865, 1.447, 'Accessible ', 'Health Services');
  body(s, 'Lorem ipsum dolor amet, consectetuer adipiscing  elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna', 1.065, 5.165, 4.865, 0.867);
  button(s, { x: 1.065, y: 3.533, w: 2.633, h: 0.616, adj: 21602, text: 'More Detail', fontSize: 20, textX: 1.378, textY: 3.623, textW: 2.007, textH: 0.438 });
  eyebrow(s, 'Healthcare Solutions', 1.065, 4.701, 2.298);
  s.addShape('roundRect', { x: 4.095, y: 3.573, w: 0.616, h: 0.616, fill: { color: C.white }, line: { type: 'none' }, rectRadius: radius(21602, 0.616, 0.616), rotate: 180, shadow: cardShadow() });
  arrowGlyph(s, 4.041, 3.841, 0.292, C.a3);
  cross(s, -1.343, -0.724, 2.646, 2.604, C.crossPale, 270, true);
  cross(s, 5.388, 6.636, 1.555, 1.53, C.crossPale, 270, true);
  arch(s, 7.421, 0.783, 1.152, 1.212, C.a3d, { rotate: 270 });
  icon(s, 'bottle', 7.752, 1.141, 0.455, 0.497);
}

function slide15(s) {
  cross(s, -0.995, -1.264, 4.789, 4.712, C.crossPale, 270, true);
  searchBadge(s);
  // Phone photograph -> flat placeholder (body plus the camera bump on its left).
  s.addShape('roundRect', { x: 1.42, y: 0.9, w: 0.5, h: 2.56, fill: { color: '181818' }, line: { type: 'none' }, rectRadius: 0.16 });
  s.addShape('roundRect', { x: 1.66, y: 0.47, w: 5.06, h: 7.4, fill: { color: '111111' }, line: { type: 'none' }, rectRadius: 0.5 });
  s.addShape('roundRect', { x: 1.86, y: 0.69, w: 4.66, h: 6.96, fill: { color: C.white }, line: { type: 'none' }, rectRadius: 0.36 });
  s.addText('[image]', { x: 1.86, y: 3.7, w: 4.66, h: 0.4, fontFace: MINOR, fontSize: 11, color: 'C8C8C8', align: 'center', valign: 'middle' });

  s.addShape('roundRect', { x: 0.679, y: 5.016, w: 3.528, h: 1.733, fill: { color: C.cardTeal }, line: { type: 'none' }, rectRadius: radius(7157, 3.528, 1.733) });
  s.addShape('roundRect', { x: 3.955, y: 1.2, w: 3.528, h: 1.733, fill: { color: C.white }, line: { type: 'none' }, rectRadius: radius(7157, 3.528, 1.733), shadow: cardShadow() });
  s.addText('$76,234,00', { x: 1.093, y: 5.333, w: 2.483, h: 0.438, fontFace: MAJOR, fontSize: 20, bold: true, color: C.white, valign: 'top' });
  body(s, LOREM.pleasure, 1.093, 5.807, 2.7, 0.625, { color: C.white });
  s.addText('$23,713,00', { x: 4.37, y: 1.535, w: 2.483, h: 0.438, fontFace: MAJOR, fontSize: 20, bold: true, color: C.a5, valign: 'top' });
  body(s, LOREM.pleasure, 4.37, 1.973, 2.7, 0.625);

  heading(s, 8.246, 1.425, 4.591, 1.447, 'Partnerships ', 'Healthcare');
  body(s, 'which toil and pain can procure him some great pleasure. To take a trivial example, which of us ever undertakes laborious', 7.219, 5.473, 4.498, 0.602);
  progressBar(s, { x: 7.337, y: 4.155, w: 4.839, fillW: 2.638, color: C.a3, label: 'undertakes', labelX: 7.219, labelY: 3.749, labelW: 1.372, pct: '55%', pctX: 11.334, pctW: 0.906 });
  progressBar(s, { x: 7.356, y: 4.897, w: 4.82, fillW: 3.289, color: C.a2, label: 'undertakes', labelX: 7.219, labelY: 4.492, labelW: 1.595, pct: '70%', pctX: 11.186, pctW: 1.054 });
  cross(s, 11.517, 6.06, 1.402, 1.379, C.crossMint, 270, true);
}

function slide16(s) {
  cross(s, 4.599, -1.766, 4.25, 4.182, C.crossPale, 270, true);
  // Hand-holding-phone photograph -> flat placeholder.
  s.addShape('roundRect', { x: 6.66, y: 0.71, w: 4.15, h: 6.3, fill: { color: '111111' }, line: { type: 'none' }, rectRadius: 0.42 });
  s.addShape('roundRect', { x: 6.85, y: 0.92, w: 3.77, h: 5.88, fill: { color: C.white }, line: { type: 'none' }, rectRadius: 0.3 });
  s.addText('[image]', { x: 6.85, y: 3.65, w: 3.77, h: 0.4, fontFace: MINOR, fontSize: 11, color: 'C8C8C8', align: 'center', valign: 'middle' });
  searchBadge(s);

  s.addShape('roundRect', { x: 1.15, y: 4.533, w: 3.833, h: 1.64, fill: { color: C.cardTeal }, line: { type: 'none' }, rectRadius: radius(9217, 3.833, 1.64) });
  s.addShape('roundRect', { x: 5.354, y: 4.533, w: 3.833, h: 1.64, fill: { color: C.white }, line: { type: 'none' }, rectRadius: radius(5153, 3.833, 1.64), shadow: cardShadow() });
  heading(s, 1.111, 1.464, 4.498, 1.447, 'Public-private ', 'Partnerships');
  body(s, 'because it is pleasure, but because those who do not know to pursue pleasure rationally encounter consequences that', 1.153, 3.421, 4.201, 0.602, { color: C.black });

  eyebrow(s, 'Analytics For Every', 5.845, 4.856, 2.918, { color: C.black });
  body(s, LOREM.painful, 5.845, 5.198, 2.918, 0.625, { color: C.black });
  eyebrow(s, 'Analytics For Every', 1.619, 4.856, 2.918, { color: C.white });
  body(s, LOREM.painful, 1.619, 5.202, 2.918, 0.625, { color: C.white });

  cross(s, -0.148, 6.342, 1.64, 1.614, C.crossMint, 270, true);
  button(s, { x: 9.474, y: 1.589, w: 2.06, h: 0.759, fill: C.a3d, text: '9373+', fontSize: 28, textX: 9.838, textY: 1.682, textW: 1.387, textH: 0.572 });
}

function slide17(s) {
  searchBadge(s);
  photo(s, { x: 0.851, y: 1.534, w: 6.941, h: 4.538, fill: C.hairline, labelColor: 'D8D8D8' }); // US map artwork
  button(s, { x: 7.839, y: 3.998, w: 4.234, h: 0.759, fill: C.a3d, text: 'Healthcare Solutions', fontSize: 20, textX: 8.288, textY: 4.172, textW: 3.338, textH: 0.438 });
  heading(s, 8.706, 1.314, 4.234, 2.121, 'Designing ', 'Sustainable Health');
  body(s, 'Lorem ipsum dolor amet, consectetuer aàelit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus Lorem ipsum dolor amet, consectetuer aàelit', 7.792, 5.319, 4.691, 0.867);
  cross(s, 4.071, 5.769, 2.597, 2.556, C.crossPale, 90, true);
  arch(s, 5.451, -0.035, 1.347, 1.417, C.a3d, { rotate: 270 });
  icon(s, 'dna', 5.839, 0.353, 0.441, 0.53);
  cross(s, -0.005, 0.574, 1.373, 1.351, C.crossMint, 90, true);
}

function slide18(s) {
  cross(s, -1.044, 4.138, 5.138, 5.055, C.crossPale, 270, true);
  searchBadge(s);
  photo(s, { x: 5.259, y: 1.358, w: 7.736, h: 5.308, fill: 'F8F8F8', labelColor: 'DEDEDE' }); // world map artwork
  heading(s, 1.03, 1.301, 5.245, 1.447, 'Accessible Health ', 'Services');

  [1.015, 4.332].forEach((x) => {
    s.addText('$8,340', { x, y: 4.481, w: 1.943, h: 0.64, fontFace: MAJOR, fontSize: 32, color: C.a5, valign: 'top' });
    body(s, 'Lorem ipsum, consectetuer adipiscing  elit. Maecenas porttitor Lorem ipsum, consectetuer', x, 5.148, 2.6, 0.867);
  });
  s.addShape('line', { x: 3.841, y: 4.432, w: 0, h: 1.767, line: { color: C.hairline, width: 1 } });
  progressBar(s, { x: 1.121, y: 3.685, w: 3.197, fillW: 2.181, color: C.a2, label: 'undertakes', labelX: 1.03, labelY: 3.28, labelW: 1.058, pct: '70%', pctX: 3.661, pctW: 0.699 });

  cross(s, 6.13, -0.319, 2.144, 2.11, C.crossMint, 270, true);
  arch(s, 11.013, 4.974, 1.347, 1.417, C.a3d, { rotate: 90, adj1: 20417 });
  icon(s, 'doc', 11.487, 5.414, 0.494, 0.539);
}

function slide19(s) {
  cross(s, 9.241, 3.61, 3.825, 3.764, C.crossPale, 270, true);
  cross(s, 2.947, -0.787, 2.655, 2.612, C.crossMint, 270, true);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Quisque non elit mauris. Cras euismod, ', 1.119, 5.489, 2.612, 0.867);

  // Contact rows: [squareX, squareY, textX, textY, textW, label, icon]
  const rows = [
    [4.708, 5.152, 5.412, 5.144, 2.885, '123 Anywhere St., Any City', 'pin'],
    [4.713, 5.877, 5.412, 5.869, 2.885, '+123-456-7890', 'phone'],
    [8.565, 5.152, 9.27, 5.093, 2.569, 'www.reallygreatsite.com', 'globe'],
    [8.565, 5.877, 9.27, 5.861, 2.569, 'hello@reallygreatsite.com', 'mail'],
  ];
  rows.forEach(([sx, sy, tx, ty, tw, label, kind]) => {
    s.addShape('roundRect', { x: sx, y: sy, w: 0.438, h: 0.438, fill: { color: C.a3d }, line: { type: 'none' }, rectRadius: radius(16667, 0.438, 0.438) });
    icon(s, kind, sx + 0.13, sy + 0.13, 0.18, 0.18);
    s.addText(label, { x: tx, y: ty, w: tw, h: 0.454, fontFace: MINOR, fontSize: 14, bold: true, color: C.black, valign: 'middle', lineSpacingMultiple: 1.5 });
  });

  heading(s, 4.708, 1.789, 7.507, 0.909, 'Contact ', 'Information', { fontSize: 48, boldFirst: true, firstColor: C.black, secondColor: C.a5 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, a maximus leo ligula at dolor. Morbi et malesuada purus. ', 4.708, 2.828, 7.507, 0.603);
  eyebrow(s, 'Patient Care', 1.119, 5.152, 2.479, { color: C.black, bold: true });
  searchBadge(s);
}

function slide20(s) {
  coverSlide(s, 'Thanks', 'For Attention', MINOR);
}

/* ==================================================================== build */
const BUILDERS = [slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'WIDE';
pptx.title = 'Healthcare Proposal Deck';

BUILDERS.forEach((build) => {
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  build(slide);
});

pptx
  .writeFile({ fileName: path.join(__dirname, '0390f017-0580-4d09-97b4-aef049866ad4_grok_final.pptx') })
  .then((f) => console.log('wrote', f));
