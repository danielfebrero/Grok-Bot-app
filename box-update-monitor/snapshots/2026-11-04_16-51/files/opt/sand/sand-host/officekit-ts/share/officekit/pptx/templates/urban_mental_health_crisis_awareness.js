/*
 * "Urban Literal | Mental Health" — 20-slide deck rebuilt with pptxgenjs.
 * Run: node 15ede332-d468-4732-a676-49dacea93008_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Design tokens
 * ------------------------------------------------------------------ */
const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

const C = {
  brown: '764B2F', // primary brand brown
  lilac: 'D4BBD7',
  pink: 'FFD3F7',
  cream: 'F1ECE7',
  paper: 'FEFCF9', // theme bg1 / bg2
  ink: '131313', // theme tx1 / tx2
  ash: 'E7E7E7', // tx2 @ 10% luminance — the light grey panels
  bar1: 'DBD5D5',
  bar2: 'A6A6A6',
  hatch: 'BCBAB7', // pinstripes filling the empty picture placeholders
  iconEdge: '565655',
  iconFrame: '9E9E9D',
  iconBlue: '4A87C4',
  iconSun: 'E8A93B',
};

const FONT = 'Karla';
const FONT_MED = 'Karla Medium';
// A handful of runs fall back to the theme fonts instead of naming Karla.
const THEME_MAJOR = 'Manrope bold';
const THEME_MINOR = 'Manrope';

// Body copy that the template repeats across many slides.
const L = {
  long: 'Lorem ipsum dolor sit dolor sit ametis asda amet, consectetuer adipiscing elitisu. Dolor sit eget ligula eget dolor sit amet consetectuer adispiscing lorem ipsum dolor sit amet.',
  mid: 'Lorem ipsum dolor sit dolor sit ametis asda amet, consectetuer adipiscing elitisu. Dolor sit eget ligula eget dolor sit amet consetectuer adispiscing.',
  cut: 'Lorem ipsum dolor sit dolor sit ametis asda amet, consectetuer adipiscing elitisu. Dolor sit eget',
  dot: 'Lorem ipsum dolor sit dolor sit ametis asda amet, consectetuer adipiscing elitisu. Dolor sit eget.',
  circle: 'Lorem ipsum dolor sit dolora sit ametis asda ametsuliasq, consectetuer adipiscing.',
  tiny: 'Lorem ipsum dolor sit dolora sit ametis as.',
  case: 'Lorem ipsum dolor sitasdwea amet, consectetuer dolor sit.',
  service: 'Lorem ipsum dolor sitasdwea amet, consectetuer',
  panel: 'Lorem ipsum dolor sit dolor sit ametis asda amet, consectetuer adipiscing elitisu. Dolor sit eget amet consetectuer adispsicing ligula eget dolor sit amet ligula.',
};

/* ------------------------------------------------------------------ *
 * Drawing helpers
 * ------------------------------------------------------------------ */
let SH = null; // pptx.ShapeType, assigned in build()

const rect = (s, x, y, w, h, fill) => s.addShape(SH.rect, { x, y, w, h, fill: { color: fill } });

// `adj` is the OOXML roundRect adjust value (fraction of the shorter side).
const roundRect = (s, x, y, w, h, fill, adj = 0.16667, rotate = 0) =>
  s.addShape(SH.roundRect, {
    x, y, w, h,
    fill: { color: fill },
    rectRadius: adj * Math.min(w, h),
    rotate,
  });

const oval = (s, x, y, w, h, fill, line) =>
  s.addShape(SH.ellipse, { x, y, w, h, fill: { color: fill }, ...(line ? { line } : {}) });

// `fit: 'resize'` mirrors the template's <a:spAutoFit/> on every text box.
const text = (s, body, o) =>
  s.addText(body, { fontFace: FONT, fontSize: 10, color: C.ink, valign: 'top', fit: 'resize', ...o });

// Each argument becomes its own paragraph.
const ml = (...parts) => parts.map(t => ({ text: t, options: { breakLine: true } }));

// Bulleted list; `indent` is the hanging indent in points.
const bullets = (items, indent) =>
  items.map(t => ({ text: t, options: { breakLine: true, bullet: { indent } } }));

/* Repeating page furniture ----------------------------------------- */
function logo(s, color) {
  oval(s, 0.76, 0.729, 0.146, 0.146, color);
  oval(s, 0.906, 0.729, 0.146, 0.146, color);
  text(s, 'Company', { x: 1.2, y: 0.652, w: 0.947, h: 0.303, fontSize: 12, color, wrap: false });
}

const dateTag = (s, color) =>
  text(s, ml('Dates:', '09/04/2045'), { x: 8.606, y: 0.623, w: 0.994, h: 0.438, color, wrap: false });

const brandTag = (s, color) =>
  text(s, ml('Urban Literal', 'Mental Health'), { x: 11.268, y: 0.623, w: 1.734, h: 0.438, color });

/* Small circular badge holding a diagonal arrow (the "next" chip). */
function arrowBadge(s, x, y, d, fill, arrow, rotate, outline) {
  oval(s, x, y, d, d * 0.945, fill, outline ? { color: outline, width: 1 } : null);
  s.addShape(SH.line, {
    x: x + d * 0.27, y: y + d * 0.47, w: d * 0.45, h: 0,
    line: { color: arrow, width: 0.75, endArrowType: 'arrow' },
    rotate,
  });
}

/* Empty picture placeholder: 45-degree pinstripes plus a small "photo" icon.
   (Raster art in the source deck is replaced by these programmatic stand-ins.) */
const HATCH_STEP = 0.085;

function hatchPanel(s, x, y, w, h) {
  rect(s, x, y, w, h, C.paper);
  for (let d = -h; d < w; d += HATCH_STEP) {
    const x0 = Math.max(x + d, x);
    const y0 = y + Math.max(-d, 0);
    const x1 = Math.min(x + d + h, x + w);
    const y1 = y + Math.min(h, w - d);
    if (x1 > x0) s.addShape(SH.line, { x: x0, y: y0, w: x1 - x0, h: y1 - y0, line: { color: C.hatch, width: 0.5 } });
  }
}

function photoIcon(s, cx, cy) {
  const w = 0.87, h = 0.69;
  const x = cx - w / 2, y = cy - h / 2;
  s.addShape(SH.rect, { x, y, w, h, fill: { color: 'FAFAFA' }, line: { color: C.iconEdge, width: 1.25 } });
  s.addShape(SH.rect, { x: x + 0.07, y: y + 0.06, w: w - 0.14, h: h - 0.12, fill: { color: 'FAFAFA' }, line: { color: C.iconFrame, width: 0.5 } });
  oval(s, x + 0.15, y + 0.11, 0.13, 0.13, 'F8DB8F', { color: C.iconSun, width: 0.75 });
  // Two overlapping mountains inside the frame.
  [[[0.17, 0.82], [0.40, 0.53], [0.58, 0.82]], [[0.32, 0.82], [0.62, 0.37], [0.84, 0.82]]].forEach(pts => {
    s.addShape(SH.custGeom, {
      x, y, w, h,
      fill: { color: '83BEEC' }, line: { color: C.iconBlue, width: 1 },
      points: pts.map(([px, py]) => ({ x: px * w, y: py * h })).concat([{ close: true }]),
    });
  });
}

function photoFrame(s, x, y, w, h) {
  hatchPanel(s, x, y, w, h);
  photoIcon(s, x + w / 2, y + h / 2);
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

// 1 — Title
function slide01(s) {
  rect(s, 7.8, 0, 5.533, 7.5, C.brown);
  roundRect(s, 7.644, 1.388, 3.48, 4.499, C.cream);
  roundRect(s, 7.166, 1.388, 3.48, 4.499, C.lilac);

  logo(s, C.brown);
  dateTag(s, C.paper);
  brandTag(s, C.paper);

  text(s, 'Understanding The Mental Health Crisis in Urban Areas',
    { x: 0.604, y: 1.746, w: 6.063, h: 3.736, fontFace: FONT_MED, fontSize: 54 });

  const contact = [
    ['Website:', 'www.example.com', 0.76, 1.413],
    ['Contact:', '+11 222 3333 4444', 2.549, 1.413],
    ['Email:', 'example@email.com', 4.337, 1.555],
  ];
  contact.forEach(([label, value, x, vw]) => {
    text(s, label, { x, y: 6.194, w: 1.413, h: 0.373, fontSize: 12, align: 'justify', lineSpacingMultiple: 1.5 });
    text(s, value, { x, y: 6.572, w: vw, h: 0.327, align: 'justify', lineSpacingMultiple: 1.5 });
  });

  roundRect(s, 10.782, 5.368, 1.794, 1.199, C.pink, 0.09983);
  arrowBadge(s, 12.226, 5.481, 0.246, C.ink, C.paper, 52.16);
  text(s, 'Effected People', { x: 10.95, y: 5.89, w: 1.392, h: 0.303, fontSize: 12 });
  text(s, '12.357K ', { x: 10.931, y: 6.062, w: 1.734, h: 0.438, fontSize: 20 });
  text(s, 'Slide To Continue', { x: 11.268, y: 6.63, w: 1.359, h: 0.269, color: C.paper });
}

// 2 — Introduction
function slide02(s) {
  rect(s, 0, 0, 13.333, 5.344, C.brown);
  roundRect(s, 2.739, 0.394, 3.255, 5.223, C.cream, 0.16667, 270);
  roundRect(s, 2.184, 0.394, 3.255, 5.223, C.lilac, 0.16667, 270);

  logo(s, C.paper);
  dateTag(s, C.paper);
  brandTag(s, C.paper);

  text(s, 'Introduction', { x: 0.76, y: 5.996, w: 5.517, h: 1.111, fontFace: FONT_MED, fontSize: 60 });
  text(s, 'Hello, We are From Company',
    { x: 8.606, y: 1.67, w: 1.461, h: 0.675, fontSize: 12, color: C.paper, lineSpacingMultiple: 1.5 });
  text(s, L.mid,
    { x: 8.606, y: 2.836, w: 3.967, h: 0.832, color: C.paper, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, 'Mental Health 2040',
    { x: 8.606, y: 4.138, w: 2.277, h: 0.327, color: C.paper, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, 'Where ideas meet impact.', { x: 10.694, y: 6.63, w: 1.879, h: 0.269 });
}

// 3 — Content index
function slide03(s) {
  rect(s, 5.278, 0, 8.056, 7.5, C.brown);

  const INDEX = [
    ['01', ['Case ', 'Studies'], 5.838, 7.046, 1.002, C.paper],
    ['02', ['Public', 'Awareness'], 8.211, 9.436, 1.301, C.lilac],
    ['03', ['Urban Habits'], 10.737, 11.948, 1.008, C.cream],
    ['04', ['Risk Groups ', 'in City Care'], 5.838, 7.047, 1.32, C.paper],
    ['05', ['Our Collaboration'], 8.211, 9.436, 1.301, C.lilac],
    ['06', ['Future Step'], 10.739, 11.95, 0.847, C.cream],
  ];
  INDEX.forEach(([num, label, nx, lx, lw, numColor], i) => {
    const ny = i < 3 ? 2.241 : 4.347;
    const ly = i < 3 ? 2.648 : 4.747;
    text(s, num, { x: nx, y: ny, w: 1.736, h: 1.212, fontSize: 66, color: numColor });
    text(s, ml(...label), { x: lx, y: ly, w: lw, h: 0.505, fontSize: 12, color: C.paper });
  });

  logo(s, C.brown);
  dateTag(s, C.paper);
  brandTag(s, C.paper);

  text(s, 'Content Index', { x: 0.616, y: 2.241, w: 3.606, h: 2.121, fontFace: FONT_MED, fontSize: 60 });
  text(s, 'Maximus lorem guis, blandit elit. Proin tellus euismod, a congue libero nisl sed do euismod.',
    { x: 0.632, y: 4.577, w: 4.15, h: 0.675, fontSize: 12, lineSpacingMultiple: 1.5 });

  roundRect(s, 0.662, 5.708, 2.035, 0.427, C.pink, 0.5);
  text(s, 'Presentation', { x: 0.804, y: 5.77, w: 1.381, h: 0.303, fontSize: 12, align: 'center' });
  arrowBadge(s, 2.326, 5.792, 0.26, C.paper, C.ink, 0, C.ink);
}

// 4 — Case studies
function slide04(s) {
  roundRect(s, 10.276, -0.12, 2.949, 3.155, C.brown, 0.16667, 270);
  roundRect(s, 7.134, -0.112, 2.932, 3.155, C.lilac, 0.16667, 270);

  text(s, 'Case Studies',
    { x: 0.561, y: 1.388, w: 5.358, h: 1.111, fontFace: FONT_MED, fontSize: 60, wrap: false });

  [['What is Our Concern', 7.477, C.ink], ['How it Benefits', 10.632, C.paper]].forEach(([head, x, color]) => {
    text(s, head, { x, y: 0.655, w: 1.151, h: 0.675, fontFace: THEME_MAJOR, fontSize: 12, color, lineSpacingMultiple: 1.5 });
    text(s, L.cut, { x, y: 1.636, w: 2.325, h: 0.832, fontFace: THEME_MINOR, color, align: 'justify', lineSpacingMultiple: 1.5 });
  });

  logo(s, C.brown);
  photoFrame(s, 0, 2.932, 13.333, 4.568);
}

// 5 — Public awareness
function slide05(s) {
  s.background = { color: C.brown };

  const CARDS = [
    [C.lilac, -1.866, '76%', 0.436, ['Public A ', 'City A'], 0.703, 1.535, 1.817],
    [C.pink, 0.518, '34%', 2.824, ['Public B', 'City B'], 3.092, 1.151, 4.204],
    [C.cream, 2.918, '12%', 5.221, ['Public C', 'City C'], 5.378, 1.151, 6.628],
  ];
  CARDS.forEach(([fill, cardY]) => roundRect(s, 8.732, cardY, 2.347, 6.478, fill, 0.16667, 270));

  text(s, ml('Public', 'Awareness'),
    { x: 0.76, y: 4.994, w: 4.53, h: 2.121, fontFace: FONT_MED, fontSize: 60, color: C.paper });
  text(s, 'Why it Matters?',
    { x: 0.76, y: 1.67, w: 1.261, h: 0.675, fontSize: 12, color: C.paper, lineSpacingMultiple: 1.5 });
  text(s, L.long,
    { x: 0.76, y: 2.598, w: 4.53, h: 0.832, color: C.paper, align: 'justify', lineSpacingMultiple: 1.5 });

  CARDS.forEach(([, , pct, pctY, label, labelY, labelW, audienceY]) => {
    text(s, pct, { x: 9.079, y: pctY, w: 3.599, h: 2.036, fontSize: 115 });
    text(s, ml(...label), { x: 7.077, y: labelY, w: labelW, h: 0.675, fontSize: 12, lineSpacingMultiple: 1.5 });
    text(s, 'Target Audiences', { x: 7.077, y: audienceY, w: 1.351, h: 0.269 });
  });

  logo(s, C.paper);
}

// 6 — People urban habits
function slide06(s) {
  rect(s, 6.322, 0, 7.011, 7.5, C.brown);
  roundRect(s, 7.776, 2.199, 3.625, 5.776, C.cream, 0.16667, 270);
  roundRect(s, 8.204, 2.199, 3.625, 5.776, C.lilac, 0.16667, 270);

  text(s, 'People Urban Habits',
    { x: 7.078, y: 0.577, w: 5.495, h: 2.121, fontFace: FONT_MED, fontSize: 60, color: C.paper });

  [['At-Risk Groups in City Care', 1.77, 1.494, 2.697], ['Urban Life and Mental ', 4.714, 1.261, 5.642]]
    .forEach(([head, hy, hw, by]) => {
      text(s, head, { x: 0.76, y: hy, w: hw, h: 0.675, fontSize: 12, lineSpacingMultiple: 1.5 });
      text(s, L.long, { x: 0.76, y: by, w: 4.53, h: 0.832, align: 'justify', lineSpacingMultiple: 1.5 });
    });

  logo(s, C.brown);
}

// 7 — Risk groups: problem / solution
function slide07(s) {
  rect(s, 0, 0, 13.333, 4.214, C.brown);
  roundRect(s, 8.426, 0.113, 2.121, 4.672, C.cream, 0.16667, 270);
  roundRect(s, 8.687, 0.113, 2.121, 4.672, C.lilac, 0.16667, 270);
  rect(s, 7.7, 4.214, 5.633, 3.286, C.ash);

  text(s, ml('Risk Groups ', 'in City Care'),
    { x: 0.76, y: 1.473, w: 4.873, h: 2.121, fontFace: FONT_MED, fontSize: 60, color: C.paper });

  text(s, 'Our Problem',
    { x: 0.76, y: 5.088, w: 0.906, h: 0.675, fontFace: FONT_MED, fontSize: 12, lineSpacingMultiple: 1.5 });
  text(s, L.long, { x: 0.76, y: 6.016, w: 4.53, h: 0.832, align: 'justify', lineSpacingMultiple: 1.5 });

  text(s, 'Our Solution',
    { x: 8.606, y: 5.088, w: 0.906, h: 0.675, fontFace: FONT_MED, fontSize: 12, lineSpacingMultiple: 1.5 });
  text(s, bullets([
    'Solution one, lorem ipsum dolor sit amet consetecuter.',
    'Solution two, lorem ipsum dolro sit amet consetectuer.',
    'Soultion three, lorem ipsum dolor sit amet consetectu.',
  ], 18), { x: 8.606, y: 6.016, w: 3.967, h: 0.832, align: 'justify', lineSpacingMultiple: 1.5 });

  logo(s, C.paper);
  dateTag(s, C.paper);
  brandTag(s, C.paper);
}

// 8 — Our collaboration
function slide08(s) {
  text(s, 'Our Collaboration', { x: 0.76, y: 1.388, w: 5.964, h: 2.121, fontFace: FONT_MED, fontSize: 60 });

  const CIRCLES = [
    [C.brown, 0.76, ['Urban Institution'], 1.547, 1.261, 1.184, C.paper],
    [C.pink, 3.742, ['Urban', 'Communities'], 4.545, 1.261, 4.146, C.ink],
    [C.lilac, 6.724, ['Urban', 'Government'], 7.448, 1.418, 7.128, C.ink],
    [C.cream, 9.706, ['Urban', 'Social Activists'], 10.43, 1.418, 10.109, C.ink],
  ];
  CIRCLES.forEach(([fill, cx]) => oval(s, cx, 4.032, 2.867, 2.867, fill));
  CIRCLES.forEach(([, , label, lx, lw, bx, color]) => {
    text(s, ml(...label),
      { x: lx, y: 4.659, w: lw, h: 0.675, fontSize: 12, color, align: 'center', lineSpacingMultiple: 1.5 });
    text(s, L.circle,
      { x: bx, y: 5.466, w: 2.06, h: 0.832, color, align: 'center', lineSpacingMultiple: 1.5 });
  });

  text(s, 'Prioritize Efforts', { x: 8.606, y: 1.445, w: 1.261, h: 0.675, fontSize: 12, lineSpacingMultiple: 1.5 });
  text(s, L.mid, { x: 8.606, y: 2.448, w: 3.967, h: 0.832, align: 'justify', lineSpacingMultiple: 1.5 });

  logo(s, C.brown);
  dateTag(s, C.ink);
  brandTag(s, C.ink);
}

// 9 — Future steps
function slide09(s) {
  text(s, 'Future Steps', { x: 0.633, y: 1.445, w: 4.306, h: 2.121, fontFace: FONT_MED, fontSize: 60 });
  rect(s, 6.244, 0, 7.089, 7.5, C.brown);

  const STEPS = [
    [['Mental Health Emergency ', 'Response'], 7.021, 2.064, 2.23, 0.505, false, 7.016, 2.613],
    [['Policy and System Gaps'], 10.118, 2.064, 2.043, 0.303, false, 10.113, 2.613],
    [['Program Impact and Evaluation'], 7.016, 4.521, 1.967, 0.505, true, 7.016, 5.07],
    [['Trauma-Informed Care'], 10.156, 4.521, 1.967, 0.303, false, 10.113, 5.07],
  ];
  STEPS.forEach(([head, hx, hy, hw, hh, wrapOn, bx, by]) => {
    text(s, ml(...head), { x: hx, y: hy, w: hw, h: hh, fontSize: 12, color: C.paper, wrap: wrapOn });
    text(s, L.dot, { x: bx, y: by, w: 2.305, h: 0.832, color: C.paper, align: 'justify', lineSpacingMultiple: 1.5 });
  });

  logo(s, C.brown);
  dateTag(s, C.paper);
  brandTag(s, C.paper);
  photoFrame(s, 0, 4.055, 6.244, 3.445);
}

// 10 — Strategic steps
function slide10(s) {
  text(s, 'Strategic Steps', { x: 0.76, y: 1.388, w: 9.817, h: 1.111, fontFace: FONT_MED, fontSize: 60 });

  const NOTES = [
    [0.76, 5.989, 2.022, 6.072, 3.265, 'Lorem ipsum dolor sit dolor sit ametis asda amet, consectetuer adipiscing elitisu. Dolor sit.'],
    [5.397, 5.195, 6.728, 5.243, 1.75, 'Lorem ipsum dolor sit dolor sit ametis asdaa.'],
    [8.606, 6.275, 9.919, 6.32, 2.654, 'Lorem ipsum dolor sit dolor sit ametis asda amet, consectetuer adipiscing.'],
  ];
  NOTES.forEach(([hx, hy, bx, by, bw, body]) => {
    text(s, 'Strategic Subject:', { x: hx, y: hy, w: 1.261, h: 0.675, fontFace: THEME_MAJOR, fontSize: 12, lineSpacingMultiple: 1.5 });
    text(s, body, { x: bx, y: by, w: bw, h: 0.58, fontFace: THEME_MINOR, align: 'justify', lineSpacingMultiple: 1.5 });
  });

  logo(s, C.brown);
  dateTag(s, C.ink);
  brandTag(s, C.ink);
}

// 11 — Grow in rates
function slide11(s) {
  rect(s, 0, 0, 4.183, 7.5, C.brown);
  rect(s, 7.967, 0, 5.367, 7.5, C.lilac);

  text(s, ml('Grow ', 'In Rates'), { x: 4.388, y: 1.455, w: 3.373, h: 2.121, fontFace: FONT_MED, fontSize: 60 });

  // Left case study (on brown) / right case study (on lilac)
  const STUDIES = [
    { color: C.paper, x: 0.924, title: 'Case Studies 1', titleY: 1.658, titleW: 1.78,
      client: 'Client: Xyz Corporation', clientY: 2.331, bodyY: 2.882, bodyW: 2.325,
      objY: 3.975, objW: 1.577, listY: 4.551,
      list: ['Increase brand awareness.', 'Improve websites.'] },
    { color: C.ink, x: 8.606, title: 'Case Studies 2', titleY: 1.711, titleW: 1.911,
      client: 'Client: Abc Enterprises', clientY: 2.384, bodyX: 8.656, bodyY: 2.854, bodyW: 2.557,
      objX: 8.694, objY: 3.947, objW: 1.734, listX: 8.656, listY: 4.523,
      list: ['Monthly branding design.', 'Attract any customers'] },
  ];
  STUDIES.forEach(d => {
    text(s, d.title, { x: d.x, y: d.titleY, w: d.titleW, h: 0.37, fontSize: 16, color: d.color });
    text(s, d.client, { x: d.x, y: d.clientY, w: 2.225, h: 0.303, fontSize: 12, color: d.color });
    text(s, L.case, { x: d.bodyX || d.x, y: d.bodyY, w: d.bodyW, h: 0.58, color: d.color, align: 'justify', lineSpacingMultiple: 1.5 });
    text(s, 'Main Objectives', { x: d.objX || d.x, y: d.objY, w: d.objW, h: 0.303, fontSize: 12, color: d.color });
    text(s, bullets(d.list, 13.5), { x: d.listX || d.x, y: d.listY, w: d.bodyW, h: 0.58, color: d.color, align: 'justify', lineSpacingMultiple: 1.5 });
  });

  text(s, '25%', { x: 0.924, y: 5.572, w: 1.169, h: 0.64, fontSize: 32, color: C.paper });
  text(s, ml('Increase in', 'Converstion'),
    { x: 2.123, y: 5.484, w: 1.162, h: 0.58, color: C.paper, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, ml('Reduction in', 'Manual Entry'),
    { x: 8.754, y: 6.124, w: 1.279, h: 0.58, align: 'justify', lineSpacingMultiple: 1.5 });

  roundRect(s, 10.984, 5.572, 1.794, 1.199, C.pink, 0.09983);
  arrowBadge(s, 12.428, 5.685, 0.246, C.ink, C.paper, 52.16);
  text(s, 'percentage', { x: 11.065, y: 5.737, w: 1.279, h: 0.327, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, '30%', { x: 11.026, y: 6.064, w: 1.169, h: 0.64, fontSize: 32 });

  logo(s, C.paper);
  dateTag(s, C.ink);
  brandTag(s, C.ink);
  photoFrame(s, 4.183, 4.055, 3.784, 3.445);
}

// 12 — Access to mental health services
function slide12(s) {
  text(s, 'Access to Mental Health Services',
    { x: 6.667, y: 3.991, w: 6.006, h: 3.13, fontFace: FONT_MED, fontSize: 60 });
  text(s, 'Why it Matters?', { x: 6.667, y: 1.749, w: 1.261, h: 0.675, fontSize: 12, lineSpacingMultiple: 1.5 });
  text(s, L.long, { x: 6.667, y: 2.677, w: 4.53, h: 0.832, align: 'justify', lineSpacingMultiple: 1.5 });

  logo(s, C.brown);
  dateTag(s, C.ink);
  brandTag(s, C.ink);

  rect(s, 0, 4.138, 5.555, 2.761, C.brown);
  rect(s, 0, 1.376, 5.555, 2.761, C.lilac);

  [['Therapy', 0.805, 2.127, 2.677, C.ink], ['Community Support', 1.764, 4.944, 5.493, C.paper]]
    .forEach(([head, hw, hy, by, color]) => {
      text(s, head, { x: 0.699, y: hy, w: hw, h: 0.303, fontSize: 12, color, wrap: false });
      text(s, L.panel, { x: 0.695, y: by, w: 3.967, h: 0.832, color, align: 'justify', lineSpacingMultiple: 1.5 });
    });
}

// 13 — Services overview
function slide13(s) {
  rect(s, 0, 0, 13.333, 5.344, C.brown);
  roundRect(s, 6.209, 0.118, 3.256, 5.776, C.cream, 0.16667, 270);
  roundRect(s, 6.521, 0.118, 3.256, 5.776, C.lilac, 0.16667, 270);
  rect(s, 0, 5.344, 4.667, 2.156, C.ash);

  text(s, 'Services Overview', { x: 5.41, y: 6.011, w: 7.386, h: 1.111, fontFace: FONT_MED, fontSize: 60 });
  text(s, 'Excellent Services',
    { x: 0.76, y: 1.67, w: 1.484, h: 0.675, fontSize: 12, color: C.paper, lineSpacingMultiple: 1.5 });
  text(s, L.mid, { x: 0.76, y: 2.836, w: 3.74, h: 0.832, color: C.paper, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, 'Bryant Davis – Head  Of Service',
    { x: 0.76, y: 4.138, w: 2.406, h: 0.327, color: C.paper, lineSpacingMultiple: 1.5 });

  [['Service 1', 0.76, 0.76], ['Service 2', 2.582, 2.61]].forEach(([label, lx, bx]) => {
    text(s, label, { x: lx, y: 5.811, w: 1.413, h: 0.373, fontSize: 12, align: 'justify', lineSpacingMultiple: 1.5 });
    text(s, L.service, { x: bx, y: 6.183, w: 1.413, h: 0.831, align: 'justify', lineSpacingMultiple: 1.5 });
  });

  logo(s, C.paper);
  dateTag(s, C.paper);
  brandTag(s, C.paper);
}

// 14 — Service pricing
function slide14(s) {
  const COLS = [
    [0, C.brown, C.paper, 'Content Marketing', 1.852, '$50.2k+', 2.191, 0.788, 6.563, 2.136],
    [4.444, C.lilac, C.ink, 'Social Media Marketing', 2.169, '$25.4k+', 2.169, 5.29, 6.563, 6.631],
    [8.889, C.paper, C.ink, 'Email Marketing', 2.469, '$35.8k+', 2.274, 9.767, 6.507, 11.104],
  ];
  COLS.forEach(([bx, fill]) => rect(s, bx, 4.554, 4.444, 2.946, fill));

  text(s, ml('Service', 'Pricing'), { x: 0.594, y: 1.629, w: 3.703, h: 2.121, fontFace: FONT_MED, fontSize: 60 });

  COLS.forEach(([, , color, label, labelW, price, priceW, tx, footY, useX]) => {
    text(s, label, { x: tx, y: 5.418, w: labelW, h: 0.303, fontSize: 12, color });
    text(s, price, { x: tx, y: 5.822, w: priceW, h: 0.707, fontSize: 36, color });
    text(s, 'Allocated Budget', { x: tx, y: footY, w: 1.341, h: 0.269, color });
    text(s, 'Uses of resources', { x: useX, y: footY, w: 1.469, h: 0.269, color });
  });

  logo(s, C.brown);
  photoFrame(s, 5.496, 0, 7.837, 4.554);
}

// 15 — Company activity campaign
function slide15(s) {
  rect(s, 7.595, 0, 5.751, 7.5, C.brown);

  text(s, ml('Company', 'Activity', 'Campaign'),
    { x: 8.211, y: 2.432, w: 4.362, h: 3.13, fontFace: FONT_MED, fontSize: 60, color: C.paper });
  text(s, 'Why it Matters?', { x: 0.76, y: 1.393, w: 1.261, h: 0.675, fontSize: 12, lineSpacingMultiple: 1.5 });
  text(s, 'Lorem ipsum dolor sit dolor sit ametis asda amet, consectetuer adipiscing elitisu. Dolor sit eget ligula eget dolor sit amet consetectuer adispiscing lorem ipsum.',
    { x: 0.76, y: 2.321, w: 3.017, h: 1.085, align: 'justify', lineSpacingMultiple: 1.5 });

  roundRect(s, 2.556, 2.94, 3.255, 5.223, C.cream, 0.16667, 270);
  roundRect(s, 2.15, 2.94, 3.255, 5.223, C.lilac, 0.16667, 270);

  roundRect(s, 10.575, 5.736, 1.794, 1.199, C.pink, 0.09983);
  arrowBadge(s, 12.02, 5.848, 0.246, C.ink, C.paper, 52.16);
  text(s, 'percentage', { x: 10.657, y: 5.9, w: 1.279, h: 0.327, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, '30%', { x: 10.618, y: 6.227, w: 1.169, h: 0.64, fontSize: 32 });

  logo(s, C.brown);
  dateTag(s, C.paper);
  brandTag(s, C.paper);
}

// 16 — People and resources
function slide16(s) {
  rect(s, 7.711, 0, 5.622, 7.5, C.brown);

  text(s, 'People and Resources', { x: 0.76, y: 1.936, w: 4.256, h: 3.13, fontFace: FONT_MED, fontSize: 60 });

  const BUBBLES = [
    [C.lilac, 6.28, 0.601, ['Community'], 6.776, 1.206, 0.372, 6.597, 1.796],
    [C.pink, 8.368, 2.498, ['Urban People'], 8.873, 3.099, 0.372, 8.694, 3.633],
    [C.cream, 6.614, 4.629, ['Networking Events'], 7.153, 5.066, 0.675, 6.974, 5.859],
  ];
  BUBBLES.forEach(([fill, cx, cy]) => oval(s, cx, cy, 2.271, 2.271, fill));
  BUBBLES.forEach(([, , , label, lx, ly, lh, bx, by]) => {
    text(s, ml(...label),
      { x: lx, y: ly, w: 1.261, h: lh, fontSize: 12, bold: true, align: 'center', lineSpacingMultiple: 1.5 });
    text(s, L.tiny, { x: bx, y: by, w: 1.619, h: 0.58, align: 'center', lineSpacingMultiple: 1.5 });
  });

  text(s, 'Lorem ipsum dolor sit dolor sit ametis asda amet, consectetuer adipiscing elitisu. Dolor sit eget ligula eget dolor sit amet consetectuer.',
    { x: 0.76, y: 6.32, w: 4.862, h: 0.58, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, 'Data Strategies', { x: 11.268, y: 6.63, w: 1.359, h: 0.269, color: C.paper });

  logo(s, C.brown);
  dateTag(s, C.paper);
  brandTag(s, C.paper);
}

// 17 — Data and evidence-based care (bar chart drawn with rounded bars)
function slide17(s) {
  rect(s, 0, 0, 8.111, 7.5, C.brown);
  rect(s, 8.111, 5.379, 5.356, 2.121, C.ash);

  roundRect(s, 1.2, 1.576, 1.794, 1.199, C.pink, 0.09983);
  arrowBadge(s, 2.644, 1.688, 0.246, C.ink, C.paper, 52.16);
  text(s, '4,76K', { x: 1.301, y: 1.774, w: 1.604, h: 0.64, fontSize: 32 });
  text(s, 'Data Statistic', { x: 1.248, y: 2.344, w: 1.488, h: 0.303, fontSize: 12 });

  text(s, 'Data and Evidence-Based Care',
    { x: 8.562, y: 1.222, w: 4.207, h: 4.14, fontFace: FONT_MED, fontSize: 60 });
  text(s, 'Lorem ipsum dolor sit dolor sit ametis asda amet, consectetuer adipiscing elitisu. Dolor',
    { x: 3.33, y: 1.471, w: 3.056, h: 0.58, color: C.paper, align: 'justify', lineSpacingMultiple: 1.5 });

  // Back bars (paper) then the coloured foreground bars.
  const BARS = [
    [1.55, 3.816, 0.816, 2.308, C.paper], [2.833, 3.019, 0.816, 3.155, C.paper],
    [4.127, 3.816, 0.816, 2.308, C.paper], [5.431, 2.995, 0.816, 3.155, C.paper],
    [4.127, 4.53, 0.816, 1.611, C.bar1], [1.543, 4.539, 0.827, 1.611, C.lilac],
    [2.82, 3.878, 0.843, 2.296, C.pink], [5.424, 4.223, 0.837, 1.928, C.bar2],
  ];
  BARS.forEach(([x, y, w, h, fill]) => roundRect(s, x, y, w, h, fill, 0.5));

  [['2020', 1.709], ['2025', 2.926], ['2030', 4.288], ['2035', 5.607]].forEach(([year, x]) =>
    text(s, year, { x, y: 6.576, w: 0.657, h: 0.327, color: C.paper, lineSpacingMultiple: 1.5 }));

  text(s, '$2.850.00', { x: 8.606, y: 6.104, w: 2.201, h: 0.572, fontSize: 28 });
  text(s, 'Conversion Rates', { x: 8.606, y: 6.576, w: 1.475, h: 0.327, lineSpacingMultiple: 1.5 });
  s.addShape(SH.line, { x: 11.022, y: 6.39, w: 0, h: 0.349, line: { color: C.ink, width: 1 } });
  text(s, '13%', { x: 11.29, y: 5.942, w: 0.978, h: 0.735, fontSize: 28, lineSpacingMultiple: 1.5 });
  text(s, 'Increase Rates', { x: 11.29, y: 6.576, w: 1.283, h: 0.327, lineSpacingMultiple: 1.5 });

  logo(s, C.paper);
  dateTag(s, C.ink);
  brandTag(s, C.ink);
}

// 18 — Achievement program
function slide18(s) {
  s.background = { color: C.cream };
  rect(s, 3.825, 0, 3.819, 3.101, C.lilac);
  rect(s, -0.006, 0, 3.832, 3.101, C.brown);

  [['Future Urban Mental Health Strategies', 0.732, 2.015, C.paper],
   ['Mental Day Urban Family', 4.551, 1.29, C.ink]].forEach(([head, x, hw, color]) => {
    text(s, head, { x, y: 0.688, w: hw, h: 0.675, fontSize: 12, color, lineSpacingMultiple: 1.5 });
    text(s, L.cut, { x, y: 1.669, w: 2.325, h: 0.832, color, align: 'justify', lineSpacingMultiple: 1.5 });
  });

  text(s, ml('Achievement', 'Program'),
    { x: 7.834, y: 0.519, w: 5.506, h: 3.13, fontFace: FONT_MED, fontSize: 60 });

  photoFrame(s, 0, 3.101, 13.333, 4.399);
}

// 19 — Powered by our team
function slide19(s) {
  s.background = { color: C.brown };

  roundRect(s, 1.869, 1.407, 4.94, 5.467, C.cream, 0.16667, 270);
  roundRect(s, 1.463, 1.407, 4.94, 5.467, C.lilac, 0.16667, 270);
  roundRect(s, 0.938, 1.407, 4.94, 5.467, C.lilac, 0.16667, 270);

  text(s, 'Powered by Our Team',
    { x: 7.911, y: 4.631, w: 5.004, h: 2.121, fontFace: THEME_MINOR, fontSize: 60, color: C.paper });
  text(s, 'We Seek Better Futures',
    { x: 8.606, y: 1.67, w: 1.405, h: 0.675, fontFace: THEME_MAJOR, fontSize: 12, color: C.paper, lineSpacingMultiple: 1.5 });
  text(s, 'Lorem ipsum dolor sit dolor sit ametis asda amet, consectetuer adipiscing elitisu. Dolor sit eget ligula eget dolor sit amet consetectuer adispiscing lorem ipsum dolor.',
    { x: 8.606, y: 2.598, w: 3.967, h: 0.832, fontFace: THEME_MINOR, color: C.paper, align: 'justify', lineSpacingMultiple: 1.5 });

  logo(s, C.paper);
  dateTag(s, C.paper);
  brandTag(s, C.paper);
}

// 20 — Thank you
function slide20(s) {
  rect(s, 0, 0, 4.415, 7.5, C.brown);

  const contact = [
    ['Website:', 'www.example.com', 7.539, 1.413],
    ['Contact:', '+11 222 3333 4444', 9.327, 1.413],
    ['Email:', 'example@email.com', 11.115, 1.555],
  ];
  contact.forEach(([label, value, x, vw]) => {
    text(s, label, { x, y: 0.652, w: 1.413, h: 0.373, fontFace: THEME_MAJOR, fontSize: 12, align: 'justify', lineSpacingMultiple: 1.5 });
    text(s, value, { x, y: 1.03, w: vw, h: 0.327, fontFace: THEME_MINOR, align: 'justify', lineSpacingMultiple: 1.5 });
  });

  text(s, 'Thank You for Your Kind Attention',
    { x: 7.539, y: 2.366, w: 5.131, h: 3.13, fontFace: THEME_MINOR, fontSize: 60 });
  text(s, 'Lorem ipsum dolor sit dolor sit ametis asda amet, consectetuer adipiscing elitisu. Dolor sit eget ligula eget dolor sit amet consetectuer dolor sit amet.',
    { x: 7.539, y: 6.32, w: 4.99, h: 0.58, fontFace: THEME_MINOR, align: 'justify', lineSpacingMultiple: 1.5 });
  text(s, 'We\u2019ll see you soon, have a great day.',
    { x: 0.76, y: 6.572, w: 2.751, h: 0.327, fontFace: THEME_MAJOR, color: C.paper, lineSpacingMultiple: 1.5 });

  roundRect(s, 1.8, 1.357, 3.7, 5.081, C.cream);
  roundRect(s, 1.269, 1.357, 3.7, 5.081, C.lilac);
  roundRect(s, 0.71, 1.357, 3.7, 5.081, C.lilac);

  logo(s, C.paper);
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */
const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
];

function build() {
  const pptx = new PptxGenJS();
  SH = pptx.ShapeType;

  pptx.defineLayout({ name: 'DECK', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'DECK';
  pptx.theme = { headFontFace: FONT_MED, bodyFontFace: FONT };

  BUILDERS.forEach(builder => {
    const slide = pptx.addSlide();
    slide.background = { color: C.paper };
    builder(slide);
  });

  return pptx.writeFile({
    fileName: path.join(__dirname, '15ede332-d468-4732-a676-49dacea93008_grok_final.pptx'),
  });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
