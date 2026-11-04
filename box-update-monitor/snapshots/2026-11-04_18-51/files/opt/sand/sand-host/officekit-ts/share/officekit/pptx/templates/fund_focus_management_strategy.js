/**
 * Fund Focus - Finance Presentation Template
 * Standalone pptxgenjs recreation of the 20-slide reference deck.
 *
 * Raster photos / logos in the source deck are replaced by flat colour
 * placeholders (see `imagePlaceholder`) - everything else is drawn with
 * native pptxgenjs shapes, text boxes and charts.
 */
const pptxgen = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ *
 * Theme
 * ------------------------------------------------------------------ */
const RED = 'C60C0C'; // accent1
const RED_MID = 'F65050'; // accent2
const RED_SOFT = 'FC7C7C'; // accent3
const RED_LIGHT = 'FE6666'; // accent4
const RED_BRIGHT = 'EE2222'; // accent5
const RED_DEEP = 'D22828'; // accent6
const RED_DARK = '940909';
const GRAD = 'E23939'; // flat stand-in for the accent1 -> accent4 gradient
const DISC_STOPS = ['CE1A1A', 'E23939', 'F45656']; // that gradient, dark -> light
const INK = '262626'; // headline grey
const BODY = '404040'; // body grey
const GREY = 'D9D9D9';
const WHITE = 'FFFFFF';

const HEAD = 'Be Vietnam Pro'; // theme major font
const TEXT = 'Open Sans Light'; // theme minor font

const W = 13.333;
const H = 7.5;

// pptxgenjs rewrites shadow objects in place while rendering, so every shape
// needs its own copy - hence functions rather than shared constants.
function shadow(opacity, blur, offset) {
  return { type: 'outer', color: '000000', opacity: opacity, blur: blur, offset: offset, angle: 90 };
}
const CARD_SHADOW = function () { return shadow(0.2, 22, 3); };
const SOFT_SHADOW = function () { return shadow(0.1, 24, 6); };
const PANEL_SHADOW = function () { return shadow(0.14, 40, 5); };
const DISC_SHADOW = function () { return shadow(0.22, 14, 2); };

/* ------------------------------------------------------------------ *
 * Small helpers
 * ------------------------------------------------------------------ */
// Text boxes in the source deck are top-anchored; pptxgenjs defaults to middle.
function txt(slide, text, o) {
  slide.addText(text, Object.assign({ valign: 'top', fontFace: TEXT, color: BODY, fontSize: 12 }, o));
}

function shape(slide, type, o) {
  slide.addShape(type, o);
}

// Diagonal gradient inside a circle: chord-length strips marching along the
// gradient axis, each rotated to lie across it. A circle is rotation-symmetric,
// so the stack fills it exactly.
function discGradient(slide, o) {
  const r = o.d / 2;
  const cx = o.x + r;
  const cy = o.y + r;
  const rad = (o.angle * Math.PI) / 180;
  const steps = Math.round(o.d * 80);
  const trim = 0.02; // keep the square strip ends just inside the rim
  slide.addShape('ellipse', {
    x: o.x, y: o.y, w: o.d, h: o.d, shadow: o.shadow,
    fill: { color: mixStops(o.stops, 0.5) }, line: { type: 'none' },
  });
  for (let i = 0; i < steps; i++) {
    const t = (i + 0.5) / steps;
    const off = (t - 0.5) * o.d;
    const half = Math.max(0, Math.sqrt(Math.max(0, r * r - off * off)) - trim);
    const th = o.d / steps + 0.01;
    slide.addShape('rect', {
      x: cx + off * Math.cos(rad) - half, y: cy + off * Math.sin(rad) - th / 2,
      w: half * 2, h: th, rotate: o.angle + 90,
      fill: { color: mixStops(o.stops, t) }, line: { type: 'none' },
    });
  }
}

function hexToRgb(c) { return [0, 2, 4].map(function (i) { return parseInt(c.substr(i, 2), 16); }); }
function rgbToHex(a) {
  return a.map(function (v) { return Math.round(v).toString(16).padStart(2, '0'); }).join('').toUpperCase();
}
function mixStops(stops, t) {
  const seg = Math.min(Math.floor(t * (stops.length - 1)), stops.length - 2);
  const local = t * (stops.length - 1) - seg;
  const a = hexToRgb(stops[seg]);
  const b = hexToRgb(stops[seg + 1]);
  return rgbToHex(a.map(function (v, k) { return v + (b[k] - v) * local; }));
}

// pptxgenjs has no gradient fill, so vertical gradients are painted as a stack
// of thin opaque strips. `radius` insets the strips along a circular arc so the
// stack keeps the silhouette of a capsule / pill.
function vGradient(slide, o) {
  const stops = o.reverse ? o.stops.slice().reverse() : o.stops;
  const r = o.radius || 0;
  const steps = o.steps || Math.round(o.h * 90);
  for (let i = 0; i < steps; i++) {
    const yMid = ((i + 0.5) / steps) * o.h;
    let inset = 0;
    if (r > 0) {
      const d = yMid < r ? r - yMid : (yMid > o.h - r ? yMid - (o.h - r) : 0);
      if (d > 0) inset = r - Math.sqrt(Math.max(0, r * r - d * d));
    }
    slide.addShape('rect', {
      x: o.x + inset, y: o.y + (o.h * i) / steps, w: Math.max(0, o.w - 2 * inset),
      h: o.h / steps + 0.006,
      fill: { color: mixStops(stops, (i + 0.5) / steps) }, line: { type: 'none' },
    });
  }
}

// Stand-in for a bitmap photo that was embedded in the reference deck.
function imagePlaceholder(slide, o) {
  slide.addShape(o.points ? 'custGeom' : (o.type || 'rect'), {
    x: o.x, y: o.y, w: o.w, h: o.h, points: o.points, rotate: o.rotate,
    rectRadius: o.rectRadius, fill: { color: o.color || 'E9E4E1' },
    line: o.line || { type: 'none' },
  });
  if (o.label !== false) {
    txt(slide, '[image]', {
      x: o.x, y: o.y + o.h / 2 - 0.18, w: o.w, h: 0.36,
      align: 'center', valign: 'middle', fontSize: 11, color: '8A8A8A',
    });
  }
}

// Red disc with a white arrow - the deck's "next" affordance.
const ARROW = [
  { x: 0.004, y: 0.136, moveTo: true },
  { x: 0.004, y: 0.116, curve: { type: 'cubic', x1: -0.001, y1: 0.13, x2: -0.001, y2: 0.121 } },
  { x: 0.024, y: 0.116, curve: { type: 'cubic', x1: 0.007, y1: 0.112, x2: 0.018, y2: 0.112 } },
  { x: 0.074, y: 0.165 }, { x: 0.074, y: 0.014 },
  { x: 0.102, y: 0.014, curve: { type: 'cubic', x1: 0.074, y1: 0.0, x2: 0.102, y2: 0.0 } },
  { x: 0.102, y: 0.165 }, { x: 0.152, y: 0.116 },
  { x: 0.172, y: 0.136, curve: { type: 'cubic', x1: 0.157, y1: 0.11, x2: 0.178, y2: 0.121 } },
  { x: 0.088, y: 0.214, curve: { type: 'cubic', x1: 0.16, y1: 0.2, x2: 0.096, y2: 0.214 } },
  { x: 0.004, y: 0.136, curve: { type: 'cubic', x1: 0.08, y1: 0.214, x2: 0.012, y2: 0.2 } },
  { close: true },
];

function arrowButton(slide, x, y, o) {
  o = o || {};
  const d = 0.527;
  const light = o.disc === WHITE;
  shape(slide, 'ellipse', {
    x: x, y: y, w: d, h: d, fill: { color: o.disc || RED }, line: { type: 'none' },
    shadow: light ? DISC_SHADOW() : CARD_SHADOW(),
  });
  shape(slide, 'custGeom', {
    x: x + 0.183, y: y + 0.166, w: 0.161, h: 0.196, rotate: 270,
    points: ARROW, fill: { color: o.arrow || WHITE }, line: { type: 'none' },
  });
}

function continueButton(slide, x, y) {
  shape(slide, 'roundRect', {
    x: x, y: y, w: 1.665, h: 0.527, rectRadius: 0.26,
    fill: { color: WHITE }, line: { color: 'F79191', width: 1 }, shadow: SOFT_SHADOW(),
  });
  txt(slide, 'Continue', { x: x + 0.245, y: y + 0.095, w: 1.185, h: 0.337, align: 'center', fontSize: 14 });
}

// The three little dots that flag a section heading.
function dots(slide, x, y) {
  [0, 0.251, 0.506].forEach(function (dx) {
    shape(slide, 'ellipse', { x: x + dx, y: y, w: 0.144, h: 0.144, fill: { color: RED_MID }, line: { type: 'none' } });
  });
}

// Simplified stand-in for the deck's small red line-art pictograms.
function iconGlyph(slide, x, y, size, color) {
  const c = color || RED;
  shape(slide, 'roundRect', {
    x: x, y: y + size * 0.12, w: size, h: size * 0.72, rectRadius: size * 0.1,
    fill: { type: 'none' }, line: { color: c, width: 1.25 },
  });
  shape(slide, 'rect', { x: x + size * 0.18, y: y + size * 0.5, w: size * 0.14, h: size * 0.26, fill: { color: c }, line: { type: 'none' } });
  shape(slide, 'rect', { x: x + size * 0.43, y: y + size * 0.34, w: size * 0.14, h: size * 0.42, fill: { color: c }, line: { type: 'none' } });
  shape(slide, 'rect', { x: x + size * 0.68, y: y + size * 0.44, w: size * 0.14, h: size * 0.32, fill: { color: c }, line: { type: 'none' } });
}

// White pictogram used inside the small coloured icon discs.
function whiteGlyph(slide, x, y, size) {
  shape(slide, 'roundRect', {
    x: x + size * 0.16, y: y + size * 0.2, w: size * 0.68, h: size * 0.6, rectRadius: size * 0.08,
    fill: { type: 'none' }, line: { color: WHITE, width: 1.5 },
  });
  shape(slide, 'rect', { x: x + size * 0.28, y: y + size * 0.42, w: size * 0.44, h: size * 0.06, fill: { color: WHITE }, line: { type: 'none' } });
  shape(slide, 'rect', { x: x + size * 0.28, y: y + size * 0.56, w: size * 0.3, h: size * 0.06, fill: { color: WHITE }, line: { type: 'none' } });
}

/* ------------------------------------------------------------------ *
 * Master chrome: divider rule, brand pill, logo, page number
 * ------------------------------------------------------------------ */
function chrome(slide, pageNo) {
  shape(slide, 'line', {
    x: 0.466, y: 0.713, w: 12.401, h: 0, line: { color: 'BFBFBF', width: 1, transparency: 70 },
  });
  shape(slide, 'roundRect', {
    x: 0.602, y: 0.307, w: 1.502, h: 0.271, rectRadius: 0.135,
    fill: { color: 'DF3535' }, line: { type: 'none' },
  });
  txt(slide, 'Fund Focus Finance -', {
    x: 0.562, y: 0.307, w: 1.583, h: 0.252, align: 'center', fontSize: 9, color: WHITE,
  });
  // brand mark: solid red disc with a white lens cut through it
  shape(slide, 'ellipse', { x: 12.501, y: 0.281, w: 0.282, h: 0.219, fill: { color: RED }, line: { type: 'none' } });
  shape(slide, 'ellipse', { x: 12.601, y: 0.291, w: 0.082, h: 0.199, fill: { color: WHITE }, line: { type: 'none' } });
  txt(slide, 'Presentation Template', {
    x: 10.989, y: 0.261, w: 1.583, h: 0.252, align: 'center', fontSize: 9,
  });
  txt(slide, String(pageNo), {
    x: 12.395, y: 6.906, w: 0.776, h: 0.303, align: 'center', fontSize: 12, color: '939393',
  });
}

/* ------------------------------------------------------------------ *
 * Reusable copy
 * ------------------------------------------------------------------ */
const LOREM_LONG = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit';
const LOREM_MED = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ';
const LOREM_BODY = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation';
const LOREM_TINY = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, ';
const LOREM_CARD = 'Lorem ipsum dolor sit amet, consectetur, adipiscing elit, sed do eiusmod tempor incididunt';

const HEADLINE = { fontFace: HEAD, fontSize: 40, color: INK, lineSpacingMultiple: 1.0 };
const KICKER = { fontFace: HEAD, fontSize: 14, color: BODY };
const PARA = { fontSize: 12, color: BODY, lineSpacingMultiple: 1.3 };

// Headline split across two paragraphs, the way the reference deck sets them.
function twoLine(a, b) {
  return [{ text: a, options: { breakLine: true } }, { text: b }];
}

/* ------------------------------------------------------------------ *
 * Slide 1 - title
 * ------------------------------------------------------------------ */
const HERO_BLOB = [
  { x: 3.475, y: 0.0, moveTo: true },
  { x: 5.932, y: 1.018, curve: { type: 'cubic', x1: 4.435, y1: 0.0, x2: 5.304, y2: 0.389 } },
  { x: 6.154, y: 1.262 }, { x: 6.154, y: 4.846 }, { x: 0.282, y: 4.846 }, { x: 0.273, y: 4.828 },
  { x: 0.0, y: 3.475, curve: { type: 'cubic', x1: 0.097, y1: 4.412, x2: 0.0, y2: 3.955 } },
  { x: 3.475, y: 0.0, curve: { type: 'cubic', x1: 0.0, y1: 1.556, x2: 1.556, y2: 0.0 } },
  { close: true },
];

function coverSlide(slide, title, dx, dy) {
  shape(slide, 'custGeom', {
    x: 7.18, y: 2.654, w: 6.154, h: 4.846, points: HERO_BLOB,
    fill: { color: GRAD }, line: { type: 'none' },
  });
  discGradient(slide, { x: 6.794, y: 1.07, d: 3.646, angle: 145, stops: DISC_STOPS, shadow: SOFT_SHADOW() });
  dots(slide, 0.926 + dx, 2.339 + dy);
  txt(slide, title, { x: 0.787 + dx, y: 2.736 + dy, w: 5.574, h: 1.195, fontFace: HEAD, fontSize: 65, bold: true, color: INK });
  txt(slide, 'Finance Presentation Template', { x: 0.867 + dx, y: 4.022 + dy, w: 5.574, h: 0.353, fontSize: 15 });
  continueButton(slide, 0.932 + dx, 4.954 + dy);
  arrowButton(slide, 2.983 + dx, 4.954 + dy);
}

function slide1(slide) {
  coverSlide(slide, 'Fund Focus', 0.12, 0.09);
}

/* ------------------------------------------------------------------ *
 * Slide 2 - section opener with a wide gradient banner
 * ------------------------------------------------------------------ */
function slide2(slide) {
  vGradient(slide, {
    x: 0.833, y: 1.39, w: 11.667, h: 2.589, radius: 1.2945,
    stops: [WHITE, 'F6BDBD', 'D34747'],
  });
  dots(slide, 1.094, 4.576);
  txt(slide, 'Core Principles of Fund Management', Object.assign({ x: 0.966, y: 4.861, w: 5.7, h: 1.447 }, HEADLINE));
  txt(slide, LOREM_LONG, Object.assign({ x: 6.793, y: 5.053, w: 5.937, h: 1.129 }, PARA));
}

/* ------------------------------------------------------------------ *
 * Slide 3 - four feature cards over a red band
 * ------------------------------------------------------------------ */
const S3_CARDS = [
  { x: 1.161, y: 3.108, title: 'Prioritizing Goals,', tx: 2.248 },
  { x: 6.858, y: 3.108, title: 'Enhancing Decision-Making,', tx: 7.936 },
  { x: 1.161, y: 4.999, title: 'Improving Resource,', tx: 2.123 },
  { x: 6.898, y: 4.999, title: 'Adapting to Change,', tx: 7.936 },
];

function slide3(slide) {
  shape(slide, 'rect', { x: 0, y: 0, w: W, h: 4.125, fill: { color: RED, transparency: 22 }, line: { type: 'none' } });
  txt(slide, 'Relevance of Strategic Focus', Object.assign({}, HEADLINE, { x: 1.093, y: 1.057, w: 5.7, h: 1.447, color: WHITE }));
  txt(slide, LOREM_BODY, Object.assign({}, PARA, { x: 6.709, y: 1.216, w: 5.454, h: 1.129, color: WHITE }));
  S3_CARDS.forEach(function (c) {
    shape(slide, 'roundRect', {
      x: c.x, y: c.y, w: 5.296, h: 1.628, rectRadius: 0.26,
      fill: { color: WHITE }, line: { type: 'none' }, shadow: CARD_SHADOW(),
    });
    iconGlyph(slide, c.x + 0.416, c.y + 0.576, 0.475);
    txt(slide, c.title, { x: c.tx, y: c.y + 0.268, w: 4.006, h: 0.37, fontFace: HEAD, fontSize: 16 });
    txt(slide, LOREM_CARD, Object.assign({}, PARA, { x: c.tx, y: c.y + 0.672, w: 4.006, h: 0.604 }));
  });
}

/* ------------------------------------------------------------------ *
 * Slide 4 - split layout with a stat chip
 * ------------------------------------------------------------------ */
function slide4(slide) {
  txt(slide, 'Introduction to Fund Allocation', Object.assign({}, HEADLINE, { x: 6.912, y: 1.132, w: 5.377, h: 1.447 }));
  txt(slide, 'Understanding the power of strategic,', Object.assign({}, KICKER, { x: 6.912, y: 2.911, w: 5.192, h: 0.337 }));
  txt(slide, LOREM_LONG, Object.assign({}, PARA, { x: 6.912, y: 3.39, w: 5.819, h: 1.129 }));
  shape(slide, 'roundRect', {
    x: 7.051, y: 5.044, w: 2.801, h: 1.323, rectRadius: 0.29,
    fill: { color: WHITE }, line: { color: RED_MID, width: 1 }, shadow: SOFT_SHADOW(),
  });
  shape(slide, 'ellipse', {
    x: 9.507, y: 4.908, w: 0.525, h: 0.525,
    fill: { color: 'E04A4A', transparency: 45 }, line: { color: RED_LIGHT, width: 1.5 },
  });
  whiteGlyph(slide, 9.567, 4.968, 0.405);
  txt(slide, '305+', { x: 7.318, y: 5.299, w: 1.596, h: 0.505, fontFace: HEAD, fontSize: 24, bold: true });
  txt(slide, 'Lorem ipsum dolor sit amet,', { x: 7.318, y: 5.822, w: 2.404, h: 0.303, fontSize: 12 });
  arrowButton(slide, 10.62, 5.447, { disc: WHITE, arrow: RED });
}

/* ------------------------------------------------------------------ *
 * Slide 5 - definition + two headline figures
 * ------------------------------------------------------------------ */
const S5_STATS = [
  { y: 1.776, label: 'Strategic Targeting,', value: '100,784,00+' },
  { y: 4.121, label: 'Enhanced Decision,', value: '202,156,00+' },
];

function slide5(slide) {
  txt(slide, twoLine('Defining ', 'Fund Focus'),
    Object.assign({}, HEADLINE, { x: 0.879, y: 1.406, w: 4.138, h: 1.447 }));
  txt(slide, 'Framework designed to provide clarity,', Object.assign({}, KICKER, { x: 0.879, y: 3.247, w: 5.192, h: 0.337 }));
  txt(slide, LOREM_MED, Object.assign({}, PARA, { x: 0.879, y: 3.717, w: 4.705, h: 1.129 }));
  continueButton(slide, 0.99, 5.407);
  arrowButton(slide, 3.241, 5.427);
  S5_STATS.forEach(function (s) {
    txt(slide, s.label, Object.assign({}, KICKER, { x: 9.185, y: s.y, w: 2.885, h: 0.337 }));
    txt(slide, s.value, { x: 9.185, y: s.y + 0.439, w: 2.795, h: 0.572, fontFace: HEAD, fontSize: 28, bold: true });
    txt(slide, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor',
      Object.assign({}, PARA, { x: 9.185, y: s.y + 1.039, w: 3.546, h: 0.604 }));
  });
}

/* ------------------------------------------------------------------ *
 * Slide 6 - centred copy flanked by two half discs
 * ------------------------------------------------------------------ */
function slide6(slide) {
  discGradient(slide, { x: 10.517, y: 0.934, d: 5.633, angle: 145, stops: DISC_STOPS });
  discGradient(slide, { x: -2.817, y: 0.934, d: 5.633, angle: 145, stops: DISC_STOPS });
  txt(slide, twoLine('Exploring ', 'Fund All Types'),
    Object.assign({}, HEADLINE, { x: 4.395, y: 1.095, w: 4.705, h: 1.447, align: 'center' }));
  [{ y: 2.882, head: 'Diverse Investment Options,' }, { y: 4.272, head: 'Maximing Opportunities,' }].forEach(function (b) {
    txt(slide, b.head, Object.assign({}, KICKER, { x: 4.071, y: b.y, w: 5.192, h: 0.38, align: 'center', lineSpacingMultiple: 1.3 }));
    txt(slide, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, ',
      Object.assign({}, PARA, { x: 3.592, y: b.y + 0.419, w: 6.149, h: 0.604, align: 'center' }));
  });
  continueButton(slide, 5.531, 5.845);
  arrowButton(slide, 7.538, 5.845);
}

/* ------------------------------------------------------------------ *
 * Slide 7 - risk / return with two tilted capsules
 * ------------------------------------------------------------------ */
function slide7(slide) {
  shape(slide, 'roundRect', {
    x: 8.118, y: 1.254, w: 1.507, h: 5.686, rectRadius: 0.75, rotate: 318,
    fill: { color: RED, transparency: 26 }, line: { type: 'none' },
  });
  shape(slide, 'roundRect', {
    x: 10.629, y: 1.344, w: 1.199, h: 2.122, rectRadius: 0.6, rotate: 318,
    fill: { color: RED, transparency: 26 }, line: { type: 'none' },
  });
  txt(slide, 'Balancing Risk and Return', Object.assign({}, HEADLINE, { x: 0.879, y: 1.406, w: 4.705, h: 1.447 }));
  txt(slide, 'Balancing risk and return is key,', Object.assign({}, KICKER, { x: 0.879, y: 3.247, w: 5.192, h: 0.337 }));
  txt(slide, LOREM_MED, Object.assign({}, PARA, { x: 0.879, y: 3.717, w: 4.705, h: 1.129 }));
  [{ x: 0.982, xr: 0.997, w: 1.488, num: '4', label: 'Trade Off-s' },
   { x: 3.267, xr: 3.278, w: 1.583, num: '9', label: 'Risk Appetite' }].forEach(function (s) {
    slide.addText([
      { text: s.num, options: { fontSize: 35, bold: true, color: BODY, fontFace: TEXT } },
      { text: '/10', options: { fontSize: 20, bold: true, color: BODY, fontFace: TEXT } },
    ], { x: s.xr, y: s.y || 5.34, w: 1.25, h: 0.69, valign: 'middle' });
    txt(slide, s.label, Object.assign({}, PARA, { x: s.x, y: 5.901, w: s.w, h: 0.345 }));
  });
  shape(slide, 'line', { x: 2.663, y: 5.442, w: 0, h: 0.746, line: { color: GREY, width: 1 } });
  arrowButton(slide, 5.424, 5.611);
}

/* ------------------------------------------------------------------ *
 * Slide 8 - diversification
 * ------------------------------------------------------------------ */
const LEFT_LOBE = [
  { x: 0.026, y: 0.0, moveTo: true },
  { x: 2.686, y: 2.66, curve: { type: 'cubic', x1: 1.495, y1: 0.0, x2: 2.686, y2: 1.191 } },
  { x: 0.026, y: 5.32, curve: { type: 'cubic', x1: 2.686, y1: 4.129, x2: 1.495, y2: 5.32 } },
  { x: 0.0, y: 5.319 }, { x: 0.0, y: 0.001 }, { close: true },
];

function slide8(slide) {
  shape(slide, 'custGeom', { x: -0.026, y: 1.493, w: 2.686, h: 5.32, points: LEFT_LOBE, fill: { color: GRAD }, line: { type: 'none' } });
  discGradient(slide, { x: 0.602, y: 1.78, d: 4.439, angle: 145, stops: DISC_STOPS });
  txt(slide, 'Power of Diversification', Object.assign({}, HEADLINE, { x: 5.677, y: 1.497, w: 6.871, h: 0.774 }));
  txt(slide, 'Diversification reduces risk and enhances portfolio,', Object.assign({}, KICKER, { x: 5.712, y: 2.825, w: 5.192, h: 0.337 }));
  txt(slide, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim',
    Object.assign({}, PARA, { x: 5.712, y: 3.361, w: 6.357, h: 0.604 }));
  txt(slide, LOREM_MED, Object.assign({}, PARA, { x: 5.712, y: 4.016, w: 6.357, h: 0.866 }));
  [{ x: 5.712, icon: 6.044, tx: 6.52, title: 'Risk Reduction,' },
   { x: 8.998, icon: 9.291, tx: 9.821, title: 'Maximized Returns' }].forEach(function (p) {
    shape(slide, 'roundRect', {
      x: p.x, y: 5.384, w: 2.938, h: 1.097, rectRadius: 0.548,
      fill: { type: 'none' }, line: { color: RED_MID, width: 1 },
    });
    iconGlyph(slide, p.icon, 5.743, 0.42);
    txt(slide, p.title, { x: p.tx, y: 5.649, w: 1.975, h: 0.28, fontFace: HEAD, fontSize: 12, lineSpacingMultiple: 1.3 });
    txt(slide, 'Lorem ipsum dolor sit amet ', { x: p.tx, y: 5.9, w: 1.975, h: 0.28, fontSize: 10, lineSpacingMultiple: 1.3 });
  });
  arrowButton(slide, 12.285, 3.702);
}

/* ------------------------------------------------------------------ *
 * Slide 9 - market dynamics
 * ------------------------------------------------------------------ */
function slide9(slide) {
  [{ x: 6.778, reverse: false }, { x: 9.667, reverse: true }].forEach(function (p) {
    vGradient(slide, {
      x: p.x, y: 1.49, w: 2.562, h: 4.854, radius: 1.281, reverse: p.reverse,
      stops: [WHITE, 'F09D9D', 'C81515'],
    });
  });
  txt(slide, 'Adapting to Market Dynamics', Object.assign({}, HEADLINE, { x: 0.879, y: 1.326, w: 5.192, h: 1.447 }));
  txt(slide, 'Stay updated with current market trends,', Object.assign({}, KICKER, { x: 0.879, y: 3.28, w: 5.192, h: 0.337 }));
  txt(slide, LOREM_SHORT, Object.assign({}, PARA, { x: 0.879, y: 3.75, w: 4.705, h: 0.866 }));
  [{ x: 0.879, title: 'One Update,' }, { x: 3.475, title: 'Two Update,' }].forEach(function (c) {
    shape(slide, 'roundRect', {
      x: c.x, y: 5.117, w: 2.221, h: 1.027, rectRadius: 0.16,
      fill: { color: WHITE }, line: { color: 'F79191', width: 2 }, shadow: SOFT_SHADOW(),
    });
    txt(slide, c.title, { x: c.x + 0.16, y: 5.177, w: 1.9, h: 0.337, fontSize: 14, bold: true, fontFace: 'Manrope', margin: 0 });
    txt(slide, '\u2611 Lorem ipsum dolor', { x: c.x + 0.16, y: 5.592, w: 2.0, h: 0.3, fontSize: 12, fontFace: 'Manrope', margin: 0 });
  });
}

/* ------------------------------------------------------------------ *
 * Slide 10 - testimonials with star ratings
 * ------------------------------------------------------------------ */
const S10_BLOCKS = [
  { y: 1.334, value: '2,306+', title: 'Diversification Matters' },
  { y: 4.147, value: '2,170+', title: 'Patience Pays Off' },
];

function slide10(slide) {
  [1.29, 4.123].forEach(function (y) {
    discGradient(slide, { x: 4.853, y: y, d: 2.48, angle: 145, stops: DISC_STOPS });
  });
  txt(slide, twoLine('Lessons ', 'from Investors'),
    Object.assign({}, HEADLINE, { x: 7.712, y: 1.216, w: 5.1, h: 1.447 }));
  txt(slide, 'Real-world stories and testimonials,', Object.assign({}, KICKER, { x: 7.712, y: 3.305, w: 4.066, h: 0.336 }));
  txt(slide, LOREM_SHORT, Object.assign({}, PARA, { x: 7.712, y: 3.775, w: 4.875, h: 0.866 }));
  continueButton(slide, 7.806, 5.382);
  arrowButton(slide, 9.811, 5.382);
  S10_BLOCKS.forEach(function (b) {
    txt(slide, b.value, { x: 1.036, y: b.y, w: 2.797, h: 0.909, fontFace: HEAD, fontSize: 48, bold: true });
    txt(slide, b.title, { x: 1.036, y: b.y + 0.72, w: 2.797, h: 0.337, fontSize: 14, bold: true });
    const tick = { bullet: { characterCode: '2713', indent: 13.5 }, breakLine: true };
    slide.addText([
      { text: 'Lorem ipsum dolor sit amet, consectetur', options: tick },
      { text: 'Lorem ipsum dolor sit amet, consectetur', options: tick },
    ], {
      x: 1.057, y: b.y + 1.094, w: 3.485, h: 0.562, valign: 'top', fontFace: TEXT, fontSize: 11,
      color: BODY, lineSpacingMultiple: 1.3, margin: 0,
    });
    for (let i = 0; i < 5; i++) {
      shape(slide, 'star5', {
        x: 1.145 + i * 0.258, y: b.y + 1.821, w: 0.198, h: 0.198,
        fill: { color: i === 4 ? GREY : RED_MID }, line: { type: 'none' },
      });
    }
  });
}

/* ------------------------------------------------------------------ *
 * Slide 11 - three feature columns with ribbon call-outs
 * ------------------------------------------------------------------ */
const S11_COLS = [
  { card: 1.517, title: 'Core Features', tx: 2.038, ribbon: 1.167, rtx: 1.589, ry: 3.978,
    above: [{ y: 3.4, h: 0.341, lines: 1 }], below: { y: 4.69, h: 1.423, lines: 3 } },
  { card: 5.33, title: 'User-Centric', tx: 5.858, ribbon: 4.993, rtx: 5.415, ry: 4.449,
    above: [{ y: 3.273, h: 0.935, lines: 2 }], below: { y: 5.16, h: 0.935, lines: 2 } },
  { card: 9.144, title: 'Operational', tx: 9.661, ribbon: 8.82, rtx: 9.216, ry: 5.023,
    above: [{ y: 3.273, h: 1.423, lines: 3 }], below: { y: 5.797, h: 0.341, lines: 1 } },
];

function bulletLines(slide, x, y, w, h, lines, spacing) {
  const rows = [];
  const opt = { bullet: { characterCode: '25BA', indent: 13.5 }, breakLine: true };
  for (let i = 0; i < lines; i++) rows.push({ text: 'Lorem ipsum dolor sit', options: opt });
  slide.addText(rows, {
    x: x, y: y, w: w, h: h, valign: 'top', align: 'center', fontFace: TEXT, fontSize: 12,
    color: BODY, lineSpacingMultiple: spacing, margin: 0,
  });
}

function slide11(slide) {
  S11_COLS.forEach(function (c) {
    shape(slide, 'roundRect', {
      x: c.card, y: 2.431, w: 3.126, h: 4.236, rectRadius: 0.3,
      fill: { color: WHITE }, line: { type: 'none' }, shadow: CARD_SHADOW(),
    });
  });
  txt(slide, 'Key The Functionalities', Object.assign({}, HEADLINE, { x: 2.965, y: 1.117, w: 7.404, h: 0.774, align: 'center' }));
  S11_COLS.forEach(function (c) {
    txt(slide, c.title, {
      x: c.tx, y: 2.706, w: 2.148, h: 0.5, align: 'center', valign: 'bottom',
      fontFace: HEAD, fontSize: 20, bold: true, lineSpacingMultiple: 1.3,
    });
    c.above.forEach(function (a) { bulletLines(slide, c.tx, a.y, 2.152, a.h, a.lines, a.lines > 1 ? 2.0 : 1.3); });
    // red ribbon with its folded corner
    shape(slide, 'rect', { x: c.ribbon, y: c.ry, w: 3.559, h: 0.585, fill: { color: RED }, line: { type: 'none' }, shadow: CARD_SHADOW() });
    shape(slide, 'rtTriangle', {
      x: c.ribbon, y: c.ry + 0.585, w: 0.347, h: 0.298, flipH: true, flipV: true,
      fill: { color: RED_DARK }, line: { type: 'none' },
    });
    txt(slide, 'Lorem ipsum dolor sit amet, ', {
      x: c.rtx, y: c.ry + 0.1, w: 2.752, h: 0.341, align: 'center', color: WHITE, lineSpacingMultiple: 1.3,
    });
    bulletLines(slide, c.tx, c.below.y, 2.152, c.below.h, c.below.lines, c.below.lines > 1 ? 2.0 : 1.3);
  });
}

/* ------------------------------------------------------------------ *
 * Slide 12 - staggered comparison cards
 * ------------------------------------------------------------------ */
const S12_CARDS = [
  { x: 2.335, y: 3.091, label: 'Targeted Investment', value: '20% - 70%', dark: false },
  { x: 5.429, y: 1.927, label: 'Expert Guidance', value: '12% - 75%', dark: true },
  { x: 8.523, y: 0.763, label: 'Tailored Strategies', value: '50% - 99%', dark: false },
];

function slide12(slide) {
  txt(slide, twoLine('Why Choose ', 'Fund Focus?'),
    Object.assign({}, HEADLINE, { x: 0.931, y: 1.079, w: 4.546, h: 1.447 }));
  S12_CARDS.forEach(function (c) {
    const fg = c.dark ? WHITE : BODY;
    shape(slide, 'roundRect', {
      x: c.x, y: c.y, w: 2.475, h: 3.646, rectRadius: 0.33,
      fill: { color: c.dark ? 'DE3232' : WHITE }, line: { color: GREY, width: 0.5 }, shadow: CARD_SHADOW(),
    });
    txt(slide, c.label, { x: c.x + 0.169, y: c.y + 0.129, w: 2.145, h: 0.303, color: fg });
    txt(slide, c.value, { x: c.x + 0.17, y: c.y + 2.886, w: 2.135, h: 0.37, fontFace: HEAD, fontSize: 16, bold: true, color: fg });
    txt(slide, 'Lorem ipsum dolor sit', { x: c.x + 0.17, y: c.y + 3.18, w: 2.135, h: 0.345, color: fg, lineSpacingMultiple: 1.3 });
  });
  // wide summary card
  shape(slide, 'roundRect', {
    x: 8.523, y: 5.107, w: 3.672, h: 1.645, rectRadius: 0.22,
    fill: { color: WHITE }, line: { color: GREY, width: 0.5 }, shadow: CARD_SHADOW(),
  });
  txt(slide, 'Diverse Opportunities,', { x: 8.78, y: 5.302, w: 2.468, h: 0.286, fontSize: 11 });
  txt(slide, '100,893,00+', { x: 8.78, y: 5.561, w: 3.622, h: 0.404, fontFace: HEAD, fontSize: 18, bold: true });
  txt(slide, 'Lorem ipsum ', { x: 8.78, y: 6.024, w: 2.468, h: 0.286, fontSize: 11 });
  txt(slide, 'Three Fund Focus', { x: 8.78, y: 6.326, w: 1.833, h: 0.303, fontFace: HEAD, fontSize: 12, bold: true });
  // floating red badge
  shape(slide, 'roundRect', {
    x: 11.089, y: 5.671, w: 1.798, h: 0.901, rectRadius: 0.18,
    fill: { color: 'D62A2A' }, line: { color: GREY, width: 0.5 }, shadow: CARD_SHADOW(),
  });
  txt(slide, 'Provides Access,', { x: 11.262, y: 5.784, w: 1.411, h: 0.286, fontSize: 11, color: WHITE, align: 'center' });
  txt(slide, '5.81%', { x: 11.43, y: 6.022, w: 1.285, h: 0.438, fontFace: HEAD, fontSize: 20, bold: true, color: WHITE, align: 'center', margin: 0 });
  shape(slide, 'triangle', { x: 11.381, y: 6.17, w: 0.137, h: 0.118, fill: { color: WHITE }, line: { type: 'none' } });
}

/* ------------------------------------------------------------------ *
 * Slide 13 - pull quote
 * ------------------------------------------------------------------ */
const QUOTE_MARK = [
  { x: 0.767, y: 0.0, moveTo: true }, { x: 0.838, y: 0.113 },
  { x: 0.707, y: 0.224, curve: { type: 'cubic', x1: 0.779, y1: 0.138, x2: 0.735, y2: 0.175 } },
  { x: 0.66, y: 0.438, curve: { type: 'cubic', x1: 0.679, y1: 0.273, x2: 0.664, y2: 0.344 } },
  { x: 0.812, y: 0.438 }, { x: 0.812, y: 0.775 }, { x: 0.501, y: 0.775 }, { x: 0.501, y: 0.509 },
  { x: 0.552, y: 0.195, curve: { type: 'cubic', x1: 0.501, y1: 0.364, x2: 0.518, y2: 0.26 } },
  { x: 0.767, y: 0.0, curve: { type: 'cubic', x1: 0.597, y1: 0.109, x2: 0.669, y2: 0.044 } },
  { close: true },
  { x: 0.266, y: 0.0, moveTo: true }, { x: 0.337, y: 0.113 },
  { x: 0.207, y: 0.224, curve: { type: 'cubic', x1: 0.278, y1: 0.138, x2: 0.235, y2: 0.175 } },
  { x: 0.16, y: 0.438, curve: { type: 'cubic', x1: 0.179, y1: 0.273, x2: 0.163, y2: 0.344 } },
  { x: 0.312, y: 0.438 }, { x: 0.312, y: 0.775 }, { x: 0.0, y: 0.775 }, { x: 0.0, y: 0.509 },
  { x: 0.052, y: 0.195, curve: { type: 'cubic', x1: 0.0, y1: 0.364, x2: 0.017, y2: 0.26 } },
  { x: 0.266, y: 0.0, curve: { type: 'cubic', x1: 0.097, y1: 0.109, x2: 0.168, y2: 0.044 } },
  { close: true },
];
const HALF_CIRCLE = [
  { x: 0.0, y: 0.0, moveTo: true },
  { x: 2.329, y: 2.329, curve: { type: 'cubic', x1: 1.286, y1: 0.0, x2: 2.329, y2: 1.043 } },
  { x: 0.0, y: 4.658, curve: { type: 'cubic', x1: 2.329, y1: 3.615, x2: 1.286, y2: 4.658 } },
  { close: true },
];

function slide13(slide) {
  imagePlaceholder(slide, { x: 0.0, y: 1.098, w: 2.329, h: 4.658, points: HALF_CIRCLE, color: 'C6CFD5' });
  shape(slide, 'custGeom', { x: 4.604, y: 1.304, w: 0.838, h: 0.775, points: QUOTE_MARK, fill: { color: 'DD3B3B' }, line: { type: 'none' } });
  shape(slide, 'custGeom', {
    x: 11.918, y: 5.392, w: 0.838, h: 0.775, points: QUOTE_MARK, rotate: 180,
    fill: { color: 'DD3B3B' }, line: { type: 'none' },
  });
  txt(slide, twoLine('Building a', 'Stable Future'),
    { x: 5.847, y: 1.537, w: 6.338, h: 1.919, fontFace: HEAD, fontSize: 54, color: INK });
  txt(slide, 'Building a successful financial portfolio requires knowledge, discipline, and the right tools. Fund Focus bridges the gap between complexity and clarity, offering investors a structured approach to achieve their goals.',
    { x: 5.963, y: 3.697, w: 5.838, h: 0.978, lineSpacingMultiple: 1.5 });
  txt(slide, 'By analyzing market trends, assessing risks, and diversifying effectively, you can secure your financial future and achieve consistent growth',
    { x: 5.963, y: 4.872, w: 5.838, h: 0.505 });
  txt(slide, 'Building a Stable Future', { x: 2.57, y: 5.124, w: 2.87, h: 0.37, align: 'right', fontFace: HEAD, fontSize: 16 });
  txt(slide, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad',
    Object.assign({}, PARA, { x: 1.426, y: 5.552, w: 4.013, h: 0.866, align: 'right' }));
  dots(slide, 4.698, 4.751);
  shape(slide, 'line', { x: 6.061, y: 6.167, w: 3.123, h: 0, line: { color: RED, width: 1 } });
  arrowButton(slide, 12.185, 3.05);
}

/* ------------------------------------------------------------------ *
 * Slide 14 - comparative analysis (native area chart)
 * ------------------------------------------------------------------ */
const MONTHS = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep'];
// The reference draws its "chart" as vector art; these series are that curve
// resampled at 30 stops so a native area chart traces the same silhouette.
const WAVE_X = Array.from({ length: 30 }, function (_, i) { return String(i); });
const WAVE_BACK = [65, 75, 88, 97, 100, 100, 98, 92, 84, 73, 63, 59, 61, 64, 66,
  68, 69, 69, 68, 67, 65, 60, 56, 50, 46, 45, 52, 59, 64, 66];
const WAVE_FRONT = [30, 34, 40, 44, 46, 46, 45, 42, 39, 34, 28, 28, 34, 40, 48,
  53, 57, 58, 58, 60, 65, 73, 83, 89, 92, 92, 91, 83, 72, 57];
const S14_STEPS = [
  { x: 1.031, title: 'Performance', n: '01' },
  { x: 5.005, title: 'Assessment', n: '02' },
  { x: 8.978, title: 'Evaluation', n: '03' },
];

function slide14(slide) {
  txt(slide, 'Comparative Analysis of Funds', Object.assign({}, HEADLINE, { x: 2.141, y: 1.092, w: 9.052, h: 0.774, align: 'center' }));
  shape(slide, 'roundRect', {
    x: 0.905, y: 2.302, w: 11.424, h: 2.794, rectRadius: 0.09,
    fill: { color: WHITE }, line: { type: 'none' }, shadow: PANEL_SHADOW(),
  });
  slide.addChart('area', [
    { name: 'Fund A', labels: WAVE_X, values: WAVE_BACK },
    { name: 'Fund B', labels: WAVE_X, values: WAVE_FRONT },
  ], {
    x: 1.68, y: 2.77, w: 9.665, h: 1.533, chartColors: ['F79191', 'F05252'], chartColorsOpacity: 85,
    showLegend: false, catAxisHidden: true, valAxisHidden: true, valAxisMaxVal: 100, valAxisMinVal: 0,
    border: { pt: 0, color: WHITE }, showValue: false,
    valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
    plotArea: { fill: { color: WHITE } }, chartArea: { fill: { color: WHITE } },
  });
  MONTHS.forEach(function (m, i) {
    txt(slide, m, { x: 1.178 + i * 1.933, y: 4.513, w: 1.237, h: 0.267, align: 'center' });
  });
  S14_STEPS.forEach(function (s) {
    shape(slide, 'ellipse', {
      x: s.x, y: 5.473, w: 0.625, h: 0.625, fill: { color: 'D62A2A' }, line: { color: GREY, width: 0.5 }, shadow: CARD_SHADOW(),
    });
    txt(slide, s.n, { x: s.x, y: 5.473, w: 0.625, h: 0.625, align: 'center', valign: 'middle', fontSize: 14, bold: true, color: WHITE });
    txt(slide, s.title, { x: s.x + 0.878, y: 5.573, w: 2.473, h: 0.425, fontFace: HEAD, fontSize: 16, bold: true, lineSpacingMultiple: 1.3 });
    txt(slide, LOREM_TINY, Object.assign({}, PARA, { x: s.x + 0.878, y: 6.096, w: 2.478, h: 0.608 }));
  });
}

/* ------------------------------------------------------------------ *
 * Slide 15 - risk vs reward dashboard (native charts)
 * ------------------------------------------------------------------ */
const SPARK_X = Array.from({ length: 30 }, function (_, i) { return String(i); });
const RISK_AREA = [36, 39, 44, 58, 75, 91, 97, 100, 100, 99, 93, 81, 64, 48, 37,
  31, 28, 28, 32, 42, 58, 72, 78, 78, 77, 74, 70, 67, 65, 64];
const REWARD_LINE = [4, 12, 20, 23, 25, 26, 30, 37, 44, 48, 50, 50, 49, 46, 37,
  28, 27, 36, 48, 57, 66, 73, 75, 75, 77, 79, 84, 89, 96, 100];
const S15_KPI = [
  { y: 4.139, value: '45,000+', tint: [RED, RED_MID] },
  { y: 5.54, value: '780+', tint: [RED_SOFT, RED_LIGHT] },
];

function slide15(slide) {
  txt(slide, 'Risk vs Reward', Object.assign({}, HEADLINE, { x: 2.141, y: 1.092, w: 9.052, h: 0.774, align: 'center' }));
  txt(slide, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea',
    Object.assign({}, PARA, { x: 1.936, y: 2.009, w: 9.456, h: 0.604, align: 'center' }));

  // right panel
  shape(slide, 'roundRect', { x: 6.033, y: 3.225, w: 6.698, h: 3.525, rectRadius: 0.11, fill: { color: WHITE }, line: { type: 'none' }, shadow: PANEL_SHADOW() });
  txt(slide, 'Reward Data', { x: 6.286, y: 3.437, w: 6.192, h: 0.417, fontFace: HEAD, fontSize: 16 });
  shape(slide, 'roundRect', { x: 6.4, y: 4.094, w: 2.425, h: 2.325, rectRadius: 0.25, fill: { type: 'none' }, line: { color: 'D3D3D3', width: 0.5 } });
  txt(slide, 'Reward Analysis', { x: 6.664, y: 4.301, w: 1.896, h: 0.339 });
  for (let i = 0; i < 5; i++) {
    shape(slide, 'line', { x: 6.693, y: 4.906 + i * 0.307, w: 1.838, h: 0, line: { color: 'EDEDED', width: 2.25 } });
  }
  slide.addChart('line', [{ name: 'Reward', labels: SPARK_X, values: REWARD_LINE }], {
    x: 6.664, y: 4.906, w: 1.847, h: 1.229, chartColors: [RED_BRIGHT], lineSize: 2.5,
    showLegend: false, catAxisHidden: true, valAxisHidden: true, valAxisMaxVal: 100, valAxisMinVal: 0,
    valGridLine: { style: 'none' }, catGridLine: { style: 'none' }, lineDataSymbol: 'none',
    border: { pt: 0, color: WHITE }, plotArea: { fill: { type: 'none' } }, chartArea: { fill: { type: 'none' } },
  });
  [[7.3, 5.467], [8.463, 4.848]].forEach(function (d) {
    shape(slide, 'ellipse', { x: d[0], y: d[1], w: 0.107, h: 0.107, fill: { color: RED_BRIGHT }, line: { type: 'none' } });
  });
  shape(slide, 'line', { x: 9.328, y: 5.256, w: 3.15, h: 0, line: { color: 'D3D3D3', width: 1, dashType: 'dash' } });
  S15_KPI.forEach(function (k) {
    shape(slide, 'roundRect', {
      x: 9.328, y: k.y, w: 0.833, h: 0.833, rectRadius: 0.42,
      fill: { color: k.tint[0] === RED ? 'DB2B2B' : 'FD7171' }, line: { type: 'none' }, shadow: CARD_SHADOW(),
    });
    whiteGlyph(slide, 9.328, k.y, 0.833);
    txt(slide, 'Lorem ipsum', { x: 10.443, y: k.y + 0.069, w: 2.035, h: 0.339 });
    txt(slide, k.value, { x: 10.443, y: k.y + 0.347, w: 2.035, h: 0.417, fontFace: HEAD, fontSize: 16 });
  });

  // left panel
  shape(slide, 'roundRect', { x: 1.099, y: 3.966, w: 4.339, h: 2.443, rectRadius: 0.08, fill: { color: WHITE }, line: { type: 'none' }, shadow: PANEL_SHADOW() });
  txt(slide, 'Risk Data', { x: 1.099, y: 3.299, w: 4.339, h: 0.417, fontFace: HEAD, fontSize: 16 });
  txt(slide, 'Monthly Analysis', { x: 1.37, y: 4.131, w: 3.797, h: 0.379, fontSize: 14 });
  for (let i = 0; i < 4; i++) {
    shape(slide, 'line', { x: 1.37, y: 4.651 + i * 0.486, w: 3.797, h: 0, line: { color: 'EDEDED', width: 1 } });
  }
  slide.addChart('area', [{ name: 'Risk', labels: SPARK_X, values: RISK_AREA }], {
    x: 1.37, y: 4.906, w: 3.797, h: 1.203, chartColors: ['F26A6A'], chartColorsOpacity: 85,
    showLegend: false, catAxisHidden: true, valAxisHidden: true, valAxisMaxVal: 100, valAxisMinVal: 0,
    valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
    border: { pt: 0, color: WHITE }, plotArea: { fill: { type: 'none' } }, chartArea: { fill: { type: 'none' } },
  });
  shape(slide, 'ellipse', { x: 3.947, y: 5.251, w: 0.229, h: 0.229, fill: { color: RED_DEEP, transparency: 80 }, line: { type: 'none' } });
  shape(slide, 'ellipse', { x: 4.013, y: 5.316, w: 0.098, h: 0.098, fill: { color: RED_BRIGHT }, line: { type: 'none' } });
  // tooltip callout
  shape(slide, 'roundRect', { x: 3.395, y: 4.494, w: 1.333, h: 0.745, rectRadius: 0.03, fill: { color: WHITE }, line: { type: 'none' }, shadow: PANEL_SHADOW() });
  shape(slide, 'triangle', { x: 3.949, y: 5.19, w: 0.228, h: 0.09, rotate: 180, fill: { color: WHITE }, line: { type: 'none' } });
  txt(slide, 'August ', { x: 3.465, y: 4.559, w: 1.193, h: 0.339, align: 'center' });
  txt(slide, '10.000', { x: 3.465, y: 4.756, w: 1.193, h: 0.417, align: 'center', fontFace: HEAD, fontSize: 16 });
}

/* ------------------------------------------------------------------ *
 * Slide 16 - step-by-step tree
 * ------------------------------------------------------------------ */
const S16_LEAVES = [
  { x: 5.859, y: 1.662, w: 1.593, h: 1.823, n: '01', color: RED },
  { x: 7.333, y: 2.644, w: 1.657, h: 1.728, n: '02', color: RED_LIGHT },
  { x: 4.26, y: 3.23, w: 1.786, h: 1.611, n: '03', color: RED_LIGHT },
  { x: 7.457, y: 4.631, w: 1.807, h: 1.599, n: '04', color: RED },
  { x: 3.907, y: 5.143, w: 1.819, h: 1.594, n: '05', color: RED },
];
const S16_LABELS = [
  { x: 7.601, y: 1.237, w: 3.004, title: 'Set Clear Goals', align: 'left' },
  { x: 9.289, y: 3.033, w: 3.124, title: 'Asssess Your Risk', align: 'left' },
  { x: 0.635, y: 3.451, w: 3.191, title: 'Research Investment', align: 'right' },
  { x: 9.84, y: 5.167, w: 2.944, title: 'Monitor Performance', align: 'left' },
  { x: 0.635, y: 5.459, w: 2.791, title: 'Create a Strategy', align: 'right' },
];

function slide16(slide) {
  // stylised grey tree: tapering trunk plus a handful of branches
  shape(slide, 'trapezoid', { x: 5.72, y: 5.9, w: 1.6, h: 1.6, fill: { color: GREY }, line: { type: 'none' } });
  shape(slide, 'trapezoid', { x: 6.12, y: 3.3, w: 0.8, h: 2.65, fill: { color: GREY }, line: { type: 'none' } });
  [{ x: 5.42, y: 3.15, w: 1.0, h: 0.2, rot: 38 }, { x: 6.62, y: 3.3, w: 0.95, h: 0.18, rot: -34 },
   { x: 5.62, y: 4.3, w: 0.95, h: 0.18, rot: 28 }, { x: 6.55, y: 4.5, w: 1.0, h: 0.18, rot: -24 },
   { x: 6.2, y: 2.75, w: 0.62, h: 0.16, rot: 10 }, { x: 5.95, y: 2.6, w: 0.5, h: 0.14, rot: 62 }]
    .forEach(function (b) {
      shape(slide, 'rect', { x: b.x, y: b.y, w: b.w, h: b.h, rotate: b.rot, fill: { color: GREY }, line: { type: 'none' } });
    });
  txt(slide, 'Step-by-Step Investment', Object.assign({}, HEADLINE, { x: 1.146, y: 1.012, w: 4.235, h: 1.447 }));
  S16_LEAVES.forEach(function (l) {
    shape(slide, 'ellipse', {
      x: l.x, y: l.y, w: l.w, h: l.h, fill: { color: l.color },
      line: { color: WHITE, width: 6 }, shadow: CARD_SHADOW(),
    });
    txt(slide, l.n, {
      x: l.x, y: l.y, w: l.w, h: l.h, align: 'center', valign: 'middle',
      fontFace: HEAD, fontSize: 28, bold: true, color: WHITE,
    });
  });
  S16_LABELS.forEach(function (t) {
    txt(slide, t.title, { x: t.x, y: t.y, w: t.w, h: 0.37, align: t.align, fontFace: HEAD, fontSize: 16, bold: true });
    txt(slide, 'Lorem ipsum dolor sit amet', Object.assign({}, PARA, { x: t.x, y: t.y + 0.33, w: t.w - 0.2, h: 0.341, align: t.align }));
  });
}

/* ------------------------------------------------------------------ *
 * Slide 17 - clock diagram
 * ------------------------------------------------------------------ */
const CLOCK_HANDS = [
  { x: 0.205, y: 1.346, moveTo: true }, { x: 0.455, y: 1.859 },
  { x: 0.575, y: 1.901, curve: { type: 'cubic', x1: 0.477, y1: 1.904, x2: 0.53, y2: 1.922 } },
  { x: 0.616, y: 1.781, curve: { type: 'cubic', x1: 0.619, y1: 1.879, x2: 0.637, y2: 1.826 } },
  { x: 0.365, y: 1.267 },
  { x: 0.335, y: 0.985, curve: { type: 'cubic', x1: 0.431, y1: 1.181, x2: 0.422, y2: 1.057 } },
  { x: 0.756, y: 0.129 },
  { x: 0.716, y: 0.009, curve: { type: 'cubic', x1: 0.778, y1: 0.085, x2: 0.76, y2: 0.031 } },
  { x: 0.596, y: 0.05, curve: { type: 'cubic', x1: 0.671, y1: -0.013, x2: 0.618, y2: 0.006 } },
  { x: 0.157, y: 0.943 },
  { x: 0.205, y: 1.346, curve: { type: 'cubic', x1: -0.079, y1: 0.999, x2: -0.038, y2: 1.347 } },
  { close: true },
];
const S17_ARCS = [
  // quarter arcs riding on the ring, drawn as thick rotated blockArc segments
  { x: 4.085, y: 2.123, w: 1.542, h: 0.932, color: '630606',
    points: [{ x: 0.0, y: 0.871, moveTo: true },
      { x: 0.637, y: 0.233, curve: { type: 'cubic', x1: 0.156, y1: 0.601, x2: 0.377, y2: 0.384 } },
      { x: 1.508, y: 0.0, curve: { type: 'cubic', x1: 0.898, y1: 0.083, x2: 1.197, y2: 0.0 } },
      { x: 1.542, y: 0.063 }, { x: 1.508, y: 0.123 },
      { x: 0.699, y: 0.34, curve: { type: 'cubic', x1: 1.219, y1: 0.123, x2: 0.941, y2: 0.2 } },
      { x: 0.106, y: 0.932, curve: { type: 'cubic', x1: 0.457, y1: 0.479, x2: 0.251, y2: 0.682 } },
      { x: 0.0, y: 0.871 }, { close: true }] },
  { x: 3.852, y: 3.402, w: 0.34, h: 1.333, color: RED_DARK,
    points: [{ x: 0.233, y: 1.333, moveTo: true },
      { x: 0.059, y: 0.913, curve: { type: 'cubic', x1: 0.157, y1: 1.2, x2: 0.099, y2: 1.059 } },
      { x: 0.0, y: 0.462, curve: { type: 'cubic', x1: 0.02, y1: 0.766, x2: 0.0, y2: 0.615 } },
      { x: 0.059, y: 0.011, curve: { type: 'cubic', x1: 0.0, y1: 0.3, x2: 0.025, y2: 0.13 } },
      { x: 0.126, y: 0.0 }, { x: 0.178, y: 0.043 },
      { x: 0.123, y: 0.462, curve: { type: 'cubic', x1: 0.146, y1: 0.181, x2: 0.123, y2: 0.32 } },
      { x: 0.34, y: 1.271, curve: { type: 'cubic', x1: 0.123, y1: 0.745, x2: 0.269, y2: 1.148 } },
      { x: 0.233, y: 1.333 }, { close: true }] },
  { x: 6.403, y: 2.633, w: 0.932, h: 2.74, color: RED_MID,
    points: [{ x: 0.422, y: 0.0, moveTo: true },
      { x: 0.932, y: 1.231, curve: { type: 'cubic', x1: 0.713, y1: 0.356, x2: 0.886, y2: 0.775 } },
      { x: 0.061, y: 2.74, curve: { type: 'cubic', x1: 0.932, y1: 1.842, x2: 0.548, y2: 2.363 } },
      { x: 0.006, y: 2.708 }, { x: 0.0, y: 2.634 },
      { x: 0.809, y: 1.231, curve: { type: 'cubic', x1: 0.453, y1: 2.283, x2: 0.809, y2: 1.521 } },
      { x: 0.335, y: 0.087, curve: { type: 'cubic', x1: 0.809, y1: 0.806, x2: 0.487, y2: 0.239 } },
      { x: 0.422, y: 0.0 }, { close: true }] },
  { x: 4.362, y: 5.009, w: 1.268, h: 0.597, color: RED,
    points: [{ x: 1.268, y: 0.597, moveTo: true },
      { x: 0.565, y: 0.464, curve: { type: 'cubic', x1: 1.243, y1: 0.597, x2: 0.775, y2: 0.551 } },
      { x: 0.0, y: 0.087, curve: { type: 'cubic', x1: 0.356, y1: 0.377, x2: 0.163, y2: 0.25 } },
      { x: 0.02, y: 0.024 }, { x: 0.087, y: 0.0 },
      { x: 0.612, y: 0.351, curve: { type: 'cubic', x1: 0.239, y1: 0.152, x2: 0.418, y2: 0.271 } },
      { x: 1.268, y: 0.474, curve: { type: 'cubic', x1: 0.806, y1: 0.432, x2: 1.016, y2: 0.474 } },
      { x: 1.268, y: 0.597 }, { close: true }] },
];
const S17_ICONS = [
  { x: 4.406, y: 2.098, color: RED_DARK },
  { x: 6.952, y: 3.523, color: RED_SOFT },
  { x: 4.673, y: 5.105, color: RED_MID },
  { x: 3.617, y: 3.783, color: RED },
];
const S17_LABELS = [
  { x: 1.236, y: 1.28, tx: 1.51, w: 2.477, title: 'Risk Reduction', align: 'right' },
  { x: 0.752, y: 3.897, tx: 1.027, w: 2.477, title: 'Stability Returns', align: 'right' },
  { x: 1.956, y: 5.843, tx: 2.231, w: 2.477, title: 'Opportunities', align: 'right' },
  { x: 8.72, y: 3.16, tx: 8.72, w: 2.88, title: 'Adaptability Market', align: 'left' },
];
const S17_ARROWS = [
  { x: 3.05, y: 1.5, w: 0.6, h: 0.45, color: RED_DARK, rot: 30 },
  { x: 2.35, y: 3.45, w: 0.55, h: 0.28, color: RED, rot: 25 },
  { x: 3.0, y: 5.52, w: 0.55, h: 0.1, color: RED_MID, rot: 0 },
  { x: 7.75, y: 3.3, w: 0.6, h: 0.32, color: RED_SOFT, rot: -155 },
];

function slide17(slide) {
  shape(slide, 'ellipse', { x: 3.774, y: 2.045, w: 3.639, h: 3.639, fill: { color: 'FECBCB' }, line: { type: 'none' } });
  shape(slide, 'ellipse', { x: 4.044, y: 2.315, w: 3.099, h: 3.099, fill: { color: WHITE }, line: { type: 'none' } });
  S17_ARCS.forEach(function (a) {
    shape(slide, 'custGeom', { x: a.x, y: a.y, w: a.w, h: a.h, points: a.points, fill: { color: a.color }, line: { type: 'none' } });
  });
  shape(slide, 'custGeom', { x: 5.392, y: 2.723, w: 0.765, h: 1.91, points: CLOCK_HANDS, fill: { color: RED }, line: { type: 'none' } });
  shape(slide, 'ellipse', { x: 5.42, y: 3.67, w: 0.4, h: 0.4, fill: { color: RED }, line: { type: 'none' } });
  S17_ICONS.forEach(function (i) {
    shape(slide, 'ellipse', { x: i.x, y: i.y, w: 0.646, h: 0.646, fill: { color: i.color }, line: { type: 'none' } });
    whiteGlyph(slide, i.x, i.y, 0.646);
  });
  S17_ARROWS.forEach(function (a) {
    shape(slide, 'line', {
      x: a.x, y: a.y, w: a.w, h: a.h, rotate: a.rot,
      line: { color: a.color, width: 2, dashType: 'sysDot', endArrowType: 'triangle' },
    });
  });
  txt(slide, 'Diversification Strategy', Object.assign({}, HEADLINE, { x: 7.301, y: 0.963, w: 4.622, h: 1.447, align: 'right' }));
  S17_LABELS.forEach(function (l) {
    txt(slide, l.title, { x: l.tx, y: l.y, w: l.w - (l.align === 'right' ? 0.274 : 0), h: 0.426, align: l.align, fontFace: HEAD, fontSize: 16, bold: true, lineSpacingMultiple: 1.3 });
    txt(slide, LOREM_TINY, Object.assign({}, PARA, { x: l.x, y: l.y + 0.385, w: l.align === 'right' ? l.w : 2.468, h: 0.604, align: l.align }));
  });
  shape(slide, 'roundRect', {
    x: 7.69, y: 4.962, w: 4.547, h: 1.288, rectRadius: 0.2,
    fill: { color: WHITE }, line: { color: GREY, width: 0.5 }, shadow: CARD_SHADOW(),
  });
  shape(slide, 'ellipse', { x: 7.947, y: 5.293, w: 0.625, h: 0.625, fill: { color: 'D62A2A' }, line: { type: 'none' }, shadow: CARD_SHADOW() });
  whiteGlyph(slide, 7.947, 5.293, 0.625);
  txt(slide, 'Diversificaion Strategy,', { x: 8.793, y: 5.266, w: 2.675, h: 0.303 });
  slide.addText([
    { text: '45% - 70%  ', options: { fontSize: 20, bold: true, fontFace: HEAD, color: BODY } },
    { text: '(Four Strategy)', options: { fontSize: 12, fontFace: HEAD, color: BODY } },
  ], { x: 8.793, y: 5.568, w: 3.202, h: 0.438, valign: 'top' });
}

/* ------------------------------------------------------------------ *
 * Slide 18 - market trends, illustration + list
 * ------------------------------------------------------------------ */
const S18_ITEMS = [
  { y: 2.223, title: 'Staying Competitive' },
  { y: 3.812, title: 'Identifying' },
  { y: 5.4, title: 'Enhancing Relevance' },
];

// Flat-illustration stand-in for the reference artwork: cream blob, open
// carton spilling paper, coins and sparkle accents.
function trendIllustration(slide, ox, oy) {
  shape(slide, 'ellipse', { x: ox - 0.1, y: oy + 0.15, w: 4.4, h: 3.8, rotate: -18, fill: { color: 'FFF6EF' }, line: { type: 'none' } });
  // papers drifting out of the box
  shape(slide, 'parallelogram', { x: ox + 1.85, y: oy + 0.55, w: 1.9, h: 0.75, rotate: -6, fill: { color: 'E2E7FC' }, line: { type: 'none' } });
  shape(slide, 'parallelogram', { x: ox + 2.4, y: oy + 1.35, w: 1.35, h: 0.7, rotate: 8, flipH: true, fill: { color: 'BBC7EA' }, line: { type: 'none' } });
  shape(slide, 'parallelogram', { x: ox + 1.15, y: oy + 1.4, w: 1.75, h: 0.66, fill: { color: 'E2E7FC' }, line: { type: 'none' } });
  // the carton: back wall, two front flaps
  shape(slide, 'trapezoid', { x: ox + 0.75, y: oy + 1.8, w: 2.9, h: 1.7, fill: { color: 'F9B282' }, line: { type: 'none' } });
  shape(slide, 'parallelogram', { x: ox + 0.75, y: oy + 1.95, w: 1.1, h: 1.5, flipH: true, fill: { color: 'D88A50' }, line: { type: 'none' } });
  shape(slide, 'parallelogram', { x: ox + 2.2, y: oy + 1.95, w: 1.45, h: 1.5, fill: { color: 'D88A50' }, line: { type: 'none' } });
  shape(slide, 'diamond', { x: ox + 1.15, y: oy + 2.25, w: 0.75, h: 0.8, fill: { color: 'E2E7FC' }, line: { type: 'none' } });
  shape(slide, 'rect', { x: ox + 2.3, y: oy + 2.35, w: 0.6, h: 0.42, rotate: -3, fill: { color: 'E2E7FC' }, line: { type: 'none' } });
  // pencil and stacked notes
  shape(slide, 'rect', { x: ox + 3.44, y: oy + 2.15, w: 0.14, h: 1.3, fill: { color: 'FFCC66' }, line: { type: 'none' } });
  shape(slide, 'triangle', { x: ox + 3.42, y: oy + 2.02, w: 0.18, h: 0.2, fill: { color: '52ABAD' }, line: { type: 'none' } });
  [3.18, 3.42].forEach(function (yy) {
    shape(slide, 'roundRect', { x: ox + 1.75, y: oy + yy, w: 0.78, h: 0.14, rectRadius: 0.07, fill: { color: 'FFCC66' }, line: { type: 'none' } });
  });
  // coins
  [[0.28, 1.42, 0.46], [2.5, 3.2, 0.5]].forEach(function (c) {
    shape(slide, 'ellipse', { x: ox + c[0], y: oy + c[1], w: c[2], h: c[2], fill: { color: 'FBC139' }, line: { type: 'none' } });
    shape(slide, 'ellipse', {
      x: ox + c[0] + c[2] * 0.15, y: oy + c[1] + c[2] * 0.15, w: c[2] * 0.7, h: c[2] * 0.7,
      fill: { type: 'none' }, line: { color: 'F4AD58', width: 1.5 },
    });
  });
  // teal leaf swooshes
  shape(slide, 'moon', { x: ox + 1.7, y: oy + 0.2, w: 0.5, h: 0.55, rotate: 205, fill: { color: '66D4CF' }, line: { type: 'none' } });
  shape(slide, 'moon', { x: ox + 2.55, y: oy + 0.85, w: 0.42, h: 0.42, rotate: 25, fill: { color: '66D4CF' }, line: { type: 'none' } });
  // sparkles and outlined bubbles
  [[1.95, 0.5, 0.4], [0.45, 1.95, 0.34], [3.05, 1.0, 0.32], [0.35, 2.55, 0.26]].forEach(function (k) {
    shape(slide, 'star4', { x: ox + k[0], y: oy + k[1], w: k[2], h: k[2], fill: { color: 'FBC139' }, line: { type: 'none' } });
  });
  [[1.05, 0.25, 0.16], [0.78, 1.28, 0.2], [2.62, 1.22, 0.17], [3.18, 0.28, 0.14]].forEach(function (c) {
    shape(slide, 'ellipse', {
      x: ox + c[0], y: oy + c[1], w: c[2], h: c[2], fill: { type: 'none' }, line: { color: '52ABAD', width: 1.25 },
    });
  });
}

function slide18(slide) {
  trendIllustration(slide, 1.34, 3.007);
  txt(slide, 'Leveraging Market Trends', Object.assign({}, HEADLINE, { x: 1.269, y: 1.18, w: 5.22, h: 1.447 }));
  dots(slide, 7.204, 1.558);
  S18_ITEMS.forEach(function (it, i) {
    txt(slide, it.title, { x: 7.127, y: it.y, w: 4.432, h: 0.404, fontFace: HEAD, fontSize: 18, bold: true });
    txt(slide, LOREM_TINY, Object.assign({}, PARA, { x: 7.127, y: it.y + 0.458, w: 4.432, h: 0.341 }));
    if (i < 2) shape(slide, 'line', { x: 7.127, y: 3.55 + i * 1.452, w: 4.319, h: 0, line: { color: 'BFBFBF', width: 1 } });
  });
  arrowButton(slide, 12.087, 3.608, { disc: WHITE, arrow: RED });
}

/* ------------------------------------------------------------------ *
 * Slide 19 - portfolio dashboard on device mock-ups
 * ------------------------------------------------------------------ */
const S19_BARS = [
  { y: 5.471, value: '90', frac: 0.868 },
  { y: 5.978, value: '85', frac: 0.812 },
];

// Blank phone mock-up standing in for the device photo in the reference.
function phoneMock(slide, x, y, w, h, rot) {
  shape(slide, 'roundRect', {
    x: x, y: y, w: w, h: h, rotate: rot, rectRadius: 0.42,
    fill: { color: '7FA8C9' }, line: { type: 'none' },
  });
  shape(slide, 'roundRect', {
    x: x + 0.03, y: y + 0.03, w: w - 0.06, h: h - 0.06, rotate: rot, rectRadius: 0.4,
    fill: { color: '111C24' }, line: { type: 'none' },
  });
  shape(slide, 'roundRect', {
    x: x + 0.11, y: y + 0.11, w: w - 0.22, h: h - 0.22, rotate: rot, rectRadius: 0.33,
    fill: { color: WHITE }, line: { type: 'none' },
  });
  // the notch must orbit the phone centre so it stays at the top when tilted
  const rad = (rot * Math.PI) / 180;
  const dx = 0;
  const dy = -(h / 2 - 0.22);
  shape(slide, 'roundRect', {
    x: x + w / 2 + dx * Math.cos(rad) - dy * Math.sin(rad) - 0.45,
    y: y + h / 2 + dx * Math.sin(rad) + dy * Math.cos(rad) - 0.11,
    w: 0.9, h: 0.22, rotate: rot, rectRadius: 0.11,
    fill: { color: '111C24' }, line: { type: 'none' },
  });
}

function slide19(slide) {
  [{ x: 7.256, y: 1.722, rot: 0 }, { x: 9.094, y: 1.61, rot: 15 }].forEach(function (p) {
    phoneMock(slide, p.x, p.y, 2.482, 4.965, p.rot);
  });
  txt(slide, twoLine('Visualizing ', 'Your Portfolio'),
    Object.assign({}, HEADLINE, { x: 1.189, y: 1.101, w: 5.22, h: 1.447 }));
  txt(slide, 'Visual representation of a user-friendly dashboard,', Object.assign({}, KICKER, { x: 1.189, y: 2.874, w: 5.478, h: 0.337 }));
  shape(slide, 'roundRect', {
    x: 1.976, y: 3.537, w: 3.956, h: 3.26, rectRadius: 0.24,
    fill: { color: WHITE }, line: { color: GREY, width: 0.5 }, shadow: CARD_SHADOW(),
  });
  txt(slide, '259+', { x: 2.238, y: 3.662, w: 3.426, h: 0.819, fontFace: HEAD, fontSize: 36, bold: true, lineSpacingMultiple: 1.3 });
  txt(slide, 'Visualizing', { x: 2.238, y: 4.442, w: 3.426, h: 0.34, fontFace: HEAD, fontSize: 12, lineSpacingMultiple: 1.3 });
  txt(slide, 'Lorem ipsum dolor sit amet, consectetur', { x: 2.238, y: 4.867, w: 3.432, h: 0.321, fontSize: 11, lineSpacingMultiple: 1.3 });
  S19_BARS.forEach(function (b) {
    txt(slide, 'Stats.', { x: 2.238, y: b.y, w: 2.21, h: 0.341, valign: 'bottom' });
    txt(slide, b.value, { x: 4.798, y: b.y, w: 0.873, h: 0.341, align: 'right', bold: true, valign: 'bottom' });
    shape(slide, 'line', { x: 2.346, y: b.y + 0.4, w: 3.209, h: 0, line: { color: 'DDDDDD', width: 5 } });
    shape(slide, 'line', { x: 2.346, y: b.y + 0.4, w: 3.209 * b.frac, h: 0, line: { color: 'E33F3F', width: 5 } });
  });
  arrowButton(slide, 1.189, 4.626, { disc: WHITE, arrow: RED });
}

/* ------------------------------------------------------------------ *
 * Slide 20 - closing
 * ------------------------------------------------------------------ */
function slide20(slide) {
  coverSlide(slide, 'Thank You', 0, 0);
}

/* ------------------------------------------------------------------ *
 * Assemble
 * ------------------------------------------------------------------ */
const BUILDERS = [slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20];

function build() {
  const pres = new pptxgen();
  pres.defineLayout({ name: 'WIDE', width: W, height: H });
  pres.layout = 'WIDE';
  pres.theme = { headFontFace: HEAD, bodyFontFace: TEXT };
  pres.author = 'Fund Focus';
  pres.title = 'Fund Focus - Finance Presentation Template';

  BUILDERS.forEach(function (builder, i) {
    const slide = pres.addSlide();
    slide.background = { color: WHITE };
    builder(slide);
    chrome(slide, i + 1);
  });

  return pres.writeFile({ fileName: path.join(__dirname, '0f1600d2-23c5-4970-8381-81880542f2f8_grok_final.pptx') });
}

build().then(function (f) { console.log('wrote', f); });
