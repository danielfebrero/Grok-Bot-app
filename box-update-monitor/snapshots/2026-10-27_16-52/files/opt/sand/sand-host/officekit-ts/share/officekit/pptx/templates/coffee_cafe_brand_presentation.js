/**
 * "cafea" coffee-brand deck — recreated with pptxgenjs.
 *
 * 30 slides, 13.333 x 7.5 in (16:9). Photographs in the original deck are
 * replaced by flat grey placeholder rectangles (see `photo`).
 *
 * Run: node 041c4842-f6c1-4ecb-badc-047c1f5e5eaa_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const C = {
  white: 'FFFFFF',
  brown: '443726',        // accent2 - the deck's dark brand brown
  tan: 'C5B197',          // accent2 lum 40/60 - hairline frames & labels
  sand: 'E6D3BD',         // accent6 - timeline rule
  cocoa: '854E2D',        // accent3
  clay: 'CD8F6A',         // accent3 lum 60/40
  mocha: '8C6A4B',        // accent5 lum 75 - donut arcs
  ink: '1D211F',          // accent1 - near-black
  camel: 'B18E6D',        // accent5
  mist: 'F3F6F5',         // lt2 - world map on slide 12
  linen: 'EFE8E2',        // world map on slide 26
  grey: '6F6F6F',         // stand-in for photography
  axis: 'BFBFBF',
  axisDark: '595959',
  grid: 'D9D9D9',
  gridFaint: 'F2F2F2',
  markerTan: 'C3A589',    // accent4 lum 60/40 - map pins
};

const F = { head: 'Roboto', body: 'Roboto Condensed' };

/* Body copy. The original deck is set entirely in lorem ipsum. */
const L = {
  a: 'Proin gravida facilisis purus, at iaculis orci convallis eu. Aenean pellentesque auctor libero, a ornare justo eleifend eget. Ut consectetur venenatis feugiat. Praesent vitae mauris orci. Mauris sit amet sem vehicula, mattis mi eu, congue augue. ',
  b: 'Phasellus massa odio, pellentesque at nunc id, aliquam tempus leo. Donec ultricies ex nec lorem rutrum scelerisque. Quisque vel augue a erat venenatis vestibulum. ',
  c: 'Phasellus massa odio, pellentesque at nunc id, aliquam tempus leo. Donec ultricies ex nec lorem rutrum scelerisque. ',
  d: 'Proin gravida facilisis purus, at iaculis orci convallis eu. Aenean pellentesque auctor libero, a ornare justo eleifend eget. Ut consectetur venenatis feugiat. ',
  e: 'Proin gravida facilisis purus, at iaculis orci convallis eu. Aenean pellentesque auctor libero, a ornare justo eleifend eget. Ut consectetur venenatis.',
  f: 'Proin gravida facilisis purus, at iaculis orci convallis eu. Aenean pellentesque auctor libero, a ornare justo eleifend eget. Ut consectetur venenatis feugiat.',
  g: 'Proin gravida facilisis purus, at iaculis orci convallis eu. Aenean pellentesque auctor libero, a ornare justo eleifend eget. Ut consectetur venenatis feugiat. Praesent vitae mauris orci. ',
  h: 'Proin gravida facilisis purus, at iaculis orci convallis eu. Aenean pellentesque auctor libero, a ornare justo eleifend eget. ',
  i: 'Proin gravida facilisis purus, at iaculis orci convallis eu. ',
  j: 'Phasellus massa odio, pellentesque at nunc id, aliquam tempus leo. Donec ultricies ex nec lorem rutrum.',
};

const HEADLINE = 'Bring you simply the best coffee in the world.';
const BEANS = 'The best coffee beans from all over the world.';
const QUALITY = 'Quality ingredients. Sustainably sourced.';

/* Text-box insets, in points, as [left, right, bottom, top]. */
const M_LABEL = [14.4, 21.6, 7.2, 7.2];  // pill labels: 0.20 / 0.30 / 0.10 / 0.10 in
const M_PANEL = [21.6, 21.6, 3.6, 3.6];  // framed paragraphs: 0.30 / 0.30 in sides

/* --------------------------------------------------------------- helpers */

const NO_LINE = { type: 'none' };
const NO_FILL = { type: 'none' };

/** Plain text block. Everything the deck draws is a rectangle with text. */
function text(s, o) {
  s.addText(o.text, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: o.font || F.body,
    fontSize: o.size || 12,
    color: o.color || C.brown,
    bold: !!o.bold,
    align: o.align || 'left',
    valign: o.valign || 'top',
    charSpacing: o.spc,
    lineSpacingMultiple: o.lnSpc,
    paraSpaceBefore: o.spcBef,
    margin: o.margin || [7.2, 7.2, 3.6, 3.6],
    wrap: o.wrap !== false,
    fill: o.fill ? { color: o.fill } : undefined,
    line: o.line || NO_LINE,
  });
}

/** 32 pt (or 48 pt) brand headline. */
function heading(s, x, y, w, h, str, color, size) {
  text(s, { x, y, w, h, text: str, font: F.head, size: size || 32, bold: true, color });
}

/** Running paragraph(s): 12 pt, 120 % leading, 12 pt space-before. */
function body(s, x, y, w, h, paras, color, extra) {
  const runs = [].concat(paras).map((t, i, all) => ({
    text: t, options: { breakLine: i < all.length - 1 },
  }));
  text(s, Object.assign({
    x, y, w, h, text: runs, size: 12, color: color || C.brown,
    lnSpc: 1.2, spcBef: 12,
  }, extra || {}));
}

/** Grey stand-in for a photograph, captioned unless `bare` is set. */
function photo(s, x, y, w, h, bare) {
  s.addShape('rect', { x, y, w, h, fill: { color: C.grey }, line: NO_LINE });
  if (!bare) {
    text(s, {
      x, y: y + h / 2 - 0.2, w, h: 0.4, text: '[image]',
      size: 11, color: '9C9C9C', align: 'center', valign: 'middle',
    });
  }
}

/* ------------------------------------------------------- brand furniture */

/**
 * The "cafea | Est. 2019" wordmark: a wide framed box holding the name plus a
 * narrow framed box holding the founding year. `scale` = 1 is 2.22 in wide.
 */
function wordmark(s, x, y, color, scale) {
  const k = scale || 1;
  const w = 2.222 * k, h = 0.858 * k, w2 = 1.181 * k, gap = 0.132 * k;
  const frame = { color, width: 4.5 * k };
  text(s, {
    x, y, w, h, text: 'cafea', font: F.head, size: 32 * k, bold: true,
    color, spc: 6, align: 'center', valign: 'middle', line: frame, wrap: false,
  });
  s.addShape('rect', { x: x + w - gap, y, w: w2, h, fill: NO_FILL, line: frame });
  text(s, {
    x: x + w + 0.132 * k, y: y + 0.143 * k, w: 0.917 * k, h: 0.572 * k,
    text: [{ text: 'Est.', options: { breakLine: true } }, { text: '2019' }],
    size: 14 * k, color, spc: 6,
  });
}

/** Just the "cafea" frame, no year box (slides 3, 4, 18, 21, 24). */
function logotype(s, x, y, color) {
  text(s, {
    x, y, w: 1.691, h: 0.668, text: 'cafea', font: F.head, size: 20, bold: true,
    color, spc: 4.5, align: 'center', valign: 'middle',
    line: { color, width: 3.5 }, wrap: false,
  });
}

/* --------------------------------------------------------- feature icons */

/** Line-art pictograms, drawn inside a `size` square at (x, y). */
function icon(s, kind, x, y, size, color, bg) {
  const R = (dx, dy, dw, dh, shape, opts) => s.addShape(shape || 'rect', Object.assign({
    x: x + dx * size, y: y + dy * size, w: dw * size, h: dh * size,
    fill: { color }, line: NO_LINE,
  }, opts || {}));

  if (kind === 'chat') {                                   // two speech bubbles
    R(0.40, 0.06, 0.58, 0.52, 'rect');                     // back bubble
    R(0.66, 0.50, 0.20, 0.22, 'triangle', { rotate: 180 });
    R(0.02, 0.16, 0.62, 0.48, 'rect');                     // front bubble
    R(0.10, 0.58, 0.18, 0.20, 'triangle', { rotate: 195 });
    R(0.13, 0.27, 0.40, 0.08, 'rect', { fill: { color: bg } });
    R(0.13, 0.44, 0.40, 0.08, 'rect', { fill: { color: bg } });
  } else if (kind === 'glasses') {                         // reading glasses
    R(0.06, 0.06, 0.19, 0.10, 'rect');                     // left temple tip
    R(0.15, 0.10, 0.10, 0.42, 'rect');                     // left temple arm
    R(0.75, 0.06, 0.19, 0.10, 'rect');                     // right temple tip
    R(0.75, 0.10, 0.10, 0.42, 'rect');                     // right temple arm
    R(0.42, 0.62, 0.16, 0.09, 'rect');                     // bridge
    R(0.05, 0.50, 0.40, 0.40, 'donut', { arcThicknessRatio: 0.26 });
    R(0.55, 0.50, 0.40, 0.40, 'donut', { arcThicknessRatio: 0.26 });
  } else if (kind === 'cup') {                             // cup with steam
    R(0.28, 0.02, 0.07, 0.22, 'rect', { rotate: 20 });
    R(0.46, 0.02, 0.07, 0.22, 'rect', { rotate: 20 });
    R(0.68, 0.42, 0.26, 0.26, 'donut', { arcThicknessRatio: 0.34 });
    R(0.08, 0.36, 0.66, 0.42, 'trapezoid', { rotate: 180 });
    R(0.02, 0.80, 0.80, 0.10, 'roundRect', { rectRadius: 0.03 * size });
  } else if (kind === 'bars') {                            // stacked to-go cup
    R(0.13, 0.08, 0.74, 0.20, 'roundRect', { rectRadius: 0.05 * size });
    R(0.17, 0.38, 0.66, 0.19, 'rect');
    R(0.21, 0.67, 0.58, 0.19, 'rect');
  }
}

/* ------------------------------------------------------- composed blocks */

/**
 * "Feature card": a square icon frame, a label frame beside it, and (when
 * `bodyW` is given) a framed paragraph underneath. This is the deck's single
 * most reused component.
 */
function card(s, o) {
  const x = o.x, y = o.y, c = o.color, bg = o.bg || C.white;
  const frame = { color: c, width: 2.25 };
  const fill = o.fill ? { color: o.fill } : NO_FILL;
  s.addShape('rect', { x, y, w: 0.567, h: 0.567, fill, line: frame });
  icon(s, o.icon, x + 0.128, y + 0.128, 0.31, c, o.iconBg || o.fill || bg);
  text(s, {
    x: x + 0.571, y, w: o.labelW || 2.214, h: 0.567, text: o.label,
    font: F.head, size: 16, bold: true, color: c, align: 'center',
    valign: 'middle', margin: M_LABEL, wrap: false,
    fill: o.fill, line: frame,
  });
  if (o.bodyW) {
    body(s, x, y + 0.567, o.bodyW, o.bodyH || 1.428, o.body, o.bodyColor || C.brown, {
      valign: 'middle', margin: M_PANEL, fill: o.fill, line: frame,
    });
  }
}

/** Label pill with no icon and no body (slides 4, 5, 11). */
function tag(s, x, y, w, label, color) {
  text(s, {
    x, y, w: w || 2.497, h: 0.471, text: label, font: F.head, size: 16,
    bold: true, color, align: 'center', valign: 'middle',
    margin: M_LABEL, wrap: false, line: { color, width: 2.25 },
  });
}

/* ------------------------------------------------------------- world map */

/* Continent outlines, 0-1000 x 0-578, flattened [x0,y0,x1,y1,...] per polygon. */
const WORLD = [
  [996,172,960,145,946,140,943,155,868,119,828,136,792,109,757,113,778,87,750,69,681,109,688,141,673,157,655,176,656,116,640,134,645,159,623,145,579,172,572,155,574,179,553,198,539,168,565,165,516,134,459,211,460,240,476,232,482,266,491,260,495,208,511,180,519,188,507,224,533,230,487,279,469,252,472,279,430,318,442,342,417,357,418,399,438,401,454,356,470,347,491,393,499,378,483,338,510,403,510,376,529,371,538,328,544,348,559,327,552,342,567,368,521,379,525,402,550,400,542,439,508,427,500,444,474,421,473,397,428,407,395,499,395,558,422,607,455,594,473,609,489,805,498,850,520,848,549,784,546,756,564,732,586,742,619,703,631,682,633,652,674,672,692,657,702,673,704,632,725,624,731,647,748,634,750,591,772,592,785,538,822,530,823,468,850,447,875,439,900,394,932,381,929,342,913,321,876,326,835,290,802,304,779,281,742,283,726,268,724,240,706,224,724,203,761,215,779,201,821,190,861,205,868,182,894,155,928,169,942,152,966,181,996,172],
  [349,17,321,29,326,44,299,42,271,63,320,79,320,101,297,105,297,132,317,153,331,144,341,161,357,158,352,199,376,244,384,285,349,335,342,393,318,405,321,455,299,462,286,528,251,565,215,573,196,543,187,489,166,477,158,436,131,415,127,341,102,300,73,290,63,251,20,247,17,215,0,206,33,146,86,142,120,163,131,155,157,161,148,133,177,133,191,116,178,90,197,80,208,53,242,58,229,26,266,20,286,42,309,22,349,17],
  [388,660,398,689,414,701,411,758,428,807,451,829,468,879,503,904,530,905,545,930,568,935,584,979,613,977,613,946,646,911,650,864,663,825,652,795,617,770,606,720,594,700,565,704,538,689,502,700,477,676,437,668,388,660],
  [741,782,745,830,779,860,806,857,818,884,844,861,853,829,835,806,806,798,780,808,741,782],
  [313,22,300,35,326,44,313,22],
  [893,803,875,822,901,850,935,833,930,809,893,803],
  [312,165,299,187,326,192,312,165],
  [227,24,196,45,231,60,266,45,227,24],
  [192,138,163,152,196,168,213,148,192,138],
  [764,625,743,647,767,663,782,646,764,625],
  [590,712,573,733,597,749,605,727,590,712],
  [604,373,589,392,612,405,619,384,604,373],
  [868,682,850,700,872,715,881,697,868,682],
  [449,251,435,268,456,283,463,264,449,251],
  [664,308,650,325,671,338,678,320,664,308],
  [601,163,590,180,610,190,616,173,601,163],
  [536,60,522,74,542,86,549,70,536,60],
  [822,530,806,548,827,561,834,543,822,530],
  [706,540,692,556,712,568,719,551,706,540],
  [447,438,434,453,453,464,459,448,447,438],
  [270,86,258,101,277,111,283,95,270,86],
  [672,251,660,265,678,275,684,260,672,251],
  [962,432,950,446,968,456,974,441,962,432],
  [531,331,520,344,537,353,543,339,531,331],
  [852,747,841,760,858,769,864,755,852,747],
  [385,297,375,309,391,318,397,305,385,297],
  [479,624,469,636,485,644,491,632,479,624],
  [714,441,704,453,720,461,726,449,714,441],
  [141,105,131,117,147,125,153,113,141,105],
  [349,570,340,581,355,589,361,577,349,570],
  [913,586,904,597,919,605,924,593,913,586],
  [263,466,254,477,269,484,274,473,263,466],
  [575,205,567,215,581,222,586,212,575,205],
  [800,364,792,374,806,381,811,371,800,364],
];

function worldMap(s, x, y, w, h, color) {
  WORLD.forEach((flat) => {
    const pts = [];
    for (let i = 0; i < flat.length; i += 2) {
      pts.push({ x: (flat[i] / 1000) * w, y: (flat[i + 1] / 578) * h });
    }
    pts.push({ close: true });
    s.addShape('custGeom', { x, y, w, h, points: pts, fill: { color }, line: NO_LINE });
  });
}

/** Soft two-ring location pin used on the slide 12 map. */
function mapPin(s, cx, cy) {
  s.addShape('ellipse', {
    x: cx - 0.194, y: cy - 0.194, w: 0.387, h: 0.387,
    fill: { color: C.markerTan, transparency: 70 }, line: NO_LINE,
  });
  s.addShape('ellipse', {
    x: cx - 0.09, y: cy - 0.09, w: 0.18, h: 0.18,
    fill: { color: C.markerTan, transparency: 30 }, line: NO_LINE,
  });
}

/* ------------------------------------------------------ device mock-ups */

/** Phone mock-up: black body, thin bezel, notch; the screen is a photo well. */
function phone(s, x, y, w, h) {
  const r = w * 0.115;
  s.addShape('roundRect', {
    x: x + w * 0.006, y, w: w * 0.986, h, rectRadius: r,
    fill: { color: '000000' }, line: NO_LINE,
  });
  s.addShape('roundRect', {                                   // inner bezel
    x: x + w * 0.048, y: y + h * 0.022, w: w * 0.904, h: h * 0.956,
    rectRadius: r * 0.9, fill: { color: '2E2E2E' }, line: NO_LINE,
  });
  s.addShape('roundRect', {                                   // screen
    x: x + w * 0.056, y: y + h * 0.027, w: w * 0.887, h: h * 0.946,
    rectRadius: r * 0.82, fill: { color: C.grey }, line: NO_LINE,
  });
  s.addShape('roundRect', {                                   // notch
    x: x + w * 0.445, y: y + h * 0.033, w: w * 0.12, h: h * 0.009,
    rectRadius: h * 0.0045, fill: { color: '000000' }, line: NO_LINE,
  });
}

/** Laptop mock-up: screen lid (`lidH` tall) sitting on a wide tapered base. */
function laptop(s, x, y, w, lidH) {
  s.addShape('roundRect', {
    x, y, w, h: lidH, rectRadius: w * 0.008,
    fill: { color: '111111' }, line: NO_LINE,
  });
  s.addShape('rect', {
    x: x + w * 0.04, y: y + lidH * 0.074, w: w * 0.921, h: lidH * 0.816,
    fill: { color: C.grey }, line: NO_LINE,
  });
  s.addShape('rect', {                                        // chin
    x: x + w * 0.007, y: y + lidH * 0.925, w: w * 0.985, h: lidH * 0.065,
    fill: { color: '181818' }, line: NO_LINE,
  });
  s.addShape('trapezoid', {                                   // base slab
    x: x - w * 0.112, y: y + lidH * 0.966, w: w * 1.224, h: lidH * 0.035,
    fill: { color: 'A5A6A8' }, line: NO_LINE,
  });
  s.addShape('roundRect', {                                   // touchpad lip
    x: x + w * 0.383, y: y + lidH * 0.966, w: w * 0.234, h: lidH * 0.013,
    rectRadius: lidH * 0.0065, fill: { color: 'CFD0D2' }, line: NO_LINE,
  });
}

/* ------------------------------------------------------------ slide 26 */

/**
 * Progress donut: pale ring, a filled `sweep`-degree arc starting at 12 o'clock,
 * and a white hub carrying the percentage.
 */
function progressDonut(s, x, y, d, sweep, label) {
  s.addShape('ellipse', {
    x, y, w: d, h: d, fill: { color: C.tan, transparency: 50 }, line: NO_LINE,
  });
  s.addShape('pie', {
    x: x + 0.002, y: y + 0.002, w: d - 0.004, h: d - 0.004,
    angleRange: [271, 271 + sweep], fill: { color: C.mocha }, line: NO_LINE,
  });
  const hub = d * 0.6212, off = (d - hub) / 2;
  s.addShape('ellipse', {
    x: x + off, y: y + off, w: hub, h: hub, fill: { color: C.white }, line: NO_LINE,
  });
  text(s, {
    x: x + off, y: y + off, w: hub, h: hub, text: label,
    font: F.head, size: 21, color: C.ink, align: 'center', valign: 'middle',
    margin: [0, 0, 0, 0],
  });
}

/* ============================================================== SLIDES */

const build = [];

/* 1 — cover over a full-bleed photograph */
build.push((s) => {
  photo(s, 0, 0, 13.333, 7.5, true);
  wordmark(s, 1.118, 2.229, C.white);
  heading(s, 1.12, 3.368, 5.125, 1.178, QUALITY, C.white);
  body(s, 1.115, 4.621, 4.448, 0.809, [L.d], C.white);
});

/* 2 — centred statement over a full-bleed photograph */
build.push((s) => {
  photo(s, 0, 0, 13.333, 7.5, true);
  const frame = { color: C.white, width: 4.5 };
  text(s, {
    x: 1.667, y: 2.786, w: 10, h: 0.858, text: HEADLINE, font: F.head,
    size: 32, bold: true, color: C.white, align: 'center', valign: 'middle', line: frame,
  });
  body(s, 1.667, 3.643, 10, 1.071, [L.a], C.white, {
    align: 'center', valign: 'middle', line: frame,
  });
});

/* 3 — half photo / half statement */
build.push((s) => {
  photo(s, 0, 0, 7.219, 7.5);
  logotype(s, 8.332, 2.456, C.tan);
  heading(s, 8.326, 3.327, 3.894, 1.717, HEADLINE, C.brown);
});

/* 4 — offset photo collage with a copy column */
build.push((s) => {
  photo(s, 0, 0, 4.443, 5.125);
  photo(s, 1.115, 1.708, 4.443, 5.125);
  photo(s, 10.552, 3.75, 2.781, 3.75);
  logotype(s, 0.53, 0.52, C.white);
  heading(s, 6.115, 1.216, 3.894, 1.717, HEADLINE, C.brown);
  tag(s, 6.115, 3.31, 2.497, 'Quality ingredients', C.tan);
  body(s, 6.115, 4.159, 3.894, 2.432, [L.a, L.b], C.brown);
});

/* 5 — solid brown field beside a photo */
build.push((s) => {
  s.background = { color: C.brown };
  photo(s, 6.115, 0, 7.219, 5.681);
  tag(s, 1.115, 2.4, 2.497, 'Quality ingredients', C.tan);
  body(s, 1.115, 3.249, 3.894, 2.432, [L.a, L.b], C.white);
});

/* 6 — four feature cards over a tinted photograph */
build.push((s) => {
  photo(s, 0, 0, 13.333, 7.5, true);
  s.addShape('rect', {
    x: 0, y: 0, w: 13.333, h: 7.5,
    fill: { color: C.brown, transparency: 30 }, line: NO_LINE,
  });
  heading(s, 1.115, 2.892, 3.894, 1.717, HEADLINE, C.white);
  const cards = [
    { x: 6.1, y: 0.995, icon: 'chat', label: 'The chat' },
    { x: 9.433, y: 0.995, icon: 'glasses', label: 'The view' },
    { x: 6.1, y: 3.564, icon: 'cup', label: 'Cozy place' },
    { x: 9.433, y: 3.564, icon: 'bars', label: 'Best coffee' },
  ];
  cards.forEach((c) => card(s, Object.assign({
    color: C.white, iconBg: '6A5B4A', bodyW: 2.786, body: [L.c], bodyColor: C.white,
  }, c)));
  body(s, 6.115, 5.937, 6.106, 0.567, [L.b], C.white);
});

/* 7 — copy column with stepped photographs on the right */
build.push((s) => {
  photo(s, 7.217, 1.236, 4.443, 6.264);
  photo(s, 8.891, 0, 4.443, 6.264);
  heading(s, 1.106, 1.413, 3.894, 1.717, HEADLINE, C.brown);
  card(s, { x: 1.106, y: 3.655, icon: 'bars', label: 'Best coffee', color: C.tan });
  body(s, 1.106, 4.463, 5.002, 1.947, [L.a, L.b], C.brown);
});

/* 8 — wide banner photo above two captions */
build.push((s) => {
  photo(s, 1.115, 0, 11.106, 3.75);
  card(s, { x: 1.115, y: 4.539, icon: 'chat', label: 'The chat', color: C.tan });
  card(s, { x: 6.667, y: 4.539, icon: 'glasses', label: 'The view', color: C.tan });
  body(s, 1.106, 5.439, 5.002, 1.052, [L.a], C.brown);
  body(s, 6.667, 5.439, 5.002, 1.052, [L.a], C.brown);
});

/* 9 — tall photo beside stacked captioned paragraphs */
build.push((s) => {
  photo(s, 1.112, 0.854, 3.9, 5.791);
  body(s, 5.55, 0.967, 6.671, 1.462, [L.a, L.b], C.brown);
  card(s, { x: 5.566, y: 3.102, icon: 'chat', label: 'The chat', color: C.tan });
  body(s, 5.557, 3.808, 6.671, 0.809, [L.a], C.brown);
  card(s, { x: 5.566, y: 5.017, icon: 'glasses', label: 'The view', color: C.tan });
  body(s, 5.566, 5.723, 6.655, 0.809, [L.a], C.brown);
});

/* 10 — three-up gallery on brown */
build.push((s) => {
  s.background = { color: C.brown };
  card(s, { x: 1.115, y: 0.902, icon: 'glasses', label: 'The view', color: C.tan, iconBg: C.brown });
  body(s, 1.115, 1.608, 11.106, 0.567, [L.a], C.white);
  [1.112, 5.0, 8.885].forEach((x) => photo(s, x, 2.563, 3.331, 4.035, false));
});

/* 11 — distribution timeline */
build.push((s) => {
  heading(s, 1.106, 0.845, 3.894, 0.64, 'Distribution line.', C.brown);
  s.addShape('line', {
    x: 1.115, y: 4.322, w: 11.125, h: 0,
    line: { color: C.sand, width: 3 },
  });
  const stops = [
    { dot: 1.115, x: 1.115, label: 'Quality beans', above: true },
    { dot: 3.786, x: 3.885, label: 'Cafea warehouse', above: false },
    { dot: 6.457, x: 6.562, label: 'Wholesalers', above: true },
    { dot: 9.128, x: 9.25, label: 'Cafes near you', above: false },
  ];
  stops.forEach((st) => {
    s.addShape('ellipse', {
      x: st.dot, y: 4.217, w: 0.21, h: 0.21,
      fill: { color: C.white }, line: { color: C.sand, width: 3 },
    });
    tag(s, st.x, st.above ? 2.119 : 4.807, 2.497, st.label, C.tan);
    body(s, st.x, st.above ? 2.834 : 5.522, 2.771, 1.052, [L.d], C.ink);
  });
});

/* 12 — sourcing map */
build.push((s) => {
  worldMap(s, 1.439, 0.845, 10.455, 6.043, C.mist);
  [[3.243, 3.208], [3.624, 3.867], [5.079, 5.104], [6.667, 3.674],
   [5.989, 2.572], [6.384, 2.853], [9.255, 4.718], [10.289, 5.671],
   [10.289, 3.299]].forEach((p) => mapPin(s, p[0], p[1]));
  heading(s, 1.106, 0.845, 5.561, 1.178, BEANS, C.brown);
  card(s, {
    x: 1.1, y: 4.762, icon: 'cup', label: 'Quality ingredients', labelW: 2.527,
    color: C.tan, bodyW: 3.098, body: [L.e],
  });
});

/* 13 — pie chart */
build.push((s, pptx) => {
  s.addChart(pptx.ChartType.pie, [{
    name: 'Sales',
    labels: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'],
    values: [8.2, 3.2, 1.4, 1.2],
  }], {
    x: 6.383, y: 1.025, w: 6.129, h: 5.764,
    chartColors: [C.sand, C.brown, C.cocoa, C.clay],
    dataBorder: { pt: 0, color: C.white },
    showLegend: true, legendPos: 'b', legendFontSize: 10, legendColor: C.axisDark,
    showTitle: false, firstSliceAng: 0,
  });
  heading(s, 1.106, 0.845, 5.561, 1.178, BEANS, C.brown);
  card(s, {
    x: 1.1, y: 2.5, icon: 'cup', label: 'Quality ingredients', labelW: 4.444,
    color: C.tan, fill: C.white, bodyW: 5.015, bodyH: 1.178, body: [L.f],
  });
  card(s, {
    x: 1.1, y: 4.593, icon: 'bars', label: 'Sustainably sourced', labelW: 4.444,
    color: C.tan, fill: C.white, bodyW: 5.015, bodyH: 1.178, body: [L.g],
  });
});

/* 14 — phone mock-up with the brand lock-up */
build.push((s) => {
  photo(s, 10.552, 0, 2.796, 3.75);
  phone(s, 1.115, 0.984, 3.885, 7.745);
  wordmark(s, 6.105, 1.463, C.tan);
  heading(s, 6.115, 2.769, 4.438, 3.332, QUALITY, C.brown, 48);
  body(s, 9.985, 5.737, 2.796, 1.052, [L.d], C.brown);
});

/* 15 — phone mock-up flanked by copy and cards */
build.push((s) => {
  phone(s, 5.171, 0.768, 2.991, 5.963);
  heading(s, 1.106, 1.685, 3.337, 2.255, HEADLINE, C.brown);
  body(s, 1.106, 4.279, 3.337, 1.537, [L.a], C.brown);
  card(s, { x: 9.448, y: 1.468, icon: 'glasses', label: 'The view', color: C.tan, bodyW: 2.786, body: [L.c] });
  card(s, { x: 9.448, y: 4.037, icon: 'bars', label: 'Best coffee', color: C.tan, bodyW: 2.786, body: [L.c] });
});

/* 16 — cover card floating on a photograph */
build.push((s) => {
  photo(s, 0, 0, 13.333, 7.5, true);
  s.addShape('rect', {
    x: 3.891, y: 1.382, w: 5.552, h: 4.776, fill: { color: C.white }, line: NO_LINE,
  });
  wordmark(s, 4.441, 2.229, C.tan);
  heading(s, 4.443, 3.368, 4.448, 1.178, QUALITY, C.brown);
  body(s, 4.438, 4.621, 4.448, 0.809, [L.d], C.brown);
});

/* 17 — framed statement over a letterboxed photograph */
build.push((s) => {
  photo(s, 1.111, 1.079, 11.111, 5.341, true);
  text(s, {
    x: 5.0, y: 2.165, w: 3.333, h: 3.311, text: HEADLINE, font: F.head, size: 32,
    bold: true, color: C.white, align: 'center', valign: 'middle',
    margin: [21.6, 21.6, 0, 0], line: { color: C.white, width: 4.5 },
  });
  ['Quality ingredients', 'Sustainably sourced'].forEach((t, i) => text(s, {
    x: i ? 4.887 : 4.955, y: i ? 6.772 : 0.366, w: i ? 3.559 : 3.422, h: 0.37,
    text: t, size: 16, color: C.brown, spc: 6, align: 'center', wrap: false,
  }));
});

/* 18 — copy column beside a single photograph */
build.push((s) => {
  photo(s, 6.115, 1.147, 6.115, 5.206);
  logotype(s, 1.12, 1.912, C.tan);
  heading(s, 1.115, 2.783, 3.894, 1.717, HEADLINE, C.brown);
  body(s, 1.106, 4.778, 3.337, 0.809, [L.h], C.brown);
});

/* 19 — three photographs, two captions */
build.push((s) => {
  photo(s, 1.115, 1.188, 4.443, 5.125);
  photo(s, 6.115, 1.188, 2.781, 3.75);
  photo(s, 9.438, 1.188, 2.781, 3.75);
  body(s, 6.104, 5.375, 2.792, 0.809, [L.j], C.brown);
  body(s, 9.427, 5.375, 2.792, 0.809, [L.j], C.brown);
});

/* 20 — split photo / brown panel with a big statement */
build.push((s) => {
  s.background = { color: C.brown };
  photo(s, 0, 0, 7.219, 7.5);
  card(s, {
    x: 1.1, y: 4.512, icon: 'cup', label: 'Quality ingredients', labelW: 2.527,
    color: C.white, iconBg: '6E6E6E', bodyW: 3.098, body: [L.e], bodyColor: C.white,
  });
  wordmark(s, 8.334, 1.463, C.tan);
  heading(s, 8.344, 2.769, 4.438, 3.332, BEANS, C.white, 48);
});

/* 21 — tinted photo, big statement and two wide cards */
build.push((s) => {
  photo(s, 0, 0, 13.333, 7.5, true);
  s.addShape('rect', {
    x: 0, y: 0, w: 13.333, h: 7.5,
    fill: { color: C.brown, transparency: 30 }, line: NO_LINE,
  });
  logotype(s, 1.12, 1.648, C.tan);
  heading(s, 1.115, 2.519, 4.437, 3.332, HEADLINE, C.white, 48);
  body(s, 6.11, 1.078, 6.104, 1.462, [L.a, L.b], C.white);
  card(s, {
    x: 6.115, y: 3.005, icon: 'cup', label: 'Quality ingredients', labelW: 5.527,
    color: C.white, iconBg: '6A5B4A', bodyW: 6.099, bodyH: 1.032, body: [L.f], bodyColor: C.white,
  });
  card(s, {
    x: 6.115, y: 4.824, icon: 'bars', label: 'Sustainably sourced', labelW: 5.527,
    color: C.white, iconBg: '6A5B4A', bodyW: 6.099, bodyH: 1.032, body: [L.g], bodyColor: C.white,
  });
});

/* 22 — two photographs with captions */
build.push((s) => {
  photo(s, 1.105, 0, 4.443, 4.903);
  photo(s, 6.115, 0, 6.113, 4.903);
  body(s, 1.105, 5.345, 4.442, 1.294, [L.a + 'Phasellus massa odio, pellentesque at nunc id, aliquam tempus leo. '], C.brown);
  body(s, 6.115, 5.345, 6.106, 1.294, [L.a + ' ' + L.b], C.brown);
});

/* 23 — half photo, mirrored caption columns */
build.push((s) => {
  photo(s, 0, 0, 7.219, 7.5);
  card(s, { x: 1.674, y: 1.963, icon: 'chat', label: 'The chat', color: C.white, iconBg: C.grey });
  body(s, 1.665, 2.863, 3.337, 2.674, [L.a, L.b], C.white);
  card(s, { x: 8.332, y: 1.963, icon: 'glasses', label: 'The view', color: C.tan, labelW: 2.211 });
  body(s, 8.332, 2.863, 3.335, 2.674, [L.a, L.b], C.brown);
});

/* 24 — brown field with a landscape photo */
build.push((s) => {
  s.background = { color: C.brown };
  photo(s, 5.566, 0.854, 6.655, 3.763);
  logotype(s, 1.12, 1.119, C.tan);
  heading(s, 1.115, 1.99, 3.894, 1.717, HEADLINE, C.white);
  body(s, 1.107, 3.949, 3.893, 2.432, [L.a, L.b], C.white);
  card(s, { x: 5.566, y: 5.017, icon: 'glasses', label: 'The view', color: C.tan, iconBg: C.brown });
  body(s, 5.566, 5.723, 6.655, 0.809, [L.a], C.white);
});

/* 25 — staggered photo mosaic */
build.push((s) => {
  card(s, { x: 1.115, y: 2.699, icon: 'glasses', label: 'The view', color: C.tan });
  body(s, 1.115, 3.506, 3.885, 1.294, [L.a], C.brown);
  photo(s, 6.115, 0, 3.331, 4.035);
  photo(s, 10.003, 0.743, 3.331, 4.035);
  photo(s, 6.115, 4.625, 3.331, 2.875);
  photo(s, 10.003, 5.368, 3.331, 2.132);
});

/* 26 — three progress donuts over a pale map */
build.push((s) => {
  worldMap(s, 2.219, 0.892, 8.896, 5.142, C.linen);
  const cols = [
    { x: 1.696, hubX: 1.559, sweep: 242, label: '70%', icon: 'chat', name: 'America Farm' },
    { x: 5.418, hubX: 5.274, sweep: 72, label: '40%', icon: 'cup', name: 'Africa Farm' },
    { x: 9.141, hubX: 8.999, sweep: 220, label: '60%', icon: 'bars', name: 'Asia Farm' },
  ];
  cols.forEach((c) => {
    progressDonut(s, c.x, 1.384, 2.503, c.sweep, c.label);
    card(s, {
      x: c.hubX, y: 4.359, icon: c.icon, label: c.name, color: C.tan,
      fill: C.white, bodyW: 2.786, body: [L.c],
    });
  });
});

/* 27 — line chart above three note cards */
build.push((s, pptx) => {
  const years = ['2017', '2018', '2019', '2020', '2021'];
  s.addChart(pptx.ChartType.line, [
    { name: 'Lorem Ipsum 1', labels: years, values: [23, 40, 76, 20, 43] },
    { name: 'Lorem Ipsum 2', labels: years, values: [35, 47, 21, 7, 37] },
    { name: 'Lorem Ipsum 3', labels: years, values: [15, 30, 40, 28, 41] },
    { name: 'Lorem Ipsum 4', labels: years, values: [12, 20, 5, 11, 13] },
  ], {
    x: 1.115, y: 1.127, w: 11.106, h: 3.317,
    chartColors: [C.camel, C.sand, C.cocoa, C.brown],
    lineDataSymbol: 'none', lineSize: 2, showLegend: false, showTitle: false,
    catAxisLabelColor: C.axis, valAxisLabelColor: C.axis,
    catAxisLabelFontSize: 11, valAxisLabelFontSize: 11,
    catAxisLineShow: false, valAxisLineShow: false,
    valGridLine: { style: 'solid', color: C.gridFaint, size: 1 },
    valAxisMaxVal: 80, valAxisMajorUnit: 10,
  });
  [1.1, 5.111, 9.122].forEach((x) => card(s, {
    x, y: 5.117, icon: 'cup', label: 'Quality ingredients', labelW: 2.527,
    color: C.tan, bodyW: 3.098, bodyH: 0.889, body: [L.i],
  }));
});

/* 28 — horizontal bar chart beside copy */
build.push((s, pptx) => {
  s.addChart(pptx.ChartType.bar, [{
    name: 'Sales',
    labels: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'],
    values: [8.2, 3.2, 1.4, 1.2],
  }], {
    x: 6.383, y: 1.025, w: 6.129, h: 5.764, barDir: 'bar', barGapWidthPct: 100,
    chartColors: [C.sand, C.brown, C.ink, C.clay],
    showLegend: true, legendPos: 'b', legendFontSize: 10, legendColor: C.axisDark,
    showTitle: false,
    catAxisLabelColor: C.axisDark, valAxisLabelColor: C.axisDark,
    catAxisLabelFontSize: 12, valAxisLabelFontSize: 12,
    catAxisLineColor: C.grid, valAxisLineShow: false,
    valGridLine: { style: 'solid', color: C.grid, size: 1 },
  });
  heading(s, 1.106, 1.582, 5.002, 1.178, BEANS, C.brown);
  card(s, { x: 1.106, y: 3.163, icon: 'bars', label: 'Best coffee', color: C.tan });
  body(s, 1.106, 3.97, 5.002, 1.947, [L.a, L.b], C.brown);
});

/* 29 — laptop mock-up with the brand lock-up */
build.push((s) => {
  laptop(s, 1.189, 1.214, 7.177, 5.068);
  wordmark(s, 8.92, 2.633, C.tan, 0.969);
  body(s, 8.92, 3.815, 3.348, 1.052, [L.d], C.brown);
});

/* 30 — laptop mock-up bleeding off the right edge */
build.push((s) => {
  laptop(s, 5.05, 0.643, 8.796, 6.21);
  card(s, { x: 1.115, y: 1.468, icon: 'glasses', label: 'The view', color: C.tan, bodyW: 2.786, body: [L.c] });
  card(s, { x: 1.115, y: 4.037, icon: 'bars', label: 'Best coffee', color: C.tan, bodyW: 2.786, body: [L.c] });
});

/* ================================================================= main */

function main() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE';
  pptx.theme = { headFontFace: F.head, bodyFontFace: F.body };
  pptx.title = 'cafea';

  build.forEach((fn) => fn(pptx.addSlide(), pptx));

  const out = path.join(__dirname, '041c4842-f6c1-4ecb-badc-047c1f5e5eaa_grok_final.pptx');
  return pptx.writeFile({ fileName: out }).then(() => console.log('wrote ' + out));
}

main();
