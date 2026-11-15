#!/usr/bin/env node
/**
 * "Xander Home Security" — 35-slide deck rebuilt with pptxgenjs.
 *
 * Layout numbers are inches on a 26.667 x 15 in (16:9) stage.
 * Raster artwork in the source deck is replaced by flat placeholder shapes.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette --
const NAVY = '262B47';   // deck background
const TEAL = '46A8C7';   // primary accent
const PURPLE = 'AB80CE'; // secondary accent
const BLEND = '7894CA';  // flat stand-in for the teal -> purple gradient
const WHITE = 'FFFFFF';
const GREY = 'BFBFBF';   // image placeholder tone
const PALE = 'EFEDF7';

const SLIDE_W = 26.667;
const SLIDE_H = 15;

const TITLE_FONT = 'Poppins';
const BODY_FONT = 'Roboto';

/** Fill object helper: `fill(TEAL, 40)` = teal at 40 % opacity. */
function fill(color, opacityPct) {
  return opacityPct === undefined
    ? { color: color }
    : { color: color, transparency: Math.round(100 - opacityPct) };
}

/** Linear interpolation between two hex colours; `t` runs 0 -> 1. */
function mix(c1, c2, t) {
  let out = '';
  for (let i = 0; i < 6; i += 2) {
    const a = parseInt(c1.substr(i, 2), 16);
    const b = parseInt(c2.substr(i, 2), 16);
    out += ('0' + Math.round(a + (b - a) * t).toString(16)).slice(-2);
  }
  return out.toUpperCase();
}

/**
 * Horizontal colour ramp, painted as adjacent vertical strips.
 * pptxgenjs has no gradient fill, so the teal -> purple washes that run
 * through this deck are rebuilt strip by strip. Each strip is shortened at
 * the top and bottom so the silhouette follows a corner radius `r`
 * (r = 0 rectangle, r = h/2 pill, r = w/2 = h/2 circle).
 * `o.archTop` rounds only the top edge; `o.opacity` fades the whole ramp.
 */
function ramp(s, x, y, w, h, c1, c2, o) {
  o = o || {};
  const r = o.r || 0;
  // Wide strips down the straight middle, narrow ones through the rounded
  // ends, so the silhouette stays smooth without exploding the shape count.
  const coarse = 0.16;
  const fine = Math.max(0.02, Math.min(coarse, r / 14));
  let sx = x;
  while (sx < x + w - 1e-6) {
    const edge = Math.min(sx - x, x + w - sx);
    const sw = Math.min(edge < r ? fine : coarse, x + w - sx);
    const mid = sx + sw / 2 - x;
    const dm = Math.min(mid, w - mid);
    const inset = dm >= r ? 0 : r - Math.sqrt(Math.max(0, r * r - (r - dm) * (r - dm)));
    s.addShape('rect', {
      x: sx, y: y + inset, w: sw + 0.006, h: h - inset * 2,
      fill: fill(mix(c1, c2, w <= sw ? 0 : mid / w), o.opacity),
      line: { type: 'none' },
    });
    sx += sw;
  }
}

/** The deck's signature teal -> purple wash. */
function wash(s, x, y, w, h, o) {
  ramp(s, x, y, w, h, TEAL, PURPLE, o);
}

/** Flat approximation of the wash, for glyphs too small to strip. */
const WASH = { color: BLEND };

/** Generic shape with no outline; `extra` carries rectRadius / rotate / arc props. */
function shape(s, kind, x, y, w, h, paint, extra) {
  const opts = { x: x, y: y, w: w, h: h, fill: paint, line: { type: 'none' } };
  Object.assign(opts, extra || {});
  s.addShape(kind, opts);
}

// -------------------------------------------------------------- typography --
/** Text styles reused across the deck (see `text()` for the defaults). */
const HEAD = { font: TITLE_FONT, size: 85, bold: true };
const HEAD_C = { font: TITLE_FONT, size: 85, bold: true, align: 'center' };
const SUB = { size: 28, bold: true };
const BODY = { valign: 'middle', ls: 1.5, margin: 3.6 };
const NOTE = { ls: 1.5 };

/** Text box. Defaults: Roboto 23 pt, white, left / top aligned, 7.2 pt inset. */
function text(s, x, y, w, h, body, st) {
  st = st || {};
  const opts = {
    x: x, y: y, w: w, h: h,
    align: st.align || 'left',
    valign: st.valign || 'top',
    fontFace: st.font || BODY_FONT,
    fontSize: st.size || 23,
    color: st.color || WHITE,
    margin: st.margin === undefined ? [7.2, 7.2, 3.6, 3.6] : st.margin,
  };
  if (st.bold) opts.bold = true;
  if (st.italic) opts.italic = true;
  if (st.ls) opts.lineSpacingMultiple = st.ls;
  s.addText(body, opts);
}

// ------------------------------------------------------- background motifs --

/**
 * Soft radial bloom: concentric discs, painted outside-in, whose stacked
 * opacity reproduces the linear centre->edge ramp of the original gradient.
 */
function glow(s, x, y, size, color, peakPct) {
  const cx = x + size / 2;
  const cy = y + size / 2;
  const steps = 12;
  let stacked = 0;
  for (let i = 1; i <= steps; i++) {
    const target = (peakPct / 100) * (i / steps);
    const layer = (target - stacked) / (1 - stacked);
    stacked = target;
    const r = (size * 0.75 / 2) * (1 - (i - 1) / steps);
    s.addShape('ellipse', {
      x: cx - r, y: cy - r, w: r * 2, h: r * 2,
      fill: fill(color, layer * 100), line: { type: 'none' },
    });
  }
}

/** Tiny white "star" highlight (0.666 in in the source deck). */
function dot(s, x, y) {
  const steps = [[0.50, 12], [0.34, 26], [0.20, 55], [0.10, 100]];
  const cx = x + 0.333;
  const cy = y + 0.333;
  steps.forEach(function (st) {
    const r = 0.666 * st[0];
    s.addShape('ellipse', {
      x: cx - r, y: cy - r, w: r * 2, h: r * 2,
      fill: fill(WHITE, st[1]), line: { type: 'none' },
    });
  });
}

/** Big "planet": two hairline rings around a translucent disc. */
function rings(s, x, y) {
  s.addShape('donut', { x: x, y: y, w: 10.489, h: 10.489, fill: fill(BLEND, 50), line: { type: 'none' }, rectRadius: 0.08 });
  s.addShape('donut', { x: x + 0.519, y: y + 0.487, w: 9.452, h: 9.452, fill: fill(BLEND, 50), line: { type: 'none' }, rectRadius: 0.073 });
  s.addShape('ellipse', { x: x + 1.014, y: y + 0.983, w: 8.461, h: 8.461, fill: fill(BLEND, 50), line: { type: 'none' } });
}

/** Pair of faint concentric hairline rings (no disc). */
function halo(s, x, y, size) {
  const inner = size * 0.8993;
  s.addShape('donut', { x: x, y: y, w: size, h: size, fill: fill(BLEND, 20), line: { type: 'none' }, rectRadius: size * 0.00767 });
  s.addShape('donut', { x: x + size * 0.0504, y: y + size * 0.0508, w: inner, h: inner, fill: fill(BLEND, 20), line: { type: 'none' }, rectRadius: inner * 0.00767 });
}

/** Draws the per-slide list of background ornaments, in back-to-front order. */
function deco(s, items) {
  items.forEach(function (it) {
    const x = it[0], y = it[1], kind = it[2];
    if (kind === 'glow') glow(s, x, y, it[3], it[4], 40);
    else if (kind === 'dot') dot(s, x, y);
    else if (kind === 'rings') rings(s, x, y);
    else if (kind === 'halo') halo(s, x, y, it[3]);
  });
}

// ------------------------------------------------------------ chrome bits --

/** Rounded "next" / "back" pill in the accent gradient. */
function pillButton(s, x, y, label) {
  wash(s, x, y, 2.35, 0.968, { r: 0.376, opacity: 50 });
  s.addText(label, {
    x: x + 0.424, y: y + 0.231, w: 1.503, h: 0.505,
    align: 'center', valign: 'top', fontFace: TITLE_FONT, fontSize: 24, bold: true, color: WHITE,
  });
}

/** Slide number, top-right corner. */
function pageNumber(s, label) {
  s.addText(label, {
    x: 25.341, y: 0.654, w: 1.144, h: 0.572,
    align: 'center', valign: 'top', fontFace: TITLE_FONT, fontSize: 28, bold: true, color: WHITE,
  });
}

/** Facebook / Twitter / LinkedIn glyph column, top-left corner. */
function socialStrip(s) {
  icon(s, 'facebook', 0.496, 0.751, 0.205, 0.39, WASH);
  icon(s, 'twitter', 0.42, 1.556, 0.357, 0.29, WASH);
  icon(s, 'linkedin', 0.447, 2.261, 0.304, 0.29, WASH);
}

/** Flat stand-in for a photograph; `r` is the corner radius in inches. */
function photo(s, x, y, w, h, shape, r) {
  const opts = { x: x, y: y, w: w, h: h, fill: fill(GREY), line: { type: 'none' } };
  if (shape === 'roundRect') opts.rectRadius = r || 0.0001;
  s.addShape(shape, opts);
}

// ------------------------------------------------------------------ icons --
/**
 * Pictograms assembled from native shapes.
 * Each part is [shape, x, y, w, h, cornerFraction, extraOpts], with x/y/w/h
 * given as fractions of the icon's bounding box.
 */
const ICONS = {
  facebook: [
    ['rect', 0.10, 0.33, 0.90, 0.19],
    ['rect', 0.34, 0.14, 0.38, 0.86],
    ['roundRect', 0.34, 0.00, 0.66, 0.30, 0.45],
  ],
  twitter: [
    ['moon', 0.00, 0.06, 0.92, 0.92, 0, { rotate: 220 }],
    ['triangle', 0.56, -0.04, 0.44, 0.44, 0, { rotate: 40 }],
    ['ellipse', 0.28, 0.22, 0.44, 0.40],
  ],
  linkedin: [
    ['rect', 0.00, 0.32, 0.26, 0.68],
    ['ellipse', 0.00, 0.00, 0.26, 0.24],
    ['roundRect', 0.36, 0.28, 0.64, 0.72, 0.45],
    ['rect', 0.36, 0.62, 0.64, 0.38],
    ['rect', 0.36, 0.32, 0.21, 0.68],
  ],
  camera: [
    ['roundRect', 0.00, 0.20, 1.00, 0.80, 0.22],
    ['roundRect', 0.26, 0.00, 0.34, 0.32, 0.30],
    ['ellipse', 0.30, 0.34, 0.42, 0.50, 0, { cut: true }],
    ['ellipse', 0.37, 0.42, 0.28, 0.34],
  ],
  video: [
    ['roundRect', 0.00, 0.10, 0.60, 0.80, 0.22],
    ['triangle', 0.54, 0.20, 0.66, 0.60, 0, { rotate: 270 }],
  ],
  speaker: [
    ['rect', 0.00, 0.34, 0.26, 0.32],
    ['triangle', -0.01, 0.25, 0.96, 0.50, 0, { rotate: 270 }],
    ['blockArc', 0.14, 0.06, 0.88, 0.88, 0, { angleRange: [300, 60], arcThicknessRatio: 0.20 }],
  ],
  cloud: [
    ['ellipse', 0.00, 0.34, 0.44, 0.44],
    ['ellipse', 0.56, 0.28, 0.44, 0.44],
    ['ellipse', 0.22, 0.18, 0.54, 0.54],
    ['rect', 0.14, 0.54, 0.72, 0.24],
    ['triangle', 0.30, 0.34, 0.40, 0.26, 0, { cut: true }],
    ['rect', 0.42, 0.54, 0.16, 0.22, 0, { cut: true }],
  ],
  archive: [
    ['roundRect', 0.00, 0.06, 1.00, 0.22, 0.16],
    ['rect', 0.09, 0.32, 0.82, 0.62],
    ['rect', 0.36, 0.48, 0.28, 0.10, 0, { cut: true }],
  ],
  phone: [
    ['blockArc', 0.00, 0.00, 1.00, 1.00, 0, { angleRange: [30, 215], arcThicknessRatio: 0.42 }],
    ['roundRect', 0.67, 0.53, 0.34, 0.34, 0.40],
    ['roundRect', 0.01, 0.10, 0.34, 0.34, 0.40],
  ],
  addressBook: [
    ['rect', 0.16, 0.00, 0.70, 1.00],
    ['rect', 0.02, 0.00, 0.07, 1.00],
    ['rect', 0.86, 0.10, 0.14, 0.10],
    ['rect', 0.86, 0.45, 0.14, 0.10],
    ['rect', 0.86, 0.80, 0.14, 0.10],
    ['ellipse', 0.42, 0.18, 0.24, 0.24, 0, { cut: true }],
    ['roundRect', 0.32, 0.50, 0.44, 0.38, 0.42, { cut: true }],
  ],
  envelope: [
    ['roundRect', 0.00, 0.10, 1.00, 0.80, 0.14],
    ['triangle', 0.05, 0.14, 0.90, 0.50, 0, { rotate: 180, cut: true }],
  ],
  globe: [
    ['ellipse', 0.00, 0.00, 1.00, 1.00],
    ['ellipse', 0.27, -0.02, 0.46, 1.04, 0, { cut: true }],
    ['rect', 0.00, 0.40, 1.00, 0.18, 0, { cut: true }],
    ['ellipse', 0.40, -0.02, 0.20, 1.04],
    ['rect', 0.00, 0.44, 1.00, 0.10],
  ],
  person: [
    ['ellipse', 0.24, 0.00, 0.52, 0.30],
    ['roundRect', 0.06, 0.30, 0.88, 0.42, 0.44],
    ['roundRect', 0.14, 0.50, 0.30, 0.50, 0.16],
    ['roundRect', 0.56, 0.50, 0.30, 0.50, 0.16],
  ],
};

/**
 * Draw one pictogram inside the given box.
 * `paint` is the glyph fill; `bg` paints the knocked-out details (lens ring,
 * envelope flap, globe meridians) and must match whatever sits behind.
 */
function icon(s, name, x, y, w, h, paint, bg) {
  (ICONS[name] || []).forEach(function (part) {
    const shape = part[0];
    const extra = part[6] || {};
    const opts = {
      x: x + part[1] * w, y: y + part[2] * h, w: part[3] * w, h: part[4] * h,
      fill: extra.cut ? fill(bg || WHITE) : paint, line: { type: 'none' },
    };
    if (part[5]) opts.rectRadius = part[5] * Math.min(part[3] * w, part[4] * h);
    if (extra.rotate) opts.rotate = extra.rotate;
    if (extra.angleRange) opts.angleRange = extra.angleRange;
    if (extra.arcThicknessRatio) opts.arcThicknessRatio = extra.arcThicknessRatio;
    s.addShape(shape, opts);
  });
}

/** Outlined rounded frame used by the step diagram on slide 28. */
function stepFrame(s, x, y, w, h, color) {
  s.addShape('roundRect', {
    x: x, y: y, w: w, h: h, rectRadius: Math.min(w, h) * 0.21,
    fill: { type: 'none' }, line: { color: color, width: 3 },
  });
}

/** Filled label tab of the step diagram. */
function stepTab(s, x, y, w, h, color) {
  s.addShape('roundRect', {
    x: x, y: y, w: w, h: h, rectRadius: Math.min(w, h) * 0.26,
    fill: fill(color), line: { type: 'none' },
  });
}

/**
 * Semicircular connector arrow between two step cards.
 * `down` bulges it below the row instead of above; each half takes its own
 * colour so the pair still reads as the teal -> purple accent gradient.
 */
function arcArrow(s, x, y, w, h, down) {
  const cy = down ? y - h : y;                 // full circle box is 2h tall
  const tipY = down ? y : y + h;
  // Angles run clockwise from 3 o'clock: 90 = bottom, 180 = left, 270 = top.
  const halves = down ? [[90, 180], [0, 90]] : [[180, 270], [270, 0]];
  const tone = down ? [PURPLE, TEAL] : [TEAL, PURPLE];
  halves.forEach(function (range, i) {
    s.addShape('blockArc', {
      x: x, y: cy, w: w, h: h * 2, fill: fill(tone[i]), line: { type: 'none' },
      angleRange: range, arcThicknessRatio: 0.075,
    });
  });
  s.addShape('triangle', {
    x: x + w - 0.28, y: tipY - (down ? 0.40 : 0.06), w: 0.56, h: 0.44,
    fill: fill(tone[1]), line: { type: 'none' }, rotate: down ? 0 : 180,
  });
}

/**
 * Arch card: semicircular top over a square body, both in the accent wash.
 * The dome is swept out of pie wedges, which keeps its edge perfectly round.
 */
function arch(s, x, y, w, h) {
  const wedges = 30;
  for (let i = 0; i < wedges; i++) {
    const a0 = 180 + (i * 180) / wedges;
    const a1 = 180 + ((i + 1) * 180) / wedges;
    s.addShape('pie', {
      x: x, y: y, w: w, h: w, angleRange: [a0, a1 >= 360 ? 0 : a1],
      fill: fill(mix(TEAL, PURPLE, (i + 0.5) / wedges)), line: { type: 'none' },
    });
  }
  wash(s, x, y + w / 2, w, h - w / 2);
}

// ----------------------------------------------------------------- charts --

/** Clustered column chart used on slide 26. */
function columnChart(s, x, y, w, h) {
  const cats = ['Category 1', 'Category 2', 'Category 3'];
  s.addChart('bar', [
    { name: 'Series 1', labels: cats, values: [4.3, 2.5, 3.5] },
    { name: 'Series 2', labels: cats, values: [2.4, 4.4, 1.8] },
  ], {
    x: x, y: y, w: w, h: h,
    barDir: 'col', barGapWidthPct: 219, barOverlapPct: -27,
    chartColors: [TEAL, PURPLE],
    showTitle: true, title: 'Chart Title', titleColor: WHITE, titleFontFace: 'Open Sans', titleFontSize: 23, titleBold: true,
    showLegend: true, legendPos: 'b', legendColor: WHITE, legendFontFace: 'Open Sans', legendFontSize: 21,
    catAxisLabelColor: WHITE, catAxisLabelFontFace: 'Open Sans', catAxisLabelFontSize: 21,
    valAxisLabelColor: WHITE, valAxisLabelFontFace: 'Open Sans', valAxisLabelFontSize: 21,
    valGridLine: { color: 'D9D9D9', style: 'solid', size: 0.75 },
    catAxisLineShow: true, valAxisLineShow: false,
    plotArea: { fill: { type: 'none' } },
  });
}

/**
 * Waterfall chart (slide 27). The source deck ships this as a flat picture,
 * so it is rebuilt here from bars, gridlines and labels.
 * Each entry is [label, floor, ceiling, colour, labelBelow].
 */
function waterfall(s, x, y, w, h) {
  const bars = [
    ['100', 0, 100, TEAL, false], ['20', 100, 120, PURPLE, false],
    ['50', 120, 170, TEAL, false], ['-40', 130, 170, PURPLE, true],
    ['130', 0, 130, TEAL, false], ['-60', 70, 130, PURPLE, true],
    ['70', 70, 140, TEAL, false], ['140', 0, 140, PURPLE, false],
  ];
  const MAX = 180;
  const left = x + 0.062 * w, right = x + 0.914 * w;
  const top = y + 0.168 * h, base = y + 0.807 * h;
  const at = function (v) { return base - (base - top) * (v / MAX); };
  const slot = 0.1068 * w, bw = 0.073 * w;

  s.addText('Chart Title', { x: x, y: y, w: w, h: 0.42, align: 'center', valign: 'middle', fontFace: 'Open Sans', fontSize: 16, color: WHITE, margin: 0 });
  s.addText([
    { text: '\u25AA Increase  ', options: { color: '92D050' } },
    { text: '\u25AA Decrease  ', options: { color: '92D050' } },
    { text: '\u25AA Total', options: { color: '92D050' } },
  ], { x: x, y: y + 0.42, w: w, h: 0.38, align: 'center', valign: 'middle', fontFace: 'Open Sans', fontSize: 13, color: WHITE, margin: 0 });

  for (let v = 0; v <= MAX; v += 20) {
    s.addShape('rect', { x: left, y: at(v), w: right - left, h: 0.012, fill: fill(WHITE, 30), line: { type: 'none' } });
    s.addText(String(v), { x: left - 0.95, y: at(v) - 0.17, w: 0.85, h: 0.34, align: 'right', valign: 'middle', fontFace: 'Open Sans', fontSize: 13, color: WHITE, margin: 0 });
  }
  s.addShape('rect', { x: left, y: top, w: 0.012, h: base - top, fill: fill(WHITE, 30), line: { type: 'none' } });

  bars.forEach(function (b, i) {
    const cx = left + 0.045 * w + i * slot;
    s.addShape('rect', { x: cx - bw / 2, y: at(b[2]), w: bw, h: at(b[1]) - at(b[2]), fill: fill(b[3]), line: { type: 'none' } });
    s.addText(b[0], {
      x: cx - slot / 2, y: (b[4] ? at(b[1]) + 0.03 : at(b[2]) - 0.37), w: slot, h: 0.34,
      align: 'center', valign: 'middle', fontFace: 'Open Sans', fontSize: 13, color: '203864', margin: 0,
    });
    s.addText('Category ' + (i + 1), {
      x: cx - slot * 0.9, y: base + (i % 2 ? 0.06 : 0.44), w: slot * 1.8, h: 0.4,
      align: 'center', valign: 'middle', fontFace: 'Open Sans', fontSize: 14, color: WHITE, margin: 0,
    });
  });
}

// ------------------------------------------------------------- slides --

function slide01(s) {
  deco(s, [[18.213, -7.664, "glow", 15.327, TEAL], [1.326, -5.729, "glow", 14.071, PURPLE], 
    [12.968, 1.445, "rings"], [21.412, 2.518, "dot"], [14.266, 3.283, "dot"], [13.55, 9.187, "dot"]]);
  photo(s, 14.914, 3.879, 14.017, 14.019, "ellipse");
  deco(s, [[-5.328, 4.606, "glow", 15.327, TEAL]]);
  text(s, 3.07, 3.242, 10.609, 2.794, "XANDER", { font: TITLE_FONT, size: 160, bold: true });
  text(s, 3.07, 5.741, 6.791, 1.01, "Home Security", { font: TITLE_FONT, size: 54 });
  text(s, 3.07, 6.753, 6.791, 0.572, "Presentation Template", { font: TITLE_FONT, size: 28 });
  deco(s, [[-3.51, 10.487, "halo", 8.795]]);
  text(s, 7.407, 9.575, 4.283, 2.359, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown.", BODY);
  socialStrip(s);
  pillButton(s, 3.07, 8.412, "next");
  pageNumber(s, "01");
}

function slide02(s) {
  deco(s, [[17.239, 7.112, "glow", 15.327, TEAL], [4.125, -7.382, "glow", 14.071, PURPLE], 
    [15.794, 1.336, "rings"], [24.237, 2.408, "dot"], [17.091, 3.173, "dot"], [25, 9.243, "dot"]]);
  photo(s, 13.299, 4.286, 8.886, 8.887, "ellipse");
  deco(s, [[-7.423, 3.116, "glow", 15.327, TEAL]]);
  text(s, 2.343, 2.523, 11.967, 2.962, "Welcome to Xander Security", HEAD);
  deco(s, [[-4.314, 10.658, "halo", 8.795]]);
  socialStrip(s);
  pillButton(s, 11.817, 12.068, "next");
  pageNumber(s, "02");
  text(s, 4.33, 6.344, 5.595, 2.939, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book.", BODY);
  text(s, 4.32, 9.71, 5.595, 1.778, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when.", BODY);
}

function slide03(s) {
  deco(s, [[-4.335, -7.74, "glow", 15.327, TEAL], [-2.19, 5.515, "glow", 14.071, PURPLE], 
    [12.552, -4.628, "glow", 15.327, TEAL], [1.517, 4.648, "rings"], [9.96, 5.72, "dot"], [2.815, 6.485, "dot"], 
    [18.641, 7.291, "halo", 10.989]]);
  photo(s, -0.041, 8.503, 7.661, 6.526, "rect");
  photo(s, 19.398, 2.421, 5.908, 9.071, "rect");
  text(s, 8.025, 1.842, 10.964, 2.962, [{ text: "History of", options: { breakLine: true } }, { text: "Xander Security" }], HEAD);
  socialStrip(s);
  pillButton(s, 21.908, 12.159, "next");
  pageNumber(s, "03");
  text(s, 13.178, 7.333, 5.595, 2.939, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book.", BODY);
  text(s, 13.168, 10.699, 5.595, 1.778, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when.", BODY);
}

function slide04(s) {
  deco(s, [[11.34, 8.546, "glow", 15.327, TEAL], [6.106, -7.42, "glow", 14.071, PURPLE], 
    [15.794, 4.068, "rings"], [24.237, 5.141, "dot"], [15.705, 10.31, "dot"], [25, 11.975, "dot"]]);
  photo(s, 13.093, 0.905, 8.886, 8.887, "ellipse");
  deco(s, [[-8.095, 2.421, "glow", 15.327, TEAL]]);
  text(s, 2.286, 2.421, 11.967, 2.962, [{ text: "History of ", options: { breakLine: true } }, { text: "Xander Security" }], HEAD);
  deco(s, [[-3.495, 9.852, "halo", 8.795]]);
  socialStrip(s);
  pillButton(s, 10.362, 11.993, "next");
  pageNumber(s, "04");
  text(s, 5.237, 6.576, 5.595, 2.939, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book.", BODY);
  text(s, 5.227, 9.942, 5.595, 1.778, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when.", BODY);
}

function slide05(s) {
  deco(s, [[4.599, -7.827, "glow", 15.327, TEAL], [-2.19, 5.515, "glow", 14.071, PURPLE], [1.14, 1.416, "rings"], 
    [9.583, 2.488, "dot"], [2.437, 3.253, "dot"], [1.066, 7.851, "dot"]]);
  photo(s, 5.895, 5.462, 7.959, 7.008, "rect");
  deco(s, [[18.885, -0.325, "glow", 15.327, TEAL]]);
  text(s, 12.631, 1.846, 10.964, 2.962, [{ text: "About", options: { breakLine: true } }, { text: "Xander Security" }], HEAD);
  deco(s, [[22.087, -4.397, "halo", 8.795]]);
  socialStrip(s);
  pillButton(s, 21.908, 12.469, "next");
  pageNumber(s, "05");
  text(s, 15.673, 6.381, 5.595, 2.939, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book.", BODY);
  text(s, 15.662, 9.747, 5.595, 1.778, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when.", BODY);
}

function slide06(s) {
  deco(s, [[14.807, -7.597, "glow", 15.327, TEAL], [-0.612, 7.214, "glow", 14.071, PURPLE], 
    [7.677, 3.377, "rings"], [16.121, 4.45, "dot"], [8.975, 5.215, "dot"], [20.863, 2.678, "halo", 10.821]]);
  photo(s, 12.4, 5.133, 14.244, 6.057, "rect");
  deco(s, [[-4.781, -8.124, "glow", 15.327, TEAL]]);
  text(s, 2.319, 2.199, 9.612, 2.962, [{ text: "About Xander", options: { breakLine: true } }, { text: "Security" }], HEAD);
  deco(s, [[7.604, 9.812, "dot"]]);
  socialStrip(s);
  pillButton(s, 2.364, 11.406, "next");
  pageNumber(s, "06");
  text(s, 2.319, 6.958, 4.78, 3.52, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book.", BODY);
}

function slide07(s) {
  deco(s, [[-2.39, 4.873, "glow", 15.327, TEAL], [-2.505, -4.622, "glow", 14.071, PURPLE], 
    [4.531, 1.926, "rings"], [12.974, 2.998, "dot"], [5.828, 3.763, "dot"], [14.731, 6.806, "dot"]]);
  photo(s, -1.13, 5.138, 14.017, 14.019, "ellipse");
  deco(s, [[19.77, -4.238, "glow", 15.327, TEAL]]);
  text(s, 15.631, 2.315, 8.279, 1.531, "Break Slide", HEAD);
  deco(s, [[22.271, 10.906, "halo", 8.795]]);
  socialStrip(s);
  pillButton(s, 16.298, 11.634, "next");
  pageNumber(s, "07");
  text(s, 16.378, 6.164, 7.114, 4.214, "“This is a quote. Words full of wisdom that someone important said.”", { font: TITLE_FONT, size: 45, italic: true, valign: "middle", margin: 7.2 });
}

function slide08(s) {
  deco(s, [[12.416, 8.518, "glow", 15.327, TEAL], [5.245, -8.076, "glow", 14.071, PURPLE], 
    [15.794, 4.068, "rings"], [24.237, 5.141, "dot"], [15.705, 10.31, "dot"], [25, 11.975, "dot"]]);
  photo(s, 15.131, 0.873, 8.886, 8.887, "ellipse");
  deco(s, [[-6.431, 6.344, "glow", 15.327, TEAL], [-4.694, 11.975, "halo", 8.795]]);
  socialStrip(s);
  pillButton(s, 13.158, 7.318, "next");
  pageNumber(s, "08");
  text(s, 2.245, 2.306, 10.964, 2.962, [{ text: "Why Choose ", options: { breakLine: true } }, { text: "Xander Security" }], HEAD);
  shape(s, "ellipse", 2.327, 8.222, 2, 2, fill(WHITE));
  icon(s, "camera", 2.956, 8.844, 0.8, 0.7, WASH, WHITE);
  text(s, 2.475, 11.311, 5.595, 1.198, "Lorem Ipsum has been the industry's standard dummy text ever since.", BODY);
  shape(s, "ellipse", 8.542, 8.222, 2, 2, fill(WHITE));
  icon(s, "video", 9.171, 8.918, 0.8, 0.6, WASH);
  text(s, 2.459, 6.027, 8.164, 1.198, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown.", BODY);
  text(s, 2.475, 10.634, 5.083, 0.572, "Lorem Ipsum Dolor", SUB);
  text(s, 8.542, 11.311, 5.595, 1.198, "Lorem Ipsum has been the industry's standard dummy text ever since.", BODY);
  text(s, 8.542, 10.634, 5.083, 0.572, "Lorem Ipsum Dolor", SUB);
}

function slide09(s) {
  photo(s, -0.024, -0.06, 11.135, 15.06, "rect");
  deco(s, [[16.348, 5.202, "glow", 14.071, PURPLE], [20.203, -3.465, "rings"], [20.885, 4.525, "dot"], 
    [25.788, 6.032, "dot"]]);
  text(s, 12.555, 1.541, 10.964, 2.962, [{ text: "Why Choose ", options: { breakLine: true } }, { text: "Xander Security" }], HEAD);
  socialStrip(s);
  pillButton(s, 12.681, 6.32, "next");
  pageNumber(s, "09");
  wash(s, 2.586, 9.289, 21.672, 4.007, { r: 2.003, opacity: 49.803 });
  shape(s, "ellipse", 3.198, 10.056, 2.5, 2.5, fill("E9E6E3"));
  text(s, 6.173, 10.185, 5.083, 0.572, "Lorem Ipsum", { font: "Roboto Black", size: 28, bold: true });
  text(s, 6.173, 10.758, 6.491, 1.781, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. sed do eiusmod tempor incididunt ", NOTE);
  shape(s, "ellipse", 13.848, 10.035, 2.5, 2.5, fill("E9E6E3"));
  text(s, 16.823, 10.164, 5.083, 0.572, "Lorem Ipsum", { font: "Roboto Black", size: 28, bold: true });
  text(s, 16.823, 10.736, 6.491, 1.781, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. sed do eiusmod tempor incididunt ", NOTE);
  icon(s, "speaker", 14.698, 10.956, 0.8, 0.7, WASH);
  icon(s, "video", 4.069, 10.956, 0.8, 0.7, WASH);
  text(s, 15.746, 6.14, 5.278, 1.778, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown.", BODY);
}

function slide10(s) {
  deco(s, [[5.634, -8.414, "glow", 15.327, TEAL], [-4.339, 6.061, "glow", 14.071, PURPLE], 
    [1.14, 1.416, "rings"], [9.583, 2.488, "dot"], [2.437, 3.253, "dot"], [1.066, 7.851, "dot"]]);
  photo(s, 6.036, 5.596, 7.742, 7.008, "rect");
  deco(s, [[21.011, 1.338, "glow", 15.327, TEAL], [22.087, -4.397, "halo", 8.795]]);
  socialStrip(s);
  pillButton(s, 21.911, 12.082, "next");
  pageNumber(s, "10");
  text(s, 12.631, 2.102, 10.964, 2.962, [{ text: "Why Choose ", options: { breakLine: true } }, { text: "Xander Security" }], HEAD);
  shape(s, "ellipse", 13.067, 6.808, 2, 2, fill(WHITE));
  icon(s, "camera", 13.696, 7.43, 0.8, 0.7, WASH, WHITE);
  text(s, 15.51, 6.828, 5.595, 1.778, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when.", BODY);
  shape(s, "ellipse", 13.067, 9.478, 2, 2, fill(WHITE));
  text(s, 15.51, 9.498, 5.595, 1.778, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when.", BODY);
  icon(s, "video", 13.696, 10.175, 0.8, 0.6, WASH);
}

function slide11(s) {
  deco(s, [[2.374, -1.681, "glow", 15.327, TEAL]]);
  photo(s, 7.479, 1.567, 19.19, 5.157, "rect");
  deco(s, [[18.821, 9.23, "glow", 15.327, TEAL], [-3.232, 7.176, "rings"], [5.212, 8.249, "dot"], 
    [0.19, 7.605, "dot"], [6.172, 13.354, "dot"], [22.606, 11.875, "halo", 8.795]]);
  socialStrip(s);
  pillButton(s, 2.643, 3.395, "next");
  pageNumber(s, "11");
  text(s, 7.362, 7.259, 10.964, 2.962, [{ text: "Why Choose ", options: { breakLine: true } }, { text: "Xander Security" }], HEAD);
  shape(s, "ellipse", 8.425, 10.675, 2, 2, fill(WHITE));
  icon(s, "camera", 9.053, 11.296, 0.8, 0.7, WASH, WHITE);
  text(s, 11.267, 11.485, 5.595, 1.198, "Lorem Ipsum has been the industry's standard dummy text ever since.", BODY);
  shape(s, "ellipse", 17.135, 10.674, 2, 2, fill(WHITE));
  icon(s, "video", 17.763, 11.371, 0.8, 0.6, WASH);
  text(s, 11.267, 10.808, 5.083, 0.572, "Lorem Ipsum Dolor", SUB);
  text(s, 19.965, 11.485, 5.595, 1.198, "Lorem Ipsum has been the industry's standard dummy text ever since.", BODY);
  text(s, 19.965, 10.808, 5.083, 0.572, "Lorem Ipsum Dolor", SUB);
}

function slide12(s) {
  deco(s, [[11.34, 8.546, "glow", 15.327, TEAL], [6.106, -7.42, "glow", 14.071, PURPLE], 
    [15.794, 4.068, "rings"], [24.237, 5.141, "dot"], [15.705, 10.31, "dot"], [25, 11.975, "dot"]]);
  photo(s, 12.952, 1.132, 8.887, 8.887, "ellipse");
  deco(s, [[-8.095, 2.421, "glow", 15.327, TEAL], [-3.495, 9.852, "halo", 8.795]]);
  socialStrip(s);
  pillButton(s, 10.602, 11.039, "next");
  pageNumber(s, "12");
  text(s, 3.311, 3.039, 8.279, 1.531, "Break Slide", HEAD);
  text(s, 5.227, 6.1, 7.114, 4.214, "“This is a quote. Words full of wisdom that someone important said.”", { font: TITLE_FONT, size: 45, italic: true, valign: "middle", margin: 7.2 });
}

function slide13(s) {
  photo(s, 11.686, 1.428, 6.782, 5.345, "roundRect", 0);
  deco(s, [[5.634, -8.414, "glow", 15.327, TEAL]]);
  photo(s, 11.688, 7.953, 6.782, 5.345, "roundRect", 0);
  deco(s, [[-5.133, 6.626, "glow", 14.071, PURPLE], [21.398, 3.149, "glow", 15.327, TEAL], 
    [0.496, 0.355, "rings"], [8.94, 1.428, "dot"], [1.794, 2.193, "dot"], [8.348, 9.538, "dot"], 
    [22.087, -4.397, "halo", 8.795]]);
  socialStrip(s);
  pillButton(s, 22.087, 7.081, "next");
  pageNumber(s, "13");
  text(s, 2.493, 11.087, 7.606, 1.531, "Our Service", HEAD);
  text(s, 19.175, 3.986, 4.421, 1.778, "Lorem Ipsum has been the industry's standard dummy text ever since.", BODY);
  text(s, 19.175, 3.149, 5.083, 0.572, "Service One", SUB);
  text(s, 19.177, 10.511, 4.421, 1.778, "Lorem Ipsum has been the industry's standard dummy text ever since.", BODY);
  text(s, 19.177, 9.674, 5.083, 0.572, "Service Two", SUB);
}

function slide14(s) {
  deco(s, [[4.646, 8.701, "rings"], [7.136, 9.005, "dot"], [4.933, 13.09, "dot"]]);
  photo(s, 9.942, 3.572, 6.782, 10.054, "roundRect", 0);
  deco(s, [[17.762, 2.406, "glow", 15.327, TEAL], [-5.122, -5.92, "glow", 14.071, PURPLE]]);
  text(s, 8.757, 1.515, 9.153, 1.531, "Our Service", HEAD_C);
  socialStrip(s);
  pillButton(s, 2.34, 8.701, "next");
  pageNumber(s, "14");
  text(s, 2.362, 4.2, 5.595, 2.939, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book.", BODY);
  deco(s, [[23.582, 1.846, "halo", 8.795]]);
  text(s, 18.825, 4.737, 4.421, 1.778, "Lorem Ipsum has been the industry's standard dummy text ever since.", BODY);
  text(s, 18.825, 4.171, 5.083, 0.572, "Service One", SUB);
  wash(s, 15.762, 4.33, 2, 2, { r: 1 });
  wash(s, 15.762, 7.684, 2, 2, { r: 1 });
  wash(s, 15.698, 11.038, 2, 2, { r: 1 });
  icon(s, "speaker", 16.391, 8.334, 0.8, 0.7, fill(WHITE));
  icon(s, "video", 16.319, 11.784, 0.8, 0.6, fill(WHITE));
  icon(s, "camera", 16.391, 4.951, 0.8, 0.7, fill(WHITE), NAVY);
  text(s, 18.825, 8.25, 4.421, 1.778, "Lorem Ipsum has been the industry's standard dummy text ever since.", BODY);
  text(s, 18.825, 7.684, 5.083, 0.572, "Service One", SUB);
  text(s, 18.824, 11.644, 4.421, 1.778, "Lorem Ipsum has been the industry's standard dummy text ever since.", BODY);
  text(s, 18.824, 11.079, 5.083, 0.572, "Service One", SUB);
}

function slide15(s) {
  deco(s, [[4.527, -9.448, "glow", 15.327, TEAL], [-4.889, 7.136, "glow", 14.071, PURPLE], 
    [1.14, 1.416, "rings"], [9.583, 2.488, "dot"], [2.437, 3.253, "dot"], [1.066, 7.851, "dot"]]);
  photo(s, 3.528, 4.137, 9.384, 9.384, "ellipse");
  deco(s, [[20.866, 1.991, "glow", 15.327, TEAL], [22.087, -4.658, "halo", 8.795]]);
  socialStrip(s);
  pillButton(s, 22.087, 12.528, "next");
  pageNumber(s, "15");
  text(s, 12.305, 2.238, 9.612, 1.531, "Our Service", HEAD);
  shape(s, "ellipse", 13.308, 7.125, 2, 2, fill(WHITE));
  icon(s, "camera", 13.937, 7.746, 0.8, 0.7, WASH, WHITE);
  text(s, 13.456, 10.199, 4.989, 1.778, "Lorem Ipsum has been the industry's standard dummy text ever since.", BODY);
  shape(s, "ellipse", 19.523, 7.124, 2, 2, fill(WHITE));
  icon(s, "video", 20.152, 7.821, 0.8, 0.6, WASH);
  text(s, 13.409, 9.536, 5.083, 0.572, "Service One", SUB);
  text(s, 19.523, 10.199, 4.989, 1.778, "Lorem Ipsum has been the industry's standard dummy text ever since.", BODY);
  text(s, 19.476, 9.536, 5.083, 0.572, "Service Two", SUB);
  text(s, 13.333, 5.025, 8.164, 1.198, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown.", BODY);
}

function slide16(s) {
  deco(s, [[17.158, -6.768, "glow", 14.071, PURPLE], [20.342, 6.185, "rings"], [21.097, 14.223, "dot"]]);
  photo(s, 2.328, 4.577, 6.702, 4.943, "roundRect", 0);
  photo(s, 9.931, 4.588, 6.702, 4.943, "roundRect", 0);
  photo(s, 17.636, 4.577, 6.702, 4.943, "roundRect", 0);
  deco(s, [[-7.047, 7.611, "glow", 15.327, TEAL], [-5.514, 11.065, "halo", 8.795]]);
  socialStrip(s);
  pageNumber(s, "16");
  text(s, 8.589, 2.312, 9.488, 1.531, "Our Service", HEAD_C);
  pillButton(s, 22.152, 3.034, "next");
  text(s, 2.328, 11.018, 4.421, 1.778, "Lorem Ipsum has been the industry's standard dummy text ever since.", BODY);
  text(s, 2.296, 10.446, 5.083, 0.572, "Service One", SUB);
  text(s, 9.931, 11.029, 4.421, 1.778, "Lorem Ipsum has been the industry's standard dummy text ever since.", BODY);
  text(s, 9.899, 10.456, 5.083, 0.572, "Service Two", SUB);
  text(s, 17.636, 11.018, 4.421, 1.778, "Lorem Ipsum has been the industry's standard dummy text ever since.", BODY);
  text(s, 17.604, 10.446, 5.083, 0.572, "Service Three", SUB);
  shape(s, "ellipse", 2.28, 8.333, 2, 2, fill(WHITE));
  shape(s, "ellipse", 9.894, 8.344, 2, 2, fill(WHITE));
  shape(s, "ellipse", 17.614, 8.327, 2, 2, fill(WHITE));
  icon(s, "speaker", 10.522, 8.994, 0.8, 0.7, WASH);
  icon(s, "video", 18.235, 9.073, 0.8, 0.6, WASH);
  icon(s, "camera", 2.909, 8.954, 0.8, 0.7, WASH, WHITE);
}

function slide17(s) {
  deco(s, [[-0.139, 9.099, "glow", 14.071, PURPLE], [0.473, 3.911, "rings"], [0.198, 8.715, "dot"], 
    [2.019, 12.048, "dot"]]);
  photo(s, 8.186, -0.06, 10.403, 15.06, "rect");
  socialStrip(s);
  pillButton(s, 4.586, 2.263, "next");
  pageNumber(s, "17");
  text(s, 19.631, 2.203, 6.327, 2.962, "Break Slide", HEAD);
  wash(s, 13.1, 8.002, 11.136, 4.916, { r: 1.251, opacity: 49.803 });
  text(s, 14.346, 8.979, 8.498, 3.139, "“This is a quote. Words full of wisdom that someone important said.”", { font: TITLE_FONT, size: 45, italic: true, valign: "middle", margin: 7.2 });
}

function slide18(s) {
  deco(s, [[17.239, 7.112, "glow", 15.327, TEAL], [5.332, -8.435, "glow", 14.071, PURPLE], 
    [15.794, 1.336, "rings"], [24.237, 2.408, "dot"], [17.091, 3.173, "dot"], [25, 9.243, "dot"]]);
  photo(s, 13.046, 4.287, 8.887, 8.887, "ellipse");
  deco(s, [[-7.417, 6.366, "glow", 15.327, TEAL]]);
  text(s, 2.334, 2.54, 10.137, 1.531, "Xander Team", HEAD);
  deco(s, [[-4.314, 10.658, "halo", 8.795]]);
  socialStrip(s);
  pillButton(s, 9.122, 11.593, "next");
  pageNumber(s, "18");
  text(s, 4.444, 7.718, 5.595, 2.939, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book.", BODY);
  text(s, 4.444, 5.623, 2.018, 0.572, "Grey John", SUB);
  text(s, 4.444, 6.368, 1.445, 0.488, "Manager", { italic: true });
}

function slide19(s) {
  deco(s, [[-4.662, 9.081, "glow", 14.071, PURPLE], [-3.244, -9.107, "glow", 15.327, TEAL], 
    [3.987, 3.837, "rings"], [12.431, 4.91, "dot"], [5.285, 5.675, "dot"], [3.914, 10.272, "dot"]]);
  photo(s, 12.336, 0.49, 6.782, 7.014, "roundRect", 0);
  photo(s, 12.336, 7.526, 6.782, 7.014, "roundRect", 0);
  deco(s, [[17.79, -7.801, "glow", 15.327, TEAL], [24.64, 2.333, "halo", 10.821]]);
  socialStrip(s);
  pillButton(s, 2.379, 12.519, "next");
  pageNumber(s, "19");
  text(s, 2.476, 1.785, 8.418, 2.962, "Our Best Team", HEAD);
  wash(s, 18.413, 1.776, 5.712, 3.686);
  shape(s, "rect", 18.414, 5.464, 5.712, 1.292, fill(WHITE));
  icon(s, "facebook", 22.554, 5.833, 0.291, 0.553, WASH);
  icon(s, "linkedin", 20.924, 5.833, 0.431, 0.412, WASH);
  icon(s, "twitter", 19.218, 5.904, 0.507, 0.412, WASH);
  text(s, 18.974, 2.317, 2.018, 0.572, "Grey John", SUB);
  text(s, 18.974, 3.653, 5.451, 1.2, "Lorem ipsum dolor sit amet, consectetur adipiscing elit.", NOTE);
  text(s, 18.974, 3.063, 1.445, 0.488, "Manager", { italic: true });
  wash(s, 18.413, 8.812, 5.712, 3.686);
  shape(s, "rect", 18.414, 12.5, 5.712, 1.292, fill(WHITE));
  icon(s, "facebook", 22.554, 12.869, 0.291, 0.553, WASH);
  icon(s, "linkedin", 20.924, 12.869, 0.431, 0.412, WASH);
  icon(s, "twitter", 19.218, 12.94, 0.507, 0.412, WASH);
  text(s, 18.974, 9.353, 1.75, 0.572, "Renaldo ", SUB);
  text(s, 18.974, 10.689, 5.451, 1.2, "Lorem ipsum dolor sit amet, consectetur adipiscing elit.", NOTE);
  text(s, 18.974, 10.099, 1.445, 0.488, "Manager", { italic: true });
}

function slide20(s) {
  deco(s, [[-5.84, 7.965, "glow", 15.327, TEAL], [16.857, -5.952, "glow", 14.071, PURPLE], 
    [20.342, 6.185, "rings"], [21.097, 14.223, "dot"], [21.926, 7.785, "dot"]]);
  photo(s, 2.399, 5.142, 6.782, 7.014, "roundRect", 0);
  photo(s, 10.044, 5.142, 6.782, 7.014, "roundRect", 0);
  photo(s, 17.591, 5.142, 6.782, 7.014, "roundRect", 0);
  deco(s, [[-5.778, -3.257, "halo", 8.795]]);
  socialStrip(s);
  pageNumber(s, "20");
  text(s, 9.085, 2.328, 9.488, 1.531, "Our Best Team", HEAD);
  pillButton(s, 2.556, 3.628, "next");
  wash(s, 2.293, 10.307, 6.888, 1.853);
  text(s, 4.7, 10.618, 2.018, 0.572, "Grey John", SUB);
  text(s, 4.945, 11.262, 1.445, 0.488, "Manager", { italic: true });
  wash(s, 9.982, 10.307, 6.874, 1.853);
  text(s, 12.345, 10.618, 2.018, 0.572, "Grey John", SUB);
  text(s, 12.591, 11.262, 1.445, 0.488, "Manager", { italic: true });
  wash(s, 17.551, 10.307, 6.888, 1.853);
  text(s, 19.892, 10.618, 2.018, 0.572, "Grey John", SUB);
  text(s, 20.137, 11.262, 1.445, 0.488, "Manager", { italic: true });
}

function slide21(s) {
  deco(s, [[2.374, -1.681, "glow", 15.327, TEAL]]);
  photo(s, 7.479, 1.567, 19.19, 5.99, "rect");
  deco(s, [[-2.379, 6.636, "glow", 14.071, PURPLE], [18.821, 9.23, "glow", 15.327, TEAL], 
    [-2.492, 6.828, "rings"], [5.952, 7.901, "dot"], [0.93, 7.257, "dot"], [6.912, 13.006, "dot"], 
    [22.606, 11.875, "halo", 8.795]]);
  socialStrip(s);
  pillButton(s, 2.643, 3.395, "next");
  pageNumber(s, "21");
  text(s, 8.489, 7.994, 7.603, 2.962, "Gallery of Product", HEAD);
  text(s, 16.534, 9.723, 5.595, 2.939, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book.", BODY);
}

function slide22(s) {
  deco(s, [[-8.095, 2.421, "glow", 15.327, TEAL], [7.916, -7.127, "glow", 14.071, PURPLE], 
    [-3.495, 9.852, "halo", 8.795]]);
  photo(s, 3.349, 8.259, 5.142, 4.923, "rect");
  photo(s, 9.032, 8.259, 5.142, 4.923, "rect");
  photo(s, 14.714, 3.598, 5.142, 4.923, "rect");
  deco(s, [[19.328, 6.185, "rings"], [20.083, 14.223, "dot"], [20.911, 7.785, "dot"]]);
  socialStrip(s);
  pillButton(s, 20.342, 2.862, "next");
  pageNumber(s, "22");
  text(s, 2.612, 2.485, 8.279, 2.962, "Gallery of Product", HEAD);
  text(s, 3.311, 6.185, 5.181, 1.781, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt", NOTE);
  text(s, 14.831, 10.234, 4.295, 2.359, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt", NOTE);
  text(s, 14.753, 9.201, 2.712, 0.572, "Lorem Ipsum", SUB);
  text(s, 8.993, 6.185, 5.181, 1.781, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt", NOTE);
}

function slide23(s) {
  deco(s, [[16.311, 7.523, "glow", 15.327, TEAL], [-0.695, 7.003, "glow", 14.071, PURPLE], 
    [18.033, -9.451, "glow", 15.327, TEAL], [-3.111, 9.826, "rings"]]);
  text(s, 2.725, 2.225, 8.802, 1.531, "Best Pricing", HEAD);
  deco(s, [[-3.63, 9.338, "rings"], [4.814, 10.411, "dot"], [0.466, 9.585, "dot"], 
    [22.253, -6.458, "halo", 10.989]]);
  socialStrip(s);
  pillButton(s, 2.971, 5.472, "next");
  pageNumber(s, "23");
  wash(s, 11.23, 4.558, 6.364, 8.067, { r: 1.171 });
  text(s, 11.789, 5.616, 5.247, 0.909, "Starter", { size: 48, bold: true, align: "center" });
  text(s, 12.605, 7.276, 3.614, 1.717, "$199", { size: 96, bold: true, align: "center" });
  text(s, 12.722, 9.669, 3.687, 1.649, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s.", { align: "center", valign: "middle", margin: 3.6 });
  wash(s, 18.169, 4.531, 6.364, 8.067, { r: 1.171 });
  text(s, 18.728, 5.588, 5.247, 0.909, "Premium", { size: 48, bold: true, align: "center" });
  text(s, 19.544, 7.248, 3.614, 1.717, "$299", { size: 96, bold: true, align: "center" });
  text(s, 19.661, 9.642, 3.687, 1.649, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s.", { align: "center", valign: "middle", margin: 3.6 });
  text(s, 5.683, 6.891, 4.728, 3.52, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book.", BODY);
}

function slide24(s) {
  deco(s, [[19.519, -2.508, "glow", 15.327, TEAL], [-7.279, -9.062, "glow", 15.327, TEAL], 
    [-5.743, 7.499, "glow", 14.071, PURPLE], [21.938, 5.708, "rings"], [24.647, 5.933, "dot"], 
    [22.538, 12.486, "dot"], [-4.998, 11.504, "halo", 8.795]]);
  socialStrip(s);
  pillButton(s, 18.963, 11.467, "next");
  pageNumber(s, "24");
  wash(s, 2.641, 4.865, 6.364, 8.067, { r: 1.171 });
  text(s, 3.199, 5.923, 5.247, 0.909, "Starter", { size: 48, bold: true, align: "center" });
  text(s, 4.015, 7.583, 3.614, 1.717, "$199", { size: 96, bold: true, align: "center" });
  text(s, 4.132, 9.976, 3.687, 1.649, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s.", { align: "center", valign: "middle", margin: 3.6 });
  wash(s, 9.579, 4.838, 6.364, 8.067, { r: 1.171 });
  text(s, 10.138, 5.896, 5.247, 0.909, "Premium", { size: 48, bold: true, align: "center" });
  text(s, 10.954, 7.556, 3.614, 1.717, "$299", { size: 96, bold: true, align: "center" });
  text(s, 11.071, 9.949, 3.687, 1.649, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s.", { align: "center", valign: "middle", margin: 3.6 });
  text(s, 16.792, 5.451, 4.728, 3.52, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book.", BODY);
  text(s, 9.004, 2.032, 8.802, 1.531, "Best Pricing", HEAD_C);
}

function slide25(s) {
  deco(s, [[15.883, -7.664, "glow", 15.327, TEAL], [-2.059, 7.552, "glow", 14.071, PURPLE], 
    [-2.038, -7.646, "glow", 15.327, TEAL], [4.624, 3.516, "rings"], [13.068, 4.588, "dot"], 
    [5.922, 5.353, "dot"], [4.551, 9.951, "dot"], [-5.378, -5.393, "halo", 10.821]]);
  photo(s, 12.4, 5.684, 14.246, 6.057, "rect");
  socialStrip(s);
  pillButton(s, 2.364, 11.406, "next");
  pageNumber(s, "25");
  text(s, 16.211, 2.933, 8.279, 1.531, "Break Slide", HEAD);
  text(s, 6.891, 6.535, 5.647, 4.214, "“This is a quote. Words full of wisdom that someone important said.”", { font: TITLE_FONT, size: 45, italic: true, valign: "middle", margin: 7.2 });
}

function slide26(s) {
  deco(s, [[19.005, 2.854, "glow", 15.327, TEAL], [-5.227, 11.515, "halo", 8.795], 
    [-7.336, -2.691, "glow", 14.071, PURPLE], [20.203, -3.465, "rings"], [19.928, 1.339, "dot"], 
    [25.58, 6.057, "dot"]]);
  socialStrip(s);
  pillButton(s, 21.983, 8.887, "next");
  pageNumber(s, "26");
  columnChart(s, 11.186, 3.666, 8.387, 5.519);
  text(s, 3.319, 6.474, 3.859, 0.572, "Lorem Ipsum Dolor", { size: 28, bold: true, color: PALE });
  text(s, 3.362, 7.175, 5.774, 1.781, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt", { color: PALE, ls: 1.5 });
  text(s, 3.276, 9.752, 3.859, 0.572, "Lorem Ipsum Dolor", { size: 28, bold: true, color: PALE });
  text(s, 3.319, 10.453, 5.774, 1.781, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt", { color: PALE, ls: 1.5 });
  text(s, 3.099, 1.962, 10.031, 2.962, "Our Company in Number", HEAD);
  shape(s, "donut", 11.397, 9.811, 3.8, 3.8, fill(WHITE, 10.98), { rectRadius: 0.542 });
  shape(s, "blockArc", 11.397, 9.811, 3.8, 3.8, fill(BLEND), { angleRange: [244.218, 0], arcThicknessRatio: 0.293, rotate: -90 });
  text(s, 12.557, 11.303, 1.552, 0.858, "80 %", { size: 45, bold: true });
  text(s, 16.001, 10.232, 2.119, 0.572, "Chart One", SUB);
  text(s, 16.044, 10.932, 6.363, 1.781, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt", NOTE);
}

function slide27(s) {
  deco(s, [[18.61, -6.387, "glow", 15.327, TEAL], [-1.485, 8.828, "glow", 14.071, PURPLE], 
    [-5.246, 7.884, "rings"], [0.777, 8.215, "dot"], [4.887, 12.765, "dot"]]);
  socialStrip(s);
  pillButton(s, 10.351, 11.533, "next");
  pageNumber(s, "27");
  waterfall(s, 2.592, 1.277, 10.295, 7.459);
  text(s, 14.798, 6.149, 6.829, 1.778, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took.", BODY);
  icon(s, "person", 14.802, 11.756, 0.5, 1.2, fill(PURPLE));
  icon(s, "person", 15.583, 11.756, 0.5, 1.2, fill(PURPLE));
  icon(s, "person", 16.383, 11.756, 0.5, 1.2, fill(PURPLE));
  icon(s, "person", 17.197, 11.756, 0.5, 1.2, fill(PURPLE));
  icon(s, "person", 18.005, 11.756, 0.5, 1.2, fill(PURPLE));
  icon(s, "person", 18.84, 11.756, 0.5, 1.2, fill(PURPLE));
  icon(s, "person", 19.706, 11.756, 0.5, 1.2, fill(PURPLE));
  icon(s, "person", 20.516, 11.756, 0.5, 1.2, fill(WHITE, 10.98));
  icon(s, "person", 21.356, 11.756, 0.5, 1.2, fill(WHITE, 10.98));
  icon(s, "person", 22.203, 11.756, 0.5, 1.2, fill(WHITE, 10.98));
  text(s, 23.194, 11.927, 1.394, 0.858, "70%", { size: 45, bold: true });
  text(s, 14.747, 10.877, 2.527, 0.572, "Lorem Ipsum", SUB);
  icon(s, "person", 14.798, 9.382, 0.5, 1.2, fill(TEAL));
  icon(s, "person", 15.579, 9.382, 0.5, 1.2, fill(TEAL));
  icon(s, "person", 16.379, 9.382, 0.5, 1.2, fill(TEAL));
  icon(s, "person", 17.193, 9.382, 0.5, 1.2, fill(TEAL));
  icon(s, "person", 18.001, 9.382, 0.5, 1.2, fill(TEAL));
  icon(s, "person", 18.836, 9.382, 0.5, 1.2, fill(TEAL));
  icon(s, "person", 19.702, 9.382, 0.5, 1.2, fill(TEAL));
  icon(s, "person", 20.512, 9.382, 0.5, 1.2, fill(TEAL));
  icon(s, "person", 21.352, 9.382, 0.5, 1.2, fill(WHITE, 10.98));
  icon(s, "person", 22.199, 9.382, 0.5, 1.2, fill(WHITE, 10.98));
  text(s, 23.19, 9.553, 1.394, 0.858, "80%", { size: 45, bold: true });
  text(s, 14.744, 8.502, 2.527, 0.572, "Lorem Ipsum", SUB);
  text(s, 14.798, 2.044, 10.031, 2.962, "Our Company in Number", HEAD);
  text(s, 8.6, 9.367, 4.702, 1.198, "Lorem Ipsum has been the industry's standard dummy.", BODY);
  deco(s, [[22.199, -4.398, "halo", 8.795]]);
}

function slide28(s) {
  deco(s, [[19.387, -2.499, "glow", 15.327, TEAL], [-7.614, -5.608, "glow", 15.327, TEAL], 
    [-5.652, 5.698, "glow", 14.071, PURPLE], [21.938, 5.708, "rings"], [24.647, 5.933, "dot"], 
    [18.953, 10.099, "dot"], [22.538, 12.486, "dot"], [-4.998, 11.504, "halo", 8.795]]);
  socialStrip(s);
  pillButton(s, 21.026, 2.589, "next");
  pageNumber(s, "28");
  stepFrame(s, 3.379, 6.988, 3.702, 3.708, TEAL);
  stepTab(s, 3.781, 6.518, 3.702, 1.24, TEAL);
  stepTab(s, 8.413, 9.824, 3.702, 1.245, PURPLE);
  stepTab(s, 12.976, 6.518, 3.702, 1.24, TEAL);
  stepTab(s, 17.67, 9.95, 3.702, 1.245, PURPLE);
  stepFrame(s, 7.971, 6.988, 3.702, 3.708, PURPLE);
  stepFrame(s, 12.562, 6.988, 3.702, 3.708, TEAL);
  stepFrame(s, 17.148, 6.988, 3.702, 3.708, PURPLE);
  arcArrow(s, 5.485, 4.291, 4.678, 2.365, false);
  arcArrow(s, 14.709, 4.291, 4.678, 2.365, false);
  text(s, 8.969, 10.018, 3.405, 0.735, "Lorem Ipsum", { size: 28, bold: true, valign: "middle", ls: 1.5, margin: 3.6 });
  text(s, 3.887, 8.324, 2.906, 1.262, "Lorem Ipsum has been the industry's standard dummy", { valign: "middle", margin: 3.6 });
  text(s, 8.454, 7.89, 2.906, 1.262, "Lorem Ipsum has been the industry's standard dummy", { valign: "middle", margin: 3.6 });
  text(s, 13.032, 8.281, 2.906, 1.262, "Lorem Ipsum has been the industry's standard dummy", { valign: "middle", margin: 3.6 });
  text(s, 17.649, 7.838, 2.906, 1.262, "Lorem Ipsum has been the industry's standard dummy", { valign: "middle", margin: 3.6 });
  text(s, 18.232, 10.109, 3.405, 0.735, "Lorem Ipsum", { size: 28, bold: true, valign: "middle", ls: 1.5, margin: 3.6 });
  text(s, 4.263, 6.697, 3.405, 0.735, "Lorem Ipsum", { size: 28, bold: true, valign: "middle", ls: 1.5, margin: 3.6 });
  text(s, 13.527, 6.691, 3.405, 0.735, "Lorem Ipsum", { size: 28, bold: true, valign: "middle", ls: 1.5, margin: 3.6 });
  arcArrow(s, 9.969, 10.841, 4.678, 2.365, true);
  text(s, 7.478, 4.734, 0.935, 1.454, "1", { size: 60, bold: true, valign: "middle", ls: 1.5, margin: 3.6 });
  text(s, 11.84, 11.158, 0.935, 1.454, "2", { size: 60, bold: true, valign: "middle", ls: 1.5, margin: 3.6 });
  text(s, 16.678, 4.746, 0.935, 1.454, "3", { size: 60, bold: true, valign: "middle", ls: 1.5, margin: 3.6 });
  text(s, 8.944, 1.701, 8.802, 1.531, "Infographic", HEAD);
}

function slide29(s) {
  deco(s, [[-4.404, -6.75, "glow", 15.327, TEAL], [-1.285, 5.956, "glow", 14.071, PURPLE], 
    [18.821, 1.323, "glow", 15.327, TEAL], [-2.55, 9.136, "rings"]]);
  text(s, 2.725, 2.739, 8.802, 1.531, "Infographic", HEAD);
  deco(s, [[-3.069, 8.648, "rings"], [5.374, 9.721, "dot"], [1.026, 8.895, "dot"], [21.181, 10.95, "halo", 10.989]]);
  socialStrip(s);
  pillButton(s, 9.939, 10.652, "next");
  pageNumber(s, "29");
  text(s, 6.951, 5.956, 5.595, 2.939, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book.", BODY);
  wash(s, 15.392, 0.01, 2.236, 13.431, { opacity: 49.803 });
  text(s, 15.392, 0.01, 2.236, 13.431, ".", { size: 14, color: "000000", valign: "middle", margin: 0 });
  shape(s, "ellipse", 14.763, 0.825, 3.495, 3.495, fill(WHITE));
  shape(s, "ellipse", 14.761, 9.127, 3.5, 3.5, fill(WHITE));
  shape(s, "ellipse", 14.761, 4.973, 3.5, 3.5, fill(WHITE));
  text(s, 19.329, 1.018, 1.821, 0.538, "Info one", { size: 28, bold: true, margin: 2.4 });
  text(s, 19.329, 9.317, 2.236, 0.538, "Info three", { size: 28, bold: true, margin: 2.4 });
  text(s, 19.361, 5.199, 1.821, 0.538, "Info two", { size: 28, bold: true, margin: 2.4 });
  text(s, 19.329, 1.564, 4.555, 2.36, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt", NOTE);
  text(s, 19.329, 5.866, 4.555, 2.36, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt", NOTE);
  text(s, 19.298, 9.893, 4.555, 2.36, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt", NOTE);
  icon(s, "archive", 16.111, 10.477, 0.8, 0.8, WASH, WHITE);
  icon(s, "speaker", 16.111, 6.373, 0.8, 0.7, WASH);
  icon(s, "camera", 16.111, 2.196, 0.8, 0.7, WASH, WHITE);
}

function slide30(s) {
  deco(s, [[13.282, 6.519, "glow", 15.327, TEAL], [-5.156, 7.096, "halo", 13.15], 
    [-7.695, 7.361, "glow", 14.071, PURPLE], [21.158, -5.379, "rings"], [21.259, 1.596, "dot"], 
    [24.624, 4.012, "dot"]]);
  socialStrip(s);
  pillButton(s, 22.822, 8.669, "next");
  pageNumber(s, "30");
  arch(s, 3.326, 5.086, 5.836, 8.531);
  shape(s, "ellipse", 4.733, 3.469, 3, 3, fill(WHITE));
  wash(s, 5.366, 4.147, 1.7, 1.7, { r: 0.85 });
  arch(s, 9.429, 5.086, 5.836, 8.531);
  arch(s, 15.552, 5.038, 5.836, 8.531);
  shape(s, "ellipse", 10.857, 3.469, 3, 3, fill(WHITE));
  shape(s, "ellipse", 16.96, 3.469, 3, 3, fill(WHITE));
  text(s, 4.201, 8.344, 4.085, 0.572, "Lorem Ipsum Dolor", { size: 28, bold: true, align: "center" });
  text(s, 4.301, 9.232, 4.197, 2.359, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt", { align: "center", ls: 1.5 });
  text(s, 10.272, 8.328, 4.085, 0.572, "Lorem Ipsum Dolor", { size: 28, bold: true, align: "center" });
  text(s, 10.373, 9.217, 4.197, 2.359, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt", { align: "center", ls: 1.5 });
  text(s, 16.33, 8.359, 4.085, 0.572, "Lorem Ipsum Dolor", { size: 28, bold: true, align: "center" });
  text(s, 16.431, 9.247, 4.197, 2.359, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt", { align: "center", ls: 1.5 });
  wash(s, 11.498, 4.127, 1.7, 1.7, { r: 0.85 });
  wash(s, 17.569, 4.12, 1.7, 1.7, { r: 0.85 });
  icon(s, "speaker", 11.913, 4.521, 0.928, 0.825, fill(WHITE));
  icon(s, "video", 5.771, 4.618, 0.928, 0.825, fill(WHITE));
  icon(s, "cloud", 17.996, 4.451, 0.928, 0.825, fill(WHITE), NAVY);
  text(s, 8.845, 1.323, 8.802, 1.531, "Infographic", HEAD);
}

function slide31(s) {
  photo(s, -0.026, -0.06, 11.136, 15.06, "rect");
  deco(s, [[10.909, 6.32, "glow", 14.071, PURPLE], [20.203, -3.465, "rings"], [19.928, 1.339, "dot"], 
    [25.58, 6.057, "dot"]]);
  socialStrip(s);
  pillButton(s, 22.023, 11.623, "next");
  pageNumber(s, "31");
  wash(s, 7.537, 7.954, 11.136, 4.916, { r: 1.251, opacity: 49.803 });
  text(s, 8.782, 8.931, 8.498, 3.139, "“This is a quote. Words full of wisdom that someone important said.”", { font: TITLE_FONT, size: 45, italic: true, valign: "middle", margin: 7.2 });
  text(s, 12.196, 3.425, 8.279, 1.531, "Break Slide", HEAD);
}

function slide32(s) {
  photo(s, 3.528, 4.137, 9.384, 9.384, "ellipse");
  deco(s, [[4.95, -8.456, "glow", 15.327, TEAL], [-3.086, 5.971, "glow", 14.071, PURPLE], 
    [20.362, 1.556, "glow", 15.327, TEAL], [1.14, 1.416, "rings"], [9.583, 2.488, "dot"], [2.437, 3.253, "dot"], 
    [1.066, 7.851, "dot"], [22.087, -4.658, "halo", 8.795]]);
  socialStrip(s);
  pillButton(s, 11.629, 12.522, "next");
  pageNumber(s, "32");
  icon(s, "phone", 14.405, 5.306, 0.8, 0.8, fill(WHITE));
  icon(s, "addressBook", 19.962, 5.258, 0.8, 0.8, fill(WHITE), NAVY);
  icon(s, "envelope", 19.962, 9.574, 0.8, 0.6, fill(WHITE), NAVY);
  icon(s, "globe", 14.327, 9.374, 0.8, 0.8, fill(WHITE), NAVY);
  text(s, 14.327, 6.34, 2.261, 0.572, "Phone", SUB);
  text(s, 19.962, 6.34, 2.574, 0.572, "Address", SUB);
  text(s, 14.327, 10.673, 2.574, 0.572, "Website", SUB);
  text(s, 19.962, 10.673, 2.792, 0.572, "Email", SUB);
  text(s, 14.327, 6.985, 4.348, 1.198, "Lorem ipsum dolor sit amet, consectetur adipiscing elit", NOTE);
  text(s, 14.327, 11.248, 4.348, 1.198, "Lorem ipsum dolor sit amet, consectetur adipiscing elit", NOTE);
  text(s, 19.962, 6.98, 4.348, 1.198, "Lorem ipsum dolor sit amet, consectetur adipiscing elit", NOTE);
  text(s, 19.962, 11.243, 4.348, 1.198, "Lorem ipsum dolor sit amet, consectetur adipiscing elit", NOTE);
  text(s, 12.305, 2.238, 9.612, 1.531, "Contact Us", HEAD);
}

function slide33(s) {
  deco(s, [[16.459, -7.827, "glow", 15.327, TEAL], [-0.462, 8.641, "glow", 14.071, PURPLE], 
    [-6.451, -9.004, "glow", 15.327, TEAL], [7.677, 3.377, "rings"], [16.121, 4.45, "dot"], 
    [8.975, 5.215, "dot"], [7.604, 9.812, "dot"]]);
  photo(s, 12.4, 5.513, 14.246, 6.057, "rect");
  text(s, 12.346, 1.723, 9.612, 1.531, "Contact Us", HEAD);
  deco(s, [[20.863, 2.678, "halo", 10.821]]);
  socialStrip(s);
  pillButton(s, 19.169, 12.476, "next");
  pageNumber(s, "33");
  text(s, 2.389, 2.574, 2.261, 0.572, "Phone", SUB);
  text(s, 2.416, 7.85, 2.574, 0.572, "Address", SUB);
  text(s, 2.389, 5.247, 2.574, 0.572, "Website", SUB);
  text(s, 2.389, 10.522, 2.792, 0.572, "Email", SUB);
  text(s, 2.389, 3.219, 4.348, 1.198, "Lorem ipsum dolor sit amet, consectetur adipiscing elit", NOTE);
  text(s, 2.389, 5.822, 4.348, 1.198, "Lorem ipsum dolor sit amet, consectetur adipiscing elit", NOTE);
  text(s, 2.416, 8.49, 4.348, 1.198, "Lorem ipsum dolor sit amet, consectetur adipiscing elit", NOTE);
  text(s, 2.389, 11.092, 4.348, 1.198, "Lorem ipsum dolor sit amet, consectetur adipiscing elit", NOTE);
}

function slide34(s) {
  deco(s, [[-4.271, 7.31, "glow", 14.071, PURPLE]]);
  photo(s, 5.296, 4.724, 7.29, 6.828, "rect");
  deco(s, [[11.292, -12.759, "glow", 15.327, TEAL]]);
  wash(s, 12.586, 3.024, 12.335, 9.852, { r: 1.631, opacity: 49.803 });
  deco(s, [[-5.395, 9.11, "rings"], [0.629, 9.441, "dot"], [4.738, 13.991, "dot"]]);
  socialStrip(s);
  pillButton(s, 2.383, 4.751, "next");
  pageNumber(s, "34");
  deco(s, [[22.817, -5.425, "halo", 8.795]]);
  icon(s, "phone", 14.129, 4.082, 0.8, 0.8, fill(WHITE));
  icon(s, "addressBook", 19.686, 4.035, 0.8, 0.8, fill(WHITE), NAVY);
  icon(s, "envelope", 19.686, 8.35, 0.8, 0.6, fill(WHITE), NAVY);
  icon(s, "globe", 14.051, 8.15, 0.8, 0.8, fill(WHITE), NAVY);
  text(s, 14.051, 5.117, 2.261, 0.572, "Phone", SUB);
  text(s, 19.686, 5.117, 2.574, 0.572, "Address", SUB);
  text(s, 14.051, 9.45, 2.574, 0.572, "Website", SUB);
  text(s, 19.686, 9.45, 2.792, 0.572, "Email", SUB);
  text(s, 14.051, 5.761, 4.348, 1.198, "Lorem ipsum dolor sit amet, consectetur adipiscing elit", NOTE);
  text(s, 14.051, 10.024, 4.348, 1.198, "Lorem ipsum dolor sit amet, consectetur adipiscing elit", NOTE);
  text(s, 19.686, 5.757, 4.348, 1.198, "Lorem ipsum dolor sit amet, consectetur adipiscing elit", NOTE);
  text(s, 19.686, 10.02, 4.348, 1.198, "Lorem ipsum dolor sit amet, consectetur adipiscing elit", NOTE);
  text(s, 3.284, 1.613, 9.612, 1.531, "Contact Us", HEAD);
}

function slide35(s) {
  deco(s, [[-4.677, 3.236, "glow", 15.327, TEAL], [1.326, -5.729, "glow", 14.071, PURPLE], 
    [2.05, 1.556, "rings"], [10.493, 2.629, "dot"], [3.347, 3.394, "dot"], [10.909, 9.006, "dot"]]);
  photo(s, -3.11, 4.974, 14.019, 14.019, "ellipse");
  deco(s, [[19.539, 5.053, "glow", 15.327, TEAL]]);
  text(s, 14.254, 5.321, 10.609, 2.794, "THANK’S", { font: TITLE_FONT, size: 160, bold: true });
  text(s, 14.254, 7.82, 8.228, 1.01, "See You Next TIme", { font: TITLE_FONT, size: 54 });
  deco(s, [[20.944, 9.006, "halo", 8.795]]);
  text(s, 14.396, 9.725, 5.584, 1.778, "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown.", BODY);
  socialStrip(s);
  pillButton(s, 14.396, 2.641, "back");
  pageNumber(s, "35");
}

const BUILDERS = [
  slide01,
  slide02,
  slide03,
  slide04,
  slide05,
  slide06,
  slide07,
  slide08,
  slide09,
  slide10,
  slide11,
  slide12,
  slide13,
  slide14,
  slide15,
  slide16,
  slide17,
  slide18,
  slide19,
  slide20,
  slide21,
  slide22,
  slide23,
  slide24,
  slide25,
  slide26,
  slide27,
  slide28,
  slide29,
  slide30,
  slide31,
  slide32,
  slide33,
  slide34,
  slide35,
];

function main() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'XANDER', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'XANDER';
  pptx.title = 'Xander Home Security';

  BUILDERS.forEach(function (build) {
    const s = pptx.addSlide();
    s.background = { color: NAVY };
    build(s);
  });

  const out = path.join(__dirname, '052922ad-bdf2-4fa3-96b4-94fb3dcf4f14_grok_final.pptx');
  return pptx.writeFile({ fileName: out }).then(function () { console.log('wrote ' + out); });
}

main().catch(function (err) { console.error(err); process.exit(1); });
