'use strict';
/*
 * Standalone re-creation of the "Krishna said" presentation with PptxGenJS.
 * 30 slides, 13.333 x 7.5 in.  Raster artwork in the source deck is replaced
 * by programmatic placeholders / native-shape approximations.
 */
const pptxgen = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ *
 * Design tokens
 * ------------------------------------------------------------------ */
const W = 13.333;
const H = 7.5;

const C = {
  blue:   '01A4E7',  // accent1
  purple: '8D5FE6',  // accent2
  teal:   '01D9CC',  // accent3
  pink:   'FD7BFC',  // accent4
  blue2:  '0084DE',  // accent5
  blue3:  '016AA1',  // accent6
  dark:   '2C2C2C',  // tx1
  white:  'FFFFFF',
  navy:   '354552',  // title colour on the photo slides
  gray:   '969696',  // tx1 lum 50 / off 50 -> caption grey
  page:   'F4F4F4',  // tx1 @5% over white -> photo-slide backdrop
  darkBlue: '005274' // accent1 lum 50
};

const F = { head: 'Oswald Regular', body: 'Lato' };

/** Blend `fg` over opaque `bg` -- emulates OOXML solid fill + alpha. */
function mix(fg, bg, alpha) {
  const p = (h, i) => parseInt(h.substr(i * 2, 2), 16);
  let out = '';
  for (let i = 0; i < 3; i++) {
    const v = Math.round(p(fg, i) * alpha + p(bg, i) * (1 - alpha));
    out += v.toString(16).padStart(2, '0');
  }
  return out.toUpperCase();
}

const FAINT_WHITE_ON_BLUE = mix(C.white, C.blue, 0.20);  // FFFFFF @20%
const FAINT_DARK_ON_WHITE = mix(C.dark, C.white, 0.10);  // 2C2C2C @10%
const BAR_REST_ON_BLUE    = mix(C.white, C.blue, 0.10);
const BAR_REST_ON_WHITE   = mix(C.dark, C.white, 0.10);

const NONE = { type: 'none' };

/* ------------------------------------------------------------------ *
 * Primitive helpers
 * ------------------------------------------------------------------ */
function shp(s, kind, o) { s.addShape(kind, Object.assign({ line: NONE }, o)); }

function rect(s, x, y, w, h, color, extra) {
  shp(s, 'rect', Object.assign({ x, y, w, h, fill: { color } }, extra));
}
/** Full-bleed horizontal colour band. */
function band(s, y, h, color) { rect(s, 0, y, W, h, color || C.blue); }
/** Full-height vertical colour panel. */
function panel(s, x, w, color) { rect(s, x, 0, w, H, color || C.blue); }

function oval(s, x, y, w, h, color, extra) {
  shp(s, 'ellipse', Object.assign({ x, y, w, h, fill: { color } }, extra));
}
/** Open arc slice: `a` = [startDeg, endDeg]. */
function arc(s, x, y, w, h, a, color, width, extra) {
  shp(s, 'arc', Object.assign({
    x, y, w, h, angleRange: a, fill: NONE,
    line: { color, width: width || 2.5, cap: 'round' }
  }, extra));
}
/** Filled pie wedge. */
function wedge(s, x, y, w, h, a, color, extra) {
  shp(s, 'pie', Object.assign({ x, y, w, h, angleRange: a, fill: { color } }, extra));
}
function seg(s, x, y, w, h, color, width, extra) {
  shp(s, 'line', Object.assign({ x, y, w, h, line: { color, width: width || 2.5, cap: 'round' } }, extra));
}

/** Build styled runs: rich([[ 'text' ], [ 'accent', C.blue ]], {fontSize:44}) */
function rich(parts, base) {
  return parts.map(function (p) {
    return { text: p[0], options: Object.assign({}, base, p[1] ? { color: p[1] } : null, p[2]) };
  });
}

const VALIGN = { t: 'top', ctr: 'middle', b: 'bottom' };

/** Plain text box (defaults mirror a PowerPoint TextBox). */
function text(s, runs, o) {
  s.addText(runs, Object.assign({
    fontFace: F.body, fontSize: 10, color: C.dark, valign: 'top', wrap: true
  }, o));
}
/** 10 pt body copy. */
function body(s, x, y, w, h, t, o) {
  text(s, t, Object.assign({ x, y, w, h }, o));
}
/** Grey italic footnote. */
function note(s, x, y, w, h, t, o) {
  text(s, t, Object.assign({ x, y, w, h, italic: true, color: C.gray }, o));
}
/** Large pull-quote. */
function quote(s, x, y, w, h, parts, size, color) {
  text(s, rich(parts, { fontFace: F.head, fontSize: size, color: color || C.dark }),
    { x, y, w, h, fontSize: size, fontFace: F.head, color: color || C.dark });
}

/** Rounded "pill" label. `o.rotate` = 270 gives the vertical tabs. */
function pill(s, x, y, w, h, label, o) {
  o = o || {};
  s.addText(label, {
    x, y, w, h, shape: 'roundRect', rectRadius: Math.min(w, h) / 2,
    fill: { color: o.fill || C.blue }, line: NONE,
    color: o.color || C.white, fontFace: o.font || F.head, fontSize: o.size || 12,
    charSpacing: o.spc === undefined ? 3 : o.spc, bold: o.bold,
    align: 'center', valign: 'middle', margin: 0, rotate: o.rotate, wrap: false
  });
}
/** Circular badge with centred caption. */
function badge(s, x, y, d, label, o) {
  o = o || {};
  s.addText(label, {
    x, y, w: d, h: d, shape: 'ellipse', fill: { color: o.fill || C.white }, line: NONE,
    color: o.color || C.blue, fontFace: o.font || F.head, fontSize: o.size || 12,
    charSpacing: o.spc === undefined ? 3 : o.spc,
    align: 'center', valign: 'middle', margin: 0, wrap: false
  });
}

/* ------------------------------------------------------------------ *
 * Icon library -- flat vector stand-ins for the deck's freeform art
 * ------------------------------------------------------------------ */
/** Outline gem / diamond. */
function iconDiamond(s, x, y, d, color) {
  const lw = Math.max(0.75, d * 4);
  shp(s, 'custGeom', {
    x, y, w: d, h: d, fill: NONE, line: { color, width: lw },
    points: [{ x: d * 0.22, y: d * 0.14 }, { x: d * 0.78, y: d * 0.14 },
             { x: d, y: d * 0.38 }, { x: d / 2, y: d * 0.92 },
             { x: 0, y: d * 0.38 }, { close: true }]
  });
  shp(s, 'custGeom', {
    x, y, w: d, h: d, fill: NONE, line: { color, width: lw * 0.7 },
    points: [{ x: 0, y: d * 0.38 }, { x: d, y: d * 0.38 },
             { x: d * 0.78, y: d * 0.14, moveTo: true }, { x: d * 0.62, y: d * 0.38 },
             { x: d / 2, y: d * 0.92 }, { x: d * 0.38, y: d * 0.38 },
             { x: d * 0.22, y: d * 0.14 }]
  });
}
/** Outline five-point star. */
function iconStar(s, x, y, d, color) {
  shp(s, 'star5', { x, y, w: d, h: d, fill: NONE, line: { color, width: 1 } });
}
/** Solid "person" pictogram (head + shoulders + legs). */
function iconPerson(s, x, y, w, h, color) {
  const hd = w * 0.46;
  oval(s, x + (w - hd) / 2, y, hd, hd, color);
  shp(s, 'roundRect', { x, y: y + h * 0.30, w, h: h * 0.40, fill: { color }, rectRadius: w * 0.30 });
  shp(s, 'roundRect', { x: x + w * 0.10, y: y + h * 0.52, w: w * 0.32, h: h * 0.48,
    fill: { color }, rectRadius: w * 0.16 });
  shp(s, 'roundRect', { x: x + w * 0.58, y: y + h * 0.52, w: w * 0.32, h: h * 0.48,
    fill: { color }, rectRadius: w * 0.16 });
}
/** Outline arrow pointing up. */
function iconArrowUp(s, x, y, w, h, color) {
  shp(s, 'custGeom', {
    x, y, w, h, fill: NONE, line: { color, width: 1 },
    points: [{ x: w / 2, y: 0 }, { x: w, y: h * 0.38 }, { x: w * 0.68, y: h * 0.38 },
             { x: w * 0.68, y: h }, { x: w * 0.32, y: h }, { x: w * 0.32, y: h * 0.38 },
             { x: 0, y: h * 0.38 }, { close: true }]
  });
}
/** Double chevron ("guillemet") quote mark. */
function iconChevrons(s, x, y, w, h, color, flip) {
  for (let i = 0; i < 2; i++) {
    shp(s, 'custGeom', {
      x: x + i * w * 0.47, y, w: w * 0.53, h, flipH: !!flip, fill: NONE,
      line: { color, width: 3 },
      points: [{ x: w * 0.5, y: 0 }, { x: 0, y: h / 2 }, { x: w * 0.5, y: h }]
    });
  }
}
/** Grey disc holding a right-pointing arrow (process connector). */
function iconArrowCircle(s, x, y, d, color) {
  oval(s, x, y, d, d, color);
  shp(s, 'rightArrow', { x: x + d * 0.22, y: y + d * 0.28, w: d * 0.56, h: d * 0.44, fill: { color: C.white } });
}

/** Small line-art glyphs used inside coloured discs / donut holes. */
function glyph(s, kind, x, y, w, h, color) {
  const lw = 1;
  const ln = { color, width: lw };
  switch (kind) {
    case 'gear':
      shp(s, 'gear6', { x, y, w, h, fill: NONE, line: ln });
      shp(s, 'ellipse', { x: x + w * 0.28, y: y + h * 0.28, w: w * 0.44, h: h * 0.44, fill: NONE, line: ln });
      break;
    case 'target':
      for (let i = 0; i < 3; i++) {
        const k = i * 0.16;
        shp(s, 'ellipse', { x: x + w * k, y: y + h * k, w: w * (1 - 2 * k), h: h * (1 - 2 * k), fill: NONE, line: ln });
      }
      break;
    case 'dollar':                       // circled "S" with a vertical bar
      shp(s, 'ellipse', { x, y, w, h, fill: NONE, line: ln });
      shp(s, 'custGeom', { x, y, w, h, fill: NONE, line: ln,
        points: [{ x: w * 0.66, y: h * 0.34 }, { x: w * 0.36, y: h * 0.28 },
                 { x: w * 0.34, y: h * 0.46 }, { x: w * 0.66, y: h * 0.54 },
                 { x: w * 0.64, y: h * 0.72 }, { x: w * 0.34, y: h * 0.66 }] });
      seg(s, x + w * 0.50, y + h * 0.18, 0, h * 0.64, color, lw);
      break;
    case 'briefcase':
      shp(s, 'rect', { x, y: y + h * 0.25, w, h: h * 0.7, fill: NONE, line: ln });
      shp(s, 'rect', { x: x + w * 0.33, y, w: w * 0.34, h: h * 0.25, fill: NONE, line: ln });
      break;
    case 'folder':
      shp(s, 'folderCorner', { x, y: y + h * 0.1, w, h: h * 0.8, fill: NONE, line: ln, flipH: true });
      break;
    case 'calendar':
      shp(s, 'rect', { x, y: y + h * 0.12, w, h: h * 0.88, fill: NONE, line: ln });
      seg(s, x, y + h * 0.38, w, 0, color, lw);
      seg(s, x + w * 0.28, y, 0, h * 0.22, color, lw);
      seg(s, x + w * 0.72, y, 0, h * 0.22, color, lw);
      break;
    case 'people':
      oval(s, x + w * 0.10, y + h * 0.12, w * 0.26, h * 0.26, color);
      oval(s, x + w * 0.64, y + h * 0.12, w * 0.26, h * 0.26, color);
      oval(s, x + w * 0.37, y, w * 0.28, h * 0.28, color);
      shp(s, 'arc', { x, y: y + h * 0.20, w, h: h * 1.4, angleRange: [180, 360], fill: NONE, line: { color, width: 1.5 } });
      break;
    case 'note':
      shp(s, 'rect', { x: x + w * 0.08, y: y + h * 0.12, w: w * 0.84, h: h * 0.88, fill: NONE, line: ln });
      for (let i = 0; i < 3; i++) seg(s, x + w * 0.22, y + h * (0.38 + i * 0.18), w * 0.56, 0, color, lw);
      for (let i = 0; i < 3; i++) seg(s, x + w * (0.24 + i * 0.26), y, 0, h * 0.2, color, lw);
      break;
    case 'camera':
      shp(s, 'roundRect', { x, y: y + h * 0.18, w, h: h * 0.7, fill: NONE, line: ln, rectRadius: w * 0.12 });
      shp(s, 'ellipse', { x: x + w * 0.3, y: y + h * 0.35, w: w * 0.4, h: h * 0.4, fill: NONE, line: ln });
      break;
    case 'safe':
      shp(s, 'rect', { x, y, w, h, fill: NONE, line: ln });
      shp(s, 'ellipse', { x: x + w * 0.22, y: y + h * 0.22, w: w * 0.56, h: h * 0.56, fill: NONE, line: ln });
      break;
    default:
      shp(s, 'ellipse', { x, y, w, h, fill: NONE, line: ln });
  }
}
/** Coloured disc with a white glyph in the middle. */
function iconDisc(s, kind, x, y, d, discColor, glyphColor) {
  oval(s, x, y, d, d, discColor);
  glyph(s, kind, x + d * 0.26, y + d * 0.26, d * 0.48, d * 0.48, glyphColor || C.white);
}

/** Row of `n` icons, `on` of them highlighted. */
function iconRow(s, draw, x, y, step, n, on, hot, cold) {
  for (let i = 0; i < n; i++) draw(x + i * step, i < on ? hot : cold);
}

/* ------------------------------------------------------------------ *
 * Composite blocks that repeat across the deck
 * ------------------------------------------------------------------ */
/** "Main Idea" heading + paragraph (white on the blue bands). */
function ideaBlock(s, x, y, w, title, copy, o) {
  o = o || {};
  text(s, title, { x, y, w, h: 0.572, fontFace: F.head, fontSize: 28,
    color: o.color || C.white, valign: 'bottom', bold: o.bold });
  body(s, x, y + 0.603, w, o.bodyH || 0.942, copy, { color: o.color || C.white, paraSpaceAfter: 6 });
}
/** "Data text" caption above a big number. */
function dataText(s, x, y, w, label, value, o) {
  o = o || {};
  text(s, label, { x, y, w, h: 0.303, fontSize: 12, color: o.color || C.dark, align: o.align });
  text(s, value, { x, y: y + 0.335, w, h: 0.404, fontFace: F.head, fontSize: 18,
    color: o.color || C.dark, valign: 'bottom', align: o.align, bold: o.bold });
}
/** small caption / huge percentage / paragraph stack. */
function statBlock(s, x, y, w, label, value, copy, o) {
  o = o || {};
  text(s, label, { x, y, w, h: 0.303, fontSize: 12, bold: true, color: o.color || C.dark, align: o.align });
  text(s, value, { x, y: y + 0.332, w, h: 1.01, fontFace: F.head, fontSize: 54,
    color: o.color || C.dark, bold: o.bold, align: o.align });
  body(s, x, y + 1.326, w, 0.606, copy, { color: o.color || C.dark, align: o.align });
}
/** Speedometer built from two arcs, tick marks and a needle. */
function gauge(s, gx, gy, valueColor, trackColor) {
  const cx = gx + 1.3605, cy = gy + 1.3575;
  arc(s, gx + 0.428, gy + 0.425, 1.865, 1.865, [180.2, 359.3], trackColor, 20);
  arc(s, gx + 0.428, gy + 0.425, 1.865, 1.865, [178.9, 318.4], valueColor, 20);
  for (let i = 0; i <= 18; i++) {
    const a = (180 + i * 10) * Math.PI / 180;
    seg(s, cx + 1.085 * Math.cos(a) - 0.0515, cy + 1.085 * Math.sin(a), 0.103, 0,
      C.white, 1, { rotate: i * 10 });
  }
  shp(s, 'triangle', { x: gx + 1.686, y: gy + 0.610, w: 0.079, h: 0.819,
    fill: { color: C.white }, rotate: 49.8 });
  oval(s, gx + 1.239, gy + 1.235, 0.246, 0.246, C.white);
  const marks = [['0', 0, 1.228, 0.259, 'right'], ['25', 0.220, 0.448, 0.315, 'right'],
                 ['50', 1.236, 0, 0.315, 'right'], ['75', 2.250, 0.448, 0.315, 'left'],
                 ['100', 2.442, 1.228, 0.576, 'left']];
  marks.forEach(function (m) {
    text(s, m[0], { x: gx + m[1], y: gy + m[2], w: m[3], h: 0.218,
      fontSize: 7, bold: true, color: C.white, align: m[4], margin: 0 });
  });
}

/* ------------------------------------------------------------------ *
 * Chart helpers
 * ------------------------------------------------------------------ */
function chart(s, type, data, o) {
  s.addChart(type, data, Object.assign({
    showLegend: false, chartArea: { roundedCorners: false },
    catAxisLabelFontFace: F.body, valAxisLabelFontFace: F.body,
    catAxisLabelFontSize: 10, valAxisLabelFontSize: 10,
    catAxisLineColor: C.dark, valAxisLineColor: C.dark,
    catAxisMajorTickMark: 'none', valAxisMajorTickMark: 'none',
    catGridLine: { style: 'none' }, valGridLine: { style: 'none' }
  }, o));
}
/** Grid-line spec matching the deck's 30 %-opacity rules. */
function grid(color, size) {
  return { color: mix(color, C.white, 0.30), size: size || 1, style: 'solid' };
}
/** Horizontal 100 %-stacked "progress" bar (filled part + faint remainder). */
function progressBar(s, x, y, w, h, o) {
  o = o || {};
  chart(s, 'bar', [
    { name: 'Actual',  labels: [o.cat || '2025'], values: [o.pct === undefined ? 1900 : o.pct] },
    { name: 'Planned', labels: [o.cat || '2025'], values: [o.rest === undefined ? 2000 : o.rest] }
  ], {
    x, y, w, h, barDir: 'bar', barGrouping: 'percentStacked', barGapWidthPct: o.gap || 75,
    chartColors: [o.fg || C.white, o.bg || BAR_REST_ON_BLUE],
    catAxisHidden: true, valAxisLabelFormatCode: '0%', valAxisLabelPos: o.lblPos || 'low',
    valAxisLineShow: false, valAxisLabelColor: o.axColor || C.white,
    valGridLine: { color: mix(o.gridColor || C.white, o.gridOver || C.blue, 0.30), size: 0.5, style: 'solid' }
  });
}
/** Small "2 of 3" stacked bar with a numeric ruler underneath. */
function rulerBar(s, x, y, w, h, o) {
  o = o || {};
  chart(s, 'bar', [
    { name: 'Actual',  labels: ['Last month'], values: [2] },
    { name: 'Planned', labels: ['Last month'], values: [1] }
  ], {
    x, y, w, h, barDir: 'bar', barGrouping: 'stacked', barGapWidthPct: 125,
    chartColors: [o.fg || C.blue, o.bg || BAR_REST_ON_WHITE],
    catAxisHidden: true, valAxisLabelPos: 'low', valAxisMajorTickMark: 'out',
    valAxisMinorTickMark: 'in', valAxisLineColor: o.axColor || C.dark,
    valAxisLabelColor: o.axColor || C.dark, valGridLine: { style: 'none' }
  });
}
/** Vertical 100 %-stacked column used as a "thermometer". */
function columnMeter(s, x, y, w, h, o) {
  o = o || {};
  chart(s, 'bar', [
    { name: 'Actual',  labels: ['A'], values: [o.pct === undefined ? 2 : o.pct] },
    { name: 'Planned', labels: ['A'], values: [o.rest === undefined ? 1 : o.rest] }
  ], {
    x, y, w, h, barDir: 'col', barGrouping: 'percentStacked', barGapWidthPct: o.gap || 125,
    chartColors: [o.fg || C.blue, o.bg || BAR_REST_ON_WHITE],
    catAxisHidden: o.showCat !== true, catAxisLineColor: o.axColor || C.dark,
    catAxisLabelColor: o.axColor || C.dark,
    valAxisLabelFormatCode: '0%', valAxisLabelPos: 'low',
    valAxisMajorTickMark: o.ticks || 'out', valAxisMinorTickMark: o.ticks ? 'none' : 'in',
    valAxisLineShow: o.axLine !== false, valAxisLineColor: o.axColor || C.dark,
    valAxisLabelColor: o.axColor || C.dark, valAxisLogScaleBase: o.log ? 10 : undefined,
    valGridLine: o.grid ? { color: o.grid, size: 0.5, style: 'solid' } : { style: 'none' }
  });
}
/** Normalise each category to 0-1 so a stacked plot reads as 100 %-stacked. */
function toPercent(series) {
  const totals = series[0].map(function (_, i) {
    return series.reduce(function (t, row) { return t + row[i]; }, 0);
  });
  return series.map(function (row) { return row.map(function (v, i) { return v / totals[i]; }); });
}

/** Two-value doughnut: coloured slice over a faint remainder ring. */
function donut(s, x, y, w, h, val, rest, color, restColor, hole) {
  chart(s, 'doughnut', [{ name: 'Sales', labels: ['Market Saze A', 'Market Saze B'], values: [val, rest] }], {
    x, y, w, h, holeSize: hole || 80, firstSliceAng: 0,
    chartColors: [color, restColor || BAR_REST_ON_WHITE], dataNoEffects: true
  });
}

/* ------------------------------------------------------------------ *
 * Shared copy
 * ------------------------------------------------------------------ */
const L = {
  full:  "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.",
  mid:   "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. ",
  ideaLong: "Lorem Ipsum\u00a0is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. ",
  idea:  "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. ",
  book:  "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type book. ",
  short: "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s.",
  scrap: "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled book.",
  scrap2:"Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled  book. ",
  spec:  "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. ",
  stdBook: "Lorem Ipsum has been the standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a book",
  leap:  "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting. ",
  printer: "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer.",
  tiny:  'Lorem Ipsum is simply dummy text.',
  text5: 'Lorem Ipsum has been the text.',
  everSince: 'Lorem Ipsum has been the standard dummy text ever.'
};

const WEEK = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
const MON1 = ['J', 'F', 'M', 'A', 'M', 'J', 'J', 'A', 'S', 'O', 'N', 'D'];
const QTR  = ['Iq', 'IIq', 'IIIq ', 'IVq'];
const SER3 = [[2.5, 3.5, 4.5, 4.3], [4.4, 1.8, 2.8, 2.4], [2, 3, 5, 2]];
const WAVE = [4.3, 2.5, 3.5, 5, 3, 2, 4];

/* ------------------------------------------------------------------ *
 * Slide 1 -- title
 * ------------------------------------------------------------------ */
function slide01(s) {
  s.background = { color: C.page };
  text(s, 'Krishna said', { x: 2.182, y: 1.84, w: 5.279, h: 1.447,
    fontFace: F.head, fontSize: 80, bold: true, color: C.navy });
  text(s, 'POWERPOINT', { x: 2.261, y: 3.288, w: 5.231, h: 0.303,
    fontSize: 12, charSpacing: 3, color: C.blue });
}

/* ------------------------------------------------------------------ *
 * Slide 2 -- line chart + "Main Idea" band
 * ------------------------------------------------------------------ */
function slide02(s) {
  s.addText(rich([['"Man is '], ['made by his belief', C.blue],
    ['. As he believes, so he is…" ']], { fontFace: F.head, fontSize: 44, bold: true }),
    { x: 0.76, y: 0.771, w: 5.394, h: 2.322, fontFace: F.head, fontSize: 44, bold: true, color: C.dark });
  body(s, 0.76, 3.292, 5.394, 1.279, L.full);

  dataText(s, 7.217, 0.81, 2.728, 'Data text', '11 330');
  chart(s, 'line', [{ name: 'stats', labels: WEEK, values: WAVE }], {
    x: 7.179, y: 1.588, w: 5.734, h: 3.073, chartColors: [C.teal],
    lineDataSymbol: 'circle', lineDataSymbolSize: 8, lineDataSymbolLineColor: C.teal,
    lineSize: 2, valGridLine: grid(C.dark)
  });
  pill(s, 9.737, 1.689, 0.821, 0.335, '+15%', { fill: C.teal });

  band(s, 4.762, 2.738);
  progressBar(s, 7.756, 5.006, 4.867, 1.313, {});
  pill(s, 6.396, 5.936, 1.986, 0.394, '+15% increase', { fill: C.white, color: C.blue, rotate: 270 });
  bulbIcon(s, 0.71, 5.243, 1.131, 1.731, mix(C.white, C.blue, 0.50), C.blue);
  ideaBlock(s, 2.35, 5.335, 3.803, 'Main Idea', L.idea);
  badge(s, 10.19, 5.176, 0.731, '100%');
  note(s, 7.756, 6.486, 4.867, 0.438, L.scrap2, { color: C.white, paraSpaceAfter: 6 });
}

/** Light bulb whose glass holds the two-lobed brain (slide 2). */
function bulbIcon(s, x, y, w, h, glassColor, lineColor) {
  const bw = w * 0.98, bx = x + (w - bw) / 2;
  oval(s, bx, y, bw, bw, glassColor);                                 // glass
  rect(s, bx + bw * 0.30, y + bw * 0.84, bw * 0.40, h * 0.16, glassColor); // neck
  brainIcon(s, bx + bw * 0.11, y + bw * 0.16, bw * 0.78, bw * 0.66, C.white, lineColor);
  for (let i = 0; i < 4; i++) {                                       // screw base
    shp(s, 'roundRect', { x: bx + bw * 0.30, y: y + h - (4 - i) * h * 0.075, w: bw * 0.40, h: h * 0.05,
      fill: { color: i % 2 ? C.blue3 : mix(C.blue3, C.dark, 0.75) }, line: NONE, rectRadius: 0.02 });
  }
}

/* ------------------------------------------------------------------ *
 * Slide 3 -- blue header band, area chart, guillemets
 * ------------------------------------------------------------------ */
function slide03(s) {
  band(s, 0, 2.738);
  brainIcon(s, 0.734, 0.5, 1.811, 1.766, C.white, C.blue, true);
  statBlock(s, 3.051, 0.537, 3.104, 'Lorem Ipsum is text', '100%', L.short, { color: C.white });
  chart(s, 'area', [{ name: 'stats', labels: WEEK, values: WAVE }], {
    x: 7.179, y: 0.408, w: 4.889, h: 1.892, chartColors: [C.white],
    catAxisLabelColor: C.white, valAxisLabelColor: C.white,
    catAxisLineColor: C.white, valAxisLineColor: C.white,
    valGridLine: grid(C.white)
  });
  pill(s, 11.697, 1.165, 1.44, 0.394, '$345k', { fill: C.white, color: C.blue, rotate: 270 });

  quote(s, 1.084, 3.05, 5.07, 3.803, [
    ['Change is the law ', C.blue], ['of the Universe. \nYou can be a millionaire, or a pauper in an instant. ']
  ], 44);
  iconChevrons(s, 0.415, 3.053, 0.691, 0.737, C.blue);
  iconChevrons(s, 6.321, 6.116, 0.691, 0.737, C.blue, true);
  dataText(s, 10.194, 3.213, 2.419, 'Data text', '11 330', { align: 'right' });
  progressBar(s, 7.179, 4.057, 5.434, 1.379, { cat: 'Last month', pct: 2, rest: 1, gap: 125,
    fg: C.blue, bg: BAR_REST_ON_WHITE, axColor: C.dark, gridColor: C.dark, gridOver: C.white });
  seg(s, 7.642, 3.213, 0, 1.13, C.blue, 1);
  pill(s, 7.493, 3.215, 0.821, 0.335, '+15%');
  note(s, 7.701, 3.614, 2.164, 0.606, L.short);
  body(s, 7.493, 5.541, 5.12, 1.279, L.full);
}

/**
 * Two-lobed "icon brain": each lobe is a cluster of overlapping discs with
 * small glyphs scattered inside, matching the artwork on slides 2 and 3.
 */
const BRAIN_LOBE = [   // unit-square disc centres + radii for one lobe
  [0.46, 0.12, 0.40], [0.16, 0.26, 0.36], [0.76, 0.24, 0.36], [0.10, 0.50, 0.38],
  [0.44, 0.40, 0.46], [0.82, 0.52, 0.36], [0.16, 0.74, 0.38], [0.48, 0.70, 0.46],
  [0.82, 0.76, 0.34], [0.32, 0.92, 0.34], [0.66, 0.92, 0.34], [0.50, 0.50, 0.60]
];
function brainIcon(s, x, y, w, h, fillColor, lineColor, withGlyphs) {
  const half = w * 0.49, d = Math.min(half, h) ;
  for (let i = 0; i < 2; i++) {
    const ox = x + i * (w - half);
    BRAIN_LOBE.forEach(function (b) {
      const r = b[2] * d;
      const bx = i === 0 ? b[0] : 1 - b[0];
      oval(s, ox + bx * half - r / 2, y + b[1] * h - r / 2, r, r, fillColor);
    });
    if (!withGlyphs) continue;
    ['gear', 'camera', 'note', 'people', 'calendar', 'folder'].forEach(function (k, j) {
      const g = [[0.30, 0.22], [0.64, 0.30], [0.26, 0.50], [0.62, 0.60], [0.34, 0.78], [0.66, 0.84]][j];
      const gx = i === 0 ? g[0] : 1 - g[0];
      const gs = d * 0.20;
      glyph(s, k, ox + gx * half - gs / 2, y + g[1] * h - gs / 2, gs, gs, lineColor);
    });
  }
}

/* ------------------------------------------------------------------ *
 * Slide 4 -- doughnut + "Main Idea" band
 * ------------------------------------------------------------------ */
function slide04(s) {
  quote(s, 0.76, 0.771, 8.926, 2.322, [
    ['"You have the '], ['right', C.blue],
    [' to perform your duties, but you are not entitled to the fruits of your actions." ']
  ], 44);
  donut(s, 9.892, 0.889, 2.141, 2.029, 9, 11, C.blue);
  text(s, '$22k', { x: 10.388, y: 1.56, w: 1.097, h: 0.404, fontFace: F.head, fontSize: 18,
    align: 'center', valign: 'bottom' });
  pill(s, 10.234, 1.958, 1.399, 0.315, 'Lorem ipsum', { size: 10, font: F.body, spc: 0 });
  pill(s, 11.585, 1.62, 1.699, 0.394, 'New York', { rotate: 270 });

  band(s, 3.393, 2.738);
  progressBar(s, 7.756, 3.637, 4.867, 1.313, {});
  pill(s, 6.396, 4.567, 1.986, 0.394, 'Growth', { fill: C.white, color: C.blue, rotate: 270 });
  bulbSolid(s, 0.858, 3.747, 1.396, 2.029, C.purple);
  ideaBlock(s, 2.581, 3.966, 3.572, 'Main Idea', L.idea);
  badge(s, 10.19, 3.807, 0.731, '100%');
  note(s, 7.756, 5.117, 4.867, 0.438, L.scrap2, { color: C.white, paraSpaceAfter: 6 });
  note(s, 0.76, 6.338, 5.394, 0.606, L.leap);
  note(s, 7.223, 6.338, 5.394, 0.606,
    'It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.');
}

/** Silhouette light bulb (solid glass + translucent screw base). */
function bulbSolid(s, x, y, w, h, color) {
  oval(s, x + (w - h * 0.657) / 2, y, h * 0.657, h * 0.657, color);
  const ghost = mix(C.white, C.blue, 0.20);
  rect(s, x + w * 0.30, y + h * 0.62, w * 0.40, h * 0.16, ghost);
  for (let i = 0; i < 3; i++) {
    shp(s, 'roundRect', { x: x + w * 0.30, y: y + h * (0.79 + i * 0.075), w: w * 0.40, h: h * 0.055,
      fill: { color: ghost }, rectRadius: 0.02 });
  }
  shp(s, 'trapezoid', { x: x + w * 0.34, y: y + h * 0.95, w: w * 0.32, h: h * 0.05, fill: { color: ghost } });
}

/* ------------------------------------------------------------------ *
 * Slide 5 -- pie + KPI cards
 * ------------------------------------------------------------------ */
function slide05(s) {
  band(s, 0.463, 3.008);
  quote(s, 0.76, 0.771, 7.852, 2.322, [
    ['"You are what you ', C.white], ['believe in', C.purple],
    ['. You become that which you believe you can become." ', C.white]
  ], 44, C.white);
  progressBar(s, 8.836, 0.797, 3.781, 1.203, { lblPos: 'high' });
  body(s, 8.836, 2.151, 3.781, 0.942, L.idea, { color: C.white, paraSpaceAfter: 6 });

  chart(s, 'pie', [{ name: 'Sales', labels: ['AS', 'FR', 'GT', 'HE'], values: [8, 11, 7, 4] }], {
    x: 0.72, y: 4.442, w: 2.619, h: 2.595, firstSliceAng: 0,
    chartColors: [C.blue, C.purple, C.teal, C.pink], dataNoEffects: true
  });
  text(s, 'Increase your Power', { x: 3.552, y: 4.535, w: 2.422, h: 0.37,
    fontFace: F.head, fontSize: 16, valign: 'bottom' });
  text(s, '2025', { x: 3.552, y: 4.905, w: 2.422, h: 0.303, fontSize: 12, color: C.blue });
  body(s, 3.552, 5.313, 2.422, 1.447, L.idea);

  iconRow(s, function (x, col) { iconDiamond(s, x, 3.887, 0.308, col); },
    0.762, 0, 0.3545, 10, 5, C.blue, C.dark);
  oval(s, 4.48, 3.829, 0.428, 0.428, C.blue);
  shp(s, 'star5', { x: 4.585, y: 3.936, w: 0.215, h: 0.215, fill: { color: C.white } });
  text(s, '50%', { x: 4.911, y: 3.828, w: 0.977, h: 0.404,
    fontFace: F.head, fontSize: 18, valign: 'bottom' });

  const cards = [
    [7.184, 3.804, '> 2 500', '+5%', C.blue,   7.224, 7.406],
    [9.157, 3.804, '> 3 970', '+7%', C.purple, 9.197, 9.379],
    [7.184, 5.543, '> 2 500', '+5%', C.teal,   7.165, 7.406],
    [9.157, 5.543, '> 3 970', '+7%', C.pink,   9.197, 9.379]
  ];
  cards.forEach(function (c) {
    pill(s, c[0], c[1], 1.299, 0.354, c[2], { fill: c[4] });
    text(s, c[3], { x: c[6], y: c[1] + 0.542, w: 1.327, h: 0.303, fontFace: F.head, fontSize: 12 });
    iconArrowUp(s, c[5], c[1] + 0.560, 0.141, 0.226, c[4]);
    body(s, c[6], c[1] + 0.843, 1.327, 0.438, L.tiny);
  });

  chart(s, 'bar', [
    { name: 'Stage I',   labels: ['Progress'], values: [389] },
    { name: 'Stage II',  labels: ['Progress'], values: [204] },
    { name: 'Stage III', labels: ['Progress'], values: [297] },
    { name: 'Stage IV',  labels: ['Progress'], values: [122] }
  ], {
    x: 10.919, y: 3.75, w: 1.698, h: 3.071, barDir: 'col', barGrouping: 'percentStacked',
    barGapWidthPct: 75, chartColors: [C.blue, C.purple, C.teal, C.pink],
    catAxisHidden: true, valAxisLabelFormatCode: '0%', valAxisLabelPos: 'low',
    valAxisLineColor: mix(C.dark, C.white, 0.30), valAxisLogScaleBase: 10,
    showValue: true, dataLabelColor: C.white, dataLabelFontSize: 12, dataLabelFontBold: true,
    dataLabelFormatCode: '#,##0'
  });
}

/* ------------------------------------------------------------------ *
 * Slide 6 -- meditation step circles
 * ------------------------------------------------------------------ */
function slide06(s) {
  quote(s, 1.462, 0.797, 8.195, 2.322, [
    ['"When '], ['meditation', C.blue],
    [' is mastered, the mind is unwavering like the flame of a lamp in a windless place." ']
  ], 44);
  note(s, 1.462, 3.385, 8.195, 0.774, L.full);
  pill(s, -0.204, 1.761, 2.322, 0.394, 'Krishna said', { rotate: 270 });
  columnMeter(s, 10.25, 0.797, 2.511, 3.504, { pct: 2, rest: 1, fg: C.blue, bg: BAR_REST_ON_WHITE });

  const steps = [
    [0.798, C.purple, 'Meditation \nLevel I'],
    [3.347, C.teal,   'Level II'],
    [5.897, C.pink,   'Anahata\nLevel'],
    [8.447, C.blue2,  'Om \nMeditation'],
    [10.996, C.blue3, 'Level III']
  ];
  steps.forEach(function (st) { oval(s, st[0], 4.676, 1.608, 1.608, st[1]); });
  steps.forEach(function (st, i) {
    s.addText(st[2], { x: st[0] - 0.2, y: 4.676, w: 2.008, h: 1.608, color: C.white,
      fontFace: F.head, fontSize: 18, align: 'center', valign: 'middle', margin: 0, wrap: false });
    if (i < 4) iconArrowCircle(s, 2.602 + i * 2.5495, 5.206, 0.547, FAINT_DARK_ON_WHITE);
    const cx = [0.789, 3.446, 6.22, 8.95, 10.988][i];
    text(s, rich([['Lorem Ipsum ', null, { bold: true }], ["has been the industry's standard dummy text. "]],
      { fontSize: 10 }), { x: cx, y: 6.406, w: 1.625, h: 0.606, align: 'center', paraSpaceAfter: 6 });
  });
}

/* ------------------------------------------------------------------ *
 * Slide 7 -- big pie wedge
 * ------------------------------------------------------------------ */
function slide07(s) {
  s.background = { color: C.page };
  chart(s, 'pie', [{ name: 'Sales', labels: ['Market Saze A', 'Market Saze B'], values: [8, 11] }], {
    x: 1.258, y: 3.593, w: 3.985, h: 3.606, firstSliceAng: 0,
    chartColors: [C.purple, C.page], dataNoEffects: true
  });
  wedge(s, 1.838, 4.001, 2.763, 2.763, [44.9, 278.6], C.blue);
  quote(s, 7.179, 1.271, 4.164, 4.342, [
    ['"When consciousness is unified, however, all vain anxiety is left behind. There is ', C.navy],
    ['no cause for worry', C.blue], [', whether things go well or ill." ', C.navy]
  ], 36, C.navy);
  pill(s, 10.914, 2.235, 2.322, 0.394, 'Krishna said', { rotate: 270 });
  text(s, '$129.45', { x: 2.467, y: 4.103, w: 0.539, h: 2.559, fontFace: F.head, fontSize: 20,
    charSpacing: 3, color: C.white, align: 'center', valign: 'bottom', vert: 'vert270' });
  text(s, '46%', { x: 3.413, y: 4.517, w: 1.247, h: 0.64, fontFace: F.head, fontSize: 32, color: C.white });
  note(s, 3.413, 5.145, 1.319, 0.438, 'Lorem Ipsum has been the text.', { color: C.white });
}

/* ------------------------------------------------------------------ *
 * Slide 8 -- maze head panel
 * ------------------------------------------------------------------ */
function slide08(s) {
  panel(s, 6.667, 6.293);
  headProfile(s, 7.054, 0.523, 3.91, 4.002, { fill: FAINT_WHITE_ON_BLUE });
  quote(s, 0.76, 1.484, 5.249, 3.13, [
    ['"It is better to '], ['live your own destiny', C.blue],
    [" imperfectly than to live an imitation of somebody else's life with perfection.\" "]
  ], 36);
  note(s, 0.76, 4.905, 5.249, 0.774, L.mid);
  pill(s, 0.76, 0.797, 2.322, 0.394, 'Krishna said');
  rulerBar(s, 0.76, 5.838, 5.249, 0.983, {});

  // Concentric maze rings
  const maze = [[7.383, 0.739, 2.731, 315.4, 307.0], [7.549, 0.905, 2.402, 246.5, 232.9],
                [7.741, 1.103, 2.003, 153.3, 135.3], [7.915, 1.278, 1.654, 347.5, 327.8],
                [8.090, 1.454, 1.302, 153.1, 127.7], [8.281, 1.645, 0.921, 291.3, 263.9],
                [8.471, 1.830, 0.551, 223.6, 179.4]];
  maze.forEach(function (m) { arc(s, m[0], m[1], m[2], m[2], [m[3], m[4]], C.white, 2.5); });
  oval(s, 8.595, 1.945, 0.306, 0.306, C.purple);

  pill(s, 7.179, 4.255, 1.009, 0.335, '+5%', { fill: C.white, color: C.dark });
  text(s, '24/7', { x: 7.179, y: 4.904, w: 2.647, h: 0.505, fontFace: F.head, fontSize: 24,
    color: C.white, valign: 'bottom' });
  body(s, 7.179, 5.451, 2.647, 0.774, L.scrap, { color: C.white });
  chart(s, 'pie', [{ name: 'Sales', labels: ['Market Saze A', 'Market Saze B'], values: [8, 11] }], {
    x: 10.004, y: 4.417, w: 2.588, h: 2.379, firstSliceAng: 0,
    chartColors: [C.white, BAR_REST_ON_BLUE], dataNoEffects: true
  });
  pill(s, 11.678, 4.868, 0.899, 0.259, '> 10', { fill: C.purple, size: 10 });
  text(s, '+15%', { x: 11.859, y: 5.232, w: 0.415, h: 0.269, fontFace: F.head, fontSize: 10,
    color: C.darkBlue, align: 'center' });
  iconArrowUp(s, 11.718, 5.251, 0.141, 0.226, C.purple);
  iconRow(s, function (x, col) { iconStar(s, x, 6.551, 0.244, col); }, 7.287, 0, 0.2846, 5, 0, C.white, C.white);
  columnMeter(s, 11.09, 0.64, 1.532, 3.615, { pct: 100, rest: 2000, gap: 75,
    fg: C.purple, bg: BAR_REST_ON_BLUE, showCat: true, axColor: C.white, axLine: false,
    ticks: 'none', grid: mix(C.white, C.blue, 0.3) });
}

/**
 * Right-facing head profile: smooth cranium (cubic curves) plus the
 * nose / lips / chin notches. Stand-in for the deck's photo artwork.
 */
function headProfile(s, x, y, w, h, opt) {
  const P = [
    { x: w * 0.06, y: h },
    { x: w * 0.06, y: h * 0.56, curve: { type: 'cubic', x1: w * 0.02, y1: h * 0.82, x2: w * 0.00, y2: h * 0.68 } },
    { x: w * 0.50, y: h * 0.02, curve: { type: 'cubic', x1: w * 0.12, y1: h * 0.22, x2: w * 0.28, y2: h * 0.02 } },
    { x: w * 0.90, y: h * 0.44, curve: { type: 'cubic', x1: w * 0.74, y1: h * 0.02, x2: w * 0.90, y2: h * 0.20 } },
    { x: w * 1.00, y: h * 0.60 },   // nose bridge
    { x: w * 0.88, y: h * 0.65 },   // nostril
    { x: w * 0.90, y: h * 0.72 },   // lips
    { x: w * 0.84, y: h * 0.74 },
    { x: w * 0.86, y: h * 0.84 },   // chin
    { x: w * 0.66, y: h * 0.90, curve: { type: 'cubic', x1: w * 0.84, y1: h * 0.90, x2: w * 0.74, y2: h * 0.90 } },
    { x: w * 0.66, y: h },
    { close: true }
  ];
  shp(s, 'custGeom', { x, y, w, h, points: P,
    fill: opt.fill ? { color: opt.fill } : NONE,
    line: opt.stroke ? { color: opt.stroke, width: 2.5, cap: 'round' } : NONE });
}

/* ------------------------------------------------------------------ *
 * Slide 9 -- bubble diagram panel
 * ------------------------------------------------------------------ */
function slide09(s) {
  rect(s, 0.374, 0, 6.293, H, C.blue);
  quote(s, 7.21, 1.484, 5.394, 3.13, [
    ['"To '], ['refrain from selfish acts ', C.blue],
    ['is one kind of renunciation, called sannyasa; to renounce the fruit of action is another, called '],
    ['Tyaga', C.blue], ['." ']
  ], 36);
  body(s, 7.21, 6.046, 5.394, 0.774, L.mid);
  pill(s, 7.21, 0.797, 2.322, 0.394, 'Krishna said');

  pill(s, 0.76, 0.674, 1.009, 0.335, '+5%', { fill: C.white, color: C.dark });
  text(s, '24/7', { x: 0.76, y: 1.323, w: 2.647, h: 0.505, fontFace: F.head, fontSize: 24,
    color: C.white, valign: 'bottom' });
  body(s, 0.76, 1.869, 2.647, 0.774, L.scrap, { color: C.white });
  iconRow(s, function (x, col) { iconStar(s, x, 2.97, 0.244, col); }, 0.869, 0, 0.2846, 5, 0, C.white, C.white);

  // Free-form connector ribbon threading the bubbles
  [[1.464, 3.951, 2.278, 182.2, 327.2], [3.510, 3.205, 1.761, 271.9, 149.3],
   [3.429, 1.210, 1.995, 96.3, 269.1], [3.510, -0.786, 1.995, 31.6, 91.6],
   [-0.814, 3.849, 2.278, 7.1, 61.2]].forEach(function (a) {
    arc(s, a[0], a[1], a[2], a[2], [a[3], a[4]], C.dark, 2.5);
  });
  badge(s, 1.717, 4.204, 1.772, '58%', { fill: C.white, color: C.dark, font: F.head, size: 36, spc: 0 });
  badge(s, 3.661, 3.353, 1.442, '16%', { fill: C.purple, color: C.white, font: F.head, size: 36, spc: 0 });
  badge(s, 3.619, 1.445, 1.525, '31%', { fill: C.teal, color: C.white, font: F.head, size: 36, spc: 0 });

  pill(s, 5.079, 5.269, 1.009, 0.335, '+9%', { fill: C.white, color: C.dark });
  text(s, '24/7', { x: 3.675, y: 5.198, w: 1.207, h: 0.505, fontFace: F.head, fontSize: 24,
    color: C.white, valign: 'bottom' });
  body(s, 3.675, 5.744, 2.647, 0.774, L.scrap, { color: C.white });
  iconRow(s, function (x, col) { iconStar(s, x, 6.845, 0.244, col); }, 3.783, 0, 0.2846, 5, 0, C.white, C.white);
  progressBar(s, 7.179, 4.734, 5.434, 1.242, { cat: 'Last month', pct: 2, rest: 1, gap: 125,
    fg: C.blue, bg: BAR_REST_ON_WHITE, axColor: C.dark, gridColor: C.dark, gridOver: C.white });
}

/* ------------------------------------------------------------------ *
 * Slide 10 -- happiness band with three charts
 * ------------------------------------------------------------------ */
function slide10(s) {
  band(s, 2.165, 2.51);
  quote(s, 0.76, 2.474, 5.394, 1.919, [
    ['"', C.white], ['Happiness', C.purple], [' is a state of ', C.white], ['mind', C.purple],
    [' and has nothing to do with the external world." ', C.white]
  ], 36, C.white);
  iconRow(s, function (x, col) { iconDiamond(s, x, 0.856, 0.308, col); },
    0.762, 0, 0.3545, 10, 6, C.blue, C.dark);
  oval(s, 4.48, 0.798, 0.428, 0.428, C.blue);
  shp(s, 'star5', { x: 4.585, y: 0.905, w: 0.215, h: 0.215, fill: { color: C.white } });
  text(s, '60%', { x: 4.911, y: 0.797, w: 0.977, h: 0.404, fontFace: F.head, fontSize: 18, valign: 'bottom' });
  body(s, 0.76, 1.362, 5.394, 0.606,
    "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap. ");

  chart(s, 'line', [{ name: 'Series 1', labels: MON1, values: [4.3, 2.5, 3.5, 4.5, 3, 2, 4, 5, 6, 4, 2, 3] }], {
    x: 7.179, y: 0.642, w: 5.443, h: 1.273, chartColors: [C.purple], lineSize: 2,
    lineDataSymbol: 'circle', lineDataSymbolSize: 8, lineDataSymbolLineColor: C.purple,
    valAxisHidden: true, catAxisLineShow: false
  });
  chart(s, 'bar', [
    { name: 'Sales',   labels: ['Market Saze A'], values: [5] },
    { name: 'Column1', labels: ['Market Saze A'], values: [7] },
    { name: 'Column2', labels: ['Market Saze A'], values: [3] }
  ], {
    x: 7.179, y: 2.862, w: 5.443, h: 1.13, barDir: 'bar', barGrouping: 'percentStacked',
    barGapWidthPct: 100, chartColors: [C.white, C.pink, C.teal],
    catAxisHidden: true, valAxisLabelFormatCode: '0%', valAxisLabelColor: C.white,
    valAxisLineShow: false, valAxisMajorTickMark: 'out',
    valGridLine: grid(C.white, 0.75)
  });
  pill(s, 7.486, 2.667, 0.767, 0.259, 'I', { fill: C.white, color: C.blue, size: 10 });
  pill(s, 10.527, 2.667, 0.767, 0.259, 'II', { fill: C.pink, size: 10 });
  pill(s, 11.496, 2.667, 0.767, 0.259, 'III', { fill: C.teal, size: 10 });
  note(s, 7.78, 4.123, 4.842, 0.269, "Lorem Ipsum has been the industry's standard dummy text.",
    { color: C.white, align: 'center', paraSpaceAfter: 6 });

  const keys = [[0.746, C.blue, 'target', 0.76], [2.607, C.purple, 'gear', 2.567], [4.373, C.teal, 'briefcase', 4.347]];
  keys.forEach(function (k) {
    iconDisc(s, k[2], k[0], 5.002, 0.553, k[1]);
    text(s, 'Key of happy', { x: k[3], y: 5.624, w: 1.645, h: 0.303, fontSize: 12, valign: 'bottom' });
  });
  chart(s, 'bar', [
    { name: 'A', labels: WEEK, values: [4.3, 2.5, 3.5, 4.5, 3, 2, 4] },
    { name: 'B', labels: WEEK, values: [3, 1, 3, 4.2, 2, 1.9, 3] },
    { name: 'C', labels: WEEK, values: [4.5, 4, 5, 2, 4, 3, 4] }
  ], {
    x: 7.179, y: 4.766, w: 5.443, h: 2.047, barDir: 'col', barGrouping: 'stacked',
    barGapWidthPct: 90, chartColors: [C.blue, C.purple, C.teal],
    valAxisHidden: true
  });
  note(s, 0.76, 6.199, 5.394, 0.606, L.leap);
}

/* ------------------------------------------------------------------ *
 * Slide 11 -- idea-cloud head
 * ------------------------------------------------------------------ */
/** Scatter of small grey glyphs that fills the outlined head on slide 11. */
const CLOUD_ICONS = [
  [2.46, 1.14, 0.30], [2.72, 0.85, 0.56], [3.21, 2.02, 0.48], [1.62, 2.52, 0.29], [1.32, 1.06, 0.28],
  [1.69, 3.40, 0.54], [2.89, 3.26, 0.27], [1.47, 1.47, 0.26], [2.96, 2.42, 0.22], [3.32, 2.71, 0.37],
  [1.08, 2.07, 0.60], [3.37, 1.41, 0.34], [2.96, 1.82, 0.30], [3.75, 2.04, 0.46], [2.34, 1.54, 0.68],
  [4.25, 2.10, 0.37], [2.07, 1.91, 0.25], [2.26, 1.36, 0.18], [1.30, 3.42, 0.30], [2.08, 3.29, 0.35],
  [3.05, 1.54, 0.28], [3.61, 3.48, 0.27], [1.31, 1.81, 0.23], [3.52, 1.82, 0.18], [1.96, 2.21, 0.32],
  [3.85, 1.41, 0.45], [3.27, 0.84, 0.21], [3.31, 1.04, 0.24], [2.80, 3.77, 0.45], [3.19, 3.53, 0.22],
  [3.00, 2.72, 0.26], [1.63, 3.06, 0.38], [1.74, 2.24, 0.19], [4.09, 2.49, 0.43], [3.75, 1.11, 0.16],
  [4.16, 1.82, 0.33], [0.90, 1.94, 0.35], [1.72, 1.91, 0.30], [3.53, 1.19, 0.20], [3.04, 0.91, 0.20],
  [2.55, 2.23, 0.28], [2.62, 3.62, 0.21], [3.36, 3.12, 0.32], [2.00, 2.71, 0.54], [2.61, 2.88, 0.32],
  [2.63, 2.54, 0.24], [2.41, 0.73, 0.28], [2.73, 1.40, 0.32], [3.79, 2.57, 0.29], [2.31, 2.25, 0.16],
  [1.02, 1.36, 0.36], [2.59, 3.22, 0.26], [3.60, 2.43, 0.23], [3.83, 3.08, 0.28], [0.96, 2.56, 0.27],
  [1.15, 2.86, 0.50], [2.11, 0.91, 0.30], [1.67, 0.85, 0.33], [3.74, 1.87, 0.17], [1.80, 1.21, 0.45]
];
const CLOUD_KINDS = ['gear', 'target', 'note', 'camera', 'calendar', 'people', 'folder', 'briefcase', 'dollar'];

function slide11(s) {
  quote(s, 6.143, 0.771, 6.481, 2.524, [
    ['"Make it easy for yourself. Get organized and '], ['live in the moment', C.blue],
    ['. Your day-to-day activities should not steal your happiness." ']
  ], 36);
  note(s, 6.143, 3.485, 6.469, 0.942, L.full);
  text(s, '11 330', { x: 4.169, y: 2.88, w: 2.482, h: 0.404, fontFace: F.head, fontSize: 18,
    charSpacing: 3, color: C.gray, valign: 'bottom', rotate: 90 });
  pill(s, 4.473, 1.132, 0.978, 0.335, '+100%');

  headProfile(s, 0.785, 0.652, 4.224, 4.323, { fill: C.white, stroke: C.dark });
  const faint = mix(C.dark, C.white, 0.18);
  CLOUD_ICONS.forEach(function (ic, i) {
    glyph(s, CLOUD_KINDS[i % CLOUD_KINDS.length], ic[0], ic[1], ic[2], ic[2], faint);
  });
  // Purple bulb at the centre of the head
  oval(s, 1.967, 1.374, 1.08, 1.15, C.purple);
  rect(s, 2.275, 2.45, 0.463, 0.47, C.purple);
  shp(s, 'custGeom', { x: 2.269, y: 2.005, w: 0.502, h: 0.858, fill: NONE,
    line: { color: C.white, width: 2 },
    points: [{ x: 0.25, y: 0.86 }, { x: 0.25, y: 0.42 }, { x: 0.02, y: 0.06 },
             { x: 0.25, y: 0.30, moveTo: true }, { x: 0.48, y: 0.06 }] });
  [2.917, 3.020, 3.124].forEach(function (yy) {
    shp(s, 'roundRect', { x: 2.275, y: yy, w: 0.463, h: 0.068, fill: { color: C.purple }, rectRadius: 0.03 });
  });
  shp(s, 'trapezoid', { x: 2.349, y: 3.227, w: 0.316, h: 0.108, fill: { color: C.purple } });

  band(s, 4.762, 2.738);
  ideaBlock(s, 1.307, 5.335, 4.224, 'My Moment', L.idea);
  chart(s, 'area', [{ name: 'stats', labels: WEEK, values: WAVE }], {
    x: 6.155, y: 5.128, w: 6.469, h: 1.912, barGrouping: 'stacked', chartColors: [C.white],
    catAxisLabelColor: C.white, valAxisLabelColor: C.white,
    catAxisLineColor: C.white, valAxisLineColor: C.white, valGridLine: grid(C.white)
  });
}

/* ------------------------------------------------------------------ *
 * Slide 12 -- gauge header + KPI cards + stacked bar
 * ------------------------------------------------------------------ */
function slide12(s) {
  band(s, 0, 2.738);
  pill(s, 0.228, 1.165, 1.44, 0.394, '$345k', { fill: C.white, color: C.blue, rotate: 270 });
  body(s, 1.344, 0.655, 4.564, 1.447, L.full, { color: C.white, valign: 'middle' });
  gauge(s, 6.731, 0.557, C.white, BAR_REST_ON_BLUE);
  iconDiamond(s, 10.326, 0.378, 0.721, C.white);
  iconPerson(s, 10.361, 1.008, 1.174, 1.464, mix(C.white, C.blue, 0.60));

  quote(s, 0.76, 3.05, 5.15, 3.736, [
    ['"A person can rise through the efforts of his own mind; or draw himself down, in the same manner. Because '],
    ['each person is his own friend or enemy', C.blue], ['." ']
  ], 36);
  [[7.330, C.blue, '13%'], [8.970, C.purple, '27%'], [10.609, C.teal, '19%']].forEach(function (c) {
    shp(s, 'roundRect', { x: c[0], y: 3.114, w: 1.452, h: 0.676, fill: { color: c[1] }, rectRadius: 0.06 });
    text(s, c[2], { x: c[0] + 0.13, y: 3.211, w: 1.203, h: 0.337, fontFace: F.head, fontSize: 14,
      color: C.white, align: 'center', valign: 'bottom' });
    note(s, c[0] + 0.13, 3.473, 1.203, 0.269, 'Lorem ipsum', { color: C.white, align: 'center' });
  });
  chart(s, 'bar', [
    { name: 'Series 1', labels: QTR, values: SER3[0] },
    { name: 'Series 2', labels: QTR, values: SER3[1] },
    { name: 'Series 3', labels: QTR, values: SER3[2] }
  ], {
    x: 6.667, y: 4.004, w: 5.946, h: 2.827, barDir: 'bar', barGrouping: 'percentStacked',
    barGapWidthPct: 150, chartColors: [C.blue, C.purple, C.teal],
    valAxisLabelFormatCode: '0%',
    catAxisMajorTickMark: 'out', valAxisMajorTickMark: 'out', valGridLine: grid(C.dark, 0.5)
  });
}

/* ------------------------------------------------------------------ *
 * Slide 13 -- target icon + KPI + blue band quote
 * ------------------------------------------------------------------ */
function slide13(s) {
  band(s, 2.984, 2.738);
  quote(s, 0.751, 3.397, 7.037, 1.919,
    [['" There is neither this world nor the world beyond nor happiness for the one who doubts." ', C.white]],
    36, C.white);
  progressBar(s, 8.366, 3.326, 4.257, 1.313, {});
  badge(s, 10.495, 3.531, 0.731, '>10hr');
  note(s, 8.366, 4.805, 4.257, 0.606, L.scrap2, { color: C.white, paraSpaceAfter: 6 });

  // Concentric target with an arrow through the bullseye
  [[0.751, 0.778, 1.827], [1.074, 1.101, 1.180], [1.396, 1.423, 0.538]].forEach(function (t, i) {
    const deg = [[330.0, 298.4], [331.2, 295.7], [344.5, 280.5]][i];
    arc(s, t[0], t[1], t[2], t[2], deg, C.dark, 2.5);
  });
  shp(s, 'custGeom', { x: 2.04, y: 0.476, w: 0.227, h: 1.458, fill: { color: C.purple }, line: NONE,
    rotate: 44.4, points: [{ x: 0.11, y: 0 }, { x: 0.227, y: 0.3 }, { x: 0.14, y: 1.458 },
                           { x: 0.09, y: 1.458 }, { x: 0, y: 0.3 }, { close: true }] });
  text(s, 'Main Idea', { x: 3.088, y: 0.799, w: 2.62, h: 0.572, fontFace: F.head, fontSize: 28,
    bold: true, valign: 'bottom' });
  body(s, 3.088, 1.402, 2.621, 1.279, L.idea, { paraSpaceAfter: 6 });

  iconDisc(s, 'gear', 6.425, 0.778, 0.653, C.purple);
  dataText(s, 7.201, 0.841, 1.627, 'Description', '21.781', { bold: true });
  text(s, '+5%', { x: 7.342, y: 2.13, w: 0.415, h: 0.269, fontFace: F.head, fontSize: 10,
    bold: true, color: C.purple, align: 'center' });
  iconArrowUp(s, 7.201, 2.149, 0.141, 0.226, C.purple);
  note(s, 7.788, 2.13, 0.971, 0.269, 'Lorem Ipsum');
  donut(s, 8.896, 0.671, 2.263, 2.048, 2, 11, C.purple);
  text(s, '7.3', { x: 9.491, y: 1.434, w: 1.097, h: 0.404, fontFace: F.head, fontSize: 18,
    bold: true, align: 'center', valign: 'bottom' });
  note(s, 9.491, 1.838, 1.097, 0.269, 'Lorem ipsum', { align: 'center' });
  body(s, 11.244, 0.891, 1.369, 1.616, L.stdBook);

  chart(s, 'line', [{ name: 'Series 1', labels: MON1, values: [4.3, 2.5, 3.5, 4.5, 3, 2, 4, 5, 6, 4, 2, 3] }], {
    x: 0.76, y: 5.935, w: 7.028, h: 0.953, chartColors: [C.purple], lineSize: 1.5,
    lineDataSymbol: 'circle', lineDataSymbolSize: 10, lineDataSymbolLineColor: C.purple,
    valAxisHidden: true, catAxisLineShow: false, catGridLine: grid(C.dark)
  });
  body(s, 8.366, 5.935, 4.247, 0.942, L.idea, { paraSpaceAfter: 6 });
}

/* ------------------------------------------------------------------ *
 * Slide 14 -- mind-map of circles
 * ------------------------------------------------------------------ */
function slide14(s) {
  const nodes = [
    [2.851, 0.559, 1.772, C.teal,   'Lorem Ipsum',  3.153, 1.154, null],
    [0.788, 2.400, 1.772, C.blue2,  'Lorem Ipsum',  1.103, 3.000, null],
    [4.852, 0.237, 1.772, C.blue,   'Lorem Ipsum',  4.947, 0.879, 'Lorem Ipsum has the dummy text.'],
    [6.628, 1.101, 1.772, C.dark,   'Lorem Ipsum',  6.719, 1.756, 'Lorem Ipsum has the dummy text.'],
    [7.267, 2.830, 1.772, C.pink,   'Lorem Ipsum',  7.367, 3.494, 'Lorem Ipsum has the dummy text.']
  ];
  // Dashed-look connector arcs & droppers
  [[3.417, 3.286, 272.1, 358.8, false], [3.725, 3.125, 272.1, 358.8, false],
   [3.721, 1.247, 284.4, 358.8, true],  [5.841, 3.260, 272.1, 358.8, false],
   [6.047, 3.694, 272.1, 358.8, false]].forEach(function (a) {
    arc(s, a[0], a[1], 1.875, 1.874, [a[2], a[3]], C.dark, 2.5, { flipH: a[4] });
  });
  [[5.291, 4.271, 2.110], [5.599, 4.110, 2.394], [5.738, 2.162, 3.765],
   [5.841, 4.244, 2.394], [6.047, 4.678, 1.749]].forEach(function (l) {
    seg(s, l[0], l[1], 0, l[2], C.dark, 2.5);
  });
  seg(s, 2.707, 3.286, 1.612, 0, C.dark, 2.5);

  nodes.forEach(function (n) {
    oval(s, n[0], n[1], n[2], n[2], n[3]);
    text(s, n[4], { x: n[5], y: n[6], w: 1.563, h: 0.337, fontFace: F.head, fontSize: 14,
      color: C.white, align: 'center', valign: n[7] ? 'bottom' : 'middle' });
    if (n[7]) body(s, n[5], n[6] + 0.235, 1.563, 0.438, n[7], { color: C.white, align: 'center' });
  });
  donut(s, 2.852, 0.651, 1.745, 1.604, 8, 11, C.white, mix(C.white, C.teal, 0.10), 90);
  donut(s, 0.815, 2.497, 1.745, 1.604, 8, 11, C.white, mix(C.white, C.blue2, 0.10), 90);
  body(s, 2.679, 3.515, 1.563, 0.438, 'Lorem Ipsum has the dummy text.', { align: 'center' });
  body(s, 0.77, 0.824, 1.973, 1.279, L.spec);

  dataText(s, 9.513, 0.81, 2.728, 'Data text', '11 330', { bold: true });
  chart(s, 'bar', [{ name: 'stats', labels: WEEK, values: WAVE }], {
    x: 9.513, y: 1.588, w: 3.1, h: 2.513, barDir: 'bar', barGrouping: 'clustered',
    barGapWidthPct: 150, chartColors: [C.blue], valGridLine: grid(C.dark)
  });
  pill(s, 11.808, 1.147, 0.821, 0.335, '+19%');

  band(s, 4.354, 2.738);
  quote(s, 0.751, 4.767, 7.037, 1.919, [
    ['" No one who does ', C.white], ['good work ', C.purple],
    ['will ever come to a bad end, either here or in the world to come." ', C.white]
  ], 36, C.white);
  iconRow(s, function (x, col) { iconPerson(s, x, 5.123, 0.266, 0.456, col); },
    8.795, 0, 0.3409, 10, 7, C.purple, FAINT_WHITE_ON_BLUE);
  body(s, 8.234, 5.777, 4.371, 0.606, L.spec, { color: C.white, align: 'center', paraSpaceAfter: 6 });
}

/* ------------------------------------------------------------------ *
 * Slide 15 -- dashboard of six donuts
 * ------------------------------------------------------------------ */
function slide15(s) {
  cloudShape(s, 7.065, 0.152, 5.868, 5.492, FAINT_DARK_ON_WHITE);
  rulerBar(s, 0.76, 0.789, 5.87, 0.983, {});
  body(s, 0.797, 1.918, 5.83, 0.942, L.ideaLong, { paraSpaceAfter: 6 });

  const smalls = [
    [1.105, C.blue,   'gear',      1.503, 0.367, 8],
    [3.159, C.purple, 'folder',    3.557, 0.300, 2],
    [5.214, C.teal,   'people',    5.611, 0.300, 25],
    [7.268, C.pink,   'calendar',  7.665, 0.366, 7],
    [9.323, C.blue2,  'diamond',   9.719, 0.366, 30],
    [11.353, C.blue3, 'diamond',  11.749, 0.366, 4]
  ];
  smalls.forEach(function (d, i) {
    donut(s, d[0], 3.139, 1.176, 1.064, d[5], 11, d[1]);
    if (d[2] === 'diamond') iconDiamond(s, d[3] + 0.04, 3.52, 0.29, C.dark);
    else glyph(s, d[2], d[3] + 0.04, 3.52 + (0.29 - d[4] * 0.8) / 2, 0.29, d[4] * 0.8, C.dark);
    const cx = [0.993, 3.074, 5.125, 7.175, 9.225, 11.254][i];
    note(s, cx, 4.349, 1.386, 0.438, L.text5, { align: 'center', paraSpaceAfter: 6 });
  });
  statBlock(s, 7.665, 0.885, 2.392, 'Lorem Ipsum is text', '100%', L.short, { align: 'center', bold: true });
  chart(s, 'doughnut', [{ name: 'Sales', labels: ['A', 'B', 'C', 'D', 'E', 'K'], values: [14, 11, 20, 18, 15, 10] }], {
    x: 10.468, y: 0.989, w: 1.756, h: 1.664, holeSize: 80,
    chartColors: [C.blue, C.purple, C.teal, C.pink, C.blue2, C.blue3], dataNoEffects: true
  });

  band(s, 5.103, 2.397);
  quote(s, 1.482, 5.607, 7.558, 1.313,
    [['"Cultivate vigour, patience, will, purity; avoid malice and pride." ', C.white]], 36, C.white);
  pill(s, 0.228, 6.066, 1.44, 0.394, 'Krishna', { fill: C.white, color: C.teal, rotate: 270 });
  gauge(s, 9.417, 5.474, C.white, BAR_REST_ON_BLUE);
}

/** Soft cloud silhouette (overlapping discs) used as an artwork stand-in. */
function cloudShape(s, x, y, w, h, color) {
  shp(s, 'cloud', { x, y, w, h, fill: { color }, line: NONE });
}

/* ------------------------------------------------------------------ *
 * Slide 16 -- column chart header + stacked bars
 * ------------------------------------------------------------------ */
function slide16(s) {
  band(s, 0, 2.738);
  pill(s, 0.228, 1.165, 1.44, 0.394, '$345k', { fill: C.white, color: C.blue, rotate: 270 });
  body(s, 1.344, 0.739, 4.811, 1.279, L.full, { color: C.white, valign: 'middle' });
  chart(s, 'bar', [{ name: 'Progress', labels: ['M', 'T', 'W', 'Th', 'F'], values: [3, 5, 3.5, 5, 4] }], {
    x: 6.667, y: 0.547, w: 5.946, h: 1.811, barDir: 'col', barGrouping: 'clustered',
    barGapWidthPct: 135, chartColors: [C.white],
    catAxisLineColor: C.white, catAxisLabelColor: C.white, valAxisHidden: true
  });
  iconDisc(s, 'note', 9.256, 0.361, 0.756, C.white, C.blue);

  quote(s, 0.76, 3.05, 5.394, 3.736, [
    ['"Whatever happened, '], ['happened for the good', C.blue],
    ['. Whatever is happening, is happening for the good. Whatever will happen, will also '],
    ['happen for the good', C.blue], ['." ']
  ], 36);
  chart(s, 'bar', [
    { name: 'A', labels: WEEK, values: [4.3, 2.5, 3.5, 4.5, 3, 2, 4] },
    { name: 'B', labels: WEEK, values: [3, 1, 3, 4.2, 2, 1.9, 3] },
    { name: 'C', labels: WEEK, values: [4.5, 4, 5, 2, 4, 3, 4] },
    { name: 'S', labels: WEEK, values: [2, 3, 2, 4, 3, 4, 2] }
  ], {
    x: 8.657, y: 2.832, w: 3.956, h: 3.99, barDir: 'bar', barGrouping: 'stacked',
    barGapWidthPct: 75, chartColors: [C.blue, C.purple, C.teal, C.pink],
    valAxisLineShow: false, valGridLine: grid(C.dark, 0.5)
  });
  chart(s, 'bar', [
    { name: 'Stage I',  labels: ['Progress'], values: [389] },
    { name: 'Stage II', labels: ['Progress'], values: [204] },
    { name: 'Stage III', labels: ['Progress'], values: [297] },
    { name: 'Column3',  labels: ['Progress'], values: [200] }
  ], {
    x: 6.667, y: 2.832, w: 1.632, h: 3.99, barDir: 'col', barGrouping: 'percentStacked',
    barGapWidthPct: 75, chartColors: [C.blue, C.purple, C.teal, C.pink],
    catAxisHidden: true, valAxisLabelFormatCode: '0%', valAxisLabelPos: 'low',
    valAxisLineColor: C.dark,
    showValue: true, dataLabelColor: C.white, dataLabelFontSize: 10, dataLabelFormatCode: '#,##0'
  });
}

/* ------------------------------------------------------------------ *
 * Slide 17 -- twin donuts + people row
 * ------------------------------------------------------------------ */
function slide17(s) {
  quote(s, 1.3, 0.771, 4.854, 1.919, [
    ['"'], ['Cultivate', C.blue], [' vigour, patience, will, purity; avoid malice and pride.." ']
  ], 36);
  [[6.614, C.blue, 5, 7.003, 7.028], [9.791, C.purple, 11, 10.179, 10.204]].forEach(function (d, i) {
    donut(s, d[0], 0.708, 1.925, 1.824, i === 0 ? 9 : 5, 11, d[1]);
    text(s, '$22k', { x: d[3], y: 1.218, w: 1.097, h: 0.404, fontFace: F.head, fontSize: 18,
      bold: true, align: 'center', valign: 'bottom' });
    pill(s, d[4], 1.675, 1.04, 0.315, 'Lorem ipsum', { fill: d[1], size: 10, font: F.body, spc: 0 });
  });
  columnMeter(s, 8.282, 0.606, 1.169, 2.029, { pct: 2, rest: 1, log: true });
  columnMeter(s, 11.458, 0.606, 1.169, 2.029, { pct: 2, rest: 1, fg: C.purple, log: true });

  band(s, 2.937, 2.738);
  pill(s, 0.108, 1.478, 1.699, 0.394, 'Krishna', { rotate: 270 });
  ideaBlock(s, 0.761, 4.042, 5.391, 'Main Idea', L.book, { bodyH: 0.606 });
  iconRow(s, function (x, col) { iconPerson(s, x, 3.339, 0.266, 0.456, col); },
    0.779, 0, 0.3409, 10, 7, C.purple, FAINT_WHITE_ON_BLUE);
  iconRow(s, function (x, col) { iconStar(s, x, 6.346, 0.244, col); }, 9.229, 0, 0.2846, 5, 0, C.white, C.white);
  chart(s, 'bar', [
    { name: 'PA', labels: ['I q', 'II q', 'III q', 'IV q'], values: [2301, 4211, 3011, 2198] },
    { name: 'PB', labels: ['I q', 'II q', 'III q', 'IV q'], values: [2301, 4211, 3011, 2198] }
  ], {
    x: 7.179, y: 3.315, w: 5.453, h: 2.058, barDir: 'bar', barGrouping: 'stacked',
    barGapWidthPct: 95, chartColors: [C.white, C.purple],
    catAxisLineColor: C.white, catAxisLabelColor: C.white, valAxisLineColor: C.white,
    valAxisLabelColor: C.white, valGridLine: grid(C.white, 0.75)
  });
  body(s, 7.179, 6.034, 5.434, 0.774, L.mid);
  chart(s, 'bar', [{ name: 'Series 1', labels: MON1, values: [4.3, 2.5, 3.5, 4.5, 3, 2, 4, 5, 6, 4, 2, 3] }], {
    x: 0.779, y: 5.923, w: 5.376, h: 1.036, barDir: 'col', barGrouping: 'clustered',
    barGapWidthPct: 150, chartColors: [C.blue],
    valAxisHidden: true, catAxisLineShow: false
  });
}

/* ------------------------------------------------------------------ *
 * Slide 18 -- growth columns with trend arrow
 * ------------------------------------------------------------------ */
function slide18(s) {
  s.background = { color: C.page };
  quote(s, 6.667, 0.797, 5.946, 3.803, [
    ['"The struggle for \n', C.navy], ['self-mastery', C.blue],
    [' that every human being must wage if he or she is to emerge from life victorious." ', C.navy]
  ], 44, C.navy);
  pill(s, 4.993, 1.761, 2.322, 0.394, 'Krishna said', { rotate: 270 });
  chart(s, 'bar', [{ name: 'A', labels: ['Week 1', 'Week 2', 'Week 3', 'Week4'], values: [2, 4, 6, 7] }], {
    x: 6.667, y: 4.827, w: 5.622, h: 1.994, barDir: 'col', barGrouping: 'stacked',
    barGapWidthPct: 130, chartColors: [C.blue],
    catAxisLineColor: C.navy, catAxisLabelColor: C.navy,
    valAxisLineColor: C.navy, valAxisLabelColor: C.navy, valAxisLabelPos: 'low',
    valAxisMajorTickMark: 'out', catAxisLabelFontSize: 12, valAxisLabelFontSize: 12,
    valGridLine: { color: mix(C.navy, C.page, 0.30), size: 1, style: 'solid' }
  });
  s.addShape('line', { x: 7.634, y: 4.827, w: 3.943, h: 1.049, flipV: true,
    line: { color: C.blue, width: 5, cap: 'round', endArrowType: 'triangle' } });
}

/* ------------------------------------------------------------------ *
 * Slide 19 -- blue quote panel + KPI dashboard
 * ------------------------------------------------------------------ */
function slide19(s) {
  rect(s, 0.374, 0, 6.293, H, C.blue);
  pill(s, 0.76, 0.674, 1.357, 0.335, 'Krishna', { fill: C.white, color: C.dark });
  quote(s, 1.216, 1.417, 4.609, 3.736,
    [['"His judgement will be better and his vision clear if he is not emotionally entangled in the outcome of what he does." ', C.white]],
    36, C.white);
  body(s, 1.216, 5.388, 4.609, 0.774, L.book, { color: C.white, paraSpaceAfter: 6 });
  iconRow(s, function (x, col) { iconStar(s, x, 6.576, 0.244, col); }, 2.806, 0, 0.2846, 5, 0, C.white, C.white);

  statBlock(s, 7.35, 0.806, 2.861, 'Lorem Ipsum is text', '100%', L.short, { align: 'center', bold: true });
  chart(s, 'doughnut', [{ name: 'Sales', labels: ['A', 'B', 'C'], values: [14, 11, 20] }], {
    x: 10.48, y: 0.674, w: 2.041, h: 1.933, holeSize: 80,
    chartColors: [C.blue, C.purple, C.teal], dataNoEffects: true
  });
  oval(s, 12.272, 0.737, 0.349, 0.349, C.blue);
  oval(s, 11.192, 1.349, 0.597, 0.597, C.blue);
  shp(s, 'star5', { x: 11.375, y: 1.538, w: 0.239, h: 0.22, fill: { color: C.white } });

  [[7.411, C.blue, 'people', '259hr'], [9.172, C.purple, 'dollar', '$45k'], [10.936, C.teal, 'briefcase', '$45k']].forEach(function (k, i) {
    iconDisc(s, k[2], k[0], 3.043, 0.553, k[1]);
    const tx = [7.991, 9.753, 11.517][i];
    text(s, k[3], { x: tx, y: 3.117, w: 1.028, h: 0.337, fontFace: F.head, fontSize: 14, valign: 'bottom' });
    note(s, tx, 3.422, 1.028, 0.269, 'Sample text');
  });
  chart(s, 'bar', [
    { name: 'Series 1', labels: QTR, values: SER3[0] },
    { name: 'Series 2', labels: QTR, values: SER3[1] },
    { name: 'Series 3', labels: QTR, values: SER3[2] }
  ], {
    x: 7.411, y: 3.995, w: 5.202, h: 2.835, barDir: 'col', barGrouping: 'stacked',
    barGapWidthPct: 150, chartColors: [C.blue, C.purple, C.teal],
    catAxisMajorTickMark: 'out', valAxisMajorTickMark: 'out', valGridLine: grid(C.dark, 0.5)
  });
}

/* ------------------------------------------------------------------ *
 * Slide 20 -- data dashboard + right blue panel
 * ------------------------------------------------------------------ */
function slide20(s) {
  rect(s, 7.561, 0, 5.398, H, C.blue);
  pill(s, 8.229, 0.797, 1.009, 0.335, '+5%', { fill: C.white, color: C.dark });
  text(s, '24/7', { x: 8.229, y: 1.446, w: 2.647, h: 0.505, fontFace: F.head, fontSize: 24,
    color: C.white, valign: 'bottom' });
  body(s, 8.229, 1.992, 2.647, 0.774, L.scrap, { color: C.white });
  iconRow(s, function (x, col) { iconStar(s, x, 3.093, 0.244, col); }, 8.337, 0, 0.2846, 5, 0, C.white, C.white);
  columnMeter(s, 11.09, 0.64, 1.532, 2.756, { pct: 100, rest: 2000, gap: 75, log: true,
    fg: C.purple, bg: BAR_REST_ON_BLUE, showCat: true, axColor: C.white, axLine: false,
    ticks: 'none', grid: mix(C.white, C.blue, 0.30) });

  const cards = [
    [0.898, C.blue,   '170hr',    0.802, 'gear',      1.586, 0.367, 8],
    [2.938, C.purple, '$23.000',  2.842, 'dollar',    3.625, 0.366, 8],
    [4.992, C.purple, '> 17 850', 4.897, 'briefcase', 5.692, 0.280, 14]
  ];
  cards.forEach(function (c, i) {
    pill(s, c[0], 0.806, 1.741, 0.406, c[2], { fill: c[1], color: C.white });
    text(s, 'Description', { x: c[0] + 0.06, y: 1.439, w: 1.62, h: 0.303,
      fontFace: F.head, fontSize: 12, align: 'center' });
    body(s, c[0] + 0.058, 1.729, 1.626, 0.942, L.printer, { align: 'center', paraSpaceAfter: 6 });
    donut(s, c[3], 2.672, 1.933, 1.655, c[7], 11, i === 2 ? C.teal : c[1]);
    glyph(s, c[4], c[5], 3.316 + (0.366 - c[6]) / 2, 0.366, c[6], C.dark);
  });
  text(s, 'Data Dashboard', { x: 0.802, y: 4.502, w: 5.874, h: 0.337,
    fontFace: F.head, fontSize: 14, bold: true, valign: 'bottom' });
  chart(s, 'bar', [
    { name: 'Series 1', labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'], values: [4, 3, 4, 5, 7, 6] },
    { name: 'Series 2', labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'], values: [3, 5, 5, 6, 5, 5] },
    { name: 'Series 3', labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'], values: [2, 4, 3, 2, 3, 4] }
  ], {
    x: 0.76, y: 4.929, w: 5.916, h: 1.892, barDir: 'bar', barGrouping: 'stacked',
    barGapWidthPct: 100, chartColors: [C.blue, C.purple, C.teal],
    catAxisMajorTickMark: 'out', valAxisMajorTickMark: 'out', catGridLine: grid(C.dark)
  });
  chart(s, 'bar', [{ name: 'Series 1',
    labels: ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'],
    values: [4.3, 2.5, 3.5, 4.5, 3, 2, 4, 5, 6, 4, 3, 2] }], {
    x: 7.953, y: 3.75, w: 4.66, h: 3.071, barDir: 'col', barGrouping: 'clustered',
    barGapWidthPct: 150,
    chartColors: [C.white, C.white, C.white, C.teal, C.white, C.white, C.white,
                  C.pink, C.purple, C.white, C.white, C.white],
    valAxisHidden: true, catAxisLineShow: false, catAxisLabelColor: C.white
  });
  [[10.952, 3.716, C.purple, 'safe'], [9.127, 4.255, C.teal, 'note'], [10.584, 4.091, C.pink, 'camera']].forEach(function (b) {
    oval(s, b[0], b[1], 0.493, 0.493, C.white);
    glyph(s, b[3], b[0] + 0.12, b[1] + 0.12, 0.253, 0.253, b[2]);
  });
}

/* ------------------------------------------------------------------ *
 * Slide 21 -- rocket illustration
 * ------------------------------------------------------------------ */
function slide21(s) {
  s.addText(rich([['"', null, { bold: true }], ['He who experiences the '],
    ['unity of life ', C.blue, { bold: true }],
    ['sees his own Self in all beings, and all beings in his own Self, and looks on everything with an impartial eye.'],
    ['" ', null, { bold: true }]], { fontFace: F.head, fontSize: 36 }),
    { x: 0.76, y: 0.771, w: 7.943, h: 2.524, fontFace: F.head, fontSize: 36, color: C.dark });
  body(s, 0.797, 3.529, 6.382, 0.774, L.ideaLong, { paraSpaceAfter: 6 });
  rocket(s, 9.754, 0.516, 2.334, 3.062);
  text(s, '44 890', { x: 11.022, y: 2.91, w: 2.482, h: 0.404, fontFace: F.head, fontSize: 18,
    charSpacing: 3, color: C.gray, valign: 'bottom', rotate: 90 });

  band(s, 4.762, 2.738);
  progressBar(s, 7.179, 5.339, 5.444, 1.313, {});
  pill(s, 0, 5.936, 1.986, 0.394, 'Krishna', { fill: C.white, color: C.blue, rotate: 270 });
  ideaBlock(s, 3.485, 5.335, 3.18, 'Main Idea', L.idea, { bodyH: 1.111 });
  badge(s, 9.932, 5.509, 0.731, '100%');
  chart(s, 'pie', [{ name: 'Sales', labels: ['Market Saze A', 'Market Saze B'], values: [13, 11] }], {
    x: 1.274, y: 5.198, w: 2.007, h: 1.845, firstSliceAng: 0,
    chartColors: [C.white, BAR_REST_ON_BLUE], dataNoEffects: true
  });
  text(s, '+5%', { x: 2.512, y: 6.009, w: 0.415, h: 0.269, fontFace: F.head, fontSize: 10,
    bold: true, color: C.darkBlue, align: 'center' });
  iconArrowUp(s, 2.371, 6.028, 0.141, 0.226, C.purple);
}

/** Outlined rocket with fins, portholes and exhaust trails. */
function rocket(s, x, y, w, h) {
  const ln = { color: C.dark, width: 2.5 };
  shp(s, 'custGeom', {
    x, y, w, h, fill: NONE, line: ln, rotate: 39.3,
    points: [{ x: w * 0.50, y: 0 }, { x: w * 0.70, y: h * 0.18 }, { x: w * 0.76, y: h * 0.55 },
             { x: w * 0.72, y: h * 0.86 }, { x: w * 0.50, y: h * 0.96 },
             { x: w * 0.28, y: h * 0.86 }, { x: w * 0.24, y: h * 0.55 },
             { x: w * 0.30, y: h * 0.18 }, { close: true }]
  });
  oval(s, 10.616, 1.297, 0.602, 0.602, C.blue);
  shp(s, 'ellipse', { x: 10.738, y: 2.032, w: 0.357, h: 0.357, fill: NONE, line: ln });
  shp(s, 'ellipse', { x: 10.804, y: 2.530, w: 0.224, h: 0.224, fill: NONE, line: ln });
  // Fins
  shp(s, 'custGeom', { x: 9.754, y: 2.448, w: 0.627, h: 1.068, fill: NONE, line: ln,
    points: [{ x: 0.627, y: 0 }, { x: 0, y: 0.75 }, { x: 0.25, y: 1.068 }, { x: 0.627, y: 0.62 }] });
  shp(s, 'custGeom', { x: 11.462, y: 2.448, w: 0.627, h: 1.068, fill: NONE, line: ln, flipH: true,
    points: [{ x: 0.627, y: 0 }, { x: 0, y: 0.75 }, { x: 0.25, y: 1.068 }, { x: 0.627, y: 0.62 }] });
  // Exhaust
  arc(s, 8.143, 2.590, 1.077, 1.249, [272.1, 358.8], C.dark, 2.5, { flipV: true, rotate: 39.1 });
  arc(s, 9.405, 3.553, 1.077, 1.250, [272.1, 358.8], C.dark, 2.5, { flipH: true, flipV: true, rotate: 38.7 });
  seg(s, 8.355, 3.430, 1.433, 1.762, C.dark, 2.5, { flipH: true });
  oval(s, 10.880, 1.366, 0.657, 0.657, C.blue);
}

/* ------------------------------------------------------------------ *
 * Slide 22 -- gauge header + monthly column chart
 * ------------------------------------------------------------------ */
function slide22(s) {
  band(s, 0, 2.738);
  pill(s, 0.228, 1.165, 1.44, 0.394, '$345k', { fill: C.white, color: C.blue, rotate: 270 });
  gauge(s, 1.266, 0.48, C.purple, BAR_REST_ON_BLUE);
  oval(s, 2.294, 1.416, 0.731, 0.731, C.white);
  glyph(s, 'people', 2.477, 1.632, 0.366, 0.3, C.purple);
  note(s, 4.407, 0.687, 2.746, 0.303, 'Lorem Ipsum is simply dummy text.', { color: C.white, fontSize: 12 });
  text(s, '73%', { x: 4.407, y: 1.165, w: 2.746, h: 1.212, fontFace: F.head, fontSize: 66,
    bold: true, color: C.white });
  body(s, 7.547, 0.842, 5.066, 1.279, L.full, { color: C.white, valign: 'middle' });

  dataText(s, 0.799, 3.008, 2.728, 'Data text', '$800k', { bold: true });
  chart(s, 'bar', [{ name: 'Performance',
    labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
    values: [7, 5, 12, 10, 14, 9, 8, 7, 5, 4, 8, 10] }], {
    x: 0.751, y: 3.867, w: 7.705, h: 2.954, barDir: 'col', barGrouping: 'clustered',
    barGapWidthPct: 150, barOverlapPct: -25, chartColors: [C.teal],
    catAxisLabelFontSize: 10, valAxisHidden: true,
    showValue: true, dataLabelFormatCode: '[$$-409]#,##0.00',
    dataLabelFontSize: 10, dataLabelColor: C.dark, dataLabelPosition: 'outEnd'
  });
  pill(s, 6.249, 4.75, 0.821, 0.335, '+15%');
  quote(s, 8.953, 3.05, 3.747, 3.736, [
    ['"True '], ['LOVE', C.blue], [' is beyond the limitations of time and space; it is eternal and unchanging." ']
  ], 36);
}

/* ------------------------------------------------------------------ *
 * Slide 23 -- area chart + bottom band
 * ------------------------------------------------------------------ */
function slide23(s) {
  pill(s, 0.76, 0.806, 1.741, 0.406, '170hr', { color: C.white });
  statBlock(s, 0.76, 1.411, 2.927, 'Lorem Ipsum is text', '100%', '', { bold: true });
  body(s, 0.76, 2.737, 2.927, 1.616, L.ideaLong, { paraSpaceAfter: 6 });
  columnMeter(s, 4.05, 0.806, 1.741, 3.608, { pct: 2, rest: 1, log: true });

  dataText(s, 6.175, 0.784, 2.728, 'Data text', '$800k', { bold: true });
  chart(s, 'area', [{ name: 'stats', labels: ['J-F', 'M-A', 'M-J', 'J-A', 'S-O', 'N-D'],
    values: [4.3, 2.5, 3.5, 5, 3, 2] }], {
    x: 6.155, y: 1.562, w: 6.759, h: 2.852, chartColors: [C.blue], valGridLine: grid(C.dark)
  });
  pill(s, 9.768, 1.654, 0.821, 0.335, '+15%');

  band(s, 4.729, 2.363);
  quote(s, 0.751, 5.267, 5.915, 1.313,
    [['"We behold what we are, and we are what we behold." ', C.white]], 36, C.white);
  iconRow(s, function (x, col) { iconPerson(s, x, 5.284, 0.266, 0.456, col); },
    8.521, 0, 0.3409, 10, 7, C.purple, FAINT_WHITE_ON_BLUE);
  body(s, 7.959, 5.938, 4.371, 0.606, L.spec, { color: C.white, align: 'center', paraSpaceAfter: 6 });
}

/* ------------------------------------------------------------------ *
 * Slide 24 -- twin KPI donuts + blue band quote
 * ------------------------------------------------------------------ */
function slide24(s) {
  band(s, 2.984, 2.738);
  [[0.76, C.blue, 1.536, 4.09, 5.151, '19.7', 10],
   [7.041, C.purple, 7.817, 10.371, 11.432, '7.3', 2]].forEach(function (k) {
    iconDisc(s, 'gear', k[0], 0.778, 0.653, k[1]);
    dataText(s, k[2], 0.841, 2.405, 'Description', '21.781', { bold: true });
    body(s, k[2], 1.778, 2.405, 0.942, L.stdBook);
    iconArrowUp(s, k[4], 1.176, 0.141, 0.226, k[1]);
    donut(s, k[3], 0.671, 2.263, 2.048, k[6], 11, k[1]);
    text(s, k[5], { x: k[3] + 0.595, y: 1.434, w: 1.097, h: 0.404, fontFace: F.head, fontSize: 18,
      bold: true, align: 'center', valign: 'bottom' });
    note(s, k[3] + 0.595, 1.838, 1.097, 0.269, 'Lorem ipsum', { align: 'center' });
  });
  quote(s, 1.396, 3.397, 6.392, 1.919,
    [['"Calmness, gentleness, silence, self-restraint and purity: these are the disciplines of the mind." ', C.white]],
    36, C.white);
  pill(s, 0.237, 4.156, 1.44, 0.394, 'Krishna', { fill: C.white, color: C.blue, rotate: 270 });
  body(s, 8.03, 3.624, 4.609, 0.774, L.book, { color: C.white, paraSpaceAfter: 6 });
  iconRow(s, function (x, col) { iconStar(s, x, 4.812, 0.244, col); }, 9.62, 0, 0.2846, 5, 0, C.white, C.white);
  note(s, 0.76, 5.935, 5.394, 0.942, L.ideaLong, { italic: false, paraSpaceAfter: 6 });
  chart(s, 'bar', [
    { name: 'Actual',  labels: ['Last month'], values: [2] },
    { name: 'Planned', labels: ['Last month'], values: [1] }
  ], {
    x: 7.179, y: 5.841, w: 5.434, h: 1.242, barDir: 'bar', barGrouping: 'percentStacked',
    barGapWidthPct: 125, chartColors: [C.blue, mix(C.dark, C.white, 0.1)],
    catAxisHidden: true, valAxisLabelFormatCode: '0%', valAxisLabelPos: 'low',
    valAxisLineShow: false, valGridLine: grid(C.dark),
    showValue: true, dataLabelPosition: 'inEnd', dataLabelColor: C.white, dataLabelFontSize: 12,
    dataLabelFormatCode: '#,##0'
  });
}

/* ------------------------------------------------------------------ *
 * Slide 25 -- pie-of-pie + small charts
 * ------------------------------------------------------------------ */
function slide25(s) {
  quote(s, 0.76, 0.901, 5.394, 4.342, [
    ['"'], ['Feelings of heat ', C.blue],
    ['and cold, pleasure and pain, are caused by the contact of the senses with their objects. They come and they go, never lasting long. You must accept them." ']
  ], 36);
  note(s, 0.76, 5.411, 5.394, 1.279,
    "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including.",
    { italic: false, paraSpaceAfter: 6 });

  chart(s, 'pie', [{ name: 'Posts', labels: ['A', 'B', 'C', 'D'], values: [8.2, 3.2, 1.4, 2] }], {
    x: 7.179, y: 0.797, w: 5.434, h: 3.398, firstSliceAng: 0,
    chartColors: [C.blue, C.purple, C.teal, C.pink], dataNoEffects: true,
    showLegend: true, legendPos: 'b', legendColor: C.dark, legendFontSize: 12
  });
  oval(s, 8.104, 1.7, 1.722, 1.722, C.white);
  pill(s, 8.332, 2.177, 1.288, 0.394, '$232k', { size: 14, spc: 0 });
  note(s, 8.332, 2.642, 1.288, 0.303, 'Lorem ipsum', { color: C.blue, fontSize: 12, align: 'center' });
  chart(s, 'pie', [{ name: 'Sales', labels: ['Australia', 'B', 'C', 'D'], values: [10, 6, 5, 3] }], {
    x: 7.179, y: 4.195, w: 2.299, h: 2.626, firstSliceAng: 0,
    chartColors: [C.blue, C.purple, C.teal, C.pink], dataNoEffects: true,
    showPercent: true, dataLabelColor: C.white, dataLabelFontSize: 12,
    dataLabelFontFace: F.head, dataLabelFormatCode: '0%'
  });
  chart(s, 'area', [{ name: 'Series 1', labels: QTR, values: [4.3, 2.5, 3.5, 4.5] }], {
    x: 9.619, y: 4.195, w: 2.994, h: 2.626, barGrouping: 'stacked', chartColors: [C.purple],
    valAxisHidden: true, catAxisLabelFontSize: 12
  });
}

/* ------------------------------------------------------------------ *
 * Slide 26 -- quote on the light page background
 * ------------------------------------------------------------------ */
function slide26(s) {
  s.background = { color: C.page };
  quote(s, 1.523, 2.783, 9.092, 2.322, [
    ['"Why do you worry unnecessarily? \nWhom do you fear? Who can kill you? \n', C.navy],
    ['The soul is neither born nor dies', C.blue], ['." ', C.navy]
  ], 44, C.navy);
  pill(s, -0.162, 3.67, 2.322, 0.394, 'Krishna said', { rotate: 270 });
  chart(s, 'pie', [{ name: 'Sales', labels: ['Market Saze A', 'Market Saze B'], values: [8, 11] }], {
    x: 10.027, y: 2.783, w: 2.588, h: 2.379, firstSliceAng: 0,
    chartColors: [C.white, mix(C.purple, C.page, 0.1)], dataNoEffects: true
  });
  pill(s, 11.533, 3.233, 0.899, 0.259, '> 10', { fill: C.purple, size: 10 });
  text(s, '+15%', { x: 11.714, y: 3.597, w: 0.415, h: 0.269, fontFace: F.head, fontSize: 10,
    color: C.darkBlue, align: 'center' });
  iconArrowUp(s, 11.573, 3.616, 0.141, 0.226, C.purple);
}

/* ------------------------------------------------------------------ *
 * Slide 27 -- arrow ribbons with +65 / +37
 * ------------------------------------------------------------------ */
function slide27(s) {
  s.addText(rich([['"', null, { bold: true }], ['No one that '], ['does', C.blue], [' '], ['good', C.blue],
    [' work will ever come to a terrible ending, either in the world to come.'], ['" ', null, { bold: true }]],
    { fontFace: F.head, fontSize: 36 }),
    { x: 0.76, y: 0.771, w: 5.394, h: 2.524, fontFace: F.head, fontSize: 36, color: C.dark });
  body(s, 0.797, 3.328, 5.364, 0.942, L.ideaLong, { paraSpaceAfter: 6 });
  rulerBar(s, 7.179, 0.789, 5.434, 0.993, {});
  body(s, 7.179, 2.016, 5.434, 0.774, L.mid, { paraSpaceAfter: 6 });
  chart(s, 'bar', [
    { name: 'Actual',  labels: ['2024'], values: [1900] },
    { name: 'Planned', labels: ['2024'], values: [2000] }
  ], {
    x: 7.179, y: 3.002, w: 5.434, h: 1.321, barDir: 'bar', barGrouping: 'percentStacked',
    barGapWidthPct: 75, chartColors: [C.purple, C.teal],
    catAxisHidden: true, valAxisLabelFormatCode: '0%', valAxisLabelPos: 'high',
    valAxisLineShow: false, valGridLine: grid(C.dark, 0.5)
  });
  iconDisc(s, 'safe', 8.846, 3.376, 0.805, C.purple);
  iconDisc(s, 'note', 10.855, 3.376, 0.805, C.teal);

  band(s, 4.537, 2.963);
  ribbonArrow(s, { color: C.purple, dir: 'left', label: 'Subtitle Text', num: '02',
    barY: 5.214, barL: 4.065, barR: 5.619, headX: 3.710, tabX: 5.324 });
  ribbonArrow(s, { color: C.teal, dir: 'right', label: 'Subtitle Text', num: '03',
    barY: 5.664, barL: 6.848, barR: 9.445, headX: 9.446, tabX: 6.313 });
  text(s, '+65%', { x: 1.174, y: 4.987, w: 2.354, h: 1.01, fontFace: F.head, fontSize: 54,
    bold: true, color: C.white, align: 'center' });
  text(s, '+37%', { x: 9.901, y: 5.475, w: 2.354, h: 1.01, fontFace: F.head, fontSize: 54,
    bold: true, color: C.white, align: 'center' });
  note(s, 1.174, 6.06, 3.991, 0.774, L.idea, { color: C.white, paraSpaceAfter: 6 });
  note(s, 7.628, 6.514, 4.627, 0.438, L.scrap2, { color: C.white, paraSpaceAfter: 6 });
}

/**
 * Banner arrow: flat bar `barL..barR`, a slanted tail on the far side, a big
 * arrow head at `headX` pointing `dir`, and a numbered tab running to the
 * bottom edge of the slide.
 */
function ribbonArrow(s, o) {
  const barH = 0.616, barY = o.barY, left = o.dir === 'left';
  rect(s, o.barL, barY, o.barR - o.barL, barH, o.color);
  shp(s, 'rtTriangle', { x: left ? o.barR : o.barL - 0.535, y: barY, w: 0.535, h: barH,
    fill: { color: o.color }, line: NONE, flipV: true, flipH: !left });
  shp(s, 'triangle', { x: o.headX, y: barY - 0.144, w: 0.357, h: 0.903,
    fill: { color: o.color }, line: NONE, rotate: left ? 270 : 90 });
  rect(s, o.tabX, barY + barH, 0.831, H - barY - barH, mix(o.color, C.dark, 0.80));
  s.addText(o.num, { x: o.tabX, y: H - 0.72, w: 0.831, h: 0.55, color: C.white,
    fontFace: F.head, fontSize: 24, align: 'center', valign: 'middle', margin: 0 });
  text(s, o.label, { x: o.barL + 0.15, y: barY, w: o.barR - o.barL - 0.3, h: barH,
    fontFace: F.head, fontSize: 16, color: C.white, valign: 'middle', wrap: false,
    align: left ? 'right' : 'left' });
}

/* ------------------------------------------------------------------ *
 * Slide 28 -- percent-stacked area
 * ------------------------------------------------------------------ */
function slide28(s) {
  band(s, 0, 2.738);
  pill(s, 0.228, 1.165, 1.44, 0.394, '$345k', { fill: C.white, color: C.blue, rotate: 270 });
  body(s, 1.344, 0.739, 4.811, 1.279, L.full, { color: C.white, valign: 'middle' });
  chart(s, 'bar', [{ name: 'data', labels: ['Iq', 'II q', 'III q', 'Ivq'], values: [0.3, 0.15, 0.35, 0.2] }], {
    x: 7.179, y: 0.38, w: 5.403, h: 1.962, barDir: 'bar', barGrouping: 'clustered',
    barGapWidthPct: 100, chartColors: [C.white, C.purple, C.white, C.white],
    catAxisLabelColor: C.white, catAxisLabelPos: 'low', catAxisOrientation: 'maxMin',
    valAxisLabelFormatCode: '0%', valAxisLabelColor: C.white, valAxisLogScaleBase: 10,
    valAxisLineShow: false, catAxisMajorTickMark: 'out', valAxisMajorTickMark: 'out',
    valGridLine: { color: mix(C.dark, C.blue, 0.30), size: 0.75, style: 'solid' }
  });
  quote(s, 0.76, 3.05, 5.394, 3.736, [
    ['"When a person '], ['responds to the joys and sorrows of others', C.blue],
    [' as though they were his own, he or she has attained the highest spiritual union." ']
  ], 36);
  text(s, 'KRISHNA said', { x: 4.913, y: 5.371, w: 2.482, h: 0.404, fontFace: F.head, fontSize: 18,
    charSpacing: 3, color: C.gray, valign: 'bottom', rotate: 90 });
  body(s, 7.179, 3.161, 5.434, 0.774, L.mid, { paraSpaceAfter: 6 });
  const pct = toPercent(SER3);
  chart(s, 'area', [
    { name: 'Series 1', labels: QTR, values: pct[0] },
    { name: 'Series 2', labels: QTR, values: pct[1] },
    { name: 'Series 3', labels: QTR, values: pct[2] }
  ], {
    x: 7.179, y: 4.092, w: 5.434, h: 2.738, barGrouping: 'stacked',
    chartColors: [C.blue, C.purple, C.teal], valAxisLabelFormatCode: '0%', valAxisMaxVal: 1,
    catAxisMajorTickMark: 'out', valAxisMajorTickMark: 'out', valGridLine: grid(C.dark, 0.5)
  });
}

/* ------------------------------------------------------------------ *
 * Slide 29 -- climbing path with milestones
 * ------------------------------------------------------------------ */
function slide29(s) {
  quote(s, 1.325, 0.771, 4.83, 4.544, [
    ['"'], ['Fear not', C.blue],
    ['. What is not real, never was and never will be. What’s true, always was and cannot be destroyed." ']
  ], 44);
  pill(s, 0.036, 1.549, 1.842, 0.394, 'Krishna said', { rotate: 270 });
  donut(s, 10.604, 1.018, 2.009, 1.818, 8, 11, C.blue);
  glyph(s, 'gear', 11.419, 1.742, 0.366, 0.367, C.dark);

  // Rising polyline with alternating milestone dots
  [[0.778, 6.306, 3.266, 0, false], [4.195, 5.361, 1.179, 1.065, true],
   [5.527, 5.361, 0.393, 0.223, false], [6.024, 4.796, 0.405, 0.689, true],
   [6.469, 4.746, 0.437, 0, false], [7.421, 4.746, 0.477, 0, false],
   [8.681, 4.178, 0.196, 0.233, false], [8.999, 3.105, 0.963, 1.264, true],
   [10.195, 3.059, 0.598, 0.271, false], [10.775, 2.215, 0.818, 1.115, true]].forEach(function (l) {
    seg(s, l[0], l[1], l[2], l[3], C.dark, 2.5, { flipV: l[4] });
  });
  bicycle(s, 6.880, 1.121, 1.818, 3.629);
  [[4.016, 6.246, C.purple], [5.237, 5.127, C.dark], [5.770, 5.459, C.purple],
   [6.290, 4.577, C.dark], [7.796, 4.567, C.dark], [8.762, 4.317, C.dark],
   [9.860, 2.856, C.purple]].forEach(function (o) { oval(s, o[0], o[1], 0.358, 0.358, o[2]); });

  text(s, '01', { x: 2.889, y: 6.238, w: 1.127, h: 0.909, fontFace: F.head, fontSize: 48,
    color: C.purple, align: 'right', valign: 'middle' });
  text(s, '02', { x: 6.403, y: 4.893, w: 1.127, h: 0.909, fontFace: F.head, fontSize: 48,
    color: C.purple, valign: 'middle' });
  text(s, '03', { x: 9.898, y: 3.453, w: 1.127, h: 0.909, fontFace: F.head, fontSize: 48,
    color: C.purple, valign: 'middle' });
  body(s, 7.28, 5.129, 2.006, 0.438, L.everSince, { valign: 'middle' });
  body(s, 10.775, 3.605, 1.609, 0.606, L.everSince, { valign: 'middle' });

  band(s, 5.694, 1.806);
  progressBar(s, 7.179, 6.0, 5.434, 1.054, {});
  body(s, 4.468, 6.278, 2.159, 0.774,
    "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley. ",
    { color: C.white, paraSpaceAfter: 6 });
  note(s, 0.754, 6.366, 2.159, 0.606,
    "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s. ",
    { color: C.white, paraSpaceAfter: 6 });
}

/** Outline cyclist stand-in for the photo on slide 29. */
function bicycle(s, x, y, w, h) {
  const ln = { color: C.dark, width: 2.5 };
  shp(s, 'ellipse', { x: x + w * 0.10, y: y + h * 0.05, w: w * 0.30, h: h * 0.16, fill: NONE, line: ln });
  shp(s, 'custGeom', { x, y: y + h * 0.18, w, h: h * 0.52, fill: NONE, line: ln,
    points: [{ x: w * 0.24, y: 0 }, { x: w * 0.52, y: h * 0.10 }, { x: w * 0.80, y: h * 0.24 },
             { x: w * 0.62, y: h * 0.32 }, { x: w * 0.30, y: h * 0.30 }, { x: w * 0.20, y: h * 0.52 }] });
  shp(s, 'rect', { x: x + w * 0.18, y: y + h * 0.68, w: w * 0.40, h: h * 0.16, fill: NONE, line: ln });
  shp(s, 'rect', { x: x + w * 0.10, y: y + h * 0.84, w: w * 0.46, h: h * 0.16, fill: NONE, line: ln });
}

/* ------------------------------------------------------------------ *
 * Slide 30 -- closing quote
 * ------------------------------------------------------------------ */
function slide30(s) {
  s.background = { color: C.page };
  quote(s, 2.253, 3.922, 8.727, 2.322, [
    ['"Krishana Didn’t make a mistake when he made you. You need to see ourself as a ', C.navy],
    ['Krishna sees you', C.blue], ['." ', C.navy]
  ], 44, C.navy);
  pill(s, 0.965, 4.886, 1.842, 0.394, 'Krishna said', { rotate: 270 });
}

/* ------------------------------------------------------------------ *
 * Assemble & write
 * ------------------------------------------------------------------ */
const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
                  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
                  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30];

const pptx = new pptxgen();
pptx.defineLayout({ name: 'WIDE', width: W, height: H });
pptx.layout = 'WIDE';
pptx.author = 'Krishna said';
pptx.title = 'Krishna said';

BUILDERS.forEach(function (build) { build(pptx.addSlide()); });

pptx.writeFile({ fileName: path.join(__dirname, '043a69a5-a31b-4fa0-b895-29c30943cb08_grok_final.pptx') })
  .then(function (f) { console.log('wrote ' + f); })
  .catch(function (e) { console.error(e); process.exit(1); });
