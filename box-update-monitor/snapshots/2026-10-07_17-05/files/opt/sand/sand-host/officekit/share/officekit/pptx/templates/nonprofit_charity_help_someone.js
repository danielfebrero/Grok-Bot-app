/**
 * "Charl" presentation template - rebuilt with pptxgenjs.
 *
 * Run:  node 0b7f592f-374d-4500-a995-3168e9b75762_grok_final.js
 * Out:  0b7f592f-374d-4500-a995-3168e9b75762_grok_final.pptx (next to this file)
 *
 * Raster artwork in the source deck (mock-ups / photos) is replaced by a
 * light placeholder rectangle labelled "[image]".
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette
const SKY       = 'B1D9ED'; // accent1 - pale blue panels
const SKY_MID   = '8FC9E5';
const SKY_DEEP  = '69B7DC';
const BLUE      = '267BA4'; // accent2 - primary brand blue
const BLUE_LINK = '267BA5';
const AZURE     = '3DA1D2'; // accent3
const STEEL     = '1C5C7B';
const NAVY      = '133E52';
const WHITE     = 'FFFFFF';
const INK       = '404040'; // headline grey
const CHARCOAL  = '2B2B2B';
const GRAY      = '959595'; // body copy
const GHOST     = 'EAEAEA'; // oversized step numerals
const TRACK     = 'E6E6E6'; // progress-bar track
const PLACE     = 'DCE9F1'; // image placeholder
const PLACE_TXT = '8AA9BC';

const HEAD = 'Roboto';    // theme major font
const BODY = 'Open Sans'; // theme minor font

// ---------------------------------------------------------- shared strings

const ADD_NAME = 'Add Your Name';
const ADD_TITLE = 'Add Tittle Here';
const BODY_A = 'PLACEHOLDER';
const BODY_B = 'lorem ipsum dolor sitasana nisan atamet nisanen nibuhase';
const BODY_C =
  'PLACEHOLDER'
  + 'PLACEHOLDER';
const BODY_D = 'lorem ipsum dolor sita amet adipiscingesa nibuhase eneti';
const BODY_E =
  'PLACEHOLDER'
  + 'PLACEHOLDER'
  + 'hanasa nisan hunsana';
const BODY_F =
  'PLACEHOLDER'
  + 'hansanil nian hane';
const BODY_G = 'PLACEHOLDER';
const BODY_H =
  'PLACEHOLDER'
  + 'tincidunt velas aliquet ';
const BODY_I = 'lorem ipsum dolor anisi sitanis nibubase enen hanen banisula';
const BODY_J = 'PLACEHOLDER';
const BODY_K =
  'PLACEHOLDER'
  + 'PLACEHOLDER'
  + 'tinciduntees velas aliqueta nibhasi tincidunt velas hanasa';
const BODY_L = 'lorem ipsum dolor sitak amet nisanen elitas sed cursu misa';
const BODY_M = 'PLACEHOLDER';
const CAPTION_A = 'lorem ipsum dolor sita amet consec';
const CAPTION_B = 'lorem ipsum dolor sita amet consecutasli  ati nasat  ';
const HEADLINE = 'We can’t help everyone, but everyone can help someone.';
const HEADLINE_2L = 'We can’t help everyone, but everyone \ncan help someone.';
const HEADLINE_LC = 'we can’t help everyone, but everyone can help someone.';


// ------------------------------------------------------------------ helpers
// Every helper is a thin wrapper over one pptxgenjs call so the per-slide
// builders below read as a list of "what is on the slide".

/** Filled rectangle (colour panel / band). */
function band(s, x, y, w, h, color, o) {
  o = o || {};
  s.addShape('rect', {
    x: x, y: y, w: w, h: h,
    fill: { color: color, transparency: o.t || 0 },
    line: o.ln ? { color: o.ln, width: o.lw || 1 } : { type: 'none' },
    rotate: o.rot || 0,
  });
}

/** Filled ellipse. */
function circle(s, x, y, w, h, color, o) {
  o = o || {};
  s.addShape('ellipse', {
    x: x, y: y, w: w, h: h,
    fill: { color: color },
    line: o.ln ? { color: o.ln, width: o.lw || 2 } : { type: 'none' },
  });
}

/** Text box. o: sz, b, c(olor), f(ont), al(ign), v(align), ls(line spacing), cs(char spacing) */
function text(s, x, y, w, h, str, o) {
  o = o || {};
  const opt = {
    x: x, y: y, w: w, h: h,
    fontSize: o.sz || 12,
    fontFace: o.f || BODY,
    color: o.c || CHARCOAL,
    bold: !!o.b,
    align: o.al || 'left',
    valign: o.v || 'top',
    wrap: !o.nowrap,
    isTextBox: true,
  };
  if (o.ls) opt.lineSpacingMultiple = o.ls;
  if (o.cs) opt.charSpacing = o.cs;
  if (o.rot) opt.rotate = o.rot;
  if (o.m !== undefined) opt.margin = o.m;
  s.addText(str, opt);
}

/** Thin horizontal hairline. */
function rule(s, x, y, w, color) {
  s.addShape('line', { x: x, y: y, w: w, h: 0, line: { color: color, width: 1 } });
}

/** Three small dots used as a page ornament. */
function dots(s, x, y, color) {
  for (let i = 0; i < 3; i++) circle(s, x + i * 0.2715, y, 0.168, 0.168, color);
}

/** "Next page" caption plus its trailing hairline. */
function nextPage(s, tx, ty, lx, ly, lw, color) {
  text(s, tx, ty, 0.905, 0.326, 'Next page', { sz: 10, c: color, al: 'justify', ls: 1.5 });
  rule(s, lx, ly, lw, color === WHITE ? WHITE : SKY);
}

/**
 * Numbered feature card: oversized ghost numeral with a bold title (and an
 * optional paragraph) overlapping its lower-right corner.
 * `rot` rotates the whole card about its own bounding box (slide 4).
 */
function card(s, x, y, numeral, title, body, o) {
  o = o || {};
  const big = !o.small;
  const tdx = o.tdx !== undefined ? o.tdx : (big ? 0.316 : 0.35);
  const tdy = o.tdy !== undefined ? o.tdy : (big ? 0.326 : 0.234);
  const boxes = [
    { x: x, y: y, w: 1.29, h: big ? 0.909 : 0.774, str: numeral,
      opt: { sz: big ? 48 : 40, b: 1, c: o.nc || GHOST, f: HEAD, al: 'center' } },
    { x: x + tdx, y: y + tdy, w: o.tw || 1.29, h: 0.303, str: title,
      opt: { sz: 12, b: 1, c: o.tc || INK, f: HEAD, al: o.tal } },
  ];
  if (body) {
    const bdy = o.bdy !== undefined ? o.bdy : (big ? 0.825 : 0.779);
    boxes.push({ x: x + tdx, y: y + bdy, w: o.bw || 2.405, h: 0.531, str: body,
                 opt: { sz: 9, c: o.bc || GRAY, al: o.bal || 'justify', ls: 1.5 } });
  }
  (o.rot ? rotateBoxes(boxes, o.rot) : boxes).forEach(function (b) {
    text(s, b.x, b.y, b.w, b.h, b.str, b.opt);
  });
}

/** Rotate a set of boxes clockwise about the centre of their common bounds. */
function rotateBoxes(boxes, deg) {
  const x0 = Math.min.apply(null, boxes.map(function (b) { return b.x; }));
  const y0 = Math.min.apply(null, boxes.map(function (b) { return b.y; }));
  const x1 = Math.max.apply(null, boxes.map(function (b) { return b.x + b.w; }));
  const y1 = Math.max.apply(null, boxes.map(function (b) { return b.y + b.h; }));
  const cx = (x0 + x1) / 2, cy = (y0 + y1) / 2;
  const rad = (deg * Math.PI) / 180, cos = Math.cos(rad), sin = Math.sin(rad);
  return boxes.map(function (b) {
    const dx = b.x + b.w / 2 - cx, dy = b.y + b.h / 2 - cy;
    return {
      x: cx + dx * cos - dy * sin - b.w / 2,
      y: cy + dx * sin + dy * cos - b.h / 2,
      w: b.w, h: b.h, str: b.str,
      opt: Object.assign({}, b.opt, { rot: deg }),
    };
  });
}

/** Stand-in for a photo / device mock-up from the original deck. */
function photo(s, x, y, w, h) {
  s.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: PLACE }, line: { type: 'none' } });
  s.addText('[image]', {
    x: x, y: y, w: w, h: h, fontSize: 11, fontFace: BODY,
    color: PLACE_TXT, align: 'center', valign: 'middle',
  });
}

/** Progress bar: grey track, dark fill, caption on the left, percentage right. */
function meter(s, x, y, w, h, pct, percentText, label) {
  band(s, x, y + 0.012, w, 0.177, TRACK);
  band(s, x, y + 0.012, w * pct, 0.177, CHARCOAL);
  text(s, x + 0.039, y, 1.2, h, label, { sz: 6, c: WHITE, v: 'middle', nowrap: 1, m: 0 });
  text(s, x, y + 0.012, w * pct, 0.177, percentText,
       { sz: 6, c: WHITE, al: 'right', v: 'middle', m: [0, 7.2, 0, 7.2] });
}

/** Freeform shape described by a normalised [cmd, ...coords] path table. */
function freeform(s, x, y, w, h, color, cmds) {
  const pts = [];
  for (const c of cmds) {
    if (c[0] === 'M') pts.push({ x: c[1] * w, y: c[2] * h, moveTo: true });
    else if (c[0] === 'L') pts.push({ x: c[1] * w, y: c[2] * h });
    else if (c[0] === 'C') pts.push({ x: c[5] * w, y: c[6] * h,
      curve: { type: 'cubic', x1: c[1] * w, y1: c[2] * h, x2: c[3] * w, y2: c[4] * h } });
    else pts.push({ close: true });
  }
  s.addShape('custGeom', {
    x: x, y: y, w: w, h: h, points: pts,
    fill: { color: color }, line: { type: 'none' },
  });
}

/** Vertical stem ending in a round dot (slide 28 timeline). */
function stem(s, x, y, h, color) {
  s.addShape('line', { x: x, y: y, w: 0, h: h, line: { color: color, width: 1 } });
  circle(s, x - 0.07, y + h - 0.07, 0.14, 0.14, color);
}

/** Percent-stacked bar chart matching the deck's two chart objects. */
function stackedChart(s, x, y, w, h, dir, labels, series, colors) {
  const data = series.map(function (sr, i) {
    return { name: 'Series ' + (i + 1), labels: labels, values: sr };
  });
  s.addChart('bar', data, {
    x: x, y: y, w: w, h: h,
    barDir: dir, barGrouping: 'percentStacked', barGapWidthPct: 150,
    chartColors: colors, showLegend: false,
    valAxisLabelFormatCode: '0%',
    catAxisLabelFontSize: 8, valAxisLabelFontSize: 8,
    catAxisLabelFontFace: BODY, valAxisLabelFontFace: BODY,
    catAxisLabelColor: '757575', valAxisLabelColor: '757575',
    catAxisLineColor: 'D9D9D9', valAxisLineShow: false,
    catAxisMajorTickMark: 'none', valAxisMajorTickMark: 'none',
    valGridLine: { color: 'D9D9D9', size: 0.75 },
    catGridLine: { style: 'none' },
    chartArea: { fill: { color: WHITE } },
  });
}

// -------------------------------------------------- freeform outlines
// Normalised path tables: [cmd, ...coords] with coords in 0..1 of the shape box.
const BANNER_LIGHT = [
  ['M', 1, 0.0639],
  ['C', 0.9546, 0, 0.9546, 0, 0.9546, 0],
  ['C', 0.8716, 0.0255, 0.6887, 0.0748, 0.4838, 0.0748],
  ['C', 0.2827, 0.0748, 0.1206, 0.0274, 0.0454, 0.0018],
  ['C', 0, 0.0639, 0, 0.0639, 0, 0.0639],
  ['C', 0, 0.9088, 0, 0.9088, 0, 0.9088],
  ['C', 0, 0.9088, 0.2049, 1, 0.4838, 1],
  ['C', 0.7613, 1, 1, 0.9088, 1, 0.9088],
  ['C', 1, 0.0639, 1, 0.0639, 1, 0.0639],
  ['Z'],
];
const BANNER_LIGHT_LIP = [
  ['M', 0.9546, 0],
  ['C', 0.8716, 0.1647, 0.6887, 0.4824, 0.4838, 0.4824],
  ['C', 0.2827, 0.4824, 0.1206, 0.1765, 0.0454, 0.0118],
  ['C', 0, 0.4118, 0, 0.4118, 0, 0.4118],
  ['C', 0, 0.4118, 0.2049, 1, 0.4838, 1],
  ['C', 0.7613, 1, 1, 0.4118, 1, 0.4118],
  ['L', 0.9546, 0],
  ['Z'],
];
const BANNER_DARK = [
  ['M', 1, 0.1025],
  ['C', 1, 0.1025, 0.7891, 0, 0.5, 0],
  ['C', 0.2109, 0, 0, 0.1025, 0, 0.1025],
  ['C', 0, 0.9335, 0, 0.9335, 0, 0.9335],
  ['C', 0.0427, 1, 0.0427, 1, 0.0427, 1],
  ['C', 0.0427, 1, 0.2299, 0.9101, 0.4846, 0.9101],
  ['C', 0.7382, 0.9101, 0.9562, 1, 0.9573, 1],
  ['C', 1, 0.9335, 1, 0.9335, 1, 0.9335],
  ['L', 1, 0.1025],
  ['Z'],
];
const BANNER_DARK_LIP = [
  ['M', 1, 0.6129],
  ['C', 0.9573, 1, 0.9573, 1, 0.9573, 1],
  ['C', 0.9573, 1, 0.7393, 0.4731, 0.4846, 0.4731],
  ['C', 0.2299, 0.4731, 0.0427, 1, 0.0427, 1],
  ['C', 0, 0.6129, 0, 0.6129, 0, 0.6129],
  ['C', 0, 0.6129, 0.2109, 0, 0.5, 0],
  ['C', 0.7891, 0, 1, 0.6129, 1, 0.6129],
  ['Z'],
];
const BANNER_EDGE_L = [
  ['M', 1, 1],
  ['L', 0, 0.9259],
  ['L', 0, 0],
  ['L', 1, 0.0721],
  ['L', 1, 1],
  ['Z'],
];
const BANNER_EDGE_R = [
  ['M', 0, 1],
  ['L', 1, 0.9259],
  ['L', 1, 0],
  ['L', 0, 0.0721],
  ['L', 0, 1],
  ['Z'],
];
const PETAL_A = [
  ['M', 0.9542, 0.8693],
  ['C', 0.9083, 0.864, 0.8667, 0.841, 0.8417, 0.8057],
  ['C', 0.8125, 0.7633, 0.8167, 0.7085, 0.8521, 0.6714],
  ['C', 0.9646, 0.553, 0.9854, 0.3869, 0.8938, 0.2473],
  ['C', 0.7708, 0.0636, 0.4938, 0, 0.2792, 0.1078],
  ['C', 0.075, 0.2085, 0, 0.4293, 0.1083, 0.6078],
  ['C', 0.1979, 0.7544, 0.3833, 0.8304, 0.5646, 0.8092],
  ['C', 0.6208, 0.8021, 0.6771, 0.8269, 0.7063, 0.8675],
  ['C', 0.7562, 0.9417, 0.8438, 0.9912, 0.9417, 1],
  ['C', 1, 0.9435, 1, 0.9435, 1, 0.9435],
  ['L', 0.9542, 0.8693],
  ['Z'],
];
const PETAL_B = [
  ['M', 0.7372, 0.5802],
  ['C', 0.6138, 0.533, 0.4621, 0.5407, 0.3545, 0.5912],
  ['C', 0.3157, 0.6088, 0.2646, 0.6099, 0.224, 0.5956],
  ['C', 0.1993, 0.5868, 0.1993, 0.5868, 0.1993, 0.5868],
  ['C', 0.157, 0.5714, 0.1305, 0.544, 0.1305, 0.5132],
  ['C', 0.1305, 0.5099, 0.1323, 0.5077, 0.1323, 0.5044],
  ['C', 0.1376, 0.4791, 0.1605, 0.4582, 0.1958, 0.4451],
  ['C', 0.1975, 0.4451, 0.1975, 0.4451, 0.1975, 0.4451],
  ['C', 0.2399, 0.4297, 0.2928, 0.4319, 0.3316, 0.4505],
  ['C', 0.448, 0.5088, 0.6138, 0.5209, 0.7531, 0.4714],
  ['C', 0.9365, 0.4077, 1, 0.2615, 0.8942, 0.1484],
  ['C', 0.7919, 0.0407, 0.5714, 0, 0.3933, 0.0582],
  ['C', 0.2469, 0.1055, 0.1728, 0.2033, 0.1922, 0.2989],
  ['C', 0.1993, 0.3286, 0.1764, 0.3582, 0.134, 0.3725],
  ['C', 0.06, 0.3989, 0.0106, 0.4462, 0.0018, 0.4978],
  ['C', 0, 0.5033, 0, 0.5099, 0, 0.5154],
  ['C', 0, 0.5747, 0.0494, 0.6264, 0.1305, 0.656],
  ['C', 0.1605, 0.667, 0.1605, 0.667, 0.1605, 0.667],
  ['C', 0.2011, 0.6813, 0.224, 0.7099, 0.2205, 0.7385],
  ['C', 0.2063, 0.8198, 0.2681, 0.9022, 0.3915, 0.9462],
  ['C', 0.5432, 1, 0.7407, 0.9714, 0.8395, 0.8813],
  ['C', 0.9489, 0.7802, 0.9012, 0.644, 0.7372, 0.5802],
  ['Z'],
];
const PETAL_C = [
  ['M', 0.9429, 0.3944],
  ['C', 0.8956, 0.2465, 0.7967, 0.1725, 0.7011, 0.1937],
  ['C', 0.6714, 0.2007, 0.6429, 0.1761, 0.6275, 0.1338],
  ['C', 0.6011, 0.0616, 0.5549, 0.0123, 0.5033, 0.0018],
  ['C', 0.4967, 0.0018, 0.4912, 0, 0.4846, 0],
  ['C', 0.4264, 0, 0.3736, 0.0493, 0.344, 0.1303],
  ['C', 0.333, 0.1602, 0.333, 0.1602, 0.333, 0.1602],
  ['C', 0.3187, 0.2007, 0.2912, 0.2254, 0.2615, 0.2201],
  ['C', 0.1802, 0.206, 0.0978, 0.2676, 0.0549, 0.3908],
  ['C', 0, 0.544, 0.0286, 0.7394, 0.1198, 0.838],
  ['C', 0.2198, 0.9489, 0.356, 0.9014, 0.4198, 0.7359],
  ['C', 0.467, 0.6144, 0.4593, 0.463, 0.4088, 0.3539],
  ['C', 0.3912, 0.3169, 0.3901, 0.2658, 0.4044, 0.2236],
  ['C', 0.4143, 0.1989, 0.4143, 0.1989, 0.4143, 0.1989],
  ['C', 0.4286, 0.1567, 0.456, 0.1303, 0.4868, 0.132],
  ['C', 0.4901, 0.132, 0.4934, 0.132, 0.4967, 0.132],
  ['C', 0.5209, 0.1373, 0.5429, 0.1602, 0.5549, 0.1972],
  ['C', 0.556, 0.1972, 0.556, 0.1972, 0.556, 0.1972],
  ['C', 0.5703, 0.2394, 0.5692, 0.2923, 0.5495, 0.331],
  ['C', 0.4912, 0.4472, 0.4791, 0.6144, 0.5286, 0.7518],
  ['C', 0.5923, 0.9366, 0.7385, 1, 0.8527, 0.8926],
  ['C', 0.9604, 0.7923, 1, 0.5704, 0.9429, 0.3944],
  ['Z'],
];
const PETAL_D = [
  ['M', 0.8697, 0.3432],
  ['C', 0.8398, 0.3333, 0.8398, 0.3333, 0.8398, 0.3333],
  ['C', 0.7993, 0.319, 0.7746, 0.2904, 0.7799, 0.2618],
  ['C', 0.794, 0.1804, 0.7324, 0.0968, 0.6092, 0.0539],
  ['C', 0.456, 0, 0.2606, 0.0286, 0.162, 0.1188],
  ['C', 0.0511, 0.22, 0.0986, 0.3564, 0.2641, 0.4191],
  ['C', 0.3856, 0.4664, 0.537, 0.4598, 0.6461, 0.4092],
  ['C', 0.6831, 0.3916, 0.7342, 0.3894, 0.7764, 0.4048],
  ['C', 0.8011, 0.4136, 0.8011, 0.4136, 0.8011, 0.4136],
  ['C', 0.8415, 0.429, 0.8697, 0.4565, 0.868, 0.4862],
  ['C', 0.868, 0.4895, 0.868, 0.4928, 0.868, 0.4961],
  ['C', 0.8627, 0.5204, 0.8398, 0.5424, 0.8028, 0.5556],
  ['C', 0.7606, 0.571, 0.7077, 0.5688, 0.669, 0.5501],
  ['C', 0.5528, 0.4917, 0.3856, 0.4796, 0.2482, 0.5281],
  ['C', 0.0634, 0.593, 0, 0.7393, 0.1074, 0.8526],
  ['C', 0.2077, 0.9604, 0.4296, 1, 0.6056, 0.9428],
  ['C', 0.7535, 0.8955, 0.8275, 0.7976, 0.8063, 0.7019],
  ['C', 0.7993, 0.6722, 0.8239, 0.6425, 0.8662, 0.6271],
  ['C', 0.9384, 0.6007, 0.9877, 0.5545, 0.9982, 0.5028],
  ['C', 0.9982, 0.4972, 1, 0.4906, 1, 0.4851],
  ['C', 1, 0.4257, 0.9507, 0.3729, 0.8697, 0.3432],
  ['Z'],
];
const PETAL_E = [
  ['M', 0.1663, 0.223],
  ['C', 0.0722, 0.3532, 0.0853, 0.513, 0.186, 0.6264],
  ['C', 0.221, 0.6673, 0.2254, 0.7212, 0.1947, 0.7639],
  ['C', 0.1772, 0.79, 0.1772, 0.79, 0.1772, 0.79],
  ['C', 0.1466, 0.8346, 0.0919, 0.8643, 0.0328, 0.8625],
  ['C', 0.0263, 0.8625, 0.0197, 0.8625, 0.0131, 0.8606],
  ['C', 0, 0.9981, 0, 0.9981, 0, 0.9981],
  ['C', 0.0109, 1, 0.0241, 1, 0.035, 1],
  ['C', 0.1532, 1, 0.2582, 0.9498, 0.3173, 0.8625],
  ['C', 0.337, 0.8327, 0.337, 0.8327, 0.337, 0.8327],
  ['C', 0.3676, 0.79, 0.4223, 0.7639, 0.4792, 0.7695],
  ['C', 0.6411, 0.7844, 0.8074, 0.7175, 0.8928, 0.5892],
  ['C', 1, 0.4275, 0.9431, 0.2212, 0.7637, 0.1171],
  ['C', 0.5646, 0, 0.291, 0.0502, 0.1663, 0.223],
  ['Z'],
];

// Slide 28 - the four banner blocks (odd ones ride lower, even ones are dark).
const BANNERS = [
  { x: 1.774, y: 3.34,  w: 2.179, h: 1.55,  lipH: 0.24,  sy: 4.861, sh: 0.257, dark: false },
  { x: 4.171, y: 3.154, w: 2.387, h: 1.573, lipH: 0.263, sy: 4.559, sh: 0.56,  dark: true  },
  { x: 6.775, y: 3.34,  w: 2.179, h: 1.55,  lipH: 0.24,  sy: 4.861, sh: 0.257, dark: false },
  { x: 9.172, y: 3.154, w: 2.387, h: 1.573, lipH: 0.263, sy: 4.559, sh: 0.56,  dark: true  },
];

// ---------------------------------------------------------- slide builders

function slide01(s) {  // title cover: brand block, oversized wordmark, ornament dots
  band(s, 2.984, 2.783, 7.365, 2.057, BLUE, { t: 15 });
  band(s, 6.343, 2.788, 0.647, 4.104, SKY, { rot: 90 });
  text(s, 4.81, 3.078, 3.713, 1.212, 'Charl', { sz: 66, b: 1, c: WHITE, f: HEAD, al: 'center', cs: 3 });
  text(s, 5.171, 4.699, 2.991, 0.281, 'Presentation Template', { sz: 8, c: BLUE, al: 'center', ls: 1.5, cs: 3 });
  dots(s, 0.662, 6.675, WHITE);
  nextPage(s, 9.602, 0.548, 10.66, 0.752, 2.674, WHITE);
}

function slide02(s) {  // quote page with left pale panel and numbered badge
  band(s, 0, 0, 3.688, 7.5, SKY);
  text(s, 8.533, 3.668, 3.718, 1.084, BODY_K, { sz: 10, c: GRAY, al: 'justify', ls: 1.5 });
  dots(s, 11.985, 0.659, SKY);
  band(s, 1.281, 2.006, 2.406, 0.444, WHITE);
  nextPage(s, 8.533, 5.904, 9.823, 6.108, 2.337, BLUE_LINK);
  text(s, 8.537, 2.222, 3.713, 1.111, HEADLINE, { sz: 20, b: 1, c: INK, f: HEAD });
  text(s, 0.669, 5.425, 1.444, 1.089, CAPTION_B, { sz: 8, c: BLUE, ls: 1.5, cs: 3 });
  card(s, 0.398, 0.826, '01', ADD_TITLE, null, { nc: SKY_DEEP, tc: WHITE, tw: 1.548 });
}

function slide03(s) {  // wide image band with right pale rail
  band(s, 10.625, 0, 2.708, 7.5, SKY);
  text(s, 11.235, 1.293, 1.527, 2.457, HEADLINE_LC, { sz: 20, b: 1, c: BLUE_LINK, f: HEAD });
  band(s, 10.625, 6.314, 2.076, 0.444, WHITE);
  text(s, 5.886, 5.385, 3.718, 1.336, 'lorem ipsum dolor sita amet conseca sanialesi teturasi adipiscingesa nibuhase eget justo sedasanis cursus mi cursus misani utasisau nasat sapien tinciduntees nisal velasonseca sanialesi teturasi adipiscingesa nibuhasen  aliqueta nibhasi tincidunt velas hanasa', { sz: 10, c: GRAY, al: 'justify', ls: 1.5 });
  text(s, 1.224, 5.385, 3.904, 0.774, HEADLINE, { sz: 20, b: 1, c: INK, f: HEAD });
  nextPage(s, 7.375, 0.347, 8.581, 0.551, 2.044, BLUE_LINK);
  dots(s, 0.662, 6.674, SKY);
  text(s, 11.235, 5.075, 1.341, 0.685, CAPTION_A, { sz: 8, c: GRAY, ls: 1.5, cs: 3 });
}

function slide04(s) {  // portrait image with rotated side badge
  dots(s, 0.662, 0.659, SKY);
  text(s, 2.872, 5.512, 1.986, 0.831, 'PLACEHOLDER', { sz: 10, c: WHITE, al: 'justify', ls: 1.5 });
  text(s, 9.252, 1.942, 1.863, 2.121, HEADLINE, { sz: 20, b: 1, c: INK, f: HEAD });
  band(s, 7.202, 6.087, 3.687, 0.684, SKY, { rot: 180 });
  band(s, 2.313, 1.54, 1.368, 0.444, WHITE);
  text(s, 9.252, 4.533, 3.086, 0.831, 'PLACEHOLDER', { sz: 10, c: GRAY, al: 'justify', ls: 1.5 });
  text(s, 9.252, 1.061, 3.269, 0.281, 'lorem ipsum dolor amet consen', { sz: 8, c: GRAY, al: 'justify', ls: 1.5, cs: 3 });
  card(s, 0.263, 5.246, '02', ADD_TITLE, null, { tw: 1.53, tdy: 0.298, rot: 270 });
}

function slide05(s) {  // left rail, quote block and paragraph on the right
  band(s, 0, 0, 2.842, 7.5, SKY);
  text(s, 8.885, 4.623, 3.086, 1.084, 'PLACEHOLDER', { sz: 10, c: GRAY, al: 'justify', ls: 1.5 });
  text(s, 8.89, 3.132, 3.422, 1.111, HEADLINE_2L, { sz: 20, b: 1, c: INK, f: HEAD });
  nextPage(s, 8.885, 6.254, 10.176, 6.479, 3.158, BLUE_LINK);
  band(s, 1.15, 3.465, 1.692, 0.444, WHITE);
  text(s, 3.964, 1.164, 3.718, 0.831, BODY_C, { sz: 10, c: GRAY, al: 'justify', ls: 1.5 });
  dots(s, 11.985, 0.659, SKY);
  text(s, 0.725, 1.008, 1.341, 1.089, CAPTION_B, { sz: 8, c: BLUE, al: 'right', ls: 1.5, cs: 3 });
}

function slide06(s) {  // mirrored layout of slide 5 with a right rail
  band(s, 9.008, 0, 4.325, 7.5, SKY);
  band(s, 6.64, 5.47, 2.714, 0.444, WHITE);
  text(s, 1.239, 4.154, 3.086, 1.336, 'PLACEHOLDER', { sz: 10, c: GRAY, al: 'justify', ls: 1.5 });
  text(s, 1.243, 2.663, 3.422, 1.111, HEADLINE_2L, { sz: 20, b: 1, c: INK, f: HEAD });
  nextPage(s, 1.228, 6.169, 2.519, 6.373, 2.823, BLUE_LINK);
  dots(s, 0.662, 0.659, SKY);
  text(s, 10.963, 1.327, 1.701, 1.336, CAPTION_B, { sz: 10, c: BLUE, ls: 1.5, cs: 3 });
  card(s, 10.672, 5.586, '03', ADD_TITLE, null, { nc: SKY_DEEP, tc: WHITE, tw: 1.548 });
}

function slide07(s) {  // three numbered services beside a centre pale column
  band(s, 5.376, 0.75, 2.097, 6.75, SKY);
  dots(s, 0.662, 0.638, SKY);
  card(s, 0.934, 1.438, '01', 'The First Service', BODY_F, { tw: 2.708, bw: 3.014, bdy: 0.889 });
  card(s, 0.934, 3.314, '02', 'The Second Service', BODY_F, { tw: 2.23, bw: 3.014, bdy: 0.889 });
  card(s, 0.934, 5.132, '03', 'The Third Service', BODY_F, { tw: 2.097, bw: 3.014, bdy: 0.889 });
  nextPage(s, 8.338, 0.642, 9.387, 0.85, 3.097, BLUE_LINK);
  text(s, 8.338, 3.113, 3.718, 0.831, BODY_C, { sz: 10, c: GRAY, al: 'justify', ls: 1.5 });
  text(s, 8.342, 1.63, 3.713, 1.111, HEADLINE, { sz: 20, b: 1, c: INK, f: HEAD });
}

function slide08(s) {  // image left, quote and two service cards right
  band(s, 4.577, 0.887, 2.026, 6.613, SKY);
  dots(s, 11.985, 0.659, SKY);
  nextPage(s, 7.749, 4.409, 9.039, 4.617, 3.657, BLUE_LINK);
  text(s, 7.744, 2.917, 3.718, 0.831, BODY_C, { sz: 10, c: GRAY, al: 'justify', ls: 1.5 });
  text(s, 7.749, 1.471, 3.713, 1.111, HEADLINE, { sz: 20, b: 1, c: INK, f: HEAD });
  card(s, 7.509, 5.416, '01', 'First Service', BODY_D, { tw: 1.29, bw: 1.867 });
  card(s, 9.975, 5.416, '02', 'Secpnd Service', BODY_D, { tw: 1.58, bw: 1.867 });
  text(s, 1.242, 6.087, 1.341, 0.685, CAPTION_A, { sz: 8, c: GRAY, ls: 1.5, cs: 3 });
}

function slide09(s) {  // wide image with two service cards under it
  band(s, 10.625, 0, 2.708, 7.5, SKY);
  card(s, 3.871, 5.482, '01', 'First Service', BODY_A, { tw: 1.29, bw: 2.405 });
  card(s, 7.002, 5.482, '02', 'Second Service', BODY_A, { tw: 1.58, bw: 2.405 });
  band(s, 10.625, 6.332, 1.094, 0.444, WHITE);
  text(s, 0.934, 2.308, 1.619, 2.457, HEADLINE, { sz: 20, b: 1, c: INK, f: HEAD });
  dots(s, 0.662, 6.674, SKY);
  text(s, 4.187, 0.724, 5.535, 0.579, 'PLACEHOLDER', { sz: 10, c: GRAY, ls: 1.5 });
  nextPage(s, 0.579, 0.616, 1.604, 0.841, 2.042, BLUE_LINK);
}

function slide10(s) {  // left rail, two service cards, image right
  band(s, 0, 0, 2.708, 7.5, SKY);
  card(s, 3.357, 0.879, '01', 'First Service', BODY_A, { tw: 1.29, bw: 2.405 });
  card(s, 3.357, 2.796, '02', 'Second Service', BODY_A, { tw: 2.23, bw: 2.405 });
  nextPage(s, 3.674, 4.884, 4.964, 5.092, 3.001, BLUE_LINK);
  text(s, 0.664, 1.007, 1.527, 2.457, HEADLINE_LC, { sz: 20, b: 1, c: BLUE_LINK, f: HEAD });
  band(s, 0.664, 6.326, 2.045, 0.444, WHITE);
  dots(s, 11.985, 6.678, SKY);
  text(s, 3.674, 6.026, 4.291, 0.831, 'PLACEHOLDER', { sz: 10, c: GRAY, al: 'justify', ls: 1.5 });
  text(s, 8.577, 6.011, 2.534, 0.831, 'PLACEHOLDER', { sz: 10, c: GRAY, al: 'justify', ls: 1.5 });
  text(s, 0.664, 4.75, 1.341, 0.685, CAPTION_A, { sz: 8, c: GRAY, ls: 1.5, cs: 3 });
}

function slide11(s) {  // three service cards, centre image column, right rail
  band(s, 6.667, 0.729, 3.25, 6.771, SKY);
  card(s, 0.903, 1.417, '01', 'First Service', BODY_A, { tw: 1.29, bw: 2.405 });
  dots(s, 0.662, 0.659, SKY);
  card(s, 0.903, 3.293, '02', 'Second Service', BODY_A, { tw: 2.23, bw: 2.405 });
  card(s, 0.903, 5.111, '03', 'Third Service', BODY_A, { tw: 1.29, bw: 2.405 });
  nextPage(s, 10.797, 6.499, 11.992, 6.707, 1.342, BLUE_LINK);
  text(s, 10.797, 2.057, 1.61, 2.457, HEADLINE, { sz: 20, b: 1, c: INK, f: HEAD });
  text(s, 10.807, 4.992, 1.714, 1.084, 'PLACEHOLDER', { sz: 10, c: GRAY, al: 'justify', ls: 1.5 });
  text(s, 10.797, 1.299, 1.995, 0.236, 'lorem ipsum at dolor', { sz: 6, c: GRAY, al: 'justify', ls: 1.5, cs: 3 });
}

function slide12(s) {  // three white-on-blue portfolio cards over a blue column
  band(s, 4.562, 0, 3.458, 7.5, SKY);
  text(s, 0.942, 1.449, 2.148, 2.524, HEADLINE, { sz: 24, b: 1, c: INK, f: HEAD });
  card(s, 4.911, 0.967, '01', 'First Portfolio', BODY_G, { nc: SKY_DEEP, tc: WHITE, tw: 1.765, bw: 2.28, bc: WHITE });
  card(s, 4.911, 2.987, '02', 'Second Portfolio', BODY_G, { nc: SKY_DEEP, tc: WHITE, tw: 1.548, bw: 2.28, bc: WHITE });
  card(s, 4.911, 4.981, '03', 'Third Portfolio', BODY_G, { nc: SKY_DEEP, tc: WHITE, tw: 1.765, bw: 2.28, bc: WHITE });
  dots(s, 0.662, 6.69, SKY);
  text(s, 0.89, 4.632, 2.615, 1.336, 'PLACEHOLDER', { sz: 10, c: GRAY, al: 'justify', ls: 1.5 });
  text(s, 0.89, 0.636, 2.876, 0.281, 'lorem ipsum dolor sit amet', { sz: 8, c: GRAY, al: 'justify', ls: 1.5, cs: 3 });
}

function slide13(s) {  // three images across the top, two portfolio cards below
  band(s, 10.625, 0, 2.708, 7.5, SKY);
  text(s, 11.266, 1.355, 1.527, 2.457, HEADLINE_LC, { sz: 20, b: 1, c: BLUE_LINK, f: HEAD });
  band(s, 10.625, 6.313, 2.076, 0.444, WHITE);
  text(s, 0.632, 5.856, 1.789, 0.831, 'lorem ipsum dolor sitasi amet conseca sanialesiti uttinedunsi', { sz: 10, c: GRAY, al: 'justify', ls: 1.5 });
  dots(s, 0.735, 5.217, SKY);
  card(s, 3.292, 5.33, '01', 'First Portfolio', BODY_A, { tw: 1.29, bw: 2.405 });
  card(s, 6.667, 5.33, '02', 'Second Portfolio', BODY_A, { tw: 1.995, bw: 2.405 });
  text(s, 11.291, 4.981, 1.341, 0.685, CAPTION_A, { sz: 8, c: GRAY, ls: 1.5, cs: 3 });
}

function slide14(s) {  // tall image left, two portfolio cards, quote right
  band(s, 0, 0, 2.81, 7.5, SKY);
  band(s, 0.881, 5.726, 1.929, 0.444, WHITE);
  dots(s, 11.985, 0.659, SKY);
  text(s, 9.105, 4.866, 3.826, 1.212, HEADLINE, { sz: 22, b: 1, c: INK, f: HEAD });
  card(s, 5.26, 0.993, '01', 'First Portfolio', BODY_A, { tw: 1.29, bw: 2.405 });
  card(s, 5.26, 2.838, '02', 'Second Portfolio', BODY_A, { tw: 2.23, bw: 2.405 });
  nextPage(s, 9.105, 6.396, 10.272, 6.618, 3.061, BLUE_LINK);
  text(s, 9.011, 1.235, 3.591, 0.236, 'lorem ipsum dolor nita amet nasail', { sz: 6, c: GRAY, al: 'center', ls: 1.5, cs: 3 });
}

function slide15(s) {  // three images with blue caption bars
  band(s, 5.385, 0.729, 2.646, 6.771, SKY);
  text(s, 0.849, 1.978, 1.785, 2.928, HEADLINE, { sz: 24, b: 1, c: INK, f: HEAD });
  dots(s, 11.985, 0.659, SKY);
  nextPage(s, 0.891, 6.394, 2.058, 6.616, 1.645, BLUE_LINK);
  text(s, 0.849, 0.569, 3.891, 0.281, CAPTION_A, { sz: 8, c: GRAY, al: 'justify', ls: 1.5, cs: 3 });
  band(s, 8.975, 3.321, 3.365, 0.538, BLUE, { t: 15 });
  band(s, 8.975, 6.079, 3.365, 0.538, BLUE, { t: 15 });
  text(s, 9.757, 3.438, 1.8, 0.303, '2. Second Portfolio', { sz: 12, b: 1, c: WHITE, f: HEAD, al: 'center' });
  text(s, 9.757, 6.196, 1.8, 0.303, '3. Third Portfolio', { sz: 12, b: 1, c: WHITE, f: HEAD, al: 'center' });
  band(s, 3.703, 6.081, 3.365, 0.538, BLUE, { t: 15 });
  text(s, 4.485, 6.199, 1.8, 0.303, '1. First Portfolio', { sz: 12, b: 1, c: WHITE, f: HEAD, al: 'center' });
}

function slide16(s) {  // three images plus a highlighted portfolio card
  band(s, 0.964, 5.034, 3.771, 1.746, SKY);
  text(s, 8.473, 4.822, 4.031, 1.313, HEADLINE, { sz: 24, b: 1, c: INK, f: HEAD });
  nextPage(s, 8.473, 6.498, 9.64, 6.72, 2.864, BLUE_LINK);
  dots(s, 11.985, 0.659, SKY);
  text(s, 8.513, 2.556, 3.891, 1.336, 'PLACEHOLDER', { sz: 10, c: GRAY, al: 'justify', ls: 1.5 });
  card(s, 1.253, 5.238, '01', 'First Portfolio', 'PLACEHOLDER', { small: 1, nc: SKY_DEEP, tc: WHITE, tw: 1.662, bw: 2.51, bdy: 0.732, bc: WHITE });
  band(s, 6.267, 1.758, 1.289, 0.376, WHITE);
  text(s, 8.513, 1.805, 3.912, 0.281, 'lorem ipsum dolor nita amet iconsebanu', { sz: 8, c: GRAY, al: 'justify', ls: 1.5, cs: 3 });
  band(s, 4.952, 6.202, 2.604, 0.538, BLUE, { t: 15 });
  text(s, 5.381, 6.319, 1.772, 0.303, '2. Second Portfolio', { sz: 12, b: 1, c: WHITE, f: HEAD, al: 'center' });
}

function slide17(s) {  // team page with two progress meters
  band(s, 0, 0, 2.708, 7.5, SKY);
  text(s, 8.213, 3.138, 3.718, 0.831, BODY_C, { sz: 10, c: GRAY, al: 'justify', ls: 1.5 });
  dots(s, 11.985, 0.659, SKY);
  nextPage(s, 8.224, 4.549, 9.515, 4.753, 2.91, BLUE_LINK);
  text(s, 8.218, 1.629, 3.713, 1.111, HEADLINE, { sz: 20, b: 1, c: INK, f: HEAD });
  meter(s, 8.35, 6.353, 3.517, 0.202, 0.887, '82%', 'Lorem ipsum dolor');
  meter(s, 8.35, 5.718, 3.517, 0.202, 0.887, '82%', 'Lorem ipsum dolor');
  band(s, 0.892, 1.749, 1.817, 0.444, WHITE);
  text(s, 0.463, 5.545, 1.337, 0.909, '- Our \nTeam', { sz: 24, b: 1, c: BLUE, f: HEAD, al: 'right' });
}

function slide18(s) {  // three "Add Your Name" cards over three images
  band(s, 10.625, 0, 2.708, 7.5, SKY);
  text(s, 11.266, 1.189, 1.527, 2.457, HEADLINE, { sz: 20, b: 1, c: BLUE_LINK, f: HEAD });
  band(s, 10.625, 6.207, 2.076, 0.444, WHITE);
  card(s, 0.934, 1.372, '01', ADD_NAME, BODY_H, { tw: 1.8, bw: 2.405, bdy: 1.03 });
  card(s, 3.852, 1.372, '03', ADD_NAME, BODY_H, { tw: 1.776, bw: 2.405, bdy: 1.03 });
  card(s, 6.771, 1.372, '02', ADD_NAME, BODY_H, { tw: 2.089, bw: 2.405, bdy: 1.03 });
  dots(s, 0.662, 0.659, SKY);
  text(s, 11.291, 4.888, 1.341, 0.685, CAPTION_A, { sz: 8, c: GRAY, ls: 1.5, cs: 3 });
}

function slide19(s) {  // four square photos with blue caption bars
  band(s, 0, 4.416, 3.548, 3.084, SKY);
  band(s, 3.803, 0, 3.642, 4.14, SKY);
  dots(s, 0.662, 0.659, SKY);
  text(s, 8.402, 1.806, 4.031, 1.313, HEADLINE, { sz: 24, b: 1, c: INK, f: HEAD });
  nextPage(s, 8.402, 3.814, 9.569, 4.036, 2.723, BLUE_LINK);
  text(s, 8.401, 5.081, 3.891, 1.589, 'PLACEHOLDER', { sz: 10, c: GRAY, al: 'justify', ls: 1.5 });
  band(s, 6.095, 0.743, 1.349, 0.444, WHITE);
  text(s, 8.472, 0.829, 3.891, 0.281, CAPTION_A, { sz: 8, c: GRAY, al: 'justify', ls: 1.5, cs: 3 });
  band(s, 3.803, 6.289, 2.292, 0.419, BLUE, { t: 15 });
  text(s, 4.112, 6.356, 1.799, 0.286, '2. Add Your Name', { sz: 10.5, b: 1, c: WHITE, f: HEAD, al: 'center' });
  band(s, 1.256, 3.721, 2.292, 0.419, BLUE, { t: 15 });
  text(s, 1.565, 3.788, 1.799, 0.278, '1. Add Your Name', { sz: 10.5, b: 1, c: WHITE, f: HEAD, al: 'center' });
}

function slide20(s) {  // three small service cards beside a tall image
  band(s, 3.512, 0, 2.894, 6.841, SKY);
  text(s, 9.428, 1.767, 2.281, 2.524, HEADLINE, { sz: 24, b: 1, c: INK, f: HEAD });
  nextPage(s, 9.449, 6.515, 10.616, 6.737, 2.723, BLUE_LINK);
  text(s, 9.443, 4.559, 2.846, 1.336, 'PLACEHOLDER', { sz: 10, c: GRAY, al: 'justify', ls: 1.5 });
  card(s, 0.561, 1.563, '01', 'First Service', BODY_I, { small: 1, tw: 1.29, bw: 1.993 });
  card(s, 0.564, 3.28, '02', 'Second Service', BODY_I, { small: 1, tw: 1.457, bw: 1.991 });
  card(s, 0.561, 4.946, '03', 'Third Service', BODY_I, { small: 1, tw: 1.29, bw: 1.991 });
  text(s, 9.443, 0.81, 3.891, 0.281, 'lorem ipsum dolor sita amet', { sz: 8, c: GRAY, al: 'justify', ls: 1.5, cs: 3 });
  dots(s, 0.693, 0.659, SKY);
}

function slide21(s) {  // left image column with caption bar, two name cards
  band(s, 0, 0, 3.494, 7.5, SKY);
  band(s, 0, 0.618, 2.708, 0.444, WHITE);
  text(s, 4.361, 1.96, 1.561, 2.457, HEADLINE, { sz: 20, b: 1, c: INK, f: HEAD });
  text(s, 4.361, 5.167, 1.561, 1.589, 'PLACEHOLDER', { sz: 10, c: GRAY, al: 'justify', ls: 1.5 });
  card(s, 7.094, 5.444, '02', ADD_NAME, BODY_L, { small: 1, tw: 1.57, tdy: 0.222, bw: 1.918, bdy: 0.767 });
  card(s, 10.045, 5.444, '03', ADD_NAME, BODY_L, { small: 1, tw: 1.669, tdy: 0.222, bw: 1.918, bdy: 0.767 });
  dots(s, 11.985, 0.659, SKY);
  text(s, 4.372, 0.602, 4.109, 0.281, 'lorem ipsum dolor sita amet conseenasil', { sz: 8, c: GRAY, al: 'justify', ls: 1.5, cs: 3 });
  band(s, 0, 6.218, 2.708, 0.524, BLUE, { t: 15 });
  text(s, 0.454, 6.337, 1.799, 0.278, '1. Add Your Name', { sz: 10.5, b: 1, c: WHITE, f: HEAD, al: 'center' });
}

function slide22(s) {  // laptop mock-up, centre pale column, three service cards
  photo(s, 0.973, 3.75, 4.57, 2.662);
  band(s, 6.416, 0.75, 1.952, 6.75, SKY);
  card(s, 9.162, 1.438, '01', 'First Service', BODY_A, { tw: 1.29, bw: 2.405 });
  dots(s, 11.88, 0.659, SKY);
  card(s, 9.162, 3.314, '02', 'Second Service', BODY_A, { tw: 2.23, bw: 2.405 });
  card(s, 9.162, 5.132, '03', 'Third Service', BODY_A, { tw: 1.29, bw: 2.405 });
  nextPage(s, 1.301, 0.605, 2.591, 0.798, 2.952, BLUE_LINK);
  text(s, 1.306, 1.541, 4.013, 1.313, HEADLINE, { sz: 24, b: 1, c: INK, f: HEAD });
}

function slide23(s) {  // tablet mock-up on a pale panel with a numbered badge
  band(s, 0, 0, 4.688, 7.5, SKY);
  photo(s, 3.546, 1.157, 3.683, 5.455);
  band(s, 1.677, 2.187, 3.006, 0.652, WHITE);
  text(s, 8.356, 3.668, 3.718, 1.084, BODY_K, { sz: 10, c: GRAY, al: 'justify', ls: 1.5 });
  dots(s, 11.985, 0.659, SKY);
  nextPage(s, 8.356, 5.904, 9.646, 6.108, 2.337, BLUE_LINK);
  text(s, 8.36, 2.222, 3.452, 1.111, HEADLINE, { sz: 20, b: 1, c: INK, f: HEAD });
  card(s, 0.657, 5.613, '01', ADD_TITLE, null, { nc: SKY_DEEP, tc: WHITE, tw: 1.548 });
  text(s, 1.58, 3.385, 1.444, 1.089, CAPTION_B, { sz: 8, c: BLUE, al: 'right', ls: 1.5, cs: 3 });
}

function slide24(s) {  // monitor mock-up on the right, two service cards left
  band(s, 6.763, 0.887, 2.026, 6.613, SKY);
  photo(s, 7.771, 0.89, 4.799, 4.088);
  dots(s, 0.671, 0.659, SKY);
  nextPage(s, 4.713, 4.409, 0.671, 4.617, 3.657, BLUE_LINK);
  text(s, 1.258, 2.948, 4.333, 0.831, BODY_E, { sz: 10, c: GRAY, ls: 1.5 });
  text(s, 1.258, 1.471, 3.713, 1.111, HEADLINE, { sz: 20, b: 1, c: INK, f: HEAD });
  card(s, 0.943, 5.416, '01', 'First Service', BODY_D, { tw: 1.29, bw: 1.867 });
  card(s, 3.408, 5.416, '02', 'Secpnd Service', BODY_D, { tw: 1.58, bw: 1.867 });
  text(s, 10.783, 6.087, 1.341, 0.685, CAPTION_A, { sz: 8, c: GRAY, ls: 1.5, cs: 3 });
}

function slide25(s) {  // phone mock-up in the centre, three service cards right
  band(s, 5.524, 0.729, 2.646, 6.771, SKY);
  photo(s, 4.518, 1.358, 2.629, 5.267);
  text(s, 0.905, 1.978, 1.785, 2.928, HEADLINE, { sz: 24, b: 1, c: INK, f: HEAD });
  dots(s, 11.985, 0.659, SKY);
  nextPage(s, 0.947, 6.394, 2.114, 6.616, 1.645, BLUE_LINK);
  text(s, 0.905, 0.569, 3.891, 0.281, CAPTION_A, { sz: 8, c: GRAY, al: 'justify', ls: 1.5, cs: 3 });
  card(s, 9.218, 1.438, '01', 'First Service', BODY_A, { tw: 1.29, bw: 2.405 });
  card(s, 9.218, 3.314, '02', 'Second Service', BODY_A, { tw: 2.23, bw: 2.405 });
  card(s, 9.218, 5.132, '03', 'Third Service', BODY_A, { tw: 1.29, bw: 2.405 });
}

function slide26(s) {  // percent-stacked column chart with two service cards
  stackedChart(s, 0.746, 0.787, 5.397, 6.245, 'col',
    ['1', '2', '3', '4'],
    [[4.3, 2.5, 3.5, 4.5], [2.4, 4.4, 1.8, 2.8], [2, 2, 3, 5]],
    [BLUE, SKY, AZURE]);
  dots(s, 11.985, 0.659, SKY);
  nextPage(s, 11.089, 4.195, 7.046, 4.404, 3.657, BLUE_LINK);
  text(s, 7.046, 2.919, 4.845, 0.831, 'PLACEHOLDER', { sz: 10, c: GRAY, al: 'justify', ls: 1.5 });
  text(s, 7.046, 1.442, 3.713, 1.111, HEADLINE, { sz: 20, b: 1, c: INK, f: HEAD });
  card(s, 6.74, 5.243, '01', 'First Service', BODY_M, { tw: 1.29, bw: 2.152 });
  card(s, 9.447, 5.243, '02', 'Secpnd Service', BODY_M, { tw: 1.58, bw: 2.152 });
}

function slide27(s) {  // percent-stacked bar chart with three service cards
  stackedChart(s, 6.562, 1.109, 5.993, 3.808, 'bar',
    ['1', '2', '3'],
    [[4.3, 2.5, 3.5], [2.4, 4.4, 1.8], [2, 2, 3]],
    [AZURE, SKY, BLUE]);
  dots(s, 0.671, 0.659, SKY);
  nextPage(s, 4.836, 4.545, 0.671, 4.753, 3.78, BLUE_LINK);
  text(s, 1.381, 3.096, 4.333, 0.831, BODY_E, { sz: 10, c: GRAY, ls: 1.5 });
  text(s, 1.381, 1.512, 3.713, 1.111, HEADLINE, { sz: 20, b: 1, c: INK, f: HEAD });
  card(s, 1.176, 5.485, '01', 'The First Service', BODY_J, { tw: 1.602, bw: 2.727 });
  card(s, 9.332, 5.485, '03', 'The First Service', BODY_J, { tw: 1.654, bw: 2.727 });
  card(s, 5.145, 5.485, '02', 'The First Service', BODY_J, { tw: 2.119, bw: 2.727 });
}

function slide28(s) {  // four-step curved banner timeline
  // four curved banners; the odd ones sit slightly lower than the even ones
  BANNERS.forEach(function (b) {
    stem(s, b.x + 1.09, b.sy, b.sh, SKY);
    freeform(s, b.x, b.y, b.w, b.h, b.dark ? BLUE : SKY, b.dark ? BANNER_DARK : BANNER_LIGHT);
    freeform(s, b.x, b.y, b.w, b.lipH, b.dark ? NAVY : SKY_MID,
             b.dark ? BANNER_DARK_LIP : BANNER_LIGHT_LIP);
    if (b.dark) {
      freeform(s, b.x, b.y + 0.161, 0.102, 1.411, STEEL, BANNER_EDGE_L);
      freeform(s, b.x + b.w - 0.102, b.y + 0.161, 0.102, 1.411, STEEL, BANNER_EDGE_R);
    }
  });
  circle(s, 2.431, 3.002, 0.865, 0.865, WHITE, { ln: SKY });
  circle(s, 4.932, 2.762, 0.865, 0.865, WHITE, { ln: STEEL });
  circle(s, 7.433, 3.002, 0.865, 0.865, WHITE, { ln: SKY });
  circle(s, 9.933, 2.762, 0.865, 0.865, WHITE, { ln: STEEL });
  dots(s, 0.591, 0.579, SKY);
  nextPage(s, 11.902, 0.543, 7.462, 0.751, 4.073, BLUE_LINK);
  text(s, 7.337, 1.318, 4.333, 0.831, BODY_E, { sz: 10, c: GRAY, ls: 1.5 });
  text(s, 1.663, 1.319, 3.98, 0.774, HEADLINE, { sz: 20, b: 1, c: INK, f: HEAD });
  text(s, 1.792, 5.791, 2.144, 0.303, 'The First Service', { sz: 12, b: 1, c: INK, f: HEAD, al: 'center' });
  text(s, 1.792, 6.242, 2.144, 0.531, BODY_B, { sz: 9, c: GRAY, al: 'center', ls: 1.5 });
  text(s, 4.292, 5.791, 2.144, 0.303, 'The Second Service', { sz: 12, b: 1, c: INK, f: HEAD, al: 'center' });
  text(s, 4.292, 6.242, 2.144, 0.531, BODY_B, { sz: 9, c: GRAY, al: 'center', ls: 1.5 });
  text(s, 6.793, 5.791, 2.144, 0.303, 'The Third Service', { sz: 12, b: 1, c: INK, f: HEAD, al: 'center' });
  text(s, 6.793, 6.242, 2.144, 0.531, BODY_B, { sz: 9, c: GRAY, al: 'center', ls: 1.5 });
  text(s, 9.294, 5.791, 2.144, 0.303, 'The Fourth Service', { sz: 12, b: 1, c: INK, f: HEAD, al: 'center' });
  text(s, 9.294, 6.242, 2.144, 0.531, BODY_B, { sz: 9, c: GRAY, al: 'center', ls: 1.5 });
  text(s, 2.431, 3.259, 0.865, 0.404, '01', { sz: 18, b: 1, c: SKY, f: HEAD, al: 'center' });
  text(s, 4.93, 2.992, 0.865, 0.404, '02', { sz: 18, b: 1, c: STEEL, f: HEAD, al: 'center' });
  text(s, 7.429, 3.232, 0.865, 0.404, '03', { sz: 18, b: 1, c: SKY, f: HEAD, al: 'center' });
  text(s, 9.927, 2.99, 0.865, 0.404, '04', { sz: 18, b: 1, c: STEEL, f: HEAD, al: 'center' });
}

function slide29(s) {  // four-lobe circular diagram with labels either side
  // interlocking petals of the circular diagram
  freeform(s, 5.046, 4.712, 1.658, 1.954, SKY,  PETAL_A);
  freeform(s, 4.786, 3.172, 1.959, 3.142, BLUE, PETAL_B);
  freeform(s, 5.146, 2.917, 3.143, 1.96,  SKY,  PETAL_C);
  freeform(s, 6.586, 3.275, 1.962, 3.139, BLUE, PETAL_D);
  freeform(s, 6.606, 4.811, 1.58,  1.858, SKY,  PETAL_E);
  dots(s, 0.591, 0.579, SKY);
  nextPage(s, 11.902, 0.543, 7.517, 0.751, 4.073, BLUE_LINK);
  text(s, 7.393, 1.318, 4.333, 0.831, BODY_E, { sz: 10, c: GRAY, ls: 1.5 });
  text(s, 1.677, 1.319, 3.98, 0.774, HEADLINE, { sz: 20, b: 1, c: INK, f: HEAD });
  text(s, 9.582, 3.351, 2.144, 0.303, 'The Second Service', { sz: 12, b: 1, c: INK, f: HEAD });
  text(s, 9.582, 3.801, 2.144, 0.531, BODY_B, { sz: 9, c: GRAY, ls: 1.5 });
  text(s, 9.582, 5.281, 2.144, 0.303, 'The Third Service', { sz: 12, b: 1, c: INK, f: HEAD });
  text(s, 9.582, 5.731, 2.144, 0.531, BODY_B, { sz: 9, c: GRAY, ls: 1.5 });
  text(s, 1.663, 3.351, 2.144, 0.303, 'The First Service', { sz: 12, b: 1, c: INK, f: HEAD, al: 'right' });
  text(s, 1.663, 3.801, 2.144, 0.531, BODY_B, { sz: 9, c: GRAY, al: 'right', ls: 1.5 });
  text(s, 1.663, 5.281, 2.144, 0.303, 'The Fourth Service', { sz: 12, b: 1, c: INK, f: HEAD, al: 'right' });
  text(s, 1.663, 5.731, 2.144, 0.531, BODY_B, { sz: 9, c: GRAY, al: 'right', ls: 1.5 });
  text(s, 5.442, 3.769, 0.865, 0.505, '01', { sz: 24, b: 1, c: WHITE, f: HEAD, al: 'center' });
  text(s, 7.041, 3.769, 0.865, 0.505, '02', { sz: 24, b: 1, c: WHITE, f: HEAD, al: 'center' });
  text(s, 5.442, 5.329, 0.865, 0.505, '04', { sz: 24, b: 1, c: WHITE, f: HEAD, al: 'center' });
  text(s, 7.041, 5.329, 0.865, 0.505, '03', { sz: 24, b: 1, c: WHITE, f: HEAD, al: 'center' });
}

function slide30(s) {  // closing "Thank You" cover
  band(s, 2.984, 2.783, 7.365, 2.057, BLUE, { t: 15 });
  band(s, 6.343, 2.788, 0.647, 4.104, SKY, { rot: 90 });
  text(s, 3.731, 3.067, 5.871, 1.212, 'Thank You', { sz: 66, b: 1, c: WHITE, f: HEAD, al: 'center' });
  text(s, 5.171, 4.699, 2.991, 0.281, 'Presentation Template', { sz: 8, c: BLUE, al: 'center', ls: 1.5, cs: 3 });
  dots(s, 0.662, 6.675, WHITE);
  nextPage(s, 9.602, 0.548, 10.66, 0.752, 2.674, WHITE);
}

const SLIDES = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
];

// ------------------------------------------------------------------ build
function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'CHARL', width: 13.333, height: 7.5 });
  pptx.layout = 'CHARL';
  pptx.author = 'pptxgenjs';
  pptx.title = 'Charl Presentation Template';

  SLIDES.forEach(function (fn) { fn(pptx.addSlide()); });

  return pptx.writeFile({
    fileName: path.join(__dirname, '0b7f592f-374d-4500-a995-3168e9b75762_grok_final.pptx'),
  });
}

build().then(function (f) { console.log('wrote ' + f); },
             function (e) { console.error(e); process.exit(1); });
