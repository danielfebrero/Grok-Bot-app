/**
 * Disease and Virus Dynamics — 30-slide deck rebuilt with pptxgenjs.
 *
 * Design system
 *   Two half-slide panels form every background: a near-black panel with a soft
 *   radial glow and a purple panel with a vertical gradient. Both gradients are
 *   emulated with banded solid shapes because pptxgenjs has no gradient fill.
 *   On top sit three recurring vector motifs — an organic "blob", a "wave"
 *   ribbon and a spiked virus icon — plus a small type scale (hero / title /
 *   eyebrow / kicker / body / tagline).
 */
const fs = require('fs');
const path = require('path');
const PptxGenJS = require('pptxgenjs');

const OUT = path.join(__dirname, '11958917-ed72-4f92-98e6-7ea4c89eb93b_grok_final.pptx');

// ---------------------------------------------------------------- palette ---
const WHITE = 'FFFFFF';
const BLACK = '000000';
const PURPLE = '691CAF';       // virus icon purple
const GRAD_TOP = '541290';     // blob gradient, top
const GRAD_BOT = '9032EA';     // blob gradient, bottom
const PANEL_TOP = '7421BF';    // purple panel gradient, top
const PANEL_BOT = '9032EA';    // purple panel gradient, bottom
const DARK_IN = '262626';      // dark panel, centre of the radial glow
const DARK_OUT = '131313';     // dark panel, outer edge
const LILAC = 'E2CBF9';        // eyebrow + tagline text
const LILAC2 = 'CBA1F5';       // tagline on the cover slides
const GRID = 'D9D9D9';         // chart gridlines

const F_HEAD = 'Merriweather Sans';
const F_KICK = 'Poppins SemiBold';
const F_BODY = 'Lato';
const F_TAG = 'Open Sans Light';

// The shadow every blob/leaf shape carries (outerShdw, 45deg, black @38%).
const BLOB_SHADOW = { type: 'outer', blur: 6, offset: 13, angle: 45, color: BLACK, opacity: 0.38 };

// ---------------------------------------------------------------- helpers ---
function mix(a, b, t) {
  let out = '';
  for (let i = 0; i < 3; i++) {
    const ca = parseInt(a.substr(i * 2, 2), 16);
    const cb = parseInt(b.substr(i * 2, 2), 16);
    out += Math.round(ca + (cb - ca) * t).toString(16).padStart(2, '0');
  }
  return out.toUpperCase();
}

/** Purple panel: vertical gradient faked with horizontal bands. */
function purplePanel(s, x, y, w, h, bands) {
  const n = bands || 30;
  for (let i = 0; i < n; i++) {
    s.addShape('rect', {
      x: x, y: y + (i * h) / n, w: w, h: h / n + 0.012,
      fill: { color: mix(PANEL_TOP, PANEL_BOT, i / (n - 1)) }, line: { type: 'none' },
    });
  }
}

/** Dark panel: radial glow faked with concentric ellipses over a flat base. */
function darkPanel(s, x, y, w, h) {
  const n = 18;
  s.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: DARK_OUT }, line: { type: 'none' } });
  for (let i = n; i >= 1; i--) {
    const f = i / n;
    s.addShape('ellipse', {
      x: x + (w * (1 - f)) / 2, y: y + (h * (1 - f)) / 2, w: w * f, h: h * f,
      fill: { color: mix(DARK_IN, DARK_OUT, f) }, line: { type: 'none' },
    });
  }
}

/** Flat colour card (solid rectangles that sit on top of a panel). */
function card(s, x, y, w, h, color) {
  s.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: color }, line: { type: 'none' } });
}

// ------------------------------------------------------------- geometries ---
// Each path is stored in a 0..1 unit square: ['m',x,y] move, ['c',x1,y1,x2,y2,x,y]
// cubic bezier, ['l',x,y] line, ['z'] close.
const BLOB_PATH = [
  ['m', 0.0427, 0.9594],
  ['c', -0.1448, 0.7975, 0.3206, -0.2044, 0.7404, 0.0375],
  ['c', 1.1603, 0.2793, 1.0163, 1.1869, 0.6943, 0.9594],
  ['c', 0.3722, 0.7318, 0.2302, 1.1212, 0.0427, 0.9594],
  ['z'],
];

const WAVE_PATH = [
  ['m', 0.5811, 0.0169],
  ['c', 0.4015, -0.0639, 0.2218, 0.1528, 0.0707, 0.4930],
  ['l', 0.0000, 0.6700],
  ['l', 1.0000, 1.0000],
  ['l', 0.9809, 0.8631],
  ['c', 0.9258, 0.5471, 0.8311, 0.2679, 0.6885, 0.1025],
  ['c', 0.6528, 0.0611, 0.6170, 0.0331, 0.5811, 0.0169],
  ['z'],
];

// One virus spike: a bar with a rounded cap, drawn pointing right.
const SPIKE_PATH = [
  ['m', 1.0000, 0.1220],
  ['l', 1.0000, 0.8780],
  ['c', 1.0000, 0.9454, 0.9454, 1.0000, 0.8780, 1.0000],
  ['c', 0.8106, 1.0000, 0.7561, 0.9454, 0.7561, 0.8780],
  ['l', 0.7561, 0.6220],
  ['l', 0.1220, 0.6220],
  ['c', 0.0546, 0.6220, 0.0000, 0.5674, 0.0000, 0.5000],
  ['c', 0.0000, 0.4327, 0.0546, 0.3780, 0.1220, 0.3780],
  ['l', 0.7561, 0.3780],
  ['l', 0.7561, 0.1220],
  ['c', 0.7561, 0.0546, 0.8106, 0.0000, 0.8780, 0.0000],
  ['c', 0.9454, 0.0000, 1.0000, 0.0546, 1.0000, 0.1220],
  ['z'],
];

/** Scale a unit-square path into a w x h box of pptxgenjs custGeom points. */
function points(unitPath, w, h) {
  return unitPath.map(function (seg) {
    if (seg[0] === 'z') return { close: true };
    if (seg[0] === 'm') return { x: seg[1] * w, y: seg[2] * h, moveTo: true };
    if (seg[0] === 'l') return { x: seg[1] * w, y: seg[2] * h };
    return {
      x: seg[5] * w, y: seg[6] * h,
      curve: { type: 'cubic', x1: seg[1] * w, y1: seg[2] * h, x2: seg[3] * w, y2: seg[4] * h },
    };
  });
}

/** Sample a unit path into a closed polygon (beziers flattened to segments). */
function flatten(unitPath, steps) {
  const poly = [];
  let cur = [0, 0];
  unitPath.forEach(function (seg) {
    if (seg[0] === 'm' || seg[0] === 'l') {
      cur = [seg[1], seg[2]];
      poly.push(cur);
    } else if (seg[0] === 'c') {
      const p = cur;
      for (let i = 1; i <= steps; i++) {
        const t = i / steps;
        const u = 1 - t;
        poly.push([
          u * u * u * p[0] + 3 * u * u * t * seg[1] + 3 * u * t * t * seg[3] + t * t * t * seg[5],
          u * u * u * p[1] + 3 * u * u * t * seg[2] + 3 * u * t * t * seg[4] + t * t * t * seg[6],
        ]);
      }
      cur = [seg[5], seg[6]];
    }
  });
  return poly;
}

/** Sutherland-Hodgman clip of a closed polygon down to the band y >= minY. */
function clipBelow(poly, minY) {
  const out = [];
  for (let i = 0; i < poly.length; i++) {
    const a = poly[i];
    const b = poly[(i + 1) % poly.length];
    if (a[1] >= minY) out.push(a);
    if ((a[1] >= minY) !== (b[1] >= minY)) {
      const t = (minY - a[1]) / (b[1] - a[1]);
      out.push([a[0] + (b[0] - a[0]) * t, minY]);
    }
  }
  return out;
}

function polyPoints(poly, w, h) {
  const pts = poly.map(function (p, i) { return { x: p[0] * w, y: p[1] * h, moveTo: i === 0 }; });
  pts.push({ close: true });
  return pts;
}

const BLOB_POLY = flatten(BLOB_PATH, 18);

/** Blob with its top-to-bottom purple gradient approximated in 14 bands. */
function gradBlob(s, o) {
  const base = { x: o.x, y: o.y, w: o.w, h: o.h, rotate: o.rot, flipH: !!o.flipH, line: { type: 'none' } };
  s.addShape('custGeom', Object.assign({}, base, {
    points: points(BLOB_PATH, o.w, o.h), fill: { color: GRAD_TOP }, shadow: BLOB_SHADOW,
  }));
  const n = 14;
  for (let i = 1; i < n; i++) {
    s.addShape('custGeom', Object.assign({}, base, {
      points: polyPoints(clipBelow(BLOB_POLY, i / n), o.w, o.h),
      fill: { color: mix(GRAD_TOP, GRAD_BOT, i / (n - 1)) },
    }));
  }
}

/** Small white "leaf" flourish that punctuates the headings. */
function accent(s, x, y, rot, w, h) {
  s.addShape('custGeom', {
    x: x, y: y, w: w || 0.4, h: h || 0.257, rotate: rot, flipH: true,
    points: points(BLOB_PATH, w || 0.4, h || 0.257),
    fill: { color: WHITE }, line: { type: 'none' }, shadow: BLOB_SHADOW,
  });
}

/** Round white bullet used for the list on the "Disease Prevention" slide. */
function dot(s, x, y) {
  s.addShape('ellipse', { x: x, y: y, w: 0.4, h: 0.257, rotate: 45, fill: { color: WHITE }, line: { type: 'none' } });
}

/**
 * The blob + wave + leaf trio that anchors the corner of nearly every slide.
 * variants: default, `mirror` (flipped, used on slide 2) and `big` (covers).
 */
function cluster(s, x, y, opt) {
  const o = opt || {};
  if (o.big) {
    gradBlob(s, { x: x, y: y, w: 4.157, h: 2.675, rot: 6.02, flipH: true });
    s.addShape('custGeom', {
      x: x + 0.42, y: y + 2.564, w: 3.089, h: 0.987, rotate: 6.02, flipH: true,
      points: points(WAVE_PATH, 3.089, 0.987), fill: { color: WHITE }, line: { type: 'none' },
    });
    gradBlob(s, { x: x - 0.233, y: y + 0.293, w: 0.879, h: 0.565, rot: 120, flipH: true });
    return;
  }
  const mirrored = !!o.mirror;
  gradBlob(s, { x: x, y: y, w: 3.105, h: 1.998, rot: mirrored ? 6.02 : 353.98, flipH: mirrored });
  s.addShape('custGeom', {
    x: x + (o.waveDx === undefined ? 0.27 : o.waveDx), y: y + 1.909, w: 2.308, h: 0.738,
    rotate: 6.02, flipH: true,
    points: points(WAVE_PATH, 2.308, 0.738), fill: { color: WHITE }, line: { type: 'none' },
  });
  if (mirrored) gradBlob(s, { x: x - 0.169, y: y + 0.268, w: 0.656, h: 0.422, rot: 120, flipH: true });
  else gradBlob(s, { x: x + 1.224, y: y + 1.656, w: 0.656, h: 0.422, rot: 30, flipH: true });
}

// ------------------------------------------------------------- virus icon ---
// Eight spikes around a disc, then the "pores" inside it. All values are
// fractions of the icon's bounding box.
const SPIKES = [
  [0.0963, 0.1324, 220.6, false],
  [0.6726, 0.1231, 139.4, true],
  [0.3753, 0.0101, 90.0, true],
  [0.6857, 0.6496, 40.6, false],
  [0.0891, 0.6489, 319.4, true],
  [0.3858, 0.7712, 270.0, true],
  [0.7617, 0.3944, 0.0, false],
  [0.0000, 0.3906, 180.0, false],
];
const SPIKE_W = 0.2383;
const SPIKE_H = 0.2188;
const PORES = [
  ['ring', 0.3438, 0.3137, 0.1199, 1.5],
  ['ring', 0.4195, 0.5661, 0.0960, 3.0],
  ['ring', 0.6179, 0.6156, 0.0683, 2.25],   // rendered as an open arc
  ['fill', 0.6141, 0.3944, 0.0328],
  ['fill', 0.2753, 0.5719, 0.0411],
  ['fill', 0.3848, 0.7254, 0.0328],
  ['fill', 0.7453, 0.5337, 0.0328],
];

function virus(s, x, y, size, body, pore) {
  const h = size * 0.9971;                    // icon is a hair wider than tall
  const lw = size / 2.2593;                   // pore outlines scale with the icon
  SPIKES.forEach(function (sp) {
    s.addShape('custGeom', {
      x: x + sp[0] * size, y: y + sp[1] * h, w: SPIKE_W * size, h: SPIKE_H * h,
      rotate: sp[2], flipH: sp[3],
      points: points(SPIKE_PATH, SPIKE_W * size, SPIKE_H * h),
      fill: { color: body }, line: { type: 'none' },
    });
  });
  s.addShape('ellipse', {
    x: x + 0.1715 * size, y: y + 0.1694 * h, w: 0.667 * size, h: 0.6687 * h,
    fill: { color: body }, line: { type: 'none' },
  });
  PORES.forEach(function (p) {
    const box = { x: x + p[1] * size, y: y + p[2] * h, w: p[3] * size, h: p[3] * h };
    if (p[0] === 'fill') s.addShape('ellipse', Object.assign(box, { fill: { color: pore }, line: { type: 'none' } }));
    else if (p[4] === 2.25) s.addShape('arc', Object.assign(box, { angleRange: [270, 183], line: { color: pore, width: p[4] * lw } }));
    else s.addShape('ellipse', Object.assign(box, { fill: { type: 'none' }, line: { color: pore, width: p[4] * lw } }));
  });
}

// ----------------------------------------------------------- backgrounds ----
// Which half-slide panels each of the deck's 28 layouts is built from.
const LAYOUTS = {
  '24': [['dark', 0, 0, 6.667, 7.5], ['purple', 6.667, 0, 6.667, 7.5]],
  '25': [['dark', 0, 0, 6.667, 7.5], ['purple', 6.667, 0, 6.667, 7.5]],
  '26': [['dark', 6.667, 0, 6.667, 7.5], ['purple', 0, 0, 6.667, 7.5]],
  '27': [['purple', 0, 0, 6.667, 7.5], ['dark', 3.952, 0, 9.381, 7.5]],
  '28': [['purple', 0, 0, 7.611, 7.5], ['dark', 7.611, 0, 5.722, 7.5]],
  '29': [['purple', 6.667, 0, 6.667, 7.5], ['dark', 0, 0, 6.667, 7.5]],
  '30': [['purple', -0, 0, 6.794, 7.5], ['dark', 4.873, 0, 8.587, 7.5]],
  '31': [['purple', 6.54, 0, 6.794, 7.5], ['dark', 0, 0, 6.54, 7.5]],
  '32': [['dark', 0, 0, 13.333, 7.5], ['purple', 0, 3.25, 8.365, 4.25]],
  '33': [['dark', 0, 0, 13.333, 7.5]],
  '34': [['dark', 0, 0, 13.333, 7.5], ['purple', 7.875, 0, 5.458, 7.5]],
  '35': [['dark', 0, 0, 13.333, 7.5], ['purple', 3.938, 0, 5.458, 7.5]],
  '36': [['dark', 0, 0, 13.333, 7.5], ['purple', 3.297, 4.063, 3.297, 3.437], ['purple', 9.891, 4.063, 3.443, 3.437]],
  '37': [['dark', 0, 0, 13.333, 7.5], ['purple', 3.297, 0, 4.957, 7.5]],
  '38': [['dark', 0, 0, 13.333, 7.5], ['purple', 0, 0, 4.957, 7.5]],
  '39': [['dark', 0, 0, 13.333, 7.5], ['purple', 0, 0, 6.667, 7.5]],
  '40': [['dark', 0, 0, 6.667, 7.5], ['purple', 6.667, 0, 6.667, 7.5]],
  '41': [['dark', 6.667, 0, 6.667, 7.5], ['purple', 0, 0, 6.667, 7.5]],
  '42': [['dark', 6.667, 0, 6.667, 7.5], ['purple', 0, 0, 6.667, 7.5]],
  '43': [['dark', 0, 0, 6.667, 7.5], ['purple', 6.667, 0, 6.667, 7.5], ['image', 0.951, 1.841, 4.981, 4.136, '[monitor]', '343236']],
  '44': [['dark', 0, 0, 6.667, 7.5], ['purple', 6.667, 0, 6.667, 7.5], ['image', 8.431, 1.35, 3.137, 4.799, '[tablet]', '672B9F']],
  '45': [['purple', 0, 0, 6.667, 7.5], ['dark', 6.667, 0, 6.667, 7.5]],
  '46': [['purple', 6.667, 0, 6.667, 7.5], ['dark', 0, 0, 7.417, 7.5]],
  '47': [['dark', 0, 0, 7.417, 7.5], ['purple', 5.917, 0, 7.417, 7.5]],
  '48': [['dark', 0, 0, 7.417, 7.5], ['purple', 5.917, 0, 7.417, 7.5]],
  '49': [['dark', 0, 0, 7.417, 7.5], ['purple', 7.417, 0, 5.917, 7.5]],
  '50': [['dark', 0, 0, 7.417, 7.5], ['purple', 7.417, 0, 5.917, 7.5]],
  '51': [['dark', 5.762, 0, 7.571, 7.5], ['purple', -0.154, 0, 5.917, 7.5]],
};

/**
 * Stand-in for a raster photo in the reference deck: a soft grey rectangle
 * carrying a short caption, positioned and sized like the original image.
 */
function imagePlaceholder(s, x, y, w, h, label, color) {
  s.addShape('rect', {
    x: x, y: y, w: w, h: h,
    fill: { color: color }, line: { color: mix(color, WHITE, 0.35), width: 0.75, dashType: 'dash' },
  });
  s.addText(label, {
    x: x, y: y, w: w, h: h, fontFace: F_BODY, fontSize: 11, color: mix(color, WHITE, 0.7),
    charSpacing: 3, align: 'center', valign: 'middle',
  });
}

function panels(s, key) {
  LAYOUTS[key].forEach(function (p) {
    if (p[0] === 'dark') darkPanel(s, p[1], p[2], p[3], p[4]);
    else if (p[0] === 'image') imagePlaceholder(s, p[1], p[2], p[3], p[4], p[5], p[6]);
    else purplePanel(s, p[1], p[2], p[3], p[4]);
  });
}

// ------------------------------------------------------------------ type ----
function hero(s, text, x, y, w, h, size, color) {
  s.addText(text, { x: x, y: y, w: w, h: h, fontFace: F_HEAD, fontSize: size, bold: true, color: color, valign: 'top' });
}

function title(s, text, x, y, w, h) {
  s.addText(text, { x: x, y: y, w: w, h: h, fontFace: F_HEAD, fontSize: 28, bold: true, color: WHITE, valign: 'top' });
}

function eyebrow(s, text, x, y, w, h) {
  s.addText(text, {
    x: x, y: y, w: w, h: h, fontFace: F_BODY, fontSize: 12, color: LILAC,
    charSpacing: 3, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top',
  });
}

function kicker(s, text, x, y, w, h, opt) {
  const o = opt || {};
  s.addText(text, {
    x: x, y: y, w: w, h: h, fontFace: F_KICK, fontSize: 12, color: WHITE,
    charSpacing: o.align ? 0 : 3, align: o.align || 'left', wrap: !!o.wrap, valign: 'top',
  });
}

function body(s, text, x, y, w, h, opt) {
  const o = opt || {};
  s.addText(text, {
    x: x, y: y, w: w, h: h, fontFace: F_BODY, fontSize: 11, color: WHITE,
    align: o.align || 'justify', lineSpacingMultiple: 1.5, valign: 'top',
  });
}

function tagline(s, text, x, y, w, h, color) {
  s.addText(text, {
    x: x, y: y, w: w, h: h, fontFace: F_TAG, fontSize: 14, italic: true,
    color: color || LILAC, charSpacing: 3, valign: 'top',
  });
}

// ----------------------------------------------------------------- chart ----
const CHART_YEARS = ['2018', '2019', '2020', '2021', '2022', '2023'];
const CHART_VALUES = [1, 2, 1, 3, 5, 6];

function lineChart(s, x, y, w, h) {
  s.addChart('line', [{ name: 'Series 1', labels: CHART_YEARS, values: CHART_VALUES }], {
    x: x, y: y, w: w, h: h,
    chartColors: [WHITE], lineSize: 2.25, lineDataSymbolSize: 5,
    lineDataSymbolLineColor: WHITE, lineSmooth: false,
    showLegend: false, showTitle: false,
    valAxisMaxVal: 7, valAxisMajorUnit: 1, valAxisLineShow: false,
    valGridLine: { color: GRID, size: 0.75 }, catGridLine: { style: 'none' },
    catAxisLineColor: GRID,
    catAxisLabelColor: WHITE, catAxisLabelFontFace: F_BODY, catAxisLabelFontSize: 10,
    valAxisLabelColor: WHITE, valAxisLabelFontFace: F_BODY, valAxisLabelFontSize: 10,
  });
}

/**
 * The reference deck embeds a shaded world-map chart here. Raster/geo artwork is
 * replaced by a labelled placeholder rectangle, per the conversion brief.
 */
function mapPlaceholder(s, x, y, w, h) {
  s.addShape('rect', {
    x: x, y: y, w: w, h: h,
    fill: { color: 'A163DB' }, line: { color: 'D9CBEE', width: 0.75, dashType: 'dash' },
  });
  s.addText('[map]', {
    x: x, y: y, w: w, h: h, fontFace: F_BODY, fontSize: 11, color: 'EFE5FB',
    charSpacing: 3, align: 'center', valign: 'middle',
  });
}

/** Solid-purple virus used as the three big labelled icons on slide 4. */
function virusPlain(s, x, y, size) {
  const h = size * 0.9779;
  SPIKES.forEach(function (sp) {
    s.addShape('custGeom', {
      x: x + sp[0] * size, y: y + sp[1] * h, w: SPIKE_W * size, h: SPIKE_H * h,
      rotate: sp[2], flipH: sp[3],
      points: points(SPIKE_PATH, SPIKE_W * size, SPIKE_H * h),
      fill: { color: PURPLE }, line: { type: 'none' },
    });
  });
  s.addShape('ellipse', {
    x: x + 0.1715 * size, y: y + 0.1626 * h, w: 0.667 * size, h: 0.6824 * h,
    fill: { color: PURPLE }, line: { type: 'none' },
  });
}

function slide01(s) {
  panels(s, '24');
  cluster(s, 5.197, 4.108, { big: true });
  virus(s, 5.163, 2.554, 1.397, '691CAF', 'FFFFFF');
  accent(s, 10.857, 2.339, 30, 0.473, 0.305);
  hero(s, 'DISEASE', 6.913, 2.492, 6.611, 1.212, 66, 'FFFFFF');
  hero(s, 'AND VIRUS', 6.944, 3.443, 6.611, 0.926, 49, 'FFFFFF');
  hero(s, 'DYNAMICS', 6.897, 4.166, 6.611, 1.01, 53, '000000');
  tagline(s, 'NAVIGATING IMPACT AND PROTECTION', 6.962, 5.463, 6.611, 0.337, 'CBA1F5');
}

function slide02(s) {
  panels(s, '25');
  cluster(s, 6.225, 4.985, { mirror: true });
  virus(s, 6.824, 4.591, 1.397, 'FFFFFF', '691CAF');
  accent(s, 4.028, 1.905, 45.05);
  title(s, 'Introduction', 1.333, 2.043, 3.768, 0.572);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 8.461, 5.99, 4.237, 0.572);
  kicker(s, 'UNDERSTANDING \nDISEASES & VIRUSES', 1.333, 3.613, 2.63, 0.505);
  body(s, 'PLACEHOLDER', 1.333, 4.059, 3.883, 1.731);
  eyebrow(s, 'About us', 1.362, 1.669, 4.035, 0.364);
}

function slide03(s) {
  panels(s, '26');
  cluster(s, 4.209, 4.985, {});
  virus(s, 5.372, 4.626, 1.397, 'FFFFFF', '691CAF');
  accent(s, 10.331, 1.746, 45.05);
  title(s, 'Common Infectious Diseases', 8.35, 1.836, 3.768, 1.515);
  kicker(s, 'FLU TO COVID 19', 8.35, 4.09, 2.181, 0.303);
  body(s, 'PLACEHOLDER', 8.35, 4.403, 3.768, 1.731);
  eyebrow(s, 'About us', 8.35, 1.43, 4.035, 0.364);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 2.952, 5.99, 4.237, 0.572);
}

function slide04(s) {
  panels(s, '27');
  cluster(s, 1.698, 4.985, {});
  virusPlain(s, 5.063, 4.004, 2.599);
  virusPlain(s, 7.565, 4.004, 2.599);
  virusPlain(s, 10.033, 4.014, 2.599);
  virus(s, 2.787, 4.685, 1.124, 'FFFFFF', '691CAF');
  accent(s, 7.065, 1.09, 45.05);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 0.441, 5.99, 4.237, 0.572);
  title(s, 'Disease Classification', 5.353, 1.229, 3.768, 1.043);
  eyebrow(s, 'About us', 5.382, 0.855, 4.035, 0.364);
  body(s, 'PLACEHOLDER', 5.373, 2.418, 5.054, 0.897);
  kicker(s, 'INFECTIOUS', 5.772, 5.172, 1.177, 0.303, { align: 'center' });
  kicker(s, 'NON \nINFECTIOUS', 8.29, 5.099, 1.177, 0.505, { align: 'center' });
  kicker(s, 'GENETIC', 10.88, 5.172, 0.908, 0.303, { align: 'center' });
}

function slide05(s) {
  panels(s, '28');
  cluster(s, 4.511, 4.985, {});
  virus(s, 5.207, 4.626, 1.397, 'FFFFFF', '691CAF');
  accent(s, 9.014, 5.518, 45.05);
  accent(s, 3.734, 2.334, 45.05);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 6.539, 5.823, 4.237, 0.572);
  title(s, 'Viral \nReplication', 1.358, 2.043, 3.768, 1.043);
  eyebrow(s, 'About us', 1.376, 1.637, 4.035, 0.364);
  kicker(s, 'HOST CELL VIRUSES', 1.348, 3.743, 2.591, 0.303);
  body(s, 'PLACEHOLDER', 1.348, 4.015, 3.025, 1.731);
}

function slide06(s) {
  panels(s, '29');
  cluster(s, 5.289, 4.985, {});
  virus(s, 5.985, 4.626, 1.397, 'FFFFFF', '691CAF');
  accent(s, 3.752, 1.934, 45.05);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 6.943, 6.027, 4.237, 0.572);
  title(s, 'Viral \nReplication', 1.321, 1.629, 3.768, 1.043);
  eyebrow(s, 'About us', 1.34, 1.223, 4.035, 0.364);
  kicker(s, 'PANDEMIC', 1.361, 3.367, 1.424, 0.303);
  body(s, 'PLACEHOLDER', 1.361, 3.639, 3.025, 1.731);
  kicker(s, 'EPIDEMIC', 8.834, 3.363, 1.315, 0.303);
  body(s, 'PLACEHOLDER', 8.834, 3.635, 3.025, 1.731);
}

function slide07(s) {
  panels(s, '30');
  cluster(s, 10.068, 4.985, {});
  virus(s, 11.312, 4.241, 1.397, '691CAF', 'FFFFFF');
  accent(s, 3.51, 2.23, 45.05);
  dot(s, 1.185, 4.031);
  dot(s, 1.167, 4.68);
  dot(s, 1.167, 5.326);
  dot(s, 1.167, 6.037);
  title(s, 'Disease Prevention', 1.134, 1.94, 3.768, 1.043);
  eyebrow(s, 'About us', 1.153, 1.534, 4.035, 0.364);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 10.423, 5.927, 4.237, 0.572);
  kicker(s, 'HAND HYGIENE', 1.75, 4.008, 1.943, 0.303);
  kicker(s, 'MASK USE', 1.731, 4.657, 1.364, 0.303);
  kicker(s, 'VACCINATION ', 1.731, 5.303, 1.939, 0.303);
  kicker(s, 'AND MORE', 1.731, 6.014, 1.426, 0.303);
  body(s, 'PLACEHOLDER', 5.992, 4.977, 3.172, 1.453);
}

function slide08(s) {
  panels(s, '31');
  card(s, 0.771, 5.317, 2.364, 2.183, '882DDD');
  cluster(s, 1.433, 4.985, {});
  virus(s, 1.572, 4.294, 1.397, 'FFFFFF', '691CAF');
  accent(s, 10.884, 2.092, 45.05);
  title(s, 'Immunization Importance', 8.016, 2.195, 3.768, 1.043);
  eyebrow(s, 'Medical', 8.034, 1.789, 4.035, 0.364);
  kicker(s, 'DEFENSES AGAINST DISEASES', 8.016, 3.927, 3.655, 0.303);
  body(s, 'PLACEHOLDER', 8.016, 4.23, 3.768, 1.731);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 2.446, 5.815, 4.237, 0.572);
}

function slide09(s) {
  panels(s, '32');
  cluster(s, 6.426, 4.985, {});
  mapPlaceholder(s, 1.014, 3.931, 4.472, 2.514);
  virus(s, 7.67, 4.241, 1.397, '691CAF', 'FFFFFF');
  accent(s, 4.174, 1.196, 45.05);
  title(s, 'Global Health Impact', 1.307, 1.299, 3.768, 1.043);
  eyebrow(s, 'Medical', 1.325, 0.893, 4.035, 0.364);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 7.549, 5.865, 4.237, 0.572);
  body(s, 'PLACEHOLDER', 5.859, 1.201, 3.768, 1.175);
}

function slide10(s) {
  panels(s, '33');
  purplePanel(s, 7.556, 1.095, 4.989, 5.508);
  cluster(s, 1.005, 1.253, {});
  virus(s, 1.014, 0.547, 1.397, '691CAF', 'FFFFFF');
  accent(s, 11.069, 2.018, 45.05);
  title(s, 'Global Health Impact', 8.201, 2.121, 3.768, 1.043);
  eyebrow(s, 'Medical', 8.219, 1.715, 4.035, 0.364);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 2.088, 1.943, 4.237, 0.572);
  kicker(s, 'ZOONOTIC DISEASE', 8.219, 3.896, 2.476, 0.303);
  body(s, 'PLACEHOLDER', 8.219, 4.199, 3.768, 1.731);
  kicker(s, 'PREPAREDNESS AND RAPID RESPONSE', 1.209, 4.601, 4.6, 0.303);
  body(s, 'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum ipsum dolor sit es lorem ipsum dolor siabes dolore magna aliqua uta enimvas ad minim veniam nostrud consectetur adipiscing elita sweden aliquie commodo maines  lorem ipsum dol', 1.209, 4.873, 4.569, 1.175);
}

function slide11(s) {
  panels(s, '34');
  cluster(s, 8.68, 4.985, {});
  lineChart(s, 1.429, 4.178, 5.056, 2.549);
  virus(s, 9.768, 4.685, 1.124, 'FFFFFF', '691CAF');
  virus(s, 5.722, 4.348, 0.512, '691CAF', 'FFFFFF');
  virus(s, 4.819, 4.685, 0.512, '691CAF', 'FFFFFF');
  virus(s, 4.257, 5.226, 0.512, '691CAF', 'FFFFFF');
  virus(s, 3.443, 5.81, 0.512, '691CAF', 'FFFFFF');
  virus(s, 2.595, 5.483, 0.512, '691CAF', 'FFFFFF');
  virus(s, 1.804, 5.818, 0.512, '691CAF', 'FFFFFF');
  accent(s, 4.34, 1.319, 45.05);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 9.171, 5.99, 4.237, 0.572);
  title(s, 'Case Study\nEbola', 1.473, 1.422, 3.768, 1.043);
  eyebrow(s, 'Disease', 1.517, 1.016, 2.817, 0.364);
  body(s, 'PLACEHOLDER', 1.475, 2.841, 5.011, 0.897);
}

function slide12(s) {
  panels(s, '35');
  cluster(s, 4.823, 4.985, {});
  virus(s, 6.067, 4.241, 1.397, '691CAF', 'FFFFFF');
  accent(s, 2.72, 1.774, 45.05);
  title(s, 'Role of Healthcare Workers', 0.739, 1.877, 2.835, 1.515);
  eyebrow(s, 'Medic', 0.783, 1.471, 2.817, 0.364);
  kicker(s, 'DISEASE CONTROL', 0.739, 4.108, 2.881, 0.303, { wrap: true });
  body(s, 'PLACEHOLDER', 0.739, 4.38, 2.546, 1.453, { align: 'left' });
  kicker(s, 'SAVES LIVES', 10.048, 4.108, 2.881, 0.303, { wrap: true });
  body(s, 'PLACEHOLDER', 10.048, 4.38, 2.546, 1.453, { align: 'left' });
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 5.946, 5.865, 4.237, 0.572);
}

function slide13(s) {
  panels(s, '36');
  cluster(s, 2.629, 4.985, {});
  virus(s, 3.717, 4.685, 1.124, 'FFFFFF', '691CAF');
  accent(s, 5.277, 1.673, 45.05);
  title(s, 'Genomic Sequencing', 3.297, 1.776, 2.835, 1.043);
  eyebrow(s, 'Medic', 3.341, 1.37, 2.817, 0.364);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 3.121, 5.99, 4.237, 0.572);
  kicker(s, 'UNCOVERS THE GENETIC MAKEUP ', 7.606, 1.372, 4.169, 0.303);
  body(s, 'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum ipsum dolor sit es lorem ipsum dolor siabes dolore magna aliqua uta enimvas ad minim veniam nostrud consectetur adipiscing elita sweden aliquie commodo maines  lorem ipsum dol', 7.606, 1.644, 4.569, 1.175);
}

function slide14(s) {
  panels(s, '37');
  cluster(s, 6.375, 4.985, {});
  virus(s, 5.265, 5.916, 1.397, '691CAF', 'FFFFFF');
  accent(s, 6.723, 1.044, 45.05);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 6.89, 5.781, 3.053, 0.572);
  title(s, 'Vaccination Myths', 4.319, 1.147, 2.835, 1.043);
  eyebrow(s, 'Medic', 4.363, 0.741, 2.817, 0.364);
  kicker(s, 'FACTS', 4.319, 2.937, 0.951, 0.303);
  body(s, 'PLACEHOLDER', 4.319, 3.24, 3.053, 1.453);
  kicker(s, 'MYTHS', 9.246, 2.937, 0.989, 0.303);
  body(s, 'PLACEHOLDER', 9.246, 3.24, 3.053, 1.453);
}

function slide15(s) {
  panels(s, '24');
  cluster(s, 5.197, 4.108, { big: true });
  virus(s, 5.163, 2.554, 1.397, '691CAF', 'FFFFFF');
  accent(s, 10.154, 3, 30, 0.473, 0.305);
  hero(s, 'BREAK', 6.913, 3.269, 6.611, 1.212, 66, 'FFFFFF');
  hero(s, 'SLIDES', 6.897, 4.166, 6.611, 1.212, 66, '000000');
  tagline(s, 'TAKE A 5 MINUTES BREAK', 6.962, 5.463, 6.611, 0.337, 'CBA1F5');
}

function slide16(s) {
  panels(s, '38');
  cluster(s, 1.866, 4.985, {});
  virus(s, 2.954, 4.685, 1.124, 'FFFFFF', '691CAF');
  accent(s, 8.159, 4.861, 45.05);
  title(s, 'Antibiotic Resistance', 6.023, 4.964, 2.835, 1.043);
  eyebrow(s, 'Medic', 6.067, 4.558, 2.817, 0.364);
  kicker(s, 'GLOBAL THREAT', 9.461, 4.726, 2.067, 0.303);
  body(s, 'PLACEHOLDER', 9.461, 5.029, 3.053, 0.897);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 2.357, 5.99, 4.237, 0.572);
}

function slide17(s) {
  panels(s, '39');
  card(s, 8.013, 0, 1.717, 3.748, '882DDD');
  cluster(s, 7.278, 4.985, {});
  virus(s, 8.367, 4.685, 1.124, 'FFFFFF', '691CAF');
  accent(s, 4.298, 1.575, 45.05);
  accent(s, 7.727, 4.388, 315);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 7.77, 5.99, 4.237, 0.572);
  title(s, 'Mental Health Impact', 1.404, 1.678, 3.754, 1.043);
  eyebrow(s, 'Medic', 1.448, 1.272, 2.817, 0.364);
  kicker(s, 'DISEASE OUTBREAKS', 1.448, 3.518, 2.616, 0.303);
  body(s, 'PLACEHOLDER', 1.448, 3.821, 3.053, 0.897);
  kicker(s, 'SUPPORT NETWORKS', 1.448, 5.149, 2.611, 0.303);
  body(s, 'PLACEHOLDER', 1.448, 5.452, 3.053, 0.897);
}

function slide18(s) {
  panels(s, '40');
  cluster(s, 6.982, 4.985, { waveDx: 0.772 });
  virus(s, 8.071, 4.685, 1.124, 'FFFFFF', '691CAF');
  accent(s, 7.431, 4.388, 315);
  accent(s, 2.693, 1.973, 45.05);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 7.976, 5.99, 4.237, 0.572);
  title(s, 'Global Collaboration', 1.208, 2.143, 3.754, 1.043);
  eyebrow(s, 'Medic', 1.252, 1.737, 2.817, 0.364);
  kicker(s, 'MYTHS', 1.257, 3.999, 0.989, 0.303);
  body(s, 'PLACEHOLDER', 1.257, 4.302, 3.277, 1.453);
}

function slide19(s) {
  panels(s, '41');
  cluster(s, 2.812, 4.985, {});
  virus(s, 4.06, 1.613, 1.124, 'FFFFFF', '691CAF');
  accent(s, 10.181, 1.717, 45.05);
  title(s, 'Disease Surveillance', 8.504, 1.888, 3.754, 1.043);
  eyebrow(s, 'Medic', 8.548, 1.482, 2.817, 0.364);
  kicker(s, 'DISEASE PATTERNS', 8.51, 4.11, 2.427, 0.303);
  body(s, 'PLACEHOLDER', 8.51, 4.413, 3.277, 1.453);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 3.304, 5.99, 4.237, 0.572);
}

function slide20(s) {
  panels(s, '42');
  cluster(s, 5.688, 4.985, {});
  virus(s, 7.113, 2.927, 1.397, '691CAF', 'FFFFFF');
  accent(s, 3.304, 1.745, 45.05);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 6.18, 5.99, 4.237, 0.572);
  title(s, 'Vaccine Development', 1.627, 1.916, 3.754, 1.043);
  eyebrow(s, 'Medic', 1.671, 1.51, 2.817, 0.364);
  kicker(s, 'DISCOVERY, PRECLINICAL, \nCLINICAL', 1.634, 4.05, 3.386, 0.505);
  body(s, 'PLACEHOLDER', 1.634, 4.555, 3.277, 1.453);
}

function slide21(s) {
  panels(s, '43');
  cluster(s, 4.517, 4.985, { waveDx: 0.772 });
  virus(s, 5.605, 4.685, 1.124, 'FFFFFF', '691CAF');
  accent(s, 10.181, 1.745, 45.05);
  accent(s, 4.966, 4.388, 315);
  title(s, 'Social \nDistancing', 8.504, 1.916, 3.754, 1.043);
  eyebrow(s, 'Medic', 8.548, 1.51, 2.817, 0.364);
  kicker(s, 'CURBS DISEASE SPREAD', 8.51, 4.138, 3.005, 0.303);
  body(s, 'PLACEHOLDER', 8.51, 4.441, 3.277, 1.453);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 5.51, 5.99, 4.237, 0.572);
}

function slide22(s) {
  panels(s, '44');
  cluster(s, 5.294, 4.985, { waveDx: 0.772 });
  virus(s, 6.383, 4.685, 1.124, 'FFFFFF', '691CAF');
  accent(s, 8.348, 5.057, 315);
  accent(s, 3.365, 1.524, 45.05);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 6.517, 5.911, 4.237, 0.572);
  title(s, 'Disease\nOutbreak Preparedness', 1.384, 1.613, 3.768, 1.515);
  kicker(s, 'PREPAREDNESS PLANS ', 1.384, 3.867, 2.867, 0.303);
  body(s, 'PLACEHOLDER', 1.384, 4.181, 3.768, 1.731);
  eyebrow(s, 'Medic', 1.416, 1.225, 4.035, 0.364);
}

function slide23(s) {
  panels(s, '45');
  cluster(s, 2.53, 4.985, {});
  virus(s, 4.63, 4.347, 1.397, '691CAF', 'FFFFFF');
  accent(s, 10.413, 1.714, 45.05);
  title(s, 'Healthcare Infrastructure', 8.146, 1.804, 3.768, 1.043);
  kicker(s, 'WELL-EQUIPPED FACILITIES', 8.146, 4.058, 3.414, 0.303);
  body(s, 'PLACEHOLDER', 8.146, 4.371, 3.768, 1.731);
  eyebrow(s, 'Medic', 8.178, 1.415, 4.035, 0.364);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 3.021, 5.99, 4.237, 0.572);
}

function slide24(s) {
  panels(s, '46');
  cluster(s, 9.172, 4.985, {});
  virus(s, 10.26, 4.685, 1.124, 'FFFFFF', '691CAF');
  accent(s, 5.324, 3.495, 45.05);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 9.663, 5.99, 4.237, 0.572);
  title(s, 'Future of Disease Control', 1.756, 3.584, 3.768, 1.043);
  kicker(s, 'ADVANCES LIKE AI\nTELEMEDICINE', 1.756, 4.941, 2.462, 0.505);
  body(s, 'PLACEHOLDER', 1.756, 5.355, 3.768, 1.175);
  eyebrow(s, 'Medic', 1.787, 3.196, 4.035, 0.364);
}

function slide25(s) {
  panels(s, '47');
  cluster(s, 3.358, 4.985, { waveDx: 0.772 });
  virus(s, 4.446, 4.685, 1.124, 'FFFFFF', '691CAF');
  accent(s, 10.032, 1.683, 45.05);
  accent(s, 3.823, 4.353, 315);
  title(s, 'Children’s\nHealth', 7.765, 1.772, 3.768, 1.043);
  kicker(s, 'CHILDREN\'S UNIQUE \nVULNERABILITIES ', 7.765, 3.834, 2.623, 0.505);
  body(s, 'PLACEHOLDER', 7.765, 4.339, 3.768, 1.731);
  eyebrow(s, 'Medic', 7.797, 1.384, 4.035, 0.364);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 4.023, 5.959, 4.237, 0.572);
}

function slide26(s) {
  panels(s, '48');
  cluster(s, 3.358, 4.985, { waveDx: 0.772 });
  virus(s, 4.446, 4.685, 1.124, 'FFFFFF', '691CAF');
  accent(s, 3.823, 4.353, 315);
  accent(s, 10.11, 1.601, 45.05);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 4.023, 5.959, 4.237, 0.572);
  title(s, 'Travel Restrictions', 7.843, 1.691, 3.768, 1.043);
  kicker(s, 'MITIGATE DISEASE SPREAD.', 7.843, 3.946, 3.429, 0.303);
  body(s, 'PLACEHOLDER', 7.843, 4.258, 3.768, 1.731);
  eyebrow(s, 'Medic', 7.874, 1.302, 4.035, 0.364);
}

function slide27(s) {
  panels(s, '49');
  cluster(s, 9.283, 4.985, {});
  lineChart(s, 1.19, 4.906, 3.905, 1.968);
  virus(s, 10.371, 4.685, 1.124, 'FFFFFF', '691CAF');
  virus(s, 4.505, 5.038, 0.396, '691CAF', 'FFFFFF');
  virus(s, 3.374, 5.716, 0.396, '691CAF', 'FFFFFF');
  virus(s, 2.09, 5.915, 0.396, '691CAF', 'FFFFFF');
  accent(s, 3.457, 1.07, 45.05);
  title(s, 'Economic \nImpact', 1.19, 1.159, 3.768, 1.043);
  eyebrow(s, 'Medic', 1.221, 0.771, 4.035, 0.364);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 10.371, 5.959, 4.237, 0.572);
  kicker(s, 'DISEASE OUTBREAKS \nDISRUPT ECONOMIES', 1.245, 3.066, 2.698, 0.505);
  body(s, 'PLACEHOLDER', 1.245, 3.51, 3.768, 1.175);
}

function slide28(s) {
  panels(s, '50');
  cluster(s, 2.646, 4.985, {});
  virus(s, 1.447, 4.983, 1.397, '691CAF', 'FFFFFF');
  accent(s, 10.213, 2.308, 45.05);
  title(s, 'Myths \nBusted', 8.761, 2.397, 3.768, 1.043);
  eyebrow(s, 'Medic', 8.792, 2.009, 4.035, 0.364);
  kicker(s, 'DISEASE MYTHS ', 8.792, 4.143, 2.12, 0.303);
  body(s, 'PLACEHOLDER', 8.792, 4.446, 3.277, 1.453);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 3.138, 5.99, 4.237, 0.572);
}

function slide29(s) {
  panels(s, '51');
  cluster(s, 7.535, 4.985, {});
  virus(s, 6.336, 4.983, 1.397, '691CAF', 'FFFFFF');
  accent(s, 3.221, 1.601, 45.05);
  title(s, 'Hope and Resilience', 0.954, 1.691, 3.768, 1.043);
  kicker(s, 'STORIES OF \nRESILIENCE IN DISEASE', 0.954, 3.915, 2.914, 0.505);
  body(s, 'PLACEHOLDER', 0.954, 4.42, 3.768, 1.731);
  eyebrow(s, 'Medic', 0.985, 1.302, 4.035, 0.364);
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 8.171, 5.702, 4.237, 0.572);
}

function slide30(s) {
  panels(s, '24');
  cluster(s, 5.197, 4.108, { big: true });
  virus(s, 5.163, 2.554, 1.397, '691CAF', 'FFFFFF');
  accent(s, 10.154, 3, 30, 0.473, 0.305);
  hero(s, 'THANK', 6.913, 3.269, 6.611, 1.212, 66, 'FFFFFF');
  hero(s, 'YOU', 6.897, 4.166, 6.611, 1.212, 66, '000000');
  tagline(s, 'NAVIGATING IMPACT \nAND PROTECTION', 7.16, 5.361, 4.237, 0.572);
}

// ------------------------------------------------------------------ build ---
const SLIDES = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'WIDE';
pptx.title = 'Disease and Virus Dynamics';

SLIDES.forEach(function (build) {
  build(pptx.addSlide());
});

pptx.writeFile({ fileName: OUT }).then(function () {
  console.log('wrote ' + OUT);
});
