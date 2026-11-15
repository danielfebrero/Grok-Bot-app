/**
 * "Digital Marketing" — Budogol presentation template, rebuilt with pptxgenjs.
 * 25 slides, 13.333in x 7.5in (16:9).
 *
 * Raster artwork from the source deck is replaced by flat placeholder rectangles
 * (see `imagePlaceholder`); everything else is drawn with native pptxgenjs shapes,
 * text boxes, tables and charts.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const C = {
  green: '839878', // accent1
  orange: 'D3663B', // accent2
  stone: 'CCC4C0', // accent3
  gold: 'A68F63', // accent4
  sage: 'D4D4C7', // accent5
  mint: 'B0D9CD', // accent6
  ink: '262626', // tx1 lum85
  body: '404040', // tx1 lum75
  grey: '595959', // tx1 lum65
  white: 'FFFFFF',
  paper: 'F2F2F2',
  placeholder: 'D9D9D9'
};

const HEAD = 'Lato'; // theme major font
const BODY = 'Open Sans'; // theme minor font

const LOREM = {
  short: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ',
  medium: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ',
  long:
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
    'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.',
  nunc: 'Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus.',
  pellen:
    'Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. ' +
    'Proin pharetra nonummy pede. Mauris et orci.',
  magna: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna.',
  fusce: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere.',
  tiny: 'Lorem ipsum dolor sit adipiscing elit. ',
  card: 'Lorem ipsum dolor sit amet, consectetuer adipiscing'
};

/* ------------------------------------------------- wave / freeform geometry */
// Paths are absolute slide inches: ['M',x,y] ['L',x,y] ['C',x1,y1,x2,y2,x,y] ['Z'].

const WAVE_SALMON = [
  ['M', 3.528, 5.636], ['C', 3.919, 5.631, 4.307, 5.657, 4.681, 5.708], ['C', 5.326, 5.796, 5.861, 5.991, 6.422, 6.189],
  ['C', 7.252, 6.482, 8.063, 6.843, 9.012, 6.985], ['C', 9.899, 7.117, 10.875, 7.077, 11.767, 6.823],
  ['C', 12.279, 6.677, 12.864, 6.424, 13.259, 6.107], ['L', 13.333, 6.041], ['L', 13.333, 7.5], ['L', 0, 7.5],
  ['L', 0, 6.812], ['L', 0.074, 6.746], ['C', 0.232, 6.623, 0.415, 6.507, 0.506, 6.448],
  ['C', 0.828, 6.24, 1.202, 6.068, 1.605, 5.937], ['C', 2.219, 5.738, 2.877, 5.645, 3.528, 5.636], ['Z']
];

const WAVE_GREEN = [
  ['M', 3.534, 5.743], ['C', 3.925, 5.75, 4.314, 5.787, 4.689, 5.849], ['C', 5.336, 5.956, 5.874, 6.167, 6.437, 6.381],
  ['C', 7.27, 6.698, 8.086, 7.083, 9.037, 7.252], ['C', 9.927, 7.411, 10.903, 7.4, 11.794, 7.173],
  ['C', 12.305, 7.043, 12.888, 6.807, 13.281, 6.502], ['L', 13.333, 6.457], ['L', 13.333, 7.523], ['L', 0, 7.523],
  ['L', 0, 6.824], ['L', 0.086, 6.749], ['C', 0.244, 6.631, 0.425, 6.52, 0.516, 6.464],
  ['C', 0.836, 6.266, 1.21, 6.105, 1.611, 5.986], ['C', 2.102, 5.842, 2.621, 5.766, 3.143, 5.747],
  ['C', 3.273, 5.742, 3.404, 5.741, 3.534, 5.743], ['Z']
];

// Slide 1 — three overlapping washes (tinted sky + salmon + green wave).
const S1_SKY = [
  ['M', 13.352, -0.003], ['L', 13.352, 1.743], ['L', 13.184, 1.766],
  ['C', 11.836, 1.931, 10.621, 1.79, 9.352, 1.657], ['C', 7.475, 1.459, 5.58, 1.097, 3.589, 1.312],
  ['C', 2.425, 1.438, 1.248, 1.736, 0.147, 2.207], ['L', -0.02, 2.281], ['L', -0.02, 0.014], ['L', 0.024, -0.003], ['Z']
];
const S1_SALMON = [
  ['M', 3.616, 4.744], ['C', 3.995, 4.739, 4.372, 4.765, 4.736, 4.816], ['C', 5.363, 4.904, 5.883, 5.099, 6.428, 5.298],
  ['C', 7.233, 5.59, 8.021, 5.951, 8.943, 6.093], ['C', 9.805, 6.225, 10.752, 6.185, 11.619, 5.931],
  ['C', 12.188, 5.765, 12.848, 5.457, 13.221, 5.075], ['L', 13.333, 4.932], ['L', 13.333, 6.757], ['L', 4.978, 7.5],
  ['L', 0.04, 7.5], ['L', -0.02, 6.193], ['C', -0.028, 6.021, 0.503, 5.674, 0.68, 5.556],
  ['C', 0.993, 5.348, 1.357, 5.176, 1.748, 5.045], ['C', 2.344, 4.846, 2.984, 4.753, 3.616, 4.744], ['Z']
];
const S1_GREEN = [
  ['M', 3.316, 4.619], ['C', 3.823, 4.63, 4.329, 4.697, 4.81, 4.809], ['C', 5.431, 4.952, 5.941, 5.193, 6.476, 5.439],
  ['C', 7.267, 5.803, 8.036, 6.233, 8.949, 6.455], ['C', 9.804, 6.664, 10.751, 6.708, 11.628, 6.532],
  ['C', 12.202, 6.416, 12.876, 6.168, 13.265, 5.82], ['L', 13.333, 5.744], ['L', 13.333, 7.508], ['L', 0.001, 7.508],
  ['L', 0.001, 5.761], ['C', 0.001, 5.588, 0.547, 5.289, 0.729, 5.187], ['C', 1.05, 5.008, 1.421, 4.868, 1.817, 4.772],
  ['C', 2.301, 4.655, 2.809, 4.609, 3.316, 4.619], ['Z']
];

// Slide 6 — single big sweep filling the lower two thirds.
const S6_SWEEP = [
  ['M', 0, 2.673], ['L', 0.38, 2.796], ['C', 2.025, 3.349, 3.497, 4.143, 5, 4.835],
  ['C', 6.252, 5.411, 7.445, 5.974, 8.901, 6.31], ['C', 10.025, 6.57, 11.21, 6.728, 12.397, 6.753],
  ['C', 12.546, 6.756, 12.694, 6.757, 12.842, 6.756], ['C', 12.991, 6.755, 13.139, 6.752, 13.287, 6.746],
  ['L', 13.333, 6.744], ['L', 13.333, 7.5], ['L', 0, 7.5], ['Z']
];

// Slide 7 — full bleed diagonal wash.
const S7_WASH = [
  ['M', 0, 0], ['C', 4.336, 8.317, 7.795, 1.557, 13.333, 5.529], ['L', 13.333, 7.5], ['L', 0, 7.5], ['L', 0, 0], ['Z']
];

// Slide 11 — taller version of the title wave pair.
const S11_SALMON = [
  ['M', 1.787, 3.078], ['C', 2.396, 3.07, 3.001, 3.111, 3.584, 3.193], ['C', 4.59, 3.334, 5.425, 3.647, 6.299, 3.965],
  ['C', 7.592, 4.435, 8.857, 5.014, 10.335, 5.241], ['C', 11.2, 5.374, 12.119, 5.398, 13.025, 5.299],
  ['L', 13.333, 5.256], ['L', 13.333, 6.667], ['L', 3.973, 7.5], ['L', 0, 7.5], ['L', 0, 3.251], ['L', 0.266, 3.204],
  ['C', 0.769, 3.125, 1.279, 3.084, 1.787, 3.078], ['Z']
];
const S11_GREEN = [
  ['M', 1, 2.875], ['C', 1.102, 2.874, 1.203, 2.875, 1.305, 2.877], ['C', 2.119, 2.894, 2.932, 3.002, 3.703, 3.18],
  ['C', 4.7, 3.411, 5.519, 3.797, 6.376, 4.192], ['C', 7.645, 4.776, 8.881, 5.466, 10.346, 5.824],
  ['C', 11.203, 6.033, 12.118, 6.139, 13.028, 6.12], ['L', 13.333, 6.105], ['L', 13.333, 7.513], ['L', 0, 7.513],
  ['L', 0, 2.933], ['L', 0.088, 2.923], ['C', 0.39, 2.893, 0.695, 2.877, 1, 2.875], ['Z']
];

// Slide 12 — rotated wedges (salmon behind, green in front).
const S12_SALMON = [
  ['M', -0.012, 2.903], ['L', 0.325, 2.974], ['C', 1.24, 3.179, 2.133, 3.457, 2.983, 3.789],
  ['C', 4.448, 4.363, 5.606, 5.133, 6.822, 5.924], ['C', 7.497, 6.363, 8.161, 6.824, 8.841, 7.275],
  ['L', 9.207, 7.512], ['L', -0.012, 7.512], ['Z']
];
const S12_GREEN = [
  ['M', 0, 0], ['L', 5.22, 0], ['L', 5.216, 0.113], ['C', 5.177, 0.915, 5.062, 1.712, 4.886, 2.473],
  ['C', 4.626, 3.598, 4.19, 4.521, 3.745, 5.488], ['C', 3.498, 6.025, 3.235, 6.556, 2.98, 7.098],
  ['L', 2.795, 7.5], ['L', 0, 7.5], ['Z']
];

// Slide 23 — the line-chart spline (stroke only, no fill).
const S23_SPLINE = [
  ['M', 0, 1.973], ['C', 0.515, 2.903, 0.913, 3.926, 1.749, 3.913], ['C', 2.586, 3.901, 3.295, 2.469, 3.907, 2.405],
  ['C', 4.518, 2.342, 4.62, 2.955, 5.109, 2.901], ['C', 5.598, 2.846, 5.912, 1.889, 6.379, 1.762]
];

/* ---------------------------------------------------------------- utilities */

/** Bounding box of an absolute path, used to place the custGeom frame. */
function bbox(cmds) {
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  for (const c of cmds) {
    for (let i = 1; i < c.length; i += 2) {
      x0 = Math.min(x0, c[i]); x1 = Math.max(x1, c[i]);
      y0 = Math.min(y0, c[i + 1]); y1 = Math.max(y1, c[i + 1]);
    }
  }
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
}

/** Absolute path commands -> pptxgenjs custGeom `points` (relative to the frame). */
function toPoints(cmds, box) {
  const rx = v => +(v - box.x).toFixed(4);
  const ry = v => +(v - box.y).toFixed(4);
  return cmds.map((c, i) => {
    if (c[0] === 'Z') return { close: true };
    if (c[0] === 'C') {
      return {
        x: rx(c[5]), y: ry(c[6]),
        curve: { type: 'cubic', x1: rx(c[1]), y1: ry(c[2]), x2: rx(c[3]), y2: ry(c[4]) }
      };
    }
    return { x: rx(c[1]), y: ry(c[2]), moveTo: c[0] === 'M' || i === 0 };
  });
}

/**
 * PowerPoint linear gradients are not exposed by pptxgenjs, so every gradient
 * freeform is rendered as a stack of solid-colour layers: the outline is
 * flattened to a polygon and then re-drawn `BANDS` times, each copy clipped a
 * little further along the gradient axis. Because the layers are nested rather
 * than butted together, no hairline seams appear between them.
 */
const BANDS = 44;
const FLATTEN = 14; // line segments per cubic bezier

function mixHex(a, b, t) {
  const ch = i => {
    const va = parseInt(a.substr(i * 2, 2), 16);
    const vb = parseInt(b.substr(i * 2, 2), 16);
    return Math.round(va + (vb - va) * t).toString(16).padStart(2, '0').toUpperCase();
  };
  return ch(0) + ch(1) + ch(2);
}

/** Sample a gradient defined as [[pos%, hex, alpha%], ...] at position `p` (0..100). */
function sampleStops(stops, p) {
  if (p <= stops[0][0]) return { color: stops[0][1], alpha: stops[0][2] / 100 };
  const last = stops[stops.length - 1];
  if (p >= last[0]) return { color: last[1], alpha: last[2] / 100 };
  for (let i = 1; i < stops.length; i++) {
    if (p <= stops[i][0]) {
      const a = stops[i - 1], b = stops[i];
      const t = (p - a[0]) / (b[0] - a[0]);
      return { color: mixHex(a[1], b[1], t), alpha: (a[2] + (b[2] - a[2]) * t) / 100 };
    }
  }
  return { color: last[1], alpha: last[2] / 100 };
}

/** Path commands -> flat polygon [[x,y], ...]. */
function flatten(cmds) {
  const out = [];
  let cx = 0, cy = 0;
  for (const c of cmds) {
    if (c[0] === 'Z') continue;
    if (c[0] === 'C') {
      for (let i = 1; i <= FLATTEN; i++) {
        const t = i / FLATTEN, u = 1 - t;
        out.push([
          u * u * u * cx + 3 * u * u * t * c[1] + 3 * u * t * t * c[3] + t * t * t * c[5],
          u * u * u * cy + 3 * u * u * t * c[2] + 3 * u * t * t * c[4] + t * t * t * c[6]
        ]);
      }
      cx = c[5]; cy = c[6];
    } else {
      out.push([c[1], c[2]]);
      cx = c[1]; cy = c[2];
    }
  }
  return simplify(out);
}

/** Ramer-Douglas-Peucker: drop points that stay within `eps` of the chord. */
function simplify(pts, eps = 0.004) {
  if (pts.length < 3) return pts;
  const first = pts[0], last = pts[pts.length - 1];
  const dx = last[0] - first[0], dy = last[1] - first[1];
  const len = Math.hypot(dx, dy);
  let far = 0, farAt = 0;
  for (let i = 1; i < pts.length - 1; i++) {
    const d = len < 1e-9
      ? Math.hypot(pts[i][0] - first[0], pts[i][1] - first[1])
      : Math.abs(dx * (first[1] - pts[i][1]) - dy * (first[0] - pts[i][0])) / len;
    if (d > far) { far = d; farAt = i; }
  }
  if (far <= eps) return [first, last];
  return simplify(pts.slice(0, farAt + 1), eps).slice(0, -1).concat(simplify(pts.slice(farAt), eps));
}

/** Sutherland-Hodgman: keep the part of `poly` where dot(p, axis) >= limit. */
function halfPlane(poly, ax, ay, limit) {
  const dist = p => p[0] * ax + p[1] * ay - limit;
  const out = [];
  for (let i = 0; i < poly.length; i++) {
    const a = poly[i], b = poly[(i + 1) % poly.length];
    const da = dist(a), db = dist(b);
    if (da >= 0) out.push(a);
    if ((da >= 0) !== (db >= 0)) {
      const t = da / (da - db);
      out.push([a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t]);
    }
  }
  return out;
}

/**
 * Draw a freeform with a linear gradient fill.
 *  stops  [[pos%, hex, alpha%], ...] measured along the gradient axis
 *  angle  axis direction in degrees (0 = left to right, 90 = top to bottom)
 */
function gradientShape(slide, cmds, stops, angle) {
  const poly = flatten(cmds);
  const rad = (angle * Math.PI) / 180;
  let ax = Math.cos(rad), ay = Math.sin(rad);

  const samples = [];
  for (let i = 0; i < BANDS; i++) samples.push(sampleStops(stops, ((i + 0.5) / BANDS) * 100));
  // Layers are stacked opaque-last, so walk the axis in the direction of rising alpha.
  const rising = samples[samples.length - 1].alpha >= samples[0].alpha;
  if (!rising) { samples.reverse(); ax = -ax; ay = -ay; }

  const proj = poly.map(p => p[0] * ax + p[1] * ay);
  const lo = Math.min(...proj), hi = Math.max(...proj);

  // Track what is already on the slide so each new layer lands on the target colour.
  let accColor = null, accAlpha = 0;
  for (let i = 0; i < BANDS; i++) {
    const tgt = samples[i];
    if (tgt.alpha <= 0.005) continue;
    let a, color;
    if (accAlpha >= 0.999) { a = 1; color = tgt.color; }
    else {
      a = (tgt.alpha - accAlpha) / (1 - accAlpha);
      if (a <= 0.004) continue;
      // Un-composite: what colour must this layer be for the stack to read as `tgt`?
      const mixed = i =>
        Math.max(0, Math.min(255, Math.round(
          (parseInt(tgt.color.substr(i * 2, 2), 16) * tgt.alpha -
            parseInt((accColor || tgt.color).substr(i * 2, 2), 16) * accAlpha * (1 - a)) / a)));
      color = [0, 1, 2].map(i => mixed(i).toString(16).padStart(2, '0').toUpperCase()).join('');
    }
    const region = i === 0 ? poly : halfPlane(poly, ax, ay, lo + ((hi - lo) * i) / BANDS);
    if (region.length >= 3) solidPolygon(slide, region, color, Math.round((1 - a) * 100));
    accColor = tgt.color;
    accAlpha = tgt.alpha;
  }
}

/** Add a closed polygon (absolute inches) as a solid custGeom shape. */
function solidPolygon(slide, poly, color, transparency) {
  const box = bbox(poly.map(p => ['L', p[0], p[1]]));
  box.w = Math.max(box.w, 0.001);
  box.h = Math.max(box.h, 0.001);
  slide.addShape('custGeom', {
    x: box.x, y: box.y, w: box.w, h: box.h,
    points: poly.map((p, i) => ({ x: +(p[0] - box.x).toFixed(3), y: +(p[1] - box.y).toFixed(3), moveTo: i === 0 }))
      .concat([{ close: true }]),
    fill: { color, transparency }, line: { type: 'none' }
  });
}

/* ------------------------------------------------------- shared slide chrome */

/** Master decoration: two stacked waves plus the footer strap and page number. */
function addChrome(slide, pageNo) {
  gradientShape(slide, WAVE_SALMON, [[0, C.orange, 100], [100, C.orange, 0]], 250);
  gradientShape(slide, WAVE_GREEN, [[24, C.green, 100], [100, C.mint, 0]], 356.76);
  slide.addText('2024 © Digital Marketing', {
    x: 0.425, y: 7.014, w: 1.983, h: 0.252, fontSize: 9, fontFace: HEAD, color: C.ink, align: 'left'
  });
  slide.addText(String(pageNo), {
    x: 11.438, y: 6.888, w: 1.666, h: 0.505, fontSize: 24, fontFace: HEAD,
    color: C.ink, transparency: 60, align: 'right'
  });
}

/** Flat stand-in for a bitmap from the source deck. */
function imagePlaceholder(slide, x, y, w, h, opts = {}) {
  slide.addText(opts.label === undefined ? '[image]' : opts.label, {
    shape: opts.shape || 'rect', x, y, w, h,
    fill: { color: opts.color || C.placeholder, transparency: opts.transparency === undefined ? 30 : opts.transparency },
    line: { type: 'none' }, align: 'center', valign: 'middle', wrap: opts.wrap !== false,
    fontSize: opts.fontSize || 14, bold: true, fontFace: BODY, color: opts.textColor || C.white,
    rectRadius: opts.rectRadius
  });
}

/**
 * Heading/label text box. Defaults match the deck: Lato, near-black, and
 * top-anchored (PowerPoint's default, whereas pptxgenjs centres vertically).
 */
function text(slide, content, o) {
  slide.addText(content, Object.assign({ fontFace: HEAD, color: C.ink, valign: 'top' }, o));
}

/** Body copy: Open Sans 12pt in the muted body grey. */
function txt(slide, content, o) {
  text(slide, content, Object.assign({ fontFace: BODY, color: C.body, fontSize: 12 }, o));
}

// Repeated heading; kept as two runs on one line exactly as the source deck stores it.
// Returns a fresh array because pptxgenjs writes resolved options back into the run objects.
const channelsHeading = () => [{ text: 'Digital ' }, { text: 'MarketingChannels' }];

/** Paragraph list -> pptxgenjs run array (one run per line). */
function paras(lines, o = {}) {
  return lines.map(t => ({ text: t, options: Object.assign({ breakLine: true }, o) }));
}

/* -------------------------------------------------------------------- charts */

/** 75%-hole ring used for the "63% / 55%" stats; `values` are the raw series. */
function donut(slide, x, y, color, values) {
  slide.addChart('doughnut', [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values }], {
    x, y, w: 1.637, h: 2.026, holeSize: 75, firstSliceAng: 0,
    chartColors: [C.paper, color], dataBorder: { pt: 0, color: C.white },
    showLegend: false, showTitle: false, showValue: false, showLabel: false
  });
}

/** Four-column bar chart with a faint value grid (slide 22). */
function barChart(slide, x, y, color, values) {
  slide.addChart('bar', [{ name: 'Series 1', labels: ['Pro A', 'Pro B', 'Pro C', 'Pro D'], values }], {
    x, y, w: 3.928, h: 2.817, barDir: 'col', barGapWidthPct: 219, barOverlapPct: -27,
    chartColors: [color], showLegend: false, showTitle: false, showValue: false,
    valAxisHidden: true, valGridLine: { style: 'solid', color: 'BFBFBF', size: 0.75 },
    catGridLine: { style: 'none' }, catAxisLineShow: true, catAxisLineColor: 'D9D9D9',
    catAxisLabelColor: C.body, catAxisLabelFontFace: BODY, catAxisLabelFontSize: 12
  });
}

/* ---------------------------------------------------------- slide builders */

function slide01(p) {
  const s = p.addSlide();
  gradientShape(s, S1_SKY, [[0, C.green, 0], [100, C.mint, 59]], 158.23);
  gradientShape(s, S1_SALMON, [[0, C.orange, 100], [100, C.orange, 0]], 250);
  gradientShape(s, S1_GREEN, [[24, C.green, 100], [100, C.mint, 100]], 0);
  text(s, 'Digital Marketing', {
    x: 1.921, y: 2.572, w: 9.492, h: 1.178, fontSize: 80, bold: true, fontFace: HEAD,
    color: C.green, align: 'center', lineSpacingMultiple: 0.8
  });
  text(s, 'Presentation Template by Budogol', {
    x: 3.048, y: 3.707, w: 7.238, h: 0.438, fontSize: 20, fontFace: HEAD, color: C.ink, align: 'center'
  });
  text(s, '©2024 Digital Marketing', {
    x: 4.763, y: 6.988, w: 3.807, h: 0.303, fontSize: 12, fontFace: HEAD, color: C.ink, align: 'center'
  });
}

function slide02(p) {
  const s = p.addSlide();
  addChrome(s, 2);
  text(s, 'Our Introduction', {
    x: 1.443, y: 3.55, w: 4.59, h: 1.555, fontSize: 54, fontFace: HEAD, color: C.ink, lineSpacingMultiple: 0.8
  });
  text(s, 'Digital Marketing Opinion', {
    x: 7.3, y: 3.959, w: 4.979, h: 0.452, fontSize: 16, fontFace: HEAD, color: C.orange, lineSpacingMultiple: 1.5
  });
  txt(s, LOREM.long, { x: 7.3, y: 4.539, w: 4.979, h: 1.28, lineSpacingMultiple: 1.5 });
}

function slide03(p) {
  const s = p.addSlide();
  s.addShape('rect', { x: 0, y: 2.816, w: 13.2, h: 2.662, fill: { color: C.paper, transparency: 65 }, line: { type: 'none' } });
  addChrome(s, 3);
  text(s, 'Introduction Digital Marketing', {
    x: 0.946, y: 1.004, w: 8.387, h: 1.555, fontSize: 54, fontFace: HEAD, color: C.ink, lineSpacingMultiple: 0.8
  });
  text(s, '2024', {
    x: 9.776, y: 1.026, w: 2.521, h: 0.382, fontSize: 14, fontFace: HEAD, color: C.green,
    align: 'right', lineSpacingMultiple: 1.3
  });
  text(s, 'Digital Marketing', {
    x: 10.359, y: 1.327, w: 1.938, h: 0.382, fontSize: 14, fontFace: HEAD, color: C.body,
    align: 'right', lineSpacingMultiple: 1.3
  });
  text(s, 'Analytics and Data-driven Marketing', {
    x: 1.036, y: 3.287, w: 7.901, h: 0.463, fontSize: 16, fontFace: HEAD, color: C.orange, lineSpacingMultiple: 1.5
  });
  txt(s, paras([
    LOREM.long,
    'Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus. Pellentesque habitant morbi tristique senectus',
    'et netus et malesuada fames ac turpis egestas. Proin pharetra nonummy pede. Mauris et orci.'
  ]), { x: 1.036, y: 3.822, w: 11.175, h: 1.28, lineSpacingMultiple: 1.5 });
}

function slide04(p) {
  const s = p.addSlide();
  addChrome(s, 4);
  text(s, 'Affiliate marketing is a performance-based', {
    x: 0.858, y: 1.889, w: 4.108, h: 2.471, fontSize: 44, fontFace: HEAD, color: C.ink, lineSpacingMultiple: 0.8
  });
  const rows = [
    { n: '01', color: C.green, badgeY: 1.278, numY: 1.466, headY: 1.102, bodyY: 1.568 },
    { n: '02', color: C.orange, badgeY: 2.728, numY: 2.915, headY: 2.551, bodyY: 3.017 },
    { n: '03', color: C.gold, badgeY: 4.217, numY: 4.404, headY: 4.04, bodyY: 4.506 }
  ];
  rows.forEach(r => {
    s.addShape('roundRect', {
      x: 6.388, y: r.badgeY, w: 0.779, h: 0.779, rectRadius: 0.256,
      fill: { color: r.color }, line: { type: 'none' }
    });
    text(s, r.n, {
      x: 6.388, y: r.numY, w: 0.779, h: 0.404, fontSize: 20, fontFace: HEAD, color: C.white,
      align: 'center', lineSpacingMultiple: 0.9
    });
    text(s, 'Marketing Content A', {
      x: 7.65, y: r.headY, w: 3.673, h: 0.37, fontSize: 16, fontFace: HEAD, color: r.color
    });
    txt(s, LOREM.medium, { x: 7.65, y: r.bodyY, w: 4.656, h: 0.604, lineSpacingMultiple: 1.3 });
  });
}

/** Donut + caption block reused on slides 5, 19 and 21. */
function marketPair(s, ox, oy) {
  const cards = [
    { color: C.green, pct: 63, values: [18, 25], label: 'Market A', dx: 0 },
    { color: C.orange, pct: 55, values: [25, 18], label: 'Market B', dx: 2.936 }
  ];
  cards.forEach(c => {
    donut(s, ox + c.dx, oy, c.color, c.values);
    text(s, [{ text: String(c.pct), options: { fontSize: 20 } }, { text: '%', options: { fontSize: 18 } }], {
      x: ox + c.dx + 0.172, y: oy + 0.795, w: 1.292, h: 0.438, fontFace: BODY, color: c.color,
      align: 'center', valign: 'middle'
    });
    text(s, c.label, {
      x: ox + c.dx - 0.264, y: oy + 1.964, w: 2.165, h: 0.343, fontSize: 16, fontFace: HEAD,
      color: c.color, align: 'center', lineSpacingMultiple: 0.9
    });
    txt(s, LOREM.short, {
      x: ox + c.dx - 0.269, y: oy + 2.276, w: 2.175, h: 0.866, align: 'center', lineSpacingMultiple: 1.3
    });
  });
}

function slide05(p) {
  const s = p.addSlide();
  addChrome(s, 5);
  text(s, 'Involves promoting products or services on social media ', {
    x: 1.764, y: 0.694, w: 9.806, h: 1.286, fontSize: 44, fontFace: HEAD, color: C.ink,
    align: 'center', lineSpacingMultiple: 0.8
  });
  txt(s, LOREM.fusce, { x: 1.083, y: 2.872, w: 5.401, h: 1.27, fontSize: 16, lineSpacingMultiple: 1.5 });
  text(s, '$ 2.145.000', { x: 1.083, y: 4.314, w: 2.99, h: 0.505, fontSize: 24, fontFace: HEAD, color: C.green });
  txt(s, LOREM.short, { x: 1.083, y: 4.813, w: 3.74, h: 0.604, lineSpacingMultiple: 1.3 });
  s.addShape('line', {
    x: 9.654, y: 2.52, w: 0, h: 3.223, line: { color: C.body, transparency: 75, width: 0.5 }
  });
  marketPair(s, 7.25, 2.455);
}

function slide06(p) {
  const s = p.addSlide();
  gradientShape(s, S6_SWEEP, [[24, C.green, 100], [100, C.orange, 100]], 0);
  text(s, 'How did we do it?', {
    x: 2.105, y: 1.196, w: 4.108, h: 1.393, fontSize: 48, fontFace: HEAD, color: C.ink, lineSpacingMultiple: 0.8
  });
  text(s, 'Content marketing focuses on creating and distributing valuable', {
    x: 7.12, y: 1.173, w: 3.849, h: 0.714, fontSize: 16, fontFace: HEAD, color: C.green, lineSpacingMultiple: 1.2
  });
  txt(s, paras([LOREM.long, '', LOREM.nunc, '', LOREM.pellen]), {
    x: 7.12, y: 2.102, w: 4.656, h: 2.967, lineSpacingMultiple: 1.3
  });
}

function slide07(p) {
  const s = p.addSlide();
  // Layout accents: three pale ticks that show through the wash.
  [[6.647, 0.214, 0.04, 0.945], [0.812, 1.597, 0.08, 1.109], [10.957, 6.724, 0.05, 0.776]].forEach(r => {
    s.addShape('rect', { x: r[0], y: r[1], w: r[2], h: r[3], fill: { color: C.paper, transparency: 80 }, line: { type: 'none' } });
  });
  gradientShape(s, S7_WASH, [[24, C.green, 86], [100, C.orange, 50]], 0);
  text(s, 'Digital Marketing ©2024', {
    x: 1.758, y: 1.137, w: 4.06, h: 0.3, fontSize: 16, fontFace: BODY, color: C.white, lineSpacingMultiple: 0.7
  });
  text(s, paras(['Search Engine Marketing', '(SEM)']), {
    x: 2.642, y: 1.57, w: 8.499, h: 1.555, fontSize: 54, fontFace: HEAD, color: C.white, lineSpacingMultiple: 0.8
  });
  txt(s, paras([LOREM.long, '', LOREM.nunc, LOREM.pellen]), {
    x: 1.531, y: 5.0, w: 9.611, h: 1.917, color: C.white, lineSpacingMultiple: 1.3
  });
}

function slide08(p) {
  const s = p.addSlide();
  addChrome(s, 8);
  text(s, 'Digital Marketing ©2024', {
    x: 1.013, y: 0.633, w: 4.06, h: 0.3, fontSize: 16, fontFace: BODY, color: C.orange, lineSpacingMultiple: 0.7
  });
  text(s, 'This includes optimizing content', {
    x: 1.013, y: 1.173, w: 6.37, h: 1.393, fontSize: 48, fontFace: HEAD, color: C.ink, lineSpacingMultiple: 0.8
  });
  txt(s, paras([
    LOREM.long, '', LOREM.nunc, 'Pellentesque netus et malesuada pede. Mauris et orci.'
  ]), { x: 1.013, y: 2.957, w: 5.784, h: 2.219, fontSize: 14, lineSpacingMultiple: 1.3 });
}

function slide09(p) {
  const s = p.addSlide();
  addChrome(s, 9);
  const items = ['Promotional Strategy', 'Product Situation', 'Marketing Strategy Proposal', 'Competitive Situation', 'Action Programs'];
  items.forEach((label, i) => {
    const y = 1.162 + i * 0.818;
    text(s, String(i + 1).padStart(2, '0'), {
      x: 0.827, y, w: 0.873, h: 0.438, fontSize: 20, bold: true, fontFace: HEAD, color: C.green, align: 'center'
    });
    txt(s, label, { x: 1.7, y, w: 4.967, h: 0.438, fontSize: 20 });
  });
  s.addText('What We Do?', {
    shape: 'round1Rect', x: 6.784, y: 3.488, w: 3.023, h: 1.788, rectRadius: 0.894,
    fill: { color: C.green }, line: { type: 'none' },
    fontSize: 24, fontFace: HEAD, color: C.white, align: 'center', valign: 'middle'
  });
}

function slide10(p) {
  const s = p.addSlide();
  addChrome(s, 10);
  const rings = [
    { x: 7.474, y: 0.924, d: 5.193, color: C.green, tag: 'TAM', tagX: 9.795, tagY: 1.178, tagW: 0.553, val: '$112 B', valX: 9.348, valY: 1.473, valW: 1.445 },
    { x: 8.168, y: 2.146, d: 3.806, color: C.orange, tag: 'SAM', tagX: 9.795, tagY: 2.426, tagW: 0.553, val: '$53 B', valX: 9.462, valY: 2.721, valW: 1.217 },
    { x: 8.851, y: 3.293, d: 2.44, color: C.stone, tag: 'SOM', tagX: 9.782, tagY: 4.041, tagW: 0.577, val: '$21 B', valX: 9.462, valY: 4.412, valW: 1.217 }
  ];
  rings.forEach(r => s.addShape('ellipse', {
    x: r.x, y: r.y, w: r.d, h: r.d, fill: { color: r.color },
    line: { color: C.white, width: 2.5 },
    shadow: { type: 'outer', blur: 46, offset: 28, angle: 270, color: '000000', opacity: 0.07 }
  }));
  rings.forEach(r => {
    text(s, r.tag, { x: r.tagX, y: r.tagY, w: r.tagW, h: 0.303, fontSize: 12, fontFace: BODY, color: C.white, align: 'center' });
    text(s, r.val, { x: r.valX, y: r.valY, w: r.valW, h: 0.572, fontSize: 28, bold: true, fontFace: HEAD, color: C.white, align: 'center' });
  });
  txt(s, LOREM.long, { x: 1.194, y: 3.965, w: 4.979, h: 1.129, lineSpacingMultiple: 1.3 });
  text(s, 'Digital Marketing Opinion', {
    x: 1.194, y: 3.386, w: 4.979, h: 0.452, fontSize: 16, fontFace: HEAD, color: C.orange, lineSpacingMultiple: 1.5
  });
  text(s, 'Affiliate marketing is a performance-based', {
    x: 1.219, y: 0.851, w: 4.108, h: 2.471, fontSize: 44, fontFace: HEAD, color: C.ink, lineSpacingMultiple: 0.8
  });
}

function slide11(p) {
  const s = p.addSlide();
  gradientShape(s, S11_SALMON, [[0, C.orange, 100], [100, C.orange, 0]], 250);
  gradientShape(s, S11_GREEN, [[24, C.green, 100], [100, C.mint, 100]], 0);
  text(s, paras(['Digital Marketing', 'Proposal']), {
    x: 1.346, y: 0.903, w: 6.032, h: 1.393, fontSize: 48, fontFace: HEAD, color: C.ink, lineSpacingMultiple: 0.8
  });
  txt(s, 'Project Proposal', { x: 7.766, y: 0.935, w: 4.979, h: 0.372, lineSpacingMultiple: 1.5 });
  text(s, 'Comprehensive Digital Marketing Service', {
    x: 7.766, y: 1.273, w: 4.979, h: 0.868, fontSize: 20, fontFace: HEAD, color: C.green, lineSpacingMultiple: 1.2
  });
  txt(s, 'Client', { x: 7.766, y: 2.275, w: 4.979, h: 0.372, lineSpacingMultiple: 1.5 });
  text(s, 'Microhard Company Co.', {
    x: 7.766, y: 2.613, w: 4.979, h: 0.464, fontSize: 20, fontFace: HEAD, color: C.orange, lineSpacingMultiple: 1.2
  });
}

function slide12(p) {
  const s = p.addSlide();
  gradientShape(s, S12_SALMON, [[0, C.orange, 100], [100, C.orange, 0]], 258.35);
  gradientShape(s, S12_GREEN, [[24, C.green, 100], [100, C.mint, 34]], 90);
  text(s, channelsHeading(), {
    x: 0.651, y: 1.119, w: 3.679, h: 2.282, fontSize: 54, fontFace: HEAD, color: C.white, lineSpacingMultiple: 0.8
  });
  s.addShape('line', { x: 0.873, y: 3.75, w: 0, h: 2.12, flipV: true, line: { color: C.white, width: 1, beginArrowType: 'oval' } });
  txt(s, paras(['Date', '02/12/24']), { x: 0.97, y: 5.185, w: 1.64, h: 0.604, color: C.white, lineSpacingMultiple: 1.3 });
  txt(s, LOREM.fusce, { x: 0.752, y: 6.219, w: 5.382, h: 0.604, color: C.white, lineSpacingMultiple: 1.3 });
}

function slide13(p) {
  const s = p.addSlide();
  addChrome(s, 13);
  const cards = [
    { x: 0.937, y: 2.116, badge: C.green, head: C.green, label: 'Digital Marketing A' },
    { x: 0.937, y: 3.821, badge: C.orange, head: C.orange, label: 'Digital Marketing B' },
    { x: 6.798, y: 2.116, badge: C.gold, head: C.gold, label: 'Digital Marketing C' },
    { x: 6.798, y: 3.821, badge: C.sage, head: C.stone, label: 'Digital Marketing D' }
  ];
  cards.forEach(c => {
    s.addShape('roundRect', {
      x: c.x, y: c.y, w: 5.598, h: 1.468, rectRadius: 0.22,
      fill: { color: C.white, transparency: 12 }, line: { type: 'none' },
      shadow: { type: 'outer', blur: 65, offset: 3, angle: 90, color: '000000', opacity: 0.12 }
    });
    s.addShape('roundRect', {
      x: c.x + 0.24, y: c.y + 0.223, w: 0.96, h: 0.96, rectRadius: 0.3,
      fill: { color: c.badge }, line: { type: 'none' }
    });
    // icon glyph from the source deck -> simple white outline mark
    s.addShape('roundRect', {
      x: c.x + 0.46, y: c.y + 0.5, w: 0.52, h: 0.41, rectRadius: 0.07,
      fill: { type: 'none' }, line: { color: C.white, width: 1.25 }
    });
    text(s, c.label, {
      x: c.x + 1.36, y: c.y + 0.245, w: 3.92, h: 0.371, fontSize: 16, fontFace: HEAD,
      color: c.head, lineSpacingMultiple: 1.1, paraSpaceAfter: 6
    });
    txt(s, LOREM.medium, {
      x: c.x + 1.36, y: c.y + 0.619, w: 4.068, h: 0.604, lineSpacingMultiple: 1.3, paraSpaceAfter: 6
    });
  });
  text(s, 'This includes optimizing content', {
    x: 1.332, y: 0.953, w: 10.669, h: 0.693, fontSize: 44, fontFace: HEAD, color: C.ink,
    align: 'center', lineSpacingMultiple: 0.8
  });
}

function slide14(p) {
  const s = p.addSlide();
  addChrome(s, 14);
  text(s, 'Affiliate marketing is a performance-based', {
    x: 4.602, y: 1.079, w: 7.107, h: 2.282, fontSize: 54, fontFace: HEAD, color: C.ink, lineSpacingMultiple: 0.8
  });
  [{ n: '01', x: 7.183, bx: 7.175 }, { n: '02', x: 10.182, bx: 10.173 }].forEach(col => {
    text(s, col.n, {
      x: col.x, y: 4.139, w: 2.178, h: 0.404, fontSize: 20, bold: true, fontFace: HEAD,
      color: C.green, lineSpacingMultiple: 0.9
    });
    txt(s, LOREM.short, { x: col.bx, y: 4.543, w: 2.62, h: 0.993, fontSize: 14, color: C.grey, lineSpacingMultiple: 1.3 });
  });
}

function slide15(p) {
  const s = p.addSlide();
  addChrome(s, 15);
  const steps = [
    { cardX: 1.227, cardY: 0.689, dotX: 2.347, dotY: 1.909, color: C.green, n: '01', headY: 0.919, bodyY: 1.259, textX: 1.587 },
    { cardX: 3.147, cardY: 3.75, dotX: 4.267, dotY: 3.27, color: C.orange, n: '02', headY: 4.34, bodyY: 4.679, textX: 3.507 },
    { cardX: 5.067, cardY: 0.689, dotX: 6.187, dotY: 1.909, color: C.stone, n: '03', headY: 0.919, bodyY: 1.259, textX: 5.427 },
    { cardX: 6.987, cardY: 3.75, dotX: 8.107, dotY: 3.27, color: C.gold, n: '04', headY: 4.34, bodyY: 4.679, textX: 7.347 },
    { cardX: 8.906, cardY: 0.689, dotX: 10.026, dotY: 1.909, color: C.sage, n: '05', headY: 0.919, bodyY: 1.259, textX: 9.266 }
  ];
  steps.forEach(st => {
    s.addShape('roundRect', {
      x: st.cardX, y: st.cardY, w: 3.177, h: 1.7, rectRadius: 0.19,
      fill: { color: C.white, transparency: 12 }, line: { type: 'none' },
      shadow: { type: 'outer', blur: 65, offset: 3, angle: 90, color: '000000', opacity: 0.12 }
    });
    s.addText(st.n, {
      shape: 'ellipse', x: st.dotX, y: st.dotY, w: 0.953, h: 0.96,
      fill: { color: st.color }, line: { type: 'none' },
      shadow: { type: 'outer', blur: 43, offset: 3, angle: 90, color: '000000', opacity: 0.07 },
      fontSize: 24, fontFace: HEAD, color: C.white, align: 'center', valign: 'middle'
    });
  });
  [{ x: 3.16, flipV: false }, { x: 5.08, flipV: true }, { x: 7.0, flipV: false }, { x: 8.92, flipV: true }].forEach(l => {
    s.addShape('line', {
      x: l.x, y: 2.729, w: 1.246, h: 0.682, flipV: l.flipV,
      line: { color: C.body, transparency: 70, width: 1.25, endArrowType: 'triangle' }
    });
  });
  steps.forEach(st => {
    text(s, 'Point Here', {
      x: st.textX, y: st.headY, w: 2.462, h: 0.346, fontSize: 14, fontFace: HEAD, color: C.ink,
      align: 'center', lineSpacingMultiple: 1.1, paraSpaceAfter: 6
    });
    txt(s, LOREM.card, {
      x: st.textX, y: st.bodyY, w: 2.462, h: 0.604, color: C.grey, align: 'center',
      lineSpacingMultiple: 1.3, paraSpaceAfter: 6
    });
  });
}

function slide16(p) {
  const s = p.addSlide();
  addChrome(s, 16);
  text(s, 'Promotional strategy', {
    x: 1.91, y: 0.678, w: 9.513, h: 0.747, fontSize: 48, fontFace: HEAD, color: C.ink,
    align: 'center', lineSpacingMultiple: 0.8
  });
  const groups = [
    { n: '01', title: 'Coupons', numX: 1.468, numY: 2.285, textX: 2.508, headY: 2.229, bodyY: 2.584 },
    { n: '02', title: 'Referral Programs', numX: 6.269, numY: 2.285, textX: 7.31, headY: 2.229, bodyY: 2.584 },
    { n: '03', title: 'Discount', numX: 2.869, numY: 4.062, textX: 3.909, headY: 4.007, bodyY: 4.361 },
    { n: '04', title: 'Loyalty Incentives', numX: 7.879, numY: 4.062, textX: 8.919, headY: 4.007, bodyY: 4.361 }
  ];
  groups.forEach(g => {
    text(s, g.n, {
      x: g.numX, y: g.numY, w: 0.879, h: 0.527, fontSize: 36, bold: true, fontFace: HEAD,
      color: C.green, lineSpacingMultiple: 0.7
    });
    text(s, g.title, {
      x: g.textX, y: g.headY, w: 2.841, h: 0.356, fontSize: 20, fontFace: HEAD, color: C.green, lineSpacingMultiple: 0.7
    });
    txt(s, 'Lorem ipsum dolor sit amet, consectetuer', { x: g.textX, y: g.bodyY, w: 3.031, h: 0.572, fontSize: 14 });
  });
}

function slide17(p) {
  const s = p.addSlide();
  addChrome(s, 17);
  gradientShape(s, [['M', 7.052, 0], ['L', 13.333, 0], ['L', 13.333, 6.367], ['L', 7.052, 6.367], ['Z']],
    [[0, C.green, 100], [100, C.green, 0]], 90);
  text(s, channelsHeading(), {
    x: 0.651, y: 0.667, w: 3.679, h: 2.282, fontSize: 54, fontFace: HEAD, color: C.ink, lineSpacingMultiple: 0.8
  });
  s.addShape('line', { x: 0.873, y: 3.299, w: 0, h: 0.959, flipV: true, line: { color: C.green, width: 1, beginArrowType: 'oval' } });
  txt(s, paras(['Date', '02/12/24']), { x: 0.97, y: 3.824, w: 1.64, h: 0.604, lineSpacingMultiple: 1.3 });
  txt(s, LOREM.fusce, { x: 0.752, y: 4.607, w: 5.382, h: 0.604, lineSpacingMultiple: 1.3 });

  text(s, 'Client & Investor', {
    x: 7.883, y: 0.667, w: 4.501, h: 0.404, fontSize: 18, fontFace: HEAD, color: C.ink, valign: 'top'
  });
  txt(s, LOREM.magna, { x: 7.883, y: 1.105, w: 4.501, h: 0.866, valign: 'top', lineSpacingMultiple: 1.3 });
  s.addShape('line', { x: 7.883, y: 2.27, w: 4.109, h: 0, line: { color: C.green, width: 0.75 } });
  text(s, 'Project Collaboration', {
    x: 7.883, y: 2.684, w: 4.501, h: 0.404, fontSize: 18, fontFace: HEAD, color: C.ink, valign: 'top'
  });
  txt(s, LOREM.magna, { x: 7.883, y: 3.122, w: 4.501, h: 0.869, valign: 'top', lineSpacingMultiple: 1.3 });
  [{ v: '350+', l: 'Material Rating', x: 7.883, lw: 1.525 }, { v: '250+', l: 'Property Rating', x: 9.794, lw: 2.061 }].forEach(k => {
    text(s, k.v, {
      x: k.x, y: 4.219, w: 1.367, h: 0.722, fontSize: 32, fontFace: HEAD, color: C.ink, lineSpacingMultiple: 1.3
    });
    txt(s, k.l, { x: k.x, y: 4.94, w: k.lw, h: 0.342, lineSpacingMultiple: 1.3 });
  });
}

function slide18(p) {
  const s = p.addSlide();
  addChrome(s, 18);
  text(s, 'Missionary Business Marketing Sales', {
    x: 0.98, y: 0.531, w: 10.702, h: 0.774, fontSize: 40, fontFace: HEAD, color: C.ink
  });
  text(s, paras(['Project', 'Collaboration']), {
    x: 0.978, y: 3.096, w: 2.609, h: 0.751, fontSize: 18, fontFace: HEAD, color: C.green,
    valign: 'top', lineSpacingMultiple: 1.1
  });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa.', {
    x: 0.978, y: 3.868, w: 2.857, h: 1.3, fontSize: 14, valign: 'top', lineSpacingMultiple: 1.3, paraSpaceAfter: 6
  });
  [{ x: 4.263, d: '14 Aug 2024' }, { x: 7.333, d: '18 Aug 2024' }, { x: 10.333, d: '26 Sep 2024' }].forEach(col => {
    text(s, col.d, {
      x: col.x, y: 1.894, w: 2.124, h: 0.459, fontSize: 18, fontFace: HEAD, color: C.green, lineSpacingMultiple: 1.3
    });
    txt(s, 'Lorem ipsum dolor sit adipiscing elit.', {
      x: col.x, y: 2.299, w: 2.254, h: 0.688, fontSize: 14, valign: 'top', lineSpacingMultiple: 1.3, paraSpaceAfter: 6
    });
  });
}

function slide19(p) {
  const s = p.addSlide();
  addChrome(s, 19);
  s.addShape('line', { x: 3.636, y: 0.773, w: 0, h: 3.223, line: { color: C.body, transparency: 75, width: 0.5 } });
  marketPair(s, 1.233, 0.708);
  txt(s, 'Project Proposal', { x: 6.667, y: 3.996, w: 4.979, h: 0.372, lineSpacingMultiple: 1.5 });
  text(s, 'Comprehensive Digital Marketing Service', {
    x: 6.667, y: 4.334, w: 4.979, h: 0.868, fontSize: 20, fontFace: HEAD, color: C.green, lineSpacingMultiple: 1.2
  });
}

function slide20(p) {
  const s = p.addSlide();
  addChrome(s, 20);
  text(s, 'Best Team Marketing', {
    x: 2.778, y: 0.803, w: 7.778, h: 0.693, fontSize: 44, fontFace: HEAD, color: C.ink,
    align: 'center', lineSpacingMultiple: 0.8
  });
  const team = [
    { photoX: 1.661, name: 'Milly Whiteley', nameX: 1.526, nameW: 1.849, role: 'Data Scientist', roleX: 1.387, roleW: 2.127, roleColor: C.orange, bodyX: 1.387 },
    { photoX: 4.472, name: 'Willem Milne', nameX: 4.337, nameW: 1.849, role: 'Data Scientist', roleX: 4.198, roleW: 2.127, roleColor: C.orange, bodyX: 4.198 },
    { photoX: 7.283, name: 'Kiah Burrows', nameX: 6.881, nameW: 2.383, role: 'Developer', roleX: 6.789, roleW: 2.567, roleColor: C.green, bodyX: 7.009 },
    { photoX: 10.094, name: 'Malia Mays', nameX: 9.959, nameW: 1.849, role: 'Content Marketer', roleX: 9.664, roleW: 2.567, roleColor: C.orange, bodyX: 9.82 }
  ];
  team.forEach(m => {
    imagePlaceholder(s, m.photoX, 2.034, 1.578, 1.578, {
      label: 'IMAGE PLACEHOLDER', transparency: 70, wrap: false, fontSize: 12, textColor: C.placeholder
    });
    text(s, m.name, {
      x: m.nameX, y: 3.957, w: m.nameW, h: 0.404, fontSize: 18, fontFace: HEAD, color: C.ink, align: 'center'
    });
    text(s, m.role, {
      x: m.roleX, y: 4.336, w: m.roleW, h: 0.303, fontSize: 12, fontFace: HEAD, color: m.roleColor, align: 'center'
    });
    txt(s, LOREM.tiny, { x: m.bodyX, y: 4.673, w: 2.127, h: 0.604, align: 'center', lineSpacingMultiple: 1.3 });
  });
}

function slide21(p) {
  const s = p.addSlide();
  imagePlaceholder(s, 0, 0, 7.518, 7.5, { color: 'EDEDED', transparency: 0, textColor: C.stone });
  addChrome(s, 21);
  s.addShape('line', { x: 10.283, y: 3.115, w: 0, h: 3.223, line: { color: C.body, transparency: 75, width: 0.5 } });
  marketPair(s, 7.879, 3.05);
  text(s, channelsHeading(), {
    x: 7.723, y: 0.667, w: 4.103, h: 2.282, fontSize: 54, fontFace: HEAD, color: C.ink, lineSpacingMultiple: 0.8
  });
}

function slide22(p) {
  const s = p.addSlide();
  addChrome(s, 22);
  text(s, 'Digital Marketing Chart Data', {
    x: 2.473, y: 0.648, w: 8.387, h: 0.774, fontSize: 40, fontFace: HEAD, color: C.ink, align: 'center'
  });
  barChart(s, 1.773, 1.769, C.green, [4.3, 2.5, 3.5, 4.5]);
  barChart(s, 7.587, 1.842, C.orange, [4.3, 3, 2, 5]);
  const stats = [
    { x: 1.799, pct: '50', color: C.green, when: 'September 2023', bodyX: 3.737, bodyY: 4.703 },
    { x: 7.646, pct: '80', color: C.orange, when: 'August 2024', bodyX: 9.116, bodyY: 4.689 }
  ];
  stats.forEach(st => {
    text(s, [{ text: st.pct, options: { fontSize: 40 } }, { text: '%', options: { fontSize: 28 } }], {
      x: st.x, y: 4.604, w: 1.498, h: 0.774, fontFace: HEAD, color: st.color
    });
    text(s, st.when, {
      x: st.x, y: 5.287, w: 1.74, h: 0.303, fontSize: 12, fontFace: HEAD, color: st.color, valign: 'top'
    });
    txt(s, LOREM.tiny, { x: st.bodyX, y: st.bodyY, w: 2.127, h: 0.604, lineSpacingMultiple: 1.3 });
  });
  s.addShape('line', { x: 6.667, y: 1.838, w: 0, h: 3.926, line: { color: C.body, transparency: 75, width: 0.5 } });
}

function slide23(p) {
  const s = p.addSlide();
  addChrome(s, 23);
  ['19 Aug 2024', '20 Aug 2024', '21 Aug 2024', '22 Aug 2024'].forEach((d, i) => {
    text(s, d, {
      x: 0.555 + i * 1.5677, y: 5.093, w: 1.121, h: 0.225, fontSize: 10.5, fontFace: HEAD,
      color: C.green, align: 'center', lineSpacingMultiple: 0.7
    });
  });
  [1.115, 2.683, 4.251, 5.818].forEach(x => {
    s.addShape('line', { x, y: 1.035, w: 0, h: 3.821, line: { color: C.paper, width: 1.5 } });
  });
  // Spline: stroke-only freeform.
  const box = bbox(S23_SPLINE);
  s.addShape('custGeom', {
    x: box.x, y: box.y, w: box.w, h: box.h, points: toPoints(S23_SPLINE, box),
    fill: { type: 'none' }, line: { color: C.green, width: 4.5 }
  });
  const callouts = [
    { label: 'Label A', boxX: 1.042, boxY: 4.079, boxW: 1.164, boxH: 0.442, tipX: 1.24, tipY: 3.958, txtX: 1.09, txtY: 4.132, dotX: 0.996, dotY: 3.552 },
    { label: 'Label B', boxX: 1.888, boxY: 2.009, boxW: 1.241, boxH: 0.519, tipX: 2.897, tipY: 2.523, txtX: 1.978, txtY: 2.1, dotX: 2.893, dotY: 2.844, tipFlipV: true },
    { label: 'Label C', boxX: 4.89, boxY: 3.242, boxW: 1.164, boxH: 0.442, tipX: 5.088, tipY: 3.121, txtX: 4.938, txtY: 3.296, dotX: 4.906, dotY: 2.724 }
  ];
  callouts.forEach(c => {
    s.addShape('roundRect', {
      x: c.boxX, y: c.boxY, w: c.boxW, h: c.boxH, rectRadius: 0,
      fill: { color: C.green }, line: { type: 'none' }
    });
    s.addShape('rtTriangle', {
      x: c.tipX, y: c.tipY, w: 0.143, h: 0.143, flipH: true, flipV: !!c.tipFlipV,
      fill: { color: C.green }, line: { type: 'none' }
    });
    text(s, c.label, {
      x: c.txtX, y: c.txtY, w: 1.062, h: 0.335, fontSize: 14, fontFace: HEAD, color: C.white, align: 'center'
    });
    s.addShape('ellipse', {
      x: c.dotX, y: c.dotY, w: 0.292, h: 0.292, fill: { color: C.green }, line: { color: C.white, width: 6 }
    });
  });
  text(s, 'Affiliate marketing is a performance-based', {
    x: 7.646, y: 1.082, w: 4.108, h: 2.471, fontSize: 44, fontFace: HEAD, color: C.ink, lineSpacingMultiple: 0.8
  });
  text(s, 'Digital Marketing Opinion', {
    x: 7.62, y: 3.616, w: 4.979, h: 0.452, fontSize: 16, fontFace: HEAD, color: C.orange, lineSpacingMultiple: 1.5
  });
  txt(s, LOREM.long, { x: 7.62, y: 4.196, w: 4.979, h: 1.129, lineSpacingMultiple: 1.3 });
}

function slide24(p) {
  const s = p.addSlide();
  addChrome(s, 24);
  s.addShape('roundRect', {
    x: 1.122, y: 1.831, w: 11.09, h: 4.815, rectRadius: 0.304,
    fill: { color: C.white, transparency: 12 }, line: { type: 'none' },
    shadow: { type: 'outer', blur: 65, offset: 3, angle: 90, color: '000000', opacity: 0.12 }
  });
  const bodyRows = [
    ['Coupons', 'Trade shows and conventions'],
    ['Sweepstakes or contests', 'Sales contests'],
    ['Premiums', 'Trade and advertising allowances'],
    ['Rebates', 'Product demonstrations'],
    ['Samples', 'Training'],
    ['Layalty programs', 'Free merchandise']
  ];
  const border = [0, 1, 2, 3].map(() => ({ type: 'solid', pt: 0.25, color: '090F21' }));
  const rows = [
    bodyRows.length >= 0
      ? ['Consumer Sales Promotions', 'Business-to-Business Sales Promotions'].map(t => ({
        text: t,
        options: { fill: { color: C.orange }, color: C.white, fontSize: 16, fontFace: HEAD, margin: 0.184, border }
      }))
      : []
  ].concat(bodyRows.map((cells, i) => {
    const shaded = i % 2 === 0;
    return cells.map(t => ({
      text: t,
      options: {
        fill: shaded ? { color: C.green } : { color: C.white, transparency: 100 },
        color: shaded ? C.white : '000000',
        fontSize: 14, fontFace: BODY, margin: 0.195, border
      }
    }));
  }));
  s.addTable(rows, {
    x: 1.408, y: 2.057, w: 10.517, colW: [5.259, 5.259],
    rowH: [0.631, 0.588, 0.588, 0.588, 0.588, 0.588, 0.588],
    valign: 'middle', align: 'left'
  });
  text(s, paras(['Search Engine Marketing', '(SEM)']), {
    x: 2.417, y: 0.56, w: 8.499, h: 1.286, fontSize: 44, fontFace: HEAD, color: C.ink,
    align: 'center', lineSpacingMultiple: 0.8
  });
}

function slide25(p) {
  const s = p.addSlide();
  addChrome(s, 25);
  text(s, 'Involves promoting products or services on social media ', {
    x: 1.764, y: 0.603, w: 9.806, h: 1.286, fontSize: 44, fontFace: HEAD, color: C.ink,
    align: 'center', lineSpacingMultiple: 0.8
  });
  // Each card holds a five-bar sparkline; `hi` marks the fully opaque bar.
  const panels = [
    {
      x: 1.309, valueW: 1.3, value: '5.321', delta: '22.45% ^', color: C.green,
      barX: 2.811, bars: [[3.143, 0.269], [3.04, 0.372], [2.984, 0.428], [2.651, 0.761], [2.791, 0.621]], hi: 3,
      labelX: 1.565, labelW: 2.933, label: 'Market Sales A', labelColor: C.green, bodyX: 1.934
    },
    {
      x: 4.945, valueW: 1.261, value: '3.121', delta: '15.45% ^', color: C.orange,
      barX: 6.447, bars: [[3.143, 0.269], [3.04, 0.372], [3.04, 0.372], [2.72, 0.692], [3.04, 0.372]], hi: 2,
      labelX: 5.408, labelW: 2.517, label: 'Market Sales B', labelColor: C.orange, bodyX: 5.57
    },
    {
      x: 8.58, valueW: 1.261, value: '2.211', delta: '23.00% ^', color: C.stone,
      barX: 10.082, bars: [[3.143, 0.269], [2.72, 0.692], [3.011, 0.401], [2.885, 0.527], [2.72, 0.692]], hi: 1,
      labelX: 8.836, labelW: 2.933, label: 'Market Sales C', labelColor: C.gold, bodyX: 9.206
    }
  ];
  panels.forEach(pn => {
    s.addShape('roundRect', {
      x: pn.x, y: 2.31, w: 3.444, h: 1.542, rectRadius: 0.2, fill: { color: C.white }, line: { type: 'none' },
      shadow: { type: 'outer', blur: 97, offset: 61, angle: 45, color: '000000', opacity: 0.15 }
    });
    txt(s, 'Option Here', { x: pn.x + 0.202, y: 2.539, w: 2.042, h: 0.381, fontSize: 14, color: C.ink, lineSpacingMultiple: 1.3 });
    text(s, pn.value, { x: pn.x + 0.202, y: 2.907, w: pn.valueW, h: 0.404, fontSize: 18, fontFace: HEAD, color: C.ink });
    text(s, pn.delta, { x: pn.x + 0.202, y: 3.301, w: 1.63, h: 0.286, fontSize: 11, fontFace: HEAD, color: pn.color });
    pn.bars.forEach((b, i) => {
      s.addShape('round2SameRect', {
        x: pn.barX + i * 0.3697, y: b[0], w: 0.208, h: b[1], rectRadius: 0.06,
        fill: { color: pn.color, transparency: i === pn.hi ? 0 : 50 }, line: { type: 'none' }
      });
    });
    text(s, pn.label, {
      x: pn.labelX, y: 4.168, w: pn.labelW, h: 0.508, fontSize: 18, fontFace: HEAD,
      color: pn.labelColor, align: 'center', lineSpacingMultiple: 1.5
    });
    txt(s, LOREM.tiny, { x: pn.bodyX, y: 4.69, w: 2.193, h: 0.604, align: 'center', lineSpacingMultiple: 1.3 });
  });
}

/* ---------------------------------------------------------------------- main */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE_16x9', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE_16x9';
  pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
  pptx.title = 'Digital Marketing';
  pptx.author = 'Budogol';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
    slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18,
    slide19, slide20, slide21, slide22, slide23, slide24, slide25].forEach(fn => fn(pptx));

  return pptx.writeFile({
    fileName: path.join(__dirname, '049ecf83-7478-412d-8a1a-d410a424cb4b_grok_final.pptx')
  });
}

build().then(f => console.log('wrote ' + f)).catch(e => { console.error(e); process.exit(1); });
