/**
 * "Dentalist" dental-clinic deck — rebuilt with pptxgenjs.
 *
 * 15 slides, 13.333 x 7.5 in (16:9 widescreen).
 *
 * Fidelity notes
 * --------------
 * - pptxgenjs can only emit solid shape fills, so every gradient of the source
 *   deck is rebuilt by ramp() / rampPath() as a mosaic of opaque tiles that
 *   interpolate the four corner colours. Those corner colours were measured
 *   from a render of the reference file.
 * - Photographs and the SVG icon set are replaced by programmatic stand-ins:
 *   outlines traced from the artwork (see PATHS) plus simple shape glyphs
 *   assembled from rectangles and ellipses (see ICONS).
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const C = {
  navy: '0E384C',   // headline / dark brand colour
  blue: '1E84B5',   // primary accent
  green: '3D8F78',  // secondary accent
  mint: '80C6AF',   // secondary accent, light
  bg: 'EFF8FF',     // page background
  vivid: '0A84FF',  // palette swatch only
  deep: '203A70',   // slide 13 list card
  sky: '93D8F1',    // slide 13 list body
  white: 'FFFFFF',
  g85: 'D9D9D9',
  g75: 'BFBFBF',
  g65: 'A6A6A6',
  g50: '808080',
  ink: '262626',    // theme tx1 lumMod 85 % / lumOff 15 %
};

const HEAD = 'Host Grotesk'; // titles, numbers, UI labels
const BODY = 'Heebo';        // paragraphs, captions

/* ------------------------------------------------------------------ helpers */

/** Blend two hex colours; t = 0 returns a, t = 1 returns b. */
function mix(a, b, t) {
  const ch = (s, i) => parseInt(s.substr(i * 2, 2), 16);
  return [0, 1, 2]
    .map((i) => Math.round(ch(a, i) + (ch(b, i) - ch(a, i)) * t))
    .map((v) => v.toString(16).padStart(2, '0').toUpperCase())
    .join('');
}

/** OOXML "adj" corner value -> pptxgenjs rectRadius (inches). */
const rr = (adj, w, h) => (adj / 100000) * Math.min(w, h);

/**
 * Gradient emulation
 * ------------------
 * pptxgenjs only writes solid fills, so a gradient is rebuilt as: one smooth
 * shape painted in the average colour, covered by a mosaic of opaque tiles
 * that bilinearly interpolate the four measured corner colours. Rounded
 * outlines stay crisp because the mosaic is inset by 0.32 * radius (a rounded
 * rectangle always contains the rectangle inset by ~0.293 * radius).
 */
const INSET = 0.32;

/** Bilinear blend of [topLeft, topRight, bottomLeft, bottomRight] at (u, v). */
function bilinear(c, u, v) {
  return mix(mix(c[0], c[1], u), mix(c[2], c[3], u), v);
}

/** Largest per-channel distance between two hex colours (0..255). */
function delta(a, b) {
  const ch = (s, i) => parseInt(s.substr(i * 2, 2), 16);
  return Math.max(...[0, 1, 2].map((i) => Math.abs(ch(a, i) - ch(b, i))));
}

/** Tiles needed over `len` inches so each step shifts the colour by <= ~5. */
function tileCount(len, spread) {
  return Math.max(1, Math.min(22, Math.round(Math.min(len * 6, spread / 5))));
}

/**
 * Draw a shape filled with a four-corner colour blend.
 *
 *   colors  [topLeft, topRight, bottomLeft, bottomRight] as seen on screen
 *   box     on-screen [x, y, w, h] for the mosaic; needed when the shape is
 *           rotated. Defaults to the shape's own box.
 *   ease    [xExp, yExp] bends the blend towards one side, for the source
 *           deck's radial fills whose focus is not the shape centre.
 */
function ramp(slide, o) {
  const c = o.colors;
  const r = o.rectRadius || 0;
  const ease = o.ease || [1, 1];
  slide.addShape(o.shape || 'roundRect', {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: r || undefined,
    rotate: o.rotate, flipH: o.flipH,
    fill: { color: bilinear(c, Math.pow(0.5, ease[0]), Math.pow(0.5, ease[1])) },
  });

  const pad = r * INSET;
  const base = o.box || [o.x, o.y, o.w, o.h];
  const bx = base[0] + pad, by = base[1] + pad;
  const bw = base[2] - 2 * pad, bh = base[3] - 2 * pad;
  // Tile only as finely as the colour actually changes along each axis.
  const nx = tileCount(bw, Math.max(delta(c[0], c[1]), delta(c[2], c[3])));
  const ny = tileCount(bh, Math.max(delta(c[0], c[2]), delta(c[1], c[3])));

  // Tiles overlap slightly so no hairlines show, but never past the safe box.
  for (let j = 0; j < ny; j++) {
    for (let i = 0; i < nx; i++) {
      const tx = bx + (bw * i) / nx;
      const ty = by + (bh * j) / ny;
      slide.addShape('rect', {
        x: tx, y: ty,
        w: Math.min(bw / nx + 0.02, bx + bw - tx),
        h: Math.min(bh / ny + 0.02, by + bh - ty),
        fill: {
          color: bilinear(c,
            Math.pow((i + 0.5) / nx, ease[0]),
            Math.pow((j + 0.5) / ny, ease[1])),
        },
      });
    }
  }
}

/**
 * Background wash: the corner tint that the slide layouts paint under the
 * content. In the source these are 45-degree gradients that stay fully
 * transparent for the first 55 % of the diagonal and then ramp up to `alpha`
 * at the `from` corner. Drawn as translucent tiles so two washes on opposite
 * corners still overlay correctly.
 */
function wash(slide, o) {
  const tiles = 14;
  const corner = { tl: [0, 0], tr: [1, 0], bl: [0, 1], br: [1, 1] }[o.from];
  for (let j = 0; j < tiles; j++) {
    for (let i = 0; i < tiles; i++) {
      const u = Math.abs((i + 0.5) / tiles - corner[0]);
      const v = Math.abs((j + 0.5) / tiles - corner[1]);
      const reach = Math.max(0, (0.45 - (u + v) / 2) / 0.45);   // 1 at the corner
      const a = o.alpha * reach * reach;
      if (a < 0.012) continue;
      slide.addShape('rect', {
        x: o.x + (o.w * i) / tiles, y: o.y + (o.h * j) / tiles,
        w: o.w / tiles + 0.02, h: o.h / tiles + 0.02,
        fill: { color: o.color, transparency: Math.round(100 - a * 100) },
      });
    }
  }
}

/* ------------------------------------------------------- tooth + icon glyphs */

/**
 * Outlines traced from the deck's vector artwork, normalised to a 1 x 1 box.
 * Commands: ['M', x, y] move, ['C', x1, y1, x2, y2, x, y] cubic, ['Z'] close.
 */
const PATHS = {
  tooth: [
    ['M', 0.760, 0.000],
    ['C', 0.850, 0.010, 0.995, 0.100, 0.999, 0.235],
    ['C', 1.002, 0.340, 0.960, 0.500, 0.933, 0.581],
    ['C', 0.900, 0.700, 0.860, 0.800, 0.810, 0.902],
    ['C', 0.780, 0.950, 0.740, 1.000, 0.690, 1.000],
    ['C', 0.561, 1.000, 0.627, 0.782, 0.523, 0.727],
    ['C', 0.510, 0.727, 0.490, 0.727, 0.477, 0.727],
    ['C', 0.470, 0.733, 0.455, 0.742, 0.444, 0.755],
    ['C', 0.387, 0.834, 0.423, 1.000, 0.310, 1.000],
    ['C', 0.236, 1.000, 0.174, 0.896, 0.123, 0.755],
    ['C', 0.084, 0.639, 0.044, 0.498, 0.026, 0.427],
    ['C', 0.014, 0.380, 0.002, 0.330, 0.001, 0.283],
    ['C', -0.009, 0.106, 0.083, 0.038, 0.177, 0.010],
    ['C', 0.290, -0.023, 0.341, 0.032, 0.399, 0.064],
    ['C', 0.423, 0.091, 0.461, 0.104, 0.500, 0.111],
    ['C', 0.559, 0.111, 0.601, 0.064, 0.646, 0.039],
    ['C', 0.688, 0.008, 0.720, 0.000, 0.760, 0.000],
    ['Z'],
  ],
  // Slide 8: soft backdrop blob, lower lobe
  blobLow: [
    ['M', 0.120, 0.376],
    ['C', 0.069, 0.438, 0.047, 0.499, 0.063, 0.553],
    ['C', 0.111, 0.724, 0.509, 0.768, 0.973, 0.661],
    ['C', 0.982, 0.659, 0.991, 0.657, 1.000, 0.654],
    ['C', 1.000, 0.770, 1.000, 0.884, 1.000, 1.000],
    ['C', 0.976, 1.000, 0.952, 1.000, 0.928, 1.000],
    ['C', 0.912, 0.932, 0.898, 0.883, 0.877, 0.883],
    ['C', 0.856, 0.883, 0.836, 0.922, 0.812, 0.975],
    ['C', 0.809, 0.983, 0.805, 0.992, 0.802, 1.000],
    ['C', 0.646, 1.000, 0.490, 1.000, 0.334, 1.000],
    ['C', 0.315, 0.948, 0.294, 0.888, 0.276, 0.824],
    ['C', 0.130, 0.794, 0.031, 0.732, 0.006, 0.644],
    ['C', -0.017, 0.562, 0.027, 0.469, 0.120, 0.376],
    ['Z'],
  ],
  // Slide 8: soft backdrop blob, upper lobe
  blobHigh: [
    ['M', 0.519, 0.000],
    ['C', 0.695, -0.006, 0.800, 0.155, 0.919, 0.144],
    ['C', 0.937, 0.143, 0.954, 0.137, 1.000, 0.113],
    ['C', 1.000, 0.268, 1.000, 0.424, 1.000, 0.579],
    ['C', 0.981, 0.583, 0.961, 0.588, 0.942, 0.592],
    ['C', 0.656, 0.654, 0.398, 0.659, 0.235, 0.611],
    ['C', 0.204, 0.346, 0.245, 0.087, 0.414, 0.023],
    ['C', 0.452, 0.008, 0.487, 0.001, 0.519, 0.000],
    ['Z'],
  ],
  // Slide 11: the ribbon that orbits the hero tooth
  swoosh: [
    ['M', 0.940, 0.998],
    ['C', 0.910, 0.992, 0.882, 0.975, 0.853, 0.962],
    ['C', 0.836, 0.954, 0.827, 0.950, 0.819, 0.946],
    ['C', 0.789, 0.932, 0.759, 0.915, 0.729, 0.898],
    ['C', 0.692, 0.876, 0.655, 0.853, 0.618, 0.828],
    ['C', 0.592, 0.811, 0.566, 0.793, 0.540, 0.774],
    ['C', 0.479, 0.730, 0.418, 0.683, 0.358, 0.633],
    ['C', 0.310, 0.592, 0.262, 0.550, 0.215, 0.504],
    ['C', 0.189, 0.477, 0.163, 0.449, 0.137, 0.420],
    ['C', 0.128, 0.408, 0.117, 0.397, 0.107, 0.386],
    ['C', 0.076, 0.351, 0.043, 0.310, 0.021, 0.236],
    ['C', -0.049, 0.008, 0.080, -0.002, 0.097, 0.000],
    ['C', 0.037, 0.077, 0.048, 0.165, 0.099, 0.256],
    ['C', 0.105, 0.267, 0.112, 0.278, 0.120, 0.289],
    ['C', 0.143, 0.323, 0.171, 0.357, 0.202, 0.391],
    ['C', 0.311, 0.508, 0.464, 0.619, 0.597, 0.702],
    ['C', 0.651, 0.736, 0.701, 0.764, 0.745, 0.787],
    ['C', 0.781, 0.805, 0.812, 0.820, 0.836, 0.829],
    ['C', 0.862, 0.839, 0.879, 0.843, 0.886, 0.840],
    ['C', 0.937, 0.819, 0.887, 0.685, 0.869, 0.641],
    ['C', 0.909, 0.708, 1.093, 1.028, 0.940, 0.998],
    ['Z'],
  ],
};

/** Draw one of PATHS scaled into { x, y, w, h }. */
function addPath(slide, o) {
  const px = (u) => u * o.w;
  const py = (v) => v * o.h;
  const pts = [];
  o.path.forEach((c) => {
    if (c[0] === 'M') pts.push({ x: px(c[1]), y: py(c[2]), moveTo: true });
    else if (c[0] === 'Z') pts.push({ close: true });
    else pts.push({
      x: px(c[5]), y: py(c[6]),
      curve: { type: 'cubic', x1: px(c[1]), y1: py(c[2]), x2: px(c[3]), y2: py(c[4]) },
    });
  });
  slide.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h, points: pts, rotate: o.rotate || 0,
    fill: { color: o.color, transparency: o.transparency || 0 },
  });
}

const addTooth = (slide, o) => addPath(slide, Object.assign({ path: PATHS.tooth }, o));

/** Approximate a PATHS entry by a polygon (16 line segments per curve). */
function flatten(path, steps = 16) {
  const poly = [];
  let cur = [0, 0];
  path.forEach((c) => {
    if (c[0] === 'M') { cur = [c[1], c[2]]; poly.push(cur); return; }
    if (c[0] === 'Z') return;
    for (let i = 1; i <= steps; i++) {
      const t = i / steps, m = 1 - t;
      poly.push([
        m * m * m * cur[0] + 3 * m * m * t * c[1] + 3 * m * t * t * c[3] + t * t * t * c[5],
        m * m * m * cur[1] + 3 * m * m * t * c[2] + 3 * m * t * t * c[4] + t * t * t * c[6],
      ]);
    }
    cur = [c[5], c[6]];
  });
  return poly;
}

/** Horizontal [from, to] runs of a polygon at height v (even-odd rule). */
function spansAt(poly, v) {
  const xs = [];
  for (let i = 0; i < poly.length; i++) {
    const a = poly[i];
    const b = poly[(i + 1) % poly.length];
    if ((a[1] <= v) !== (b[1] <= v)) xs.push(a[0] + ((v - a[1]) / (b[1] - a[1])) * (b[0] - a[0]));
  }
  xs.sort((p, q) => p - q);
  const out = [];
  for (let i = 0; i + 1 < xs.length; i += 2) out.push([xs[i], xs[i + 1]]);
  return out;
}

/**
 * Fill a PATHS outline with a four-corner colour blend: the smooth silhouette
 * is painted first in the average colour, then a mosaic of tiles clipped to
 * the outline (via scanline spans) paints the blend inside it.
 */
function rampPath(slide, o) {
  addPath(slide, Object.assign({}, o, { color: bilinear(o.colors, 0.5, 0.5) }));

  const poly = flatten(o.path);
  const rows = o.rows || Math.max(8, Math.round(o.h * 6));
  const edge = 0.010;                                   // keep the outline crisp
  for (let j = 0; j < rows; j++) {
    const top = spansAt(poly, j / rows);
    const bottom = spansAt(poly, (j + 1) / rows);
    if (top.length !== bottom.length) continue;         // span count changed: skip
    top.forEach((sp, k) => {
      const x0 = Math.max(sp[0], bottom[k][0]) + edge;
      const x1 = Math.min(sp[1], bottom[k][1]) - edge;
      if (x1 - x0 < 0.03) return;
      const cols = Math.max(1, Math.round((x1 - x0) * o.w * 6));
      for (let i = 0; i < cols; i++) {
        slide.addShape('rect', {
          x: o.x + (x0 + ((x1 - x0) * i) / cols) * o.w,
          y: o.y + (j / rows) * o.h,
          w: ((x1 - x0) / cols) * o.w + 0.015,
          h: o.h / rows + 0.015,
          fill: { color: bilinear(o.colors, x0 + ((x1 - x0) * (i + 0.5)) / cols, (j + 0.5) / rows) },
        });
      }
    });
  }
}

/**
 * Stand-ins for the deck's icon artwork. Every part is
 * [x, y, w, h, shape, rotate, useCutColour] in 0..1 icon space.
 */
const ICONS = {
  tools: [[0.04, 0.00, 0.12, 0.95, 'rect'], [0.26, 0.00, 0.12, 0.95, 'rect'],
          [0.50, 0.28, 0.10, 0.67, 'rect'], [0.44, 0.00, 0.22, 0.30, 'ellipse'],
          [0.78, 0.00, 0.20, 0.95, 'roundRect']],
  bag: [[0.32, 0.00, 0.36, 0.16, 'roundRect'], [0.00, 0.16, 1.00, 0.84, 'roundRect'],
        [0.43, 0.40, 0.14, 0.36, 'rect', 0, 1], [0.32, 0.51, 0.36, 0.14, 'rect', 0, 1]],
  brush: [[0.10, 0.43, 0.80, 0.13, 'rect', -45], [0.62, 0.10, 0.30, 0.26, 'roundRect', -45]],
  braces: [[0.00, 0.40, 1.00, 0.13, 'rect'], [0.03, 0.22, 0.17, 0.52, 'roundRect'],
           [0.29, 0.22, 0.17, 0.52, 'roundRect'], [0.55, 0.22, 0.17, 0.52, 'roundRect'],
           [0.81, 0.22, 0.17, 0.52, 'roundRect']],
  card: [[0.00, 0.14, 1.00, 0.72, 'roundRect'], [0.10, 0.30, 0.26, 0.26, 'ellipse', 0, 1],
         [0.52, 0.32, 0.36, 0.09, 'rect', 0, 1], [0.52, 0.54, 0.36, 0.09, 'rect', 0, 1]],
  phone: [[0.04, 0.04, 0.34, 0.30, 'roundRect'], [0.22, 0.28, 0.28, 0.40, 'rect', -20],
          [0.44, 0.60, 0.38, 0.32, 'roundRect']],
  pin: [[0.14, 0.00, 0.72, 0.74, 'ellipse'], [0.34, 0.50, 0.32, 0.50, 'triangle', 180],
        [0.38, 0.20, 0.24, 0.24, 'ellipse', 0, 1]],
  mail: [[0.00, 0.18, 1.00, 0.64, 'roundRect'], [0.10, 0.26, 0.80, 0.34, 'triangle', 180, 1]],
  cloud: [[0.05, 0.26, 0.90, 0.34, 'roundRect'], [0.16, 0.04, 0.44, 0.40, 'ellipse'],
          [0.50, 0.12, 0.38, 0.34, 'ellipse'], [0.44, 0.50, 0.12, 0.22, 'rect'],
          [0.32, 0.66, 0.36, 0.32, 'triangle', 180]],
};

function addIcon(slide, kind, x, y, size, color, cut) {
  if (kind === 'tooth') { addTooth(slide, { x, y, w: size, h: size, color }); return; }
  ICONS[kind].forEach((p) => {
    slide.addShape(p[4], {
      x: x + p[0] * size, y: y + p[1] * size, w: p[2] * size, h: p[3] * size,
      rotate: p[5] || 0, fill: { color: p[6] ? (cut || C.bg) : color },
      rectRadius: p[4] === 'roundRect' ? size * 0.07 : undefined,
    });
  });
}

/* ------------------------------------------------------- reusable furniture */

/** Brand mark: tooth glyph with a sparkle, as used in the header logo. */
function addLogoMark(slide, x, y, w, h, color) {
  addTooth(slide, { x, y: y + h * 0.06, w: w * 0.78, h: h * 0.94, color });
  slide.addShape('star4', {
    x: x + w * 0.66, y: y + h * 0.00, w: w * 0.36, h: h * 0.48,
    rectRadius: w * 0.05, fill: { color },
  });
}

/** Small circular "go" button with a diagonal arrow. */
function addArrowButton(slide, x, y, d) {
  slide.addShape('ellipse', {
    x, y, w: d, h: d, fill: { color: C.white, transparency: 94 }, line: { color: C.bg, width: 1 },
  });
  slide.addShape('line', {
    x: x + d * 0.3, y: y + d * 0.3, w: d * 0.4, h: d * 0.4, flipV: true,
    line: { color: C.bg, width: 1, endArrowType: 'triangle' },
  });
}

/** Rounded eyebrow pill: "ABOUT US", "OUR TEAM", "OUR SERVICE", ... */
function addEyebrow(slide, x, y, label, o = {}) {
  const w = o.w || 1.469;
  slide.addShape('roundRect', {
    x, y, w, h: 0.383, rectRadius: 0.1915,
    fill: { color: o.chip || C.navy, transparency: 86 },
  });
  slide.addText(label, {
    x: x + 0.12, y, w: w - 0.24, h: 0.383, valign: 'middle', align: o.align || 'left',
    fontFace: BODY, fontSize: 9, charSpacing: 1, color: o.color || C.blue,
  });
}

/** Gradient pill button with a label and (optionally) a trailing arrow button. */
function addPillButton(slide, o) {
  ramp(slide, {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: rr(o.adj || 30739, o.w, o.h), colors: o.colors,
  });
  slide.addText(o.label, {
    x: o.tx, y: o.ty, w: o.tw || 1.455, h: 0.287,
    fontFace: o.font || BODY, fontSize: o.size || 9, color: o.color || C.white,
    lineSpacingMultiple: 1.3,
  });
  if (o.arrow) addArrowButton(slide, o.arrow[0], o.arrow[1], o.arrow[2]);
}

/** Header row: brand mark + wordmark, hamburger rules, "Book appoitment" CTA. */
function addChrome(slide, o = {}) {
  const dx = o.dx || 0;
  const dy = o.dy || 0;
  addLogoMark(slide, 0.425 + dx, 0.291 + dy, 0.313, 0.239, o.mark || mix(C.blue, C.bg, 0.45));
  slide.addText('Dentalist', {
    x: 0.773 + dx, y: 0.251 + dy, w: 0.95, h: 0.303,
    fontFace: HEAD, fontSize: 12, color: o.dark ? C.white : C.navy,
  });
  if (o.nav !== false) {
    [0.346, 0.403, 0.460].forEach((ly) => slide.addShape('line', {
      x: 6.509, y: ly + dy, w: 0.315, h: 0, line: { color: o.dark ? C.g85 : C.green, width: 1 },
    }));
  }
  if (o.cta !== false) {
    const cx = o.ctaX !== undefined ? o.ctaX : 11.403;
    ramp(slide, {
      x: cx, y: 0.220 + dy, w: 1.680, h: 0.365, rectRadius: 0.1825,
      colors: ['0E384C', 'B7BEC1', '325565', 'DBDADA'],
    });
    slide.addText('Book appoitment', {
      x: cx + 0.174, y: 0.276 + dy, w: 1.331, h: 0.252,
      fontFace: HEAD, fontSize: 9, color: C.bg, align: 'center',
    });
  }
}

/** The "Say Goodbay / Dental Problem With / Dentalist ..." tri-colour headline. */
function goodbayHeadline(tail) {
  return [
    { text: 'Say Goodbay ', options: { color: C.navy } },
    { text: 'Dental Problem With ', options: { color: C.g65 } },
    { text: tail, options: { color: C.navy } },
  ];
}

const LOREM_LONG =
  'PLACEHOLDER' +
  'laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi architecto';
const LOREM_XL = LOREM_LONG + ' beatae vitae dicta sunt explicabo. Nemo enim ipsam voluptatem quia';

/* --------------------------------------------------------------- slide 1 */

function slide01(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.bg };

  // Hero card — mint/green diagonal, only the top corners are rounded.
  ramp(s, {
    shape: 'round2SameRect', x: 0.773, y: 2.111, w: 6.846, h: 5.389,
    rectRadius: rr(10039, 6.846, 5.389),
    colors: ['8ACAB6', 'C6E6E2', 'BAE0D9', 'F1F8FF'],
  });

  addChrome(s, { dy: 0.1, nav: false, cta: false });

  addEyebrow(s, 1.638, 2.826, 'WELCOME TO DENTALIST', { w: 2.653, color: C.bg, align: 'center' });
  s.addText('Dentalist', {
    x: 1.531, y: 3.315, w: 5.928, h: 1.582, fontFace: HEAD, fontSize: 88, color: C.navy,
  });
  s.addText(
    'PLACEHOLDER' +
    'doloremque laudantium, totam rem aperiam, eaque ipsa',
    { x: 1.577, y: 5.003, w: 4.403, h: 0.536, fontFace: BODY, fontSize: 9, color: C.green, lineSpacingMultiple: 1.5 }
  );

  s.addShape('roundRect', { x: 1.661, y: 5.793, w: 1.680, h: 0.365, rectRadius: 0.1825, fill: { color: C.navy } });
  s.addText('Book appoitment', {
    x: 1.836, y: 5.849, w: 1.331, h: 0.252, fontFace: HEAD, fontSize: 9, color: C.bg, align: 'center',
  });
  s.addShape('roundRect', {
    x: 3.535, y: 5.793, w: 1.583, h: 0.365, rectRadius: 0.1825, line: { color: C.navy, width: 1 },
  });
  s.addText('Explore now', {
    x: 3.727, y: 5.849, w: 1.063, h: 0.252, fontFace: HEAD, fontSize: 9, color: C.navy,
  });

  s.addShape('star4', {
    x: 6.712, y: 6.752, w: 0.326, h: 0.326, rotate: 20,
    rectRadius: rr(16757, 0.326, 0.326), fill: { color: C.mint },
  });
  s.addShape('star4', {
    x: 1.368, y: 2.411, w: 0.326, h: 0.326, rotate: 330,
    rectRadius: rr(16757, 0.326, 0.326), fill: { color: C.green, transparency: 87 },
  });
}

/* --------------------------------------------------------------- slide 2 */

function slide02(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.bg };
  addChrome(s);

  s.addText('Dental and Dentist', {
    x: 0.371, y: 0.952, w: 4.475, h: 1.919, fontFace: HEAD, fontSize: 54, color: C.navy,
  });
  [[5.410, 1.333, 0.278], [4.508, 2.308, 0.291]].forEach(([x, y, d]) => s.addShape('star4', {
    x, y, w: d, h: d, rectRadius: rr(14130, d, d), fill: { color: C.g85 },
  }));

  // Round seal badge with a sparkle in the middle
  ramp(s, {
    shape: 'ellipse', x: 5.859, y: 3.355, w: 1.170, h: 1.170, rectRadius: 0.585,
    colors: ['3D8F78', '98C1B6', '98C1B6', 'F4F3F3'],
  });
  s.addText('ALIGNING TEETH FOR BETTER HEALTH AND AESTHETICS', {
    x: 6.045, y: 3.400, w: 0.796, h: 0.34, fontFace: BODY, fontSize: 4,
    charSpacing: 0.6, color: C.bg, align: 'center',
  });
  s.addShape('ellipse', { x: 6.216, y: 3.712, w: 0.456, h: 0.456, fill: { color: C.mint } });
  s.addShape('star4', {
    x: 6.264, y: 3.760, w: 0.359, h: 0.359, rectRadius: rr(8397, 0.359, 0.359), fill: { color: C.bg },
  });

  // Wide blue feature card
  ramp(s, {
    x: 0.371, y: 5.032, w: 6.011, h: 2.194, rectRadius: rr(5349, 6.011, 2.194),
    colors: ['1D83B5', '97CCE4', '499EC6', 'C2E5F5'],
  });
  s.addText('Dental and Dentist', {
    x: 0.646, y: 5.277, w: 2.327, h: 0.252, fontFace: HEAD, fontSize: 9, color: C.bg,
  });
  s.addText('Experience Dental Excellence Gentle Touch', {
    x: 0.640, y: 6.131, w: 2.013, h: 0.909, fontFace: HEAD, fontSize: 16, color: C.bg,
  });
  addPillButton(s, {
    x: 4.370, y: 6.407, w: 1.825, h: 0.525, colors: ['1E84B5', '124862', '1A739D', '0E364A'],
    label: 'Get Started!', tx: 4.571, ty: 6.528, tw: 1.227, arrow: [5.794, 6.518, 0.303],
  });

  // Two small feature cards, each with a tooth watermark and three tag chips
  const cards = [
    { x: 6.870, colors: ['838383', 'CCCECE', 'B1B2B2', 'F4F6F7'], title: 'Experienced Team', tw: 1.768,
      tooth: [8.738, 5.627, 0.866], arrow: 9.252,
      tags: [[7.067, 6.056, 0.793, 'toothcare'], [7.067, 6.452, 0.889, 'dentalclinic'], [8.038, 6.452, 0.793, 'doctor']] },
    { x: 9.862, colors: ['42927C', 'B6D1CB', '8BB9AD', 'F6F5F8'], title: 'Comprehensive Services', tw: 2.013,
      tooth: [11.874, 5.715, 0.707], arrow: 12.266,
      tags: [[10.028, 6.086, 0.793, 'toothcare'], [10.028, 6.483, 0.889, 'dentalclinic'], [10.998, 6.483, 0.793, 'doctor']] },
  ];
  cards.forEach((card) => {
    ramp(s, {
      x: card.x, y: 5.065, w: 2.865, h: 1.842, rectRadius: rr(5309, 2.865, 1.842), colors: card.colors,
    });
    s.addText(card.title, {
      x: card.x + 0.11, y: 5.201, w: card.tw, h: 0.572, fontFace: HEAD, fontSize: 14, color: C.bg,
    });
    addTooth(s, { x: card.tooth[0], y: card.tooth[1], w: card.tooth[2], h: card.tooth[2], color: C.bg, transparency: 25 });
    addArrowButton(s, card.arrow, 5.243, 0.359);
    card.tags.forEach(([tx, ty, tw, label]) => {
      s.addShape('roundRect', { x: tx, y: ty, w: tw, h: 0.299, rectRadius: 0.048, line: { color: 'F4F3F9', width: 1 } });
      s.addText(label, {
        x: tx, y: ty, w: tw, h: 0.299, fontFace: BODY, fontSize: 9, color: 'F2F2F2',
        align: 'center', valign: 'middle',
      });
    });
  });
}

/* --------------------------------------------------------------- slide 3 */

function slide03(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.bg };
  wash(s, { x: 0, y: 0.833, w: 7.5, h: 6.667, from: 'bl', color: C.mint, alpha: 0.51 });
  wash(s, { x: 6.667, y: 0, w: 6.667, h: 7.5, from: 'br', color: C.blue, alpha: 0.38 });
  addChrome(s);

  addEyebrow(s, 6.853, 1.745, 'ABOUT US');
  s.addText([
    { text: 'Your Smile, ', options: { color: C.navy } },
    { text: 'Our Commitment ', options: { color: C.g65 } },
    { text: 'to Care', options: { color: C.navy } },
  ], { x: 6.761, y: 2.273, w: 5.283, h: 1.178, fontFace: HEAD, fontSize: 32 });
  s.addText(LOREM_XL, {
    x: 6.847, y: 3.549, w: 5.479, h: 0.764, fontFace: BODY, fontSize: 9, color: C.g65, lineSpacingMultiple: 1.5,
  });

  [6.833, 9.870].forEach((x) => {
    s.addText('Your Title Here', {
      x, y: 4.757, w: 2.457, h: 0.303, fontFace: HEAD, fontSize: 12, color: C.ink,
    });
    s.addText(
      'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem',
      { x, y: 5.060, w: 2.457, h: 0.764, fontFace: BODY, fontSize: 9, color: C.g65, lineSpacingMultiple: 1.5 }
    );
  });

  // Photo card stand-in
  ramp(s, {
    x: 0.833, y: 4.660, w: 4.678, h: 2.213, rectRadius: rr(5349, 4.678, 2.213),
    colors: ['3C8E77', 'B9D3CB', '78AF9F', 'F0F2F2'],
  });
  s.addText('Experience Dental Excellence Gentle Touch', {
    x: 1.007, y: 4.879, w: 2.992, h: 0.640, fontFace: HEAD, fontSize: 16, color: C.bg,
  });
  s.addText('Sed ut perspiciatis unde omnis iste natus error sit volupt', {
    x: 1.007, y: 5.968, w: 2.028, h: 0.536, fontFace: BODY, fontSize: 9, color: C.bg, lineSpacingMultiple: 1.5,
  });
  addPillButton(s, {
    x: 3.891, y: 6.110, w: 1.392, h: 0.456, adj: 21039, colors: ['80C6AF', '4E9C86', '6FB8A1', '3C8E77'],
    label: 'Read More', tx: 3.999, ty: 6.195, tw: 1.067, arrow: [4.934, 6.207, 0.264],
  });
}

/* --------------------------------------------------------------- slide 4 */

function slide04(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.bg };
  wash(s, { x: 0, y: 0, w: 6.667, h: 5.889, from: 'tl', color: C.blue, alpha: 0.38 });
  wash(s, { x: 0, y: 0.833, w: 7.5, h: 6.667, from: 'bl', color: C.mint, alpha: 0.51 });

  // Right-hand image panel with a faint tooth watermark
  ramp(s, {
    x: 6.509, y: 0.806, w: 6.399, h: 6.235, rectRadius: rr(4681, 6.399, 6.235),
    colors: ['1D83B4', '74B6D6', '71B5D5', 'C3E6F6'],
  });
  rampPath(s, {
    path: PATHS.tooth, x: 6.988, y: 1.861, w: 5.441, h: 4.345, rows: 22,
    colors: ['2F8FBC', '7EBCD9', '68B0D2', 'CDE6F4'],
  });

  addChrome(s);
  addArrowButton(s, 12.164, 1.011, 0.532);

  addEyebrow(s, 1.089, 1.613, 'ABOUT US');
  s.addText(goodbayHeadline('Dentalist'), {
    x: 0.959, y: 2.203, w: 4.160, h: 1.717, fontFace: HEAD, fontSize: 32,
  });
  s.addText(LOREM_XL, {
    x: 0.942, y: 4.071, w: 4.177, h: 0.991, fontFace: BODY, fontSize: 9, color: C.g65, lineSpacingMultiple: 1.5,
  });
  addPillButton(s, {
    x: 1.089, y: 5.451, w: 1.825, h: 0.525, colors: ['1E84B5', '124862', '1A739E', '0E374A'],
    label: 'Get Started!', tx: 1.291, ty: 5.572, tw: 1.227, arrow: [2.514, 5.562, 0.303],
  });

  // "24 years of experience" seal
  s.addShape('ellipse', { x: 10.053, y: 4.569, w: 1.900, h: 1.902, fill: { color: C.bg }, line: { color: C.blue, width: 1 } });
  s.addText('LIVE AUCTION . LIVE AUCTION. LIVE AUCTION.', {
    x: 10.266, y: 4.640, w: 1.485, h: 0.30, fontFace: HEAD, fontSize: 4, bold: true,
    charSpacing: 0.6, color: C.blue, align: 'center',
  });
  s.addShape('ellipse', {
    x: 10.433, y: 4.948, w: 1.141, h: 1.143,
    fill: { color: '3690BA' }, line: { color: C.blue, width: 1 },
  });
  s.addText('24', {
    x: 10.249, y: 5.06, w: 1.503, h: 0.55, fontFace: HEAD, fontSize: 32, bold: true,
    color: C.bg, align: 'center', valign: 'middle',
  });
  s.addShape('rect', { x: 10.441, y: 5.677, w: 1.134, h: 0.225, fill: { color: C.blue } });
  s.addText('Experience', {
    x: 10.030, y: 5.677, w: 1.941, h: 0.225, fontFace: HEAD, fontSize: 9,
    color: 'FEFCF4', align: 'center', valign: 'middle',
  });
}

/* --------------------------------------------------------------- slide 5 */

function slide05(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.bg };
  wash(s, { x: 0, y: 0, w: 6.667, h: 5.889, from: 'tl', color: C.blue, alpha: 0.38 });
  wash(s, { x: 6.667, y: 0, w: 6.667, h: 7.5, from: 'br', color: C.mint, alpha: 0.51 });
  addChrome(s);

  addEyebrow(s, 8.534, 1.462, 'ABOUT US');
  s.addText(goodbayHeadline('Dentalist'), {
    x: 8.404, y: 2.006, w: 4.331, h: 1.717, fontFace: HEAD, fontSize: 32,
  });

  // Checklist card
  ramp(s, {
    x: 8.405, y: 4.615, w: 4.331, h: 2.213, rectRadius: rr(5349, 4.331, 2.213),
    colors: ['3C8E77', 'B5D1C9', '7AB0A1', 'F3F3F3'],
  });
  [[5.059, 'Comfortable Waiting Room'],
   [5.560, 'Advanced Dental Equipment'],
   [6.069, 'Dentalic Contact Information']].forEach(([y, label]) => {
    s.addShape('ellipse', { x: 8.883, y: y + 0.055, w: 0.227, h: 0.227, fill: { color: C.bg } });
    s.addShape('rect', { x: 8.949, y: y + 0.145, w: 0.032, h: 0.070, rotate: 320, fill: { color: C.green } });
    s.addShape('rect', { x: 8.986, y: y + 0.105, w: 0.032, h: 0.110, rotate: 35, fill: { color: C.green } });
    s.addText(label, { x: 9.180, y, w: 3.042, h: 0.303, fontFace: HEAD, fontSize: 12, color: C.white });
  });

  // Stat chips (they sit on a photo in the original)
  [[0.773, 3.436, 0.979, 3.527, 1.005, 3.974, '25+', 'Years experience'],
   [0.748, 4.566, 0.954, 4.658, 0.980, 5.105, '677K+', 'Positive review'],
   [0.748, 5.722, 0.954, 5.813, 0.980, 6.260, '144+', 'Professional Dentist']]
    .forEach(([bx, by, nx, ny, lx, ly, num, label]) => {
      s.addShape('roundRect', {
        x: bx, y: by, w: 2.0, h: 1.0, rectRadius: rr(16024, 2.0, 1.0),
        fill: { color: '0D0D0D', transparency: 40 },
      });
      s.addText(num, { x: nx, y: ny, w: 1.529, h: 0.572, fontFace: HEAD, fontSize: 28, color: C.white });
      s.addText(label, {
        x: lx, y: ly, w: 1.545, h: 0.309, fontFace: BODY, fontSize: 9, color: C.g85, lineSpacingMultiple: 1.5,
      });
    });
}

/* --------------------------------------------------------------- slide 6 */

function slide06(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.bg };
  wash(s, { x: 0, y: 0, w: 6.667, h: 5.889, from: 'tl', color: C.blue, alpha: 0.38 });
  wash(s, { x: 5.833, y: 0, w: 7.5, h: 6.667, from: 'tr', color: C.vivid, alpha: 0.20 });

  // Full-width bottom band
  ramp(s, { shape: 'rect', x: 0, y: 4.681, w: 13.333, h: 2.819, colors: ['1D85B6', 'A7D6EC', '368BB4', 'B7D8E6'] });

  addChrome(s);
  addEyebrow(s, 0.773, 1.259, 'ABOUT US');
  s.addText(goodbayHeadline('Dentalist Commitment'), {
    x: 6.500, y: 1.212, w: 6.060, h: 1.717, fontFace: HEAD, fontSize: 32,
  });
  s.addText(LOREM_LONG, {
    x: 0.694, y: 1.761, w: 3.996, h: 0.764, fontFace: BODY, fontSize: 9, color: C.g65, lineSpacingMultiple: 1.5,
  });

  [{ x: 0.748, colors: ['F0F9FF', '318EBC', 'BEDCED', '0476AC'], icon: 1.171, tx: 1.972 },
   { x: 6.791, colors: ['6FBEA3', 'D4ECEC', '83C7B2', 'F2F9FF'], icon: 7.089, tx: 7.890 }].forEach((c) => {
    ramp(s, {
      x: c.x, y: 5.351, w: 5.769, h: 1.425, rectRadius: rr(6972, 4.148, 1.425), colors: c.colors,
    });
    s.addShape('ellipse', { x: c.icon, y: 5.797, w: 0.534, h: 0.534, fill: { color: C.bg } });
    addIcon(s, 'tooth', c.icon + 0.142, 5.939, 0.250, C.navy);
    s.addText('Common Dental Problems & How to Prevent Them', {
      x: c.tx, y: 5.744, w: 3.513, h: 0.640, fontFace: HEAD, fontSize: 16, color: C.bg,
    });
  });
}

/* --------------------------------------------------------------- slide 7 */

function slide07(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.bg };
  addChrome(s);

  // Doctor profile card
  ramp(s, {
    x: 7.072, y: 1.042, w: 5.782, h: 5.826, rectRadius: rr(5349, 5.782, 5.826),
    colors: ['3C8E77', '98C1B5', '98C1B6', 'F2F3F3'],
  });
  addArrowButton(s, 12.112, 1.293, 0.532);
  s.addText('Dr. Ismael Braun', {
    x: 7.461, y: 1.500, w: 4.119, h: 1.582, fontFace: HEAD, fontSize: 44, color: C.bg,
  });
  s.addText('About Me', {
    x: 7.559, y: 3.986, w: 2.474, h: 0.431, fontFace: BODY, fontSize: 16, bold: true,
    color: C.bg, lineSpacingMultiple: 1.3,
  });
  s.addText(
    'PLACEHOLDER' +
    'laudantium, totam rem aperiam, eaque ipsa quae ab illo Sed ut perspiciatis unde omnis iste ' +
    'natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo',
    { x: 7.559, y: 4.395, w: 4.831, h: 1.218, fontFace: BODY, fontSize: 9, color: C.bg, lineSpacingMultiple: 1.5 }
  );
  addPillButton(s, {
    x: 7.461, y: 5.919, w: 5.005, h: 0.631, colors: ['EFF8FF', '509A86', 'DBECF0', '3B8E77'],
    label: 'Book Appoitment now', tx: 7.833, ty: 6.082, tw: 3.056, size: 10, color: C.navy,
    arrow: [11.935, 6.063, 0.303],
  });

  addEyebrow(s, 0.607, 1.463, 'OUR TEAM');
  s.addText('Our Team', {
    x: 0.543, y: 2.008, w: 2.705, h: 0.640, fontFace: HEAD, fontSize: 32, color: C.navy,
  });

  // Two experience tiles
  [{ x: 0.868, adj: 5913, colors: ['3C8E77', '65B099', '58A58E', '81C7B0'], num: [{ text: '10+' }] },
   { x: 2.583, adj: 4772, colors: ['808080', 'C4C4C4', 'AFAFAF', 'F3F3F3'],
     num: [{ text: '9' }, { text: '/10', options: { fontSize: 16 } }] }].forEach((t) => {
    ramp(s, {
      x: t.x, y: 5.552, w: 1.538, h: 1.058, rectRadius: rr(t.adj, 1.538, 1.058), colors: t.colors,
    });
    s.addText(t.num, {
      x: t.x + 0.177, y: 5.711, w: 1.183, h: 0.495, fontFace: HEAD, fontSize: 28,
      color: C.white, align: 'center', lineSpacingMultiple: 0.8,
    });
    s.addText('Experience', {
      x: t.x + 0.177, y: 6.106, w: 1.183, h: 0.346, fontFace: HEAD, fontSize: 12,
      color: C.white, align: 'center', lineSpacingMultiple: 1.3,
    });
  });
}

/* --------------------------------------------------------------- slide 8 */

function slide08(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.bg };

  // Soft grey blob behind the right-hand cards
  [PATHS.blobLow, PATHS.blobHigh].forEach((lobe) => rampPath(s, {
    path: lobe, x: 6.828, y: 0.610, w: 6.505, h: 6.890, rows: 22,
    colors: ['D5D7D9', 'EEF7FE', 'E0E4E9', 'E3E9ED'],
  }));

  addChrome(s);
  addEyebrow(s, 0.669, 1.204, 'OUR DENTIST');
  s.addText('Our Dentist', {
    x: 0.539, y: 1.749, w: 2.873, h: 0.640, fontFace: HEAD, fontSize: 32, color: C.navy,
  });

  // Four dentist cards; the middle one is taller and mint-toned
  [{ x: 0.704, y: 2.922, w: 2.539, h: 3.412, colors: ['0B7AAF', '6CB0D1', '8EC2DD', 'F1F9FF'],
    name: 'Dr. Liana Johns', nx: 0.704, nw: 2.539, role: 'Sedation, Botox', rx: 1.134 },
   { x: 3.515, y: 2.624, w: 2.982, h: 4.008, ease: [0.6, 2.2], colors: ['378C74', '3D8F78', '2C856B', 'F8FDFF'],
     name: 'Dr. Sammie Grady', nx: 3.846, nw: 2.321, role: 'Sedation, Gum Lift', rx: 4.167 },
   { x: 6.770, y: 2.922, w: 2.539, h: 3.412, colors: ['0677AD', '69AED0', '8CC1DC', 'F1F9FF'],
     name: 'Dr. Noe Eichmann', nx: 6.879, nw: 2.321, role: 'Crowns, Inlays & Onlays', rx: 7.200 },
   { x: 9.581, y: 2.922, w: 2.539, h: 3.412, colors: ['0B7AAF', '6CB0D1', '8EC2DD', 'F1F9FF'],
     name: 'Dr. Antonina Bell', nx: 9.690, nw: 2.321, role: 'Full Mouth Rehabilitation', rx: 10.012 },
  ].forEach((c) => {
    ramp(s, {
      x: c.x, y: c.y, w: c.w, h: c.h, rectRadius: rr(11751, c.w, c.h), colors: c.colors,
    });
    s.addText(c.name, {
      x: c.nx, y: 5.084, w: c.nw, h: 0.370, fontFace: HEAD, fontSize: 16, color: C.white, align: 'center',
    });
    s.addText(c.role, {
      x: c.rx, y: 5.499, w: 1.678, h: 0.252, fontFace: BODY, fontSize: 9, color: C.white, align: 'center',
    });
  });
}

/* --------------------------------------------------------------- slide 9 */

function slide09(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.bg };
  addChrome(s);

  addEyebrow(s, 0.824, 1.164, 'OUR SERVICE');
  s.addText('Dental Care Service', {
    x: 0.694, y: 1.708, w: 4.478, h: 0.640, fontFace: HEAD, fontSize: 32, color: C.navy,
  });
  s.addText(LOREM_LONG, {
    x: 0.738, y: 2.395, w: 3.996, h: 0.764, fontFace: BODY, fontSize: 9, color: C.g65, lineSpacingMultiple: 1.5,
  });

  // Photo tile stand-in
  ramp(s, {
    x: 3.248, y: 4.040, w: 2.326, h: 2.778, rectRadius: rr(7227, 2.326, 2.778),
    colors: ['1D84B5', '7EB9D7', '90C3DD', 'F1F9FF'],
  });
  s.addText('Goodbay Dental Problem', {
    x: 3.451, y: 5.671, w: 2.025, h: 0.640, fontFace: HEAD, fontSize: 16, color: C.white,
  });

  const blue = ['EFF8FF', '4097C1', 'CBE4F2', '1C83B4'];
  [{ x: 5.924, y: 3.036, w: 3.112, colors: blue, label: 'Dental Implant', tx: 6.378, ax: 8.506 },
   { x: 9.625, y: 3.036, w: 3.112, colors: blue, label: 'General Dental Care', tx: 10.079, ax: 12.207 },
   { x: 0.882, y: 6.062, w: 2.116, colors: blue, label: 'Dental Cleaning', tx: 1.130, ax: 2.468 },
   { x: 5.924, y: 5.991, w: 6.663, colors: ['F0F8FF', '4B9783', 'E0EFF4', '3B8E77'], label: 'General Dental Care', tx: 6.146, ax: 12.057 },
  ].forEach((p) => {
    addPillButton(s, {
      x: p.x, y: p.y, w: p.w, h: 0.631, colors: p.colors,
      label: p.label, tx: p.tx, ty: p.y + 0.162, tw: 1.788, size: 10, color: C.navy,
      arrow: [p.ax, p.y + 0.144, 0.303],
    });
  });
  s.addShape('line', {
    x: 8.712, y: 6.307, w: 2.366, h: 0, line: { color: C.bg, width: 1, endArrowType: 'triangle' },
  });
}

/* -------------------------------------------------------------- slide 10 */

function slide10(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.bg };
  addChrome(s);

  addEyebrow(s, 5.932, 1.056, 'OUR SERVICE', { align: 'center' });
  s.addText(goodbayHeadline('Dentalist Commitment'), {
    x: 2.544, y: 1.602, w: 8.244, h: 1.178, fontFace: HEAD, fontSize: 32, align: 'center',
  });

  [{ x: 0.872, blue: true, icon: 'bag', ix: 1.196, title: 'General Dentistry', tx: 1.196, ty: 4.639, bx: 1.154, bty: 5.912, lx: 1.683, ly: 6.001 },
   { x: 3.805, blue: false, icon: 'braces', ix: 4.157, title: 'Orthodontics', tx: 4.129, ty: 4.627, bx: 4.087, bty: 5.900, lx: 4.616, ly: 5.989 },
   { x: 6.761, blue: true, icon: 'brush', ix: 7.139, title: 'Implant Restoration', tx: 7.086, ty: 4.633, bx: 7.044, bty: 5.906, lx: 7.573, ly: 5.995 },
   { x: 9.694, blue: false, icon: 'tooth', ix: 10.044, title: 'Cosmetic Dentisry', tx: 10.019, ty: 4.621, bx: 9.977, bty: 5.894, lx: 10.506, ly: 5.983 },
  ].forEach((sv) => {
    const accent = sv.blue ? C.blue : C.g50;
    ramp(s, {
      x: sv.x, y: 3.079, w: 2.750, h: 3.496, rectRadius: rr(8500, 2.750, 3.496),
      colors: sv.blue ? ['F0F8FF', '92C5DE', '79B7D5', '1C83B4']
                       : ['EFF8FF', 'CFD4D8', 'C6CACD', 'A5A5A4'],
    });
    addIcon(s, sv.icon, sv.ix, 3.279, 0.463, accent, sv.blue ? 'CFE6F5' : 'E4E8EC');
    s.addText(sv.title, {
      x: sv.tx, y: sv.ty, w: 1.848, h: 0.320, fontFace: HEAD, fontSize: 13, color: accent,
    });
    s.addText('Sed do ut eiusmod aliqua et incididunt labo', {
      x: sv.tx, y: sv.ty + 0.431, w: 1.848, h: 0.585, fontFace: BODY, fontSize: 10,
      color: accent, lineSpacingMultiple: 1.5,
    });
    ramp(s, {
      x: sv.bx, y: sv.bty, w: 2.184, h: 0.465, rectRadius: 0.2325,
      colors: sv.blue ? ['EFF8FF', '4097C1', 'C9E3F1', '1B82B4']
                       : ['EFF8FF', '929495', 'DBE2E8', '7E7E7D'],
    });
    s.addText('Read more', {
      x: sv.lx, y: sv.ly, w: 1.126, h: 0.287, fontFace: BODY, fontSize: 9,
      color: C.navy, align: 'center', lineSpacingMultiple: 1.3,
    });
  });
}

/* -------------------------------------------------------------- slide 11 */

function slide11(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.bg };

  s.addShape('roundRect', {
    x: 0.264, y: 0.305, w: 12.806, h: 6.891, rectRadius: rr(2961, 12.806, 6.891), fill: { color: C.navy },
  });
  addChrome(s, { dx: 0.107, dy: 0.306, dark: true, ctaX: 11.122 });

  addEyebrow(s, 5.411, 1.297, 'OUR INFOGRAPHIC', { w: 2.510, chip: C.bg, color: C.bg, align: 'center' });
  s.addText('Problem With Dentalist Commitment', {
    x: 2.544, y: 1.998, w: 8.244, h: 0.640, fontFace: HEAD, fontSize: 32, color: C.bg, align: 'center',
  });

  // Hero tooth with its orbiting ribbon
  rampPath(s, {
    path: PATHS.tooth, x: 4.955, y: 3.332, w: 3.119, h: 3.248, rows: 22,
    colors: ['1D84B5', '84BDD9', '89C0DB', 'F1F8FF'],
  });
  rampPath(s, {
    path: PATHS.swoosh, x: 4.570, y: 3.769, w: 3.901, h: 1.891, rows: 20,
    colors: ['1B82B4', 'AED4E8', '74B4D4', 'B3D6E9'],
  });
  [[8.002, 4.672, 0.377], [6.141, 3.280, 0.546]].forEach(([x, y, d]) => s.addShape('star4', {
    x, y, w: d, h: d, rectRadius: rr(20115, d, d), fill: { color: C.blue },
  }));

  [{ icon: 2.805, glyph: 2.995, text: 0.880, tw: 2.870, label: 1.098, align: 'right', title: 'Featured Point Two', kind: 'tooth' },
   { icon: 9.470, glyph: 9.666, text: 9.470, tw: 2.906, label: 9.470, align: 'left', title: 'Featured Point One', kind: 'tools' },
  ].forEach((f) => {
    s.addShape('roundRect', { x: f.icon, y: 3.845, w: 0.797, h: 0.797, rectRadius: 0.19, fill: { color: C.bg } });
    addIcon(s, f.kind, f.glyph, 4.030, 0.420, C.navy);
    s.addText(f.title, {
      x: f.label, y: 4.998, w: 2.651, h: 0.337, fontFace: HEAD, fontSize: 14, bold: true,
      color: C.bg, align: f.align,
    });
    s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Ut sollicitudin iaculis est', {
      x: f.text, y: 5.344, w: f.tw, h: 0.532, fontFace: BODY, fontSize: 9,
      color: C.g50, align: f.align, lineSpacingMultiple: 1.5,
    });
  });
}

/* -------------------------------------------------------------- slide 12 */

function slide12(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.bg };
  addChrome(s);

  addEyebrow(s, 5.932, 1.056, 'OUR SERVICE', { align: 'center' });
  s.addText('Building Healthy Habits Early', {
    x: 2.544, y: 1.602, w: 8.244, h: 0.640, fontFace: HEAD, fontSize: 32, color: C.navy, align: 'center',
  });

  // Doughnut: "Most Ache Part"
  s.addChart(pptx.ChartType.doughnut,
    [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'], values: [0.25, 0.32, 0.41, 0.21] }],
    {
      x: 4.742, y: 2.677, w: 4.017, h: 3.787,
      holeSize: 50, firstSliceAng: 0, showLegend: false, showValue: false,
      chartColors: ['6EAFD1', '6D8997', 'C8CDD0', 'BBE1DA'],
      dataBorder: { pt: 1.5, color: C.white },
    });
  s.addText([{ text: 'Most', options: { breakLine: true } }, { text: 'Ache Part' }], {
    x: 6.084, y: 4.395, w: 1.333, h: 0.471, fontFace: HEAD, fontSize: 10.5, color: C.blue, align: 'center',
  });
  [['25%', 7.214, 3.312, 0.696], ['32%', 7.540, 5.083, 0.694],
   ['41%', 5.335, 5.103, 0.680], ['21%', 5.685, 3.331, 0.680]].forEach(([t, x, y, w]) => s.addText(t, {
    x, y, w, h: 0.303, fontFace: HEAD, fontSize: 12, color: C.bg, align: 'center',
  }));

  // Four labelled callouts around the chart
  [{ icon: 'brush', ix: 2.341, iy: 2.079, tx: 1.387, ty: 2.700, bx: 1.186, by: 3.034,
     body: 'PLACEHOLDER' },
   { icon: 'braces', ix: 2.341, iy: 4.768, tx: 1.387, ty: 5.264, bx: 1.186, by: 5.598,
     body: 'PLACEHOLDER' },
   { icon: 'tools', ix: 10.656, iy: 2.093, tx: 9.703, ty: 2.700, bx: 9.502, by: 3.034,
     body: 'PLACEHOLDER' },
   { icon: 'tooth', ix: 10.656, iy: 4.667, tx: 9.703, ty: 5.264, bx: 9.502, by: 5.598,
     body: 'PLACEHOLDER' },
  ].forEach((n) => {
    addIcon(s, n.icon, n.ix, n.iy, 0.519, C.navy);
    s.addText('Description', {
      x: n.tx, y: n.ty, w: 2.427, h: 0.337, fontFace: HEAD, fontSize: 14, color: C.navy, align: 'center',
    });
    s.addText(n.body, {
      x: n.bx, y: n.by, w: 2.827, h: 0.536, fontFace: BODY, fontSize: 9,
      color: C.g65, align: 'center', lineSpacingMultiple: 1.5,
    });
  });
}

/* -------------------------------------------------------------- slide 13 */

function slide13(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.bg };

  // Two big panels, each rotated so a single corner stays square. `box` is the
  // upright screen rectangle they end up occupying, used to lay out the ramp.
  ramp(s, {
    shape: 'round1Rect', x: 5.971, y: 0.126, w: 6.336, h: 7.765, rotate: 90, flipH: true,
    rectRadius: rr(4596, 6.336, 7.765), box: [5.2565, 0.8405, 7.765, 6.336], ease: [1.3, 1.7],
    colors: ['298AB9', 'B0D5E8', 'DDEEF9', 'E4F2FB'],
  });
  ramp(s, {
    shape: 'round1Rect', x: 1.936, y: 1.682, w: 3.871, h: 7.119, rotate: 270, flipH: true,
    rectRadius: rr(4596, 3.871, 7.119), box: [0.312, 3.306, 7.119, 3.871],
    colors: ['1C83B4', 'A6CFE5', '68ADCF', 'F1F9FF'],
  });

  addChrome(s);
  addEyebrow(s, 0.999, 4.136, 'PORTFOLIO', { color: C.bg });
  s.addText('Say Goodbay Dental Problem With Dentalist', {
    x: 0.869, y: 4.680, w: 4.331, h: 1.717, fontFace: HEAD, fontSize: 32, color: C.bg,
  });

  // Numbered list card
  s.addShape('roundRect', { x: 0.770, y: 1.331, w: 4.025, h: 2.424, rectRadius: rr(5925, 4.025, 2.424), fill: { color: C.deep } });
  s.addShape('roundRect', {
    x: 0.773, y: 2.039, w: 4.025, h: 1.717, rectRadius: rr(8004, 4.025, 1.717),
    fill: { color: C.sky, transparency: 15 },
  });
  s.addText('List title here', {
    x: 1.034, y: 1.534, w: 1.763, h: 0.337, fontFace: HEAD, fontSize: 14, color: C.bg,
  });
  s.addText([
    'Lorem ipsum dolor sit amet.',
    'Qui sint neque a velit modi quo',
    'Non exercitationem reiciendis qui.',
    'Consequatur repudiandae.',
  ].map((t) => ({ text: t, options: { breakLine: true, bullet: { type: 'number', style: 'arabicPeriod', indent: 18 } } })), {
    x: 1.058, y: 2.288, w: 3.115, h: 0.980, fontFace: BODY, fontSize: 11, color: C.deep, lineSpacingMultiple: 1.2,
  });

  // Two KPI figures
  [['90%', 9.271, 9.587, 9.224, 1.218, 'Donec Viverra'],
   ['70%', 10.969, 11.285, 10.494, 2.022, 'Dolor Mollis Nec']].forEach(([pct, px, ix, lx, lw, label]) => {
    s.addText(pct, {
      x: px, y: 1.307, w: 1.125, h: 0.505, fontFace: HEAD, fontSize: 24, color: C.deep, align: 'center',
    });
    s.addShape('mathPlus', { x: ix, y: 1.887, w: 0.494, h: 0.494, fill: { color: C.deep } });
    s.addText(label, {
      x: lx, y: 2.555, w: lw, h: 0.307, fontFace: HEAD, fontSize: 9,
      color: C.deep, align: 'center', lineSpacingMultiple: 1.5,
    });
  });

  s.addShape('roundRect', {
    x: 9.583, y: 5.985, w: 2.301, h: 0.527, rectRadius: rr(23583, 2.301, 0.527), fill: { color: C.navy },
  });
  s.addText('Download Apps Now', {
    x: 9.697, y: 6.087, w: 1.817, h: 0.309, fontFace: BODY, fontSize: 9, color: C.bg, lineSpacingMultiple: 1.5,
  });
  addIcon(s, 'cloud', 11.421, 6.130, 0.238, C.bg, C.navy);
}

/* -------------------------------------------------------------- slide 14 */

function slide14(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.bg };

  ramp(s, {
    x: 0.264, y: 0.305, w: 12.806, h: 6.891, rectRadius: rr(2961, 12.806, 6.891),
    colors: ['F1F8FF', '6DB0D0', 'ACD3E7', '1F84B5'],
  });
  addChrome(s, { dx: 0.25, dy: 0.471, ctaX: 10.979 });

  s.addText('Thank You', {
    x: 5.665, y: 1.931, w: 7.649, h: 1.582, fontFace: HEAD, fontSize: 88, color: C.navy, align: 'center',
  });

  ramp(s, {
    x: 6.667, y: 5.139, w: 6.217, h: 1.832, rectRadius: rr(11577, 6.217, 1.832),
    colors: ['1D84B5', 'C1DEEE', '4D9EC6', 'F1F9FF'],
  });
  [{ icon: 'card', ix: 7.174, iy: 5.530, tx: 7.528, ty: 5.462, tw: 1.459, text: 'Reymundo Haley', cut: '4E9EC6' },
   { icon: 'pin', ix: 7.174, iy: 6.138, tx: 7.528, ty: 6.069, tw: 1.961, text: '81012 Gerhold Drive Suite 135, Italy', cut: '62AACD' },
   { icon: 'phone', ix: 10.278, iy: 5.530, tx: 10.632, ty: 5.462, tw: 1.471, text: '+1-274-400-1747', cut: '8CC0DB' },
   { icon: 'mail', ix: 10.278, iy: 6.138, tx: 10.632, ty: 6.069, tw: 1.853, text: 'admin@dentcare.com', cut: 'A2CDE4' },
  ].forEach((c) => {
    addIcon(s, c.icon, c.ix, c.iy, 0.250, C.bg, c.cut);
    s.addText(c.text, {
      x: c.tx, y: c.ty, w: c.tw, h: 0.341, fontFace: BODY, fontSize: 12, color: C.bg, lineSpacingMultiple: 1.25,
    });
  });
}

/* -------------------------------------------------------------- slide 15 */

function slide15(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.bg };

  // Rules that split the style sheet into three regions
  s.addShape('line', { x: 4.875, y: 0, w: 0, h: 7.5, line: { color: C.g85, width: 0.75 } });
  s.addShape('line', { x: 0, y: 4.309, w: 4.867, h: 0, line: { color: C.g85, width: 0.75 } });
  s.addShape('line', { x: 4.867, y: 3.776, w: 8.467, h: 0, line: { color: C.g85, width: 0.75 } });

  // Logo lock-up
  ramp(s, {
    x: 1.893, y: 1.120, w: 0.806, h: 0.815, rectRadius: rr(13303, 0.806, 0.815),
    colors: ['1E84B5', '124862', '1A739D', '0E364A'],
  });
  addLogoMark(s, 2.031, 1.335, 0.531, 0.426, C.bg);
  s.addText('Dentalist', {
    x: 1.070, y: 1.998, w: 2.452, h: 0.707, fontFace: HEAD, fontSize: 36, color: C.navy, align: 'center',
  });
  s.addText('Logo – Host Grotesk', {
    x: 0.959, y: 2.633, w: 2.675, h: 0.356, fontFace: BODY, fontSize: 11,
    color: C.blue, align: 'center', lineSpacingMultiple: 1.5,
  });

  // Palette swatches
  [[1.663, 5.214, C.navy], [2.122, 5.214, C.vivid], [2.581, 5.214, '5DA992'],
   [1.663, 5.633, C.blue], [2.122, 5.633, C.g65], [2.581, 5.633, C.g75],
  ].forEach(([x, y, color]) => s.addShape('ellipse', { x, y, w: 0.348, h: 0.348, fill: { color } }));
  s.addText('Pallete', {
    x: 1.703, y: 6.053, w: 1.189, h: 0.356, fontFace: BODY, fontSize: 11,
    color: C.blue, align: 'center', lineSpacingMultiple: 1.5,
  });

  // Type specimens
  s.addText('Tittle  – Host Grotesk', {
    x: 5.691, y: 1.049, w: 4.240, h: 0.356, fontFace: BODY, fontSize: 11, color: C.blue, lineSpacingMultiple: 1.5,
  });
  s.addText('Sed Ut Perspiciatis Unde Omnis Iste Natus Error Sit', {
    x: 5.691, y: 1.414, w: 7.071, h: 1.313, fontFace: HEAD, fontSize: 36, color: C.navy,
  });
  s.addText('Long Text - Heebo', {
    x: 5.691, y: 4.637, w: 2.188, h: 0.356, fontFace: BODY, fontSize: 11, color: C.blue, lineSpacingMultiple: 1.5,
  });
  s.addText(
    'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, ' +
    'totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi architecto beatae vitae dicta ' +
    'sunt explicabo. Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut odit aut fugit, sed quia ' +
    'consequuntur magni dolores eos qui ratione voluptatem sequi nesciunt. Neque porro quisquam est, qui ' +
    'dolorem ipsum quia dolor sit amet, consectetur, adipisci velit, sed quia non numquam eius modi tempora ' +
    'incidunt ut labore et dolore magnam',
    { x: 5.691, y: 5.099, w: 6.432, h: 1.218, fontFace: BODY, fontSize: 9, color: C.g65, lineSpacingMultiple: 1.5 }
  );
}

/* ------------------------------------------------------------------- main */

function build() {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5 in
  pptx.title = 'Dentalist';
  pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
   slide09, slide10, slide11, slide12, slide13, slide14, slide15].forEach((fn) => fn(pptx));

  return pptx.writeFile({
    fileName: path.join(__dirname, '0f86b0a4-d74c-4cbc-99b0-45c2ed6f156a_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
