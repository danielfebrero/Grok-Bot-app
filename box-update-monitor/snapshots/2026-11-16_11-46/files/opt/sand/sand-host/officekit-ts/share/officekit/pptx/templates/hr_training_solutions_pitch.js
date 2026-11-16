/**
 * HR Training deck — recreated with pptxgenjs.
 * Slide canvas: 26.667 x 15 in (widescreen, 4x scale of 6.667x3.75).
 * Run:  node 108af3ff-3f8e-4070-9257-1847ac54b1e4_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- design tokens
const W = 26.667;              // slide width  (in)
const H = 15;                  // slide height (in)
const FONT = 'Archivo';        // theme major/minor face

const C = {
  bg: '09192C',                // page background
  card: '0C2441',              // panel / card fill
  purple: 'A73BF6',            // accent 1
  orange: 'FF9246',            // accent 2
  yellow: 'FFDC5C',            // accent 3
  gray: '666666',              // accent 4
  white: 'FFFFFF',
  dim: 'A6A6A6',               // white lumMod 65%
  soft: 'D9D9D9',              // white lumMod 85% / photo placeholder
  pale: 'F2F2F2',              // white lumMod 95%
  muted: '8A9AA8',             // caption grey-blue
  ink: '404040',               // black lumMod 75 / lumOff 25
};

// ---------------------------------------------------------------- tiny helpers
const R = (text, options) => ({ text, options });

/** Text box: OOXML default anchoring is top-left, pptxgenjs defaults to middle. */
function T(slide, text, o) {
  slide.addText(text, Object.assign({ fontFace: FONT, color: C.white, valign: 'top' }, o));
}

/** The tinted "glass" pill / donut stroke used all over the deck. */
const GLASS_FILL = { color: C.white, transparency: 91 };
const GLASS_LINE = { color: C.white, transparency: 48, width: 1 };

/** Grey block standing in for a photo (rasters are not embedded). */
function photo(slide, x, y, w, h, o = {}) {
  slide.addShape(o.shape || 'rect', Object.assign(
    { x, y, w, h, fill: { color: C.soft } },
    o.rectRadius ? { rectRadius: o.rectRadius } : {},
    o.rotate ? { rotate: o.rotate } : {}
  ));
  if (o.label !== false) {
    T(slide, '[image]', {
      x, y: y + h / 2 - 0.32, w, h: 0.64,
      align: 'center', valign: 'middle', fontSize: o.labelSize || 16, color: '8C9196',
    });
  }
}

/** Brand mark: three slanted bars (traced from the deck's SVG logo, 226x170). */
function logo(slide, x, y, w, h, color) {
  const bars = [
    [[102.5, 0], [189.3, 0], [142.4, 50.4], [55.6, 50.4]],
    [[47.1, 59.8], [226, 59.8], [178.9, 110.2], [0, 110.2]],
    [[83.3, 119.6], [170.4, 119.6], [123.4, 170], [36.7, 170]],
  ];
  bars.forEach((pts) => slide.addShape('custGeom', {
    x, y, w, h,
    fill: { color },
    points: pts.map(([px, py]) => ({ x: (px / 226) * w, y: (py / 170) * h })).concat([{ close: true }]),
  }));
}

/** Outline circle with an arrow inside (the "read more" icon). */
function arrowCircle(slide, x, y, d, color, dir) {
  slide.addShape('ellipse', { x, y, w: d, h: d, line: { color, width: Math.max(0.75, d * 1.6) } });
  const m = d * 0.24, mid = y + d / 2;
  slide.addShape('line', {
    x: x + m, y: mid, w: d - 2 * m, h: 0, flipH: dir === 'left',
    line: { color, width: Math.max(0.75, d * 1.6), endArrowType: 'triangle' },
  });
}

/**
 * Pictogram library — the source art is SVG, so each icon is rebuilt from a few
 * primitives. Parts are [shape, x, y, w, h, 'f'ill | 'o'utline, rotate] in 0..1 units.
 */
const ICONS = {
  thumb: [['roundRect', 0.26, 0.34, 0.48, 0.56, 'f'], ['roundRect', 0.30, 0.04, 0.20, 0.36, 'f', -20], ['roundRect', 0.02, 0.50, 0.22, 0.42, 'f']],
  doc: [['rect', 0.16, 0.06, 0.60, 0.88, 'o'], ['rect', 0.26, 0.26, 0.40, 0.05, 'f'], ['rect', 0.26, 0.44, 0.40, 0.05, 'f'], ['rect', 0.26, 0.62, 0.26, 0.05, 'f']],
  monitor: [['rect', 0.04, 0.12, 0.92, 0.60, 'o'], ['rect', 0.42, 0.74, 0.16, 0.10, 'f'], ['rect', 0.26, 0.86, 0.48, 0.07, 'f']],
  gear: [['gear9', 0.00, 0.00, 1.00, 1.00, 'o'], ['ellipse', 0.34, 0.34, 0.32, 0.32, 'o']],
  people: [['ellipse', 0.06, 0.14, 0.22, 0.22, 'o'], ['ellipse', 0.72, 0.14, 0.22, 0.22, 'o'], ['ellipse', 0.39, 0.04, 0.22, 0.22, 'o'],
           ['roundRect', 0.02, 0.44, 0.30, 0.50, 'o'], ['roundRect', 0.68, 0.44, 0.30, 0.50, 'o'], ['roundRect', 0.34, 0.34, 0.32, 0.60, 'o']],
  mouse: [['roundRect', 0.28, 0.02, 0.44, 0.96, 'o'], ['roundRect', 0.46, 0.16, 0.08, 0.22, 'f']],
  wifi: [['arc', 0.02, 0.10, 0.96, 0.96, 'o'], ['arc', 0.20, 0.28, 0.60, 0.60, 'o'], ['ellipse', 0.42, 0.68, 0.16, 0.16, 'f']],
  save: [['rect', 0.06, 0.06, 0.88, 0.88, 'f'], ['rect', 0.30, 0.06, 0.40, 0.30, 'o'], ['rect', 0.22, 0.52, 0.56, 0.42, 'o']],
  heart: [['heart', 0.04, 0.08, 0.92, 0.84, 'f']],
  globe: [['ellipse', 0.04, 0.04, 0.92, 0.92, 'o'], ['ellipse', 0.34, 0.04, 0.32, 0.92, 'o'], ['rect', 0.04, 0.48, 0.92, 0.04, 'f']],
  cloud: [['cloud', 0.02, 0.14, 0.96, 0.72, 'f']],
  usb: [['roundRect', 0.16, 0.34, 0.68, 0.52, 'f'], ['rect', 0.44, 0.06, 0.12, 0.28, 'f']],
  briefcase: [['roundRect', 0.02, 0.26, 0.96, 0.64, 'f'], ['rect', 0.34, 0.08, 0.32, 0.18, 'f'], ['rect', 0.44, 0.46, 0.12, 0.16, 'o']],
  bird: [['teardrop', 0.06, 0.20, 0.62, 0.62, 'o', 225], ['triangle', 0.40, 0.00, 0.58, 0.40, 'o', 120]],
  play: [['triangle', 0.28, 0.24, 0.48, 0.52, 'f', 90]],
  pencil: [['rect', 0.24, 0.20, 0.62, 0.28, 'f', -45], ['triangle', 0.04, 0.62, 0.30, 0.30, 'f', -135]],
};

function icon(slide, kind, x, y, s, color) {
  ICONS[kind].forEach(([shape, dx, dy, dw, dh, mode, rot]) => slide.addShape(shape, Object.assign(
    { x: x + dx * s, y: y + dy * s, w: dw * s, h: dh * s },
    mode === 'o' ? { line: { color, width: Math.max(0.75, s * 1.4) } } : { fill: { color } },
    rot ? { rotate: rot } : {}
  )));
}

/** The rich "Empower your organization…" headline reused on most slides. */
function empower(size) {
  const f = { fontFace: FONT, fontSize: size, color: C.white };
  return [
    R('Empower', Object.assign({}, f, { color: C.purple })),
    R(' your organization with ', f),
    R('our tailored HR training solutions', Object.assign({}, f, { transparency: 84 })),
    R(". Whether you're looking to build ", f),
    R('leadership', Object.assign({}, f, { color: C.orange })),
    R(' skills, streamline ', f),
    R('recruitment', Object.assign({}, f, { transparency: 77 })),
    R(' ', f),
    R('processes', Object.assign({}, f, { color: C.yellow })),
  ];
}

/** Headline shared by slides 8-30: centred, 54pt, 18.848in wide. */
function headline(slide, y, o = {}) {
  T(slide, empower(o.size || 54), Object.assign(
    { x: o.x !== undefined ? o.x : 3.909, y, w: o.w || 18.848, h: 2.827, align: o.align || 'center' }, o.extra || {}
  ));
}

/** Master furniture: brand tag top-left, page chip bottom-left. */
function chrome(slide, num) {
  T(slide, 'HR Training', { x: 0.85, y: 0.82, w: 1.683, h: 0.438, fontSize: 20, wrap: false });
  slide.addShape('rect', {
    x: 0.983, y: 13.595, w: 3.317, h: 0.586,
    fill: { color: C.soft, transparency: 90 }, line: { color: C.dim, transparency: 50, width: 0.5 },
  });
  T(slide, 'www.website.com   /', { x: 0.922, y: 13.686, w: 2.728, h: 0.404, fontSize: 18, align: 'right' });
  T(slide, String(num), { x: 3.253, y: 13.686, w: 0.939, h: 0.404, fontSize: 18, align: 'right', valign: 'middle' });
}

/**
 * Big translucent ring used as a background flourish.
 * `band` is the ring thickness as a fraction of the diameter (the donut `adj`).
 */
function ring(slide, x, y, d, band) {
  slide.addShape('donut', { x, y, w: d, h: d, rectRadius: band * d, fill: GLASS_FILL, line: GLASS_LINE });
}

/** Cog wheel: `teeth` trapezoidal teeth around a disc (the gear9 preset is too coarse). */
function gear(slide, x, y, d, teeth, color) {
  const R = d / 2, ROOT = R * 0.79, PITCH = (2 * Math.PI) / teeth, TIP = PITCH * 0.20, GAP = PITCH * 0.34;
  const pts = [];
  const at = (ang, rad) => pts.push({ x: R + rad * Math.cos(ang), y: R + rad * Math.sin(ang) });
  for (let i = 0; i < teeth; i++) {
    const c = i * PITCH;
    at(c - TIP, R); at(c + TIP, R);           // tooth tip
    at(c + GAP, ROOT); at(c + PITCH - GAP, ROOT); // valley
  }
  pts[0].moveTo = true;
  pts.push({ close: true });
  slide.addShape('custGeom', { x, y, w: d, h: d, fill: { color }, points: pts });
}

/** Cubic spline drawn from normalised control points (used by the line chart). */
function spline(slide, x, y, w, h, norm, color) {
  const pts = [{ x: norm[0][0] * w, y: norm[0][1] * h, moveTo: true }];
  for (let i = 1; i + 2 < norm.length + 1; i += 3) {
    pts.push({
      x: norm[i + 2][0] * w, y: norm[i + 2][1] * h,
      curve: { type: 'cubic', x1: norm[i][0] * w, y1: norm[i][1] * h, x2: norm[i + 1][0] * w, y2: norm[i + 1][1] * h },
    });
  }
  slide.addShape('custGeom', { x, y, w, h, points: pts, line: { color, width: 1.5 } });
}

// ================================================================= slide builders
const slides = [];
const S = (fn) => slides.push(fn);

// -- 1 --------------------------------------------------------- cover
S((s) => {
  T(s, 'www.site.com', { x: 2.44, y: 2.057, w: 2.602, h: 0.572, fontSize: 28, wrap: false });
  logo(s, 2.44, 6.028, 4.105, 3.085, C.purple);
  // wrap:false in the original lets the line overflow its box symmetrically,
  // so the box is widened here to land the glyphs in the same place.
  T(s, [R('HR', { bold: true }), R(' Training')], {
    x: 1.72, y: 9.493, w: 17.9, h: 3.45, fontSize: 199,
  });
  ring(s, 12.444, 3.333, 22.278, 0.134);
});

// -- 2 --------------------------------------------------------- intro
S((s) => {
  ring(s, 12.444, 3.333, 22.278, 0.134);
  logo(s, 2.331, 2.511, 1.629, 1.224, C.purple);
  T(s, 'HR Training', { x: 4.189, y: 2.803, w: 2.576, h: 0.64, fontSize: 32, wrap: false });
  T(s, empower(54), { x: 2.331, y: 6.936, w: 9.403, h: 5.554 });
  T(s, 'www.site.com', { x: 13.173, y: 2.837, w: 2.602, h: 0.572, fontSize: 28, color: C.yellow, wrap: false });
});

// -- 3 --------------------------------------------------------- "What is HR Training?"
S((s) => {
  ring(s, 7.585, -13.049, 22.278, 0.181);
  logo(s, 2.129, 2.272, 1.629, 1.224, C.purple);
  T(s, 'HR Training', { x: 3.987, y: 2.564, w: 2.576, h: 0.64, fontSize: 32, wrap: false });
  T(s, 'What is', { x: 1.982, y: 7.831, w: 9.161, h: 2.794, fontSize: 160 });
  T(s, 'HR Training?', { x: 1.982, y: 9.934, w: 14.296, h: 2.794, fontSize: 160, bold: true, color: C.purple });
  photo(s, 18.918, -0.023, 7.748, 15.023);
});

// -- 4 --------------------------------------------------------- copy-heavy
S((s) => {
  ring(s, 17.719, 2.475, 22.278, 0.181);
  T(s, empower(54), { x: 2.63, y: 3.015, w: 18.544, h: 2.827 });
  T(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. ', { x: 2.63, y: 6.661, w: 14.688, h: 0.774, fontSize: 40 });
  T(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence in this spot. ',
    { x: 2.63, y: 7.89, w: 15.052, h: 1.098, fontSize: 24, lineSpacingMultiple: 1.3 });
  T(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone',
    { x: 2.63, y: 9.807, w: 7.954, h: 1.623, fontSize: 24, lineSpacingMultiple: 1.3 });
  T(s, '- A wonderful serenity has taken possession',
    { x: 10.96, y: 10.857, w: 7.954, h: 0.573, fontSize: 24, italic: true, lineSpacingMultiple: 1.3 });
});

/** The tilted "65+" badge that appears on slides 5 and 6. */
function badge65(s, ovalX, ovalY, ringColor, textColor, capColor) {
  s.addShape('ellipse', { x: ovalX, y: ovalY, w: 5.049, h: 5.049, fill: { color: ringColor } });
  const cx = ovalX + 2.5245, cy = ovalY + 2.5245;
  T(s, '65', { x: cx - 1.5265, y: cy - 1.3435, w: 2.004, h: 1.717, fontSize: 96, bold: true, align: 'center', color: textColor, rotate: -16.6 });
  T(s, '+', { x: cx + 0.5113 - 0.6685, y: cy - 0.5564 - 0.8585, w: 1.337, h: 1.717, fontSize: 96, bold: true, align: 'center', color: C.orange, rotate: -30.3 });
  T(s, 'Lorem ipsum dolor sit amet, consectetuer.', {
    x: cx + 0.2601 - 1.893, y: cy + 0.8706 - 0.549, w: 3.786, h: 1.098,
    fontSize: 24, align: 'center', color: capColor, lineSpacingMultiple: 1.3, rotate: -16.6,
  });
}

// -- 5 --------------------------------------------------------- split photo / purple panel
S((s) => {
  photo(s, 0, 0, 14.899, 15, { labelSize: 26 });
  ring(s, 23.344, 0.152, 10.893, 0.130);
  s.addShape('rect', { x: 14.9, y: 7.5, w: 11.767, h: 7.5, fill: { color: C.purple } });
  T(s, empower(48), { x: 16.625, y: 1.061, w: 8.316, h: 4.948 });
  badge65(s, 18.259, 8.726, C.white, '000000', '000000');
});

// -- 6 --------------------------------------------------------- purple page + notched blob
S((s) => {
  s.background = { color: C.purple };
  // Layout blob: rounded panel with a circular bite taken out of the bottom edge.
  s.addShape('custGeom', {
    x: 9.543, y: 0, w: 17.124, h: 15, fill: { color: C.soft },
    points: [
      { x: 7.683, y: 0, moveTo: true }, { x: 14.733, y: 0 }, { x: 15.062, y: 0.111 },
      { x: 17.023, y: 1.055, curve: { type: 'cubic', x1: 15.749, y1: 0.363, x2: 16.405, y2: 0.680 } },
      { x: 17.124, y: 1.119 }, { x: 17.124, y: 15 }, { x: 13.435, y: 15 }, { x: 13.543, y: 14.945 },
      { x: 16.105, y: 10.641, curve: { type: 'cubic', x1: 15.069, y1: 14.117, x2: 16.105, y2: 12.500 } },
      { x: 11.211, y: 5.745, curve: { type: 'cubic', x1: 16.105, y1: 7.937, x2: 13.913, y2: 5.745 } },
      { x: 6.316, y: 10.641, curve: { type: 'cubic', x1: 8.506, y1: 5.745, x2: 6.316, y2: 7.937 } },
      { x: 8.878, y: 14.945, curve: { type: 'cubic', x1: 6.316, y1: 12.500, x2: 7.352, y2: 14.117 } },
      { x: 8.985, y: 15 }, { x: 0.880, y: 15 }, { x: 0.680, y: 14.495 },
      { x: 0, y: 10.641, curve: { type: 'cubic', x1: 0.240, y1: 13.293, x2: 0, y2: 12.000 } },
      { x: 7.355, y: 0.111, curve: { type: 'cubic', x1: 0, y1: 5.805, x2: 3.064, y2: 1.683 } },
      { close: true },
    ],
  });
  T(s, '[image]', { x: 14.5, y: 5.2, w: 3, h: 0.6, align: 'center', valign: 'middle', fontSize: 18, color: '8C9196' });
  badge65(s, 6.392, 7.507, C.bg, C.white, C.white);
  T(s, [
    R('Unlock Your Team\u2019s ', { fontFace: FONT, fontSize: 66, color: C.white }),
    R('Potential', { fontFace: FONT, fontSize: 66, color: C.orange }),
    R(' with Expert HR Training!', { fontFace: FONT, fontSize: 66, color: C.white }),
  ], { x: 1.905, y: 2.444, w: 8.476, h: 3.433 });
  logo(s, 18.7, 9.098, 4.105, 3.085, C.white);
});

// -- 7 --------------------------------------------------------- photo grid
S((s) => {
  s.background = { color: C.purple };
  [[0, 0, 13.333, 4.852], [13.333, 0, 13.333, 4.852], [17.778, 4.852, 8.889, 10.148],
   [8.889, 4.852, 8.889, 10.148], [0, 4.852, 8.889, 10.148]].forEach(([x, y, w, h]) =>
    photo(s, x, y, w, h, { labelSize: 20 }));
  T(s, empower(54), { x: 1.692, y: 8.941, w: 13.333, h: 4.570, fill: { color: C.bg } });
});

// -- 8 --------------------------------------------------------- team
S((s) => {
  T(s, 'Meet our team Training', { x: 4.259, y: 1.574, w: 18.148, h: 1.717, fontSize: 96, align: 'center' });
  [['Our Best Team', 8.978, true], ['Customer Support', 11.981, false], ['Group Strategy', 14.983, false]]
    .forEach(([label, x, active]) => {
      s.addShape('roundRect', { x, y: 3.719, w: 2.706, h: 0.658, rectRadius: 0.329, fill: active ? { color: C.card } : { type: 'none' } });
      T(s, label, { x, y: 3.719, w: 2.706, h: 0.658, fontSize: 20, align: 'center', valign: 'middle' });
    });
  [2.874, 10.147, 17.421].forEach((x) => {
    s.addShape('rect', { x, y: 5.224, w: 6.372, h: 8.202, fill: { color: C.card } });
    T(s, 'Sokovia Nguen', { x: x + 0.687, y: 5.794, w: 4.719, h: 0.841, fontSize: 44 });
    photo(s, x + 1.141, 7.281, 4.088, 4.088, { shape: 'ellipse' });
    T(s, 'Sales Trainer', { x: x + 0.687, y: 12.217, w: 3.208, h: 0.64, fontSize: 32 });
    arrowCircle(s, x + 5.127, 12.258, 0.557, C.white);
  });
});

// -- 9 --------------------------------------------------------- two feature cards
S((s) => {
  [1.212, 13.508].forEach((x) => s.addShape('rect', { x, y: 6.692, w: 11.946, h: 6.357, fill: { color: C.card } }));
  T(s, empower(54), { x: 1.212, y: 1.951, w: 15.807, h: 3.736 });
  const body = 'PLACEHOLDER';
  T(s, body, { x: 20.621, y: 2.745, w: 4.833, h: 2.148, fontSize: 24, lineSpacingMultiple: 1.3 });
  // The label box is narrow, so "Features" wraps to a first line that the photo
  // (drawn last, as in the original) covers — only the second word shows.
  [['Features One', 1.543], ['Features Two', 13.839]].forEach(([label, x]) => {
    T(s, label, { x, y: 9.979, w: 3.646, h: 0.774, fontSize: 40, fill: { color: C.purple } });
    T(s, body, { x, y: 11.208, w: 7.622, h: 1.098, fontSize: 24, lineSpacingMultiple: 1.3 });
    photo(s, x, 7.003, 11.285, 3.75, { labelSize: 20 });
  });
});

// -- 10 -------------------------------------------------------- three thumbnails
S((s) => {
  T(s, empower(54), { x: 2.682, y: 2.162, w: 19.409, h: 2.827 });
  T(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut',
    { x: 2.682, y: 5.877, w: 10.242, h: 1.623, fontSize: 24, color: C.dim });
  arrowCircle(s, 20.727, 6.773, 1.284, C.dim, 'left');
  arrowCircle(s, 22.382, 6.773, 1.284, C.white);
  [2.682, 9.84, 17.0].forEach((x) => photo(s, x, 8.847, 6.667, 3.991, { labelSize: 18 }));
});

// -- 11 -------------------------------------------------------- programme cards
S((s) => {
  T(s, empower(54), { x: 1.697, y: 2.164, w: 19.545, h: 2.827 });
  [1.697, 9.424, 17.152, 24.879].forEach((x) => {
    s.addShape('rect', { x, y: 5.594, w: 7.242, h: 7.242, fill: { color: C.purple } });
    s.addShape('ellipse', { x: x + 4.414, y: 6.632, w: 1.588, h: 1.588, fill: { color: C.bg } });
    icon(s, 'thumb', x + 4.85, 7.02, 0.78, C.yellow);
    photo(s, x + 0.66, 5.989, 4.569, 4.569, { shape: 'ellipse' });
    s.addShape('roundRect', { x: x + 3.108, y: 9.685, w: 3.475, h: 1.064, rectRadius: 0.532, rotate: -16.4, fill: { color: C.orange } });
    T(s, '+188K Class', { x: x + 3.108, y: 9.685, w: 3.475, h: 1.064, fontSize: 32, align: 'center', valign: 'middle', rotate: -16.4 });
    T(s, 'Customized Programs', { x: x + 0.702, y: 11.566, w: 5.838, h: 0.774, fontSize: 40, align: 'center' });
  });
});

// -- 12 -------------------------------------------------------- text + big blob
S((s) => {
  // Layout blob: full-height panel with a circular bite on the left edge.
  s.addShape('custGeom', {
    x: 13.333, y: 0, w: 13.333, h: 15, fill: { color: C.soft },
    points: [
      { x: 2.904, y: 0, moveTo: true }, { x: 13.333, y: 0 }, { x: 13.333, y: 4.610 }, { x: 13.170, y: 4.488 },
      { x: 11.139, y: 3.867, curve: { type: 'cubic', x1: 12.590, y1: 4.095, x2: 11.890, y2: 3.867 } },
      { x: 7.505, y: 7.5, curve: { type: 'cubic', x1: 9.132, y1: 3.867, x2: 7.505, y2: 5.493 } },
      { x: 11.139, y: 11.133, curve: { type: 'cubic', x1: 7.505, y1: 9.507, x2: 9.132, y2: 11.133 } },
      { x: 13.170, y: 10.512, curve: { type: 'cubic', x1: 11.890, y1: 11.133, x2: 12.590, y2: 10.905 } },
      { x: 13.333, y: 10.390 }, { x: 13.333, y: 15 }, { x: 2.904, y: 15 }, { x: 2.893, y: 14.990 },
      { x: 0, y: 7.5, curve: { type: 'cubic', x1: 1.096, y1: 13.011, x2: 0, y2: 10.383 } },
      { x: 2.893, y: 0.010, curve: { type: 'cubic', x1: 0, y1: 4.617, x2: 1.096, y2: 1.989 } },
      { close: true },
    ],
  });
  T(s, '[image]', { x: 18.4, y: 7.2, w: 3, h: 0.6, align: 'center', valign: 'middle', fontSize: 18, color: '8C9196' });
  logo(s, 23.13, 6.491, 2.685, 2.018, C.purple);
  logo(s, 2.221, 2.285, 0.922, 0.693, C.purple);
  T(s, 'HR Training', { x: 3.377, y: 2.379, w: 1.981, h: 0.505, fontSize: 24, wrap: false });
  T(s, empower(54), { x: 2.145, y: 3.433, w: 10.215, h: 5.554 });
  T(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.',
    { x: 2.145, y: 9.502, w: 10.242, h: 1.098, fontSize: 24, color: C.dim });
  T(s, 'www.site.com', { x: 2.145, y: 12.143, w: 2.602, h: 0.572, fontSize: 28, color: C.yellow, wrap: false });
});

/** Glass pill with a circled arrow — slides 13 and 14. */
function glassPill(s, x, y, label) {
  s.addShape('roundRect', { x, y, w: 5.251, h: 1.036, rectRadius: 0.518, fill: GLASS_FILL, line: GLASS_LINE });
  T(s, label, { x, y, w: 5.251, h: 1.036, fontSize: 28, color: C.pale, align: 'right', valign: 'middle' });
  arrowCircle(s, x + 0.258, y + 0.186, 0.664, C.white);
}

// -- 13 -------------------------------------------------------- ring + pills
S((s) => {
  T(s, [
    R('Unlock Your Team\u2019s ', { fontFace: FONT, fontSize: 66, color: C.white }),
    R('Potential', { fontFace: FONT, fontSize: 66, color: C.purple }),
    R(' with Expert HR Training!', { fontFace: FONT, fontSize: 66, color: C.white }),
  ], { x: 6.868, y: 1.445, w: 12.93, h: 2.322, align: 'center' });
  photo(s, 8.833, 4.555, 9, 9, { shape: 'ellipse', labelSize: 30 });
  s.addShape('ellipse', { x: 11.921, y: 7.643, w: 2.824, h: 2.824, fill: { color: C.bg } });
  logo(s, 12.722, 8.596, 1.223, 0.919, C.purple);
  glassPill(s, 3.777, 5.910, 'Customize Programs');
  glassPill(s, 2.925, 7.734, 'Customize Programs');
  glassPill(s, 17.638, 5.910, 'Experienced Trainers');
  glassPill(s, 18.490, 7.734, 'Experienced Trainers');
  [[4.701, 'right'], [18.490, 'left']].forEach(([x, align]) => {
    T(s, [R('+188K ', { fontFace: FONT, fontSize: 66, bold: true, color: C.purple }),
          R('Class', { fontFace: FONT, fontSize: 66, bold: true, color: C.white })],
      { x, y: 9.909, w: 3.475, h: 2.293, align, valign: 'middle' });
  });
});

// -- 14 -------------------------------------------------------- phone mock + pills
S((s) => {
  ring(s, 3.909, 7.548, 18.848, 0.177);
  T(s, empower(54), { x: 3.909, y: 2.829, w: 18.848, h: 2.827, align: 'center' });
  logo(s, 11.765, 1.642, 0.922, 0.693, C.purple);
  T(s, 'HR Training', { x: 12.921, y: 1.736, w: 1.981, h: 0.505, fontSize: 24, wrap: false });
  glassPill(s, 3.777, 8.430, 'Customize Programs');
  glassPill(s, 2.925, 10.254, 'Customize Programs');
  glassPill(s, 17.638, 8.430, 'Experienced Trainers');
  glassPill(s, 18.490, 10.254, 'Experienced Trainers');
  s.addShape('roundRect', { x: 10.220, y: 6.953, w: 6.269, h: 12.818, rectRadius: 0.7, fill: { color: '0E2137' }, line: { color: '3F8BD6', width: 1.5 } });
  photo(s, 10.387, 7.120, 5.935, 12.484, { rectRadius: 0.6, labelSize: 24 });
});

// -- 15 -------------------------------------------------------- five phones
S((s) => {
  T(s, [
    R('Empower', { fontFace: FONT, fontSize: 54, color: C.purple }),
    R(' your organization with ', { fontFace: FONT, fontSize: 54, color: C.white }),
    R('our tailored HR training solutions.', { fontFace: FONT, fontSize: 54, color: C.white, transparency: 84 }),
  ], { x: 6.583, y: 1.772, w: 13.5, h: 1.919, align: 'center' });
  [[5.015, 5.762, 2.109, 4.217, 5.167, 5.913, 1.807, 3.791],
   [7.962, 4.994, 2.835, 5.670, 8.190, 5.224, 2.434, 5.111],
   [11.583, 4.443, 3.500, 7.000, 11.839, 4.700, 2.989, 6.486],
   [15.842, 4.994, 2.835, 5.670, 16.064, 5.224, 2.434, 5.111],
   [19.543, 5.762, 2.109, 4.217, 19.692, 5.913, 1.807, 3.791]]
    .forEach(([fx, fy, fw, fh, sx, sy, sw, sh]) => {
      s.addShape('roundRect', { x: fx, y: fy, w: fw, h: fh, rectRadius: fw * 0.16, fill: { color: '0E2137' }, line: { color: '3F8BD6', width: 1.25 } });
      photo(s, sx, sy, sw, sh, { rectRadius: sw * 0.13, labelSize: 16 });
    });
  T(s, 'PLACEHOLDER',
    { x: 8.19, y: 12.131, w: 10.287, h: 1.098, fontSize: 24, align: 'center', lineSpacingMultiple: 1.3 });
});

// -- 16 -------------------------------------------------------- percentage blocks
S((s) => {
  headline(s, 1.856);
  [[11.036, 6.533, 5.344, C.purple, '80%', 80], [8.120, 9.395, 3.749, C.orange, '20%', 32],
   [15.242, 9.839, 3.305, C.yellow, '30%', 36], [9.627, 5.182, 2.377, C.gray, '40%', 36]]
    .forEach(([x, y, d, color, label, size]) => {
      s.addShape('rect', { x, y, w: d, h: d, fill: { color } });
      T(s, label, { x, y, w: d, h: d, fontSize: size, align: 'center', valign: 'middle' });
    });
  [[18.036, 11.270], [15.851, 7.449], [7.405, 5.925], [5.268, 10.863]].forEach(([x, y]) =>
    s.addShape('line', { x, y, w: 1.584, h: 0, line: { color: C.dim, transparency: 50, width: 1 } }));
  [['Definition', 1.783, 10.475, 1.783, 11.336], ['Impact', 19.970, 10.765, 19.970, 11.626],
   ['Key Principles', 17.911, 7.037, 17.911, 7.898], ['Best Practices', 2.491, 5.543, 2.491, 6.404]]
    .forEach(([title, tx, ty, bx, by]) => {
      T(s, title, { x: tx, y: ty, w: 3.972, h: 0.707, fontSize: 36 });
      T(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ', { x: bx, y: by, w: 4.913, h: 1.098, fontSize: 24, lineSpacingMultiple: 1.3 });
    });
});

// -- 17 -------------------------------------------------------- bar chart panel
S((s) => {
  headline(s, 1.402);
  s.addShape('rect', { x: 1.927, y: 5.027, w: 11.352, h: 8.571, fill: { color: C.card } });
  s.addShape('line', { x: 3.477, y: 5.339, w: 0, h: 7.129, line: { color: C.soft, width: 1 } });
  s.addShape('line', { x: 3.477, y: 12.468, w: 9.173, h: 0, line: { color: C.soft, width: 1 } });
  // full-height track behind every bar
  [4.126, 5.544, 6.962, 8.380, 9.798, 11.216].forEach((x) =>
    s.addShape('rect', { x, y: 5.975, w: 0.657, h: 6.493, fill: { color: C.bg } }));
  // value bars: [x, y, w, h, colour]
  [[4.147, 10.295, 0.657, 2.173, C.orange], [5.561, 9.695, 0.657, 2.773, C.purple],
   [6.962, 7.749, 0.669, 4.719, C.orange], [8.390, 6.646, 0.647, 5.822, C.purple],
   [9.804, 8.422, 0.651, 4.046, C.orange], [11.218, 9.384, 0.655, 3.084, C.purple]]
    .forEach(([x, y, w, h, color]) => s.addShape('rect', { x, y, w, h, fill: { color } }));
  ['2019', '2020', '2021', '2022', '2023', '2024'].forEach((year, i) =>
    T(s, year, { x: [3.928, 5.338, 6.752, 8.175, 9.578, 11.013][i], y: 12.60, w: 1.09, h: 0.572, fontSize: 28, align: 'center', wrap: false }));
  ['0', '10', '20', '30', '40', '50', '60', '70', '80', '90', '100'].forEach((v, i) =>
    T(s, v, { x: 2.115, y: 12.204 - i * 0.634, w: 1.258, h: 0.572, fontSize: 28, align: 'right' }));
  s.addShape('rect', { x: 14.028, y: 5.027, w: 10.712, h: 4.031, fill: { color: C.purple } });
  s.addShape('rect', { x: 14.028, y: 9.567, w: 10.712, h: 4.031, fill: { color: C.orange } });
  T(s, [R('Unlock Your Team\u2019s ', { fontFace: FONT, fontSize: 54, color: C.white }),
        R('Potential', { fontFace: FONT, fontSize: 54, color: C.orange }),
        R(' with Expert HR Training!', { fontFace: FONT, fontSize: 54, color: C.white })],
    { x: 15.094, y: 5.629, w: 8.578, h: 2.827 });
  logo(s, 16.354, 10.970, 1.629, 1.224, C.purple);
  T(s, 'HR Training', { x: 18.211, y: 11.077, w: 4.202, h: 1.010, fontSize: 54, wrap: false });
});

// -- 18 -------------------------------------------------------- funnel arrows
S((s) => {
  headline(s, 2.026);
  [[16.854, 5.641, 7.333, C.purple], [17.520, 6.308, 6.000, C.orange], [18.381, 7.169, 4.278, C.yellow]]
    .forEach(([x, y, d, color]) => s.addShape('roundRect', { x, y, w: d, h: d, rectRadius: d * 0.32, fill: { color } }));
  T(s, 'Influencer Webinarin Digital Marketing', { x: 18.381, y: 7.169, w: 4.278, h: 4.278, fontSize: 32, align: 'center', valign: 'middle', color: '000000' });
  [[14.436, 6.655, 3.084, C.purple], [15.245, 10.044, 3.084, C.orange], [12.554, 8.433, 6.000, C.yellow]]
    .forEach(([x, y, w, color]) => s.addShape('rightArrow', { x, y, w, h: 1.528, rotate: 180, fill: { color } }));
  [['200k', C.purple, 5.917, 6.085, null], ['220k', C.yellow, 8.517, 8.685, C.yellow], ['190k', C.orange, 11.116, 11.284, C.orange]]
    .forEach(([value, color, ny, gy, titleColor]) => {
      T(s, value, { x: 2.480, y: ny, w: 3.892, h: 1.582, fontSize: 88, bold: true, color, align: 'right' });
      T(s, [R('Text Tittle Here', { fontFace: FONT, fontSize: 28, bold: true, italic: true, color: titleColor || C.purple })],
        { x: 6.577, y: gy, w: 3.477, h: 0.651 });
      T(s, 'Ut wisi enim ad minim veniam, quis', { x: 6.577, y: gy + 0.673, w: 5.895, h: 0.573, fontSize: 24, color: C.ink, lineSpacingMultiple: 1.3 });
    });
});

// -- 19 -------------------------------------------------------- gears
S((s) => {
  headline(s, 2.136);
  [[2.834, C.purple, 'doc', 10.557], [8.497, C.orange, 'monitor', 10.441],
   [14.170, C.purple, 'gear', 10.441], [19.833, C.orange, 'people', 10.441]]
    .forEach(([x, color, kind, ty]) => {
      gear(s, x, 5.790, 4.0, 12, color);
      s.addShape('ellipse', { x: x + 0.855, y: 6.603, w: 2.407, h: 2.407, fill: { color: C.bg } });
      icon(s, kind, x + 1.554, 7.302, 1.01, color);
      T(s, 'Your Title Text', { x: x - 0.196, y: ty, w: 4.398, h: 0.573, fontSize: 24, bold: true, italic: true, color, align: 'center', lineSpacingMultiple: 1.3 });
      T(s, 'A wonderful serenity has taken possession of my entire soul',
        { x: x - 0.196, y: 11.24, w: 4.398, h: 1.623, fontSize: 24, align: 'center', lineSpacingMultiple: 1.3 });
    });
});

// -- 20 -------------------------------------------------------- stacked layers
S((s) => {
  headline(s, 1.797);
  [[9.592, C.gray], [8.555, C.yellow], [7.518, C.orange]].forEach(([y, color]) =>
    s.addShape('roundRect', { x: 10.112, y, w: 6.171, h: 3.612, rectRadius: 1.806, fill: { color } }));
  s.addShape('roundRect', { x: 10.129, y: 4.958, w: 6.137, h: 5.120, rectRadius: 2.560, fill: { color: C.purple } });
  T(s, '\u201Cum sociis natoque penatibus et magnis dis parturient montes\u201D.',
    { x: 10.129, y: 4.958, w: 6.137, h: 5.120, fontSize: 36, align: 'center', valign: 'middle' });
  // [number, colour, numX, numY, lineX, lineY, titleX, titleY, titleAlign, bodyX, bodyAlign]
  [['01', C.purple, 18.560, 5.137, 16.032, 6.900, 18.560, 6.551, 'left', 18.560, 'left', 7.284],
   ['02', C.orange, 6.702, 5.137, 8.357, 6.900, 4.891, 6.551, 'right', 2.279, 'right', 7.284],
   ['03', C.gray, 6.702, 9.544, 8.123, 11.307, 4.891, 10.958, 'left', 2.279, 'right', 11.691],
   ['04', C.yellow, 18.560, 9.544, 16.177, 11.266, 18.560, 10.958, 'left', 18.560, 'left', 11.691]]
    .forEach(([num, color, nx, ny, lx, ly, tx, ty, tAlign, bx, bAlign, by]) => {
      T(s, num, { x: nx, y: ny, w: 1.405, h: 1.313, fontSize: 72, bold: true, color, wrap: false, align: nx < 12 ? 'right' : 'left' });
      s.addShape('line', { x: lx, y: ly, w: 2.007, h: 0, line: { color, transparency: 60, width: 1.5, dashType: 'sysDot' } });
      T(s, 'Awesome text', { x: tx, y: ty, w: tAlign === 'right' ? 3.215 : 4.718, h: 0.573, fontSize: 24, color, align: tAlign });
      T(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ', { x: bx, y: by, w: 5.828, h: 1.239, fontSize: 24, align: bAlign, lineSpacingMultiple: 1.5 });
    });
});

// -- 21 -------------------------------------------------------- pricing
S((s) => {
  headline(s, 1.480);
  // [cardX, cardY, cardW, cardH, name, nameY, price, priceY, ruleY, bodyY, bodyH, body]
  [[2.8615, 5.4205, 6.564, 7.565, 'Basic ', 6.230, '$22.99', 7.159, 10.049, 10.430, 1.098, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ', 3.510],
   [17.2415, 5.4205, 6.564, 7.565, 'Enterprise', 6.230, '$55.99', 7.159, 10.049, 10.430, 1.098, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ', 17.890],
   [10.0515, 4.8855, 6.564, 8.635, 'Pro', 5.004, '$30.99', 5.932, 8.823, 9.203, 2.148, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. ', 10.700]]
    .forEach(([cx, cy, cw, ch, name, ny, price, py, ry, by, bh, body, tx]) => {
      s.addShape('rect', { x: cx, y: cy, w: cw, h: ch, fill: { color: C.card } });
      T(s, name, { x: tx, y: ny, w: 5.267, h: 0.573, fontSize: 24, italic: true, color: C.purple, align: 'center', lineSpacingMultiple: 1.3 });
      T(s, price, { x: tx, y: py, w: 5.267, h: 1.447, fontSize: 80, color: C.purple, align: 'center' });
      T(s, 'per month', { x: tx, y: py + 1.493, w: 5.267, h: 0.572, fontSize: 28, color: C.muted, align: 'center' });
      s.addShape('line', { x: tx + 0.069, y: ry, w: 5.128, h: 0, line: { color: C.purple, transparency: 60, width: 1 } });
      T(s, body, { x: tx, y: by, w: 5.267, h: bh, fontSize: 24, color: C.muted, align: 'center', lineSpacingMultiple: 1.3 });
    });
  s.addShape('rect', { x: 11.586, y: 11.885, w: 3.495, h: 0.989, fill: { color: C.purple } });
  T(s, 'Best Deal', { x: 11.586, y: 11.885, w: 3.495, h: 0.989, fontSize: 36, align: 'center', valign: 'middle' });
  s.addShape('ellipse', { x: 15.338, y: 4.508, w: 1.871, h: 1.871, fill: { color: C.white } });
  icon(s, 'thumb', 15.808, 4.923, 1.04, C.yellow);
});

// -- 22 -------------------------------------------------------- four arrows
S((s) => {
  headline(s, 1.665);
  // rightArrow spun to each compass point, plus its white glyph
  [[10.232, 5.325, 270, C.purple], [13.301, 6.431, 0, C.orange],
   [12.264, 9.496, 90, C.yellow], [9.194, 8.389, 180, C.gray]]
    .forEach(([x, y, rot, color]) => s.addShape('rightArrow', { x, y, w: 4.171, h: 3.507, rotate: rot, fill: { color } }));
  s.addShape('roundRect', { x: 12.723, y: 8.554, w: 1.220, h: 1.220, rectRadius: 0.4, fill: { color: C.card } });
  icon(s, 'globe', 11.900, 6.243, 0.833, C.white);
  icon(s, 'cloud', 15.044, 7.702, 0.870, C.white);
  icon(s, 'briefcase', 13.901, 10.804, 0.833, C.white);
  icon(s, 'usb', 10.703, 9.726, 0.800, C.white);
  [[2.353, 6.178, 'right'], [2.353, 9.691, 'right'], [18.232, 6.210, 'left'], [18.232, 9.659, 'left']]
    .forEach(([x, y, align]) => {
      T(s, 'Awesome text', { x: align === 'right' ? x + 1.158 : x, y, w: 4.923, h: 0.707, fontSize: 36, bold: true, align });
      T(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ', { x, y: y + 0.861, w: 6.082, h: 1.239, fontSize: 24, align, lineSpacingMultiple: 1.5 });
    });
});

// -- 23 -------------------------------------------------------- SWOT
S((s) => {
  headline(s, 1.444);
  [[3.636, C.purple, 'Strength', C.white], [8.450, C.orange, 'Weakness', C.white],
   [13.264, C.yellow, 'Opportunity', '000000'], [18.078, C.gray, 'Threats ', C.white]]
    .forEach(([x, color, title, ink]) => {
      s.addShape('rect', { x, y: 5.151, w: 4.952, h: 4.952, fill: { color } });
      T(s, title, { x: x + 0.364, y: 6.426, w: 4.225, h: 0.64, fontSize: 32, bold: true, color: ink, align: 'center' });
      T(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ',
        { x: x + 0.636, y: 7.165, w: 3.680, h: 1.700, fontSize: 22, color: ink, align: 'center', lineSpacingMultiple: 1.5 });
    });
  T(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. ',
    { x: 6.979, y: 11.140, w: 12.708, h: 1.264, fontSize: 28, align: 'center', lineSpacingMultiple: 1.3 });
  T(s, 'A wonderful serenity has taken possession',
    { x: 9.356, y: 12.983, w: 7.954, h: 0.573, fontSize: 24, italic: true, color: C.purple, align: 'center', lineSpacingMultiple: 1.3 });
});

/**
 * Flat vector figure holding a laptop (slide 24). Rebuilt from primitives laid
 * out on a 0..1 grid covering the box the original artwork occupied; parts are
 * painted back-to-front: [dx, dy, dw, dh, colour, shape, rotate].
 */
const FIGURE = [
  [0.446, 0.000, 0.349, 0.103, '191816', 'roundRect'],   // hair
  [0.462, 0.063, 0.317, 0.132, 'F7C289', 'roundRect'],   // face
  [0.551, 0.161, 0.120, 0.078, 'F2AF6B'],                // neck
  [0.188, 0.223, 0.812, 0.249, 'CC9B2F', 'roundRect'],   // shirt
  [0.605, 0.207, 0.110, 0.058, 'B78727', 'triangle', 180], // left collar
  [0.506, 0.205, 0.100, 0.058, 'B78727', 'triangle', 180], // right collar
  [0.598, 0.235, 0.014, 0.247, 'B78727'],                // placket
  [0.681, 0.286, 0.117, 0.011, '191816'],                // name badge
  [0.676, 0.313, 0.127, 0.061, 'B78727'],                // left pocket
  [0.401, 0.312, 0.127, 0.062, 'B78727'],                // right pocket
  [0.409, 0.471, 0.371, 0.013, '191816'],                // belt
  [0.564, 0.471, 0.058, 0.014, '383734'],                // buckle
  [0.364, 0.478, 0.461, 0.483, 'CC9B2F'],                // trousers
  [0.575, 0.478, 0.027, 0.478, 'B78727'],                // trouser seam
  [0.624, 0.948, 0.371, 0.052, '191816', 'roundRect'],   // left shoe
  [0.188, 0.948, 0.371, 0.052, '191816', 'roundRect'],   // right shoe
  [0.000, 0.286, 0.532, 0.126, '565554', 'parallelogram', 12], // laptop lid
  [0.215, 0.338, 0.064, 0.026, 'E0DEDE', 'ellipse'],     // lid badge
  [0.076, 0.412, 0.603, 0.011, '848383', 'roundRect'],   // laptop base
  [0.509, 0.378, 0.466, 0.068, 'F7C289', 'roundRect'],   // forearm
  [0.326, 0.403, 0.150, 0.035, 'F7C289', 'roundRect'],   // hand
];

function presenter(s, x, y, w, h) {
  FIGURE.forEach(([dx, dy, dw, dh, color, shape, rot]) => s.addShape(shape || 'rect', Object.assign(
    { x: x + dx * w, y: y + dy * h, w: dw * w, h: dh * h, fill: { color } }, rot ? { rotate: rot } : {}
  )));
}

// -- 24 -------------------------------------------------------- presenter
S((s) => {
  s.addShape('roundRect', { x: 13.811, y: 3.500, w: 8.096, h: 7.999, rectRadius: 2.6, fill: { color: C.card } });
  s.addShape('roundRect', { x: 13.339, y: 3.972, w: 7.152, h: 7.999, rectRadius: 2.6, fill: { color: C.orange } });
  s.addShape('roundRect', { x: 15.481, y: 5.170, w: 2.378, h: 4.756, rectRadius: 1.1, fill: { color: C.purple } });
  presenter(s, 14.368, 1.989, 4.557, 11.023);
  [['01', C.purple, 4.433, 'Online Meeting'], ['02', C.orange, 6.780, 'Online Class'], ['03', C.yellow, 9.127, 'Online Test']]
    .forEach(([num, color, y, label]) => {
      s.addShape('roundRect', { x: 19.097, y, w: 1.441, h: 1.441, rectRadius: 0.42, fill: { color } });
      T(s, num, { x: 19.097, y, w: 1.441, h: 1.441, fontSize: 36, align: 'center', valign: 'middle', color: num === '03' ? '000000' : C.white });
      T(s, label, { x: 21.208, y: y + 0.25, w: 3.415, h: 0.722, fontSize: 28, bold: true, lineSpacingMultiple: 1.5 });
    });
  T(s, empower(54), { x: 2.044, y: 2.435, w: 11.283, h: 4.645 });
  T(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. ',
    { x: 2.044, y: 7.739, w: 8.417, h: 1.877, fontSize: 28, lineSpacingMultiple: 1.3 });
  s.addShape('rect', { x: 2.060, y: 10.416, w: 8.812, h: 2.149, fill: { color: C.card } });
  s.addShape('roundRect', { x: 2.809, y: 11.079, w: 0.100, h: 0.823, rectRadius: 0.05, fill: { color: C.purple } });
  T(s, '\u201Cum sociis natoque penatibus et magnis dis parturient montes\u201D.',
    { x: 3.442, y: 10.853, w: 6.680, h: 1.181, fontSize: 28, bold: true, lineSpacingMultiple: 1.2 });
});

// -- 25 -------------------------------------------------------- four circles
S((s) => {
  headline(s, 1.775);
  [[9.473, 5.555, C.purple], [13.331, 5.555, C.orange], [9.473, 9.365, C.yellow], [13.331, 9.365, C.gray]]
    .forEach(([x, y, color]) => s.addShape('ellipse', { x, y, w: 3.860, h: 3.860, fill: { color } }));
  s.addShape('ellipse', { x: 11.736, y: 7.848, w: 3.189, h: 3.189, fill: { color: C.card } });
  icon(s, 'pencil', 12.784, 8.890, 1.10, C.white);
  [['01', C.purple, 2.140, 6.223, 4.252, 6.179], ['04', C.yellow, 2.140, 10.928, 4.252, 10.884],
   ['02', C.orange, 17.543, 6.223, 19.654, 6.179], ['03', C.gray, 17.543, 10.928, 19.654, 10.884]]
    .forEach(([num, color, bx, by, tx, ty]) => {
      s.addShape('rect', { x: bx, y: by, w: 1.441, h: 1.441, fill: { color } });
      T(s, num, { x: bx, y: by, w: 1.441, h: 1.441, fontSize: 36, align: 'center', valign: 'middle' });
      T(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit', { x: tx, y: ty, w: 4.872, h: 1.717, fontSize: 32 });
    });
});

// -- 26 -------------------------------------------------------- social channels
S((s) => {
  ring(s, 12.717, -7.278, 22.278, 0.181);
  T(s, empower(54), { x: 2.811, y: 3.064, w: 10.523, h: 5.554 });
  T(s, 'Data Social Media', { x: 2.811, y: 9.598, w: 6.406, h: 0.64, fontSize: 32, valign: 'middle' });
  T(s, 'Leverage agile frameworks to provide a robust synopsis for high level overviews. Iterative approaches to corporate strategy foster collaborative',
    { x: 2.811, y: 10.313, w: 8.593, h: 1.623, fontSize: 24, lineSpacingMultiple: 1.3 });
  const blurb = 'Leverage agile frameworks to provide a robust synopsis for high level overviews. Iterative approaches to corporate strategy';
  [['Twitter', 1.596, 2.773, 3.488, 'bird'], ['Facebook', 5.658, 6.859, 7.575, 'f'], ['Youtube', 9.774, 10.905, 11.620, 'play']]
    .forEach(([name, iconY, titleY, bodyY, kind]) => {
      if (kind === 'f') {
        s.addShape('roundRect', { x: 14.936, y: iconY, w: 0.860, h: 0.860, rectRadius: 0.25, line: { color: C.purple, width: 2 } });
        T(s, 'f', { x: 14.936, y: iconY, w: 0.860, h: 0.860, fontSize: 32, bold: true, align: 'center', valign: 'middle', color: C.purple });
      } else if (kind === 'play') {
        s.addShape('roundRect', { x: 14.936, y: iconY, w: 0.900, h: 0.691, rectRadius: 0.2, line: { color: C.purple, width: 2 } });
        icon(s, 'play', 15.190, 9.930, 0.40, C.purple);
      } else {
        icon(s, 'bird', 14.936, iconY, 0.860, C.purple);
      }
      T(s, name, { x: 14.751, y: titleY, w: 2.599, h: 0.64, fontSize: 32, valign: 'middle' });
      T(s, blurb, { x: 14.751, y: bodyY, w: 9.105, h: 1.015, fontSize: 22, lineSpacingMultiple: 1.3 });
    });
});

// -- 27 -------------------------------------------------------- puzzle pencil
S((s) => {
  headline(s, 1.868);
  // Four interlocking tiles marching down the diagonal, then the pencil tip.
  [['01', C.purple, 15.538, 6.781], ['02', C.orange, 13.861, 8.458],
   ['03', C.yellow, 12.184, 10.135], ['04', C.gray, 10.507, 11.812]]
    .forEach(([num, color, cx, cy]) => {
      s.addShape('rect', { x: cx - 1.36, y: cy - 1.36, w: 2.72, h: 2.72, rotate: 45, fill: { color } });
      T(s, num, { x: cx - 0.8, y: cy - 0.45, w: 1.6, h: 0.9, fontSize: 45, bold: true, align: 'center', valign: 'middle' });
    });
  s.addShape('triangle', { x: 8.20, y: 12.25, w: 1.88, h: 1.88, rotate: -135, fill: { color: 'FDDCBA' } });
  // [title, titleX, titleY, bodyX, bodyY, align, lineX, lineY, lineW]
  [['Great Score', 18.994, 7.140, 18.994, 7.874, 'left', 17.296, 7.423, 1.516],
   ['Time Is Money', 15.682, 11.133, 15.682, 11.868, 'left', 13.257, 11.416, 2.186],
   ['Prestige Award', 5.436, 6.421, 3.170, 7.155, 'right', 10.621, 6.750, 2.484],
   ['Fresh Graduate', 3.193, 10.091, 0.927, 10.826, 'right', 8.124, 10.422, 1.802]]
    .forEach(([title, tx, ty, bx, by, align, lx, ly, lw]) => {
      T(s, title, { x: tx, y: ty, w: 4.796, h: 0.64, fontSize: 32, bold: true, align });
      T(s, 'A wonderful serenity has taken possession of my entire soul', { x: bx, y: by, w: 7.062, h: 1.264, fontSize: 28, align, lineSpacingMultiple: 1.3 });
      s.addShape('line', { x: lx, y: ly, w: lw, h: 0, line: { color: C.dim, transparency: 55, width: 0.75, endArrowType: 'triangle' } });
    });
});

// -- 28 -------------------------------------------------------- hex cluster
S((s) => {
  // honeycomb: pointy-top hexagons
  [[6.183, 5.594], [9.867, 5.594], [11.674, 2.356], [11.705, 8.832], [13.552, 5.594]]
    .forEach(([x, y]) => s.addShape('hexagon', { x: x - 0.2035, y: y + 0.2035, w: 3.869, h: 3.462, rotate: 90, fill: { color: C.card } }));
  [['mouse', C.gray, 12.685, 3.570], ['wifi', C.purple, 10.878, 6.808],
   ['save', C.yellow, 14.563, 6.808], ['heart', C.orange, 12.716, 10.046]]
    .forEach(([kind, color, x, y]) => {
      s.addShape('roundRect', { x, y, w: 1.441, h: 1.441, rectRadius: 0.42, fill: { color } });
      icon(s, kind, x + 0.38, y + 0.35, 0.72, C.white);
    });
  T(s, '11', { x: 6.504, y: 6.343, w: 2.703, h: 1.717, fontSize: 96, bold: true, align: 'center' });
  T(s, 'Optimizing', { x: 6.504, y: 7.916, w: 2.703, h: 0.505, fontSize: 24, bold: true, align: 'center' });
  [[14.827, 2.878, 3.333, 1.496], [14.926, 10.588, 3.407, 0.725], [17.153, 6.224, 2.359, 1.320],
   [6.334, 10.735, 2.373, 0.671], [8.083, 3.763, 2.428, 1.976]]
    .forEach(([x, y, w, h]) => s.addShape('line', { x, y, w, h, line: { color: C.purple, transparency: 60, width: 1.5, dashType: 'sysDot' } }));
  [['right', 3.093, 2.878, 3.093, 3.565], ['right', 1.911, 10.276, 1.911, 10.964],
   ['left', 18.500, 2.299, 18.500, 2.986], ['left', 18.959, 10.372, 18.959, 11.059],
   ['left', 19.833, 5.471, 19.833, 6.158]]
    .forEach(([align, tx, ty, bx, by]) => {
      T(s, 'Awesome text', { x: tx, y: ty, w: 4.923, h: 0.64, fontSize: 32, align });
      T(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ', { x: bx, y: by, w: 4.923, h: 1.239, fontSize: 24, align, lineSpacingMultiple: 1.5 });
    });
});

// -- 29 -------------------------------------------------------- spline chart
S((s) => {
  headline(s, 1.430);
  s.addShape('line', { x: 4.682, y: 5.177, w: 0, h: 4.888, line: { color: C.soft, width: 1 } });
  s.addShape('line', { x: 4.682, y: 10.066, w: 18.276, h: 0, line: { color: C.soft, width: 1 } });
  spline(s, 5.879, 6.952, 15.095, 3.069, [[0, 0.3478],
    [0.0533, 0.3576], [0.1066, 0.3674], [0.1626, 0.3679], [0.2186, 0.3685], [0.2835, 0.3853], [0.336, 0.3511],
    [0.3886, 0.3169], [0.4377, 0.2178], [0.478, 0.1629], [0.5183, 0.1079], [0.5457, 0.0457], [0.5777, 0.0217],
    [0.6098, -0.0024], [0.6405, -0.0103], [0.6703, 0.0183], [0.7001, 0.0469], [0.7218, 0.0984], [0.7565, 0.1931],
    [0.7913, 0.2878], [0.8381, 0.452], [0.8787, 0.5865], [0.9193, 0.721], [1, 1], [1, 1]], 'FFC000');
  spline(s, 5.913, 5.625, 15.109, 4.464, [[0, 0.0277],
    [0.088, 0.3643], [0.1759, 0.701], [0.237, 0.8596], [0.298, 1.0183], [0.3208, 1.016], [0.3662, 0.9798],
    [0.4117, 0.9436], [0.4674, 0.7541], [0.5099, 0.6424], [0.5524, 0.5307], [0.5817, 0.3967], [0.6212, 0.3097],
    [0.6607, 0.2226], [0.6837, 0.1718], [0.7469, 0.1202], [0.81, 0.0686], [0.905, 0.0343], [1, 0]], '8AA5D7');
  spline(s, 5.869, 5.596, 14.994, 2.659, [[0, 0.7607],
    [0.0693, 0.5473], [0.1386, 0.334], [0.1886, 0.2079], [0.2386, 0.0818], [0.2639, 0.0207], [0.2999, 0.0042],
    [0.3358, -0.0123], [0.3673, 0.0188], [0.4043, 0.109], [0.4414, 0.1992], [0.5224, 0.5454], [0.5224, 0.5454],
    [0.5561, 0.6705], [0.5776, 0.784], [0.6065, 0.8596], [0.6355, 0.9352], [0.661, 0.9925], [0.6961, 0.9992],
    [0.7311, 1.006], [0.7662, 0.9663], [0.8168, 0.9003], [0.8675, 0.8344], [0.9649, 0.6938], [1, 0.6036]], 'EFAC7E');
  [['FFC000', [[5.684, 7.878], [10.895, 7.899], [15.748, 6.882], [20.670, 9.824]]],
   ['8AA5D7', [[5.684, 5.578], [10.895, 9.922], [15.748, 6.461], [20.665, 5.487]]],
   [C.orange, [[5.671, 7.457], [10.895, 5.499], [15.748, 8.072], [20.653, 7.100]]]]
    .forEach(([color, dots]) => dots.forEach(([x, y]) =>
      s.addShape('ellipse', { x, y, w: 0.387, h: 0.294, fill: { color } })));
  ['2014', '2016', '2018', '2020', '2022', '2024', '2026', '2028', '2030', '2032'].forEach((year, i) =>
    T(s, year, { x: 4.325 + i * 2.031, y: 10.423, w: 0.80, h: 0.438, fontSize: 20, color: C.pale, align: 'center', wrap: false }));
  ['5.0', '4.0', '3.0', '2.0', '1.0', '0.0'].forEach((v, i) =>
    T(s, v, { x: 3.34, y: 4.827 + i * 0.956, w: 0.57, h: 0.438, fontSize: 20, color: C.pale, align: 'center', wrap: false }));
  T(s, 'A wonderful serenity has taken possession of my entire soul  \u2933',
    { x: 4.070, y: 11.725, w: 9.941, h: 0.579, fontSize: 24, bold: true, italic: true, color: C.purple, lineSpacingMultiple: 1.3, wrap: false });
  T(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Et dolor diam ultricies sed quisque. Tortor cursus sed blandit..',
    { x: 15.786, y: 11.725, w: 6.811, h: 1.845, fontSize: 24, lineSpacingMultiple: 1.5 });
});

// -- 30 -------------------------------------------------------- closing
S((s) => {
  ring(s, 2.194, -3.639, 22.278, 0.134);
  logo(s, 11.284, 2.367, 1.560, 1.172, C.purple);
  T(s, 'HR Training', { x: 13.133, y: 2.667, w: 2.250, h: 0.572, fontSize: 28, align: 'right', wrap: false });
  T(s, 'Thank You', { x: 6.552, y: 5.775, w: 13.562, h: 3.45, fontSize: 199, align: 'center', wrap: false });
  T(s, 'www.site.com', { x: 12.032, y: 12.061, w: 2.602, h: 0.572, fontSize: 28, align: 'right', wrap: false });
});

// ================================================================= build
const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'HR', width: W, height: H });
pptx.layout = 'HR';
pptx.theme = { headFontFace: FONT, bodyFontFace: FONT };
pptx.title = 'HR Training';

slides.forEach((build, i) => {
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  build(slide);
  chrome(slide, i + 1);
});

pptx.writeFile({ fileName: path.join(__dirname, '108af3ff-3f8e-4070-9257-1847ac54b1e4_grok_final.pptx') })
  .then((f) => console.log('wrote', f));
