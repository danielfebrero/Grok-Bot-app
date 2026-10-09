/**
 * AVDIJA — Sport Presentation Template (36 slides, 13.333 x 7.5 in)
 *
 * Recreated with pptxgenjs. Photographic content in the original deck is
 * replaced by flat "[image]" placeholder rectangles of matching geometry.
 *
 *   node <this file>   ->  writes the .pptx next to this script
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette
const ORANGE  = 'FE712F';   // brand accent
const ORANGE2 = 'FF6600';   // slightly different orange used on slide 32 arcs
const BLACK   = '040404';
const DARK    = '292929';
const WHITE   = 'FFFFFF';
const INK     = '000000';
const INK2    = '0C0C0C';
const GREY    = '808080';
const GREY2   = '404040';
const SILVER  = 'F2F2F2';
const NAVY    = '072C62';
const RULE    = '6B4D9D';
const PLACEHOLDER_FILL = 'D9D9D9';

// ---------------------------------------------------------------- typography
const BODY   = 'Poppins';
const MED    = 'Poppins Medium';
const SEMI   = 'Poppins SemiBold';
const OSSEMI = 'Open Sans SemiBold';

// Text-box insets of the source deck: 0.1" left/right, 0.05" top/bottom,
// expressed the way pptxgenjs wants them - points, ordered [L, R, B, T].
const INSET = [7.2, 7.2, 3.6, 3.6];

function newSlide (pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  return s;
}

/**
 * Solid-filled decorative geometry (banners, stripes, cards, pills, ...).
 *
 * `adj` is the shape's first adjust handle as a fraction of its shorter side:
 * the corner radius of a roundRect, the shear of a parallelogram, and so on.
 * `arc` / `thick` drive blockArc (start+sweep in degrees, ring thickness).
 */
function deco (s, shape, x, y, w, h, o) {
  o = o || {};
  const opts = { x, y, w, h };
  opts.fill = o.fill ? { color: o.fill, transparency: o.alpha || 0 } : { type: 'none' };
  opts.line = o.line ? { color: o.line, width: o.lineW || 1 } : { type: 'none' };
  if (o.rotate) opts.rotate = o.rotate;
  if (o.flipH) opts.flipH = true;
  if (o.flipV) opts.flipV = true;
  if (o.adj) opts.rectRadius = o.adj * Math.min(w, h);
  if (o.arc) { opts.angleRange = o.arc; opts.arcThicknessRatio = o.thick; }
  s.addShape(shape, opts);
}

/** A single-format text box. */
function text (s, x, y, w, h, str, sz, face, colour, o) {
  o = o || {};
  s.addText(str, {
    x, y, w, h,
    fontSize: sz, fontFace: face, color: colour,
    align: o.align || 'left', valign: 'top',
    lineSpacingMultiple: o.lh,
    margin: INSET,
    wrap: o.wrap !== false,
    isTextBox: true,
    fit: 'resize',
  });
}

/** A text box whose runs carry individual colours / faces / line breaks. */
function richText (s, x, y, w, h, runs, o) {
  o = o || {};
  s.addText(runs.map(([str, r]) => ({
    text: str,
    options: {
      fontSize: r.sz, fontFace: r.f, color: r.col,
      bold: !!r.b, italic: !!r.i, breakLine: !!r.br,
      align: o.align || 'left', lineSpacingMultiple: o.lh,
    },
  })), {
    x, y, w, h,
    align: o.align || 'left', valign: 'top',
    lineSpacingMultiple: o.lh,
    margin: INSET,
    wrap: o.wrap !== false,
    isTextBox: true,
    fit: 'resize',
  });
}

/** Stand-in for a photo in the original deck. */
function photo (s, x, y, w, h) {
  s.addShape('rect', { x, y, w, h, fill: { color: PLACEHOLDER_FILL }, line: { type: 'none' } });
  s.addText('[image]', {
    x, y, w, h, fontSize: 12, fontFace: BODY, color: '8A8A8A',
    align: 'center', valign: 'middle',
  });
}

/** Free vector outline (cubic beziers + lines), coordinates relative to x/y. */
function freeform (s, x, y, w, h, colour, points) {
  s.addShape('custGeom', {
    x, y, w, h, points,
    fill: { color: colour }, line: { type: 'none' },
  });
}
/* ------------------------------------------------------------------ icons
 * The template uses flat pictogram glyphs. Each one is rebuilt here from
 * native PowerPoint shapes inside a 0..1 unit box:
 *
 *     ['<prst>',  x, y, w, h, { ... }]   a preset shape
 *     ['poly',    [[x, y], ...]        ] a filled polygon
 *     ['stroke',  [[x, y], ...],  wdt  ] a poly-line
 *
 * `bg: true` paints with the colour behind the icon (used for cut-outs).
 */
const ICON_ART = {
  target: [
    ['ellipse', 0.00, 0.14, 0.78, 0.78],
    ['ellipse', 0.10, 0.24, 0.58, 0.58, { bg: true }],
    ['ellipse', 0.19, 0.33, 0.40, 0.40],
    ['ellipse', 0.27, 0.41, 0.24, 0.24, { bg: true }],
    ['ellipse', 0.33, 0.47, 0.12, 0.12],
    ['stroke', [[0.39, 0.53], [0.80, 0.12]], 0.11],
    ['poly', [[0.72, 0.00], [1.00, 0.00], [1.00, 0.28], [0.93, 0.14], [0.86, 0.07]]],
  ],
  stopwatch: [
    ['rect', 0.34, 0.00, 0.32, 0.09],
    ['rect', 0.43, 0.07, 0.14, 0.14],
    ['stroke', [[0.72, 0.10], [0.86, 0.22]], 0.09],
    ['ellipse', 0.02, 0.16, 0.96, 0.84],
    ['ellipse', 0.13, 0.27, 0.74, 0.62, { bg: true }],
    ['ellipse', 0.20, 0.34, 0.60, 0.48],
    ['ellipse', 0.26, 0.40, 0.48, 0.36, { bg: true }],
    ['stroke', [[0.50, 0.58], [0.68, 0.46]], 0.06],
    ['ellipse', 0.45, 0.53, 0.10, 0.10],
  ],
  shoe: [
    ['poly', [[0.02, 0.60], [0.10, 0.40], [0.30, 0.28], [0.52, 0.24], [0.64, 0.08],
              [0.82, 0.06], [0.98, 0.34], [1.00, 0.66], [0.02, 0.66]]],
    ['roundRect', 0.00, 0.58, 1.00, 0.24, { r: 0.4 }],
    ['rect', 0.04, 0.82, 0.92, 0.07],
    ['stroke', [[0.34, 0.30], [0.44, 0.56]], 0.05, { bg: true }],
    ['stroke', [[0.47, 0.28], [0.56, 0.56]], 0.05, { bg: true }],
    ['stroke', [[0.62, 0.22], [0.70, 0.56]], 0.05, { bg: true }],
  ],
  book: [
    ['can', 0.00, 0.00, 0.30, 1.00],
    ['rect', 0.30, 0.00, 0.44, 1.00],
    ['rect', 0.80, 0.00, 0.20, 1.00],
    ['ellipse', 0.05, 0.72, 0.20, 0.16, { bg: true }],
  ],
  trophy: [
    ['ellipse', 0.00, 0.14, 0.26, 0.30, { hollow: 0.065 }],
    ['ellipse', 0.74, 0.14, 0.26, 0.30, { hollow: 0.065 }],
    ['rect', 0.12, 0.04, 0.76, 0.13],
    ['pie', 0.16, -0.30, 0.68, 0.98, { angles: [0, 180] }],
    ['rect', 0.44, 0.56, 0.12, 0.18],
    ['rect', 0.28, 0.72, 0.44, 0.09],
    ['roundRect', 0.14, 0.87, 0.72, 0.13, { r: 0.05 }],
    ['star5', 0.35, 0.22, 0.30, 0.30, { bg: true }],
  ],
  sunrise: [
    ['pie', 0.26, 0.10, 0.48, 0.66, { angles: [180, 360] }],
    ['stroke', [[0.50, 0.00], [0.50, 0.07]], 0.05],
    ['stroke', [[0.32, 0.04], [0.36, 0.12]], 0.05],
    ['stroke', [[0.68, 0.04], [0.64, 0.12]], 0.05],
    ['stroke', [[0.16, 0.16], [0.22, 0.22]], 0.05],
    ['stroke', [[0.84, 0.16], [0.78, 0.22]], 0.05],
    ['stroke', [[0.06, 0.36], [0.15, 0.36]], 0.05],
    ['stroke', [[0.94, 0.36], [0.85, 0.36]], 0.05],
    ['wave', 0.00, 0.50, 1.00, 0.09],
    ['wave', 0.00, 0.65, 1.00, 0.09],
    ['wave', 0.00, 0.80, 1.00, 0.09],
    ['wave', 0.00, 0.95, 1.00, 0.09],
  ],
  whistle: [
    ['roundRect', 0.30, 0.10, 0.70, 0.80, { r: 0.22 }],
    ['ellipse', 0.52, 0.28, 0.42, 0.42, { bg: true }],
    ['ellipse', 0.60, 0.36, 0.26, 0.26],
    ['poly', [[0.00, 0.30], [0.34, 0.24], [0.34, 0.62], [0.00, 0.56]]],
  ],
  cloud: [
    ['ellipse', 0.10, 0.16, 0.44, 0.44],
    ['ellipse', 0.40, 0.04, 0.50, 0.50],
    ['ellipse', 0.00, 0.32, 0.36, 0.34],
    ['rect', 0.10, 0.36, 0.80, 0.24],
    ['stroke', [[0.16, 0.78], [0.70, 0.78]], 0.09],
    ['ellipse', 0.62, 0.62, 0.30, 0.34, { hollow: 0.09 }],
    ['rect', 0.62, 0.62, 0.32, 0.16, { bg: true }],
  ],
  bottle: [
    ['rect', 0.32, 0.00, 0.36, 0.10],
    ['rect', 0.38, 0.09, 0.24, 0.09],
    ['roundRect', 0.08, 0.17, 0.84, 0.83, { r: 0.18 }],
    ['rect', 0.08, 0.36, 0.84, 0.05, { bg: true }],
    ['rect', 0.08, 0.60, 0.84, 0.05, { bg: true }],
    ['lightningBolt', 0.38, 0.42, 0.24, 0.17, { bg: true }],
  ],
  masks: [
    ['ellipse', 0.00, 0.14, 0.48, 0.60],
    ['poly', [[0.03, 0.58], [0.45, 0.58], [0.24, 0.86]]],
    ['ellipse', 0.09, 0.34, 0.08, 0.10, { bg: true }],
    ['ellipse', 0.31, 0.34, 0.08, 0.10, { bg: true }],
    ['pie', 0.12, 0.46, 0.24, 0.18, { angles: [180, 0], bg: true }],
    ['ellipse', 0.52, 0.04, 0.48, 0.60],
    ['poly', [[0.55, 0.48], [0.97, 0.48], [0.76, 0.76]]],
    ['ellipse', 0.61, 0.24, 0.08, 0.10, { bg: true }],
    ['ellipse', 0.83, 0.24, 0.08, 0.10, { bg: true }],
    ['pie', 0.64, 0.34, 0.24, 0.18, { angles: [0, 180], bg: true }],
  ],
  camera: [
    ['rect', 0.20, 0.00, 0.30, 0.16],
    ['roundRect', 0.00, 0.12, 1.00, 0.88, { r: 0.09 }],
    ['rect', 0.08, 0.22, 0.14, 0.12, { bg: true }],
    ['ellipse', 0.30, 0.30, 0.42, 0.42, { bg: true }],
    ['ellipse', 0.37, 0.37, 0.28, 0.28],
    ['ellipse', 0.43, 0.43, 0.16, 0.16, { bg: true }],
  ],
  rocket: [
    ['teardrop', 0.22, 0.04, 0.78, 0.78],
    ['ellipse', 0.56, 0.22, 0.18, 0.18, { bg: true }],
    ['poly', [[0.08, 0.44], [0.40, 0.54], [0.30, 0.72]]],
    ['poly', [[0.46, 0.94], [0.42, 0.60], [0.26, 0.78]]],
    ['stroke', [[0.00, 0.98], [0.10, 0.88]], 0.11],
    ['stroke', [[0.02, 0.66], [0.09, 0.59]], 0.11],
    ['stroke', [[0.36, 1.00], [0.43, 0.93]], 0.11],
  ],
  envelope: [
    ['rect', 0.00, 0.00, 1.00, 1.00],
    ['stroke', [[0.10, 0.14], [0.50, 0.54]], 0.10, { bg: true }],
    ['stroke', [[0.50, 0.54], [0.90, 0.14]], 0.10, { bg: true }],
  ],
  megaphone: [
    ['poly', [[0.34, 0.30], [1.00, 0.00], [1.00, 1.00], [0.34, 0.70]]],
    ['roundRect', 0.00, 0.26, 0.42, 0.48, { r: 0.1 }],
    ['roundRect', 0.10, 0.70, 0.22, 0.30, { r: 0.07 }],
  ],
  medkit: [
    ['rect', 0.32, 0.00, 0.36, 0.16],
    ['rect', 0.40, 0.05, 0.20, 0.11, { bg: true }],
    ['roundRect', 0.00, 0.14, 1.00, 0.86, { r: 0.08 }],
    ['mathPlus', 0.30, 0.36, 0.40, 0.40, { bg: true }],
  ],
  globe: [
    ['ellipse', 0.00, 0.00, 1.00, 1.00],
    ['ellipse', 0.28, 0.00, 0.44, 1.00, { hollow: 0.06, bg: true }],
    ['ellipse', 0.00, 0.28, 1.00, 0.44, { hollow: 0.06, bg: true }],
    ['stroke', [[0.02, 0.50], [0.98, 0.50]], 0.06, { bg: true }],
    ['stroke', [[0.50, 0.02], [0.50, 0.98]], 0.06, { bg: true }],
  ],
  globePin: [
    ['ellipse', 0.00, 0.00, 0.86, 0.86],
    ['ellipse', 0.24, 0.00, 0.38, 0.86, { hollow: 0.06, bg: true }],
    ['stroke', [[0.02, 0.43], [0.84, 0.43]], 0.06, { bg: true }],
    ['teardrop', 0.44, 0.44, 0.56, 0.56, { rotate: 135 }],
    ['ellipse', 0.62, 0.62, 0.20, 0.20, { bg: true }],
  ],
  care: [
    ['ellipse', 0.38, 0.02, 0.24, 0.24],
    ['poly', [[0.29, 0.68], [0.37, 0.26], [0.63, 0.26], [0.71, 0.68]]],
    ['blockArc', -0.10, 0.24, 0.50, 0.76, { arc: [340, 170], thick: 0.5 }],
    ['blockArc', 0.60, 0.24, 0.50, 0.76, { arc: [10, 200], thick: 0.5 }],
  ],
  people: [
    ['ellipse', 0.03, 0.34, 0.15, 0.15],
    ['poly', [[0.00, 0.78], [0.04, 0.52], [0.19, 0.52], [0.23, 0.78]]],
    ['ellipse', 0.82, 0.34, 0.15, 0.15],
    ['poly', [[0.77, 0.78], [0.81, 0.52], [0.96, 0.52], [1.00, 0.78]]],
    ['roundRect', 0.18, 0.00, 0.34, 0.24, { r: 0.07 }],
    ['poly', [[0.24, 0.22], [0.36, 0.22], [0.25, 0.34]]],
    ['roundRect', 0.58, 0.14, 0.28, 0.20, { r: 0.07 }],
    ['poly', [[0.70, 0.32], [0.80, 0.32], [0.79, 0.44]]],
    ['rect', 0.26, 0.62, 0.48, 0.07],
    ['rect', 0.29, 0.69, 0.05, 0.29],
    ['rect', 0.66, 0.69, 0.05, 0.29],
    ['rect', 0.00, 0.84, 0.23, 0.07],
    ['rect', 0.03, 0.91, 0.05, 0.09],
    ['rect', 0.15, 0.91, 0.05, 0.09],
    ['rect', 0.77, 0.84, 0.23, 0.07],
    ['rect', 0.80, 0.91, 0.05, 0.09],
    ['rect', 0.92, 0.91, 0.05, 0.09],
  ],
  handshake: [
    ['stroke', [[0.20, 0.02], [0.26, 0.20]], 0.09],
    ['stroke', [[0.50, 0.00], [0.50, 0.18]], 0.09],
    ['stroke', [[0.80, 0.02], [0.74, 0.20]], 0.09],
    ['rect', 0.00, 0.38, 0.20, 0.38],
    ['rect', 0.80, 0.38, 0.20, 0.38],
    ['poly', [[0.16, 0.34], [0.46, 0.34], [0.70, 0.52], [0.70, 0.80], [0.16, 0.80]]],
    ['poly', [[0.84, 0.48], [0.58, 0.48], [0.34, 0.62], [0.34, 0.86], [0.84, 0.86]]],
    ['stroke', [[0.34, 0.46], [0.60, 0.64]], 0.07, { bg: true }],
  ],
  monitorChart: [
    ['rect', 0.00, 0.00, 1.00, 0.70],
    ['rect', 0.08, 0.08, 0.84, 0.54, { bg: true }],
    ['stroke', [[0.18, 0.48], [0.38, 0.26], [0.58, 0.42], [0.82, 0.18]], 0.05],
    ['ellipse', 0.13, 0.43, 0.10, 0.10],
    ['ellipse', 0.33, 0.21, 0.10, 0.10],
    ['ellipse', 0.53, 0.37, 0.10, 0.10],
    ['ellipse', 0.77, 0.13, 0.10, 0.10],
    ['rect', 0.42, 0.70, 0.16, 0.14],
    ['roundRect', 0.24, 0.84, 0.52, 0.16, { r: 0.04 }],
  ],
  monitorPulse: [
    ['rect', 0.00, 0.00, 1.00, 0.72],
    ['rect', 0.08, 0.09, 0.84, 0.54, { bg: true }],
    ['stroke', [[0.14, 0.38], [0.28, 0.38], [0.36, 0.16], [0.48, 0.56], [0.58, 0.30], [0.66, 0.38], [0.86, 0.38]], 0.05],
    ['rect', 0.42, 0.72, 0.16, 0.14],
    ['roundRect', 0.24, 0.86, 0.52, 0.14, { r: 0.04 }],
  ],
  tabletPie: [
    ['roundRect', 0.00, 0.00, 1.00, 1.00, { r: 0.10 }],
    ['rect', 0.07, 0.09, 0.86, 0.75, { bg: true }],
    ['ellipse', 0.14, 0.30, 0.64, 0.44],
    ['pie', 0.14, 0.30, 0.64, 0.44, { angles: [270, 0], bg: true }],
    ['pie', 0.22, 0.22, 0.64, 0.44, { angles: [270, 0] }],
    ['ellipse', 0.42, 0.87, 0.16, 0.10],
  ],
  presentation: [
    ['rect', 0.00, 0.00, 1.00, 0.08],
    ['rect', 0.06, 0.08, 0.06, 0.48],
    ['rect', 0.88, 0.08, 0.06, 0.48],
    ['rect', 0.06, 0.50, 0.88, 0.06],
    ['rect', 0.20, 0.28, 0.13, 0.22],
    ['rect', 0.44, 0.18, 0.13, 0.32],
    ['rect', 0.68, 0.34, 0.13, 0.16],
    ['rect', 0.47, 0.56, 0.06, 0.28],
    ['ellipse', 0.42, 0.84, 0.16, 0.16],
  ],
  plane: [
    ['poly', [[0.00, 0.42], [1.00, 0.00], [0.52, 1.00], [0.38, 0.60]]],
    ['stroke', [[0.02, 0.86], [0.16, 0.72]], 0.06],
    ['stroke', [[0.14, 0.98], [0.24, 0.88]], 0.06],
  ],
  deskPhone: [
    ['roundRect', 0.24, 0.18, 0.76, 0.82, { r: 0.06 }],
    ['roundRect', 0.00, 0.00, 0.22, 0.66, { r: 0.09 }],
    ['ellipse', 0.02, 0.74, 0.20, 0.20],
    ['rect', 0.34, 0.30, 0.12, 0.12, { bg: true }],
    ['rect', 0.54, 0.30, 0.12, 0.12, { bg: true }],
    ['rect', 0.74, 0.30, 0.12, 0.12, { bg: true }],
    ['rect', 0.34, 0.50, 0.12, 0.12, { bg: true }],
    ['rect', 0.54, 0.50, 0.12, 0.12, { bg: true }],
    ['rect', 0.74, 0.50, 0.12, 0.12, { bg: true }],
    ['rect', 0.34, 0.70, 0.12, 0.12, { bg: true }],
    ['rect', 0.54, 0.70, 0.12, 0.12, { bg: true }],
    ['rect', 0.74, 0.70, 0.12, 0.12, { bg: true }],
  ],
};

/** Draw one pictogram: `name` from ICON_ART, tinted `colour` on a `bg` field. */
function icon (s, name, x, y, w, h, colour, bg) {
  const X = (u) => x + u * w;
  const Y = (v) => y + v * h;
  const scale = Math.min(w, h);
  ICON_ART[name].forEach((part) => {
    const kind = part[0];
    if (kind === 'poly') {
      const o = part[2] || {};
      freeform(s, x, y, w, h, o.bg ? bg : colour,
        part[1].map(([u, v], i) => ({ x: u * w, y: v * h, moveTo: i === 0 }))
          .concat([{ close: true }]));
      return;
    }
    if (kind === 'stroke') {
      const pts = part[1];
      const o = part[3] || {};
      for (let i = 1; i < pts.length; i++) {
        s.addShape('line', {
          x: X(Math.min(pts[i - 1][0], pts[i][0])), y: Y(Math.min(pts[i - 1][1], pts[i][1])),
          w: Math.abs(pts[i][0] - pts[i - 1][0]) * w, h: Math.abs(pts[i][1] - pts[i - 1][1]) * h,
          flipV: (pts[i][0] - pts[i - 1][0]) * (pts[i][1] - pts[i - 1][1]) < 0,
          line: { color: o.bg ? bg : colour, width: part[2] * scale * 72, endArrowType: o.arrow },
        });
      }
      return;
    }
    const [shape, u, v, uw, vh, o = {}] = part;
    const opts = {
      x: X(u), y: Y(v), w: uw * w, h: vh * h,
      fill: { color: o.bg ? bg : colour }, line: { type: 'none' },
    };
    if (o.hollow) {
      opts.fill = { type: 'none' };
      opts.line = { color: o.bg ? bg : colour, width: o.hollow * scale * 72 };
    }
    if (o.rotate) opts.rotate = o.rotate;
    if (o.flipH) opts.flipH = true;
    if (o.r) opts.rectRadius = o.r * Math.min(uw * w, vh * h);
    if (o.angles) { opts.angleRange = o.angles; }
    s.addShape(shape, opts);
  });
}

function slide01(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  deco(s, "parallelogram", 6.687, 0, 6.635, 2.73, { fill: BLACK, adj: 0.677 });
  deco(s, "parallelogram", 3.487, 2.931, 7.835, 4.569, { fill: DARK, adj: 0.677 });
  deco(s, "parallelogram", 7.899, 3.454, 3.381, 4.046, { fill: ORANGE, adj: 0.805 });
  // --- slide content
  text(s, 1.006, 2.074, 5.089, 1.683, "AVDIJA", 94, SEMI, ORANGE, { wrap: false });
  text(s, 1.088, 3.642, 3.88, 0.37, "Sport Presentation Template", 16, MED, INK);
}

function slide02(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  deco(s, "parallelogram", 0.809, 3.938, 5.111, 3.562, { fill: ORANGE, adj: 0.381 });
  // picture frame (empty in the source deck): 0, 0  6.729 x 7.5
  // --- slide content
  text(s, 6.877, 4.074, 5.062, 1.666, "There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour, or randomised words which don't look even slightly believable. If you are going to use a passage of Lorem Ipsum, you need to be sure there isn't anything embarrassing hidden in the middle of text. It uses a dictionary of over 200 Latin words.", 10.5, BODY, INK, { align: "justify", lh: 1.5 });
  text(s, 6.877, 3.514, 1.381, 0.471, "About Us", 16, MED, INK, { lh: 1.5 });
  richText(s, 6.877, 1.344, 4.367, 1.582, [
    ["WELCOME TO ", { sz: 44, f: SEMI, col: ORANGE }],
    ["AVDIJA", { sz: 44, f: SEMI, col: INK }],
  ]);
}

function slide03(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  deco(s, "rect", 9.688, 0, 3.646, 7.513, { fill: BLACK });
  // picture frame (empty in the source deck): 7.729, 0.729  4.672 x 6.229
  // --- slide content
  text(s, 0.856, 2.977, 5.26, 1.666, "There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour, or randomised words which don't look even slightly believable. If you are going to use a passage of Lorem Ipsum, you need to be sure there isn't anything embarrassing hidden in the middle of text.", 10.5, BODY, INK, { align: "justify", lh: 1.5 });
  text(s, 0.856, 2.388, 1.539, 0.505, "About Us", 16, MED, INK, { lh: 1.5 });
  richText(s, 0.856, 0.913, 4.865, 0.841, [
    ["ABOUT ", { sz: 44, f: SEMI, col: INK }],
    ["AVDIJA", { sz: 44, f: SEMI, col: ORANGE }],
  ]);
  deco(s, "rect", 6.46, 5.095, 3.365, 2.405, { fill: ORANGE });
  text(s, 6.773, 5.432, 1.539, 0.471, "About Us", 16, MED, WHITE, { lh: 1.5 });
  text(s, 6.773, 5.903, 2.703, 1.136, "You are going to use a passage of Lorem Ipsum, you need to be sure there isn't anything embarrassing hidden in the middle of text.", 10.5, BODY, WHITE, { align: "justify", lh: 1.5 });
  text(s, 0.856, 5.6, 5.26, 0.871, "All the Lorem Ipsum generators on the Internet tend to more than repeat predefined chunks as necessary, making this the first true generator on the Internet.", 10.5, BODY, GREY2, { align: "justify", lh: 1.5 });
  text(s, 0.856, 5.095, 1.539, 0.471, "Sporty", 16, MED, INK, { lh: 1.5 });
}

function slide04(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  // picture frame (empty in the source deck): 7.49, 0  5.311 x 7.5
  deco(s, "rect", 10.258, 5.688, 2.55, 1.812, { fill: ORANGE });
  // --- slide content
  richText(s, 0.856, 1.164, 5.357, 0.841, [
    ["WHAT IS ", { sz: 44, f: SEMI, col: INK }],
    ["AVDIJA", { sz: 44, f: SEMI, col: ORANGE }],
  ]);
  text(s, 0.856, 3.061, 5.621, 1.666, "All the Lorem Ipsum generators on the Internet tend to repeat predefined chunks as necessary, making this the first true generator on the Internet. It uses a dictionary of over 200 Latin words, combined with a handful of model sentence structures, to generate Lorem Ipsum which looks reasonable. The generated Lorem Ipsum is therefore always free from repetition, injected humour, or non-characteristic words etc.", 10.5, BODY, INK, { align: "justify", lh: 1.5 });
  text(s, 0.856, 2.448, 1.539, 0.505, "About Us", 16, MED, INK, { lh: 1.5 });
  text(s, 0.856, 4.954, 5.504, 0.606, "If you are going to use a passage of Lorem Ipsum, you need to be sure there isn't anything embarrassing hidden in the middle of text.", 10.5, BODY, INK2, { align: "justify", lh: 1.5 });
  deco(s, "roundRect", 0.895, 6.037, 1.802, 0.472, { fill: ORANGE });
  text(s, 0.895, 6.037, 1.802, 0.471, "Sport Avdija", 16, MED, INK, { align: "center", lh: 1.5 });
}

function slide05(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  deco(s, "rect", 2.17, 0.347, 3.02, 5.678, { fill: ORANGE, rotate: 30 });
  // picture frame (empty in the source deck): 0, 0  6.385 x 6.004
  // --- slide content
  text(s, 6.655, 3.182, 5.633, 1.136, "Lorep  ipsum duis aute irure dolor in kauselih oiluek epreh deriti vols esse cill inure dolorlaboru sit amet. Duis autelo irusitakus rephenderi Voluptate loremil. Lorep  ipsum duis aute irure dolor in kauselih oiluek Lorep  ipsum duis aute irure dolor in kauselih.", 10.5, BODY, INK, { align: "justify", lh: 1.5 });
  text(s, 6.655, 2.62, 1.829, 0.471, "About Dunk", 16, MED, INK, { lh: 1.5 });
  richText(s, 6.655, 1.283, 5.076, 0.841, [
    ["THE SLAM ", { sz: 44, f: SEMI, col: INK }],
    ["DUNK", { sz: 44, f: SEMI, col: ORANGE }],
  ]);
  text(s, 6.655, 4.644, 5.633, 0.606, "Lorep  ipsum duis aute irure dolor in kauselih oiluek epreh deriti vols esse cill inure dolorlaboru sit amet. Duis autelo irusitakus rephenderi.", 10.5, BODY, INK, { lh: 1.5 });
}

function slide06(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  // picture frame (empty in the source deck): 0.646, 2.875  3.812 x 3.981
  // picture frame (empty in the source deck): 4.688, 2.878  2.854 x 3.978
  deco(s, "rect", 7.771, 2.875, 5.562, 3.981, { fill: DARK });
  // --- slide content
  richText(s, 0.85, 1.862, 5.816, 0.908, [
    ["There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour, or randomised words which don't look even slightly believable", { sz: 10.5, f: BODY, col: INK2 }],
    [".", { sz: 11, f: BODY, col: INK2 }],
  ], { align: "justify", lh: 1.5 });
  text(s, 0.834, 1.391, 2.15, 0.505, "About Driblbing", 16, MED, INK, { lh: 1.5 });
  richText(s, 0.834, 0.314, 6.691, 0.841, [
    ["DRIBBLING ", { sz: 44, f: SEMI, col: INK }],
    ["OFFENCE", { sz: 44, f: SEMI, col: ORANGE }],
  ]);
  text(s, 8.326, 3.267, 2.15, 0.471, "Change Of Pace", 16, MED, ORANGE, { lh: 1.5 });
  text(s, 8.326, 3.974, 4.116, 1.401, "Lorep  ipsum duis aute irure dolor in kauselih oiluek epreh deriti vols esse cill inure dolorlaboru sit amet. Duis autelo irusitakus rephenderi Voluptate loremil. Lorep  ipsum duis aute irure dolor in kauselih oiluek Lorep  ipsum duis aute irure dolor in kauselih", 10.5, BODY, WHITE, { align: "justify", lh: 1.5 });
  text(s, 8.326, 5.699, 2.15, 0.471, "Basic Dribbling", 16, MED, ORANGE, { lh: 1.5 });
}

function slide07(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  // picture frame (empty in the source deck): 0, 0  9.479 x 7.5
  // picture frame (empty in the source deck): 8.312, 1.761  3.625 x 3.978
  deco(s, "rect", 12.271, 1.758, 1.062, 3.981, { fill: ORANGE });
  // --- slide content
  deco(s, "rect", 0, 0, 9.479, 7.5, { fill: BLACK, alpha: 20 });
  text(s, 0.914, 3.074, 5.668, 1.931, "There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour, or randomised words which don't look even slightly believable. If you are going to use a passage of Lorem Ipsum, you need to be sure there isn't anything embarrassing hidden in the middle of text. All the Lorem Ipsum generators on the Internet tend to repeat predefined chunks as necessary, making this the first true generator on the Internet.", 10.5, BODY, WHITE, { align: "justify", lh: 1.5 });
  richText(s, 0.914, 1.126, 5.966, 0.841, [
    ["THREE POINT ", { sz: 44, f: SEMI, col: WHITE }],
    ["SHOT", { sz: 44, f: SEMI, col: ORANGE }],
  ]);
  text(s, 0.914, 2.574, 2.15, 0.471, "New Shot ", 16, MED, WHITE, { lh: 1.5 });
  text(s, 0.914, 5.329, 1.652, 0.505, "New Skils", 16, MED, WHITE, { lh: 1.5 });
  text(s, 0.914, 5.834, 5.668, 0.871, "There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour, or randomised words which don't look even slightly believable.", 10.5, BODY, WHITE, { align: "justify", lh: 1.5 });
}

function slide08(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  deco(s, "parallelogram", 0.951, 0, 5.111, 3.562, { fill: ORANGE, adj: 0.182 });
  deco(s, "parallelogram", 0.708, 3.938, 5.111, 3.562, { fill: DARK, adj: 0.182 });
  // picture frame (empty in the source deck): 0.708, 0  5.354 x 7.5
  // --- slide content
  text(s, 6.594, 3.656, 5.354, 1.931, "If you are going to use a passage of Lorem Ipsum, you need to be sure there isn't anything embarrassing hidden in the middle of text. All the Lorem Ipsum generators on the Internet tend to repeat predefined chunks as necessary, making this the first true generator on the Internet. It uses a dictionary of over 200 Latin words, combined with a handful of model sentence structures, to generate Lorem Ipsum which looks reasonable", 10.5, BODY, INK, { align: "justify", lh: 1.5 });
  text(s, 6.594, 3.185, 2.108, 0.505, "About Strenght", 16, MED, INK, { lh: 1.5 });
  richText(s, 6.594, 1.687, 5.52, 0.841, [
    ["BODY ", { sz: 44, f: SEMI, col: INK }],
    ["STRENGHT", { sz: 44, f: SEMI, col: ORANGE }],
  ]);
  text(s, 6.594, 5.73, 5.354, 0.63, "Lorep  ipsum duis aute irure dolor in kauselih oiluek epreh deriti vols esse cill inure dolorlaboru sit amet. Duis autelo irusitakus rephenderi. ", 10.5, BODY, INK, { align: "justify", lh: 1.5 });
}

function slide09(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  deco(s, "rect", 0, 0, 4.438, 7.5, { fill: BLACK });
  // picture frame (empty in the source deck): 0.775, 1.104  4.85 x 5.221
  deco(s, "rect", 5.958, 1.104, 7.375, 5.221, { fill: ORANGE });
  // --- slide content
  richText(s, 6.818, 1.723, 4.85, 1.582, [
    ["BODY ", { sz: 44, f: SEMI, col: WHITE }],
    ["ENDURANCE", { sz: 44, f: SEMI, col: INK }],
  ]);
  text(s, 6.818, 4.099, 4.85, 1.401, "If you are going to use a passage of Lorem Ipsum, you need to be sure there isn't anything embarrassing hidden in the middle of text. All the Lorem Ipsum generators on the Internet tend to repeat predefined chunks as necessary, making this the first true generator on the Internet.", 10.5, BODY, WHITE, { align: "justify", lh: 1.5 });
  text(s, 6.818, 3.628, 2.368, 0.467, "About Endurance", 16, MED, WHITE, { lh: 1.5 });
}

function slide10(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  deco(s, "rect", 0, 2.875, 13.333, 3.312, { fill: ORANGE });
  // picture frame (empty in the source deck): 2.035, 3.229  2.683 x 2.562
  // picture frame (empty in the source deck): 5.325, 3.229  2.683 x 2.562
  // picture frame (empty in the source deck): 8.616, 3.229  2.683 x 2.562
  // --- slide content
  richText(s, 4.035, 0.59, 5.263, 0.841, [
    ["OUR ", { sz: 44, f: SEMI, col: INK }],
    ["TEAM", { sz: 44, f: SEMI, col: ORANGE }],
  ], { align: "center" });
  text(s, 3.2, 1.822, 6.933, 0.606, "Lorep  ipsum duis aute irure dolor in kauselih oiluek epreh deriti vols esse cill inure dolorlaboru sit amet. Duis autelo irusitakus rephenderi", 10.5, BODY, INK, { align: "center", lh: 1.5 });
}

function slide11(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  deco(s, "rect", 0, 2.49, 9.896, 3.312, { fill: DARK });
  // picture frame (empty in the source deck): 3.625, 2.042  2.683 x 2.667
  // picture frame (empty in the source deck): 0.419, 2.042  2.683 x 2.667
  // picture frame (empty in the source deck): 6.831, 2.042  2.683 x 2.667
  // --- slide content
  richText(s, 4.021, 0.611, 5.263, 0.841, [
    ["NEW ", { sz: 44, f: SEMI, col: INK }],
    ["TEAM", { sz: 44, f: SEMI, col: ORANGE }],
  ], { align: "center" });
  text(s, 3.2, 6.41, 6.933, 0.63, "Lorep  ipsum duis aute irure dolor in kauselih oiluek epreh deriti vols esse cill inure dolorlaboru sit amet. Duis autelo irusitakus rephenderi", 10.5, BODY, INK, { align: "center", lh: 1.5 });
  richText(s, 0.863, 5.246, 1.554, 0.309, [
    ["Denver, ", { sz: 9, f: SEMI, col: WHITE, i: true }],
    ["03/14/1990", { sz: 9, f: OSSEMI, col: WHITE, i: true }],
  ], { align: "center", lh: 1.5 });
  text(s, 0.648, 4.91, 1.985, 0.337, "Fran Victor", 14, MED, WHITE, { align: "center" });
  richText(s, 4.236, 5.246, 1.621, 0.306, [
    ["Minnesota, ", { sz: 9, f: SEMI, col: WHITE, i: true }],
    ["03/14/1990", { sz: 9, f: OSSEMI, col: WHITE, i: true }],
  ], { align: "center", lh: 1.5 });
  text(s, 4.021, 4.91, 1.985, 0.337, "Eduard Honwa", 14, MED, WHITE, { align: "center" });
  richText(s, 7.403, 5.246, 1.554, 0.309, [
    ["Chigago, ", { sz: 9, f: SEMI, col: WHITE, i: true }],
    ["03/14/1990", { sz: 9, f: OSSEMI, col: WHITE, i: true }],
  ], { align: "center", lh: 1.5 });
  text(s, 7.187, 4.91, 1.985, 0.337, "Vierra Jo", 14, MED, WHITE, { align: "center" });
}

function slide12(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  deco(s, "rect", 8.312, 5.907, 4.167, 1.233, { fill: DARK, rotate: 180 });
  deco(s, "rect", 9.554, 0, 3.779, 4.896, { fill: ORANGE, rotate: 180 });
  // picture frame (empty in the source deck): 8.312, 0.36  4.167 x 5.337
  // --- slide content
  richText(s, 9.75, 6.593, 1.554, 0.306, [
    ["Toronto, ", { sz: 9, f: SEMI, col: WHITE, i: true }],
    ["03/14/1990", { sz: 9, f: OSSEMI, col: WHITE, i: true }],
  ], { align: "center", lh: 1.5 });
  text(s, 9.535, 6.257, 1.985, 0.337, "Jhone Dio", 14, MED, WHITE, { align: "center" });
  richText(s, 1.115, 1.167, 4.171, 0.841, [
    ["TOP ", { sz: 44, f: SEMI, col: INK }],
    ["SCORES", { sz: 44, f: SEMI, col: ORANGE }],
  ]);
  text(s, 1.115, 2.45, 2.187, 0.505, "New Best Player", 16, MED, INK, { lh: 1.5 });
  text(s, 1.115, 2.99, 5.432, 1.931, "All the Lorem Ipsum generators on the Internet tend to repeat predefined chunks as necessary, making this the first true generator on the Internet. It uses a dictionary of over 200 Latin words, combined with a handful of model sentence structures, to generate Lorem Ipsum which looks reasonable. The generated Lorem Ipsum is therefore always free from repetition, injected humour, or non-characteristic words etc.", 10.5, BODY, INK, { align: "justify", lh: 1.5 });
  text(s, 1.056, 5.091, 5.49, 0.871, "There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour, or randomised words which don't look even slightly believable. ", 10.5, BODY, INK2, { align: "justify", lh: 1.5 });
}

function slide13(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  deco(s, "rect", 0, 2.604, 13.333, 4.896, { fill: DARK });
  // picture frame (empty in the source deck): 0, 0  4.771 x 5.979
  deco(s, "rect", 5.273, 4.625, 7.185, 1.354, { fill: ORANGE, rotate: 180 });
  // --- slide content
  richText(s, 5.591, 5.333, 1.554, 0.309, [
    ["Dallas, ", { sz: 9, f: SEMI, col: WHITE, i: true }],
    ["03/14/1990", { sz: 9, f: OSSEMI, col: WHITE, i: true }],
  ], { align: "center", lh: 1.5 });
  text(s, 5.376, 4.996, 1.985, 0.337, "Carloos Samba", 14, MED, WHITE, { align: "center" });
  richText(s, 5.474, 0.843, 4.171, 0.841, [
    ["BEST ", { sz: 44, f: SEMI, col: INK }],
    ["PLAYER", { sz: 44, f: SEMI, col: ORANGE }],
  ]);
  text(s, 5.591, 2.871, 5.002, 1.401, "If you are going to use a passage of Lorem Ipsum, you need to be sure there isn't anything embarrassing hidden in the middle of text. All the Lorem Ipsum generators on the Internet tend to repeat predefined chunks as necessary, making this the first true generator on the Internet.", 10.5, BODY, WHITE, { align: "justify", lh: 1.5 });
  text(s, 7.56, 4.986, 4.6, 0.63, "Lorep  ipsum duis aute irure dolor in kauselih oiluek epreh deriti vols esse cill inure dolorlaboru sit amet. Duis autelo", 10.5, BODY, WHITE, { align: "justify", lh: 1.5 });
}

function slide14(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  // picture frame (empty in the source deck): 0, 0  13.333 x 3.562
  // --- slide content
  deco(s, "rect", 0, 0, 13.333, 3.562, { fill: BLACK, alpha: 20 });
  deco(s, "rect", 2.436, 3.042, 1.664, 1.665, { fill: ORANGE, rotate: 180 });
  deco(s, "rect", 5.834, 3.042, 1.664, 1.665, { fill: ORANGE, rotate: 180 });
  deco(s, "rect", 9.232, 3.042, 1.664, 1.665, { fill: ORANGE, rotate: 180 });
  text(s, 5.482, 5.671, 2.655, 1.136, "If you are going to use a passage of Lorem Ipsum, you need to be sure there isn't anything embarrassing hidden.", 10.5, BODY, INK, { align: "center", lh: 1.5 });
  text(s, 2.089, 5.671, 2.528, 1.136, "The generated Lorem Ipsum is therefore always free from repetition, injected humour, or non-characteristic words etc.", 10.5, BODY, INK, { align: "center", lh: 1.5 });
  text(s, 8.875, 5.671, 2.655, 1.136, "All the Lorem Ipsum generators on the Internet tend to repeat predefined chunks as necessary, making this the first", 10.5, BODY, INK, { align: "center", lh: 1.5 });
  text(s, 2.493, 5.046, 1.552, 0.471, "Our Service", 16, MED, INK, { align: "center", lh: 1.5 });
  text(s, 5.886, 5.046, 1.552, 0.471, "Our Service", 16, MED, INK, { align: "center", lh: 1.5 });
  text(s, 9.279, 5.046, 1.552, 0.471, "Our Service", 16, MED, INK, { align: "center", lh: 1.5 });
  text(s, 4.031, 0.76, 5.272, 0.841, "ABOUT SERVICE", 44, SEMI, WHITE, { align: "center" });
  icon(s, "target", 2.902, 3.534, 0.686, 0.682, BLACK, ORANGE);
  icon(s, "shoe", 6.191, 3.64, 1.032, 0.574, BLACK, WHITE);
  icon(s, "stopwatch", 9.723, 3.441, 0.694, 0.797, BLACK, ORANGE);
}

function slide15(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  deco(s, "rect", 0.458, 0, 5.542, 3.021, { fill: DARK });
  // picture frame (empty in the source deck): 0.458, 3.312  5.542 x 3.396
  deco(s, "rect", 6.354, 5.208, 6.104, 1.5, { fill: ORANGE });
  deco(s, "rect", 6.354, 3.312, 6.104, 1.5, { fill: ORANGE });
  // --- slide content
  text(s, 7.489, 3.618, 4.491, 0.871, "All the Lorem Ipsum generators on the Internet tend to repeat predefined chunks as necessary, making this the first true generator on the Internet.", 10.5, BODY, WHITE, { align: "justify", lh: 1.5 });
  text(s, 7.489, 5.468, 4.491, 0.871, "If you are going to use a passage of Lorem Ipsum, you need to be sure there isn't anything embarrassing hidden in the middle of text.", 10.5, BODY, WHITE, { align: "justify", lh: 1.5 });
  richText(s, 0.882, 0.731, 4.171, 0.841, [
    ["OUR ", { sz: 44, f: SEMI, col: WHITE }],
    ["SERVICE", { sz: 44, f: SEMI, col: ORANGE }],
  ]);
  text(s, 0.882, 1.622, 1.831, 0.471, "New Service", 16, MED, WHITE, { lh: 1.5 });
  text(s, 6.596, 1.572, 5.6, 1.401, "There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour, or randomised words which don't look even slightly believable. If you are going to use a passage of Lorem Ipsum, you need to be sure there isn't anything embarrassing hidden in the middle of text. ", 10.5, BODY, INK, { align: "justify", lh: 1.5 });
  text(s, 6.596, 1.026, 1.831, 0.471, "About Service", 16, MED, ORANGE, { lh: 1.5 });
  icon(s, "book", 6.727, 3.802, 0.595, 0.564, WHITE, ORANGE);
  icon(s, "trophy", 6.727, 5.674, 0.559, 0.564, WHITE, ORANGE);
}

function slide16(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  // picture frame (empty in the source deck): 0, 0  5.771 x 7.5
  // --- slide content
  deco(s, "rect", 4.833, 5.312, 6.771, 1.5, { fill: ORANGE });
  deco(s, "rect", 4.833, 3.417, 6.771, 1.5, { fill: BLACK });
  text(s, 6.242, 2.255, 5.196, 0.871, "Lorep  ipsum duis aute irure dolor in kauselih oiluek epreh deriti vols esse cill inure dolorlaboru sit amet. Duis autelo irusitakus rephenderi Voluptate loremil. ", 10.5, BODY, INK, { align: "justify", lh: 1.5 });
  text(s, 6.242, 1.784, 1.831, 0.471, "New Service", 16, MED, BLACK, { lh: 1.5 });
  richText(s, 6.242, 0.736, 4.171, 0.841, [
    ["OUR ", { sz: 44, f: SEMI, col: INK }],
    ["SERVICE", { sz: 44, f: SEMI, col: ORANGE }],
  ]);
  text(s, 6.242, 3.835, 4.71, 0.63, "Lorep  ipsum duis aute irure dolor in kauselih oiluek epreh deriti vols esse cill inure dolorlaboru sit amet. Duis autelo", 10.5, BODY, WHITE, { align: "justify", lh: 1.5 });
  text(s, 6.242, 5.754, 4.71, 0.63, "Lorep  ipsum duis aute irure dolor in kauselih oiluek epreh deriti vols esse cill inure dolorlaboru sit amet. Duis autelo", 10.5, BODY, WHITE, { align: "justify", lh: 1.5 });
  icon(s, "sunrise", 5.353, 3.904, 0.496, 0.462, ORANGE, BLACK);
  icon(s, "whistle", 5.354, 5.878, 0.497, 0.41, BLACK, ORANGE);
}

function slide17(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  deco(s, "rect", 5.804, 5.342, 1.178, 1.179, { fill: ORANGE });
  deco(s, "rect", 5.804, 3.54, 1.178, 1.179, { fill: ORANGE });
  // picture frame (empty in the source deck): 0, 0  7.292 x 7.292
  // --- slide content
  deco(s, "diagStripe", 0, 0, 7.619, 7.5, { fill: BLACK });
  text(s, 7.196, 3.773, 5.311, 1.136, "There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour, or randomised words which don't look even slightly believable.", 10.5, BODY, INK, { align: "justify", lh: 1.5 });
  text(s, 7.292, 3.302, 1.831, 0.471, "Best Service", 16, MED, BLACK, { lh: 1.5 });
  text(s, 7.196, 5.668, 5.311, 0.871, "All the Lorem Ipsum generators on the Internet tend to repeat predefined chunks as necessary, making this the first true generator on the Internet.", 10.5, BODY, INK, { align: "justify", lh: 1.5 });
  text(s, 7.292, 5.196, 1.831, 0.471, "Best Service", 16, MED, BLACK, { lh: 1.5 });
  richText(s, 7.196, 0.8, 4.547, 0.841, [
    ["BEST ", { sz: 44, f: SEMI, col: INK }],
    ["SERVICE", { sz: 44, f: SEMI, col: ORANGE }],
  ]);
  text(s, 7.196, 2.258, 5.311, 0.606, "Lorep  ipsum duis aute irure dolor in kauselih oiluek epreh deriti vols esse cill inure dolorlaboru sit amet. Duis autelo irusitakus rephenderi", 10.5, BODY, INK, { align: "justify", lh: 1.5 });
  text(s, 7.196, 1.786, 1.831, 0.471, "New Service", 16, MED, BLACK, { lh: 1.5 });
  icon(s, "cloud", 6.156, 3.952, 0.483, 0.48, WHITE, ORANGE);
  icon(s, "bottle", 6.283, 5.611, 0.23, 0.501, WHITE, ORANGE);
}

function slide18(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  // picture frame (empty in the source deck): 0, 0  13.333 x 7.5
  // --- slide content
  deco(s, "rect", 0, 0.031, 13.333, 7.5, { fill: BLACK, alpha: 38 });
  deco(s, "diagStripe", 5.682, 2.431, 5.038, 5.038, { fill: ORANGE, rotate: 180 });
  freeform(s, 10.974, 0, 2.359, 4.719, DARK, [
    { x: 2.248, y: 0, moveTo: true },
    { x: 2.359, y: 0 },
    { x: 2.359, y: 2.36 },
    { x: 0, y: 4.719 },
    { x: 0, y: 2.248 },
    { close: true },
  ]);
  text(s, 1.088, 2.385, 7.779, 1.717, "Break Slide", 94, SEMI, WHITE, { wrap: false });
  text(s, 1.088, 3.931, 3.88, 0.404, "Sport Presentation Template", 18, MED, WHITE);
}

function slide19(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  // picture frame (empty in the source deck): 0.813, 0.771  3.134 x 5.883
  // picture frame (empty in the source deck): 4.224, 3.625  4.276 x 3.029
  // picture frame (empty in the source deck): 8.778, 3.625  3.514 x 3.029
  deco(s, "rect", 4.224, 0.771, 8.068, 2.635, { fill: DARK });
  // --- slide content
  richText(s, 4.934, 0.992, 4.547, 0.841, [
    ["BEST ", { sz: 44, f: SEMI, col: WHITE }],
    ["GALLERY", { sz: 44, f: SEMI, col: ORANGE }],
  ]);
  text(s, 4.934, 2.497, 5.6, 0.63, "Lorep  ipsum duis aute irure dolor in kauselih oiluek epreh deriti vols esse cill inure dolorlaboru sit amet. Duis autelo irusitakus rephenderi", 10.5, BODY, WHITE, { align: "justify", lh: 1.5 });
  text(s, 4.934, 1.993, 1.831, 0.471, "New Portfolio", 16, MED, WHITE, { lh: 1.5 });
}

function slide20(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  // picture frame (empty in the source deck): 0.958, 3.688  3.134 x 3.312
  // picture frame (empty in the source deck): 4.356, 3.688  3.134 x 3.312
  // picture frame (empty in the source deck): 7.753, 3.688  5.58 x 3.312
  // picture frame (empty in the source deck): 7.753, 1.125  5.58 x 2.188
  // --- slide content
  richText(s, 0.958, 1.133, 6.333, 0.841, [
    ["GALLERY ", { sz: 44, f: SEMI, col: INK }],
    ["PORTFOLIO", { sz: 44, f: SEMI, col: ORANGE }],
  ]);
  text(s, 0.958, 2.656, 5.6, 0.63, "If you are going to use a passage of Lorem Ipsum, you need to be sure there isn't anything embarrassing hidden in the middle of text.", 10.5, BODY, INK, { align: "justify", lh: 1.5 });
  text(s, 0.958, 2.185, 1.831, 0.471, "New Portfolio", 16, MED, INK, { lh: 1.5 });
}

function slide21(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  // picture frame (empty in the source deck): 8.333, 4.229  4.062 x 2.625
  // picture frame (empty in the source deck): 0.958, 1.375  3.354 x 2.604
  // picture frame (empty in the source deck): 0.958, 4.229  7.146 x 2.625
  // picture frame (empty in the source deck): 4.531, 1.375  7.865 x 2.604
  // --- slide content
  richText(s, 3.5, 0.346, 6.333, 0.841, [
    ["GALLERY ", { sz: 44, f: SEMI, col: INK }],
    ["PORTFOLIO", { sz: 44, f: SEMI, col: ORANGE }],
  ], { align: "center" });
}

function slide22(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  // picture frame (empty in the source deck): 4.688, 0.646  4.062 x 3.708
  // picture frame (empty in the source deck): 9, 0.646  3.375 x 3.708
  deco(s, "rect", 0.854, 0.646, 3.583, 3.708, { fill: ORANGE });
  // --- slide content
  richText(s, 0.86, 4.952, 4.146, 1.582, [
    ["AVDIJA ", { sz: 44, f: SEMI, col: ORANGE }],
    ["PORTFOLIO", { sz: 44, f: SEMI, col: INK }],
  ]);
  text(s, 6.13, 5.523, 5.6, 1.136, "It uses a dictionary of over 200 Latin words, combined with a handful of model sentence structures, to generate Lorem Ipsum which looks reasonable. The generated Lorem Ipsum is therefore always free from repetition, injected humour, or non-characteristic words etc.", 10.5, BODY, INK, { align: "justify", lh: 1.5 });
  text(s, 6.13, 4.952, 1.831, 0.471, "New Portfolio", 16, MED, INK, { lh: 1.5 });
  text(s, 1.192, 1.937, 2.959, 1.666, "There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour, or randomised words which don't look even.", 10.5, BODY, WHITE, { align: "justify", lh: 1.5 });
  text(s, 1.192, 1.285, 1.831, 0.471, "New Gallery", 16, MED, INK, { lh: 1.5 });
}

function slide23(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  // picture frame (empty in the source deck): 8.354, 0.604  4.979 x 2.807
  // picture frame (empty in the source deck): 8.354, 3.945  4.979 x 3.658
  deco(s, "rect", 0, 0.604, 8.104, 2.807, { fill: DARK });
  deco(s, "rect", 0, 3.928, 8.104, 3.572, { fill: ORANGE });
  // --- slide content
  richText(s, 0.646, 1.326, 6.562, 0.841, [
    ["AVDIJA ", { sz: 44, f: SEMI, col: WHITE }],
    ["COLLECTION", { sz: 44, f: SEMI, col: ORANGE }],
  ]);
  text(s, 0.764, 4.569, 1.966, 0.467, "New Collection", 16, MED, WHITE, { lh: 1.5 });
  text(s, 0.764, 5.083, 6.327, 1.666, "There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour, or randomised words which don't look even slightly believable. If you are going to use a passage of Lorem Ipsum, you need to be sure there isn't anything embarrassing hidden in the middle of text. All the Lorem Ipsum generators on the Internet tend to repeat predefined chunks as necessary, making this the first true generator on the Internet.", 10.5, BODY, WHITE, { align: "justify", lh: 1.5 });
}

function slide24(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  // picture frame (empty in the source deck): 0.625, 0.562  3.188 x 4.292
  // picture frame (empty in the source deck): 4.208, 0.562  3.188 x 4.292
  deco(s, "rect", 0.625, 5.062, 6.771, 2.438, { fill: ORANGE });
  // picture frame (empty in the source deck): 7.792, 0.562  4.792 x 4.292
  // --- slide content
  text(s, 1.74, 5.521, 4.146, 1.582, "GALLERY LAYOUT", 44, SEMI, INK, { align: "center" });
  text(s, 7.792, 5.613, 4.66, 1.401, "Lorep  ipsum duis aute irure dolor in kauselih oiluek epreh deriti vols esse cill inure dolorlaboru sit amet. Duis autelo irusitakus rephenderi Voluptate loremil. Lorep  ipsum duis aute irure dolor in kauselih oiluek Lorep  ipsum duis aute irure dolor in kauselih.", 10.5, BODY, INK, { align: "justify", lh: 1.5 });
  text(s, 7.792, 5.099, 1.831, 0.471, "New Gallery", 16, MED, INK, { lh: 1.5 });
}

function slide25(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  deco(s, "diagStripe", 9.627, 3.779, 3.707, 3.707, { fill: DARK, rotate: 90 });
  deco(s, "diagStripe", 8.295, 0, 5.038, 5.038, { fill: ORANGE, rotate: 90 });
  // picture frame (empty in the source deck): 7.224, 2.336  4.428 x 2.769
  // --- slide content
  photo(s, 6.631, 2.16, 5.572, 3.225);
  richText(s, 0.998, 1.731, 5.058, 1.582, [
    ["NEW", { sz: 44, f: SEMI, col: ORANGE }],
    [" MOCKUP LAPTOP", { sz: 44, f: SEMI, col: INK }],
  ]);
  text(s, 0.998, 3.809, 5.058, 1.401, "All the Lorem Ipsum generators on the Internet tend to repeat predefined chunks as necessary, making this the first true generator on the Internet. It uses a dictionary of over 200 Latin words, combined with a handful of model sentence structures, to generate Lorem Ipsum which looks reasonable.", 10.5, BODY, INK, { align: "justify", lh: 1.5 });
}

function slide26(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  deco(s, "rtTriangle", 5.833, 0, 7.5, 7.5, { fill: ORANGE, rotate: 270 });
  // picture frame (empty in the source deck): 7.204, 2.336  4.428 x 2.769
  // --- slide content
  photo(s, 6.997, 2.216, 4.882, 4.055);
  richText(s, 0.998, 1.731, 5.058, 1.582, [
    ["NEW", { sz: 44, f: SEMI, col: ORANGE }],
    [" MOCKUP DEKSTOP", { sz: 44, f: SEMI, col: INK }],
  ]);
  text(s, 0.998, 3.75, 5.058, 1.401, "It uses a dictionary of over 200 Latin words, combined with a handful of model sentence structures, to generate Lorem Ipsum which looks reasonable. The generated Lorem Ipsum is therefore always free from repetition, injected humour, or non-characteristic words etc.", 10.5, BODY, INK, { align: "justify", lh: 1.5 });
}

function slide27(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  // picture frame (empty in the source deck): 1.289, 1.317  3.159 x 4.81
  deco(s, "rect", 4.917, 1.317, 8.417, 4.81, { fill: DARK });
  deco(s, "rect", 5.427, 5.574, 3.157, 1.106, { fill: ORANGE });
  // --- slide content
  photo(s, 1.135, 1.099, 3.466, 5.303);
  richText(s, 5.756, 1.806, 5.343, 1.582, [
    ["NEW MOCKUP ", { sz: 44, f: SEMI, col: WHITE }],
    ["TABLET", { sz: 44, f: SEMI, col: ORANGE }],
  ]);
  text(s, 5.756, 3.548, 5.343, 1.666, "There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour, or randomised words which don't look even slightly believable. If you are going to use a passage of Lorem Ipsum, you need to be sure there isn't anything embarrassing hidden in the middle of text.", 10.5, BODY, WHITE, { align: "justify", lh: 1.5 });
  text(s, 5.671, 5.894, 2.494, 0.467, "New Mockup", 16, MED, INK, { align: "center", lh: 1.5 });
}

function slide28(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  deco(s, "parallelogram", 2.26, 3.833, 5.06, 3.667, { fill: ORANGE, flipH: true, adj: 0.467 });
  deco(s, "parallelogram", 0.479, 0, 5.06, 3.417, { fill: DARK, flipH: true, adj: 0.467 });
  // picture frame (empty in the source deck): 1.188, 1.111  2.651 x 5.444
  // --- slide content
  photo(s, 1.107, 1.055, 2.813, 5.542);
  richText(s, 6.665, 1.725, 4.883, 1.582, [
    ["NEW", { sz: 44, f: SEMI, col: ORANGE }],
    [" MOCKUP PHONE", { sz: 44, f: SEMI, col: INK }],
  ]);
  richText(s, 6.665, 3.722, 4.993, 1.741, [
    ["There are many variations of passages of Lorem Ipsum available, but the majority have suffered alteration in some form, by injected humour, or randomised words which don't look even slightly believable. If you are going to use a passage of Lorem Ipsum, you need to be sure there isn't anything embarrassing hidden in the middle of text", { sz: 10.5, f: BODY, col: INK }],
    [".", { sz: 11, f: BODY, col: INK }],
  ], { align: "justify", lh: 1.5 });
}

function slide29(pptx) {
  const s = newSlide(pptx);
  // --- slide content
  text(s, 4.472, 0.629, 4.39, 0.841, "INFOGRAPHIC", 44, SEMI, INK);
  deco(s, "teardrop", 4.492, 4.454, 2.038, 2.038, { fill: BLACK });
  deco(s, "teardrop", 4.492, 2.174, 2.038, 2.038, { fill: ORANGE, rotate: 90 });
  deco(s, "teardrop", 6.761, 2.174, 2.038, 2.038, { fill: BLACK, rotate: 180 });
  deco(s, "teardrop", 6.803, 4.428, 2.038, 2.038, { fill: ORANGE, rotate: 270 });
  richText(s, 1.407, 2.731, 2.654, 0.834, [
    ["Lorem ipsum dolor sit amet, con adipiscing elit, sed do eiusmoun ", { sz: 10, f: BODY, col: GREY2, br: true }],
    ["aliqua. Ut enim ad", { sz: 10, f: BODY, col: GREY2 }],
  ], { align: "right", lh: 1.5 });
  text(s, 2.702, 2.502, 1.359, 0.269, "TEXT TITLE HERE", 10, SEMI, INK, { align: "right", wrap: false });
  richText(s, 1.407, 4.919, 2.654, 0.834, [
    ["Lorem ipsum dolor sit amet, con adipiscing elit, sed do eiusmoun ", { sz: 10, f: BODY, col: GREY2, br: true }],
    ["aliqua. Ut enim ad", { sz: 10, f: BODY, col: GREY2 }],
  ], { align: "right", lh: 1.5 });
  text(s, 2.702, 4.69, 1.359, 0.269, "TEXT TITLE HERE", 10, SEMI, INK, { align: "right", wrap: false });
  richText(s, 9.346, 2.79, 2.654, 0.834, [
    ["Lorem ipsum dolor sit amet, con adipiscing elit, sed do eiusmoun ", { sz: 10, f: BODY, col: GREY2, br: true }],
    ["aliqua. Ut enim ad", { sz: 10, f: BODY, col: GREY2 }],
  ], { lh: 1.5 });
  text(s, 9.346, 2.562, 1.359, 0.269, "TEXT TITLE HERE", 10, SEMI, INK, { wrap: false });
  richText(s, 9.293, 4.976, 2.654, 0.834, [
    ["Lorem ipsum dolor sit amet, con adipiscing elit, sed do eiusmoun ", { sz: 10, f: BODY, col: GREY2, br: true }],
    ["aliqua. Ut enim ad", { sz: 10, f: BODY, col: GREY2 }],
  ], { lh: 1.5 });
  text(s, 9.293, 4.748, 1.359, 0.269, "TEXT TITLE HERE", 10, SEMI, INK, { wrap: false });
  icon(s, "masks", 5.179, 2.942, 0.696, 0.696, WHITE, ORANGE);
  icon(s, "camera", 7.356, 3.003, 0.697, 0.575, WHITE, BLACK);
  icon(s, "trophy", 5.213, 5.073, 0.61, 0.687, WHITE, BLACK);
  icon(s, "rocket", 7.421, 5.047, 0.702, 0.699, WHITE, ORANGE);
}

function slide30(pptx) {
  const s = newSlide(pptx);
  // --- slide content
  deco(s, "donut", 4.788, 2.698, 3.75, 3.75, { fill: BLACK, adj: 0.107 });
  freeform(s, 6.057, 2.269, 1.259, 2.319, DARK, [
    { x: 0.63, y: 0, moveTo: true },
    { x: 0.042, y: 0.587, curve: { type: "cubic", x1: 0.306, y1: 0, x2: 0.042, y2: 0.263 } },
    { x: 0.385, y: 1.123, curve: { type: "cubic", x1: 0.042, y1: 0.826, x2: 0.184, y2: 1.026 } },
    { x: 0.552, y: 1.382, curve: { type: "cubic", x1: 0.485, y1: 1.169, x2: 0.552, y2: 1.269 } },
    { x: 0.376, y: 1.637, curve: { type: "cubic", x1: 0.552, y1: 1.495, x2: 0.481, y2: 1.599 } },
    { x: 0, y: 1.955, curve: { type: "cubic", x1: 0.217, y1: 1.696, x2: 0.084, y2: 1.809 } },
    { x: 0.632, y: 2.319 },
    { x: 1.259, y: 1.955 },
    { x: 0.887, y: 1.637, curve: { type: "cubic", x1: 1.175, y1: 1.809, x2: 1.041, y2: 1.696 } },
    { x: 0.707, y: 1.382, curve: { type: "cubic", x1: 0.778, y1: 1.6, x2: 0.707, y2: 1.495 } },
    { x: 0.707, y: 1.374 },
    { x: 0.87, y: 1.123, curve: { type: "cubic", x1: 0.707, y1: 1.265, x2: 0.774, y2: 1.169 } },
    { x: 1.217, y: 0.545, curve: { type: "cubic", x1: 1.087, y1: 1.026, x2: 1.234, y2: 0.8 } },
    { x: 0.673, y: 0.002, curve: { type: "cubic", x1: 1.196, y1: 0.257, x2: 0.962, y2: 0.022 } },
    { x: 0.63, y: 0, curve: { type: "cubic", x1: 0.659, y1: 0.001, x2: 0.644, y2: 0 } },
    { close: true },
  ]);
  freeform(s, 4.531, 4.224, 2.159, 1.816, ORANGE, [
    { x: 1.527, y: 0, moveTo: true },
    { x: 1.439, y: 0.481, curve: { type: "cubic", x1: 1.443, y1: 0.147, x2: 1.414, y2: 0.314 } },
    { x: 1.305, y: 0.761, curve: { type: "cubic", x1: 1.46, y1: 0.594, x2: 1.406, y2: 0.707 } },
    { x: 1.301, y: 0.766 },
    { x: 1.159, y: 0.803, curve: { type: "cubic", x1: 1.258, y1: 0.79, x2: 1.209, y2: 0.803 } },
    { x: 1, y: 0.753, curve: { type: "cubic", x1: 1.102, y1: 0.804, x2: 1.046, y2: 0.786 } },
    { x: 0.655, y: 0.641, curve: { type: "cubic", x1: 0.899, y1: 0.68, x2: 0.778, y2: 0.641 } },
    { x: 0.33, y: 0.741, curve: { type: "cubic", x1: 0.544, y1: 0.641, x2: 0.432, y2: 0.673 } },
    { x: 0.13, y: 1.485, curve: { type: "cubic", x1: 0.088, y1: 0.903, x2: 0, y2: 1.221 } },
    { x: 0.656, y: 1.816, curve: { type: "cubic", x1: 0.232, y1: 1.696, x2: 0.441, y2: 1.816 } },
    { x: 0.95, y: 1.736, curve: { type: "cubic", x1: 0.756, y1: 1.816, x2: 0.857, y2: 1.79 } },
    { x: 1.242, y: 1.171, curve: { type: "cubic", x1: 1.159, y1: 1.619, x2: 1.264, y2: 1.397 } },
    { x: 1.385, y: 0.899, curve: { type: "cubic", x1: 1.234, y1: 1.062, x2: 1.289, y2: 0.954 } },
    { x: 1.524, y: 0.861, curve: { type: "cubic", x1: 1.429, y1: 0.874, x2: 1.477, y2: 0.861 } },
    { x: 1.694, y: 0.925, curve: { type: "cubic", x1: 1.585, y1: 0.861, x2: 1.645, y2: 0.882 } },
    { x: 2.159, y: 1.087, curve: { type: "cubic", x1: 1.828, y1: 1.029, x2: 1.992, y2: 1.087 } },
    { x: 2.159, y: 0.364 },
    { close: true },
  ]);
  freeform(s, 6.689, 4.224, 2.175, 1.817, ORANGE, [
    { x: 0.627, y: 0, moveTo: true },
    { x: 0, y: 0.364 },
    { x: 0, y: 1.088 },
    { x: 0.46, y: 0.925, curve: { type: "cubic", x1: 0.167, y1: 1.088, x2: 0.33, y2: 1.029 } },
    { x: 0.77, y: 0.9, curve: { type: "cubic", x1: 0.547, y1: 0.85, x2: 0.672, y2: 0.84 } },
    { x: 0.778, y: 0.904 },
    { x: 0.912, y: 1.171, curve: { type: "cubic", x1: 0.87, y1: 0.958, x2: 0.924, y2: 1.063 } },
    { x: 1.238, y: 1.757, curve: { type: "cubic", x1: 0.891, y1: 1.406, x2: 1.008, y2: 1.644 } },
    { x: 1.495, y: 1.817, curve: { type: "cubic", x1: 1.318, y1: 1.796, x2: 1.406, y2: 1.817 } },
    { x: 1.983, y: 1.56, curve: { type: "cubic", x1: 1.685, y1: 1.817, x2: 1.871, y2: 1.724 } },
    { x: 1.79, y: 0.72, curve: { type: "cubic", x1: 2.175, y1: 1.272, x2: 2.083, y2: 0.887 } },
    { x: 1.496, y: 0.641, curve: { type: "cubic", x1: 1.697, y1: 0.666, x2: 1.596, y2: 0.641 } },
    { x: 1.158, y: 0.749, curve: { type: "cubic", x1: 1.375, y1: 0.641, x2: 1.257, y2: 0.678 } },
    { x: 0.849, y: 0.761, curve: { type: "cubic", x1: 1.067, y1: 0.815, x2: 0.945, y2: 0.82 } },
    { x: 0.715, y: 0.481, curve: { type: "cubic", x1: 0.749, y1: 0.707, x2: 0.694, y2: 0.594 } },
    { x: 0.627, y: 0, curve: { type: "cubic", x1: 0.74, y1: 0.314, x2: 0.711, y2: 0.147 } },
    { close: true },
  ]);
  text(s, 1.334, 3.379, 2.579, 0.707, "Lorep  ipsum duis aute irure dolor in kauselih oilusioi reprehenderitilores voluptates esse", 8, BODY, GREY, { lh: 1.5 });
  deco(s, "line", 1.434, 3.256, 2.109, 0, { line: RULE, lineW: 1.75 });
  text(s, 1.334, 5.629, 2.579, 0.707, "Lorep  ipsum duis aute irure dolor in kauselih oilusioi reprehenderitilores voluptates esse", 8, BODY, GREY, { lh: 1.5 });
  deco(s, "line", 1.434, 5.505, 2.109, 0, { line: RULE, lineW: 1.75 });
  text(s, 9.659, 5.629, 2.579, 0.707, "Lorep  ipsum duis aute irure dolor in kauselih oilusioi reprehenderitilores voluptates esse", 8, BODY, GREY, { lh: 1.5 });
  deco(s, "line", 9.759, 5.505, 2.109, 0, { line: RULE, lineW: 1.75 });
  text(s, 9.659, 3.407, 2.579, 0.707, "Lorep  ipsum duis aute irure dolor in kauselih oilusioi reprehenderitilores voluptates esse", 8, BODY, GREY, { lh: 1.5 });
  deco(s, "roundRect", 1.388, 4.928, 1.501, 0.392, { fill: BLACK, adj: 0.5 });
  text(s, 1.521, 5.032, 1.236, 0.236, "NEW TEXT", 8, MED, WHITE, { align: "center" });
  deco(s, "roundRect", 1.388, 2.712, 1.501, 0.392, { fill: ORANGE, adj: 0.5 });
  text(s, 1.521, 2.816, 1.236, 0.236, "NEW TEXT", 8, MED, WHITE, { align: "center" });
  deco(s, "roundRect", 9.759, 2.712, 1.501, 0.392, { fill: ORANGE, adj: 0.5 });
  text(s, 9.892, 2.816, 1.236, 0.236, "NEW TEXT", 8, MED, WHITE, { align: "center" });
  deco(s, "roundRect", 9.759, 4.962, 1.501, 0.392, { fill: BLACK, adj: 0.5 });
  text(s, 9.892, 5.066, 1.236, 0.236, "NEW TEXT", 8, MED, WHITE, { align: "center" });
  deco(s, "ellipse", 6.214, 4.136, 0.945, 0.945, { fill: SILVER });
  deco(s, "line", 9.76, 3.256, 2.109, 0, { line: RULE, lineW: 1.75 });
  text(s, 4.472, 0.629, 4.39, 0.841, "INFOGRAPHIC", 44, SEMI, INK);
  icon(s, "envelope", 6.402, 2.606, 0.571, 0.436, WHITE, DARK);
  icon(s, "megaphone", 4.93, 5.196, 0.463, 0.453, WHITE, ORANGE);
  icon(s, "medkit", 7.942, 5.18, 0.551, 0.484, WHITE, ORANGE);
  icon(s, "globe", 6.452, 4.372, 0.474, 0.474, NAVY, SILVER);
}

function slide31(pptx) {
  const s = newSlide(pptx);
  // --- slide content
  freeform(s, 6.771, 2.383, 1.773, 2.151, ORANGE, [
    { x: 0.005, y: 0, moveTo: true },
    { x: 0.29, y: 0.522 },
    { x: 0, y: 1.059 },
    { x: 0.716, y: 1.861, curve: { type: "cubic", x1: 0.403, y1: 1.11, x2: 0.709, y2: 1.451 } },
    { x: 1.249, y: 2.151 },
    { x: 1.773, y: 1.861 },
    { x: 0.005, y: 0, curve: { type: "cubic", x1: 1.766, y1: 0.869, x2: 0.992, y2: 0.056 } },
    { close: true },
  ]);
  freeform(s, 4.794, 2.38, 2.151, 1.776, DARK, [
    { x: 1.861, y: 0, moveTo: true },
    { x: 0, y: 1.771, curve: { type: "cubic", x1: 0.869, y1: 0.007, x2: 0.056, y2: 0.781 } },
    { x: 0.522, y: 1.484 },
    { x: 1.059, y: 1.776 },
    { x: 1.861, y: 1.057, curve: { type: "cubic", x1: 1.11, y1: 1.37, x2: 1.451, y2: 1.064 } },
    { x: 2.151, y: 0.524 },
    { x: 1.861, y: 0 },
    { close: true },
  ]);
  freeform(s, 6.393, 4.36, 2.149, 1.776, DARK, [
    { x: 1.09, y: 0, moveTo: true },
    { x: 0.29, y: 0.719, curve: { type: "cubic", x1: 1.038, y1: 0.404, x2: 0.698, y2: 0.71 } },
    { x: 0, y: 1.25 },
    { x: 0.288, y: 1.776 },
    { x: 2.149, y: 0.005, curve: { type: "cubic", x1: 1.277, y1: 1.769, x2: 2.093, y2: 0.992 } },
    { x: 2.149, y: 0.005 },
    { x: 1.627, y: 0.29 },
    { x: 1.09, y: 0 },
    { close: true },
  ]);
  freeform(s, 4.789, 3.982, 1.778, 2.149, ORANGE, [
    { x: 0.526, y: 0, moveTo: true },
    { x: 0, y: 0.29 },
    { x: 1.771, y: 2.149, curve: { type: "cubic", x1: 0.009, y1: 1.28, x2: 0.784, y2: 2.093 } },
    { x: 1.486, y: 1.627 },
    { x: 1.778, y: 1.089 },
    { x: 1.059, y: 0.29, curve: { type: "cubic", x1: 1.372, y1: 1.038, x2: 1.066, y2: 0.698 } },
    { x: 0.526, y: 0 },
    { close: true },
  ]);
  text(s, 9.292, 3.34, 2.57, 0.858, "Lorep  ipsum duis aute irure dolor in kauselih oiluek eprehfa hreer deriti vols daqeesse cill.", 10, BODY, GREY, { lh: 1.5 });
  text(s, 9.292, 5.458, 2.57, 0.858, "Lorep  ipsum duis aute irure dolor in kauselih oiluek eprehfa hreer deriti vols daqeesse cill.", 10, BODY, GREY, { lh: 1.5 });
  text(s, 1.323, 3.34, 2.668, 0.858, "Lorep  ipsum duis aute irure dolor in kauselih oiluek eprehfa hreer deriti vols daqeesse cill.", 10, BODY, GREY, { align: "right", lh: 1.5 });
  text(s, 1.323, 5.458, 2.668, 0.858, "Lorep  ipsum duis aute irure dolor in kauselih oiluek eprehfa hreer deriti vols daqeesse cill.", 10, BODY, GREY, { align: "right", lh: 1.5 });
  deco(s, "roundRect", 2.484, 2.712, 1.501, 0.392, { fill: BLACK, adj: 0.5 });
  text(s, 2.617, 2.816, 1.236, 0.236, "NEW TEXT", 8, MED, WHITE, { align: "center" });
  deco(s, "roundRect", 9.287, 2.712, 1.501, 0.392, { fill: BLACK, adj: 0.5 });
  text(s, 9.419, 2.816, 1.236, 0.236, "NEW TEXT", 8, MED, WHITE, { align: "center" });
  deco(s, "roundRect", 2.484, 4.9, 1.501, 0.392, { fill: ORANGE, adj: 0.5 });
  text(s, 2.617, 5.004, 1.236, 0.236, "NEW TEXT", 8, MED, INK, { align: "center" });
  deco(s, "roundRect", 9.287, 4.9, 1.501, 0.392, { fill: ORANGE, adj: 0.5 });
  text(s, 9.419, 5.004, 1.236, 0.236, "NEW TEXT", 8, MED, INK, { align: "center" });
  text(s, 4.472, 0.629, 4.39, 0.841, "INFOGRAPHIC", 44, SEMI, INK);
  icon(s, "care", 5.598, 2.874, 0.575, 0.575, WHITE, DARK);
  icon(s, "trophy", 7.483, 3.21, 0.584, 0.594, WHITE, ORANGE);
  icon(s, "people", 5.296, 4.629, 0.574, 0.57, WHITE, ORANGE);
  icon(s, "handshake", 7.183, 4.984, 0.569, 0.528, WHITE, DARK);
}

function slide32(pptx) {
  const s = newSlide(pptx);
  // --- slide content
  text(s, 4.472, 0.629, 4.39, 0.841, "INFOGRAPHIC", 44, SEMI, INK);
  deco(s, "blockArc", 3.573, 2.197, 1.889, 1.889, { fill: ORANGE2, arc: [180, 4.788], thick: 0.195 });
  deco(s, "blockArc", 5.812, 3.625, 1.889, 1.889, { fill: BLACK, flipV: true, arc: [180, 4.788], thick: 0.195 });
  deco(s, "blockArc", 8.336, 2.17, 1.889, 1.889, { fill: ORANGE2, arc: [180, 4.788], thick: 0.195 });
  deco(s, "ellipse", 2.54, 2.957, 1.815, 1.785, { fill: BLACK });
  deco(s, "ellipse", 7.063, 2.906, 1.885, 1.87, { fill: BLACK });
  deco(s, "ellipse", 9.371, 2.945, 1.893, 1.851, { fill: ORANGE });
  deco(s, "ellipse", 4.786, 2.934, 1.853, 1.833, { fill: ORANGE });
  text(s, 5.045, 4.146, 1.333, 0.303, "Lorem Ipsum", 12, SEMI, WHITE, { align: "center" });
  text(s, 7.339, 4.146, 1.333, 0.303, "Lorem Ipsum", 12, SEMI, WHITE, { align: "center" });
  text(s, 2.781, 4.146, 1.333, 0.303, "Lorem Ipsum", 12, SEMI, WHITE, { align: "center" });
  text(s, 9.651, 4.146, 1.333, 0.303, "Lorem Ipsum", 12, SEMI, WHITE, { align: "center" });
  text(s, 2.301, 5.043, 2.053, 0.988, "PLACEHOLDER", 9, BODY, INK, { align: "center", lh: 1.5 });
  text(s, 4.685, 5.731, 2.053, 0.988, "PLACEHOLDER", 9, BODY, INK, { align: "center", lh: 1.5 });
  text(s, 6.997, 5.731, 2.053, 0.988, "PLACEHOLDER", 9, BODY, INK, { align: "center", lh: 1.5 });
  text(s, 9.309, 5.043, 2.053, 0.988, "PLACEHOLDER", 9, BODY, INK, { align: "center", lh: 1.5 });
  icon(s, "tabletPie", 3.212, 3.385, 0.434, 0.704, WHITE, BLACK);
  icon(s, "monitorChart", 5.392, 3.445, 0.639, 0.584, WHITE, ORANGE);
  icon(s, "monitorPulse", 7.655, 3.45, 0.7, 0.574, WHITE, BLACK);
  icon(s, "presentation", 10.012, 3.432, 0.612, 0.61, WHITE, ORANGE);
}

function slide33(pptx) {
  const s = newSlide(pptx);
  // --- slide content
  text(s, 4.472, 0.629, 4.39, 0.841, "INFOGRAPHIC", 44, SEMI, INK);
  deco(s, "roundRect", 1.667, 2.373, 2.333, 4.004, { fill: ORANGE });
  deco(s, "roundRect", 4.256, 2.373, 2.333, 4.004, { fill: BLACK });
  deco(s, "roundRect", 6.846, 2.373, 2.333, 4.004, { fill: ORANGE });
  deco(s, "roundRect", 9.435, 2.373, 2.333, 4.004, { fill: BLACK });
  text(s, 1.976, 4.401, 1.838, 0.988, "Lorep  ipsum duis aute irure dolor in kauselih oilue epreh  deriti vols esse cill inure", 9, BODY, SILVER, { align: "center", lh: 1.5 });
  text(s, 2.375, 4.069, 0.916, 0.332, "NEW TEXT", 10, SEMI, WHITE, { lh: 1.5 });
  text(s, 4.574, 4.401, 1.838, 0.988, "Lorep  ipsum duis aute irure dolor in kauselih oilue epreh  deriti vols esse cill inure", 9, BODY, WHITE, { align: "center", lh: 1.5 });
  text(s, 4.974, 4.069, 0.916, 0.332, "NEW TEXT", 10, SEMI, WHITE, { lh: 1.5 });
  text(s, 7.126, 4.401, 1.838, 0.988, "Lorep  ipsum duis aute irure dolor in kauselih oilue epreh  deriti vols esse cill inure", 9, BODY, SILVER, { align: "center", lh: 1.5 });
  text(s, 7.525, 4.069, 0.916, 0.332, "NEW TEXT", 10, SEMI, WHITE, { lh: 1.5 });
  text(s, 9.759, 4.401, 1.838, 0.988, "Lorep  ipsum duis aute irure dolor in kauselih oilue epreh  deriti vols esse cill inure", 9, BODY, WHITE, { align: "center", lh: 1.5 });
  text(s, 10.158, 4.069, 0.916, 0.332, "NEW TEXT", 10, SEMI, WHITE, { lh: 1.5 });
  icon(s, "plane", 2.578, 3.176, 0.511, 0.529, WHITE, ORANGE);
  icon(s, "megaphone", 5.257, 3.232, 0.508, 0.473, WHITE, BLACK);
  icon(s, "globe", 7.76, 3.176, 0.505, 0.529, WHITE, ORANGE);
  icon(s, "target", 10.345, 3.176, 0.514, 0.529, WHITE, BLACK);
}

function slide34(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  deco(s, "rect", 0, 0.683, 7.438, 6.817, { fill: ORANGE });
  // picture frame (empty in the source deck): 0, 1.422  6.208 x 5.396
  deco(s, "rect", 6.5, 1.393, 6.833, 5.396, { fill: DARK });
  // --- slide content
  text(s, 6.811, 3.582, 5.821, 1.678, "“Character cannot be developed easily and quietly. Only through the experience of trial and suffering can the soul be strengthened, ambition inspired, and success achieved.”", 16, MED, WHITE, { align: "justify", lh: 1.5 });
  richText(s, 6.948, 2.156, 5.027, 0.841, [
    ["QUOTE ", { sz: 44, f: SEMI, col: WHITE }],
    ["AVDIJA", { sz: 44, f: SEMI, col: ORANGE }],
  ]);
}

function slide35(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  deco(s, "rect", 5.333, 3.104, 8, 2.292, { fill: ORANGE });
  // picture frame (empty in the source deck): 0.708, 0  6.208 x 5.396
  deco(s, "rect", 0.708, 5.729, 4.938, 1.771, { fill: DARK });
  // --- slide content
  text(s, 6.65, 4.173, 1.534, 0.303, "OUR ADDRESS", 12, SEMI, SILVER, { align: "center", wrap: false });
  richText(s, 6.711, 4.476, 1.412, 0.555, [
    ["123, Rev Avenue, ", { sz: 9, f: BODY, col: SILVER, br: true }],
    ["Kolabagan LA , 267", { sz: 9, f: BODY, col: SILVER }],
  ], { align: "center", lh: 1.5, wrap: false });
  text(s, 8.86, 4.17, 1.279, 0.303, "FOLLOW US", 12, SEMI, SILVER, { align: "center", wrap: false });
  richText(s, 8.693, 4.498, 1.613, 0.555, [
    ["www.yourtext.com", { sz: 9, f: BODY, col: SILVER, br: true }],
    ["example@gmail.com", { sz: 9, f: BODY, col: SILVER }],
  ], { align: "center", lh: 1.5, wrap: false });
  text(s, 10.815, 4.17, 1.551, 0.303, "GET IN TOUCH", 12, SEMI, SILVER, { align: "center", wrap: false });
  richText(s, 10.869, 4.473, 1.443, 0.534, [
    ["(+11) 185 6554 3435", { sz: 9, f: BODY, col: SILVER, br: true }],
    ["(+11) 189 6398 3432", { sz: 9, f: BODY, col: SILVER }],
  ], { align: "center", lh: 1.5, wrap: false });
  richText(s, 7.111, 1.242, 4.39, 0.841, [
    ["CONTACT", { sz: 44, f: SEMI, col: ORANGE }],
    [" US", { sz: 44, f: SEMI, col: INK }],
  ]);
  text(s, 6.741, 5.959, 5.625, 0.871, "Lorep  ipsum duis aute irure dolor in kauselih oiluek epreh deriti vols esse cill inure dolorlaboru sit amet. Duis autelo irusitakus rephenderi Voluptate loremil. Lorep  ipsum duis aute irure dolor in kauselih.", 10.5, BODY, INK, { align: "justify", lh: 1.5 });
  icon(s, "globePin", 7.238, 3.542, 0.358, 0.359, WHITE, ORANGE);
  icon(s, "megaphone", 9.306, 3.518, 0.388, 0.361, WHITE, ORANGE);
  icon(s, "deskPhone", 11.364, 3.518, 0.406, 0.409, WHITE, ORANGE);
}

function slide36(pptx) {
  const s = newSlide(pptx);
  // --- layout artwork
  deco(s, "parallelogram", 8.49, 0, 4.843, 2.73, { fill: DARK, adj: 0.677 });
  deco(s, "parallelogram", 0, 4.397, 4.116, 3.119, { fill: DARK, adj: 0.677 });
  deco(s, "parallelogram", 2.07, 4.397, 2.607, 3.119, { fill: ORANGE, adj: 0.805 });
  // --- slide content
  richText(s, 3.125, 2.746, 7.083, 1.649, [
    ["Thank ", { sz: 92, f: SEMI, col: INK }],
    ["You", { sz: 92, f: SEMI, col: ORANGE }],
  ], { align: "center", wrap: false });
  text(s, 4.727, 4.285, 3.88, 0.404, "Sport Presentation Template", 18, MED, INK, { align: "center" });
}
// ------------------------------------------------------------------ output
const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
  slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18,
  slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27,
  slide28, slide29, slide30, slide31, slide32, slide33, slide34, slide35, slide36,
];

function build () {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE';
  pptx.title = 'AVDIJA — Sport Presentation Template';
  BUILDERS.forEach((fn) => fn(pptx));
  return pptx.writeFile({
    fileName: path.join(__dirname, '0ea12660-ef56-4b1b-8898-8cde9cfb702f_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
