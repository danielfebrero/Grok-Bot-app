/**
 * "Sustainable Energy" deck — rebuilt with pptxgenjs.
 *
 * Slide size 13.333 x 7.5 in (16:9).  Theme "#greentosca":
 *   accent1 #22647D · accent2 #00B7C2 · accent3 #0096AE
 * Fonts: Marcellus (major / display) and Montserrat (minor / body).
 *
 * The reference deck fills many shapes with a 225deg accent1 -> accent2
 * gradient.  pptxgenjs has no gradient fill, so those shapes use BRAND,
 * the visual midpoint of the two stops.  Raster icons in the reference are
 * redrawn here from native pptxgenjs shapes (see the `ICONS` table).
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette
const C = {
  accent1: '22647D', // deep teal-blue
  accent2: '00B7C2', // bright cyan
  accent3: '0096AE',
  BRAND: '2196A7', // accent1 -> accent2 gradient, flattened
  LOGO: '1F849B', // accent1 -> accent3 gradient, flattened
  ICE: 'C8E5F0', // accent1 lum 20/80 — icon glyphs on brand fills
  ICE_CYAN: 'BCF6FF', // accent3 lum 20/80 — hairline connector
  INK: '262626', // tx1 lum 85/15 — display headings
  GRAY: 'A6A6A6', // bg1 lum 65   — body copy
  GRAY_MID: '808080', // tx1 lum 50/50 — body copy variant
  WHITE: 'FFFFFF',
};

const FONT_DISPLAY = 'Marcellus';
const FONT_BODY = 'Montserrat';
const FONT_BOLD = 'Montserrat SemiBold';

// Reusable text recipes. Sizes/faces/colours taken from the reference runs.
const T = {
  display: { fontFace: FONT_DISPLAY, fontSize: 40, color: C.INK },
  heading: { fontFace: FONT_BOLD, fontSize: 14, bold: true, color: C.accent1 },
  body: { fontFace: FONT_BODY, fontSize: 11, color: C.GRAY, lineSpacing: 22.2 },
  bodyMid: { fontFace: FONT_BODY, fontSize: 11, color: C.GRAY_MID, lineSpacing: 22.2 },
  tiny: { fontFace: FONT_BODY, fontSize: 11, color: C.GRAY },
};

const LOREM_LONG =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
  'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud ' +
  'exercitation ullamco laboris nisi ut aliquip';
const LOREM_SHORT = 'Lorem ipsum dolor sit ut amet, adipiscing elit, sed enm.';
const LOREM_CARD = 'Lorem ipsum dolor amet, consectetur veniam.';

const pptx = new PptxGenJS();
const S = pptx.ShapeType;

// -------------------------------------------------------------- primitives

/** Text box. The reference boxes are top-anchored; pptxgenjs defaults to middle. */
function txt(slide, text, opts) {
  slide.addText(text, Object.assign({ valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] }, opts));
}

function oval(slide, x, y, d, fill) {
  slide.addShape(S.ellipse, { x, y, w: d, h: d, fill: { color: fill }, line: { type: 'none' } });
}

/** Concentric decorative rings: alternating solid / dashed hairlines. */
const RING_SETS = {
  3: [0.4188, 0.6689, 1],
  4: [0.1378, 0.3907, 0.712, 1],
};
function rings(slide, cx, cy, outer, kind) {
  RING_SETS[kind].forEach((ratio, i) => {
    const d = outer * ratio;
    slide.addShape(S.ellipse, {
      x: cx - d / 2, y: cy - d / 2, w: d, h: d,
      fill: { type: 'none' },
      line: { color: C.ICE, width: 0.75, dashType: i % 2 ? 'dash' : 'solid' },
    });
  });
}

/** The "SUSTAINABLE" arch mark: a dome with an arched notch cut out of it. */
function logoMark(slide, x, y, w, h) {
  slide.addShape(S.custGeom, {
    x, y, w, h,
    fill: { color: C.LOGO },
    line: { type: 'none' },
    points: [
      { x: 0, y: h },
      { x: w, y: h, curve: { type: 'cubic', x1: 0, y1: -h * 0.42, x2: w, y2: -h * 0.42 } },
      { x: w * 0.79, y: h },
      { x: w * 0.21, y: h, curve: { type: 'cubic', x1: w * 0.66, y1: h * 0.30, x2: w * 0.34, y2: h * 0.30 } },
      { close: true },
    ],
  });
}

/** Header lockup repeated on every slide. */
function brandBar(slide) {
  logoMark(slide, 0.364, 0.359, 0.318, 0.241);
  txt(slide, 'SUSTAINABLE', {
    x: 0.682, y: 0.329, w: 2.4, h: 0.303,
    fontFace: FONT_DISPLAY, fontSize: 12, color: C.GRAY, charSpacing: 1,
  });
}

// ------------------------------------------------------------------- icons
// Each entry draws a small pictogram from native shapes inside the box
// (x, y, size) using `col` as the stroke/fill colour.

/** Pointed leaf blade: two quadratic curves meeting tip to tail. */
function leafBlade(slide, x, y, w, h, rot, col, filled) {
  slide.addShape(S.custGeom, {
    x, y, w, h, rotate: rot,
    fill: filled ? { color: col } : { type: 'none' },
    line: filled ? { type: 'none' } : { color: col, width: Math.max(1, h * 7) },
    points: [
      { x: 0, y: h },
      { x: w, y: 0, curve: { type: 'quadratic', x1: 0, y1: 0 } },
      { x: 0, y: h, curve: { type: 'quadratic', x1: w, y1: h } },
      { close: true },
    ],
  });
}

/** Horizontal bar, the building block of the wind / bench glyphs. */
function bar(slide, x, y, w, h, col) {
  slide.addShape(S.roundRect, {
    x, y, w, h, rectRadius: h / 2, fill: { color: col }, line: { type: 'none' },
  });
}

const ICONS = {
  // circled down-arrow — the "scroll" cue in the lower-left corner
  scroll(slide, x, y, s, col) {
    slide.addShape(S.ellipse, {
      x, y, w: s, h: s, fill: { type: 'none' },
      line: { color: col, width: 2.6, dashType: 'dash' },
    });
    slide.addShape(S.downArrow, {
      x: x + s * 0.28, y: y + s * 0.24, w: s * 0.44, h: s * 0.52,
      fill: { color: col }, line: { type: 'none' },
    });
  },
  // two-leaf sprout on a stem
  sprout(slide, x, y, s, col) {
    leafBlade(slide, x + s * 0.02, y + s * 0.16, s * 0.46, s * 0.30, 20, col, false);
    leafBlade(slide, x + s * 0.52, y + s * 0.16, s * 0.46, s * 0.30, 250, col, false);
    slide.addShape(S.rect, {
      x: x + s * 0.47, y: y + s * 0.34, w: s * 0.06, h: s * 0.58,
      fill: { color: col }, line: { type: 'none' },
    });
  },
  // coin with a dollar glyph
  dollar(slide, x, y, s, col) {
    slide.addShape(S.ellipse, {
      x, y, w: s, h: s, fill: { type: 'none' }, line: { color: col, width: 1.6 },
    });
    slide.addText('$', {
      x, y: y + s * 0.06, w: s, h: s * 0.88, align: 'center', valign: 'middle',
      fontFace: FONT_BODY, fontSize: s * 52, color: col, margin: 0,
    });
  },
  // recycling triangle with an inner triangle and a dot
  recycle(slide, x, y, s, col) {
    slide.addShape(S.triangle, {
      x, y: y + s * 0.12, w: s, h: s * 0.82, fill: { type: 'none' }, line: { color: col, width: 1.7 },
    });
    slide.addShape(S.triangle, {
      x: x + s * 0.46, y: y + s * 0.42, w: s * 0.48, h: s * 0.42,
      fill: { type: 'none' }, line: { color: col, width: 1.7 },
    });
    slide.addShape(S.ellipse, {
      x: x + s * 0.66, y: y, w: s * 0.28, h: s * 0.28, fill: { color: col }, line: { type: 'none' },
    });
  },
  // seedling in a flower pot
  pot(slide, x, y, s, col) {
    slide.addShape(S.trapezoid, {
      x: x + s * 0.22, y: y + s * 0.62, w: s * 0.56, h: s * 0.28, rotate: 180,
      fill: { type: 'none' }, line: { color: col, width: 2.4 },
    });
    slide.addShape(S.rect, {
      x: x + s * 0.12, y: y + s * 0.44, w: s * 0.76, h: s * 0.16,
      fill: { type: 'none' }, line: { color: col, width: 2.4 },
    });
    leafBlade(slide, x + s * 0.08, y + s * 0.10, s * 0.42, s * 0.26, 20, col, true);
    leafBlade(slide, x + s * 0.52, y + s * 0.10, s * 0.42, s * 0.26, 250, col, true);
    slide.addShape(S.rect, {
      x: x + s * 0.47, y: y + s * 0.26, w: s * 0.06, h: s * 0.20,
      fill: { color: col }, line: { type: 'none' },
    });
  },
  // broadleaf tree: three-lobed canopy over a trunk with two branches
  tree(slide, x, y, s, col) {
    const w = s, h = s * 0.66;
    slide.addShape(S.custGeom, {
      x, y, w, h, fill: { type: 'none' }, line: { color: col, width: 2.6 },
      points: [
        { x: w * 0.14, y: h },
        { x: w * 0.02, y: h * 0.62, curve: { type: 'quadratic', x1: 0, y1: h * 0.94 } },
        { x: w * 0.24, y: h * 0.16, curve: { type: 'quadratic', x1: w * 0.02, y1: h * 0.28 } },
        { x: w * 0.50, y: h * 0.02, curve: { type: 'quadratic', x1: w * 0.34, y1: 0 } },
        { x: w * 0.76, y: h * 0.16, curve: { type: 'quadratic', x1: w * 0.66, y1: 0 } },
        { x: w * 0.98, y: h * 0.62, curve: { type: 'quadratic', x1: w * 0.98, y1: h * 0.28 } },
        { x: w * 0.86, y: h, curve: { type: 'quadratic', x1: w, y1: h * 0.94 } },
        { close: true },
      ],
    });
    slide.addShape(S.rect, {
      x: x + s * 0.46, y: y + s * 0.34, w: s * 0.08, h: s * 0.66,
      fill: { color: col }, line: { type: 'none' },
    });
    [[0.28, 0.50, 32], [0.64, 0.50, -32]].forEach(([dx, dy, rot]) => {
      slide.addShape(S.rect, {
        x: x + s * dx, y: y + s * dy, w: s * 0.20, h: s * 0.07, rotate: rot,
        fill: { color: col }, line: { type: 'none' },
      });
    });
  },
  // park bench beside a shrub
  bench(slide, x, y, s, col) {
    bar(slide, x, y + s * 0.44, s * 0.56, s * 0.09, col);
    bar(slide, x + s * 0.04, y + s * 0.64, s * 0.48, s * 0.09, col);
    [0.06, 0.44].forEach(dx => slide.addShape(S.rect, {
      x: x + s * dx, y: y + s * 0.50, w: s * 0.06, h: s * 0.44,
      fill: { color: col }, line: { type: 'none' },
    }));
    slide.addShape(S.ellipse, {
      x: x + s * 0.64, y: y + s * 0.04, w: s * 0.32, h: s * 0.48,
      fill: { type: 'none' }, line: { color: col, width: 2.2 },
    });
    slide.addShape(S.rect, {
      x: x + s * 0.77, y: y + s * 0.40, w: s * 0.06, h: s * 0.54,
      fill: { color: col }, line: { type: 'none' },
    });
  },
  // ">" step separator: two bars pivoted into a caret
  chevron(slide, x, y, s, col) {
    [[0.10, 0.30, 45], [0.10, 0.66, -45]].forEach(([dx, dy, rot]) => {
      slide.addShape(S.roundRect, {
        x: x + s * dx, y: y + s * dy, w: s * 0.56, h: s * 0.10, rotate: rot,
        rectRadius: s * 0.05, fill: { color: col }, line: { type: 'none' },
      });
    });
  },
  // solar panel on a stand
  solar(slide, x, y, s, col) {
    slide.addShape(S.rect, {
      x, y, w: s, h: s * 0.66, fill: { type: 'none' }, line: { color: col, width: 2 },
    });
    slide.addShape(S.rect, {
      x, y: y + s * 0.31, w: s, h: s * 0.04, fill: { color: col }, line: { type: 'none' },
    });
    slide.addShape(S.rect, {
      x: x + s * 0.48, y: y, w: s * 0.04, h: s * 0.66, fill: { color: col }, line: { type: 'none' },
    });
    slide.addShape(S.rect, {
      x: x + s * 0.46, y: y + s * 0.66, w: s * 0.08, h: s * 0.24, fill: { color: col }, line: { type: 'none' },
    });
    slide.addShape(S.rect, {
      x: x + s * 0.20, y: y + s * 0.90, w: s * 0.60, h: s * 0.08, fill: { color: col }, line: { type: 'none' },
    });
  },
  // stylised wind gusts: three bars, each ending in a curl
  wind(slide, x, y, s, col) {
    const th = s * 0.13, r = s * 0.17;
    [[0.17, 0.30, 180, 100], [0.50, 0.62, 180, 100], [0.83, 0.20, 350, 260]]
      .forEach(([cy, len, from, to]) => {
        bar(slide, x, y + s * cy - th / 2, s * len, th, col);
        slide.addShape(S.arc, {
          x: x + s * len, y: y + s * cy - r, w: 2 * r, h: 2 * r,
          angleRange: [from, to], line: { color: col, width: th * 72 },
        });
      });
  },
  // battery cell with a bolt and a charging lead
  battery(slide, x, y, s, col) {
    slide.addShape(S.roundRect, {
      x: x + s * 0.04, y, w: s * 0.62, h: s, rectRadius: s * 0.16,
      fill: { type: 'none' }, line: { color: col, width: 2 },
    });
    slide.addShape(S.lightningBolt, {
      x: x + s * 0.24, y: y + s * 0.22, w: s * 0.22, h: s * 0.54,
      fill: { color: col }, line: { type: 'none' },
    });
    slide.addShape(S.rect, {
      x: x + s * 0.76, y: y + s * 0.26, w: s * 0.07, h: s * 0.68,
      fill: { color: col }, line: { type: 'none' },
    });
    slide.addShape(S.rect, {
      x: x + s * 0.66, y: y + s * 0.14, w: s * 0.28, h: s * 0.07,
      fill: { color: col }, line: { type: 'none' },
    });
  },
  // single leaf
  leaf(slide, x, y, s, col) {
    leafBlade(slide, x, y + s * 0.18, s, s * 0.64, 330, col, true);
  },
  // water droplet
  drop(slide, x, y, s, col) {
    slide.addShape(S.teardrop, {
      x: x + s * 0.08, y: y + s * 0.08, w: s * 0.82, h: s * 0.82, rotate: 315,
      fill: { type: 'none' }, line: { color: col, width: 2.2 },
    });
  },
  // factory silhouette: saw-tooth roof block with a chimney
  factory(slide, x, y, s, col) {
    slide.addShape(S.rect, {
      x, y: y + s * 0.42, w: s * 0.92, h: s * 0.46, fill: { type: 'none' }, line: { color: col, width: 2.2 },
    });
    slide.addShape(S.rect, {
      x: x + s * 0.62, y: y + s * 0.08, w: s * 0.22, h: s * 0.36,
      fill: { color: col }, line: { type: 'none' },
    });
    [0.12, 0.36, 0.60].forEach(dx => slide.addShape(S.rect, {
      x: x + s * dx, y: y + s * 0.58, w: s * 0.10, h: s * 0.18,
      fill: { color: col }, line: { type: 'none' },
    }));
  },
  // a small stand of conifers
  conifers(slide, x, y, s, col) {
    [[0.00, 0.26, 0.38], [0.28, 0.02, 0.44], [0.64, 0.26, 0.36]].forEach(([dx, dy, w]) => {
      slide.addShape(S.triangle, {
        x: x + s * dx, y: y + s * dy, w: s * w, h: s * (0.88 - dy),
        fill: { type: 'none' }, line: { color: col, width: 2 },
      });
    });
    slide.addShape(S.rect, {
      x: x + s * 0.46, y: y + s * 0.86, w: s * 0.07, h: s * 0.14,
      fill: { color: col }, line: { type: 'none' },
    });
  },
  // recycling arrow curling around two leaves
  loop(slide, x, y, s, col) {
    const r = s * 0.36, th = s * 0.16;
    slide.addShape(S.arc, {
      x: x + s * 0.5 - r, y: y + s * 0.5 - r, w: 2 * r, h: 2 * r,
      angleRange: [130, 60], line: { color: col, width: th * 72 },
    });
    slide.addShape(S.triangle, {
      x: x + s * 0.52, y: y + s * 0.02, w: s * 0.30, h: s * 0.26, rotate: 105,
      fill: { color: col }, line: { type: 'none' },
    });
    leafBlade(slide, x + s * 0.22, y + s * 0.48, s * 0.28, s * 0.20, 20, col, true);
    leafBlade(slide, x + s * 0.48, y + s * 0.44, s * 0.30, s * 0.21, 250, col, true);
  },
  // filled circle with a white tick
  check(slide, x, y, s, col) {
    slide.addShape(S.ellipse, { x, y, w: s, h: s, fill: { color: col }, line: { type: 'none' } });
    slide.addShape(S.rect, {
      x: x + s * 0.18, y: y + s * 0.50, w: s * 0.28, h: s * 0.11, rotate: 45,
      fill: { color: C.WHITE }, line: { type: 'none' },
    });
    slide.addShape(S.rect, {
      x: x + s * 0.38, y: y + s * 0.42, w: s * 0.44, h: s * 0.11, rotate: -45,
      fill: { color: C.WHITE }, line: { type: 'none' },
    });
  },
};
function icon(slide, kind, x, y, size, col) {
  ICONS[kind](slide, x, y, size, col);
}

// --------------------------------------------------------------- slide 1
// Landing page: nav bar, oversized wordmark and the big teal arch.
function slide1() {
  const s = pptx.addSlide();
  rings(s, 10.821, 5.062, 3.222, 4);
  brandBar(s);

  [['HOME', 7.879, 0.897], ['ABOUT', 9.200, 0.897],
   ['SERVICES', 10.434, 1.159], ['CONTACT', 11.866, 1.159]].forEach(([label, x, w]) => {
    txt(s, label, {
      x, y: 0.329, w, h: 0.286,
      fontFace: FONT_BOLD, fontSize: 11, bold: true, color: C.GRAY,
    });
  });
  icon(s, 'scroll', 0.335, 6.617, 0.488, C.accent1);

  txt(s, [
    { text: 'SUSTAINABLE', options: { breakLine: true } },
    { text: 'ENERGY' },
  ], {
    x: 1.556, y: 1.394, w: 5.574, h: 1.582,
    fontFace: FONT_BODY, fontSize: 44, bold: true, color: C.accent1, charSpacing: 1.5,
  });
  txt(s, 'SUSTAINABLE ENERGY PRESENTATION TEMPLATE ', {
    x: 7.451, y: 1.394, w: 3.741, h: 0.64,
    fontFace: FONT_BODY, fontSize: 16, bold: true, color: C.accent1, charSpacing: 1,
  });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore', {
    x: 7.451, y: 2.033, w: 4.594, h: 0.679, ...T.bodyMid,
  });

  // Semicircular arch whose legs run off the bottom of the slide.
  s.addShape(S.arc, {
    x: 2.306, y: 3.079, w: 8.720, h: 8.720, angleRange: [180, 360],
    line: { color: C.BRAND, width: 15 },
  });
  rings(s, 3.0825, 5.0115, 1.995, 3);
}

// --------------------------------------------------------------- slide 2
// "Benefits": display title plus three icon/label/copy rows.
const BENEFITS = [
  { y: 3.523, title: 'Reduces Pollution', glyph: 'sprout', gx: 1.805, gy: 3.697, gs: 0.317 },
  { y: 4.528, title: 'Saves Money', glyph: 'dollar', gx: 1.790, gy: 4.688, gs: 0.348 },
  { y: 5.534, title: 'Creates Jobs', glyph: 'recycle', gx: 1.803, gy: 5.683, gs: 0.321 },
];
function slide2() {
  const s = pptx.addSlide();
  rings(s, 7.5975, 4.5715, 1.779, 3);
  rings(s, 11.187, 1.893, 3.222, 4);
  brandBar(s);

  txt(s, 'The Sustainable Energy Benefits', { x: 1.631, y: 1.356, w: 4.113, h: 1.447, ...T.display });

  BENEFITS.forEach(b => {
    oval(s, 1.631, b.y, 0.667, C.BRAND);
    icon(s, b.glyph, b.gx, b.gy, b.gs, C.ICE);
    txt(s, b.title, { x: 2.346, y: b.y + 0.003, w: 2.481, h: 0.337, ...T.heading });
    txt(s, 'Lorem ipsum dolor tempor eiusmod veniam.',
      { x: 2.344, y: b.y + 0.377, w: 3.878, h: 0.286, ...T.tiny });
  });

  icon(s, 'scroll', 0.335, 6.617, 0.488, C.accent1);
  oval(s, 6.649, 0.631, 0.836, C.BRAND);
  icon(s, 'pot', 6.817, 0.799, 0.501, C.WHITE);
}

// --------------------------------------------------------------- slide 3
// "Big Challenges": three numbered steps rising to the right, giant ring,
// and a single-corner-rounded gradient card carrying the title.
const CHALLENGES = [
  { n: '03', cx: 2.950, cy: 1.327, tx: 4.286, ty: 1.332, tw: 2.481,
    title: 'Public Acceptance', body: 'Lorem ipsum dolor sit amet, adipiscing elit eiusmod.' },
  { n: '01', cx: 7.899, cy: 5.218, tx: 9.236, ty: 5.223, tw: 2.481,
    title: 'Infrastucture', body: 'Lorem ipsum dolor sit amet, adipiscing elit eiusmod.' },
  { n: '02', cx: 4.887, cy: 3.493, tx: 6.219, ty: 3.503, tw: 1.808,
    title: 'The Cost', body: 'Lorem ipsum dolor amet eiusmod.' },
];
function slide3() {
  const s = pptx.addSlide();
  rings(s, 1.424, 5.073, 4.624, 4);
  s.addShape(S.round1Rect, {
    x: 0, y: 5.103, w: 6.542, h: 2.397, rectRadius: 0.2467,
    fill: { color: C.BRAND }, line: { type: 'none' },
  });
  brandBar(s);

  CHALLENGES.forEach(c => {
    txt(s, c.body, { x: c.tx, y: c.ty + 0.381, w: c.tw, h: 0.679, ...T.bodyMid });
    txt(s, c.title, { x: c.tx, y: c.ty, w: 2.481, h: 0.337, ...T.heading });
    oval(s, c.cx, c.cy, 1.070, C.BRAND);
    txt(s, c.n, {
      x: c.cx - 0.046, y: c.cy + 0.249, w: 1.161, h: 0.572, align: 'center',
      fontFace: FONT_DISPLAY, fontSize: 28, color: C.WHITE,
    });
  });

  txt(s, 'Sustainable Energy Big Challenges',
    { x: 1.556, y: 5.477, w: 4.763, h: 1.447, ...T.display, color: C.WHITE });

  // Oversized ring bleeding off the top-right corner.
  s.addShape(S.ellipse, {
    x: 7.337, y: -2.770, w: 7.393, h: 7.393,
    fill: { type: 'none' }, line: { color: C.BRAND, width: 39.5 },
  });
  rings(s, 13.07, 7.287, 2.174, 3);
  icon(s, 'scroll', 0.335, 6.617, 0.488, C.ICE);
}

// --------------------------------------------------------------- slide 4
// "How it works": copy column on the left, two icon notes on the right.
// Soft drop shadow under the white cards. pptxgenjs rewrites the object it is
// handed into EMU, so hand out a fresh one every time.
const cardShadow = () =>
  ({ type: 'outer', blur: 12, offset: 4, angle: 90, color: '000000', opacity: 0.2 });
function slide4() {
  const s = pptx.addSlide();
  rings(s, 6.4715, 1.2995, 1.527, 3);
  s.addShape(S.roundRect, {
    x: 6.471, y: 1.153, w: 3.556, h: 2.431, rectRadius: 0.2403,
    fill: { color: C.WHITE }, line: { type: 'none' }, shadow: cardShadow(),
  });
  brandBar(s);

  txt(s, 'How Sustainable Energy Works', { x: 1.469, y: 1.835, w: 5.003, h: 1.447, ...T.display });
  txt(s, LOREM_LONG + '..', { x: 1.469, y: 3.890, w: 4.771, h: 1.296, ...T.body });

  s.addShape(S.roundRect, {
    x: 1.486, y: 5.731, w: 2.007, h: 0.429, rectRadius: 0.2145,
    fill: { color: C.BRAND }, line: { type: 'none' },
  });
  txt(s, 'More Details', {
    x: 1.745, y: 5.794, w: 1.489, h: 0.303, align: 'center',
    fontFace: FONT_BOLD, fontSize: 12, bold: true, color: C.WHITE,
  });

  [{ y: 1.970, w: 2.195, title: 'The Trend Towards Sustain Energy' },
   { y: 4.687, w: 2.362, title: 'Sustainable Energy Is The Real Future' }].forEach(n => {
    txt(s, n.title, { x: 6.966, y: n.y, w: n.w, h: 0.572, ...T.heading });
    txt(s, LOREM_SHORT, { x: 6.966, y: n.y + 0.637, w: 2.568, h: 0.679, ...T.body });
  });
  icon(s, 'tree', 7.077, 1.451, 0.456, C.accent1);
  icon(s, 'bench', 7.063, 4.071, 0.552, C.accent1);

  rings(s, 5.215, 7.500, 3.284, 4);
  icon(s, 'scroll', 0.335, 6.617, 0.488, C.accent1);
}

// --------------------------------------------------------------- slide 5
// Two stat cards beside a right-hand title block.
function slide5() {
  const s = pptx.addSlide();
  rings(s, 12.1855, 4.6405, 1.227, 3);
  rings(s, 5.9565, 6.0415, 2.245, 4);
  brandBar(s);

  txt(s, 'What Is Solar Energy About?', { x: 7.254, y: 3.868, w: 4.113, h: 1.447, ...T.display });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
    'incididunt ut labore et dolore magna aliqua. Ut enim minim veniam, quis.',
    { x: 7.254, y: 5.605, w: 4.530, h: 0.988, ...T.body });

  s.addShape(S.roundRect, {
    x: 3.830, y: 4.291, w: 2.167, h: 1.878, rectRadius: 0.1856,
    fill: { color: C.WHITE }, line: { type: 'none' }, shadow: cardShadow(),
  });
  s.addShape(S.roundRect, {
    x: 1.488, y: 4.291, w: 2.167, h: 1.878, rectRadius: 0.2158,
    fill: { color: C.BRAND }, line: { type: 'none' },
  });

  [{ x: 1.761, y: 4.709, big: '735+', small: 'The Facilities', col: C.WHITE, lab: C.WHITE },
   { x: 4.103, y: 4.706, big: '82.3K', small: 'Total People', col: C.accent1, lab: C.accent1 }].forEach(k => {
    txt(s, k.big, {
      x: k.x, y: k.y, w: 1.621, h: 0.774,
      fontFace: FONT_DISPLAY, fontSize: 40, bold: true, color: k.col,
    });
    txt(s, k.small, {
      x: k.x, y: k.y + 0.708, w: 1.621, h: 0.337,
      fontFace: FONT_BOLD, fontSize: 14, bold: true, color: k.lab,
    });
  });
  icon(s, 'scroll', 0.335, 6.617, 0.488, C.accent1);
}

// --------------------------------------------------------------- slide 6
// Four-step process chain: circles on a hairline, captions alternating
// above and below the axis.
const PROCESS = [
  { x: 2.141, title: 'Carbondioxide', above: true, glyph: 'drop', gx: 2.586, gy: 4.312, gs: 0.562 },
  { x: 4.748, title: 'Biomass Process', above: false, glyph: 'factory', gx: 5.137, gy: 4.312, gs: 0.562 },
  { x: 7.355, title: 'Trees and Plants', above: true, glyph: 'conifers', gx: 7.687, gy: 4.312, gs: 0.606 },
  { x: 9.961, title: 'The Biomass', above: false, glyph: 'loop', gx: 10.364, gy: 4.291, gs: 0.575 },
];
function slide6() {
  const s = pptx.addSlide();
  brandBar(s);
  txt(s, 'Sustainable Energy Is The Real Future', {
    x: 2.435, y: 0.918, w: 8.619, h: 0.774, align: 'center', ...T.display,
  });
  s.addShape(S.line, {
    x: 2.141, y: 4.594, w: 9.161, h: 0, line: { color: C.ICE_CYAN, width: 1.5 },
  });

  PROCESS.forEach(p => {
    oval(s, p.x, 3.923, 1.341, C.BRAND);
    icon(s, p.glyph, p.gx, p.gy, p.gs, C.ICE);
    const headY = p.above ? 2.749 : 5.418;
    txt(s, p.title, {
      x: p.x - 0.570, y: headY, w: 2.481, h: 0.337, align: 'center', ...T.heading,
    });
    txt(s, 'Lorem ipsum dolor enim amet, adipiscing elit eiusmod.', {
      x: p.x - 0.639, y: headY + 0.341, w: 2.618, h: 0.679, align: 'center', ...T.body,
    });
  });

  [3.753, 6.434, 9.116].forEach(x => icon(s, 'chevron', x, 4.343, 0.501, C.accent3));
  rings(s, 13.3275, 2.1445, 3.799, 3);
  rings(s, 0.875, 6.809, 3.284, 4);
}

// --------------------------------------------------------------- slide 7
// Statement slide: left display title, right paragraph plus two ticked tags.
function slide7() {
  const s = pptx.addSlide();
  rings(s, 5.6005, 3.2205, 2.131, 4);
  brandBar(s);

  [{ x: 7.414, w: 1.651, label: 'Wind Energy', ix: 7.069 },
   { x: 9.627, w: 1.949, label: 'Biomass Energy', ix: 9.282 }].forEach(tag => {
    txt(s, tag.label, { x: tag.x, y: 5.959, w: tag.w, h: 0.337, ...T.heading });
    icon(s, 'check', tag.ix, 5.976, 0.281, C.BRAND);
  });

  txt(s, 'Trend Towards This Idea Is Clear', { x: 1.469, y: 1.273, w: 4.661, h: 1.447, ...T.display });
  txt(s, LOREM_LONG + ' ex ea commodo.', { x: 6.929, y: 4.286, w: 5.003, h: 1.296, ...T.body });
  rings(s, 12.719, 6.861, 2.286, 3);
}

// --------------------------------------------------------------- slide 8
// 2x2 card grid; alternate cards are filled, the others white with a shadow.
const TYPE_CARDS = [
  { x: 6.662, y: 0.941, filled: true, radius: 0.2021, title: 'Solar Energy',
    glyph: 'solar', gx: 7.780, gy: 1.497, gs: 0.330, ax: 7.526, ay: 1.240 },
  { x: 9.392, y: 0.941, filled: false, radius: 0.2190, title: 'Wind Energy',
    glyph: 'wind', gx: 10.508, gy: 1.511, gs: 0.334, ax: 10.257, ay: 1.243 },
  { x: 6.662, y: 3.853, filled: false, radius: 0.2190, title: 'Hydroelectric',
    glyph: 'bench', gx: 7.665, gy: 4.366, gs: 0.378, ax: 7.436, ay: 4.158 },
  { x: 9.392, y: 3.868, filled: true, radius: 0.2021, title: 'Geothermal',
    glyph: 'battery', gx: 10.479, gy: 4.432, gs: 0.303, ax: 10.212, ay: 4.176 },
];
function slide8() {
  const s = pptx.addSlide();
  rings(s, 5.016, 4.081, 2.286, 3);
  rings(s, 12.1675, 2.1865, 1.975, 4);
  brandBar(s);
  txt(s, 'Sustainable Energy Types', { x: 1.549, y: 1.353, w: 4.415, h: 1.447, ...T.display });

  TYPE_CARDS.forEach(c => {
    s.addShape(S.roundRect, Object.assign({
      x: c.x, y: c.y, w: 2.569, h: 2.712, rectRadius: c.radius, line: { type: 'none' },
      fill: { color: c.filled ? C.BRAND : C.WHITE },
    }, c.filled ? {} : { shadow: cardShadow() }));
    // Open ring behind the glyph (302deg sweep starting at the top).
    s.addShape(S.arc, {
      x: c.ax, y: c.ay, w: 0.837, h: 0.837, angleRange: [270, 212.2756],
      line: { color: c.filled ? C.ICE : C.BRAND, width: 15 },
    });
    icon(s, c.glyph, c.gx, c.gy, c.gs, c.filled ? C.ICE : C.accent1);
    txt(s, c.title, {
      x: c.x + 0.046, y: c.y + 1.367, w: 2.481, h: 0.337, align: 'center',
      ...T.heading, color: c.filled ? C.WHITE : C.accent1,
    });
    txt(s, LOREM_CARD, {
      x: c.x + 0.044, y: c.y + 1.735, w: 2.481, h: 0.679, align: 'center',
      ...T.body, color: c.filled ? C.WHITE : C.GRAY_MID,
    });
  });
}

// --------------------------------------------------------------- slide 9
// Display title with a pill-shaped highlight panel.
function slide9() {
  const s = pptx.addSlide();
  brandBar(s);
  txt(s, [
    { text: 'Sustainable Energy ', options: { breakLine: true } },
    { text: 'Is The Real Future' },
  ], { x: 1.549, y: 2.170, w: 5.340, h: 1.447, ...T.display });

  s.addShape(S.roundRect, {
    x: 5.831, y: 4.164, w: 5.954, h: 2.107, rectRadius: 1.0535,
    fill: { color: C.BRAND }, line: { type: 'none' },
  });
  txt(s, 'Sustainable Development', {
    x: 6.508, y: 4.520, w: 3.096, h: 0.337, ...T.heading, color: C.WHITE,
  });
  txt(s, LOREM_LONG.replace(' exercitation ullamco laboris nisi ut aliquip', '') + '.', {
    x: 6.508, y: 4.913, w: 5.003, h: 0.988, ...T.body, color: C.WHITE,
  });

  rings(s, 12.717, 6.894, 2.522, 4);
  rings(s, 4.410, 0.165, 3.564, 3);
}

// -------------------------------------------------------------- slide 10
// Two overlapping circular stat bubbles plus a right-hand note.
function slide10() {
  const s = pptx.addSlide();
  rings(s, 1.002, 6.706, 2.994, 3);
  brandBar(s);
  txt(s, 'Industrial Can Help Us to Save Planet',
    { x: 1.454, y: 1.638, w: 5.070, h: 1.447, ...T.display });

  oval(s, 1.071, 3.371, 3.592, C.BRAND);
  txt(s, '550+', {
    x: 2.286, y: 3.966, w: 1.274, h: 0.707, align: 'center',
    fontFace: FONT_DISPLAY, fontSize: 36, color: C.ICE,
  });
  txt(s, 'Contributions', {
    x: 1.867, y: 4.725, w: 1.999, h: 0.337, align: 'center',
    fontFace: FONT_BOLD, fontSize: 14, bold: true, color: C.WHITE,
  });
  txt(s, 'Lorem ipsum dolor sit amet elit, sed do eiusmod tempor ut labore enim magna aliqua veniam sed.',
    { x: 1.508, y: 5.072, w: 2.718, h: 1.296, align: 'center', ...T.body, color: C.WHITE });

  oval(s, 4.802, 4.091, 2.152, C.BRAND);
  icon(s, 'leaf', 5.582, 4.505, 0.593, C.ICE);
  txt(s, 'Lorem ipsum dolor sit amet elit, sed.',
    { x: 4.996, y: 5.124, w: 1.783, h: 0.679, align: 'center', ...T.body, color: C.WHITE });

  txt(s, 'Lorem ipsum dolor sit amet veniam, sed do eiusmod tempor ut labore enim magna aliqua veniam sed veniam enim minim do.',
    { x: 7.431, y: 1.496, w: 4.348, h: 0.988, ...T.body });
  txt(s, 'Lorem Ipsum Dolor Sit Amet Enim Minim.',
    { x: 7.431, y: 2.799, w: 4.906, h: 0.337, ...T.heading });

  rings(s, 6.779, 0.000, 3.342, 4);
}

// ------------------------------------------------------------------ build
pptx.defineLayout({ name: 'DECK', width: 13.333, height: 7.5 });
pptx.layout = 'DECK';
pptx.theme = { headFontFace: FONT_DISPLAY, bodyFontFace: FONT_BODY };
pptx.title = 'Sustainable Energy Presentation Template';

[slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10]
  .forEach(build => build());

pptx.writeFile({
  fileName: path.join(__dirname, '0d8ee907-c808-43a3-b0fe-672118007a16_grok_final.pptx'),
}).then(f => console.log('wrote', f));
