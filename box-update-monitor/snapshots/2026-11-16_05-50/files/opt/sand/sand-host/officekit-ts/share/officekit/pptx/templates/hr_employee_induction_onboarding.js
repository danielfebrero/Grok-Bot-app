#!/usr/bin/env node
/**
 * "Induction" presentation template - rebuilt with pptxgenjs.
 *
 * 20 slides, 10 x 5.625in (16:9). Raster photography in the source deck is
 * replaced with flat "[image]" placeholders; every other element (the wave
 * motifs, arches, bars, cards, rules and type) is drawn natively.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const C = {
  blue: '4793AF', // accent1 - primary brand blue
  cream: 'FFE9C9', // accent2
  red: 'DD5746', // accent3
  white: 'FFFFFF',
  black: '000000',
  ink: '0C0C0C', // body copy on light backgrounds
  blueLt: 'B3D3E0', // chart tints
  blueMd: '8DBFD1',
  photo: 'F2F2F2', // image-placeholder fill
  photoInk: '9AA0A6', // image-placeholder caption
  lid: '333333', // laptop mockup: lid shell, then the two base slabs
  deck: '6A6967',
  wedge: '565553',
};

const F = { head: 'Bricolage Grotesque', body: 'Open Sans', num: 'Sora' };

/** Text insets used throughout the source deck: 0.075in / 0.0375in, in points. */
const INSET = [5.4, 5.4, 2.7, 2.7];

const SIZE = { w: 10, h: 5.625 };

/* -------------------------------------------------------------- text tools */

function text(slide, body, opts) {
  slide.addText(body, Object.assign({ margin: INSET, valign: 'top', isTextBox: true }, opts));
}

/** Display heading - Bricolage Grotesque, 90% leading. */
function title(slide, body, opts) {
  text(slide, body, Object.assign({ fontFace: F.head, fontSize: 50, lineSpacingMultiple: 0.9 }, opts));
}

/** Small caps-height label / eyebrow. */
function label(slide, body, opts) {
  text(slide, body, Object.assign({ fontFace: F.head, fontSize: 14 }, opts));
}

/** Running copy - Open Sans, 130% leading. */
function copy(slide, body, opts) {
  text(slide, body, Object.assign({ fontFace: F.body, fontSize: 11, lineSpacingMultiple: 1.3 }, opts));
}

/** Big "40,5%" style figure: full-size number with a smaller percent sign. */
function figure(slide, value, opts) {
  const runs = [
    { text: value, options: { fontSize: 86 } },
    { text: '%', options: { fontSize: 60 } },
    { text: '  ', options: { fontSize: 86 } },
  ];
  text(slide, runs, Object.assign({ fontFace: F.head, color: C.white }, opts));
}

/* ---------------------------------------------------------- path utilities */

const S45 = Math.SQRT1_2;
const KCIR = 0.5522847; // cubic handle length for a 90-deg arc

/** Appends a circular arc (a0 -> a1, radians) as one cubic segment. */
function arcSeg(pts, cx, cy, r, a0, a1) {
  const dir = a1 > a0 ? 1 : -1;
  const k = r * (4 / 3) * Math.tan((a1 - a0) / 4) * dir;
  const x0 = cx + r * Math.cos(a0);
  const y0 = cy + r * Math.sin(a0);
  const x1 = cx + r * Math.cos(a1);
  const y1 = cy + r * Math.sin(a1);
  pts.push({
    x: x1,
    y: y1,
    curve: {
      type: 'cubic',
      x1: x0 - k * Math.sin(a0) * dir,
      y1: y0 + k * Math.cos(a0) * dir,
      x2: x1 + k * Math.sin(a1) * dir,
      y2: y1 - k * Math.cos(a1) * dir,
    },
  });
}

/** Walks a point list backwards, mirroring cubic control points. */
function reversePath(pts) {
  const out = [];
  for (let i = pts.length - 1; i > 0; i--) {
    const seg = pts[i];
    const prev = pts[i - 1];
    out.push(
      seg.curve
        ? {
            x: prev.x,
            y: prev.y,
            curve: { type: 'cubic', x1: seg.curve.x2, y1: seg.curve.y2, x2: seg.curve.x1, y2: seg.curve.y1 },
          }
        : { x: prev.x, y: prev.y }
    );
  }
  return out;
}

/** Rectangle with a semicircular top - the deck's arched photo frames. */
function archPath(w, h) {
  const r = w / 2;
  const k = KCIR * r;
  return [
    { x: 0, y: h, moveTo: true },
    { x: 0, y: r },
    { x: r, y: 0, curve: { type: 'cubic', x1: 0, y1: r - k, x2: r - k, y2: 0 } },
    { x: w, y: r, curve: { type: 'cubic', x1: r + k, y1: 0, x2: w, y2: r - k } },
    { x: w, y: h },
    { close: true },
  ];
}

/** Rectangle with one stadium end - progress bars and wide photo frames. */
function pillPath(w, h, roundedSide) {
  const r = h / 2;
  const k = KCIR * r;
  if (roundedSide === 'left') {
    return [
      { x: w, y: 0, moveTo: true },
      { x: r, y: 0 },
      { x: 0, y: r, curve: { type: 'cubic', x1: r - k, y1: 0, x2: 0, y2: r - k } },
      { x: r, y: h, curve: { type: 'cubic', x1: 0, y1: r + k, x2: r - k, y2: h } },
      { x: w, y: h },
      { close: true },
    ];
  }
  return [
    { x: 0, y: 0, moveTo: true },
    { x: w - r, y: 0 },
    { x: w, y: r, curve: { type: 'cubic', x1: w - r + k, y1: 0, x2: w, y2: r - k } },
    { x: w - r, y: h, curve: { type: 'cubic', x1: w, y1: r + k, x2: w - r + k, y2: h } },
    { x: 0, y: h },
    { close: true },
  ];
}

/** Filled box with an elliptical bite taken out of it - clips the wave disc. */
function discMask(boxW, boxH, cx, cy, rx, ry) {
  const kx = KCIR * rx;
  const ky = KCIR * ry;
  return [
    { x: 0, y: 0, moveTo: true },
    { x: boxW, y: 0 },
    { x: boxW, y: boxH },
    { x: 0, y: boxH },
    { close: true },
    { x: cx, y: cy - ry, moveTo: true },
    { x: cx - rx, y: cy, curve: { type: 'cubic', x1: cx - kx, y1: cy - ry, x2: cx - rx, y2: cy - ky } },
    { x: cx, y: cy + ry, curve: { type: 'cubic', x1: cx - rx, y1: cy + ky, x2: cx - kx, y2: cy + ry } },
    { x: cx + rx, y: cy, curve: { type: 'cubic', x1: cx + kx, y1: cy + ry, x2: cx + rx, y2: cy + ky } },
    { x: cx, y: cy - ry, curve: { type: 'cubic', x1: cx + rx, y1: cy - ky, x2: cx + kx, y2: cy - ry } },
    { close: true },
  ];
}

/* --------------------------------------------------------- the wave ribbon */

/**
 * The deck's signature "≈" stroke: 45-degree straights joined by quarter-round
 * crests and troughs. Lengths are fractions of the wave unit `u`, which equals
 * the width of a two-crest ribbon.
 */
const WAVE = { R: 0.208, RUN: 0.1087, THICK: 0.1156, PITCH: 0.2899 };
WAVE.STEP = 2 * WAVE.R * S45 + WAVE.RUN; // crest -> trough, horizontally
WAVE.AMP = 2 * WAVE.R * (1 - S45) + WAVE.RUN; // crest -> trough, vertically
WAVE.LEAD = (WAVE.R + WAVE.THICK / 2) * S45 + WAVE.RUN; // overhang past an end apex
WAVE.RIB = WAVE.AMP + WAVE.THICK; // one ribbon's height

const ribbonWidth = (halves, u) => ((halves - 1) * WAVE.STEP + 2 * WAVE.LEAD) * u;

/** One side of a ribbon, walked left to right; `off` is signed from the centreline. */
function waveEdge(x0, yMid, u, halves, off) {
  const pts = [];
  for (let k = 0; k < halves; k++) {
    const crest = k % 2 === 0;
    const cx = x0 + k * WAVE.STEP * u;
    const cy = yMid + (crest ? WAVE.R : WAVE.AMP - WAVE.R) * u;
    const r = (crest ? WAVE.R - off : WAVE.R + off) * u;
    const a0 = (crest ? -0.75 : 0.75) * Math.PI;
    const a1 = (crest ? -0.25 : 0.25) * Math.PI;
    const sx = cx + r * Math.cos(a0);
    const sy = cy + r * Math.sin(a0);
    const run = WAVE.RUN * u;
    if (k === 0) pts.push({ x: sx - run, y: sy + (crest ? run : -run), moveTo: true });
    pts.push({ x: sx, y: sy });
    arcSeg(pts, cx, cy, r, a0, a1);
    if (k === halves - 1) {
      pts.push({ x: cx + r * Math.cos(a1) + run, y: cy + r * Math.sin(a1) + (crest ? run : -run) });
    }
  }
  return pts;
}

/**
 * A closed ribbon spanning `halves` half-waves. `xCrest` is the x of the first
 * crest apex; the stroke overhangs it by WAVE.LEAD on each end.
 */
function ribbon(xCrest, yTop, u, halves) {
  const yMid = yTop + (WAVE.THICK / 2) * u;
  const upper = waveEdge(xCrest, yMid, u, halves, -WAVE.THICK / 2);
  const lower = waveEdge(xCrest, yMid, u, halves, WAVE.THICK / 2);
  const tail = lower[lower.length - 1];
  return upper.concat([{ x: tail.x, y: tail.y }], reversePath(lower), [{ close: true }]);
}

/** Scales a point list so its natural extent fills exactly w x h. */
function fitPoints(pts, natW, natH, w, h) {
  const sx = w / natW;
  const sy = h / natH;
  return pts.map((pt) =>
    pt.close
      ? pt
      : Object.assign({}, pt, { x: pt.x * sx, y: pt.y * sy },
          pt.curve
            ? { curve: { type: 'cubic', x1: pt.curve.x1 * sx, y1: pt.curve.y1 * sy, x2: pt.curve.x2 * sx, y2: pt.curve.y2 * sy } }
            : null)
  );
}

/** Three stacked ribbons - the corner motif used on most content slides. */
function waveMark(slide, x, y, w, h, color) {
  const natW = WAVE.STEP + 2 * WAVE.LEAD;
  const natH = 2 * WAVE.PITCH + WAVE.RIB;
  let pts = [];
  for (let i = 0; i < 3; i++) pts = pts.concat(ribbon(WAVE.LEAD, i * WAVE.PITCH, 1, 2));
  slide.addShape('custGeom', {
    x, y, w, h, points: fitPoints(pts, natW, natH, w, h), fill: { color }, line: { type: 'none' },
  });
}

/**
 * Six wider ribbons clipped to a disc - the title/closing motif. Drawn as an
 * oversized ribbon field, then a background-coloured plate with a round hole
 * punched in it does the clipping. All fractions are of the disc diameter.
 */
const DISC = { U: 0.6063, CREST: 0.1332, TOP: -0.0372, PITCH: 0.1729, PAD: 0.1 };

function waveDisc(slide, x, y, d, color, bg) {
  const u = DISC.U * d;
  const pad = DISC.PAD * d;
  let pts = [];
  for (let i = 0; i < 6; i++) {
    pts = pts.concat(ribbon(pad + DISC.CREST * d, pad + (DISC.TOP + i * DISC.PITCH) * d, u, 4));
  }
  const box = { x: x - pad, y: y - pad, w: d + 2 * pad, h: d + 2 * pad };
  slide.addShape('custGeom', Object.assign({ points: pts, fill: { color }, line: { type: 'none' } }, box));
  slide.addShape('custGeom', Object.assign(
    { points: discMask(box.w, box.h, box.w / 2, box.h / 2, d / 2, d / 2), fill: { color: bg }, line: { type: 'none' } },
    box
  ));
}

/* -------------------------------------------------------- image stand-ins */

/** Flat placeholder standing in for a photograph. `shape`: rect|arch|ellipse|pillLeft */
function photo(slide, x, y, w, h, shape) {
  const base = { x, y, w, h, fill: { color: C.photo }, line: { type: 'none' } };
  if (shape === 'arch') slide.addShape('custGeom', Object.assign({ points: archPath(w, h) }, base));
  else if (shape === 'ellipse') slide.addShape('ellipse', base);
  else if (shape === 'pillLeft') slide.addShape('custGeom', Object.assign({ points: pillPath(w, h, 'left') }, base));
  else slide.addShape('rect', base);
  slide.addText('[image]', {
    x, y, w, h, align: 'center', valign: 'middle',
    fontFace: F.body, fontSize: w < 2.2 ? 9 : 11, color: C.photoInk,
  });
}

/* ------------------------------------------------------------ page chrome */

function page(pptx, bg, numColor) {
  const slide = pptx.addSlide();
  slide.background = { color: bg };
  if (numColor) {
    slide.slideNumber = {
      x: 9.353, y: 5.354, w: 0.527, h: 0.202,
      align: 'right', fontFace: F.head, fontSize: 8, color: numColor, margin: INSET,
    };
  }
  return slide;
}

/* ------------------------------------------------------- reusable clusters */

/** Numbered column: big ordinal, heading, paragraph. Used on slides 5 and 10. */
function numberedColumn(slide, x, item) {
  text(slide, item.n, { x, y: 3.31, w: item.nw, h: 0.757, fontFace: F.head, fontSize: 41, color: C.white });
  label(slide, item.head, { x, y: 4.128, w: 2.022, h: 0.303, color: C.white });
  copy(slide, item.body, { x, y: 4.491, w: item.bw, h: 0.742, color: C.white });
}

const STEPS = [
  { n: '01', nw: 0.742, head: 'Your New Beginning' },
  { n: '02', nw: 1.034, head: 'Your First Day' },
];
const STEP_BODY = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit sed do eiusmod tempor.';

/** Right-aligned "2025 / kicker / paragraph" stack shared by slides 6 and 12. */
function yearBlock(slide, y, kicker, body, bodyBox) {
  text(slide, '2025', { x: 6.659, y, w: 2.862, h: 0.833, fontFace: F.head, fontSize: 45, color: C.white, align: 'right' });
  label(slide, kicker, { x: 6.503, y: y + 0.882, w: 3.019, h: 0.328, fontSize: 15, color: C.white, align: 'right' });
  copy(slide, body, Object.assign({ h: 0.502, color: C.white, align: 'right' }, bodyBox));
}

/* ------------------------------------------------------------------ slides */

function slide01(pptx) {
  const s = page(pptx, C.blue, null);
  waveDisc(s, 6.512, 0.451, 2.917, C.cream, C.blue);
  text(s, 'Induction', { x: 0.392, y: 3.787, w: 7.108, h: 1.818, fontFace: F.head, fontSize: 104, color: C.white });
  text(s, 'Presentation Template', { x: 0.476, y: 3.228, w: 5.674, h: 0.682, fontFace: F.head, fontSize: 36, color: C.white });
  label(s, 'Embarking on a New Journey', { x: 0.485, y: 0.375, w: 3.117, h: 0.303, color: C.white });
  copy(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod', {
    x: 0.485, y: 0.649, w: 3.619, h: 0.535, color: C.white,
  });
}

function slide02(pptx) {
  const s = page(pptx, C.blue, C.white);
  waveMark(s, 2.734, 3.791, 1.492, 1.383, C.cream);
  title(s, '\u201CThe goal of an induction program is to ensure that individuals feel  prepared to succeed in their new roles.\u201D', {
    x: 2.726, y: 0.307, w: 6.811, h: 3.484, fontSize: 41, color: C.white, align: 'right', lineSpacingMultiple: 1,
  });
  s.addShape('line', { x: 4.595, y: 4.288, w: 3.462, h: 0, line: { color: C.white, width: 1 } });
  label(s, 'Introduction', { x: 8.144, y: 4.124, w: 1.392, h: 0.328, fontSize: 15, color: C.white });
  photo(s, 0.57, 1.391, 2.333, 3.783, 'arch');
}

function slide03(pptx) {
  const s = page(pptx, C.red, C.ink);
  waveMark(s, 0.776, 3.924, 1.854, 1.719, C.cream);
  title(s, 'Embarking on a New Journey', { x: 0.445, y: 0.341, w: 5.147, h: 1.591, color: C.white });
  label(s, 'Your Journey Begins Here', { x: 6.417, y: 0.378, w: 3.098, h: 0.328, fontSize: 15, color: C.white, align: 'right' });
  copy(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore et dolore magna aliqua enim ad minim veniam.', {
    x: 5.592, y: 0.719, w: 3.91, h: 0.742, color: C.white, align: 'right',
  });
  photo(s, 2.48, 2.431, 7.52, 3.194, 'rect');
}

function slide04(pptx) {
  const s = page(pptx, C.cream, C.blue);
  title(s, 'The First Step to Your Bright Future', { x: 0.497, y: 3.714, w: 6.147, h: 1.575, color: C.blue });
  waveMark(s, -0.168, 0.456, 1.633, 1.515, C.blue);

  const cards = [
    { head: 'Strategy One', x: 3.755, y: 0.451 },
    { head: 'Strategy Three', x: 3.755, y: 1.641 },
    { head: 'Strategy Two', x: 6.755, y: 0.451 },
    { head: 'Strategy Four', x: 6.755, y: 1.641 },
  ];
  cards.forEach((card) => {
    s.addShape('rect', { x: card.x, y: card.y, w: 2.674, h: 1.028, fill: { color: C.blue }, line: { type: 'none' } });
    label(s, card.head, { x: card.x + 0.077, y: card.y + 0.129, w: 1.932, h: 0.303, color: C.white });
    copy(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eiusmod tempor.', {
      x: card.x + 0.077, y: card.y + 0.429, w: 2.519, h: 0.47, fontSize: 9, color: C.white,
    });
  });

  label(s, 'Embrace the Future', { x: 6.755, y: 4.42, w: 2.835, h: 0.311, color: C.blue });
  copy(s, 'Lorem ipsum dolor amet dipiscing elit sed do eiusmod tempor incididunt.', {
    x: 6.755, y: 4.729, w: 3.024, h: 0.502, color: C.ink,
  });
}

function slide05(pptx) {
  const s = page(pptx, C.blue, C.white);
  title(s, 'Deep Dive into Your New Position', { x: 0.497, y: 0.344, w: 6.393, h: 1.591, color: C.white });
  waveDisc(s, 7.604, -0.725, 3.164, C.cream, C.blue);
  [0.508, 2.997].forEach((x, i) => {
    numberedColumn(s, x, Object.assign({ body: STEP_BODY, bw: 2.231 }, STEPS[i]));
  });
  photo(s, 6.266, 2.812, 3.164, 2.362, 'rect');
}

function slide06(pptx) {
  const s = page(pptx, C.red, C.white);
  title(s, 'Glimpse into Our Organization', { x: 0.466, y: 0.373, w: 5.945, h: 1.591, color: C.white });
  waveMark(s, 8.303, 0, 1.854, 1.719, C.cream);

  [{ head: 'Program One', y: 3.279 }, { head: 'Program Two', y: 4.332 }].forEach((row) => {
    s.addShape('line', {
      x: 0.57, y: row.y + 0.151, w: 0.334, h: 0,
      line: { color: C.white, width: 1, endArrowType: 'stealth' },
    });
    label(s, row.head, { x: 1.011, y: row.y, w: 1.499, h: 0.303, color: C.white });
    copy(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor.', {
      x: 1.011, y: row.y + 0.359, w: 3.237, h: 0.502, color: C.white,
    });
  });

  yearBlock(
    s, 3.458, 'Understanding Our Business',
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua enim ad minim.',
    { x: 4.13, y: 4.747, w: 5.391 }
  );
}

function slide07(pptx) {
  const s = page(pptx, C.blue, C.white);
  title(s, 'Understanding Our Culture', { x: 3.981, y: 0.373, w: 5.57, h: 1.575, color: C.white, align: 'right' });
  waveMark(s, -0.168, -0.011, 1.633, 1.515, C.cream);

  [{ head: 'Culture One', y: 2.812 }, { head: 'Culture Two', y: 4.068 }].forEach((row) => {
    label(s, row.head, { x: 2.891, y: row.y, w: 1.688, h: 0.303, color: C.white });
    copy(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.', {
      x: 2.891, y: row.y + 0.342, w: 2.269, h: 0.742, color: C.white,
    });
    photo(s, 0.57, row.y, 1.992, 1.093, 'rect');
  });

  text(s, '80%', { x: 7.438, y: 3.869, w: 2.109, h: 0.909, fontFace: F.head, fontSize: 50, color: C.white, align: 'right' });
  copy(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.', {
    x: 6.171, y: 4.747, w: 3.351, h: 0.535, color: C.white, align: 'right',
  });
}

function slide08(pptx) {
  const s = page(pptx, C.cream, C.blue);
  title(s, 'Meet Our CO Founder', { x: 0.497, y: 0.355, w: 5.108, h: 1.893, fontSize: 60, color: C.blue });
  waveDisc(s, 7.604, -0.725, 3.164, C.blue, C.cream);
  text(s, 'Alberto Maturazzi', { x: 2.652, y: 3.638, w: 2.953, h: 0.429, fontFace: F.head, fontSize: 21, color: C.blue });
  text(s, 'Best Founder of The Year', { x: 2.652, y: 4.067, w: 2.792, h: 0.303, fontFace: F.body, fontSize: 14, color: C.blue });
  copy(s, '\u201CLorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore.\u201D', {
    x: 2.652, y: 4.496, w: 3.008, h: 0.765, color: C.ink, italic: true,
  });
  text(s, '10+Years', { x: 6.386, y: 4.115, w: 3.162, h: 0.833, fontFace: F.head, fontSize: 45, color: C.blue, align: 'right' });
  label(s, 'Work Experience', { x: 6.49, y: 4.948, w: 3.008, h: 0.303, color: C.black, align: 'right' });
  photo(s, 0.57, 2.438, 1.956, 2.737, 'arch');
}

function slide09(pptx) {
  const s = page(pptx, C.cream, C.blue);
  title(s, 'Meet Our Expert Team', { x: 0.432, y: 0.316, w: 5.108, h: 1.893, fontSize: 60, color: C.blue });
  waveDisc(s, -0.874, 3.255, 3.164, C.blue, C.cream);
  text(s, '10+Years', { x: 6.386, y: 0.247, w: 3.162, h: 0.833, fontFace: F.head, fontSize: 45, color: C.blue, align: 'right' });
  label(s, 'Work Experience', { x: 6.49, y: 1.08, w: 3.008, h: 0.303, color: C.black, align: 'right' });

  [{ name: 'Bryan Alvarez ', x: 5.0 }, { name: 'Karina Sofaya', x: 7.354 }].forEach((person) => {
    photo(s, person.x, 2.349, 2.076, 2.076, 'arch');
    s.addShape('rect', { x: person.x, y: 4.425, w: 2.076, h: 0.75, fill: { color: C.blue }, line: { type: 'none' } });
    text(s, person.name, {
      x: person.x + 0.138, y: 4.507, w: 1.799, h: 0.379,
      fontFace: F.head, fontSize: 18, color: C.cream, align: 'center',
    });
    text(s, 'Job Position', {
      x: person.x + 0.418, y: 4.839, w: 1.266, h: 0.252,
      fontFace: F.head, fontSize: 11, color: C.cream, align: 'center',
    });
  });
}

function slide10(pptx) {
  const s = page(pptx, C.blue, C.white);
  title(s, 'Navigating Your New Role', { x: 0.466, y: 0.373, w: 5.05, h: 1.575, color: C.white });
  [2.851, 5.341].forEach((x, i) => {
    numberedColumn(s, x, Object.assign({ body: STEP_BODY, bw: i ? 2.286 : 2.225 }, STEPS[i]));
  });
  waveMark(s, 8.613, 3.66, 1.633, 1.515, C.cream);
  photo(s, 0.57, 3.054, 1.979, 2.121, 'arch');
  photo(s, 6.883, 0.451, 3.117, 2.169, 'pillLeft');
}

function slide11(pptx) {
  const s = page(pptx, C.red, C.white);
  title(s, 'Our Gallery Activity Company', { x: 3.246, y: 3.836, w: 6.276, h: 1.575, color: C.white, align: 'right' });
  waveMark(s, -0.168, 0.451, 1.633, 1.515, C.cream);
  copy(s, 'Lorem ipsum dolor amet consectetur adipiscing elit do eiusmod tempor.', {
    x: 0.5, y: 2.375, w: 2.964, h: 0.502, color: C.white,
  });
  photo(s, 3.246, 0.451, 2.646, 2.646, 'ellipse');
  photo(s, 0.57, 3.433, 1.742, 1.742, 'ellipse');
  photo(s, 6.281, 0.451, 3.719, 2.646, 'pillLeft');
}

function slide12(pptx) {
  const s = page(pptx, C.blue, C.white);
  waveMark(s, -0.516, 0.692, 2.823, 2.618, C.cream);
  text(s, 'Break Slide', {
    x: 0.398, y: 3.956, w: 8.964, h: 1.643,
    fontFace: F.head, fontSize: 104, color: C.white, lineSpacingMultiple: 0.9,
  });
  yearBlock(
    s, 0.266, ' Introducing New Employees',
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua',
    { x: 4.703, y: 1.556, w: 4.818 }
  );
}

function slide13(pptx) {
  const s = page(pptx, C.cream, C.blue);
  // x/y/h = the column, px/py = its "70%" caption, yx = the year + note below.
  const bars = [
    { pct: '70%', year: '2024', x: 3.441, y: 2.366, h: 3.263, px: 3.518, py: 2.473, yx: 4.113, nx: 3.38, fill: C.blueLt },
    { pct: '65%', year: '2023', x: 5.264, y: 3.055, h: 2.574, px: 5.372, py: 3.165, yx: 5.992, nx: 5.259, fill: C.blueMd },
    { pct: '80%', year: '2025', x: 7.191, y: 1.661, h: 3.967, px: 7.316, py: 1.766, yx: 8.244, nx: 7.51, fill: C.blue },
  ];
  bars.forEach((bar) => {
    s.addShape('rect', { x: bar.x, y: bar.y, w: 2.239, h: bar.h, fill: { color: bar.fill }, line: { type: 'none' } });
    text(s, bar.pct, { x: bar.px, y: bar.py, w: 2.009, h: 0.833, fontFace: F.head, fontSize: 45, color: C.white });
    text(s, bar.year, {
      x: bar.yx, y: 4.53, w: 1.082, h: 0.48,
      fontFace: F.head, fontSize: 24, color: C.white, align: 'right',
    });
    copy(s, 'Lorem ipsum dolor sit conse adipiscing elit. ', {
      x: bar.nx, y: 4.986, w: 1.815, h: 0.502, color: C.white, align: 'right',
    });
  });

  title(s, 'Growth Project Onboarding', { x: 0.432, y: 0.316, w: 6.884, h: 1.575, color: C.blue });
  label(s, ' Completed Projects', { x: 0.432, y: 4.442, w: 3.019, h: 0.328, fontSize: 15, color: C.blue });
  copy(s, 'Lorem ipsum dolor amet consectetur adipiscing elit, sed do eiusmod.', {
    x: 0.497, y: 4.733, w: 2.883, h: 0.535, color: C.ink,
  });
}

function slide14(pptx) {
  const s = page(pptx, C.cream, C.blue);
  const bars = [
    { pct: '80%', w: 5.792, y: 3.135, fill: C.blue },
    { pct: '70%', w: 4.53, y: 4.347, fill: C.blueMd },
  ];
  bars.forEach((bar) => {
    s.addShape('custGeom', {
      x: 0, y: bar.y, w: bar.w, h: 0.828,
      points: pillPath(bar.w, 0.828, 'right'), fill: { color: bar.fill }, line: { type: 'none' },
    });
    text(s, bar.pct, {
      x: 0.483, y: bar.y + 0.126, w: 1.165, h: 0.581,
      fontFace: F.head, fontSize: 30, color: C.white,
    });
  });

  title(s, 'Increasing Employee Performance', { x: 0.455, y: 0.352, w: 7.624, h: 1.575, color: C.blue });
  [{ head: ' Progress One', y: 3.136 }, { head: ' Progress Two', y: 4.407 }].forEach((row) => {
    label(s, row.head, { x: 6.498, y: row.y, w: 3.019, h: 0.328, fontSize: 15, color: C.blue, align: 'right' });
    copy(s, 'Lorem ipsum dolor amet consectetur adipiscing elit, sed do eiusmod.', {
      x: 6.634, y: row.y + 0.292, w: 2.883, h: 0.535, color: C.ink, align: 'right',
    });
  });
  waveMark(s, 8.613, 0.456, 1.633, 1.515, C.blue);
}

function slide15(pptx) {
  const s = page(pptx, C.blue, C.white);
  figure(s, '40,5', { x: 0.469, y: 2.647, w: 4.113, h: 1.527 });
  figure(s, '59,5', { x: 5.558, y: 2.647, w: 3.974, h: 1.527, align: 'right' });
  s.addShape('rect', { x: 0.571, y: 4.126, w: 3.603, h: 0.646, fill: { color: C.blueMd }, line: { type: 'none' } });
  s.addShape('rect', { x: 4.173, y: 4.126, w: 5.257, h: 0.646, fill: { color: C.cream }, line: { type: 'none' } });
  title(s, 'Data Visual Comparison', { x: 0.469, y: 0.352, w: 5.854, h: 1.575, color: C.white });
  copy(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore.', {
    x: 5.534, y: 1.139, w: 3.974, h: 0.502, color: C.white, align: 'right',
  });
  text(s, '2025', { x: 6.673, y: 0.262, w: 2.862, h: 0.833, fontFace: F.head, fontSize: 45, color: C.white, align: 'right' });
  text(s, '2024', { x: 0.486, y: 4.894, w: 2.055, h: 0.379, fontFace: F.head, fontSize: 18, color: C.white });
  text(s, '2025', { x: 7.459, y: 4.904, w: 2.055, h: 0.379, fontFace: F.head, fontSize: 18, color: C.white, align: 'right' });
}

function slide16(pptx) {
  const s = page(pptx, C.blue, C.white);
  s.addShape('line', { x: 0, y: 3.832, w: 10, h: 0, line: { color: C.white, width: 1 } });

  const nodes = [
    { pct: '40%', head: 'Progress One', x: 0.488, y: 3.447, hw: 1.725 },
    { pct: '50%', head: 'Progress Two', x: 2.895, y: 3.435, hw: 1.622 },
    { pct: '60%', head: 'Progress Three', x: 5.302, y: 3.423, hw: 1.979 },
    { pct: '70%', head: 'Progress Four', x: 7.709, y: 3.411, hw: 1.721 },
  ];
  nodes.forEach((node) => {
    s.addShape('ellipse', {
      x: node.x + 0.082, y: node.y, w: 0.732, h: 0.732, fill: { color: C.cream }, line: { type: 'none' },
    });
    text(s, node.pct, {
      x: node.x + 0.082, y: node.y + 0.202, w: 0.732, h: 0.328,
      fontFace: F.head, fontSize: 15, color: C.blue, align: 'center',
    });
    label(s, node.head, { x: node.x, y: node.y + 0.978, w: node.hw, h: 0.311, color: C.white });
    copy(s, 'Lorem ipsum dolor sit amet consectetur adipiscing.', {
      x: node.x, y: node.y + 1.281, w: 2.208, h: 0.502, color: C.white,
    });
  });

  title(s, 'Employee Induction Schedule 2025', { x: 0.469, y: 0.352, w: 6.812, h: 1.575, color: C.white });
  waveMark(s, 8.022, 0.763, 2.208, 2.048, C.cream);
}

function slide17(pptx) {
  const s = page(pptx, C.red, C.white);
  title(s, 'Data Visual Comparison', { x: 0.947, y: 0.352, w: 8.105, h: 0.825, color: C.white, align: 'center' });
  copy(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea.', {
    x: 0.947, y: 1.243, w: 8.105, h: 0.502, color: C.white, align: 'center',
  });

  text(s, [
    { text: '20,2', options: { fontSize: 86 } },
    { text: '%', options: { fontSize: 60 } },
    { text: '  ', options: { fontSize: 86 } },
  ], { x: 0.469, y: 2.935, w: 4.113, h: 1.527, fontFace: F.num, color: C.white, align: 'right' });
  text(s, [
    { text: '47,2', options: { fontSize: 86 } },
    { text: '%', options: { fontSize: 60 } },
    { text: '  ', options: { fontSize: 86 } },
  ], { x: 5.418, y: 2.935, w: 3.974, h: 1.527, fontFace: F.num, color: C.white });

  s.addShape('line', { x: 5, y: 2.329, w: 0, h: 3.261, line: { color: C.white, width: 2 } });
  s.addShape('ellipse', { x: 4.944, y: 4.462, w: 0.113, h: 0.113, fill: { color: C.white }, line: { type: 'none' } });

  const footer = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore';
  copy(s, footer, { x: 0.469, y: 4.731, w: 4.113, h: 0.502, color: C.white, align: 'right' });
  copy(s, footer, { x: 5.418, y: 4.731, w: 4.113, h: 0.502, color: C.white });
}

function slide18(pptx) {
  const s = page(pptx, C.blue, C.white);
  waveMark(s, 3.191, 2.556, 2.823, 2.618, C.cream);

  // Laptop mockup: lid shell, screen, then the two base slabs that run off-slide.
  s.addShape('roundRect', { x: 5.616, y: 1.79, w: 5.305, h: 3.596, rectRadius: 0.143, fill: { color: C.lid }, line: { type: 'none' } });
  photo(s, 5.736, 1.997, 5.068, 3.177, 'rect');
  s.addShape('rect', { x: 5.014, y: 5.375, w: 6.5, h: 0.042, fill: { color: C.deck }, line: { type: 'none' } });
  s.addShape('custGeom', {
    x: 5.111, y: 5.417, w: 6.4, h: 0.11,
    points: [{ x: 0, y: 0, moveTo: true }, { x: 6.4, y: 0 }, { x: 6.4, y: 0.11 }, { x: 0.25, y: 0.11 }, { close: true }],
    fill: { color: C.wedge }, line: { type: 'none' },
  });

  title(s, 'Device Laptop Mockup', { x: 0.469, y: 0.352, w: 4.809, h: 1.575, color: C.white });
  copy(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore.', {
    x: 5.719, y: 1.139, w: 4.072, h: 0.502, color: C.white,
  });
  text(s, '80%', { x: 0.469, y: 3.565, w: 1.912, h: 0.833, fontFace: F.body, fontSize: 45, color: C.white });
  label(s, 'Device Performance', { x: 0.499, y: 4.426, w: 2.094, h: 0.303, color: C.white });
  copy(s, 'Lorem ipsum dolor amet  consec adipiscing elit, sed do eiusmod.', {
    x: 0.499, y: 4.729, w: 2.579, h: 0.535, color: C.white,
  });
}

function slide19(pptx) {
  const s = page(pptx, C.blue, C.white);
  title(s, 'Get in Touch With Us', { x: 0.469, y: 0.305, w: 4.809, h: 3.995, fontSize: 86, color: C.white });
  waveMark(s, 7.414, 0.451, 2.823, 2.618, C.cream);
  ['(34) 4574 7136 8193', '(34) 4546 6135 5753'].forEach((line, i) => {
    label(s, line, { x: 7.276, y: 4.582 + i * 0.373, w: 2.237, h: 0.311, color: C.white, align: 'right' });
  });
  copy(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore', {
    x: 0.487, y: 4.731, w: 4.298, h: 0.502, color: C.white,
  });
}

function slide20(pptx) {
  const s = page(pptx, C.blue, C.white);
  waveDisc(s, 6.32, 1.934, 3.11, C.cream, C.blue);
  text(s, 'Thank You', {
    x: 0.477, y: 0.565, w: 6.568, h: 3.114,
    fontFace: F.head, fontSize: 125, color: C.white, lineSpacingMultiple: 0.7,
  });
  text(s, 'For Your Attention', { x: 0.477, y: 4.701, w: 4.273, h: 0.631, fontFace: F.head, fontSize: 33, color: C.white });
}

/* -------------------------------------------------------------------- main */

const DECK = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK_16x9', width: SIZE.w, height: SIZE.h });
  pptx.layout = 'DECK_16x9';
  pptx.author = 'pptxgenjs';
  pptx.title = 'Induction - Presentation Template';
  DECK.forEach((fn) => fn(pptx));
  return pptx;
}

build()
  .writeFile({ fileName: path.join(__dirname, '02418321-abc1-4b4d-9795-80fa732653ff_grok_final.pptx') })
  .then((f) => console.log('wrote', f));
