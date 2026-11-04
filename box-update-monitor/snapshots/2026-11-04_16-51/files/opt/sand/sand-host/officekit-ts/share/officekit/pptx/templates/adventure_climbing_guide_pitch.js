/**
 * "Summito" mountain-guide presentation template - 20 slides, 16:9 (13.333 x 7.5 in).
 * Rebuilt with pptxgenjs only. Photographs in the original deck are replaced by
 * flat grey placeholder shapes ("[image]"), everything else is native geometry.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const C = {
  dark: '283732',      // theme accent1 - deep forest green
  mid: '6C816E',       // accent2 - sage green
  sage: '8FA68A',      // accent3 - light sage
  grey: '949A9A',      // accent4
  navy: '44546A',      // theme dk2
  paleGrey: 'E7E6E6',  // theme lt2
  ink: '262626',       // text 85% black
  body: '808080',      // text 50% black
  white: 'FFFFFF',
  white85: 'D9D9D9',
  white75: 'BFBFBF',
  gold: 'FFCD15',
  photo: 'CCCCCC',     // stand-in for the deck's photographs
  photoDark: '767676',
  photoInk: 'A0A0A0',
  bullet: 'A3BF8F',
  slate: '547469',
  midPale: 'A6B4A7',
  sagePale: 'BCCAB9',
  greyPale: 'BFC2C2',
  greyDark: '6E7575',
};

const F = {
  xbold: 'Montserrat ExtraBold',
  semi: 'Montserrat SemiBold',
  med: 'Montserrat Medium',
  sans: 'Montserrat',
  body: 'Poppins',
  bodyMed: 'Poppins Medium',
  bodyBold: 'Poppins Bold',
};

// pptxgenjs mutates shadow option objects in place, so each shape needs its own copy.
const shadow = () => ({ type: 'outer', color: '000000', opacity: 0.2, blur: 15, offset: 3, angle: 45 });
const softShadow = () => ({ type: 'outer', color: '000000', opacity: 0.18, blur: 5, offset: 3, angle: 90 });

/* ------------------------------------------------------------------ helpers */

/** Text box. Reference boxes are top-anchored, pptxgenjs defaults to middle. */
function tx(slide, text, o) {
  slide.addText(text, Object.assign({ valign: 'top', fontFace: F.body, color: C.ink }, o));
}

/** Two-tone section heading: first run inherits `color`, second run is sage. */
function heading(slide, a, b, o) {
  tx(slide, [
    { text: a, options: { color: o.color || C.ink } },
    { text: b, options: { color: C.sage } },
  ], Object.assign({ fontFace: F.xbold, fontSize: 40, bold: true, color: C.ink }, o));
}

function rect(slide, o) {
  slide.addShape('rect', Object.assign({ line: { type: 'none' } }, o));
}

/** Polygon whose points are given as fractions of the bounding box. */
function poly(slide, o) {
  const opt = Object.assign({}, o);
  delete opt.pts;
  // custGeom path coordinates are relative to the shape's own top-left corner.
  opt.points = o.pts.map(([fx, fy]) => ({ x: fx * o.w, y: fy * o.h }));
  opt.points.push({ close: true });
  slide.addShape('custGeom', Object.assign({ line: { type: 'none' } }, opt));
}

/** Quadrilateral with vertical left/right edges optionally slanted. */
function quadPts(q) {
  return [[q.tl, 0], [q.tr, 0], [q.br, 1], [q.bl, 1]];
}

/** Grey stand-in for one of the deck's photographs. */
function photo(slide, o) {
  const pts = o.pts || quadPts({ tl: 0, tr: 1, br: 1, bl: 0 });
  poly(slide, { x: o.x, y: o.y, w: o.w, h: o.h, pts, fill: { color: C.photo } });
  if (o.label !== false) {
    tx(slide, '[image]', {
      x: o.x, y: o.y + o.h / 2 - 0.2, w: o.w, h: 0.4,
      align: 'center', valign: 'middle', fontSize: 11, color: C.photoInk,
    });
  }
}

function mixHex(a, b, f) {
  const ch = (h, i) => parseInt(h.substr(i * 2, 2), 16);
  return [0, 1, 2]
    .map((i) => Math.round(ch(a, i) + (ch(b, i) - ch(a, i)) * f).toString(16).padStart(2, '0'))
    .join('').toUpperCase();
}

/**
 * The reference deck fades a photo behind a dark-green overlay. pptxgenjs cannot emit
 * gradient fills, so the ramp is drawn as a row of solid slices whose colours are
 * interpolated between `from` and `to` (left edge -> right edge of the panel).
 */
function gradPanel(slide, o) {
  const steps = o.steps || 40;
  const q = o.quad || { tl: 0, tr: 1, br: 1, bl: 0 };
  const lerp = (a, b, f) => a + (b - a) * f;
  for (let i = 0; i < steps; i++) {
    const f0 = i / steps;
    const f1 = Math.min(1, (i + 1) / steps + 0.6 / steps);
    const ax = [lerp(q.tl, q.tr, f0), lerp(q.tl, q.tr, f1), lerp(q.bl, q.br, f1), lerp(q.bl, q.br, f0)]
      .map((f) => o.x + f * o.w);
    const bx = Math.min.apply(null, ax);
    const bw = Math.max.apply(null, ax) - bx;
    poly(slide, {
      x: bx, y: o.y, w: bw, h: o.h,
      pts: [[(ax[0] - bx) / bw, 0], [(ax[1] - bx) / bw, 0],
            [(ax[2] - bx) / bw, 1], [(ax[3] - bx) / bw, 1]],
      fill: { color: mixHex(o.from, o.to, (i + 0.5) / steps) },
    });
  }
}

function line(slide, o) {
  slide.addShape('line', {
    x: o.x, y: o.y, w: o.w || 0, h: o.h || 0,
    line: { color: o.color, width: o.width || 1, dashType: o.dash || 'solid',
      beginArrowType: o.arrowLeft ? 'triangle' : undefined,
      endArrowType: o.arrow ? 'triangle' : undefined },
  });
}

/** Three-segment elbow arrow: out, across, in. */
function elbowArrow(slide, x1, y1, x2, y2, color) {
  const mx = (x1 + x2) / 2;
  line(slide, { x: x1, y: y1, w: mx - x1, color });
  line(slide, { x: mx, y: Math.min(y1, y2), h: Math.abs(y2 - y1), color });
  line(slide, { x: mx, y: y2, w: x2 - mx, color, arrow: true });
}

/**
 * Small pictogram; `kind` picks one of the deck's icon motifs. `color` paints the
 * glyph, `bg` paints the punched-out details so they read against the tile behind.
 */
function icon(slide, kind, x, y, s, color, bg) {
  const c = color || C.white;
  const k = bg || C.dark;
  const R = (dx, dy, dw, dh, extra) => rect(slide, Object.assign({
    x: x + dx * s, y: y + dy * s, w: dw * s, h: dh * s, fill: { color: c } }, extra));
  const E = (dx, dy, dw, dh, extra) => slide.addShape('ellipse', Object.assign({
    x: x + dx * s, y: y + dy * s, w: dw * s, h: dh * s, fill: { color: c },
    line: { type: 'none' } }, extra));
  const cut = (dx, dy, dw, dh) => rect(slide, {
    x: x + dx * s, y: y + dy * s, w: dw * s, h: dh * s, fill: { color: k } });
  switch (kind) {
    case 'bars':
      R(0.10, 0.45, 0.16, 0.45); R(0.34, 0.20, 0.16, 0.70);
      R(0.58, 0.35, 0.16, 0.55); R(0.82, 0.10, 0.12, 0.80); break;
    case 'pie':
      E(0.05, 0.05, 0.9, 0.9); cut(0.52, 0.03, 0.45, 0.45); break;
    case 'trophy':
      slide.addShape('flowChartManualOperation', { x: x + 0.24 * s, y: y + 0.03 * s,
        w: 0.52 * s, h: 0.46 * s, fill: { color: c }, line: { type: 'none' } });
      slide.addShape('ellipse', { x: x - 0.02 * s, y: y + 0.04 * s, w: 0.3 * s, h: 0.3 * s,
        fill: { type: 'none' }, line: { color: c, width: 1.4 } });
      slide.addShape('ellipse', { x: x + 0.72 * s, y: y + 0.04 * s, w: 0.3 * s, h: 0.3 * s,
        fill: { type: 'none' }, line: { color: c, width: 1.4 } });
      R(0.44, 0.47, 0.12, 0.26); R(0.26, 0.73, 0.48, 0.13); break;
    case 'people':
      E(0.02, 0.16, 0.24, 0.24); E(0.74, 0.16, 0.24, 0.24);
      slide.addShape('flowChartManualOperation', { x, y: y + 0.42 * s, w: 0.32 * s, h: 0.34 * s,
        fill: { color: c }, line: { type: 'none' }, rotate: 180 });
      slide.addShape('flowChartManualOperation', { x: x + 0.68 * s, y: y + 0.42 * s,
        w: 0.32 * s, h: 0.34 * s, fill: { color: c }, line: { type: 'none' }, rotate: 180 });
      E(0.36, 0.02, 0.28, 0.28);
      slide.addShape('flowChartManualOperation', { x: x + 0.22 * s, y: y + 0.32 * s,
        w: 0.56 * s, h: 0.42 * s, fill: { color: c }, line: { type: 'none' }, rotate: 180 });
      break;
    case 'building':
      R(0.18, 0.02, 0.64, 0.96);
      [0.16, 0.34, 0.52].forEach((wy) => { cut(0.29, wy, 0.15, 0.11); cut(0.56, wy, 0.15, 0.11); });
      cut(0.42, 0.70, 0.16, 0.28);
      break;
    case 'gear':
      E(0.05, 0.05, 0.90, 0.90);
      slide.addShape('ellipse', { x: x + 0.34 * s, y: y + 0.34 * s, w: 0.32 * s, h: 0.32 * s,
        fill: { color: k }, line: { type: 'none' } });
      break;
    case 'target':
      slide.addShape('ellipse', { x: x + 0.02 * s, y: y + 0.02 * s, w: 0.96 * s, h: 0.96 * s,
        fill: { type: 'none' }, line: { color: c, width: 1.5 } });
      E(0.34, 0.34, 0.32, 0.32); break;
    case 'head':
      E(0.10, 0.05, 0.66, 0.72); R(0.18, 0.62, 0.5, 0.33);
      slide.addShape('ellipse', { x: x + 0.62 * s, y: y + 0.10 * s, w: 0.3 * s, h: 0.3 * s,
        fill: { type: 'none' }, line: { color: c, width: 1.2 } });
      break;
    case 'hiker':
      E(0.42, 0.03, 0.22, 0.22); R(0.34, 0.28, 0.22, 0.36);
      R(0.20, 0.62, 0.18, 0.32); R(0.54, 0.60, 0.16, 0.34); R(0.66, 0.10, 0.06, 0.80); break;
    case 'cloud':
      E(0.06, 0.34, 0.42, 0.42); E(0.30, 0.18, 0.50, 0.50); E(0.52, 0.38, 0.42, 0.38);
      R(0.10, 0.58, 0.80, 0.20); break;
    case 'truck':
      R(0.04, 0.28, 0.56, 0.40); R(0.60, 0.42, 0.32, 0.26);
      E(0.12, 0.62, 0.24, 0.24); E(0.62, 0.62, 0.24, 0.24); break;
    case 'link':
      slide.addShape('roundRect', { x: x + 0.02 * s, y: y + 0.30 * s, w: 0.52 * s, h: 0.34 * s,
        fill: { type: 'none' }, line: { color: c, width: 1.6 }, rectRadius: 0.05 * s });
      slide.addShape('roundRect', { x: x + 0.44 * s, y: y + 0.40 * s, w: 0.52 * s, h: 0.34 * s,
        fill: { type: 'none' }, line: { color: c, width: 1.6 }, rectRadius: 0.05 * s });
      break;
    case 'clipboard':
      slide.addShape('roundRect', { x: x + 0.02 * s, y: y + 0.09 * s, w: 0.42 * s, h: 0.62 * s,
        fill: { color: c }, line: { type: 'none' }, rectRadius: 0.05 * s });
      R(0.14, 0.01, 0.18, 0.12);
      E(0.18, 0.00, 0.10, 0.10, { fill: { color: k } });
      cut(0.32, 0.26, 0.06, 0.48);
      poly(slide, { x: x + 0.36 * s, y: y + 0.29 * s, w: 0.6 * s, h: 0.71 * s,
        pts: [[0, 0], [0.68, 0], [1, 0.26], [1, 1], [0, 1]], fill: { color: c } });
      cut(0.76, 0.29, 0.20, 0.03);
      break;
    case 'reply':
      slide.addShape('leftArrow', { x: x + 0.05 * s, y: y + 0.28 * s, w: 0.9 * s, h: 0.44 * s,
        fill: { color: c }, line: { type: 'none' } }); break;
    case 'rocket':
      slide.addShape('teardrop', { x: x + 0.22 * s, y: y + 0.04 * s, w: 0.74 * s, h: 0.74 * s,
        fill: { color: c }, line: { type: 'none' } });
      E(0.60, 0.19, 0.19, 0.19, { fill: { color: k } });
      poly(slide, { x: x + 0.02 * s, y: y + 0.26 * s, w: 0.34 * s, h: 0.24 * s,
        pts: [[0, 1], [0.35, 0], [1, 0], [1, 1]], fill: { color: c } });
      poly(slide, { x: x + 0.20 * s, y: y + 0.62 * s, w: 0.24 * s, h: 0.34 * s,
        pts: [[1, 0], [1, 1], [0, 0.65]], fill: { color: c } });
      break;
    case 'podcast':
      slide.addShape('ellipse', { x: x + 0.05 * s, y: y + 0.05 * s, w: 0.9 * s, h: 0.9 * s,
        fill: { type: 'none' }, line: { color: c, width: 1.6 } });
      E(0.36, 0.36, 0.28, 0.28); break;
    case 'tools':
      // screwdriver: blade + shaft running top-left to bottom-right
      R(0.05, 0.02, 0.15, 0.2, { rotate: 45 });
      R(0.26, 0.26, 0.12, 0.62, { rotate: 45 });
      E(0.10, 0.76, 0.18, 0.18);
      // wrench: open jaw + shaft running top-right to bottom-left
      slide.addShape('ellipse', { x: x + 0.56 * s, y: y + 0.02 * s, w: 0.36 * s, h: 0.36 * s,
        fill: { type: 'none' }, line: { color: c, width: 3 } });
      cut(0.55, 0.00, 0.20, 0.12);
      R(0.42, 0.28, 0.16, 0.62, { rotate: 315 });
      E(0.70, 0.76, 0.18, 0.18);
      break;
    case 'signpost':
      E(0.44, 0.00, 0.12, 0.12); R(0.46, 0.06, 0.08, 0.84);
      slide.addShape('pentagon', { x: x + 0.08 * s, y: y + 0.15 * s, w: 0.72 * s, h: 0.19 * s,
        fill: { color: c }, line: { type: 'none' }, flipH: true });
      slide.addShape('pentagon', { x: x + 0.20 * s, y: y + 0.43 * s, w: 0.72 * s, h: 0.19 * s,
        fill: { color: c }, line: { type: 'none' } });
      R(0.26, 0.90, 0.48, 0.10);
      break;
    case 'briefcase':
      R(0.04, 0.26, 0.92, 0.60);
      slide.addShape('rect', { x: x + 0.30 * s, y: y + 0.04 * s, w: 0.4 * s, h: 0.24 * s,
        fill: { type: 'none' }, line: { color: c, width: 1.2 } });
      cut(0.42, 0.48, 0.16, 0.14);
      break;
    case 'binocular':
      slide.addShape('roundRect', { x: x + 0.02 * s, y: y + 0.24 * s, w: 0.38 * s, h: 0.62 * s,
        fill: { color: c }, line: { type: 'none' }, rectRadius: 0.09 * s });
      slide.addShape('roundRect', { x: x + 0.60 * s, y: y + 0.24 * s, w: 0.38 * s, h: 0.62 * s,
        fill: { color: c }, line: { type: 'none' }, rectRadius: 0.09 * s });
      R(0.36, 0.34, 0.28, 0.18); R(0.10, 0.04, 0.22, 0.24); R(0.68, 0.04, 0.22, 0.24);
      cut(0.10, 0.52, 0.22, 0.22); cut(0.68, 0.52, 0.22, 0.22); break;
    case 'chat':
      slide.addShape('ellipse', { x: x + 0.02 * s, y: y + 0.20 * s, w: 0.46 * s, h: 0.46 * s,
        fill: { type: 'none' }, line: { color: c, width: 2 } });
      slide.addShape('triangle', { x: x + 0.06 * s, y: y + 0.62 * s, w: 0.2 * s, h: 0.24 * s,
        fill: { color: c }, line: { type: 'none' }, rotate: 200 });
      E(0.28, 0.06, 0.66, 0.66);
      slide.addShape('triangle', { x: x + 0.40 * s, y: y + 0.62 * s, w: 0.24 * s, h: 0.3 * s,
        fill: { color: c }, line: { type: 'none' }, rotate: 200 });
      break;
    case 'pin':
      slide.addShape('ellipse', { x: x + 0.12 * s, y: y + 0.02 * s, w: 0.76 * s, h: 0.76 * s,
        fill: { type: 'none' }, line: { color: c, width: 1.6 } });
      E(0.40, 0.28, 0.20, 0.20);
      slide.addShape('triangle', { x: x + 0.34 * s, y: y + 0.55 * s, w: 0.32 * s, h: 0.42 * s,
        fill: { color: c }, line: { type: 'none' }, rotate: 180 });
      break;
    case 'hash':
      R(0.24, 0.02, 0.10, 0.96, { rotate: 12 }); R(0.62, 0.02, 0.10, 0.96, { rotate: 12 });
      R(0.02, 0.28, 0.96, 0.10); R(0.02, 0.62, 0.96, 0.10); break;
    case 'globe':
      slide.addShape('ellipse', { x: x + 0.02 * s, y: y + 0.02 * s, w: 0.96 * s, h: 0.96 * s,
        fill: { type: 'none' }, line: { color: c, width: 1.5 } });
      slide.addShape('ellipse', { x: x + 0.32 * s, y: y + 0.02 * s, w: 0.36 * s, h: 0.96 * s,
        fill: { type: 'none' }, line: { color: c, width: 1.2 } });
      rect(slide, { x: x + 0.02 * s, y: y + 0.47 * s, w: 0.96 * s, h: 0.05 * s, fill: { color: c } });
      break;
    case 'mail':
      slide.addShape('rect', { x: x + 0.02 * s, y: y + 0.18 * s, w: 0.96 * s, h: 0.64 * s,
        fill: { type: 'none' }, line: { color: c, width: 1.5 } });
      line(slide, { x: x + 0.02 * s, y: y + 0.18 * s, w: 0.48 * s, h: 0.34 * s, color: c, width: 1.5 });
      line(slide, { x: x + 0.50 * s, y: y + 0.52 * s, w: 0.48 * s, h: -0.34 * s, color: c, width: 1.5 });
      break;
    case 'clock':
      slide.addShape('ellipse', { x: x + 0.02 * s, y: y + 0.02 * s, w: 0.96 * s, h: 0.96 * s,
        fill: { type: 'none' }, line: { color: c, width: 1.2 } });
      rect(slide, { x: x + 0.46 * s, y: y + 0.22 * s, w: 0.08 * s, h: 0.32 * s, fill: { color: c } });
      break;
    case 'level':
      R(0.06, 0.62, 0.18, 0.32); R(0.40, 0.38, 0.18, 0.56); R(0.74, 0.14, 0.18, 0.80); break;
    default:
      E(0.15, 0.15, 0.7, 0.7);
  }
}

/** Brand bar (logo + wordmark), language switch and page number - on every slide. */
function chrome(slide, n, opt) {
  const o = opt || {};
  const ink = o.dark ? C.white : C.dark;
  slide.addShape('triangle', { x: 0.687, y: 0.30, w: 0.21, h: 0.17,
    fill: { type: 'none' }, line: { color: ink, width: 1.5 } });
  slide.addShape('ellipse', { x: 0.845, y: 0.262, w: 0.075, h: 0.075,
    fill: { type: 'none' }, line: { color: ink, width: 1.2 } });
  tx(slide, 'Summito', { x: 0.938, y: 0.252, w: 1.358, h: 0.337,
    fontFace: F.semi, fontSize: 14, color: ink });

  slide.addShape('ellipse', { x: 11.912, y: 0.345, w: 0.135, h: 0.135,
    fill: { type: 'none' }, line: { color: ink, width: 1.2 } });
  line(slide, { x: 12.03, y: 0.462, w: 0.06, h: 0.06, color: ink, width: 1.2 });
  line(slide, { x: 12.286, y: 0.299, h: 0.248, color: ink, width: 1.25 });
  tx(slide, 'EN', { x: 12.343, y: 0.261, w: 0.517, h: 0.337,
    fontFace: F.sans, fontSize: 14, color: ink });

  tx(slide, String(o.page || n).padStart(2, '0'), { x: 12.388, y: 6.902, w: 0.593, h: 0.337,
    fontFace: F.semi, fontSize: 14, color: o.pageColor || (o.dark ? C.white : C.ink) });
}

/** "INFOGRAPHIC SECTION" heading shared by slides 9-18. */
function infographicTitle(slide) {
  heading(slide, 'INFOGRAPHIC ', 'SECTION',
    { x: 3.003, y: 0.807, w: 7.328, h: 0.774, align: 'center' });
}

/* ------------------------------------------------------- shared copy strings */
const LOREM_LONG = 'Lorem ipsum dolor sit, consectetur adipisicing elit, sed do eiusmod tempor. ' +
  'incididunt labore dolore magna. Lorem ipsum dolor sit, cons maecenas.';
const LOREM_MED = 'Lorem ipsum dolor sit, consectetur adipisicing elit, sed do eiusmod tempor.';
const LOREM_SHORT = 'Lorem ipsum dolor sit, consectetur adipis';
const LOREM_CARD = 'Lorem ipsum dolor sit amet, consectetur';
const QUOTE = '\u2018\u2019Lorem ipsum dolor sit, consectetur adipis maecenas porttitor congue massa.\u2019\u2019';

/* ================================================================= slide 01 */
function slide01(pptx) {
  const s = pptx.addSlide();
  gradPanel(s, { x: 0, y: 0, w: 13.333, h: 7.5, from: C.dark, to: C.photo });
  rect(s, { x: 8.231, y: 1.0, w: 3.253, h: 6.5, fill: { color: C.photo } });
  tx(s, '[image]', { x: 8.231, y: 3.9, w: 3.253, h: 0.4,
    align: 'center', fontSize: 11, color: C.photoInk });

  chrome(s, 1, { dark: true });
  tx(s, 'SUMMITO', { x: 0.667, y: 1.0, w: 5.621, h: 1.313,
    fontFace: F.xbold, fontSize: 72, color: C.white });
  tx(s, 'Lorem ipsum dolor sit, consectetur adipisicing elit, sed do eius tempor incididunt ' +
        'labore dolore magna. Lorem dolor sit, amet consectetur adipisicing elit.',
    { x: 0.731, y: 2.299, w: 5.143, h: 0.908, fontSize: 11, color: C.white85, lineSpacingMultiple: 1.5 });

  s.addShape('rect', { x: 0.731, y: 3.546, w: 1.717, h: 0.475,
    fill: { type: 'none' }, line: { color: C.white, width: 1.5 } });
  tx(s, 'Learn More', { x: 0.731, y: 3.615, w: 1.717, h: 0.337,
    align: 'center', fontFace: F.semi, fontSize: 14, color: C.white });

  tx(s, 'Presentation Template 2025', { x: 0.667, y: 6.534, w: 3.844, h: 0.341,
    fontSize: 10, color: C.white85, charSpacing: 3.5, lineSpacingMultiple: 1.5 });
  tx(s, 'Mount Semeru', { x: 7.432, y: 4.943, w: 1.958, h: 0.337,
    fontFace: F.xbold, fontSize: 14, color: C.white });
  tx(s, 'East Java, Indonesia', { x: 7.432, y: 5.175, w: 1.893, h: 0.33,
    fontSize: 10, italic: true, color: C.white, lineSpacingMultiple: 1.5 });
  tx(s, 'Lorem dolor sit, amet cons adipisicing elit. Magna.',
    { x: 10.514, y: 2.688, w: 2.248, h: 0.63, fontSize: 11, color: C.white, lineSpacingMultiple: 1.5 });
}

/* ================================================================= slide 02 */
function slide02(pptx) {
  const s = pptx.addSlide();
  photo(s, { x: 0.667, y: 1.0, w: 12.042, h: 3.27,
    pts: [[0, 0], [0.765, 0], [0.818, 0.194], [0.818, 0], [1, 0], [1, 1], [0, 1]] });
  rect(s, { x: 0.667, y: 3.699, w: 12.042, h: 2.754, fill: { color: C.dark }, shadow: shadow() });
  chrome(s, 2);

  heading(s, 'YOUR TRUSTED ', 'CLIMBING GUIDE',
    { x: 6.976, y: 4.399, w: 5.413, h: 1.447, color: C.white });
  tx(s, 'Lorem ipsum dolor sit, consectetur adipisicing elit, sed do eiusmod tempor incididunt ' +
        'labore dolore magna. Lorem ipsum dolor sit, consectetur.',
    { x: 1.503, y: 4.637, w: 4.321, h: 0.908, fontSize: 11, color: C.white, lineSpacingMultiple: 1.5 });

  rect(s, { x: 2.226, y: 2.497, w: 2.874, h: 1.617, fill: { color: C.mid }, shadow: shadow() });
  tx(s, 'Expert mountain guides for all levels since', { x: 2.559, y: 2.756, w: 2.208, h: 0.678,
    fontSize: 12, color: C.white, lineSpacingMultiple: 1.5 });
  tx(s, '2010', { x: 2.559, y: 3.297, w: 2.208, h: 0.558,
    fontFace: F.bodyBold, fontSize: 20, color: C.white, lineSpacingMultiple: 1.5 });
}

/* ================================================================= slide 03 */
function slide03(pptx) {
  const s = pptx.addSlide();
  rect(s, { x: 0.667, y: 1.0, w: 6.0, h: 5.5, fill: { color: C.dark }, shadow: shadow() });
  photo(s, { x: 6.667, y: 1.0, w: 6.0, h: 2.75,
    pts: [[0, 0], [0.799, 0], [0.893, 0.205], [0.893, 0], [1, 0], [1, 1], [0, 1]] });
  photo(s, { x: 0.812, y: 3.871, w: 5.71, h: 2.475,
    pts: [[0, 0], [1, 0], [1, 1], [0.233, 1], [0.151, 0.812], [0.151, 1], [0, 1]] });
  chrome(s, 3);

  heading(s, 'WHY SUMMITO ', 'STANDS OUT',
    { x: 1.109, y: 1.309, w: 5.413, h: 1.447, color: C.white });
  tx(s, LOREM_MED, { x: 1.13, y: 2.874, w: 4.065, h: 0.63,
    fontSize: 11, color: C.white, lineSpacingMultiple: 1.5 });

  [['01.', C.ink, 4.02, 0.566], ['02.', C.mid, 4.807, 0.674], ['03.', C.ink, 5.595, 0.674]]
    .forEach(([num, col, y, w]) => {
      tx(s, num, { x: 7.426, y: y + 0.041, w, h: 0.501,
        fontFace: F.xbold, fontSize: 18, color: col, lineSpacingMultiple: 1.5 });
      tx(s, LOREM_MED, { x: 8.322, y, w: 3.92, h: 0.63,
        fontSize: 11, color: C.ink, lineSpacingMultiple: 1.5 });
    });
}

/* ================================================================= slide 04 */
function slide04(pptx) {
  const s = pptx.addSlide();
  photo(s, { x: 8.143, y: 1.012, w: 4.524, h: 5.446,
    pts: [[0, 0], [1, 0], [1, 1], [0, 1], [0, 0.264], [0.128, 0.264], [0, 0.157]] });
  chrome(s, 4);

  heading(s, 'SUMMITO\u2019S EXPERT ', 'CLIMBING SERVICES',
    { x: 0.952, y: 1.238, w: 6.315, h: 1.447 });
  tx(s, LOREM_LONG, { x: 0.98, y: 2.808, w: 5.988, h: 0.63,
    fontSize: 11, color: C.body, lineSpacingMultiple: 1.5 });

  const services = [
    { n: '01.', nCol: C.ink, title: 'Guided Expeditions ', col: 0, row: 0, nw: 0.614 },
    { n: '02.', nCol: C.dark, title: 'Personalized Training', col: 0, row: 1, nw: 0.614 },
    { n: '03.', nCol: C.dark, title: 'Equipment Rental', col: 1, row: 0, nw: 0.614 },
    { n: '04.', nCol: C.ink, title: 'Adventure Packages', col: 1, row: 1, nw: 0.757 },
  ];
  services.forEach((it) => {
    const tX = [1.535, 4.764][it.col];
    const nX = [0.954, 4.183][it.col];
    const yT = [4.143, 5.306][it.row];
    tx(s, it.n, { x: nX, y: yT + 0.178, w: it.nw, h: 0.501,
      fontFace: F.xbold, fontSize: 18, color: it.nCol, lineSpacingMultiple: 1.5 });
    tx(s, it.title, { x: tX, y: yT, w: 2.523, h: 0.368,
      fontFace: F.xbold, fontSize: 12, color: C.ink, lineSpacingMultiple: 1.5 });
    tx(s, LOREM_SHORT, { x: tX, y: yT + 0.326, w: 2.523, h: 0.63,
      fontSize: 11, color: C.body, lineSpacingMultiple: 1.5 });
  });
}

/* ================================================================= slide 05 */
function slide05(pptx) {
  const s = pptx.addSlide();
  const topQuad = { tl: 0.103, tr: 1, br: 1, bl: 0 };
  const botQuad = { tl: 0, tr: 1, br: 0.85, bl: 0 };
  gradPanel(s, { x: 5.234, y: 0.999, w: 7.433, h: 2.759, quad: topQuad, from: C.photo, to: C.dark });
  gradPanel(s, { x: 0.679, y: 3.758, w: 4.567, h: 2.743, quad: botQuad, from: C.photo, to: C.dark });
  chrome(s, 5);

  heading(s, 'MUST-VISIT CLIMBING ', 'DESTINATIONS',
    { x: 5.686, y: 4.159, w: 6.968, h: 1.447 });
  tx(s, 'Lorem ipsum dolor sit, consectetur adipisicing elit, sed do eiusmod tempor. incididunt ' +
        'labore dolore magna. Lorem ipsum dolor sit, consectetur',
    { x: 5.714, y: 5.729, w: 6.247, h: 0.63, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5 });

  tx(s, '40%', { x: 1.725, y: 1.454, w: 1.246, h: 0.768,
    fontFace: F.xbold, fontSize: 30, color: C.dark, lineSpacingMultiple: 1.5 });
  tx(s, 'Lorem ipsum dolor sit, amet consectet adipisicing elit',
    { x: 1.725, y: 2.205, w: 2.54, h: 0.63, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5 });

  const peak = (nameA, nameB, tx0, ty0, days, level) => {
    tx(s, [{ text: nameA }, { text: nameB }], { x: tx0, y: ty0, w: 1.958, h: 0.412,
      fontFace: F.xbold, fontSize: 14, color: C.white, lineSpacingMultiple: 1.5 });
    icon(s, 'clock', tx0 + 0.109, ty0 + 0.523, 0.18);
    tx(s, days, { x: tx0 + 0.365, y: ty0 + 0.415, w: 2.528, h: 0.352,
      fontSize: 11, color: C.white, lineSpacingMultiple: 1.5 });
    icon(s, 'level', tx0 + 0.109, ty0 + 0.813, 0.18);
    tx(s, level, { x: tx0 + 0.365, y: ty0 + 0.714, w: 2.528, h: 0.352,
      fontSize: 11, color: C.white, lineSpacingMultiple: 1.5 });
  };
  peak('Mount ', 'Rinjani', 6.077, 2.428, '5 Days', 'Moderate to challenging');
  peak('Mount ', 'Sindoro', 1.113, 5.202, '2 Days', 'Moderate');
}

/* ================================================================= slide 06 */
function slide06(pptx) {
  const s = pptx.addSlide();
  chrome(s, 6);
  heading(s, 'CHOOSE YOUR ', 'CLIMBING EXPERIENCE', { x: 0.667, y: 1.0, w: 7.129, h: 1.447 });
  tx(s, 'Lorem ipsum dolor sit, consectetur adipisicing elit, sed do eiusmod tempor. incididunt ' +
        'labore dolore magna. Lorem ipsum dolor sit, cons maecenas porttitor congue massa. ' +
        'Fusce posuere, magna sed pulvinar',
    { x: 0.694, y: 2.57, w: 6.247, h: 0.908, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5 });
  photo(s, { x: 8.668, y: 1.0, w: 4.666, h: 2.75,
    pts: [[0, 0], [1, 0], [1, 1], [0.457, 1], [0.324, 0.775], [0.324, 1], [0, 1]] });

  poly(s, { x: 0.681, y: 3.953, w: 6.386, h: 2.501,
    pts: quadPts({ tl: 0, tr: 0.85, br: 1, bl: 0 }), fill: { color: C.dark }, shadow: shadow() });
  poly(s, { x: 6.299, y: 3.958, w: 6.386, h: 2.501,
    pts: quadPts({ tl: 0, tr: 1, br: 1, bl: 0.15 }), fill: { color: C.mid }, shadow: shadow() });

  const pkg = (x0, title, price, bullets, bulletColor) => {
    tx(s, title, { x: x0, y: 4.101, w: 2.231, h: 0.368,
      fontFace: F.xbold, fontSize: 12, color: C.white, lineSpacingMultiple: 1.5 });
    tx(s, price, { x: x0, y: 4.389, w: 2.528, h: 0.457,
      fontFace: F.xbold, fontSize: 16, color: C.white, lineSpacingMultiple: 1.5 });
    bullets.forEach((b, i) => {
      const bx = x0 + 0.079 + (i > 1 ? 2.767 : 0);
      const by = i % 2 === 0 ? 5.121 : 5.712;
      s.addShape('parallelogram', { x: bx, y: by, w: 0.141, h: 0.106,
        fill: { color: bulletColor }, line: { type: 'none' } });
      tx(s, b, { x: bx + 0.243, y: by - 0.153, w: 1.852, h: 0.627,
        fontSize: 11, color: C.white, lineSpacingMultiple: 1.5 });
    });
  };
  pkg(0.969, 'Beginner Package', '$500 per person', [
    'Easy trails with scenic views', 'Basic climbing techniques training',
    'Guided hikes with safety instructions', 'Ideal for families and casual adventurers',
  ], C.bullet);
  pkg(7.404, 'Intermediate Package', '$1200 per person', [
    'Moderate to challenging routes', 'Altitude training and acclimatization sessions',
    'Multi-day trekking with expert guides', 'Ideal for climbers leveling up',
  ], C.dark);
}

/* ================================================================= slide 07 */
function slide07(pptx) {
  const s = pptx.addSlide();
  photo(s, { x: 6.667, y: 3.905, w: 6.667, h: 3.595,
    pts: [[0, 0], [0.836, 0], [0.836, 0.175], [0.931, 0], [1, 0], [1, 1], [0, 1]] });
  chrome(s, 7, { pageColor: C.white });

  heading(s, 'CLIMB SAFELY ', 'WITH SUMMITO', { x: 6.827, y: 1.018, w: 5.132, h: 1.447 });
  tx(s, LOREM_LONG, { x: 6.854, y: 2.587, w: 4.451, h: 0.908,
    fontSize: 11, color: C.body, lineSpacingMultiple: 1.5 });

  const features = [
    { y: 1.351, fill: C.dark, glyph: 'hiker', title: 'Experienced & Certified Guides' },
    { y: 2.723, fill: C.mid, glyph: 'cloud', title: 'Real-Time Weather Monitoring' },
    { y: 4.088, fill: C.dark, glyph: 'truck', title: 'Emergency & Evacuation Plans' },
    { y: 5.454, fill: C.mid, glyph: 'link', title: 'High-Quality Safety Gear' },
  ];
  features.forEach((f, i) => {
    rect(s, { x: 0.667, y: f.y, w: 0.638, h: 0.638, fill: { color: f.fill }, shadow: shadow() });
    icon(s, f.glyph, 0.836, f.y + 0.169, 0.3, C.white, f.fill);
    tx(s, f.title, { x: 1.519, y: f.y - 0.183, w: 3.041, h: 0.368,
      fontFace: F.xbold, fontSize: 12, color: f.fill, lineSpacingMultiple: 1.5 });
    tx(s, 'Lorem ipsum dolor sit, consectetur adipisicing elit, sed do eiusmod tempor. incididunt labore.',
      { x: 1.519, y: f.y + 0.184, w: 3.922, h: 0.63, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5 });
    if (i < 3) {
      line(s, { x: 0.986, y: f.y + 0.777, h: 0.456, color: f.fill, width: 1.25, dash: 'dash' });
    }
  });
}

/* ================================================================= slide 08 */
function slide08(pptx) {
  const s = pptx.addSlide();
  photo(s, { x: 0.667, y: 3.724, w: 12.042, h: 2.153,
    pts: [[0, 0], [1, 0], [1, 1], [0.965, 1], [0.965, 0.727], [0.916, 1], [0, 1]] });
  rect(s, { x: 0.667, y: 1.0, w: 12.042, h: 2.75, fill: { color: C.dark }, shadow: shadow() });
  chrome(s, 8);

  heading(s, 'WHAT OUR ', 'CLIMBERS SAY',
    { x: 1.198, y: 1.639, w: 5.132, h: 1.447, color: C.white });
  tx(s, '90%', { x: 7.839, y: 1.297, w: 1.246, h: 0.768,
    fontFace: F.xbold, fontSize: 30, color: C.white, lineSpacingMultiple: 1.5 });
  tx(s, 'Lorem ipsum dolor sit, amet consectet adipisicing elit',
    { x: 7.839, y: 2.048, w: 2.54, h: 0.63, fontSize: 11, color: C.white, lineSpacingMultiple: 1.5 });

  const stars = (x0, y0) => {
    for (let i = 0; i < 5; i++) {
      s.addShape('star5', { x: x0 + i * 0.2855, y: y0, w: 0.27, h: 0.27,
        fill: { color: C.gold }, line: { type: 'none' } });
    }
  };
  const card = (x0, y0, fill, name, sub, starX, starY) => {
    rect(s, { x: x0, y: y0, w: 3.957, h: 1.953, fill: { color: fill }, shadow: shadow() });
    tx(s, name, { x: x0 + 0.226, y: y0 + 0.144, w: 2.523, h: 0.368,
      fontFace: F.xbold, fontSize: 12, color: C.white, lineSpacingMultiple: 1.5 });
    tx(s, sub, { x: x0 + 0.226, y: y0 + 0.391, w: 1.925, h: 0.33,
      fontSize: 10, italic: true, color: C.white85, lineSpacingMultiple: 1.5 });
    tx(s, QUOTE, { x: x0 + 0.226, y: y0 + 0.653, w: 3.397, h: 0.63,
      fontSize: 11, color: C.white, lineSpacingMultiple: 1.5 });
    stars(starX, starY);
  };
  card(7.613, 2.963, C.mid, 'Caroline Harlow', 'Mountain Adventure 01', 9.809, 4.392);
  card(1.553, 4.505, C.dark, 'Carl Johanson', 'Mountain Adventure 02', 3.765, 5.933);
}

/* ================================================================= slide 09 */
function slide09(pptx) {
  const s = pptx.addSlide();
  chrome(s, 9);
  infographicTitle(s);

  // Four big diamonds pinwheeled around a central circle.
  [[4.684, 3.324, C.mid], [5.843, 2.135, C.dark], [5.872, 4.483, C.dark], [7.032, 3.295, C.mid]]
    .forEach(([x, y, col]) => rect(s, { x, y, w: 1.618, h: 1.618, rotate: 44.3,
      fill: { color: col }, shadow: shadow() }));
  [[5.694, 3.75, 0.788, 0.767, 'E8EDE4'], [6.273, 3.155, 0.788, 0.767, 'E9EDE8'],
   [6.288, 4.308, 0.788, 0.808, 'E9EEE5'], [6.868, 3.714, 0.788, 0.808, 'E8EDE4']]
    .forEach(([x, y, w, h, col]) => rect(s, { x, y, w, h, rotate: 134.3, fill: { color: col } }));
  s.addShape('ellipse', { x: 6.18, y: 3.632, w: 0.973, h: 0.973,
    fill: { color: C.paleGrey }, line: { type: 'none' }, shadow: shadow() });

  icon(s, 'bars', 4.876, 3.99, 0.36, C.white, C.mid);
  icon(s, 'building', 6.505, 2.34, 0.34, C.white, C.dark);
  icon(s, 'people', 6.503, 5.578, 0.34, C.white, C.dark);
  icon(s, 'pie', 8.153, 3.932, 0.35, C.white, C.mid);
  icon(s, 'trophy', 6.483, 3.93, 0.37, C.dark, C.paleGrey);

  const projects = [
    { x: 1.2, y: 2.162, title: 'Project 02', col: C.mid },
    { x: 1.2, y: 4.672, title: 'Project 01', col: C.dark },
    { x: 9.446, y: 2.162, title: 'Project 03', col: C.mid },
    { x: 9.446, y: 4.672, title: 'Project 04', col: C.dark },
  ];
  projects.forEach((p) => {
    rect(s, { x: p.x, y: p.y, w: 2.688, h: 1.402, fill: { color: C.white }, shadow: shadow() });
    tx(s, p.title, { x: p.x + 0.16, y: p.y + 0.207, w: 1.331, h: 0.337,
      fontFace: F.sans, fontSize: 14, bold: true, color: p.col });
    tx(s, LOREM_CARD, { x: p.x + 0.16, y: p.y + 0.521, w: 2.368, h: 0.678,
      fontSize: 12, color: C.body, lineSpacingMultiple: 1.5 });
  });
}

/* ================================================================= slide 10 */
function slide10(pptx) {
  const s = pptx.addSlide();
  chrome(s, 10);
  infographicTitle(s);
  line(s, { x: 2.86, y: 3.984, w: 7.4, color: C.white85, width: 3 });

  const targets = [
    { x: 1.568, up: true, fill: C.dark, glyph: 'gear', label: 'Target 01', tX: 1.393, bX: 1.038 },
    { x: 3.741, up: false, fill: C.mid, glyph: 'building', label: 'Target 02', tX: 3.566, bX: 3.318 },
    { x: 5.915, up: true, fill: C.dark, glyph: 'head', label: 'Target 03', tX: 5.739, bX: 5.385 },
    { x: 8.088, up: false, fill: C.mid, glyph: 'bars', label: 'Target 04', tX: 7.912, bX: 7.558 },
    { x: 10.261, up: true, fill: C.dark, glyph: 'target', label: 'Target 05', tX: 10.085, bX: 9.731 },
  ];
  targets.forEach((t) => {
    const y = t.up ? 3.338 : 3.343;
    rect(s, { x: t.x, y, w: 1.292, h: 1.292, fill: { color: C.white }, shadow: shadow() });
    rect(s, { x: t.x + 0.124, y: y + 0.124, w: 1.045, h: 1.045, fill: { color: t.fill }, shadow: shadow() });
    icon(s, t.glyph, t.x + 0.43, y + 0.43, 0.44, C.white, t.fill);
    const ty = t.up ? 1.874 : 5.115;
    const by = t.up ? 2.199 : 5.44;
    tx(s, t.label, { x: t.tX, y: ty, w: 1.643, h: 0.337,
      align: 'center', fontFace: F.sans, fontSize: 14, bold: true, color: t.fill });
    tx(s, 'Lorem ipsum dolor sit amet, et ud consectetur tempor adipiscing',
      { x: t.bX, y: by, w: 2.352, h: 0.981, align: 'center', fontSize: 12,
        color: C.body, lineSpacingMultiple: 1.5 });
  });
}

/* ================================================================= slide 11 */
function slide11(pptx) {
  const s = pptx.addSlide();
  chrome(s, 11);
  infographicTitle(s);

  // Quartered circle with a diamond hub.
  const cx = 4.851, cy = 2.324, d = 3.632;
  [[180, 270, C.mid], [270, 360, C.dark], [90, 180, C.dark], [0, 90, C.mid]]
    .forEach(([a, b, col]) => s.addShape('pie', { x: cx, y: cy, w: d, h: d,
      angleRange: [a, b], fill: { color: col }, line: { type: 'none' }, shadow: shadow() }));
  line(s, { x: cx, y: cy + d / 2, w: d, color: C.white, width: 2 });
  line(s, { x: cx + d / 2, y: cy, h: d, color: C.white, width: 2 });
  s.addShape('diamond', { x: 5.909, y: 3.382, w: 1.516, h: 1.516,
    fill: { color: C.sage }, line: { type: 'none' }, shadow: shadow() });

  icon(s, 'chat', 5.56, 3.048, 0.36, C.white, C.mid);
  icon(s, 'pie', 7.409, 3.047, 0.35, C.white, C.dark);
  icon(s, 'binocular', 5.554, 4.895, 0.36, C.white, C.dark);
  icon(s, 'briefcase', 7.418, 4.891, 0.34, C.white, C.mid);
  icon(s, 'signpost', 6.45, 3.74, 0.62, C.white, C.sage);

  const cards = [
    { x: 1.37, y: 1.912, title: 'Data Analysis 01', col: C.mid, from: [5.456, 2.799], to: [4.549, 2.626] },
    { x: 8.784, y: 1.912, title: 'Data Analysis 02', col: C.dark, from: [7.846, 2.799], to: [8.753, 2.626] },
    { x: 8.784, y: 5.001, title: 'Data Analysis 03', col: C.mid, from: [7.846, 5.498], to: [8.753, 5.715] },
    { x: 1.37, y: 5.001, title: 'Data Analysis 04', col: C.dark, from: [5.456, 5.498], to: [4.549, 5.715] },
  ];
  cards.forEach((c) => {
    rect(s, { x: c.x, y: c.y, w: 3.179, h: 1.409, fill: { color: C.white }, shadow: shadow() });
    tx(s, c.title, { x: c.x + 0.478, y: c.y + 0.151, w: 2.224, h: 0.337,
      align: 'center', fontFace: F.sans, fontSize: 14, bold: true, color: c.col });
    tx(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. ',
      { x: c.x + 0.185, y: c.y + 0.488, w: 2.809, h: 0.678, align: 'center',
        fontSize: 12, color: C.body, lineSpacingMultiple: 1.5 });
    elbowArrow(s, c.from[0], c.from[1], c.to[0], c.to[1], c.col);
  });
}

/* ================================================================= slide 12 */
function slide12(pptx) {
  const s = pptx.addSlide();
  chrome(s, 12);
  infographicTitle(s);

  const stages = ['01', '03', '05', '02', '04', '06'];
  const fills = [C.dark, C.mid, C.dark, C.mid, C.dark, C.mid];
  stages.forEach((num, i) => {
    const x = [1.766, 5.099, 8.432][i % 3];
    const y = i < 3 ? 2.226 : 4.268;
    rect(s, { x, y, w: 3.135, h: 1.844, fill: { color: fills[i] } });
    tx(s, num, { x: x + 0.384, y: y + 0.206, w: 0.766, h: 0.589,
      fontFace: F.sans, fontSize: 32, bold: true, color: C.white, margin: [3.6, 3.6, 1.8, 1.8] });
    tx(s, 'PLANNING STAGE', { x: x + 0.384, y: y + 0.847, w: 2.369, h: 0.286,
      fontSize: 14, color: C.white, margin: [3.6, 3.6, 1.8, 1.8] });
    tx(s, 'Mauris quam dolor, cursus at porta et, luctus eget',
      { x: x + 0.384, y: y + 1.162, w: 2.369, h: 0.421,
        fontSize: 11, color: C.white, margin: [3.6, 3.6, 1.8, 1.8] });
  });
}

/* ================================================================= slide 13 */
function slide13(pptx) {
  const s = pptx.addSlide();
  chrome(s, 13, { page: 14 });
  infographicTitle(s);

  const steps = [
    { x: 1.13, y: 3.54, fill: C.dark, glyph: 'clipboard', label: 'Planning 01',
      lx: 0.545, ly: 1.955, bx: 0.686, by: 2.308 },
    { x: 3.546, y: 2.466, fill: C.mid, glyph: 'reply', label: 'Planning 02',
      lx: 2.961, ly: 5.164, bx: 3.102, by: 5.517 },
    { x: 5.801, y: 3.783, fill: C.dark, glyph: 'rocket', label: 'Planning 03',
      lx: 5.216, ly: 1.96, bx: 5.357, by: 2.313 },
    { x: 8.098, y: 2.709, fill: C.mid, glyph: 'podcast', label: 'Planning 04',
      lx: 7.512, ly: 5.161, bx: 7.653, by: 5.514 },
    { x: 10.436, y: 3.783, fill: C.dark, glyph: 'tools', label: 'Planning 05',
      lx: 9.85, ly: 1.96, bx: 9.991, by: 2.313 },
  ];
  steps.forEach((st) => {
    rect(s, { x: st.x, y: st.y, w: 1.768, h: 1.768, rotate: 315, fill: { color: st.fill } });
    rect(s, { x: st.x + 0.207, y: st.y + 0.204, w: 1.359, h: 1.359, rotate: 315,
      fill: { color: C.white } });
    icon(s, st.glyph, st.x + 0.53, st.y + 0.52, 0.72, st.fill, C.white);
    tx(s, st.label, { x: st.lx, y: st.ly, w: 2.939, h: 0.337,
      align: 'center', fontFace: F.sans, fontSize: 14, bold: true, color: C.ink });
    tx(s, 'Lorem ipsum dolor sit amet, elit, sed do eiusmod',
      { x: st.bx, y: st.by, w: 2.657, h: 0.707, align: 'center',
        fontSize: 12, color: C.body, lineSpacingMultiple: 1.5 });
  });
}

/* ================================================================= slide 14 */
function slide14(pptx) {
  const s = pptx.addSlide();
  chrome(s, 14);
  infographicTitle(s);
  line(s, { x: 6.629, y: 5.8, w: 5.381, color: C.white75, width: 1 });

  // Pentagon "arrow" bars, authored horizontally then rotated upright (as in the source).
  const bars = [
    { x: 5.723, y: 3.845, len: 3.21, fill: C.dark, lx: 7.026, pct: '80%', pctY: 2.197, m: 'Jan' },
    { x: 6.742, y: 4.064, len: 2.765, fill: C.mid, lx: 7.822, pct: '68%', pctY: 2.642, m: 'Feb' },
    { x: 7.391, y: 3.917, len: 3.06, fill: C.dark, lx: 8.619, pct: '74%', pctY: 2.347, m: 'Mar' },
    { x: 7.965, y: 3.695, len: 3.504, fill: C.mid, lx: 9.415, pct: '97%', pctY: 1.903, m: 'Apr' },
    { x: 9.343, y: 4.276, len: 2.341, fill: C.dark, lx: 10.211, pct: '52%', pctY: 3.067, m: 'May' },
    { x: 9.558, y: 3.695, len: 3.504, fill: C.mid, lx: 11.008, pct: '97%', pctY: 1.903, m: 'Jun' },
  ];
  bars.forEach((b) => {
    s.addShape('homePlate', { x: b.x, y: b.y, w: b.len, h: 0.707, rotate: 270,
      fill: { color: b.fill }, line: { type: 'none' } });
    tx(s, b.pct, { x: b.lx, y: b.pctY, w: 0.605, h: 0.375,
      align: 'center', fontSize: 12, color: C.body, lineSpacingMultiple: 1.5 });
    tx(s, b.m, { x: b.lx, y: 5.827, w: 0.605, h: 0.375,
      align: 'center', fontSize: 12, color: C.body, lineSpacingMultiple: 1.5 });
  });

  const months = [
    { m: 'January', x: 1.324, y: 2.49, col: C.dark }, { m: 'February', x: 3.764, y: 2.49, col: C.mid },
    { m: 'March', x: 1.324, y: 3.575, col: C.dark }, { m: 'April', x: 3.764, y: 3.575, col: C.mid },
    { m: 'May', x: 1.324, y: 4.661, col: C.dark }, { m: 'June', x: 3.764, y: 4.661, col: C.mid },
  ];
  months.forEach((it) => {
    tx(s, it.m, { x: it.x, y: it.y, w: 1.147, h: 0.337,
      fontFace: F.sans, fontSize: 14, bold: true, color: it.col });
    tx(s, 'Lorem ipsum dolor sit ut amet, elit. Sed',
      { x: it.x, y: it.y + 0.277, w: 2.153, h: 0.678,
        fontSize: 12, color: C.body, lineSpacingMultiple: 1.5 });
  });
}

/* ================================================================= slide 15 */
function slide15(pptx) {
  const s = pptx.addSlide();
  chrome(s, 15);
  infographicTitle(s);

  const rows = [1.837, 2.943, 4.049, 5.156];
  const leftNums = ['01', '03', '05', '07'];
  const rightNums = ['02', '04', '06', '08'];
  const alt = [C.dark, C.mid, C.dark, C.mid];

  rows.forEach((y, i) => {
    // left column
    s.addShape(i === 1 ? 'roundRect' : 'rect', { x: 1.077, y, w: 3.179, h: 0.948,
      fill: { color: C.white }, line: { type: 'none' }, rectRadius: 0.1, shadow: shadow() });
    rect(s, { x: 3.765, y: y + 0.083, w: 0.958, h: 0.782, fill: { color: alt[i] }, shadow: shadow() });
    tx(s, 'Lorem ipsum dolor sit, consectetur adipiscing', { x: 1.336, y: y + 0.121, w: 2.171,
      h: 0.678, align: 'right', fontSize: 12, color: C.body, lineSpacingMultiple: 1.5 });
    tx(s, leftNums[i], { x: 3.944, y: y + 0.255, w: 0.599, h: 0.438,
      align: 'center', fontFace: F.sans, fontSize: 20, bold: true, color: C.white });
    // right column (mirrored)
    rect(s, { x: 9.077, y, w: 3.179, h: 0.948, fill: { color: C.white }, shadow: shadow() });
    rect(s, { x: 8.611, y: y + 0.083, w: 0.958, h: 0.782,
      fill: { color: i % 2 === 0 ? C.mid : C.dark }, shadow: shadow() });
    tx(s, 'Lorem ipsum dolor sit, consectetur adipiscing', { x: 9.827, y: y + 0.121, w: 2.171,
      h: 0.678, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5 });
    tx(s, rightNums[i], { x: 8.79, y: y + 0.255, w: 0.599, h: 0.438,
      align: 'center', fontFace: F.sans, fontSize: 20, bold: true, color: C.white });
  });

  // Connectors between the number tiles and the hub. Rows 2 and 3 run straight in;
  // rows 1 and 4 step around the circle with an elbow.
  line(s, { x: 4.723, y: 2.311, h: 0.742, color: C.dark, width: 1 });
  line(s, { x: 4.723, y: 3.053, w: 1.027, color: C.dark, width: 1 });
  line(s, { x: 4.723, y: 3.417, w: 0.799, color: C.mid, width: 1, arrowLeft: true });
  line(s, { x: 4.723, y: 4.523, w: 0.799, color: C.dark, width: 1, arrowLeft: true });
  line(s, { x: 5.15, y: 4.883, w: 0.6, color: C.mid, width: 1 });
  line(s, { x: 5.15, y: 4.883, h: 0.746, color: C.mid, width: 1 });
  line(s, { x: 4.723, y: 5.629, w: 0.427, color: C.mid, width: 1, arrowLeft: true });
  line(s, { x: 7.583, y: 3.053, w: 0.634, color: C.mid, width: 1 });
  line(s, { x: 8.217, y: 2.311, h: 0.742, color: C.mid, width: 1 });
  line(s, { x: 8.217, y: 2.311, w: 0.394, color: C.mid, width: 1, arrow: true });
  line(s, { x: 7.811, y: 3.417, w: 0.8, color: C.dark, width: 1, arrow: true });
  line(s, { x: 7.811, y: 4.523, w: 0.8, color: C.mid, width: 1, arrow: true });
  line(s, { x: 7.583, y: 4.883, w: 1.0, color: C.dark, width: 1 });
  line(s, { x: 8.583, y: 4.883, h: 0.746, color: C.dark, width: 1 });
  line(s, { x: 8.583, y: 5.629, w: 0.4, color: C.dark, width: 1, arrow: true });

  s.addShape('ellipse', { x: 5.369, y: 2.673, w: 2.595, h: 2.595,
    fill: { color: C.white }, line: { type: 'none' }, shadow: shadow() });
  s.addShape('ellipse', { x: 5.509, y: 2.812, w: 2.316, h: 2.316,
    fill: { color: C.white }, line: { type: 'none' }, shadow: shadow() });
  icon(s, 'people', 6.07, 3.34, 1.0, C.dark);
  icon(s, 'gear', 6.42, 4.14, 0.42, C.dark, C.white);
}

/* ================================================================= slide 16 */
function slide16(pptx) {
  const s = pptx.addSlide();
  chrome(s, 16);

  /** Card header: product label plus the three-dot menu. */
  const cardHead = (label, lx, ly, dotsX, dotsY) => {
    tx(s, label, { x: lx, y: ly, w: 2.548, h: 0.337, fontFace: F.bodyMed, fontSize: 14, color: C.navy });
    for (let i = 0; i < 3; i++) {
      s.addShape('ellipse', { x: dotsX + i * 0.246, y: dotsY, w: 0.133, h: 0.133,
        fill: { color: C.mid }, line: { type: 'none' } });
    }
  };
  /** Smooth line-chart trace drawn with cubic segments (fractions of the box). */
  const wave = (x, y, w, h, pts, color) => {
    const P = pts.map(([fx, fy]) => [fx * w, fy * h]);
    const opts = [{ x: P[0][0], y: P[0][1], moveTo: true }];
    for (let i = 1; i + 2 < P.length; i += 3) {
      opts.push({ x: P[i + 2][0], y: P[i + 2][1],
        curve: { type: 'cubic', x1: P[i][0], y1: P[i][1], x2: P[i + 1][0], y2: P[i + 1][1] } });
    }
    s.addShape('custGeom', { x, y, w, h, points: opts,
      fill: { type: 'none' }, line: { color, width: 1.5 } });
  };
  const traceA = [[0, 0.81], [0.082, 0.391], [0.165, -0.029], [0.262, 0.002],
    [0.359, 0.032], [0.493, 0.916], [0.582, 0.992], [0.672, 1.068], [0.727, 0.576], [0.797, 0.456],
    [0.866, 0.336], [0.933, 0.305], [1, 0.274]];
  const traceB = [[0, 0.19], [0.065, 0.457], [0.13, 0.723], [0.205, 0.758],
    [0.28, 0.793], [0.383, 0.466], [0.451, 0.4], [0.518, 0.333], [0.544, 0.258], [0.61, 0.358],
    [0.675, 0.458], [0.789, 1.035], [0.844, 0.998], [0.899, 0.962], [0.914, 0.305], [0.941, 0.14],
    [0.967, -0.024], [0.983, -0.007], [1, 0.011]];
  const flip = (t) => t.map(([fx, fy]) => [1 - fx, 1 - fy]).reverse();
  /** Big money figure preceded by a coloured tick. */
  const figure = (text, tX, tY, tickX, tickY, color) => {
    rect(s, { x: tickX, y: tickY, w: 0.088, h: 0.417, fill: { color } });
    tx(s, text, { x: tX, y: tY, w: 2.366, h: 0.572, fontFace: F.med, fontSize: 28, color });
  };

  rect(s, { x: 1.002, y: 0.763, w: 5.318, h: 2.835, fill: { color: C.white }, shadow: shadow() });
  cardHead('Product 01', 1.29, 0.984, 5.369, 1.069);
  wave(1.558, 1.793, 4.377, 0.585, traceA, C.dark);
  wave(1.555, 1.464, 4.377, 1.156, traceB, C.mid);
  figure('$543,00', 1.657, 2.789, 1.555, 2.855, C.dark);
  figure('785,000', 4.051, 2.789, 3.949, 2.855, C.mid);

  rect(s, { x: 1.002, y: 3.775, w: 5.318, h: 2.835, fill: { color: C.white }, shadow: shadow() });
  cardHead('Product 02', 1.29, 3.996, 5.369, 4.081);
  wave(1.555, 4.717, 4.377, 0.585, flip(traceA), C.sage);
  wave(1.558, 4.476, 4.377, 1.156, flip(traceB), C.dark);
  figure('$543,00', 1.657, 5.801, 1.555, 5.868, C.sage);
  figure('785,000', 4.051, 5.801, 3.949, 5.868, C.dark);

  rect(s, { x: 6.904, y: 3.775, w: 4.85, h: 2.835, fill: { color: C.white }, shadow: shadow() });
  tx(s, 'Our Planning', { x: 7.125, y: 4.003, w: 2.548, h: 0.337,
    fontFace: F.bodyMed, fontSize: 14, color: C.navy });
  [[7.19, 4.582, 1.319, C.dark, 7.269, 4.707], [7.933, 5.303, 0.598, C.mid, 8.011, 5.401],
   [8.675, 4.861, 1.039, C.mid, 8.754, 4.998], [9.417, 5.12, 0.78, C.dark, 9.496, 5.169],
   [10.16, 4.668, 1.232, C.mid, 10.238, 4.707], [10.902, 5.056, 0.844, C.dark, 10.981, 5.098]]
    .forEach(([bx, by, bh, col, lx, ly]) => {
      rect(s, { x: bx, y: by, w: 0.577, h: bh, fill: { color: col } });
      tx(s, '75', { x: lx, y: ly, w: 0.419, h: 0.303, align: 'center', fontSize: 12, color: C.white });
    });
  [[7.247, 6.214, C.dark, 7.449, 6.064], [9.216, 6.205, C.mid, 9.417, 6.055]]
    .forEach(([dx, dy, col, lx, ly]) => {
      s.addShape('ellipse', { x: dx, y: dy, w: 0.157, h: 0.157,
        fill: { color: col }, line: { type: 'none' } });
      tx(s, 'Lorem ipsum dolor', { x: lx, y: ly, w: 1.767, h: 0.404,
        fontSize: 12, color: C.navy, lineSpacingMultiple: 1.5 });
    });

  tx(s, 'BUILDING KPI DASHBOARDS', { x: 6.838, y: 1.032, w: 5.493, h: 1.447,
    fontFace: F.xbold, fontSize: 40, color: C.ink });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut.',
    { x: 6.838, y: 2.622, w: 4.562, h: 0.707, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5 });
}

/* ================================================================= slide 17 */
function slide17(pptx) {
  const s = pptx.addSlide();
  chrome(s, 17);
  infographicTitle(s);

  const UP_A = [[3.027, 1.017], [3.458, 0.586], [3.212, 0.833], [3.715, 0.329], [3.146, 0.899]];
  const UP_B = [[3.715, 0.329], [3.59, 0.455], [2.959, 1.085], [3.715, 0.329], [3.458, 0.586]];
  const DOWN = [0.403, 0.751, 1.017, 0.833, 0.652];

  const cards = [
    { x: 0.76, color: C.dark, up: UP_A, axis: '3F3F3F', big: '$63,04', badge: '+132',
      legend: [['78,03%', C.dark, C.grey], ['32,23%', C.mid, C.dark]] },
    { x: 4.937, color: C.mid, up: UP_B, axis: C.greyDark, big: '$8943', badge: '849',
      legend: [['$392,120', C.dark, C.grey], ['$905,50', C.mid, C.mid]] },
    { x: 9.114, color: C.sage, up: UP_A, axis: C.greyDark, big: '834,34', badge: '930',
      legend: [['+45,4%', C.sage, C.grey], ['-854,3%', C.sage, C.sage]] },
  ];
  cards.forEach((c) => {
    rect(s, { x: c.x, y: 1.755, w: 3.424, h: 4.703, fill: { color: C.white }, shadow: shadow() });
    c.up.forEach(([top, h], i) => {
      const bx = c.x + 0.304 + i * 0.6;
      rect(s, { x: bx, y: top, w: 0.479, h, fill: { color: c.color } });
      rect(s, { x: bx, y: 4.044, w: 0.479, h: DOWN[i], fill: { color: C.grey } });
    });
    line(s, { x: c.x + 0.304, y: 4.044, w: 2.879, color: c.axis, width: 1.5 });
    tx(s, c.big, { x: c.x + 0.283, y: 2.141, w: 1.785, h: 0.64,
      fontFace: F.med, fontSize: 32, color: c.color });
    s.addShape('roundRect', { x: c.x + 2.611, y: 2.345, w: 0.6, h: 0.269,
      fill: { color: c.color }, line: { type: 'none' }, rectRadius: 0.134 });
    tx(s, c.badge, { x: c.x + 2.638, y: 2.345, w: 0.545, h: 0.269,
      align: 'center', fontSize: 10, color: C.white });
    c.legend.forEach(([label, textCol, dotCol], i) => {
      const y = i === 0 ? 5.372 : 5.801;
      s.addShape('ellipse', { x: c.x + 0.369, y: y + 0.044, w: 0.175, h: 0.175,
        fill: { color: dotCol }, line: { type: 'none' } });
      tx(s, label, { x: c.x + 0.628, y, w: 0.955, h: 0.303, fontSize: 12, color: textCol });
    });
    tx(s, 'Lorem ipsum sit amet, elit, sed', { x: c.x + 1.782, y: 5.36, w: 1.514, h: 0.707,
      fontSize: 12, color: C.body, lineSpacingMultiple: 1.5 });
  });
}

/* ================================================================= slide 18 */
function slide18(pptx) {
  const s = pptx.addSlide();
  chrome(s, 18);
  infographicTitle(s);

  const L = 1.688, W = 9.956, ROW = 0.412, TOP = 2.624;

  rect(s, { x: L, y: 1.817, w: W, h: 0.395, fill: { color: C.navy }, shadow: softShadow() });
  tx(s, 'March 2025', { x: 5.812, y: 1.846, w: 1.909, h: 0.337,
    align: 'center', fontFace: F.sans, fontSize: 14, bold: true, color: C.white });

  const weeks = [
    { x: 1.684, w: 2.496, head: C.dark, sub: C.slate, label: 'Week 01', lx: 2.373, day0: 1.744 },
    { x: 4.196, w: 2.483, head: C.mid, sub: C.midPale, label: 'Week 02', lx: 4.873, day0: 4.263 },
    { x: 6.696, w: 2.483, head: C.sage, sub: C.sagePale, label: 'Week 03', lx: 7.373, day0: 6.75 },
    { x: 9.196, w: 2.448, head: C.grey, sub: C.greyPale, label: 'Week 04', lx: 9.855, day0: 9.237 },
  ];
  rect(s, { x: L, y: TOP, w: W, h: ROW, fill: { color: C.white85 }, shadow: softShadow() });
  weeks.forEach((wk) => {
    rect(s, { x: wk.x, y: 2.223, w: wk.w, h: 0.386, fill: { color: wk.head }, shadow: softShadow() });
    rect(s, { x: wk.x, y: 2.625, w: wk.w, h: 0.408, fill: { color: wk.sub }, shadow: softShadow() });
    tx(s, wk.label, { x: wk.lx, y: 2.247, w: 1.129, h: 0.337,
      align: 'center', fontFace: F.sans, fontSize: 14, bold: true, color: C.white });
    for (let d = 0; d < 7; d++) {
      tx(s, String(d + 1), { x: wk.day0 + d * 0.356, y: 2.693, w: 0.259, h: 0.303,
        align: 'center', fontSize: 12, color: C.white });
    }
  });

  // Alternating task rows.
  for (let r = 0; r < 8; r++) {
    rect(s, { x: L, y: 3.036 + r * ROW, w: W, h: ROW,
      fill: { color: r % 2 === 0 ? C.paleGrey : C.white85 }, shadow: softShadow() });
  }

  // Scheduled task blocks.
  [[2.768, 3.046, 1.411, C.dark], [4.553, 3.456, 1.777, C.mid], [4.194, 3.864, 1.413, C.mid],
   [9.545, 3.866, 2.101, C.grey], [6.691, 4.689, 1.068, C.sage], [1.688, 5.107, 1.071, C.dark],
   [9.557, 5.521, 0.704, C.grey], [7.411, 5.927, 1.771, C.sage]]
    .forEach(([x, y, w, col]) => rect(s, { x, y, w, h: 0.401, fill: { color: col }, shadow: softShadow() }));

  // White grid.
  for (let r = 0; r <= 10; r++) {
    line(s, { x: 1.696, y: 2.211 + r * ROW, w: 9.947, color: C.white, width: 1 });
  }
  for (let d = 1; d < 28; d++) {
    const x = 1.688 + d * 0.357;
    const isWeekEdge = d % 7 === 0;
    line(s, { x, y: isWeekEdge ? 2.22 : TOP, h: isWeekEdge ? 4.114 : 3.711, color: C.white, width: 1 });
  }
}

/* ================================================================= slide 19 */
function slide19(pptx) {
  const s = pptx.addSlide();
  gradPanel(s, { x: 0.667, y: 1.0, w: 6.0, h: 5.458, from: C.dark, to: C.photo });
  // notch cut out of the panel's top edge
  poly(s, { x: 0.667, y: 1.0, w: 6.0, h: 5.458, fill: { color: C.white },
    pts: [[0.671, 0], [0.809, 0], [0.809, 0.152]] });
  photo(s, { x: 7.088, y: 4.069, w: 5.62, h: 2.389,
    pts: [[0, 0], [1, 0], [1, 1], [0.425, 1], [0.322, 0.757], [0.322, 1], [0, 1]] });
  chrome(s, 19);

  heading(s, 'STAY CONNECTED ', 'WITH SUMMITO', { x: 7.06, y: 1.285, w: 5.648, h: 1.447 });
  tx(s, 'Lorem ipsum dolor sit, consectetur adipisicing elit, sed do eiusmod tempor. incididunt ' +
        'labore dolore magna. Lorem ipsum dolor sit, cons maecenas. , consectetur adipisicing. labore dolore.',
    { x: 7.088, y: 2.854, w: 5.306, h: 0.908, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5 });

  const contacts = [
    { y: 1.655, glyph: 'pin', title: 'Location', sub: 'Your Business Location', subW: 2.796 },
    { y: 2.79, glyph: 'hash', title: 'Social Media', sub: '@yourbusinesssocial', subW: 2.18 },
    { y: 3.925, glyph: 'globe', title: 'Sites', sub: 'www.yourgreatsite.com', subW: 2.18 },
    { y: 5.06, glyph: 'mail', title: 'Email', sub: 'yourbusinessemail@gmail.com', subW: 2.882 },
  ];
  contacts.forEach((c) => {
    icon(s, c.glyph, 1.086, c.y + 0.259, 0.3, C.white);
    tx(s, c.title, { x: 1.548, y: c.y + 0.09, w: 2.108, h: 0.421,
      fontFace: F.sans, fontSize: 14, bold: true, color: C.white });
    tx(s, c.sub, { x: 1.559, y: c.y + 0.409, w: c.subW, h: 0.375,
      fontSize: 12, color: C.white, lineSpacingMultiple: 1.5 });
  });
}

/* ================================================================= slide 20 */
function slide20(pptx) {
  const s = pptx.addSlide();
  photo(s, { x: 0.667, y: 1.0, w: 4.022, h: 2.75, pts: quadPts({ tl: 0, tr: 0.904, br: 1, bl: 0 }) });
  poly(s, { x: 4.689, y: 3.758, w: 8.02, h: 2.75,
    pts: quadPts({ tl: 0, tr: 1, br: 1, bl: 0.056 }), fill: { color: C.photoDark } });
  tx(s, '[image]', { x: 4.689, y: 4.95, w: 8.02, h: 0.4,
    align: 'center', fontSize: 11, color: 'A8A8A8' });
  chrome(s, 20);

  heading(s, 'THANK YOU ', 'FOR ATTENTION',
    { x: 5.312, y: 1.201, w: 6.73, h: 1.919, fontSize: 54 });
  tx(s, 'We look forward to guiding you on your adventure',
    { x: 7.97, y: 4.695, w: 2.566, h: 0.678, fontFace: F.bodyMed, fontSize: 12,
      color: C.white, lineSpacingMultiple: 1.5 });
  tx(s, 'Lorem ipsum dolor sit amet, cons adipisicing elit, sed do eiusmod tempor incididunt ut. Magna.',
    { x: 1.046, y: 4.266, w: 3.135, h: 1.01, fontSize: 12, color: C.body, lineSpacingMultiple: 1.5 });

  rect(s, { x: 1.067, y: 5.455, w: 3.033, h: 0.475, fill: { color: C.dark } });
  tx(s, 'End Presentation', { x: 1.598, y: 5.524, w: 1.972, h: 0.337,
    align: 'center', fontFace: F.semi, fontSize: 14, color: C.white });
}

/* ==================================================================== build */
function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE';
  pptx.author = 'Summito';
  pptx.title = 'Summito Presentation Template 2025';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
   slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach((fn) => fn(pptx));

  return pptx.writeFile({
    fileName: path.join(__dirname, '01aaf313-e75a-4348-b300-7db7386803ac_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
