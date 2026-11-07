/*
 * Avast - Business Multipurpose  (34 slides, 20 x 11.25 in)
 *
 * Standalone re-creation of the reference deck with pptxgenjs only.
 * Run:  node 0026fe40-e808-4023-8e07-b59f8424918c_grok_final.js
 *
 * Raster artwork in the original (logo marks, icon glyphs, photo frames) is
 * redrawn here with native pptxgenjs shapes -- no embedded image data.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const BLUE  = '1E5CE3';   // brand accent
const INK   = '0D0D0D';   // headings
const DARK  = '262626';   // secondary dark
const BODY  = '595959';   // body copy
const WHITE = 'FFFFFF';
const LIGHT = 'F2F2F2';   // body copy on blue
const PAGE  = 'FBFBFB';   // infographic page tint
const GRAY  = '808080';
const SLATE = '404040';
const BLACK = '000000';
const FRAME = 'F5F5F5';   // hairline for the deck's empty picture slots
const SHEER = '4A7BE9';   // same hairline, legible on a blue panel

/* -------------------------------------------------------------------- fonts */
const XB = 'Plus Jakarta Sans ExtraBold';
const SB = 'Plus Jakarta Sans SemiBold';
const MD = 'Plus Jakarta Sans Medium';

/* --------------------------------------------------------------- body copy */
/* Filler paragraphs reused across the deck. */
const T1 =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas port itor congue massa. Fusce posuere, magna sed pulvinar ultricies';
const T2 =
  'Welcome                       About                       Services                       Team                       Portfolio                       Contact';
const T3 =
  'Lorem ipsum dolor sit amet, consectetue adipiscing elit. Maecenas porttitor congu massa. Fusce posuere, magna sed';
const T4 =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas port itor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lei ctus malesuada libero, sit amet commodo magna eros quis urna. Nuncya viverra imperdiet enim. Fusce est. Vivamus a tellus.';
const T5 =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas port itor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lei ctus malesuada libero, sit amet commodo magna eros quis urna. Nuncya viverra imperdiet enim. Fusce est.';
const T6 =
  'About                            Services                            Team                            Portfolio                            Contact';
const T7 =
  'Lorem ipsum dolor sit amet, consectetue adipisc ing elit. Maecenas porttitor congu massa. Fusce posuere, magna sed';
const T8 =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas port itor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lei ctus malesuada libero, sit amet commodo magna eros';
const T9 =
  'About                     Services                     Team                     Portfolio';
const T10 =
  'Lorem ipsum dolor sit amet, consectetue adipiscing elit. Maecenas porttitor congu massa. Fusce posuere, ';
const T11 =
  'Lorem ipsum dolor sit amet, consectetue adipiscing elit. ';
const T12 =
  'Lorem ipsum dolor sit amet, consec tetue adipiscing elit. Maecenas';
const T13 =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas port itor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lei ctus malesuada libero, sit amet commodo magna eros quis urna. ';
const T14 =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas port itor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lei ctus malesuada libero, ';
const T15 =
  'Lorem ipsum dolor sit amet, conse adipiscing elit. Maecenas porttitor massa. Fusce posuere';
const T16 =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero';
const T17 =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congu massa. Fusce posuere, ';
const T18 =
  'Lorem ipsum dolor sit amet, consectet uer adipiscing elit. Maecenas porttitor';
const T19 =
  'Lorem ipsum dolor sit amet, consectet uer adipiscing';
const T20 =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenasasr porttitoras congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus males uada libero, sit amet';
const T21 =
  'Fusce posuere, magna sed pulvinar ultricies, purus';
const T22 =
  'Lorem ipsum dolor sit amet, consect etue adipiscing elit. Maecenas portti tor congu massa. Fusce';
const T23 =
  'Lorem ipsum dolor sit ametar we, conse adipiscing elit. Maec enas porttitor massa';
const T24 =
  'Lorem ipsum dolor sit amet, conse ctetue adipiscing elit. Maecenas';
const T25 =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas port itor co ngue massa. Fusce posuere, magna sed pulvinar ultricies, purus lei ctus malesua da libero, sit amet commodo magna eros quis urna. ';
const T26 =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas port itor co ngue massa. Fusce posuere, magna sed pulvinar';
const T27 =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitoras congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus males uada libero, sit amet commodo magna eros quis urna.';
const T28 =
  'Lorem ipsum dolor sit ameta, conse ctetue adipiscing elit.';
const T29 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris';
const T30 =
  'Lorem ipsum dolor sit amet, cons ectetur adipiscing elit, ';
const T31 =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus.';
const T32 =
  'Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin pharetra nonummy pede. Mauris et';
const T33 =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fuscee posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus.';
const T34 =
  'Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin pharetra nonummy pede. Mauris etmagna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. ';
const T35 =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.';
const T36 =
  'Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin pharetra nonummy pede. Mauris et orci.';
const T37 =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fuscee posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. ';
const T38 =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitoras congue massa. Fusce posuere, magna sed pulvinar ultricies, ';

/* -------------------------------------------------- custom shape geometries */
/* Outlines normalised to the 0..1 box of the shape they are drawn into.      */
/* `null` splits the outline into separate (independently filled) pieces.     */
const GEOM = {
  // one leaf of the double price-tag glyph (drawn twice by tag())
  tag: [[0.968, 0.653], [0.653, 0.979], [0.063, 0.505], [0, 0.358], [0, 0.084],
        [0.084, 0], [0.358, 0], [0.505, 0.063], [0.968, 0.537], [1, 0.589]],
  // statistics diagram: lit face of a peak + its shaded right face
  spike:      [[0.496, 0], [1, 1], [0, 1]],
  spikeShade: [[0, 0], [0, 1], [1, 1]],
  // horizontal process diagram markers
  pennantDown: [[0, 0], [1, 0], [1, 0.5], [0.5, 1], [0, 0.5]],
  pennantUp:   [[0.5, 0], [1, 0.5], [1, 1], [0, 1], [0, 0.5]],
  // "process with half frames" chevrons (outer band + inner band)
  chevron:   [[0.5, 0], [1, 1], [0.875, 1], [0.5, 0.25], [0.125, 1], [0, 1]],
  chevronIn: [[0.5, 0], [1, 1], [0.946, 1], [0.5, 0.108], [0.054, 1], [0, 1]],
  // four curved-up arrows infographic
  arrow1: [[0.762, 0.007], [0.599, 0.101], [0.633, 0.127], [0.674, 0.144], [0.674, 0.739],
           [0.623, 0.778], [0, 1], [0.689, 1], [0.911, 0.766], [0.918, 0.752],
           [0.918, 0.144], [0.959, 0.127], [0.992, 0.101], [0.829, 0.007]],
  arrow2: [[0.262, 0.831], [0.171, 0.779], [0.171, 0.114], [0.193, 0.100], [0.211, 0.079],
           [0.125, 0.006], [0.090, 0.006], [0.004, 0.079], [0.022, 0.100], [0.043, 0.114],
           [0.043, 0.786], [0.103, 0.829], [0.638, 1], [1, 1]],
  arrow3: [[0.326, 0.687], [0.326, 0.173], [0.367, 0.153], [0.401, 0.121], [0.238, 0.008],
           [0.171, 0.008], [0.008, 0.121], [0.041, 0.153], [0.082, 0.173], [0.082, 0.702],
           [0.089, 0.719], [0.311, 1], [1, 1], [0.377, 0.733]],
  arrow4: [[0.957, 0.754], [0.957, 0.131], [0.978, 0.116], [0.996, 0.091], [0.910, 0.006],
           [0.875, 0.006], [0.789, 0.091], [0.807, 0.116], [0.828, 0.131], [0.828, 0.746],
           [0.738, 0.806], [0, 1], [0.362, 1], [0.897, 0.804]],
};

/* ------------------------------------------------------------------ helpers */
const NONE = { type: 'none' };
const INSET = [7.2, 7.2, 3.6, 3.6];        // l, r, b, t text inset in points

function newSlide(p, bg) {
  const s = p.addSlide();
  s.background = { color: bg || WHITE };
  return s;
}

/** Rectangle / ellipse / triangle / … with solid, translucent or no fill. */
function box(s, r, o) {
  o = o || {};
  const opt = { x: r[0], y: r[1], w: r[2], h: r[3], line: NONE };
  opt.fill = o.fill ? { color: o.fill } : NONE;
  if (o.alpha !== undefined) opt.fill.transparency = o.alpha;
  if (o.stroke) opt.line = { color: o.stroke, width: o.weight || 1 };
  if (o.rotate) opt.rotate = o.rotate;
  if (o.radius) opt.rectRadius = o.radius;
  if (o.fade) return fade(s, r, o.fade);
  s.addShape(o.shape || 'rect', opt);
}

/** Vertical transparent -> solid wash, stacked from translucent slices because
 *  pptxgenjs has no gradient fill. */
function fade(s, r, color) {
  const bands = 24;
  const step = r[3] / bands;
  for (let i = 0; i < bands; i++) {
    s.addShape('rect', {
      x: r[0], y: r[1] + step * i, w: r[2], h: step + 0.008,
      fill: { color, transparency: 100 - Math.round((100 * (i + 1)) / bands) }, line: NONE,
    });
  }
}

/** Free-form outline from the GEOM table, scaled into the given rectangle. */
function poly(s, r, color, kind, o) {
  o = o || {};
  const fill = { color };
  if (o.alpha) fill.transparency = o.alpha;
  const pieces = [[]];
  GEOM[kind].forEach((pt) => (pt === null ? pieces.push([]) : pieces[pieces.length - 1].push(pt)));
  pieces.forEach((piece) => {
    const pts = piece.map((pt, i) => ({ x: r[2] * pt[0], y: r[3] * pt[1], moveTo: i === 0 }));
    pts.push({ close: true });
    s.addShape('custGeom', {
      x: r[0], y: r[1], w: r[2], h: r[3], points: pts, fill, rotate: o.rotate || 0,
      line: o.edge ? { color: o.edge, width: 1 } : NONE,
    });
  });
}

/** Double price-tag glyph: two overlapping leaves plus the punch hole. */
function tag(s, r, color) {
  const back = color === WHITE ? BLUE : WHITE;   // the panel the glyph is stamped on
  const leaf = r[2] * 0.78;
  poly(s, [r[0] + r[2] - leaf, r[1], leaf, r[3]], color, 'tag');
  poly(s, [r[0], r[1], leaf, r[3]], color, 'tag', { edge: back });
  s.addShape('ellipse', {
    x: r[0] + r[2] * 0.1, y: r[1] + r[3] * 0.14, w: r[2] * 0.13, h: r[3] * 0.16,
    fill: { color: back }, line: NONE,
  });
}

/** Straight rule / arrow. */
function rule(s, r, o) {
  s.addShape('line', {
    x: r[0], y: r[1], w: r[2], h: r[3], flipH: !!o.flipH, flipV: !!o.flipV,
    line: { color: o.color, width: o.width || 1, endArrowType: o.arrow || 'none' },
  });
}

/** Multi-run paragraph block; each run is [text, size, color, font, breakLine]. */
function rich(s, r, runs, o) {
  o = o || {};
  s.addText(
    runs.map((t) => ({
      text: t[0],
      options: { fontSize: t[1], color: t[2], fontFace: t[3], breakLine: !!t[4] },
    })),
    {
      x: r[0], y: r[1], w: r[2], h: r[3], margin: INSET, isTextBox: true,
      align: o.align || 'left', valign: o.valign || 'top',
      wrap: o.wrap !== 0, lineSpacingMultiple: o.line,
    }
  );
}

/** Single-run shorthand. */
function txt(s, r, text, size, color, font, o) {
  rich(s, r, [[text, size, color, font]], o);
}

/** Stand-in for an icon glyph, drawn entirely with native shapes. */
function icon(s, r, color, kind) {
  const [x, y, w, h] = r;
  const at = (fx, fy, fw, fh, shape, opt) => s.addShape(shape, Object.assign(
    { x: x + w * fx, y: y + h * fy, w: w * fw, h: h * fh, fill: { color }, line: NONE }, opt || {}
  ));
  const stroke = Math.max(1, w * 3);
  const ring = (fx, fw, sweep) => at(fx, fx, fw, fw, sweep ? 'arc' : 'ellipse',
    { fill: NONE, line: { color, width: stroke }, angleRange: sweep });
  switch (kind) {
    case 'target':                                       // dartboard pierced by a dart
      ring(0, 1, [30, 340]);
      ring(0.16, 0.68, [30, 300]);
      ring(0.31, 0.38, [30, 260]);
      at(0.44, 0.42, 0.14, 0.14, 'ellipse');
      at(0.46, 0.16, 0.42, 0.42, 'line',
        { fill: NONE, line: { color, width: stroke, endArrowType: 'triangle' }, flipV: true });
      break;
    case 'gem':                                          // faceted diamond
      at(0.05, 0.1, 0.9, 0.8, 'diamond');
      break;
    case 'till':                                         // cash register
      at(0.1, 0.08, 0.8, 0.3, 'roundRect', { rectRadius: 0.06 });
      at(0.02, 0.44, 0.96, 0.48, 'trapezoid');
      break;
    case 'runner':                                       // sprinter off the blocks
      at(0.74, 0.04, 0.2, 0.2, 'ellipse');
      at(0.24, 0.3, 0.56, 0.16, 'roundRect', { rotate: 15, rectRadius: 0.08 });
      at(0.06, 0.36, 0.28, 0.14, 'roundRect', { rotate: 330, rectRadius: 0.07 });
      at(0.3, 0.52, 0.16, 0.42, 'roundRect', { rotate: 25, rectRadius: 0.07 });
      at(0.6, 0.5, 0.16, 0.44, 'roundRect', { rotate: 335, rectRadius: 0.07 });
      break;
    case 'chat':                                         // two speech bubbles
      at(0, 0.04, 0.64, 0.46, 'wedgeRoundRectCallout', { rectRadius: 0.12 });
      at(0.36, 0.46, 0.64, 0.46, 'wedgeRoundRectCallout', { rectRadius: 0.12 });
      break;
    case 'coins':                                        // stack of coins
      [[0.06, 0.56], [0.2, 0.34], [0.34, 0.12]].forEach((o) =>
        at(o[0], o[1], 0.6, 0.3, 'donut', { fill: NONE, line: { color, width: 4 } }));
      break;
    case 'people':                                       // group of three
      [0, 0.34, 0.68].forEach((fx, i) => {
        const fy = i === 1 ? 0.24 : 0.1;
        at(fx + 0.04, fy, 0.24, 0.24, 'ellipse');
        at(fx, fy + 0.3, 0.32, 0.34, 'round2SameRect');
      });
      break;
    case 'puzzle':                                       // jigsaw piece with knobs
      at(0.16, 0.16, 0.68, 0.68, 'rect');
      at(0.62, 0.02, 0.3, 0.3, 'ellipse');
      at(0.72, 0.36, 0.28, 0.28, 'ellipse');
      at(0.36, 0.7, 0.28, 0.28, 'ellipse');
      at(0.0, 0.36, 0.28, 0.28, 'ellipse');
      break;
    case 'bulb':                                         // lightbulb
      ring(0.14, 0, 0.72, Math.max(1, w * 3));
      at(0.34, 0.66, 0.32, 0.1, 'rect');
      at(0.38, 0.82, 0.24, 0.1, 'rect');
      break;
    case 'flame':                                        // flame with a curl
      at(0.1, 0.14, 0.62, 0.76, 'teardrop', { rotate: 225 });
      at(0.5, 0.4, 0.4, 0.5, 'teardrop', { rotate: 225 });
      break;
    case 'tablet':                                       // tablet with stylus alongside
      at(0.005, 0.09, 0.03, 0.45, 'roundRect', { fill: { color: GRAY }, rectRadius: 0.02 });
      at(0.06, 0.02, 0.92, 0.96, 'roundRect',
        { fill: NONE, line: { color, width: 14 }, rectRadius: 0.2 });
      break;
    default:
      at(0.12, 0.12, 0.76, 0.76, 'roundRect', { rectRadius: 0.1 });
  }
}

/** Picture slot. The reference deck ships these placeholders empty, so only the
 *  frame is drawn -- nothing is painted over the artwork beneath it. */
function photo(s, r, tone) {
  s.addShape('rect', {
    x: r[0], y: r[1], w: r[2], h: r[3], fill: NONE, line: { color: tone || FRAME, width: 0.75 },
  });
}

/** Header logo, hamburger bars and page footer shared by most slides. */
function chrome(s, o) {
  if (o.logo !== undefined) {
    txt(s, [0.952, o.logo, 1.157, 0.505], 'Avast', 24, o.logoColor || BLUE, XB, { wrap: 0 });
  }
  if (o.mark) icon(s, [1.347, 0.47, 0.368, 0.368], o.mark, 'target');
  if (o.bars) {
    const bx = o.barsX === undefined ? 18.57 : o.barsX;
    box(s, [bx, 0.929, 0.479, 0.079], { fill: o.bars });
    box(s, [bx, 1.138, 0.479, 0.079], { fill: o.bars });
  }
  if (o.foot) {
    txt(s, [0.952, 10.104, 2.421, 0.37], 'Business', 16, o.foot, XB);
    txt(s, [16.628, 10.104, 2.421, 0.37], 'Presentation', 16, o.foot, XB, { align: 'right' });
  }
}

/* ------------------------------------------------------------------- slides */

function slide01(p) {
  const s = newSlide(p);
  txt(s, [6.601, 3.343, 6.797, 2.895], 'Avast', 166, BLUE, XB, { align: 'center', wrap: 0 });
  txt(s, [6.889, 6.285, 6.223, 0.707], 'Business Multipurpose', 36, INK, SB, { align: 'center' });
  icon(s, [9.063, 1.834, 1.875, 1.875], INK, 'target');
  chrome(s, { logo: 0.775, bars: INK, foot: INK, mark: INK });
  txt(s, [6.081, 7.494, 7.926, 0.855], T1, 16, BODY, MD, { align: 'center', line: 1.5 });
}

function slide02(p) {
  const s = newSlide(p);
  box(s, [9.968, 9.086, 0.476, 0.927], { shape: 'roundRect', stroke: BLUE, weight: 1, radius: 0.24 });
  box(s, [10.044, 9.58, 0.328, 0.328], { shape: 'ellipse', fill: BLUE });
  chrome(s, { logo: 0.713, bars: INK, foot: INK });
  txt(s, [1.333, 3.577, 7.926, 2.794], 'Multipurpose Business', 80, BLUE, XB);
  txt(s, [5.137, 0.837, 9.815, 0.381], T2, 16, INK, SB, { align: 'center' });
  txt(s, [1.438, 2.757, 4.907, 0.572], 'Presentation Template.', 28, INK, XB);
  txt(s, [12.2, 3.577, 6.467, 0.761], T3, 14, BODY, MD, { line: 1.5 });
  txt(s, [12.2, 2.933, 2.195, 0.404], 'CopyRight 2023', 18, INK, XB);
  txt(s, [12.2, 4.742, 2.195, 0.404], 'Read More', 18, BLUE, SB);
  rule(s, [13.946, 4.944, 0.584, 0], { color: BLUE, width: 3, arrow: 'triangle' });
  txt(s, [1.438, 6.902, 7.926, 1.663], T4, 16, BODY, MD, { line: 1.5 });
  txt(s, [12.2, 6.786, 6.467, 0.761], T3, 14, BODY, MD, { line: 1.5 });
  txt(s, [12.2, 6.142, 2.195, 0.404], 'CopyRight 2023', 18, INK, XB);
  txt(s, [12.2, 7.952, 2.195, 0.404], 'Read More', 18, BLUE, SB);
  rule(s, [13.946, 8.154, 0.584, 0], { color: BLUE, width: 3, arrow: 'triangle' });
}

function slide03(p) {
  const s = newSlide(p);
  box(s, [15.694, 0, 4.306, 11.25], { fill: BLUE });
  chrome(s, { logo: 0.713, bars: WHITE, barsX: 18.443 });
  rich(s, [1.33, 2.46, 7.926, 2.12], [
    ['About Business ', 60, INK, XB],
    ['Multipurpose', 60, BLUE, XB],
  ]);
  txt(s, [1.33, 5.301, 10.109, 1.259], T5, 16, BODY, MD, { line: 1.5 });
  txt(s, [3.692, 0.837, 9.815, 0.381], T6, 16, INK, SB);
  txt(s, [1.33, 8.319, 4.987, 1.259], T3, 16, BODY, MD, { line: 1.5 });
  txt(s, [1.33, 7.675, 3.522, 0.438], 'Section One', 20, INK, XB);
  txt(s, [6.52, 8.319, 4.987, 1.259], T3, 16, BODY, MD, { line: 1.5 });
  txt(s, [6.52, 7.675, 3.522, 0.438], 'Section Two', 20, INK, XB);
  photo(s, [12.903, 2.222, 6.111, 8.191]);
}

function slide04(p) {
  const s = newSlide(p);
  box(s, [0, 0, 5.972, 11.25], { fill: BLUE });
  txt(s, [1.314, 1.538, 3.987, 1.447], 'Multipurpose Business', 40, WHITE, XB);
  rich(s, [11.473, 4.556, 7.213, 2.12], [
    ['About Business ', 60, INK, XB],
    ['Multipurpose', 60, BLUE, XB],
  ]);
  txt(s, [11.473, 7.323, 7.448, 1.663], T5, 16, BODY, MD, { line: 1.5 });
  txt(s, [6.965, 1.817, 5.495, 1.259], T7, 16, BODY, MD, { line: 1.5 });
  txt(s, [6.965, 1.173, 3.522, 0.438], 'Section One', 20, INK, XB);
  txt(s, [13.47, 1.817, 5.495, 1.259], T7, 16, BODY, MD, { line: 1.5 });
  txt(s, [13.47, 1.173, 3.522, 0.438], 'Section Two', 20, INK, XB);
  photo(s, [1.313, 4.189, 8.274, 5.889]);
}

function slide05(p) {
  const s = newSlide(p);
  box(s, [10.695, 0.889, 3.361, 6.107], { fill: BLUE });
  rich(s, [1.33, 2.498, 7.926, 2.12], [
    ['About Business ', 60, INK, XB],
    ['Multipurpose', 60, BLUE, XB],
  ]);
  txt(s, [1.33, 5.34, 8.21, 1.663], T5, 16, BODY, MD, { line: 1.5 });
  txt(s, [1.33, 7.564, 8.21, 1.259], T8, 16, BODY, MD, { line: 1.5 });
  txt(s, [3.026, 0.78, 6.752, 0.37], T9, 16, INK, SB);
  chrome(s, { logo: 0.713, bars: INK, foot: INK });
  photo(s, [16.639, 2.312, 3.361, 6.627]);
  photo(s, [11.667, 2.312, 4.431, 6.627]);
}

function slide06(p) {
  const s = newSlide(p);
  box(s, [0, 8.038, 6.603, 2.65], { fill: BLUE });
  box(s, [4.569, 1.54, 2.79, 2.65], { fill: BLUE });
  rich(s, [8.399, 1.54, 10.187, 1.582], [
    ['The most important investment ', 44, INK, XB],
    ['you can make is in yourself.', 44, BLUE, XB],
  ]);
  txt(s, [8.399, 4, 9.996, 1.259], T5, 16, BODY, MD, { line: 1.5 });
  txt(s, [8.399, 7.053, 5.823, 1.259], T3, 16, BODY, MD, { line: 1.5 });
  txt(s, [8.399, 6.409, 2.195, 0.404], 'CopyRight 2023', 18, INK, XB);
  txt(s, [8.399, 8.552, 2.195, 0.404], 'Read More', 18, BLUE, SB);
  rule(s, [10.145, 8.754, 0.584, 0], { color: BLUE, width: 3, arrow: 'triangle' });
  txt(s, [15.013, 7.113, 3.381, 1.663], T10, 16, BODY, MD, { line: 1.5 });
  txt(s, [15.013, 6.469, 3.018, 0.438], 'Section Here', 20, INK, XB);
  photo(s, [0, 0, 6.603, 11.25]);
}

function slide07(p) {
  const s = newSlide(p);
  photo(s, [10, 0, 10, 11.25]);
  rich(s, [1.288, 1.555, 7.982, 2.322], [
    ['The most important investment ', 44, INK, XB],
    ['you can make is in yourself.', 44, BLUE, XB],
  ]);
  txt(s, [1.288, 4.762, 7.833, 1.663], T5, 16, BODY, MD, { line: 1.5 });
  box(s, [1.288, 7.372, 15.347, 2.758], { shape: 'parallelogram', fill: BLUE });
  txt(s, [2.76, 8.652, 3.225, 0.855], T11, 16, LIGHT, MD, { line: 1.5 });
  txt(s, [2.76, 8.008, 3.018, 0.438], 'Section One', 20, WHITE, XB);
  txt(s, [7.22, 8.652, 3.225, 0.855], T11, 16, LIGHT, MD, { line: 1.5 });
  txt(s, [7.22, 8.008, 3.018, 0.438], 'Section  Two', 20, WHITE, XB);
  txt(s, [11.679, 8.652, 3.225, 0.855], T11, 16, LIGHT, MD, { line: 1.5 });
  txt(s, [11.679, 8.008, 3.018, 0.438], 'Section Three', 20, WHITE, XB);
}

function slide08(p) {
  const s = newSlide(p);
  box(s, [0, 0, 3.361, 6.107], { fill: BLUE });
  rich(s, [10.861, 1.333, 7.926, 2.524], [
    ['Risk comes from not ', 48, INK, XB],
    ['knowing what you\'re doing.', 48, BLUE, XB],
  ]);
  txt(s, [10.861, 4.444, 8.21, 1.663], T5, 16, BODY, MD, { line: 1.5 });
  txt(s, [1.083, 8.367, 4.987, 1.259], T3, 16, BODY, MD, { line: 1.5 });
  txt(s, [1.083, 7.723, 3.522, 0.438], 'Section One', 20, INK, XB);
  txt(s, [7.507, 8.367, 4.987, 1.259], T3, 16, BODY, MD, { line: 1.5 });
  txt(s, [7.507, 7.723, 3.522, 0.438], 'Section Two', 20, INK, XB);
  txt(s, [13.93, 8.367, 4.987, 1.259], T3, 16, BODY, MD, { line: 1.5 });
  txt(s, [13.93, 7.723, 3.522, 0.438], 'Section Three', 20, INK, XB);
  photo(s, [1.083, 0, 8.056, 5.389]);
}

function slide09(p) {
  const s = newSlide(p);
  box(s, [16.875, 0, 3.125, 11.25], { fill: BLUE });
  rich(s, [1.306, 1.991, 8.568, 1.919], [
    ['Price is what you pay; ', 54, INK, XB],
    ['value is what you get.', 54, BLUE, XB],
  ]);
  txt(s, [1.306, 4.716, 8.694, 1.663], T5, 16, BODY, MD, { line: 1.5 });
  icon(s, [1.494, 6.821, 0.631, 0.631], BLUE, 'gem');
  txt(s, [1.306, 8.404, 4.012, 0.855], T12, 16, BODY, MD, { line: 1.5 });
  txt(s, [1.306, 7.76, 2.64, 0.438], 'Our Section One', 20, INK, XB);
  txt(s, [6.005, 8.404, 4.012, 0.855], T12, 16, BODY, MD, { line: 1.5 });
  txt(s, [6.005, 7.76, 2.64, 0.438], 'Our Section Two', 20, INK, XB);
  icon(s, [6.194, 6.821, 0.631, 0.631], BLUE, 'till');
  photo(s, [11.069, 1.075, 3.686, 6.647]);
  photo(s, [15.361, 3.311, 3.686, 6.647]);
}

function slide10(p) {
  const s = newSlide(p);
  box(s, [6.633, 5.824, 8.568, 1.033], { fill: BLUE });
  box(s, [0, 0, 5.069, 3.698], { fill: BLUE });
  box(s, [17.335, 7.54, 2.665, 3.71], { fill: BLUE });
  txt(s, [0.952, 9.16, 3.366, 0.855], T11, 16, BODY, MD, { line: 1.5 });
  txt(s, [0.952, 8.629, 2.195, 0.404], 'Lorem Ipsum', 18, INK, XB);
  tag(s, [1.218, 7.96, 0.504, 0.404], BLUE);
  txt(s, [4.803, 9.16, 3.366, 0.855], T11, 16, BODY, MD, { line: 1.5 });
  txt(s, [4.803, 8.629, 2.195, 0.404], 'Lorem Ipsum', 18, INK, XB);
  tag(s, [5.069, 7.96, 0.504, 0.404], BLUE);
  txt(s, [9.236, 8.095, 3.858, 0.404], 'Consectetue adipiscing elit. ', 18, INK, XB);
  txt(s, [9.236, 8.759, 3.858, 0.404], 'Fusce posuere, magna sed', 18, INK, XB);
  txt(s, [9.236, 9.423, 3.858, 0.404], 'Purus lectus malesuada libero', 18, INK, XB);
  rich(s, [10, 1.438, 8.568, 1.717], [
    ['The best investment you ', 48, INK, XB],
    ['can make is in yourself.', 48, BLUE, XB],
  ]);
  txt(s, [10, 3.635, 8.694, 1.663], T5, 16, BODY, MD, { line: 1.5 });
  photo(s, [0.877, 0.872, 7.615, 5.985]);
  photo(s, [14.417, 5.778, 4.365, 4.683]);
}

function slide11(p) {
  const s = newSlide(p);
  box(s, [0, 7.514, 20, 3.736], { fill: BLUE });
  rich(s, [1.681, 1.912, 11.819, 1.717], [
    ['The goal of retirement is to live off ', 48, INK, XB],
    ['your assets, not on them.', 48, BLUE, XB],
  ]);
  txt(s, [1.681, 4.157, 11.965, 0.855], T13, 16, BODY, MD, { line: 1.5 });
  txt(s, [11.847, 8.914, 7.622, 1.259], T14, 16, WHITE, MD, { line: 1.5 });
  txt(s, [11.847, 8.126, 6.472, 0.505], 'Fusce posuere, magna sed pulvinar', 24, WHITE, XB);
  icon(s, [15.025, 0.725, 0.631, 0.631], BLUE, 'gem');
  txt(s, [14.836, 2.308, 4.012, 0.855], T12, 16, BODY, MD, { line: 1.5 });
  txt(s, [14.836, 1.664, 2.64, 0.438], 'Our Section One', 20, INK, XB);
  txt(s, [14.836, 5.625, 4.012, 0.855], T12, 16, BODY, MD, { line: 1.5 });
  txt(s, [14.836, 4.981, 2.64, 0.438], 'Our Section Two', 20, INK, XB);
  icon(s, [15.025, 4.042, 0.631, 0.631], BLUE, 'till');
  txt(s, [1.681, 1.04, 4.333, 0.505], 'Business Multipurpose', 24, BLUE, SB);
  photo(s, [1.681, 6.014, 4.042, 4.224]);
  photo(s, [6.764, 6.014, 4.042, 4.224]);
}

function slide12(p) {
  const s = newSlide(p);
  photo(s, [0, 0, 20, 11.25]);
  box(s, [0, 0, 20, 11.25], { fill: BLUE, alpha: 28 });
  txt(s, [4.089, 3.611, 11.821, 2.423], 'Break Slides', 138, WHITE, XB, { align: 'center', wrap: 0 });
  txt(s, [6.889, 6.285, 6.223, 0.707], 'Business Multipurpose', 36, WHITE, SB, { align: 'center' });
  icon(s, [9.063, 1.834, 1.875, 1.875], WHITE, 'target');
  chrome(s, { logo: 0.775, logoColor: WHITE, bars: WHITE, foot: WHITE, mark: WHITE });
  txt(s, [6.081, 7.494, 7.926, 0.855], T1, 16, WHITE, MD, { align: 'center', line: 1.5 });
}

function slide13(p) {
  const s = newSlide(p);
  rich(s, [1.304, 1.349, 9.357, 2.12], [
    ['Our Services Business ', 60, INK, XB],
    ['Multipurpose', 60, BLUE, XB],
  ]);
  box(s, [1.423, 6.747, 0.785, 0.785], { fill: BLUE });
  txt(s, [1.321, 8.786, 3.46, 1.114], T15, 14, BODY, MD, { line: 1.5 });
  txt(s, [1.321, 8.001, 3.46, 0.505], 'Our Services One', 24, INK, XB);
  tag(s, [1.633, 6.992, 0.366, 0.293], WHITE);
  box(s, [15.484, 6.747, 0.785, 0.785], { fill: BLUE });
  txt(s, [15.382, 8.786, 3.46, 1.114], T15, 14, BODY, MD, { line: 1.5 });
  txt(s, [15.382, 8.001, 3.46, 0.505], 'Our Services Four', 24, INK, XB);
  tag(s, [15.694, 6.992, 0.366, 0.293], WHITE);
  box(s, [6.042, 6.747, 0.785, 0.785], { fill: BLUE });
  txt(s, [5.939, 8.786, 3.46, 1.114], T15, 14, BODY, MD, { line: 1.5 });
  txt(s, [5.939, 8.001, 3.46, 0.505], 'Our Services Two', 24, INK, XB);
  tag(s, [6.252, 6.992, 0.366, 0.293], WHITE);
  box(s, [10.763, 6.747, 0.785, 0.785], { fill: BLUE });
  txt(s, [10.66, 8.786, 3.46, 1.114], T15, 14, BODY, MD, { line: 1.5 });
  txt(s, [10.66, 8.001, 3.46, 0.505], 'Our Services Three', 24, INK, XB);
  tag(s, [10.973, 6.992, 0.366, 0.293], WHITE);
  txt(s, [12.54, 2.948, 6.608, 1.259], T16, 16, BODY, MD, { line: 1.5 });
  txt(s, [12.54, 2.218, 5.743, 0.438], 'Fusce posuere magna sed pulvinar ', 20, INK, XB);
  txt(s, [1.321, 4.338, 9.34, 1.259], T13, 16, BODY, MD, { line: 1.5 });
  txt(s, [12.54, 4.677, 2.195, 0.404], 'Read More', 18, BLUE, SB);
  rule(s, [14.286, 4.879, 0.584, 0], { color: BLUE, width: 3, arrow: 'triangle' });
}

function slide14(p) {
  const s = newSlide(p);
  rich(s, [1.271, 5.605, 9.357, 2.12], [
    ['Our Services Business ', 60, INK, XB],
    ['Multipurpose', 60, BLUE, XB],
  ]);
  txt(s, [1.288, 8.594, 9.34, 1.259], T13, 16, BODY, MD, { line: 1.5 });
  txt(s, [1.288, 2.128, 4.792, 1.259], T17, 16, BODY, MD, { line: 1.5 });
  txt(s, [1.288, 1.397, 4.792, 0.438], 'Fusce posuere magna sed pulvi', 20, INK, XB);
  txt(s, [1.288, 3.856, 2.195, 0.404], 'Read More', 18, BLUE, SB);
  rule(s, [3.034, 4.058, 0.584, 0], { color: BLUE, width: 3, arrow: 'triangle' });
  txt(s, [6.353, 2.128, 4.792, 1.259], T17, 16, BODY, MD, { line: 1.5 });
  txt(s, [6.353, 1.397, 4.792, 0.438], 'Fusce posuere magna sed pulvi', 20, INK, XB);
  txt(s, [6.353, 3.856, 2.195, 0.404], 'Read More', 18, BLUE, SB);
  rule(s, [8.099, 4.058, 0.584, 0], { color: BLUE, width: 3, arrow: 'triangle' });
  box(s, [12.149, 1.444, 6.98, 8.481], { fill: BLUE });
  icon(s, [13.062, 2.369, 0.631, 0.631], WHITE, 'gem');
  txt(s, [13.88, 3.001, 4.653, 0.855], T18, 16, WHITE, MD, { line: 1.5 });
  txt(s, [13.88, 2.357, 2.64, 0.438], 'Our Services One', 20, WHITE, XB);
  icon(s, [13.062, 7.407, 0.631, 0.631], WHITE, 'gem');
  txt(s, [13.88, 8.038, 4.653, 0.855], T18, 16, WHITE, MD, { line: 1.5 });
  txt(s, [13.88, 7.394, 3.28, 0.438], 'Our Services Three', 20, WHITE, XB);
  icon(s, [13.062, 4.888, 0.631, 0.631], WHITE, 'gem');
  txt(s, [13.88, 5.519, 4.653, 0.855], T18, 16, WHITE, MD, { line: 1.5 });
  txt(s, [13.88, 4.875, 2.64, 0.438], 'Our Services Two', 20, WHITE, XB);
}

function slide15(p) {
  const s = newSlide(p);
  rich(s, [3.715, 2.153, 12.569, 2.12], [
    ['Our Services Avast Business ', 60, INK, XB],
    ['Multipurpose', 60, BLUE, XB],
  ], { align: 'center' });
  txt(s, [7.833, 1.315, 4.333, 0.505], 'Business Multipurpose', 24, BLUE, SB, { align: 'center' });
  chrome(s, { logo: 0.775, bars: INK, foot: INK, mark: INK });
  txt(s, [3.168, 4.837, 13.663, 0.855], T13, 16, BODY, MD, { align: 'center', line: 1.5 });
  box(s, [0.911, 6.548, 18.177, 2.701], { fill: BLUE });
  icon(s, [2.165, 7.127, 0.631, 0.631], WHITE, 'gem');
  txt(s, [2.983, 7.759, 3.516, 0.855], T19, 16, WHITE, MD, { line: 1.5 });
  txt(s, [2.983, 7.115, 2.64, 0.438], 'Our Services One', 20, WHITE, XB);
  icon(s, [13.501, 7.127, 0.631, 0.631], WHITE, 'gem');
  txt(s, [14.318, 7.759, 3.516, 0.855], T19, 16, WHITE, MD, { line: 1.5 });
  txt(s, [14.318, 7.115, 2.64, 0.438], 'Our Services One', 20, WHITE, XB);
  icon(s, [7.832, 7.127, 0.631, 0.631], WHITE, 'gem');
  txt(s, [8.65, 7.759, 3.516, 0.855], T19, 16, WHITE, MD, { line: 1.5 });
  txt(s, [8.65, 7.115, 2.64, 0.438], 'Our Services One', 20, WHITE, XB);
}

function slide16(p) {
  const s = newSlide(p);
  rich(s, [2.071, 6.984, 8.261, 2.524], [
    ['SWOT', 72, BLUE, XB],
    [' ', 72, INK, XB, 1],
    ['Analysis Slides', 72, INK, XB],
  ]);
  txt(s, [10.806, 8.07, 7.768, 1.259], T20, 16, BODY, MD, { line: 1.5 });
  txt(s, [10.806, 7.177, 7.768, 0.505], 'Fusce posuere, magna sed pulvinar ultricies, ', 24, INK, XB);
  chrome(s, { logo: 0.713, bars: INK, foot: INK });
  txt(s, [5.137, 0.837, 9.815, 0.381], T2, 16, INK, SB, { align: 'center' });
  photo(s, [2.071, 2.34, 15.858, 3.704]);
}

function slide17(p) {
  const s = newSlide(p);
  box(s, [12.73, 2.819, 7.27, 6.258], { fill: BLUE });
  txt(s, [14.43, 2.744, 3.87, 5.89], 'S', 344, WHITE, XB, { align: 'center' });
  rich(s, [1.127, 3.283, 9.079, 3.063], [
    ['Strengths ', 88, INK, XB, 1],
    ['Analysis Slides', 88, BLUE, XB],
  ]);
  txt(s, [1.127, 8.088, 9.227, 0.855], T16, 16, BODY, MD, { line: 1.5 });
  txt(s, [1.127, 7.254, 8.873, 0.505], T21, 24, INK, XB);
  chrome(s, { logo: 0.713, bars: INK, foot: INK });
  txt(s, [5.137, 0.837, 9.815, 0.381], T2, 16, INK, SB, { align: 'center' });
  txt(s, [1.127, 2.492, 4.333, 0.505], 'Business Multipurpose', 24, BLUE, SB);
}

function slide18(p) {
  const s = newSlide(p);
  box(s, [0, 2.819, 7.27, 6.258], { fill: BLUE });
  txt(s, [1.7, 2.744, 3.87, 5.89], 'W', 344, WHITE, XB, { align: 'center' });
  chrome(s, { logo: 0.713, bars: INK, foot: INK });
  txt(s, [5.137, 0.837, 9.815, 0.381], T2, 16, INK, SB, { align: 'center' });
  rich(s, [8.866, 3.283, 9.079, 3.063], [
    ['Weakneses ', 88, INK, XB, 1],
    ['Analysis Slides', 88, BLUE, XB],
  ]);
  txt(s, [8.866, 2.492, 4.333, 0.505], 'Business Multipurpose', 24, BLUE, SB);
  txt(s, [8.866, 7.624, 4.333, 1.259], T22, 16, BODY, MD, { line: 1.5 });
  txt(s, [8.866, 6.98, 3.522, 0.438], 'Section One', 20, INK, XB);
  txt(s, [13.998, 7.624, 4.333, 1.259], T22, 16, BODY, MD, { line: 1.5 });
  txt(s, [13.998, 6.98, 3.522, 0.438], 'Section One', 20, INK, XB);
}

function slide19(p) {
  const s = newSlide(p);
  box(s, [12.73, 2.819, 7.27, 6.258], { fill: BLUE });
  txt(s, [14.43, 2.744, 3.87, 5.89], 'O', 344, WHITE, XB, { align: 'center' });
  rich(s, [1.127, 3.283, 9.079, 3.063], [
    ['Opportunities ', 88, INK, XB, 1],
    ['Analysis Slides', 88, BLUE, XB],
  ]);
  chrome(s, { logo: 0.713, bars: INK, foot: INK });
  txt(s, [5.137, 0.837, 9.815, 0.381], T2, 16, INK, SB, { align: 'center' });
  txt(s, [1.127, 2.492, 4.333, 0.505], 'Business Multipurpose', 24, BLUE, SB);
  txt(s, [2.021, 7.864, 3.87, 1.259], T23, 16, BODY, MD, { line: 1.5 });
  txt(s, [2.021, 7.242, 3.46, 0.505], 'Our Services One', 24, INK, XB);
  box(s, [1.216, 7.298, 0.623, 0.623], { fill: BLUE });
  tag(s, [1.382, 7.493, 0.29, 0.232], WHITE);
  txt(s, [7.196, 7.864, 3.87, 1.259], T23, 16, BODY, MD, { line: 1.5 });
  txt(s, [7.196, 7.242, 3.46, 0.505], 'Our Services One', 24, INK, XB);
  box(s, [6.391, 7.298, 0.623, 0.623], { fill: BLUE });
  tag(s, [6.557, 7.493, 0.29, 0.232], WHITE);
}

function slide20(p) {
  const s = newSlide(p);
  box(s, [0, 2.819, 7.27, 6.258], { fill: BLUE });
  txt(s, [1.7, 2.744, 3.87, 5.89], 'T', 344, WHITE, XB, { align: 'center' });
  chrome(s, { logo: 0.713, bars: INK, foot: INK });
  txt(s, [5.137, 0.837, 9.815, 0.381], T2, 16, INK, SB, { align: 'center' });
  rich(s, [8.803, 3.283, 9.079, 3.063], [
    ['Threats ', 88, INK, XB, 1],
    ['Analysis Slides', 88, BLUE, XB],
  ]);
  txt(s, [8.803, 2.492, 4.333, 0.505], 'Business Multipurpose', 24, BLUE, SB);
  txt(s, [8.803, 7.307, 3.858, 0.404], 'Consectetue adipiscing elit. ', 18, INK, XB);
  txt(s, [8.803, 7.971, 3.858, 0.404], 'Fusce posuere, magna sed', 18, INK, XB);
  txt(s, [8.803, 8.635, 3.858, 0.404], 'Purus lectus malesuada libero', 18, INK, XB);
  box(s, [13.406, 6.875, 5.095, 2.701], { fill: BLUE });
  icon(s, [13.966, 7.454, 0.631, 0.631], WHITE, 'gem');
  txt(s, [14.784, 8.085, 3.368, 0.855], T19, 16, WHITE, MD, { line: 1.5 });
  txt(s, [14.784, 7.442, 2.82, 0.438], 'Our Services  Here', 20, WHITE, XB);
}

function slide21(p) {
  const s = newSlide(p);
  box(s, [0.589, 7.312, 1.633, 1.398], { fill: BLUE });
  box(s, [0.589, 3.962, 1.633, 1.398], { fill: BLUE });
  box(s, [0.589, 0.545, 1.633, 1.398], { fill: BLUE });
  txt(s, [4.705, 1.74, 4.375, 0.855], T24, 16, BODY, MD, { line: 1.5 });
  txt(s, [4.705, 1.134, 3.133, 0.404], 'Great Team Work One', 18, INK, XB);
  txt(s, [4.705, 2.835, 2.195, 0.404], 'Read More', 18, BLUE, SB);
  rule(s, [6.451, 3.037, 0.584, 0], { color: BLUE, width: 3, arrow: 'triangle' });
  txt(s, [4.705, 8.617, 4.375, 0.855], T24, 16, BODY, MD, { line: 1.5 });
  txt(s, [4.705, 8.011, 3.133, 0.404], 'Great Team Work Three', 18, INK, XB);
  txt(s, [4.705, 9.712, 2.195, 0.404], 'Read More', 18, BLUE, SB);
  rule(s, [6.451, 9.914, 0.584, 0], { color: BLUE, width: 3, arrow: 'triangle' });
  txt(s, [4.705, 5.065, 4.375, 0.855], T24, 16, BODY, MD, { line: 1.5 });
  txt(s, [4.705, 4.459, 3.133, 0.404], 'Great Team Work Two', 18, INK, XB);
  txt(s, [4.705, 6.16, 2.195, 0.404], 'Read More', 18, BLUE, SB);
  rule(s, [6.451, 6.362, 0.584, 0], { color: BLUE, width: 3, arrow: 'triangle' });
  rich(s, [9.866, 1.538, 8.844, 1.919], [
    ['Our Team Work Avast ', 54, INK, XB],
    ['Business Multipurpose', 54, BLUE, XB],
  ]);
  txt(s, [9.866, 3.912, 8.731, 1.259], T25, 16, BODY, MD, { line: 1.5 });
  box(s, [10.103, 6.187, 0.785, 0.785], { fill: BLUE });
  txt(s, [10, 8.226, 4.161, 1.259], T15, 16, BODY, MD, { line: 1.5 });
  txt(s, [10, 7.441, 3.46, 0.505], 'Our Services One', 24, INK, XB);
  tag(s, [10.312, 6.433, 0.366, 0.293], WHITE);
  box(s, [14.748, 6.187, 0.785, 0.785], { fill: BLUE });
  txt(s, [14.646, 8.226, 4.161, 1.259], T15, 16, BODY, MD, { line: 1.5 });
  txt(s, [14.646, 7.441, 3.46, 0.505], 'Our Services One', 24, INK, XB);
  tag(s, [14.958, 6.433, 0.366, 0.293], WHITE);
  photo(s, [1.141, 1.028, 2.778, 2.452]);
  photo(s, [1.141, 4.399, 2.778, 2.452]);
  photo(s, [1.141, 7.77, 2.778, 2.452]);
}

function slide22(p) {
  const s = newSlide(p);
  box(s, [0, 0, 20, 2.629], { fill: BLUE });
  txt(s, [1.141, 5.143, 4.714, 0.855], T24, 16, BODY, MD, { line: 1.5 });
  txt(s, [1.141, 4.537, 3.133, 0.404], 'Great Team Work One', 18, INK, XB);
  txt(s, [1.141, 6.239, 2.195, 0.404], 'Read More', 18, BLUE, SB);
  rule(s, [2.887, 6.44, 0.584, 0], { color: BLUE, width: 3, arrow: 'triangle' });
  txt(s, [7.54, 5.143, 4.714, 0.855], T24, 16, BODY, MD, { line: 1.5 });
  txt(s, [7.54, 4.537, 3.133, 0.404], 'Great Team Work Two', 18, INK, XB);
  txt(s, [7.54, 6.239, 2.195, 0.404], 'Read More', 18, BLUE, SB);
  rule(s, [9.285, 6.44, 0.584, 0], { color: BLUE, width: 3, arrow: 'triangle' });
  txt(s, [13.938, 5.143, 4.714, 0.855], T24, 16, BODY, MD, { line: 1.5 });
  txt(s, [13.938, 4.537, 3.133, 0.404], 'Great Team Work Three', 18, INK, XB);
  txt(s, [13.938, 6.239, 2.195, 0.404], 'Read More', 18, BLUE, SB);
  rule(s, [15.684, 6.44, 0.584, 0], { color: BLUE, width: 3, arrow: 'triangle' });
  rich(s, [1.026, 7.557, 7.41, 1.582], [
    ['Our Team Work Avast ', 44, INK, XB],
    ['Business Multipurpose', 44, BLUE, XB],
  ]);
  txt(s, [1.026, 9.508, 7.668, 0.855], T26, 16, BODY, MD, { line: 1.5 });
  txt(s, [10, 8.736, 8.974, 1.259], T27, 16, BODY, MD, { line: 1.5 });
  txt(s, [10, 7.843, 8.652, 0.505], T21, 24, INK, XB);
  photo(s, [1.141, 1.028, 4.921, 3.02]);
  photo(s, [7.54, 1.028, 4.921, 3.02]);
  photo(s, [13.938, 1.028, 4.921, 3.02]);
}

function slide23(p) {
  const s = newSlide(p);
  photo(s, [1.601, 5.212, 3.145, 2.575]);
  photo(s, [6.268, 5.212, 3.145, 2.575]);
  photo(s, [10.935, 5.212, 3.145, 2.575]);
  photo(s, [15.601, 5.212, 3.145, 2.575]);
  txt(s, [1.49, 9.011, 3.399, 0.855], T28, 16, BODY, MD, { line: 1.5 });
  txt(s, [1.49, 8.405, 3.133, 0.404], 'Great Team Work One', 18, BLUE, XB);
  txt(s, [6.268, 9.011, 3.399, 0.855], T28, 16, BODY, MD, { line: 1.5 });
  txt(s, [6.268, 8.405, 3.133, 0.404], 'Great Team Work Two', 18, BLUE, XB);
  txt(s, [10.935, 9.011, 3.399, 0.855], T28, 16, BODY, MD, { line: 1.5 });
  txt(s, [10.935, 8.405, 3.133, 0.404], 'Great Team Work Three', 18, BLUE, XB);
  txt(s, [15.601, 9.011, 3.399, 0.855], T28, 16, BODY, MD, { line: 1.5 });
  txt(s, [15.601, 8.405, 3.133, 0.404], 'Great Team Work Four', 18, BLUE, XB);
  box(s, [3.961, 5.212, 0.785, 0.785], { fill: BLUE });
  tag(s, [4.17, 5.458, 0.366, 0.293], WHITE);
  box(s, [8.616, 5.212, 0.785, 0.785], { fill: BLUE });
  tag(s, [8.826, 5.458, 0.366, 0.293], WHITE);
  box(s, [13.345, 5.212, 0.785, 0.785], { fill: BLUE });
  tag(s, [13.555, 5.458, 0.366, 0.293], WHITE);
  box(s, [17.963, 5.212, 0.785, 0.785], { fill: BLUE });
  tag(s, [18.173, 5.458, 0.366, 0.293], WHITE);
  rich(s, [1.57, 1.383, 7.41, 1.582], [
    ['Our Team Work Avast ', 44, INK, XB],
    ['Business Multipurpose', 44, BLUE, XB],
  ]);
  txt(s, [1.57, 3.333, 7.668, 0.855], T26, 16, BODY, MD, { line: 1.5 });
  txt(s, [10.369, 2.562, 8.974, 1.259], T27, 16, BODY, MD, { line: 1.5 });
  txt(s, [10.369, 1.67, 8.652, 0.505], T21, 24, INK, XB);
}

function slide24(p) {
  const s = newSlide(p, PAGE);
  poly(s, [1.093, 3.875, 3.796, 4.01], BLUE, 'spike');
  poly(s, [3.005, 3.906, 1.884, 3.979], BLACK, 'spikeShade', { alpha: 80 });
  poly(s, [3.896, 5.796, 3.796, 2.089], DARK, 'spike');
  poly(s, [5.808, 5.812, 1.884, 2.073], BLACK, 'spikeShade', { alpha: 80 });
  poly(s, [6.7, 5.019, 3.796, 2.866], BLUE, 'spike');
  poly(s, [8.612, 5.041, 1.884, 2.844], BLACK, 'spikeShade', { alpha: 80 });
  poly(s, [9.504, 6.518, 3.796, 1.367], DARK, 'spike');
  poly(s, [11.416, 6.528, 1.884, 1.357], BLACK, 'spikeShade', { alpha: 80 });
  poly(s, [12.307, 4.257, 3.796, 3.629], BLUE, 'spike');
  poly(s, [14.22, 4.284, 1.884, 3.601], BLACK, 'spikeShade', { alpha: 80 });
  poly(s, [15.111, 4.692, 3.796, 3.193], DARK, 'spike');
  poly(s, [17.023, 4.717, 1.884, 3.168], BLACK, 'spikeShade', { alpha: 80 });
  box(s, [2.122, 2.883, 1.737, 0.586], { fill: BLUE });
  txt(s, [2.122, 2.883, 1.737, 0.586], '82%', 20, WHITE, XB, { align: 'center', valign: 'middle' });
  box(s, [2.654, 3.464, 0.674, 0.196], { shape: 'triangle', fill: BLUE, rotate: 180 });
  box(s, [4.876, 4.717, 1.737, 0.586], { fill: DARK });
  txt(s, [4.876, 4.717, 1.737, 0.586], '46%', 20, WHITE, XB, { align: 'center', valign: 'middle' });
  box(s, [5.407, 5.298, 0.674, 0.196], { shape: 'triangle', fill: DARK, rotate: 180 });
  box(s, [7.73, 3.941, 1.737, 0.586], { fill: BLUE });
  txt(s, [7.73, 3.941, 1.737, 0.586], '70%', 20, WHITE, XB, { align: 'center', valign: 'middle' });
  box(s, [8.261, 4.521, 0.674, 0.196], { shape: 'triangle', fill: BLUE, rotate: 180 });
  box(s, [10.52, 5.493, 1.737, 0.586], { fill: DARK });
  txt(s, [10.52, 5.493, 1.737, 0.586], '31%', 20, WHITE, XB, { align: 'center', valign: 'middle' });
  box(s, [11.051, 6.074, 0.674, 0.196], { shape: 'triangle', fill: DARK, rotate: 180 });
  box(s, [13.337, 3.204, 1.737, 0.586], { fill: BLUE });
  txt(s, [13.337, 3.204, 1.737, 0.586], '75%', 20, WHITE, XB, { align: 'center', valign: 'middle' });
  box(s, [13.868, 3.785, 0.674, 0.196], { shape: 'triangle', fill: BLUE, rotate: 180 });
  box(s, [16.104, 3.712, 1.737, 0.586], { fill: DARK });
  txt(s, [16.104, 3.712, 1.737, 0.586], '72%', 20, WHITE, XB, { align: 'center', valign: 'middle' });
  box(s, [16.635, 4.293, 0.674, 0.196], { shape: 'triangle', fill: DARK, rotate: 180 });
  txt(s, [2.421, 8.321, 1.14, 0.597], '2017', 24, DARK, XB, { align: 'center', wrap: 0 });
  txt(s, [5.224, 8.327, 1.169, 0.597], '2018', 24, DARK, XB, { align: 'center', wrap: 0 });
  txt(s, [8.023, 8.321, 1.151, 0.597], '2019', 24, DARK, XB, { align: 'center', wrap: 0 });
  txt(s, [10.783, 8.327, 1.266, 0.597], '2020', 24, DARK, XB, { align: 'center', wrap: 0 });
  txt(s, [13.666, 8.321, 1.153, 0.597], '2021', 24, DARK, XB, { align: 'center', wrap: 0 });
  txt(s, [16.447, 8.327, 1.225, 0.597], '2022', 24, DARK, XB, { align: 'center', wrap: 0 });
  txt(s, [3.007, 9.45, 13.987, 0.871], T29, 16, BODY, SB, { align: 'center', line: 1.5 });
  txt(s, [3.715, 1.344, 12.569, 1.111], 'Stastitics diagram', 60, INK, XB, { align: 'center' });
  txt(s, [7.833, 0.586, 4.333, 0.505], 'Business Multipurpose', 24, BLUE, SB, { align: 'center' });
  chrome(s, { logo: 0.775, bars: INK, mark: INK });
}

function slide25(p) {
  const s = newSlide(p, PAGE);
  rule(s, [4.551, 5.434, 0, 1.13], { color: GRAY, width: 2.25, flipH: 1, flipV: 1 });
  rule(s, [11.816, 5.434, 0, 1.13], { color: GRAY, width: 2.25, flipH: 1, flipV: 1 });
  rule(s, [8.184, 7.215, 0, 1.13], { color: GRAY, width: 2.25, flipV: 1 });
  rule(s, [15.449, 7.215, 0, 1.13], { color: GRAY, width: 2.25, flipV: 1 });
  rule(s, [0.918, 6.89, 18.164, 0], { color: SLATE, width: 6 });
  box(s, [7.858, 6.564, 0.651, 0.651], { shape: 'ellipse', fill: WHITE, stroke: SLATE, weight: 2.25 });
  poly(s, [10.605, 3.668, 2.422, 1.766], BLUE, 'pennantDown');
  poly(s, [3.34, 3.668, 2.422, 1.766], BLUE, 'pennantDown');
  poly(s, [14.238, 8.346, 2.422, 1.766], INK, 'pennantUp');
  poly(s, [6.973, 8.346, 2.422, 1.766], INK, 'pennantUp');
  box(s, [0.677, 6.649, 0.482, 0.482], { shape: 'ellipse', fill: WHITE, stroke: SLATE, weight: 1 });
  box(s, [0.593, 6.564, 0.651, 0.651], { shape: 'ellipse', fill: WHITE, stroke: SLATE, weight: 2.25 });
  box(s, [4.31, 6.649, 0.482, 0.482], { shape: 'ellipse', fill: WHITE, stroke: SLATE, weight: 1 });
  box(s, [4.225, 6.564, 0.651, 0.651], { shape: 'ellipse', fill: WHITE, stroke: SLATE, weight: 2.25 });
  box(s, [7.943, 6.649, 0.482, 0.482], { shape: 'ellipse', fill: INK });
  box(s, [15.208, 6.649, 0.482, 0.482], { shape: 'ellipse', fill: WHITE, stroke: SLATE, weight: 1 });
  box(s, [15.124, 6.564, 0.651, 0.651], { shape: 'ellipse', fill: WHITE, stroke: SLATE, weight: 2.25 });
  box(s, [11.576, 6.649, 0.482, 0.482], { shape: 'ellipse', fill: WHITE, stroke: SLATE, weight: 1 });
  box(s, [11.491, 6.564, 0.651, 0.651], { shape: 'ellipse', fill: WHITE, stroke: SLATE, weight: 2.25 });
  box(s, [18.841, 6.649, 0.482, 0.482], { shape: 'ellipse', fill: WHITE, stroke: SLATE, weight: 1 });
  box(s, [18.756, 6.564, 0.651, 0.651], { shape: 'ellipse', fill: WHITE, stroke: SLATE, weight: 2.25 });
  icon(s, [3.924, 3.696, 1.254, 1.254], WHITE, 'runner');
  icon(s, [7.556, 8.79, 1.254, 1.254], WHITE, 'chat');
  icon(s, [11.189, 3.755, 1.254, 1.254], WHITE, 'target');
  icon(s, [14.822, 8.79, 1.254, 1.254], WHITE, 'coins');
  box(s, [0.677, 6.649, 0.482, 0.482], { shape: 'ellipse', fill: INK });
  box(s, [4.292, 6.649, 0.482, 0.482], { shape: 'ellipse', fill: BLUE });
  box(s, [11.56, 6.649, 0.482, 0.482], { shape: 'ellipse', fill: BLUE });
  box(s, [15.208, 6.649, 0.482, 0.482], { shape: 'ellipse', fill: INK });
  box(s, [18.841, 6.649, 0.482, 0.482], { shape: 'ellipse', fill: BLUE });
  txt(s, [2.412, 8.455, 4.194, 0.871], T30, 16, BODY, MD, { align: 'center', line: 1.5 });
  txt(s, [3.16, 7.789, 2.699, 0.505], 'Section One', 24, DARK, XB, { align: 'center' });
  txt(s, [6.086, 4.823, 4.194, 0.871], T30, 16, BODY, MD, { align: 'center', line: 1.5 });
  txt(s, [6.827, 4.265, 2.699, 0.505], 'Section Two', 24, DARK, XB, { align: 'center' });
  txt(s, [9.72, 8.455, 4.194, 0.871], T30, 16, BODY, MD, { align: 'center', line: 1.5 });
  txt(s, [10.467, 7.789, 2.699, 0.505], 'Section Three', 24, DARK, XB, { align: 'center' });
  txt(s, [13.394, 4.823, 4.194, 0.871], T30, 16, BODY, MD, { align: 'center', line: 1.5 });
  txt(s, [14.135, 4.265, 2.699, 0.505], 'Section Four', 24, DARK, XB, { align: 'center' });
  txt(s, [3.715, 1.344, 12.569, 1.111], 'Horizontal process diagram ', 60, INK, XB, { align: 'center' });
  txt(s, [7.833, 0.586, 4.333, 0.505], 'Business Multipurpose', 24, BLUE, SB, { align: 'center' });
  chrome(s, { logo: 0.775, bars: INK, mark: INK });
}

function slide26(p) {
  const s = newSlide(p, PAGE);
  icon(s, [15.551, 6.492, 1.5, 1.5], BLUE, 'people');
  icon(s, [11.348, 5.607, 1.5, 1.5], INK, 'puzzle');
  icon(s, [7.152, 6.492, 1.5, 1.5], BLUE, 'bulb');
  poly(s, [1.398, 4.309, 4.618, 2.309], BLUE, 'chevron');
  poly(s, [1.766, 4.677, 3.881, 1.94], BLUE, 'chevronIn');
  poly(s, [9.789, 4.309, 4.618, 2.309], BLUE, 'chevron');
  poly(s, [10.157, 4.677, 3.881, 1.94], BLUE, 'chevronIn');
  poly(s, [5.593, 6.833, 4.618, 2.309], INK, 'chevron', { rotate: 180 });
  poly(s, [5.962, 6.833, 3.881, 1.94], INK, 'chevronIn', { rotate: 180 });
  poly(s, [13.985, 6.833, 4.618, 2.309], INK, 'chevron', { rotate: 180 });
  poly(s, [14.353, 6.833, 3.881, 1.94], INK, 'chevronIn', { rotate: 180 });
  icon(s, [2.957, 5.607, 1.5, 1.5], INK, 'flame');
  txt(s, [1.619, 8.271, 4.194, 0.871], T30, 16, BODY, MD, { align: 'center', line: 1.5 });
  txt(s, [2.366, 7.606, 2.699, 0.505], 'Section One', 24, DARK, XB, { align: 'center' });
  txt(s, [5.812, 4.951, 4.194, 0.871], T30, 16, BODY, MD, { align: 'center', line: 1.5 });
  txt(s, [6.553, 4.393, 2.699, 0.505], 'Section Two', 24, DARK, XB, { align: 'center' });
  txt(s, [10.165, 8.271, 4.194, 0.871], T30, 16, BODY, MD, { align: 'center', line: 1.5 });
  txt(s, [10.912, 7.606, 2.699, 0.505], 'Section Three', 24, DARK, XB, { align: 'center' });
  txt(s, [14.082, 4.954, 4.194, 0.871], T30, 16, BODY, MD, { align: 'center', line: 1.5 });
  txt(s, [14.824, 4.397, 2.699, 0.505], 'Section Four', 24, DARK, XB, { align: 'center' });
  txt(s, [3.715, 2.076, 12.569, 1.111], 'Process with half frames', 60, INK, XB, { align: 'center' });
  txt(s, [7.833, 1.318, 4.333, 0.505], 'Business Multipurpose', 24, BLUE, SB, { align: 'center' });
  chrome(s, { logo: 0.775, bars: INK, mark: INK });
}

function slide27(p) {
  const s = newSlide(p, PAGE);
  poly(s, [7.37, 5.758, 2.267, 5.492], INK, 'arrow1');
  poly(s, [10.933, 4.279, 4.312, 6.971], INK, 'arrow2');
  poly(s, [9.828, 6.677, 2.267, 4.573], BLUE, 'arrow3');
  poly(s, [4.203, 5.19, 4.312, 6.06], BLUE, 'arrow4');
  txt(s, [3.715, 1.662, 12.569, 2.12], 'Four curved up arrows infographic', 60, INK, XB, { align: 'center' });
  txt(s, [7.833, 0.904, 4.333, 0.505], 'Business Multipurpose', 24, BLUE, SB, { align: 'center' });
  chrome(s, { logo: 0.775, bars: INK, mark: INK });
  box(s, [1.092, 4.247, 0.785, 0.785], { fill: BLUE });
  txt(s, [2.087, 4.945, 4.161, 1.259], T15, 16, BODY, MD, { line: 1.5 });
  txt(s, [2.087, 4.16, 3.46, 0.505], 'Our Section One', 24, INK, XB);
  tag(s, [1.302, 4.493, 0.366, 0.293], WHITE);
  box(s, [1.092, 7.423, 0.785, 0.785], { fill: BLUE });
  txt(s, [2.087, 8.121, 4.161, 1.259], T15, 16, BODY, MD, { line: 1.5 });
  txt(s, [2.087, 7.336, 3.46, 0.505], 'Our Section Three', 24, INK, XB);
  tag(s, [1.302, 7.669, 0.366, 0.293], WHITE);
  box(s, [13.751, 4.247, 0.785, 0.785], { fill: BLUE });
  txt(s, [14.746, 4.945, 4.161, 1.259], T15, 16, BODY, MD, { line: 1.5 });
  txt(s, [14.746, 4.16, 3.46, 0.505], 'Our Section Two', 24, INK, XB);
  tag(s, [13.961, 4.493, 0.366, 0.293], WHITE);
  box(s, [13.751, 7.423, 0.785, 0.785], { fill: BLUE });
  txt(s, [14.746, 8.121, 4.161, 1.259], T15, 16, BODY, MD, { line: 1.5 });
  txt(s, [14.746, 7.336, 3.46, 0.505], 'Our Section Four', 24, INK, XB);
  tag(s, [13.961, 7.669, 0.366, 0.293], WHITE);
}

function slide28(p) {
  const s = newSlide(p);
  box(s, [0, 0, 2.587, 11.25], { fill: BLUE });
  rich(s, [11.368, 2.541, 7.254, 2.827], [
    ['The best way to predict ', 54, INK, XB],
    ['the future is to create it.', 54, BLUE, XB],
  ]);
  rich(s, [11.368, 6.088, 7.528, 2.875], [
    [T31, 16, BODY, MD, 1],
    ['', 16, INK, MD, 1],
    [T32, 16, BODY, MD],
  ], { line: 1.5 });
  txt(s, [11.368, 1.669, 4.333, 0.505], 'Business Multipurpose', 24, BLUE, SB);
  photo(s, [1.378, 6.27, 4.124, 4.98]);
  photo(s, [6.169, 1.222, 3.497, 8.292]);
  photo(s, [1.378, 1.222, 4.124, 4.292]);
}

function slide29(p) {
  const s = newSlide(p);
  box(s, [9.46, 0, 4.064, 3.032], { fill: BLUE });
  photo(s, [9.46, 0.944, 9.571, 4.015]);
  rich(s, [1.114, 2.541, 7.254, 2.827], [
    ['The best way to predict ', 54, INK, XB],
    ['the future is to create it.', 54, BLUE, XB],
  ]);
  rich(s, [1.114, 6.182, 10.849, 2.875], [
    [T33, 16, BODY, MD, 1],
    ['', 16, INK, MD, 1],
    [T34, 16, BODY, MD],
  ], { line: 1.5 });
  txt(s, [1.114, 1.669, 4.333, 0.505], 'Business Multipurpose', 24, BLUE, SB);
  box(s, [13.02, 2.769, 6.98, 8.481], { fill: BLUE });
  icon(s, [13.933, 3.694, 0.631, 0.631], WHITE, 'gem');
  txt(s, [14.751, 4.325, 4.653, 0.855], T18, 16, WHITE, MD, { line: 1.5 });
  txt(s, [14.751, 3.681, 2.64, 0.438], 'Our Services One', 20, WHITE, XB);
  icon(s, [13.933, 8.731, 0.631, 0.631], WHITE, 'gem');
  txt(s, [14.751, 9.363, 4.653, 0.855], T18, 16, WHITE, MD, { line: 1.5 });
  txt(s, [14.751, 8.719, 3.28, 0.438], 'Our Services Three', 20, WHITE, XB);
  icon(s, [13.933, 6.213, 0.631, 0.631], WHITE, 'gem');
  txt(s, [14.751, 6.844, 4.653, 0.855], T18, 16, WHITE, MD, { line: 1.5 });
  txt(s, [14.751, 6.2, 2.64, 0.438], 'Our Services Two', 20, WHITE, XB);
}

function slide30(p) {
  const s = newSlide(p);
  photo(s, [0, 6.54, 2.997, 3.871]);
  photo(s, [3.492, 6.54, 4.524, 3.871]);
  photo(s, [8.511, 6.54, 4.524, 3.871]);
  box(s, [0, 7.202, 2.997, 3.215], { fade: BLUE });
  txt(s, [0.239, 9.1, 2.111, 0.404], 'Portfolio One', 18, WHITE, XB);
  txt(s, [0.239, 9.504, 2.111, 0.408], '+ Lorem ipsum dolor', 14, WHITE, MD, { line: 1.5 });
  box(s, [3.492, 7.202, 4.524, 3.215], { fade: BLUE });
  txt(s, [3.731, 9.1, 2.111, 0.404], 'Portfolio Two', 18, WHITE, XB);
  txt(s, [3.731, 9.504, 2.111, 0.408], '+ Lorem ipsum dolor', 14, WHITE, MD, { line: 1.5 });
  box(s, [8.511, 7.202, 4.524, 3.215], { fade: BLUE });
  txt(s, [8.75, 9.1, 2.111, 0.404], 'Portfolio Three', 18, WHITE, XB);
  txt(s, [8.75, 9.504, 2.111, 0.408], '+ Lorem ipsum dolor', 14, WHITE, MD, { line: 1.5 });
  icon(s, [15.025, 0.725, 0.631, 0.631], BLUE, 'gem');
  txt(s, [14.836, 2.308, 4.012, 0.855], T12, 16, BODY, MD, { line: 1.5 });
  txt(s, [14.836, 1.664, 2.64, 0.438], 'Our Section One', 20, INK, XB);
  txt(s, [14.836, 5.625, 4.012, 0.855], T12, 16, BODY, MD, { line: 1.5 });
  txt(s, [14.836, 4.981, 2.64, 0.438], 'Our Section Two', 20, INK, XB);
  icon(s, [15.025, 4.042, 0.631, 0.631], BLUE, 'till');
  icon(s, [15.025, 7.359, 0.631, 0.631], BLUE, 'gem');
  txt(s, [14.836, 8.942, 4.012, 0.855], T12, 16, BODY, MD, { line: 1.5 });
  txt(s, [14.836, 8.298, 2.925, 0.438], 'Our Section Three', 20, INK, XB);
  rich(s, [0.987, 1.944, 12.047, 2.12], [
    ['Our Gallery Portfolio Avast   ', 60, INK, XB],
    ['Business Multipurpose', 60, BLUE, XB],
  ]);
  txt(s, [0.987, 0.974, 4.333, 0.505], 'Business Multipurpose', 24, BLUE, SB);
  txt(s, [0.987, 4.53, 12.047, 0.855], T27, 16, BODY, MD, { line: 1.5 });
}

function slide31(p) {
  const s = newSlide(p);
  photo(s, [11.762, 0.006, 3.746, 3.028]);
  photo(s, [15.825, 0, 3.746, 4.27]);
  photo(s, [15.825, 4.556, 3.746, 3.381]);
  photo(s, [11.762, 3.317, 3.746, 3.381]);
  photo(s, [11.762, 6.98, 3.746, 4.27]);
  photo(s, [15.825, 8.222, 3.746, 3.028]);
  box(s, [11.762, 9.943, 3.746, 1.317], { fill: BLUE });
  txt(s, [12.096, 10.147, 2.111, 0.404], 'Portfolio Three', 18, WHITE, XB);
  txt(s, [12.096, 10.551, 2.111, 0.408], '+ Lorem ipsum dolor', 14, WHITE, MD, { line: 1.5 });
  box(s, [15.825, 9.943, 3.746, 1.317], { fill: BLUE });
  txt(s, [16.16, 10.147, 2.111, 0.404], 'Portfolio Six', 18, WHITE, XB);
  txt(s, [16.16, 10.551, 2.111, 0.408], '+ Lorem ipsum dolor', 14, WHITE, MD, { line: 1.5 });
  box(s, [15.825, 6.625, 3.746, 1.317], { fill: BLUE });
  txt(s, [16.16, 6.829, 2.111, 0.404], 'Portfolio Five', 18, WHITE, XB);
  txt(s, [16.16, 7.233, 2.111, 0.408], '+ Lorem ipsum dolor', 14, WHITE, MD, { line: 1.5 });
  box(s, [15.825, 2.956, 3.746, 1.317], { fill: BLUE });
  txt(s, [16.16, 3.16, 2.111, 0.404], 'Portfolio Four', 18, WHITE, XB);
  txt(s, [16.16, 3.564, 2.111, 0.408], '+ Lorem ipsum dolor', 14, WHITE, MD, { line: 1.5 });
  box(s, [11.762, 5.386, 3.746, 1.317], { fill: BLUE });
  txt(s, [12.096, 5.59, 2.111, 0.404], 'Portfolio Two', 18, WHITE, XB);
  txt(s, [12.096, 5.994, 2.111, 0.408], '+ Lorem ipsum dolor', 14, WHITE, MD, { line: 1.5 });
  box(s, [11.762, 1.72, 3.746, 1.317], { fill: BLUE });
  txt(s, [12.096, 1.924, 2.111, 0.404], 'Portfolio One', 18, WHITE, XB);
  txt(s, [12.096, 2.328, 2.111, 0.408], '+ Lorem ipsum dolor', 14, WHITE, MD, { line: 1.5 });
  rich(s, [1.258, 2.139, 9.013, 1.717], [
    ['Our Gallery Portfolio Avast ', 48, INK, XB],
    ['Business Multipurpose', 48, BLUE, XB],
  ]);
  txt(s, [1.258, 1.169, 4.333, 0.505], 'Business Multipurpose', 24, BLUE, SB);
  rich(s, [1.258, 4.556, 9.203, 2.471], [
    [T35, 16, BODY, MD, 1],
    ['', 16, INK, MD, 1],
    [T36, 16, BODY, MD],
  ], { line: 1.5 });
  txt(s, [1.258, 8.572, 9.203, 1.259], T35, 16, BODY, MD, { line: 1.5 });
  txt(s, [1.258, 7.679, 8.652, 0.505], T21, 24, INK, XB);
}

function slide32(p) {
  const s = newSlide(p);
  box(s, [0, 8.076, 20, 3.174], { fill: BLUE });
  txt(s, [1.114, 2.334, 9.37, 1.447], 'Mockup Slides', 80, INK, XB);
  txt(s, [1.114, 4.362, 9.209, 1.259], T37, 16, BODY, MD, { line: 1.5 });
  txt(s, [1.114, 1.462, 4.333, 0.505], 'Business Multipurpose', 24, BLUE, SB);
  box(s, [1.252, 6.176, 2.022, 0.633], { fill: BLUE });
  txt(s, [1.252, 6.176, 2.022, 0.633], 'Read More', 18, WHITE, XB, { align: 'center', valign: 'middle' });
  box(s, [4.043, 6.176, 2.022, 0.633], { fill: BLUE });
  txt(s, [4.043, 6.176, 2.022, 0.633], 'Learn More', 18, WHITE, XB, { align: 'center', valign: 'middle' });
  txt(s, [1.252, 9.413, 8.974, 0.855], T38, 16, WHITE, MD, { line: 1.5 });
  txt(s, [1.252, 8.653, 8.652, 0.505], T21, 24, WHITE, XB);
  photo(s, [12.273, 1.697, 5.939, 7.848]);
  icon(s, [11.851, 1.514, 6.513, 8.221], INK, 'tablet');
}

function slide33(p) {
  const s = newSlide(p, BLUE);
  txt(s, [10.51, 2.492, 8.928, 1.919], 'Get In Touch Avast   Business Multipurpose', 54, WHITE, XB);
  txt(s, [10.51, 1.522, 4.333, 0.505], 'Business Multipurpose', 24, WHITE, SB);
  box(s, [10.643, 7.174, 1.666, 0.535], { fill: WHITE });
  txt(s, [10.643, 7.174, 1.666, 0.535], '  Contact ', 18, BLUE, XB, { valign: 'middle' });
  rule(s, [12.528, 7.441, 0.59, 0], { color: 'A6A6A6', width: 3 });
  txt(s, [10.643, 8.138, 0.416, 0.438], 'T', 20, WHITE, XB);
  txt(s, [10.999, 8.167, 2.591, 0.37], ':  +0123 4567 8901', 16, WHITE, MD);
  txt(s, [10.643, 8.637, 0.416, 0.438], 'F', 20, WHITE, XB);
  txt(s, [10.999, 8.665, 2.591, 0.37], ':  +0011  2233 4455', 16, WHITE, MD);
  txt(s, [10.643, 9.157, 0.416, 0.438], 'E', 20, WHITE, XB);
  txt(s, [10.999, 9.186, 2.966, 0.37], ':  business@example.com', 16, WHITE, MD);
  box(s, [15.409, 7.174, 1.666, 0.535], { fill: WHITE });
  txt(s, [15.409, 7.174, 1.666, 0.535], '  Address ', 18, BLUE, XB, { valign: 'middle' });
  rule(s, [17.294, 7.441, 0.59, 0], { color: 'A6A6A6', width: 3 });
  txt(s, [15.443, 8.099, 3.656, 0.37], 'Address Line One Lorem Ipsum', 16, WHITE, MD);
  txt(s, [15.443, 9.133, 2.153, 0.37], 'Asia, Indonesian', 16, WHITE, MD);
  txt(s, [15.443, 8.593, 3.656, 0.37], 'Address Line Two Lorem', 16, WHITE, MD);
  txt(s, [10.51, 4.922, 8.589, 1.259], T27, 16, LIGHT, MD, { line: 1.5 });
  photo(s, [0, 0, 9.111, 11.25], SHEER);
}

function slide34(p) {
  const s = newSlide(p, BLUE);
  txt(s, [5.903, 3.343, 8.194, 2.895], 'Thanks', 166, WHITE, XB, { align: 'center', wrap: 0 });
  txt(s, [6.889, 6.285, 6.223, 0.707], 'Business Multipurpose', 36, WHITE, SB, { align: 'center' });
  icon(s, [9.063, 1.834, 1.875, 1.875], WHITE, 'target');
  chrome(s, { logo: 0.775, logoColor: WHITE, bars: WHITE, foot: WHITE, mark: WHITE });
  txt(s, [6.081, 7.494, 7.926, 0.855], T1, 16, WHITE, MD, { align: 'center', line: 1.5 });
}

/* -------------------------------------------------------------------- build */
const SLIDES = [
  slide01, slide02, slide03, slide04, slide05, slide06,
  slide07, slide08, slide09, slide10, slide11, slide12,
  slide13, slide14, slide15, slide16, slide17, slide18,
  slide19, slide20, slide21, slide22, slide23, slide24,
  slide25, slide26, slide27, slide28, slide29, slide30,
  slide31, slide32, slide33, slide34,
];

function build() {
  const p = new PptxGenJS();
  p.defineLayout({ name: 'WIDE20', width: 20, height: 11.25 });
  p.layout = 'WIDE20';
  p.author = 'Avast';
  p.title = 'Business Multipurpose';
  SLIDES.forEach((fn) => fn(p));
  return p.writeFile({
    fileName: path.join(__dirname, '0026fe40-e808-4023-8e07-b59f8424918c_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f));
