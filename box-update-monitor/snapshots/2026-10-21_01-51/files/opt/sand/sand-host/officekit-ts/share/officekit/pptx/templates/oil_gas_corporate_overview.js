/**
 * Recreation of "Oil & Gas — HydraX" (16 slides, 13.333 x 7.5 in) with pptxgenjs.
 *
 * Everything is expressed as plain shape/text calls plus small data tables:
 *   - PAL / FONT           : theme palette and fonts
 *   - generic helpers      : para(), snipTag(), txt(), iconMark(), ...
 *   - chrome()             : the navigation bar + page number repeated on every slide
 *   - SLIDES[]             : one builder function per slide, in deck order
 *
 * Raster images from the original deck are replaced by programmatic placeholders.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const PAL = {
  dark: '0E2E34',       // accent1
  yellow: 'FBD508',     // accent2
  white: 'FFFFFF',
  teal: '4FBDBA',       // accent6
  orange: 'E85C0D',     // accent5
  peach: 'FCDECD',      // accent5 @ 20%
  paleYellow: 'FDE66B', // accent2 lumMod 60 / lumOff 40
  black: '000000',
  grey: '808080',       // tx1 lumMod 50 / lumOff 50
  greyDark: '404040',   // tx1 lumMod 75 / lumOff 25
  gridGrey: 'AFABAB',   // bg2 lumMod 75
  lightGrey: 'E7E6E6',  // bg2
  mapGrey: 'EFEFEF',
  road: '595959',
  road2: '767171',
  photo: 'E4E4E4',      // stand-in fill for removed photographs
};

const FONT = 'Poppins';
const FONT_B = 'Poppins SemiBold';

const SHADOW = { type: 'outer', blur: 15, offset: 3, angle: 45, color: '000000', opacity: 0.15 };
const SHADOW_SOFT = { type: 'outer', blur: 23, offset: 3, angle: 45, color: '000000', opacity: 0.25 };
const SHADOW_DEEP = { type: 'outer', blur: 23, offset: 3, angle: 45, color: '000000', opacity: 0.35 };
const SHADOW_CARD = { type: 'outer', blur: 15, offset: 3, angle: 135, color: '000000', opacity: 0.1 };

/* ---------------------------------------------------------------- helpers */

/** Text box. The deck's text boxes are top-anchored and use the Poppins family. */
function txt(slide, text, o) {
  slide.addText(text, Object.assign({ fontFace: FONT, color: PAL.dark, valign: 'top' }, o));
}

/**
 * Parallelogram. `adj` is the OOXML adjust value (skew = min(w,h) * adj);
 * 0.25 is the preset default, so those keep the native prstGeom shape.
 */
function para(slide, o) {
  const adj = o.adj === undefined ? 0.25 : o.adj;
  const opts = { x: o.x, y: o.y, w: o.w, h: o.h };
  if (o.fill) opts.fill = o.fill;
  if (o.shadow) opts.shadow = o.shadow;
  if (adj === 0.25) return slide.addShape('parallelogram', opts);
  const s = Math.min(o.w, o.h) * adj;
  opts.points = [{ x: 0, y: o.h }, { x: s, y: 0 }, { x: o.w, y: 0 }, { x: o.w - s, y: o.h }, { close: true }];
  return slide.addShape('custGeom', opts);
}

/** Closed polygon given as normalised [0..1] points inside the x/y/w/h box. */
function poly(slide, x, y, w, h, pts, o) {
  slide.addShape('custGeom', Object.assign({
    x: x, y: y, w: w, h: h,
    points: pts.map(function (p) { return { x: p[0] * w, y: p[1] * h }; }).concat([{ close: true }]),
  }, o));
}

/** Open polyline (a stroked curve) given as normalised [0..1] points. */
function polyline(slide, x, y, w, h, pts, line) {
  slide.addShape('custGeom', {
    x: x, y: y, w: w, h: h, line: line,
    points: pts.map(function (p) { return { x: p[0] * w, y: p[1] * h }; }),
  });
}

/** Curve described by evenly spaced y-samples across the box (used by slide 8). */
function curve(slide, x, y, w, h, ys, line) {
  polyline(slide, x, y, w, h, ys.map(function (v, i) { return [i / (ys.length - 1), v]; }), line);
}

/** Rectangle with the top-right / bottom-left corners snipped (the "About." tag). */
function snipTag(slide, x, y, w, h, color) {
  const s = Math.min(w, h) * 0.25178;
  slide.addShape('custGeom', {
    x: x, y: y, w: w, h: h, line: { color: color, width: 1.5 },
    points: [{ x: 0, y: 0 }, { x: w - s, y: 0 }, { x: w, y: s }, { x: w, y: h },
             { x: s, y: h }, { x: 0, y: h - s }, { close: true }],
  });
}

/** Small square icon placeholder standing in for an original vector/raster icon. */
function iconMark(slide, x, y, size, color) {
  slide.addShape('roundRect', {
    x: x + size * 0.05, y: y + size * 0.05, w: size * 0.9, h: size * 0.9,
    rectRadius: size * 0.16, line: { color: color, width: Math.max(0.75, size * 2.6) },
  });
  slide.addShape('ellipse', { x: x + size * 0.33, y: y + size * 0.33, w: size * 0.34, h: size * 0.34, fill: { color: color } });
}

/**
 * Placeholder for the deck's only photograph (an open laptop). Raster images
 * are never embedded; the silhouette is redrawn as two flat grey quads.
 */
const LAPTOP_SCREEN = [[0.935, 0.00], [0.945, 0.63], [0.487, 0.63], [0.490, 0.28]];
const LAPTOP_BASE = [[0.484, 0.63], [0.893, 0.63], [0.293, 0.95], [0.118, 0.95]];

function photo(slide, x, y, w, h, label) {
  poly(slide, x, y, w, h, LAPTOP_SCREEN, { fill: { color: PAL.photo }, line: { color: '6E7679', width: 3 } });
  poly(slide, x, y, w, h, LAPTOP_BASE, { fill: { color: 'A9AEB0' }, line: { color: '6E7679', width: 1.5 } });
  txt(slide, label || '[image]', { x: x + w * 0.5, y: y + h * 0.4, w: w * 0.4, h: 0.3, fontSize: 11, color: '4A5153', align: 'center', valign: 'middle' });
}

/** Dashed-outline square enclosing a right arrow — the deck's "go" badge. */
function arrowBadge(slide, x, y, size, color) {
  const c = color || PAL.dark;
  slide.addShape('rect', { x: x, y: y, w: size, h: size, line: { color: c, width: 1, dashType: 'dash' } });
  slide.addShape('line', {
    x: x + size * 0.2, y: y + size / 2, w: size * 0.6, h: 0,
    line: { color: c, width: 1.25, endArrowType: 'triangle' },
  });
}

/** Yellow call-to-action pill: parallelogram + label + arrow badge. */
function ctaButton(slide, o) {
  para(slide, { x: o.x, y: o.y, w: o.w, h: o.h, adj: o.adj, fill: { color: PAL.yellow }, shadow: SHADOW });
  txt(slide, o.label, { x: o.tx, y: o.ty, w: o.tw, h: o.th, fontFace: FONT_B, fontSize: o.size, color: PAL.dark });
  arrowBadge(slide, o.ax, o.ay, o.as);
}

/* ------------------------------------------------------- repeated chrome */

const NAV_LINKS = [
  { text: 'Services.', x: 9.396 },
  { text: 'Contact.', x: 10.612 },
];

/** Top navigation ("HydraX ... Services. Contact. About.") plus the page number. */
function chrome(slide, page, o) {
  const nav = (o && o.nav) || PAL.black;
  const brand = (o && o.brand) || nav;
  const pageColor = (o && o.page) || nav;
  iconMark(slide, 0.335, 0.243, 0.232, PAL.yellow);
  txt(slide, 'HydraX', { x: 0.567, y: 0.191, w: 0.905, h: 0.337, fontSize: 14, color: brand });
  NAV_LINKS.forEach(function (l) {
    txt(slide, l.text, { x: l.x, y: 0.191, w: 1.073, h: 0.337, fontSize: 14, color: nav, align: 'center' });
  });
  snipTag(slide, 11.917, 0.191, 1.098, 0.337, PAL.yellow);
  txt(slide, 'About.', { x: 12.022, y: 0.191, w: 0.888, h: 0.337, fontSize: 14, color: nav, align: 'center' });
  txt(slide, 'Page ' + page, { x: 11.685, y: 6.973, w: 1.419, h: 0.337, fontSize: 14, color: pageColor, align: 'right' });
}

/* ------------------------------------------------------- slide 1 : cover */

const COVER_STRIPES = [
  [0.807, -0.016, 3.679, 2.754], [-1.415, 3.363, 3.679, 2.754],
  [2.455, 1.017, 3.868, 2.879], [0.0, 4.762, 3.679, 2.754],
];

function slide01(s) {
  s.addShape('rect', { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: PAL.dark, transparency: 50 } });
  COVER_STRIPES.forEach(function (r) {
    para(s, { x: r[0], y: r[1], w: r[2], h: r[3], adj: 0.66848, fill: { color: PAL.yellow } });
  });
  chrome(s, 1, { nav: PAL.white });

  // Dark banner: a "top corners snipped" rectangle rotated 90 deg -> snips on the right edge.
  poly(s, 0, 1.442, 11.685, 4.415,
    [[0, 0], [0.937, 0], [1, 0.167], [1, 0.833], [0.937, 1], [0, 1]], { fill: { color: PAL.dark } });

  txt(s, 'Oil & Gas', { x: 3.445, y: 1.803, w: 7.556, h: 2.036, fontFace: FONT_B, fontSize: 115, color: PAL.yellow });
  txt(s, '\u201CDriving Tomorrow\u2019s Energy!\u201D', { x: 3.451, y: 3.595, w: 7.556, h: 0.505, fontSize: 24, color: PAL.white });
  txt(s, 'Presentation Template', { x: 4.528, y: 4.33, w: 3.118, h: 0.404, fontFace: FONT_B, fontSize: 18, color: PAL.yellow });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit.', { x: 4.528, y: 4.768, w: 3.679, h: 0.505, fontSize: 12, color: PAL.white });

  para(s, { x: 8.403, y: 4.481, w: 3.679, h: 0.738, fill: { color: PAL.yellow }, shadow: SHADOW });
  para(s, { x: 0.814, y: -0.032, w: 3.679, h: 2.754, adj: 0.66848, fill: { color: PAL.yellow } });
  para(s, { x: -1.408, y: 3.347, w: 3.679, h: 2.754, adj: 0.66848, fill: { color: PAL.yellow } });

  s.addShape('rect', { x: 1.805, y: 2.562, w: 1.248, h: 1.248, fill: { color: PAL.dark }, line: { color: PAL.yellow, width: 6 }, shadow: SHADOW });
  iconMark(s, 2.114, 2.853, 0.681, PAL.yellow);
  para(s, { x: 0.007, y: 4.326, w: 4.138, h: 3.174, adj: 0.66848, fill: { color: PAL.yellow } });

  txt(s, 'Learn More', { x: 8.757, y: 4.565, w: 2.444, h: 0.572, fontFace: FONT_B, fontSize: 28, color: PAL.dark });
  arrowBadge(s, 11.289, 4.631, 0.439);
}

/* ------------------------------------------------ slide 2 : intro / split */

const INTRO_RUNS = [
  ['Welcome to Energy Flow, ', PAL.yellow],
  ['a future-driven company that ', PAL.black],
  ['explores, develops, and powers sustainable oil & gas solutions.', PAL.yellow],
  [' Our ', PAL.black],
  ['technology-driven approach ', PAL.yellow],
  ['PLACEHOLDER', PAL.black],
  ['from exploration to distribution, ', PAL.yellow],
  ['ensuring ', PAL.black],
  ['efficiency, safety, and sustainability ', PAL.yellow],
  ['with impact.', PAL.black],
];

function slide02(s) {
  poly(s, 0, 0, 7.292, 7.5, [[0, 0], [1, 0], [0.7188, 1], [0, 1]], { fill: { color: PAL.dark } });
  chrome(s, 2, { nav: PAL.black, brand: PAL.white });

  txt(s, INTRO_RUNS.map(function (r) { return { text: r[0], options: { color: r[1] } }; }),
    { x: 6.667, y: 0.872, w: 6.243, h: 5.756, fontSize: 28, align: 'right' });

  txt(s, '\u201CPowering Energy, Fuelling Progress\u201D',
    { x: 0.696, y: 0.872, w: 5.079, h: 2.322, fontFace: FONT_B, fontSize: 44, color: PAL.yellow });

  // Three yellow info tiles across the middle of the dark panel.
  s.addShape('rect', { x: -0.021, y: 3.631, w: 3.953, h: 1.349, fill: { color: PAL.yellow }, shadow: SHADOW });
  s.addShape('rect', { x: 4.108, y: 3.631, w: 1.868, h: 1.349, fill: { color: PAL.yellow }, shadow: SHADOW });
  s.addShape('rect', { x: 6.157, y: 3.631, w: 1.349, h: 1.349, fill: { color: PAL.yellow }, shadow: SHADOW });
  iconMark(s, 6.408, 3.882, 0.846, PAL.dark);
  iconMark(s, 0.272, 3.882, 0.847, PAL.dark);

  txt(s, 'Your Tagline', { x: 1.319, y: 3.824, w: 2.319, h: 0.37, fontFace: 'Oswald', fontSize: 20, lineSpacingMultiple: 0.8 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer.', { x: 1.319, y: 4.177, w: 2.319, h: 0.61, fontSize: 14, lineSpacingMultiple: 1.1 });
  txt(s, [{ text: 'Funfact!', options: { fontFace: FONT_B, fontSize: 20, breakLine: true } },
          { text: 'First oil well: 1859, Pennsylvania.', options: { fontSize: 14 } }],
    { x: 4.191, y: 3.733, w: 1.703, h: 1.144, align: 'center' });

  txt(s, 'Your Subtitle Here', { x: 0.702, y: 5.668, w: 3.118, h: 0.404, fontFace: FONT_B, fontSize: 18, color: PAL.white });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor.',
    { x: 0.702, y: 6.105, w: 3.679, h: 0.707, fontSize: 12, color: PAL.white });
}

/* --------------------------------------------- slide 3 : facts + diagonal */

function slide03(s) {
  para(s, { x: 2.127, y: 0, w: 7.269, h: 7.5, adj: 0.43227, fill: { color: PAL.dark, transparency: 50 } });
  para(s, { x: 6.667, y: 4.364, w: 4.54, h: 3.136, adj: 0.43227, fill: { color: PAL.dark, transparency: 50 } });
  chrome(s, 3);

  txt(s, 'Fuelling Progress, Together.', { x: 0.488, y: 0.711, w: 3.766, h: 2.322, fontFace: FONT_B, fontSize: 44 });
  txt(s, 'Your Subtitle Here', { x: 0.488, y: 3.744, w: 2.512, h: 0.404, fontFace: FONT_B, fontSize: 18 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ', { x: 0.488, y: 4.148, w: 2.512, h: 0.707, fontSize: 12, color: PAL.grey });

  ctaButton(s, { x: 0.566, y: 6.051, w: 3.92, h: 0.738, label: 'Explore Now', size: 28,
    tx: 0.92, ty: 6.134, tw: 2.684, th: 0.572, ax: 3.566, ay: 6.2, as: 0.439 });

  txt(s, '99%', { x: 9.378, y: 0.991, w: 2.368, h: 0.774, fontFace: FONT_B, fontSize: 40 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ', { x: 9.396, y: 1.662, w: 2.979, h: 0.572, fontSize: 14, color: PAL.grey });

  s.addShape('rect', { x: 10.146, y: 5.939, w: 0.872, h: 0.872, fill: { color: PAL.dark }, line: { color: PAL.white, width: 3 }, shadow: SHADOW });
  s.addShape('rect', { x: 4.157, y: 0.507, w: 0.872, h: 0.872, fill: { color: PAL.yellow }, line: { color: PAL.white, width: 3 }, shadow: SHADOW });
  iconMark(s, 4.355, 0.702, 0.476, PAL.dark);

  // Two fact cards that run off the right edge of the slide.
  s.addShape('rect', { x: 6.548, y: 2.588, w: 3.953, h: 1.349, fill: { color: PAL.dark }, shadow: SHADOW });
  iconMark(s, 6.842, 2.839, 0.847, PAL.yellow);
  txt(s, 'Oil Reserves', { x: 7.834, y: 2.782, w: 2.319, h: 0.379, fontFace: FONT_B, fontSize: 20, color: PAL.yellow, lineSpacingMultiple: 0.8 });
  txt(s, 'The world consumes 100M barrels of oil daily.', { x: 7.834, y: 3.134, w: 2.512, h: 0.61, fontSize: 14, color: PAL.yellow, lineSpacingMultiple: 1.1 });

  s.addShape('rect', { x: 10.725, y: 2.588, w: 5.294, h: 1.349, fill: { color: PAL.white }, shadow: SHADOW });
  txt(s, 'Energy Fact', { x: 12.065, y: 2.782, w: 2.319, h: 0.381, fontFace: FONT_B, fontSize: 20, color: PAL.yellow, lineSpacingMultiple: 0.8 });
  txt(s, 'One barrel of crude oil (159 L) can make 6,000+ products, not just fuel.',
    { x: 12.065, y: 3.134, w: 3.953, h: 0.61, fontSize: 14, color: PAL.yellow, lineSpacingMultiple: 1.1 });
  iconMark(s, 11.104, 2.804, 0.918, PAL.yellow);
  iconMark(s, 10.318, 6.11, 0.529, PAL.yellow);
}

/* -------------------------------------------- slide 4 : four numbered cols */

const GLANCE_COLS = [
  { n: '01', x: 0.896, nw: 0.958, tx: 1.854, tw: 1.543 },
  { n: '02', x: 3.930, nw: 1.109, tx: 4.888, tw: 1.509 },
  { n: '03', x: 6.964, nw: 1.109, tx: 7.922, tw: 1.523 },
  { n: '04', x: 9.997, nw: 1.109, tx: 10.955, tw: 1.482 },
];
const GLANCE_BODY = 'Lorem ipsum dolor sit amet, o consectetuer adipiscing ipsum elit. Aenean o commodo ligula eget dolor sit. Aenean massa.';

function slide04(s) {
  s.addShape('rect', { x: 0, y: -0.025, w: 13.333, h: 4.058, fill: { color: PAL.dark, transparency: 50 } });
  chrome(s, 4, { nav: PAL.white, page: PAL.black });

  GLANCE_COLS.forEach(function (c) {
    txt(s, c.n, { x: c.x, y: 4.369, w: c.nw, h: 0.909, fontFace: FONT_B, fontSize: 48, color: PAL.yellow });
    txt(s, 'Custom Itineraries', { x: c.tx, y: 4.504, w: c.tw, h: 0.64, fontFace: FONT_B, fontSize: 16, color: PAL.yellow });
    txt(s, GLANCE_BODY, { x: c.x, y: 5.191, w: 2.314, h: 1.782, fontSize: 14, color: PAL.grey, align: 'justify', lineSpacingMultiple: 1.2 });
  });

  txt(s, 'The Oil & Gas Industry at a Glance: Past, Present, and Future',
    { x: 0.567, y: 1.877, w: 10.898, h: 1.582, fontFace: FONT_B, fontSize: 44, color: PAL.yellow });
  s.addShape('rect', { x: 11.465, y: 2.883, w: 1.153, h: 1.153, fill: { color: PAL.dark }, line: { color: PAL.white, width: 3 }, shadow: SHADOW });
  iconMark(s, 11.665, 3.083, 0.752, PAL.yellow);
}

/* ------------------------------------------------ slide 5 : photo strip   */

function slide05(s) {
  para(s, { x: 2.726, y: 3.75, w: 6.67, h: 2.869, fill: { color: PAL.dark, transparency: 50 } });
  chrome(s, 5);
  para(s, { x: 9.05, y: 4.466, w: 3.61, h: 2.153, fill: { color: PAL.dark } });

  txt(s, '$15M+', { x: 9.819, y: 5.072, w: 2.19, h: 0.718, fontFace: FONT_B, fontSize: 44, color: PAL.yellow, align: 'right', lineSpacingMultiple: 0.8 });
  txt(s, 'The world consumes 15M barrels of oil daily.',
    { x: 9.383, y: 5.79, w: 2.537, h: 0.61, fontSize: 14, color: PAL.yellow, align: 'right', lineSpacingMultiple: 1.1 });

  s.addShape('rect', { x: 11.15, y: 3.816, w: 0.872, h: 0.872, fill: { color: PAL.dark }, line: { color: PAL.white, width: 3 }, shadow: SHADOW });
  iconMark(s, 11.321, 3.988, 0.529, PAL.yellow);

  txt(s, 'Energy That Powers the World and Shapes the Future',
    { x: 0.62, y: 0.881, w: 9.252, h: 1.582, fontFace: FONT_B, fontSize: 44 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. ' +
        'Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. ',
    { x: 0.62, y: 2.588, w: 9.199, h: 0.808, fontSize: 14, color: PAL.grey });

  poly(s, 0, 3.75, 3.14, 2.869, [[0, 0], [1, 0], [0.7716, 1], [0, 1]], { fill: { color: PAL.dark, transparency: 50 } });
}

/* ------------------------------------------------- slide 6 : team gallery */

const TEAM = [
  { x: 0.723, fill: PAL.yellow, tag: PAL.dark, tagText: PAL.white, name: 'Anastashia', tx: 0.839 },
  { x: 4.596, fill: PAL.dark, tag: PAL.yellow, tagText: PAL.dark, name: 'Dorione', tx: 4.709 },
  { x: 8.468, fill: PAL.yellow, tag: PAL.dark, tagText: PAL.white, name: 'William', tx: 8.6 },
];

function slide06(s) {
  chrome(s, 6);
  TEAM.forEach(function (t) {
    para(s, { x: t.x, y: 2.258, w: 4.142, h: 3.682, fill: { color: t.fill } });
  });

  txt(s, 'Behind Every Drop: Our Experts', { x: 1.667, y: 1.0, w: 10.0, h: 0.841, fontFace: FONT_B, fontSize: 44, align: 'center' });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. ',
    { x: 3.017, y: 6.273, w: 7.299, h: 0.505, fontSize: 12, color: PAL.grey, align: 'center' });

  // Carousel chevrons ("<" left, ">" right) built from four mirrored line segments.
  [[12.598, 3.287, false, false], [12.598, 4.099, false, true],
   [0.411, 3.287, true, false], [0.411, 4.099, true, true]].forEach(function (c) {
    s.addShape('line', { x: c[0], y: c[1], w: 0.312, h: 0.812, flipH: c[2], flipV: c[3], line: { color: PAL.dark, width: 2.25 } });
  });

  TEAM.forEach(function (t) {
    para(s, { x: t.tx, y: 5.076, w: 1.682, h: 0.389, adj: 0.25892, fill: { color: t.tag } });
    txt(s, t.name, { x: t.tx, y: 5.076, w: 1.682, h: 0.389, fontSize: 14, color: t.tagText, align: 'center', valign: 'middle' });
  });
}

/* --------------------------------------------- slide 7 : product showcase */

const STRATEGIES = [
  { label: 'Strategy 1', ty: 2.09, tw: 1.132, by: 2.393, iy: 2.136, ih: 0.528 },
  { label: 'Strategy 2', ty: 3.297, tw: 1.145, by: 3.6, iy: 3.347, ih: 0.601 },
  { label: 'Strategy 3', ty: 4.507, tw: 1.147, by: 4.81, iy: 4.589, ih: 0.601 },
];

function slide07(s) {
  chrome(s, 7);
  para(s, { x: 0.527, y: 1.66, w: 12.278, h: 4.18, adj: 0.20409, fill: { color: PAL.dark } });
  photo(s, 2.789, 0.929, 5.122, 5.342, '[image]');

  STRATEGIES.forEach(function (st) {
    txt(s, st.label, { x: 9.006, y: st.ty, w: st.tw, h: 0.303, fontFace: FONT_B, fontSize: 18, color: PAL.white, wrap: false, margin: 0 });
    txt(s, 'Lorem ipsum dolor sit amet, consectetuer.', { x: 8.933, y: st.by, w: 2.752, h: 0.61, fontSize: 14, color: PAL.white, lineSpacingMultiple: 1.1 });
    iconMark(s, 8.139, st.iy, st.ih, PAL.white);
  });

  txt(s, 'Smart Solutions in Action', { x: 1.905, y: 2.136, w: 3.506, h: 2.322, fontFace: FONT_B, fontSize: 44, color: PAL.white });
  ctaButton(s, { x: 1.135, y: 4.64, w: 2.406, h: 0.483, adj: 0.50884, label: 'Try it Now!', size: 16,
    tx: 1.407, ty: 4.696, tw: 1.598, th: 0.37, ax: 2.915, ay: 4.732, as: 0.287 });
}

/* ----------------------------------------------------- slide 8 : charts   */

const TREND_DARK = [0.740, 0.748, 0.714, 0.640, 0.606, 0.571, 0.518, 0.401, 0.196, 0.036, 0.003, 0.077,
  0.283, 0.525, 0.663, 0.708, 0.763, 0.838, 0.910, 0.955, 0.964, 0.936, 0.879, 0.814, 0.785, 0.756,
  0.704, 0.639, 0.607, 0.656, 0.811, 0.961, 0.998];
const TREND_YEL = [0.939, 0.878, 0.632, 0.378, 0.297, 0.348, 0.456, 0.542, 0.573, 0.477, 0.213, 0.418,
  0.959, 0.776, 0.529, 0.438, 0.420, 0.443, 0.490, 0.564, 0.669, 0.756, 0.859, 0.807, 0.369, 0.003,
  0.303, 0.508, 0.463, 0.351, 0.472, 0.493, 0.514];
const HILL_YEL = [1.000, 0.984, 0.942, 0.862, 0.763, 0.639, 0.530, 0.477, 0.447, 0.440, 0.416, 0.361,
  0.238, 0.125, 0.046, 0.007, 0.001, 0.018, 0.081, 0.198, 0.336, 0.491, 0.575, 0.616, 0.627, 0.640,
  0.667, 0.726, 0.796, 0.863, 0.909, 0.933, 0.941];
const HILL_DARK = [1.000, 0.972, 0.896, 0.718, 0.463, 0.216, 0.081, 0.016, 0.000, 0.011, 0.034, 0.094,
  0.163, 0.217, 0.250, 0.265, 0.270, 0.293, 0.348, 0.453, 0.566, 0.648, 0.692, 0.705, 0.708, 0.723,
  0.753, 0.792, 0.842, 0.881, 0.906, 0.926, 0.930];

const MONTHS = [['Jan', 1.741], ['Feb', 2.674], ['Mar', 3.567], ['Apr', 4.457], ['May', 5.403], ['Jun', 6.298]];
const KM_LABELS = [['75 km', 4.305], ['50 km', 4.953], ['25 km', 5.601], ['0 km', 6.248]];
const DEPTHS = [['-70', 8.895], ['-60', 9.320], ['-50', 9.746], ['-40', 10.171], ['-30', 10.596], ['-20', 11.021], ['-10', 11.446]];

function slide08(s) {
  poly(s, 5.479, 0, 7.854, 7.5, [[0.2387, 0], [1, 0], [1, 1], [0, 1]], { fill: { color: PAL.dark, transparency: 50 } });
  chrome(s, 8, { nav: PAL.white, brand: PAL.black, page: PAL.white });

  txt(s, 'Lorem ipsum dolor sit amet', { x: 1.006, y: 0.801, w: 2.889, h: 0.337, fontSize: 14, color: PAL.grey, wrap: false });
  s.addShape('line', { x: 0.63, y: 1.003, w: 0.377, h: 0, line: { color: PAL.yellow, width: 2.25 } });
  txt(s, 'Global Oil Production & Consumption Trends',
    { x: 0.476, y: 1.052, w: 6.414, h: 1.414, fontFace: FONT_B, fontSize: 36, color: PAL.black, margin: [14.4, 14.4, 7.2, 7.2] });

  // ---- left card: line chart drawn from sampled curves
  s.addShape('rect', { x: 0.567, y: 2.769, w: 6.414, h: 4.204, fill: { color: PAL.white }, shadow: SHADOW_CARD });
  [29, 28, 27, 26, 25].forEach(function (v, i) {
    const gy = 3.215 + i * 0.70075;
    s.addShape('line', { x: 1.596, y: gy, w: 4.903, h: 0, line: { color: PAL.gridGrey, width: 1.85 } });
    txt(s, String(v), { x: 0.85, y: gy - 0.12, w: 0.385, h: 0.24, fontSize: 12, bold: true, color: PAL.black, align: 'right', valign: 'middle', margin: 0 });
  });
  curve(s, 1.675, 3.954, 4.824, 1.268, TREND_DARK, { color: PAL.dark, width: 3.7 });
  curve(s, 1.629, 4.529, 4.824, 0.771, TREND_YEL, { color: PAL.yellow, width: 3.7 });
  MONTHS.forEach(function (m) {
    txt(s, m[0], { x: m[1] - 0.35, y: 6.36, w: 0.7, h: 0.25, fontSize: 12, bold: true, color: PAL.black, align: 'center', valign: 'middle', margin: 0 });
  });

  // ---- right card: area/depth chart
  s.addShape('rect', { x: 7.579, y: 3.867, w: 4.552, h: 3.076, fill: { color: PAL.white }, shadow: SHADOW_CARD });
  for (let i = 0; i < 8; i++) {
    s.addShape('line', { x: 8.432 + i * 0.43014, y: 4.279, w: 0, h: 2.048, line: { color: PAL.lightGrey, width: 0.75 } });
  }
  curve(s, 8.538, 4.626, 2.904, 0.868, HILL_YEL, { color: PAL.yellow, width: 25.5, transparency: 70 });
  curve(s, 8.538, 5.008, 2.892, 0.694, HILL_DARK, { color: PAL.dark, width: 4.36 });
  curve(s, 8.538, 4.626, 2.904, 0.868, HILL_YEL, { color: PAL.yellow, width: 4.36 });
  para(s, { x: 9.492, y: 4.215, w: 0.937, h: 0.289, adj: 0.48051, fill: { color: PAL.yellow } });
  s.addShape('ellipse', { x: 9.874, y: 4.545, w: 0.15, h: 0.15, fill: { color: PAL.paleYellow }, line: { color: PAL.white, width: 1.5 } });
  KM_LABELS.forEach(function (l) {
    txt(s, l[0], { x: 7.55, y: l[1] - 0.06, w: 0.7, h: 0.24, fontSize: 10, bold: true, color: PAL.black, align: 'right', valign: 'middle', margin: 0 });
  });
  DEPTHS.forEach(function (d) {
    txt(s, d[0], { x: d[1] - 0.3, y: 6.43, w: 0.6, h: 0.22, fontSize: 9, bold: true, color: PAL.black, align: 'center', valign: 'middle', margin: 0 });
  });

  s.addShape('rect', { x: 6.312, y: 2.559, w: 0.872, h: 0.872, fill: { color: PAL.yellow }, line: { color: PAL.dark, width: 3 }, shadow: SHADOW });
  iconMark(s, 6.483, 2.731, 0.529, PAL.dark);
  s.addShape('rect', { x: 11.895, y: 5.643, w: 0.872, h: 0.872, fill: { color: PAL.yellow }, line: { color: PAL.dark, width: 3 }, shadow: SHADOW });
  iconMark(s, 12.067, 5.814, 0.529, PAL.dark);
}

/* ------------------------------------------- slide 9 : doughnut infograph */

const DONUT_OPTS = {
  holeSize: 75, firstSliceAng: 0, showLegend: false, showTitle: false,
  showValue: false, showPercent: false, showLabel: false,
  chartColors: [PAL.yellow, PAL.dark], dataBorder: { pt: 1.5, color: PAL.white },
};

function slide09(s, pptx) {
  poly(s, 5.492, 0, 7.842, 7.5, [[0, 0], [1, 0], [1, 1], [0.255, 1]], { fill: { color: PAL.dark, transparency: 50 } });
  chrome(s, 9, { nav: PAL.white, brand: PAL.black });

  txt(s, 'Lorem ipsum dolor sit amet', { x: 1.198, y: 0.979, w: 2.889, h: 0.337, fontSize: 14, color: PAL.grey, wrap: false });
  s.addShape('line', { x: 0.821, y: 1.181, w: 0.377, h: 0, line: { color: PAL.yellow, width: 2.25 } });
  txt(s, 'Infographic Chart Section', { x: 0.705, y: 1.432, w: 5.089, h: 1.717, fontFace: FONT_B, fontSize: 48 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor.',
    { x: 0.705, y: 3.353, w: 4.353, h: 0.869, fontSize: 14, color: PAL.grey, lineSpacingMultiple: 1.1 });

  s.addShape('roundRect', { x: 1.019, y: 4.782, w: 3.585, h: 1.717, rectRadius: 0.17, fill: { color: PAL.white }, shadow: SHADOW_SOFT });
  s.addShape('roundRect', { x: 4.954, y: 4.782, w: 3.585, h: 1.717, rectRadius: 0.17, fill: { color: PAL.dark }, shadow: SHADOW_DEEP });

  [[1.353, 5.146, 1.555, 5.348, PAL.dark], [5.283, 5.146, 5.492, 5.353, PAL.yellow]].forEach(function (d) {
    s.addShape('donut', { x: d[0], y: d[1], w: 0.978, h: 0.978, fill: { color: d[4] } });
    s.addShape('ellipse', { x: d[2], y: d[3], w: 0.574, h: 0.574, fill: { color: d[4] } });
  });

  txt(s, '70%', { x: 2.568, y: 5.021, w: 1.734, h: 1.01, fontFace: FONT_B, fontSize: 54, align: 'center', wrap: false });
  txt(s, 'Subtitle Here', { x: 3.059, y: 5.876, w: 1.431, h: 0.337, fontSize: 14, color: PAL.grey, wrap: false });
  txt(s, '50%', { x: 6.476, y: 5.001, w: 1.806, h: 1.01, fontFace: FONT_B, fontSize: 54, color: PAL.yellow, align: 'center', wrap: false });
  txt(s, 'Subtitle Here', { x: 6.77, y: 5.869, w: 1.601, h: 0.37, fontSize: 16, color: PAL.white, wrap: false });

  iconMark(s, 1.663, 5.467, 0.376, PAL.white);
  iconMark(s, 5.606, 5.462, 0.376, PAL.white);

  s.addChart(pptx.ChartType.doughnut, [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: [8.2, 3.2] }],
    Object.assign({ x: 1.116, y: 4.895, w: 1.452, h: 1.481 }, DONUT_OPTS));
  s.addChart(pptx.ChartType.doughnut, [{ name: 'Sales', labels: ['2nd Qtr', '3rd Qtr'], values: [3.2, 3.2] }],
    Object.assign({ x: 5.054, y: 4.941, w: 1.436, h: 1.417 }, DONUT_OPTS));
}

/* --------------------------------------------------- slide 10 : world map */

/* Mosaic stand-in for the vector world map: one character per 0.1" cell.
   g = neutral land, y = yellow, c = teal, d = dark, . = empty */
const MAP_ROWS = [
  '.........................................yy.......ggg',
  '........................................yy..ggggggggggg',
  '...........................................gggggggggggg...............................dd',
  '.............................................ggggggggg............................dddddddd',
  '................d...........yyy..y..yyyy.....gggggggg..........................d.ddddddddddddddddddddd.d',
  '.............dddddd.yyyyyyy..yy..yyyy..yy....gggggg..............cccddd....dddddddddddddddddddddddddddddddddddd',
  '............dddddd.yyyyyyyyyyyyyyyy...yy....ggggg...............cccccd..ddddddddddddddddddddddddddddddddddddddddd',
  '...........ddddddyyyyyyyyyyyyyyy.......y.....gg...............cccc.ccdddddddddddddddddddddddddddddddddddddddddddd',
  '..........dddd.d.yyyyyyyyyyyyyy.....yy.......gg...............cccc.c.ddddddddddddddddddddddddddddddddddddddd..ddd',
  '..................yyyyyyyyyyyy......yyyy........................c....dddddddddddddddddddddddddddddddddddd.....d',
  '..................yyyyyyyyyyyyyy...yyyyy...........................cccdddddddddddddddddddddddddddddddddd......dd',
  '.................yyyyyyyyyyyyyyyy.yyyyyyy..................c...ccccccc.dddddddddyyyyyydddddddddddddddddddd',
  '.................yyyyyyyyyyyyyyyyyyyyyyy......................ccccccyyyydddddyddyyyyyyy.dddddgddddddddddddd',
  '...................dddddddd.yyyyyyyyy......................yyy.ccc..yyyyy.ddyyyyyyyyyyyyy.ggggggggggdddddddd',
  '.................dddddddddddd..yyyyy.......................yyy.cc.cccc.y.dddd.yyyyyyyyyyyddggggggggggddddddd',
  '................ddddddddddddd.d..ddd........................yy.cc..ccc....ddd.yy.cyyyyyydddddggggggdddddddd',
  '................ddddddddddddddddddd......................cccc...c...c...d.......ggccygggddddddddddddddddd',
  '...............dddddddddddddddddd........................ccc..........ddddddg..ggggccggdddddddddddddddd..g',
  '................dddddddddddddddd.........................cc...........dddddddd..dgggccgddddddddddddddddd..g',
  '................dddddddddddddddd..........................yyyyyy.........ggggddddd.cccyy.ddddddddddddddd',
  'ggg..............ddddddddddddd...........................yyyyyyy.........ggggddddd.cccyycdddddddddddddddd',
  '..................yyyddddddddd..........................yyyyyyy.gggggddd.gyyggdddddccyyyccddddddddddddddd',
  '..................yyyydd.....d..........................yyyyyyyygggggddddyyyyy.dddd.yyyccc..dddc.ddddddddd',
  '...................yyyy................................ygyyyyyyyggggggddd.yyyyy...d.yy.cccccc.cc.dddddddd',
  '.......................................................gggcyyyyyggggggddd..yyyyygg....cccccccc..gdddddddd',
  '......................................................ggggccyyycccgggggggg.yyyyyygg....cccccc..gggg...d',
  '......................................................ggggccc.ccccgggggggg..yyyygg......ccccc...ggyg',
  '......................................................ggggccccccccggggggggg..gggg.......cccc....ggyy',
  '........................................................cccccccccggggggggggggggg.........cc.......yyy',
  '..........................gg............................ccyy.cccc.ggggggggggg............cc',
  '...........................g...ycc.c..................ggggy..ccccggggggggggggggg..........c',
  '..............................yycccc...................ggggc.ccccggggggggggggggg',
  '..............................yyycccgg..................ggg...c.ggggggggggggggg',
  '..............................yyyycd.gg.........................gggggggggccggg.....................y.....y',
  '..............................yyy..ddddd........................ggggggggggccg.....................yy...yyy',
  '.............................ccyydddddddd......................ggggggggg.ccc.......................yy..yyy',
  '.............................cyy.ddddddddddd....................gggggggggggc........................y...y',
  '.............................yyydddddddddddddd...................gggggggggg',
  '.............................yyddddddddddddddd...................gggggggggg',
  '..............................yydddddddddddddd...................ggggggggggg',
  '..............................yyyccdddddddddd....................ggggggggggg...................................y',
  '...............................yyccc.dddddddd....................ggggggggggg..................................yy',
  '................................yccccdddddddd....................gggggggggg..cc.............................yyyyy',
  '..................................ccccddddddd....................ggggggggg...cc............................yyyyyy',
  '..................................ccyydddddd.....................gggggggg....c...........................yyyyyyyy',
  '.................................ycccyydddd......................ggggggygg.............................yyyyyyyyyy',
  '.................................yccccy.dd........................ggg.yyg..............................yyyyyyyyyy',
  '..................................cccccddd........................g.yyyy...............................yyyyyyyyyy',
  '..................................cccccdd..........................yyy.y...............................yyyyyyyyyy',
  '..................................ccccyy...........................yyyy................................yyyy..yyyy',
  '..................................cccccy...............................................................y.......yy',
  '.................................yccccc........................................................................yy',
  '..................................cccc',
  '..................................ccc',
  '...................................cc',
  '...................................cc',
  '...................................cc',
  '...................................cc',
];
const MAP_COLORS = { g: PAL.mapGrey, y: PAL.yellow, c: PAL.teal, d: PAL.dark };
const MAP_ORIGIN = { x: 2.05, y: 0.75, cell: 0.10 };

function worldMap(s) {
  MAP_ROWS.forEach(function (row, j) {
    let i = 0;
    while (i < row.length) {
      const ch = row[i];
      if (ch === '.') { i++; continue; }
      let k = i;
      while (k + 1 < row.length && row[k + 1] === ch) k++;
      s.addShape('rect', {
        x: MAP_ORIGIN.x + i * MAP_ORIGIN.cell, y: MAP_ORIGIN.y + j * MAP_ORIGIN.cell,
        w: (k - i + 1) * MAP_ORIGIN.cell, h: MAP_ORIGIN.cell, fill: { color: MAP_COLORS[ch] },
      });
      i = k + 1;
    }
  });
}

function slide10(s) {
  chrome(s, 10, { page: PAL.dark });
  worldMap(s);
  txt(s, [{ text: 'There are 1,349,166 barrels of oil', options: { breakLine: true } },
          { text: 'consumed globally every day as of April 2025.' }],
    { x: 0.567, y: 3.227, w: 3.787, h: 3.871, fontFace: FONT_B, fontSize: 32 });
  txt(s, [{ text: 'That\u2019s +70,640 barrels (+3.8%) more than the previous day,', options: { breakLine: true } },
          { text: 'and 88,281 barrels more compared to two days ago.' }],
    { x: 6.484, y: 6.441, w: 6.07, h: 0.572, fontSize: 14, color: PAL.grey });
  ctaButton(s, { x: -0.24, y: 2.132, w: 2.958, h: 0.594, adj: 0.50884, label: 'Learn More', size: 20,
    tx: 0.095, ty: 2.201, tw: 1.965, th: 0.438, ax: 1.949, ay: 2.245, as: 0.353 });
}

/* ------------------------------------------- slide 11 : arrow step (rows) */

/* Long horizontal arrow with a stepped tail. The notches sit at fixed
   distances from the left edge, so the profile is built in inches and then
   normalised by the row width. */
function stepArrow(w) {
  const f = function (inches) { return inches / w; };
  const head = f(w - 1.61);
  return [[head, 0], [1, 0.3595], [head, 0.7189], [head, 0.6132], [f(3.77), 0.6132],
          [f(1.66), 1], [0, 1], [0, 0.4925], [f(1.37), 0.4925], [f(3.47), 0.1066], [head, 0.1057]];
}
const ARROW_ROWS = [
  { y: 3.141, w: 11.917, color: PAL.yellow },
  { y: 1.942, w: 10.789, color: PAL.dark },
  { y: 4.340, w: 10.247, color: PAL.teal },
];
const ARROW_STEPS = [
  { ty: 2.311, by: 2.642, bw: 4.57, iy: 2.398, num: '01.', ny: 2.961, nx: 1.91, glyph: PAL.dark },
  { ty: 3.539, by: 3.870, bw: 4.67, iy: 3.663, num: '02.', ny: 4.164, nx: 1.91, glyph: PAL.yellow },
  { ty: 4.709, by: 5.039, bw: 4.57, iy: 4.862, num: '03.', ny: 5.367, nx: 1.933, glyph: PAL.teal },
];

function slide11(s) {
  chrome(s, 11);
  ARROW_ROWS.forEach(function (a) {
    poly(s, 0, a.y, a.w, 2.365, stepArrow(a.w), { fill: { color: a.color } });
  });
  ARROW_STEPS.forEach(function (st) {
    txt(s, 'Your Title', { x: 4.403, y: st.ty, w: 1.347, h: 0.362, fontFace: FONT_B, fontSize: 14, bold: true, color: PAL.white, lineSpacingMultiple: 1.2 });
    txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor.',
      { x: 4.403, y: st.by, w: st.bw, h: 0.572, fontSize: 12, color: PAL.white, lineSpacingMultiple: 1.2 });
    s.addShape('roundRect', { x: 3.515, y: st.iy, w: 0.639, h: 0.639, rectRadius: 0.16, fill: { color: PAL.white }, shadow: SHADOW_SOFT });
    iconMark(s, 3.615, st.iy + 0.1, 0.44, st.glyph);
    txt(s, st.num, { x: st.nx, y: st.ny, w: 0.945, h: 0.707, fontFace: FONT_B, fontSize: 36, color: PAL.white, align: 'center', wrap: false });
  });
  txt(s, [{ text: 'Arrow Step ', options: { color: PAL.greyDark } }, { text: 'Infographic.', options: { color: PAL.dark } }],
    { x: 1.272, y: 0.853, w: 10.789, h: 0.774, fontFace: FONT_B, fontSize: 40, align: 'center' });
  ctaButton(s, { x: 9.933, y: 6.054, w: 2.958, h: 0.594, adj: 0.50884, label: 'Learn More', size: 20,
    tx: 10.268, ty: 6.123, tw: 1.965, th: 0.438, ax: 12.122, ay: 6.167, as: 0.353 });
}

/* ------------------------------------- slide 12 : staircase arrow diagram */

const STAIR = [[0.8396, 0.0000], [1.0000, 0.0965], [0.8396, 0.1930], [0.8396, 0.1484], [0.6039, 0.1484],
  [0.6039, 0.4199], [0.5883, 0.4372], [0.3475, 0.4372], [0.3475, 0.6964], [0.3319, 0.6964],
  [0.0936, 0.6964], [0.0936, 0.9827], [0.0780, 1.0000], [0.0156, 1.0000], [0.0000, 0.9827],
  [0.0000, 0.5924], [0.0156, 0.5924], [0.2539, 0.5924], [0.2539, 0.3333], [0.2695, 0.3333],
  [0.5103, 0.3333], [0.5103, 0.0445], [0.5259, 0.0445], [0.8396, 0.0445]];

const STAIR_OPTIONS = [
  { title: 'Option 1', color: PAL.dark, tx: 4.747, ty: 1.035, bx: 3.068, by: 1.387, bw: 3.016, align: 'right', ox: 6.220, oy: 1.128 },
  { title: 'Option 2', color: PAL.dark, tx: 3.273, ty: 2.703, bx: 1.594, by: 3.055, bw: 3.016, align: 'right', ox: 4.746, oy: 2.796 },
  { title: 'Option 3', color: PAL.yellow, tx: 10.257, ty: 3.750, bx: 10.257, by: 4.102, bw: 2.758, align: 'left', ox: 9.504, oy: 3.807 },
  { title: 'Option 4', color: PAL.yellow, tx: 8.794, ty: 5.266, bx: 8.794, by: 5.618, bw: 2.570, align: 'left', ox: 8.041, oy: 5.309 },
];

function slide12(s) {
  chrome(s, 12);
  txt(s, 'Arrow Step Infographic.', { x: 0.899, y: 5.346, w: 3.879, h: 1.447, fontFace: FONT_B, fontSize: 40 });
  poly(s, 5.564, 2.36, 5.8, 5.3, STAIR, { fill: { color: PAL.yellow } });
  poly(s, 3.389, -0.16, 5.8, 5.3, STAIR, { fill: { color: PAL.dark }, flipH: true, flipV: true });

  STAIR_OPTIONS.forEach(function (op) {
    txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ',
      { x: op.bx, y: op.by, w: op.bw, h: 0.572, fontSize: 12, color: PAL.grey, align: op.align, lineSpacingMultiple: 1.2 });
    txt(s, op.title, { x: op.tx, y: op.ty, w: 1.577, h: 0.37, fontFace: FONT_B, fontSize: 16, color: op.color, align: op.align });
    s.addShape('ellipse', { x: op.ox, y: op.oy, w: 0.591, h: 0.591, fill: { color: PAL.white }, shadow: SHADOW });
    iconMark(s, op.ox + 0.105, op.oy + 0.105, 0.38, PAL.greyDark);
  });
}

/* ---------------------------------------- slide 13 : road-map with images */

const ROAD13_OUTLINE = [[0.0000, 0.0000], [0.0000, 0.2932], [0.5731, 0.2932], [0.5731, 0.3536], [0.3948, 0.3536],
  [0.3759, 0.3818], [0.3644, 0.4536], [0.3628, 0.4998], [0.3628, 0.8534], [0.3689, 0.9400], [0.3847, 0.9925],
  [0.3948, 1.0000], [1.0000, 1.0000], [1.0000, 0.7068], [0.4269, 0.7068], [0.4269, 0.6464], [0.6052, 0.6464],
  [0.6241, 0.6182], [0.6356, 0.5462], [0.6373, 0.4998], [0.6373, 0.1466], [0.6311, 0.0600], [0.6153, 0.0075],
  [0.6052, 0.0000]];
const ROAD13_SURFACE = [[0.0000, 0.0000], [0.0000, 0.1867], [0.5874, 0.1867], [0.5874, 0.4069], [0.3948, 0.4069],
  [0.3843, 0.4248], [0.3780, 0.4704], [0.3771, 0.4998], [0.3771, 0.9067], [0.3805, 0.9618], [0.3892, 0.9952],
  [0.3948, 1.0000], [1.0000, 1.0000], [1.0000, 0.8133], [0.4125, 0.8133], [0.4125, 0.5931], [0.6052, 0.5931],
  [0.6156, 0.5752], [0.6220, 0.5294], [0.6229, 0.4998], [0.6229, 0.0934], [0.6195, 0.0383], [0.6108, 0.0048],
  [0.6052, 0.0000]];

/* Centre-line dash segments: [x, y, w, h, count] (count 1 == corner marker). */
const ROAD13_DASHES = [
  [5.647, 3.970, 4.319, 0.048, 10], [10.245, 3.970, 2.938, 0.048, 7], [0.182, 1.906, 2.938, 0.048, 7],
  [5.246, 3.905, 0.112, 0.113, 1], [5.246, 3.380, 0.047, 0.196, 1], [5.246, 2.938, 0.112, 0.112, 1],
  [5.651, 2.938, 2.047, 0.047, 5], [7.988, 2.873, 0.113, 0.112, 1], [8.054, 2.347, 0.048, 0.197, 1],
  [7.988, 1.906, 0.113, 0.112, 1], [3.374, 1.906, 4.327, 0.048, 10],
];
const ROAD13_BADGES = [[5.477, 1.389, PAL.yellow], [7.517, 3.490, PAL.dark], [11.186, 3.490, PAL.yellow], [1.600, 1.389, PAL.dark]];
const ROAD13_STEPS = [
  { label: 'Step 1', x: 0.423, color: PAL.dark, body: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit' },
  { label: 'Step 2', x: 3.730, color: PAL.yellow, body: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit' },
  { label: 'Step 3', x: 7.037, color: PAL.dark, body: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit' },
  { label: 'Step 4', x: 10.327, color: PAL.yellow, body: 'Lorem ipsum dolor sit amet, consectetuer adipiscing' },
];

function slide13(s) {
  chrome(s, 13, { page: PAL.white });
  poly(s, 0, 1.501, 13.348, 2.921, ROAD13_OUTLINE, { fill: { color: PAL.lightGrey } });
  poly(s, 0, 1.693, 13.348, 2.538, ROAD13_SURFACE, { fill: { color: PAL.road } });

  ROAD13_DASHES.forEach(function (d) {
    const n = d[4];
    const horizontal = d[2] >= d[3];
    const span = horizontal ? d[2] : d[3];
    // Dash occupies 37% of its period, so period = span / (n - 1 + 0.37).
    const period = n === 1 ? span : span / (n - 1 + 0.374);
    const len = n === 1 ? span : period * 0.374;
    for (let i = 0; i < n; i++) {
      const off = i * period;
      s.addShape('rect', {
        x: d[0] + (horizontal ? off : 0), y: d[1] + (horizontal ? 0 : off),
        w: horizontal ? len : d[2], h: horizontal ? d[3] : len, fill: { color: PAL.white },
      });
    }
  });

  ROAD13_BADGES.forEach(function (b) {
    s.addShape('ellipse', { x: b[0], y: b[1], w: 1.056, h: 1.056, fill: { color: b[2] } });
    iconMark(s, b[0] + 0.26, b[1] + 0.26, 0.534, PAL.white);
  });

  txt(s, 'Roadmap Infographic Section.', { x: 7.61, y: 0.936, w: 5.337, h: 2.106, fontFace: FONT_B, fontSize: 44, align: 'right', lineSpacingMultiple: 0.9 });
  txt(s, 'Subtitle Here', { x: 0.609, y: 3.132, w: 3.135, h: 0.407, fontFace: FONT_B, fontSize: 20, lineSpacingMultiple: 0.9 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. ',
    { x: 0.609, y: 3.533, w: 3.244, h: 1.128, fontSize: 14, color: PAL.grey, lineSpacingMultiple: 1.1 });

  poly(s, -0.01, 5.47, 3.648, 2.03, [[0, 0], [1, 0], [0.8609, 1], [0, 1]], { fill: { color: PAL.yellow } });
  para(s, { x: 3.099, y: 5.47, w: 3.846, h: 2.03, fill: { color: PAL.dark } });
  para(s, { x: 6.406, y: 5.47, w: 3.846, h: 2.03, fill: { color: PAL.yellow } });
  poly(s, 9.695, 5.47, 3.652, 2.03, [[0.139, 0], [1, 0], [1, 1], [0, 1]], { fill: { color: PAL.dark } });

  ROAD13_STEPS.forEach(function (st) {
    txt(s, st.label, { x: st.x, y: 5.861, w: 2.583, h: 0.464, fontFace: FONT_B, fontSize: 24, color: st.color, align: 'center', lineSpacingMultiple: 0.9 });
    txt(s, st.body, { x: st.x, y: 6.240, w: 2.583, h: 0.869, fontSize: 14, color: st.color, align: 'center', lineSpacingMultiple: 1.1 });
  });
}

/* ------------------------------------------- slide 14 : winding road pins */

const ROAD14 = [[0.0000, 1.0000], [0.7465, 0.6358], [0.7526, 0.6259], [0.7479, 0.6148], [0.5002, 0.4470],
  [0.4683, 0.4210], [0.4703, 0.3990], [0.5049, 0.3831], [0.9029, 0.2441], [0.9069, 0.2378], [0.9034, 0.2311],
  [0.6723, 0.1184], [0.6678, 0.1101], [0.6725, 0.1025], [0.9248, 0.0000], [0.7057, 0.1053], [0.7032, 0.1098],
  [0.7058, 0.1139], [0.9771, 0.2222], [0.9947, 0.2326], [0.9948, 0.2509], [0.9764, 0.2618], [0.6055, 0.4147],
  [0.6004, 0.4228], [0.6046, 0.4315], [0.8791, 0.5931], [0.9288, 0.6266], [0.9282, 0.6509], [0.8753, 0.6816],
  [0.2223, 1.0000]];

/* White centre-line dashes along the winding road: [x, y, w, h]. */
const ROAD14_DASHES = [
  [1.167, 7.354, 0.215, 0.104], [1.572, 7.226, 0.213, 0.104], [1.973, 7.099, 0.214, 0.102], [2.377, 6.971, 0.213, 0.102],
  [2.741, 6.854, 0.215, 0.104], [3.144, 6.726, 0.213, 0.104], [3.546, 6.598, 0.215, 0.102], [3.949, 6.470, 0.213, 0.103],
  [4.351, 6.343, 0.213, 0.101], [4.754, 6.213, 0.213, 0.101], [5.156, 6.086, 0.213, 0.101], [5.560, 5.958, 0.213, 0.099],
  [5.961, 5.830, 0.213, 0.099], [6.365, 5.703, 0.211, 0.097], [6.766, 5.575, 0.213, 0.097], [7.170, 5.447, 0.211, 0.095],
  [7.572, 5.318, 0.211, 0.095], [7.975, 5.190, 0.211, 0.095], [8.377, 5.063, 0.211, 0.093], [8.780, 4.935, 0.211, 0.093],
  [9.182, 4.808, 0.211, 0.092], [9.585, 4.582, 0.085, 0.189], [9.211, 4.418, 0.206, 0.109], [8.823, 4.252, 0.205, 0.109],
  [8.433, 4.087, 0.206, 0.107], [8.045, 3.922, 0.204, 0.108], [7.657, 3.757, 0.204, 0.105], [7.269, 3.592, 0.204, 0.105],
  [6.880, 3.425, 0.203, 0.105], [6.491, 3.260, 0.205, 0.106], [6.103, 3.095, 0.205, 0.104], [6.086, 2.856, 0.211, 0.080],
  [6.496, 2.753, 0.210, 0.072], [6.904, 2.647, 0.210, 0.073], [7.315, 2.544, 0.210, 0.071], [7.723, 2.438, 0.210, 0.072],
  [8.133, 2.334, 0.208, 0.070], [8.542, 2.230, 0.210, 0.068], [8.952, 2.125, 0.208, 0.068], [9.361, 2.021, 0.210, 0.068],
  [9.769, 1.915, 0.209, 0.068], [10.180, 1.811, 0.207, 0.067], [10.588, 1.706, 0.209, 0.066], [10.666, 1.513, 0.208, 0.074],
  [10.261, 1.394, 0.206, 0.072], [9.856, 1.277, 0.206, 0.070], [9.451, 1.158, 0.206, 0.070], [9.046, 1.039, 0.205, 0.070],
  [8.641, 0.920, 0.205, 0.070], [8.235, 0.802, 0.205, 0.066], [7.858, 0.651, 0.177, 0.099], [8.028, 0.491, 0.206, 0.064],
  [8.435, 0.377, 0.204, 0.063], [8.842, 0.261, 0.204, 0.061], [9.249, 0.146, 0.202, 0.060], [9.654, 0.031, 0.204, 0.060],
];

/* Map pins: [x, y, w, h, colour] */
const ROAD14_PINS = [
  [9.686, 2.919, 1.057, 1.859, PAL.peach], [1.766, 4.559, 1.311, 2.299, PAL.dark],
  [4.833, 3.758, 1.189, 2.084, PAL.yellow], [7.554, 0.897, 0.849, 1.484, PAL.teal],
  [5.270, 1.330, 1.000, 1.748, PAL.orange],
];

function mapPin(s, x, y, w, h, color) {
  s.addShape('ellipse', { x: x, y: y, w: w, h: w, fill: { color: color } });
  s.addShape('ellipse', { x: x + w * 0.115, y: y + w * 0.115, w: w * 0.77, h: w * 0.77, fill: { color: PAL.white } });
  s.addShape('rect', { x: x + w / 2 - w * 0.035, y: y + w * 0.94, w: w * 0.07, h: h - w * 1.06, fill: { color: color } });
  s.addShape('ellipse', { x: x + w / 2 - w * 0.1, y: y + h - w * 0.2, w: w * 0.2, h: w * 0.2, fill: { color: color } });
  iconMark(s, x + w * 0.27, y + w * 0.27, w * 0.46, color);
}

function slide14(s) {
  chrome(s, 14);
  poly(s, 0, -0.192, 11.452, 7.692, ROAD14, { fill: { color: PAL.road2 } });
  ROAD14_DASHES.forEach(function (d) {
    // Each dash follows the local road slope, so draw it as a short stroke
    // running from the bottom-left to the top-right of its box.
    s.addShape('line', { x: d[0], y: d[1] + d[3], w: d[2], h: d[3], flipV: true, line: { color: PAL.white, width: 3 } });
  });
  ROAD14_PINS.forEach(function (p) { mapPin(s, p[0], p[1], p[2], p[3], p[4]); });

  txt(s, 'Roadmap Infographic Section.', { x: 0.627, y: 0.879, w: 3.928, h: 2.084, fontFace: FONT_B, fontSize: 44, valign: 'middle' });
  txt(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings.',
    { x: 0.627, y: 3.057, w: 3.384, h: 1.154, fontSize: 14, color: PAL.grey, align: 'justify' });
  txt(s, [{ text: 'That\u2019s +70,640 barrels (+3.8%) more than the previous day,', options: { breakLine: true } },
          { text: 'and 88,281 barrels more compared to two days ago.' }],
    { x: 6.237, y: 6.619, w: 5.948, h: 0.572, fontSize: 14, color: PAL.grey });
  ctaButton(s, { x: 9.952, y: 5.65, w: 2.958, h: 0.594, adj: 0.50884, label: 'Learn More', size: 20,
    tx: 10.287, ty: 5.719, tw: 1.965, th: 0.438, ax: 12.141, ay: 5.763, as: 0.353 });
}

/* --------------------------------------------- slide 15 : testimonials    */

const QUOTES = [
  { bx: 1.019, by: 2.088, px: 1.211, py: 2.252, tx: 2.363, ty: 2.263, ix: 4.783, iy: 2.758,
    text: '\u201CEnergy Flow helped us cut downtime and boost efficiency.\u201D' },
  { bx: 7.532, by: 4.857, px: 7.724, py: 5.022, tx: 8.876, ty: 5.032, ix: 11.296, iy: 5.528,
    text: '\u201CTheir solutions lowered costs while keeping safety first.\u201D' },
];

function slide15(s) {
  chrome(s, 15);
  txt(s, 'Trusted by Leading Energy Partners', { x: 7.196, y: 0.873, w: 5.714, h: 1.447, fontFace: FONT_B, fontSize: 40, align: 'right' });
  para(s, { x: 2.328, y: -0.018, w: 4.833, h: 5.07, fill: { color: PAL.yellow, transparency: 50 } });
  para(s, { x: 5.779, y: 2.667, w: 4.833, h: 4.833, fill: { color: PAL.yellow, transparency: 50 } });

  QUOTES.forEach(function (q) {
    para(s, { x: q.bx, y: q.by, w: 4.222, h: 1.158, fill: { color: PAL.dark } });
    para(s, { x: q.px, y: q.py, w: 1.117, h: 0.829, fill: { color: PAL.yellow } });
    txt(s, q.text, { x: q.tx, y: q.ty, w: 2.542, h: 0.808, fontSize: 14, color: PAL.white });
    s.addShape('rect', { x: q.ix, y: q.iy, w: 0.625, h: 0.625, fill: { color: PAL.yellow }, line: { color: PAL.dark, width: 3 }, shadow: SHADOW });
    iconMark(s, q.ix + 0.122, q.iy + 0.123, 0.379, PAL.dark);
  });

  txt(s, '\u201CPartner with Us for Smarter Energy Solutions\u201D',
    { x: 0.835, y: 5.497, w: 4.608, h: 0.707, fontSize: 18, color: PAL.grey });
  ctaButton(s, { x: -0.502, y: 6.379, w: 4.408, h: 0.594, adj: 0.50884, label: 'Discover Insights', size: 20,
    tx: 0.256, ty: 6.457, tw: 2.892, th: 0.438, ax: 3.139, ay: 6.5, as: 0.36 });
}

/* --------------------------------------------------- slide 16 : closing   */

const CLOSING_STRIPES = [
  [2.219, 1.399, 3.868, 2.879], [0.807, -0.016, 3.679, 2.754],
  [-1.415, 3.363, 3.679, 2.754], [0.0, 4.279, 4.021, 3.237],
];

function slide16(s) {
  s.addShape('rect', { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: PAL.dark, transparency: 50 } });
  chrome(s, 16, { nav: PAL.white });

  para(s, { x: CLOSING_STRIPES[0][0], y: CLOSING_STRIPES[0][1], w: CLOSING_STRIPES[0][2], h: CLOSING_STRIPES[0][3], adj: 0.66848, fill: { color: PAL.white } });
  para(s, { x: 1.294, y: 2.021, w: 10.745, h: 3.457, fill: { color: PAL.yellow } });
  CLOSING_STRIPES.slice(1).forEach(function (r) {
    para(s, { x: r[0], y: r[1], w: r[2], h: r[3], adj: 0.66848, fill: { color: PAL.white } });
  });

  txt(s, 'Fueling Progress, Together Towards the Future!',
    { x: 1.818, y: 2.172, w: 9.698, h: 2.827, fontFace: FONT_B, fontSize: 54, align: 'center' });

  para(s, { x: 5.583, y: 5.151, w: 5.267, h: 0.541, fill: { color: PAL.dark }, shadow: SHADOW_DEEP });
  iconMark(s, 5.768, 5.299, 0.254, PAL.white);
  txt(s, '+123 456 789 012 345', { x: 6.01, y: 5.258, w: 2.273, h: 0.337, fontSize: 14, color: PAL.white, valign: 'middle' });
  iconMark(s, 8.21, 5.283, 0.268, PAL.white);
  txt(s, 'hydrax@gmail.com', { x: 8.478, y: 5.248, w: 2.188, h: 0.337, fontSize: 14, color: PAL.white, valign: 'middle' });
}

/* -------------------------------------------------------------- assembly */

const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
                slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE';
  pptx.author = 'HydraX';
  pptx.title = 'Oil & Gas';

  SLIDES.forEach(function (fn) {
    const slide = pptx.addSlide();
    slide.background = { color: PAL.white };
    fn(slide, pptx);
  });

  return pptx.writeFile({ fileName: path.join(__dirname, '09d9b551-e9c8-453c-ac22-f4309a9f2779_grok_final.pptx') });
}

build().then(function (f) { console.log('wrote ' + f); }).catch(function (e) { console.error(e); process.exit(1); });
