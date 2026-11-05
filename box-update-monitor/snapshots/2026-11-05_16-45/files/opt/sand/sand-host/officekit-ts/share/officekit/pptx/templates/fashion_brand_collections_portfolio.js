#!/usr/bin/env node
/**
 * "Lloris Fashion" — 42-slide template, rebuilt with pptxgenjs.
 *
 * Every slide is a small builder function; all geometry, colour and copy are
 * plain literals so the deck's design can be read straight out of this file.
 *
 *   node 168aa179-6900-4175-ae0d-38b72b804700_grok_final.js
 *
 * Photographic placeholders in the original are intentionally left empty (they
 * are empty picture placeholders there too); the three device mock-ups that the
 * source draws with vector art are redrawn from native shapes.
 */
'use strict';

const pptxgen = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ palette */
const DARK      = '443C29';  // deep espresso brown — the deck's anchor colour
const CREAM     = 'F2EDE7';  // warm off-white page
const PEACH     = 'FDE5CD';  // pale peach block
const APRICOT   = 'FBCC9A';  // saturated peach block
const SAND      = 'D3CBB6';  // muted sand band
const OLIVE     = '887952';  // olive / brass accent
const KHAKI     = 'BAAD8D';  // khaki full-bleed
const STEEL     = 'BAC9D0';  // progress-bar track
const SILVER    = 'BFBFBF';  // dotted timeline rule
const SLATE     = '595959';  // phone body
const GRAPHITE  = '383838';  // laptop screen
const SCREEN    = 'D9D9D9';  // device screen fill
const MIST      = 'F3F3F3';  // faint grey page
const WHITE     = 'FFFFFF';
const OFFWHITE  = 'F2F2F2';
const INK       = '000000';
const ICON      = '111111';  // pictogram black
const COAL      = '1B1810';
const NEARBLACK = '0D0D0D';
const CHARCOAL  = '262626';
const GRAY      = '404040';
const BRONZE    = '996633';

/* --------------------------------------------------------------------- type */
const SERIF = 'Cormorant Garamond';  // display face
const SANS  = 'Lato';                // body face
const BODY_SIZE = 18;                // the deck's inherited default size

/* --------------------------------------------------------------- primitives */

/** Flat colour block; `o.type` selects another preset geometry. */
function box(s, x, y, w, h, color, o) {
  o = o || {};
  const fill = { color: color };
  if (o.transparency) fill.transparency = o.transparency;
  s.addShape(o.type || 'rect', { x: x, y: y, w: w, h: h, fill: fill, line: { type: 'none' } });
}

/** Outline-only shape (icon strokes, device frames). */
function stroke(s, type, x, y, w, h, color, width, o) {
  s.addShape(type, Object.assign({
    x: x, y: y, w: w, h: h,
    fill: { type: 'none' },
    line: { color: color, width: width }
  }, o || {}));
}

/**
 * Text block. Options mirror the source shape properties:
 *   face size color bold align valign wrap ls (line spacing) rot fill
 * `content` is a string ('\n' starts a paragraph) or an array of
 * { t, size, bold, br } run objects when a line mixes formats.
 */
function text(s, content, o) {
  const body = Array.isArray(content)
    ? content.map(function (r) {
        const ro = {};
        if (r.size) ro.fontSize = r.size;
        if (r.bold !== undefined) ro.bold = r.bold;
        if (r.br) ro.breakLine = true;
        return { text: r.t, options: ro };
      })
    : content;
  const opt = {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: o.face || SANS,
    fontSize: o.size || BODY_SIZE,
    color: o.color || INK,
    bold: !!o.bold,
    align: o.align || 'left',
    valign: o.valign || 'top'
  };
  if (o.ls) opt.lineSpacingMultiple = o.ls;
  if (o.rot) opt.rotate = o.rot;
  if (o.wrap === false) opt.wrap = false;
  if (o.fill) opt.fill = { color: o.fill };
  s.addText(body, opt);
}

/** Straight rule / connector. */
function rule(s, o) {
  s.addShape('line', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    line: { color: o.color, width: o.width || 1, dashType: o.dash || 'solid' }
  });
}

/** Centred lettermark, used by the wordmark-style icons. */
function glyph(s, ch, x, y, w, h, c, face, sizeFactor) {
  s.addText(ch, {
    x: x - w * 0.5, y: y - h * 0.15, w: w * 2, h: h * 1.3,
    fontFace: face, fontSize: Math.round(h * (sizeFactor || 78)), bold: true,
    color: c, align: 'center', valign: 'middle', wrap: false
  });
}

/* -------------------------------------------------------------------- icons */
/* The source deck draws its pictograms with custom geometry; each one is
   rebuilt here from native shapes inside the same box the original occupied. */

const icon = {
  twitter: function (s, x, y, w, h, c) {              // bird in flight
    s.addShape('moon', { x: x, y: y + h * 0.05, w: w * 0.78, h: h * 0.9,
      fill: { color: c }, line: { type: 'none' }, rotate: 200 });
    s.addShape('triangle', { x: x + w * 0.34, y: y + h * 0.38, w: w * 0.62, h: h * 0.56,
      fill: { color: c }, line: { type: 'none' }, rotate: 215 });
  },
  instagram: function (s, x, y, w, h, c) {            // rounded square + lens
    stroke(s, 'roundRect', x, y, w, h, c, 1.5, { rectRadius: Math.min(w, h) * 0.22 });
    stroke(s, 'ellipse', x + w * 0.27, y + h * 0.3, w * 0.46, h * 0.46, c, 1.5);
    box(s, x + w * 0.66, y + h * 0.15, w * 0.16, h * 0.16, c);
  },
  tumblr:   function (s, x, y, w, h, c) { glyph(s, 't',  x, y, w, h, c, SERIF, 86); },
  linkedin: function (s, x, y, w, h, c) { glyph(s, 'in', x, y, w, h, c, SANS,  72); },
  dribbble: function (s, x, y, w, h, c) {             // ball with swooping arcs
    stroke(s, 'ellipse', x, y, w, h, c, 1.5);
    stroke(s, 'arc', x + w * 0.1, y + h * 0.15, w * 1.1, h * 0.9, c, 1.2, { rotate: 120 });
    stroke(s, 'arc', x - w * 0.15, y + h * 0.05, w, h, c, 1.2, { rotate: 20 });
  },
  check: function (s, x, y, w, h, c) {                // tick mark
    s.addShape('rect', { x: x + w * 0.1, y: y + h * 0.5, w: w * 0.44, h: h * 0.17,
      fill: { color: c }, line: { type: 'none' }, rotate: 45 });
    s.addShape('rect', { x: x + w * 0.33, y: y + h * 0.36, w: w * 0.74, h: h * 0.17,
      fill: { color: c }, line: { type: 'none' }, rotate: -50 });
  },
  bag: function (s, x, y, w, h, c) {                  // shopping bag
    stroke(s, 'roundRect', x, y + h * 0.26, w, h * 0.74, c, 1.5, { rectRadius: w * 0.1 });
    stroke(s, 'blockArc', x + w * 0.24, y, w * 0.52, h * 0.5, c, 1.5);
    rule(s, { x: x + w * 0.06, y: y + h * 0.55, w: w * 0.88, h: 0, color: c, width: 1.2 });
  },
  cart: function (s, x, y, w, h, c) {                 // shopping cart
    stroke(s, 'trapezoid', x + w * 0.18, y + h * 0.2, w * 0.78, h * 0.46, c, 1.4, { flipV: true });
    rule(s, { x: x, y: y + h * 0.2, w: w * 0.22, h: 0, color: c, width: 1.4 });
    stroke(s, 'ellipse', x + w * 0.34, y + h * 0.78, w * 0.15, h * 0.15, c, 1.2);
    stroke(s, 'ellipse', x + w * 0.68, y + h * 0.78, w * 0.15, h * 0.15, c, 1.2);
  },
  browser: function (s, x, y, w, h, c) {              // window chrome
    stroke(s, 'roundRect', x, y, w, h, c, 1.5, { rectRadius: w * 0.1 });
    rule(s, { x: x, y: y + h * 0.34, w: w, h: 0, color: c, width: 1.3 });
    for (let i = 0; i < 3; i++) {
      box(s, x + w * (0.13 + i * 0.14), y + h * 0.13, w * 0.08, h * 0.08, c, { type: 'ellipse' });
    }
  },
  camera: function (s, x, y, w, h, c) {               // camera body + lens
    stroke(s, 'roundRect', x, y + h * 0.2, w, h * 0.8, c, 2, { rectRadius: w * 0.1 });
    stroke(s, 'trapezoid', x + w * 0.28, y + h * 0.02, w * 0.44, h * 0.24, c, 2);
    stroke(s, 'ellipse', x + w * 0.3, y + h * 0.36, w * 0.4, h * 0.46, c, 2);
  },
  clock: function (s, x, y, w, h, c) {                // clock face + hands
    stroke(s, 'ellipse', x, y, w, h, c, 3);
    box(s, x + w * 0.46, y + h * 0.24, w * 0.08, h * 0.3, c);
    box(s, x + w * 0.48, y + h * 0.46, w * 0.26, h * 0.08, c);
  },
  badgeInstagram: function (s, x, y, w, h) {
    stroke(s, 'roundRect', x, y, w, h, WHITE, 1, { rectRadius: w * 0.18 });
    icon.instagram(s, x + w * 0.26, y + h * 0.26, w * 0.48, h * 0.48, WHITE);
  },
  badgeFacebook: function (s, x, y, w, h) {
    stroke(s, 'roundRect', x, y, w, h, WHITE, 1, { rectRadius: w * 0.18 });
    glyph(s, 'f', x + w * 0.3, y + h * 0.2, w * 0.4, h * 0.6, WHITE, SANS, 76);
  },
  badgeGoogle: function (s, x, y, w, h) {
    stroke(s, 'roundRect', x, y, w, h, WHITE, 1, { rectRadius: w * 0.18 });
    glyph(s, 'g+', x + w * 0.2, y + h * 0.22, w * 0.6, h * 0.56, WHITE, SANS, 66);
  }
};

/* ------------------------------------------------- icon-library grid slides */
/* Slides 39-42 are pictogram catalogue sheets: a dense lattice of small line
   icons.  The repertoire below is cycled across the lattice of each sheet.   */

const GRID_ICONS = [
  function (s, x, y, u, c) { stroke(s, 'rect', x, y, u, u * 0.8, c, 1); box(s, x, y, u, u * 0.2, c); },
  function (s, x, y, u, c) { stroke(s, 'ellipse', x, y, u, u, c, 1);
    box(s, x + u * 0.47, y + u * 0.2, u * 0.06, u * 0.34, c); box(s, x + u * 0.5, y + u * 0.47, u * 0.26, u * 0.06, c); },
  function (s, x, y, u, c) { stroke(s, 'roundRect', x + u * 0.2, y, u * 0.6, u, c, 1, { rectRadius: u * 0.1 });
    box(s, x + u * 0.36, y + u * 0.86, u * 0.28, u * 0.05, c); },
  function (s, x, y, u, c) { stroke(s, 'rect', x, y + u * 0.1, u, u * 0.62, c, 1);
    box(s, x + u * 0.25, y + u * 0.86, u * 0.5, u * 0.06, c); box(s, x + u * 0.46, y + u * 0.72, u * 0.08, u * 0.14, c); },
  function (s, x, y, u, c) { stroke(s, 'ellipse', x, y, u, u, c, 1); stroke(s, 'ellipse', x + u * 0.3, y + u * 0.3, u * 0.4, u * 0.4, c, 1); },
  function (s, x, y, u, c) { box(s, x + u * 0.1, y, u * 0.8, u, c, { type: 'triangle' }); },
  function (s, x, y, u, c) { stroke(s, 'rect', x + u * 0.12, y, u * 0.76, u, c, 1);
    for (let i = 0; i < 3; i++) box(s, x + u * 0.24, y + u * (0.22 + i * 0.22), u * (i === 2 ? 0.32 : 0.52), u * 0.06, c); },
  function (s, x, y, u, c) { stroke(s, 'ellipse', x, y, u * 0.72, u * 0.72, c, 1);
    s.addShape('rect', { x: x + u * 0.56, y: y + u * 0.62, w: u * 0.38, h: u * 0.09, fill: { color: c }, line: { type: 'none' }, rotate: 45 }); },
  function (s, x, y, u, c) { stroke(s, 'rect', x, y + u * 0.16, u, u * 0.68, c, 1);
    stroke(s, 'triangle', x, y + u * 0.16, u, u * 0.5, c, 1, { rotate: 180 }); },
  function (s, x, y, u, c) { box(s, x, y, u, u, c, { type: 'star5' }); },
  function (s, x, y, u, c) { box(s, x, y + u * 0.08, u, u * 0.86, c, { type: 'heart' }); },
  function (s, x, y, u, c) { stroke(s, 'roundRect', x, y + u * 0.22, u, u * 0.78, c, 1, { rectRadius: u * 0.1 });
    stroke(s, 'blockArc', x + u * 0.24, y, u * 0.52, u * 0.46, c, 1); },
  function (s, x, y, u, c) { stroke(s, 'diamond', x, y, u, u, c, 1); },
  function (s, x, y, u, c) { box(s, x, y, u, u, c, { type: 'pie' }); },
  function (s, x, y, u, c) { stroke(s, 'hexagon', x, y, u, u, c, 1); },
  function (s, x, y, u, c) { box(s, x, y + u * 0.42, u, u * 0.14, c); box(s, x + u * 0.43, y, u * 0.14, u, c); },
  function (s, x, y, u, c) { stroke(s, 'ellipse', x, y, u, u, c, 1); box(s, x + u * 0.2, y + u * 0.45, u * 0.6, u * 0.1, c); },
  function (s, x, y, u, c) { stroke(s, 'cloud', x, y + u * 0.12, u, u * 0.72, c, 1); },
  function (s, x, y, u, c) { box(s, x, y, u, u, c, { type: 'moon' }); },
  function (s, x, y, u, c) { stroke(s, 'trapezoid', x + u * 0.16, y + u * 0.2, u * 0.8, u * 0.44, c, 1);
    stroke(s, 'ellipse', x + u * 0.24, y + u * 0.8, u * 0.16, u * 0.16, c, 1);
    stroke(s, 'ellipse', x + u * 0.64, y + u * 0.8, u * 0.16, u * 0.16, c, 1); },
  function (s, x, y, u, c) { stroke(s, 'rect', x, y, u, u, c, 1);
    box(s, x + u * 0.16, y + u * 0.56, u * 0.15, u * 0.28, c);
    box(s, x + u * 0.43, y + u * 0.36, u * 0.15, u * 0.48, c);
    box(s, x + u * 0.7, y + u * 0.2, u * 0.15, u * 0.64, c); },
  function (s, x, y, u, c) { stroke(s, 'ellipse', x + u * 0.28, y, u * 0.44, u * 0.44, c, 1);
    stroke(s, 'blockArc', x, y + u * 0.42, u, u * 0.62, c, 1); },
  function (s, x, y, u, c) { box(s, x, y + u * 0.44, u * 0.8, u * 0.1, c);
    box(s, x + u * 0.5, y + u * 0.24, u * 0.5, u * 0.5, c, { type: 'triangle' }); },
  function (s, x, y, u, c) { stroke(s, 'roundRect', x, y, u, u, c, 1, { rectRadius: u * 0.2 });
    box(s, x + u * 0.34, y + u * 0.34, u * 0.32, u * 0.32, c, { type: 'ellipse' }); }
];

function iconGrid(s, g) {
  let n = 0;
  for (let r = 0; r < g.rows; r++) {
    const cols = (r === g.rows - 1 && g.lastRow) ? g.lastRow : g.cols;
    for (let c = 0; c < cols; c++) {
      GRID_ICONS[n % GRID_ICONS.length](s, g.x + c * g.dx, g.y + r * g.dy, g.size, g.color);
      n++;
    }
  }
}

/* ---------------------------------------------------------- device mock-ups */
/* Slides 34-36 sit on layouts carrying a laptop / phone / tablet illustration. */

function laptopMockup(s) {
  box(s, -3.7, 1.039, 8.61, 5.161, GRAPHITE, { type: 'round2SameRect' });
  box(s, -3.701, 6.2, 9.524, 0.2, SCREEN);
  box(s, -3.701, 6.4, 9.524, 0.06, 'A5A5A5');
  box(s, 0.379, 6.205, 1.317, 0.104, OFFWHITE);
  box(s, 1.02, 1.19, 0.05, 0.05, OFFWHITE, { type: 'ellipse' });
}

function phoneMockup(s) {
  box(s, 2.792, 1.312, 2.771, 5.601, SLATE, { type: 'roundRect' });
  box(s, 2.998, 1.425, 2.435, 5.289, SCREEN);
  text(s, '[image]', { x: 2.998, y: 3.87, w: 2.435, h: 0.4, size: 11, align: 'center', color: GRAY });
}

function tabletMockup(s) {
  box(s, 4.86, 2.97, 5.0, 4.53, WHITE, { type: 'roundRect' });
  stroke(s, 'roundRect', 4.86, 2.97, 5.0, 4.53, ICON, 1.5);
  box(s, 5.16, 3.58, 4.4, 3.92, '9B9B9B');
  stroke(s, 'rect', 5.16, 3.58, 4.4, 3.92, ICON, 1.5);
  box(s, 7.32, 3.26, 0.07, 0.07, ICON, { type: 'ellipse' });
}

function slide01(s) {
  s.background = { color: DARK };
  box(s, 8.254, 0, 5.08, 7.5, CREAM);
  text(s, 'Lloris', { x: 8.147, y: 2.767, w: 5.293, h: 1.966, face: SERIF, size: 115, color: GRAY, bold: true, align: 'center', valign: 'middle', rot: 90 });
  box(s, 0, 1.349, 4.667, 4.825, PEACH);
  text(s, 'Lloris Fashion ', { x: -0.397, y: 2.514, w: 4.54, h: 0.652, face: SERIF, size: 44, color: INK, align: 'right', valign: 'middle' });
  text(s, 'Modern Fashion Clean Creative Design', { x: 0.899, y: 3.839, w: 3.245, h: 0.652, size: 16, color: INK, align: 'right', valign: 'middle' });
}

function slide02(s) {
  s.background = { color: MIST };
  box(s, 7.438, 0.646, 5.521, 6.188, APRICOT);
  text(s, 'About Our\nFashion', { x: 0.774, y: 0.635, w: 4.915, h: 2.542, face: SERIF, size: 44, color: INK, bold: true, valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Every season deserves its own look transition into your fall wardrobe with these top picks', { x: 0.774, y: 5.12, w: 5.899, h: 1.784, size: 16, ls: 1.25 });
  text(s, 'PLACEHOLDER', { x: 0.844, y: 3.487, w: 3.691, h: 0.471, size: 11 });
}

function slide03(s) {
  s.background = { color: CREAM };
  text(s, 'The Team', { x: 0.604, y: 2.791, w: 2.479, h: 1.582, face: SERIF, size: 44, bold: true });
  text(s, 'PLACEHOLDER', { x: 0.833, y: 5.053, w: 4.75, h: 0.639, face: SERIF, size: 16 });
  text(s, 'Lala', { x: 6.625, y: 0.913, w: 2.458, h: 0.639, face: SERIF, size: 32 });
  text(s, 'llorisia', { x: 9.771, y: 0.913, w: 2.667, h: 0.639, face: SERIF, size: 32 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed', { x: 6.667, y: 1.785, w: 2.417, h: 0.544, size: 11, ls: 1.25 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed', { x: 9.792, y: 1.785, w: 2.417, h: 0.544, size: 11, ls: 1.25 });
  icon.tumblr(s, 8.098, 5.858, 0.194, 0.315, ICON);
  icon.instagram(s, 7.491, 5.858, 0.315, 0.315, ICON);
  icon.twitter(s, 6.945, 5.885, 0.313, 0.256, ICON);
  icon.linkedin(s, 8.585, 5.844, 0.315, 0.297, ICON);
  icon.tumblr(s, 11.176, 5.858, 0.194, 0.315, ICON);
  icon.instagram(s, 10.568, 5.858, 0.315, 0.315, ICON);
  icon.twitter(s, 10.022, 5.885, 0.313, 0.256, ICON);
  icon.linkedin(s, 11.663, 5.844, 0.315, 0.297, ICON);
}

function slide04(s) {
  s.background = { color: CREAM };
  text(s, 'Every season deserves its own look transition.', { x: 2.562, y: 0.647, w: 5.979, h: 1.043, face: SERIF, size: 28, align: 'right' });
  text(s, 'Jancuk', { x: 1.062, y: 5.331, w: 2.354, h: 0.505, face: SERIF, size: 24, align: 'center' });
  text(s, 'Anomaly', { x: 3.708, y: 5.331, w: 2.271, h: 0.505, face: SERIF, size: 24, align: 'center' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing ', { x: 1.146, y: 5.915, w: 2.125, h: 0.539, size: 11, align: 'center', ls: 1.25 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing ', { x: 3.708, y: 5.915, w: 2.271, h: 0.539, size: 11, align: 'center', ls: 1.25 });
  text(s, 'Reynhard', { x: 6.271, y: 5.331, w: 2.271, h: 0.505, face: SERIF, size: 24, align: 'center' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing', { x: 6.271, y: 5.855, w: 2.271, h: 0.539, size: 11, align: 'center', ls: 1.25 });
  text(s, 'This is our Team', { x: 9.708, y: 2.271, w: 2.583, h: 2.542, face: SERIF, size: 44, color: INK, bold: true, align: 'right', valign: 'middle' });
}

function slide05(s) {
  s.background = { color: PEACH };
  box(s, 4.897, 0, 3.54, 7.5, DARK);
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Every season deserves its own look transition into your fall wardrobe with these top picks', { x: 0.835, y: 0.704, w: 3.205, h: 1.721, size: 11, align: 'right', ls: 1.25 });
  text(s, 'Lloris – We Are  Fashion ', { x: 4.297, y: 3.43, w: 4.739, h: 0.639, face: SERIF, size: 32, color: WHITE, bold: true, align: 'center', wrap: false, rot: 270 });
  text(s, 'We Make\nFashion', { x: 9.707, y: 3.489, w: 2.712, h: 1.582, face: SERIF, size: 44, bold: true, align: 'right' });
}

function slide06(s) {
  s.background = { color: PEACH };
  box(s, 6.667, 0, 6.667, 7.5, CREAM);
  text(s, 'Collections.', { x: 0.896, y: 4.066, w: 3.83, h: 0.652, face: SERIF, size: 44, color: INK, bold: true, valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 0.896, y: 5.165, w: 4.814, h: 0.817, ls: 1.25 });
  text(s, 'Lloris – Fashion Design ', { x: 5.932, y: 3.566, w: 3.83, h: 0.652, face: SERIF, size: 24, color: INK, bold: true, align: 'center', valign: 'middle', rot: 270 });
}

function slide07(s) {
  s.background = { color: DARK };
  box(s, 6.667, 0, 6.667, 7.5, CREAM);
  text(s, 'Collections.', { x: 0.385, y: 4.202, w: 3.83, h: 0.652, face: SERIF, size: 44, color: WHITE, bold: true, valign: 'middle' });
  text(s, [{ t: 'Lorem ipsum dolor sit amet, consectetur ' }, { t: 'adipiscing', size: 11 }, { t: ' elit, sed do eiusmod tempo' }], { x: 0.488, y: 5.612, w: 4.75, h: 0.658, size: 14, color: WHITE, ls: 1.25 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur ', { x: 0.488, y: 1.365, w: 4.75, h: 0.332, size: 12, color: WHITE, ls: 1.25 });
}

function slide08(s) {
  s.background = { color: DARK };
  box(s, 0, 2.078, 13.333, 5.422, SAND);
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ', { x: 7.375, y: 0.755, w: 3.539, h: 0.831, size: 12, color: WHITE, ls: 1.25 });
  text(s, 'New Collection', { x: 0.604, y: 0.417, w: 5.146, h: 1.661, face: SERIF, size: 44, color: WHITE, bold: true, valign: 'middle' });
}

function slide09(s) {
  s.background = { color: CREAM };
  text(s, 'Lloris \nFashion.', { x: 1.125, y: 1.216, w: 1.855, h: 1.313, face: SERIF, size: 36, bold: true, wrap: false });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 7.8, y: 4.756, w: 4.971, h: 0.95, ls: 1.5 });
  text(s, 'New Collections.', { x: 4.842, y: 3.397, w: 3.622, h: 0.707, face: SERIF, size: 36, bold: true, wrap: false, rot: 270 });
}

function slide10(s) {
  s.background = { color: DARK };
  text(s, 'Influincer Colaborations.', { x: 0.917, y: 2.25, w: 3.83, h: 0.652, face: SERIF, size: 44, color: WHITE, bold: true, valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 3.985, y: 4.598, w: 2.521, h: 1.279, size: 14, color: WHITE, ls: 1.25 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 10.152, y: 4.598, w: 2.521, h: 1.279, size: 14, color: WHITE, ls: 1.25 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 10.152, y: 1.514, w: 2.521, h: 1.279, size: 14, color: WHITE, ls: 1.25 });
}

function slide11(s) {
  s.background = { color: DARK };
  icon.bag(s, 6.458, 0.786, 0.439, 0.575, WHITE);
  text(s, 'Visit Our Online Shop', { x: 2.351, y: 1.767, w: 8.652, h: 0.806, face: SERIF, size: 44, color: OFFWHITE, bold: true, align: 'center', valign: 'middle' });
}

function slide12(s) {
  s.background = { color: CREAM };
  box(s, 6.667, 5.292, 2.458, 1.583, PEACH);
  icon.bag(s, 8.279, 5.925, 0.249, 0.326, ICON);
  icon.cart(s, 7.7, 5.925, 0.334, 0.326, ICON);
  icon.cart(s, 7.166, 5.934, 0.334, 0.311, ICON);
  text(s, [{ t: 'Every season deserves its ' }, { t: 'own', size: 20 }], { x: 2.889, y: 2.42, w: 6.236, h: 0.505, size: 24, bold: true, align: 'right' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ', { x: 4.874, y: 3.557, w: 4.251, h: 0.985, size: 14, align: 'right', ls: 1.25 });
  text(s, 'Our Online Shop.', { x: 0.638, y: 4.05, w: 2.896, h: 1.661, face: SERIF, size: 44, color: INK, bold: true, align: 'right', valign: 'middle' });
}

function slide13(s) {
  s.background = { color: CREAM };
  box(s, 0.049, 0, 3.21, 7.5, DARK);
  box(s, 8.271, 1.417, 3.938, 4.688, PEACH);
  text(s, 'Social Media', { x: 5.018, y: 3.34, w: 3.34, h: 0.842, face: SERIF, size: 44, bold: true, wrap: false, rot: 270 });
  text(s, 'Find us on \nSocial Media', { x: 9.164, y: 1.833, w: 2.152, h: 1.245, size: 24, ls: 1.5 });
  icon.tumblr(s, 10.534, 3.426, 0.194, 0.315, ICON);
  icon.instagram(s, 9.926, 3.426, 0.315, 0.315, ICON);
  icon.dribbble(s, 9.323, 3.439, 0.311, 0.311, ICON);
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit,', { x: 9.143, y: 4.238, w: 2.654, h: 1.268, size: 16, ls: 1.5 });
}

function slide14(s) {
  s.background = { color: CREAM };
  text(s, 'Marketing', { x: 1.62, y: 1.034, w: 3.83, h: 0.652, face: SERIF, size: 48, color: INK, bold: true, valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 3.055, y: 3.289, w: 3.46, h: 1.237, ls: 1.25 });
  text(s, 'A', { x: 2.075, y: 2.966, w: 1.182, h: 1.145, size: 32, color: INK, align: 'center', valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 3.055, y: 5.116, w: 3.46, h: 1.237, ls: 1.25 });
  text(s, 'B', { x: 2.075, y: 4.792, w: 1.182, h: 1.145, size: 32, color: INK, align: 'center', valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 8.532, y: 1.684, w: 3.46, h: 1.237, ls: 1.25 });
  text(s, 'C', { x: 7.552, y: 1.36, w: 1.182, h: 1.145, size: 32, color: INK, align: 'center', valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 8.532, y: 3.289, w: 3.46, h: 1.237, ls: 1.25 });
  text(s, 'D', { x: 7.552, y: 2.966, w: 1.182, h: 1.145, size: 32, color: INK, align: 'center', valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, ', { x: 1.62, y: 1.742, w: 3.394, h: 0.579, size: 12, ls: 1.25 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 8.532, y: 5.116, w: 3.46, h: 1.237, ls: 1.25 });
  text(s, 'E', { x: 7.552, y: 4.792, w: 1.182, h: 1.145, size: 32, color: INK, align: 'center', valign: 'middle' });
}

function slide15(s) {
  s.background = { color: SAND };
  box(s, 6.667, 0, 3.479, 7.5, PEACH, { type: 'flowChartProcess' });
  text(s, 'Our Workflow', { x: 6.073, y: 3.243, w: 4.667, h: 0.806, face: SERIF, size: 44, color: INK, bold: true, align: 'center', valign: 'middle', rot: 90 });
  text(s, 'Fashion', { x: 1.012, y: 0.736, w: 4.667, h: 0.806, face: SERIF, size: 40, color: INK, bold: true, valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 3.51, y: 0.734, w: 2.611, h: 1.411, size: 16, align: 'justify', ls: 1.25 });
  text(s, 'Style', { x: 1.012, y: 2.965, w: 4.667, h: 0.806, face: SERIF, size: 40, color: INK, bold: true, valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 3.51, y: 2.963, w: 2.611, h: 1.411, size: 16, align: 'justify', ls: 1.25 });
  text(s, 'Design', { x: 1.012, y: 5.199, w: 4.667, h: 0.806, face: SERIF, size: 40, color: INK, bold: true, valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 3.51, y: 5.197, w: 2.611, h: 1.411, size: 16, align: 'justify', ls: 1.25 });
}

function slide16(s) {
  s.background = { color: CREAM };
  box(s, 9.793, 0, 3.54, 7.5, PEACH);
  text(s, 'Rene Maulesten', { x: 9.741, y: 2.127, w: 4.284, h: 0.639, face: SERIF, size: 32, bold: true, align: 'center', rot: 90 });
  icon.tumblr(s, 12.062, 5.308, 0.194, 0.315, ICON);
  icon.instagram(s, 11.455, 5.308, 0.315, 0.315, ICON);
  icon.twitter(s, 10.909, 5.335, 0.313, 0.256, ICON);
  icon.linkedin(s, 12.549, 5.294, 0.315, 0.297, ICON);
  text(s, 'Fashion & Style', { x: 0.999, y: 0.183, w: 3.771, h: 2.542, face: SERIF, size: 44, color: INK, bold: true, valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ', { x: 0.999, y: 2.558, w: 4.251, h: 0.985, size: 14, ls: 1.25 });
  box(s, 0.847, 5.596, 4.782, 0.102, STEEL, { type: 'roundRect' });
  box(s, 0.847, 5.596, 4.198, 0.102, OLIVE, { type: 'roundRect' });
  text(s, '90%', { x: 4.924, y: 5.175, w: 1.162, h: 0.421, size: 11, color: COAL });
  text(s, 'Design', { x: 0.653, y: 5.108, w: 1.484, h: 0.452, size: 16, color: COAL, bold: true, ls: 1.5 });
  box(s, 0.847, 4.757, 4.782, 0.102, STEEL, { type: 'roundRect' });
  box(s, 0.847, 4.757, 3.582, 0.102, OLIVE, { type: 'roundRect' });
  text(s, '87%', { x: 4.924, y: 4.336, w: 1.162, h: 0.421, size: 11, color: COAL });
  text(s, 'Fashion', { x: 0.653, y: 4.206, w: 1.81, h: 0.452, size: 16, color: COAL, bold: true, ls: 1.5 });
  box(s, 0.847, 6.501, 4.782, 0.102, STEEL, { type: 'roundRect' });
  box(s, 0.847, 6.501, 3.582, 0.102, OLIVE, { type: 'roundRect' });
  text(s, '87%', { x: 4.924, y: 6.079, w: 1.162, h: 0.421, size: 11, color: COAL });
  text(s, 'Style', { x: 0.653, y: 5.949, w: 1.81, h: 0.452, size: 16, color: COAL, bold: true, ls: 1.5 });
}

function slide17(s) {
  s.background = { color: CREAM };
  box(s, 1.146, 5.312, 4.812, 1.583, OFFWHITE);
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ', { x: 7.228, y: 2.558, w: 4.251, h: 0.985, size: 14, ls: 1.25 });
  text(s, 'Fashion & Design', { x: 7.228, y: 1.395, w: 3.771, h: 1.25, face: SERIF, size: 20, color: INK, valign: 'middle' });
  text(s, 'Rene Maulesten', { x: 7.228, y: 0.7, w: 4.148, h: 1.661, face: SERIF, size: 44, color: INK, bold: true, valign: 'middle' });
  icon.tumblr(s, 3.728, 5.953, 0.194, 0.315, ICON);
  icon.instagram(s, 3.12, 5.953, 0.315, 0.315, ICON);
  icon.twitter(s, 2.574, 5.981, 0.313, 0.256, ICON);
  icon.linkedin(s, 4.215, 5.94, 0.315, 0.297, ICON);
  box(s, 7.423, 5.885, 4.782, 0.102, STEEL, { type: 'roundRect' });
  box(s, 7.423, 5.885, 4.198, 0.102, OLIVE, { type: 'roundRect' });
  text(s, '90%', { x: 11.499, y: 5.464, w: 1.162, h: 0.286, size: 11, color: COAL });
  text(s, 'Design', { x: 7.228, y: 5.396, w: 1.484, h: 0.452, size: 16, color: COAL, bold: true, ls: 1.5 });
  box(s, 7.423, 5.046, 4.782, 0.102, STEEL, { type: 'roundRect' });
  box(s, 7.423, 5.046, 3.582, 0.102, OLIVE, { type: 'roundRect' });
  text(s, '87%', { x: 11.499, y: 4.625, w: 1.162, h: 0.286, size: 11, color: COAL });
  text(s, 'Fashion', { x: 7.228, y: 4.494, w: 1.81, h: 0.452, size: 16, color: COAL, bold: true, ls: 1.5 });
  box(s, 7.423, 6.789, 4.782, 0.102, STEEL, { type: 'roundRect' });
  box(s, 7.423, 6.789, 3.582, 0.102, OLIVE, { type: 'roundRect' });
  text(s, '87%', { x: 11.499, y: 6.368, w: 1.162, h: 0.286, size: 11, color: COAL });
  text(s, 'Marketing', { x: 7.228, y: 6.238, w: 1.81, h: 0.452, size: 16, color: COAL, bold: true, ls: 1.5 });
}

function slide18(s) {
  s.background = { color: WHITE };
  box(s, 1.167, 1.417, 11, 4.667, OLIVE, { transparency: 18, type: 'flowChartProcess' });
  text(s, [{ t: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore. ', br: true }, { t: '', size: 44 }], { x: 1.875, y: 2.453, w: 9.583, h: 2.904, face: SERIF, size: 44, color: WHITE, bold: true, align: 'center', valign: 'middle' });
}

function slide19(s) {
  s.background = { color: CREAM };
  box(s, 0, 0, 4.643, 7.5, PEACH);
  text(s, 'The New \nStyle.', { x: 0.837, y: 5.017, w: 2.477, h: 0.652, face: SERIF, size: 44, color: INK, bold: true, valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 9.72, y: 1.834, w: 3.46, h: 1.237, ls: 1.25 });
  text(s, 'Clean and \nCreative style \ndesign', { x: 0.837, y: 1.834, w: 1.967, h: 1.205, bold: true, ls: 1.25 });
  icon.check(s, 9.01, 1.865, 0.39, 0.411, ICON);
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 9.72, y: 4.106, w: 3.46, h: 1.237, ls: 1.25 });
  icon.check(s, 9.01, 4.137, 0.39, 0.411, ICON);
  text(s, 'lacus nulla ac netus nibh aliquet, porttitor ligula justo libero vivamus porttitor dolor', { x: 6.667, y: 5.996, w: 5.704, h: 0.278, size: 10.5 });
}

function slide20(s) {
  s.background = { color: PEACH };
  box(s, 6.667, 0, 6.667, 7.5, CREAM);
  text(s, 'Lloris Fashion\nStyle.', { x: 7.29, y: 2.649, w: 4.186, h: 0.652, face: SERIF, size: 44, color: INK, bold: true, valign: 'middle' });
  text(s, [{ t: 'Lorem ipsum dolor sit amet, consectetur ' }, { t: 'adipiscing', size: 12 }, { t: ' elit, sed do eiusmod tempo' }], { x: 7.29, y: 4.23, w: 5.583, h: 0.738, size: 16, bold: true, ls: 1.25 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 8.143, y: 5.651, w: 3.333, h: 0.585, size: 12, ls: 1.25 });
  icon.check(s, 7.436, 5.809, 0.25, 0.264, BRONZE);
}

function slide21(s) {
  s.background = { color: PEACH };
  box(s, 7.438, 0.646, 5.521, 2.771, OLIVE);
  text(s, 'New Style.', { x: 7.625, y: 2.268, w: 4.383, h: 0.652, face: SERIF, size: 44, color: WHITE, bold: true, valign: 'middle' });
  text(s, 'Every season deserves its own look transition.', { x: 8.479, y: 4.188, w: 3.771, h: 0.707 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ', { x: 8.479, y: 5.323, w: 4.251, h: 0.837, size: 12, ls: 1.25 });
}

function slide22(s) {
  s.background = { color: GRAY };
  box(s, 10.047, 3.587, 2.477, 3.306, OLIVE);
  text(s, 'Lorem ipsum dolor sit amet, Lloris Modern\nDesign.', { x: 10.301, y: 4.447, w: 2.143, h: 2.313, color: WHITE, ls: 1.5 });
  text(s, 'Lloris  Style.', { x: 9.343, y: 0.716, w: 3.243, h: 2.442, face: SERIF, size: 72, color: WHITE, bold: true, valign: 'middle' });
}

function slide23(s) {
  s.background = { color: DARK };
  box(s, 10.148, 2.974, 2.477, 3.306, PEACH);
  text(s, 'Our Service', { x: 0.866, y: 1.172, w: 3.83, h: 0.652, face: SERIF, size: 44, color: WHITE, bold: true, valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 1.402, y: 2.887, w: 3.46, h: 1.237, color: WHITE, ls: 1.25 });
  text(s, 'A', { x: 0.422, y: 2.563, w: 1.182, h: 1.145, size: 32, color: WHITE, align: 'center', valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 6.069, y: 2.865, w: 3.46, h: 1.237, color: WHITE, ls: 1.25 });
  text(s, 'B', { x: 5.089, y: 2.541, w: 1.182, h: 1.145, size: 32, color: WHITE, align: 'center', valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 1.402, y: 5.02, w: 3.46, h: 1.237, color: WHITE, ls: 1.25 });
  text(s, 'C', { x: 0.422, y: 4.696, w: 1.182, h: 1.145, size: 32, color: WHITE, align: 'center', valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 6.069, y: 5.02, w: 3.46, h: 1.237, color: WHITE, ls: 1.25 });
  text(s, 'D', { x: 5.089, y: 4.696, w: 1.182, h: 1.145, size: 32, color: WHITE, align: 'center', valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, ', { x: 9.528, y: 1.291, w: 3.46, h: 0.746, size: 16, color: WHITE, ls: 1.25 });
  text(s, 'This is our New Clean and \nCreative style \ndesign', { x: 10.508, y: 3.9, w: 1.903, h: 1.616, color: GRAY, ls: 1.25 });
}

function slide24(s) {
  s.background = { color: CREAM };
  box(s, 2.061, 0.957, 5.521, 5.564, DARK);
  text(s, 'Lloris - Fashion ', { x: -1.518, y: 3.424, w: 5.155, h: 0.652, face: SERIF, size: 44, color: CHARCOAL, bold: true, align: 'center', valign: 'middle', rot: 270 });
  text(s, 'Lorem ipsum dolor sit amet, Disney Modern Clean Creative design.', { x: 2.953, y: 1.809, w: 2.585, h: 2.373, color: OFFWHITE, ls: 1.5 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing ', { x: 2.953, y: 4.836, w: 3.039, h: 0.579, size: 12, color: OFFWHITE, ls: 1.25 });
}

function slide25(s) {
  s.background = { color: PEACH };
  text(s, 'Lloris – Fashion Template ', { x: 9.067, y: 3.566, w: 5.199, h: 0.652, face: SERIF, size: 32, color: INK, bold: true, align: 'center', valign: 'middle', rot: 90 });
  box(s, 1.19, 1.008, 4.984, 5.484, DARK);
  text(s, 'Lloris Fashion.', { x: 1.407, y: 2.649, w: 4.186, h: 0.652, face: SERIF, size: 44, color: WHITE, bold: true, valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur elit, sed do eiusmod tempo', { x: 1.976, y: 4.309, w: 3.783, h: 0.658, size: 14, color: WHITE, align: 'right', ls: 1.25 });
}

function slide26(s) {
  s.background = { color: CREAM };
  box(s, 1.159, 0.62, 3.206, 6.285, DARK);
  text(s, 'Lloris– Fashion Template ', { x: 0.385, y: 3.566, w: 4.754, h: 0.652, face: SERIF, size: 32, color: WHITE, bold: true, align: 'center', valign: 'middle', rot: 270 });
  icon.browser(s, 4.809, 3.565, 0.389, 0.394, ICON);
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 4.809, y: 4.686, w: 2.762, h: 1.584, ls: 1.25 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 4.809, y: 1.255, w: 2.762, h: 1.584, ls: 1.25 });
}

function slide27(s) {
  s.background = { color: CREAM };
  box(s, 2.508, 0.651, 4.984, 6.238, PEACH);
  text(s, 'Lloris – Fashion Template ', { x: -1.714, y: 3.444, w: 5.84, h: 0.652, face: SERIF, size: 32, color: INK, bold: true, align: 'center', valign: 'middle', rot: 270 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 3.159, y: 1.382, w: 2.714, h: 1.919, ls: 1.5 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur elit, sed do eiusmod tempo', { x: 3.159, y: 5.194, w: 3.651, h: 0.658, size: 14, ls: 1.25 });
}

function slide28(s) {
  s.background = { color: DARK };
  box(s, 0, 0, 5.872, 7.5, CREAM);
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 2.198, y: 1.215, w: 2.991, h: 1.26, size: 16, ls: 1.5 });
  text(s, 'Our Solutions', { x: 9.544, y: 3.286, w: 3.664, h: 0.842, face: SERIF, size: 44, color: OFFWHITE, bold: true, align: 'center', wrap: false, rot: 90 });
  text(s, 'One.', { x: 0.579, y: 1.215, w: 4.186, h: 0.652, face: SERIF, size: 44, color: INK, bold: true, valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 2.198, y: 4.357, w: 2.991, h: 1.26, size: 16, ls: 1.5 });
  text(s, 'Two.', { x: 0.579, y: 4.357, w: 4.186, h: 0.652, face: SERIF, size: 44, color: INK, bold: true, valign: 'middle' });
}

function slide29(s) {
  s.background = { color: PEACH };
  box(s, 0, 0, 3.479, 7.5, CREAM, { type: 'flowChartProcess' });
  text(s, 'Vertical Timeline', { x: -2.01, y: 3.347, w: 7.5, h: 0.806, face: SERIF, size: 44, color: INK, bold: true, align: 'center', valign: 'middle', rot: 270 });
  text(s, 'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo libero vivamus porttitor dolor, conubia mollit. Sapien nam suspendisse, tincidunt eget ante tincidunt, eros in auctor fringilla praesent at diam. ', { x: 4.481, y: 3.471, w: 2.577, h: 1.111, size: 8, color: ICON, align: 'right', ls: 1.5 });
  text(s, 'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo libero vivamus porttitor dolor, conubia mollit. Sapien nam suspendisse, tincidunt eget ante tincidunt, eros in auctor fringilla praesent at diam. mollit. Sapien nam suspendisse, tincidunt eget ante tincidunt, eros in auctor fringilla praesent at diam. ', { x: 8.876, y: 1.356, w: 2.577, h: 1.488, size: 8, color: ICON, ls: 1.5 });
  text(s, 'Design and Style', { x: 4.489, y: 3.182, w: 2.569, h: 0.303, size: 12, color: NEARBLACK, bold: true, align: 'right' });
  text(s, 'Fashion ', { x: 8.876, y: 0.783, w: 2.569, h: 0.303, size: 12, color: NEARBLACK, bold: true });
  text(s, 'STARTING POINT', { x: 7.103, y: 0.887, w: 1.62, h: 0.296, size: 12, color: WHITE, align: 'center', valign: 'middle', fill: DARK });
  rule(s, { x: 7.913, y: 1.183, w: 0, h: 6.317, color: SILVER, width: 2.25, dash: 'sysDot' });
  text(s, '12 March 2020', { x: 7.317, y: 3.22, w: 1.191, h: 0.405, size: 10, color: WHITE, align: 'center', valign: 'middle', fill: DARK });
  text(s, 'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo libero vivamus porttitor dolor, conubia mollit. Sapien nam suspendisse, tincidunt eget ante tincidunt, eros in auctor fringilla praesent at diam. ', { x: 8.868, y: 5.673, w: 2.577, h: 1.111, size: 8, color: ICON, ls: 1.5 });
  text(s, 'Fashion and Design', { x: 8.876, y: 5.385, w: 2.569, h: 0.303, size: 12, color: NEARBLACK, bold: true });
  text(s, '12 March 2021', { x: 7.317, y: 5.423, w: 1.191, h: 0.405, size: 10, color: WHITE, align: 'center', valign: 'middle', fill: DARK });
}

function slide30(s) {
  s.background = { color: PEACH };
  text(s, 'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo libero vivamus porttitor dolor, conubia mollit. Sapien nam suspendisse, tincidunt eget ante tincidunt, eros in auctor fringilla', { x: 8.821, y: 1.356, w: 2.577, h: 1.111, size: 8, color: NEARBLACK, ls: 1.5 });
  text(s, 'NULLA NETUS NIBH ALIQUET PORTTITOR', { x: 8.821, y: 0.783, w: 2.569, h: 0.505, size: 12, color: NEARBLACK });
  rule(s, { x: 7.865, y: -0.139, w: 0, h: 5.327, color: SILVER, width: 2.25, dash: 'sysDot' });
  text(s, '12 March 2021', { x: 7.27, y: 0.783, w: 1.191, h: 0.405, size: 10, color: WHITE, align: 'center', valign: 'middle', fill: DARK });
  text(s, 'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo libero vivamus porttitor dolor, conubia mollit. Sapien nam suspendisse, tincidunt eget ante tincidunt, eros in auctor fringilla praesent at diam. ', { x: 8.821, y: 5.464, w: 2.577, h: 1.111, size: 8, color: NEARBLACK, ls: 1.5 });
  text(s, 'NULLA NETUS NIBH', { x: 8.829, y: 5.176, w: 2.569, h: 0.303, size: 12, color: NEARBLACK });
  box(s, 7.037, 5.188, 1.657, 0.405, DARK);
  icon.check(s, 7.74, 5.259, 0.25, 0.264, WHITE);
}

function slide31(s) {
  s.background = { color: PEACH };
  box(s, 0.816, 0.618, 11.701, 6.285, DARK);
  icon.camera(s, 6.354, 3.478, 0.624, 0.557, WHITE);
  text(s, 'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, ', { x: 7.236, y: 2.401, w: 4.143, h: 2.447, face: SERIF, size: 32, color: WHITE, align: 'right', ls: 1.5 });
}

function slide32(s) {
  s.background = { color: CREAM };
  box(s, 5.896, 0.646, 1.583, 6.219, PEACH);
  text(s, 'Design Concept', { x: 4.495, y: 3.334, w: 4.344, h: 0.842, face: SERIF, size: 44, bold: true, align: 'center', rot: 270 });
}

function slide33(s) {
  s.background = { color: CREAM };
  box(s, 6.667, 3.75, 6.667, 3.75, DARK);
  text(s, 'Design Concept.', { x: 7.8, y: 1.308, w: 3.822, h: 0.774, face: SERIF, size: 40, wrap: false });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 7.8, y: 4.401, w: 4.971, h: 0.856, size: 16, color: WHITE, ls: 1.5 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 7.8, y: 5.984, w: 4.971, h: 0.856, size: 16, color: WHITE, ls: 1.5 });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit,', { x: 0.848, y: 5.594, w: 4.971, h: 1.245, size: 24, ls: 1.5 });
}

function slide34(s) {
  s.background = { color: DARK };
  laptopMockup(s);
  box(s, 6.667, 0, 6.667, 7.5, CREAM, { type: 'flowChartProcess' });
  text(s, 'Destop App.', { x: 3.578, y: 3.474, w: 4.667, h: 0.806, face: SERIF, size: 44, color: WHITE, bold: true, align: 'center', valign: 'middle', rot: 270 });
  text(s, 'A. ', { x: 7.797, y: 1.056, w: 4.667, h: 0.806, size: 40, color: INK, valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 9.307, y: 1.019, w: 3.509, h: 1.237, ls: 1.25 });
  text(s, 'B. ', { x: 7.797, y: 2.977, w: 4.667, h: 0.806, size: 40, color: INK, valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 9.307, y: 2.94, w: 3.509, h: 1.237, ls: 1.25 });
  text(s, 'C. ', { x: 7.797, y: 4.942, w: 4.667, h: 0.806, size: 40, color: INK, valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 9.307, y: 4.905, w: 3.509, h: 1.237, ls: 1.25 });
}

function slide35(s) {
  s.background = { color: KHAKI };
  phoneMockup(s);
  box(s, 6.667, 0, 6.667, 7.5, CREAM, { type: 'flowChartProcess' });
  text(s, 'Social Media', { x: -1.264, y: 3.243, w: 4.667, h: 0.806, face: SERIF, size: 44, color: WHITE, bold: true, align: 'center', valign: 'middle', rot: 270 });
}

function slide36(s) {
  s.background = { color: CREAM };
  tabletMockup(s);
  box(s, 10.619, 0, 2.714, 7.5, DARK);
  text(s, 'Tablet App.', { x: 10.061, y: 3.828, w: 3.83, h: 0.652, face: SERIF, size: 44, color: WHITE, bold: true, valign: 'middle', rot: 90 });
  text(s, 'New.', { x: 4.153, y: 0.987, w: 1.574, h: 0.652, face: SERIF, size: 32, color: INK, bold: true, valign: 'middle' });
  rule(s, { x: 1.435, y: 3.611, w: 2.718, h: 0, color: DARK, width: 2.5 });
  text(s, 'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo libero vivamus porttitor dolor, conubia mollit. Sapien nam suspendisse, tincidunt eget ante tincidunt, eros in auctor fringilla praesent at diam. In et quam est eget mi. Pellentesque nunc orci eu enim, eget in fringilla vitae, et eros praesent dolor porttitor. Lacinia lectus nonummy, accumsan mauris in sed justotincidunt, eros in auctor fringilla praesent at diam. ', { x: 1.435, y: 3.876, w: 2.847, h: 2.373, size: 9, color: COAL, align: 'justify', ls: 1.5 });
  text(s, 'Tablet Device', { x: 1.305, y: 2.869, w: 2.976, h: 1.007, size: 32, color: COAL });
  text(s, 'Our Now.', { x: 1.305, y: 1.251, w: 1.904, h: 0.652, face: SERIF, size: 44, color: INK, bold: true, valign: 'middle' });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', { x: 5.353, y: 1.859, w: 4.106, h: 0.658, size: 14, ls: 1.25 });
}

function slide37(s) {
  s.background = { color: CREAM };
  box(s, 8.563, 0.021, 4.77, 7.479, DARK);
  text(s, 'Contact & Social Media', { x: 3.576, y: 3.34, w: 6.052, h: 0.842, face: SERIF, size: 44, bold: true, wrap: false, rot: 270 });
  text(s, 'New York, 5 Ave 321', { x: 9.995, y: 4.544, w: 2.152, h: 0.415, size: 14, color: WHITE, ls: 1.5 });
  text(s, 'Address :', { x: 9.018, y: 3.884, w: 1.569, h: 0.505, size: 24, color: WHITE, wrap: false });
  icon.badgeFacebook(s, 10.565, 2.475, 0.767, 0.767, WHITE);
  icon.badgeGoogle(s, 11.681, 2.475, 0.767, 0.767, WHITE);
  icon.badgeInstagram(s, 9.448, 2.475, 0.767, 0.767, WHITE);
  text(s, '+01 10345 678', { x: 9.995, y: 4.959, w: 2.152, h: 0.415, size: 14, color: WHITE, ls: 1.5 });
  text(s, 'www.llorisfashion.com', { x: 9.995, y: 5.438, w: 2.324, h: 0.408, size: 14, color: WHITE, ls: 1.5 });
}

function slide38(s) {
  s.background = { color: WHITE };
  box(s, 4.646, 1.25, 4.042, 5, PEACH);
  icon.clock(s, 6.375, 2.365, 0.583, 0.588, ICON);
  text(s, 'Break Time', { x: 4.917, y: 3.75, w: 3.5, h: 0.652, face: SERIF, size: 44, color: INK, bold: true, align: 'center', valign: 'middle' });
}
function slide39(s) {
  s.background = { color: CREAM };
  iconGrid(s, { x: 1.146, y: 1.806, dx: 0.595, dy: 0.597, size: 0.259, cols: 19, rows: 7, color: ICON });
}

function slide40(s) {
  s.background = { color: CREAM };
  iconGrid(s, { x: 1.26, y: 1.735, dx: 0.54, dy: 0.538, size: 0.298, cols: 20, rows: 8, color: ICON });
}

function slide41(s) {
  s.background = { color: DARK };
  iconGrid(s, { x: 1.686, y: 1.883, dx: 0.551, dy: 0.553, size: 0.28, cols: 19, rows: 7, lastRow: 18, color: WHITE });
}

function slide42(s) {
  s.background = { color: DARK };
  iconGrid(s, { x: 1.82, y: 1.192, dx: 0.551, dy: 0.548, size: 0.294, cols: 18, rows: 10, lastRow: 6, color: WHITE });
}

/* --------------------------------------------------------------------- build */

const SLIDES = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
  slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
  slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24,
  slide25, slide26, slide27, slide28, slide29, slide30, slide31, slide32,
  slide33, slide34, slide35, slide36, slide37, slide38, slide39, slide40,
  slide41, slide42
];

function build() {
  const pres = new pptxgen();
  pres.defineLayout({ name: 'WIDE16x9', width: 13.3333, height: 7.5 });
  pres.layout = 'WIDE16x9';
  pres.author = 'Lloris Fashion';
  pres.title = 'Lloris Fashion';
  SLIDES.forEach(function (fn) { fn(pres.addSlide()); });
  return pres.writeFile({
    fileName: path.join(__dirname, '168aa179-6900-4175-ae0d-38b72b804700_grok_final.pptx')
  });
}

build().then(function (f) { console.log('wrote ' + f); })
       .catch(function (e) { console.error(e); process.exit(1); });
