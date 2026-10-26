// Recreates the "Yanti" ballet presentation template (36 slides, 13.33 x 7.5 in).
// Raster photos in the original are replaced with programmatic placeholders.
'use strict';
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette & type
const PURPLE  = 'A537C4';   // theme accent 1
const INK     = '262626';   // near-black headings
const SLATE   = '555555';
const BODY    = '404040';   // body copy
const GREY    = 'D9D9D9';   // rails, dividers
const MIDGREY = 'A6A6A6';
const WHITE   = 'FFFFFF';

const SERIF = 'Playfair Display';  // headings / labels
const SANS  = 'Open Sans';         // body copy

// Card drop shadow used by every floating white panel.
const SHADOW = { type: 'outer', color: GREY, opacity: 0.4, blur: 15, offset: 0, angle: 0 };

// ---------------------------------------------------------------- repeated copy
const H_A     = 'Ballet Is A Dance Executed ';
const H_B     = 'By The Human Soul';
const H_C     = 'Ballet Is A Dance Executed By The ';
const H_D     = 'Human Soul';
const H_FULL  = 'Ballet Is A Dance Executed By The Human Soul';
const H_E     = 'Ballet Is A Dance Executed By The Human ';
const H_F     = 'Soul';
const SIMPLE  = 'Simple Text Here';
const SHORT   = 'Lorem ipsum dolor sit amet, consectet adipiscing';
const BODY1   =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore eta dolore  magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nullai excepteur sint occaecat cupidatat non proident, sunt in culpa qui oficia.';
const BODY2   =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore eta dolore  magna aliqua. Ut enim ad minim veniam, quis nostrud exercit ullamco laboris nisi ut aliquip commodo consequat. Duis aute irure dolor in reprehenderiti voluptate velit esse cillum dolore eu fugiat nullai excepteur sint occaecat.';
const BODY3   =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore eta dolore  magna aliqua. Ut enim adine minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip commodo consequat. Duis aute irure dolor in reprehenderit voluptate dolore eu fugiat nullai excepteur sint occaecat cupidatat non proident, sunt in culpa qui oficia.';
const BODY4   =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore eta dolore  magna aliqua ute enim minim veniam quis nostrud exercitation ullamco laboris nisi aliquip commodo consequat. Duis aute irure dol reprehenderit in voluptate velit esse cillum dolore eu fugiat nullai excepteur sint occaecat cupidatat non.';
const BODY5   =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididu labore eta dolore  magna aliqua ute enim minim veniam quis nostrud exercitatin ullamco laboris nisi aliquip commodo sint occaecat cupidatat non provident.';
const BODY6   =
  'Lorem ipsum dolor sit amet, consect adipiscing elit, sed do eiusmod tempor incididu labore eta dolore  magna aliqua ute enim minim';
const BODY7   =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididu labore eta dolore  magna in aliqua ute enim minim veniam quis nostrud exercit ullamco laboris nisi aliquip commodo sint occaecat.';
const BODY8   =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididu labore eta dolore';
const BODY9   =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore eta dol  magna ali enim ad minim veniam, quis nostrud exercit ullamco labori nisi ut aliquip commodo.';
const BODY10  =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eiusmo tempor incididunt ut labore eta dolore  magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip commo consequat. Duis aute irure dolor in reprehenderit volupta velit.';
const BODY11  =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eiusmo tempor incididunt ut labore eta dolore  magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip commo consequat. Duis irurete dolor in reprehenderit volupta velit.';
const BODY12  =
  'Lorem ipsum dolor sit amet, consect adipiscing eiusmod tempor';
const BODY13  =
  'Lorem ipsum dolor sit amet, consectet adipiscing elit, sed do eiusmod tempor incididunt ut labore eta dolore  magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip commodo consequat. Duis aute irure dolor in reprehenderit volupta velit esse cillum dolore eu fugiat nullai excepteur sint occaecat.';
const BODY14  =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore eta dolore  magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip commodo consequat. Duis aute irure dolor in reprehenderit volupta velit esse cillum dolore eu fugiat nullai excepteur sint occaecat.';
const BODY15  =
  'Lorem ipsum dolor sit amet, consectet adipiscing elit, sed do eiusmod tempor incididu labore et dolore  magna in aliqua ute enim minim veniam quis nostrud exercit ullamco laboris.';

// ---------------------------------------------------------------- primitives

// Plain filled rectangle / preset shape.
function box(sl, o) {
  sl.addShape(o.shape || 'rect', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: o.fill, line: { type: 'none' },
    rotate: o.rotate, rectRadius: o.rectRadius,
  });
}

// Floating white panel with the deck's soft drop shadow.
function card(sl, o) {
  sl.addShape(o.shape || 'rect', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: o.fill || { color: WHITE }, line: { type: 'none' },
    rotate: o.rotate, rectRadius: o.rectRadius,
    shadow: Object.assign({}, SHADOW, {
      offset: o.offset || 0, angle: o.angle || 0,
    }),
  });
}

// Two-tone serif headline. `parts` is [[text, color], ...].
function heading(sl, parts, o) {
  sl.addText(parts.map(([text, color]) => ({ text, options: { color } })), {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: SERIF, fontSize: o.fontSize, bold: true,
    align: o.align || 'left', valign: o.valign || 'top',
    charSpacing: o.charSpacing, margin: o.margin,
    shape: o.shape, fill: o.fill, rectRadius: o.rectRadius,
    line: o.fill ? { type: 'none' } : undefined,
  });
}

// Open Sans body copy (1.5 line spacing by default).
function paragraph(sl, text, o) {
  sl.addText(text, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: SANS, fontSize: o.fontSize || 12, color: o.color || BODY,
    lineSpacingMultiple: o.lineSpacingMultiple === undefined ? 1.5 : o.lineSpacingMultiple,
    align: o.align || 'left', valign: o.valign || 'top', margin: o.margin,
  });
}

// Playfair label / caption (bold unless told otherwise).
function label(sl, text, o) {
  sl.addText(text, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: SERIF, fontSize: o.fontSize || 12,
    color: o.color || WHITE, bold: o.bold === undefined ? true : o.bold,
    lineSpacingMultiple: o.lineSpacingMultiple,
    align: o.align || 'left', valign: o.valign || 'top',
    charSpacing: o.charSpacing, margin: o.margin,
    shape: o.shape, fill: o.fill, rectRadius: o.rectRadius,
    line: o.fill ? { type: 'none' } : undefined,
  });
}

// ---------------------------------------------------------------- composites

// Big number + "Simple Text Here" caption, as used across the stat strips.
function stat(sl, x, y, value, o) {
  o = o || {};
  const align = o.align || 'left';
  const color = o.color || WHITE;
  const vw = o.w === undefined ? 0.8574 : o.w;
  const cw = o.cw === undefined ? 1.5391 : o.cw;
  const cx = align === 'right' ? x + vw - cw : x;
  label(sl, value, { x, y, w: vw, h: 0.3702, fontSize: 16, color, align });
  label(sl, SIMPLE, { x: cx, y: y + 0.3847, w: cw, h: 0.3029, fontSize: 12, color, bold: false, align });
}

// Bold serif tile title with a two-line Open Sans blurb underneath.
function tile(sl, title, body, o) {
  const align = o.align || 'left';
  const tw = 1.8206;
  const tx = align === 'center' ? o.x + (o.w - tw) / 2 : (align === 'right' ? o.x + o.w - tw : o.x);
  label(sl, title, { x: tx, y: o.y, w: tw, h: 0.3366, fontSize: 14, color: o.color || INK, align });
  paragraph(sl, body, { x: o.x, y: o.y + 0.3881, w: o.w, h: o.bh || 0.6746, color: o.bodyColor, align });
}

// ---------------------------------------------------------------- decorations

// The big "L" bracket on slide 7 (thin left leg + thin bottom foot).
function cornerBracket(sl, o) {
  const t = 0.0805;  // stroke thickness as a fraction of the square
  sl.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h, fill: o.fill, line: { type: 'none' },
    points: [
      { x: 0, y: 0, moveTo: true },
      { x: o.w * t, y: 0 },
      { x: o.w * t, y: o.h * 0.9262 },
      { x: o.w, y: o.h * 0.9262 },
      { x: o.w, y: o.h },
      { x: 0, y: o.h },
      { close: true },
    ],
  });
}

// Ring-shaped map pin (slide 34 timeline): teardrop outline with a hole.
function mapPin(sl, o) {
  const ring = (cx, cy, rx, ry) => ([
    { x: cx - rx, y: cy, moveTo: true },
    { x: cx, y: cy - ry, curve: { type: 'cubic', x1: cx - rx, y1: cy - ry * 0.552, x2: cx - rx * 0.552, y2: cy - ry } },
    { x: cx + rx, y: cy, curve: { type: 'cubic', x1: cx + rx * 0.552, y1: cy - ry, x2: cx + rx, y2: cy - ry * 0.552 } },
    { x: cx, y: cy + ry, curve: { type: 'cubic', x1: cx + rx, y1: cy + ry * 0.552, x2: cx + rx * 0.552, y2: cy + ry } },
    { x: cx - rx, y: cy, curve: { type: 'cubic', x1: cx - rx * 0.552, y1: cy + ry, x2: cx - rx, y2: cy + ry * 0.552 } },
    { close: true },
  ]);
  const w = o.w, h = o.h;
  sl.addShape('custGeom', {
    x: o.x, y: o.y, w, h, fill: o.fill, line: { type: 'none' },
    points: [
      // outer teardrop: circular head that tapers to a point at the bottom
      { x: w * 0.497, y: 0, moveTo: true },
      { x: w * 0.994, y: h * 0.444, curve: { type: 'cubic', x1: w * 0.772, y1: 0, x2: w * 0.990, y2: h * 0.199 } },
      { x: w * 0.497, y: h, curve: { type: 'cubic', x1: w, y1: h * 0.799, x2: w * 0.515, y2: h } },
      { x: 0, y: h * 0.444, curve: { type: 'cubic', x1: w * 0.478, y1: h, x2: 0, y2: h * 0.762 } },
      { x: w * 0.497, y: 0, curve: { type: 'cubic', x1: 0, y1: h * 0.199, x2: w * 0.222, y2: 0 } },
      { close: true },
      // inner hole
      ...ring(w * 0.497, h * 0.441, w * 0.353, h * 0.315),
    ],
  });
}

// Dotted connector between two map pins.
function dottedTrack(sl, o) {
  sl.addShape('line', {
    x: o.x, y: o.y + o.h / 2, w: o.w, h: 0,
    line: { color: o.fill.color, width: 4, dashType: 'sysDot' },
  });
}

// Small solid triangle terminating a dotted connector.
function arrowHead(sl, o) {
  sl.addShape('triangle', {
    x: o.x, y: o.y, w: o.w, h: o.h, rotate: 90,
    fill: o.fill, line: { type: 'none' },
  });
}

// Overlapping circle of the slide-35 "venn" infographic.
function venn(sl, text, o) {
  sl.addShape('ellipse', { x: o.x, y: o.y, w: o.w, h: o.h, fill: o.fill, line: { type: 'none' } });
  label(sl, text, {
    x: o.x, y: o.y, w: o.w, h: o.h, fontSize: 14, color: WHITE, bold: false,
    align: 'center', valign: 'middle',
  });
}

// ---------------------------------------------------------------- icon set
// The reference deck uses purple ballet line-art PNGs; these are the same
// motifs drawn from native shapes so the script stays free of embedded images.

const icon = {
  // dancer at the barre, one leg lifted
  dancer(sl, x, y, sz, color) {
    const f = { color }, ln = { type: 'none' };
    const X = u => x + sz * u, Y = v => y + sz * v, S = u => sz * u;
    sl.addShape('rect',     { x: X(0.14), y: Y(0.00), w: S(0.07), h: S(0.30), fill: f, line: ln, rotate: -12 });
    sl.addShape('ellipse',  { x: X(0.34), y: Y(0.10), w: S(0.19), h: S(0.19), fill: f, line: ln });
    sl.addShape('rect',     { x: X(0.36), y: Y(0.27), w: S(0.16), h: S(0.22), fill: f, line: ln });
    sl.addShape('rect',     { x: X(0.46), y: Y(0.26), w: S(0.42), h: S(0.07), fill: f, line: ln, rotate: 16 });
    sl.addShape('rect',     { x: X(0.52), y: Y(0.34), w: S(0.30), h: S(0.06), fill: f, line: ln, rotate: -22 });
    sl.addShape('triangle', { x: X(0.20), y: Y(0.45), w: S(0.50), h: S(0.20), fill: f, line: ln });
    sl.addShape('rect',     { x: X(0.15), y: Y(0.54), w: S(0.30), h: S(0.07), fill: f, line: ln, rotate: -30 });
    sl.addShape('rect',     { x: X(0.38), y: Y(0.62), w: S(0.12), h: S(0.24), fill: f, line: ln });
    sl.addShape('rect',     { x: X(0.00), y: Y(0.76), w: S(0.78), h: S(0.06), fill: f, line: ln });
    sl.addShape('rect',     { x: X(0.38), y: Y(0.84), w: S(0.12), h: S(0.16), fill: f, line: ln });
  },

  // dancer in a tutu with both arms raised
  arms(sl, x, y, sz, color) {
    const f = { color }, ln = { type: 'none' };
    const X = u => x + sz * u, Y = v => y + sz * v, S = u => sz * u;
    sl.addShape('rect',     { x: X(0.12), y: Y(0.00), w: S(0.07), h: S(0.26), fill: f, line: ln, rotate: 12 });
    sl.addShape('rect',     { x: X(0.81), y: Y(0.00), w: S(0.07), h: S(0.26), fill: f, line: ln, rotate: -12 });
    sl.addShape('ellipse',  { x: X(0.36), y: Y(0.01), w: S(0.26), h: S(0.24), fill: f, line: ln });
    sl.addShape('rect',     { x: X(0.20), y: Y(0.20), w: S(0.60), h: S(0.07), fill: f, line: ln });
    sl.addShape('rect',     { x: X(0.38), y: Y(0.24), w: S(0.22), h: S(0.22), fill: f, line: ln });
    sl.addShape('triangle', { x: X(0.06), y: Y(0.42), w: S(0.86), h: S(0.22), fill: f, line: ln });
    sl.addShape('rect',     { x: X(0.32), y: Y(0.62), w: S(0.34), h: S(0.10), fill: f, line: ln });
    sl.addShape('rect',     { x: X(0.43), y: Y(0.70), w: S(0.11), h: S(0.24), fill: f, line: ln });
    sl.addShape('ellipse',  { x: X(0.38), y: Y(0.90), w: S(0.21), h: S(0.10), fill: f, line: ln });
  },

  // pair of pointe shoes drawn as outlines
  shoes(sl, x, y, sz, color) {
    const ln = { color, width: 1.25 }, hollow = { type: 'none' };
    const X = u => x + sz * u, Y = v => y + sz * v, S = u => sz * u;
    // front-facing slipper: leg warmer over a rounded toe box
    sl.addShape('custGeom', { x: X(0.02), y: Y(0), w: S(0.30), h: sz, fill: hollow, line: ln, points: [
      { x: S(0.02), y: 0, moveTo: true },
      { x: S(0.28), y: 0 },
      { x: S(0.25), y: S(0.42) },
      { x: S(0.30), y: S(0.68), curve: { type: 'cubic', x1: S(0.30), y1: S(0.50), x2: S(0.30), y2: S(0.60) } },
      { x: S(0.15), y: S(1.00), curve: { type: 'cubic', x1: S(0.30), y1: S(0.88), x2: S(0.24), y2: S(1.00) } },
      { x: S(0.00), y: S(0.68), curve: { type: 'cubic', x1: S(0.06), y1: S(1.00), x2: S(0.00), y2: S(0.88) } },
      { x: S(0.05), y: S(0.42), curve: { type: 'cubic', x1: S(0.00), y1: S(0.60), x2: S(0.00), y2: S(0.50) } },
      { close: true },
    ]});
    sl.addShape('line', { x: X(0.05), y: Y(0.44), w:  S(0.24), h: S(0.16), line: ln });
    sl.addShape('line', { x: X(0.29), y: Y(0.44), w: -S(0.24), h: S(0.16), line: ln });
    // side-on slipper, toe pointed, with ankle ribbons
    sl.addShape('custGeom', { x: X(0.34), y: Y(0.18), w: S(0.66), h: S(0.62), fill: hollow, line: ln, points: [
      { x: 0, y: S(0.30), moveTo: true },
      { x: S(0.20), y: 0,       curve: { type: 'cubic', x1: 0,        y1: S(0.10), x2: S(0.08), y2: 0 } },
      { x: S(0.42), y: S(0.20), curve: { type: 'cubic', x1: S(0.30),  y1: 0,       x2: S(0.36), y2: S(0.10) } },
      { x: S(0.66), y: S(0.52), curve: { type: 'cubic', x1: S(0.52),  y1: S(0.32), x2: S(0.62), y2: S(0.44) } },
      { x: S(0.56), y: S(0.62), curve: { type: 'cubic', x1: S(0.68),  y1: S(0.58), x2: S(0.64), y2: S(0.62) } },
      { x: S(0.16), y: S(0.46), curve: { type: 'cubic', x1: S(0.38),  y1: S(0.62), x2: S(0.24), y2: S(0.56) } },
      { x: 0, y: S(0.30),       curve: { type: 'cubic', x1: S(0.06),  y1: S(0.40), x2: 0,       y2: S(0.36) } },
      { close: true },
    ]});
    sl.addShape('line', { x: X(0.40), y: Y(0.02), w:  S(0.10), h: S(0.22), line: ln });
    sl.addShape('line', { x: X(0.62), y: Y(0.06), w: -S(0.10), h: S(0.26), line: ln });
  },

  // eighth note
  note(sl, x, y, h, color) {
    const f = { color }, ln = { type: 'none' }, w = h * 0.636;
    sl.addShape('rect',    { x: x + w * 0.42, y, w: w * 0.10, h: h * 0.80, fill: f, line: ln });
    sl.addShape('rect',    { x: x + w * 0.48, y, w: w * 0.48, h: h * 0.20, fill: f, line: ln });
    sl.addShape('ellipse', { x, y: y + h * 0.74, w: w * 0.52, h: h * 0.26, fill: f, line: ln });
  },

  // rounded-square social badges: hairline box with a white/purple glyph
  twitter(sl, x, y, s, color) {
    badge(sl, x, y, s, color);
    bird(sl, x + s * 0.23, y + s * 0.30, s * 0.54, color);
  },
  facebook(sl, x, y, s, color) {
    badge(sl, x, y, s, color);
    const f = { color }, ln = { type: 'none' }, S = u => s * u;
    sl.addShape('rect', { x: x + S(0.44), y: y + S(0.26), w: S(0.15), h: S(0.50), fill: f, line: ln });
    sl.addShape('rect', { x: x + S(0.44), y: y + S(0.26), w: S(0.18), h: S(0.10), fill: f, line: ln });
    sl.addShape('rect', { x: x + S(0.56), y: y + S(0.26), w: S(0.08), h: S(0.14), fill: f, line: ln });
    sl.addShape('rect', { x: x + S(0.34), y: y + S(0.42), w: S(0.28), h: S(0.10), fill: f, line: ln });
  },
  pinterest(sl, x, y, s, color) {
    badge(sl, x, y, s, color);
    const S = u => s * u, R = S(0.27);
    // filled disc with the lowercase "p" knocked out of it
    sl.addShape('custGeom', {
      x: x + S(0.23), y: y + S(0.23), w: S(0.54), h: S(0.54),
      fill: { color }, line: { type: 'none' },
      points: [
        ...ellipsePath(R, R, R, R),
        ...ellipsePath(S(0.27), S(0.22), S(0.115), S(0.135)),
        { x: S(0.180), y: S(0.280), moveTo: true },
        { x: S(0.245), y: S(0.300) },
        { x: S(0.175), y: S(0.500) },
        { x: S(0.135), y: S(0.485) },
        { close: true },
        ...ellipsePath(S(0.27), S(0.22), S(0.05), S(0.065)),
      ],
    });
  },
};

function badge(sl, x, y, s, color) {
  sl.addShape('roundRect', {
    x, y, w: s, h: s, rectRadius: 0.12,
    fill: { type: 'none' }, line: { color, width: 1.5 },
  });
}

// Closed elliptical subpath, for custGeom shapes that need holes.
function ellipsePath(cx, cy, rx, ry) {
  const k = 0.5523;  // circle-to-bezier constant
  return [
    { x: cx - rx, y: cy, moveTo: true },
    { x: cx, y: cy - ry, curve: { type: 'cubic', x1: cx - rx, y1: cy - ry * k, x2: cx - rx * k, y2: cy - ry } },
    { x: cx + rx, y: cy, curve: { type: 'cubic', x1: cx + rx * k, y1: cy - ry, x2: cx + rx, y2: cy - ry * k } },
    { x: cx, y: cy + ry, curve: { type: 'cubic', x1: cx + rx, y1: cy + ry * k, x2: cx + rx * k, y2: cy + ry } },
    { x: cx - rx, y: cy, curve: { type: 'cubic', x1: cx - rx * k, y1: cy + ry, x2: cx - rx, y2: cy + ry * k } },
    { close: true },
  ];
}

// Silhouette of a bird in flight, used inside the twitter badge.
function bird(sl, x, y, w, color) {
  const S = u => w * u;
  sl.addShape('custGeom', { x, y, w, h: w * 0.82, fill: { color }, line: { type: 'none' }, points: [
    { x: 0, y: S(0.62), moveTo: true },
    { x: S(0.34), y: S(0.72), curve: { type: 'cubic', x1: S(0.14), y1: S(0.74), x2: S(0.24), y2: S(0.74) } },
    { x: S(0.06), y: S(0.42), curve: { type: 'cubic', x1: S(0.16), y1: S(0.66), x2: S(0.06), y2: S(0.56) } },
    { x: S(0.20), y: S(0.46), curve: { type: 'cubic', x1: S(0.12), y1: S(0.46), x2: S(0.16), y2: S(0.47) } },
    { x: S(0.09), y: S(0.20), curve: { type: 'cubic', x1: S(0.10), y1: S(0.40), x2: S(0.06), y2: S(0.30) } },
    { x: S(0.22), y: S(0.26), curve: { type: 'cubic', x1: S(0.14), y1: S(0.24), x2: S(0.18), y2: S(0.25) } },
    { x: S(0.52), y: S(0.06), curve: { type: 'cubic', x1: S(0.16), y1: S(0.06), x2: S(0.34), y2: 0 } },
    { x: S(0.86), y: S(0.14), curve: { type: 'cubic', x1: S(0.68), y1: S(0.10), x2: S(0.80), y2: S(0.14) } },
    { x: S(1.00), y: S(0.08), curve: { type: 'cubic', x1: S(0.92), y1: S(0.13), x2: S(0.97), y2: S(0.10) } },
    { x: S(0.88), y: S(0.26), curve: { type: 'cubic', x1: S(0.97), y1: S(0.18), x2: S(0.92), y2: S(0.23) } },
    { x: S(0.98), y: S(0.22), curve: { type: 'cubic', x1: S(0.93), y1: S(0.25), x2: S(0.96), y2: S(0.23) } },
    { x: S(0.88), y: S(0.36), curve: { type: 'cubic', x1: S(0.95), y1: S(0.29), x2: S(0.92), y2: S(0.33) } },
    { x: 0, y: S(0.62),       curve: { type: 'cubic', x1: S(0.88), y1: S(0.80), x2: S(0.30), y2: S(0.86) } },
    { close: true },
  ]});
}

// ---------------------------------------------------------------- device mockups
// Stand-ins for the phone / laptop photographs used in the reference deck.

function phoneMockup(sl, o) {
  const r = 0.10;
  sl.addShape('roundRect', {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: r,
    fill: { color: 'D5D8DB' }, line: { type: 'none' },
  });
  sl.addShape('roundRect', {
    x: o.x + o.w * 0.022, y: o.y + o.h * 0.010, w: o.w * 0.956, h: o.h * 0.980,
    rectRadius: r * 0.95, fill: { color: '17181A' }, line: { type: 'none' },
  });
  sl.addShape('roundRect', {
    x: o.x + o.w * 0.052, y: o.y + o.h * 0.024, w: o.w * 0.896, h: o.h * 0.952,
    rectRadius: r * 0.8, fill: { color: WHITE }, line: { type: 'none' },
  });
  // front-camera notch
  sl.addShape('roundRect', {
    x: o.x + o.w * 0.30, y: o.y + o.h * 0.024, w: o.w * 0.40, h: o.h * 0.030,
    rectRadius: 0.5, fill: { color: '17181A' }, line: { type: 'none' },
  });
  sl.addText('[image]', {
    x: o.x, y: o.y + o.h * 0.45, w: o.w, h: 0.28,
    fontFace: SANS, fontSize: 10, color: MIDGREY, align: 'center',
  });
}

function laptopMockup(sl, o) {
  const X = u => o.x + o.w * u, Y = v => o.y + o.h * v, W = u => o.w * u, H = v => o.h * v;
  // Dark lid drawn as a frame: the inner rectangle is a knockout so whatever
  // sits behind the laptop (a purple band, on slide 31) shows through.
  sl.addShape('custGeom', {
    x: X(0.097), y: Y(0), w: W(0.806), h: H(0.90),
    fill: { color: '1F2124' }, line: { type: 'none' },
    points: [
      { x: 0, y: 0, moveTo: true }, { x: W(0.806), y: 0 },
      { x: W(0.806), y: H(0.90) }, { x: 0, y: H(0.90) }, { close: true },
      { x: W(0.036), y: H(0.055), moveTo: true }, { x: W(0.770), y: H(0.055) },
      { x: W(0.770), y: H(0.845) }, { x: W(0.036), y: H(0.845) }, { close: true },
    ],
  });
  // tapered silver base
  sl.addShape('custGeom', {
    x: X(0.008), y: Y(0.90), w: W(0.983), h: H(0.075),
    fill: { color: 'CDD1D5' }, line: { type: 'none' },
    points: [
      { x: 0, y: 0, moveTo: true }, { x: W(0.983), y: 0 },
      { x: W(0.967), y: H(0.075) }, { x: W(0.016), y: H(0.075) }, { close: true },
    ],
  });
  sl.addShape('roundRect', {
    x: X(0.44), y: Y(0.905), w: W(0.12), h: H(0.022),
    rectRadius: 0.5, fill: { color: 'AAB0B6' }, line: { type: 'none' },
  });
  sl.addText('[image]', {
    x: X(0.13), y: Y(0.40), w: W(0.74), h: 0.28,
    fontFace: SANS, fontSize: 11, color: MIDGREY, align: 'center',
  });
}

// ---------------------------------------------------------------- charts

// 3-D pie, quarterly sales split (slide 32).
function pieChart(sl, o) {
  sl.addChart('pie', [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr'], values: [5, 4, 2] }], {
    x: o.x, y: o.y, w: o.w, h: o.h,
    chartColors: [INK, PURPLE, MIDGREY],
    dataBorder: { pt: 2, color: WHITE },
    showLegend: true, legendPos: 'b', legendFontFace: SANS, legendFontSize: 12, legendColor: '595959',
    showValue: false, showPercent: false,
  });
}

// 3-D clustered bar, three series over four categories (slide 33).
function barChart(sl, o) {
  const labels = ['Category 1', 'Category 2', 'Category 3', 'Category 4'];
  sl.addChart('bar3D', [
    { name: 'Series 1', labels, values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels, values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels, values: [2, 2, 3, 5] },
  ], {
    x: o.x, y: o.y, w: o.w, h: o.h,
    barDir: 'bar', bar3DShape: 'box', barGapWidthPct: 150,
    chartColors: ['595959', PURPLE, INK],
    v3DRotX: 15, v3DRotY: 20, v3DRAngAx: true, v3DPerspective: 0, serAxisHidden: true,
    catAxisLineShow: false, valAxisLineShow: false, serAxisLineShow: false,
    valGridLine: { color: 'E6E6E6', size: 0.75 }, catGridLine: { style: 'none' },
    catAxisLabelFontFace: SANS, catAxisLabelFontSize: 12, catAxisLabelColor: '595959',
    valAxisLabelFontFace: SANS, valAxisLabelFontSize: 12, valAxisLabelColor: '595959',
    showLegend: true, legendPos: 'b', legendFontFace: SANS, legendFontSize: 12, legendColor: '595959',
  });
}

// ================================================================ slides
function slide01(pres) {
  const sl = pres.addSlide();
  box(sl, { x: 0, y: 0, w: 13.3333, h: 7.5, fill: { color: WHITE, transparency: 30 } });
  heading(sl, [['YANTI', PURPLE]], { x: 3.8074, y: 2.5115, w: 5.7185, h: 1.4473, fontSize: 80, align: 'center', charSpacing: 6 });
  box(sl, { x: 0, y: 4.0739, w: 13.3333, h: 0.5547, fill: { color: PURPLE } });
  label(sl, 'Ballet Presentation Template', { x: 4.5152, y: 4.146, w: 4.3028, h: 0.4104, fontSize: 14, bold: false, lineSpacingMultiple: 1.5, align: 'center', charSpacing: 3 });
}

function slide02(pres) {
  const sl = pres.addSlide();
  card(sl, { x: 2.24, y: 0.6301, w: 11.0933, h: 4.3699, offset: 3, angle: 225 });
  box(sl, { x: 0, y: 0, w: 2.24, h: 0.6301, fill: { color: PURPLE } });
  label(sl, SIMPLE, { x: 0.3501, y: 0.1667, w: 1.5399, h: 0.3029, bold: false, align: 'center' });
  label(sl, SIMPLE, { x: 0.3501, y: 0.7963, w: 1.5399, h: 0.3029, color: BODY, bold: false, align: 'center' });
  label(sl, SIMPLE, { x: 0.3501, y: 1.4306, w: 1.5399, h: 0.3029, color: BODY, bold: false, align: 'center' });
  heading(sl, [[H_A, INK], [H_B, PURPLE]], { x: 3.402, y: 1.3956, w: 5.878, h: 1.1781, fontSize: 32 });
  paragraph(sl, BODY1, { x: 3.402, y: 2.9541, w: 8.838, h: 1.2804 });
}

function slide03(pres) {
  const sl = pres.addSlide();
  card(sl, { x: 3.84, y: 1.5651, w: 9.4933, h: 4.3699 });
  heading(sl, [[H_A, INK], [H_B, PURPLE]], { x: 4.922, y: 2.3305, w: 5.878, h: 1.1781, fontSize: 32 });
  paragraph(sl, BODY2, { x: 4.922, y: 3.889, w: 7.438, h: 1.2804 });
}

function slide04(pres) {
  const sl = pres.addSlide();
  box(sl, { x: 0, y: 5.02, w: 4.94, h: 1.42, fill: { color: PURPLE } });
  icon.shoes(sl, 2.1511, 5.4111, 0.625, WHITE);
  icon.arms(sl, 3.64, 5.4111, 0.6379, WHITE);
  icon.dancer(sl, 0.6621, 5.4111, 0.625, WHITE);
  heading(sl, [[H_A, INK], [H_B, PURPLE]], { x: 1.2075, y: 1.1505, w: 2.725, h: 2.7937, fontSize: 32, align: 'right' });
}

function slide05(pres) {
  const sl = pres.addSlide();
  box(sl, { x: 1.9522, y: 3.75, w: 9.4289, h: 1.31, fill: { color: PURPLE } });
  stat(sl, 3.2289, 4.0612, '250K', { align: 'center', w: 1.0629, cw: 1.9081 });
  stat(sl, 6.1352, 4.0612, '570K', { align: 'center', w: 1.0629, cw: 1.9081 });
  stat(sl, 9.0415, 4.0612, '700K', { align: 'center', w: 1.0629, cw: 1.9081 });
  paragraph(sl, BODY3, { x: 1.0567, y: 5.7393, w: 11.22, h: 0.9775, align: 'center' });
}

function slide06(pres) {
  const sl = pres.addSlide();
  card(sl, { x: 0, y: 0.68, w: 12.52, h: 3.92, shape: 'round1Rect', offset: 3, angle: 315 });
  heading(sl, [[H_A, INK], [H_B, PURPLE]], { x: 1.122, y: 1.4247, w: 5.198, h: 1.0434, fontSize: 28 });
  paragraph(sl, BODY4, { x: 1.122, y: 2.8778, w: 10.318, h: 0.9775 });
}

function slide07(pres) {
  const sl = pres.addSlide();
  cornerBracket(sl, { x: 1.23, y: 1.98, w: 2.98, h: 2.98, fill: { color: PURPLE } });
  heading(sl, [[H_A, INK], [H_B, PURPLE]], { x: 5.242, y: 4.2156, w: 5.878, h: 1.1781, fontSize: 32 });
  paragraph(sl, BODY5, { x: 5.242, y: 5.7625, w: 7.118, h: 0.9775 });
}

function slide08(pres) {
  const sl = pres.addSlide();
  card(sl, { x: 0.6242, y: 3.44, w: 12.0849, h: 3.58, shape: 'round1Rect', rotate: 180 });
  heading(sl, [[H_E, PURPLE], [H_F, INK]], { x: 3.8105, y: 4.3363, w: 5.7123, h: 0.7742, fontSize: 20, align: 'center' });
  paragraph(sl, BODY4, { x: 1.5077, y: 5.4562, w: 10.318, h: 0.9775, align: 'center' });
}

function slide09(pres) {
  const sl = pres.addSlide();
  box(sl, { x: 0.9333, y: 2.58, w: 6.7, h: 4.094, fill: { color: PURPLE } });
  heading(sl, [[H_A, INK], [H_B, PURPLE]], { x: 0.9133, y: 0.7544, w: 6.02, h: 1.1781, fontSize: 32 });
  paragraph(sl, BODY6, { x: 8.5133, y: 1.5259, w: 3.92, h: 0.9775 });
  icon.dancer(sl, 1.5741, 3.4234, 0.625, WHITE);
  stat(sl, 2.4457, 3.3913, '250K', { w: 1.0629, cw: 1.9081 });
  icon.arms(sl, 4.48, 3.4226, 0.625, WHITE);
  stat(sl, 5.3516, 3.3904, '700K', { w: 1.0629, cw: 1.9081 });
  icon.note(sl, 4.5936, 5.1763, 0.625, WHITE);
  stat(sl, 5.3516, 5.1449, '860K', { w: 1.0629, cw: 1.9081 });
  icon.shoes(sl, 1.5742, 5.2068, 0.625, WHITE);
  stat(sl, 2.4458, 5.1759, '570K', { w: 1.0629, cw: 1.9081 });
}

function slide10(pres) {
  const sl = pres.addSlide();
  box(sl, { x: 0, y: 2.8393, w: 13.3333, h: 1.8214, fill: { color: PURPLE } });
  icon.dancer(sl, 0.8443, 3.4386, 0.625, WHITE);
  stat(sl, 1.7158, 3.4064, '250K', {  });
  icon.arms(sl, 7.0228, 3.4381, 0.625, WHITE);
  stat(sl, 7.8944, 3.406, '700K', {  });
  icon.note(sl, 10.1121, 3.4373, 0.625, WHITE);
  stat(sl, 10.8701, 3.406, '860K', {  });
  icon.shoes(sl, 3.9336, 3.4373, 0.625, WHITE);
  stat(sl, 4.8051, 3.4064, '570K', {  });
  heading(sl, [[H_A, INK], [H_B, PURPLE]], { x: 1.4693, y: 0.7633, w: 4.0331, h: 1.3127, fontSize: 24, align: 'right' });
  paragraph(sl, BODY7, { x: 7.5454, y: 5.4401, w: 4.889, h: 1.2804 });
}

function slide11(pres) {
  const sl = pres.addSlide();
  card(sl, { x: 1.2317, y: 1.11, w: 10.8699, h: 5.28 });
  icon.dancer(sl, 3.1083, 2.982, 0.625, PURPLE);
  icon.arms(sl, 9.6, 2.982, 0.625, PURPLE);
  icon.shoes(sl, 6.3542, 2.982, 0.625, PURPLE);
  heading(sl, [[H_C, INK], [H_D, PURPLE]], { x: 3.7049, y: 1.6541, w: 5.9236, h: 0.9088, fontSize: 24, align: 'center' });
  tile(sl, 'Great Service', SHORT, { x: 5.4325, y: 3.838, w: 2.4683, align: 'center', color: INK });
  tile(sl, 'Great Service', SHORT, { x: 2.1866, y: 3.838, w: 2.4683, align: 'center', color: INK });
  tile(sl, 'Great Service', SHORT, { x: 8.6784, y: 3.838, w: 2.4683, align: 'center', color: INK });
  label(sl, '250K', { x: 6.0496, y: 5.3773, w: 1.2342, h: 0.4286, fontSize: 10.5, align: 'center', valign: 'middle', charSpacing: 6, shape: 'round2DiagRect', fill: { color: PURPLE } });
}

function slide12(pres) {
  const sl = pres.addSlide();
  card(sl, { x: 2.8036, y: 1.375, w: 10.5298, h: 3.2857 });
  icon.dancer(sl, 3.5241, 2.0407, 0.625, PURPLE);
  icon.arms(sl, 10.0159, 2.0407, 0.625, PURPLE);
  icon.shoes(sl, 6.77, 2.0407, 0.625, PURPLE);
  tile(sl, 'Great Service', SHORT, { x: 6.77, y: 2.9324, w: 2.4683, color: INK });
  tile(sl, 'Great Service', SHORT, { x: 3.5241, y: 2.9324, w: 2.4683, color: INK });
  tile(sl, 'Great Service', SHORT, { x: 10.0159, y: 2.9324, w: 2.4683, color: INK });
  paragraph(sl, BODY7, { x: 7.5952, y: 5.4401, w: 4.889, h: 1.2804 });
  box(sl, { x: 0, y: 0.7449, w: 2.8036, h: 0.6301, fill: { color: PURPLE } });
  label(sl, SIMPLE, { x: 0.6319, y: 0.9116, w: 1.5399, h: 0.3029, bold: false, align: 'center' });
  label(sl, SIMPLE, { x: 0.6319, y: 1.5412, w: 1.5399, h: 0.3029, color: BODY, bold: false, align: 'center' });
  label(sl, SIMPLE, { x: 0.6319, y: 2.1754, w: 1.5399, h: 0.3029, color: BODY, bold: false, align: 'center' });
}

function slide13(pres) {
  const sl = pres.addSlide();
  box(sl, { x: 0.9951, y: 3.75, w: 4.0763, h: 3.75, fill: { color: PURPLE } });
  heading(sl, [[H_FULL, WHITE]], { x: 1.9495, y: 4.5647, w: 2.1675, h: 2.1205, fontSize: 24, align: 'center' });
  icon.dancer(sl, 6.292, 4.6478, 0.625, PURPLE);
  icon.shoes(sl, 9.7041, 4.6478, 0.625, PURPLE);
  tile(sl, 'Great Service', SHORT, { x: 9.7041, y: 5.5395, w: 2.4683, color: INK });
  tile(sl, 'Great Service', SHORT, { x: 6.292, y: 5.5395, w: 2.4683, color: INK });
}

function slide14(pres) {
  const sl = pres.addSlide();
  heading(sl, [[H_A, INK], [H_B, PURPLE]], { x: 9.3792, y: 0.7042, w: 2.7279, h: 1.7166, fontSize: 24 });
  box(sl, { x: 6.9065, y: 3.125, w: 6.4268, h: 1.7321, fill: { color: PURPLE } });
  paragraph(sl, BODY8, { x: 7.6933, y: 3.6538, w: 4.889, h: 0.6746, color: WHITE });
  icon.dancer(sl, 0.8443, 6.1171, 0.625, PURPLE);
  stat(sl, 1.7158, 6.0849, '250K', { color: INK });
  icon.arms(sl, 7.0228, 6.1167, 0.625, PURPLE);
  stat(sl, 7.8944, 6.0845, '700K', { color: INK });
  icon.note(sl, 10.1121, 6.1159, 0.625, PURPLE);
  stat(sl, 10.8701, 6.0845, '860K', { color: INK });
  icon.shoes(sl, 3.9336, 6.1159, 0.625, PURPLE);
  stat(sl, 4.8051, 6.0849, '570K', { color: INK });
}

function slide15(pres) {
  const sl = pres.addSlide();
  card(sl, { x: 2.4295, y: 1.8359, w: 10.4634, h: 5.32 });
  box(sl, { x: 0, y: 0, w: 13.3333, h: 1.8214, fill: { color: PURPLE } });
  icon.dancer(sl, 0.8443, 0.5993, 0.625, WHITE);
  stat(sl, 1.7158, 0.5671, '250K', {  });
  icon.arms(sl, 7.0228, 0.5988, 0.625, WHITE);
  stat(sl, 7.8944, 0.5667, '700K', {  });
  icon.note(sl, 10.1121, 0.598, 0.625, WHITE);
  stat(sl, 10.8701, 0.5667, '860K', {  });
  icon.shoes(sl, 3.9336, 0.598, 0.625, WHITE);
  stat(sl, 4.8051, 0.5671, '570K', {  });
  heading(sl, [[H_A, INK], [H_B, PURPLE]], { x: 7.4837, y: 2.7842, w: 3.719, h: 1.7166, fontSize: 32 });
  paragraph(sl, BODY9, { x: 7.4837, y: 4.9272, w: 4.5978, h: 1.2804 });
}

function slide16(pres) {
  const sl = pres.addSlide();
  box(sl, { x: 0, y: 2.8291, w: 13.3333, h: 1.875, fill: { color: PURPLE } });
  tile(sl, 'Awesome Team', SHORT, { x: 9.8087, y: 0.9015, w: 2.4683, color: INK });
  tile(sl, 'Awesome Team', SHORT, { x: 9.8087, y: 5.6041, w: 2.4683, color: INK });
  tile(sl, 'Awesome Team', SHORT, { x: 0.989, y: 0.9015, w: 2.4683, align: 'right', color: INK });
  tile(sl, 'Awesome Team', SHORT, { x: 0.989, y: 5.6041, w: 2.4683, align: 'right', color: INK });
  icon.twitter(sl, 2.5206, 3.5309, 0.4844, WHITE);
  stat(sl, 1.3591, 3.4293, '250K', { align: 'right' });
  icon.facebook(sl, 10.3346, 3.5244, 0.4844, WHITE);
  stat(sl, 11.1231, 3.4228, '570K', {  });
}

function slide17(pres) {
  const sl = pres.addSlide();
  paragraph(sl, SHORT, { x: 0.7533, y: 6.1607, w: 2.4683, h: 0.6746, align: 'center' });
  paragraph(sl, SHORT, { x: 3.8725, y: 6.1607, w: 2.4683, h: 0.6746, align: 'center' });
  paragraph(sl, SHORT, { x: 6.9916, y: 6.1607, w: 2.4683, h: 0.6746, align: 'center' });
  paragraph(sl, SHORT, { x: 10.1118, y: 6.1607, w: 2.4683, h: 0.6746, align: 'center' });
  heading(sl, [[H_C, INK], [H_D, PURPLE]], { x: 3.7049, y: 0.6006, w: 5.9236, h: 0.9088, fontSize: 24, align: 'center' });
  box(sl, { x: 0, y: 2.2737, w: 13.3333, h: 0.5512, fill: { color: PURPLE } });
  label(sl, 'Awesome Team', { x: 1.012, y: 2.381, w: 1.9491, h: 0.3366, fontSize: 14, align: 'center' });
  label(sl, 'Awesome Team', { x: 4.1321, y: 2.381, w: 1.9491, h: 0.3366, fontSize: 14, align: 'center' });
  label(sl, 'Awesome Team', { x: 7.2513, y: 2.381, w: 1.9491, h: 0.3366, fontSize: 14, align: 'center' });
  label(sl, 'Awesome Team', { x: 10.3714, y: 2.381, w: 1.9491, h: 0.3366, fontSize: 14, align: 'center' });
}

function slide18(pres) {
  const sl = pres.addSlide();
  tile(sl, 'Awesome Team', SHORT, { x: 0.7316, y: 5.7408, w: 2.4683, color: INK });
  tile(sl, 'Awesome Team', SHORT, { x: 3.8655, y: 5.7408, w: 2.4683, color: INK });
  tile(sl, 'Awesome Team', SHORT, { x: 6.9995, y: 5.7408, w: 2.4683, color: INK });
  tile(sl, 'Awesome Team', SHORT, { x: 10.1334, y: 5.7408, w: 2.4683, color: INK });
  heading(sl, [[H_C, INK], [H_D, PURPLE]], { x: 3.7238, y: 0.6006, w: 5.9236, h: 0.9088, fontSize: 24, align: 'center' });
  box(sl, { x: 0, y: 2.0074, w: 13.3333, h: 1.2632, fill: { color: PURPLE } });
  icon.twitter(sl, 1.4198, 2.3956, 0.4844, WHITE);
  stat(sl, 2.2083, 2.2939, '250K', {  });
  icon.facebook(sl, 5.5028, 2.398, 0.4844, WHITE);
  stat(sl, 6.2913, 2.2964, '570K', {  });
  icon.pinterest(sl, 9.5885, 2.3956, 0.4844, WHITE);
  stat(sl, 10.377, 2.2939, '700K', {  });
}

function slide19(pres) {
  const sl = pres.addSlide();
  box(sl, { x: 4.676, y: 0.7411, w: 7.8214, h: 6.0179, fill: { color: PURPLE } });
  heading(sl, [[H_FULL, WHITE]], { x: 5.6784, y: 1.6541, w: 4.8751, h: 0.9088, fontSize: 24 });
  paragraph(sl, BODY10, { x: 5.6784, y: 3.2582, w: 6.0716, h: 1.2804, color: WHITE });
  icon.twitter(sl, 5.7677, 5.2363, 0.4844, WHITE);
  stat(sl, 6.5562, 5.1347, '250K', {  });
  icon.facebook(sl, 9.3898, 5.2387, 0.4844, WHITE);
  stat(sl, 10.1782, 5.1371, '570K', {  });
}

function slide20(pres) {
  const sl = pres.addSlide();
  box(sl, { x: 4.676, y: 0, w: 8.6574, h: 5.9821, fill: { color: PURPLE } });
  icon.twitter(sl, 3.129, 0.6685, 0.4844, PURPLE);
  stat(sl, 1.9675, 0.5669, '250K', { align: 'right', color: INK });
  icon.facebook(sl, 9.1623, 6.5066, 0.4844, PURPLE);
  stat(sl, 9.9507, 6.405, '570K', { color: INK });
  heading(sl, [[H_FULL, WHITE]], { x: 7.6893, y: 1.3974, w: 4.4714, h: 0.9088, fontSize: 24 });
  paragraph(sl, BODY11, { x: 7.6893, y: 3.0014, w: 4.8821, h: 1.5834, color: WHITE });
}

function slide21(pres) {
  const sl = pres.addSlide();
  box(sl, { x: 1.204, y: 6.2623, w: 4.961, h: 0.3414, shape: 'roundRect', fill: { color: GREY }, rectRadius: 0.5 });
  label(sl, '80%', { x: 1.204, y: 6.2623, w: 3.8405, h: 0.3414, align: 'right', valign: 'middle', shape: 'roundRect', fill: { color: PURPLE }, rectRadius: 0.5, margin: [0, 7.2, 0, 7.2] });
  label(sl, 'Teamwork', { x: 1.2805, y: 6.2895, w: 1.2299, h: 0.2524, valign: 'middle', margin: [0, 7.2, 0, 7.2] });
  box(sl, { x: 7.2482, y: 6.2623, w: 4.961, h: 0.3414, shape: 'roundRect', fill: { color: GREY }, rectRadius: 0.5 });
  label(sl, '94%', { x: 7.2482, y: 6.2623, w: 4.529, h: 0.3414, align: 'right', valign: 'middle', shape: 'roundRect', fill: { color: PURPLE }, rectRadius: 0.5, margin: [0, 7.2, 0, 7.2] });
  label(sl, 'Working Times', { x: 7.3246, y: 6.2998, w: 1.6476, h: 0.2524, valign: 'middle', margin: [0, 7.2, 0, 7.2] });
  heading(sl, [[H_A, INK], [H_B, PURPLE]], { x: 9.3792, y: 0.7042, w: 2.7279, h: 1.7166, fontSize: 24 });
  box(sl, { x: 6.9065, y: 3.125, w: 6.4268, h: 1.7321, fill: { color: PURPLE } });
  paragraph(sl, BODY8, { x: 7.6933, y: 3.6538, w: 4.889, h: 0.6746, color: WHITE });
}

function slide22(pres) {
  const sl = pres.addSlide();
  heading(sl, [[H_A, INK], [H_B, PURPLE]], { x: 0.896, y: 0.7633, w: 4.0331, h: 1.3127, fontSize: 24 });
  paragraph(sl, BODY7, { x: 0.896, y: 2.6764, w: 4.889, h: 1.2804 });
  icon.twitter(sl, 7.5987, 4.023, 0.4844, PURPLE);
  stat(sl, 8.2383, 3.9213, '250K', { color: INK });
  icon.facebook(sl, 10.2304, 4.0254, 0.4844, PURPLE);
  stat(sl, 10.8701, 3.9238, '570K', { color: INK });
  box(sl, { x: 7.5987, y: 5.4014, w: 4.6772, h: 0.3414, shape: 'roundRect', fill: { color: GREY }, rectRadius: 0.5 });
  label(sl, '80%', { x: 7.5987, y: 5.4014, w: 3.6208, h: 0.3414, align: 'right', valign: 'middle', shape: 'roundRect', fill: { color: PURPLE }, rectRadius: 0.5, margin: [0, 7.2, 0, 7.2] });
  label(sl, 'Teamwork', { x: 7.6707, y: 5.4286, w: 1.1595, h: 0.2524, valign: 'middle', margin: [0, 7.2, 0, 7.2] });
  box(sl, { x: 7.5987, y: 6.2106, w: 4.6772, h: 0.3414, shape: 'roundRect', fill: { color: GREY }, rectRadius: 0.5 });
  label(sl, '94%', { x: 7.5987, y: 6.2106, w: 4.2699, h: 0.3414, align: 'right', valign: 'middle', shape: 'roundRect', fill: { color: PURPLE }, rectRadius: 0.5, margin: [0, 7.2, 0, 7.2] });
  label(sl, 'Working Times', { x: 7.6707, y: 6.2481, w: 1.5533, h: 0.2524, valign: 'middle', margin: [0, 7.2, 0, 7.2] });
}

function slide23(pres) {
  const sl = pres.addSlide();
  heading(sl, [[H_A, SLATE], [H_B, PURPLE]], { x: 5.4285, y: 0.8156, w: 6.701, h: 1.3127, fontSize: 36, align: 'center' });
  tile(sl, 'Great Portfolio', BODY12, { x: 7.686, y: 3.0672, w: 2.186, align: 'center', color: INK, bh: 0.9775 });
  tile(sl, 'Great Portfolio', BODY12, { x: 10.3257, y: 3.0672, w: 2.186, align: 'center', color: INK, bh: 0.9775 });
  tile(sl, 'Great Portfolio', BODY12, { x: 5.0463, y: 3.0672, w: 2.186, align: 'center', color: INK, bh: 0.9775 });
}

function slide24(pres) {
  const sl = pres.addSlide();
  card(sl, { x: 8.658, y: 4.2759, w: 3.9828, h: 1.7241 });
  box(sl, { x: 4.6753, y: 5.0345, w: 3.9828, h: 1.7241, fill: { color: PURPLE } });
  card(sl, { x: 0.6925, y: 4.2759, w: 3.9827, h: 1.7241 });
  tile(sl, 'Great Portfolio', SHORT, { x: 1.4497, y: 4.6069, w: 2.4683, align: 'center', color: INK });
  tile(sl, 'Great Portfolio', SHORT, { x: 5.4325, y: 5.38, w: 2.4683, align: 'center', bodyColor: WHITE });
  tile(sl, 'Great Portfolio', SHORT, { x: 9.4153, y: 4.6069, w: 2.4683, align: 'center', color: INK });
  heading(sl, [[H_C, INK], [H_D, PURPLE]], { x: 3.6649, y: 0.6006, w: 5.9236, h: 0.9088, fontSize: 24, align: 'center' });
}

function slide25(pres) {
  const sl = pres.addSlide();
  box(sl, { x: 5.8103, y: 3.7615, w: 3.7615, h: 2.3132, fill: { color: PURPLE } });
  tile(sl, 'Great Portfolio', BODY12, { x: 6.5981, y: 4.2353, w: 2.186, align: 'center', bodyColor: WHITE, bh: 0.9775 });
  box(sl, { x: 9.5718, y: 1.4253, w: 3.7615, h: 2.3132, fill: { color: PURPLE } });
  tile(sl, 'Great Portfolio', BODY12, { x: 10.3596, y: 1.8991, w: 2.186, align: 'center', bodyColor: WHITE, bh: 0.9775 });
  heading(sl, [[H_A, INK], [H_B, PURPLE]], { x: 0.862, y: 0.7689, w: 4.0917, h: 1.3127, fontSize: 24 });
  paragraph(sl, BODY13, { x: 0.862, y: 2.7021, w: 4.0917, h: 2.4922, align: 'justify' });
}

function slide26(pres) {
  const sl = pres.addSlide();
  box(sl, { x: 9.6034, y: 3.5376, w: 3.7299, h: 1.8825, fill: { color: PURPLE } });
  tile(sl, 'Great Portfolio', SHORT, { x: 10.2342, y: 3.9476, w: 2.4683, bodyColor: WHITE });
  box(sl, { x: 9.6034, y: 5.6175, w: 3.7299, h: 1.8825, fill: { color: PURPLE } });
  tile(sl, 'Great Portfolio', SHORT, { x: 10.2342, y: 6.0274, w: 2.4683, bodyColor: WHITE });
  heading(sl, [[H_E, PURPLE], [H_F, INK]], { x: 1.054, y: 0.7026, w: 5.7123, h: 0.7742, fontSize: 20 });
  paragraph(sl, BODY4, { x: 1.054, y: 1.8226, w: 11.2564, h: 0.9775 });
}

function slide27(pres) {
  const sl = pres.addSlide();
  heading(sl, [[H_C, INK], [H_D, PURPLE]], { x: 3.7269, y: 0.6006, w: 5.9236, h: 0.9088, fontSize: 24, align: 'center' });
  paragraph(sl, BODY4, { x: 1.5297, y: 1.9582, w: 10.318, h: 0.9775, align: 'center' });
  label(sl, 'Great Portfolio', { x: 8.1351, y: 4.3621, w: 3.7299, h: 0.6379, fontSize: 14, align: 'center', valign: 'middle', shape: 'rect', fill: { color: PURPLE } });
  label(sl, 'Great Portfolio', { x: 1.4677, y: 4.3621, w: 3.7299, h: 0.6379, fontSize: 14, align: 'center', valign: 'middle', shape: 'rect', fill: { color: PURPLE } });
  label(sl, '250K', { x: 6.0716, y: 3.4996, w: 1.2342, h: 0.4286, fontSize: 10.5, align: 'center', valign: 'middle', charSpacing: 6, shape: 'round2DiagRect', fill: { color: PURPLE } });
}

function slide28(pres) {
  const sl = pres.addSlide();
  box(sl, { x: 0, y: 0.7181, w: 4.94, h: 1.42, fill: { color: PURPLE } });
  icon.shoes(sl, 2.1511, 1.1092, 0.625, WHITE);
  icon.arms(sl, 3.64, 1.1092, 0.6379, WHITE);
  icon.dancer(sl, 0.6621, 1.1092, 0.625, WHITE);
  card(sl, { x: 0.6621, y: 2.1381, w: 6.2262, h: 4.8053 });
  heading(sl, [[H_A, INK], [H_B, PURPLE]], { x: 1.4792, y: 2.829, w: 3.719, h: 1.7166, fontSize: 32 });
  paragraph(sl, BODY9, { x: 1.4792, y: 4.972, w: 4.5978, h: 1.2804 });
}

function slide29(pres) {
  const sl = pres.addSlide();
  box(sl, { x: 0, y: 6.236, w: 13.3333, h: 0.5512, fill: { color: PURPLE } });
  heading(sl, [[H_C, INK], [H_D, PURPLE]], { x: 3.7269, y: 0.6006, w: 5.9236, h: 0.9088, fontSize: 24, align: 'center' });
  paragraph(sl, BODY4, { x: 1.5297, y: 1.9582, w: 10.318, h: 0.9775, align: 'center' });
  label(sl, 'Great Portfolio', { x: 1.012, y: 6.3433, w: 1.9491, h: 0.3366, fontSize: 14, align: 'center' });
  label(sl, 'Great Portfolio', { x: 4.1321, y: 6.3433, w: 1.9491, h: 0.3366, fontSize: 14, align: 'center' });
  label(sl, 'Great Portfolio', { x: 7.2513, y: 6.3433, w: 1.9491, h: 0.3366, fontSize: 14, align: 'center' });
  label(sl, 'Great Portfolio', { x: 10.3714, y: 6.3433, w: 1.9491, h: 0.3366, fontSize: 14, align: 'center' });
}

function slide30(pres) {
  const sl = pres.addSlide();
  phoneMockup(sl, { x: 1.0505, y: 3.6042, w: 2.6628, h: 5.3717 });
  heading(sl, [[H_A, SLATE], [H_B, PURPLE]], { x: 4.762, y: 3.65, w: 5.878, h: 1.1781, fontSize: 32 });
  paragraph(sl, BODY14, { x: 4.762, y: 5.3485, w: 7.458, h: 1.2804 });
}

function slide31(pres) {
  const sl = pres.addSlide();
  box(sl, { x: 0, y: 4.8782, w: 13.3333, h: 1.2632, fill: { color: PURPLE } });
  heading(sl, [[H_C, INK], [H_D, PURPLE]], { x: 3.7269, y: 0.6006, w: 5.9236, h: 0.9088, fontSize: 24, align: 'center' });
  paragraph(sl, BODY15, { x: 8.3539, y: 2.8661, w: 4.035, h: 1.2804 });
  stat(sl, 10.76, 5.1648, '700K', {  });
  stat(sl, 8.3539, 5.1648, '250K', {  });
  laptopMockup(sl, { x: 0.4417, y: 2.1529, w: 7.6529, h: 4.6434 });
}

function slide32(pres) {
  const sl = pres.addSlide();
  pieChart(sl, { x: 6.8491, y: 0.8759, w: 5.9493, h: 3.9483 });
  icon.dancer(sl, 0.946, 5.7638, 0.7776, PURPLE);
  tile(sl, 'Device Chart', SHORT, { x: 1.7986, y: 5.6213, w: 2.4683, color: INK });
  icon.arms(sl, 9.1908, 5.7638, 0.7776, PURPLE);
  tile(sl, 'Device Chart', SHORT, { x: 10.0434, y: 5.6213, w: 2.4683, color: INK });
  icon.shoes(sl, 5.0684, 5.7638, 0.7776, PURPLE);
  tile(sl, 'Device Chart', SHORT, { x: 5.9144, y: 5.6213, w: 2.4683, color: INK });
  card(sl, { x: 0, y: 0.0188, w: 6.2262, h: 4.8053 });
  heading(sl, [[H_A, INK], [H_B, PURPLE]], { x: 0.8171, y: 0.7098, w: 3.719, h: 1.7166, fontSize: 32 });
  paragraph(sl, BODY9, { x: 0.8171, y: 2.8528, w: 4.5978, h: 1.2804 });
}

function slide33(pres) {
  const sl = pres.addSlide();
  icon.dancer(sl, 0.946, 5.7638, 0.7776, PURPLE);
  tile(sl, 'Device Chart', SHORT, { x: 1.7986, y: 5.6213, w: 2.4683, color: INK });
  icon.arms(sl, 9.1908, 5.7638, 0.7776, PURPLE);
  tile(sl, 'Device Chart', SHORT, { x: 10.0434, y: 5.6213, w: 2.4683, color: INK });
  icon.shoes(sl, 5.0684, 5.7638, 0.7776, PURPLE);
  tile(sl, 'Device Chart', SHORT, { x: 5.9144, y: 5.6213, w: 2.4683, color: INK });
  card(sl, { x: 0, y: 0.0188, w: 6.2262, h: 4.8053 });
  heading(sl, [[H_A, INK], [H_B, PURPLE]], { x: 0.8171, y: 0.7098, w: 3.719, h: 1.7166, fontSize: 32 });
  paragraph(sl, BODY9, { x: 0.8171, y: 2.8528, w: 4.5978, h: 1.2804 });
  barChart(sl, { x: 7.1485, y: 0.7606, w: 5.3632, h: 3.8967 });
}

function slide34(pres) {
  const sl = pres.addSlide();
  card(sl, { x: 0.0355, y: 5.011, w: 13.3065, h: 1.9336 });
  mapPin(sl, { x: 1.1974, y: 2.5273, w: 1.7294, h: 1.9357, fill: { color: PURPLE } });
  mapPin(sl, { x: 4.2664, y: 2.5273, w: 1.7294, h: 1.9357, fill: { color: PURPLE } });
  mapPin(sl, { x: 7.3376, y: 2.5273, w: 1.7271, h: 1.9357, fill: { color: PURPLE } });
  mapPin(sl, { x: 10.4065, y: 2.5273, w: 1.7294, h: 1.9357, fill: { color: PURPLE } });
  arrowHead(sl, { x: 10.1776, y: 3.2866, w: 0.1315, h: 0.2652, fill: { color: GREY } });
  dottedTrack(sl, { x: 9.1848, y: 3.3909, w: 0.9089, h: 0.0589, fill: { color: GREY } });
  arrowHead(sl, { x: 7.1064, y: 3.2866, w: 0.1337, h: 0.2652, fill: { color: GREY } });
  dottedTrack(sl, { x: 6.1136, y: 3.3909, w: 0.9112, h: 0.0589, fill: { color: GREY } });
  arrowHead(sl, { x: 4.0148, y: 3.2866, w: 0.1315, h: 0.2652, fill: { color: GREY } });
  dottedTrack(sl, { x: 3.022, y: 3.3909, w: 0.9112, h: 0.0589, fill: { color: GREY } });
  icon.dancer(sl, 1.7182, 3.0753, 0.6878, PURPLE);
  icon.arms(sl, 7.8584, 3.0753, 0.6878, PURPLE);
  icon.note(sl, 11.0524, 3.0753, 0.6878, PURPLE);
  icon.shoes(sl, 4.7872, 3.0753, 0.6878, PURPLE);
  heading(sl, [[H_C, INK], [H_D, PURPLE]], { x: 3.7269, y: 0.6006, w: 5.9236, h: 0.9088, fontSize: 24, align: 'center' });
  tile(sl, 'Great Infographic', SHORT, { x: 10.037, y: 5.4464, w: 2.4683, align: 'center', color: INK });
  tile(sl, 'Great Infographic', SHORT, { x: 6.967, y: 5.4464, w: 2.4683, align: 'center', color: INK });
  tile(sl, 'Great Infographic', SHORT, { x: 3.8969, y: 5.4464, w: 2.4683, align: 'center', color: INK });
  tile(sl, 'Great Infographic', SHORT, { x: 0.828, y: 5.4464, w: 2.4683, align: 'center', color: INK });
}

function slide35(pres) {
  const sl = pres.addSlide();
  card(sl, { x: 0.0355, y: 5.011, w: 13.3065, h: 1.9336 });
  heading(sl, [[H_C, INK], [H_D, PURPLE]], { x: 3.7269, y: 0.6006, w: 5.9236, h: 0.9088, fontSize: 24, align: 'center' });
  tile(sl, 'Great Infographic', SHORT, { x: 10.037, y: 5.4464, w: 2.4683, align: 'center', color: INK });
  tile(sl, 'Great Infographic', SHORT, { x: 6.967, y: 5.4464, w: 2.4683, align: 'center', color: INK });
  tile(sl, 'Great Infographic', SHORT, { x: 3.8969, y: 5.4464, w: 2.4683, align: 'center', color: INK });
  tile(sl, 'Great Infographic', SHORT, { x: 0.828, y: 5.4464, w: 2.4683, align: 'center', color: INK });
  venn(sl, '75% Data One', { x: 0.855, y: 2.3269, w: 1.8665, h: 1.8665, fill: { color: PURPLE } });
  box(sl, { x: 2.8731, y: 2.7189, w: 1.0826, h: 1.0826, shape: 'mathPlus', fill: { color: GREY } });
  venn(sl, '80% Data Two', { x: 4.1073, y: 2.3269, w: 1.8665, h: 1.8665, fill: { color: PURPLE } });
  box(sl, { x: 6.1254, y: 2.7189, w: 1.0826, h: 1.0826, shape: 'mathPlus', fill: { color: GREY } });
  venn(sl, '90% Data Three', { x: 7.3595, y: 2.3269, w: 1.8665, h: 1.8665, fill: { color: PURPLE } });
  box(sl, { x: 9.3776, y: 2.7189, w: 1.0826, h: 1.0826, shape: 'mathEqual', fill: { color: GREY } });
  venn(sl, 'Data Series', { x: 10.6118, y: 2.3269, w: 1.8665, h: 1.8665, fill: { color: PURPLE } });
}

function slide36(pres) {
  const sl = pres.addSlide();
  box(sl, { x: 0, y: 0, w: 13.3333, h: 7.5, fill: { color: WHITE, transparency: 30 } });
  heading(sl, [['THANK YOU', PURPLE]], { x: 2.7704, y: 2.5115, w: 7.7926, h: 1.4473, fontSize: 80, align: 'center', charSpacing: 6 });
  box(sl, { x: 0, y: 4.0739, w: 13.3333, h: 0.5547, fill: { color: PURPLE } });
  label(sl, 'Ballet Presentation Template', { x: 4.5152, y: 4.146, w: 4.3028, h: 0.4104, fontSize: 14, bold: false, lineSpacingMultiple: 1.5, align: 'center', charSpacing: 3 });
}

// ================================================================ build & save
function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'WIDE', width: 13.3333, height: 7.5 });
  pres.layout = 'WIDE';
  pres.title = 'Yanti - Ballet Presentation Template';

  const builders = [
    slide01, slide02, slide03, slide04, slide05, slide06,
    slide07, slide08, slide09, slide10, slide11, slide12,
    slide13, slide14, slide15, slide16, slide17, slide18,
    slide19, slide20, slide21, slide22, slide23, slide24,
    slide25, slide26, slide27, slide28, slide29, slide30,
    slide31, slide32, slide33, slide34, slide35, slide36,
  ];
  builders.forEach(fn => fn(pres));

  return pres.writeFile({ fileName: path.join(__dirname, '0aa96e45-f006-493b-a27d-34461748db8d_grok_final.pptx') });
}

build().then(f => console.log('wrote', f)).catch(e => { console.error(e); process.exit(1); });
