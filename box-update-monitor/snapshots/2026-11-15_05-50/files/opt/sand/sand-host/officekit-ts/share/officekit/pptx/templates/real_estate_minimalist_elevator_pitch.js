/**
 * "REAL ESTATE" — Modern Minimalist Real Estate Presentation Template
 * Rebuilt with pptxgenjs. 23 slides, 13.333" x 7.5" (16:9).
 *
 * Raster artwork in the source deck (icon PNGs, phone/laptop mock-ups and the
 * photography placeholders) is re-drawn here with native pptxgenjs shapes.
 */
const pptxgen = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ *
 * Design tokens
 * ------------------------------------------------------------------ */
const HEAD = 'Inter 24pt';   // theme major font
const BODY = 'Open Sans';    // theme minor font
const MONO = 'Poppins';      // the KPI block on the phone mock-up slide
const CHART_FONT = 'Be Vietnam Pro';

const C = {
  white: 'FFFFFF',
  black: '000000',
  cyan: '00C6E8',      // accent3
  cyanLt: '58E7FF',    // accent3 @ 60% lum  (the pale mint used on cards)
  teal: '27E3DE',      // accent2
  blue: '4B6CFD',      // accent4
  ink: '404040',       // text1 @ 75% lum — body copy
  ink65: '595959',     // text1 @ 65% lum
  hairline: '226964',  // outline colour of the thin frames
  veil1: '595959',     // cover veil    (black @ 65%)
  veil2: '4C4C4C',     // title veil    (black @ 70%)
  veil3: '404040',     // closing veil  (black @ 75%)
  device: '1A1A1A',    // mock-up device body
  screen: 'FFFFFF',    // empty picture placeholder inside a mock-up
  grid: 'D9D9D9'
};

const NONE = { type: 'none' };
// The deck's signature card shadow. Returned fresh each call: pptxgenjs
// rewrites the object it is handed, so a shared literal would compound.
function softShadow() {
  return { type: 'outer', blur: 25, offset: 25, angle: 45, color: C.black, opacity: 0.05 };
}

const LOREM_FULL =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
  'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo ' +
  'magna eros quis urna. Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus.';
const LOREM_MED =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
  'Fusce posuere, magna sed pulvinar.';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ';
const LOREM_TINY = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit.';
const LOREM_MIN = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas.';

/* ------------------------------------------------------------------ *
 * Generic drawing helpers
 * ------------------------------------------------------------------ */
// Solid card with the very soft drop shadow used throughout the deck.
function card(s, x, y, w, h, fill) {
  s.addShape('rect', { x, y, w, h, fill: { color: fill }, line: NONE, shadow: softShadow() });
}

// Thin open frame — the recurring "outline rectangle" motif.
function frame(s, x, y, w, h, width) {
  s.addShape('rect', { x, y, w, h, fill: NONE, line: { color: C.hairline, width: width || 1 } });
}

// Fully rounded pill (roundRect with adj = 50000).
function pill(s, x, y, w, h, opts) {
  s.addShape('roundRect', {
    x, y, w, h, rectRadius: Math.min(w, h) / 2,
    fill: opts.fill ? { color: opts.fill } : NONE,
    line: opts.line ? { color: opts.line, width: 1 } : NONE
  });
}

// Label pill: capsule plus its centred caption (two shapes in the original).
function tagPill(s, x, y, w, text, opts) {
  pill(s, x, y, w, 0.303, { fill: opts.fill, line: opts.line });
  s.addText(text, {
    x: x + 0.044, y, w: w - 0.088, h: 0.303, align: 'center', valign: 'middle',
    fontFace: BODY, fontSize: 12, bold: true, italic: true, color: opts.color
  });
}

// Section kicker: 12pt bold-italic grey line above/below a heading.
function kicker(s, x, y, w, text, align) {
  s.addText(text, {
    x, y, w, h: 0.372, align: align || 'left', valign: 'top', lineSpacingMultiple: 1.5,
    fontFace: BODY, fontSize: 12, bold: true, italic: true, color: C.ink
  });
}

// Body copy: 12pt italic grey.
function para(s, x, y, w, h, text, opts) {
  opts = opts || {};
  s.addText(text, {
    x, y, w, h, align: opts.align || 'left', valign: 'top',
    fontFace: BODY, fontSize: 12, italic: true, color: opts.color || C.ink
  });
}

// Display heading in the major font. `runs` = [text, color, bold?, breakAfter?].
function heading(s, x, y, w, h, runs, size, align) {
  s.addText(runs.map(function (r) {
    return {
      text: r[0],
      options: { fontFace: HEAD, fontSize: size || 48, color: r[1] || C.black, bold: !!r[2], breakLine: !!r[3] }
    };
  }), { x, y, w, h, align: align || 'left', valign: 'top' });
}

// Square number chip ("01." … "04.").
function numBadge(s, x, y, fill, num, numColor, align) {
  s.addShape('rect', { x, y, w: 0.712, h: 0.729, fill: { color: fill }, line: NONE, shadow: softShadow() });
  s.addText(num, {
    x: align === 'right' ? x - 0.185 : x + 0.073, y: y + 0.035, w: 0.824, h: 0.558,
    align: align === 'right' ? 'right' : 'left', valign: 'top', lineSpacingMultiple: 1.5,
    fontFace: BODY, fontSize: 20, bold: true, italic: true, color: numColor
  });
}

/* ------------------------------------------------------------------ *
 * Pictograms (redrawn from the deck's small PNG/SVG artwork)
 * ------------------------------------------------------------------ */
// Thin ">" mark, built from two rotated bars.
function chevron(s, x, y, size, color) {
  const bar = size * 0.075;
  [36.7, -36.7].forEach(function (deg, i) {
    const cy = y + size * (i === 0 ? 0.325 : 0.675);
    s.addShape('rect', {
      x: x + size * 0.485 - size * 0.293, y: cy - bar / 2,
      w: size * 0.586, h: bar, rotate: deg, fill: { color }, line: NONE
    });
  });
}

/**
 * 0..1 normalised mini-drawings, scaled into a `size` square.
 * `bg` is the colour the icon sits on, used for knocked-out details.
 */
function iconGlyph(s, kind, x, y, size, color, bg) {
  const knock = bg || C.white;
  const put = function (c, fx, fy, fw, fh, shape, rot) {
    s.addShape(shape || 'rect', {
      x: x + fx * size, y: y + fy * size, w: fw * size, h: fh * size,
      rotate: rot, fill: { color: c }, line: NONE
    });
  };
  const fill = function (fx, fy, fw, fh, shape, rot) { put(color, fx, fy, fw, fh, shape, rot); };
  const cut = function (fx, fy, fw, fh, shape, rot) { put(knock, fx, fy, fw, fh, shape, rot); };

  if (kind === 'people') {              // speech panel above an audience
    fill(0.10, 0.04, 0.80, 0.32, 'roundRect');
    fill(0.37, 0.33, 0.18, 0.09, 'triangle', 180);
    [[0.20, 0.52], [0.20, 0.66], [0.20, 0.44]].forEach(function (l, i) {
      cut(l[0], 0.11 + i * 0.07, l[1], 0.045);
    });
    [[0.05, 0.47], [0.71, 0.47]].forEach(function (p) {      // outer figures
      fill(p[0] + 0.04, p[1], 0.18, 0.18, 'ellipse');
      fill(p[0], p[1] + 0.20, 0.26, 0.34, 'ellipse');
    });
    fill(0.40, 0.53, 0.20, 0.20, 'ellipse');                  // front figure
    fill(0.32, 0.74, 0.36, 0.34, 'ellipse');
  } else if (kind === 'monitor') {      // screen, avatar, doc card and cursor
    fill(0.04, 0.18, 0.64, 0.10);                             // screen top rail
    fill(0.04, 0.18, 0.10, 0.46);                             // left edge
    fill(0.04, 0.56, 0.64, 0.09);                             // bottom edge
    fill(0.30, 0.65, 0.12, 0.09);                             // stand
    fill(0.17, 0.74, 0.38, 0.09);                             // base
    fill(0.00, 0.24, 0.28, 0.22);                             // doc card
    cut(0.03, 0.27, 0.22, 0.16);
    fill(0.06, 0.30, 0.16, 0.035);
    fill(0.06, 0.36, 0.10, 0.035);
    cut(0.28, 0.24, 0.30, 0.30, 'ellipse');                   // halo behind avatar
    fill(0.31, 0.27, 0.24, 0.24, 'ellipse');
    cut(0.39, 0.31, 0.08, 0.08, 'ellipse');
    cut(0.36, 0.42, 0.14, 0.10, 'ellipse');
    fill(0.52, 0.42, 0.15, 0.18, 'triangle', 160);            // cursor
    [0.00, 0.08, 0.16, 0.24].forEach(function (d) {           // dotted selection
      fill(0.70 + d, 0.02, 0.05, 0.05);
      fill(0.94, 0.02 + d, 0.05, 0.05);
    });
    fill(0.70, 0.10, 0.05, 0.05);
    fill(0.70, 0.18, 0.05, 0.05);
    fill(0.86, 0.26, 0.05, 0.05);
  } else if (kind === 'bars') {         // ascending bars with a trend arrow
    fill(0.02, 0.84, 0.96, 0.11);                             // baseline
    fill(0.86, 0.06, 0.11, 0.84);                             // right rule
    fill(0.10, 0.55, 0.18, 0.30);
    fill(0.35, 0.36, 0.18, 0.49);
    fill(0.60, 0.08, 0.18, 0.77);
    fill(0.12, 0.20, 0.48, 0.07, 'rect', -38);                // arrow shaft
    fill(0.02, 0.30, 0.26, 0.26, 'triangle', 225);            // arrow head
  } else if (kind === 'network') {      // hub with five satellites
    fill(0.465, 0.14, 0.07, 0.36);                            // spokes
    fill(0.125, 0.415, 0.38, 0.07, 'rect', 14);
    fill(0.495, 0.415, 0.38, 0.07, 'rect', -14);
    fill(0.135, 0.63, 0.44, 0.07, 'rect', -49);
    fill(0.425, 0.63, 0.44, 0.07, 'rect', 49);
    fill(0.34, 0.34, 0.32, 0.32, 'ellipse');                  // hub
    cut(0.44, 0.39, 0.10, 0.10, 'ellipse');
    cut(0.41, 0.51, 0.16, 0.10, 'ellipse');
    [[0.38, 0.02], [0.02, 0.30], [0.76, 0.30], [0.10, 0.72], [0.68, 0.72]].forEach(function (p) {
      fill(p[0], p[1], 0.23, 0.23, 'ellipse');
    });
  }
}

/**
 * Device mock-up stand-in. The originals are photographs of a phone / laptop;
 * here the hardware is a plain dark rounded rectangle and the display area is
 * the empty white picture placeholder that the template ships with.
 */
function deviceMockup(s, body, screen) {
  s.addShape('roundRect', {
    x: body.x, y: body.y, w: body.w, h: body.h, rectRadius: body.r,
    fill: { color: C.device }, line: NONE
  });
  s.addShape(screen.r ? 'roundRect' : 'rect', {
    x: screen.x, y: screen.y, w: screen.w, h: screen.h, rectRadius: screen.r,
    fill: { color: C.screen }, line: NONE
  });
}

/* ------------------------------------------------------------------ *
 * Master furniture (top nav, footer, page number) — hidden on 1/2/23
 * ------------------------------------------------------------------ */
const NAV = [['Business', false], ['Estate', true], ['Collaboration', false]];

function chrome(s, pageNo) {
  NAV.forEach(function (item, i) {
    s.addText(item[0], {
      x: 0.839 + i, y: 0.186, w: 1.4, h: 0.286, valign: 'top',
      fontFace: BODY, fontSize: 10.5, bold: item[1], italic: true, color: C.black
    });
  });
  s.addText('Elevator Pitch Presentation Template', {
    x: 0.839, y: 7.123, w: 3.548, h: 0.236, valign: 'top',
    fontFace: BODY, fontSize: 8, italic: true, color: C.black
  });
  s.addText([
    { text: 'Page', options: { fontSize: 10.5, italic: true } },
    { text: '_', options: { fontSize: 10.5, bold: true, italic: true } },
    { text: ' ' + pageNo, options: { fontSize: 8, bold: true } }
  ], { x: 8.985, y: 7.098, w: 3.548, h: 0.286, align: 'right', valign: 'top', fontFace: BODY, color: C.black });
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

// 1 — Cover
function slide01(s) {
  s.addShape('rect', { x: 0, y: 0.633, w: 13.333, h: 6.266, fill: { color: C.veil1 }, line: NONE });
  s.addText([
    { text: 'REAL ', options: { fontFace: HEAD, fontSize: 88, bold: true, color: C.white } },
    { text: 'ESTATE', options: { fontFace: HEAD, fontSize: 80, color: C.white } }
  ], { x: 0.729, y: 2.026, w: 6.644, h: 2.928, valign: 'top' });
  s.addText('Redefining Spaces with Simplicity and Style', {
    x: 0.766, y: 4.884, w: 5.758, h: 0.37, valign: 'top',
    fontFace: BODY, fontSize: 16, bold: true, color: C.cyanLt
  });
  [4.051, 4.202, 4.353].forEach(function (x) { chevron(s, x, 2.885, 0.474, C.cyanLt); });
  s.addShape('ellipse', { x: 10.646, y: 3.445, w: 0.609, h: 0.609, fill: NONE, line: { color: C.white, width: 3 } });
  chevron(s, 10.752, 3.471, 0.557, C.white);
}

// 2 — Title
function slide02(s) {
  s.addShape('rect', { x: 0, y: 0, w: 13.333, h: 7.533, fill: { color: C.veil2 }, line: NONE });
  s.addText('REAL ESTATE', {
    x: 1.143, y: 2.908, w: 11.046, h: 1.717, align: 'center', valign: 'top',
    fontFace: HEAD, fontSize: 96, bold: true, color: C.white
  });
  s.addText('Modern Minimalist Real Estate Presentation Template', {
    x: 6.53, y: 4.415, w: 4.56, h: 0.961, align: 'right', valign: 'top', lineSpacingMultiple: 1.5,
    fontFace: BODY, fontSize: 18, italic: true, color: C.cyanLt
  });
  [['left', 0.839, 7.057], ['right', 8.947, 0.19]].forEach(function (f) {
    s.addText('Elevator Pitch Presentation Template', {
      x: f[1], y: f[2], w: 3.548, h: 0.286, align: f[0], valign: 'top',
      fontFace: BODY, fontSize: 11, italic: true, color: C.black
    });
  });
  pill(s, 2.241, 2.542, 2.001, 0.366, { line: C.white });
  s.addText('Presentation', {
    x: 2.432, y: 2.53, w: 2.562, h: 0.37, valign: 'top',
    fontFace: BODY, fontSize: 16, bold: true, italic: true, color: C.cyanLt
  });
}

// 3 — Shifting Preferences
function slide03(s) {
  frame(s, -0.114, 1.2, 7.026, 2.908, 0.5);
  frame(s, 12.495, 3.959, 5.859, 2.908, 0.5);
  tagPill(s, 7.744, 1.971, 1.464, 'Introduction', { fill: C.cyan, color: C.white });
  heading(s, 7.58, 2.391, 4.23, 1.717, [['Shifting Preferences', C.black]]);
  kicker(s, 7.58, 4.107, 5.388, 'minimalist properties in urban and suburban areas.');
  para(s, 7.58, 4.713, 4.304, 1.111, LOREM_FULL);
}

// 4 — Real Estate Trends
function slide04(s) {
  frame(s, 12.495, 3.959, 5.859, 2.593, 0.5);
  kicker(s, 0.839, 1.38, 5.828, 'The Rise of Simplicity and Functionality');
  para(s, 0.839, 1.838, 6.478, 0.909, LOREM_FULL);
  pill(s, 8.747, 4.151, 1.464, 0.286, { line: C.cyan });
  s.addText('Real Estate', {
    x: 8.821, y: 4.131, w: 1.317, h: 0.303, align: 'center', valign: 'middle',
    fontFace: BODY, fontSize: 12, bold: true, italic: true, color: C.ink65
  });
  heading(s, 8.578, 4.496, 3.917, 1.717, [['Real Estate ', C.black], ['Trends', C.cyan]]);
}

// 5 — Why Minimalism?
function slide05(s) {
  frame(s, 3.621, 4.444, 10.149, 2.289, 0.5);
  card(s, 0.942, 4.444, 2.346, 2.289, C.cyanLt);

  const FEATURES = [
    { x: 1.204, ix: 1.245, glyph: 'people', ink: C.white, glyphColor: C.white, bg: C.cyanLt, ty: 5.34, by: 5.766 },
    { x: 4.026, ix: 4.124, glyph: 'monitor', ink: C.black, glyphColor: C.cyan, ty: 5.34, by: 5.766 },
    { x: 6.849, ix: 6.925, glyph: 'bars', ink: C.black, glyphColor: C.cyan, ty: 5.322, by: 5.748 },
    { x: 9.672, ix: 9.807, glyph: 'network', ink: C.black, glyphColor: C.cyan, ty: 5.322, by: 5.748 }
  ];
  FEATURES.forEach(function (f) {
    iconGlyph(s, f.glyph, f.ix, 4.78, 0.542, f.glyphColor, f.bg);
    s.addText('Option Here', {
      x: f.x, y: f.ty, w: 1.74, h: 0.375, valign: 'top', lineSpacingMultiple: 1.5,
      fontFace: BODY, fontSize: 12, bold: true, italic: true, color: f.ink
    });
    para(s, f.x, f.by, 2.228, 0.707, LOREM_TINY, { color: f.ink === C.white ? C.white : C.ink });
  });

  pill(s, 7.795, 1.338, 1.592, 0.286, { line: C.cyan });
  s.addText('Real Estate', {
    x: 7.842, y: 1.317, w: 1.498, h: 0.303, align: 'center', valign: 'middle',
    fontFace: BODY, fontSize: 12, bold: true, italic: true, color: C.ink65
  });
  heading(s, 7.719, 1.688, 4.23, 1.717, [['Why ', C.cyan], ['Minimalism', C.black], ['?', C.cyan]]);
  kicker(s, 7.719, 3.405, 4.425, 'Focus on affordability, sustainability, and buyer preference for clean, functional homes.');
}

// 6 — A Modern Lifestyle Choice
function slide06(s) {
  frame(s, -0.256, 1.841, 6.188, 1.126, 1);
  const ROWS = [
    { ty: 1.967, tx: 2.734, by: 2.299, bx: 0.825, bax: 5.562, bay: 2.04, fill: C.cyan, num: '01.' },
    { ty: 3.375, tx: 2.391, by: 3.707, bx: 0.482, bax: 5.22, bay: 3.448, fill: C.cyanLt, num: '02.' },
    { ty: 4.783, tx: 2.391, by: 5.115, bx: 0.482, bax: 5.22, bay: 4.856, fill: C.cyanLt, num: '03.' }
  ];
  ROWS.forEach(function (r) {
    s.addText('Option Here', {
      x: r.tx, y: r.ty, w: 2.721, h: 0.375, align: 'right', valign: 'top', lineSpacingMultiple: 1.5,
      fontFace: BODY, fontSize: 12, bold: true, italic: true, color: C.ink
    });
    para(s, r.bx, r.by, 4.63, 0.505, LOREM_SHORT, { align: 'right' });
    numBadge(s, r.bax, r.bay, r.fill, r.num, C.white, 'right');
  });
  kicker(s, 7.058, 2.023, 4.425, 'Minimalist Residential Spaces');
  heading(s, 7.058, 2.516, 5.475, 1.717, [['A', C.cyan, true], [' Modern Lifestyle Choice', C.black]]);
  para(s, 7.058, 4.367, 4.63, 1.111, LOREM_FULL);
}

// 7 — Urban Real Estate Trends (four cards)
function slide07(s) {
  frame(s, 6.767, 2.655, 6.789, 2.537, 1);
  const CARDS = [
    { x: 5.375, y: 1.611, fill: C.cyanLt, fg: C.white, bw: 2.302, badge: C.white, numColor: C.ink, num: '01.' },
    { x: 9.293, y: 1.611, fill: C.white, fg: C.ink, bw: 2.43, badge: C.cyan, numColor: C.white, num: '02.' },
    { x: 5.375, y: 4.212, fill: C.white, fg: C.ink, bw: 2.43, badge: C.cyan, numColor: C.white, num: '03.' },
    { x: 9.293, y: 4.212, fill: C.white, fg: C.ink, bw: 2.43, badge: C.cyan, numColor: C.white, num: '04.' }
  ];
  CARDS.forEach(function (c) {
    card(s, c.x, c.y, 3.061, 1.774, c.fill);
    s.addText('Option Here', {
      x: c.x + 0.481, y: c.y + 0.084, w: 2.721, h: 0.375, valign: 'top', lineSpacingMultiple: 1.5,
      fontFace: BODY, fontSize: 12, bold: true, italic: true, color: c.fg
    });
    para(s, c.x + 0.481, c.y + 0.459, c.bw, 1.111, LOREM_MED, { color: c.fg });
    numBadge(s, c.x - 0.291, c.y - 0.098, c.badge, c.num, c.numColor);
  });

  pill(s, 0.939, 1.54, 1.592, 0.286, { line: C.cyan });
  s.addText('Real Estate', {
    x: 0.986, y: 1.519, w: 1.498, h: 0.303, align: 'center', valign: 'middle',
    fontFace: BODY, fontSize: 12, bold: true, italic: true, color: C.ink65
  });
  heading(s, 0.863, 1.891, 2.875, 3.736,
    [['Urban', C.black], [' ', C.ink], ['Real ', C.black], ['Estate', C.cyan], [' Trends', C.black]], 54);
  kicker(s, 0.863, 5.609, 4.425, 'Compact Yet Functional Living');
}

// 8 — Premium Minimalist Design
function slide08(s) {
  frame(s, 7.493, 0.38, 5.556, 5.873, 1);
  heading(s, 0.839, 1.09, 5.113, 2.524, [['Premium', C.black], [' Minimalist ', C.cyan], ['Design', C.black]]);
  kicker(s, 0.839, 3.886, 5.388, 'Luxury Real Estate Trends');
  para(s, 0.839, 4.346, 4.715, 0.707,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
    'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero.');
  tagPill(s, 8.195, 6.127, 1.464, 'Real Estate', { fill: C.cyan, color: C.white });
  s.addShape('ellipse', { x: 6.622, y: 2.658, w: 0.39, h: 0.39, fill: { color: C.white }, line: { color: C.cyan, width: 1.5 } });
  chevron(s, 6.688, 2.675, 0.357, C.cyan);
}

// 9 — Commercial Market Trends
function slide09(s) {
  frame(s, -0.178, 1.367, 6.066, 2.633, 1);
  kicker(s, 6.446, 1.247, 4.425, 'Minimalist Residential Spaces');
  heading(s, 6.446, 1.673, 5.475, 1.717, [['Commercial ', C.black], ['Market', C.cyan], [' Trends', C.black]]);
  para(s, 6.446, 3.495, 5.424, 0.505, LOREM_MED);

  const CARDS = [
    { x: 6.188, fill: C.cyanLt, fg: C.white, bw: 1.904, badge: C.white, numColor: C.ink, num: '01.' },
    { x: 9.423, fill: C.white, fg: C.black, bw: 2.009, badge: C.cyan, numColor: C.white, num: '02.' }
  ];
  CARDS.forEach(function (c) {
    card(s, c.x, 4.651, 2.653, 1.538, c.fill);
    s.addText('Option Here', {
      x: c.x + 0.481, y: 4.735, w: 2.721, h: 0.375, valign: 'top', lineSpacingMultiple: 1.5,
      fontFace: BODY, fontSize: 12, bold: true, italic: true, color: c.fg
    });
    para(s, c.x + 0.481, 5.11, c.bw, 0.909, LOREM_MIN, { color: c.fg === C.white ? C.white : C.ink });
    numBadge(s, c.x - 0.291, 4.553, c.badge, c.num, c.numColor);
  });
}

// 10 — Minimalism & Investment
function slide10(s) {
  frame(s, 7.553, 0.694, 6.201, 5.873, 1);
  heading(s, 8.322, 1.128, 5.006, 1.717,
    [['Minimalism', C.black, false, true], ['&', C.cyan, true], [' Investment', C.black]]);

  [{ x: 0.926, fill: C.cyanLt, color: C.white },
   { x: 3.253, fill: C.white, color: C.black },
   { x: 5.58, fill: C.white, color: C.black }].forEach(function (t) {
    card(s, t.x, 4.074, 2.205, 0.435, t.fill);
    s.addText('Option Here', {
      x: t.x + 0.172, y: 4.049, w: 1.839, h: 0.372, align: 'center', valign: 'top', lineSpacingMultiple: 1.5,
      fontFace: BODY, fontSize: 12, bold: true, italic: true, color: t.color
    });
  });

  kicker(s, 0.839, 4.836, 5.828, 'High Returns in Real Estate');
  s.addText([
    { text: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
            'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero,', options: { breakLine: true } },
    { text: 'sit amet commodo magna eros quis urna. Nunc viverra imperdiet enim. Fusce est.' }
  ], { x: 0.839, y: 5.295, w: 6.525, h: 0.707, valign: 'top', fontFace: BODY, fontSize: 12, italic: true, color: C.ink });
}

// 11 — Sustainability in Minimalist Design
function slide11(s) {
  frame(s, 7.609, 2.618, 5.258, 2.816, 1);
  heading(s, 0.839, 1.256, 6.173, 1.717,
    [['Sustainability in ', C.black], ['Minimalist', C.cyan], [' Design', C.black]]);
  tagPill(s, 10.671, 2.467, 1.464, 'Real Estate', { fill: C.cyan, color: C.white });
  card(s, 8.55, 3.75, 3.945, 2.311, C.white);
  kicker(s, 8.847, 3.982, 3.648, 'Eco-Conscious Real Estate');
  para(s, 8.847, 4.458, 3.37, 1.313,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
    'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo ' +
    'magna eros quis urna. Nunc viverra imperdiet enim. Fusce est. ');
}

// 12 — Affordable Minimalist Housing (two blurbs)
function slide12(s) {
  frame(s, -0.354, 3.933, 3.623, 2.08, 1);
  card(s, 3.564, 3.933, 4.621, 2.08, C.white);
  kicker(s, 0.909, 1.223, 4.425, 'Redefining Cost-Efficiency');
  heading(s, 0.909, 1.716, 6.324, 1.717,
    [['Affordable ', C.black], ['Minimalist', C.cyan], [' Housing', C.black]]);
  kicker(s, 0.839, 4.105, 2.471, 'Option Here');
  para(s, 0.909, 4.523, 2.003, 1.111, LOREM_SHORT);
  kicker(s, 3.784, 4.126, 2.721, 'Option Here');
  para(s, 3.762, 4.537, 4.225, 1.111,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
    'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo ' +
    'magna eros quis urna. Nunc viverra imperdiet enim. Fusce est.');
}

// 13 — Understanding Today's Market (line chart)
function slide13(s) {
  card(s, 5.88, 3.34, 7.545, 3.241, C.white);
  s.addChart('line', [{
    name: 'Column1',
    labels: ['2019', '2020', '2021', '2022', '2023'],
    values: [55, 25, 65, 80, 90]
  }], {
    x: 6.478, y: 4.195, w: 5.512, h: 2.082,
    chartColors: [C.blue], lineSize: 1.25, lineSmooth: false,
    lineDataSymbol: 'circle', lineDataSymbolSize: 5, lineDataSymbolLineSize: 0.75,
    showLegend: false, showTitle: false, showValue: true,
    dataLabelPosition: 't', dataLabelFormatCode: '#,###\\k',
    dataLabelFontFace: CHART_FONT, dataLabelFontSize: 12,
    dataLabelFontBold: true, dataLabelFontItalic: true, dataLabelColor: C.ink,
    catAxisLabelFontFace: CHART_FONT, catAxisLabelFontSize: 12, catAxisLabelColor: C.ink,
    catAxisLineColor: C.grid, catAxisMajorTickMark: 'none', catAxisMinorTickMark: 'none',
    valAxisMinVal: 0, valAxisMaxVal: 100, valAxisMajorUnit: 50,
    valAxisLabelFormatCode: '#,##0\\k',
    valAxisLabelFontFace: CHART_FONT, valAxisLabelFontSize: 12, valAxisLabelColor: C.ink,
    valAxisLineShow: false, valAxisMajorTickMark: 'none',
    valGridLine: { color: C.grid, size: 0.75 }, catGridLine: { style: 'none' },
    chartArea: { fill: { color: C.white }, border: { pt: 0, color: C.white } }
  });
  s.addText('Dashboard Overview', {
    x: 6.569, y: 3.735, w: 2.609, h: 0.274, margin: 0, valign: 'top', lineSpacingMultiple: 1.5,
    fontFace: BODY, fontSize: 12, bold: true, italic: true, color: C.ink
  });
  kicker(s, 7.44, 1.009, 4.425, 'Buyer Preferences');
  heading(s, 7.44, 1.502, 5.475, 1.717,
    [['Understanding ', C.black], ['Today\u2019s', C.cyan], [' Market', C.black]]);

  [{ y: 4.163, iy: 4.274, glyph: 'monitor' }, { y: 5.39, iy: 5.499, glyph: 'bars' }].forEach(function (r) {
    iconGlyph(s, r.glyph, 0.932, r.iy, 0.707, C.cyan);
    kicker(s, 1.866, r.y, 3.453, 'Option Here');
    para(s, 1.866, r.y + 0.332, 3.921, 0.707, LOREM_SHORT);
  });
}

// 14 — Affordable Minimalist Housing (banner rows)
function slide14(s) {
  kicker(s, 6.106, 1.223, 4.425, 'Redefining Cost-Efficiency');
  heading(s, 6.106, 1.716, 6.324, 1.717,
    [['Affordable ', C.black], ['Minimalist', C.cyan], [' Housing', C.black]]);
  frame(s, -0.111, 3.747, 11.24, 1.126, 1);

  s.addText('Option Here', {
    x: 7.931, y: 3.872, w: 2.721, h: 0.375, align: 'right', valign: 'top', lineSpacingMultiple: 1.5,
    fontFace: BODY, fontSize: 12, bold: true, italic: true, color: C.ink
  });
  para(s, 6.022, 4.204, 4.63, 0.505, LOREM_SHORT, { align: 'right' });
  numBadge(s, 10.76, 3.945, C.cyan, '01.', C.white, 'right');

  kicker(s, 7.01, 5.18, 2.721, 'Option Here');
  para(s, 7.009, 5.512, 3.643, 0.505, LOREM_MIN);
  numBadge(s, 6.19, 5.253, C.cyanLt, '02.', C.white);

  card(s, 1.1, 4.631, 2.721, 1.576, C.white);
  kicker(s, 1.32, 4.76, 2.721, 'Option Here');
  para(s, 1.298, 5.17, 2.311, 0.707,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue.');
}

// 15 — Frequently Asked Questions
function slide15(s) {
  card(s, 0.839, 3.259, 4.42, 2.992, C.white);
  heading(s, 0.839, 1.311, 5.508, 1.717,
    [['Frequently', C.black, false, true], ['Asked', C.cyan], [' Questions', C.black]]);
  s.addText('FAQ\u2019s', {
    x: 4.341, y: 1.505, w: 1.707, h: 0.512, valign: 'top', lineSpacingMultiple: 1.5,
    fontFace: BODY, fontSize: 18, italic: true, color: C.cyan
  });
  s.addText([
    'Lorem ipsum dolor sit amet', 'Consectetuer adipiscing elit',
    'Maecenas porttitor congue massa.', 'Fusce posuere, magna sed'
  ].map(function (t) {
    return { text: t, options: { breakLine: true, bullet: { characterCode: '25AA', indent: 13.5 } } };
  }), {
    x: 1.41, y: 3.627, w: 3.71, h: 2.038, valign: 'top', lineSpacingMultiple: 2.5,
    fontFace: BODY, fontSize: 12, bold: true, italic: true, color: C.ink
  });

  frame(s, 6.987, 1.597, 6.911, 1.126, 1);
  [{ x: 7.464, y: 1.722, bw: 4.63, bx: 6.644, by: 1.796, fill: C.cyan, num: '01.' },
   { x: 8.102, y: 3.22, bw: 4.168, bx: 7.282, by: 3.293, fill: C.cyanLt, num: '02.' },
   { x: 8.102, y: 4.331, bw: 4.168, bx: 7.282, by: 4.404, fill: C.cyanLt, num: '03.' },
   { x: 8.102, y: 5.441, bw: 4.168, bx: 7.282, by: 5.514, fill: C.cyanLt, num: '04.' }].forEach(function (r) {
    kicker(s, r.x, r.y, 2.721, 'Option Here');
    para(s, r.x, r.y + 0.333, r.bw, 0.505, LOREM_SHORT);
    numBadge(s, r.bx, r.by, r.fill, r.num, C.white);
  });
}

// 16 — Modern Apartment Trends
function slide16(s) {
  tagPill(s, 7.457, 1.271, 1.464, 'Introduction', { fill: C.cyan, color: C.white });
  heading(s, 7.293, 1.691, 4.23, 2.524, [['Modern ', C.black], ['Apartment', C.cyan], [' Trends', C.black]]);
  kicker(s, 8.815, 4.622, 5.388, 'The Shift to Space Optimization');
  para(s, 8.815, 5.228, 3.23, 1.313,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
    'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo ' +
    'magna eros quis urna. ');

  card(s, 0.839, 1.691, 2.326, 2.224, C.white);
  s.addText('+1M', { x: 1.083, y: 1.835, w: 1.742, h: 0.774, valign: 'top', fontFace: HEAD, fontSize: 40, color: C.black });
  s.addText('Option Here', {
    x: 1.084, y: 2.564, w: 1.742, h: 0.303, valign: 'top',
    fontFace: BODY, fontSize: 12, bold: true, color: C.ink
  });
  para(s, 1.085, 2.9, 2.08, 0.707, LOREM_TINY);
}

// 17-20 — S.W.O.T pages (mirrored left/right variants)
function swotSlide(s, word, side) {
  const left = side === 'left';
  frame(s, left ? 8.544 : -0.557, 1.304, 5.354, 4.892, 1);
  const tx = left ? 0.856 : 8.13;
  card(s, left ? 0.693 : 7.443, 2.752, 5.206, 1.256, C.white);
  kicker(s, tx, 2.226, 4.691, 'S.W.O.T Analysis Template');
  heading(s, tx, 2.844, 5.811, 1.01, [[word, C.cyan]], 54);
  para(s, tx, left ? 4.223 : 4.252, 4.477, 1.111, LOREM_FULL);
}

// 21 — Mock-Up (phone) + KPI panel
function slide21(s) {
  heading(s, 0.82, 2.745, 4.631, 1.01, [['Mock-Up', C.cyan]], 54);
  kicker(s, 0.887, 3.63, 4.631, 'Mock-Up Template Template');
  para(s, 0.887, 4.176, 3.289, 0.909,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
    'Fusce posuere, magna sed pulvinar ultricies, purus lectus');

  card(s, 8.369, 2.263, 4.132, 3.499, C.white);
  [{ label: 'Average Revenue Per Unit', lw: 2.636, value: '570.00$', fill: C.teal, ly: 2.748, vy: 3.098 },
   { label: 'Customer Lifetime Value', lw: 2.522, value: '429.00$', fill: C.cyan, ly: 3.676, vy: 4.026 },
   { label: 'Customer Acquisition Cost', lw: 2.735, value: '483.00$', fill: C.blue, ly: 4.603, vy: 4.953 }].forEach(function (k) {
    s.addText(k.label, {
      x: 9.46, y: k.ly, w: k.lw, h: 0.202, margin: 0, valign: 'top',
      fontFace: MONO, fontSize: 12, bold: true, italic: true, color: C.ink
    });
    pill(s, 9.46, k.vy, 1.292, 0.331, { fill: k.fill });
    s.addText(k.value, {
      x: 9.46, y: k.vy, w: 1.292, h: 0.331, valign: 'middle',
      fontFace: MONO, fontSize: 12, bold: true, color: C.white
    });
  });

  // phone mock-up — drawn last so it overlaps the KPI card, as in the original.
  // A second handset peeks out on the left, as in the photographed original.
  s.addShape('roundRect', {
    x: 4.844, y: 2.013, w: 0.547, h: 5.487, rectRadius: 0.3,
    fill: { color: C.device }, line: NONE
  });
  deviceMockup(s,
    { x: 5.437, y: 1.553, w: 3.272, h: 5.947, r: 0.42 },
    { x: 5.544, y: 1.74, w: 2.985, h: 5.76, r: 0.28 });
  s.addShape('roundRect', {                              // front-camera notch
    x: 6.204, y: 1.6, w: 1.665, h: 0.31, rectRadius: 0.14,
    fill: { color: C.device }, line: NONE
  });
}

// 22 — Mock-Up (laptop)
function slide22(s) {
  deviceMockup(s,
    { x: -0.747, y: 1.973, w: 6.691, h: 4.074, r: 0.12 },
    { x: 0.173, y: 2.26, w: 5.544, h: 3.467, r: 0 });
  s.addShape('rect', {                                   // laptop base / hinge
    x: -0.747, y: 6.047, w: 7.377, h: 0.146, fill: { color: '5E5E62' }, line: NONE
  });

  heading(s, 6.667, 2.083, 5.591, 1.111, [['Mock-Up', C.cyan]], 60);
  kicker(s, 6.727, 3.111, 4.691, 'Mock-Up Template Template');
  frame(s, 8.544, 3.932, 5.354, 2.077, 1);
  card(s, 5.367, 4.161, 4.963, 1.543, C.white);
  para(s, 5.611, 4.347, 4.493, 1.111,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
    'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo ' +
    'magna eros quis urna. Nunc viverra imperdiet enim. Fusce est.', { align: 'justify' });
  heading(s, 10.574, 4.448, 2.526, 0.909, [['+', C.ink, true], ['1M', C.black, true]]);
}

// 23 — Thank You
function slide23(s) {
  s.addShape('rect', { x: 0, y: 0.633, w: 13.333, h: 6.266, fill: { color: C.veil3 }, line: NONE });
  s.addText('Thank You!', {
    x: 1.143, y: 2.748, w: 11.046, h: 2.036, align: 'center', valign: 'top',
    fontFace: HEAD, fontSize: 115, bold: true, color: C.white
  });
  s.addText('Real Estate Simplified for a Modern Future', {
    x: 7.548, y: 4.479, w: 3.452, h: 0.961, align: 'right', valign: 'top', lineSpacingMultiple: 1.5,
    fontFace: BODY, fontSize: 18, italic: true, color: C.cyanLt
  });
  s.addText('Elevator Pitch Presentation Template', {
    x: 0.839, y: 7.057, w: 3.548, h: 0.286, valign: 'top',
    fontFace: BODY, fontSize: 11, italic: true, color: C.black
  });
  pill(s, 2.413, 2.559, 1.792, 0.366, { line: C.white });
  s.addText('Conclusion ', {
    x: 2.604, y: 2.547, w: 2.562, h: 0.37, valign: 'top',
    fontFace: BODY, fontSize: 16, bold: true, italic: true, color: C.cyanLt
  });
}

/* ------------------------------------------------------------------ *
 * Deck assembly
 * ------------------------------------------------------------------ */
const SLIDES = [
  { build: slide01, chrome: false },
  { build: slide02, chrome: false },
  { build: slide03 },
  { build: slide04 },
  { build: slide05 },
  { build: slide06 },
  { build: slide07 },
  { build: slide08 },
  { build: slide09 },
  { build: slide10 },
  { build: slide11 },
  { build: slide12 },
  { build: slide13 },
  { build: slide14 },
  { build: slide15 },
  { build: slide16 },
  { build: function (s) { swotSlide(s, 'STRENGTH', 'left'); } },
  { build: function (s) { swotSlide(s, 'Weakness', 'right'); } },
  { build: function (s) { swotSlide(s, 'Opportunity', 'left'); } },
  { build: function (s) { swotSlide(s, 'Threat', 'right'); } },
  { build: slide21 },
  { build: slide22 },
  { build: slide23, chrome: false }
];

function buildDeck() {
  const pptx = new pptxgen();
  pptx.defineLayout({ name: 'DECK_16x9', width: 13.333, height: 7.5 });
  pptx.layout = 'DECK_16x9';
  pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
  pptx.title = 'REAL ESTATE — Modern Minimalist Real Estate Presentation Template';

  SLIDES.forEach(function (def, i) {
    const s = pptx.addSlide();
    s.background = { color: C.white };
    if (def.chrome !== false) chrome(s, i + 1);
    def.build(s);
  });

  return pptx.writeFile({
    fileName: path.join(__dirname, '01708432-192f-46f0-8fb8-84aaf926cc95_grok_final.pptx')
  });
}

buildDeck().then(function (f) { console.log('wrote ' + f); }, function (e) { console.error(e); process.exit(1); });
