/**
 * "Medical Health Care" — 15-slide deck rebuilt with pptxgenjs.
 * Slide size 13.333 x 7.5 in (16:9). Theme fonts: Poppins Medium / Poppins Light.
 * Raster photos in the original are replaced by flat grey "[image]" placeholders.
 */
'use strict';

const path = require('path');
const pptxgen = require('pptxgenjs');

// ---------------------------------------------------------------- palette ---
const C = {
  navy: '26275D',   // accent1
  pink: 'E4BAD2',   // accent2
  yellow: 'EBDE6C', // accent3
  mint: 'A3DEBE',   // accent4
  peach: 'F9C0AB',  // accent5
  blue: '94BBF1',   // accent6
  white: 'FFFFFF',
  black: '000000',
  grey: '595959',   // tx1 lum 65/35 — body copy
  greyDark: '404040', // tx1 lum 75/25
  greyMid: '3F3F3F',
  greyInk: '262626',
  greySoft: 'D0CECE',
  navyPale: 'C9CAE9', // navy lum 20/80
  mintDeep: '5CC58C', // mint lum 75
  pinkDeep: 'C770A1', // pink lum 75
  placeholder: 'D9D9D9', // stand-in for photographic content
};

const FONT_H = 'Poppins Medium'; // +mj-lt
const FONT_B = 'Poppins Light';  // +mn-lt

const NO_LINE = { type: 'none' };
// PowerPoint's default text-body insets, in points: [left, right, bottom, top].
const DEF_INSET = [7.2, 7.2, 3.6, 3.6];

// -------------------------------------------------------------- primitives ---
const fill = (color, transparency) => (transparency ? { color, transparency } : { color });

/** Solid / translucent rectangle. */
function rect(s, x, y, w, h, color, transparency) {
  s.addShape('rect', { x, y, w, h, fill: fill(color, transparency), line: NO_LINE });
}

/** Rounded rectangle; `r` is the corner radius in inches. */
function roundRect(s, x, y, w, h, r, color, transparency) {
  s.addShape('roundRect', {
    x, y, w, h, rectRadius: r, fill: fill(color, transparency), line: NO_LINE,
  });
}

/** Filled circle of diameter `d` centred nowhere in particular — x/y is top-left. */
function circle(s, x, y, d, color, transparency, line) {
  s.addShape('ellipse', {
    x, y, w: d, h: d, fill: fill(color, transparency), line: line || NO_LINE,
  });
}

/**
 * Free-form polygon from normalised [0..1] points scaled into the x/y/w/h box.
 * custGeom path coordinates are relative to the shape's own bounding box.
 */
function poly(s, x, y, w, h, pts, color, opts) {
  s.addShape('custGeom', Object.assign({
    x, y, w, h,
    fill: color ? { color } : { type: 'none' },
    line: NO_LINE,
    points: pts.map(([px, py]) => ({ x: +(px * w).toFixed(4), y: +(py * h).toFixed(4) }))
      .concat([{ close: true }]),
  }, opts || {}));
}

/** Circle outline (no fill) — e.g. the chevron rings on the contents slide. */
function ring(s, x, y, d, color, width) {
  s.addShape('ellipse', {
    x, y, w: d, h: d, fill: { type: 'none' }, line: { color, width: width || 1 },
  });
}

/**
 * Text box. `runs` is a string or an array of {text, ...runProps}.
 * Defaults mirror the deck: Poppins Light, top-anchored, no internal padding.
 */
function text(s, runs, o) {
  const body = typeof runs === 'string' ? [{ text: runs }] : runs;
  s.addText(
    body.map(r => ({
      text: r.text,
      options: {
        fontFace: r.font || o.font || FONT_B,
        fontSize: r.size || o.size || 18,
        bold: r.bold !== undefined ? r.bold : !!o.bold,
        color: r.color || o.color || C.black,
        superscript: !!r.sup,
        charSpacing: r.spc || o.spc,
        breakLine: !!r.br,
      },
    })),
    {
      x: o.x, y: o.y, w: o.w, h: o.h || 0.35,
      align: o.align || 'left',
      valign: o.valign || 'top',
      fontFace: o.font || FONT_B,
      fontSize: o.size || 18,
      color: o.color || C.black,
      bold: !!o.bold,
      lineSpacingMultiple: o.lnSpc,
      margin: o.margin || DEF_INSET,
      wrap: o.wrap !== false,
      bullet: o.bullet || false,
    }
  );
}

/**
 * Grey block standing in for a photo that the original deck embedded.
 * `corners` selects which corners are rounded: 'all' | 'top' | 'right' | 'tr'.
 * round2SameRect only rounds the *top* pair, so 'right' rotates it a quarter turn;
 * a rotated shape spins about its centre, hence the swapped pre-rotation box.
 */
const CORNER_SHAPES = {
  all: ['roundRect', 0], top: ['round2SameRect', 0],
  right: ['round2SameRect', 90], tr: ['round1Rect', 0],
};
function imagePlaceholder(s, x, y, w, h, opts) {
  const o = opts || {};
  const [shape, rotate] = CORNER_SHAPES[o.corners] || ['rect', 0];
  const box = rotate === 90
    ? { x: x + (w - h) / 2, y: y + (h - w) / 2, w: h, h: w }
    : { x, y, w, h };
  s.addShape(shape, Object.assign({}, box, {
    rotate, rectRadius: o.radius,
    fill: fill(o.color || C.placeholder), line: NO_LINE,
  }));
  if (o.label !== false) {
    text(s, '[image]', {
      x, y: y + h / 2 - 0.18, w, h: 0.36, align: 'center', valign: 'middle',
      size: 12, color: 'A6A6A6',
    });
  }
}

// ----------------------------------------------------------- shared chrome ---
/** Diamond "Health With Me" mark: two facing triangles split by a plus sign. */
function logoMark(s, x, y, d, color) {
  const c = color || C.navy;
  poly(s, x, y, d, 0.40 * d, [[0.5, 0], [1, 1], [0, 1]], c);
  poly(s, x, y + 0.60 * d, d, 0.40 * d, [[0, 0], [1, 0], [0.5, 1]], c);
  poly(s, x + 0.33 * d, y + 0.33 * d, 0.34 * d, 0.34 * d,
    [[0.34, 0], [0.66, 0], [0.66, 0.34], [1, 0.34], [1, 0.66], [0.66, 0.66],
      [0.66, 1], [0.34, 1], [0.34, 0.66], [0, 0.66], [0, 0.34], [0.34, 0.34]], c);
}

/** Microscope glyph for the pink header bubble. */
function microscopeIcon(s, x, y, d) {
  rect(s, x + 0.10 * d, y + 0.84 * d, 0.80 * d, 0.12 * d, C.white);
  poly(s, x + 0.22 * d, y + 0.56 * d, 0.46 * d, 0.28 * d,
    [[0.25, 0], [0.75, 0], [1, 1], [0, 1]], C.white);
  poly(s, x + 0.36 * d, y + 0.10 * d, 0.44 * d, 0.50 * d,
    [[0.45, 0], [1, 0.28], [0.55, 1], [0, 0.72]], C.white);
  rect(s, x + 0.62 * d, y + 0.56 * d, 0.30 * d, 0.10 * d, C.white);
}

/** House glyph for the yellow header bubble and the contact "Address" chip. */
function houseIcon(s, x, y, d) {
  poly(s, x, y + 0.08 * d, d, 0.46 * d, [[0.5, 0], [1, 1], [0, 1]], C.white);
  rect(s, x + 0.15 * d, y + 0.50 * d, 0.70 * d, 0.42 * d, C.white);
  rect(s, x + 0.38 * d, y + 0.64 * d, 0.24 * d, 0.28 * d, C.yellow);
}

/** Magnifier glyph inside the mint search button. */
function searchIcon(s, x, y, d) {
  ring(s, x, y, 0.68 * d, C.white, 1.5);
  poly(s, x + 0.56 * d, y + 0.56 * d, 0.44 * d, 0.44 * d,
    [[0, 0.34], [0.34, 0], [1, 0.66], [0.66, 1]], C.white);
}

/** Hospital building glyph: an "H" tower flanked by two wings of windows. */
function hospitalIcon(s, x, y, d, bg) {
  const cut = bg || C.mint;
  rect(s, x + 0.30 * d, y, 0.40 * d, 0.44 * d, C.white);
  poly(s, x + 0.39 * d, y + 0.09 * d, 0.22 * d, 0.24 * d,
    [[0, 0], [0.26, 0], [0.26, 0.38], [0.74, 0.38], [0.74, 0], [1, 0], [1, 1],
      [0.74, 1], [0.74, 0.62], [0.26, 0.62], [0.26, 1], [0, 1]], cut);
  rect(s, x, y + 0.44 * d, d, 0.56 * d, C.white);
  [0.10, 0.38, 0.66].forEach(cx => {
    [0.56, 0.76].forEach(cy => rect(s, x + cx * d, y + cy * d, 0.24 * d, 0.12 * d, cut));
  });
}

/** Umbrella sheltering three heads (slide 5, step one). */
function umbrellaIcon(s, x, y, d) {
  poly(s, x + 0.06 * d, y, 0.88 * d, 0.34 * d,
    [[0.5, 0], [0.78, 0.1], [1, 0.42], [1, 0.62], [0.88, 0.42], [0.75, 0.62],
      [0.62, 0.42], [0.5, 0.62], [0.38, 0.42], [0.25, 0.62], [0.12, 0.42],
      [0, 0.62], [0, 0.42], [0.22, 0.1]], C.white);
  circle(s, x + 0.06 * d, y + 0.42 * d, 0.24 * d, C.white);
  circle(s, x + 0.70 * d, y + 0.42 * d, 0.24 * d, C.white);
  circle(s, x + 0.38 * d, y + 0.38 * d, 0.26 * d, C.white);
  rect(s, x, y + 0.72 * d, 0.30 * d, 0.28 * d, C.white);
  rect(s, x + 0.70 * d, y + 0.72 * d, 0.30 * d, 0.28 * d, C.white);
  rect(s, x + 0.34 * d, y + 0.68 * d, 0.32 * d, 0.32 * d, C.white);
}

/** Cocktail glass (slide 5, "Stay Hydrated"). */
function glassIcon(s, x, y, d) {
  poly(s, x, y, d, 0.56 * d, [[0, 0], [1, 0], [0.5, 1]], C.white);
  rect(s, x + 0.45 * d, y + 0.54 * d, 0.10 * d, 0.36 * d, C.white);
  rect(s, x + 0.22 * d, y + 0.88 * d, 0.56 * d, 0.12 * d, C.white);
}

/** Nurse bust with a cross badge on the coat (slides 5, 6 and 15). */
function nurseIcon(s, x, y, d, bg) {
  const cut = bg || C.yellow;
  circle(s, x + 0.24 * d, y, 0.52 * d, C.white);
  circle(s, x + 0.38 * d, y + 0.20 * d, 0.06 * d, cut);
  circle(s, x + 0.56 * d, y + 0.20 * d, 0.06 * d, cut);
  rect(s, x, y + 0.60 * d, d, 0.40 * d, C.white);
  poly(s, x + 0.08 * d, y + 0.68 * d, 0.24 * d, 0.24 * d,
    [[0.36, 0], [0.64, 0], [0.64, 0.36], [1, 0.36], [1, 0.64], [0.64, 0.64],
      [0.64, 1], [0.36, 1], [0.36, 0.64], [0, 0.64], [0, 0.36], [0.36, 0.36]], cut);
  rect(s, x + 0.46 * d, y + 0.60 * d, 0.08 * d, 0.40 * d, cut);
}

/** Eye glyph (slide 5, "Get Enough Sleep"). */
function eyeIcon(s, x, y, d) {
  poly(s, x, y + 0.18 * d, d, 0.64 * d,
    [[0, 0.5], [0.25, 0.08], [0.5, 0], [0.75, 0.08], [1, 0.5],
      [0.75, 0.92], [0.5, 1], [0.25, 0.92]], C.white);
  circle(s, x + 0.36 * d, y + 0.36 * d, 0.28 * d, C.mint);
}

/** Circled X used on the "End Presentation" chip. */
function closeIcon(s, x, y, d, color) {
  poly(s, x, y, d, d,
    [[0.14, 0], [0.5, 0.36], [0.86, 0], [1, 0.14], [0.64, 0.5], [1, 0.86],
      [0.86, 1], [0.5, 0.64], [0.14, 1], [0, 0.86], [0.36, 0.5], [0, 0.14]], color);
}

/** Doctor bust glyph: head over shoulders. */
function doctorIcon(s, x, y, d) {
  circle(s, x + 0.30 * d, y, 0.40 * d, C.white);
  poly(s, x, y + 0.48 * d, d, 0.52 * d,
    [[0.5, 0], [0.8, 0.14], [1, 0.55], [1, 1], [0, 1], [0, 0.55], [0.2, 0.14]], C.white);
}

/** Telephone handset glyph for the contact list. */
function phoneIcon(s, x, y, d) {
  poly(s, x, y, d, d,
    [[0.06, 0.2], [0.26, 0], [0.5, 0.26], [0.34, 0.44], [0.58, 0.7], [0.76, 0.52],
      [1, 0.76], [0.8, 0.96], [0.6, 1], [0.3, 0.83], [0.12, 0.6], [0.02, 0.36]], C.white);
}

/** Envelope glyph for the contact list. */
function mailIcon(s, x, y, d) {
  rect(s, x, y + 0.18 * d, d, 0.62 * d, C.white);
  poly(s, x + 0.06 * d, y + 0.24 * d, 0.88 * d, 0.38 * d,
    [[0, 0], [1, 0], [0.5, 1]], C.navy);
}

/** Stethoscope glyph: tubing loop plus earpieces and a bell. */
function stethoscopeIcon(s, x, y, d) {
  poly(s, x + 0.08 * d, y + 0.05 * d, 0.52 * d, 0.55 * d,
    [[0, 0], [0.16, 0], [0.16, 0.5], [0.5, 0.9], [0.84, 0.5], [0.84, 0], [1, 0],
      [1, 0.56], [0.5, 1.09], [0, 0.56]], C.white);
  circle(s, x + 0.05 * d, y, 0.14 * d, C.white);
  circle(s, x + 0.47 * d, y, 0.14 * d, C.white);
  rect(s, x + 0.30 * d, y + 0.60 * d, 0.08 * d, 0.14 * d, C.white);
  poly(s, x + 0.34 * d, y + 0.72 * d, 0.52 * d, 0.28 * d,
    [[0, 0], [0.14, 0], [0.14, 0.7], [1, 0.7], [1, 0.86], [0, 0.86]], C.white);
  circle(s, x + 0.74 * d, y + 0.48 * d, 0.26 * d, C.white);
}

/** Family glyph: an adult figure sheltering a smaller one. */
function familyIcon(s, x, y, d) {
  circle(s, x + 0.10 * d, y + 0.06 * d, 0.30 * d, C.white);
  poly(s, x, y + 0.44 * d, 0.56 * d, 0.46 * d,
    [[0.45, 0], [0.78, 0.16], [1, 0.6], [1, 1], [0, 1], [0, 0.6], [0.16, 0.16]], C.white);
  circle(s, x + 0.60 * d, y + 0.32 * d, 0.24 * d, C.white);
  poly(s, x + 0.52 * d, y + 0.62 * d, 0.44 * d, 0.32 * d,
    [[0.5, 0], [0.85, 0.2], [1, 1], [0, 1], [0.15, 0.2]], C.white);
}

/** Group-of-people glyph: three heads over one shared body. */
function groupIcon(s, x, y, d) {
  circle(s, x + 0.36 * d, y, 0.28 * d, C.white);
  circle(s, x + 0.02 * d, y + 0.20 * d, 0.24 * d, C.white);
  circle(s, x + 0.74 * d, y + 0.20 * d, 0.24 * d, C.white);
  poly(s, x + 0.18 * d, y + 0.42 * d, 0.64 * d, 0.36 * d,
    [[0.5, 0], [0.82, 0.2], [1, 0.8], [1, 1], [0, 1], [0, 0.8], [0.18, 0.2]], C.white);
  poly(s, x, y + 0.58 * d, 0.34 * d, 0.28 * d, [[0.7, 0], [1, 0.5], [1, 1], [0, 1], [0, 0.4]], C.white);
  poly(s, x + 0.66 * d, y + 0.58 * d, 0.34 * d, 0.28 * d, [[0.3, 0], [1, 0.4], [1, 1], [0, 1], [0, 0.5]], C.white);
}

/** Globe glyph for the contact list. */
function globeIcon(s, x, y, d) {
  ring(s, x, y, d, C.white, 1.5);
  s.addShape('ellipse', {
    x: x + 0.3 * d, y, w: 0.4 * d, h: d, fill: { type: 'none' },
    line: { color: C.white, width: 1 },
  });
  rect(s, x, y + 0.47 * d, d, 0.02 * d, C.white);
}

/** White top bar with the brand mark, two icon bubbles and the search pill. */
function header(s) {
  rect(s, 0, 0, 13.333, 1.032, C.white);
  logoMark(s, 0.597, 0.354, 0.30);
  text(s, 'Health With Me', {
    x: 0.91, y: 0.364, w: 1.424, h: 0.303, size: 12, wrap: false, valign: 'middle',
  });

  circle(s, 9.858, 0.283, 0.466, C.pink);
  microscopeIcon(s, 9.968, 0.375, 0.28);
  circle(s, 10.418, 0.283, 0.466, C.yellow);
  houseIcon(s, 10.528, 0.375, 0.28);

  s.addShape('roundRect', {
    x: 11.024, y: 0.309, w: 1.689, h: 0.414, rectRadius: 0.207,
    fill: { type: 'none' }, line: { color: C.black, width: 1 },
  });
  text(s, 'Search More', {
    x: 11.101, y: 0.373, w: 1.295, h: 0.286, size: 11, valign: 'middle',
  });
  circle(s, 12.349, 0.357, 0.318, C.mint);
  searchIcon(s, 12.427, 0.432, 0.165);
}

// ------------------------------------------------------------ shape tables ---
/** Head-in-profile puzzle (slide 10): six pieces plus three connector nubs. */
const HEAD_PATHS = {
  tl: [[1, 0], [1, 1], [0, 1], [0.0983, 0.81], [0.277, 0.574], [0.181, 0.529], [0.124, 0.447], [0.117, 0.349], [0.175, 0.25], [0.336, 0.145], [0.514, 0.075], [0.732, 0.0231], [1, 0]],
  tr: [[0, 0], [0.176, 0.0263], [0.362, 0.0976], [0.5, 0.186], [0.626, 0.298], [0.738, 0.427], [0.874, 0.645], [0.942, 0.798], [1, 1], [0, 1], [0, 0]],
  ml: [[0.251, 0], [1, 0], [1, 1], [0.117, 1], [0.151, 0.958], [0.11, 0.92], [0.0873, 0.86], [0.0868, 0.793], [0.11, 0.703], [0.0912, 0.67], [0.0114, 0.608], [0.0032, 0.542], [0.206, 0.207], [0.251, 0]],
  mr: [[0, 0], [0.991, 0], [0.995, 0.243], [0.954, 0.47], [0.86, 0.742], [0.733, 1], [0, 1], [0, 0]],
  bl: [[0.0542, 0], [1, 0], [1, 0.998], [0.79, 0.989], [0.559, 0.938], [0.488, 0.708], [0.388, 0.519], [0.295, 0.456], [0.141, 0.453], [0.0564, 0.411], [0.0197, 0.355], [0, 0.263], [0.0112, 0.194], [0.0628, 0.084], [0.0542, 0]],
  br: [[0, 0], [1, 0], [0.77, 0.255], [0.723, 0.362], [0.71, 0.473], [0.72, 0.58], [0.793, 0.806], [0.591, 0.886], [0.384, 0.946], [0, 1], [0, 0]],
};
/** Interlocking tab shared by the head / body puzzles. */
const PUZZLE_NUB = [[0.809, 0.78], [0.919, 0.871], [1, 1], [1, 0], [0.89, 0.16], [0.762, 0.239], [0.637, 0.235], [0.427, 0.128], [0.325, 0.105], [0.215, 0.128], [0.0949, 0.22], [0.0189, 0.364], [0.00218, 0.543], [0.0504, 0.707], [0.15, 0.827], [0.25, 0.878], [0.341, 0.888], [0.429, 0.866], [0.658, 0.758], [0.741, 0.755], [0.809, 0.78]];

/** Standing figure puzzle (slide 11). */
const BODY_PATHS = {
  head: [[1, 0.616], [0.97, 0.823], [0.89, 1], [0.107, 0.994], [0.0285, 0.818], [0, 0.614], [0.0299, 0.407], [0.0789, 0.285], [0.148, 0.18], [0.233, 0.0957], [0.332, 0.0358], [0.5, 0], [0.668, 0.0359], [0.767, 0.0958], [0.853, 0.18], [0.922, 0.285], [0.971, 0.408], [1, 0.616]],
  armL: [[1, 0], [1, 0.984], [0.429, 1], [0.489, 0.599], [0.152, 0.554], [0.079, 0.514], [0.0262, 0.442], [0.00112, 0.348], [0.00434, 0.272], [0.0365, 0.183], [0.0944, 0.118], [0.171, 0.0859], [0.641, 0.0179], [0.631, 0.00358], [1, 0]],
  armR: [[1, 0.322], [0.974, 0.442], [0.937, 0.498], [0.887, 0.538], [0.51, 0.599], [0.57, 1], [0, 0.984], [0, 0], [0.367, 0.00895], [0.36, 0.0197], [0.85, 0.092], [0.906, 0.119], [0.964, 0.184], [1, 0.322]],
  legL: [[1, 0], [1, 0.604], [0.618, 0.924], [0.507, 0.98], [0.381, 1], [0.26, 0.985], [0.183, 0.956], [0.0896, 0.889], [0.0135, 0.769], [0.000881, 0.664], [0.141, 0.0149], [1, 0]],
  legR: [[0.697, 0.994], [0.567, 0.996], [0.423, 0.95], [0, 0.604], [0, 0], [0.86, 0.0167], [0.994, 0.63], [0.995, 0.737], [0.955, 0.834], [0.881, 0.916], [0.778, 0.973], [0.697, 0.994]],
};

/** Heart puzzle (slide 12): x, y, w, h, colour, outline. */
const HEART_PIECES = [
  [4.881, 3.625, 0.782, 1.285, C.navy, [[1, 0], [0.125, 0], [0.0578, 0.117], [0.0158, 0.241], [0.00835, 0.491], [0.045, 0.632], [0.111, 0.776], [0.314, 0.776], [0.26, 0.812], [0.233, 0.858], [0.236, 0.904], [0.276, 0.953], [0.345, 0.987], [0.435, 1], [0.58, 0.964], [0.628, 0.918], [0.639, 0.867], [0.617, 0.819], [0.556, 0.776], [1, 0.776], [1, 0]]],
  [4.979, 3.072, 0.684, 0.847, C.navy, [[1, 0], [0.772, 0.0482], [0.539, 0.138], [0.401, 0.217], [0.222, 0.359], [0.122, 0.469], [0, 0.653], [0.107, 0.653], [0.0378, 0.708], [0.00254, 0.783], [0.00617, 0.855], [0.0513, 0.929], [0.131, 0.981], [0.234, 1], [0.312, 0.989], [0.399, 0.945], [0.453, 0.875], [0.466, 0.797], [0.438, 0.72], [0.36, 0.653], [1, 0.653], [1, 0]]],
  [5.397, 3.625, 1.293, 1.275, C.navy, [[1, 0], [0.206, 0], [0.206, 0.285], [0.143, 0.255], [0.0692, 0.266], [0.0192, 0.311], [0, 0.378], [0.0192, 0.446], [0.0692, 0.491], [0.143, 0.502], [0.206, 0.472], [0.206, 0.782], [0.52, 0.782], [0.483, 0.844], [0.486, 0.917], [0.537, 0.98], [0.603, 1], [0.669, 0.98], [0.719, 0.917], [0.723, 0.844], [0.686, 0.782], [1, 0.782], [1, 0]]],
  [5.663, 3.06, 1.027, 0.859, C.navy, [[1, 0.63], [0.909, 0.475], [0.787, 0.323], [0.539, 0.118], [0.296, 0.0179], [0.111, 0.000536], [0, 0.0138], [0, 0.658], [0.416, 0.658], [0.359, 0.735], [0.345, 0.836], [0.369, 0.914], [0.431, 0.981], [0.536, 0.995], [0.631, 0.914], [0.652, 0.773], [0.584, 0.658], [1, 0.658], [1, 0.63]]],
  [6.414, 3.625, 1.322, 1.275, C.pink, [[0.909, 0.504], [0.815, 0.457], [0.791, 0.35], [0.824, 0.29], [0.881, 0.256], [0.946, 0.259], [1, 0.296], [1, 0], [0.208, 0], [0.208, 0.305], [0.15, 0.269], [0.0801, 0.273], [0.0266, 0.312], [0.000814, 0.376], [0.0123, 0.446], [0.0677, 0.503], [0.143, 0.514], [0.208, 0.477], [0.208, 0.782], [0.523, 0.782], [0.487, 0.844], [0.49, 0.917], [0.529, 0.972], [0.59, 0.999], [0.657, 0.987], [0.713, 0.93], [0.723, 0.851], [0.685, 0.782], [1, 0.782], [1, 0.461], [0.909, 0.504]]],
  [6.69, 3.01, 1.046, 0.88, C.pink, [[1, 0.029], [0.852, 0.00157], [0.713, 0.00694], [0.584, 0.0398], [0.427, 0.12], [0.295, 0.228], [0.188, 0.349], [0.0464, 0.572], [0, 0.699], [0.402, 0.699], [0.364, 0.818], [0.398, 0.932], [0.499, 0.999], [0.599, 0.972], [0.646, 0.916], [0.669, 0.84], [0.662, 0.762], [0.632, 0.699], [1, 0.699], [1, 0.029]]],
  [7.456, 3.624, 0.992, 0.996, C.pink, [[0.997, 0.367], [0.973, 0.181], [0.921, 0], [0.282, 0], [0.282, 0.378], [0.191, 0.327], [0.0902, 0.34], [0.0164, 0.414], [0.00108, 0.503], [0.0354, 0.584], [0.107, 0.635], [0.201, 0.64], [0.282, 0.59], [0.282, 1], [0.9, 1], [0.983, 0.682], [0.997, 0.367]]],
  [7.736, 3.035, 0.66, 0.854, C.pink, [[0.94, 0.69], [0.96, 0.69], [0.833, 0.484], [0.607, 0.261], [0.315, 0.0938], [0, 0], [0, 0.69], [0.576, 0.69], [0.517, 0.793], [0.54, 0.895], [0.628, 0.971], [0.758, 1], [0.888, 0.971], [0.975, 0.895], [0.997, 0.783], [0.94, 0.69]]],
  [4.968, 4.621, 0.695, 0.913, C.mint, [[0, 0], [0.156, 0.271], [0.386, 0.543], [0.66, 0.778], [1, 1], [1, 0], [0, 0]]],
  [5.663, 4.621, 1.027, 0.997, C.mint, [[1, 0], [0, 0], [0, 0.916], [0.102, 1], [1, 1], [1, 0]]],
  [5.768, 5.35, 0.922, 0.876, C.mint, [[0.538, 0.305], [0.581, 0.212], [0.566, 0.102], [0.503, 0.0284], [0.41, 0], [0.317, 0.0284], [0.254, 0.102], [0.239, 0.212], [0.282, 0.305], [0, 0.305], [0.578, 0.75], [1, 1], [1, 0.305], [0.538, 0.305]]],
  [6.426, 4.621, 1.556, 1.283, C.yellow, [[0.897, 0.264], [0.842, 0.283], [0.842, 0], [0.17, 0], [0.17, 0.294], [0.108, 0.264], [0.0476, 0.283], [0.00599, 0.346], [0.00271, 0.417], [0.0226, 0.466], [0.0681, 0.506], [0.124, 0.51], [0.17, 0.483], [0.17, 0.776], [0.444, 0.776], [0.407, 0.842], [0.409, 0.917], [0.451, 0.981], [0.507, 1], [0.606, 0.904], [0.603, 0.834], [0.568, 0.776], [0.732, 0.776], [0.842, 0.659], [0.842, 0.493], [0.921, 0.509], [0.984, 0.455], [0.99, 0.333], [0.952, 0.283], [0.897, 0.264]]],
  [6.69, 5.617, 0.875, 0.616, C.yellow, [[0, 0.987], [0.0254, 1], [0.0844, 0.961], [0.343, 0.735], [0.654, 0.408], [1, 0], [0, 0], [0, 0.987]]],
  [7.736, 4.621, 0.613, 0.846, C.yellow, [[0, 1], [0.35, 0.758], [0.63, 0.509], [0.843, 0.257], [1, 0], [0, 0], [0, 1]]],
];

/** Brain puzzle (slide 13): x, y, w, h, colour, outline. */
const BRAIN_PIECES = [
  [7.09, 2.118, 2.868, 2.04, C.navy, [[1, 0.774], [0.548, 0.774], [0.579, 0.831], [0.58, 0.906], [0.552, 0.966], [0.499, 1], [0.444, 0.97], [0.412, 0.911], [0.41, 0.84], [0.442, 0.774], [0.000956, 0.774], [0.00956, 0.628], [0.0468, 0.496], [0.11, 0.386], [0.215, 0.3], [0.335, 0.258], [0.384, 0.167], [0.448, 0.0968], [0.495, 0.0712], [0.565, 0.0605], [0.631, 0.0739], [0.707, 0.112], [0.812, 0.0202], [0.902, 0.00134], [0.959, 0.0255], [1, 0.0605], [1, 0.286], [0.957, 0.261], [0.892, 0.274], [0.852, 0.333], [0.84, 0.425], [0.86, 0.499], [0.892, 0.539], [0.944, 0.559], [1, 0.531], [1, 0.774]]],
  [9.549, 2.083, 2.786, 1.604, C.pink, [[0.167, 1], [0.167, 0.602], [0.149, 0.65], [0.119, 0.679], [0.0896, 0.687], [0.0531, 0.677], [0.00886, 0.609], [0.00295, 0.492], [0.0305, 0.421], [0.0856, 0.383], [0.13, 0.403], [0.167, 0.46], [0.167, 0.094], [0.209, 0.0564], [0.284, 0.0154], [0.348, 0], [0.402, 0.00342], [0.494, 0.0513], [0.55, 0.111], [0.604, 0.2], [0.736, 0.277], [0.839, 0.397], [0.9, 0.571], [0.913, 0.829], [0.975, 0.918], [1, 1], [0.659, 1], [0.685, 0.903], [0.676, 0.797], [0.637, 0.721], [0.575, 0.685], [0.52, 0.718], [0.476, 0.814], [0.472, 0.882], [0.5, 1], [0.167, 1]]],
  [7.09, 3.76, 3.345, 1.64, C.mint, [[0.857, 1], [0.813, 0.967], [0.781, 0.921], [0.748, 0.839], [0.735, 0.756], [0.632, 0.794], [0.53, 0.771], [0.483, 0.727], [0.451, 0.677], [0.399, 0.49], [0.314, 0.48], [0.221, 0.446], [0.132, 0.375], [0.0475, 0.231], [0, 0], [0.346, 0], [0.333, 0.0803], [0.345, 0.169], [0.389, 0.253], [0.429, 0.269], [0.468, 0.247], [0.498, 0.191], [0.514, 0.105], [0.507, 0], [0.857, 0], [0.857, 0.405], [0.904, 0.336], [0.941, 0.331], [0.982, 0.383], [1, 0.483], [0.977, 0.585], [0.936, 0.627], [0.889, 0.612], [0.857, 0.548], [0.857, 1]]],
  [10.026, 3.253, 2.407, 2.879, C.yellow, [[0, 0.176], [0.424, 0.176], [0.38, 0.13], [0.367, 0.0695], [0.407, 0.0152], [0.469, 0], [0.541, 0.0181], [0.573, 0.0638], [0.567, 0.122], [0.523, 0.176], [0.981, 0.176], [1, 0.272], [0.986, 0.348], [0.931, 0.422], [0.815, 0.483], [0.802, 0.609], [0.744, 0.704], [0.679, 0.75], [0.583, 0.785], [0.372, 0.803], [0.61, 1], [0.5, 1], [0.0934, 0.746], [0, 0.746], [0, 0.538], [0.0569, 0.55], [0.138, 0.532], [0.175, 0.502], [0.191, 0.45], [0.171, 0.395], [0.128, 0.36], [0.0626, 0.348], [0, 0.362], [0, 0.176]]],
];

// ------------------------------------------------------------ shared bits ----
const LOREM_LONG = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.';
const LOREM_MED = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ';
const LOREM_AENEAN = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque';

/** Section heading in Poppins Medium 40pt — used on almost every content slide. */
function heading(s, str, x, y, w, h, align) {
  text(s, str, { x, y, w, h, size: 40, font: FONT_H, align: align || 'left' });
}

/** Body paragraph at 12pt / 150% leading. */
function body(s, str, x, y, w, h, align, color) {
  text(s, str, {
    x, y, w, h, size: 12, color: color || C.grey, lnSpc: 1.5, align: align || 'left',
  });
}

/** Navy "Learn More" pill. */
function learnMore(s, x, y) {
  roundRect(s, x, y, 1.281, 0.336, 0.168, C.navy);
  text(s, 'Learn More', {
    x: x + 0.064, y: y + 0.016, w: 1.152, h: 0.303, size: 12, color: C.white, align: 'center',
  });
}

/** Circled check mark used for tick lists and stat callouts. */
function checkBadge(s, x, y, d, color) {
  circle(s, x, y, d, color);
  poly(s, x + 0.2 * d, y + 0.28 * d, 0.6 * d, 0.44 * d,
    [[0, 0.42], [0.15, 0.27], [0.38, 0.52], [0.85, 0], [1, 0.16], [0.38, 0.86]], C.white);
}

/** "o" style bullet row (Courier bullet glyph in the source); indent is in points. */
function bulletRow(s, str, x, y, w, indentPt) {
  text(s, str, {
    x, y, w, h: 0.337, size: 14, valign: 'middle',
    bullet: { characterCode: '006F', indent: indentPt },
  });
}

/**
 * Pale card whose two corners on one side are rounded (round2SameRect rotated 90°).
 * x/y/w/h describe the shape *before* rotation, matching the source geometry.
 */
function sideRoundedCard(s, x, y, w, h, color, transparency) {
  s.addShape('round2SameRect', {
    x, y, w, h, rotate: 90, rectRadius: Math.min(w, h) * 0.16667,
    fill: fill(color, transparency), line: NO_LINE,
  });
}

// ----------------------------------------------------------------- slides ----
function slide01(pres) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  imagePlaceholder(s, 0, 0, 13.333, 7.5, { label: false, color: 'C8C8C8' });
  // Navy scrim over the photo: opaque at the left edge, fading out to the right.
  // Approximated with abutting vertical bands (pptxgenjs has no gradient fill).
  const bands = 48, bandW = 13.333 / bands;
  for (let i = 0; i < bands; i++) {
    const t = (i + 0.5) / bands;                     // 0 at left, 1 at right
    const alpha = t <= 0.58 ? 54 * (t / 0.58) : 54 + 46 * ((t - 0.58) / 0.42);
    rect(s, i * bandW, 1.065, bandW, 6.435, C.navy, Math.round(alpha));
  }
  circle(s, -1.313, -0.088, 2.154, C.white, 85);
  circle(s, -1.76, 5.433, 4.094, C.white, 97);
  circle(s, 10.791, 4.953, 5.395, C.white, 70);

  // Hospital chip
  roundRect(s, 0.939, 2.541, 2.686, 0.604, 0.302, C.white);
  circle(s, 1.007, 2.616, 0.454, C.mint);
  hospitalIcon(s, 1.098, 2.702, 0.272, C.mint);
  text(s, 'Healthywlife Hospital', {
    x: 1.529, y: 2.691, w: 1.969, h: 0.303, size: 12, bold: true, valign: 'middle',
  });

  text(s, [
    { text: 'Medical', color: C.mint, br: true },
    { text: 'Health Care', color: C.white },
  ], { x: 0.842, y: 3.774, w: 9.107, h: 2.74, size: 96, font: FONT_H, bold: true, lnSpc: 0.8 });

  text(s, "DON'T WAIT UNTIL YOU'RE SICK, IT'S BETTER TO CHECK YOUR HEALTH IMMEDIATELY", {
    x: 0.842, y: 6.628, w: 9.646, h: 0.303, size: 12, color: C.white, spc: 3, wrap: false,
  });
  header(s);
}

function slide02(pres) {
  const s = pres.addSlide();
  s.background = { color: C.navy };
  circle(s, 11.349, 0.04, 2.751, C.white, 94);
  circle(s, -0.355, 7.145, 0.709, C.white, 85);

  text(s, 'Table Of Content', {
    x: 0.812, y: 1.667, w: 5.454, h: 0.841, size: 44, font: FONT_H, color: C.white, wrap: false,
  });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus',
    0.9, 2.508, 7.187, 0.678, 'left', 'D9D9D9');

  const items = [
    'Introduction To Medical Healthcare',
    'Importance of Medical Healthcare',
    'Common Health Tips',
    'Role of Medical Professionals',
  ];
  items.forEach((label, i) => {
    const y = 3.749 + i * 0.8005;
    roundRect(s, 0.9, y, 11.083, 0.633, 0.161, C.white, 85);
    text(s, label, { x: 1.337, y: y + 0.098, w: 5.587, h: 0.438, size: 20, color: C.white });
    // chevron-in-circle affordance on the right
    ring(s, 11.349, y + 0.158, 0.317, C.white, 1.5);
    poly(s, 11.427, y + 0.262, 0.161, 0.09,
      [[0, 0], [0.5, 0.72], [1, 0], [1, 0.28], [0.5, 1], [0, 0.28]], C.white);
  });
  header(s);
}

function slide03(pres) {
  const s = pres.addSlide();
  imagePlaceholder(s, 8.907, 0, 4.427, 7.5);
  heading(s, 'Introduction To Medical Healthcare', 0.86, 1.937, 6.187, 1.447);
  body(s, LOREM_LONG, 0.86, 3.489, 7.187, 0.981);

  roundRect(s, 0.86, 6.008, 2.938, 0.526, 0.134, C.pink, 85);
  const ticks = [
    ['General Consultation', C.navy, 4.785],
    ['Emergency Care', C.yellow, 5.183],
    ['Laboratory & Diagnostics', C.mint, 5.58],
    ['Surgery & Intensive Care', C.pinkDeep, 6.102],
  ];
  ticks.forEach(([label, color, y]) => {
    checkBadge(s, 0.984, y, 0.22, color);
    text(s, label, { x: 1.286, y, w: 2.635, h: 0.303, size: 12 });
  });

  roundRect(s, 5.282, 4.772, 4.668, 1.715, 0.208, C.navyPale);
  text(s, 'Overview of the healthcare industry', {
    x: 5.673, y: 5.113, w: 3.78, h: 0.337, size: 14, bold: true, valign: 'middle', wrap: false,
  });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue',
    5.673, 5.469, 3.884, 0.678, 'left', C.greyDark);
  header(s);
}

function slide04(pres) {
  const s = pres.addSlide();
  header(s);
  heading(s, 'Importance Of Medical Healthcare', 0.91, 1.629, 4.062, 2.121);
  body(s, LOREM_MED, 0.91, 3.929, 4.784, 0.981);
  learnMore(s, 0.998, 5.254);

  const cards = [
    ['Early Detection Of Diseases Saves Lives', C.navy, 1.715, 3.983, stethoscopeIcon],
    ['Improves Quality Of Life', C.pink, 3.518, 2.488, familyIcon],
    ['Promotes Overall Well-being', C.mint, 5.321, 2.981, groupIcon],
  ];
  cards.forEach(([title, color, y, titleW, icon]) => {
    roundRect(s, 7.344, y, 5.399, 1.688, 0.24, color, 85);
    text(s, title, {
      x: 8.09, y: y + 0.326, w: titleW, h: 0.337, size: 14, bold: true, valign: 'middle', wrap: false,
    });
    body(s, LOREM_SHORT, 8.09, y + 0.683, 4.43, 0.678, 'left', C.greyDark);
    circle(s, 6.889, y + 0.378, 0.931, color, 0, { color: C.white, width: 4.25 });
    icon(s, 7.128, y + 0.617, 0.45);
  });
}

function slide05(pres) {
  const s = pres.addSlide();
  header(s);
  heading(s, 'Common Health Tips', 3.479, 1.421, 6.376, 0.774, 'center');
  body(s, LOREM_LONG, 1.167, 2.302, 10.999, 0.678, 'center');

  const steps = [
    ['1', 'st', 'Schedule Regular Check-Ups', C.navy, 0.619, umbrellaIcon],
    ['2', 'nd', 'Stay Hydrated', C.pink, 3.746, glassIcon],
    ['3th', '', 'Get Vaccinated', C.yellow, 6.872, nurseIcon],
    ['4th', '', 'Get Enough Sleep', C.mint, 9.933, eyeIcon],
  ];
  steps.forEach(([num, ord, title, color, x, icon]) => {
    roundRect(s, x, 3.716, 2.781, 3.2, 0.395, color, 85);
    circle(s, x + 1.068, 4.095, 0.645, color);
    icon(s, x + 1.229, 4.246, 0.325, color);
    const runs = ord
      ? [{ text: num }, { text: ord, sup: true }, { text: ' Step' }]
      : [{ text: num, sup: true }, { text: ' Step' }];
    text(s, runs, { x: x + 0.482, y: 4.916, w: 1.817, h: 0.346, size: 16, align: 'center', lnSpc: 0.9 });
    text(s, title, {
      x: x + 0.142, y: 5.267, w: 2.497, h: 0.572, size: 14, bold: true, align: 'center',
    });
    text(s, 'PLACEHOLDER', {
      x: x + 0.239, y: 5.829, w: 2.304, h: 0.707, size: 12, color: C.grey, align: 'center',
    });
  });
}

function slide06(pres) {
  const s = pres.addSlide();
  header(s);
  imagePlaceholder(s, 0, 1.789, 4.931, 2.392, { corners: 'right', radius: 0.18 });
  heading(s, 'Role of Medical Professionals', 6.428, 1.789, 4.784, 1.447);

  circle(s, 6.546, 3.635, 0.548, C.navy);
  nurseIcon(s, 6.681, 3.77, 0.28, C.navy);
  text(s, 'Diagnose And Treat Medical Conditions', {
    x: 7.287, y: 3.671, w: 5.187, h: 0.404, size: 18, bold: true,
  });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed',
    7.287, 4.181, 5.436, 0.678);
  [['Doctors, nurses, and specialists', 5.27, 3.552],
    ['Monitor recovery and manage chronic diseases', 5.84, 5.198],
    ['Guide patients on healthy lifestyle choices', 6.374, 4.63],
  ].forEach(([str, y, w]) => bulletRow(s, str, 7.287, y, w, 22.5));

  // Pink card anchored off the left edge, rounded on its right side only.
  sideRoundedCard(s, 2.046, 2.57, 1.991, 6.073, C.pink, 85);
  circle(s, 0.426, 4.916, 0.628, C.pink);
  hospitalIcon(s, 0.568, 5.058, 0.344, C.pink);
  text(s, 'Telemedicine And Virtual Consultations', {
    x: 1.197, y: 4.972, w: 2.901, h: 0.572, size: 14, bold: true, valign: 'middle',
  });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa fusce posuere',
    0.426, 5.618, 5.23, 0.678);
}

function slide07(pres) {
  const s = pres.addSlide();
  header(s);
  imagePlaceholder(s, 0, 5.161, 13.333, 2.339, { corners: 'top', radius: 0.26 });
  heading(s, 'Innovation In Healthcare', 0.91, 1.401, 4.062, 1.447);
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna',
    0.91, 3.031, 6.094, 0.981);
  learnMore(s, 0.998, 4.355);

  roundRect(s, 7.69, 1.874, 5.023, 2.468, 0.35, C.navy, 85);
  ['Use of electronic health records (EHR)', 'AI-assisted diagnostics',
    'Remote patient monitoring', 'Mobile health applications',
  ].forEach((str, i) => bulletRow(s, str, 8.051, 2.309 + i * 0.42, 4.3, 27));
}

function slide08(pres) {
  const s = pres.addSlide();
  header(s);
  heading(s, 'Meet Our Medical Team', 3.049, 1.401, 7.236, 0.774, 'center');
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet',
    2.147, 2.237, 9.039, 0.678, 'center');

  [['Ava Allen', C.navy, C.white], ['Jennifer Harris', C.pink, C.white],
    ['Joshua Martin', C.yellow, C.white], ['Charlotte Nelson', C.mint, C.white],
  ].forEach(([name, color, ink], i) => {
    const x = 1.126 + i * 2.8335;
    imagePlaceholder(s, x, 3.38, 2.581, 3.39, { corners: 'all', radius: 0.083 });
    roundRect(s, x, 6.099, 2.581, 0.671, 0.081, color);
    text(s, name, {
      x: x + 0.145, y: 6.233, w: 2.291, h: 0.404, size: 18, bold: true, color: ink,
      align: 'center', valign: 'middle',
    });
  });
}

function slide09(pres) {
  const s = pres.addSlide();
  header(s);
  // Pale pink stage behind the mock-ups (inherited from the layout in the source).
  sideRoundedCard(s, 0.361, 3.528, 3.616, 4.328, C.pink, 85);
  imagePlaceholder(s, 1.017, 2.058, 2.007, 4.366, { corners: 'all', radius: 0.17 }); // phone
  imagePlaceholder(s, 2.485, 3.884, 2.485, 2.781, { color: 'BFBFBF' });              // watch
  heading(s, 'Medical Healthcare Mockup Design', 6.053, 1.511, 6.094, 1.447);
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna',
    6.053, 3.141, 6.094, 0.981);

  checkBadge(s, 6.209, 4.54, 0.311, C.mintDeep);
  text(s, '90.9%', {
    x: 6.599, y: 4.462, w: 2.187, h: 0.774, size: 44, bold: true, color: C.mintDeep, lnSpc: 0.9,
  });
  text(s, 'Patients Who Were Helped By Early Checkups', {
    x: 6.657, y: 5.154, w: 2.327, h: 0.513, size: 10.5, color: C.greyDark, lnSpc: 1.2,
  });

  roundRect(s, 9.212, 4.378, 2.741, 2.328, 0.327, C.mint, 85);
  checkBadge(s, 9.431, 4.784, 0.19, C.navy);
  text(s, '1.2M+', {
    x: 9.695, y: 4.585, w: 1.784, h: 0.712, size: 40, bold: true, color: C.navy, lnSpc: 0.9,
  });
  text(s, 'Spread In All Cities', {
    x: 9.695, y: 5.217, w: 1.534, h: 0.301, size: 10.5, bold: true, color: C.greyDark, lnSpc: 1.2,
  });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit.', 9.695, 5.518, 2.107, 0.981);
}

function slide10(pres) {
  const s = pres.addSlide();
  header(s);
  heading(s, 'Infographic Section', 3.049, 1.567, 7.236, 0.774, 'center');

  // Left: six stat cards with a coloured spine.
  const stats = [
    ['40', C.navy, 2.074, 2.859], ['50', C.pink, 4.689, 2.859],
    ['60', C.yellow, 2.074, 4.221], ['70', C.mint, 4.689, 4.221],
    ['80', C.peach, 2.074, 5.582], ['90', C.blue, 4.689, 5.582],
  ];
  stats.forEach(([num, color, x, y]) => {
    roundRect(s, x, y, 2.237, 1.18, 0.148, 'F2F2F2', 70);
    roundRect(s, x - 0.006, y - 0.004, 0.089, 1.18, 0.015, color);
    text(s, [
      { text: '\u25B2', size: 28 }, { text: ' ', size: 32 },
      { text: num, size: 28 }, { text: 'K', size: 20 },
    ], { x: x + 0.264, y: y, w: 1.6, h: 0.533, bold: true, color });
    text(s, 'Total Expense', {
      x: x + 0.264, y: y + 0.536, w: 1.6, h: 0.303, size: 12, bold: true, color: C.greyInk,
    });
    text(s, 'Lorem ipsum dolor sit amet', {
      x: x + 0.264, y: y + 0.781, w: 1.947, h: 0.276, size: 8, color: C.greyMid, lnSpc: 1.3,
    });
  });

  // Right: head-in-profile puzzle with percentage labels.
  const pieces = [
    ['tl', 8.283, 2.859, 1.221, 1.305, C.navy], ['tr', 9.594, 2.862, 1.637, 1.303, C.pink],
    ['ml', 7.833, 4.255, 1.671, 1.237, C.yellow], ['mr', 9.594, 4.255, 1.671, 1.237, C.mint],
    ['bl', 7.941, 5.582, 1.563, 1.306, C.peach], ['br', 9.594, 5.582, 1.156, 1.298, C.blue],
  ];
  pieces.forEach(([k, x, y, w, h, color]) => poly(s, x, y, w, h, HEAD_PATHS[k], color));
  [[9.504, 3.36, C.navy, false], [9.504, 6.231, C.peach, false], [9.225, 4.722, C.mint, true]]
    .forEach(([x, y, color, flipH]) =>
      poly(s, x, y, 0.369, 0.304, PUZZLE_NUB, color, { rotate: 180, flipH }));

  [['25', 8.447, 3.314], ['45', 9.78, 3.408], ['55', 8.09, 4.63],
    ['60', 9.74, 4.625], ['65', 8.329, 5.675], ['70', 9.421, 5.688],
  ].forEach(([n, x, y]) => {
    text(s, [{ text: n, size: 28 }, { text: '%', size: 20 }], {
      x, y, w: 1.265, h: 0.509, bold: true, color: C.white, align: 'center', valign: 'middle',
    });
  });
}

function slide11(pres) {
  const s = pres.addSlide();
  header(s);
  heading(s, 'Infographic Section', 6.099, 2.023, 5.783, 0.774);
  body(s, LOREM_AENEAN, 6.099, 2.828, 6.071, 0.506, 'left', C.greyMid);

  // Standing-figure puzzle on the left.
  [['head', 1.92, 1.636, 2.123, 1.725, C.greySoft],
    ['armL', 0.73, 3.345, 2.249, 1.842, C.yellow],
    ['armR', 2.982, 3.345, 2.252, 1.842, C.mint],
    ['legL', 1.489, 5.158, 1.493, 1.775, C.blue],
    ['legR', 2.986, 5.161, 1.493, 1.775, C.pink],
  ].forEach(([k, x, y, w, h, color]) => poly(s, x, y, w, h, BODY_PATHS[k], color));
  poly(s, 2.981, 4.034, 0.467, 0.389, PUZZLE_NUB, C.yellow, { flipH: true });
  poly(s, 2.515, 5.529, 0.467, 0.389, PUZZLE_NUB, C.pink);

  // Four labelled progress bars on the right.
  [['Data 01', '65%', C.navy, 2.624, 3.905], ['Data 02', '70%', C.pink, 2.995, 4.542],
    ['Data 03', '75%', C.yellow, 3.414, 5.179], ['Data 04', '80%', C.mint, 3.726, 5.816],
  ].forEach(([label, pct, color, barW, y]) => {
    roundRect(s, 6.452, y + 0.404, 4.153, 0.071, 0.036, C.greySoft);
    roundRect(s, 6.174, y + 0.404, barW, 0.071, 0.036, color);
    text(s, label, { x: 6.099, y: y + 0.013, w: 1.441, h: 0.505, size: 16, bold: true, color: C.greyMid });
    text(s, pct, {
      x: 6.174 + barW - 0.834, y, w: 0.834, h: 0.505, size: 16, bold: true, color: C.greyMid,
      align: 'center',
    });
  });
}

function slide12(pres) {
  const s = pres.addSlide();
  header(s);
  heading(s, 'Infographic Section', 3.049, 1.567, 7.236, 0.774, 'center');
  HEART_PIECES.forEach(([x, y, w, h, color, pts]) =>
    poly(s, x, y, w, h, pts, color, { line: { color: C.white, width: 3 } }));

  // Four stat cards flanking the heart, mirrored left/right.
  const cards = [
    ['44', 'K', C.navy, 2.756, 3.129, 0.17, 'right'],
    ['65', 'k', C.pink, 8.64, 3.129, 10.634, 'left'],
    ['58', 'k', C.mint, 3.14, 5.173, 0.554, 'right'],
    ['70', 'k', C.yellow, 8.189, 5.173, 10.183, 'left'],
  ];
  cards.forEach(([num, suffix, color, cardX, cardY, textX, align]) => {
    roundRect(s, cardX + 0.162, cardY, 1.593, 1.309, 0.132, C.white);
    text(s, [{ text: num, size: 32 }, { text: suffix, size: 20 }], {
      x: cardX, y: cardY + 0.159, w: 1.916, h: 0.639, bold: true, color, align: 'center',
    });
    text(s, 'Projects total', {
      x: cardX, y: cardY + 0.673, w: 1.916, h: 0.404, size: 12, color: C.greyMid,
      align: 'center', lnSpc: 1.5,
    });
    text(s, 'Your text Here', {
      x: textX, y: cardY - 0.067, w: 2.508, h: 0.505, size: 16, bold: true, color: C.greyMid,
      align, lnSpc: 1.5,
    });
    text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing', {
      x: align === 'right' ? textX + 0.459 : textX, y: cardY + 0.34, w: 2.049, h: 0.889,
      size: 12, color: C.greyMid, align, lnSpc: 1.3,
    });
  });
}

function slide13(pres) {
  const s = pres.addSlide();
  header(s);
  heading(s, 'Infographic Section', -0.036, 1.65, 7.236, 0.774, 'center');
  body(s, LOREM_AENEAN, 0.77, 2.522, 5.678, 0.598, 'left', C.greyMid);

  BRAIN_PIECES.forEach(([x, y, w, h, color, pts]) => poly(s, x, y, w, h, pts, color));
  [['45', 7.674, 2.712], ['25', 9.801, 2.413], ['65', 8.356, 4.122], ['35', 10.272, 4.284]]
    .forEach(([n, x, y]) => {
      text(s, [{ text: n, size: 28 }, { text: '%', size: 20 }], {
        x, y, w: 1.916, h: 0.572, bold: true, font: FONT_H, color: C.white,
        align: 'center', valign: 'middle',
      });
    });

  // Two columns of "Your Text Here" bullets under the intro paragraph.
  [[C.navy, 0.909, 3.953, 1.256, 1.235], [C.pink, 3.768, 3.953, 4.115, 4.094],
    [C.mint, 0.909, 5.399, 1.256, 1.235], [C.yellow, 3.768, 5.399, 4.115, 4.094],
  ].forEach(([color, dotX, dotY, titleX, bodyX]) => {
    circle(s, dotX, dotY, 0.262, color);
    poly(s, dotX + 0.096, dotY + 0.099, 0.07, 0.06, [[0, 0], [1, 0.5], [0, 1]], C.white);
    text(s, 'Your Text Here', {
      x: titleX, y: dotY - 0.071, w: 2.192, h: 0.37, size: 16, bold: true,
      font: FONT_H, color: C.greyInk,
    });
    text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit', {
      x: bodyX, y: dotY + 0.25, w: 2.282, h: 0.626, size: 12, color: C.greyMid, lnSpc: 1.3,
    });
  });
}

function slide14(pres) {
  const s = pres.addSlide();
  header(s);
  imagePlaceholder(s, 0, 3.806, 6.479, 3.694, { corners: 'tr', radius: 0.51 });
  heading(s, 'Get Our Contact Information\u2019s', 0.91, 1.774, 5.569, 1.447);
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero,',
    7.515, 2.007, 5.278, 0.981);

  [['Address', ': 123 Street Name, City Name', C.pink, 3.609, houseIcon],
    ['Phone', ': (123) 123-4567', C.yellow, 4.479, phoneIcon],
    ['Mail', ': info@yourmail.com', C.navy, 5.349, mailIcon],
    ['Website', ': yoursitehere.com', C.mint, 6.22, globeIcon],
  ].forEach(([label, value, color, y, icon]) => {
    roundRect(s, 7.515, y, 5.05, 0.708, 0.354, color, 85);
    circle(s, 7.673, y + 0.123, 0.47, color);
    icon(s, 7.795, y + 0.245, 0.226);
    text(s, label, { x: 8.251, y: y + 0.197, w: 1.089, h: 0.337, size: 14, font: FONT_H });
    text(s, value, { x: 9.321, y: y + 0.197, w: 3.096, h: 0.337, size: 14, font: FONT_H });
  });
}

function slide15(pres) {
  const s = pres.addSlide();
  s.background = { color: C.navy };
  circle(s, 8.551, 4.172, 5.395, C.white, 85); // large glow from the layout
  imagePlaceholder(s, 8.551, 1.341, 4.5, 6.159);
  circle(s, -1.249, -0.23, 2.154, C.white, 85);
  circle(s, 12.24, 3.153, 0.709, C.white, 85);
  header(s);

  text(s, [{ text: 'Thank', color: C.pink }, { text: ' you', color: C.white }], {
    x: 0.551, y: 1.662, w: 8.903, h: 2.036, size: 115, font: FONT_H, bold: true,
  });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet',
    0.884, 3.63, 8.367, 0.678, 'left', 'D9D9D9');

  // "End Presentation" pill (bottom-left)
  roundRect(s, 0.947, 6.242, 3.0, 0.708, 0.354, C.white, 85);
  circle(s, 1.105, 6.365, 0.47, C.mint);
  closeIcon(s, 1.191, 6.45, 0.297, 'FBF8E2');
  text(s, 'End Presentation', {
    x: 1.774, y: 6.439, w: 1.834, h: 0.337, size: 14, bold: true, color: C.white, wrap: false,
  });

  // Yellow speaker chip
  roundRect(s, 7.264, 5.387, 2.686, 0.604, 0.302, C.yellow);
  circle(s, 7.332, 5.463, 0.454, C.white);
  nurseIcon(s, 7.442, 5.573, 0.23, 'E1CD21');
  text(s, 'Dr. Michael Johnson', {
    x: 7.904, y: 5.538, w: 1.885, h: 0.303, size: 12, bold: true, valign: 'middle',
  });
}

// ------------------------------------------------------------------ build ----
function build() {
  const pres = new pptxgen();
  pres.defineLayout({ name: 'WIDE_16x9', width: 13.333, height: 7.5 });
  pres.layout = 'WIDE_16x9';
  pres.theme = { headFontFace: FONT_H, bodyFontFace: FONT_B };
  pres.title = 'Medical Health Care';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
    slide09, slide10, slide11, slide12, slide13, slide14, slide15,
  ].forEach(fn => fn(pres));

  return pres.writeFile({
    fileName: path.join(__dirname, '0eb79ff7-b526-4a4d-bd27-5a35961b70b2_grok_final.pptx'),
  });
}

build().then(f => console.log('wrote', f)).catch(err => {
  console.error(err);
  process.exit(1);
});
