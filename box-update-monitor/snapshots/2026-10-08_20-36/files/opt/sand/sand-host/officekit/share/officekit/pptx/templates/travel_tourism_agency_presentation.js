/* ============================================================================
 * "Travel Tourism" — 16-slide deck, 13.333 x 7.5 in, rebuilt with pptxgenjs.
 *
 * The reference deck leans on raster art: soft blurred colour blobs, a faint
 * world map, line-art travel icons and one phone photograph.  None of that is
 * embedded here — every mark is drawn with native pptxgenjs shapes (see the
 * `glow`, `worldMap`, `ICONS` and `photo` helpers) so the design stays legible
 * in the source.
 * ==========================================================================*/

const pptxgen = require('pptxgenjs');
const path = require('path');

/* --------------------------------------------------------------- palette */
const C = {
  bg: 'F6F8FA',        // slide background
  ink: '000000',
  ink90: '0D0D0D',
  dark: '214856',      // accent2 — deep teal
  lime: 'D9ED82',      // accent1
  limeDeep: 'C1E132',  // accent1 @ lumMod 75%
  limeLine: 'E6F43A',  // dashed decoration rings
  mist: 'CFE2E9',      // accent3
  blueBlob: 'CFE1E8',  // soft blue background glow
  white: 'FFFFFF',
  hair: 'E7E6E6',      // thin cell separators
  rule: 'BFBFBF',      // long horizontal rules (slide 14)
  faint: 'F2F2F2',     // empty star / pale disc
  slate: '32394E',     // dotted connector arcs
  pebble: 'C6CBCF',    // small grey dots
  steel: '44546A',
  land: '2B3B44'       // world-map land tint (drawn at 93% transparency)
};
const HEAD = 'Raleway SemiBold';   // theme major font
const BODY = 'Poppins';            // theme minor font
const SHADOW = { type: 'outer', color: '000000', blur: 18, offset: 3, angle: 90, opacity: 0.11 };
const SOFT_SHADOW = { type: 'outer', color: '000000', blur: 75, offset: 3, angle: 90, opacity: 0.2 };

const LOREM = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean massa commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus ipsum dolor';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ';
const LOREM_TINY = 'Lorem ipsum dolor sit amet, consectetuer';

/* --------------------------------------------------------------- helpers */
/* Text boxes in the source use PowerPoint's default 0.1in side / 0.05in
 * top-bottom insets; pptxgenjs takes them as [left, right, bottom, top] pt. */
const INSETS = [7.2, 7.2, 3.6, 3.6];
const txt = (s, text, o) => s.addText(text, Object.assign({
  fontFace: BODY, fontSize: 12, color: C.ink, valign: 'top', margin: INSETS
}, o));

const rrect = (s, x, y, w, h, radius, fill, shadow) => s.addShape('roundRect', {
  x, y, w, h, rectRadius: radius, fill: { color: fill }, line: { type: 'none' },
  shadow: shadow === undefined ? SHADOW : shadow
});

const disc = (s, x, y, d, fill, shadow) => s.addShape('ellipse', {
  x, y, w: d, h: d, fill: { color: fill }, line: { type: 'none' },
  shadow: shadow === undefined ? SHADOW : shadow
});

/* Soft radial glow.  Nested translucent discs whose per-layer opacity is solved
 * so the stacked coverage follows a gaussian — reproduces the blurred PNG blobs
 * of the original without visible banding. */
function glow(s, x, y, d, color) {
  const N = 14, cx = x + d / 2, cy = y + d / 2;
  let outerCoverage = 0;
  for (let k = N; k >= 1; k--) {
    const r = (k - 0.5) / N;
    const target = 0.8 * Math.exp(-3.4 * r * r);
    const layer = 1 - (1 - target) / (1 - outerCoverage);
    outerCoverage = target;
    const dd = d * k / N;
    s.addShape('ellipse', {
      x: cx - dd / 2, y: cy - dd / 2, w: dd, h: dd,
      fill: { color, transparency: Math.round((1 - layer) * 100) }, line: { type: 'none' }
    });
  }
}

/* Four concentric dashed rings ("radar" motif) filling a square of side d. */
const RING_STEPS = [1, 0.7816, 0.5564, 0.3649];
function rings(s, x, y, d, color) {
  const cx = x + d / 2, cy = y + d / 2;
  RING_STEPS.forEach(f => {
    const dd = d * f;
    s.addShape('ellipse', {
      x: cx - dd / 2, y: cy - dd / 2, w: dd, h: dd,
      fill: { type: 'none' }, line: { color: color || C.limeLine, width: 1.5, dashType: 'dash' }
    });
  });
}

/* Very faint world map wash: [x, y, w, h, rotate] per land mass, expressed as
 * fractions of the map rectangle. */
const CONTINENTS = [
  [0.16, 0.09, 0.15, 0.10, 0],    // Canada
  [0.19, 0.15, 0.11, 0.10, 20],   // United States
  [0.36, 0.07, 0.06, 0.06, 0],    // Greenland
  [0.28, 0.24, 0.07, 0.06, 25],   // Central America
  [0.32, 0.32, 0.07, 0.10, 15],   // northern South America
  [0.33, 0.40, 0.05, 0.14, -8],   // southern South America
  [0.46, 0.13, 0.09, 0.07, 0],    // Europe
  [0.48, 0.22, 0.09, 0.10, 0],    // North Africa
  [0.50, 0.30, 0.06, 0.14, 8],    // southern Africa
  [0.56, 0.09, 0.19, 0.11, 0],    // Russia / north Asia
  [0.60, 0.18, 0.10, 0.09, 0],    // central Asia
  [0.66, 0.24, 0.05, 0.07, 0],    // India
  [0.72, 0.20, 0.07, 0.08, 0],    // China / east Asia
  [0.72, 0.31, 0.09, 0.05, 0],    // south-east Asia
  [0.76, 0.45, 0.09, 0.08, 0]     // Australia
];
function worldMap(s, x, y, w, h) {
  CONTINENTS.forEach(([fx, fy, fw, fh, rotate]) => s.addShape('ellipse', {
    x: x + fx * w, y: y + fy * h, w: fw * w, h: fh * h, rotate,
    fill: { color: C.land, transparency: 96 }, line: { type: 'none' }
  }));
}

/* Free-form silhouettes drawn as polygons in 0..1 glyph space. */
const PLANE = [[1, .50], [.78, .38], [.55, .35], [.30, .05], [.18, .08], [.28, .38], [.08, .42],
  [.02, .34], [0, .38], [0, .62], [.02, .66], [.08, .58], [.28, .62], [.18, .92], [.30, .95],
  [.55, .65], [.78, .62]];
const CURSOR = [[1, 0], [0, .45], [.45, .58], [.58, 1]];
function poly(s, pts, x, y, w, h, color, rotate) {
  s.addShape('custGeom', {
    x, y, w, h, rotate: rotate || 0, fill: { color }, line: { type: 'none' },
    points: pts.map(([fx, fy], i) => ({ x: fx * w, y: fy * h, moveTo: i === 0 })).concat([{ close: true }])
  });
}
const plane = (s, x, y, w, h, rotate, color) => poly(s, PLANE, x, y, w, h, color, rotate);

/* --------------------------------------------------------------- icon set */
/* Every travel icon is a list of [shape, x, y, w, h, mode] in 0..1 glyph
 * space.  Mode: 'f' fill in the icon colour, 'o' outline in the icon colour,
 * 'b' fill in the backdrop colour (knock-out), 'B' outline in the backdrop
 * colour.  'poly' entries carry their own point list. */
const ICONS = {
  plane: [['poly', .02, .10, .96, .80, 'f', PLANE]],
  hotel: [['rect', .12, .22, .76, .74, 'o'], ['rect', .30, .06, .40, .16, 'o'],
    ['rect', .24, .36, .13, .13, 'f'], ['rect', .44, .36, .13, .13, 'f'], ['rect', .64, .36, .13, .13, 'f'],
    ['rect', .24, .58, .13, .13, 'f'], ['rect', .64, .58, .13, .13, 'f'], ['rect', .43, .62, .15, .34, 'o']],
  suitcase: [['roundRect', .06, .30, .40, .64, 'o'], ['roundRect', .50, .20, .42, .74, 'o'],
    ['rect', .18, .20, .16, .10, 'o'], ['rect', .62, .10, .18, .10, 'o']],
  umbrella: [['pie', .03, .20, .94, .92, 'o'], ['rect', .47, .62, .05, .28, 'f'], ['ellipse', .28, .84, .22, .14, 'o']],
  mobile: [['roundRect', .22, .04, .50, .92, 'o'], ['ellipse', .30, .26, .34, .34, 'o'],
    ['rect', .38, .84, .18, .04, 'f'], ['triangle', .56, .50, .32, .42, 'f']],
  cap: [['triangle', .02, .10, .96, .38, 'f'], ['pie', .22, .44, .56, .40, 'f'],
    ['rect', .86, .30, .05, .42, 'f'], ['ellipse', .82, .68, .12, .12, 'f']],
  doc: [['rect', .04, .02, .60, .80, 'f'], ['rect', .13, .16, .40, .05, 'b'], ['rect', .13, .30, .40, .05, 'b'],
    ['rect', .13, .44, .24, .05, 'b'], ['ellipse', .44, .42, .40, .40, 'f'], ['ellipse', .50, .48, .28, .28, 'b'],
    ['rect', .78, .80, .18, .08, 'f']],
  person: [['ellipse', .36, .00, .28, .28, 'f'], ['pie', .22, .24, .56, .34, 'f'],
    ['rect', .04, .50, .92, .08, 'f'], ['pie', .04, .58, .44, .42, 'f'], ['pie', .52, .58, .44, .42, 'f'],
    ['rect', .46, .56, .08, .44, 'b']],
  globe: [['ellipse', .06, .00, .88, .82, 'f'], ['ellipse', .34, .00, .32, .82, 'B'],
    ['rect', .07, .36, .86, .05, 'b'], ['rect', .04, .84, .92, .09, 'f'], ['rect', .22, .80, .56, .05, 'f']],
  ticket: [['roundRect', .02, .30, .60, .36, 'o'], ['roundRect', .36, .44, .60, .36, 'o'],
    ['rect', .10, .44, .30, .05, 'f'], ['rect', .44, .58, .30, .05, 'f'], ['ellipse', .46, .34, .10, .10, 'f']],
  desk: [['ellipse', .30, .04, .22, .22, 'f'], ['pie', .18, .24, .46, .28, 'f'], ['rect', .04, .56, .92, .08, 'f'],
    ['rect', .10, .64, .80, .30, 'o'], ['rect', .66, .14, .28, .34, 'o']]
};
function icon(s, name, x, y, d, color, backdrop) {
  const bg = backdrop || C.bg;
  ICONS[name].forEach(([kind, fx, fy, fw, fh, mode, pts]) => {
    const paint = mode === 'b' || mode === 'B' ? bg : color;
    if (kind === 'poly') { poly(s, pts, x + fx * d, y + fy * d, fw * d, fh * d, paint, 0); return; }
    const filled = mode === 'f' || mode === 'b';
    const o = {
      x: x + fx * d, y: y + fy * d, w: fw * d, h: fh * d,
      fill: filled ? { color: paint } : { type: 'none' },
      line: filled ? { type: 'none' } : { color: paint, width: Math.max(0.6, d * 2) }
    };
    if (kind === 'pie') o.angleRange = [180, 360];
    if (kind === 'roundRect') o.rectRadius = fw * d * 0.25;
    s.addShape(kind, o);
  });
}

/* Map pin. */
function pin(s, x, y, d, color) {
  s.addShape('teardrop', { x, y, w: d, h: d, rotate: 135, fill: { color }, line: { type: 'none' } });
  s.addShape('ellipse', { x: x + d * 0.32, y: y + d * 0.24, w: d * 0.34, h: d * 0.34, fill: { color: C.white }, line: { type: 'none' } });
}

/* Paint a rectangle from a coarse colour grid, bilinearly resampled to
 * cols x rows tiles.  Used for the phone wallpaper on slide 9. */
const mixHex = (a, b, t) => [0, 2, 4].map(i => {
  const va = parseInt(a.substr(i, 2), 16), vb = parseInt(b.substr(i, 2), 16);
  return Math.round(va + (vb - va) * t).toString(16).padStart(2, '0');
}).join('').toUpperCase();

function colorGrid(s, x, y, w, h, grid, cols, rows) {
  const gr = grid.length - 1, gc = grid[0].length - 1;
  for (let j = 0; j < rows; j++) {
    for (let i = 0; i < cols; i++) {
      const py = (j / (rows - 1)) * gr, px = (i / (cols - 1)) * gc;
      const j0 = Math.min(gr - 1, Math.floor(py)), i0 = Math.min(gc - 1, Math.floor(px));
      const top = mixHex(grid[j0][i0], grid[j0][i0 + 1], px - i0);
      const bot = mixHex(grid[j0 + 1][i0], grid[j0 + 1][i0 + 1], px - i0);
      s.addShape('rect', {
        x: x + (i / cols) * w, y: y + (j / rows) * h,
        w: w / cols + 0.008, h: h / rows + 0.008,
        fill: { color: mixHex(top, bot, py - j0) }, line: { type: 'none' }
      });
    }
  }
}

/* Rectangle standing in for a photograph. */
function photo(s, x, y, w, h, o) {
  o = o || {};
  const opts = { x, y, w, h, fill: { color: o.color || 'DCE3E8' }, line: { type: 'none' } };
  if (o.round) opts.rectRadius = o.round;
  if (o.shadow) opts.shadow = o.shadow;
  s.addShape(o.round ? 'roundRect' : 'rect', opts);
}

/* Dark pill reading "Learn More", with a circled arrow. */
function learnMore(s, x, y) {
  rrect(s, x, y, 1.702, 0.434, 0.217, C.dark);
  disc(s, x + 1.322, y + 0.065, 0.303, C.white);
  txt(s, 'Learn More', { x: x + 0.006, y: y + 0.065, w: 1.316, h: 0.303, align: 'center', fontFace: HEAD, color: C.white });
  s.addShape('line', { x: x + 1.384, y: y + 0.217, w: 0.18, h: 0, line: { color: C.dark, width: 1, endArrowType: 'arrow' } });
}

const viewNow = (s, x, y) => txt(s, 'View Now', {
  x, y, w: 1.316, h: 0.303, align: 'center', fontFace: HEAD, underline: { style: 'sng' }
});

/* White pill reading "Explore the world!". */
function explorePill(s, x, y) {
  rrect(s, x, y, 1.885, 0.434, 0.217, C.white);
  txt(s, 'Explore the world!', { x: x + 0.088, y: y + 0.065, w: 1.709, h: 0.303, align: 'center', fontFace: HEAD });
}

/* Four filled stars plus one pale one — a 4.5-of-5 rating. */
function stars(s, x, y) {
  for (let i = 0; i < 5; i++) {
    s.addShape('star5', {
      x: x + i * 0.2132, y, w: 0.192, h: 0.191,
      fill: { color: i < 4 ? C.dark : C.faint }, line: { type: 'none' }
    });
  }
}

/* "1 room | 2 persons | Check-in 13:00WIB" style meta strip. */
function metaRow(s, x, y, cells) {
  let cursor = x;
  cells.forEach((cell, i) => {
    if (i > 0) s.addShape('line', { x: cursor - 0.038, y: y + 0.035, w: 0, h: 0.182, line: { color: C.hair, width: 0.75 } });
    txt(s, cell.t, { x: cursor, y, w: cell.w, h: 0.252, fontSize: 9 });
    cursor += cell.w + 0.076;
  });
}

/* Master furniture: brand mark, top nav, footer URL, page number. */
const NAV = ['Home', 'About Us', 'Service', 'Contact'];
function chrome(s, num) {
  txt(s, 'Travel Tourism', { x: 0.364, y: 0.418, w: 2.5, h: 0.345, fontFace: HEAD, lineSpacingMultiple: 1.3 });
  NAV.forEach((label, i) => txt(s, label, {
    x: 6.755 + i * 1.703, y: 0.417, w: 1.104, h: 0.345,
    align: 'center', fontFace: HEAD, lineSpacingMultiple: 1.3
  }));
  txt(s, 'www.TravelTourism.com', { x: 0.364, y: 6.91, w: 3.5, h: 0.286, fontSize: 11 });
  txt(s, String(num), { x: 11.89, y: 6.91, w: 1.079, h: 0.286, align: 'right', fontSize: 11, bold: true, fontFace: HEAD });
}

/* Centred kicker above a large centred heading. */
function sectionHead(s, kicker, title, y) {
  txt(s, kicker, { x: 5.745, y, w: 1.844, h: 0.37, align: 'center', fontSize: 16, fontFace: HEAD });
  txt(s, title, { x: 2.718, y: y + 0.29, w: 7.898, h: 0.841, align: 'center', fontSize: 44, fontFace: HEAD });
}

/* ======================================================= slide builders */

function slide01(s) {
  glow(s, -1.852, -1.278, 6.375, C.blueBlob);
  worldMap(s, 0, -0.673, 12.75, 8.989);
  [[6.004, 1.432], [6.167, 3.875], [9.255, 2.073], [1.957, 4.407]].forEach(([x, y]) => glow(s, x, y, 3.354, C.blueBlob));
  rings(s, 7.095, 1.26, 4.281);
  rings(s, 10.324, 4.432, 4.281);
  txt(s, 'TRAVELTOURISMPRESENTATION', { x: 0.588, y: 1.62, w: 6.271, h: 0.37, fontSize: 16, charSpacing: 3 });
  txt(s, 'Travel Tourism', { x: 0.588, y: 2.006, w: 7.455, h: 1.212, fontSize: 66, fontFace: HEAD });
  txt(s, LOREM + ' sit amet, consectetuer adipiscing elit. ', { x: 0.588, y: 3.329, w: 6.335, h: 0.87, lineSpacingMultiple: 1.3 });
  learnMore(s, 0.721, 4.669);
  viewNow(s, 2.656, 4.734);
}

function slide02(s) {
  glow(s, -2.715, -1.584, 7.435, C.blueBlob);
  glow(s, 4.219, -1.839, 7.691, C.lime);
  glow(s, 8.67, 2.625, 7.344, C.blueBlob);
  txt(s, 'Introduction to Our Tourism Concept', { x: 0.63, y: 1.589, w: 6.665, h: 1.582, fontSize: 44, fontFace: HEAD });
  disc(s, 7.594, 2.38, 3.917, C.dark);
  txt(s, LOREM, { x: 0.63, y: 3.315, w: 5.11, h: 0.87, lineSpacingMultiple: 1.3 });
  learnMore(s, 0.63, 4.564);
  viewNow(s, 2.487, 4.63);
  rings(s, 5.269, 3.587, 4.281);
  rings(s, 9.266, 1.001, 4.281);
  rrect(s, 9.55, 3.587, 2.84, 2.427, 0.146, C.white);   // photo card (empty in the source)
  rrect(s, 6.036, 3.935, 2.374, 0.526, 0.263, C.white); // floating location chip
  pin(s, 6.222, 3.993, 0.409, C.ink);
  txt(s, 'Hallstatter See', { x: 6.693, y: 4.046, w: 1.531, h: 0.303, fontFace: HEAD });
  txt(s, 'Hallstatter See', { x: 9.657, y: 5.547, w: 1.531, h: 0.303, fontFace: HEAD });
  s.addShape('star5', { x: 11.649, y: 5.594, w: 0.208, h: 0.208, fill: { color: C.lime }, line: { type: 'none' } });
  txt(s, '4.5', { x: 11.867, y: 5.564, w: 0.416, h: 0.269, fontSize: 10, italic: true });
}

const S3_CARDS = [
  { x: 1.811, place: 'Rome, Italty', price: '$5,42k', trip: '10 Days Trip' },
  { x: 5.228, place: 'London, UK', price: '$4.2k', trip: '12 Days Trip' },
  { x: 8.645, place: 'Full Europe', price: '$15k', trip: '28 Days Trip' }
];
function slide03(s) {
  rings(s, 0.284, 4.358, 2.984);
  rings(s, 9.946, 1.941, 2.984);
  glow(s, 8.715, 2.496, 7.435, C.lime);
  glow(s, -2.207, -0.929, 7.435, C.blueBlob);
  sectionHead(s, 'Top Selling', 'Top Travel Destinations', 0.994);
  S3_CARDS.forEach(card => {
    rrect(s, card.x, 2.496, 2.878, 4.011, 0.215, C.white);
    txt(s, card.place, { x: card.x + 0.071, y: 5.681, w: 1.531, h: 0.337, fontSize: 14, fontFace: HEAD });
    txt(s, card.price, { x: card.x + 1.987, y: 5.698, w: 0.821, h: 0.303, align: 'right', fontFace: HEAD });
    poly(s, CURSOR, card.x + 0.177, 6.08, 0.197, 0.205, C.ink, 0);
    txt(s, card.trip, { x: card.x + 0.374, y: 6.048, w: 1.103, h: 0.278, fontSize: 10.5 });
  });
}

function slide04(s) {
  glow(s, 3.724, 2.675, 7.435, C.lime);
  glow(s, -2.296, -1.24, 7.435, C.blueBlob);
  s.addShape('arc', { x: 6.213, y: 1.686, w: 1.559, h: 1.582, angleRange: [55, 305], line: { color: 'C9DC5B', width: 2, dashType: 'sysDot' } });
  s.addShape('arc', { x: 11.55, y: 3.35, w: 1.25, h: 2.0, angleRange: [300, 120], line: { color: C.dark, width: 2, dashType: 'sysDot' } });
  plane(s, 7.569, 1.474, 0.427, 0.375, 0, C.lime);
  disc(s, 8.255, 1.256, 0.479, C.white);
  icon(s, 'hotel', 8.362, 1.362, 0.265, C.ink, C.white);
  plane(s, 12.389, 3.141, 0.427, 0.375, 248, C.dark);
  explorePill(s, 0.781, 1.858);
  txt(s, 'From Southeast Asia to the World.', { x: 0.781, y: 2.406, w: 5.709, h: 1.582, fontSize: 44, fontFace: HEAD });
  txt(s, LOREM, { x: 0.781, y: 4.102, w: 5.11, h: 0.87, lineSpacingMultiple: 1.3 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean massa commodo ligula eget dolor. Aenean',
    { x: 0.781, y: 5.035, w: 5.11, h: 0.608, lineSpacingMultiple: 1.3 });
}

const S5_ROOM_CELLS = [{ t: '1 room', w: 0.608 }, { t: '2 persons', w: 0.898 }, { t: 'Check-in 13:00WIB', w: 1.361 }];
const S5_TICKETS = [{ x: 2.684, y: 2.988 }, { x: 8.03, y: 3.289 }, { x: 6.504, y: 5.192 }];
/* Three dashed arcs chain together into the looping flight route. */
const S5_ARCS = [
  { x: 1.749, y: 2.791, d: 2.785, rotate: 218.37, range: [270, 80.41] },
  { x: 4.142, y: 1.543, d: 3.946, rotate: 90, range: [350.94, 80.41] },
  { x: 7.026, y: 2.675, d: 2.785, rotate: 255.42, flipH: true, range: [343.4, 80.41] }
];
function slide05(s) {
  glow(s, 7.918, -1.603, 7.435, C.lime);
  glow(s, -1.974, 1.41, 7.435, C.blueBlob);
  worldMap(s, 1.306, 1.36, 10.721, 7.558);
  txt(s, 'Plan Your Vacation!', { x: 3.405, y: 1.051, w: 6.524, h: 0.841, align: 'center', fontSize: 44, fontFace: HEAD });
  txt(s, LOREM, { x: 2.648, y: 1.893, w: 8.038, h: 0.608, align: 'center', lineSpacingMultiple: 1.3 });
  S5_ARCS.forEach(a => s.addShape('arc', {
    x: a.x, y: a.y, w: a.d, h: a.d, rotate: a.rotate, flipH: !!a.flipH,
    angleRange: a.range, line: { color: C.lime, width: 2, dashType: 'dash' }
  }));
  S5_TICKETS.forEach(t => {
    rrect(s, t.x, t.y, 3.895, 0.762, 0.057, C.white);
    txt(s, 'Rome, Italty', { x: t.x + 0.743, y: t.y + 0.114, w: 1.531, h: 0.303, fontFace: HEAD });
    metaRow(s, t.x + 0.748, t.y + 0.43, S5_ROOM_CELLS);
  });
  rrect(s, 1.408, 5.329, 3.475, 0.762, 0.057, C.white);
  txt(s, 'CGK - BDO', { x: 1.54, y: 5.425, w: 1.531, h: 0.303, fontFace: HEAD });
  txt(s, 'Pesawat', { x: 2.547, y: 5.451, w: 0.785, h: 0.252, fontSize: 9, italic: true, color: C.steel });
  metaRow(s, 1.54, 5.741, [{ t: 'ID-762 ( Batik Air) ', w: 1.323 }, { t: 'Ekonomi', w: 0.785 }, { t: '13:00WIB', w: 0.8 }]);
  plane(s, 4.271, 4.496, 0.687, 0.603, 58, C.lime);
}

const S6_STEPS = [
  { y: 3.23, title: 'Choose Destination ', glyph: 'hotel', highlight: false },
  { y: 4.457, title: 'Make payment', glyph: 'suitcase', highlight: true },
  { y: 5.684, title: 'Reach Airport on Selected Date', glyph: 'umbrella', highlight: false }
];
function slide06(s) {
  glow(s, -1.969, -1.829, 7.435, C.blueBlob);
  glow(s, 3.661, 1.363, 8.832, C.lime);
  glow(s, 6.117, -1.186, 8.832, C.lime);
  txt(s, 'Book Your Next Trip In 3 Easy Steps', { x: 0.961, y: 1.147, w: 5.9, h: 1.582, fontSize: 44, fontFace: HEAD });
  rrect(s, 0.961, 4.327, 4.783, 0.907, 0.068, C.lime);   // highlighted step band
  S6_STEPS.forEach(step => {
    if (step.highlight) {
      icon(s, step.glyph, 1.143, 4.529, 0.505, C.ink, C.lime);
    } else {
      rrect(s, 1.08, step.y + 0.074, 0.5, 0.5, 0.083, C.white);
      icon(s, step.glyph, 1.185, step.y + 0.179, 0.291, C.ink, C.white);
    }
    txt(s, step.title, { x: 1.829, y: step.y, w: 3.741, h: 0.337, fontSize: 14, fontFace: HEAD });
    txt(s, LOREM_TINY, { x: 1.829, y: step.y + 0.303, w: 4.106, h: 0.345, lineSpacingMultiple: 1.3 });
  });
  rrect(s, 7.485, 1.479, 4.173, 4.778, 0.19, C.white);   // trip card
  txt(s, 'Trip To Grace ', { x: 7.672, y: 4.524, w: 1.577, h: 0.337, fontSize: 14, fontFace: HEAD });
  metaRow(s, 7.672, 4.88, [{ t: '14-29 June', w: 0.898 }, { t: 'By Robbin Jr', w: 1.361 }]);
  [['ticket', 7.751, C.white], ['desk', 8.29, C.dark], ['desk', 8.829, C.white]].forEach(([glyph, x, fill], i) => {
    disc(s, x, 5.263, 0.42, fill);
    icon(s, glyph, x + 0.094, 5.356, 0.232, i === 1 ? C.white : C.ink, fill);
  });
  txt(s, '24 People going', { x: 7.672, y: 5.779, w: 1.361, h: 0.252, fontSize: 9 });
  rrect(s, 9.868, 4.734, 2.504, 0.854, 0.064, C.white); // progress callout
  txt(s, 'Trip To Rome  ', { x: 10.547, y: 4.867, w: 1.577, h: 0.337, fontSize: 14, fontFace: HEAD });
  txt(s, '40% Completed', { x: 10.547, y: 5.204, w: 1.361, h: 0.252, fontSize: 9 });
}

const S7_CARDS = [
  { x: 0.806, title: 'Calculated Weather', glyph: 'suitcase', card: C.white, tile: C.lime },
  { x: 3.778, title: 'Best Flights', glyph: 'hotel', card: C.lime, tile: C.white },
  { x: 6.75, title: 'Local Events', glyph: 'umbrella', card: C.white, tile: C.lime },
  { x: 9.722, title: 'Customization', glyph: 'mobile', card: C.white, tile: C.lime }
];
function slide07(s) {
  glow(s, -2.207, -0.929, 7.435, C.blueBlob);
  rings(s, 0.284, 4.358, 2.984);
  rings(s, 9.946, 1.941, 2.984);
  glow(s, 8.715, 2.496, 7.435, C.lime);
  sectionHead(s, 'Category', 'We Offer Best Services', 1.099);
  S7_CARDS.forEach(card => {
    rrect(s, card.x, 2.675, 2.806, 3.541, 0.2, card.card);
    rrect(s, card.x + 0.294, 2.986, 0.828, 0.828, 0.138, card.tile, card.tile === C.white ? SHADOW : null);
    icon(s, card.glyph, card.x + 0.456, 3.148, 0.505, C.ink, card.tile);
    txt(s, card.title, { x: card.x + 0.222, y: 4.127, w: 2.074, h: 0.337, fontSize: 14, fontFace: HEAD });
    txt(s, 'Lorem ipsum dolor sit amet, elit consectetuer adipiscing elit. Aenean massa commodo ligula eget dolor. Aenean',
      { x: card.x + 0.222, y: 4.509, w: 2.35, h: 1.395, align: 'justify', lineSpacingMultiple: 1.3 });
  });
}

function slide08(s) {
  s.addShape('ellipse', { x: 6.952, y: 1.216, w: 5.069, h: 5.069, fill: { type: 'none' }, line: { color: C.lime, width: 1.5, dashType: 'dash' } });
  explorePill(s, 0.765, 1.216);
  txt(s, 'Meet Our Travel Experts and Guides', { x: 0.765, y: 1.793, w: 6.524, h: 1.582, fontSize: 44, fontFace: HEAD });
  plane(s, 7.72, 1.417, 0.636, 0.558, 332.84, C.lime);
  plane(s, 11.279, 4.897, 0.636, 0.558, 118.38, C.lime);
  rrect(s, 6.945, 3.617, 2.084, 0.526, 0.263, C.white);  // rating chip
  stars(s, 7.141, 3.784);
  txt(s, '4.5', { x: 8.258, y: 3.694, w: 0.574, h: 0.337, fontSize: 14, fontFace: HEAD });
  rrect(s, 6.077, 4.328, 2.084, 0.526, 0.263, C.white);  // name chip
  icon(s, 'desk', 6.347, 4.447, 0.288, C.ink, C.white);
  txt(s, 'Mark Medison', { x: 6.649, y: 4.439, w: 1.433, h: 0.303, fontFace: HEAD });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing. Aenean massa commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus ipsum',
    { x: 0.77, y: 3.518, w: 5.11, h: 0.87, lineSpacingMultiple: 1.3 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing. Aenean massa commodo ligula eget dolor. ',
    { x: 0.77, y: 4.451, w: 5.11, h: 0.608, lineSpacingMultiple: 1.3 });
  learnMore(s, 0.853, 5.392);
}

/* Phone mock-up.  The reference drops a photograph in here; its swirling
 * wallpaper is approximated by resampling this 5x3 colour grid. */
const PHONE_SCREEN = [
  ['E4E7F0', 'C08D90', 'B45D77', 'DC2A83', 'E62860'],
  ['FBEFE9', 'F7E6D9', 'F4C3A2', 'E94166', 'E83380'],
  ['E97060', 'F1BEAF', 'EEDBDC', '9EAAC7', '3258AB']
];
function slide09(s) {
  worldMap(s, -0.375, -0.803, 14.271, 10.061);
  glow(s, -3.257, -2.345, 9.005, C.blueBlob);
  glow(s, 6.667, 0.757, 6.375, C.blueBlob);
  rings(s, 7.792, 4.34, 2.984);
  rings(s, 4.099, 1.711, 2.984);
  photo(s, 1.462, 3.501, 5.59, 2.872, { round: 0.32, color: '18191B', shadow: SOFT_SHADOW });
  colorGrid(s, 1.60, 3.632, 5.31, 2.61, PHONE_SCREEN, 24, 12);
}

const S10_SEGMENTS = [
  { range: [180, 224], color: C.limeDeep }, { range: [227, 269], color: C.lime },
  { range: [272, 316], color: C.limeDeep }, { range: [319, 360], color: C.lime }
];
const S10_LABELS = [
  { num: '01', glyph: 'cap', seg: C.limeDeep, ix: 4.217, iy: 5.046, id: 0.541, numX: 1.931, numY: 5.186, ttlX: 0.527, ttlY: 5.358, bodyX: 0.542, bodyY: 5.713, align: 'right' },
  { num: '02', glyph: 'doc', seg: C.lime, ix: 5.73, iy: 3.535, id: 0.541, numX: 2.439, numY: 3.356, ttlX: 1.035, ttlY: 3.527, bodyX: 1.051, bodyY: 3.883, align: 'right' },
  { num: '03', glyph: 'person', seg: C.limeDeep, ix: 7.09, iy: 3.559, id: 0.541, numX: 9.92, numY: 3.32, ttlX: 10.582, ttlY: 3.491, bodyX: 9.92, bodyY: 3.847, align: 'left' },
  { num: '04', glyph: 'globe', seg: C.lime, ix: 8.601, iy: 5.072, id: 0.539, numX: 10.379, numY: 5.151, ttlX: 11.041, ttlY: 5.322, bodyX: 10.379, bodyY: 5.678, align: 'left' }
];
function slide10(s) {
  S10_SEGMENTS.forEach(seg => s.addShape('blockArc', {
    x: 3.663, y: 2.954, w: 5.979, h: 5.979, angleRange: seg.range, arcThicknessRatio: 0.515,
    fill: { color: seg.color }, line: { type: 'none' }, shadow: SHADOW
  }));
  [[9.454, 5.085, 0.812], [7.847, 3.236, 2.0]].forEach(([x, y, w]) => s.addShape('line', {
    x, y, w, h: 0, line: { color: C.ink, width: 2, dashType: 'dash', endArrowType: 'oval' }
  }));
  [[3.067, 5.085, 0.812], [3.488, 3.236, 2.0]].forEach(([x, y, w]) => s.addShape('line', {
    x, y, w, h: 0, line: { color: C.ink, width: 2, dashType: 'dash', beginArrowType: 'oval' }
  }));
  S10_LABELS.forEach(l => {
    icon(s, l.glyph, l.ix, l.iy, l.id, C.white, l.seg);
    txt(s, l.num, { x: l.numX, y: l.numY, w: 1.024, h: 0.643, fontSize: 32, bold: true, fontFace: HEAD, align: l.align });
    txt(s, 'Subtitle Here', { x: l.ttlX, y: l.ttlY, w: 1.765, h: 0.362, fontSize: 14, bold: true, fontFace: HEAD, align: l.align, lineSpacingMultiple: 1.2 });
    txt(s, LOREM_TINY, { x: l.bodyX, y: l.bodyY, w: 2.412, h: 0.608, align: l.align, lineSpacingMultiple: 1.3 });
  });
  txt(s, 'Infographic Selection', { x: 1.165, y: 1.156, w: 11.002, h: 0.841, align: 'center', fontSize: 44, fontFace: HEAD });
}

const S11_TILES = [
  { x: 0.802, y: 2.857, fill: C.white, shadow: true, num: '149', numX: 0.95, numY: 3.037, numW: 2.938, bodyX: 0.995, bodyY: 4.927 },
  { x: 4.962, y: 2.094, fill: C.lime, shadow: false, num: '217', numX: 5.306, numY: 2.274, numW: 2.504, bodyX: 5.133, bodyY: 4.164 },
  { x: 9.122, y: 1.394, fill: C.white, shadow: true, num: '312', numX: 9.47, numY: 1.574, numW: 2.505, bodyX: 9.298, bodyY: 3.464 }
];
function slide11(s) {
  S11_TILES.forEach(t => {
    s.addShape('rect', { x: t.x, y: t.y, w: 3.281, h: 3.249, fill: { color: t.fill }, line: { type: 'none' }, shadow: t.shadow ? SOFT_SHADOW : null });
    s.addText([
      { text: '+', options: { fontSize: 36 } },
      { text: t.num, options: { fontSize: 66 } },
      { text: '%', options: { fontSize: 60 } }
    ], { x: t.numX, y: t.numY, w: t.numW, h: 1.212, align: 'center', valign: 'top', fontFace: HEAD, color: C.ink });
    txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget',
      { x: t.bodyX, y: t.bodyY, w: 2.938, h: 0.87, lineSpacingMultiple: 1.3 });
  });
  // hooked connectors between the tiles
  s.addShape('arc', { x: 3.905, y: 1.613, w: 1.61, h: 1.691, rotate: 330.32, angleRange: [219.62, 316.17], line: { color: C.slate, width: 1.5, transparency: 80 } });
  disc(s, 3.825, 2.415, 0.198, C.pebble, null);
  disc(s, 5.03, 1.7, 0.198, C.lime, null);
  s.addShape('arc', { x: 7.837, y: 4.182, w: 1.61, h: 1.691, rotate: 150.32, angleRange: [219.62, 316.17], line: { color: C.slate, width: 1.5, transparency: 80 } });
  disc(s, 9.33, 4.874, 0.198, C.pebble, null);
  disc(s, 8.125, 5.588, 0.198, C.lime, null);
  txt(s, LOREM_SHORT, { x: 0.808, y: 1.539, w: 2.901, h: 0.608, lineSpacingMultiple: 1.3 });
  txt(s, 'Lorem ipsum dolor a amet consectetuer adipiscing elit ', { x: 9.631, y: 5.39, w: 2.901, h: 0.608, lineSpacingMultiple: 1.3 });
}

const S12_NODES = [
  { num: '01', x: 4.236, y: 5.113, fill: C.white, numX: 4.501, numY: 5.41, numW: 0.637 },
  { num: '02', x: 5.022, y: 3.457, fill: C.lime, numX: 5.263, numY: 3.754, numW: 0.684 },
  { num: '03', x: 6.852, y: 3.457, fill: C.white, numX: 7.074, numY: 3.754, numW: 0.723 },
  { num: '04', x: 7.686, y: 5.113, fill: C.lime, numX: 7.937, numY: 5.41, numW: 0.663 }
];
const S12_TEXTS = [
  { title: 'Leadership', tX: 1.825, tY: 5.124, tW: 2.134, bX: 1.111, bY: 5.494, bW: 2.849, align: 'right' },
  { title: 'People', tX: 1.92, tY: 3.44, tW: 2.944, bX: 2.015, bY: 3.78, bW: 2.849, align: 'right' },
  { title: 'Strategy', tX: 8.275, tY: 3.41, tW: 2.944, bX: 8.275, bY: 3.78, bW: 2.944, align: 'left' },
  { title: 'Implementation', tX: 9.179, tY: 5.124, tW: 2.473, bX: 9.179, bY: 5.494, bW: 2.944, align: 'left' }
];
function slide12(s) {
  txt(s, 'Section Infographic', { x: 1.023, y: 1.104, w: 11.287, h: 0.767, align: 'center', fontSize: 44, fontFace: HEAD, lineSpacingMultiple: 0.9 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus dolor sit amet, ',
    { x: 2.976, y: 1.997, w: 7.382, h: 0.608, align: 'center', lineSpacingMultiple: 1.3 });
  [[6.829, 4.853, 203.88], [6.077, 4.835, 156.12], [7.263, 5.575, 270], [5.647, 5.575, 90]].forEach(([x, y, rotate]) =>
    s.addShape('downArrow', { x, y, w: 0.233, h: 0.241, rotate, fill: { color: C.lime }, line: { type: 'none' } }));
  S12_NODES.forEach(n => {
    disc(s, n.x, n.y, 1.167, n.fill, n.fill === C.white ? SOFT_SHADOW : null);
    txt(s, n.num, { x: n.numX, y: n.numY, w: n.numW, h: 0.572, align: 'center', fontSize: 28, fontFace: HEAD });
  });
  disc(s, 5.986, 5.113, 1.167, C.white, SOFT_SHADOW);
  txt(s, 'Target', { x: 5.955, y: 5.504, w: 1.228, h: 0.37, fontSize: 16, color: C.ink90 });
  S12_TEXTS.forEach(t => {
    txt(s, t.title, { x: t.tX, y: t.tY, w: t.tW, h: 0.404, fontSize: 18, fontFace: HEAD, color: C.ink90, align: t.align, lineSpacingMultiple: 1.0 });
    txt(s, LOREM_SHORT, { x: t.bX, y: t.bY, w: t.bW, h: 0.608, color: C.ink90, align: t.align, lineSpacingMultiple: 1.3 });
  });
}

const S13_NODES = [
  { glyph: 'plane', cx: 4.298, cy: 1.734, d: 1.551, ix: 4.75, iy: 2.186, id: 0.647, fill: C.lime },
  { glyph: 'suitcase', cx: 3.766, cy: 3.603, d: 0.913, ix: 3.986, iy: 3.824, id: 0.472, fill: C.dark },
  { glyph: 'ticket', cx: 4.094, cy: 5.025, d: 0.862, ix: 4.268, iy: 5.199, id: 0.514, fill: C.mist }
];
const S13_BRANCHES = [
  { ttlX: 1.85, ttlY: 1.817, bodyX: 1.521, bodyY: 2.156, dotX: 3.519, dotY: 2.155, dot: C.lime },
  { ttlX: 1.366, ttlY: 3.423, bodyX: 1.036, bodyY: 3.763, dotX: 3.035, dotY: 3.762, dot: C.dark },
  { ttlX: 1.636, ttlY: 5.129, bodyX: 1.306, bodyY: 5.468, dotX: 3.305, dotY: 5.467, dot: C.mist }
];
function slide13(s) {
  s.addShape('ellipse', { x: 4.186, y: 1.871, w: 4.602, h: 4.687, fill: { type: 'none' }, line: { color: 'A6A6A6', width: 2, dashType: 'sysDot' } });
  disc(s, 6.693, 1.93, 3.095, C.faint, null);
  txt(s, 'Your Title', { x: 7.176, y: 2.522, w: 2.088, h: 0.724, align: 'center', valign: 'middle', fontSize: 24, fontFace: HEAD });
  txt(s, LOREM_SHORT, { x: 7.216, y: 3.433, w: 2.088, h: 1.0, align: 'center', lineSpacingMultiple: 1.2 });
  s.addShape('ellipse', { x: 6.544, y: 1.791, w: 3.392, h: 3.392, fill: { type: 'none' }, line: { color: 'F5F5F7', width: 1, dashType: 'dash' } });
  S13_NODES.forEach(n => {
    disc(s, n.cx, n.cy, n.d, n.fill, null);
    icon(s, n.glyph, n.ix, n.iy, n.id, C.white, n.fill);
  });
  S13_BRANCHES.forEach(b => {
    txt(s, 'Description', { x: b.ttlX, y: b.ttlY, w: 1.508, h: 0.339, align: 'right', fontSize: 16, fontFace: HEAD });
    txt(s, LOREM_TINY, { x: b.bodyX, y: b.bodyY, w: 1.822, h: 0.503, align: 'right', fontSize: 11, lineSpacingMultiple: 1.3 });
    disc(s, b.dotX, b.dotY, 0.152, b.dot, null);
  });
  txt(s, 'Selection Infographic', { x: 5.971, y: 5.129, w: 6.783, h: 0.841, fontSize: 44, fontFace: HEAD });
}

const S14_ROWS = [{ y: 1.375, value: '$500k' }, { y: 2.807, value: '68%' },
  { y: 4.239, value: '$300K' }, { y: 5.671, value: '$100k' }];
function slide14(s) {
  s.addShape('line', { x: 9.869, y: 2.849, w: 0.882, h: 0, line: { color: C.white, width: 1.5, dashType: 'dash' } });
  s.addShape('line', { x: 5.3, y: 4.339, w: 0.882, h: 0, line: { color: C.white, width: 1.5, dashType: 'dash' } });
  txt(s, 'Growth Roadmap for Tourism Services', { x: 0.506, y: 1.375, w: 5.536, h: 2.322, fontSize: 44, fontFace: HEAD });
  S14_ROWS.forEach(r => {
    txt(s, r.value, { x: 10.31, y: r.y, w: 2.67, h: 0.909, align: 'right', fontSize: 48, fontFace: HEAD });
    txt(s, 'Your Description Here', { x: 6.583, y: r.y + 0.268, w: 2.346, h: 0.372, fontSize: 14, fontFace: HEAD, lineSpacingMultiple: 1.3 });
    s.addShape('line', { x: 6.583, y: r.y + 0.909, w: 6.33, h: 0, line: { color: C.rule, width: 0.75 } });
  });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus Aenean commodo ligula eget',
    { x: 0.506, y: 4.428, w: 5.257, h: 0.87, align: 'justify', lineSpacingMultiple: 1.3 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum',
    { x: 0.506, y: 5.331, w: 5.257, h: 0.608, align: 'justify', lineSpacingMultiple: 1.3 });
}

function slide15(s) {
  rrect(s, 1.605, 2.633, 10.122, 3.5, 0.165, C.white);
  sectionHead(s, 'Testimonials', 'Trust Our Clients', 0.994);
  [[11.316, false], [1.194, true]].forEach(([x, flipH]) => {
    disc(s, x, 3.971, 0.823, C.lime);
    s.addShape('rightArrow', { x: x + 0.247, y: 4.301, w: 0.33, h: 0.163, flipH, fill: { color: C.white }, line: { type: 'none' } });
  });
  txt(s, 'Mark Medison', { x: 5.631, y: 4.208, w: 2.07, h: 0.337, align: 'center', fontSize: 14, fontFace: HEAD });
  stars(s, 5.821, 4.636);
  txt(s, '4.5', { x: 6.938, y: 4.545, w: 0.574, h: 0.337, fontSize: 14, fontFace: HEAD });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing. Aenean massa commodo ligula eget dolor. Aenean massa. Cum sociis natoque',
    { x: 3.577, y: 5.016, w: 6.18, h: 0.608, align: 'center', lineSpacingMultiple: 1.3 });
}

function slide16(s) {
  glow(s, 8.03, 1.94, 6.375, C.lime);
  glow(s, -1.852, -1.278, 6.375, C.blueBlob);
  worldMap(s, 0, -0.673, 12.75, 8.989);
  txt(s, 'TRAVELTOURISMPRESENTATION', { x: 0.754, y: 1.513, w: 5.582, h: 0.37, fontSize: 16, charSpacing: 3 });
  txt(s, 'Thank You!!', { x: 0.754, y: 1.775, w: 5.771, h: 1.212, fontSize: 66, fontFace: HEAD });
  txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean massa commodo ligula eget dolor. Aenean massa. Cum sociis',
    { x: 6.853, y: 1.7, w: 5.726, h: 0.608, lineSpacingMultiple: 1.3 });
  txt(s, 'From 2025', { x: 6.852, y: 2.426, w: 1.177, h: 0.375, lineSpacingMultiple: 1.5 });
  txt(s, '+123 456 7890', { x: 8.357, y: 2.441, w: 1.868, h: 0.345, align: 'center', lineSpacingMultiple: 1.3 });
  txt(s, '12 Your Street Name', { x: 10.552, y: 2.441, w: 1.868, h: 0.345, align: 'right', lineSpacingMultiple: 1.3 });
}

/* ------------------------------------------------------------------ build */
const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
  slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16];

const pres = new pptxgen();
pres.defineLayout({ name: 'W16x9', width: 13.333, height: 7.5 });
pres.layout = 'W16x9';
pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };
pres.author = 'Travel Tourism';
pres.title = 'Travel Tourism';

BUILDERS.forEach((build, i) => {
  const slide = pres.addSlide();
  slide.background = { color: C.bg };
  chrome(slide, i + 1);   // master furniture sits *under* the slide artwork
  build(slide);
});

pres.writeFile({ fileName: path.join(__dirname, '1072e4f1-618c-4463-94e7-a380b55949a3_grok_final.pptx') })
  .then(f => console.log('wrote', f));
