/**
 * Recreation of "BUILDINGZ — Multipurpose Presentation Template" (23 slides,
 * 26.6" x 15") using pptxgenjs only.
 *
 * The original deck's flat-vector illustrations are rebuilt from plain shape
 * tables (see ILLUS below); its raster photo placeholders are replaced by
 * simple placeholder rectangles.
 *
 * Run: node 13e96605-dab3-4250-9f85-b06554d25c85_grok_final.js
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------------------
// Deck constants
// ---------------------------------------------------------------------------
const SLIDE_W = 26.6007;
const SLIDE_H = 15.0;

// Theme palette ("Red Orange")
const C = {
  bg1: 'FFFFFF',
  tx2: '505046',
  accent1: 'E84C22', // red-orange
  accent2: 'FFBD47', // light amber
  accent3: 'B64926', // brick
  accent4: 'FF8427', // orange
  accent5: 'CC9900', // gold
  accent6: 'B22600', // deep red
};

const FONT = 'Lato';
const FONT_BLACK = 'Lato Black';
const FONT_LIGHT = 'Lato Light';

// Text-frame insets (points) used by the original text boxes.
const M_BODY = [17.28, 17.28, 8.64, 8.64]; // l, r, b, t
const M_HEAD = [14.4, 14.4, 7.2, 7.2];

// Repeated copy.
const BODY_1 = 'PLACEHOLDER';
const BODY_2 = 'Lorem Ipsum has two main data statistical this methodologies important.';
const BODY_3 = 'Lorem Ipsum has two main data statistical this methodologies';
const BODY_4 = 'Lorem Ipsum has two main data statistical';
const BODY_5 = 'Lorem Ipsum has two main data statistical important.';
const LOREM_INTRO = ', consectetur adipiscing elit. Ut efficitur ipsum vitae tortor accumsan, ' +
  'a pulvinar lorem lacinia. Donec eu arcu justo. Fusce eget consequat risus Proin';
const LOREM_LONG = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Aliquam tincidunt ante nec ' +
  'sem congue convallis. Pellentesque vel mauris quis nisl ornare rutrum in id risus. Proin vehicula ' +
  'ut sem et tempus. Interdum et malesuada fames ac ante ipsum primis in faucibus. Pellentesque.';
const LOREM_MED = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Aliquam tincidunt ante nec ' +
  'sem congue convallis. Pellentesque vel mauris quis nisl ornare rutrum in id risus. Proin vehicula ' +
  'ut sem et tempus. ';

// ---------------------------------------------------------------------------
// Drawing helpers
// ---------------------------------------------------------------------------
let SH; // pptx.ShapeType, assigned once the deck is created

const NOLINE = { type: 'none' };

function rect(s, x, y, w, h, fill, opts = {}) {
  s.addShape(SH.rect, { x, y, w, h, fill: { color: fill }, line: NOLINE, ...opts });
}

function oval(s, x, y, w, h, fill, opts = {}) {
  s.addShape(SH.ellipse, { x, y, w, h, fill: { color: fill }, line: NOLINE, ...opts });
}

// PowerPoint anchors text boxes to the top; pptxgenjs defaults to centred.
function text(s, runs, o) {
  s.addText(runs, { fontFace: FONT, color: C.bg1, isTextBox: true, valign: 'top', ...o });
}

/** Shorthand for a styled run inside a mixed-format paragraph. */
function run(t, size, color, face = FONT, bold = false, extra = {}) {
  return { text: t, options: { fontSize: size, color, fontFace: face, bold, ...extra } };
}

/**
 * Banner with two small corner notches on the right edge — the deck's
 * "Freeform 6" ribbon. `flipH` mirrors it for left-facing rows.
 */
function ribbon(s, x, y, w, h, color, flipH = false) {
  const n = Math.min(0.33, w / 30);
  s.addShape(SH.custGeom, {
    x, y, w, h, flipH, fill: { color }, line: NOLINE,
    points: [
      { x: 0, y: 0 }, { x: w - n, y: 0 }, { x: w, y: n },
      { x: w, y: h - n }, { x: w - n, y: h }, { x: 0, y: h }, { close: true },
    ],
  });
}

/** Left- or right-pointing arrow band (prstGeom homePlate). */
function chevron(s, x, y, w, h, color, pointsLeft = false) {
  s.addShape(SH.homePlate, {
    x, y, w, h, rotate: pointsLeft ? 180 : 0, fill: { color }, line: NOLINE,
  });
}

/** Slim vertical accent bar placed beside a statistic. */
function accentBar(s, x, y, w, h, color) {
  s.addShape(SH.round2SameRect, {
    x, y, w, h, rotate: 180, flipH: true, fill: { color }, line: NOLINE,
  });
}

/** Stand-in for an empty picture placeholder in the original template. */
function photoBox(s, x, y, w, h) {
  s.addShape(SH.rect, { x, y, w, h, fill: { color: '353535' }, line: NOLINE });
  s.addText('[image]', {
    x, y, w, h, align: 'center', valign: 'middle',
    fontFace: FONT, fontSize: 14, color: '4A4A4A',
  });
}

// ---------------------------------------------------------------------------
// Slide furniture
// ---------------------------------------------------------------------------

/**
 * Master background. The original is a photo of a near-black wall that fades
 * from #3F3F3F at the top-left to #272727 at the bottom-right; a flat mid-tone
 * reproduces it without banding.
 */
function background(s) {
  s.background = { color: '333333' };
}

/** Centred "CONSTRUCTION PRESENT" header used on most slides. */
function header(s) {
  text(s, [
    run('CONSTRUCTION ', 30, C.accent2, FONT_BLACK, true),
    run('PRESENT', 30, C.accent1, FONT_BLACK, true),
  ], { x: 8.30, y: 0.337, w: 10.11, h: 0.707, align: 'center', margin: M_HEAD });
  text(s, 'Your great subtitle in this line', {
    x: 9.217, y: 0.955, w: 8.165, h: 0.545, align: 'center',
    fontSize: 18, fontFace: FONT_LIGHT, margin: M_BODY,
  });
}

/** "Description Here" + lorem intro paragraph that opens slides 2–15. */
function introBlock(s, y, w, titleX) {
  text(s, 'Description Here', {
    x: titleX, y, w: 6.754, h: 0.64, fontSize: 32, bold: true, color: C.accent5,
  });
  text(s, [
    run('Lorem ipsum dolor sit amet', 26, C.bg1, FONT_BLACK, true),
    run(LOREM_INTRO, 26, C.bg1),
  ], { x: 1.494, y: y + 0.553, w, h: 2.034, margin: M_BODY });
}

/** "DESCRIPTION  HERE" heading + body paragraph pair. */
function descBlock(s, x, y, bodyW, bodyH, color, body, titleSize = 28, bodySize = 24) {
  text(s, 'DESCRIPTION  HERE', {
    x, y, w: 4.61, h: 0.67, fontSize: titleSize, bold: true, color,
    fontFace: FONT_BLACK, margin: M_HEAD,
  });
  text(s, body, { x, y: y + 0.583, w: bodyW, h: bodyH, fontSize: bodySize, margin: M_BODY });
}

/** White glossy ring + coloured disc + white line-art glyph. */
function iconBadge(s, x, y, size, discColor, glyph) {
  oval(s, x, y, size, size, 'F1F1F2');
  const d = size * 0.554;
  const dx = x + (size - d) / 2;
  const dy = y + (size - d) / 2;
  oval(s, dx, dy, d, d, discColor);
  glyphArt(s, dx + d * 0.22, dy + d * 0.22, d * 0.56, glyph);
}

/** The four white outline glyphs used throughout the deck. */
function glyphArt(s, x, y, sz, kind) {
  const W = C.bg1;
  const bar = (bx, by, bw, bh) => rect(s, bx, by, bw, bh, W);
  const stroke = { fill: NOLINE, line: { color: W, width: Math.max(1, sz * 1.6) } };
  if (kind === 'calendar') {
    s.addShape(SH.rect, { x, y: y + sz * 0.1, w: sz, h: sz * 0.85, ...stroke });
    bar(x, y + sz * 0.1, sz, sz * 0.16);
    for (let r = 0; r < 2; r++) {
      for (let c = 0; c < 3; c++) bar(x + sz * (0.15 + c * 0.28), y + sz * (0.44 + r * 0.22), sz * 0.15, sz * 0.13);
    }
  } else if (kind === 'hourglass') {
    s.addShape(SH.custGeom, {
      x, y, w: sz, h: sz, ...stroke,
      points: [
        { x: 0, y: 0 }, { x: sz, y: 0 }, { x: sz * 0.5, y: sz * 0.5 },
        { x: sz, y: sz }, { x: 0, y: sz }, { x: sz * 0.5, y: sz * 0.5 }, { close: true },
      ],
    });
    bar(x - sz * 0.08, y - sz * 0.04, sz * 1.16, sz * 0.09);
    bar(x - sz * 0.08, y + sz * 0.95, sz * 1.16, sz * 0.09);
  } else if (kind === 'flask') {
    s.addShape(SH.rect, { x: x + sz * 0.06, y, w: sz * 0.26, h: sz, ...stroke });
    s.addShape(SH.rect, { x: x + sz * 0.68, y, w: sz * 0.26, h: sz, ...stroke });
    bar(x + sz * 0.32, y, sz * 0.36, sz * 0.09);
    bar(x + sz * 0.44, y, sz * 0.12, sz * 0.62);
    bar(x + sz * 0.32, y + sz * 0.53, sz * 0.36, sz * 0.09);
  } else { // sitemap
    const cr = sz * 0.24;
    const ring = (cx, cy) => s.addShape(SH.ellipse, { x: cx - cr / 2, y: cy - cr / 2, w: cr, h: cr, ...stroke });
    ring(x + sz / 2, y + cr / 2);
    bar(x + sz * 0.48, y + cr, sz * 0.04, sz * 0.32);
    bar(x + sz * 0.16, y + sz * 0.44, sz * 0.68, sz * 0.04);
    [0.16, 0.5, 0.84].forEach((f) => {
      bar(x + sz * f - sz * 0.02, y + sz * 0.44, sz * 0.04, sz * 0.2);
      ring(x + sz * f, y + sz * 0.76);
    });
  }
}

/** Yellow radiation warning sign: black trefoil on a yellow disc. */
function radiationSign(s, x, y, size) {
  oval(s, x, y, size, size, '2B2C2A');
  oval(s, x + size * 0.056, y + size * 0.056, size * 0.888, size * 0.888, 'FFCF13');
  // Three 60°-wide blades, 120° apart, with a gap around the central hub.
  [210, 330, 90].forEach((start) => {
    s.addShape(SH.pie, {
      x: x + size * 0.11, y: y + size * 0.11, w: size * 0.78, h: size * 0.78,
      angleRange: [start, start + 60], fill: { color: '2B2C2A' }, line: NOLINE,
    });
  });
  oval(s, x + size * 0.30, y + size * 0.30, size * 0.40, size * 0.40, 'FFCF13');
  oval(s, x + size * 0.39, y + size * 0.39, size * 0.22, size * 0.22, '2B2C2A');
}

/** Coloured disc with a white tick. */
function checkBullet(s, x, y, size, color) {
  oval(s, x, y, size, size, color);
  const lw = size * 3.4;
  s.addShape(SH.line, {
    x: x + size * 0.24, y: y + size * 0.48, w: size * 0.19, h: size * 0.22,
    line: { color: C.bg1, width: lw },
  });
  s.addShape(SH.line, {
    x: x + size * 0.43, y: y + size * 0.3, w: size * 0.33, h: size * 0.4,
    line: { color: C.bg1, width: lw }, flipV: true,
  });
}

// ---------------------------------------------------------------------------
// Flat-vector illustrations
//
// Every entry is [x, y, w, h, fill, shape?, rotate?] expressed in the
// coordinates of the slide the artwork first appears on (see SRC below).
// `shape` is one of: '' rectangle, 'o' ellipse, 'r' rounded rectangle,
// 't' triangle, 'z' trapezoid (narrow top), 'u' trapezoid (narrow bottom),
// 'd' dome (flat-bottomed half ellipse).
// `illus()` re-maps a whole table onto any target box.
// ---------------------------------------------------------------------------
const ILLUS = {
  excavator: [
    // crawler tracks
    [11.22, 11.72, 8.74, 1.46, '42383D', 'r'], [12.91, 12.19, 6.38, 0.47, '525252', 'r'],
    // body, engine cover and cab
    [10.88, 8.30, 7.99, 3.20, 'E37A25'], [10.88, 8.30, 7.99, 0.60, 'F8951D'],
    [11.35, 9.60, 3.60, 0.30, '993621'], [15.30, 7.60, 2.70, 2.80, 'E37A25'],
    [15.60, 7.90, 2.10, 1.90, 'C6E7F7'], [15.60, 7.90, 1.00, 1.90, '9BD8F2'],
    [14.55, 6.95, 1.10, 0.75, '9BD8F2'],
    // boom, stick and bucket
    [15.95, 4.35, 7.10, 1.10, 'F8951D', '', -30], [21.30, 4.20, 4.60, 0.95, 'E37A25', '', 62],
    [17.60, 7.05, 1.60, 1.60, 'FAB01F', 'o'], [22.55, 2.85, 0.55, 0.55, '42383D', 'o'],
    [23.55, 6.35, 1.80, 1.55, 'FAB01F', 'z'], [23.35, 6.20, 2.20, 0.35, 'E37A25'],
    // counter-jib strut over the cab
    [16.60, 3.05, 6.30, 0.32, '42383D', '', -28], [19.30, 1.85, 1.70, 1.35, '848484', '', -28],
  ],
  factory: [
    [14.59, 2.79, 2.03, 3.33, 'C2C2C5'], [8.97, 4.40, 3.63, 1.68, 'B0AFB2'], [12.59, 3.60, 1.78, 2.48, 'C2C2C5'],
    [16.80, 3.75, 1.45, 2.38, 'C2C2C5'], [12.23, 1.81, 1.23, 2.67, 'C2C2C5'], [14.90, 3.75, 1.97, 0.29, '919396'],
    [14.90, 4.28, 1.97, 0.29, '919396'], [14.93, 4.82, 1.97, 0.29, '919396'], [14.90, 3.21, 1.95, 0.29, '919396'],
    [11.95, 2.06, 1.78, 0.31, '636466'], [11.95, 2.62, 1.78, 0.29, '636466'], [12.87, 4.32, 1.74, 0.21, '919396'],
    [12.87, 4.71, 1.74, 0.21, '919396'], [12.89, 5.11, 1.74, 0.21, '919396'], [12.87, 3.91, 1.71, 0.21, '919396'],
    [17.02, 4.44, 1.41, 0.21, '919396'],
  ],
  laptop: [
    [3.63, 8.41, 7.21, 5.22, 'E9ECED'], [4.16, 8.89, 6.13, 2.82, 'BBC0C4'], [2.75, 7.86, 8.75, 0.55, 'BBC0C4'],
    [6.10, 11.86, 2.24, 1.54, 'BBC0C4'], [4.81, 8.39, 4.88, 0.21, '41414A'], [6.12, 11.08, 1.76, 0.45, '313F52'],
    [8.89, 10.51, 1.23, 0.45, '313F52'], [4.41, 10.51, 1.01, 0.45, '313F52'], [9.29, 9.97, 0.83, 0.46, '313F52'],
    [4.41, 9.97, 0.73, 0.46, '313F52'], [5.49, 11.08, 0.55, 0.45, '313F52'], [7.96, 11.08, 0.55, 0.45, '313F52'],
    [4.41, 9.42, 0.55, 0.45, '313F52'], [6.42, 9.97, 0.45, 0.46, '313F52'],
  ],
  envelope: [
    [4.05, 4.64, 5.00, 7.51, 'E9D096'], [4.05, 4.64, 5.00, 2.00, 'EFDDB4'], [4.08, 10.79, 4.94, 1.36, 'EFDDB4'],
    [4.28, 6.03, 4.52, 1.00, 'D9B77B'], [6.30, 5.70, 0.75, 0.78, 'FFFFFF', 'o'],
    [6.47, 5.86, 0.42, 0.45, 'AE7A5D', 'o'],
  ],
  dumptruck: [
    // tipper body (angled front) and its darker ribs
    [11.81, 6.69, 11.71, 3.14, 'FAA21B'], [11.81, 6.69, 1.60, 3.14, 'E18726', 'z35'],
    [13.40, 7.10, 0.42, 2.35, 'E18726'], [14.35, 7.10, 0.42, 2.35, 'E18726'],
    [15.30, 7.10, 0.42, 2.35, 'E18726'], [16.25, 7.10, 0.42, 2.35, 'E18726'],
    [17.20, 7.10, 0.42, 2.35, 'E18726'], [18.15, 7.10, 0.42, 2.35, 'E18726'],
    // cab, chassis and lights
    [20.04, 8.35, 3.21, 2.11, 'FAA21B'], [20.40, 8.59, 2.55, 0.48, 'F37121'],
    [15.17, 9.85, 8.23, 0.90, 'FAA21B'], [16.20, 9.48, 3.60, 0.55, 'EF4725'],
    // wheels: dark tyre, light rim, hub
    [12.74, 9.75, 4.16, 4.12, '543842', 'o'], [19.06, 9.75, 4.11, 4.12, '543842', 'o'],
    [13.68, 10.69, 2.28, 2.24, 'CBE7F9', 'o'], [19.98, 10.69, 2.27, 2.24, 'CBE7F9', 'o'],
    [14.30, 11.31, 1.04, 1.00, 'A3C2D3', 'o'], [20.60, 11.31, 1.03, 1.00, 'A3C2D3', 'o'],
  ],
  // Tower crane: raked jib, counterweight, glazed cab, lattice mast, feet.
  crane: [
    [4.30, 4.15, 4.60, 0.55, 'FAB01F', '', -18], [4.30, 4.55, 4.60, 0.30, 'E37A25', '', -18],
    [1.15, 5.75, 2.00, 0.80, 'F8951D', '', -22], [1.15, 6.25, 1.60, 0.42, 'E37A25', '', -22],
    [3.20, 5.35, 2.80, 0.42, 'E37A25', '', -30],
    [7.32, 3.90, 0.10, 5.90, '111111'], [7.20, 9.60, 0.34, 0.55, 'F8951D'],
    [1.38, 6.90, 2.75, 1.55, 'F8951D'], [1.52, 7.15, 0.62, 1.05, '9BD8F2'],
    [3.40, 7.15, 0.58, 1.05, '9BD8F2'], [2.36, 7.15, 0.52, 1.05, '993621'],
    [1.38, 8.15, 2.75, 0.40, 'FAB01F'],
    [2.90, 8.55, 0.24, 4.85, 'E37A25'], [3.52, 8.55, 0.24, 4.85, 'E37A25'],
    [2.98, 9.15, 0.70, 0.16, '993621'], [2.98, 10.25, 0.70, 0.16, '993621'],
    [2.98, 11.35, 0.70, 0.16, '993621'], [2.98, 12.45, 0.70, 0.16, '993621'],
    [2.86, 13.40, 1.94, 0.42, 'FAB01F'],
    [2.34, 13.75, 0.42, 1.25, 'F8951D', '', 12], [4.30, 13.75, 0.42, 1.25, 'F8951D', '', -12],
  ],
  boxes: [
    [5.97, 12.49, 3.78, 2.60, 'FFD9AA'], [6.31, 10.63, 3.06, 1.92, 'E7C69A'], [6.71, 9.25, 2.37, 1.37, 'D2B48D'],
    [9.54, 12.54, 1.00, 2.51, 'E7C69A'], [9.25, 10.63, 0.86, 1.92, 'D5B690'], [7.48, 12.51, 0.92, 1.00, 'C5CBD4'],
    [9.03, 9.25, 0.57, 1.37, 'C2A682'], [7.48, 10.63, 0.71, 0.74, 'B1AFBA'], [7.62, 9.25, 0.60, 0.57, 'C9C6D4'],
    [6.37, 14.57, 0.83, 0.31, 'F3F3F3'], [8.80, 10.74, 0.43, 0.37, 'F3F3F3'], [9.11, 12.66, 0.40, 0.37, 'F3F3F3'],
  ],
  crate: [
    [20.64, 4.27, 3.84, 2.64, 'FFD9AA'], [24.24, 4.35, 1.02, 2.56, 'E7C69A'], [22.18, 4.30, 0.93, 1.02, 'C5CBD4'],
    [21.05, 6.40, 0.84, 0.32, 'F3F3F3'], [23.84, 4.44, 0.41, 0.38, 'F3F3F3'],
  ],
  forklift: [
    // roll cage, mast and forks
    [3.60, 5.26, 0.42, 5.19, '4E4E4E'], [8.10, 5.26, 0.42, 5.19, '4E4E4E'],
    [3.60, 5.26, 4.92, 0.36, '4E4E4E'], [9.15, 4.96, 0.55, 6.20, '4E4E4E'],
    [9.75, 4.96, 0.55, 6.20, '4E4E4E'], [9.60, 7.33, 1.05, 3.48, 'F2C222'],
    // chassis, bonnet, seat and cushion
    [2.04, 8.49, 5.25, 2.67, 'F2C222', 'r'], [2.04, 9.60, 3.20, 1.56, 'D9A42D'],
    [4.14, 7.17, 1.05, 1.60, '676767'], [4.14, 8.55, 2.15, 0.52, '676767'],
    [7.05, 8.20, 0.42, 2.20, '676767'], [6.20, 8.05, 1.65, 0.42, '676767'],
    [2.35, 9.10, 1.60, 0.20, 'D9A42D'], [2.35, 9.45, 1.60, 0.20, 'D9A42D'],
    // wheels
    [7.41, 9.87, 1.86, 1.98, '393D3F', 'o'], [3.51, 10.23, 1.74, 1.62, '393D3F', 'o'],
    [7.85, 10.34, 0.98, 1.04, '9AA4AA', 'o'], [3.94, 10.65, 0.88, 0.78, '9AA4AA', 'o'],
  ],
  // Standing engineer: helmet, face, suit, tie, arms, trousers, shoes.
  engineerA: [
    [12.71, 4.55, 1.45, 2.30, 'F6CE9D', 'r'], [12.90, 4.20, 1.10, 0.55, '502E15', 'd'],
    [12.55, 3.60, 0.90, 1.05, 'F7DC54', 'dl'], [13.45, 3.60, 0.90, 1.05, 'EBBB1E', 'dr'],
    [12.42, 4.55, 2.06, 0.30, 'F2B71B', 'r'], [12.87, 6.65, 1.07, 0.60, 'E8EEEE'],
    [12.33, 7.05, 2.15, 3.38, '6C7177', 'r'], [13.24, 7.05, 0.42, 2.60, 'FFBD47'],
    [11.73, 7.25, 0.68, 2.75, '6C7177', 'r'], [14.28, 7.25, 0.68, 2.75, '6C7177', 'r'],
    [11.70, 9.85, 0.58, 0.70, 'FDD8C2', 'o'], [14.30, 9.85, 0.58, 0.70, 'FDD8C2', 'o'],
    [13.10, 8.40, 1.70, 2.10, '0070C0', '', 12],
    [12.55, 10.30, 0.82, 3.55, '353333'], [13.48, 10.30, 0.82, 3.55, '353333'],
    [12.40, 13.70, 1.00, 0.50, '3E2512', 'r'], [13.42, 13.70, 1.00, 0.50, '3E2512', 'r'],
  ],
  coolingtower: [
    // flared chimney, rim and base slab
    [9.98, 7.04, 6.66, 7.35, 'CCCCCA', 'z62'], [10.60, 7.04, 5.40, 7.35, 'F0EFEF', 'z55'],
    [12.10, 6.85, 2.42, 0.34, 'B6B7B7', 'r'], [9.30, 14.24, 8.11, 0.76, 'B6B7B7', 'r'],
    // steam cloud and the radiation sign on the shaft
    [12.40, 4.79, 3.20, 1.60, 'AEDBF6', 'o'], [13.60, 5.20, 2.00, 1.20, 'AEDBF6', 'o'],
  ],
  cone: [
    // tapering cone with two reflective white bands, on a rounded base
    [2.29, 3.83, 5.34, 9.46, 'FCB116', 'z12'], [3.46, 6.16, 2.98, 1.67, 'FFFFFF', 'z62'],
    [2.65, 9.92, 4.62, 1.64, 'FFFFFF', 'z80'], [1.13, 13.30, 7.57, 1.15, 'FFCC49', 'r'],
  ],
  rulerset: [
    [1.10, 8.57, 0.83, 5.92, 'E7E7E7'], [1.86, 8.57, 0.07, 5.92, 'D2D3D4'], [2.28, 8.99, 0.15, 4.89, 'CE3F27'],
    [2.42, 8.99, 0.12, 4.89, 'E04B2D'], [2.55, 8.99, 0.15, 4.89, 'F15D3B'], [2.28, 8.43, 0.42, 0.56, 'FFFFFF'],
    [2.40, 8.08, 0.17, 0.37, '2E323C'], [2.28, 14.02, 0.42, 0.49, '3F4450'], [3.11, 12.77, 1.10, 0.51, 'FFFFFF'],
    [3.11, 13.29, 1.10, 1.23, '48505A'],
  ],
  hardhat: [
    // ruler and pencil standing behind the helmet
    [23.03, 0.67, 0.59, 4.00, 'E7E7E7'], [23.09, 1.20, 0.30, 0.07, 'D2D3D4'],
    [23.09, 1.70, 0.30, 0.07, 'D2D3D4'], [23.09, 2.20, 0.30, 0.07, 'D2D3D4'],
    [23.09, 2.70, 0.30, 0.07, 'D2D3D4'], [23.09, 3.20, 0.30, 0.07, 'D2D3D4'],
    [22.60, 1.79, 0.22, 2.42, 'E04B2D'], [22.60, 1.55, 0.22, 0.28, 'FFFFFF', 't'],
    [21.60, 2.20, 0.72, 1.10, 'FCB116'], [21.60, 2.05, 0.72, 0.28, 'FFFFFF'],
    // dome shell, ribs and brim
    [20.93, 2.62, 1.80, 2.30, 'F7DA55', 'dl'], [22.73, 2.62, 1.80, 2.30, 'F7BF1C', 'dr'],
    [22.30, 3.05, 0.90, 1.87, 'F99B23'], [21.65, 3.28, 0.34, 1.14, 'F99B23', 't'],
    [23.42, 3.25, 0.36, 1.16, 'F99B23', 't'], [20.88, 4.65, 3.71, 0.61, 'F99B23', 'r'],
  ],
  // Engineer standing in front of a blue blueprint board with a yellow plan.
  blueprint: [
    [11.93, 8.38, 8.92, 6.00, '477ABD'], [11.24, 14.35, 9.91, 0.69, '4678BC'],
    [12.60, 8.90, 7.60, 0.06, '6E9BD0'], [12.60, 10.10, 7.60, 0.06, '6E9BD0'],
    [12.60, 11.30, 7.60, 0.06, '6E9BD0'], [12.60, 12.50, 7.60, 0.06, '6E9BD0'],
    [13.60, 8.60, 0.06, 5.20, '6E9BD0'], [15.60, 8.60, 0.06, 5.20, '6E9BD0'],
    [17.60, 8.60, 0.06, 5.20, '6E9BD0'], [19.60, 8.60, 0.06, 5.20, '6E9BD0'],
    [14.86, 10.25, 6.46, 4.33, 'FCC454'], [14.36, 14.58, 7.18, 0.49, 'DC9A27'],
    [15.60, 10.90, 1.60, 1.30, 'FDB71D'], [17.60, 10.90, 1.60, 1.30, 'FDB71D'],
    [15.60, 12.60, 1.60, 1.30, 'FDB71D'], [17.60, 12.60, 1.60, 1.30, 'FDB71D'],
    [11.97, 10.94, 0.60, 4.10, 'E7E7E7'],
    [19.63, 4.95, 2.69, 3.95, 'F6CE9D', 'r'], [20.00, 4.35, 1.95, 1.05, '502E15', 'd'],
    [19.30, 3.35, 1.70, 1.85, 'F7DC54', 'dl'], [21.00, 3.35, 1.75, 1.85, 'EBBB1E', 'dr'],
    [19.12, 4.95, 3.66, 0.55, 'F2B71B', 'r'], [19.92, 8.60, 1.98, 1.10, 'E8EEEE'],
    [18.92, 9.17, 3.99, 5.85, '6C7177', 'r'], [20.60, 9.17, 0.72, 4.80, 'FFBD47'],
    [17.77, 9.60, 1.35, 5.45, '6C7177', 'r'], [22.70, 9.60, 1.35, 5.45, '6C7177', 'r'],
    [21.85, 12.40, 1.35, 1.05, 'FDD8C2', 'o'],
  ],
  // Same engineer drawn larger (torso runs off the bottom of the slide).
  engineerB: [
    [5.75, 5.10, 2.10, 3.75, 'F6CE9D', 'r'], [6.00, 4.55, 1.60, 0.95, '502E15', 'd'],
    [5.45, 3.75, 1.35, 1.65, 'F7DC54', 'dl'], [6.80, 3.75, 1.35, 1.65, 'EBBB1E', 'dr'],
    [5.30, 5.05, 3.00, 0.48, 'F2B71B', 'r'], [6.05, 8.55, 1.55, 1.00, 'E8EEEE'],
    [4.84, 9.20, 3.76, 5.80, '6C7177', 'r'], [6.50, 9.20, 0.62, 4.60, 'FFBD47'],
    [3.75, 9.60, 1.25, 5.40, '6C7177', 'r'], [8.45, 9.60, 1.25, 5.40, '6C7177', 'r'],
    [3.60, 14.55, 1.20, 0.95, 'FDD8C2', 'o'], [8.40, 14.55, 1.20, 0.95, 'FDD8C2', 'o'],
  ],
  // Screwdriver seen head-on: grey shaft with a chunky yellow handle.
  screwdriver: [
    [13.01, 5.66, 0.67, 4.10, 'B2B2AF'], [12.09, 4.02, 2.77, 1.89, 'B2B2AF', 'r'],
    [11.73, 3.82, 0.36, 1.49, 'C7C6C6', 'r'], [12.86, 5.56, 0.97, 0.20, 'C7C6C6'],
    [12.60, 9.19, 1.43, 0.41, 'FCB116', 'r'], [12.60, 9.60, 1.43, 5.12, 'FFC72F', 'r'],
    [12.76, 10.27, 0.20, 3.53, 'E5A724'], [13.22, 10.32, 0.20, 3.53, 'E5A724'],
    [13.68, 10.32, 0.20, 3.53, 'E5A724'],
  ],
  // Wind turbine: tapering tower, three blades 120° apart, hub.
  windmill: [
    [13.52, 7.43, 0.77, 7.07, 'C9C8C7', 'z55'], [13.37, 14.03, 1.07, 0.89, 'A5A2A1', 'z70'],
    [13.06, 14.65, 1.69, 0.35, 'AEACAB', 'r'],
    [13.85, 4.55, 0.28, 3.00, 'AEACAB', '', 10],
    [13.95, 7.60, 0.28, 3.00, 'AEACAB', '', -130],
    [11.05, 7.35, 3.00, 0.28, 'AEACAB', '', -20],
    [13.68, 7.35, 0.46, 0.42, 'A5A2A1', 'o'], [13.75, 7.39, 0.31, 0.31, 'EDEDED', 'o'],
  ],
  // Smoke stack with a stylised flame curling above it.
  chimney: [
    [1.32, 10.53, 1.87, 4.47, '919396', 'z55'],
    [2.02, 9.55, 1.26, 0.28, 'F37340', 'r', -20],
    [2.15, 9.10, 1.00, 0.26, 'F37340', 'r', -20],
    [2.30, 8.70, 0.80, 0.24, 'F37340', 'r', -20],
  ],
};

// Source bounding box each ILLUS table was captured in.
const SRC = {
  excavator: [10.88, 1.67, 14.23, 11.65], factory: [8.97, 1.81, 9.48, 4.33],
  laptop: [2.75, 7.86, 8.75, 5.77], envelope: [4.05, 4.64, 5.00, 7.51],
  dumptruck: [11.81, 6.69, 11.71, 7.18], crane: [1.13, 3.66, 7.92, 11.33],
  boxes: [5.97, 9.25, 4.57, 5.83], crate: [20.64, 4.27, 4.63, 2.64],
  forklift: [2.04, 4.96, 10.02, 6.90], engineerA: [11.72, 3.82, 3.83, 10.60],
  coolingtower: [9.30, 4.79, 8.11, 10.21], cone: [1.13, 3.83, 7.57, 10.61],
  rulerset: [1.10, 8.08, 3.11, 6.43], hardhat: [20.88, 0.67, 3.72, 4.60],
  blueprint: [11.24, 2.80, 13.64, 12.37], engineerB: [3.75, 3.62, 6.71, 11.79],
  screwdriver: [11.73, 3.82, 3.12, 10.91], windmill: [10.87, 4.55, 5.22, 10.45],
  chimney: [1.32, 8.83, 1.97, 6.17],
};

/** Symmetric trapezoid whose narrow edge is `narrow` x the wide one. */
function taperPoints(w, h, narrow, narrowOnTop) {
  const inset = (w * (1 - narrow)) / 2;
  return narrowOnTop
    ? [{ x: inset, y: 0 }, { x: w - inset, y: 0 }, { x: w, y: h }, { x: 0, y: h }, { close: true }]
    : [{ x: 0, y: 0 }, { x: w, y: 0 }, { x: w - inset, y: h }, { x: inset, y: h }, { close: true }];
}

/**
 * Shape-code -> [preset, extra options].
 *  'zNN' / 'uNN' are trapezoids narrowing to NN% at the top / bottom.
 *  The half-dome codes use `pie`, whose 180°–270° and 270°–360° wedges fill
 *  only one quadrant of their box, so they grow and shift it to compensate.
 */
function primitive(code, geo) {
  const taper = /^([zu])(\d*)$/.exec(code || '');
  if (taper) {
    const narrow = (Number(taper[2]) || 50) / 100;
    return [SH.custGeom, { points: taperPoints(geo.w, geo.h, narrow, taper[1] === 'z') }];
  }
  switch (code) {
    case 'o': return [SH.ellipse, {}];
    case 'r': return [SH.roundRect, { rectRadius: Math.min(0.25, geo.h / 2) }];
    case 't': return [SH.triangle, {}];
    case 'd': return [SH.chord, { angleRange: [180, 360] }]; // flat-bottom dome
    case 'dl': return [SH.pie, { w: geo.w * 2, h: geo.h * 2, angleRange: [180, 270] }];
    case 'dr': return [SH.pie, { x: geo.x - geo.w, w: geo.w * 2, h: geo.h * 2, angleRange: [270, 360] }];
    default: return [SH.rect, {}];
  }
}

/**
 * Draw an illustration. With no `box` it lands at its captured position;
 * otherwise it is scaled into `box` = [x, y, w, h].
 */
function illus(s, key, box) {
  const [sx, sy, sw, sh] = SRC[key];
  const [tx, ty, tw, th] = box || SRC[key];
  const kx = tw / sw;
  const ky = th / sh;
  ILLUS[key].forEach(([x, y, w, h, fill, code, rotate]) => {
    const geo = { x: tx + (x - sx) * kx, y: ty + (y - sy) * ky, w: w * kx, h: h * ky };
    const [shape, extra] = primitive(code, geo);
    s.addShape(shape, {
      ...geo, fill: { color: fill }, line: NOLINE,
      ...(rotate ? { rotate } : {}), ...extra,
    });
  });
}

// ===========================================================================
// Slides
// ===========================================================================

// Slide 1 — cover
function slide01(pptx) {
  const s = pptx.addSlide();
  background(s);
  illus(s, 'excavator');

  text(s, 'BUILDINGZ', { x: 0.884, y: 1.167, w: 13.102, h: 2.625, fontSize: 150, bold: true, color: C.accent2, fontFace: FONT_BLACK });
  text(s, 'PRESENT', { x: 1.027, y: 3.296, w: 6.44, h: 1.784, fontSize: 100, bold: true, color: C.accent6, fontFace: FONT_BLACK });
  text(s, 'CLEAN PRESENTATIONs', { x: 1.111, y: 4.971, w: 8.356, h: 0.909, fontSize: 48, bold: true, color: C.accent3, fontFace: FONT_BLACK });
  text(s, 'Multipurpose PresentationTemplate.', {
    x: 0.916, y: 6.022, w: 6.958, h: 0.794, fontSize: 24, valign: 'middle',
    lineSpacingMultiple: 1.3, margin: [19.2, 19.2, 9.6, 9.6],
  });
  text(s, LOREM_LONG, {
    x: 1.006, y: 7.185, w: 6.869, h: 3.945, fontSize: 24, valign: 'middle',
    lineSpacingMultiple: 1.3, margin: [19.2, 19.2, 9.6, 9.6],
  });
  text(s, 'BUILDING', { x: 1.05, y: 11.385, w: 6.112, h: 1.531, fontSize: 85, bold: true, color: C.accent4, fontFace: FONT_BLACK });
  oval(s, 7.217, 11.211, 1.833, 1.833, C.accent1);
  text(s, '1', { x: 7.634, y: 11.129, w: 1.25, h: 1.784, fontSize: 100, bold: true, fontFace: FONT_BLACK, align: 'center' });
}

// Slide 2 — factory, four vertical ribbons, four data callouts
function slide02(pptx) {
  const s = pptx.addSlide();
  background(s);

  // Four coloured ribbons hanging below the ground slab, with rotated labels.
  const RIBBONS = [
    { x: 9.05, color: C.accent2 }, { x: 11.37, color: C.accent5 },
    { x: 13.61, color: C.accent6 }, { x: 15.93, color: C.accent1 },
  ];
  RIBBONS.forEach(({ x, color }, i) => {
    const dx = i * 2.245;
    rect(s, x, 6.86, 2.28, 6.55, color);
    text(s, 'DESCRIPTION', {
      x: 7.465 + dx, y: 10.351, w: 4.686, h: 0.818, rotate: 270,
      fontSize: 35, bold: true, fontFace: FONT_BLACK, margin: M_HEAD,
    });
    text(s, 'Lorem Ipsum has two main data statisticat methodologies', {
      x: 7.734 + dx, y: 9.833, w: 5.582, h: 1.05, rotate: 270, fontSize: 24, margin: M_BODY,
    });
  });

  // Factory and the radiation warning sign
  illus(s, 'factory');
  radiationSign(s, 21.50, 1.88, 3.42);

  // Ground slab with soil below
  rect(s, 6.94, 6.386, 13.171, 1.181, '8F441F');
  rect(s, 6.134, 6.083, 14.75, 0.78, 'C2C2C5');

  introBlock(s, 2.413, 6.89, 1.579);

  const CALLOUTS = [
    { bx: 1.876, by: 7.484, tx: 5.080, ty: 7.924, color: C.accent6, glyph: 'hourglass', tw: 3.082 },
    { bx: 1.876, by: 11.688, tx: 5.058, ty: 12.105, color: C.accent5, glyph: 'flask', tw: 3.082 },
    { bx: 19.325, by: 7.236, tx: 22.376, ty: 7.741, color: C.accent1, glyph: 'calendar', tw: 3.393 },
    { bx: 19.339, by: 11.665, tx: 22.353, ty: 12.007, color: C.accent3, glyph: 'sitemap', tw: 3.393 },
  ];
  CALLOUTS.forEach((c) => {
    iconBadge(s, c.bx, c.by, 3.001, c.color, c.glyph);
    text(s, 'Data Here', { x: c.tx, y: c.ty, w: 2.782, h: 0.774, fontSize: 40, bold: true, color: c.color });
    text(s, 'Lorem Ipsum has two main data statistical important.', {
      x: c.tx - 0.112, y: c.ty + 0.77, w: c.tw, h: 1.858, fontSize: 24,
    });
  });

  header(s);
}

// Slide 3 — laptop + 3D percent-stacked bar chart
function slide03(pptx) {
  const s = pptx.addSlide();
  background(s);
  illus(s, 'envelope');
  illus(s, 'laptop');

  s.addChart(pptx.ChartType.bar3d, series([210, 120, 150], [470, 430, 400], [230, 125, 420], [100, 215, 300]), {
    ...CHART_BASE,
    x: 11.5, y: 3.447, w: 14.041, h: 7.261,
    barDir: 'bar', barGrouping: 'percentStacked', bar3DShape: 'box',
    valAxisLabelFormatCode: '0%',
    chartColors: [C.accent5, C.accent2, C.accent6, C.accent1],
  });

  introBlock(s, 1.833, 22.64, 1.689);

  const DATA = [
    { x: 12.757, y: 10.804, color: C.accent5 }, { x: 12.717, y: 12.825, color: C.accent6 },
    { x: 18.292, y: 10.804, color: C.accent2 }, { x: 18.252, y: 12.825, color: C.accent1 },
  ];
  DATA.forEach((d) => {
    text(s, 'Data Here', { x: d.x, y: d.y, w: 2.197, h: 0.522, fontSize: 25, bold: true, color: d.color });
    rect(s, d.x - 0.13, d.y + 0.15, 0.1, 0.1, d.color);
    text(s, BODY_4, { x: d.x + 0.318, y: d.y + 0.416, w: 4.857, h: 1.05, fontSize: 23, margin: M_BODY });
  });
  header(s);
}

// Slide 4 — dump truck + four full-bleed information ribbons
function slide04(pptx) {
  const s = pptx.addSlide();
  background(s);
  introBlock(s, 1.833, 22.64, 1.689);
  illus(s, 'dumptruck');

  const ROWS = [
    { y: 4.75, ty: 4.94, by: 4.58, color: C.accent6, glyph: 'hourglass' },
    { y: 6.83, ty: 7.03, by: 6.72, color: C.accent5, glyph: 'flask' },
    { y: 8.93, ty: 9.13, by: 8.83, color: C.accent2, glyph: 'sitemap' },
    { y: 11.05, ty: 11.24, by: 11.00, color: C.accent1, glyph: 'calendar' },
  ];
  ROWS.forEach((r) => {
    ribbon(s, -0.03, r.y, 9.38, 2.12, r.color);
    text(s, 'Information', { x: 0.15, y: r.ty, w: 4.58, h: 0.87, fontSize: 33, bold: true, margin: M_HEAD });
    text(s, BODY_2, { x: 0.15, y: r.ty + 0.6, w: 7.26, h: 1.2, fontSize: 24, margin: M_BODY });
    iconBadge(s, 8.20, r.by, 2.53, r.color, r.glyph);
  });

  text(s, '94%', { x: 12.44, y: 4.58, w: 3.24, h: 1.7, fontSize: 95, bold: true, color: C.accent1, fontFace: FONT_BLACK });
  accentBar(s, 15.46, 4.86, 0.13, 1.12, C.accent1);
  text(s, BODY_3, { x: 15.68, y: 4.83, w: 6.37, h: 1.05, fontSize: 24, margin: M_BODY });
  header(s);
}

// Slide 5 — crane and boxes, three chevrons, three bucket stats
function slide05(pptx) {
  const s = pptx.addSlide();
  background(s);
  introBlock(s, 1.833, 22.64, 1.395);
  illus(s, 'crane');
  illus(s, 'boxes');

  const ROWS = [
    { y: 6.61, cy: 6.25, py: 7.84, ty: 7.00, pct: '73%', color: C.accent2 },
    { y: 9.11, cy: 8.84, py: 10.42, ty: 9.44, pct: '62%', color: C.accent1 },
    { y: 11.52, cy: 11.25, py: 12.84, ty: 11.93, pct: '78%', color: C.accent6 },
  ];
  ROWS.forEach((r) => {
    chevron(s, 11.28, r.y, 10.73, 2.53, r.color, true);
    descBlock(s, 12.88, r.ty, 7.34, 1.05, C.bg1, BODY_3 + ' important', 30);
    // Truncated-cone bucket: elliptical rim plus a tapering body.
    s.addShape(SH.trapezoid, {
      x: 20.44, y: r.cy + 0.63, w: 5.46, h: 2.60, flipV: true,
      fill: { color: r.color }, line: NOLINE,
    });
    oval(s, 20.44, r.cy, 5.46, 1.26, r.color);
    text(s, r.pct, { x: 21.55, y: r.py, w: 3.65, h: 1.19, fontSize: 59, bold: true, fontFace: FONT_BLACK, align: 'center', margin: M_HEAD });
  });

  illus(s, 'crate');
  text(s, [run('TOTAL ', 70, C.accent3, FONT_BLACK, true), run('DATA', 70, C.accent5, FONT_BLACK, true)],
    { x: 12.37, y: 3.60, w: 8.09, h: 1.28 });
  descBlock(s, 12.48, 4.70, 7.54, 1.05, C.accent1, BODY_1);
  header(s);
}

// Slide 6 — forklift, three bands, TOTAL 79%, three check bullets
function slide06(pptx) {
  const s = pptx.addSlide();
  background(s);
  introBlock(s, 1.833, 22.64, 1.689);

  rect(s, 2.76, 11.99, 13.171, 1.18, '8F441F');
  rect(s, 1.96, 11.68, 14.75, 0.78, 'C2C2C5');

  const BANDS = [
    { y: 5.38, h: 1.81, ty: 5.41, color: C.accent5 },
    { y: 7.13, h: 1.82, ty: 7.26, color: C.accent1 },
    { y: 8.94, h: 1.82, ty: 9.08, color: C.accent2 },
  ];
  BANDS.forEach((b) => rect(s, 9.58, b.y, 8.97, b.h, b.color));
  illus(s, 'forklift');
  BANDS.forEach((b) => descBlock(s, 11.91, b.ty, 5.26, 1.05, C.bg1, BODY_3 + ' ', 30));

  text(s, [run('TOTAL ', 70, C.accent1, FONT_BLACK, true), run('79%', 70, C.accent5, FONT_BLACK, true)],
    { x: 19.04, y: 3.92, w: 6.59, h: 1.28 });
  descBlock(s, 19.14, 5.21, 5.96, 1.45, C.accent1, BODY_1);

  const CHECKS = [
    { y: 8.60, ty: 8.27, color: C.accent2 },
    { y: 10.06, ty: 9.73, color: C.accent1 },
    { y: 11.61, ty: 11.28, color: C.accent5 },
  ];
  CHECKS.forEach((c) => {
    checkBullet(s, 19.40, c.y, 0.89, c.color);
    text(s, BODY_4, { x: 20.46, y: c.ty, w: 3.93, h: 1.05, fontSize: 24, margin: M_BODY });
  });
  header(s);
}

/** The "Information" caption pair that sits on each ribbon of slides 7/8/9. */
function infoCaption(s, x, y) {
  text(s, 'Information', { x, y, w: 4.58, h: 0.85, fontSize: 32, bold: true, margin: M_HEAD });
  text(s, 'Lorem Ipsum has two main data statistical this methodologies.', {
    x, y: y + 0.6, w: 6.13, h: 1.05, fontSize: 24, margin: M_BODY,
  });
}

// Slide 7 — engineer surrounded by four ribbons and four check bullets
function slide07(pptx) {
  const s = pptx.addSlide();
  background(s);
  introBlock(s, 1.833, 22.64, 1.689);

  const ROWS = [
    { x: 13.47, y: 5.64, w: 8.74, tx: 15.89, ty: 5.86, bx: 21.31, by: 5.58, color: C.accent1, glyph: 'calendar', flip: false },
    { x: 13.65, y: 8.27, w: 8.74, tx: 15.97, ty: 8.50, bx: 21.37, by: 8.31, color: C.accent6, glyph: 'hourglass', flip: false },
    { x: 4.10, y: 6.84, w: 9.35, tx: 5.42, ty: 7.00, bx: 2.55, by: 6.92, color: C.accent5, glyph: 'flask', flip: true },
    { x: 3.78, y: 10.55, w: 9.35, tx: 5.50, ty: 10.73, bx: 2.48, by: 10.58, color: C.accent2, glyph: 'sitemap', flip: true },
  ];
  ROWS.forEach((b) => ribbon(s, b.x, b.y, b.w, 2.12, b.color, b.flip));
  illus(s, 'engineerA');
  ROWS.forEach((b) => {
    infoCaption(s, b.tx, b.ty);
    iconBadge(s, b.bx, b.by, 2.53, b.color, b.glyph);
  });

  const CHECKS = [
    { cx: 1.88, cy: 4.31, tx: 2.89, ty: 3.99, w: 6.28, color: C.accent2 },
    { cx: 1.86, cy: 5.65, tx: 2.88, ty: 5.33, w: 5.40, color: C.accent6 },
    { cx: 17.17, cy: 11.48, tx: 18.18, ty: 11.17, w: 5.12, color: C.accent5 },
    { cx: 17.21, cy: 12.82, tx: 18.22, ty: 12.50, w: 4.67, color: C.accent1 },
  ];
  CHECKS.forEach((c) => {
    checkBullet(s, c.cx, c.cy, 0.85, c.color);
    text(s, BODY_4, { x: c.tx, y: c.ty, w: c.w, h: 1.05, fontSize: 24, margin: M_BODY });
  });
  header(s);
}

// Slide 8 — cooling tower, POLLUTION 67%, four ribbons
function slide08(pptx) {
  const s = pptx.addSlide();
  background(s);
  introBlock(s, 1.833, 22.64, 1.689);

  const ROWS = [
    { x: 13.47, y: 7.09, w: 8.74, tx: 15.89, ty: 7.31, bx: 21.31, by: 7.03, color: C.accent1, glyph: 'calendar', flip: false },
    { x: 13.65, y: 9.72, w: 8.74, tx: 15.97, ty: 9.95, bx: 21.37, by: 9.76, color: C.accent6, glyph: 'hourglass', flip: false },
    { x: 4.10, y: 8.29, w: 9.35, tx: 5.05, ty: 8.45, bx: 2.55, by: 8.37, color: C.accent5, glyph: 'flask', flip: true },
    { x: 3.78, y: 12.00, w: 9.35, tx: 5.13, ty: 12.18, bx: 2.48, by: 12.03, color: C.accent2, glyph: 'sitemap', flip: true },
  ];
  ROWS.forEach((b) => {
    ribbon(s, b.x, b.y, b.w, 2.12, b.color, b.flip);
    infoCaption(s, b.tx, b.ty);
    iconBadge(s, b.bx, b.by, 2.53, b.color, b.glyph);
  });
  illus(s, 'coolingtower');
  radiationSign(s, 11.95, 9.40, 2.70);

  text(s, [
    run('POLLUTION', 70, C.accent4, FONT_BLACK, true),
    run(' ', 70, C.accent1, FONT_BLACK, true),
    run('67%', 70, C.accent5, FONT_BLACK, true),
  ], { x: 1.74, y: 4.33, w: 8.24, h: 1.28 });
  descBlock(s, 1.72, 5.63, 7.40, 1.05, C.accent1, BODY_1);
  header(s);
}

// Slide 9 — traffic cone, four ribbons, four percentages
function slide09(pptx) {
  const s = pptx.addSlide();
  background(s);
  introBlock(s, 1.833, 22.64, 1.689);

  const ROWS = [
    { y: 4.86, ty: 4.97, by: 4.77, ay: 5.12, py: 4.96, pct: '87', color: C.accent1, glyph: 'calendar' },
    { y: 6.94, ty: 7.06, by: 6.88, ay: 7.37, py: 7.21, pct: '73', color: C.accent6, glyph: 'hourglass' },
    { y: 9.04, ty: 9.16, by: 9.02, ay: 9.67, py: 9.50, pct: '64', color: C.accent5, glyph: 'flask' },
    { y: 11.14, ty: 11.25, by: 11.13, ay: 11.92, py: 11.75, pct: '78', color: C.accent2, glyph: 'sitemap' },
  ];
  ROWS.forEach((r) => {
    ribbon(s, 5.13, r.y, 9.97, 2.12, r.color);
    text(s, 'Information', { x: 7.95, y: r.ty, w: 4.58, h: 0.87, fontSize: 33, bold: true, margin: M_HEAD });
    text(s, BODY_2, { x: 7.95, y: r.ty + 0.6, w: 6.45, h: 1.05, fontSize: 24, margin: M_BODY });
    iconBadge(s, 14.33, r.by, 2.53, r.color, r.glyph);
    text(s, r.pct + '%', { x: 16.78, y: r.py, w: 2.6, h: 1.36, fontSize: 75, bold: true, color: r.color, fontFace: FONT_BLACK });
    accentBar(s, 19.36, r.ay, 0.13, 1.12, r.color);
    text(s, BODY_3, { x: 19.59, y: r.ay - 0.03, w: 6.37, h: 1.05, fontSize: 24, margin: M_BODY });
  });

  illus(s, 'cone');
  illus(s, 'rulerset');
  header(s);
}

// Slide 10 — blue band, hard hat, ENGINEERING HERE
function slide10(pptx) {
  const s = pptx.addSlide();
  background(s);
  introBlock(s, 1.833, 22.64, 1.689);

  rect(s, 0.01, 5.45, 26.59, 9.63, '0070C0');
  illus(s, 'hardhat', [16.00, 3.25, 9.55, 11.79]);

  text(s, 'ENGINEERING HERE', {
    x: 8.00, y: 6.36, w: 7.53, h: 2.39, fontSize: 65, bold: true, color: C.accent1, margin: M_HEAD,
  });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Ut efficitur ipsum vitae tortor ' +
    'accumsan, a pulvinar lorem lacinia. Donec eu arcu justo. Fusce eget consequat risus Proin est lacus, ' +
    'interdum vitae feugiat quis, faucibus vel mi.  Vivamus accumsan',
    { x: 7.91, y: 8.96, w: 8.21, h: 2.26, fontSize: 24, margin: M_BODY });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Ut efficitur ipsum vitae tortor ' +
    'accumsan, a pulvinar lorem lacinia. Donec eu arcu justo. Fusce eget consequat risus Proin est lacus, ',
    { x: 7.91, y: 11.73, w: 8.21, h: 1.86, fontSize: 24, margin: M_BODY });

  glyphArt(s, 3.18, 6.75, 1.5, 'flask');
  text(s, 'DATA 75%', { x: 1.54, y: 8.92, w: 5.10, h: 1.28, fontSize: 70, bold: true, color: C.accent2, fontFace: FONT_BLACK, align: 'center' });
  descBlock(s, 1.49, 10.49, 5.26, 1.86, C.accent1,
    'Lorem Ipsum has two main data statistical this methodologies important Data analysis which summarizes.');
  header(s);
}

// Slide 11 — blueprint engineer, four 3D blocks, PLANNING HERE
function slide11(pptx) {
  const s = pptx.addSlide();
  background(s);
  introBlock(s, 1.833, 16.71, 1.689);
  illus(s, 'blueprint');

  const BLOCKS = [
    { x: 1.67, y: 4.17, pct: '78%', label: 'TARGETING', ty: 4.63, by: 5.18, color: C.accent6 },
    { x: 1.61, y: 6.75, pct: '54%', label: 'ADVERTISE', ty: 7.00, by: 7.55, color: C.accent5 },
    { x: 1.60, y: 9.34, pct: '87%', label: 'TARGETING', ty: 9.66, by: 10.26, color: C.accent3 },
    { x: 1.58, y: 12.03, pct: '69%', label: 'ADVERTISE', ty: 12.03, by: 12.63, color: C.accent2 },
  ];
  BLOCKS.forEach((b) => {
    s.addShape(SH.trapezoid, { x: b.x, y: b.y, w: 4.41, h: 0.52, flipV: true, fill: { color: b.color }, line: NOLINE });
    rect(s, b.x, b.y + 0.52, 4.41, 2.33, b.color);
    text(s, b.pct, { x: b.x + 0.37, y: b.y + 1.07, w: 3.65, h: 0.77, fontSize: 34, bold: true, fontFace: FONT_BLACK, align: 'center', margin: M_HEAD });
    text(s, b.label, { x: b.x, y: b.y + 1.70, w: 4.17, h: 0.71, fontSize: 30, bold: true, fontFace: FONT_BLACK, align: 'center', margin: M_HEAD });
    text(s, 'Data Here', { x: 6.55, y: b.ty, w: 3.46, h: 0.71, fontSize: 30, bold: true, color: b.color, fontFace: FONT_BLACK, margin: M_HEAD });
    text(s, BODY_3, { x: 6.57, y: b.by, w: 4.53, h: 1.45, fontSize: 24, margin: M_BODY });
  });

  accentBar(s, 12.80, 4.74, 0.23, 1.93, C.accent4);
  text(s, [run('PLANNING', 65, C.accent2, FONT, true), run(' HERE', 65, C.accent1, FONT, true)],
    { x: 13.21, y: 4.42, w: 5.72, h: 2.39, margin: M_HEAD });
  header(s);
}

// Slide 12 — engineer at left, four ribbons, four side notes
function slide12(pptx) {
  const s = pptx.addSlide();
  background(s);
  introBlock(s, 1.833, 22.64, 1.689);

  const ROWS = [
    { y: 5.18, ty: 5.29, by: 5.08, ay: 5.42, color: C.accent1, glyph: 'calendar' },
    { y: 7.26, ty: 7.38, by: 7.20, ay: 7.67, color: C.accent6, glyph: 'hourglass' },
    { y: 9.36, ty: 9.48, by: 9.33, ay: 9.97, color: C.accent5, glyph: 'flask' },
    { y: 11.46, ty: 11.57, by: 11.45, ay: 12.22, color: C.accent2, glyph: 'sitemap' },
  ];
  ROWS.forEach((r) => {
    ribbon(s, 7.31, r.y, 9.97, 2.12, r.color);
    text(s, 'Information', { x: 10.14, y: r.ty, w: 4.58, h: 0.87, fontSize: 33, bold: true, margin: M_HEAD });
    text(s, BODY_2, { x: 10.14, y: r.ty + 0.6, w: 6.45, h: 1.05, fontSize: 24, margin: M_BODY });
    iconBadge(s, 16.51, r.by, 2.53, r.color, r.glyph);
    accentBar(s, 19.22, r.ay, 0.13, 1.12, r.color);
    text(s, BODY_3, { x: 19.59, y: r.ay - 0.17, w: 4.62, h: 1.45, fontSize: 24, margin: M_BODY });
  });

  illus(s, 'engineerB');
  header(s);
}

// Slide 13 — gear wheel with four percentage bands and four callouts
function slide13(pptx) {
  const s = pptx.addSlide();
  background(s);
  introBlock(s, 1.833, 22.64, 1.689);

  s.addShape(SH.gear9, { x: 8.52, y: 3.81, w: 10.25, h: 10.19, fill: { color: C.accent3 }, line: NOLINE });
  oval(s, 10.26, 5.51, 6.88, 6.88, C.accent3); // hub the colour bands sit in

  const BANDS = [
    { x: 10.68, y: 5.54, w: 6.02, h: 1.76, ty: 5.98, pct: '59%', color: C.accent6 },
    { x: 10.26, y: 7.27, w: 6.88, h: 1.73, ty: 7.68, pct: '78%', color: C.accent5 },
    { x: 10.26, y: 9.00, w: 6.88, h: 1.76, ty: 9.31, pct: '39%', color: C.accent1 },
    { x: 10.81, y: 10.76, w: 5.78, h: 1.60, ty: 11.01, pct: '74%', color: C.accent2 },
  ];
  BANDS.forEach((b) => {
    rect(s, b.x, b.y, b.w, b.h, b.color);
    text(s, b.pct, { x: 12.98, y: b.ty, w: 1.54, h: 0.91, fontSize: 48, bold: true, align: 'center', margin: 0 });
  });

  const NOTES = [
    { x: 2.60, y: 6.36, tx: 3.03, ty: 6.55, color: C.accent6, left: true },
    { x: 16.90, y: 6.78, tx: 19.02, ty: 7.00, color: C.accent5, left: false },
    { x: 2.70, y: 10.48, tx: 3.11, ty: 10.66, color: C.accent2, left: true },
    { x: 15.83, y: 10.80, tx: 18.08, ty: 10.90, color: C.accent1, left: false },
  ];
  NOTES.forEach((n) => {
    chevron(s, n.x, n.y, 8.06, 1.81, n.color, n.left);
    const align = n.left ? 'right' : 'left';
    text(s, 'DESCRIPTION  HERE', {
      x: n.tx, y: n.ty, w: 5.53, h: 0.62, fontSize: 25, bold: true, fontFace: FONT_BLACK, align, margin: M_HEAD,
    });
    text(s, 'Lorem Ipsum has two main data methodologies .', {
      x: n.tx + (n.left ? 0.25 : 0), y: n.ty + 0.43, w: 5.30, h: 1.05, fontSize: 24, align, margin: M_BODY,
    });
  });
  header(s);
}

// Slide 14 — 3D cylinder column chart + four badge callouts
function slide14(pptx) {
  const s = pptx.addSlide();
  background(s);

  const ROWS = [
    { by: 3.89, ty: 4.33, color: C.accent6, glyph: 'hourglass' },
    { by: 6.42, ty: 6.83, color: C.accent5, glyph: 'flask' },
    { by: 9.00, ty: 9.59, color: C.accent1, glyph: 'calendar' },
    { by: 11.75, ty: 12.17, color: C.accent2, glyph: 'sitemap' },
  ];
  ROWS.forEach((r) => {
    text(s, 'Data Here', { x: 17.05, y: r.ty, w: 2.78, h: 0.77, fontSize: 40, bold: true, color: r.color });
    text(s, BODY_5, { x: 16.94, y: r.ty + 0.79, w: 7.40, h: 1.05, fontSize: 24, margin: M_BODY });
    iconBadge(s, 13.84, r.by, 3.001, r.color, r.glyph);
  });

  introBlock(s, 1.833, 22.64, 1.689);

  s.addChart(pptx.ChartType.bar3d, series([80, 120, 150], [500, 430, 400], [230, 125, 420], [100, 215, 300]), {
    ...CHART_BASE,
    x: 1.04, y: 4.30, w: 12.30, h: 11.06,
    barDir: 'col', barGrouping: 'clustered', bar3DShape: 'cylinder',
    chartColors: [C.accent1, C.accent5, C.accent6, C.accent3],
  });
  header(s);
}

// Slide 15 — screwdriver, BUILDING 87%, six lettered chevrons
function slide15(pptx) {
  const s = pptx.addSlide();
  background(s);
  introBlock(s, 1.833, 22.64, 1.689);

  const ROWS = [
    { right: true, y: 6.20, letter: 'A', lx: 20.21, tx: 14.84, ty: 6.59, color: C.accent1 },
    { right: false, y: 6.25, letter: 'B', lx: 5.79, tx: 6.85, ty: 6.55, color: C.accent6 },
    { right: true, y: 7.95, letter: 'C', lx: 20.23, tx: 14.85, ty: 8.28, color: C.accent5 },
    { right: false, y: 7.98, letter: 'D', lx: 5.68, tx: 6.75, ty: 8.43, color: C.accent2 },
    { right: true, y: 9.70, letter: 'E', lx: 20.21, tx: 14.84, ty: 10.09, color: C.accent3 },
    { right: false, y: 9.73, letter: 'F', lx: 5.69, tx: 6.75, ty: 10.18, color: C.accent4 },
  ];
  ROWS.forEach((r) => {
    chevron(s, r.right ? 13.62 : 5.06, r.y, 8.06, 1.82, r.color, !r.right);
    text(s, r.letter, { x: r.lx, y: r.y + 0.36, w: 1.15, h: 1.04, fontSize: 50, bold: true, margin: M_HEAD });
    text(s, 'Lorem Ipsum has two main data statistical this methodologies.', {
      x: r.tx, y: r.ty, w: 5.38, h: 1.08, fontSize: 24, align: r.right ? 'right' : 'left', margin: M_BODY,
    });
  });

  illus(s, 'screwdriver');
  text(s, [run('BUILDING ', 70, C.accent1, FONT_BLACK, true), run('87%', 70, C.accent5, FONT_BLACK, true)],
    { x: 1.74, y: 4.33, w: 8.24, h: 1.28 });
  descBlock(s, 1.50, 12.42, 7.40, 1.05, C.accent5, BODY_1);
  descBlock(s, 19.92, 3.83, 6.00, 1.45, C.accent1, BODY_1);
  descBlock(s, 20.08, 12.16, 6.00, 1.45, C.accent3, BODY_1);
  header(s);
}

/**
 * Slides 16–22 share one anatomy: photo placeholders plus a solid olive panel
 * carrying a headline, two description blocks and a long lorem paragraph.
 */
function panelSlide(pptx, cfg) {
  const s = pptx.addSlide();
  background(s);
  (cfg.before || []).forEach((fn) => fn(s));

  const [px, py, pw, ph] = cfg.panel;
  rect(s, px, py, pw, ph, C.tx2);

  text(s, cfg.title, { x: cfg.title_[0], y: cfg.title_[1], w: cfg.title_[2], h: cfg.title_[3] });
  descBlock(s, cfg.desc[0], cfg.desc[1], cfg.desc[2], cfg.desc[3], C.accent1, BODY_1);
  text(s, 'DESCRIPTION  HERE', {
    x: cfg.desc2[0], y: cfg.desc2[1], w: 5.82, h: 0.67,
    fontSize: 28, bold: true, color: C.accent2, fontFace: FONT_BLACK, margin: M_HEAD,
  });
  text(s, cfg.body || LOREM_LONG, {
    x: cfg.bodyAt[0], y: cfg.bodyAt[1], w: cfg.bodyAt[2], h: cfg.bodyAt[3],
    fontSize: 24, valign: 'middle', lineSpacingMultiple: 1.3, margin: [19.2, 19.2, 9.6, 9.6],
  });
  if (cfg.glyph) glyphArt(s, cfg.glyph[0], cfg.glyph[1], cfg.glyph[2], 'sitemap');

  (cfg.after || []).forEach((fn) => fn(s));
  if (cfg.header !== false) header(s);
}

/** Two-line headline (word on one line, percentage on the next). */
function stackedTitle(word, pct, pctColor) {
  return [
    run(word, 70, C.accent1, FONT_BLACK, true, { breakLine: true }),
    run(pct, 70, pctColor, FONT_BLACK, true),
  ];
}

function slide16(pptx) {
  panelSlide(pptx, {
    panel: [0.97, 0.0, 8.25, 11.76],
    title: [run('BUILDING ', 70, C.accent1, FONT_BLACK, true), run('92%', 70, C.accent5, FONT_BLACK, true)],
    title_: [1.31, 0.97, 7.74, 1.28],
    desc: [1.31, 2.58, 7.03, 1.05],
    desc2: [1.61, 6.83],
    bodyAt: [1.47, 7.50, 6.87, 3.94],
    glyph: [4.02, 4.82, 1.54],
    after: [(s) => photoBox(s, 12.91, 0.50, 13.70, 12.33)],
  });
}

function slide17(pptx) {
  panelSlide(pptx, {
    panel: [14.97, 3.24, 9.67, 11.76],
    title: [run('CONSTRUCTION ', 70, C.accent1, FONT_BLACK, true), run('97%', 70, C.accent5, FONT_BLACK, true)],
    title_: [15.38, 3.75, 9.08, 2.46],
    desc: [15.47, 6.40, 7.03, 1.05],
    desc2: [15.52, 8.17],
    bodyAt: [15.38, 8.83, 8.67, 3.42],
    after: [(s) => photoBox(s, 0.01, 2.83, 11.95, 10.75)],
  });
}

function slide18(pptx) {
  panelSlide(pptx, {
    header: false,
    panel: [14.97, 5.19, 9.67, 9.81],
    title: stackedTitle('REAL ESTATE', '59%', C.accent2),
    title_: [15.47, 6.08, 9.08, 2.46],
    desc: [15.55, 8.74, 7.03, 1.05],
    desc2: [15.61, 10.51],
    bodyAt: [15.47, 11.16, 8.67, 3.42],
    before: [(s) => photoBox(s, 0.0, 0.0, 13.31, 15.0)],
    after: [(s) => illus(s, 'hardhat')],
  });
}

function slide19(pptx) {
  panelSlide(pptx, {
    panel: [4.35, 5.27, 9.04, 9.81],
    title: stackedTitle('REAL ESTATE', '91%', C.accent2),
    title_: [4.63, 6.06, 8.67, 2.46],
    desc: [4.72, 8.71, 7.03, 1.05],
    desc2: [4.77, 10.48],
    bodyAt: [4.63, 11.14, 8.67, 3.42],
    before: [
      (s) => photoBox(s, 14.05, 3.58, 8.91, 7.51),
      (s) => illus(s, 'crane', [0.25, 1.66, 10.38, 13.34]),
    ],
    after: [(s) => descBlock(s, 13.92, 12.17, 7.40, 1.05, C.accent5, BODY_1)],
  });
}

function slide20(pptx) {
  panelSlide(pptx, {
    panel: [18.22, 5.19, 8.00, 9.81],
    title: stackedTitle('BUILDING', '59%', C.accent2),
    title_: [19.32, 5.67, 6.65, 2.46],
    desc: [19.32, 8.38, 6.26, 1.45],
    desc2: [19.38, 10.39],
    bodyAt: [19.32, 11.19, 6.56, 2.89],
    body: LOREM_MED,
    before: [
      (s) => photoBox(s, 1.46, 3.50, 5.67, 7.63),
      (s) => photoBox(s, 6.55, 3.50, 5.67, 7.63),
      (s) => photoBox(s, 11.71, 3.50, 5.67, 7.63),
    ],
    after: [(s) => illus(s, 'hardhat', [14.13, 8.36, 5.42, 6.69])],
  });
}

function slide21(pptx) {
  panelSlide(pptx, {
    header: false,
    panel: [16.47, -0.03, 8.25, 11.76],
    title: [run('ECOLOGY ', 70, C.accent1, FONT_BLACK, true), run('92%', 70, C.accent5, FONT_BLACK, true)],
    title_: [16.81, 0.94, 7.74, 1.28],
    desc: [16.81, 2.55, 7.03, 1.05],
    desc2: [17.11, 6.79],
    bodyAt: [16.97, 7.47, 6.87, 3.94],
    glyph: [19.52, 4.79, 1.54],
    before: [
      (s) => photoBox(s, 0.63, -0.03, 6.42, 5.88),
      (s) => photoBox(s, 7.55, -0.03, 6.42, 5.88),
      (s) => photoBox(s, 4.05, 0.10, 6.45, 5.83),
      (s) => photoBox(s, 4.01, 6.17, 6.42, 5.88),
    ],
    after: [
      (s) => illus(s, 'windmill'),
      (s) => illus(s, 'windmill', [8.97, 8.40, 3.29, 6.65]),
      (s) => illus(s, 'chimney'),
    ],
  });
}

function slide22(pptx) {
  panelSlide(pptx, {
    panel: [13.09, 3.02, 13.51, 7.22],
    title: [run('BUILDING ', 70, C.accent1, FONT_BLACK, true), run('59%', 70, C.accent2, FONT_BLACK, true)],
    title_: [14.33, 3.50, 8.63, 1.28],
    desc: [14.25, 4.87, 10.92, 1.05],
    desc2: [14.22, 6.58],
    bodyAt: [14.16, 7.19, 11.47, 1.84],
    body: LOREM_MED,
    before: [(s) => photoBox(s, 2.63, 2.28, 8.58, 11.55)],
    after: [
      (s) => illus(s, 'hardhat', [9.58, 7.29, 5.42, 6.69]),
      (s) => {
        // Screwdriver lying on its side, plus a ruler and an eraser.
        rect(s, 21.55, 11.45, 2.10, 0.42, 'B2B2AF');
        rect(s, 23.55, 11.33, 1.80, 0.66, 'FFC72F');
        rect(s, 23.40, 11.33, 0.25, 0.66, 'FCB116');
        rect(s, 21.35, 11.30, 0.50, 0.72, 'C7C6C6');
        rect(s, 19.87, 11.90, 5.60, 0.62, 'E7E7E7');
        for (let i = 0; i < 14; i++) rect(s, 20.05 + i * 0.38, 11.93, 0.05, 0.18, 'D2D3D4');
        rect(s, 14.02, 9.99, 1.10, 1.10, '48505A');
        rect(s, 14.02, 9.99, 1.10, 0.35, 'FFFFFF');
      },
    ],
  });
}

// Slide 23 — CHARTS HERE panel + 2D clustered column chart + four callouts
function slide23(pptx) {
  const s = pptx.addSlide();
  background(s);

  rect(s, 0.42, 0.01, 8.50, 4.02, C.tx2);
  text(s, [run('CHARTS ', 70, C.accent1, FONT_BLACK, true), run('HERE', 70, C.accent2, FONT_BLACK, true)],
    { x: 0.72, y: 0.48, w: 8.05, h: 1.28 });
  descBlock(s, 0.80, 1.87, 7.24, 1.05, C.accent1, BODY_1);
  descBlock(s, 10.17, 2.58, 13.66, 0.65, C.accent5, BODY_1);

  s.addChart(pptx.ChartType.bar, series([80, 120, 150], [500, 430, 400], [230, 125, 420], [100, 215, 300]), {
    ...CHART_BASE,
    x: 2.41, y: 5.08, w: 21.14, h: 5.51,
    barDir: 'col', barGrouping: 'clustered',
    chartColors: [C.accent5, C.accent2, C.accent1, C.accent3],
  });

  const CARDS = [
    { bx: 1.53, tx: 3.72, ty: 11.54, color: C.accent6, glyph: 'hourglass' },
    { bx: 7.30, tx: 9.85, ty: 11.57, color: C.accent5, glyph: 'flask' },
    { bx: 13.37, tx: 15.63, ty: 11.52, color: C.accent1, glyph: 'calendar' },
    { bx: 19.03, tx: 21.77, ty: 11.54, color: C.accent2, glyph: 'sitemap' },
  ];
  CARDS.forEach((c) => {
    iconBadge(s, c.bx, 11.28, 2.52, c.color, c.glyph);
    text(s, 'Data Here', { x: c.tx + 0.11, y: c.ty, w: 2.78, h: 0.77, fontSize: 40, bold: true, color: c.color });
    text(s, BODY_4, { x: c.tx, y: c.ty + 0.65, w: 3.78, h: 1.05, fontSize: 24, margin: M_BODY });
  });
  header(s);
}

// ---------------------------------------------------------------------------
// Charts
// ---------------------------------------------------------------------------
const CHART_CATS = ['Group 1', 'Group 2', 'Group 3'];

function series(...valueSets) {
  const names = ['Data 1', 'Data 2', ' Data 3', 'Data 4'];
  return valueSets.map((values, i) => ({ name: names[i], labels: CHART_CATS, values }));
}

const CHART_BASE = {
  showLegend: false,
  showTitle: false,
  catAxisLabelColor: 'BFBFBF',
  valAxisLabelColor: 'BFBFBF',
  catAxisLabelFontFace: FONT,
  valAxisLabelFontFace: FONT,
  catAxisLabelFontSize: 12,
  valAxisLabelFontSize: 12,
  catGridLine: { color: 'F2F2F2', size: 1 },
  valGridLine: { color: 'F2F2F2', size: 1 },
  catAxisLineColor: 'F2F2F2',
  valAxisLineColor: 'F2F2F2',
  barGapWidthPct: 150,
  v3DRotX: 15,
  v3DRotY: 20,
  v3DPerspective: 30,
};

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------
const pptx = new PptxGenJS();
SH = pptx.ShapeType;
pptx.defineLayout({ name: 'DECK', width: SLIDE_W, height: SLIDE_H });
pptx.layout = 'DECK';
pptx.title = 'BUILDINGZ — Multipurpose Presentation Template';

[
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
  slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
  slide17, slide18, slide19, slide20, slide21, slide22, slide23,
].forEach((build) => build(pptx));

pptx.writeFile({
  fileName: path.join(__dirname, '13e96605-dab3-4250-9f85-b06554d25c85_grok_final.pptx'),
}).then((f) => console.log('wrote', f));
