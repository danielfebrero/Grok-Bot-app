/**
 * "Music Max" presentation template - 30 slides, 13.333 x 7.5 in.
 * Rebuilt with pptxgenjs only.
 *
 * The source deck leaves every picture placeholder empty (they render blank),
 * so only the two real bitmaps - the phone and desktop mockups on slides 27/28 -
 * become grey "[image]" stand-ins here.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const INK = '0C0C0C';       // near-black headings / body
const CHARCOAL = '242424';  // dark puzzle blocks, page numbers
const SLATE = '494949';     // muted dark puzzle block (slide 12)
const BLUE = '22A7F0';      // theme accent
const SKY = '7ACAF6';       // light accent puzzle pieces
const PALE = 'D3EDFC';      // palest accent puzzle pieces
const SILVER = 'C2C2C2';    // grey puzzle pieces / muted strap-line
const GREY = '868686';      // body copy
const TRACK = 'E6E6E6';     // progress-bar track
const WHITE = 'FFFFFF';

const SERIF = 'PT Serif';   // theme major font
const SANS = 'Raleway';     // theme minor font

/* --------------------------------------------------------------- boilerplate */
const L1 = 'PLACEHOLDER';
const L2 = L1 + 'PLACEHOLDER';
const L3 = L2 + ' cursus mi ut';
const L4 = L3 + ' sapien tincidunt velas aliqueta nisa adipiscinge elit sedasisa';
const L5 = L4 + ' cursus misani utasisau nasat sapienan';
const L6 = L5 + ' tincidunt velas';
const L7 = L6 + ' contan man';
const L8 = L4 + ' asi adipiscingesa nibuhase eget justo sed';
const L_CARD = 'Lorem ipsum dolor sita amet hanisa nisi';
const L_PORT = 'Lorem ipsum dolor sita amet hanisa  onse caila teturasi adipiscingesa';
const NOTE = 'Write  Something Here';
const MUSIC_TAG = '246 . 324 : Music';

/* ------------------------------------------------- freeform puzzle outlines */
/* Normalised 0..1 outlines: M move, L line, C cubic (c1x,c1y,c2x,c2y,x,y), Z close */

/** Tall corner tab: round notch in the right edge, round tab hanging below. */
const TAB_PANEL = [
  ['M', 0, 0], ['L', 1, 0], ['L', 1, 0.1423], ['L', 0.9673, 0.1447],
  ['C', 0.8933, 0.1558, 0.8376, 0.2036, 0.8376, 0.2610],
  ['C', 0.8376, 0.3183, 0.8933, 0.3661, 0.9673, 0.3772],
  ['L', 1, 0.3796], ['L', 1, 0.8814], ['L', 0.4978, 0.8814],
  ['C', 0.4978, 0.9469, 0.4251, 1, 0.3354, 1],
  ['C', 0.2457, 1, 0.1730, 0.9469, 0.1730, 0.8814],
  ['L', 0, 0.8814], ['Z'],
];

/** Service-card block: same jigsaw language, squarer proportions. */
const CARD_PANEL = [
  ['M', 0, 0], ['L', 1, 0], ['L', 1, 0.1423], ['L', 0.9999, 0.1423],
  ['C', 0.9280, 0.1423, 0.8696, 0.1954, 0.8696, 0.2610],
  ['C', 0.8696, 0.3265, 0.9280, 0.3796, 0.9999, 0.3796],
  ['L', 1, 0.3796], ['L', 1, 0.8814], ['L', 0.4139, 0.8814],
  ['L', 0.4113, 0.9053],
  ['C', 0.3991, 0.9593, 0.3465, 1, 0.2835, 1],
  ['C', 0.2205, 1, 0.1679, 0.9593, 0.1558, 0.9053],
  ['L', 0.1531, 0.8814], ['L', 0, 0.8814], ['Z'],
];

/** Wide banner strip with a tab on top and a tab on the right (slide 25). */
const BANNER_PANEL = [
  ['M', 0, 0], ['L', 0.2320, 0], ['L', 0.2309, 0.0172],
  ['C', 0.2309, 0.1322, 0.2597, 0.2254, 0.2952, 0.2254],
  ['C', 0.3307, 0.2254, 0.3594, 0.1322, 0.3594, 0.0172],
  ['L', 0.3584, 0], ['L', 0.9357, 0], ['L', 0.9357, 0.4509],
  ['C', 0.9712, 0.4509, 1, 0.5441, 1, 0.6590],
  ['C', 1, 0.7741, 0.9712, 0.8673, 0.9357, 0.8673],
  ['L', 0.9357, 1], ['L', 0, 1], ['Z'],
];

/* ------------------------------------------------------------------ helpers */
const pptx = new PptxGenJS();

function freeform(s, outline, x, y, w, h, color, rotate) {
  const points = outline.map((seg) => {
    if (seg[0] === 'M') return { x: seg[1] * w, y: seg[2] * h, moveTo: true };
    if (seg[0] === 'L') return { x: seg[1] * w, y: seg[2] * h };
    if (seg[0] === 'C') {
      return {
        x: seg[5] * w, y: seg[6] * h,
        curve: { type: 'cubic', x1: seg[1] * w, y1: seg[2] * h, x2: seg[3] * w, y2: seg[4] * h },
      };
    }
    return { close: true };
  });
  s.addShape('custGeom', { x, y, w, h, points, fill: { color }, rotate: rotate || 0 });
}

const rect = (s, x, y, w, h, color) => s.addShape('rect', { x, y, w, h, fill: { color } });
const circle = (s, x, y, d, color) => s.addShape('ellipse', { x, y, w: d, h: d, fill: { color } });
const hline = (s, x, y, w) => s.addShape('line', { x, y, w, h: 0, line: { color: BLUE, width: 1.5 } });

/** Square block with a round notch in one edge and a matching tab on another. */
function puzzleBlock(s, block, notch, tab) {
  rect(s, block[0], block[1], block[2], block[3], block[4]);
  circle(s, notch[0], notch[1], notch[2], WHITE);
  circle(s, tab[0], tab[1], tab[2], block[4]);
}

function txt(s, content, opts) {
  s.addText(content, Object.assign(
    { fontFace: SANS, fontSize: 10.5, color: GREY, valign: 'top' }, opts));
}

/** Run options for the PT Serif display face. */
const serif = (fontSize, color) => ({ fontFace: SERIF, fontSize, bold: true, color: color || INK });

/** Two-tone display heading; each part is [text, colour, breakAfter]. */
function heading(s, box, size, parts, align) {
  const runs = parts.map((p) => ({
    text: p[0],
    options: Object.assign(serif(size, p[1]), { breakLine: !!p[2] }),
  }));
  txt(s, runs, { x: box[0], y: box[1], w: box[2], h: box[3], align: align || 'left' });
}

/** Grey body copy at 150% leading, justified unless told otherwise. */
function body(s, box, str, align) {
  txt(s, str, {
    x: box[0], y: box[1], w: box[2], h: box[3],
    align: align || 'justify', lineSpacingMultiple: 1.5,
  });
}

/** Letter-spaced "Write Something Here" caption. */
function caption(s, box, opts) {
  txt(s, NOTE, Object.assign({
    x: box[0], y: box[1], w: box[2], h: box[3],
    fontSize: 11, charSpacing: 3, lineSpacingMultiple: 1.5,
  }, opts));
}

/** Letter-spaced "Presentation Template" kicker on the cover slides. */
function kicker(s, box) {
  txt(s, 'Presentation Template', {
    x: box[0], y: box[1], w: box[2], h: box[3],
    align: 'center', charSpacing: 3, lineSpacingMultiple: 1.5,
  });
}

/** Serif "246 . 324 : Music" strap-line. */
function strap(s, box, color) {
  txt(s, MUSIC_TAG, Object.assign(
    { x: box[0], y: box[1], w: box[2], h: box[3], align: 'justify', lineSpacingMultiple: 1.5 },
    serif(14, color || GREY)));
}

function readMore(s, x, y) {
  txt(s, 'Read More', {
    x, y, w: 1.792, h: 0.531, shape: 'roundRect', fill: { color: BLUE },
    align: 'center', valign: 'middle', fontSize: 14, color: WHITE,
  });
}

/** Labelled percentage bar: caption row on top, track + coloured fill below. */
function progressBar(s, o) {
  txt(s, 'Lorem ipsum dolor', {
    x: o.x + o.labelDx, y: o.y, w: o.labelW, h: 0.236,
    fontSize: 8, bold: true, color: INK, valign: 'middle', wrap: false,
  });
  txt(s, '82%', {
    x: o.x + o.pctDx, y: o.y, w: o.pctW, h: 0.236,
    fontSize: 8, bold: true, color: INK, align: 'center', valign: 'middle',
  });
  rect(s, o.x, o.y + 0.379, o.w, 0.238, TRACK);
  rect(s, o.x, o.y + 0.379, o.fill, 0.238, o.color);
}

const wideBar = (s, x, y, fill, color) =>
  progressBar(s, { x, y, w: 4.169, fill, color, labelDx: 0.046, labelW: 1.222, pctDx: 3.599, pctW: 0.571 });
const slimBar = (s, x, y, fill, color) =>
  progressBar(s, { x, y, w: 2.286, fill, color, labelDx: 0.025, labelW: 1.344, pctDx: 1.869, pctW: 0.521 });

/** Number / "Service" / micro-copy lockup sitting on a jigsaw tile. */
function serviceCard(s, x, y, num, opts) {
  const o = opts || {};
  txt(s, num, {
    x, y, w: 0.697, h: 0.636, fontSize: 24,
    color: o.numColor || WHITE, align: 'center', lineSpacingMultiple: 1.5,
  });
  txt(s, o.title || 'Service',
    Object.assign({ x: x + 0.066, y: y + 0.816, w: 1.55, h: 0.337 }, serif(14, o.titleColor || WHITE)));
  txt(s, L_CARD, {
    x: x + 0.066, y: y + 1.138, w: 1.492, h: 0.576, fontSize: 10,
    color: o.bodyColor || WHITE, align: 'justify', lineSpacingMultiple: 1.5,
  });
}

/** Numbered list entry used on the light "service" slides. */
function numberedItem(s, o) {
  const dx = o.dx === undefined ? 0.067 : o.dx;
  const titleDy = o.titleDy === undefined ? 0.816 : o.titleDy;
  const color = o.color || INK;
  txt(s, o.num, { x: o.x, y: o.y, w: 0.697, h: 0.636, fontSize: 24, color, lineSpacingMultiple: 1.5 });
  txt(s, o.title, Object.assign({ x: o.x + dx, y: o.y + titleDy, w: o.titleW || 1.55, h: 0.37 }, serif(16, color)));
  body(s, [o.x + dx, o.y + titleDy + 0.476, 2.313, 0.896], L1);
}

/** "portfolio" caption + micro copy pair. */
function portfolioItem(s, x, y) {
  txt(s, 'portfolio', Object.assign({ x, y, w: 1.55, h: 0.37 }, serif(16)));
  txt(s, L_PORT, { x, y: y + 0.476, w: 1.779, h: 0.865, lineSpacingMultiple: 1.5 });
}

/** Stand-in for a bitmap mockup from the source deck. */
function photo(s, x, y, w, h) {
  txt(s, '[image]', {
    x, y, w, h, shape: 'rect', fill: { color: 'ECECEC' }, line: { color: 'D5D5D5', width: 1 },
    align: 'center', valign: 'middle', fontSize: 12, color: '9E9E9E',
  });
}

/** Bold serif page number, bottom-right, inherited from the slide master. */
function pageNumber(s, n) {
  txt(s, n + '.', Object.assign(
    { x: 12.528, y: 6.844, w: 0.346, h: 0.303, align: 'center', wrap: false },
    serif(12, CHARCOAL)));
}

/* ---------------------------------------------------- recurring decorations */
const darkCorner = (s) =>
  puzzleBlock(s, [0, 6.233, 1.05, 1.267, CHARCOAL], [0.879, 6.954, 0.341], [0.182, 6.062, 0.341]);
const skyCornerTR = (s) =>
  puzzleBlock(s, [11.583, 0, 1.751, 1.451, SKY], [12.579, 1.215, 0.471], [11.347, 0.251, 0.471]);

/* ------------------------------------------------------------------- slides */
const slides = [];

// 1 - cover
slides.push((s) => {
  kicker(s, [4.604, 2.046, 4.125, 0.346]);
  heading(s, [5.148, 2.683, 3.037, 2.121], 60, [['Music', INK, true], ['Max', BLUE]], 'center');
});

// 2 - What Is Music? (empty portrait photo panel on the left)
slides.push((s) => {
  heading(s, [7.012, 1.748, 3.037, 1.717], 48, [['What Is ', INK, true], ['Music ?', INK]]);
  body(s, [7.012, 3.865, 4.278, 1.395], L6);
  readMore(s, 7.125, 5.781);
  caption(s, [10.65, 4.496, 4.125, 0.346], { rotate: 90 });
  circle(s, 5.164, 1.703, 1.032, WHITE);
  circle(s, 12.431, 1.455, 0.564, WHITE);
});

// 3 - What Is Music? (left aligned)
slides.push((s) => {
  darkCorner(s);
  heading(s, [1.459, 1.368, 4.636, 0.774], 40, [['What Is Music ?', INK]]);
  body(s, [1.459, 2.84, 4.278, 1.395], L6);
  readMore(s, 1.573, 4.947);
  freeform(s, TAB_PANEL, 11.587, 0, 1.761, 2.411, SKY);
  caption(s, [4.094, 6.698, 3.429, 0.379]);
  circle(s, 7.086, 4.477, 0.825, BLUE);
  hline(s, 7.499, 0.979, 2.019);
});

// 4 - About My Studio (photo left, copy right)
slides.push((s) => {
  heading(s, [6.08, 2.401, 5.508, 0.841], 44, [['About My Studio', INK]]);
  body(s, [6.08, 3.75, 3.487, 1.426], L4);
  readMore(s, 6.193, 5.865);
  caption(s, [0.808, 0.844, 3.429, 0.379]);
  hline(s, 4.478, 1.076, 3.126);
  freeform(s, TAB_PANEL, 11.872, 0, 1.476, 2.021, SKY);
});

// 5 - About My Studio + statistic bars
slides.push((s) => {
  heading(s, [1.201, 0.995, 5.508, 0.841], 44, [['About My Studio', INK]]);
  body(s, [1.201, 2.345, 5.279, 1.161], L8);
  strap(s, [9.867, 3.119, 2.409, 0.417], SILVER);
  wideBar(s, 7.782, 4.458, 3.598, INK);
  wideBar(s, 7.782, 5.667, 3.598, BLUE);
});

// 6 - About My Studio (right aligned, blue lead-in)
slides.push((s) => {
  heading(s, [8.809, 1.833, 3.566, 1.582], 44, [['About My ', BLUE, true], ['Studio', INK]]);
  body(s, [8.809, 4.24, 3.487, 1.426], L4);
  darkCorner(s);
  freeform(s, TAB_PANEL, 12.603, 0, 0.746, 1.021, SKY);
  caption(s, [-0.831, 2.74, 2.854, 0.392], { rotate: 270 });
  strap(s, [5.257, 0.65, 2.792, 0.454]);
  hline(s, 1.937, 0.882, 2.438);
});

// 7 - About My Studio (copy on the right rail)
slides.push((s) => {
  heading(s, [1.044, 2.168, 3.566, 1.582], 44, [['About My Studio', INK]]);
  body(s, [10.282, 3.75, 2.248, 2.487], L4);
  hline(s, 10.297, 2.781, 2.099);
  freeform(s, TAB_PANEL, 11.872, 0, 1.476, 2.021, SILVER);
  readMore(s, 1.044, 5.664);
  txt(s, 'Lorem Ipsum Dolor Sita', {
    x: 1.044, y: 1.64, w: 3.309, h: 0.335, bold: true, charSpacing: 3, lineSpacingMultiple: 1.5,
  });
});

// 8 - Our Studio Service: four jigsaw tiles in a row
slides.push((s) => {
  puzzleBlock(s, [1.916, 2.4, 2.164, 2.096, BLUE], [3.797, 2.739, 0.564], [2.247, 4.214, 0.564]);
  rect(s, 4.362, 2.4, 2.164, 2.096, INK);
  rect(s, 6.808, 2.4, 2.164, 2.096, BLUE);
  puzzleBlock(s, [9.254, 2.4, 2.164, 2.096, INK], [8.972, 3.593, 0.564], [10.522, 2.118, 0.564]);
  heading(s, [3.022, 0.667, 7.397, 0.909], 48, [['Our Studio Service', INK]], 'center');
  body(s, [2.247, 5.761, 8.839, 0.865], L7, 'center');
  caption(s, [3.342, 5.206, 6.649, 0.346], { align: 'center' });
  [[2.181, 2.48, '01.'], [4.635, 2.478, '02.'], [7.081, 2.48, '03.'], [9.698, 2.478, '04.']]
    .forEach((c) => serviceCard(s, c[0], c[1], c[2]));
});

// 9 - Numbered services, photo panel right
slides.push((s) => {
  numberedItem(s, { x: 0.912, y: 1.032, num: '01.', title: 'Service One' });
  numberedItem(s, { x: 0.912, y: 4.074, num: '02.', title: 'Service Two' });
  numberedItem(s, { x: 4.203, y: 1.032, num: '03.', title: 'Service Three', titleW: 1.834, color: BLUE });
  numberedItem(s, { x: 4.203, y: 4.074, num: '04.', title: 'Service Four', color: BLUE });
  heading(s, [7.315, 6.294, 4.678, 0.572], 28, [['Our Studio Service', INK]], 'center');
  circle(s, 8.394, 0.821, 0.789, BLUE);
});

// 10 - Staggered jigsaw tiles above the headline
slides.push((s) => {
  freeform(s, CARD_PANEL, 1.534, 0.525, 2.164, 2.378, BLUE);
  rect(s, 4.262, 1.703, 2.164, 2.096, INK);
  rect(s, 6.99, 0.525, 2.164, 2.096, BLUE);
  freeform(s, CARD_PANEL, 9.719, 1.421, 2.164, 2.378, INK, 180);
  [[1.799, 0.605, '01.'], [4.536, 1.781, '02.'], [7.264, 0.605, '03.'], [10.163, 1.781, '04.']]
    .forEach((c) => serviceCard(s, c[0], c[1], c[2]));
  heading(s, [1.103, 4.784, 3.574, 1.582], 44, [['Our Studio Service', INK]]);
  body(s, [6.426, 5.335, 5.865, 1.161], L6);
  strap(s, [6.414, 4.695, 2.409, 0.417], BLUE);
});

// 11 - Two numbered services + headline block
slides.push((s) => {
  numberedItem(s, { x: 0.912, y: 1.032, num: '01.', title: 'Service One' });
  numberedItem(s, { x: 0.912, y: 4.074, num: '02.', title: 'Service Two' });
  heading(s, [4.382, 1.094, 3.574, 1.582], 44, [['Our Studio Service', INK]]);
  body(s, [4.423, 3.812, 3.41, 1.161], L3);
  strap(s, [4.411, 3.172, 2.409, 0.454]);
  readMore(s, 4.486, 5.641);
});

// 12 - 2x2 jigsaw grid
slides.push((s) => {
  [[1.043, 0.886, SLATE, 1.308, 0.966, '01.'],
   [1.043, 4.160, BLUE, 1.308, 4.240, '02.'],
   [4.151, 0.886, BLUE, 4.416, 0.966, '03.'],
   [4.151, 4.160, INK, 4.416, 4.240, '04.']].forEach((t) => {
    freeform(s, CARD_PANEL, t[0], t[1], 2.164, 2.378, t[2]);
    serviceCard(s, t[3], t[4], t[5]);
  });
  heading(s, [7.697, 1.388, 3.574, 1.582], 44, [['Our Studio Service', INK]]);
  body(s, [7.697, 3.542, 4.278, 1.395], L6);
  readMore(s, 7.81, 5.459);
  freeform(s, TAB_PANEL, 11.976, 0, 1.373, 1.879, SILVER);
});

// 13 - Four numbered services across the top
slides.push((s) => {
  [['01.', 'Service One', 0.825, INK, 1.55],
   ['02.', 'Service Two ', 3.826, BLUE, 1.55],
   ['03.', 'Service Three', 6.827, INK, 1.856],
   ['04.', 'Service Four', 9.827, BLUE, 1.55]].forEach((c) => {
    numberedItem(s, { x: c[2], y: 0.778, dx: 0, titleDy: 0.8, num: c[0], title: c[1], color: c[3], titleW: c[4] });
  });
});

// 14 - Gallery (white type sits over the photo band)
slides.push((s) => {
  heading(s, [2.968, 1.486, 7.397, 0.909], 48, [['Gallery', WHITE]], 'center');
  body(s, [2.247, 6.195, 8.839, 0.6], L4, 'center');
  caption(s, [3.342, 5.688, 6.649, 0.346], { align: 'center' });
  circle(s, 4.159, 1.554, 0.772, WHITE);
  circle(s, 8.402, 1.554, 0.772, WHITE);
});

// 15 - New Music Studio, copy right
slides.push((s) => {
  heading(s, [9.029, 1.798, 3.721, 1.447], 40, [['New Music ', BLUE, true], ['Studio', INK]]);
  body(s, [9.041, 4.598, 3.313, 1.161], L3);
  strap(s, [9.029, 3.958, 2.409, 0.454]);
  freeform(s, TAB_PANEL, 11.945, -0.216, 1.172, 1.604, SILVER, 90);
});

// 16 - New Music Studio, portfolio rail right
slides.push((s) => {
  heading(s, [0.966, 1.486, 3.721, 1.447], 40, [['New Music ', BLUE, true], ['Studio', INK]]);
  portfolioItem(s, 10.854, 1.557);
  portfolioItem(s, 10.854, 4.499);
  body(s, [0.966, 3.396, 3.326, 1.161], L3);
  readMore(s, 1.029, 5.225);
});

// 17 - New Music Studio, big photo left
slides.push((s) => {
  caption(s, [-0.834, 1.989, 2.878, 0.379], { rotate: 270 });
  heading(s, [8.749, 2.17, 3.721, 1.447], 40, [['New Music ', BLUE, true], ['Studio', INK]]);
  body(s, [8.761, 4.971, 3.313, 1.161], L3);
  strap(s, [8.749, 4.331, 2.409, 0.454]);
  freeform(s, TAB_PANEL, 12.075, 0, 1.274, 1.744, SILVER);
  circle(s, 7.229, 1.757, 0.513, BLUE);
});

// 18 - Wide banner photo + three portfolio blurbs
slides.push((s) => {
  heading(s, [0.977, 0.926, 5.386, 0.774], 40, [['New Music ', BLUE], ['Studio', INK]]);
  portfolioItem(s, 4.484, 5.263);
  portfolioItem(s, 7.499, 5.263);
  portfolioItem(s, 10.513, 5.263);
  hline(s, 6.907, 1.313, 5.546);
  body(s, [0.977, 2.623, 2.546, 1.426], L2);
});

// 19 - Photo bottom-left, portfolio pair right
slides.push((s) => {
  portfolioItem(s, 7.796, 4.721);
  portfolioItem(s, 10.51, 4.721);
  heading(s, [1.31, 0.97, 5.386, 0.774], 40, [['New Music ', BLUE], ['Studio', INK]]);
  freeform(s, TAB_PANEL, 12.075, 0, 1.274, 1.744, SILVER);
  body(s, [7.796, 2.75, 4.362, 0.896], L3);
  strap(s, [7.796, 1.929, 2.409, 0.454]);
});

// 20 - Four-up photo strip with centred caption
slides.push((s) => {
  heading(s, [4.015, 5.305, 5.386, 0.774], 40, [['New Music ', BLUE], ['Studio', INK]], 'center');
  body(s, [2.907, 6.184, 7.602, 0.631], L3, 'center');
});

// 21 - Tall photo left, copy right
slides.push((s) => {
  heading(s, [6.708, 1.719, 3.721, 1.447], 40, [['New Music ', BLUE, true], ['Studio', INK]]);
  body(s, [6.721, 4.519, 3.313, 1.161], L3);
  strap(s, [6.708, 3.879, 2.409, 0.454]);
  freeform(s, TAB_PANEL, 12.075, 0, 1.274, 1.744, SILVER);
});

// 22 - Centred intro over a four-up photo row
slides.push((s) => {
  puzzleBlock(s, [0, 6.049, 1.751, 1.451, SKY], [0.283, 5.813, 0.471], [1.515, 6.778, 0.471]);
  heading(s, [4.015, 0.526, 5.386, 0.774], 40, [['New Music ', BLUE], ['Studio', INK]], 'center');
  body(s, [2.661, 1.728, 8.095, 0.896], L6, 'center');
  caption(s, [5.276, 3.072, 2.878, 0.346], { align: 'center' });
  hline(s, 8.625, 3.287, 0.776);
  hline(s, 3.958, 3.287, 0.776);
});

// 23 - Profile: Stevano
slides.push((s) => {
  circle(s, 5.164, 1.703, 1.032, WHITE);
  heading(s, [6.97, 1.435, 4.072, 0.909], 48, [['Stevano', INK]]);
  body(s, [6.97, 2.87, 4.278, 1.395], L5);
  wideBar(s, 7.012, 4.934, 3.598, INK);
  wideBar(s, 7.012, 5.809, 3.598, BLUE);
  caption(s, [10.65, 2.562, 4.125, 0.346], { rotate: 90 });
});

// 24 - My Teams In Studio, staggered portraits
slides.push((s) => {
  freeform(s, CARD_PANEL, 10.445, 0.654, 2.164, 2.378, PALE, 180);
  heading(s, [1.103, 4.784, 3.574, 1.582], 44, [['My Teams In Studio', INK]]);
  body(s, [6.426, 5.335, 5.865, 1.161], L6);
  strap(s, [6.414, 4.695, 2.409, 0.454]);
  hline(s, 1.534, 3.75, 0.987);
  hline(s, 6.99, 3.75, 0.987);
});

// 25 - Profile: Jessica
slides.push((s) => {
  freeform(s, BANNER_PANEL, 0, 6.146, 4.386, 1.354, PALE);
  heading(s, [4.86, 1.563, 2.078, 0.707], 36, [['Jessica ', INK]]);
  body(s, [4.86, 2.461, 2.286, 0.896], L1);
  slimBar(s, 4.86, 4.083, 1.973, BLUE);
  slimBar(s, 4.86, 5.115, 1.973, INK);
  heading(s, [7.842, 0.497, 3.123, 0.438], 20, [['My Team In Studio', BLUE]]);
  hline(s, 10.965, 0.715, 1.828);
});

// 26 - Team grid with names underneath
slides.push((s) => {
  heading(s, [3.696, 0.546, 6.024, 0.841], 44, [['My Teams ', BLUE], ['In Studio', INK]], 'center');
  [['Stevani', 1.155], ['Joe Brams', 4.156], ['Christina', 7.157], ['Jhonatan', 10.158]].forEach((m) => {
    txt(s, m[0], Object.assign({ x: m[1], y: 5.244, w: 2.103, h: 0.37, align: 'center' }, serif(16)));
    body(s, [m[1] - 0.117, 5.713, 2.338, 0.896], L1, 'center');
  });
});

// 27 - Phone mockups
slides.push((s) => {
  puzzleBlock(s, [0, 3.828, 3.044, 3.672, CHARCOAL], [2.549, 5.918, 0.989], [0.526, 3.333, 0.989]);
  photo(s, 0.901, 1.062, 2.664, 5.375);
  photo(s, 3.916, 1.062, 2.664, 5.375);
  skyCornerTR(s);
  heading(s, [7.664, 1.687, 3.683, 1.447], 40, [['About The', INK, true], ['Mockup', INK]]);
  body(s, [7.664, 3.698, 4.278, 1.395], L6);
  readMore(s, 7.777, 5.642);
  caption(s, [10.783, 4.103, 4.125, 0.346], { rotate: 90 });
});

// 28 - Desktop mockup
slides.push((s) => {
  photo(s, 6.542, 1.089, 4.735, 4.735);
  heading(s, [1.185, 1.431, 3.536, 1.447], 40, [['About The', INK, true], ['Mockup', INK]]);
  body(s, [1.185, 3.312, 4.278, 1.395], L6);
  readMore(s, 1.298, 5.559);
  skyCornerTR(s);
  caption(s, [10.65, 5.097, 4.125, 0.346], { rotate: 90 });
});

// 29 - 100% stacked bar chart
slides.push((s) => {
  s.addChart(pptx.ChartType.bar, [
    { name: 'Series 1', labels: ['1', '2', '3'], values: [4.3, 2.5, 3.5] },
    { name: 'Series 2', labels: ['1', '2', '3'], values: [2.4, 4.4, 1.8] },
    { name: 'Series 3', labels: ['1', '2', '3'], values: [2, 2, 3] },
  ], {
    x: 1.615, y: 0.692, w: 10.733, h: 3.431,
    barDir: 'bar', barGrouping: 'percentStacked', barGapWidthPct: 150, barOverlapPct: 100,
    chartColors: [BLUE, 'E7E7E7', CHARCOAL],
    showLegend: false, showTitle: false, showValue: false,
    catAxisLineShow: true, catAxisLineColor: 'DBDBDB', catAxisMajorTickMark: 'none',
    catAxisLabelColor: '616161', catAxisLabelFontFace: SANS, catAxisLabelFontSize: 8,
    valAxisLineShow: false, valAxisLabelFormatCode: '0%', valAxisMajorTickMark: 'none',
    valAxisLabelPos: 'nextTo',
    valAxisLabelColor: '616161', valAxisLabelFontFace: SANS, valAxisLabelFontSize: 8,
    catGridLine: { style: 'none' }, valGridLine: { style: 'none' },
  });
  puzzleBlock(s, [0, 6.162, 1.615, 1.338, SKY], [0.261, 5.944, 0.435], [1.397, 6.834, 0.435]);
  caption(s, [-1.537, 3.577, 4.125, 0.346], { rotate: 90, align: 'right' });
  [[6.538, '01.'], [8.538, '02.'], [10.538, '03.']].forEach((c) => {
    serviceCard(s, c[0], 4.933, c[1], { title: 'Chart', numColor: INK, titleColor: INK, bodyColor: GREY });
  });
  heading(s, [2.181, 4.883, 3.935, 0.909], 48, [['Our Chart', INK]]);
});

// 30 - Thank You!
slides.push((s) => {
  kicker(s, [4.604, 2.046, 4.125, 0.346]);
  heading(s, [5.148, 2.683, 3.037, 2.121], 60, [['Thank', INK, true], ['You!', BLUE]], 'center');
});

/* --------------------------------------------------------------------- build */
pptx.defineLayout({ name: 'WIDE_16x9', width: 40 / 3, height: 7.5 }); // 12192000 x 6858000 EMU
pptx.layout = 'WIDE_16x9';
pptx.title = 'Music Max - Presentation Template';

slides.forEach((build, i) => {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  build(s);
  pageNumber(s, i + 1);
});

pptx.writeFile({
  fileName: path.join(__dirname, '13797095-aa23-45c2-aa45-2f2f821c5c64_grok_final.pptx'),
});
