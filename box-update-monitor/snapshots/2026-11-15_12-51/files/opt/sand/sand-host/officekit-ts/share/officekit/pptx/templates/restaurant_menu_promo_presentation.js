/**
 * "restaurant" deck - recreated with pptxgenjs.
 *
 * Slide size 13.333 x 7.5 in (16:9). Black background, white type,
 * red (#FF0000 / #EE0000) accent blocks. Titles are Playfair Display,
 * body copy is Source Sans Pro.
 *
 * Raster photos in the source deck are drawn here as flat "[image]"
 * placeholder rectangles; vector artwork (world map, icons) is rebuilt
 * from native pptxgenjs shapes.
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ */
/* palette / type                                                      */
/* ------------------------------------------------------------------ */

const BLACK = '000000';
const NEAR_BLACK = '0D0D0D'; // tx1 lumMod 95% - closing slide background
const WHITE = 'FFFFFF';
const RED = 'FF0000';
const RED_DEEP = 'EE0000';
const GREY_95 = 'F2F2F2';
const GREY_85 = 'D9D9D9';
const GREY_65 = 'A6A6A6';
const GREY_50 = '808080';
const GREY_25 = '404040';
// Stand-in tone for the two photographs, picked to sit at roughly the
// average brightness of the originals against the black page.
const PLACEHOLDER_FILL = '2F2C29';
const PLACEHOLDER_TEXT = '8C8C8C';

const SERIF = 'Playfair Display';
const SANS = 'Source Sans Pro';

// PowerPoint's stock text-box insets, in points: [left, right, bottom, top].
const PPT_INSET = [7.2, 7.2, 3.6, 3.6];

/* ------------------------------------------------------------------ */
/* copy                                                                */
/* ------------------------------------------------------------------ */

const TITLE = 'Good Food Can Make Our Day Cheerful';

const LOREM =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
  'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud ' +
  'exercitation ullamco laboris nisi ut aliquip ex ea. Lorem ipsum dolor sit amet, ' +
  'consectetur adipiscing elit.';

const LOREM_HALF =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
  'incididunt ut labore et dolore magna aliqua.';

const LOREM_CLIPPED =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
  'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud ' +
  'exercitation ullamco laboris nisi ut aliquip ex ea.';

const LOREM_LONG = LOREM_CLIPPED + ' ' + LOREM_HALF;

const LOREM_ALT =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
  'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud ' +
  'exercitation ullamco laboris nisi ut aliquipas ex ea. Lorem ipsum dolor sit amet ' +
  'consectetur adipiscing elit.';

const LOREM_EXA =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
  'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud ' +
  'exercitation ullamco laboris nisi ut aliquip exa. Lorem ipsum dolor sit amet, ' +
  'consectetur adipiscing elit.';

const LOREM_UTIN =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
  'incididunt ut labore et dolore magna aliqua. utin enim ad minim veniam, quis nostrud ' +
  'exercitation ullamco laboris nisi ut aliquip.';

const LOREM_TINY =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
  'incididunt ut labore et dolore.';

const LOREM_NOSTRU =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
  'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam quis nostru ' +
  'exercitation ullamco laboris nisi ut aliquip ex ea.';

const COMPANY =
  'A company is an association or collection of individuals, whether natural persons. ';

/* ------------------------------------------------------------------ */
/* generic helpers                                                     */
/* ------------------------------------------------------------------ */

function rect(slide, x, y, w, h, color) {
  slide.addShape('rect', { x, y, w, h, fill: { color }, line: { type: 'none' } });
}

/**
 * Every text box in this deck is white Source Sans Pro, top-anchored and
 * grown to fit its copy, so those are the defaults here.
 */
function text(slide, str, opts) {
  slide.addText(str, Object.assign({
    valign: 'top', fit: 'resize', fontFace: SANS, color: WHITE,
  }, opts));
}

/** Section / slide headline - Playfair Display 28pt bold. */
function title(slide, x, y, opts = {}) {
  text(slide, opts.text || TITLE, {
    x, y,
    w: opts.w || 4.301,
    h: opts.h || 0.942,
    margin: 0,
    align: opts.align || 'left',
    fontFace: SERIF, fontSize: 28, bold: true, charSpacing: 1.2,
  });
}

/** Justified 11pt body paragraph on double leading. */
function body(slide, x, y, w, h, str, opts = {}) {
  text(slide, str, {
    x, y, w, h,
    margin: opts.margin === undefined ? 0 : opts.margin,
    align: opts.align || 'justify',
    lineSpacingMultiple: opts.lineSpacingMultiple || 2,
    fontSize: 11,
  });
}

/** Slide-12 style caption: centred, inset like a plain PowerPoint shape. */
function caption(slide, x, y, w, h, str) {
  body(slide, x, y, w, h, str, { align: 'center', margin: PPT_INSET });
}

/** "$25 / SHORT TITLE / Lorem ipsum dolor sit" figure block. */
function statBlock(slide, x, y) {
  text(slide, '$25', { x, y, w: 1.394, h: 0.841, fontSize: 44, bold: true });
  text(slide, 'SHORT TITLE', {
    x: x + 1.273, y: y + 0.107, w: 1.75, h: 0.37, fontSize: 16, bold: true,
  });
  text(slide, 'Lorem ipsum dolor sit', {
    x: x + 1.287, y: y + 0.374, w: 1.75, h: 0.349,
    lineSpacingMultiple: 1.5, fontSize: 11,
  });
}

/** "$67.568 / Lorem ipsum" caption pair. */
function figureCaption(slide, x, y, opts = {}) {
  const align = opts.align || 'left';
  text(slide, '$67.568', { x, y, w: 1.537, h: 0.438, align, fontSize: 20, bold: true });
  text(slide, 'Lorem ipsum', {
    x: x + (opts.dx || 0), y: y + (opts.dy || 0.315), w: opts.labelW || 1.537, h: 0.349,
    align, lineSpacingMultiple: 1.5, fontSize: 11,
  });
}

/**
 * Pair of thin progress tracks (75% / 50%) with labels.
 * (x, y) is the top-left corner of the upper track.
 */
function progressBars(slide, x, y) {
  const bars = [
    { dy: 0, dx: 0, fillW: 2.513, label: '75%' },
    { dy: 0.837, dx: -0.013, fillW: 1.684, label: '50%' },
  ];
  bars.forEach((b, i) => {
    slide.addShape('roundRect', {
      x: x + (i === 0 ? 0 : b.dx), y: y + b.dy, w: 3.36, h: 0.087,
      rectRadius: 0.0435, fill: { color: GREY_85 }, line: { type: 'none' },
    });
    slide.addShape('roundRect', {
      x: x - 0.013, y: y + b.dy, w: b.fillW, h: 0.087,
      rectRadius: 0.0435, fill: { color: RED_DEEP }, line: { type: 'none' },
    });
    text(slide, 'SHORT TEXT HERE', {
      x: x - 0.083 - i * 0.013, y: y + b.dy - 0.372, w: 1.549, h: 0.286,
      fontSize: 11, bold: true,
    });
    text(slide, b.label, {
      x: x + 2.71 - i * 0.013, y: y + b.dy - 0.489, w: 0.76, h: 0.404,
      align: 'center', fontSize: 18, bold: true,
    });
  });
}

/** Centred deck header used by the infographic slides. */
function sectionHeader(slide, heading, opts = {}) {
  text(slide, heading, {
    x: opts.x || 4.295, y: 0.493, w: opts.w || 4.833, h: 0.572,
    align: 'center', fontFace: SERIF, fontSize: 28, bold: true,
  });
  text(slide, 'Good Food', {
    x: 6.033, y: 1.083, w: 1.266, h: 0.37, align: 'center', fontSize: 16,
  });
}

/** Flat stand-in for a photograph that is not embedded in this script. */
function imagePlaceholder(slide, x, y, w, h, caption = '[image]') {
  slide.addText(caption, {
    x, y, w, h,
    align: 'center', valign: 'middle',
    fill: { color: PLACEHOLDER_FILL }, line: { type: 'none' },
    fontFace: SANS, fontSize: 12, color: PLACEHOLDER_TEXT,
  });
}

/* ------------------------------------------------------------------ */
/* icon stand-ins (the source deck uses freeform vector icons)         */
/* ------------------------------------------------------------------ */

function flagIcon(slide, cx, cy, color) {
  slide.addShape('rect', { x: cx - 0.23, y: cy - 0.23, w: 0.055, h: 0.46, fill: { color }, line: { type: 'none' } });
  slide.addShape('homePlate', { x: cx - 0.175, y: cy - 0.21, w: 0.4, h: 0.24, flipH: true, fill: { color }, line: { type: 'none' } });
}

function repeatIcon(slide, cx, cy, color) {
  slide.addShape('circularArrow', { x: cx - 0.28, y: cy - 0.24, w: 0.56, h: 0.48, fill: { color }, line: { type: 'none' } });
}

function gavelIcon(slide, cx, cy, color) {
  slide.addShape('rect', { x: cx - 0.05, y: cy - 0.26, w: 0.1, h: 0.52, rotate: 45, fill: { color }, line: { type: 'none' } });
  slide.addShape('rect', { x: cx - 0.02, y: cy - 0.26, w: 0.3, h: 0.16, rotate: 45, fill: { color }, line: { type: 'none' } });
}

function moneyBagIcon(slide, cx, cy, color) {
  slide.addShape('rect', { x: cx - 0.09, y: cy - 0.21, w: 0.18, h: 0.07, fill: { color }, line: { type: 'none' } });
  slide.addShape('ellipse', { x: cx - 0.16, y: cy - 0.14, w: 0.32, h: 0.34, fill: { color }, line: { type: 'none' } });
}

function handshakeIcon(slide, cx, cy, color) {
  slide.addShape('roundRect', { x: cx - 0.24, y: cy - 0.06, w: 0.28, h: 0.11, rotate: -12, rectRadius: 0.05, fill: { color }, line: { type: 'none' } });
  slide.addShape('roundRect', { x: cx - 0.04, y: cy - 0.05, w: 0.28, h: 0.11, rotate: 12, rectRadius: 0.05, fill: { color }, line: { type: 'none' } });
}

function peopleIcon(slide, cx, cy, color) {
  [-0.17, 0, 0.17].forEach((dx, i) => {
    const r = i === 1 ? 0.11 : 0.09;
    slide.addShape('ellipse', { x: cx + dx - r / 2, y: cy - 0.19 + (i === 1 ? -0.02 : 0), w: r, h: r, fill: { color }, line: { type: 'none' } });
    slide.addShape('roundRect', { x: cx + dx - 0.075, y: cy - 0.06, w: 0.15, h: 0.22, rectRadius: 0.04, fill: { color }, line: { type: 'none' } });
  });
}

function rocketIcon(slide, cx, cy, color) {
  slide.addShape('teardrop', { x: cx - 0.34, y: cy - 0.34, w: 0.68, h: 0.68, rotate: -45, fill: { color }, line: { type: 'none' } });
  slide.addShape('triangle', { x: cx - 0.42, y: cy + 0.02, w: 0.26, h: 0.22, rotate: -135, fill: { color }, line: { type: 'none' } });
  slide.addShape('triangle', { x: cx + 0.13, y: cy - 0.32, w: 0.26, h: 0.22, rotate: 45, fill: { color }, line: { type: 'none' } });
  slide.addShape('ellipse', { x: cx + 0.02, y: cy - 0.15, w: 0.15, h: 0.15, fill: { color: BLACK }, line: { type: 'none' } });
}

/**
 * One quadrant petal of the slide-26 ring: a thick arc band that tapers to an
 * arrow tip. Traced in inches inside its own 1.618 x 2.502 box, pointing right;
 * the ring's centre sits PETAL_PIVOT to the left of the box centre, so each
 * copy is rotated about its own centre and offset back onto the ring.
 */
const PETAL_PIVOT = 1.639;
const PETAL_W = 1.618;
const PETAL_H = 2.502;
const PETAL_PATH = [
  { x: 0.616, y: 0.000, moveTo: true },
  { x: 0.634, y: 0.020 },
  { x: 1.048, y: 0.870, curve: { type: 'cubic', x1: 0.838, y1: 0.262, x2: 0.984, y2: 0.553 } },
  { x: 1.070, y: 1.013 },
  { x: 1.233, y: 1.089 },
  { x: 1.618, y: 1.244, curve: { type: 'cubic', x1: 1.365, y1: 1.148, x2: 1.495, y2: 1.200 } },
  { x: 1.233, y: 1.405, curve: { type: 'cubic', x1: 1.495, y1: 1.290, x2: 1.365, y2: 1.344 } },
  { x: 1.069, y: 1.483 },
  { x: 1.048, y: 1.623 },
  { x: 0.634, y: 2.478, curve: { type: 'cubic', x1: 0.984, y1: 1.945, x2: 0.838, y2: 2.236 } },
  { x: 0.612, y: 2.502 },
  { x: 0.000, y: 1.890 },
  { x: 0.052, y: 1.815 },
  { x: 0.217, y: 1.244, curve: { type: 'cubic', x1: 0.156, y1: 1.649, x2: 0.217, y2: 1.455 } },
  { x: 0.052, y: 0.684, curve: { type: 'cubic', x1: 0.217, y1: 1.040, x2: 0.156, y2: 0.849 } },
  { x: 0.001, y: 0.615 },
  { close: true },
];

/** Place a petal on the ring centred at (cx, cy), tip pointing at `angle`. */
function ringPetal(slide, cx, cy, angle, color) {
  const rad = (angle * Math.PI) / 180;
  slide.addShape('custGeom', {
    x: cx + PETAL_PIVOT * Math.cos(rad) - PETAL_W / 2,
    y: cy - PETAL_PIVOT * Math.sin(rad) - PETAL_H / 2,
    w: PETAL_W,
    h: PETAL_H,
    rotate: (360 - angle) % 360,
    points: PETAL_PATH,
    fill: { color },
    line: { type: 'none' },
  });
}

/** Circle-with-tick badge (white disc, red mark). */
function checkBadge(slide, x, y, d) {
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color: WHITE }, line: { type: 'none' } });
  const cx = x + d / 2;
  const cy = y + d / 2;
  const r = d * 0.29;
  slide.addShape('ellipse', {
    x: cx - r, y: cy - r, w: r * 2, h: r * 2,
    fill: { type: 'none' }, line: { color: RED, width: 2 },
  });
  slide.addShape('rect', { x: cx - 0.13, y: cy - 0.01, w: 0.13, h: 0.035, rotate: 45, fill: { color: RED }, line: { type: 'none' } });
  slide.addShape('rect', { x: cx - 0.04, y: cy + 0.02, w: 0.24, h: 0.035, rotate: -45, fill: { color: RED }, line: { type: 'none' } });
}

/* ------------------------------------------------------------------ */
/* slide builders                                                      */
/* ------------------------------------------------------------------ */

/** 1 - cover, 30 - closing: one huge Playfair word, centred. */
function coverSlide(slide, word, x) {
  // Half-opaque black scrim; in the original it dims a full-bleed photo.
  slide.addShape('rect', {
    x: 0, y: 0, w: 13.333, h: 7.5,
    fill: { color: BLACK, transparency: 50 }, line: { type: 'none' },
  });
  text(slide, word, {
    x, y: 2.746, w: 6.219, h: 1.616,
    margin: [19.2, 19.2, 9.6, 9.6],
    align: 'center', fontFace: SERIF, fontSize: 80, bold: true,
  });
}

/**
 * The recurring "headline + one or two body columns" layout.
 * `blocks` is a list of { y, w, h, text } body paragraphs.
 */
function textColumn(slide, x, titleY, blocks, opts = {}) {
  title(slide, x, titleY, opts);
  blocks.forEach(b => body(slide, b.x === undefined ? x : b.x, b.y, b.w, b.h, b.text));
}

const slideBuilders = [
  // 1 - cover
  s => coverSlide(s, 'restaurant', 3.557),

  // 2 - red block left, copy right
  s => {
    rect(s, 0, 1.937, 1.594, 1.937, RED);
    textColumn(s, 7.023, 1.825, [
      { y: 3.159, w: 4.716, h: 1.429, text: LOREM_ALT },
      { y: 4.946, w: 4.716, h: 1.429, text: LOREM_ALT },
    ]);
  },

  // 3 - copy left, red block on the right edge
  s => {
    rect(s, 11.804, 3.627, 1.53, 1.937, RED);
    textColumn(s, 1.53, 0.99, [
      { y: 2.319, w: 4.91, h: 1.799, text: LOREM_LONG },
      { y: 4.505, w: 4.91, h: 1.058, text: LOREM_CLIPPED },
    ]);
  },

  // 4 - full-width band on top, headline + copy underneath
  s => {
    rect(s, 11.812, 0, 1.53, 3.75, RED);
    title(s, 1.521, 4.858);
    body(s, 7.133, 4.858, 4.68, 1.429, LOREM);
  },

  // 5 - tall red bar left, copy + figure right
  s => {
    rect(s, 0, 1.302, 1.375, 4.896, RED);
    title(s, 7.062, 1.22);
    body(s, 7.074, 2.425, 4.884, 1.429, LOREM_EXA);
    body(s, 7.062, 4.115, 4.884, 0.688, LOREM_HALF);
    statBlock(s, 7.062, 5.439);
  },

  // 6 - red diamond
  s => {
    s.addShape('diamond', { x: 4.042, y: 2.219, w: 3.125, h: 3.125, fill: { color: RED_DEEP }, line: { type: 'none' } });
    textColumn(s, 7.781, 1.015, [
      { y: 2.22, w: 4.301, h: 1.799, text: LOREM },
      { y: 4.327, w: 4.301, h: 1.799, text: LOREM },
    ]);
  },

  // 7 - copy left, red square right
  s => {
    s.addShape('rect', { x: 7.531, y: 3.962, w: 2.17, h: 2.17, fill: { color: RED_DEEP }, line: { type: 'none' } });
    textColumn(s, 1.377, 1.114, [
      { y: 2.321, w: 4.83, h: 1.429, text: LOREM },
      { y: 4.17, w: 4.83, h: 1.429, text: LOREM },
    ]);
  },

  // 8 - two stacked paragraphs, right hand column
  s => {
    textColumn(s, 7.672, 1.083, [
      { y: 2.352, w: 4.301, h: 1.799, text: LOREM },
      { y: 4.472, w: 4.301, h: 1.799, text: LOREM },
    ]);
  },

  // 9 - copy + progress bars, right
  s => {
    textColumn(s, 7.203, 1.019, [{ y: 2.191, w: 4.301, h: 1.429, text: LOREM_CLIPPED }]);
    progressBars(s, 7.212, 5.309);
  },

  // 10 - mirror of 9
  s => {
    textColumn(s, 1.516, 1.019, [{ y: 2.221, w: 4.301, h: 1.429, text: LOREM_CLIPPED }]);
    progressBars(s, 1.525, 5.309);
  },

  // 11 - full-bleed image band, two captions below
  s => {
    title(s, 1.042, 4.362);
    body(s, 1.042, 5.545, 5.0, 1.058, LOREM_UTIN);
    body(s, 7.291, 5.487, 5.0, 1.058, LOREM_UTIN);
  },

  // 12 - three centred captions
  s => {
    title(s, 4.427, 0.61, { align: 'center' });
    [1.301, 5.148, 9.214].forEach(x => caption(s, x, 4.952, 3.037, 1.159, LOREM_TINY));
  },

  // 13 - red panel right
  s => {
    rect(s, 8.656, 0, 4.677, 7.5, RED);
    title(s, 1.361, 2.163);
    body(s, 1.375, 3.344, 4.301, 1.799, LOREM);
    statBlock(s, 1.32, 5.533);
  },

  // 14 - red panel left
  s => {
    rect(s, 0, 0, 4.677, 7.5, RED);
    title(s, 7.825, 0.711);
    body(s, 7.825, 1.921, 4.301, 1.799, LOREM);
    body(s, 7.825, 4.061, 4.301, 0.688, LOREM_HALF);
    statBlock(s, 7.685, 5.81);
  },

  // 15 - three red icon discs with figures
  s => {
    const icons = [flagIcon, repeatIcon, gavelIcon];
    [1.187, 3.233, 5.28].forEach((y, i) => {
      s.addShape('ellipse', { x: 10.134, y, w: 1.033, h: 1.033, fill: { color: RED }, line: { type: 'none' } });
      icons[i](s, 10.65, y + 0.517, WHITE);
      figureCaption(s, 11.167, y + 0.096, { align: 'center', dx: -0.03, dy: 0.331 });
    });
  },

  // 16 - copy right of a full-height image
  s => {
    title(s, 5.874, 0.978);
    body(s, 6.979, 2.301, 4.301, 1.799, LOREM);
    body(s, 7.77, 4.421, 4.301, 1.799, LOREM);
  },

  // 17 - copy left of a full-height image
  s => {
    textColumn(s, 1.521, 0.916, [
      { y: 2.163, w: 4.301, h: 1.799, text: LOREM },
      { y: 4.266, w: 4.301, h: 1.799, text: LOREM },
    ]);
  },

  // 18 - red band across the top
  s => {
    rect(s, 0, 0, 13.333, 2.979, RED);
    title(s, 1.535, 5.194);
    body(s, 7.147, 5.149, 4.68, 1.429, LOREM);
  },

  // 19 - pull quote
  s => {
    text(
      s,
      [
        {
          text: '"Cooking is like love. It should be entered into with abandon or not at all.\u201C',
          options: { fontFace: SERIF, bold: true, breakLine: true },
        },
        { text: '\u2014 Harriet Van Horne', options: { fontFace: SANS } },
      ],
      {
        x: 1.354, y: 1.404, w: 6.257, h: 2.827,
        margin: 0, lineSpacingMultiple: 1.5, fontSize: 28,
      }
    );
  },

  // 20 - red headline plate top right
  s => {
    rect(s, 6.875, 0.782, 5.521, 1.885, RED);
    title(s, 7.485, 1.185, { align: 'center' });
    body(s, 1.058, 3.218, 4.808, 1.429, LOREM);
    body(s, 1.058, 5.048, 4.808, 1.429, LOREM);
  },

  // 21 - big red card
  s => {
    rect(s, 5.104, 1.042, 7.188, 5.417, RED);
    title(s, 5.844, 1.472);
    body(s, 5.859, 2.75, 5.664, 1.429, LOREM);
    body(s, 5.859, 4.483, 5.664, 1.429, LOREM);
  },

  // 22 - laptop photo left
  s => {
    imagePlaceholder(s, 0.93, 1.42, 5.79, 5.06);
    textColumn(s, 7.7, 1.209, [
      { y: 2.477, w: 4.301, h: 1.799, text: LOREM },
      { y: 4.497, w: 4.301, h: 1.799, text: LOREM },
    ]);
  },

  // 23 - tablet photo left, copy + progress bars right
  s => {
    imagePlaceholder(s, 0.9, 1.85, 3.74, 5.64);
    title(s, 2.757, 0.573, { w: 7.972, h: 0.471, align: 'center' });
    body(s, 6.043, 1.786, 5.915, 1.058, LOREM_NOSTRU);
    body(s, 6.043, 3.104, 5.915, 1.058, LOREM_NOSTRU);
    progressBars(s, 6.017, 5.652);
  },

  // 24 - F O O D grid
  s => {
    sectionHeader(s, 'Restaurant Infographic');
    const cards = [
      { x: 0.687, y: 3.004, hx: 1.254, hy: 2.654, align: 'right' },
      { x: 0.782, y: 5.396, hx: 1.35, hy: 5.046, align: 'right' },
      { x: 9.357, y: 3.07, hx: 9.357, hy: 2.72, align: 'left' },
      { x: 9.357, y: 5.605, hx: 9.357, hy: 5.255, align: 'left' },
    ];
    cards.forEach(c => {
      text(s, 'Short Text', {
        x: c.hx, y: c.hy, w: 2.662, h: 0.417,
        align: c.align, lineSpacingMultiple: 1.14,
        fontFace: SERIF, fontSize: 18, bold: true,
      });
      text(s, COMPANY, {
        x: c.x, y: c.y, w: 3.23, h: 0.789,
        align: c.align, lineSpacingMultiple: 2, fontSize: 11,
      });
    });
    const letters = [
      { t: 'F', x: 4.485, y: 2.352 },
      { t: 'O', x: 6.921, y: 2.352 },
      { t: 'O', x: 4.488, y: 4.779 },
      { t: 'D', x: 6.921, y: 4.779 },
    ];
    letters.forEach(l => {
      s.addText(l.t, {
        x: l.x, y: l.y, w: 1.921, h: 1.921,
        shape: 'rect', fill: { color: RED_DEEP }, line: { type: 'none' },
        align: 'center', valign: 'middle',
        fontFace: SANS, fontSize: 96, bold: true, color: WHITE,
      });
    });
  },

  // 25 - milestone path
  s => {
    const dash = { color: GREY_85, width: 3, dashType: 'dash' };
    s.addShape('line', { x: 1.935, y: 3.568, w: 2.411, h: 1.161, flipV: true, line: dash });
    s.addShape('line', { x: 4.407, y: 3.704, w: 2.411, h: 1.251, line: dash });
    s.addShape('line', { x: 7.007, y: 3.941, w: 2.386, h: 0.92, flipV: true, line: dash });

    const stops = [
      { x: 1.472, y: 4.336, icon: moneyBagIcon },
      { x: 3.852, y: 3.163, icon: handshakeIcon },
      { x: 6.232, y: 4.397, icon: peopleIcon },
    ];
    stops.forEach(st => {
      s.addShape('ellipse', { x: st.x, y: st.y, w: 0.987, h: 0.987, fill: { color: GREY_95 }, line: { type: 'none' } });
      st.icon(s, st.x + 0.4935, st.y + 0.4935, RED);
    });
    rocketIcon(s, 9.898, 3.586, RED);

    const subtitles = [
      { x: 1.122, y: 5.554, align: 'center' },
      { x: 3.479, y: 4.39, align: 'center' },
      { x: 5.86, y: 5.557, align: 'center' },
      { x: 9.064, y: 4.589, align: 'left' },
    ];
    subtitles.forEach(t => {
      text(s, 'Subtitle Here', {
        x: t.x, y: t.y, w: 1.732, h: 0.438,
        wrap: false, align: t.align, fontSize: 20,
      });
    });
    [
      { t: '70%', x: 9.257, y: 2.029 },
      { t: '$450M', x: 10.855, y: 3.094 },
    ].forEach(f => {
      text(s, f.t, {
        x: f.x, y: f.y, w: 1.552, h: 0.909,
        align: 'center', valign: 'middle', lineSpacingMultiple: 1.5,
        fontSize: 32, bold: true,
      });
    });
    sectionHeader(s, 'Restaurant Infographic');
  },

  // 26 - four-segment ring
  s => {
    sectionHeader(s, 'Restaurant Infographic');
    body(s, 1.151, 2.471, 5.719, 1.429, LOREM + ' ');
    // Four quadrant petals around a hollow centre at (10.299, 4.292).
    ringPetal(s, 10.299, 4.292, 45, GREY_95);
    ringPetal(s, 10.299, 4.292, 135, RED_DEEP);
    ringPetal(s, 10.299, 4.292, 225, RED_DEEP);
    ringPetal(s, 10.299, 4.292, 315, GREY_65);
    text(s, '50%', {
      x: 9.477, y: 3.933, w: 1.656, h: 0.673,
      margin: 0, align: 'center', fontSize: 40, bold: true,
    });
    progressBars(s, 1.149, 5.096);
  },

  // 27 - clustered bar chart
  s => {
    sectionHeader(s, 'Restaurant Infographic');
    s.addChart(
      'bar',
      [
        { name: 'Series 1', labels: ['Category 1', 'Category 2', 'Category 3'], values: [4.3, 2.5, 3.5] },
        { name: 'Series 2', labels: ['Category 1', 'Category 2', 'Category 3'], values: [2.4, 4.4, 1.8] },
        { name: 'Series 3', labels: ['Category 1', 'Category 2', 'Category 3'], values: [2, 2, 3] },
      ],
      {
        x: 1.107, y: 2.06, w: 5.477, h: 4.858,
        barDir: 'bar',
        barGrouping: 'clustered',
        barGapWidthPct: 326,
        barOverlapPct: -58,
        chartColors: [GREY_95, GREY_65, 'F20203'],
        showLegend: true,
        legendPos: 'b',
        legendColor: WHITE,
        legendFontFace: SANS,
        legendFontSize: 11,
        showValue: false,
        catAxisLabelColor: WHITE,
        catAxisLabelFontFace: SANS,
        catAxisLabelFontSize: 11,
        catAxisLineColor: GREY_85,
        valAxisLabelColor: WHITE,
        valAxisLabelFontFace: SANS,
        valAxisLabelFontSize: 11,
        valAxisLineShow: false,
        valGridLine: { style: 'solid', color: GREY_25, size: 0.75 },
        catGridLine: { style: 'none' },
        dataBorder: { pt: 0, color: BLACK },
      }
    );
    body(s, 7.618, 2.0, 4.301, 2.169, LOREM_LONG);
    body(s, 7.618, 4.429, 4.301, 2.169, LOREM_LONG);
  },

  // 28 - doughnut with call-outs
  s => {
    sectionHeader(s, 'Restaurant Infographic');
    s.addChart(
      'doughnut',
      [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr'], values: [7.1, 3.2, 3] }],
      {
        x: -0.27, y: 1.89, w: 7.361, h: 4.907,
        holeSize: 90,
        chartColors: [GREY_85, GREY_85, RED_DEEP],
        dataBorder: { pt: 0, color: BLACK },
        showLegend: false,
        showValue: false,
      }
    );
    checkBadge(s, 1.754, 2.183, 0.956);
    checkBadge(s, 2.153, 5.832, 0.956);
    checkBadge(s, 5.154, 3.986, 0.956);
    figureCaption(s, 1.873, 3.329);
    figureCaption(s, 2.228, 4.999, { labelW: 1.257 });
    figureCaption(s, 3.52, 4.134, { align: 'right' });
    body(s, 7.356, 1.926, 4.301, 2.239, LOREM, { lineSpacingMultiple: 2.5 });
    statBlock(s, 7.284, 5.753);
  },

  // 29 - world map with share bubbles
  s => {
    worldMap(s);
    const arcLine = { color: RED, width: 2.25, dashType: 'sysDash' };
    s.addShape('arc', { x: 2.226, y: 2.26, w: 1.623, h: 1.623, angleRange: [200, 342.5], line: arcLine });
    s.addShape('arc', { x: 2.235, y: 3.179, w: 3.089, h: 3.089, angleRange: [42.4, 106.5], line: arcLine });
    s.addShape('arc', { x: 8.204, y: 3.213, w: 3.089, h: 3.089, angleRange: [70.3, 147], line: arcLine });
    s.addShape('arc', { x: 8.852, y: 2.094, w: 2.679, h: 2.679, angleRange: [218.4, 295.7], line: arcLine });

    const bubbles = [
      { x: 3.562, y: 2.963, d: 0.704, fill: GREY_25, pct: '20%', tx: 3.302, ty: 3.039, tw: 1.255, th: 0.505, size: 16 },
      { x: 8.548, y: 3.032, d: 0.81, fill: GREY_65, pct: '15%', tx: 8.347, ty: 3.161, tw: 1.255, th: 0.505, size: 16 },
      { x: 4.739, y: 4.576, d: 1.043, fill: RED, pct: '45%', tx: 4.639, ty: 4.692, tw: 1.255, th: 0.707, size: 24 },
      { x: 7.88, y: 4.558, d: 0.903, fill: GREY_50, pct: '30%', tx: 7.722, ty: 4.681, tw: 1.255, th: 0.606, size: 20 },
    ];
    bubbles.forEach(b => {
      s.addShape('ellipse', {
        x: b.x, y: b.y, w: b.d, h: b.d,
        fill: { color: b.fill }, line: { color: WHITE, width: 2.25 },
      });
      text(s, b.pct, {
        x: b.tx, y: b.ty, w: b.tw, h: b.th,
        align: 'center', lineSpacingMultiple: 1.5, fontSize: b.size,
      });
    });
    [
      { x: 1.287, y: 2.92 },
      { x: 10.479, y: 2.62 },
      { x: 1.918, y: 5.665 },
      { x: 9.839, y: 5.54 },
    ].forEach(t => {
      text(s, 'Subtitle Here', {
        x: t.x, y: t.y, w: 1.732, h: 0.438, wrap: false, fontSize: 20,
      });
    });
    sectionHeader(s, 'Where Our Customers Are', { x: 4.039, w: 5.243 });
  },

  // 30 - closing
  s => coverSlide(s, 'thank you', 3.677),
];

/**
 * Vector world map (the reference deck draws it as a freeform group).
 * Outlines are stored as flat [x0,y0, x1,y1, ...] lists in a 0..1 space and
 * scaled into MAP_BOX, so the shape of the artwork is visible in the source.
 */
const MAP_BOX = { x: 3.306, y: 2.847, w: 7.005, h: 3.287 };

const WORLD_MAP = [
  [ // eurasia
    0.488,0.361, 0.513,0.315, 0.516,0.396, 0.537,0.290, 0.578,0.485, 0.656,0.497, 0.571,0.573,
    0.603,0.662, 0.626,0.637, 0.601,0.565, 0.652,0.611, 0.703,0.514, 0.793,0.628, 0.797,0.528,
    0.851,0.485, 0.863,0.380, 0.931,0.342, 0.910,0.434, 1.000,0.313, 0.868,0.245, 0.794,0.308,
    0.735,0.207, 0.676,0.313, 0.674,0.245, 0.577,0.334, 0.541,0.265,
  ],
  [ // north america
    0.200,0.144, 0.198,0.190, 0.138,0.177, 0.118,0.235, 0.078,0.256, 0.124,0.345, 0.256,0.376,
    0.220,0.400, 0.198,0.346, 0.119,0.358, 0.126,0.434, 0.184,0.477, 0.151,0.472, 0.174,0.542,
    0.206,0.561, 0.219,0.535, 0.198,0.548, 0.192,0.486, 0.230,0.472, 0.239,0.500, 0.270,0.351,
    0.308,0.314, 0.259,0.238, 0.244,0.330, 0.208,0.283, 0.204,0.221, 0.235,0.172, 0.219,0.190,
  ],
  [ // africa
    0.457,0.554, 0.443,0.597, 0.466,0.624, 0.460,0.672, 0.441,0.682, 0.449,0.701, 0.427,0.679,
    0.448,0.715, 0.461,0.687, 0.455,0.734, 0.473,0.727, 0.469,0.675, 0.502,0.689, 0.482,0.718,
    0.508,0.718, 0.514,0.768, 0.500,0.772, 0.521,0.806, 0.508,0.869, 0.524,0.966, 0.541,0.966,
    0.582,0.854, 0.578,0.744, 0.588,0.766, 0.612,0.704, 0.586,0.697, 0.565,0.579, 0.543,0.577,
    0.543,0.630, 0.567,0.634, 0.543,0.641, 0.537,0.706, 0.535,0.648, 0.501,0.632, 0.539,0.645,
    0.539,0.575, 0.506,0.572, 0.503,0.621, 0.500,0.541, 0.456,0.587,
  ],
  [ // greenland
    0.428,0.025, 0.400,0.038, 0.401,0.023, 0.375,0.023, 0.401,0.011, 0.391,0.000, 0.340,0.004,
    0.341,0.025, 0.328,0.014, 0.317,0.028, 0.297,0.021, 0.283,0.038, 0.282,0.059, 0.265,0.070,
    0.280,0.075, 0.275,0.092, 0.297,0.092, 0.325,0.168, 0.317,0.189, 0.322,0.218, 0.343,0.252,
    0.353,0.207, 0.397,0.168, 0.395,0.148, 0.404,0.152, 0.396,0.131, 0.410,0.100, 0.403,0.087,
    0.409,0.082, 0.404,0.066, 0.410,0.042,
  ],
  [ // australia
    0.898,0.937, 0.895,0.917, 0.876,0.877, 0.868,0.835, 0.863,0.870, 0.849,0.861, 0.848,0.844,
    0.840,0.837, 0.831,0.858, 0.823,0.851, 0.808,0.882, 0.787,0.903, 0.794,0.965, 0.809,0.970,
    0.828,0.954, 0.843,0.958, 0.849,0.969, 0.855,0.962, 0.861,0.992, 0.868,1.000, 0.886,0.996,
  ],
  [ // alaska
    0.001,0.177, 0.015,0.200, 0.010,0.207, 0.003,0.203, 0.000,0.211, 0.016,0.214, 0.016,0.224,
    0.003,0.244, 0.006,0.255, 0.015,0.265, 0.024,0.265, 0.025,0.279, 0.043,0.244, 0.046,0.249,
    0.042,0.259, 0.048,0.256, 0.050,0.245, 0.065,0.256, 0.069,0.254, 0.047,0.220, 0.032,0.156,
    0.021,0.156, 0.011,0.165, 0.007,0.176,
  ],
  [ // arabia + levant
    0.514,0.421, 0.518,0.442, 0.512,0.445, 0.512,0.454, 0.525,0.459, 0.537,0.446, 0.537,0.459,
    0.553,0.461, 0.549,0.469, 0.537,0.465, 0.531,0.473, 0.531,0.461, 0.520,0.469, 0.523,0.476,
    0.531,0.472, 0.537,0.489, 0.549,0.489, 0.549,0.469, 0.553,0.463, 0.568,0.482, 0.582,0.456,
    0.564,0.432, 0.557,0.434, 0.560,0.418, 0.543,0.385, 0.547,0.372, 0.541,0.372, 0.543,0.383,
    0.539,0.393, 0.533,0.390, 0.532,0.403, 0.543,0.413, 0.539,0.427, 0.549,0.435, 0.541,0.435,
    0.537,0.448, 0.535,0.415,
  ],
  [ // patagonia
    0.297,0.845, 0.267,0.849, 0.262,0.858, 0.258,0.886, 0.258,0.908, 0.262,0.904, 0.264,0.942,
    0.260,0.956, 0.258,0.932, 0.256,0.955, 0.260,0.966, 0.258,0.976, 0.263,0.987, 0.267,0.985,
    0.276,0.958, 0.273,0.945, 0.279,0.930, 0.279,0.911, 0.286,0.907, 0.289,0.894, 0.295,0.894,
    0.301,0.887, 0.297,0.869,
  ],
  [ // arctic islands
    0.291,0.013, 0.276,0.004, 0.251,0.004, 0.213,0.021, 0.215,0.032, 0.225,0.037, 0.244,0.032,
    0.236,0.051, 0.225,0.042, 0.225,0.049, 0.234,0.063, 0.223,0.063, 0.229,0.077, 0.221,0.076,
    0.217,0.087, 0.240,0.092, 0.244,0.085, 0.236,0.080, 0.254,0.066, 0.251,0.054, 0.282,0.028,
    0.279,0.020,
  ],
  [ // hudson bay coast
    0.214,0.142, 0.223,0.159, 0.245,0.156, 0.260,0.176, 0.263,0.193, 0.271,0.197, 0.268,0.203,
    0.261,0.199, 0.257,0.210, 0.249,0.214, 0.260,0.214, 0.278,0.237, 0.275,0.223, 0.283,0.224,
    0.275,0.201, 0.280,0.197, 0.287,0.208, 0.292,0.197, 0.277,0.182, 0.272,0.156, 0.264,0.155,
    0.250,0.135, 0.241,0.139, 0.235,0.124, 0.227,0.131, 0.226,0.148, 0.222,0.138, 0.224,0.123,
  ],
  [ // west africa
    0.492,0.456, 0.480,0.445, 0.468,0.461, 0.461,0.461, 0.466,0.463, 0.471,0.477, 0.470,0.494,
    0.478,0.501, 0.449,0.499, 0.450,0.504, 0.456,0.510, 0.454,0.518, 0.455,0.539, 0.468,0.535,
    0.472,0.515, 0.483,0.493, 0.491,0.490, 0.489,0.472, 0.493,0.461,
  ],
  [ // gulf of guinea
    0.455,0.614, 0.451,0.610, 0.448,0.613, 0.442,0.613, 0.440,0.615, 0.440,0.625, 0.438,0.632,
    0.438,0.638, 0.436,0.642, 0.430,0.642, 0.428,0.645, 0.428,0.663, 0.430,0.666, 0.434,0.666,
    0.440,0.675, 0.443,0.670, 0.454,0.670, 0.455,0.668, 0.455,0.645, 0.453,0.638, 0.453,0.620,
  ],
  [ // canadian isles
    0.180,0.161, 0.173,0.151, 0.170,0.139, 0.171,0.123, 0.165,0.134, 0.164,0.145, 0.160,0.137,
    0.156,0.139, 0.152,0.132, 0.147,0.139, 0.143,0.131, 0.141,0.134, 0.141,0.152, 0.151,0.156,
    0.151,0.161, 0.142,0.161, 0.141,0.165, 0.150,0.176, 0.157,0.176, 0.166,0.168, 0.178,0.172,
    0.177,0.165, 0.180,0.163,
  ],
  [ // arabian peninsula
    0.547,0.523, 0.547,0.532, 0.553,0.539, 0.581,0.539, 0.584,0.535, 0.593,0.535, 0.594,0.532,
    0.594,0.520, 0.589,0.511, 0.586,0.511, 0.583,0.515, 0.576,0.515, 0.573,0.511, 0.564,0.507,
    0.557,0.515, 0.555,0.515,
  ],
  [ // iberia
    0.501,0.414, 0.498,0.423, 0.495,0.423, 0.494,0.425, 0.494,0.431, 0.492,0.438, 0.492,0.448,
    0.497,0.456, 0.496,0.461, 0.497,0.463, 0.506,0.463, 0.508,0.459, 0.507,0.452, 0.508,0.444,
    0.512,0.438, 0.513,0.434, 0.511,0.427, 0.511,0.421, 0.509,0.415, 0.505,0.418,
  ],
];

function worldMap(slide) {
  WORLD_MAP.forEach(outline => {
    const points = [];
    for (let i = 0; i < outline.length; i += 2) {
      points.push({ x: MAP_BOX.x + outline[i] * MAP_BOX.w, y: MAP_BOX.y + outline[i + 1] * MAP_BOX.h });
    }
    points.push({ close: true });
    slide.addShape('custGeom', {
      x: 0, y: 0, w: 13.333, h: 7.5,
      points,
      fill: { color: GREY_85 },
      line: { type: 'none' },
    });
  });
}

/* ------------------------------------------------------------------ */
/* build                                                               */
/* ------------------------------------------------------------------ */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK_16x9', width: 13.333, height: 7.5 });
  pptx.layout = 'DECK_16x9';
  pptx.theme = { headFontFace: SERIF, bodyFontFace: SANS };

  slideBuilders.forEach((buildSlide, i) => {
    const slide = pptx.addSlide();
    slide.background = { color: i === slideBuilders.length - 1 ? NEAR_BLACK : BLACK };
    buildSlide(slide);
  });

  return pptx.writeFile({
    fileName: path.join(__dirname, '11778d31-df3a-4569-96a6-b81b5f7db5c4_grok_final.pptx'),
  });
}

build().then(f => console.log('wrote', f)).catch(err => {
  console.error(err);
  process.exit(1);
});
