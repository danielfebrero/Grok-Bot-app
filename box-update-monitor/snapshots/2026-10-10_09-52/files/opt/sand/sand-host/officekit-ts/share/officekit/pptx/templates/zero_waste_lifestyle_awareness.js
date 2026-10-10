/**
 * "Zero Waste" deck — recreated with pptxgenjs.
 *
 * Slide size 13.333 x 7.5 in (16:9). Deep teal background comes from the master.
 * Photographs in the original are empty picture placeholders; they are drawn here
 * as flat light panels labelled "[image]".
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const TEAL = '18635C'; // slide background / dark accents
const TEAL_SOFT = '1A6C64'; // decorative circles on the teal background
const GREEN = 'B7CD6F'; // lime accent (bars, cards, circles)
const WHITE = 'FFFFFF';
const BLACK = '000000';
const INK = '262626'; // tx1 lightened 15% — body copy on light panels
const PANEL = 'FFFFFF'; // photo panels: white with a faint blue pinstripe
const PANEL_RULE = 'C6D5EC';
const PANEL_INK = 'AFC3E6'; // "[image]" caption on those panels

const HEAD = 'Poppins'; // display face
const BODY = 'Roboto'; // copy face

/* -------------------------------------------------------------- copy blocks */

const LOREM =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed non odio justo. ' +
  'Fusce purus ligula, pellentesque in nunc sit amet, congue auctor justo. ' +
  'Aenean eu augue tellus. Proin vestibulum ex non eros ultrices tristique. ' +
  'Etiam ut risus tellus. Donec nec ex arcu. Sed lacinia bibendum magna, eu egestas massa.\u00A0';

const LOREM_SHORT =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed non odio justo. ' +
  'Fusce purus ligula, pellentesque in nunc sit amet, congue auctor justo. ' +
  'Aenean eu augue tellus. ';

const LOREM_CHART = LOREM.replace(' massa.\u00A0', '.');

/* ------------------------------------------------------------------ helpers */

/** Positive modulo. */
function mod(a, n) {
  return ((a % n) + n) % n;
}

/** Solid rectangle. */
function rect(slide, x, y, w, h, fill) {
  slide.addShape('rect', { x, y, w, h, fill: { color: fill }, line: { type: 'none' } });
}

/** Circle (the deck's decorative blobs — many bleed off the slide). */
function circle(slide, x, y, d, fill) {
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color: fill }, line: { type: 'none' } });
}

/**
 * Rounded card. `adj` is the OOXML roundRect adjust value (0-50000);
 * pptxgenjs wants the corner radius in inches instead.
 */
function card(slide, x, y, w, h, fill, adj) {
  slide.addShape('roundRect', {
    x, y, w, h,
    fill: { color: fill },
    line: { type: 'none' },
    rectRadius: (adj / 100000) * Math.min(w, h),
  });
}

/** Display headline; `lines` is one string per paragraph. Never wraps. */
function headline(slide, lines, box, opts) {
  const o = opts || {};
  slide.addText(
    lines.map((line, i) => ({ text: line, options: { breakLine: i < lines.length - 1 } })),
    {
      x: box[0], y: box[1], w: box[2], h: box[3],
      fontFace: HEAD, fontSize: o.size || 60, bold: true,
      color: o.color || WHITE,
      align: o.align || 'left', valign: 'top', wrap: false,
    }
  );
}

/** Paragraph copy. */
function copy(slide, text, box, opts) {
  const o = opts || {};
  slide.addText(text, {
    x: box[0], y: box[1], w: box[2], h: box[3],
    fontFace: BODY, fontSize: o.size || 14, bold: false, italic: false,
    color: o.color || INK,
    align: o.align || 'justify', valign: 'top',
  });
}

/**
 * Panel standing in for a photo placeholder. The template fills these with a
 * 5% diagonal pinstripe, redrawn here as evenly spaced 45° rules clipped to
 * the panel rectangle.
 */
const RULE_STEP = 0.1389; // gap between pinstripes, inches
const RULE_PHASE = 0.07; // pattern is anchored to the slide, not to the panel

function imagePanel(slide, x, y, w, h) {
  rect(slide, x, y, w, h, PANEL);
  const start = mod(RULE_PHASE - (x + y), RULE_STEP);
  for (let c = start; c < w + h; c += RULE_STEP) {
    const xEnd = Math.min(c, w); //   the rule runs from (xStart, yEnd)
    const yStart = c - xEnd; //       up-right to (xEnd, yStart)
    const yEnd = Math.min(c, h);
    const xStart = c - yEnd;
    slide.addShape('line', {
      x: x + xStart, y: y + yStart, w: xEnd - xStart, h: yEnd - yStart,
      flipV: true, line: { color: PANEL_RULE, width: 0.75 },
    });
  }
  slide.addText('[image]', {
    x, y, w, h,
    fontFace: BODY, fontSize: 12, color: PANEL_INK,
    align: 'center', valign: 'middle',
  });
}

/** Big number + caption used on the stat cards. */
function stat(slide, value, valueBox, captionBox, size) {
  slide.addText(value, {
    x: valueBox[0], y: valueBox[1], w: valueBox[2], h: valueBox[3],
    fontFace: HEAD, fontSize: size, bold: true, color: TEAL,
    align: valueBox[4] || 'left', valign: 'top', wrap: false,
  });
  copy(slide, 'Your Text', captionBox, { align: 'center' });
}

/* --------------------------------------------- contact icons (slide 14) ---
 * Small line-art marks, each drawn inside a 0.5 x 0.5 in box at (x, y).
 */

function phoneIcon(slide, x, y) {
  slide.addShape('blockArc', {
    x: x + 0.04, y: y + 0.11, w: 0.42, h: 0.30,
    angleRange: [180, 360], arcThicknessRatio: 0.3,
    fill: { color: TEAL }, line: { type: 'none' },
  });
  slide.addShape('roundRect', {
    x: x + 0.11, y: y + 0.23, w: 0.28, h: 0.19, rectRadius: 0.035,
    fill: { color: TEAL }, line: { type: 'none' },
  });
}

function laptopIcon(slide, x, y) {
  slide.addShape('rect', {
    x: x + 0.08, y: y + 0.10, w: 0.34, h: 0.24,
    fill: { type: 'none' }, line: { color: TEAL, width: 1.5 },
  });
  slide.addShape('ellipse', {
    x: x + 0.18, y: y + 0.15, w: 0.14, h: 0.14,
    fill: { type: 'none' }, line: { color: TEAL, width: 1 },
  });
  slide.addShape('trapezoid', {
    x: x + 0.02, y: y + 0.34, w: 0.46, h: 0.06,
    fill: { color: TEAL }, line: { type: 'none' },
  });
}

function mailIcon(slide, x, y) {
  slide.addShape('rect', {
    x: x + 0.05, y: y + 0.13, w: 0.36, h: 0.24,
    fill: { type: 'none' }, line: { color: TEAL, width: 1.5 },
  });
  slide.addShape('line', { x: x + 0.05, y: y + 0.13, w: 0.18, h: 0.12, line: { color: TEAL, width: 1.5 } });
  slide.addShape('line', { x: x + 0.23, y: y + 0.13, w: 0.18, h: 0.12, flipH: true, line: { color: TEAL, width: 1.5 } });
}

/* ------------------------------------------------------------- slide bodies */

/** Leaf mark of the "Your Logo" lockup (slides 1 and 15). */
function logoLockup(slide) {
  const x = 7.879, y = 2.114, d = 0.474; // blade box
  slide.addShape('custGeom', {
    x, y, w: d, h: d, fill: { color: GREEN }, line: { type: 'none' },
    points: [
      { x: 0.06 * d, y: 0.94 * d },
      { curve: { type: 'quadratic', x1: 0.02 * d, y1: 0.20 * d }, x: 0.94 * d, y: 0.06 * d },
      { curve: { type: 'quadratic', x1: 0.98 * d, y1: 0.80 * d }, x: 0.06 * d, y: 0.94 * d },
      { close: true },
    ],
  });
  slide.addShape('line', {
    x: x + 0.10 * d, y: y + 0.42 * d, w: 0.48 * d, h: 0.48 * d, flipV: true,
    line: { color: TEAL, width: 1.25 },
  });
  slide.addText('Your Logo', {
    x: 8.358, y: 2.211, w: 1.224, h: 0.37,
    fontFace: BODY, fontSize: 16, color: WHITE, valign: 'top', wrap: false,
  });
}

function slide01(slide) {
  headline(slide, ['Zero'], [7.714, 2.513, 3.268, 1.717], { size: 96 });
  headline(slide, ['Waste'], [7.714, 3.629, 4.66, 1.717], { size: 96 });
  rect(slide, 7.714, 5.27, 2.808, 0.212, GREEN);
  logoLockup(slide);
  circle(slide, 10.12, -1.717, 4.66, GREEN);
  circle(slide, 3.998, 5.779, 4.66, TEAL_SOFT);
  circle(slide, -2.514, 4.167, 4.66, GREEN);
}

function slide02(slide) {
  circle(slide, 5.122, -3.039, 4.66, TEAL_SOFT);
  imagePanel(slide, 0, -0.011, 7.562, 7.511);
  rect(slide, 4.241, 2.014, 9.093, 3.429, GREEN);
  headline(slide, ['Introduction'], [5.234, 2.648, 6.674, 1.313], { size: 72, color: TEAL_SOFT });
  copy(slide, LOREM, [5.234, 3.826, 7.057, 1.043], { color: BLACK });
  circle(slide, 9.782, 5.785, 4.66, TEAL_SOFT);
}

function slide03(slide) {
  rect(slide, 0, 0, 6.667, 3.75, GREEN);
  headline(slide, ['Problem'], [0.55, 0.813, 3.854, 1.111], { color: TEAL });
  copy(slide, LOREM, [0.55, 1.865, 5.855, 1.279]);
  card(slide, 6.928, 4.655, 2.869, 1.904, GREEN, 20825);
  card(slide, 10.005, 4.655, 2.869, 1.904, GREEN, 20825);
  stat(slide, '999+', [7.222, 4.925, 2.281, 1.111], [7.557, 5.867, 1.612, 0.337], 60);
  stat(slide, '1000+', [10.183, 4.859, 2.691, 1.111], [10.655, 5.818, 1.612, 0.337], 60);
}

function slide04(slide) {
  circle(slide, 9.063, -2.087, 5.553, GREEN);
  circle(slide, -0.739, 4.487, 5.553, TEAL_SOFT);
  headline(slide, ['What is', 'Zero Waste?'], [0.675, 1.942, 5.537, 2.121]);
  copy(slide, LOREM, [0.675, 4.063, 5.855, 1.279], { color: WHITE });
}

function slide05(slide) {
  circle(slide, -3.699, -0.383, 8.698, TEAL_SOFT);
  headline(slide, ['Benefit of', 'Zero Waste'], [0.74, 1.846, 5.084, 2.121]);
  copy(slide, LOREM, [0.74, 3.966, 4.601, 1.75], { color: WHITE });

  [
    { label: 'Benefit 01', y: 1.667, labelW: 2.113 },
    { label: 'Benefit 02', y: 3.812, labelW: 2.19 },
  ].forEach(b => {
    card(slide, 8.308, b.y, 4.654, 1.904, GREEN, 22819);
    headline(slide, [b.label], [8.753, b.y + 0.241, b.labelW, 0.572], { size: 28, color: TEAL_SOFT });
    copy(slide, LOREM_SHORT, [8.753, b.y + 0.779, 4.019, 0.909], { size: 12, color: BLACK });
  });
}

function slide06(slide) {
  circle(slide, 3.213, -2.468, 5.553, TEAL_SOFT);
  imagePanel(slide, 6.418, -0.011, 6.916, 7.511);
  circle(slide, -1.145, 4.724, 5.553, TEAL_SOFT);
  rect(slide, 1.464, 4.112, 6.523, 1.9, GREEN);
  headline(slide, ['Zero Waste', 'At Home'], [0.905, 1.815, 5.084, 2.121]);
  copy(slide, LOREM, [1.798, 4.431, 5.855, 1.279]);
}

function slide07(slide) {
  imagePanel(slide, 8.249, -0.011, 5.084, 7.511);
  headline(slide, ['Zero Waste', 'In Business'], [2.424, 1.718, 5.084, 2.121]);
  rect(slide, 0, 0, 1.734, 7.5, GREEN);
  rect(slide, 3.3, 4.036, 6.523, 1.9, GREEN);
  copy(slide, LOREM, [3.634, 4.355, 5.855, 1.279]);
}

function slide08(slide) {
  circle(slide, -2.825, -0.497, 8.698, TEAL_SOFT);
  headline(slide, ['Zero Waste', 'Shopping'], [0.88, 1.985, 5.084, 2.121]);
  copy(slide, LOREM, [0.88, 4.105, 4.994, 1.515], { color: WHITE });
  circle(slide, 11.74, 1.747, 4.005, GREEN);
}

function slide09(slide) {
  headline(slide, ['Composting'], [1.088, 0.666, 5.579, 1.111]);
  copy(slide, LOREM, [6.842, 5.764, 5.855, 1.279], { color: WHITE });
  circle(slide, 7.144, -3.932, 5.553, TEAL_SOFT);
  circle(slide, 0.637, 5.822, 5.553, TEAL_SOFT);
  imagePanel(slide, 0, 2.0, 13.333, 3.443);
}

function slide10(slide) {
  headline(slide, ['The Role', 'Of Recycling'], [1.082, 0.863, 5.598, 2.121]);
  copy(slide, LOREM, [1.082, 2.984, 5.855, 1.279], { color: WHITE });
  circle(slide, 1.193, 4.506, 2.158, GREEN);
  circle(slide, 3.549, 4.506, 2.158, GREEN);
  stat(slide, '85%', [1.44, 5.019, 1.662, 0.909], [1.466, 5.759, 1.612, 0.337], 48);
  stat(slide, '90%', [3.827, 5.019, 1.639, 0.909], [3.852, 5.759, 1.612, 0.337], 48);
}

function slide11(slide) {
  headline(slide, ['Break Slide'], [0.981, 3.34, 7.253, 1.582], { size: 88 });
  circle(slide, -1.414, -3.419, 6.49, TEAL_SOFT);
  circle(slide, 2.601, 5.191, 6.49, TEAL_SOFT);
  circle(slide, 5.975, 0.869, 1.383, GREEN);
  circle(slide, 1.295, 5.191, 0.936, GREEN);
  imagePanel(slide, 8.975, -0.011, 4.359, 7.511);
}

function slide12(slide) {
  circle(slide, 7.031, -1.77, 8.151, TEAL_SOFT);
  headline(slide, ['Circular', 'Economy'], [6.98, 1.823, 4.229, 2.121]);
  copy(slide, LOREM, [6.98, 4.037, 5.855, 1.279], { color: WHITE });
}

function slide13(slide, pptx) {
  headline(slide, ['Our Data'], [0.912, 1.282, 4.066, 1.111]);
  copy(slide, LOREM_CHART, [0.912, 2.333, 4.938, 1.515], { color: WHITE });

  const cats = ['Category 1', 'Category 2', 'Category 3', 'Category 4'];
  slide.addChart(
    [
      {
        type: pptx.ChartType.bar,
        data: [
          { name: 'Series 1', labels: cats, values: [4.3, 2.5, 3.5, 4.5] },
          { name: 'Series 2', labels: cats, values: [2.4, 4.4, 1.8, 2.8] },
        ],
        options: { chartColors: [GREEN, WHITE], barGapWidthPct: 219, barOverlapPct: -27 },
      },
      {
        type: pptx.ChartType.line,
        data: [{ name: 'Series 3', labels: cats, values: [2, 2, 3, 5] }],
        options: { chartColors: ['A5A5A5'], lineSize: 2.25, lineDataSymbol: 'none' },
      },
    ],
    {
      x: 6.4, y: 0.7, w: 6.47, h: 6.362,
      showTitle: true, title: 'Chart Title', titleColor: WHITE, titleFontFace: 'Calibri', titleFontSize: 18.62,
      catAxisLabelColor: WHITE, catAxisLabelFontFace: 'Calibri', catAxisLabelFontSize: 11.97,
      catAxisLineColor: 'D9D9D9', catAxisMajorTickMark: 'none', catAxisMinorTickMark: 'none',
      valAxisLabelColor: WHITE, valAxisLabelFontFace: 'Calibri', valAxisLabelFontSize: 11.97,
      valAxisLineShow: false, valAxisMajorTickMark: 'none', valAxisMinorTickMark: 'none',
      valGridLine: { color: 'D9D9D9', size: 1 }, catGridLine: { style: 'none' },
      showLegend: true, legendPos: 'b', legendColor: WHITE, legendFontFace: 'Calibri', legendFontSize: 11.97,
    }
  );

  card(slide, 0.912, 4.271, 4.938, 1.904, GREEN, 20825);
  slide.addText('1500++', {
    x: 1.206, y: 4.541, w: 3.926, h: 1.111,
    fontFace: HEAD, fontSize: 60, bold: true, color: TEAL, align: 'center', valign: 'top',
  });
  copy(slide, 'Lorem ipsum dolor sit amet', [1.994, 5.558, 2.774, 0.337], { align: 'center' });
}

function slide14(slide) {
  circle(slide, -0.953, 3.75, 6.497, TEAL_SOFT);
  rect(slide, 7.072, 3.839, 6.262, 2.105, GREEN);
  headline(slide, ['Keep ', 'In Touch'], [1.143, 0.854, 3.876, 2.121]);

  phoneIcon(slide, 8.104, 4.2);
  laptopIcon(slide, 8.104, 4.655);
  mailIcon(slide, 8.127, 5.155);

  [
    { text: '+123-4567-890', y: 4.247, w: 3.898 },
    { text: 'www.yourwebsite.com', y: 4.707, w: 4.376 },
    { text: 'name@yourwebsite.com', y: 5.181, w: 4.647 },
  ].forEach(row => {
    slide.addText(row.text, {
      x: 8.686, y: row.y, w: row.w, h: 0.404,
      fontFace: BODY, fontSize: 18, bold: false, color: INK, valign: 'top',
    });
  });

  circle(slide, 7.387, -3.113, 6.497, TEAL_SOFT);
}

function slide15(slide) {
  headline(slide, ['Thank'], [7.714, 2.513, 4.56, 1.717], { size: 96 });
  headline(slide, ['You!'], [7.714, 3.629, 3.398, 1.717], { size: 96 });
  rect(slide, 7.714, 5.27, 2.808, 0.212, GREEN);
  logoLockup(slide);
  circle(slide, 9.944, -2.147, 4.66, GREEN);
  circle(slide, 3.998, 5.779, 4.66, TEAL_SOFT);
  circle(slide, -2.514, 4.167, 4.66, GREEN);
}

/* ---------------------------------------------------------------- assembly */

const SLIDES = [
  slide01, slide02, slide03, slide04, slide05,
  slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15,
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK', width: 13.333333, height: 7.5 });
  pptx.layout = 'DECK';
  pptx.title = 'Zero Waste';

  SLIDES.forEach(builder => {
    const slide = pptx.addSlide();
    slide.background = { color: TEAL };
    builder(slide, pptx);
  });

  return pptx.writeFile({ fileName: path.join(__dirname, '0faee6da-d16c-4cc5-a14c-478ff517c0e8_grok_final.pptx') });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
