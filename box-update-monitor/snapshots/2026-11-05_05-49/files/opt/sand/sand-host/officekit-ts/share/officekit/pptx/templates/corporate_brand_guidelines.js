/**
 * "BRAND GUIDELINES" deck - rebuilt with PptxGenJS.
 * 20 slides, 13.333in x 7.5in (16:9).
 *
 *   node 0a9b8a23-22a6-4a07-ab14-9b6837b956d3_grok_final.js
 *      -> 0a9b8a23-22a6-4a07-ab14-9b6837b956d3_grok_final.pptx (next to this file)
 */

const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ *
 * Theme
 * ------------------------------------------------------------------ */
const DARK  = '30110D';   // accent2 - deep brown, fills the divider slides
const TAN   = 'B5AFA7';   // accent1 - warm grey
const SAND  = 'A19180';   // accent3
const STONE = 'A6A099';   // accent4
const PAPER = 'E1DFDC';   // accent1 @ lumMod 40 / lumOff 60 - light slide background
const WHITE = 'FFFFFF';
const BLACK = '000000';
const BODY  = '262626';   // small paragraph copy
const GREY  = 'BFBFBF';

const F     = 'Poppins';   // theme major + minor latin font

const SW = 13.333, SH = 7.5;            // slide size (inches)

/**
 * addText with the deck's defaults (Poppins, top-anchored like PowerPoint's
 * own text boxes). Every call below therefore only states what differs.
 */
function txt(s, content, opts) {
  s.addText(content, Object.assign({ fontFace: F, valign: 'top', color: BLACK }, opts));
}

/* ------------------------------------------------------------------ *
 * Shared chrome
 * ------------------------------------------------------------------ */

// x offsets (inches, from the group origin) of the 99 bars making up the
// decorative barcode in the top-left corner of every slide.
const BARCODE = [
  0.0000, 0.0266, 0.0550, 0.0677, 0.0735, 0.1100, 0.1227, 0.1250, 0.1568, 0.1649, 0.1800, 0.2031,
  0.2037, 0.2118, 0.2459, 0.2581, 0.2587, 0.2604, 0.2622, 0.3009, 0.3154, 0.3171, 0.3247, 0.3715,
  0.3796, 0.4097, 0.4236, 0.4363, 0.4705, 0.4774, 0.4832, 0.5069, 0.5347, 0.5538, 0.6007, 0.6128,
  0.6134, 0.6429, 0.6557, 0.6574, 0.6678, 0.6701, 0.6719, 0.7106, 0.7173, 0.7251, 0.7269, 0.7344,
  0.7439, 0.7723, 0.7812, 0.7850, 0.7894, 0.7908, 0.8272, 0.8333, 0.8400, 0.8423, 0.8741, 0.8973,
  0.9204, 0.9210, 0.9291, 0.9632, 0.9754, 0.9760, 0.9777, 0.9794, 1.0182, 1.0327, 1.0344, 1.0419,
  1.0888, 1.0969, 1.1270, 1.1409, 1.1536, 1.1878, 1.1947, 1.2242, 1.2520, 1.2711, 1.2838, 1.3180,
  1.3301, 1.3307, 1.3388, 1.3602, 1.3730, 1.3747, 1.3851, 1.3857, 1.3874, 1.3892, 1.4279, 1.4424,
  1.4441, 1.4517, 1.4985,
];

/**
 * Header strip shared by all 20 slides: barcode + year on the left, the wide
 * colour bar carrying "Created / V1-01 / Page nn" on the right.
 */
function addHeader(s, opts) {
  const { page, barColor, ink, cornerInk } = opts;

  BARCODE.forEach((dx) => {
    s.addShape('line', {
      x: 0.407 + dx, y: 0.298, w: 0, h: 0.327, line: { color: cornerInk, width: 0.5 },
    });
  });
  txt(s, '2028/2030', {
    x: 0.294, y: 0.662, w: 1.347, h: 0.269,
    fontSize: 10, charSpacing: 3, color: cornerInk, lineSpacing: 12,
  });

  s.addShape('rect', { x: 6.247, y: 0.292, w: 7.087, h: 0.361, fill: { color: barColor } });
  txt(s, 'Created : 25 July 2025', { x: 6.736, y: 0.323, w: 3.152, h: 0.303, fontSize: 12, color: ink });
  txt(s, 'V1-01', {
    x: 9.743, y: 0.328, w: 1.195, h: 0.303, fontSize: 12, charSpacing: 3, color: ink, align: 'center',
  });
  txt(s, `Page ${page}`, {
    x: 11.313, y: 0.338, w: 1.335, h: 0.269, fontSize: 10, charSpacing: 3, color: ink, align: 'right',
  });
}

/** Kicker + big headline, present on every slide. */
function addHeadline(s, opts) {
  const { kicker, kickerSize, kickerColor, kickerW, title, titleSize, titleColor, titleW, titleH } = opts;
  txt(s, kicker, {
    x: 1.475, y: 1.375, w: kickerW, h: kickerSize >= 28 ? 0.59 : 0.558,
    fontSize: kickerSize, color: kickerColor, lineSpacing: 36,
  });
  txt(s, title, {
    x: 1.442, y: 1.749, w: titleW, h: titleH,
    fontSize: titleSize, bold: true, color: titleColor,
  });
}

/** The "CONTRARY TO POPULAR BELIEF..." block on the lower left of content slides. */
function addSideNote(s) {
  txt(s, [
    { text: 'CONTRARY TO POPULAR BELIEF, LOREM IPSUM IS ALOREM', options: { breakLine: true } },
    { text: 'IPSUM PASSAGE, ', options: { underline: { style: 'sng' } } },
  ], {
    x: 1.442, y: 5.015, w: 3.839, h: 1.005,
    fontSize: 14.67, bold: true, color: DARK, lineSpacingMultiple: 1.2,
  });
  txt(s, LOREM_NOTE, {
    x: 1.442, y: 6.027, w: 4.686, h: 0.662, fontSize: 9, color: BODY, lineSpacingMultiple: 1.2,
  });
}

const LOREM_NOTE =
  'contrary to popular belief, lorem ipsum is not simply obscure latin connecter, ' +
  'alarm ipsum passage, and going through the cites of the popular belief, lorem ipsum is not simply';

/**
 * Stand-in for one of the deck's picture frames: a small light card centred in
 * the frame, plus the frame's prompt caption where the reference shows one.
 */
function addImagePlaceholder(s, x, y, w, h, caption) {
  const cw = Math.min(0.87, w * 0.8), ch = Math.min(0.71, h * 0.8);
  s.addShape('rect', {
    x: x + (w - cw) / 2, y: y + (h - ch) / 2, w: cw, h: ch,
    fill: { color: 'F7F7F7' }, line: { color: '595959', width: 0.75 },
  });
  txt(s, '[image]', {
    x: x + (w - cw) / 2, y: y + (h - ch) / 2, w: cw, h: ch,
    fontSize: 8, color: '595959', align: 'center', valign: 'middle',
  });
  if (caption) {
    txt(s, caption, { x, y: y + 0.06, w, h: 0.28, fontSize: 12, align: 'center' });
  }
}

/**
 * The "C" monogram: a blackletter capital traced as a closed vector outline.
 * Rows are normalised (0..1) coordinates inside the shape's bounding box -
 * [x, y] draws a line, [x, y, cx1, cy1, cx2, cy2] a cubic bezier.
 */
const MONOGRAM = [
  [0.5884,0.9843], [0.4248,0.895,0.5569,0.9643,0.4507,0.9064], [0.377,0.8809,0.4137,0.8902,0.3924,0.884],
  [0.3511,0.8695,0.3498,0.8757,0.3491,0.8754], [0.3626,0.727,0.3606,0.8428,0.3626,0.8187], [0.3252,0.62,0.3626,0.6169,0.3671,0.6303],
  [0.2459,0.5745,0.2918,0.6117,0.2816,0.6057], [0.2092,0.5421,0.2285,0.559,0.2118,0.5445], [0.2246,0.5402,0.2013,0.5354,0.2072,0.5347],
  [0.3452,0.4711,0.2796,0.5573,0.3268,0.5304], [0.3449,0.3372,0.3527,0.447,0.3524,0.3732], [0.3006,0.2536,0.3363,0.2981,0.3196,0.2659],
  [0.2662,0.2695,0.2937,0.249,0.2937,0.249], [0.1364,0.4535,0.2046,0.3153,0.1567,0.3832], [0.1295,0.5702,0.1276,0.484,0.1243,0.5383],
  [0.2239,0.8216,0.1403,0.6384,0.1852,0.7575], [0.2773,0.884,0.241,0.8499,0.2544,0.8654], [0.2967,0.9083,0.299,0.9012,0.3003,0.9026],
  [0.2908,0.9145,0.2944,0.9116,0.2918,0.9145], [0.1925,0.8602,0.2898,0.9145,0.2455,0.89], [0.0964,0.8061],
  [0.0738,0.7584], [0.0381,0.6834,0.0614,0.7322,0.0453,0.6984], [0.0066,0.5869,0.024,0.6531,0.0142,0.6229],
  [0.0017,0.4368,-0.0003,0.5528,-0.0029,0.4716], [0.2806,0.1228,0.022,0.2895,0.1135,0.1866], [0.5405,0.0687,0.3691,0.0889,0.4658,0.0687],
  [0.569,0.0687], [0.569,0.0341], [0.569,-0.0004],
  [0.5821,-0.0004], [0.5953,-0.0004], [0.5953,0.0339],
  [0.5953,0.0682], [0.629,0.0699], [0.7411,0.0911,0.6733,0.0722,0.705,0.0782],
  [0.965,0.2421,0.7896,0.1087,0.8116,0.1235], [0.9977,0.2752,0.9994,0.2688,1.0023,0.2717], [0.9813,0.2698,0.9928,0.2786,0.9915,0.2781],
  [0.9109,0.2159,0.941,0.2362,0.9305,0.2281], [0.7499,0.1647,0.8748,0.1933,0.8139,0.174], [0.611,0.1623,0.7195,0.1604,0.6385,0.159],
  [0.5953,0.1642], [0.5953,0.4906], [0.5953,0.8168],
  [0.5821,0.8168], [0.569,0.8168], [0.569,0.4918],
  [0.5667,0.1656,0.569,0.3129,0.568,0.1661], [0.4451,0.1892,0.5625,0.164,0.4786,0.1802], [0.3183,0.2436,0.3898,0.2042,0.3183,0.2347],
  [0.3586,0.2736,0.3183,0.2452,0.3363,0.2588], [0.4569,0.3736,0.4156,0.3119,0.4383,0.3346], [0.4904,0.7725,0.4822,0.4258,0.4904,0.5233],
  [0.4904,0.9131], [0.5009,0.9195], [0.5117,0.9259],
  [0.5356,0.9095], [0.6824,0.8547,0.5838,0.8769,0.6293,0.8597], [0.7735,0.8685,0.7165,0.8514,0.7339,0.854],
  [0.9305,0.9498,0.8214,0.8859,0.9024,0.9279], [0.9417,0.9691,0.9473,0.9629,0.9482,0.965], [0.9266,0.9636,0.9384,0.971,0.9348,0.9698],
  [0.8555,0.939,0.9119,0.9524,0.8863,0.9436], [0.7968,0.9379,0.8335,0.9357,0.8243,0.9357], [0.6621,0.976,0.7434,0.9421,0.7188,0.9491],
  [0.6126,0.9996], [0.5884,0.9843],
];

function addMonogram(s, x, y, w, h, color) {
  s.addShape('custGeom', {
    x, y, w, h,
    fill: { color },
    flipV: true,   // the source glyph is stored upside down and mirrored back
    points: MONOGRAM.map((p, i) => (
      i === 0 ? { x: w * p[0], y: h * p[1], moveTo: true }
      : p.length === 2 ? { x: w * p[0], y: h * p[1] }
      : { x: w * p[0], y: h * p[1],
          curve: { type: 'cubic', x1: w * p[2], y1: h * p[3], x2: w * p[4], y2: h * p[5] } }
    )).concat([{ close: true }]),
  });
}

/* ------------------------------------------------------------------ *
 * Slide skeletons
 * ------------------------------------------------------------------ */

/** Dark chapter-divider slide, with the oversized section number bleeding off. */
function dividerSlide(pptx, cfg) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: SW, h: SH, fill: { color: DARK } });

  txt(s, cfg.num, {
    x: cfg.numX, y: 3.812, w: cfg.numW, h: 5.01,
    fontSize: 300, charSpacing: 3, color: TAN, align: 'right', lineSpacing: 350,
  });
  s.addShape('line', { x: 0, y: 3.331, w: SW, h: 0, line: { color: WHITE, width: 0.25 } });

  addHeader(s, { page: cfg.page, barColor: TAN, ink: BLACK, cornerInk: WHITE });
  addHeadline(s, {
    kicker: 'BRAND', kickerSize: 28, kickerColor: WHITE, kickerW: 1.797,
    title: cfg.title, titleSize: 87, titleColor: WHITE, titleW: cfg.titleW, titleH: cfg.titleH,
  });

  txt(s, [
    { text: '© COPY ALL RIGHT. COMPLY ', options: { bold: true } },
    { text: 'BRAND GUIDELINE 2023' },
  ], { x: 1.442, y: 3.695, w: 4.951, h: 0.286, fontSize: 11, color: WHITE });

  return s;
}

/** The rule + two-column lorem block sitting under six of the divider slides. */
const LOREM_COL =
  'Dolor sit, elitdolor sit,erere dolorese iser eiusmodsjonsectetur adipiscing elit dolor ' +
  'PLACEHOLDER' +
  'elit dolor sitesp amet, consec glores volor eiusmodsjonsecteturer asser adipiscing elit dolor ' +
  'sitespis gooder doloresre gooi eadipiscis elit, sed do eiusmodsjonsectetur adipiscing elit ' +
  'dolor sitesp amet, ';

function addDividerBody(s) {
  s.addShape('line', { x: 1.522, y: 4.559, w: 1.852, h: 0, line: { color: WHITE, width: 1 } });
  txt(s, 'LOREM IPSUM DOLOR', {
    x: 1.428, y: 4.674, w: 1.695, h: 0.269, fontSize: 10, bold: true, color: WHITE,
  });
  s.addShape('line', { x: 1.522, y: 5.059, w: 1.852, h: 0, line: { color: WHITE, width: 1 } });
  txt(s, 'THERE ARE MANY VARIATIONS OF PASSAGES OF LOREM IPSUM AVAILABLE, BUT THE MAJORITY HAVE SUFFERED', {
    x: 1.425, y: 5.35, w: 2.467, h: 0.928, fontSize: 9, color: WHITE, lineSpacing: 15,
  });
  txt(s, LOREM_COL, {
    x: 4.021, y: 4.441, w: 2.886, h: 2.187, fontSize: 9, color: WHITE, lineSpacing: 15,
  });
}

/** Light content slide: paper background, short rule at mid-height, headline. */
function contentSlide(pptx, cfg) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: SW, h: SH, fill: { color: PAPER } });
  s.addShape('line', { x: 0, y: 3.331, w: 3.06, h: 0, line: { color: DARK, width: 0.25 } });

  addHeader(s, { page: cfg.page, barColor: DARK, ink: WHITE, cornerInk: DARK });
  addHeadline(s, {
    kicker: 'COMPLY BRAND', kickerSize: 20, kickerColor: DARK, kickerW: 3.794,
    title: cfg.title, titleSize: 60, titleColor: BLACK, titleW: cfg.titleW, titleH: 1.111,
  });
  addSideNote(s);
  return s;
}

/* ------------------------------------------------------------------ *
 * Per-slide builders
 * ------------------------------------------------------------------ */

// 01 - cover
function slide01(pptx) {
  const s = dividerSlide(pptx, {
    page: '01', num: '01', numX: 9.462, numW: 4.438,
    title: 'GUIDELINES', titleW: 8.151, titleH: 1.582,
  });

  txt(s, [
    { text: 'Lorem ipsum, or\u00A0lapsus\u00A0as it is ', options: { bold: true } },
    { text: 'sometimes known, is dummy text used in laying out print, graphic or web designs. ' +
            'PLACEHOLDER' },
  ], { x: 1.449, y: 4.978, w: 5.287, h: 0.757, fontSize: 10, color: WHITE, lineSpacing: 16 });

  [
    { x: 1.442, head: 'BRAND GUIDELINES', lines: ['Client Name', 'Version – 00.3'] },
    { x: 3.744, head: 'LAST UPDATE',      lines: ['June 22', '2028/2030 December'] },
  ].forEach((m) => {
    txt(s, m.head, {
      x: m.x, y: 6.211, w: 1.885, h: 0.371,
      fontFace: 'Inter', fontSize: 9, bold: true, color: WHITE, lineSpacingMultiple: 1.78,
    });
    txt(s, m.lines.join('\n'), {
      x: m.x + 0.011, y: 6.471, w: 1.379, h: 0.37,
      fontFace: 'Inter', fontSize: 8, color: WHITE, lineSpacingMultiple: 1.0,
    });
  });
}

// 02 - content list
function slide02(pptx) {
  const s = contentSlide(pptx, { page: '02', title: 'CONTENT LIST', titleW: 6.771 });

  const rows = [
    { label: ' 01. Brand',      pages: '01-02', bold: true,  x: 7.631, w: 1.672, size: 18 },
    { label: '02.  Logo',       pages: '03-04', bold: false, x: 8.052, w: 1.158, size: 14 },
    { label: ' 03. Colour',     pages: '05-06', bold: false, x: 8.052, w: 1.323, size: 14 },
    { label: '04. Typography',  pages: '07-08', bold: false, x: 8.052, w: 1.833, size: 14 },
    { label: '05. Application', pages: '09-10', bold: false, x: 8.052, w: 1.740, size: 14 },
    { label: '06. Gallery',     pages: '11-12', bold: false, x: 8.052, w: 1.404, size: 14 },
    { label: '07. Others',      pages: '13-14', bold: false, x: 8.052, w: 1.672, size: 14 },
  ];

  rows.forEach((r, i) => {
    const y = 3.454 + i * 0.4832;
    txt(s, r.label, { x: r.x, y, w: r.w, h: 0.337, fontSize: r.size, bold: r.bold });
    s.addShape('line', {
      x: 10.4, y: 3.622 + i * 0.4842, w: 0.456, h: 0,
      line: { color: DARK, width: 0.5, endArrowType: 'triangle' },
    });
    txt(s, r.pages, { x: 11.45, y, w: 1.022, h: 0.337, fontSize: 14, bold: r.bold });
  });
}

// 03 - welcomes (empty portrait picture frame at 6.347 / 3.679, 2.96 x 3.01in)
function slide03(pptx) {
  const s = contentSlide(pptx, { page: '03', title: 'WELCOMES', titleW: 6.771 });

  txt(s, 'ASHLYN ANDERSON', { x: 9.583, y: 5.079, w: 3.126, h: 0.438, fontSize: 20, bold: true });
  txt(s, 'YOUR TITLE HERE', { x: 9.604, y: 5.551, w: 2.601, h: 0.253, fontSize: 12, lineSpacing: 10 });
  s.addShape('ellipse', { x: 9.706, y: 6.022, w: 0.089, h: 0.089, fill: { color: DARK } });
  txt(s, [
    { text: 'LOREM IPSUM', options: { bold: true } },
    { text: '\u00A0IS SIMPLY DUMMY TEXT OF THE PRINTING AND TYPESETTING.' },
  ], { x: 9.599, y: 6.282, w: 3.36, h: 0.459, fontSize: 10, lineSpacing: 13 });
}

// 04 / 08 / 11 / 13 / 15 / 18 - chapter dividers with the two-column body
function chapterDivider(pptx, cfg) {
  addDividerBody(dividerSlide(pptx, Object.assign({ numX: 8.047, numW: 5.715 }, cfg)));
}

// 05 - primary logo
function slide05(pptx) {
  const s = contentSlide(pptx, { page: '05', title: 'PRIMARY LOGO', titleW: 6.771 });
  s.addShape('rect', { x: 6.724, y: 3.18, w: 5.528, h: 3.509, fill: { color: TAN } });
  addMonogram(s, 8.424, 3.471, 2.128, 2.927, DARK);
}

// 06 - logo mark construction grid
function slide06(pptx) {
  const s = contentSlide(pptx, { page: '06', title: 'LOGO MARK', titleW: 6.771 });
  const rule = { color: DARK, width: 0.5 };

  [3.221, 3.452, 4.06, 5.688, 6.311, 6.542].forEach((y) => {
    s.addShape('line', { x: 6.417, y, w: 6.238, h: 0, line: rule });
  });
  [7.996, 8.2, 11.038, 11.222].forEach((x) => {
    s.addShape('line', { x, y: 2.851, w: 0, h: 4.033, line: rule });
  });

  addMonogram(s, 8.944, 4.06, 1.184, 1.628, DARK);

  [
    { t: 'Brand',      x: 9.189,  y: 3.611, w: 0.859, align: 'center' },
    { t: 'Brand',      x: 7.076,  y: 4.695, w: 0.833, align: 'left' },
    { t: 'Guidelines', x: 11.307, y: 4.695, w: 1.267, align: 'left' },
    { t: 'Logo',       x: 9.189,  y: 5.820, w: 0.859, align: 'center' },
  ].forEach((l) => {
    txt(s, l.t, { x: l.x, y: l.y, w: l.w, h: 0.259, fontSize: 12, align: l.align });
  });
}

// 07 - logo usage swatches
function slide07(pptx) {
  const s = contentSlide(pptx, { page: '07', title: 'LOGO USED', titleW: 5.849 });

  [
    { x: 6.667, y: 2.930, fill: DARK,  line: null, mark: WHITE, label: WHITE },
    { x: 9.403, y: 2.930, fill: STONE, line: null, mark: DARK,  label: BLACK },
    { x: 6.667, y: 5.013, fill: null,  line: DARK, mark: BLACK, label: BLACK },
    { x: 9.403, y: 5.013, fill: null,  line: null, mark: DARK,  label: BLACK },
  ].forEach((c) => {
    if (c.fill || c.line) {
      s.addShape('rect', {
        x: c.x, y: c.y, w: 2.292, h: 1.613,
        fill: c.fill ? { color: c.fill } : { type: 'none' },
        line: c.line ? { color: c.line, width: 0.75 } : undefined,
      });
    }
    addMonogram(s, c.x + 0.791, c.y + 0.151, 0.709, 0.975, c.mark);
    txt(s, 'BRAND LOGO', {
      x: c.x + 0.405, y: c.y + 1.184, w: 1.481, h: 0.311,
      fontSize: 14, color: c.label, align: 'center', lineSpacing: 15,
    });
  });
}

// 09 - hero colour
function slide09(pptx) {
  const s = contentSlide(pptx, { page: '09', title: 'HERO COLOR', titleW: 5.849 });

  [
    { x: 7.627,  y: 2.314, w: 2.657, h: 2.278, fill: TAN,   line: null, hex: '#E85322', tx: 7.721,  ty: 2.411, hw: 1.095, ink: BLACK },
    { x: 10.284, y: 2.314, w: 2.256, h: 1.518, fill: PAPER, line: DARK, hex: '#112F1C', tx: 10.349, ty: 2.411, hw: 0.953, ink: DARK },
    { x: 10.284, y: 3.832, w: 2.256, h: 3.037, fill: BLACK, line: null, hex: '#CEE4C8', tx: 10.349, ty: 3.918, hw: 1.158, ink: WHITE },
    { x: 6.247,  y: 4.591, w: 4.037, h: 2.278, fill: DARK,  line: null, hex: '#C8E4E0', tx: 6.394,  ty: 4.717, hw: 1.240, ink: WHITE },
  ].forEach((c) => {
    s.addShape('rect', {
      x: c.x, y: c.y, w: c.w, h: c.h,
      fill: { color: c.fill },
      line: c.line ? { color: c.line, width: 0.75 } : undefined,
    });
    txt(s, c.hex, { x: c.tx, y: c.ty, w: c.hw, h: 0.337, fontSize: 14, color: c.ink });
  });
}

// 10 - colour palettes (three tint ramps)
function slide10(pptx) {
  const s = contentSlide(pptx, { page: '10', title: 'COLOR PALLETES', titleW: 8.979 });

  const columns = [
    { x: 7.520,  w: 1.638, base: DARK,  hex: '#1A1817', hexX: 7.607,  labelX: 7.607 },
    { x: 9.151,  w: 1.630, base: TAN,   hex: '#EFB090', hexX: 9.217,  labelX: 9.211, faintest: SAND },
    { x: 10.782, w: 1.630, base: BLACK, hex: '#EFB090', hexX: 10.847, labelX: 10.841 },
  ];
  const rows = [
    { y: 3.252, h: 1.226, transparency: 0 },
    { y: 4.477, h: 0.539, transparency: 20 },
    { y: 5.014, h: 0.539, transparency: 40 },
    { y: 5.552, h: 0.539, transparency: 60 },
    { y: 6.090, h: 0.539, transparency: 80 },
  ];

  columns.forEach((c) => {
    rows.forEach((r, i) => {
      const color = i === rows.length - 1 && c.faintest ? c.faintest : c.base;
      s.addShape('rect', { x: c.x, y: r.y, w: c.w, h: r.h, fill: { color, transparency: r.transparency } });
    });
    txt(s, 'COLOR', { x: c.labelX, y: 3.409, w: 1.414, h: 0.269, fontSize: 10, color: WHITE });
    txt(s, c.hex,    { x: c.hexX,   y: 3.589, w: 1.204, h: 0.236, fontSize: 8,  color: WHITE });
  });

  ['100%', '80%', '60%', '40%', '20%'].forEach((label, i) => {
    txt(s, label, { x: 6.548, y: 3.61 + i * 0.695, w: 1.153, h: 0.3, fontSize: 14 });
  });
}

// 12 - typography specimen
function slide12(pptx) {
  const s = contentSlide(pptx, { page: '12', title: 'TYPOGRAPHY USED', titleW: 8.979 });

  txt(s, 'POPPINS REGULAR ', { x: 7.125, y: 3.444, w: 3.378, h: 0.438, fontSize: 20, color: DARK });
  txt(s, 'A B C D E F G H I J K L M N O P Q R S \nT U V W X Y Z A B', {
    x: 7.125, y: 3.838, w: 6.034, h: 0.735, fontSize: 14, charSpacing: 3, lineSpacingMultiple: 1.5,
  });
  txt(s, 'POPPINS BOLD', { x: 7.125, y: 4.771, w: 2.873, h: 0.438, fontSize: 20, bold: true, color: DARK });
  txt(s, 'a b c d e f g h i j k l m n o p q r s \nt u v w x y z a b', {
    x: 7.125, y: 5.252, w: 6.034, h: 0.774, fontSize: 14, bold: true, charSpacing: 3, lineSpacingMultiple: 1.5,
  });
  txt(s, '1 2 3 4 5 6 7 8 9 0 ! @ # $ % ^ & * ( ) _ + < > ? / \\', {
    x: 7.125, y: 6.231, w: 6.669, h: 0.482, fontFace: 'Poppins SemiBold', fontSize: 14, lineSpacing: 31,
  });
}

// 14 - image gallery (four frames; only the first carries a prompt)
function slide14(pptx) {
  const s = contentSlide(pptx, { page: '14', title: 'IMAGE GALLERY', titleW: 8.979 });
  addImagePlaceholder(s, 1.595, 3.648, 1.586, 1.111);
}

// 16 - web apps: laptop mock-up with an empty screen
function slide16(pptx) {
  const s = contentSlide(pptx, { page: '16', title: 'WEB APPS', titleW: 8.979 });

  s.addShape('roundRect', { x: 6.894, y: 3.364, w: 4.955, h: 3.365, rectRadius: 0.06, fill: { color: '1D1D1B' } });
  s.addShape('rect',      { x: 6.945, y: 3.414, w: 4.852, h: 3.176, fill: { color: PAPER } });
  s.addShape('rect',      { x: 6.894, y: 6.591, w: 4.955, h: 0.138, fill: { color: '1D1D1B' } });
  s.addShape('roundRect', { x: 6.321, y: 6.725, w: 6.094, h: 0.121, rectRadius: 0.06, fill: { color: '575756' } });
  s.addShape('rect',      { x: 6.344, y: 6.768, w: 6.048, h: 0.078, fill: { color: '1D1D1B' } });
  s.addShape('rect',      { x: 8.788, y: 6.725, w: 1.160, h: 0.043, fill: { color: '646363' } });
}

// 17 - phone & tablet mock-ups
function slide17(pptx) {
  const s = contentSlide(pptx, { page: '17', title: 'PHONE & TAB APPS', titleW: 8.979 });

  s.addShape('roundRect', { x: 6.207, y: 3.527, w: 1.559, h: 3.306, rectRadius: 0.18, fill: { color: BLACK } });
  s.addShape('roundRect', { x: 6.243, y: 3.567, w: 1.486, h: 3.231, rectRadius: 0.16, fill: { color: PAPER } });
  s.addShape('roundRect', { x: 6.607, y: 3.567, w: 0.746, h: 0.093, rectRadius: 0.3, fill: { color: BLACK } });
  s.addShape('roundRect', { x: 6.800, y: 3.578, w: 0.372, h: 0.040, rectRadius: 0.5, fill: { color: '484747' } });
  addImagePlaceholder(s, 6.243, 3.567, 1.486, 3.231, 'Picture');

  s.addShape('roundRect', { x: 8.215, y: 3.527, w: 4.318, h: 3.306, rectRadius: 0.10, fill: { color: '231F20' } });
  s.addShape('roundRect', { x: 8.286, y: 3.593, w: 4.203, h: 3.205, rectRadius: 0.08, fill: { color: PAPER } });
  addImagePlaceholder(s, 8.286, 3.593, 4.203, 3.205, 'Picture');
}

// 19 - tag & t-shirt
function slide19(pptx) {
  const s = contentSlide(pptx, { page: '19', title: 'TAG & T-SHIRT', titleW: 8.979 });

  // t-shirt silhouette: right shoulder -> neckline -> left shoulder -> sleeve
  // -> body -> sleeve, traced clockwise inside a 4.513 x 4.344in box
  s.addShape('custGeom', {
    x: 8.241, y: 2.381, w: 4.513, h: 4.344, fill: { color: DARK },
    points: [
      { x: 3.666, y: 0.289, moveTo: true },
      { x: 2.739, y: 0.000, curve: { type: 'cubic', x1: 3.405, y1: 0.079, x2: 2.739, y2: 0.000 } },
      { x: 1.774, y: 0.000, curve: { type: 'cubic', x1: 2.094, y1: 0.122, x2: 1.774, y2: 0.000 } },
      { x: 0.847, y: 0.289, curve: { type: 'cubic', x1: 1.774, y1: 0.000, x2: 1.108, y2: 0.079 } },
      { x: 0.000, y: 1.171, curve: { type: 'cubic', x1: 0.847, y1: 0.289, x2: 0.664, y2: 0.394 } },
      { x: 0.475, y: 1.868 },
      { x: 0.847, y: 1.742 },
      { x: 0.847, y: 4.344 },
      { x: 3.665, y: 4.344 },
      { x: 3.665, y: 1.742 },
      { x: 4.037, y: 1.867 },
      { x: 4.512, y: 1.169 },
      { x: 3.666, y: 0.289, curve: { type: 'cubic', x1: 3.849, y1: 0.394, x2: 3.666, y2: 0.289 } },
      { close: true },
    ],
  });
  // collar: the lower half of a disc, in background colour, notched into the
  // neckline - hence a full-height ellipse box whose top half is clipped away
  s.addShape('chord', {
    x: 10.016, y: 1.904, w: 0.963, h: 0.962, angleRange: [0, 180], fill: { color: TAN },
  });
  addMonogram(s, 10.143, 3.535, 0.709, 0.975, WHITE);
  txt(s, 'EXAMPLES TO POPULAR BELIEF, LOREM IPSUM IS NOT SIMPLY', {
    x: 9.684, y: 4.596, w: 1.641, h: 0.551, fontSize: 8, color: WHITE, align: 'center', lineSpacing: 11,
  });
  txt(s, 'COMPLY', {
    x: 10.421, y: 6.251, w: 1.47, h: 0.37,
    fontSize: 16, bold: true, charSpacing: 2, color: WHITE, align: 'center',
  });

  // hang tag: looped string, rounded card, eyelet
  s.addShape('ellipse', {
    x: 7.008, y: 3.321, w: 0.316, h: 1.508, fill: { type: 'none' }, line: { color: BLACK, width: 2.5 },
  });
  s.addShape('roundRect', { x: 6.477, y: 4.485, w: 1.379, h: 2.24, rectRadius: 0.12, fill: { color: TAN } });
  s.addShape('ellipse',   { x: 7.008, y: 4.709, w: 0.239, h: 0.239, fill: { color: GREY } });
  addMonogram(s, 6.884, 5.147, 0.564, 0.776, DARK);
  txt(s, 'EXAMPLES TO POPULAR BELIEF, LOREM IPSUM', {
    x: 6.607, y: 6.045, w: 1.118, h: 0.557, fontSize: 8, align: 'center', lineSpacing: 11,
  });
}

// 20 - thank you
function slide20(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: SW, h: SH, fill: { color: PAPER } });
  txt(s, '08', {
    x: 8.129, y: 3.83, w: 5.715, h: 5.01,
    fontSize: 300, charSpacing: 3, color: DARK, align: 'right', lineSpacing: 350,
  });
  s.addShape('line', { x: 0, y: 3.331, w: SW, h: 0, line: { color: DARK, width: 0.25 } });

  addHeader(s, { page: '20', barColor: DARK, ink: WHITE, cornerInk: DARK });
  addHeadline(s, {
    kicker: 'WE\u2019VE SAY ', kickerSize: 28, kickerColor: DARK, kickerW: 3.192,
    title: 'THANK YOU', titleSize: 87, titleColor: BLACK, titleW: 8.151, titleH: 1.582,
  });

  [
    { x: 1.548, ruleX: 1.639, head: 'MEDIA BY :',    w: 1.502 },
    { x: 3.760, ruleX: 3.864, head: 'PREPARED BY :', w: 1.763 },
  ].forEach((c) => {
    txt(s, c.head, { x: c.x, y: 3.714, w: c.w, h: 0.337, fontSize: 14, bold: true });
    s.addShape('line', { x: c.ruleX, y: 4.094, w: 1.575, h: 0, line: { color: DARK, width: 0.5 } });
    txt(s, 'ADRIAN BRAND', { x: c.x, y: 4.151, w: 1.502, h: 0.269, fontSize: 10 });
  });

  txt(s, LOREM_NOTE, {
    x: 1.468, y: 4.759, w: 6.661, h: 0.48, fontSize: 9, color: BODY, lineSpacingMultiple: 1.2,
  });

  [
    { x: 1.468, w: 1.815, head: 'CONTACT', headW: 1.080,
      lines: ['PHONE : 789-87-6', 'FAX : 789-87-6', '-', 'MAIL : YOUR@MAIL.COM'] },
    { x: 4.075, w: 2.243, head: 'ADDRESS', headW: 1.022,
      lines: ['FLEXIRA. BRAND', '356 DOWNTOWN STREET BLVD', 'SCRANTON, PA', 'UNITED STATES'] },
  ].forEach((b) => {
    txt(s, b.head, { x: b.x, y: 5.575, w: b.headW, h: 0.303, fontSize: 12, bold: true });
    txt(s, b.lines.join('\n'), {
      x: b.x, y: 5.911, w: b.w, h: 0.961, fontSize: 10, lineSpacingMultiple: 1.3,
    });
  });

  [3.655, 6.372].forEach((x) => {
    s.addShape('line', { x, y: 5.695, w: 0, h: 1.05, line: { color: DARK, width: 0.5 } });
  });
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */
function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK', width: SW, height: SH });
  pptx.layout = 'DECK';
  pptx.title = 'BRAND GUIDELINES';

  slide01(pptx);
  slide02(pptx);
  slide03(pptx);
  chapterDivider(pptx, { page: '04', num: '02', title: 'LOGO',        titleW: 8.151, titleH: 1.582 });
  slide05(pptx);
  slide06(pptx);
  slide07(pptx);
  chapterDivider(pptx, { page: '08', num: '03', title: 'COLORS',      titleW: 8.151, titleH: 1.582 });
  slide09(pptx);
  slide10(pptx);
  chapterDivider(pptx, { page: '11', num: '04', title: 'TYPEFACE',    titleW: 8.151, titleH: 1.582 });
  slide12(pptx);
  chapterDivider(pptx, { page: '13', num: '05', title: 'IMAGERY',     titleW: 8.151, titleH: 1.582 });
  slide14(pptx);
  chapterDivider(pptx, { page: '15', num: '06', title: 'APPLICATION', titleW: 8.835, titleH: 1.565 });
  slide16(pptx);
  slide17(pptx);
  chapterDivider(pptx, { page: '18', num: '07', title: 'STATIONARY',  titleW: 8.835, titleH: 1.565 });
  slide19(pptx);
  slide20(pptx);

  return pptx.writeFile({
    fileName: path.join(__dirname, '0a9b8a23-22a6-4a07-ab14-9b6837b956d3_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
