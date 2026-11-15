/**
 * "CEO Report" deck — rebuilt with pptxgenjs.
 *
 * Slide size 13.333 x 7.5 in (16:9). Raster artwork from the original deck is
 * replaced by native-shape placeholders (see `laptopMockup` and `glyph`).
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const BLUE = '362BD2'; // accent1
const LIME = '8BFE3F'; // accent2
const GREEN = '00B46F'; // accent3
const WHITE = 'FFFFFF';
const BLACK = '000000';
const GREY = 'E7E6E6'; // bg2, used for the big open arc on slide 10
const GREY95 = 'F2F2F2'; // bg1 lumMod 95%
const BLUE_L = '867FE5'; // accent1 lum 60/40
const BLUE_D1 = '28209D'; // accent1 lum 75%
const BLUE_D2 = '1B1669'; // accent1 lum 50%

const HEAD = 'Plus Jakarta Sans SemiBold'; // theme major font
const BODY = 'Poppins Light'; // theme minor font

// pptxgenjs rewrites shadow options in place while serialising, so every shape
// needs its own copy — hence these are factories rather than constants.

/** The soft drop shadow the template puts under nearly every card / blob. */
const cardShadow = () => ({ type: 'outer', angle: 90, blur: 18, offset: 3, color: BLACK, opacity: 0.11 });
/** Wider, off-axis shadow used by the ring diagram on slide 10. */
const ringShadow = () => ({ type: 'outer', angle: 50, blur: 35, offset: 10, color: BLACK, opacity: 0.15 });

/* --------------------------------------------------------------- body copy */

const L_XL =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. ' +
  'Aenean massa. Cum sociis natoque penatibus ipsum dolor sit amet, consectetuer adipiscing elit. Aenean';
const L_LG =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. ' +
  'Aenean massa. Cum sociis natoque penatibus ipsum dolor sit amet.';
const L_MD =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. ' +
  'Aenean massa. Cum sociis natoque penatibus ipsum dolor sit';
const L_SM =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. ' +
  'Aenean massa. Cum sociis natoque penatibus';
const L_COL =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. ' +
  'Aenean massa. Cum sociis';
const L_QUOTE =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. ' +
  'Aenean massa. Cum sociis natoque';
const L_LEAD = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean';
const L_CARD = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget';
const L_TAG = 'Lorem ipsum dolor sit dolor sit amet, consectetuer';
const L_TINY = 'Lorem ipsum dolor sit amet, consectetuer';
const L_FEATURE = 'lorem ipsum dolor a amet consectetuer adipiscing elt';

/* ------------------------------------------------------------------ helpers */

// `roundRect` corner radius: PowerPoint stores an "adj" fraction of the short
// side; pptxgenjs wants that radius expressed in inches.
const radius = (adj, w, h) => adj * Math.min(w, h);

/** Text box. Defaults mirror the deck: body font, 12 pt, black, top aligned. */
function text(slide, content, opts) {
  slide.addText(content, Object.assign({ fontFace: BODY, fontSize: 12, color: BLACK, valign: 'top' }, opts));
}

/** Headline text (major font). */
function title(slide, content, opts) {
  text(slide, content, Object.assign({ fontFace: HEAD }, opts));
}

/** Header / footer furniture that the slide master paints on every page. */
function chrome(slide, pageNo) {
  title(slide, 'CEO Report', { x: 0.484, y: 0.307, w: 2.219, h: 0.303 });
  ['Home', 'About Us', 'Service', 'Contact'].forEach((label, i) => {
    title(slide, label, {
      x: 6.636 + i * 1.703,
      y: 0.241,
      w: 1.104,
      h: 0.345,
      align: 'center',
      lineSpacingMultiple: 1.3,
    });
  });
  text(slide, 'www. CEO Report.com', { x: 0.442, y: 6.956, w: 4.083, h: 0.303 });
  title(slide, `Page ${pageNo}`, { x: 11.256, y: 6.956, w: 1.635, h: 0.303, align: 'right' });
}

/**
 * The three nested "quarter blob" arcs used as background decoration.
 * All three derive from the width of the largest one.
 */
const BLOB_BIG = [
  // normalised outline of the large blob (0..1 of its own box)
  { x: 0.2429, y: 0 },
  { x: 1, y: 0.7702, curve: { type: 'cubic', x1: 0.661, y1: 0, x2: 1, y2: 0.3448 } },
  { x: 0.966, y: 0.9992, curve: { type: 'cubic', x1: 1, y1: 0.85, x2: 0.9881, y2: 0.9269 } },
  { x: 0.9657, y: 1 },
  { x: 0, y: 1 },
  { x: 0, y: 0.0408 },
  { x: 0.0178, y: 0.0346 },
  { x: 0.2429, y: 0, curve: { type: 'cubic', x1: 0.0889, y1: 0.0121, x2: 0.1645, y2: 0 } },
  { close: true },
];
const BLOB_SMALL = [
  { x: 0.3685, y: 0 },
  { x: 1, y: 0.6483, curve: { type: 'cubic', x1: 0.7172, y1: 0, x2: 1, y2: 0.2903 } },
  { x: 0.9238, y: 0.9573, curve: { type: 'cubic', x1: 1, y1: 0.7602, x2: 0.9724, y2: 0.8654 } },
  { x: 0.8985, y: 1 },
  { x: 0, y: 1 },
  { x: 0, y: 0.1225 },
  { x: 0.0153, y: 0.1107 },
  { x: 0.3685, y: 0, curve: { type: 'cubic', x1: 0.1161, y1: 0.0408, x2: 0.2376, y2: 0 } },
  { close: true },
];

function blob(slide, outline, x, y, w, h) {
  const pts = outline.map((p) => {
    if (p.close) return p;
    const pt = { x: +(p.x * w).toFixed(4), y: +(p.y * h).toFixed(4) };
    if (p.curve) {
      pt.curve = {
        type: 'cubic',
        x1: +(p.curve.x1 * w).toFixed(4),
        y1: +(p.curve.y1 * h).toFixed(4),
        x2: +(p.curve.x2 * w).toFixed(4),
        y2: +(p.curve.y2 * h).toFixed(4),
      };
    }
    return pt;
  });
  slide.addShape('custGeom', { x, y, w, h, points: pts, fill: { color: WHITE }, shadow: cardShadow() });
}

/** Nested decorative blobs anchored at (x, y); `w` sizes the outermost one. */
function blobTrio(slide, x, y, w) {
  blob(slide, BLOB_BIG, x, y, w, w * 0.9829);
  blob(slide, BLOB_SMALL, x, y + w * 0.3174, w * 0.6875, w * 0.6697);
  blob(slide, BLOB_SMALL, x, y + w * 0.5822, w * 0.4157, w * 0.4049);
}

/** Pill-shaped label chip (rounded rectangle + centred caption). */
function chip(slide, o) {
  const w = o.w;
  const h = o.h === undefined ? 0.515 : o.h;
  const adj = o.adj === undefined ? 0.16667 : o.adj;
  slide.addShape('roundRect', {
    x: o.x,
    y: o.y,
    w: w,
    h: h,
    rectRadius: radius(adj, w, h),
    fill: { color: o.fill || LIME },
    shadow: o.flat ? undefined : cardShadow(),
    rotate: o.rotate,
  });
  // Buttons that carry a chevron badge nudge their caption left of centre.
  title(slide, o.label, {
    x: o.x + (o.labelDx === undefined ? (w - o.labelW) / 2 : o.labelDx),
    y: o.y + (h - 0.303) / 2,
    w: o.labelW,
    h: 0.303,
    fontSize: o.fontSize || 12,
    align: 'center',
    rotate: o.rotate,
  });
}

/** Small white disc holding a lime chevron — the "go" affordance on buttons. */
function chevronBadge(slide, x, y) {
  slide.addShape('ellipse', { x: x, y: y, w: 0.285, h: 0.285, fill: { color: WHITE }, rotate: 90 });
  slide.addShape('chevron', { x: x + 0.107, y: y + 0.085, w: 0.071, h: 0.114, fill: { color: LIME } });
}

/** Tilted "Report"/"Achieve" sticker used across the opening slides. */
function sticker(slide, x, y, label, rotate) {
  chip(slide, {
    x: x,
    y: y,
    w: 1.223,
    h: 0.434,
    adj: 0.13467,
    label: label,
    labelW: label === 'Achieve' ? 1.0 : 0.868,
    fontSize: 14,
    rotate: rotate,
  });
}

/** Big-number + caption pair (skills row on slide 8, KPI row on slide 12). */
function stat(slide, x, y, w, value, caption) {
  title(slide, value, { x: x, y: y, w: w, h: 0.505, fontSize: 24 });
  text(slide, caption, { x: x, y: y + 0.445, w: w * 0.806, h: 0.33, lineSpacingMultiple: 1.2, wrap: false });
}

/** Circular "→" bullet that precedes the italic "Subtitle Here" captions. */
function arrowDot(slide, x, y, ringColor) {
  slide.addShape('ellipse', { x: x, y: y, w: 0.229, h: 0.229, fill: { color: ringColor } });
  slide.addShape('rightArrow', { x: x + 0.056, y: y + 0.094, w: 0.117, h: 0.041, fill: { color: WHITE } });
}

/**
 * Placeholder for a raster icon: an approximate native-shape stand-in drawn in
 * the icon's own colour and bounding box.
 */
function glyph(slide, kind, x, y, w, h, color) {
  const fill = { color: color };
  const box = (dx, dy, dw, dh, shape, extra) =>
    slide.addShape(shape || 'rect', Object.assign({ x: x + w * dx, y: y + h * dy, w: w * dw, h: h * dh, fill: fill }, extra));

  switch (kind) {
    case 'gear':
      box(0, 0, 1, 1, 'gear6');
      break;
    case 'rocket': // capsule body with porthole, side fins and exhaust flame
      box(0.22, 0, 0.56, 0.72, 'round2SameRect', { rectRadius: w * 0.28 });
      box(0.38, 0.14, 0.24, 0.19, 'ellipse', { fill: { color: WHITE } });
      box(0, 0.42, 0.3, 0.3, 'rtTriangle', { flipH: true });
      box(0.7, 0.42, 0.3, 0.3, 'rtTriangle');
      box(0.22, 0.7, 0.56, 0.12);
      box(0.34, 0.83, 0.32, 0.17, 'triangle', { rotate: 180 });
      break;
    case 'people': // disc with two heads knocked out in white
      box(0, 0, 1, 1, 'ellipse');
      box(0.2, 0.25, 0.24, 0.24, 'ellipse', { fill: { color: WHITE } });
      box(0.54, 0.3, 0.2, 0.2, 'ellipse', { fill: { color: WHITE } });
      box(0.12, 0.62, 0.34, 0.38, 'ellipse', { fill: { color: WHITE } });
      box(0.54, 0.66, 0.3, 0.34, 'ellipse', { fill: { color: WHITE } });
      break;
    case 'briefcase':
      box(0.3, 0.06, 0.4, 0.18, 'roundRect', { rectRadius: w * 0.04 });
      box(0, 0.24, 1, 0.62, 'roundRect', { rectRadius: w * 0.06 });
      break;
    case 'layers':
      [0, 0.3, 0.6].forEach((dy) => box(0, dy, 1, 0.4, 'diamond'));
      break;
    case 'pie': // three-quarter disc with a detached quadrant top-right
      box(0, 0.1, 0.9, 0.9, 'pie', { angleRange: [0, 270] });
      box(0.55, 0.1, 0.45, 0.45, 'pieWedge', { rotate: 180 });
      break;
    case 'plane': // paper-plane dart
      box(0, 0, 1, 1, 'triangle', { rotate: 120 });
      break;
    case 'globe': // ringed sphere with meridian and equator
      box(0, 0, 1, 1, 'donut');
      slide.addShape('ellipse', {
        x: x + w * 0.28,
        y: y + h * 0.06,
        w: w * 0.44,
        h: h * 0.88,
        fill: { type: 'none' },
        line: { color: color, width: 1.2 },
      });
      box(0.05, 0.46, 0.9, 0.07);
      break;
  }
}

/**
 * Stand-in for the laptop photo on slide 9 (the original raster is not
 * embedded): dark bezel rectangle plus a flat cyan "screen".
 */
function laptopMockup(slide) {
  slide.addShape('roundRect', { x: 6.1, y: 2.72, w: 7.24, h: 4.78, rectRadius: 0.16, fill: { color: '111111' } });
  slide.addShape('rect', { x: 6.36, y: 3.09, w: 6.98, h: 4.41, fill: { color: '00BEF2' } });
  text(slide, '[image]', { x: 12.2, y: 7.1, w: 1.0, h: 0.28, fontSize: 10, color: WHITE, align: 'right' });
}

/* ---------------------------------------------------------------- slide 1 */

function slide01(pptx) {
  const s = pptx.addSlide();
  chrome(s, 1);
  blobTrio(s, 0, 0.847, 6.74);

  title(s, 'CEO Report', { x: 5.472, y: 2.775, w: 6.665, h: 1.212, fontSize: 66, align: 'center' });
  [3.965, 4.614, 5.263].forEach((x) =>
    s.addShape('ellipse', { x: x, y: 3.131, w: 0.5, h: 0.5, fill: { color: BLUE }, shadow: cardShadow() })
  );

  text(s, 'Template Presentations', { x: 7.35, y: 4.357, w: 2.359, h: 0.303 });
  chip(s, { x: 9.647, y: 4.291, w: 1.739, h: 0.434, adj: 0.13467, label: 'Learn More', labelW: 1.14, labelDx: 0.126 });
  chevronBadge(s, 10.957, 4.366);
}

/* ---------------------------------------------------------------- slide 2 */

function slide02(pptx) {
  const s = pptx.addSlide();
  chrome(s, 2);
  blobTrio(s, 8.591, 2.819, 4.742);

  text(s, '- Reflecting on Achievements', { x: 0.731, y: 0.985, w: 3.685, h: 0.303 });
  title(s, 'CEO Opening Statement', { x: 0.731, y: 1.402, w: 8.973, h: 0.909, fontSize: 48 });
  text(s, L_XL, { x: 0.731, y: 2.376, w: 8.78, h: 0.608, lineSpacingMultiple: 1.3 });

  title(s, 'Q2', { x: 0.731, y: 3.065, w: 4.81, h: 3.45, fontSize: 199, color: BLUE });
  sticker(s, 2.364, 4.455, 'Report', 349.07);

  chip(s, { x: 9.207, y: 3.215, w: 2.807, label: 'Reflecting on Achievements', labelW: 2.623 });
}

/* ---------------------------------------------------------------- slide 3 */

function slide03(pptx) {
  const s = pptx.addSlide();
  chrome(s, 3);

  title(
    s,
    [
      { text: 'Comprehensive ' },
      { text: 'Company Overview', options: { color: BLUE } },
      { text: ' and ' },
      { text: 'Vision for Sustainable', options: { color: BLUE } },
      { text: ' Growth and Industry Leadership' },
    ],
    { x: 0.459, y: 1.402, w: 8.096, h: 1.717, fontSize: 32 }
  );

  title(s, '2025', { x: 8.544, y: 1.167, w: 4.507, h: 2.036, fontSize: 115, color: BLUE, align: 'center' });
  s.addShape('rect', { x: 8.827, y: 2.306, w: 4.224, h: 0.658, fill: { color: WHITE }, shadow: cardShadow() });
  sticker(s, 10.201, 2.082, 'Report', 349.07);

  [0.459, 4.763, 9.068].forEach((x) => {
    text(s, L_COL, { x: x, y: 3.604, w: 3.806, h: 0.87, lineSpacingMultiple: 1.3 });
    title(s, '- Reflecting on Achievements', { x: x, y: 4.609, w: 3.685, h: 0.303 });
  });
}

/* ---------------------------------------------------------------- slide 4 */

function slide04(pptx) {
  const s = pptx.addSlide();
  chrome(s, 4);

  text(s, L_LG, { x: 0.615, y: 6.014, w: 7.29, h: 0.608, lineSpacingMultiple: 1.3 });
  chip(s, { x: 5.813, y: 5.323, w: 2.807, label: 'Reflecting on Achievements', labelW: 2.623 });
  title(s, 'Q2', { x: 8.246, y: 5.763, w: 1.665, h: 1.111, fontSize: 60, color: BLUE, align: 'right' });
  text(s, L_TINY, { x: 9.911, y: 6.014, w: 2.74, h: 0.608, lineSpacingMultiple: 1.3 });
}

/* ---------------------------------------------------------------- slide 5 */

function slide05(pptx) {
  const s = pptx.addSlide();
  chrome(s, 5);
  blobTrio(s, 6.535, 0.79, 6.798);

  title(
    s,
    [
      { text: 'Comprehensive ' },
      { text: 'Operational Highlights', options: { color: BLUE } },
      { text: ' and ' },
      { text: 'Business Performance', options: { color: BLUE } },
      { text: ' Analysis' },
    ],
    { x: 0.563, y: 1.402, w: 6.57, h: 1.717, fontSize: 32 }
  );
  sticker(s, 5.228, 2.373, 'Achieve', 349.07);
  text(s, L_XL, { x: 0.563, y: 3.615, w: 4.74, h: 1.133, align: 'justify', lineSpacingMultiple: 1.3 });

  s.addShape('roundRect', {
    x: 5.995,
    y: 3.615,
    w: 2.276,
    h: 2.019,
    rectRadius: radius(0.07101, 2.276, 2.019),
    fill: { color: LIME },
    shadow: cardShadow(),
  });
  title(s, '80% of Businesses Anticipate', { x: 6.168, y: 4.119, w: 1.93, h: 1.01, fontSize: 18, align: 'center' });

  s.addShape('roundRect', {
    x: 8.59,
    y: 2.442,
    w: 2.163,
    h: 1.354,
    rectRadius: radius(0.07101, 2.163, 1.354),
    fill: { color: BLUE },
    shadow: cardShadow(),
  });
  title(s, '$1 Million or More Annually', {
    x: 8.707,
    y: 2.766,
    w: 1.93,
    h: 0.64,
    fontSize: 16,
    color: WHITE,
    align: 'center',
  });
}

/* ---------------------------------------------------------------- slide 6 */

function slide06(pptx) {
  const s = pptx.addSlide();
  chrome(s, 6);

  s.addShape('roundRect', {
    x: 2.542,
    y: 2.0,
    w: 8.25,
    h: 3.5,
    rectRadius: radius(0.03869, 8.25, 3.5),
    fill: { color: BLUE },
    shadow: cardShadow(),
  });
  text(s, 'Quarterly Report Overview', { x: 4.824, y: 2.762, w: 3.685, h: 0.303, color: WHITE, align: 'center' });
  title(s, 'PLACEHOLDER', {
    x: 3.382,
    y: 3.022,
    w: 6.57,
    h: 1.717,
    fontSize: 32,
    color: WHITE,
    align: 'center',
  });

  sticker(s, 3.152, 2.822, 'Achieve', 349.07);
  sticker(s, 8.856, 4.088, 'Report', 8.32);
}

/* ---------------------------------------------------------------- slide 7 */

function slide07(pptx) {
  const s = pptx.addSlide();
  chrome(s, 7);
  blobTrio(s, 0, 3.113, 4.444);

  title(
    s,
    [
      { text: 'Shaping the Next ' },
      { text: 'Chapter of Sustainable', options: { color: BLUE } },
      { text: ' Expansion' },
    ],
    { x: 0.563, y: 1.402, w: 5.885, h: 1.178, fontSize: 32 }
  );
  text(s, L_MD, { x: 0.563, y: 2.858, w: 4.74, h: 0.87, align: 'justify', lineSpacingMultiple: 1.3 });

  // Two stacked "25%" cards on the right.
  [
    { y: 2.869, fill: LIME, fg: BLACK },
    { y: 4.523, fill: BLUE, fg: WHITE },
  ].forEach((card) => {
    s.addShape('roundRect', {
      x: 8.639,
      y: card.y,
      w: 4.0,
      h: 1.354,
      rectRadius: radius(0.07101, 4.0, 1.354),
      fill: { color: card.fill },
      shadow: cardShadow(),
    });
    title(s, '25%', { x: 8.868, y: card.y + 0.257, w: 1.77, h: 0.841, fontSize: 44, color: card.fg });
    text(s, L_TAG, {
      x: 10.639,
      y: card.y + 0.242,
      w: 1.77,
      h: 0.87,
      color: card.fg,
      align: 'justify',
      lineSpacingMultiple: 1.3,
    });
  });
}

/* ---------------------------------------------------------------- slide 8 */

function slide08(pptx) {
  const s = pptx.addSlide();
  chrome(s, 8);
  blobTrio(s, 6.535, 0.79, 6.798);

  title(s, 'Guiding the Company Toward Sustainable Success', {
    x: 0.563,
    y: 1.402,
    w: 5.497,
    h: 1.717,
    fontSize: 32,
  });
  chip(s, { x: 2.842, y: 2.841, w: 2.145, label: 'Mark Medison', labelW: 1.501, labelDx: 0.18 });
  chevronBadge(s, 4.523, 2.956);
  text(s, L_SM, { x: 0.563, y: 3.837, w: 4.425, h: 0.87, align: 'justify', lineSpacingMultiple: 1.3 });

  [
    { x: 0.563, value: '80%' },
    { x: 2.268, value: '90%' },
    { x: 3.973, value: '90%' },
  ].forEach((col) => stat(s, col.x, 5.038, 1.147, col.value, 'Your Skills'));
  [1.88, 3.698].forEach((x) =>
    s.addShape('line', { x: x, y: 5.038, w: 0, h: 0.834, line: { color: BLACK, width: 1 } })
  );
}

/* ---------------------------------------------------------------- slide 9 */

function slide09(pptx) {
  const s = pptx.addSlide();
  chrome(s, 9);
  blobTrio(s, 0, 3.601, 3.95);
  laptopMockup(s);

  title(
    s,
    [
      { text: 'Driving Market ' },
      { text: 'Leadership Through', options: { color: BLUE } },
      { text: ' Continuous Innovation' },
    ],
    { x: 0.781, y: 1.402, w: 5.885, h: 1.717, fontSize: 32 }
  );
  text(s, L_LEAD, { x: 7.093, y: 1.543, w: 4.74, h: 0.608, align: 'justify', lineSpacingMultiple: 1.3 });

  s.addShape('roundRect', {
    x: 3.749,
    y: 3.5,
    w: 4.0,
    h: 1.354,
    rectRadius: radius(0.07101, 4.0, 1.354),
    fill: { color: LIME },
    shadow: cardShadow(),
  });
  title(s, '25%', { x: 3.978, y: 3.757, w: 1.77, h: 0.841, fontSize: 44 });
  text(s, L_TAG, { x: 5.749, y: 3.743, w: 1.77, h: 0.87, align: 'justify', lineSpacingMultiple: 1.3 });
}

/* --------------------------------------------------------------- slide 10 */

function slide10(pptx) {
  const s = pptx.addSlide();
  chrome(s, 10);

  // Concentric ring diagram.
  s.addShape('arc', {
    x: 3.256,
    y: 1.63,
    w: 4.522,
    h: 4.522,
    angleRange: [224.4912, 137.4024],
    flipH: true,
    line: { color: GREY, width: 12 },
  });
  s.addShape('flowChartConnector', { x: 3.533, y: 1.908, w: 3.967, h: 3.967, fill: { color: GREY95 } });
  s.addShape('flowChartConnector', {
    x: 4.212,
    y: 2.591,
    w: 2.6,
    h: 2.6,
    fill: { type: 'none' },
    line: { color: WHITE, width: 2.5 },
    shadow: ringShadow(),
  });
  s.addShape('flowChartConnector', {
    x: 4.397,
    y: 2.771,
    w: 2.24,
    h: 2.24,
    fill: { color: WHITE },
    shadow: ringShadow(),
  });

  title(s, 'Revenue Growth Playbook Winning Moves for Modern', {
    x: 5.549,
    y: 3.302,
    w: 6.802,
    h: 1.313,
    fontSize: 36,
  });

  // Icon discs sitting on the ring.
  const disc = (x, y, d) =>
    s.addShape('ellipse', { x: x, y: y, w: d, h: d, fill: { color: WHITE }, shadow: ringShadow() });

  disc(3.826, 1.487, 0.881);
  glyph(s, 'gear', 4.061, 1.711, 0.41, 0.432, LIME);
  disc(2.884, 3.427, 0.881);
  glyph(s, 'briefcase', 3.069, 3.62, 0.495, 0.495, BLUE);
  disc(3.956, 5.41, 0.881);
  glyph(s, 'rocket', 4.236, 5.628, 0.322, 0.444, LIME);
  disc(6.764, 5.0, 0.881);
  glyph(s, 'people', 6.971, 5.207, 0.468, 0.468, BLUE);

  disc(6.386, 1.5, 1.531);
  title(
    s,
    [
      { text: '80', options: { fontSize: 36 } },
      { text: '%', options: { fontSize: 24 } },
    ],
    { x: 6.538, y: 1.912, w: 1.226, h: 0.707, color: GREEN, align: 'center', wrap: false }
  );

  // Callout copy around the diagram.
  text(s, 'Lorem ipsum dolor sit amet cum consectetuer legit', {
    x: 0.951,
    y: 1.487,
    w: 2.639,
    h: 0.608,
    align: 'right',
    lineSpacingMultiple: 1.3,
  });
  text(s, 'Lorem ipsum dolor sit eget et consectetuer', {
    x: 1.056,
    y: 3.432,
    w: 1.551,
    h: 0.87,
    align: 'right',
    lineSpacingMultiple: 1.3,
  });
  text(s, 'Lorem ipsum dolor sit amet cum consectetuer legit', {
    x: 0.951,
    y: 5.641,
    w: 2.639,
    h: 0.608,
    align: 'right',
    lineSpacingMultiple: 1.3,
  });
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula', {
    x: 8.134,
    y: 1.451,
    w: 4.217,
    h: 0.608,
    lineSpacingMultiple: 1.3,
  });
  arrowDot(s, 8.352, 2.33, 'C2BFF2');
  text(s, 'Subtitle Here', {
    x: 8.591,
    y: 2.274,
    w: 1.652,
    h: 0.345,
    color: GREEN,
    italic: true,
    align: 'justify',
    lineSpacingMultiple: 1.3,
  });
  text(s, 'Lorem ipsum dolor sit amet, consectetuer et adipiscing elit. Aenean commodo ligula legit', {
    x: 8.056,
    y: 4.949,
    w: 4.541,
    h: 0.608,
    lineSpacingMultiple: 1.3,
  });
}

/* --------------------------------------------------------------- slide 11 */

function slide11(pptx) {
  const s = pptx.addSlide();
  chrome(s, 11);

  title(s, 'Revenue Growth Playbook Winning Moves for Modern Companies.', {
    x: 0.602,
    y: 1.348,
    w: 8.919,
    h: 1.313,
    fontSize: 36,
  });
  title(s, '2025', { x: 0.602, y: 4.932, w: 1.52, h: 0.505, fontSize: 24, bold: true });
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo', {
    x: 0.602,
    y: 5.432,
    w: 2.804,
    h: 0.866,
    lineSpacingMultiple: 1.3,
  });

  // Three step cards.
  const cards = [
    { x: 4.558, fill: BLUE, fg: WHITE, dot: '7169DE' },
    { x: 7.38, fill: LIME, fg: BLACK, dot: '61B22C' },
    { x: 10.202, fill: GREY95, fg: BLACK },
  ];
  cards.forEach((c) =>
    s.addShape('roundRect', {
      x: c.x,
      y: 3.502,
      w: 2.57,
      h: 2.977,
      rectRadius: radius(0.04101, 2.57, 2.977),
      fill: { color: c.fill },
    })
  );

  cards.slice(0, 2).forEach((c, i) => {
    chip(s, {
      x: c.x + 1.122,
      y: 3.313,
      w: 1.115,
      h: 0.378,
      adj: 0.5,
      fill: WHITE,
      label: `Step 0${i + 1}`,
      labelW: 0.925,
    });
    title(s, 'Your Target', {
      x: c.x + 0.207,
      y: 4.094,
      w: c.fill === BLUE ? 1.587 : 1.45,
      h: 0.37,
      fontSize: 16,
      bold: true,
      color: c.fg,
      wrap: false,
    });
    text(s, L_CARD, {
      x: c.x + 0.207,
      y: 4.465,
      w: 2.156,
      h: 1.129,
      color: c.fg,
      lineSpacingMultiple: 1.3,
    });
    arrowDot(s, c.x + 0.307, 5.751, c.dot);
    text(s, 'Subtitle Here', {
      x: c.x + 0.546,
      y: 5.695,
      w: 1.652,
      h: 0.345,
      color: c.fg,
      italic: true,
      align: 'justify',
      lineSpacingMultiple: 1.3,
    });
  });

  // Third card: metric + dark icon disc.
  s.addShape('ellipse', { x: 10.621, y: 3.847, w: 0.749, h: 0.749, fill: { color: BLACK } });
  glyph(s, 'gear', 10.812, 4.029, 0.366, 0.385, WHITE);
  title(s, '36K', { x: 10.485, y: 4.666, w: 1.708, h: 0.841, fontSize: 44, bold: true });
  text(s, L_TINY, { x: 10.485, y: 5.432, w: 2.156, h: 0.604, lineSpacingMultiple: 1.3 });
}

/* --------------------------------------------------------------- slide 12 */

function slide12(pptx) {
  const s = pptx.addSlide();
  chrome(s, 12);

  title(s, 'Market Trends and Competitive Insights', {
    x: 1.152,
    y: 1.874,
    w: 5.26,
    h: 1.192,
    fontSize: 36,
    lineSpacingMultiple: 0.9,
  });
  text(s, L_MD, { x: 1.152, y: 3.383, w: 4.74, h: 0.87, align: 'justify', lineSpacingMultiple: 1.3 });
  stat(s, 1.152, 4.851, 2.396, '30% Margin', 'Lorem ipsum dolor ');
  stat(s, 3.812, 4.851, 2.396, '$ 5 Million', 'Lorem ipsum dolor ');

  // Nested bubble chart: outermost (lightest) first.
  const bubbles = [
    { x: 7.053, y: 1.186, d: 5.129, fill: BLUE_L, label: '$67.5K', ly: 1.492, lw: 1.357, lh: 0.505 },
    { x: 7.565, y: 2.209, d: 4.105, fill: BLUE, label: '$49.2K', ly: 2.597, lw: 1.334, lh: 0.487 },
    { x: 8.113, y: 3.299, d: 3.015, fill: BLUE_D1, label: '$33.4K', ly: 3.69, lw: 1.341, lh: 0.487 },
    { x: 8.656, y: 4.392, d: 1.922, fill: BLUE_D2, label: '$24.8K', ly: 5.11, lw: 1.344, lh: 0.487 },
  ];
  bubbles.forEach((b) =>
    s.addShape('flowChartConnector', { x: b.x, y: b.y, w: b.d, h: b.d, fill: { color: b.fill } })
  );
  bubbles.forEach((b) =>
    title(s, b.label, { x: 8.945, y: b.ly, w: b.lw, h: b.lh, fontSize: 24, color: WHITE, wrap: false })
  );
}

/* --------------------------------------------------------------- slide 13 */

function slide13(pptx) {
  const s = pptx.addSlide();
  chrome(s, 13);

  title(s, 'Strategic Goals for the Next Quarter', {
    x: 0.77,
    y: 1.535,
    w: 4.973,
    h: 1.192,
    fontSize: 36,
    lineSpacingMultiple: 0.9,
  });
  title(s, '90+', { x: 0.77, y: 4.503, w: 1.44, h: 0.646, fontSize: 36, color: BLUE, lineSpacingMultiple: 0.9 });
  text(s, L_SM, { x: 0.77, y: 5.19, w: 4.425, h: 0.87, align: 'justify', lineSpacingMultiple: 1.3 });

  // Five columns; each is a bar rising to `top` with a year tag above it.
  const bars = [
    { x: 6.122, top: 3.198, year: '2024', fill: WHITE, fg: BLACK, labelY: 5.164 },
    { x: 7.361, top: 2.229, year: '2025', fill: BLUE, fg: WHITE, labelY: 4.678 },
    { x: 8.601, top: 3.75, year: '2026', fill: WHITE, fg: BLACK, labelY: 5.44 },
    { x: 9.841, top: 3.064, year: '2027', fill: LIME, fg: BLACK, labelY: 5.097 },
    { x: 11.08, top: 4.438, year: '2028', fill: WHITE, fg: BLACK, labelY: 5.784 },
  ];
  bars.forEach((b) => {
    s.addShape('round2SameRect', {
      x: b.x,
      y: b.top,
      w: 1.115,
      h: 7.5 - b.top,
      rectRadius: 0.083,
      fill: { color: b.fill },
      shadow: b.fill === BLUE ? undefined : cardShadow(),
    });
    s.addShape('roundRect', {
      x: b.x,
      y: b.top - 0.511,
      w: 1.115,
      h: 0.41,
      rectRadius: radius(0.23199, 1.115, 0.41),
      fill: { color: b.fill },
      shadow: b.fill === BLUE ? undefined : cardShadow(),
    });
    title(s, b.year, {
      x: b.x + 0.155,
      y: b.top - 0.491,
      w: 0.775,
      h: 0.37,
      fontSize: 16,
      color: b.fg,
      wrap: false,
    });
    // Vertical caption inside the bar (rotated 270°).
    title(s, 'Your Text Here', {
      x: b.x - 0.342,
      y: b.labelY,
      w: 1.773,
      h: 0.37,
      fontSize: 16,
      color: b.fg,
      align: 'center',
      rotate: 270,
      wrap: false,
    });
  });
}

/* --------------------------------------------------------------- slide 14 */

function slide14(pptx) {
  const s = pptx.addSlide();
  chrome(s, 14);

  title(s, 'User Growth and Engagement Stats', {
    x: 1.031,
    y: 1.576,
    w: 11.145,
    h: 0.646,
    fontSize: 36,
    align: 'center',
    lineSpacingMultiple: 0.9,
  });

  // Five columns alternate: blue bubble above the rule (tail down) with caption
  // underneath, lime bubble below the rule (tail up) with caption on top.
  const columns = [
    { rule: 0.941, color: BLUE, icon: 'briefcase', heading: 'Accuracy', bubble: [1.429, 3.064], head: 4.757 },
    { rule: 3.317, color: LIME, icon: 'pie', heading: 'Refrences', bubble: [3.805, 4.905], head: 2.991 },
    { rule: 5.692, color: BLUE, icon: 'layers', heading: 'Objectivity', bubble: [6.18, 3.064], head: 4.757 },
    { rule: 8.068, color: LIME, icon: 'plane', heading: 'Representation', bubble: [8.556, 4.852], head: 2.938 },
    { rule: 10.443, color: BLUE, icon: 'globe', heading: 'Accessibility', bubble: [10.931, 3.064], head: 4.757 },
  ];

  columns.forEach((c) => {
    const tailUp = c.color === LIME;
    s.addShape('roundRect', {
      x: c.rule,
      y: 4.403,
      w: 1.95,
      h: 0.086,
      rectRadius: radius(0.5, 1.95, 0.086),
      fill: { color: c.color },
    });
    s.addShape('wedgeRoundRectCallout', {
      x: c.bubble[0],
      y: c.bubble[1],
      w: 0.974,
      h: 0.975,
      fill: { color: c.color },
      flipV: tailUp,
      shadow: tailUp ? cardShadow() : undefined,
    });
    glyph(s, c.icon, c.bubble[0] + 0.277, c.bubble[1] + (tailUp ? 0.278 : 0.26), 0.42, 0.44, WHITE);

    title(s, c.heading, {
      x: c.rule - 0.057,
      y: c.head,
      w: 2.065,
      h: 0.325,
      fontSize: 16,
      bold: true,
      align: 'center',
      lineSpacingMultiple: 0.8,
    });
    text(s, L_FEATURE, {
      x: c.rule - 0.153,
      y: c.head + 0.297,
      w: 2.256,
      h: 0.87,
      align: 'center',
      lineSpacingMultiple: 1.3,
    });
  });
}

/* --------------------------------------------------------------- slide 15 */

function slide15(pptx) {
  const s = pptx.addSlide();
  chrome(s, 15);

  // Oversized quotation mark behind the testimonial.
  s.addShape('custGeom', {
    x: 5.338,
    y: 3.199,
    w: 2.658,
    h: 2.218,
    fill: { color: BLACK, transparency: 90 },
    points: [
      { x: 2.658, y: 0 },
      { x: 2.658, y: 0.457 },
      { x: 2.06, y: 1.138, curve: { type: 'cubic', x1: 2.48, y1: 0.501, x2: 2.081, y2: 0.966 } },
      { x: 2.491, y: 1.138 },
      { x: 2.491, y: 2.218 },
      { x: 1.536, y: 2.218 },
      { x: 1.536, y: 1.512 },
      { x: 2.658, y: 0, curve: { type: 'cubic', x1: 1.536, y1: 1.108, x2: 1.632, y2: 0.51 } },
      { close: true },
      { x: 1.121, y: 0, moveTo: true },
      { x: 1.121, y: 0.457 },
      { x: 0.523, y: 1.138, curve: { type: 'cubic', x1: 0.944, y1: 0.501, x2: 0.545, y2: 0.966 } },
      { x: 0.955, y: 1.138 },
      { x: 0.955, y: 2.218 },
      { x: 0, y: 2.218 },
      { x: 0, y: 1.512 },
      { x: 1.121, y: 0, curve: { type: 'cubic', x1: 0, y1: 1.108, x2: 0.095, y2: 0.51 } },
      { close: true },
    ],
  });

  blobTrio(s, 0, 0.79, 6.798);

  title(s, 'What Our Clients and Partners Say About Our Brand', {
    x: 6.226,
    y: 1.787,
    w: 6.649,
    h: 1.07,
    fontSize: 32,
    lineSpacingMultiple: 0.9,
  });
  text(s, L_QUOTE, { x: 8.396, y: 3.199, w: 4.281, h: 0.87, align: 'justify', lineSpacingMultiple: 1.3 });
  text(s, L_CARD, { x: 8.396, y: 4.131, w: 4.281, h: 0.608, align: 'justify', lineSpacingMultiple: 1.3 });
  chip(s, { x: 8.396, y: 5.008, w: 2.145, label: 'Mark Medison', labelW: 1.501, labelDx: 0.18 });
  chevronBadge(s, 10.077, 5.123);
}

/* --------------------------------------------------------------- slide 16 */

function slide16(pptx) {
  const s = pptx.addSlide();
  chrome(s, 16);
  blobTrio(s, 0, 0.847, 6.74);

  title(s, 'Thank You!!', { x: 5.472, y: 2.775, w: 6.665, h: 1.212, fontSize: 66, align: 'center' });
  [3.965, 4.614, 5.263].forEach((x) =>
    s.addShape('ellipse', { x: x, y: 3.131, w: 0.5, h: 0.5, fill: { color: BLUE }, shadow: cardShadow() })
  );

  [
    { x: 3.869, w: 1.393, label: '+123 456 7890' },
    { x: 5.376, w: 1.932, label: '12 Your Street Name' },
    { x: 7.421, w: 2.081, label: 'www. Your Name.com' },
  ].forEach((item) =>
    text(s, item.label, {
      x: item.x,
      y: 4.275,
      w: item.w,
      h: 0.375,
      align: 'center',
      lineSpacingMultiple: 1.5,
    })
  );

  chip(s, { x: 9.647, y: 4.291, w: 1.739, h: 0.434, adj: 0.13467, label: 'Learn More', labelW: 1.14, labelDx: 0.126 });
  chevronBadge(s, 10.957, 4.366);
}

/* ------------------------------------------------------------------- build */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK', width: 13.333, height: 7.5 });
  pptx.layout = 'DECK';
  pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
  pptx.title = 'CEO Report';

  [
    slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
    slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
  ].forEach((fn) => fn(pptx));

  return pptx.writeFile({
    fileName: path.join(__dirname, '07fbfd28-9dcd-4921-a83a-73ae37b73566_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f));
