/**
 * "Wanderlust" travel deck - rebuilt with pptxgenjs.
 *
 * Slide size 13.333 x 7.5 in (16:9). Every measurement below is in inches and
 * comes straight from the reference deck.
 *
 * Raster art in the original (icons, decorative "photo card" silhouettes) is
 * redrawn here with native pptxgenjs shapes; the empty picture placeholders are
 * drawn as the light-grey notched cards the slide layouts define for them.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const TEAL = '27A59E';
const ORANGE = 'EF8F38';
const INK = '262626'; // tx1 @ 85% lum -- headline colour
const GREY = '808080'; // bg1 @ 50% lum -- body copy
const WHITE = 'FFFFFF';
const CREAM = 'FCE9D7'; // faint accent used for the footer URL
const PHOTO = 'F2F2F2'; // bg1 @ 95% lum -- empty picture placeholder fill

const SERIF = 'Marcellus';
const LIGHT = 'Montserrat Light';
const MEDIUM = 'Montserrat Medium';
const SEMI = 'Montserrat SemiBold';
const BOOK = 'Montserrat';

// pptxgenjs rewrites the shadow object in place, so hand it a fresh one each time.
const softShadow = () => ({ type: 'outer', blur: 15, offset: 3, angle: 45, color: '000000', opacity: 0.1 });

/* ---------------------------------------------------------------- helpers */

/** Body/heading text box: top anchored, PowerPoint default insets. */
function text(slide, body, opts) {
  slide.addText(body, Object.assign({ valign: 'top', fontFace: LIGHT, fontSize: 10, color: GREY }, opts));
}

/** The ".Wanderlust" wordmark that repeats on every slide. */
function wordmark(slide) {
  text(slide, [
    { text: '.Wander', options: { fontFace: SEMI, bold: true, color: ORANGE } },
    { text: 'lust', options: { fontFace: MEDIUM, color: TEAL } }
  ], { x: 0.35, y: 0.424, w: 1.605, h: 0.286, fontSize: 11, charSpacing: 2 });
}

/** Bottom-right footer URL (every slide except the cover). */
function footerUrl(slide) {
  text(slide, 'www.yourgreatsite.com', {
    x: 9.691, y: 6.929, w: 3.283, h: 0.269,
    fontSize: 10, fontFace: BOOK, color: CREAM, charSpacing: 3, align: 'right'
  });
}

/** Pill button with centred label. */
function pill(slide, label, o) {
  slide.addText(label, {
    shape: 'roundRect', rectRadius: 0.18014 * Math.min(o.w, o.h),
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: { color: o.fill || TEAL }, line: { type: 'none' },
    align: 'center', valign: 'middle',
    fontFace: BOOK, fontSize: 11, color: o.color || WHITE
  });
}

/** Small orange rounded square used for step numbers / icon chips. */
function badge(slide, o) {
  slide.addShape('roundRect', {
    x: o.x, y: o.y, w: o.s, h: o.s, rectRadius: 0.25074 * o.s,
    fill: { color: o.fill || ORANGE }, line: { type: 'none' }
  });
}

/* --------------------------------------------------- notched "photo" card */

const K = 0.5523; // circle -> cubic bezier magic constant

/**
 * Corner list (clockwise) for a rounded rectangle whose corners may be replaced
 * by an L-shaped "bite".
 *   corner bite  tl|tr|br|bl -> [width, height, filletRadius]
 *   edge bite    tc|bc       -> [xStart, xEnd, depth, filletRadius]
 */
function cardCorners(w, h, r, notch) {
  const n = notch || {};
  const v = [];
  const bite = (key, plain, stepped) => (n[key] ? stepped.apply(null, n[key]) : v.push(plain));

  bite('tl', { x: 0, y: 0, r: r }, (nw, nh, nr) => v.push({ x: 0, y: nh, r: r }, { x: nw, y: nh, r: nr }, { x: nw, y: 0, r: r }));
  if (n.tc) { const [a, b, d, nr] = n.tc; v.push({ x: a, y: 0, r: r }, { x: a, y: d, r: nr }, { x: b, y: d, r: nr }, { x: b, y: 0, r: r }); }
  bite('tr', { x: w, y: 0, r: r }, (nw, nh, nr) => v.push({ x: w - nw, y: 0, r: r }, { x: w - nw, y: nh, r: nr }, { x: w, y: nh, r: r }));
  bite('br', { x: w, y: h, r: r }, (nw, nh, nr) => v.push({ x: w, y: h - nh, r: r }, { x: w - nw, y: h - nh, r: nr }, { x: w - nw, y: h, r: r }));
  if (n.bc) { const [a, b, d, nr] = n.bc; v.push({ x: b, y: h, r: r }, { x: b, y: h - d, r: nr }, { x: a, y: h - d, r: nr }, { x: a, y: h, r: r }); }
  bite('bl', { x: 0, y: h, r: r }, (nw, nh, nr) => v.push({ x: nw, y: h, r: r }, { x: nw, y: h - nh, r: nr }, { x: 0, y: h - nh, r: r }));
  return v;
}

/** Turn a clockwise corner list into a closed custGeom point list. */
function filletedPath(corners) {
  const towards = (from, to, d) => {
    const dx = to.x - from.x;
    const dy = to.y - from.y;
    const len = Math.hypot(dx, dy) || 1;
    return { x: from.x + (dx / len) * Math.min(d, len / 2), y: from.y + (dy / len) * Math.min(d, len / 2) };
  };
  const pts = [];
  corners.forEach((v, i) => {
    const prev = corners[(i - 1 + corners.length) % corners.length];
    const next = corners[(i + 1) % corners.length];
    const a = towards(v, prev, v.r);
    const b = towards(v, next, v.r);
    pts.push(i === 0 ? { x: a.x, y: a.y, moveTo: true } : { x: a.x, y: a.y });
    pts.push({
      x: b.x, y: b.y,
      curve: { type: 'cubic', x1: a.x + K * (v.x - a.x), y1: a.y + K * (v.y - a.y), x2: b.x + K * (v.x - b.x), y2: b.y + K * (v.y - b.y) }
    });
  });
  pts.push({ close: true });
  return pts;
}

/** Rounded card, optionally with corner bites. `o.notch` = { tl:[w,h,r], ... } */
function card(slide, o) {
  slide.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    points: filletedPath(cardCorners(o.w, o.h, o.r === undefined ? 0.25 : o.r, o.notch)),
    fill: { color: o.fill }, line: { type: 'none' },
    shadow: o.shadow ? softShadow() : undefined
  });
}

/** Empty picture placeholder: same silhouette, light grey. */
function photoCard(slide, o) {
  card(slide, Object.assign({ fill: PHOTO }, o));
}

/* -------------------------------------------------------------- line icons */

function stroke(slide, o) {
  slide.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h, points: o.points,
    fill: { type: 'none' }, line: { color: o.color, width: o.width, cap: 'round' }
  });
}

/** Three stacked bars ("menu"). */
function iconMenu(slide, x, y, s) {
  [0.25, 0.5, 0.75].forEach(fy => {
    slide.addShape('roundRect', {
      x: x + 0.11 * s, y: y + (fy - 0.033) * s, w: 0.775 * s, h: 0.066 * s,
      rectRadius: 0.033 * s, fill: { color: ORANGE }, line: { type: 'none' }
    });
  });
}

/** Magnifying glass: ring plus handle. */
function iconSearch(slide, x, y, s) {
  slide.addShape('ellipse', {
    x: x + 0.11 * s, y: y + 0.11 * s, w: 0.55 * s, h: 0.55 * s,
    fill: { type: 'none' }, line: { color: ORANGE, width: 0.085 * s * 72 }
  });
  stroke(slide, {
    x: x, y: y, w: s, h: s, color: ORANGE, width: 0.085 * s * 72,
    points: [{ x: 0.60 * s, y: 0.60 * s, moveTo: true }, { x: 0.85 * s, y: 0.85 * s }]
  });
}

/** Two mountain outlines. */
function iconMountains(slide, x, y, s, color) {
  const w = 0.09 * s * 72;
  stroke(slide, {
    x: x, y: y, w: s, h: s, color: color, width: w,
    points: [{ x: 0.09 * s, y: 0.72 * s, moveTo: true }, { x: 0.34 * s, y: 0.30 * s }, { x: 0.59 * s, y: 0.72 * s }, { close: true }]
  });
  stroke(slide, {
    x: x, y: y, w: s, h: s, color: color, width: w,
    points: [{ x: 0.42 * s, y: 0.72 * s, moveTo: true }, { x: 0.62 * s, y: 0.36 * s }, { x: 0.88 * s, y: 0.72 * s }, { close: true }]
  });
}

/** Sailing boat. */
function iconSailboat(slide, x, y, s, color) {
  const w = 0.075 * s * 72;
  stroke(slide, {
    x: x, y: y, w: s, h: s, color: color, width: w,
    points: [{ x: 0.44 * s, y: 0.15 * s, moveTo: true }, { x: 0.44 * s, y: 0.68 * s }]
  });
  stroke(slide, {
    x: x, y: y, w: s, h: s, color: color, width: w,
    points: [{ x: 0.50 * s, y: 0.22 * s, moveTo: true }, { x: 0.80 * s, y: 0.68 * s }, { x: 0.50 * s, y: 0.68 * s }, { close: true }]
  });
  stroke(slide, {
    x: x, y: y, w: s, h: s, color: color, width: w,
    points: [{ x: 0.14 * s, y: 0.74 * s, moveTo: true }, { x: 0.88 * s, y: 0.74 * s }, { x: 0.74 * s, y: 0.85 * s }, { x: 0.26 * s, y: 0.85 * s }, { close: true }]
  });
}

/** ">" chevron. */
function iconChevron(slide, x, y, s) {
  stroke(slide, {
    x: x, y: y, w: s, h: s, color: ORANGE, width: 0.10 * s * 72,
    points: [{ x: 0.33 * s, y: 0.16 * s, moveTo: true }, { x: 0.66 * s, y: 0.50 * s }, { x: 0.33 * s, y: 0.84 * s }]
  });
}

/** Target rings with a mouse cursor (step 1 of "How to book"). */
function iconTarget(slide, x, y, s, color) {
  [[0.02, 0.96], [0.20, 0.60]].forEach(([o, d]) => {
    slide.addShape('ellipse', {
      x: x + o * s, y: y + o * s, w: d * s, h: d * s,
      fill: { type: 'none' }, line: { color: color, width: 0.075 * s * 72 }
    });
  });
  slide.addShape('custGeom', {
    x: x, y: y, w: s, h: s, fill: { color: color }, line: { color: color, width: 1 },
    points: [
      { x: 0.44 * s, y: 0.40 * s, moveTo: true }, { x: 0.94 * s, y: 0.62 * s },
      { x: 0.72 * s, y: 0.70 * s }, { x: 0.63 * s, y: 0.93 * s }, { close: true }
    ]
  });
}

/** Bullet list with a pencil (step 2 of "How to book"). */
function iconChecklist(slide, x, y, w, h, color) {
  const lw = 0.10 * h * 72;
  [[0.00, 0.62, 0.06], [0.00, 0.45, 0.36], [0.00, 0.62, 0.66]].forEach(([x0, x1, fy]) => {
    stroke(slide, {
      x: x, y: y, w: w, h: h, color: color, width: lw,
      points: [{ x: x0 * w, y: fy * h, moveTo: true }, { x: x1 * w, y: fy * h }]
    });
  });
  stroke(slide, {
    x: x, y: y, w: w, h: h, color: color, width: lw,
    points: [{ x: 0.95 * w, y: 0.18 * h, moveTo: true }, { x: 0.62 * w, y: 0.72 * h }, { x: 0.55 * w, y: 0.92 * h }, { x: 0.74 * w, y: 0.82 * h }, { close: true }]
  });
}

/** Rolling suitcase (step 3 of "How to book"). */
function iconSuitcase(slide, x, y, w, h, color) {
  const lw = 0.055 * h * 72;
  stroke(slide, {
    x: x, y: y, w: w, h: h, color: color, width: lw,
    points: [{ x: 0.36 * w, y: 0.16 * h, moveTo: true }, { x: 0.36 * w, y: 0.00, }, { x: 0.64 * w, y: 0.00 }, { x: 0.64 * w, y: 0.16 * h }]
  });
  slide.addShape('roundRect', {
    x: x, y: y + 0.16 * h, w: w, h: 0.70 * h, rectRadius: 0.06 * w,
    fill: { type: 'none' }, line: { color: color, width: lw }
  });
  [0.30, 0.50, 0.70].forEach(fx => {
    stroke(slide, {
      x: x, y: y, w: w, h: h, color: color, width: lw,
      points: [{ x: fx * w, y: 0.32 * h, moveTo: true }, { x: fx * w, y: 0.70 * h }]
    });
  });
  [0.22, 0.78].forEach(fx => {
    stroke(slide, {
      x: x, y: y, w: w, h: h, color: color, width: lw,
      points: [{ x: fx * w, y: 0.86 * h, moveTo: true }, { x: fx * w, y: 0.99 * h }]
    });
  });
}

/** Oversized faded question-mark speech bubble (FAQ slide watermark). */
function iconQuestionMark(slide, x, y, s) {
  const ring = 0.057 * s * 72;
  slide.addShape('ellipse', {
    x: x + 0.154 * s, y: y + 0.112 * s, w: 0.646 * s, h: 0.646 * s,
    fill: { type: 'none' }, line: { color: CREAM, width: ring }
  });
  stroke(slide, { // speech-bubble tail
    x: x, y: y, w: s, h: s, color: CREAM, width: ring,
    points: [{ x: 0.545 * s, y: 0.735 * s, moveTo: true }, { x: 0.560 * s, y: 0.885 * s }, { x: 0.735 * s, y: 0.700 * s }]
  });
  stroke(slide, { // "?" hook and descender
    x: x, y: y, w: s, h: s, color: CREAM, width: 0.050 * s * 72,
    points: [
      { x: 0.393 * s, y: 0.365 * s, moveTo: true },
      { x: 0.586 * s, y: 0.335 * s, curve: { type: 'cubic', x1: 0.393 * s, y1: 0.225 * s, x2: 0.586 * s, y2: 0.215 * s } },
      { x: 0.490 * s, y: 0.520 * s, curve: { type: 'cubic', x1: 0.586 * s, y1: 0.430 * s, x2: 0.505 * s, y2: 0.440 * s } }
    ]
  });
  slide.addShape('ellipse', { // "?" dot
    x: x + 0.452 * s, y: y + 0.612 * s, w: 0.066 * s, h: 0.066 * s,
    fill: { color: CREAM }, line: { type: 'none' }
  });
}

/* ------------------------------------------------------------- slide 1/10 */

function slideCover(pptx) {
  const s = pptx.addSlide();
  // Layout art: two interlocking grey panels behind the strapline.
  photoCard(s, { x: 1.629, y: 2.668, w: 5.390, h: 1.942, r: 0.345, notch: { br: [0.798, 0.856, 0.200] } });
  photoCard(s, { x: 6.298, y: 2.962, w: 5.390, h: 1.942, r: 0.345, notch: { tl: [0.800, 0.854, 0.206] } });

  text(s, 'Wanderlust', { x: 3.992, y: 0.918, w: 5.35, h: 1.111, fontFace: SERIF, fontSize: 60, color: INK, align: 'center' });
  text(s, 'Journey and Travel Presentation', { x: 3.761, y: 2.029, w: 5.812, h: 0.303, fontFace: MEDIUM, fontSize: 12, charSpacing: 2.5, align: 'center' });
  text(s, 'The sun shone brightly as the birds chirped in the trees. A gentle breeze rustled the leaves, ' +
    'creating a soothing melody. The town bustled with life as people went about their daily routines. ',
  { x: 2.617, y: 5.139, w: 8.099, h: 0.625, fontSize: 11, lineSpacingMultiple: 1.5, align: 'center' });
  pill(s, 'Start Your Journey', { x: 5.479, y: 6.13, w: 2.376, h: 0.397 });

  wordmark(s);
  iconSearch(s, 12.147, 0.424, 0.243);
  iconMenu(s, 12.534, 0.424, 0.243);
  return s;
}

function slideAboutUs(pptx) {
  const s = pptx.addSlide();
  photoCard(s, { x: 2.51, y: 1.128, w: 3.799, h: 4.826, r: 0.355, notch: { tr: [0.908, 0.870, 0.212], bl: [2.391, 1.142, 0.194] } });
  card(s, { x: 1.641, y: 4.917, w: 3.156, h: 1.621, r: 0.315, fill: TEAL });
  text(s, 'Welcome to Wanderlust Adventures, your go-to travel agency for unforgettable journeys.',
    { x: 1.898, y: 5.313, w: 2.642, h: 0.83, color: WHITE, lineSpacingMultiple: 1.5 });

  text(s, 'About Us', { x: 7.485, y: 1.604, w: 3.664, h: 0.909, fontFace: SERIF, fontSize: 48, color: INK });
  text(s, 'Your Trusted Travel Partner', { x: 7.485, y: 3.379, w: 3.283, h: 0.303, fontFace: MEDIUM, fontSize: 12, color: INK });
  text(s, "PLACEHOLDER" +
    "tailored to your desires. Let us be your guide in discovering the world's wonders.",
  { x: 7.485, y: 3.979, w: 3.575, h: 1.082, lineSpacingMultiple: 1.5 });

  // Five-star rating on a translucent grey pill.
  s.addShape('roundRect', {
    x: 7.523, y: 5.461, w: 1.938, h: 0.455, rectRadius: 0.2275,
    fill: { color: PHOTO, transparency: 65 }, line: { type: 'none' }
  });
  for (let i = 0; i < 5; i++) {
    s.addShape('star5', { x: 7.743 + i * 0.3065, y: 5.542, w: 0.269, h: 0.269, fill: { color: ORANGE }, line: { type: 'none' } });
  }

  wordmark(s);
  footerUrl(s);
  return s;
}

function slideExploring(pptx) {
  const s = pptx.addSlide();
  text(s, 'Exploring The World', { x: 3.512, y: 1.056, w: 6.309, h: 0.774, fontFace: SERIF, fontSize: 40, color: INK, align: 'center' });
  text(s, 'Discover Enhancing Destinations', { x: 4.662, y: 2.093, w: 4.01, h: 0.303, fontFace: MEDIUM, fontSize: 12, color: INK, align: 'center' });

  // Four photo slots, each with a caption chip straddling its lower edge.
  const spots = [
    { x: 1.548, chip: TEAL, label: 'Bali, Indonesia', labelColor: WHITE, cx: 2.107, tx: 2.072, tw: 1.303 },
    { x: 4.177, chip: TEAL, label: 'Bangkok, Thailand', labelColor: WHITE, cx: 4.736, tx: 4.701, tw: 1.303 },
    { x: 6.805, chip: WHITE, label: 'Hawaii, USA', labelColor: TEAL, cx: 7.364, tx: 7.479, tw: 1.004 },
    { x: 9.434, chip: TEAL, label: 'Tokyo, Japan', labelColor: WHITE, cx: 9.993, tx: 9.993, tw: 1.233 }
  ];
  spots.forEach(spot => {
    // The chip sits in a shallow bite along the bottom edge of the photo slot.
    photoCard(s, { x: spot.x, y: 2.984, w: 2.35, h: 2.646, r: 0.325, notch: { bc: [0.520, 1.833, 0.315, 0.255] } });
    card(s, { x: spot.cx, y: 5.385, w: 1.233, h: 0.794, r: 0.24, fill: spot.chip, shadow: spot.chip === WHITE });
    text(s, spot.label, { x: spot.tx, y: 5.53, w: spot.tw, h: 0.505, fontFace: MEDIUM, fontSize: 12, color: spot.labelColor, align: 'center' });
  });

  wordmark(s);
  footerUrl(s);
  return s;
}

function slidePackages(pptx) {
  const s = pptx.addSlide();
  text(s, 'Our Travel Package', { x: 1.63, y: 1.867, w: 3.664, h: 1.582, fontFace: SERIF, fontSize: 44, color: INK });
  text(s, 'Customized Travel Packages', { x: 1.63, y: 3.951, w: 3.283, h: 0.303, fontFace: MEDIUM, fontSize: 12, color: INK });
  text(s, "PLACEHOLDER" +
    "tailored to your desires. Let us be your guide in discovering the world's wonders.",
  { x: 1.63, y: 4.551, w: 3.404, h: 1.082, lineSpacingMultiple: 1.5 });

  // Three staggered price cards; the corner bite hosts the step number badge.
  // `tx`/`ty` is the top-left of the tier's own (independently placed) text block.
  const tiers = [
    { name: 'Bronze', price: ['10', '.5k'], x: 5.8, y: 1.146, notch: 'tr', fill: TEAL, ink: WHITE, tx: 6.074, ty: 2.181, badge: { n: '01', x: 7.137, y: 1.157 } },
    { name: 'Silver', price: ['13', '.2k'], x: 7.823, y: 2.267, notch: 'bl', fill: TEAL, ink: WHITE, tx: 8.097, ty: 2.787, badge: { n: '02', x: 7.786, y: 5.392 } },
    { name: 'Gold', price: ['15', '.3k'], x: 9.858, y: 1.466, notch: 'br', fill: WHITE, ink: TEAL, tx: 10.133, ty: 1.986, badge: { n: '03', x: 11.196, y: 4.593 } }
  ];
  tiers.forEach(t => {
    const notch = {};
    notch[t.notch] = [0.643, 0.684, 0.176];
    card(s, { x: t.x, y: t.y, w: 1.932, h: 3.766, r: 0.325, fill: t.fill, notch: notch, shadow: t.fill === WHITE });

    text(s, t.name, { x: t.tx, y: t.ty, w: 1.383, h: 0.37, fontFace: MEDIUM, fontSize: 16, color: t.ink, align: 'center' });
    text(s, [
      { text: '$', options: { fontSize: 24, superscript: true } },
      { text: t.price[0], options: { fontSize: 24 } },
      { text: t.price[1], options: { fontSize: 12 } }
    ], { x: t.tx, y: t.ty + 0.493, w: 1.383, h: 0.505, fontFace: MEDIUM, color: t.ink, align: 'center' });
    text(s, 'Package detail can be written here, as you wish, as you want. ',
      { x: t.tx - 0.092, y: t.ty + 1.082, w: 1.567, h: 1.082, color: t.fill === WHITE ? GREY : WHITE, lineSpacingMultiple: 1.5, align: 'center' });

    badge(s, { x: t.badge.x, y: t.badge.y, s: 0.623 });
    text(s, t.badge.n, { x: t.badge.x - 0.025, y: t.badge.y + 0.113, w: 0.673, h: 0.404, fontFace: MEDIUM, fontSize: 18, color: WHITE, align: 'center' });
  });

  wordmark(s);
  footerUrl(s);
  return s;
}

function slideTestimonials(pptx) {
  const s = pptx.addSlide();
  photoCard(s, { x: 2.051, y: 3.065, w: 2.4, h: 2.896, r: 0.320, notch: { tl: [1.934, 0.576, 0.197] } });
  photoCard(s, { x: 4.533, y: 1.739, w: 2.4, h: 2.896, r: 0.320, notch: { br: [1.930, 0.579, 0.191] } });

  const quotes = [
    { x: 1.424, y: 1.739, tx: 1.623, ty: 2.020, quote: '\u201CThe trip exceeded all expectations! Every detail was thoughtfully arranged."', who: 'Sarah Doe' },
    { x: 5.039, y: 4.121, tx: 5.272, ty: 4.401, quote: '\u201CThanks to Wanderlust Adventures, I had the best vacation of my life!"', who: 'Mark Doe' }
  ];
  quotes.forEach(q => {
    card(s, { x: q.x, y: q.y, w: 2.522, h: 1.841, r: 0.310, fill: TEAL });
    text(s, q.quote, { x: q.tx, y: q.ty, w: 2.249, h: 0.83, color: WHITE, lineSpacingMultiple: 1.5 });
    text(s, q.who, { x: q.tx, y: q.ty + 0.977, w: 1.526, h: 0.303, fontFace: SEMI, bold: true, fontSize: 12, color: WHITE });
  });

  text(s, 'Customers Testimonials', { x: 8.277, y: 2.02, w: 3.899, h: 1.447, fontFace: SERIF, fontSize: 40, color: INK });
  text(s, 'Hear from Our Travelers', { x: 8.277, y: 4.031, w: 3.283, h: 0.303, fontFace: MEDIUM, fontSize: 12, color: INK });
  text(s, 'PLACEHOLDER' +
    'experiences with our services, expressing their genuine delight and satisfaction.',
  { x: 8.277, y: 4.568, w: 3.575, h: 1.082, lineSpacingMultiple: 1.5 });

  wordmark(s);
  footerUrl(s);
  return s;
}

function slidePortfolio(pptx) {
  const s = pptx.addSlide();
  // Six-slot photo mosaic on the right.
  const mosaic = [
    { x: 6.687, y: 0.802, w: 1.627, h: 2.715, r: 0.290, notch: { tl: [0.436, 0.414, 0.221] } },
    { x: 8.405, y: 0.802, w: 1.635, h: 3.044, r: 0.310, notch: { tr: [0.434, 0.461, 0.233] } },
    { x: 10.131, y: 0.805, w: 1.623, h: 2.364, r: 0.300, notch: { br: [0.435, 0.441, 0.221] } },
    { x: 6.697, y: 3.603, w: 1.615, h: 3.095, r: 0.270, notch: { tr: [0.429, 0.412, 0.230] } },
    { x: 8.405, y: 3.932, w: 1.639, h: 2.767, r: 0.325, notch: { bl: [0.434, 0.500, 0.245] } },
    { x: 10.131, y: 3.267, w: 1.650, h: 2.914, r: 0.295, notch: { bl: [0.442, 0.446, 0.227] } }
  ];
  mosaic.forEach(m => photoCard(s, m));

  text(s, 'Portfolio', { x: 1.965, y: 1.936, w: 3.664, h: 1.111, fontFace: SERIF, fontSize: 60, color: INK });
  text(s, 'Photo Compilation of Our Services', { x: 1.965, y: 3.617, w: 3.283, h: 0.303, fontFace: MEDIUM, fontSize: 12, color: INK });
  text(s, 'PLACEHOLDER' +
    'impeccable services and travel experiences.',
  { x: 1.965, y: 4.123, w: 3.575, h: 0.83, lineSpacingMultiple: 1.5 });
  pill(s, 'Our Gallery', { x: 2.053, y: 5.369, w: 1.704, h: 0.397 });

  badge(s, { x: 11.518, y: 4.782, s: 0.534 });
  iconMountains(s, 11.606, 4.87, 0.357, WHITE);

  wordmark(s);
  footerUrl(s);
  return s;
}

function slideTraditions(pptx) {
  const s = pptx.addSlide();
  photoCard(s, { x: 2.205, y: 1.836, w: 3.893, h: 1.883, r: 0.315, notch: { tl: [0.969, 0.683, 0.179] } });
  photoCard(s, {
    x: 6.815, y: 4.151, w: 4.882, h: 1.998, r: 0.365,
    notch: { tl: [0.585, 0.513, 0.224], tr: [0.469, 0.413, 0.260], br: [0.475, 0.415, 0.260] }
  });

  card(s, { x: 1.642, y: 1.266, w: 1.479, h: 1.2, r: 0.270, fill: TEAL });
  text(s, 'Immerse Yourself in Culture', { x: 1.795, y: 1.513, w: 1.175, h: 0.707, fontFace: MEDIUM, fontSize: 12, color: WHITE, align: 'center' });

  text(s, 'Embrace Local Traditions', { x: 7.109, y: 2.034, w: 4.448, h: 1.582, fontFace: SERIF, fontSize: 44, color: INK });
  text(s, 'At Wanderlust Adventures, we believe in fostering meaningful connections with the destinations ' +
    'you visit. Immerse yourself in local cultures through authentic experiences, culinary delights, and ' +
    'interactions with friendly locals.',
  { x: 2.131, y: 4.587, w: 3.893, h: 1.339, lineSpacingMultiple: 1.5 });

  wordmark(s);
  footerUrl(s);
  return s;
}

function slideAdventure(pptx) {
  const s = pptx.addSlide();
  photoCard(s, { x: 6.852, y: 1.434, w: 5.019, h: 2.065, r: 0.320, notch: { tl: [0.597, 0.554, 0.188], br: [0.609, 0.559, 0.206] } });
  photoCard(s, { x: 5.881, y: 3.613, w: 5.06, h: 2.062, r: 0.320, notch: { br: [2.253, 1.270, 0.179] } });

  text(s, 'Adventure Tours', { x: 1.652, y: 2.031, w: 3.671, h: 1.582, fontFace: SERIF, fontSize: 44, color: INK });
  text(s, 'Adventures for The Brave', { x: 1.652, y: 4.114, w: 3.671, h: 0.303, fontFace: MEDIUM, fontSize: 12, color: INK });
  text(s, 'For adrenaline seekers, we offer thrilling adventure tours that push boundaries and create ' +
    'memories that last a lifetime. ',
  { x: 1.652, y: 4.639, w: 3.526, h: 0.83, lineSpacingMultiple: 1.5 });

  card(s, { x: 8.775, y: 4.48, w: 3.096, h: 1.712, r: 0.315, fill: TEAL });
  text(s, 'From exhilarating hikes and safaris to daring water sports, our expertly crafted tours guarantee ' +
    'an extraordinary rush.',
  { x: 9.001, y: 4.795, w: 2.644, h: 1.082, color: WHITE, lineSpacingMultiple: 1.5 });

  badge(s, { x: 6.756, y: 1.291, s: 0.623 });
  iconSailboat(s, 6.844, 1.379, 0.446, WHITE);

  wordmark(s);
  footerUrl(s);
  return s;
}

function slideHowToBook(pptx) {
  const s = pptx.addSlide();
  text(s, 'How to Book with Us', { x: 3.232, y: 1.121, w: 6.87, h: 0.841, fontFace: SERIF, fontSize: 44, color: INK, align: 'center' });
  text(s, 'Start Your Journey with Us Today', { x: 4.037, y: 2.099, w: 5.259, h: 0.303, fontFace: MEDIUM, fontSize: 12, color: INK, align: 'center' });

  const steps = [
    {
      x: 1.553, y: 3.072, w: 2.483, h: 3.010, r: 0.335, fill: WHITE, ink: GREY,
      notch: { tl: [0.494, 0.486, 0.245], br: [0.479, 0.480, 0.260] },
      body: 'Visit our website or contact our friendly travel consultants to discuss your travel aspirations.\u00A0',
      tx: 1.812, ty: 4.165, tw: 1.911, th: 1.335
    },
    {
      x: 5.427, y: 3.072, w: 2.479, h: 3.006, r: 0.320, fill: TEAL, ink: WHITE,
      notch: { tl: [0.487, 0.491, 0.266], tr: [0.478, 0.485, 0.236] },
      body: 'Choosing the perfect itinerary and tailor it to your preferences.',
      tx: 5.622, ty: 4.438, tw: 2.09, th: 0.83
    },
    {
      x: 9.296, y: 3.049, w: 2.484, h: 3.032, r: 0.320, fill: TEAL, ink: WHITE,
      notch: { tl: [0.680, 0.684, 0.191] },
      body: 'Enjoy your traveling and get amazing experiences with our plan',
      tx: 9.521, ty: 4.438, tw: 2.09, th: 0.83
    }
  ];
  steps.forEach(st => {
    card(s, { x: st.x, y: st.y, w: st.w, h: st.h, r: st.r, fill: st.fill, notch: st.notch, shadow: st.fill === WHITE });
    text(s, st.body, { x: st.tx, y: st.ty, w: st.tw, h: st.th, color: st.ink, lineSpacingMultiple: 1.5, align: 'center' });
  });

  iconTarget(s, 2.584, 3.589, 0.427, TEAL);
  iconChecklist(s, 6.472, 3.871, 0.389, 0.321, WHITE);
  iconSuitcase(s, 10.42, 3.823, 0.292, 0.417, WHITE);

  [4.377, 8.251].forEach(x => {
    iconChevron(s, x, 4.356, 0.406);
    iconChevron(s, x + 0.219, 4.356, 0.406);
  });

  wordmark(s);
  footerUrl(s);
  return s;
}

function slideFaq(pptx) {
  const s = pptx.addSlide();
  photoCard(s, {
    x: 5.007, y: 0.667, w: 3.319, h: 6.167, r: 0.345,
    notch: { tl: [0.846, 0.795, 0.206], tr: [0.703, 0.572, 0.242], br: [0.688, 1.321, 0.200] }
  });
  iconQuestionMark(s, 10.237, 1.517, 2.737);

  text(s, 'Frequently Asked Question', { x: 8.814, y: 2.69, w: 2.952, h: 2.121, fontFace: SERIF, fontSize: 40, color: INK });

  const faqs = [
    {
      y: 1.801, fill: TEAL, shadow: false, r: 0.325,
      notch: { tl: [0.482, 0.448, 0.260], br: [0.468, 0.447, 0.236] },
      q: 'How do travel agencies save time and money for travellers?', qColor: WHITE, qy: 2.071, qw: 2.798,
      a: 'They have access to exclusive deals, compare prices, and handle all planning and booking details.',
      aColor: WHITE, ay: 2.725, aw: 2.951
    },
    {
      y: 4.236, fill: WHITE, shadow: true, r: 0.335,
      notch: { tl: [0.472, 0.437, 0.215], bl: [0.472, 0.442, 0.230], br: [0.484, 0.453, 0.260] },
      q: 'Why use a travel agency for complex or international trips?', qColor: TEAL, qy: 4.506, qw: 2.951,
      a: 'They provide expertise on destinations, handle visas, and create detailed itineraries for a smoother trip.',
      aColor: GREY, ay: 5.16, aw: 3.128
    }
  ];
  faqs.forEach(f => {
    card(s, { x: 1.456, y: f.y, w: 4.135, h: 2.089, r: f.r, fill: f.fill, notch: f.notch, shadow: f.shadow });
    text(s, f.q, { x: 2.116, y: f.qy, w: f.qw, h: 0.577, fontFace: MEDIUM, color: f.qColor, lineSpacingMultiple: 1.5 });
    text(s, f.a, { x: 2.116, y: f.ay, w: f.aw, h: 0.83, color: f.aColor, lineSpacingMultiple: 1.5 });
  });

  wordmark(s);
  footerUrl(s);
  return s;
}

/* -------------------------------------------------------------------- run */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE_16x9', width: 13.3333, height: 7.5 });
  pptx.layout = 'WIDE_16x9';
  pptx.title = 'Wanderlust';

  slideCover(pptx);
  slideAboutUs(pptx);
  slideExploring(pptx);
  slidePackages(pptx);
  slideTestimonials(pptx);
  slidePortfolio(pptx);
  slideTraditions(pptx);
  slideAdventure(pptx);
  slideHowToBook(pptx);
  slideFaq(pptx);

  return pptx.writeFile({ fileName: path.join(__dirname, '0af8df23-bb56-4514-82fe-14b4d6671835_grok_final.pptx') });
}

build().then(f => console.log('wrote ' + f)).catch(err => { console.error(err); process.exit(1); });
