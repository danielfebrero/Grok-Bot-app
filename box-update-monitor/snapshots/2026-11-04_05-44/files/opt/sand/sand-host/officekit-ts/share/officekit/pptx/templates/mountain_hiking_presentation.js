/**
 * HorizonTrail - Mountain & Hiking Presentation (20 slides, 13.333 x 7.5 in)
 * Rebuilt with pptxgenjs. Raster icons/photos from the source deck are replaced
 * by native shape placeholders (see `mark()` and `photoBox()`).
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const GREEN = '465337';
const YELLOW = 'FFDA22';
const INK = '212121';
const INK85 = '262626';
const WHITE = 'FFFFFF';
const BLACK = '000000';
const OVERLAY = '343E29'; // accent1 darkened 75%, painted at 50% transparency
const FONT = 'Open Sans';

const NONE = { type: 'none' };

// pptxgenjs rewrites shadow objects in place while serialising, so every shape
// needs its own copy. blur/offset are points, angle degrees, opacity 0..1.
const shadow = (blur, offset, opacity) => ({ type: 'outer', blur, offset, angle: 45, color: BLACK, opacity });
const SHADOW = () => shadow(15, 3, 0.2); // tight contact shadow
const SOFT = () => shadow(20, 15, 0.1); // wide drop shadow
const WIDE = () => shadow(20, 20, 0.1);

/* ------------------------------------------------------------------ helpers */
function txt(slide, text, o) {
  slide.addText(text, Object.assign({ fontFace: FONT, valign: 'top', isTextBox: true }, o));
}

function shape(slide, kind, o) {
  slide.addShape(kind, Object.assign({ line: NONE }, o));
}

function rect(slide, x, y, w, h, fill, extra) {
  shape(slide, 'rect', Object.assign({ x, y, w, h, fill }, extra));
}

function oval(slide, x, y, w, h, color, extra) {
  shape(slide, 'ellipse', Object.assign({ x, y, w, h, fill: { color } }, extra));
}

function round(slide, x, y, w, h, color, radius, extra) {
  shape(slide, 'roundRect', Object.assign({ x, y, w, h, fill: { color }, rectRadius: radius }, extra));
}

/** Stand-in emblem for the deck's icon artwork: the HorizonTrail twin-peak mountain. */
function mark(slide, cx, cy, size, color) {
  const w = size;
  const h = size * 0.62;
  const x = cx - w / 2;
  const y = cy - h / 2;
  shape(slide, 'triangle', { x, y, w: w * 0.64, h, fill: { color } });
  shape(slide, 'triangle', { x: x + w * 0.36, y: y + h * 0.24, w: w * 0.64, h: h * 0.76, fill: { color } });
}

/**
 * Stand-in for a picture frame. The reference deck's picture placeholders are
 * empty (no image is embedded, they render as bare slide background), so the
 * frame is drawn as a hairline outline only — same box, no invented content.
 */
function photoBox(slide, x, y, w, h, radius) {
  shape(slide, 'roundRect', {
    x, y, w, h, rectRadius: radius || 0.12,
    fill: { color: WHITE, transparency: 100 },
    line: { color: 'E6E6E6', width: 0.75 },
  });
}

/** Mix `color` into white at the given opacity and return the flat hex result. */
function blend(color, alpha) {
  let out = '';
  for (let i = 0; i < 6; i += 2) {
    const c = parseInt(color.substr(i, 2), 16) * alpha + 255 * (1 - alpha);
    out += ('0' + Math.round(c).toString(16)).slice(-2);
  }
  return out.toUpperCase();
}

/**
 * Vertical white->color fade, built from stacked opaque bands because pptxgenjs
 * has no gradient fill. Bands are plain rectangles (their top edge is invisible
 * against the previous, nearly identical shade); the rounded silhouette comes
 * from the white base card and the rounded band that finishes the bottom.
 * The fade only begins at `startFrac` down the card.
 */
function fadeCard(slide, x, y, w, h, color, maxAlpha, startFrac, radius, steps) {
  round(slide, x, y, w, h, WHITE, radius);
  // Rounded bottom cap first; the bands drawn after it hide its upper corners.
  round(slide, x, y + h - 2 * radius, w, 2 * radius, blend(color, maxAlpha), radius);
  const top0 = y + h * startFrac;
  const span = h * (1 - startFrac);
  for (let i = 0; i < steps; i++) {
    const top = top0 + (span * i) / steps;
    rect(slide, x, top, w, y + h - radius - top, { color: blend(color, (maxAlpha * (i + 1)) / steps) });
  }
}

/* --------------------------------------------------------- shared page furniture */
const NAV = [
  ['Home', 4.757, 0.756],
  ['About Us', 5.734, 1.01],
  ['Services', 6.966, 0.955],
  ['Explore', 8.142, 0.831],
];

function chrome(slide, footerColor) {
  rect(slide, 0, 0, 13.333, 0.844, { color: GREEN });
  mark(slide, 0.788, 0.378, 0.372, WHITE);
  txt(slide, 'HorizonTrail', { x: 0.975, y: 0.286, w: 1.236, h: 0.303, fontSize: 12, bold: true, color: WHITE });
  NAV.forEach(([label, x, w]) =>
    txt(slide, label, { x, y: 0.286, w, h: 0.303, fontSize: 12, bold: true, color: WHITE, align: 'center' })
  );
  txt(slide, 'Contact Us', { x: 11.749, y: 0.286, w: 1.114, h: 0.303, fontSize: 12, bold: true, color: YELLOW, align: 'center' });
  txt(slide, 'www.horizontrail.com', { x: 0.492, y: 6.661, w: 2.074, h: 0.303, fontSize: 12, bold: true, color: footerColor });
  txt(slide, 'Presentation Template', { x: 10.736, y: 6.665, w: 2.132, h: 0.303, fontSize: 12, bold: true, color: footerColor, align: 'right' });
}

/** Section heading used on every infographic slide. */
function heading(slide, text, x, w) {
  txt(slide, text, { x, y: 1.15, w, h: 0.64, fontSize: 32, bold: true, color: GREEN, align: 'center' });
}

const BODY = { fontSize: 12, color: INK, lineSpacingMultiple: 1.5 };

/* =================================================================== slides */

// 1 - Cover
function slide01(pptx) {
  const s = pptx.addSlide();
  rect(s, 0, 0.844, 13.333, 6.656, { color: OVERLAY, transparency: 50 });
  chrome(s, WHITE);
  mark(s, 6.666, 2.281, 0.981, WHITE);
  txt(s, 'HorizonTrail', { x: 3.333, y: 2.534, w: 6.667, h: 1.313, fontSize: 72, bold: true, color: WHITE, align: 'center' });
  txt(s, 'Mountain & Hiking Presentation ', {
    x: 3.987, y: 3.783, w: 5.359, h: 0.337, fontSize: 14, bold: true, color: WHITE, align: 'center', charSpacing: 3,
  });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut', {
    x: 4.458, y: 4.166, w: 4.417, h: 0.675, fontSize: 12, color: WHITE, align: 'center', lineSpacingMultiple: 1.5,
  });
  round(s, 5.389, 5.045, 2.555, 0.494, GREEN, 0.147);
  txt(s, 'Start Presentation', { x: 5.587, y: 5.14, w: 2.159, h: 0.303, fontSize: 12, bold: true, color: WHITE, align: 'center' });
}

// 2 - Why Hiking is the Ultimate Escape
function slide02(pptx) {
  const s = pptx.addSlide();
  chrome(s, GREEN);
  txt(s, 'Why Hiking is the Ultimate Escape', { x: 3.505, y: 1.638, w: 4.835, h: 1.178, fontSize: 32, bold: true, color: GREEN });
  txt(s, 'Benefits of Exploring the Great Outdoors', { x: 3.505, y: 2.873, w: 6.433, h: 0.337, fontSize: 14, bold: true, color: GREEN });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim',
    Object.assign({ x: 3.505, y: 3.218, w: 4.175, h: 0.978 }, BODY));

  round(s, 0.602, 4.645, 4.175, 1.688, GREEN, 0.237);
  txt(s, '\u201CIn every walk with nature, one receives far more than he seeks.\u201D', {
    x: 0.905, y: 4.829, w: 3.803, h: 0.866, fontSize: 16, italic: true, color: WHITE, fontFace: 'Open Sans SemiBold', lineSpacingMultiple: 1.5,
  });
  shape(s, 'line', { x: 1.0, y: 5.849, w: 0.202, h: 0, line: { color: WHITE, width: 1.25 } });
  txt(s, 'John Muir', { x: 1.202, y: 5.602, w: 1.75, h: 0.417, fontSize: 14, bold: true, color: WHITE, lineSpacingMultiple: 1.5 });

  txt(s, 'Key Benefits from Hiking', { x: 5.167, y: 4.841, w: 3.026, h: 0.337, fontSize: 14, bold: true, color: GREEN });
  [['Improves cardiovascular fitness', 5.257, 5.316, GREEN, 5.649],
   ['Encourages bonding with friends or family', 5.737, 5.796, YELLOW, 4.829]].forEach(([label, ty, sy, color, w]) => {
    rect(s, 5.292, sy, 0.201, 0.201, { color });
    txt(s, label, { x: 5.621, y: ty, w, h: 0.303, fontSize: 12, color: BLACK });
  });

  round(s, 8.123, 3.197, 1.973, 1.604, WHITE, 0.2, { line: { color: GREEN, width: 1.5 } });
  txt(s, '20K+', { x: 8.309, y: 3.354, w: 1.67, h: 0.707, fontSize: 36, bold: true, color: GREEN, align: 'center' });
  txt(s, 'Experience These Benefits', {
    x: 8.231, y: 3.876, w: 1.825, h: 0.77, fontSize: 14, bold: true, color: GREEN, align: 'center', lineSpacingMultiple: 1.5,
  });
}

// 3 - Top Hiking Destinations
function slide03(pptx) {
  const s = pptx.addSlide();
  const cards = [
    { x: 4.251, y: 1.329, w: 2.701, h: 4.842, name: 'Inca Trail, Peru', tw: 1.739, tx: 4.432, bx: 4.44, ax: 6.584 },
    { x: 7.158, y: 1.25, w: 2.673, h: 4.923, name: 'Everest Trek, Nepal', tw: 3.05, tx: 7.311, bx: 7.318, ax: 9.462 },
    { x: 10.046, y: 1.336, w: 2.685, h: 4.836, name: 'Mount Fuji, Japan', tw: 2.219, tx: 10.189, bx: 10.197, ax: 12.341 },
  ];
  cards.forEach((c) => fadeCard(s, c.x, c.y, c.w, c.h, GREEN, 0.8, 0.42, 0.3, 24));
  chrome(s, GREEN);
  txt(s, 'Top Hiking Destinations Around the World', { x: 0.602, y: 2.584, w: 3.649, h: 2.255, fontSize: 32, bold: true, color: GREEN });
  cards.forEach((c) => {
    txt(s, c.name, { x: c.tx, y: 4.985, w: c.tw, h: 0.337, fontSize: 14, bold: true, color: WHITE });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, ', {
      x: c.bx, y: 5.254, w: 2.444, h: 0.675, fontSize: 12, color: WHITE, lineSpacingMultiple: 1.5,
    });
    oval(s, c.ax - 0.051, 1.493, 0.315, 0.315, WHITE, { fill: { color: WHITE, transparency: 20 } });
    shape(s, 'line', { x: c.ax, y: 1.526, w: 0.25, h: 0.25, flipV: true, line: { color: INK, width: 2, endArrowType: 'triangle' } });
  });
}

// 4 - Ultimate Guide to Prepping for a Hike
function slide04(pptx) {
  const s = pptx.addSlide();
  chrome(s, GREEN);
  photoBox(s, 0.658, 1.298, 3.289, 5.083, 0.3);
  photoBox(s, 8.152, 1.25, 4.523, 3.472, 0.2);
  txt(s, 'Ultimate Guide to Prepping for a Hike', { x: 8.334, y: 4.943, w: 4.63, h: 1.178, fontSize: 32, bold: true, color: GREEN });

  const steps = [
    { y: 1.75, ox: 4.478, tx: 4.922, tw: 2.77, dot: GREEN, glyph: WHITE, title: 'Choose Your Trail Wisely' },
    { y: 2.971, ox: 4.46, tx: 4.904, tw: 2.929, dot: YELLOW, glyph: GREEN, title: 'Plan Your Route' },
    { y: 4.192, ox: 4.443, tx: 4.887, tw: 2.77, dot: GREEN, glyph: WHITE, title: 'Pack the Essentials' },
  ];
  steps.forEach((it) => {
    oval(s, it.ox, it.y, 0.411, 0.411, it.dot);
    mark(s, it.ox + 0.206, it.y + 0.206, 0.2, it.glyph);
    txt(s, it.title, { x: it.tx, y: it.y, w: it.tw, h: 0.337, fontSize: 14, bold: true, color: GREEN });
    txt(s, 'Lorem ipsum dolor sit amet, adipiscing elit, sed do eiusmod',
      Object.assign({ x: it.tx, y: it.y + 0.273, w: 3.176, h: 0.675 }, BODY));
  });

  round(s, 4.922, 5.414, 1.689, 0.485, YELLOW, 0.081);
  txt(s, 'See More', { x: 5.14, y: 5.488, w: 1.252, h: 0.337, fontSize: 14, bold: true, color: GREEN, align: 'center' });
}

// 5 - Essential Survival and Safety Tips
function slide05(pptx) {
  const s = pptx.addSlide();
  chrome(s, GREEN);
  photoBox(s, 9.762, 1.25, 3.007, 5.125, 0.25);
  txt(s, 'Essential Survival and Safety Tips for Hikers', { x: 0.606, y: 1.341, w: 5.743, h: 1.178, fontSize: 32, bold: true, color: GREEN });
  txt(s, 'Prepared On With Hiking Safety Tips', { x: 0.626, y: 2.609, w: 6.433, h: 0.337, fontSize: 14, bold: true, color: GREEN });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut',
    Object.assign({ x: 0.626, y: 2.892, w: 4.689, h: 0.675 }, BODY));

  const tips = [
    { bx: 0.732, tx: 0.918, y: 3.799, tw: 2.333, color: GREEN, title: 'Stay on Marked Trails' },
    { bx: 3.81, tx: 3.997, y: 3.799, tw: 2.996, color: YELLOW, title: 'Pack Emergency' },
    { bx: 0.732, tx: 0.918, y: 5.197, tw: 2.333, color: YELLOW, title: 'Wildlife Awareness' },
    { bx: 3.81, tx: 3.997, y: 5.197, tw: 2.333, color: GREEN, title: 'Know Basic First Aid' },
  ];
  tips.forEach((t) => {
    round(s, t.bx, t.y + 0.056, 0.064, 0.926, t.color, 0.032);
    txt(s, t.title, { x: t.tx, y: t.y, w: t.tw, h: 0.337, fontSize: 14, bold: true, color: t.color });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing',
      Object.assign({ x: t.tx, y: t.y + 0.351, w: 2.588, h: 0.675 }, BODY));
  });

  const cards = [
    { y: 1.254, bg: YELLOW, dot: GREEN, title: 'Knowing Basic Survival Skills', tc: GREEN, bc: INK },
    { y: 3.905, bg: GREEN, dot: YELLOW, title: 'Preparation Is Key Save Yourself', tc: YELLOW, bc: WHITE },
  ];
  cards.forEach((c) => {
    round(s, 7.229, c.y, 2.3, 2.47, c.bg, 0.211);
    round(s, 7.405, c.y + 0.218, 0.406, 0.094, c.dot, 0.047);
    oval(s, 7.892, c.y + 0.218, 0.094, 0.094, c.dot);
    oval(s, 8.97, c.y + 0.105, 0.424, 0.424, c.dot);
    txt(s, c.title, { x: 7.329, y: c.y + 0.583, w: 2.076, h: 0.572, fontSize: 14, bold: true, color: c.tc });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do', {
      x: 7.329, y: c.y + 1.137, w: 1.917, h: 0.978, fontSize: 12, color: c.bc, lineSpacingMultiple: 1.5,
    });
  });
}

// 6 - Eco-Conscious Hiking
function slide06(pptx) {
  const s = pptx.addSlide();
  chrome(s, GREEN);
  photoBox(s, 5.148, 1.247, 3.134, 5.128, 0.25);
  rect(s, 0.772, 1.25, 0.067, 5.125, { color: YELLOW, transparency: 32 });
  [['Stick to the Trails', 1.56, 2.078],
   ['Minimize Campfire Impact', 3.137, 3.644],
   ['Conserve Water Sources', 4.714, 5.207]].forEach(([title, ty, dy]) => {
    oval(s, 0.688, dy, 0.219, 0.219, GREEN);
    txt(s, title, { x: 1.229, y: ty, w: 2.833, h: 0.421, fontSize: 14, bold: true, color: GREEN, lineSpacingMultiple: 1.5 });
    txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipisci elit. Maecenas porttitor congue.',
      Object.assign({ x: 1.229, y: ty + 0.43, w: 3.465, h: 0.678 }, BODY));
  });

  txt(s, 'Eco-Conscious Hiking for Nature Lovers', { x: 8.752, y: 1.244, w: 4.258, h: 1.919, fontSize: 36, bold: true, color: GREEN });
  txt(s, 'Leave No Trace Behind Nature', { x: 8.768, y: 3.196, w: 5.153, h: 0.337, fontSize: 14, bold: true, color: GREEN });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore',
    Object.assign({ x: 8.768, y: 3.533, w: 3.764, h: 0.978 }, BODY));

  round(s, 8.881, 4.75, 3.134, 1.625, GREEN, 0.179);
  round(s, 9.229, 4.952, 0.518, 0.518, WHITE, 0.086);
  mark(s, 9.488, 5.211, 0.336, GREEN);
  txt(s, 'Wildlife', { x: 9.784, y: 4.853, w: 1.861, h: 0.417, fontSize: 14, bold: true, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Respect', { x: 9.784, y: 5.132, w: 1.861, h: 0.379, fontSize: 12, color: WHITE, lineSpacingMultiple: 1.5 });
  txt(s, 'Lorem  ipsum  dolor sit  amet consect adipiscing elital', {
    x: 9.129, y: 5.501, w: 3.009, h: 0.675, fontSize: 12, color: WHITE, lineSpacingMultiple: 1.5,
  });
}

// 7 - Collection of Hiking Memories
function slide07(pptx) {
  const s = pptx.addSlide();
  chrome(s, GREEN);
  [[-0.025, 3.297, 2.079, 2.646], [2.289, 2.98, 2.698, 3.279], [5.222, 3.297, 2.889, 2.646],
   [8.346, 2.98, 2.698, 3.279], [11.279, 3.297, 2.079, 2.646]].forEach(([x, y, w, h]) => photoBox(s, x, y, w, h, 0.18));
  txt(s, 'Exploring A Collection of Hiking Memories', { x: 0.674, y: 1.341, w: 5.477, h: 1.178, fontSize: 32, bold: true, color: GREEN });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation',
    Object.assign({ x: 7.539, y: 1.366, w: 5.302, h: 0.978 }, BODY));
}

// 8 - Personal Hiking Experiences
function slide08(pptx) {
  const s = pptx.addSlide();
  chrome(s, GREEN);
  const cards = [
    { cx: 0.987, tx: 1.187, px: 3.379, dx: 3.866, bg: GREEN, accent: YELLOW, quoteColor: WHITE,
      name: 'Logan Spencer', quote: '\u201CThe solitude of the mountains felt freeing, just endless trail ahead.\u201D', photo: 1.15 },
    { cx: 9.028, tx: 9.228, px: 11.42, dx: 11.906, bg: YELLOW, accent: GREEN, quoteColor: INK,
      name: 'Lily Charlotte', quote: '\u201CHiking let me clear my mind, with only nature as my companion\u201D', photo: 9.201 },
  ];
  cards.forEach((c) => {
    round(s, c.cx, 1.584, 3.306, 4.456, c.bg, 0.208);
    photoBox(s, c.photo, 1.679, 2.977, 3.056, 0.18);
    txt(s, c.name, { x: c.tx, y: 4.88, w: 1.871, h: 0.337, fontSize: 14, bold: true, color: c.accent });
    txt(s, c.quote, { x: c.tx, y: 5.141, w: 3.017, h: 0.675, fontSize: 12, italic: true, color: c.quoteColor, lineSpacingMultiple: 1.5 });
    round(s, c.px, 5.009, 0.406, 0.094, c.accent, 0.047);
    oval(s, c.dx, 5.009, 0.094, 0.094, c.accent);
  });
  txt(s, 'Personal Hiking Experiences', { x: 4.778, y: 2.409, w: 3.73, h: 1.178, fontSize: 32, bold: true, color: GREEN, align: 'center' });
  txt(s, 'Stories from the Trail', { x: 4.778, y: 3.587, w: 3.73, h: 0.337, fontSize: 14, bold: true, color: GREEN, align: 'center' });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et',
    { x: 4.953, y: 3.924, w: 3.38, h: 0.978, fontSize: 12, color: INK, align: 'center', lineSpacingMultiple: 1.5 });
}

// 9 - Infographic Analysis (four quarter-round tiles)
function slide09(pptx) {
  const s = pptx.addSlide();
  chrome(s, GREEN);
  heading(s, 'Infographic Analysis', 3.441, 6.451);

  const tiles = [
    { x: 4.603, y: 2.361, size: 1.933, flipH: true, flipV: false, color: GREEN },
    { x: 6.644, y: 2.208, size: 2.086, flipH: false, flipV: false, color: YELLOW },
    { x: 4.603, y: 4.4, size: 1.933, flipH: true, flipV: true, color: YELLOW },
    { x: 6.646, y: 4.4, size: 1.933, flipH: false, flipV: true, color: GREEN },
  ];
  tiles.forEach((t) => {
    const inner = t.size - 0.154;
    shape(s, 'round1Rect', {
      x: t.x, y: t.y, w: t.size, h: t.size, flipH: t.flipH, flipV: t.flipV,
      fill: { color: WHITE }, rectRadius: t.size / 2, shadow: SHADOW(),
    });
    shape(s, 'round1Rect', {
      x: t.x + 0.077, y: t.y + 0.077, w: inner, h: inner, flipH: t.flipH, flipV: t.flipV,
      fill: { color: t.color }, rectRadius: inner / 2,
    });
    mark(s, t.x + t.size / 2, t.y + t.size / 2, 0.46, WHITE);
  });

  [['Analysis 01', 1.361, 2.549, GREEN], ['Analysis 03', 1.361, 4.649, YELLOW],
   ['Analysis 02', 9.653, 2.549, YELLOW], ['Analysis 04', 9.649, 4.649, GREEN]].forEach(([label, x, y, color]) => {
    txt(s, label, { x, y, w: 1.636, h: 0.337, fontSize: 14, bold: true, color });
    txt(s, 'Lorem ipsum dolor sit amet, elit. Sed do eiusmod tempor', Object.assign({ x, y: y + 0.366, w: 2.402, h: 0.675 }, BODY));
  });
}

// 10 - Infographic Section (five map pins)
function slide10(pptx) {
  const s = pptx.addSlide();
  chrome(s, GREEN);
  heading(s, 'Infographic Section', 3.165, 7.003);

  for (let i = 0; i < 5; i++) {
    const x = 1.289 + i * 2.242;
    const color = i % 2 === 0 ? YELLOW : GREEN;
    const cx = x + 0.894;
    oval(s, x, 2.096, 1.787, 1.787, color, { shadow: SOFT() });
    rect(s, x + 0.791, 3.873, 0.205, 0.253, { color });
    shape(s, 'triangle', { x: x + 0.689, y: 4.126, w: 0.409, h: 0.205, fill: { color }, flipV: true });
    oval(s, cx - 0.425, 2.562, 0.85, 0.85, WHITE);
    mark(s, cx, 2.987, 0.38, color);
    txt(s, 'Idea 0' + (i + 1), { x: x + 0.397, y: 4.612, w: 0.993, h: 0.37, fontSize: 16, bold: true, color: INK85, align: 'center' });
    txt(s, 'Lorem ipsum dolor sit amet, elit. Sed vestibule eros eget adipiscing, ', {
      x: x - 0.137, y: 4.922, w: 2.062, h: 1.01, fontSize: 12, color: INK85, align: 'center', lineSpacingMultiple: 1.5,
    });
  }
}

// 11 - Mind Map Infographic Section
function slide11(pptx) {
  const s = pptx.addSlide();
  chrome(s, GREEN);
  heading(s, 'Mind Map Infographic Section', 3.165, 7.003);

  const bands = [
    { x: 2.587, y: 3.214, rot: 126.745, color: GREEN },
    { x: 4.746, y: 3.401, rot: 46.889, color: YELLOW },
    { x: 6.902, y: 3.214, rot: 126.745, color: GREEN },
    { x: 9.057, y: 3.401, rot: 46.889, color: YELLOW },
  ];
  bands.forEach((b) => shape(s, 'trapezoid', { x: b.x, y: b.y, w: 0.97, h: 1.63, rotate: b.rot, fill: { color: b.color }, shadow: SHADOW() }));

  const nodes = [
    { x: 1.355, y: 2.494, color: GREEN, label: 'Planning 01', lc: GREEN, tx: 1.55, ty: 4.807, bx: 1.387, by: 5.155 },
    { x: 3.578, y: 3.833, color: YELLOW, label: 'Planning 02', lc: YELLOW, tx: 3.773, ty: 2.497, bx: 3.61, by: 2.846 },
    { x: 5.669, y: 2.494, color: GREEN, label: 'Planning 03', lc: GREEN, tx: 5.865, ty: 4.807, bx: 5.702, by: 5.155 },
    { x: 7.889, y: 3.833, color: YELLOW, label: 'Planning 04', lc: YELLOW, tx: 8.084, ty: 2.501, bx: 7.921, by: 2.849 },
    { x: 9.984, y: 2.494, color: GREEN, label: 'Planning 06', lc: GREEN, tx: 10.18, ty: 4.807, bx: 10.017, by: 5.155 },
  ];
  nodes.forEach((n) => {
    oval(s, n.x, n.y, 1.994, 1.994, n.color, { shadow: SHADOW() });
    oval(s, n.x + 0.204, n.y + 0.204, 1.587, 1.587, WHITE);
    mark(s, n.x + 0.997, n.y + 0.997, 0.62, n.color);
    txt(s, n.label, { x: n.tx, y: n.ty, w: 1.604, h: 0.337, fontSize: 14, bold: true, color: n.lc, align: 'center' });
    txt(s, 'Lorem ipsum dolor sit amet, ', {
      x: n.bx, y: n.by, w: 1.929, h: 0.673, fontSize: 12, color: INK, align: 'center', lineSpacingMultiple: 1.5,
    });
  });
}

// 12 - Infographic Section (hill)
function slide12(pptx) {
  const s = pptx.addSlide();
  chrome(s, GREEN);
  heading(s, 'Infographic Section', 3.165, 7.003);

  // Dome outline: symmetric bezier hill, drawn twice (white halo + green body).
  const hill = (x, y, w, h, color, shadow) =>
    shape(s, 'custGeom', {
      x, y, w, h, fill: { color }, shadow,
      points: [
        { x: w * 0.5, y: 0 },
        { x: w * 0.99659, y: h * 0.97174, curve: { type: 'cubic', x1: w * 0.728, y1: 0, x2: w * 0.9225, y2: h * 0.4047 } },
        { x: w, y: h },
        { x: 0, y: h },
        { x: w * 0.00341, y: h * 0.97174 },
        { x: w * 0.5, y: 0, curve: { type: 'cubic', x1: w * 0.07753, y1: h * 0.4047, x2: w * 0.27189, y2: 0 } },
        { close: true },
      ],
    });
  hill(2.694, 4.663, 7.947, 2.837, WHITE, SOFT());
  hill(2.786, 4.729, 7.763, 2.771, GREEN, SHADOW());

  const nodes = [
    { ox: 2.635, oy: 5.532, color: GREEN, label: 'Strategy 01', lx: 2.52, ly: 4.136, bx: 2.463, by: 4.445 },
    { ox: 4.814, oy: 4.181, color: YELLOW, label: 'Strategy 02', lx: 4.687, ly: 2.784, bx: 4.629, by: 3.093 },
    { ox: 7.213, oy: 4.181, color: GREEN, label: 'Strategy 03', lx: 7.086, ly: 2.784, bx: 7.029, by: 3.093 },
    { ox: 9.394, oy: 5.532, color: YELLOW, label: 'Strategy 04', lx: 9.267, ly: 4.136, bx: 9.21, by: 4.445 },
  ];
  nodes.forEach((n) => {
    oval(s, n.ox, n.oy, 1.308, 1.308, WHITE, { shadow: SOFT() });
    oval(s, n.ox + 0.071, n.oy + 0.072, 1.165, 1.165, n.color);
    mark(s, n.ox + 0.654, n.oy + 0.654, 0.45, WHITE);
    txt(s, n.label, { x: n.lx, y: n.ly, w: 1.562, h: 0.337, fontSize: 14, bold: true, color: INK85, align: 'center' });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur', {
      x: n.bx, y: n.by, w: 1.676, h: 0.976, fontSize: 12, color: INK, align: 'center', lineSpacingMultiple: 1.5,
    });
  });

  mark(s, 6.743, 5.952, 0.855, WHITE);
  txt(s, 'Strategy Marketing', { x: 5.296, y: 6.652, w: 2.895, h: 0.37, fontSize: 16, bold: true, color: WHITE, align: 'center' });
}

// 13 - Timeline Data Analysis Infographic
function slide13(pptx) {
  const s = pptx.addSlide();
  chrome(s, GREEN);
  txt(s, 'Timeline Data Analysis Infographic', { x: 6.823, y: 1.425, w: 5.481, h: 1.313, fontSize: 36, bold: true, color: GREEN });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do tempor incididunt ut labore et dolore magna aliqua. Ut enim ad',
    { x: 6.823, y: 2.829, w: 5.78, h: 0.707, fontSize: 12, color: INK85, lineSpacingMultiple: 1.5 });

  const stats = [
    { rx: 1.234, ry: 1.568, tx: 3.459, ty: 1.785, color: GREEN, label: 'First Our Data', value: '673,000++' },
    { rx: 1.234, ry: 4.221, tx: 3.459, ty: 4.438, color: YELLOW, label: 'Second Our Data', value: '830,000,00' },
    { rx: 6.928, ry: 4.221, tx: 9.152, ty: 4.438, color: GREEN, label: 'Third Our Data', value: '484,399,02k' },
  ];
  stats.forEach((b) => {
    oval(s, b.rx, b.ry, 1.942, 1.942, b.color);
    oval(s, b.rx + 0.19, b.ry + 0.19, 1.562, 1.562, WHITE);
    mark(s, b.rx + 0.971, b.ry + 0.971, 0.85, b.color);
    txt(s, b.label, { x: b.tx, y: b.ty, w: 2.939, h: 0.337, fontSize: 14, bold: true, color: INK85 });
    txt(s, b.value, { x: b.tx, y: b.ty + 0.4, w: 2.939, h: 0.572, fontSize: 28, bold: true, color: b.color });
    txt(s, 'Lorem ipsum dolor sit amet, elit, sed do eiusmod tempor',
      { x: b.tx, y: b.ty + 0.981, w: 2.939, h: 0.707, fontSize: 12, color: INK85, lineSpacingMultiple: 1.5 });
  });
}

// 14 - Infographic Section (four support cards)
function slide14(pptx) {
  const s = pptx.addSlide();
  chrome(s, GREEN);
  heading(s, 'Infographic Section', 3.165, 7.003);

  const cards = [
    { x: 0.963, color: GREEN, down: false, label: 'Support 01', body: 'Lorem ipsum dolor sit amet, et ud consectetur adipiscing. ', bw: 2.342, bx: 1.072 },
    { x: 3.912, color: YELLOW, down: true, label: 'Support 02', body: 'Lorem ipsum dolor sit amet, et ud consectetur adipiscing elit. ', bw: 2.229, bx: 4.078 },
    { x: 6.861, color: GREEN, down: false, label: 'Support 03', body: 'Lorem ipsum dolor sit amet, et ud consectetur adipiscing elit. ', bw: 2.335, bx: 6.974 },
    { x: 9.81, color: YELLOW, down: true, label: 'Support 04', body: 'Lorem ipsum dolor sit amet, et ud consectetur adipiscing elit. ', bw: 2.229, bx: 9.976 },
  ];
  cards.forEach((c) => {
    const backY = c.down ? 3.005 : 2.331;
    const frontY = c.down ? 2.843 : 2.677;
    const circleY = c.down ? 2.331 : 4.637;
    round(s, c.x, backY, 2.561, 2.652, c.color, 0.253, { flipV: c.down, shadow: SOFT() });
    round(s, c.x + 0.132, frontY, 2.297, 2.468, WHITE, 0.189, { flipV: c.down, shadow: SHADOW() });
    oval(s, c.x + 0.77, circleY, 1.021, 1.021, WHITE);
    oval(s, c.x + 0.829, circleY + 0.059, 0.902, 0.902, c.color);
    mark(s, c.x + 1.28, circleY + 0.51, 0.42, WHITE);
    txt(s, c.label, { x: c.x + 0.432, y: c.down ? 3.496 : 2.928, w: 1.695, h: 0.337, fontSize: 14, bold: true, color: INK85, align: 'center' });
    txt(s, c.body, {
      x: c.bx, y: c.down ? 3.835 : 3.289, w: c.bw, h: 0.978, fontSize: 12, color: INK85, align: 'center', lineSpacingMultiple: 1.5,
    });
  });
}

// 15 - Geometric Infographic (four translucent circles)
function slide15(pptx) {
  const s = pptx.addSlide();
  chrome(s, GREEN);
  heading(s, 'Geometric Infographic', 3.165, 7.003);

  [[5.623, 2.303, YELLOW], [6.546, 3.225, YELLOW], [5.623, 4.148, GREEN], [4.701, 3.225, GREEN]].forEach(([x, y, color]) =>
    shape(s, 'ellipse', { x, y, w: 2.086, h: 2.086, fill: { color, transparency: 25 }, line: { color: WHITE, width: 1 }, shadow: SOFT() })
  );

  const items = [
    { label: 'Advantage  02', y: 2.828, color: GREEN, side: 'left' },
    { label: 'Advantage  01', y: 4.721, color: GREEN, side: 'left' },
    { label: 'Advantage  03', y: 2.828, color: YELLOW, side: 'right' },
    { label: 'Advantage  04', y: 4.721, color: YELLOW, side: 'right' },
  ];
  items.forEach((it) => {
    const left = it.side === 'left';
    round(s, left ? 4.023 : 9.26, it.y + 0.094, 0.05, 0.819, it.color, 0.008, { shadow: SOFT() });
    txt(s, it.label, { x: left ? 1.969 : 9.365, y: it.y, w: 2.0, h: 0.337, fontSize: 14, bold: true, color: it.color, align: left ? 'right' : 'left' });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit,', {
      x: left ? 1.236 : 9.365, y: it.y + 0.314, w: left ? 2.732 : 3.066, h: 0.675,
      fontSize: 12, color: INK, align: left ? 'right' : 'left', lineSpacingMultiple: 1.5,
    });
  });
}

// 16 - Sales Infographic
function slide16(pptx) {
  const s = pptx.addSlide();
  chrome(s, GREEN);
  heading(s, 'Sales Infographic', 3.165, 7.003);

  const rows = [2.202, 3.308, 4.414, 5.521];
  rows.forEach((y, i) => {
    [{ side: 'l' }, { side: 'r' }].forEach((col) => {
      const left = col.side === 'l';
      const num = left ? 2 * i + 1 : 2 * i + 2;
      const chipColor = (num % 4 === 1 || num % 4 === 0) ? YELLOW : GREEN;
      round(s, left ? 1.077 : 9.077, y, 3.179, 0.948, WHITE, 0.106, { flipH: !left, shadow: WIDE() });
      round(s, left ? 3.765 : 8.611, y + 0.083, 0.958, 0.782, chipColor, 0.088, { flipH: !left, shadow: SHADOW() });
      txt(s, i === 0 && left ? 'Lorem ipsum dolor sit, adipiscing consectetur' : 'Lorem ipsum dolor sit, consectetur adipiscing', {
        x: left ? 1.336 : 9.827, y: y + 0.089, w: 2.171, h: 0.678, fontSize: 12, color: INK, lineSpacingMultiple: 1.5,
      });
      txt(s, ('0' + num).slice(-2), { x: left ? 3.944 : 8.79, y: y + 0.255, w: 0.599, h: 0.438, fontSize: 20, bold: true, color: WHITE, align: 'center' });
    });
  });

  // Connectors run from behind the hub out to each chip; the hub is drawn last
  // and covers their inner ends. Elbows are two segments, arrow on the last one.
  const wire = (x1, y1, x2, y2, color, arrow) =>
    shape(s, 'line', {
      x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
      flipH: x2 < x1, flipV: y2 < y1, line: { color, width: 1, endArrowType: arrow ? 'triangle' : 'none' },
    });
  [[5.7495, 3.4185, 4.7225, 2.6765, YELLOW], [7.5845, 3.4185, 8.6115, 2.6765, GREEN],
   [5.7495, 5.2535, 4.7225, 5.9955, GREEN], [7.5845, 5.2535, 8.6115, 5.9955, YELLOW]].forEach(([xh, yh, xc, yc, color]) => {
    wire(xh, yh, xc, yh, color, false);
    wire(xc, yh, xc, yc, color, true);
  });
  wire(5.522, 3.782, 4.723, 3.793, GREEN, true);
  wire(7.827, 3.782, 8.611, 3.793, YELLOW, true);
  wire(5.522, 4.889, 4.723, 4.889, YELLOW, true);
  wire(7.827, 4.889, 8.611, 4.889, GREEN, true);

  oval(s, 5.369, 3.038, 2.595, 2.595, WHITE, { shadow: WIDE() });
  oval(s, 5.509, 3.177, 2.316, 2.316, WHITE, { shadow: SHADOW() });
  mark(s, 6.666, 4.335, 1.193, GREEN);
}

// 17 - Development Marketing Infographic
function slide17(pptx) {
  const s = pptx.addSlide();
  chrome(s, GREEN);
  heading(s, 'Development Marketing Infographic', 2.419, 8.496);

  const arms = [
    { bar: 2.861, barY: 1.332, rot: 270, flipH: true, color: YELLOW, ox: 4.359, oy: 2.168, tx: 1.984, ty: 2.416, align: 'right', tc: INK, num: '02', nx: 4.716, ny: 2.652 },
    { bar: 9.151, barY: 1.332, rot: 90, flipH: false, color: GREEN, ox: 7.502, oy: 2.168, tx: 9.009, ty: 2.416, align: 'left', tc: WHITE, num: '03', nx: 7.859, ny: 2.652 },
    { bar: 2.861, barY: 3.664, rot: 270, flipH: true, color: GREEN, ox: 4.359, oy: 4.501, tx: 1.984, ty: 4.749, align: 'right', tc: WHITE, num: '01', nx: 4.716, ny: 4.985 },
    { bar: 9.151, barY: 3.664, rot: 90, flipH: false, color: YELLOW, ox: 7.502, oy: 4.501, tx: 9.009, ty: 4.749, align: 'left', tc: INK, num: '04', nx: 7.859, ny: 4.985 },
  ];
  arms.forEach((a) =>
    shape(s, 'round2SameRect', { x: a.bar, y: a.barY, w: 1.322, h: 3.146, rotate: a.rot, flipH: a.flipH, fill: { color: a.color }, shadow: SHADOW() })
  );
  oval(s, 4.716, 2.274, 3.902, 3.902, WHITE, { shadow: SHADOW() });
  oval(s, 4.968, 2.526, 3.398, 3.398, GREEN, { shadow: SHADOW() });
  arms.forEach((a) => {
    oval(s, a.ox, a.oy, 1.472, 1.472, WHITE, { flipH: a.flipH, shadow: SHADOW() });
    oval(s, a.ox + 0.095, a.oy + 0.095, 1.282, 1.282, a.color, { flipH: a.flipH, shadow: SHADOW() });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do', {
      x: a.tx, y: a.ty, w: 2.341, h: 0.981, fontSize: 12, color: a.tc, align: a.align, lineSpacingMultiple: 1.5,
    });
    txt(s, a.num, { x: a.nx, y: a.ny, w: 0.758, h: 0.505, fontSize: 24, bold: true, color: WHITE, align: 'center' });
  });
  mark(s, 6.666, 4.225, 1.001, WHITE);
}

// 18 - Analysis Process Infographic
function slide18(pptx) {
  const s = pptx.addSlide();
  chrome(s, GREEN);
  heading(s, 'Analysis Process Infographic', 2.419, 8.496);

  for (let i = 0; i < 5; i++) {
    const up = i % 2 === 0;
    const color = up ? GREEN : YELLOW;
    const ax = 1.938 + i * 1.8385;
    const ay = up ? 3.177 : 3.292;
    shape(s, 'blockArc', {
      x: ax, y: ay, w: 2.104, h: 2.104, flipV: !up, fill: { color }, shadow: SHADOW(),
      angleRange: [180, 1.131], arcThicknessRatio: 0.2474,
    });
    oval(s, ax + 0.329, ay + 0.33, 1.444, 1.444, WHITE, { flipV: !up, shadow: SHADOW() });
    oval(s, ax + 0.393, ay + 0.393, 1.318, 1.318, color, { flipV: !up, shadow: SHADOW() });
    mark(s, ax + 1.052, ay + 1.052, 0.5, WHITE);
    const ty = up ? 1.975 : 5.612;
    txt(s, 'Analysis 0' + (i + 1), { x: ax + 0.381, y: ty, w: 1.34, h: 0.337, fontSize: 14, bold: true, color: INK, align: 'center' });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur', {
      x: ax + 0.122, y: ty + 0.312, w: 1.859, h: 0.675, fontSize: 12, color: INK, align: 'center', lineSpacingMultiple: 1.5,
    });
  }
}

// 19 - Stay Connected with Us
function slide19(pptx) {
  const s = pptx.addSlide();
  chrome(s, GREEN);
  photoBox(s, 7.107, 1.702, 5.667, 4.244, 0.2);
  txt(s, 'Stay Connected with Us Anytime and Anywhere', { x: 0.722, y: 1.559, w: 5.864, h: 1.178, fontSize: 32, bold: true, color: GREEN });
  txt(s, 'We\u2019re Always Here to Help', { x: 0.722, y: 2.776, w: 3.026, h: 0.337, fontSize: 14, bold: true, color: GREEN });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore',
    Object.assign({ x: 0.722, y: 3.213, w: 5.464, h: 0.675 }, BODY));

  const rows = [
    { y: 4.133, radius: 0.175, label: 'Email :', ly: 4.209, lw: 0.884, lsize: 14, vx: 2.427, vw: 2.03, value: 'yourwebsite@mail.com' },
    { y: 4.792, radius: 0.15, label: 'Phone :', ly: 4.872, lw: 1.082, lsize: 14, vx: 2.422, vw: 1.917, value: '0953 \u2013 576970800000' },
    { y: 5.457, radius: 0.161, label: 'Website :', ly: 5.55, lw: 0.974, lsize: 12, vx: 2.457, vw: 2.094, value: 'www.yourwebsite.com' },
  ];
  rows.forEach((r) => {
    round(s, 0.802, r.y, 3.977, 0.489, GREEN, r.radius);
    mark(s, 1.283, r.y + 0.245, 0.24, WHITE);
    txt(s, r.label, { x: 1.538, y: r.ly, w: r.lw, h: 0.337, fontSize: r.lsize, bold: true, color: WHITE });
    txt(s, r.value, { x: r.vx, y: r.y + 0.093, w: r.vw, h: 0.303, fontSize: 12, color: WHITE });
  });
}

// 20 - Thank You
function slide20(pptx) {
  const s = pptx.addSlide();
  rect(s, 0, 0.844, 13.333, 6.656, { color: OVERLAY, transparency: 50 });
  chrome(s, WHITE);
  txt(s, 'Mountain & Hiking Presentation ', {
    x: 3.615, y: 1.772, w: 6.103, h: 0.337, fontSize: 14, bold: true, color: WHITE, align: 'center', charSpacing: 3,
  });
  txt(s, 'Thank You For Your Attention', { x: 2.002, y: 2.109, w: 9.329, h: 2.524, fontSize: 72, bold: true, color: WHITE, align: 'center' });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation',
    { x: 2.954, y: 4.519, w: 7.425, h: 0.675, fontSize: 12, color: WHITE, align: 'center', lineSpacingMultiple: 1.5 });
  round(s, 5.429, 5.512, 2.555, 0.494, GREEN, 0.147);
  txt(s, 'See You Next Time', { x: 5.627, y: 5.608, w: 2.159, h: 0.303, fontSize: 12, bold: true, color: WHITE, align: 'center' });
}

/* ------------------------------------------------------------------- build */
function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'HORIZON', width: 13.333, height: 7.5 });
  pptx.layout = 'HORIZON';
  pptx.title = 'HorizonTrail - Mountain & Hiking Presentation';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
   slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach((fn) => fn(pptx));

  return pptx.writeFile({ fileName: path.join(__dirname, '090d2459-35a7-493f-80d9-d0a6836b9b68_grok_final.pptx') });
}

build().then((f) => console.log('wrote', f));
