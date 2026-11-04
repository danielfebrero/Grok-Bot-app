/**
 * "Sublm 24x7 / Fashion Lookbook" deck - rebuilt with pptxgenjs.
 *
 * Slide canvas is 26.667 x 15 in (24384000 x 13716000 EMU).
 * Everything below is plain data + small helpers: no external assets are read,
 * raster artwork from the original is represented by `imageSlot()` rectangles.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

const SLIDE_W = 26 + 2 / 3; // 24384000 EMU
const SLIDE_H = 15; // 13716000 EMU

/* ------------------------------------------------------------------ palette */

const C = {
  black: '000000', // pure black text on the white cover pages
  ink: '0A0A0A', // near-black used for headings and full-bleed bands
  paper: 'FAFAFA', // off-white page background / knock-out text
  white: 'FFFFFF',
};

/* The original leaves every photo frame empty, so the stand-in rectangles are
 * drawn as a faint tint of whatever surface they sit on. */
const SLOT_TINT = { FFFFFF: 'F5F5F5', FAFAFA: 'F1F1F1', '0A0A0A': '151515' };

const FONT = {
  display: 'Anton Regular', // theme minor font - all the oversized headlines
  body: 'Poppins Regular',
  bodyBold: 'Poppins Bold',
};

/* Text roles taken from the original master / list styles. */
const STYLE = {
  hero: { fontFace: FONT.display, fontSize: 287, charSpacing: -8.6, lineSpacingMultiple: 0.7, caps: true },
  title: { fontFace: FONT.display, fontSize: 177, charSpacing: -5.31, lineSpacingMultiple: 0.8, caps: true },
  head: { fontFace: FONT.display, fontSize: 110, charSpacing: -3.3, lineSpacingMultiple: 0.7, caps: true },
  sub: { fontFace: FONT.display, fontSize: 68, charSpacing: -2.04, lineSpacingMultiple: 0.7, caps: true },
  lead: { fontFace: FONT.body, fontSize: 26 },
  body: { fontFace: FONT.body, fontSize: 16 },
  label: { fontFace: FONT.body, fontSize: 16, caps: true },
  labelBold: { fontFace: FONT.bodyBold, fontSize: 16, caps: true },
};

/* Copy blocks that the deck repeats verbatim. */
const COPY = {
  intro:
    'In the ever-evolving realm of fashion, the landscape for men has undergone a significant ' +
    'transformation in recent years.',
  introLong:
    'In the ever-evolving realm of fashion, the landscape for men has undergone a significant ' +
    'transformation in recent years. The modern man is no longer confined to the limited choices ' +
    'of yesteryears but is now presented with a diverse',
  tradition:
    "Traditionally, men's fashion has often been associated with a more conservative and restrained " +
    'approach. However, the modern man has shattered these preconceived notions, embracing a style ' +
    'that is as varied as his interests.',
  traditionShort:
    "Traditionally, men's fashion has often been associated with a more conservative and restrained " +
    'approach.',
  stamp: "Fashion Lookbook\nyamamoto '24",
  brand: 'Sublm 24x7 fashion trends',
  page: 'Page 2',
  krux: 'The Jimmy Krux Jeunking in The Hills',
  evolution: 'The Evolution of Modern Menswear',
  memories: 'memories on keep your eyes open',
  divide: 'Devide et impera',
  italy: 'Italy / rome / delpi',
};

/* ------------------------------------------------------------------ helpers */

/** Text box. `role` selects a STYLE preset, `box` carries geometry + overrides. */
function txt(slide, role, text, box) {
  const s = STYLE[role];
  const opts = Object.assign(
    {
      x: 0,
      y: 0,
      w: 1,
      h: 1,
      color: C.black,
      align: 'left',
      valign: 'middle',
      fontFace: s.fontFace,
      fontSize: s.fontSize,
      charSpacing: s.charSpacing,
      lineSpacingMultiple: s.lineSpacingMultiple,
      inset: 0.0556,
      fit: 'resize',
      wrap: true,
    },
    box
  );
  slide.addText(s.caps ? text.toUpperCase() : text, opts);
}

/** The "PAGE 2" / "SUBLM 24X7 FASHION TRENDS" pair that runs across the deck. */
function runningHead(slide, y, color) {
  txt(slide, 'label', COPY.page, { x: 1.286, y: y, w: 4.725, h: 0.436, color: color });
  txt(slide, 'label', COPY.brand, { x: 20.656, y: y, w: 4.725, h: 0.436, color: color, align: 'right' });
}

/** Two-line "FASHION LOOKBOOK / YAMAMOTO '24" credit stamp. */
function stamp(slide, x, y, box) {
  txt(slide, 'label', COPY.stamp, Object.assign({ x: x, y: y, w: 4.725, h: 0.783 }, box));
}

/** Flat colour band / panel. */
function band(slide, x, y, w, h, color) {
  slide.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: color }, line: { type: 'none' } });
}

/** Placeholder standing in for a photo frame of the original deck. */
function imageSlot(slide, x, y, w, h, surface) {
  slide.addShape('rect', {
    x: x, y: y, w: w, h: h,
    fill: { color: SLOT_TINT[surface || slide._surface] },
    line: { type: 'none' },
  });
}

/** Outlined circle - the deck's line-icon motif (slides 1, 21 and 22). */
function iconRing(slide, x, y, d, glyph) {
  const t = 8; // stroke weight in points
  const pad = t / 144; // half the stroke, in inches, so the outer edge lands on `d`
  slide.addShape('ellipse', {
    x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad,
    fill: { type: 'none' },
    line: { color: C.black, width: t },
  });
  if (glyph) {
    slide.addText(glyph, {
      x: x, y: y, w: d, h: d,
      fontFace: FONT.body, fontSize: Math.round(d * 42), bold: true, color: C.black,
      align: 'center', valign: 'middle', inset: 0, fit: 'none',
    });
  }
}

/** Regular grid of icon rings; `glyphs` supplies one row of symbols at a time. */
function iconGrid(slide, originX, rows, stepX, d, glyphs) {
  rows.forEach(function (y, r) {
    glyphs[r].forEach(function (g, i) { iconRing(slide, originX + i * stepX, y, d, g); });
  });
}

function newSlide(pptx, bg) {
  const slide = pptx.addSlide();
  slide.background = { color: bg };
  slide._surface = bg;
  return slide;
}

/* ------------------------------------------------------------- slide builders */

// 1 - cover: giant wordmark, arrow icon, wide photo band
function slide01(pptx) {
  const s = newSlide(pptx, C.white);
  imageSlot(s, 1.286, 6.502, 24.095, 7.39);
  txt(s, 'title', 'Sublm 24x7', { x: 1.286, y: 1.979, w: 11.941, h: 3.806 });
  txt(s, 'title', '*', { x: 23.567, y: 1.979, w: 1.814, h: 3.806, align: 'right' });
  iconRing(s, 22.697, 2.845, 1.316, '\u2199');
  txt(s, 'label', 'Fashion Lookbook Presentation Template', {
    x: 13.439, y: 4.193, w: 4.051, h: 0.783, charSpacing: 3.2,
  });
  txt(s, 'label', COPY.brand, { x: 20.656, y: 1.108, w: 4.725, h: 0.436, align: 'right' });
  txt(s, 'label', 'new trends \nartcl 196', { x: 1.286, y: 1.108, w: 4.725, h: 0.783 });
}

// 2 - "Discovery Very Rarely" + two-column lead paragraph
function slide02(pptx) {
  const s = newSlide(pptx, C.white);
  imageSlot(s, 1.286, 6.349, 13.06, 7.542);
  runningHead(s, 1.108, C.black);
  txt(s, 'label', 'The snake hole', { x: 1.286, y: 2.434, w: 4.725, h: 0.436 });
  txt(s, 'head', 'Discovery Very Rarely', { x: 1.286, y: 2.87, w: 13.594, h: 2.403, color: C.ink });
  // One 9.916" wide two-column text frame in the original; rebuilt as two boxes
  // (column width 4.655", gutter 0.496").
  const col = { w: 4.766, h: 6.47, color: C.ink, valign: 'top', fit: 'none' };
  txt(s, 'lead',
    'In the ever-evolving realm of fashion, the landscape for men has undergone a significant ' +
    'transformation in recent years. The modern man is no longer confined to the limited choices ' +
    'PLACEHOLDER' +
    'personality, lifestyle, and a keen sense of individuality.',
    Object.assign({ x: 15.465, y: 6.349 }, col));
  txt(s, 'lead',
    "Let's take a journey through the nuances of modern men's fashion, exploring the trends, " +
    'influences, and the essence of self-expression.',
    Object.assign({ x: 20.615, y: 6.349 }, col));
  stamp(s, 15.465, 13.108);
}

// 3 - two photo frames over an oversized "MODERN MENSWEAR"
function slide03(pptx) {
  const s = newSlide(pptx, C.paper);
  imageSlot(s, 1.286, 3.729, 13.967, 7.542);
  imageSlot(s, 15.465, 3.729, 9.916, 7.542);
  runningHead(s, 1.108, C.black);
  txt(s, 'hero', 'Modern Menswear', { x: -1.368, y: 10.448, w: 29.614, h: 6.111 });
  stamp(s, 1.286, 2.513);
  txt(s, 'label', COPY.memories, { x: 6.473, y: 2.513, w: 2.703, h: 0.783, align: 'right' });
  txt(s, 'sub', '2024', { x: 11.787, y: 2.141, w: 13.594, h: 1.528, color: C.ink, align: 'right' });
}

// 4 - date splash with a left copy rail
function slide04(pptx) {
  const s = newSlide(pptx, C.paper);
  imageSlot(s, 9.388, 8.433, 17.279, 6.567);
  txt(s, 'hero', '*05/20/24', { x: 8.548, y: 3.677, w: 16.833, h: 6.111, align: 'right' });
  runningHead(s, 1.108, C.black);
  stamp(s, 1.286, 2.27, { fontFace: FONT.bodyBold });
  txt(s, 'body', COPY.intro, { x: 1.286, y: 3.192, w: 5.865, h: 1.131 });
  txt(s, 'labelBold', COPY.divide, { x: 1.286, y: 4.909, w: 4.725, h: 0.436 });
  txt(s, 'body', COPY.tradition, { x: 1.286, y: 5.484, w: 5.865, h: 1.825 });
  txt(s, 'sub', COPY.italy, { x: 1.286, y: 12.364, w: 7.89, h: 1.528, color: C.ink });
}

// 5 - dark page, three stacked photo frames
function slide05(pptx) {
  const s = newSlide(pptx, C.paper);
  band(s, 0, 1.94, 26.667, 13.06, C.ink);
  runningHead(s, 0.752, C.black);
  txt(s, 'body', COPY.tradition, { x: 1.286, y: 10.491, w: 5.865, h: 1.825, color: C.paper });
  txt(s, 'sub', COPY.italy, { x: 17.702, y: 3.356, w: 7.679, h: 1.528, color: C.paper, align: 'right' });
  imageSlot(s, 9.494, 3.707, 7.679, 9.526, C.ink);
  imageSlot(s, 1.286, 3.707, 7.679, 6.368, C.ink);
  imageSlot(s, 17.702, 5.151, 7.679, 8.082, C.ink);
}

// 6 - full-bleed image with centred knock-out headline
function slide06(pptx) {
  const s = newSlide(pptx, C.paper);
  imageSlot(s, 0, 0, 26.667, 15);
  txt(s, 'hero', 'This is a sublm ', {
    x: 4.917, y: 4.312, w: 16.833, h: 7.93, color: C.paper, align: 'center', lineSpacingMultiple: 0.8,
  });
  txt(s, 'body', COPY.intro, { x: 9.569, y: 2.149, w: 7.529, h: 0.783, color: C.paper, align: 'center' });
  txt(s, 'body', '05 / 20 / 2024', { x: 9.569, y: 12.242, w: 7.529, h: 0.436, color: C.paper, align: 'center' });
}

// 7 - dark page, three frames plus two light caption cards
function slide07(pptx) {
  const s = newSlide(pptx, C.paper);
  band(s, 0, 1.94, 26.667, 13.06, C.ink);
  runningHead(s, 0.752, C.black);
  band(s, 8.27, 11.005, 8.279, 2.532, C.paper);
  band(s, 17.102, 3.404, 8.279, 2.532, C.paper);
  txt(s, 'head', COPY.evolution, {
    x: 1.286, y: 3.976, w: 6.639, h: 6.071, color: C.paper, lineSpacingMultiple: 0.8,
  });
  stamp(s, 8.691, 11.33, { h: 0.718, fontFace: FONT.bodyBold, lineSpacingMultiple: 0.8 });
  txt(s, 'body', COPY.intro, { x: 8.691, y: 12.081, w: 5.865, h: 1.131 });
  txt(s, 'sub', 'Modernist', { x: 17.518, y: 3.906, w: 4.725, h: 1.528, color: C.ink });
  txt(s, 'labelBold', '001', { x: 15.224, y: 12.775, w: 0.726, h: 0.436, align: 'right', lineSpacingMultiple: 0.8 });
  imageSlot(s, 8.27, 3.404, 8.279, 3.636, C.ink);
  imageSlot(s, 8.27, 7.369, 8.279, 3.636, C.ink);
  imageSlot(s, 17.102, 5.935, 8.279, 7.601, C.ink);
}

// 8 - section marker over a wide photo band
function slide08(pptx) {
  const s = newSlide(pptx, C.paper);
  imageSlot(s, 0, 4.028, 26.667, 10.972);
  txt(s, 'sub', COPY.evolution, { x: 1.286, y: 2.022, w: 13.594, h: 1.528, color: C.ink });
  runningHead(s, 1.108, C.black);
}

// 9 - headline reversed out of a top-bleed photo
function slide09(pptx) {
  const s = newSlide(pptx, C.paper);
  imageSlot(s, 1.286, 0, 24.095, 11.882);
  txt(s, 'hero', 'Modern Menswear', {
    x: 2.471, y: 4.136, w: 16.833, h: 7.93, color: C.paper, lineSpacingMultiple: 0.8,
  });
  txt(s, 'label', COPY.brand, { x: 19.347, y: 1.108, w: 4.725, h: 0.436, align: 'right' });
  txt(s, 'label', COPY.page, { x: 2.595, y: 1.108, w: 4.725, h: 0.436 });
  stamp(s, 1.286, 12.736);
  txt(s, 'label', COPY.memories, { x: 6.473, y: 12.736, w: 2.703, h: 0.783, align: 'right' });
  txt(s, 'sub', '2024', { x: 11.787, y: 12.364, w: 13.594, h: 1.528, color: C.ink, align: 'right' });
}

// 10 - copy rail left, one tall + three stacked frames right
function slide10(pptx) {
  const s = newSlide(pptx, C.paper);
  imageSlot(s, 17.49, 2.01, 7.89, 11.882);
  [2.01, 6.071, 10.132].forEach(function (y) { imageSlot(s, 9.388, y, 7.89, 3.76); });
  runningHead(s, 1.108, C.black);
  txt(s, 'head', COPY.evolution, {
    x: 1.286, y: 2.117, w: 7.164, h: 6.071, color: C.ink, lineSpacingMultiple: 0.8,
  });
  txt(s, 'labelBold', COPY.divide, { x: 1.286, y: 9.097, w: 4.725, h: 0.436 });
  txt(s, 'body', COPY.tradition, { x: 1.286, y: 9.811, w: 6.75, h: 1.478 });
  stamp(s, 1.286, 13.108);
}

// 11 - "MODERNISM" opener
function slide11(pptx) {
  const s = newSlide(pptx, C.paper);
  imageSlot(s, 13.439, 2.01, 11.941, 4.163);
  imageSlot(s, 9.388, 7.679, 5.865, 6.213);
  runningHead(s, 1.108, C.black);
  txt(s, 'hero', 'Modernism', { x: 1.147, y: 1.406, w: 18.647, h: 6.111 });
  txt(s, 'sub', 'Fashion Trends on this year', {
    x: 1.286, y: 7.804, w: 5.554, h: 1.965, color: C.ink, lineSpacingMultiple: 0.8,
  });
  txt(s, 'body', COPY.introLong, { x: 1.286, y: 10.414, w: 6.1, h: 1.825 });
  txt(s, 'sub', COPY.italy, { x: 17.49, y: 12.655, w: 7.89, h: 1.528, color: C.ink, align: 'right' });
}

// 12 - repeating marquee strip along the bottom
function slide12(pptx) {
  const s = newSlide(pptx, C.paper);
  band(s, 0, 11.175, 26.667, 3.825, C.ink);
  txt(s, 'head', COPY.evolution + ' ' + COPY.evolution + ' The Evolution of Modern ', {
    x: -1.623, y: 11.533, w: 29.701, h: 3.109, color: C.paper, align: 'center', lineSpacingMultiple: 0.8,
  });
  stamp(s, 1.286, 5.854);
  txt(s, 'sub', COPY.krux, {
    x: 17.49, y: 7.777, w: 7.89, h: 1.965, color: C.ink, align: 'right', lineSpacingMultiple: 0.8,
  });
  txt(s, 'body', COPY.tradition, { x: 1.286, y: 8.021, w: 6.75, h: 1.478 });
  imageSlot(s, 0, 0, 26.667, 5.255);
}

// 13 - wide photo, caption pair beneath
function slide13(pptx) {
  const s = newSlide(pptx, C.paper);
  imageSlot(s, 1.286, 1.108, 24.095, 7.929);
  stamp(s, 1.286, 9.454);
  txt(s, 'sub', COPY.krux, { x: 1.286, y: 11.95, w: 7.89, h: 1.965, color: C.ink, lineSpacingMultiple: 0.8 });
  txt(s, 'body', COPY.tradition, { x: 18.63, y: 12.194, w: 6.75, h: 1.478 });
}

// 14 - centred "PERFECT TIME"
function slide14(pptx) {
  const s = newSlide(pptx, C.paper);
  imageSlot(s, 9.388, 5.675, 8.113, 3.957);
  txt(s, 'title', 'Perfect time', {
    x: 5.521, y: 0.346, w: 15.413, h: 3.806, align: 'center', lineSpacingMultiple: 0.7,
  });
  stamp(s, 10.971, 10.148, { align: 'center' });
  txt(s, 'label', 'Vision abroad on 2024', { x: 10.971, y: 4.724, w: 4.725, h: 0.436, align: 'center' });
  txt(s, 'sub', COPY.krux, {
    x: 9.388, y: 11.847, w: 7.89, h: 2.172, color: C.ink, align: 'center', lineSpacingMultiple: 0.9,
  });
}

// 15 - full-bleed photo with a dark centre column
function slide15(pptx) {
  const s = newSlide(pptx, C.paper);
  imageSlot(s, 0, 0, 26.667, 15);
  band(s, 9.311, 1.216, 8.044, 12.569, C.ink);
  txt(s, 'sub', COPY.krux, {
    x: 9.388, y: 11.046, w: 7.89, h: 2.172, color: C.paper, align: 'center', lineSpacingMultiple: 0.9,
  });
  stamp(s, 10.971, 1.633, { color: C.paper, align: 'center' });
  txt(s, 'body', COPY.traditionShort, {
    x: 10.732, y: 2.694, w: 5.203, h: 1.131, color: C.paper, align: 'center',
  });
  runningHead(s, 1.108, C.paper);
}

// 16 - large photo, reversed caption block
function slide16(pptx) {
  const s = newSlide(pptx, C.paper);
  imageSlot(s, 1.286, 2.217, 24.095, 12.783);
  runningHead(s, 1.108, C.black);
  stamp(s, 2.193, 9.2, { color: C.paper });
  txt(s, 'sub', COPY.krux, { x: 2.193, y: 11.546, w: 7.89, h: 2.172, color: C.paper, lineSpacingMultiple: 0.9 });
}

// 17 - row of five square frames, copy rail, footer band
function slide17(pptx) {
  const s = newSlide(pptx, C.paper);
  runningHead(s, 1.108, C.black);
  band(s, 0, 13.892, 26.667, 1.108, C.ink);
  txt(s, 'sub', COPY.krux, {
    x: 9.388, y: 3.014, w: 7.89, h: 2.172, color: C.ink, align: 'center', lineSpacingMultiple: 0.9,
  });
  stamp(s, 10.971, 5.637, { align: 'center' });
  stamp(s, 8.039, 7.12, { fontFace: FONT.bodyBold });
  txt(s, 'body', COPY.intro, { x: 8.039, y: 8.042, w: 5.865, h: 1.131 });
  txt(s, 'labelBold', COPY.divide, { x: 8.039, y: 9.76, w: 4.725, h: 0.436 });
  txt(s, 'body', COPY.tradition, { x: 8.039, y: 10.335, w: 5.865, h: 1.825 });
  [15.424, 20.615, 25.806, 2.152, -3.018].forEach(function (x) { imageSlot(s, x, 7.12, 5.04, 5.04); });
}

// 18 - centred chapter title over a strip of frames
function slide18(pptx) {
  const s = newSlide(pptx, C.paper);
  runningHead(s, 1.108, C.black);
  txt(s, 'title', COPY.evolution, {
    x: 2.853, y: 0.948, w: 20.961, h: 5.474, align: 'center', lineSpacingMultiple: 0.9,
  });
  stamp(s, 2.193, 7.88, { fontFace: FONT.bodyBold });
  txt(s, 'body', COPY.intro, { x: 2.193, y: 8.802, w: 5.865, h: 1.131 });
  txt(s, 'labelBold', COPY.divide, { x: 2.193, y: 10.52, w: 4.725, h: 0.436 });
  txt(s, 'body', COPY.tradition, { x: 2.193, y: 11.095, w: 5.865, h: 1.825 });
  imageSlot(s, 8.291, 7.88, 12.132, 6.012);
  imageSlot(s, 20.656, 7.88, 4.958, 6.012);
  imageSlot(s, 25.73, 7.88, 4.958, 6.012);
  imageSlot(s, -2.977, 7.88, 4.725, 6.012);
}

// 19 - closing photo page with the year set huge
function slide19(pptx) {
  const s = newSlide(pptx, C.paper);
  imageSlot(s, 0, 0, 26.667, 11.604);
  txt(s, 'label', COPY.memories, { x: 1.286, y: 12.369, w: 2.703, h: 0.783 });
  txt(s, 'hero', '2024', { x: 15.526, y: 9.095, w: 9.855, h: 6.111, align: 'right' });
  txt(s, 'sub', 'Perfect time for buy some menswears', {
    x: 1.286, y: 3.263, w: 8.369, h: 2.172, color: C.paper, lineSpacingMultiple: 0.9,
  });
  runningHead(s, 1.108, C.paper);
  stamp(s, 1.286, 2.306, { color: C.paper });
}

// 20 - contact details
function slide20(pptx) {
  const s = newSlide(pptx, C.paper);
  txt(s, 'label', 'Contact Us', { x: 1.286, y: 1.108, w: 4.725, h: 0.436 });
  txt(s, 'label', COPY.brand, { x: 20.656, y: 1.108, w: 4.725, h: 0.436, align: 'right' });
  const contacts = [
    { x: 1.286, label: 'Website', value: 'www.launchtech.com', w: 6.043 },
    { x: 5.337, label: 'Email', value: 'info@launchtech.com', w: 6.043 },
    {
      x: 9.388, label: 'Address', w: 5.867, h: 0.783,
      value: 'Launch Tech Inc. 123 Tech Boulevard Innovation City, Techland 54321 United States',
    },
  ];
  contacts.forEach(function (c) {
    txt(s, 'body', c.label, { x: c.x, y: 2.04, w: 2.473, h: 0.436 });
    txt(s, 'labelBold', c.value, { x: c.x, y: c.h ? 2.476 : 2.483, w: c.w, h: c.h || 0.436 });
  });
  imageSlot(s, 1.286, 3.775, 20.043, 10.117);
  imageSlot(s, 21.329, 3.775, 4.051, 10.117);
}

// 21 & 22 - the template's line-icon reference sheets (10 x 5 rings each)
function slide21(pptx) {
  const s = newSlide(pptx, C.white);
  iconGrid(s, 2.67, [2.16, 4.55, 6.85, 9.13, 11.52], 2.221, 1.316, [
    ['\u25A4', '\u2261', '\u2640', '\u2715', '\u2630', '\u2702', 'i', '\u2302', '\u2665', '\u263A'],
    ['\u2699', '\u2699', '\u25CE', '\u25A1', 'P', 'D', '\u2642', '\u21C4', '\u25A4', '\u2193'],
    ['\u2193', '\u2299', '\u00A7', '\u2601', '\u2193', '\u2713', '\u2298', '\u2715', '\u25A3', '\u25A6'],
    ['\u2691', '\u266A', '\u2197', '\u002B', '\u2191', '\u25A5', '\u21BA', '\u21BB', '\u2192', '\u21AA'],
    ['\u21A9', '\u2190', '\u2197', '\u2193', '\u2199', '\u2261', '\u0021', '\u0040', '\u002B', '\u2295'],
  ]);
}

function slide22(pptx) {
  const s = newSlide(pptx, C.white);
  iconGrid(s, 1.47, [1.72, 4.32, 6.74, 9.42, 12.0], 2.496, 1.282, [
    ['\u2261', '\u0021', '\u25B6', '\u263A', '\u2191', '\u2691', 'I', 'T', '\u2299', '\u25BC'],
    ['T', '\u2295', '\u2601', '\u2600', '\u25A4', '\u25A3', '\u25C6', '\u25CE', '\u25B2', '\u25BC'],
    ['\u25A3', '\u2299', '\u2212', '\u263B', '\u266B', '\u003F', '\u23FB', '\u25B6', '\u2691', '\u260E'],
    ['\u25B2', '\u25A1', '\u270E', '\u2712', '\u25C6', '\u2261', '\u2630', '\u2261', '\u2630', '\u2699'],
    ['\u2302', '\u25A3', '\u2026', '\u2699', '\u25AA', '\u266A', '\u266B', '\u2299', '\u25A4', '\u263A'],
  ]);
}

/* ---------------------------------------------------------------------- build */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'LOOKBOOK', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'LOOKBOOK';
  pptx.theme = { headFontFace: FONT.display, bodyFontFace: FONT.body };
  pptx.title = 'Fashion Lookbook Presentation Template';

  [
    slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
    slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
    slide17, slide18, slide19, slide20, slide21, slide22,
  ].forEach(function (buildSlide) { buildSlide(pptx); });

  return pptx.writeFile({
    fileName: path.join(__dirname, '00f4406b-5bd5-40e4-b70e-edc9869ba59e_grok_final.pptx'),
  });
}

build().then(function (f) { console.log('wrote ' + f); });
