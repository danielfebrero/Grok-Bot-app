/**
 * "Furniture Lookbook" — interior presentation template.
 * 10 slides, 16:9 widescreen (13.333in x 7.5in), cream/brown editorial palette.
 *
 * Standalone pptxgenjs re-creation of the reference deck. The reference carries
 * no raster media at all — its photo slots are empty picture placeholders — so
 * they are re-created here as unpainted rectangles of identical geometry.
 *
 *   node 01a99be8-cc22-495b-aa2b-33e5999611c9_grok_final.js
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const C = {
  cream: 'F6E7D4', // theme accent2 — slide background, light panels, reversed text
  brown: '44372D', // theme accent1 — headings, dark panels, icons
  body: '262626', // tx1 @ 85% luminance — body copy
  white: 'FFFFFF', // body copy reversed out of the brown cards
};

const F = {
  serif: 'Playfair Display', // display face: titles and the wordmark
  sans: 'Work Sans', // text face: everything else
};

const SZ = {
  display: 96, // "Furniture Lookbook" / "Thank You!"
  title: 37, // section titles
  lead: 18, // bold sub-headings (deck default size)
  button: 16, // pill/button captions on slide 10
  tab: 14, // button captions on slide 1
  brand: 12, // nav wordmark
  body: 10.5, // body copy, nav links, footer
};

const NONE = { type: 'none' }; // pptxgenjs "no fill" / "no line"
const INSETS = [7.2, 7.2, 3.6, 3.6]; // PowerPoint default 0.1in / 0.05in text insets

/* ---------------------------------------------------------------- helpers */

/** Auto-sizing text box with no fill — the deck's universal text primitive. */
function label(slide, text, o) {
  slide.addText(text, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: o.font || F.sans,
    fontSize: o.size || SZ.body,
    color: o.color || C.body,
    bold: o.bold || false,
    align: o.align || 'left',
    valign: 'top',
    wrap: !!o.wrap,
    lineSpacingMultiple: o.lineSpacing,
    margin: INSETS,
    fit: 'resize',
  });
}

/** Justified 10.5pt copy at 1.5 line spacing — the deck's standard paragraph. */
function bodyCopy(slide, text, o) {
  label(slide, text, { ...o, size: SZ.body, align: 'justify', wrap: true, lineSpacing: 1.5 });
}

/** Section title set in Playfair Display. */
function title(slide, text, o) {
  label(slide, text, { ...o, font: F.serif, size: SZ.title, color: C.brown, wrap: true });
}

/** Bold sub-heading ("1. Vision Title", "About Our Planning", "2. Service", ...). */
function subhead(slide, text, o) {
  label(slide, text, { ...o, size: SZ.lead, bold: true, color: o.color || C.brown });
}

/** Filled rectangle, optionally carrying a centred caption. */
function panel(slide, o) {
  const opts = {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: { color: o.fill },
    line: o.line ? { color: o.line, width: 1 } : NONE,
  };
  if (o.text === undefined) {
    slide.addShape('rect', opts);
  } else {
    slide.addText(o.text, {
      ...opts, shape: 'rect', margin: INSETS,
      fontFace: F.sans, fontSize: o.size || SZ.tab, color: o.color || C.cream,
      bold: !!o.bold, align: 'center', valign: 'middle',
    });
  }
}

/**
 * One of the deck's photo frames. The reference ships these as *empty* picture
 * placeholders — no raster data — so they reserve space without painting
 * anything. Reproduced here as unfilled rectangles of identical geometry.
 */
function pictureFrame(slide, o) {
  slide.addShape('rect', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: NONE, line: NONE,
    objectName: '[image] frame',
  });
}

/** Free-form polygon; `pts` are [x, y] fractions of the w x h bounding box. */
function poly(slide, o) {
  slide.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: { color: o.fill }, line: NONE,
    points: [...o.pts.map(([px, py]) => ({ x: px * o.w, y: py * o.h })), { close: true }],
  });
}

function newSlide(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.cream };
  return s;
}

/* -------------------------------------------------- chrome shared by every slide */

const NAV_LEFT = [
  { text: 'Profile', x: 1.053, w: 0.665 },
  { text: 'About', x: 2.199, w: 0.631 },
  { text: 'Lookbook', x: 3.312, w: 0.901 },
];
const NAV_RIGHT = [
  { text: 'News', x: 9.86, y: 0.458, w: 0.595 },
  { text: 'Contact', x: 10.899, y: 0.463, w: 0.772 },
];

/** Small "user account" glyph closing the navigation bar: head over shoulders. */
function accountIcon(slide) {
  const x = 11.974, y = 0.481, w = 0.21, h = 0.233;
  const ink = { color: C.brown };
  slide.addShape('ellipse', { x: x + w * 0.24, y: y, w: w * 0.52, h: h * 0.5, fill: ink, line: NONE });
  slide.addShape('round2SameRect', {
    x: x + w * 0.06, y: y + h * 0.52, w: w * 0.88, h: h * 0.48, rectRadius: 0.035,
    fill: ink, line: NONE,
  });
}

/** Top navigation strip. On slide 8 the left links sit on a brown panel. */
function navBar(slide, leftColor) {
  NAV_LEFT.forEach(n => label(slide, n.text, { x: n.x, y: 0.463, w: n.w, h: 0.278, color: leftColor }));
  NAV_RIGHT.forEach(n => label(slide, n.text, { x: n.x, y: n.y, w: n.w, h: 0.278 }));
  label(slide, 'Furniture Lookbook', {
    x: 5.8, y: 0.451, w: 1.732, h: 0.303,
    font: F.serif, size: SZ.brand, color: C.brown, align: 'center',
  });
  accountIcon(slide);
}

function chrome(slide, opts = {}) {
  const leftColor = opts.leftColor || C.body;
  navBar(slide, leftColor);
  label(slide, 'Best Interior 2025', { x: 1.053, y: 6.717, w: 1.497, h: 0.278, color: leftColor });
  label(slide, 'Furniture', { x: 11.415, y: 6.717, w: 0.868, h: 0.278 });
}

/* ------------------------------------------------------- decorative icons */

/** Ellipse positioned with fractional coordinates inside an x/y/w/h box. */
function dot(slide, box, fx, fy, fw, fh, color) {
  slide.addShape('ellipse', {
    x: box.x + box.w * fx, y: box.y + box.h * fy, w: box.w * fw, h: box.h * fh,
    fill: { color }, line: NONE,
  });
}

/** Slide 7 service badges: brown glyphs sitting on the cream corner tile. */
function serviceIcon(slide, kind, x, y, w, h) {
  const box = { x, y, w, h };
  const ink = { color: C.brown };
  if (kind === 'documents') {
    // Clipboard behind, then a sheet whose top-right corner is folded back.
    slide.addShape('roundRect', { x: x, y: y + h * 0.03, w: w * 0.6, h: h * 0.76, rectRadius: 0.02, fill: ink, line: NONE });
    slide.addShape('roundRect', {
      x: x + w * 0.13, y: y + h * 0.12, w: w * 0.34, h: h * 0.09, rectRadius: 0.01,
      fill: { color: C.cream }, line: NONE,
    });
    poly(slide, { ...box, fill: C.brown, pts: [[0.36, 0.28], [0.70, 0.28], [1.0, 0.58], [1.0, 1.0], [0.36, 1.0]] });
    poly(slide, { ...box, fill: C.cream, pts: [[0.45, 0.37], [0.64, 0.37], [0.64, 0.62], [0.91, 0.62], [0.91, 0.91], [0.45, 0.91]] });
    poly(slide, { ...box, fill: C.cream, pts: [[0.70, 0.34], [0.93, 0.57], [0.70, 0.57]] });
  } else if (kind === 'cloud') {
    // Three lobes over a rounded slab, with a download arrow knocked out.
    dot(slide, box, 0.18, 0.0, 0.5, 0.72, C.brown);
    dot(slide, box, 0.56, 0.22, 0.36, 0.52, C.brown);
    dot(slide, box, 0.0, 0.28, 0.36, 0.54, C.brown);
    slide.addShape('roundRect', { x: x, y: y + h * 0.4, w: w, h: h * 0.6, rectRadius: 0.09, fill: ink, line: NONE });
    slide.addShape('downArrow', {
      x: x + w * 0.3, y: y + h * 0.26, w: w * 0.36, h: h * 0.64,
      fill: { color: C.cream }, line: NONE,
    });
  } else if (kind === 'gears') {
    // A large cogwheel plus two smaller ones, each with a hollow hub.
    slide.addShape('gear9', { x: x, y: y + h * 0.18, w: w * 0.66, h: h * 0.82, fill: ink, line: NONE });
    dot(slide, box, 0.2, 0.44, 0.26, 0.3, C.cream);
    slide.addShape('gear9', { x: x + w * 0.58, y: y, w: w * 0.42, h: h * 0.48, fill: ink, line: NONE });
    dot(slide, box, 0.72, 0.16, 0.14, 0.16, C.cream);
    slide.addShape('gear9', { x: x + w * 0.56, y: y + h * 0.52, w: w * 0.42, h: h * 0.48, fill: ink, line: NONE });
    dot(slide, box, 0.7, 0.68, 0.14, 0.16, C.cream);
  }
}

/** Slide 9 contact badges: cream glyphs on the brown square. */
function contactIcon(slide, kind, x, y, w, h) {
  const box = { x, y, w, h };
  if (kind === 'phone') {
    // Handset: a thick arc for the cord with a squared earpiece at each end.
    slide.addShape('roundRect', { x: x, y: y, w: w, h: h, rectRadius: 0.055, fill: { color: C.cream }, line: NONE });
    slide.addShape('blockArc', {
      x: x + w * 0.18, y: y + h * 0.18, w: w * 0.64, h: h * 0.64,
      angleRange: [35, 215], arcThicknessRatio: 0.46,
      fill: { color: C.brown }, line: NONE,
    });
    slide.addShape('roundRect', { x: x + w * 0.19, y: y + h * 0.17, w: w * 0.26, h: h * 0.32, rectRadius: 0.06, rotate: -18, fill: { color: C.brown }, line: NONE });
    slide.addShape('roundRect', { x: x + w * 0.53, y: y + h * 0.52, w: w * 0.32, h: h * 0.26, rectRadius: 0.06, rotate: -18, fill: { color: C.brown }, line: NONE });
  } else if (kind === 'plane') {
    // Paper plane: upper wing plus the folded-under tail fin, split by a crease.
    poly(slide, { ...box, fill: C.cream, pts: [[0.02, 0.52], [0.98, 0.04], [0.34, 0.66]] });
    poly(slide, { ...box, fill: C.cream, pts: [[0.38, 0.70], [0.98, 0.06], [0.46, 0.99], [0.38, 0.86]] });
  } else if (kind === 'globe') {
    // Cream disc, brown ocean, cream landmasses either side of the meridian.
    dot(slide, box, 0, 0, 1, 1, C.cream);
    dot(slide, box, 0.1, 0.1, 0.8, 0.8, C.brown);
    poly(slide, { ...box, fill: C.cream, pts: [[0.12, 0.32], [0.22, 0.22], [0.34, 0.28], [0.46, 0.24], [0.48, 0.38], [0.36, 0.46], [0.42, 0.60], [0.32, 0.78], [0.26, 0.58], [0.16, 0.46]] });
    poly(slide, { ...box, fill: C.cream, pts: [[0.56, 0.18], [0.74, 0.24], [0.66, 0.36], [0.82, 0.42], [0.78, 0.58], [0.62, 0.68], [0.54, 0.48], [0.60, 0.30]] });
    poly(slide, { ...box, fill: C.cream, pts: [[0.30, 0.80], [0.68, 0.78], [0.60, 0.89], [0.38, 0.89]] });
  }
}

/* ---------------------------------------------------------------- copy deck */

const LOREM = {
  short: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut eiusmod tempor',
  medium: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut eiusmod tempor incididunt ut consectetur consec',
  long: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut eiusmod tempor incididunt ut consectetur elit Aliquam tincidunt ante nec sem congue convallis. Pellentesque vel mauris quis nisl ornare rutrum in id risus. Proin vehicula ut sem et tempus. Interdum et malesuada fames ac ante nec sem congue convallis. ',
  facility: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, do eiusmod tempor',
  product: 'Lorem ipsum dolor sit amet, consectetur elit, sed mauris do eiusmod ',
  productLead: 'Lorem ipsum dolor sit amet, consectetur elit, sed mauris do eiusmod eiusmod tempor',
  service: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ut',
  portfolio: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut eiusmod tempor incididunt ut consectetur consec adipiscing elit sed do ',
  contact: 'Lorem ipsum dolor sit amet, consectetur elit, sed mauris do elit, sed',
};

/* --------------------------------------------------------- slide builders */

/** 1 — cover: oversized wordmark, two brown buttons, portrait photo frame. */
function slide01(pptx) {
  const s = newSlide(pptx);
  chrome(s);
  pictureFrame(s, { x: 9.769, y: 1.886, w: 2.415, h: 4.245 });
  label(s, 'Furniture', { x: 0.989, y: 1.665, w: 6.087, h: 1.717, font: F.serif, size: SZ.display, color: C.brown });
  label(s, 'Lookbook', { x: 2.831, y: 3.217, w: 6.183, h: 1.717, font: F.serif, size: SZ.display, color: C.brown });
  panel(s, { x: 1.149, y: 5.602, w: 3.946, h: 0.494, fill: C.brown, text: 'Interior Presentation Template' });
  panel(s, { x: 9.769, y: 1.392, w: 2.415, h: 0.494, fill: C.brown, text: 'Start Presentation' });
}

/** 2 — welcome: photo left, headline + two paragraphs right, brown edge tab. */
function slide02(pptx) {
  const s = newSlide(pptx);
  chrome(s);
  pictureFrame(s, { x: 1.167, y: 1.403, w: 4.736, h: 4.694 });
  title(s, 'Welcome To Furniture Lookbook', { x: 7.376, y: 1.253, w: 5.053, h: 1.346 });
  bodyCopy(s, LOREM.short, { x: 7.338, y: 2.956, w: 4.945, h: 0.601 });
  subhead(s, 'About Our Planning', { x: 7.371, y: 3.964, w: 2.835, h: 0.404, wrap: true });
  bodyCopy(s, LOREM.long, { x: 7.36, y: 4.531, w: 4.945, h: 1.662 });
  // Brown tab bleeding off the left edge of the slide
  panel(s, { x: 0, y: 3.128, w: 2.705, h: 1.201, fill: C.brown });
  label(s, 'Furniture', { x: 0.614, y: 3.486, w: 1.476, h: 0.438, size: 20, color: C.cream, align: 'center' });
  s.addShape('line', { x: 0.503, y: 3.972, w: 1.699, h: 0, line: { color: C.cream, width: 0.5 } });
}

/** 3 — vision & mission: two numbered blocks left, stacked photo frames right. */
function slide03(pptx) {
  const s = newSlide(pptx);
  chrome(s);
  pictureFrame(s, { x: 6.667, y: 1.389, w: 5.517, h: 4.764 });
  panel(s, { x: 9.421, y: 3.768, w: 2.763, h: 2.386, fill: C.cream });
  pictureFrame(s, { x: 9.588, y: 3.935, w: 2.596, h: 2.219 });
  title(s, 'Vision & Mission', { x: 1.074, y: 1.221, w: 5.053, h: 0.724 });
  [
    { head: '1.  Vision Title', hx: 1.062, hy: 2.604, hw: 1.937, bx: 1.073, by: 3.213 },
    { head: '2. Mission Title', hx: 1.055, hy: 4.678, hw: 2.101, bx: 1.066, by: 5.287 },
  ].forEach(block => {
    subhead(s, block.head, { x: block.hx, y: block.hy, w: block.hw, h: 0.404 });
    bodyCopy(s, LOREM.medium, { x: block.bx, y: block.by, w: 4.329, h: 0.867 });
  });
}

/** 4 — facilities: twin photo columns left, two brown cards right. */
function slide04(pptx) {
  const s = newSlide(pptx);
  chrome(s);
  pictureFrame(s, { x: 1.158, y: 1.4, w: 2.552, h: 4.713 });
  pictureFrame(s, { x: 4.115, y: 1.4, w: 2.552, h: 4.713 });
  title(s, 'Best Our Facilities', { x: 7.354, y: 1.221, w: 5.053, h: 0.724 });
  [
    { head: '1.  Facilites', y: 2.573, hy: 2.806, by: 3.253 },
    { head: '2.  Facilites', y: 4.541, hy: 4.775, by: 5.221 },
  ].forEach(card => {
    panel(s, { x: 7.474, y: card.y, w: 4.707, h: 1.551, fill: C.brown });
    subhead(s, card.head, { x: 7.976, y: card.hy, w: 1.724, h: 0.404, color: C.cream, wrap: true });
    bodyCopy(s, LOREM.facility, { x: 8.007, y: card.by, w: 3.674, h: 0.601, color: C.white });
  });
}

/** 5 — products: wide photo left, 01/02/03 numbered rows right. */
function slide05(pptx) {
  const s = newSlide(pptx);
  chrome(s);
  pictureFrame(s, { x: 1.179, y: 2.516, w: 6.238, h: 3.622 });
  title(s, 'Best Our Product', { x: 1.074, y: 1.221, w: 5.053, h: 0.724 });
  bodyCopy(s, LOREM.productLead, { x: 8.045, y: 1.282, w: 4.215, h: 0.601 });
  [
    { n: '01', bx: 8.151, by: 2.516, tx: 9.109, ty: 2.559 },
    { n: '02', bx: 8.153, by: 3.94, tx: 9.111, ty: 3.983 },
    { n: '03', bx: 8.154, by: 5.364, tx: 9.113, ty: 5.407 },
  ].forEach(row => {
    panel(s, { x: row.bx, y: row.by, w: 0.774, h: 0.774, fill: C.brown, text: row.n, size: 24, bold: true });
    bodyCopy(s, LOREM.product, { x: row.tx, y: row.ty, w: 3.169, h: 0.601 });
  });
}

/** 6 — team: three brown portrait tiles above outlined name plates. */
function slide06(pptx) {
  const s = newSlide(pptx);
  chrome(s);
  title(s, 'Meet Our Team', { x: 4.586, y: 1.221, w: 4.162, h: 0.724, align: 'center' });
  [
    { name: 'Alfredo Torres', tileX: 1.153, plateX: 1.153, frame: { x: 1.319, w: 3.266, h: 2.625 } },
    { name: 'Brigitte Schwartz', tileX: 4.95, plateX: 4.95, frame: { x: 4.95, w: 3.433, h: 2.464 } },
    { name: 'Eleanor Fitzgerald', tileX: 8.739, plateX: 8.748, frame: { x: 8.739, w: 3.266, h: 2.625 } },
  ].forEach(member => {
    panel(s, { x: member.tileX, y: 2.514, w: 3.433, h: 2.792, fill: C.brown });
    pictureFrame(s, { ...member.frame, y: 2.681 });
    panel(s, {
      x: member.plateX, y: 5.625, w: 3.433, h: 0.611,
      fill: C.cream, line: C.brown,
      text: member.name, color: C.brown, size: SZ.lead, bold: true,
    });
  });
}

/** 7 — services: three brown cards, each with a cream icon tile in the corner. */
function slide07(pptx) {
  const s = newSlide(pptx);
  chrome(s);
  pictureFrame(s, { x: 5.885, y: 1.375, w: 6.284, h: 2.367 });
  title(s, 'Best Our Service', { x: 1.074, y: 1.221, w: 4.358, h: 0.724 });
  bodyCopy(s, LOREM.service, { x: 1.074, y: 2.374, w: 4.104, h: 0.601 });
  [
    { head: '1. Service', card: 1.179, headX: 2.185, headW: 1.378, copyX: 1.627, icon: 'documents', ix: 1.405, iy: 3.992, iw: 0.314, ih: 0.314 },
    { head: '2. Service', card: 4.879, headX: 5.885, headW: 1.42, copyX: 5.327, icon: 'cloud', ix: 5.091, iy: 4.025, iw: 0.34, ih: 0.248 },
    { head: '3. Service', card: 8.622, headX: 9.628, headW: 1.419, copyX: 9.07, icon: 'gears', ix: 8.834, iy: 3.994, iw: 0.34, ih: 0.31 },
  ].forEach(card => {
    panel(s, { x: card.card, y: 3.743, w: 3.547, h: 2.331, fill: C.brown });
    panel(s, { x: card.card, y: 3.742, w: 0.765, h: 0.814, fill: C.cream });
    serviceIcon(s, card.icon, card.ix, card.iy, card.iw, card.ih);
    subhead(s, card.head, { x: card.headX, y: 4.172, w: card.headW, h: 0.404, color: C.cream });
    bodyCopy(s, LOREM.service, { x: card.copyX, y: 4.874, w: 2.65, h: 0.867, color: C.white });
  });
}

/** 8 — portfolio: full-bleed brown panel left, photo mosaic + copy right. */
function slide08(pptx) {
  const s = newSlide(pptx);
  panel(s, { x: 0, y: 0, w: 5.356, h: 7.5, fill: C.brown });
  chrome(s, { leftColor: C.cream });
  [
    { x: 1.168, y: 1.381, w: 5.498, h: 2.251 },
    { x: 1.168, y: 3.863, w: 2.623, h: 2.251 },
    { x: 4.044, y: 3.863, w: 2.623, h: 2.251 },
    { x: 6.91, y: 1.381, w: 2.507, h: 2.251 },
    { x: 9.658, y: 1.381, w: 2.507, h: 2.251 },
  ].forEach(frame => pictureFrame(s, frame));
  title(s, 'Best Our Portfolio', { x: 7.333, y: 4.089, w: 5.053, h: 0.724 });
  bodyCopy(s, LOREM.portfolio, { x: 7.333, y: 4.898, w: 4.949, h: 0.867 });
}

/** 9 — contact: three icon + outlined-field rows left, photo frames right. */
function slide09(pptx) {
  const s = newSlide(pptx);
  chrome(s);
  panel(s, { x: 5.379, y: 2.483, w: 3.653, h: 2.552, fill: C.cream });
  pictureFrame(s, { x: 6.667, y: 1.402, w: 5.517, h: 4.714 });
  pictureFrame(s, { x: 5.546, y: 2.653, w: 3.315, h: 2.211 });
  title(s, 'Contact Us', { x: 1.074, y: 1.221, w: 3.242, h: 0.724 });
  bodyCopy(s, LOREM.contact, { x: 1.054, y: 2.411, w: 3.342, h: 0.601 });
  [
    { icon: 'phone', y: 3.563, text: '+123 4567 8910', ty: 3.651, tw: 1.334, th: 0.286, ix: 1.268, iy: 3.683, iw: 0.224, ih: 0.223 },
    { icon: 'plane', y: 4.602, text: 'yourmail@mail.com', ty: 4.694, tw: 1.643, th: 0.278, ix: 1.25, iy: 4.703, iw: 0.26, ih: 0.26 },
    { icon: 'globe', y: 5.641, text: 'www.yourwebsite.com', ty: 5.733, tw: 1.904, th: 0.278, ix: 1.257, iy: 5.748, iw: 0.247, ih: 0.249 },
  ].forEach(row => {
    panel(s, { x: 1.149, y: row.y, w: 0.462, h: 0.462, fill: C.brown });
    panel(s, { x: 1.732, y: row.y, w: 2.334, h: 0.462, fill: C.cream, line: C.brown });
    contactIcon(s, row.icon, row.ix, row.iy, row.iw, row.ih);
    label(s, row.text, { x: 1.864, y: row.ty, w: row.tw, h: row.th });
  });
}

/** 10 — closing: two buttons framing the oversized "Thank You!" wordmark. */
function slide10(pptx) {
  const s = newSlide(pptx);
  chrome(s);
  pictureFrame(s, { x: 0, y: 2.892, w: 2.434, h: 1.717 });
  pictureFrame(s, { x: 10.899, y: 2.892, w: 2.434, h: 1.717 });
  panel(s, { x: 5.141, y: 1.981, w: 3.051, h: 0.537, fill: C.brown, text: 'See You Next Time', size: SZ.button });
  label(s, 'Thank You!', {
    x: 3.144, y: 2.892, w: 7.046, h: 1.717,
    font: F.serif, size: SZ.display, color: C.brown, align: 'center',
  });
  panel(s, { x: 5.679, y: 5.19, w: 1.975, h: 0.537, fill: C.brown, text: 'The End', size: SZ.button });
  const rule = { color: C.brown, width: 0.5, endArrowType: 'triangle' };
  s.addShape('line', { x: 8.254, y: 5.459, w: 1.262, h: 0.015, line: rule });
  s.addShape('line', { x: 3.818, y: 5.474, w: 1.262, h: 0.015, flipH: true, line: rule });
}

/* ------------------------------------------------------------------- main */

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE_16x9', width: 13.3333333, height: 7.5 }); // 12192000 x 6858000 EMU
pptx.layout = 'WIDE_16x9';
pptx.title = 'Furniture Lookbook';

[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10]
  .forEach(buildSlide => buildSlide(pptx));

pptx.writeFile({ fileName: path.join(__dirname, '01a99be8-cc22-495b-aa2b-33e5999611c9_grok_final.pptx') })
  .then(file => console.log('wrote', file));
