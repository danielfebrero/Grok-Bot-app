/**
 * "Machiato" — coffee shop business presentation template.
 * 15 slides, 16:9 widescreen (13.333in x 7.5in). Rebuilt with pptxgenjs only.
 *
 * Imagery note: the source file is a *template* — every photo area is an empty
 * picture placeholder that renders transparent, so those regions are reproduced
 * as the bare background they actually show (some carry the small "img" prompt).
 * The one genuinely embedded raster, the phone mock-up on slide 13, is replaced
 * by a programmatic placeholder made of native shapes and labelled "[image]".
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------------------
// Palette & typography
// ---------------------------------------------------------------------------
const C = {
  brown: 'BC713E',       // primary accent
  brownDark: 'A26336',   // buttons, rings, panels
  brownRust: 'B45A38',   // eyebrow labels, prices
  charcoal: '383838',    // dark slide background
  white: 'FFFFFF',
  cream: 'FBF7F5',       // 5% brown wash over white
  ink: '262626',         // headings          (tx1 lumMod 85%)
  inkSoft: '404040',     // sub-heads         (tx1 lumMod 75%)
  body: '808080',        // body copy         (tx1 lumMod 50%)
  bodyLight: 'F2F2F2',   // body copy on dark (bg1 lumMod 95%)
  bodyLight2: 'D9D9D9',  // body copy on dark (bg1 lumMod 85%)
  rule: 'A6A6A6',
  star: 'FFC000',
  prompt: 'F6F2F8',      // 5% purple hatch of the empty "img" placeholders
};

const F = {
  logo: 'Comfortaa SemiBold',   // "Machiato" wordmark
  head: 'Montserrat SemiBold',  // slide headlines
  sans: 'Montserrat',
  body: 'Lato',                 // paragraph copy
  label: 'Poppins',             // ALL-CAPS sub-heads
  labelMed: 'Poppins Medium',   // buttons, prices, roles
  eyebrow: 'Work Sans Medium',  // "About Us" / "Our Services"
  name: 'Work Sans SemiBold',   // people names
};

// pptxgenjs rewrites the shadow object it is handed, so hand it a fresh one each time.
const cardShadow = (opacity, blur) =>
  ({ type: 'outer', color: '000000', opacity, blur, offset: 0, angle: 90, rotateWithShape: false });
const SOFT = () => cardShadow(0.07, 10);   // large diffuse card shadow
const LIFT = () => cardShadow(0.1, 9);     // tighter shadow on small elements

// Lorem strings reused across several slides.
const LOREM_LONG =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ' +
  'ut labore et dolore magna aliqua. Suspendisse potenti nullam ac tortor vitae purus faucibus ' +
  'ornare. Tellus elementum sagittis vitae et leo duis ut diam quam. ';
const LOREM_TAIL =
  'Nisl condimentum id venenatis. Feugiat sed lectus vestibulum mattis ullamcorper velit sed ' +
  'potenti nullam ac tortor vitae. ';
const DUIS =
  'PLACEHOLDER' +
  'PLACEHOLDER';
const LOREM_MIX =
  'PLACEHOLDER' +
  'adipiscing elita sed do eiusmod tempor';
const VOLUPTATE = 'PLACEHOLDER';
const CARD_BLURB = 'Duis  irure dolor reprehenderit voluptate velcillum dolor.';
const CONSECTETUR = 'Consectetur adipiscing siabes im doloremag faubicus  eiusmod tempor incididunt ut labore.';
const VELITESSE = 'voluptate velitesse clum ipsum dolor sit anim ipsum doloria nostrud minimasa veniam nostrud.';
const TESTIMONIAL =
  '\u201CLorem dosiamet, consectetur adipiscing elit, eiusmod tempor labore dolore magna aliqua. ' +
  'Suspendisse potenti nullam ac tortor vitae purus faucibus ornare tempor\u201C';

// ---------------------------------------------------------------------------
// Helpers — thin wrappers over native pptxgenjs calls
// ---------------------------------------------------------------------------
const rect = (s, o) => s.addShape('rect', o);
const roundRect = (s, o) => s.addShape('roundRect', o);
const oval = (s, o) => s.addShape('ellipse', o);
const line = (s, o) => s.addShape('line', o);
const freeform = (s, points, o) => s.addShape('custGeom', Object.assign({ points }, o));

/** Paragraph copy: 11 pt Lato, 1.5 line spacing, justified by default. */
function body(s, text, o) {
  s.addText(text, Object.assign({
    fontFace: F.body, fontSize: 11, color: C.body,
    align: 'justify', lineSpacingMultiple: 1.5, valign: 'top',
  }, o));
}

/** ALL-CAPS bold Poppins sub-head. */
function label(s, text, o) {
  s.addText(text, Object.assign({
    fontFace: F.label, fontSize: 12, bold: true, charSpacing: 1,
    color: C.inkSoft, valign: 'top',
  }, o));
}

/** Slide headline: 36 pt Montserrat SemiBold. */
function heading(s, text, o) {
  s.addText(text, Object.assign({
    fontFace: F.head, fontSize: 36, color: C.ink, valign: 'top',
  }, o));
}

/** Small brown eyebrow line above a headline ("About Us", "Our Services"…). */
function eyebrow(s, text, x, y, color) {
  s.addText(text, {
    x, y, w: 1.6, h: 0.353, fontFace: F.eyebrow, fontSize: 15,
    color: color || C.brownRust, valign: 'top',
  });
}

/** 1.143in x 0.05in accent bar sitting under several headlines. */
const accentBar = (s, x, y, color) =>
  rect(s, { x, y, w: 1.143, h: 0.05, fill: { color: color || C.brown } });

/** Pill button with centred, letter-spaced caps. */
function button(s, text, x, y, w, color) {
  roundRect(s, { x, y, w, h: 0.391, rectRadius: 0.05, fill: { color: color || C.brown } });
  s.addText(text, {
    x, y, w, h: 0.391, align: 'center', valign: 'middle',
    fontFace: F.labelMed, fontSize: 11, charSpacing: 1.5, color: C.white,
  });
}

/**
 * Empty picture placeholder from the template: a 5%-hatch frame carrying the
 * "img" prompt plus PowerPoint's picture glyph (framed sky, sun and mountains).
 */
function imgPrompt(s, o) {
  const gs = o.glyphScale || 0.62;
  const g = { x: o.x + o.w * (1 - gs) / 2, y: o.y + o.h * 0.28, w: o.w * gs, h: o.h * gs * 0.86 };
  s.addShape(o.shape || 'rect', { x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: C.prompt } });
  s.addText('img', {
    x: o.x, y: o.y + o.h * 0.06, w: o.w, h: o.h * 0.24, align: 'center', valign: 'top',
    fontFace: 'Arial', fontSize: 10.5, color: C.inkSoft,
  });
  rect(s, { x: g.x, y: g.y, w: g.w, h: g.h, fill: { color: C.white }, line: { color: C.body, width: 0.75 } });
  rect(s, { x: g.x + g.w * 0.08, y: g.y + g.h * 0.1, w: g.w * 0.84, h: g.h * 0.8, fill: { color: 'EAF3FB' } });
  oval(s, { x: g.x + g.w * 0.16, y: g.y + g.h * 0.2, w: g.w * 0.2, h: g.h * 0.22, fill: { color: 'F7C873' } });
  s.addShape('triangle', { x: g.x + g.w * 0.1, y: g.y + g.h * 0.48, w: g.w * 0.4, h: g.h * 0.42, fill: { color: '78ADDD' } });
  s.addShape('triangle', { x: g.x + g.w * 0.36, y: g.y + g.h * 0.3, w: g.w * 0.56, h: g.h * 0.6, fill: { color: '78ADDD' } });
}

/**
 * Programmatic stand-in for the deck's one embedded raster (a phone mock-up):
 * dark bezel, white screen and the notch, all from native shapes.
 */
function phonePlaceholder(s, o) {
  const { x, y, w, h } = o;
  const bezel = 0.18; // the screen stays open so the brown disc shows through, as in the original
  roundRect(s, {
    x: x + bezel / 2, y: y + bezel / 2, w: w - bezel, h: h - bezel, rectRadius: 0.36,
    fill: { type: 'none' }, line: { color: '1C1C1C', width: bezel * 72 },
  });
  roundRect(s, { // speaker notch
    x: x + w / 2 - 0.7, y: y + bezel - 0.02, w: 1.4, h: 0.15, rectRadius: 0.06, fill: { color: '1C1C1C' },
  });
  s.addText('[image]', {
    x, y: y + h - 0.62, w, h: 0.32, align: 'center', valign: 'middle',
    fontFace: F.body, fontSize: 11, color: C.body,
  });
}

/**
 * Vertical gradient faked with horizontal bands — the cover slides use a
 * black gradient (45% -> 65% alpha) over white, i.e. gray #8C8C8C -> #595959.
 */
function verticalGradient(s, from, to, steps) {
  const lerp = (a, b, t) => Math.round(a + (b - a) * t);
  const hex = (v) => v.toString(16).toUpperCase().padStart(2, '0');
  const c0 = [0, 2, 4].map((i) => parseInt(from.substr(i, 2), 16));
  const c1 = [0, 2, 4].map((i) => parseInt(to.substr(i, 2), 16));
  const bandH = 7.5 / steps;
  for (let i = 0; i < steps; i++) {
    const t = i / (steps - 1);
    rect(s, {
      x: 0, y: i * bandH, w: 13.333, h: bandH + 0.01,
      fill: { color: c0.map((v, k) => hex(lerp(v, c1[k], t))).join('') },
    });
  }
}

/**
 * The template's coffee-bean icon: three outlined beans with a diagonal crease,
 * used as a watermark. Sizes are fractions of the icon's 96x96 viewBox.
 */
function coffeeBeans(s, o) {
  const { x, y, w, h, color, transparency = 0, rotate = 0 } = o;
  const stroke = Math.min(5, Math.max(0.9, w * 1.15));
  const beans = [
    { cx: 0.714, cy: 0.293, w: 0.316, h: 0.378, rot: 20 },
    { cx: 0.626, cy: 0.731, w: 0.382, h: 0.345, rot: -35 },
    { cx: 0.286, cy: 0.440, w: 0.366, h: 0.358, rot: 35 },
  ];
  beans.forEach((b) => {
    const bw = b.w * w, bh = b.h * h;
    const bx = x + b.cx * w - bw / 2, by = y + b.cy * h - bh / 2;
    const pen = { color, width: stroke, transparency };
    oval(s, { x: bx, y: by, w: bw, h: bh, rotate: b.rot + rotate, fill: { type: 'none' }, line: pen });
    line(s, {
      x: bx + bw * 0.16, y: by + bh * 0.84, w: bw * 0.68, h: -bh * 0.68,
      rotate: b.rot + rotate, line: pen,
    });
  });
}

/** Row of small 5-point stars (rating widget). */
function stars(s, o) {
  const size = o.size || 0.175, pitch = o.pitch || 0.1715;
  for (let i = 0; i < o.count; i++) {
    s.addShape('star5', {
      x: o.x + i * pitch, y: o.y, w: size, h: size,
      fill: o.outline ? { type: 'none' } : { color: o.color },
      line: o.outline ? { color: o.color, width: 0.75 } : undefined,
    });
  }
}

/** Line-art takeaway cup (medallion on slide 5, first row of slide 8). */
function cupIcon(s, o) {
  const { x, y, w, h, color } = o;
  const pen = { color, width: 1.25 };
  s.addShape('trapezoid', {
    x: x + w * 0.22, y: y + h * 0.28, w: w * 0.56, h: h * 0.58,
    rotate: 180, fill: { type: 'none' }, line: pen,
  });
  rect(s, { x: x + w * 0.16, y: y + h * 0.16, w: w * 0.68, h: h * 0.12, fill: { type: 'none' }, line: pen });
  rect(s, { x: x + w * 0.27, y: y + h * 0.06, w: w * 0.46, h: h * 0.1, fill: { type: 'none' }, line: pen });
}

// ---------------------------------------------------------------------------
// Freeform outlines lifted from the template's layouts (inches, shape-local)
// ---------------------------------------------------------------------------
const SHAPE_RIGHT_BLOB = [ // slide 3 — brown organic panel hugging the right edge
  { x: 1.294, y: 0.0, moveTo: true },
  { x: 3.804, y: 0.0 },
  { x: 3.804, y: 7.5 },
  { x: 1.338, y: 7.5 },
  { x: 1.193, y: 7.316 },
  { x: 0.0, y: 3.722, curve: { type: 'cubic', x1: 0.444, y1: 6.314, x2: 0.0, y2: 5.07 } },
  { x: 1.193, y: 0.129, curve: { type: 'cubic', x1: 0.0, y1: 2.375, x2: 0.444, y2: 1.131 } },
  { close: true },
];

const SHAPE_SWOOSH = [ // slide 8 — charcoal panel with a sweeping bottom edge
  { x: 0.0, y: 0.0, moveTo: true },
  { x: 13.333, y: 0.0 },
  { x: 13.333, y: 4.15 },
  { x: 13.189, y: 4.219 },
  { x: 1.464, y: 6.778, curve: { type: 'cubic', x1: 9.613, y1: 5.863, x2: 5.644, y2: 6.778 } },
  { x: 0.729, y: 6.768, curve: { type: 'cubic', x1: 1.218, y1: 6.778, x2: 0.973, y2: 6.775 } },
  { x: 0.0, y: 6.74 },
  { x: 0.0, y: 1.646 },
  { close: true },
];

const SHAPE_SWOOSH_BROWN = [ // slide 8 — brown edge peeking out beneath the swoosh
  { x: 0.0, y: 0.0, moveTo: true },
  { x: 12.754, y: 0.0 },
  { x: 12.754, y: 4.412 },
  { x: 12.514, y: 4.519 },
  { x: 1.464, y: 6.778, curve: { type: 'cubic', x1: 9.116, y1: 5.974, x2: 5.382, y2: 6.778 } },
  { x: 0.729, y: 6.768, curve: { type: 'cubic', x1: 1.218, y1: 6.778, x2: 0.973, y2: 6.775 } },
  { x: 0.0, y: 6.74 },
  { x: 0.0, y: 1.646 },
  { close: true },
];

const SHAPE_DOME = [ // slide 14 — full-width banner with a convex bottom edge
  { x: 0.0, y: 0.0, moveTo: true },
  { x: 13.333, y: 0.0 },
  { x: 13.333, y: 2.784 },
  { x: 12.932, y: 2.891 },
  { x: 6.667, y: 3.667, curve: { type: 'cubic', x1: 10.978, y1: 3.393, x2: 8.868, y2: 3.667 } },
  { x: 0.402, y: 2.891, curve: { type: 'cubic', x1: 4.465, y1: 3.667, x2: 2.355, y2: 3.393 } },
  { x: 0.0, y: 2.784 },
  { close: true },
];

const SHAPE_DOME_WIDE = [ // slide 10 — same idea, slightly flatter and deeper at the sides
  { x: 0.0, y: 0.0, moveTo: true },
  { x: 13.333, y: 0.0 },
  { x: 13.333, y: 3.05 },
  { x: 13.104, y: 3.102 },
  { x: 7.183, y: 3.667, curve: { type: 'cubic', x1: 11.36, y1: 3.461, x2: 9.338, y2: 3.667 } },
  { x: 6.667, y: 3.66 },
  { x: 6.168, y: 3.667 },
  { x: 5.236, y: 3.654 },
  { x: 0.23, y: 3.102, curve: { type: 'cubic', x1: 3.423, y1: 3.605, x2: 1.724, y2: 3.41 } },
  { x: 0.0, y: 3.05 },
  { close: true },
];

const SHAPE_QUOTE_CARD = [ // slide 11 — rounded card with a semicircular tab on top
  { x: 1.918, y: 0.0, moveTo: true },
  { x: 2.318, y: 0.4, curve: { type: 'cubic', x1: 2.139, y1: 0.0, x2: 2.318, y2: 0.179 } },
  { x: 2.317, y: 0.412 },
  { x: 3.757, y: 0.412 },
  { x: 3.803, y: 0.457, curve: { type: 'cubic', x1: 3.783, y1: 0.412, x2: 3.803, y2: 0.432 } },
  { x: 3.803, y: 2.462 },
  { x: 3.757, y: 2.508, curve: { type: 'cubic', x1: 3.803, y1: 2.487, x2: 3.783, y2: 2.508 } },
  { x: 0.046, y: 2.508 },
  { x: 0.0, y: 2.462, curve: { type: 'cubic', x1: 0.02, y1: 2.508, x2: 0.0, y2: 2.487 } },
  { x: 0.0, y: 0.457 },
  { x: 0.046, y: 0.412, curve: { type: 'cubic', x1: 0.0, y1: 0.432, x2: 0.02, y2: 0.412 } },
  { x: 1.519, y: 0.412 },
  { x: 1.518, y: 0.4 },
  { x: 1.918, y: 0.0, curve: { type: 'cubic', x1: 1.518, y1: 0.179, x2: 1.697, y2: 0.0 } },
  { close: true },
];

// ---------------------------------------------------------------------------
// Slide builders
// ---------------------------------------------------------------------------

// 1 — Cover: dark gradient with the Machiato wordmark.
function slide01(pptx) {
  const s = pptx.addSlide();
  verticalGradient(s, '8C8C8C', '595959', 30);
  coffeeBeans(s, { x: 9.29, y: 2.767, w: 1.0, h: 1.0, color: C.white, rotate: 9.5 });
  s.addText('Machiato', {
    x: 3.543, y: 2.99, w: 6.247, h: 1.582, align: 'center', valign: 'top',
    fontFace: F.logo, fontSize: 88, charSpacing: -1, color: C.white,
  });
  s.addText('COFFEE SHOP BUSINESS PRESENTATION TEMPLATE', {
    x: 3.655, y: 4.364, w: 6.055, h: 0.303, align: 'center', valign: 'top',
    fontFace: F.logo, fontSize: 12, charSpacing: 2, color: C.white,
  });
}

// 2 — About Us: 65%-black panel over the left half of the slide.
function slide02(pptx) {
  const s = pptx.addSlide();
  rect(s, { x: 0, y: 0, w: 6.846, h: 7.5, fill: { color: '000000', transparency: 35 } });
  eyebrow(s, 'About Us', 0.739, 1.683, C.brown);
  heading(s, 'Welcome to Machiato Coffee', { x: 0.725, y: 1.967, w: 5.001, h: 1.313, color: C.white });
  accentBar(s, 0.837, 3.509);
  body(s, LOREM_LONG + 'Nisl condimentum id venenatis a. Feugiat sed lectus vestibulum mattis ullamcorper velit sed. ',
    { x: 0.739, y: 3.779, w: 5.212, h: 1.453, color: C.white });
  button(s, 'EXPLORE', 0.826, 5.426, 1.147);
}

// 3 — About Us: light slide, brown organic panel right, bean watermark top-left.
function slide03(pptx) {
  const s = pptx.addSlide();
  freeform(s, SHAPE_RIGHT_BLOB, { x: 9.53, y: 0, w: 3.804, h: 7.5, fill: { color: C.brownDark } });
  coffeeBeans(s, { x: 0.0, y: 0.0, w: 2.747, h: 3.0, color: C.bodyLight, transparency: 30 });
  coffeeBeans(s, { x: 1.562, y: 0.0, w: 1.853, h: 1.403, color: C.bodyLight, transparency: 30 });
  eyebrow(s, 'About Us', 1.672, 1.654);
  heading(s, 'Enjoy the taste \nof our coffee', { x: 1.626, y: 1.97, w: 4.813, h: 1.313 });
  body(s, DUIS, { x: 1.655, y: 3.455, w: 4.783, h: 0.903 });
  body(s, LOREM_MIX + 'PLACEHOLDER',
    { x: 1.655, y: 4.373, w: 4.783, h: 0.897 });
  button(s, 'ORDER NOW', 1.72, 5.456, 1.448, C.brownDark);
}

// 4 — Hero band, product card on the right, white info bar below.
function slide04(pptx) {
  const s = pptx.addSlide();
  rect(s, { x: 0, y: 0, w: 13.333, h: 4.937, fill: { color: '000000', transparency: 75 } });
  roundRect(s, { x: 1.413, y: 4.502, w: 10.508, h: 2.095, rectRadius: 0.19, fill: { color: C.white }, shadow: SOFT() });
  roundRect(s, { x: 8.873, y: 0.802, w: 3.603, h: 4.135, rectRadius: 0.12, fill: { color: C.white }, shadow: SOFT() });

  heading(s, 'There\u2019s always time for good coffee', { x: 1.654, y: 1.144, w: 5.396, h: 1.313, color: C.white });
  body(s, DUIS + LOREM_MIX, { x: 1.683, y: 2.63, w: 5.396, h: 1.175, color: C.white });

  label(s, 'SPECIAL BLEND', { x: 2.072, y: 4.97, w: 1.867, h: 0.303, color: C.ink });
  body(s, 'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum magna enimvas minimasa consectetur.',
    { x: 2.082, y: 5.234, w: 2.726, h: 0.904 });
  label(s, 'AWESOME TASTE', { x: 5.468, y: 4.968, w: 1.867, h: 0.303, color: C.ink });
  body(s, DUIS + 'lorem ipsum dolor siabes dolore magna aliqua.', { x: 5.478, y: 5.232, w: 5.773, h: 0.897 });

  // product-card footer: name, price, rating pill, calorie note
  s.addText('Machiato', {
    x: 9.056, y: 3.901, w: 1.883, h: 0.505, fontFace: F.logo, fontSize: 24, color: C.ink, valign: 'top',
  });
  s.addText('$15', {
    x: 9.07, y: 4.346, w: 0.616, h: 0.404, fontFace: F.logo, fontSize: 18, color: C.brownRust, valign: 'top',
  });
  roundRect(s, { x: 11.384, y: 3.975, w: 0.832, h: 0.331, rectRadius: 0.165, fill: { color: C.body } });
  s.addShape('star5', { x: 11.534, y: 4.039, w: 0.174, h: 0.175, fill: { color: C.star } });
  s.addText('4.6', {
    x: 11.684, y: 4.009, w: 0.474, h: 0.269,
    fontFace: 'Poppins SemiBold', fontSize: 10, charSpacing: 1, color: C.white, valign: 'top',
  });
  s.addText('Low Calories', {
    x: 10.842, y: 4.351, w: 1.457, h: 0.347, align: 'right',
    fontFace: F.sans, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top',
  });
}

// 5 — Badge medallion over the left photo column, copy on the right.
function slide05(pptx) {
  const s = pptx.addSlide();
  coffeeBeans(s, { x: 10.179, y: -0.244, w: 3.177, h: 2.493, color: C.bodyLight, transparency: 30, rotate: 9.5 });

  oval(s, { x: 4.68, y: 0.629, w: 1.406, h: 1.406, fill: { color: C.white }, shadow: SOFT() });
  oval(s, { x: 4.803, y: 0.758, w: 1.159, h: 1.159, fill: { type: 'none' }, line: { color: C.brownRust, width: 1 } });
  cupIcon(s, { x: 5.073, y: 0.911, w: 0.6, h: 0.6, color: C.brown });
  stars(s, { x: 5.036, y: 1.541, count: 4, color: C.brown });

  heading(s, 'Great ideas comes from a cup of coffee', { x: 6.795, y: 1.483, w: 5.429, h: 1.313 });
  accentBar(s, 6.925, 3.033);
  label(s, 'BEST QUALITY OF COFFEE', { x: 6.841, y: 3.441, w: 2.653, h: 0.303, charSpacing: 1.2 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ' +
    'ut labore et dolore magna aliqua. Suspendisse potenti nullam ac tortor vitae purus faucibus ornare. ' +
    'Tellus elementum sagittis vitae et leo duis ut diam quam nisl condimentum id venenatis. ',
    { x: 6.841, y: 3.699, w: 5.429, h: 1.175 });
  label(s, 'SPECIAL BLEND', { x: 6.826, y: 5.093, w: 1.867, h: 0.303 });
  body(s, 'Duis irure dolor reprehenderit voluptate velit cillum magna', { x: 6.837, y: 5.357, w: 2.371, h: 0.626 });
  label(s, 'AWESOME TASTE', { x: 9.912, y: 5.09, w: 1.902, h: 0.303 });
  body(s, 'Duis irure dolor reprehenderit voluptate velit cillum magna', { x: 9.899, y: 5.353, w: 2.371, h: 0.626 });
}

// 6 — "Our best menu": charcoal slide with four white product cards.
function slide06(pptx) {
  const s = pptx.addSlide();
  rect(s, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: C.charcoal } });
  coffeeBeans(s, { x: 10.873, y: -0.24, w: 3.149, h: 1.969, color: C.white, transparency: 96, rotate: -9.5 });

  eyebrow(s, 'Our Services', 1.083, 0.852);
  heading(s, 'Our best menu', { x: 1.052, y: 1.082, w: 4.213, h: 0.707, color: C.white });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore aliqua',
    { x: 7.922, y: 1.095, w: 4.302, h: 0.62, color: C.bodyLight });

  const menu = [
    { x: 0.758, name: 'Machiato', price: '$12' },
    { x: 3.818, name: 'Cappuccino', price: '$15' },
    { x: 6.879, name: 'Latte', price: '$13' },
    { x: 9.955, name: 'Espresso', price: '$12' },
  ];
  menu.forEach((m) => {
    roundRect(s, { x: m.x, y: 2.187, w: 2.7, h: 4.314, rectRadius: 0.2, fill: { color: C.white }, shadow: SOFT() });
    s.addText(m.name, {
      x: m.x + 0.195, y: 4.886, w: 2.2, h: 0.438,
      fontFace: F.head, fontSize: 20, color: C.ink, valign: 'top',
    });
    body(s, CARD_BLURB, { x: m.x + 0.2, y: 5.323, w: 2.293, h: 0.62 });
    stars(s, { x: m.x + 0.294, y: 6.083, count: 5, color: C.brown });
    s.addText(m.price, {
      x: m.x + 1.978, y: 6.002, w: 0.52, h: 0.337,
      fontFace: F.labelMed, fontSize: 14, color: C.brownRust, valign: 'top',
    });
  });
}

// 7 — "High quality of ingredients": copy left, three ringed thumbnails right.
function slide07(pptx) {
  const s = pptx.addSlide();
  eyebrow(s, 'Our Services', 1.432, 2.431, C.brown);
  heading(s, 'High quality of ingredients', { x: 1.401, y: 2.744, w: 4.666, h: 1.313, bold: true });
  label(s, 'BEST QUALITY OF COFFEE', { x: 1.448, y: 4.267, w: 2.653, h: 0.303, charSpacing: 1.2 });
  body(s, LOREM_LONG + LOREM_TAIL, { x: 1.464, y: 4.543, w: 5.393, h: 1.453 });

  const items = [
    { y: 2.056, title: 'HIGH QUALITY BLEND' },
    { y: 3.703, title: 'TOP QUALITY BEANS' },
    { y: 5.346, title: 'LATTEE ART' },
  ];
  items.forEach((it) => {
    oval(s, { x: 7.979, y: it.y, w: 1.3, h: 1.3, fill: { type: 'none' }, line: { color: C.brownDark, width: 1.5 } });
    imgPrompt(s, { x: 8.054, y: it.y + 0.075, w: 1.15, h: 1.15, shape: 'ellipse' });
    oval(s, { x: 8.926, y: it.y, w: 0.354, h: 0.354, fill: { color: C.brownDark } });
    label(s, it.title, { x: 9.602, y: it.y + 0.083, w: 2.46, h: 0.303 });
    body(s, VELITESSE, { x: 9.602, y: it.y + 0.321, w: 2.843, h: 0.897 });
  });
}

// 8 — Charcoal swoosh panel, three icon rows, photo card right.
function slide08(pptx) {
  const s = pptx.addSlide();
  freeform(s, SHAPE_SWOOSH_BROWN, { x: 0.579, y: 0, w: 12.754, h: 6.778, fill: { color: C.brownDark } });
  freeform(s, SHAPE_SWOOSH, { x: 0, y: 0, w: 13.333, h: 6.778, fill: { color: C.charcoal } });
  coffeeBeans(s, { x: -0.189, y: -0.257, w: 2.879, h: 3.022, color: C.bodyLight, transparency: 95, rotate: 9.5 });
  coffeeBeans(s, { x: 2.225, y: -0.27, w: 2.31, h: 2.839, color: C.bodyLight, transparency: 95, rotate: -9.5 });

  eyebrow(s, 'Our Services', 1.565, 0.98, C.brown);
  heading(s, 'Serve you with best quality coffee', { x: 1.541, y: 1.236, w: 5.028, h: 1.313, color: C.white });
  accentBar(s, 1.64, 2.804, C.brownDark);

  const rows = [
    { y: 3.268, title: 'SPECIAL BLEND', icon: 'cup' },
    { y: 4.289, title: 'BEST INGRIDIENTS', icon: 'beans' },
    { y: 5.309, title: 'AWESOME TASTE', icon: 'stars' },
  ];
  rows.forEach((r) => {
    if (r.icon === 'cup') cupIcon(s, { x: 1.586, y: r.y, w: 0.7, h: 0.7, color: C.brownDark });
    if (r.icon === 'beans') coffeeBeans(s, { x: 1.586, y: r.y, w: 0.7, h: 0.7, color: C.brownDark });
    if (r.icon === 'stars') stars(s, { x: 1.6, y: r.y + 0.26, count: 3, size: 0.19, pitch: 0.185, color: C.brownDark, outline: true });
    label(s, r.title, { x: 2.647, y: r.y - 0.084, w: 2.46, h: 0.303, color: C.bodyLight });
    body(s, CONSECTETUR, { x: 2.647, y: r.y + 0.154, w: 3.376, h: 0.62, color: C.bodyLight2 });
  });
}

// 9 — "Our Team": three staggered white cards left, copy right.
function slide09(pptx) {
  const s = pptx.addSlide();
  coffeeBeans(s, { x: 10.272, y: -0.234, w: 3.106, h: 2.546, color: C.bodyLight, transparency: 30, rotate: 9.5 });
  coffeeBeans(s, { x: 12.123, y: 0.978, w: 1.766, h: 1.882, color: C.bodyLight, transparency: 30, rotate: 57.9 });

  const team = [
    { cardX: 1.481, cardY: 1.055, textX: 3.144, name: 'Brian Forster', role: 'BARISTA' },
    { cardX: 2.156, cardY: 2.949, textX: 3.83, name: 'Abigail Madison', role: 'OWNER' },
    { cardX: 0.965, cardY: 4.842, textX: 2.633, name: 'Aubrey Grace', role: 'MANAGER' },
  ];
  team.forEach((t) => {
    roundRect(s, { x: t.cardX, y: t.cardY, w: 4.356, h: 1.603, rectRadius: 0.14, fill: { color: C.white }, shadow: SOFT() });
    s.addText(t.name, {
      x: t.textX, y: t.cardY + 0.2, w: 2.088, h: 0.37,
      fontFace: F.name, fontSize: 16, color: C.inkSoft, valign: 'top',
    });
    s.addText(t.role, {
      x: t.textX + 0.012, y: t.cardY + 0.495, w: 1.906, h: 0.252,
      fontFace: F.labelMed, fontSize: 9, charSpacing: 1, color: C.brown, valign: 'top',
    });
    line(s, { x: t.textX + 0.106, y: t.cardY + 0.806, w: 2.044, h: 0, line: { color: C.rule, width: 0.75 } });
    s.addText(VOLUPTATE, {
      x: t.textX + 0.021, y: t.cardY + 0.827, w: 2.335, h: 0.573,
      fontFace: F.body, fontSize: 10, color: C.body, lineSpacingMultiple: 1.5, valign: 'top',
    });
  });

  eyebrow(s, 'Our Team', 7.592, 1.785);
  heading(s, 'We serve your with best effort', { x: 7.576, y: 2.123, w: 4.666, h: 1.313, bold: true });
  label(s, 'GIVE YOU THE BEST QUALITY OF COFFEE', { x: 7.623, y: 3.708, w: 4.045, h: 0.303, charSpacing: 1.2 });
  body(s, LOREM_LONG + LOREM_TAIL, { x: 7.639, y: 3.984, w: 4.603, h: 1.731 });
}

// 10 — "Meet our team": brown dome banner, three portraits with social badges.
function slide10(pptx) {
  const s = pptx.addSlide();
  coffeeBeans(s, { x: -0.257, y: -0.471, w: 13.839, h: 8.333, color: C.bodyLight, transparency: 50, rotate: -4.1 });
  freeform(s, SHAPE_DOME_WIDE, { x: 0, y: 0, w: 13.333, h: 3.667, fill: { color: C.brown } });

  s.addText('Meet our team', {
    x: 3.768, y: 0.86, w: 5.802, h: 0.707, align: 'center',
    fontFace: F.head, fontSize: 36, color: C.white, valign: 'top',
  });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ' +
    'ut labore et dolore magna aliqua Suspendisse potenti nullam  tortor vitae purus faucibus ornare.',
    { x: 3.378, y: 1.651, w: 6.842, h: 0.62, color: C.bodyLight, align: 'center' });

  const people = [
    { badge: 2.778, textX: 2.078, name: 'Brian Forster', role: 'BARISTA' },
    { badge: 6.415, textX: 5.714, name: 'Abigail Madison', role: 'OWNER' },
    { badge: 10.052, textX: 9.35, name: 'Daniel Smith', role: 'BARISTA' },
  ];
  people.forEach((p) => {
    oval(s, { x: p.badge, y: 5.42, w: 0.504, h: 0.504, fill: { color: C.white }, shadow: LIFT() });
    oval(s, { x: p.badge + 0.061, y: 5.479, w: 0.388, h: 0.388, fill: { type: 'none' }, line: { color: C.brown, width: 0.5 } });
    roundRect(s, { // Facebook mark
      x: p.badge + 0.168, y: 5.601, w: 0.167, h: 0.167, rectRadius: 0.03, fill: { color: C.brown },
    });
    s.addText('f', {
      x: p.badge + 0.168, y: 5.591, w: 0.167, h: 0.167, align: 'center', valign: 'middle',
      fontFace: F.label, fontSize: 9, bold: true, color: C.white,
    });
    s.addText(p.name, {
      x: p.textX, y: 5.996, w: 1.906, h: 0.337, align: 'center',
      fontFace: F.name, fontSize: 14, color: C.inkSoft, valign: 'top',
    });
    s.addText(p.role, {
      x: p.textX, y: 6.256, w: 1.906, h: 0.252, align: 'center',
      fontFace: F.labelMed, fontSize: 9, charSpacing: 1, color: C.brown, valign: 'top',
    });
  });
}

// 11 — "What clients say": three quote cards with avatar tabs and star ratings.
function slide11(pptx) {
  const s = pptx.addSlide();
  rect(s, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: C.cream } });
  coffeeBeans(s, { x: 9.732, y: -0.28, w: 3.66, h: 2.289, color: C.brown, transparency: 55, rotate: 9.5 });
  coffeeBeans(s, { x: 0.098, y: -0.432, w: 3.66, h: 2.244, color: C.brown, transparency: 55, rotate: 12.8 });

  s.addText('What clients say', {
    x: 3.956, y: 1.155, w: 5.421, h: 0.707, align: 'center',
    fontFace: F.head, fontSize: 36, bold: true, color: C.ink, valign: 'top',
  });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ' +
    'ut labore et dolore magna aliqua. Suspendisse potenti nullam ac tortor vitae purus faucibus ornare. ',
    { x: 3.271, y: 1.985, w: 6.79, h: 0.62, align: 'center' });

  const quotes = [
    { x: 0.724, avatarX: 2.292, name: 'Abigail Nester' },
    { x: 4.805, avatarX: 6.373, name: 'Alfred Julian' },
    { x: 8.887, avatarX: 10.454, name: 'John Erwin' },
  ];
  quotes.forEach((q) => {
    freeform(s, SHAPE_QUOTE_CARD, { x: q.x, y: 3.312, w: 3.803, h: 2.508, fill: { color: C.white }, shadow: LIFT() });
    imgPrompt(s, { x: q.avatarX, y: 3.36, w: 0.7, h: 0.7, glyphScale: 0.9 });
    body(s, TESTIMONIAL, { x: q.x + 0.278, y: 4.083, w: 3.27, h: 1.175, align: 'left' });
    line(s, { x: q.x + 0.381, y: 5.479, w: 0.25, h: 0, line: { color: C.inkSoft, width: 1.5 } });
    s.addText(q.name, {
      x: q.x + 0.682, y: 5.32, w: 1.7, h: 0.337,
      fontFace: F.label, fontSize: 14, bold: true, color: C.inkSoft, valign: 'top',
    });
    stars(s, { x: q.x + 2.737, y: 5.401, count: 4, color: C.star });
  });
}

// 12 — "Our Charts": white text card left, native clustered bar chart right.
function slide12(pptx) {
  const s = pptx.addSlide();
  rect(s, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: C.cream } });
  coffeeBeans(s, { x: 10.699, y: -0.24, w: 3.149, h: 1.969, color: C.brown, transparency: 55, rotate: -9.5 });
  roundRect(s, { x: 0.962, y: 1.022, w: 4.867, h: 5.456, rectRadius: 0.1, fill: { color: C.white }, shadow: SOFT() });

  heading(s, 'Our Charts', { x: 1.327, y: 1.483, w: 4.175, h: 0.707 });
  accentBar(s, 1.472, 2.33);
  label(s, 'OUR PROGRESS IN 2020', { x: 1.374, y: 2.721, w: 2.477, h: 0.303, charSpacing: 1.2 });
  body(s, 'PLACEHOLDER' +
    'siabes dolore magna aliqua enimvas minimasa veniam nostrud adipiscinc.',
    { x: 1.367, y: 3.024, w: 4.115, h: 0.897 });

  const bullets = [
    'Consectetur adipiscing siabes doloremag faubicus  eiusmod tempor incididunt ut labore.',
    'voluptate velitesse clum ipsum dolor sit anim lorem ipsum dolor nostrud minimasa veniam nostrud.',
    'Duis aute irure dolor reprehenderit voluptate velit esse cillum ipsum dolor lorem.',
  ];
  bullets.forEach((t, i) => {
    const y = 4.046 + i * 0.608;
    oval(s, { x: 1.552, y: y + 0.164, w: 0.1, h: 0.1, fill: { color: C.brown } });
    body(s, t, { x: 1.832, y, w: 3.65, h: 0.62 });
  });

  const labels = ['Graphic 1', 'Graphic 2', 'Graphic 3', 'Graphic 4'];
  s.addChart(pptx.ChartType.bar, [
    { name: 'Series 1', labels, values: [2, 2.5, 3.5, 4] },
    { name: 'Series 2', labels, values: [2.4, 3.4, 1.8, 3] },
    { name: 'Series 3', labels, values: [3, 2, 3, 5] },
  ], {
    x: 6.42, y: 1.483, w: 6.051, h: 4.534,
    barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 208, barOverlapPct: -27,
    chartColors: [C.brown, 'D39C77', '824E2A'],
    chartArea: { fill: { color: C.white, transparency: 100 } },
    plotArea: { fill: { color: C.white, transparency: 100 } },
    showLegend: true, legendPos: 'b', legendFontFace: 'Arial', legendFontSize: 12, legendColor: C.body,
    catAxisLabelFontFace: 'Arial', catAxisLabelFontSize: 12, catAxisLabelColor: C.body,
    catAxisLineColor: C.body, catAxisMajorTickMark: 'none', catAxisMinorTickMark: 'none',
    valAxisLabelFontFace: 'Arial', valAxisLabelFontSize: 12, valAxisLabelColor: C.body,
    valAxisLineShow: false, valAxisMajorTickMark: 'none', valAxisMinorTickMark: 'none',
    valGridLine: { style: 'solid', color: 'BFBFBF', size: 0.75 },
    catGridLine: { style: 'none' },
    showValue: false,
  });
}

// 13 — Phone mock-up centred on a brown disc, six feature callouts around it.
function slide13(pptx) {
  const s = pptx.addSlide();
  oval(s, { x: 4.705, y: 1.715, w: 3.924, h: 3.924, fill: { color: C.brownRust } });
  phonePlaceholder(s, { x: 5.173, y: 0.667, w: 2.987, h: 6.02 });

  const features = [
    { textX: 1.322, textY: 1.488, align: 'right', iconX: 3.878, iconY: 1.572, icon: 'cup' },
    { textX: 0.943, textY: 3.234, align: 'right', iconX: 3.566, iconY: 3.327, icon: 'stars' },
    { textX: 1.322, textY: 5.023, align: 'right', iconX: 3.878, iconY: 5.082, icon: 'measure' },
    { textX: 9.548, textY: 1.488, align: 'left', iconX: 8.77, iconY: 1.572, icon: 'frappe' },
    { textX: 9.939, textY: 3.234, align: 'left', iconX: 9.115, iconY: 3.327, icon: 'beans' },
    { textX: 9.548, textY: 5.023, align: 'left', iconX: 8.77, iconY: 5.082, icon: 'mug' },
  ];
  features.forEach((f) => {
    label(s, 'TEXT TITTLE HERE', { x: f.textX, y: f.textY, w: 2.46, h: 0.303, align: f.align });
    body(s, VOLUPTATE, { x: f.textX, y: f.textY + 0.238, w: 2.46, h: 0.62, align: f.align });
    drinkIcon(s, f.icon, { x: f.iconX, y: f.iconY, w: 0.7, h: 0.7, color: C.brownRust });
  });
}

/** Line-art drink icons used on slides 8 and 13. */
function drinkIcon(s, kind, o) {
  const { x, y, w, h, color } = o;
  const pen = { color, width: 1.25 };
  if (kind === 'cup') return cupIcon(s, o);
  if (kind === 'beans') return coffeeBeans(s, { x, y, w, h, color });
  if (kind === 'stars') return stars(s, { x: x + 0.02, y: y + h * 0.34, count: 3, size: w * 0.3, pitch: w * 0.28, color, outline: true });
  if (kind === 'frappe') { // tall cup, domed lid and a straw
    s.addShape('trapezoid', { x: x + w * 0.26, y: y + h * 0.3, w: w * 0.48, h: h * 0.58, rotate: 180, fill: { type: 'none' }, line: pen });
    oval(s, { x: x + w * 0.22, y: y + h * 0.16, w: w * 0.56, h: h * 0.16, fill: { type: 'none' }, line: pen });
    line(s, { x: x + w * 0.5, y: y + h * 0.22, w: w * 0.16, h: -h * 0.18, line: pen });
  }
  if (kind === 'measure') { // measuring jug with handle and gauge marks
    rect(s, { x: x + w * 0.2, y: y + h * 0.3, w: w * 0.5, h: h * 0.46, fill: { type: 'none' }, line: pen });
    s.addShape('rightBracket', { x: x + w * 0.68, y: y + h * 0.4, w: w * 0.16, h: h * 0.24, fill: { type: 'none' }, line: pen });
    line(s, { x: x + w * 0.28, y: y + h * 0.48, w: w * 0.2, h: 0, line: pen });
    line(s, { x: x + w * 0.28, y: y + h * 0.6, w: w * 0.2, h: 0, line: pen });
  }
  if (kind === 'mug') { // mug with three steam strokes
    rect(s, { x: x + w * 0.2, y: y + h * 0.42, w: w * 0.44, h: h * 0.36, fill: { type: 'none' }, line: pen });
    s.addShape('rightBracket', { x: x + w * 0.62, y: y + h * 0.5, w: w * 0.14, h: h * 0.2, fill: { type: 'none' }, line: pen });
    [0.28, 0.4, 0.52].forEach((fx) => line(s, { x: x + w * fx, y: y + h * 0.34, w: 0, h: -h * 0.14, line: pen }));
  }
}

// 14 — "Get in touch": dome banner plus four bordered contact cards.
function slide14(pptx) {
  const s = pptx.addSlide();
  freeform(s, SHAPE_DOME, { x: 0, y: 0, w: 13.333, h: 3.667, fill: { color: '000000', transparency: 65 } });

  s.addText('Get in touch', {
    x: 4.143, y: 1.063, w: 5.047, h: 1.01, align: 'center',
    fontFace: F.head, fontSize: 54, color: C.white, valign: 'top',
  });
  s.addText('Consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore magna aliqua ' +
    'suspendisse potenti nullam vitae faucibus. ', {
    x: 3.762, y: 2.073, w: 5.81, h: 0.62, align: 'center',
    fontFace: F.sans, fontSize: 11, color: C.white, lineSpacingMultiple: 1.5, valign: 'top',
  });

  const contacts = [
    { x: 1.175, icon: 'marker', title: 'OUR ADDRESS', lines: ['Street Name, 45, Building 85, 445566, Soho, NY'] },
    { x: 4.014, icon: 'laptop', title: 'OFFICE HOURS', lines: ['Monday \u2013 Friday', '08:00 \u2013 17:00'] },
    { x: 6.861, icon: 'envelope', title: 'EMAIL', lines: ['email@machiato.com', 'service@machiato.com'] },
    { x: 9.715, icon: 'phone', title: 'PHONE', lines: ['(123) 456 7890', '(800) 123 4567'] },
  ];
  contacts.forEach((c) => {
    roundRect(s, {
      x: c.x, y: 4.356, w: 2.442, h: 2.302, rectRadius: 0.16,
      fill: { color: C.white }, line: { color: C.brownDark, width: 0.75 }, shadow: LIFT(),
    });
    oval(s, { x: c.x + 0.929, y: 4.694, w: 0.618, h: 0.618, fill: { color: C.brownDark, transparency: 69 } });
    contactIcon(s, c.icon, { x: c.x + 1.037, y: 4.796, w: 0.4, h: 0.4, color: C.brownDark });
    s.addText(c.title, {
      x: c.x, y: 5.454, w: 2.442, h: 0.337, align: 'center',
      fontFace: F.label, fontSize: 14, bold: true, charSpacing: 1, color: C.ink, valign: 'top',
    });
    s.addText(c.lines.join('\n'), {
      x: c.x + 0.24, y: 5.703, w: 1.965, h: 0.62, align: 'center',
      fontFace: F.body, fontSize: 11, color: C.body, lineSpacingMultiple: 1.5, valign: 'top',
    });
  });
}

/** Solid glyphs inside the contact-card circles. */
function contactIcon(s, kind, o) {
  const { x, y, w, h, color } = o;
  const fill = { color };
  if (kind === 'marker') {
    oval(s, { x: x + w * 0.18, y: y + h * 0.04, w: w * 0.64, h: h * 0.64, fill });
    s.addShape('triangle', { x: x + w * 0.3, y: y + h * 0.48, w: w * 0.4, h: h * 0.46, rotate: 180, fill });
    oval(s, { x: x + w * 0.38, y: y + h * 0.24, w: w * 0.24, h: h * 0.24, fill: { color: C.white } });
  }
  if (kind === 'laptop') {
    rect(s, { x: x + w * 0.16, y: y + h * 0.2, w: w * 0.68, h: h * 0.46, fill });
    rect(s, { x: x + w * 0.26, y: y + h * 0.3, w: w * 0.48, h: h * 0.26, fill: { color: C.white } });
    rect(s, { x: x + w * 0.04, y: y + h * 0.68, w: w * 0.92, h: h * 0.12, fill });
  }
  if (kind === 'envelope') {
    rect(s, { x: x + w * 0.08, y: y + h * 0.24, w: w * 0.84, h: h * 0.52, fill });
    s.addShape('triangle', { x: x + w * 0.08, y: y + h * 0.24, w: w * 0.84, h: h * 0.32, rotate: 180, fill: { color: C.white } });
  }
  if (kind === 'phone') { // handset: angled bar with an earpiece at each end
    roundRect(s, { x: x + w * 0.18, y: y + h * 0.42, w: w * 0.64, h: h * 0.16, rectRadius: 0.03, rotate: -40, fill });
    oval(s, { x: x + w * 0.1, y: y + h * 0.56, w: w * 0.3, h: h * 0.3, fill });
    oval(s, { x: x + w * 0.6, y: y + h * 0.14, w: w * 0.3, h: h * 0.3, fill });
  }
}

// 15 — Closing: same dark gradient cover, "Thank You" set right of centre.
function slide15(pptx) {
  const s = pptx.addSlide();
  verticalGradient(s, '8C8C8C', '595959', 30);
  s.addText('Thank You', {
    x: 7.15, y: 2.946, w: 4.765, h: 1.111,
    fontFace: F.logo, fontSize: 60, charSpacing: -1, color: C.white, valign: 'top',
  });
  line(s, { x: 7.51, y: 4.148, w: 0, h: 0.4, line: { color: C.white, width: 1 } });
  s.addText('Never say goodbye, because saying goodbye means going away and going away means forgetting.', {
    x: 7.552, y: 3.987, w: 4.363, h: 0.625, align: 'justify',
    fontFace: F.sans, fontSize: 11, color: C.white, lineSpacingMultiple: 1.5, valign: 'top',
  });
}

// ---------------------------------------------------------------------------
// Build & save
// ---------------------------------------------------------------------------
function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'MACHIATO', width: 13.333, height: 7.5 });
  pptx.layout = 'MACHIATO';
  pptx.title = 'Machiato — Coffee Shop Business Presentation Template';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
   slide09, slide10, slide11, slide12, slide13, slide14, slide15]
    .forEach((fn) => fn(pptx));

  return pptx.writeFile({
    fileName: path.join(__dirname, '177b2bc6-2152-4015-80b1-2175684f47af_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
