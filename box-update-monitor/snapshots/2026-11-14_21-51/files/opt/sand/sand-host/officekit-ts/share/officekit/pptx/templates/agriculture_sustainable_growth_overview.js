/**
 * "Agriculture — Innovating Sustainable Growth For The Future"
 * A 10-slide 16:9 (13.333" x 7.5") deck rebuilt with pptxgenjs.
 *
 * Photos/icons of the original deck are re-created as programmatic placeholders:
 *   - picturePlaceholder() -> the empty picture frames (5% pattern fill in the original)
 *   - iconPlaceholder()    -> the small white raster/vector glyphs sitting on orange chips
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Design tokens
 * ------------------------------------------------------------------ */
const SLIDE_W = 13.333333;
const SLIDE_H = 7.5;

const C = {
  teal: '1E504D', // primary brand green
  orange: 'EB6C37', // accent
  ink: '000000', // default body/heading text
  white: 'FFFFFF',
  grey: '595959', // tx1 lumMod 65%  – body copy on light backgrounds
  white95: 'F2F2F2', // bg1 lumMod 95% – body copy on coloured cards
  white85: 'D9D9D9', // bg1 lumMod 85% – body copy on the teal slides
  frame: 'F6F8FC', // empty picture frame: average of the 5% accent pattern fill
};

const F = { head: 'Oswald', body: 'Open Sans' };

/* ------------------------------------------------------------------ *
 * Shape / text helpers
 * ------------------------------------------------------------------ */
function rect(slide, x, y, w, h, color) {
  slide.addShape('rect', { x, y, w, h, fill: { color }, line: { type: 'none' } });
}

/**
 * Empty picture frame of the original template. The reference deck leaves these
 * placeholders unfilled (a 5% pattern), so they are drawn as a barely-there tint
 * that works on both the light and the dark slides.
 */
function picturePlaceholder(slide, x, y, w, h) {
  slide.addShape('rect', {
    x, y, w, h,
    fill: { color: C.frame, transparency: 96 },
    line: { type: 'none' },
  });
}

/**
 * Stand-ins for the small white pictograms of the original deck (png/svg glyphs).
 * Every icon is a list of native shapes whose x/y/w/h are fractions of the chip
 * size. `solid` parts are filled white, the rest are drawn as white outlines.
 */
function bust(x, y, w) {
  // Head + shoulders — the "people" motif of the two icons on slide 4.
  return [
    { shape: 'ellipse', x, y, w, h: w, solid: true },
    { shape: 'chord', x: x - w * 0.4, y: y + w * 1.15, w: w * 1.8, h: w, solid: true },
  ];
}

function wheatEar() {
  // A diagonal stalk with four pairs of grains.
  const parts = [{ shape: 'rect', x: 0, y: 0.72, w: 0.62, h: 0.035, rotate: -42, solid: true }];
  for (let k = 0; k < 4; k += 1) {
    parts.push({ shape: 'ellipse', x: 0.30 - k * 0.135, y: 0.02 + k * 0.145, w: 0.30, h: 0.13, rotate: 325, solid: true });
    parts.push({ shape: 'ellipse', x: 0.48 - k * 0.135, y: 0.10 + k * 0.145, w: 0.30, h: 0.13, rotate: 35, solid: true });
  }
  return parts;
}

const ICONS = {
  // speech bubble + two people
  question: [
    { shape: 'wedgeRectCallout', x: 0.30, y: 0.00, w: 0.70, h: 0.44 },
    ...bust(0.02, 0.20, 0.18),
    ...bust(0.34, 0.50, 0.24),
  ],
  // light bulb above three people
  idea: [
    { shape: 'ellipse', x: 0.37, y: 0.00, w: 0.26, h: 0.26 },
    { shape: 'rect', x: 0.44, y: 0.26, w: 0.12, h: 0.07, solid: true },
    ...bust(0.03, 0.42, 0.19),
    ...bust(0.40, 0.42, 0.19),
    ...bust(0.77, 0.42, 0.19),
  ],
  // two leaves cupped by a hand
  sprout: [
    { shape: 'teardrop', x: 0.50, y: 0.02, w: 0.34, h: 0.34, rotate: 315, solid: true },
    { shape: 'teardrop', x: 0.22, y: 0.10, w: 0.30, h: 0.30, rotate: 225, solid: true },
    { shape: 'blockArc', x: -0.02, y: 0.34, w: 1.04, h: 0.95 },
  ],
  wheat: wheatEar(),
  // monitor + cloud with sync arrows
  cloudSync: [
    { shape: 'cloud', x: 0.30, y: 0.00, w: 0.68, h: 0.42 },
    { shape: 'rect', x: 0.02, y: 0.38, w: 0.52, h: 0.38 },
    { shape: 'rect', x: 0.22, y: 0.80, w: 0.12, h: 0.06, solid: true },
  ],
  envelope: [
    { shape: 'rect', x: 0.05, y: 0.24, w: 0.90, h: 0.52 },
    { shape: 'line', x: 0.05, y: 0.24, w: 0.45, h: 0.30 },
    { shape: 'line', x: 0.50, y: 0.24, w: 0.45, h: 0.30, flipV: true },
  ],
  globe: [
    { shape: 'ellipse', x: 0.05, y: 0.05, w: 0.90, h: 0.90 },
    { shape: 'ellipse', x: 0.33, y: 0.05, w: 0.34, h: 0.90 },
    { shape: 'line', x: 0.06, y: 0.50, w: 0.88, h: 0 },
  ],
};

function iconPlaceholder(slide, name, x, y, size) {
  ICONS[name].forEach((part) => {
    slide.addShape(part.shape, {
      x: x + part.x * size, y: y + part.y * size,
      w: part.w * size, h: part.h * size,
      rotate: part.rotate || 0,
      flipV: part.flipV || false,
      fill: part.solid ? { color: C.white } : { type: 'none' },
      line: part.solid ? { type: 'none' } : { color: C.white, width: 1.4 },
    });
  });
}

/** Big Oswald display type. `text` may contain \n for extra lines. */
function title(slide, o) {
  slide.addText(o.text, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: F.head, fontSize: o.size || 60, bold: true,
    color: o.color || C.ink, align: o.align,
    valign: 'top', wrap: false, fit: 'resize',
  });
}

/** Small bold Oswald label (card titles, list headings, buttons). */
function heading(slide, o) {
  slide.addText(o.text, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: F.head, fontSize: o.size || 16, bold: true,
    color: o.color || C.ink, align: o.align,
    valign: 'top', fit: 'resize',
  });
}

/** Open Sans paragraph, 150% line spacing like the original. */
function body(slide, o) {
  slide.addText(o.text, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: F.body, fontSize: o.size || 12,
    color: o.color || C.grey, italic: o.italic || false,
    align: o.align, valign: 'top',
    lineSpacingMultiple: o.leading || 1.5, fit: 'resize',
  });
}

/* Boiler-plate copy that the template repeats all over the deck. */
const LOREM = 'Beneath the waning light of an amber sky, the village lay nestled in a hush older than memory.';
const LOREM_2 = LOREM + ' Weather-worn cottages leaned gently against one another.';
const LOREM_3 = LOREM + ' Weather-worn cottages leaned gently against one another, their chimneys whispering thin trails of smoke into the still air. ';

/* ------------------------------------------------------------------ *
 * Slide 1 — Title
 * ------------------------------------------------------------------ */
function slide1(pptx) {
  const s = pptx.addSlide();
  rect(s, 0, 5.374, 13.333, 2.126, C.teal);
  picturePlaceholder(s, 0.602, 0.601, 4.449, 6.298);

  body(s, { x: 5.814, y: 0.823, w: 3.256, h: 0.37, text: 'Company Name', size: 16, color: C.ink, leading: 1 });
  title(s, { x: 5.814, y: 1.625, w: 6.467, h: 1.717, text: 'Agriculture', size: 96 });
  body(s, {
    x: 5.814, y: 3.369, w: 6.473, h: 1.63,
    text: 'Innovating Sustainable Growth For The Future', size: 32, italic: true,
  });

  rect(s, 5.814, 6.047, 3.701, 0.551, C.orange);
  s.addText(
    [
      { text: 'Presented', options: { italic: true } },
      { text: ' by ' },
      { text: 'Thomas Hardy', options: { bold: true } },
    ],
    {
      x: 5.911, y: 6.138, w: 3.506, h: 0.37,
      fontFace: F.body, fontSize: 16, color: C.white,
      align: 'center', valign: 'top', fit: 'resize',
    }
  );
}

/* ------------------------------------------------------------------ *
 * Slide 2 — Welcome / two coloured cards
 * ------------------------------------------------------------------ */
function slide2(pptx) {
  const s = pptx.addSlide();
  title(s, { x: 0.602, y: 0.618, w: 5.452, h: 3.13, text: 'Welcome to \nthe Agriculture\nIndustry' });

  heading(s, { x: 6.968, y: 2.101, w: 5.763, h: 0.37, text: 'Global Impact' });
  body(s, { x: 6.968, y: 2.471, w: 5.763, h: 0.978, text: LOREM_2 });

  // Two cards, each: colour block + empty photo frame + label + copy.
  const cards = [
    { x: 0.602, color: C.teal, label: 'Modern Advancements' },
    { x: 6.968, color: C.orange, label: 'Experienced Practices' },
  ];
  cards.forEach((card) => {
    rect(s, card.x, 4.026, 5.763, 2.874, card.color);
    picturePlaceholder(s, card.x, 4.026, 2.362, 2.874);
    heading(s, { x: card.x + 2.748, y: 4.637, w: 2.63, h: 0.37, text: card.label, color: C.white });
    body(s, { x: card.x + 2.748, y: 5.008, w: 2.63, h: 1.28, text: LOREM, color: C.white95 });
  });
}

/* ------------------------------------------------------------------ *
 * Slide 3 — Types of agriculture (numbered columns)
 * ------------------------------------------------------------------ */
function slide3(pptx) {
  const s = pptx.addSlide();
  picturePlaceholder(s, 8.006, 0.601, 4.724, 6.297);
  title(s, { x: 0.602, y: 1.027, w: 4.118, h: 2.121, text: 'Type of\nAgriculture' });

  rect(s, 8.4, 4.064, 3.228, 2.441, C.orange);
  body(s, {
    x: 8.699, y: 4.448, w: 2.63, h: 1.674, size: 16, italic: true, color: C.white95,
    text: 'From basic methods to modern tech, each type helps improve farming and sustainability.',
  });

  const columns = [
    { x: 0.602, num: '01', label: 'Crop Farming' },
    { x: 3.001, num: '02', label: 'Livestock Farming' },
    { x: 5.399, num: '03', label: 'Organic Farming' },
  ];
  columns.forEach((col) => {
    rect(s, col.x + 0.107, 3.75, 0.709, 0.709, C.teal);
    heading(s, {
      x: col.x + 0.12, y: 3.818, w: 0.682, h: 0.572,
      text: col.num, size: 28, color: C.white, align: 'center',
    });
    heading(s, { x: col.x, y: 4.643, w: 2.005, h: 0.37, text: col.label });
    body(s, {
      x: col.x, y: 5.013, w: 2.005, h: 1.886,
      text: LOREM + ' Weather-worn cottages.',
    });
  });
}

/* ------------------------------------------------------------------ *
 * Slide 4 — Current industry challenges (teal slide, orange banner)
 * ------------------------------------------------------------------ */
function slide4(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.teal };
  rect(s, 0, 0, 13.333, 2.312, C.orange);
  title(s, {
    x: 1.709, y: 0.6, w: 9.916, h: 1.111,
    text: 'Current Industry Challenges', color: C.white, align: 'center',
  });
  picturePlaceholder(s, 0.602, 2.312, 6.064, 4.587);

  const items = [
    { y: 2.918, icon: 'question', label: 'Organic Farming' },
    { y: 5.212, icon: 'idea', label: 'Solution' },
  ];
  items.forEach((it) => {
    rect(s, 7.489, it.y, 0.551, 0.551, C.orange);
    iconPlaceholder(s, it.icon, 7.567, it.y + 0.079, 0.394);
    heading(s, { x: 7.383, y: it.y + 0.642, w: 5.348, h: 0.37, text: it.label, color: C.white });
    body(s, { x: 7.383, y: it.y + 1.012, w: 5.348, h: 0.675, text: LOREM + ' ', color: C.white85 });
  });
}

/* ------------------------------------------------------------------ *
 * Slide 5 — Modern agricultural technologies
 * ------------------------------------------------------------------ */
function slide5(pptx) {
  const s = pptx.addSlide();
  picturePlaceholder(s, 8.504, 0, 4.829, 4.154);
  rect(s, 0, 4.154, 8.504, 3.346, C.teal);
  title(s, {
    x: 0.602, y: 4.767, w: 7.111, h: 2.121,
    text: 'Modern Agricultural\nTechnologies', color: C.white,
  });

  const bullets = [
    { y: 0.6, label: 'Smart Farming Tools' },
    { y: 2.377, label: 'AI & Data Analytics' },
  ];
  bullets.forEach((b) => {
    rect(s, 0.602, b.y + 0.048, 0.276, 0.276, C.teal);
    heading(s, { x: 0.988, y: b.y, w: 6.914, h: 0.37, text: b.label });
    body(s, { x: 0.988, y: b.y + 0.3702, w: 6.914, h: 0.675, text: LOREM_2 });
  });

  heading(s, { x: 9.106, y: 5.018, w: 3.625, h: 0.64, text: '40%', size: 32, color: C.orange });
  body(s, { x: 9.106, y: 5.658, w: 3.625, h: 0.978, text: LOREM + ' ' });
}

/* ------------------------------------------------------------------ *
 * Slide 6 — Industry statistics
 * ------------------------------------------------------------------ */
function slide6(pptx) {
  const s = pptx.addSlide();
  picturePlaceholder(s, 0, 0, 5.354, 7.5);
  rect(s, 5.354, 4.744, 3.989, 2.756, C.teal);
  rect(s, 9.344, 4.744, 3.99, 2.756, C.orange);

  title(s, { x: 6.064, y: 0.6, w: 6.196, h: 2.121, text: 'Agriculture\nIndustry Statistic' });
  body(s, { x: 6.064, y: 3.104, w: 6.667, h: 0.978, text: LOREM_3 });

  const stats = [
    { x: 6.01, value: '65%' },
    { x: 10.0, value: '+2,713' },
  ];
  stats.forEach((st) => {
    heading(s, { x: st.x, y: 5.465, w: 2.677, h: 0.64, text: st.value, size: 32, color: C.white });
    body(s, {
      x: st.x, y: 6.105, w: 2.677, h: 0.675,
      text: 'Beneath the waning light of an amber sky.', color: C.white,
    });
  });
}

/* ------------------------------------------------------------------ *
 * Slide 7 — Sustainable solutions
 * ------------------------------------------------------------------ */
function slide7(pptx) {
  const s = pptx.addSlide();
  title(s, { x: 0.602, y: 0.919, w: 7.537, h: 2.121, text: 'Sustainable\nAgriculture Solutions' });
  body(s, {
    x: 9.09, y: 1.456, w: 3.641, h: 1.583,
    text: LOREM + ' Weather-worn cottages leaned gently against one another, their chimneys whispering thin trails of smoke.',
  });
  picturePlaceholder(s, 9.09, 3.433, 3.641, 3.148);

  const cards = [
    { x: 0.602, color: C.teal, label: 'Solution 01' },
    { x: 4.733, color: C.orange, label: 'Solution 02' },
  ];
  cards.forEach((card) => {
    rect(s, card.x, 3.787, 3.74, 2.441, card.color);
    heading(s, { x: card.x + 0.364, y: 4.333, w: 3.012, h: 0.37, text: card.label, color: C.white });
    body(s, { x: card.x + 0.364, y: 4.703, w: 3.012, h: 0.978, text: LOREM + ' ', color: C.white95 });
  });
}

/* ------------------------------------------------------------------ *
 * Slide 8 — Farming practices (teal slide, three icon columns)
 * ------------------------------------------------------------------ */
function slide8(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.teal };
  picturePlaceholder(s, 0.602, 0, 3.641, 3.449);
  title(s, {
    x: 4.846, y: 1.328, w: 6.415, h: 2.121,
    text: 'Agriculture\nFarming Practices', color: C.white,
  });

  const columns = [
    { x: 0.602, icon: 'sprout', label: 'Crop Rotation' },
    { x: 4.846, icon: 'wheat', label: 'Organic Fertilization' },
    { x: 9.09, icon: 'cloudSync', label: 'Smart Irrigation Systems' },
  ];
  columns.forEach((col) => {
    rect(s, col.x + 0.092, 4.317, 0.709, 0.709, C.orange);
    iconPlaceholder(s, col.icon, col.x + 0.171, 4.395, 0.551);
    heading(s, { x: col.x, y: 5.249, w: 3.641, h: 0.37, text: col.label, color: C.white });
    body(s, { x: col.x, y: 5.619, w: 3.641, h: 1.28, text: LOREM_2, color: C.white95 });
  });
}

/* ------------------------------------------------------------------ *
 * Slide 9 — Product showcase
 * ------------------------------------------------------------------ */
function slide9(pptx) {
  const s = pptx.addSlide();
  picturePlaceholder(s, 4.961, 0.6, 3.679, 2.953);
  picturePlaceholder(s, 9.052, 0.6, 3.679, 2.953);
  picturePlaceholder(s, 0.602, 3.947, 6.772, 2.953);

  title(s, { x: 0.602, y: 0.6, w: 3.508, h: 3.13, text: 'Our\nProduct\nShowcase' });

  // Chips sitting on the photo frames.
  const chips = [
    { x: 4.961, w: 2.283, tx: 5.125, tw: 1.955, y: 2.687, color: C.teal, label: 'Product 01' },
    { x: 9.052, w: 2.283, tx: 9.216, tw: 1.955, y: 2.687, color: C.teal, label: 'Product 02' },
    { x: 0.602, w: 3.307, tx: 0.895, tw: 2.723, y: 6.033, color: C.orange, label: 'Product Highlight' },
  ];
  chips.forEach((chip) => {
    rect(s, chip.x, chip.y, chip.w, 0.866, chip.color);
    heading(s, {
      x: chip.tx, y: chip.y + 0.215, w: chip.tw, h: 0.438,
      text: chip.label, size: 20, color: C.white, align: 'center',
    });
  });

  const descriptions = [
    { y: 4.385, label: 'Product 01 Description' },
    { y: 5.855, label: 'Product 02 Description' },
  ];
  descriptions.forEach((d) => {
    heading(s, { x: 7.976, y: d.y, w: 4.754, h: 0.37, text: d.label });
    body(s, { x: 7.976, y: d.y + 0.37, w: 4.754, h: 0.675, text: LOREM + ' ' });
  });
}

/* ------------------------------------------------------------------ *
 * Slide 10 — Closing / contact
 * ------------------------------------------------------------------ */
function slide10(pptx) {
  const s = pptx.addSlide();
  picturePlaceholder(s, 0, 0, 13.333, 7.5);
  rect(s, 0.761, 0, 5.276, 5.826, C.teal);

  title(s, {
    x: 1.339, y: 0.648, w: 4.12, h: 2.524, size: 48, color: C.white,
    text: 'Let\u2019s Grow\nThe Future\nOf Agriculture',
  });
  heading(s, { x: 1.339, y: 3.726, w: 3.736, h: 0.438, text: 'Get In Touch', size: 20, color: C.white });

  const contacts = [
    { y: 4.243, icon: 'envelope', text: 'info@yourcompanyname.com' },
    { y: 4.785, icon: 'globe', text: 'www.yourcompanyname.com' },
  ];
  contacts.forEach((c) => {
    rect(s, 1.422, c.y, 0.394, 0.394, C.orange);
    iconPlaceholder(s, c.icon, 1.468, c.y + 0.045, 0.303);
    body(s, { x: 1.904, y: c.y + 0.045, w: 2.962, h: 0.303, text: c.text, color: C.white95, leading: 1 });
  });
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */
function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK_16x9', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'DECK_16x9';
  pptx.title = 'Agriculture';

  [slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10]
    .forEach((fn) => fn(pptx));

  return pptx.writeFile({
    fileName: path.join(__dirname, '16ce1c5a-f7d1-42ae-bbfc-61512ec70065_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => {
  console.error(e);
  process.exit(1);
});
