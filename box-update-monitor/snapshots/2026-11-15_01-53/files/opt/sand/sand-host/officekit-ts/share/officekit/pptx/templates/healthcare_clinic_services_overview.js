/**
 * "Dr.Cure" medical pitch deck — recreated with pptxgenjs.
 * Slide size 13.3333 x 7.5 in (16:9). Run: node <this file>
 * The raster/SVG icons of the source deck are redrawn here with native shapes.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette / type
const DARK = '16191B';  // theme accent1 — cards, pills, logo mark
const GREEN = '16CA9F'; // theme accent2 — highlights, icons
const INK = '262626';   // headings (tx1 lumMod 15%)
const GRAY = '808080';  // body copy (bg1 lumMod -50%)
const WHITE = 'FFFFFF';

const HEAD = 'Poppins Medium'; // theme major font
const BODY = 'Nunito';         // theme minor font

// PowerPoint text-box insets: 0.1" left/right, 0.05" top/bottom -> [l, r, b, t] in points.
const INSET = { margin: [7.2, 7.2, 3.6, 3.6] };

// ---------------------------------------------------------------- text helpers
/** Heading text (Poppins Medium). `runs` = string, or array of {text, color, size, br}. */
function head(slide, runs, o) {
  const parts = (typeof runs === 'string' ? [{ text: runs }] : runs).map(r => ({
    text: r.text,
    options: { color: r.color || o.color || INK, fontSize: r.size || o.size, breakLine: !!r.br }
  }));
  slide.addText(parts, Object.assign({
    fontFace: HEAD, fontSize: o.size, color: o.color || INK, valign: 'top'
  }, INSET, o));
}

/** Body copy (Nunito, 150% line spacing by default). */
function body(slide, text, o) {
  slide.addText(text, Object.assign({
    fontFace: BODY, fontSize: 12, color: GRAY,
    lineSpacingMultiple: o.lineSpacingMultiple === undefined ? 1.5 : o.lineSpacingMultiple,
    valign: 'top'
  }, INSET, o));
}

// ---------------------------------------------------------------- shape helpers
const NOLINE = { type: 'none' };

function card(slide, x, y, w, h, fill, radius) {
  slide.addShape('roundRect', { x, y, w, h, fill: { color: fill }, rectRadius: radius, line: NOLINE });
}

function pill(slide, x, y, w, h, fill) {
  slide.addShape('roundRect', { x, y, w, h, fill: { color: fill }, rectRadius: h / 2, line: NOLINE });
}

function dot(slide, x, y, d, fill) {
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color: fill }, line: NOLINE });
}

/** Rounded bar drawn between two fractional points [0..1] of an s-sized icon box. */
function bar(slide, x, y, s, a, b, thick, color) {
  const dx = (b[0] - a[0]) * s, dy = (b[1] - a[1]) * s, len = Math.hypot(dx, dy);
  slide.addShape('roundRect', {
    x: x + (a[0] + b[0]) / 2 * s - len / 2, y: y + (a[1] + b[1]) / 2 * s - thick * s / 2,
    w: len, h: thick * s, rectRadius: thick * s / 2, fill: { color }, line: NOLINE,
    rotate: Math.atan2(dy, dx) * 180 / Math.PI
  });
}

/** Straight segment between two fractional points, optionally arrow-tipped. */
function seg(slide, x, y, s, a, b, color, weight, arrow) {
  slide.addShape('line', {
    x: x + Math.min(a[0], b[0]) * s, y: y + Math.min(a[1], b[1]) * s,
    w: Math.abs(b[0] - a[0]) * s, h: Math.abs(b[1] - a[1]) * s,
    flipV: (b[0] - a[0]) * (b[1] - a[1]) < 0,
    line: { color, width: weight, endArrowType: arrow ? 'arrow' : undefined }
  });
}

// ---------------------------------------------------------------- icon primitives
/** Stethoscope logo mark: ear tips, U-shaped binaural, tubing and chest piece. */
function iconStethoscope(slide, x, y, s, color) {
  const f = { color }, t = 0.10 * s;
  [0.04, 0.35].forEach(px => slide.addShape('rect', { x: x + px * s, y: y + 0.06 * s, w: 0.23 * s, h: 0.13 * s, fill: f, line: NOLINE }));
  [0.04, 0.48].forEach(px => slide.addShape('rect', { x: x + px * s, y: y + 0.06 * s, w: t, h: 0.30 * s, fill: f, line: NOLINE }));
  slide.addShape('blockArc', { x: x + 0.04 * s, y: y + 0.08 * s, w: 0.54 * s, h: 0.54 * s, fill: f, line: NOLINE, angleRange: [0, 180], arcThicknessRatio: 0.37 });
  slide.addShape('rect', { x: x + 0.27 * s, y: y + 0.55 * s, w: t, h: 0.26 * s, fill: f, line: NOLINE });
  slide.addShape('blockArc', { x: x + 0.28 * s, y: y + 0.62 * s, w: 0.56 * s, h: 0.34 * s, fill: f, line: NOLINE, angleRange: [0, 180], arcThicknessRatio: 0.5 });
  slide.addShape('rect', { x: x + 0.74 * s, y: y + 0.50 * s, w: t, h: 0.30 * s, fill: f, line: NOLINE });
  dot(slide, x + 0.66 * s, y + 0.26 * s, 0.28 * s, color);
}

/** Arrow with a chevron head: 'ne' (the default) or 'e'. */
function iconArrow(slide, x, y, s, color, dir) {
  const t = 0.115;
  const shaft = dir === 'e' ? [[0.12, 0.5], [0.88, 0.5]] : [[0.16, 0.84], [0.84, 0.16]];
  const head = dir === 'e' ? [[0.60, 0.22], [0.60, 0.78]] : [[0.40, 0.16], [0.84, 0.60]];
  bar(slide, x, y, s, shaft[0], shaft[1], t, color);
  head.forEach(p => bar(slide, x, y, s, p, shaft[1], t, color));
}

/** Head-and-shoulders avatar. */
function iconPerson(slide, x, y, s, color) {
  dot(slide, x + 0.32 * s, y + 0.13 * s, 0.36 * s, color);
  slide.addShape('blockArc', { x: x + 0.14 * s, y: y + 0.44 * s, w: 0.72 * s, h: 0.68 * s, fill: { color }, line: NOLINE, angleRange: [180, 360], arcThicknessRatio: 0.95 });
}

/** Rising zig-zag trend line. */
function iconChart(slide, x, y, s, color) {
  const pts = [[0.10, 0.72], [0.36, 0.44], [0.54, 0.58], [0.88, 0.24]];
  for (let i = 0; i < 3; i++) bar(slide, x, y, s, pts[i], pts[i + 1], 0.11, color);
}

/** Circled tick used for the check-list bullets. */
function iconCheckCircle(slide, x, y, s, ring, tick) {
  dot(slide, x, y, s, ring);
  bar(slide, x, y, s, [0.28, 0.52], [0.44, 0.68], 0.11, tick);
  bar(slide, x, y, s, [0.42, 0.68], [0.74, 0.34], 0.11, tick);
}

/** Thumbs-up: cuff bar on the left, palm block and the raised thumb. */
function iconThumbsUp(slide, x, y, s, color) {
  const f = { color };
  slide.addShape('rect', { x: x + 0.06 * s, y: y + 0.38 * s, w: 0.15 * s, h: 0.53 * s, fill: f, line: NOLINE });
  slide.addShape('roundRect', { x: x + 0.26 * s, y: y + 0.30 * s, w: 0.68 * s, h: 0.61 * s, fill: f, line: NOLINE, rectRadius: 0.10 * s });
  slide.addShape('round2SameRect', { x: x + 0.33 * s, y: y + 0.09 * s, w: 0.22 * s, h: 0.34 * s, fill: f, line: NOLINE, rotate: 15 });
}

/** Map pin. */
function iconPin(slide, x, y, s, color, hole) {
  slide.addShape('teardrop', { x: x + 0.14 * s, y: y + 0.04 * s, w: 0.72 * s, h: 0.72 * s, fill: { color }, line: NOLINE, rotate: 135 });
  dot(slide, x + 0.39 * s, y + 0.20 * s, 0.22 * s, hole);
}

/** Envelope with a V-shaped flap. */
function iconMail(slide, x, y, s, color, flap) {
  const w = Math.max(1, 7 * s);
  slide.addShape('roundRect', { x: x + 0.04 * s, y: y + 0.18 * s, w: 0.92 * s, h: 0.64 * s, fill: { color }, line: NOLINE, rectRadius: 0.09 * s });
  seg(slide, x, y, s, [0.14, 0.26], [0.50, 0.50], flap, w);
  seg(slide, x, y, s, [0.50, 0.50], [0.86, 0.26], flap, w);
}

/** Telephone handset: a quarter arc with a receiver cap at each end. */
function iconPhone(slide, x, y, s, color) {
  const f = { color }, r = 0.80 * s, cx = x + 0.90 * s, cy = y + 0.12 * s;
  slide.addShape('blockArc', { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, fill: f, line: NOLINE, angleRange: [90, 180], arcThicknessRatio: 0.32 });
  slide.addShape('roundRect', { x: x + 0.04 * s, y: y + 0.06 * s, w: 0.30 * s, h: 0.28 * s, fill: f, line: NOLINE, rectRadius: 0.08 * s });
  slide.addShape('roundRect', { x: x + 0.62 * s, y: y + 0.64 * s, w: 0.32 * s, h: 0.28 * s, fill: f, line: NOLINE, rectRadius: 0.08 * s });
}

// ---------------------------------------------------------------- chrome on every slide
function brand(slide) {
  iconStethoscope(slide, 0.527, 0.418, 0.316, DARK);
  head(slide, 'Dr.Cure', { x: 0.832, y: 0.432, w: 1.105, h: 0.286, size: 11, color: GREEN });
}

function pageBadge(slide, num, disc, ink) {
  dot(slide, 12.424, 6.657, 0.485, disc || DARK);
  head(slide, num, { x: 12.424, y: 6.748, w: 0.485, h: 0.303, size: 12, color: ink || WHITE, align: 'center' });
}

// ================================================================= slides
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore.';
const LOREM_MED = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim';
const LOREM_TINY = 'Lorem ipsum dolor sit amet, consectetur adipiscing.';

// --- 1 : hero -----------------------------------------------------
function slide1(pres) {
  const s = pres.addSlide();
  brand(s);
  head(s, 'Menu', { x: 11.444, y: 0.432, w: 1.105, h: 0.286, size: 11, align: 'center' });
  [0.496, 0.575, 0.655].forEach(y => {
    s.addShape('line', { x: 12.539, y, w: 0.282, h: 0, line: { color: DARK, width: 1.5 } });
  });

  head(s, [
    { text: 'Innovative Approaches ', br: true },
    { text: 'to ' }, { text: 'Health Care', color: GREEN }
  ], { x: 1.488, y: 1.762, w: 5.664, h: 2.524, size: 48 });
  body(s, LOREM_SHORT, { x: 1.488, y: 4.566, w: 4.876, h: 0.682 });

  head(s, 'Check Up Your Health Now', { x: 1.485, y: 5.686, w: 2.679, h: 0.303, size: 12 });
  iconArrow(s, 3.930, 5.721, 0.234, GREEN);

  card(s, 7.465, 1.648, 1.377, 1.150, DARK, 0.101);
  head(s, '35K', { x: 7.546, y: 1.709, w: 1.265, h: 0.572, size: 28, color: GREEN });
  head(s, 'Has Already Contact Us', { x: 7.546, y: 2.199, w: 1.265, h: 0.505, size: 12, color: WHITE });

  pill(s, 6.673, 5.601, 2.411, 0.474, DARK);
  head(s, 'Dr. Samantha Ann', { x: 7.206, y: 5.686, w: 1.955, h: 0.303, size: 12, color: WHITE });
  dot(s, 6.719, 5.640, 0.396, GREEN);
  iconPerson(s, 6.762, 5.683, 0.309, WHITE);

  pageBadge(s, '01');
}

// --- 2 : table of content ----------------------------------------
const TOC = ['About Us', 'What We Do', 'Great Service', 'Our Doctors', 'Pricing', 'Our Portfolio', 'Great Quotes', 'Get in Touch'];

function slide2(pres) {
  const s = pres.addSlide();
  brand(s);
  head(s, 'Table of Content', { x: 1.776, y: 4.247, w: 6.240, h: 0.909, size: 48 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.',
    { x: 1.806, y: 5.460, w: 5.588, h: 0.682 });

  card(s, 8.258, 0.981, 3.478, 5.663, DARK, 0.221);
  TOC.forEach((label, i) => {
    const y = 1.451 + i * 0.6044;
    head(s, '0' + (i + 1) + '.', { x: 8.828, y: y + 0.062, w: 0.750, h: 0.404, size: 18, color: GREEN });
    head(s, label, { x: 9.478, y, w: 2.685, h: 0.462, size: 16, color: WHITE, lineSpacingMultiple: 1.5 });
  });

  pageBadge(s, '02');
}

// --- 3 : about us -------------------------------------------------
function slide3(pres) {
  const s = pres.addSlide();
  brand(s);
  head(s, 'About Us', { x: 1.689, y: 1.727, w: 4.978, h: 1.010, size: 54 });

  card(s, 6.506, 1.005, 2.034, 2.501, DARK, 0.156);
  card(s, 6.756, 1.247, 0.476, 0.476, GREEN, 0.079);
  iconChart(s, 6.786, 1.277, 0.417, WHITE);
  head(s, '148K', { x: 7.369, y: 1.258, w: 1.184, h: 0.505, size: 24, color: WHITE });
  body(s, 'Lorem ipsum dolor sit amet, adipiscing elit, aliqua. Ut', { x: 6.679, y: 1.862, w: 1.767, h: 0.985, color: WHITE });
  head(s, 'Read More', { x: 6.679, y: 3.007, w: 1.767, h: 0.286, size: 11, color: WHITE, italic: true });

  head(s, 'General Surgeon', { x: 6.894, y: 4.363, w: 3.194, h: 0.337, size: 14 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo',
    { x: 6.894, y: 4.993, w: 4.725, h: 1.287 });

  pageBadge(s, '03');
}

// --- 4 : smarter diagnosis ---------------------------------------
function slide4(pres) {
  const s = pres.addSlide();
  brand(s);
  head(s, 'Smarter Diagnosis', { x: 1.382, y: 1.526, w: 4.087, h: 1.717, size: 48 });
  body(s, LOREM_MED, { x: 1.382, y: 4.066, w: 4.087, h: 0.985 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut', { x: 1.382, y: 5.176, w: 4.087, h: 0.682 });

  card(s, 9.152, 1.375, 3.048, 2.421, DARK, 0.141);
  card(s, 11.525, 1.561, 0.476, 0.476, GREEN, 0.079);
  iconArrow(s, 11.555, 1.590, 0.417, WHITE);
  head(s, 'Analysing Vast Amounts of Data, Enabling Smarter Diagnosis', { x: 9.320, y: 2.748, w: 2.835, h: 0.808, size: 14, color: WHITE });

  head(s, 'Coordination of Care', { x: 9.320, y: 4.321, w: 3.194, h: 0.337, size: 14 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod.', { x: 9.320, y: 4.873, w: 2.298, h: 0.985 });

  pageBadge(s, '04');
}

// --- 5 : our service ----------------------------------------------
function slide5(pres) {
  const s = pres.addSlide();
  brand(s);
  head(s, 'Our Service', { x: 1.428, y: 1.513, w: 3.048, h: 1.717, size: 48 });
  body(s, LOREM_MED, { x: 1.428, y: 3.951, w: 3.352, h: 1.287 });
  pill(s, 1.512, 5.630, 1.915, 0.474, DARK);
  head(s, 'Read More', { x: 1.654, y: 5.715, w: 1.279, h: 0.303, size: 12, color: WHITE });
  dot(s, 2.983, 5.669, 0.396, GREEN);
  iconArrow(s, 3.046, 5.732, 0.269, WHITE, 'e');

  card(s, 5.145, 1.191, 3.048, 5.525, DARK, 0.178);
  head(s, 'Key Feature of Effective Health Care', { x: 5.393, y: 1.629, w: 2.298, h: 0.572, size: 14, color: WHITE });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod.', { x: 5.393, y: 2.324, w: 2.298, h: 0.985, color: WHITE });
  s.addShape('line', { x: 5.393, y: 3.733, w: 2.396, h: 0, line: { color: WHITE, width: 0.5 } });
  card(s, 7.313, 4.009, 0.476, 0.476, GREEN, 0.079);
  iconArrow(s, 7.343, 4.038, 0.417, WHITE);
  head(s, 'Providing High- Quality, Evidence-Based Care that Meets Professional Standards and Patient Needs.',
    { x: 5.393, y: 5.034, w: 2.718, h: 1.279, size: 14, color: WHITE });

  [1.507, 3.391, 5.274].forEach(y => {
    head(s, 'Service #01', { x: 8.935, y, w: 2.298, h: 0.337, size: 14 });
    body(s, 'Lorem ipsum dolor sit amet, adipiscing elit, sed.', { x: 8.935, y: y + 0.438, w: 2.298, h: 0.682 });
  });

  pageBadge(s, '05');
}

// --- 6 : meet the doctors ----------------------------------------
const DOCTORS = [
  { card: 6.128, text: 6.406, name: 'Dr . Samsul Fernand', role: '(Internal Medicine Doctor)' },
  { card: 9.166, text: 9.463, name: 'Dr . Andy Samuel', role: '(Gastroenterologist)' }
];

function slide6(pres) {
  const s = pres.addSlide();
  brand(s);
  head(s, 'Meet The Doctors', { x: 1.488, y: 1.693, w: 3.951, h: 1.717, size: 48 });
  iconCheckCircle(s, 1.553, 4.512, 0.417, GREEN, WHITE);
  head(s, '10 Years of Experience', { x: 2.013, y: 4.552, w: 2.528, h: 0.337, size: 14 });
  body(s, LOREM_MED, { x: 1.488, y: 5.028, w: 4.103, h: 0.985 });

  DOCTORS.forEach(d => {
    card(s, d.card, 1.139, 2.854, 5.222, DARK, 0.167);
    head(s, d.name, { x: d.text, y: 4.452, w: 2.298, h: 0.337, size: 14, color: WHITE, align: 'center' });
    body(s, d.role, { x: d.text, y: 4.709, w: 2.298, h: 0.379, color: WHITE, align: 'center' });
    body(s, 'Lorem ipsum dolor sit amet, elit, sed do eiusmod.', { x: d.text, y: 5.266, w: 2.298, h: 0.682, color: WHITE, align: 'center' });
  });

  pageBadge(s, '06');
}

// --- 7 : pricing ---------------------------------------------------
function slide7(pres) {
  const s = pres.addSlide();
  s.addShape('rect', { x: 9.735, y: 0, w: 3.598, h: 7.5, fill: { color: DARK }, line: NOLINE });
  brand(s);
  head(s, 'Consultation Pricing Plan', { x: 1.658, y: 1.659, w: 6.177, h: 1.717, size: 48 });
  head(s, 'Book Your Consultation With Our Available Doctors', { x: 1.658, y: 3.591, w: 3.756, h: 0.572, size: 14 });
  [{ n: '01.', x: 1.658 }, { n: '02.', x: 4.243 }].forEach(c => {
    head(s, c.n, { x: c.x, y: 4.916, w: 0.684, h: 0.337, size: 14 });
    body(s, LOREM_TINY, { x: c.x, y: 5.274, w: 2.342, h: 0.682 });
  });

  card(s, 7.834, 1.360, 4.135, 5.018, GREEN, 0.213);
  head(s, [{ text: '$66' }, { text: '/month', size: 11 }],
    { x: 7.834, y: 1.750, w: 4.135, h: 1.111, size: 60, color: WHITE, align: 'center' });
  head(s, 'Professionals Price', { x: 7.834, y: 2.955, w: 4.135, h: 0.438, size: 20, color: WHITE, align: 'center' });
  s.addShape('line', { x: 8.594, y: 3.749, w: 2.615, h: 0, line: { color: DARK, width: 3 } });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et.',
    { x: 8.226, y: 4.130, w: 3.352, h: 0.985, color: WHITE, align: 'center' });
  pill(s, 8.898, 5.513, 2.006, 0.474, DARK);
  head(s, 'Book Now', { x: 8.898, y: 5.513, w: 2.006, h: 0.474, size: 11, color: WHITE, align: 'center', valign: 'middle' });

  pageBadge(s, '07', WHITE, INK);
}

// --- 8 : portfolio details ----------------------------------------
function slide8(pres) {
  const s = pres.addSlide();
  brand(s);
  head(s, 'Portfolio Details', { x: 8.284, y: 1.429, w: 3.478, h: 1.582, size: 44 });
  [[3.846, 3.726], [4.664, 4.544], [5.483, 5.363]].forEach(([iconY, textY]) => {
    iconCheckCircle(s, 8.388, iconY, 0.417, GREEN, WHITE);
    body(s, LOREM_TINY, { x: 8.971, y: textY, w: 2.342, h: 0.682 });
  });

  card(s, 3.297, 1.481, 1.964, 1.519, DARK, 0.078);
  iconThumbsUp(s, 4.098, 1.696, 0.363, GREEN);
  head(s, [{ text: 'Best Rating ', br: true }, { text: 'All-Time' }],
    { x: 3.404, y: 2.245, w: 1.752, h: 0.572, size: 14, color: WHITE, align: 'center' });

  pageBadge(s, '08');
}

// --- 9 : quote -----------------------------------------------------
function slide9(pres) {
  const s = pres.addSlide();
  brand(s);
  pill(s, 1.804, 1.808, 2.527, 0.474, DARK);
  head(s, 'Dr. Vega Vanderman', { x: 2.293, y: 1.894, w: 2.202, h: 0.303, size: 12, color: WHITE });
  dot(s, 1.849, 1.847, 0.396, GREEN);
  iconPerson(s, 1.893, 1.891, 0.309, WHITE);

  head(s, [
    { text: '\u201C' }, { text: 'Recovery', color: GREEN },
    { text: ' Often Requires Time, Yet Sometimes it Also ' }, { text: 'Depends', color: GREEN },
    { text: ' on The Right Moment.\u201D' }
  ], { x: 1.685, y: 2.532, w: 10.009, h: 2.322, size: 44 });

  pageBadge(s, '09');
}

// --- 10 : contact --------------------------------------------------
const CONTACT = [
  { icon: iconPin,   iconX: 8.767, iconY: 2.830, y: 2.760, label: 'Address',      lines: [{ t: 'Street Wagir, 14, Malang, Indonesia', y: 3.042, w: 2.067, h: 0.471 }] },
  { icon: iconMail,  iconX: 8.767, iconY: 3.807, y: 3.762, label: 'Email',        lines: [{ t: 'yourgreate@mail.com', y: 4.013, w: 2.065, h: 0.286 }] },
  { icon: iconPhone, iconX: 8.792, iconY: 4.623, y: 4.548, label: 'Phone Number', lines: [{ t: '+123 456 789 10', y: 4.830, w: 1.940, h: 0.286 }, { t: '+123 8 52 219 22', y: 5.103, w: 1.940, h: 0.286 }] }
];

function slide10(pres) {
  const s = pres.addSlide();
  brand(s);
  head(s, 'Get in Touch!', { x: 1.888, y: 1.417, w: 6.013, h: 1.010, size: 54 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim.',
    { x: 1.948, y: 2.785, w: 5.953, h: 0.682 });

  card(s, 8.424, 1.094, 2.952, 4.612, DARK, 0.152);
  head(s, 'Reach and Get Offer', { x: 8.701, y: 1.416, w: 2.690, h: 0.303, size: 12, color: WHITE, bold: true });
  body(s, LOREM_TINY, { x: 8.701, y: 1.794, w: 2.281, h: 0.633, fontSize: 11, color: WHITE });

  CONTACT.forEach(c => {
    c.icon(s, c.iconX, c.iconY, 0.311, GREEN, DARK);
    head(s, c.label, { x: 9.138, y: c.y, w: 1.860, h: 0.303, size: 12, color: WHITE, bold: true });
    c.lines.forEach(l => {
      body(s, l.t, { x: 9.138, y: l.y, w: l.w, h: l.h, fontSize: 11, color: WHITE, lineSpacingMultiple: 1 });
    });
  });

  pageBadge(s, '10');
}

// ================================================================= build
const pres = new PptxGenJS();
pres.defineLayout({ name: 'DECK', width: 13.333333, height: 7.5 });
pres.layout = 'DECK';
pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };
pres.author = 'Dr.Cure';
pres.title = 'Innovative Approaches to Health Care';

[slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10].forEach(fn => fn(pres));

pres.writeFile({ fileName: path.join(__dirname, '08f00f4d-bf3d-4ef7-8d45-f7bcc0959498_grok_final.pptx') })
  .then(f => console.log('wrote', f));
