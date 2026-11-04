/**
 * Recreation of "Building Construction" industry presentation template (20 slides, 16:9).
 * Pure pptxgenjs — every shape, colour and string below is a plain literal.
 * Raster photos in the original deck are replaced by grey "[image]" placeholders.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const BLUE = '4D5ADC'; // accent1 / accent3
const ORANGE = 'F97104'; // accent2
const NAVY = '2532BA'; // darker blue variant
const DEEP = '18217C'; // darkest blue variant
const LILAC = '949CEA'; // light blue variant
const PEACH = 'FDAA67'; // light orange variant
const RUST = 'BB5503'; // dark orange variant
const INK = '262626'; // body text (tx1 lum 85/15)
const WHITE = 'FFFFFF';
const PANEL = 'F2F2F2'; // title-slide background panel
const IMG_FILL = 'CFCFCF'; // photo placeholder body
const IMG_TEXT = 'A8A8A8'; // photo placeholder caption

const HEAD = 'Archivo'; // display face
const BODY = 'Work Sans'; // text face

// pptxgenjs rewrites the shadow object it is handed, so hand it a fresh copy each time.
const shadow = (blur, offset, angle) => ({ type: 'outer', blur: blur, offset: offset, angle: angle, color: '000000', opacity: 0.2 });
const SHADOW_CARD = () => shadow(50, 20, 45); // soft drop under the white info cards
const SHADOW_SOFT = () => shadow(17, 5, 45); // tight drop under small tiles
const SHADOW_BAR = () => shadow(23, 10, 180); // sideways drop under the ribbon bars

/* --------------------------------------------------------------- primitives */

// Solid rectangle / preset shape.
function box(s, shape, o) {
  s.addShape(shape, o);
}

// Text block. `runs` is a string or an array of {text, options} pptxgenjs runs.
// `wrap: false` marks the deck's shrink-to-fit labels, whose box hugs the glyphs,
// so they are centred in the box to survive font-substitution.
function text(s, runs, o) {
  const hug = o.wrap === false ? { align: 'center', valign: 'middle' } : {};
  s.addText(runs, Object.assign({ fontFace: BODY, color: INK, fontSize: 18 }, hug, o));
}

// Box widths below are inches rounded to 3 decimals, so a box can end up half a
// thousandth of an inch narrower than the original and drop a trailing word to
// the next line. SLACK covers that rounding error without changing the layout.
const SLACK = 0.0005;

// Paragraph of justified 10.5 pt body copy on 1.5 line spacing (the deck's default).
function para(s, str, x, y, w, h, extra) {
  text(s, str, Object.assign({
    x: x, y: y, w: w + SLACK, h: h, fontSize: 10.5, align: 'justify',
    lineSpacingMultiple: 1.5, valign: 'top',
  }, extra || {}));
}

// Bold sub-heading (18 pt Work Sans) used above body copy in the content cards.
function subHead(s, str, x, y, w, color) {
  text(s, str, { x: x, y: y, w: w, h: 0.404, fontSize: 18, bold: true, color: color, wrap: false });
}

// Grey rectangle standing in for a photograph in the source deck.
function photo(s, x, y, w, h) {
  box(s, 'rect', { x: x, y: y, w: w, h: h, fill: { color: IMG_FILL }, line: { type: 'none' } });
  text(s, '[image]', { x: x, y: y, w: w, h: h, fontSize: 11, color: IMG_TEXT, align: 'center', valign: 'middle' });
}

// Free-form polygon; `pts` are fractions of the bounding box.
function poly(s, x, y, w, h, color, pts, extra) {
  s.addShape('custGeom', Object.assign({
    x: x, y: y, w: w, h: h, fill: { color: color }, line: { type: 'none' },
    points: pts.map(p => ({ x: w * p[0], y: h * p[1] })).concat([{ close: true }]),
  }, extra || {}));
}
const quad = poly; // alias used where the polygon is a simple four-sided slab

// Block arrow with an explicit head height and shaft width (both fractions of
// the box), since pptxgenjs cannot set preset-shape adjustment handles.
function blockArrow(s, x, y, w, h, color, headH, shaftW, dir) {
  const a = (1 - shaftW) / 2;
  const b = 1 - a;
  const up = [[0, headH], [0.5, 0], [1, headH], [b, headH], [b, 1], [a, 1], [a, headH]];
  const down = up.map(p => [p[0], 1 - p[1]]);
  poly(s, x, y, w, h, color, dir === 'down' ? down : up);
}

// Trapezoid with a given top inset (fraction of the width); `flipV` puts the
// wide edge on top, matching the funnel slices in the deck.
function trap(s, x, y, w, h, color, inset, flipV) {
  const pts = [[0, 1], [inset, 0], [1 - inset, 0], [1, 1]];
  poly(s, x, y, w, h, color, flipV ? pts.map(p => [p[0], 1 - p[1]]) : pts);
}

// Two-tone 37 pt page title, e.g. "Best Our" (blue) + "Project" (orange).
function pageTitle(s, blueText, orangeText, x, y, w, h, align) {
  text(s, [
    { text: blueText, options: { color: BLUE } },
    { text: orangeText, options: { color: ORANGE } },
  ], { x: x, y: y, w: w + SLACK, h: h, fontSize: 37, fontFace: HEAD, align: align || 'left', valign: 'top' });
}

/* ------------------------------------------------------------- page furniture */

const NAV_LINKS = [
  { label: 'Home', w: 0.623 },
  { label: 'Service', w: 0.723 },
  { label: 'About', w: 0.631 },
  { label: 'Contact', w: 0.772 },
];

// The 3x3 orange dot-grid logo mark.
function logoMark(s, x, y) {
  const step = 0.1035;
  const dot = 0.069;
  for (let r = 0; r < 3; r++) {
    for (let c = 0; c < 3; c++) {
      box(s, 'roundRect', {
        x: x + c * step, y: y + r * step, w: dot, h: dot,
        fill: { color: ORANGE }, line: { type: 'none' }, rectRadius: 0.012,
      });
    }
  }
}

// Top navigation bar. Layout "wide" (slides 2-19) also carries the dot-grid logo.
function navBar(s, layout) {
  const wide = layout === 'wide';
  const brandX = wide ? 1.468 : 1.838;
  const linkX = wide ? [7.991, 9.024, 10.274, 11.524] : [7.202, 8.234, 9.484, 10.734];
  if (wide) logoMark(s, 1.155, 0.463);
  text(s, [
    { text: 'Building', options: { color: ORANGE } },
    { text: ' ', options: { color: INK } },
    { text: 'Construction', options: { color: BLUE } },
  ], { x: brandX, y: 0.464, w: 1.613, h: 0.278, fontSize: 10.5, fontFace: HEAD, wrap: false });
  NAV_LINKS.forEach((link, i) => {
    text(s, link.label, { x: linkX[i], y: 0.464, w: link.w, h: 0.278, fontSize: 10.5, wrap: false });
  });
}

/* -------------------------------------------------------- decorative icon set */

// The reference deck draws its pictograms as freeform vector groups. Each is
// rebuilt here from native shapes; `kind` selects the motif and `bg` is the
// colour behind the icon, used to punch out negative space.
function icon(s, kind, x, y, size, color, bg) {
  const back = bg || WHITE;
  const S = (fx, fy, fw, fh, shape, extra) => box(s, shape || 'rect', Object.assign({
    x: x + size * fx, y: y + size * fy, w: size * fw, h: size * fh,
    fill: { color: color }, line: { type: 'none' },
  }, extra || {}));
  const cut = (fx, fy, fw, fh, shape, extra) => S(fx, fy, fw, fh, shape, Object.assign({ fill: { color: back } }, extra || {}));
  const stroke = (fx, fy, fw, fh, shape, extra) => S(fx, fy, fw, fh, shape,
    Object.assign({ fill: { type: 'none' }, line: { color: color, width: size * 4 } }, extra || {}));

  switch (kind) {
    case 'bars': // rising bar chart under a trend line
      S(0.04, 0.56, 0.2, 0.44); S(0.34, 0.36, 0.2, 0.64); S(0.64, 0.16, 0.2, 0.84);
      S(0.2, 0.0, 0.8, 0.06, null, { rotate: -28 });
      break;
    case 'chartbox': // bar chart inside a frame
      stroke(0.02, 0.06, 0.96, 0.88);
      S(0.24, 0.52, 0.14, 0.26); S(0.44, 0.34, 0.14, 0.44); S(0.64, 0.44, 0.14, 0.34);
      break;
    case 'linechart': // zig-zag line over an axis
      S(0.06, 0.16, 0.06, 0.72); S(0.06, 0.82, 0.88, 0.06);
      S(0.18, 0.5, 0.3, 0.05, null, { rotate: 30 });
      S(0.4, 0.42, 0.24, 0.05, null, { rotate: -34 });
      S(0.58, 0.3, 0.32, 0.05, null, { rotate: -22 });
      break;
    case 'clipboard': // ticked task list on a clipboard
      S(0.1, 0.1, 0.8, 0.88, 'roundRect', { rectRadius: size * 0.06 });
      S(0.36, 0.0, 0.28, 0.14, 'roundRect', { rectRadius: size * 0.03 });
      cut(0.28, 0.5, 0.24, 0.09, null, { rotate: 45 });
      cut(0.4, 0.36, 0.38, 0.09, null, { rotate: -45 });
      break;
    case 'person': // head-and-shoulders silhouette
      S(0.36, 0.06, 0.3, 0.3, 'ellipse');
      S(0.16, 0.46, 0.7, 0.54, 'blockArc', { angleRange: [180, 360] });
      break;
    case 'badge': // rosette / award seal
      S(0.0, 0.0, 1.0, 1.0, 'star16');
      cut(0.16, 0.16, 0.68, 0.68, 'ellipse');
      S(0.28, 0.28, 0.44, 0.44, 'star5');
      break;
    case 'cap': // graduation mortar-board with a tassel
      S(0.2, 0.38, 0.6, 0.3, 'trapezoid');
      S(0.0, 0.1, 1.0, 0.42, 'diamond');
      S(0.86, 0.3, 0.05, 0.4);
      S(0.82, 0.66, 0.13, 0.18, 'ellipse');
      break;
    case 'gear': // cog wheel
      S(0.0, 0.0, 1.0, 1.0, 'star8');
      cut(0.3, 0.3, 0.4, 0.4, 'ellipse');
      break;
    case 'flask': // laboratory flask with a rounded belly
      S(0.4, 0.0, 0.2, 0.34);
      S(0.33, 0.0, 0.34, 0.08);
      S(0.1, 0.26, 0.8, 0.56, 'triangle');
      S(0.14, 0.52, 0.72, 0.48, 'ellipse');
      break;
    case 'atom': // nucleus crossed by two elliptical orbits
      stroke(0.0, 0.0, 1.0, 1.0, 'ellipse');
      [45, 135].forEach(a => stroke(0.0, 0.32, 1.0, 0.36, 'ellipse', { rotate: a }));
      S(0.38, 0.38, 0.24, 0.24, 'ellipse');
      break;
    case 'puzzle': // interlocking piece
      S(0.04, 0.12, 0.84, 0.84, 'roundRect', { rectRadius: size * 0.1 });
      S(0.32, 0.0, 0.28, 0.28, 'ellipse');
      cut(0.74, 0.4, 0.28, 0.28, 'ellipse');
      break;
    case 'pie': // disc with one wedge lifted out
      S(0.0, 0.1, 0.9, 0.9, 'pie', { angleRange: [0, 270] });
      S(0.1, 0.0, 0.9, 0.9, 'pie', { angleRange: [270, 360] });
      break;
    case 'bulb': // light bulb
      stroke(0.14, 0.0, 0.72, 0.72, 'ellipse');
      S(0.36, 0.66, 0.28, 0.16);
      S(0.38, 0.86, 0.24, 0.12);
      break;
    case 'briefcase': // case with handle
      stroke(0.34, 0.0, 0.32, 0.24, 'roundRect', { rectRadius: size * 0.04 });
      S(0.02, 0.2, 0.96, 0.72, 'roundRect', { rectRadius: size * 0.06 });
      cut(0.44, 0.42, 0.12, 0.26);
      break;
    case 'globe': // globe with meridian and equator carved out
      S(0.0, 0.0, 1.0, 1.0, 'ellipse');
      cut(0.3, 0.0, 0.4, 1.0, 'ellipse', { fill: { type: 'none' }, line: { color: back, width: size * 3 } });
      cut(0.02, 0.44, 0.96, 0.08);
      break;
    case 'safe': // strongbox with dial
      S(0.02, 0.04, 0.96, 0.92, 'roundRect', { rectRadius: size * 0.08 });
      cut(0.22, 0.24, 0.5, 0.5, 'roundRect', { rectRadius: size * 0.06 });
      S(0.37, 0.39, 0.2, 0.2, 'ellipse');
      break;
    case 'pin': // map marker
      S(0.1, 0.0, 0.8, 0.8, 'ellipse');
      S(0.26, 0.5, 0.48, 0.5, 'triangle', { flipV: true });
      cut(0.38, 0.24, 0.24, 0.24, 'ellipse');
      break;
    case 'notebook': // ring-bound notebook
      S(0.16, 0.0, 0.84, 1.0);
      cut(0.56, 0.08, 0.3, 0.26);
      S(0.0, 0.1, 0.16, 0.12); S(0.0, 0.44, 0.16, 0.12); S(0.0, 0.78, 0.16, 0.12);
      break;
    case 'phone': // rotary handset inside a dial ring
      stroke(0.0, 0.0, 1.0, 1.0, 'ellipse');
      S(0.22, 0.46, 0.56, 0.3, 'trapezoid');
      S(0.16, 0.24, 0.68, 0.16, 'roundRect', { rectRadius: size * 0.07 });
      S(0.12, 0.28, 0.2, 0.22, 'ellipse');
      S(0.68, 0.28, 0.2, 0.22, 'ellipse');
      cut(0.42, 0.54, 0.16, 0.16, 'ellipse');
      break;
    case 'mail': // envelope, flap folded down into a V
      S(0.0, 0.14, 1.0, 0.72);
      cut(0.0, 0.14, 1.0, 0.62, 'triangle', { flipV: true });
      S(0.06, 0.14, 0.88, 0.48, 'triangle', { flipV: true });
      break;
    case 'house': // home with a door
      S(0.0, 0.0, 1.0, 0.56, 'triangle');
      S(0.14, 0.5, 0.72, 0.5);
      cut(0.4, 0.66, 0.2, 0.34);
      break;
    case 'megaphone': // bow-tie speaker mark beside a disc
      S(0.42, 0.1, 0.58, 0.58, 'ellipse');
      S(0.0, 0.18, 0.46, 0.42, 'triangle', { rotate: 90 });
      S(0.26, 0.18, 0.46, 0.42, 'triangle', { rotate: 270 });
      break;
    default: // fallback: outlined tile with a solid core
      stroke(0.0, 0.0, 1.0, 1.0, 'roundRect', { rectRadius: size * 0.22 });
      S(0.32, 0.32, 0.36, 0.36, 'ellipse');
  }
}

// Filled disc holding a white tick.
function checkDot(s, x, y, size, color) {
  box(s, 'ellipse', { x: x, y: y, w: size, h: size, fill: { color: color }, line: { type: 'none' } });
  text(s, '✓', { x: x, y: y, w: size, h: size, fontSize: size * 46, bold: true, color: WHITE, align: 'center', valign: 'middle' });
}

/* ================================================================== slide 01 */
function slide01(pptx) {
  const s = pptx.addSlide();
  box(s, 'rect', { x: 1.179, y: 0, w: 11.011, h: 7.5, fill: { color: PANEL }, line: { type: 'none' } });
  box(s, 'rect', { x: 11.189, y: 2.947, w: 1, h: 1, fill: { color: BLUE }, line: { type: 'none' } });
  box(s, 'rightArrow', { x: 11.47, y: 3.266, w: 0.439, h: 0.363, fill: { color: WHITE }, line: { type: 'none' } });
  navBar(s, 'narrow');

  text(s, [
    { text: 'BUILDING', options: { color: BLUE } },
    { text: ' ', options: { color: INK } },
    { text: 'CONSTRUCTION', options: { color: ORANGE } },
  ], { x: 1.795, y: 1.159, w: 9.262, h: 0.909, fontSize: 48, bold: true, fontFace: HEAD, wrap: false });
  text(s, 'Industry Presentation Template', { x: 1.821, y: 2.104, w: 4.221, h: 0.404, valign: 'top' });

  photo(s, 1.949, 2.947, 5.505, 4.553);
  photo(s, 8.234, 2.508, 2.187, 2.071);

  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod', 8.144, 5.059, 3.362, 0.601);
  box(s, 'rect', { x: 8.234, y: 6.095, w: 1.533, h: 0.453, fill: { color: ORANGE }, line: { type: 'none' } });
  text(s, 'Get Started', { x: 8.234, y: 6.095, w: 1.533, h: 0.453, fontSize: 12, color: WHITE, align: 'center', valign: 'middle' });
}

/* ================================================================== slide 02 */
function slide02(pptx) {
  const s = pptx.addSlide();
  navBar(s, 'wide');
  [
    [1.155, 1.389, 2.375, 1.589], [3.839, 1.389, 3.96, 2.774],
    [1.155, 3.316, 2.375, 2.774], [3.894, 4.5, 3.906, 1.589],
  ].forEach(p => photo(s, p[0], p[1], p[2], p[3]));

  pageTitle(s, 'Welcome To ', 'Bulding Construction', 8.945, 1.232, 3.707, 1.969);
  subHead(s, 'About Our Planning', 8.945, 3.658, 2.604, BLUE);
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor '
    + 'PLACEHOLDER'
    + 'adipiscing elit, sed do eiusmod tempor incididunt eiusmod tempor', 8.945, 4.519, 3.35, 1.662);
}

/* ================================================================== slide 03 */
function slide03(pptx) {
  const s = pptx.addSlide();
  box(s, 'rect', { x: 6.158, y: 1.384, w: 5.54, h: 4.732, fill: { color: BLUE }, line: { type: 'none' } });
  navBar(s, 'wide');
  pageTitle(s, 'Vision & ', 'Mission', 1.053, 1.222, 4.789, 0.724);

  const LOREM = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor';
  subHead(s, '1.  Vision Title', 6.901, 1.999, 1.937, WHITE);
  para(s, LOREM, 6.933, 2.608, 2.457, 0.867, { color: WHITE });
  subHead(s, '2. Mission Title', 6.884, 4.025, 2.101, WHITE);
  para(s, LOREM, 6.915, 4.635, 2.457, 0.867, { color: WHITE });

  photo(s, 10.116, 2.194, 3.218, 3.133);
  photo(s, 1.167, 2.421, 3.936, 3.695);
}

/* ================================================================== slide 04 */
function slide04(pptx) {
  const s = pptx.addSlide();
  photo(s, 1.141, 1.389, 2.399, 1.605);
  photo(s, 4.305, 0, 3.171, 6.105);
  navBar(s, 'wide');
  pageTitle(s, 'Best ', 'Facilities', 8.095, 1.253, 3.674, 0.724);
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut',
    8.131, 2.393, 4.164, 0.601);

  const CARD = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt';
  [
    { x: 1.141, tx: 1.757, fill: BLUE, label: '1.  Facilities', lw: 1.666 },
    { x: 8.242, tx: 8.858, fill: ORANGE, label: '2.  Facilities', lw: 1.708 },
  ].forEach(c => {
    box(s, 'rect', { x: c.x, y: 3.65, w: 3.954, h: 2.455, fill: { color: c.fill }, line: { type: 'none' } });
    subHead(s, c.label, c.tx, 4.152, c.lw, WHITE);
    para(s, CARD, c.tx, 4.746, 2.89, 0.867, { color: WHITE });
  });
}

/* ================================================================== slide 05 */
function slide05(pptx) {
  const s = pptx.addSlide();
  navBar(s, 'wide');
  pageTitle(s, 'Best Our ', 'Project', 1.053, 1.222, 4.789, 0.724);
  photo(s, 0, 2.426, 5.095, 3.686);
  photo(s, 10.614, 1.383, 2.719, 4.728);

  const COPY = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt';
  [
    { y: 1.383, ty: 2.048, fill: BLUE, label: '1. Product Title' },
    { y: 3.136, ty: 3.801, fill: ORANGE, label: '2. Product Title' },
    { y: 4.89, ty: 5.555, fill: BLUE, label: '3. Product Title' },
  ].forEach(r => {
    box(s, 'rect', { x: 5.886, y: r.y, w: 2.403, h: 0.494, fill: { color: r.fill }, line: { type: 'none' } });
    text(s, r.label, { x: 5.886, y: r.y, w: 2.403, h: 0.494, fontSize: 18, bold: true, color: WHITE, align: 'center', valign: 'middle' });
    para(s, COPY, 5.8, r.ty, 4.023, 0.601);
  });
}

/* ================================================================== slide 06 */
function slide06(pptx) {
  const s = pptx.addSlide();
  navBar(s, 'wide');
  pageTitle(s, 'Meet Our ', 'Team', 4.604, 1.222, 4.126, 0.724, 'center');
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut adipiscing elit, sed',
    { x: 4.252, y: 2.397, w: 4.83, h: 0.601, fontSize: 10.5, align: 'center', lineSpacingMultiple: 1.5, valign: 'top' });

  [
    { x: 1.168, py: 2.188, ph: 3.76, fill: BLUE, name: 'Alfredo Torres' },
    { x: 3.998, py: 3.75, ph: 2.197, fill: ORANGE, name: 'Bailey Dupont' },
    { x: 6.828, py: 3.75, ph: 2.197, fill: BLUE, name: 'Benjamin Shah' },
    { x: 9.658, py: 2.188, ph: 3.76, fill: ORANGE, name: 'Dani Martinez' },
  ].forEach(m => {
    photo(s, m.x, m.py, 2.507, m.ph);
    box(s, 'rect', { x: m.x, y: 6.137, w: 2.507, h: 0.554, fill: { color: m.fill }, line: { type: 'none' } });
    text(s, m.name, { x: m.x, y: 6.137, w: 2.507, h: 0.554, fontSize: 16, bold: true, color: WHITE, align: 'center', valign: 'middle' });
  });
}

/* ================================================================== slide 07 */
function slide07(pptx) {
  const s = pptx.addSlide();
  photo(s, 6.667, 2.426, 6.667, 5.074);
  navBar(s, 'wide');
  pageTitle(s, 'Best Our ', 'Service', 1.053, 1.222, 4.789, 0.724);

  box(s, 'rect', { x: 1.155, y: 2.426, w: 4.774, h: 5.074, fill: { color: BLUE }, line: { type: 'none' } });
  const COPY = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut eiusmod';
  icon(s, 'bars', 2.049, 3.222, 0.322, WHITE, BLUE);
  subHead(s, '1. Service', 2.464, 3.188, 1.378, WHITE);
  para(s, COPY, 1.971, 3.785, 3.164, 0.867, { color: WHITE });
  icon(s, 'clipboard', 2.016, 5.395, 0.3, WHITE, BLUE);
  subHead(s, '2. Service', 2.464, 5.376, 1.42, WHITE);
  para(s, COPY, 1.95, 5.967, 3.164, 0.867, { color: WHITE });

  box(s, 'rect', { x: 7.421, y: 1.4, w: 4.757, h: 1.868, fill: { color: ORANGE }, line: { type: 'none' } });
  para(s, 'Lorem ipsum dolor sit amet, sed do eiusmod tempor incididunt ut eiusmod consectetur '
    + 'adipiscing elit, sed do eiusmod tempor incididunt ut eiusmodsmod adipiscing elit',
    7.869, 1.768, 3.861, 1.132, { color: WHITE });
}

/* ================================================================== slide 08 */
function slide08(pptx) {
  const s = pptx.addSlide();
  navBar(s, 'wide');
  photo(s, 1.158, 1.389, 2.788, 4.784);
  photo(s, 4.695, 1.389, 2.788, 4.784);
  photo(s, 8.232, 1.389, 3.947, 2.627);
  pageTitle(s, 'Best ', 'Portfolio', 8.095, 4.432, 3.674, 0.724);
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut',
    8.131, 5.572, 4.164, 0.601);
}

/* ================================================================== slide 09 */
function slide09(pptx) {
  const s = pptx.addSlide();
  navBar(s, 'wide');
  pageTitle(s, 'Infographic ', 'Here', 1.053, 1.222, 3.331, 1.346);
  const BLURB = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttmassa. '
    + 'Fusce posuere, magna sed pulvinar osuere, magna magna posuere, magna sed';
  para(s, BLURB, 1.07, 3.021, 3.422, 1.212, { fontSize: 11 });
  para(s, BLURB, 1.07, 4.691, 3.422, 1.212, { fontSize: 11 });

  // Regular (orange) plan, behind the highlighted card.
  box(s, 'rect', { x: 9.451, y: 1.894, w: 3.054, h: 3.79, fill: { color: ORANGE }, line: { type: 'none' } });
  pricePlan(s, {
    tab: [9.791, 2.266, 2.436, 0.571], tabText: 'Regular Class', tabSize: 18,
    price: [9.816, 2.96, 2.386, 0.505], amount: 'FREE ', amountSize: 24,
    bullets: [3.702, 4.132], bulletX: 10.316, dotX: 10.199, dotYOffset: 0.09,
    cta: [10.062, 4.718, 1.832, 0.602], ctaSize: 12,
  });

  // Premium (blue) plan.
  box(s, 'rect', { x: 5.703, y: 1.417, w: 3.748, h: 4.7, fill: { color: BLUE }, line: { type: 'none' } });
  pricePlan(s, {
    tab: [6.118, 1.803, 2.937, 0.674], tabText: 'Premium Class', tabSize: 18,
    price: [6.31, 2.58, 2.386, 0.572], amount: '$500 ', amountSize: 28,
    bullets: [3.442, 3.906, 4.371], bulletX: 6.819, dotX: 6.702, dotYOffset: 0.105,
    cta: [6.335, 5.037, 2.515, 0.726], ctaSize: 16,
  });
}

// One pricing column: translucent tab, price line, bulleted services and CTA button.
function pricePlan(s, p) {
  box(s, 'rect', { x: p.tab[0], y: p.tab[1], w: p.tab[2], h: p.tab[3], fill: { color: WHITE, transparency: 90 }, line: { type: 'none' } });
  text(s, p.tabText, { x: p.tab[0], y: p.tab[1], w: p.tab[2], h: p.tab[3], fontSize: p.tabSize, bold: true, color: WHITE, align: 'center', valign: 'middle' });

  text(s, [
    { text: p.amount, options: { fontSize: p.amountSize } },
    { text: '/Month', options: { fontSize: 16 } },
  ], { x: p.price[0], y: p.price[1], w: p.price[2], h: p.price[3], bold: true, color: WHITE, align: 'center', valign: 'top' });

  p.bullets.forEach((y, i) => {
    box(s, 'ellipse', { x: p.dotX, y: y + p.dotYOffset, w: 0.097, h: 0.097, fill: { color: WHITE }, line: { type: 'none' } });
    text(s, 'Your Service Title 0' + (i + 1), { x: p.bulletX, y: y, w: 1.886, h: 0.278, fontSize: 10.5, color: WHITE, valign: 'top' });
  });

  box(s, 'rect', { x: p.cta[0], y: p.cta[1], w: p.cta[2], h: p.cta[3], fill: { color: WHITE, transparency: 90 }, line: { type: 'none' } });
  text(s, 'Register Now', { x: p.cta[0], y: p.cta[1], w: p.cta[2], h: p.cta[3], fontSize: p.ctaSize, bold: true, color: WHITE, align: 'center', valign: 'middle' });
}

/* ================================================================== slide 10 */
function slide10(pptx) {
  const s = pptx.addSlide();
  navBar(s, 'wide');
  pageTitle(s, 'Infographic ', 'Here', 3.357, 1.222, 6.619, 0.724, 'center');
  text(s, 'Lorem ipsum dolor sit amet, consect dolor sit amet, consect adipiscing elit. dolor dolor sit amet, '
    + 'consect dolor sit amet, consect adipiscing elit. dolor dolor sit amet, consect dolor sit amet, consect adipiscing elit. ',
    { x: 2.208, y: 2.747, w: 8.917, h: 0.603, fontSize: 10.5, align: 'center', lineSpacingMultiple: 1.5, valign: 'top' });

  [
    { x: 1.275, y: 3.944, pct: '44%', bar: ORANGE },
    { x: 1.275, y: 5.203, pct: '85%', bar: BLUE },
    { x: 6.844, y: 3.944, pct: '56%', bar: BLUE },
    { x: 6.844, y: 5.203, pct: '74%', bar: ORANGE },
  ].forEach(c => {
    box(s, 'rect', { x: c.x, y: c.y, w: 5.225, h: 0.894, fill: { color: WHITE }, line: { type: 'none' }, shadow: SHADOW_CARD() });
    box(s, 'rect', { x: c.x + 0.194, y: c.y + 0.155, w: 0.115, h: 0.594, fill: { color: c.bar }, line: { type: 'none' } });
    text(s, c.pct, { x: c.x + 0.576, y: c.y + 0.195, w: 1.21, h: 0.505, fontSize: 24, bold: true, valign: 'top' });
    para(s, 'Lorem ipsum dolor sit amet, consectetuer massa. Fusce posuere, ', c.x + 1.744, c.y + 0.123, 3.134, 0.601);
  });
}

/* ================================================================== slide 11 */
function slide11(pptx) {
  const s = pptx.addSlide();
  navBar(s, 'wide');
  pageTitle(s, 'Infographic ', 'Here', 3.357, 1.222, 6.619, 0.724, 'center');

  // Pinwheel of four rotated "flowchart data" parallelograms.
  box(s, 'flowChartInputOutput', { x: 2.078, y: 2.965, w: 1.727, h: 1.121, rotate: 270, fill: { color: LILAC }, line: { type: 'none' } });
  box(s, 'flowChartInputOutput', { x: 1.184, y: 3.829, w: 1.787, h: 1.121, rotate: 180, fill: { color: BLUE }, line: { type: 'none' } });
  box(s, 'flowChartInputOutput', { x: 2.911, y: 3.829, w: 1.787, h: 1.121, rotate: 180, flipH: true, flipV: true, fill: { color: BLUE }, line: { type: 'none' } });
  box(s, 'flowChartInputOutput', { x: 2.078, y: 4.692, w: 1.727, h: 1.121, rotate: 270, flipH: true, flipV: true, fill: { color: PEACH }, line: { type: 'none' } });
  icon(s, 'badge', 1.835, 4.166, 0.4, WHITE, BLUE);
  icon(s, 'person', 2.678, 3.27, 0.44, WHITE, LILAC);
  icon(s, 'chartbox', 3.716, 4.203, 0.372, WHITE, BLUE);
  icon(s, 'linechart', 2.723, 5.157, 0.4, WHITE, PEACH);

  const COPY = 'Lorem ipsum dolor sit amet, consectetuer adi piscing elit. ipsum adi piscing elit. ';
  [
    { x: 5.968, tx: 6.424, y: 2.666, color: LILAC },
    { x: 9.544, tx: 10.0, y: 2.666, color: PEACH },
    { x: 5.968, tx: 6.424, y: 4.846, color: BLUE },
    { x: 9.544, tx: 10.0, y: 4.846, color: ORANGE },
  ].forEach(c => {
    box(s, 'rect', { x: c.x, y: c.y, w: 0.05, h: 1.012, fill: { color: c.color }, line: { type: 'none' } });
    text(s, 'Title Here', { x: c.tx, y: c.y, w: 1.653, h: 0.303, fontSize: 12, bold: true, color: c.color, valign: 'top' });
    para(s, COPY, c.tx, c.y + 0.358, 2.333, 1.182, { fontSize: 11 });
  });
}

/* ================================================================== slide 12 */
function slide12(pptx) {
  const s = pptx.addSlide();
  navBar(s, 'wide');
  pageTitle(s, 'Infographic ', 'Here', 3.357, 1.222, 6.619, 0.724, 'center');

  // Four slanted quads forming the central counter block.
  quad(s, 4.443, 2.91, 2.483, 1.29, ORANGE, [[0, 0], [1, 0], [0.877, 1], [0, 1]]);
  quad(s, 6.72, 2.91, 2.179, 1.29, BLUE, [[0.14, 0], [1, 0], [1, 1], [0, 1]]);
  quad(s, 4.443, 4.304, 2.179, 1.29, BLUE, [[0, 0], [1, 0], [0.86, 1], [0, 1]]);
  quad(s, 6.415, 4.304, 2.483, 1.29, ORANGE, [[0.123, 0], [1, 0], [1, 1], [0, 1]]);

  [['284+', 5.203], ['843+', 6.926]].forEach(n => {
    text(s, n[0], { x: n[1], y: 3.205, w: 1.473, h: 0.64, fontSize: 32, bold: true, color: WHITE, align: 'center', valign: 'top' });
  });
  [['632+', 5.068], ['195+', 6.792]].forEach(n => {
    text(s, n[0], { x: n[1], y: 4.63, w: 1.473, h: 0.64, fontSize: 32, bold: true, color: WHITE, align: 'center', valign: 'top' });
  });

  // White "manual input" tabs carrying the four icons; the slanted edge always
  // faces away from the centre, so the outer pair is mirrored.
  [
    { x: 3.457, y: 2.403, mirror: false }, { x: 3.457, y: 4.541, mirror: false },
    { x: 8.677, y: 2.451, mirror: true }, { x: 8.677, y: 4.589, mirror: true },
  ].forEach(t => {
    box(s, 'flowChartManualInput', {
      x: t.x, y: t.y, w: 1.178, h: 1.675, rotate: 270, flipV: t.mirror,
      fill: { color: WHITE }, line: { type: 'none' }, shadow: SHADOW_CARD(),
    });
  });
  icon(s, 'cap', 3.909, 3.053, 0.55, BLUE);
  icon(s, 'flask', 3.94, 5.1, 0.55, ORANGE);
  icon(s, 'gear', 8.92, 3.062, 0.55, ORANGE);
  icon(s, 'atom', 8.848, 5.128, 0.55, BLUE);

  const NOTE = 'Lorem ipsum dolor sit amet, adipiscing elit. cenas';
  [
    { x: 1.254, lx: 1.456, y: 2.664, label: '01. Title Here', color: BLUE, align: 'right' },
    { x: 1.254, lx: 1.456, y: 4.843, label: '02. Title Here', color: ORANGE, align: 'right' },
    { x: 10.606, lx: 10.606, y: 2.664, label: '03. Title Here', color: ORANGE, align: 'left' },
    { x: 10.606, lx: 10.606, y: 4.843, label: '04. Title Here', color: BLUE, align: 'left' },
  ].forEach(c => {
    text(s, c.label, { x: c.lx, y: c.y, w: 1.473, h: 0.31, fontSize: 12, bold: true, color: c.color, align: c.align, valign: 'top' });
    text(s, NOTE, { x: c.x, y: c.y + 0.309, w: 1.675, h: 0.904, fontSize: 11, align: c.align, lineSpacingMultiple: 1.5, valign: 'top' });
  });
}

/* ================================================================== slide 13 */
function slide13(pptx) {
  const s = pptx.addSlide();
  navBar(s, 'wide');
  pageTitle(s, 'Infographic ', 'Here', 3.357, 1.222, 6.619, 0.724, 'center');

  // Three full-width guide rails.
  [2.952, 3.806, 4.699].forEach(y => {
    box(s, 'rect', { x: 0, y: y, w: 13.333, h: 0.131, fill: { color: BLUE }, line: { type: 'none' } });
  });

  // Rotated cubes reading as 3-D blocks sitting on the rails.
  [
    { x: 2.619, y: 2.606, w: 0.814, h: 2.531, fill: NAVY, pct: '30%', px: 2.392, py: 3.587, pw: 1.268 },
    { x: 7.202, y: 1.711, w: 0.814, h: 2.531, fill: LILAC, pct: '32%', px: 6.842, py: 2.647, pw: 1.268 },
    { x: 5.758, y: 3.639, w: 0.831, h: 2.271, fill: ORANGE, pct: '20%', px: 5.66, py: 4.48, pw: 1.028 },
    { x: 10.304, y: 2.944, w: 0.814, h: 1.855, fill: BLUE, pct: '18%', px: 10.077, py: 3.587, pw: 1.268 },
  ].forEach(c => {
    box(s, 'cube', { x: c.x, y: c.y, w: c.w, h: c.h, rotate: 90, fill: { color: c.fill }, line: { type: 'none' } });
    text(s, c.pct, { x: c.px, y: c.py, w: c.pw, h: 0.438, fontSize: 20, bold: true, color: WHITE, align: 'center', valign: 'top' });
  });

  [
    { ix: 1.223, tx: 1.713, color: LILAC, kind: 'puzzle' },
    { ix: 4.151, tx: 4.651, color: BLUE, kind: 'gear' },
    { ix: 7.09, tx: 7.566, color: NAVY, kind: 'pie' },
    { ix: 10.005, tx: 10.415, color: ORANGE, kind: 'bulb' },
  ].forEach(l => {
    icon(s, l.kind, l.ix, 5.6, 0.38, l.color);
    text(s, 'Your Title Here', { x: l.tx, y: 5.609, w: 2.271, h: 0.37, fontSize: 16, bold: true, color: l.color, valign: 'top' });
  });
}

/* ================================================================== slide 14 */
function slide14(pptx) {
  const s = pptx.addSlide();
  navBar(s, 'wide');
  pageTitle(s, 'Infographic ', 'Here', 3.357, 1.222, 6.619, 0.724, 'center');

  // Ribbon of four segments; alternating segments point up / down.
  [
    { x: 2.252, fill: LILAC, up: true, num: '01', numX: 3.295, numY: 2.757, boxY: 2.569, noteX: 2.39, noteY: 4.569, kind: 'briefcase' },
    { x: 4.46, fill: NAVY, up: false, num: '02', numX: 5.529, numY: 5.272, boxY: 5.054, noteX: 4.598, noteY: 3.413, kind: 'globe' },
    { x: 6.668, fill: DEEP, up: true, num: '03', numX: 7.711, numY: 2.757, boxY: 2.569, noteX: 6.806, noteY: 4.569, kind: 'safe' },
    { x: 8.876, fill: ORANGE, up: false, num: '04', numX: 9.944, numY: 5.272, boxY: 5.054, noteX: 9.014, noteY: 3.413, kind: 'pin' },
  ].forEach(seg => {
    box(s, 'rect', { x: seg.x, y: 4.185, w: 2.205, h: 0.335, fill: { color: seg.fill }, line: { type: 'none' }, shadow: SHADOW_BAR() });
    box(s, 'triangle', {
      x: seg.x, y: seg.up ? 3.755 : 4.615, w: 2.205, h: 0.335,
      flipV: !seg.up, fill: { color: seg.fill }, line: { type: 'none' },
    });
    box(s, 'rect', { x: seg.x, y: seg.boxY, w: 2.205, h: 1.081, fill: { type: 'none' }, line: { color: seg.fill, width: 1 } });
    text(s, seg.num, { x: seg.numX, y: seg.numY, w: 0.998, h: 0.707, fontSize: 36, bold: true, color: seg.fill, valign: 'top' });
    icon(s, seg.kind, seg.numX - 0.63, seg.numY + 0.13, 0.44, seg.fill);
    para(s, 'Lorem ipsum dolor sit amet, adipiscing', seg.noteX, seg.noteY, 1.903, 0.627, { fontSize: 11 });
  });
}

/* ================================================================== slide 15 */
function slide15(pptx) {
  const s = pptx.addSlide();
  navBar(s, 'wide');
  pageTitle(s, 'Infographic ', 'Here', 3.357, 1.222, 6.619, 0.724, 'center');

  // Left column: four ticked mini-facts plus a closing paragraph.
  [
    { x: 1.176, tx: 1.522, y: 2.682, ty: 2.618, label: '01 Title Here', color: ORANGE },
    { x: 3.907, tx: 4.253, y: 2.685, ty: 2.621, label: '02 Title Here', color: BLUE },
    { x: 1.176, tx: 1.522, y: 3.762, ty: 3.698, label: '03 Title Here', color: ORANGE },
    { x: 3.907, tx: 4.253, y: 3.765, ty: 3.701, label: '04 Title Here', color: BLUE },
  ].forEach(f => {
    checkDot(s, f.x, f.y, 0.276, f.color);
    text(s, f.label, { x: f.tx, y: f.ty, w: 1.498, h: 0.303, fontSize: 12, bold: true, valign: 'top' });
    para(s, 'Lorem ipsum dolor sit.', f.tx, f.ty + 0.236, 1.81, 0.601);
  });
  para(s, 'Lorem ipsum dolor sit amet, consect dolor sit amet, consect adipiscing elit. dolor dolor sit amet, '
    + 'consect dolor sit dolor sit amet, consect dolor sit amet, consect adipiscing elit. dolor dolor sit amet, '
    + 'dolor sit amet, consect dolor sit amet. Lorem ipsum dolor sit amet, consect dolor sit amet, consect adipiscing elit. ',
    1.086, 4.777, 4.889, 1.398);

  // Right: funnel of stacked trapezoids with a heading plate on top.
  box(s, 'rect', { x: 8.548, y: 2.648, w: 2.026, h: 3.46, fill: { color: NAVY }, line: { type: 'none' } });
  [
    { x: 7.144, w: 4.835, y: 3.18, fill: ORANGE, label: '01 Title Here', ly: 3.359 },
    { x: 7.493, w: 4.136, y: 3.936, fill: BLUE, label: '02 Title Here', ly: 4.114 },
    { x: 7.835, w: 3.452, y: 4.691, fill: ORANGE, label: '03 Title Here', ly: 4.87 },
    { x: 8.179, w: 2.764, y: 5.447, fill: BLUE, label: '04 Title Here', ly: 5.626 },
  ].forEach(t => {
    trap(s, t.x, t.y, t.w, 0.66, t.fill, 0.336 / t.w, true);
    text(s, t.label, { x: 8.494, y: t.ly, w: 2.133, h: 0.303, fontSize: 12, bold: true, color: WHITE, align: 'center', valign: 'top' });
  });
  text(s, 'Main Heading', { x: 8.685, y: 2.738, w: 1.753, h: 0.337, fontSize: 14, bold: true, color: WHITE, align: 'center', valign: 'top' });

  // Two diagonal class arrows flanking the funnel.
  box(s, 'rect', { x: 6.823, y: 3.193, w: 0.186, h: 0.826, rotate: 335.4, fill: { color: BLUE }, line: { type: 'none' } });
  box(s, 'downArrow', { x: 7.68, y: 5.168, w: 0.411, h: 0.999, rotate: 335.1, fill: { color: BLUE }, line: { type: 'none' } });
  text(s, 'Regular Class', { x: 6.618, y: 4.453, w: 1.515, h: 0.286, rotate: 64.7, fontSize: 10.5, bold: true, color: WHITE, align: 'center', valign: 'middle', fill: { color: ORANGE } });
  box(s, 'rect', { x: 11.1, y: 5.295, w: 0.186, h: 0.826, rotate: 24.6, flipV: true, fill: { color: BLUE }, line: { type: 'none' } });
  box(s, 'downArrow', { x: 11.957, y: 3.147, w: 0.411, h: 0.999, rotate: 24.9, flipV: true, fill: { color: BLUE }, line: { type: 'none' } });
  text(s, 'Premium Class', { x: 10.938, y: 4.579, w: 1.429, h: 0.278, rotate: 295.3, fontSize: 10.5, bold: true, color: WHITE, align: 'center', valign: 'middle', fill: { color: ORANGE } });
}

/* ================================================================== slide 16 */
function slide16(pptx) {
  const s = pptx.addSlide();
  navBar(s, 'wide');
  pageTitle(s, 'Infographic ', 'Here', 3.357, 1.222, 6.619, 0.724, 'center');

  // Two-tone "process" block plus three chevron pairs (top light / bottom dark).
  quad(s, 1.185, 2.983, 3.759, 1.398, ORANGE,
    [[0, 0], [0.737, 0], [0.737, 0.428], [0.868, 0.428], [1, 0.99], [0.997, 1], [0, 1]]);
  quad(s, 1.185, 4.381, 3.751, 1.371, RUST,
    [[0, 0], [1, 0], [0.87, 0.564], [0.738, 0.564], [0.738, 1], [0, 1]]);
  box(s, 'rect', { x: 1.625, y: 3.423, w: 1.888, h: 1.888, fill: { type: 'none' }, line: { color: WHITE, width: 1 } });
  box(s, 'rect', { x: 2.029, y: 3.827, w: 1.08, h: 1.08, fill: { color: WHITE }, line: { type: 'none' }, shadow: SHADOW_SOFT() });
  icon(s, 'megaphone', 2.295, 4.216, 0.53, ORANGE);

  const UP = [[0, 0], [0.676, 0], [1, 0.982], [0.995, 1], [0.318, 1], [0.324, 0.982]];
  const DOWN = [[0.32, 0], [1, 0], [0.68, 1], [0, 1]];
  [
    { x: 4.541, top: BLUE, bot: NAVY, tw: 1.572, bw: 1.563 },
    { x: 5.698, top: ORANGE, bot: RUST, tw: 1.572, bw: 1.563 },
    { x: 6.852, top: BLUE, bot: NAVY, tw: 1.624, bw: 1.619 },
  ].forEach(c => {
    quad(s, c.x, 3.581, c.tw, 0.8, c.top, UP);
    quad(s, c.x, 4.381, c.bw, 0.772, c.bot, DOWN);
  });
  // Battlement detail crowning the last chevron.
  [[7.08, 3.245, 0.211, 0.336], [7.465, 3.371, 0.211, 0.21], [7.85, 3.245, 0.211, 0.336]].forEach(r => {
    box(s, 'rect', { x: r[0], y: r[1], w: r[2], h: r[3], fill: { color: BLUE }, line: { type: 'none' } });
  });
  [[7.158, 3.111], [7.804, 3.111]].forEach(r => {
    box(s, 'rect', { x: r[0], y: r[1], w: 0.179, h: 0.335, rotate: 90, fill: { color: BLUE }, line: { type: 'none' } });
  });

  text(s, 'Title Here', { x: 4.465, y: 2.809, w: 1.191, h: 0.303, fontSize: 12, bold: true, color: BLUE, align: 'center', valign: 'top' });
  text(s, 'Lorem ipsum dolor', { x: 4.262, y: 3.028, w: 1.599, h: 0.338, fontSize: 10.5, italic: true, align: 'center', lineSpacingMultiple: 1.5, valign: 'top' });
  text(s, 'Title Here', { x: 5.722, y: 5.376, w: 1.191, h: 0.303, fontSize: 12, bold: true, color: ORANGE, align: 'center', valign: 'top' });
  text(s, 'Lorem ipsum dolor', { x: 5.518, y: 5.595, w: 1.599, h: 0.338, fontSize: 10.5, italic: true, align: 'center', lineSpacingMultiple: 1.5, valign: 'top' });

  [{ y: 2.746, color: BLUE }, { y: 3.937, color: NAVY }, { y: 5.127, color: DEEP }].forEach(r => {
    checkDot(s, 9.267, r.y + 0.09, 0.239, r.color);
    para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttmassa. ', 9.607, r.y, 2.669, 0.868);
  });
}

/* ================================================================== slide 17 */
function slide17(pptx) {
  const s = pptx.addSlide();
  navBar(s, 'wide');
  pageTitle(s, 'Infographic ', 'Here', 3.357, 1.222, 6.619, 0.724, 'center');

  // Four stacked panels of decreasing width, largest on the left.
  [
    { x: 1.172, w: 5.494, fill: ORANGE, num: '4', nx: 1.049, nw: 1.0, icon: 6.184 },
    { x: 6.667, w: 2.539, fill: NAVY, num: '3', nx: 6.522, nw: 0.933, icon: 8.719 },
    { x: 9.205, w: 1.892, fill: BLUE, num: '2', nx: 9.058, nw: 0.938, icon: 10.613 },
    { x: 11.097, w: 1.064, fill: LILAC, num: '1', nx: 10.924, nw: 0.733, icon: 11.68 },
  ].forEach(p => {
    box(s, 'rect', { x: p.x, y: 3.022, w: p.w, h: 1.875, flipH: true, fill: { color: p.fill }, line: { type: 'none' }, shadow: SHADOW_BAR() });
    text(s, p.num, { x: p.nx, y: 3.69, w: p.nw, h: 1.582, fontSize: 88, bold: true, color: WHITE, wrap: false });
    icon(s, 'notebook', p.icon, 3.147, 0.36, WHITE, p.fill);
  });

  [
    { x: 1.166, tx: 1.386, color: ORANGE },
    { x: 4.221, tx: 4.441, color: NAVY },
    { x: 7.276, tx: 7.496, color: BLUE },
    { x: 10.332, tx: 10.552, color: LILAC },
  ].forEach(l => {
    box(s, 'rect', { x: l.x, y: 5.502, w: 0.166, h: 0.166, fill: { color: l.color }, line: { type: 'none' } });
    text(s, 'Title Here', { x: l.tx, y: 5.407, w: 1.498, h: 0.303, fontSize: 12, bold: true, valign: 'top' });
    para(s, 'Lorem ipsum dolor sit.', l.tx, 5.643, 1.81, 0.601);
  });
}

/* ================================================================== slide 18 */
function slide18(pptx) {
  const s = pptx.addSlide();
  navBar(s, 'wide');
  pageTitle(s, 'Infographic ', 'Here', 3.357, 1.222, 6.619, 0.724, 'center');

  // Three overlapping house-shaped arrows, one per year.
  [
    { x: 1.083, y: 2.648, w: 3.647, h: 3.454, head: 0.322, shaft: 0.529, fill: BLUE, year: '2024', yx: 1.876, yw: 0.83 },
    { x: 2.689, y: 3.144, w: 3.123, h: 2.958, head: 0.337, shaft: 0.474, fill: NAVY, year: '2025', yx: 3.405, yw: 0.814 },
    { x: 3.984, y: 3.517, w: 2.729, h: 2.585, head: 0.429, shaft: 0.474, fill: ORANGE, year: '2026', yx: 4.588, yw: 0.823 },
  ].forEach(a => {
    blockArrow(s, a.x, a.y, a.w, a.h, a.fill, a.head, a.shaft, 'up');
    text(s, a.year, { x: a.yx, y: 5.736, w: a.yw, h: 0.404, fontSize: 18, bold: true, color: WHITE, wrap: false });
  });

  // Progress tracks on the right.
  [
    { y: 2.935, ty: 2.52, label: 'Your Title Here 01', pct: '60%', w: 2.395, knob: 9.856, fill: BLUE, track: 1.043 },
    { y: 3.81, ty: 3.395, label: 'Your Title Here 02', pct: '72%', w: 3.101, knob: 10.58, fill: NAVY, track: 1.918 },
    { y: 4.685, ty: 4.27, label: 'Your Title Here 03', pct: '93%', w: 3.525, knob: 11.007, fill: ORANGE, track: 2.792 },
  ].forEach(b => {
    box(s, 'rect', { x: 9.484, y: b.track, w: 0.102, h: 3.919, rotate: 90, fill: { color: 'D9D9D9' }, line: { type: 'none' } });
    box(s, 'rect', { x: 7.575, y: b.y, w: b.w, h: 0.118, fill: { color: b.fill }, line: { type: 'none' } });
    box(s, 'rect', { x: b.knob, y: b.y - 0.045, w: 0.197, h: 0.197, fill: { color: b.fill }, line: { type: 'none' } });
    text(s, b.label, { x: 7.499, y: b.ty, w: 2.596, h: 0.345, fontSize: 12, valign: 'top' });
    text(s, b.pct, { x: 11.736, y: b.y - 0.11, w: 0.704, h: 0.321, fontSize: 11, valign: 'top' });
  });

  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt '
    + 'ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation',
    7.462, 5.323, 4.789, 0.868);
}

/* ================================================================== slide 19 */
function slide19(pptx) {
  const s = pptx.addSlide();
  navBar(s, 'wide');
  pageTitle(s, 'Contact ', 'Us', 1.053, 1.222, 3.331, 0.724);
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod', 1.06, 2.425, 3.298, 0.601);

  box(s, 'rect', { x: 5.081, y: 1.405, w: 7.088, h: 2.345, fill: { color: BLUE }, line: { type: 'none' } });
  [
    { ix: 5.993, iy: 2.089, tx: 6.5, ty: 2.107, w: 1.633, value: '+12345678910', kind: 'phone' },
    { ix: 5.969, iy: 2.883, tx: 6.5, ty: 2.826, w: 1.953, value: 'yourmail@gmail.com', kind: 'mail' },
    { ix: 8.631, iy: 2.118, tx: 9.187, ty: 2.128, w: 1.953, value: 'www.yourwebsite.com', kind: 'globe' },
    { ix: 8.612, iy: 2.815, tx: 9.187, ty: 2.826, w: 2.259, value: '123 Anywhere ST,, Any City', kind: 'house' },
  ].forEach(row => {
    icon(s, row.kind, row.ix, row.iy, 0.31, WHITE, BLUE);
    text(s, row.value, { x: row.tx, y: row.ty, w: row.w, h: 0.278, fontSize: 10.5, color: WHITE, valign: 'top' });
  });

  photo(s, 1.155, 3.75, 3.136, 3.144);
  photo(s, 5.099, 4.568, 7.07, 2.325);
}

/* ================================================================== slide 20 */
function slide20(pptx) {
  const s = pptx.addSlide();
  box(s, 'rect', { x: 1.179, y: 0, w: 11.011, h: 7.5, fill: { color: PANEL }, line: { type: 'none' } });
  photo(s, 7.444, 3.076, 3.981, 3.05);
  photo(s, 1.179, 3.75, 5.502, 2.376);
  navBar(s, 'narrow');

  text(s, [
    { text: 'THANK', options: { color: BLUE } },
    { text: ' ', options: { color: WHITE } },
    { text: 'YOU', options: { color: ORANGE } },
  ], { x: 1.832, y: 0.891, w: 9.935, h: 2.036, fontSize: 114, bold: true, fontFace: HEAD, wrap: false });
  text(s, 'FOR WATCHING US', { x: 1.849, y: 2.843, w: 2.635, h: 0.404, fontSize: 18, bold: true, fontFace: HEAD, wrap: false });

  box(s, 'rect', { x: 3.811, y: 5.587, w: 2.87, h: 0.539, fill: { color: ORANGE }, line: { type: 'none' } });
  text(s, 'SEE YOU NEXT TIME', { x: 3.811, y: 5.587, w: 2.87, h: 0.539, fontSize: 16, bold: true, color: WHITE, align: 'center', valign: 'middle' });
}

/* ===================================================================== build */
function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK16x9', width: 13.333, height: 7.5 });
  pptx.layout = 'DECK16x9';
  pptx.title = 'Building Construction';
  pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach(fn => fn(pptx));

  return pptx.writeFile({
    fileName: path.join(__dirname, '0f2608e9-dba0-475f-80ad-08e141856cbe_grok_final.pptx'),
  });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
