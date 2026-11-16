/**
 * Skin Beauty — 20-slide deck rebuilt with pptxgenjs.
 * Raster photos in the source deck are replaced by grey shape placeholders.
 */
const PptxGenJS = require('pptxgenjs');
const path = require('path');

// ---------------------------------------------------------------- palette ---
const C = {
  bg: 'FFEEE9',      // page background (theme accent6)
  brown: 'C37960',    // theme accent2
  orange: 'FE9737',   // theme accent1
  orangeDk: 'E77001', // accent1, 25% shade — used for headline accents
  peach: 'FEC187',
  peachLt: 'FFD5AF',
  rose: 'DBAFA0',
  roseLt: 'E7C9BF',
  rust: '9E553C',
  cream: 'FFEAD7',
  grey: 'F2F2F2',
  dark: '262626',
  white: 'FFFFFF',
  img: 'CCCCCC'       // photo placeholder
};
const F = { serif: 'Playfair Display', sans: 'Open Sans', alt: 'Work Sans' };

const pptx = new PptxGenJS();
const S = pptx.ShapeType;

// ---------------------------------------------------------------- helpers ---
const KAPPA = 0.5523; // circle-to-bezier constant

/** Points for a rectangle with per-corner radii: [tl, tr, br, bl]; each r or [rx, ry]. */
function roundedRect(w, h, corners) {
  const [tl, tr, br, bl] = corners.map(c => (Array.isArray(c) ? c : [c, c]));
  const p = [{ x: tl[0], y: 0, moveTo: true }];
  p.push({ x: w - tr[0], y: 0 });
  if (tr[0]) p.push({ x: w, y: tr[1], curve: { type: 'cubic', x1: w - tr[0] * (1 - KAPPA), y1: 0, x2: w, y2: tr[1] * (1 - KAPPA) } });
  p.push({ x: w, y: h - br[1] });
  if (br[0]) p.push({ x: w - br[0], y: h, curve: { type: 'cubic', x1: w, y1: h - br[1] * (1 - KAPPA), x2: w - br[0] * (1 - KAPPA), y2: h } });
  p.push({ x: bl[0], y: h });
  if (bl[0]) p.push({ x: 0, y: h - bl[1], curve: { type: 'cubic', x1: bl[0] * (1 - KAPPA), y1: h, x2: 0, y2: h - bl[1] * (1 - KAPPA) } });
  p.push({ x: 0, y: tl[1] });
  if (tl[0]) p.push({ x: tl[0], y: 0, curve: { type: 'cubic', x1: 0, y1: tl[1] * (1 - KAPPA), x2: tl[0] * (1 - KAPPA), y2: 0 } });
  p.push({ close: true });
  return p;
}

/** Points from a normalised outline: [x,y] = line, ['c',x1,y1,x2,y2,x,y] = cubic. */
function outline(w, h, steps) {
  return steps.map((s, i) => (s[0] === 'c'
    ? { x: s[5] * w, y: s[6] * h, curve: { type: 'cubic', x1: s[1] * w, y1: s[2] * h, x2: s[3] * w, y2: s[4] * h } }
    : { x: s[0] * w, y: s[1] * h, moveTo: i === 0 })).concat([{ close: true }]);
}

/** Grey stand-in for a photo, clipped to the same rounded outline as the original. */
function photo(slide, x, y, w, h, corners) {
  slide.addShape(S.custGeom, { x, y, w, h, points: roundedRect(w, h, corners || [0, 0, 0, 0]), fill: { color: C.img } });
  slide.addText('[image]', { x, y, w, h, align: 'center', valign: 'middle', fontFace: F.sans, fontSize: 11, color: C.white });
}

/** Body copy: 10.5pt Open Sans, justified, 1.5 line spacing. */
function body(slide, text, o) {
  slide.addText(text, Object.assign({
    fontFace: F.sans, fontSize: 10.5, color: C.dark, align: 'justify',
    lineSpacingMultiple: 1.5, valign: 'top'
  }, o));
}

/** Small bold label (18pt unless overridden). */
function label(slide, text, o) {
  slide.addText(text, Object.assign({ fontFace: F.sans, fontSize: 18, bold: true, color: C.orangeDk, valign: 'top' }, o));
}

/** Two-tone Playfair headline; `parts` = [[text, color, breakAfter], ...]. */
function headline(slide, parts, o) {
  slide.addText(parts.map(([text, color, breakLine]) => ({ text, options: { color, breakLine } })), Object.assign({
    fontFace: F.serif, fontSize: 40, bold: true, color: C.dark, valign: 'top'
  }, o));
}

// Top navigation bar, repeated on every slide.
const NAV_LINKS = [['Home', 4.249, 0.669], ['About', 5.539, 0.669], ['Service', 6.83, 0.872], ['Contact', 8.324, 0.872]];
function navbar(slide) {
  slide.addText('SKIN BEAUTY', { x: 1.082, y: 0.308, w: 1.342, h: 0.303, fontFace: F.sans, fontSize: 12, bold: true, color: C.dark, valign: 'top' });
  NAV_LINKS.forEach(([text, x, w]) => {
    slide.addText(text, { x, y: 0.317, w, h: 0.286, fontFace: F.sans, fontSize: 11, color: C.dark, valign: 'top' });
  });
  [0.401, 0.445, 0.489].forEach(y => slide.addShape(S.rect, { x: 11.942, y, w: 0.236, h: 0.029, fill: { color: C.dark } }));
  slide.addShape(S.line, { x: 1.151, y: 0.616, w: 11.082, h: 0, line: { color: C.brown } });
}

// "Infographic Section" heading used by the infographic slides (9–18).
function sectionTitle(slide) {
  headline(slide, [['Infographic ', C.dark], ['Section', C.orangeDk], [' ', C.orange]], { x: 2.899, y: 1.237, w: 7.536, h: 0.774, align: 'center' });
}

// Tiny white pictograms standing in for the deck's icon artwork. All parts are
// specified in a 0..1 box so one call can place an icon at any size.
const ICONS = {
  //         shape,        x,    y,    w,    h,   options ({ back } paints in the badge colour)
  person: [[S.pie, 0.10, 0.06, 0.80, 0.46, { angleRange: [180, 360] }], [S.ellipse, 0.10, 0.30, 0.24, 0.22],
           [S.ellipse, 0.66, 0.30, 0.24, 0.22], [S.ellipse, 0.24, 0.06, 0.52, 0.46],
           [S.ellipse, 0.35, 0.24, 0.09, 0.09, { back: true }], [S.ellipse, 0.56, 0.24, 0.09, 0.09, { back: true }],
           [S.pieWedge, 0.00, 0.58, 0.44, 0.42, { rotate: 90 }], [S.pieWedge, 0.56, 0.58, 0.44, 0.42]],
  doc: [[S.snip1Rect, 0.02, 0.02, 0.62, 0.96], [S.rect, 0.10, 0.11, 0.46, 0.78, { back: true }],
        [S.rect, 0.17, 0.26, 0.32, 0.06], [S.rect, 0.17, 0.40, 0.32, 0.06], [S.rect, 0.17, 0.54, 0.32, 0.06],
        [S.rect, 0.60, 0.62, 0.10, 0.36], [S.rect, 0.78, 0.62, 0.10, 0.36], [S.donut, 0.52, 0.34, 0.44, 0.44]],
  bars: [[S.rect, 0.00, 0.00, 0.06, 1.00], [S.rect, 0.00, 0.94, 1.00, 0.06],
         [S.rect, 0.17, 0.63, 0.14, 0.31], [S.rect, 0.38, 0.24, 0.14, 0.70],
         [S.rect, 0.59, 0.44, 0.14, 0.50], [S.rect, 0.80, 0.14, 0.14, 0.80]],
  music: [[S.ellipse, 0.00, 0.68, 0.34, 0.28], [S.ellipse, 0.48, 0.58, 0.34, 0.28],
          [S.rect, 0.28, 0.16, 0.07, 0.66], [S.rect, 0.76, 0.06, 0.07, 0.66],
          [S.rect, 0.28, 0.05, 0.55, 0.17, { rotate: -10 }]],
  gear: [[S.gear9, 0.02, 0.10, 0.40, 0.40], [S.gear9, 0.06, 0.56, 0.28, 0.28], [S.gear9, 0.38, 0.34, 0.60, 0.60]],
  // Three arrows laid along the edges of a triangle to form the recycling loop.
  recycle: [[S.rightArrow, 0.30, 0.02, 0.46, 0.26, { rotate: 45 }], [S.rightArrow, 0.27, 0.62, 0.46, 0.26, { rotate: 180 }],
            [S.rightArrow, 0.00, 0.30, 0.46, 0.26, { rotate: 295 }]],
  phone: [[S.teardrop, 0.02, 0.30, 0.56, 0.56, { rotate: 135 }], [S.rect, 0.02, 0.20, 0.22, 0.28, { rotate: -35 }],
          [S.rect, 0.34, 0.62, 0.28, 0.22, { rotate: -35 }],
          [S.blockArc, 0.34, 0.02, 0.64, 0.64, { angleRange: [280, 40], thickness: 0.22 }],
          [S.blockArc, 0.47, 0.15, 0.38, 0.38, { angleRange: [280, 40], thickness: 0.35 }]],
  mail: [[S.rect, 0.02, 0.22, 0.96, 0.56], [S.triangle, 0.02, 0.22, 0.96, 0.44, { flipV: true, back: true }]],
  globe: [[S.ellipse, 0.02, 0.02, 0.96, 0.96], [S.ellipse, 0.30, 0.02, 0.40, 0.96, { back: true }],
          [S.ellipse, 0.36, 0.02, 0.28, 0.96], [S.rect, 0.02, 0.30, 0.96, 0.05, { back: true }],
          [S.rect, 0.02, 0.63, 0.96, 0.05, { back: true }], [S.rect, 0.46, 0.02, 0.08, 0.96]],
  home: [[S.triangle, 0.00, 0.06, 1.00, 0.48], [S.rect, 0.16, 0.50, 0.68, 0.46], [S.rect, 0.40, 0.66, 0.20, 0.30, { back: true }]]
};
/** Draw `kind` in a box of `sw` x `sh` inches at (x, y); `bg` fills cut-out parts. */
function icon(slide, kind, x, y, sw, sh, bg) {
  ICONS[kind].forEach(([shape, dx, dy, dw, dh, o = {}]) => {
    const opt = { x: x + dx * sw, y: y + dy * sh, w: dw * sw, h: dh * sh, fill: { color: o.back ? bg : C.white } };
    if (o.rotate) opt.rotate = o.rotate;
    if (o.flipV) opt.flipV = true;
    if (o.angleRange) { opt.angleRange = o.angleRange; opt.arcThicknessRatio = o.thickness; }
    slide.addShape(shape, opt);
  });
}

// -------------------------------------------------------------- the deck ---
pptx.defineLayout({ name: 'W16x9', width: 13.333, height: 7.5 });
pptx.layout = 'W16x9';

function newSlide() {
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  return slide;
}

const LOREM_LONG = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris culpa qui officia deserunt mollit anim id est laborum.';
const LOREM_MED = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut mollit anim laborum.';
const LOREM_TITLE = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttmassa. Fusce posuere, magna porttmassa. ';

// 1 — Cover
function slide01() {
  const s = newSlide();
  photo(s, 8.416, 1.366, 3.767, 4.767, [1.883, 1.883, 1.883, 1.883]);
  photo(s, 4.745, 1.366, 3.767, 4.767, [1.883, 1.883, 1.883, 1.883]);
  navbar(s);
  s.addShape(S.roundRect, { x: 1.151, y: 1.366, w: 3.767, h: 4.767, rectRadius: 1.883, fill: { color: C.brown } });
  s.addText('Skin Beauty', { x: 1.79, y: 2.727, w: 2.671, h: 1.919, fontFace: F.serif, fontSize: 54, bold: true, color: C.white, valign: 'top' });
  s.addShape(S.line, { x: 1.954, y: 4.845, w: 1.473, h: 0, line: { color: C.white, endArrowType: 'triangle' } });
  s.addText('SKIN CLINIC', { x: 1.069, y: 6.858, w: 2.7, h: 0.303, fontFace: F.sans, fontSize: 12, color: C.dark, charSpacing: 6, valign: 'top' });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing', { x: 5.071, y: 6.79, w: 4.125, h: 0.338 });
  s.addShape(S.ellipse, { x: 11.844, y: 6.89, w: 0.339, h: 0.339, fill: { color: C.brown } });
  s.addShape(S.custGeom, {
    x: 11.961, y: 6.951, w: 0.145, h: 0.217, fill: { color: C.bg },
    points: outline(0.145, 0.217, [[0.28, 0], [1, 0.5], [0.28, 1], [0, 0.83], [0.45, 0.5], [0, 0.17]])
  });
  s.addText('Learn More', {
    shape: S.roundRect, x: 4.443, y: 4.076, w: 2.041, h: 2.057, rectRadius: 1.02,
    fill: { type: 'none' }, line: { color: C.brown, width: 1 },
    align: 'center', valign: 'middle', fontFace: F.sans, fontSize: 14, bold: true, color: C.dark
  });
}

// 2 — Welcome / About us
function slide02() {
  const s = newSlide();
  navbar(s);
  s.addShape(S.rect, { x: 10.597, y: 3.75, w: 2.736, h: 2.354, fill: { color: C.brown } });
  headline(s, [['Welcome ', C.dark, true], ['To ', C.dark], ['Skin Beauty', C.orangeDk]], { x: 1.047, y: 1.195, w: 4.809, h: 1.447 });
  label(s, 'About us', { x: 1.047, y: 3.221, w: 1.745, h: 0.404 });
  body(s, LOREM_LONG, { x: 1.047, y: 3.914, w: 4.826, h: 1.133 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do enim ad minim veniam, quis nostrud exercitation ullamco laboris culpa qui officia deserunt mollit anim id est laborum.', { x: 1.047, y: 5.329, w: 4.826, h: 0.868 });
  photo(s, 7.442, 1.396, 3.866, 4.709, [[1.933, 2.354], [1.933, 2.354], 0, 0]);
}

// 3 — Vision & Mission
function slide03() {
  const s = newSlide();
  s.addShape(S.rect, { x: 2.251, y: 3.759, w: 11.082, h: 2.338, fill: { color: C.brown } });
  navbar(s);
  [['Our Vision', 1.047], ['Our Mission', 7.356]].forEach(([text, x]) => {
    label(s, text, { x, y: 1.309, w: 1.745, h: 0.404 });
    body(s, LOREM_LONG, { x, y: 1.889, w: 4.93, h: 1.133 });
  });
  headline(s, [['Our Vision And Mission', C.white]], { x: 5.027, y: 4.541, w: 6.594, h: 0.774 });
  photo(s, 1.151, 3.627, 2.491, 3.873, [1.246, 1.246, 0, 0]);
}

// 4 — Facilities
function slide04() {
  const s = newSlide();
  photo(s, 8.347, 1.384, 4.986, 2.366, [1.183, 0, 0, 1.183]);
  photo(s, 0, 3.75, 4.986, 2.366, [0, 1.183, 1.183, 0]);
  navbar(s);
  [['01', 3.823, 3.75], ['02', 9.02, 1.384]].forEach(([text, x, y]) => {
    s.addText(text, {
      shape: S.roundRect, x, y, w: 0.851, h: 0.761, rectRadius: 0.216, fill: { color: C.brown },
      align: 'center', valign: 'middle', fontFace: F.sans, fontSize: 28, bold: true, color: C.white
    });
  });
  [['01. Facilities', 1.047, 1.309, 1.889], ['02. Facilities', 4.233, 1.304, 1.883]].forEach(([text, x, y, ty]) => {
    label(s, text, { x, y, w: 1.745, h: 0.404 });
    body(s, LOREM_MED, { x, y: ty, w: 2.464, h: 1.133 });
  });
  headline(s, [['Skin Beauty ', C.dark], ['Facilities', C.orangeDk]], { x: 6.555, y: 4.405, w: 6.063, h: 0.774 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris culpa qui officia deserunt', { x: 6.555, y: 5.331, w: 5.678, h: 0.868 });
}

// 5 — Our product's
function slide05() {
  const s = newSlide();
  s.addShape(S.rect, { x: 8.324, y: 1.384, w: 5.01, h: 2.229, fill: { color: C.brown } });
  s.addShape(S.rect, { x: 8.324, y: 3.873, w: 5.01, h: 2.243, fill: { color: C.brown } });
  navbar(s);
  headline(s, [['Our', C.orangeDk], [' product\u2019s', C.dark]], { x: 1.05, y: 4.352, w: 4.024, h: 0.774 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore ullamco laboris culpa qui officia deserunt', { x: 1.05, y: 5.321, w: 4.141, h: 0.868 });
  [['01. Product', 1.903, 2.392], ['02. Product', 4.352, 4.841]].forEach(([text, y, by]) => {
    label(s, text, { x: 9.367, y, w: 1.745, h: 0.404, color: C.white });
    body(s, LOREM_MED, { x: 9.367, y: by, w: 2.971, h: 0.868, color: C.white });
  });
  const tile = [0.374, 0.612, 0.374, 0.612];
  photo(s, 6.667, 1.384, 2.333, 2.243, tile);
  photo(s, 6.667, 3.873, 2.333, 2.243, tile);
  photo(s, 0, 1.384, 4.986, 2.366, [0, 1.183, 1.183, 0]);
}

// 6 — Our Team
function slide06() {
  const s = newSlide();
  navbar(s);
  [['01. LAURA', 1.151, 5.261], ['02. CINTA', 4.176, 5.261], ['03. BELLA', 7.2, 5.249]].forEach(([text, x, y]) => {
    s.addText(text, {
      shape: S.rect, x, y, w: 2.491, h: 0.851, fill: { color: C.brown },
      align: 'center', valign: 'middle', fontFace: F.sans, fontSize: 18, bold: true, color: C.white, charSpacing: 3
    });
  });
  s.addText([{ text: 'Our ', options: { color: C.dark } }, { text: 'Team', options: { color: C.orangeDk } }], {
    x: 9.247, y: 3.181, w: 5.069, h: 1.313, rotate: 90, fontFace: F.serif, fontSize: 72, bold: true, valign: 'top'
  });
  [1.151, 4.176, 7.2].forEach(x => photo(s, x, 1.388, 2.491, 3.873, [1.246, 1.246, 0, 0]));
}

// 7 — Best Service
function slide07() {
  const s = newSlide();
  navbar(s);
  s.addShape(S.custGeom, {
    x: 6.667, y: 0.883, w: 6.667, h: 6.617,
    points: roundedRect(6.667, 6.617, [[0.675, 0.72], 0, 0, [0.675, 0.72]]), fill: { color: C.brown }
  });
  headline(s, [['Best ', C.dark], ['Service', C.orangeDk]], { x: 1.04, y: 1.187, w: 4.822, h: 0.774 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore exercitation ullamco laboris officia deserunt', { x: 1.04, y: 2.146, w: 4.205, h: 0.868 });
  [['01. Service', 1.882, 2.411, 'person', 1.773], ['02. Service', 3.457, 3.986, 'doc', 3.357], ['03. Service', 5.034, 5.563, 'person', 4.92]]
    .forEach(([text, y, by, ico, iy]) => {
      label(s, text, { x: 7.99, y, w: 1.745, h: 0.404, color: C.white });
      body(s, LOREM_MED, { x: 7.373, y: by, w: 4.921, h: 0.603, color: C.white });
      icon(s, ico, 7.5, iy, 0.4, 0.42, C.brown);
    });
  photo(s, 1.151, 3.75, 3.984, 2.377, [0, 1.188, 0, 1.188]);
}

// 8 — Portfolio
function slide08() {
  const s = newSlide();
  navbar(s);
  const leaf = [0, 1.188, 0, 1.188];
  s.addShape(S.custGeom, { x: 1.151, y: 3.75, w: 3.984, h: 2.377, points: roundedRect(3.984, 2.377, leaf), fill: { color: C.brown } });
  s.addText('Our Best\nPortfolio', { x: 1.852, y: 4.215, w: 2.513, h: 1.447, fontFace: F.serif, fontSize: 40, bold: true, color: C.white, valign: 'top' });
  photo(s, 9.63, 3.75, 2.541, 3.75, [1.27, 1.27, 0, 0]);
  photo(s, 5.39, 3.75, 3.984, 2.377, leaf);
  [1.151, 3.998, 6.845, 9.692].forEach(x => photo(s, x, 1.373, 2.491, 2.254, [1.127, 1.127, 0, 0]));
}

// 9 — Percentage cards
function slide09() {
  const s = newSlide();
  [['67%', 1.195, C.orange], ['93%', 4.061, C.brown], ['55%', 6.926, C.orange], ['86%', 9.791, C.brown]]
    .forEach(([pct, x, fill], i) => {
      s.addShape(S.round2SameRect, { x, y: 2.979, w: 2.396, h: 2.161, fill: { color: fill } });
      s.addText(pct, { x: x + 0.232, y: 3.512, w: 1.932, h: 0.774, align: 'center', valign: 'top', fontFace: F.sans, fontSize: 40, bold: true, color: C.white });
      s.addText('Title Here', { x: x + 0.506, y: 4.193, w: 1.384, h: 0.37, align: 'center', valign: 'top', fontFace: F.sans, fontSize: 16, italic: true, color: C.white });
      body(s, 'Lorem ipsum dolor sit amet, consectetuer', { x: x + 0.327, y: 5.492, w: 2.069, h: 0.627 });
      s.addText(String(i + 1), {
        shape: S.ellipse, x: x + 0.005, y: 5.599, w: 0.302, h: 0.302, fill: { color: fill },
        align: 'center', valign: 'middle', fontFace: F.sans, fontSize: 10, color: C.white
      });
    });
  navbar(s);
  sectionTitle(s);
}

// 10 — Numbered tiles
function slide10() {
  const s = newSlide();
  [1.171, 4.949, 8.722].forEach((x, col) => {
    [2.984, 4.66].forEach((y, row) => {
      const n = row * 3 + col + 1;
      s.addShape(S.roundRect, { x, y, w: 3.44, h: 1.297, fill: { color: n % 2 ? C.orange : C.brown } });
      s.addText('0' + n + '.', { x: x + 0.395, y: y + 0.317, w: 1.036, h: 0.64, fontFace: F.sans, fontSize: 32, bold: true, color: C.white, valign: 'top' });
      body(s, 'Lorem ipsum dolor sit amet, adipi', { x: x + 1.405, y: y + 0.317, w: 1.655, h: 0.627, color: C.white });
    });
  });
  navbar(s);
  sectionTitle(s);
}

// 11 — Cards with "Learn More"
function slide11() {
  const s = newSlide();
  [1.155, 5.055, 8.954].forEach((x, i) => {
    const fill = i === 1 ? C.brown : C.orange;
    s.addShape(S.roundRect, { x, y: 2.983, w: 3.235, h: 2.158, fill: { color: fill } });
    s.addText(String(i + 1), {
      shape: S.ellipse, x: x + 0.261, y: 3.225, w: 0.499, h: 0.499, fill: { color: C.white, transparency: 80 },
      align: 'center', valign: 'middle', fontFace: F.sans, fontSize: 12, bold: true, color: C.white
    });
    s.addText('Title Here', { x: x + 0.875, y: 3.225, w: 1.841, h: 0.37, fontFace: F.sans, fontSize: 12, bold: true, color: C.white, lineSpacingMultiple: 1.5, valign: 'top' });
    body(s, 'Lorem ipsum dolor sit am et dolor ipsum dolor ', { x: x + 0.875, y: 3.6, w: 2.08, h: 0.627, color: C.white });
    s.addText('Learn More', {
      shape: S.roundRect, x: x + 0.919, y: 4.4, w: 1.186, h: 0.42, rectRadius: 0.21, fill: { color: C.white, transparency: 80 },
      align: 'center', valign: 'middle', fontFace: F.sans, fontSize: 10, bold: true, color: C.white
    });
    body(s, 'Lorem nsect ipsum ame co nsect etuer adip', { x: x + 0.852, y: 5.445, w: 2.102, h: 0.627 });
    s.addText('0' + (i + 1), {
      shape: S.ellipse, x: x + 0.013, y: 5.476, w: 0.582, h: 0.582, fill: { color: fill },
      align: 'center', valign: 'middle', fontFace: F.sans, fontSize: 10.5, bold: true, color: C.white
    });
  });
  navbar(s);
  sectionTitle(s);
}

// 12 — Circle icon + text, 2x2
function slide12() {
  const s = newSlide();
  [[1.156, 2.987, C.orange], [1.156, 4.868, C.brown], [7.187, 2.99, C.brown], [7.187, 4.871, C.orange]]
    .forEach(([x, y, fill]) => {
      s.addShape(S.ellipse, { x, y, w: 1.281, h: 1.281, fill: { color: fill } });
      icon(s, 'bars', x + 0.4, y + 0.45, 0.48, 0.36, fill);
      s.addText('Title Here', { x: x + 1.49, y: y + 0.029, w: 2.479, h: 0.303, fontFace: F.sans, fontSize: 12, bold: true, color: fill, valign: 'top' });
      body(s, LOREM_TITLE, { x: x + 1.49, y: y + 0.349, w: 3.5, h: 0.868 });
    });
  navbar(s);
  sectionTitle(s);
}

// 13 — Price columns
function slide13() {
  const s = newSlide();
  [['$3.500', 1.201, C.orange], ['$6.500', 5.258, C.brown], ['$9.500', 9.315, C.orange]]
    .forEach(([price, x, fill]) => {
      body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Mae cenas porttmassa. Fusce posuere, magna porttmassa. Fusce posuere, magna sed pulvinar posuere, magna portt massa. Fusce posuere, posuere, magna portt massa magna portt massa', { x: x - 0.094, y: 4.11, w: 3.018, h: 1.928 });
      s.addShape(S.ellipse, { x, y: 2.999, w: 0.741, h: 0.741, fill: { color: fill } });
      s.addShape(S.ellipse, { x: x + 0.169, y: 3.168, w: 0.403, h: 0.403, fill: { type: 'none' }, line: { color: C.white, width: 1 } });
      s.addText('$', { x: x + 0.169, y: 3.168, w: 0.403, h: 0.403, align: 'center', valign: 'middle', fontFace: F.sans, fontSize: 12, bold: true, color: C.white });
      s.addText(price, { x: x + 0.94, y: 2.998, w: 1.589, h: 0.505, fontFace: F.sans, fontSize: 24, bold: true, color: C.dark, valign: 'top' });
      s.addText('Your Title Here', { x: x + 0.94, y: 3.455, w: 1.589, h: 0.286, fontFace: F.sans, fontSize: 11, italic: true, color: C.dark, valign: 'top' });
    });
  navbar(s);
  sectionTitle(s);
}

// 14 — Numbered circles with arrows
function slide14() {
  const s = newSlide();
  [1.696, 4.655, 7.614, 10.574].forEach((x, i) => {
    const fill = i % 2 ? C.brown : C.orange;
    s.addText('0' + (i + 1), {
      shape: S.ellipse, x, y: 2.986, w: 1.092, h: 1.092, fill: { color: fill },
      align: 'center', valign: 'middle', fontFace: F.sans, fontSize: 20, bold: true, color: C.white
    });
    s.addShape(S.line, { x: x + 0.546, y: 4.247, w: 0, h: 0.4, line: { color: fill, endArrowType: 'arrow' } });
    s.addText('Title Here', { x: x - 0.648, y: 4.816, w: 2.387, h: 0.372, align: 'center', valign: 'top', fontFace: F.sans, fontSize: 12, bold: true, color: C.dark, lineSpacingMultiple: 1.5 });
    body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas portt. ', { x: x - 0.648, y: 5.191, w: 2.387, h: 0.868, align: 'center' });
  });
  navbar(s);
  sectionTitle(s);
}

// 15 — Segmented donut + progress bars
const DONUT = [
  { x: 1.565, y: 2.981, w: 1.303, h: 1.212, fill: C.orange, steps: [[0.9895, 0], [1, 0.8016], [0.9857, 0.8096], [0.978, 0.8088], ['c', 0.829, 0.8088, 0.6941, 0.8737, 0.5964, 0.9787], [0.5801, 1], [0, 0.5626], ['c', 0.2267, 0.2152, 0.5948, 0.0059, 0.9895, 0]] },
  { x: 1.233, y: 3.706, w: 1.048, h: 1.422, fill: C.brown, steps: [[0.287, 0], [1, 0.3769], [0.9769, 0.3975], ['c', 0.9045, 0.4764, 0.8622, 0.5716, 0.8622, 0.674], ['c', 0.8622, 0.7081, 0.8669, 0.7415, 0.8758, 0.7737], [0.8901, 0.8076], [0.0726, 1], ['c', -0.0747, 0.6605, 0.0048, 0.2896, 0.287, 0]] },
  { x: 1.331, y: 4.915, w: 1.462, h: 1.384, fill: C.orange, steps: [[0.5832, 0], [0.5882, 0.017], ['c', 0.649, 0.169, 0.7779, 0.2832, 0.9345, 0.3171], [1, 0.3241], [0.9332, 1], ['c', 0.497, 0.9519, 0.1299, 0.6337, 0, 0.1911], [0.5832, 0]] },
  { x: 2.758, y: 5.07, w: 1.458, h: 1.245, fill: C.brown, steps: [[0.4488, 0], [1, 0.4658], ['c', 0.7708, 0.8379, 0.3898, 1.039, 0, 0.9937], [0.0641, 0.2383], [0.153, 0.2278], ['c', 0.2472, 0.2053, 0.3314, 0.1505, 0.3968, 0.0738], [0.4488, 0]] }
];
function slide15() {
  const s = newSlide();
  DONUT.forEach(d => s.addShape(S.custGeom, { x: d.x, y: d.y, w: d.w, h: d.h, points: outline(d.w, d.h, d.steps), fill: { color: d.fill } }));
  [['77%', 1.817, 3.378], ['70%', 1.187, 4.238], ['97%', 1.631, 5.305], ['78%', 2.893, 5.47]].forEach(([t, x, y]) => {
    s.addText(t, { x, y, w: 1.043, h: 0.438, align: 'center', valign: 'top', fontFace: F.sans, fontSize: 20, bold: true, color: C.white });
  });
  s.addText('Title Here ', { x: 2.216, y: 4.346, w: 1.155, h: 0.707, align: 'center', valign: 'top', fontFace: F.sans, fontSize: 18, bold: true, color: C.orange });

  [['01', 5.044, 3.021, C.orange], ['02', 8.806, 3.017, C.brown], ['03', 5.044, 4.211, C.brown], ['04', 8.806, 4.207, C.orange]]
    .forEach(([n, x, y, col]) => {
      s.addText(' Title Here ' + n, { x, y, w: 1.597, h: 0.303, fontFace: F.sans, fontSize: 12, bold: true, color: col, valign: 'top' });
      body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing .aecenas porttitor.', { x: x + 0.066, y: y + 0.23, w: 3.338, h: 0.603 });
    });

  // Progress bars: grey track + coloured fill + right-aligned percentage.
  [[5.205, 5.456, 1.872, C.orange, '77%', 7.959, 5.416], [5.205, 5.893, 1.717, C.brown, '70%', 7.95, 5.855],
   [8.964, 5.462, 2.473, C.orange, '97%', 11.735, 5.416], [8.964, 5.899, 1.92, C.brown, '78%', 11.735, 5.855]]
    .forEach(([x, y, fw, col, pct, tx, ty]) => {
      s.addShape(S.roundRect, { x, y, w: 2.726, h: 0.211, rectRadius: 0.105, fill: { color: C.grey } });
      s.addShape(S.roundRect, { x, y, w: fw, h: 0.211, rectRadius: 0.105, fill: { color: col } });
      s.addText(pct, { x: tx, y: ty, w: 0.502, h: 0.278, align: 'right', wrap: false, valign: 'top', fontFace: F.alt, fontSize: 10.5, color: C.dark });
    });
  navbar(s);
  sectionTitle(s);
}

// 16 — Tabbed panel + arrow rows
function slide16() {
  const s = newSlide();
  [[3.119, C.peach], [4.105, C.rose], [5.12, C.peach]].forEach(([y, col], i) => {
    s.addShape(S.homePlate, { x: i === 0 ? 5.547 : 6.333, y, w: i === 0 ? 6.63 : 5.844, h: i === 2 ? 0.823 : 0.851, fill: { color: col } });
  });
  s.addShape(S.round1Rect, { x: 1.522, y: 2.962, w: 4.811, h: 3.138, fill: { color: C.orange } });
  [['01', 3.119, C.peach, 3.777, C.orangeDk], ['02', 4.105, C.rose, 4.764, C.rust], ['03', 5.092, C.peach, 5.75, C.orangeDk]]
    .forEach(([n, y, col, ty, tcol]) => {
      s.addShape(S.rtTriangle, { x: 1.156, y: ty, w: 0.367, h: 0.193, flipH: true, flipV: true, fill: { color: tcol } });
      s.addText(n, {
        shape: S.round1Rect, x: 1.156, y, w: 0.894, h: 0.658, fill: { color: col },
        align: 'center', valign: 'middle', fontFace: F.sans, fontSize: 18, bold: true, color: C.dark
      });
    });
  [3.146, 4.133, 5.12].forEach(y => body(s, 'Lorem ipsum dolor sit amet, elit consectetuer adipiscing elit dolor sit amet', { x: 2.2, y, w: 3.744, h: 0.603, color: C.white }));
  [['25%', 3.215, C.peachLt, 3.405], ['35%', 4.202, C.roseLt, 4.392], ['45%', 5.203, C.peachLt, 5.393]]
    .forEach(([pct, y, col, ty]) => {
      s.addText(pct, {
        shape: S.roundRect, x: 6.47, y, w: 0.658, h: 0.658, fill: { color: col },
        align: 'center', valign: 'middle', fontFace: F.sans, fontSize: 12, bold: true, color: C.dark
      });
      s.addText('Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ', { x: 7.265, y: ty, w: 4.444, h: 0.278, align: 'justify', valign: 'top', fontFace: F.sans, fontSize: 10.5, color: C.dark });
    });
  navbar(s);
  sectionTitle(s);
}

// 17 — Concentric rings
function slide17() {
  const s = newSlide();
  s.addShape(S.roundRect, { x: 6.463, y: 2.962, w: 2.407, h: 0.353, rectRadius: 0.177, fill: { color: C.orange } });
  s.addShape(S.ellipse, { x: 5.026, y: 2.962, w: 3.172, h: 3.172, fill: { color: C.orange } });
  s.addShape(S.ellipse, { x: 5.361, y: 3.296, w: 2.503, h: 2.503, fill: { color: C.brown } });
  s.addText('Title Here', {
    shape: S.ellipse, x: 5.714, y: 3.65, w: 1.796, h: 1.796, fill: { color: C.white },
    align: 'center', valign: 'middle', fontFace: F.sans, fontSize: 24, bold: true, color: C.brown
  });
  s.addShape(S.roundRect, { x: 6.521, y: 3.296, w: 2.83, h: 0.353, rectRadius: 0.177, fill: { color: C.brown } });
  s.addShape(S.roundRect, { x: 4.296, y: 5.782, w: 2.407, h: 0.353, rectRadius: 0.177, fill: { color: C.orange } });
  s.addShape(S.roundRect, { x: 3.983, y: 5.439, w: 2.83, h: 0.353, rectRadius: 0.177, fill: { color: C.brown } });
  [1.11, 9.973].forEach((x, side) => {
    [2.85, 4.194, 5.539].forEach(y => body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing', { x, y, w: 2.305, h: 0.603 }));
    [[3.665, side ? C.orange : C.brown], [5.009, side ? C.brown : C.orange]].forEach(([y, col]) => {
      s.addShape(S.rightArrow, {
        x: side ? 10.955 : 2.092, y, w: 0.34, h: 0.317, fill: { color: col },
        rotate: side ? 90 : 270, flipV: !!side
      });
    });
  });
  navbar(s);
  sectionTitle(s);
}

// 18 — Icon rows with progress pills
function slide18() {
  const s = newSlide();
  s.addShape(S.custGeom, {
    x: 0, y: 2.955, w: 4.455, h: 4.545, fill: { color: C.cream },
    points: outline(4.455, 4.545, [[0, 0], [1, 0.3126], [0.7763, 1], [0.5909, 1], [0.7781, 0.4249], [0, 0.1817]])
  });
  [['music', 2.969, 3.256, C.orange, 3.187, 2.953, C.orange],
   ['gear', 4.182, 4.468, C.brown, 4.416, 4.182, C.brown],
   ['recycle', 5.394, 5.68, C.orange, 5.564, 5.33, C.orange]]
    .forEach(([ico, y, ay, col, by, ly]) => {
      s.addShape(S.ellipse, { x: 1.936, y, w: 0.781, h: 0.781, fill: { color: col } });
      icon(s, ico, 2.126, y + 0.19, 0.4, 0.4, col);
      s.addShape(S.rightArrow, { x: 2.99, y: ay, w: 1.597, h: 0.208, fill: { color: col } });
      s.addText('Title Here ', { x: 5.023, y: ly, w: 1.448, h: 0.303, fontFace: F.sans, fontSize: 12, bold: true, color: col, valign: 'top' });
      body(s, 'Lorem ipsum dolor sit amet, consectetuer dolor. sed do eiusmod tempor inci.', { x: 5.023, y: by, w: 3.487, h: 0.603 });
    });
  [['70% ', 2.969, C.orange], ['90% ', 4.182, C.brown], ['85% ', 5.394, C.orange]].forEach(([pct, y, col]) => {
    s.addShape(S.flowChartTerminator, { x: 9.078, y, w: 3.1, h: 0.781, fill: { color: col } });
    s.addText([{ text: pct, options: { fontSize: 20, bold: true } }, { text: 'in progress', options: { fontSize: 14 } }],
      { x: 9.368, y: y + 0.172, w: 2.522, h: 0.438, align: 'center', valign: 'top', fontFace: F.sans, color: C.white });
  });
  navbar(s);
  sectionTitle(s);
}

// 19 — Get in Touch
function slide19() {
  const s = newSlide();
  navbar(s);
  headline(s, [['Get in ', C.dark], ['Touch', C.orangeDk]], { x: 1.04, y: 1.187, w: 4.822, h: 0.774 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore exercitation ullamco laboris officia deserunt', { x: 1.04, y: 2.146, w: 4.918, h: 0.868 });
  [['phone', '+123 4567 8910', 1.378, 1.466, 1.334], ['mail', 'yourmail@mail.com', 1.928, 2.02, 1.569],
   ['globe', 'www.yourwebsite.com', 2.477, 2.57, 1.759], ['home', '123 Anywhere St., Any City', 3.027, 3.119, 2.018]]
    .forEach(([ico, text, y, ty, tw]) => {
      s.addShape(S.ellipse, { x: 7.472, y, w: 0.462, h: 0.462, fill: { color: C.dark } });
      icon(s, ico, 7.578, y + 0.106, 0.25, 0.25, C.dark);
      s.addShape(S.round1Rect, { x: 8.055, y, w: 4.128, h: 0.462, rectRadius: 0.231, fill: { color: C.brown } });
      s.addText(text, { x: 8.186, y: ty, w: tw, h: 0.278, wrap: false, valign: 'top', fontFace: F.sans, fontSize: 10.5, color: C.white });
    });
  photo(s, 8.186, 4.55, 3.984, 2.377, [0, 1.188, 0, 1.188]);
  photo(s, 1.151, 3.75, 5.516, 3.75, [0, 1.875, 0, 1.875]);
}

// 20 — Thanks For Watching
function slide20() {
  const s = newSlide();
  navbar(s);
  headline(s, [['Thanks For ', C.dark], ['Watching', C.orangeDk]], { x: 3.256, y: 1.578, w: 6.822, h: 0.774, align: 'center' });
  body(s, LOREM_LONG, { x: 1.151, y: 2.433, w: 11.032, h: 0.603, align: 'center' });
  photo(s, 8.199, 4.543, 3.984, 2.957, [[1.49, 1.48], 0, [1.29, 1.47], 0]);
  photo(s, 5.098, 3.75, 3.137, 3.75, [1.569, 1.569, 0, 0]);
  photo(s, 1.151, 4.543, 3.947, 2.957, [0, [1.48, 1.48], 0, [1.29, 1.47]]);
}

[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
 slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
  .forEach(build => build());

pptx.writeFile({ fileName: path.join(__dirname, '0e4596f5-8501-4ad0-98b3-87ea843e14d9_grok_final.pptx') })
  .then(f => console.log('wrote', f));
