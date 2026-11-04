/**
 * Recreation of "Product Demo" (16:9, 16 slides) with pptxgenjs.
 * Raster images in the source deck are replaced by flat colour placeholders.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */
const PINK = 'F493C6';
const BLUE = '5093E1';
const LTBLUE = '5B9BD5';
const DKBLUE = '3C6EA9';
const MIDBLUE = '96BEED';
const PALEBLUE = 'B9D4F3';
const ICEBLUE = 'DCE9F9';
const YELLOW = 'FBCC31';
const GREEN = '70AD47';
const SLATE = '576579';
const WHITE = 'FFFFFF';
const BLACK = '000000';
const INK = '262626';
const BG = 'F2F2F2';        // master background (bg1 lumMod 95%)
const SKIN = 'FFBC96';
const PALEPINK = 'F8BEDD';

const MAJOR = 'Cormorant Garamond SemiBold';   // theme major latin font
const MINOR = 'Poppins';                       // theme minor latin font

/* ---------------------------------------------------------------- helpers */
const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'W16x9', width: 13.333, height: 7.5 });
pptx.layout = 'W16x9';
pptx.author = 'pptxgenjs';
pptx.title = 'Product Demo';

const S = pptx.ShapeType;

/** roundRect adjustment value (0-100000) -> pptxgenjs rectRadius in inches */
const radius = (adj, w, h) => (adj / 100000) * Math.min(w, h);

/* the deck uses three drop-shadow recipes */
const SOFT = { shadow: { type: 'outer', blur: 50, offset: 0, angle: 90, color: BLACK, opacity: 0.05 } };
const DROP = { shadow: { type: 'outer', blur: 35, offset: 10, angle: 50, color: BLACK, opacity: 0.05 } };
const LIFT = { shadow: { type: 'outer', blur: 35, offset: 3, angle: 45, color: BLACK, opacity: 0.10 } };

/** solid shape */
function shape(s, type, o) {
  return s.addShape(type, o);
}
function rect(s, x, y, w, h, color, extra) {
  return s.addShape(S.rect, Object.assign({ x, y, w, h, fill: { color } }, extra));
}
function roundRect(s, x, y, w, h, color, adj, extra) {
  return s.addShape(S.roundRect, Object.assign(
    { x, y, w, h, fill: { color }, rectRadius: radius(adj, w, h) }, extra));
}
function oval(s, x, y, w, h, color, extra) {
  return s.addShape(S.ellipse, Object.assign({ x, y, w, h, fill: { color } }, extra));
}
function line(s, x, y, w, h, color, width, dash, extra) {
  return s.addShape(S.line, Object.assign(
    { x, y, w, h, line: { color, width, dashType: dash || 'solid' } }, extra));
}
/** free-form polygon from [[x,y],...] normalised 0..1 coordinates */
function poly(s, x, y, w, h, color, pts, extra) {
  return s.addShape(S.custGeom, Object.assign({
    x, y, w, h, fill: { color },
    points: pts.map(p => ({ x: +(p[0] * w).toFixed(3), y: +(p[1] * h).toFixed(3) }))
      .concat([{ close: true }]),
  }, extra));
}

/** text box: PowerPoint defaults (top anchored, 0.1/0.05 insets) */
function text(s, str, o) {
  return s.addText(str, Object.assign({
    fontFace: MINOR, fontSize: 18, color: BLACK, valign: 'top', wrap: true,
  }, o));
}
/** heading in the display serif face */
function head(s, str, o) {
  return text(s, str, Object.assign({ fontFace: MAJOR }, o));
}
/** 12pt Poppins body copy */
function body(s, str, o) {
  return text(s, str, Object.assign({ fontSize: 12, lineSpacingMultiple: 1.3 }, o));
}

/* --------------------------------------------------- icon mini-vocabulary */
/* Every icon draws inside the square (x, y, d) in one colour, so the same
   glyph can be reused knocked-out white on pink or pink on white.          */
const ICON = {
  target: (s, x, y, d, c) => {
    s.addShape(S.donut, { x, y, w: d, h: d, fill: { color: c }, rectRadius: d * 0.07 });
    s.addShape(S.donut, { x: x + d * 0.24, y: y + d * 0.24, w: d * 0.52, h: d * 0.52, fill: { color: c }, rectRadius: d * 0.07 });
    oval(s, x + d * 0.42, y + d * 0.42, d * 0.16, d * 0.16, c);
  },
  trophy: (s, x, y, d, c) => {
    poly(s, x + d * 0.24, y + d * 0.06, d * 0.52, d * 0.48, c, [[0, 0], [1, 0], [0.8, 1], [0.2, 1]]);
    s.addShape(S.blockArc, { x: x + d * 0.02, y: y + d * 0.04, w: d * 0.32, h: d * 0.34, fill: { color: c }, rotate: 90 });
    s.addShape(S.blockArc, { x: x + d * 0.66, y: y + d * 0.04, w: d * 0.32, h: d * 0.34, fill: { color: c }, rotate: 270 });
    rect(s, x + d * 0.44, y + d * 0.52, d * 0.12, d * 0.24, c);
    roundRect(s, x + d * 0.26, y + d * 0.76, d * 0.48, d * 0.12, c, 30000);
  },
  medal: (s, x, y, d, c) => {
    poly(s, x + d * 0.26, y, d * 0.28, d * 0.52, c, [[0.55, 0], [1, 0.05], [0.45, 1], [0, 0.9]]);
    poly(s, x + d * 0.46, y, d * 0.28, d * 0.52, c, [[0, 0.05], [0.45, 0], [1, 0.9], [0.55, 1]]);
    oval(s, x + d * 0.2, y + d * 0.4, d * 0.6, d * 0.6, c);
    s.addShape(S.star5, { x: x + d * 0.34, y: y + d * 0.54, w: d * 0.32, h: d * 0.32, fill: { color: WHITE } });
  },
  rocket: (s, x, y, d, c) => {
    poly(s, x + d * 0.34, y, d * 0.32, d * 0.72, c, [[0.5, 0], [1, 0.35], [1, 1], [0, 1], [0, 0.35]]);
    poly(s, x + d * 0.1, y + d * 0.42, d * 0.26, d * 0.34, c, [[1, 0], [1, 1], [0, 1]]);
    poly(s, x + d * 0.64, y + d * 0.42, d * 0.26, d * 0.34, c, [[0, 0], [1, 1], [0, 1]]);
    poly(s, x + d * 0.4, y + d * 0.74, d * 0.2, d * 0.24, c, [[0, 0], [1, 0], [0.5, 1]]);
    oval(s, x + d * 0.42, y + d * 0.22, d * 0.16, d * 0.16, c, { fill: { color: WHITE } });
  },
  dumbbell: (s, x, y, d, c) => {
    rect(s, x + d * 0.18, y + d * 0.43, d * 0.64, d * 0.14, c);
    roundRect(s, x + d * 0.04, y + d * 0.3, d * 0.15, d * 0.4, c, 30000);
    roundRect(s, x + d * 0.81, y + d * 0.3, d * 0.15, d * 0.4, c, 30000);
    roundRect(s, x, y + d * 0.38, d * 0.07, d * 0.24, c, 40000);
    roundRect(s, x + d * 0.93, y + d * 0.38, d * 0.07, d * 0.24, c, 40000);
  },
  briefcase: (s, x, y, d, c) => {
    s.addShape(S.custGeom, {
      x: x + d * 0.32, y: y + d * 0.08, w: d * 0.36, h: d * 0.2,
      fill: { type: 'none' }, line: { color: c, width: d * 3 },
      points: [{ x: 0, y: d * 0.2 }, { x: 0, y: d * 0.04 }, { x: d * 0.06, y: 0 },
      { x: d * 0.3, y: 0 }, { x: d * 0.36, y: d * 0.04 }, { x: d * 0.36, y: d * 0.2 }],
    });
    roundRect(s, x + d * 0.04, y + d * 0.26, d * 0.92, d * 0.6, c, 12000);
    rect(s, x + d * 0.42, y + d * 0.48, d * 0.16, d * 0.14, WHITE);
  },
  people: (s, x, y, d, c) => {
    [[0.15, 0.3, 0.2], [0.5, 0.18, 0.24], [0.85, 0.3, 0.2]].forEach(([cx, cy, r]) => {
      oval(s, x + d * (cx - r / 2), y + d * cy, d * r, d * r, c);
      poly(s, x + d * (cx - 0.22), y + d * (cy + r), d * 0.44, d * 0.38, c,
        [[0.5, 0], [0.9, 0.3], [1, 1], [0, 1], [0.1, 0.3]]);
    });
  },
  bars: (s, x, y, d, c) => {
    [[0.1, 0.34], [0.41, 0.56], [0.72, 0.78]].forEach(([bx, bh]) =>
      rect(s, x + d * bx, y + d * (0.9 - bh), d * 0.18, d * bh, c));
  },
  pie: (s, x, y, d, c) => {           // 3/4 disc with the quarter slice pulled out
    s.addShape(S.pie, { x, y: y + d * 0.12, w: d * 0.84, h: d * 0.84, fill: { color: c } });
    s.addShape(S.pieWedge, { x: x + d * 0.5, y: y + d * 0.04, w: d * 0.42, h: d * 0.42, fill: { color: c } });
  },
  bulb: (s, x, y, d, c) => {
    oval(s, x + d * 0.2, y + d * 0.06, d * 0.6, d * 0.6, c);
    rect(s, x + d * 0.38, y + d * 0.6, d * 0.24, d * 0.2, c);
    roundRect(s, x + d * 0.38, y + d * 0.82, d * 0.24, d * 0.12, c, 35000);
  },
  clock: (s, x, y, d, c) => {
    s.addShape(S.donut, { x, y, w: d, h: d, fill: { color: c }, rectRadius: d * 0.06 });
    rect(s, x + d * 0.47, y + d * 0.26, d * 0.06, d * 0.28, c);
    rect(s, x + d * 0.5, y + d * 0.48, d * 0.2, d * 0.06, c);
  },
  calendar: (s, x, y, d, c) => {
    rect(s, x, y + d * 0.08, d, d * 0.16, c);
    rect(s, x, y + d * 0.24, d, d * 0.7, c, { fill: { type: 'none' }, line: { color: c, width: 1.25 } });
    for (let r = 0; r < 3; r++) for (let k = 0; k < 4; k++)
      rect(s, x + d * (0.09 + k * 0.22), y + d * (0.36 + r * 0.19), d * 0.12, d * 0.1, c);
  },
  list: (s, x, y, d, c) => {
    s.addShape(S.donut, { x, y, w: d, h: d, fill: { color: c }, rectRadius: d * 0.06 });
    rect(s, x + d * 0.28, y + d * 0.4, d * 0.44, d * 0.06, c);
    rect(s, x + d * 0.28, y + d * 0.55, d * 0.44, d * 0.06, c);
  },
  dollar: (s, x, y, d, c) => {
    s.addShape(S.donut, { x, y, w: d, h: d, fill: { color: c }, rectRadius: d * 0.06 });
    text(s, '$', { x: x + d * 0.15, y: y + d * 0.15, w: d * 0.7, h: d * 0.7, align: 'center', valign: 'middle', fontSize: d * 44, bold: true, color: c });
  },
  handshake: (s, x, y, d, c) => {       // clasped hands inside a ring, badge above
    s.addShape(S.donut, { x, y, w: d, h: d, fill: { color: c }, rectRadius: d * 0.04 });
    rect(s, x + d * 0.16, y + d * 0.46, d * 0.32, d * 0.1, c, { rotate: 20 });
    rect(s, x + d * 0.52, y + d * 0.46, d * 0.32, d * 0.1, c, { rotate: -20 });
    roundRect(s, x + d * 0.36, y + d * 0.46, d * 0.28, d * 0.16, c, 40000);
    oval(s, x + d * 0.6, y + d * 0.2, d * 0.22, d * 0.22, c);
  },
  atom: (s, x, y, d, c) => {
    [30, 90, 150].forEach(a => s.addShape(S.ellipse, {
      x: x + d * 0.03, y: y + d * 0.28, w: d * 0.94, h: d * 0.44,
      fill: { type: 'none' }, line: { color: c, width: d * 2.6 }, rotate: a,
    }));
    oval(s, x + d * 0.38, y + d * 0.38, d * 0.24, d * 0.24, c);
  },
  growth: (s, x, y, d, c) => {        // bar chart, rising arrow and a dollar sign
    [[0.06, 0.34], [0.36, 0.5], [0.66, 0.66]].forEach(([bx, bh]) =>
      rect(s, x + d * bx, y + d * (0.94 - bh), d * 0.2, d * bh, c));
    poly(s, x + d * 0.3, y + d * 0.06, d * 0.68, d * 0.4, c,
      [[0.55, 0], [1, 0], [1, 0.45], [0.84, 0.3], [0.42, 1], [0.16, 0.72], [0, 0.9], [0, 0.6], [0.16, 0.42], [0.42, 0.7], [0.76, 0.13]]);
    text(s, '$', { x: x - d * 0.04, y: y + d * 0.06, w: d * 0.36, h: d * 0.4, align: 'center', valign: 'middle', fontSize: d * 30, bold: true, color: c });
  },
  gear: (s, x, y, d, c) => {
    s.addShape(S.gear9, { x, y, w: d, h: d, fill: { color: c } });
    oval(s, x + d * 0.34, y + d * 0.29, d * 0.26, d * 0.26, c, { fill: { color: WHITE } });
  },
  bell: (s, x, y, d, c) => {
    poly(s, x + d * 0.12, y + d * 0.12, d * 0.76, d * 0.58, c,
      [[0.5, 0], [0.8, 0.15], [0.88, 0.75], [1, 1], [0, 1], [0.12, 0.75], [0.2, 0.15]]);
    roundRect(s, x + d * 0.06, y + d * 0.68, d * 0.88, d * 0.1, c, 40000);
    oval(s, x + d * 0.41, y + d * 0.8, d * 0.18, d * 0.18, c);
    oval(s, x + d * 0.44, y + d * 0.02, d * 0.12, d * 0.12, c);
  },
  flower: (s, x, y, d, c) => {
    [0, 72, 144, 216, 288].forEach(a => {
      const r = (a - 90) * Math.PI / 180;
      oval(s, x + d * (0.5 + 0.3 * Math.cos(r) - 0.2), y + d * (0.5 + 0.3 * Math.sin(r) - 0.2),
        d * 0.4, d * 0.4, c);
    });
    oval(s, x + d * 0.33, y + d * 0.33, d * 0.34, d * 0.34, c, { fill: { color: WHITE } });
  },
  pin: (s, x, y, d, c) => {
    poly(s, x + d * 0.14, y + d * 0.04, d * 0.72, d * 0.92, c,
      [[0.5, 1], [0.08, 0.45], [0.03, 0.28], [0.2, 0.06], [0.5, 0], [0.8, 0.06], [0.97, 0.28], [0.92, 0.45]]);
    oval(s, x + d * 0.36, y + d * 0.2, d * 0.28, d * 0.28, c, { fill: { color: WHITE } });
  },
  /* folded map sheet; the pin is knocked out, so it needs the backdrop colour */
  map: (s, x, y, d, c, bg) => {
    poly(s, x, y + d * 0.2, d, d * 0.62, c,
      [[0, 0.08], [0.25, 0], [0.5, 0.55], [0.75, 0], [1, 0.08], [1, 0.85], [0.75, 1], [0.5, 0.78], [0.25, 1], [0, 0.85]]);
    poly(s, x + d * 0.26, y, d * 0.48, d * 0.7, bg, [[0.5, 1], [0.04, 0.42], [0.5, 0], [0.96, 0.42]]);
    poly(s, x + d * 0.33, y + d * 0.06, d * 0.34, d * 0.5, c, [[0.5, 1], [0.05, 0.4], [0.5, 0], [0.95, 0.4]]);
  },
  phone: (s, x, y, d, c) => poly(s, x + d * 0.1, y + d * 0.1, d * 0.8, d * 0.8, c,
    [[0.05, 0], [0.32, 0], [0.42, 0.28], [0.28, 0.4], [0.6, 0.72], [0.72, 0.58], [1, 0.68],
    [1, 0.95], [0.8, 1], [0.35, 0.8], [0.08, 0.42]]),
  split: (s, x, y, d, c) => s.addShape(S.custGeom, {   // one stem forking into two arrows
    x, y, w: d, h: d, fill: { color: c },
    points: [{ x: d * 0.5, y: d },
    { x: d * 0.5, y: d * 0.62, curve: { type: 'cubic', x1: d * 0.5, y1: d * 0.85, x2: d * 0.44, y2: d * 0.68 } },
    { x: d * 0.22, y: d * 0.5 }, { x: d * 0.22, y: d * 0.62 }, { x: 0, y: d * 0.4 },
    { x: d * 0.22, y: d * 0.18 }, { x: d * 0.22, y: d * 0.3 },
    { x: d * 0.5, y: d * 0.46, curve: { type: 'cubic', x1: d * 0.4, y1: d * 0.32, x2: d * 0.46, y2: d * 0.4 } },
    { x: d * 0.78, y: d * 0.3, curve: { type: 'cubic', x1: d * 0.54, y1: d * 0.4, x2: d * 0.6, y2: d * 0.32 } },
    { x: d * 0.78, y: d * 0.18 }, { x: d, y: d * 0.4 }, { x: d * 0.78, y: d * 0.62 }, { x: d * 0.78, y: d * 0.5 },
    { x: d * 0.62, y: d * 0.62, curve: { type: 'cubic', x1: d * 0.68, y1: d * 0.52, x2: d * 0.62, y2: d * 0.55 } },
    { x: d * 0.62, y: d }, { close: true }],
  }),
  screen: (s, x, y, d, c) => {          // presentation board: outline + bars
    s.addShape(S.roundRect, {
      x: x + d * 0.02, y: y + d * 0.08, w: d * 0.96, h: d * 0.66,
      fill: { type: 'none' }, line: { color: c, width: d * 3.2 }, rectRadius: d * 0.06,
    });
    rect(s, x + d * 0.44, y + d * 0.74, d * 0.12, d * 0.16, c);
    roundRect(s, x + d * 0.26, y + d * 0.88, d * 0.48, d * 0.1, c, 40000);
    [[0.2, 0.2], [0.42, 0.32], [0.64, 0.14]].forEach(([bx, bh]) =>
      rect(s, x + d * bx, y + d * (0.62 - bh), d * 0.14, d * bh, c));
  },
};

/* ------------------------------------------------- master chrome per slide */
const NAV = [['Home', 9.344, 0.819], ['About', 10.342, 0.819], ['Contact', 11.296, 0.923]];

function logoMark(s, x, y, d, color) {         // "ribbon" award mark
  poly(s, x + d * 0.18, y + d * 0.5, d * 0.3, d * 0.5, color, [[0.45, 0], [1, 0.15], [0.6, 1], [0.3, 0.62], [0, 0.8]]);
  poly(s, x + d * 0.52, y + d * 0.5, d * 0.3, d * 0.5, color, [[0.55, 0], [1, 0.8], [0.7, 0.62], [0.4, 1], [0, 0.15]]);
  s.addShape(S.donut, { x: x + d * 0.11, y: y + d * 0.03, w: d * 0.78, h: d * 0.78, fill: { color } });
}
function hamburger(s, x, y, d, color) {
  [0, 0.294, 0.588, 0.882].forEach(o => rect(s, x, y + d * o, d, d * 0.118, color));
}
/** header bar + page number, exactly as on the slide master */
function chrome(s, pageNo) {
  logoMark(s, 0.540, 0.353, 0.303, BLACK);
  text(s, 'Product Demo', { x: 0.931, y: 0.353, w: 1.684, h: 0.303, fontSize: 12 });
  NAV.forEach(([label, x, w]) => text(s, label, { x, y: 0.353, w, h: 0.303, fontSize: 12 }));
  hamburger(s, 12.486, 0.404, 0.227, BLACK);
  text(s, 'Page ' + pageNo, { x: 12.305, y: 6.926, w: 1.006, h: 0.303, fontSize: 12 });
  rect(s, 13.200, 6.833, 0.133, 0.489, PINK);
}
function newSlide(pageNo) {
  const s = pptx.addSlide();
  s.background = { color: BG };
  chrome(s, pageNo);
  return s;
}
/** stand-in for a photograph that the original deck embedded */
function photo(s, x, y, w, h, label) {
  rect(s, x, y, w, h, 'D9D9D9', { line: { color: 'BFBFBF', width: 1 } });
  text(s, label, {
    x, y: y + h / 2 - 0.2, w, h: 0.4, align: 'center', fontSize: 12, color: '7F7F7F',
  });
}

const LOREM = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor';
const LOREM_LONG = LOREM + ' incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation consectetur adipiscing elit.';
const LOREM_MED = LOREM + ' incididunt ut labore et dolore magna aliqua. ';
const LOREM_SHORT = LOREM + '.';
const LOREM_AENEAN = 'Lorem ipsum dolor sit amet, consectetuer adipiscing. Aenean commodo ligula';
const LOREM_TINY = 'Lorem ipsum dolor sit amet, consectetuer adipiscing. ';
const LOREM_DOLOR = 'Lorem ipsum dolor sit amet, dolor consectetuer adipiscing. ';

/* ============================================================== slide 1 */
function slide1() {
  const s = newSlide(1);
  rect(s, -0.027, 0, 13.361, 7.5, PINK, { fill: { color: PINK, transparency: 50 } });
  roundRect(s, 0.179, 0.207, 12.975, 0.598, WHITE, 50000);
  logoMark(s, 0.540, 0.353, 0.303, BLACK);
  text(s, 'Product Demo', { x: 0.931, y: 0.353, w: 1.684, h: 0.303, fontSize: 12 });
  NAV.forEach(([label, x, w]) => text(s, label, { x, y: 0.353, w, h: 0.303, fontSize: 12 }));
  hamburger(s, 12.486, 0.404, 0.227, BLACK);

  head(s, 'PRODUCT DEMO', { x: 0.540, y: 5.164, w: 8.014, h: 1.212, fontSize: 66, color: WHITE });
  body(s, LOREM + ' incididunt ut labore et dolore magna.',
    { x: 0.603, y: 6.162, w: 5.401, h: 0.608, color: WHITE, align: 'justify' });
  roundRect(s, 0.608, 4.599, 2.983, 0.501, WHITE, 50000);
  text(s, 'Presentation Template', { x: 1.033, y: 4.700, w: 2.133, h: 0.308, fontSize: 12 });
}

/* ============================================================== slide 2 */
const S2_ROWS = [
  { y: 2.917, num: '01', numX: 0.935, numW: 0.625, title: 'DISCOVERING INNOVATION IN EVERY DETAIL' },
  { y: 4.279, num: '02', numX: 0.907, numW: 0.697, title: 'EXPLORING THE FUTURE OF BEAUTY TOGETHER' },
  { y: 5.602, num: '03', numX: 0.914, numW: 0.778, title: 'A JOURNEY TOWARDS CONFIDENCE AND RADIANCE', dx: 0.007 },
];
function slide2() {
  const s = newSlide(2);
  head(s, 'INTRODUCTION TO OUR PRODUCT DEMO', { x: 0.531, y: 1.059, w: 7.608, h: 1.582, fontSize: 44 });
  S2_ROWS.forEach(r => {
    const dx = r.dx || 0;
    roundRect(s, 0.639 + dx, r.y, 12.042, 1.153, PINK, 431, SOFT);
    head(s, r.title, { x: 1.833 + dx, y: r.y + 0.209, w: 3.719, h: 0.707, color: WHITE });
    rect(s, 0.847 + dx, r.y + 0.209, 0.778, 0.707, WHITE);
    head(s, r.num, { x: r.numX, y: r.y + (r.num === '03' ? 0.128 : 0.170), w: r.numW, h: 0.707, fontSize: 36 });
    body(s, LOREM_SHORT, { x: 8.551 + dx, y: r.y + 0.225, w: 3.847, h: 0.608, color: WHITE, align: 'justify' });
  });
}

/* ============================================================== slide 3 */
function slide3() {
  const s = newSlide(3);
  head(s, 'EXPERIENCE THE FUTURE OF SKINCARE', { x: 0.586, y: 3.750, w: 7.011, h: 1.582, fontSize: 44 });
  head(s, 'Your Text Here', { x: 0.642, y: 5.403, w: 3.499, h: 0.438, fontSize: 20 });
  body(s, LOREM_LONG, { x: 0.642, y: 5.861, w: 6.442, h: 0.870, align: 'justify' });
  [['INNOVATIVE FORMULAS, TIMELESS BEAUTY', 4.436, 'atom'],
  ['REDEFINING SKINCARE FOR THE MODERN ERA', 5.543, 'split']].forEach(([t, y, icon]) => {
    rect(s, 8.111, y, 0.847, 0.792, PINK);
    ICON[icon](s, 8.258, y + 0.131, 0.553, WHITE);
    head(s, t, { x: 9.167, y: y + 0.042, w: 3.525, h: 0.707 });
  });
}

/* ============================================================== slide 4 */
function slide4() {
  const s = newSlide(4);
  head(s, 'YOUR BEAUTY, OUR INNOVATION', { x: 6.667, y: 1.236, w: 5.986, h: 1.582, fontSize: 44 });
  body(s, LOREM_LONG.replace(' adipiscing elit.', '.'), { x: 6.694, y: 2.887, w: 5.986, h: 0.870, align: 'justify' });
  roundRect(s, 4.597, 4.438, 5.583, 1.729, WHITE, 431, SOFT);
  roundRect(s, 10.375, 4.201, 2.278, 2.201, PINK, 431, SOFT);
  rect(s, 4.736, 4.582, 1.372, 1.439, PINK);
  ICON.growth(s, 5.017, 4.956, 0.820, WHITE);
  head(s, 'EMPOWERING CONFIDENCE THROUGH SKINCARE TECHNOLOGY', { x: 6.403, y: 4.853, w: 3.486, h: 1.010 });
  head(s, '95%', { x: 10.670, y: 4.329, w: 1.688, h: 0.909, fontSize: 48, color: WHITE, align: 'center' });
  text(s, 'Users Report Visible Skin Improvement', { x: 10.670, y: 5.358, w: 1.688, h: 0.707, fontSize: 12, color: WHITE, align: 'center' });
}

/* ============================================================== slide 5 */
function slide5() {
  const s = newSlide(5);
  rect(s, 6.181, 2.448, 6.572, 4.319, WHITE, SOFT);
  head(s, 'EXPLORING COSMETIC INNOVATIONS', { x: 0.581, y: 4.608, w: 5.128, h: 2.322, fontSize: 44 });
  [['01', 3.079, 'INNOVATIVE TRENDS SHAPING THE FUTURE OF BEAUTY'],
  ['02', 4.914, 'WHERE SCIENCE MEETS CREATIVITY IN SKINCARE']].forEach(([n, y, t]) => {
    rect(s, 6.569, y, 0.903, 0.844, PINK);
    head(s, n, { x: 6.689, y: y + 0.012, w: 0.674, h: 0.707, fontSize: 36, color: WHITE });
    head(s, t, { x: 7.731, y: y - 0.042, w: 4.747, h: 0.707 });
    body(s, 'Lorem ipsum dolor sit amet, sed do consectetur adipiscing elit, sed do eiusmod tempor incididunt.',
      { x: 7.738, y: y + 0.679, w: 4.407, h: 0.608, align: 'justify' });
  });
}

/* ============================================================== slide 6 */
function slide6() {
  const s = newSlide(6);
  head(s, 'MEET OUR TEAM 2025', { x: 0.520, y: 1.269, w: 4.072, h: 1.582, fontSize: 44 });
  body(s, LOREM_MED, { x: 0.620, y: 2.872, w: 3.944, h: 0.870, align: 'justify' });
  rect(s, 0.752, 4.228, 3.687, 2.335, WHITE, SOFT);
  head(s, 'Targets That Have Been Achieved ', { x: 1.079, y: 4.550, w: 3.183, h: 0.707, color: '1F1F1F' });
  s.addText(
    [' Lorem ipsum dolor sit amet', ' Elit, sed do eiusmod tempor ', ' Dolore magna aliqua. ']
      .map(t => ({ text: t, options: { bullet: { characterCode: '2713', indent: 13.5 }, breakLine: true } })),
    { x: 1.079, y: 5.304, w: 3.183, h: 0.870, fontFace: MINOR, fontSize: 12, color: BLACK, valign: 'top', lineSpacingMultiple: 1.3, align: 'justify' });

  [['01', 'ISABELLA', 5.389, 4.847], ['02', 'MIRRECEL', 9.082, 4.851]].forEach(([n, name, x, y]) => {
    rect(s, x, y, 3.125, 0.980, WHITE, SOFT);
    roundRect(s, x, y + 0.110, 0.847, 0.778, PINK, 0);
    head(s, n, { x: x + 0.098, y: y + 0.073, w: 0.720, h: 0.707, fontSize: 36, bold: true, color: WHITE });
    head(s, name, { x: x + 1.004, y: y + 0.145, w: 1.774, h: 0.404, bold: true });
    body(s, 'Our Position', { x: x + 1.012, y: y + 0.474, w: 1.500, h: 0.345, align: 'justify' });
  });
}

/* ============================================================== slide 7 */
function slide7() {
  const s = newSlide(7);
  photo(s, 0.0, 1.972, 7.640, 3.757, '[laptop mockup]');
  photo(s, 5.585, 2.820, 1.783, 4.166, '[phone]');
  head(s, 'MOCKUP DEVICES', { x: 7.745, y: 1.131, w: 4.976, h: 0.841, fontSize: 44, bold: true, charSpacing: -2.87 });
  body(s, LOREM_MED, { x: 7.763, y: 1.950, w: 4.779, h: 0.870, align: 'justify' });
  [['250+ ', 3.255, 'EXCLUSIVE BEAUTY PRODUCTS LAUNCHED', 3.415, 7.259],
  ['120+ ', 4.846, 'DEMO SESSIONS CONDUCTED', 3.089, 7.245]].forEach(([n, y, t, tw, nx]) => {
    roundRect(s, 6.931, y, 5.583, 1.425, WHITE, 431, SOFT);
    head(s, n, { x: nx, y: y + 0.217, w: 1.688, h: 0.909, fontSize: 48, align: 'center' });
    head(s, t, { x: 8.946, y: y + 0.373, w: tw, h: 0.707 });
  });
}

/* ============================================================== slide 8 */
const S8_BARS = [
  { x: 0.909, y: 5.472, h: 1.484, c: PINK, v: '+25.2' },
  { x: 2.000, y: 5.344, h: 1.611, c: BLUE, v: '+32.7' },
  { x: 3.092, y: 4.904, h: 2.051, c: PINK, v: '+37.5' },
  { x: 4.183, y: 4.488, h: 2.467, c: BLUE, v: '+50.9' },
  { x: 5.275, y: 3.782, h: 3.174, c: PINK, v: '+72.4' },
];
const S8_CARDS = [
  { y: 1.729, bg: WHITE, fg: PINK, txt: BLACK },
  { y: 3.380, bg: PINK, fg: WHITE, txt: WHITE },
  { y: 5.009, bg: WHITE, fg: PINK, txt: BLACK },
];
function slide8() {
  const s = newSlide(8);
  // growth arrow sweeping over the bars
  s.addShape(S.custGeom, {
    x: 0.909, y: 2.428, w: 5.078, h: 2.446, fill: { color: SLATE },
    points: [{ x: 5.032, y: 0 }, { x: 4.345, y: 0.17 }, { x: 4.329, y: 0.24 },
    { x: 4.440, y: 0.348 }, { x: 0, y: 2.446, curve: { type: 'cubic', x1: 2.416, y1: 2.370, x2: 0, y2: 2.446 } },
    { x: 4.748, y: 0.654, curve: { type: 'cubic', x1: 2.708, y1: 2.442, x2: 4.587, y2: 0.798 } },
    { x: 4.846, y: 0.749 }, { x: 4.912, y: 0.732 }, { x: 5.072, y: 0.049 }, { x: 5.036, y: 0 }, { close: true }],
  });
  S8_BARS.forEach(b => {
    roundRect(s, b.x, b.y, 0.956, b.h, b.c, 9587);
    head(s, b.v, { x: b.x, y: 6.283, w: 0.956, h: 0.525, fontSize: 16, bold: true, color: WHITE, align: 'center', lineSpacingMultiple: 1.5, inset: 0.1 });
  });
  roundRect(s, 0.553, 6.879, 6.032, 0.076, SLATE, 50000);
  head(s, 'SECTION INFOGRAPHIC.', { x: 0.551, y: 1.079, w: 4.976, h: 1.443, fontSize: 44, bold: true, lineSpacingMultiple: 0.9 });

  S8_CARDS.forEach(c => {
    roundRect(s, 7.146, c.y, 5.494, 1.425, c.bg, 431, SOFT);
    oval(s, 7.371, c.y + 0.204, 1.025, 1.025, c.fg);
    ICON.handshake(s, 7.596, c.y + 0.429, 0.575, c.bg);
    head(s, 'YOUR TITLE HERE', { x: 8.542, y: c.y + 0.067, w: 2.893, h: 0.615, bold: true, color: c.txt, lineSpacingMultiple: 1.5, inset: 0.1 });
    body(s, LOREM_AENEAN, { x: 8.660, y: c.y + 0.602, w: 3.779, h: 0.608, color: c.txt });
  });
}

/* ============================================================== slide 9 */
const ICEBERG = [
  { x: 4.984, y: 1.781, w: 3.366, h: 1.852, c: ICEBLUE, p: [[0.999, 1], [1, 0.999], [0.997, 0.994], [0.697, 0.239], [0.662, 0.322], [0.5, 0], [0.234, 0.533], [0.205, 0.467], [0.002, 0.994], [0, 0.999], [0, 1]] },
  { x: 6.264, y: 1.781, w: 2.085, h: 1.852, c: MIDBLUE, p: [[0.175, 0.513], [0, 1], [0.999, 1], [1, 0.999], [0.995, 0.994], [0.51, 0.239], [0.454, 0.322], [0.193, 0], [0.031, 0.588]] },
  { x: 6.329, y: 2.527, w: 0.166, h: 0.342, c: MIDBLUE, p: [[1, 0.775], [0.638, 0], [0, 1]] },
  { x: 6.264, y: 2.731, w: 1.009, h: 0.902, c: PALEBLUE, p: [[0.362, 0], [0, 1], [1, 1]] },
  { x: 6.530, y: 1.781, w: 1.480, h: 1.852, c: PALEBLUE, p: [[0.097, 0.004], [0.093, 0], [0.089, 0.007], [0, 0.237], [0.695, 1], [1, 1]] },
  { x: 4.984, y: 3.633, w: 3.364, h: 3.495, c: BLUE, p: [[0.941, 0.22], [0.992, 0.016], [1, 0], [0, 0], [0.008, 0.016], [0.06, 0.22], [0.096, 0.192], [0.149, 0.297], [0.201, 0.504], [0.237, 0.475], [0.312, 0.624], [0.36, 0.812], [0.394, 0.786], [0.5, 1], [0.607, 0.786], [0.641, 0.812], [0.688, 0.624], [0.762, 0.475], [0.8, 0.504], [0.852, 0.297], [0.904, 0.192]] },
  { x: 6.667, y: 3.633, w: 1.681, h: 3.495, c: DKBLUE, p: [[0, 1], [0.214, 0.786], [0.281, 0.812], [0.376, 0.624], [0.525, 0.475], [0.599, 0.504], [0.704, 0.297], [0.808, 0.192], [0.881, 0.22], [0.984, 0.016], [1, 0], [0.26, 0]] },
  { x: 6.150, y: 3.633, w: 0.686, h: 1.997, c: DKBLUE, p: [[0.399, 0.102], [0.109, 0.306], [0.692, 1], [1, 0], [0.166, 0], [0, 0.102]] },
  { x: 4.984, y: 3.633, w: 1.281, h: 1.243, c: MIDBLUE, p: [[0, 0], [0.022, 0.045], [0.158, 0.619], [0.253, 0.539], [0.515, 0.3], [0.386, 0.826], [0.39, 0.834], [0.431, 1], [1, 0]] },
  { x: 5.655, y: 3.837, w: 0.769, h: 1.551, c: MIDBLUE, p: [[0.003, 1], [1, 0], [0.789, 0], [0, 0.988]] },
];
const S9_NODES = [   // side, circle y, icon, circle colour, connector colour + span
  { side: 'l', y: 2.152, icon: 'trophy', c: BLUE, ty: 1.881, cy: 2.517, cx: 4.813, cw: 1.321, cc: BLUE },
  { side: 'r', y: 2.152, icon: 'medal', c: PINK, ty: 1.993, cy: 2.517, cx: 7.203, cw: 1.321, cc: PINK },
  { side: 'l', y: 4.086, icon: 'dumbbell', c: PINK, ty: 3.859, cy: 4.499, cx: 4.813, cw: 1.054, cc: PINK },
  { side: 'r', y: 4.086, icon: 'rocket', c: BLUE, ty: 3.870, cy: 4.499, cx: 7.483, cw: 1.054, cc: BLUE },
  { side: 'l', y: 5.907, icon: 'briefcase', c: BLUE, ty: 5.736, cy: 6.299, cx: 4.802, cw: 1.598, cc: LTBLUE },
  { side: 'r', y: 5.907, icon: 'target', c: PINK, ty: 5.748, cy: 6.299, cx: 6.928, cw: 1.609, cc: GREEN },
];
function slide9() {
  const s = newSlide(9);
  ICEBERG.forEach(f => poly(s, f.x, f.y, f.w, f.h, f.c, f.p));
  S9_NODES.forEach(n => {
    line(s, n.cx, n.cy, n.cw, 0, n.cc, 1);
    const cx = n.side === 'l' ? 3.737 : 8.798;
    oval(s, cx, n.y, 0.799, 0.798, n.c);
    ICON[n.icon](s, cx + 0.175, n.y + 0.175, 0.448, WHITE);
    const tx = n.side === 'l' ? 0.60 : 9.75;
    const align = n.side === 'l' ? 'right' : 'left';
    head(s, 'YOUR TITLE HERE', { x: tx + 0.06, y: n.ty, w: 2.893, h: 0.615, bold: true, align, lineSpacingMultiple: 1.5, inset: 0.1 });
    body(s, LOREM_TINY, { x: tx - 0.03, y: n.ty + 0.535, w: 2.893, h: 0.608, align });
  });
  head(s, 'SECTION INFOGRAPHIC.', { x: 2.805, y: 0.828, w: 8.247, h: 0.776, fontSize: 44, bold: true, align: 'center', lineSpacingMultiple: 0.9 });
}

/* ============================================================= slide 10 */
/* "which way?" scene: signpost, businessman seen from behind, floating icons */
const SIGN_ARROW = [[0, 0], [0.79, 0], [1, 0.5], [0.79, 1], [0, 1]];
const PERSON = [ // [x, y, w, h, colour, outline]
  { x: 6.870, y: 3.918, w: 0.680, h: 0.149, c: GREEN, p: [[0.08, 0.59], [0.33, 0.43], [0.58, 0], [0.87, 0.14], [1, 0.55], [0.85, 1], [0.64, 0.67], [0.36, 0.73], [0.29, 0.92], [0.11, 0.86], [0, 0.59]] },
  { x: 6.550, y: 3.787, w: 0.214, h: 0.151, c: SKIN, p: [[0, 0.13], [0, 1], [1, 1], [1, 0]] },
  { x: 6.442, y: 3.647, w: 0.116, h: 0.143, c: SKIN, p: [[0, 0], [1, 0], [1, 1], [0, 1]] },
  { x: 6.769, y: 3.647, w: 0.116, h: 0.143, c: SKIN, p: [[0, 0], [1, 0], [1, 1], [0, 1]] },
  { x: 6.459, y: 3.325, w: 0.423, h: 0.523, c: INK, p: [[0.34, 0.16], [0.14, 0.22], [0, 0.52], [0.13, 0.86], [0.29, 1], [0.66, 1], [0.79, 0.9], [0.94, 0.63], [0.94, 0.44], [1, 0.46], [1, 0.32], [0.81, 0.2], [0.9, 0.17], [0.77, 0.1], [0.56, 0.12], [0.53, 0]] },
  { x: 6.229, y: 4.764, w: 0.716, h: 2.028, c: INK, p: [[0.1, 0.03], [0.07, 0.54], [0, 0.62], [0.1, 0.94], [0.11, 0.99], [0.32, 0.99], [0.34, 0.7], [0.48, 0.23], [0.54, 0.22], [0.63, 0.61], [0.74, 0.96], [1, 0.93], [0.95, 0.89], [0.99, 0.17], [0.9, 0], [0.12, 0.01]] },
  { x: 5.938, y: 3.933, w: 1.352, h: 0.942, c: WHITE, p: [[0.5, 0], [0.39, 0.04], [0.14, 0.22], [0, 0.64], [0.1, 0.72], [0.24, 0.49], [0.28, 0.8], [0.28, 1], [0.72, 0.98], [0.72, 0.78], [0.77, 0.43], [0.91, 0.71], [1, 0.61], [0.85, 0.13], [0.64, 0.02]] },
  { x: 6.459, y: 3.885, w: 0.376, h: 0.167, c: WHITE, p: [[0, 1], [1, 0.97], [0.91, 0.21], [0.75, 0], [0.26, 0], [0.07, 0.34]] },
  { x: 5.921, y: 4.601, w: 0.547, h: 0.430, c: SKIN, p: [[0, 0.05], [0.27, 0.44], [0.59, 0.82], [0.59, 0.91], [0.68, 1], [0.72, 0.81], [0.95, 0.75], [0.8, 0.61], [0.64, 0.61], [0.35, 0]] },
  { x: 5.894, y: 4.447, w: 0.248, h: 0.241, c: WHITE, p: [[0.39, 0], [0.21, 0.07], [0, 0.39], [0.12, 0.88], [0.73, 1], [0.94, 0.8], [1, 0.37], [0.78, 0.53], [0.44, 0.46], [0.29, 0.34]] },
  { x: 6.780, y: 4.601, w: 0.548, h: 0.439, c: SKIN, p: [[1, 0.05], [0.73, 0.43], [0.41, 0.8], [0.41, 0.89], [0.29, 1], [0.28, 0.79], [0.05, 0.74], [0.2, 0.6], [0.36, 0.6], [0.65, 0]] },
  { x: 7.115, y: 4.456, w: 0.238, h: 0.233, c: WHITE, p: [[0.74, 0], [1, 0.36], [0.88, 0.88], [0.24, 1], [0.07, 0.86], [0, 0.36], [0.2, 0.51], [0.57, 0.5], [0.66, 0.19]] },
  { x: 6.291, y: 6.669, w: 0.165, h: 0.123, c: '0D0D0D', p: [[0.88, 0.9], [1, 0.83], [0.95, 0], [0.25, 0.19], [0.05, 0], [0, 0.83], [0.11, 1], [0.74, 1]] },
  { x: 6.745, y: 6.663, w: 0.270, h: 0.140, c: '0D0D0D', p: [[0.71, 0], [0.89, 0.1], [1, 0.31], [0.98, 0.83], [0.53, 0.9], [0.09, 1], [0, 0.71], [0.04, 0.08]] },
];
const TUFTS = [  // [x, y, w, h, colour, leaf-pair rows]
  { x: 8.318, y: 5.986, w: 0.367, h: 0.806, c: LTBLUE, rows: 4 },
  { x: 4.322, y: 5.922, w: 0.395, h: 0.870, c: PINK, rows: 4 },
  { x: 4.735, y: 6.276, w: 0.395, h: 0.513, c: BLUE, rows: 2 },
  { x: 7.920, y: 6.314, w: 0.363, h: 0.475, c: 'F27530', rows: 2 },
  { x: 8.706, y: 6.314, w: 0.364, h: 0.475, c: GREEN, rows: 2 },
];
const LEAF = [[0, 0.5], [0.25, 0.05], [0.75, 0], [1, 0.28], [0.75, 0.9], [0.3, 1]];
function crossroadsScene(s) {
  s.addShape(S.custGeom, {                                        // soft grey backdrop
    x: 4.712, y: 3.781, w: 3.741, h: 3.017, fill: { color: '808080', transparency: 90 },
    points: [{ x: 2.040, y: 0 },
    { x: 3.558, y: 0.856, curve: { type: 'cubic', x1: 2.362, y1: 0.022, x2: 2.937, y2: 0.264 } },
    { x: 3.741, y: 1.628, curve: { type: 'cubic', x1: 3.705, y1: 1.045, x2: 3.741, y2: 1.230 } },
    { x: 2.956, y: 2.870, curve: { type: 'cubic', x1: 3.741, y1: 2.032, x2: 3.192, y2: 2.531 } },
    { x: 2.681, y: 3.017 }, { x: 0.560, y: 3.017 },
    { x: 0.006, y: 2.224, curve: { type: 'cubic', x1: 0.206, y1: 2.770, x2: 0.046, y2: 2.533 } },
    { x: 0.464, y: 1.325, curve: { type: 'cubic', x1: 0, y1: 1.794, x2: 0.200, y2: 1.508 } },
    { x: 1.261, y: 0.298, curve: { type: 'cubic', x1: 0.971, y1: 0.828, x2: 1.053, y2: 0.520 } },
    { x: 2.040, y: 0, curve: { type: 'cubic', x1: 1.465, y1: 0.078, x2: 1.751, y2: 0 } }, { close: true }],
  });
  roundRect(s, 6.588, 1.447, 0.122, 5.340, 'D0D4DA', 50000);       // post
  poly(s, 6.709, 1.634, 1.570, 0.553, PINK, SIGN_ARROW);
  poly(s, 6.952, 1.739, 1.087, 0.343, WHITE, [[0, 0.22], [0.72, 0.22], [0.72, 0], [1, 0.51], [0.72, 1], [0.72, 0.78], [0, 0.78]]);
  poly(s, 6.709, 2.412, 0.875, 0.776, LTBLUE, SIGN_ARROW);
  poly(s, 6.810, 2.589, 0.595, 0.422, WHITE, [[0, 0.26], [0.68, 0.26], [0.68, 0], [1, 0.5], [0.68, 1], [0.68, 0.79], [0, 0.73]]);
  poly(s, 5.716, 2.533, 0.875, 0.776, BLUE, SIGN_ARROW.map(p => [1 - p[0], p[1]]));
  poly(s, 5.817, 2.710, 0.595, 0.422, WHITE, [[1, 0.26], [0.32, 0.26], [0.32, 0], [0, 0.5], [0.32, 1], [0.32, 0.79], [1, 0.73]]);
  PERSON.forEach(f => poly(s, f.x, f.y, f.w, f.h, f.c, f.p));
  ICON.clock(s, 7.867, 3.584, 0.584, SLATE);
  ICON.calendar(s, 4.805, 3.407, 0.587, SLATE);
  ICON.list(s, 4.874, 4.247, 0.449, SLATE);
  ICON.bars(s, 7.775, 4.411, 0.607, SLATE);
  ICON.target(s, 4.849, 2.723, 0.497, SLATE);
  ICON.dollar(s, 7.860, 5.248, 0.418, SLATE);
  TUFTS.forEach(t => {                                             // grass tufts
    rect(s, t.x + t.w * 0.48, t.y + t.h * 0.24, t.w * 0.03, t.h * 0.76, t.c);
    poly(s, t.x + t.w * 0.4, t.y, t.w * 0.4, t.h * 0.26, t.c, LEAF);
    for (let r = 0; r < t.rows; r++) {
      const ly = t.y + t.h - (r + 1) * t.h * 0.2;
      poly(s, t.x, ly, t.w * 0.5, t.h * 0.17, t.c, LEAF.map(p => [1 - p[0], p[1]]));
      poly(s, t.x + t.w * 0.5, ly, t.w * 0.5, t.h * 0.17, t.c, LEAF);
    }
  });
  rect(s, 4.292, 6.782, 4.976, 0.013, INK);                        // ground line
}
function slide10() {
  const s = newSlide(10);
  crossroadsScene(s);
  [{ x: 10.757, y: 1.731, c: PINK, icon: 'target', tx: 9.851, ty: 2.575, bx: 9.949, by: 3.121 },
  { x: 1.699, y: 4.254, c: BLUE, icon: 'bulb', tx: 0.826, ty: 5.107, bx: 0.923, by: 5.653 },
  { x: 10.757, y: 4.254, c: LTBLUE, icon: 'screen', tx: 9.851, ty: 5.107, bx: 9.949, by: 5.653 }]
    .forEach(n => {
      oval(s, n.x, n.y, 0.703, 0.703, n.c);
      ICON[n.icon](s, n.x + 0.154, n.y + 0.154, 0.394, WHITE);
      head(s, 'YOUR TITLE HERE', { x: n.tx, y: n.ty, w: 2.532, h: 0.615, bold: true, align: 'center', lineSpacingMultiple: 1.5, inset: 0.1 });
      body(s, LOREM_DOLOR, { x: n.bx, y: n.by, w: 2.337, h: 0.870, align: 'center' });
    });
  head(s, 'SECTION INFOGRAPHIC.', { x: 0.495, y: 1.076, w: 4.976, h: 1.443, fontSize: 44, bold: true, lineSpacingMultiple: 0.9 });
}

/* ============================================================= slide 11 */
/* character sprinting with a trophy, drawn back-to-front */
const RUNNER = [
  { x: 2.418, y: 4.902, w: 0.331, h: 0.487, c: YELLOW, p: [[0.46, 0], [0, 0.48], [0.22, 0.77], [0.69, 1], [0.73, 0.82], [1, 0.59], [0.57, 0.42]] },
  { x: 2.647, y: 5.144, w: 0.504, h: 0.360, c: SKIN, p: [[0.89, 0], [0.58, 0.26], [0.36, 0.24], [0.2, 0.12], [0.03, 0.44], [0, 0.68], [0.5, 1], [1, 0.19]] },
  { x: 3.062, y: 5.006, w: 0.252, h: 0.244, c: SKIN, p: [[0.17, 0.74], [0, 0.43], [0.33, 0.11], [0.41, 0.25], [0.67, 0], [1, 0.44], [0.63, 1], [0.25, 0.87]] },
  { x: 0.885, y: 6.698, w: 0.173, h: 0.374, c: '0D0D0D', p: [[0.12, 0], [0, 0.06], [0.32, 0.72], [0.87, 0.99], [1, 0.99], [0.68, 0.8], [0.11, 0.11]] },
  { x: 0.902, y: 6.662, w: 0.215, h: 0.404, c: INK, p: [[0.58, 0.09], [0.3, 0], [0.01, 0.09], [0.04, 0.36], [0.47, 0.83], [0.74, 0.98], [0.82, 0.82], [1, 0.39]] },
  { x: 0.992, y: 6.614, w: 0.139, h: 0.151, c: SKIN, p: [[0, 0.2], [0.6, 0], [1, 0.5], [0.7, 1], [0.1, 0.9]] },
  { x: 3.032, y: 6.109, w: 0.273, h: 0.341, c: '0D0D0D', p: [[0, 0.97], [0.12, 0.99], [0.84, 0.47], [1, 0.03], [0.95, 0], [0.82, 0.25], [0.36, 0.78], [0.15, 0.91]] },
  { x: 2.966, y: 6.107, w: 0.325, h: 0.334, c: INK, p: [[0, 0.66], [0, 0.89], [0.2, 1], [0.51, 0.8], [0.89, 0.26], [0.97, 0.01], [0.76, 0.05], [0.21, 0.22]] },
  { x: 2.874, y: 6.195, w: 0.163, h: 0.140, c: SKIN, p: [[0, 0.3], [0.5, 0], [1, 0.4], [0.8, 1], [0.15, 0.95]] },
  { x: 1.038, y: 5.684, w: 0.987, h: 1.037, c: INK, p: [[0.66, 0], [0.55, 0.38], [0.47, 0.54], [0.34, 0.56], [0.18, 0.65], [0, 0.91], [0.1, 1], [0.6, 0.67], [1, 0.23], [0.98, 0.02]] },
  { x: 1.680, y: 5.622, w: 1.285, h: 0.666, c: INK, p: [[0.03, 0], [0, 0.17], [0.03, 0.62], [0.61, 0.47], [0.65, 0.37], [0.62, 0.66], [0.94, 1], [1, 0.85], [0.78, 0.14], [0.72, 0.02], [0.53, 0.08], [0.3, 0.13]] },
  { x: 1.616, y: 4.767, w: 0.955, h: 1.017, c: YELLOW, p: [[0.41, 0.01], [0.12, 0.07], [0, 0.17], [0.13, 0.33], [0.3, 0.25], [0.15, 0.69], [0.06, 0.92], [0.36, 1], [0.53, 0.91], [0.95, 0.35], [1, 0.13], [0.66, 0]] },
  { x: 1.440, y: 4.939, w: 0.394, h: 0.718, c: SKIN, p: [[0.95, 0.86], [0.71, 0.75], [0.43, 0.4], [0.77, 0.22], [0.45, 0], [0.12, 0.16], [0, 0.42], [0.57, 0.79], [0.54, 0.88], [0.73, 1], [1, 0.92]] },
  { x: 2.153, y: 4.641, w: 0.222, h: 0.284, c: SKIN, p: [[0.44, 0], [0.03, 0.65], [0, 0.69], [0.15, 0.86], [0.39, 1], [0.8, 0.9], [0.89, 0.81], [1, 0.23]] },
  { x: 2.177, y: 4.430, w: 0.309, h: 0.329, c: SKIN, p: [[0.75, 0], [0.93, 0.38], [0.95, 0.79], [1, 0.92], [0.5, 1], [0.3, 0.88], [0.18, 0.68], [0, 0.37], [0.08, 0.12]] },
  { x: 2.077, y: 4.292, w: 0.419, h: 0.388, c: INK, p: [[0.43, 0.81], [0.38, 0.59], [0.53, 0.53], [0.86, 0.45], [1, 0.2], [0.94, 0.05], [0.78, 0.02], [0.7, 0.12], [0.31, 0.19], [0.13, 0.46], [0, 0.54], [0.16, 0.83], [0.1, 0.89], [0.29, 1], [0.4, 0.98]] },
  { x: 3.006, y: 4.767, w: 0.375, h: 0.421, c: PALEPINK, p: [[0.83, 0.42], [0.65, 0], [0, 0.46], [0.15, 0.77], [0.49, 0.89], [0.84, 1], [1, 0.76]] },
  { x: 2.776, y: 4.005, w: 0.818, h: 0.950, c: PINK, p: [[0, 0.05], [0, 0.78], [0.82, 1], [1, 0.42], [0.99, 0]] },
  { x: 2.917, y: 5.248, w: 0.535, h: 0.064, c: PALEPINK, p: [[0, 0], [1, 0], [1, 1], [0, 1]] },
  { x: 2.917, y: 5.175, w: 0.535, h: 0.073, c: PINK, p: [[1, 1], [0, 1], [0.25, 0], [0.75, 0]] },
  { x: 2.684, y: 3.903, w: 1.022, h: 0.199, c: PALEPINK, p: [[0.05, 0.28], [0, 0.4], [0, 0.91], [0.05, 0.99], [0.95, 0.74], [1, 0.63], [1, 0.09], [0.95, 0.01]] },
  { x: 3.302, y: 4.460, w: 0.482, h: 0.441, c: PALEPINK, p: [[0.09, 1], [0, 0.93], [0.03, 0.84], [0.73, 0.76], [0.8, 0.32], [0.68, 0.17], [0.5, 0.24], [0.49, 0.39], [0.31, 0.38], [0.28, 0.27], [0.33, 0.16], [0.55, 0.01], [0.71, 0], [0.89, 0.1], [1, 0.19], [0.93, 0.62], [0.77, 0.79], [0.3, 0.98]] },
  { x: 2.591, y: 4.460, w: 0.483, h: 0.441, c: PALEPINK, p: [[0.91, 1], [1, 0.88], [0.91, 0.83], [0.27, 0.76], [0.2, 0.32], [0.31, 0.17], [0.5, 0.24], [0.51, 0.39], [0.69, 0.38], [0.72, 0.27], [0.67, 0.16], [0.45, 0.01], [0.28, 0], [0.11, 0.1], [0, 0.19], [0.07, 0.62], [0.23, 0.79], [0.7, 0.98]] },
  { x: 2.873, y: 4.229, w: 0.290, h: 0.545, c: WHITE, p: [[0.28, 0], [0.72, 0], [0.72, 0.1], [1, 0.35], [0.6, 0.62], [0.6, 0.9], [0.4, 0.9], [0.4, 0.62], [0, 0.35], [0.28, 0.1]] },
  { x: 3.137, y: 4.983, w: 0.181, h: 0.231, c: SKIN, p: [[0.2, 0], [0.9, 0.15], [1, 0.6], [0.6, 1], [0, 0.7]] },
];
function winnerScene(s) {
  s.addShape(S.custGeom, {
    x: 0.526, y: 2.408, w: 5.057, h: 4.836, fill: { color: '808080', transparency: 90 },
    points: [{ x: 4.265, y: 4.028 },
    { x: 4.011, y: 4.242, curve: { type: 'cubic', x1: 4.171, y1: 4.126, x2: 4.081, y2: 4.191 } },
    { x: 1.989, y: 4.621, curve: { type: 'cubic', x1: 3.180, y1: 4.836, x2: 2.197, y2: 4.662 } },
    { x: 0.653, y: 3.931, curve: { type: 'cubic', x1: 1.679, y1: 4.558, x2: 1.084, y2: 4.437 } },
    { x: 0.608, y: 1.162, curve: { type: 'cubic', x1: 0, y1: 3.165, x2: 0.101, y2: 1.959 } },
    { x: 2.183, y: 0.079, curve: { type: 'cubic', x1: 0.701, y1: 1.017, x2: 1.189, y2: 0.208 } },
    { x: 3.992, y: 0.583, curve: { type: 'cubic', x1: 2.784, y1: 0, x2: 3.484, y2: 0.158 } },
    { x: 4.265, y: 4.028, curve: { type: 'cubic', x1: 5.057, y1: 1.474, x2: 4.982, y2: 3.274 } }, { close: true }],
  });
  // dartboard, arrow and coins
  oval(s, 3.148, 2.934, 1.853, 1.853, LTBLUE, LIFT);
  [[3.201, 2.987, 1.746, WHITE], [3.511, 3.295, 1.128, LTBLUE],
  [3.556, 3.340, 1.038, WHITE], [3.856, 3.641, 0.438, LTBLUE]].forEach(([x, y, d, c]) => oval(s, x, y, d, d, c));
  poly(s, 4.252, 2.486, 0.969, 1.179, SLATE, [[0.97, 0], [1, 0.02], [0.03, 1], [0, 0.98]]);
  poly(s, 4.958, 2.510, 0.404, 0.325, GREEN, [[0.65, 0], [1, 0], [0.43, 0.87], [0.22, 0.99], [0, 1]]);
  poly(s, 4.907, 2.346, 0.281, 0.464, GREEN, [[1, 0.3], [0.9, 0], [0.08, 0.61], [0, 0.81], [0.07, 1]]);
  poly(s, 4.067, 3.553, 0.293, 0.341, YELLOW, [[0.67, 0.05], [0.72, 0], [1, 0.2], [0.95, 0.25], [0.05, 1], [0.02, 0.93]]);
  [[2.505, 2.812, 0.399], [2.834, 2.163, 0.696], [3.102, 2.945, 0.717], [3.679, 2.397, 0.524]]
    .forEach(([x, y, d]) => {
      oval(s, x, y, d, d, WHITE, LIFT);
      oval(s, x + d * 0.07, y + d * 0.07, d * 0.86, d * 0.86, PINK);
      text(s, '$', { x, y: y + d * 0.1, w: d, h: d * 0.8, align: 'center', valign: 'middle', fontSize: d * 46, bold: true, color: WHITE });
    });
  // stylised foliage behind the runner
  poly(s, 1.081, 4.007, 1.010, 2.653, 'FDE083', [[1, 1], [0.58, 0.89], [0.55, 0.44], [0.66, 0.13], [0.02, 0], [0.12, 0.2], [0.06, 0.51], [0.21, 0.93]]);
  poly(s, 1.211, 4.235, 0.534, 2.222, 'FEF5D6', [[1, 1], [0.55, 0.8], [0.5, 0.3], [0.7, 0], [0, 0.35], [0.15, 0.75]]);
  poly(s, 0.742, 4.108, 1.009, 2.409, PALEPINK, [[0, 0.01], [0.01, 0.81], [0.97, 1], [0.99, 0.9], [0.93, 0.63], [0.18, 0.13]]);
  poly(s, 0.905, 4.677, 0.846, 1.797, 'FDE9F4', [[0.05, 0.05], [0.06, 0.8], [1, 1], [0.98, 0.85], [0.9, 0.6], [0.2, 0.18]]);
  // blue zig-zag "V" arrow
  poly(s, 3.708, 3.429, 1.773, 3.355, BLUE, [[0, 0.9], [0.51, 0.38], [0.39, 0.36], [0.93, 0.02], [1, 0.03], [0.96, 0.39], [0.81, 0.4], [0.3, 0.99], [0.06, 1]]);
  poly(s, 2.962, 4.990, 1.295, 1.795, MIDBLUE, [[0, 0.14], [0.52, 0.97], [0.66, 1], [0.97, 0.99], [0.94, 0.93], [0.28, 0]]);
  poly(s, 1.694, 4.730, 1.717, 2.045, BLUE, [[1, 0.21], [0.88, 0.22], [0.29, 0.97], [0.06, 1], [0, 0.85], [0.51, 0.2], [0.77, 0], [0.96, 0.12]]);
  poly(s, 0.740, 5.410, 1.531, 1.379, MIDBLUE, [[0.07, 0.07], [0.26, 0], [0.35, 0.04], [0.73, 0.78], [0.85, 0.95], [0.99, 0.88], [0.88, 1], [0.63, 1], [0.54, 0.95], [0.17, 0.39], [0, 0.19], [0.01, 0.1]]);
  poly(s, 4.556, 3.533, 0.858, 1.061, MIDBLUE, [[0.01, 1], [0, 0.99], [0.93, 0.02], [1, 0.04], [0.95, 0.03]]);
  poly(s, 2.753, 5.020, 0.540, 0.655, MIDBLUE, [[0.02, 1], [0.01, 0.98], [0.76, 0.05], [0.93, 0.01], [0.79, 0.07]]);
  ICON.gear(s, 0.670, 2.700, 0.565, YELLOW);
  ICON.gear(s, 1.212, 3.162, 1.016, PINK);
  // orange / blue leaf pair at the right
  poly(s, 4.129, 5.202, 0.957, 1.520, 'F7AC83', [[0, 0.98], [0.06, 0.84], [0.32, 0.31], [0.73, 0], [0.87, 0.2], [1, 0.68], [0.02, 1]]);
  poly(s, 3.834, 5.306, 1.655, 1.546, '9DC3E6', [[0.99, 0.02], [0.89, 0.02], [0.64, 0.1], [0.3, 0.84], [0.15, 0.94], [0, 0.96], [0.15, 1], [0.73, 0.7], [0.88, 0.15]]);
  poly(s, 4.253, 5.464, 1.035, 1.245, 'DEEBF7', [[0.02, 1], [0.85, 0.44], [0.96, 0], [1, 0.02], [0.97, 0.12], [0.72, 0.46], [0.48, 0.71]]);
  RUNNER.forEach(f => poly(s, f.x, f.y, f.w, f.h, f.c, f.p));
}
function slide11() {
  const s = newSlide(11);
  winnerScene(s);
  [{ x: 5.942, cx: 7.179, val: '$8m', vx: 6.634, tx: 6.250, c: PINK, icon: 'trophy', ix: 7.382, oy: 2.753, vy: 3.402, ty: 4.681 },
  { x: 9.519, cx: 10.651, val: '66.6%', vx: 10.185, tx: 9.827, c: LTBLUE, icon: 'target', ix: 10.827, oy: 2.760, vy: 3.430, ty: 4.699 }]
    .forEach(cd => {
      roundRect(s, cd.x, cd.oy - 0.252, 3.288, 3.818, WHITE, 431, SOFT);
      oval(s, cd.cx, cd.oy, 0.802, 0.802, cd.c);
      ICON[cd.icon](s, cd.ix, cd.oy + 0.176, 0.449, WHITE);
      head(s, cd.val, { x: cd.vx, y: cd.vy, w: 1.892, h: 1.302, fontSize: 48, bold: true, align: 'center', lineSpacingMultiple: 1.5, inset: 0.1 });
      body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore.',
        { x: cd.tx, y: cd.ty, w: 2.672, h: 1.133, align: 'center' });
    });
  head(s, 'SECTION INFOGRAPHIC.', { x: 2.805, y: 0.939, w: 8.247, h: 0.776, fontSize: 44, bold: true, align: 'center', lineSpacingMultiple: 0.9 });
}

/* ============================================================= slide 12 */
const S12_STEPS = [
  { y: 1.016, c: PINK, name: 'PROJECT ONE', num: '01', icon: 'people', numC: PINK },
  { y: 2.611, c: BLUE, name: 'PROJECT TWO', num: '02', icon: 'bars', numC: BLUE },
  { y: 4.205, c: PINK, name: 'PROJECT THREE', num: '03', icon: 'pie', numC: PINK },
];
function slide12() {
  const s = newSlide(12);
  S12_STEPS.forEach((st, i) => {
    roundRect(s, 2.815, st.y, 3.904, 1.432, st.c, 50000);
    oval(s, 5.442, st.y + 0.179, 1.084, 1.073, WHITE, LIFT);
    ICON[st.icon](s, 5.741, st.y + 0.468, 0.486, st.c);
    head(s, st.name, { x: 3.038, y: st.y + 0.179, w: 2.582, h: 0.337, fontSize: 14, color: WHITE, align: 'center' });
    body(s, 'A wonderful serenity has taken possession of my entire soul.',
      { x: 3.271, y: st.y + 0.516, w: 2.117, h: 0.707, color: WHITE, align: 'center', lineSpacingMultiple: 1 });
    head(s, 'STEP', { x: 1.211, y: st.y + 0.063 + i * 0.0075, w: 1.283, h: 0.438, fontSize: 20, color: INK, align: 'center' });
    head(s, st.num, { x: 1.211, y: st.y + 0.282 + i * 0.0075, w: 1.283, h: 1.111, fontSize: 60, color: st.numC, align: 'center' });
  });
  // hub
  oval(s, 9.560, 1.983, 2.562, 2.562, 'F2F2F2');
  oval(s, 9.846, 2.279, 2.140, 2.118, WHITE, LIFT);
  head(s, 'TIMELINE INFOGRAPHIC.', { x: 9.764, y: 2.927, w: 2.290, h: 0.707, align: 'center' });
  // dashed leaders + end dots
  line(s, 6.718, 1.702, 1.590, 0.031, PINK, 2, 'dash', { flipV: true });
  line(s, 8.309, 1.702, 1.627, 0.656, PINK, 2, 'dash');
  oval(s, 9.772, 2.243, 0.204, 0.204, PINK);
  line(s, 6.718, 3.264, 2.842, 0.063, BLUE, 2, 'dash', { flipV: true });
  oval(s, 9.458, 3.149, 0.204, 0.204, BLUE);
  line(s, 6.718, 4.949, 1.631, 0.001, PINK, 2, 'dash');
  line(s, 8.424, 4.205, 1.523, 0.697, PINK, 2, 'dash', { flipV: true });
  oval(s, 9.874, 4.082, 0.204, 0.204, PINK);
  // footnote
  head(s, 'WHAT’S PROBLEM', { x: 9.662, y: 5.198, w: 2.460, h: 0.385, fontSize: 16, bold: true, valign: 'middle', lineSpacingMultiple: 1.1 });
  oval(s, 9.233, 5.275, 0.225, 0.232, WHITE);
  s.addShape(S.custGeom, {          // small ">" tick, an open stroked path
    x: 9.322, y: 5.344, w: 0.081, h: 0.162, fill: { type: 'none' },
    line: { color: '8A9AA8', width: 1.25 },
    points: [{ x: 0, y: 0 }, { x: 0.081, y: 0.081 }, { x: 0, y: 0.162 }],
  });
  body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings',
    { x: 9.662, y: 5.659, w: 2.787, h: 0.869 });
}

/* ============================================================= slide 13 */
function slide13() {
  const s = newSlide(13);
  rect(s, 2.264, 3.083, 8.806, 0.884, WHITE, LIFT);
  [['01', 1.551, PINK, 'flower', 2.196], ['02', 4.274, BLUE, 'gear', 4.934],
  ['03', 6.997, PINK, 'bell', 7.657], ['04', 9.720, BLUE, 'pin', 10.358]].forEach(([n, x, c, icon, nx]) => {
    s.addShape(S.chevron, { x, y: 2.500, w: 2.063, h: 2.063, fill: { color: c } });
    oval(s, x + 1.370, 3.870, 0.693, 0.693, WHITE, DROP);
    ICON[icon](s, x + 1.570, 4.070, 0.293, c);
    text(s, n, { x: nx, y: 3.083, w: 1.216, h: 0.871, fontSize: 40, fontFace: 'Montserrat SemiBold', color: WHITE, align: 'center' });
  });
  [0.989, 3.730, 6.569, 9.311].forEach(x => {
    head(s, 'YOUR TITLE HERE', { x, y: 4.952, w: 2.532, h: 0.615, bold: true, align: 'center', lineSpacingMultiple: 1.5, inset: 0.1 });
    body(s, LOREM_DOLOR, { x: x + 0.097, y: 5.498, w: 2.337, h: 0.870, align: 'center' });
  });
  head(s, 'SECTION INFOGRAPHIC.', { x: 2.805, y: 1.078, w: 8.247, h: 0.776, fontSize: 44, bold: true, align: 'center', lineSpacingMultiple: 0.9 });
}

/* ============================================================= slide 14 */
function slide14() {
  const s = newSlide(14);
  head(s, 'SUB TITLE ONE', { x: 1.345, y: 2.183, w: 2.511, h: 0.404, align: 'right' });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ', { x: 0.464, y: 2.574, w: 3.392, h: 0.608, align: 'right' });
  head(s, 'SUB TITLE ONE', { x: 9.478, y: 2.183, w: 2.511, h: 0.404 });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit', { x: 9.478, y: 2.574, w: 3.392, h: 0.608 });

  s.addShape(S.bentArrow, { x: 6.981, y: 3.596, w: 2.788, h: 2.976, rotate: 90, flipH: true, flipV: true, fill: { color: BLUE } });
  s.addShape(S.bentArrow, { x: 3.246, y: 3.596, w: 2.788, h: 2.976, rotate: 270, flipV: true, fill: { color: PINK } });
  head(s, 'ARROWS INFOGRAPHIC.', { x: 4.548, y: 5.009, w: 3.933, h: 0.505, fontSize: 24, charSpacing: -1.5, rotate: 270 });

  [[1.116, '82%', PINK, 1.460, 1.294, 2.391, 1.360], [9.478, '56%', BLUE, 9.822, 9.680, 2.334, 9.722]]
    .forEach(([x, val, c, vx, lx, lw, ox]) => {
      roundRect(s, x, 3.661, 2.740, 3.287, WHITE, 10204, DROP);
      head(s, 'OPTION 1', { x: ox, y: 3.958, w: 2.252, h: 0.484, fontSize: 20, align: 'center', lineSpacingMultiple: 1.2 });
      head(s, val, { x: vx, y: 4.510, w: 2.052, h: 0.841, fontSize: 44, color: c });
      body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ',
        { x: lx, y: 5.601, w: lw, h: 0.815, align: 'center', lineSpacingMultiple: 1.2 });
    });
  head(s, 'SECTION INFOGRAPHIC.', { x: 2.805, y: 0.966, w: 8.247, h: 0.776, fontSize: 44, bold: true, align: 'center', lineSpacingMultiple: 0.9 });
}

/* ============================================================= slide 15 */
function slide15() {
  const s = newSlide(15);
  head(s, 'WHAT OUR CLIENTS SAY?', { x: 0.527, y: 1.119, w: 4.862, h: 1.582, fontSize: 44, bold: true });
  rect(s, 5.056, 1.875, 3.972, 5.625, PINK, SOFT);
  poly(s, 7.667, 3.319, 5.000, 1.861, WHITE,
    [[0.069, 0], [0.088, 0.209], [1, 0.209], [1, 1], [0, 1], [0, 0.209], [0.051, 0.209]], SOFT);
  body(s, LOREM_MED, { x: 7.928, y: 3.997, w: 4.477, h: 0.870, align: 'justify' });
  text(s, '“', { x: 11.639, y: 3.321, w: 1.028, h: 1.717, fontSize: 96, fontFace: 'Anton' });
  rect(s, 0.527, 5.306, 1.875, 1.681, PINK, SOFT);
  rect(s, 2.597, 5.306, 2.069, 1.681, PINK, SOFT);
  head(s, '4,0/5', { x: 0.887, y: 5.601, w: 1.154, h: 0.640, fontSize: 32, bold: true, color: WHITE });
  [0.896, 1.199, 1.497, 1.800].forEach(x => s.addShape(S.star5, { x, y: 6.264, w: 0.229, h: 0.217, fill: { color: WHITE } }));
  head(s, '$1358', { x: 3.055, y: 5.475, w: 1.154, h: 0.640, fontSize: 32, bold: true, color: WHITE });
  body(s, 'Lorem ipsum dolor sit amet, adipiscing', { x: 2.716, y: 6.115, w: 1.833, h: 0.608, color: WHITE });
}

/* ============================================================= slide 16 */
function slide16() {
  const s = newSlide(16);
  rect(s, -0.014, 0, 13.361, 7.5, PINK, { fill: { color: PINK, transparency: 35 } });
  roundRect(s, 0.179, 0.207, 12.975, 0.598, WHITE, 50000);
  logoMark(s, 0.540, 0.353, 0.303, BLACK);
  text(s, 'Product Demo', { x: 0.931, y: 0.353, w: 1.684, h: 0.303, fontSize: 12 });
  NAV.forEach(([label, x, w]) => text(s, label, { x, y: 0.353, w, h: 0.303, fontSize: 12 }));
  hamburger(s, 12.486, 0.404, 0.227, BLACK);

  head(s, 'THANK YOU FOR ATTENTION', { x: 0.691, y: 1.428, w: 8.014, h: 2.322, fontSize: 66, color: WHITE });
  body(s, LOREM + ' incididunt ut labore et dolore magna.',
    { x: 0.691, y: 3.726, w: 5.401, h: 0.608, color: WHITE, align: 'justify' });
  roundRect(s, 0.691, 4.643, 2.212, 0.404, WHITE, 50000);
  text(s, 'Contact Us :', { x: 1.140, y: 4.693, w: 1.314, h: 0.303, fontSize: 12 });

  const TINT = 'F3B4D5';               // PINK @ 35% transparency over the grey master
  [['phone', 0.843, 5.583, 0.333, '+123 456 7890', 1.287, 5.541, 1.545],
  ['map', 3.702, 5.588, 0.333, '12 Your Street Name', 4.230, 5.537, 1.932],
  ['pie', 7.024, 5.546, 0.375, 'www. Your Name.com', 7.625, 5.546, 2.183]]
    .forEach(([icon, ix, iy, id, label, tx, ty, tw]) => {
      ICON[icon](s, ix, iy, id, WHITE, TINT);
      text(s, label, { x: tx, y: ty, w: tw, h: 0.375, fontSize: 12, color: WHITE, lineSpacingMultiple: 1.5 });
    });
}

/* -------------------------------------------------------------- assemble */
[slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8,
  slide9, slide10, slide11, slide12, slide13, slide14, slide15, slide16]
  .forEach(fn => fn());

pptx.writeFile({ fileName: path.join(__dirname, '1545cffa-d998-4d05-a2b5-c6c5adbf4c94_grok_final.pptx') })
  .then(f => console.log('written:', f));
