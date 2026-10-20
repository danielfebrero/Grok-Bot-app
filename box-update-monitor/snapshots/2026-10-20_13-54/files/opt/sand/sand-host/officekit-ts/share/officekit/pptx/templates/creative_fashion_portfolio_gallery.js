/**
 * "MAXX / Creative Powerpoint" deck rebuilt with pptxgenjs.
 * 20 slides, 13.333 x 7.5 in (16:9).
 *
 * Raster photos of the original are replaced by flat grey placeholder shapes
 * that keep the original geometry (diamonds, triangles, parallelograms, ...).
 */
const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ *
 * Palette (theme "Grayscale" + the deck's signature red)
 * ------------------------------------------------------------------ */
const RED = 'C00000';
const WHITE = 'FFFFFF';
const BLACK = '000000';
const OFFWHITE = 'F8F8F8'; // theme bg2
const DARK = '404040'; // tx1 lum 75/25 - headline grey
const NEARBLACK = '1A1A1A'; // tx1 lum 90/10 - body grey
const INK = '191919'; // bg2 lum 10 - body grey (slide 13)
const TITLEBG = '373737'; // accent1 lum 25 - slide 1 background
const TRACK = 'D9D9D9'; // bg1 lum 85 - progress-bar track
const PH_FILL = 'D4D4D4'; // photo placeholder body
const PH_TEXT = '9E9E9E'; // photo placeholder caption
const GREY1 = 'DDDDDD'; // accent1
const GREY2 = 'B2B2B2'; // accent2
const GREY3 = '969696'; // accent3
const GREY4 = '808080'; // accent4
const GREY5 = '5F5F5F'; // accent5
const GREY6 = '4D4D4D'; // accent6

const LATO = 'Lato';
const LATO_BLACK = 'Lato Black';
const MONT = 'Montserrat';
const ROBOTO = 'Roboto';

/* ------------------------------------------------------------------ *
 * Small helpers
 * ------------------------------------------------------------------ */
let S; // pres.ShapeType, assigned in build()

/** Text box. Defaults match PowerPoint (top anchored, Lato). */
function text(slide, runs, o) {
  slide.addText(runs, Object.assign({ fontFace: LATO, valign: 'top', color: BLACK }, o));
}

/** Filled/outlined auto shape. */
function shape(slide, type, o) {
  slide.addShape(type, o);
}

/** Free-form polygon from normalised (0..1) points. */
function poly(slide, pts, o) {
  const points = pts.map(function (p) { return { x: p[0] * o.w, y: p[1] * o.h }; });
  points.push({ close: true });
  slide.addShape(S.custGeom, Object.assign({}, o, { points: points }));
}

/** A "plus" rotated 45deg - the X mark used all over the deck. */
function crossMark(slide, x, y, size, color) {
  const a = 0.36578 * size; // matches the original prstGeom adjust value
  const b = size - a;
  poly(slide, [
    [0, a / size], [a / size, a / size], [a / size, 0], [b / size, 0], [b / size, a / size],
    [1, a / size], [1, b / size], [b / size, b / size], [b / size, 1], [a / size, 1],
    [a / size, b / size], [0, b / size],
  ], { x: x, y: y, w: size, h: size, fill: { color: color }, line: { type: 'none' }, rotate: 45 });
}

/** Grey stand-in for a photo: optional polygon outline, centred caption. */
function photo(slide, o) {
  const base = { x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: PH_FILL }, line: { type: 'none' } };
  if (o.pts) poly(slide, o.pts, base);
  else shape(slide, o.shape || S.rect, Object.assign(base, o.rotate ? { rotate: o.rotate } : {}));
  text(slide, '[image]', {
    x: o.x, y: o.y + o.h / 2 - 0.16, w: o.w, h: 0.32,
    align: 'center', valign: 'middle', fontSize: 10, color: PH_TEXT,
  });
}

/** The rotated "Creative Powerpoint" side tab. `dir` is 90 (right) or 270 (left). */
function sideTab(slide, x, y, dir, runs) {
  text(slide, runs, {
    x: x, y: y, w: 3.126, h: 0.337, rotate: dir, align: 'center', wrap: false,
    fontFace: MONT, fontSize: 14, bold: true, charSpacing: 3,
  });
}

/** Two-line "MAXX / WORD" headline (red first line, grey second). */
function headline(slide, o) {
  text(slide, [
    { text: 'MAXX', options: { color: o.color1 || RED, breakLine: true } },
    { text: o.word, options: { color: o.color2 || DARK } },
  ], {
    x: o.x, y: o.y, w: o.w, h: o.h, align: o.align || 'left',
    fontFace: LATO_BLACK, fontSize: o.fontSize || 48, bold: true,
  });
}

/** Skill/progress bar: grey track, dark fill, right aligned percentage. */
function progressBar(slide, o) {
  text(slide, o.label, {
    x: o.x - 0.013, y: o.y - 0.324, w: o.labelW, h: 0.286,
    fontSize: 11, charSpacing: 3, wrap: false,
  });
  shape(slide, S.roundRect, {
    x: o.x, y: o.y, w: 3.471, h: 0.182, rectRadius: 0.091,
    fill: { color: TRACK }, line: { type: 'none' },
  });
  shape(slide, S.roundRect, {
    x: o.x, y: o.y, w: o.value, h: 0.182, rectRadius: 0.091,
    fill: { color: BLACK }, line: { type: 'none' },
  });
  text(slide, o.pct, {
    x: o.x, y: o.y, w: o.value - 0.06, h: 0.182,
    align: 'right', valign: 'middle', fontSize: 10.5, bold: true, color: WHITE,
  });
}

/* ------------------------------------------------------------------ *
 * Icons - drawn from primitives / font glyphs, centred on (cx, cy)
 * ------------------------------------------------------------------ */
const GLYPH_ICONS = {
  phone: '\u260E', envelope: '\u2709', database: '\u26C3', home: '\u2302',
  bolt: '\u26A1', gear: '\u2699', send: '\u27A4', rocket: '\u2708', sliders: '\u2630',
};

function icon(slide, name, cx, cy, s, color) {
  const solid = { fill: { color: color }, line: { type: 'none' } };
  const stroke = { fill: { type: 'none' }, line: { color: color, width: Math.max(1, s * 4) } };
  const ring = function (d, dx, dy) {
    shape(slide, S.donut, Object.assign({
      x: cx - d / 2 + (dx || 0), y: cy - d / 2 + (dy || 0), w: d, h: d,
    }, solid));
  };

  if (GLYPH_ICONS[name]) {
    text(slide, GLYPH_ICONS[name], {
      x: cx - s, y: cy - s, w: s * 2, h: s * 2, align: 'center', valign: 'middle',
      fontSize: s * 66, color: color, fontFace: 'DejaVu Sans',
    });
    return;
  }
  switch (name) {
    case 'clock':
      shape(slide, S.ellipse, Object.assign({ x: cx - s / 2, y: cy - s / 2, w: s, h: s }, stroke));
      shape(slide, S.rect, Object.assign({ x: cx - 0.012, y: cy - s * 0.32, w: 0.024, h: s * 0.34 }, solid));
      shape(slide, S.rect, Object.assign({ x: cx, y: cy - 0.012, w: s * 0.25, h: 0.024 }, solid));
      break;
    case 'lock':
      shape(slide, S.blockArc, Object.assign({
        x: cx - s * 0.3, y: cy - s * 0.56, w: s * 0.6, h: s * 0.6,
        angleRange: [180, 360], arcThicknessRatio: 0.3,
      }, solid));
      shape(slide, S.roundRect, Object.assign({
        x: cx - s * 0.44, y: cy - s * 0.22, w: s * 0.88, h: s * 0.64, rectRadius: s * 0.09,
      }, solid));
      break;
    case 'lifebuoy':
      shape(slide, S.ellipse, Object.assign({ x: cx - s / 2, y: cy - s / 2, w: s, h: s }, stroke));
      shape(slide, S.ellipse, Object.assign({ x: cx - s * 0.2, y: cy - s * 0.2, w: s * 0.4, h: s * 0.4 }, stroke));
      break;
    case 'wifi':
      [1, 0.68, 0.36].forEach(function (k) {
        shape(slide, S.blockArc, Object.assign({
          x: cx - s * k / 2, y: cy + s * 0.28 - s * k / 2, w: s * k, h: s * k,
          angleRange: [200, 340], arcThicknessRatio: 0.22,
        }, solid));
      });
      shape(slide, S.ellipse, Object.assign({ x: cx - s * 0.09, y: cy + s * 0.19, w: s * 0.18, h: s * 0.18 }, solid));
      break;
    case 'camera':
      shape(slide, S.rect, Object.assign({ x: cx - s * 0.18, y: cy - s * 0.46, w: s * 0.34, h: s * 0.14 }, solid));
      shape(slide, S.roundRect, Object.assign({
        x: cx - s * 0.55, y: cy - s * 0.36, w: s * 1.1, h: s * 0.8, rectRadius: s * 0.11,
      }, stroke));
      ring(s * 0.48);
      break;
    case 'chat':
      shape(slide, S.wedgeRoundRectCallout, Object.assign({
        x: cx - s * 0.55, y: cy - s * 0.5, w: s * 0.72, h: s * 0.58,
      }, solid));
      shape(slide, S.wedgeRoundRectCallout, Object.assign({
        x: cx - s * 0.07, y: cy - s * 0.12, w: s * 0.62, h: s * 0.52,
      }, solid));
      break;
    case 'bulb':
      shape(slide, S.ellipse, Object.assign({ x: cx - s * 0.38, y: cy - s * 0.54, w: s * 0.76, h: s * 0.76 }, stroke));
      shape(slide, S.rect, Object.assign({ x: cx - s * 0.18, y: cy + s * 0.22, w: s * 0.36, h: s * 0.11 }, solid));
      shape(slide, S.rect, Object.assign({ x: cx - s * 0.18, y: cy + s * 0.4, w: s * 0.36, h: s * 0.11 }, solid));
      break;
    case 'folder':
      poly(slide, [[0, 1], [0, 0.18], [0.44, 0.18], [0.56, 0], [1, 0], [1, 1]], {
        x: cx - s * 0.55, y: cy - s * 0.38, w: s * 1.1, h: s * 0.76,
        fill: { type: 'none' }, line: { color: color, width: Math.max(1, s * 4) },
      });
      break;
    case 'megaphone':
      shape(slide, S.trapezoid, Object.assign({
        x: cx - s * 0.4, y: cy - s * 0.32, w: s * 0.8, h: s * 0.64, rotate: 270,
      }, solid));
      shape(slide, S.rect, Object.assign({ x: cx - s * 0.52, y: cy + s * 0.08, w: s * 0.2, h: s * 0.4 }, solid));
      break;
    case 'map':
      [0, 1, 2].forEach(function (k) {
        const up = k % 2 === 0;
        poly(slide, [[0, up ? 0 : 0.2], [1, up ? 0.2 : 0], [1, up ? 0.8 : 1], [0, up ? 1 : 0.8]], {
          x: cx - s * 0.55 + k * s * 0.37, y: cy - s * 0.45, w: s * 0.37, h: s * 0.9,
          fill: { color: color }, line: { type: 'none' },
        });
      });
      break;
    case 'diamond':
      shape(slide, S.diamond, Object.assign({ x: cx - s * 0.5, y: cy - s * 0.42, w: s, h: s * 0.84 }, stroke));
      shape(slide, S.rect, Object.assign({ x: cx - s * 0.5, y: cy - s * 0.18, w: s, h: 0.014 }, solid));
      break;
    case 'bell':
      shape(slide, S.trapezoid, Object.assign({ x: cx - s * 0.4, y: cy - s * 0.42, w: s * 0.8, h: s * 0.62 }, stroke));
      shape(slide, S.ellipse, Object.assign({ x: cx - s * 0.11, y: cy + s * 0.24, w: s * 0.22, h: s * 0.22 }, solid));
      break;
    case 'inbox':
      poly(slide, [[0, 0], [1, 0], [1, 0.55], [0.68, 0.55], [0.6, 0.8], [0.4, 0.8], [0.32, 0.55], [0, 0.55]], {
        x: cx - s * 0.5, y: cy - s * 0.06, w: s, h: s * 0.56,
        fill: { color: color }, line: { type: 'none' },
      });
      shape(slide, S.downArrow, Object.assign({
        x: cx - s * 0.22, y: cy - s * 0.48, w: s * 0.44, h: s * 0.42,
      }, solid));
      break;
    default:
      break;
  }
}

/** Five-petal blossom used as an icon backdrop on slide 6. */
function blossom(slide, cx, cy, d, color) {
  const petal = d * 0.52;
  for (let k = 0; k < 5; k += 1) {
    const a = (Math.PI * 2 * k) / 5 - Math.PI / 2;
    shape(slide, S.ellipse, {
      x: cx + Math.cos(a) * d * 0.26 - petal / 2, y: cy + Math.sin(a) * d * 0.26 - petal / 2,
      w: petal, h: petal, fill: { color: color }, line: { type: 'none' },
    });
  }
  shape(slide, S.ellipse, {
    x: cx - d * 0.36, y: cy - d * 0.36, w: d * 0.72, h: d * 0.72,
    fill: { color: color }, line: { type: 'none' },
  });
}

/** Circular icon badge: outlined ring + icon (slides 15/16). */
function iconCircle(slide, o) {
  shape(slide, S.ellipse, {
    x: o.x, y: o.y, w: o.d, h: o.d, fill: { type: 'none' }, line: { color: o.ring, width: 3 },
  });
  icon(slide, o.icon, o.x + o.d / 2, o.y + o.d / 2, o.d * 0.42, o.color);
}

/**
 * Smartphone mock-up standing in for the device photo (slides 15/16):
 * metal rim, black glass, ear-piece slot and front camera.
 */
function phoneMockup(slide, o) {
  const cx = o.x + o.w / 2;
  shape(slide, S.roundRect, { // outer metal frame
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.5,
    fill: { color: 'A09FA4' }, line: { type: 'none' },
  });
  shape(slide, S.roundRect, { // dark bezel
    x: o.x + 0.055, y: o.y + 0.055, w: o.w - 0.11, h: o.h - 0.11, rectRadius: 0.46,
    fill: { color: '737278' }, line: { type: 'none' },
  });
  shape(slide, S.roundRect, { // glass
    x: o.x + 0.1, y: o.y + 0.1, w: o.w - 0.2, h: o.h - 0.2, rectRadius: 0.42,
    fill: { color: '0D0D0D' }, line: { type: 'none' },
  });
  poly(slide, [[0, 0], [0.587, 0], [0.488, 1], [0, 1]], { // glare above the screen
    x: o.x + 0.1, y: o.y + 0.1, w: o.w - 0.2, h: 1.44,
    fill: { color: '696969' }, line: { type: 'none' },
  });
  shape(slide, S.ellipse, { // front camera
    x: cx - 0.055, y: o.y + 0.51, w: 0.11, h: 0.11,
    fill: { color: '515151' }, line: { type: 'none' },
  });
  shape(slide, S.roundRect, { // ear-piece slot
    x: cx - o.w * 0.097, y: o.y + 0.88, w: o.w * 0.194, h: 0.15, rectRadius: 0.075,
    fill: { color: '1A1A1A' }, line: { color: '515151', width: 1 },
  });
  [1.6, 2.35, 2.95].forEach(function (dy) { // volume / power buttons
    shape(slide, S.roundRect, {
      x: o.x - 0.04, y: o.y + dy, w: 0.08, h: 0.32, rectRadius: 0.04,
      fill: { color: '9A99A0' }, line: { type: 'none' },
    });
  });
}

/* ------------------------------------------------------------------ *
 * Re-used photo-placeholder outlines (normalised polygons)
 * ------------------------------------------------------------------ */
const SHAPES = {
  star8: [ // slide 4: four-pointed "X" burst
    [0, 0.2388], [0.2397, 0], [0.4997, 0.2542], [0.7583, 0.0015], [0.9972, 0.2432],
    [0.7448, 0.4985], [1, 0.7568], [0.7569, 0.9971], [0.4997, 0.7356], [0.2426, 1],
    [0.0024, 0.7537], [0.2576, 0.4956],
  ],
  notchRight: [[0, 0], [1, 0], [1, 1], [0.5044, 1]], // slide 6
  triTopLeft: [[0, 0], [1, 0], [0, 1]], // slide 7
  wedgeRight: [[0, 0], [0.6952, 0], [1, 1], [0, 1]], // slide 8
  slant: [[0, 0], [0.75, 0], [1, 1], [0.25, 1]], // slide 9
  ribbon: [[0.25, 0], [1, 0], [1, 0.0004], [0.7501, 1], [0, 1]], // slide 10
  bandTop: [[0, 0], [0.8653, 0], [1, 1], [0, 1]], // slide 11 upper
  bandBottom: [[0, 0], [1, 0], [1, 1], [0.1343, 1]], // slide 11 lower
  pentaArrow: [[0, 0.326], [0.7795, 0], [1, 0.2345], [0.8303, 0.6658], [0, 1]], // slide 18
  pentaFold: [[0, 0], [0.8304, 0.3590], [0.2304, 1], [1, 0.3522]], // slide 18
  // Slide 5: sweeping paint stroke - one blob plus two flicked-off slivers.
  brushBody: [
    [0.624, 0.002], [0.369, 0.121], [0.22, 0.219], [0.166, 0.274], [0.247, 0.281], [0.242, 0.288],
    [0.066, 0.442], [0.17, 0.47], [0.159, 0.498], [0.211, 0.505], [0.121, 0.575], [0.03, 0.701],
    [0.001, 0.771], [0.001, 0.855], [0.015, 0.911], [0.037, 0.946], [0.082, 0.981], [0.123, 0.995],
    [0.175, 0.995], [0.251, 0.967], [0.613, 0.764], [0.428, 0.757], [0.47, 0.729], [0.455, 0.715],
    [0.626, 0.617], [0.624, 0.61], [0.434, 0.603], [0.488, 0.533], [0.461, 0.526], [0.457, 0.512],
    [0.516, 0.498], [0.741, 0.219], [0.764, 0.142], [0.759, 0.1], [0.714, 0.037], [0.651, 0.002],
  ],
  brushTail1: [[0.813, 0.494], [0.77, 0.508], [0.45, 0.731], [0.55, 0.6]],
  brushTail2: [[0.962, 0.564], [0.937, 0.578], [0.572, 0.759], [0.68, 0.648]],
};

/* ------------------------------------------------------------------ *
 * Repeated copy
 * ------------------------------------------------------------------ */
const LOREM_LONG = 'lacus nulla ac netus nibh aliquet, porttitor ligula justo libero vivamus '
  + 'porttitor dolor, conubia mollit. Sapien nam suspendisse, tincidunt eget ante tincidunt, '
  + 'eros in auctor fringilla praesent at diam. In et quam est eget mi.';
const LOREM_MAURIS = 'Mauris quam dolor, cursus at porta et, luctus eget purus. Nunc tempor '
  + 'luctus interdum. Duis libero leo, consequat ut accumsan eu, viverra et erat. ';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor '
  + 'ligula justo libero vivamus porttitor dolor, ';
const LOREM_ALIQUAM = 'Aliquam varius adipiscing tempor. Vivamus id ipsum sit amet massa '
  + 'consectetur porta. Class aptent taciti sociosqu ad litora torque.';
const LOREM_VENENATIS = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. '
  + 'In venenatis tortor et elit eleifend dictum. Praesent.';
const LOREM_VOKALIA = 'Lorem Ipsum Dolor Far far away, behind the word mountains, '
  + 'far from the countries Vokalia and';
const CREATIVE_SPLIT = [
  { text: 'Creative', options: { color: RED } },
  { text: ' Powerpoint', options: { color: DARK } },
];

/* ================================================================== *
 * Slides
 * ================================================================== */

function slide01(pres) { // Title
  const s = pres.addSlide();
  s.background = { color: TITLEBG };
  text(s, 'MAXX', {
    x: 5.219, y: 2.837, w: 2.895, h: 1.212, align: 'center', wrap: false,
    fontSize: 66, bold: true, charSpacing: -1.5, color: TRACK,
  });
  text(s, 'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula '
    + 'justo libero vivamus porttitor dolor, conubia mollit. Sapien nam suspendisse, tincidunt '
    + 'eget ante tincidunt, eros in auctor fringilla praesent at diam. In et quam est eget mi. '
    + 'Pellentesque nunc orci eu enim, eget in fringilla vitae, et ero, ', {
    x: 4.304, y: 4.355, w: 4.726, h: 1.01, align: 'center',
    fontSize: 9, color: WHITE, lineSpacingMultiple: 1.5,
  });
  crossMark(s, 6.275, 1.827, 0.783, RED);
  sideTab(s, 10.906, 3.378, 90, [{ text: 'Creative Powerpoint', options: { color: WHITE } }]);
  sideTab(s, -0.698, 3.378, 270, [{ text: 'Creative Powerpoint', options: { color: WHITE } }]);
}

function slide02(pres) { // Gallery - diamonds on a red band
  const s = pres.addSlide();
  shape(s, S.rect, { x: 0, y: 0, w: 13.333, h: 2.636, fill: { color: RED }, line: { type: 'none' } });
  photo(s, { x: 2.949, y: 1.199, w: 2.907, h: 2.872, shape: S.diamond });
  photo(s, { x: 7.223, y: 1.199, w: 2.907, h: 2.872, shape: S.diamond });
  text(s, LOREM_MAURIS, {
    x: 6.984, y: 5.206, w: 3.56, h: 0.864, fontSize: 11,
    lineSpacingMultiple: 1.5, margin: [4.8, 4.8, 2.4, 2.4],
  });
  headline(s, { x: 3.236, y: 4.865, w: 2.976, h: 1.582, word: 'GALLERY', align: 'right', fontSize: 44 });
  crossMark(s, 11.278, 5.309, 0.783, RED);
  photo(s, { x: 4.426, y: 0.492, w: 4.338, h: 4.285, shape: S.diamond });
}

function slide03(pres) { // Fashion - dark side bar
  const s = pres.addSlide();
  shape(s, S.rect, { x: 0, y: 0, w: 3.597, h: 7.5, fill: { color: DARK }, line: { type: 'none' } });
  photo(s, { x: 1.173, y: 1.545, w: 4.41, h: 4.409 });
  headline(s, { x: 6.86, y: 1.83, w: 4.001, h: 1.717, word: 'FASHION' });
  text(s, LOREM_LONG, {
    x: 6.86, y: 3.848, w: 3.783, h: 1.01, align: 'justify', fontSize: 9, lineSpacingMultiple: 1.5,
  });
  crossMark(s, 6.706, 5.813, 0.783, RED);
  sideTab(s, 10.906, 3.378, 90, CREATIVE_SPLIT);
}

function slide04(pres) { // Creative - red slide, burst photo
  const s = pres.addSlide();
  s.background = { color: RED };
  sideTab(s, -0.828, 3.295, 270, [{ text: 'Creative Powerpoint', options: { color: WHITE } }]);
  headline(s, {
    x: 2.499, y: 2.033, w: 4.001, h: 1.717, word: 'CREATIVE', color1: WHITE, color2: WHITE,
  });
  text(s, LOREM_LONG, {
    x: 2.499, y: 4.051, w: 3.783, h: 1.01, align: 'justify',
    fontSize: 9, color: WHITE, lineSpacingMultiple: 1.5,
  });
  photo(s, { x: 7.041, y: 1.294, w: 4.746, h: 4.753, pts: SHAPES.star8 });
}

function slide05(pres) { // Personal creative - brush stroke + skills
  const s = pres.addSlide();
  const brush = { x: 0.575, y: 0.171, w: 7.377, h: 7.149 };
  [SHAPES.brushTail1, SHAPES.brushTail2].forEach(function (tail) {
    poly(s, tail, Object.assign({ fill: { color: PH_FILL }, line: { type: 'none' } }, brush));
  });
  photo(s, Object.assign({ pts: SHAPES.brushBody }, brush));
  text(s, [
    { text: 'MAXX', options: { color: RED, breakLine: true } },
    { text: 'PERSONAL', options: { color: DARK, breakLine: true } },
    { text: 'CREATIVE', options: { color: DARK } },
  ], { x: 7.307, y: 1.001, w: 4.001, h: 2.524, fontFace: LATO_BLACK, fontSize: 48, bold: true });
  progressBar(s, { x: 8.276, y: 4.495, value: 2.779, pct: '80%', label: 'SKILL NAME', labelW: 1.91 });
  progressBar(s, { x: 8.276, y: 5.222, value: 2.063, pct: '60%', label: 'SKILL NAME', labelW: 1.966 });
}

function slide06(pres) { // Gallery - feature list
  const s = pres.addSlide();
  photo(s, { x: 4.603, y: 0, w: 8.73, h: 7.5, pts: SHAPES.notchRight });
  s.addShape(S.line, {
    x: 6.206, y: 3.317, w: 2.476, h: 4.183, line: { color: BLACK, width: 2.25 },
  });
  [
    { y: 4.406, glyph: 'diamond', ty: 4.317, by: 4.683, bh: 0.667 },
    { y: 5.655, glyph: 'bell', ty: 5.675, by: 6.041, bh: 0.62 },
  ].forEach(function (row) {
    blossom(s, 1.2745, row.y + 0.4455, 0.86, BLACK);
    icon(s, row.glyph, 1.2745, row.y + 0.4455, 0.4, WHITE);
    text(s, 'Inceptos Hemanoes', { x: 1.897, y: row.ty, w: 2.765, h: 0.37, fontSize: 16 });
    text(s, LOREM_VENENATIS, {
      x: 1.897, y: row.by, w: 4.588, h: row.bh, fontSize: 11,
      color: NEARBLACK, lineSpacingMultiple: 1.5,
    });
  });
  text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. In venenatis tortor et elit '
    + 'eleifend dictum. Praesent mollis velit et nisi placerat feugiat. In id tortor.', {
    x: 0.873, y: 2.813, w: 4.715, h: 0.897, fontSize: 11, color: NEARBLACK, lineSpacingMultiple: 1.5,
  });
  headline(s, { x: 0.873, y: 0.738, w: 4.001, h: 1.717, word: 'GALLERY' });
}

function slide07(pres) { // Profile - diagonal photo
  const s = pres.addSlide();
  photo(s, { x: 0, y: 0, w: 7.5, h: 7.5, pts: SHAPES.triTopLeft });
  text(s, LOREM_LONG, {
    x: 6.008, y: 4.605, w: 3.783, h: 1.01, align: 'justify', fontSize: 9, lineSpacingMultiple: 1.5,
  });
  headline(s, { x: 5.899, y: 2.415, w: 4.001, h: 1.717, word: 'PROFILE' });
  crossMark(s, 3.481, 5.223, 0.783, RED);
  sideTab(s, 9.656, 3.809, 90, CREATIVE_SPLIT);
}

function slide08(pres) { // Gallery - two icon bullets
  const s = pres.addSlide();
  photo(s, { x: 0, y: 0, w: 6.125, h: 7.5, pts: SHAPES.wedgeRight });
  [
    { oy: 3.413, ly: 3.37, sy: 3.757, glyph: 'phone' },
    { oy: 5.251, ly: 5.208, sy: 5.629, glyph: 'inbox' },
  ].forEach(function (row) {
    shape(s, S.ellipse, {
      x: 6.591, y: row.oy, w: 0.614, h: 0.614, fill: { color: BLACK }, line: { type: 'none' },
    });
    icon(s, row.glyph, 6.898, row.oy + 0.307, 0.3, WHITE);
    text(s, 'Also called investment', {
      x: 7.613, y: row.ly, w: 2.588, h: 0.286, valign: 'middle', wrap: false,
      fontSize: 11, bold: true, italic: true, charSpacing: 3,
    });
    text(s, LOREM_VOKALIA, {
      x: 7.534, y: row.sy, w: 3.938, h: 0.645, fontSize: 11,
      lineSpacingMultiple: 1.5, margin: [8.6, 8.6, 4.3, 4.3],
    });
  });
  headline(s, { x: 6.591, y: 0.954, w: 4.001, h: 1.717, word: 'GALLERY' });
  sideTab(s, 10.906, 3.378, 90, CREATIVE_SPLIT);
}

function slide09(pres) { // Gallery - slanted photos + skills
  const s = pres.addSlide();
  progressBar(s, { x: 1.293, y: 4.738, value: 2.779, pct: '80%', label: 'EXPOSURE', labelW: 1.327 });
  progressBar(s, { x: 1.293, y: 5.465, value: 2.063, pct: '60%', label: 'COMPOSURE', labelW: 1.55 });
  text(s, LOREM_MAURIS, {
    x: 1.231, y: 2.949, w: 3.56, h: 0.864, fontSize: 11,
    lineSpacingMultiple: 1.5, margin: [4.8, 4.8, 2.4, 2.4],
  });
  headline(s, { x: 1.194, y: 0.926, w: 2.976, h: 1.582, word: 'GALLERY', fontSize: 44 });
  crossMark(s, 11.748, 0.928, 0.783, RED);
  photo(s, { x: 5.302, y: 1.081, w: 3.705, h: 5.337, pts: SHAPES.slant });
  photo(s, { x: 8.176, y: 1.081, w: 3.705, h: 5.337, pts: SHAPES.slant });
}

function slide10(pres) { // Gallery - red slide with white slashes
  const s = pres.addSlide();
  s.background = { color: RED };
  shape(s, S.rect, {
    x: 4.722, y: -0.246, w: 0.102, h: 3.459, rotate: 12.9,
    fill: { color: OFFWHITE }, line: { type: 'none' },
  });
  shape(s, S.rect, {
    x: 9.197, y: 4.172, w: 0.102, h: 3.459, rotate: 13.2,
    fill: { color: OFFWHITE }, line: { type: 'none' },
  });
  text(s, 'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo '
    + 'libero vivamus porttitor dolor, conubia mollit. Sapien nam suspendisse', {
    x: 0.745, y: 3.237, w: 2.814, h: 1.132, align: 'right',
    fontFace: ROBOTO, fontSize: 10.5, color: WHITE, lineSpacingMultiple: 1.5,
  });
  text(s, 'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo '
    + 'libero vivamus porttitor dolor, conubia mollit. Sapien nam suspendisse, tincidunt eget '
    + 'ante tincidunt, eros in auctor, ', {
    x: 9.967, y: 5.333, w: 2.869, h: 1.662,
    fontFace: ROBOTO, fontSize: 10.5, color: WHITE, lineSpacingMultiple: 1.5,
  });
  headline(s, {
    x: 0.583, y: 1.163, w: 2.976, h: 1.582, word: 'GALLERY',
    align: 'right', fontSize: 44, color1: WHITE, color2: WHITE,
  });
  crossMark(s, 10.125, 3.358, 0.783, OFFWHITE);
  sideTab(s, 10.857, 2.879, 90, [{ text: 'Creative Powerpoint', options: { color: WHITE } }]);
  photo(s, { x: 3.619, y: 0, w: 6.794, h: 7.5, pts: SHAPES.ribbon });
}

function slide11(pres) { // Two full-width photo bands
  const s = pres.addSlide();
  [
    { align: 'right', tx: 9.898, ty: 1.854, hx: 10.885, hy: 1.285 },
    { align: 'left', tx: 0.462, ty: 4.912, hx: 0.462, hy: 4.374 },
  ].forEach(function (b) {
    text(s, LOREM_ALIQUAM, {
      x: b.tx, y: b.ty, w: 2.955, h: 0.828, align: b.align,
      fontSize: 10.5, lineSpacingMultiple: 1.5, margin: [4.8, 4.8, 2.4, 2.4],
    });
    text(s, [
      { text: 'MAXX', options: { color: RED } },
      { text: ' GALLERY', options: { color: BLACK } },
    ], {
      x: b.hx, y: b.hy, w: 1.968, h: 0.37, align: b.align, wrap: false,
      fontFace: MONT, fontSize: 18, bold: true, charSpacing: -1.5,
      margin: [4.8, 4.8, 2.4, 2.4],
    });
  });
  photo(s, { x: 0, y: 1.086, w: 10.289, h: 2.624, pts: SHAPES.bandTop });
  photo(s, { x: 3.013, y: 3.79, w: 10.321, h: 2.624, pts: SHAPES.bandBottom });
}

function slide12(pres) { // Pictures - 2 x 2 checkerboard
  const s = pres.addSlide();
  text(s, 'AWESOME TITLE', {
    x: 7.02, y: 1.16, w: 4.033, h: 0.715, fontFace: LATO_BLACK,
    fontSize: 28, charSpacing: 3, color: DARK, lineSpacingMultiple: 1.5,
  });
  text(s, LOREM_SHORT, {
    x: 7.02, y: 2.039, w: 3.367, h: 0.896, align: 'justify', fontSize: 10.5, lineSpacingMultiple: 1.5,
  });
  text(s, [
    { text: 'MAXX', options: { color: RED } },
    { text: ' PICTURES', options: { color: BLACK } },
  ], { x: 1.4, y: 4.241, w: 4.96, h: 0.89, fontFace: LATO_BLACK, fontSize: 36, lineSpacingMultiple: 1.5 });
  text(s, LOREM_SHORT, {
    x: 1.406, y: 5.331, w: 4.7, h: 0.631, align: 'justify', fontSize: 10.5, lineSpacingMultiple: 1.5,
  });
  sideTab(s, 11.077, 2.318, 90, CREATIVE_SPLIT);
  photo(s, { x: 1.4, y: 1.16, w: 4.96, h: 2.655 });
  photo(s, { x: 7.02, y: 3.814, w: 4.96, h: 2.655 });
}

function slide13(pres) { // Pictures - three across
  const s = pres.addSlide();
  text(s, 'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo '
    + 'libero vivamus porttitor dolor, conubia mollit. Sapien nam suspendisse, tincidunt eget ante '
    + 'tincidunt, eros in auctor fringilla praesent at diam. In et quam est eget mi. Pellentesque '
    + 'nunc orci eu enim, eget in fringilla vitae, et eros praesent dolor porttitor. Lacinia lectus '
    + 'nonummy, accumsan mauris in sed justotincidunt, eros in auctor fringilla praesent at diam. ', {
    x: 2.11, y: 5.533, w: 9.127, h: 0.783, align: 'center',
    fontSize: 9, color: INK, lineSpacingMultiple: 1.5,
  });
  text(s, [
    { text: 'MAXX', options: { color: RED } },
    { text: ' PICTURES', options: { color: BLACK } },
  ], {
    x: 4.194, y: 0.85, w: 4.96, h: 1.065, align: 'center', fontFace: LATO_BLACK,
    fontSize: 44, bold: true, charSpacing: -1.5, lineSpacingMultiple: 1.5,
  });
  photo(s, { x: 0, y: 2.486, w: 4.333, h: 2.444 });
  photo(s, { x: 4.5, y: 2.486, w: 4.347, h: 2.444 });
  photo(s, { x: 9, y: 2.486, w: 4.333, h: 2.444 });
}

function slide14(pres) { // Gallery photos - red mosaic
  const s = pres.addSlide();
  s.background = { color: RED };
  text(s, 'MAXX', {
    x: 0.841, y: 6.361, w: 1.632, h: 0.707, wrap: false,
    fontSize: 36, bold: true, charSpacing: -1.5, color: WHITE,
  });
  crossMark(s, 12.388, 6.523, 0.783, OFFWHITE);
  sideTab(s, 11.216, 4.477, 90, [{ text: 'Creative Powerpoint', options: { color: WHITE } }]);
  text(s, 'GALLERY PHOTOS', {
    x: 7.853, y: 0.24, w: 4.372, h: 0.774, align: 'right', wrap: false,
    fontSize: 40, bold: true, italic: true, charSpacing: -1.5, color: WHITE,
  });
  photo(s, { x: 0.957, y: 3.361, w: 3.292, h: 2.847 });
  photo(s, { x: 0.957, y: 1.125, w: 11.419, h: 2.014 });
  photo(s, { x: 4.513, y: 3.361, w: 7.863, h: 2.847 });
}

function slide15(pres) { // Mockup - red, phone centred, four features
  const s = pres.addSlide();
  s.background = { color: RED };
  phoneMockup(s, { x: 4.107, y: 2.007, w: 5.16, h: 6.0 });
  photo(s, { x: 4.46, y: 3.524, w: 4.413, h: 3.976 });
  [
    { cx: 1.674, cy: 2.306, ix: 1.997, iy: 2.672, ty: 3.6, by: 3.964, tx: 1.092, bx: 1.106, icon: 'camera' },
    { cx: 1.688, cy: 5.004, ix: 2.046, iy: 5.371, ty: 6.298, by: 6.662, tx: 1.106, bx: 1.121, icon: 'chat' },
    { cx: 10.853, cy: 2.306, ix: 11.147, iy: 2.676, ty: 3.6, by: 3.964, tx: 10.271, bx: 10.286, icon: 'map' },
    { cx: 10.867, cy: 5.004, ix: 11.236, iy: 5.333, ty: 6.298, by: 6.662, tx: 10.286, bx: 10.3, icon: 'database' },
  ].forEach(function (f) {
    iconCircle(s, { x: f.cx, y: f.cy, d: 1.097, ring: GREY1, color: OFFWHITE, icon: f.icon });
    text(s, 'Lorem Ipsum', {
      x: f.tx, y: f.ty, w: 2.372, h: 0.364, align: 'center',
      fontSize: 12, charSpacing: 3, color: WHITE, lineSpacingMultiple: 1.5,
    });
    text(s, 'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh,', {
      x: f.bx, y: f.by, w: 2.343, h: 0.596, align: 'center',
      fontSize: 10.5, color: WHITE, lineSpacingMultiple: 1.5,
    });
  });
  text(s, 'MAXX MOCKUP', {
    x: 4.066, y: 0.771, w: 5.202, h: 0.774, align: 'center',
    fontSize: 40, bold: true, color: WHITE,
  });
  crossMark(s, 3.641, 0.767, 0.783, OFFWHITE);
}

function slide16(pres) { // Mockup - white, phone left, six icons
  const s = pres.addSlide();
  phoneMockup(s, { x: 1.02, y: 1.046, w: 5.16, h: 6.96 });
  photo(s, { x: 1.38, y: 2.587, w: 4.413, h: 4.913 });
  text(s, 'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo '
    + 'libero vivamus porttitor dolor, conubia mollit. Sapien nam suspendisse, tincidunt eget ante '
    + 'tincidunt, eros in auctor fringilla praesent at diam. In et quam est eget mi. Pellentesque '
    + 'nunc orci eu enim, eget in fringilla vitae, et eros praesent dolor porttitor. Lacinia '
    + 'lectus nonummy, ', {
    x: 6.804, y: 2.02, w: 5.375, h: 1.363, align: 'justify', fontSize: 10, lineSpacingMultiple: 1.5,
  });
  const cols = [7.155, 8.959, 10.763];
  [
    { y: 4.049, icons: ['bulb', 'folder', 'envelope'] },
    { y: 5.673, icons: ['wifi', 'megaphone', 'phone'] },
  ].forEach(function (row) {
    row.icons.forEach(function (name, i) {
      iconCircle(s, { x: cols[i], y: row.y, d: 1.052, ring: BLACK, color: BLACK, icon: name });
    });
  });
  text(s, [
    { text: 'MAXX', options: { color: RED } },
    { text: ' MOCKUP', options: { color: DARK } },
  ], { x: 6.358, y: 0.894, w: 5.202, h: 0.774, align: 'center', fontSize: 40, bold: true });
  sideTab(s, -1.032, 3.982, 270, CREATIVE_SPLIT);
}

function slide17(pres) { // SWOT
  const s = pres.addSlide();
  text(s, 'INFOGRAPHIC', {
    x: 4.525, y: 0.459, w: 4.427, h: 0.774, align: 'center',
    fontSize: 40, bold: true, italic: true, color: RED,
  });
  const cards = [
    { x: 4.135, y: 2.048, fill: RED, letter: 'S', size: 66, ly: 2.431, lh: 1.212 },
    { x: 6.651, y: 2.048, fill: BLACK, letter: 'W', size: 60, ly: 2.494, lh: 1.111 },
    { x: 4.135, y: 4.254, fill: BLACK, letter: 'O', size: 60, ly: 4.672, lh: 1.111 },
    { x: 6.651, y: 4.254, fill: RED, letter: 'T', size: 60, ly: 4.672, lh: 1.111 },
  ];
  cards.forEach(function (c) {
    shape(s, S.roundRect, {
      x: c.x, y: c.y, w: 2.333, h: 2.048, rectRadius: 0.341,
      fill: { color: c.fill }, line: { type: 'none' },
    });
    text(s, c.letter, {
      x: c.x + 0.186, y: c.ly, w: 1.96, h: c.lh, align: 'center',
      fontSize: c.size, charSpacing: 3, color: WHITE,
    });
  });
  const CALLOUT = 'Lorem ipsum dolor sit dudu amet, consectetur adipiscing elit. ';
  [
    {
      label: 'STRENGHT', align: 'right', lx: 0.844, ly: 2.133, tx: 1.443, ty: 2.431, tw: 1.523,
      dot: [2.999, 2.234], bar: [3.131, 2.316], leg: [4.55, 2.313, 0.409, 0.386, false, false],
    },
    {
      label: 'OPPORTUNITY', align: 'right', lx: 1.012, ly: 5.847, tx: 1.586, ty: 6.145, tw: 1.547,
      dot: [3.167, 5.986], bar: [3.27, 6.066], leg: [4.688, 5.702, 0.307, 0.364, true, false],
    },
    {
      label: 'WEAKNESS', align: 'justify', lx: 10.309, ly: 2.133, tx: 10.309, ty: 2.4, tw: 1.706,
      dot: [10.137, 2.261], bar: [8.798, 2.354], leg: [8.383, 2.354, 0.415, 0.447, false, true],
    },
    {
      label: 'TREAD', align: 'justify', lx: 10.075, ly: 5.753, tx: 10.075, ty: 6.021, tw: 1.706,
      dot: [9.857, 5.876], bar: [8.456, 5.954], leg: [8.083, 5.576, 0.373, 0.39, false, false],
    },
  ].forEach(function (c) {
    s.addShape(S.line, {
      x: c.leg[0], y: c.leg[1], w: c.leg[2], h: c.leg[3],
      flipH: c.leg[4], flipV: c.leg[5], line: { color: BLACK, width: 1.5 },
    });
    s.addShape(S.line, {
      x: c.bar[0], y: c.bar[1], w: 1.419, h: 0, line: { color: BLACK, width: 1.5 },
    });
    shape(s, S.ellipse, {
      x: c.dot[0], y: c.dot[1], w: 0.159, h: 0.159, fill: { color: BLACK }, line: { color: BLACK },
    });
    text(s, c.label, {
      x: c.lx, y: c.ly, w: 2.121, h: 0.379, align: c.align,
      fontSize: 11, charSpacing: 3, lineSpacingMultiple: 1.5,
    });
    text(s, CALLOUT, { // PowerPoint auto-shrinks these boxes to 92.5% / 0.88 line spacing
      x: c.tx - 0.11, y: c.ty, w: c.tw + 0.22, h: 0.765, align: c.align,
      fontSize: 9.5, lineSpacingMultiple: 0.88, margin: [7.2, 7.2, 3.6, 3.6],
    });
  });
}

function slide18(pres) { // Stacked arrow steps
  const s = pres.addSlide();
  const steps = [
    {
      bar: [3.798, 2.496, 6.693, 1.027], barFill: DARK,
      penta: [3.798, 2.001, 1.423, 1.518], pentaFill: GREY4,
      fold: [4.906, 2.001, 0.314, 1.011], foldFill: 'EBEBEB',
      icon: 'home', num: '01', numX: 5.219, numY: 2.697, lineX: 6.087, lineY: 2.755, txX: 6.165, txY: 2.65,
    },
    {
      bar: [2.843, 3.523, 7.648, 1.028], barFill: GREY2,
      penta: [2.843, 3.029, 1.421, 1.515], pentaFill: '868686',
      fold: [3.961, 3.04, 0.303, 0.998], foldFill: 'D1D1D1',
      icon: 'clock', num: '02', numX: 5.172, numY: 3.725, lineX: 6.04, lineY: 3.783, txX: 6.118, txY: 3.678,
    },
    {
      bar: [4.066, 4.551, 6.424, 1.027], barFill: GREY3,
      penta: [4.066, 4.059, 1.421, 1.517], pentaFill: '717171',
      fold: [5.178, 4.063, 0.31, 1.007], foldFill: 'C0C0C0',
      icon: 'rocket', num: '03', numX: 5.706, numY: 4.752, lineX: 6.574, lineY: 4.81, txX: 6.652, txY: 4.705,
    },
    {
      bar: [4.641, 5.578, 5.85, 1.028], barFill: GREY4,
      penta: [4.641, 5.094, 1.42, 1.515], pentaFill: '606060',
      fold: [5.747, 5.094, 0.314, 1.011], foldFill: 'B3B3B3',
      icon: 'bolt', num: '04', numX: 6.021, numY: 5.78, lineX: 6.889, lineY: 5.837, txX: 6.967, txY: 5.733,
    },
  ];
  steps.forEach(function (st) {
    shape(s, S.rect, {
      x: st.bar[0], y: st.bar[1], w: st.bar[2], h: st.bar[3],
      fill: { color: st.barFill }, line: { type: 'none' },
    });
  });
  steps.forEach(function (st) {
    poly(s, SHAPES.pentaFold, {
      x: st.fold[0], y: st.fold[1], w: st.fold[2], h: st.fold[3],
      fill: { color: st.foldFill }, line: { type: 'none' },
    });
    poly(s, SHAPES.pentaArrow, {
      x: st.penta[0], y: st.penta[1], w: st.penta[2], h: st.penta[3],
      fill: { color: st.pentaFill }, line: { type: 'none' },
    });
    icon(s, st.icon, st.penta[0] + st.penta[2] * 0.44, st.penta[1] + st.penta[3] * 0.44, 0.47, WHITE);
    text(s, st.num, {
      x: st.numX, y: st.numY, w: 0.806, h: 0.505, align: 'right', fontSize: 24, color: WHITE,
    });
    s.addShape(S.line, {
      x: st.lineX, y: st.lineY, w: 0, h: 0.457, line: { color: WHITE, width: 2 },
    });
    text(s, [
      { text: 'Lorem ipsum dolor sit amet, consectetur ', options: { breakLine: true } },
      { text: 'adipiscing ed id risus quis' },
    ], {
      x: st.txX, y: st.txY, w: 4.246, h: 0.619,
      fontSize: 11, color: WHITE, lineSpacingMultiple: 1.4,
    });
  });
  text(s, 'INFOGRAPHIC', {
    x: 4.525, y: 0.459, w: 4.427, h: 0.774, align: 'center',
    fontSize: 40, bold: true, italic: true, color: RED,
  });
}

function slide19(pres) { // Hexagon ring
  const s = pres.addSlide();
  const HEX_W = 1.333;
  const HEX_H = 1.539;
  const hexes = [
    { cx: 5.9565, cy: 3.2655, fill: GREY6, dark: '3A3A3A', icon: 'send', num: '6', nx: 5.972, ny: 3.458 },
    { cx: 7.3635, cy: 3.2655, fill: RED, dark: '900000', icon: 'clock', num: '1', nx: 6.899, ny: 3.458 },
    { cx: 8.0735, cy: 4.4905, fill: RED, dark: '900000', icon: 'sliders', num: '2', nx: 7.363, ny: 4.26 },
    { cx: 7.3635, cy: 5.7165, fill: GREY3, dark: '717171', icon: 'gear', num: '3', nx: 6.899, ny: 5.063 },
    { cx: 5.9625, cy: 5.7165, fill: GREY4, dark: '606060', icon: 'lock', num: '4', nx: 5.972, ny: 5.063 },
    { cx: 5.2595, cy: 4.4905, fill: GREY5, dark: '474747', icon: 'lifebuoy', num: '5', nx: 5.509, ny: 4.26 },
  ];
  hexes.forEach(function (h) {
    shape(s, S.hexagon, {
      x: h.cx - HEX_H / 2, y: h.cy - HEX_W / 2, w: HEX_H, h: HEX_W, rotate: 90,
      fill: { color: h.fill }, line: { type: 'none' },
    });
    icon(s, h.icon, h.cx, h.cy - 0.12, 0.42, WHITE);
  });
  hexes.forEach(function (h) {
    shape(s, S.ellipse, { // darker halo left by the hexagons overlapping
      x: h.nx - 0.05, y: h.ny - 0.05, w: 0.562, h: 0.562,
      fill: { color: h.dark }, line: { type: 'none' },
    });
    shape(s, S.ellipse, {
      x: h.nx, y: h.ny, w: 0.462, h: 0.462, fill: { color: WHITE }, line: { type: 'none' },
    });
    text(s, h.num, {
      x: h.nx - 0.17, y: h.ny + 0.029, w: 0.802, h: 0.404,
      align: 'center', fontSize: 18, charSpacing: 0.3,
    });
  });
  [
    { x: 8.987, y: 2.497, align: 'left', tx: 8.987 },
    { x: 9.686, y: 3.95, align: 'left', tx: 9.686 },
    { x: 8.996, y: 5.435, align: 'left', tx: 8.996 },
    { x: 2.279, y: 2.497, align: 'right', tx: 2.732 },
    { x: 1.722, y: 3.95, align: 'right', tx: 2.175 },
    { x: 2.279, y: 5.435, align: 'right', tx: 2.732 },
  ].forEach(function (c) { calloutText(s, c, 0.335); });
  text(s, 'INFOGRAPHIC', {
    x: 4.525, y: 0.459, w: 4.427, h: 0.774, align: 'center',
    fontSize: 40, bold: true, italic: true, color: RED,
  });
}

/** "Write your title" + two grey lines, used on slides 19 and 20. */
function calloutText(slide, c, gap) {
  text(slide, 'Write your title', {
    x: c.tx, y: c.y, w: 1.598, h: 0.337, align: c.align, wrap: false,
    fontSize: 14, charSpacing: 0.3,
  });
  text(slide, [
    { text: 'Lorem ipsum dolor, consec ', options: { breakLine: true } },
    { text: c.align === 'right' ? 'adipiscing elit roin ' : 'adipiscing elit roin. ' },
  ], {
    x: c.x, y: c.y + gap, w: 2.051, h: 0.619, align: c.align, wrap: false,
    fontSize: 11, color: '848484', lineSpacingMultiple: 1.4,
  });
}

function slide20(pres) { // Light-bulb arrow funnel
  const s = pres.addSlide();
  const parts = [
    { pts: [[1, 0.5], [0.75, 1], [0.75, 0.75], [0.2744, 0.75], [0, 0.75], [0, 0.2345], [0.3503, 0.2345], [0.7493, 0.2345], [0.7487, 0]], box: [5.754, 2.764, 2.612, 0.601], fill: RED },
    { pts: [[0.003, 0.2345], [0.4025, 0.2345], [0.7493, 0.2345], [0.7487, 0], [1, 0.5026], [0.75, 1], [0.75, 0.75], [0, 0.75], [0, 0.2345]], box: [5.754, 3.831, 2.612, 0.601], fill: GREY3 },
    { pts: [[1, 0.0052], [0.3937, 0.4052], [0.3937, 1], [0, 1], [0, 0.1195], [0.1772, 0]], box: [6.846, 4.814, 0.786, 0.596], fill: GREY5 },
    { pts: [[0.3307, 0.2474], [0.6989, 0.2423], [1, 0.2397], [1, 0.7552], [0.9727, 0.7552], [0.5833, 0.7603], [0.2513, 0.7629], [0.2519, 1], [0, 0.5], [0.2496, 0], [0.2496, 0.2474]], box: [4.9, 3.307, 2.612, 0.601], fill: GREY2 },
    { pts: [[1, 0], [1, 0.0294], [0.6218, 0.9853], [0.6218, 1], [0, 1], [0.3939, 0.0147]], box: [5.762, 3.761, 1.678, 0.211], fill: '717171' },
    { pts: [[1, 0.0064], [0.5804, 0.9872], [0, 1], [0.4271, 0]], box: [5.764, 3.214, 1.657, 0.242], fill: 'A6A6A6' },
    { pts: [[0.9237, 0], [1, 0.6246], [0.6538, 1], [0, 1], [0, 0.9967]], box: [5.754, 2.439, 1.4, 0.466], fill: RED },
    { pts: [[1, 0.1195], [0.9961, 1], [0.6039, 1], [0.6039, 0.4052], [0, 0.0052], [0.8196, 0]], box: [5.748, 4.814, 0.79, 0.596], fill: GREY5 },
    { pts: [[0, 1], [0.4282, 0], [1, 0.0064], [0.5821, 0.9873]], box: [5.923, 4.281, 1.66, 0.243], fill: '606060' },
    { pts: [[0, 0.504], [0.25, 0], [0.25, 0.2506], [1, 0.2429], [1, 0.7571], [0.2506, 0.7674], [0.2519, 1]], box: [5.143, 4.374, 2.49, 0.599], fill: GREY4 },
  ];
  parts.forEach(function (p) {
    poly(s, p.pts, {
      x: p.box[0], y: p.box[1], w: p.box[2], h: p.box[3],
      fill: { color: p.fill }, line: { type: 'none' },
    });
  });
  shape(s, S.roundRect, { // bulb neck
    x: 6.112, y: 5.445, w: 1.127, h: 0.291, rectRadius: 0.145,
    fill: { color: RED }, line: { type: 'none' },
  });
  shape(s, S.roundRect, { // bulb base
    x: 6.217, y: 5.832, w: 0.904, h: 0.56, rectRadius: 0.2,
    fill: { color: RED }, line: { type: 'none' },
  });
  ['Step 01', 'Step 02', 'Step 03', 'Step 04'].forEach(function (label, i) {
    text(s, label, {
      x: 5.92, y: [2.866, 3.409, 3.933, 4.475][i], w: 1.493, h: 0.337,
      align: 'center', fontSize: 14, color: WHITE,
    });
  });
  [
    { x: 8.697, y: 2.573, align: 'left', tx: 8.697 },
    { x: 8.697, y: 4.315, align: 'left', tx: 8.697 },
    { x: 2.517, y: 3.079, align: 'right', tx: 2.971 },
    { x: 2.517, y: 4.687, align: 'right', tx: 2.971 },
  ].forEach(function (c) { calloutText(s, c, 0.378); });
  text(s, 'INFOGRAPHIC', {
    x: 4.525, y: 0.459, w: 4.427, h: 0.774, align: 'center',
    fontSize: 40, bold: true, italic: true, color: RED,
  });
}

/* ================================================================== */

function build() {
  const pres = new PptxGenJS();
  S = pres.ShapeType;
  pres.defineLayout({ name: 'MAXX_16x9', width: 13.333333, height: 7.5 });
  pres.layout = 'MAXX_16x9';
  pres.title = 'MAXX Creative Powerpoint';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach(function (fn) { fn(pres); });

  return pres.writeFile({
    fileName: path.join(__dirname, '08dbeba6-20f2-4b01-b1f9-3798176b4acf_grok_final.pptx'),
  });
}

build().then(function (f) { console.log('wrote', f); }).catch(function (e) {
  console.error(e);
  process.exit(1);
});
