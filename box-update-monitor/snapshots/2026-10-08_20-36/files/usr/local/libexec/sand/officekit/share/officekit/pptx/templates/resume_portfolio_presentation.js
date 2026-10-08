/**
 * "RESUME" portfolio template - rebuilt with pptxgenjs.
 * 22 slides, 13.333in x 7.5in (16:9), fonts Public Sans (display) / Work Sans (text).
 *
 * The original deck's picture placeholders are empty, so they stay empty here.
 * The raster artwork that IS embedded (corner logo, line icons, phone/laptop
 * product shots) is redrawn with native pptxgenjs shapes instead of bitmaps.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Palette - theme colours with lumMod/lumOff/alpha already flattened
 * ------------------------------------------------------------------ */
const C = {
  white: 'FFFFFF',
  black: '000000',
  wash: 'F4F4F3',     // accent1 20/80 @75% alpha over white (cover wash)
  ink: '262626',      // tx1 85/15  - body + headline text
  charcoal: '302E2B', // accent1 lumMod 25%
  gray: 'BAB7B2',     // accent1
  grayDk: '8F8A82',   // accent1 lumMod 75%
  rail: 'F2F2F2',     // bg2 lumMod 95% (skill-bar track)
  tan: 'BCA58E',      // accent4
  tanDk: '9B7C5C',    // accent4 lumMod 75%
  cream: 'FAFAF0',    // accent2
  creamDk: 'E8E8B9',  // accent2 lumMod 85% (gauge track)
  brown: '553717',    // accent3 lumMod 25% (arrow fold)
};

const HEAD = 'Public Sans'; // +mj-lt
const BODY = 'Work Sans';   // +mn-lt

/* Boilerplate copy, reused across the template */
const LEAD_LONG =
  'A well-crafted resume highlights your skills, experience, and achievements, ' +
  'helping you stand out in a competitive job market. Clarity, structure, and ' +
  'relevance are key to making a lasting impact on employers.';
const LEAD =
  'A well-crafted resume highlights your skills, experience, and achievements, ' +
  'helping you stand out in a competitive job market. ';
const TAG = 'A resume is your first impression\u2014make it count.';
const TAG_SHORT = 'A resume is your first impression';
const SUB = 'Subtitle Here';

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'W16x9', width: 13.333, height: 7.5 });
pptx.layout = 'W16x9';

/* ------------------------------------------------------------------ *
 * Drawing helpers
 * ------------------------------------------------------------------ */
function rect(s, x, y, w, h, fill, line) {
  s.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: fill }, line });
}

function ellipse(s, x, y, w, h, fill) {
  s.addShape(pptx.ShapeType.ellipse, { x, y, w, h, fill: { color: fill } });
}

function roundRect(s, x, y, w, h, fill, radius) {
  s.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: radius || 0, fill: { color: fill } });
}

/** Open stroked polyline from fractional (0..1) coordinates - used for icon art. */
function stroke(s, x, y, w, h, pts, color, width) {
  s.addShape(pptx.ShapeType.custGeom, {
    x, y, w, h, line: { color, width: width || 1.5 },
    points: pts.map(p => ({ x: p[0] * w, y: p[1] * h })),
  });
}

/** Filled polygon from fractional (0..1) coordinates. */
function poly(s, x, y, w, h, pts, fill) {
  s.addShape(pptx.ShapeType.custGeom, {
    x, y, w, h, fill: { color: fill },
    points: pts.map(p => ({ x: p[0] * w, y: p[1] * h })).concat([{ close: true }]),
  });
}

/** Plain top-anchored text box, matching the template's text-box defaults. */
function txt(s, text, x, y, w, h, o = {}) {
  s.addText(text, {
    x, y, w, h,
    fontFace: o.face || BODY,
    fontSize: o.size || 12,
    color: o.color || C.ink,
    bold: !!o.bold,
    align: o.align || 'left',
    valign: o.valign || 'top',
    lineSpacingMultiple: o.lsp,
    charSpacing: o.spc,
    margin: o.margin,
    rotate: o.rotate,
  });
}

/** Small bold caption that sits above every section headline. */
function eyebrow(s, text, x, y, w, o = {}) {
  txt(s, text, x, y, w, 0.303, { size: 12, bold: true, color: o.color || C.tan, align: o.align });
}

/** Two-tone display headline: [text, colour, breakLineAfter?] runs. */
function headline(s, parts, x, y, w, h, o = {}) {
  s.addText(
    parts.map(p => ({ text: p[0], options: { color: p[1] || C.ink, breakLine: !!p[2] } })),
    {
      x, y, w, h, fontFace: o.face || HEAD, fontSize: o.size || 36,
      align: o.align || 'left', valign: 'top', lineSpacingMultiple: o.lsp,
    }
  );
}

/** "Subtitle Here" + tagline pair, the deck's most repeated text unit. */
function subBlock(s, x, y, w, color, tagline) {
  txt(s, SUB, x, y, w, 0.324, { size: 12, bold: true, color, lsp: 1.2 });
  txt(s, tagline === undefined ? TAG : tagline, x, y + 0.327, w, 0.567, { size: 12, color, lsp: 1.2 });
}

/** Page number, bottom-left, as defined on the slide master. */
function pageNum(s, n) {
  txt(s, String(n), 0, 6.951, 0.76, 0.399, {
    size: 12, bold: true, color: C.tan, align: 'center', valign: 'middle',
  });
}

/* --- icon art: each helper draws inside a d x d box at (x, y) ------- */
/** Rotatable rounded bar - the building block for most of the icon art. */
function bar(s, x, y, w, h, color, rot) {
  s.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, rectRadius: Math.min(w, h) / 2, fill: { color }, rotate: rot || 0,
  });
}

/** Four-pointed sparkle. */
function star(s, x, y, d, color) {
  s.addShape(pptx.ShapeType.star4, { x, y, w: d, h: d, fill: { color }, line: { color, width: 0.5 } });
}

/**
 * Line-art icons. Each entry draws inside a d x d box anchored at (x, y),
 * standing in for the template's embedded SVG/PNG icons.
 */
const ICON = {
  chevronR: (s, x, y, d, c) => stroke(s, x + d * 0.38, y + d * 0.26, d * 0.22, d * 0.48, [[0, 0], [1, 0.5], [0, 1]], c, d * 4),
  chevronL: (s, x, y, d, c) => stroke(s, x + d * 0.40, y + d * 0.26, d * 0.22, d * 0.48, [[1, 0], [0, 0.5], [1, 1]], c, d * 4),
  check: (s, x, y, d, c) => stroke(s, x + d * 0.27, y + d * 0.40, d * 0.46, d * 0.22, [[0, 0.2], [0.38, 1], [1, 0]], c, d * 4),

  pin: (s, x, y, d, c, hole) => {
    ellipse(s, x + d * 0.22, y + d * 0.08, d * 0.56, d * 0.56, c);
    poly(s, x + d * 0.33, y + d * 0.44, d * 0.34, d * 0.48, [[0, 0], [1, 0], [0.5, 1]], c);
    ellipse(s, x + d * 0.40, y + d * 0.26, d * 0.20, d * 0.20, hole);
  },
  phone: (s, x, y, d, c) => {
    s.addShape(pptx.ShapeType.blockArc, {
      x: x + d * 0.02, y: y + d * 0.14, w: d * 0.68, h: d * 0.68, fill: { color: c },
      angleRange: [205, 20], arcThicknessRatio: 0.66, rotate: 232,
    });
    stroke(s, x + d * 0.56, y + d * 0.22, d * 0.14, d * 0.22, [[0, 0], [1, 0.5], [0, 1]], c, d * 1.8);
    stroke(s, x + d * 0.68, y + d * 0.12, d * 0.16, d * 0.42, [[0, 0], [1, 0.5], [0, 1]], c, d * 1.8);
  },
  mail: (s, x, y, d, c) => {
    roundRect(s, x + d * 0.10, y + d * 0.26, d * 0.80, d * 0.50, c, d * 0.10);
    stroke(s, x + d * 0.20, y + d * 0.34, d * 0.60, d * 0.22, [[0, 0], [0.5, 1], [1, 0]], C.gray, d * 2.6);
  },
  globe: (s, x, y, d, c) => {
    s.addShape(pptx.ShapeType.ellipse, {
      x: x + d * 0.08, y: y + d * 0.08, w: d * 0.78, h: d * 0.78, line: { color: c, width: d * 2.6 },
    });
    stroke(s, x + d * 0.10, y + d * 0.47, d * 0.74, 0, [[0, 0], [1, 0]], c, d * 2.2);
    s.addShape(pptx.ShapeType.ellipse, {
      x: x + d * 0.32, y: y + d * 0.08, w: d * 0.30, h: d * 0.78, line: { color: c, width: d * 2.2 },
    });
    poly(s, x + d * 0.52, y + d * 0.52, d * 0.38, d * 0.42, [[0, 0], [1, 0.45], [0.45, 0.6], [0.3, 1]], c);
  },

  wand: (s, x, y, d, c) => {
    bar(s, x + d * 0.02, y + d * 0.40, d * 0.62, d * 0.17, c, -40);
    star(s, x + d * 0.44, y + d * 0.02, d * 0.24, c);
    star(s, x + d * 0.68, y + d * 0.20, d * 0.20, c);
    star(s, x + d * 0.66, y + d * 0.50, d * 0.22, c);
  },
  brush: (s, x, y, d, c) => {
    bar(s, x + d * 0.30, y + d * 0.10, d * 0.16, d * 0.60, c, 40);
    poly(s, x + d * 0.14, y + d * 0.56, d * 0.28, d * 0.30, [[1, 0], [1, 1], [0, 1]], c);
    stroke(s, x + d * 0.60, y + d * 0.14, d * 0.24, d * 0.20, [[0, 1], [1, 0]], c, d * 2.4);
    stroke(s, x + d * 0.60, y + d * 0.62, d * 0.24, d * 0.22, [[0, 0], [1, 1]], c, d * 2.4);
  },
  bulb: (s, x, y, d, c) => {
    s.addShape(pptx.ShapeType.ellipse, {
      x: x + d * 0.28, y: y + d * 0.20, w: d * 0.42, h: d * 0.42, line: { color: c, width: d * 2.6 },
    });
    stroke(s, x + d * 0.40, y + d * 0.58, d * 0.18, d * 0.18, [[0, 0], [0, 1], [1, 1], [1, 0]], c, d * 2.6);
    stroke(s, x + d * 0.38, y + d * 0.84, d * 0.22, 0, [[0, 0], [1, 0]], c, d * 2.4);
    [[0.44, 0.02, 0.06, 0.10], [0.20, 0.16, 0.10, 0.08], [0.72, 0.16, 0.10, 0.08]].forEach(([fx, fy, fw, fh]) =>
      stroke(s, x + d * fx, y + d * fy, d * fw, d * fh, [[0, 0], [1, 1]], c, d * 1.8));
  },
  bars: (s, x, y, d, c) => {
    [[0.16, 0.40], [0.31, 0.26], [0.46, 0.12], [0.61, 0.32], [0.76, 0.24]].forEach(([fx, ty]) =>
      bar(s, x + d * fx, y + ty * d, d * 0.09, d * (0.90 - ty), c));
  },
  brain: (s, x, y, d, c) => {
    poly(s, x + d * 0.12, y + d * 0.12, d * 0.76, d * 0.78,
      [[0.45, 0], [0.85, 0.3], [1, 0.75], [0.7, 0.75], [0.7, 1], [0.15, 1], [0, 0.55]], c);
    ellipse(s, x + d * 0.24, y + d * 0.26, d * 0.36, d * 0.36, C.white);
    ellipse(s, x + d * 0.34, y + d * 0.36, d * 0.16, d * 0.16, c);
    stroke(s, x + d * 0.42, y + d * 0.14, 0, d * 0.16, [[0, 0], [0, 1]], c, d * 2.2);
  },

  camera: (s, x, y, d, c) => {
    roundRect(s, x + d * 0.10, y + d * 0.42, d * 0.52, d * 0.36, c, d * 0.10);
    poly(s, x + d * 0.64, y + d * 0.46, d * 0.24, d * 0.28, [[1, 0], [1, 1], [0, 0.55]], c);
    ellipse(s, x + d * 0.14, y + d * 0.14, d * 0.20, d * 0.20, c);
    ellipse(s, x + d * 0.38, y + d * 0.14, d * 0.20, d * 0.20, c);
  },
  film: (s, x, y, d, c) => {
    s.addShape(pptx.ShapeType.roundRect, {
      x: x + d * 0.28, y: y + d * 0.10, w: d * 0.60, h: d * 0.58, rectRadius: d * 0.10,
      line: { color: c, width: d * 2.6 },
    });
    [0.38, 0.60, 0.78].forEach(fx =>
      stroke(s, x + d * fx, y + d * 0.16, 0, d * 0.46, [[0, 0], [0, 1]], c, d * 1.6));
    stroke(s, x + d * 0.10, y + d * 0.28, d * 0.42, d * 0.56, [[0, 0], [0, 0.75], [1, 1]], c, d * 2.6);
  },
  pencil: (s, x, y, d, c) => {
    stroke(s, x + d * 0.08, y + d * 0.14, d * 0.60, d * 0.62,
      [[0.30, 0], [0, 0], [0, 0.75], [0.30, 1], [0.45, 1.2], [0.60, 1], [1, 1], [1, 0.55]], c, d * 2.6);
    bar(s, x + d * 0.50, y + d * 0.08, d * 0.14, d * 0.46, c, 40);
  },
  node: (s, x, y, d, c) => {
    s.addShape(pptx.ShapeType.rect, {
      x: x + d * 0.20, y: y + d * 0.20, w: d * 0.60, h: d * 0.60, line: { color: c, width: d * 2.6 },
    });
    [[0.08, 0.08], [0.72, 0.08], [0.08, 0.72], [0.72, 0.72]].forEach(([fx, fy]) => {
      ellipse(s, x + d * fx, y + d * fy, d * 0.20, d * 0.20, c);
      ellipse(s, x + d * (fx + 0.05), y + d * (fy + 0.05), d * 0.10, d * 0.10, C.white);
    });
    s.addShape(pptx.ShapeType.roundRect, {
      x: x + d * 0.30, y: y + d * 0.34, w: d * 0.24, h: d * 0.16, rectRadius: d * 0.04,
      line: { color: c, width: d * 2 },
    });
    stroke(s, x + d * 0.44, y + d * 0.54, d * 0.16, d * 0.14, [[1, 0], [1, 1], [0, 1]], c, d * 2);
  },
  music: (s, x, y, d, c) => {
    s.addShape(pptx.ShapeType.roundRect, {
      x: x + d * 0.34, y: y + d * 0.10, w: d * 0.54, h: d * 0.44, rectRadius: d * 0.10,
      line: { color: c, width: d * 2.6 },
    });
    stroke(s, x + d * 0.40, y + d * 0.30, d * 0.42, d * 0.18, [[0, 1], [0.4, 0.2], [0.7, 0.6], [1, 0.25]], c, d * 1.8);
    ellipse(s, x + d * 0.42, y + d * 0.68, d * 0.22, d * 0.20, c);
    stroke(s, x + d * 0.62, y + d * 0.42, 0, d * 0.36, [[0, 1], [0, 0]], c, d * 2.4);
    stroke(s, x + d * 0.62, y + d * 0.42, d * 0.20, d * 0.10, [[0, 0], [1, 1]], c, d * 2.4);
    stroke(s, x + d * 0.10, y + d * 0.52, 0, d * 0.36, [[0, 0], [0, 1]], c, d * 2.6);
  },
  clapper: (s, x, y, d, c) => {
    poly(s, x + d * 0.08, y + d * 0.22, d * 0.82, d * 0.22,
      [[0, 0.7], [0.06, 0], [1, 0.3], [0.94, 1]], c);
    [0.28, 0.50, 0.72].forEach(fx =>
      stroke(s, x + d * (fx + 0.05), y + d * 0.25, -d * 0.06, d * 0.17, [[0, 0], [1, 1]], C.white, d * 1.4));
    s.addShape(pptx.ShapeType.roundRect, {
      x: x + d * 0.14, y: y + d * 0.46, w: d * 0.70, h: d * 0.38, rectRadius: d * 0.08,
      line: { color: c, width: d * 2.6 },
    });
  },
};

/** Corner "Logoipsum" lockup used on the cover and the closing slide. */
function logoipsum(s, x, y) {
  ellipse(s, x, y - 0.015, 0.26, 0.26, C.charcoal);
  ellipse(s, x + 0.085, y + 0.07, 0.09, 0.09, C.wash);
  txt(s, 'Logoipsum', x + 0.3, y - 0.05, 0.9, 0.33, {
    face: HEAD, size: 12, bold: true, color: C.charcoal, valign: 'middle', margin: 0,
  });
}

/**
 * Split wordmark: the cover and closing titles are cut by a shallow diagonal,
 * approximated here by colouring the leading letters one tone and the trailing
 * letters the other.
 */
function wordmark(s, head, tail, headColor, tailColor, x, y, w, h, size, spc) {
  s.addText(
    [{ text: head, options: { color: headColor } }, { text: tail, options: { color: tailColor } }],
    {
      x, y, w, h, fontFace: HEAD, fontSize: size, bold: true,
      charSpacing: spc, valign: 'middle', align: 'left', margin: 0,
    }
  );
}

/** Pill button. `outline` renders the ghost variant used next to it. */
function button(s, label, x, y, w, h, o = {}) {
  if (o.outline) {
    roundRect(s, x, y, w, h, C.white, 0);
    s.addShape(pptx.ShapeType.roundRect, {
      x, y, w, h, rectRadius: 0, line: { color: C.tan, width: 1.5 },
    });
    txt(s, label, x, y, w, h, {
      face: HEAD, size: 14, bold: true, color: C.tan, align: 'center', valign: 'middle',
    });
  } else {
    rect(s, x, y, w, h, C.tan);
    txt(s, label, x + 0.151, y + 0.148, w - 0.42, 0.303, { size: 12, bold: true, color: C.white });
    ICON.chevronR(s, x + w - 0.49, y + h / 2 - 0.163, 0.326, C.white);
  }
}

/* ================================================================== *
 * 1 - Cover: giant RESUME wordmark
 * ================================================================== */
function slide01() {
  const s = pptx.addSlide();
  rect(s, 0, 0, 13.333, 7.5, C.wash);
  wordmark(s, 'RES', 'UME', C.tan, C.charcoal, 0.82, 2.359, 11.6, 2.491, 152.7, 5.1);
  txt(s, 'PORTFOLIO PRESENTATION TEMPLATE', 0.931, 6.4, 8.323, 0.303, {
    size: 12, color: C.charcoal, spc: 3,
  });
  logoipsum(s, 11.889, 0.297);
  pageNum(s, 1);
}

/* ================================================================== *
 * 2 - About Me + two numbered cards
 * ================================================================== */
function slide02() {
  const s = pptx.addSlide();
  eyebrow(s, 'About Me', 6.381, 1.197, 5.906);
  headline(s, [['Welcome to ', C.ink], ['My', C.tan], [' ', C.ink], ['Professional', C.tan], [' Story', C.ink]],
    6.381, 1.5, 5.265, 1.498, { lsp: 1.2 });
  txt(s, LEAD_LONG, 6.404, 3.223, 5.052, 1.051, { lsp: 1.2 });

  [['01', 3.973, C.tan], ['02', 8.624, C.gray]].forEach(([num, x, fill]) => {
    rect(s, x, 4.965, 3.937, 1.492, fill);
    txt(s, num, x + 0.295, 5.391, 2.999, 0.64, { face: HEAD, size: 32, color: C.white });
    subBlock(s, x + 1.119, 5.264, 2.648, C.white);
  });
  pageNum(s, 2);
}

/* ================================================================== *
 * 3 - Who I Am + two name plates
 * ================================================================== */
function slide03() {
  const s = pptx.addSlide();
  [[0.76, 0.797, C.tan], [4.179, 5.497, C.gray]].forEach(([x, y, fill]) => {
    rect(s, x + 0.271, y, 3.149, 1.206, fill);
    txt(s, 'Your Name Here', x, y + 0.318, 3.69, 0.269, {
      face: HEAD, size: 16, color: C.white, align: 'center', margin: 0,
    });
    txt(s, 'Your Title Here', x, y + 0.686, 3.69, 0.202, {
      size: 12, color: C.white, align: 'center', margin: 0,
    });
  });

  eyebrow(s, 'About Me', 8.431, 1.843, 3.677);
  headline(s, [['Who', C.tan], [' I Am', C.ink]], 8.431, 2.145, 3.455, 0.983, { size: 48, lsp: 1.2 });
  txt(s, LEAD_LONG, 8.465, 3.447, 4.016, 1.294, { lsp: 1.2 });
  button(s, 'Learn More', 8.543, 5.059, 1.846, 0.598);
  pageNum(s, 3);
}

/* ================================================================== *
 * 4 - Career Objective + skill meters
 * ================================================================== */
function slide04() {
  const s = pptx.addSlide();
  rect(s, 0, 0, 5.043, 7.5, C.gray);
  rect(s, 9.019, 1.388, 3.554, 4.724, C.white, { color: C.gray, width: 3 });
  rect(s, 4.314, 1.388, 4.468, 4.724, C.tan);

  headline(s, [['Career', C.white, true], ['Objective', C.white]], 5.043, 1.922, 4.136, 1.489, { lsp: 1.2 });
  eyebrow(s, 'My Professional Goals', 5.043, 3.807, 5.906, { color: C.white });
  txt(s, LEAD, 5.043, 4.328, 3.015, 1.052, { color: C.white, lsp: 1.2 });

  [[2.083, '90%', 2.126, C.tan], [3.028, '75%', 1.772, C.gray],
   [3.987, '90%', 2.126, C.gray], [4.932, '75%', 1.772, C.gray]].forEach(([y, pct, filled, color]) => {
    txt(s, 'Skill Here', 9.565, y, 1.619, 0.324, { lsp: 1.2 });
    txt(s, pct, 10.981, y, 1.036, 0.384, { face: HEAD, size: 14, align: 'right', lsp: 1.2 });
    s.addShape(pptx.ShapeType.line, { x: 9.654, y: y + 0.528, w: 2.362, h: 0, line: { color: C.rail, width: 8 } });
    s.addShape(pptx.ShapeType.line, { x: 9.654, y: y + 0.528, w: filled, h: 0, line: { color, width: 8 } });
  });
  pageNum(s, 4);
}

/* ================================================================== *
 * 5 - A Brief Personal Introduction + footer band
 * ================================================================== */
function slide05() {
  const s = pptx.addSlide();
  rect(s, 0, 4.587, 13.333, 2.116, C.gray);
  eyebrow(s, 'Introduction', 3.714, 0.891, 5.906, { align: 'center' });
  headline(s, [['A', C.tan], [' Brief ', C.ink], ['Personal', C.tan], [' Introduction', C.ink]],
    3.027, 1.15, 7.28, 0.762, { align: 'center', lsp: 1.2 });

  [[0.968, 'brush'], [7.268, 'bulb']].forEach(([x, kind]) => {
    rect(s, x, 5.829, 0.257, 0.257, C.grayDk);
    ICON[kind](s, x + 0.132, 5.757, 0.386, C.white);
    subBlock(s, x + 0.933, 5.641, 4.371, C.white);
  });
  pageNum(s, 5);
}

/* ================================================================== *
 * 6 - My Career Path, three steps
 * ================================================================== */
function slide06() {
  const s = pptx.addSlide();
  eyebrow(s, 'Work Experience Overview', 0.76, 0.797, 11.812, { align: 'center', color: C.gray });
  headline(s, [['My ', C.ink], ['Career', C.tan], [' Path', C.ink]],
    0.76, 1.1, 11.812, 0.707, { align: 'center' });

  const steps = [
    { x: 1.291, ty: 2.405, tag: TAG_SHORT, tx: 1.291, tw: 3.15, num: '1', nx: 2.546, ny: 5.939, ns: 0.64, nf: C.gray, nz: 24 },
    { x: 5.092, ty: 2.165, tag: TAG, tx: 5.300, tw: 2.733, num: '2', nx: 6.273, ny: 6.006, ns: 0.787, nf: C.tan, nz: 32 },
    { x: 8.893, ty: 2.405, tag: TAG_SHORT, tx: 8.893, tw: 3.15, num: '3', nx: 10.148, ny: 5.939, ns: 0.64, nf: C.gray, nz: 24 },
  ];
  steps.forEach(c => {
    txt(s, SUB, c.x, c.ty, 3.15, 0.324, { size: 12, bold: true, align: 'center', lsp: 1.2 });
    txt(s, c.tag, c.tx, c.ty + 0.327, c.tw, 0.567, { size: 12, align: 'center', lsp: 1.2 });
    rect(s, c.nx, c.ny, c.ns, c.ns, c.nf);
    txt(s, c.num, c.nx - 0.09, c.ny, c.ns + 0.18, c.ns, {
      size: c.nz, color: C.white, align: 'center', valign: 'middle',
    });
  });
  pageNum(s, 6);
}

/* ================================================================== *
 * 7 - My Academic Journey, three right-aligned steps
 * ================================================================== */
function slide07() {
  const s = pptx.addSlide();
  rect(s, 0.76, 0.694, 5.906, 1.575, C.gray);

  [{ y: 1.033, color: C.white, badge: C.gray, chip: '01', by: 1.284 },
   { y: 3.302, color: C.ink, badge: C.tan, chip: '02', by: 3.553 },
   { y: 5.571, color: C.ink, badge: C.tan, chip: '03', by: 5.822 }].forEach(r => {
    txt(s, SUB, 1.451, r.y, 2.938, 0.324, { size: 12, bold: true, color: r.color, align: 'right', lsp: 1.2 });
    txt(s, TAG, 1.451, r.y + 0.33, 2.938, 0.567, { size: 12, color: r.color, align: 'right', lsp: 1.2 });
    roundRect(s, 6.638, r.by, 0.394, 0.394, r.badge);
    txt(s, r.chip, 6.513, r.by + 0.046, 0.643, 0.303, {
      face: HEAD, size: 12, color: C.white, align: 'center', margin: 0,
    });
  });

  eyebrow(s, 'Educational Background', 7.982, 1.524, 5.906);
  headline(s, [['My Academic ', C.ink], ['Journey', C.tan]], 7.982, 1.827, 3.918, 1.489, { lsp: 1.2 });
  txt(s, LEAD_LONG, 8.001, 3.602, 4.176, 1.294, { lsp: 1.2 });
  button(s, 'Read More', 8.058, 5.386, 1.575, 0.591, { outline: true });
  button(s, 'Learn More', 9.836, 5.369, 1.846, 0.598);
  pageNum(s, 7);
}

/* ================================================================== *
 * 8 - What I Bring to the Table, three percentage tiles
 * ================================================================== */
function slide08() {
  const s = pptx.addSlide();
  eyebrow(s, 'Skills & Expertise', 4.942, 0.939, 5.906);
  headline(s, [['What', C.ink, true], ['I ', C.ink], ['Bring', C.tan, true], ['to the ', C.ink], ['Table', C.tan]],
    4.942, 1.242, 6.16, 2.216, { lsp: 1.2 });
  txt(s, LEAD, 4.942, 4.052, 6.314, 0.567, { lsp: 1.2 });

  [['75%', 4.869, C.tan], ['85%', 7.551, C.gray], ['90%', 10.234, C.gray]].forEach(([pct, x, fill]) => {
    rect(s, x, 5.07, 2.339, 1.633, fill);
    txt(s, pct, x, 5.396, 2.339, 0.539, { face: HEAD, size: 32, color: C.white, align: 'center', margin: 0 });
    txt(s, SUB, x, 6.175, 2.339, 0.202, { size: 12, bold: true, color: C.white, align: 'center', margin: 0 });
  });
  pageNum(s, 8);
}

/* ================================================================== *
 * 9 - Building Industry Connections, numbered list
 * ================================================================== */
function slide09() {
  const s = pptx.addSlide();
  rect(s, 7.178, 3.161, 5.395, 1.178, C.gray);
  eyebrow(s, 'Professional Networking', 7.588, 1.104, 5.906);
  headline(s, [['Building ', C.ink], ['Industry', C.tan], [' Connections', C.ink]],
    7.588, 1.481, 5.281, 1.313);

  [['01', 3.265, C.white, C.white], ['02', 4.46, C.ink, C.gray], ['03', 5.605, C.ink, C.gray]]
    .forEach(([num, y, textColor, numColor]) => {
      txt(s, num, 7.178, y + 0.131, 1.051, 0.64, { face: HEAD, size: 32, color: numColor, align: 'right' });
      subBlock(s, 8.373, y, 2.919, textColor);
    });
  pageNum(s, 9);
}

/* ================================================================== *
 * 10 - Personal Strengths, two arc gauges
 * ================================================================== */
function slide10() {
  const s = pptx.addSlide();
  eyebrow(s, 'Soft Skills', 7.279, 0.957, 5.906, { color: C.gray });
  headline(s, [['Personal ', C.black], ['Strengths', C.tan]], 7.279, 1.26, 4.176, 1.313, { face: BODY });

  // Cream track arc (near-full circle) with a white value arc drawn over it.
  [[2.916, C.gray, '60%', 131.5], [4.764, C.tan, '90%', 225.2]].forEach(([y, fill, pct, endAng]) => {
    rect(s, 7.279, y, 1.891, 1.575, fill);
    [[C.creamDk, 268.9], [C.white, endAng]].forEach(([color, ang]) => {
      s.addShape(pptx.ShapeType.arc, {
        x: 7.754, y: y + 0.316, w: 0.943, h: 0.943,
        line: { color, width: 10 }, angleRange: [270, ang],
      });
    });
    txt(s, pct, 7.788, y + 0.606, 0.874, 0.362, {
      size: 14, bold: true, color: C.white, align: 'center', lsp: 1.2,
    });
    subBlock(s, 9.461, y + 0.341, 2.919, C.ink);
  });
  pageNum(s, 10);
}

/* ================================================================== *
 * 11 - Career timeline 2024 / 2025 / 2026
 * ================================================================== */
function slide11() {
  const s = pptx.addSlide();
  s.addShape(pptx.ShapeType.line, { x: 8.752, y: 1.462, w: 0, h: 2.762, line: { color: C.gray, width: 1 } });

  [{ y: 0.797, label: '2024', boxed: true, dot: C.tan },
   { y: 2.354, label: '2025', boxed: false, dot: C.gray },
   { y: 3.921, label: '2026', boxed: false, dot: C.gray }].forEach(yr => {
    if (yr.boxed) rect(s, 6.19, yr.y, 2.562, 1.076, C.gray);
    txt(s, yr.label, 6.716, yr.y + 0.243, 1.454, 0.64, {
      face: HEAD, size: 32, color: yr.boxed ? C.white : C.ink,
    });
    ellipse(s, 8.621, yr.y + 0.428, 0.262, 0.238, yr.dot);
    subBlock(s, 9.579, yr.y + 0.125, 2.872, C.ink);
  });

  eyebrow(s, 'Career Timeline', 0.76, 5.54, 5.906, { color: C.gray });
  headline(s, [['A ', C.ink], ['Glimpse', C.tan], [' into My ', C.ink], ['Journey', C.tan]],
    0.76, 5.868, 6.948, 0.707);

  [[8.399, C.gray], [10.61, C.tan]].forEach(([x, fill]) => {
    rect(s, x, 5.332, 1.963, 1.371, fill);
    txt(s, '75%', x, 5.606, 1.963, 0.452, { face: HEAD, size: 32, color: C.white, align: 'center', margin: 0 });
    txt(s, SUB, x, 6.259, 1.963, 0.17, { size: 12, bold: true, color: C.white, align: 'center', margin: 0 });
  });
  pageNum(s, 11);
}

/* ================================================================== *
 * 12 - Awards & Achievements
 * ================================================================== */
function slide12() {
  const s = pptx.addSlide();
  rect(s, 8.527, 0, 3.937, 0.797, C.gray);
  rect(s, 8.527, 6.703, 3.937, 0.797, C.gray);
  rect(s, 0.76, 0.797, 6.762, 1.178, C.gray);
  txt(s, '01', 0.76, 1.032, 1.051, 0.64, { face: HEAD, size: 32, color: C.white, align: 'right' });
  subBlock(s, 1.956, 0.901, 2.919, C.white);

  eyebrow(s, 'Recognitions & Milestones', 8.586, 1.38, 5.906, { color: C.gray });
  headline(s, [['Awards & ', C.black], ['Achievements', C.tan]], 8.586, 1.683, 4.176, 1.313, { face: BODY });
  txt(s, LEAD, 8.613, 3.355, 3.96, 0.809, { lsp: 1.2 });

  [[4.651, C.tan], [5.531, C.gray]].forEach(([y, color]) => {
    ellipse(s, 8.613, y, 0.394, 0.394, color);
    ICON.check(s, 8.613, y, 0.394, C.white);
    txt(s, 'Your Subtitle Here', 9.068, y + 0.034, 2.125, 0.324, { size: 12, bold: true, lsp: 1.2 });
  });
  pageNum(s, 12);
}

/* ================================================================== *
 * 13 - Problem-Solving Skills, numbered chips
 * ================================================================== */
function slide13() {
  const s = pptx.addSlide();
  rect(s, 0, 0, 3.937, 0.797, C.gray);
  rect(s, 0, 6.703, 3.937, 0.797, C.gray);

  eyebrow(s, 'Overcoming Challenges', 0.76, 1.815, 4.445, { color: C.gray });
  headline(s, [['Problem-', C.ink, true], ['Solving ', C.ink], ['Skills', C.tan]],
    0.76, 2.118, 4.445, 1.313);
  txt(s, LEAD, 0.76, 4.392, 2.784, 1.294, { lsp: 1.2 });

  [['1', 1.081, C.tan], ['2', 3.149, C.gray], ['3', 5.216, C.gray]].forEach(([num, y, fill]) => {
    rect(s, 7.911, y, 0.657, 0.657, fill);
    txt(s, num, 7.656, y + 0.092, 1.172, 0.473, {
      size: 20, bold: true, color: C.white, align: 'center', lsp: 1.2,
    });
    subBlock(s, 9.082, y, 2.919, C.ink);
  });
  pageNum(s, 13);
}

/* ================================================================== *
 * 14 - Building Industry Connections, two icon cards
 * ================================================================== */
function slide14() {
  const s = pptx.addSlide();
  [[0.766, 3.947, C.tan, 'film'], [3.917, 0.797, C.gray, 'camera']].forEach(([x, y, fill, kind]) => {
    rect(s, x, y, 2.756, 2.756, fill);
    ICON[kind](s, x + 1.083, y + 0.463, 0.591, C.white);
    txt(s, SUB, x + 0.316, y + 1.399, 2.124, 0.343, {
      face: HEAD, size: 12, color: C.white, align: 'center', lsp: 1.2,
    });
    txt(s, TAG, x + 0.316, y + 1.727, 2.124, 0.809, {
      size: 12, color: C.white, align: 'center', lsp: 1.2,
    });
  });

  eyebrow(s, 'Professional Networking', 7.062, 0.923, 5.906, { color: C.gray });
  headline(s, [['Building ', C.ink], ['Industry', C.tan], [' Connections', C.ink]],
    7.062, 1.226, 5.906, 1.313);
  txt(s, LEAD, 7.062, 2.847, 5.716, 0.567, { lsp: 1.2 });
  pageNum(s, 14);
}

/* ================================================================== *
 * 15 - My Work in Action, full-bleed showcase
 * ================================================================== */
function slide15() {
  const s = pptx.addSlide();
  rect(s, 9.797, 0, 3.536, 7.5, C.gray);
  eyebrow(s, 'My Work in Action', 1.155, 0.743, 2.919, { color: C.gray });
  txt(s, TAG, 1.121, 6.42, 4.355, 0.324, { lsp: 1.2 });
  txt(s, 'by Aaron Loeb', 7.823, 6.432, 4.355, 0.324, { color: C.white, align: 'right', lsp: 1.2 });
  pageNum(s, 15);
}

/* ================================================================== *
 * 16 - Personal Branding, phone mockup between two arrow buttons
 * ================================================================== */
function slide16() {
  const s = pptx.addSlide();
  rect(s, 0, 0, 2.038, 7.5, C.gray);

  // Landscape phone: dark bezel wrapping a white screen.
  roundRect(s, 3.233, 2.24, 6.814, 3.387, C.charcoal, 0.42);
  roundRect(s, 3.333, 2.34, 6.614, 3.187, C.white, 0.34);

  eyebrow(s, 'How I Present Myself', 0.76, 0.797, 11.812, { align: 'center', color: C.gray });
  headline(s, [['Personal ', C.ink], ['Branding', C.tan]], 0.76, 1.1, 11.812, 0.707, { align: 'center' });

  ellipse(s, 0.76, 3.544, 0.787, 0.787, C.cream);
  ICON.chevronL(s, 0.76, 3.544, 0.787, C.gray);
  ellipse(s, 11.786, 3.544, 0.787, 0.787, C.tan);
  ICON.chevronR(s, 11.786, 3.544, 0.787, C.white);

  txt(s, LEAD, 3.167, 6.117, 7.0, 0.567, { align: 'center', lsp: 1.2 });
  pageNum(s, 16);
}

/* ================================================================== *
 * 17 - Insights from My Experience, laptop mockup
 * ================================================================== */
function slide17() {
  const s = pptx.addSlide();
  rect(s, 9.797, 0, 3.536, 7.5, C.gray);
  rect(s, 0, 3.699, 6.165, 1.594, C.gray);

  // Laptop: dark lid, white screen, light base bar.
  roundRect(s, 6.09, 2.42, 7.243, 3.99, C.charcoal, 0.1);
  rect(s, 6.28, 2.6, 3.517, 3.62, C.white);
  rect(s, 9.797, 2.6, 3.343, 3.62, C.gray);
  roundRect(s, 6.014, 6.41, 7.319, 0.22, C.grayDk, 0.08);

  eyebrow(s, ' Lessons Learned', 1.414, 0.939, 11.812, { color: C.gray });
  headline(s, [['Insights from My ', C.ink], ['Experience', C.tan]], 1.414, 1.242, 11.812, 0.707);

  [{ y: 2.42, chip: C.gray, color: C.ink, kind: 'wand' },
   { y: 3.965, chip: C.grayDk, color: C.white, kind: 'pencil' },
   { y: 5.51, chip: C.gray, color: C.ink, kind: 'node' }].forEach(r => {
    rect(s, 1.414, r.y + 0.12, 0.394, 0.394, r.chip);
    ICON[r.kind](s, 1.512, r.y + 0.218, 0.591, r.color === C.white ? C.cream : C.black);
    subBlock(s, 2.554, r.y, 2.919, r.color);
  });
  pageNum(s, 17);
}

/* ================================================================== *
 * 18 - What's Next for Me, four folded arrow bars
 * ================================================================== */
function slide18() {
  const s = pptx.addSlide();
  rect(s, 0, 3.823, 6.456, 1.353, C.gray);

  // Every bar is authored pointing right, then mirrored: chevron body,
  // folded "ribbon" corner, and a detached tail square at the far right.
  const mirrorX = pts => pts.map(([px, py]) => [1 - px, py]);
  const ARROW = mirrorX([[0.8778, 1], [0, 1], [0, 0], [0.8778, 0], [1, 0.5]]);
  [{ pct: '90%', bx: 7.652, by: 0.814, bw: 3.922, bh: 1.25, body: C.tan,
     fold: [11.574, 0.801, 0.875, 1.708, C.brown, [[0, 1], [0, 0.5061], [1, 0], [1, 0.7317]]],
     tail: [12.449, 1.678, 0.948, 0.844, C.tan] },
   { pct: '25%', bx: 9.495, by: 2.397, bw: 2.079, bh: 1.25, body: C.gray,
     fold: [11.574, 2.397, 0.875, 1.25, C.ink, [[1, 1], [0, 1], [0, 0.3667], [1, 0]]],
     tail: [12.449, 2.856, 0.948, 0.792, C.gray] },
   { pct: '50%', bx: 8.921, by: 3.926, bw: 2.652, bh: 1.25, body: C.gray,
     fold: [11.574, 3.926, 0.875, 1.25, C.ink, [[1, 1], [0, 0.6333], [0, 0], [1, 0]]],
     tail: [12.449, 3.926, 0.948, 0.792, C.gray] },
   { pct: '75%', bx: 8.429, by: 5.455, bw: 3.144, bh: 1.25, body: C.gray,
     fold: [11.574, 4.997, 0.875, 1.708, C.ink, [[1, 1], [0, 0.4939], [0, 0], [1, 0.2683]]],
     tail: [12.449, 4.997, 0.948, 0.844, C.gray] }].forEach(b => {
    rect(s, b.tail[0], b.tail[1], b.tail[2], b.tail[3], b.tail[4]);
    poly(s, b.fold[0], b.fold[1], b.fold[2], b.fold[3], mirrorX(b.fold[5]), b.fold[4]);
    poly(s, b.bx, b.by, b.bw, b.bh, ARROW, b.body);
    txt(s, b.pct, b.bx, b.by + 0.305, 1.969, 0.64, {
      size: 32, bold: true, color: C.white, align: 'center',
    });
  });

  eyebrow(s, 'Future Goals & Aspirations', 0.843, 1.043, 5.906, { color: C.gray });
  headline(s, [['What\u2019s', C.ink, true], ['Next', C.tan], [' for Me?', C.ink]],
    0.843, 1.346, 5.906, 1.313);
  txt(s, LEAD, 0.843, 2.847, 5.613, 0.567, { lsp: 1.2 });

  [[4.019, C.grayDk, C.white, 'bars'], [5.564, C.gray, C.ink, 'brain']].forEach(([y, chip, color, kind]) => {
    rect(s, 1.0, y + 0.12, 0.394, 0.394, chip);
    ICON[kind](s, 1.098, y + 0.218, 0.591, color === C.white ? C.cream : C.black);
    subBlock(s, 2.141, y, 2.919, color);
  });
  pageNum(s, 18);
}

/* ================================================================== *
 * 19 - Evolving with the Industry, two doughnut charts
 * ================================================================== */
function slide19() {
  const s = pptx.addSlide();
  rect(s, 7.443, 0, 5.89, 7.5, C.gray);
  rect(s, 0, 0.797, 3.312, 5.906, C.white);

  [[1.113, '80%', 8, 2, C.gray], [3.908, '60%', 6, 4, C.tan]].forEach(([y, label, a, b, color]) => {
    s.addChart(pptx.ChartType.doughnut,
      [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: [a, b] }],
      {
        x: 0.229, y, w: 2.855, h: 1.903,
        holeSize: 75, firstSliceAng: 0,
        showLegend: false, showTitle: false, showValue: false,
        chartColors: [color, C.cream],
        dataBorder: { pt: 1.5, color: C.white },
      });
    txt(s, label, 1.052, y + 0.765, 1.307, 0.438, { face: HEAD, size: 20, align: 'center' });
    txt(s, SUB, 0.574, y + 1.835, 2.165, 0.328, { size: 12, bold: true, align: 'center', lsp: 1.2 });
  });

  eyebrow(s, 'Adaptability & Growth', 8.906, 1.763, 3.492, { color: C.white });
  headline(s, [['Evolving with the Industry', C.white]], 8.906, 2.066, 3.492, 1.178, { size: 32 });
  txt(s, LEAD, 8.906, 3.548, 2.781, 1.294, { color: C.white, lsp: 1.2 });
  roundRect(s, 8.986, 5.146, 1.575, 0.591, C.tan);
  txt(s, 'Read More', 8.986, 5.146, 1.575, 0.591, {
    face: HEAD, size: 14, bold: true, color: C.white, align: 'center', valign: 'middle',
  });
  pageNum(s, 19);
}

/* ================================================================== *
 * 20 - Business & Entrepreneurship
 * ================================================================== */
function slide20() {
  const s = pptx.addSlide();
  rect(s, 8.796, 0, 4.537, 7.5, C.gray);

  eyebrow(s, 'My Entrepreneurial Experience', 1.284, 0.998, 5.906, { color: C.gray });
  headline(s, [['Business & ', C.ink], ['Entrepreneurship', C.tan]], 1.284, 1.301, 5.906, 1.313);
  txt(s, LEAD, 1.284, 3.036, 5.906, 0.567, { lsp: 1.2 });

  [[4.159, 'music'], [5.609, 'clapper']].forEach(([y, kind]) => {
    rect(s, 1.284, y + 0.12, 0.394, 0.394, C.gray);
    ICON[kind](s, 1.382, y + 0.218, 0.591, C.black);
    subBlock(s, 2.424, y, 2.919, C.ink);
  });
  pageNum(s, 20);
}

/* ================================================================== *
 * 21 - Let's Connect and Collaborate, contact rows
 * ================================================================== */
function slide21() {
  const s = pptx.addSlide();
  rect(s, 0, 0, 4.132, 6.703, C.gray);
  rect(s, 6.878, 2.621, 6.456, 0.975, C.gray);

  eyebrow(s, 'Get in Touch', 7.01, 0.797, 5.906, { color: C.gray });
  headline(s, [['Let\u2019s', C.tan], [' Connect', C.ink, true], ['and ', C.ink], ['Collaborate', C.tan]],
    7.01, 1.1, 5.906, 1.313);

  [{ y: 2.891, label: 'Location ', value: 'City, Country', color: C.white, chip: false, kind: 'pin' },
   { y: 3.904, label: 'Phone', value: '+62 947 1591 2578', color: C.ink, chip: true, kind: 'phone' },
   { y: 5.012, label: 'E-Mail', value: 'lorem@youremail.com', color: C.ink, chip: true, kind: 'mail' },
   { y: 6.12, label: 'Website', value: 'yourwebsite.com', color: C.ink, chip: true, kind: 'globe' }]
    .forEach(r => {
      const dy = r.chip ? 0.098 : 0.003;
      if (r.chip) roundRect(s, 7.158, r.y, 0.591, 0.591, C.gray);
      ICON[r.kind](s, 7.257, r.y + dy, 0.394, C.white, C.gray);
      txt(s, r.label, 8.056, r.y + (r.chip ? 0.096 : 0), 2.215, 0.399, {
        size: 16, bold: true, color: r.color, lsp: 1.2,
      });
      txt(s, r.value, 9.7, r.y + (r.chip ? 0.109 : 0.014), 2.67, 0.372, {
        size: 14, color: r.color, align: 'right', lsp: 1.2,
      });
    });
  pageNum(s, 21);
}

/* ================================================================== *
 * 22 - Thank You! closing wordmark
 * ================================================================== */
function slide22() {
  const s = pptx.addSlide();
  rect(s, 0, 0, 13.333, 7.5, C.wash);
  rect(s, 0, 0, 0.76, 7.5, C.gray);

  wordmark(s, 'THAN', 'K YOU!', C.charcoal, C.tanDk, 1.326, 2.835, 11.0, 1.474, 90.3, -1.7);
  txt(s, LEAD, 1.241, 4.405, 4.718, 0.808, { size: 14, color: C.black });
  txt(s, 'PORTFOLIO', 8.584, 3.296, 7.07, 0.909, {
    face: HEAD, size: 48, color: C.white, align: 'center', spc: 16, rotate: 90,
  });
  logoipsum(s, 1.364, 0.297);
  pageNum(s, 22);
}

/* ------------------------------------------------------------------ */
[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
 slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
 slide17, slide18, slide19, slide20, slide21, slide22].forEach(build => build());

pptx.writeFile({ fileName: path.join(__dirname, '03935622-b4e6-4645-8fa7-89e1fd6d423c_grok_final.pptx') })
  .then(f => console.log('wrote', f));
