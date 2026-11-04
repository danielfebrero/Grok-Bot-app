/**
 * "AGRIFY" pitch deck - 10 slides, 13.333 x 7.5 in (16:9).
 * Recreated with pptxgenjs. Photographs in the source deck are drawn here as
 * grey rounded-rectangle stand-ins labelled "[image]".
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const DARK = '547533'; // accent1 - deep olive green
const LIME = '8EB03A'; // accent2 - fresh green
const PAPER = 'F9F9F9'; // accent3 - off white (light slide background)
const INK = '595959'; // body copy / nav links on light slides (black @ 65% lum)
const INK_DARKER = '404040'; // the two paragraphs on slide 2 (black @ 75% lum)
const PHOTO = 'CCCCCC'; // image placeholder fill
const PHOTO_LABEL = 'A6A6A6';
const TRACK = 'AFABAB'; // progress-bar track

const HEAD = 'Montserrat'; // headings / numbers (bold)
const BODY = 'Montserrat Medium'; // paragraph copy

/* Lorem strings reused across the deck. */
const L_SHORT = 'Lorem ipsum dolor sit amet, consectetuer adipiscing.';
const L_HALF = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttmassa. ';
const L_MED = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttmassa. Fusce posuere, magna sed.';
const L_PULV = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttmassa. Fusce posuere, magna sed pulvinar';
const L_LONG = L_PULV + ' osuere.';
const L_LONGER = L_PULV + ' osuere, magna magna sed pulvinar.';
const L_XLONG = L_PULV + ' osuere, magna magna sed pulvinar osuere, magna porttmassa.';

/* ----------------------------------------------------------------- helpers */

/** Heading / label text. Placeholders in the deck use 90% line spacing. */
function label(slide, o) {
  slide.addText(o.text, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: o.face || HEAD, fontSize: o.size, color: o.color,
    bold: o.bold !== false, align: o.align || 'left', valign: 'top',
    lineSpacingMultiple: o.spacing === undefined ? 0.9 : o.spacing,
  });
}

/** 11pt paragraph copy - always 150% line spacing in this deck. */
function body(slide, o) {
  slide.addText(o.text, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: BODY, fontSize: o.size || 11, color: o.color,
    align: o.align || 'left', valign: o.valign || 'top', lineSpacingMultiple: 1.5,
  });
}

/**
 * Two-line section title: first line dark green, second line lime.
 * `spaceBefore` reproduces the 10pt paragraph gap of the source placeholders.
 */
function title(slide, o) {
  const runs = o.lines.map((line, i) => ({
    text: line.text,
    options: {
      color: line.color,
      breakLine: i < o.lines.length - 1,
      paraSpaceBefore: i > 0 ? o.spaceBefore : undefined,
    },
  }));
  slide.addText(runs, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: HEAD, fontSize: o.size, bold: true, valign: 'top',
    lineSpacingMultiple: o.spacing === undefined ? 0.9 : o.spacing,
  });
}

/** Rounded rectangle (radius in inches). */
function card(slide, pptx, o) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: { color: o.fill }, line: { type: 'none' }, rectRadius: o.r,
  });
}

/** Caption written across the middle of an image stand-in. */
function photoLabel(slide, x, y, w, h) {
  slide.addText('[image]', {
    x, y, w, h, fontFace: BODY, fontSize: 10, color: PHOTO_LABEL,
    align: 'center', valign: 'middle',
  });
}

/** Rounded-rectangle image stand-in (`r` = corner radius in inches). */
function photoRect(slide, pptx, x, y, w, h, r) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, fill: { color: PHOTO }, line: { type: 'none' }, rectRadius: r,
  });
  photoLabel(slide, x, y, w, h);
}

/** Fully rounded ("pill") image stand-in. */
function photoPill(slide, pptx, x, y, w, h) {
  photoRect(slide, pptx, x, y, w, h, Math.min(w, h) / 2);
}

/**
 * Image stand-in with a free-form outline. `pts` are offsets from the frame's
 * own top-left corner: [x, y] draws a straight segment, [x, y, c1x, c1y, c2x,
 * c2y] a cubic curve (used for the rounded and filleted corners).
 */
function photoPath(slide, pptx, x, y, w, h, pts) {
  const points = pts.map((p) => (p.length === 2
    ? { x: p[0], y: p[1] }
    : { x: p[0], y: p[1], curve: { type: 'cubic', x1: p[2], y1: p[3], x2: p[4], y2: p[5] } }));
  points.push({ close: true });
  slide.addShape(pptx.ShapeType.custGeom, {
    x, y, w, h, points, fill: { color: PHOTO }, line: { type: 'none' },
  });
  photoLabel(slide, x, y, w, h);
}

/**
 * Top navigation bar, repeated on every slide.
 * `light` = white lettering on the dark green title slides.
 */
function navBar(slide, pptx, light) {
  const linkColor = light ? PAPER : INK;
  label(slide, { x: 0.567, y: 0.576, w: 1.617, h: 0.33, text: 'AGRIFY', size: 14, color: light ? PAPER : LIME });
  [['Overview', 2.234, 0.586, 1.211, 0.266], ['Solution', 4.296, 0.577, 1.305, 0.266], ['Impact', 6.077, 0.577, 1.76, 0.304]]
    .forEach(([text, x, y, w, h]) => {
      label(slide, { x, y, w, h, text, size: 12, color: linkColor, bold: false, face: BODY, align: 'center' });
    });

  card(slide, pptx, { x: 10.711, y: 0.552, w: 2.055, h: 0.401, fill: light ? LIME : DARK, r: 0.091 });
  magnifier(slide, pptx, 10.857, 0.667, 0.16);
  label(slide, { x: 11.067, y: 0.613, w: 0.797, h: 0.304, text: 'Search', size: 12, color: PAPER, bold: false, face: BODY });
  leaf(slide, pptx, 12.446, 0.649, 0.173, 0.181);
}

/** Magnifying-glass glyph: a ring plus a short diagonal handle. */
function magnifier(slide, pptx, x, y, d) {
  slide.addShape(pptx.ShapeType.ellipse, {
    x, y, w: d * 0.78, h: d * 0.78,
    fill: { type: 'none' }, line: { color: PAPER, width: 1.25 },
  });
  slide.addShape(pptx.ShapeType.line, {
    x: x + d * 0.62, y: y + d * 0.62, w: d * 0.34, h: d * 0.34,
    line: { color: PAPER, width: 1.25 },
  });
}

/**
 * Sprout mark at the right edge of the search pill: a large blade leaning to
 * the upper right, a small leaflet on the left and the stem between them.
 */
function leaf(slide, pptx, x, y, w, h) {
  // Blade: a lens between tip (x1,y1) and base (x0,y0), bowed out either side.
  const blade = (x0, y0, x1, y1, bulge) => {
    const [dx, dy] = [(x1 - x0) * w, (y1 - y0) * h];
    const [nx, ny] = [-dy * bulge, dx * bulge];
    return {
      x, y, w, h, fill: { color: PAPER }, line: { type: 'none' },
      points: [
        { x: x1 * w, y: y1 * h },
        { x: x0 * w, y: y0 * h, curve: { type: 'cubic', x1: x1 * w + nx, y1: y1 * h + ny, x2: x0 * w + nx, y2: y0 * h + ny } },
        { x: x1 * w, y: y1 * h, curve: { type: 'cubic', x1: x0 * w - nx, y1: y0 * h - ny, x2: x1 * w - nx, y2: y1 * h - ny } },
        { close: true },
      ],
    };
  };
  slide.addShape(pptx.ShapeType.custGeom, blade(0.4, 0.66, 1, 0.04, 0.28));
  slide.addShape(pptx.ShapeType.custGeom, blade(0.42, 0.74, 0.04, 0.52, 0.3));
  slide.addShape(pptx.ShapeType.line, {
    x: x + 0.22 * w, y: y + h, w: 0.22 * w, h: 0.26 * h,
    line: { color: PAPER, width: 1.25 }, flipV: true,
  });
}

/** ">" chevron pointing at (cx, cy); negative `w` points it the other way. */
function chevron(slide, pptx, cx, cy, w, h, color, weight) {
  const line = { color, width: weight || 1 };
  const x = Math.min(cx, cx - w);
  const back = w < 0;
  slide.addShape(pptx.ShapeType.line, { x, y: cy - h, w: Math.abs(w), h, line, flipV: back });
  slide.addShape(pptx.ShapeType.line, { x, y: cy, w: Math.abs(w), h, line, flipV: !back });
}

/** Pill button with a circled arrow beside it (slides 1 and 10). */
function navButton(slide, pptx, o) {
  card(slide, pptx, { x: o.btnX, y: 6.046, w: 2.449, h: 0.478, fill: LIME, r: 0.239 });
  slide.addText(o.text, {
    x: o.btnX - 0.007, y: 6.155, w: 2.449, h: 0.337,
    fontFace: 'Poppins', fontSize: 14, color: PAPER, align: 'center', valign: 'top',
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: o.circleX, y: 6.046, w: 0.478, h: 0.478,
    fill: { type: 'none' }, line: { color: LIME, width: 1 },
  });
  // Slim arrow glyph: a shaft with an open chevron head.
  const tipX = o.circleX + (o.back ? 0.127 : 0.352);
  slide.addShape(pptx.ShapeType.line, {
    x: o.circleX + 0.127, y: 6.285, w: 0.225, h: 0, line: { color: LIME, width: 1.75 },
  });
  chevron(slide, pptx, tipX, 6.285, o.back ? -0.09 : 0.09, 0.088, LIME, 1.75);
}

/* ------------------------------------------------------------------ slides */

function slide1(pptx) {
  const s = pptx.addSlide();
  s.background = { color: DARK };
  navBar(s, pptx, true);

  title(s, {
    x: 0.682, y: 1.286, w: 12.085, h: 2.464, size: 72, spaceBefore: 10,
    lines: [{ text: 'Agriculture          For A', color: PAPER }, { text: 'Sustainable Future', color: PAPER }],
  });

  photoPill(s, pptx, 7.009, 1.325, 2.312, 0.988);

  card(s, pptx, { x: 0.742, y: 4.75, w: 2.984, h: 1.774, fill: LIME, r: 0.199 });
  label(s, { x: 0.957, y: 5.102, w: 1.424, h: 0.468, text: '20.5K', size: 28, color: PAPER });
  body(s, { x: 0.957, y: 5.659, w: 2.784, h: 0.633, text: L_SHORT, color: PAPER });

  // Stepped frame: full-height right column with a lower-left block attached.
  photoPath(s, pptx, 4.224, 3.788, 2.564, 2.736, [
    [1.466, 0], [2.341, 0], [2.56, 0.179, 2.449, 0, 2.539, 0.077], [2.564, 0.22],
    [2.564, 2.523], [2.348, 2.736, 2.56, 2.662, 2.453, 2.736], [0.216, 2.736],
    [0, 2.523, 0.111, 2.736, 0, 2.662], [0, 1.175],
    [0.216, 0.962, 0.024, 1.037, 0.111, 0.962], [0.87, 0.962],
    [1.242, 0.587, 0.87, 0.752, 1.035, 0.587], [1.242, 0.224],
    [1.466, 0, 1.242, 0.1, 1.342, 0],
  ]);
  // Stepped frame: wide band with a step cut out of the bottom-right.
  photoPath(s, pptx, 7.244, 3.788, 5.362, 2.736, [
    [0.237, 0], [5.125, 0], [5.362, 0.242, 5.279, 0.027, 5.362, 0.125], [5.362, 1.743],
    [5.119, 1.985, 5.362, 1.876, 5.253, 1.985], [2.332, 1.985],
    [1.903, 2.428, 2.332, 2.236, 2.14, 2.428], [1.876, 2.507],
    [1.647, 2.736, 1.876, 2.633, 1.774, 2.736], [0.229, 2.736],
    [0, 2.507, 0.103, 2.736, 0, 2.633], [0, 0.242],
    [0.237, 0, 0, 0.125, 0.106, 0],
  ]);

  body(s, { x: 0.773, y: 3.75, w: 4.602, h: 0.633, text: L_MED, color: PAPER });
  navButton(s, pptx, { text: 'START PRESENTATION', btnX: 9.409, circleX: 12.14 });
}

function slide2(pptx) {
  const s = pptx.addSlide();
  s.background = { color: PAPER };
  navBar(s, pptx, false);

  title(s, {
    x: 0.682, y: 1.286, w: 6.255, h: 1.589, size: 40, spaceBefore: 10,
    lines: [{ text: 'Say Hello to', color: DARK }, { text: 'Smarter Agriculture', color: LIME }],
  });
  body(s, { x: 0.742, y: 2.95, w: 5.454, h: 0.903, text: L_XLONG, color: INK_DARKER });
  body(s, { x: 7.11, y: 5.922, w: 5.454, h: 0.625, text: L_LONG, color: INK_DARKER });

  photoPill(s, pptx, 0.762, 4.25, 5.905, 2.297);
  photoRect(s, pptx, 7.137, 1.396, 5.63, 4.198, 0.324);
}

function slide3(pptx) {
  const s = pptx.addSlide();
  s.background = { color: PAPER };
  navBar(s, pptx, false);

  // Stepped frame: wide band on top, narrower column dropping down the left.
  photoPath(s, pptx, 0.832, 1.405, 4.918, 3.324, [
    [0.226, 0], [4.692, 0], [4.918, 0.226, 4.786, 0, 4.918, 0.101], [4.918, 1.256],
    [4.692, 1.481, 4.918, 1.38, 4.817, 1.481], [3.803, 1.481],
    [3.376, 1.909, 3.803, 1.718, 3.612, 1.909], [3.343, 1.909], [3.343, 3.108],
    [3.129, 3.324, 3.343, 3.213, 3.243, 3.324], [0.214, 3.324],
    [0, 3.108, 0.101, 3.324, 0, 3.213], [0, 0.226],
    [0.226, 0, 0, 0.101, 0.101, 0],
  ]);
  photoPill(s, pptx, 9.201, 5.128, 3.3, 1.418);

  title(s, {
    x: 6.077, y: 1.296, w: 6.861, h: 1.589, size: 40, spaceBefore: 10,
    lines: [{ text: 'Driven by Purpose,', color: DARK }, { text: 'Guided by Innovation', color: LIME }],
  });

  label(s, { x: 0.832, y: 5.054, w: 2.613, h: 0.789, text: '01 VISION', size: 32, color: DARK });
  body(s, { x: 0.832, y: 5.702, w: 3.463, h: 0.903, text: L_PULV, color: INK });
  label(s, { x: 9.158, y: 3.214, w: 3.248, h: 0.789, text: '02 MISSION', size: 32, color: LIME });
  body(s, { x: 9.158, y: 3.862, w: 3.463, h: 0.903, text: L_PULV, color: INK });

  photoRect(s, pptx, 4.703, 3.302, 3.938, 3.245, 0.236);
}

function slide4(pptx) {
  const s = pptx.addSlide();
  s.background = { color: PAPER };
  navBar(s, pptx, false);

  photoPill(s, pptx, 0.832, 1.385, 5.027, 1.466);
  photoRect(s, pptx, 0.832, 3.33, 5.027, 3.217, 0.21);

  title(s, {
    x: 6.589, y: 1.571, w: 5.972, h: 1.432, size: 40, spaceBefore: 10,
    lines: [{ text: 'Behind the', color: DARK }, { text: 'Declining Harvests', color: LIME }],
  });
  body(s, { x: 6.646, y: 3.139, w: 5.537, h: 0.633, text: L_LONG, color: INK, valign: 'middle' });

  const problems = [
    { x: 4.475, tx: 4.624, fill: DARK, num: '01', head: 'Limited access tools and data', headW: 3.582, numY: 4.325 },
    { x: 8.799, tx: 8.999, fill: LIME, num: '02', head: 'Climate change affecting', headW: 2.866, numY: 4.33 },
  ];
  problems.forEach((p) => {
    card(s, pptx, { x: p.x, y: 4.167, w: 3.825, h: 2.019, fill: p.fill, r: 0.191 });
    label(s, { x: p.tx, y: p.numY, w: 0.928, h: 0.572, text: p.num, size: 28, color: PAPER, spacing: 1 });
    label(s, { x: p.tx, y: 4.957, w: p.headW, h: 0.329, text: p.head, size: 14, color: PAPER });
    body(s, { x: p.tx, y: 5.317, w: 3.582, h: 0.625, text: L_HALF, color: PAPER });
  });
}

function slide5(pptx) {
  const s = pptx.addSlide();
  s.background = { color: PAPER };
  navBar(s, pptx, false);

  card(s, pptx, { x: 0.742, y: 3.435, w: 6.258, h: 1.583, fill: DARK, r: 0.264 });
  title(s, {
    x: 0.652, y: 1.64, w: 6.177, h: 1.447, size: 40, spacing: 1,
    lines: [{ text: 'Turning Problems', color: DARK }, { text: 'into Progress', color: LIME }],
  });

  const steps = [
    { num: '01', numX: 1.065, numY: 3.834, color: PAPER, head: 'Drone-assisted surveillance & spraying', headW: 4.429, headY: 3.74, bodyY: 4.078, bodyColor: PAPER },
    { num: '02', numX: 1.011, numY: 5.573, color: LIME, head: 'Access to real-time market data', headW: 4.35, headY: 5.494, bodyY: 5.833, bodyColor: INK },
  ];
  steps.forEach((it) => {
    label(s, { x: it.numX, y: it.numY, w: 1.012, h: 0.774, text: it.num, size: 40, color: it.color, spacing: 1 });
    label(s, { x: 2.178, y: it.headY, w: it.headW, h: 0.218, text: it.head, size: 14, color: it.color });
    body(s, { x: 2.178, y: it.bodyY, w: 4.576, h: 0.633, text: L_MED, color: it.bodyColor, valign: 'middle' });
  });

  // Top-rounded frame running off the bottom edge of the slide.
  s.addShape(pptx.ShapeType.round2SameRect, {
    x: 10.361, y: 2.181, w: 2.244, h: 5.319,
    fill: { color: PHOTO }, line: { type: 'none' }, rectRadius: 1.122,
  });
  photoLabel(s, 10.361, 2.181, 2.244, 5.319);
  photoPill(s, pptx, 7.602, 1.389, 2.244, 5.167);
}

function slide6(pptx) {
  const s = pptx.addSlide();
  s.background = { color: PAPER };
  navBar(s, pptx, false);

  title(s, {
    x: 0.8, y: 1.291, w: 5.736, h: 1.589, size: 44, spaceBefore: 10,
    lines: [{ text: 'Inside the Smart', color: DARK }, { text: 'Farming Toolbox', color: LIME }],
  });
  body(s, { x: 0.8, y: 2.99, w: 5.454, h: 0.633, text: L_PULV + ' osuere', color: INK });
  photoRect(s, pptx, 0.832, 4.125, 5.701, 2.422, 0.162);

  const tools = [
    { x: 6.901, y: 1.391, fill: DARK, name: 'Smart sensors', icon: 'tube' },
    { x: 9.939, y: 1.393, fill: LIME, name: 'Mobile app', icon: 'phone' },
    { x: 6.901, y: 4.127, fill: LIME, name: 'Marketplace', icon: 'cart' },
    { x: 9.935, y: 4.125, fill: DARK, name: 'Analytics', icon: 'chart' },
  ];
  tools.forEach((t) => {
    card(s, pptx, { x: t.x, y: t.y, w: 2.667, h: 2.359, fill: t.fill, r: 0.174 });
    toolIcon(s, pptx, t.icon, t.x + 0.282, t.y + 0.291, t.fill);
    label(s, { x: t.x + 0.15, y: t.y + 1.177, w: 1.942, h: 0.329, text: t.name, size: 14, color: PAPER });
    body(s, { x: t.x + 0.15, y: t.y + 1.507, w: 2.475, h: 0.625, text: L_SHORT, color: PAPER });
  });
}

/**
 * Outline pictograms built from primitives, white on the coloured cards.
 * `cardFill` is needed where a shape is knocked out of a solid glyph.
 */
function toolIcon(slide, pptx, kind, x, y, cardFill) {
  const stroke = { color: PAPER, width: 1.5 };
  const bar = (dx, dy, w, h) => slide.addShape(pptx.ShapeType.rect, {
    x: x + dx, y: y + dy, w, h, fill: { color: PAPER }, line: { type: 'none' },
  });
  if (kind === 'tube') {
    slide.addShape(pptx.ShapeType.round2SameRect, { x: x + 0.035, y: y + 0.055, w: 0.175, h: 0.4, fill: { type: 'none' }, line: stroke, rectRadius: 0.088, rotate: 180 });
    slide.addShape(pptx.ShapeType.roundRect, { x: x + 0.005, y, w: 0.235, h: 0.075, fill: { type: 'none' }, line: stroke, rectRadius: 0.022 });
    slide.addShape(pptx.ShapeType.round2SameRect, { x: x + 0.075, y: y + 0.29, w: 0.1, h: 0.13, fill: { color: PAPER }, line: { type: 'none' }, rectRadius: 0.05, rotate: 180 });
    [0.13, 0.19, 0.25].forEach((dy) => bar(0.135, dy, 0.075, 0.024)); // graduations
  } else if (kind === 'phone') {
    slide.addShape(pptx.ShapeType.roundRect, { x: x + 0.035, y, w: 0.205, h: 0.44, fill: { color: PAPER }, line: { type: 'none' }, rectRadius: 0.028 });
    slide.addShape(pptx.ShapeType.rect, { x: x + 0.055, y: y + 0.06, w: 0.165, h: 0.31, fill: { color: cardFill }, line: { type: 'none' } });
    slide.addShape(pptx.ShapeType.roundRect, { x: x + 0.115, y: y + 0.028, w: 0.045, h: 0.014, fill: { color: cardFill }, line: { type: 'none' }, rectRadius: 0.007 });
    slide.addShape(pptx.ShapeType.ellipse, { x: x + 0.127, y: y + 0.393, w: 0.022, h: 0.022, fill: { color: cardFill }, line: { type: 'none' } });
  } else if (kind === 'cart') {
    slide.addShape(pptx.ShapeType.line, { x: x + 0.02, y: y + 0.045, w: 0.075, h: 0, line: stroke });
    slide.addShape(pptx.ShapeType.line, { x: x + 0.09, y: y + 0.045, w: 0.135, h: 0.29, line: stroke });
    slide.addShape(pptx.ShapeType.line, { x: x + 0.21, y: y + 0.335, w: 0.29, h: 0, line: stroke });
    slide.addShape(pptx.ShapeType.trapezoid, { x: x + 0.155, y: y + 0.115, w: 0.36, h: 0.185, fill: { color: PAPER }, line: { type: 'none' }, rotate: 180 });
    [0.21, 0.4].forEach((dx) => {
      slide.addShape(pptx.ShapeType.ellipse, { x: x + dx, y: y + 0.375, w: 0.085, h: 0.085, fill: { color: PAPER }, line: { type: 'none' } });
    });
  } else if (kind === 'chart') {
    bar(0.0, 0.0, 0.045, 0.44); // y axis
    bar(0.0, 0.395, 0.5, 0.045); // x axis
    [[0.075, 0.245], [0.185, 0.32], [0.295, 0.185]].forEach(([dx, h]) => bar(dx, 0.395 - h, 0.09, h));
  }
}

function slide7(pptx) {
  const s = pptx.addSlide();
  s.background = { color: PAPER };
  navBar(s, pptx, false);

  photoRect(s, pptx, 0.832, 1.391, 4.48, 2.309, 0.255);
  photoRect(s, pptx, 0.832, 4.234, 4.48, 2.309, 0.255);
  photoRect(s, pptx, 7.109, 4.231, 5.454, 2.309, 0.228);

  title(s, {
    x: 6.995, y: 1.26, w: 4.834, h: 1.432, size: 40, spaceBefore: 10,
    lines: [{ text: 'Real Benefits', color: DARK }, { text: 'on the Field', color: LIME }],
  });
  body(s, { x: 7.006, y: 2.805, w: 5.454, h: 0.911, text: L_XLONG + ' ', color: INK });

  card(s, pptx, { x: 3.739, y: 2.227, w: 2.928, h: 3.48, fill: DARK, r: 0.287 });
  label(s, { x: 3.948, y: 2.585, w: 2.322, h: 0.218, text: 'Benefits for Farmers', size: 14, color: PAPER });

  const benefits = [
    ['Increased crop yield', 4.052, 3.081, 1.875],
    ['Cost-effective resource', 4.052, 3.585, 2.019],
    ['Reduced labor efforts', 4.052, 4.097, 1.927],
    ['Better market access', 4.047, 4.614, 1.88],
    ['Eco-friendly farming', 4.055, 5.123, 1.927],
  ];
  benefits.forEach(([text, ix, iy, w]) => {
    s.addShape(pptx.ShapeType.ellipse, { x: ix, y: iy, w: 0.227, h: 0.227, fill: { color: PAPER }, line: { type: 'none' } });
    chevron(s, pptx, ix + 0.142, iy + 0.113, 0.05, 0.045, DARK, 1);
    label(s, { x: ix + 0.289, y: iy - 0.012, w, h: 0.383, text, size: 11, color: PAPER, bold: false, face: BODY });
  });
}

function slide8(pptx) {
  const s = pptx.addSlide();
  s.background = { color: PAPER };
  navBar(s, pptx, false);

  s.addText(
    [{ text: 'Agrify ', options: { color: DARK } }, { text: 'Growth', options: { color: LIME } }],
    { x: 0.745, y: 1.26, w: 4.734, h: 0.798, fontFace: HEAD, fontSize: 40, bold: true, valign: 'top', lineSpacingMultiple: 0.9 }
  );
  body(s, { x: 0.649, y: 2.268, w: 5.537, h: 0.911, text: L_XLONG, color: INK, valign: 'middle' });
  body(s, { x: 0.649, y: 3.439, w: 5.537, h: 0.633, text: L_LONG, color: INK, valign: 'middle' });

  card(s, pptx, { x: 6.951, y: 1.408, w: 2.638, h: 1.545, fill: DARK, r: 0.159 });
  label(s, { x: 7.097, y: 1.838, w: 2.266, h: 0.286, text: 'Impact So Far', size: 14, color: PAPER });
  body(s, { x: 7.097, y: 2.155, w: 2.467, h: 0.625, text: L_SHORT, color: PAPER });

  const stats = [
    { x: 0.731, y: 4.485, tx: 0.885, ty: 4.697, vx: 0.893, vy: 5.033, vw: 1.861, fill: DARK, head: 'Farmers onboarded', value: '+5.5K', by: 5.748 },
    { x: 3.83, y: 4.497, tx: 3.984, ty: 4.709, vx: 3.992, vy: 5.045, vw: 1.413, fill: LIME, head: 'Regions accross', value: '+150', by: 5.759 },
  ];
  stats.forEach((st) => {
    card(s, pptx, { x: st.x, y: st.y, w: 2.651, h: 2.092, fill: st.fill, r: 0.227 });
    label(s, { x: st.tx, y: st.ty, w: 2.26, h: 0.218, text: st.head, size: 14, color: PAPER });
    label(s, { x: st.vx, y: st.vy, w: st.vw, h: 0.572, text: st.value, size: 28, color: PAPER, spacing: 1 });
    body(s, { x: st.tx, y: st.by, w: 3.098, h: 0.633, text: L_SHORT, color: PAPER, valign: 'middle' });
  });

  photoRect(s, pptx, 9.969, 1.393, 2.628, 3.104, 0.229);
  photoPill(s, pptx, 6.951, 3.323, 2.628, 1.162);

  const bars = [
    { y: 5.404, labelY: 4.871, fill: DARK, done: 5.029, name: 'Impact 01', nameW: 1.844, pct: '94%' },
    { y: 6.457, labelY: 5.917, fill: LIME, done: 4.412, name: 'Impact 02', nameW: 2.11, pct: '72%' },
  ];
  bars.forEach((b) => {
    card(s, pptx, { x: 6.952, y: b.y, w: 5.649, h: 0.132, fill: TRACK, r: 0.066 });
    card(s, pptx, { x: 6.952, y: b.y, w: b.done, h: 0.131, fill: b.fill, r: 0.066 });
    body(s, { x: 6.859, y: b.labelY, w: b.nameW, h: 0.425, text: b.name, size: 14, color: INK });
    body(s, { x: 11.982, y: b.labelY, w: 0.644, h: 0.425, text: b.pct, size: 14, color: INK, align: 'right' });
  });
}

function slide9(pptx) {
  const s = pptx.addSlide();
  s.background = { color: PAPER };
  navBar(s, pptx, false);

  title(s, {
    x: 6.794, y: 1.26, w: 5.922, h: 1.475, size: 40, spaceBefore: 10,
    lines: [{ text: 'The Future We\u2019re', color: DARK }, { text: 'Cultivating', color: LIME }],
  });

  const plans = [
    { num: '01', numX: 0.752, numY: 3.763, headY: 3.833, headW: 4.37, head: 'Collaborate with agri-finance partners', bodyY: 4.178, bodyH: 0.911, text: L_LONGER },
    { num: '02', numX: 0.754, numY: 5.454, headY: 5.547, headW: 4.258, head: 'Introduce AI-powered pest prediction', bodyY: 5.914, bodyH: 0.633, text: L_MED },
  ];
  plans.forEach((p) => {
    label(s, { x: p.numX, y: p.numY, w: 0.948, h: 0.572, text: p.num, size: 28, color: DARK, spacing: 1 });
    label(s, { x: 1.657, y: p.headY, w: p.headW, h: 0.218, text: p.head, size: 14, color: DARK });
    body(s, { x: 1.657, y: p.bodyY, w: 4.687, h: p.bodyH, text: p.text, color: INK, valign: 'middle' });
  });

  photoPill(s, pptx, 0.745, 1.391, 5.479, 1.76);
  photoRect(s, pptx, 6.905, 3.151, 5.701, 3.396, 0.222);
}

function slide10(pptx) {
  const s = pptx.addSlide();
  s.background = { color: DARK };
  navBar(s, pptx, true);

  photoPill(s, pptx, 8.922, 1.36, 3.684, 0.999);
  photoPill(s, pptx, 0.728, 2.581, 3.272, 0.999);
  // Stepped frame: wide band on top with a block hanging off the right.
  photoPath(s, pptx, 0.728, 4.083, 5.348, 2.44, [
    [0.263, 0], [5.086, 0], [5.348, 0.264, 5.223, 0.023, 5.348, 0.121], [5.348, 2.259],
    [5.166, 2.44, 5.348, 2.36, 5.242, 2.44], [3.791, 2.44],
    [3.608, 2.258, 3.69, 2.44, 3.608, 2.359], [3.608, 2.03],
    [3.233, 1.69, 3.386, 2.03, 3.233, 1.878], [0.265, 1.69],
    [0, 1.425, 0.119, 1.69, 0, 1.571], [0, 0.265],
    [0.263, 0, 0, 0.121, 0.114, 0],
  ]);
  photoRect(s, pptx, 6.667, 4.726, 5.939, 1.798, 0.245);

  title(s, {
    x: 0.682, y: 1.286, w: 12.085, h: 2.464, size: 72, spaceBefore: 10,
    lines: [{ text: 'Thank You For', color: PAPER }, { text: '             Your Attention!', color: PAPER }],
  });
  body(s, { x: 6.695, y: 3.75, w: 5.602, h: 0.625, text: L_LONG, color: PAPER });
  navButton(s, pptx, { text: 'END PRESENTATION', btnX: 0.739, circleX: 3.47, back: true });
}

/* ------------------------------------------------------------------- build */

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE', width: 13.333333, height: 7.5 });
pptx.layout = 'WIDE';
pptx.title = 'AGRIFY';

[slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10]
  .forEach((build) => { build(pptx); });

pptx.writeFile({ fileName: path.join(__dirname, '0bb3510a-b433-4192-a305-30b165618d23_grok_final.pptx') })
  .then((f) => { console.log('wrote', f); });
