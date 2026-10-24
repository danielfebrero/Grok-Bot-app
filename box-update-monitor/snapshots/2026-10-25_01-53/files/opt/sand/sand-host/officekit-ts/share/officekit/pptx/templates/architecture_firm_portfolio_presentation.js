/**
 * "GIGANT Architecture" — 25-slide deck rebuilt with pptxgenjs.
 * Raster photos / device photos in the original are re-drawn as plain shapes.
 *
 *   node 05bb0aca-479e-4d28-a063-1c6556bf2e47_grok_final.js
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const RED = 'C00000';
const DARK = '0D0D0D';
const BLACK = '000000';
const GRAY = '808080';
const WHITE = 'FFFFFF';
const INK = '262626';
const INK2 = '111111';
const GRAY40 = '404040';
const OVAL_BG = 'E8E8EA';
const PANEL_BG = 'E1E2E6';
const DOTLINE = 'BFBFBF';

const MAB = 'Montserrat Alternates Black';
const MSB = 'Montserrat SemiBold';
const MBL = 'Montserrat Black';
const MEL = 'Montserrat ExtraLight';
const LATO = 'Lato';

/* ----------------------------------------------------------- copy strings */

const LOREM1 =
  'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo ' +
  'libero vivamus porttitor dolor, conubia mollit. Sapien nam suspendisse, tincidunt eget ante tincidunt, ';
const LOREM2 = LOREM1 + 'eros in lacus nulla ac netus nibh aliquet, ';
const LOREM_LONG =
  LOREM1 +
  'eros in lacus nulla ac netus nibh aliquet, porttitor ligula justo libero vivamus porttitor dolor, ' +
  'conubia mollit. Sapien nam suspendisse, tincidunt eget ante tincidunt';
const NOTE = 'PLACEHOLDER';
const TIMELINE =
  'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo libero ' +
  'vivamus porttitor dolor, conubia mollit. Sapien nam suspendisse, tincidunt eget ante tincidunt, ' +
  'eros in auctor fringilla praesent at diam. ';
const TIMELINE_LONG =
  TIMELINE + 'mollit. Sapien nam suspendisse, tincidunt eget ante tincidunt, eros in auctor fringilla praesent at diam. ';
const TIMELINE_SHORT =
  'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo libero ' +
  'vivamus porttitor dolor, conubia mollit. Sapien nam suspendisse, tincidunt eget ante tincidunt, ' +
  'eros in auctor fringilla';
const CARD = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor';
const CARD_ITALIC = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, ';

/* ---------------------------------------------------------------- helpers */

/** Text box: top-anchored, matching the source deck's default body settings. */
function txt(slide, text, opts) {
  slide.addText(text, Object.assign({ valign: 'top', color: DARK, fontFace: LATO }, opts));
}

/** Solid rectangle. */
function rect(slide, x, y, w, h, color, extra) {
  slide.addShape('rect', Object.assign({ x, y, w, h, fill: { color } }, extra || {}));
}

/** Straight connector (w:0 = vertical, h:0 = horizontal). */
function connector(slide, x, y, w, h, color, extra) {
  slide.addShape('line', Object.assign({ x, y, w, h, line: Object.assign({ color }, extra || {}) }));
}

/** Red chip with centred white label (timeline markers). */
function chip(slide, x, y, w, h, text, fontSize) {
  slide.addText(text, {
    x, y, w, h, fill: { color: RED }, fontSize, fontFace: LATO,
    color: WHITE, align: 'center', valign: 'middle',
  });
}

/** Small outlined ">" marker used in the red bars. */
function chevron(slide, x, y, color) {
  slide.addShape('chevron', { x, y, w: 0.215, h: 0.286, fill: { type: 'none' }, line: { color, width: 1 } });
}

/** "GIGANT" wordmark, top-left of every slide. */
function wordmark(slide, color) {
  txt(slide, 'GIGANT', { x: 0.302, y: 0.121, w: 0.869, h: 0.286, fontSize: 11, bold: true, fontFace: MAB, color });
}

/** "Multipurpose / Presentation / New style" trio. */
function footer(slide, x, y, color) {
  const items = [['Multipurpose', 0, 1.512], ['Presentation', 2.012, 1.389], ['New style', 3.696, 1.389]];
  items.forEach(([label, dx, w]) =>
    txt(slide, label, { x: x + dx, y, w, h: 0.252, fontSize: 9, fontFace: MAB, color })
  );
}

/** Two grey lorem paragraphs (left aligned + justified) stacked 0.832" apart. */
function bodyPair(slide, x, y, color) {
  const c = color || GRAY;
  txt(slide, LOREM1, { x, y, w: 2.811, h: 0.707, fontSize: 9, color: c });
  txt(slide, LOREM2, { x, y: y + 0.832, w: 3.521, h: 0.707, fontSize: 9, color: c, align: 'justify' });
}

/**
 * The recurring editorial block: grey eyebrow, two 32pt headline lines and
 * (optionally) the two lorem paragraphs.
 *  body: 'both' (default) | 'first' | 'none'
 */
function titleBlock(slide, o) {
  const x = o.x, y = o.y;
  txt(slide, 'New Concept', { x, y, w: 4.22, h: 0.269, fontSize: 10, bold: true, fontFace: MAB, color: GRAY });
  txt(slide, o.line1, { x, y: y + 0.624, w: o.w1 || 4.22, h: 0.64, fontSize: 32, bold: true, fontFace: MAB });
  txt(slide, o.line2, { x, y: y + 1.065, w: 5.189, h: 0.64, fontSize: 32, bold: true, fontFace: MAB });
  if (o.body === 'first') {
    txt(slide, LOREM1, { x, y: y + 2.021, w: 2.811, h: 0.707, fontSize: 9, color: GRAY });
  } else if (o.body !== 'none') {
    bodyPair(slide, x, y + 2.021);
  }
}

/** Three "- Lorem / - ipsum dolor / - lacus nulla" lines. */
function dashList(slide, x, y, color) {
  const rows = [['- Lorem', 0, 1.499], ['- ipsum dolor', 0.229, 1.499], ['- lacus nulla', 0.476, 1.914]];
  rows.forEach(([label, dy, w]) =>
    txt(slide, label, { x, y: y + dy, w, h: 0.303, fontSize: 12, fontFace: MSB, color })
  );
}

/** The little 3-line caption used inside coloured panels. */
function note(slide, x, y, color) {
  txt(slide, NOTE, { x, y, w: 1.643, h: 0.555, fontSize: 9, color });
}

/** White tick drawn as a closed polygon (freeform in the source deck). */
function checkMark(slide, x, y, w, h) {
  slide.addShape('custGeom', {
    x, y, w, h, fill: { color: WHITE },
    points: [
      { x: 0, y: 0.50 * h }, { x: 0.14 * w, y: 0.37 * h }, { x: 0.38 * w, y: 0.66 * h },
      { x: 0.86 * w, y: 0 }, { x: w, y: 0.13 * h }, { x: 0.40 * w, y: h }, { close: true },
    ],
  });
}

/** Heading + caption pair ("Mission 01", "Vision 02", "Option 03", ...). */
function miniBlock(slide, x, y, heading, color) {
  txt(slide, heading, {
    x, y, w: 1.102, h: 0.373, fontSize: 12, fontFace: MSB, color, lineSpacingMultiple: 1.5,
  });
  note(slide, x, y + 0.354, color);
}

/* ---------------------------------------------- device mock-up placeholders
 * The original deck uses transparent-PNG product photos; they are recreated
 * from primitive shapes so no binary assets are needed.                     */

/**
 * Handset: 2.471 x 4.954, see-through screen. The 0.163" bezel is two
 * concentric strokes — brushed metal outside, matte black inside.
 */
function phoneMockup(slide, x, y) {
  const w = 2.471, h = 4.954;
  const band = (inset, thick, color, radius) =>
    slide.addShape('roundRect', {
      x: x + inset + thick / 2, y: y + inset + thick / 2,
      w: w - 2 * (inset + thick / 2), h: h - 2 * (inset + thick / 2),
      rectRadius: radius, fill: { type: 'none' }, line: { color, width: thick * 72 },
    });
  band(0, 0.062, 'BEBFC1', 0.36);
  band(0.062, 0.101, '181818', 0.30);
  // Notch: hangs off the top bezel, rounded at the bottom.
  slide.addShape('roundRect', {
    x: x + 0.718, y: y + 0.09, w: 1.05, h: 0.238, rectRadius: 0.075, fill: { color: '181818' },
  });
}

/** Tablet: dark rounded body, white screen. */
function tabletMockup(slide, x, y, w, h) {
  slide.addShape('roundRect', { x, y, w, h, rectRadius: 0.26, fill: { color: '1C1C1E' } });
  slide.addShape('roundRect', {
    x: x + 0.2, y: y + 0.209, w: w - 0.415, h: h - 0.485, rectRadius: 0.04, fill: { color: WHITE },
  });
}

/** Open laptop: dark lid, white screen, silver deck sticking out on both sides. */
function laptopMockup(slide, x, y, w, h) {
  slide.addShape('roundRect', { x, y, w, h, rectRadius: 0.03, fill: { color: '141414' } });
  rect(slide, x + 0.227, y + 0.292, w - 0.462, h - 0.591, WHITE);
  slide.addShape('roundRect', {
    x: x - 0.875, y: y + h, w: w + 1.734, h: 0.321, rectRadius: 0.35, fill: { color: '9C9EA1' },
  });
  slide.addShape('roundRect', {
    x: x - 0.875, y: y + h, w: w + 1.734, h: 0.24, rectRadius: 0.12, fill: { color: 'EFEFEF' },
  });
  slide.addShape('roundRect', {
    x: x + w / 2 - 0.63, y: y + h, w: 1.26, h: 0.1, rectRadius: 0.05, fill: { color: 'CFD0D2' },
  });
}

/** All-in-one desktop: screen, chin, neck and elliptical foot. */
function monitorMockup(slide, x, y, w, h) {
  slide.addShape('roundRect', { x, y, w, h, rectRadius: 0.02, fill: { color: '111111' } });
  rect(slide, x + 0.196, y + 0.22, w - 0.393, h - 0.433, WHITE);
  rect(slide, x, y + h, w, 0.527, 'C4C4C4');                       // chin
  slide.addShape('ellipse', {                                      // soft floor shadow
    x: x + 0.55, y: y + h + 1.04, w: 4.51, h: 0.38, fill: { color: 'E6E6E6' },
  });
  const nx = x + w / 2, ny = y + h + 0.527;                        // flared neck
  slide.addShape('custGeom', {
    x: nx - 0.95, y: ny, w: 1.9, h: 0.62, fill: { color: 'B0B0B0' },
    points: [
      { x: 0.365, y: 0 }, { x: 1.535, y: 0 }, { x: 1.9, y: 0.62 }, { x: 0, y: 0.62 }, { close: true },
    ],
  });
}

/* ----------------------------------------------------------------- slides */

const slides = [];

// 1 — cover
slides.push((s) => {
  wordmark(s, DARK);
  rect(s, 7.476, 4.476, 5.857, 1.786, RED);
  note(s, 8.037, 5.091, WHITE);
  connector(s, 9.948, 4.895, 0, 0.948, WHITE);
  txt(s, '2024', { x: 10.503, y: 5.226, w: 0.869, h: 0.37, fontSize: 16, bold: true, fontFace: MAB, color: WHITE });
  chevron(s, 12.5, 5.226, WHITE);
  titleBlock(s, { x: 0.954, y: 2.38, line1: 'Gigant', line2: 'Architecture', body: 'first' });
  txt(s, 'Read More', { x: 0.954, y: 5.327, w: 4.22, h: 0.303, fontSize: 12, bold: true, fontFace: MAB });
  footer(s, 7.242, 6.745, DARK);
});

// 2 — welcome
slides.push((s) => {
  rect(s, 0, 6.052, 5.698, 1.448, RED);
  wordmark(s, DARK);
  titleBlock(s, { x: 6.667, y: 2.036, line1: 'Welcome to', line2: 'Gigant Architecture' });
  footer(s, 7.242, 6.745, DARK);
  dashList(s, 0.517, 6.352, WHITE);
  note(s, 3.385, 6.424, WHITE);
  connector(s, 2.885, 6.296, 0, 0.948, WHITE);
});

// 3 — landscape architecture
slides.push((s) => {
  rect(s, 0, 6.188, 6.25, 1.312, RED);
  footer(s, 0.377, 6.745, WHITE);
  titleBlock(s, { x: 1.146, y: 2.182, line1: 'Landscape', line2: 'Architecture' });
  wordmark(s, DARK);
  chevron(s, 5.593, 6.701, WHITE);
});

// 4 — minimal architecture
slides.push((s) => {
  wordmark(s, DARK);
  titleBlock(s, { x: 7.708, y: 1.338, line1: 'Minimal', line2: 'Architecture' });
  rect(s, 1.246, 4.698, 2.661, 2.841, RED);
  footer(s, 7.242, 6.745, DARK);
  note(s, 1.704, 5.362, WHITE);
  connector(s, 1.812, 6.159, 1.271, 0, WHITE);
  dashList(s, 1.7, 6.392, WHITE);
});

// 5 — history
slides.push((s) => {
  rect(s, 6.667, 6.145, 6.667, 1.355, RED);
  wordmark(s, DARK);
  titleBlock(s, { x: 7.556, y: 1.355, line1: 'History', line2: 'Gigant Studio' });
  footer(s, 7.242, 6.745, WHITE);
  chevron(s, 12.622, 6.701, WHITE);
});

// 6 — how we work
slides.push((s) => {
  wordmark(s, DARK);
  titleBlock(s, { x: 6.667, y: 2.175, line1: 'How we work', line2: 'With Gigant' });
  footer(s, 7.242, 6.745, DARK);
  rect(s, 0, 4.278, 2.663, 3.222, RED);
  dashList(s, 0.46, 4.71, WHITE);
});

// 7 — quality material
slides.push((s) => {
  wordmark(s, DARK);
  titleBlock(s, { x: 7.573, y: 2.144, line1: 'We use Quality', line2: 'Material' });
  footer(s, 8.148, 6.714, DARK);
  rect(s, 6.952, 0, 6.381, 0.792, RED);
});

// 8 — mission
slides.push((s) => {
  rect(s, 6.685, 1.615, 6.649, 5.885, RED);
  [0.741, 2.497].forEach((x) =>
    [3.564, 5.082].forEach((y) => miniBlock(s, x, y, 'Mission 01', DARK))
  );
  txt(s, LOREM_LONG, { x: 11.27, y: 2.451, w: 1.651, h: 2.373, fontSize: 9, color: WHITE, align: 'justify' });
  wordmark(s, DARK);
  titleBlock(s, { x: 0.741, y: 1.509, line1: 'Mission', line2: 'Gigant Studio', body: 'none' });
  footer(s, 7.242, 6.745, WHITE);
});

// 9 — vision
slides.push((s) => {
  ['Vision 01', 'Vision 02', 'Vision 03'].forEach((label, i) =>
    miniBlock(s, [7.287, 9.193, 10.836][i], 4.005, label, DARK)
  );
  rect(s, 3.488, 5.35, 3.179, 2.15, RED);
  txt(s, LOREM2, { x: 3.831, y: 5.898, w: 2.129, h: 1.161, fontSize: 9, color: WHITE, align: 'justify' });
  wordmark(s, DARK);
  titleBlock(s, { x: 7.29, y: 1.631, line1: 'Vision', line2: 'Gigant Studio', body: 'none' });
  txt(s, LOREM2, { x: 7.287, y: 5.393, w: 3.521, h: 0.707, fontSize: 9, color: GRAY, align: 'justify' });
  footer(s, 7.242, 6.745, DARK);
});

// 10 — our services (options on red)
slides.push((s) => {
  rect(s, 5.054, 4.533, 7.696, 1.854, RED);
  wordmark(s, DARK);
  ['Option 01', 'Option 02', 'Option 03'].forEach((label, i) =>
    miniBlock(s, [5.9, 7.806, 9.449][i], 4.974, label, WHITE)
  );
  titleBlock(s, { x: 5.903, y: 2.342, line1: 'Our Services', line2: 'Gigant Studio', body: 'none' });
  footer(s, 7.242, 6.745, DARK);
});

// 11 — our services (red column)
slides.push((s) => {
  rect(s, 8.178, 0, 5.155, 6.219, RED);
  wordmark(s, DARK);
  titleBlock(s, { x: 1.651, y: 2.013, line1: 'Our Services', line2: 'Gigant Studio' });
  footer(s, 7.242, 6.745, DARK);
});

// 12 — numbered service grid
slides.push((s) => {
  const cells = [
    { n: '1', ox: 0.914, oy: 2.743, nx: 0.673, ny: 2.563, tx: 1.856, ty: 2.804 },
    { n: '3', ox: 5.316, oy: 2.743, nx: 5.089, ny: 2.541, tx: 6.271, ty: 2.782 },
    { n: '2', ox: 0.914, oy: 4.92, nx: 0.673, ny: 4.696, tx: 1.856, ty: 4.937 },
    { n: '4', ox: 5.316, oy: 4.92, nx: 5.089, ny: 4.696, tx: 6.271, ty: 4.937 },
  ];
  cells.forEach((c) => {
    s.addShape('ellipse', { x: c.ox, y: c.oy, w: 0.699, h: 0.699, fill: { color: OVAL_BG } });
    s.addText(c.n, {
      x: c.nx, y: c.ny, w: 1.182, h: 1.145, fontSize: 32, fontFace: MBL,
      color: RED, align: 'center', valign: 'middle',
    });
    txt(s, CARD, { x: c.tx, y: c.ty, w: 2.793, h: 0.499, fontSize: 10, color: INK, lineSpacingMultiple: 1.25 });
  });
  rect(s, 10.148, 2.563, 2.477, 3.306, RED);
  s.addText('OUR SERVICE', {
    x: 1.117, y: 1.172, w: 3.83, h: 1.237, fontSize: 32, bold: true, fontFace: MAB,
    color: GRAY40, valign: 'middle',
  });
  txt(s, 'New Concept', { x: 1.171, y: 1.06, w: 4.22, h: 0.269, fontSize: 10, bold: true, fontFace: MAB, color: GRAY });
  txt(s, CARD_ITALIC, {
    x: 10.148, y: 1.526, w: 2.477, h: 0.539, fontSize: 11, italic: true, color: INK, lineSpacingMultiple: 1.25,
  });
  txt(s, [
    { text: 'This is our New Clean and ', options: { breakLine: true } },
    { text: 'Creative style ', options: { breakLine: true } },
    { text: 'design' },
  ], { x: 10.508, y: 3.489, w: 1.903, h: 1.616, fontSize: 18, color: WHITE, lineSpacingMultiple: 1.25 });
  wordmark(s, DARK);
  footer(s, 7.242, 6.745, DARK);
});

// 13 — vertical timeline (top half)
slides.push((s) => {
  s.addText([
    { text: 'VERTICAL ', options: { breakLine: true } },
    { text: 'TIMELINE' },
  ], { x: 0.731, y: 3.423, w: 7.5, h: 0.806, fontSize: 32, bold: true, fontFace: MAB, color: RED, valign: 'middle' });
  connector(s, 7.913, 1.183, 0, 6.317, DOTLINE, { width: 2, dashType: 'sysDot' });

  txt(s, 'Fashion ', { x: 8.876, y: 0.783, w: 2.569, h: 0.303, fontSize: 12, bold: true });
  txt(s, TIMELINE_LONG, { x: 8.876, y: 1.356, w: 2.577, h: 1.488, fontSize: 8, color: INK2, lineSpacingMultiple: 1.5 });
  chip(s, 7.103, 0.887, 1.62, 0.296, 'STARTING POINT', 12);

  txt(s, 'Design and Style', { x: 4.489, y: 3.182, w: 2.569, h: 0.303, fontSize: 12, bold: true, align: 'right' });
  txt(s, TIMELINE, {
    x: 4.481, y: 3.471, w: 2.577, h: 1.111, fontSize: 8, color: INK2, align: 'right', lineSpacingMultiple: 1.5,
  });
  chip(s, 7.317, 3.22, 1.191, 0.405, '12 March 2020', 10);

  txt(s, 'Fashion and Design', { x: 8.876, y: 5.385, w: 2.569, h: 0.303, fontSize: 12, bold: true });
  txt(s, TIMELINE, { x: 8.868, y: 5.673, w: 2.577, h: 1.111, fontSize: 8, color: INK2, lineSpacingMultiple: 1.5 });
  chip(s, 7.317, 5.423, 1.191, 0.405, '12 March 2021', 10);
  wordmark(s, DARK);
});

// 14 — vertical timeline (continued)
slides.push((s) => {
  connector(s, 7.865, 0, 0, 5.327, DOTLINE, { width: 2, dashType: 'sysDot' });

  txt(s, 'NULLA NETUS NIBH ALIQUET PORTTITOR', { x: 8.821, y: 0.922, w: 2.569, h: 0.505, fontSize: 12 });
  txt(s, TIMELINE_SHORT, { x: 8.821, y: 1.495, w: 2.577, h: 1.111, fontSize: 8, lineSpacingMultiple: 1.5 });
  chip(s, 7.27, 0.922, 1.191, 0.405, '12 March 2021', 10);

  txt(s, 'NULLA NETUS NIBH', { x: 8.829, y: 5.315, w: 2.569, h: 0.303, fontSize: 12 });
  txt(s, TIMELINE, { x: 8.821, y: 5.603, w: 2.577, h: 1.111, fontSize: 8, lineSpacingMultiple: 1.5 });
  rect(s, 7.037, 5.327, 1.657, 0.405, RED);
  checkMark(s, 7.74, 5.398, 0.25, 0.264);

  footer(s, 1.85, 1.075, DARK);
  rect(s, 0, 1.704, 1.817, 5.796, RED);
  wordmark(s, DARK);
});

// 15 — our founder
slides.push((s) => {
  rect(s, 0, 4.944, 3.278, 2.556, RED);
  rect(s, 3.278, 6.063, 3.222, 1.438, PANEL_BG);
  dashList(s, 0.462, 6.098, WHITE);
  wordmark(s, DARK);
  titleBlock(s, { x: 7.138, y: 1.438, line1: 'Our Founder', line2: 'John Doe' });
  footer(s, 7.242, 6.745, DARK);
  note(s, 4.067, 6.429, DARK);
});

// 16 — marketing / quote
slides.push((s) => {
  rect(s, 2.618, 5.964, 4.049, 1.536, RED);
  txt(s, 'All art is a struggle to be, in a particular sort of way, virtuous.', {
    x: 2.995, y: 6.35, w: 2.244, h: 0.808, fontSize: 14, bold: true, fontFace: MEL, color: WHITE,
  });
  wordmark(s, DARK);
  titleBlock(s, { x: 6.888, y: 1.682, line1: 'Marketing Gigant', line2: 'Natalia Aderson', w1: 5.189 });
  chevron(s, 6.108, 6.582, WHITE);
  rect(s, 0, 4.659, 2.62, 2.841, PANEL_BG);
  note(s, 0.418, 5.323, BLACK);
  connector(s, 0.527, 6.12, 1.271, 0, DARK);
  dashList(s, 0.415, 6.353, BLACK);
  footer(s, 7.242, 6.745, DARK);
});

// 17 — two portraits
slides.push((s) => {
  wordmark(s, DARK);
  footer(s, 7.242, 6.745, DARK);
  titleBlock(s, { x: 8.315, y: 1.953, line1: 'Achitecture', line2: 'Gigant Studio', w1: 3.237, body: 'none' });
  txt(s, 'Natalia And Gissela', { x: 8.315, y: 3.856, w: 3.075, h: 0.37, fontSize: 16, bold: true, fontFace: MAB });
  bodyPair(s, 8.315, 4.225);
});

// 18 — break time
slides.push((s) => {
  s.addShape('rect', { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: BLACK, transparency: 54 } });
  s.addText('BREAK TIME', {
    x: 4.968, y: 2.322, w: 3.397, h: 1.681, fontSize: 32, bold: true, fontFace: MAB,
    color: WHITE, align: 'center', valign: 'middle',
  });
  txt(s, LOREM2, { x: 3.28, y: 4.003, w: 6.773, h: 0.404, fontSize: 9, color: WHITE, align: 'center' });
  wordmark(s, WHITE);
  footer(s, 7.242, 6.745, WHITE);
  chevron(s, 12.622, 6.701, WHITE);
});

// 19 — phone mockup
slides.push((s) => {
  rect(s, 0, 0.577, 5.113, 6.923, RED);
  phoneMockup(s, 1.092, 1.39);
  phoneMockup(s, 3.863, 1.39);
  wordmark(s, DARK);
  titleBlock(s, { x: 7.138, y: 1.438, line1: 'Phone', line2: 'Mockup' });
  footer(s, 7.242, 6.745, DARK);
});

// 20 — ipad mockup
slides.push((s) => {
  tabletMockup(s, 1.647, 0.892, 4.125, 5.789);
  rect(s, 6.667, 6.062, 6.667, 1.438, RED);
  wordmark(s, DARK);
  titleBlock(s, { x: 7.138, y: 1.438, line1: 'Ipad', line2: 'Mockup' });
  footer(s, 7.242, 6.745, WHITE);
});

// 21 — laptop mockup
slides.push((s) => {
  laptopMockup(s, -0.721, 1.199, 7.469, 4.872);
  wordmark(s, DARK);
  titleBlock(s, { x: 7.492, y: 1.562, line1: 'Laptop', line2: 'Mockup' });
  rect(s, 0.667, 4.944, 1.961, 2.556, RED);
  dashList(s, 1.129, 6.098, WHITE);
  footer(s, 7.242, 6.745, DARK);
});

// 22 — desktop mockup
slides.push((s) => {
  monitorMockup(s, 6.491, 1.348, 5.607, 3.435);
  wordmark(s, DARK);
  titleBlock(s, { x: 2.158, y: 1.552, line1: 'Dekstop', line2: 'Mockup' });
  footer(s, 7.242, 6.745, DARK);
});

// 23 — follow and contact us
slides.push((s) => {
  rect(s, 6.667, 5.958, 6.667, 1.542, RED);
  titleBlock(s, { x: 1.646, y: 1.687, line1: 'Follow and', line2: 'Contact us' });
  wordmark(s, DARK);
  footer(s, 7.242, 6.745, WHITE);
});

// 24 — contact & social media
slides.push((s) => {
  rect(s, 0, 1.394, 7.519, 6.106, RED);
  txt(s, 'CONTACT & SOCIAL MEDIA', {
    x: 0.847, y: 2.404, w: 3.512, h: 1.178, fontSize: 32, bold: true, fontFace: MAB, color: WHITE,
  });
  const y0Icon = 4.006;
  [1.246, 1.811, 2.374].forEach((x, i) => {
    s.addShape('roundRect', {
      x, y: 3.905, w: 0.387, h: 0.387, rectRadius: 0.06,
      fill: { type: 'none' }, line: { color: WHITE, width: 1 },
    });
    if (i === 0) {
      // Instagram: rounded square outline + lens circle + corner dot.
      s.addShape('roundRect', {
        x: x + 0.101, y: y0Icon, w: 0.185, h: 0.185, rectRadius: 0.05,
        fill: { type: 'none' }, line: { color: WHITE, width: 1 },
      });
      s.addShape('ellipse', {
        x: x + 0.145, y: y0Icon + 0.044, w: 0.097, h: 0.097,
        fill: { type: 'none' }, line: { color: WHITE, width: 1 },
      });
      s.addShape('ellipse', { x: x + 0.246, y: y0Icon + 0.021, w: 0.024, h: 0.024, fill: { color: WHITE } });
    } else {
      s.addText(['f', 'g+'][i - 1], {
        x, y: 3.905, w: 0.387, h: 0.387, fontSize: 13, bold: true, fontFace: LATO,
        color: WHITE, align: 'center', valign: 'middle',
      });
    }
  });
  txt(s, 'Address :', { x: 3.896, y: 4.615, w: 1.569, h: 0.505, fontSize: 24, color: WHITE });
  [['New York, 5 Ave 321', 5.275, 2.152], ['+01 10345 678', 5.691, 2.152], ['www.gigant.com', 6.101, 2.324]]
    .forEach(([label, y, w]) =>
      txt(s, label, { x: 4.873, y, w, h: 0.415, fontSize: 14, color: WHITE, lineSpacingMultiple: 1.5 })
    );
  wordmark(s, DARK);
});

// 25 — thank you
slides.push((s) => {
  wordmark(s, DARK);
  connector(s, 9.948, 4.895, 0, 0.948, WHITE);
  txt(s, [
    { text: 'THANK YOU', options: { breakLine: true } },
    { text: 'FOR WATCHING' },
  ], { x: 4.557, y: 2.466, w: 4.22, h: 1.582, fontSize: 44, bold: true, fontFace: MAB, align: 'center' });
  footer(s, 4.288, 0.155, DARK);
  rect(s, 7.476, 5.378, 5.857, 1.786, RED);
  note(s, 8.037, 5.993, WHITE);
  connector(s, 9.948, 5.797, 0, 0.948, WHITE);
  txt(s, '2024', { x: 10.503, y: 6.127, w: 0.869, h: 0.37, fontSize: 16, bold: true, fontFace: MAB, color: WHITE });
  chevron(s, 12.5, 6.127, WHITE);
});

/* ------------------------------------------------------------------ build */

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE_16x9', width: 13.333, height: 7.5 });
pptx.layout = 'WIDE_16x9';
pptx.title = 'Gigant Architecture';

slides.forEach((build) => {
  const slide = pptx.addSlide();
  slide.background = { color: WHITE };
  build(slide);
});

pptx.writeFile({
  fileName: path.join(__dirname, '05bb0aca-479e-4d28-a063-1c6556bf2e47_grok_final.pptx'),
});
