/*
 * "Automobi" - Car Repair Presentation Template (20 slides, 13.333" x 7.5").
 *
 * Standalone pptxgenjs re-creation of the reference deck.
 * Photographs in the original are replaced by flat grey "[image]" placeholder
 * rectangles at the same position/size; vector icons are replaced by unicode
 * glyphs drawn with native shapes.
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const C = {
  ink: '404040', // tx1 lumMod 75% - body/headline grey
  white: 'FFFFFF',
  accent1: '220901', // near-black brown
  accent2: '621708',
  accent3: '941B0C',
  accent4: 'BC3908', // signature orange
  accent5: 'F6AA1C',
  star: 'FFC000',
  card: 'FFFFFF',
  ph: 'CCCCCC', // image placeholder grey
  phInk: 'A8A8A8', // image placeholder caption
  rule: 'F2F2F2', // bg1 lumMod 95% - progress-bar track
};

const MAJOR = 'Lexend SemiBold'; // theme major font (headings)
const MINOR = 'Plus Jakarta Sans'; // theme minor font (body)

// Left-to-right scrim used on the title / thank-you / banner slides.
const SCRIM_FROM = '621708';
const SCRIM_TO = '8D6F66';

// Soft "floating card" drop shadow used by every white panel in the deck.
// Returns a fresh object each call: pptxgenjs rewrites the props in place.
const cardShadow = () => ({ type: 'outer', color: '000000', opacity: 0.07, blur: 36, offset: 12, angle: 90 });

/* ---------------------------------------------------------------- helpers */

// Heading: theme major font, tight leading, mixed-colour runs.
function heading(slide, x, y, w, h, runs, size, opts = {}) {
  slide.addText(
    runs.map((r) => ({ text: r[0], options: { color: r[1] || C.ink } })),
    Object.assign(
      { x, y, w, h, fontFace: MAJOR, fontSize: size, color: C.ink, valign: 'top', margin: [3.5, 7, 3.5, 7] },
      opts
    )
  );
}

// Body copy: minor font, 12pt, 1.5 line spacing.
function body(slide, x, y, w, h, text, opts = {}) {
  slide.addText(
    text,
    Object.assign(
      {
        x, y, w, h,
        fontFace: MINOR, fontSize: 12, color: C.ink,
        lineSpacingMultiple: 1.5, valign: 'top', margin: [3.5, 7, 3.5, 7],
      },
      opts
    )
  );
}

// Plain one-liner in the minor/major font (labels, names, list items).
function label(slide, x, y, w, h, text, opts = {}) {
  slide.addText(
    text,
    Object.assign(
      { x, y, w, h, fontFace: MAJOR, fontSize: 14, color: C.ink, valign: 'top', margin: [3.5, 7, 3.5, 7] },
      opts
    )
  );
}

// White rounded panel with the deck's signature soft shadow.
function card(slide, x, y, w, h, radius) {
  slide.addShape('roundRect', { x, y, w, h, rectRadius: radius, fill: { color: C.card }, line: { type: 'none' }, shadow: cardShadow() });
}

// Stand-in for a photograph: flat grey block labelled "[image]".
function imagePlaceholder(slide, x, y, w, h, radius = 0.15) {
  if (radius > 0) {
    slide.addShape('roundRect', { x, y, w, h, rectRadius: Math.min(radius, w / 2, h / 2), fill: { color: C.ph }, line: { type: 'none' } });
  } else {
    slide.addShape('rect', { x, y, w, h, fill: { color: C.ph }, line: { type: 'none' } });
  }
  if (w > 0.9 && h > 0.5) {
    slide.addText('[image]', {
      x, y: y + h / 2 - 0.18, w, h: 0.36,
      align: 'center', valign: 'middle', fontFace: MINOR, fontSize: 12, color: C.phInk,
    });
  }
}

function mixHex(from, to, t) {
  const chan = (i) => Math.round(parseInt(from.substr(i, 2), 16) * (1 - t) + parseInt(to.substr(i, 2), 16) * t);
  return [chan(0), chan(2), chan(4)].map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
}

// pptxgenjs has no gradient fill, so the horizontal scrim is painted as a run
// of narrow solid strips interpolating between the two stop colours. Rounded
// variants get a solid end-cap at each side so the corners stay round.
function gradientScrim(slide, x, y, w, h, radius = 0, steps = 48) {
  if (radius > 0) {
    slide.addShape('roundRect', { x, y, w: radius * 2, h, rectRadius: radius, fill: { color: SCRIM_FROM }, line: { type: 'none' } });
    slide.addShape('roundRect', { x: x + w - radius * 2, y, w: radius * 2, h, rectRadius: radius, fill: { color: SCRIM_TO }, line: { type: 'none' } });
  }
  const stripW = (w - radius * 2) / steps;
  for (let i = 0; i < steps; i++) {
    slide.addShape('rect', {
      x: x + radius + i * stripW, y, w: stripW + 0.012, h,
      fill: { color: mixHex(SCRIM_FROM, SCRIM_TO, i / (steps - 1)) }, line: { type: 'none' },
    });
  }
}

// Filled disc + centred glyph - the deck's round icon buttons.
function iconDisc(slide, x, y, d, glyph, fill = C.accent4, glyphColor = C.white) {
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color: fill }, line: { type: 'none' } });
  slide.addText(glyph, {
    x, y, w: d, h: d, align: 'center', valign: 'middle',
    fontFace: 'DejaVu Sans', fontSize: Math.max(5, Math.round(d * 44)), color: glyphColor,
  });
}

// Bare (un-circled) accent glyph, used for the service icons.
function iconGlyph(slide, x, y, d, glyph, color = C.accent4) {
  slide.addText(glyph, {
    x, y, w: d, h: d, align: 'center', valign: 'middle',
    fontFace: 'DejaVu Sans', fontSize: Math.round(d * 62), color,
  });
}

// The tilted orange lozenge ("Explore Now!" / "Order Now!" / "Job Position").
function pill(slide, x, y, text) {
  slide.addShape('roundRect', {
    x, y, w: 1.069, h: 0.262, rectRadius: 0.131, rotate: -7.43,
    fill: { color: C.accent4 }, line: { type: 'none' },
  });
  slide.addText(text, {
    x: x - 0.062, y: y + 0.022, w: 1.193, h: 0.219, rotate: -7,
    align: 'center', valign: 'middle', fontFace: MINOR, fontSize: 7, color: C.white,
  });
}

// "+8590" big number over a small caption - repeated all over the deck.
function stat(slide, x, y, value, caption, opts = {}) {
  const size = opts.size || 24;
  const align = opts.align || 'left';
  slide.addText(value, {
    x, y, w: 2.927, h: 0.505,
    fontFace: MAJOR, fontSize: size, color: opts.color || C.ink, align,
    valign: 'top', margin: [3.5, 7, 3.5, 7],
  });
  body(slide, x, y + (opts.gap || size / 56), opts.capW || 2.698, opts.capH || 0.364, caption, { color: opts.capColor || opts.color || C.ink, align });
}

// "Discover The Automobi" heading + supporting paragraph.
function discover(slide, x, y, w, text, title = 'Discover The Automobi') {
  label(slide, x, y, 3.84, 0.337, title);
  body(slide, x, y + 0.337, w, 0.667, text);
}

// Big decorative opening quote mark.
function quoteMark(slide, x, y) {
  slide.addText('\u201C', {
    x: x - 0.09, y: y - 0.16, w: 0.5, h: 0.42,
    fontFace: 'DejaVu Sans', fontSize: 28, bold: true, color: C.accent4, valign: 'top', margin: 0,
  });
}

// Page number (bottom-left) + circular "next" button (bottom-right) from the master.
function masterChrome(slide, n) {
  slide.addText(String(n), {
    x: 0.147, y: 6.951, w: 0.598, h: 0.399,
    align: 'center', valign: 'middle', fontFace: MINOR, fontSize: 10, color: C.ink,
  });
  iconDisc(slide, 12.786, 7.057, 0.202, '\u2192');
}

/* ------------------------------------------------- repeated copy fragments */

const LOREM_LONG =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
  'Fusce posuere, magna sed pulvinar ultricies, purus';
const LOREM_MED = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit.';
const MAECENAS = 'Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar';
const MAECENAS_LECTUS = 'Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus';
const CAPTION = 'Maecenas porttitor congue';
const TAGLINE = 'Don\u2019t Wait for a Breakdown Save Big on Preventative Repairs!';

const WRENCH = '\u2692';
const HAMMER = '\u2699';
const BOLT = '\u26A1';
const ARROW = '\u2192';

/* ------------------------------------------------------------ slide 1 - 20 */

function slide01(pptx) {
  const s = pptx.addSlide();
  gradientScrim(s, 0, 0, 13.333, 7.5); // hero photo is fully covered by the scrim
  heading(s, 0.903, 2.628, 9.634, 1.717, [['Automobi', C.white]], 96);
  pill(s, 6.233, 3.84, 'Explore Now!');
  body(s, 0.884, 4.078, 4.523, 0.303, 'Car Repair Presentation Template', { color: C.white, lineSpacingMultiple: 1 });
  heading(s, 9.5, 5.502, 3.426, 0.808, [[TAGLINE, C.white]], 14, { align: 'right' });

  s.addShape('roundRect', { x: 11.44, y: 6.725, w: 1.008, h: 0.262, rectRadius: 0.131, fill: { type: 'none' }, line: { color: C.white, width: 0.75 } });
  s.addText('Next Slide', { x: 11.444, y: 6.738, w: 1.0, h: 0.236, align: 'center', valign: 'middle', fontFace: MINOR, fontSize: 8, color: C.white });
  s.addShape('ellipse', { x: 12.555, y: 6.725, w: 0.262, h: 0.262, fill: { type: 'none' }, line: { color: C.white, width: 0.75 } });
  s.addText(ARROW, { x: 12.555, y: 6.725, w: 0.262, h: 0.262, align: 'center', valign: 'middle', fontFace: 'DejaVu Sans', fontSize: 8, color: C.white });
}

function slide02(pptx) {
  const s = pptx.addSlide();
  imagePlaceholder(s, 7.852, 1.063, 4.35, 5.18, 0.16);
  heading(s, 0.958, 1.772, 5.708, 2.827, [['Your ', C.ink], ['Trusted', C.accent4], [' Partner in Car Care', C.ink]], 54);
  body(s, 0.958, 4.758, 4.523, 0.97, LOREM_LONG);
  iconDisc(s, 7.923, 1.089, 0.598, ARROW);
  pill(s, 10.62, 6.112, 'Explore Now!');
  masterChrome(s, 2);
}

function slide03(pptx) {
  const s = pptx.addSlide();
  imagePlaceholder(s, 5.375, 0.774, 7.208, 2.976, 0.16);
  heading(s, 0.99, 0.83, 4.094, 2.827, [['Why', C.accent4], [' Choose Us?', C.ink]], 54);
  pill(s, 11.12, 3.619, 'Explore Now!');

  // Two feature cards + a plain text column.
  [
    { x: 0.99, title: 'Expert Technicians', icon: WRENCH },
    { x: 4.5, title: 'Comprehensive Warranty', icon: HAMMER },
  ].forEach((f) => {
    card(s, f.x, 4.283, 3.219, 2.387, 0.17);
    iconGlyph(s, f.x + 0.415, 4.643, 0.33, f.icon);
    label(s, f.x + 0.291, 5.163, 2.927, 0.337, f.title);
    body(s, f.x + 0.291, 5.485, 2.698, 0.97, MAECENAS);
  });
  discover(s, 8.252, 4.824, 3.84, LOREM_MED);

  // Three small stat chips floating over the photo.
  [['80%', 5.744], ['+678', 6.796], ['80%', 7.849]].forEach(([value, x]) => {
    card(s, x, 2.594, 0.923, 0.894, 0.151);
    body(s, x + 0.053, 2.728, 0.955, 0.202, 'Up to', { fontFace: MAJOR, fontSize: 6, lineSpacingMultiple: 1 });
    heading(s, x + 0.053, 2.855, 0.955, 0.37, [[value, C.accent4]], 16);
    body(s, x + 0.053, 3.151, 0.955, 0.202, 'Your Text Here', { fontFace: MAJOR, fontSize: 6, lineSpacingMultiple: 1 });
  });
  masterChrome(s, 3);
}

function slide04(pptx) {
  const s = pptx.addSlide();
  imagePlaceholder(s, 10.342, 0, 3.012, 3.75, 0.16);
  imagePlaceholder(s, 0.72, 0.853, 3.821, 5.794, 0.16);
  heading(s, 5.168, 1.112, 5.708, 1.717, [['Discount', C.accent4], [' Breakdown', C.ink]], 48);
  body(s, 5.168, 2.907, 4.62, 0.667, LOREM_MED);

  s.addShape('ellipse', { x: 3.962, y: 4.901, w: 1.16, h: 1.16, fill: { color: C.accent4 }, line: { type: 'none' } });
  heading(s, 3.845, 5.229, 1.394, 0.505, [['80%', C.white]], 24, { align: 'center', valign: 'middle', rotate: -4.07 });
  discover(s, 5.526, 4.991, 4.523, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ');
  iconDisc(s, 10.342, 3.056, 0.637, ARROW);
  masterChrome(s, 4);
}

function slide05(pptx) {
  const s = pptx.addSlide();
  imagePlaceholder(s, 0.877, 0.684, 5.241, 3.6, 0.16);
  heading(s, 7.037, 1.16, 5.789, 1.582, [['Comprehensive', C.accent4], [' Repair Services', C.ink]], 44);

  stat(s, 1.142, 4.693, '+8590', CAPTION);
  stat(s, 1.142, 5.767, '+8590', CAPTION);

  card(s, 4.711, 3.332, 4.653, 3.503, 0.17);
  imagePlaceholder(s, 4.946, 3.552, 4.174, 2.144, 0.16);
  heading(s, 5.203, 5.991, 2.927, 0.572, [['70%', C.ink]], 28);
  label(s, 6.421, 6.114, 2.58, 0.337, 'Repair Services Offered');
  pill(s, 7.711, 5.539, 'Explore Now!');

  label(s, 9.899, 4.624, 3.84, 0.337, 'Discover The Automobi');
  body(s, 9.899, 4.999, 2.716, 1.273, LOREM_MED);
  masterChrome(s, 5);
}

function slide06(pptx) {
  const s = pptx.addSlide();
  gradientScrim(s, 0.877, 0.684, 11.472, 3.6, 0.167); // banner photo is fully covered by the scrim
  heading(s, 1.561, 1.323, 4.298, 2.322, [['Keep Your Car Running Smoothly', C.white]], 44);
  card(s, 6.98, 2.283, 5.081, 1.363, 0.098);
  iconDisc(s, 7.401, 2.646, 0.637, ARROW);
  discover(s, 8.28, 2.463, 3.84, LOREM_SHORT);
  pill(s, 10.635, 3.458, 'Explore Now!');

  discover(s, 1.561, 5.128, 3.873, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ');
  stat(s, 6.849, 5.344, '+8590', CAPTION);
  stat(s, 9.776, 5.344, '+8590', CAPTION);
  masterChrome(s, 6);
}

function slide07(pptx) {
  const s = pptx.addSlide();
  heading(s, 2.149, 0.62, 9.036, 0.774, [['Specialized', C.accent4], [' Services', C.ink]], 40, { align: 'center' });
  body(s, 2.486, 1.535, 8.361, 0.667,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
    'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero',
    { align: 'center' });

  [
    { x: 1.175, icon: WRENCH },
    { x: 5.147, icon: HAMMER },
    { x: 9.119, icon: BOLT },
  ].forEach((col) => {
    imagePlaceholder(s, col.x, 2.819, 3.04, 3.653, 0.16);
    card(s, col.x + 2.352, 2.983, 0.502, 0.486, 0.082);
    iconGlyph(s, col.x + 2.451, 3.075, 0.302, col.icon);
    pill(s, col.x + 1.587, 6.341, 'Explore Now!');
  });
  masterChrome(s, 7);
}

function slide08(pptx) {
  const s = pptx.addSlide();
  imagePlaceholder(s, 6.363, 5.569, 6.098, 1.931, 0.16);
  imagePlaceholder(s, 0.877, 0.906, 5.342, 3.378, 0.16);
  heading(s, 7.215, 1.278, 4.828, 2.524, [['Precision ', C.ink], ['Care', C.accent4], [' for Every Vehicle', C.ink]], 48);
  body(s, 7.215, 4.029, 4.62, 0.667, LOREM_MED);
  iconDisc(s, 5.605, 3.689, 0.522, ARROW);
  pill(s, 10.688, 5.44, 'Explore Now!');

  // Numbered list, each row ending with a small dark bullet-arrow.
  [
    { n: '01', text: 'Advanced Diagnostics', y: 4.8, discX: 4.354 },
    { n: '02', text: 'Hybrid & Electric Vehicle Expertise', y: 5.524, discX: 5.354 },
    { n: '03', text: 'Performance Enhancements', y: 6.247, discX: 4.865 },
  ].forEach((row) => {
    heading(s, 1.138, row.y, 0.891, 0.505, [[row.n, C.ink]], 24);
    body(s, 1.828, row.y + 0.084, 3.84, 0.337, row.text, { fontSize: 14, lineSpacingMultiple: 1 });
    iconDisc(s, row.discX, row.y + 0.158, 0.188, ARROW, C.accent1);
  });
  masterChrome(s, 8);
}

function slide09(pptx) {
  const s = pptx.addSlide();
  // The two gallery frames here are empty picture placeholders in the original,
  // so nothing is drawn for them.
  heading(s, 6.313, 1.121, 4.489, 1.582, [['Before and After ', C.ink], ['Gallery', C.accent4]], 44);

  card(s, 0.96, 4.739, 5.081, 1.363, 0.098);
  iconDisc(s, 1.381, 5.102, 0.637, ARROW);
  discover(s, 2.26, 4.919, 3.84, LOREM_SHORT);
  pill(s, 4.615, 5.914, 'Explore Now!');

  stat(s, 8.292, 3.897, '+8590', MAECENAS, { color: C.accent4, capColor: C.ink, size: 18, gap: 0.404, capW: 3.844, capH: 0.667 });
  stat(s, 8.292, 5.236, '+8590', MAECENAS, { color: C.accent4, capColor: C.ink, size: 18, gap: 0.404, capW: 3.844, capH: 0.667 });
  masterChrome(s, 9);
}

function slide10(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 7.027, y: 2.917, w: 6.306, h: 2.617, fill: { color: C.accent4 }, line: { type: 'none' } });
  imagePlaceholder(s, 4.461, 2.594, 5.708, 3.25, 0);
  heading(s, 0.662, 1.312, 7.306, 2.322, [['Limited-Time ', C.ink], ['Offer', C.accent4]], 66);
  iconDisc(s, 0.763, 3.888, 0.469, ARROW);
  body(s, 0.662, 4.675, 3.712, 0.97, LOREM_MED);

  s.addShape('ellipse', { x: 4.594, y: 2.917, w: 1.16, h: 1.16, fill: { color: C.white }, line: { type: 'none' }, shadow: cardShadow() });
  heading(s, 4.477, 3.244, 1.394, 0.505, [['80%', C.accent4]], 24, { align: 'center', valign: 'middle', rotate: -4.07 });

  stat(s, 10.331, 3.236, '+8590', CAPTION, { color: C.white });
  stat(s, 10.331, 4.421, '+8590', CAPTION, { color: C.white });
  masterChrome(s, 10);
}

function slide11(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 6.306, y: 0, w: 1.62, h: 7.5, fill: { color: C.accent4 }, line: { type: 'none' } });
  imagePlaceholder(s, 4.706, 1.389, 4.169, 6.111, 0.16);

  heading(s, 0.828, 1.312, 5.186, 1.582, [['Our Expert ', C.ink], ['Team', C.accent4], [' Leader', C.ink]], 44);
  body(s, 0.828, 3.233, 4.061, 0.97, MAECENAS_LECTUS);

  // Two skill bars: grey track with a coloured progress segment.
  [
    { y: 4.784, value: '+8590', color: C.accent4, fill: 1.669 },
    { y: 5.533, value: '+8590', color: C.ink, fill: 2.617 },
  ].forEach((bar) => {
    heading(s, 0.851, bar.y, 1.2, 0.337, [[bar.value, bar.color]], 14);
    s.addShape('line', { x: 0.97, y: bar.y + 0.505, w: 3.242, h: 0, line: { color: C.rule, width: 5 } });
    s.addShape('line', { x: 0.97, y: bar.y + 0.505, w: bar.fill, h: 0, line: { color: bar.color, width: 5 } });
  });

  quoteMark(s, 8.627, 1.63);
  s.addText(
    [
      { text: 'Drive with confidence', options: { bold: true } },
      { text: '\u2014our expert repairs make every journey safe.' },
    ],
    { x: 8.51, y: 1.865, w: 4.22, h: 0.667, fontFace: MINOR, fontSize: 12, color: C.ink, lineSpacingMultiple: 1.5, valign: 'top', margin: [3.5, 7, 3.5, 7] }
  );
  iconDisc(s, 9.238, 3.431, 0.637, ARROW);
  stat(s, 10.189, 3.321, '90%', CAPTION);

  card(s, 7.926, 5.33, 2.351, 0.781, 0.096);
  heading(s, 7.941, 5.499, 2.321, 0.37, [['Bailey Dupont', C.ink]], 16, { align: 'center' });
  pill(s, 9.003, 5.98, 'Job Position');
  masterChrome(s, 11);
}

function slide12(pptx) {
  const s = pptx.addSlide();
  heading(s, 0.862, 0.8, 5.805, 1.447, [['Our Expert ', C.ink], ['Automobi', C.accent4], [' Team', C.ink]], 40);
  body(s, 0.862, 2.701, 3.409, 0.97, MAECENAS_LECTUS);
  stat(s, 0.893, 4.123, '+8590', CAPTION);
  stat(s, 0.893, 5.165, '+8590', CAPTION);

  [
    { x: 4.866, name: 'Bartholomew' },
    { x: 7.542, name: 'Harper Russo' },
    { x: 10.218, name: 'Reese Miller' },
  ].forEach((p, i) => {
    imagePlaceholder(s, p.x, 2.565, 2.451, 3.653, 0.16);
    card(s, p.x + 0.649, 6.0, 1.622, 0.63, 0.077);
    heading(s, p.x + 0.3, 6.137, 2.321, 0.303, [[p.name, C.ink]], 12, { align: 'center' });
    pill(s, p.x + 1.039, 6.503, 'Job Position');
    if (i < 2) iconDisc(s, p.x + 2.323, 4.134, 0.481, ARROW);
  });
  masterChrome(s, 12);
}

function slide13(pptx) {
  const s = pptx.addSlide();
  [[0.96, 1.935], [3.157, 1.935], [5.354, 2.335]].forEach(([x, w]) => imagePlaceholder(s, x, 2.755, w, 3.552, 0.16));
  heading(s, 0.862, 0.748, 4.492, 1.582, [['Customer', C.accent4], [' Satisfaction', C.ink]], 44);

  quoteMark(s, 8.485, 3.514);
  s.addText(
    [
      { text: 'Their team is incredibly knowledgeable ', options: { bold: true } },
      { text: 'and always goes the extra mile to ensure my vehicle runs perfectly.' },
    ],
    { x: 8.386, y: 3.722, w: 4.792, h: 1.114, fontFace: MINOR, fontSize: 14, color: C.ink, lineSpacingMultiple: 1.5, valign: 'top', margin: [3.5, 7, 3.5, 7] }
  );
  for (let i = 0; i < 5; i++) {
    s.addShape('star5', { x: 8.668 + i * 0.174, y: 5.28, w: 0.159, h: 0.159, fill: { color: C.star }, line: { type: 'none' } });
  }

  card(s, 6.628, 5.045, 1.758, 0.63, 0.077);
  heading(s, 6.346, 5.172, 2.321, 0.303, [['Ronald Jose', C.ink]], 12, { align: 'center' });
  pill(s, 7.167, 5.585, 'Job Position');
  masterChrome(s, 13);
}

// Flat-vector engine hoist: a stack of plain rectangles, ordered back-to-front.
const HOIST_PARTS = [
  [5.51, 3.23, 0.33, 3.07, 'BC3908'], [6.18, 5.14, 0.91, 1.07, 'BC3908'], [5.36, 6.03, 2.65, 0.27, 'BC3908'],
  [5.62, 3.23, 0.22, 3.07, 'BC3908'], [5.34, 1.71, 2.85, 0.21, 'BC3908'], [6.05, 6.03, 1.96, 0.27, 'BC3908'],
  [5.82, 1.71, 2.47, 0.21, 'BC3908'], [5.56, 1.51, 0.22, 2.14, '2D373A'], [5.63, 1.51, 0.16, 2.14, '677477'],
  [5.21, 4.54, 0.36, 1.64, 'FFFFFF'], [5.74, 2.30, 0.57, 0.57, 'BC3908'], [5.79, 2.35, 0.52, 0.52, 'BC3908'],
  [5.42, 3.75, 0.26, 0.86, '677477'], [5.42, 3.75, 0.21, 0.86, 'B8C6C9'], [5.51, 1.28, 0.33, 0.33, '2D373A'],
  [5.56, 1.28, 0.33, 0.33, '677477'], [5.45, 2.67, 0.46, 0.20, 'BC3908'], [5.59, 2.67, 0.32, 0.20, 'BC3908'],
  [5.56, 1.20, 0.15, 0.42, 'BC3908'], [5.63, 1.20, 0.15, 0.42, 'BC3908'], [6.42, 1.83, 0.34, 0.22, 'BC3908'],
  [6.51, 1.83, 0.26, 0.22, 'BC3908'], [5.32, 4.25, 0.23, 0.17, '3C4749'], [5.26, 4.19, 0.15, 0.15, '3C4749'],
  [5.48, 4.24, 0.09, 0.29, '677477'], [5.48, 4.24, 0.07, 0.29, '212A2C'],
  // hoist arm + mast ladder rungs
  [7.27, 1.49, 0.51, 0.32, 'BC3908'], [7.40, 1.49, 0.38, 0.32, 'BC3908'], [7.34, 1.80, 0.36, 0.50, '384749'],
  [7.47, 1.80, 0.23, 0.50, '212A2C'], [7.29, 2.27, 0.48, 0.10, '384749'], [7.46, 2.27, 0.31, 0.10, '212A2C'],
  // engine block
  [6.81, 3.46, 1.27, 0.39, '859599'], [6.89, 3.18, 1.23, 0.27, '444F51'], [7.44, 3.18, 0.67, 0.27, '2D373A'],
  [6.76, 3.29, 0.87, 0.18, '677477'], [7.28, 3.29, 0.87, 0.18, '2D373A'], [6.72, 3.34, 0.96, 0.12, '96A7AA'],
  [7.24, 3.33, 0.96, 0.14, '677477'], [7.16, 3.49, 0.82, 0.77, '2D373A'], [7.38, 3.38, 0.70, 0.54, '2D373A'],
  [6.94, 3.98, 1.00, 0.25, '444F51'], [7.21, 3.80, 0.91, 0.27, '677477'], [6.78, 3.90, 1.00, 0.17, '677477'],
  [6.72, 3.82, 0.96, 0.12, '96A7AA'], [7.91, 3.38, 0.33, 0.37, '2D373A'], [7.93, 3.41, 0.27, 0.41, '4C585B'],
  [8.04, 3.33, 0.27, 0.38, '4C585B'], [7.83, 3.39, 0.21, 0.25, '3C4749'], [7.91, 3.39, 0.25, 0.25, 'C6C6C6'],
  [7.92, 3.40, 0.23, 0.23, '4C585B'], [8.09, 3.21, 0.14, 0.14, '677477'], [8.09, 3.24, 0.18, 0.14, '677477'],
  [8.09, 3.30, 0.18, 0.14, '677477'], [6.88, 3.09, 0.25, 0.19, '677477'], [7.04, 3.09, 0.25, 0.19, '677477'],
  [7.20, 3.09, 0.25, 0.19, '677477'], [7.36, 3.09, 0.25, 0.19, '677477'], [7.49, 3.09, 0.19, 0.25, '677477'],
  [6.88, 3.51, 0.07, 0.26, '677477'], [7.00, 3.51, 0.07, 0.26, '677477'], [7.13, 3.51, 0.07, 0.26, '677477'],
  [7.26, 3.51, 0.07, 0.26, '677477'],
  // pulleys + cradle
  [7.52, 3.84, 0.41, 0.41, '3C4749'], [7.57, 3.84, 0.41, 0.41, 'C6C6C6'], [7.60, 3.86, 0.36, 0.36, '4C585B'],
  [7.12, 3.56, 0.33, 0.33, '3C4749'], [7.17, 3.56, 0.33, 0.33, 'C6C6C6'], [7.19, 3.58, 0.29, 0.29, '4C585B'],
  [7.41, 3.37, 0.27, 0.32, '3C4749'], [7.50, 3.37, 0.32, 0.32, 'C6C6C6'], [7.51, 3.38, 0.29, 0.29, '4C585B'],
  [6.64, 4.19, 1.75, 0.13, 'BC3908'], [7.14, 4.19, 1.25, 0.13, 'BC3908'], [6.80, 4.07, 0.18, 0.18, 'BC3908'],
  [8.07, 4.07, 0.18, 0.18, 'BC3908'],
];

function slide14(pptx) {
  const s = pptx.addSlide();
  heading(s, 0.636, 2.38, 4.3, 1.582, [['Expert ', C.ink], ['Engine', C.accent4], [' Care', C.ink]], 44);
  body(s, 0.636, 4.151, 4.061, 0.97, MAECENAS_LECTUS);

  HOIST_PARTS.forEach(([x, y, w, h, color]) => {
    s.addShape('rect', { x, y, w, h, fill: { color }, line: { type: 'none' } });
  });
  // Four lifting chains fan out from the hook down to the cradle, over the engine.
  [[7.42, 6.89], [7.50, 7.46], [7.55, 7.57], [7.62, 8.16]].forEach(([topX, botX]) => {
    const len = Math.hypot(botX - topX, 1.78);
    s.addShape('rect', {
      x: (topX + botX) / 2 - len / 2, y: 3.215, w: len, h: 0.09,
      rotate: (Math.atan2(1.78, botX - topX) * 180) / Math.PI,
      fill: { color: 'C0D4D8' }, line: { type: 'none' },
    });
  });
  // diagonal brace of the mast base and the strut under the jib arm
  s.addShape('rect', { x: 6.21, y: 4.46, w: 0.23, h: 1.96, rotate: -45, fill: { color: C.accent4 }, line: { type: 'none' } });
  s.addShape('rect', { x: 6.01, y: 2.16, w: 0.78, h: 0.10, rotate: -45, fill: { color: '2D373A' }, line: { type: 'none' } });
  // black-and-yellow hazard stripes on the mast
  for (let i = 0; i < 9; i++) {
    s.addShape('rect', { x: 5.62, y: 3.51 + i * 0.26, w: 0.22, h: 0.24, fill: { color: '212A2C' }, line: { type: 'none' } });
  }

  [[2.087, WRENCH], [3.405, HAMMER], [4.773, BOLT]].forEach(([y, glyph], i) => {
    iconDisc(s, 9.264, y, 0.657, glyph);
    stat(s, 10.164, [1.986, 3.353, 4.721][i], '+8590', CAPTION);
  });
  masterChrome(s, 14);
}

// Speedometer: six block-arc segments sharing one ellipse, plus a needle.
// The band thickens as it sweeps clockwise from the yellow "10%" end.
const GAUGE = { x: 3.805, y: 3.443, w: 6.73, h: 5.398 };
const GAUGE_SEGMENTS = [
  { color: C.accent5, from: 180, to: 210.5, thickness: 0.037 },
  { color: C.accent4, from: 210, to: 239.5, thickness: 0.049 },
  { color: C.accent3, from: 238, to: 266.0, thickness: 0.169 },
  { color: C.accent2, from: 265.5, to: 293.5, thickness: 0.179 },
  { color: '491106', from: 292, to: 324.5, thickness: 0.196 },
  { color: C.accent1, from: 321, to: 357.5, thickness: 0.287 },
];

function slide15(pptx) {
  const s = pptx.addSlide();
  heading(s, 0.584, 0.474, 4.992, 1.447, [['Expert Car ', C.ink], ['Tuning', C.accent4], [' Services', C.ink]], 40);
  body(s, 0.584, 1.927, 4.933, 0.667, MAECENAS);

  // Ghosted step numbers behind the dial.
  heading(s, 5.266, 2.672, 1.5, 0.64, [['03', C.rule]], 32, { align: 'center' });
  heading(s, 10.166, 4.731, 1.5, 0.64, [['06', C.rule]], 32, { align: 'center' });

  GAUGE_SEGMENTS.forEach((seg) => {
    s.addShape('blockArc', {
      x: GAUGE.x, y: GAUGE.y, w: GAUGE.w, h: GAUGE.h,
      angleRange: [seg.from, seg.to], arcThicknessRatio: seg.thickness,
      fill: { color: seg.color }, line: { type: 'none' },
    });
  });

  // Needle pointing at ~80%, with its pivot hub (rotation is about the centre).
  s.addShape('rect', { x: 6.417, y: 5.575, w: 2.35, h: 0.13, rotate: -45.3, fill: { color: C.rule }, line: { type: 'none' } });
  s.addShape('ellipse', { x: 6.49, y: 6.2, w: 0.55, h: 0.55, fill: { color: C.rule }, line: { type: 'none' } });

  // Tick labels follow the curve of the dial.
  [
    ['10%', 3.674, 5.512, -59],
    ['20%', 4.508, 4.612, -33.7],
    ['40%', 5.594, 4.147, -10.8],
    ['70%', 6.813, 4.197, 15.6],
    ['80%', 7.793, 4.664, 36.8],
    ['100%', 8.531, 5.456, 58.3],
  ].forEach(([text, x, y, rot]) => {
    heading(s, x, y, 1.519, 0.337, [[text, C.ink]], 14, { align: 'center', rotate: rot });
  });

  // Four call-outs around the dial: coloured dot + figure + caption.
  [
    { dot: [3.197, 4.969], color: C.accent5, glyph: WRENCH, statX: 0.031, statY: 4.969, align: 'right' },
    { dot: [4.271, 3.599], color: C.accent4, glyph: HAMMER, statX: 1.148, statY: 3.533, align: 'right' },
    { dot: [7.451, 2.563], color: C.accent2, glyph: HAMMER, statX: 8.256, statY: 2.498, align: 'left' },
    { dot: [9.601, 3.485], color: '491106', glyph: BOLT, statX: 10.406, statY: 3.419, align: 'left' },
  ].forEach((cb) => {
    s.addShape('ellipse', { x: cb.dot[0], y: cb.dot[1], w: 0.554, h: 0.554, fill: { color: cb.color }, line: { type: 'none' }, shadow: cardShadow() });
    iconGlyph(s, cb.dot[0] + 0.127, cb.dot[1] + 0.127, 0.3, cb.glyph, C.white);
    stat(s, cb.statX, cb.statY, '+8590', CAPTION, { size: 14, align: cb.align });
  });
  masterChrome(s, 15);
}

const PLANS = [
  { x: 1.211, title: 'Basic Tune-Up ', price: '$299', items: ['ECU diagnostics', 'Air-fuel ratio optimization', 'Ignition timing adjustments'] },
  { x: 5.072, title: 'Performance Boost ', price: '$599', items: ['Comprehensive ECU remapping', 'Custom exhaust tuning', 'Intake enhancements'] },
  { x: 8.932, title: 'Ultimate Tuning ', price: '$999', items: ['Full vehicle performance', 'Suspension adjustments', 'Dyno testing'] },
];

function slide16(pptx) {
  const s = pptx.addSlide();
  heading(s, 3.605, 0.62, 6.124, 1.313, [['The Perfect Plan for ', C.ink], ['Your Needs', C.accent4]], 36, { align: 'center' });

  PLANS.forEach((plan) => {
    card(s, plan.x, 2.578, 3.19, 3.984, 0.171);
    heading(s, plan.x + 0.283, 2.913, 2.251, 0.707, [[plan.title, C.ink], ['Package', C.accent4]], 18);
    heading(s, plan.x + 0.283, 3.77, 2.387, 0.707, [[plan.price, C.ink]], 36, { valign: 'middle' });
    body(s, plan.x + 0.283, 4.443, 1.487, 0.202, 'Your Subtitle Here', { fontSize: 6, lineSpacingMultiple: 1 });
    body(s, plan.x + 0.283, 4.733, 1.487, 0.303, 'Includes:', { lineSpacingMultiple: 1 });
    plan.items.forEach((item, i) => {
      const y = 5.124 + i * 0.341;
      s.addShape('ellipse', { x: plan.x + 0.384, y: y + 0.076, w: 0.102, h: 0.102, fill: { color: C.accent4 }, line: { type: 'none' } });
      body(s, plan.x + 0.506, y, 1.815, 0.252, item, { fontSize: 9, lineSpacingMultiple: 1 });
    });
    pill(s, plan.x + 1.786, 6.473, 'Order Now!');
  });

  iconDisc(s, 7.82, 2.386, 0.637, ARROW);
  masterChrome(s, 16);
}

function slide17(pptx) {
  const s = pptx.addSlide();
  // Tablet mock-up bleeding off the left edge: black bezel + grey screen.
  s.addShape('roundRect', { x: -0.4, y: 1.386, w: 6.42, h: 4.73, rectRadius: 0.2, fill: { color: '111111' }, line: { type: 'none' } });
  imagePlaceholder(s, -0.4, 1.61, 5.99, 4.28, 0);
  iconDisc(s, 2.459, 3.558, 0.637, '\u25B6');

  heading(s, 7.016, 1.629, 4.885, 2.121, [['Short Video From A ', C.ink], ['Happy Customer', C.accent4]], 40);
  body(s, 7.016, 3.876, 4.687, 0.667, MAECENAS_LECTUS);

  card(s, 3.935, 5.135, 5.081, 1.363, 0.098);
  iconDisc(s, 4.356, 5.498, 0.637, ARROW);
  discover(s, 5.235, 5.315, 3.84, LOREM_SHORT);
  pill(s, 7.59, 6.31, 'Explore Now!');
  stat(s, 9.552, 5.496, '+8590', CAPTION);
  masterChrome(s, 17);
}

function slide18(pptx) {
  const s = pptx.addSlide();
  imagePlaceholder(s, 1.046, 3.75, 11.241, 2.464, 0.16);
  heading(s, 1.415, 0.989, 8.231, 1.313, [['Break ', C.ink], ['Slide', C.accent4]], 72);
  iconDisc(s, 7.772, 1.453, 0.549, ARROW);

  card(s, 1.583, 2.869, 3.625, 1.763, 0.127);
  quoteMark(s, 1.996, 3.205);
  heading(s, 1.875, 3.502, 3.426, 0.808, [[TAGLINE, C.ink]], 14);

  card(s, 5.69, 5.479, 5.833, 1.47, 0.106);
  heading(s, 5.209, 5.709, 6.797, 1.01, [['60 Minutes', C.accent4]], 54, { align: 'center', valign: 'middle' });
  masterChrome(s, 18);
}

function slide19(pptx) {
  const s = pptx.addSlide();
  imagePlaceholder(s, 7.111, 0, 6.222, 7.5, 0);
  heading(s, 1.217, 0.834, 5.297, 1.447, [['Keep in Touch ', C.ink], ['with Us', C.accent4]], 40);
  body(s, 1.217, 2.37, 4.687, 0.667, MAECENAS_LECTUS);

  card(s, 1.217, 3.622, 8.214, 2.433, 0.175);
  [
    { x: 1.804, y: 4.182, glyph: '@', text: 'loremipsum@yourmail.com' },
    { x: 1.804, y: 5.091, glyph: '\u260E', text: '+0123 4567 890' },
    { x: 5.887, y: 4.182, glyph: '\u2295', text: 'Your Location Here' },
    { x: 5.887, y: 5.091, glyph: '\u2318', text: 'www.yourwebsite.com' },
  ].forEach((row) => {
    s.addShape('roundRect', { x: row.x, y: row.y, w: 0.511, h: 0.511, rectRadius: 0.09, fill: { color: C.accent4 }, line: { type: 'none' } });
    s.addText(row.glyph, { x: row.x, y: row.y, w: 0.511, h: 0.511, align: 'center', valign: 'middle', fontFace: 'DejaVu Sans', fontSize: 16, color: C.white });
    body(s, row.x + 0.741, row.y + 0.041, 3.486, 0.376, row.text, { fontFace: MAJOR });
  });
  masterChrome(s, 19);
}

function slide20(pptx) {
  const s = pptx.addSlide();
  gradientScrim(s, 0, 0, 13.333, 7.5); // closing photo is fully covered by the scrim
  heading(s, 0.903, 3.026, 9.634, 1.447, [['Thank You', C.white]], 80);
  body(s, 1.152, 4.286, 3.389, 0.303, 'For Joining this Presentation', { color: C.white, lineSpacingMultiple: 1 });
  pill(s, 6.132, 4.077, 'Join Now!');
}

/* -------------------------------------------------------------------- main */

const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'AUTOMOBI', width: 13.333, height: 7.5 });
  pptx.layout = 'AUTOMOBI';
  pptx.theme = { headFontFace: MAJOR, bodyFontFace: MINOR };
  pptx.title = 'Automobi - Car Repair Presentation Template';

  BUILDERS.forEach((fn) => fn(pptx));

  return pptx.writeFile({ fileName: path.join(__dirname, '15be576d-0df1-4acd-ae68-69b20b836408_grok_final.pptx') });
}

build().then((f) => console.log('wrote', f));
