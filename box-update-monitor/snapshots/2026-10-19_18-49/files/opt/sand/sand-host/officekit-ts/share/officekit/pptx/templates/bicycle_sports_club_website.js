/**
 * "Jengky" presentation template — rebuilt with PptxGenJS.
 *
 *   node 07c58c37-7d3c-4d8c-a3f5-e83bc8197bee_grok_final.js
 *      -> 07c58c37-7d3c-4d8c-a3f5-e83bc8197bee_grok_final.pptx
 *
 * 30 slides, 13.333 x 7.5 in. Photographs in the original deck are drawn here as
 * labelled "[image]" placeholder blocks; every other element is a native shape.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const INK = '262626'; // tx1 85% — body/heading text
const INK_SOFT = '404040'; // tx1 75% — paragraph copy
const BLACK = '0D0D0D'; // tx1 95% — dark backgrounds & icons
const WHITE = 'FFFFFF';
const YELLOW = 'FFC000'; // brand yellow (accents on dark)
const AMBER = 'F0A20A'; // brand amber (accents on light)
const PLACEHOLDER_FILL = 'E6E6E6';
const PLACEHOLDER_LINE = 'BFBFBF';

const BODY = 'Work Sans';
const SEMI = 'Work Sans SemiBold';

/* ------------------------------------------------------------------- copy */

const LOREM = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
  'incididunt ut labore et dolore magna ali qua. Ut enim ad minaun veniam, quis nostrud ' +
  'exercitation u llamco laboris nisi ut aliquip ex ea commodo consequat.';
const LOREM_MED = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
  'incididunt ut labore et dolore magna ali qua. Ut enim ad minaun veniam, ';
const LOREM_CARD = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
  'incididunt ut labore et dolore magna ali qua. Ut enim ad minaun veniam, quis nostrud ' +
  'exercitation u llamco laboris nisi ut aliqui';
const LOREM_TINY = 'Lorem ipsum dolor sata amet, consectetur adipiscing elit, ';
const TEMPLATE_TAG = 'Jengky Template';
const CLUB = 'Bicycle Sports Club';
const NAV_4 = 'HOME                    ABOUT                    SERVICES                    CONTACT US';
const NAV_3 = 'HOME                    ABOUT                    CONTACT US';

/* ---------------------------------------------------------------- helpers */

/** Plain text box. Reference boxes are top-anchored and grow to fit their text. */
function text(slide, runs, o) {
  slide.addText(runs, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: o.face || BODY, fontSize: o.size, bold: o.bold || false,
    color: o.color || INK, align: o.align || 'left', valign: 'top',
    lineSpacingMultiple: o.lnSpc, wrap: o.wrap !== false, fit: 'resize',
  });
}

function rect(slide, o) {
  slide.addShape('rect', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: { color: o.fill, transparency: o.transparency || 0 },
  });
}

/** Straight connector (the thin rules under the footer labels). */
function rule(slide, x, y, w, color) {
  slide.addShape('line', { x: x, y: y, w: w, h: 0, line: { color: color, width: 0.75 } });
}

/**
 * Freeform shape. Each entry of `pts` is normalised to the shape box: either
 * [x, y] for a straight segment or [x, y, cx1, cy1, cx2, cy2] for a cubic.
 */
function freeform(slide, o) {
  const pts = o.pts.map((p) => {
    const pt = { x: p[0] * o.w, y: p[1] * o.h };
    if (p.length === 2) return pt;
    return Object.assign(pt, {
      curve: { type: 'cubic', x1: p[2] * o.w, y1: p[3] * o.h, x2: p[4] * o.w, y2: p[5] * o.h },
    });
  });
  pts.push({ close: true });
  slide.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h, points: pts,
    fill: o.fill ? { color: o.fill, transparency: o.transparency || 0 } : { type: 'none' },
    line: o.line ? { color: o.line, width: o.lineWidth || 1, transparency: o.lineTransparency || 0 } : null,
  });
}

/** Bike-saddle logo mark, drawn from scratch so no bitmap is needed. */
function logo(slide, x, y, color) {
  const s = 0.4414;
  freeform(slide, {
    x: x, y: y, w: s, h: s, line: color, lineWidth: 1.25,
    pts: [[0.06, 0.30],
      [0.34, 0.10, 0.06, 0.16, 0.17, 0.10],
      [0.98, 0.21, 0.60, 0.10, 0.88, 0.14],
      [0.60, 0.34, 1.0, 0.30, 0.80, 0.33],
      [0.24, 0.50, 0.44, 0.36, 0.34, 0.48],
      [0.06, 0.30, 0.12, 0.52, 0.03, 0.43]],
  });
  // seat-post clamp and post
  slide.addShape('rect', { x: x + 0.20 * s, y: y + 0.46 * s, w: 0.30 * s, h: 0.10 * s,
    fill: { type: 'none' }, line: { color: color, width: 1.25 } });
  slide.addShape('line', { x: x + 0.35 * s, y: y + 0.56 * s, w: 0, h: 0.38 * s, line: { color: color, width: 1.75 } });
  text(slide, 'Jengky', { x: x + 0.2206, y: y + 0.1385, w: 0.8856, h: 0.3366, face: SEMI, size: 14, color: color, wrap: false });
}

/** Footer: club label plus its underline. */
function footer(slide, o) {
  text(slide, CLUB, {
    x: o.x, y: o.y, w: 2.3065, h: 0.3029, face: SEMI, size: 12,
    color: o.color, align: o.align || 'left',
  });
  rule(slide, o.lineX, o.lineY, o.lineW || 4.8818, o.lineColor);
}

/** Footer flavour used by the light "content" slides (amber rule, ink label). */
function footerLight(slide, side, lineW) {
  if (side === 'left') footer(slide, { x: 0.5133, y: 6.4254, color: INK, lineX: 0.5994, lineY: 6.8798, lineW: lineW, lineColor: AMBER });
  else footer(slide, { x: 10.3878, y: 6.4254, color: INK, align: 'right', lineX: 7.8125, lineY: 6.8798, lineW: lineW, lineColor: AMBER });
}

/** Footer flavour used by the dark cover slides (white rule + label + blurb). */
function footerDark(slide) {
  footer(slide, { x: 0.5133, y: 6.2825, color: WHITE, lineX: 0.5994, lineY: 6.7369, lineColor: WHITE });
  text(slide, LOREM_TINY, { x: 10.3117, y: 6.2728, w: 2.5084, h: 0.6253, size: 11, color: WHITE, align: 'right', lnSpc: 1.5 });
}

/** Top-right navigation strip; `active` is the item painted in the accent colour. */
function nav(slide, o) {
  let runs = o.items;
  if (o.active) {
    const [before, after] = o.items.split(o.active);
    runs = [{ text: before, options: { color: o.color } },
      { text: o.active, options: { color: o.accent } },
      { text: after, options: { color: o.color } }].filter((r) => r.text !== '');
  }
  text(slide, runs, { x: o.x, y: o.y, w: o.w, h: 0.3252, face: SEMI, size: 10, color: o.color, lnSpc: 1.5 });
}

/** The "Jengky Template" eyebrow that sits above most headlines. */
function eyebrow(slide, x, y) {
  text(slide, TEMPLATE_TAG, { x: x, y: y, w: 1.7193, h: 0.3384, size: 10.5, bold: true, color: AMBER, lnSpc: 1.5 });
}

/** The recurring two-tone headline, used at several sizes. */
function beautyHeadline(slide, o) {
  const runs = [
    ['There is beauty in ', INK], ['silence and there is silence ', AMBER],
    ['in beauty and ', INK], ['you can find ', AMBER], ['both in a bicycle', INK],
  ].map((r) => ({ text: r[0], options: { color: r[1] } }));
  text(slide, runs, { x: o.x, y: o.y, w: o.w, h: o.h, size: o.size, bold: true });
}

/** Body paragraph (justified, 1.5 line spacing) used all over the deck. */
function para(slide, body, o) {
  text(slide, body, {
    x: o.x, y: o.y, w: o.w, h: o.h, size: o.size || 10.5,
    color: o.color || INK_SOFT, align: 'justify', lnSpc: 1.5,
  });
}

/** Yellow call-to-action button. */
function button(slide, x, y, label) {
  slide.addShape('rect', { x: x, y: y, w: 1.9251, h: 0.4611, fill: { color: YELLOW } });
  slide.addText(label, {
    x: x, y: y, w: 1.9251, h: 0.4611, fontFace: SEMI, fontSize: 14,
    color: BLACK, align: 'center', valign: 'middle',
  });
}

/**
 * Stand-in for a photograph / device mock-up from the original deck. It is kept
 * translucent so the coloured panels the originals sit on still read through.
 */
function imagePlaceholder(slide, o) {
  slide.addShape('rect', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: { color: PLACEHOLDER_FILL, transparency: 45 }, line: { color: PLACEHOLDER_LINE, width: 1 },
  });
  slide.addText('[image]', {
    x: o.x, y: o.y, w: o.w, h: o.h, fontFace: BODY, fontSize: 12,
    color: '9E9E9E', align: 'center', valign: 'middle',
  });
}

/* --------------------------------------------------------- slide builders */

// 1 — cover: dark band across a white page.
function slide01(s) {
  rect(s, { x: 0, y: 0.807, w: 13.3333, h: 5.42, fill: BLACK });
  text(s, [{ text: 'Jen', options: { color: WHITE } }, { text: 'gky', options: { color: YELLOW } }],
    { x: 0.8494, y: 1.6935, w: 4.9983, h: 1.7166, size: 96, bold: true, wrap: false });
  text(s, 'Presentation Template.', { x: 0.8494, y: 3.2586, w: 2.3065, h: 0.3029, size: 12, color: WHITE });
  nav(s, { x: 7.7222, y: 0.2194, w: 5.26, items: NAV_4, active: 'HOME', accent: AMBER, color: INK });
  rule(s, 0.8948, 4.7342, 4.8818, WHITE);
  text(s, CLUB, { x: 0.8494, y: 4.2798, w: 2.3065, h: 0.3029, face: SEMI, size: 12, color: WHITE });
  text(s, LOREM_TINY, { x: 0.8494, y: 4.8857, w: 2.5084, h: 0.6253, size: 11, color: WHITE, align: 'justify', lnSpc: 1.5 });
  logo(s, 0.5133, 0.1318, INK);
}

// 2 — cover on a dimmed photo (70% dark wash).
function slide02(s) {
  rect(s, { x: 0, y: 0, w: 13.3333, h: 7.5, fill: BLACK, transparency: 30 });
  text(s, [{ text: 'Jen', options: { color: WHITE } }, { text: 'gky', options: { color: YELLOW } }],
    { x: 3.9532, y: 2.4588, w: 4.9983, h: 1.7166, size: 96, bold: true, wrap: false });
  text(s, 'Presentation Template.', { x: 3.9532, y: 4.0239, w: 2.3065, h: 0.3029, size: 12, color: WHITE });
  nav(s, { x: 9.2426, y: 0.3827, w: 3.7982, items: NAV_3, active: 'HOME', accent: YELLOW, color: WHITE });
  logo(s, 0.5133, 0.295, WHITE);
  footerDark(s);
}

// 3 — same cover, flat dark background.
function slide03(s) {
  rect(s, { x: 0, y: 0, w: 13.3333, h: 7.5, fill: BLACK });
  text(s, [{ text: 'Jen', options: { color: WHITE } }, { text: 'gky', options: { color: YELLOW } }],
    { x: 3.9532, y: 2.4588, w: 4.9983, h: 1.7166, size: 96, bold: true, wrap: false });
  text(s, 'Presentation Template.', { x: 3.9532, y: 4.0239, w: 2.3065, h: 0.3029, size: 12, color: WHITE });
  nav(s, { x: 9.2426, y: 0.3827, w: 3.7982, items: NAV_3, active: 'HOME', accent: YELLOW, color: WHITE });
  logo(s, 0.5133, 0.295, WHITE);
  footerDark(s);
}

// 4 — "About Jengky": yellow column left, copy right.
function slide04(s) {
  rect(s, { x: 0, y: 0, w: 3.0833, h: 7.5, fill: YELLOW });
  text(s, 'About Jengky', { x: 7.1583, y: 1.966, w: 4.9227, h: 0.9088, size: 48, bold: true });
  eyebrow(s, 7.1583, 1.6275);
  para(s, LOREM, { x: 7.1583, y: 3.0518, w: 4.75, h: 1.1316 });
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
    'incididunt ut labore et dolore magna ali qua. Ut enim ad minaun veniam.',
    { x: 7.1583, y: 4.3605, w: 4.75, h: 0.8665 });
  logo(s, 11.7654, 0.4073, INK);
  footerLight(s, 'right');
}

// 5 — headline + "Read More", yellow column on the right.
function slide05(s) {
  rect(s, { x: 10.2717, y: 0, w: 3.0833, h: 7.5, fill: YELLOW });
  eyebrow(s, 0.9546, 1.7736);
  beautyHeadline(s, { x: 0.9546, y: 2.1405, w: 6.3854, h: 1.3127, size: 24 });
  para(s, LOREM, { x: 0.9546, y: 3.6047, w: 6.1711, h: 0.8665 });
  button(s, 1.0443, 4.7347, 'Read More');
  nav(s, { x: 4.3217, y: 0.4806, w: 5.26, items: NAV_4, active: 'ABOUT', accent: AMBER, color: INK });
  logo(s, 0.5133, 0.4073, INK);
  logo(s, 11.7654, 0.4073, INK);
  footerLight(s, 'left');
}

// 6 — full-bleed photo left, two paragraphs right.
function slide06(s) {
  eyebrow(s, 6.4036, 1.4654);
  beautyHeadline(s, { x: 6.4036, y: 1.8324, w: 6.3854, h: 1.3127, size: 24 });
  para(s, LOREM, { x: 6.4036, y: 3.2966, w: 6.1711, h: 0.8665 });
  para(s, LOREM, { x: 6.4036, y: 4.3145, w: 6.1711, h: 0.8665 });
  logo(s, 11.7654, 0.4073, INK);
  footerLight(s, 'right');
}

// 7 — headline, sub-head and copy; yellow column right.
function slide07(s) {
  rect(s, { x: 10.6944, y: 0, w: 2.6616, h: 7.5, fill: YELLOW });
  eyebrow(s, 1.0669, 1.7736);
  beautyHeadline(s, { x: 1.0669, y: 2.1405, w: 6.3854, h: 1.3127, size: 24 });
  text(s, 'Laboris nisi ut aliquip ex ea commodo.',
    { x: 1.0669, y: 3.6594, w: 6.1711, h: 0.4597, face: SEMI, size: 16, align: 'justify', lnSpc: 1.5 });
  para(s, LOREM, { x: 1.0669, y: 4.18, w: 6.1711, h: 0.8665 });
  logo(s, 0.5133, 0.4073, INK);
  footerLight(s, 'left');
}

// 8 — yellow card with two stacked blurbs.
function slide08(s) {
  s.addShape('roundRect', { x: 3.3658, y: 3.398, w: 9.073, h: 3.591, fill: { color: YELLOW }, rectRadius: 0.141 });
  eyebrow(s, 6.4864, 1.1158);
  beautyHeadline(s, { x: 6.4864, y: 1.4827, w: 6.3854, h: 1.5146, size: 28 });
  [4.0294, 5.3696].forEach((y) => {
    text(s, CLUB, { x: 3.9326, y: y, w: 2.9069, h: 0.3702, size: 16, bold: true });
    para(s, LOREM_CARD, { x: 3.9326, y: y + 0.4004, w: 7.8328, h: 0.6014, color: INK });
  });
  logo(s, 11.7654, 0.4073, INK);
}

// 9 — yellow column plus a pair of faint decorative curves from the layout.
function slide09(s) {
  rect(s, { x: 10.6944, y: 0, w: 2.6616, h: 7.5, fill: YELLOW });
  [4.7423, 4.9872].forEach((y) => freeform(s, {
    x: 9.3271, y: y, w: 4.0106, h: 2.4463, line: 'ADC6E5', lineWidth: 6, lineTransparency: 95,
    pts: [[0, 0.536], [0.1932, 0.012, 0.0514, 0.2359, 0.1028, -0.0642],
      [0.5425, 0.9931, 0.2837, 0.0882, 0.4132, 0.9206],
      [0.9689, 0.4468, 0.6718, 1.0655, 0.8936, 0.5496], [0.9999, 0.3873]],
  }));
  eyebrow(s, 0.5994, 1.295);
  beautyHeadline(s, { x: 0.5994, y: 1.662, w: 6.3854, h: 1.3127, size: 24 });
  para(s, LOREM, { x: 0.5994, y: 3.1261, w: 6.1711, h: 0.8665 });
  para(s, LOREM_MED, { x: 0.5994, y: 4.2905, w: 3.3087, h: 1.1316 });
  logo(s, 0.5133, 0.4073, INK);
  footerLight(s, 'left', 3.3087);
}

// 10 — pull quote next to a full-height photo.
function slide10(s) {
  eyebrow(s, 0.7339, 1.8042);
  beautyHeadline(s, { x: 0.7339, y: 2.1712, w: 5.229, h: 2.7937, size: 32 });
  text(s, '\u201D', { x: 2.6308, y: 4.2376, w: 1.4354, h: 2.0364, size: 115, bold: true });
  logo(s, 0.5133, 0.4073, INK);
  footerLight(s, 'left');
}

// 11 — "Break Slides" section divider on dark.
function slide11(s) {
  rect(s, { x: 0, y: 0, w: 13.3333, h: 7.5, fill: BLACK });
  text(s, [{ text: 'Break ', options: { color: WHITE } }, { text: 'Slides', options: { color: YELLOW } }],
    { x: 2.9356, y: 2.7667, w: 7.462, h: 1.4473, size: 80, bold: true, align: 'center' });
  nav(s, { x: 9.2426, y: 0.3827, w: 3.7982, items: NAV_3, color: WHITE });
  logo(s, 0.5133, 0.295, WHITE);
  footerDark(s);
}

// 12 — same divider stacked over a dark photo panel.
function slide12(s) {
  rect(s, { x: 0.0052, y: 0, w: 13.3281, h: 7.5, fill: BLACK });
  text(s, 'Presentation Template.', { x: 7.8124, y: 1.7485, w: 2.3065, h: 0.3029, size: 12, color: WHITE });
  text(s, [{ text: 'Break', options: { color: WHITE, breakLine: true } }, { text: 'Slides', options: { color: YELLOW } }],
    { x: 7.8124, y: 1.8999, w: 4.2428, h: 3.3322, size: 96, bold: true, wrap: false });
  logo(s, 12.0552, 0.295, WHITE);
  text(s, LOREM_TINY, { x: 10.3117, y: 6.2728, w: 2.5084, h: 0.6253, size: 11, color: WHITE, align: 'right', lnSpc: 1.5 });
}

// 13 — "Time to break." over a photo (white half of the title is intentional).
function slide13(s) {
  text(s, [{ text: 'Time to ', options: { color: WHITE } }, { text: 'break.', options: { color: YELLOW } }],
    { x: 1.2577, y: 2.1953, w: 4.8437, h: 2.7937, size: 80, bold: true });
  logo(s, 0.5133, 0.295, WHITE);
  footer(s, { x: 0.5133, y: 6.2825, color: WHITE, lineX: 0.5994, lineY: 6.7369, lineColor: WHITE });
}

// 14 — wide photo band on top, headline + note below.
function slide14(s) {
  eyebrow(s, 0.7003, 5.139);
  beautyHeadline(s, { x: 0.7003, y: 5.506, w: 6.3854, h: 1.3127, size: 24 });
  text(s, 'Sed do eiusmod tempor incididunt',
    { x: 7.7279, y: 5.3903, w: 4.9418, h: 0.4148, face: SEMI, size: 14, align: 'justify', lnSpc: 1.5 });
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
    'incididunt ut labore et dolore magna ali qua. Ut enim ad minaun veniam, quis nostrud',
    { x: 7.7279, y: 5.8706, w: 4.9418, h: 0.8665 });
}

// 15 — photo top-left, headline right, note bottom-left.
function slide15(s) {
  eyebrow(s, 6.3089, 1.5664);
  beautyHeadline(s, { x: 6.3089, y: 1.9333, w: 6.3854, h: 1.3127, size: 24 });
  para(s, LOREM, { x: 6.3089, y: 3.3975, w: 6.1711, h: 0.8665 });
  para(s, LOREM, { x: 6.3089, y: 4.4154, w: 6.1711, h: 0.8665 });
  text(s, 'Laboris nisi ut aliquip ex ea commodo.',
    { x: 0.5806, y: 5.0879, w: 4.7221, h: 0.4597, face: SEMI, size: 16, align: 'justify', lnSpc: 1.5 });
  para(s, LOREM, { x: 0.5806, y: 5.6085, w: 4.823, h: 1.1316 });
  logo(s, 11.7654, 0.4073, INK);
  footerLight(s, 'right');
}

// 16 — mirror of slide 8: yellow card on the left.
function slide16(s) {
  s.addShape('roundRect', { x: 0.6685, y: 3.3796, w: 9.073, h: 3.591, fill: { color: YELLOW }, rectRadius: 0.141 });
  eyebrow(s, 0.904, 1.1469);
  beautyHeadline(s, { x: 0.904, y: 1.5138, w: 5.1969, h: 1.3127, size: 24 });
  [4.011, 5.3513].forEach((y) => {
    text(s, CLUB, { x: 1.2354, y: y, w: 2.9069, h: 0.3702, size: 16, bold: true });
    para(s, LOREM_CARD, { x: 1.2354, y: y + 0.4004, w: 7.8328, h: 0.6014, color: INK });
  });
  logo(s, 0.5133, 0.4073, INK);
}

// 17 — "Statistic diagram": six spire bars with callout labels.
const BARS = [
  { year: '2020', label: '82%', top: 1.9645, badge: 1.5940, dark: false },
  { year: '2021', label: '46%', top: 3.6368, badge: 3.2090, dark: true },
  { year: '2022', label: '70%', top: 2.9608, badge: 2.5357, dark: false },
  { year: '2023', label: '30%', top: 4.2653, badge: 3.8244, dark: true },
  { year: '2024', label: '75%', top: 2.2966, badge: 1.8994, dark: false },
  { year: '2025', label: '72%', top: 2.6761, badge: 2.2639, dark: true },
];
const BAR_X0 = 6.7908; // left edge of the first spire
const BAR_STEP = 0.9697;
const BAR_W = 1.313;
const BAR_BASE = 5.4558; // shared baseline

function slide17(s) {
  BARS.forEach((bar, i) => {
    const x = BAR_X0 + i * BAR_STEP;
    const h = BAR_BASE - bar.top;
    const color = bar.dark ? INK : YELLOW;
    // spire + a darker sliver that fakes the folded-paper shading
    freeform(s, { x: x, y: bar.top, w: BAR_W, h: h, fill: color, pts: [[0.4963, 0], [1, 1], [0, 1]] });
    freeform(s, { x: x + 0.6613, y: bar.top + 0.0213, w: 0.6517, h: h - 0.0213, fill: BLACK, transparency: 80, pts: [[0, 0], [0, 1], [1, 1]] });
    // callout badge with its little pointer
    s.addShape('rect', { x: x + 0.3483, y: bar.badge, w: 0.6214, h: 0.3645, fill: { color: color } });
    s.addText(bar.label, {
      x: x + 0.3483, y: bar.badge, w: 0.6214, h: 0.3645, fontFace: BODY, fontSize: 11,
      bold: true, color: bar.dark ? WHITE : INK, align: 'center', valign: 'middle',
    });
    s.addShape('triangle', { x: x + 0.5384, y: bar.badge + 0.3614, w: 0.2411, h: 0.1217, fill: { color: color }, rotate: 180 });
    text(s, bar.year, { x: x + 0.2565, y: 5.5772, w: 0.8, h: 0.2608, size: 11, bold: true, color: '3F3F3F', align: 'center' });
  });
  text(s, 'Statistic diagram', { x: 0.8632, y: 2.1809, w: 5.4822, h: 0.8415, size: 44, bold: true });
  eyebrow(s, 0.8632, 1.8139);
  para(s, LOREM, { x: 0.8632, y: 3.1443, w: 5.4006, h: 1.1316 });
  para(s, LOREM_MED, { x: 0.8632, y: 4.4273, w: 5.4006, h: 0.8665 });
  logo(s, 0.5133, 0.4073, INK);
  footerLight(s, 'left', 3.3087);
}

// 18 — "Cycle matrix & cubes": a quartered disc surrounded by four tumbling cubes.
const CYCLE_SHAPES = [
  // pie quadrants (clockwise from bottom-right)
  { x: 9.7509, y: 3.9507, w: 1.2221, h: 1.2054, fill: INK,
    pts: [[0, 0], [1, 0], [0.9951, 0.0981], [0.0093, 1, 0.9444, 0.6047, 0.5224, 1], [0, 0.9995]] },
  { x: 9.7507, y: 2.7352, w: 1.2224, h: 1.2166, fill: YELLOW,
    pts: [[0.0093, 0], [1, 0.9954, 0.5564, 0, 1, 0.4456], [0.9998, 1], [0, 1], [0, 0.0005]] },
  { x: 8.5595, y: 3.9507, w: 1.1994, h: 1.2054, fill: YELLOW,
    pts: [[0, 0], [1, 0], [1, 1], [0.9062, 0.9953], [0.0044, 0.0922, 0.4289, 0.9471, 0.0499, 0.568]] },
  { x: 8.5595, y: 2.7352, w: 1.1996, h: 1.2155, fill: INK,
    pts: [[1, 0], [1, 1], [0.0002, 1], [0, 0.9959], [0.9063, 0.0047, 0, 0.48, 0.3972, 0.0557]] },
  // top-right cube
  { x: 10.088, y: 1.4743, w: 0.5794, h: 2.089, fill: INK, pts: [[0.9649, 0], [0.0083, 0.8837], [0, 1], [1, 0.7908]] },
  { x: 10.0805, y: 3.126, w: 2.2985, h: 0.4369, fill: INK, pts: [[0.2521, 0], [0, 0.9945], [0.113, 1], [1, 0.1041]] },
  { x: 10.6391, y: 1.4759, w: 1.7394, h: 1.6987, fill: YELLOW, pts: [[0, 0], [0.0117, 0.9725], [1, 1], [0.9209, 0.1071]] },
  // top-left cube
  { x: 7.7095, y: 2.9798, w: 1.8124, h: 0.5794, fill: YELLOW, pts: [[0, 0.0558], [0.8587, 1], [1, 0.9959], [0.6057, 0]] },
  { x: 8.8033, y: 1.9367, w: 0.7147, h: 1.6209, fill: YELLOW, pts: [[0, 0.644], [1, 1], [0.9966, 0.8516], [0.0184, 0]] },
  { x: 7.7096, y: 1.9392, w: 1.1121, h: 1.075, fill: INK, pts: [[0.0431, 0.0679], [0, 1], [0.9871, 0.971], [1, 0]] },
  // bottom-left cube
  { x: 9.1325, y: 4.2185, w: 0.4058, h: 1.9669, fill: INK, pts: [[0.0265, 1], [0.9941, 0.1223], [1, 0], [0, 0.1601]] },
  { x: 7.4127, y: 4.2149, w: 2.1297, h: 0.3172, fill: INK, pts: [[0.8094, 1], [1, 0.0075], [0.8791, 0], [0, 0.8792]] },
  { x: 7.4186, y: 4.4821, w: 1.7393, h: 1.6939, fill: YELLOW,
    pts: [[0, 0], [0.0859, 0.8855, 0.0279, 0.2812, 0.058, 0.6043], [1, 1], [0.9931, 0.0248, 0.9977, 0.6749, 0.9954, 0.3499]] },
  // bottom-right cube
  { x: 10.0996, y: 4.2089, w: 0.3903, h: 1.1157, fill: YELLOW, pts: [[1, 0.2178], [0, 0], [0.0061, 0.22], [1, 1]] },
  { x: 10.0991, y: 4.2117, w: 1.3073, h: 0.243, fill: YELLOW, pts: [[1, 0.936], [0.1987, 0], [0, 0], [0.2985, 1]] },
  { x: 10.4894, y: 4.4364, w: 0.917, h: 0.8883, fill: INK, pts: [[0, 0.0175], [0, 1], [0.9726, 0.9461], [1, 0]] },
];

function slide18(s) {
  CYCLE_SHAPES.forEach((shape) => freeform(s, shape));
  text(s, [{ text: 'Cycle matrix', options: { color: INK } }, { text: ' & cubes ', options: { color: AMBER } },
    { text: 'infographic', options: { color: INK } }],
    { x: 0.8632, y: 2.1809, w: 5.9882, h: 1.582, size: 44, bold: true });
  eyebrow(s, 0.8632, 1.8139);
  para(s, LOREM, { x: 0.8632, y: 3.977, w: 5.7082, h: 1.1316 });
  logo(s, 0.5133, 0.4073, INK);
  footerLight(s, 'left', 3.3087);
}

// 19 — device mock-ups with three feature rows.
const MOCKUPS = ['Mockup One', 'Mockup Two', 'Mockup Three'];

function slide19(s) {
  rect(s, { x: 0, y: 0, w: 2.6616, h: 7.5, fill: YELLOW });
  imagePlaceholder(s, { x: 1.2417, y: 0.6315, w: 2.9652, h: 6.237 });
  imagePlaceholder(s, { x: 8.8346, y: 0.8641, w: 4.2322, h: 2.4497 });
  imagePlaceholder(s, { x: 8.8346, y: 4.1587, w: 4.2322, h: 2.4497 });
  MOCKUPS.forEach((title, i) => {
    const y = 1.2051 + i * 1.8449;
    s.addShape('rect', { x: 4.6488, y: y, w: 1.1511, h: 1.0938, fill: { color: YELLOW } });
    // handset glyph: rounded body, earpiece slot and home dot
    s.addShape('roundRect', { x: 5.0601, y: y + 0.2871, w: 0.3285, h: 0.5196, fill: { type: 'none' }, line: { color: BLACK, width: 2 }, rectRadius: 0.05 });
    s.addShape('roundRect', { x: 5.1271, y: y + 0.3357, w: 0.1961, h: 0.3419, fill: { type: 'none' }, line: { color: BLACK, width: 1.25 }, rectRadius: 0.02 });
    s.addShape('line', { x: 5.1925, y: y + 0.3357, w: 0.0654, h: 0, line: { color: BLACK, width: 1.25 } });
    s.addShape('ellipse', { x: 5.2093, y: y + 0.7246, w: 0.0318, h: 0.0335, fill: { color: BLACK } });
    text(s, title, { x: 5.9836, y: y, w: 2.6824, h: 0.4039, size: 18, bold: true });
    para(s, 'Lorem ipsum dolor sit ametas, consectetur adipi cing elit, sed',
      { x: 5.9836, y: y + 0.4004, w: 2.5126, h: 0.6014 });
  });
}

// 20 — laptop mock-up with a yellow column behind it.
function slide20(s) {
  rect(s, { x: 9.2708, y: 0, w: 4.0846, h: 7.5, fill: YELLOW });
  imagePlaceholder(s, { x: 5.4962, y: 1.219, w: 7.488, h: 4.3034 });
  eyebrow(s, 0.8632, 1.7629);
  text(s, 'Open with laptop slides', { x: 0.8632, y: 2.1299, w: 4.633, h: 1.3127, size: 36, bold: true });
  para(s, LOREM, { x: 0.8632, y: 3.594, w: 4.4633, h: 1.3966 });
  logo(s, 0.5133, 0.4073, INK);
  footerLight(s, 'left', 3.3087);
}

// 21 — two phone mock-ups plus a download call to action.
function slide21(s) {
  imagePlaceholder(s, { x: 0.2, y: 0, w: 3.036, h: 5.127 });
  imagePlaceholder(s, { x: 3.763, y: 2.145, w: 3.026, h: 5.355 });
  eyebrow(s, 7.628, 1.5911);
  beautyHeadline(s, { x: 7.628, y: 1.958, w: 5.1435, h: 1.3127, size: 24 });
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
    'incididunt ut labore et dolore magna ali qua. Ut enim ad minaun veniam, quis nostrud exercitation',
    { x: 7.628, y: 3.4222, w: 5.1435, h: 0.8665 });
  button(s, 7.7177, 4.5522, 'Download');
  logo(s, 11.7654, 0.4073, INK);
  footerLight(s, 'right');
}

// 22 — contact strip below a wide photo.
const CONTACTS = [
  { x: 4.4073, title: 'Email', lines: ['Jengky@mail.com', 'World.jengky@mail.com'] },
  { x: 7.3482, title: 'Social', lines: ['www.jengky.com', 'https://jengky.com'] },
  { x: 10.2642, title: 'Location', lines: ['Lampung,', 'Indonesian, Asia'] },
];

function slide22(s) {
  // envelope
  s.addShape('rect', { x: 5.1347, y: 4.8913, w: 0.9332, h: 0.678, fill: { type: 'none' }, line: { color: YELLOW, width: 2.25 } });
  s.addShape('line', { x: 5.1347, y: 4.8913, w: 0.4666, h: 0.339, line: { color: YELLOW, width: 2.25 } });
  s.addShape('line', { x: 5.6013, y: 4.8913, w: 0.4666, h: 0.339, line: { color: YELLOW, width: 2.25 }, flipH: true });
  // globe: outer circle, meridian and two latitude lines
  s.addShape('ellipse', { x: 8.0825, y: 4.8403, w: 0.8102, h: 0.8102, fill: { type: 'none' }, line: { color: YELLOW, width: 2.25 } });
  s.addShape('ellipse', { x: 8.2825, y: 4.8403, w: 0.4102, h: 0.8102, fill: { type: 'none' }, line: { color: YELLOW, width: 2.25 } });
  [5.0403, 5.2454, 5.4505].forEach((y) => s.addShape('line', {
    x: 8.0825 + (y === 5.2454 ? 0 : 0.09), y: y, w: 0.8102 - (y === 5.2454 ? 0 : 0.18), h: 0,
    line: { color: YELLOW, width: 2.25 },
  }));
  // map pin: folded map panels with a marker on top
  s.addShape('rect', { x: 10.9574, y: 4.9402, w: 0.9691, h: 0.7253, fill: { type: 'none' }, line: { color: YELLOW, width: 2.25 } });
  [11.28, 11.604].forEach((x) => s.addShape('line', { x: x, y: 4.9402, w: 0, h: 0.7253, line: { color: YELLOW, width: 2.25 } }));
  s.addShape('teardrop', { x: 11.2168, y: 4.6995, w: 0.4439, h: 0.4439, fill: { color: WHITE }, line: { color: YELLOW, width: 2.25 }, rotate: 135 });
  s.addShape('ellipse', { x: 11.3294, y: 4.8058, w: 0.2188, h: 0.2188, fill: { type: 'none' }, line: { color: YELLOW, width: 2 } });
  eyebrow(s, 0.4608, 4.7365);
  text(s, 'Contact Us', { x: 0.4608, y: 5.1095, w: 3.3861, h: 0.7068, size: 36, bold: true });
  para(s, 'Lorem ipsum dolor sit amet, consectetw adipiscing elit, sed eiusm od tempor inci ' +
    'didunt ut labore et dolore magna.', { x: 0.4608, y: 5.854, w: 3.258, h: 0.8665 });
  CONTACTS.forEach((col) => {
    text(s, col.title, { x: col.x, y: 5.7486, w: 2.3555, h: 0.3366, size: 14, bold: true, align: 'center' });
    text(s, [{ text: col.lines[0], options: { breakLine: true } }, { text: col.lines[1] }],
      { x: col.x + 0.0445, y: 6.0886, w: 2.2666, h: 0.6014, size: 10.5, color: INK_SOFT, align: 'center', lnSpc: 1.5 });
  });
}

// 23 — "Icons Slides" divider on dark.
function slide23(s) {
  rect(s, { x: 0, y: 0, w: 13.3333, h: 7.5, fill: BLACK });
  text(s, [{ text: 'Icons ', options: { color: WHITE } }, { text: 'Slides', options: { color: YELLOW } }],
    { x: 2.9356, y: 2.7667, w: 7.462, h: 1.4473, size: 80, bold: true, align: 'center' });
  nav(s, { x: 9.2426, y: 0.3827, w: 3.7982, items: NAV_3, color: WHITE });
  logo(s, 0.5133, 0.295, WHITE);
  footerDark(s);
}

// 24-27 — icon library pages. The originals are hundreds of hand-drawn glyph
// paths; here every cell is a preset shape on the same grid, at the same cell
// size and in the same near-black. Pages 24-25 mix filled and line glyphs like
// the reference; pages 26-27 are solid throughout.
const ICON_SHAPES = ['ellipse', 'star5', 'heart', 'sun', 'moon', 'cloud', 'lightningBolt',
  'smileyFace', 'diamond', 'hexagon', 'pentagon', 'plus', 'donut', 'triangle', 'star4', 'chevron'];

function iconGrid(s, o) {
  o.rows.forEach((cy, r) => o.cols.forEach((cx, c) => {
    const outline = o.mixed && (r + c) % 2 === 1;
    s.addShape(ICON_SHAPES[(r * o.cols.length + c) % ICON_SHAPES.length], {
      x: cx - o.size / 2, y: cy - o.size / 2, w: o.size, h: o.size,
      fill: outline ? { type: 'none' } : { color: BLACK },
      line: { color: BLACK, width: outline ? 2 : 0.75 },
    });
  }));
}

function slide24(s) {
  iconGrid(s, { mixed: true, size: 0.38,
    cols: [0.749, 1.724, 2.768, 3.857, 4.941, 6.066, 7.159, 8.291, 9.368, 10.345, 11.427, 12.559],
    rows: [0.632, 1.511, 2.427, 3.358, 4.297, 5.235, 6.07, 6.842] });
}

function slide25(s) {
  iconGrid(s, { mixed: true, size: 0.375,
    cols: [0.58, 1.606, 2.707, 3.879, 5.037, 6.149, 7.259, 8.429, 9.577, 10.639, 11.702, 12.723],
    rows: [0.626, 1.548, 2.391, 3.279, 4.172, 5.012, 5.924, 6.783] });
}

function slide26(s) {
  iconGrid(s, { size: 0.39,
    cols: [1.027, 2.21, 3.315, 4.438, 5.608, 6.75, 7.768, 8.8, 9.849, 10.976, 12.208],
    rows: [0.799, 1.827, 2.955, 4.106, 5.246, 6.321] });
}

function slide27(s) {
  iconGrid(s, { size: 0.4,
    cols: [1.144, 2.428, 3.575, 4.634, 5.712, 6.884, 7.984, 9.075, 10.059, 11.108, 12.182],
    rows: [0.986, 2.041, 3.162, 4.288, 5.341, 6.347] });
}

// 28 — "Thanks" closing card, dark band like slide 1.
function slide28(s) {
  rect(s, { x: 0, y: 0.807, w: 13.3333, h: 5.42, fill: BLACK });
  text(s, [{ text: 'Tha', options: { color: WHITE } }, { text: 'nks', options: { color: YELLOW } }],
    { x: 0.8494, y: 1.6935, w: 5.086, h: 1.7166, size: 96, bold: true, wrap: false });
  text(s, 'Presentation Template.', { x: 0.8494, y: 3.2586, w: 2.3065, h: 0.3029, size: 12, color: WHITE });
  nav(s, { x: 7.7222, y: 0.2194, w: 5.26, items: NAV_4, color: INK });
  rule(s, 0.8948, 4.7342, 4.8818, WHITE);
  text(s, CLUB, { x: 0.8494, y: 4.2798, w: 2.3065, h: 0.3029, face: SEMI, size: 12, color: YELLOW });
  text(s, LOREM_TINY, { x: 0.8494, y: 4.8857, w: 2.5084, h: 0.6253, size: 11, color: WHITE, align: 'justify', lnSpc: 1.5 });
  logo(s, 0.5133, 0.1318, INK);
}

// 29 — "Thanks You" over a dimmed photo.
function slide29(s) {
  rect(s, { x: 0, y: 0, w: 13.3333, h: 7.5, fill: BLACK, transparency: 30 });
  text(s, [{ text: 'Thanks ', options: { color: WHITE } }, { text: 'You', options: { color: YELLOW } }],
    { x: 2.6686, y: 2.632, w: 7.9961, h: 1.7166, size: 96, bold: true, wrap: false });
  nav(s, { x: 9.2426, y: 0.3827, w: 3.7982, items: NAV_3, color: WHITE });
  logo(s, 0.5133, 0.295, WHITE);
  footerDark(s);
}

// 30 — the same closing card on flat dark.
function slide30(s) {
  rect(s, { x: 0, y: 0, w: 13.3333, h: 7.5, fill: BLACK });
  text(s, [{ text: 'Thanks ', options: { color: WHITE } }, { text: 'You', options: { color: YELLOW } }],
    { x: 2.6686, y: 2.632, w: 7.9961, h: 1.7166, size: 96, bold: true, wrap: false });
  nav(s, { x: 9.2426, y: 0.3827, w: 3.7982, items: NAV_3, color: WHITE });
  logo(s, 0.5133, 0.295, WHITE);
  footerDark(s);
}

/* -------------------------------------------------------------------- main */

const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
  slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18,
  slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28,
  slide29, slide30];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'JENGKY', width: 13.3333, height: 7.5 });
  pptx.layout = 'JENGKY';
  pptx.title = 'Jengky Presentation Template';
  SLIDES.forEach((builder) => builder(pptx.addSlide()));
  return pptx.writeFile({ fileName: path.join(__dirname, '07c58c37-7d3c-4d8c-a3f5-e83bc8197bee_grok_final.pptx') });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
