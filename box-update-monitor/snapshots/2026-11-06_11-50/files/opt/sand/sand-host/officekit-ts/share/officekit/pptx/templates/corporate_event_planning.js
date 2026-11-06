/**
 * "Event Plan" deck rebuilt with pptxgenjs.
 *   node 0eab9463-a6d7-453c-9557-e4b800c8a7d8_grok_final.js
 * writes 0eab9463-a6d7-453c-9557-e4b800c8a7d8_grok_final.pptx next to this file.
 *
 * The raster mock-ups of the original (laptop, phones, jigsaw, world map,
 * icons) are redrawn from native pptxgenjs shapes; nothing is embedded.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const ORANGE = 'F88A44'; // accent1
const DARK = '262626'; // accent2
const YELLOW = 'F8CF28'; // accent3
const WHITE = 'FFFFFF';
const BLACK = '000000';
const GREY_TEXT = '404040'; // tx1 @ 75% luminance
const GREY_LINE = 'BFBFBF'; // bg1 @ 75% luminance
const TRACK = 'F2F2F2'; // progress-bar trough
const MAP_GREY = 'F1F1F1';

const HEAD = 'Space Grotesk'; // +mj-lt
const BODY = 'Inter'; // +mn-lt

const W = 13.333;
const H = 7.5;

const HAIRLINE = { color: DARK, width: 0.5 };
const NO_LINE = { type: 'none' };
const INSETS = [7.2, 7.2, 3.6, 3.6]; // PowerPoint "normal" text insets, in points

/* Blend two hex colours; t = 0 -> a, t = 1 -> b. */
function mix(a, b, t) {
  return [0, 2, 4]
    .map((i) => {
      const v = Math.round(parseInt(a.substr(i, 2), 16) * (1 - t) + parseInt(b.substr(i, 2), 16) * t);
      return v.toString(16).padStart(2, '0').toUpperCase();
    })
    .join('');
}

/* Every gradient in the deck is the same 45-degree ramp: yellow in the
   top-left corner sliding into flat orange towards the bottom-right. The
   only thing that varies is `stop`, the percentage at which the ramp has
   fully reached orange.                                                  */
const rampColor = (stop, f) => mix(ORANGE, YELLOW, f);
const rampAverage = (stop) => mix(ORANGE, YELLOW, 0.45);

/* ------------------------------------------------------- drawing helpers */

/**
 * Region of a w x h box on the top-left side of the 45-degree line
 * (dx + dy = d), as a pptxgenjs custGeom point list.
 */
function cornerRegion(w, h, d) {
  const pts = [{ x: 0, y: 0 }];
  if (d > w) pts.push({ x: w, y: 0 });
  pts.push({ x: Math.min(d, w), y: Math.max(0, d - w) }); // meets top or right edge
  pts.push({ x: Math.max(0, d - h), y: Math.min(d, h) }); // meets left or bottom edge
  if (d > h) pts.push({ x: 0, y: h });
  pts.push({ close: true });
  return pts;
}

/**
 * pptxgenjs cannot emit gradient fills, so the ramp is painted as a stack of
 * 45-degree bands: an orange base plate with progressively smaller, more
 * yellow corner regions layered on top.
 */
function gradientRect(s, x, y, w, h, stop, line) {
  const span = w + h;
  const ramp = span * (1 - stop / 100);
  const bands = Math.min(32, Math.max(8, Math.round(ramp * 4)));
  s.addShape('rect', { x, y, w, h, fill: { color: ORANGE }, line: NO_LINE });
  for (let i = bands; i >= 1; i--) {
    s.addShape('custGeom', {
      x, y, w, h,
      fill: { color: rampColor(stop, 1 - (i - 0.5) / bands) },
      line: NO_LINE,
      points: cornerRegion(w, h, (ramp * i) / bands),
    });
  }
  if (line !== 'none') s.addShape('rect', { x, y, w, h, line: line || HAIRLINE });
}

/** Solid rectangle carrying the deck's hairline outline. */
function panel(s, x, y, w, h, color) {
  s.addShape('rect', { x, y, w, h, fill: { color }, line: HAIRLINE });
}

/** Thin rule; `w` or `h` is 0 for a horizontal / vertical line. */
function rule(s, x, y, w, h) {
  s.addShape('line', { x, y, w, h, line: HAIRLINE });
}

/** Text box. `o` mirrors the source run properties. */
function txt(s, body, x, y, w, h, o) {
  o = o || {};
  s.addText(body, {
    x, y, w, h,
    fontFace: o.face || BODY,
    fontSize: o.size || 12,
    color: o.color || DARK,
    align: o.align || 'left',
    valign: o.valign || 'top',
    lineSpacingMultiple: o.lnSpc,
    transparency: o.transparency,
    margin: INSETS,
    isTextBox: true,
  });
}

/** Two-tone slide heading: orange lead-in, then the dark remainder. */
function heading(s, lead, rest, x, y, w, h, size, align) {
  s.addText(
    [
      { text: lead, options: { color: ORANGE } },
      { text: rest, options: { color: DARK } },
    ],
    {
      x, y, w, h,
      fontFace: HEAD, fontSize: size || 60,
      align: align || 'left', valign: 'top',
      margin: INSETS, isTextBox: true,
    }
  );
}

/** Big figure ("76,8%") plus optional 16pt caption and 12pt paragraph. */
function statBlock(s, x, y, stat, caption, para, o) {
  o = o || {};
  const c = o.color || DARK;
  if (stat) txt(s, stat, x, y, o.statW || 1.37, 0.505, { face: HEAD, size: 24, color: c });
  if (caption) txt(s, caption, x, o.capY, o.capW || 3.33, 0.37, { face: HEAD, size: 16, color: c });
  if (para) txt(s, para, x, o.paraY, o.paraW, o.paraH, { size: 12, color: c, lnSpc: 1.3 });
}

/** Grey trough, dark fill and a right-hand percentage caption. */
function progressBar(s, x, y, w, h, done, labelX, labelY, label) {
  s.addShape('rect', { x, y, w, h, fill: { color: TRACK, transparency: 12 }, line: NO_LINE });
  s.addShape('rect', { x, y, w: done, h, fill: { color: DARK }, line: NO_LINE });
  txt(s, label, labelX, labelY, 0.677, 0.358, { face: HEAD, size: 14, color: GREY_TEXT, lnSpc: 1.2 });
}

/** Numbered card: ghosted 48pt number, 16pt title, 12pt paragraph. */
function numberCard(s, x, y, o) {
  const fg = o.onColour ? WHITE : DARK;
  s.addText(o.num, {
    x: x + (o.numDx || 0), y, w: o.numW || 2.134, h: 0.909,
    fontFace: HEAD, fontSize: 48, color: o.onColour ? WHITE : o.numColor || DARK,
    transparency: 73, align: 'right', valign: 'top', margin: INSETS, isTextBox: true,
  });
  txt(s, o.title, x + (o.textDx || 0), o.titleY, o.titleW, 0.64, { face: HEAD, size: 16, color: fg });
  txt(s, o.body, x + (o.textDx || 0), o.bodyY, o.bodyW, o.bodyH, { size: 12, color: fg, lnSpc: 1.3 });
}

/** Master furniture: page number and deck name, top-left. */
function slideChrome(s, n) {
  txt(s, String(n), 0.235, 0.141, 0.39, 0.286, { face: HEAD, size: 11 });
  txt(s, '- Event Plan ', 0.454, 0.141, 1.256, 0.286, { face: HEAD, size: 11 });
}

/* --------------------------------------------------------- shared strings */

const L = {
  lorem1: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed iaculis vel neque ac cursus. Donec aliquam consectetur adipiscing elit. Sed iaculis vel neque sapien.',
  lorem2: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed iaculis vel neque ac cursus. Donec aliquam sapien a nisi varius pellentesque',
  long: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin bibendum risus sit amet venenatis tristique. Morbi tortor justo, porta eget conse quat ac, semper sit amet sem. Proin porta odio arcu, sit amet finibus urna semper vitae. Mauris ipsum est, scelerisque sed nunc.',
  med: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin bibendum risus sit amet venenatis tristique. Morbi tortor justo.',
  card: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin bibendum risus sit amet venenatis tristique. Morbi tortor justo, porta eget.',
  short: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin bibendum risus.',
  shortSit: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin bibendum risus sit.',
  tiny: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. ',
};

/* ================================================================= slides */

function slide01(s) {
  gradientRect(s, 0, 4.625, W, 2.875, 29);
  s.addText(
    [
      { text: 'EVENT', options: { color: ORANGE } },
      { text: ' PLAN ', options: { color: DARK } },
    ],
    { x: 4.612, y: 1.213, w: 7.97, h: 1.561, fontFace: HEAD, fontSize: 96,
      lineSpacingMultiple: 0.9, valign: 'top', margin: INSETS, isTextBox: true }
  );
  txt(s, L.lorem1, 4.697, 2.774, 7.801, 0.601, { color: GREY_TEXT, lnSpc: 1.3 });
  txt(s, '01', 9.25, 5.137, 3.55, 2.036, { face: HEAD, size: 115, color: WHITE, align: 'right', transparency: 65 });
  txt(s, 'COMPANY NAME ', 4.847, 6.529, 2.18, 0.303, { face: HEAD, color: WHITE });
}

function slide02(s) {
  rule(s, 0, 6.875, W, 0);
  gradientRect(s, 0, 0.625, W, 4.038, 29);
  txt(s, 'WE CAN CREATE A WORTHY AND APPROPRIATE EVENT', 5.477, 1.022, 7.19, 3.376,
    { face: HEAD, size: 54, color: WHITE, lnSpc: 0.9 });
  statBlock(s, 5.649, 5.5, '65%', null, L.lorem2, { paraY: 6.122, paraW: 6.907, paraH: 0.604 });
}

function slide03(s) {
  heading(s, 'EVENT PLAN', ' CONTENT DESCRIPTION ', 0.502, 0.832, 10.863, 2.121);
  txt(s, L.long, 0.588, 4.346, 5.883, 1.126, { lnSpc: 1.3 });
  txt(s, L.med, 0.588, 5.827, 5.883, 0.601, { lnSpc: 1.3 });
  txt(s, 'EVENT PLAN PRESENTATION TEMPLATE ', 0.588, 3.765, 4.935, 0.404, { face: HEAD, size: 18 });
  rule(s, 0, 3.364, W, 0);
  gradientRect(s, 7.002, 3.364, 2.929, 4.149, 29);
  statBlock(s, 7.231, 4.168, '2388+', null, L.card,
    { color: WHITE, statW: 1.718, paraY: 5.037, paraW: 2.621, paraH: 1.388 });
  statBlock(s, 10.294, 4.168, '$3778', null, L.card,
    { statW: 1.718, paraY: 5.037, paraW: 2.621, paraH: 1.388 });
}

function slide04(s) {
  rule(s, 0, 3.001, W, 0);
  heading(s, 'PROJECT ', 'EVENT PLANNING ', 0.588, 0.953, 11.143, 1.111);
  const cards = [
    { x: 4.474, num: '01', title: 'EVENT PLANNING ORGANIZING', titleW: 2.134, onColour: true },
    { x: 7.42, num: '02', title: 'MARKET EVENT CONTROLLING ', titleW: 2.134, onColour: true },
    { x: 10.497, num: '03', title: 'CUSTOMER POPULAR REVIEW OPINION', titleW: 2.453, onColour: false },
  ];
  cards.forEach((c) => {
    if (c.onColour) gradientRect(s, c.x, 3.001, 2.946, 4.499, 48);
    numberCard(s, c.x, 3.105, {
      num: c.num, onColour: c.onColour,
      numDx: c.onColour ? 0.71 : 0.547, textDx: c.onColour ? 0.162 : 0,
      title: c.title, titleY: 4.51, titleW: c.titleW,
      body: L.card, bodyY: 5.232, bodyW: 2.621, bodyH: 1.388,
    });
  });
}

function slide05(s) {
  rule(s, 0, 4.208, W, 0);
  gradientRect(s, 3.936, 1.36, 2.813, 4.769, 43);
  s.addText('”', { x: 4.05, y: 1.34, w: 1.2, h: 1.2, fontFace: 'Arial', fontSize: 120,
    color: WHITE, transparency: 25, valign: 'top', margin: 0, isTextBox: true });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adip iscing elit. In vehicula sem sit amet hendrerit dictum. Donec',
    4.126, 3.095, 2.429, 1.603, { size: 14, color: WHITE, lnSpc: 1.3 });
  txt(s, '- COMPANY NAME ', 4.126, 5.401, 1.778, 0.303, { face: HEAD, color: WHITE });
  heading(s, 'PROJECT ', 'EVENT PLAN  ', 7.192, 1.428, 5.53, 2.121);
  statBlock(s, 7.186, 4.766, '76,8%', null,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin biben dum risus sit amet venenatis tristique. Morbi tortor justo.',
    { paraY: 5.523, paraW: 5.53, paraH: 0.601 });
}

function slide06(s) {
  rule(s, 0, 2.682, W, 0);
  heading(s, 'EVENT ANNUAL ', 'RECAP ', 0.522, 1.057, 11.876, 1.111);
  ['2023', '2024', '2025', '2026'].forEach((year, i) => {
    const y = 3.363 + i * 0.946;
    panel(s, 0.667, y, 1.516, 0.626, ORANGE);
    txt(s, year, 0.773, y + 0.111, 1.304, 0.404, { face: HEAD, size: 18, color: WHITE, align: 'center' });
    txt(s, L.card, 2.329, y + 0.013, 6.016, 0.601, { lnSpc: 1.3 });
  });
  rule(s, 8.83, 2.682, 0, 4.818);
  statBlock(s, 9.159, 4.055, '76,8%', 'EVENT PLAN PRESENTATION ',
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin bibendum risus sit amet venenatis tristique. Morbi tortor justo, porta eget conse quat ac, semper sit amet sem. Proin porta odio arcu.',
    { capY: 4.865, capW: 3.327, paraY: 5.372, paraW: 3.495, paraH: 1.388 });
}

function slide07(s) {
  gradientRect(s, 0, 1.365, 5.111, 4.769, 46);
  ['01.', '02.', '03.', '04.', '05.'].forEach((n, i) => {
    const y = 1.761 + i * 0.8455;
    txt(s, n, 0.373, y + 0.048, 0.715, 0.505, { face: HEAD, size: 24, color: WHITE });
    txt(s, L.shortSit, 1.044, y, 3.694, 0.601, { color: WHITE, lnSpc: 1.3 });
  });
  heading(s, 'PROJECT EVENT ', 'CONTENT PLAN ', 5.778, 1.366, 6.879, 2.121);
  rule(s, 5.111, 4.053, 8.222, 0);
  statBlock(s, 5.778, 4.393, '3288+', null,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin bibendum risus sit amet venenatis tristique. Morbi tortor justo, porta eget conse quat ac, semper sit amet sem. Proin porta odio arcu, sit amet finibus urna semper vitae. Mauris ipsum est.',
    { paraY: 5.143, paraW: 7.183, paraH: 0.863 });
}

function slide08(s) {
  rule(s, 0, 3.435, 9.472, 0);
  heading(s, 'MARKETING ', 'PLAN CONTENT', 0.667, 0.792, 6.505, 2.121);
  const cards = [
    { x: 0.667, num: '01', title: 'CONTENT DESCRIPTION ', titleW: 2.118, onColour: true },
    { x: 3.59, num: '02', title: 'CONTENT MARKET MAPPING  ', titleW: 2.142, onColour: true },
    { x: 6.686, num: '03', title: 'PLAN EXECUTION FOR PUBLIC MARKET ', titleW: 2.444, onColour: false },
  ];
  cards.forEach((c) => {
    if (c.onColour) gradientRect(s, c.x, 3.435, 2.947, 4.065, 46);
    numberCard(s, c.x, 3.583, {
      num: c.num, onColour: c.onColour, numColor: ORANGE, numW: 2.259,
      numDx: c.onColour ? 0.545 : 0.374, textDx: c.onColour ? 0.172 : 0,
      title: c.title, titleY: 4.711, titleW: c.titleW,
      body: L.short, bodyY: 5.43, bodyW: 2.602, bodyH: 0.863,
    });
  });
}

function slide09(s) {
  heading(s, 'EVENT PLAN ', 'QUALITY CONTROL', 0.743, 0.86, 10.698, 2.121);
  rule(s, 0, 3.389, W, 0);
  [{ x: 3.584, num: '01' }, { x: 9.222, num: '02' }].forEach((c) => {
    gradientRect(s, c.x, 3.39, 2.946, 4.119, 51);
    numberCard(s, c.x, 3.594, {
      num: c.num, onColour: true, numDx: 0.71, textDx: 0.163,
      title: 'MARKET EVENT CONTROLLING ', titleY: 4.941, titleW: 2.134,
      body: L.short, bodyY: 5.73, bodyW: 2.621, bodyH: 0.863,
    });
  });
}

function slide10(s) {
  gradientRect(s, 5.011, 0.625, 8.307, 6.25, 52);
  heading(s, 'EVENT ', 'PLANNING PROGRESS', 0.563, 0.86, 4.268, 2.827, 54);
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin bibendum risus sit amet venenatis tristique.',
    0.557, 5.348, 3.625, 0.863, { lnSpc: 1.3 });
  rule(s, 0, 4.298, 5.011, 0);
  [{ x: 5.522, num: '01.', tw: 3.089 }, { x: 9.812, num: '02.', tw: 2.771 }].forEach((c) => {
    txt(s, c.num, c.x, 1.147, 2.259, 0.505, { face: HEAD, size: 24, color: WHITE });
    txt(s, 'EVENT PLAN PROGRESS', c.x, 1.804, c.tw, 0.37, { face: HEAD, size: 16, color: WHITE });
    txt(s, 'Lorem ipsum dolor sit amet, consec tetur adipiscing elit. Proin bibendum adipiscing elit. Proin risus.',
      c.x, 2.214, 3.169, 0.863, { color: WHITE, lnSpc: 1.3 });
  });
  txt(s, 'EVENT PLAN PRESENTATION ', 0.563, 4.844, 3.39, 0.37, { face: HEAD, size: 16 });
}

function slide11(s) {
  gradientRect(s, 5.72, 3.054, 7.613, 4.446, 42);
  txt(s, 'PORTFOLIO EVENT GALLERY', 8.362, 3.546, 4.535, 3.13, { face: HEAD, size: 60, color: WHITE });
  txt(s, 'EVENT PLAN PORTFOLIO IMAGE GALLERY ', 8.362, 1.035, 4.455, 0.37, { face: HEAD, size: 16 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipi scing elit. Proin bibendum risus sit.',
    8.362, 1.59, 3.971, 0.601, { lnSpc: 1.3 });
}

function slide12(s) {
  rule(s, 0, 3.612, W, 0);
  gradientRect(s, 10.078, 0, 3.256, 7.5, 43);
  heading(s, 'MEET OUR ', 'GREAT MANAGER ', 0.882, 1.034, 6.946, 2.121);
  txt(s, 'EVENT PERCENTAGE DATA ', 0.988, 3.992, 3.923, 0.405, { face: HEAD, size: 16, color: GREY_TEXT, lnSpc: 1.2 });
  progressBar(s, 0.988, 4.596, 5.585, 0.079, 4.05, 6.626, 4.457, '85%');
  s.addText(
    [
      { text: '15+ ', options: { fontSize: 24 } },
      { text: 'YEARS EXPERIENCE', options: { fontSize: 16 } },
    ],
    { x: 0.882, y: 4.966, w: 3.043, h: 0.505, fontFace: HEAD, color: DARK,
      valign: 'top', margin: INSETS, isTextBox: true }
  );
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin bibendum risus sit amet venenatis tristique. Morbi tortor justo, porta eget semper sit amet sem. ',
    0.882, 5.63, 7.183, 0.601, { lnSpc: 1.3 });
}

function slide13(s) {
  const team = [
    { x: 0.667, name: 'LUIZ SUAREZ ', nameW: 1.811, onColour: true },
    { x: 3.37, name: 'ELIZABETH LIU', nameW: 1.811, onColour: false },
    { x: 5.923, name: 'EDUARDO GIMENEZ', nameW: 2.377, onColour: true },
  ];
  team.forEach((m) => {
    if (m.onColour) gradientRect(s, m.x, 5.158, 2.628, 1.532, 29);
    const dx = m.onColour ? 0.075 : 0;
    txt(s, m.name, m.x + dx, 5.362, m.nameW, 0.37,
      { face: HEAD, size: 16, color: m.onColour ? WHITE : BLACK });
    txt(s, L.tiny, m.x + dx, 5.884, 2.52, 0.601,
      { color: m.onColour ? WHITE : DARK, lnSpc: 1.3 });
  });
  rule(s, 0, 6.689, W, 0);
  heading(s, 'MEET OUR ', 'TEAM ', 8.882, 0.812, 3.387, 3.13);
  rule(s, 8.551, 4.105, 4.782, 0);
  s.addText(
    [
      { text: '10+ ', options: { fontSize: 24 } },
      { text: 'YEARS EXPERIENCE', options: { fontSize: 16 } },
    ],
    { x: 9.0, y: 5.131, w: 3.269, h: 0.505, fontFace: HEAD, color: DARK,
      valign: 'top', margin: INSETS, isTextBox: true }
  );
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin Bibendum.',
    9.0, 5.764, 3.76, 0.601, { lnSpc: 1.3 });
}

/**
 * Open-lid laptop mock-up. Only the frame is drawn: the screen is left open
 * so whatever sits behind it on the slide shows through, as in the original.
 */
function laptop(s, x, y) {
  const lidX = x + 1.478, lidW = 5.929, lidH = 3.963;
  const bezel = 0.203, topBar = 0.241, chin = 0.278;
  const shell = { color: '111213' };
  const frame = (bx, by, bw, bh) => s.addShape('rect', { x: bx, y: by, w: bw, h: bh, fill: shell, line: NO_LINE });
  frame(lidX, y, lidW, topBar);
  frame(lidX, y + lidH - chin, lidW, chin);
  frame(lidX, y + topBar, bezel, lidH - topBar - chin);
  frame(lidX + lidW - bezel, y + topBar, bezel, lidH - topBar - chin);
  s.addShape('ellipse', { x: lidX + lidW / 2 - 0.03, y: y + 0.105, w: 0.06, h: 0.06,
    fill: { color: '3A3D40' }, line: NO_LINE });
  // Aluminium base: a shallow trapezoid with a lighter top edge and a thumb notch.
  const bY = y + lidH, bH = 0.212;
  s.addShape('custGeom', { x: x + 0.808, y: bY, w: 7.259, h: bH, fill: { color: 'B9BCC2' }, line: NO_LINE,
    points: [{ x: 0, y: 0 }, { x: 7.259, y: 0 }, { x: 6.916, y: bH }, { x: 0.156, y: bH }, { close: true }] });
  s.addShape('rect', { x: x + 0.808, y: bY, w: 7.259, h: 0.055, fill: { color: 'E4E5E8' }, line: NO_LINE });
  s.addShape('roundRect', { x: x + 4.068, y: bY - 0.03, w: 0.863, h: 0.085,
    fill: { color: 'DDDEE1' }, line: NO_LINE, rectRadius: 0.03 });
}

function slide14(s) {
  gradientRect(s, 0, 0, 3.256, 7.5, 45);
  rule(s, 6.811, 3.847, 6.522, 0);
  heading(s, 'EVENT PLAN ', 'MONITORING  ', 7.864, 1.274, 5.56, 2.121);
  laptop(s, -0.56, 1.578);
  statBlock(s, 7.855, 4.119, '3599+', null,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin bibendum risus sit amet venenatis tristique. ',
    { paraY: 4.689, paraW: 4.751, paraH: 0.601 });
  txt(s, '“Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin risus sit amet tristique.” ',
    7.855, 5.466, 4.751, 0.691, { face: HEAD, size: 14, color: ORANGE, lnSpc: 1.3 });
}

/** Phone mock-up: three nested rounded shells, notch island and side buttons. */
function phone(s, x, y, w, h) {
  s.addShape('roundRect', { x, y, w, h, fill: { color: '5E465D' }, line: NO_LINE, rectRadius: 0.42 });
  s.addShape('roundRect', { x: x + 0.012, y: y + 0.011, w: w - 0.024, h: h - 0.024,
    fill: { color: '706F6F' }, line: NO_LINE, rectRadius: 0.41 });
  s.addShape('roundRect', { x: x + 0.042, y: y + 0.039, w: w - 0.084, h: h - 0.078,
    fill: { color: '1D1D1B' }, line: NO_LINE, rectRadius: 0.38 });
  s.addShape('roundRect', { x: x + w / 2 - 0.28, y: y + 0.195, w: 0.561, h: 0.162,
    fill: { color: '141413' }, line: NO_LINE, rectRadius: 0.081 });
  s.addShape('ellipse', { x: x + w / 2 + 0.16, y: y + 0.242, w: 0.071, h: 0.071,
    fill: { color: '3A3A3A' }, line: NO_LINE });
  [[0.808, 0.201], [1.177, 0.391], [1.668, 0.391]].forEach(([dy, bh]) =>
    s.addShape('rect', { x: x - 0.014, y: y + dy, w: 0.021, h: bh, fill: { color: '68595F' }, line: NO_LINE }));
  s.addShape('rect', { x: x + w - 0.007, y: y + 1.306, w: 0.021, h: 0.617,
    fill: { color: '68595F' }, line: NO_LINE });
}

function slide15(s) {
  gradientRect(s, 0, 3.75, W, 2.543, 49);
  phone(s, 7.044, 0.965, 2.536, 5.115);
  phone(s, 9.867, 1.432, 2.536, 5.115);
  heading(s, 'EVENT PLAN ', 'TASK JOURNAL ', 0.642, 1.081, 6.188, 2.121);
  [{ x: 0.849, num: '01.', title: 'EVENT PLAN TASK', tw: 2.184 },
   { x: 3.956, num: '02.', title: 'EVENT PLAN JOURNAL', tw: 2.599 }].forEach((c) => {
    txt(s, c.num, c.x, 4.056, 2.259, 0.505, { face: HEAD, size: 24, color: WHITE });
    txt(s, c.title, c.x, 4.713, c.tw, 0.37, { face: HEAD, size: 16, color: WHITE });
    txt(s, 'Lorem ipsum dolor sit amet, consec etur adipiscing elit. Proin bibe proin risus.',
      c.x, 5.123, 2.493, 0.863, { color: WHITE, lnSpc: 1.3 });
  });
}

const TABLE_COLS = [2.384, 1.43, 1.092, 0.935, 1.078, 0.895];
const TABLE_HEAD = ['EVENT NAME ', 'EVENT DATE', 'AUDIENCE ', 'PRICE', 'TYPE', 'CHECK '];
const TABLE_ROWS = [
  ['Event Plan #1', '12 Jan 2025', '3588+', '$ 460', 'Type A', ''],
  ['Event Plan #2', '24 Feb 2025', '2388+', '$ 210', 'Type E', ''],
  ['Event Plan #3', '18 Mar 2025', '4288+', '$ 360', 'Type D', ''],
  ['Event Plan #4', '12 Apr 2025', '3289+ ', '$ 450', 'Type C', ''],
  ['Event Plan #5', '30 Jun 2025', '2399+ ', '$ 270', 'Type B ', ''],
];

function slide16(s) {
  heading(s, 'PROJECT EVENT ', 'DATA TABLE', 0.522, 1.057, 11.596, 1.111);
  rule(s, 0, 2.871, W, 0);
  rule(s, 0, 6.875, W, 0);
  rule(s, 0.667, 2.869, 0, 4.006);
  rule(s, 8.481, 2.869, 0, 4.006);

  // Each header cell owns its own gradient, so paint them behind the table.
  let hx = 0.667;
  TABLE_COLS.forEach((cw) => {
    gradientRect(s, hx, 2.871, cw, 0.667, 49, 'none');
    hx += cw;
  });

  const rows = [
    TABLE_HEAD.map((t) => ({ text: t, options: { fill: { type: 'none' }, color: WHITE, fontFace: HEAD } })),
    ...TABLE_ROWS.map((r) => r.map((t) => ({ text: t, options: { fill: { color: WHITE }, color: GREY_TEXT, fontFace: BODY } }))),
  ];
  s.addTable(rows, {
    x: 0.667, y: 2.871, w: 7.814, colW: TABLE_COLS, rowH: 0.667, fontSize: 12,
    border: { type: 'solid', color: GREY_LINE, pt: 0.75 }, valign: 'middle',
    margin: [3.6, 7.2, 3.6, 7.2],
  });

  // Tick badges in the CHECK column.
  for (let i = 0; i < 5; i++) {
    const cy = 3.782 + i * 0.648;
    s.addShape('ellipse', { x: 7.949, y: cy, w: 0.21, h: 0.21, fill: { color: ORANGE }, line: HAIRLINE });
    s.addShape('custGeom', { x: 7.949, y: cy, w: 0.21, h: 0.21, line: { color: WHITE, width: 1 },
      points: [{ x: 0.055, y: 0.105 }, { x: 0.093, y: 0.145 }, { x: 0.155, y: 0.068 }] });
  }

  statBlock(s, 8.845, 3.74, '3288+', 'EVENT PLAN DATA PRESENTATION ',
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin bibendum risus sit amet venenatis tristique. Morbi tortor justo. ',
    { capY: 4.998, capW: 3.822, paraY: 5.571, paraW: 3.822, paraH: 0.863 });
  progressBar(s, 8.946, 4.6, 2.843, 0.07, 2.062, 12.038, 4.456, '75%');
}

function slide17(s) {
  gradientRect(s, 0, 1.167, 5.611, 5.306, 46);
  const cats = ['Data 1', 'Data 2', 'Data 3', 'Data 4'];
  s.addChart('bar', [
    { name: 'Series 1', labels: cats, values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels: cats, values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels: cats, values: [2, 2, 3, 5] },
  ], {
    x: 0.32, y: 1.52, w: 4.971, h: 4.6,
    barDir: 'bar', barGrouping: 'clustered', barGapWidthPct: 182,
    chartColors: [WHITE, WHITE, WHITE], chartColorsOpacity: 50,
    showLegend: false, showTitle: false,
    catAxisLabelColor: WHITE, catAxisLabelFontFace: HEAD, catAxisLabelFontSize: 12,
    catAxisLineColor: 'E4E4E4',
    valAxisLabelColor: WHITE, valAxisLabelFontFace: BODY, valAxisLabelFontSize: 12,
    valAxisLineShow: false,
    valGridLine: { color: 'FBC08D', style: 'solid', size: 0.75 },
    catGridLine: { style: 'none' },
    plotArea: { fill: { color: WHITE, transparency: 100 } },
    chartArea: { fill: { color: WHITE, transparency: 100 } },
  });
  heading(s, 'PROJECT EVENT ', 'OFFICE CHART', 6.295, 1.389, 6.533, 2.121);
  rule(s, 5.611, 3.915, 7.722, 0);
  statBlock(s, 6.295, 4.356, '$4500+', null,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin bibendum risus sit amet venenatis tristique. Morbi tortor justo, porta eget conse quat.',
    { statW: 1.673, paraY: 5.68, paraW: 6.427, paraH: 0.601 });
  progressBar(s, 6.375, 5.267, 4.856, 0.083, 4.044, 11.511, 5.123, '90%');
}

/**
 * One jigsaw piece. `tabs` is [top, right, bottom, left]; +1 is a knob poking
 * outwards, -1 a matching socket, 0 a plain edge. x/y are the top-left of the
 * piece's *core* square, so neighbouring pieces simply step by `size`.
 */
function jigsawPiece(s, x, y, size, tabs, fill) {
  const k = 0.44; // how far a knob reaches past the edge
  const lo = 0.36 * size, hi = 0.64 * size; // where the knob interrupts the edge
  const bulge = 0.62; // bezier handle spread, tuned for a round knob
  const pts = [];
  const push = (px, py) => pts.push({ x: k + px, y: k + py });
  const curve = (x1, y1, x2, y2, px, py) =>
    pts.push({ x: k + px, y: k + py, curve: { type: 'cubic', x1: k + x1, y1: k + y1, x2: k + x2, y2: k + y2 } });

  // Edge walk, clockwise from the top-left corner of the core square.
  const sides = [
    { from: [0, 0], to: [size, 0], at: (u, o) => [u, -o] }, // top
    { from: [size, 0], to: [size, size], at: (u, o) => [size + o, u] }, // right
    { from: [size, size], to: [0, size], at: (u, o) => [size - u, size + o] }, // bottom
    { from: [0, size], to: [0, 0], at: (u, o) => [-o, size - u] }, // left
  ];
  push(...sides[0].from);
  sides.forEach((side, i) => {
    const t = tabs[i];
    if (t) {
      push(...side.at(lo, 0));
      curve(...side.at(lo - bulge * size * 0.18, k * t * 1.35),
            ...side.at(hi + bulge * size * 0.18, k * t * 1.35),
            ...side.at(hi, 0));
    }
    push(...side.to);
  });
  pts.push({ close: true });
  s.addShape('custGeom', { x: x - k, y: y - k, w: size + 2 * k, h: size + 2 * k, fill, line: HAIRLINE, points: pts });
}

/** Small white money glyph dropped inside a jigsaw piece. */
function puzzleIcon(s, kind, x, y, d) {
  const white = { color: WHITE };
  if (kind === 'piggy') {
    s.addShape('ellipse', { x, y: y + d * 0.24, w: d * 0.78, h: d * 0.5, fill: white, line: NO_LINE });
    s.addShape('ellipse', { x: x + d * 0.62, y: y + d * 0.34, w: d * 0.3, h: d * 0.22, fill: white, line: NO_LINE });
    s.addShape('rect', { x: x + d * 0.14, y: y + d * 0.66, w: d * 0.13, h: d * 0.16, fill: white, line: NO_LINE });
    s.addShape('rect', { x: x + d * 0.5, y: y + d * 0.66, w: d * 0.13, h: d * 0.16, fill: white, line: NO_LINE });
    s.addShape('custGeom', { x: x + d * 0.2, y: y + d * 0.1, w: d * 0.3, h: d * 0.18, fill: white, line: NO_LINE,
      points: [{ x: 0, y: d * 0.18 }, { x: d * 0.16, y: 0 }, { x: d * 0.3, y: d * 0.16 }, { close: true }] });
  } else if (kind === 'coins') {
    [0, 1, 2].forEach((i) =>
      s.addShape('ellipse', { x: x + d * 0.11 * i, y: y + d * (0.6 - 0.17 * i), w: d * 0.58, h: d * 0.22,
        fill: white, line: { color: DARK, width: 0.5 } }));
  } else if (kind === 'stack') {
    [0, 1, 2].forEach((i) =>
      s.addShape('ellipse', { x, y: y + d * (0.16 + 0.23 * i), w: d * 0.7, h: d * 0.24,
        fill: white, line: { color: DARK, width: 0.5 } }));
  } else {
    s.addShape('rect', { x, y: y + d * 0.34, w: d * 0.88, h: d * 0.4, fill: white, line: NO_LINE });
    s.addShape('ellipse', { x: x + d * 0.35, y: y + d * 0.45, w: d * 0.18, h: d * 0.18, fill: { color: DARK }, line: NO_LINE });
    s.addShape('rect', { x: x + d * 0.12, y: y + d * 0.2, w: d * 0.68, h: d * 0.07, fill: white, line: NO_LINE, rotate: -9 });
  }
}

function slide18(s) {
  rule(s, 0, 4.633, W, 0);
  const cell = 1.5755, x0 = 5.09, y0 = 3.058; // top-left core square origin
  const gradFill = { color: rampAverage(29) };
  jigsawPiece(s, x0, y0, cell, [1, -1, -1, 1], { color: DARK });
  jigsawPiece(s, x0 + cell, y0, cell, [-1, -1, 1, 1], gradFill);
  jigsawPiece(s, x0, y0 + cell, cell, [1, 1, -1, -1], gradFill);
  jigsawPiece(s, x0 + cell, y0 + cell, cell, [-1, 1, 1, -1], { color: DARK });
  puzzleIcon(s, 'piggy', 5.546, 3.469, 0.486);
  puzzleIcon(s, 'coins', 7.16, 3.846, 0.486);
  puzzleIcon(s, 'stack', 5.726, 5.031, 0.486);
  puzzleIcon(s, 'note', 7.431, 5.208, 0.486);
  // Direction chevrons sitting in the four knobs.
  const chevron = (x, y, w, h, pts) => s.addShape('custGeom', { x, y, w, h,
    line: { color: DARK, width: 1.75 }, points: pts });
  chevron(6.338, 3.779, 0.073, 0.147, [{ x: 0.073, y: 0 }, { x: 0, y: 0.0735 }, { x: 0.073, y: 0.147 }]);
  chevron(6.899, 5.346, 0.073, 0.147, [{ x: 0, y: 0 }, { x: 0.073, y: 0.0735 }, { x: 0, y: 0.147 }]);
  chevron(7.37, 4.889, 0.147, 0.073, [{ x: 0, y: 0 }, { x: 0.0735, y: 0.073 }, { x: 0.147, y: 0 }]);
  chevron(5.811, 4.313, 0.147, 0.073, [{ x: 0, y: 0.073 }, { x: 0.0735, y: 0 }, { x: 0.147, y: 0.073 }]);

  heading(s, 'EVENT DATA ', 'INFOGRAPHIC', 1.28, 1.027, 10.774, 1.111, 60, 'center');
  [
    { title: 'EVENT DATA 01', tx: 2.268, bx: 0.795, y: 2.967, align: 'right' },
    { title: 'EVENT DATA 02', tx: 8.931, bx: 8.931, y: 2.967, align: 'left' },
    { title: 'EVENT DATA 03', tx: 2.268, bx: 0.795, y: 5.291, align: 'right' },
    { title: 'EVENT DATA 04', tx: 8.931, bx: 8.931, y: 5.291, align: 'left' },
  ].forEach((l) => {
    txt(s, l.title, l.tx, l.y, 2.058, 0.37, { face: HEAD, size: 16, align: l.align });
    txt(s, L.short, l.bx, l.y + 0.46, 3.532, 0.601, { lnSpc: 1.3, align: l.align });
  });
}

function slide19(s) {
  const bars = [
    { x: 1.039, label: 'Q1', pct: '45%', y: 3.711, h: 1.887, stop: 51 },
    { x: 2.116, label: 'Q2', pct: '65%', y: 2.917, h: 2.682, stop: null },
    { x: 3.193, label: 'Q3', pct: '50%', y: 3.206, h: 2.392, stop: 42 },
    { x: 4.27, label: 'Q4', pct: '75%', y: 2.5, h: 3.098, stop: null },
    { x: 5.347, label: 'Q5', pct: '85%', y: 2.12, h: 3.478, stop: 45 },
  ];
  bars.forEach((b) => {
    if (b.stop === null) {
      s.addShape('rect', { x: b.x, y: b.y, w: 0.694, h: b.h, fill: { color: DARK }, line: NO_LINE,
        shadow: { type: 'outer', color: BLACK, opacity: 0.23, blur: 20, offset: 16, angle: 90 } });
    } else {
      gradientRect(s, b.x, b.y, 0.694, b.h, b.stop);
    }
    s.addShape('rect', { x: b.x, y: 1.38, w: 0.694, h: 4.219, line: HAIRLINE });
    txt(s, b.pct, b.x + 0.035, b.y + 0.253, 0.623, 0.283,
      { color: WHITE, align: 'center', valign: 'middle', lnSpc: 0.9 });
    txt(s, b.label, b.x, 5.777, 0.694, 0.343,
      { face: HEAD, size: 16, align: 'center', valign: 'middle', lnSpc: 0.9 });
  });
  rule(s, 0, 0.998, 6.674, 0);
  rule(s, 0, 6.502, 6.667, 0);
  rule(s, 6.674, 0.998, 0, 5.503);
  rule(s, 6.674, 3.964, 6.659, 0);
  heading(s, 'CUSTOM DATA', ' CHART', 7.168, 1.446, 5.4, 2.121);
  txt(s, '$3499+', 7.168, 4.332, 2.464, 0.505, { size: 24 });
  txt(s, 'EVENT PLAN PRESENTATION TEMPLATE ', 7.168, 5.033, 4.741, 0.37, { size: 16 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin bibendum risus sit amet venenatis tristique.',
    7.168, 5.532, 5.4, 0.601, { lnSpc: 1.3 });
}

/* Continent outlines traced from the source artwork, as x,y pairs in
   fractions of the map box. */
const MAP_LAND_PATHS = [
  [0.322,0.844, 0.336,0.794, 0.356,0.772, 0.373,0.681, 0.328,0.653, 0.327,0.626, 0.304,0.639, 0.302,0.62, 0.287,0.644, 0.275,0.64, 0.275,0.673, 0.263,0.691, 0.273,0.711, 0.288,0.705, 0.307,0.745, 0.32,0.804, 0.31,0.824, 0.322,0.844],
  [0.319,0.796, 0.306,0.806, 0.294,0.774, 0.278,0.79, 0.269,0.972, 0.278,0.974, 0.288,0.892, 0.311,0.866, 0.308,0.827, 0.319,0.796],
  [0.269,0.971, 0.268,0.892, 0.282,0.779, 0.275,0.749, 0.259,0.945, 0.27,0.978, 0.264,0.986, 0.273,0.976, 0.269,0.99, 0.282,0.995, 0.269,0.971],
  [0.254,0.647, 0.275,0.673, 0.281,0.616, 0.265,0.598, 0.271,0.579, 0.252,0.608, 0.254,0.647],
  [0.309,0.75, 0.288,0.705, 0.276,0.711, 0.28,0.78, 0.309,0.75],
  [0.275,0.749, 0.278,0.719, 0.263,0.691, 0.274,0.665, 0.259,0.651, 0.243,0.682, 0.258,0.732, 0.275,0.749],
  [0.268,0.585, 0.284,0.643, 0.301,0.621, 0.296,0.59, 0.268,0.585],
  [0.895,0.812, 0.865,0.71, 0.859,0.751, 0.836,0.712, 0.829,0.735, 0.821,0.728, 0.785,0.774, 0.782,0.799, 0.79,0.855, 0.832,0.833, 0.875,0.882, 0.885,0.873, 0.895,0.812],
  [0.558,0.803, 0.556,0.776, 0.53,0.802, 0.525,0.791, 0.525,0.811, 0.515,0.814, 0.521,0.848, 0.543,0.845, 0.558,0.803],
  [0.502,0.46, 0.499,0.506, 0.539,0.534, 0.54,0.462, 0.502,0.46],
  [0.499,0.61, 0.51,0.574, 0.483,0.572, 0.48,0.611, 0.488,0.624, 0.499,0.61],
  [0.567,0.476, 0.54,0.463, 0.539,0.522, 0.564,0.524, 0.56,0.473, 0.564,0.488, 0.567,0.476],
  [0.564,0.607, 0.579,0.629, 0.602,0.605, 0.576,0.565, 0.564,0.607],
  [0.579,0.703, 0.574,0.665, 0.554,0.655, 0.553,0.695, 0.57,0.715, 0.579,0.703],
  [0.213,0.558, 0.227,0.527, 0.206,0.545, 0.189,0.481, 0.143,0.46, 0.163,0.517, 0.154,0.467, 0.177,0.539, 0.213,0.558],
  [0.428,0.032, 0.404,0.049, 0.407,0.03, 0.379,0.035, 0.41,0.024, 0.398,0.013, 0.341,0.016, 0.346,0.033, 0.299,0.035, 0.282,0.055, 0.286,0.068, 0.267,0.079, 0.285,0.09, 0.271,0.093, 0.279,0.105, 0.311,0.12, 0.317,0.155, 0.327,0.163, 0.319,0.163, 0.329,0.177, 0.32,0.197, 0.325,0.224, 0.349,0.263, 0.364,0.208, 0.408,0.172, 0.398,0.168, 0.409,0.166, 0.403,0.146, 0.41,0.141, 0.402,0.137, 0.416,0.126, 0.408,0.098, 0.42,0.096, 0.409,0.084, 0.421,0.061, 0.414,0.06, 0.437,0.037, 0.428,0.032],
  [0.184,0.172, 0.177,0.132, 0.168,0.154, 0.151,0.136, 0.138,0.15, 0.158,0.167, 0.142,0.172, 0.158,0.186, 0.184,0.172],
  [0.216,0.035, 0.248,0.043, 0.242,0.059, 0.229,0.053, 0.24,0.075, 0.225,0.079, 0.234,0.09, 0.219,0.099, 0.243,0.102, 0.262,0.071, 0.254,0.068, 0.297,0.024, 0.247,0.018, 0.216,0.035],
  [0.285,0.208, 0.299,0.204, 0.27,0.157, 0.244,0.149, 0.242,0.133, 0.231,0.138, 0.233,0.155, 0.232,0.132, 0.222,0.137, 0.22,0.162, 0.249,0.166, 0.276,0.204, 0.252,0.224, 0.285,0.247, 0.278,0.232, 0.29,0.234, 0.285,0.208],
  [0.302,0.368, 0.292,0.37, 0.29,0.349, 0.272,0.363, 0.314,0.326, 0.29,0.259, 0.281,0.277, 0.276,0.255, 0.253,0.241, 0.247,0.333, 0.239,0.302, 0.211,0.288, 0.205,0.261, 0.216,0.234, 0.208,0.227, 0.223,0.227, 0.22,0.21, 0.238,0.204, 0.243,0.184, 0.231,0.172, 0.226,0.198, 0.217,0.187, 0.206,0.15, 0.214,0.129, 0.203,0.132, 0.208,0.184, 0.202,0.197, 0.198,0.174, 0.195,0.192, 0.167,0.186, 0.174,0.185, 0.169,0.202, 0.114,0.166, 0.098,0.184, 0.108,0.17, 0.075,0.175, 0.075,0.26, 0.093,0.268, 0.127,0.35, 0.223,0.349, 0.245,0.372, 0.24,0.396, 0.277,0.36, 0.292,0.375, 0.286,0.386, 0.302,0.368],
  [0.281,0.362, 0.242,0.4, 0.232,0.374, 0.226,0.399, 0.225,0.378, 0.235,0.368, 0.213,0.364, 0.22,0.356, 0.205,0.346, 0.127,0.349, 0.124,0.412, 0.143,0.46, 0.173,0.464, 0.198,0.5, 0.208,0.475, 0.236,0.475, 0.245,0.504, 0.242,0.47, 0.258,0.44, 0.256,0.415, 0.259,0.425, 0.283,0.377, 0.281,0.362],
  [0.106,0.299, 0.093,0.267, 0.076,0.26, 0.076,0.175, 0.032,0.158, 0.005,0.183, 0.021,0.207, 0,0.214, 0.019,0.221, 0.008,0.259, 0.029,0.273, 0.014,0.304, 0.048,0.252, 0.046,0.27, 0.057,0.254, 0.08,0.263, 0.093,0.273, 0.093,0.292, 0.097,0.283, 0.098,0.306, 0.106,0.299],
  [0.58,0.357, 0.586,0.343, 0.569,0.32, 0.539,0.326, 0.535,0.349, 0.55,0.349, 0.552,0.372, 0.563,0.362, 0.567,0.378, 0.58,0.357],
  [0.557,0.241, 0.555,0.168, 0.531,0.173, 0.543,0.21, 0.532,0.251, 0.557,0.241],
  [0.495,0.346, 0.479,0.33, 0.459,0.35, 0.467,0.387, 0.482,0.393, 0.494,0.384, 0.495,0.346],
  [0.54,0.201, 0.532,0.173, 0.514,0.202, 0.505,0.266, 0.51,0.295, 0.54,0.201],
  [0.559,0.156, 0.516,0.172, 0.506,0.216, 0.487,0.242, 0.488,0.267, 0.505,0.267, 0.508,0.224, 0.525,0.177, 0.559,0.156],
  [0.797,0.361, 0.789,0.361, 0.792,0.345, 0.743,0.329, 0.738,0.346, 0.714,0.351, 0.735,0.396, 0.758,0.404, 0.797,0.361],
  [0.593,0.427, 0.589,0.403, 0.544,0.415, 0.552,0.437, 0.593,0.427],
  [0.737,0.49, 0.701,0.493, 0.689,0.464, 0.694,0.446, 0.677,0.451, 0.661,0.515, 0.672,0.524, 0.685,0.605, 0.694,0.561, 0.716,0.526, 0.715,0.501, 0.726,0.526, 0.737,0.49],
  [0.601,0.483, 0.581,0.463, 0.567,0.478, 0.588,0.557, 0.623,0.536, 0.601,0.483],
  [0.646,0.493, 0.641,0.436, 0.614,0.436, 0.594,0.414, 0.597,0.454, 0.613,0.489, 0.639,0.506, 0.646,0.493],
  [0.711,0.347, 0.665,0.305, 0.641,0.314, 0.64,0.34, 0.613,0.332, 0.6,0.355, 0.618,0.368, 0.61,0.385, 0.618,0.405, 0.637,0.372, 0.661,0.408, 0.694,0.399, 0.711,0.347],
  [0.997,0.197, 0.958,0.155, 0.916,0.164, 0.908,0.142, 0.86,0.118, 0.856,0.136, 0.832,0.144, 0.813,0.105, 0.761,0.118, 0.784,0.083, 0.756,0.056, 0.697,0.11, 0.701,0.131, 0.688,0.122, 0.69,0.142, 0.681,0.119, 0.677,0.162, 0.69,0.18, 0.679,0.167, 0.673,0.193, 0.664,0.189, 0.676,0.171, 0.671,0.118, 0.657,0.142, 0.662,0.172, 0.635,0.148, 0.64,0.167, 0.604,0.186, 0.592,0.167, 0.594,0.194, 0.574,0.204, 0.574,0.217, 0.56,0.182, 0.577,0.194, 0.585,0.179, 0.557,0.153, 0.55,0.167, 0.558,0.23, 0.547,0.277, 0.56,0.328, 0.582,0.351, 0.573,0.385, 0.603,0.416, 0.602,0.346, 0.613,0.333, 0.64,0.341, 0.641,0.312, 0.665,0.301, 0.713,0.354, 0.738,0.349, 0.743,0.33, 0.769,0.352, 0.795,0.351, 0.802,0.321, 0.816,0.321, 0.823,0.351, 0.841,0.365, 0.831,0.394, 0.837,0.407, 0.86,0.328, 0.844,0.303, 0.862,0.27, 0.899,0.266, 0.905,0.242, 0.913,0.251, 0.926,0.233, 0.904,0.279, 0.903,0.341, 0.918,0.308, 0.923,0.259, 0.967,0.234, 0.958,0.211, 0.971,0.195, 0.988,0.216, 0.997,0.197],
  [0.684,0.441, 0.67,0.438, 0.656,0.477, 0.64,0.479, 0.642,0.506, 0.668,0.51, 0.664,0.495, 0.684,0.441],
];
const MAP_DARK_PATHS = [
  [0.506,0.745, 0.534,0.749, 0.531,0.722, 0.536,0.721, 0.536,0.709, 0.532,0.711, 0.53,0.69, 0.52,0.694, 0.516,0.681, 0.506,0.684, 0.509,0.713, 0.506,0.745],
  [0.467,0.451, 0.467,0.462, 0.447,0.483, 0.447,0.492, 0.48,0.541, 0.505,0.514, 0.498,0.498, 0.498,0.473, 0.492,0.449, 0.495,0.431, 0.475,0.434, 0.465,0.443, 0.467,0.451],
  [0.543,0.613, 0.533,0.586, 0.511,0.615, 0.516,0.634, 0.524,0.62, 0.546,0.619, 0.543,0.613],
  [0.426,0.555, 0.436,0.565, 0.456,0.561, 0.452,0.505, 0.457,0.505, 0.447,0.491, 0.447,0.499, 0.437,0.499, 0.434,0.527, 0.423,0.527, 0.426,0.555],
  [0.48,0.541, 0.452,0.505, 0.456,0.561, 0.437,0.565, 0.439,0.58, 0.455,0.59, 0.46,0.572, 0.482,0.559, 0.48,0.541],
  [0.513,0.523, 0.505,0.514, 0.483,0.54, 0.482,0.559, 0.472,0.564, 0.475,0.575, 0.506,0.574, 0.514,0.553, 0.513,0.523],
  [0.536,0.536, 0.513,0.516, 0.514,0.553, 0.509,0.571, 0.514,0.606, 0.533,0.586, 0.531,0.568, 0.536,0.536],
  [0.567,0.517, 0.564,0.524, 0.539,0.522, 0.531,0.569, 0.535,0.598, 0.555,0.628, 0.569,0.622, 0.561,0.604, 0.576,0.546, 0.567,0.517],
  [0.551,0.688, 0.55,0.662, 0.556,0.636, 0.545,0.619, 0.524,0.62, 0.514,0.671, 0.505,0.677, 0.506,0.684, 0.516,0.682, 0.52,0.694, 0.53,0.69, 0.532,0.712, 0.55,0.723, 0.551,0.688],
];
const MAP_HOT_PATHS = [
  [0.766,0.527, 0.76,0.517, 0.751,0.523, 0.765,0.56, 0.767,0.575, 0.757,0.592, 0.759,0.601, 0.771,0.581, 0.761,0.539, 0.766,0.527],
  [0.759,0.589, 0.765,0.569, 0.752,0.573, 0.753,0.587, 0.759,0.589],
  [0.759,0.545, 0.751,0.523, 0.745,0.533, 0.747,0.55, 0.758,0.551, 0.761,0.572, 0.765,0.56, 0.759,0.545],
  [0.752,0.578, 0.76,0.564, 0.756,0.547, 0.747,0.55, 0.745,0.533, 0.739,0.54, 0.743,0.581, 0.74,0.603, 0.751,0.615, 0.743,0.597, 0.745,0.577, 0.752,0.578],
  [0.755,0.624, 0.745,0.613, 0.751,0.637, 0.756,0.643, 0.755,0.624],
  [0.796,0.618, 0.791,0.612, 0.772,0.643, 0.783,0.643, 0.796,0.618],
  [0.841,0.36, 0.831,0.36, 0.816,0.322, 0.805,0.32, 0.789,0.354, 0.801,0.368, 0.758,0.404, 0.735,0.396, 0.721,0.377, 0.722,0.367, 0.712,0.349, 0.709,0.365, 0.694,0.381, 0.694,0.4, 0.678,0.413, 0.677,0.422, 0.684,0.441, 0.694,0.446, 0.689,0.464, 0.695,0.477, 0.716,0.495, 0.735,0.482, 0.742,0.495, 0.738,0.511, 0.743,0.524, 0.749,0.53, 0.76,0.518, 0.772,0.535, 0.793,0.515, 0.806,0.478, 0.799,0.448, 0.807,0.434, 0.8,0.433, 0.795,0.42, 0.806,0.409, 0.806,0.421, 0.831,0.396, 0.841,0.36],
  [0.798,0.67, 0.8,0.683, 0.801,0.667, 0.806,0.676, 0.805,0.66, 0.809,0.655, 0.8,0.652, 0.814,0.642, 0.801,0.646, 0.798,0.67],
  [0.761,0.675, 0.761,0.659, 0.732,0.619, 0.751,0.671, 0.757,0.683, 0.761,0.675],
  [0.785,0.695, 0.76,0.688, 0.785,0.695],
  [0.849,0.659, 0.842,0.669, 0.839,0.655, 0.831,0.657, 0.838,0.663, 0.836,0.671, 0.85,0.681, 0.849,0.697, 0.859,0.701, 0.859,0.666, 0.849,0.659],
  [0.791,0.659, 0.797,0.645, 0.793,0.626, 0.783,0.643, 0.771,0.639, 0.769,0.647, 0.778,0.67, 0.789,0.67, 0.791,0.659],
];

function slide20(s) {
  const mx = -1.746, my = 1.339, mw = 9.115, mh = 4.822;
  // Paths arrive as flat [x0,y0, x1,y1, ...] fractions of the map box.
  const land = (flat, color) => {
    const pts = [];
    for (let i = 0; i < flat.length; i += 2) pts.push({ x: mw * flat[i], y: mh * flat[i + 1] });
    pts.push({ close: true });
    s.addShape('custGeom', { x: mx, y: my, w: mw, h: mh, fill: { color }, line: NO_LINE, points: pts });
  };
  MAP_LAND_PATHS.forEach((c) => land(c, MAP_GREY));
  MAP_HOT_PATHS.forEach((c) => land(c, rampAverage(29)));
  MAP_DARK_PATHS.forEach((c) => land(c, DARK));
  const pin = (x, y, label, fill, line) => {
    s.addShape('teardrop', { x, y, w: 0.351, h: 0.351, fill, line, rotate: 135 });
    txt(s, label, x - 0.009, y + 0.049, 0.37, 0.252, { face: HEAD, size: 9, color: WHITE, align: 'center' });
  };
  pin(4.976, 2.987, '01', { color: rampAverage(29) }, HAIRLINE);
  pin(2.437, 3.202, '02', { color: DARK }, { color: ORANGE, width: 0.5 });

  gradientRect(s, 7.602, 3.808, 5.731, 1.2, 29);
  // Two donut gauges: a ring pie, a white "gap" pie and a solid hub.
  const gauge = (x, y, ring, hub, pct, arc) => {
    s.addShape('pie', { x, y, w: 0.751, h: 0.751, fill: { color: ring }, line: NO_LINE, angleRange: arc });
    s.addShape('pie', { x, y, w: 0.751, h: 0.751, fill: { color: WHITE }, line: NO_LINE, angleRange: [268.96, 328.25] });
    s.addShape('ellipse', { x: x + 0.12, y: y + 0.121, w: 0.51, h: 0.51, fill: { color: hub }, line: NO_LINE });
    txt(s, pct, x + 0.09, y + 0.224, 0.572, 0.303, { face: HEAD, color: WHITE, align: 'center' });
  };
  gauge(7.96, 4.032, DARK, ORANGE, '85%', [326.5, 270]);
  txt(s, 'MARKET AREA ONE ', 9.132, 3.963, 2.458, 0.37, { face: HEAD, size: 16, color: WHITE });
  txt(s, '$538 - $1064', 9.132, 4.348, 2.512, 0.505, { face: HEAD, size: 24, color: WHITE });
  gauge(7.96, 5.308, ORANGE, DARK, '75%', [0.01, 270]);
  txt(s, 'MARKET AREA TWO ', 9.132, 5.239, 2.458, 0.37, { face: HEAD, size: 16 });
  txt(s, '$538 - $1064', 9.132, 5.624, 2.651, 0.505, { face: HEAD, size: 24, color: GREY_TEXT });

  heading(s, 'WORLD DATA ', 'HEAT MAP ', 7.931, 1.294, 4.921, 1.919, 54);
  rule(s, 7.602, 0, 0, 7.5);
  rule(s, 7.602, 6.354, 5.731, 0);
}

function slide21(s) {
  gradientRect(s, 0, 0, W, 3.316, 29);
  s.addText(
    [
      { text: 'THANK', options: { color: ORANGE } },
      { text: ' YOU ', options: { color: DARK } },
    ],
    { x: 0.779, y: 3.75, w: 6.722, h: 3.253, fontFace: HEAD, fontSize: 115,
      align: 'right', lineSpacingMultiple: 0.8, valign: 'top', margin: INSETS, isTextBox: true }
  );
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed iaculis vel neque ac cursus. Donec aliquam consectetur adipiscing elit.',
    0.667, 0.854, 5.742, 0.601, { color: WHITE, lnSpc: 1.3 });
  txt(s, 'Company Name ', 0.667, 6.343, 2.18, 0.303, { face: HEAD, color: BLACK });
}

/* =================================================================== build */

const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07,
  slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15,
  slide16, slide17, slide18, slide19, slide20, slide21];

/* Slides 1 and 21 switch the master's page number / deck name off. */
const NO_CHROME = new Set([1, 21]);

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE_13x7.5', width: W, height: H });
pptx.layout = 'WIDE_13x7.5';
pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
pptx.title = 'Event Plan';

BUILDERS.forEach((build, i) => {
  const slide = pptx.addSlide();
  slide.background = { color: WHITE };
  if (!NO_CHROME.has(i + 1)) slideChrome(slide, i + 1);
  build(slide);
});

pptx.writeFile({ fileName: path.join(__dirname, '0eab9463-a6d7-453c-9557-e4b800c8a7d8_grok_final.pptx') })
  .then((f) => console.log('wrote', f));
