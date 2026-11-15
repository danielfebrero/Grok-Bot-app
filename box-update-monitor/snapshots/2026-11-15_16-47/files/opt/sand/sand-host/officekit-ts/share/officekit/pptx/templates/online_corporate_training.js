/**
 * Rebuild of "Online Training - Presentation Template" (20 slides, 10 x 5.625in)
 * with PptxGenJS. Photographs in the source deck are replaced by grey
 * "[image]" placeholder shapes that keep the original clipping geometry.
 */
const pptxgen = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ palette */
const C = {
  pink:      'D813BA', // primary magenta
  magenta:   'EF4BD6', // accent1
  pinkSoft:  'D04CBC',
  lime:      'AFEE26', // accent2
  limeSoft:  'D4F66E',
  limeDark:  '88BF0F',
  limeDark2: '88C00F',
  blue:      '0E32AD', // accent3
  black:     '000000',
  ink:       '0C0C0C',
  grey:      '3F3F3F',
  grey2:     '7F7F7F',
  grey3:     '595959',
  silver:    'BFBFBF',
  hair:      'E5E5E5',
  white:     'FFFFFF',
  panel:     'F9F9F9',
  photo:     'C0C0C0', // stand-in colour for photographs
  photoTxt:  '8A8A8A',
};

const HEAD = 'Parkinsans';   // display / heading face
const BODY = 'Noto Sans';    // body face
const INSET = [5.4, 5.4, 2.7, 2.7];  // pt - matches lIns/rIns 68575, tIns/bIns 34275 EMU

/* ------------------------------------------------------------------ helpers */
const fill = (color, transparency) => (transparency === undefined ? { color } : { color, transparency });

function txt(slide, content, o) {
  slide.addText(content, Object.assign({
    fontFace: BODY, fontSize: 9, color: C.grey,
    align: 'left', valign: 'top', margin: INSET, isTextBox: true,
  }, o));
}

/** 9pt body copy (130% leading, as in the source deck) */
const body = (s, t, x, y, w, h, o) => txt(s, t, Object.assign({ x, y, w, h, lineSpacingMultiple: 1.3 }, o));

/** run list for the two-tone display titles */
const runs = (parts, size = 41) => parts.map(([text, color, brk]) => ({
  text, options: { fontFace: HEAD, fontSize: size, color, breakLine: !!brk },
}));

/** rounded rectangle; `adj` is the OOXML corner adjust value (1/100000 of the short side) */
function roundRect(slide, x, y, w, h, colour, adj, extra) {
  slide.addShape('roundRect', Object.assign({
    x, y, w, h, fill: colour, rectRadius: (adj / 100000) * Math.min(w, h),
  }, extra));
}

/** the pink "highlighter" plate that sits behind one word of every title */
const plate = (slide, x, y, w, h, colour) => roundRect(slide, x, y, w, h, colour || fill(C.pink), 16667);

/** straight connector between two points */
function line(slide, x1, y1, x2, y2, o) {
  if (x2 < x1) { const tx = x1, ty = y1; x1 = x2; y1 = y2; x2 = tx; y2 = ty; }
  slide.addShape('line', { x: x1, y: Math.min(y1, y2), w: x2 - x1, h: Math.abs(y2 - y1), flipV: y2 < y1, line: o });
}

/** custom outline; `pathData` holds 0..1 coordinates - M/L/C/Z like an SVG path */
function customPath(slide, pathData, o) {
  const pts = [];
  pathData.forEach((seg) => {
    if (seg[0] === 'M') pts.push({ x: seg[1] * o.w, y: seg[2] * o.h, moveTo: true });
    else if (seg[0] === 'L') pts.push({ x: seg[1] * o.w, y: seg[2] * o.h });
    else if (seg[0] === 'C') pts.push({ x: seg[5] * o.w, y: seg[6] * o.h, curve: { type: 'cubic', x1: seg[1] * o.w, y1: seg[2] * o.h, x2: seg[3] * o.w, y2: seg[4] * o.h } });
    else pts.push({ close: true });
  });
  slide.addShape('custGeom', Object.assign({}, o, { points: pts }));
}

/* ------------------------------------------------------- reusable geometry */
const P = {
  // slide 1 - scalloped ribbon behind the cover title
  ribbon: [['M', 1, 0], ['C', .861, .009, .801, .249, .765, .394], ['C', .762, .406, .758, .42, .755, .434],
    ['C', .751, .42, .748, .406, .745, .394], ['C', .708, .247, .646, 0, .502, 0], ['C', .357, 0, .296, .247, .259, .394],
    ['C', .256, .406, .253, .42, .249, .434], ['C', .245, .42, .242, .406, .239, .394], ['C', .203, .248, .142, .004, 0, 0],
    ['L', 0, .581], ['C', .002, .59, .004, .598, .006, .606], ['C', .043, .753, .105, 1, .249, 1],
    ['C', .394, 1, .455, .753, .492, .606], ['C', .495, .594, .498, .58, .502, .566], ['C', .505, .58, .509, .594, .512, .606],
    ['C', .549, .753, .61, 1, .755, 1], ['C', .899, 1, .961, .753, .997, .606], ['C', .998, .603, .999, .599, 1, .595],
    ['L', 1, 0], ['Z']],
  // slide 1 - photo frame with two stepped tabs on the top edge
  frameTab: [['M', .07, 0], ['L', .785, 0], ['C', .811, 0, .832, .02, .832, .044], ['L', .832, .116],
    ['C', .832, .136, .849, .152, .87, .152], ['L', .964, .152], ['C', .985, .156, 1, .173, 1, .193], ['L', 1, .935],
    ['C', 1, .971, .969, 1, .93, 1], ['L', .07, 1], ['C', .031, 1, 0, .971, 0, .935], ['L', 0, .065],
    ['C', 0, .029, .031, 0, .07, 0], ['Z']],
  // slide 2 - wide photo frame with a notch on the top edge
  frameNotch: [['M', .406, 0], ['L', .931, 0], ['C', .969, 0, 1, .046, 1, .103], ['L', 1, .897],
    ['C', 1, .954, .969, 1, .931, 1], ['L', .069, 1], ['C', .031, 1, 0, .954, 0, .897], ['L', 0, .287],
    ['C', 0, .24, .026, .201, .058, .201], ['L', .305, .201], ['L', .318, .194], ['C', .331, .185, .342, .168, .347, .147],
    ['L', .349, .133], ['L', .349, .085], ['C', .349, .038, .375, 0, .406, 0], ['Z']],
  // slide 4 - card whose top-left corner is a quarter circle
  quarter: [['M', .119, 0], ['C', .575, 0, .95, .347, .996, .793], ['L', 1, .88], ['L', .998, .903],
    ['C', .986, .958, .937, 1, .878, 1], ['L', .122, 1], ['C', .055, 1, 0, .945, 0, .878], ['L', 0, .12],
    ['C', 0, .069, .031, .026, .074, .007], ['L', .095, .001], ['Z']],
  // slide 5 - photo frame with a side tab
  frameSide: [['M', .48, 0], ['L', .945, 0], ['C', .976, 0, 1, .025, 1, .055], ['L', 1, .945],
    ['C', 1, .975, .976, 1, .945, 1], ['L', .48, 1], ['C', .45, 1, .425, .975, .425, .945], ['L', .424, .904],
    ['C', .419, .879, .397, .86, .371, .86], ['L', .044, .86], ['C', .019, .854, 0, .832, 0, .805], ['L', 0, .195],
    ['C', 0, .164, .024, .139, .055, .139], ['L', .371, .139], ['C', .401, .139, .425, .115, .425, .084], ['L', .425, .055],
    ['C', .425, .025, .45, 0, .48, 0], ['Z']],
  // slide 7 - two half-round photo panels stacked with a gap
  splitLeaf: [['M', .12, .52], ['L', .878, .52], ['C', .945, .52, 1, .546, 1, .579], ['L', 1, .942],
    ['C', 1, .97, .958, .993, .903, .999], ['L', .88, 1], ['L', .793, .998], ['C', .347, .976, 0, .796, 0, .577],
    ['L', .007, .556], ['C', .026, .535, .069, .52, .12, .52], ['Z'],
    ['M', .88, 0], ['L', .903, .001], ['C', .958, .007, 1, .03, 1, .058], ['L', 1, .421],
    ['C', 1, .454, .945, .48, .878, .48], ['L', .12, .48], ['C', .069, .48, .026, .465, .007, .444], ['L', 0, .423],
    ['C', 0, .204, .347, .024, .793, .002], ['Z']],
  // slide 8 - banner with two rounded corners on one side only
  banner: [['M', 0, 0], ['L', 1, 0], ['L', 1, .82], ['C', 1, .92, .966, 1, .925, 1], ['L', .075, 1],
    ['C', .034, 1, 0, .92, 0, .82], ['Z']],
  // slide 9 - tall panel with one big rounded corner
  leaf: [['M', .61, 0], ['C', .757, 0, .892, .025, .998, .066], ['L', 1, .069], ['L', 1, .931],
    ['C', 1, .969, .935, 1, .855, 1], ['L', .145, 1], ['C', .065, 1, 0, .969, 0, .931], ['L', 0, .29],
    ['C', 0, .13, .273, 0, .61, 0], ['Z']],
  // slide 19 - soft "hill" area behind the trend chart
  hill: [['M', 1, 0], ['L', 1, .92], ['C', 1, .964, .983, 1, .961, 1], ['L', .039, 1], ['C', .017, 1, 0, .964, 0, .92],
    ['L', 0, .764], ['L', .058, .759], ['C', .455, .713, .794, .44, .975, .055], ['L', 1, 0], ['Z']],
  // small ">" arrow glyph used inside every round button
  arrow: [['M', 1, .5], ['L', .487, 1], ['L', .45, .964], ['L', .901, .525], ['L', 0, .525], ['L', 0, .475],
    ['L', .901, .475], ['L', .45, .036], ['L', .487, 0], ['Z']],
};

/* --------------------------------------------------- picture placeholders */
function caption(slide, x, y, w, h) {
  if (w < 1.1 || h < 0.7) return;
  txt(slide, '[image]', { x, y: y + h / 2 - 0.16, w, h: 0.32, align: 'center', valign: 'middle', color: C.photoTxt, fontSize: 10 });
}
/** rectangular photo stand-in */
function photoRect(slide, x, y, w, h) {
  slide.addShape('rect', { x, y, w, h, fill: fill(C.photo) });
  caption(slide, x, y, w, h);
}
/** rounded photo stand-in */
function photoRound(slide, x, y, w, h, adj) {
  roundRect(slide, x, y, w, h, fill(C.photo), adj);
  caption(slide, x, y, w, h);
}
/** photo stand-in clipped to one of the custom outlines above */
function photoPath(slide, pathData, x, y, w, h, extra) {
  customPath(slide, pathData, Object.assign({ x, y, w, h, fill: fill(C.photo) }, extra));
  caption(slide, x, y, w, h);
}
/** small vector icon stand-in (source deck uses tiny png glyphs) */
function icon(slide, x, y, size, colour) {
  roundRect(slide, x, y, size, size, fill(colour), 22000);
}

/** circular icon badge: soft halo + solid disc + glyph */
function iconBadge(slide, x, y, halo, disc, haloColour, discColour, glyphColour) {
  slide.addShape('ellipse', { x, y, w: halo, h: halo, fill: haloColour });
  const off = (halo - disc) / 2;
  slide.addShape('ellipse', { x: x + off, y: y + off, w: disc, h: disc, fill: fill(discColour) });
  const g = disc * 0.39;
  icon(slide, x + (halo - g) / 2, y + (halo - g) / 2, g, glyphColour);
}

/** round arrow chip (lime disc + white arrow) */
function arrowChip(slide, x, y, d, discColour, arrowFill, arrowLine) {
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: fill(discColour) });
  const w = d * 0.231, h = d * 0.238;
  customPath(slide, P.arrow, {
    x: x + d * 0.384, y: y + d * 0.381, w, h, rotate: -46.09,
    fill: fill(arrowFill), line: { color: arrowLine, width: 0.75 },
  });
}

/** "Read more" pill + arrow chip */
function readMore(slide, x, y, labelColour) {
  roundRect(slide, x, y, 1.282, 0.393, fill(C.pink), 50000, { line: { color: C.silver, width: 0.75 } });
  txt(slide, 'Read more', { x: x + 0.02, y: y + 0.086, w: 0.93, h: 0.227, align: 'center', fontFace: HEAD, color: labelColour });
  arrowChip(slide, x + 0.942, y + 0.05, 0.294, C.lime, C.white, C.grey);
}

const pres = new pptxgen();
pres.defineLayout({ name: 'DECK', width: 10, height: 5.625 });
pres.layout = 'DECK';
const newSlide = () => pres.addSlide();

/* ============================================================ 1. cover */
function slide01() {
  const s = newSlide();
  customPath(s, P.ribbon, { x: 2.941, y: 2.022, w: 4.21, h: 2.2, fill: fill(C.pink, 53) });
  plate(s, 5.0, 1.914, 4.206, 1.257);
  photoPath(s, P.frameTab, 0.517, 0.568, 4.206, 4.49);
  txt(s, runs([['Online ', C.pink], ['Training', C.lime]], 66), { x: 5.147, y: 0.874, w: 4.299, h: 2.297 });
  txt(s, 'Presentation Template', { x: 5.277, y: 3.458, w: 3.747, h: 0.429, fontFace: HEAD, fontSize: 21, color: C.ink });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula', 5.277, 4.063, 3.803, 0.453);
}

/* ================================================= 2. training journey */
function slide02() {
  const s = newSlide();
  s.addShape('rect', { x: 5.931, y: 0, w: 4.069, h: 5.625, fill: fill(C.lime) });
  plate(s, 2.828, 1.322, 2.626, 0.698);
  txt(s, runs([['Our Online Training ', C.black], ['Journey', C.lime]]), { x: 0.462, y: 0.579, w: 5.0, h: 1.439 });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis', 0.512, 2.052, 4.432, 0.644);
  photoPath(s, P.frameNotch, 0.512, 2.812, 4.488, 2.245);

  ['01', '02', '03'].forEach((n, i) => {
    const y = 0.629 + i * 1.544;
    s.addText(n, { shape: 'star24', x: 6.307, y, w: 0.698, h: 0.698, fill: fill(C.pink),
      fontFace: HEAD, fontSize: 12, color: C.white, align: 'center', valign: 'middle', margin: INSET });
    txt(s, 'Journey ' + n, { x: 7.206, y: y + 0.084, w: 2.48, h: 0.379, fontFace: i ? BODY : HEAD, fontSize: 18 });
    body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula', 6.307, y + 0.832, 3.378, 0.453);
  });
}

/* ============================================== 3. online learning intro */
function slide03() {
  const s = newSlide();
  s.addShape('rect', { x: 0, y: 0, w: 2.417, h: 5.625, fill: fill(C.pink) });
  photoRect(s, 0.608, 0.568, 3.535, 4.49);
  plate(s, 4.625, 1.818, 3.4, 0.655);
  txt(s, [
    { text: 'Upskill Your Future', options: { fontFace: HEAD, fontSize: 14, color: C.black, breakLine: true } },
    ...runs([['Online Learning ', C.black], ['Starts Here', C.lime]]),
  ], { x: 4.625, y: 0.866, w: 5.0, h: 1.666 });
  iconBadge(s, 4.625, 2.893, 0.602, 0.427, fill(C.limeSoft, 82), C.pink, C.white);
  txt(s, 'Mastering Skills From Anywhere', { x: 5.346, y: 3.042, w: 3.287, h: 0.303, fontSize: 14, color: C.black });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis', 5.346, 3.424, 3.937, 0.644);
}

/* ============================================ 4. new employee orientation */
function slide04() {
  const s = newSlide();
  photoPath(s, P.quarter, 4.933, 0.557, 2.161, 2.157, { rotate: 180, flipH: true, fill: fill(C.pink) });
  txt(s, 'More Volunteers Joined the Cause', { x: 5.064, y: 0.743, w: 1.764, h: 0.53, fontFace: HEAD, fontSize: 14, color: C.white });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing', 5.064, 1.535, 1.542, 0.65, { color: C.white });

  customPath(s, P.quarter, { x: 4.933, y: 2.9, w: 2.161, h: 2.157, fill: fill(C.lime) });
  arrowChip(s, 5.1, 3.101, 0.526, C.pink, C.blue, C.blue);
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo', 5.1, 4.218, 1.818, 0.644);

  plate(s, 0.433, 2.158, 3.499, 0.742);
  txt(s, runs([['New Employee Online Training ', C.black], ['Orientation', C.lime]]), { x: 0.433, y: 0.819, w: 5.0, h: 2.121 });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et', 0.457, 3.193, 3.589, 0.644);
  readMore(s, 0.512, 3.935, C.lime);
  photoPath(s, P.frameSide, 7.179, 0.568, 2.309, 4.5);
}

/* ================================================= 5. workplace success */
function slide05() {
  const s = newSlide();
  photoPath(s, P.frameSide, 0.55, 0.568, 4.45, 3.893);

  // date card
  roundRect(s, 0.553, 3.986, 1.829, 0.932, fill(C.lime), 12268);
  txt(s, 'Online Training', { x: 0.66, y: 4.076, w: 1.383, h: 0.227, fontFace: HEAD, color: C.ink });
  line(s, 0.66, 4.353, 2.043, 4.353, { color: C.grey2, transparency: 78, width: 0.75 });
  iconBadge(s, 0.66, 4.47, 0.329, 0.233, fill(C.magenta, 82), C.magenta, C.white);
  txt(s, '25 August 2025', { x: 1.068, y: 4.48, w: 1.315, h: 0.215, fontFace: HEAD, fontSize: 8, color: C.ink });
  line(s, 1.119, 4.757, 2.043, 4.757, { color: 'D8D8D8', transparency: 60, width: 3 });
  line(s, 1.119, 4.757, 1.62, 4.757, { color: C.magenta, width: 3 });

  plate(s, 5.35, 2.03, 3.883, 0.782);
  txt(s, runs([['Workplace Success Starts ', C.black], ['With Training', C.lime]]), { x: 5.35, y: 0.692, w: 4.358, h: 2.121 });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes', 5.386, 3.0, 4.232, 0.644);
  readMore(s, 5.386, 3.986, C.white);
}

/* ============================================ 6. compliance & policy */
function slide06() {
  const s = newSlide();
  photoRound(s, 5.258, 0.568, 4.229, 4.5, 5154);
  plate(s, 0.596, 1.697, 4.404, 0.761);
  txt(s, [
    { text: 'Digital Edition ', options: { fontFace: HEAD, fontSize: 14, color: C.black, breakLine: true } },
    ...runs([['Compliance & ', C.black], ['Policy Training', C.lime]]),
  ], { x: 0.578, y: 0.757, w: 5.0, h: 1.666 });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis', 0.562, 2.713, 4.06, 0.644);

  [
    { x: 0.578, card: C.pink,  disc: C.lime, ink: C.white, label: 'Compliance' },
    { x: 3.96,  card: C.lime,  disc: C.pink, ink: C.grey,  label: 'Policy Training' },
  ].forEach((c) => {
    roundRect(s, c.x, 3.596, 3.189, 1.175, fill(c.card), 17566);
    iconBadge(s, c.x + 0.148, 3.882, 0.602, 0.427, fill(C.limeSoft, 82), c.disc, c.ink === C.white ? C.grey : C.white);
    txt(s, c.label, { x: c.x + 0.857, y: 3.769, w: 2.184, h: 0.303, fontFace: HEAD, fontSize: 14, color: c.ink });
    body(s, 'Lorem ipsum dolor sit amet adipiscing elit. Aenean commodo', c.x + 0.857, 4.15, 2.139, 0.453, { color: c.ink });
  });
}

/* ============================================= 7. online executive training */
function slide07() {
  const s = newSlide();
  photoPath(s, P.splitLeaf, 0.535, 0.554, 2.157, 4.504);
  plate(s, 6.812, 1.282, 2.659, 0.766);
  txt(s, runs([['Online Executive ', C.black], ['Training', C.lime]]), { x: 2.915, y: 0.617, w: 6.441, h: 1.439, align: 'right' });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis', 5.145, 2.305, 3.922, 0.644);

  ['Awareness Workshops', 'Fundraising Events', 'Community Clean-Up Drives'].forEach((label, i) => {
    const y = 3.463 + i * 0.606;
    txt(s, label, { x: 5.119, y, w: 3.539, h: 0.303, fontFace: HEAD, fontSize: 14, color: C.ink });
    arrowChip(s, 8.648, y + 0.004, 0.294, C.blue, C.white, C.white);
    if (i < 2) line(s, 5.119, y + 0.454, 8.942, y + 0.454, { color: C.grey2, transparency: 78, width: 0.75 });
  });

  roundRect(s, 2.915, 1.513, 2.089, 3.555, fill(C.lime), 11181);
  iconBadge(s, 3.122, 1.736, 0.602, 0.427, fill(C.magenta, 66), C.pink, C.white);
  txt(s, 'Mobile Outreach Programs', { x: 3.122, y: 2.437, w: 1.77, h: 0.53, fontFace: HEAD, fontSize: 14 });
  body(s, [
    'Lorem ipsum dolor sit amet, consectetuer ', 'adipiscing elit. Aenean ', 'commodo ligula eget. ',
    'Aenean massa. ', 'Cum sociis natoque penatibus et magnis',
  ].map((t) => ({ text: t, options: { bullet: { indent: 9.5 }, breakLine: true } })), 3.122, 3.127, 1.675, 1.768);
}

/* ============================================== 8. photo strip + banner */
function slide08() {
  const s = newSlide();
  photoPath(s, P.banner, 5.762, 1.6, 3.727, 3.473, { rotate: 0, flipH: false });
  [0.537, 2.253, 3.969].forEach((x) => photoPath(s, P.leaf, x, 0.573, 1.548, 4.5, { rotate: 0 }));
  [[0.512, 1.598], [2.228, 1.586], [3.96, 1.548]].forEach(([x, w]) =>
    roundRect(s, x, 0.552, w, 4.516, fill(C.ink, 31), 13891));

  customPath(s, P.banner, { x: 5.762, y: 0.557, w: 3.727, h: 1.563, rotate: 180, flipH: true, fill: fill(C.pink) });
  txt(s, 'Online Executive Training', { x: 5.982, y: 0.842, w: 5.0, h: 0.303, fontFace: HEAD, fontSize: 14, color: C.lime });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis', 5.982, 1.207, 3.287, 0.644, { color: C.white });

  [1.6, 3.316, 5.032].forEach((x) => arrowChip(s, x, 0.662, 0.379, C.lime, C.white, C.white));
  arrowChip(s, 8.987, 0.652, 0.379, C.lime, C.blue, C.white);
}

/* ============================================= 9. interactive learning */
function slide09() {
  const s = newSlide();
  photoPath(s, P.leaf, 7.334, 0.562, 2.139, 4.5, { rotate: 180 });
  customPath(s, P.leaf, { x: 7.319, y: 0.562, w: 2.169, h: 4.5, rotate: 180, fill: fill(C.ink, 31) });
  customPath(s, P.leaf, { x: 5.1, y: 0.562, w: 2.139, h: 4.5, fill: fill(C.pink) });

  plate(s, 2.325, 1.542, 2.675, 0.845);
  txt(s, runs([['Interactive Online ', C.black], ['Learning ', C.lime]]), { x: 0.396, y: 0.86, w: 5.0, h: 1.439 });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes', 0.682, 2.725, 4.232, 0.644);
  readMore(s, 0.682, 3.78, C.white);

  // two ring gauges
  [
    { x: 5.592, tx: 5.308, pct: '70%', label: 'Non-Profits', track: fill(C.silver, 70), ring: C.lime,  sweep: 151.92 },
    { x: 7.811, tx: 7.527, pct: '85%', label: 'Government', track: fill(C.white, 60),  ring: C.white, sweep: 209.93 },
  ].forEach((g) => {
    s.addText(g.pct, { shape: 'ellipse', x: g.x, y: 1.49, w: 1.155, h: 1.155, fill: fill(C.white, 100),
      line: g.track === undefined ? undefined : { color: g.track.color, transparency: g.track.transparency, width: 6 },
      fontFace: HEAD, fontSize: 14, color: C.white, align: 'center', valign: 'middle', margin: INSET });
    s.addShape('arc', { x: g.x, y: 1.49, w: 1.155, h: 1.155, angleRange: [270, g.sweep], line: { color: g.ring, width: 6 } });
    txt(s, g.label, { x: g.tx, y: 3.047, w: 1.723, h: 0.303, fontFace: HEAD, fontSize: 14, color: C.white, align: 'center' });
    body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit.', g.tx, 3.389, 1.723, 0.65, { color: C.white, align: 'center' });
  });
}

/* ============================================== 10. learn together online */
function slide10() {
  const s = newSlide();
  s.addShape('rect', { x: 0, y: 0, w: 1.675, h: 5.625, fill: fill(C.pink) });
  [[2.733, 1.167, 1.588, 3.29], [4.52, 1.167, 1.588, 3.29]].forEach((r) => photoPath(s, P.leaf, r[0], r[1], r[2], r[3]));
  photoPath(s, P.leaf, 0.62, 0.823, 1.898, 3.933);
  // phone bezel that frames the first screenshot
  roundRect(s, 0.512, 0.75, 2.115, 4.125, fill(C.black), 9500);
  roundRect(s, 0.577, 0.815, 1.985, 3.995, fill(C.photo), 8500);
  roundRect(s, 1.32, 0.86, 0.5, 0.1, fill(C.black), 50000);

  plate(s, 6.307, 1.556, 2.659, 0.655);
  txt(s, runs([['Let\u2019s Learn ', C.black], ['Together ', C.lime], ['Online', C.black]]), { x: 6.307, y: 0.823, w: 4.235, h: 2.121 });
  txt(s, '222K', { x: 6.328, y: 3.053, w: 3.539, h: 0.48, fontFace: HEAD, fontSize: 24 });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis', 6.334, 3.573, 2.821, 0.847);
}

/* ==================================================== 11. image gallery */
function slide11() {
  const s = newSlide();
  plate(s, 6.028, 1.287, 2.428, 0.719);
  txt(s, runs([['Our Image ', C.ink], ['Gallery', C.lime]]), { x: 6.031, y: 0.568, w: 3.785, h: 1.439 });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque.', 6.028, 2.157, 3.241, 0.65);
  [
    [0.512, 0.613, 1.841, 2.164], [2.571, 0.568, 3.241, 2.47], [0.512, 2.91, 1.841, 2.164],
    [2.56, 3.225, 3.504, 1.832], [6.282, 3.225, 3.206, 1.832],
  ].forEach((r) => photoRound(s, r[0], r[1], r[2], r[3], 3400));
}

/* ======================================================= 12. best team */
function slide12() {
  const s = newSlide();
  plate(s, 5.408, 1.021, 1.745, 0.655);
  txt(s, runs([['Our Best ', C.ink], ['Team', C.lime]]), { x: 2.03, y: 0.968, w: 5.928, h: 0.757, align: 'center' });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis', 2.03, 1.866, 5.928, 0.453, { align: 'center' });

  [
    { x: 0.512, card: C.pink, ink: C.white, name: 'Arindra Natha',  ny: 3.017, tx: 2.711, px: 0.682, adj: 7971 },
    { x: 4.994, card: C.lime, ink: C.grey,  name: 'Kenia Lazuardhi', ny: 2.942, tx: 7.137, px: 5.17,  adj: 7488 },
  ].forEach((c) => {
    roundRect(s, c.x, 2.767, 4.324, 2.075, fill(c.card), 7060);
    photoRound(s, c.px, 2.942, 1.859, 1.725, c.adj);
    txt(s, c.name, { x: c.tx, y: c.ny, w: 2.181, h: 0.379, fontFace: HEAD, fontSize: 18, color: c.ink });
    txt(s, 'Job Position', { x: c.tx, y: c.ny + 0.381, w: 1.859, h: 0.227, color: c.ink });
    body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget', c.tx, 3.915, 2.016, 0.644, { color: c.ink });
  });
}

/* ================================================= 13. customer testimony */
function slide13() {
  const s = newSlide();
  plate(s, 4.96, 0.68, 3.065, 0.723);
  txt(s, runs([['Customer ', C.black], ['Testimony', C.lime]]), { x: 1.74, y: 0.646, w: 6.521, h: 0.757, align: 'center' });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis', 2.134, 1.546, 5.853, 0.453, { align: 'center' });

  [
    { x: 0.562, y: 2.175, h: 1.949, card: C.pink, ink: C.white, star: C.white,   name: 'Budi Santono', adj: 11178 },
    { x: 3.587, y: 2.785, h: 1.989, card: C.lime, ink: C.grey,  star: C.limeDark, name: 'Siti Wahyuni', adj: 10002 },
    { x: 6.612, y: 2.175, h: 1.949, card: C.pink, ink: C.white, star: 'F2F2F2',  name: 'Yanto Triono', adj: 10394 },
  ].forEach((c) => {
    roundRect(s, c.x, c.y, 2.825, c.h, fill(c.card), c.adj);
    slidePhotoCircle(s, c.x + 0.188, c.y + 0.145, 0.51);
    txt(s, c.name, { x: c.x + 0.789, y: c.y + 0.081, w: 1.495, h: 0.303, fontFace: HEAD, fontSize: 14, color: c.ink });
    txt(s, 'Your Company Here', { x: c.x + 0.789, y: c.y + 0.383, w: 1.566, h: 0.227, color: c.ink });
    body(s, 'Lorem ipsum dolor sit amet, adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. ', c.x + 0.188, c.y + 0.733, 2.45, 0.65, { color: c.ink });
    for (let i = 0; i < 5; i++) {
      s.addShape('star5', { x: c.x + 0.254 + i * 0.1376, y: c.y + 1.638, w: 0.105, h: 0.105, fill: fill(c.star) });
    }
  });
}
function slidePhotoCircle(s, x, y, d) {
  s.addShape('ellipse', { x, y, w: d, h: d, fill: fill(C.photo) });
}

/* ================================================= 14. learning path rings */
function slide14() {
  const s = newSlide();
  [[2.389, 1.698, 2.218], [1.455, 0.764, 4.086]].forEach(([x, y, d]) =>
    s.addShape('ellipse', { x, y, w: d, h: d, fill: { type: 'none' }, line: { color: C.silver, width: 1 } }));

  // percentage chips arranged around the rings: [chipX, chipY, card, ink, pctColour, pctX, pctAlign, textX, textW, dotX, dotY]
  [
    { x: 3.631, y: 0.807, card: C.lime, ink: C.grey,  pct: '25%', pctCol: C.grey, pctX: 3.714, tX: 4.462, tW: 1.28,  tAlign: 'left',  dot: [3.549, 0.738, C.lime] },
    { x: 0.433, y: 1.491, card: C.pink, ink: C.white, pct: '76%', pctCol: C.lime, pctX: 1.805, tX: 0.519, tW: 1.286, tAlign: 'right', dot: [2.636, 2.034, C.pink] },
    { x: 2.004, y: 3.144, card: C.lime, ink: C.grey,  pct: '63%', pctCol: C.grey, pctX: 2.087, tX: 2.835, tW: 1.277, tAlign: 'left',  dot: [4.269, 3.507, C.lime] },
    { x: 2.389, y: 4.069, card: C.pink, ink: C.white, pct: '97%', pctCol: C.lime, pctX: 2.472, tX: 3.22,  tW: 1.346, tAlign: 'left',  dot: [2.205, 4.389, C.pink] },
  ].forEach((c) => {
    roundRect(s, c.x, c.y, 2.177, 0.584, fill(c.card), 17566);
    txt(s, c.pct, { x: c.pctX, y: c.y + 0.103, w: 0.719, h: 0.379, fontFace: HEAD, fontSize: 18, color: c.pctCol, align: 'right' });
    body(s, 'Lorem ipsum dolor sit amet adipiscing', c.tX, c.y + 0.068, c.tW, 0.453, { color: c.ink, align: c.tAlign });
    s.addShape('ellipse', { x: c.dot[0], y: c.dot[1], w: 0.083, h: 0.083, fill: fill(c.dot[2]) });
  });

  plate(s, 5.761, 2.97, 1.664, 0.655);
  txt(s, runs([['Launching Your Learning ', C.black], ['Path', C.lime]]), { x: 5.808, y: 1.594, w: 4.086, h: 2.121 });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis', 5.808, 3.98, 3.716, 0.644);
}

/* ============================================ 15. training for performance */
function slide15() {
  const s = newSlide();
  // three identical bar panels, only the month labels change
  const rows = [
    { seg: [0.139, 0.550, 0.130] },
    { seg: [0.819, 0.405, 0.556] },
    { seg: [0.139, 0.680, 0.280] },
    { seg: [0.139, 0.139, 0.268] },
  ];
  const segColour = [C.magenta, C.lime, C.limeDark2];
  const panels = [
    { x: 0.562, bar: 1.276, months: ['Apr', 'Marc', 'Feb', 'Jan'], pill: [1.030, C.magenta,  'Future-Oriented', C.white] },
    { x: 3.570, bar: 4.118, months: ['Aug', 'Jul', 'Jun', 'May'],  pill: [4.038, C.lime,     'Data and Trends', C.grey] },
    { x: 6.578, bar: 7.100, months: ['Dec', 'Nop', 'Oct', 'Sep'],  pill: [7.045, C.limeDark, 'Decision-Making', C.white] },
  ];
  panels.forEach((p) => {
    roundRect(s, p.x, 2.264, 2.859, 2.027, fill(C.panel), 7653);
    line(s, p.bar, 2.46, p.bar, 3.99, { color: C.grey3, width: 0.75 });
    rows.forEach((r, ri) => {
      const y = 2.543 + ri * 0.3915;
      txt(s, p.months[ri], { x: p.bar - 0.61, y: y - 0.028, w: 0.535, h: 0.227, fontFace: HEAD, color: C.grey2, align: 'right' });
      let x = p.bar;
      r.seg.forEach((w, si) => { s.addShape('rect', { x, y, w, h: 0.172, fill: fill(segColour[si]) }); x += w; });
    });
    roundRect(s, p.pill[0], 4.109, 1.925, 0.335, fill(p.pill[1]), 50000);
    txt(s, p.pill[2], { x: p.pill[0] + 0.134, y: 4.15, w: 1.656, h: 0.252, fontFace: HEAD, fontSize: 11, color: p.pill[3], align: 'center' });
  });

  plate(s, 3.336, 1.502, 3.344, 0.655);
  txt(s, [
    { text: 'Online Edition ', options: { fontFace: HEAD, fontSize: 14, color: C.black, breakLine: true, align: 'center' } },
    ...runs([['Training For Performance ', C.black], ['Excellence', C.lime]]),
  ], { x: 1.028, y: 0.561, w: 7.941, h: 1.666, align: 'center' });
  body(s, '\u201CLorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis\u201D',
    2.221, 4.576, 5.555, 0.421, { fontSize: 8, align: 'center', italic: true });
}

/* ==================================================== 16. sharpen your edge */
function slide16() {
  const s = newSlide();
  [
    { x: 0.522, y: 3.374, h: 1.683, colour: C.pink,     pct: '49.8%' },
    { x: 1.985, y: 2.897, h: 2.160, colour: C.magenta,  pct: '61.8%' },
    { x: 3.451, y: 2.307, h: 2.750, colour: C.limeDark, pct: '70.3%' },
    { x: 4.916, y: 1.916, h: 3.142, colour: C.lime,     pct: '87.5%' },
  ].forEach((b) => {
    roundRect(s, b.x, b.y, 1.245, b.h, fill(b.colour), 16667, { flipH: true });
    line(s, b.x, b.y - 1.151, b.x, b.y, { color: C.silver, width: 0.75 });
    txt(s, b.pct, { x: b.x + 0.103, y: b.y - 1.051, w: 1.05, h: 0.303, fontFace: HEAD, fontSize: 14, color: C.ink });
    body(s, 'Lorem ipsum dolor sit ', b.x + 0.102, b.y - 0.732, 1.036, 0.447, { color: C.grey2 });
  });

  plate(s, 6.523, 1.912, 2.988, 0.797);
  txt(s, runs([['Sharpen ', C.black], ['Your Edge', C.lime]]), { x: 5.881, y: 1.221, w: 3.619, h: 1.439, align: 'right' });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis',
    6.588, 2.965, 2.912, 1.044, { align: 'right' });
}

/* ==================================================== 17. online course intro */
function slide17() {
  const s = newSlide();
  plate(s, 6.277, 0.735, 1.622, 0.655);
  txt(s, runs([['Online Course ', C.black], ['Intro', C.lime]]), { x: 1.934, y: 0.713, w: 6.132, h: 0.757, align: 'center' });
  txt(s, 'Lorem ipsum dolor sit amet, adipiscing elit. Aenean commodo ligula eget', { x: 4.484, y: 1.879, w: 5.102, h: 0.227, color: C.grey2 });

  // three KPI cards
  [
    { y: 1.826, colour: C.limeDark, ink: C.white, value: '+71%' },
    { y: 2.951, colour: C.lime,     ink: C.grey,  value: '>820K' },
    { y: 4.076, colour: C.pink,     ink: C.white, value: '10M' },
  ].forEach((c) => {
    roundRect(s, 0.568, c.y, 3.258, 0.987, fill(c.colour), 16667);
    txt(s, c.value, { x: 0.724, y: c.y + 0.167, w: 1.495, h: 0.379, fontFace: HEAD, fontSize: 18, color: c.ink });
    txt(s, 'Lorem ipsum dolor amet', { x: 0.724, y: c.y + 0.567, w: 1.875, h: 0.252, fontSize: 11, color: c.ink });
  });

  // card 1: white zig-zag sparkline
  const zig = [[2.817, 2.489], [2.897, 2.224], [3.016, 2.496], [3.124, 2.388], [3.220, 2.471], [3.243, 2.147],
    [3.370, 2.471], [3.391, 2.138], [3.535, 2.318], [3.606, 2.272]];
  zig.forEach((pt, i) => {
    if (i) line(s, zig[i - 1][0], zig[i - 1][1], pt[0], pt[1], { color: C.white, width: 1.25 });
    s.addShape('ellipse', { x: pt[0] - 0.012, y: pt[1] - 0.012, w: 0.024, h: 0.025, fill: fill(C.white) });
  });
  // card 2: bar sparkline
  [[2.800, 3.518, 0.208], [2.932, 3.474, 0.252], [3.063, 3.518, 0.208], [3.195, 3.177, 0.549],
    [3.326, 3.330, 0.396], [3.457, 3.566, 0.160], [3.589, 3.680, 0.047]].forEach((b, i) =>
    roundRect(s, b[0], b[1], 0.086, b[2], i === 3 ? fill(C.white) : fill('EEFBD3', 54), 16667));
  // card 3: donut gauge
  s.addShape('blockArc', { x: 2.913, y: 4.246, w: 0.647, h: 0.647, fill: fill(C.pinkSoft), flipH: true,
    angleRange: [162.1, 270], arcThicknessRatio: 0.251 });
  s.addShape('blockArc', { x: 2.913, y: 4.246, w: 0.647, h: 0.647, fill: fill(C.blue),
    angleRange: [18, 270], arcThicknessRatio: 0.251 });
  txt(s, '70%', { x: 2.975, y: 4.456, w: 0.526, h: 0.227, fontFace: HEAD, color: C.white, align: 'center' });

  // right-hand line plot, drawn the same way as in the source deck
  const PX0 = 4.77, PDX = 1.5265;      // x = 0 and one unit of x
  const PY0 = 4.6305, PDY = 0.5593;    // y = 0 and one unit of y
  const px = (v) => PX0 + v * PDX;
  const py = (v) => PY0 - v * PDY;
  for (let i = 0; i <= 8; i++) {                       // horizontal gridlines + y labels
    const y = 2.3935 + i * 0.2796;
    line(s, 4.77, y, 9.35, y, { color: C.hair, width: 0.75, dashType: i === 8 ? 'solid' : 'dash' });
    line(s, 4.726, y, 4.77, y, { color: C.hair, width: 0.75 });
    txt(s, ['4', '3,5', '3', '2,5', '2', '1,5', '1', '0,5', '0'][i], { x: 4.257, y: y - 0.119, w: 0.46, h: 0.227, color: C.grey2, align: 'right' });
  }
  line(s, 3.629, 4.675, 3.629, 2.393, { color: C.hair, width: 0.75 });
  ['0', '0,5', '1', '1,5', '2', '2,5'].forEach((lab, i) => {                // x labels + ticks
    const x = px(i * 0.5);
    if (i) line(s, x, 4.655, x, 4.696, { color: C.hair, width: 0.75 });
    txt(s, lab, { x: x - 0.25, y: 4.699, w: 0.5, h: 0.227, color: C.grey2, align: 'center' });
  });
  [
    { colour: C.magenta, pts: [[0, 0], [1, 1], [2, 1], [2.6, 3]] },
    { colour: C.lime,    pts: [[0, 0], [1, 2], [2, 2], [2.6, 1]] },
    { colour: C.blue,    pts: [[0, 0], [1, 3], [2, 3.1], [2.6, 3.5]] },
  ].forEach((sr) => {
    sr.pts.forEach((p, i) => {
      if (i) line(s, px(sr.pts[i - 1][0]), py(sr.pts[i - 1][1]), px(p[0]), py(p[1]), { color: sr.colour, width: 1.5 });
    });
    sr.pts.forEach((p) => s.addShape('ellipse', { x: px(p[0]) - 0.0515, y: py(p[1]) - 0.0515, w: 0.103, h: 0.103,
      fill: fill(sr.colour), line: { color: sr.colour, width: 0.75 } }));
  });
}

/* ======================================================= 18. get ready to learn */
function slide18() {
  const s = newSlide();
  // nested TAM / SAM / SOM bubbles
  [
    { x: 5.286, y: 0.781, d: 4.062, colour: C.pink,     tag: 'TAM', value: '$512m', ty: 0.989, tw: 1.841, tx: 6.397 },
    { x: 5.867, y: 1.943, d: 2.901, colour: C.lime,     tag: 'SAM', value: '$333m', ty: 2.249, tw: 1.841, tx: 6.397 },
    { x: 6.424, y: 3.058, d: 1.786, colour: C.limeDark, tag: 'SOM', value: '$65m',  ty: 3.630, tw: 1.521, tx: 6.557 },
  ].forEach((b) => {
    s.addShape('ellipse', { x: b.x, y: b.y, w: b.d, h: b.d, fill: fill(b.colour) });
    txt(s, b.tag, { x: b.tx, y: b.ty, w: b.tw, h: 0.278, fontFace: HEAD, fontSize: 12, color: C.white, align: 'center', valign: 'bottom' });
    txt(s, b.value, { x: b.tx, y: b.ty + 0.219, w: b.tw, h: 0.429, fontFace: HEAD, fontSize: 21, color: C.white, align: 'center' });
  });

  plate(s, 0.831, 1.343, 1.856, 0.655);
  txt(s, runs([['Get Ready To ', C.black], ['Learn', C.lime]]), { x: 0.867, y: 0.604, w: 5.0, h: 1.439 });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis', 0.875, 2.132, 3.905, 0.453);
  txt(s, 'Proposed Allocation', { x: 1.165, y: 2.698, w: 2.475, h: 0.278, fontFace: HEAD, fontSize: 12, color: C.ink, align: 'center', valign: 'bottom' });

  [1.185, 2.385, 3.586].forEach((x) => s.addShape('ellipse', { x, y: 3.127, w: 0.045, h: 0.045, fill: fill('D8D8D8') }));
  [1.322, 2.523].forEach((x) => line(s, x, 3.15, x + 0.971, 3.15, { color: '121214', transparency: 30, width: 0.75, dashType: 'dash' }));

  [
    { x: 0.651, colour: C.pink,     ink: C.white, value: '$125', adj: 9168, glyph: C.white },
    { x: 1.852, colour: C.lime,     ink: C.grey,  value: '$678', adj: 8231, glyph: C.white },
    { x: 3.053, colour: C.limeDark, ink: C.white, value: '$24',  adj: 9637, glyph: C.white },
  ].forEach((c) => {
    roundRect(s, c.x, 3.604, 1.111, 1.206, fill(c.colour), c.adj);
    s.addShape('ellipse', { x: c.x + 0.277, y: 3.325, w: 0.559, h: 0.559, fill: fill(C.white) });
    s.addShape('ellipse', { x: c.x + 0.343, y: 3.391, w: 0.425, h: 0.425, fill: fill(c.colour) });
    icon(s, c.x + 0.489, 3.537, 0.135, c.glyph);
    body(s, 'Your Statement', c.x + 0.022, 3.936, 1.068, 0.256, { color: c.ink, align: 'center' });
    txt(s, c.value, { x: c.x + 0.022, y: 4.436, w: 1.068, h: 0.278, fontFace: HEAD, fontSize: 12, color: c.ink, align: 'center' });
  });
}

/* ==================================================== 19. grow your knowledge */
function slide19() {
  const s = newSlide();
  customPath(s, P.hill, { x: 3.645, y: 2.231, w: 5.844, h: 2.826, fill: fill(C.panel) });
  txt(s, '678K', { x: 3.755, y: 2.46, w: 1.297, h: 0.48, fontFace: HEAD, fontSize: 24, color: 'EC2ECF' });
  txt(s, 'A/B Testing ', { x: 4.638, y: 2.603, w: 2.063, h: 0.278, fontFace: HEAD, fontSize: 12, color: C.ink, valign: 'bottom' });
  body(s, 'Lorem ipsum dolor consectetuer adipiscing elit aenean', 3.755, 3.019, 2.195, 0.453);

  [
    { x: 4.504, top: 4.331, dot: 4.102, colour: C.pink,     pct: '43%', ink: C.white },
    { x: 5.913, top: 4.102, dot: 3.838, colour: C.magenta,  pct: '52%', ink: C.white },
    { x: 7.322, top: 3.661, dot: 3.374, colour: C.limeDark, pct: '65%', ink: C.white },
    { x: 8.629, top: 2.969, dot: 2.689, colour: C.lime,     pct: '87%', ink: C.grey },
  ].forEach((m) => {
    line(s, m.x, m.top, m.x, 5.058, { color: C.silver, width: 0.75 });
    s.addShape('ellipse', { x: m.x - 0.197, y: m.dot, w: 0.394, h: 0.394, fill: fill(m.colour) });
    txt(s, m.pct, { x: m.x - 0.197, y: m.dot + 0.096, w: 0.394, h: 0.202, fontFace: HEAD, fontSize: 8, color: m.ink, align: 'center' });
  });

  txt(s, 'PLACEHOLDER', { x: 0.648, y: 2.463, w: 2.52, h: 0.757, fontFace: HEAD, fontSize: 14 });
  roundRect(s, 0.555, 3.381, 2.613, 0.607, fill(C.lime), 17515);
  body(s, 'Lorem ipsum dolor sit amet, adipiscing elit. Aenean commodo', 0.711, 3.459, 2.324, 0.453);
  body(s, ['Lorem ipsum dolor sit consectetuer. ', 'Aenean commodo ligula eget dolor. ', 'Cum sociis natoque penatibus et magnis ']
    .map((t) => ({ text: t, options: { bullet: { indent: 9.5 }, breakLine: true } })), 0.648, 4.202, 2.52, 0.847);

  plate(s, 4.846, 0.917, 3.461, 0.744);
  txt(s, runs([['Grow Your ', C.black], ['Knowledge', C.lime]]), { x: 1.641, y: 0.894, w: 6.718, h: 0.757, align: 'center' });
}

/* ============================================================= 20. thank you */
function slide20() {
  const s = newSlide();
  s.addShape('rect', { x: 0, y: 0, w: 2.506, h: 5.625, fill: fill(C.pink) });
  photoRect(s, 0.607, 0.568, 3.599, 4.49);
  plate(s, 7.613, 1.155, 2.061, 0.975);
  txt(s, [
    ...runs([['Thank ', C.black], ['You', C.lime], [' ', C.black]], 72),
    { text: '', options: { breakLine: true } },
    { text: 'For Your Attention', options: { fontFace: HEAD, fontSize: 18, color: C.ink } },
  ], { x: 4.347, y: 1.028, w: 5.549, h: 1.641 });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis', 4.381, 2.774, 4.543, 0.65);
  ['Pringle Drive Chicago, IL 60606', 'LaverneRGilmore@armyspy.com', 'www.northact.com'].forEach((t, i) =>
    txt(s, t, { x: 4.381, y: 3.712 + i * 0.3575, w: 2.922, h: 0.252, fontFace: HEAD, fontSize: 11, color: C.ink }));
}

/* -------------------------------------------------------------------- build */
[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
 slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
  .forEach((fn) => fn());

pres.writeFile({ fileName: path.join(__dirname, '081beaf4-c82f-432d-9844-18df0ba40c7b_grok_final.pptx') })
  .then((f) => console.log('wrote ' + f));
