/**
 * "Living Real Estate" — 16-slide deck rebuilt with pptxgenjs.
 *
 * Slide size 13.333 x 7.5 in (16:9). Theme "Custom 123" / fonts Raleway + Poppins.
 * Photographs in the original are replaced with grey "[image]" placeholder blocks.
 *
 * Run:  node 09698e65-4467-4e99-8725-6b0a2bfa9387_grok_final.js
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ---------------------------------------------------------------- palette */

const DARK = '0C3A2D'; // accent1 - deep forest green
const GREEN = '6D9773'; // accent2 - sage green
const TAN = 'BB8A52'; // accent3
const TEAL = '19525B'; // accent4
const WHITE = 'FFFFFF';
const BLACK = '000000';
const PHOTO_BG = 'D9D9D9'; // stand-in for photographs
const PHOTO_FG = 'AEAEAE';

const HEAD = 'Raleway'; // theme major font
const BODY = 'Poppins'; // theme minor font

const NOLINE = { type: 'none' };

/* ------------------------------------------------------------ lorem bank */

const L = {
  short: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ',
  row: 'Lorem ipsum ligula dolor sit amet',
  elitLong:
    'Lorem elit ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque',
  elitMid: 'Lorem elit ipsum dolor sit amet, ligula consectetuer adipiscing elit. Aenean',
  elitCum:
    'Lorem elit ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum',
  natoque:
    'Lorem ipsum ligula dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. ligula Cum sociis natoque',
  penatibusFull:
    'Lorem ipsum ligula dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. ligula Cum sociis natoque penatibus ligula dolor sit amet, ligula dolor sit amet, consectetuer adipiscing',
  penatibusShort:
    'Lorem ipsum ligula dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. ligula Cum sociis natoque penatibus ligula dolor sit amet, ligula',
  penatibusSit:
    'Lorem ipsum ligula dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. ligula Cum sociis natoque penatibus ligula dolor sit amet, ligula dolor sit',
  penatibusComma:
    'Lorem ipsum ligula dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. ligula Cum sociis natoque penatibus ligula dolor sit amet, ',
  penatibusEnd:
    'Lorem ipsum ligula dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. ligula Cum sociis natoque penatibus',
  dolorEnd: 'Lorem ipsum ligula dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. ',
  elitEnd: 'Lorem ipsum ligula dolor sit amet, consectetuer adipiscing elit. ',
  team: 'Lorem ipsum ligula dolor sit amet, ligula consectetuer adipiscing Aenean. Aenean',
  icon: 'Lorem ipsum ligula dolor sit amet, consectetuer adipiscing',
  egetLong:
    'Lorem ipsum ligula dolor sit amet, consectetuer eget adipiscing elit. Aenean commodo eget ligula eget dolor. Aenean massa. ligula Cum sociis natoque',
  egetShort:
    'Lorem ipsum ligula dolor sit amet, consectetuer eget adipiscing elit. Aenean commodo eget ligula eget dolor. Aenean',
  eget: 'Lorem ipsum ligula dolor sit amet, consectetuer eget',
  card: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula',
  quarter: 'Lorem sit ipsum ligula dolor sit amet, sit consectetuer eget adipiscing elit. Aenean',
  tick: ['amet, consectetuer eget', 'Lorem ipsum ligula dolor sit ', 'Lorem ipsum ligula dolor sit '],
  goal: ['Lorem ipsum ligula dolor sit amet, consectetuer', ' adipiscing elit. Aenean commodo ligula eget ', 'dolor. Aenean massa. ligula Cum sociis natoque'],
};

const TAGLINE = 'Creating Spaces That Combine Elegance, Comfort, and Value';
const TITLE_HERE = 'Your Title Here';
const TITLE_here = 'Your Title here';

/* ---------------------------------------------------------------- helpers */

/** Body copy: Poppins 12pt, top-anchored like a PowerPoint text box. */
function body(slide, text, o) {
  slide.addText(text, Object.assign({ fontFace: BODY, fontSize: 12, color: BLACK, valign: 'top', lineSpacingMultiple: 1.3 }, o));
}

/** Display copy: Raleway, top-anchored. */
function head(slide, text, o) {
  slide.addText(text, Object.assign({ fontFace: HEAD, fontSize: 54, color: BLACK, valign: 'top' }, o));
}

/** Solid block of colour. */
function block(slide, x, y, w, h, color) {
  slide.addShape('rect', { x, y, w, h, fill: { color }, line: NOLINE });
}

/** Rounded card. `adj` is the OOXML corner adjustment (fraction of the short side). */
function card(slide, x, y, w, h, color, adj) {
  slide.addShape('roundRect', { x, y, w, h, fill: { color }, line: NOLINE, rectRadius: adj * Math.min(w, h) });
}

/** Hairline rule. */
function rule(slide, x, y, w, h, color, width) {
  slide.addShape('line', { x, y, w, h, line: { color, width: width || 0.5 } });
}

/** Grey stand-in for a photograph. */
function photo(slide, x, y, w, h, opts) {
  const o = opts || {};
  slide.addShape(o.shape || 'rect', { x, y, w, h, fill: { color: PHOTO_BG }, line: NOLINE, rectRadius: o.radius });
  slide.addText('[image]', { x, y, w, h, align: 'center', valign: 'middle', fontFace: BODY, fontSize: 11, color: PHOTO_FG });
}

/** Check-mark bullets used throughout the deck (Wingdings "ü" in the original). */
const TICK = { characterCode: '2713', indent: 13.5 };

/**
 * Small pictograms standing in for the deck's SVG icons.
 * `bg` is the colour behind the icon, used to knock details out of solid shapes.
 */
function icon(slide, kind, x, y, s, color, bg) {
  const fill = { color }, ln = NOLINE, hole = { color: bg };
  if (kind === 'house') {
    slide.addShape('triangle', { x, y, w: s, h: s * 0.46, fill, line: ln });
    slide.addShape('rect', { x: x + s * 0.19, y: y + s * 0.42, w: s * 0.62, h: s * 0.58, fill, line: ln });
    slide.addShape('rect', { x: x + s * 0.41, y: y + s * 0.66, w: s * 0.18, h: s * 0.34, fill: hole, line: ln });
  } else if (kind === 'coins') {
    [0, 0.19, 0.38].forEach(dy => slide.addShape('ellipse', { x, y: y + s * dy, w: s * 0.66, h: s * 0.22, fill, line: ln }));
    slide.addShape('ellipse', { x: x + s * 0.40, y: y + s * 0.40, w: s * 0.60, h: s * 0.60, fill, line: ln });
    slide.addText('$', { x: x + s * 0.40, y: y + s * 0.40, w: s * 0.60, h: s * 0.60, align: 'center', valign: 'middle', fontFace: BODY, fontSize: 7, bold: true, color: bg });
  } else if (kind === 'percent') {
    slide.addShape('star16', { x, y, w: s, h: s, fill, line: ln });
    slide.addText('%', { x, y, w: s, h: s, align: 'center', valign: 'middle', fontFace: BODY, fontSize: 9, bold: true, color: bg });
  } else if (kind === 'target') {
    // Concentric rings: a white disc with two knocked-out rings in the background colour.
    slide.addShape('ellipse', { x, y, w: s, h: s, fill, line: ln });
    slide.addShape('donut', { x: x + s * 0.11, y: y + s * 0.11, w: s * 0.78, h: s * 0.78, fill: hole, line: ln });
    slide.addShape('donut', { x: x + s * 0.30, y: y + s * 0.30, w: s * 0.40, h: s * 0.40, fill: hole, line: ln });
  }
}

/** Footer that the slide master paints on every slide. */
function footer(slide) {
  slide.addText('Living Real Estate', { x: 0.4069, y: 6.8965, w: 1.9196, h: 0.3366, fontFace: HEAD, fontSize: 14, color: BLACK, valign: 'top' });
}

/* ------------------------------------------------------------- 1: cover */

function slide01(slide) {
  photo(slide, 8.4253, 0, 4.908, 7.5);
  block(slide, 0, 4.0634, 8.4261, 1.9444, GREEN);
  head(slide, 'Real Estate', { x: 1.4273, y: 4.2446, w: 6.8615, h: 1.582, fontSize: 88, color: WHITE });
  head(slide, 'Living', { x: 1.4273, y: 2.4622, w: 4.2722, h: 1.582, fontSize: 88 });
  block(slide, 0, 0, 1.2435, 4.0634, DARK);
  body(slide, L.short, { x: 5.4534, y: 3.1328, w: 2.7632, h: 0.6075, align: 'justify' });
}

/* ------------------------------------------------- 2: agenda / welcome */

function slide02(slide) {
  block(slide, 0, 0, 13.3333, 0.9126, DARK);
  photo(slide, 0, 0.5938, 13.3333, 3.6441);
  head(slide, 'Welcome to Living Real Estate', { x: 5.7018, y: 4.8839, w: 7.3073, h: 1.9186, align: 'right' });
  block(slide, 0.5195, 2.4286, 4.8776, 5.0714, GREEN);

  // Four numbered agenda rows inside the green panel.
  ['01.', '02.', '03.', '04.'].forEach((num, i) => {
    const top = 3.038 + i * 0.7437;
    slide.addText(num, { x: 0.7578, y: top, w: 0.6129, h: 0.3777, fontFace: HEAD, fontSize: 14, color: WHITE, valign: 'top', lineSpacingMultiple: 1.3 });
    body(slide, L.row, { x: 1.4467, y: top + 0.0136, w: 3.1651, h: 0.345, color: WHITE });
    slide.addShape('triangle', { x: 4.9785, y: top + 0.1286, w: 0.1333, h: 0.1149, rotate: 180, fill: { color: WHITE }, line: NOLINE });
    rule(slide, 0.8418, top + 0.3931, 4.3168, 0, WHITE, 1);
  });

  body(slide, L.elitLong, { x: 0.7578, y: 5.9066, w: 4.4008, h: 0.8701, color: WHITE, align: 'justify' });
  head(slide, '2025', { x: 5.7478, y: 1.9684, w: 7.3073, h: 2.8947, fontSize: 166, bold: true, color: WHITE, align: 'right' });
  body(slide, TAGLINE, { x: 5.9165, y: 6.0379, w: 2.9772, h: 0.6075, align: 'right' });
}

/* ---------------------------------------------------- 3: vision + stats */

function slide03(slide) {
  block(slide, 0, 3.75, 10.2708, 2.9541, DARK);
  head(slide, 'Our Vision for Modern Living', { x: 7.5093, y: 1.1546, w: 5.9649, h: 1.9186 });

  const stats = [
    { x: 0.7148, rule: 0.551, ruleColor: DARK, value: '30+', w: 1.6406 },
    { x: 3.9079, rule: 3.5768, ruleColor: WHITE, value: '130K', w: 2.1329 },
    { x: 7.101, rule: 6.77, ruleColor: WHITE, value: '4,9', w: 1.6406 },
  ];
  stats.forEach(s => {
    head(slide, s.value, { x: s.x, y: 4.3463, w: s.w, h: 1.0098, color: WHITE });
    body(slide, TITLE_here, { x: s.x, y: 4.1501, w: 1.4722, h: 0.345, color: WHITE });
    body(slide, L.elitMid, { x: s.x, y: 5.4339, w: 2.5939, h: 0.8701, color: WHITE, align: 'justify' });
    rule(slide, s.rule, 3.9745, 0, 2.4949, s.ruleColor);
  });

  photo(slide, 0.5503, 0.9253, 6.4219, 2.3767);
  photo(slide, 10.2708, 3.75, 3.0625, 3.75);
}

/* ------------------------------------------------ 4: featured properties */

function slide04(slide) {
  block(slide, 5.3571, 0, 7.9762, 4.0714, DARK);
  head(slide, 'Featured Properties and Premium Locations', { x: 5.751, y: 0.7547, w: 7.1885, h: 2.8273, color: WHITE });
  body(slide, TAGLINE, { x: 9.6716, y: 2.7419, w: 2.9772, h: 0.6075, color: WHITE });
  head(slide, '30%', { x: 5.751, y: 4.4008, w: 1.6406, h: 1.0098 });
  head(slide, '130+', { x: 5.751, y: 5.8999, w: 2.1329, h: 1.0098 });
  body(slide, L.elitCum, { x: 7.6558, y: 4.6019, w: 5.3952, h: 0.6075 });
  body(slide, L.elitCum, { x: 7.6558, y: 6.101, w: 5.3952, h: 0.6075 });
  rule(slide, 5.6735, 5.6939, 7.6599, 0, DARK, 1);
  photo(slide, 0, 0, 5.3576, 7.5);
}

/* --------------------------------------------------- 5: design philosophy */

function slide05(slide) {
  photo(slide, 0, 0, 13.3333, 7.5);
  head(slide, 'Design Philosophy and Interior Concepts', { x: 0.6943, y: 0.6779, w: 6.8474, h: 2.8273 });
  icon(slide, 'house', 8.1926, 1.0686, 0.4737, BLACK, PHOTO_BG);
  body(slide, L.elitLong, { x: 8.1321, y: 1.6565, w: 4.4008, h: 0.8701, align: 'justify' });
  body(slide, 'How We Bring Comfort and Aesthetics Together', { x: 8.8439, y: 1.0017, w: 2.9772, h: 0.6075 });
}

/* --------------------------------------------------- 6: monthly report */

function slide06(slide) {
  head(slide, 'Monthly Business Report Titles', { x: 6.2568, y: 1.0737, w: 7.1885, h: 1.9186 });
  block(slide, 0, 5.8333, 5.7812, 1.6667, DARK);
  body(slide, L.penatibusFull, { x: 6.2568, y: 4.8125, w: 6.7432, h: 0.8701, align: 'justify' });
  body(slide, L.natoque, { x: 6.2568, y: 5.8333, w: 6.7432, h: 0.6075, align: 'justify' });
  head(slide, '290K', { x: 6.2568, y: 3.652, w: 2.4724, h: 1.0098 });
  body(slide, TITLE_HERE, { x: 8.1701, y: 4.1569, w: 1.4583, h: 0.345, italic: true });
  head(slide, 'Real Estate That Delivers Long-Term Financial Benefits', { x: 0.7934, y: 6.3132, w: 3.6945, h: 0.7068, fontSize: 18, color: WHITE });
  photo(slide, 0, 0, 5.7812, 5.8333);
}

/* ------------------------------------------------- 7: smart living goals */

function slide07(slide) {
  block(slide, 4.0938, 4.3802, 8.8229, 2.5312, DARK);
  head(slide, 'Smart Living Features', { x: 4.413, y: 0.8163, w: 4.6911, h: 1.9186 });
  body(slide, L.natoque, { x: 4.413, y: 2.824, w: 4.5072, h: 0.8701, align: 'justify' });
  rule(slide, 4.0937, 5.6458, 8.8229, 0, WHITE, 1);

  ['01. Goal health', '02. Goal health'].forEach((label, i) => {
    head(slide, label, { x: 4.9938, y: 4.823 + i * 1.2417, w: 2.0245, h: 0.4039, fontSize: 18, color: WHITE });
    slide.addText(L.goal.map(t => ({ text: t, options: { bullet: TICK, breakLine: true } })), {
      x: 7.5094, y: 4.59 + i * 1.2416, w: 4.5072, h: 0.8701,
      fontFace: BODY, fontSize: 12, color: WHITE, valign: 'top', align: 'justify', lineSpacingMultiple: 1.3,
    });
  });

  photo(slide, 0.4167, 0.5885, 3.6771, 6.3229);
  photo(slide, 9.2396, 0.5885, 3.6771, 3.7917);
}

/* --------------------------------------------------------- 8: our team */

function slide08(slide) {
  head(slide, 'Our Team', { x: 0.8297, y: 1.18, w: 4.0036, h: 1.0098 });
  body(slide, L.penatibusShort, { x: 0.8297, y: 2.3104, w: 5.4203, h: 0.8701, align: 'justify' });
  block(slide, 0.6701, 3.5625, 5.9966, 2.767, DARK);
  block(slide, 6.6667, 0.7967, 5.9966, 2.767, GREEN);
  body(slide, L.penatibusSit, { x: 7.0068, y: 4.9719, w: 5.6565, h: 0.8701, align: 'justify' });
  head(slide, '12K', { x: 7.0068, y: 4.0501, w: 1.7328, h: 1.0098 });
  body(slide, TITLE_HERE, { x: 8.4617, y: 4.5549, w: 1.4583, h: 0.345, italic: true });

  // Two identical "person" cards, offset into the dark and the green panel.
  [{ x: 0.8297, y: 3.7909 }, { x: 6.8991, y: 1.0251 }].forEach(p => {
    head(slide, 'Madison Marques', { x: p.x, y: p.y, w: 2.2588, h: 0.7068, fontSize: 18, color: WHITE });
    body(slide, 'Your Position Here', { x: p.x, y: p.y + 0.3509, w: 1.8317, h: 0.3247, fontSize: 11, italic: true, color: WHITE, align: 'justify' });
    body(slide, L.team, { x: p.x, y: p.y + 1.4401, w: 2.5314, h: 0.8701, color: WHITE, align: 'justify' });
  });

  photo(slide, 3.9236, 3.5642, 2.7431, 2.7656);
  photo(slide, 9.9201, 0.7969, 2.7431, 2.7656);
}

/* -------------------------------------------------- 9: smartphone mockup */

function slide09(slide) {
  block(slide, 0.6771, 0.894, 5.2812, 5.7119, DARK);

  // Phone mock-up: side buttons, chassis, bezel, screen, notch with speaker + camera.
  [[1.486, 0.208], [1.930, 0.431], [2.500, 0.417]].forEach(([y, h]) => block(slide, 5.0981, y, 0.0492, h, '919191'));
  block(slide, 8.1849, 2.087, 0.0503, 0.6961, '919191');
  card(slide, 5.1236, 0.624, 3.085, 6.252, '414041', 0.14);
  card(slide, 5.1705, 0.671, 2.9915, 6.1577, '131313', 0.13);
  photo(slide, 5.3194, 0.8194, 2.6806, 5.8472, { shape: 'roundRect', radius: 0.3 });
  slide.addShape('round2SameRect', { x: 5.875, y: 0.70, w: 1.569, h: 0.30, fill: { color: '131313' }, line: NOLINE, rectRadius: 0.09, flipV: true });
  slide.addShape('roundRect', { x: 6.4752, y: 0.8286, w: 0.3768, h: 0.055, fill: { color: '414041' }, line: NOLINE, rectRadius: 0.027 });
  slide.addShape('ellipse', { x: 6.9455, y: 0.8249, w: 0.0882, h: 0.0882, fill: { color: '414041' }, line: NOLINE });

  ['coins', 'house', 'percent', 'target'].forEach((kind, i) => {
    icon(slide, kind, 1.193, 1.701 + i * 1.2283, 0.4067, WHITE, DARK);
    body(slide, L.icon, { x: 1.9208, y: 1.6006 + i * 1.2304, w: 2.7792, h: 0.6075, color: WHITE, align: 'justify' });
  });

  head(slide, 'Special Smart Phone Home', { x: 8.7671, y: 0.9947, w: 4.1617, h: 3.7361 });
  body(slide, L.egetLong, { x: 8.7671, y: 5.3727, w: 3.7641, h: 1.1326, align: 'justify' });
}

/* ------------------------------------------------------ 10: price cards */

function slide10(slide) {
  head(slide, 'Buying Process Simplified', { x: 1.894, y: 0.716, w: 9.5453, h: 1.0098, align: 'center' });

  const plans = [
    { x: 1.4615, tx: 1.7899, yearX: 3.4036, lineX: 1.4615, lineW: 3.375, nameW: 1.0331, fill: GREEN, name: 'Basic ', price: '$120/' },
    { x: 4.9792, tx: 5.3267, yearX: 6.9405, lineX: 4.9792, lineW: 3.375, nameW: 1.642, fill: DARK, name: 'Premium', price: '$220/' },
    { x: 8.4969, tx: 8.8348, yearX: 10.4486, lineX: 8.5064, lineW: 3.3654, nameW: 1.0331, fill: GREEN, name: 'Extra', price: '$300/' },
  ];

  plans.forEach(p => card(slide, p.x, 2.3228, 3.375, 4.3592, p.fill, 0.04916));
  plans.forEach(p => {
    head(slide, p.name, { x: p.tx, y: 2.5772, w: p.nameW, h: 0.4376, fontSize: 20, color: WHITE });
    rule(slide, p.lineX, 3.0987, p.lineW, 0, WHITE);
    head(slide, p.price, { x: p.tx, y: 3.3379, w: 1.8768, h: 0.7742, fontSize: 40, color: WHITE });
    head(slide, 'Year', { x: p.yearX, y: 3.6996, w: 0.6687, h: 0.3366, fontSize: 14, color: WHITE });
    body(slide, L.eget, { x: p.tx, y: 4.1299, w: 2.7888, h: 0.6075, color: WHITE });
    L.tick.forEach((t, i) => {
      body(slide, t, { x: p.tx, y: [4.9111, 5.2997, 5.7335][i], w: 2.7888, h: 0.345, color: WHITE, bullet: TICK });
    });
  });
}

/* ---------------------------------------------------- 11: market growth */

function slide11(slide) {
  const tiles = [
    { x: 8.2056, y: 1.1302, fill: DARK, vx: 8.7193, vy: 1.2321, lx: 8.7486, ly: 2.1336, lh: 0.3857, lFont: BODY, dx: 8.4442, dy: 2.4647 },
    { x: 6.6667, y: 3.6095, fill: GREEN, vx: 7.1804, vy: 3.7114, lx: 7.2097, ly: 4.7146, lh: 0.3777, lFont: HEAD, dx: 6.9053, dy: 5.0405 },
    { x: 9.7445, y: 3.6095, fill: TAN, vx: 10.2582, vy: 3.7114, lx: 10.2875, ly: 4.7146, lh: 0.3777, lFont: HEAD, dx: 9.9831, dy: 5.0405 },
  ];
  tiles.forEach(t => card(slide, t.x, t.y, 2.7602, 2.7602, t.fill, 0.05946));

  head(slide, 'Market Growth and Real Estate Trends', { x: 0.5594, y: 0.6271, w: 6.7745, h: 2.8273 });
  body(slide, L.penatibusComma, { x: 0.5594, y: 5.4997, w: 5.2235, h: 0.8701, align: 'justify' });
  head(slide, '12K', { x: 0.5594, y: 4.5779, w: 1.7328, h: 1.0098 });
  body(slide, TITLE_HERE, { x: 2.0143, y: 5.0828, w: 1.4583, h: 0.345, italic: true });

  tiles.forEach(t => {
    head(slide, '$2,392', { x: t.vx, y: t.vy, w: 1.7328, h: 0.7068, fontSize: 36, color: WHITE, align: 'center' });
    body(slide, TITLE_HERE, { x: t.lx, y: t.ly, w: 1.6741, h: t.lh, fontFace: t.lFont, fontSize: 14, italic: true, color: WHITE, align: 'center' });
    rule(slide, t.x, t.y + 0.895, 2.7602, 0, WHITE);
    body(slide, L.elitEnd, { x: t.dx, y: t.dy, w: 2.2829, h: 0.8701, color: WHITE, align: 'center' });
  });
}

/* ------------------------------------------------ 12: location highlights */

function slide12(slide) {
  const cols = [
    { x: 1.2941, fill: DARK, pct: '68', label: 'Mission' },
    { x: 4.1098, fill: GREEN, pct: '72', label: 'Strategy' },
    { x: 6.9256, fill: DARK, pct: '83', label: 'Strategy Map' },
    { x: 9.7413, fill: GREEN, pct: '94', label: 'Balance Score' },
  ];
  cols.forEach(c => {
    const tx = c.x + 0.1831;
    block(slide, c.x, 3.5385, 2.5518, 3.9615, c.fill);
    slide.addText(
      [{ text: c.pct, options: { fontSize: 60 } }, { text: '%', options: { fontSize: 40 } }],
      { x: tx, y: 3.9202, w: 2.1855, h: 1.1107, fontFace: HEAD, color: WHITE, align: 'center', valign: 'top' }
    );
    head(slide, c.label, { x: tx, y: 5.0102, w: 2.1855, h: 0.3702, fontSize: 16, color: WHITE, align: 'center' });
    body(slide, L.card, { x: tx, y: 5.5486, w: 2.1855, h: 1.0569, color: WHITE, align: 'center', lineSpacingMultiple: 1.2 });
  });

  head(slide, 'Location Highlights and Connectivity', { x: 1.0129, y: 0.8414, w: 7.1965, h: 1.9186 });
  body(slide, L.egetShort, { x: 8.5563, y: 1.3656, w: 3.7641, h: 0.8701, align: 'justify' });
}

/* ------------------------------------------------------------- 13: ROI */

function slide13(slide) {
  head(slide, 'ROI and Future Value Projections', { x: 0.72, y: 4.7688, w: 7.0257, h: 1.9186 });

  const steps = [
    { x: 8.7177, h: 5.2296, fill: GREEN, label: 'Q3', ly: 3.8133, tx: 9.0411, ty: 2.3598 },
    { x: 5.119, h: 4.1454, fill: DARK, label: 'Q2', ly: 2.7948, tx: 5.4204, ty: 1.2861 },
    { x: 1.5204, h: 2.8557, fill: GREEN, label: 'Q1', ly: 1.4126, tx: 1.8119, ty: 0.5425 },
  ];
  steps.forEach(s => {
    slide.addShape('snip1Rect', { x: s.x, y: -0.0153, w: 3.5986, h: s.h, fill: { color: s.fill }, line: NOLINE, flipH: true, flipV: true });
  });
  steps.forEach(s => head(slide, s.label, { x: s.tx, y: s.ly, w: 1.3412, h: 1.0098, color: WHITE }));
  steps.forEach(s => body(slide, L.quarter, { x: s.tx, y: s.ty, w: 2.882, h: 0.8701, color: WHITE, align: 'justify' }));

  body(slide, L.egetShort, { x: 7.6073, y: 5.728, w: 3.7641, h: 0.8701, align: 'justify' });
}

/* ------------------------------------------------------ 14: price plans */

function slide14(slide) {
  head(slide, 'Price Plans and Ownership Options', { x: 6.7842, y: 1.7451, w: 5.5952, h: 2.8273 });
  body(slide, L.penatibusSit, { x: 6.7842, y: 4.8848, w: 5.6565, h: 0.8701, align: 'justify' });

  const quads = [
    { x: 1.2838, y: 1.4557, fill: DARK, pct: '91%', tx: 1.5698, ty: 2.1217, tw: 1.5888 },
    { x: 3.6947, y: 1.4557, fill: TEAL, pct: '81%', tx: 3.9297, ty: 2.1217, tw: 1.6909 },
    { x: 1.2838, y: 3.8592, fill: TAN, pct: '61%', tx: 1.367, ty: 4.4609, tw: 1.9944 },
    { x: 3.6947, y: 3.8728, fill: GREEN, pct: '71%', tx: 3.7779, ty: 4.4744, tw: 1.9944 },
  ];
  quads.forEach(q => card(slide, q.x, q.y, 2.1609, 2.1609, q.fill, 0.10281));
  quads.forEach(q => {
    slide.addText(q.pct, { x: q.tx, y: q.ty, w: q.tw, h: 0.7068, fontFace: 'Montserrat SemiBold', fontSize: 36, color: WHITE, align: 'center', valign: 'top' });
  });

  rule(slide, 3.5697, 1.2564, 0, 5.0208, DARK);
  rule(slide, 1.0679, 3.7447, 4.9842, 0, DARK);

  [
    { x: 2.8308, y: 0.8099, rotate: 0 },
    { x: 5.5907, y: 3.5943, rotate: 90 },
    { x: 0.0469, y: 3.5943, rotate: 270 },
    { x: 2.8308, y: 6.3451, rotate: 0 },
  ].forEach(a => body(slide, TITLE_HERE, { x: a.x, y: a.y, w: 1.4583, h: 0.345, italic: true, rotate: a.rotate }));
}

/* ------------------------------------------------------- 15: testimonial */

function slide15(slide) {
  slide.addText('\u201C', { x: 9.4877, y: -0.114, w: 3.1535, h: 12.1172, fontFace: HEAD, fontSize: 714, color: BLACK, transparency: 83, valign: 'top' });
  block(slide, 5.5868, 5.1808, 7.7465, 1.3817, GREEN);
  head(slide, 'Client Experiences', { x: 5.9403, y: 1.1369, w: 7.0257, h: 1.0098 });
  block(slide, 0, 0, 1.6859, 7.5, DARK);
  head(slide, '20+', { x: 5.9403, y: 5.3509, w: 1.7328, h: 0.8415, fontSize: 44, color: WHITE });
  body(slide, TITLE_HERE, { x: 5.9403, y: 6.0474, w: 1.4583, h: 0.345, italic: true, color: WHITE });
  body(slide, L.dolorEnd, { x: 7.521, y: 5.5679, w: 4.8019, h: 0.6075, color: WHITE, align: 'justify' });
  body(slide, L.penatibusFull, { x: 5.9403, y: 2.9166, w: 6.7432, h: 0.8701, align: 'justify' });
  body(slide, L.natoque, { x: 5.9403, y: 3.9375, w: 6.7432, h: 0.6075, align: 'justify' });
  head(slide, 'Madison Marques', { x: 5.9403, y: 2.4958, w: 2.368, h: 0.4039, fontSize: 18 });
  photo(slide, 1.0104, 0, 4.5764, 7.5);
}

/* --------------------------------------------------------- 16: thank you */

function slide16(slide) {
  photo(slide, 6.6667, 0, 6.6667, 7.5);
  head(slide, 'Thank You', { x: 1.5019, y: 0.7787, w: 4.5493, h: 3.063, fontSize: 88 });
  body(slide, L.penatibusEnd, { x: 0.6217, y: 4.6136, w: 4.6951, h: 0.8701, align: 'justify' });
  body(slide, 'Let\u2019s Build a Better Living Together', { x: 4.0511, y: 2.8274, w: 2.0001, h: 0.6075 });
  block(slide, 6.6667, 5.5387, 6.6667, 1.0142, GREEN);

  [
    { x: 7.0283, w: 1.653, text: '+123 456 7890' },
    { x: 8.9347, w: 2.1178, text: 'www. Storytelling.com' },
    { x: 11.1868, w: 1.7849, text: '12 Your Street Name' },
  ].forEach(c => body(slide, c.text, { x: c.x, y: 5.8696, w: c.w, h: 0.3524, fontSize: 11, color: WHITE, lineSpacingMultiple: 1.5 }));

  block(slide, 0, 0, 1.2435, 4.0634, DARK);
  body(slide, L.dolorEnd, { x: 0.6217, y: 5.5432, w: 4.6951, h: 0.6075, align: 'justify' });
}

/* ------------------------------------------------------------------ build */

const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
  slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK', width: 13.3333, height: 7.5 });
  pptx.layout = 'DECK';
  pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
  pptx.title = 'Living Real Estate';

  BUILDERS.forEach(builder => {
    const slide = pptx.addSlide();
    slide.background = { color: WHITE };
    footer(slide); // master shape: sits behind everything the slide draws
    builder(slide);
  });

  return pptx.writeFile({ fileName: path.join(__dirname, '09698e65-4467-4e99-8725-6b0a2bfa9387_grok_final.pptx') });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
