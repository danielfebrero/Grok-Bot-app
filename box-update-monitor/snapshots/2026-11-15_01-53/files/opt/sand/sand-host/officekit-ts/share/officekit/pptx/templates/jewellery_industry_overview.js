/*
 * "Jewellery" presentation template - rebuilt with pptxgenjs.
 *
 * 20 slides, 13.333 x 7.5 in (16:9).  Theme: Libre Baskerville headings /
 * Albert Sans body on a cream / teal / mustard palette.
 *
 * The source deck contains empty PowerPoint picture placeholders (there is no
 * ppt/media part at all), so there is nothing raster to re-embed here.
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const TEAL = '008081'; // accent1
const GOLD = 'FFCE48'; // accent2
const CREAM = 'FFF5DA'; // accent3
const WHITE = 'FFFFFF'; // bg1
const INK = '0D0D0D'; // tx1 lumMod 95%
const INK2 = '262626'; // tx1 lumMod 85%
const BLACK = '000000'; // tx1
const GREY = 'D9D9D9'; // bg1 lumMod 85%

const HEAD = 'Libre Baskerville'; // +mj-lt
const BODY = 'Albert Sans'; // +mn-lt

/* Body copy used throughout the template. */
const L = {
  full: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur.',
  parturient: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient.',
  penatibusEt: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et.',
  cumSociis: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis.',
  massa: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa.',
  massaSp: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. ',
  sociisNatoque: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa sociis natoque.',
  ligula: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula.',
  dolorEnean: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor enean.',
  elitSp: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ',
  consectetuer: 'Lorem ipsum dolor sit amet, consectetuer.',
  commoMagnis: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus magnis dis.',
  commoEnean: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commo ligula eget dolor enean.',
  commoLigula: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commo ligula.',
  eneanCommodo: 'Lorem ipsum dolor consectetuer adipiscing enean commodo.',
  eneanShort: 'Lorem consectetuer adipiscing enean.',
  ipsumCommodo: 'Lorem ipsum consectetuer adipiscing commodo.',
  sitConsectetuer: 'Lorem ipsum dolor sit consectetuer.',
};

/* --------------------------------------------------------------- helpers */

/* Rounded-rectangle corner radius: PowerPoint stores it as a fraction of the
 * shorter side, pptxgenjs wants it in inches. */
const adj = (w, h, pct) => pct * Math.min(w, h);

/* Every text box in the source deck is an auto-fitting ("resize shape to fit
 * text") box anchored at the top. */
const TEXTBOX = { valign: 'top', fit: 'resize' };

/** Body paragraph: 14 pt Albert Sans, 130% leading, top aligned. */
function para(slide, text, o) {
  slide.addText(text, {
    fontFace: BODY, fontSize: 14, color: BLACK, lineSpacingMultiple: 1.3, ...TEXTBOX, ...o,
  });
}

/** Anything set in the display face (headings, numbers, labels, buttons). */
function head(slide, text, o) {
  slide.addText(text, { fontFace: HEAD, fontSize: 18, color: BLACK, ...TEXTBOX, ...o });
}

/** Two-tone heading, e.g. "Introduction " + italic "to Jewellery". */
function title(slide, roman, italic, o) {
  head(slide, [{ text: roman }, { text: italic, options: { italic: true } }], o);
}

/** Filled rounded card. */
function card(slide, x, y, w, h, fill, pct) {
  slide.addShape('roundRect', { x, y, w, h, fill: { color: fill }, rectRadius: adj(w, h, pct) });
}

/** Fully-rounded "Learn More" style button (roundRect, adj = 50%). */
function button(slide, x, y, w, h, label, fill, color) {
  slide.addShape('roundRect', { x, y, w, h, fill: { color: fill }, rectRadius: h / 2 });
  head(slide, label, {
    x: x + 0.1, y: y + 0.04, w: w - 0.2, h: h - 0.08,
    fontSize: 14, color, align: 'center',
  });
}

/** Bottom-rounded tab.  Drawn as a roundRect grown upwards by the corner
 *  radius so the (unwanted) top corners land above the slide edge. */
function bottomTab(slide, x, y, w, h, r, fill) {
  slide.addShape('roundRect', { x, y: y - r, w, h: h + r, fill: { color: fill }, rectRadius: r });
}

/** The "Jewellery / 2025" tabs that hang off the top edge of most slides. */
function tagPills(slide, fill, color) {
  const pills = [
    { x: 0.76, w: 1.517, tx: 0.924, tw: 1.189, label: 'Jewellery' },
    { x: 2.508, w: 0.977, tx: 2.639, tw: 0.716, label: '2025' },
  ];
  pills.forEach((p) => {
    bottomTab(slide, p.x, 0, p.w, 0.448, 0.176, fill);
    head(slide, p.label, {
      x: p.tx, y: 0.056, w: p.tw, h: 0.337, fontSize: 14, color, align: 'center',
    });
  });
}

function pageNumber(slide, n, color) {
  head(slide, String(n), {
    x: 12.749, y: 7.055, w: 0.462, h: 0.286, fontSize: 11, color, align: 'center',
  });
}

/* Slides 2/5/10/15 hide the master furniture and carry their own variants. */
const CHROME = {
  5: { pill: GOLD, pillText: INK, num: WHITE },
  10: { pill: null, num: WHITE },
  15: { pill: GOLD, pillText: INK, num: WHITE },
};

function chrome(slide, n) {
  const c = CHROME[n] || { pill: TEAL, pillText: WHITE, num: INK };
  if (c.pill) tagPills(slide, c.pill, c.pillText);
  pageNumber(slide, n, c.num);
}

/** The faint "16/20" counter tucked into the bottom-right of slides 16 & 19. */
function slashTwenty(slide, n) {
  head(slide, `${n}/20`, {
    x: 11.916, y: 6.909, w: 0.864, h: 0.337, fontSize: 14, color: CREAM, align: 'right',
  });
}

/** Elbow connector, reproduced with the deck's own xfrm / rotation / flip. */
function elbow(slide, x, y, w, h, rotate, flipH, flipV, color) {
  slide.addShape('bentConnector3', {
    x, y, w, h, rotate, flipH, flipV, line: { color, width: 0.5 },
  });
}

/* Percentage / caption / copy stack used inside the coloured cards. */
function statBlock(slide, o) {
  head(slide, o.value, { x: o.x, y: o.y, w: o.vw, h: o.vh, fontSize: o.vsize, color: o.color });
  head(slide, o.label, { x: o.lx || o.x, y: o.ly, w: o.lw, h: 0.404, color: o.color });
  para(slide, o.text, { x: o.lx || o.x, y: o.ty, w: o.tw, h: o.th || 0.686, color: o.color });
}

/* ------------------------------------------------------------ the slides */

const SLIDES = [];

// 1 - cover
SLIDES.push((s) => {
  elbow(s, 6.402, 2.146, 7.873, 3.228, 270, true, false, INK);
  s.addShape('ellipse', { x: 7.811, y: 1.223, w: 5.054, h: 5.054, fill: { color: GOLD } });
  head(s, [{ text: 'Jewellery', options: { italic: true } }],
    { x: 0.76, y: 2.014, w: 6.177, h: 1.582, fontSize: 88 });
  head(s, 'Presentation Template', { x: 0.76, y: 3.949, w: 4.857, h: 0.438, fontSize: 20 });
  para(s, L.full, { x: 0.76, y: 4.494, w: 6.073, h: 0.992 });
});

// 2 - introduction
SLIDES.push((s) => {
  elbow(s, -0.327, 1.42, 7.924, 4.628, 270, true, false, CREAM);
  title(s, 'Introduction ', 'to Jewellery', { x: 6.51, y: 1.67, w: 5.906, h: 2.121, fontSize: 60 });
  para(s, L.full, { x: 6.51, y: 3.962, w: 5.998, h: 0.992 });
  button(s, 6.646, 5.413, 1.639, 0.417, 'Learn More', TEAL, WHITE);
});

// 3 - history
SLIDES.push((s) => {
  s.addShape('ellipse', { x: 7.473, y: 5.099, w: 4.293, h: 4.293, fill: { color: GOLD } });
  title(s, 'History of ', 'Jewellery', { x: 0.76, y: 1.575, w: 5.148, h: 2.121, fontSize: 60 });
  para(s, L.commoMagnis, { x: 0.76, y: 3.977, w: 5.148, h: 0.992 });
  para(s, L.commoEnean, { x: 0.76, y: 5.239, w: 5.148, h: 0.686 });
});

// 4 - materials
SLIDES.push((s) => {
  elbow(s, 0.125, -0.221, 13.208, 3.971, 180, true, true, INK);
  title(s, 'Materials Used ', 'in Jewellery', { x: 4.53, y: 1.328, w: 6.631, h: 2.121, fontSize: 60 });
  const items = [
    { x: 4.53, num: '01', dot: TEAL, ink: WHITE, label: 'Necklaces', lw: 1.486 },
    { x: 8.522, num: '02', dot: GOLD, ink: INK, label: 'Bracelets', lw: 2.141 },
  ];
  items.forEach((it) => {
    s.addShape('ellipse', { x: it.x + 0.011, y: 4.315, w: 0.538, h: 0.538, fill: { color: it.dot } });
    head(s, [{ text: it.num, options: { italic: true } }],
      { x: it.x, y: 4.399, w: 0.559, h: 0.37, fontSize: 16, color: it.ink, align: 'center' });
    head(s, it.label, { x: it.x + 0.795, y: 4.399, w: it.lw, h: 0.404 });
    para(s, L.commoLigula, { x: it.x + 0.795, y: 4.947, w: 2.838, h: 0.992 });
  });
});

// 5 - sustainability (teal slide)
SLIDES.push((s) => {
  title(s, 'Sustainability ', 'in Jewellery',
    { x: 5.345, y: 1.572, w: 6.29, h: 2.121, fontSize: 60, color: WHITE });
  para(s, L.full, { x: 5.345, y: 3.97, w: 6.073, h: 0.992, color: WHITE });
  button(s, 5.5, 5.511, 1.639, 0.417, 'Learn More', GOLD, INK);
});

// 6 - challenges: three teal cards
SLIDES.push((s) => {
  s.addText([
    { text: 'Challenges in the ' },
    { text: 'Jewellery Industry', options: { italic: true } },
  ], { x: 0.76, y: 1.086, w: 7.781, h: 2.121, fontFace: HEAD, fontSize: 60, color: INK, valign: 'top' });
  const cards = [
    { x: 0.76, value: '67%', label: 'Challenges One', lw: 2.278 },
    { x: 4.77, value: '75%', label: 'Challenges Two', lw: 2.278 },
    { x: 8.78, value: '80%', label: 'Challenges Three', lw: 2.614 },
  ];
  cards.forEach((c) => {
    card(s, c.x, 3.75, 3.6, 2.906, TEAL, 0.11025);
    statBlock(s, {
      x: c.x + 0.248, y: 4.037, vw: 1.984, vh: 1.111, vsize: 60, value: c.value, color: WHITE,
      label: c.label, ly: 5.213, lw: c.lw, text: L.eneanCommodo, ty: 5.683, tw: 3.104,
    });
  });
});

// 7 - popular styles
SLIDES.push((s) => {
  elbow(s, 0.189, 2.132, 7.848, 3.204, 270, true, false, INK);
  s.addText([
    { text: 'Popular ' },
    { text: 'Jewellery Styles', options: { italic: true } },
  ], { x: 4.595, y: 0.791, w: 6.5, h: 2.121, fontFace: HEAD, fontSize: 60, color: INK, valign: 'top' });
  para(s, L.penatibusEt, { x: 4.595, y: 3.206, w: 7.087, h: 0.686 });
  const cards = [
    { x: 6.187, value: '01', fill: TEAL, ink: WHITE },
    { x: 9.096, value: '02', fill: GOLD, ink: INK },
  ];
  cards.forEach((c) => {
    card(s, c.x, 4.392, 2.63, 2.508, c.fill, 0.11025);
    statBlock(s, {
      x: c.x + 0.279, y: 4.609, vw: 1.984, vh: 0.841, vsize: 44, value: c.value, color: c.ink,
      label: 'Popular Style ', ly: 5.527, lw: 2.071, text: L.eneanShort, ty: 5.996, tw: 2.071,
    });
  });
});

// 8 - iconic brands
SLIDES.push((s) => {
  elbow(s, -0.633, 2.373, 14.595, 4.099, 180, true, false, INK);
  s.addText([
    { text: 'Iconic Jewellery ' },
    { text: 'Brands 2025', options: { italic: true } },
  ], { x: 0.76, y: 1.247, w: 6.434, h: 1.919, fontFace: HEAD, fontSize: 54, color: INK, valign: 'top' });
  const rows = [
    { y: 3.521, value: '78%', vw: 1.306, brand: 'Van Cleef & Arpels', fill: TEAL, ink: WHITE },
    { y: 4.968, value: '89%', vw: 1.402, brand: 'Tiffany & Co', fill: GOLD, ink: INK },
  ];
  rows.forEach((r) => {
    card(s, 0.825, r.y, 6.196, 1.229, r.fill, 0.18652);
    head(s, r.value, { x: 1.139, y: r.y + 0.194, w: r.vw, h: 0.841, fontSize: 44, color: r.ink });
    head(s, r.brand, { x: 2.637, y: r.y + 0.204, w: 2.888, h: 0.37, fontSize: 16, color: r.ink });
    para(s, L.consectetuer, { x: 2.637, y: r.y + 0.645, w: 4.069, h: 0.38, color: r.ink });
  });
});

// 9 - marketing strategies
SLIDES.push((s) => {
  s.addText([
    { text: 'Marketing Strategies ' },
    { text: 'for Jewellery', options: { italic: true } },
  ], { x: 0.76, y: 0.993, w: 8.573, h: 1.919, fontFace: HEAD, fontSize: 54, color: INK, valign: 'top' });
  para(s, L.parturient, { x: 0.76, y: 3.45, w: 5.378, h: 0.992 });
  para(s, L.massa, { x: 0.76, y: 4.682, w: 5.378, h: 0.686 });
  button(s, 0.861, 5.809, 1.639, 0.417, 'Learn More', TEAL, WHITE);
  const rows = [
    { y: 3.209, value: '78%', vw: 1.306, label: 'Marketing Strategy 01' },
    { y: 5.076, value: '80%', vw: 1.483, label: 'Marketing Strategy 02' },
  ];
  rows.forEach((r) => {
    card(s, 6.667, r.y, 5.906, 1.569, TEAL, 0.15072);
    head(s, r.value, { x: 6.86, y: r.y + 0.231, w: r.vw, h: 0.841, fontSize: 44, color: WHITE });
    head(s, r.label, { x: 8.43, y: r.y + 0.241, w: 2.888, h: 0.37, fontSize: 16, color: WHITE });
    para(s, L.ligula, { x: 8.43, y: r.y + 0.652, w: 3.949, h: 0.686, color: WHITE });
  });
});

// 10 - consumer preferences (teal slide, no tabs)
SLIDES.push((s) => {
  title(s, 'Consumer ', 'Preferences',
    { x: 5.973, y: 0.601, w: 5.503, h: 2.121, fontSize: 60, color: WHITE });
  para(s, L.massaSp, { x: 5.973, y: 3.02, w: 5.378, h: 0.686, color: WHITE });
  const cards = [
    { x: 5.973, value: '452+', label: 'Consumer One' },
    { x: 9.303, value: '671+', label: 'Consumer Two' },
  ];
  cards.forEach((c) => {
    card(s, c.x, 4.281, 2.961, 2.515, GOLD, 0.11025);
    statBlock(s, {
      x: c.x + 0.188, y: 4.455, vw: 1.984, vh: 0.909, vsize: 48, value: c.value, color: INK2,
      label: c.label, ly: 5.448, lw: 2.278, text: L.ipsumCommodo, ty: 5.937, tw: 2.585,
    });
  });
});

// 11 - future trends
SLIDES.push((s) => {
  const bars = [
    { x: 0.76, fill: TEAL, tx: 0.912, color: WHITE },
    { x: 6.891, fill: GOLD, tx: 7.043, color: INK },
  ];
  bars.forEach((b) => {
    s.addShape('round2SameRect',
      { x: b.x, y: 3.581, w: 5.682, h: 0.888, fill: { color: b.fill }, flipV: true });
    para(s, L.massa, { x: b.tx, y: 3.682, w: 5.378, h: 0.686, color: b.color });
  });
  s.addText([
    { text: 'Future Trends ' },
    { text: 'in Jewellery', options: { italic: true } },
  ], { x: 0.76, y: 4.99, w: 6.434, h: 1.919, fontFace: HEAD, fontSize: 54, color: INK, valign: 'top' });
  head(s, '2025', { x: 7.043, y: 5.119, w: 3.282, h: 1.01, fontSize: 54, color: INK });
  para(s, L.massa, { x: 7.043, y: 6.205, w: 5.378, h: 0.686, color: INK });
});

// 12 - design process
SLIDES.push((s) => {
  s.addText([
    { text: ' Jewellery ' },
    { text: 'Design Process', options: { italic: true } },
  ], { x: 0.76, y: 0.88, w: 5.906, h: 1.919, fontFace: HEAD, fontSize: 54, color: INK, valign: 'top' });
  para(s, L.massa, { x: 0.76, y: 3.003, w: 5.378, h: 0.686, color: INK });
  const cards = [
    { x: 0.76, y: 4.113, value: '56%', label: 'Progress 01' },
    { x: 3.665, y: 4.103, value: '75%', label: 'Progress 02' },
  ];
  cards.forEach((c) => {
    card(s, c.x, c.y, 2.63, 2.508, TEAL, 0.11025);
    statBlock(s, {
      x: c.x + 0.28, y: c.y + 0.217, vw: 1.984, vh: 0.841, vsize: 44, value: c.value, color: WHITE,
      label: c.label, ly: c.y + 1.135, lw: 2.071, text: L.eneanShort, ty: c.y + 1.604, tw: 2.071,
    });
  });
});

// 13 - influencer collaborations
SLIDES.push((s) => {
  title(s, 'Influencer ', 'Collaborations', { x: 0.76, y: 0.962, w: 6.108, h: 2.121, fontSize: 60, color: INK });
  para(s, L.full, { x: 0.76, y: 3.254, w: 6.073, h: 0.992 });
  para(s, L.cumSociis, { x: 0.76, y: 4.417, w: 6.073, h: 0.686 });
  button(s, 0.869, 5.581, 1.639, 0.417, 'Learn More', TEAL, WHITE);
  const cols = [
    { x: 7.071, tx: 7.294, fill: TEAL, ink: WHITE, value: '01' },
    { x: 9.941, tx: 10.164, fill: GOLD, ink: INK, value: '02' },
  ];
  cols.forEach((c) => {
    card(s, c.x, 0.975, 2.632, 5.808, c.fill, 0.11025);
    head(s, c.value, { x: c.tx, y: 1.263, w: 1.803, h: 1.447, fontSize: 80, color: c.ink });
    head(s, 'Production Collaboration ', { x: c.tx, y: 4.621, w: 2.045, h: 0.707, color: c.ink });
    para(s, L.elitSp, { x: c.tx, y: 5.503, w: 2.185, h: 0.992, color: c.ink });
  });
});

// 14 - break slide
SLIDES.push((s) => {
  card(s, 0.76, 0.952, 11.812, 5.595, TEAL, 0.06738);
  head(s, [{ text: 'Break Slide', options: { italic: true } }], {
    x: 2.068, y: 2.659, w: 9.198, h: 1.717, fontSize: 96, color: GOLD, align: 'center',
  });
  head(s, 'It Is Time to Break Until 10 Minute', {
    x: 4.087, y: 4.437, w: 5.159, h: 0.404, color: WHITE, align: 'center',
  });
});

// 15 - market analysis (teal slide with mustard columns)
SLIDES.push((s) => {
  const cols = [
    { x: 0.76, y: 0.937, h: 6.563, value: '85%' },
    { x: 3.343, y: 1.778, h: 5.722, value: '72%' },
    { x: 5.927, y: 2.667, h: 4.833, value: '65%' },
  ];
  cols.forEach((c) => {
    s.addShape('round2SameRect', { x: c.x, y: c.y, w: 2.186, h: c.h, fill: { color: GOLD } });
    head(s, c.value, { x: c.x + 0.151, y: 5.203, w: 1.748, h: 1.01, fontSize: 54, color: INK });
    para(s, L.sitConsectetuer, { x: c.x + 0.151, y: 6.213, w: 1.885, h: 0.686 });
  });
  s.addText([
    { text: 'Jewellery ' },
    { text: 'Market Analysis', options: { italic: true } },
  ], {
    x: 6.778, y: 0.601, w: 5.765, h: 1.717,
    fontFace: HEAD, fontSize: 48, color: WHITE, align: 'right', valign: 'top',
  });
  head(s, '2025', { x: 9.398, y: 4.898, w: 3.175, h: 0.841, fontSize: 44, color: WHITE, align: 'right' });
  para(s, L.cumSociis, { x: 8.635, y: 5.907, w: 3.938, h: 0.992, color: WHITE, align: 'right' });
});

// 16 - investment value: two donut gauges
SLIDES.push((s) => {
  const donuts = [
    { y: 1.031, ring: TEAL, sweepEnd: 200.82, value: '85', ly: 1.458, ty: 1.947, label: 'Invesment One' },
    { y: 4.134, ring: GOLD, sweepEnd: 147.42, value: '65', ly: 4.561, ty: 5.05, label: 'Invesment Two' },
  ];
  donuts.forEach((d) => {
    s.addShape('pie', { x: 7.087, y: d.y, w: 2.343, h: 2.335, fill: { color: GREY }, angleRange: [10.38, 356.88] });
    s.addShape('pie', { x: 7.087, y: d.y, w: 2.343, h: 2.335, fill: { color: d.ring }, angleRange: [270.63, d.sweepEnd] });
    s.addShape('ellipse', { x: 7.508, y: d.y + 0.42, w: 1.502, h: 1.496, fill: { color: CREAM } });
    s.addText([
      { text: d.value, options: { fontSize: 40 } },
      { text: '%', options: { fontSize: 28 } },
    ], { x: 7.631, y: d.y + 0.781, w: 1.256, h: 0.774, fontFace: HEAD, color: BLACK, valign: 'top' });
    head(s, d.label, { x: 9.752, y: d.ly, w: 2.278, h: 0.404, color: INK2 });
    para(s, L.ligula, { x: 9.752, y: d.ty, w: 2.821, h: 0.992, color: INK2 });
  });
  s.addText([
    { text: 'Investment ' },
    { text: 'Value of Jewellery', options: { italic: true } },
  ], { x: 0.76, y: 1.458, w: 5.486, h: 1.582, fontFace: HEAD, fontSize: 44, color: INK, valign: 'top' });
  para(s, L.parturient, { x: 0.76, y: 3.366, w: 5.556, h: 0.992 });
  para(s, L.massaSp, { x: 0.76, y: 4.576, w: 5.556, h: 0.686 });
  button(s, 0.869, 5.625, 1.639, 0.417, 'Learn More', TEAL, WHITE);
  slashTwenty(s, 16);
});

// 17 - pricing factors: combo bar + line chart
SLIDES.push((s, pptx) => {
  s.addText([
    { text: 'Pricing Factors ' },
    { text: 'in Jewellery', options: { italic: true } },
  ], { x: 0.76, y: 0.95, w: 6.255, h: 1.919, fontFace: HEAD, fontSize: 54, color: INK, valign: 'top' });

  // "DATA ONE" is a teal->dark-teal gradient in the source; approximated with
  // its mid tone since pptxgenjs only emits solid fills.
  const legend = [
    { x: 0.76, w: 1.274, label: 'DATA ONE', fill: '016060', ink: WHITE },
    { x: 5.239, w: 1.431, label: 'DATA TWO', fill: GOLD, ink: INK2 },
  ];
  legend.forEach((g) => {
    s.addShape('rect', { x: g.x, y: 3.332, w: g.w, h: 0.402, fill: { color: g.fill } });
    head(s, g.label, {
      x: g.x, y: 3.332, w: g.w, h: 0.402,
      fontSize: 12, color: g.ink, align: 'center', valign: 'middle', fit: 'none',
    });
  });

  const years = ['2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024'];
  const bars = [
    { name: 'Series 1', labels: years, values: [15, 16, 30, 27, 44, 25, 16, 50, 80] },
    { name: 'Series 2', labels: years, values: [17, 14, 28, 25, 46, 27, 14, 48, 49] },
  ];
  const trend = [{ name: 'Series 3', labels: years, values: [34, 28, 56, 50, 72, 54, 28, 74, 84] }];
  s.addChart(
    [
      { type: pptx.ChartType.bar, data: bars, options: { chartColors: [TEAL, GOLD], barGapWidthPct: 219, barOverlapPct: -27 } },
      { type: pptx.ChartType.line, data: trend, options: { chartColors: [WHITE], lineDash: 'dash', lineSize: 1, lineDataSymbolSize: 3 } },
    ],
    {
      x: 0.76, y: 4.021, w: 5.83, h: 2.816,
      showLegend: false, catAxisHidden: true, valAxisHidden: true, valAxisMaxVal: 100,
      valGridLine: { color: 'EFE7D2', size: 0.5 }, catGridLine: { style: 'none' },
      chartArea: { fill: { type: 'none' } }, plotArea: { fill: { type: 'none' } },
    },
  );

  [
    { y: 2.604, value: '$48.000' },
    { y: 4.896, value: '$68.000' },
  ].forEach((b) => {
    head(s, b.value, { x: 7.667, y: b.y, w: 4.206, h: 1.111, fontSize: 60 });
    para(s, L.dolorEnean, { x: 7.667, y: b.y + 1.256, w: 4.906, h: 0.686 });
  });
});

// 18 - pull quote
SLIDES.push((s) => {
  head(s, '\u201CEvery Piece of Jewellery is a Whisper of Art and a Shout of Style\u201D', {
    x: 0.81, y: 1.267, w: 10.983, h: 3.13, fontSize: 60,
  });
  para(s, L.parturient, { x: 7.017, y: 5.907, w: 5.556, h: 0.992, align: 'right' });
});

// 19 - contact
SLIDES.push((s) => {
  slashTwenty(s, 19);
  title(s, 'Get in Touch ', 'With Us', { x: 6.421, y: 1.519, w: 5.495, h: 1.919, fontSize: 54, color: INK });
  para(s, L.sociisNatoque, { x: 6.421, y: 3.669, w: 6.012, h: 0.686 });
  const contacts = [
    { x: 6.421, fill: TEAL, ink: WHITE, label: 'Your Phone', value: '+123 839 3920 123', vw: 2.781 },
    { x: 9.469, fill: GOLD, ink: INK, label: 'Your Email', value: 'jewellery@email.com', vw: 3.104 },
  ];
  contacts.forEach((c) => {
    s.addShape('roundRect', { x: c.x, y: 4.853, w: 2.078, h: 0.496, fill: { color: c.fill }, rectRadius: 0.248 });
    head(s, c.label, { x: c.x + 0.143, y: 4.901, w: 1.792, h: 0.4, color: c.ink, align: 'center' });
    head(s, c.value, { x: c.x, y: 5.577, w: c.vw, h: 0.404 });
  });
});

// 20 - thank you
SLIDES.push((s) => {
  elbow(s, 6.402, 2.146, 7.873, 3.228, 270, true, false, INK);
  s.addShape('ellipse', { x: 7.811, y: 1.223, w: 5.054, h: 5.054, fill: { color: GOLD } });
  head(s, [{ text: 'Thank You!', options: { italic: true } }],
    { x: 0.76, y: 2.148, w: 6.398, h: 1.447, fontSize: 80 });
  head(s, 'Presentation Template', { x: 0.76, y: 3.905, w: 4.857, h: 0.438, fontSize: 20 });
  para(s, L.full, { x: 0.76, y: 4.449, w: 6.073, h: 0.992 });
});

/* Slide background, inherited from each slide's layout in the source deck. */
const BACKGROUNDS = [
  CREAM, GOLD, CREAM, CREAM, TEAL, GOLD, CREAM, CREAM, GOLD, TEAL,
  CREAM, GOLD, CREAM, CREAM, TEAL, CREAM, CREAM, GOLD, CREAM, CREAM,
];

/* -------------------------------------------------------------- assemble */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE_16x9', width: 13.3333333, height: 7.5 }); // 12192000 x 6858000 EMU
  pptx.layout = 'WIDE_16x9';
  pptx.title = 'Jewellery Presentation Template';

  SLIDES.forEach((draw, i) => {
    const slide = pptx.addSlide();
    slide.background = { color: BACKGROUNDS[i] };
    draw(slide, pptx);
    chrome(slide, i + 1);
  });

  return pptx.writeFile({
    fileName: path.join(__dirname, '16f32fa6-47e1-4df5-9a0b-eb85690740eb_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => {
  console.error(e);
  process.exit(1);
});
