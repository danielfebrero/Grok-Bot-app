#!/usr/bin/env node
/**
 * "Pictorial Infographic - Multipurpose presentation template"
 * Rebuilt as a standalone pptxgenjs script (21 slides, 13.333in x 7.5in).
 *
 * Raster/vector illustrations from the source deck are replaced by labelled
 * placeholder rectangles (see `illustration()`); everything else - shapes,
 * text, tables and charts - is recreated with native pptxgenjs calls.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Theme
 * ------------------------------------------------------------------ */

// Theme palette "SWOT Analysis" from the source deck.
const C = {
  teal: '69AFC0',      // accent1
  green: '7EB977',     // accent2
  grass: '4FA04A',     // accent3
  lime: 'A9CF38',      // accent4
  deepTeal: '2C8991',  // accent5
  cardLine: '3DA6B6',  // accent6 - used for every card outline
  navy: '32394E',      // headline navy
  ink: '31456A',       // infographic label navy
  black: '000000',
  near: '0D0D0D',
  dark: '262626',
  gray: '595959',
  slate: '8A9AA8',
  mute: '808080',
  white: 'FFFFFF',
  footer: '4F8390',    // accent1 @ 75% luminance (master footer / page no.)
};

const MAJOR = 'Montserrat';  // +mj-lt
const MINOR = 'Open Sans';   // +mn-lt

const FOOTER_TEXT = 'Pictorial Infographic – Multipurpose presentation template';

// Card drop shadow used by every rounded "content card" in the deck.
const CARD_SHADOW = { type: 'outer', color: C.black, opacity: 0.1, blur: 40, offset: 18, angle: 50 };
const SOFT_SHADOW = { type: 'outer', color: C.black, opacity: 0.14, blur: 30, offset: 14, angle: 45 };

/* ------------------------------------------------------------------ *
 * Helpers
 * ------------------------------------------------------------------ */

// Every text box in the source deck is "resize shape to fit text"; without it
// LibreOffice/PowerPoint vertically centres overflowing copy instead of
// growing downwards, which shifts every headline off its baseline.
const AUTOFIT = 'resize';

/** Master furniture: left footer caption + right page number. */
function chrome(slide, pageNo) {
  slide.addText(FOOTER_TEXT, {
    x: 1.0, y: 6.819, w: 4.039, h: 0.269, fit: AUTOFIT,
    fontSize: 10, fontFace: MINOR, color: C.footer,
  });
  slide.addText(String(pageNo), {
    x: 11.566, y: 6.77, w: 1.145, h: 0.337, fit: AUTOFIT,
    fontSize: 14, fontFace: MAJOR, color: C.footer, align: 'center',
  });
}

/**
 * Placeholder standing in for one of the deck's hand-drawn illustrations.
 * Keeps the original bounding box so the page composition still reads right.
 */
function illustration(slide, x, y, w, h, caption, captionAt) {
  slide.addShape('rect', {
    x, y, w, h,
    fill: { color: 'F4F5F6' },
    line: { color: 'C9CFD4', width: 1, dashType: 'dash' },
  });
  slide.addText(caption || '[image]', {
    x, y: y + h * (captionAt === undefined ? 0.5 : captionAt) - 0.16, w, h: 0.32,
    align: 'center', valign: 'middle',
    fontSize: 12, fontFace: MINOR, color: '9AA4AC',
  });
}

/** White rounded card with the deck's signature thin teal outline. */
function card(slide, o) {
  slide.addShape('roundRect', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    rectRadius: o.r === undefined ? 0.17 : o.r,
    fill: { color: o.fill || C.white },
    line: { color: o.line === null ? undefined : (o.line || C.cardLine), width: 1 },
    shadow: o.shadow === null ? undefined : (o.shadow || CARD_SHADOW),
  });
}

/** Outlined stand-in for one of the deck's monoline pictograms. */
function glyph(slide, x, y, size, color) {
  slide.addShape('roundRect', {
    x: x + size * 0.08, y: y + size * 0.14, w: size * 0.84, h: size * 0.72,
    rectRadius: 0.02, fill: { type: 'none' }, line: { color, width: 1 },
  });
  slide.addShape('line', {
    x: x + size * 0.24, y: y + size * 0.42, w: size * 0.52, h: 0, line: { color, width: 1 },
  });
  slide.addShape('line', {
    x: x + size * 0.24, y: y + size * 0.62, w: size * 0.34, h: 0, line: { color, width: 1 },
  });
}

/** Two strokes forming a check mark, centred on (cx, cy). */
function checkMark(slide, cx, cy, size, color) {
  const ln = { color, width: 1.5 };
  slide.addShape('line', { x: cx - size * 0.45, y: cy, w: size * 0.33, h: size * 0.32, line: ln });
  slide.addShape('line', { x: cx - size * 0.12, y: cy - size * 0.34, w: size * 0.57, h: size * 0.66, line: ln, flipV: true });
}

/**
 * Oversized typographic quote mark sized to land its ink inside (x, y, w, h).
 * Montserrat's opening double-quote measures ~0.0044in tall per point and
 * rides ~0.0025in per point above the text box centre, hence the offsets.
 */
function quoteMark(slide, x, y, w, h, color) {
  const size = Math.round(h / 0.00435);
  const boxH = h * 3;
  slide.addText('\u201C', {
    x: x - 0.115, y: y + h / 2 + 0.0025 * size - boxH / 2, w: w + 0.4, h: boxH,
    fontSize: size, fontFace: MAJOR, bold: true, color, align: 'left', valign: 'middle',
  });
}

/** Section heading: "Presentation by Jhon Doe" + short rule to its left. */
function bylineBlock(slide, x, y) {
  slide.addShape('line', { x, y: y + 0.181, w: 0.635, h: 0, line: { color: C.teal, width: 2.5 } });
  slide.addText(
    [
      { text: 'Presentation by ', options: { fontSize: 12, fontFace: MINOR, color: C.gray } },
      { text: 'Jhon Doe', options: { fontSize: 14, fontFace: MAJOR, color: '2E3252' } },
    ],
    { x: x + 0.798, y, w: 2.858, h: 0.362, lineSpacingMultiple: 1.2, fit: AUTOFIT }
  );
}

/**
 * Big display headline (Montserrat, 90% leading).
 * `track` mirrors the source's `spc="-200"` tracking, used on the navy titles.
 */
function headline(slide, text, o) {
  slide.addText(text, {
    x: o.x, y: o.y, w: o.w, h: o.h, fit: AUTOFIT,
    fontSize: o.size, fontFace: MAJOR, color: o.color || C.navy,
    charSpacing: o.track ? -2 : 0,
    lineSpacingMultiple: o.line || 0.9,
    align: o.align || 'left', bold: o.bold || false,
  });
}

/** Body copy (Open Sans). */
function body(slide, text, o) {
  slide.addText(text, {
    x: o.x, y: o.y, w: o.w, h: o.h, fit: AUTOFIT,
    fontSize: o.size || 14, fontFace: MINOR, color: o.color || C.dark,
    lineSpacingMultiple: o.line || 1.3,
    paraSpaceBefore: o.before, align: o.align || 'left', valign: o.valign || 'top',
  });
}

/** Plain text box that grows with its content (spAutoFit), like the source. */
function label(slide, text, o) {
  slide.addText(text, Object.assign({ fit: AUTOFIT, fontFace: MINOR }, o));
}

/** "62" + superscript "%" pair used on several stat slides. */
function pct(value, size, color, opts) {
  return [
    { text: value, options: Object.assign({ fontSize: size, fontFace: MAJOR, color }, opts) },
    { text: '%', options: Object.assign({ fontSize: size, fontFace: MAJOR, color, superscript: true }, opts) },
  ];
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

// 1 - Cover
function slide01(pptx) {
  const s = pptx.addSlide();
  s.addShape('ellipse', {
    x: 0.812, y: -2.104, w: 11.708, h: 11.708,
    fill: { type: 'none' }, line: { color: '345860', width: 2 },
  });
  label(s, 'PICTORIAL\nINFOGRAPHIC', {
    x: 3.164, y: 0.663, w: 7.006, h: 2.322,
    fontSize: 66, fontFace: MAJOR, color: C.black, align: 'center',
  });
  label(s, 'MULTIPURPOSE PRESENTATION TEMPLATE', {
    x: 3.946, y: 2.993, w: 5.442, h: 0.404,
    fontSize: 18, color: C.dark, align: 'center',
  });
  illustration(s, 3.472, 3.75, 6.389, 3.75, '[illustration: two doctors]');
}

// 2 - Company Objective
function slide02(pptx) {
  const s = pptx.addSlide();
  s.addShape('line', { x: 4.081, y: 0, w: 0.074, h: 4.377, line: { color: C.black, width: 1.25 } });
  s.addShape('line', { x: -0.017, y: 5.264, w: 2.513, h: 1.339, line: { color: C.black, width: 1.25 }, flipH: true });
  illustration(s, 1.155, 1.803, 5.224, 4.048, '[illustration: agent at desk]');

  headline(s, 'Company\nObjective', { x: 7.046, y: 1.045, w: 4.523, h: 1.555, size: 60, track: true });
  bylineBlock(s, 7.225, 2.996);

  [
    { y: 3.754, fill: C.teal, title: 'Company mission A' },
    { y: 5.268, fill: C.green, title: 'Company mission B' },
  ].forEach(function (row) {
    s.addShape('roundRect', {
      x: 7.153, y: row.y + 0.08, w: 0.435, h: 0.435, rectRadius: 0.129,
      fill: { color: row.fill }, line: { color: row.fill, width: 0.5 },
    });
    checkMark(s, 7.371, row.y + 0.3, 0.24, C.white);
    label(s, row.title, {
      x: 7.756, y: row.y, w: 4.377, h: 0.419,
      fontSize: 16, fontFace: MAJOR, color: C.navy, lineSpacingMultiple: 1.3,
    });
    body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ',
      { x: 7.756, y: row.y + 0.478, w: 4.341, h: 0.686, color: C.gray });
  });
  chrome(s, 2);
}

// 3 - Strategy Market
function slide03(pptx) {
  const s = pptx.addSlide();
  illustration(s, 6.053, 2.355, 6.386, 4.251, '[illustration: cheering woman]');

  [
    { x: 5.718, y: 1.709, d: 1.535, big: '500K', small: 'New Register', fill: C.teal },
    { x: 11.003, y: 1.869, d: 1.535, big: '+123K', small: 'Join Visitor', fill: C.teal },
  ].forEach(function (b) {
    s.addShape('ellipse', {
      x: b.x, y: b.y, w: b.d, h: b.d, fill: { color: b.fill },
      line: { color: b.fill, width: 0.5 }, shadow: SOFT_SHADOW,
    });
    label(s, b.big, {
      x: b.x, y: b.y + 0.407, w: b.d, h: 0.505,
      fontSize: 24, fontFace: MAJOR, color: C.white, align: 'center',
    });
    label(s, b.small, {
      x: b.x, y: b.y + 0.79, w: b.d, h: 0.286,
      fontSize: 11, color: C.white, align: 'center',
    });
  });
  s.addShape('ellipse', {
    x: 7.292, y: 1.064, w: 1.131, h: 1.131, fill: { color: C.green },
    line: { color: C.green, width: 0.5 }, shadow: SOFT_SHADOW,
  });
  s.addText('+725', {
    x: 7.292, y: 1.064, w: 1.131, h: 1.131,
    fontSize: 18, fontFace: MINOR, color: C.white, align: 'center', valign: 'middle',
  });

  headline(s, 'Strategy Market ', { x: 1.231, y: 1.188, w: 4.019, h: 1.919, size: 60, track: true });

  s.addShape('roundRect', {
    x: 1.323, y: 3.548, w: 2.164, h: 0.622, rectRadius: 0.203,
    fill: { color: C.teal }, line: { color: C.teal, width: 0.5 },
  });
  label(s, 'Get Started!', {
    x: 1.482, y: 3.669, w: 1.455, h: 0.38,
    fontSize: 14, color: C.white, lineSpacingMultiple: 1.3,
  });
  s.addShape('ellipse', { x: 2.993, y: 3.679, w: 0.359, h: 0.359, fill: { color: C.white }, line: { color: C.white, width: 0.5 } });
  glyph(s, 3.046, 3.733, 0.252, C.teal);

  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus',
    { x: 1.331, y: 4.73, w: 3.782, h: 1.3, color: C.navy });
  chrome(s, 3);
}

// 4 - Company Target
function slide04(pptx) {
  const s = pptx.addSlide();
  illustration(s, 1.154, 0.973, 4.789, 5.554, '[illustration: woman with coins]');

  headline(s, 'Company Target', { x: 6.818, y: 1.431, w: 4.523, h: 1.919, size: 60, track: true });
  bylineBlock(s, 6.997, 3.374);

  card(s, { x: 6.818, y: 4.551, w: 4.922, h: 1.048, r: 0.175 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor',
    { x: 7.086, y: 4.732, w: 4.387, h: 0.688, color: C.gray });
  quoteMark(s, 7.143, 4.375, 0.411, 0.312, C.teal);
  chrome(s, 4);
}

// 5 - Write Something Here (two cards)
function slide05(pptx) {
  const s = pptx.addSlide();
  illustration(s, 4.293, 3.118, 4.747, 4.382, '[illustration: two people talking]');
  headline(s, 'Write Something Here', { x: 1.365, y: 0.968, w: 10.602, h: 0.909, size: 48, align: 'center', track: true, line: 1.0 });

  [
    {
      cx: 0.947, cy: 2.453, cw: 3.641, ch: 3.597, r: 0.113,
      tx: 1.318, ty: 2.944, tw: 2.899, title: 'Explore New Ideas',
      bx: 1.37, by: 4.023, bw: 2.899,
      ix: 3.458, iy: 2.056, id: 0.817, ic: C.teal,
    },
    {
      cx: 8.636, cy: 2.352, cw: 3.75, ch: 3.704, r: 0.113,
      tx: 9.018, ty: 2.774, tw: 2.986, title: 'Looks For More New Things',
      bx: 9.018, by: 4.004, bw: 2.986,
      ix: 11.222, iy: 1.943, id: 0.841, ic: C.green,
    },
  ].forEach(function (c) {
    card(s, { x: c.cx, y: c.cy, w: c.cw, h: c.ch, r: c.r });
    label(s, c.title, {
      x: c.tx, y: c.ty, w: c.tw, h: 0.909,
      fontSize: 24, bold: true, color: C.dark,
    });
    body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart',
      { x: c.bx, y: c.by, w: c.bw, h: 1.607 });
    s.addShape('ellipse', {
      x: c.ix, y: c.iy, w: c.id, h: c.id,
      fill: { color: c.ic }, line: { color: c.ic, width: 0.5 }, shadow: SOFT_SHADOW,
    });
    glyph(s, c.ix + c.id * 0.27, c.iy + c.id * 0.27, c.id * 0.46, C.white);
  });
}

// 6 - Two stacked feature cards
function slide06(pptx) {
  const s = pptx.addSlide();
  illustration(s, 0, 1.791, 6.155, 3.918, '[illustration: couple at laptop]');
  s.addShape('line', { x: 5.256, y: 5.43, w: 0, h: 2.07, line: { color: C.black, width: 1.25 } });

  [
    {
      y: 1.018, h: 2.691, badge: C.teal, by: 1.251,
      title: 'A wonderful serenity has taken possession of my entire soul',
      ty: 1.186,
      copy: 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. Ithe charm of existence.',
      cy: 2.07, ch: 1.281, note: null,
    },
    {
      y: 3.809, h: 2.673, badge: C.green, by: 4.027,
      title: 'A wonderful serenity has taken possession of my entire soul',
      ty: 4.004,
      copy: 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings which I enjoy with my whole heart.',
      cy: 4.886, ch: 0.978, note: '- I am alone, and feel the charm of existence.',
    },
  ].forEach(function (b) {
    card(s, { x: 7.394, y: b.y, w: 5.249, h: b.h, r: 0.084 });
    s.addShape('roundRect', {
      x: 6.97, y: b.by, w: 0.832, h: 0.643, rectRadius: 0.138,
      fill: { color: b.badge }, line: { color: b.badge, width: 0.5 }, shadow: SOFT_SHADOW,
    });
    glyph(s, 7.31, b.by + 0.14, 0.37, C.white);
    s.addShape('line', { x: 7.085, y: b.by + 0.067, w: 0, h: 0.51, line: { color: C.white, width: 2.5 } });
    label(s, b.title, {
      x: 8.09, y: b.ty, w: 3.591, h: 0.686,
      fontSize: 14, fontFace: MAJOR, color: C.dark, lineSpacingMultiple: 1.3,
    });
    body(s, b.copy, { x: 8.09, y: b.cy, w: 4.285, h: b.ch, size: 12, line: 1.5 });
    if (b.note) body(s, b.note, { x: 8.025, y: 5.861, w: 4.291, h: 0.373, size: 12, line: 1.5 });
  });
  chrome(s, 6);
}

// 7 - Lorem ipsum headline + note card
function slide07(pptx) {
  const s = pptx.addSlide();
  illustration(s, 6.968, 1.762, 5.295, 5.24, '[illustration: woman waving]');
  headline(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ',
    { x: 1.14, y: 1.083, w: 7.554, h: 2.121, size: 40, line: 1.0 });
  card(s, { x: 1.14, y: 3.729, w: 4.642, h: 1.962, r: 0.25 });
  body(s, 'Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin pharetra nonummy pede. Mauris et orci.',
    { x: 1.739, y: 4.061, w: 3.895, h: 1.299, color: C.gray });
  s.addShape('ellipse', { x: 1.431, y: 4.141, w: 0.245, h: 0.245, fill: { type: 'none' }, line: { color: C.dark, width: 1 } });
  s.addText('i', { x: 1.431, y: 4.141, w: 0.245, h: 0.245, fontSize: 9, fontFace: MINOR, color: C.dark, align: 'center', valign: 'middle' });
  chrome(s, 7);
}

// 8 - Market size diagram (stacked bar chart)
function slide08(pptx) {
  const s = pptx.addSlide();
  illustration(s, 1.272, 2.318, 4.887, 4.262, '[illustration: woman in cart]');
  headline(s, 'Market size diagram', { x: 0.863, y: 0.826, w: 6.576, h: 0.909, size: 48, color: C.dark, line: 1.0, track: true });

  card(s, { x: 6.929, y: 1.908, w: 5.563, h: 3.688, r: 0.386 });
  label(s, 'Data Market 2023', {
    x: 7.238, y: 2.224, w: 4.502, h: 0.37,
    fontSize: 16, fontFace: MAJOR, color: C.navy,
  });

  const cats = ['Text 1', 'Text 2', 'Text 3', 'Text 4'];
  s.addChart(pptx.ChartType.bar, [
    { name: 'Series 1', labels: cats, values: [1, 1, 6, 1] },
    { name: 'Series 2', labels: cats, values: [1, 5, 3, 4] },
    { name: 'Series 3', labels: cats, values: [2, 2, 4, 1] },
  ], {
    x: 7.186, y: 2.631, w: 5.049, h: 2.731,
    barDir: 'bar', barGrouping: 'stacked', barGapWidthPct: 150, barOverlapPct: 100,
    chartColors: [C.teal, C.green, C.grass],
    valAxisMaxVal: 15, valAxisMajorUnit: 2,
    catAxisLabelColor: C.navy, catAxisLabelFontFace: MINOR, catAxisLabelFontSize: 12,
    valAxisLabelColor: C.navy, valAxisLabelFontFace: MINOR, valAxisLabelFontSize: 12,
    valGridLine: { color: 'DCE0E4', style: 'solid', size: 1 },
    catGridLine: { style: 'none' },
    catAxisLineShow: true, catAxisLineColor: 'DCE0E4', valAxisLineShow: false,
    showLegend: false, showValue: true,
    dataLabelColor: C.white, dataLabelFontFace: MINOR, dataLabelFontSize: 12, dataLabelPosition: 'ctr',
  });

  [
    { x: 7.379, label: 'Datatype 01', color: C.teal },
    { x: 9.029, label: 'Datatype 02', color: C.green },
    { x: 10.777, label: 'Datatype 03', color: C.grass },
  ].forEach(function (l) {
    s.addShape('ellipse', { x: l.x, y: 6.025, w: 0.121, h: 0.121, fill: { color: l.color }, line: { color: l.color, width: 0.5 } });
    label(s, l.label, { x: l.x + 0.131, y: 5.935, w: 1.327, h: 0.303, fontSize: 12, color: C.navy });
  });
  chrome(s, 8);
}

// 9 - Table Price
function slide09(pptx) {
  const s = pptx.addSlide();
  illustration(s, 0.953, 1.588, 4.179, 5.912, '[illustration: woman with book]', 0.35);
  headline(s, 'Table Price', { x: 5.759, y: 1.485, w: 6.576, h: 0.774, size: 40, color: C.dark, align: 'center', line: 1.0, track: true });

  [
    {
      cx: 5.902, cy: 2.791, cw: 2.833, ch: 3.466, cf: C.white,
      name: 'Basic', nameColor: C.dark, price: '$7', priceColor: C.teal,
      btnFill: C.green, btnText: C.white,
      icx: 6.868, icy: 3.014, icFill: C.green, tx: 6.109, bx: 6.3,
    },
    {
      cx: 9.059, cy: 2.682, cw: 3.062, ch: 3.745, cf: C.teal,
      name: 'Creative', nameColor: C.white, price: '$45', priceColor: C.white,
      btnFill: C.white, btnText: C.teal,
      icx: 10.127, icy: 2.903, icFill: C.deepTeal, tx: 9.368, bx: 9.559,
    },
  ].forEach(function (p) {
    card(s, { x: p.cx, y: p.cy, w: p.cw, h: p.ch, r: 0.198, fill: p.cf });
    s.addShape('ellipse', {
      x: p.icx, y: p.icy, w: 0.901, h: 0.901,
      fill: { color: p.icFill }, line: { color: p.icFill, width: 0.5 },
    });
    glyph(s, p.icx + 0.289, p.icy + 0.289, 0.324, p.cf === C.white ? C.white : 'D8E9EC');
    label(s, p.name, { x: p.tx, y: 4.188, w: 2.42, h: 0.337, fontSize: 14, color: p.nameColor, align: 'center' });
    label(s, [
      { text: p.price, options: { fontSize: 32, fontFace: MAJOR, color: p.priceColor } },
      { text: ' /month', options: { fontSize: 20, fontFace: MAJOR, color: p.priceColor } },
    ], { x: p.tx, y: 4.49, w: 2.42, h: 0.64, align: 'center' });
    s.addShape('roundRect', {
      x: p.bx, y: 5.414, w: 2.037, h: 0.518, rectRadius: 0.113,
      fill: { color: p.btnFill }, line: { color: p.btnFill, width: 0.5 }, shadow: SOFT_SHADOW,
    });
    label(s, 'Perfect 2022', {
      x: p.bx, y: 5.505, w: 2.037, h: 0.337,
      fontSize: 14, fontFace: MAJOR, bold: true, color: p.btnText, align: 'center',
    });
  });
  chrome(s, 9);
}

// 10 - Knowledge is Power
function slide10(pptx) {
  const s = pptx.addSlide();
  illustration(s, 5.433, 1.943, 7.2, 5.557, '[illustration: woman with books]');
  headline(s, 'Knowledge is Power', { x: 1.341, y: 1.37, w: 5.246, h: 1.56, size: 48, color: C.near });

  label(s, '88M+', {
    x: 1.341, y: 3.19, w: 3.457, h: 0.909,
    fontSize: 48, fontFace: MAJOR, bold: true, charSpacing: -3, color: C.teal,
  });
  label(s, pct('43', 48, C.green, { bold: true, charSpacing: -3 }), {
    x: 3.813, y: 3.19, w: 3.457, h: 0.909,
  });
  label(s, 'Participants in total.', {
    x: 1.341, y: 4.574, w: 3.037, h: 0.37,
    fontSize: 16, fontFace: MAJOR, color: C.dark, lineSpacingMultiple: 1.0,
  });
  body(s, 'Lorem ipsum dolor sit amet. Qui sint neque a velit modi quo numquam. Non exercitationem reiciendis qui.',
    { x: 1.341, y: 5.118, w: 3.426, h: 0.933, line: 1.2 });
  chrome(s, 10);
}

// 11 - Percentages + area chart
function slide11(pptx) {
  const s = pptx.addSlide();
  illustration(s, 0.634, 0.018, 5.575, 7.482, '[illustration: woman presenting]', 0.4);

  card(s, { x: 5.04, y: 3.648, w: 7.198, h: 2.814, r: 0.22 });
  const years = ['2019', '2020', '2021', '2022', '2023', '2024', '2025', '2026'];
  s.addChart(pptx.ChartType.area, [
    { name: 'Growth', labels: years, values: [58, 78, 44, 55, 70, 52, 74, 46] },
  ], {
    x: 5.331, y: 4.062, w: 6.71, h: 1.729,
    chartColors: [C.green], chartColorsOpacity: 30,
    lineSize: 1, lineDataSymbol: 'none',
    valAxisHidden: true, catAxisHidden: true,
    valAxisMaxVal: 100, valAxisMinVal: 0,
    catGridLine: { style: 'none' }, valGridLine: { style: 'none' },
    showLegend: false, showValue: false,
    layout: { x: 0, y: 0, w: 1, h: 1 },
  });
  years.forEach(function (yr, i) {
    label(s, yr, {
      x: 5.331 + i * 0.87, y: 5.868, w: 0.62, h: 0.332,
      fontSize: 9, color: C.mute, align: 'center', lineSpacingMultiple: 1.2,
    });
  });

  [
    { y: 1.389, ty: 1.444, value: '62%', color: C.teal, copy: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue' },
    { y: 2.287, ty: 2.342, value: '78%', color: C.green, copy: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue' },
  ].forEach(function (r) {
    label(s, r.value, {
      x: 6.78, y: r.y, w: 3.798, h: 0.774,
      fontSize: 40, fontFace: MAJOR, charSpacing: -3, color: r.color,
    });
    body(s, r.copy, { x: 8.44, y: r.ty, w: 3.798, h: 0.604, size: 12, color: C.slate });
  });
}

// 12 - Target rings
function slide12(pptx) {
  const s = pptx.addSlide();
  illustration(s, 6.912, 0.717, 5.901, 6.783, '[illustration: archer]');

  // Concentric target: alternating white / teal ellipses with a brown outline.
  [
    { x: 0.890, y: 1.473, w: 1.641, h: 4.558, fill: C.white },
    { x: 1.198, y: 1.858, w: 1.205, h: 3.770, fill: C.teal },
    { x: 1.321, y: 2.259, w: 0.959, h: 2.976, fill: C.white },
    { x: 1.436, y: 2.615, w: 0.728, h: 2.264, fill: C.teal },
    { x: 1.562, y: 3.006, w: 0.477, h: 1.480, fill: C.white },
    { x: 1.675, y: 3.358, w: 0.250, h: 0.784, fill: C.teal },
  ].forEach(function (r) {
    s.addShape('ellipse', { x: r.x, y: r.y, w: r.w, h: r.h, fill: { color: r.fill }, line: { color: '56302A', width: 1 } });
  });

  label(s, '88M+', {
    x: 2.98, y: 2.61, w: 3.457, h: 0.909,
    fontSize: 48, fontFace: MAJOR, bold: true, charSpacing: -3, color: C.green,
  });
  label(s, 'Participants in total.', {
    x: 2.98, y: 3.396, w: 3.037, h: 0.37,
    fontSize: 16, fontFace: MAJOR, color: C.near, lineSpacingMultiple: 1.0,
  });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar',
    { x: 2.98, y: 4.193, w: 3.426, h: 1.212, line: 1.2 });
  chrome(s, 12);
}

// 13 - Follow your passions
function slide13(pptx) {
  const s = pptx.addSlide();
  illustration(s, 0.48, 1.377, 6.361, 6.123, '[illustration: man with trophy]');
  label(s, [
    { text: 'Follow your ', options: { fontSize: 50, fontFace: MAJOR, color: C.near } },
    { text: 'passions ', options: { fontSize: 50, fontFace: MAJOR, color: C.green } },
  ], { x: 7.453, y: 1.076, w: 5.246, h: 1.616, lineSpacingMultiple: 0.9, charSpacing: -2 });

  [3.103, 5.005].forEach(function (y) {
    label(s, '01. Type Something', {
      x: 7.473, y: y, w: 4.416, h: 0.436,
      fontSize: 24, fontFace: MAJOR, color: C.near, lineSpacingMultiple: 0.8,
    });
    body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar',
      { x: 7.473, y: y + 0.598, w: 4.546, h: 0.994 });
  });
}

// 14 - Data Market (clustered column chart)
function slide14(pptx) {
  const s = pptx.addSlide();
  illustration(s, 0.371, 1.122, 5.144, 5.004, '[illustration: woman at desk]');
  // Desk edge: a hairline top plus three legs running off the bottom.
  s.addShape('rect', { x: 0, y: 6.126, w: 6.119, h: 0.05, fill: { color: C.black }, line: { type: 'none' } });
  [
    { x: 5.16, y: 6.159, w: 0.951, h: 1.341 },
    { x: 6.108, y: 6.176, w: 0.004, h: 1.324 },
    { x: 5.494, y: 6.584, w: 0.615, h: 0.93 },
  ].forEach(function (l) {
    s.addShape('line', { x: l.x, y: l.y, w: l.w, h: l.h, line: { color: C.near, width: 2 }, flipH: true });
  });

  headline(s, 'Data Market', { x: 7.05, y: 1.377, w: 5.246, h: 0.828, size: 48, color: C.near });
  body(s, 'Lorem ipsum dolor sit amet. Qui sint neque a velit modi quo numquam. Non exercitationem reiciendis qui.',
    { x: 7.05, y: 2.354, w: 5.246, h: 0.646, line: 1.2 });

  card(s, { x: 7.156, y: 3.444, w: 5.14, h: 2.858, r: 0.182 });
  const cats = ['DATA 1', 'DATA 2', 'DATA 3', 'DATA 4'];
  s.addChart(pptx.ChartType.bar, [
    { name: 'Series 1', labels: cats, values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels: cats, values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels: cats, values: [2, 2, 3, 5] },
  ], {
    x: 7.487, y: 3.652, w: 4.479, h: 2.443,
    barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 219, barOverlapPct: -27,
    chartColors: [C.teal, 'A5CFD9', '4F8390'],
    catAxisLabelColor: C.ink, catAxisLabelFontFace: MINOR, catAxisLabelFontSize: 12, catAxisLabelFontBold: true,
    valAxisLabelColor: C.slate, valAxisLabelFontFace: MINOR, valAxisLabelFontSize: 12,
    valGridLine: { color: 'E6E9EB', style: 'solid', size: 1 },
    catGridLine: { style: 'none' },
    catAxisLineShow: true, catAxisLineColor: 'E0E0E0', valAxisLineShow: false,
    showLegend: false, showValue: false,
  });
  chrome(s, 14);
}

// 15 - Write Your Stuff Here (three doughnuts)
function slide15(pptx) {
  const s = pptx.addSlide();
  illustration(s, 1.373, 1.311, 4.558, 6.189, '[illustration: thinking woman]');
  headline(s, 'Write Your\nStuff Here', { x: 6.511, y: 1.121, w: 6.028, h: 1.919, size: 60, color: C.near });

  [
    { x: 6.444, tx: 6.779, lx: 6.359, color: C.deepTeal },
    { x: 8.410, tx: 8.745, lx: 8.325, color: C.green },
    { x: 10.376, tx: 10.711, lx: 10.291, color: 'ABD0D3' },
  ].forEach(function (d) {
    s.addChart(pptx.ChartType.doughnut, [
      { name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: [7, 3] },
    ], {
      x: d.x, y: 3.415, w: 1.808, h: 1.682,
      holeSize: 75, firstSliceAng: 0,
      chartColors: [d.color, 'F2F2F2'],
      showLegend: false, showValue: false, dataBorder: { pt: 0, color: C.white },
      layout: { x: 0, y: 0, w: 1, h: 1 },
    });
    label(s, pct('62', 32, C.near), { x: d.tx, y: 3.936, w: 1.138, h: 0.64, align: 'center' });
    label(s, 'Total Value', {
      x: d.lx, y: 5.092, w: 1.979, h: 0.325,
      fontSize: 16, fontFace: MAJOR, color: C.near, align: 'center', lineSpacingMultiple: 0.8,
    });
  });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna',
    { x: 6.485, y: 6.002, w: 5.66, h: 0.646, line: 1.2 });
}

// 16 - Hexagon product cards
function slide16(pptx) {
  const s = pptx.addSlide();
  illustration(s, 3.46, 1.375, 6.414, 4.75, '[illustration: shrugging man]');
  s.addShape('line', { x: 0, y: 6.125, w: 13.333, h: 0, line: { color: C.near, width: 2.5 } });

  [
    { hx: 0.625, hy: 2.065, rot: 90, fill: C.teal, tx: 0.900, ty: 2.500, title: 'Product A', ix: 1.080, iy: 1.774, id: 0.637 },
    { hx: 3.117, hy: 0.921, rot: 90, fill: C.teal, tx: 3.392, ty: 1.462, title: 'Product B', ix: 3.519, iy: 0.677, id: 0.644 },
    { hx: 7.535, hy: 0.969, rot: 270, fill: C.green, tx: 7.810, ty: 1.404, title: 'Product B', ix: 9.164, iy: 0.733, id: 0.648 },
    { hx: 10.028, hy: 2.113, rot: 270, fill: C.green, tx: 10.302, ty: 2.548, title: 'Product A', ix: 11.700, iy: 1.906, id: 0.652 },
  ].forEach(function (h) {
    s.addShape('hexagon', {
      x: h.hx, y: h.hy, w: 2.68, h: 2.33, rotate: h.rot,
      fill: { color: h.fill }, line: { color: h.fill, width: 0.5 }, shadow: SOFT_SHADOW,
    });
    label(s, h.title, {
      x: h.tx, y: h.ty, w: 2.131, h: 0.44,
      fontSize: 18, fontFace: MAJOR, bold: true, color: C.white, align: 'center', lineSpacingMultiple: 1.2,
    });
    label(s, 'Ut wisi enim ad minim veniam, quis nostrud exerci tation', {
      x: h.tx, y: h.ty + 0.538, w: 2.131, h: 0.809,
      fontSize: 12, color: C.white, align: 'center', lineSpacingMultiple: 1.2,
    });
    s.addShape('ellipse', {
      x: h.ix, y: h.iy, w: h.id, h: h.id,
      fill: { color: C.white }, line: { color: C.white, width: 0.5 }, shadow: SOFT_SHADOW,
    });
    glyph(s, h.ix + h.id * 0.26, h.iy + h.id * 0.26, h.id * 0.48, h.fill);
  });
  chrome(s, 16);
}

// 17 - Strength / Weakness split
function slide17(pptx) {
  const s = pptx.addSlide();
  s.addShape('rect', { x: 0, y: 0, w: 6.688, h: 7.5, fill: { color: C.teal }, line: { type: 'none' } });
  illustration(s, 0.0, 2.839, 4.51, 3.5, '[illustration: thumbs up]');
  illustration(s, 8.613, 0.928, 4.72, 3.683, '[illustration: thumbs down]');

  [
    { title: 'STRENGTH', tc: C.white, x: 2.239, y: 1.148, bw: 3.4, bh: 0.65, bc: C.white, cx: 1.371, cy: 1.162, cf: C.dark,
      copy: 'Lorem ipsum dolor sit amet. Qui sint neque a velit modi quo numquam. ' },
    { title: 'WEAKNESS', tc: C.near, x: 9.019, y: 4.912, bw: 3.607, bh: 0.929, bc: C.dark, cx: 8.164, cy: 4.829, cf: C.near,
      copy: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue' },
  ].forEach(function (b) {
    label(s, b.title, { x: b.x, y: b.y, w: 4.181, h: 0.505, fontSize: 24, fontFace: MAJOR, color: b.tc });
    body(s, b.copy, { x: b.x, y: b.y + 0.559, w: b.bw, h: b.bh, line: 1.2, color: b.bc });
    s.addShape('ellipse', { x: b.cx, y: b.cy, w: 0.639, h: 0.639, fill: { color: b.cf }, line: { color: b.cf, width: 0.5 } });
    glyph(s, b.cx + 0.127, b.cy + 0.127, 0.386, C.white);
  });
  chrome(s, 17);
}

// 18 - Venn diagram
function slide18(pptx) {
  const s = pptx.addSlide();
  illustration(s, 1.337, 1.52, 5.33, 5.98, '[illustration: woman with laptop]');

  label(s, 'Venn\nDiagram', {
    x: 5.11, y: 1.168, w: 3.114, h: 0.902,
    fontSize: 20, fontFace: MAJOR, bold: true, color: C.ink,
    lineSpacingMultiple: 1.1, paraSpaceBefore: 6, valign: 'middle',
  });
  body(s, 'A wonderful serenity has taken possession of my entire soul,',
    { x: 5.11, y: 2.291, w: 2.831, h: 0.994, color: C.slate });

  [
    { x: 7.932, y: 1.158, color: C.grass, tx: 8.984, ty: 1.793 },
    { x: 6.943, y: 2.857, color: C.teal, tx: 7.222, ty: 4.618 },
    { x: 8.920, y: 2.857, color: C.lime, tx: 10.682, ty: 4.618 },
  ].forEach(function (h) {
    s.addShape('hexagon', {
      x: h.x, y: h.y, w: 3.956, h: 3.523, rotate: 90,
      fill: { color: h.color, transparency: 40 }, line: { type: 'none' },
    });
    label(s, 'Far far away behind', {
      x: h.tx, y: h.ty, w: 1.851, h: 0.72,
      fontSize: 16, color: C.white, align: 'center', lineSpacingMultiple: 1.2,
    });
  });

  [
    { n: '01', x: 8.148, y: 3.154 },
    { n: '02', x: 10.488, y: 3.154 },
    { n: '03', x: 9.341, y: 5.013 },
  ].forEach(function (l) {
    label(s, l.n, { x: l.x, y: l.y, w: 1.182, h: 0.64, fontSize: 32, fontFace: MAJOR, color: C.white, align: 'center' });
  });

  s.addShape('hexagon', {
    x: 9.036, y: 3.253, w: 1.743, h: 1.546, rotate: 90,
    fill: { color: C.white }, line: { color: C.white, width: 0.5 },
  });
  label(s, 'Venn Diagram', {
    x: 9.196, y: 3.706, w: 1.423, h: 0.64,
    fontSize: 16, fontFace: MAJOR, bold: true, color: C.ink, align: 'center',
  });
}

// 19 - Write Somethin Here (horizontal bar chart)
function slide19(pptx) {
  const s = pptx.addSlide();
  illustration(s, 7.526, 1.603, 5.233, 5.136, '[illustration: woman with calculator]', 0.62);
  s.addShape('rect', { x: -0.026, y: 5.372, w: 13.359, h: 0.05, fill: { color: C.near }, line: { type: 'none' } });

  headline(s, 'Write Somethin Here', { x: 1.085, y: 1.155, w: 7.548, h: 0.707, size: 40, color: C.near });

  card(s, { x: 1.201, y: 2.406, w: 6.305, h: 2.484, r: 0.061 });
  const cats = ['Data 1', 'Data 2'];
  s.addChart(pptx.ChartType.bar, [
    { name: 'Series 1', labels: cats, values: [4.3, 2.5] },
    { name: 'Series 2', labels: cats, values: [2.4, 4.4] },
    { name: 'Series 3', labels: cats, values: [2, 2] },
  ], {
    x: 1.508, y: 2.553, w: 5.964, h: 2.19,
    barDir: 'bar', barGrouping: 'clustered', barGapWidthPct: 182,
    chartColors: [C.green, C.grass, C.teal],
    catAxisLabelColor: C.ink, catAxisLabelFontFace: MINOR, catAxisLabelFontSize: 14,
    valAxisLabelColor: C.slate, valAxisLabelFontFace: MINOR, valAxisLabelFontSize: 12,
    valGridLine: { color: 'EFEFEF', style: 'solid', size: 1 },
    catGridLine: { style: 'none' },
    catAxisLineShow: false, valAxisLineShow: false,
    showLegend: false, showValue: false,
  });

  quoteMark(s, 1.26, 6.072, 0.333, 0.271, C.near);
  body(s, 'Lorem ipsum dolor sit amet. Qui sint neque a velit modi quo numquam. Non exercitationem reiciendis',
    { x: 1.781, y: 5.971, w: 4.987, h: 0.65, line: 1.2, before: 12 });
  chrome(s, 19);
}

// 20 - Diagram Infographic (arrow columns)
function slide20(pptx) {
  const s = pptx.addSlide();
  illustration(s, 0.697, 1.343, 4.99, 6.157, '[illustration: woman with piggy bank]');

  // homePlate arrows rotated 270deg become upward-pointing columns.
  [
    { x: 6.665, y: 4.975, w: 1.097, color: C.teal, label: 'January', lx: 6.886, lw: 0.654, vy: 4.401, vx: 6.946, vw: 0.535, v: '+21.5' },
    { x: 7.432, y: 4.674, w: 1.699, color: C.green, label: 'February', lx: 7.940, lw: 0.722, vy: 3.817, vx: 8.014, vw: 0.535, v: '+36.5' },
    { x: 8.021, y: 4.195, w: 2.657, color: C.grass, label: 'March', lx: 9.105, lw: 0.508, vy: 2.841, vx: 9.046, vw: 0.606, v: '+48.3' },
    { x: 8.576, y: 3.682, w: 3.683, color: C.lime, label: 'April', lx: 10.229, lw: 0.377, vy: 1.858, vx: 10.150, vw: 0.535, v: '+64.5' },
    { x: 9.987, y: 4.025, w: 2.997, color: C.deepTeal, label: 'June', lx: 11.287, lw: 0.396, vy: 2.461, vx: 11.218, vw: 0.535, v: '+52.1' },
  ].forEach(function (b) {
    s.addShape('homePlate', {
      x: b.x, y: b.y, w: b.w, h: 0.802, rotate: 270,
      fill: { color: b.color }, line: { color: b.color, width: 0.5 },
    });
    label(s, b.label, {
      x: b.lx - 0.25, y: 6.145, w: b.lw + 0.5, h: 0.342, margin: 0,
      fontSize: 12, bold: true, color: C.ink, align: 'center', lineSpacingMultiple: 1.3,
    });
    label(s, pct(b.v, 12, b.color, { bold: true, fontFace: MINOR }), {
      x: b.vx - 0.15, y: b.vy, w: b.vw + 0.3, h: 0.341, margin: 0, align: 'center', lineSpacingMultiple: 1.3,
    });
  });
  s.addShape('rect', { x: 6.583, y: 5.924, w: 5.552, h: 0.05, fill: { color: C.white }, line: { type: 'none' } });

  label(s, 'Diagram Infographic', {
    x: 6.457, y: 1.21, w: 3.114, h: 0.818,
    fontSize: 20, fontFace: MAJOR, bold: true, color: C.ink,
    lineSpacingMultiple: 1.1, paraSpaceBefore: 6, valign: 'middle',
  });
  body(s, 'A wonderful serenity has taken possession of my entire soul,',
    { x: 6.457, y: 2.149, w: 2.831, h: 0.994, color: C.slate });
}

// 21 - Statistic Team dashboard
function slide21(pptx) {
  const s = pptx.addSlide();
  illustration(s, 6.771, 1.66, 6.138, 5.84, '[illustration: man with globe]');

  card(s, { x: 0.813, y: 1.358, w: 5.62, h: 3.267, r: 0.124 });
  label(s, 'Statistic Team', {
    x: 1.051, y: 1.404, w: 2.529, h: 0.428,
    fontSize: 18, fontFace: MAJOR, bold: true, color: C.ink, valign: 'middle',
  });

  s.addChart(pptx.ChartType.doughnut, [
    { name: 'Teams', labels: ['Team One', 'Team Three', 'Team Two'], values: [25, 30, 45] },
  ], {
    x: 1.145, y: 2.011, w: 2.343, h: 2.337,
    holeSize: 65, firstSliceAng: 0,
    chartColors: [C.teal, C.grass, C.green],
    showLegend: false, showValue: false, dataBorder: { pt: 0, color: C.white },
    layout: { x: 0, y: 0, w: 1, h: 1 },
  });
  glyph(s, 2.083, 2.946, 0.467, C.slate);

  label(s, '75%', {
    x: 3.799, y: 1.857, w: 2.343, h: 0.752,
    fontSize: 36, fontFace: MAJOR, bold: true, color: C.teal,
  });
  s.addShape('roundRect', { x: 3.887, y: 2.715, w: 2.167, h: 0.072, rectRadius: 0.036, fill: { color: 'E9E9E9' }, line: { type: 'none' } });
  s.addShape('roundRect', { x: 3.887, y: 2.715, w: 1.514, h: 0.072, rectRadius: 0.036, fill: { color: C.teal }, line: { type: 'none' } });

  [
    { y: 3.049, name: 'Team One', nw: 1.222, value: '25%', color: C.teal, rule: 3.439 },
    { y: 3.536, name: 'Team Two', nw: 1.170, value: '45%', color: C.green, rule: 3.925 },
    { y: 4.022, name: 'Team Three', nw: 1.587, value: '30%', color: C.grass, rule: null },
  ].forEach(function (r) {
    label(s, r.name, { x: 3.813, y: r.y, w: r.nw, h: 0.317, fontSize: 12, color: C.slate });
    label(s, r.value, { x: 5.565, y: r.y, w: 0.571, h: 0.317, fontSize: 12, bold: true, color: r.color, align: 'right' });
    if (r.rule !== null) {
      s.addShape('line', { x: 3.813, y: r.rule, w: 2.323, h: 0, line: { color: 'D3D3D3', width: 1, dashType: 'dash' } });
    }
  });

  [
    { x: 0.821, tx: 1.026, big: '20K', small: 'Stats', r: 0.111 },
    { x: 3.761, tx: 3.966, big: '597', small: 'Personals', r: 0.13 },
  ].forEach(function (t) {
    card(s, { x: t.x, y: 4.881, w: 2.68, h: 1.549, r: t.r });
    label(s, t.big, {
      x: t.tx, y: 5.195, w: 2.27, h: 0.578,
      fontSize: 24, fontFace: MAJOR, bold: true, color: C.ink, align: 'center', valign: 'bottom',
    });
    label(s, t.small, {
      x: t.tx, y: 5.721, w: 2.27, h: 0.422,
      fontSize: 16, color: C.slate, align: 'center',
    });
  });
  chrome(s, 21);
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK', width: 13.333, height: 7.5 });
  pptx.layout = 'DECK';
  pptx.title = 'Pictorial Infographic';
  pptx.theme = { headFontFace: MAJOR, bodyFontFace: MINOR };

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07,
    slide08, slide09, slide10, slide11, slide12, slide13, slide14,
    slide15, slide16, slide17, slide18, slide19, slide20, slide21,
  ].forEach(function (fn) { fn(pptx); });

  return pptx.writeFile({
    fileName: path.join(__dirname, '179066c9-93d5-4e84-ab95-a58a42ddce12_grok_final.pptx'),
  });
}

build().then(function (f) { console.log('wrote', f); }).catch(function (e) {
  console.error(e);
  process.exit(1);
});
