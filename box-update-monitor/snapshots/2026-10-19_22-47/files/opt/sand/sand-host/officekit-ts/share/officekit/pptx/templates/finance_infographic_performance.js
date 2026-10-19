/**
 * "Financial Infographic" — 20-slide deck rebuilt with pptxgenjs.
 * Run: node 00097b98-1350-413a-8855-266d8a8b59d3_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Palette / typography (from the deck theme)
 * ------------------------------------------------------------------ */
const C = {
  navy: '142030',      // accent1
  pink: 'FF5B8D',      // accent2
  plum: '732654',      // accent3
  steel: '203644',     // accent4
  navyDark: '0F1824',
  white: 'FFFFFF',
  page: 'F2F2F2',      // slide background (bg1 lumMod 95%)
  ink: '0D0D0D',       // tx1 lumMod 95%
  graphite: '262626',
  slate: '595959',
  body: '000000',
  grey: '808080',
  greyLt: 'A6A6A6',
  greyMd: 'BFBFBF',
  silver: 'D9D9D9',
  track: 'EDEDED',
  pinkPale: 'FFDEE8',
  pinkDeep: 'FF0451',
  bluePale: 'E0E8F2',
  blueSoft: 'B1C4DE',
  bluePastel: 'C1D0E5',
};

const FONT = { head: 'Lexend', body: 'Poppins Light' };

const SHADOW = { type: 'outer', color: '000000', opacity: 0.1, blur: 80, offset: 30, angle: 35 };
const SHADOW_SM = { type: 'outer', color: '000000', opacity: 0.1, blur: 50, offset: 20, angle: 45 };

const LOREM_LONG = 'Lorem ipsum dolor sit sed do amet, consectetur adipiscing elit, ad sed do amet eiusmod tempor incididunt ut labore';
const LOREM_MID = 'Lorem ipsum dolor sit sed do amet, consectetur adipiscing elit, ad sed do amet eiusmod';
const LOREM_SHORT = 'Lorem ipsum dolor sit sed do amet, consectetur adipiscing.';
const LOREM_TEMPOR = 'Lorem ipsum dolor sit sed do amet, consectetur adipiscing elit, ad sed do amet eiusmod tempor.';

/* ------------------------------------------------------------------ *
 * Small builders shared by many slides
 * ------------------------------------------------------------------ */

// Decorative 5-square checkerboard chevron used as a corner motif.
const CHEVRON = {
  left: [[0, 2], [1, 1], [1, 3], [2, 0], [2, 4]],
  right: [[0, 0], [0, 4], [1, 1], [1, 3], [2, 2]],
  up: [[0, 0], [1, 1], [2, 2], [3, 1], [4, 0]],
  down: [[0, 2], [1, 1], [2, 0], [3, 1], [4, 2]],
};
const DECO_FILL = { color: '000000', transparency: 90 };

function chevron(slide, dir, x, y, cw, ch, fills) {
  CHEVRON[dir].forEach(([col, row], i) => {
    slide.addShape('rect', {
      x: x + col * cw, y: y + row * ch, w: cw, h: ch,
      fill: Array.isArray(fills) ? { color: fills[i] } : (fills || DECO_FILL),
    });
  });
}

// White rounded "card" with the deck's soft drop shadow.
function card(slide, x, y, w, h, radius, opts) {
  slide.addShape('roundRect', Object.assign({
    x, y, w, h, rectRadius: radius, fill: { color: C.white }, shadow: SHADOW,
  }, opts || {}));
}

/* ------------------------------------------------------------------ *
 * Text helpers — every text box in the deck is top-anchored and uses
 * the theme fonts, so route all text through one place.
 * ------------------------------------------------------------------ */
function text(slide, content, opts) {
  slide.addText(content, Object.assign({ valign: 'top', fontFace: FONT.body, color: C.body }, opts));
}

function head(slide, content, opts) {
  text(slide, content, Object.assign({ fontFace: FONT.head, fontSize: 32, lineSpacingMultiple: 1.0 }, opts));
}

function label(slide, content, opts) {
  text(slide, content, Object.assign({ fontFace: FONT.head, fontSize: 16, color: C.ink }, opts));
}

function para(slide, content, opts) {
  text(slide, content, Object.assign({ fontSize: 12, align: 'justify', lineSpacingMultiple: 1.3 }, opts));
}

// Bulleted list: each line needs its own bullet flag in pptxgenjs.
function bullets(lines) {
  return lines.map((t) => ({ text: t, options: { bullet: { characterCode: '2022', indent: 13.5 }, breakLine: true } }));
}

// Rounded speech-bubble marker ("Value Title / 80%") with a small pointer.
function callout(slide, x, y, w, h, dir) {
  const tip = 0.08;
  const bodyY = dir === 'up' ? y + tip : y;
  slide.addShape('roundRect', {
    x, y: bodyY, w, h: h - tip, rectRadius: 0.05,
    fill: { color: C.white }, shadow: SHADOW_SM,
  });
  slide.addShape('triangle', {
    x: x + w / 2 - 0.06, y: dir === 'up' ? y : y + h - tip, w: 0.12, h: tip,
    fill: { color: C.white }, rotate: dir === 'up' ? 0 : 180,
  });
}

// Circular badge with a small up-arrow, used on the KPI cards.
function arrowBadge(slide, x, y, color, flip) {
  slide.addShape('roundRect', {
    x, y, w: 0.333, h: 0.333, rectRadius: 0.1665,
    fill: { color: C.white }, line: { color, width: 0.5 }, shadow: SHADOW_SM,
  });
  slide.addShape('upArrow', {
    x: x + 0.091, y: y + 0.082, w: 0.151, h: 0.167, fill: { color }, flipV: !!flip,
  });
}

/* ------------------------------------------------------------------ *
 * Flat vector icon glyphs (drawn from primitives, no bitmaps)
 * ------------------------------------------------------------------ */
function icon(slide, kind, cx, cy, size, color) {
  const s = size, x = cx - s / 2, y = cy - s / 2, fill = { color };
  if (kind === 'briefcase') {
    slide.addShape('roundRect', { x, y: y + 0.2 * s, w: s, h: 0.75 * s, rectRadius: 0.06 * s, fill });
    slide.addShape('rect', { x: x + 0.26 * s, y, w: 0.48 * s, h: 0.16 * s, fill });
    slide.addShape('rect', { x: x + 0.26 * s, y: y + 0.1 * s, w: 0.1 * s, h: 0.14 * s, fill });
    slide.addShape('rect', { x: x + 0.64 * s, y: y + 0.1 * s, w: 0.1 * s, h: 0.14 * s, fill });
  } else if (kind === 'pie') {
    // three-quarter disc plus a detached slice
    slide.addShape('pie', { x, y, w: s, h: s, fill, angleRange: [275, 185] });
    slide.addShape('pie', { x: x + 0.08 * s, y: y - 0.08 * s, w: s, h: s, fill, angleRange: [190, 270] });
  } else if (kind === 'bars') {
    [[0.05, 0.45], [0.37, 0.62], [0.69, 0.85]].forEach(([bx, bh]) => {
      slide.addShape('rect', { x: x + bx * s, y: y + (0.9 - bh) * s, w: 0.26 * s, h: bh * s, fill });
    });
    slide.addShape('rect', { x, y: y + 0.92 * s, w: s, h: 0.1 * s, fill });
  } else if (kind === 'layers') {
    [0, 0.22, 0.44].forEach((dy, i) => {
      slide.addShape('diamond', { x, y: y + dy * s + (i === 0 ? 0 : 0.12 * s), w: s, h: 0.56 * s, fill });
    });
  } else if (kind === 'bag') {
    slide.addShape('blockArc', { x: x + 0.2 * s, y: y - 0.02 * s, w: 0.6 * s, h: 0.6 * s, fill, angleRange: [180, 0], arcThicknessRatio: 0.22 });
    slide.addShape('roundRect', { x, y: y + 0.26 * s, w: s, h: 0.74 * s, rectRadius: 0.1 * s, fill });
  } else if (kind === 'phone') {
    slide.addShape('blockArc', { x, y, w: s, h: s, fill, angleRange: [200, 340], arcThicknessRatio: 0.55, rotate: 45 });
  } else if (kind === 'pin') {
    slide.addShape('ellipse', { x, y, w: s, h: s * 0.86, fill });
    slide.addShape('triangle', { x: x + 0.28 * s, y: y + 0.62 * s, w: 0.44 * s, h: 0.42 * s, fill, rotate: 180 });
    slide.addShape('ellipse', { x: x + 0.33 * s, y: y + 0.26 * s, w: 0.34 * s, h: 0.34 * s, fill: { color: C.pink } });
  } else if (kind === 'person') {
    slide.addShape('ellipse', { x: x + 0.3 * s, y, w: 0.4 * s, h: 0.28 * s, fill });
    slide.addShape('roundRect', { x: x + 0.2 * s, y: y + 0.31 * s, w: 0.6 * s, h: 0.42 * s, rectRadius: 0.12 * s, fill });
    slide.addShape('rect', { x: x + 0.3 * s, y: y + 0.66 * s, w: 0.15 * s, h: 0.34 * s, fill });
    slide.addShape('rect', { x: x + 0.55 * s, y: y + 0.66 * s, w: 0.15 * s, h: 0.34 * s, fill });
  } else if (kind === 'personF') {
    slide.addShape('ellipse', { x: x + 0.3 * s, y, w: 0.4 * s, h: 0.28 * s, fill });
    slide.addShape('triangle', { x: x + 0.12 * s, y: y + 0.31 * s, w: 0.76 * s, h: 0.45 * s, fill, rotate: 180 });
    slide.addShape('rect', { x: x + 0.3 * s, y: y + 0.72 * s, w: 0.15 * s, h: 0.28 * s, fill });
    slide.addShape('rect', { x: x + 0.55 * s, y: y + 0.72 * s, w: 0.15 * s, h: 0.28 * s, fill });
  }
}

// Pale circle + glyph, the deck's standard bullet icon.
function iconCircle(slide, kind, x, y, d, bg, fg) {
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color: bg } });
  icon(slide, kind, x + d / 2, y + d / 2, d * 0.47, fg);
}

/* ------------------------------------------------------------------ *
 * Chart option presets
 * ------------------------------------------------------------------ */
const AXIS_TEXT = { catAxisLabelFontFace: FONT.body, valAxisLabelFontFace: FONT.body };

function pctBar(slide, x, y, w, h, colors) {
  slide.addChart('bar', [
    { name: 'Series 1', labels: ['Category 1'], values: [7] },
    { name: 'Series 2', labels: ['Category 1'], values: [3] },
  ], {
    x, y, w, h,
    barDir: 'bar', barGrouping: 'percentStacked', barGapWidthPct: 14,
    chartColors: colors, showLegend: false,
    catAxisHidden: true, valAxisHidden: true, valGridLine: { style: 'none' },
  });
}

function ring(slide, x, y, w, h, holeSize, values, colors) {
  slide.addChart('doughnut', [{ name: 'Sales', labels: values.map((_, i) => `Part ${i + 1}`), values }], {
    x, y, w, h, holeSize, firstSliceAng: 0,
    chartColors: colors, showLegend: false, showValue: false, dataBorder: { pt: 0, color: C.white },
  });
}

// Same as ring(), but the caller gives the drawn circle's box: the chart frame is
// padded out because pptxgenjs auto-fits the plot inside its frame.
function ringAt(slide, cx, cy, d, holeSize, values, colors) {
  const f = d * 1.05;
  ring(slide, cx - f / 2, cy - f / 2, f, f, holeSize, values, colors);
}

/* ================================================================== *
 * SLIDES
 * ================================================================== */

// 1 — Cover
function slide01(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.navy };
  pageNum(s, 1);
  chevron(s, 'left', 9.201, 0, 1.378, 1.5, [C.pink, C.white, C.white, C.pink, C.pink]);
  text(s, 'Presentation Template', { x: 0.389, y: 0.346, w: 3.431, h: 0.37, fontFace: FONT.body, fontSize: 16, color: C.white });
  text(s, 'Financial Infographic', {
    x: 0.754, y: 4.281, w: 8.566, h: 2.524,
    fontFace: FONT.head, fontSize: 80, color: C.white, lineSpacingMultiple: 0.9,
  });
}

// 2 — Three KPI cards with mini progress bars
function slide02(pptx) {
  const s = newSlide(pptx, 2);
  chevron(s, 'down', 0, 4.921, 0.936, 0.86);
  head(s, 'Visual Insights That Reflect Our Financial Strength', { x: 0.986, y: 0.723, w: 11.36, h: 0.64, align: 'center' });

  const cards = [
    { x: 1.155, dark: false, pct: '75%', bars: [C.pink, C.track] },
    { x: 4.956, dark: true, pct: '89%', bars: [C.white, 'FF87AC'] },
    { x: 8.757, dark: false, pct: '70%', bars: [C.pink, C.track] },
  ];
  cards.forEach((c) => {
    card(s, c.x, 1.98, 3.421, 4.281, 0.115, { fill: { color: c.dark ? C.pink : C.white } });
    pctBar(s, c.x + 0.328, 3.053, 2.786, 0.402, c.bars);
    text(s, c.pct, {
      x: c.x + 0.375, y: 2.206, w: 2.703, h: 0.912,
      fontFace: FONT.head, fontSize: 40, bold: true, color: c.dark ? C.white : C.navy, lineSpacingMultiple: 1.3,
    });
    label(s, 'Your Text Here', { x: c.x + 0.384, y: 3.405, w: 2.694, h: 0.37, color: c.dark ? C.white : C.ink });
    para(s, LOREM_LONG, { x: c.x + 0.375, y: 4.648, w: 2.694, h: 1.133, color: c.dark ? C.white : C.body });
  });
}

// 3 — "Current Activity" combo chart + result card
function slide03(pptx) {
  const s = newSlide(pptx, 3);
  chevron(s, 'left', 11.184, 0, 0.716, 0.78);
  head(s, 'Uncovering Business Growth Through Financial Visualization', { x: 0.792, y: 0.7, w: 8.162, h: 1.178 });

  card(s, 0.943, 2.482, 7.833, 3.901, 0.117);
  text(s, 'Current Activity', { x: 1.141, y: 2.615, w: 4.723, h: 0.507, fontFace: FONT.head, fontSize: 20, bold: true, color: C.ink, lineSpacingMultiple: 1.3 });

  const years = ['2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023'];
  s.addChart([
    {
      type: pptx.ChartType.bar,
      data: [
        { name: 'Series 1', labels: years, values: [15, 16, 30, 27, 44, 25, 16, 50] },
        { name: 'Series 2', labels: years, values: [17, 14, 28, 25, 46, 27, 14, 48] },
      ],
      options: { chartColors: [C.navy, C.pink], barGapWidthPct: 219, barOverlapPct: -27 },
    },
    {
      type: pptx.ChartType.line,
      data: [{ name: 'Series 3', labels: years, values: [34, 28, 56, 50, 72, 54, 28, 74] }],
      options: {
        chartColors: [C.greyMd], lineSize: 1.5, lineDataSymbolSize: 5,
        lineDataSymbolLineColor: C.greyMd, lineSmooth: false,
      },
    },
  ], Object.assign({
    x: 1.236, y: 3.295, w: 7.248, h: 2.88,
    showLegend: false, valAxisMaxVal: 100,
    catAxisLabelColor: C.slate, catAxisLabelFontSize: 9, catAxisLineColor: 'E4E4E4',
    valAxisLabelColor: C.slate, valAxisLabelFontSize: 9, valAxisLineShow: false,
    valGridLine: { color: 'E8E8E8', size: 0.25 },
  }, AXIS_TEXT));

  callout(s, 7.249, 2.916, 1.333, 0.826, 'down');
  text(s, 'Dec 2025', { x: 7.319, y: 2.968, w: 1.193, h: 0.344, fontFace: FONT.head, fontSize: 12, color: C.graphite, align: 'center', lineSpacingMultiple: 1.3 });
  text(s, '10.000', { x: 7.319, y: 3.178, w: 1.193, h: 0.426, fontFace: FONT.head, fontSize: 16, bold: true, color: C.navy, align: 'center', lineSpacingMultiple: 1.3 });
  s.addShape('ellipse', { x: 7.801, y: 3.97, w: 0.229, h: 0.229, fill: { color: C.pink, transparency: 80 } });
  s.addShape('ellipse', { x: 7.866, y: 4.035, w: 0.098, h: 0.098, fill: { color: C.pink } });

  card(s, 9.067, 4.833, 3.31, 1.55, 0.052);
  text(s, 'Analytical Results', { x: 9.416, y: 5.053, w: 2.576, h: 0.37, fontFace: FONT.body, fontSize: 16, color: '1F1F1F' });
  text(s, '239,308,984,00', { x: 9.416, y: 5.423, w: 2.576, h: 0.438, fontFace: FONT.head, fontSize: 20, bold: true, color: C.pink });
  text(s, 'Dec 10, 2025 at 8:02am', { x: 9.416, y: 5.876, w: 2.474, h: 0.303, fontFace: FONT.body, fontSize: 12, italic: true, color: C.navy });
}

// 4 — Four KPI tiles + doughnut panel
function slide04(pptx) {
  const s = newSlide(pptx, 4);
  chevron(s, 'down', 0, 5.841, 0.606, 0.557);
  head(s, 'Analyzing Financial Performance Through Diagrams', { x: 0.749, y: 0.667, w: 7.416, h: 1.178, color: '1F1F1F' });

  const tiles = [
    { x: 0.86, y: 2.361, value: '40,214', color: C.navy, badge: C.navy, flip: true },
    { x: 4.467, y: 2.377, value: '51,297', color: C.pink, badge: C.pink, flip: false },
    { x: 0.86, y: 4.507, value: '98,162', color: C.pink, badge: C.steel, flip: false },
    { x: 4.467, y: 4.507, value: '93,397', color: C.navy, badge: C.plum, flip: true },
  ];
  tiles.forEach((t) => {
    card(s, t.x, t.y, 3.448, 1.987, 0.0756);
    text(s, t.value, {
      x: t.x + 0.313, y: t.y + 0.385, w: 2.288, h: 0.64,
      fontFace: FONT.head, fontSize: 32, bold: true, color: t.color,
    });
    para(s, LOREM_SHORT, { x: t.x + 0.313, y: t.y + 1.055, w: 2.801, h: 0.608 });
    arrowBadge(s, t.x + 2.945, t.y + 0.121, t.badge, t.flip);
  });

  card(s, 8.324, 1.748, 4.149, 4.742, 0.125);
  text(s, 'Your Subtitle Here', { x: 8.595, y: 2.015, w: 3.095, h: 0.417, fontFace: FONT.head, fontSize: 16, bold: true, color: C.ink, lineSpacingMultiple: 1.3 });
  para(s, 'Lorem ipsum dolor sit sed do amet, consectetur adipiscing elit.', { x: 8.62, y: 2.446, w: 3.07, h: 0.608 });
  ringAt(s, 10.389, 4.722, 2.653, 75, [8.2, 3.2, 1.4, 1.2], [C.navy, C.pink, C.navy, C.pink]);

  // two value flags pinned to the ring
  callout(s, 9.361, 4.029, 1.17, 0.449, 'up');
  text(s, '$1,400,000', { x: 9.403, y: 4.15, w: 1.087, h: 0.272, fontFace: 'Inter SemiBold', fontSize: 12, color: C.plum });
  callout(s, 10.266, 5.122, 1.17, 0.449, 'down');
  text(s, '$1,400,000', { x: 10.307, y: 5.179, w: 1.087, h: 0.272, fontFace: 'Inter SemiBold', fontSize: 12, color: C.navy, align: 'right' });
}

// 5 — Wide monthly line chart + two footnotes
function slide05(pptx) {
  const s = newSlide(pptx, 5);
  chevron(s, 'down', 9.457, 5.361, 0.776, 0.713);
  head(s, 'PLACEHOLDER',
    { x: 0.749, y: 0.667, w: 11.294, h: 1.178, align: 'center' });

  card(s, 0.924, 2.374, 11.486, 2.751, 0.105);
  s.addChart('line', [{
    name: 'Series 1',
    labels: ['Oct 2020', 'Nov 2020', 'Dec 2020', 'Jan 2021', 'Feb 2021', 'Mar 2021', 'Apr 2021',
      'May 2021', 'Jun 2021', 'Jul 2021', 'Aug 2021', 'Sep 2021', 'Oct 2021', 'Nov 2021'],
    values: [1.5, 4.6, 2.4, 5.4, 4.23, 5.13, 3.9, 6.5, 8.5, 7.32, 4.7, 8.56, 3.89, 6.45],
  }], Object.assign({
    x: 0.924, y: 2.499, w: 11.31, h: 2.502,
    chartColors: [C.navy], lineSize: 2.25, lineSmooth: false,
    lineDataSymbol: 'circle', lineDataSymbolSize: 5, lineDataSymbolLineColor: C.navy,
    showLegend: false,
    catAxisLabelColor: C.ink, catAxisLabelFontSize: 10, catAxisLineColor: 'E4E4E4',
    valAxisLabelColor: C.ink, valAxisLabelFontSize: 10, valAxisLineShow: false,
    valAxisMajorUnit: 2, valAxisLabelFormatCode: '"$"#,##0.00',
    valGridLine: { color: 'EDEDED', size: 0.75, style: 'dash' },
  }, AXIS_TEXT));

  [{ x: 0.924, kind: 'briefcase' }, { x: 5.06, kind: 'bag' }].forEach((f) => {
    s.addShape('ellipse', { x: f.x, y: 5.691, w: 0.822, h: 0.822, fill: { color: C.pink } });
    icon(s, f.kind, f.x + 0.411, 6.102, 0.38, C.white);
    text(s, 'Your Subtitle Here', { x: f.x + 1.041, y: 5.655, w: 3.095, h: 0.417, fontFace: FONT.head, fontSize: 16, bold: true, color: C.ink, lineSpacingMultiple: 1.3 });
    para(s, LOREM_SHORT, { x: f.x + 1.041, y: 6.066, w: 2.694, h: 0.608 });
  });
}

// 6 — Growth funnel + summary strip
function slide06(pptx) {
  const s = newSlide(pptx, 6);
  chevron(s, 'up', 0, 0.004, 0.598, 0.549);
  head(s, 'Aligning Financial Performance With Business Vision', { x: 4.663, y: 0.597, w: 7.514, h: 1.178, align: 'right' });

  card(s, 1.157, 2.156, 11.02, 2.855, 0.139);
  s.addShape('triangle', { x: 6.179, y: -1.481, w: 1.284, h: 9.463, fill: { color: C.silver }, rotate: 270 });

  [4.497, 8.019, 11.545].forEach((x) => {
    s.addShape('line', { x, y: 2.293, w: 0, h: 2.425, line: { color: C.greyMd, width: 1 } });
  });
  s.addShape('line', { x: 1.865, y: 3.251, w: 9.912, h: 0, line: { color: C.greyMd, width: 1 } });

  text(s, 'Growth', { x: 1.988, y: 2.269, w: 2.979, h: 0.423, fontFace: FONT.head, fontSize: 16, color: C.ink, lineSpacingMultiple: 1.3 });
  text(s, '6,292+', { x: 1.988, y: 2.664, w: 1.795, h: 0.48, fontFace: FONT.head, fontSize: 22.5, color: C.ink });

  const steps = [
    { x: 2.115, ly: 3.484, py: 3.756, pw: 2.395, ph: 0.87, text: 'Pretium aenean pharetra magna ac placerat. Eget gravida cum sociis' },
    { x: 5.009, ly: 3.747, py: 4.018, pw: 2.893, ph: 0.608, text: 'Pretium aenean pharetra magna ac placerat. ' },
    { x: 8.217, ly: 4.009, py: 4.281, pw: 3.246, ph: 0.345, text: 'Pretium aenean pharetra magna.' },
  ];
  steps.forEach((st) => {
    text(s, 'Your Text Here', { x: st.x, y: st.ly, w: 2.142, h: 0.337, fontFace: FONT.head, fontSize: 14, color: C.ink, valign: 'middle' });
    text(s, st.text, { x: st.x, y: st.py, w: st.pw, h: st.ph, fontFace: FONT.body, fontSize: 12, color: C.ink, lineSpacingMultiple: 1.3 });
  });

  s.addShape('ellipse', { x: 4.456, y: 3.087, w: 0.086, h: 0.325, fill: { color: C.navy } });
  s.addShape('ellipse', { x: 7.94, y: 2.846, w: 0.159, h: 0.807, fill: { color: C.pink } });
  s.addShape('ellipse', { x: 11.467, y: 2.606, w: 0.159, h: 1.287, fill: { color: C.plum } });

  card(s, 1.157, 5.349, 5.994, 1.314, 0.108, { shadow: SHADOW_SM });
  s.addShape('line', { x: 4.009, y: 5.661, w: 0, h: 0.689, line: { color: C.greyLt, width: 1, dashType: 'dash' } });
  iconCircle(s, 'briefcase', 1.502, 5.656, 0.74, C.blueSoft, C.navy);
  text(s, '$3,261', { x: 2.387, y: 5.652, w: 1.607, h: 0.438, fontFace: FONT.head, fontSize: 20, bold: true, color: C.navy });
  text(s, 'Your Text Here', { x: 2.387, y: 6.1, w: 1.709, h: 0.303, fontFace: FONT.body, fontSize: 12, color: C.ink });
  iconCircle(s, 'pie', 4.36, 5.656, 0.74, C.pinkPale, C.pinkDeep);
  text(s, '$56,840', { x: 5.205, y: 5.652, w: 1.709, h: 0.438, fontFace: FONT.head, fontSize: 20, bold: true, color: C.pink });
  text(s, 'Your Text Here', { x: 5.205, y: 6.1, w: 1.709, h: 0.303, fontFace: FONT.body, fontSize: 12, color: C.ink });

  label(s, 'Your Text Here', { x: 7.903, y: 5.543, w: 2.694, h: 0.37 });
  para(s, LOREM_TEMPOR, { x: 7.903, y: 5.861, w: 4.274, h: 0.608 });
}

// 7 — Gauge card + numbered notes
function slide07(pptx) {
  const s = newSlide(pptx, 7);
  chevron(s, 'up', 10.346, 0, 0.598, 0.549);
  card(s, 0.859, 1.0, 3.076, 5.5, 0.117);

  const gauges = [
    { y: 1.581, value: '33.50', type: 'Type A', color: C.navy, end: 274, typeColor: C.navy, valColor: C.ink, capY: 2.978 },
    { y: 4.317, value: '29.50', type: 'Type B', color: C.pink, end: 345, typeColor: C.ink, valColor: C.graphite, capY: 5.714 },
  ];
  gauges.forEach((g) => {
    s.addShape('arc', { x: 1.45, y: g.y, w: 1.785, h: 1.785, angleRange: [161, 19], line: { color: C.greyLt, width: 7.5, transparency: 75 } });
    s.addShape('arc', { x: 1.45, y: g.y, w: 1.785, h: 1.785, angleRange: [161, g.end], line: { color: g.color, width: 7.5 } });
    text(s, g.type, { x: 1.593, y: g.y + 0.292, w: 1.5, h: 0.375, fontFace: FONT.head, fontSize: 12, color: g.typeColor, align: 'center', lineSpacingMultiple: 1.5 });
    text(s, g.value, { x: 1.485, y: g.y + 0.582, w: 1.716, h: 0.572, fontFace: FONT.head, fontSize: 28, color: g.valColor, align: 'center' });
    s.addShape('line', { x: 1.968, y: g.y + 1.215, w: 0.75, h: 0, line: { color: C.greyLt, width: 1, transparency: 60 } });
    text(s, 'Data Analysis – 12,70%', { x: 1.018, y: g.capY, w: 2.756, h: 0.375, fontFace: FONT.head, fontSize: 12, color: C.ink, align: 'center', lineSpacingMultiple: 1.5 });
  });
  s.addShape('line', { x: 1.154, y: 3.75, w: 2.377, h: 0, line: { color: C.greyLt, width: 1, dashType: 'dash', transparency: 50 } });

  head(s, 'Data That Drives Growth: A Financial Visualization Summary', { x: 4.794, y: 2.292, w: 7.521, h: 1.178 });

  [{ x: 4.926, tx: 5.8, num: '01', color: C.navy }, { x: 8.936, tx: 9.781, num: '02', color: C.pink }].forEach((n) => {
    s.addShape('roundRect', { x: n.x, y: 4.025, w: 0.653, h: 0.617, rectRadius: 0.074, fill: { color: n.color } });
    text(s, n.num, { x: n.x + 0.048, y: 4.119, w: 0.653, h: 0.438, fontFace: FONT.head, fontSize: 20, color: C.white });
    label(s, 'Your Text Here', { x: n.tx, y: 4.021, w: 2.694, h: 0.37 });
    para(s, LOREM_MID, { x: n.tx, y: 4.339, w: 2.694, h: 0.87 });
  });
}

// 8 — Pie chart with four callouts
function slide08(pptx) {
  const s = newSlide(pptx, 8);
  chevron(s, 'up', 10.346, 0, 0.598, 0.549);
  head(s, 'PLACEHOLDER', { x: 0.849, y: 0.74, w: 9.415, h: 1.178 });

  s.addShape('ellipse', { x: 0.801, y: 2.386, w: 4.225, h: 4.225, fill: { color: C.white }, shadow: SHADOW });
  s.addChart('pie', [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'], values: [0.7, 0.2, 0.1, 0.12] }], {
    x: 0.958, y: 2.52, w: 3.877, h: 3.877,
    chartColors: [C.navy, C.pink, C.navy, C.pink], showLegend: false,
    showValue: true, dataLabelColor: C.white, dataLabelFontSize: 12, dataLabelFontFace: FONT.head,
    dataLabelFormatCode: '0%', dataLabelPosition: 'ctr',
  });
  s.addShape('ellipse', { x: 2.443, y: 4.072, w: 0.94, h: 0.94, fill: { color: C.white }, shadow: SHADOW });

  const notes = [
    { x: 5.378, y: 2.476, pct: '70%', color: C.navy },
    { x: 5.378, y: 4.674, pct: '20%', color: C.pink },
    { x: 9.189, y: 2.476, pct: '12%', color: C.pink },
    { x: 9.189, y: 4.674, pct: '10%', color: C.navy },
  ];
  notes.forEach((n) => {
    text(s, n.pct, { x: n.x, y: n.y, w: 1.336, h: 0.64, fontFace: FONT.head, fontSize: 32, bold: true, color: n.color });
    label(s, 'Your Text Here', { x: n.x + 1.454, y: n.y + 0.135, w: 2.046, h: 0.37 });
    para(s, LOREM_MID, { x: n.x, y: n.y + 0.766, w: 3.292, h: 0.87 });
  });
}

// 9 — Clustered column chart with bulleted legend
function slide09(pptx) {
  const s = newSlide(pptx, 9);
  chevron(s, 'right', 0, 0, 0.549, 0.598);
  head(s, 'Strategic Financial Insights That Drive Long-Term Value Creation', { x: 1.869, y: 0.697, w: 9.595, h: 1.178, align: 'center' });

  card(s, 1.0, 2.293, 11.333, 2.914, 0.0978);
  const cats = ['Data 1', 'Data 2', 'Data 3', 'Data 4'];
  s.addChart('bar', [
    { name: 'Series 1', labels: cats, values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels: cats, values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels: cats, values: [2, 2, 3, 5] },
  ], Object.assign({
    x: 1.198, y: 2.446, w: 10.937, h: 2.608,
    chartColors: [C.navy, C.pink, C.plum], barGapWidthPct: 219, barOverlapPct: -27, showLegend: false,
    catAxisLabelColor: C.grey, catAxisLabelFontSize: 12, catAxisLineColor: 'E4E4E4',
    valAxisLabelColor: C.grey, valAxisLabelFontSize: 12, valAxisLineShow: false, valAxisMajorUnit: 2,
    valGridLine: { color: 'EDEDED', size: 0.75 },
  }, AXIS_TEXT));

  cats.forEach((name, i) => {
    const x = 1.98 + i * 2.625;
    text(s, name, { x, y: 5.529, w: 1.659, h: 0.37, fontFace: FONT.head, fontSize: 16, color: C.body });
    text(s, bullets(['Lorem ipsum dolor ', 'Sit sed do amet ', 'consectetur adipiscing']), {
      x: x - 0.028, y: 5.919, w: 2.395, h: 0.87, fontFace: FONT.body, fontSize: 12, color: C.body,
      align: 'justify', lineSpacingMultiple: 1.3,
    });
  });
}

// 10 — Two overlapping stat circles + benefit rows
function slide10(pptx) {
  const s = newSlide(pptx, 10);
  chevron(s, 'down', 9.452, 5.361, 0.776, 0.713);
  head(s, 'Leveraging Financial Intelligence to Empower Strategic Decision-Making', { x: 0.9, y: 1.361, w: 5.95, h: 1.717 });

  s.addShape('ellipse', { x: 7.444, y: 1.361, w: 2.941, h: 2.941, fill: { color: C.navy } });
  text(s, '80%', { x: 8.264, y: 2.071, w: 1.528, h: 0.774, fontFace: FONT.head, fontSize: 40, color: C.white });
  text(s, 'Your Text Here', { x: 7.956, y: 2.845, w: 2.046, h: 0.37, fontFace: FONT.head, fontSize: 16, color: C.white, align: 'center' });

  s.addShape('ellipse', { x: 9.125, y: 3.139, w: 2.941, h: 2.941, fill: { color: C.pink } });
  text(s, '75%', { x: 9.95, y: 3.976, w: 1.528, h: 0.774, fontFace: FONT.head, fontSize: 40, color: C.white });
  text(s, 'Your Text Here', { x: 9.642, y: 4.75, w: 2.046, h: 0.37, fontFace: FONT.head, fontSize: 16, color: C.white, align: 'center' });

  [{ y: 3.625, name: 'Benefit One', color: C.body }, { y: 5.075, name: 'Benefit Two', color: C.pink }].forEach((b) => {
    text(s, b.name, { x: 0.9, y: b.y, w: 1.804, h: 0.37, fontFace: FONT.head, fontSize: 16, color: b.color });
    text(s, '80%', { x: 0.902, y: b.y + 0.37, w: 1.319, h: 0.707, fontFace: FONT.head, fontSize: 36, color: b.color });
    para(s, LOREM_MID, { x: 2.704, y: b.y + 0.061, w: 3.076, h: 0.87 });
  });
}

// 11 — Revenue segmentation: three KPI cards + wide line chart
function slide11(pptx) {
  const s = newSlide(pptx, 11);
  chevron(s, 'right', -0.023, 0, 0.549, 0.598);
  head(s, 'Revenue Segmentation Overview ', { x: 2.119, y: 0.635, w: 9.096, h: 0.64, align: 'center' });

  const tiles = [
    { x: 1.0, radius: 0.192, value: '40,214', color: C.navy, badge: C.navy },
    { x: 4.943, radius: 0.151, value: '91,720', color: C.pink, badge: C.pink },
    { x: 8.885, radius: 0.151, value: '80,910', color: C.navy, badge: C.plum },
  ];
  tiles.forEach((t) => {
    card(s, t.x, 1.773, 3.448, 1.987, t.radius);
    arrowBadge(s, t.x + 2.945, 1.877, t.badge, false);
    text(s, t.value, { x: t.x + 0.313, y: 2.122, w: 2.288, h: 0.64, fontFace: FONT.head, fontSize: 32, bold: true, color: t.color });
    para(s, LOREM_SHORT, { x: t.x + 0.313, y: 2.806, w: 2.783, h: 0.608 });
  });

  card(s, 1.0, 4.178, 11.333, 2.646, 0.189);
  s.addChart('line', [{
    name: 'Series 1',
    labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
    values: [1.5, 4.6, 2.4, 5.4, 4.23, 5.13, 3.9, 6.5, 8.5, 7.32, 4.7, 8.56],
  }], Object.assign({
    x: 1.229, y: 4.365, w: 10.875, h: 2.271,
    chartColors: [C.navy], lineSize: 2.5, lineSmooth: false,
    lineDataSymbol: 'circle', lineDataSymbolSize: 5, lineDataSymbolLineColor: C.navy,
    showLegend: false,
    catAxisLabelColor: C.ink, catAxisLabelFontSize: 10, catAxisLineColor: 'E4E4E4',
    valAxisLabelColor: C.ink, valAxisLabelFontSize: 10, valAxisLineShow: false,
    valAxisMajorUnit: 2, valAxisLabelFormatCode: '"$"#,##0.00',
    valGridLine: { color: 'EDEDED', size: 0.5 },
  }, AXIS_TEXT));

  callout(s, 7.281, 5.282, 1.76, 0.826, 'up');
  text(s, 'Value Title Here', { x: 7.343, y: 5.431, w: 1.635, h: 0.345, fontFace: FONT.body, fontSize: 12, color: C.grey, align: 'center', lineSpacingMultiple: 1.3 });
  text(s, '80%', { x: 7.564, y: 5.628, w: 1.193, h: 0.426, fontFace: FONT.head, fontSize: 16, bold: true, color: C.navy, align: 'center', lineSpacingMultiple: 1.3 });
}

// 12 — Three staggered ring cards + bullet list
function slide12(pptx) {
  const s = newSlide(pptx, 12);
  chevron(s, 'right', 0, 4.517, 0.549, 0.598);
  head(s, 'Unveiling the Numbers Behind Our Strategic Business Growth', { x: 7.123, y: 0.718, w: 5.493, h: 1.717, align: 'right' });

  const panels = [
    { x: 0.718, y: 0.849, radius: 0.172, ring: C.navy, num: '70', numSize: 14, bold: false, value: C.navy },
    { x: 4.69, y: 2.611, radius: 0.191, ring: C.pink, num: '80', numSize: 14, bold: false, value: C.pink },
    { x: 8.662, y: 3.661, radius: 0.158, ring: C.navy, num: '70', numSize: 17.5, bold: true, value: C.navy },
  ];
  panels.forEach((p, i) => {
    card(s, p.x, p.y, 3.646, 2.901, p.radius);
    ring(s, p.x + 0.215, p.y + 0.175, 1.278, 1.831, 60, [22, 78], ['E4E4E4', p.ring]);
    text(s, p.num, {
      x: p.x + 0.496, y: p.y + 0.875, w: 0.686, h: 0.42,
      fontFace: FONT.body, fontSize: p.numSize, bold: p.bold, color: p.ring, align: i === 2 ? 'center' : 'left',
    });
    text(s, '6,292', { x: p.x + 1.63, y: p.y + 0.526, w: 1.795, h: 0.5, fontFace: FONT.head, fontSize: 22.5, color: p.value });
    text(s, 'Your Section Here', { x: p.x + 1.63, y: p.y + 1.032, w: 1.299, h: 0.505, fontFace: FONT.head, fontSize: 12, color: C.ink });
    text(s, '80%', { x: p.x + 2.967, y: p.y + 1.332, w: 0.6, h: 0.401, fontFace: FONT.head, fontSize: 12.5, color: C.body, valign: 'middle' });
    s.addShape('line', { x: p.x + 0.177, y: p.y + 1.882, w: 3.269, h: 0, line: { color: '8A9AA8', width: 1, transparency: 59 } });
    para(s, 'Lorem ipsum dolor sit sed do amet, consectetur adipiscing elit', { x: p.x + 0.349, y: p.y + 2.052, w: 3.076, h: 0.608 });
  });

  label(s, 'Your Text Here', { x: 0.823, y: 4.585, w: 2.694, h: 0.37 });
  text(s, bullets(['Lorem ipsum dolor sit amet', 'Consectetur adipiscing elit ', 'Sed do eiusmod tempor ', 'Incididunt ut labore et dolore']), {
    x: 0.824, y: 4.992, w: 3.049, h: 1.133, fontFace: FONT.body, fontSize: 12, color: C.body, lineSpacingMultiple: 1.3,
  });
}

// 13 — Horizontal bar chart + two analysis cards
function slide13(pptx) {
  const s = newSlide(pptx, 13);
  chevron(s, 'up', 10.346, 0, 0.598, 0.549);
  head(s, 'PLACEHOLDER', { x: 1.019, y: 0.732, w: 9.049, h: 1.178 });

  card(s, 1.13, 2.413, 7.178, 4.124, 0.138);
  text(s, 'Performance Indicator', { x: 1.549, y: 2.591, w: 3.831, h: 0.426, fontFace: FONT.head, fontSize: 16, bold: true, color: C.ink, lineSpacingMultiple: 1.3 });
  const months = ['Jan', 'Feb', 'Mar', 'Apr'];
  s.addChart('bar', [
    { name: 'Series 2', labels: months, values: [3, 4, 2, 3.2] },
    { name: 'Series 3', labels: months, values: [3, 2, 3.4, 4] },
  ], Object.assign({
    x: 1.477, y: 2.982, w: 6.527, h: 3.31,
    barDir: 'bar', chartColors: [C.pink, C.navy], barGapWidthPct: 150, showLegend: false,
    catAxisLabelColor: C.slate, catAxisLabelFontSize: 12, catAxisLineColor: 'E4E4E4',
    valAxisLabelColor: C.slate, valAxisLabelFontSize: 12, valAxisLineShow: false,
    valGridLine: { color: 'E4E4E4', size: 0.75 },
  }, AXIS_TEXT));

  [{ y: 2.195, title: 'Data and Progress', tw: 2.069 }, { y: 4.581, title: 'Analysis Profit', tw: 2.303 }].forEach((p) => {
    card(s, 8.726, p.y, 3.477, 2.174, 0.121);
    text(s, p.title, { x: 8.917, y: p.y + 0.3, w: p.tw, h: 0.386, fontFace: FONT.head, fontSize: 14, bold: true, color: C.ink, lineSpacingMultiple: 1.3 });
    [['Analysis 01', '25%', 0.91], ['Analysis 02', '45%', 1.534]].forEach(([name, pct, dy]) => {
      text(s, name, { x: 8.935, y: p.y + dy, w: 1.817, h: 0.345, fontFace: FONT.body, fontSize: 12, color: C.ink, lineSpacingMultiple: 1.3 });
      text(s, pct, { x: 10.986, y: p.y + dy, w: 1.027, h: 0.345, fontFace: FONT.head, fontSize: 12, bold: true, color: C.ink, align: 'right', lineSpacingMultiple: 1.3 });
    });
    s.addShape('line', { x: 9.039, y: p.y + 1.391, w: 2.876, h: 0, line: { color: C.greyLt, width: 1, dashType: 'dash', transparency: 50 } });
  });
}

// 14 — Customer relation ring + market line chart
function slide14(pptx) {
  const s = newSlide(pptx, 14);
  chevron(s, 'left', 11.707, 2.256, 0.549, 0.598);
  head(s, 'PLACEHOLDER', { x: 6.196, y: 1.467, w: 6.286, h: 1.717 });

  card(s, 0.913, 0.98, 4.697, 2.64, 0.145);
  text(s, 'Customer Relation', { x: 1.177, y: 1.141, w: 3.75, h: 0.426, fontFace: FONT.head, fontSize: 16, bold: true, color: C.ink, lineSpacingMultiple: 1.3 });
  ringAt(s, 2.083, 2.528, 1.371, 69, [90, 10], [C.navy, C.pink]);
  text(s, 'Q1', { x: 1.623, y: 2.302, w: 0.914, h: 0.46, fontFace: FONT.head, fontSize: 18, color: C.graphite, align: 'center', valign: 'middle' });

  [{ y: 1.67, color: C.navy, kind: 'person' }, { y: 2.541, color: C.pink, kind: 'personF' }].forEach((r) => {
    icon(s, r.kind, 3.26, r.y + 0.31, 0.44, r.color);
    text(s, '+86,214', { x: 3.419, y: r.y, w: 1.601, h: 0.496, fontFace: FONT.head, fontSize: 20, color: r.color, valign: 'bottom' });
    text(s, 'Total User', { x: 3.419, y: r.y + 0.426, w: 1.658, h: 0.338, fontFace: FONT.body, fontSize: 12, color: C.ink, lineSpacingMultiple: 1.3 });
  });

  card(s, 0.913, 3.899, 6.163, 2.64, 0.118);
  s.addChart('line', [{ name: 'Sep-17 (Hundred Thousand)', labels: ['1st', '2nd', '3rd', '4th'], values: [150, 180, 125, 160] }], Object.assign({
    x: 1.099, y: 4.513, w: 3.699, h: 1.669,
    chartColors: [C.navy], lineSize: 2.25, lineSmooth: false,
    lineDataSymbol: 'circle', lineDataSymbolSize: 5, lineDataSymbolLineColor: C.navy,
    showLegend: true, legendPos: 'b', legendFontSize: 9, legendFontFace: FONT.body, legendColor: C.slate,
    catAxisLabelColor: C.slate, catAxisLabelFontSize: 9, catAxisLineColor: 'E4E4E4',
    valAxisHidden: true, valAxisMinVal: 100, valAxisMaxVal: 200, valAxisMajorUnit: 20,
    valGridLine: { style: 'none' },
  }, AXIS_TEXT));
  callout(s, 2.965, 4.271, 1.197, 0.644, 'down');
  text(s, 'Option ', { x: 2.856, y: 4.323, w: 1.401, h: 0.303, fontFace: FONT.head, fontSize: 12, color: C.ink, align: 'center' });
  text(s, '3.514.000', { x: 2.925, y: 4.575, w: 1.265, h: 0.286, fontFace: FONT.head, fontSize: 10.5, bold: true, color: C.navy, align: 'center' });

  text(s, '2024, Market', { x: 5.06, y: 4.449, w: 2.215, h: 0.37, fontFace: FONT.head, fontSize: 16, color: C.ink });
  text(s, '@JohnDoe', { x: 5.06, y: 4.764, w: 2.032, h: 0.34, fontFace: FONT.body, fontSize: 12, bold: true, color: C.ink, lineSpacingMultiple: 1.3 });
  s.addShape('line', { x: 4.985, y: 5.499, w: 1.874, h: 0, line: { color: C.greyLt, width: 1, dashType: 'dash', transparency: 50 } });
  text(s, '2.00 ETH', { x: 5.06, y: 5.355, w: 2.032, h: 0.37, fontFace: FONT.head, fontSize: 16, color: C.navy });
  text(s, 'Sold', { x: 5.06, y: 5.649, w: 2.032, h: 0.34, fontFace: FONT.body, fontSize: 12, color: C.slate, lineSpacingMultiple: 1.3 });

  label(s, 'Your Text Here', { x: 8.078, y: 4.619, w: 2.694, h: 0.37 });
  para(s, LOREM_LONG, { x: 8.078, y: 5.064, w: 4.342, h: 0.87 });
}

// 15 — Stacked-disc illustration + icon list
function slide15(pptx) {
  const s = newSlide(pptx, 15);
  chevron(s, 'down', 6.623, 5.854, 0.598, 0.549);
  head(s, 'Simplifying Financial Data Into Intuitive Visual Insights', { x: 0.779, y: 0.786, w: 6.721, h: 1.178 });

  // horizontal "disc stack": back ellipse + body + front face, drawn back-to-front
  function disc(xLeft, w, y, h, faceW, bodyColor, faceColor) {
    s.addShape('ellipse', { x: xLeft, y, w: faceW, h, fill: { color: bodyColor } });
    s.addShape('rect', { x: xLeft + faceW / 2, y, w: w - faceW, h, fill: { color: bodyColor } });
    s.addShape('ellipse', { x: xLeft + w - faceW, y, w: faceW, h, fill: { color: faceColor } });
  }
  s.addShape('leftArrow', { x: 7.176, y: 3.588, w: 1.385, h: 0.878, fill: { color: C.grey } });
  s.addShape('ellipse', { x: 8.536, y: 2.988, w: 0.539, h: 2.22, fill: { color: C.grey } });
  disc(8.511, 1.484, 2.698, 2.8, 0.68, C.navy, C.greyLt);
  disc(9.273, 1.767, 2.274, 3.649, 0.884, C.pink, C.greyMd);
  disc(10.118, 1.986, 1.846, 4.504, 1.087, C.navy, C.navyDark);
  s.addShape('ellipse', { x: 11.393, y: 3.409, w: 0.335, h: 1.378, fill: { color: C.navy } });
  s.addShape('rect', { x: 11.393, y: 3.887, w: 1.161, h: 0.421, fill: { color: C.grey } });

  [{ x: 8.648, kind: 'layers', bg: C.bluePale, fg: C.navy },
   { x: 9.445, kind: 'pie', bg: C.pinkPale, fg: C.pinkDeep },
   { x: 10.241, kind: 'briefcase', bg: C.bluePale, fg: C.navy }].forEach((b) => {
    iconCircle(s, b.kind, b.x, 3.719, 0.601, b.bg, b.fg);
  });

  const rows = [
    { y: 2.726, kind: 'briefcase', bg: C.bluePale, fg: C.navy, tx: 1.937, px: 1.969 },
    { y: 4.043, kind: 'pie', bg: C.pinkPale, fg: C.pinkDeep, tx: 1.92, px: 1.952 },
    { y: 5.359, kind: 'layers', bg: C.bluePale, fg: C.navy, tx: 1.903, px: 1.935 },
  ];
  rows.forEach((r, i) => {
    iconCircle(s, r.kind, [0.877, 0.86, 0.838][i], r.y, 0.671, r.bg, r.fg);
    label(s, 'Your Text Here', { x: r.tx, y: r.y, w: 2.694, h: 0.37 });
    para(s, LOREM_TEMPOR, { x: r.px, y: r.y + 0.383, w: 4.342, h: 0.608 });
  });
}

// 16 — Column chart card + two icon panels
function slide16(pptx) {
  const s = newSlide(pptx, 16);
  chevron(s, 'up', 9.293, 0, 0.808, 0.742);
  head(s, 'Transforming Data Into Design That Informs', { x: 0.617, y: 0.715, w: 6.721, h: 1.178 });

  card(s, 0.63, 2.387, 5.845, 4.238, 0.142, { shadow: SHADOW_SM });
  text(s, 'Subjects Status', { x: 0.776, y: 2.626, w: 5.555, h: 0.426, fontFace: FONT.head, fontSize: 16, bold: true, color: C.graphite, lineSpacingMultiple: 1.3 });
  s.addChart('bar', [{
    name: 'Series 1',
    labels: ['Data 1', 'Data 2', 'Data 3', 'Data 4', 'Data 5', 'Data 6', 'Data 7', 'Data 8'],
    values: [9, 5, 10, 6, 7, 5, 8, 7],
  }], Object.assign({
    x: 0.872, y: 3.519, w: 5.555, h: 2.868,
    chartColors: [C.navy, C.pink, C.navy, C.pink, C.navy, C.pink, C.navy, C.pink],
    barGapWidthPct: 150, showLegend: false,
    catAxisLabelColor: C.slate, catAxisLabelFontSize: 10, catAxisLineColor: 'E4E4E4',
    valAxisLabelColor: C.slate, valAxisLabelFontSize: 10, valAxisMaxVal: 10, valAxisMajorUnit: 5,
    valAxisLineShow: false, valGridLine: { color: 'EDEDED', size: 0.75 },
  }, AXIS_TEXT));

  [{ x: 6.718, tile: 6.944, tileY: 2.616, fill: C.navy, kind: 'pie', px: 6.957 },
   { x: 9.853, tile: 10.092, tileY: 2.626, fill: C.pink, kind: 'bars', px: 10.057 }].forEach((p) => {
    card(s, p.x, 2.387, 2.894, 4.238, 0.097, { shadow: SHADOW_SM });
    s.addShape('roundRect', { x: p.tile, y: p.tileY, w: 0.791, h: 0.791, rectRadius: 0.0815, fill: { color: p.fill } });
    icon(s, p.kind, p.tile + 0.396, p.tileY + 0.396, 0.42, C.white);
    text(s, '+8192', { x: p.tile, y: 4.481, w: 2.042, h: 0.572, fontFace: FONT.head, fontSize: 28, color: C.body });
    para(s, 'Lorem ipsum dolor sit sed do amet, lab consectetur adipiscing elit, ad sed do amet eiusmod tempor',
      { x: p.px, y: 5.15, w: 2.404, h: 1.133 });
  });
}

// 17 — Skill bars + two mini trend cards
function slide17(pptx) {
  const s = newSlide(pptx, 17);
  chevron(s, 'left', 11.694, 4.512, 0.549, 0.598);
  head(s, 'Growth Projection by Sector ', { x: 0.793, y: 0.878, w: 4.347, h: 1.178 });
  label(s, 'Your Text Here', { x: 0.797, y: 2.328, w: 2.694, h: 0.37 });
  para(s, LOREM_LONG, { x: 0.797, y: 2.773, w: 4.342, h: 0.87 });

  const skills = [
    { y: 4.663, name: 'Creative Skill  01', pct: '85%', fill: 2.952, color: C.navy },
    { y: 5.488, name: 'Creative Skill  02', pct: '60%', fill: 2.063, color: C.pink },
    { y: 6.314, name: 'Creative Skill  03', pct: '72%', fill: 2.514, color: C.plum },
  ];
  skills.forEach((k) => {
    text(s, k.name, { x: 0.816, y: k.y - 0.399, w: 2.25, h: 0.342, fontFace: FONT.head, fontSize: 12, color: C.ink, lineSpacingMultiple: 1.3 });
    text(s, k.pct, { x: 3.919, y: k.y - 0.371, w: 0.703, h: 0.337, fontFace: FONT.head, fontSize: 14, color: C.ink });
    s.addShape('roundRect', { x: 0.892, y: k.y, w: 3.637, h: 0.093, rectRadius: 0.0465, fill: { color: C.greyLt, transparency: 80 } });
    s.addShape('roundRect', { x: 0.892, y: k.y, w: k.fill, h: 0.093, rectRadius: 0.0465, fill: { color: k.color } });
    s.addShape('ellipse', {
      x: 0.892 + k.fill - 0.146, y: k.y - 0.04, w: 0.172, h: 0.172,
      fill: { color: k.color }, line: { color: C.white, width: 5 }, shadow: SHADOW_SM,
    });
  });

  card(s, 5.662, 1.866, 6.875, 3.771, 0.127);
  text(s, 'Learn With Effectively With Us', { x: 5.955, y: 2.093, w: 6.146, h: 0.426, fontFace: FONT.head, fontSize: 16, bold: true, color: C.ink, lineSpacingMultiple: 1.3 });

  const panels = [
    { x: 6.097, y: 2.795, line: C.navy, dot: C.navy, curve: [[0.12, 0.92], [0.3, 0.63], [0.45, 0.7], [0.62, 0.42], [0.86, 0.05]] },
    { x: 9.29, y: 2.766, line: C.pink, dot: C.plum, curve: [[0.1, 0.95], [0.3, 0.7], [0.5, 0.28], [0.68, 0.42], [0.88, 0.05]] },
  ];
  panels.forEach((p) => {
    s.addShape('roundRect', { x: p.x, y: p.y, w: 2.811, h: 2.487, rectRadius: 0.267, fill: { color: 'F2F2F2', transparency: 70 } });
    text(s, 'Data Analysis', { x: p.x + 0.306, y: p.y + 0.083, w: 2.198, h: 0.414, fontFace: FONT.head, fontSize: 16, color: C.ink, lineSpacingMultiple: 1.3 });
    const gx = p.x + 0.339, gy = p.y + 0.869, gw = 2.131, gh = 1.315;
    for (let i = 0; i < 5; i++) {
      s.addShape('line', { x: gx, y: gy + i * (gh / 4), w: gw, h: 0, line: { color: C.greyLt, width: 2.25, transparency: 80 } });
    }
    const pts = p.curve.map(([fx, fy]) => ({ x: gx + fx * gw, y: gy + fy * gh }));
    s.addShape('custGeom', {
      x: gx, y: gy, w: gw, h: gh, line: { color: p.line, width: 3.25 },
      points: pts.map((pt, i) => (i === 0 ? { x: pt.x - gx, y: pt.y - gy, moveTo: true } : {
        x: pt.x - gx, y: pt.y - gy,
        curve: { type: 'cubic', x1: pts[i - 1].x - gx + 0.16, y1: pts[i - 1].y - gy, x2: pt.x - gx - 0.16, y2: pt.y - gy },
      })),
    });
    [pts[1], pts[4]].forEach((pt) => {
      s.addShape('ellipse', {
        x: pt.x - 0.062, y: pt.y - 0.057, w: 0.124, h: 0.114,
        fill: { color: p.dot }, line: { color: p.dot, width: 3.5, transparency: 50 },
      });
    });
  });
}

// 18 — Segmented progress bar + three statistic cards
function slide18(pptx) {
  const s = newSlide(pptx, 18);
  chevron(s, 'right', -0.011, 0, 0.549, 0.598);
  head(s, 'Progress and Performance Indicators Across Market Segments', { x: 2.493, y: 0.878, w: 8.346, h: 1.178, align: 'center' });

  card(s, 1.0, 2.612, 11.333, 2.254, 0.0756);

  // 9 segments: 3 navy, 3 pink, 3 plum — rounded caps at both ends
  const segX = [3.882, 4.805, 5.727, 6.65, 7.572, 8.495, 9.418, 10.34, 11.263];
  segX.forEach((x, i) => {
    const color = i < 3 ? C.navy : i < 6 ? C.pink : C.plum;
    if (i === 0) {
      s.addShape('round2SameRect', { x: x + 0.216, y: 3.5585, w: 0.405, h: 0.836, rectRadius: 0.2025, fill: { color }, rotate: 270 });
    } else if (i === 8) {
      s.addShape('round2SameRect', { x: x + 0.216, y: 3.5585, w: 0.405, h: 0.836, rectRadius: 0.2025, fill: { color }, rotate: 90 });
    } else {
      s.addShape('rect', { x, y: 3.77, w: 0.836, h: 0.405, fill: { color } });
    }
  });
  text(s, 'Pretium aenean pharetra magna ac placerat. Eget gravida cum sociis natoque ',
    { x: 4.268, y: 4.299, w: 7.445, h: 0.34, fontFace: FONT.body, fontSize: 12, color: C.ink, align: 'center', lineSpacingMultiple: 1.3 });

  s.addShape('roundRect', {
    x: 1.234, y: 2.819, w: 2.218, h: 1.84, rectRadius: 0.0827,
    fill: { color: C.white, transparency: 100 }, line: { color: C.greyLt, width: 0.5, transparency: 50 },
  });
  iconCircle(s, 'briefcase', 1.963, 2.97, 0.762, C.bluePastel, C.navyDark);
  text(s, 'Timing and Progress', { x: 1.452, y: 3.772, w: 1.783, h: 0.776, fontFace: FONT.head, fontSize: 16, color: C.navy, align: 'center', lineSpacingMultiple: 1.3 });

  [{ x: 4.556, pct: '30%', color: C.navy }, { x: 7.324, pct: '50%', color: C.pink }, { x: 10.092, pct: '40%', color: C.plum }].forEach((v) => {
    callout(s, v.x, 2.819, 1.333, 0.826, 'down');
    text(s, 'Value Title', { x: v.x + 0.07, y: 2.853, w: 1.193, h: 0.339, fontFace: FONT.body, fontSize: 12, color: C.ink, align: 'center', lineSpacingMultiple: 1.3 });
    text(s, v.pct, { x: v.x + 0.07, y: 3.082, w: 1.193, h: 0.426, fontFace: FONT.head, fontSize: 16, bold: true, color: v.color, align: 'center', lineSpacingMultiple: 1.3 });
  });

  const stats = [
    { x: 1.071, y: 5.201, year: '2022 Statistic', yearColor: C.ink, rate: '11m/Month', value: '+$210', color: C.navy },
    { x: 5.001, y: 5.197, year: '2023 Statistic', yearColor: C.ink, rate: '21m/Month', value: '+$810', color: C.pink },
    { x: 8.931, y: 5.199, year: '2024 Statistic', yearColor: C.navy, rate: '41m/Month', value: '+$540', color: C.plum },
  ];
  stats.forEach((st) => {
    card(s, st.x, st.y, 3.342, 1.173, 0.0536);
    text(s, st.year, { x: st.x + 0.128, y: st.y + 0.097, w: 1.897, h: 0.385, fontFace: FONT.head, fontSize: 14, color: st.yearColor, lineSpacingMultiple: 1.3 });
    text(s, st.rate, { x: st.x + 0.128, y: st.y + 0.415, w: 1.769, h: 0.304, fontFace: FONT.head, fontSize: 10, color: C.ink, lineSpacingMultiple: 1.3 });
    text(s, st.value, { x: st.x + 1.622, y: st.y + 0.246, w: 1.512, h: 0.507, fontFace: FONT.head, fontSize: 20, bold: true, color: st.color, align: 'right', valign: 'bottom', lineSpacingMultiple: 1.3 });
    text(s, 'Your Text Here', { x: st.x + 1.622, y: st.y + 0.7, w: 1.512, h: 0.345, fontFace: FONT.body, fontSize: 12, color: C.ink, align: 'right', lineSpacingMultiple: 1.3 });
  });
}

// 19 — 100% stacked bar chart with value labels
function slide19(pptx) {
  const s = newSlide(pptx, 19);
  chevron(s, 'down', 0, 5.836, 0.598, 0.549);
  head(s, 'Revenue Growth Chart: From Stability to Expansion', { x: 0.923, y: 0.762, w: 6.341, h: 1.178 });
  label(s, 'Your Text Here', { x: 8.301, y: 0.807, w: 2.694, h: 0.37 });
  para(s, LOREM_TEMPOR, { x: 8.301, y: 1.252, w: 4.342, h: 0.608 });

  card(s, 0.691, 2.162, 11.952, 4.335, 0.0956);
  const cats = ['Category 1', 'Category 2', 'Category 3', 'Category 4'];
  s.addChart('bar', [
    { name: 'Series 1', labels: cats, values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels: cats, values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels: cats, values: [2, 2, 3, 5] },
  ], Object.assign({
    x: 1.119, y: 2.341, w: 11.095, h: 3.978,
    barDir: 'bar', barGrouping: 'percentStacked', barGapWidthPct: 79,
    chartColors: [C.navy, C.pink, C.navy], showLegend: false,
    showValue: true, dataLabelColor: C.white, dataLabelFontSize: 10.6, dataLabelFontFace: FONT.body,
    dataLabelPosition: 'ctr', dataLabelFormatCode: 'General',
    catAxisLabelColor: C.ink, catAxisLabelFontSize: 10.6, catAxisLabelFontBold: true, catAxisLineColor: 'E4E4E4',
    valAxisHidden: true, valGridLine: { style: 'none' }, catGridLine: { color: 'E4E4E4', size: 0.75 },
  }, AXIS_TEXT));
}

// 20 — Closing slide
function slide20(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.navy };
  pageNum(s, 20);
  chevron(s, 'down', 0, 4.589, 1.057, 0.97, [C.pink, C.white, C.pink, C.white, C.pink]);
  text(s, 'Thank You for Your Attention', {
    x: 4.227, y: 1.257, w: 8.538, h: 2.524,
    fontFace: FONT.head, fontSize: 80, color: C.white, lineSpacingMultiple: 0.9,
  });
  s.addShape('flowChartConnector', { x: 4.535, y: 3.903, w: 0.551, h: 0.528, fill: { color: C.pink } });
  icon(s, 'pin', 4.808, 4.155, 0.3, C.body);
  text(s, '12 Your Street Name, 1234', { x: 5.284, y: 3.999, w: 3.231, h: 0.337, fontFace: FONT.body, fontSize: 14, color: C.white });
  s.addShape('flowChartConnector', { x: 8.713, y: 3.903, w: 0.551, h: 0.528, fill: { color: C.pink } });
  icon(s, 'phone', 8.988, 4.177, 0.26, C.body);
  text(s, '08951739149/ 067 526/172', { x: 9.464, y: 3.999, w: 2.71, h: 0.337, fontFace: FONT.body, fontSize: 14, color: C.white });
}

/* ------------------------------------------------------------------ *
 * Deck assembly
 * ------------------------------------------------------------------ */

// Every slide carries a large, near-transparent page number (from the master).
function pageNum(slide, num) {
  slide.addText(String(num), {
    x: 11.29, y: 6.527, w: 1.631, h: 0.64,
    fontFace: FONT.head, fontSize: 32, bold: true, color: C.body, transparency: 85, align: 'right',
  });
}

// Content slides share the light page background.
function newSlide(pptx, num) {
  const s = pptx.addSlide();
  s.background = { color: C.page };
  pageNum(s, num);
  return s;
}

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE';
  pptx.author = 'pptxgenjs';
  pptx.title = 'Financial Infographic';
  pptx.theme = { headFontFace: FONT.head, bodyFontFace: FONT.body };

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach((fn) => fn(pptx));

  return pptx.writeFile({ fileName: path.join(__dirname, '00097b98-1350-413a-8855-266d8a8b59d3_grok_final.pptx') });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
