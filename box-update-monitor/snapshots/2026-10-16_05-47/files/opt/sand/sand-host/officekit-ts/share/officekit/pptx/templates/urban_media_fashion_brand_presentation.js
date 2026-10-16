/**
 * SCARS 20/20 - urban culture presentation template (35 slides, 26.67 x 15 in).
 * Standalone rebuild with pptxgenjs. Photographs in the original deck are
 * replaced by flat colour placeholders labelled [image].
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette
const O = 'E24824';        // signature orange
const W = 'FFFFFF';
const BLACK = '000000';
const BG = '332F2E';       // deck background
const PANEL = '3B3838';    // dark panel behind photos
const INK = '262626';      // near-black logo panel
const GREY = 'A6A6A6';     // body copy
const DIM = '808080';      // ghosted repeated headline
const MAPGREY = '767171';  // map silhouettes
const LIGHT = 'F2F2F2';    // light photo placeholder
const MID = '999999';      // mid-grey photo placeholder

const HEAD = 'Russo One';
const BODY = 'Open Sans';

// ---------------------------------------------------------------- text styles
const S = {
  tag:     { fontFace: HEAD, fontSize: 20,  charSpacing: 3 },
  strip:   { fontFace: HEAD, fontSize: 24,  charSpacing: 6 },
  kicker:  { fontFace: HEAD, fontSize: 28,  charSpacing: 6 },
  sub28:   { fontFace: HEAD, fontSize: 28 },
  label:   { fontFace: HEAD, fontSize: 36,  charSpacing: 6 },
  head72:  { fontFace: HEAD, fontSize: 72 },
  head96:  { fontFace: HEAD, fontSize: 96 },
  head96s: { fontFace: HEAD, fontSize: 96,  charSpacing: 3 },
  head96w: { fontFace: HEAD, fontSize: 96,  charSpacing: 6 },
  mega:    { fontFace: HEAD, fontSize: 287, charSpacing: 32 },
  hero:    { fontFace: HEAD, fontSize: 344, charSpacing: 32 },
  body:    { fontFace: BODY, fontSize: 20 },
  stat28:  { fontFace: BODY, fontSize: 28 },
  stat54:  { fontFace: BODY, fontSize: 54 },
  stat72:  { fontFace: BODY, fontSize: 72 },
  semi20:  { fontFace: 'Open Sans Semibold', fontSize: 20 },
  brand:   { fontFace: HEAD, fontSize: 20, charSpacing: 2 },
};

// ---------------------------------------------------------------- filler copy
const LOREM = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Morbi posuere eros non ipsum ' +
  'viverra dignissim. Pellentesque metus tortor, ullamcorper in orci ut, euismod hendrerit arcu. ' +
  'Sed justo ligula, blandit eget ante non, convallis egestas ex. Aenean orci urna, venenatis ' +
  'vitae condimentum sit amet, aliquet quis enim. In hac habitasse platea dictumst. Proin ' +
  'dignissim, velit sit amet convallis aliquet, diam lectus tincidunt nulla, fringilla consequat ' +
  'lacus erat vitae sapien. Ut mollis nisi justo, et vulputate sem iaculis vel. Sed ornare ' +
  'egestas libero eget suscipit. Ut ex quam, sollicitudin ac scelerisque vehicula, aliquet ' +
  'euismod felis. Praesent a congue justo. Integer hendrerit sapien quis scelerisque ullamcorper. ' +
  'Quisque tempus, ligula non eleifend mollis, ligula neque finibus est, non gravida mauris lacus ' +
  'id dolor. ';
const QUIS = 'Quisque porta lacinia ligula. Maecenas lorem justo, varius id ex nec, elementum molestie ' +
  'massa. Suspendisse laoreet rhoncus ipsum. Morbi vulputate semper pulvinar. Donec et eros neque. ' +
  'Morbi lacinia lacus et nisl faucibus, vitae rhoncus dolor efficitur. Nulla lobortis lectus eu ' +
  'turpis dapibus, non suscipit leo maximus. Donec ullamcorper, nisl eu auctor lobortis, metus ex ' +
  'commodo enim, sit amet consequat quam metus in ex. Fusce eu lectus ante. Duis erat diam, rutrum ' +
  'a lacinia vel, dictum vel metus. Maecenas maximus risus vitae semper cursus. Sed at diam sed ' +
  'risus ultrices auctor. In hendrerit nisl ut sapien pulvinar, nec tincidunt quam laoreet. Donec ' +
  'mi mi, bibendum vel placerat eget, eleifend nec purus. Etiam id sapien quis mi dapibus lacinia. ' +
  'Fusce pulvinar neque justo, ut bibendum elit interdum ut. Donec dapibus ligula mi, in pulvinar ' +
  'dolor cong';
const PRAESENT = 'Praesent eros ex, sodales vitae lorem ut, luctus elementum eros. Mauris lacus lorem, commodo';
const BR = null;   // marks a line break between runs

// ---------------------------------------------------------------- primitives
const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'SCARS', width: 26.67, height: 15 });
pptx.layout = 'SCARS';

/** Solid shape. `opts.shape` picks the geometry (default rect); rest is merged. */
function BOX(s, x, y, w, h, color, opts) {
  const { shape = 'rect', ...rest } = opts || {};
  s.addShape(pptx.ShapeType[shape],
    Object.assign({ x, y, w, h, fill: { color }, line: { type: 'none' } }, rest));
}

/** Photo placeholder: flat rectangle carrying an "[image]" caption. */
function IMG(s, x, y, w, h, tone) {
  s.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: tone }, line: { type: 'none' } });
  s.addText('[image]', {
    x, y, w, h, align: 'center', valign: 'middle',
    fontFace: HEAD, fontSize: Math.max(11, Math.min(26, w * 2.5)),
    color: tone === LIGHT ? 'BFBFBF' : '737373',
  });
}

/**
 * Device mockup (tablet / phone / watch / laptop). Drawn as a hollow rounded
 * bezel so the screen placeholder already on the slide stays visible.
 */
function DEVICE(s, x, y, w, h) {
  const bez = Math.min(w, h) * 0.06;
  s.addShape(pptx.ShapeType.roundRect, {
    x: x + bez / 2, y: y + bez / 2, w: w - bez, h: h - bez,
    rectRadius: Math.min(w, h) * 0.09,
    fill: { type: 'none' }, line: { color: '1A1A1A', width: bez * 72 },
  });
}

/**
 * Rich text block. `runs` is a list of [text, color] pairs; the BR marker
 * starts a new paragraph.
 */
function T(s, x, y, w, h, runs, opts) {
  const o = Object.assign({ x, y, w, h, valign: 'top' }, opts);
  const parts = [];
  runs.forEach(r => {
    if (r === BR) { if (parts.length) parts[parts.length - 1].options.breakLine = true; return; }
    parts.push({ text: r[0], options: { color: r[1], breakLine: false } });
  });
  s.addText(parts, o);
}

/** Page number in the bottom gutter, e.g. "//7". */
function PAGE(s, n) {
  T(s, 10.89, 13.854, 4.89, 0.44,
    [['//', O], [String(n), W]], { ...S.tag, align: 'center' });
}

/** Grey world/US map silhouette stand-in for the vector map artwork. */
function MAP(s, x, y, w, h) {
  s.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: MAPGREY }, line: { type: 'none' } });
  s.addText('[map]', { x, y, w, h, align: 'center', valign: 'middle',
    fontFace: HEAD, fontSize: 24, color: '8F8B8B' });
}

/** Orange location pin. */
function PIN(s, x, y) {
  BOX(s, x, y, 0.313, 0.313, O, { shape: 'ellipse' });
  BOX(s, x + 0.115, y + 0.24, 0.083, 0.19, O);
  BOX(s, x + 0.104, y + 0.09, 0.105, 0.105, W, { shape: 'ellipse' });
}

// social glyphs (the originals are freeform vector artwork)
function GLYPH(s, x, y, w, ch) {
  s.addText(ch, { x: x - w, y: y - w, w: w * 3, h: w * 3, align: 'center', valign: 'middle',
    fontFace: BODY, fontSize: w * 130, bold: true, color: W });
}
function ICON_TW(s, x, y, w) { GLYPH(s, x, y, w, 't'); }
function ICON_FB(s, x, y, w) { GLYPH(s, x, y, w, 'f'); }
function ICON_IG(s, x, y, w) {
  s.addShape(pptx.ShapeType.roundRect, { x, y, w, h: w, rectRadius: 0.08,
    fill: { type: 'none' }, line: { color: W, width: 1.5 } });
  s.addShape(pptx.ShapeType.ellipse, { x: x + w * 0.27, y: y + w * 0.27, w: w * 0.46, h: w * 0.46,
    fill: { type: 'none' }, line: { color: W, width: 1.5 } });
}

// ---------------------------------------------------------------- charts
const CHART_BASE = {
  showLegend: false, showTitle: false, chartArea: { fill: { type: 'none' } },
  catAxisLabelFontFace: BODY, valAxisLabelFontFace: BODY,
};

/**
 * Slide 28 - yearly volume column chart. The original holds six clustered
 * series and only the fourth carries data, which is what pushes the thin
 * orange bars right of each category tick.
 */
function CHART_28(s) {
  const years = ['2013', '2014', '2015', '2016', '2017', '2018', '2019'];
  const empty = years.map(() => null);
  s.addChart(pptx.ChartType.bar, [
    { name: 'Column4', labels: years, values: empty },
    { name: 'Column3', labels: years, values: empty },
    { name: 'Series 1', labels: years, values: empty },
    { name: 'Series 2', labels: years, values: [30, 90, 60, 120, 30, 120, 130] },
    { name: 'Column1', labels: years, values: empty },
    { name: 'Column12', labels: years, values: empty },
  ], {
    ...CHART_BASE, x: 13.359, y: 4.206, w: 9.207, h: 5.098,
    barDir: 'col', barGapWidthPct: 219, barOverlapPct: -27,
    chartColors: ['5B9BD5', 'ED7D31', 'A5A5A5', O, '4472C4', '70AD47'],
    catAxisLabelColor: GREY, catAxisLabelFontSize: 12, catAxisLineColor: '807B7B',
    valAxisLabelColor: GREY, valAxisLabelFontSize: 12, valAxisLineShow: false,
    valAxisMaxVal: 140, valAxisMajorUnit: 20,
    valGridLine: { color: GREY, size: 0.75 }, catGridLine: { style: 'none' },
  });
}

/** Slide 32 - monthly selling progress lines. */
function CHART_32(s) {
  s.addChart(pptx.ChartType.line, [
    { name: 'Robusta',
      labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
      values: [0.05, 0.1, 0.15, 0.22, 0.3, 0.28, 0.22, 0.32, 0.45, 0.6, 0.65, 0.55] },
    { name: 'Arabica',
      labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
      values: [0.28, 0.38, 0.4, 0.48, 0.65, 0.65, 0.78, 0.85, 0.86, 0.85, 0.9, 0.85] },
  ], {
    ...CHART_BASE, x: 0.968, y: 6.344, w: 10.164, h: 7.51,
    chartColors: [O, W], lineSize: 2, lineDataSymbol: 'circle', lineDataSymbolSize: 8,
    lineSmooth: false,
    showValue: true, dataLabelPosition: 't', dataLabelFormatCode: '0%',
    dataLabelColor: DIM, dataLabelFontSize: 11, dataLabelFontFace: BODY,
    catAxisLabelColor: DIM, catAxisLabelFontSize: 11, catAxisLineShow: false,
    valAxisHidden: true, valAxisLabelFormatCode: '0%',
    catGridLine: { style: 'none' }, valGridLine: { style: 'none' },
    showLegend: true, legendPos: 'b', legendColor: DIM, legendFontSize: 11,
    legendFontFace: BODY,
  });
}

/** Slide 33 - stacked market area chart. */
function CHART_33(s) {
  const days = ['1/5/2019', '1/6/2019', '1/7/2019', '1/8/2019', '1/9/2019'];
  s.addChart(pptx.ChartType.area, [
    { name: 'Series 1', labels: days, values: [22, 32, 28, 12, 15] },
    { name: 'Series 2', labels: days, values: [12, 20, 12, 21, 28] },
  ], {
    ...CHART_BASE, x: 17.035, y: 7.869, w: 8.085, h: 5.663,
    chartColors: [MAPGREY, O],
    catAxisLabelColor: GREY, catAxisLabelFontSize: 12, catAxisLineColor: '4A4645',
    valAxisLabelColor: GREY, valAxisLabelFontSize: 12, valAxisLineShow: false,
    catGridLine: { style: 'none' }, valGridLine: { style: 'none' },
  });
}

/** Slide 1 */
function slide01(s) {
  BOX(s, 21, 0.75, 3.318, 10.438, O);
  IMG(s, 0, 0, 23.75, 10.639, LIGHT);
  T(s, 2.241, 8.465, 22.9, 5.89, [['SCARS', O]], { ...S.hero, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 2.423, 7.803, 3, 0.57, [['20/20', W]], { ...S.kicker, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 16.881, 7.213, 16.45, 0.57, [['PRESENTATION TEMPLATE ', W], ['//', O], [' PRESENTATION TEMPLATE', W]], { ...S.strip, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 90 });
}

/** Slide 2 */
function slide02(s) {
  IMG(s, 17.577, 0, 9.093, 15, LIGHT);
  BOX(s, 15.737, 3.187, 7.311, 8.625, O);
  PAGE(s, 2);
  T(s, 9.586, 7.231, 7.22, 0.57, [['SCARS - MEDIA', W]], { ...S.kicker, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 270 });
  T(s, 1.886, 4.189, 8.64, 6.66, [[LOREM, GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5 });
  T(s, 1.886, 1.18, 11.06, 1.72, [['//', O], ['ABOUT', W]], { fontFace: 'Russo One', fontSize: 96, charSpacing: 32, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  IMG(s, 13.752, 3.908, 8.544, 7.184, LIGHT);
}

/** Slide 3 */
function slide03(s) {
  BOX(s, 8.488, 1.434, 6.101, 1.156, O);
  BOX(s, 18.234, 5.679, 6.101, 1.156, O);
  BOX(s, 6.715, 9.451, 6.101, 1.156, O);
  PAGE(s, 3);
  T(s, 8.392, 2.811, 6.2, 1.62, [[LOREM.slice(0, 130) + '.', GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 21.854, 1.326, 3, 0.57, [['- SCARS', W]], { ...S.kicker, align: 'right', lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 8.803, 1.656, 5.47, 0.71, [['EDUCATION', W]], { ...S.label, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 7.142, 1.656, 1, 0.71, [['1', '595959']], { ...S.label, align: 'right', lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 18.135, 7.002, 6.2, 1.62, [[LOREM.slice(0, 130) + '.', GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 18.55, 5.901, 5.47, 0.71, [['PROGRESS', W]], { ...S.label, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 16.892, 5.901, 1, 0.71, [['2', '595959']], { ...S.label, align: 'right', lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 7.031, 9.673, 5.47, 0.71, [['PRODUCTIFITY', W]], { ...S.label, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 5.373, 9.673, 1, 0.71, [['3', '595959']], { ...S.label, align: 'right', lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 6.62, 10.774, 6.2, 1.62, [[LOREM.slice(0, 130) + '.', GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, -6.675, 4.682, 18.28, 4.93, [['GOALS', O]], { ...S.mega, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 270 });
  BOX(s, 14.8, 1.434, 0.308, 1.156, O);
  BOX(s, 24.546, 5.679, 0.308, 1.156, O);
  BOX(s, 13.027, 9.451, 0.308, 1.156, O);
  IMG(s, 18.234, 12.481, 6.62, 2.519, LIGHT);
}

/** Slide 4 */
function slide04(s) {
  PAGE(s, 4);
  T(s, 7.806, 1.436, 11.06, 1.72, [['//', O], ['AGENDA', W]], { fontFace: 'Russo One', fontSize: 96, charSpacing: 32, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 0.558, 11.403, 4.47, 1.62, [[LOREM.slice(0, 79), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 0.553, 5.148, 5.47, 1.31, [['URBAN PHOTO HUNT!', O]], { ...S.label, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 7.496, 11.403, 4.47, 1.62, [[LOREM.slice(0, 79), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 14.434, 11.403, 4.47, 1.62, [[LOREM.slice(0, 79), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 21.373, 11.403, 4.47, 1.62, [[LOREM.slice(0, 79), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 7.492, 5.148, 5.47, 1.31, [['DARK STREET STYLE', O]], { ...S.label, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 14.43, 5.148, 5.47, 1.31, [['MANHATTAN URBN FASHION', O]], { ...S.label, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 21.369, 5.148, 5.47, 1.31, [['HYPE BEAST PARADE', O]], { ...S.label, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 0.581, 4.595, 3, 0.44, [['02 / 02', W]], { fontFace: 'Russo One', fontSize: 20, charSpacing: 6, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 7.492, 4.595, 3, 0.44, [['12 / 02', W]], { fontFace: 'Russo One', fontSize: 20, charSpacing: 6, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 14.403, 4.595, 3, 0.44, [['21 / 02', W]], { fontFace: 'Russo One', fontSize: 20, charSpacing: 6, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 21.313, 4.595, 3, 0.44, [['24 / 02', W]], { fontFace: 'Russo One', fontSize: 20, charSpacing: 6, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  IMG(s, 0, 7.064, 6.711, 3.513, MID);
  IMG(s, 6.658, 7.064, 6.711, 3.513, LIGHT);
  IMG(s, 13.326, 7.064, 6.711, 3.513, MID);
  IMG(s, 19.993, 7.064, 6.711, 3.513, LIGHT);
}

/** Slide 5 */
function slide05(s) {
  PAGE(s, 5);
  T(s, -1.17, 0.482, 28.9, 4.93, [['MILESTONE', O]], { fontFace: 'Russo One', fontSize: 287, charSpacing: 6, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 1.189, 7.067, 6.2, 1.06, [[LOREM.slice(0, 85), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 1.189, 6.086, 5.47, 0.71, [['//', W], [' 2015', O]], { ...S.label, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 1.189, 11.085, 6.2, 1.06, [[LOREM.slice(0, 85), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 1.189, 10.104, 5.47, 0.71, [['//', W], [' 2018', O]], { ...S.label, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 10.457, 7.067, 6.2, 1.06, [[LOREM.slice(0, 85), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 10.457, 6.086, 5.47, 0.71, [['//', W], [' 2016', O]], { ...S.label, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 10.457, 11.085, 6.2, 1.06, [[LOREM.slice(0, 85), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 10.457, 10.104, 5.47, 0.71, [['//', W], [' 2019', O]], { ...S.label, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 19.245, 7.067, 6.2, 1.06, [[LOREM.slice(0, 85), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 19.245, 6.086, 5.47, 0.71, [['//', W], [' 2017', O]], { ...S.label, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 19.245, 11.085, 6.2, 1.06, [[LOREM.slice(0, 85), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 19.245, 10.104, 5.47, 0.71, [['//', W], [' 2020', O]], { ...S.label, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
}

/** Slide 6 */
function slide06(s) {
  PAGE(s, 6);
  T(s, 2.013, 3.059, 10.9, 4.65, [[LOREM.slice(0, 737) + '.', GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5 });
  T(s, 1.886, 8.396, 22.9, 4.95, [['Designing For Humans ', O], ['//', W], [' ', O], BR, ['An Introduction to Human -Centered Design', O]], { ...S.head96, lineSpacingMultiple: 1 });
  T(s, 13.517, 3.059, 10.9, 4.65, [[QUIS.slice(0, 713), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5 });
  T(s, 2.013, 1.795, 3, 0.57, [['20 / 20', W]], { ...S.kicker, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
}

/** Slide 7 */
function slide07(s) {
  PAGE(s, 7);
  T(s, 1.886, 3.722, 10.14, 5.15, [[LOREM.slice(0, 765) + '.', GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 15.661, 1.748, 8.73, 7.12, [[QUIS, GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 12.685, 4.01, 4.89, 0.57, [['SCARS - MEDIA', W]], { ...S.kicker, align: 'right', lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 270 });
  T(s, 1.886, 1.436, 11.06, 1.72, [['Decision', W], [' //', O]], { ...S.head96, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 5.026, 9.151, 25.29, 4.93, [['MAKERS', O]], { ...S.mega, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
}

/** Slide 8 */
function slide08(s) {
  IMG(s, 8.544, 5.325, 12.511, 7.184, LIGHT);
  IMG(s, 0, 5.325, 8.544, 7.184, MID);
  PAGE(s, 8);
  T(s, 1.896, 2.124, 8.99, 1.62, [[LOREM.slice(0, 180), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 14.8, 8.083, 11.06, 1.72, [['PORTOFOLIO', O]], { ...S.head96w, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 1.944, 1.46, 4.89, 0.57, [['// ', O], ['IMAGE', W]], { ...S.kicker, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  IMG(s, 13.335, 0, 7.72, 5.325, MID);
  IMG(s, 21.056, 0, 5.614, 5.325, LIGHT);
}

/** Slide 9 */
function slide09(s) {
  PAGE(s, 9);
  T(s, 1.563, 7.257, 6.09, 1.62, [[LOREM.slice(0, 123), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 1.499, 9.012, 6.56, 4.72, [['PORTO-FOLIO', O]], { ...S.head96w, lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 23.303, 2.999, 2.93, 0.57, [['// ', O], ['IMAGE', W]], { ...S.kicker, align: 'right', lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 270 });
  IMG(s, 1.75, 0, 9.972, 5.506, MID);
  IMG(s, 11.722, 2, 8.774, 7.917, LIGHT);
  IMG(s, 20.497, 4.833, 6.174, 5.083, MID);
  IMG(s, 20.497, 9.917, 4.413, 5.083, LIGHT);
  IMG(s, 16.083, 9.917, 4.413, 3.51, MID);
}

/** Slide 10 */
function slide10(s) {
  IMG(s, 4.45, 2.895, 7.748, 10.522, LIGHT);
  PAGE(s, 10);
  T(s, 15.78, 7.927, 9.57, 1.56, [[LOREM.slice(0, 156), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5 });
  T(s, 15.734, 6.273, 11.06, 1.72, [['PORTOFOLIO', O]], { ...S.head96w, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 15.78, 5.67, 4.89, 0.57, [['// ', O], ['IMAGE', W]], { ...S.kicker, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  IMG(s, 0, 10, 4.45, 5, MID);
  IMG(s, 3.228, 0, 4.915, 5.028, LIGHT);
  IMG(s, 0, 5.028, 5.781, 4.972, LIGHT);
  IMG(s, 5.671, 5.028, 8.774, 4.972, MID);
}

/** Slide 11 */
function slide11(s) {
  IMG(s, 22.442, 4.35, 4.228, 6.3, LIGHT);
  BOX(s, -2.468, 6.468, 7, 2.064, O, { rotate: 90 });
  PAGE(s, 11);
  T(s, 9.447, 8.016, 8.62, 2.12, [[LOREM.slice(0, 243), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 9.427, 4.321, 8.93, 3.74, [['The Rise of Urban Culture', O]], { ...S.head72, lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 16.911, 7.247, 16.45, 0.51, [['SCARS URBAN CULTURE ', W], ['//', O], [' SCARS URBAN CULTURE', W]], { ...S.strip, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 90 });
  IMG(s, 2.064, 1.972, 7.959, 11.056, LIGHT);
}

/** Slide 12 */
function slide12(s) {
  IMG(s, 14.08, 1.875, 10.886, 7.4, LIGHT);
  IMG(s, 16.081, 3.8, 10.585, 7.4, LIGHT);
  T(s, -0.682, 5.256, 16.759, 4.93, [['SCARS', BLACK]], { ...S.mega, align: 'justify', lineSpacingMultiple: 1, paraSpaceBefore: 20, outline: { color: MAPGREY, size: 1.5 }, transparency: 100 });
  PAGE(s, 12);
  T(s, 2.611, 9.655, 8.62, 1.06, [[LOREM.slice(0, 123), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 2.611, 4.321, 8.93, 3.74, [['The Age of Urban Culture', O]], { ...S.head72, lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 2.611, 3.839, 4.89, 0.57, [['// ', O], ['SCARS', W]], { ...S.kicker, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 23, 10.978, 3.75, 0.44, [['//', O], ['www.scars.shop', W]], { ...S.tag, lineSpacingMultiple: 1, rotate: 270 });
}

/** Slide 13 */
function slide13(s) {
  IMG(s, 0, 0, 7.775, 15, LIGHT);
  IMG(s, 2.449, 3.8, 10.886, 7.4, LIGHT);
  PAGE(s, 13);
  T(s, 16.461, 10.09, 7.41, 1.11, [[LOREM.slice(0, 105), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 16.461, 4.321, 8.01, 5.38, [['URBAN CONCEPT AND CULTURE', O]], { ...S.head72, lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 16.461, 3.839, 4.89, 0.57, [['//', O], ['13', W]], { ...S.kicker, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, -5.086, 4.807, 18.28, 4.93, [['SCARS', O]], { ...S.mega, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20, transparency: 15, rotate: 270 });
}

/** Slide 14 */
function slide14(s) {
  IMG(s, 13.335, 2.525, 13.335, 9.95, LIGHT);
  PAGE(s, 14);
  T(s, 3.361, 9.589, 7.41, 1.11, [[LOREM.slice(0, 105), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 3.361, 3.825, 8.01, 5.38, [['URBAN CONCEPT AND CULTURE', O]], { ...S.head72, lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 3.361, 3.337, 4.89, 0.57, [['//', O], ['SCARS', W]], { ...S.kicker, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  BOX(s, 15.139, 5.936, 6.3, 4.228, O, { rotate: 90 });
  IMG(s, 13.335, 0.75, 6.595, 9.95, LIGHT);
  T(s, 17.57, 3.337, 4.89, 0.57, [['13 ', W], ['// ', O], ['20', W]], { ...S.kicker, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20 });
}

/** Slide 15 */
function slide15(s) {
  IMG(s, 0, 2.525, 7.563, 9.95, LIGHT);
  PAGE(s, 15);
  T(s, -6.909, 7.247, 16.45, 0.51, [['SCARS URBAN CULTURE ', W], ['//', O], [' SCARS URBAN CULTURE', W]], { ...S.strip, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 90 });
  T(s, 10.965, 9.331, 9.57, 1.56, [[LOREM.slice(0, 156), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5 });
  T(s, 10.859, 4.182, 4.89, 0.57, [['// ', O], ['SCARS', W]], { ...S.kicker, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 7.563, 4.692, 19.897, 4.93, [['URBAN', O]], { ...S.mega, align: 'right', lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  IMG(s, 22.125, 4.094, 4.545, 6.813, LIGHT);
}

/** Slide 16 */
function slide16(s) {
  IMG(s, 17.012, 0, 9.658, 9.95, MID);
  IMG(s, 14.287, 2.525, 9.658, 9.95, LIGHT);
  PAGE(s, 16);
  T(s, 3.361, 10.089, 7.41, 1.11, [[LOREM.slice(0, 105), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 3.361, 3.837, 4.89, 0.57, [['//', O], ['SCARS', W]], { ...S.kicker, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 3.361, 4.325, 8.01, 5.38, [['URBAN CONCEPT AND CULTURE', O]], { ...S.head72, lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 14.224, 1.672, 3.75, 0.44, [['//', O], ['www.scars.shop', W]], { ...S.tag, align: 'right', lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 16.57, 9.946, 11.532, 1.72, [['RIGHT NOW', O]], { ...S.head96w, align: 'right', lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 270 });
  T(s, 17.864, 9.946, 11.532, 1.72, [['RIGHT NOW', BLACK]], { ...S.head96w, align: 'right', lineSpacingMultiple: 1, paraSpaceBefore: 20, outline: { color: O, size: 1.5 }, transparency: 100, rotate: 270 });
  T(s, 19.159, 9.946, 11.532, 1.72, [['RIGHT NOW', BLACK]], { ...S.head96w, align: 'right', lineSpacingMultiple: 1, paraSpaceBefore: 20, outline: { color: O, size: 1.5 }, transparency: 100, rotate: 270 });
}

/** Slide 17 */
function slide17(s) {
  PAGE(s, 17);
  T(s, 15.628, 2.952, 9.26, 4.14, [[LOREM.slice(0, 529), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 15.452, 9.364, 11.394, 4.93, [['50%', O]], { ...S.mega, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 3.151, 1.182, 7.74, 1.72, [['ARRIVAL', O]], { ...S.head96w, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, -0.985, 7.238, 6.56, 1.72, [['NEW', O]], { ...S.head96w, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 270 });
  T(s, 15.586, 9.139, 4.89, 0.57, [['//', O], ['DISC UP TO', W]], { ...S.kicker, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  IMG(s, 3.328, 3.125, 7.562, 9.95, LIGHT);
}

/** Slide 18 */
function slide18(s) {
  BOX(s, 2.578, 6.703, 5.844, 7.438, O, { rotate: 90 });
  PAGE(s, 18);
  T(s, 17.739, 11.193, 7.21, 1.06, [[LOREM.slice(0, 93), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 17.718, 7.241, 8.07, 3.56, [['The Part of Urban Culture', O]], { ...S.head72, lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  IMG(s, 2.906, 4.688, 11.25, 7.562, LIGHT);
  IMG(s, 17.938, 0, 8.733, 4.688, LIGHT);
  T(s, 10.984, 1.34, 4.89, 0.57, [['// ', O], ['SCARS', W]], { ...S.kicker, align: 'right', lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, -0.323, 1.85, 19.897, 4.93, [['URBAN', O]], { ...S.mega, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
}

/** Slide 19 */
function slide19(s) {
  BOX(s, 10.89, 0, 6.86, 15, O);
  IMG(s, 10.89, 1.875, 8.81, 11.25, LIGHT);
  PAGE(s, 19);
  T(s, 1.638, 8.106, 6.8, 1.62, [[LOREM.slice(0, 123), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 15.973, 3.057, 7.74, 1.72, [['URBAN', O]], { ...S.head96w, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 18.709, 9.432, 8.916, 1.72, [[' // ', O], ['FSHN', W]], { ...S.head96w, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 1.638, 5.231, 4.89, 0.57, [['// ', O], ['SCARS', W]], { ...S.kicker, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 1.622, 6.223, 9.268, 1.72, [['CULTURE', O]], { ...S.head96w, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 2.066, 7.243, 16.45, 0.51, [['SCARS URBAN CULTURE ', W], ['//', O], [' SCARS URBAN CULTURE', W]], { ...S.strip, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 90 });
}

/** Slide 20 */
function slide20(s) {
  BOX(s, 17.213, 0.833, 5.234, 8.101, O, { rotate: 90 });
  IMG(s, 9.3, 3.225, 13.575, 8.55, LIGHT);
  PAGE(s, 20);
  T(s, 1.629, 6.536, 7.02, 1.62, [[LOREM.slice(0, 123), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 1.629, 2.537, 8.93, 3.56, [['The Part of Urban Design', O]], { ...S.head72, lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 16.911, 7.243, 16.45, 0.51, [['SCARS URBAN CULTURE ', W], ['//', O], [' SCARS URBAN CULTURE', W]], { ...S.strip, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 90 });
  T(s, -0.348, 8.601, 19.897, 4.93, [['SCARS', O]], { ...S.mega, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
}

/** Slide 21 */
function slide21(s) {
  BOX(s, 13.4, 1.6, 9.85, 11.8, O, { rotate: 90 });
  IMG(s, 13.095, 3.225, 13.575, 8.55, LIGHT);
  PAGE(s, 21);
  T(s, 17.486, 6.555, 15.92, 1.72, [['TIME TO GET BREAK', O]], { ...S.head96w, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 270 });
  T(s, 2.714, 7.031, 6.8, 1.62, [[LOREM.slice(0, 123), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 2.714, 6.183, 6.062, 0.57, [['// ', O], ['TAKE YOUR TIME', W]], { ...S.kicker, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
}

/** Slide 22 */
function slide22(s) {
  PAGE(s, 22);
  T(s, -1.332, 4.692, 29.59, 4.93, [['GET BREAK', O]], { ...S.mega, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 1.131, 3.839, 6.062, 0.57, [['// ', O], ['TIME TO', W]], { ...S.kicker, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 19.15, 9.624, 6.33, 1.11, [[LOREM.slice(0, 85), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  BOX(s, 7.664, 2.977, 11.376, 9.046, PANEL, { rotate: 90 });
  IMG(s, 9.554, 2.525, 7.562, 9.95, LIGHT);
}

/** Slide 23 */
function slide23(s) {
  IMG(s, 3.337, 2.851, 19.997, 8.549, LIGHT);
  PAGE(s, 23);
  T(s, -0.69, 6.263, 28.11, 1.72, [['-IT\u2019S TIME TO GET BREAK SECTION-', O]], { ...S.head96w, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 5.853, 11.892, 14.96, 1.11, [[LOREM.slice(0, 123), GREY]], { ...S.body, align: 'center', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 9.853, 1.454, 6.87, 0.57, [['// ', O], ['TAKE YOUR TIME', W]], { ...S.kicker, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20 });
}

/** Slide 24 */
function slide24(s) {
  BOX(s, 11.367, 3.78, 11.376, 7.44, PANEL, { rotate: 90 });
  IMG(s, 14.036, 2.525, 7.563, 9.95, LIGHT);
  PAGE(s, 24);
  T(s, 3.318, 2.049, 6.062, 0.57, [['// ', O], ['TIME TO', W]], { ...S.kicker, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 2.962, 2.621, 20.333, 9.76, [['GET', O], BR, ['BREAK', O]], { ...S.mega, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 18.177, 6.945, 9.95, 1.11, [[LOREM.slice(0, 123), GREY]], { ...S.body, align: 'center', lineSpacingMultiple: 1.5, paraSpaceBefore: 20, rotate: 270 });
}

/** Slide 25 */
function slide25(s) {
  IMG(s, 0, 0, 18.139, 10.552, MID);
  IMG(s, 5.944, 2.632, 12.194, 7.92, LIGHT);
  BOX(s, 20.002, 8.689, 2.503, 6.23, O, { rotate: 90 });
  PAGE(s, 25);
  T(s, 21.974, 4.245, 6.7, 0.57, [['WE WORK WITH BIG', W]], { ...S.kicker, lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 90 });
  T(s, 20.794, 7.454, 6.7, 0.57, [['//', O], [' AND SMALL ALIKE', W]], { ...S.kicker, lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 90 });
  BOX(s, 11.781, 0.451, 7.92, 15.795, INK, { rotate: 90 });
  T(s, 9.369, 10.249, 2.795, 0.295, [['SUTERRA', W]], { ...S.brand, fontSize: 22, align: 'center', valign: 'middle' });
  T(s, 14.977, 10.123, 1.882, 0.547, [['PACT', W]], { ...S.brand, fontSize: 30, align: 'center', valign: 'middle' });
  T(s, 14.904, 6.136, 2.029, 0.989, [['DOUBLE-O', W]], { ...S.brand, fontSize: 22, align: 'center', valign: 'middle' });
  T(s, 19.631, 6.515, 2.754, 0.625, [['AM MEDIA', W]], { ...S.brand, fontSize: 24, align: 'center', valign: 'middle' });
  T(s, 19.631, 10.277, 2.752, 0.393, [['BESTELLER', W]], { ...S.brand, fontSize: 22, align: 'center', valign: 'middle' });
  T(s, 9.745, 6.436, 2.041, 0.782, [['Glitchera', W]], { ...S.brand, fontSize: 26, align: 'center', valign: 'middle', italic: true, fontFace: 'Rubik' });
  T(s, -2.156, 4.14, 10, 1.72, [['OUR BRAND', O]], { ...S.head96w, lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 270 });
}

/** Slide 26 */
function slide26(s) {
  IMG(s, 9.535, 7.5, 5.865, 3.906, LIGHT);
  PAGE(s, 26);
  T(s, 2.042, 2.654, 10, 3.33, [['WE WORK WITH', O]], { ...S.head96w, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 2.013, 1.797, 4.89, 0.57, [['MEET THE TEAM', W]], { ...S.kicker, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 15.288, 4.04, 4.89, 0.51, [['EMER CARLOS', W]], { ...S.strip, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 9.479, 6.909, 4.89, 0.51, [['ALISTON MORESSY', W]], { ...S.strip, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 3.997, 10.776, 4.89, 0.51, [['KLINT ELEVENTH', W]], { ...S.strip, align: 'justify', lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 270 });
  T(s, 19.811, 7.19, 4.89, 0.51, [['MATHILDA REY', W]], { ...S.strip, align: 'right', lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 270 });
  T(s, 13.51, 10.627, 11.244, 1.72, [['THE EXPERT', O]], { ...S.head96w, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  IMG(s, 15.399, 0, 5.865, 3.906, LIGHT);
  IMG(s, 22.764, 5.194, 3.906, 5.392, LIGHT);
  IMG(s, 2.047, 7.972, 3.906, 5.392, LIGHT);
}

/** Slide 27 */
function slide27(s) {
  IMG(s, 19.25, 0, 7.42, 11.056, MID);
  IMG(s, 17.139, 1.972, 7.42, 11.056, LIGHT);
  PAGE(s, 27);
  T(s, 1.886, 12.279, 10.14, 0.57, [['// ', O], ['CREATYVE DIVISION', W]], { ...S.kicker, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 1.886, 5.713, 10.14, 5.66, [[LOREM, GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5 });
  T(s, 1.886, 1.652, 10.14, 3.33, [['Klint', W], BR, ['Eleventh ', W]], { ...S.head96s, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 15.043, 12.065, 11.957, 1.72, [['SCARS', O]], { fontFace: 'Russo One', fontSize: 96, charSpacing: 32, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 19.557, 4.972, 10, 1.72, [['PROFIL', O]], { fontFace: 'Russo One', fontSize: 96, charSpacing: 32, align: 'right', lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 270 });
}

/** Slide 28 */
function slide28(s) {
  T(s, -6.675, 4.682, 18.28, 4.93, [['SCARS', O]], { ...S.mega, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 270 });
  PAGE(s, 28);
  T(s, 13.482, 10.585, 8.69, 1.74, [[LOREM.slice(0, 57) + PRAESENT, GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 13.335, 2.913, 5, 0.57, [['INFOGRAPHIC', W]], { ...S.kicker, align: 'justify', lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 20.6, 7.216, 10.14, 0.57, [['// ', O], ['SCARS MEDIA', W]], { ...S.kicker, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 270 });
  T(s, 13.482, 9.368, 5.95, 1.09, [['3.000.189', O]], { ...S.stat72, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 18.287, 9.964, 2.29, 0.54, [['Customer', W]], { ...S.semi20, lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  CHART_28(s);
  IMG(s, 2.556, 3.508, 5.968, 7.952, LIGHT);
  DEVICE(s, 2.066, 2.625, 6.954, 9.749);
}

/** Slide 29 */
function slide29(s) {
  IMG(s, 18.194, 0, 8.476, 15, MID);
  IMG(s, 15, 2.653, 4.306, 9.181, LIGHT);
  PAGE(s, 29);
  T(s, 2.401, 5.065, 10, 1.72, [['SNAP-STORY', O]], { ...S.head96w, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 2.372, 4.214, 4.89, 0.57, [['INSTA', W]], { ...S.kicker, align: 'justify', lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 2.401, 7.064, 10.14, 3.64, [[LOREM.slice(0, 529), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  DEVICE(s, 14.578, 2.104, 5.171, 10.291);
}

/** Slide 30 */
function slide30(s) {
  IMG(s, 5.667, 0, 5.792, 15, LIGHT);
  BOX(s, 15.053, 10.497, 2.807, 0.51, W, { shape: 'roundRect', rectRadius: 0.255 });
  BOX(s, -2.038, 7.295, 15, 0.409, O, { rotate: 90 });
  PAGE(s, 30);
  T(s, -3.528, 7.214, 10.14, 0.57, [['// ', O], ['SCARS WATCH', W]], { ...S.kicker, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 270 });
  T(s, 14.925, 8.248, 8.19, 1.62, [[LOREM.slice(0, 154) + '.', GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 14.878, 3.304, 9.09, 4.95, [['The Next Big Privacy', O]], { ...S.head96s, lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 14.64, 10.518, 3.59, 0.51, [['Shop Now', BG]], { ...S.semi20, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20, bold: true });
  BOX(s, 3.274, 5.451, 5.083, 4.097, O, { rotate: 90 });
  DEVICE(s, 2.659, 1.644, 6.312, 11.713);
  T(s, 4.576, 7.243, 2.479, 0.492, [['SCARS', W]], { ...S.brand, fontSize: 20, align: 'center', valign: 'middle', rotate: 270 });
}

/** Slide 31 */
function slide31(s) {
  IMG(s, 12.722, 4.55, 9.361, 5.297, LIGHT);
  PAGE(s, 31);
  T(s, 1.701, 8.477, 8.78, 1.11, [[LOREM.slice(0, 123), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 1.654, 3.525, 9.09, 4.72, [['Just Shop in Our Web!', O]], { ...S.head96s, lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 1.654, 1.163, 3.75, 0.44, [['//', O], ['www.scars.shop', W]], { ...S.tag, lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  BOX(s, 24.475, 0, 0.201, 15, O);
  BOX(s, 1.829, 10.366, 2.807, 0.51, W, { shape: 'roundRect', rectRadius: 0.255 });
  T(s, 1.416, 10.387, 3.59, 0.51, [['Shop Now', BG]], { ...S.semi20, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20, bold: true });
  T(s, 17.624, 6.64, 14, 1.72, [['ONLINE - STORE', W]], { ...S.head96w, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 270 });
  DEVICE(s, 11.838, 4.332, 11.121, 6.544);
}

/** Slide 32 */
function slide32(s) {
  IMG(s, 15.15, 0, 11.52, 5.646, MID);
  IMG(s, 13.335, 1.854, 10.495, 5.646, LIGHT);
  PAGE(s, 32);
  T(s, 13.24, 9.604, 10.59, 2.12, [[LOREM.slice(0, 318), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 1.654, 0.923, 11.2, 4.95, [['Yearly Scars Progress', O]], { ...S.head96s, lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 18.17, 7.215, 15, 0.57, [['INFOGRAPHIC ', W], ['// ', O], ['INFOGRAPHIC ', W], ['// ', O], ['INFOGRAPHIC', W]], { ...S.kicker, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 270 });
  T(s, 13.24, 8.923, 5, 0.57, [['Selling Progress', W]], { ...S.sub28, align: 'justify', lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  CHART_32(s);
}

/** Slide 33 */
function slide33(s) {
  T(s, -6.675, 4.682, 18.28, 4.93, [['SCARS', O]], { ...S.mega, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 270 });
  MAP(s, 3.271, 4.618, 10.585, 6.501);
  BOX(s, 3.969, 4.536, 1.357, 1.357, O, { shape: 'ellipse', transparency: 15 });
  BOX(s, 11.115, 4.882, 2.964, 2.964, O, { shape: 'ellipse', transparency: 15 });
  BOX(s, 7.381, 8.46, 2.74, 2.74, O, { shape: 'ellipse', transparency: 15 });
  PAGE(s, 33);
  T(s, 2.542, 2.309, 3.75, 0.44, [['//', O], ['www.scars.shop', W]], { ...S.tag, lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 270 });
  T(s, 11.54, 0.824, 13.58, 4.95, [['We\u2019re Dominating US Market', O]], { ...S.head96s, align: 'right', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 4.403, 11.657, 8.9, 1.74, [[LOREM.slice(0, 57) + PRAESENT, GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 3.526, 4.926, 2.32, 0.63, [['25%', W]], { ...S.stat28, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20, bold: true });
  T(s, 11.513, 5.775, 2.32, 1.17, [['45%', W]], { ...S.stat54, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20, bold: true });
  T(s, 7.669, 9.347, 2.32, 1.17, [['50%', W]], { ...S.stat54, align: 'center', lineSpacingMultiple: 1, paraSpaceBefore: 20, bold: true });
  CHART_33(s);
  T(s, -3.914, 2.312, 3.746, 0.438, [['//', O], ['www.scars.shop', W]], { ...S.tag, rotate: 270 });
}

/** Slide 34 */
function slide34(s) {
  IMG(s, 0, 0, 6.466, 15, MID);
  PAGE(s, 34);
  T(s, 16.883, 8.451, 8.62, 2.12, [[LOREM.slice(0, 243), GREY]], { ...S.body, align: 'justify', lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 16.863, 4.289, 8.93, 3.74, [['Worldwide Shipping Service', O]], { ...S.head72, lineSpacingMultiple: 1.5, paraSpaceBefore: 20 });
  T(s, 21.826, 0.859, 3.75, 0.44, [['//', O], ['www.scars.shop', W]], { ...S.tag, align: 'right', lineSpacingMultiple: 1, paraSpaceBefore: 20 });
  T(s, 0.896, 0.113, 13.05, 14.64, [['WORLD MARKET', O], BR, ['WORLD MARKET', DIM], BR, ['WORLD MARKET', O], BR, ['WORLD MARKET', DIM], BR, ['WORLD MARKET', O], BR, ['WORLD MARKET', DIM]], { ...S.head96s, lineSpacingMultiple: 1.5 });
  BOX(s, 2.846, 3.075, 12.625, 8.85, BG);
  MAP(s, 3.689, 4.224, 12.207, 6.651);
  PIN(s, 5.141, 6.283);
  PIN(s, 7.441, 8.445);
  PIN(s, 6.916, 8.807);
  PIN(s, 9.845, 9.252);
  PIN(s, 10.11, 7);
  PIN(s, 11.103, 7.255);
  PIN(s, 12.606, 7.889);
  PIN(s, 12.765, 8.354);
  PIN(s, 14.116, 9.259);
  PIN(s, 13.142, 6.903);
  PIN(s, 13.805, 6.693);
  PIN(s, 13.976, 6.045);
  BOX(s, 2.846, 2.875, 12.625, 0.207, O);
}

/** Slide 35 */
function slide35(s) {
  IMG(s, 10.405, 3.985, 9.509, 5.646, MID);
  PAGE(s, 35);
  T(s, 18.376, 8.155, 5.507, 1.72, [['THANK', O]], { ...S.head96s, lineSpacingMultiple: 1, paraSpaceBefore: 20, rotate: 270 });
  T(s, 17.998, 2.271, 6.261, 1.72, [['YOU', BLACK]], { ...S.head96s, lineSpacingMultiple: 1, paraSpaceBefore: 20, outline: { color: O, size: 1.5 }, transparency: 100, rotate: 270 });
  T(s, 5.904, 5.912, 3.31, 0.5, [['Scars Media', GREY]], { ...S.body, lineSpacingMultiple: 1 });
  T(s, 5.904, 7.296, 3.31, 0.5, [['Scars_US', GREY]], { ...S.body, lineSpacingMultiple: 1 });
  T(s, 5.904, 8.636, 3.31, 0.5, [['Scars.Seattle', GREY]], { ...S.body, lineSpacingMultiple: 1 });
  BOX(s, 5.102, 5.787, 0.657, 0.657, O, { shape: 'ellipse' });
  ICON_TW(s, 5.277, 6.002, 0.306);
  BOX(s, 5.102, 7.172, 0.657, 0.657, O, { shape: 'ellipse' });
  ICON_IG(s, 5.28, 7.35, 0.301);
  BOX(s, 5.102, 8.556, 0.657, 0.657, O, { shape: 'ellipse' });
  ICON_FB(s, 5.361, 8.755, 0.14);
  IMG(s, 8.821, 5.369, 9.509, 5.646, LIGHT);
}

// ---------------------------------------------------------------- build
const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30, slide31, slide32, slide33, slide34, slide35];

SLIDES.forEach(build => {
  const s = pptx.addSlide();
  s.background = { color: BG };
  build(s);
});

pptx.writeFile({ fileName: path.join(__dirname, '0d907aef-405d-4d5e-bba2-a696d4f2a6f1_grok_final.pptx') })
  .then(f => console.log('wrote ' + f));
