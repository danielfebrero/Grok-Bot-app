/**
 * "Hometown Real Estate" — 15-slide deck rebuilt with pptxgenjs.
 *
 * Slide size: 13.3333 x 7.5 in (16:9). Palette is navy + gold on white, with
 * Libre Baskerville headlines and Ovo body copy.
 *
 * Every picture placeholder in the source .pptx is empty (the file ships no
 * media parts at all), so `imageSlot()` re-creates each one as an unfilled
 * frame of the same geometry and preset shape rather than a photo.
 */
const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ palette */
const NAVY = '19222D';  // panels, dark canvases, first chart series
const GOLD = '9B784C';  // accent circles, columns, cards, pills
const BLUSH = 'CBBAB1'; // third chart series
const WHITE = 'FFFFFF'; // text on navy/gold
const BLACK = '000000'; // headlines that inherit the default text colour
const INK = '404040';   // body copy on white (theme tx1, lighter 25%)
const GRAY = '595959';  // chart title / axis / legend labels
const GRID = 'D9D9D9';  // chart gridlines and category axis line

/* -------------------------------------------------------------------- type */
const SERIF = 'Libre Baskerville';
const BODY = 'Ovo';
const CHART_FONT = 'Calibri'; // theme minor font, used by the chart

const H1 = 48; // cover headline
const H2 = 40; // section headline
const H3 = 20; // card headline
const TXT = 12; // body copy

/* -------------------------------------------------- copy (all lorem ipsum) */
const S = [
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent imperdiet quis eros sed pellentesque. Quisque sed pretium purus, nec luctus lorem. ',
  'Morbi molestie tincidunt hendrerit. Maecenas nisi massa, tempor in convallis eu, accumsan sit amet purus',
  '. Aliquam erat volutpat. Donec venenatis fermentum lacus id congue. ',
  'Etiam porttitor massa nec turpis malesuada, molestie venenatis turpis',
  ' blandit. Donec accumsan sollicitudin ante, non sagittis sapien dignissim et. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia curae; Duis at neque in felis semper tincidunt. ',
  'Suspendisse bibendum tempor gravida.',
];
const LOREM_1 = S[0];                                     // one sentence
const LOREM_2 = S[0] + S[1];                              // card copy
const LOREM_3 = S[0] + S[1] + S[2];                       // short paragraph
const LOREM_4 = S[0] + S[1] + S[2] + S[3] + '.';          // medium paragraph
const LOREM_5 = S[0] + S[1] + S[2] + S[3] + S[4];         // long paragraph
const LOREM_6 = LOREM_5 + S[5];                           // longest paragraph

/* ----------------------------------------------------------------- helpers */
// No shape in this deck has an outline, so fill colour is the only knob.
function rect(slide, x, y, w, h, color) {
  slide.addShape('rect', { x, y, w, h, fill: { color } });
}
function oval(slide, x, y, w, h, color) {
  slide.addShape('ellipse', { x, y, w, h, fill: { color } });
}
// `radius` in inches; omit it for PowerPoint's default corner rounding.
function roundRect(slide, x, y, w, h, color, radius) {
  slide.addShape('roundRect', { x, y, w, h, fill: { color }, rectRadius: radius });
}

// Every text box in the source is a top-anchored, auto-height (spAutoFit) box.
// Headlines sit on a single line and never wrap.
function headline(slide, x, y, w, h, text, color, size) {
  slide.addText(text, {
    x, y, w, h, fontFace: SERIF, fontSize: size || H2, bold: true, color,
    valign: 'top', wrap: false, fit: 'resize',
  });
}
function para(slide, x, y, w, h, text, color, align) {
  slide.addText(text, {
    x, y, w, h, fontFace: BODY, fontSize: TXT, color,
    valign: 'top', align: align || 'left', fit: 'resize',
  });
}

// Fully rounded 0.5in button: a capsule plus a centred caption laid over it.
function pill(slide, o) {
  roundRect(slide, o.x, o.y, o.w, 0.5, o.fill, 0.25);
  para(slide, o.labelX, o.labelY, o.labelW, 0.3029, o.label, WHITE, 'center');
}

// Picture placeholder from the source deck: correct frame, no artwork.
function imageSlot(slide, x, y, w, h, shape) {
  slide.addShape(shape || 'roundRect', {
    x, y, w, h, fill: { color: WHITE, transparency: 100 },
    objectName: '[image] placeholder',
  });
}

// Slides 1 and 15 share one cover template: a two-line 48pt headline, a
// one-sentence blurb and a pill, all measured from a single origin (x, y).
function coverText(slide, o) {
  headline(slide, o.x, o.y, o.w1, 0.9088, o.line1, WHITE, H1);
  headline(slide, o.x, o.y + 0.6592, o.w2, 0.9088, o.line2, WHITE, H1);
  para(slide, o.x, o.y + 1.5229, 4.2985, 0.7068, LOREM_1, WHITE);
  pill(slide, {
    x: o.x + 0.1022, y: o.y + 2.3843, w: o.pillW, fill: GOLD, label: o.label,
    labelX: o.x + 0.2272, labelY: o.y + 2.4828, labelW: o.labelW,
  });
}

/* ------------------------------------------------------------------ slides */

// 1 — Cover: navy canvas with three gold circles bleeding off the edges.
function slide01(slide) {
  slide.background = { color: NAVY };
  oval(slide, 6.1241, -3.1361, 9.0986, 9.0986, GOLD);  // large, top-right
  coverText(slide, {
    x: 0.842, y: 2.568, line1: 'Hometown', w1: 4.2112, line2: 'Real Estate', w2: 4.071,
    pillW: 1.8125, label: 'Start Presentation', labelW: 1.5625,
  });
  oval(slide, 2.2552, 5.4112, 4.9698, 4.9698, GOLD);   // bottom-centre
  oval(slide, -0.7556, -1.5849, 3.1953, 3.1953, GOLD); // small, top-left
  imageSlot(slide, 6.5822, 0.5615, 6.467, 6.3771, 'flowChartConnector');
  imageSlot(slide, 5.4694, 2.3359, 2.8281, 2.8281, 'ellipse');
}

// 2 — Introduction: text column left, gold band right, navy footer.
function slide02(slide) {
  rect(slide, 7.9792, 0, 5.3542, 7.5, GOLD);
  headline(slide, 1.2185, 1.1151, 4.0114, 0.7742, 'Introduction', NAVY);
  para(slide, 1.2185, 1.9168, 4.7919, 2.3225, LOREM_6, INK);
  pill(slide, {
    x: 1.301, y: 4.4669, w: 2.1483, fill: NAVY, label: 'www.yourwebsite.com',
    labelX: 1.4052, labelY: 4.5655, labelW: 1.94,
  });
  rect(slide, 0, 5.373, 7.9792, 2.127, NAVY);
  imageSlot(slide, 6.6667, 0.709, 5.9569, 6.0819, 'flowChartDelay');
}

// 3 — Importance of Hometown: navy bar across the top, navy copy panel below.
function slide03(slide) {
  rect(slide, 0, 4.4551, 7.2222, 2.4656, NAVY);
  headline(slide, 0.9628, 2.4812, 4.4584, 0.7742, 'Importance of', NAVY);
  headline(slide, 0.9628, 3.1399, 3.5415, 0.7742, 'Hometown', NAVY);
  para(slide, 0.6349, 4.8783, 6.419, 1.7166, LOREM_6, WHITE);
  rect(slide, 0, -0.0139, 13.3333, 1.6983, NAVY);
  imageSlot(slide, 7.6975, 2.0101, 5.0637, 4.9105, 'round2SameRect');
}

// 4 — Elements of Hometown: three numbered navy cards down the right, each a
// rounded card + gold circle + numeral + heading + copy, stepped by 2.1921in.
function slide04(slide) {
  headline(slide, 0.7262, 1.6106, 3.7572, 0.7742, 'Elements of', NAVY);
  headline(slide, 0.7262, 2.2693, 3.5415, 0.7742, 'Hometown', NAVY);
  ['1', '2', '3'].forEach((numeral, i) => {
    const dy = i * 2.1921;
    roundRect(slide, 7.4127, 0.5518 + dy, 5.1944, 2.0767, NAVY);
    oval(slide, 6.7937, 0.9711 + dy, 1.2381, 1.2381, GOLD);
    slide.addText(numeral, {
      x: 7.0214, y: 1.2031 + dy, w: 0.7826, h: 0.7742,
      fontFace: SERIF, fontSize: H2, bold: true, color: WHITE,
      align: 'center', valign: 'top', fit: 'resize',
    });
    headline(slide, 8.1953, 0.8514 + dy, 1.7008, 0.4376, 'Your Text', WHITE, H3);
    para(slide, 8.1953, 1.2031 + dy, 4.1349, 1.1107, LOREM_2, WHITE);
  });
  imageSlot(slide, 0.8631, 3.481, 5.44, 3.1957);
}

// 5 — Interior Design Style: two gold circles, headline block, navy quote panel.
function slide05(slide) {
  oval(slide, 9.9186, 4.2571, 4.9524, 4.9524, GOLD);  // bottom-right
  oval(slide, 4.1587, -2.4494, 4.9524, 4.9524, GOLD); // top-centre
  headline(slide, 1.3133, 1.928, 2.523, 0.7742, 'Interior', NAVY);
  headline(slide, 1.3133, 2.5868, 3.8851, 0.7742, 'Design Style', NAVY);
  para(slide, 1.3133, 3.3609, 4.7919, 1.3127, LOREM_3, INK);
  rect(slide, 0, 4.9943, 5.7178, 1.8917, NAVY);
  para(slide, 0.3211, 5.1842, 5.171, 1.3127, LOREM_4, WHITE);
  imageSlot(slide, 7.0694, 0.7667, 5.6174, 2.9833, 'round2SameRect');
  imageSlot(slide, 7.0694, 3.9175, 5.6174, 2.9833);
}

// 6 — Building For The Future: headline left, copy right, navy band below.
function slide06(slide) {
  rect(slide, 0, 5.1111, 13.3333, 2.3889, NAVY);
  headline(slide, 1.4762, 1.4519, 3.9763, 0.7742, 'Building For', NAVY);
  headline(slide, 1.4762, 2.1106, 3.545, 0.7742, 'The Future', NAVY);
  para(slide, 5.8571, 1.3673, 6.873, 1.5146, LOREM_5, INK);
  imageSlot(slide, 1.4762, 3.5464, 10.7056, 3.1957);
}

// 7 — Room Inspiration: photo left, headline right, navy panel bottom-right.
function slide07(slide) {
  imageSlot(slide, 0.7837, 0.7667, 6.2968, 6.023);
  headline(slide, 7.7577, 1.9916, 2.0199, 0.7742, 'Room', NAVY);
  headline(slide, 7.7577, 2.6503, 3.503, 0.7742, 'Inspiration', NAVY);
  para(slide, 7.7577, 3.493, 4.7919, 1.3127, LOREM_3, INK);
  rect(slide, 6.1111, 5.0503, 7.2222, 2.4656, NAVY);
  para(slide, 6.746, 5.4736, 6.419, 1.7166, LOREM_6, WHITE);
}

// 8 — Room Inspiration: headline left, navy panel top-right, copy below it.
function slide08(slide) {
  headline(slide, 1.1734, 1.2449, 2.0199, 0.7742, 'Room', NAVY);
  headline(slide, 1.1734, 1.9036, 3.503, 0.7742, 'Inspiration', NAVY);
  para(slide, 7.1875, 4.2899, 5.1746, 2.1205, LOREM_6, INK);
  rect(slide, 7.1875, -0.0139, 6.1458, 2.9306, NAVY);
  para(slide, 7.6163, 0.4921, 5.2882, 1.9186, LOREM_5, WHITE);
  imageSlot(slide, 1.2359, 2.9188, 5.4308, 3.804, 'round2SameRect');
}

// 9 — Room Inspiration: navy circle top-left, gold column right.
function slide09(slide) {
  imageSlot(slide, 0.7816, 0.8212, 5.5328, 6.0955);
  rect(slide, 10.8588, 0, 2.4746, 7.5, GOLD);
  headline(slide, 6.8156, 2.2731, 2.0199, 0.7742, 'Room', NAVY);
  headline(slide, 6.8156, 2.9318, 3.503, 0.7742, 'Inspiration', NAVY);
  para(slide, 6.8156, 3.75, 3.8399, 1.5146, LOREM_3, INK);
  oval(slide, -1.6395, -1.456, 4.529, 4.529, NAVY);
}

// 10 — Hometown Project: two gold project cards with centred text.
function slide10(slide) {
  rect(slide, 0, 5.8508, 13.3333, 1.6492, NAVY);
  headline(slide, 0.7497, 1.3345, 3.5415, 0.7742, 'Hometown', NAVY);
  headline(slide, 0.7497, 1.9932, 2.2951, 0.7742, 'Project', NAVY);
  [
    { x: 5.1585, label: 'Project 01', labelX: 6.214, labelW: 1.6623, textX: 5.5515 },
    { x: 9.0712, label: 'Project 02', labelX: 10.0969, labelW: 1.7219, textX: 9.4642 },
  ].forEach(card => {
    roundRect(slide, card.x, 0.4694, 3.75, 3.0, GOLD);
    slide.addText(card.label, {
      x: card.labelX, y: 0.9474, w: card.labelW, h: 0.4376,
      fontFace: SERIF, fontSize: H3, bold: true, color: WHITE,
      align: 'center', valign: 'top', wrap: false, fit: 'resize',
    });
    para(slide, card.textX, 1.3686, 2.9873, 1.7166, LOREM_2, WHITE, 'center');
  });
  imageSlot(slide, 1.4762, 4.0202, 10.7056, 2.7219);
}

// 11 — Break slide: full-height gold column right, 60pt headline bottom-left.
function slide11(slide) {
  rect(slide, 7.3673, 0, 5.966, 7.5, GOLD);
  headline(slide, 1.1525, 5.1807, 5.2876, 1.1107, 'Break Slide', BLACK, 60);
  para(slide, 1.1525, 6.2637, 5.7433, 0.5049, LOREM_1, INK);
  imageSlot(slide, 1.1517, 0.713, 11.5567, 4.0161, 'round2DiagRect');
}

// 12 — Our Chart: 100% stacked column chart left, headline and copy right.
function slide12(slide, pptx) {
  headline(slide, 8.7926, 2.6135, 3.233, 0.7742, 'Our Chart', BLACK);
  para(slide, 8.7926, 3.4918, 3.8399, 1.5146, LOREM_3, INK);

  const cats = ['Category 1', 'Category 2', 'Category 3', 'Category 4'];
  const series = [
    { name: 'Series 1', labels: cats, values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels: cats, values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels: cats, values: [2, 2, 3, 5] },
  ];
  slide.addChart(pptx.ChartType.bar, series, {
    x: 0.8651, y: 0.7551, w: 6.9717, h: 6.1415,
    barDir: 'col', barGrouping: 'percentStacked', barGapWidthPct: 150,
    chartColors: [NAVY, GOLD, BLUSH],
    chartArea: { roundedCorners: false },
    showTitle: true, title: 'Chart Title',
    titleFontFace: CHART_FONT, titleFontSize: 18.6, titleColor: GRAY,
    showLegend: true, legendPos: 'b',
    legendFontFace: CHART_FONT, legendFontSize: 12, legendColor: GRAY,
    catAxisLabelFontFace: CHART_FONT, catAxisLabelColor: GRAY, catAxisLabelFontSize: 12,
    catAxisLabelPos: 'nextTo', catAxisLineColor: GRID, catAxisLineSize: 0.75,
    catAxisMajorTickMark: 'none', catAxisMinorTickMark: 'none',
    valAxisLabelFontFace: CHART_FONT, valAxisLabelColor: GRAY, valAxisLabelFontSize: 12,
    valAxisLabelFormatCode: '0%', valAxisLineShow: false,
    valAxisMajorTickMark: 'none', valAxisMinorTickMark: 'none',
    valGridLine: { color: GRID, size: 0.75 },
    showValue: false,
  });
}

// 13 — Make It Comfortable: full-height navy column on the left.
function slide13(slide) {
  rect(slide, -0.0513, 0, 5.9493, 7.5, NAVY);
  headline(slide, 0.8089, 1.9905, 2.4687, 0.7742, 'Make It', WHITE);
  headline(slide, 0.8089, 2.6492, 3.9938, 0.7742, 'Comfortable', WHITE);
  para(slide, 0.8089, 3.5473, 4.6399, 2.3225, LOREM_6, WHITE);
  imageSlot(slide, 6.9375, 1.4583, 2.4913, 4.4115, 'rect');
  imageSlot(slide, 6.6667, 0.7677, 5.7486, 5.9899);
}

// 14 — Conclusion: headline and copy left, wide navy summary panel below.
function slide14(slide) {
  imageSlot(slide, 7.2292, 0.9356, 5.4308, 5.8352);
  headline(slide, 1.5583, 1.8403, 3.5976, 0.7742, 'Conclusion', BLACK);
  para(slide, 1.5583, 2.7187, 4.3792, 1.3127, LOREM_3, INK);
  rect(slide, 0, 4.4269, 9.6667, 1.9699, NAVY);
  para(slide, 0.6419, 4.7866, 8.5854, 1.3127, LOREM_6, WHITE);
}

// 15 — Closing: mirror of the cover, gold circles on the left.
function slide15(slide) {
  slide.background = { color: NAVY };
  oval(slide, -2.35, -3.1224, 9.4262, 9.4262, GOLD);  // large, top-left
  coverText(slide, {
    x: 7.6222, y: 2.622, line1: 'Thanks for', w1: 4.1306, line2: 'Your Attention', w2: 5.526,
    pillW: 2.2103, label: 'www.yourwebsite.com', labelW: 2.0124,
  });
  oval(slide, 10.8411, 5.2563, 3.1953, 3.1953, GOLD); // bottom-right
  imageSlot(slide, 0.3802, 0.7497, 6.2864, 6.2712, 'ellipse');
  imageSlot(slide, 4.5547, 2.4715, 2.8281, 2.8281, 'ellipse');
}

/* -------------------------------------------------------------------- main */
const SLIDES = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
  slide09, slide10, slide11, slide12, slide13, slide14, slide15,
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE_16x9', width: 13.3333333, height: 7.5 });
  pptx.layout = 'WIDE_16x9';
  pptx.title = 'Hometown Real Estate';

  SLIDES.forEach(buildSlide => buildSlide(pptx.addSlide(), pptx));

  return pptx.writeFile({ fileName: path.join(__dirname, '02fa68f9-995b-4fca-b277-492a2732c35f_grok_final.pptx') });
}

build().then(file => console.log('wrote', file));
