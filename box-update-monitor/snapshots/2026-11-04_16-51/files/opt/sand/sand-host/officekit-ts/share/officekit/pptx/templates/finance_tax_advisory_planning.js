/**
 * "Tax Advisory" deck — 20 slides, 13.333 x 7.5 in (16:9).
 * Rebuilt with pptxgenjs. Photographs in the original are replaced by
 * light-grey placeholder rectangles labelled "[image]".
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette ---
const C = {
  dark: '534447',        // theme accent1 — deep plum/brown
  sand: 'D9C7AF',        // theme accent2 — sand
  brown: '835643',       // theme accent3
  lime: 'C4D473',        // theme accent4
  ochre: '876A4A',       // theme accent6
  mauve: '9D878C',       // secondary rose-grey
  ink: '292223',         // body headline near-black
  black: '000000',
  white: 'FFFFFF',
  gold: '84663F',        // eyebrow / accent text
  cocoa: '624032',       // header text on light slides
  pearl: 'DED6D9',       // header + body text on dark slides
  charcoal: '3E3335',    // dark slide background
  orange: 'E88F4A',
  rose: 'BEAFB2',
  tan: 'BFA589',
  clay: 'BF9482',
  cream: 'E7DCCE',       // table zebra row
  grid: 'F2F2F2',        // chart grid line
  hair: 'D8D8D8',        // thin connector line
  grey59: '595959',      // chart axis labels
  imgFill: 'F1F1F1',     // photo placeholder fill
  imgText: 'A9A2A2',
};

const SERIF = 'Bodoni Moda';   // display / headline face
const SANS = 'League Spartan'; // body face
const TABLE_FONT = 'Open Sans Light';

// ------------------------------------------------------------ helper libs ---
// Text-box insets used throughout the source deck: 7.2pt sides, 3.6pt top/bottom.
// pptxgenjs reads the array as [left, right, bottom, top] in points.
const INSETS = [7.2, 7.2, 3.6, 3.6];

/** Text box mirroring the source: top-anchored, no bullet, tight insets. */
function txt(slide, text, o) {
  slide.addText(text, Object.assign({
    fontFace: SERIF, valign: 'top', margin: INSETS,
    lineSpacingMultiple: 0.9, isTextBox: true,
  }, o));
}

/** Body copy: sans face, 1.3 line spacing. */
function body(slide, text, o) {
  txt(slide, text, Object.assign({
    fontFace: SANS, fontSize: 12, color: C.black, lineSpacingMultiple: 1.3,
  }, o));
}

/** Solid rectangle. */
function rect(slide, o) {
  slide.addShape('rect', o);
}

/** Grey placeholder standing in for a photograph in the reference deck. */
function photo(slide, x, y, w, h, opts) {
  const o = opts || {};
  slide.addShape('rect', { x, y, w, h, fill: { color: o.fill || C.imgFill }, line: o.line });
  slide.addText('[image]', {
    x, y, w, h, align: 'center', valign: 'middle',
    fontFace: SANS, fontSize: 11, color: o.textColor || C.imgText,
  });
}

/** The two-bar logo mark that sits in a slide corner. */
function logo(slide, x, y) {
  slide.addShape('snip2DiagRect', {
    x: x, y: y + 0.149, w: 0.44, h: 0.087, rotate: 180,
    fill: { color: C.dark }, line: { type: 'none' },
  });
  slide.addShape('snip2DiagRect', {
    x: x + 0.117, y: y, w: 0.44, h: 0.087, rotate: 180,
    fill: { color: C.mauve }, line: { type: 'none' },
  });
}

/** "TAX ADVISORY / Maximizing Compliance & Savings" corner lockup. */
function header(slide, color) {
  txt(slide, 'TAX ADVISORY', {
    x: 11.508, y: 0.172, w: 1.648, h: 0.286,
    fontSize: 10.5, bold: true, color, align: 'right', charSpacing: 1,
  });
  txt(slide, 'Maximizing Compliance & Savings', {
    x: 11.017, y: 0.399, w: 2.141, h: 0.236,
    fontSize: 8, italic: true, color, align: 'right',
  });
}

/** Page number, bottom-right, as the master defines it. */
function pageNum(slide, n, color) {
  txt(slide, String(n), {
    x: 12.726, y: 7.082, w: 0.569, h: 0.303,
    fontSize: 12, color, align: 'center', lineSpacingMultiple: 1,
  });
}

/** Standard chrome: logo top-left, header top-right, page number. */
function chrome(slide, n, opts) {
  const o = opts || {};
  logo(slide, 0.382, -0.006);
  header(slide, o.headerColor || C.cocoa);
  pageNum(slide, n, o.pageColor || C.black);
}

/** Eyebrow + big serif title, the recurring page-opening pattern. */
function titleBlock(slide, eyebrow, title, o) {
  const s = o || {};
  txt(slide, eyebrow, {
    x: s.x, y: s.y, w: s.ew || 2.527, h: 0.475,
    fontSize: 24, color: s.eyebrowColor || C.gold, align: s.align || 'left',
  });
  txt(slide, title, {
    x: s.tx !== undefined ? s.tx : s.x, y: s.y + 0.474, w: s.tw, h: s.th || 0.848,
    fontSize: s.size || 48, color: s.color || C.ink, align: s.align || 'left',
  });
}

/** Small caps-ish serif label above a paragraph (e.g. "OVERVIEW"). */
function label(slide, text, o) {
  txt(slide, text, Object.assign({ fontSize: 18, color: C.ink, h: 0.381 }, o));
}

const LOREM_LONG =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nunc ullamcorper sit amet ' +
  'orci et consequat. Morbi semper eros vitae tincidunt porta. Mauris euismod. Cum sociis ' +
  'natoque penatibus. veniam minim as tempor incididunt ut labore';
const LOREM_MED =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nunc ullamcorper sit amet ' +
  'orci et consequat. Morbi semper eros vitae tincidunt porta. Mauris euismod. ';
const LOREM_SHORT =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nunc ullamcorper sit amet ' +
  'orci et consequat. ';
const FUSCE =
  'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet ' +
  'commodo magna eros quis urna. Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus. ' +
  'PLACEHOLDER' +
  'egestas. Proin pharetra nonummy pede. Mauris et orci.';
const NUNC =
  'Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus. Pellentesque habitant morbi ' +
  'tristique senectus et netus et malesuada fames ac turpis egestas. Proin pharetra ' +
  'nonummy pede. Mauris et orci.';
const NUNC_ULLAM =
  'Nunc ullamcorper sit amet orci et consequat. Morbi semper eros vitae tincidunt porta. ' +
  'Mauris euismod. ';
const MORBI = 'Morbi semper eros vitae tincidunt porta. Mauris euismod. ';

// ================================================================ slides ====

// 1 — cover
function slide01(pres) {
  const s = pres.addSlide();
  s.background = { color: C.sand };
  photo(s, 6.667, 0, 5.191, 7.064, { fill: C.imgFill });
  s.addShape('line', {
    x: 2.601, y: 7.052, w: 4.383, h: 0, line: { color: C.dark, width: 1 },
  });
  txt(s, [
    { text: 'Tax', options: { color: C.gold } },
    { text: ' ', options: { color: C.ink } },
  ], { x: 0.484, y: 1.446, w: 2.688, h: 1.471, fontSize: 88 });
  txt(s, 'Advisory ', { x: 0.484, y: 2.494, w: 5.513, h: 1.471, fontSize: 88, color: C.ink });
  body(s, LOREM_LONG, { x: 0.574, y: 4.121, w: 4.662, h: 1.005, fontSize: 10.5, color: C.dark });
  txt(s, 'Strategic Tax Planning for Businesses & Individuals', {
    x: 0.513, y: 6.706, w: 2.088, h: 0.471, fontSize: 11, color: C.dark, lineSpacingMultiple: 1,
  });
  txt(s, 'TAX ADVISORY', {
    x: 0.144, y: 0.172, w: 1.648, h: 0.286, fontSize: 10.5, bold: true, color: C.cocoa,
    charSpacing: 1, lineSpacingMultiple: 1,
  });
  txt(s, 'Maximizing Compliance & Savings', {
    x: 0.131, y: 0.399, w: 2.141, h: 0.236, fontSize: 8, italic: true, color: C.cocoa,
    lineSpacingMultiple: 1,
  });
  logo(s, 12.517, -0.01);
  txt(s, [
    { text: 'Vol. 01 ', options: { fontSize: 14 } },
    { text: '2025', options: { fontSize: 20 } },
  ], { x: 11.994, y: 6.427, w: 1.142, h: 0.673, color: C.gold, lineSpacingMultiple: 1 });
  txt(s, 'Strategic Tax Planning for Businesses & Individuals', {
    x: 11.994, y: 0.637, w: 1.121, h: 0.774, fontSize: 10, color: C.dark, lineSpacingMultiple: 1,
  });
}

// 2 — introduction
function slide02(pres) {
  const s = pres.addSlide();
  chrome(s, 2);
  titleBlock(s, 'Expert ', 'Tax Advisory ', { x: 0.397, y: 1.167, tw: 5.962, th: 1.128, size: 66 });
  txt(s, 'Maximizing Compliance & Savings: Strategic Tax Planning for Businesses & Individuals', {
    x: 5.601, y: 3.151, w: 6.316, h: 1.202, fontSize: 24, color: C.ink,
  });
  body(s, MORBI, {
    x: 8.188, y: 2.037, w: 2.153, h: 0.524, fontSize: 10, italic: true, align: 'right',
  });
  photo(s, 0.503, 3.151, 4.715, 3.833);
  photo(s, 10.521, 1.0, 2.526, 1.561);
  [['INTRODUCTIONS', 5.601], ['OVERVIEW', 9.031]].forEach(function (col) {
    label(s, col[0], { x: col[1], y: 4.939, w: 3.233 });
    body(s, LOREM_LONG, { x: col[1], y: 5.32, w: 3.233, h: 1.659 });
  });
}

// 3 — what is tax advisory
function slide03(pres) {
  const s = pres.addSlide();
  chrome(s, 3);
  titleBlock(s, 'What is', 'Tax Advisory? ', { x: 0.397, y: 1.167, tw: 6.667, th: 1.128, size: 66 });
  label(s, 'OVERVIEW', { x: 1.648, y: 3.215, w: 3.233 });
  body(s, LOREM_MED, { x: 1.648, y: 3.596, w: 4.706, h: 0.871 });
  body(s, FUSCE, { x: 1.648, y: 4.647, w: 4.706, h: 1.396 });
  txt(s, 'Strategic Tax Planning for Businesses & Individuals', {
    x: 1.643, y: 6.333, w: 2.088, h: 0.471, fontSize: 11, italic: true, color: C.gold,
    lineSpacingMultiple: 1,
  });
  [['COMPLIANCE', 7.144, 7.46, 2.153], ['PLANNING', 10.203, 10.416, 2.051]].forEach(function (card) {
    rect(s, { x: card[1], y: 1.167, w: 2.843, h: 4.223, fill: { color: C.sand } });
    label(s, card[0], { x: card[2], y: 1.734, w: card[3], h: 0.412, fontSize: 20 });
    body(s, LOREM_LONG, { x: card[2], y: 2.284, w: 2.264, h: 2.447, align: 'justify' });
  });
  txt(s, 'Maximizing Compliance & Saving with Planning and Risk Mitigation', {
    x: 7.064, y: 5.727, w: 5.872, h: 0.838, fontSize: 24, color: C.ink,
  });
}

// 4 — why is tax advisory important (dark right panel)
function slide04(pres) {
  const s = pres.addSlide();
  rect(s, { x: 6.667, y: 0, w: 6.667, h: 7.5, fill: { color: C.dark } });
  chrome(s, 4, { headerColor: C.pearl, pageColor: C.pearl });
  titleBlock(s, 'Why is', 'Tax Advisory ', { x: 0.397, y: 1.167, tw: 4.993 });
  txt(s, 'is very Important?', { x: 0.484, y: 2.262, w: 6.688, h: 0.848, fontSize: 48, color: C.ink });
  body(s, LOREM_MED, { x: 1.648, y: 3.596, w: 4.534, h: 0.871, align: 'justify' });
  body(s, FUSCE, { x: 1.648, y: 4.647, w: 4.534, h: 1.659, align: 'justify' });

  const items = [
    ['01', 'Enhancing financial efficiency ', 1.641, 7.224, 0.793],
    ['02', 'Maximizing deductions and credits', 3.242, 7.172, 0.898],
    ['03', 'Avoiding legal penalties', 4.934, 7.172, 0.898],
  ];
  items.forEach(function (it) {
    txt(s, it[0], {
      x: it[3], y: it[2], w: it[4], h: 0.707, fontSize: 36, bold: true, italic: true,
      color: C.pearl, align: 'center', lineSpacingMultiple: 1,
    });
    label(s, it[1], { x: 8.481, y: it[2], w: 4.368, color: C.pearl });
    body(s, LOREM_MED, {
      x: 8.481, y: it[2] + 0.381, w: 4.06, h: 0.871, color: C.pearl, align: 'justify',
    });
  });
}

// 5 — profile / efficiency
function slide05(pres) {
  const s = pres.addSlide();
  rect(s, { x: 0, y: 0, w: 3.619, h: 7.5, fill: { color: C.sand } });
  chrome(s, 5);
  txt(s, 'Enhancing financial efficiency', {
    x: 0.397, y: 1.167, w: 2.952, h: 1.569, fontSize: 32, color: C.ink,
  });
  photo(s, 3.826, 1.167, 2.658, 4.696);
  photo(s, 0.484, 3.196, 2.658, 2.667);
  titleBlock(s, 'Why is', 'Tax Advisory ', { x: 6.667, y: 1.167, tw: 4.993 });
  txt(s, 'is very Important?', { x: 6.754, y: 2.262, w: 6.688, h: 0.848, fontSize: 48, color: C.ink });
  label(s, 'EFFICIENCY', { x: 6.754, y: 3.492, w: 3.233 });
  body(s, FUSCE, { x: 6.754, y: 3.873, w: 5.071, h: 1.396 });
  body(s, NUNC, { x: 6.754, y: 5.651, w: 5.071, h: 0.871 });
  body(s, 'Vivamus a tellus. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas.', {
    x: 0.484, y: 6.086, w: 2.658, h: 0.743, fontSize: 10, italic: true,
  });
  txt(s, 'Martina Azzolini', { x: 3.826, y: 6.143, w: 2.658, h: 0.381, fontSize: 18, color: C.gold });
  txt(s, 'Chief Executive Officer', {
    x: 3.837, y: 6.501, w: 2.658, h: 0.319, fontSize: 14, italic: true, color: C.ink,
  });
}

// 6 — digital transformation (full dark slide)
function slide06(pres) {
  const s = pres.addSlide();
  s.background = { color: C.charcoal };
  chrome(s, 6, { headerColor: C.pearl, pageColor: C.pearl });
  photo(s, 4.565, 3.125, 4.203, 3.854, { line: { color: '231C1D', width: 1.5 } });
  txt(s, 'Transformation: ', {
    x: 5.13, y: 1.144, w: 3.074, h: 0.465, fontSize: 24, color: C.orange, align: 'center',
  });
  txt(s, 'Digital Transformation in Tax Advisory', {
    x: 3.148, y: 1.59, w: 7.037, h: 1.434, fontSize: 44, color: C.pearl, align: 'center',
  });
  [['AI-based tax planning', 4.691], ['Automation in tax filings', 5.39],
   ['Cloud-based tax management', 6.235]].forEach(function (it) {
    txt(s, it[0], {
      x: 1.455, y: it[1], w: 2.42, h: 0.654, fontSize: 18, color: C.pearl, align: 'right',
    });
  });
  body(s, 'Fusce est. Vivamus a tellus. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin pharetra nonummy pede. Mauris et orci.', {
    x: 9.049, y: 5.26, w: 3.528, h: 1.134, color: C.pearl,
  });
}

// 7 — what is digital transformation
function slide07(pres) {
  const s = pres.addSlide();
  chrome(s, 7);
  titleBlock(s, 'What is', 'Digital Transformation in Tax Advisory?', {
    x: 0.397, y: 1.083, ew: 2.527, tw: 7.4, th: 1.555,   // wraps after 'Transformation'
  });
  body(s, 'Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin nonummy pede. ', {
    x: 2.158, y: 3.854, w: 4.883, h: 0.871, italic: true, color: C.gold,
  });
  photo(s, 9.651, 2.398, 3.682, 4.582);
  const cols = [
    [2.155, 'Automation \nin tax filings'],
    [4.558, 'AI-based tax planning'],
    [6.962, 'Cloud-based tax management'],
  ];
  cols.forEach(function (c) {
    txt(s, c[1], { x: c[0], y: 5.205, w: 2.168, h: 0.654, fontSize: 18, color: C.ink });
    body(s, LOREM_SHORT, { x: c[0], y: 5.926, w: 2.168, h: 1.134 });
  });
}

// 8 — common tax mistakes
function slide08(pres) {
  const s = pres.addSlide();
  chrome(s, 8);
  titleBlock(s, 'How To Avoid', 'Common Tax Mistakes? ', {
    x: 0.397, y: 1.167, tw: 5.431, th: 1.944, size: 60,
  });
  const rows = [
    ['Missing deadlines', 2.086, 2.022, 1.597, LOREM_MED, 1.134],
    ['Incorrect deductions', 3.402, 3.37, 1.597,
      'Vivamus a tellus. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin pharetra nonummy pede. Mauris et orci. ', 1.134],
    ['Lack of documentation', 4.717, 4.717, 1.946,
      'Fusce est. Vivamus a tellus. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis  est orc', 0.871],
  ];
  rows.forEach(function (r) {
    txt(s, r[0], { x: 7.081, y: r[1], w: r[3], h: 0.654, fontSize: 18, color: C.black });
    body(s, r[4], { x: 9.175, y: r[2], w: 3.648, h: r[5], align: 'justify' });
  });
  photo(s, 3.175, 4.755, 2.356, 2.745);
  txt(s, 'Common Tax Mistakes & How to Avoid Them', {
    x: 5.207, y: 5.209, w: 1.728, h: 0.743, fontSize: 14, color: C.ink, rotate: 90,
  });
}

// 9 — different industries
function slide09(pres) {
  const s = pres.addSlide();
  rect(s, { x: 8.265, y: 0, w: 5.069, h: 7.5, fill: { color: C.sand } });
  chrome(s, 9, { logo: false });
  logo(s, 0.382, -0.006);
  titleBlock(s, 'What is', 'Tax Advisory in Different Industries', {
    x: 0.397, y: 1.167, tw: 7.2, th: 1.575,          // wraps after 'Advisory in'
  });
  txt(s, 'Strategic Tax Planning for Businesses & Individuals in Different Industries', {
    x: 2.632, y: 3.578, w: 4.933, h: 1.202, fontSize: 24, color: C.ink,
  });
  txt(s, 'Industry-specific tax benefits', {
    x: 2.632, y: 5.199, w: 2.456, h: 0.654, fontSize: 18, color: C.ink,
  });
  body(s, NUNC_ULLAM, { x: 2.632, y: 5.859, w: 2.215, h: 1.134 });
  body(s, LOREM_MED, { x: 4.848, y: 5.845, w: 3.123, h: 1.134 });
  photo(s, 8.854, 1.734, 3.911, 2.516);
  photo(s, 8.854, 4.547, 3.911, 2.953);
}

// 10 — future of tax advisory (dark right)
function slide10(pres) {
  const s = pres.addSlide();
  rect(s, { x: 3.921, y: 0, w: 9.413, h: 7.5, fill: { color: C.dark } });
  chrome(s, 10, { headerColor: C.pearl, pageColor: C.pearl });
  titleBlock(s, 'What is', 'Future of Tax Advisory', {
    x: 0.397, y: 1.167, tw: 3.047, th: 2.302,
  });
  txt(s, 'Role of blockchain in tax compliance', {
    x: 0.484, y: 5.414, w: 2.02, h: 1.565, fontSize: 24, color: C.gold,
  });
  photo(s, 4.786, 0.521, 3.889, 6.458);
  label(s, 'TAX REGULATIONS', { x: 9.175, y: 2.78, w: 3.233, color: C.pearl });
  body(s, FUSCE, { x: 9.175, y: 3.162, w: 3.709, h: 1.921, color: C.pearl });
  label(s, 'DIGITALIZATION', { x: 9.173, y: 5.464, w: 3.233, color: C.pearl });
  body(s, 'Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin pharetra nonummy pede. ', {
    x: 9.173, y: 5.845, w: 3.709, h: 1.134, color: C.pearl,
  });
}

// 11 — learn about industries
function slide11(pres) {
  const s = pres.addSlide();
  rect(s, { x: 0, y: 0, w: 6.667, h: 7.5, fill: { color: C.sand } });
  chrome(s, 11);
  titleBlock(s, 'Learn About', 'Tax Advisory in Different Industries', {
    x: 0.397, y: 1.167, tw: 6.269, th: 1.452, size: 44,
  });
  photo(s, 3.556, 3.458, 2.635, 2.401);
  photo(s, 10.3, 1.167, 2.635, 5.812);
  txt(s, 'Tax planning in healthcare, tech, and real estate', {
    x: 6.926, y: 3.459, w: 2.852, h: 1.202, fontSize: 24, color: C.ink,
  });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nunc ullamcorper sit amet orci et consequat. Morbi semper eros', {
    x: 6.926, y: 4.725, w: 3.043, h: 0.871,
  });
  body(s, [
    { text: 'Nunc viverra imperdiet enim. ', options: { bullet: { indent: 13.5 } } },
    { text: 'Fusce est. Vivamus a tellus', options: { bullet: { indent: 13.5 } } },
    { text: 'Pellentesque habitant morbi ', options: { bullet: { indent: 13.5 } } },
    { text: 'tristique senectus et netus', options: { bullet: { indent: 13.5 } } },
  ], { x: 6.926, y: 5.923, w: 3.043, h: 1.134 });
  [['3.2 M+', 3.443, 1.585], ['3210+', 4.959, 1.354]].forEach(function (st) {
    txt(s, st[0], {
      x: st[1], y: 6.161, w: st[2], h: 0.505, fontSize: 24, color: C.charcoal,
      lineSpacingMultiple: 1,
    });
    txt(s, 'Lorem Ipsum', {
      x: st[1], y: 6.594, w: st[2], h: 0.328, fontSize: 11, color: C.charcoal,
      lineSpacingMultiple: 1.3,
    });
  });
}

// 12 — three industry cards on dark
function slide12(pres) {
  const s = pres.addSlide();
  s.background = { color: C.charcoal };
  chrome(s, 12, { headerColor: C.pearl, pageColor: C.pearl });
  txt(s, 'Learn About', {
    x: 5.383, y: 1.261, w: 2.527, h: 0.475, fontSize: 24, color: C.orange, align: 'center',
  });
  txt(s, 'Tax Advisory in Different Industries', {
    x: 3.762, y: 1.736, w: 6.269, h: 1.452, fontSize: 44, color: C.pearl, align: 'center',
  });
  const cards = [
    ['HEALTCARE', 'Nunc ullamcorper sit amet orci et consequat. Morbi semper eros vitae tincidunt porta. Mauris euismod. Cum sociis natoque penatibus. ', 1.396, '25%', 2.04, 2.356, 2.346, 3.433],
    ['TECH', 'Morbi semper eros vitae tincidunt porta. Mauris euismod. Cum sociis natoque penatibus. veniam', 1.134, '35%', 5.199, 5.515, 5.515, 6.601],
    ['REAL ESTATE', 'Mauris euismod. Cum sociis natoque penatibus. veniam minim as tempor incididunt ut labore', 1.134, '30%', 8.403, 8.719, 8.719, 9.805],
  ];
  cards.forEach(function (c) {
    rect(s, { x: c[4], y: 3.568, w: 2.843, h: 3.932, fill: { color: C.sand } });
    label(s, c[0], { x: c[5], y: 3.983, w: 2.153, h: 0.412, fontSize: 20 });
    body(s, c[1], { x: c[5], y: 4.533, w: 2.264, h: c[2], align: 'justify' });
    txt(s, [
      { text: 'Tax ', options: { fontSize: 18 } },
      { text: c[3], options: { fontSize: 28 } },
    ], { x: c[6], y: 6.173, w: 1.142, h: 0.875, color: C.ink, lineSpacingMultiple: 1 });
    body(s, MORBI, { x: c[7], y: 6.309, w: 1.188, h: 0.614, fontSize: 8, italic: true });
  });
}

// --- device mock-ups drawn with native shapes (replacing the PNG mock-ups) ---
const DEVICE_FRAME = '3B3036';

/**
 * Phone shell built around the screen rectangle (the coordinates the reference
 * uses for its screenshot): body, notch and side buttons hang off that rect.
 */
function phoneMock(slide, sx, sy, sw, sh) {
  const x = sx - 0.119, y = sy - 0.103, w = sw + 0.252, h = sh + 0.193;
  slide.addShape('roundRect', {
    x, y, w, h, rectRadius: 0.3, fill: { color: DEVICE_FRAME }, line: { type: 'none' },
  });
  slide.addShape('roundRect', {
    x: sx, y: sy, w: sw, h: sh, rectRadius: 0.24,
    fill: { color: C.imgFill }, line: { type: 'none' },
  });
  slide.addShape('roundRect', {
    x: x + w / 2 - 0.21, y: y + 0.145, w: 0.42, h: 0.095, rectRadius: 0.047,
    fill: { color: '1D1B1E' }, line: { type: 'none' },
  });
  [[x - 0.025, y + 1.0, 0.4], [x + w, y + 0.85, 0.25], [x + w, y + 1.25, 0.55]]
    .forEach(function (b) {
      slide.addShape('rect', {
        x: b[0], y: b[1], w: 0.025, h: b[2], fill: { color: DEVICE_FRAME }, line: { type: 'none' },
      });
    });
  slide.addText('[image]', {
    x: sx, y: sy, w: sw, h: sh, align: 'center', valign: 'middle',
    fontFace: SANS, fontSize: 11, color: C.imgText,
  });
}

/**
 * Laptop shell built around the screen rectangle: dark lid, tapered silver
 * base, keyboard block and trackpad lip.
 */
function laptopMock(slide, sx, sy, sw, sh) {
  const lidX = sx - 0.130, lidY = sy - 0.107, lidW = sw + 0.258, lidH = sh + 0.145;
  const baseY = lidY + lidH;         // 5.72 in the source artwork
  const baseH = 0.827;               // base runs down to 6.547
  slide.addShape('rect', {
    x: lidX, y: lidY, w: lidW, h: lidH, fill: { color: '2C2B31' }, line: { type: 'none' },
  });
  slide.addShape('rect', { x: sx, y: sy, w: sw, h: sh, fill: { color: C.imgFill } });
  slide.addShape('trapezoid', {
    x: lidX - 0.472, y: baseY, w: lidW + 0.944, h: baseH,
    fill: { color: 'D9D6D4' }, line: { type: 'none' },
  });
  slide.addShape('trapezoid', {
    x: lidX + 0.1, y: baseY - 0.04, w: lidW - 0.2, h: 0.435,
    fill: { color: '2F2E33' }, line: { type: 'none' },
  });
  slide.addShape('roundRect', {
    x: lidX + lidW / 2 - 0.52, y: baseY + 0.66, w: 1.04, h: 0.045, rectRadius: 0.022,
    fill: { color: 'C4C1BF' }, line: { type: 'none' },
  });
  slide.addText('[image]', {
    x: sx, y: sy, w: sw, h: sh, align: 'center', valign: 'middle',
    fontFace: SANS, fontSize: 11, color: C.imgText,
  });
}

// 13 — phone mockup
function slide13(pres) {
  const s = pres.addSlide();
  rect(s, { x: 0, y: 0, w: 8.281, h: 7.5, fill: { color: C.sand } });
  chrome(s, 13);
  txt(s, 'Tax Advisory', { x: 0.484, y: 1.271, w: 2.527, h: 0.475, fontSize: 24, color: C.gold });
  txt(s, 'Phone Mockup', { x: 0.484, y: 1.746, w: 5.512, h: 0.848, fontSize: 48, color: C.ink });
  body(s, FUSCE, { x: 1.648, y: 2.855, w: 4.534, h: 1.659, align: 'justify' });
  txt(s, '275.292', {
    x: 1.65, y: 4.775, w: 3.048, h: 0.707, fontSize: 36, bold: true, color: C.gold,
    lineSpacingMultiple: 1,
  });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ', {
    x: 1.648, y: 5.443, w: 4.492, h: 0.599, italic: true, color: C.gold,
  });
  phoneMock(s, 7.483, 0.955, 2.059, 4.458);
  phoneMock(s, 10.021, 2.304, 2.059, 4.457);
}

// 14 — desktop mockup
function slide14(pres) {
  const s = pres.addSlide();
  s.background = { color: C.sand };
  chrome(s, 14, { pageColor: C.ink });
  txt(s, 'Tax Advisory', {
    x: 5.403, y: 0.996, w: 2.527, h: 0.475, fontSize: 24, color: C.gold, align: 'center',
  });
  txt(s, 'Dekstop Mockup', {
    x: 3.566, y: 1.319, w: 6.269, h: 0.786, fontSize: 44, color: C.black, align: 'center',
  });
  laptopMock(s, 4.212, 2.557, 4.948, 3.125);
  const blurb = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nunc ullamcorper sit amet orci et consequat. Morbi semper eros vitae.';
  body(s, blurb, { x: 0.904, y: 3.336, w: 2.662, h: 1.396, align: 'right' });
  body(s, blurb, { x: 9.806, y: 3.336, w: 2.729, h: 1.134 });
  txt(s, '3210+', {
    x: 1.455, y: 4.941, w: 2.076, h: 0.64, fontSize: 32, bold: true, color: C.gold,
    align: 'right', lineSpacingMultiple: 1,
  });
  txt(s, 'Lorem Ipsum', {
    x: 1.845, y: 5.441, w: 1.695, h: 0.345, fontSize: 12, color: C.gold, align: 'right',
    lineSpacingMultiple: 1.3,
  });
  txt(s, '3.2 M+', {
    x: 9.802, y: 4.941, w: 2.076, h: 0.64, fontSize: 32, bold: true, color: C.gold,
    lineSpacingMultiple: 1,
  });
  txt(s, 'Lorem Ipsum', {
    x: 9.802, y: 5.441, w: 1.695, h: 0.345, fontSize: 12, color: C.gold, lineSpacingMultiple: 1.3,
  });
}

// 15 — report analysis (native line chart)
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const SERIES_DARK = [0, 5000, 5300, 10000, 1350, 12000, 9600, 9600, 10600, 6500, 3100, 15000];
const SERIES_ROSE = [0, 4000, 4000, 8100, 12900, 10000, 11000, 8500, 13900, 5100, 12600, 13800];

function slide15(pres) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  chrome(s, 15);
  txt(s, '2025 Data Report', {
    x: 1.003, y: 1.136, w: 4.956, h: 0.488, fontSize: 20, bold: true, color: C.dark,
    align: 'center', lineSpacingMultiple: 1,
  });
  // Manual layout pins the plot box onto the same grid the original artwork used:
  // gridlines span x 1.564 -> 5.777 and y 2.101 (20000) -> 5.971 (0).
  s.addChart('line', [
    { name: 'Series 1', labels: MONTHS, values: SERIES_DARK },
    { name: 'Series 2', labels: MONTHS, values: SERIES_ROSE },
  ], {
    x: 0.917, y: 1.85, w: 5.042, h: 4.6,
    layout: { x: 0.128, y: 0.055, w: 0.836, h: 0.841 },
    chartColors: [C.dark, C.mauve],
    lineSize: 2.25, lineDataSymbol: 'circle', lineDataSymbolSize: 4,
    showLegend: false, showTitle: false,
    valAxisMinVal: 0, valAxisMaxVal: 20000, valAxisMajorUnit: 5000,
    valAxisLabelColor: C.grey59, valAxisLabelFontFace: SANS, valAxisLabelFontSize: 10,
    catAxisLabelColor: C.grey59, catAxisLabelFontFace: SERIF, catAxisLabelFontSize: 10.5,
    valGridLine: { style: 'solid', color: C.grid, size: 0.75 },
    catGridLine: { style: 'none' },
    valAxisLineShow: false, catAxisLineShow: false,
  });
  // "+15%" callout bubble hovering above the chart
  s.addShape('roundRect', {
    x: 3.958, y: 2.224, w: 1.23, h: 0.582, rectRadius: 0.09,
    fill: { color: C.brown }, line: { type: 'none' },
  });
  s.addShape('triangle', {
    x: 4.454, y: 2.786, w: 0.238, h: 0.247, rotate: 180,
    fill: { color: C.brown }, line: { type: 'none' },
  });
  txt(s, '+15%', {
    x: 4.185, y: 2.239, w: 0.776, h: 0.338, fontSize: 12, bold: true, color: C.white,
    align: 'center', lineSpacingMultiple: 1,
  });
  txt(s, 'Your text', {
    x: 4.091, y: 2.466, w: 0.964, h: 0.301, fontFace: SANS, fontSize: 10, color: C.white,
    align: 'center', lineSpacingMultiple: 1,
  });
  txt(s, 'Infograpic', { x: 6.842, y: 3.396, w: 2.527, h: 0.475, fontSize: 24, color: C.gold });
  txt(s, 'Report Analysis', { x: 6.842, y: 3.871, w: 5.512, h: 0.848, fontSize: 48, color: C.ink });
  label(s, 'TAX REGULATIONS', { x: 6.842, y: 5.02, w: 4.326, color: C.black });
  body(s, 'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin pharetra nonummy pede. ', {
    x: 6.842, y: 5.401, w: 5.512, h: 1.134,
  });
}

// 16 — finance comparison: three sand panels, each with four hand-drawn bars.
// Bar heights come straight from the source artwork (they are not strictly
// proportional to the printed percentages).
const BAR_BASE = 6.54;   // y where every bar meets its dark quarter block
const BAR_FOOT = 0.233;  // height of that dark block
const BAR_W = 0.399;
const YEAR_PANELS = [
  { title: 'Year 2023', x: 0.518, titleX: 0.636, barX: 0.842,
    bars: [[56, 0.938, C.dark], [93, 1.894, C.rose], [70, 1.316, C.brown], [84, 1.676, C.ochre]] },
  { title: 'Year 2024', x: 4.077, titleX: 4.191, barX: 4.396,
    bars: [[40, 0.938, C.dark], [90, 1.894, C.rose], [70, 1.316, C.brown], [83, 1.676, C.lime]] },
  { title: 'Year 2025', x: 7.636, titleX: 7.793, barX: 7.999,
    bars: [[40, 0.938, C.dark], [90, 1.821, C.rose], [70, 0.938, C.brown], [70, 1.521, C.ochre]] },
];

function slide16(pres) {
  const s = pres.addSlide();
  chrome(s, 16);
  txt(s, 'Infograpic', { x: 0.484, y: 1.271, w: 2.527, h: 0.475, fontSize: 24, color: C.gold });
  txt(s, 'Finance Comparison', { x: 0.484, y: 1.746, w: 7.325, h: 0.848, fontSize: 48, color: C.ink });

  YEAR_PANELS.forEach(function (panel) {
    rect(s, { x: panel.x, y: 3.19, w: 2.843, h: 3.79, fill: { color: C.sand } });
    txt(s, panel.title, {
      x: panel.titleX, y: 3.563, w: 2.528, h: 0.337, fontSize: 14, bold: true, color: C.black,
      align: 'center', lineSpacingMultiple: 1,
    });
    panel.bars.forEach(function (bar, i) {
      const pct = bar[0], bh = bar[1];
      const bx = panel.barX + i * 0.5687;
      const cx = bx + BAR_W / 2 - 0.333;   // centre a 0.666"-wide caption on the bar
      s.addShape('rect', { x: bx, y: BAR_BASE - bh, w: BAR_W, h: bh, fill: { color: bar[2] } });
      s.addShape('rect', { x: bx, y: BAR_BASE, w: BAR_W, h: BAR_FOOT, fill: { color: '3F3F3F' } });
      s.addText([
        { text: String(pct), options: { fontSize: 14, bold: true } },
        { text: '%', options: { fontSize: 14, bold: true, superscript: true } },
      ], {
        x: cx, y: BAR_BASE - bh - 0.388, w: 0.666, h: 0.387,
        fontFace: SANS, color: C.black, align: 'center', valign: 'middle',
        margin: INSETS, lineSpacingMultiple: 1.3,
      });
      s.addText('Q' + (i + 1), {
        x: cx, y: BAR_BASE - 0.026, w: 0.666, h: 0.269,
        fontFace: SANS, fontSize: 10, color: C.white, align: 'center', valign: 'middle',
        margin: INSETS,
      });
    });
  });
}

/**
 * The three small white glyphs sitting on the pie wedges: overlapping circles,
 * a pair of cogs, and a light bulb — each drawn from primitives.
 */
function pieIcon(slide, kind, x, y) {
  const W = { color: C.white }, none = { type: 'none' };
  if (kind === 0) {          // two cogs
    [[0.00, 0.10, 0.17], [0.16, 0.00, 0.20]].forEach(function (c) {
      slide.addShape('gear9', {
        x: x + c[0], y: y + c[1], w: c[2], h: c[2], fill: W, line: none,
      });
    });
  } else if (kind === 1) {   // three overlapping circles
    [[0.11, 0.00], [0.00, 0.14], [0.22, 0.14]].forEach(function (c) {
      slide.addShape('ellipse', {
        x: x + c[0], y: y + c[1], w: 0.21, h: 0.21,
        fill: none, line: { color: C.white, width: 1.75 },
      });
    });
  } else {                   // light bulb
    slide.addShape('ellipse', { x: x + 0.09, y: y + 0.02, w: 0.26, h: 0.26, fill: none, line: { color: C.white, width: 1.75 } });
    slide.addShape('rect', { x: x + 0.16, y: y + 0.27, w: 0.12, h: 0.04, fill: W, line: none });
    slide.addShape('rect', { x: x + 0.17, y: y + 0.33, w: 0.10, h: 0.04, fill: W, line: none });
  }
}

// 17 — half-donut pie chart made of three pie wedges
function slide17(pres) {
  const s = pres.addSlide();
  chrome(s, 17);
  txt(s, 'Infograpic', { x: 0.484, y: 1.029, w: 2.527, h: 0.475, fontSize: 24, color: C.gold });
  txt(s, 'Pie Chart', { x: 0.484, y: 1.504, w: 5.512, h: 0.848, fontSize: 48, color: C.ink });
  txt(s, 'Tax planning in healthcare, tech, and real estate', {
    x: 0.484, y: 2.484, w: 2.852, h: 1.202, fontSize: 24, color: C.ink,
  });
  body(s, LOREM_SHORT, { x: 0.484, y: 3.75, w: 3.234, h: 0.871 });

  // wedges: [x, y, size, startAngle, endAngle, colour, iconX, iconY]
  const wedges = [
    [5.148, 4.945, 5.134, 180.0, 229.7, C.mauve, 5.932, 6.315],
    [5.213, 4.839, 5.264, 229.6, 299.1, C.dark, 7.485, 5.391],
    [5.405, 4.930, 5.134, 299.0, 359.8, C.tan, 9.137, 6.176],
  ];
  wedges.forEach(function (w, i) {
    s.addShape('pie', {
      x: w[0], y: w[1], w: w[2], h: w[2],
      angleRange: [w[3], w[4]], fill: { color: w[5] }, line: { type: 'none' },
    });
    pieIcon(s, i, w[6], w[7]);
  });
  [['35', 5.628, 6.719], ['25', 7.106, 5.816], ['45', 8.758, 6.670]].forEach(function (v) {
    s.addText([
      { text: v[0], options: { fontSize: 20, bold: true } },
      { text: '%', options: { fontSize: 20, bold: true, superscript: true } },
    ], {
      x: v[1], y: v[2], w: 1.196, h: 0.569, fontFace: SANS, color: C.white,
      align: 'center', valign: 'middle', margin: 0,
    });
  });
  const callouts = [
    ['Tax-saving investment options', 6.46, 2.657, 2.76, 'center', 0.871],
    ['Income tax deductions', 2.701, 4.731, 2.215, 'right', 1.134],
    ['Industry-specific tax benefits', 10.669, 4.731, 2.456, 'left', 1.134],
  ];
  callouts.forEach(function (c) {
    txt(s, c[0], {
      x: c[1], y: c[2], w: c[3], h: 0.654, fontSize: 18, color: C.ink, align: c[4],
    });
    body(s, NUNC_ULLAM, {
      x: c[1], y: c[2] + 0.66, w: c[4] === 'left' ? 2.215 : c[3], h: c[5], align: c[4],
    });
  });
}

// 18 — three growing circles
function slide18(pres) {
  const s = pres.addSlide();
  chrome(s, 18);
  txt(s, 'Infograpic', { x: 0.484, y: 1.029, w: 2.527, h: 0.475, fontSize: 24, color: C.gold });
  txt(s, 'Stage Tax', { x: 0.484, y: 1.504, w: 5.512, h: 0.848, fontSize: 48, color: C.ink });
  // the two hairlines fanning out behind the circles (both rise to the right)
  s.addShape('line', {
    x: 3.326, y: 1.086, w: 7.508, h: 2.635, flipV: true, line: { color: C.hair, width: 1 },
  });
  s.addShape('line', {
    x: 3.771, y: 3.941, w: 8.090, h: 1.262, flipV: true, line: { color: C.hair, width: 1 },
  });
  const stages = [
    { pct: '40%', size: 1.630, cx: 2.956, cy: 3.574, color: C.mauve, fontSize: 36,
      label: 'Healthcare Tax', lx: 2.956, ly: 5.678, ty: 4.040, tw: 1.630 },
    { pct: '55%', size: 1.966, cx: 6.230, cy: 2.522, color: C.clay, fontSize: 40,
      label: 'Tech Tax', lx: 6.230, ly: 5.648, ty: 3.139, tw: 1.969 },
    { pct: '60%', size: 2.489, cx: 9.711, cy: 1.372, color: C.tan, fontSize: 40,
      label: 'Real Estate Tax', lx: 9.711, ly: 5.648, ty: 2.270, tw: 1.969 },
  ];
  stages.forEach(function (st) {
    s.addShape('ellipse', {
      x: st.cx, y: st.cy, w: st.size, h: st.size, fill: { color: st.color }, line: { type: 'none' },
    });
    txt(s, st.pct, {
      x: st.cx, y: st.ty, w: st.tw, h: st.size > 1.8 ? 0.774 : 0.707, fontSize: st.fontSize,
      bold: true, color: C.white, align: 'center', lineSpacingMultiple: 1,
    });
    txt(s, st.label, { x: st.lx, y: st.ly, w: 2.939, h: 0.381, fontSize: 18, color: C.ink });
    body(s, NUNC_ULLAM, { x: st.lx, y: st.ly + 0.428, w: 2.939, h: 0.871 });
  });
}

// 19 — pricing table
const TABLE_HEAD = ['Regular', 'Professional', 'Promotion', 'Business', 'Advance'];
const TABLE_BODY = [
  ['First Service', 'First Service', 'First Service', 'First Service', 'First Service'],
  ['-', 'Second Service', 'Second Service', 'Second Service', 'Second Service'],
  ['-', '-', 'Third Service', 'Third Service', 'Third Service'],
  ['-', '-', '-', 'Fourth Service', 'Fourth Service'],
  ['-', '-', '-', 'Fifth Service', 'Fifth Service'],
  ['-', '-', '-', '-', 'Sixth Service'],
  ['-', '-', '-', '-', 'Seventh Service'],
];
const TABLE_FOOT = ['$15', '$29', '$39', '$49', '$199'];

function slide19(pres) {
  const s = pres.addSlide();
  chrome(s, 19);
  txt(s, 'Infograpic', { x: 0.699, y: 0.974, w: 2.527, h: 0.475, fontSize: 24, color: C.gold });
  txt(s, 'Table Plan', { x: 0.699, y: 1.449, w: 5.512, h: 0.848, fontSize: 48, color: C.ink });

  const border = [{ type: 'solid', color: C.hair, pt: 0.75 }];   // L, R, T, B
  const headFills = [C.dark, C.mauve, C.dark, C.mauve, C.dark];
  const rows = [];
  rows.push(TABLE_HEAD.map(function (t, i) {
    return { text: t, options: { fill: { color: headFills[i] }, color: C.white, fontSize: 15 } };
  }));
  TABLE_BODY.forEach(function (r) {
    rows.push(r.map(function (t, i) {
      return { text: t, options: { fill: { color: i % 2 === 0 ? C.cream : C.white }, color: C.black, fontSize: 11 } };
    }));
  });
  rows.push(TABLE_FOOT.map(function (t) {
    return { text: t, options: { fill: { color: C.ochre }, color: C.white, fontSize: 16 } };
  }));

  s.addTable(rows, {
    x: 0.699, y: 2.551, colW: [2.019, 2.019, 2.019, 2.043, 1.996], rowH: 0.475,
    align: 'center', valign: 'middle', fontFace: TABLE_FONT, border: border, margin: 1.8,
    autoPage: false,
  });
}

// 20 — contact
function slide20(pres) {
  const s = pres.addSlide();
  rect(s, { x: 8.921, y: 0, w: 4.413, h: 7.5, fill: { color: C.dark } });
  chrome(s, 20, { headerColor: C.pearl, pageColor: C.ink });
  txt(s, [
    { text: 'Contact', options: { color: C.gold } },
    { text: ' ', options: { color: C.ink } },
  ], { x: 0.484, y: 3.626, w: 4.944, h: 1.346, fontSize: 80 });
  txt(s, 'Our Company ', { x: 0.484, y: 4.695, w: 8.705, h: 1.346, fontSize: 80, color: C.ink });
  txt(s, 'Strategic Tax Planning for Businesses & Individuals', {
    x: 0.484, y: 6.125, w: 5.262, h: 0.828, fontSize: 24, color: C.gold,
  });
  const contact = [
    ['Office Address:', 'New York, Downtown Main Street 233, USA', 3.399, 0.693],
    ['Phone:', '+22 1234 5678 9000', 4.695, 0.387],
    ['E-mail:', 'taxadvisory@company.com', 5.685, 0.387],
  ];
  contact.forEach(function (c) {
    body(s, c[0], { x: 9.422, y: c[2], w: 3.427, h: 0.469, fontSize: 18, color: C.pearl });
    body(s, c[1], { x: 9.422, y: c[2] + 0.44, w: 3.427, h: c[3], fontSize: 14, color: C.pearl });
  });
}

// ================================================================== build ===
function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
  pres.layout = 'WIDE';
  pres.title = 'Tax Advisory';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
   slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach(function (fn) { fn(pres); });

  const out = path.join(__dirname, '16a07a18-341d-4541-b4b2-09d12196f5cd_grok_final.pptx');
  return pres.writeFile({ fileName: out }).then(function () { console.log('wrote', out); });
}

build().catch(function (e) { console.error(e); process.exit(1); });
