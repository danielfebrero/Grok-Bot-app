/**
 * "Iduna Creative Studio - Branding Guideline" - 44 slides, 26.667" x 15".
 *
 * Rebuilt from scratch with pptxgenjs. Every position is in inches and matches
 * the source deck. The deck's picture placeholders are empty in the original
 * (no media is embedded anywhere in the file), so `photo()` lays out the same
 * frames without painting anything over the page.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const C = {
  tomato: 'FF644E',
  tomatoMid: 'FF846B',
  tomatoPale: 'FFA489',
  ink: '000000',
  charcoal: '363636',
  grey: '5E5E5E',
  greyMid: 'B0B5B3',
  silver: 'D5D5D5',
  sand: 'E6E6E3',
  mist: 'EDEAEA',
  paper: 'F6F7F7',
  white: 'FFFFFF',
};

const F = {
  head: 'Dosis ExtraLight Bold', // theme font of the source deck
  dosis: 'Dosis',
  dosisLight: 'Dosis Light',
  dosisMedium: 'Dosis Medium',
  dosisSemi: 'Dosis SemiBold',
  dosisXBold: 'Dosis ExtraBold',
  sans: 'Open Sans Regular',
  sansLight: 'Open Sans Light',
  sansSemi: 'Open Sans SemiBold',
  sansBold: 'Open Sans Bold',
  sansXBold: 'Open Sans ExtraBold',
};

const INSET = 5.625;   // 0.078125" text inset, used by every box in the deck
const DIM = 20;        // 80% opacity black body copy
const LS = 1.5;        // the deck's default paragraph line spacing (150%)

/* ------------------------------------------------------------ text presets */

const S = {
  // 177pt page numeral on the section dividers
  numeral: { fontFace: F.dosis, fontSize: 177, bold: true, color: C.tomato, lineSpacingMultiple: 0.8, wrap: false },
  // 110pt display type
  display: { fontFace: F.dosis, fontSize: 110, bold: true, color: C.tomato, lineSpacingMultiple: 0.9, charSpacing: -2.2, wrap: false },
  // 42pt tomato section heading
  head: { fontFace: F.dosis, fontSize: 42, bold: true, color: C.tomato, lineSpacingMultiple: 1 },
  // 42pt black sub heading
  sub: { fontFace: F.dosis, fontSize: 42, bold: true, color: C.ink, lineSpacingMultiple: 1 },
  // 42pt table-of-contents row
  index: { fontFace: F.dosis, fontSize: 42, color: C.ink, lineSpacingMultiple: 1 },
  // 26pt bold caption (26pt is the deck-wide default text size)
  caption: { fontFace: F.dosis, fontSize: 26, bold: true, color: C.ink, lineSpacingMultiple: LS },
  // 16pt running text
  body: { fontFace: F.sansLight, fontSize: 16, color: C.ink, transparency: DIM, lineSpacingMultiple: LS },
  bodyBold: { fontFace: F.sansBold, fontSize: 16, color: C.ink, transparency: DIM, lineSpacingMultiple: LS },
  // 16pt label set in the theme font
  label: { fontFace: F.head, fontSize: 16, color: C.ink, transparency: DIM, lineSpacingMultiple: LS },
};

/* --------------------------------------------------------------- primitives */

/**
 * Text box. Every text frame in the source deck is "resize shape to fit text"
 * with a 0.078" inset and vertically centred copy; unstyled runs fall back to
 * the deck-wide default of 26pt / 150% in the theme font.
 *
 * pptxgenjs writes back into the option objects it is handed, so the shared
 * presets in `S` are copied on the way in.
 */
function text(slide, runs, o) {
  const opts = Object.assign({
    margin: INSET, valign: 'middle', fit: 'resize',
    fontFace: F.head, fontSize: 26, color: C.ink, lineSpacingMultiple: LS,
  }, o);
  const body = Array.isArray(runs)
    ? runs.map((r) => ({ text: r.text, options: Object.assign({}, r.options) }))
    : runs;
  slide.addText(body, opts);
}

/**
 * Flattens `[[run, run], [run]]` into one run list with a hard paragraph break
 * after each group. A group of `null` becomes a blank line; it carries the
 * style of the preceding paragraph so it takes up the same height.
 */
function paragraphs(groups) {
  const out = [];
  let last = {};
  groups.forEach((group, i) => {
    const runs = group || [{ text: ' ', options: last }];
    last = runs[runs.length - 1].options || last;
    runs.forEach((run, j) => {
      const options = Object.assign({}, run.options);
      if (j === runs.length - 1 && i < groups.length - 1) options.breakLine = true;
      out.push({ text: run.text, options });
    });
  });
  return out;
}

function rect(slide, x, y, w, h, color, o) {
  slide.addShape('rect', Object.assign({ x, y, w, h, fill: { color } }, o));
}

function circle(slide, x, y, d, color, o) {
  slide.addShape('ellipse', Object.assign({ x, y, w: d, h: d, fill: { color } }, o));
}

/** Empty picture placeholder from the source deck - frame only, no artwork. */
function photo(slide, x, y, w, h) {
  slide.addShape('rect', { x, y, w, h });
}

/** The short 0.79" x 0.04" rule that sits above most sub headings. */
function rule(slide, x, y, color) {
  rect(slide, x, y, 0.7874, 0.0394, color || C.ink);
}

/* -------------------------------------------------------------- page chrome */

// Bottom-left section label + slide number: [label, labelWidth, numberWidth].
const FOOTER = {
  2: ['Table of Content', 1.6488, 0.2824], 3: ['Table of Content', 1.6488, 0.2859],
  4: ['Brand', 0.7133, 0.2817], 5: ['Brand Overview', 1.6008, 0.283],
  6: ['Brand History', 1.4099, 0.2921], 7: ['Mission Statement', 1.8421, 0.2806],
  8: ['Vision', 0.7184, 0.2886], 9: ['Tone of Voice', 1.3559, 0.2924],
  10: ['Personality', 1.191, 0.4242], 11: ['Key Values', 1.1448, 0.2977],
  12: ['Logo', 0.5999, 0.3461], 13: ['Logo Usage', 1.1841, 0.3497],
  14: ['Primary Logo', 1.3724, 0.3455], 15: ['Logo Variations', 1.5524, 0.3468],
  16: ['Logo Breakdowns', 1.7586, 0.3559], 17: ['Colour', 0.7681, 0.3444],
  18: ['Introduction : Color', 1.9104, 0.3524], 19: ['Color Palette', 1.3435, 0.3561],
  20: ['Color Palette', 1.3435, 0.4066], 21: ['Logo Colour', 1.2424, 0.3461],
  22: ['Typography', 1.2255, 0.3946], 23: ['Introduction : Typography', 2.4841, 0.3981],
  24: ['Primary Typeface', 1.7501, 0.389], 25: ['Primary Typeface', 1.7501, 0.3953],
  26: ['Secondary Typeface', 1.9764, 0.4044], 27: ['Secondary Typeface', 1.9764, 0.3928],
  28: ['Use of Type', 1.2017, 0.4008], 29: ['Typography Hierarchy', 2.1473, 0.4046],
  30: ['Print', 0.619, 0.4101], 31: ['Introduction : Print', 1.8777, 0.3497],
  32: ['Letterhead', 1.1697, 0.3981], 33: ['Stationery', 1.1175, 0.4017],
  34: ['Digital', 0.741, 0.3975], 35: ['Introduction : Digital', 1.9997, 0.3988],
  36: ['Web Design', 1.2237, 0.4079], 37: ['Mobile Design', 1.4099, 0.3964],
  38: ['Imagery', 0.9093, 0.4044], 39: ['Introduction : Imagery', 2.1679, 0.4081],
  40: ['Inspirational', 1.3006, 0.4059], 41: ['Photography Style', 1.8384, 0.3419],
  42: ['Photography Style', 1.8384, 0.3939], 43: ['Quote', 0.7073, 0.3986],
};

// Slides whose footer label / slide number / wordmark are set in bold.
const BOLD_FOOTER = new Set([2, 3, 4, 5, 8, 10, 12, 13, 15, 17, 18, 19, 20, 22, 23,
  24, 25, 27, 28, 29, 30, 31, 34, 38, 39]);

/**
 * Header rule, "IDUNA Creative Studio" wordmark, footer rule, section label,
 * slide number and the small logo lockup in the bottom-right corner.
 */
function chrome(slide, n, dark) {
  const fg = dark ? C.white : C.ink;
  const bold = BOLD_FOOTER.has(n);
  const [label, labelW, numW] = FOOTER[n];
  const wide = n === 2 || n === 3;  // these two pages use the taller header box

  rect(slide, 1.6475, 1.9893, 23.3717, 0.0139, fg);
  rect(slide, 1.6475, 13.0145, 23.3717, 0.0139, fg);

  // The three placeholders (wordmark, section label, page number) are the only
  // frames in the deck without "resize shape to fit text".
  text(slide, [
    { text: 'IDUNA  ', options: { fontFace: F.head, bold } },
    { text: 'Creative Studio', options: { fontFace: F.sansLight, bold: bold && dark } },
  ], {
    x: 1.6475, y: wide ? 1.1213 : 1.1502, w: wide ? 2.4192 : 2.348, h: wide ? 0.5196 : 0.4618,
    fontSize: 16, color: fg, transparency: dark ? 0 : DIM, lineSpacingMultiple: 1.5, fit: null,
  });

  text(slide, label, {
    x: 2.2076, y: n === 2 ? 13.356 : 13.4115, w: labelW, h: 0.434, fit: null,
    fontFace: F.head, fontSize: 16, bold, color: fg, lineSpacingMultiple: 1.5,
  });
  text(slide, String(n), {
    x: 25.0192 - 23.3717 - numW / 2 + 0.0722, y: 13.4115, w: numW, h: 0.434, fit: null,
    fontFace: F.head, fontSize: 16, bold, color: fg, align: 'center',
  });

  circle(slide, 23.3667, 13.4115, 0.5885, dark ? C.mist : C.silver);
  circle(slide, 23.0724, 13.4115, 0.5885, dark ? C.white : C.tomato,
    { fill: { color: dark ? C.white : C.tomato, transparency: 50 } });
  text(slide, 'IDUNA', {
    x: 24.0941, y: wide ? 13.3686 : 13.4115, w: 0.7906, h: wide ? 0.5199 : 0.434,
    fontFace: F.head, fontSize: 16, bold: n === 2, color: fg, transparency: dark ? 0 : DIM, wrap: false,
  });
  text(slide, 'Creative Studio', {
    x: 24.0941, y: 13.691, w: 0.9251, h: 0.309,
    fontFace: F.sansLight, fontSize: 8, color: fg, transparency: dark ? 0 : DIM, wrap: false,
  });
}

/** Two overlapping circles: the Iduna brand mark. */
function logoMark(slide, x, y, d) {
  circle(slide, x + d * 0.5, y, d, C.mist);
  circle(slide, x, y, d, C.tomato, { fill: { color: C.tomato, transparency: 50 } });
}

/* -------------------------------------------- shared body copy of the deck */

const T = {};
T.head = 'Iduna Creative Studio';
T.a = ' is dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ';
T.b = 'Dui ut ornare lectus sit amet est placerat.';
T.c = ' Bibendum enim facilisis gravida neque convallis a cras semper. Condimentum mattis pellentesque id nibh tortor. Sed cras ornare arcu dui. Ornare lectus sit amet est. Senectus et netus et malesuada fames ac t';
T.d = 'urpis egestas maecenas. Consequat interdum varius sit amet mattis';
T.e = ' vulputate enim nulla. Urna cursus eget nunc scelerisque viverra mauris in. Interdum velit euismod in pellentesque massa placerat duis. Odio morbi quis commodo odio aenean sed. ';
T.f = 'Commodo quis imperdiet massa tincidunt. Gravida rutrum quisque non tellus orci. Erat nam at lectus urna duis convallis convallis tellus. Eget est lorem ipsum dolor sit amet. Erat nam at lectus urna duis convallis';
T.ultrices = 'Ultrices gravida dictum fusce ut placerat. Nascetur ridiculus mus mauris vitae ultricies leo. ';
T.aliquam = 'Aliquam sem et tortor consequat id porta. Sed euismod nisi porta lorem mollis aliquam ut porttitor leo. Imperdiet proin fermentum leo vel orci ';

T.short = T.head + T.a + T.b;                              // 2 sentences
T.short2 = T.short + ' Bibendum enim facilisis gravida neque convallis a cras semper.';
T.medium = T.head + T.a + T.b + T.c + 'urpis';             // trimmed mid-sentence
T.long = T.head + T.a + T.b + T.c + T.d + T.e;
T.longer = T.long + T.f + ' ' + T.ultrices;
T.stub = T.head + ' is dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore';

T.logoCopy = T.short + ' ' + T.aliquam + 'Aliquam sem et tortor consequat id porta. Sed euismod nisi porta lorem mollis aliquam ut porttitor leo.';
T.logoCopyLong = T.logoCopy + ' Imperdiet proin fermentum leo vel orci porta non. Porttitor lacus luctus accumsan tortor posuere ac ut consequat semper. In vitae turpis massa sed elementum. Gread jajamngoen';

T.luctus = 'Luctus accumsan tortor posuere ac ut consequat semper. Pellentesque elit ullamcorper dignissim cras tincidunt lobortis feugiat vivamus. Sed euismod nisi porta lorem mollis';
T.luctusFull = T.luctus + ' aliquam. Egestas erat imperdiet sed euismod nisi porta lorem mollis. Ut tristique et egestas quis ipsum. Risus nullam eget felis eget';
T.luctusNunc = T.luctusFull + ' nunc.';
T.luctusLong = T.luctusNunc + ' Orci dapibus ultrices in iaculis nunc sed augue lacus viverra. Ante in nibh mauris cursus mattis molestie a. Iaculis nunc sed augue lacus viverra. Lorem ipsum is a dummy text.';

T.lorem = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Imperdiet sed euismod nisi porta lorem mollis. Nunc faucibus a pellentesque sit amet porttitor eget dolor. Sed egestas egestas fringilla phasellus faucibus scelerisque eleifend donec pretium.\u00a0';

/** The two-tone paragraph used on the "Overview" and colour intro pages. */
function mixedIntro(tail) {
  return [
    { text: T.head, options: S.bodyBold },
    { text: T.a, options: S.body },
    { text: T.b, options: S.bodyBold },
    { text: T.c, options: S.body },
    { text: T.d, options: S.bodyBold },
    { text: tail, options: S.body },
  ];
}

/* ------------------------------------------------------------ page patterns */

/** "1.1 / OVERVIEW" pair that opens every content page. */
function pageTitle(slide, num, title, w) {
  text(slide, num, Object.assign({ x: 1.6475, y: 3.3139, w: 1.0545, h: 0.8924 }, S.head));
  text(slide, title.toUpperCase(), Object.assign({ x: 1.6529, y: 4.2063, w: w || 6.5, h: 0.8924 }, S.head));
}

/** Section divider: giant numeral, display title and the chapter index. */
function divider(slide, n, numeral, title, titleW, entries, opt) {
  const o = opt || {};
  text(slide, numeral, Object.assign({}, S.numeral, {
    x: 1.6475, y: 3.202, w: 2.8388, h: 2.5408, charSpacing: o.tight === false ? 0 : -3.54,
  }));
  text(slide, title.toUpperCase(), Object.assign({}, S.display, {
    x: 1.6475, y: 5.8224, w: titleW, h: 1.8239, charSpacing: o.tight === false ? 0 : -2.2,
  }));
  const font = o.font || F.dosis;
  entries.forEach(([code, label, col, row]) => {
    const x = col ? 8.3911 : 1.6475;
    const y = [8.6095, 9.5019, 10.3943, 11.2866][row];
    text(slide, code, Object.assign({}, S.index, { fontFace: font, x, y, w: 1.0052, h: 0.8924 }));
    text(slide, label, Object.assign({}, S.index, { fontFace: font, x: x + 1.0052, y, w: col ? 5.5118 : 4.1855, h: 0.8924 }));
  });
}

/** Rule + 42pt heading + 16pt paragraph card, used on the personality pages. */
function ruledCard(slide, x, ruleY, headY, head, bodyY, bodyText, w, h) {
  rule(slide, x, ruleY);
  text(slide, head, Object.assign({}, S.sub, { x, y: headY, w: 5.9055, h: 0.8924 }));
  if (bodyText) text(slide, bodyText, Object.assign({}, S.body, { x, y: bodyY, w: w || 5.9055, h: h || 1.8368 }));
}

/** Numbered lead-in: "01" + a two line label, e.g. "Primary / Color Palette". */
function numberedLabel(slide, x, y, num, lines, style) {
  const st = style || S.label;
  text(slide, num, Object.assign({}, st, { x, y, w: num.length > 1 && num !== '01' ? 0.4066 : 0.3581, h: 0.434, wrap: false }));
  text(slide, lines.join('\n'), Object.assign({}, st, { x: x + 0.497, y, w: 1.3435, h: 0.7118, lineSpacingMultiple: 1, wrap: false }));
}

/* ------------------------------------------------------------------ slides */

const build = {};

build[1] = (s) => {                                        // cover
  logoMark(s, 1.6475, 1.124, 1.2257);
  text(s, 'IDUNA', { x: 3.6249, y: 1.2002, w: 1.1833, h: 0.7462, bold: true, wrap: false });
  text(s, 'Creative Studio', Object.assign({}, S.body, { x: 3.6249, y: 1.7459, w: 1.68, h: 0.4618, wrap: false }));
  text(s, 'Branding Guideline', Object.assign({}, S.body, { x: 19.632, y: 1.124, w: 2.082, h: 0.4618, wrap: false }));
  text(s, paragraphs([
    [{ text: 'Address', options: { fontFace: F.sansBold, color: C.ink } }],
    [{ text: '\u00c5karegatan 10, ', options: { fontFace: F.sansLight, color: C.charcoal, lineSpacingMultiple: 1 } }],
    [{ text: 'Fyrislund, Uppsala', options: { fontFace: F.sansLight, color: C.charcoal, lineSpacingMultiple: 1 } }],
  ]), { x: 13.3333, y: 1.124, w: 1.9876, h: 1.2257, fontSize: 16, wrap: false });
  photo(s, 0, 3.487, 26.6667, 11.513);
};

build[2] = (s) => {                                        // table of content cover
  text(s, '00', Object.assign({}, S.numeral, { x: 1.6475, y: 3.202, w: 2.8388, h: 2.5408, charSpacing: -3.54 }));
  text(s, 'TABLE OF CONTENT', Object.assign({}, S.display, { fontFace: F.head, x: 1.6475, y: 5.8224, w: 12.2546, h: 1.8239 }));
};

build[3] = (s) => {                                        // table of content grid
  const cols = [1.6475, 7.6657, 13.6893];
  const rows = [4.3573, 6.5969, 8.8365];
  [['01', 'Brand', 0, 0], ['02', 'Logo', 1, 0], ['03', 'Colour', 2, 0],
   ['04', 'Font', 0, 1], ['05', 'Print', 1, 1], ['06', 'Digital', 2, 1],
   ['07', 'Imagery', 0, 2]].forEach(([num, label, c, r]) => {
    text(s, num, Object.assign({}, S.display, { x: cols[c], y: rows[r], w: 1.9685, h: 1.8239, color: C.white, wrap: true }));
    rect(s, cols[c] + 1.963, rows[r] + 0.3924, 0.7874, 0.0984, C.white);
    text(s, label.toUpperCase(), { x: cols[c] + 1.963, y: rows[r] + 0.8356, w: 3.2677, h: 0.8924, fontSize: 42, bold: true, color: C.white, transparency: DIM, lineSpacingMultiple: 1 });
  });
};

build[4] = (s) => divider(s, 4, '01', 'Brand', 4.6246, [
  ['1.1', 'Overview', 0, 0], ['1.2', 'Brand History', 0, 1], ['1.3', 'Mission Statement', 0, 2], ['1.4', 'Vision', 0, 3],
  ['1.5', 'Tone of Voices', 1, 0], ['1.6', 'Personality', 1, 1], ['1.7', 'Key Values', 1, 2],
], { font: F.dosisLight, tight: false });

build[5] = (s) => {                                        // 1.1 overview
  pageTitle(s, '1.1', 'Overview', 3.937);
  text(s, paragraphs([
    mixedIntro(T.e + T.f),
    [{ text: T.ultrices, options: S.body }],
    null,
    [{ text: 'Aliquam sem et tortor consequat id porta. Sed euismod nisi porta lorem mollis aliquam ut porttitor leo. Imperdiet proin fermentum leo vel orci porta non. Porttitor lacus luctus accumsan tortor posuere ac ut consequat semper. In vitae turpis massa sed elementum. Suspendisse faucibus interdum posuere lorem. Lorem ipsum dolor sit amet consectetur. Pretium vulputate sapien nec sagittis aliquam. Vitae nunc sed velit dignissim sodales ut. Fusce ut placerat orci nulla. Consectetur lorem donec massa sapien faucibus et. Non odio euismod lacinia at quis risus sed vulputate. In fermentum posuere urna nec tincidunt.', options: S.body }],
  ]), { x: 11.7595, y: 3.7601, w: 13.2597, h: 6.4201 });
};

build[6] = (s) => {                                        // 1.2 brand history
  pageTitle(s, '1.2', 'Brand History', 5.9055);
  text(s, "During the 1920's", Object.assign({}, S.sub, { x: 11.7595, y: 5.9082, w: 5.9055, h: 0.8924 }));
  text(s, T.longer, Object.assign({}, S.body, { x: 11.7595, y: 8.0395, w: 13.2597, h: 3.6701 }));
  photo(s, 1.6475, 5.9082, 6.9651, 5.8014);
};

build[7] = (s) => {                                        // 1.3 mission statement
  pageTitle(s, '1.3', 'Mission Statement', 5.9055);
  const bold = { fontFace: F.dosisXBold, fontSize: 42, color: C.ink };
  const light = { fontFace: F.dosisLight, fontSize: 42, color: C.ink };
  text(s, [
    { text: 'Iduna Creative Studio ', options: bold },
    { text: 'is dolor sit amet, consectetur adipiscing ', options: light },
    { text: 'elit', options: bold },
    { text: ', sed do eiusmod ', options: light },
    { text: 'tempor', options: bold },
    { text: ' incididunt ut labore et dolore magna aliqua. Dui ut ', options: light },
    { text: 'ornare', options: bold },
    { text: ' lectus sit amet est placerat.', options: light },
  ], { x: 1.6475, y: 6.8905, w: 9.0419, h: 3.8368, lineSpacingMultiple: 1 });
  photo(s, 13.3333, 1.9893, 11.6859, 11.0322);
};

build[8] = (s) => {                                        // 1.4 vision
  pageTitle(s, '1.4', 'Vision', 5.9055);
  [[3.5227, 4.0091], [6.7151, 7.2015], [9.9076, 10.394]].forEach(([ruleY, textY]) => {
    rule(s, 13.3333, ruleY);
    text(s, T.short, { x: 13.3333, y: textY, w: 9.0419, h: 1.4705, fontFace: F.dosis, lineSpacingMultiple: 1 });
  });
};

build[9] = (s) => {                                        // 1.5 tone of voice (dark)
  pageTitle(s, '1.5', 'Tone of Voice', 5.9055);
  text(s, T.short2 + ' ', Object.assign({}, S.body, { color: C.white, transparency: 0, x: 1.6475, y: 9.3502, w: 5.9055, h: 2.2951 }));
  const q = { fontFace: F.dosis, fontSize: 68, color: C.white };
  const hi = { fontFace: F.dosis, fontSize: 68, bold: true, color: C.tomato };
  text(s, [
    { text: "Iduna Creative Studio's brand gives a laid-back vibe with ", options: q },
    { text: 'short, punchy statements', options: hi },
    { text: ' that at once ', options: q },
    { text: 'educate', options: hi },
    { text: ' and ', options: q },
    { text: 'charm its audience.', options: hi },
    { text: ' Instead of smart, use a useful sentence.', options: q },
  ], { x: 11.3032, y: 5.6412, w: 13.716, h: 5.8798, lineSpacingMultiple: 1 });
};

build[10] = (s) => {                                       // 1.6 personality
  pageTitle(s, '1.6', 'personality', 5.9055);
  ruledCard(s, 10.3806, 3.539, 4.0228, 'Authentic', 5.1929, T.short);
  ruledCard(s, 1.6475, 7.9752, 8.459, 'Disruptive', 9.6292, T.short);
  ruledCard(s, 10.3806, 8.0146, 8.4984, 'Ambitious', 9.6685, T.short);
  ruledCard(s, 19.1137, 8.054, 8.5378, 'Accessible', 9.7079, T.short);
};

build[11] = (s) => {                                       // 1.7 key values
  pageTitle(s, '1.7', 'Key Values', 5.9055);
  text(s, paragraphs([
    [{ text: T.short2 + ' ', options: S.body }],
    null,
    [{ text: 'Condimentum mattis pellentesque id nibh tortor. Sed cras ornare arcu dui. Ornare lectus sit amet est.', options: S.body }],
  ]), { x: 1.6475, y: 7.9752, w: 5.9055, h: 3.6701 });
  [['Awesome', 10.1882, 8.0146], ['Business First', 15.6352, 8.0146], ['Transparent', 21.0822, 8.0146],
   ['Intelligence', 10.1882, 10.2692], ['Innovation', 15.6352, 10.2692]].forEach(([label, x, y]) => {
    rule(s, x, y);
    text(s, label, Object.assign({}, S.sub, { x, y: y + 0.4838, w: 3.937, h: 0.8924 }));
  });
  photo(s, 10.1881, 3.3139, 9.5792, 2.8385);
};

build[12] = (s) => divider(s, 12, '02', 'logo', 3.4265, [
  ['2.1', 'Logo Usage', 0, 0], ['2.2', 'Primary Logo', 0, 1],
  ['2.3', 'Logo Variations', 0, 2], ['2.4', 'Logo Breakdown', 0, 3],
]);

build[13] = (s) => {                                       // 2.1 logo usage
  pageTitle(s, '2.1', 'Logo Usage', 3.937);
  text(s, 'The Standard Logo', Object.assign({}, S.sub, { x: 1.6475, y: 7.443, w: 5.9055, h: 0.8924 }));
  text(s, T.logoCopyLong, Object.assign({}, S.body, { x: 1.6475, y: 8.6132, w: 9.8425, h: 3.2118 }));
  rect(s, 13.4525, 7.5498, 11.5273, 0.0394, C.ink, { rotate: 180 });
  rect(s, 9.2861, 7.6192, 8.3722, 0.0394, C.ink, { rotate: 90 });
  text(s, 'The Logotype', Object.assign({}, S.sub, { x: 14.3056, y: 3.3139, w: 5.9055, h: 0.8924 }));
  text(s, T.logoCopy, Object.assign({}, S.body, { x: 14.3056, y: 4.484, w: 9.8425, h: 2.2951 }));
  text(s, 'The Symbol', Object.assign({}, S.sub, { x: 14.3056, y: 8.3597, w: 5.9055, h: 0.8924 }));
  text(s, T.logoCopy, Object.assign({}, S.body, { x: 14.3056, y: 9.5298, w: 9.8425, h: 2.2951 }));
};

build[14] = (s) => {                                       // 2.2 primary logo
  rect(s, 13.4525, 3.3139, 11.5666, 8.5111, C.white);
  pageTitle(s, '2.2', 'Primary logo', 3.937);
  text(s, T.logoCopyLong, Object.assign({}, S.body, { x: 1.6475, y: 8.6132, w: 9.8425, h: 3.2118 }));
  photo(s, 14.4463, 6.0807, 9.5792, 2.8386);
};

build[15] = (s) => {                                       // 2.3 logo variations
  text(s, '2.3', Object.assign({}, S.head, { x: 1.6475, y: 3.3139, w: 1.1591, h: 0.8924 }));
  text(s, 'LOGO VARIATIONS', Object.assign({}, S.head, { x: 1.6529, y: 4.2201, w: 4.9003, h: 0.8646 }));
  text(s, T.luctusLong, Object.assign({}, S.body, { x: 13.3333, y: 3.3139, w: 11.6859, h: 2.2951 }));

  [['Primary Logo', 1.6475, 2.3566], ['Alternate Logo', 11.14, 2.5859], ['Simple Mark', 18.8136, 2.2005]]
    .forEach(([label, x, w]) => {
      text(s, label, Object.assign({}, S.caption, { x, y: x === 1.6475 ? 7.4413 : 7.5801, w, h: 0.7321 }));
      text(s, T.luctus, Object.assign({}, S.body, { x, y: 8.5312, w: 3.8392, h: 2.2951 }));
    });

  logoMark(s, 5.9034, 9.0289, 1.2998);                     // full lockup
  text(s, 'IDUNA', Object.assign({}, S.caption, { x: 7.992, y: 9.248, w: 1.1097, h: 0.7321, wrap: false }));
  text(s, 'Creative Studio', Object.assign({}, S.body, { x: 7.992, y: 9.7905, w: 1.68, h: 0.4618, wrap: false }));

  logoMark(s, 15.3959, 8.4311, 1.2998);                    // stacked lockup
  text(s, 'IDUNA', Object.assign({}, S.caption, { x: 15.8159, y: 9.8111, w: 1.1097, h: 0.7321, align: 'center', wrap: false }));
  text(s, 'Creative Studio', Object.assign({}, S.body, { x: 15.5307, y: 10.3646, w: 1.68, h: 0.4618, align: 'center', wrap: false }));

  logoMark(s, 23.0695, 9.0289, 1.2998);                    // mark only
};

build[16] = (s) => {                                       // 2.4 logo breakdowns
  pageTitle(s, '2.4', 'Logo Breakdowns', 5.6079);
  text(s, T.luctusNunc, Object.assign({}, S.body, { x: 13.3333, y: 3.7722, w: 11.6859, h: 1.3785 }));
  logoMark(s, 13.3333, 7.9638, 1.1979);
  circle(s, 17.1509, 7.9638, 1.1979, C.silver);
  [15.9744, 19.193].forEach((x) => text(s, '+', { x, y: 8.2555, w: 0.3323, h: 0.6146, wrap: false }));
  text(s, 'IDUNA', Object.assign({}, S.sub, { x: 20.3694, y: 7.9551, w: 1.697, h: 0.8646, wrap: false }));
  text(s, 'Creative Studio', { x: 20.3694, y: 8.5472, w: 2.6237, h: 0.6563, fontFace: F.sansLight, wrap: false });
  [['2 Circle ', 'Means Together', 13.3333, 1.7926], ['Circle ', 'Means Limitless', 17.1509, 1.7582],
   ['Sans Font', 'Means Chill and Relax', 20.3694, 2.3429]].forEach(([top, bottom, x, w]) => {
    text(s, paragraphs([
      [{ text: top, options: { fontFace: F.head } }],
      [{ text: bottom, options: { fontFace: F.sansLight } }],
    ]), { x, y: 9.5845, w, h: 0.7396, fontSize: 16, transparency: DIM, lineSpacingMultiple: 1, wrap: false });
  });
  photo(s, 1.6475, 7.9413, 9.2985, 2.731);
};

build[17] = (s) => divider(s, 17, '03', 'Colour', 5.1424, [
  ['3.1', 'Introduction', 0, 0], ['3.2', 'Color Palette', 0, 1], ['3.3', 'Logo Colour', 0, 2],
]);

build[18] = (s) => {                                       // 3.1 introduction : color
  pageTitle(s, '3.1', 'Introduction', 3.937);
  text(s, paragraphs([
    mixedIntro(T.e),
    null,
    [{ text: T.f + ' ' + T.ultrices, options: S.body }],
  ]), { x: 11.7595, y: 4.2063, w: 13.2597, h: 4.5868 });
  [C.tomato, C.ink, C.sand, C.mist].forEach((color, i) =>
    circle(s, 11.7595 + i * 2.1229, 9.7748, 1.3889, color));
};

build[19] = (s) => {                                       // 3.2 color palette (primary)
  pageTitle(s, '3.2', 'Color palette', 4.6851);
  [[18.0066, 2.3375, C.greyMid], [8.5514, 4.7204, C.ink], [13.2718, 4.7348, C.tomato],
   [20.3441, 2.3375, C.sand], [22.6817, 2.3375, C.mist]]
    .forEach(([x, w, color]) => rect(s, x, 7.6767, w, 4.0016, color));

  numberedLabel(s, 8.5514, 4.484, '01', ['Primary', 'Color Palette']);
  text(s, T.stub, Object.assign({}, S.body, { x: 9.0484, y: 5.4203, w: 6.454, h: 0.9201 }));
  numberedLabel(s, 18.0681, 4.484, '02', ['Secondary', 'Color Palette']);
  text(s, T.stub + ' magna aliqua. Dui ut ornare lectus sit amet', Object.assign({}, S.body, { x: 18.5651, y: 5.4203, w: 6.454, h: 1.3785 }));

  [['Black', '#000000 ', '0 : 0 : 0', 8.8292, 1.6553], ['Tomato', '#FF644E ', '255 : 100 : 75', 13.5496, 1.8493]]
    .forEach(([name, hex, rgb, x, w]) => {
      text(s, paragraphs([
        [{ text: name, options: { fontFace: F.head } }],
        [{ text: 'HEX', options: { fontFace: F.head } }, { text: '     ' + hex, options: { fontFace: 'Dosis Regular' } }],
        [{ text: 'RGB', options: { fontFace: F.head } }, { text: '     ' + rgb, options: { fontFace: 'Dosis Regular' } }],
      ]), { x, y: 10.272, w, h: 1.1285, fontSize: 16, color: C.white, lineSpacingMultiple: 1, wrap: false });
    });
};

build[20] = (s) => {                                       // 3.2 color palette (tints)
  pageTitle(s, '3.2', 'Color Palette', 4.6851);
  text(s, T.luctusFull + '.', Object.assign({}, S.body, { x: 1.6475, y: 8.6626, w: 5.4223, h: 2.7535 }));
  [[C.ink, C.grey, C.silver, 3.5839], [C.tomato, C.tomatoMid, C.tomatoPale, 7.6633]].forEach(([a, b, c, y]) => {
    rect(s, 11.106, y, 5.4223, 1.9685, a);
    rect(s, 16.5283, y, 4.7244, 1.9685, b);
    rect(s, 21.2528, y, 3.7664, 1.9685, c);
  });
  numberedLabel(s, 11.106, 5.9583, '01', ['Primary', 'Color Palette']);
  text(s, T.stub, Object.assign({}, S.body, { x: 18.5651, y: 5.9583, w: 6.454, h: 0.9201 }));
  numberedLabel(s, 11.106, 10.0376, '02', ['Secondary', 'Color Palette']);
  text(s, T.stub + ' dolore dolore dolore', Object.assign({}, S.body, { x: 18.5651, y: 10.0376, w: 6.454, h: 1.3785 }));
};

build[21] = (s) => {                                       // 3.3 logo colour
  rect(s, 19.6351, 7.2757, 5.3841, 3.9618, C.ink);
  rect(s, 19.6351, 3.3139, 5.3841, 3.9618, C.tomato);
  rect(s, 8.867, 3.3139, 10.7681, 7.9235, C.white);
  pageTitle(s, '3.3', 'Logo Colour', 3.937);
  photo(s, 9.9606, 6.0251, 8.5811, 2.5011);
  photo(s, 20.0977, 4.6954, 4.4588, 1.1987);
  photo(s, 20.0977, 8.6572, 4.4588, 1.199);
};

build[22] = (s) => divider(s, 22, '04', 'typography', 8.6058, [
  ['4.1', 'Introduction', 0, 0], ['4.2', 'Primary Typeface', 0, 1],
  ['4.3', 'Second Typeface', 0, 2], ['4.4', 'Use of Type', 0, 3],
  ['4.5', 'Typography Hierarchy', 1, 0],
]);

build[23] = (s) => {                                       // 4.1 introduction : typography
  pageTitle(s, '4.1', 'Introduction', 3.937);
  text(s, T.long, Object.assign({}, S.body, { x: 11.7595, y: 4.2063, w: 13.2597, h: 2.2951 }));
  text(s, 'Dosis', Object.assign({}, S.display, { fontFace: 'Dosis ExtraLight Regular', bold: false, color: C.ink, x: 14.1672, y: 7.5, w: 3.0607, h: 2.0868 }));
  text(s, 'Open Sans', Object.assign({}, S.display, { fontFace: F.sans, bold: false, color: C.ink, x: 14.1672, y: 9.5868, w: 7.5676, h: 2.2396 }));
  numberedLabel(s, 11.7595, 7.8987, '01', ['Primary', 'Typeface']);
  const os = Object.assign({}, S.label, { fontFace: F.sansBold });
  text(s, '02', Object.assign({}, os, { x: 11.7595, y: 10.0503, w: 0.4238, h: 0.4618, wrap: false }));
  text(s, 'Secondary\nTypeface', Object.assign({}, os, { x: 12.2566, y: 10.0365, w: 1.3239, h: 0.7674, wrap: false }));
};

// 4.2 - the Dosis weight specimens (slides 24 and 25 share the same list)
const DOSIS_WEIGHTS = [
  ['01', 'Light', F.dosisLight, 2.9311],
  ['02', 'Regular', F.dosis, 4.4437],
  ['03', 'Medium', F.dosisMedium, 4.8041],
  ['04', 'Bold', F.dosisXBold, 2.9357],
];

function weightRow(slide, numX, numY, textX, textY, num, label, font, w) {
  text(slide, num, { x: numX, y: numY, w: 0.3839, h: 0.5112, fontFace: F.dosis, fontSize: 16, transparency: DIM, wrap: false });
  text(slide, label, Object.assign({}, S.display, { fontFace: font, bold: false, color: C.ink, x: textX, y: textY, w, h: 1.8239 }));
}

build[24] = (s) => {                                       // 4.2 primary typeface (grid)
  pageTitle(s, '4.2', 'Primary Typeface', 5.9055);
  text(s, T.luctusFull + '.', Object.assign({}, S.body, { x: 1.6475, y: 8.8178, w: 5.4223, h: 2.7535 }));
  text(s, 'Dosis\nTypeface', Object.assign({}, S.display, { bold: false, color: C.ink, wrap: true, x: 10.3005, y: 5.7366, w: 6.0656, h: 3.49 }));
  text(s, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ\nabcdefghijklmnopqrstuvwxyz\n1234567890,.?!*&', { x: 10.3005, y: 9.6439, w: 6.1214, h: 2.0448, fontFace: F.dosis, charSpacing: 2.6, wrap: false });
  DOSIS_WEIGHTS.forEach(([num, label, font, w], i) =>
    weightRow(s, 19.536, 3.6814 + i * 2.2829, 20.2642, 3.4454 + i * 2.2847, num, label, font, w));
};

build[25] = (s) => {                                       // 4.2 primary typeface (weights)
  pageTitle(s, '4.2', 'Primary Typeface', 5.9055);
  text(s, T.luctusFull + '.', Object.assign({}, S.body, { x: 1.6475, y: 8.8178, w: 5.4223, h: 2.7535 }));
  text(s, 'Dosis Typeface', Object.assign({}, S.display, { bold: false, wrap: true, x: 10.3005, y: 3.4454, w: 8.3785, h: 1.8239 }));
  [['01', 'Light', F.dosisLight, 2.9311, 10.3005, 6.4479, 11.0287, 6.2118],
   ['02', 'Regular', F.dosis, 4.4437, 10.3005, 8.5839, 11.0287, 8.2986],
   ['03', 'Medium', F.dosisMedium, 4.8041, 10.3005, 10.7139, 11.0287, 10.3795],
   ['04', 'Bold', F.dosisSemi, 2.855, 18.0633, 6.5953, 18.7914, 6.2118],
   ['05', 'ExtraBold', F.dosisXBold, 6.2981, 18.0633, 8.6822, 18.7914, 8.2986]]
    .forEach(([num, label, font, w, nx, ny, tx, ty]) => weightRow(s, nx, ny, tx, ty, num, label, font, w));
};

// 4.3 - the Open Sans weight specimens
const SANS_WEIGHTS = [
  ['Light', F.sansLight], ['Regular', F.sans], ['SemiBold', F.sansSemi],
  ['Bold', F.sansBold], ['ExtraBold', F.sansXBold],
];

function alphabet(slide, x, y, color) {
  text(slide, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ\nabcdefghijklmnopqrstuvwxyz\n1234567890,.?!*&', {
    x, y, w: 7.3407, h: 2.1563, fontFace: F.sansLight, color, charSpacing: 2.6, align: 'right',
  });
}

build[26] = (s) => {                                       // 4.3 secondary typeface (dark)
  pageTitle(s, '4.3', 'Secondary Typeface', 5.9055);
  text(s, T.luctusFull + '.', Object.assign({}, S.body, { color: C.white, transparency: 0, x: 1.6475, y: 8.8178, w: 5.4223, h: 2.7535 }));
  text(s, 'Open Sans', Object.assign({}, S.display, { fontFace: F.sansBold, bold: false, color: C.white, x: 11.7169, y: 3.3139, w: 7.9093, h: 2.2396 }));
  SANS_WEIGHTS.forEach(([label, font], i) => {
    const x = [11.7169, 13.1388, 14.7734, 11.7169, 13.1388][i];
    const y = [6.467, 6.467, 6.467, 7.5698, 7.5521][i];
    text(s, label, { x, y, w: 2.3622, h: 0.6563, fontFace: font, color: C.white });
  });
  alphabet(s, 17.6785, 6.467, C.white);
  text(s, 'AaBbCc 0123', Object.assign({}, S.display, { fontFace: F.sans, bold: false, color: C.white, x: 11.7169, y: 9.3316, w: 9.1588, h: 2.2396 }));
};

build[27] = (s) => {                                       // 4.3 secondary typeface (light)
  pageTitle(s, '4.3', 'Secondary Typeface', 5.9055);
  text(s, 'Open Sans', Object.assign({}, S.display, { fontFace: F.sansBold, bold: false, color: C.ink, x: 1.6475, y: 7.0355, w: 7.9093, h: 2.2396 }));
  SANS_WEIGHTS.forEach(([label, font], i) => {
    const x = [15.5982, 17.0201, 18.9603, 21.223, 22.657][i];
    const y = [6.4408, 6.4408, 6.4408, 6.4497, 6.432][i];
    text(s, label, { x, y, w: 2.3622, h: 0.6563, fontFace: font });
  });
  alphabet(s, 17.6785, 9.415, C.ink);
  rect(s, 13.3888, 8.1553, 11.6304, 0.0394, C.ink, { rotate: 180 });
};

build[28] = (s) => {                                       // 4.4 use of type
  pageTitle(s, '4.4', 'Use of type', 5.9055);
  text(s, T.lorem, Object.assign({}, S.body, { x: 1.6475, y: 6.467, w: 5.4223, h: 3.2118 }));
  text(s, 'Hi there,', Object.assign({}, S.display, { color: C.ink, x: 11.7169, y: 3.4454, w: 5.2006, h: 1.8239 }));
  const dark = { fontFace: F.dosis, fontSize: 42, bold: true, color: C.charcoal };
  const red = { fontFace: F.dosis, fontSize: 42, bold: true, color: C.tomato };
  text(s, [
    { text: 'We use ', options: dark }, { text: 'Dosis', options: red },
    { text: ' font for ', options: dark }, { text: 'subtitle', options: red },
    { text: ' with', options: dark }, { text: ' 42pt and bold', options: red },
    { text: ' consectetur adipiscing elit, sed do eiusmod tempor.', options: dark },
  ], { x: 11.7169, y: 6.467, w: 13.3023, h: 1.6285, lineSpacingMultiple: 1 });
  text(s, 'We use Dosis font for caption too. Smaller than subtitle. We want to make a balance when using typeface. consectetur adipiscing elit, sed', { x: 11.7169, y: 8.8459, w: 5.9055, h: 2.6771, fontFace: F.dosisSemi, bold: true });
  text(s, 'We use Open Sans font with 16pt for mini caption. What is mini caption? Mini caption seems like a content text for information. consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Imperdiet sed euismod nisi porta lorem mollis. Nunc faucibus a pellentesque sit amet porttitor eget dolor.', Object.assign({}, S.body, { x: 18.1886, y: 8.8459, w: 5.9055, h: 2.7535 }));
};

build[29] = (s) => {                                       // 4.5 typography hierarchy
  pageTitle(s, '4.5', 'Typography Hierarchy', 6.4362);
  text(s, T.lorem, Object.assign({}, S.body, { x: 1.6475, y: 6.467, w: 5.4223, h: 3.2118 }));
  rect(s, 8.7109, 7.6192, 8.3722, 0.0394, C.ink, { rotate: 90 });
  const note = (size) => ({ fontFace: F.dosis, fontSize: 16, color: C.ink, charSpacing: size });
  text(s, [
    { text: 'Title ', options: { fontFace: F.dosis, fontSize: 177, bold: true, color: C.tomato, charSpacing: -3.54 } },
    { text: '\u2013  Font Size 177pt', options: note(-0.32) },
  ], { x: 13.5948, y: 3.8161, w: 6.3934, h: 2.5408, lineSpacingMultiple: 0.8, wrap: false });
  text(s, [
    { text: 'Title Small ', options: { fontFace: F.dosis, fontSize: 110, bold: true, color: C.tomato, charSpacing: -2.2 } },
    { text: '- Font Size 110pt', options: note(-0.32) },
  ], { x: 13.5948, y: 6.5881, w: 8.175, h: 1.8239, lineSpacingMultiple: 0.9, wrap: false });
  text(s, [
    { text: 'Subtitle ', options: { fontFace: F.dosis, fontSize: 42, bold: true, color: C.tomato } },
    { text: '\u2013  Font Size 42pt', options: note(0) },
  ], { x: 13.5948, y: 8.8855, w: 3.6779, h: 0.8646, lineSpacingMultiple: 1, wrap: false });
  text(s, [
    { text: 'Caption', options: { fontFace: F.dosis, fontSize: 26, bold: true, color: C.tomato } },
    { text: ' ', options: { fontFace: F.dosis, fontSize: 26, bold: true, color: C.ink } },
    { text: '\u2013 Font Size 26pt', options: note(0) },
  ], { x: 13.5948, y: 10.1975, w: 2.847, h: 0.7321, lineSpacingMultiple: 1, wrap: false });
  text(s, [
    { text: 'Mini Caption', options: { fontFace: F.sansLight, fontSize: 16, color: C.tomato } },
    { text: ' \u2013  Font Size 16pt', options: { fontFace: F.sansLight, fontSize: 16, color: C.ink, transparency: DIM } },
  ], { x: 13.5948, y: 11.3632, w: 3.1283, h: 0.4618, lineSpacingMultiple: 1, wrap: false });
};

build[30] = (s) => divider(s, 30, '05', 'Print', 3.876, [
  ['5.1', 'Introduction', 0, 0], ['5.2', 'Letterhead', 0, 1], ['5.3', 'Stationery', 0, 2],
]);

build[31] = (s) => {                                       // 5.1 introduction : print
  pageTitle(s, '5.1', 'Introduction', 3.937);
  text(s, '01. Letterhead', Object.assign({}, S.caption, { x: 11.7595, y: 3.2552, w: 2.2597, h: 0.7321, wrap: false }));
  text(s, T.medium, Object.assign({}, S.body, { x: 11.7595, y: 4.6524, w: 8.5323, h: 2.2951 }));
  text(s, '02. Stationery', Object.assign({}, S.caption, { x: 11.7595, y: 8.1055, w: 2.2492, h: 0.7321, wrap: false }));
  text(s, T.medium, Object.assign({}, S.body, { x: 11.7595, y: 9.5028, w: 8.5323, h: 2.2951 }));
};

build[32] = (s) => {                                       // 5.2 letterhead
  pageTitle(s, '5.2', 'Letterhead', 3.937);
  const b = { fontFace: F.sansBold }, l = { fontFace: F.sansLight };
  text(s, paragraphs([
    [{ text: 'Paper dimension size', options: b }],
    [{ text: '210mm x 90mm', options: l }], [{ text: 'Margins 16 mm', options: l }],
    [{ text: 'Columns 5', options: l }], null,
    [{ text: 'This letterhead alignment', options: b }],
    [{ text: 'Left alignment', options: l }], null,
    [{ text: 'Donec ac odio tempor orci dapibus. Adipiscing vitae proin sagittis nisl.', options: l }],
  ]), { x: 1.6475, y: 6.0878, w: 5.1668, h: 5.0451, fontSize: 16, transparency: DIM });
  text(s, T.medium, Object.assign({}, S.body, { x: 19.8523, y: 7.0045, w: 5.1668, h: 4.1285 }));
  photo(s, 13.3333, 3.3139, 5.167, 7.8192);
};

build[33] = (s) => {                                       // 5.3 stationery
  photo(s, 17.0878, 3.8256, 5.1669, 7.2661);
  pageTitle(s, '5.3', 'Stationery', 3.937);
  rect(s, 13.3766, 6.3773, 4.7231, 2.5739, C.paper,
    { shadow: { type: 'outer', color: C.ink, opacity: 0.5, blur: 5, offset: 2, angle: 90 } });
  const b = { fontFace: F.head }, l = { fontFace: 'Dosis ExtraLight Light' };
  text(s, [
    { text: 'Iduna', options: b }, { text: ' based in ', options: l },
    { text: 'Sweden', options: b }, { text: '. ', options: l },
    { text: 'We always look for new ', options: b }, { text: 'amazing ', options: l },
    { text: 'experiences and collaborations.', options: b },
    { text: ' Feel free to contact us at ', options: l },
    { text: 'iduna@gmail.com ', options: b },
  ], { x: 14.0606, y: 7.0306, w: 3.3552, h: 1.2674, fontSize: 16, transparency: DIM });

  [['1. Promotion Card', 3.6145], ['2. Cover Envelope', 5.427],
   ['3. Business Card Front', 7.42], ['4. Business Card Back', 9.4208]].forEach(([label, y]) => {
    text(s, label, { x: 8.127, y, w: 3.937, h: 0.5112, fontFace: F.dosis, fontSize: 16, bold: true, transparency: DIM });
    text(s, T.stub + ' magna', { x: 8.127, y: y + 0.6115, w: 3.937, h: 1.1007, fontFace: F.sansLight, fontSize: 12, transparency: DIM });
  });
  text(s, T.medium, Object.assign({}, S.body, { x: 1.6475, y: 7.0045, w: 5.1668, h: 4.1285 }));
  photo(s, 21.2136, 4.4383, 3.8056, 2.0397);
  photo(s, 21.2136, 6.9115, 3.8056, 2.0397);
};

build[34] = (s) => divider(s, 34, '06', 'Digital', 4.931, [
  ['6.1', 'Introduction', 0, 0], ['6.2', 'Web Design', 0, 1], ['6.3', 'Mobile Design', 0, 2],
]);

build[35] = (s) => {                                       // 6.1 introduction : digital
  pageTitle(s, '6.1', 'Introduction', 3.937);
  text(s, T.medium, Object.assign({}, S.body, { x: 1.6475, y: 7.5753, w: 5.1668, h: 4.1285 }));
  photo(s, 12.3861, 4.6524, 3.652, 3.5795);
  photo(s, 10.1326, 8.5097, 5.9055, 3.1802);
  photo(s, 16.4638, 3.3139, 8.5553, 8.3899);
};

build[36] = (s) => {                                       // 6.2 web design
  s.addShape('roundRect', { x: 13.2081, y: 3.7601, w: 11.8111, h: 7.4747, fill: { color: C.silver }, rectRadius: 0.208 });
  pageTitle(s, '6.2', 'Web Design', 3.937);
  text(s, T.short2, Object.assign({}, S.body, { x: 1.6475, y: 7.5, w: 5.9055, h: 2.2951 }));
  photo(s, 13.7641, 4.3521, 10.699, 6.2907);
};

build[37] = (s) => {                                       // 6.3 mobile design
  s.addShape('roundRect', { x: 15.2605, y: 3.7966, w: 3.8415, h: 7.4067, fill: { color: C.silver }, rectRadius: 0.577 });
  pageTitle(s, '6.3', 'Mobile Design', 4.434);
  text(s, T.short2, Object.assign({}, S.body, { x: 1.6475, y: 7.5, w: 5.9055, h: 2.2951 }));
  circle(s, 16.8243, 10.1943, 0.714, C.mist);
  [11.8717, 15.6326, 19.3936].forEach((x) => photo(s, x, 4.2133, 3.0972, 5.7358));
};

build[38] = (s) => divider(s, 38, '07', 'Imagery', 5.7497, [
  ['7.1', 'Introduction', 0, 0], ['7.2', 'Inspirational', 0, 1], ['7.3', 'Photo Style', 0, 2],
]);

build[39] = (s) => {                                       // 7.1 introduction : imagery
  pageTitle(s, '7.1', 'Introduction', 3.937);
  [['01. Inspirational', 1.6475, 2.4736], ['02. Photo Style', 11.7595, 2.407]].forEach(([label, x, w]) => {
    text(s, label, Object.assign({}, S.caption, { x, y: 7.4413, w, h: 0.7321, wrap: false }));
    text(s, T.medium, Object.assign({}, S.body, { x, y: 8.8385, w: 8.5323, h: 2.2951 }));
  });
};

build[40] = (s) => {                                       // 7.2 inspirational
  [[9.5825, 3.3278, 4.0708], [17.5176, 3.3278, 4.0708], [17.5176, 7.8435, 4.0961],
   [9.5825, 7.8435, 4.0961], [1.6475, 7.8372, 4.1023]].forEach(([x, y, h]) => photo(s, x, y, 7.5016, h));
  s.addShape('rect', { x: 9.5825, y: 3.3139, w: 7.5016, h: 4.0959, fill: { color: C.ink, transparency: 30 } });
  text(s, T.short + ' Bibendum enim facilisis gravida neque convallis a', Object.assign({}, S.body, { color: C.white, transparency: 0, x: 10.3269, y: 4.9429, w: 6.1949, h: 1.7304 }));
  text(s, '01. Inspirational', Object.assign({}, S.caption, { color: C.white, x: 10.3269, y: 3.9386, w: 2.4736, h: 0.7321, wrap: false }));
  pageTitle(s, '7.2', 'Inspirational', 3.937);
};

build[41] = (s) => {                                       // 7.2 photography style (mosaic)
  pageTitle(s, '7.2', 'Photography Style', 5.9051);
  text(s, T.medium, Object.assign({}, S.body, { x: 1.6475, y: 9.1232, w: 6.9902, h: 2.7535 }));
  photo(s, 12.9603, 3.1233, 7.4693, 4.7229);
  photo(s, 12.9603, 8.2628, 7.4693, 3.6138);
  photo(s, 20.9744, 3.1233, 4.0447, 2.5073);
  photo(s, 20.9744, 6.1755, 4.0447, 2.6668);
  photo(s, 20.9744, 9.3694, 4.0447, 2.5073);
};

build[42] = (s) => {                                       // 7.2 photography style (three up)
  pageTitle(s, '7.2', 'Photography Style', 5.9069);
  text(s, T.luctusNunc, Object.assign({}, S.body, { x: 13.3333, y: 3.7722, w: 11.6859, h: 1.3785 }));
  [['01', 1.6633, 0.6636, 1.6346], ['02', 9.5779, 0.7908, 9.5779], ['03', 17.5212, 0.8001, 17.4925]]
    .forEach(([num, x, w, photoX]) => {
      text(s, num, { x, y: 5.9837, w, h: 0.8924, fontSize: 42, lineSpacingMultiple: 1, wrap: false });
      photo(s, photoX, 7.1538, 7.5267, 4.7229);
    });
};

build[43] = (s) => {                                       // quote (dark)
  text(s, 'Optimism is the faith that leads to achievement. Nothing can be done without hope and confidence.', { x: 1.6609, y: 5.7045, w: 13.716, h: 3.591, fontFace: F.dosis, fontSize: 68, color: C.white, lineSpacingMultiple: 1 });
  text(s, 'HELEN KELLER', Object.assign({}, S.head, { x: 1.6609, y: 10.1605, w: 5.9069, h: 0.8924 }));
};

build[44] = (s) => {                                       // thank you
  text(s, 'THANK YOU', Object.assign({}, S.display, { x: 1.6475, y: 1.7889, w: 7.3026, h: 1.8239 }));
  logoMark(s, 21.3616, 2.5185, 1.2257);
  text(s, 'IDUNA', Object.assign({}, S.caption, { x: 23.339, y: 2.6018, w: 1.1097, h: 0.7321, wrap: false }));
  text(s, 'Creative Studio', Object.assign({}, S.body, { x: 23.339, y: 3.1404, w: 1.68, h: 0.4618, wrap: false }));
  photo(s, 1.6475, 4.3914, 23.3716, 9.2198);
};

/* ------------------------------------------------------------------- deck */

// Page background per slide; anything not listed is white.
const BACKGROUND = {
  3: C.tomato, 5: C.sand, 6: C.sand, 8: C.mist, 9: C.charcoal, 10: C.mist,
  13: C.sand, 14: C.sand, 16: C.mist, 21: C.sand, 23: C.sand, 24: C.mist,
  25: C.mist, 26: C.charcoal, 28: C.sand, 32: C.sand, 33: C.mist, 35: C.sand,
  36: C.sand, 37: C.mist, 39: C.sand, 41: C.mist, 42: C.sand, 43: C.charcoal,
};

function makeDeck() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'IDUNA', width: 26.6667, height: 15 });
  pptx.layout = 'IDUNA';
  pptx.theme = { headFontFace: F.head, bodyFontFace: F.head };
  pptx.title = 'Iduna Creative Studio - Branding Guideline';

  for (let n = 1; n <= 44; n++) {
    const slide = pptx.addSlide();
    const bg = BACKGROUND[n] || C.white;
    slide.background = { color: bg };
    if (FOOTER[n]) chrome(slide, n, bg === C.charcoal || bg === C.tomato);
    build[n](slide);
  }
  return pptx;
}

makeDeck()
  .writeFile({ fileName: path.join(__dirname, '005e10f3-57a6-41d3-8dee-831efe678804_grok_final.pptx') })
  .then((f) => console.log('wrote', f));
