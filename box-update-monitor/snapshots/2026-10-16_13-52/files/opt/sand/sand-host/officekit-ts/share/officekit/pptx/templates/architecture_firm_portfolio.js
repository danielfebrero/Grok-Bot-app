/**
 * "Beyond Walls" — MinimalArchitecture deck (15 slides, 13.333 x 7.5 in).
 * Rebuilt with pptxgenjs only. Photographs in the original are replaced by
 * flat grey "[image]" placeholder rectangles at the same position and size.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const DISPLAY = 'Red Hat Display'; // theme major font
const BODY = 'Open Sans'; // theme minor font

const C = {
  bg: 'DEDEDE', // lt1 / slide background
  ink: '000000', // dk1 / body text
  white: 'FFFFFF',
  orange: 'FF5D38', // accent1
  orangeDim: 'BF462A', // accent1 @ lumMod 75%
  greyBlock: 'C8C8C8', // bg1 @ lumMod 90% — filled panels & hairlines
  greyNum: 'A6A6A6', // bg1 @ lumMod 75% — ghost numerals on slide 4
  photo: 'CECECE', // stand-in for the main stock photo
  photoDark: 'B2B2B2', // stand-in for the darker gallery photo
  photoMid: 'C1C1C1', // stand-in for the mid-grey gallery photo
  caption: '9E9E9E', // "[image]" caption on placeholders
  deviceWhite: 'F5F6F8',
  deviceSilver: 'D2D4D5',
  deviceGrey: 'BFC1C4',
  deviceDark: '0D0D0D',
};

const HAIRLINE = { color: C.greyBlock, width: 0.75 };
// Every card / button in the deck carries the same very soft drop shadow.
const SOFT_SHADOW = { type: 'outer', blur: 20, offset: 15, angle: 45, color: C.ink, opacity: 0.1 };

/* ------------------------------------------------- lorem re-used verbatim */

const LOREM = {
  // slide 3 / 4 / 14 intro paragraph
  intro:
    'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, ' +
    'a maximus leo ligula at dolor. Morbi et',
  // slide 1 / 15 hero paragraph
  hero:
    'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, ' +
    'a maximus leo ligula at dolor. Morbi et malesuada purus. Phasellus a lacus sit amet urna ' +
    'tempor sollicitudin. Cras pretium Quisque non elit mauris. Cras euismod, ',
  // slide 2 table-of-contents paragraph
  toc:
    'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, ' +
    'a maximus leo ligula at dolor. Morbi et malesuada purus. Phasellus a lacus sit amet urna ' +
    'tempor sollicitudin. Cras finibus finibus, felis dui suscipit purus metus ac finibus finibus ' +
    'malesuada purus',
  // slide 7 column paragraph (used four times)
  column:
    'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, ' +
    'a maximus leo ligula at dolor. Morbi et malesuada purus. Phasellus a lacusmetus ac finibus ' +
    'finibus, felis dui suscipit purus finibus finibus, felis dui suscipit purus euismod, metus  metus ac',
  // slide 13 right-hand paragraph (used twice)
  wide:
    'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui susci pit purus, ' +
    'a maximus leo ligula at dolor. Morbi et malesuada purus. Phasellus a lacus sit amet urna ' +
    'tempor sollicitudin. Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui ',
  // slide 6 numbered-card paragraph (used twice)
  card:
    'Quisque non elit mauris. Cras euismod, metus ac finibus fini bus, felis dui suscipit purus, ' +
    'a maximus leo ligula euismod, metus euismod, metus ac finibus ',
  // slide 13 numbered-card paragraph (used twice)
  cardShort:
    'Quisque non elit mauris. Cras euismod, metus ac finibus fini bus, felis dui suscipit purus, ' +
    'a maximus leo ligula euismod, metus',
  tiny: 'Quisque non elit mauris. Cras',
  tinier: 'Quisque non elit mauris. Cras euismod, metus ac',
};

/* --------------------------------------------------------------- helpers */

/** Display-font heading (theme major font, top aligned). */
function heading(slide, text, x, y, w, h, size, opts) {
  slide.addText(text, Object.assign({
    x, y, w, h, fontFace: DISPLAY, fontSize: size, color: C.ink, valign: 'top',
  }, opts));
}

/** Body copy: Open Sans 10pt with the deck's 1.5 line spacing. */
function body(slide, text, x, y, w, h, opts) {
  slide.addText(text, Object.assign({
    x, y, w, h, fontFace: BODY, fontSize: 10, color: C.ink,
    valign: 'top', lineSpacingMultiple: 1.5,
  }, opts));
}

function hline(slide, x, y, w, opts) {
  slide.addShape('line', Object.assign({ x, y, w, h: 0, line: HAIRLINE }, opts));
}

function vline(slide, x, y, h, opts) {
  slide.addShape('line', Object.assign({ x, y, w: 0, h, line: HAIRLINE }, opts));
}

/** The small orange bullet that terminates most hairlines. */
function dot(slide, cx, cy) {
  slide.addShape('ellipse', { x: cx - 0.061, y: cy - 0.061, w: 0.122, h: 0.122, fill: { color: C.orange } });
}

/** Hairline that ends in an orange dot (horizontal). */
function hlineDot(slide, x, y, w, dotAtStart) {
  hline(slide, x, y, w);
  dot(slide, dotAtStart ? x : x + w, y);
}

/** Hairline that ends in an orange dot (vertical, dot at the bottom). */
function vlineDot(slide, x, y, h) {
  vline(slide, x, y, h);
  dot(slide, x, y + h);
}

/** Flat grey stand-in for a photograph, labelled "[image]". */
function photo(slide, x, y, w, h, color) {
  slide.addShape('rect', { x, y, w, h, fill: { color: color || C.photo } });
  slide.addText('[image]', {
    x, y: y + h / 2 - 0.16, w, h: 0.32,
    fontFace: BODY, fontSize: 11, color: C.caption, align: 'center', valign: 'middle',
  });
}

/**
 * Small filled square carrying a light pictograph, used for the feature and
 * contact icons. `glyph` picks one of the line-art marks drawn with shapes.
 */
function iconTile(slide, x, y, size, fill, glyph) {
  slide.addShape('rect', { x, y, w: size, h: size, fill: { color: fill }, shadow: SOFT_SHADOW });
  const m = size * 0.22; // inner margin
  const ix = x + m;
  const iy = y + m;
  const is = size - 2 * m;
  const stroke = { color: C.bg, width: 1 };
  if (glyph === 'shield') {
    slide.addShape('ellipse', { x: ix, y: iy, w: is, h: is, fill: { type: 'none' }, line: stroke });
    slide.addShape('homePlate', {
      x: ix + is * 0.16, y: iy + is * 0.18, w: is * 0.68, h: is * 0.6, rotate: 90,
      fill: { type: 'none' }, line: stroke,
    });
    slide.addShape('line', { x: ix + is * 0.5, y: iy + is * 0.2, w: 0, h: is * 0.35, line: stroke });
  } else if (glyph === 'stack') {
    slide.addShape('can', { x: ix + is * 0.1, y: iy, w: is * 0.8, h: is, fill: { type: 'none' }, line: stroke });
    slide.addShape('line', { x: ix + is * 0.1, y: iy + is * 0.62, w: is * 0.8, h: 0, line: stroke });
  } else if (glyph === 'cloud') {
    slide.addShape('cloud', { x: ix, y: iy + is * 0.1, w: is, h: is * 0.72, fill: { type: 'none' }, line: stroke });
    slide.addShape('upArrow', {
      x: ix + is * 0.4, y: iy + is * 0.3, w: is * 0.2, h: is * 0.44, fill: { type: 'none' }, line: stroke,
    });
  } else if (glyph === 'pin') {
    slide.addShape('teardrop', {
      x: ix + is * 0.12, y: iy, w: is * 0.76, h: is * 0.76, rotate: 135, fill: { color: C.bg },
    });
    slide.addShape('ellipse', {
      x: ix + is * 0.36, y: iy + is * 0.22, w: is * 0.28, h: is * 0.28, fill: { color: fill },
    });
  } else if (glyph === 'phone') {
    // handset: a curved body with a square ear-piece at each end
    slide.addShape('arc', {
      x: ix - is * 0.06, y: iy - is * 0.06, w: is * 1.12, h: is * 1.12, rotate: 80,
      line: { color: C.bg, width: 2.5 }, angleRange: [0, 130],
    });
    slide.addShape('rect', { x: ix - is * 0.02, y: iy, w: is * 0.32, h: is * 0.22, rotate: -40, fill: { color: C.bg } });
    slide.addShape('rect', { x: ix + is * 0.7, y: iy + is * 0.72, w: is * 0.32, h: is * 0.22, rotate: -40, fill: { color: C.bg } });
  } else if (glyph === 'globe') {
    slide.addShape('ellipse', { x: ix, y: iy, w: is, h: is, fill: { type: 'none' }, line: stroke });
    slide.addShape('ellipse', {
      x: ix + is * 0.32, y: iy, w: is * 0.36, h: is, fill: { type: 'none' }, line: stroke,
    });
    slide.addShape('line', { x: ix, y: iy + is / 2, w: is, h: 0, line: stroke });
  } else if (glyph === 'mail') {
    slide.addShape('rect', { x: ix, y: iy + is * 0.16, w: is, h: is * 0.68, fill: { type: 'none' }, line: stroke });
    slide.addShape('line', { x: ix, y: iy + is * 0.16, w: is / 2, h: is * 0.34, line: stroke });
    slide.addShape('line', { x: ix + is / 2, y: iy + is * 0.5, w: is / 2, h: -is * 0.34, line: stroke });
  }
}

/** Outlined "send" paper plane, drawn as five strokes inside an s x s box. */
function paperPlane(slide, x, y, s) {
  const line = { color: C.ink, width: 0.75 };
  const tip = [x, y + 0.72 * s]; // trailing wingtip
  const nose = [x + s, y]; // nose, pointing up-right
  const tail = [x + 0.45 * s, y + s];
  const fold = [x + 0.38 * s, y + 0.62 * s];
  [[tip, nose], [nose, tail], [tail, fold], [fold, tip], [fold, nose]].forEach(([a, b]) => {
    slide.addShape('line', { x: a[0], y: a[1], w: b[0] - a[0], h: b[1] - a[1], line });
  });
}

/** Dark pill button + orange circle holding the "send" arrow (slides 1 and 15). */
function heroButtons(slide, label) {
  slide.addShape('roundRect', {
    x: 1.524, y: 5.722, w: 2.227, h: 0.467, rectRadius: 0.2335,
    fill: { color: C.ink }, shadow: SOFT_SHADOW,
  });
  slide.addText(label, {
    x: 1.598, y: 5.804, w: 1.73, h: 0.303,
    fontFace: BODY, fontSize: 12, color: C.white, valign: 'top',
  });
  slide.addShape('ellipse', { x: 3.32, y: 5.781, w: 0.349, h: 0.349, fill: { color: C.orange } });
  paperPlane(slide, 3.4, 5.855, 0.19);
  slide.addShape('roundRect', {
    x: 3.953, y: 5.722, w: 1.366, h: 0.467, rectRadius: 0.2335,
    fill: { type: 'none' }, line: { color: C.orange, width: 1 }, shadow: SOFT_SHADOW,
  });
  slide.addText('Learn More', {
    x: 4.051, y: 5.804, w: 1.172, h: 0.303,
    fontFace: BODY, fontSize: 12, color: C.ink, align: 'center', valign: 'top',
  });
}

/** Small black "Learn More" chip (slides 5 and 13). */
function learnMoreChip(slide, x) {
  slide.addShape('rect', { x, y: 5.5105, w: 1.076, h: 0.353, fill: { color: C.ink }, shadow: SOFT_SHADOW });
  slide.addText('Learn More', {
    x: x - 0.036, y: 5.501, w: 1.149, h: 0.326,
    fontFace: BODY, fontSize: 10, color: C.white, align: 'center', valign: 'top', lineSpacingMultiple: 1.5,
  });
}

/**
 * Persistent chrome: logo, top navigation, footer URL and page number.
 * Identical on every slide apart from the page number.
 */
function chrome(slide, pageNo) {
  slide.background = { color: C.bg };
  slide.addText(
    [
      { text: 'Minimal', options: { bold: true } },
      { text: 'Architecture', options: {} },
    ],
    { x: 0.383, y: 0.453, w: 1.781, h: 0.303, fontFace: DISPLAY, fontSize: 12, color: C.ink, valign: 'top' }
  );
  [['Home', 8.575, true], ['About Us', 10.16, false], ['Experience', 11.745, false]].forEach(
    ([label, x, bold]) => {
      slide.addText(label, {
        x, y: 0.473, w: 1.277, h: 0.269, fontFace: BODY, fontSize: 10,
        bold, charSpacing: 1, color: C.ink, align: 'center', valign: 'bottom',
      });
    }
  );
  slide.addText('www.reallygreatsite.com', {
    x: 0.394, y: 6.791, w: 2.731, h: 0.269,
    fontFace: BODY, fontSize: 10, charSpacing: 1, color: C.ink, valign: 'middle',
  });
  slide.addText(
    [
      { text: 'Page  |  ', options: { charSpacing: 1 } },
      { text: String(pageNo), options: { bold: true } },
    ],
    { x: 11.623, y: 6.717, w: 1.301, h: 0.399, fontFace: BODY, fontSize: 12, color: C.ink, align: 'right', valign: 'middle' }
  );
}

/* -------------------------------------------------------------- slide 01 */

function slide01(pptx) {
  const s = pptx.addSlide();
  chrome(s, 1);
  heading(s, 'Beyond Walls', 1.417, 1.763, 10.194, 2.036, 115);
  heading(s, 'The Power of Built Environments', 1.417, 1.505, 4.417, 0.404, 18);
  body(s, LOREM.hero, 1.417, 4.085, 9.528, 0.579);
  heroButtons(s, 'Start Presentation');
  hlineDot(s, 0, 5.159, 5.319);
  hlineDot(s, 7.683, 3.782, 5.651, true);
  vline(s, 11.611, 0, 7.5);
  dot(s, 11.611, 3.782);
  vlineDot(s, 5.667, 0, 1.763);
}

/* -------------------------------------------------------------- slide 02 */

function slide02(pptx) {
  const s = pptx.addSlide();
  chrome(s, 2);

  heading(s, 'Table Of', 0.917, 1.471, 2.491, 0.774, 40);
  heading(s, 'Content', 0.917, 2.117, 4.773, 1.447, 80);
  body(s, LOREM.toc, 7.018, 1.908, 5.426, 1.084);

  // Six numbered contents entries laid out in three columns of two rows.
  const columns = [
    { numX: 0.842, numW: [0.909, 1.156], textX: 1.851 },
    { numX: 4.802, numW: [1.082, 1.008], textX: 5.81 },
    { numX: 8.752, numW: [1.082, 1.008], textX: 9.76 },
  ];
  const rows = [
    { numY: 4.136, textY: 4.338 },
    { numY: 5.519, textY: 5.721 },
  ];
  columns.forEach((col, ci) => {
    rows.forEach((row, ri) => {
      const n = ci * 2 + ri + 1;
      // second row of columns 2 & 3 sits a hair lower than column 1
      const shift = ri === 1 && ci > 0 ? 0.036 : 0;
      s.addText('0' + n, {
        x: col.numX, y: row.numY + shift, w: col.numW[ri], h: 0.774,
        fontFace: DISPLAY, fontSize: 40, bold: true, color: C.ink, transparency: 95, valign: 'top',
      });
      heading(s, 'Your Subtitle Goes Here', col.textX, row.textY + shift, 3.147, 0.37, 16);
    });
  });

  s.addShape('line', { x: 0.976, y: 5.184, w: 11.357, h: 0, line: { color: C.ink, transparency: 70, width: 0.5 } });
  vlineDot(s, 6.167, 0, 3.162);
}

/* -------------------------------------------------------------- slide 03 */

function slide03(pptx) {
  const s = pptx.addSlide();
  photo(s, 1.056, 1.465, 4.569, 4.569);
  chrome(s, 3);

  s.addShape('rect', { x: 5.625, y: 4.514, w: 7.708, h: 1.521, fill: { color: C.greyBlock } });
  heading(s, 'Shaping the Modern Skyline', 6.317, 1.904, 4.877, 1.447, 40);
  body(s, LOREM.intro, 6.317, 3.569, 6.099, 0.579);

  s.addShape('rect', { x: 5.625, y: 4.514, w: 3.847, h: 1.521, fill: { color: C.orange } });
  heading(s, '123,45K+', 6.317, 4.823, 2.358, 0.64, 32, { color: C.white });
  body(s, 'Quisque non elit mauris. Cras euismod', 6.317, 5.404, 2.758, 0.326, { color: C.white });
  body(s, 'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, a',
    9.852, 4.899, 2.747, 0.831);

  hlineDot(s, 8.575, 1.534, 4.759, true);
  vline(s, 12.605, 1.534, 4.501);
  dot(s, 12.605, 1.534);
}

/* -------------------------------------------------------------- slide 04 */

function slide04(pptx) {
  const s = pptx.addSlide();
  vline(s, 9.885, 1.181, 6.337);
  chrome(s, 4);

  // Two stat cards, ghost numeral + title + copy, one grey and one orange.
  const cards = [
    { x: 7.556, fill: C.greyBlock, numX: 7.655, numW: 0.888, numColor: C.greyNum, textX: 7.901, title: 'Urban Code', ink: C.ink },
    { x: 9.885, fill: C.orange, numX: 9.984, numW: 1.12, numColor: C.orangeDim, textX: 10.208, title: 'Urban Pulse', ink: C.white },
  ];
  // orange card is drawn first in the original, then the grey one overlaps it
  s.addShape('rect', { x: cards[1].x, y: 1.465, w: 2.337, h: 1.521, fill: { color: cards[1].fill } });
  s.addShape('rect', { x: cards[0].x, y: 1.465, w: 2.33, h: 1.521, fill: { color: cards[0].fill } });

  heading(s, 'Designing with the Earth in Mind', 0.901, 1.654, 4.877, 1.447, 40);
  body(s, LOREM.intro, 0.901, 3.251, 5.274, 0.579);

  // Two feature icons with heading + copy under the hero paragraph.
  const features = [
    { x: 0.901, tileX: 0.989, fill: C.orange, glyph: 'shield', title: 'Modernism Reborn' },
    { x: 3.671, tileX: 3.749, fill: C.ink, glyph: 'stack', title: 'Concrete Futures' },
  ];
  features.forEach((f) => {
    iconTile(s, f.tileX, 4.281, 0.43, f.fill, f.glyph);
    heading(s, f.title, f.x, 4.908, 2.13, 0.337, 14);
    body(s, LOREM.tinier, f.x, 5.267, 2.342, 0.579);
  });

  cards.forEach((c, i) => {
    s.addText('0' + (i + 1), {
      x: c.numX, y: 1.484, w: c.numW, h: 0.909,
      fontFace: DISPLAY, fontSize: 48, color: c.numColor, transparency: 75, valign: 'top',
    });
    heading(s, c.title, c.textX, 1.772, 2.061, 0.337, 14, { color: c.ink });
    body(s, 'Quisque non elit mauri. Cras euismod, met', c.textX, 2.13, 1.951, 0.579, { color: c.ink });
  });

  hlineDot(s, 3.569, 1.173, 9.764, true);
  photo(s, 7.556, 2.986, 4.667, 3.049);
}

/* -------------------------------------------------------------- slide 05 */

function slide05(pptx) {
  const s = pptx.addSlide();
  vline(s, 3.556, 0, 6.493);
  chrome(s, 5);

  s.addShape('rect', { x: 3.556, y: 1.465, w: 2.5, h: 4.569, fill: { color: C.orange } });
  heading(s, 'Where Aesthetics Meets Function', 6.984, 1.637, 4.761, 1.447, 40);
  heading(s, 'Residential Architecture Now', 6.984, 3.336, 2.96, 0.337, 14);
  body(s,
    'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, ' +
    'a maximus leo ligula at dolor. Morbi et malesuada purus. Phasellus a lacus sit amet urna ' +
    'tempor sollicitudin ac finibus finibus, felis dui suscipit ', 6.984, 3.719, 5.612, 0.831);
  body(s,
    'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, ' +
    'a maximus leo ligula at dolor. Morbi et malesuada purus. ', 6.984, 4.642, 5.612, 0.579);
  learnMoreChip(s, 7.0985);

  // Two statistics stacked inside the orange panel.
  heading(s, '123,4K+', 3.946, 1.837, 1.915, 0.64, 32, { color: C.white });
  body(s, 'Quisque non elit mauris.', 3.946, 2.418, 1.86, 0.326, { color: C.white });
  body(s, 'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, ',
    3.946, 3.222, 1.915, 1.084, { color: C.white });
  heading(s, '50%', 3.947, 4.696, 1.143, 0.64, 32, { color: C.white });
  body(s, 'Quisque none', 3.946, 5.278, 1.143, 0.326, { color: C.white });

  hlineDot(s, 0, 6.493, 8.211);
  photo(s, 1.056, 1.465, 2.5, 4.569);
}

/* -------------------------------------------------------------- slide 06 */

function slide06(pptx) {
  const s = pptx.addSlide();
  vline(s, 10.012, 0, 2.57);
  chrome(s, 6);

  // Two numbered cards: ghost numeral, short colour rule, title and copy.
  const cards = [
    { numX: 1.358, numW: 0.909, ruleX: 1.016, ruleColor: C.orange, textX: 2.294, title: 'Cityscapes of Tomorrow', titleW: 2.458 },
    { numX: 7.416, numW: 1.177, ruleX: 7.073, ruleColor: C.ink, textX: 8.351, title: 'Shaping the Modern Skyline', titleW: 3.076 },
  ];
  cards.forEach((c, i) => {
    heading(s, c.title, c.textX, 4.791, c.titleW, 0.337, 14);
    body(s, LOREM.card, c.textX, 5.078, 3.966, 0.831);
    s.addText('0' + (i + 1), {
      x: c.numX, y: 4.691, w: c.numW, h: 0.774,
      fontFace: DISPLAY, fontSize: 40, bold: true, color: C.ink, transparency: 95, valign: 'top',
    });
    s.addShape('rect', { x: c.ruleX, y: 5.044, w: 0.252, h: 0.068, fill: { color: c.ruleColor } });
  });

  heading(s, 'Minimalism in the Metropolis', 0.85, 1.627, 4.877, 1.447, 40);
  body(s,
    'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, ' +
    'a maximus leo ligula at dolor. Morbi et malesuada purus. Phasellus a lacus sit amet urna ' +
    'tempor sollicitudin ac finibus finibus, felis dui suscipit Cras euismod, metus ac finibus finibus, felis ',
    0.85, 3.692, 11.644, 0.579);

  hlineDot(s, 3.81, 1.138, 6.202, true);
  photo(s, 7.073, 1.59, 5.244, 1.623);
}

/* -------------------------------------------------------------- slide 07 */

function slide07(pptx) {
  const s = pptx.addSlide();
  chrome(s, 7);
  heading(s, 'Interior Architecture Where Aesthetics Meets Function', 0.849, 1.485, 7.726, 1.447, 40);

  // Two text columns, each with a subtitle and two identical paragraphs.
  [
    { x: 0.849, title: 'The Beauty of Small Spaces' },
    { x: 7.199, title: 'From Concept to Comfort' },
  ].forEach((col) => {
    heading(s, col.title, col.x, 3.319, 3.938, 0.37, 16);
    body(s, LOREM.column, col.x, 3.793, 5.214, 1.084);
    body(s, LOREM.column, col.x, 5.062, 5.214, 1.084);
  });

  vline(s, 11.545, 0, 2.57);
  hlineDot(s, 2.967, 1.138, 8.579, true);
  dot(s, 11.545, 2.57);
}

/* -------------------------------------------------------------- slide 08 */

function slide08(pptx) {
  const s = pptx.addSlide();
  hlineDot(s, 1.0, 3.766, 12.314, true);
  chrome(s, 8);

  const team = [
    { photoX: 1.942, nameX: 1.517, nameW: 3.216, copyX: 1.624, name: 'Benjamin Esteban' },
    { photoX: 5.483, nameX: 5.141, nameW: 3.052, copyX: 5.166, name: 'Bryan Govanchy' },
    { photoX: 9.025, nameX: 8.601, nameW: 3.216, copyX: 8.708, name: 'Benjamin Esteban' },
  ];
  team.forEach((m) => {
    heading(s, m.name, m.nameX, 5.198, m.nameW, 0.37, 16, { align: 'center' });
    body(s, LOREM.tiny, m.copyX, 5.539, 3.004, 0.326, { align: 'center' });
  });

  heading(s, 'Meet Our Best Team Here', 2.804, 1.277, 7.726, 0.774, 40, { align: 'center' });
  team.forEach((m) => photo(s, m.photoX, 2.567, 2.367, 2.367));
}

/* -------------------------------------------------------------- slide 09 */

function slide09(pptx) {
  const s = pptx.addSlide();
  photo(s, 7.861, 1.465, 5.473, 4.569);
  chrome(s, 9);

  s.addShape('rect', { x: 1.354, y: 1.465, w: 6.506, h: 4.569, fill: { color: C.greyBlock } });
  heading(s, 'It\u2019s Time to Break', 2.221, 2.092, 4.773, 2.794, 80);
  heading(s, '30 Minutes Coffee Break', 2.221, 4.955, 2.458, 0.337, 14, { color: 'FF0000' });

  hlineDot(s, 0, 6.449, 6.667);
  vline(s, 0.697, 2.519, 3.93);
  dot(s, 0.697, 2.58);
}

/* -------------------------------------------------------------- slide 10 */

function slide10(pptx) {
  const s = pptx.addSlide();
  // Four-tile photo collage filling the right two thirds of the slide.
  photo(s, 7.882, 3.75, 2.726, 2.285, C.photoMid);
  photo(s, 5.156, 3.75, 2.726, 2.285);
  photo(s, 10.608, 1.465, 2.726, 4.569);
  photo(s, 5.156, 1.465, 5.451, 2.285, C.photoDark);
  chrome(s, 10);

  heading(s, 'Our Company Gallery Here', 0.849, 3.026, 3.86, 1.447, 40);
  hlineDot(s, 0, 4.93, 4.0);
  vlineDot(s, 4.0, 0, 2.57);
}

/* -------------------------------------------------------------- slide 11 */

/** Flat stand-in for the tablet product shot: edge view, rear unit, front unit. */
function tabletMockup(s) {
  s.addShape('roundRect', {
    x: 1.04, y: 1.6, w: 0.16, h: 4.33, rectRadius: 0.05, rotate: 7,
    fill: { color: C.deviceGrey },
  });
  s.addShape('roundRect', { x: 4.3, y: 1.6, w: 1.9, h: 4.33, rectRadius: 0.14, fill: { color: C.deviceSilver } });
  s.addShape('roundRect', {
    x: 1.5, y: 1.59, w: 2.9, h: 4.34, rectRadius: 0.14,
    fill: { color: C.deviceWhite }, shadow: { type: 'outer', blur: 16, offset: 4, angle: 90, color: C.ink, opacity: 0.12 },
  });
  s.addShape('ellipse', { x: 2.925, y: 1.735, w: 0.05, h: 0.05, fill: { color: '9AA0A3' } }); // camera
  s.addShape('ellipse', {
    x: 2.845, y: 5.6, w: 0.21, h: 0.21, fill: { type: 'none' }, line: { color: 'D2D2D2', width: 1 },
  }); // home button
  photo(s, 1.64, 2.0, 2.658, 3.51);
}

function slide11(pptx) {
  const s = pptx.addSlide();
  tabletMockup(s);
  chrome(s, 11);

  heading(s, 'The Beauty of Small Spaces', 6.919, 1.92, 4.41, 1.447, 40);
  body(s,
    'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscip it p urus, ' +
    'a maximus leo ligula at dolor. Morbi et', 6.919, 3.728, 5.663, 0.579);
  s.addText('123,45K+', {
    x: 6.919, y: 4.531, w: 2.358, h: 0.64, fontFace: BODY, fontSize: 32, color: C.ink, valign: 'top',
  });
  body(s, LOREM.tiny, 6.919, 5.112, 2.29, 0.326);
  body(s,
    'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui Quisque non elit ' +
    'mauris. Cras euismod, metus', 9.277, 4.607, 3.305, 0.831);

  vline(s, 11.545, 0, 3.101);
  hlineDot(s, 8.148, 1.138, 3.398, true);
  dot(s, 11.545, 3.101);
}

/* -------------------------------------------------------------- slide 12 */

/** Flat stand-in for the desktop-computer product shot: bezel, chin, neck, foot. */
function imacMockup(s) {
  s.addShape('rect', {
    x: 7.118, y: 1.861, w: 5.198, h: 2.949, fill: { color: C.deviceDark },
    shadow: { type: 'outer', blur: 40, offset: 6, angle: 90, color: C.ink, opacity: 0.2 },
  });
  s.addShape('ellipse', { x: 9.695, y: 1.9, w: 0.045, h: 0.045, fill: { color: '4A4A4A' } }); // camera
  s.addShape('rect', { x: 7.115, y: 4.81, w: 5.39, h: 0.47, fill: { color: C.deviceGrey } });
  s.addShape('trapezoid', { x: 8.48, y: 5.27, w: 2.74, h: 0.44, flipV: true, fill: { color: C.deviceGrey } });
  s.addShape('ellipse', { x: 8.42, y: 5.66, w: 2.8, h: 0.1, fill: { color: C.deviceGrey } });
  photo(s, 7.303, 2.032, 4.844, 2.618);
}

function slide12(pptx) {
  const s = pptx.addSlide();
  imacMockup(s);
  chrome(s, 12);

  // Orange call-to-action bar with a dark icon tile.
  s.addShape('rect', { x: 1.035, y: 5.168, w: 3.699, h: 0.828, fill: { color: C.orange }, shadow: SOFT_SHADOW });
  heading(s, 'Crafting Personal Spaces', 1.824, 5.393, 2.69, 0.37, 16, { color: C.bg });
  iconTile(s, 1.234, 5.35, 0.458, C.ink, 'cloud');

  heading(s, 'Human-Centered Living Spaces', 0.942, 1.68, 4.595, 1.447, 40);
  heading(s, 'Designing the Dream Home', 0.942, 3.429, 3.938, 0.37, 16);
  body(s,
    'Quisque non elit mauris. Cras euismod, metus ac finibus fini bus, felis dui suscipit purus, ' +
    'a maximus leo ligula euis mod, metus euismod, metus ac finibus finibus fini bus, felis dui ' +
    'suscipit maximus leo ligula', 0.942, 3.903, 5.683, 0.831);

  vline(s, 6.208, 0, 2.848);
  hlineDot(s, 2.811, 1.138, 3.398, true);
  dot(s, 6.208, 2.848);
}

/* -------------------------------------------------------------- slide 13 */

function slide13(pptx) {
  const s = pptx.addSlide();
  chrome(s, 13);

  // Two stacked numbered blocks in the left column, separated by a hairline.
  [
    { numY: 1.462, numW: 0.909, titleY: 2.23, copyY: 2.518, title: 'Smart Cities, Smart Design' },
    { numY: 4.057, numW: 1.409, titleY: 4.825, copyY: 5.113, title: 'Cityscapes of Tomorrow' },
  ].forEach((blk, i) => {
    s.addText('0' + (i + 1), {
      x: 0.954, y: blk.numY, w: blk.numW, h: 0.774,
      fontFace: DISPLAY, fontSize: 40, bold: true, color: C.ink, transparency: 95, valign: 'top',
    });
    heading(s, blk.title, 0.954, blk.titleY, 2.51, 0.337, 14);
    body(s, LOREM.cardShort, 0.954, blk.copyY, 3.539, 0.831);
  });
  s.addShape('line', { x: 1.063, y: 3.775, w: 3.278, h: 0, line: { color: C.greyBlock, width: 0.5 } });

  heading(s, 'Vertical Living The Rise of High-Rise', 5.565, 1.68, 5.967, 1.447, 40);
  body(s, LOREM.wide, 5.551, 3.361, 6.828, 0.831);
  body(s, LOREM.wide, 5.551, 4.33, 6.828, 0.831);
  learnMoreChip(s, 5.6685);

  vline(s, 5.006, 0, 2.447);
  hlineDot(s, 4.999, 1.138, 3.398);
  dot(s, 5.006, 2.447);
}

/* -------------------------------------------------------------- slide 14 */

function slide14(pptx) {
  const s = pptx.addSlide();
  photo(s, 1.056, 1.465, 4.569, 4.569);
  chrome(s, 14);

  heading(s, 'Our Contact Information Here', 6.55, 1.68, 5.073, 1.447, 40);
  body(s, LOREM.intro, 6.586, 3.484, 5.692, 0.579);

  // 2 x 2 grid of contact details, each an icon tile plus a line of text.
  const contacts = [
    { tileX: 6.68, y: 4.629, fill: C.orange, glyph: 'pin', textX: 7.24, text: '123 Anywhere St., Any City' },
    { tileX: 6.686, y: 5.351, fill: C.ink, glyph: 'phone', textX: 7.24, text: '+123-456-7890' },
    { tileX: 9.467, y: 4.629, fill: C.ink, glyph: 'globe', textX: 10.027, text: 'www.reallygreatsite.com' },
    { tileX: 9.467, y: 5.351, fill: C.orange, glyph: 'mail', textX: 10.027, text: 'hello@reallygreatsite.com' },
  ];
  contacts.forEach((c) => {
    iconTile(s, c.tileX, c.y, 0.438, c.fill, c.glyph);
    s.addText(c.text, {
      x: c.textX, y: c.y + 0.042, w: 1.977, h: 0.353,
      fontFace: BODY, fontSize: 10, color: C.ink, valign: 'middle', lineSpacingMultiple: 1.5,
    });
  });

  vline(s, 11.545, 0, 2.917);
  hlineDot(s, 7.817, 1.138, 3.729, true);
  dot(s, 11.545, 2.917);
}

/* -------------------------------------------------------------- slide 15 */

function slide15(pptx) {
  const s = pptx.addSlide();
  chrome(s, 15);
  heading(s, 'Thanks a Lot', 1.417, 1.763, 10.194, 2.036, 115);
  heading(s, 'The Power of Built Environments', 1.417, 1.505, 4.417, 0.404, 18);
  body(s, LOREM.hero, 1.417, 4.085, 9.528, 0.579);
  heroButtons(s, 'End Presentation');
  hlineDot(s, 0, 5.159, 5.319);
  hlineDot(s, 8.183, 3.782, 5.15, true);
  vline(s, 11.611, 0, 7.5);
  dot(s, 11.611, 3.782);
  vlineDot(s, 5.667, 0, 1.763);
}

/* ------------------------------------------------------------------ main */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE_13_33x7_5', width: 13.333333, height: 7.5 }); // 12192000 x 6858000 EMU
  pptx.layout = 'WIDE_13_33x7_5';
  pptx.title = 'Beyond Walls';
  pptx.company = 'MinimalArchitecture';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
    slide09, slide10, slide11, slide12, slide13, slide14, slide15].forEach((fn) => fn(pptx));

  return pptx;
}

const outFile = path.join(__dirname, '014293f6-4918-4681-aa01-609e99002b0d_grok_final.pptx');
build()
  .writeFile({ fileName: outFile })
  .then(() => console.log('wrote ' + outFile))
  .catch((err) => {
    console.error(err);
    process.exit(1);
  });
