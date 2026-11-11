/**
 * "Victory Arena" e-sport deck — rebuilt with pptxgenjs.
 * 20 slides, 13.333in x 7.5in (16:9), Bebas Neue display + Public Sans body.
 *
 * The four raster photos of the original (smart watch, phone mock-ups, game
 * controller, laptop) are replaced by flat "[image]" placeholder rectangles
 * that keep the original position and size. Everything else — gradients,
 * chevrons, pill tags, icon tiles, colour grid — is drawn with native shapes.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ================================================================== *
 * Design tokens (theme "Custom 99" + the tints the deck actually uses)
 * ================================================================== */
const C = {
  purple: 'B357FA',      // accent1
  purpleSoft: 'D19AFC',  // accent1 @ 85% lum — nav links on a purple ground
  ink: '262626',         // tx1 @ 85% lum — the deck's "black"
  inkDeep: '0D0D0D',     // bottom stop of the dark-slide gradient
  gray40: '404040',      // tx1 @ 75% lum
  gray50: '595959',
  gray80: '808080',      // tx1 @ 50% lum
  grayBF: 'BFBFBF',      // muted labels + slide numbers
  grayD9: 'D9D9D9',      // nav links on dark
  grayF2: 'F2F2F2',      // card / swatch grey
  white: 'FFFFFF',
};

const F = { head: 'Bebas Neue', body: 'Public Sans' };

const SLIDE_W = 13.3333;
const SLIDE_H = 7.5;

// Text boxes keep PowerPoint's default 0.1"/0.05" insets.
// pptxgenjs takes them in points, ordered [left, right, bottom, top].
const INSET = [7.2, 7.2, 3.6, 3.6];

/* ================================================================== *
 * Primitives
 * ================================================================== */

/** Filled, borderless shape — `rect` unless `shape` says otherwise. */
function box(slide, o) {
  const opts = {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: o.alpha === undefined ? { color: o.fill } : { color: o.fill, transparency: o.alpha },
    line: { type: 'none' },
  };
  if (o.rotate) opts.rotate = o.rotate;
  if (o.flipH) opts.flipH = true;
  if (o.radius !== undefined) opts.rectRadius = o.radius;
  slide.addShape(o.shape || 'rect', opts);
}

/**
 * Text box. `content` is a string, an array of paragraphs, or — for mixed
 * colour/weight/alignment lines — an array where a paragraph is itself an
 * array of {t, color, bold, align} fragments.
 */
function text(slide, content, o) {
  const paras = Array.isArray(content) ? content : [content];
  const rich = [];
  paras.forEach((para, pi) => {
    const frags = Array.isArray(para) ? para : [{ t: para }];
    const paraAlign = frags[0].align;
    frags.forEach((fr, fi) => {
      rich.push({
        text: fr.t,
        options: {
          color: fr.color || o.color || C.ink,
          bold: fr.bold || false,
          align: paraAlign || o.align || 'left',
          breakLine: fi === frags.length - 1 && pi < paras.length - 1,
        },
      });
    });
  });
  const opts = {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: o.font || F.body,
    fontSize: o.size,
    color: o.color || C.ink,
    align: o.align || 'left',
    valign: o.valign || 'top',
    margin: INSET,
    isTextBox: true,
    wrap: o.wrap !== false,
    fit: 'resize',
  };
  if (o.lineSpacing) opts.lineSpacingMultiple = o.lineSpacing;
  if (o.rotate) opts.rotate = o.rotate;
  if (o.flipH) opts.flipH = true;
  slide.addText(rich, opts);
}

/** Display type — Bebas Neue. */
function heading(slide, content, o) {
  text(slide, content, Object.assign({ font: F.head, color: C.ink }, o));
}

/** 12pt body copy; the deck always runs these at 130% leading. */
function body(slide, content, o) {
  text(slide, content, Object.assign({ size: 12, lineSpacing: 1.3 }, o));
}

/* ================================================================== *
 * Recurring composite elements
 * ================================================================== */

/** Purple pill reading "Victory creates legends" — appears on 11 slides. */
function tagPill(slide, x, y) {
  box(slide, { x: x + 0.2116, y: y, w: 1.8759, h: 0.4026, fill: C.purple });
  heading(slide, 'Victory creates legends', {
    x: x, y: y + 0.0577, w: 2.299, h: 0.3029, size: 12, color: C.white, align: 'center',
  });
}

/**
 * Square badge holding a white gamepad pictogram. The source uses custom
 * geometry glyphs (gamepad / mouse / medal / headset); one rounded gamepad
 * mark stands in for all of them.
 */
function iconTile(slide, x, y, s, tileColor, markColor) {
  box(slide, { x: x, y: y, w: s, h: s, fill: tileColor });
  const bw = s * 0.60, bh = s * 0.32;
  const bx = x + (s - bw) / 2, by = y + s * 0.34;
  // body + two grips
  box(slide, { x: bx, y: by, w: bw, h: bh, fill: markColor, shape: 'roundRect', radius: bh * 0.38 });
  box(slide, { x: bx, y: by + bh * 0.62, w: s * 0.22, h: s * 0.22, fill: markColor, shape: 'ellipse' });
  box(slide, { x: bx + bw - s * 0.22, y: by + bh * 0.62, w: s * 0.22, h: s * 0.22, fill: markColor, shape: 'ellipse' });
  // knocked-out d-pad and two buttons
  box(slide, { x: x + s * 0.265, y: by + bh * 0.40, w: s * 0.15, h: s * 0.05, fill: tileColor });
  box(slide, { x: x + s * 0.315, y: by + bh * 0.13, w: s * 0.05, h: s * 0.30, fill: tileColor });
  box(slide, { x: x + s * 0.625, y: by + bh * 0.20, w: s * 0.08, h: s * 0.08, fill: tileColor, shape: 'ellipse' });
  box(slide, { x: x + s * 0.625, y: by + bh * 0.62, w: s * 0.08, h: s * 0.08, fill: tileColor, shape: 'ellipse' });
}

/** "Vicotry|x" wordmark and its hexagon glyph (inherited from the master). */
function logo(slide, wordColor, markColor) {
  box(slide, { x: 0.6426, y: 0.4093, w: 0.2196, h: 0.2396, fill: markColor || C.purple, shape: 'hexagon' });
  heading(slide, [[{ t: 'Vicotry', color: wordColor }, { t: 'x', color: C.purple }]], {
    x: 0.9019, y: 0.3777, w: 0.9757, h: 0.3029, size: 12,
  });
}

/** contact / about / home, right aligned in the header. */
function navLinks(slide, color, o) {
  const opt = o || {};
  let items = [['contact', 10.6334], ['about', 11.3558], ['home', 12.0781]];
  if (opt.skipContact) items = items.slice(1);
  items.forEach(([label, x]) => {
    heading(slide, label, {
      x: x, y: 0.3777, w: 0.7223, h: 0.3029, size: 12, align: 'right',
      color: label === 'home' ? (opt.home || color) : color,
    });
  });
}

/** Logo + nav, as inherited from the slide master. */
function header(slide, o) {
  const opt = o || {};
  logo(slide, opt.word || C.gray40, opt.mark);
  navLinks(slide, opt.nav || C.grayD9, { home: opt.home });
}

/** Page number in the bottom-right master placeholder. */
function pageNo(slide, n, color) {
  text(slide, String(n), {
    x: 9.8005, y: 6.9514, w: 3.0, h: 0.3993, size: 12,
    align: 'right', valign: 'middle', color: color || C.grayBF,
  });
}

/** Full-bleed 262626 -> 0D0D0D vertical gradient of the dark slides. */
function darkBackdrop(slide) {
  slide.background = { color: C.ink };
  const bands = 10;
  for (let i = 0; i < bands; i++) {
    const t = i / (bands - 1);
    const v = Math.round(0x26 + (0x0d - 0x26) * t);
    const hex = v.toString(16).toUpperCase().padStart(2, '0').repeat(3);
    box(slide, { x: 0, y: (SLIDE_H / bands) * i, w: SLIDE_W, h: SLIDE_H / bands + 0.02, fill: hex });
  }
}

/** "Victory • Esport • creates • champions" strip (slides 1 and 20). */
function wordStrip(slide) {
  [['Victory', 3.3861, 1.6472], ['Esport', 4.8569, 1.8241],
    ['creates', 6.5047, 1.7027], ['champions', 8.1836, 1.8802]].forEach(([w, x, cw]) => {
    heading(slide, w, { x: x, y: 2.0474, w: cw, h: 0.5049, size: 24, color: C.white, align: 'center' });
  });
  [4.9454, 6.5047, 8.1265].forEach((x) => {
    box(slide, { x: x, y: 2.225, w: 0.0879, h: 0.0879, fill: C.white, shape: 'ellipse' });
  });
}

/** Chip rotated ~4.5 degrees anticlockwise ("Game force", "60 minutes"). */
function tiltChip(slide, x, y, label, fill, textColor) {
  box(slide, { x: x, y: y, w: 1.8212, h: 0.643, fill: fill, rotate: 355.56 });
  heading(slide, label, {
    x: x + 0.0505, y: y + 0.0868, w: 1.7202, h: 0.5049, size: 24,
    color: textColor, align: 'center', valign: 'middle', rotate: 355.56,
  });
}

/** Big number where the trailing "+" is accent purple, e.g. "620+". */
function statPlus(slide, x, y, value, plus, o) {
  const op = o || {};
  heading(slide, [[{ t: value, color: op.color || C.ink }, { t: plus, color: C.purple }]], {
    x: x, y: y, w: op.w || 2.224, h: 0.8415, size: 44,
  });
}

/**
 * Stand-in for one of the deck's four embedded photographs. The originals are
 * cut-outs on a transparent ground, so the block is drawn semi-transparent to
 * keep whatever sits behind it (purple panel, dark halo) readable.
 */
function imagePlaceholder(slide, o) {
  box(slide, { x: o.x, y: o.y, w: o.w, h: o.h, fill: o.fill, alpha: o.alpha });
  text(slide, '[image]', {
    x: o.x, y: o.y + o.h / 2 - 0.18, w: o.w, h: 0.36, size: 12,
    color: o.label || C.gray80, align: 'center',
  });
}

/* ================================================================== *
 * Slide builders
 * ================================================================== */

// 1 — Cover: VICTORY ARENA on the dark gradient.
function slide01(p) {
  const s = p.addSlide();
  darkBackdrop(s);
  heading(s, 'VICTORY ARENA ', {
    x: 0, y: 2.5523, w: SLIDE_W, h: 2.7937, size: 160, color: C.purple, align: 'center',
  });
  heading(s, 'Victory Esport Unites Players and Builds Legends', {
    x: 1.7986, y: 5.0671, w: 9.7361, h: 0.4376, size: 20, color: C.purple, align: 'center',
  });
  wordStrip(s);
  tiltChip(s, 7.887, 3.4285, 'Game force', C.gray40, C.white);
  logo(s, C.white);
  navLinks(s, C.gray80, { home: C.purple });
  heading(s, 'www.yourwebsite.com', {
    x: 0.5433, y: 6.8194, w: 2.0262, h: 0.3029, size: 12, color: C.gray80,
  });
  heading(s, 'E-sport presentation / 2029', {
    x: 10.7743, y: 6.8194, w: 2.0262, h: 0.3029, size: 12, color: C.gray80, align: 'right',
  });
}

// 2 — Table of contents.
function slide02(p) {
  const s = p.addSlide();
  tagPill(s, 0.4308, 1.4051);
  heading(s, 'Table of content overview', { x: 0.5413, y: 2.0243, w: 6.9518, h: 0.9088, size: 48 });
  body(s, 'Victory E-sport Builds Global Champions', { x: 0.5413, y: 3.0341, w: 4.0254, h: 0.3365 });

  const toc = [
    ['01', 'Introduction to Victory E-Sport', 2.9688, false],
    ['02', 'Vision, Mission, and Core Values', 3.9271, false],
    ['03', 'Competitive Gaming Landscape', 4.8854, true],
    ['04', 'Our Core Strategies for Success', 5.8438, false],
  ];
  toc.forEach(([num, label, y, active]) => {
    const fg = active ? C.white : C.grayBF;
    box(s, { x: 7.809, y: y, w: 4.8819, h: 0.8037, fill: active ? C.purple : C.grayF2 });
    heading(s, num, { x: 8.1486, y: y + 0.1998, w: 0.6417, h: 0.4039, size: 18, color: fg });
    heading(s, label, { x: 8.7903, y: y + 0.2167, w: 3.7861, h: 0.3702, size: 16, color: fg });
  });

  iconTile(s, 0.6459, 4.9902, 0.559, C.purple, C.white);
  heading(s, '35.8%', { x: 1.4852, y: 4.7784, w: 5.8389, h: 1.1107, size: 60 });
  body(s, 'Victory E-Sport inspires gamers worldwide, building strong communities, fostering teamwork, and shaping the future.',
    { x: 0.535, y: 6.0484, w: 5.2886, h: 0.5991 });
  body(s, ['The Victory E-Sport inspires ', 'gamers, building communities.'],
    { x: 10.25, y: 2.0449, w: 3.2828, h: 0.5991 });
  header(s);
  pageNo(s, 2);
}

// 3 — Team: four tilted name chips under empty portrait frames.
function slide03(p) {
  const s = p.addSlide();
  heading(s, 'MEET OUR VICTORY TEAM', {
    x: 4.2491, y: 1.1019, w: 4.8352, h: 0.7068, size: 36, align: 'center',
  });
  body(s, ['Victory E-Sport inspires gamers worldwide, building strong communities ',
    'and shaping the future of competitive gaming.'],
  { x: 2.9375, y: 1.8971, w: 7.4583, h: 0.5991, align: 'center' });

  [['Chiaki ', 'sato', 1.9343, 2.0141], ['Bailey ', 'dupont', 4.859, 4.9388],
    ['Matt ', 'zhang', 7.806, 7.8858], ['Neil ', 'tran', 10.7368, 10.8166]]
    .forEach(([first, last, tx, bx]) => {
      box(s, { x: bx, y: 5.9739, w: 1.5606, h: 0.5509, fill: C.purple, rotate: 356 });
      heading(s, [[{ t: first }, { t: last }]], {
        x: tx, y: 6.0551, w: 1.7202, h: 0.4039, size: 18, color: C.white,
        align: 'center', valign: 'middle', rotate: 356,
      });
    });
  header(s);
  pageNo(s, 3);
}

// 4 — Esport industry innovation + the 05.2029 purple block.
function slide04(p) {
  const s = p.addSlide();
  box(s, { x: 5.2454, y: 5.3698, w: 4.5404, h: 2.1302, fill: C.purple });
  heading(s, '05.2029', {
    x: 5.2454, y: 5.829, w: 4.5404, h: 1.2117, size: 66, color: C.white, align: 'center',
  });

  tagPill(s, 0.4308, 1.3179);
  heading(s, ['Esport ', 'Industry ', 'innovation'], { x: 0.5163, y: 1.9371, w: 5.5282, h: 3.1303, size: 60 });
  body(s, ['The Victory E-Sport inspires gamers, building ', 'communities, fostering teamwork, and shaping ',
    'the future of competitive gaming.'], { x: 0.5413, y: 5.3205, w: 5.2886, h: 0.8616 });

  statPlus(s, 10.245, 3.2878, '620', '+');
  body(s, ['The Victory E-Sport ', 'inspires gamers, building ', 'communities.'],
    { x: 10.245, y: 4.2058, w: 3.2828, h: 0.8616 });
  header(s);
  pageNo(s, 4);
}

// 5 — Dark slide: market-growth chevron band and two footer stats.
function slide05(p) {
  const s = p.addSlide();
  darkBackdrop(s);
  heading(s, ['Esport Market ', 'Growth Trends Analysis'], {
    x: 0.5168, y: 1.5307, w: 6.5311, h: 1.582, size: 44, color: C.white,
  });
  box(s, { x: 8.0757, y: 1.6085, w: 1.8759, h: 0.4026, fill: C.purple });
  heading(s, 'Victory creates legends', {
    x: 7.8641, y: 1.6661, w: 2.299, h: 0.3029, size: 12, color: C.white, align: 'center',
  });
  body(s, ['Victory E-Sport inspires gamers worldwide, building strong ',
    'communities and shaping the future of competitive gaming.'],
  { x: 7.9725, y: 2.3447, w: 7.4583, h: 0.5991, color: C.white });

  // Band: flat-backed left slab, purple chevron, grey chevron on the right.
  box(s, { x: 7.784, y: 3.5801, w: 4.0558, h: 1.3453, fill: C.gray40, shape: 'chevron' });
  box(s, { x: 0.6424, y: 3.5801, w: 5.8897, h: 1.3453, fill: C.gray40, shape: 'homePlate' });
  box(s, { x: 6.0007, y: 3.5801, w: 2.3147, h: 1.3453, fill: C.purple, shape: 'chevron' });
  [['TEAM STRATEGY', 'Victory E-sport Empowers Gamers', 0.6327],
    ['MARKET', 'Victory E-sport Rises', 7.0478]].forEach(([title, sub, x]) => {
    heading(s, title, { x: x, y: 3.8448, w: 5.5282, h: 0.5722, size: 28, color: C.white, align: 'center' });
    text(s, sub, { x: x, y: 4.3578, w: 5.5282, h: 0.3029, size: 12, color: C.gray80, align: 'center' });
  });
  heading(s, '60%', { x: 4.5546, y: 3.8993, w: 5.5282, h: 0.7068, size: 36, color: C.white, align: 'center' });

  iconTile(s, 0.6431, 5.9061, 0.559, C.purple, C.white);
  heading(s, '532m', { x: 1.52, y: 5.7968, w: 3.4282, h: 0.8415, size: 44, color: C.white });
  body(s, ['Victory inspires gamers, ', 'building communities.'],
    { x: 2.9759, y: 5.8704, w: 4.8612, h: 0.5991, color: C.white });
  iconTile(s, 8.0757, 5.9061, 0.559, C.gray40, C.white);
  heading(s, '532m', { x: 8.9526, y: 5.8117, w: 2.385, h: 0.8415, size: 44, color: C.white });
  body(s, ['Victory inspires gamers, ', 'building communities.'],
    { x: 10.4085, y: 5.8704, w: 4.8612, h: 0.5991, color: C.white });

  header(s, { word: C.white });
  pageNo(s, 5);
}

// 6 — About the journey.
function slide06(p) {
  const s = p.addSlide();
  box(s, { x: 6.6667, y: 3.75, w: 3.7136, h: 0.5729, fill: C.purple });
  heading(s, 'Victory      creates      legends', {
    x: 6.6667, y: 3.8345, w: 3.7136, h: 0.4039, size: 18, color: C.white, align: 'center',
  });

  tagPill(s, 0.4308, 1.4051);
  heading(s, ['About the Victory E-Sport ', 'Journey and Vision'], {
    x: 0.5413, y: 2.0243, w: 5.5282, h: 1.3127, size: 36,
  });
  body(s, 'A Victory E-Sport inspires gamers worldwide, building strong communities, fostering teamwork, and shaping the future of competitive gaming.',
    { x: 7.4024, y: 5.0312, w: 5.2886, h: 0.8616 });

  [['victory', 7.5061, 7.8022], ['champions', 9.2491, 9.5452]].forEach(([w, dot, tx]) => {
    box(s, { x: dot, y: 6.2812, w: 0.1528, h: 0.1528, fill: C.gray40, shape: 'ellipse' });
    heading(s, w, { x: tx, y: 6.1765, w: 1.2082, h: 0.4039, size: 18 });
  });

  statPlus(s, 10.7031, 2.0243, '620', '+');
  body(s, ['The Victory E-sport', 'Success'], { x: 10.7205, y: 2.8254, w: 2.2066, h: 0.5991 });
  header(s);
  pageNo(s, 6);
}

// 7 — Competitive landscape: a diagonal staircase of squares.
function slide07(p) {
  const s = p.addSlide();
  tagPill(s, 0.4377, 1.8183);
  heading(s, ['Competitive Gaming Landscape ', 'and Future Trends'], {
    x: 0.5302, y: 2.5508, w: 7.2476, h: 1.3127, size: 36,
  });
  heading(s, '258+', { x: 0.4817, y: 4.869, w: 3.4282, h: 2.0364, size: 115 });

  [[6.1111, -0.5476, C.ink], [7.9167, 1.2579, C.purple],
    [9.7222, 3.0635, C.purple], [11.5278, 4.869, C.ink]]
    .forEach(([x, y, fill]) => box(s, { x: x, y: y, w: 1.8056, h: 1.8056, fill: fill }));
  box(s, { x: 6.6748, y: 0.0879, w: 0.6781, h: 0.5346, fill: C.purple, shape: 'roundRect', radius: 0.17 });
  box(s, { x: 12.0913, y: 5.4326, w: 0.6786, h: 0.6786, fill: C.purple, shape: 'star5' });

  heading(s, '80%', { x: 7.1054, y: 1.74, w: 3.4282, h: 0.8415, size: 44, color: C.white, align: 'center' });
  heading(s, '80%', { x: 8.9109, y: 3.5455, w: 3.4282, h: 0.8415, size: 44, color: C.white, align: 'center' });
  body(s, ['Victory E-Sport builds ', 'champions.'], { x: 10.1547, y: 1.8612, w: 3.2828, h: 0.5991 });
  body(s, ['Victory E-Sport builds ', 'champions.'],
    { x: 6.0069, y: 3.6667, w: 3.2828, h: 0.5991, align: 'right', flipH: true });

  iconTile(s, 4.2559, 5.0387, 0.559, C.purple, C.white);
  body(s, ['The Victory E-Sport inspires gamers, building communities ',
    'and fostering teamwork for success worldwide.'], { x: 4.154, y: 5.8225, w: 4.8612, h: 0.5991 });
  header(s);
  pageNo(s, 7);
}

// 8 — The esport industry today: purple column + floating progress card.
function slide08(p) {
  const s = p.addSlide();
  box(s, { x: 4.9437, y: 0, w: 3.446, h: 6.0282, fill: C.purple });

  tagPill(s, 0.4308, 1.4051);
  heading(s, ['The Esport ', 'Industry Today'], { x: 0.5413, y: 2.0243, w: 5.5282, h: 1.7166, size: 48 });
  body(s, 'Victory E-sport Builds Global Champions', { x: 0.5413, y: 3.832, w: 4.0254, h: 0.3365 });
  iconTile(s, 0.6452, 5.0575, 0.559, C.ink, C.purple);
  body(s, ['The Victory E-Sport inspires ', 'gamers, building communities.'],
    { x: 0.5364, y: 5.8515, w: 2.5879, h: 0.5991, wrap: false });

  heading(s, '15.20m', { x: 9.2914, y: 2.1962, w: 3.3967, h: 1.1107, size: 60 });
  body(s, ['Victory E-Sport drives innovation, builds ', 'teamwork, inspires players, and shapes ', 'the future'],
    { x: 9.2914, y: 3.3069, w: 5.2886, h: 0.8616 });

  box(s, { x: 7.9127, y: 4.9152, w: 4.7754, h: 1.7849, fill: C.white });
  heading(s, 'Emotional Resilience', { x: 8.2145, y: 5.2141, w: 2.7828, h: 0.3029, size: 12 });
  box(s, { x: 8.2729, y: 5.7267, w: 4.0773, h: 0.597, fill: C.grayF2, alpha: 50 });
  box(s, { x: 8.2729, y: 5.7699, w: 3.1862, h: 0.5162, fill: C.ink });
  heading(s, '75%', { x: 8.3513, y: 5.8905, w: 0.6863, h: 0.3029, size: 12, color: C.white });
  box(s, { x: 11.6645, y: 5.2293, w: 0.6941, h: 0.264, fill: C.purple, shape: 'roundRect', radius: 0.132 });
  text(s, '+125', { x: 11.6022, y: 5.2267, w: 0.8186, h: 0.2693, size: 10, color: C.white, align: 'center' });
  header(s);
  pageNo(s, 8);
}

// 9 — Dark innovation slide: purple panel plus the smart-watch shot.
function slide09(p) {
  const s = p.addSlide();
  s.background = { color: C.ink };
  box(s, { x: 8.5873, y: 0, w: 4.746, h: SLIDE_H, fill: C.purple });

  box(s, { x: 0.642, y: 1.4255, w: 1.8759, h: 0.4026, fill: C.purple });
  heading(s, 'Victory creates legends', {
    x: 0.4305, y: 1.4832, w: 2.299, h: 0.3029, size: 12, color: C.white, align: 'center',
  });
  heading(s, 'Innovation and Technology Advancement Hub', {
    x: 0.5159, y: 2.0447, w: 5.5282, h: 1.3127, size: 36, color: C.white,
  });

  iconTile(s, 0.6459, 4.4547, 0.559, C.purple, C.white);
  heading(s, '120k', { x: 1.4879, y: 4.3578, w: 3.4282, h: 0.8415, size: 44, color: C.white });
  iconTile(s, 3.3838, 4.4547, 0.559, C.purple, C.white);
  heading(s, '532m', { x: 4.2258, y: 4.3578, w: 3.4282, h: 0.8415, size: 44, color: C.white });
  body(s, ['The Victory E-Sport inspires gamers, building communities ',
    'and fostering teamwork for success worldwide.'],
  { x: 0.5493, y: 5.4754, w: 4.8612, h: 0.5991, color: C.white });

  imagePlaceholder(s, { x: 6.544, y: 1.642, w: 5.551, h: 5.845, fill: '4C4C4C', alpha: 45, label: C.grayD9 });

  logo(s, C.white);
  navLinks(s, C.purpleSoft, { skipContact: true });
  pageNo(s, 9, C.purpleSoft);
}

// 10 — Community engagement, with a vertical purple label bar.
function slide10(p) {
  const s = p.addSlide();
  tagPill(s, 0.4377, 1.2963);
  heading(s, 'Community Engagement and Development Strategy', {
    x: 0.5302, y: 2.0289, w: 7.2476, h: 1.4473, size: 40,
  });
  iconTile(s, 0.6459, 4.6405, 0.559, C.purple, C.white);
  heading(s, '35,8%', { x: 1.4852, y: 4.4288, w: 5.8389, h: 1.1107, size: 60 });
  body(s, ['A Victory E-Sport inspires gamers worldwide, building ',
    'strong communities, fostering teamwork, and shaping ', 'the future of competitive gaming.'],
  { x: 0.535, y: 5.6305, w: 5.2886, h: 0.8616 });

  statPlus(s, 8.2964, 2.0289, '620', '+', { w: 1.7828 });
  body(s, ['The Victory E-Sport inspires ', 'gamers, building communities.'],
    { x: 9.7672, y: 2.0702, w: 3.2828, h: 0.5991 });

  box(s, { x: 4.9113, y: 4.7367, w: 3.0159, h: 0.4948, fill: C.purple, rotate: 270 });
  heading(s, 'Victory               e-sport               gamers', {
    x: 4.9113, y: 4.8327, w: 3.0159, h: 0.3029, size: 12, color: C.white, align: 'center', rotate: 270,
  });
  header(s);
  pageNo(s, 10);
}

// 11 — Future expansion plans: three phone mock-ups.
function slide11(p) {
  const s = p.addSlide();
  [[4.896, 0, 2.418, 3.585], [4.896, 3.842, 2.418, 3.658], [7.77, 1.245, 2.418, 4.934]]
    .forEach(([x, y, w, h]) => imagePlaceholder(s, { x: x, y: y, w: w, h: h, fill: 'ECECEC' }));

  tagPill(s, 0.4308, 1.3179);
  heading(s, ['The Future ', 'Expansion ', 'Plans'], { x: 0.5329, y: 1.9371, w: 5.5282, h: 3.1303, size: 60 });
  body(s, ['The Victory E-Sport inspires gamers, building ', 'communities, fostering teamwork, and shaping ',
    'the future of competitive gaming.'], { x: 0.5413, y: 5.3205, w: 5.2886, h: 0.8616 });

  statPlus(s, 10.6663, 2.8602, '620', '+');
  body(s, ['The Victory E-Sport ', 'inspires gamers, building ', 'communities.'],
    { x: 10.6663, y: 3.7782, w: 3.2828, h: 0.8616 });
  header(s);
  pageNo(s, 11);
}

// 12 — STRATEGY / FOR GROWTH, the two halves pushed to opposite edges.
function slide12(p) {
  const s = p.addSlide();
  tagPill(s, 0.4377, 1.4908);
  heading(s, [[{ t: 'Strategy' }], [{ t: 'FOR GROWTH ', align: 'right' }]],
    { x: 0.5302, y: 2.2233, w: 7.8638, h: 3.063, size: 88 });

  box(s, { x: 5.2273, y: 2.8883, w: 5.0455, h: 0.4948, fill: C.purple });
  heading(s, 'Victory\tE-Sport\tinspires\tgamers', {
    x: 5.2273, y: 2.9674, w: 5.0455, h: 0.3366, size: 14, color: C.white, align: 'center',
  });

  statPlus(s, 10.6716, 1.6215, '620', '+ ');
  body(s, ['The Victory E-Sport ', [{ t: 'inspires gamers', bold: true }]],
    { x: 10.6917, y: 2.3721, w: 2.2831, h: 0.5161, size: 10 });

  iconTile(s, 9.6528, 5.439, 0.559, C.purple, C.white);
  heading(s, '$2.158', { x: 10.3683, y: 5.2307, w: 5.8389, h: 1.1107, size: 60 });
  body(s, 'Victory E-Sport unites inspires success', { x: 9.5717, y: 6.3349, w: 4.8612, h: 0.3365 });
  body(s, ['Victory inspires gamers, building strong ', 'communities, and shaping the future.'],
    { x: 5.2273, y: 6.0723, w: 5.2886, h: 0.5991 });
  header(s);
  pageNo(s, 12);
}

// 13 — Dark team slide: controller inside a halo of translucent circles.
function slide13(p) {
  const s = p.addSlide();
  s.background = { color: C.ink };
  box(s, { x: 4.625, y: 2.8134, w: 4.0833, h: 4.0833, fill: C.gray40, alpha: 40, shape: 'ellipse' });
  box(s, { x: 5.2604, y: 3.4488, w: 2.8125, h: 2.8125, fill: C.gray50, alpha: 60, shape: 'ellipse' });
  imagePlaceholder(s, { x: 5.034, y: 3.758, w: 3.319, h: 2.421, fill: 'ABBAC5', alpha: 35, label: '3B3B3B' });

  heading(s, 'MEET OUR VICTORY TEAM', {
    x: 4.2491, y: 1.1019, w: 4.8352, h: 0.7068, size: 36, color: C.white, align: 'center',
  });
  body(s, 'Victory E-Sport inspires gamers worldwide, building strong communities,',
    { x: 2.9375, y: 1.8971, w: 7.4583, h: 0.3397, color: C.white, align: 'center' });

  [[9.4636, 2.9947, 9.3659, 3.7541, 'left'], [9.4636, 5.1111, 9.3659, 5.8704, 'left'],
    [3.3404, 2.9947, -0.9914, 3.7541, 'right'], [3.3404, 5.1111, -0.9914, 5.8704, 'right']]
    .forEach(([ix, iy, tx, ty, align]) => {
      iconTile(s, ix, iy, 0.4317, C.purple, C.white);
      body(s, ['Victory E-Sport unites gamers and ', 'inspires success worldwide today.'],
        { x: tx, y: ty, w: 4.8612, h: 0.5991, color: C.white, align: align, flipH: align === 'right' });
    });

  header(s, { word: C.white });
  pageNo(s, 13);
}

// 14 — Our core strategy for growth.
function slide14(p) {
  const s = p.addSlide();
  tagPill(s, 0.4377, 1.4908);
  heading(s, ['Our Core Strategy ', 'for Growth'], { x: 0.5302, y: 2.2233, w: 6.1365, h: 1.7166, size: 48 });
  body(s, 'The Victory E-Sport inspires gamers.', { x: 0.5413, y: 3.9399, w: 5.2886, h: 0.3365 });
  body(s, ['A Victory E-Sport inspires gamers worldwide, building ',
    'strong communities, fostering teamwork, and shaping ', 'the future of competitive gaming.'],
  { x: 0.535, y: 5.6987, w: 5.2886, h: 0.8616 });

  statPlus(s, 10.6663, 2.4095, '620', '+');
  body(s, ['The Victory E-Sport ', 'inspires gamers, building ', 'communities.'],
    { x: 10.6663, y: 3.3275, w: 3.2828, h: 0.8616 });
  iconTile(s, 9.5415, 4.5694, 0.7428, C.purple, C.white);

  heading(s, '$2.158.90', { x: 5.5204, y: 5.1196, w: 5.8389, h: 1.1107, size: 60 });
  body(s, 'Victory E-Sport unites inspires success', { x: 5.5204, y: 6.2238, w: 4.8612, h: 0.3365 });
  header(s);
}

// 15 — Player performance insights: a 7x4 colour-swatch grid.
function slide15(p) {
  const s = p.addSlide();
  tagPill(s, 0.4377, 1.6077);
  heading(s, ['The Player Performance ', 'Insights Report'], { x: 0.5302, y: 2.3403, w: 7.2476, h: 1.4473, size: 40 });
  body(s, ['The Victory E-Sport inspires gamers, building communities ',
    'and fostering teamwork for success worldwide.'], { x: 0.544, y: 4.1176, w: 4.8612, h: 0.5991 });
  heading(s, '35.8%', { x: 1.9722, y: 5.1861, w: 4.4107, h: 2.0364, size: 115, align: 'right' });
  body(s, [[{ t: 'The Victory E-Sport ', bold: true }, { t: 'inspires ' }], 'gamers, building communities '],
    { x: 0.544, y: 6.0015, w: 4.8612, h: 0.5161, size: 10 });

  // f = light grey, p = purple, d = near black — one character per cell.
  const rows = ['ffpdfdf', 'pdfffff', 'fppffpd', 'ffpddfp'];
  const swatch = { f: C.grayF2, p: C.purple, d: C.ink };
  const colX = [6.9565, 7.9445, 8.9325, 9.9205, 10.9085, 11.8965, 12.9204];
  const rowY = [1.6077, 2.6728, 3.7378, 4.8028];
  rows.forEach((row, r) => row.split('').forEach((code, c) => {
    box(s, { x: colX[c], y: rowY[r], w: 0.7833, h: 0.7833, fill: swatch[code] });
  }));

  iconTile(s, 6.9565, 5.8679, 0.7833, C.ink, C.white);
  heading(s, 'Victory E-Sport unites gamers and inspires future success', {
    x: 7.943, y: 6.0542, w: 5.9969, h: 0.4106, size: 16,
  });
  header(s);
  pageNo(s, 15);
}

// 16 — BREAK / SLIDE divider.
function slide16(p) {
  const s = p.addSlide();
  darkBackdrop(s);
  heading(s, 'Break', { x: 0.4488, y: 1.9997, w: 6.4549, h: 2.8947, size: 166, color: C.purple });
  heading(s, 'slide', { x: 6.0794, y: 3.569, w: 6.7767, h: 2.8947, size: 166, color: C.white, align: 'right' });
  heading(s, 'future of competitive gaming', { x: 0.5398, y: 4.3297, w: 4.0296, h: 0.472, size: 18, color: C.white });
  heading(s, 'The Victory E-Sport', {
    x: 9.9375, y: 1.6035, w: 2.8576, h: 0.4721, size: 18, color: C.purple, align: 'right',
  });
  body(s, 'The Victory E-Sport inspires gamers, building communities and fostering teamwork.',
    { x: 9.5177, y: 2.1027, w: 3.2828, h: 0.5161, size: 10, color: C.white, align: 'right' });
  tiltChip(s, 10.3619, 5.2389, '60 minutes', C.purple, C.ink);
  body(s, [[{ t: 'Victory ' }, { t: 'Building Communities', bold: true }]],
    { x: 0.5398, y: 6.6755, w: 3.2828, h: 0.3365, color: C.gray80 });

  header(s, { word: C.white, nav: C.gray80 });
  pageNo(s, 16);
}

// 17 — Two money stats over the community-strategy block.
function slide17(p) {
  const s = p.addSlide();
  [['$2.158', 5.2684], ['$5.239', 9.1816]].forEach(([value, x]) => {
    heading(s, value, { x: x, y: 1.1732, w: 3.4885, h: 1.1107, size: 60 });
  });
  [5.2934, 9.2065].forEach((x) => {
    body(s, ['Victory inspires gamers, building strong ', 'communities and fostering teamwork.'],
      { x: x, y: 2.4184, w: 5.2886, h: 0.5991 });
  });

  tagPill(s, 5.1871, 3.75);
  heading(s, 'Community Engagement and Development Strategy', {
    x: 5.2795, y: 4.4826, w: 7.2476, h: 1.4473, size: 40,
  });
  ['Competitive Innovation', 'Collaborative Strategy', 'Transformational Growth'].forEach((label, i) => {
    body(s, label, { x: 5.2934 + i * 2.6088, y: 6.4182, w: 2.6087, h: 0.3365, color: C.grayBF });
  });

  navLinks(s, C.grayD9);
  pageNo(s, 17);
}

// 18 — Innovation hub: purple sidebar and the laptop shot.
function slide18(p) {
  const s = p.addSlide();
  s.background = { color: C.white };
  box(s, { x: 0, y: 0, w: 3.5906, h: SLIDE_H, fill: C.purple });
  imagePlaceholder(s, { x: 1.42, y: 0.99, w: 4.259, h: 5.554, fill: '2B2B2B', alpha: 20, label: C.white });

  box(s, { x: 6.7251, y: 1.4255, w: 1.8759, h: 0.4026, fill: C.purple });
  heading(s, 'Victory creates legends', {
    x: 6.5135, y: 1.4832, w: 2.299, h: 0.3029, size: 12, color: C.white, align: 'center',
  });
  heading(s, 'Innovation and Technology Advancement Hub', {
    x: 6.599, y: 2.0447, w: 5.5282, h: 1.3127, size: 36,
  });
  iconTile(s, 6.7008, 4.4547, 0.559, C.purple, C.white);
  heading(s, '120k', { x: 7.5376, y: 4.3578, w: 3.4282, h: 0.8415, size: 44 });
  iconTile(s, 9.4387, 4.4547, 0.559, C.purple, C.white);
  heading(s, '532m', { x: 10.2755, y: 4.3578, w: 3.4282, h: 0.8415, size: 44 });
  body(s, ['The Victory E-Sport inspires gamers, building communities ',
    'and fostering teamwork for success worldwide.'], { x: 6.599, y: 5.4754, w: 4.8612, h: 0.5991 });

  box(s, { x: 0.6426, y: 0.4093, w: 0.2196, h: 0.2396, fill: C.white, shape: 'hexagon' });
  heading(s, 'Vicotryx', { x: 0.9019, y: 0.3777, w: 0.9757, h: 0.3029, size: 12, color: C.white });
  navLinks(s, C.grayD9);
  pageNo(s, 18);
}

// 19 — Contact details.
function slide19(p) {
  const s = p.addSlide();
  tagPill(s, 0.4308, 1.4051);
  heading(s, 'stay Connected and Get in Touch', { x: 0.5413, y: 2.0243, w: 9.9796, h: 0.9088, size: 48 });
  body(s, 'Victory E-sport Builds Global Champions', { x: 0.5413, y: 3.09, w: 4.0254, h: 0.3365 });
  iconTile(s, 0.6424, 4.5669, 0.4866, C.purple, C.white);

  [['Phone Number\t: +123-4567-8900', 5.3686],
    ['Email Address\t: yourmail@addmail.com', 5.8026],
    ['Website Link\t: www.yourwebsite.com', 6.2366]].forEach(([line, y]) => {
    text(s, line, { x: 0.5329, y: y, w: 5.7338, h: 0.3029, size: 12 });
  });
  body(s, ['The Victory E-Sport inspires gamers, ', 'building communities.'],
    { x: 9.7583, y: 2.125, w: 3.7208, h: 0.5991 });

  box(s, { x: 11.0538, y: 5.2205, w: 3.9167, h: 0.6424, fill: C.purple, rotate: 270 });
  heading(s, 'Victory\tE-Sport\tgamers', {
    x: 11.1538, y: 5.3734, w: 3.7167, h: 0.3366, size: 14, color: C.white, align: 'center', rotate: 270,
  });
  header(s);
}

// 20 — Thank you.
function slide20(p) {
  const s = p.addSlide();
  darkBackdrop(s);
  heading(s, 'Thank you', {
    x: 0, y: 2.5523, w: SLIDE_W, h: 3.3322, size: 192, color: C.purple, align: 'center',
  });
  heading(s, 'Victory Esport Unites Players and Builds Legends', {
    x: 1.7986, y: 5.417, w: 9.7361, h: 0.4376, size: 20, color: C.purple, align: 'center',
  });
  wordStrip(s);
  tiltChip(s, 8.865, 3.7431, 'Game force', C.gray40, C.white);
  logo(s, C.white);
  navLinks(s, C.gray80);
  heading(s, 'www.yourwebsite.com', {
    x: 0.5433, y: 6.8194, w: 2.0262, h: 0.3029, size: 12, color: C.gray80,
  });
  heading(s, 'E-sport presentation / 2029', {
    x: 10.7743, y: 6.8194, w: 2.0262, h: 0.3029, size: 12, color: C.gray80, align: 'right',
  });
}

/* ================================================================== *
 * Build
 * ================================================================== */
function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'WIDE';
  pptx.theme = { headFontFace: F.head, bodyFontFace: F.body };
  pptx.title = 'Victory Arena';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach((fn) => fn(pptx));

  return pptx.writeFile({
    fileName: path.join(__dirname, '0ba7fedc-b5c7-468e-8bb8-d2e46a629aa1_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f));
