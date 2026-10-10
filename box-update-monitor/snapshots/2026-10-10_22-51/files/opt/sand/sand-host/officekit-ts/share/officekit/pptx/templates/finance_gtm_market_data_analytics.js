/**
 * "Go To Market" pitch deck — recreated with pptxgenjs.
 *
 * Slide size 13.333 x 7.5 in (16:9). Brand green #00B04E on white, headline
 * face "Funnel Display", body face "Montserrat".
 *
 * Photographs in the original are empty picture placeholders (they render as
 * blank cards) and every logo / icon is vector art, so nothing here is a raster
 * embed: the brand mark is redrawn with custom geometry and the small pictogram
 * icons become coloured discs.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Palette, type and reusable effects
 * ------------------------------------------------------------------ */

const GREEN = '00B04E';
const GREEN_DARK = '005827';
const DARK = '262626'; // tx1 lum 85%
const WHITE = 'FFFFFF';
const BLACK = '000000';
const GREY_05 = 'F2F2F2';
const GREY_15 = 'D9D9D9';
const GREY_35 = 'A6A6A6';

const HEAD = 'Funnel Display';
const BODY = 'Montserrat';

const TITLE_SIZE = 40;
const BODY_SIZE = 12;

// The deck uses one soft card shadow, thrown down-right on most slides and
// up-right on the two "hero" panels. pptxgenjs rewrites the shadow object it is
// handed, so shapes are always given a private copy (see `shadow()`).
const CARD_SHADOW = { type: 'outer', color: BLACK, opacity: 0.05, blur: 10, offset: 10, angle: 40 };
const CARD_SHADOW_UP = { type: 'outer', color: BLACK, opacity: 0.05, blur: 10, offset: 10, angle: 135 };
const PHONE_SHADOW = { type: 'outer', color: BLACK, opacity: 0.2, blur: 35, offset: 30, angle: 45 };

function shadow(preset) {
  return preset ? Object.assign({}, preset) : undefined;
}

const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consec tetur';
const LOREM_CARD = 'Lorem ipsum dolor sit amet, consec tetur adipiscing elit. Praesent erat';
const LOREM_ROW = 'Lorem ipsum dolor sit amet, consec tetur adipiscing';
const LOREM_INTRO = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent erat quam, feugiat sit amet justo non, interdum dignissim';
const LOREM_LONG = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Donec eu porta augue, at eleifend est. Vivamus ornare lorem vitae ante porta tincidunt. Morbi facilisis';
const LOREM_FUNNEL = 'Lorem ipsum dolor sit amet, consec tetur adipiscing elit';
const LOREM_CARD_SHORT = 'Lorem ipsum dolor sit amet, consec tetur adipiscing elit. ';
const LOREM_VALUE = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Donec eu porta augue, at eleifend est. Vivamus orna re lorem vitae ante porta';

/* ------------------------------------------------------------------ *
 * Drawing helpers
 * ------------------------------------------------------------------ */

/** Plain filled rectangle. */
function rect(slide, o) {
  slide.addShape('rect', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: { color: o.fill, transparency: o.transparency },
    line: o.line || { type: 'none' },
    shadow: shadow(o.shadow),
  });
}

/** Rounded rectangle; `r` is the corner radius as a fraction of the short side. */
function roundRect(slide, o) {
  slide.addShape('roundRect', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    rectRadius: o.r === undefined ? 0.5 * Math.min(o.w, o.h) : o.r,
    fill: o.fill ? { color: o.fill, transparency: o.transparency } : { type: 'none' },
    line: o.line || { type: 'none' },
    shadow: shadow(o.shadow),
  });
}

/**
 * Text block. Runs are given as [text, color] pairs so a two-tone headline
 * reads as data; everything else comes from the shape-level options.
 */
function text(slide, parts, o) {
  const runs = (typeof parts === 'string' ? [[parts, o.color || DARK]] : parts)
    .map(([t, color, extra]) => ({ text: t, options: Object.assign({ color: color || DARK }, extra) }));
  slide.addText(runs, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: o.font || BODY,
    fontSize: o.size || BODY_SIZE,
    bold: o.bold,
    align: o.align || 'left',
    valign: o.valign || 'top',
    lineSpacingMultiple: o.ls,
    paraSpaceAfter: o.after,
    charSpacing: o.charSpacing,
  });
}

/** Two-tone 40pt headline (the deck's standard slide title). */
function heading(slide, parts, o) {
  text(slide, parts, Object.assign({ font: HEAD, size: TITLE_SIZE, ls: 0.9 }, o));
}

/** 12pt Montserrat paragraph with the deck's 125% leading. */
function para(slide, body, o) {
  text(slide, body, Object.assign({ size: BODY_SIZE, ls: 1.25, after: 8 }, o));
}

/** Bold 12pt label ("Subtitle Here", "Your Subtitle Here", ...). */
function label(slide, body, o) {
  para(slide, body, Object.assign({ bold: true, h: 0.332 }, o));
}

/** Oversized statistic ("90%", "12.5M", "$435K", ...). */
function stat(slide, body, o) {
  text(slide, body, Object.assign({ font: HEAD, ls: 0.9 }, o));
}

/** Wingdings-style square bullets, one paragraph per list entry. */
function bulletList(slide, items, o) {
  slide.addText(
    items.map(t => ({ text: t, options: { bullet: { characterCode: '25AA', indent: 22.5 } } })),
    {
      x: o.x, y: o.y, w: o.w, h: o.h,
      fontFace: BODY, fontSize: o.size || BODY_SIZE, color: DARK,
      valign: 'top', lineSpacingMultiple: 2.5, paraSpaceAfter: 4,
    }
  );
}

/**
 * Soft vertical wash standing in for the deck's black gradient overlays.
 * pptxgenjs cannot emit a gradient fill, so the band is painted as a stack of
 * translucent slices. `stops` are [fractionOfHeight, opacityPercent] pairs read
 * top-to-bottom and linearly interpolated between.
 */
function gradientBand(slide, o) {
  const steps = 72;
  const stops = o.stops;
  const opacityAt = t => {
    if (t <= stops[0][0]) return stops[0][1];
    for (let i = 1; i < stops.length; i++) {
      if (t <= stops[i][0]) {
        const [p0, a0] = stops[i - 1];
        const [p1, a1] = stops[i];
        return a0 + ((a1 - a0) * (t - p0)) / (p1 - p0);
      }
    }
    return stops[stops.length - 1][1];
  };
  for (let i = 0; i < steps; i++) {
    const opacity = opacityAt((i + 0.5) / steps);
    if (opacity < 0.5) continue;
    slide.addShape('rect', {
      x: o.x, y: o.y + (o.h * i) / steps, w: o.w, h: o.h / steps + 0.004,
      fill: { color: BLACK, transparency: Math.round(100 - opacity) },
      line: { type: 'none' },
    });
  }
}

/* --- brand mark ---------------------------------------------------- */

// The "G" monogram, traced from the brand SVG (43 x 40 user units).
// Each entry is a closed sub-path: ['M'|'L', x, y] or ['C', x, y, c1x, c1y, c2x, c2y].
const LOGO_W = 43;
const LOGO_H = 40;
const LOGO_PATHS = [
  [['M', 3.30, 17.13], ['C', 7.21, 14.00, 3.71, 15.30, 5.33, 14.00], ['L', 12.00, 14.00],
   ['L', 8.00, 32.00], ['L', 4.99, 32.00], ['C', 1.08, 27.13, 2.43, 32.00, 0.53, 29.63],
   ['L', 3.30, 17.13]],
  [['M', 39.29, 8.00], ['L', 21.29, 8.00], ['L', 20.68, 10.84],
   ['C', 16.77, 14.00, 20.28, 12.68, 18.65, 14.00], ['L', 12.00, 14.00], ['L', 14.32, 3.16],
   ['C', 18.23, 0.00, 14.72, 1.32, 16.35, 0.00], ['L', 36.05, 0.00],
   ['C', 39.96, 4.84, 38.60, 0.00, 40.50, 2.35], ['L', 39.29, 8.00]],
  [['M', 30.96, 36.84], ['C', 27.05, 40.00, 30.57, 38.68, 28.94, 40.00], ['L', 11.23, 40.00],
   ['C', 7.32, 35.16, 8.69, 40.00, 6.79, 37.65], ['L', 8.00, 32.00], ['L', 32.00, 32.00],
   ['L', 30.96, 36.84]],
  [['M', 38.63, 17.00], ['C', 42.51, 21.97, 41.23, 17.00, 43.00, 19.45], ['L', 40.76, 28.97],
   ['C', 36.88, 32.00, 40.31, 30.75, 38.71, 32.00], ['L', 32.00, 32.00], ['L', 33.75, 25.00],
   ['L', 14.00, 25.00], ['L', 15.04, 20.16], ['C', 18.95, 17.00, 15.43, 18.32, 17.06, 17.00],
   ['L', 38.63, 17.00]],
];

/** Draw the monogram at (x, y) with the given width; height follows the artwork ratio. */
function logo(slide, x, y, w, color, transparency) {
  const h = (w * LOGO_H) / LOGO_W;
  const sx = v => (v / LOGO_W) * w;
  const sy = v => (v / LOGO_H) * h;
  LOGO_PATHS.forEach(cmds => {
    const points = cmds.map(c => (c[0] === 'C'
      ? { x: sx(c[1]), y: sy(c[2]), curve: { type: 'cubic', x1: sx(c[3]), y1: sy(c[4]), x2: sx(c[5]), y2: sy(c[6]) } }
      : { x: sx(c[1]), y: sy(c[2]), moveTo: c[0] === 'M' }));
    points.push({ close: true });
    slide.addShape('custGeom', {
      x, y, w, h, points,
      fill: { color: color || GREEN, transparency },
      line: { type: 'none' },
    });
  });
}

/**
 * Stand-in for the deck's pictogram icons (piggy bank, light bulb, gear, ...):
 * a native shape in the icon's own colour, optionally carrying a glyph.
 * `shape` defaults to a disc; trend arrows use 'swooshArrow'.
 */
function icon(slide, o) {
  slide.addShape(o.shape || 'ellipse', {
    x: o.x, y: o.y, w: o.size, h: o.size,
    flipV: o.flipV,
    fill: { color: o.color }, line: { type: 'none' },
  });
  if (o.glyph) {
    slide.addText(o.glyph, {
      x: o.x, y: o.y, w: o.size, h: o.size,
      fontFace: HEAD, fontSize: Math.round(o.size * 46), bold: true,
      color: o.glyphColor || WHITE, align: 'center', valign: 'middle',
      margin: 0,
    });
  }
}

/* --- master furniture ---------------------------------------------- */

/**
 * Header / footer repeated on every slide but the cover: brand mark and name
 * top-left, "Presentation Template" top-right, "Go To Market" and the page
 * number along the bottom. A few slides restate it in white over dark artwork.
 */
function chrome(slide, pageNo, color, parts) {
  const c = color || DARK;
  const want = k => !parts || parts.indexOf(k) >= 0;
  if (want('brand')) {
    logo(slide, 0.423, 0.226, 0.296, c);
    text(slide, [['Company Brand', c]], { x: 0.747, y: 0.226, w: 4.306, h: 0.269, font: HEAD, size: 10, valign: 'middle' });
  }
  if (want('template')) {
    text(slide, [['Presentation Template', c]], { x: 8.163, y: 0.235, w: 4.699, h: 0.252, size: 9, align: 'right', valign: 'middle' });
  }
  if (want('footer')) {
    text(slide, [['Go To Market', c]], { x: 0.471, y: 7.013, w: 4.699, h: 0.252, size: 9, valign: 'middle' });
  }
  if (want('page')) {
    text(slide, [['page ', c], [String(pageNo), c, { bold: true }]],
      { x: 11.5, y: 7.013, w: 1.632, h: 0.252, size: 9, align: 'right', valign: 'middle' });
  }
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

/** 1 — Cover: photo band with a white title card sitting over its lower edge. */
function slide01(pres) {
  const s = pres.addSlide();
  gradientBand(s, { x: 1.549, y: 0.757, w: 11.024, h: 5.512, stops: [[0, 60], [0.6, 0]] });
  rect(s, { x: 0.76, y: 4.381, w: 6.299, h: 2.362, fill: WHITE, shadow: CARD_SHADOW_UP });
  rect(s, { x: 0.76, y: 4.381, w: 0.492, h: 2.362, fill: GREEN });

  logo(s, 1.811, 5.054, 0.576, GREEN);
  heading(s, [['o', GREEN], [' To Market', DARK]], { x: 2.32, y: 4.961, w: 4.479, h: 0.828, size: 48 });
  para(s, 'Real Time Market Data, Financial News, and Analytics', { x: 1.811, y: 5.831, w: 4.971, h: 0.332 });

  text(s, [['2030', WHITE]], { x: 9.885, y: 1.153, w: 2.39, h: 0.438, font: HEAD, size: 20, align: 'right' });
  text(s, [['Biggest Market Share Holder', WHITE]], { x: 9.885, y: 1.591, w: 2.39, h: 0.774, font: HEAD, size: 20, align: 'right' });
}

/** 2 — Agenda: five numbered rows beside a dark photo column. */
function slide02(pres) {
  const s = pres.addSlide();
  chrome(s, 2);
  gradientBand(s, { x: 0, y: 0.772, w: 4.724, h: 6.728, stops: [[0.4, 0], [1, 100]] });

  heading(s, [['Today\u2019s ', DARK], ['Agenda', GREEN]], { x: 5.41, y: 0.772, w: 6.132, h: 0.707 });

  ['01', '02', '03', '04', '05'].forEach((num, i) => {
    const y = 1.941 + i * 1.053;
    rect(s, { x: 5.514, y, w: 1.181, h: 0.591, fill: GREEN });
    text(s, [[num, WHITE]], { x: 5.809, y: y + 0.124, w: 0.591, h: 0.343, font: HEAD, size: 16, align: 'center', ls: 0.9 });
    para(s, [[LOREM_ROW, DARK]], { x: 6.99, y: y + 0.144, w: 4.979, h: 0.303, bold: true, ls: undefined, after: undefined });
  });

  text(s, [['\u201C', WHITE]], { x: 0.76, y: 5.095, w: 2.39, h: 0.942, font: HEAD, size: 50 });
  text(s, [['Aiming to Achieve Biggest Market Share', WHITE]], { x: 0.76, y: 5.632, w: 2.39, h: 1.111, font: HEAD, size: 20 });
}

/** 3 — Demand generation: 2x2 grid of soft cards, one of them inverted. */
function slide03(pres) {
  const s = pres.addSlide();
  chrome(s, 3);
  heading(s, [['Demand Generation ', DARK], ['Strategy', GREEN]], { x: 0.76, y: 0.757, w: 5.906, h: 0.374 });
  para(s, LOREM_INTRO, { x: 0.76, y: 2.249, w: 5.906, h: 0.584 });

  const cards = [
    { x: 0.76, y: 3.297, green: false },
    { x: 4.698, y: 3.297, green: false },
    { x: 0.76, y: 5.168, green: true },
    { x: 4.698, y: 5.168, green: false },
  ];
  cards.forEach(c => {
    rect(s, { x: c.x, y: c.y, w: 3.543, h: 1.575, fill: c.green ? GREEN : WHITE, shadow: c.green ? undefined : CARD_SHADOW });
    label(s, [['Your Subtitle Here', c.green ? WHITE : GREEN]], { x: c.x + 0.164, y: c.y + 0.294, w: 3.379 });
    para(s, [[LOREM_CARD, c.green ? WHITE : DARK]], { x: c.x + 0.164, y: c.y + 0.697, w: 3.379, h: 0.584 });
  });
}

/** 4 — Staircase of four percentage columns, alternating green and grey. */
function slide04(pres) {
  const s = pres.addSlide();
  chrome(s, 4);
  heading(s, [['Driving Sustainable ', DARK], ['Growth', GREEN]], { x: 6.667, y: 0.757, w: 5.906, h: 1.313, align: 'right' });
  para(s, LOREM_INTRO, { x: 6.667, y: 2.249, w: 5.906, h: 0.584, align: 'right' });

  // Drawn right-to-left so each taller column overlaps the one before it.
  const columns = [
    { x: 9.817, y: 4.381, h: 2.362, fill: GREY_05, value: '36%', labelY: 4.566, statY: 4.975 },
    { x: 6.798, y: 3.397, h: 3.346, fill: GREEN, value: '55%', labelY: 3.659, statY: 4.068 },
    { x: 3.779, y: 1.625, h: 5.118, fill: GREY_05, value: '72%', labelY: 2.026, statY: 2.436 },
    { x: 0.76, y: 0.757, h: 5.986, fill: GREEN, value: '90%', labelY: 1.227, statY: 1.636 },
  ];
  columns.forEach(c => {
    const ink = c.fill === GREEN ? WHITE : DARK;
    const tx = c.x + 0.209;
    rect(s, { x: c.x, y: c.y, w: 2.756, h: c.h, fill: c.fill });
    label(s, [['Subtitle Here', ink]], { x: tx, y: c.labelY, w: 2.403 });
    stat(s, [[c.value, ink]], { x: tx, y: c.statY, w: 2.399, h: 0.942, size: 54 });
    para(s, [[LOREM_SHORT, ink]], { x: tx, y: 5.822, w: 2.399, h: 0.584 });
  });

  logo(s, 1.068, 4.994, 0.317, WHITE);
  label(s, [['Current State', WHITE]], { x: 1.411, y: 4.975, w: 1.96 });
}

/** 5 — Team grid: four portrait cards under numbered badges. */
function slide05(pres) {
  const s = pres.addSlide();
  chrome(s, 5);
  rect(s, { x: 5.289, y: 0.757, w: 2.756, h: 0.591, fill: GREEN });
  label(s, [['Growth Inspired Minds', WHITE]], { x: 5.486, y: 0.886, w: 2.362, align: 'center' });
  heading(s, [['Meet the Special ', DARK], ['Team', GREEN]], { x: 2.583, y: 1.511, w: 8.167, h: 0.707, align: 'center' });
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent erat', { x: 3.413, y: 2.275, w: 6.508, h: 0.332, align: 'center' });

  [0, 1, 2, 3].forEach(i => {
    const cardX = 0.76 + i * 3.1505;
    rect(s, { x: cardX, y: 3.297, w: 2.362, h: 2.362, fill: WHITE, shadow: CARD_SHADOW });
    rect(s, { x: cardX + 0.432, y: 3.1, w: 0.591, h: 0.591, fill: GREEN });
    text(s, [['0' + (i + 1), WHITE]], { x: cardX + 0.432, y: 3.224, w: 0.591, h: 0.343, font: HEAD, size: 16, align: 'center', ls: 0.9 });
    para(s, [['Your Name Here', DARK]], { x: cardX - 0.02, y: 5.985, w: 2.383, h: 0.369, size: 14, bold: true, font: HEAD, align: 'center' });
    para(s, [['Your Title Here', DARK]], { x: cardX - 0.02, y: 6.406, w: 2.383, h: 0.313, size: 11, align: 'center' });
  });
}

/** 6 — Market potential: headline stat plus a green two-column panel. */
function slide06(pres) {
  const s = pres.addSlide();
  chrome(s, 6);
  heading(s, [['Unlocking Market ', DARK], ['Potential', GREEN]], { x: 6.667, y: 0.994, w: 5.906, h: 1.313 });
  stat(s, [['200+', GREEN]], { x: 6.667, y: 2.676, w: 2.399, h: 0.942, size: 54 });
  para(s, [['Regional Market', DARK]], { x: 8.849, y: 2.756, w: 1.454, h: 0.584, bold: true });
  para(s, LOREM_SHORT, { x: 10.143, y: 2.756, w: 2.399, h: 0.584 });

  rect(s, { x: 3.951, y: 3.987, w: 8.622, h: 2.52, fill: GREEN });
  [{ x: 4.652, tx: 4.597 }, { x: 8.603, tx: 8.548 }].forEach(col => {
    icon(s, { x: col.x, y: 4.377, size: 0.492, color: WHITE, glyph: '$', glyphColor: GREEN });
    label(s, [['Your Subtitle Here', WHITE]], { x: col.tx, y: 5.128, w: 3.379 });
    para(s, [[LOREM_CARD, WHITE]], { x: col.tx, y: 5.531, w: 3.379, h: 0.584 });
  });

  chrome(s, 6, WHITE, ['brand', 'footer']);
}

/** 7 — Launch metrics: three green tiles beside copy and a bullet grid. */
function slide07(pres) {
  const s = pres.addSlide();
  chrome(s, 7);
  heading(s, [['Launching with ', DARK], ['Confidence', GREEN]], { x: 0.76, y: 1.837, w: 4.719, h: 1.313 });
  para(s, LOREM_INTRO, { x: 0.76, y: 3.342, w: 4.719, h: 0.837 });
  bulletList(s, ['First Point', 'Second Point'], { x: 0.76, y: 4.581, w: 1.948, h: 1.082 });
  bulletList(s, ['Third Point', 'Fourth Point'], { x: 2.765, y: 4.581, w: 1.948, h: 1.082 });

  const tiles = [
    { x: 6.274, y: 0.757, w: 3.543, h: 2.051, value: '12.5M', size: 54, valueY: 1.116, valueW: 2.753, labelY: 2.117, tx: 6.667, lx: 6.666, lw: 2.758 },
    { x: 10.211, y: 0.757, w: 2.362, h: 3.543, value: '200K', size: 36, valueY: 1.116, valueW: 2.186, labelY: 3.584, tx: 10.387, lx: 10.387, lw: 2.186 },
    { x: 10.211, y: 4.692, w: 2.362, h: 2.051, value: '410K', size: 36, valueY: 4.971, valueW: 2.186, labelY: 6.052, tx: 10.387, lx: 10.387, lw: 2.186 },
    { x: 6.274, y: 3.2, w: 3.543, h: 3.543, value: '36.8M', size: 54, valueY: 3.584, valueW: 2.753, labelY: 6.052, tx: 6.667, lx: 6.666, lw: 2.758 },
  ];
  tiles.forEach(t => {
    rect(s, { x: t.x, y: t.y, w: t.w, h: t.h, fill: GREEN });
    stat(s, [[t.value, WHITE]], { x: t.tx, y: t.valueY, w: t.valueW, h: t.size === 54 ? 0.919 : 0.646, size: t.size });
    label(s, [['Subtitle Here', WHITE]], { x: t.lx, y: t.labelY, w: t.lw });
  });

  logo(s, 6.777, 4.544, 0.317, WHITE);
  label(s, [['Growth Target', WHITE]], { x: 7.12, y: 4.526, w: 1.96 });
}

/** 8 — Distribution: two icon tiles, the first inside a green banner. */
function slide08(pres) {
  const s = pres.addSlide();
  chrome(s, 8);
  rect(s, { x: 6.667, y: 0.757, w: 6.667, h: 2.993, fill: GREEN });
  heading(s, [['Distribution', GREEN], [' & Channel Strategy', DARK]], { x: 0.76, y: 1.215, w: 5.906, h: 1.313 });
  para(s, LOREM_INTRO, { x: 0.76, y: 2.707, w: 5.906, h: 0.584 });

  const rows = [
    { y: 1.663, tile: WHITE, iconColor: GREEN, glyph: '!', glyphColor: WHITE, ink: WHITE },
    { y: 4.656, tile: GREEN, iconColor: WHITE, glyph: '$', glyphColor: GREEN, ink: DARK },
  ];
  rows.forEach(r => {
    rect(s, { x: 7.492, y: r.y, w: 1.181, h: 1.181, fill: r.tile });
    icon(s, { x: 7.787, y: r.y + 0.295, size: 0.591, color: r.iconColor, glyph: r.glyph, glyphColor: r.glyphColor });
    label(s, [['Your Subtitle Here', r.ink]], { x: 9.194, y: r.y + 0.057, w: 3.379 });
    para(s, [[LOREM_CARD, r.ink]], { x: 9.194, y: r.y + 0.46, w: 3.379, h: 0.584 });
  });
}

/** 9 — Segmentation funnel built from rotated freeform ribbons. */
function slide09(pres) {
  const s = pres.addSlide();
  chrome(s, 9);

  // Seven ribbons: three dark connectors behind four bright bands. Each is
  // authored flat (x/y/w/h are the un-rotated box) then turned 90° so the funnel
  // tapers downward. Point lists are in the original artwork's own units.
  const ribbons = [
    { x: 5.636, y: 0.757, w: 4.368, h: 2.109, fill: GREEN_DARK, uw: 928221, uh: 448161,
      pts: [[92478, 276282], [0, 448161], [928221, 171878], [835743, 0]] },
    { x: 6.336, y: 2.057, w: 2.969, h: 2.109, fill: GREEN_DARK, uw: 630917, uh: 448161,
      pts: [[92478, 276282], [0, 448161], [630917, 171869], [538439, 0]] },
    { x: 7.035, y: 3.357, w: 1.570, h: 2.109, fill: GREEN_DARK, uw: 333613, uh: 448151,
      pts: [[92478, 276282], [0, 448151], [333613, 171869], [241135, 0]] },
    { x: 4.937, y: 0.757, w: 5.068, h: 0.809, fill: GREEN, uw: 1076877, uh: 171879,
      pts: [[1076877, 171879], [0, 171879], [92478, 0], [984399, 0]] },
    { x: 5.636, y: 2.057, w: 3.669, h: 0.809, fill: GREEN, uw: 779564, uh: 171878,
      pts: [[0, 171878], [92478, 0], [687086, 0], [779564, 171878]] },
    { x: 6.336, y: 3.357, w: 2.270, h: 0.809, fill: GREEN, uw: 482270, uh: 171869,
      pts: [[0, 171869], [92478, 0], [389792, 0], [482270, 171869]] },
    { x: 7.035, y: 4.658, w: 0.870, h: 0.809, fill: GREEN, uw: 184956, uh: 171869,
      pts: [[0, 171869], [92478, 0], [184956, 171869]] },
  ];
  ribbons.forEach(r => {
    const points = r.pts.map((p, i) => ({ x: (p[0] / r.uw) * r.w, y: (p[1] / r.uh) * r.h, moveTo: i === 0 }));
    points.push({ close: true });
    s.addShape('custGeom', {
      x: r.x, y: r.y, w: r.w, h: r.h, rotate: 90, points,
      fill: { color: r.fill }, line: { type: 'none' },
    });
  });

  // Quarter badges pinned to each funnel step.
  const badges = [
    { x: 4.589, y: 1.073, tag: 'Q1', tx: 4.341, ty: 1.230 },
    { x: 8.937, y: 2.218, tag: 'Q2', tx: 8.689, ty: 2.375 },
    { x: 5.968, y: 3.588, tag: 'Q3', tx: 5.720, ty: 3.745 },
    { x: 7.061, y: 5.181, tag: null },
  ];
  badges.forEach(b => {
    rect(s, { x: b.x, y: b.y, w: 0.819, h: 0.819, fill: GREEN, line: { color: WHITE, width: 3 } });
    if (b.tag) text(s, [[b.tag, WHITE]], { x: b.tx, y: b.ty, w: 1.315, h: 0.505, font: HEAD, size: 24, align: 'center' });
  });
  logo(s, 7.274, 5.407, 0.394, WHITE);

  const callouts = [
    { x: 1.847, y: 1.275, align: 'right' },
    { x: 3.216, y: 3.750, align: 'right' },
    { x: 10.004, y: 2.458, align: 'left' },
    { x: 8.151, y: 5.392, align: 'left' },
  ];
  callouts.forEach(c => {
    label(s, [['Subtitle Here', DARK]], { x: c.x, y: c.y, w: 2.568, align: c.align });
    para(s, LOREM_FUNNEL, { x: c.x, y: c.y + 0.403, w: 2.568, h: 0.584, align: c.align });
  });

  heading(s, [['Market ', DARK], ['Segmentation ', GREEN], ['Funnel', DARK]], { x: 0.76, y: 5.432, w: 5.906, h: 1.313 });
}

/** 10 — Product/market bridge: phone mock-up on the left, numbered tiles right. */
function slide10(pres) {
  const s = pres.addSlide();
  chrome(s, 10);

  // Phone mock-up: body, side buttons, bezel, earpiece slot and camera dot.
  // Only the top half is on-slide; the frame runs off the bottom edge.
  roundRect(s, { x: 1.578, y: 1.362, w: 3.632, h: 7.87, r: 0.45, fill: BLACK, shadow: PHONE_SHADOW });
  [
    { x: 5.174, y: 2.223, h: 0.335 }, { x: 5.174, y: 2.819, h: 0.605 },
    { x: 5.174, y: 3.595, h: 0.605 }, { x: 1.344, y: 3.011, h: 1.014 },
  ].forEach(b => roundRect(s, { x: b.x, y: b.y, w: 0.305, h: b.h, r: 0.06, fill: GREY_35 }));
  roundRect(s, { x: 1.367, y: 1.09, w: 4.09, h: 8.368, r: 0.5, fill: BLACK });
  roundRect(s, { x: 1.407, y: 1.132, w: 4.008, h: 8.282, r: 0.48, line: { color: GREY_35, width: 1.5 } });
  roundRect(s, { x: 3.169, y: 1.395, w: 0.501, h: 0.077, r: 0.038, fill: WHITE, transparency: 90 });
  s.addShape('ellipse', { x: 2.913, y: 1.384, w: 0.101, h: 0.098, fill: { color: WHITE, transparency: 90 }, line: { type: 'none' } });
  [[1.383, 1.934], [5.384, 1.934], [1.378, 8.622], [5.380, 8.622]].forEach(([x, y]) =>
    s.addShape('line', { x, y, w: 0.06, h: 0, line: { color: BLACK, transparency: 60, width: 2.5 } }));

  heading(s, [['Bridging ', DARK], ['Product', GREEN, { breakLine: true }], ['and Market', DARK]],
    { x: 6.745, y: 1.619, w: 5.906, h: 0.374 });

  const tiles = [
    { x: 6.745, y: 3.279, num: '01', green: false },
    { x: 9.571, y: 3.279, num: '02', green: false },
    { x: 6.745, y: 4.7, num: '03', green: true },
  ];
  tiles.forEach(t => {
    const ink = t.green ? WHITE : DARK;
    rect(s, { x: t.x, y: t.y, w: 2.559, h: 1.181, fill: t.green ? GREEN : WHITE, shadow: t.green ? undefined : CARD_SHADOW });
    text(s, [[t.num, t.green ? WHITE : GREEN]], { x: t.x + 0.196, y: t.y, w: 2.362, h: 0.693, font: HEAD, size: 32, ls: 1.2 });
    text(s, [['Subtitle Here', ink]], { x: t.x + 0.196, y: t.y + 0.693, w: 2.362, h: 0.324, bold: true, ls: 1.2 });
  });
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent erat quam, feugiat', { x: 9.571, y: 4.872, w: 2.878, h: 0.837 });

  rect(s, { x: 0, y: 4.7, w: 4.149, h: 1.181, fill: GREEN });
  para(s, [['Go To Market', WHITE]], { x: 0.76, y: 5.049, w: 2.272, h: 0.484, size: 20, font: HEAD });
  logo(s, 3.1, 5.023, 0.576, WHITE);
}

/** 11 — Full-bleed green support plan with four white stat cards. */
function slide11(pres) {
  const s = pres.addSlide();
  rect(s, { x: 0, y: 0, w: 13.333, h: 7.5, fill: GREEN });
  logo(s, -1.198, 3.248, 6.012, WHITE, 70);

  const cards = [
    { x: 5.77, y: 1.086, tx: 6.088, value: '400+', statY: 1.396, labelY: 2.136, bodyY: 2.468 },
    { x: 9.396, y: 1.086, tx: 9.768, value: '9/10', statY: 1.396, labelY: 2.136, bodyY: 2.468 },
    { x: 5.77, y: 4.052, tx: 6.088, value: '824K', statY: 4.448, labelY: 5.188, bodyY: 5.52 },
    { x: 9.396, y: 4.052, tx: 9.768, value: '#1', statY: 4.448, labelY: 5.188, bodyY: 5.52 },
  ];
  cards.forEach(c => {
    rect(s, { x: c.x, y: c.y, w: 3.177, h: 2.362, fill: WHITE });
    para(s, [[c.value, GREEN]], { x: c.tx, y: c.statY, w: 2.756, h: 0.74, size: 32, font: HEAD });
    label(s, [['Subtitle Here', DARK]], { x: c.tx, y: c.labelY, w: 2.756 });
    para(s, LOREM_CARD_SHORT, { x: c.tx, y: c.bodyY, w: 2.756, h: 0.584 });
  });

  heading(s, [['Post-Launch Support Plan', WHITE]], { x: 0.76, y: 1.297, w: 4.561, h: 1.313 });
  para(s, [['\u201C', WHITE]], { x: 0.76, y: 4.905, w: 0.948, h: 1.51, size: 72, font: HEAD });
  para(s, [
    ['Lorem ipsum dolor sit amet', WHITE, { bold: true }],
    [', consectetur adipiscing elit. Donec tempor metus urna, vel convallis mauris imperdiet sit amet. Cras imperdiet lacinia libero, nec blandit', WHITE],
  ], { x: 1.319, y: 5.248, w: 3.812, h: 1.089 });

  chrome(s, 11, WHITE);
}

/** 12 — Framework: numbered list on the left, green KPI rail on the right. */
function slide12(pres) {
  const s = pres.addSlide();
  chrome(s, 12);
  rect(s, { x: 9.861, y: 0.757, w: 2.712, h: 5.986, fill: GREEN });

  [
    { value: '72%', y: 1.429 },
    { value: '12K', y: 3.2 },
    { value: '84+', y: 4.972 },
  ].forEach(k => {
    icon(s, { x: 10.308, y: k.y + 0.065, size: 0.394, color: WHITE });
    stat(s, [[k.value, WHITE]], { x: 10.736, y: k.y, w: 1.698, h: 0.767, size: 44 });
    label(s, [['Your Subtitle Here', WHITE]], { x: 10.308, y: k.y + 0.767, w: 2.126 });
  });

  heading(s, [['The Go-To-Market ', DARK], ['Framework', GREEN]], { x: 0.76, y: 1.297, w: 5.111, h: 1.313 });
  [
    { num: '01', color: GREEN, y: 3.033 },
    { num: '02', color: DARK, y: 4.231 },
    { num: '03', color: GREEN, y: 5.429 },
  ].forEach(r => {
    text(s, [[r.num, r.color]], { x: 0.76, y: r.y, w: 1.271, h: 0.774, font: HEAD, size: 40, bold: true, ls: 1.0 });
    para(s, LOREM_CARD, { x: 1.76, y: r.y + 0.095, w: 3.421, h: 0.584 });
  });
}

/** 13 — Brief market introduction: photo band above a white text card. */
function slide13(pres) {
  const s = pres.addSlide();
  chrome(s, 13);
  gradientBand(s, { x: 0, y: 0, w: 13.333, h: 4.724, stops: [[0, 60], [0.6, 0]] });

  rect(s, { x: 3.911, y: 3.593, w: 8.661, h: 3.15, fill: WHITE, shadow: CARD_SHADOW_UP });
  rect(s, { x: 12.278, y: 3.593, w: 0.295, h: 3.15, fill: GREEN });
  heading(s, [['Brief Market ', DARK], ['Introduction', GREEN]], { x: 4.307, y: 4.058, w: 7.485, h: 0.707 });
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Donec eu porta augue, at eleifend est. Vivamus ornare lorem vitae ante porta tincidunt. Morbi facilisis lacinia nulla, sed sollicitudin nisl vehicula vel. Aliquam et neque at odio tempor vestibulum. Quisque in risus in ante egestas fringilla. Aliquam euismod velit a nibh vehicula, at dapibus turpis dignissim. Suspendisse euismod mauris id libero gravida',
    { x: 4.314, y: 4.937, w: 7.293, h: 1.342 });

  rect(s, { x: 1.365, y: 4.134, w: 1.181, h: 1.181, fill: GREEN });
  logo(s, 1.668, 4.457, 0.576, WHITE);
  para(s, [['2030 ', GREEN], ['Market value', DARK]], { x: 1.063, y: 5.418, w: 1.786, h: 1.325, size: 20, font: HEAD, align: 'center' });

  chrome(s, 13, WHITE, ['brand', 'template']);
}

/** 14 — Five-year Gantt-style timeline of pill bars. */
function slide14(pres) {
  const s = pres.addSlide();
  chrome(s, 14);
  heading(s, [['Yearly ', DARK], ['Timeline', GREEN]], { x: 2.583, y: 0.757, w: 8.167, h: 0.707, align: 'center' });
  para(s, LOREM_LONG, { x: 2.583, y: 1.521, w: 8.167, h: 0.584, align: 'center' });

  rect(s, { x: 0.76, y: 3.397, w: 11.812, h: 3.346, fill: GREY_05, transparency: 35 });
  rect(s, { x: 0.76, y: 2.82, w: 11.812, h: 0.48, fill: GREEN });
  ['2025', '2026', '2027', '2028', '2029'].forEach((year, i) => {
    text(s, [[year, WHITE]], { x: [0.877, 2.957, 5.354, 7.752, 10.149][i], y: 2.892, w: 2.424, h: 0.337, font: HEAD, size: 14, align: 'center', valign: 'middle' });
  });

  const bars = [
    { x: 0.877, y: 3.602, w: 3.543, green: false, tx: 1.271, tw: 2.756 },
    { x: 4.026, y: 4.243, w: 5.281, green: true, tx: 4.954, tw: 3.426 },
    { x: 8.502, y: 6.166, w: 3.543, green: false, tx: 8.896, tw: 2.756 },
    { x: 2.470, y: 5.525, w: 5.281, green: true, tx: 3.398, tw: 3.426 },
    { x: 5.297, y: 4.884, w: 5.281, green: false, tx: 6.224, tw: 3.426 },
  ];
  bars.forEach(b => {
    roundRect(s, { x: b.x, y: b.y, w: b.w, h: 0.295, fill: b.green ? GREEN : WHITE, shadow: CARD_SHADOW });
    text(s, [['Range of Deadline Here', b.green ? WHITE : GREEN]], { x: b.tx, y: b.y + 0.013, w: b.tw, h: 0.269, size: 10, bold: true, align: 'center' });
  });
}

/** 15 — Gallery slide: copy on the left, single green data tile bottom-right. */
function slide15(pres) {
  const s = pres.addSlide();
  chrome(s, 15);
  heading(s, [['Expansion Market ', DARK], ['Gallery', GREEN]], { x: 0.76, y: 1.277, w: 7.485, h: 0.707 });
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Donec eu porta augue, at eleifend est. Vivamus ornare lorem vitae ante porta tincidunt. Morbi facilisis lacinia nulla, sed sollicitudin nisl vehicula vel. Aliquam et neque',
    { x: 0.767, y: 2.156, w: 7.293, h: 0.837 });

  rect(s, { x: 8.911, y: 3.987, w: 3.661, h: 2.756, fill: GREEN });
  text(s, [['2600+', WHITE]], { x: 9.331, y: 4.447, w: 2.823, h: 1.01, font: HEAD, size: 54 });
  text(s, [['Your Data Here', WHITE]], { x: 9.331, y: 5.509, w: 2.823, h: 0.324, ls: 1.2 });
  text(s, [['Subtitle Here', WHITE]], { x: 9.331, y: 5.913, w: 2.823, h: 0.37, size: 16, bold: true });
}

/** 16 — KPI bars of increasing weight plus a "Key Data" list. */
function slide16(pres) {
  const s = pres.addSlide();
  chrome(s, 16);
  heading(s, [['KPIs & Performance ', DARK], ['Metrics', GREEN]], { x: 0.76, y: 0.772, w: 5.531, h: 1.313 });

  icon(s, { x: 7.843, y: 1.057, size: 0.394, color: GREEN });
  stat(s, [['$52M', GREEN]], { x: 8.236, y: 0.942, w: 3.55, h: 1.192, size: 72 });
  para(s, [['Capital Invest', DARK]], { x: 11.112, y: 1.323, w: 1.042, h: 0.584, bold: true });

  stat(s, [['Key Data:', DARK]], { x: 9.75, y: 3.347, w: 2.823, h: 0.525, size: 28 });
  bulletList(s, ['First Data', 'Second Data', 'Third Data'], { x: 9.75, y: 4.082, w: 2.404, h: 2.12, size: 16 });

  const bars = [
    { x: 0.76, w: 1.575, fill: GREY_05, value: '20%', ink: DARK, tx: 0.896, ix: 1.007 },
    { x: 2.335, w: 2.362, fill: GREY_15, value: '25%', ink: DARK, tx: 2.471, ix: 2.582 },
    { x: 4.697, w: 3.937, fill: GREEN, value: '55%', ink: WHITE, tx: 4.828, ix: 4.939 },
  ];
  bars.forEach(b => {
    rect(s, { x: b.x, y: 2.806, w: b.w, h: 3.937, fill: b.fill });
    icon(s, { x: b.ix, y: 3.15, size: 0.394, color: b.ink });
    text(s, [[b.value, b.ink]], { x: b.tx, y: 6.036, w: 1.439, h: 0.64, font: HEAD, size: 32 });
  });
  logo(s, 4.939, 5.494, 0.317, WHITE);
  label(s, [['Market Share', WHITE]], { x: 5.283, y: 5.475, w: 1.96 });
}

/** 17 — Value proposition: three numbered header/body card pairs. */
function slide17(pres) {
  const s = pres.addSlide();
  chrome(s, 17);
  heading(s, [['Product Value ', DARK], ['Proposition', GREEN]], { x: 0.76, y: 1.282, w: 4.485, h: 1.313 });
  rect(s, { x: 5.595, y: 0.757, w: 1.181, h: 2.362, fill: GREEN });
  logo(s, 5.898, 1.67, 0.576, WHITE);

  [
    { x: 0.76, headX: 0.761, num: '01', numX: 0.992, labelX: 1.815, bodyX: 1.089 },
    { x: 4.843, headX: 4.843, num: '02', numX: 5.074, labelX: 5.898, bodyX: 5.171 },
    { x: 9.322, headX: 9.322, num: '03', numX: 9.553, labelX: 10.377, bodyX: 9.650 },
  ].forEach(c => {
    rect(s, { x: c.x, y: 4.879, w: 3.251, h: 1.864, fill: WHITE, shadow: CARD_SHADOW });
    rect(s, { x: c.headX, y: 3.846, w: 3.251, h: 0.787, fill: GREEN });
    text(s, [[c.num, WHITE]], { x: c.numX, y: 3.987, w: 0.824, h: 0.505, font: HEAD, size: 24, bold: true });
    text(s, [['Subtitle Here', WHITE]], { x: c.labelX, y: 4.073, w: 1.466, h: 0.333, bold: true, ls: 1.2 });
    para(s, LOREM_VALUE, { x: c.bodyX, y: 5.14, w: 2.594, h: 1.342 });
  });
}

/** 18 — Risk mitigation: native clustered column chart with alternating bars. */
function slide18(pres) {
  const s = pres.addSlide();
  chrome(s, 18);

  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep'];
  const values = [4.3, 2.5, 3.5, 4.5, 3, 2.5, 4, 3, 2];
  const barColors = [GREEN, GREY_05, GREY_05, GREEN, GREY_05, GREY_05, GREEN, GREY_05, GREEN];
  s.addChart('bar', [{ name: 'Series 1', labels: months, values }], {
    x: 0.76, y: 1.771, w: 7.081, h: 4.972,
    barDir: 'col', barGapWidthPct: 100, barOverlapPct: -27,
    chartColors: barColors,
    showLegend: false, showTitle: false,
    catAxisLineShow: true, catAxisLineColor: GREY_15,
    catAxisLabelColor: '595959', catAxisLabelFontFace: BODY, catAxisLabelFontSize: 12,
    valAxisLineShow: false,
    valAxisLabelColor: '595959', valAxisLabelFontFace: BODY, valAxisLabelFontSize: 12,
    valAxisMinVal: 0, valAxisMaxVal: 5, valAxisMajorUnit: 0.5,
    catGridLine: { style: 'none' }, valGridLine: { style: 'none' },
  });

  roundRect(s, { x: 7.007, y: 3.859, w: 1.098, h: 0.438, r: 0.07, fill: WHITE, shadow: CARD_SHADOW });
  text(s, [['40%', GREEN]], { x: 7.105, y: 3.909, w: 0.999, h: 0.337, font: HEAD, size: 14, valign: 'middle' });
  icon(s, { x: 7.769, y: 3.979, size: 0.197, color: GREEN });

  rect(s, { x: 1.151, y: 0.757, w: 6.299, h: 0.787, fill: GREEN });
  text(s, [['2030', WHITE]], { x: 1.458, y: 0.898, w: 1.466, h: 0.505, font: HEAD, size: 24, bold: true });
  text(s, [['Market Growth Target & Stakeholder', WHITE]], { x: 2.771, y: 0.984, w: 3.896, h: 0.324, bold: true, ls: 1.2 });
  logo(s, 6.741, 1.003, 0.317, WHITE);

  heading(s, [['Risk Mitigation ', DARK], ['Alignment', GREEN]], { x: 8.508, y: 1.544, w: 4.065, h: 1.192, size: 36 });
  label(s, [['Put Your Subtitle Here', DARK]], { x: 8.508, y: 4.438, w: 4.065 });
  stat(s, [['$435K', GREEN]], { x: 8.508, y: 4.833, w: 3.456, h: 1.01, size: 54 });
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Donec tempor metus urna convallis mauris imperdiet sit amet',
    { x: 8.508, y: 5.906, w: 4.065, h: 0.837 });
}

/** 19 — Countdown: three time tiles above a centred headline. */
function slide19(pres) {
  const s = pres.addSlide();
  chrome(s, 19);

  [
    { x: 4.392, tx: 4.591, value: '24', unit: 'Sec', green: false },
    { x: 5.978, tx: 6.177, value: '35', unit: 'Min', green: true },
    { x: 7.563, tx: 7.762, value: '12', unit: 'Hour', green: false },
  ].forEach(t => {
    const ink = t.green ? WHITE : DARK;
    rect(s, { x: t.x, y: 3.316, w: 1.378, h: 1.181, fill: t.green ? GREEN : WHITE, shadow: t.green ? undefined : CARD_SHADOW });
    text(s, [[t.value, ink]], { x: t.tx, y: 3.316, w: 0.98, h: 0.707, font: HEAD, size: 32, align: 'center', ls: 1.2 });
    text(s, [[t.unit, ink]], { x: t.tx, y: 4.009, w: 0.98, h: 0.328, bold: true, align: 'center', ls: 1.2 });
  });

  heading(s, [['Product Launch ', DARK], ['Countdown', GREEN]], { x: 2.583, y: 5.117, w: 8.167, h: 0.707, align: 'center' });
  para(s, LOREM_LONG, { x: 2.583, y: 5.881, w: 8.167, h: 0.584, align: 'center' });
}

/** 20 — Closing slide with contact line and oversized thank-you. */
function slide20(pres) {
  const s = pres.addSlide();
  chrome(s, 20);
  rect(s, { x: 6.766, y: 0.757, w: 5.807, h: 0.295, fill: GREEN });

  [
    { x: 6.007, w: 2.275, body: 'YOUR LOCATION' },
    { x: 7.894, w: 2.039, body: '+123 456 7890' },
    { x: 9.696, w: 2.877, body: 'WWW.YOURWEBSITE.COM' },
  ].forEach(c => text(s, [[c.body, DARK]], { x: c.x, y: 3.272, w: c.w, h: 0.314, font: HEAD, align: 'right', ls: 1.1, after: 8 }));

  heading(s, [['Thank You Everyone!', GREEN]], { x: 6.135, y: 3.824, w: 6.438, h: 2.1, size: 66, align: 'right' });
  para(s, LOREM_INTRO, { x: 6.135, y: 6.159, w: 6.438, h: 0.584, align: 'right' });

  rect(s, { x: 2.756, y: 4.48, w: 0.787, h: 0.787, fill: GREEN });
  logo(s, 2.991, 4.726, 0.317, WHITE);

  chrome(s, 20, WHITE, ['brand', 'footer']);
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */

function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'WIDE_13x7_5', width: 13.333, height: 7.5 });
  pres.layout = 'WIDE_13x7_5';
  pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };
  pres.author = 'Go To Market';
  pres.title = 'Go To Market';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
   slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach(fn => fn(pres));

  return pres.writeFile({ fileName: path.join(__dirname, '0944d93d-177d-4e00-9d53-21304016e12e_grok_final.pptx') });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
