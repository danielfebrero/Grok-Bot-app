/**
 * DigEdu — "Shaping Future Generations" pitch deck, rebuilt with pptxgenjs.
 *
 * Slide size: 26.667 x 15 in (widescreen, 24384000 x 13716000 EMU).
 * Raster photos in the original are replaced with flat placeholder rectangles.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- theme ----
const C = {
  navy: '2A3782',        // accent1 - deep blue band / globe sphere
  blue: '284298',        // accent2 - card blue
  amber: 'FFAB00',       // accent3 - highlight orange
  ink: '16195F',         // accent4 - headline text
  teal: '04D1BC',        // accent5 - glow + map pins
  white: 'FFFFFF',
  black: '000000',
  violet: '7867BC',      // wave start cap on the timeline
  photo: 'D9DDE8',       // stand-in tone for replaced photographs
  photoTxt: '8A93AD',
};

const FONT = 'Poppins';
const FONT_MED = 'Poppins Medium';

const W = 26.667;   // slide width  (in)
const H = 15;       // slide height (in)

// Soft drop shadow used by every white/coloured card in the deck.
// A fresh object per shape: pptxgenjs rewrites shadow props in place.
const shadow = opacity => ({ type: 'outer', color: C.black, blur: 20, offset: 0, angle: 90, opacity });
const cardShadow = () => shadow(0.1);
const btnShadow = () => shadow(0.2);
const NO_LINE = { type: 'none' };

// -------------------------------------------------------------- helpers ----
const rect = (s, o) => s.addShape('rect', Object.assign({ line: NO_LINE }, o));
const oval = (s, o) => s.addShape('ellipse', Object.assign({ line: NO_LINE }, o));

/** Rounded rectangle. `r` is the PowerPoint "adj" fraction of the short side. */
function roundRect(s, o) {
  const r = o.r === undefined ? 0.1 : o.r;
  const opts = Object.assign({ line: NO_LINE }, o);
  delete opts.r;
  opts.rectRadius = r * Math.min(o.w, o.h);
  s.addShape('roundRect', opts);
}

/** Body copy: 1.5 line spacing, regular weight. */
function body(s, text, o) {
  s.addText(text, Object.assign({
    fontFace: FONT, fontSize: 24, color: C.black, lineSpacingMultiple: 1.5,
    valign: 'top', align: 'left',
  }, o));
}

/** Bold Poppins label ("Your tittle here", years, names, ...). */
function label(s, text, o) {
  s.addText(text, Object.assign({
    fontFace: FONT, fontSize: 24, bold: true, color: C.ink,
    valign: 'top', align: 'left',
  }, o));
}

/** Big statistic ("65%", "1043+", "$8M"). */
function stat(s, text, o) {
  s.addText(text, Object.assign({
    fontFace: FONT, fontSize: 88, bold: true, color: C.ink, valign: 'top',
  }, o));
}

/** Section eyebrow + headline pair, as used on almost every slide. */
function heading(s, o) {
  s.addText(o.eyebrow, {
    x: o.x, y: o.eyebrowY, w: o.eyebrowW || o.w, h: 0.707,
    fontFace: FONT_MED, fontSize: 36, color: C.ink, valign: 'top',
    align: o.align || 'left',
  });
  s.addText(o.title, {
    x: o.x, y: o.titleY, w: o.w, h: o.titleH || 1.313,
    fontFace: FONT, fontSize: 72, bold: true, color: C.ink, valign: 'top',
    align: o.align || 'left',
  });
}

/** Pill-shaped call to action. */
function pillButton(s, text, x, y, w, h, fontSize) {
  s.addText(text, {
    shape: 'roundRect', rectRadius: Math.min(w, h) / 2,
    x, y, w, h, fill: { color: C.amber }, line: NO_LINE, shadow: btnShadow(),
    fontFace: FONT_MED, fontSize: fontSize || 24, color: C.ink,
    align: 'center', valign: 'middle',
  });
}

/**
 * Radial teal "glow" blob behind the content on every slide. pptxgenjs has no
 * radial gradient, so it is stacked from concentric translucent discs whose
 * accumulated opacity follows the original centre-out falloff.
 * `peak` is the opacity at the centre, 0..1.
 */
function glow(s, x, y, d, peak) {
  const RINGS = 26;
  let painted = 0;
  for (let i = 0; i < RINGS; i++) {
    const f = (RINGS - i) / RINGS;                       // outer ring first
    const target = peak * (Math.exp(-3.1 * f * f) - Math.exp(-3.1)) / (1 - Math.exp(-3.1));
    const a = 1 - (1 - target) / (1 - painted);
    painted = target;
    if (a < 0.004) continue;
    oval(s, {
      x: x + d * (1 - f) / 2, y: y + d * (1 - f) / 2, w: d * f, h: d * f,
      fill: { color: C.teal, transparency: 100 - a * 100 },
    });
  }
}

/**
 * Free-form outline. `pts` is a list of commands whose coordinates are
 * fractions of the shape box (custGeom points are shape-relative):
 *   ['M', x, y] ['L', x, y] ['C', x1, y1, x2, y2, x, y] ['Z']
 */
function freeform(s, o) {
  const pts = [];
  o.pts.forEach(cmd => {
    const k = cmd[0];
    const fx = i => cmd[i] * o.w;
    const fy = i => cmd[i] * o.h;
    if (k === 'M') pts.push({ x: fx(1), y: fy(2), moveTo: true });
    else if (k === 'L') pts.push({ x: fx(1), y: fy(2) });
    else if (k === 'C') pts.push({ x: fx(5), y: fy(6), curve: { type: 'cubic', x1: fx(1), y1: fy(2), x2: fx(3), y2: fy(4) } });
    else pts.push({ close: true });
  });
  s.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h, points: pts,
    fill: o.fill, line: NO_LINE, shadow: o.shadow,
  });
}

/** Filled polygon from a flat [x0,y0,x1,y1,...] list in per-mille of w/h. */
function polygon(s, box, flat, color) {
  const pts = [];
  for (let i = 0; i < flat.length; i += 2) {
    pts.push({ x: (flat[i] / 1000) * box.w, y: (flat[i + 1] / 1000) * box.h });
  }
  pts[0].moveTo = true;
  pts.push({ close: true });
  s.addShape('custGeom', { x: box.x, y: box.y, w: box.w, h: box.h, points: pts, fill: { color }, line: NO_LINE });
}

/** DigEdu wordmark + mortarboard glyph, top-right on most slides. */
function logo(s, x) {
  const lx = x === undefined ? 22.168 : x;
  const y = 1.313, u = 1.032 / 12;   // glyph drawn on a 12-unit grid
  // Mortarboard: outer diamond with a diamond-shaped hole, then the "book" body.
  freeform(s, {
    x: lx, y: y, w: 1.032, h: 1.032, fill: { color: C.ink },
    pts: [
      ['M', 0.00, 0.34], ['L', 0.50, 0.09], ['L', 1.00, 0.34], ['L', 0.50, 0.59], ['Z'],
      ['M', 0.19, 0.34], ['L', 0.50, 0.19], ['L', 0.81, 0.34], ['L', 0.50, 0.49], ['Z'],
    ],
  });
  freeform(s, {
    x: lx + 0.6 * u, y: y + 5.6 * u, w: 8.6 * u, h: 4.0 * u, fill: { color: C.ink },
    pts: [
      ['M', 0.00, 0.00], ['L', 0.00, 0.62],
      ['C', 0.20, 0.90, 0.36, 0.90, 0.50, 0.72],
      ['C', 0.64, 0.90, 0.80, 0.90, 1.00, 0.62], ['L', 1.00, 0.00],
      ['C', 0.80, 0.34, 0.64, 0.34, 0.50, 0.16],
      ['C', 0.36, 0.34, 0.20, 0.34, 0.00, 0.00], ['Z'],
      ['M', 0.13, 0.26], ['L', 0.13, 0.62],
      ['C', 0.26, 0.74, 0.38, 0.72, 0.44, 0.60], ['L', 0.44, 0.32],
      ['C', 0.32, 0.44, 0.22, 0.40, 0.13, 0.26], ['Z'],
      ['M', 0.87, 0.26], ['L', 0.87, 0.62],
      ['C', 0.74, 0.74, 0.62, 0.72, 0.56, 0.60], ['L', 0.56, 0.32],
      ['C', 0.68, 0.44, 0.78, 0.40, 0.87, 0.26], ['Z'],
    ],
  });
  rect(s, { x: lx + 9.5 * u, y: y + 4.1 * u, w: 0.55 * u, h: 5.0 * u, fill: { color: C.ink } });
  roundRect(s, { x: lx + 8.85 * u, y: y + 8.7 * u, w: 1.9 * u, h: 2.6 * u, r: 0.3, fill: { color: C.ink } });
  oval(s, { x: lx + 9.45 * u, y: y + 9.5 * u, w: 0.7 * u, h: 0.7 * u, fill: { color: C.white } });
  s.addText([{ text: 'DigEdu', options: {} }, { text: '.', options: {} }], {
    x: lx + 0.933, y: y + 0.2, w: 1.869, h: 0.572,
    fontFace: FONT, fontSize: 28, bold: true, color: C.ink, valign: 'top', margin: 0,
  });
}

/**
 * Stand-in for a photograph in the source deck: a flat rounded rectangle
 * labelled "[image]".
 */
function imagePlaceholder(s, x, y, w, h, r) {
  roundRect(s, {
    x, y, w, h, r: r === undefined ? 0.12 : r, fill: { color: C.photo },
  });
  s.addText('[image]', {
    x, y, w, h, fontFace: FONT, fontSize: 20, color: C.photoTxt,
    align: 'center', valign: 'middle',
  });
}

/**
 * Line-art pictogram standing in for the deck's icon set: a thin-stroked
 * "presentation board" glyph on a stand with chart bars inside. Built from
 * plain rectangles so the stroke weight scales with `size`.
 * `bg` is the surface behind the icon; it fills the hollow of the frame.
 */
function icon(s, x, y, size, color, bg) {
  const c = color || C.ink;
  const hollow = bg || C.white;
  const t = size * 0.062;                       // stroke weight
  const by = y + size * 0.155, bh = size * 0.60;
  roundRect(s, { x, y: by, w: size, h: bh, r: 0.16, fill: { color: c } });
  roundRect(s, { x: x + t, y: by + t, w: size - 2 * t, h: bh - 2 * t, r: 0.13, fill: { color: hollow } });
  [[0.24, 0.34], [0.44, 0.52], [0.64, 0.42]].forEach(([fx, fh]) => {
    rect(s, { x: x + size * fx, y: by + bh - t * 1.9 - bh * fh, w: size * 0.085, h: bh * fh, fill: { color: c } });
  });
  rect(s, { x: x + size * 0.5 - t / 2, y: y + size * 0.755, w: t, h: size * 0.15, fill: { color: c } });
  rect(s, { x: x + size * 0.28, y: y + size * 0.90, w: size * 0.44, h: t, fill: { color: c } });
  rect(s, { x: x + size * 0.5 - t / 2, y: y + size * 0.02, w: t, h: size * 0.135, fill: { color: c } });
}

/** Location pin: a circular head tapering to a point (used on the world map). */
function mapPin(s, x, y, w, h) {
  freeform(s, {
    x, y, w, h, fill: { color: C.teal },
    pts: [
      ['M', 0.50, 1.00], ['L', 0.12, 0.59],
      ['C', 0.04, 0.47, 0.00, 0.32, 0.00, 0.25],
      ['C', 0.00, 0.11, 0.22, 0.00, 0.50, 0.00],
      ['C', 0.78, 0.00, 1.00, 0.11, 1.00, 0.25],
      ['C', 1.00, 0.32, 0.96, 0.47, 0.88, 0.59], ['Z'],
    ],
  });
}

/** Coloured disc with a small icon inside — used on every feature card. */
function iconBadge(s, cx, cy, d, ix, iy, isize, discColor) {
  const disc = discColor || C.amber;
  oval(s, { x: cx, y: cy, w: d, h: d, fill: { color: disc } });
  icon(s, ix, iy, isize, C.ink, disc);
}

// ----------------------------------------------- shared artwork fragments ---

// Outline of the big white "portal" panel that frames slides 1 and 20.
const PORTAL_OUTER = [
  ['M', 0.2051, 0], ['L', 0.8587, 0], ['L', 0.8749, 0.0171],
  ['C', 0.9254, 0.0728, 0.9679, 0.1411, 0.9997, 0.2184],
  ['L', 1, 0.2192], ['L', 1, 0.8873], ['L', 0.9887, 0.9136],
  ['C', 0.9773, 0.9387, 0.9648, 0.9628, 0.9512, 0.9856],
  ['L', 0.9421, 1], ['L', 0.1216, 1], ['L', 0.1185, 0.9953],
  ['C', 0.0445, 0.8752, 0, 0.7212, 0, 0.5533],
  ['C', 0, 0.3374, 0.0735, 0.1445, 0.1889, 0.0171], ['Z'],
];
const PORTAL_INNER = [
  ['M', 0.4879, 0], ['L', 0.5121, 0],
  ['C', 0.7815, 0, 1, 0.2435, 1, 0.5438],
  ['C', 1, 0.7315, 0.9147, 0.897, 0.7849, 0.9947],
  ['L', 0.7775, 1], ['L', 0.2225, 1], ['L', 0.2151, 0.9947],
  ['C', 0.0853, 0.897, 0, 0.7315, 0, 0.5438],
  ['C', 0, 0.2435, 0.2185, 0, 0.4879, 0], ['Z'],
];

/** Slides 1 & 20 share the same background: nested white panels + glow. */
function portalBackdrop(s) {
  freeform(s, { x: 6.585, y: 0, w: 20.082, h: 15, pts: PORTAL_OUTER, fill: { color: C.white }, shadow: cardShadow() });
  freeform(s, { x: 9.366, y: 0.821, w: 15.803, h: 14.179, pts: PORTAL_INNER, fill: { color: C.white }, shadow: cardShadow() });
  roundRect(s, { x: 12.352, y: 3.685, w: 9.83, h: 9.592, r: 0.5, fill: { color: C.white }, shadow: cardShadow() });
  glow(s, 9.306, 2.887, 10.613, 0.61);
}

const LOREM = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.';
const LOREM_LONG = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit,.';

// Continents of the flat world map on slide 6 (per-mille of the map box).
const WORLD_MAP = [
  [935,81,877,78,881,95,870,84,849,125,813,83,790,86,778,125,758,98,711,123,715,82,683,64,652,143,
   662,163,641,157,649,229,633,158,636,210,609,193,614,213,585,235,576,219,564,266,552,241,570,231,
   535,199,505,224,483,286,504,332,458,385,467,417,445,420,446,457,465,459,497,413,514,460,523,438,
   508,405,535,459,542,434,551,459,575,459,573,488,526,498,499,458,454,467,419,582,446,636,500,642,
   525,846,567,813,594,737,590,686,627,609,603,600,571,511,604,593,651,541,640,519,627,531,617,497,
   692,558,747,539,756,558,800,542,807,558,837,494,833,438,817,428,827,413,854,451,845,419,867,367,
   866,313,883,359,865,291,842,291,850,245,883,231,898,182,895,259,912,308,905,214,936,165,920,116,
   940,111],
  [35,190,37,228,48,233,16,270,68,220,64,234,80,219,105,257,117,353,109,430,141,534,130,482,154,556,
   234,616,225,695,260,767,259,947,266,982,292,993,280,976,284,893,349,778,365,691,357,683,336,734,
   312,633,260,597,223,613,219,580,203,576,210,545,186,564,177,532,183,503,208,492,230,523,230,485,
   257,430,301,396,278,375,313,358,304,384,321,392,301,286,290,301,280,263,268,262,249,353,246,322,
   218,284,252,227,251,249,265,246,255,218,269,213,271,182,285,182,293,208,273,241,304,267,300,233,
   318,231,294,142,270,161,274,133,255,207,251,150,267,129,257,123,227,194,213,175,203,193,187,165,
   173,199,140,190,107,245,70,190],
  [402,21,393,23,391,32,387,31,388,50,382,33,378,46,375,45,373,31,357,37,354,49,349,48,341,57,340,
   62,344,68,339,77,324,80,321,85,327,98,321,103,322,116,325,120,335,119,343,130,345,167,341,182,
   346,182,348,190,342,190,340,203,344,207,347,199,349,206,339,224,337,238,343,281,356,295,369,251,
   385,241,393,226,406,222,416,211,409,206,411,197,415,205,419,205,416,187,419,182,418,174,423,172,
   423,165,427,161,427,140,424,135,428,130,430,140,431,126,428,123,428,114,439,77,447,66,445,59,439,
   57,419,108,399,54],
  [935,806,933,805,929,790,925,787,924,776,917,768,913,741,910,737,904,737,904,749,900,760,897,759,
   896,754,893,755,886,745,858,792,840,768,824,777,817,784,814,796,815,806,813,810,817,846,813,859,
   816,864,820,866,829,859,838,859,846,850,864,846,872,853,875,869,882,859,878,872,883,870,882,876,
   884,876,886,893,893,902,899,898,902,905,907,899,915,894,918,880,930,852,934,829],
  [359,16,349,5,343,5,339,1,336,4,332,1,330,5,327,2,322,11,319,5,317,5,315,13,306,11,303,14,303,32,
   308,39,318,43,316,46,312,45,311,49,305,41,301,44,302,38,298,31,298,20,295,17,295,22,292,23,291,
   29,287,30,287,36,284,46,284,50,287,55,284,56,282,68,285,71,292,69,292,76,295,81,293,84,291,79,
   289,79,289,88,281,92,275,92,270,83,267,82,267,92,272,97,268,116,277,124,287,128,293,127,296,119,
   294,112,281,111,275,101,280,97,287,103,296,104,297,108,305,103,304,96,310,86,315,86,315,83,318,
   80,319,72,322,71,321,64,330,64,339,51,349,43,349,37,360,31,361,25],
  [194,90,192,98,187,106,187,128,188,128,192,121,199,114,200,117,197,119,193,127,193,131,195,132,
   196,135,192,136,191,141,193,144,200,145,204,150,204,152,202,153,196,149,190,149,189,152,190,160,
   195,164,194,172,195,173,203,174,213,169,215,176,223,178,224,168,227,171,229,168,229,165,225,156,
   224,149,226,144,226,138,228,131,234,121,234,117,229,115,228,121,225,120,223,123,220,140,218,140,
   219,123,215,128,214,127,214,121,210,118,208,98,204,96,200,90],
  [621,725,620,725,619,730,614,741,606,747,604,753,605,766,601,774,601,783,601,791,605,796,611,795,
   612,793,620,757,621,748,623,744,623,731],
  [522,84,517,81,514,81,513,79,512,83,510,83,507,79,506,82,505,82,504,89,500,87,499,92,494,90,491,
   93,490,98,492,102,491,109,493,111,494,110,497,114,498,126,503,131,504,136,505,135,506,126,507,
   123,508,110,509,107,511,108,512,113,513,123,516,123,517,124,519,118,517,115,516,112,514,110,514,
   107,513,106,510,100,507,98,505,91,507,90,508,93,511,96,517,98,520,94],
  [613,108,609,110,608,115,606,119,600,121,596,129,595,135,591,140,592,144,592,151,590,156,591,161,
   590,168,591,176,588,182,590,186,592,186,593,188,599,194,605,193,605,190,599,182,597,173,597,165,
   598,159,600,145,602,142,607,129,616,116,615,111],
  [883,401,882,401,881,405,881,413,882,417,881,427,879,428,880,429,880,433,878,436,876,436,875,434,
   874,444,871,447,866,448,861,456,860,460,858,465,860,469,859,472,862,479,864,476,864,464,863,461,
   863,459,866,457,866,464,868,465,869,461,871,461,872,459,876,460,879,452,884,451,885,448,887,449,
   889,446,888,432,887,429,887,426,888,423,888,414],
  [461,315,457,315,454,322,454,329,455,331,454,334,457,337,457,342,460,342,462,347,461,353,458,353,
   459,359,457,362,457,363,459,365,460,367,456,375,459,375,463,371,467,371,474,368,473,363,475,358,
   473,354,471,354,469,346,467,342,465,336,463,333,464,323,464,321,461,321,460,320],
  [214,83,214,85,217,90,221,90,220,93,217,94,217,97,219,99,222,100,230,96,236,98,238,97,243,89,242,
   85,238,85,238,80,239,78,239,71,237,72,235,75,235,81,233,83,234,86,233,88,229,85,228,74,225,68,
   220,72,217,78,217,81],
  [981,937,979,936,978,933,977,933,968,951,959,960,953,968,951,975,956,982,960,981,963,976,969,961,
   974,959,974,956,980,943],
  [410,250,410,253,414,253,414,257,411,257,411,259,414,262,414,265,413,267,416,268,420,271,424,271,
   433,264,435,260,436,254,434,249,430,245,427,249,420,249,419,253,417,252,417,249,415,245,412,247],
  [644,20,643,30,644,39,647,40,647,42,645,43,646,47,648,48,650,44,651,44,651,49,654,50,656,53,665,
   55,667,51,665,47,664,43,662,37,659,35,654,35,650,25,645,19],
  [764,41,763,43,762,49,765,58,771,64,774,64,778,56,780,56,783,49,782,43,776,39,775,41,773,41,771,
   39,770,40,769,46],
  [251,119,244,120,241,124,242,131,241,134,239,132,238,129,236,130,235,138,237,142,239,154,246,148,
   249,141,249,135,247,131,252,125,251,123],
  [987,881,987,891,989,907,987,916,983,920,983,924,986,927,986,932,983,937,984,940,987,939,991,930,
   993,922,996,921,999,913,999,909,997,911,994,909,993,908,992,901,990,901,990,889],
  [249,78,247,85,247,92,249,95,251,95,252,97,250,101,250,104,254,106,255,106,257,101,258,101,262,86,
   256,80,255,88,253,85,251,88,250,85,250,79],
  [229,51,226,51,223,50,215,57,208,61,207,64,207,66,213,70,214,68,215,68,217,63,219,62,220,64,219,
   68,220,68,229,56],
  [879,369,879,371,880,375,880,385,879,387,878,393,880,400,882,397,880,394,881,392,884,391,888,394,
   888,390,892,385,893,382,891,377,887,378],
  [454,339,448,339,446,345,444,345,444,350,445,354,442,361,443,363,446,364,453,359,454,356,454,347,
   455,345],
  [263,38,261,43,262,46,262,50,259,50,258,53,264,59,264,67,266,67,270,55,270,45,266,43,266,38],
  [905,922,899,922,895,918,894,935,896,942,901,938,905,925],
  [670,42,669,44,670,50,672,64,673,66,679,57,680,51,677,46],
  [216,543,224,539,226,541,234,544,236,548,238,550,238,554,247,553,246,550,231,537,225,535,220,536],
  [247,562,248,564,253,564,255,566,258,563,265,562,265,560,261,556,254,555,252,553,251,554,252,560],
  [566,64,564,60,563,60,562,65,559,63,556,63,554,68,555,70,559,70,559,75,562,76,563,72,566,74,566,
   72,564,72,564,70]
];

// Land masses and the highlighted region of the 3-D globe on slide 3.
const GLOBE_LAND = [
  [259,63,159,134,97,241,171,175,159,195,187,174,226,200,238,303,197,394,264,552,242,475,307,594,
   495,674,471,781,556,857,542,988,588,984,742,839,789,753,687,734,698,707,611,654,467,671,455,623,
   411,615,425,574,366,592,349,546,418,490,463,495,482,538,481,481,537,413,646,352,590,325,692,289,
   597,179,534,290,463,215,474,171,577,126,565,101,538,120,531,76,509,111,425,114,359,71,308,92],
  [590,11,655,50,668,62,666,72,670,76,689,79,721,115,717,131,726,142,728,152,719,170,720,187,759,
   235,765,236,781,189,822,189,833,184,834,174,822,156,829,137,827,126,780,87,693,39]
];
const GLOBE_HIGHLIGHT = [
  [930,453,927,458,936,478,936,495,900,548,897,559,900,583,887,607,885,621,892,630,919,648,952,654,
   960,664,962,683,980,633,990,590,997,542,999,494,966,463,942,453]
];

// ------------------------------------------------------------- slides ------

// 1 — Title
function slide01(pres) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  portalBackdrop(s);
  s.addText('Digital Education Reimagined', {
    x: 2.63, y: 5.722, w: 12.003, h: 1.01,
    fontFace: FONT_MED, fontSize: 54, color: C.ink, valign: 'top',
  });
  s.addText([
    { text: 'Shaping Future', options: { breakLine: true } },
    { text: 'Generations.', options: {} },
  ], {
    x: 2.63, y: 6.898, w: 13.596, h: 3.972,
    fontFace: FONT, fontSize: 115, bold: true, color: C.ink, valign: 'top',
  });
  icon(s, 2.699, 3.691, 1.609);
  pillButton(s, 'Learn More', 2.699, 11.309, 3.439, 1.084);
  logo(s);
}

// 2 — The Challenge (three issue cards)
function slide02(pres) {
  const s = pres.addSlide();
  imagePlaceholder(s, 16, 0, 10.667, 15, 0);
  glow(s, 4.831, 1.125, 9.849, 0.42);
  rect(s, { x: 16, y: 0, w: 10.667, h: 15, fill: { color: C.navy, transparency: 32 } });

  heading(s, { x: 1.867, w: 8.8, eyebrowY: 7.23, titleY: 5.917, eyebrow: 'Educational Access and Quality', title: 'The Challenge' });
  icon(s, 1.669, 4.203, 1.477);
  label(s, 'Your tittle here', { x: 1.923, y: 8.833, w: 3.434, h: 0.572, fontSize: 28 });
  body(s, LOREM_LONG, { x: 1.893, y: 9.459, w: 7.913, h: 1.262 });
  pillButton(s, 'Learn More', 1.997, 11.085, 3.439, 1.084);
  logo(s, 1.929);

  const cards = [
    { y: 2.906, x: 12.000, dark: true, tx: 14.366, tw: 4.303, cx: 12.667, cy: 3.556, ix: 12.953, iy: 3.818, id: 0.615, t: 'Uneven quality of education globally' },
    { y: 6.365, x: 11.986, dark: false, tx: 14.352, tw: 4.773, cx: 12.653, cy: 7.015, ix: 12.953, iy: 7.284, id: 0.615, t: 'High cost of educational materials' },
    { y: 9.824, x: 11.963, dark: true, tx: 14.329, tw: 4.773, cx: 12.630, cy: 10.473, ix: 13.000, iy: 10.876, id: 0.517, t: 'Limited access to expert knowledge' },
  ];
  cards.forEach(c => {
    roundRect(s, {
      x: c.x, y: c.y, w: 7.335, h: 2.594, r: 0.20513,
      fill: { color: c.dark ? C.blue : C.white }, shadow: c.dark ? undefined : cardShadow(),
    });
    body(s, c.t, { x: c.tx, y: c.y + 0.573, w: c.tw, h: 1.456, fontSize: 28, color: c.dark ? C.white : C.ink });
    iconBadge(s, c.cx, c.cy, 1.222, c.ix, c.iy, c.id);
  });
}

// 3 — Uneven quality (globe + stat card)
function slide03(pres) {
  const s = pres.addSlide();
  glow(s, 9.767, -0.064, 8.163, 0.42);
  rect(s, { x: 0, y: 10.167, w: W, h: 4.833, fill: { color: C.navy } });

  const gb = { x: 2.025, y: 3.403, w: 9.528, h: 9.528 };
  oval(s, Object.assign({ fill: { color: '6675CD' } }, gb));
  GLOBE_LAND.forEach(p => polygon(s, gb, p, 'FCF4E0'));
  GLOBE_HIGHLIGHT.forEach(p => polygon(s, gb, p, C.amber));

  roundRect(s, { x: 12.712, y: 2.842, w: 12.703, h: 10.658, r: 0.07764, fill: { color: C.white }, shadow: cardShadow() });
  icon(s, 14.104, 4.086, 0.942);
  heading(s, {
    x: 14.049, w: 10.8, titleH: 2.524, eyebrowY: 5.141, titleY: 5.848,
    eyebrow: 'The Challenge', title: 'Uneven quality of education globally',
  });
  roundRect(s, { x: 14.15, y: 8.833, w: 9.83, h: 3.333, r: 0.16667, fill: { color: C.white }, shadow: cardShadow() });
  stat(s, '65%', { x: 14.955, y: 9.726, w: 2.927, h: 1.582 });
  label(s, 'Your tittle here', { x: 18.122, y: 9.662, w: 3.434, h: 0.505 });
  body(s, LOREM_SHORT, { x: 18.118, y: 10.167, w: 5.29, h: 1.262 });
  logo(s);
}

// 4 — Our Solution (two feature cards over a navy band)
function slide04(pres) {
  const s = pres.addSlide();
  glow(s, 14.166, 0.844, 10.27, 0.38);
  rect(s, { x: 0, y: 10.833, w: W, h: 4.167, fill: { color: C.navy } });
  imagePlaceholder(s, 2.025, 2.167, 7.973, 11.333, 0.08863);

  icon(s, 10.881, 2.778, 1.786);
  heading(s, { x: 11.202, w: 8.8, eyebrowY: 6.048, titleY: 4.736, eyebrow: 'The Global Learning Ecosystem', title: 'Our Solution' });
  pillButton(s, 'Learn More', 11.356, 7.248, 3.955, 1.321, 28);

  const cards = [
    { x: 11.359, dark: true, tx: 12.196, tw: 5.335, cx: 12.111, ix: 12.443, iy: 10.272, id: 0.561, t: ['Affordable and scalable learning platform'] },
    { x: 18.359, dark: false, tx: 19.308, tw: 4.261, cx: 19.302, ix: 19.638, iy: 10.249, id: 0.615, t: ['Interactive content', 'with AR/VR'] },
  ];
  cards.forEach(c => {
    roundRect(s, {
      x: c.x, y: 9.343, w: 6.616, h: 4.194, r: 0.12764,
      fill: { color: c.dark ? C.blue : C.white }, shadow: c.dark ? undefined : cardShadow(),
    });
    iconBadge(s, c.cx, 9.967, 1.222, c.ix, c.iy, c.id);
    body(s, c.t.map((t, i) => ({ text: t, options: { breakLine: i < c.t.length - 1 } })), {
      x: c.tx, y: 11.326, w: c.tw, h: 1.456, fontSize: 28, color: c.dark ? C.white : C.ink,
    });
  });
  logo(s);
}

// 5 — Interactive content with AR/VR
function slide05(pres) {
  const s = pres.addSlide();
  roundRect(s, { x: 15.085, y: 9.789, w: 9.546, h: 3.698, r: 0.12764, fill: { color: C.blue } });
  glow(s, 6.371, 2.085, 10.27, 0.38);
  icon(s, 2.052, 1.952, 1.548);
  heading(s, {
    x: 1.867, w: 10.133, titleH: 2.524, eyebrowY: 3.575, titleY: 4.381,
    eyebrow: 'The Global Learning Ecosystem', title: 'Interactive content with AR/VR',
  });
  imagePlaceholder(s, 12.667, 2.816, 12.67, 6.716, 0.11525);
  imagePlaceholder(s, 1.997, 7.538, 12.67, 5.962, 0.11525);
  stat(s, '50%', { x: 15.695, y: 10.833, w: 2.927, h: 1.582, color: C.white });
  label(s, 'Your tittle here', { x: 18.914, y: 10.835, w: 3.434, h: 0.505, color: C.white });
  body(s, LOREM, { x: 18.91, y: 11.34, w: 5.384, h: 1.262, color: C.white });
  logo(s);
}

// 6 — Market Opportunity (world map with callouts)
function slide06(pres) {
  const s = pres.addSlide();
  glow(s, 9.24, 1.698, 10.27, 0.38);

  const box = { x: 6.459, y: 1.875, w: 19.081, h: 11.25 };
  WORLD_MAP.forEach(p => polygon(s, box, p, C.blue));
  [[22.273, 2.167], [8.872, 5.5], [16.402, 9.181], [19.563, 4.456], [12.508, 9.023], [14.082, 1.974]]
    .forEach(([x, y]) => mapPin(s, x, y, 0.79, 1.133));

  icon(s, 1.77, 6.833, 1.56);
  heading(s, {
    x: 1.859, w: 7.472, titleH: 2.524, eyebrowY: 11.373, eyebrowW: 8.774, titleY: 8.524,
    eyebrow: 'EdTech Industry Overview', title: 'Market Opportunity',
  });
  body(s, LOREM_LONG, { x: 1.893, y: 12.334, w: 7.913, h: 1.262 });

  [
    { x: 4.692, y: 2.009, tx: 5.058, ty: 2.378, t: 'Online learning growing by 25% year-over-year' },
    { x: 19.008, y: 8.167, tx: 19.374, ty: 8.536, t: 'Projected to be a $350 Billion industry by 2025' },
  ].forEach(c => {
    s.addShape('wedgeRoundRectCallout', {
      x: c.x, y: c.y, w: 5.335, h: 2.0, fill: { color: C.amber }, line: NO_LINE,
    });
    body(s, c.t, { x: c.tx, y: c.ty, w: 4.604, h: 1.262, align: 'center' });
  });
  logo(s);
}

// 7 — Revenue Tactics (three monetisation cards)
function slide07(pres) {
  const s = pres.addSlide();
  glow(s, 5.998, 1.698, 10.27, 0.38);
  imagePlaceholder(s, 12.667, 1.552, 12.003, 6.615, 0.11525);
  icon(s, 1.997, 3.5, 1.786);
  heading(s, { x: 1.997, w: 9.474, eyebrowY: 7.008, titleY: 5.657, eyebrow: 'Monetization Through Knowledge', title: 'Revenue Tactics' });

  [
    { x: 2.023, dark: true, tx: 3.334, tw: 4.512, cx: 3.330, ix: 3.678, iy: 10.147, t: ['Monthly subscription', 'plans'] },
    { x: 9.687, dark: false, tx: 10.844, tw: 4.820, cx: 10.865, ix: 11.236, iy: 10.133, t: ['Pay-per-course', 'premium content'] },
    { x: 17.351, dark: true, tx: 18.508, tw: 4.820, cx: 18.448, ix: 18.814, iy: 10.186, t: ['Corporate training', 'packages'] },
  ].forEach(c => {
    roundRect(s, {
      x: c.x, y: 8.833, w: 7.293, h: 4.667, r: 0.12764,
      fill: { color: c.dark ? C.blue : C.white }, shadow: cardShadow(),
    });
    oval(s, { x: c.cx, y: 9.8, w: 1.222, h: 1.222, fill: { color: C.amber } });
    icon(s, c.ix, c.iy, 0.521);
    s.addText(c.t.map((t, i) => ({ text: t, options: { breakLine: i < c.t.length - 1 } })), {
      x: c.tx, y: 11.252, w: c.tw, h: 1.043, fontFace: FONT, fontSize: 28, bold: true,
      color: c.dark ? C.white : C.ink, valign: 'top',
    });
  });
  logo(s, 1.929);
}

// 8 — Monthly subscription plans (clustered bar chart)
function slide08(pres) {
  const s = pres.addSlide();
  glow(s, 15.76, 0.532, 10.27, 0.22);
  roundRect(s, { x: 1.997, y: 1.5, w: 11.337, h: 12.0, r: 0.09118, fill: { color: C.white }, shadow: cardShadow() });
  s.addChart(pres.ChartType.bar, [
    { name: 'Series 1', labels: ['Category 1', 'Category 2', 'Category 3'], values: [4.3, 2.5, 3.5] },
    { name: 'Series 2', labels: ['Category 1', 'Category 2', 'Category 3'], values: [2.4, 4.4, 1.8] },
  ], {
    x: 3.023, y: 2.5, w: 9.434, h: 9.667,
    chartColors: [C.amber, C.blue], barGapWidthPct: 90, barOverlapPct: -12,
    showLegend: false, showTitle: false,
    catAxisLabelColor: '595959', valAxisLabelColor: '595959',
    catAxisLabelFontSize: 12, valAxisLabelFontSize: 12, catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT,
    catAxisLineColor: 'D9D9D9', valAxisLineShow: false,
    valGridLine: { color: 'D9D9D9', size: 0.75 }, catGridLine: { style: 'none' },
  });

  icon(s, 14.488, 2.833, 1.691);
  heading(s, {
    x: 14.498, w: 10.367, titleH: 2.524, eyebrowY: 4.611, eyebrowW: 4.642, titleY: 5.318,
    eyebrow: 'Revenue Tactics', title: 'Monthly subscription plans',
  });

  [
    { x: 14.488, fill: C.amber, fg: C.ink, pct: '80%', sx: 15.263, tx: 15.571, bx: 15.134 },
    { x: 20.002, fill: C.blue, fg: C.white, pct: '20%', sx: 20.777, tx: 21.084, bx: 20.647 },
  ].forEach(c => {
    roundRect(s, { x: c.x, y: 8.167, w: 5.335, h: 5.333, r: 0.12764, fill: { color: c.fill }, shadow: cardShadow() });
    stat(s, c.pct, { x: c.sx, y: 9.103, w: 3.786, h: 1.582, color: c.fg, align: 'center' });
    label(s, 'Your tittle here', { x: c.tx, y: 10.685, w: 3.17, h: 0.505, color: c.fg, align: 'center' });
    body(s, [
      { text: 'Lorem ipsum dolor', options: { breakLine: true } },
      { text: 'sit amet.', options: {} },
    ], { x: c.bx, y: 11.431, w: 4.044, h: 1.262, color: c.fg, align: 'center' });
  });
  logo(s);
}

// 9 — Educational Technology (S-curve timeline, 2020-2025)
function slide09(pres) {
  const s = pres.addSlide();
  glow(s, 17.043, -0.65, 10.27, 0.22);
  roundRect(s, { x: 1.997, y: 4.167, w: 22.674, h: 9.333, r: 0.09118, fill: { color: C.white }, shadow: cardShadow() });
  heading(s, {
    x: 6.741, w: 13.185, eyebrowY: 1.665, eyebrowW: 7.576, titleY: 2.372, align: 'center',
    eyebrow: 'Leveraging EdTech Tools', title: 'Educational Technology',
  });

  // Half-period arcs, alternating crest/trough, form the ribbon.
  [[3.368, 6.952, false], [6.566, 8.838, true], [9.757, 6.952, false],
   [12.958, 8.838, true], [16.163, 6.952, false], [19.365, 8.838, true]]
    .forEach(([x, y, flip]) => s.addShape('curvedDownArrow', {
      x, y, w: 4.137, h: 1.886, flipV: flip,
      angleRange: [0.8333, 0.8333], arcThicknessRatio: 1e-9,   // adj1 = adj2 = 50000, adj3 = 0
      fill: { color: C.blue }, line: NO_LINE,
    }));
  s.addShape('triangle', { x: 3.363, y: 8.838, w: 0.951, h: 0.644, flipV: true, fill: { color: C.violet }, line: NO_LINE });
  s.addShape('triangle', { x: 22.559, y: 8.193, w: 0.951, h: 0.644, fill: { color: C.blue }, line: NO_LINE });

  // Milestone icons sit on the ribbon; captions alternate above / below.
  [[4.867, 8.194, 1.139], [8.065, 8.194, 1.139], [11.256, 8.194, 1.139],
   [14.457, 8.194, 1.139], [17.620, 8.151, 1.224], [20.848, 8.151, 1.224]]
    .forEach(([x, y, d]) => icon(s, x, y, d, '6675CD'));

  [
    { yearX: 4.295, yearY: 9.994, year: '2020', tx: 3.855, ty: 4.681, bx: 2.977, by: 5.075 },
    { yearX: 7.538, yearY: 7.025, year: '2021', tx: 7.053, ty: 11.435, bx: 6.175, by: 11.829 },
    { yearX: 10.728, yearY: 10.021, year: '2022', tx: 10.244, ty: 4.651, bx: 9.366, by: 5.046 },
    { yearX: 13.930, yearY: 7.025, year: '2023', tx: 13.446, ty: 11.405, bx: 12.568, by: 11.799 },
    { yearX: 17.135, yearY: 10.105, year: '2024', tx: 16.651, ty: 4.651, bx: 15.773, by: 5.046 },
    { yearX: 20.365, yearY: 7.137, year: '2025', tx: 19.649, ty: 11.552, bx: 18.771, by: 11.947 },
  ].forEach(m => {
    label(s, m.year, { x: m.yearX, y: m.yearY, w: 2.194, h: 0.505, color: C.navy, align: 'center' });
    label(s, 'Tittle here', { x: m.tx, y: m.ty, w: 3.162, h: 0.505, align: 'center' });
    body(s, LOREM, { x: m.bx, y: m.by, w: 4.918, h: 1.069, fontSize: 20, align: 'center' });
  });
}

// 10 — Adaptive learning algorithms
function slide10(pres) {
  const s = pres.addSlide();
  glow(s, 9.116, 0.729, 13.246, 0.22);
  imagePlaceholder(s, 1.997, 2.167, 5.373, 7.114, 0.15745);
  roundRect(s, { x: 7.631, y: 2.833, w: 2.387, h: 2.333, r: 0.2297, fill: { color: C.amber } });
  imagePlaceholder(s, 7.631, 5.5, 5.373, 7.333, 0.15745);
  icon(s, 8.165, 3.243, 1.513);
  roundRect(s, { x: 14.0, y: 8.833, w: 10.67, h: 4.0, r: 0.14351, fill: { color: C.white }, shadow: cardShadow() });

  icon(s, 14.0, 3.577, 1.386);
  heading(s, {
    x: 13.874, w: 10.13, titleH: 2.524, eyebrowY: 5.142, eyebrowW: 7.576, titleY: 5.849,
    eyebrow: 'Educational Technology', title: 'Adaptive learning algorithms',
  });
  stat(s, '1043+', { x: 14.547, y: 9.5, w: 4.072, h: 1.582 });
  label(s, 'Your tittle here', { x: 14.547, y: 10.944, w: 3.434, h: 0.572, fontSize: 28 });
  body(s, LOREM, { x: 19.031, y: 9.576, w: 5.227, h: 1.262 });
  pillButton(s, 'Learn More', 19.093, 11.083, 3.439, 1.084);

  roundRect(s, { x: 3.33, y: 9.5, w: 7.337, h: 2.667, r: 0.16667, fill: { color: C.blue } });
  oval(s, { x: 3.966, y: 10.207, w: 1.222, h: 1.222, fill: { color: C.amber } });
  icon(s, 4.284, 10.524, 0.587);
  body(s, LOREM, { x: 5.383, y: 10.122, w: 4.935, h: 1.262, color: C.white });
  logo(s);
}

// 11 — Gamification for increased engagement
function slide11(pres) {
  const s = pres.addSlide();
  glow(s, 3.648, 3.084, 13.246, 0.22);
  glow(s, 0.666, 0.26, 10.667, 0.22);
  imagePlaceholder(s, 1.997, 2.869, 7.335, 6.249, 0.15745);
  imagePlaceholder(s, 9.711, 2.869, 14.959, 6.249, 0.15745);
  heading(s, {
    x: 1.997, w: 13.305, titleH: 2.524, eyebrowY: 9.943, eyebrowW: 7.576, titleY: 10.739,
    eyebrow: 'Educational Technology', title: 'Gamification for increased engagement',
  });
  label(s, 'Your tittle here', { x: 16.03, y: 9.746, w: 3.434, h: 0.572, fontSize: 28 });
  body(s, LOREM_LONG, { x: 16.0, y: 10.297, w: 7.913, h: 1.262 });
  pillButton(s, 'Learn More', 16.0, 11.794, 3.439, 1.084);
  logo(s);
}

// 12 — Growth Roadmap (amber rail with quarter markers)
function slide12(pres) {
  const s = pres.addSlide();
  glow(s, 14.22, 0.066, 11.515, 0.22);
  roundRect(s, { x: 1.33, y: 5.5, w: 24.007, h: 8.0, r: 0.10956, fill: { color: C.white }, shadow: cardShadow() });
  heading(s, { x: 1.898, w: 13.305, eyebrowY: 3.722, eyebrowW: 11.042, titleY: 2.255, eyebrow: 'Charting the Course of Education', title: 'Growth Roadmap' });
  label(s, 'Your tittle here', { x: 16.72, y: 2.372, w: 3.434, h: 0.572, fontSize: 28 });
  body(s, LOREM_LONG, { x: 16.69, y: 2.923, w: 7.913, h: 1.262 });

  roundRect(s, { x: 1.997, y: 8.885, w: 22.674, h: 1.333, r: 0.5, fill: { color: C.amber } });
  s.addShape('line', { x: 2.395, y: 9.526, w: 21.519, h: 0, line: { color: C.black, width: 1.75 } });

  [
    { flagX: 3.521, flagY: 6.778, dotX: 5.646, dotY: 9.302, stemY: 8.219, capX: 4.117, capY: 10.758, capW: 3.506, bodyX: 3.834, bodyY: 11.437, q: 'Q1 - 2023', cap: 'Mobile app release', txtX: 4.341, txtY: 7.232 },
    { flagX: 10.444, flagY: 10.878, dotX: 12.568, dotY: 9.327, stemY: 9.768, capX: 10.462, capY: 6.767, capW: 4.662, bodyX: 10.757, bodyY: 7.277, q: 'Q3 - 2023', cap: 'Partnership with schools', txtX: 11.264, txtY: 11.332 },
    { flagX: 17.908, flagY: 6.813, dotX: 20.032, dotY: 9.336, stemY: 8.253, capX: 17.367, capY: 10.792, capW: 5.778, bodyX: 18.221, bodyY: 11.472, q: 'Q1 - 2024', cap: 'Global language support', txtX: 18.728, txtY: 7.267 },
  ].forEach(m => {
    s.addShape('line', { x: m.dotX + 0.224, y: m.stemY, w: 0, h: 1.083, line: { color: C.black, width: 1.75 } });
    s.addShape('flowChartDisplay', { x: m.flagX, y: m.flagY, w: 4.697, h: 1.44, flipH: true, fill: { color: C.blue }, line: NO_LINE });
    s.addText(m.q, {
      x: m.txtX, y: m.txtY, w: 3.057, h: 0.64,
      fontFace: FONT, fontSize: 32, bold: true, color: C.white, align: 'center', valign: 'top',
    });
    oval(s, { x: m.dotX, y: m.dotY, w: 0.448, h: 0.448, fill: { color: C.navy } });
    label(s, m.cap, { x: m.capX, y: m.capY, w: m.capW, h: 0.505, align: 'center' });
    body(s, LOREM, { x: m.bodyX, y: m.bodyY, w: 4.071, h: 1.069, fontSize: 20, align: 'center' });
  });
  logo(s);
}

// 13 — Competitive Positioning (comparison table)
function slide13(pres) {
  const s = pres.addSlide();
  glow(s, 7.998, -0.692, 13.246, 0.22);
  icon(s, 1.997, 1.853, 1.616);
  heading(s, { x: 1.997, w: 13.305, eyebrowY: 3.5, eyebrowW: 11.042, titleY: 4.387, eyebrow: 'Standing Out in the Crowd', title: 'Competitive Positioning' });
  imagePlaceholder(s, 16.668, 2.869, 8.002, 10.667, 0.15745);

  const cell = (text, fill, color, bold) => ({
    text, options: { fill: { color }, color: color === undefined ? C.black : color },
  });
  const bodyCell = fill => ({
    text: [
      { text: 'Lorem ipsum dolor sit amet,', options: { breakLine: true } },
      { text: 'consectetur adipiscing elit.', options: {} },
    ],
    options: { fill: { color: fill }, color: C.black, fontSize: 24, lineSpacingMultiple: 1.5 },
  });
  s.addTable([
    [
      { text: 'Coursera', options: { fill: { color: C.navy }, color: C.white, fontSize: 32, bold: true } },
      { text: 'Udemy', options: { fill: { color: C.amber }, color: C.black, fontSize: 32, bold: true } },
    ],
    [bodyCell('CDCDD8'), bodyCell('CDCDD8')],
    [bodyCell('E7E8EC'), bodyCell('E7E8EC')],
    [bodyCell('CDCDD8'), bodyCell('CDCDD8')],
  ], {
    x: 1.997, y: 6.176, w: 14.003, colW: [7.002, 7.002], rowH: [1.831, 1.831, 1.831, 1.831],
    fontFace: FONT, align: 'center', valign: 'middle', border: { type: 'none' },
  });
  logo(s);
}

// 14 — Competitive Positioning (stat cards)
function slide14(pres) {
  const s = pres.addSlide();
  imagePlaceholder(s, 8.665, 6.869, 16.005, 6.667, 0.16162);
  glow(s, 12.837, -0.531, 13.246, 0.22);
  icon(s, 12.037, 2.141, 1.963);
  heading(s, { x: 11.9, w: 13.305, eyebrowY: 4.055, eyebrowW: 11.042, titleY: 4.942, eyebrow: 'Standing Out in the Crowd', title: 'Competitive Positioning' });

  roundRect(s, { x: 2.663, y: 2.141, w: 8.002, h: 6.692, r: 0.11772, fill: { color: C.blue } });
  stat(s, '65%', { x: 3.893, y: 3.971, w: 3.439, h: 1.582, color: C.white });
  label(s, 'Your tittle here', { x: 3.897, y: 5.662, w: 3.434, h: 0.505, color: C.white });
  body(s, LOREM_SHORT, { x: 3.893, y: 6.167, w: 5.29, h: 1.262, color: C.white });

  roundRect(s, { x: 1.895, y: 9.172, w: 6.37, h: 3.661, r: 0.1521, fill: { color: C.amber } });
  icon(s, 2.497, 9.675, 0.731);
  label(s, 'Your tittle here', { x: 2.589, y: 10.521, w: 3.434, h: 0.505 });
  body(s, LOREM_SHORT, { x: 2.585, y: 11.026, w: 5.29, h: 1.262 });
  logo(s);
}

// 15 — High cost of educational materials (exploded pie)
function slide15(pres) {
  const s = pres.addSlide();
  glow(s, 5.086, 0.254, 13.246, 0.22);
  icon(s, 1.997, 2.823, 1.213);
  heading(s, {
    x: 1.867, w: 12.133, titleH: 2.524, eyebrowY: 4.204, eyebrowW: 4.282, titleY: 4.904,
    eyebrow: 'The Challenge', title: [
      { text: 'High cost of', options: { breakLine: true } },
      { text: 'Educational materials', options: {} },
    ],
  });
  s.addChart(pres.ChartType.pie, [
    { name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: [60, 40] },
  ], {
    x: 12.03, y: 2.822, w: 15.778, h: 10.519,
    chartColors: [C.navy, C.amber], dataBorder: { pt: 1.5, color: C.white },
    showLegend: false, showTitle: false, firstSliceAng: 0,
  });
  icon(s, 21.513, 8.288, 1.786, C.white);
  icon(s, 16.847, 6.062, 1.786);

  roundRect(s, { x: 1.997, y: 8.167, w: 11.337, h: 4.667, r: 0.10956, fill: { color: C.white }, shadow: cardShadow() });
  [{ x: 2.599, dot: C.amber }, { x: 7.95, dot: C.blue }].forEach(c => {
    oval(s, { x: c.x + 0.042, y: 9.254, w: 0.667, h: 0.667, fill: { color: c.dot } });
    label(s, 'Your tittle here', { x: c.x + 0.004, y: 10.118, w: 3.434, h: 0.505 });
    body(s, LOREM, { x: c.x, y: 10.623, w: 5.29, h: 1.262 });
  });
  logo(s);
}

// 16 — Core Team (single profile card)
function slide16(pres) {
  const s = pres.addSlide();
  roundRect(s, { x: 1.997, y: 1.5, w: 4.216, h: 4.0, r: 0.20438, fill: { color: C.amber }, shadow: cardShadow() });
  imagePlaceholder(s, 3.144, 2.833, 10.189, 10.0, 0.14111);
  glow(s, 15.294, 1.896, 10.27, 0.38);
  icon(s, 14.882, 4.607, 1.786);
  heading(s, { x: 14.616, w: 6.768, eyebrowY: 7.653, eyebrowW: 10.189, titleY: 6.34, eyebrow: 'Educators and Technologists United', title: 'Core Team' });

  roundRect(s, { x: 11.993, y: 9.5, w: 6.002, h: 2.667, r: 0.18825, fill: { color: C.blue }, shadow: cardShadow() });
  s.addText([
    { text: 'Michael Roberts', options: { breakLine: true } },
    { text: '(EdTech Innovator)', options: {} },
  ], { x: 12.764, y: 10.244, w: 4.959, h: 1.178, fontFace: FONT, fontSize: 32, bold: true, color: C.white, valign: 'top' });
  roundRect(s, { x: 10.996, y: 8.705, w: 1.581, h: 1.5, r: 0.18825, fill: { color: C.amber }, shadow: cardShadow() });
  s.addText('CEO', { x: 11.127, y: 9.136, w: 1.318, h: 0.64, fontFace: FONT, fontSize: 32, bold: true, color: C.ink, align: 'center', valign: 'top' });

  label(s, 'Your tittle here', { x: 18.669, y: 10.067, w: 3.434, h: 0.572, fontSize: 28 });
  body(s, LOREM_LONG, { x: 18.639, y: 10.618, w: 7.365, h: 1.262 });
  logo(s);
}

// 17 — Core Team (roster on a blue panel)
function slide17(pres) {
  const s = pres.addSlide();
  imagePlaceholder(s, 0, 0, 12, 15, 0);
  glow(s, 10.87, 3.079, 10.27, 0.38);
  roundRect(s, { x: 15.333, y: 8.851, w: 9.976, h: 4.0, r: 0.14351, fill: { color: C.white }, shadow: cardShadow() });
  rect(s, { x: 0.037, y: 0, w: 12, h: 15, fill: { color: C.blue, transparency: 12 } });
  icon(s, 15.333, 4.381, 1.786);
  heading(s, { x: 15.202, w: 6.768, eyebrowY: 7.652, eyebrowW: 9.976, titleY: 6.339, eyebrow: 'Educators and Technologists United', title: 'Core Team' });

  [
    { y: 2.586, ry: 3.319, line: 4.167, x: 2.663, name: 'Michael Roberts (CEO)', role: 'EdTech Innovator' },
    { y: 6.881, ry: 7.614, line: 8.667, x: 2.693, name: 'Rachel Kim (CTO)', role: 'Software Development' },
    { y: 11.176, ry: 11.909, line: 12.851, x: 2.595, name: 'David Zha (CMO)', role: 'Marketing Strategist' },
  ].forEach(p => {
    s.addText(p.name, { x: p.x, y: p.y, w: 5.499, h: 0.64, fontFace: FONT, fontSize: 32, bold: true, color: C.white, valign: 'top' });
    label(s, p.role, { x: p.x, y: p.ry, w: 4.959, h: 0.505, color: C.white });
    s.addShape('line', { x: 2.693, y: p.line, w: 6.639, h: 0, line: { color: C.white, width: 2 } });
  });

  stat(s, '1043+', { x: 15.929, y: 9.433, w: 3.724, h: 1.582 });
  label(s, 'Your tittle here', { x: 15.929, y: 11.152, w: 3.434, h: 0.572, fontSize: 28 });
  body(s, LOREM, { x: 19.955, y: 9.783, w: 5.227, h: 1.262 });
  pillButton(s, 'Learn More', 19.934, 11.101, 3.439, 1.084);
  logo(s);
}

// 18 — Investment Details (allocation pie)
function slide18(pres) {
  const s = pres.addSlide();
  glow(s, 9.682, 0.247, 10.27, 0.38);
  roundRect(s, { x: 1.997, y: 1.5, w: 11.337, h: 12.0, r: 0.10956, fill: { color: C.white }, shadow: cardShadow() });
  s.addChart(pres.ChartType.pie, [
    { name: 'Sales', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr'], values: [35, 45, 20] },
  ], {
    x: 0.879, y: 3.101, w: 13.571, h: 9.048,
    chartColors: [C.teal, C.blue, C.amber], dataBorder: { pt: 1.5, color: C.white },
    showLegend: false, showTitle: false, firstSliceAng: 0,
  });
  [['35%', 8.331, 6.383], ['45%', 4.997, 9.266], ['20%', 4.330, 5.070]].forEach(([t, x, y]) => {
    s.addText(t, { x, y, w: 3.335, h: 1.212, fontFace: FONT, fontSize: 66, bold: true, color: C.white, align: 'center', valign: 'top' });
  });

  icon(s, 14.595, 4.667, 1.786);
  heading(s, { x: 14.595, w: 9.976, eyebrowY: 6.477, titleY: 7.184, eyebrow: 'Investing in the Future of Education', title: 'Investment Details' });
  stat(s, '$8M', { x: 14.595, y: 9.433, w: 3.406, h: 1.582 });
  label(s, 'Your tittle here', { x: 18.235, y: 9.37, w: 3.434, h: 0.572, fontSize: 28 });
  body(s, LOREM_LONG, { x: 18.205, y: 9.921, w: 7.365, h: 1.262 });
  logo(s);
}

// 19 — Investment Details (photo layout)
function slide19(pres) {
  const s = pres.addSlide();
  glow(s, 0.396, 0.833, 10.27, 0.38);
  imagePlaceholder(s, 2.015, 6.167, 15.32, 7.428, 0.16265);
  imagePlaceholder(s, 14.0, 2.833, 11.319, 7.428, 0.13338);
  roundRect(s, { x: 12.836, y: 1.758, w: 2.889, h: 2.742, r: 0.20438, fill: { color: C.amber }, shadow: cardShadow() });
  icon(s, 1.92, 2.071, 1.41);
  heading(s, { x: 1.859, w: 9.976, eyebrowY: 3.48, titleY: 4.187, eyebrow: 'Investing in the Future of Education', title: 'Investment Details' });

  roundRect(s, { x: 17.709, y: 10.598, w: 4.96, h: 2.366, r: 0.20198, fill: { color: C.blue } });
  icon(s, 18.216, 11.28, 1.119, C.white);
  body(s, [
    { text: 'Seeking $8M', options: { breakLine: true } },
    { text: 'for 25% equity', options: {} },
  ], { x: 19.335, y: 11.208, w: 2.841, h: 1.262, color: C.white });
  logo(s);
}

// 20 — Thank You
function slide20(pres) {
  const s = pres.addSlide();
  portalBackdrop(s);
  s.addText('Learn more about us', {
    x: 2.583, y: 7.481, w: 12.003, h: 1.01,
    fontFace: FONT_MED, fontSize: 54, color: C.ink, valign: 'top',
  });
  s.addText('Thank You', {
    x: 2.583, y: 8.658, w: 13.596, h: 2.036,
    fontFace: FONT, fontSize: 115, bold: true, color: C.ink, valign: 'top',
  });
  icon(s, 2.718, 5.683, 1.604);
  pillButton(s, 'Join Us!', 2.652, 11.176, 3.439, 1.084);
  logo(s);
}

// ---------------------------------------------------------------- build ----
function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'DIGEDU', width: W, height: H });
  pres.layout = 'DIGEDU';
  pres.theme = { headFontFace: 'Michroma', bodyFontFace: FONT };
  pres.title = 'Shaping Future Generations';
  pres.author = 'DigEdu';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
   slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach(fn => fn(pres));

  const out = path.join(__dirname, '0d3f6f24-0a45-420b-a4a3-447d08f2b529_grok_final.pptx');
  return pres.writeFile({ fileName: out }).then(() => console.log('wrote', out));
}

build().catch(err => { console.error(err); process.exit(1); });
