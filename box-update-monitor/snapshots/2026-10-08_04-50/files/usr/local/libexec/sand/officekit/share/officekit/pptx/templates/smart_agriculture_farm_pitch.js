/**
 * "Green Valley Farm" — Smart Agriculture business deck.
 * Standalone pptxgenjs recreation: 20 slides, 20 x 11.25 in (16:9 widescreen).
 *
 * The photo frames in the source deck are empty picture placeholders (they
 * carry no bitmap and render as transparent); the one real raster asset — the
 * sprout icon on slide 15 — is replaced by a labelled placeholder rectangle.
 */
'use strict';

const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ *
 * Theme
 * ------------------------------------------------------------------ */
const W = 20, H = 11.25;                       // slide size (inches)

const WHITE = 'FFFFFF';
const BLACK = '000000';
const GREY_TX = '404040';                      // theme dk2
const OLIVE = '565F3B';                        // accent3 - brand green
const OLIVE_MID = '6A7448';                    // accent2
const SAGE = '9FB191';                         // accent1
const SAGE_75 = '768D64';                      // accent1 lum 75%
const SAGE_PALE = 'C5D0BD';                    // accent1 lum 60/40
const OLIVE_PALE = 'ABB588';                   // accent2 lum 60/40
const MINT_PALE = 'E0E4D3';                    // accent3 lum 20/80
const OLIVE_LIGHT = 'C0C8A7';                  // accent3 lum 40/60
const OLIVE_MID2 = '606A42';                   // accent3 lum 95/5
const OLIVE_DEEP = '343923';                   // accent3 lum 60
const PHOTO = 'CFCFCF';                        // placeholder grey

const HEAD = 'Space Grotesk';                  // theme major font
const BODY = 'Work Sans';                      // theme minor font

/* ------------------------------------------------------------------ *
 * Copy used on several slides
 * ------------------------------------------------------------------ */
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, conse lectus Pellentesque scelerisque malesuada libero a libero.';
const LOREM_CARD = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.';
const LOREM_GRID = 'Lorem ipsum dolor sit amet, conse lectus ornare, viverra ctetur adipiscing elit. Pellentesque scelerisque.';
const PRAESENT = 'Praesent convallis ligula eu diam porttitor sodales. Cras imperdiet vel metus ac condimentum. Curabitur fermentum sapien diam, sed volutpat est scelerisque vitae. Suspendisse in sagittis nunc.';
const PRAESENT_PURUS = PRAESENT.replace('nunc.', 'nunc, nec tempus purus.');
const PRAESENT_SHORT = 'Praesent convallis ligula eu diam porttitor sodales. Cras imperdiet vel metus ac condimentum.';
const DUIS = 'Duis nec tempus nunc, non tempus risus. Nam felis velit, ultricies quis leo eget, volutpat auctor lorem. Curabitur ut gravida lectus. Cras leo lorem, vulputate ut consectetur dignissim, eleifend vel magna. Fusce ut porttitor quam, a laoreet velit.';
const CURABITUR = 'Curabitur suscipit tellus enim, id aliquet nisi maximus eu. Aliquam erat volutpat. Sed feugiat semper lacus, eget blandit neque vehicula ac.';
const NUNC = 'Nunc euismod risus eu nibh bibendum, at condimentum dolor interdum ullamcorper.';
const FUTURE_LINE = 'The future of farming lies in the hands of innovation';
const HARVEST = 'The harvest of the future begins with the ideas of today.';
const WISEST = 'Agriculture is our wisest pursuit, because it will in the end contribute most to real wealth';
const EVOLVES = 'Agriculture evolves when innovation leads.';
const ROOT_OF_GROWTH = 'Innovation at the Root of Growth';

/* ------------------------------------------------------------------ *
 * Small helpers
 * ------------------------------------------------------------------ */
const hex2rgb = (h) => [0, 2, 4].map((i) => parseInt(h.substr(i, 2), 16));
const rgb2hex = (c) => c.map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).toUpperCase().padStart(2, '0')).join('');
const mix = (a, b, t) => rgb2hex(hex2rgb(a).map((v, i) => v + (hex2rgb(b)[i] - v) * t));
const darken = (base, alpha) => mix(base, BLACK, alpha);

/** colour of a multi-stop gradient at position t (0..1) */
function gradientAt(stops, t) {
  for (let i = 1; i < stops.length; i++) {
    if (t <= stops[i][0] || i === stops.length - 1) {
      const [p0, c0] = stops[i - 1], [p1, c1] = stops[i];
      return mix(c0, c1, Math.max(0, Math.min(1, (t - p0) / (p1 - p0))));
    }
  }
  return stops[0][1];
}

/** text box: source deck uses zero insets, top anchored, auto-fit boxes */
function txt(slide, str, o) {
  slide.addText(str, Object.assign({
    margin: 0, valign: 'top', wrap: true, fontFace: BODY, fontSize: 18,
    color: WHITE, charSpacing: -0.44
  }, o));
}
/** display headline (Space Grotesk bold) */
function head(slide, str, o) {
  txt(slide, str, Object.assign({
    fontFace: HEAD, bold: true, fontSize: 56, lineSpacingMultiple: 0.85, charSpacing: 0
  }, o));
}
/** 28pt light "quote" line (Work Sans) */
function quote(slide, str, o) {
  txt(slide, str, Object.assign({ fontSize: 28, lineSpacingMultiple: 1.1, charSpacing: -0.89 }, o));
}
/**
 * 24pt bold section label. Unlike the other text boxes these keep PowerPoint's
 * default 0.05" top/bottom inset, which nudges the baseline down slightly.
 */
function subhead(slide, str, o) {
  txt(slide, str, Object.assign({
    fontFace: HEAD, bold: true, fontSize: 24, charSpacing: 0, margin: [0, 0, 3.6, 3.6]
  }, o));
}
function rect(slide, o) { slide.addShape('rect', o); }
function round(slide, o) { slide.addShape('roundRect', o); }

/** stand-in for the deck's one embedded raster asset */
function imagePlaceholder(slide, x, y, w, h, radius) {
  round(slide, { x, y, w, h, rectRadius: radius || 0.1, fill: { color: PHOTO } });
  txt(slide, '[image]', {
    x, y: y + h / 2 - 0.2, w, h: 0.4, align: 'center', fontSize: 12,
    color: '6E6E6E', charSpacing: 0
  });
}

/**
 * Full-bleed photo backdrop: the source slides place a picture behind a radial
 * black scrim that fades from 20% (centre) to 70% (corners). Rebuilt here as a
 * stack of concentric discs since pptxgenjs has no gradient fill.
 */
function scrim(slide, base, steps) {
  const N = steps || 72, R = Math.hypot(W, H) / 2;
  rect(slide, { x: 0, y: 0, w: W, h: H, fill: { color: darken(base, 0.7) } });
  for (let k = N - 1; k >= 0; k--) {
    const t = k / N, r = R * t;
    slide.addShape('ellipse', {
      x: W / 2 - r, y: H / 2 - r, w: 2 * r, h: 2 * r,
      fill: { color: darken(base, 0.2 + 0.5 * t) }
    });
  }
}

/** radial colour gradient (slide 17) built the same way */
function radial(slide, stops, steps) {
  const N = steps || 72, R = Math.hypot(W, H) / 2;
  rect(slide, { x: 0, y: 0, w: W, h: H, fill: { color: gradientAt(stops, 1) } });
  for (let k = N - 1; k >= 0; k--) {
    const t = k / N, r = R * t;
    slide.addShape('ellipse', {
      x: W / 2 - r, y: H / 2 - r, w: 2 * r, h: 2 * r,
      fill: { color: gradientAt(stops, t) }
    });
  }
}

/**
 * Rectangular gradient whose focus sits in the bottom-right corner (slide 18).
 * Rebuilt as nested rectangles pinned to that corner, which reproduces the
 * L-shaped bands of the original path="rect" gradient fill.
 */
function cornerGlow(slide, stops, steps) {
  const N = steps || 80, SPAN_X = 0.62 * W, SPAN_Y = 0.62 * H;
  for (let k = N; k >= 0; k--) {
    const t = k / N;                       // t = 1 at the outer edge of the glow
    rect(slide, {
      x: W - SPAN_X * t, y: H - SPAN_Y * t, w: SPAN_X * t, h: SPAN_Y * t,
      fill: { color: gradientAt(stops, Math.pow(1 - t, 0.55)) }
    });
  }
}

/**
 * Band of light along the top edge, fading downwards and dying out towards the
 * left (slide 3). Each row is one wide rectangle plus a short ramp of narrow
 * ones on its left flank.
 */
function topGlow(slide, x0, depth, from, to) {
  const ROWS = 26, RAMP = 8, cw = 0.26;
  for (let r = 0; r < ROWS; r++) {
    const lit = mix(from, to, r / (ROWS - 1));
    const y = (depth * r) / ROWS, h = depth / ROWS + 0.02;
    rect(slide, { x: x0 + RAMP * cw, y, w: W - x0 - RAMP * cw, h, fill: { color: lit } });
    for (let c = 0; c < RAMP; c++) {
      rect(slide, { x: x0 + c * cw, y, w: cw + 0.01, h, fill: { color: mix(to, lit, (c + 1) / RAMP) } });
    }
  }
}

/** top navigation bar, present on every slide */
function navbar(slide, color, brandOnly) {
  round(slide, {
    x: 0.727, y: 0.758, w: 3.189, h: 0.486, rectRadius: 0.243,
    line: { color: color, width: 1.5 }
  });
  txt(slide, 'Green Valley Farm', {
    x: 0.727, y: 0.833, w: 3.189, h: 0.337, align: 'center',
    fontFace: HEAD, bold: true, fontSize: 20, color, lineSpacingMultiple: 1
  });
  if (brandOnly) return;
  [['Agriculture', 4.922, 2.186], ['Business Project', 7.108, 2.599],
   ['2025', 10.295, 1.595], ['Contact Us', 17.679, 1.595]].forEach(([s, x, w]) => {
    txt(slide, s, {
      x, y: 0.833, w, h: 0.337, fontFace: HEAD, bold: true, fontSize: 20,
      color, lineSpacingMultiple: 1
    });
  });
}

/** full-width strip of three repeated taglines at the foot of the slide */
function tagStrip(slide, band, ink) {
  rect(slide, { x: 0, y: 10.359, w: W, h: 0.891, fill: { color: band } });
  [0.727, 7.106, 13.484].forEach((x) => {
    txt(slide, ROOT_OF_GROWTH, {
      x, y: 10.63, w: 5.788, h: 0.35, fontFace: HEAD, bold: true, fontSize: 24,
      color: ink, lineSpacingMultiple: 0.85, charSpacing: 0
    });
  });
}

/** big statistic + rule + caption (slides 3 and 8) */
function statBlock(slide, x, opts) {
  head(slide, opts.value, { x, y: opts.y, w: opts.w, h: 1.166, fontSize: 80, color: WHITE });
  slide.addShape('line', {
    x, y: opts.y + opts.ruleDy, w: 0.777, h: 0,
    line: { color: WHITE, width: 2.25 }
  });
  subhead(slide, opts.label, { x, y: opts.y + opts.ruleDy + 0.394, w: opts.w, h: 0.909 });
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

// 1 — Cover
function slide01(pres) {
  const s = pres.addSlide();
  scrim(s, WHITE);
  head(s, 'Sustainable Growth Through Smart Agriculture',
    { x: 0.727, y: 3.554, w: 13.762, h: 4.143, fontSize: 96 });
  navbar(s, WHITE);
  round(s, { x: 15.078, y: 3.804, w: 4.196, h: 4.196, rectRadius: 0.699, fill: { color: WHITE, transparency: 50 } });
  quote(s, FUTURE_LINE, { x: 0.727, y: 8.166, w: 10.573, h: 0.401, lineSpacingMultiple: 0.85, charSpacing: 0 });
  round(s, { x: 15.581, y: 6.617, w: 3.191, h: 0.8, rectRadius: 0.25, fill: { color: OLIVE } });
  txt(s, 'Smart Agriculture', {
    x: 15.581, y: 6.617, w: 3.191, h: 0.8, align: 'center', valign: 'middle',
    fontSize: 18, charSpacing: 0
  });
}

// 2 — Agenda, four numbered cards
function slide02(pres) {
  const s = pres.addSlide();
  s.background = { color: OLIVE };
  scrim(s, OLIVE);
  head(s, 'Growing Business from the Ground Up', { x: 0.727, y: 1.652, w: 12.167, h: 1.618 });
  [['Brand Guidelines', 'for ABC Company', 13.488], ['by  John Smith', '2025', 16.677]].forEach(([a, b, x]) => {
    txt(s, [{ text: a, options: { breakLine: true } }, { text: b }], {
      x, y: 1.652, w: 2.597, h: 0.725, fontFace: HEAD, bold: true, fontSize: 20,
      lineSpacingMultiple: 1.1, charSpacing: -0.89
    });
  });
  const cards = [
    { x: 0.727, n: '01', label: 'Introduction', lw: 2.601, dy: 0 },
    { x: 5.510, n: '02', label: 'Opportunity', lw: 2.601, dy: 0.018 },
    { x: 10.293, n: '03', label: 'Our Products', lw: 3.615, dy: 0.018 },
    { x: 15.076, n: '04', label: 'Conclusion', lw: 2.601, dy: 0.018 }
  ];
  cards.forEach((c) => {
    round(s, { x: c.x, y: 4.02, w: 4.196, h: 6.506, rectRadius: 0.699, fill: { color: WHITE, transparency: 80 } });
    head(s, c.n, { x: c.x + 0.582, y: 4.799 + c.dy, w: 2.412, h: 1.05, fontSize: 72, valign: 'middle' });
    txt(s, c.label, {
      x: c.x + 0.582, y: 7.827 + c.dy, w: c.lw, h: 0.471, fontFace: HEAD, bold: true,
      fontSize: 28, lineSpacingMultiple: 1, charSpacing: -0.69
    });
    txt(s, LOREM_SHORT, { x: c.x + 0.582, y: 8.499 + c.dy, w: 3.034, h: 1.212, lineSpacingMultiple: 1 });
  });
  navbar(s, WHITE);
}

// 3 — Farming the Future with Smart Tech
function slide03(pres) {
  const s = pres.addSlide();
  s.background = { color: OLIVE_DEEP };
  topGlow(s, 4.9, 1.35, '52583F', OLIVE_DEEP);   // faint band of light across the header
  head(s, 'Farming the Future with Smart Tech', { x: 0.727, y: 1.674, w: 7.54, h: 1.605 });
  txt(s, PRAESENT, { x: 15.078, y: 8.708, w: 4.22, h: 1.818, lineSpacingMultiple: 1 });
  statBlock(s, 0.727, { value: '89%', y: 7.142, w: 2.677, ruleDy: 2.081, label: 'Promotion Management' });
  quote(s, FUTURE_LINE, { x: 15.138, y: 4.384, w: 4.136, h: 1.414, lineSpacingMultiple: 1, charSpacing: -0.44 });
  head(s, '2025', { x: 15.078, y: 7.16, w: 4.196, h: 0.816 });
  navbar(s, WHITE);
}

// 4 — Sustainable Growth through Smart Agriculture (four pills)
function slide04(pres) {
  const s = pres.addSlide();
  s.background = { color: OLIVE };
  head(s, 'Sustainable Growth through Smart Agriculture',
    { x: 7.108, y: 2.177, w: 12.167, h: 1.618, align: 'center' });
  quote(s, 'Smart farming sows the seeds of a better world',
    { x: 7.108, y: 4.32, w: 12.167, h: 0.486, align: 'center' });
  [{ n: '01', x: 7.108, y: 5.821 }, { n: '02', x: 13.484, y: 5.825 },
   { n: '03', x: 7.104, y: 7.993 }, { n: '04', x: 13.484, y: 8.004 }].forEach((c) => {
    round(s, { x: c.x, y: c.y, w: 5.79, h: 1.583, rectRadius: 0.355, fill: { color: WHITE, transparency: 80 } });
    txt(s, c.n, {
      x: c.x + 0.505, y: c.y + 0.337, w: 1.092, h: 0.471, fontFace: HEAD, bold: true,
      fontSize: 28, lineSpacingMultiple: 1, charSpacing: -0.52
    });
    txt(s, LOREM_CARD, { x: c.x + 1.597, y: c.y + 0.337, w: 3.986, h: 0.909, lineSpacingMultiple: 1 });
  });
  // circular badge
  slide04Badge(s);
  tagStrip(s, WHITE, BLACK);
  navbar(s, WHITE);
}
function slide04Badge(s) {
  s.addShape('ellipse', { x: 5.791, y: 2.801, w: 1.449, h: 1.449, fill: { color: WHITE } });
  txt(s, 'From the ground we grow, from strategy we thrive', {
    x: 5.791, y: 2.801, w: 1.449, h: 1.449, align: 'center', valign: 'middle',
    fontFace: HEAD, bold: true, fontSize: 7, color: BLACK, lineSpacingMultiple: 1.1
  });
}

// 5 — Smart Agriculture, Stronger Yields
function slide05(pres) {
  const s = pres.addSlide();
  round(s, { x: 7.108, y: 4.741, w: 5.785, h: 0.943, rectRadius: 0.408, fill: { color: OLIVE } });
  round(s, { x: 13.633, y: 4.691, w: 5.785, h: 0.943, rectRadius: 0.408, fill: { color: OLIVE } });
  txt(s, 'From Fields to Finance', {
    x: 8.114, y: 5.037, w: 3.772, h: 0.35, fontFace: HEAD, bold: true, fontSize: 24,
    lineSpacingMultiple: 0.85, charSpacing: 0
  });
  txt(s, 'The Future of Food Starts Here', {
    x: 13.987, y: 4.987, w: 5.077, h: 0.35, fontFace: HEAD, bold: true, fontSize: 24,
    lineSpacingMultiple: 0.85, charSpacing: 0
  });
  quote(s, EVOLVES, { x: 1.299, y: 2.185, w: 4.608, h: 1.013 });
  head(s, 'Smart Agriculture, Stronger Yields', { x: 1.299, y: 7.27, w: 7.974, h: 1.618 });
  txt(s, [
    { text: DUIS + ' ', options: { breakLine: true, paraSpaceAfter: 12 } },
    { text: 'Ut vehicula enim eu dui volutpat efficitur. Maecenas scelerisque eget risus sed convallis. Proin venenatis accumsan augue, vel rhoncus elit consectetur eu.' }
  ], { x: 10.295, y: 7.27, w: 7.974, h: 2.289, lineSpacingMultiple: 1 });
  navbar(s, BLACK);
}

// 6 — Reimagining Agriculture with Technology
function slide06(pres) {
  const s = pres.addSlide();
  head(s, 'Reimagining Agriculture with Technology', { x: 0.727, y: 1.711, w: 18.547, h: 0.816, color: BLACK });
  subhead(s, 'Crowd Funding', { x: 0.727, y: 3.122, w: 2.599, h: 0.505, color: BLACK });
  txt(s, 'Mauris ultrices turpis mi, a ultricies leo fringilla non. Nulla sodales ullamcorper diam vel maximus.',
    { x: 0.727, y: 4.038, w: 2.599, h: 1.515, color: GREY_TX, lineSpacingMultiple: 1 });
  round(s, { x: 3.767, y: 4.85, w: 3.781, h: 5.676, rectRadius: 0.63, fill: { color: SAGE_PALE } });
  quote(s, 'The harvest of the future begins with the ideas of today."', { x: 4.294, y: 5.41, w: 2.813, h: 2.041, color: BLACK });
  txt(s, 'Nunc euismod risus eu nibh bibendum, at condimentum dolor interdum ullamcorper.',
    { x: 4.294, y: 8.72, w: 2.813, h: 1.212, color: BLACK, lineSpacingMultiple: 1 });
  head(s, '30%', { x: 13.484, y: 7.505, w: 2.356, h: 0.801 });
  s.addShape('upArrow', { x: 16.674, y: 4.835, w: 0.724, h: 0.772, fill: { color: BLACK } });
  head(s, '64%', { x: 17.595, y: 4.883, w: 2.27, h: 0.913, fontSize: 60, color: BLACK, lineSpacingMultiple: 0.9 });
  round(s, { x: 16.675, y: 6.367, w: 2.599, h: 4.159, rectRadius: 0.433, fill: { color: SAGE_PALE } });
  quote(s, 'Innovation is the true fertilizer of sustainable growth.', { x: 16.976, y: 6.816, w: 2.13, h: 2.56, color: BLACK });
  navbar(s, BLACK);
}

// 7 — Farming Innovation for a Better Tomorrow
function slide07(pres) {
  const s = pres.addSlide();
  head(s, 'Farming Innovation for a Better Tomorrow', { x: 0.715, y: 2.142, w: 8.285, h: 1.618, color: BLACK });
  quote(s, WISEST, { x: 0.717, y: 7.023, w: 6.806, h: 1.531, bold: true, color: BLACK });
  txt(s, DUIS, { x: 0.717, y: 9.314, w: 7.669, h: 1.212, color: BLACK, lineSpacingMultiple: 1 });
  navbar(s, BLACK);
}

// 8 — Next-Generation Farming Starts Here
function slide08(pres) {
  const s = pres.addSlide();
  head(s, 'Next-Generation Farming Starts Here', { x: 0.727, y: 2.083, w: 8.979, h: 1.618, color: BLACK });
  [['2024', 10.295], ['2025', 15.054]].forEach(([y, x], i) => {
    head(s, y, { x, y: i ? 2.146 : 2.13, w: 2.876, h: 0.816, color: BLACK });
    txt(s, PRAESENT, { x, y: i ? 3.051 : 3.034, w: 4.22, h: 1.818, color: BLACK, lineSpacingMultiple: 1 });
  });
  round(s, { x: 0.727, y: 5.633, w: 8.979, h: 4.893, rectRadius: 0.815, fill: { color: OLIVE_MID } });
  round(s, { x: 10.295, y: 5.633, w: 8.979, h: 4.893, rectRadius: 0.815, fill: { color: OLIVE_MID } });
  head(s, '103K', { x: 1.72, y: 6.306, w: 2.876, h: 1.166, fontSize: 80 });
  s.addShape('line', { x: 1.72, y: 8.324, w: 0.777, h: 0, line: { color: WHITE, width: 2.25 } });
  txt(s, PRAESENT_SHORT, { x: 1.72, y: 8.665, w: 3.212, h: 1.212, lineSpacingMultiple: 1 });
  statBlock(s, 11.3, { value: '89%', y: 6.345, w: 2.677, ruleDy: 2.229, label: 'Promotion Management' });
  navbar(s, BLACK);
}

// 9 — Feeding the Market, Sustaining the Planet
function slide09(pres) {
  const s = pres.addSlide();
  head(s, 'Feeding the Market, Sustaining the Planet', { x: 0.727, y: 2.144, w: 10.573, h: 1.605, color: BLACK });
  txt(s, 'Lorem ipsum dolor sit amet, conse lectus ornare, viverra ctetur adipiscing elit. Pellentesque scelerisque malesuada libero a pellentesque. Morbi orci dui, fermentum eget lectus ornare, viverr viverra ctetur adipiscing elit. ',
    { x: 11.891, y: 2.046, w: 7.384, h: 1.212, color: BLACK, lineSpacingMultiple: 1 });
  round(s, { x: 0.727, y: 4.5, w: 10.574, h: 5.026, rectRadius: 0.838, fill: { color: OLIVE_MID } });
  txt(s, 'Competitive Advantage', {
    x: 1.733, y: 4.925, w: 3.68, h: 1.091, fontFace: HEAD, fontSize: 28,
    lineSpacingMultiple: 1.2, charSpacing: -0.89
  });
  [['More comprehensive and integrated collaboration and data synchronization solution.', 5.016, 0.909],
   ['Competitive pricing and outstanding customer support.', 6.697, 0.625],
   ['Competitive pricing and outstanding customer support.', 8.266, 0.625]].forEach(([t, y, h]) => {
    txt(s, t, { x: 6.516, y, w: 3.842, h, lineSpacingMultiple: 1 });
  });
  tagStrip(s, OLIVE_MID, WHITE);
  navbar(s, BLACK);
}

// 10 — Harvesting Opportunities Today
function slide10(pres) {
  const s = pres.addSlide();
  s.background = { color: OLIVE };
  head(s, 'Harvesting Opportunities Today', { x: 0.696, y: 1.65, w: 12.198, h: 0.816 });
  txt(s, 'Integer aliquet, lectus sit amet dapibus cursus, erat leo dapibus ex, vel lacinia augue nisl vel arcu. Aenean vestibulum odio dapibus eros cursus, non egestas odio dictum. Nullam ut lectus maximus, tincidunt sem sit amet, imperdiet urna. Vivamus ac commodo leo. Nulla blandit lobortis tincidunt. Sed in diam sed ex pulvinar mattis accumsan a quam. Donec auctor tempor lorem.',
    { x: 11.891, y: 3.173, w: 7.384, h: 2.121, lineSpacingMultiple: 1 });
  subhead(s, 'Our Team', { x: 14.51, y: 7.068, w: 2.61, h: 0.505 });
  txt(s, 'Nunc euismod risus eu nibh bibendum, at condimentum . Vestibulum ante posuere dolor.',
    { x: 14.49, y: 7.863, w: 4.193, h: 0.909, lineSpacingMultiple: 1 });
  tagStrip(s, WHITE, BLACK);
  navbar(s, WHITE);
}

// 11 — Bridging Tradition with Technology
function slide11(pres) {
  const s = pres.addSlide();
  rect(s, { x: 0, y: 0, w: W, h: 7.2, fill: { color: OLIVE } });
  head(s, 'Bridging Tradition with Technology', { x: 0.726, y: 4.013, w: 5.337, h: 2.419 });
  [{ v: '94%', x: 7.106, lw: 3.189, lines: ['deaths worldwide due to', 'vehicle crashes every year'] },
   { v: '36,096', x: 11.891, lw: 3.185, lines: ['road deaths in', 'the U.S. in 2019'] },
   { v: '1.35 m', x: 16.083, lw: 3.19, lines: ['crashes involve', 'human error in the US'] }].forEach((c) => {
    head(s, c.v, { x: c.x, y: 2.097, w: 4.194, h: 0.816 });
    txt(s, c.lines.map((t, i) => ({ text: t, options: { breakLine: i === 0 } })),
      { x: c.x, y: 2.962, w: c.lw, h: 0.606, lineSpacingMultiple: 1 });
  });
  txt(s, '"The best way to predict the future is to create it." — Peter Drucker',
    { x: 0.727, y: 7.927, w: 4.195, h: 0.606, color: BLACK, align: 'center', lineSpacingMultiple: 1 });
  round(s, { x: 0.727, y: 9.446, w: 4.195, h: 0.676, rectRadius: 0.338, fill: { color: OLIVE_MID } });
  txt(s, 'View our Monthly Report', {
    x: 0.727, y: 9.446, w: 4.195, h: 0.676, align: 'center', valign: 'middle',
    fontFace: HEAD, fontSize: 18, charSpacing: 0
  });
  navbar(s, WHITE);
}

// 12 — Sustainable Growth through Smart Agriculture (2 x 2 grid on glass card)
function slide12(pres) {
  const s = pres.addSlide();
  s.background = { color: OLIVE };
  scrim(s, OLIVE);
  round(s, { x: 9.706, y: 4.353, w: 9.568, h: 6.173, rectRadius: 0.729, fill: { color: WHITE, transparency: 75 } });
  head(s, 'Sustainable Growth through Smart Agriculture', { x: 0.727, y: 2.42, w: 9.384, h: 2.419 });
  quote(s, HARVEST, { x: 0.727, y: 7.108, w: 7.523, h: 1.005 });
  [{ n: '01', t: 'Root Growth', x: 10.608, y: 4.822, tw: 3.577 },
   { n: '02', t: 'Agricultural', x: 14.814, y: 4.822, tw: 3.577 },
   { n: '03', t: 'Sustainable Farm', x: 10.608, y: 7.587, tw: 3.577 },
   { n: '04', t: 'Farming Method', x: 14.814, y: 7.587, tw: 3.702 }].forEach((c) => {
    const o = { fontFace: HEAD, bold: true, fontSize: 28, lineSpacingMultiple: 1, charSpacing: -0.69 };
    txt(s, c.n, Object.assign({ x: c.x, y: c.y, w: 1.005, h: 0.471 }, o));
    txt(s, c.t, Object.assign({ x: c.x, y: c.y + 0.734, w: c.tw, h: 0.471 }, o));
    txt(s, LOREM_GRID, { x: c.x, y: c.y + 1.34, w: 3.463, h: 1.212, lineSpacingMultiple: 1 });
  });
  txt(s, PRAESENT_PURUS, { x: 0.722, y: 8.817, w: 7.387, h: 1.212, lineSpacingMultiple: 1 });
  navbar(s, WHITE);
}

// 13 — Technology-Driven Farming for a New Era (bar chart drawn with rectangles)
function slide13(pres) {
  const s = pres.addSlide();
  s.background = { color: OLIVE };
  head(s, 'Technology-Driven Farming for a New Era',
    { x: 5.089, y: 1.883, w: 9.823, h: 1.618, align: 'center' });
  const bars = [
    { x: 0.728, y: 3.925, h: 6.032, c: OLIVE_MID, v: '95%', vx: 1.099, vy: 3.029, lx: 0.727, l: 'Client retention' },
    { x: 3.917, y: 6.259, h: 3.699, c: SAGE, v: '69%', vx: 4.286, vy: 5.374, lx: 3.913, l: 'Budget growth' },
    { x: 7.108, y: 5.515, h: 4.443, c: OLIVE_MID, v: '75%', vx: 7.478, vy: 4.630, lx: 7.108, l: 'ROI' },
    { x: 10.295, y: 4.977, h: 4.981, c: SAGE, v: '89%', vx: 10.662, vy: 4.092, lx: 10.290, l: 'Demands' },
    { x: 13.483, y: 5.284, h: 4.673, c: OLIVE_MID, v: '82%', vx: 13.853, vy: 4.421, lx: 13.472, l: 'Growth' },
    { x: 16.670, y: 3.925, h: 6.032, c: SAGE, v: '94%', vx: 17.041, vy: 3.029, lx: 16.655, l: 'Satisfaction' }
  ];
  bars.forEach((b) => {
    rect(s, { x: b.x, y: b.y, w: 2.599, h: b.h, fill: { color: b.c } });
    txt(s, b.v, {
      x: b.vx, y: b.vy, w: 1.858, h: 0.568, align: 'center', fontFace: HEAD,
      fontSize: 32, lineSpacingMultiple: 1.1, charSpacing: -0.89
    });
    subhead(s, b.l, { x: b.lx, y: 10.17, w: 2.602, h: 0.505, align: 'center' });
  });
  navbar(s, WHITE);
}

// 14 — Empowering Farmers through Innovation (two doughnut charts)
function slide14(pres) {
  const s = pres.addSlide();
  s.background = { color: OLIVE };
  rect(s, { x: 0, y: 9.051, w: W, h: 2.199, fill: { color: OLIVE_PALE } });
  head(s, 'Empowering Farmers through Innovation', { x: 0.724, y: 2.584, w: 7.385, h: 2.419 });
  txt(s, "At Procoach, we're a team of passionate professionals helping over 10,000 partners since 2015 unlock their potential, crush their goals, and elevate their performance.",
    { x: 0.71, y: 6.846, w: 7.79, h: 1.15, fontFace: HEAD, fontSize: 21, lineSpacingMultiple: 1.1, charSpacing: -0.89 });
  [{ title: 'Products', values: [43, 41], x: 9.676, ix: 11.973 },
   { title: 'Share', values: [2017, 2091], x: 14.13, ix: 16.443 }].forEach((c) => {
    s.addChart(pres.ChartType.doughnut,
      [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: c.values }], {
        x: c.x, y: 2.257, w: 5.211, h: 4.53,
        holeSize: 85, chartColors: [OLIVE_MID, SAGE_75], dataBorder: { pt: 0, color: OLIVE },
        showLegend: false, showValue: false,
        showTitle: true, title: c.title, titleColor: WHITE, titleFontFace: BODY, titleFontSize: 18
      });
    (c.title === 'Products' ? boxIcon : shareIcon)(s, c.ix, 4.564);
  });
  [11.668, 16.122].forEach((x, i) => {
    txt(s, '$1,999', { x, y: i ? 7.544 : 7.525, w: 1.227, h: 0.471, fontSize: 28, charSpacing: 0, wrap: false });
  });
  txt(s, 'Nam aliquam euismod dui et tincidunt. Nam vehicula mollis sem, sit amet tristique nibh pretium tempor. ', {
    x: 3.684, y: 9.466, w: 12.632, h: 0.808, align: 'center', fontFace: HEAD,
    fontSize: 21, color: BLACK, charSpacing: 0, margin: [7.2, 7.2, 3.6, 3.6]
  });
  navbar(s, WHITE);
}
/* white line icons sitting in the doughnut holes */
const ICON_LINE = { color: WHITE, width: 1.25 };
function boxIcon(s, x, y) {
  s.addShape('cube', { x: x + 0.06, y: y + 0.07, w: 0.48, h: 0.45, line: ICON_LINE });
}
function shareIcon(s, x, y) {
  const dots = [[0.40, 0.02], [0.02, 0.20], [0.40, 0.38]];
  s.addShape('line', { x: x + 0.10, y: y + 0.14, w: 0.34, h: 0.16, line: ICON_LINE, flipV: true });
  s.addShape('line', { x: x + 0.10, y: y + 0.30, w: 0.34, h: 0.16, line: ICON_LINE });
  dots.forEach(([dx, dy]) => {
    s.addShape('ellipse', { x: x + dx, y: y + dy, w: 0.18, h: 0.18, fill: { color: OLIVE }, line: ICON_LINE });
  });
}

// 15 — Where Agribusiness Meets Strategy
function slide15(pres) {
  const s = pres.addSlide();
  s.background = { color: OLIVE };
  head(s, 'Where Agribusiness Meets Strategy', { x: 0.727, y: 2.04, w: 18.547, h: 0.816 });
  txt(s, NUNC, { x: 15.078, y: 2.04, w: 4.196, h: 0.909, lineSpacingMultiple: 1 });
  round(s, { x: 4.922, y: 4.8, w: 3.395, h: 5.726, rectRadius: 0.716, fill: { color: MINT_PALE } });
  round(s, { x: 11.683, y: 4.8, w: 3.395, h: 5.726, rectRadius: 0.716, fill: { color: MINT_PALE } });
  quote(s, 'Agri-Projects That Make an Impact', { x: 5.296, y: 5.41, w: 2.813, h: 1.531, color: BLACK });
  txt(s, 'Nunc euismod eu nibh bibendum, at condimentum dolor interdum.',
    { x: 5.296, y: 8.72, w: 2.579, h: 1.212, color: BLACK, lineSpacingMultiple: 1 });
  quote(s, 'Agriculture is our wisest pursuit.', { x: 8.701, y: 4.758, w: 2.599, h: 1.414, lineSpacingMultiple: 1, charSpacing: -0.44 });
  imagePlaceholder(s, 12.196, 5.41, 1.585, 1.585, 0.15);   // sprout line-art icon
  quote(s, 'From Fields to Finance', { x: 12.209, y: 8.72, w: 2.588, h: 1.013, color: BLACK });
  navbar(s, WHITE);
}

// 16 — From Soil to Market: A Scalable Plan
function slide16(pres) {
  const s = pres.addSlide();
  s.background = { color: OLIVE };
  scrim(s, OLIVE);
  head(s, 'From Soil to Market: A Scalable Plan', { x: 0.727, y: 2.685, w: 5.424, h: 2.419 });
  txt(s, 'Curabitur suscipit tellus enim, id aliquet nisi maximus eu. Aliquam erat volutpat.z',
    { x: 0.727, y: 9.617, w: 4.194, h: 0.909, lineSpacingMultiple: 1 });
  round(s, { x: 8.701, y: 7.05, w: 10.584, h: 3.476, rectRadius: 0.41, fill: { color: WHITE, transparency: 80 } });
  quote(s, WISEST, { x: 9.433, y: 7.603, w: 9.217, h: 0.942, lineSpacingMultiple: 1, charSpacing: 0 });
  txt(s, PRAESENT_PURUS, { x: 9.433, y: 9.051, w: 9.161, h: 0.909, lineSpacingMultiple: 1 });
  round(s, { x: 8.701, y: 1.7, w: 5.788, h: 4.783, rectRadius: 0.33, fill: { color: WHITE, transparency: 80 } });
  round(s, { x: 15.078, y: 1.672, w: 4.196, h: 2.094, rectRadius: 0.247, fill: { color: WHITE, transparency: 80 } });
  quote(s, EVOLVES, { x: 15.606, y: 2.013, w: 3.141, h: 1.414, lineSpacingMultiple: 1, charSpacing: -0.44 });
  round(s, { x: 9.243, y: 1.99, w: 4.706, h: 2.872, rectRadius: 0.394, fill: { color: OLIVE } });
  txt(s, 'Project Clients', {
    x: 9.886, y: 2.492, w: 2.599, h: 0.337, fontFace: HEAD, bold: true, fontSize: 20, lineSpacingMultiple: 1
  });
  txt(s, PRAESENT_SHORT, { x: 9.886, y: 3.108, w: 3.418, h: 1.212, lineSpacingMultiple: 1 });
  // rating chip
  round(s, { x: 9.262, y: 4.856, w: 4.667, h: 1.317, rectRadius: 0.348, fill: { color: GREY_TX } });
  s.addShape('ellipse', { x: 9.504, y: 4.985, w: 1.125, h: 1.058, fill: { color: OLIVE_MID } });
  txt(s, '9.6', {
    x: 9.504, y: 4.985, w: 1.125, h: 1.058, align: 'center', valign: 'middle',
    fontFace: HEAD, bold: true, fontSize: 18, charSpacing: 0
  });
  txt(s, 'Our clients rate us', { x: 11.058, y: 5.234, w: 2.536, h: 0.313, lineSpacingMultiple: 1.1, charSpacing: -0.89 });
  txt(s, 'More than 10k reviews', { x: 11.058, y: 5.565, w: 2.536, h: 0.243, fontSize: 14, lineSpacingMultiple: 1.1, charSpacing: -0.89 });
  navbar(s, WHITE);
}

// 17 — Why we're here (radial gradient background)
function slide17(pres) {
  const s = pres.addSlide();
  s.background = { color: OLIVE };
  radial(s, [[0, OLIVE_LIGHT], [0.69, OLIVE_MID2], [1, OLIVE_DEEP]]);
  txt(s, 'Why we\u2019re here', {
    x: 0.727, y: 1.968, w: 2.906, h: 0.471, fontFace: HEAD, bold: true, fontSize: 28,
    lineSpacingMultiple: 1, charSpacing: -0.69
  });
  head(s, 'Farming Innovation for a Better Tomorrow', { x: 0.727, y: 2.818, w: 4.783, h: 3.22 });
  quote(s, WISEST, { x: 0.727, y: 8.641, w: 4.954, h: 1.885, lineSpacingMultiple: 1, charSpacing: -0.44 });
  [{ v: '$924', x: 15.082, y: 1.638, l: 'Growing Value with Every Season', lh: 0.303 },
   { v: '73%', x: 15.080, y: 4.012, l: 'Agri-Projects That Make an Impact', lh: 0.303 },
   { v: '94%', x: 15.080, y: 6.396, l: 'Feeding the Market, Sustaining the Planet', lh: 0.606 },
   { v: '1.35 m', x: 15.078, y: 8.771, l: 'Bridging Tradition with Technology', lh: 0.303 }].forEach((c) => {
    head(s, c.v, { x: c.x, y: c.y, w: 4.194, h: 0.816 });
    txt(s, c.l, { x: c.x, y: c.y + 0.866, w: 4.192, h: c.lh, lineSpacingMultiple: 1 });
  });
  navbar(s, WHITE, true);
}

// 18 — From Soil to Success: Powered by Innovation
function slide18(pres) {
  const s = pres.addSlide();
  s.background = { color: OLIVE };
  cornerGlow(s, [[0, OLIVE], [0.496, OLIVE_MID], [1, SAGE]]);
  head(s, 'From Soil to Success: Powered by Innovation', { x: 0.727, y: 1.657, w: 5.788, h: 3.22 });
  quote(s, EVOLVES, { x: 0.727, y: 5.631, w: 6.377, h: 1.005 });
  txt(s, 'Technology-Driven Farming', {
    x: 0.727, y: 8.858, w: 5.477, h: 0.471, fontFace: HEAD, bold: true, fontSize: 28,
    lineSpacingMultiple: 1, charSpacing: -0.52
  });
  txt(s, 'ullam tempor felis est, quis pellentesque diam tincidunt sed. Nam purus dui, hendrerit sed magna',
    { x: 0.728, y: 9.617, w: 4.782, h: 0.909, lineSpacingMultiple: 1 });
  navbar(s, WHITE);
}

// 19 — Contact, phone mockup
function slide19(pres) {
  const s = pres.addSlide();
  s.background = { color: OLIVE };
  txt(s, 'Farming Innovation for a Better Tomorrow', {
    x: 0.727, y: 1.453, w: 6.789, h: 3.067, fontFace: HEAD, bold: true, fontSize: 56,
    lineSpacingMultiple: 1.1, charSpacing: -0.89
  });
  [['Contact Us', 6.271, 6.928], ['What We Serve', 8.658, 9.314]].forEach(([label, ly, by]) => {
    subhead(s, label, { x: 0.727, y: ly, w: 4.194, h: 0.505 });
    txt(s, CURABITUR, { x: 0.727, y: by, w: 4.785, h: 1.212, lineSpacingMultiple: 1 });
  });
  // collage of photos on the right
  // phone mock-up
  round(s, { x: 8.178, y: 0.785, w: 4.715, h: 9.671, rectRadius: 0.72, fill: { color: 'B2B2B2' } });
  round(s, { x: 8.276, y: 0.888, w: 4.519, h: 9.465, rectRadius: 0.68, fill: { color: '191919' } });
  round(s, { x: 8.346, y: 0.972, w: 4.376, h: 9.306, rectRadius: 0.64, fill: { color: OLIVE } });
  [[2.789, 0.464], [3.590, 0.719], [4.504, 0.719]].forEach(([y, h]) => {
    rect(s, { x: 8.110, y, w: 0.068, h, fill: { color: SAGE } });   // side buttons
  });
  navbar(s, WHITE, true);
}

// 20 — Thank you
function slide20(pres) {
  const s = pres.addSlide();
  scrim(s, WHITE);
  head(s, 'Thank You For Your Attention.', {
    x: 3.917, y: 3.186, w: 12.167, h: 2.874, fontSize: 96, align: 'center',
    margin: [7.2, 7.2, 3.6, 3.6]
  });
  txt(s, 'Almost all of our participants report significant improvements.', {
    x: 6.81, y: 6.336, w: 6.379, h: 1.015, align: 'center', fontFace: HEAD, fontSize: 28,
    lineSpacingMultiple: 1.1
  });
  round(s, { x: 7.903, y: 8.557, w: 4.194, h: 0.676, rectRadius: 0.338, fill: { color: OLIVE } });
  txt(s, 'Latest Social Project', {
    x: 7.903, y: 8.557, w: 4.194, h: 0.676, align: 'center', valign: 'middle',
    fontFace: HEAD, fontSize: 20, charSpacing: 0
  });
  navbar(s, WHITE);
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */
function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'WIDE20', width: W, height: H });
  pres.layout = 'WIDE20';
  pres.title = 'Sustainable Growth Through Smart Agriculture';
  pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
   slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach((fn) => fn(pres));

  return pres.writeFile({
    fileName: path.join(__dirname, '16d1cd11-a501-4023-a34b-5e36a495962d_grok_final.pptx')
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
