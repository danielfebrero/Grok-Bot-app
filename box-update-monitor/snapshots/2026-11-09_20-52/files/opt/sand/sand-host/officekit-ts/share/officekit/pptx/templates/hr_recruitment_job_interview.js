/*
 * Carrera - Job Interview Presentation Template  (30 slides, 13.333 x 7.5 in)
 * Rebuilt from scratch with pptxgenjs. Raster images / SVG icon art from the
 * original deck are replaced by native shape placeholders.
 *
 *   node 16452d51-98ac-4c8f-8895-bd4132915e61_grok_final.js
 */
'use strict';

const pptxgen = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ theme */

const C = {
  navy: '0E13CC', // accent1
  purple: '3A0CA3', // accent2
  indigo: '3F37C9', // accent3
  blue: '4895EF', // accent4
  cyan: '4CC9F0', // accent5
  blueDk: '136DD6', // accent4 lumMod 75%
  cyanDk: '12ABDB', // accent5 lumMod 75%
  cyanLt: '94DFF6', // accent5 lumMod 60% / lumOff 40%  (the pale band colour)
  cyanPale: 'B7E9F9', // accent5 lumMod 40% / lumOff 60%
  indigoLt: '5B5FF4', // accent1 lumMod 60% / lumOff 40%
  violet: '7944F1', // accent2 lumMod 60% / lumOff 40%
  lilac: 'C8CAFB', // decorative connector / map fill
  black: '000000',
  white: 'FFFFFF',
  grey: 'D9D9D9',
  greyLt: 'F2F2F2',
  greyMid: 'A6A6A6',
  greyDk: 'AFABAB',
};

const HEAD = 'Poppins'; // major font
const BODY = 'Roboto'; // minor font

/* Lorem strings that repeat all over the deck */
const T = {
  s: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, ',
  intro:
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, veniam minim as tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco\u00a0',
  para:
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, ullamco\u00a0consectetur adipiscing elit. Integer vitae justo ullamcorper, scelerisque mi quis, ornare erat. Lorem ipsum dolor consectetur adipiscing elit. ',
  half: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, veniam minim as',
  tiny: 'Lorem ipsum dolor sit amet, consectetuer.',
  map: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer vitae justo ullamcorper, scelerisque mi. ',
};

/* Continent silhouettes for the two map slides, drawn as a coarse pixel grid
 * over the full 13.333 x 7.5 in canvas (64 columns x 30 rows). */
const WORLD = [
  '                                                                ',
  '                                                                ',
  '                      ###   ####                                ',
  '                    #############    ##           #             ',
  '                ## #### ########                   ##           ',
  '               ### ####   ######            #     ######  #     ',
  '        ###     ###  ###   #####      ###          #      ######',
  '        ###      # #    #  ###  #    ########      #############',
  '        ###      #### ###  ##       #########      #############',
  '         #      ##########        #  # #######     #######   #  ',
  '               ############        #######################      ',
  '               ##########          #######################      ',
  '               #########          ##  ##################        ',
  '               ########           ###    ##############         ',
  '                #######          ######################         ',
  '                  ##  #         ############# ########          ',
  '                   ###          ############   ##  ##           ',
  '                     #####       ##########                     ',
  '                      #####            ####        # ##         ',
  '                      #######           #            #   ##     ',
  '                      #######           ##              #       ',
  '                            #           # #            ####     ',
  '                                       ## #           ######    ',
  '                                      ##              ######    ',
  '                                                          #     ',
  '                                                                ',
  '                       #                                        ',
  '                       #                                        ',
];
const EUROPE = [
  '                                                            #   ',
  '                                              ###           #   ',
  '                                            ######             #',
  '                                           #########           #',
  '                                          ############ ## ######',
  '                                         #############  ########',
  '                        # ##             ##########  ###########',
  '                        ####            ##### ##### ############',
  '                         ##             ##### ##################',
  '                                      #####  ###################',
  '                                     ######  ###################',
  '                                     ####### ## ################',
  '                                     ######   ##################',
  '                                ##      ###  ###################',
  '                                ##     # #    ##################',
  '                               # #     #       #################',
  '                              ## ##  #####    ##################',
  '                              #  #   #####    ##################',
  '                                     ######   ############      ',
  '                                      ###################       ',
  '                                      ####################      ',
  '                                      ### ####### # #####       ',
  '                               ###      #  #####      ###       ',
  '                               #####     ## #### #### ####      ',
  '                              #####       # ## ###########      ',
  '                               ###       #   #  ########        ',
];

/* --------------------------------------------------------------- helpers */

/** preset-geometry adjust value (0-100000) expressed in inches, as pptxgenjs wants it */
const adj = (val, w, h) => (val / 1e5) * Math.min(w, h);

/** shape */
function sh(s, shape, x, y, w, h, o) {
  s.addShape(shape, Object.assign({ x, y, w, h }, o || {}));
}

/** text box - top anchored like every text frame in the source deck */
function tx(s, text, x, y, w, h, o) {
  s.addText(text, Object.assign(
    { x, y, w, h, valign: 'top', fontFace: BODY, fontSize: 18, color: C.black },
    o || {}
  ));
}

/** slide title (Poppins 44pt, 90% line spacing). `lines` may be a string or array */
function title(s, lines, x, y, w, h, o) {
  const arr = (Array.isArray(lines) ? lines : [lines]).map((t, i) => ({
    text: t,
    options: { breakLine: i < (Array.isArray(lines) ? lines.length : 1) - 1 },
  }));
  tx(s, arr, x, y, w, h, Object.assign(
    { fontFace: HEAD, fontSize: 44, lineSpacingMultiple: 0.9 }, o || {}
  ));
}

/** the blue disc with a white chevron used as a paragraph marker */
function arrowDisc(s, x, y, d, fill, arrowColor) {
  sh(s, 'ellipse', x, y, d, d, { fill: { color: fill || C.blueDk } });
  sh(s, 'chevron', x + d * 0.34, y + d * 0.28, d * 0.26, d * 0.44,
    { fill: { color: arrowColor || C.greyLt } });
}

/** outlined circle + arrow used on the "Search Jobs" / "Best employers" pills */
function arrowOutline(s, x, y, d, color) {
  sh(s, 'ellipse', x, y, d, d, { line: { color: color, width: 1.5 } });
  sh(s, 'rect', x + d * 0.24, y + d * 0.46, d * 0.36, d * 0.08, { fill: { color: color } });
  sh(s, 'triangle', x + d * 0.52, y + d * 0.35, d * 0.18, d * 0.3, { fill: { color: color }, rotate: 90 });
}

/** placeholder standing in for one of the deck's line-art SVG icons */
function iconGlyph(s, x, y, d, color) {
  sh(s, 'roundRect', x + d * 0.06, y + d * 0.12, d * 0.88, d * 0.7,
    { line: { color: color, width: Math.max(0.75, d * 2) }, rectRadius: d * 0.12 });
  sh(s, 'rect', x + d * 0.3, y + d * 0.86, d * 0.4, d * 0.06, { fill: { color: color } });
}

/** placeholder standing in for a photograph */
function photoBox(s, x, y, w, h, o) {
  o = o || {};
  sh(s, o.shape || 'rect', x, y, w, h,
    { fill: { color: o.color || C.greyMid }, rectRadius: o.radius });
  tx(s, '[image]', x, y + h / 2 - 0.2, w, 0.4,
    { align: 'center', fontSize: 12, color: o.label || C.white });
}

/** bold 16pt heading + 12pt body - the deck's ubiquitous "feature" pair */
function feature(s, o) {
  tx(s, o.title, o.x, o.y, o.tw || 2.727, 0.37,
    { fontSize: 16, bold: true, color: o.titleColor || C.black, align: o.align });
  tx(s, o.body, o.x, o.y + (o.gap || 0.337), o.w || 2.637, o.bh || 0.602,
    { fontSize: 12, lineSpacingMultiple: 1.3, align: o.align });
}

/** feature preceded by a numbered disc */
function numberedFeature(s, o) {
  sh(s, 'ellipse', o.x, o.y + 0.058, 0.608, 0.608, { fill: { color: o.discColor || C.blueDk } });
  tx(s, o.num, o.x + 0.053, o.y + 0.151, 0.502, 0.404,
    { align: 'center', bold: true, fontSize: 18, color: C.white });
  feature(s, { x: o.x + 0.763, y: o.y, tw: 3.093, w: o.w || 3.859, title: o.title, body: o.body });
}

/** icon disc + feature text */
function iconFeature(s, o) {
  sh(s, 'ellipse', o.x, o.y + 0.03, 0.608, 0.608, { fill: { color: C.blueDk } });
  iconGlyph(s, o.x + 0.129, o.y + 0.168, 0.332, C.white);
  feature(s, { x: o.x + (o.dx || 0.647), y: o.y, w: o.w, bh: o.bh, title: o.title, body: o.body });
}

/** the "Carrera" handshake mark + wordmark that sits in every top-right corner */
function logo(s, color) {
  sh(s, 'chevron', 11.22, 0.50, 0.30, 0.28, { fill: { color: color } });
  sh(s, 'chevron', 11.40, 0.50, 0.30, 0.28, { fill: { color: color } });
  tx(s, 'Carrera', 11.734, 0.444, 1.209, 0.388, { fontFace: HEAD, fontSize: 18, color: color });
}

/** paint one of the pixel-grid maps above as pale land tiles */
function landmass(s, grid) {
  const cw = 13.333 / 64, ch = 7.5 / 30;
  grid.forEach(function (row, r) {
    for (let c = 0; c < row.length; c++) {
      if (row[c] !== '#') continue;
      let run = 1;
      while (row[c + run] === '#') run++; // merge horizontal runs into one shape
      sh(s, 'roundRect', c * cw, r * ch, run * cw + 0.02, ch + 0.02,
        { fill: { color: C.lilac }, rectRadius: 0.05 });
      c += run - 1;
    }
  });
}

/** thin progress bar (track + value) */
function progress(s, x, y, w, pct, color) {
  sh(s, 'line', x, y, w, 0, { line: { color: 'E7E6E6', width: 8 } });
  sh(s, 'line', x, y, w * pct, 0, { line: { color: color || C.navy, width: 8 } });
}

/* --------------------------------------------------------- deck assembly */

function build() {
  const pres = new pptxgen();
  pres.defineLayout({ name: 'W16x9', width: 13.333, height: 7.5 });
  pres.layout = 'W16x9';
  pres.author = 'Carrera';
  pres.title = 'Job Interview Presentation Template';

  // Master carries the page-number chip; the logo is stamped per slide so it can
  // switch colour on the dark cover / contact slides.
  pres.defineSlideMaster({
    title: 'CARRERA',
    objects: [{ rect: { x: 12.315, y: 6.902, w: 0.628, h: 0.598, fill: { color: C.cyan } } }],
    slideNumber: { x: 12.353, y: 6.999, w: 0.551, h: 0.404, align: 'center', fontFace: BODY, fontSize: 18, color: C.white },
  });

  const pendingLogos = [];
  function add(logoColor) {
    const s = pres.addSlide({ masterName: 'CARRERA' });
    pendingLogos.push([s, logoColor || C.cyan]);
    return s;
  }

  /* -- 1 - cover ---------------------------------------------------- */
  (function () {
    const s = add(C.white);
    sh(s, 'rect', 0.01, 0, 13.333, 7.5, { fill: { color: C.blueDk, transparency: 40 } });
    title(s, 'Carrera', 3.574, 2.239, 6.009, 1.587, { fontSize: 100, color: C.white, align: 'center' });
    tx(s, 'Job Interview Presentation Template', 3.048, 4.0, 7.236, 0.565,
      { fontFace: HEAD, fontSize: 24, color: C.white, align: 'center', lineSpacingMultiple: 0.9 });
    sh(s, 'roundRect', 3.222, 4.565, 6.91, 0.565, { fill: { color: C.white }, rectRadius: 0.094 });
    tx(s, 'Search Jobs', 3.484, 4.674, 2.791, 0.467, { fontFace: HEAD, fontSize: 20 });
    arrowOutline(s, 9.583, 4.641, 0.363, C.black);
  })();

  /* -- 2 - we are Carrera ------------------------------------------- */
  (function () {
    const s = add();
    sh(s, 'rect', 0, 0, 5.655, 3.295, { fill: { color: C.cyanLt } });
    title(s, 'We are Carrera', 7.734, 0.968, 3.856, 1.531);
    feature(s, { x: 0.812, y: 5.154, w: 5.854, bh: 0.865, gap: 0.328, title: 'Innovative Recruiter', body: T.intro });
    [['01', 2.974], ['02', 4.218], ['03', 5.463]].forEach(function (r) {
      numberedFeature(s, { x: 7.868, y: r[1], num: r[0], title: 'Lorem ipsum', body: T.s });
    });
  })();

  /* -- 3 - discover the best employers ------------------------------ */
  (function () {
    const s = add();
    sh(s, 'round1Rect', 6.423, 0, 6.91, 4.393,
      { fill: { color: C.cyanLt }, rectRadius: adj(16667, 6.91, 4.393), flipH: true, flipV: true });
    title(s, ['Discover ', 'the Best Employers'], 0.812, 0.958, 4.67, 2.648);
    arrowDisc(s, 0.931, 4.135, 0.556);
    feature(s, { x: 1.708, y: 4.065, w: 4.458, bh: 1.128, gap: 0.328, title: 'Innovative Recruiter', body: T.intro });
    sh(s, 'roundRect', 0.981, 6.145, 6.91, 0.565, { fill: { color: C.cyanPale }, rectRadius: 0.094 });
    tx(s, 'Best employers', 1.209, 6.283, 2.481, 0.467, { fontFace: HEAD, fontSize: 16 });
    arrowOutline(s, 7.342, 6.221, 0.363, C.black);
  })();

  /* -- 4 - discover your dream jobs --------------------------------- */
  (function () {
    const s = add();
    sh(s, 'round1Rect', 1.767, 4.862, 11.566, 2.634,
      { fill: { color: C.cyanLt }, rectRadius: adj(35651, 11.566, 2.634), flipH: true });
    title(s, 'Discover Your Dream Jobs', 0.812, 1.251, 5.502, 1.538);
    arrowDisc(s, 0.931, 3.135, 0.556);
    feature(s, { x: 1.708, y: 3.065, w: 4.958, bh: 0.865, gap: 0.328, title: 'Innovative Recruiter', body: T.intro });
    s.addText([
      { text: '87', options: { fontSize: 166, bold: true, charSpacing: -1.5, color: C.navy } },
      { text: '%', options: { fontSize: 138, bold: true, charSpacing: -1.5, baseline: 800, color: C.navy } },
    ], { x: 8.629, y: 1.043, w: 3.825, h: 2.895, valign: 'top', align: 'center', fontFace: BODY });
    tx(s, 'Get Recruited', 8.966, 3.544, 2.727, 0.572, { fontSize: 28, bold: true, color: C.navy });
    [[2.664, 0.807], [6.231, 0.775], [9.766, 0.775]].forEach(function (i) {
      iconGlyph(s, i[0], 5.186, i[1], C.navy);
    });
    [[2.521, 6.163], [5.971, 6.183], [9.766, 6.144]].forEach(function (p) {
      feature(s, { x: p[0], y: p[1], title: 'Lorem ipsum', body: T.s });
    });
  })();

  /* -- 5 - effective strategies for hiring -------------------------- */
  (function () {
    const s = add();
    sh(s, 'round1Rect', 0, 3.75, 8.033, 3.75,
      { fill: { color: C.cyanLt }, rectRadius: adj(35651, 8.033, 3.75) });
    title(s, ['Effective Strategies ', 'for Hiring'], 0.835, 0.967, 5.961, 2.214);
    [[0.957, 4.421], [0.957, 5.801], [4.466, 4.451], [4.466, 5.83]].forEach(function (p) {
      iconFeature(s, { x: p[0], y: p[1], title: 'Lorem ipsum', body: T.s });
    });
    tx(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, veniam minim as tempor incididunt ut labore et dolore magna ',
      8.375, 5.705, 3.827, 1.128, { fontSize: 12, lineSpacingMultiple: 1.3 });
  })();

  /* -- 6 - professional team ---------------------------------------- */
  (function () {
    const s = add();
    sh(s, 'round2SameRect', 0, 4.866, 13.333, 2.634, { fill: { color: C.cyanLt }, rectRadius: 0.94 });
    title(s, 'Professional Team', 2.218, 0.815, 8.898, 0.791, { align: 'center' });
    const people = [
      { x: 2.478, fill: C.blue, name: 'Victoria Angelis', role: 'Consultan', a: '200', b: '73' },
      { x: 6.953, fill: C.blueDk, name: 'Ethaneskin', role: 'Content Planer', a: '100', b: '65' },
    ];
    people.forEach(function (p) {
      sh(s, 'roundRect', p.x, 2.145, 3.936, 4.764, { fill: { color: p.fill }, rectRadius: adj(12922, 3.936, 4.764) });
      tx(s, p.name, p.x + 0.319, 4.461, 3.336, 0.505, { fontSize: 24, color: C.white, align: 'center' });
      tx(s, p.role, p.x + 0.665, 4.891, 2.701, 0.37, { fontSize: 16, color: C.white, align: 'center' });
      sh(s, 'line', p.x + 2.053, 5.492, 0, 0.949, { line: { color: C.grey, width: 0.5 } });
      tx(s, p.a, p.x + 0.348, 5.368, 1.342, 0.572, { fontSize: 28, color: C.white, align: 'center' });
      tx(s, 'Feedback Given', p.x + 0.348, 5.869, 1.342, 0.572, { fontSize: 14, color: C.white, align: 'center' });
      tx(s, p.b, p.x + 2.246, 5.368, 1.342, 0.572, { fontSize: 28, color: C.white, align: 'center' });
      tx(s, 'Hired', p.x + 2.246, 5.869, 1.342, 0.337, { fontSize: 14, color: C.white, align: 'center' });
    });
    sh(s, 'ellipse', 2.19, 1.944, 1.15, 1.15, { fill: { color: C.navy } });
    iconGlyph(s, 2.569, 2.25, 0.423, C.white);
  })();

  /* -- 7 - expert for hiring the best talents (doughnut) ------------- */
  (function () {
    const s = add();
    sh(s, 'round1Rect', 0, 4.12, 13.333, 3.38, { fill: { color: C.cyanLt }, rectRadius: adj(16667, 13.333, 3.38) });
    title(s, 'Expert for Hiring The Best Talents', 0.88, 0.963, 5.787, 1.585);
    s.addChart(pres.ChartType.doughnut,
      [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: [5.2, 3.2] }],
      {
        x: 7.231, y: 1.884, w: 2.957, h: 2.1,
        chartColors: [C.navy, C.grey], holeSize: 65, showLegend: false, showValue: false,
        chartArea: { fill: { color: C.white, transparency: 100 } },
        layout: { x: 0.249, y: 0, w: 0.584, h: 0.876 },
      });
    tx(s, '75%', 8.185, 2.731, 1.3, 0.55,
      { fontFace: HEAD, fontSize: 40, bold: true, superscript: true, color: C.navy, align: 'center' });
    tx(s, 'Chart Data', 9.86, 2.329, 2.015, 0.438, { fontFace: HEAD, fontSize: 20, bold: true, color: C.navy });
    tx(s, 'Lorem ipsum dolor sit amet, consectetur', 9.86, 2.677, 2.015, 0.602, { fontSize: 12, lineSpacingMultiple: 1.3 });
    arrowDisc(s, 7.085, 4.656, 0.556);
    feature(s, {
      x: 7.863, y: 4.585, w: 4.252, bh: 0.865, gap: 0.329, title: 'Innovative Recruiter',
      body: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, veniam minim as tempor incididunt exercitation ullamco. veniam minim as tempor incididunt ullamco\u00a0',
    });
  })();

  /* -- 8 - expert strategies for hiring ----------------------------- */
  (function () {
    const s = add();
    sh(s, 'roundRect', 4.551, 2.644, 9.3, 3.893, { fill: { color: C.cyanLt }, rectRadius: 0.6 });
    title(s, ['Expert Strategies ', 'for Hiring'], 0.88, 0.963, 8.608, 1.585);
    arrowDisc(s, 0.932, 2.834, 0.556);
    feature(s, {
      x: 1.709, y: 2.764, w: 2.368, bh: 1.128, gap: 0.328, title: 'Innovative Recruiter',
      body: 'Lorem ipsum dolor amet, consectetur adipiscing elit, veniam minim as tempor incididunt ut exercitation',
    });
    tx(s, '270K', 1.705, 5.014, 2.372, 1.111, { fontFace: HEAD, fontSize: 60, bold: true, color: C.navy });
    tx(s, 'Best Employers', 1.705, 5.872, 2.532, 0.37, { fontFace: HEAD, fontSize: 16, bold: true, color: C.navy });
    [[5.697, 3.373], [5.697, 4.752], [9.206, 3.402], [9.206, 4.782]].forEach(function (p) {
      iconFeature(s, { x: p[0], y: p[1], title: 'Lorem ipsum', body: T.s });
    });
  })();

  /* -- 9 - find your career with us --------------------------------- */
  (function () {
    const s = add();
    title(s, 'Find Your Career With Us', 6.667, 1.522, 6.243, 1.585);
    arrowDisc(s, 6.795, 3.177, 0.556);
    feature(s, { x: 7.572, y: 3.107, w: 4.706, gap: 0.328, title: 'Innovative Recruiter', body: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, veniam minim as tempor incididunt ut exercitation ullamco\u00a0' });
    sh(s, 'round2SameRect', 7.609, 0.62, 1.895, 9.554, { fill: { color: C.cyanLt }, rotate: 270, rectRadius: 0.316 });
    tx(s, '365', 4.431, 4.732, 2.681, 1.111, { fontFace: HEAD, fontSize: 60, bold: true, color: C.navy });
    tx(s, 'Best Companies', 4.431, 5.59, 2.532, 0.37, { fontFace: HEAD, fontSize: 16, bold: true, color: C.navy });
    tx(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, veniam minim as tempor incididunt ut exercitation ullamco\u00a0',
      6.795, 4.947, 5.012, 0.602, { fontSize: 12, lineSpacingMultiple: 1.3 });
  })();

  /* -- 10 - job fair interview -------------------------------------- */
  (function () {
    const s = add();
    sh(s, 'round1Rect', 0, 0, 13.333, 2.326,
      { fill: { color: C.cyanLt }, rectRadius: adj(34469, 13.333, 2.326), flipH: true, flipV: true });
    title(s, 'Job Fair Interview', 0.835, 1.27, 5.831, 0.884);
    [[0.918, 2.85, 4.379], [0.918, 4.858, 4.379], [7.298, 4.858, 4.25]].forEach(function (p) {
      iconFeature(s, { x: p[0], y: p[1], w: p[2], dx: 0.738, bh: 1.39, title: 'Lorem ipsum', body: T.para });
    });
  })();

  /* -- 11 - best company and great management ----------------------- */
  (function () {
    const s = add();
    title(s, 'Best Company and Great Management', 0.835, 0.967, 5.562, 2.239);
    arrowDisc(s, 0.932, 3.82, 0.556);
    feature(s, {
      x: 1.709, y: 3.75, w: 4.39, bh: 1.653, gap: 0.328, title: 'Innovative Recruiter',
      body: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, veniam minim as tempor incididunt ullamco\u00a0consectetur adipiscing elit. Integer vitae justo ullamcorpe sceleruss scelerisque mi quis, ornare erat. Lorem ipsum dolor consectetur adipiscing elit. ',
    });
    iconGlyph(s, 3.952, 5.75, 0.332, C.greyMid);
    [['01', -0.095, 1.795], ['02', 1.599, 3.489], ['03', 3.292, 5.182]].forEach(function (r) {
      sh(s, 'round2SameRect', 10.484, r[1], 1.242, 4.456, { fill: { color: C.cyanLt }, rotate: 270, rectRadius: 0.207 });
      numberedFeature(s, { x: 9.213, y: r[2], num: r[0], w: 2.959, title: 'Lorem ipsum', body: 'Lorem ipsum dolor sit amet, ' });
    });
  })();

  /* -- 12 - expert team and great management ------------------------ */
  (function () {
    const s = add();
    sh(s, 'round1Rect', 0, 0, 13.333, 3.75,
      { fill: { color: C.cyanLt }, rectRadius: adj(16667, 13.333, 3.75), flipH: true, flipV: true });
    title(s, 'Expert Team and Great Management', 0.808, 1.065, 5.248, 2.148);
    arrowDisc(s, 0.898, 4.514, 0.556);
    feature(s, {
      x: 1.676, y: 4.444, tw: 3.519, w: 4.226, bh: 1.39, gap: 0.328, title: 'Best Team and Management',
      body: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, veniam minim as tempor incididunt ut exercitation ullamco\u00a0consectetur adipiscing elit. Integer vitae justo ullamcorper, scelerisque mi quis, ornare erat. Lorem ipsum dolor consectetur',
    });
    tx(s, 'Team Growth', 6.667, 5.615, 2.816, 0.37, { fontFace: HEAD, fontSize: 16, bold: true, color: C.navy });
    tx(s, '85%', 10.744, 5.615, 1.296, 0.37, { fontFace: HEAD, fontSize: 16, bold: true, color: C.navy, align: 'right' });
    progress(s, 6.667, 6.112, 5.373, 0.821);
  })();

  /* -- 13 - boost your opportunity ---------------------------------- */
  (function () {
    const s = add();
    title(s, 'Boost Your Opportunity', 0.835, 0.967, 4.455, 1.641);
    arrowDisc(s, 0.932, 2.867, 0.556);
    feature(s, {
      x: 1.709, y: 2.796, w: 3.357, bh: 0.865, gap: 0.328, title: 'Innovative Recruiter',
      body: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, veniam minim as tempor incididunt ut',
    });
    sh(s, 'round2SameRect', 1.424, 3.082, 1.899, 4.747,
      { fill: { color: C.cyanLt }, rotate: 270, flipV: true, rectRadius: 0.317 });
    tx(s, 'Feedbacks', 0.86, 4.963, 1.696, 0.404, { fontFace: HEAD, fontSize: 18, bold: true, color: C.navy });
    tx(s, '47M', 0.932, 5.281, 1.592, 0.774, { fontFace: HEAD, fontSize: 40, bold: true, color: C.navy });
    tx(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. ', 2.595, 5.019, 2.01, 1.128,
      { fontSize: 12, color: C.navy, lineSpacingMultiple: 1.3 });
    iconGlyph(s, 3.952, 4.796, 0.332, C.greyMid);
    // phone mock-up (photo placeholder in the layout of the original deck)
    // phone mock-up: dark body running off the bottom edge, screen replaced by
    // a banded colour placeholder sampled from the original wallpaper
    sh(s, 'roundRect', 5.41, 0.93, 3.56, 7.0, { fill: { color: '1D1D1F' }, rectRadius: 0.42 });
    sh(s, 'roundRect', 5.626, 1.133, 3.13, 6.6, { fill: { color: 'D1859B' }, rectRadius: 0.3 });
    ['95326D', 'BA4A8C', 'DB819B', 'E49F9D', 'E8B7AB', 'EECAC1'].forEach(function (band, i) {
      sh(s, 'rect', 5.68, 1.33 + i * 1.03, 3.02, 1.03, { fill: { color: band } });
    });
    tx(s, '[image]', 5.626, 3.9, 3.13, 0.4, { align: 'center', fontSize: 12, color: C.white });
    [[10.692, 2.519], [10.683, 4.818]].forEach(function (p, i) {
      sh(s, 'ellipse', p[0] + 0.009, p[1] - 0.829, 0.608, 0.608, { fill: { color: C.blueDk } });
      iconGlyph(s, p[0] + 0.138, p[1] - 0.691, 0.332, C.white);
      feature(s, { x: p[0] - 1.005, y: p[1], title: 'Lorem ipsum', body: T.s, align: 'center' });
    });
  })();

  /* -- 14 - the best recruiter -------------------------------------- */
  (function () {
    const s = add();
    sh(s, 'round2SameRect', -0.417, 0.417, 7.5, 6.667,
      { fill: { color: C.cyanLt }, rotate: 270, flipV: true, rectRadius: 1.111 });
    title(s, 'The Best Recruiter', 8.063, 1.879, 4.455, 1.641);
    arrowDisc(s, 8.16, 3.613, 0.556);
    feature(s, {
      x: 8.937, y: 3.542, w: 3.483, bh: 0.865, gap: 0.328, title: 'Innovative Recruiter',
      body: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, veniam minim as tempor incididunt ut',
    });
    tx(s, '10+ Years Experiences', 8.16, 5.342, 2.727, 0.37, { fontFace: HEAD, fontSize: 16, bold: true, color: C.navy });
    progress(s, 8.16, 5.84, 4.261, 0.821);
  })();

  /* -- 15 - connecting job seekers and companies -------------------- */
  (function () {
    const s = add();
    sh(s, 'round1Rect', -0.083, 4.067, 4.231, 3.528,
      { fill: { color: C.cyanLt }, rectRadius: adj(35651, 4.231, 3.528) });
    title(s, ['Connecting ', 'Job Seekers ', 'and Companies'], 6.583, 1.176, 7.018, 2.268);
    feature(s, { x: 6.583, y: 3.66, tw: 3.513, w: 6.109, bh: 1.128, title: 'Lorem ipsum', body: T.para });
    [['578k', 'Feedback Given', 6.583], ['340k', 'Employers', 8.562], ['2345', 'Company', 10.54]].forEach(function (r) {
      tx(s, r[0], r[2], 5.231, 1.434, 0.572, { fontFace: HEAD, fontSize: 28, bold: true, color: C.navy });
      tx(s, r[1], r[2], 5.834, 2.354, 0.337, { fontFace: HEAD, fontSize: 14, color: C.navy });
    });
  })();

  /* -- 16 - job hiring goals ---------------------------------------- */
  (function () {
    const s = add();
    [[1.258, C.cyanDk, '367+'], [3.174, C.blueDk, '476+'], [5.091, C.blue, '476+']].forEach(function (r) {
      sh(s, 'roundRect', 6.381, r[0], 4.701, 1.278, { fill: { color: r[1] }, rectRadius: adj(26701, 4.701, 1.278) });
      iconGlyph(s, 6.95, r[0] + 0.31, 0.66, C.white);
      tx(s, r[2], 8.413, r[0] + 0.162, 1.248, 0.438, { fontFace: HEAD, fontSize: 20, bold: true, color: C.white });
      tx(s, 'Manage Money', 8.413, r[0] + 0.625, 1.569, 0.337, { fontSize: 14, color: C.white });
    });
    title(s, 'Job Hiring Goals', 0.891, 2.198, 4.185, 1.615);
    arrowDisc(s, 0.895, 4.603, 0.556);
    feature(s, {
      x: 1.673, y: 4.532, w: 3.958, bh: 1.39, gap: 0.328, title: 'Innovative Recruiter',
      body: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, veniam minim as tempor incididunt ut exercitation ullamco\u00a0consectetur adipiscing elit. Integer vitae justo ullamcorper, scelerisque mi quis, ornare erat. Lorem ipsum dolor consectetur',
    });
    iconGlyph(s, 3.916, 6.532, 0.332, C.greyMid);
  })();

  /* -- 17 - table pricing ------------------------------------------- */
  (function () {
    const s = add();
    const plans = [
      { x: 1.515, fill: C.cyan, name: 'Agency', price: '$75' },
      { x: 5.091, fill: C.cyanDk, name: 'Premium', price: '$85' },
      { x: 8.715, fill: C.blueDk, name: 'Ultimate', price: '$95' },
    ];
    plans.forEach(function (p) {
      const dx = p.x - 1.515;
      sh(s, 'roundRect', p.x, 1.975, 3.185, 4.554, { fill: { color: p.fill }, rectRadius: adj(13162, 3.185, 4.554) });
      tx(s, p.name, 1.847 + dx, 2.292, 2.466, 0.707, { fontFace: HEAD, fontSize: 36, color: C.white, align: 'center' });
      sh(s, 'line', 2.523 + dx, 3.094, 1.864, 0, { line: { color: C.grey, width: 2.5 } });
      sh(s, 'line', 1.772 + dx, 3.094, 1.322, 0, { line: { color: C.navy, width: 2.5 } });
      s.addText([
        { text: p.price, options: { fontSize: 54, bold: true } },
        { text: '/', options: { fontSize: 16, bold: true, charSpacing: 3 } },
        { text: 'mo', options: { fontSize: 16, charSpacing: 2 } },
      ], { x: 1.886 + dx, y: 3.272, w: 2.427, h: 1.01, valign: 'top', align: 'center', fontFace: HEAD, color: C.white });
      [4.455, 4.988, 5.521].forEach(function (y) {
        sh(s, 'ellipse', 2.343 + dx, y + 0.037, 0.238, 0.238, { fill: { color: C.white, transparency: 55 } });
        sh(s, 'ellipse', 2.382 + dx, y + 0.077, 0.159, 0.159, { fill: { color: C.white } });
        tx(s, 'Lorem Ipsum', 2.638 + dx, y, 2.674, 0.337, { fontSize: 14, color: C.white });
      });
      sh(s, 'roundRect', 2.037 + dx, 6.256, 2.14, 0.535, { fill: { color: C.navy }, rectRadius: 0.2675 });
      tx(s, 'Try Now', 2.125 + dx, 6.322, 1.93, 0.404, { color: C.white, align: 'center' });
    });
    title(s, 'Table Pricing', 3.767, 0.836, 5.831, 0.859, { align: 'center' });
  })();

  /* -- 18 - break slide --------------------------------------------- */
  (function () {
    const s = add();
    s.background = { color: C.cyanLt };
    tx(s, [{ text: 'BREAK', options: { breakLine: true } }, { text: 'SLIDE' }],
      0.743, 2.62, 5.924, 3.063, { fontFace: HEAD, fontSize: 88, bold: true, color: C.navy });
    arrowDisc(s, 0.918, 5.683, 0.556);
    tx(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, veniam minim as tempor incididunt ut',
      1.577, 5.683, 4.292, 0.602, { fontSize: 12, color: C.navy, lineSpacingMultiple: 1.3 });
  })();

  /* -- 19 - our agenda this month ----------------------------------- */
  (function () {
    const s = add();
    sh(s, 'round1Rect', 0, 5.037, 13.333, 2.463, { fill: { color: C.cyanLt }, rectRadius: adj(16667, 13.333, 2.463) });
    const cards = [
      { x: 2.496, y: 3.401, h: 2.935, r: 13351, fill: C.cyan, date: '16 August', cx: 3.283, cy: 3.833 },
      { x: 5.096, y: 3.427, h: 2.935, r: 12635, fill: C.blueDk, date: '20 August', cx: 5.871, cy: 3.847 },
      { x: 7.697, y: 2.521, h: 3.815, r: 10483, fill: C.purple, date: '22 August', cx: 8.487, cy: 2.933, tall: true },
      { x: 10.232, y: 3.401, h: 2.962, r: 10287, fill: C.indigoLt, date: '26 August', cx: 11.021, cy: 3.823 },
    ];
    cards.forEach(function (c) {
      sh(s, 'roundRect', c.x, c.y, 2.25, c.h, { fill: { color: c.fill }, rectRadius: adj(c.r, 2.25, c.h) });
      tx(s, '2023', c.x + 1.408, c.y + 0.196, 0.883, 0.316, { fontSize: 14, color: C.white, align: 'center' });
      sh(s, 'ellipse', c.cx, c.cy, 0.671, 0.671, { fill: { color: C.white } });
      iconGlyph(s, c.cx + 0.18, c.cy + 0.18, 0.312, c.fill);
      if (c.tall) {
        tx(s, c.date, c.x + 0.177, 3.722, 1.896, 0.463, { fontSize: 14, color: C.white, align: 'center' });
        tx(s, 'Agenda List', c.x + 0.177, 4.141, 1.896, 0.382, { fontSize: 14, color: C.white, align: 'center' });
        ['Name Consultant', 'Idea Evaluation', 'Analysis Team'].forEach(function (t, i) {
          tx(s, t, c.x + 0.13, 4.717 + i * 0.388, 2.073, 0.303, { fontSize: 14, color: C.white, align: 'center' });
        });
      } else {
        tx(s, c.date, c.x + 0.177, c.y + 1.14, 1.896, 0.411, { fontSize: 14, color: C.white, align: 'center' });
        tx(s, 'Agenda List', c.x + 0.177, c.y + 1.505, 1.896, 0.334, { fontSize: 14, color: C.white, align: 'center' });
        tx(s, 'Lorem ipsum dolor sit amet', c.x + 0.177, c.y + 1.91, 1.896, 0.596, { fontSize: 14, color: C.white, align: 'center' });
      }
    });
    title(s, 'Our Agenda This Month', 0.843, 0.832, 4.067, 1.379);
    tx(s, T.half, 0.9, 2.38, 3.59, 0.602, { fontSize: 12, lineSpacingMultiple: 1.3 });
  })();

  /* -- 20 - recruitment flow ---------------------------------------- */
  (function () {
    const s = add();
    title(s, 'Recruitment Flow', 2.295, 1.002, 8.744, 0.774, { align: 'center', color: C.black });
    // the snaking pipe: horizontal runs, diagonal risers and rounded U-turns
    const W = 20;
    sh(s, 'line', -0.164, 2.378, 7.584, 0, { line: { color: C.navy, width: W } });
    sh(s, 'rightBracket', 7.42, 2.378, 0.539, 0.949,
      { line: { color: '240A82', width: W }, rectRadius: adj(66081, 0.539, 0.949) });
    sh(s, 'line', 5.916, 3.327, 1.504, 0, { line: { color: C.purple, width: W } });
    sh(s, 'rightBracket', 5.377, 3.329, 0.539, 0.949,
      { line: { color: '3A22C6', width: W }, rotate: 180, rectRadius: adj(63506, 0.539, 0.949) });
    sh(s, 'line', 5.916, 4.276, 1.509, 0, { line: { color: C.indigo, width: W } });
    sh(s, 'rightBracket', 7.425, 4.276, 0.539, 0.949,
      { line: { color: '4466DC', width: W }, rectRadius: adj(62219, 0.539, 0.949) });
    sh(s, 'line', 5.884, 5.225, 1.541, 0, { line: { color: C.blue, width: W } });
    sh(s, 'rightBracket', 5.369, 5.225, 0.515, 0.949,
      { line: { color: '4AAFF0', width: W }, rotate: 180, rectRadius: adj(69596, 0.515, 0.949) });
    sh(s, 'line', 5.916, 6.174, 2.351, 0, { line: { color: C.cyan, width: W } });
    // diagonals connecting each U-turn to the icon disc above it
    sh(s, 'line', 5.812, 3.803, 1.745, -0.948, { line: { color: C.purple, width: W } });
    sh(s, 'line', 5.812, 5.699, 1.745, -0.948, { line: { color: C.blue, width: W } });
    [[7.221, 2.516], [5.476, 3.467], [7.221, 4.415], [5.476, 5.363]].forEach(function (p) {
      sh(s, 'ellipse', p[0], p[1], 0.671, 0.671, { fill: { color: C.white } });
      iconGlyph(s, p[0] + 0.18, p[1] + 0.18, 0.312, C.navy);
    });
    const stats = [
      { v: '3,328', c: C.navy, x: 8.495, y: 2.311, al: 'left' },
      { v: '720', c: C.indigo, x: 8.495, y: 4.174, al: 'left' },
      { v: '400+', c: C.purple, x: 2.54, y: 3.186, al: 'right' },
      { v: '1,500', c: C.blue, x: 2.54, y: 5.088, al: 'right' },
    ];
    stats.forEach(function (st) {
      tx(s, st.v, st.x, st.y, 2.309, 0.505, { fontSize: 24, color: st.c, align: st.al });
      tx(s, T.tiny, st.x - 0.053, st.y + 0.487, 2.309, 0.608, { fontSize: 12, align: st.al, lineSpacingMultiple: 1.3 });
    });
    sh(s, 'roundRect', 8.143, 5.573, 4.342, 1.109, { fill: { color: C.cyan }, rectRadius: adj(28532, 4.342, 1.109) });
    tx(s, 'Feature title three', 8.495, 5.811, 2.283, 0.337, { fontSize: 14, bold: true, color: C.white });
    tx(s, 'Lorem ipsum dolor sit amet, consectetuer', 8.495, 6.115, 3.787, 0.345,
      { fontSize: 12, color: C.white, lineSpacingMultiple: 1.3 });
  })();

  /* -- 21 - process infographic ------------------------------------- */
  (function () {
    const s = add();
    title(s, 'Process Infographic', 2.295, 1.002, 8.744, 0.774, { align: 'center', color: C.black });
    // pale ribbons linking the four disc centres
    [[6.665, 3.171, 9.175, 3.883], [6.665, 3.171, 6.691, 5.846], [6.691, 5.846, 4.184, 5.846]]
      .forEach(function (l) {
        sh(s, 'line', l[0], l[1], l[2] - l[0], l[3] - l[1], { line: { color: C.lilac, width: 38 } });
      });
    const nodes = [
      { x: 5.671, y: 2.177, d: 1.987, fill: C.indigo, ring: 0.132 },
      { x: 8.184, y: 2.892, d: 1.982, fill: C.blue, ring: 0.189 },
      { x: 5.7, y: 4.855, d: 1.982, fill: C.purple, ring: 0.189 },
      { x: 3.193, y: 4.855, d: 1.982, fill: C.navy, ring: 0.189 },
    ];
    nodes.forEach(function (n) {
      sh(s, 'ellipse', n.x, n.y, n.d, n.d, { fill: { color: C.white }, line: { color: C.grey, width: 1 } });
      sh(s, 'ellipse', n.x + n.ring, n.y + n.ring, n.d - n.ring * 2, n.d - n.ring * 2, { fill: { color: n.fill } });
      iconGlyph(s, n.x + n.d * 0.32, n.y + n.d * 0.32, n.d * 0.36, C.white);
    });
    const steps = [
      { t: 'Step One', c: C.navy, x: 0.294, y: 5.341, al: 'right' },
      { t: 'Step Two', c: C.purple, x: 7.808, y: 5.341, al: 'left' },
      { t: 'Step Three', c: C.indigo, x: 2.538, y: 2.705, al: 'right' },
      { t: 'Step Four', c: C.blue, x: 10.41, y: 3.224, al: 'left' },
    ];
    steps.forEach(function (st) {
      feature(s, { x: st.x, y: st.y, w: 2.779, title: st.t, titleColor: st.c, body: T.s, align: st.al });
    });
  })();

  /* -- 22 - employers growth ---------------------------------------- */
  (function () {
    const s = add();
    title(s, ['Employers ', 'Growth'], 0.81, 0.813, 4.067, 0.791);
    tx(s, T.half, 0.866, 2.361, 3.413, 0.602, { fontSize: 12, lineSpacingMultiple: 1.3 });
    // grey "shadow" folds behind the ribbon
    [[8.635, 2.488], [6.319, 3.036], [3.995, 3.582]].forEach(function (p) {
      sh(s, 'parallelogram', p[0], p[1], 2.216, 1.714,
        { fill: { color: C.greyDk }, flipH: true, rectRadius: 1.005 });
    });
    // rising ribbon segments
    const bands = [
      { x: 2.682, y: 3.566, fill: C.navy, label: '15,242', lx: 2.797, ly: 4.309 },
      { x: 5.001, y: 3.026, fill: C.purple, label: '25,682', lx: 5.116, ly: 3.769 },
      { x: 7.311, y: 2.488, fill: C.indigo, label: '35,152', lx: 7.427, ly: 3.206 },
      { x: 9.649, y: 2.172, fill: C.blue, label: '45,442', lx: 9.757, ly: 2.749, wide: true },
    ];
    bands.forEach(function (b) {
      const w = b.wide ? 2.414 : 2.525, h = b.wide ? 2.032 : 2.262;
      sh(s, 'parallelogram', b.x, b.y, w, h,
        { fill: { color: b.fill }, rectRadius: adj(58268, w, h) });
      tx(s, b.label, b.lx, b.ly, 2.295, 0.774,
        { fontFace: HEAD, fontSize: 40, color: C.white, align: 'center', rotate: 300, paraSpaceBefore: 12 });
    });
    // arrow head at the top of the ribbon
    sh(s, 'triangle', 10.455, 1.169, 1.974, 1.006,
      { fill: { color: C.blue }, rectRadius: adj(75991, 1.974, 1.006) });
    [['Revenue 1', 3.26, 5.923], ['Revenue 2', 5.499, 5.439], ['Revenue 2', 7.818, 4.823], ['Revenue 4', 10.191, 4.323]]
      .forEach(function (r) { feature(s, { x: r[1], y: r[2], title: r[0], body: T.s }); });
  })();

  /* -- 23 - six step chain ------------------------------------------ */
  (function () {
    const s = add();
    // pale chain behind the numbered discs
    [[2.072, 3.4, 3.763, 4.156], [3.763, 4.156, 5.453, 3.4], [5.453, 3.4, 7.143, 4.156],
     [7.143, 4.156, 8.834, 3.4], [8.834, 3.4, 10.524, 4.156]].forEach(function (l) {
      sh(s, 'line', l[0], l[1], l[2] - l[0], l[3] - l[1], { line: { color: C.lilac, width: 40 } });
    });
    const nodes = [
      { n: '01', c: C.navy, x: 1.549, y: 2.877, tx: 0.683, ty: 1.388, up: true },
      { n: '02', c: C.purple, x: 3.24, y: 3.633, tx: 2.373, ty: 5.318, up: false },
      { n: '03', c: C.indigo, x: 4.93, y: 2.877, tx: 4.064, ty: 1.386, up: true },
      { n: '04', c: C.blueDk, x: 6.62, y: 3.633, tx: 5.754, ty: 5.318, up: false },
      { n: '05', c: C.blue, x: 8.311, y: 2.877, tx: 7.444, ty: 1.386, up: true },
      { n: '06', c: C.cyan, x: 10.001, y: 3.633, tx: 9.135, ty: 5.318, up: false },
    ];
    nodes.forEach(function (nd) {
      sh(s, 'ellipse', nd.x - 0.18, nd.y - 0.18, 1.406, 1.406, { fill: { color: C.lilac } });
      sh(s, 'ellipse', nd.x, nd.y, 1.046, 1.046, { fill: { color: nd.c } });
      tx(s, nd.n, nd.x, nd.y + 0.25, 1.046, 0.55, { fontFace: HEAD, fontSize: 28, color: C.white, align: 'center' });
      feature(s, { x: nd.tx, y: nd.ty, w: 2.779, title: 'Option One', titleColor: nd.c, body: T.s, align: 'center' });
    });
  })();

  /* -- 24 - chart data (bar + line combo) --------------------------- */
  (function () {
    const s = add();
    sh(s, 'roundRect', 0.918, 3.509, 10.678, 3.542, { fill: { color: C.cyanLt }, rectRadius: adj(8433, 10.678, 3.542) });
    const cats = ['2019', '2020', '2021', '2022', '2023', '2024'];
    s.addChart([
      { type: pres.ChartType.bar, data: [
        { name: 'Series 1', labels: cats, values: [2.6, 5, 3.5, 4.5, 2, 4] },
        { name: 'Series 3', labels: cats, values: [1.8, 4, 3, 5, 4, 5] },
      ], options: { chartColors: [C.navy, C.violet], barGapWidthPct: 269 } },
      { type: pres.ChartType.line, data: [
        { name: 'Series 2', labels: cats, values: [2.4, 4.7, 4, 2, 3.3, 3.8] },
      ], options: { chartColors: [C.purple], lineSize: 3, lineDataSymbol: 'circle', lineDataSymbolSize: 7 } },
    ], {
      x: 1.736, y: 4.249, w: 8.867, h: 2.375,
      showLegend: false, showValue: false,
      chartArea: { fill: { color: C.cyanLt } },
      plotArea: { fill: { color: C.cyanLt } },
      catAxisLabelFontFace: BODY, catAxisLabelFontSize: 20, catAxisLabelColor: C.black,
      catAxisLineColor: C.grey, catAxisMajorTickMark: 'none',
      valAxisLabelFontFace: BODY, valAxisLabelFontSize: 20, valAxisLabelColor: C.black,
      valAxisLineShow: false, valAxisMajorTickMark: 'none',
      valGridLine: { color: 'E4E4E4', size: 1 },
    });
    tx(s, 'Some project title', 1.736, 3.763, 2.648, 0.353, { fontFace: HEAD, fontSize: 15 });
    iconGlyph(s, 8.455, 3.794, 0.29, C.navy);
    s.addText([
      { text: 'December' }, { text: ' ', options: { color: '595959' } }, { text: '2024' },
    ], { x: 8.799, y: 3.785, w: 1.74, h: 0.311, valign: 'top', fontFace: HEAD, fontSize: 12.5, color: C.black });
    sh(s, 'triangle', 8.131, 3.847, 0.129, 0.157, { fill: { color: C.greyMid, transparency: 50 }, rotate: 270 });
    sh(s, 'triangle', 10.395, 3.847, 0.129, 0.157, { fill: { color: C.black }, rotate: 90 });
    title(s, 'Chart Data', 0.814, 1.621, 4.067, 0.734);
    tx(s, T.half, 0.87, 2.359, 3.413, 0.602, { fontSize: 12, lineSpacingMultiple: 1.3 });
    [['3,328', C.navy, 9.293, 9.24], ['720K', C.violet, 6.719, 6.667]].forEach(function (r) {
      tx(s, r[0], r[2], 2.19, 2.309, 0.505, { fontFace: HEAD, fontSize: 24, bold: true, color: r[1] });
      tx(s, T.tiny, r[3], 2.65, 2.309, 0.608, { fontSize: 12, lineSpacingMultiple: 1.3 });
    });
  })();

  /* -- 25 - world map infographic ----------------------------------- */
  (function () {
    const s = add();
    // stylised land masses standing in for the vector world map
    landmass(s, WORLD);
    const gauges = [
      { x: 2.162, y: 1.33, c: C.navy, pct: '50%', label: 'Region -1 ', start: 89.4, end: 270, lx: 3.613, ly: 1.63 },
      { x: 9.358, y: 1.269, c: C.indigo, pct: '60%', label: 'Region -3 ', start: 63.1, end: 270, lx: 10.809, ly: 1.569 },
      { x: 7.035, y: 4.418, c: '2F2997', pct: '70%', label: 'Region -2 ', start: 5.2, end: 270, lx: 8.486, ly: 4.718 },
    ];
    gauges.forEach(function (g) {
      sh(s, 'ellipse', g.x, g.y, 1.366, 1.366, { fill: { color: C.greyLt } });
      sh(s, 'pie', g.x, g.y, 1.366, 1.366, { fill: { color: g.c }, angleRange: [g.start, g.end] });
      sh(s, 'ellipse', g.x + 0.105, g.y + 0.105, 1.155, 1.155, { fill: { color: C.white } });
      tx(s, g.pct, g.x + 0.142, g.y + 0.431, 1.082, 0.505,
        { fontFace: HEAD, fontSize: 24, bold: true, color: g.c, align: 'center' });
      sh(s, 'round1Rect', g.lx, g.ly, 1.35, 0.303, { fill: { color: g.c }, rectRadius: adj(39591, 1.35, 0.303) });
      tx(s, g.label, g.lx, g.ly - 0.03, 1.35, 0.303, { fontFace: HEAD, fontSize: 12, color: C.white, align: 'center' });
    });
    sh(s, 'round2SameRect', 2.296, 2.906, 1.335, 6.027,
      { fill: { color: C.cyanLt }, rotate: 270, flipV: true, rectRadius: adj(21659, 1.335, 6.027) });
    title(s, 'Map Infographic', 0.805, 3.678, 4.067, 1.421);
    tx(s, T.map, 0.886, 5.484, 4.279, 0.673, { fontSize: 12, lineSpacingMultiple: 1.5 });
  })();

  /* -- 26 - europe map infographic ---------------------------------- */
  (function () {
    const s = add();
    landmass(s, EUROPE);
    // the two highlighted nations
    sh(s, 'roundRect', 8.644, 3.878, 0.984, 0.928, { fill: { color: C.blueDk }, rectRadius: 0.18 });
    sh(s, 'roundRect', 6.755, 4.486, 1.441, 1.382, { fill: { color: C.cyan }, rectRadius: 0.25 });
    title(s, 'Map Infographic', 0.934, 1.156, 4.067, 1.421);
    tx(s, T.map, 0.952, 2.676, 4.279, 0.673, { fontSize: 12, lineSpacingMultiple: 1.5 });
    [{ x: 0.938, fill: C.blueDk, name: 'Nation One' }, { x: 3.284, fill: C.cyan, name: 'Nation Two' }]
      .forEach(function (card, i) {
        const y = 4.083 + i * 0.006;
        sh(s, 'roundRect', card.x, y, 2.102, 2.34, { fill: { color: card.fill }, rectRadius: adj(13162, 2.102, 2.34) });
        tx(s, card.name, card.x + 0.195, y + 0.311, 1.879, 0.337, { fontSize: 14, bold: true, charSpacing: 0.5, color: C.white });
        tx(s, 'Lorem ipsum dolor sit amet, veniam elit a elit posuere gravida.', card.x + 0.195, y + 0.628, 1.879, 0.803,
          { fontSize: 11, color: C.white, lineSpacingMultiple: 1.3 });
        tx(s, 'Stats', card.x + 0.198, y + 1.554, 0.978, 0.321, { fontSize: 11, color: C.white, lineSpacingMultiple: 1.3 });
        tx(s, '80%', card.x + 1.399, y + 1.554, 0.52, 0.321, { fontSize: 11, color: C.white, align: 'right', lineSpacingMultiple: 1.3 });
        sh(s, 'rect', card.x + 0.236, y + 1.922, 1.674, 0.09, { fill: { color: C.white, transparency: 80 } });
        sh(s, 'rect', card.x + 0.223, y + 1.922, 1.338, 0.09, { fill: { color: C.white } });
      });
    [{ n: '02', x: 7.086, y: 4.19 }, { n: '01', x: 9.19, y: 3.519 }].forEach(function (p) {
      sh(s, 'ellipse', p.x, p.y, 0.608, 0.608, { fill: { color: C.navy } });
      tx(s, p.n, p.x + 0.053, p.y + 0.092, 0.502, 0.404, { align: 'center', bold: true, color: C.white });
    });
  })();

  /* -- 27 - chart data report (3-D bars) ---------------------------- */
  (function () {
    const s = add();
    title(s, 'Chart Data Report', 0.843, 1.444, 4.067, 1.379);
    tx(s, T.half, 0.9, 2.991, 3.59, 0.602, { fontSize: 12, lineSpacingMultiple: 1.3 });
    const cats = ['Category 1', 'Category 2', 'Category 3', 'Category 4'];
    s.addChart(pres.ChartType.bar3d, [
      { name: 'Series 1', labels: cats, values: [2.6, 3.8, 4.4, 1.5] },
      { name: 'Series 2', labels: cats, values: [3.3, 1.2, 3.6, 4.7] },
      { name: 'Series 3', labels: cats, values: [2, 3.2, 1.7, 1.4] },
    ], {
      x: 6.958, y: 1.041, w: 4.903, h: 3.106,
      barGrouping: 'standard', bar3DShape: 'box', barGapWidthPct: 150,
      v3DRotX: 15, v3DRotY: 20, v3DPerspective: 30, v3DRAngAx: false,
      chartColors: [C.navy, C.cyanDk, C.indigo],
      showLegend: false, showValue: false, catAxisHidden: true, serAxisHidden: true,
      valAxisLabelFontFace: BODY, valAxisLabelFontSize: 12, valAxisLabelColor: C.black,
      valAxisLineShow: false, valGridLine: { color: 'BFBFBF', size: 1 },
    });
    [{ n: '01', c: C.navy, x: 2.811, tx: 1.797 }, { n: '02', c: C.indigo, x: 6.363, tx: 5.348 },
     { n: '03', c: C.cyan, x: 9.914, tx: 8.899 }].forEach(function (p) {
      sh(s, 'ellipse', p.x, 4.466, 0.608, 0.608, { fill: { color: p.c } });
      tx(s, p.n, p.x + 0.053, 4.558, 0.502, 0.404, { align: 'center', bold: true, color: C.white });
      feature(s, { x: p.tx, y: 5.21, title: 'Lorem ipsum', body: T.s, align: 'center' });
    });
  })();

  /* -- 28 - we specialize in problem solving ------------------------ */
  (function () {
    const s = add();
    sh(s, 'line', 0, 3.216, 7.785, 4.284, { line: { color: C.lilac, width: 4 } });
    const pennants = [
      { x: 1.412, y: 2.072, w: 1.075, h: 2.481, c: C.navy, ring: '5B5FF4', cx: 1.607, cy: 2.501 },
      { x: 3.117, y: 3.39, w: 1.062, h: 2.115, c: C.purple, ring: C.violet, cx: 3.306, cy: 3.859 },
      { x: 4.94, y: 4.533, w: 1.022, h: 1.969, c: C.blue, ring: 'B2AFE9', cx: 5.109, cy: 4.94 },
      { x: 6.645, y: 5.735, w: 1.022, h: 1.708, c: C.cyan, ring: 'B6D5F9', cx: 6.814, cy: 6.121 },
    ];
    pennants.forEach(function (p) {
      const head = p.h * 0.21, tail = p.h * 0.25;
      sh(s, 'triangle', p.x, p.y, p.w, head, { fill: { color: p.c } });
      sh(s, 'rect', p.x, p.y + head, p.w, p.h - head - tail, { fill: { color: p.c } });
      sh(s, 'rtTriangle', p.x, p.y + p.h - tail, p.w, tail, { fill: { color: p.c }, flipH: true, flipV: true });
      sh(s, 'ellipse', p.cx, p.cy, 0.685, 0.685, { fill: { color: C.white }, line: { color: p.ring, width: 3 } });
      iconGlyph(s, p.cx + 0.19, p.cy + 0.19, 0.303, p.c);
    });
    tx(s, [{ text: 'We Specialize', options: { breakLine: true } }, { text: 'In Problem Solving' }],
      7.499, 1.243, 4.554, 2.121, { fontFace: HEAD, fontSize: 44, align: 'right', lineSpacingMultiple: 0.9 });
    [{ x: 2.088, y: 1.047, c: C.navy }, { x: 3.922, y: 2.292, c: C.purple },
     { x: 5.686, y: 3.464, c: C.blue }, { x: 7.498, y: 4.733, c: C.cyan }].forEach(function (p) {
      sh(s, 'line', p.x - 0.136, p.y + 0.31, 0, 0.582, { line: { color: p.c, width: 3.25 } });
      feature(s, { x: p.x, y: p.y, title: 'Lorem ipsum', titleColor: p.c, body: T.s });
    });
    tx(s, T.s, 1.336, 6.112, 2.508, 0.602, { fontSize: 12, lineSpacingMultiple: 1.3 });
  })();

  /* -- 29 - creative infographic (exploded doughnut) ---------------- */
  (function () {
    const s = add();
    sh(s, 'round2SameRect', 8.862, 1.929, 1.776, 7.167,
      { fill: { color: C.cyanLt }, rotate: 270, rectRadius: adj(24877, 1.776, 7.167) });
    sh(s, 'ellipse', 1.454, 2.531, 3.333, 3.333, { fill: { color: C.lilac }, line: { color: C.greyLt, width: 25 } });
    sh(s, 'pie', 1.294, 2.044, 4.18, 4.18, { fill: { color: C.navy }, angleRange: [90.4, 151.3], rotate: 224.9 });
    sh(s, 'pie', 1.183, 1.933, 4.403, 4.403, { fill: { color: C.cyan }, angleRange: [106.2, 188.4], rotate: 270 });
    sh(s, 'pie', 1.075, 1.841, 4.619, 4.619, { fill: { color: C.blue }, angleRange: [6, 135.4], rotate: 180 });
    sh(s, 'ellipse', 2.617, 3.368, 1.533, 1.533, { fill: { color: C.white } });
    iconGlyph(s, 2.987, 3.71, 0.774, C.navy);
    [['57%', 2.444, 2.572], ['18%', 4.217, 3.528], ['25%', 3.582, 5.1]].forEach(function (r) {
      tx(s, r[0], r[1], r[2], 1.138, 0.616, { fontFace: HEAD, fontSize: 24, color: C.white, align: 'center' });
    });
    title(s, 'Creative Infographic', 6.571, 1.478, 4.314, 1.582);
    tx(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer vitae justo ullamcorper, scelerisque mi quis, ornare erat. Lorem ipsum dolor consectetur adipiscing elit. Integer vitae justo ullamcorper, scelerisque ',
      6.625, 3.196, 5.634, 0.976, { fontSize: 12, lineSpacingMultiple: 1.5 });
    [['85+', 'Lorem ipsum dolor sit amet,.', 6.576], ['273', 'Lorem ipsum dolor sit amet,', 8.388],
     ['53K', 'Lorem ipsum dolor sit amet,', 10.193]].forEach(function (r, i) {
      tx(s, r[0], r[2], 4.896, 1.385, 0.572, { fontFace: HEAD, fontSize: 28, bold: true, color: C.navy });
      tx(s, r[1], r[2], 5.483, 1.646, 0.605, { fontSize: 12, color: C.navy, lineSpacingMultiple: 1.3 });
      if (i > 0) sh(s, 'line', r[2] - 0.183, 4.955, 0, 1.144, { line: { color: C.navy, width: 2 } });
    });
  })();

  /* -- 30 - contact ------------------------------------------------- */
  (function () {
    const s = add();
    sh(s, 'rect', 0, 0, 13.333, 7.5, { fill: { color: C.blueDk, transparency: 40 } });
    title(s, 'Contact Information ', 1.203, 1.58, 4.656, 1.552, { color: C.white });
    [['Company address 1234, A', 3.508], ['www.companysite.com', 4.295],
     ['+0782 4022 XXXX XXX', 5.081], ['mail@companysite.com', 5.866]].forEach(function (r) {
      sh(s, 'ellipse', 1.249, r[1], 0.534, 0.534, { fill: { color: C.white } });
      iconGlyph(s, 1.383, r[1] + 0.128, 0.28, C.blueDk);
      tx(s, r[0], 1.941, r[1] + 0.084, 2.961, 0.368,
        { fontFace: HEAD, fontSize: 14, color: C.white, lineSpacingMultiple: 1.2 });
    });
    title(s, 'Carrera', 6.725, 2.973, 6.009, 1.587, { fontSize: 100, color: C.white, align: 'center' });
  })();

  // the logo always sits on top of whatever artwork the slide draws
  pendingLogos.forEach(function (p) { logo(p[0], p[1]); });

  return pres;
}

/* ------------------------------------------------------------------ main */

build()
  .writeFile({ fileName: path.join(__dirname, '16452d51-98ac-4c8f-8895-bd4132915e61_grok_final.pptx') })
  .then(function (f) { console.log('wrote ' + f); })
  .catch(function (e) { console.error(e); process.exit(1); });
