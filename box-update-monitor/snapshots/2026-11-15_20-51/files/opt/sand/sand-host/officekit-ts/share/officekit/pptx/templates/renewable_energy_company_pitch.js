/**
 * Culka — Solar & Renewable Energy business presentation (25 slides, 13.333" x 7.5").
 *
 * Standalone pptxgenjs re-creation of the reference deck. Run with:
 *     node 0ca89487-79d9-4745-928c-5592fee7451e_grok_final.js
 * which writes 0ca89487-79d9-4745-928c-5592fee7451e_grok_final.pptx next to this file.
 *
 * Layout notes:
 *  - Every slide carries the same footer furniture (see `chrome`): a grey pill with
 *    "Culka Company", the "Business Presentation / Template" block and the date block.
 *  - The deck has no embedded bitmaps; its picture placeholders are drawn by `imageBox`
 *    as labelled outlines at the original placeholder geometry.
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const GREEN = '4B896D';      // brand green
const WHITE = 'FFFFFF';
const BLACK = '000000';
const INK = '262626';        // body copy on light backgrounds (tx1 lumMod 85%)
const GREY05 = 'F2F2F2';     // bg1 lumMod 95% — panels, pills, number tiles
const GREY15 = 'D9D9D9';     // bar-chart middle series
const GREY25 = 'BFBFBF';     // hairline rules, "Get Started" buttons
const GREY35 = 'A6A6A6';     // bg1 lumMod 85% — progress-bar troughs
const NEARBLACK = '0D0D0D';  // tx1 lumMod 95% divider rules
const SILVER = 'A5A5A5';     // bar-chart third series
const ICE = 'F3FCFF';        // rating stars

/* -------------------------------------------------------------------- fonts */
const REG = 'Maven Pro';
const MED = 'Maven Pro Medium';
const SEMI = 'Maven Pro SemiBold';
const BARLOW = 'Barlow';

/* ------------------------------------------------- repeated body copy blocks */
const COPY1 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eiusmod tempor incididunt ut labore et dolore magna aliqua. Dignissim diam quis enim lobortis. Tincidunt tortor aliquam nulla facilisi cras fermentum odio.';
const COPY2 =
  'Laoreet suspendisse interdum consectetur libero id faucibus nisl tincidunt. Arcu risus quis varius quam quisque id diam velolutpat.';
const COPY3 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eiusmod tempor incididunt ut labore et dolore magna aliqua. Dignissim diam quis enim lobortis. Tincidunt tortor aliquam nulla.';
const COPY4 =
  'Vel turpis nunc eget lorem dolor sed viverra ipsum. Laoreet suspendisse interdum consectetur libero id faucibus nisl tincidunt. Arcu risus quis varius quam quisque id diam vel. Volutpat consequat mauris nunc.';
const COPY5 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eiusmod tempor incididunt ut labore et dolore magna aliqua. Dignissim diam quis enim lobortis. Tincidunt tortor aliquam nulla facilisi cras fermentum odio. ';
const COPY6 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eiusmod tempor';
const COPY7 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eiusmod tempor incididunt ut labore et.';
const COPY8 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eius.';
const COPY9 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eiusmod tempor incididunt ut labore.';

/* ------------------------------------------------------------- primitives */

/** Text box: the deck always uses zero insets and top-anchored text. */
function tx(s, text, opts) {
  s.addText(text, Object.assign({ margin: 0, valign: 'top' }, opts));
}

/** Body copy: 12pt Maven Pro at 1.5 line spacing — the deck's default paragraph. */
function body(s, text, x, y, w, h, color) {
  tx(s, text, { x, y, w, h, fontSize: 12, fontFace: REG, lineSpacingMultiple: 1.5, color, fit: 'resize' });
}

/** Display heading: Maven Pro SemiBold, never wrapped. */
function head(s, text, x, y, w, h, fontSize, color) {
  tx(s, text, { x, y, w, h, fontSize, fontFace: SEMI, color, wrap: false, fit: 'resize' });
}

/** Flat rectangle (no outline). */
function box(s, x, y, w, h, fill) {
  s.addShape('rect', { x, y, w, h, fill: fill ? { color: fill } : { type: 'none' } });
}

/** Straight connector. */
function rule(s, x, y, w, h, color, width) {
  s.addShape('line', { x, y, w, h, line: { color, width } });
}

/**
 * Stand-in for one of the deck's picture placeholders: a labelled outline at the
 * original placeholder geometry. (The reference deck ships these placeholders empty.)
 */
function imageBox(s, x, y, w, h) {
  s.addShape('rect', { x, y, w, h, fill: { type: 'none' }, line: { color: GREY15, width: 1 } });
  tx(s, '[image]', { x, y: y + h / 2 - 0.13, w, h: 0.26, fontSize: 9, fontFace: REG, color: GREY25, align: 'center' });
}

/* ------------------------------------------------------- repeated furniture */

/** Grey square with "S+" followed by a green "The Best Company" bar. */
function badge(s, x, y) {
  box(s, x, y, 0.622, 0.622, GREY05);
  box(s, x + 0.744, y, 3.069, 0.622, GREEN);
  tx(s, 'S+', { x: x + 0.171, y: y + 0.16, w: 0.281, h: 0.303, fontFace: MED, align: 'center', wrap: false });
  tx(s, 'The Best Company', {
    x: x + 1.176, y: y + 0.162, w: 2.205, h: 0.303, fontFace: MED, color: WHITE, align: 'center', wrap: false
  });
}

/** Small green call-to-action button. */
function learnMore(s, x, y) {
  box(s, x, y, 1.229, 0.375, GREEN);
  tx(s, 'Lean More', {
    x: x + 0.158, y: y + 0.078, w: 0.887, h: 0.219, fontSize: 13, fontFace: MED, color: WHITE, wrap: false
  });
}

/** Footer furniture repeated on every slide. */
function chrome(s) {
  s.addShape('rect', { x: 0.677, y: 6.572, w: 1.879, h: 0.387, fill: { color: GREY05 }, line: { color: WHITE, width: 1 } });
  tx(s, 'Culka Company', { x: 1.018, y: 6.664, w: 1.196, h: 0.202, fontSize: 12, fontFace: REG, align: 'center', wrap: false });
  tx(s, [
    { text: 'Business', options: { fontSize: 10, fontFace: REG, bold: true } },
    { text: ' Presentation', options: { fontSize: 10, fontFace: REG, breakLine: true } },
    { text: 'Template', options: { fontSize: 10, fontFace: REG } }
  ], { x: 9.174, y: 6.619, w: 1.434, h: 0.337, wrap: false });
  tx(s, [
    { text: 'Date', options: { fontSize: 10, fontFace: REG, bold: true, breakLine: true } },
    { text: '4 Mar 2023', options: { fontSize: 10, fontFace: REG } }
  ], { x: 11.924, y: 6.619, w: 0.738, h: 0.337, wrap: false });
}

/* ---------------------------------------------------------------- slides */

/* Slide 1 — Cover — CULKA */
function slide01(s) {
  head(s, 'CULKA', 0.678, 3.098, 6.25, 2.322, 138);
  rule(s, 0.667, 3.005, 12, 0, GREY25, 1.5);
  tx(s, 'Solar & Renewable Energy Presentation', { fontSize: 20, fontFace: MED, color: INK, x: 0.667, y: 5.249, w: 5.173, h: 0.337, wrap: false, fit: 'resize' });
  body(s, COPY3, 7.711, 3.625, 4.955, 0.869);
  badge(s, 7.711, 4.963);
  imageBox(s, 0.672, 0, 11.985, 2.722);
  chrome(s);
}

/* Slide 2 — Welcome To Culka */
function slide02(s) {
  box(s, 10.144, 0, 3.189, 2.256, GREEN);
  tx(s, 'Welcome To Culka', { fontSize: 28, fontFace: MED, x: 0.678, y: 0.716, w: 3.385, h: 0.471, wrap: false, fit: 'resize' });
  head(s, 'The Best Solution For', 0.678, 1.327, 6.255, 0.741, 44);
  head(s, 'Wind & Solar Energy', 0.678, 2.019, 5.95, 0.741, 44);
  body(s, COPY1, 0.677, 3.265, 5.601, 0.869);
  body(s, COPY2, 0.677, 4.463, 5.012, 0.566);
  learnMore(s, 0.68, 5.582);
  imageBox(s, 8.722, 0.542, 3.932, 5.415);
  chrome(s);
}

/* Slide 3 — How We Started Industry */
function slide03(s) {
  box(s, 0.677, 0.542, 4.912, 2.403, GREEN);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eiusmod tempor incididunt ut labore et dolore magna aliqua. Laoreet suspendisse interdum', 1.121, 1.633, 3.94, 0.869, WHITE);
  head(s, 'How We Started Industry', 1.121, 0.985, 4.023, 0.404, 24, WHITE);
  head(s, 'Green Technology', 9.407, 0.55, 1.588, 0.219, 13);
  body(s, COPY6, 9.407, 0.89, 3.25, 0.581);
  head(s, 'Renewable Technology', 9.407, 1.833, 2.03, 0.219, 13);
  body(s, COPY6, 9.407, 2.173, 3.25, 0.581);
  box(s, 8.438, 0.545, 0.625, 0.625, GREY05);
  tx(s, '01', { fontSize: 24, fontFace: SEMI, align: 'center', x: 8.561, y: 0.655, w: 0.379, h: 0.404, wrap: false, fit: 'resize' });
  box(s, 8.438, 1.836, 0.625, 0.625, GREY05);
  tx(s, '02', { fontSize: 24, fontFace: SEMI, align: 'center', x: 8.539, y: 1.946, w: 0.424, h: 0.404, wrap: false, fit: 'resize' });
  imageBox(s, 0.667, 3.289, 12.024, 2.933);
  chrome(s);
}

/* Slide 4 — About Culka Company */
function slide04(s) {
  box(s, 9.905, 0, 3.402, 7.5, GREY05);
  box(s, 0.678, 3.735, 6.209, 1.333, GREEN);
  tx(s, 'About Culka Company', { fontFace: MED, x: 0.678, y: 0.542, w: 2.607, h: 0.303, wrap: false, fit: 'resize' });
  head(s, 'We Have The Best In Class', 0.678, 1.013, 6.327, 0.606, 36);
  head(s, 'Product And Solution', 0.678, 1.591, 5.144, 0.606, 36);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eiusmod tempor incididunt ut labore et dolore magna aliqua. Dignissim diam quis enim lobortis. Tincidunt tortor aliquam nulla facilisi cras fermentum odio consectetur.', 0.678, 2.524, 6.209, 0.869, INK);
  body(s, 'Dignissim diam quis enim lobortis. Tincidunt tortor aliquam nulla facilisi cras fermentum odio. Consectetur adipiscing elit, sed.', 0.678, 5.343, 5.327, 0.566, INK);
  s.addShape('star5', { x: 3.769, y: 4.389, w: 0.232, h: 0.232, fill: { color: ICE } });
  s.addShape('star5', { x: 4.093, y: 4.389, w: 0.232, h: 0.232, fill: { color: ICE } });
  s.addShape('star5', { x: 4.42, y: 4.389, w: 0.232, h: 0.232, fill: { color: ICE } });
  s.addShape('star5', { x: 4.74, y: 4.389, w: 0.232, h: 0.232, fill: { color: ICE } });
  s.addShape('star5', { x: 5.06, y: 4.389, w: 0.232, h: 0.232, fill: { color: ICE } });
  tx(s, '46K ', {
    fontSize: 36,
    fontFace: SEMI,
    color: WHITE,
    lineSpacingMultiple: 1.5,
    x: 1.557,
    y: 3.783,
    w: 1.12,
    h: 0.813,
    wrap: false,
    fit: 'resize'
  });
  tx(s, 'Megawatss Of Capacity', { fontSize: 12, fontFace: MED, color: WHITE, x: 1.173, y: 4.635, w: 1.888, h: 0.202, wrap: false, fit: 'resize' });
  tx(s, '9.8/10', { fontSize: 12, fontFace: MED, color: WHITE, x: 5.822, y: 4.404, w: 0.507, h: 0.202, wrap: false, fit: 'resize' });
  imageBox(s, 8.721, 0.542, 3.932, 5.415);
  chrome(s);
}

/* Slide 5 — We Build Sustainable Energy For Future */
function slide05(s) {
  box(s, 2.633, 2.561, 4.744, 1.367, GREEN);
  head(s, 'We Build Sustainable', 0.677, 0.634, 6.15, 0.741, 44);
  head(s, 'Energy For Future', 0.677, 1.245, 5.282, 0.741, 44);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eiusmod tempor incididunt ut labore et dolore magna aliqua. Dignissim diam quis enim.', 3.032, 2.788, 3.969, 0.869, WHITE);
  head(s, 'Best Energy Solution', 0.677, 4.503, 2.479, 0.303, 18);
  body(s, COPY9, 0.677, 4.953, 2.842, 0.869);
  head(s, 'End – To End- Service', 4.172, 4.503, 2.612, 0.303, 18);
  body(s, COPY9, 4.172, 4.953, 2.842, 0.869);
  imageBox(s, 0.68, 2.583, 1.876, 1.367);
  imageBox(s, 7.777, 0.542, 4.877, 5.415);
  chrome(s);
}

/* Slide 6 — A Leader In The Wind Energy Space */
function slide06(s) {
  box(s, 3.183, 4.267, 2.109, 1.175, GREEN);
  box(s, 0.677, 4.267, 2.109, 1.175, GREEN);
  head(s, 'A Leader In The Wind', 0.677, 1.052, 5.689, 0.673, 40);
  head(s, 'Energy Space Since 2020', 0.677, 1.643, 6.721, 0.673, 40);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eiusmod tempor incididunt ut labore et dolore magna aliqua. Dignissim diam quis enim lobortis. ', 0.677, 2.768, 4.216, 0.869);
  head(s, '16300+', 0.923, 4.476, 1.39, 0.471, 28, WHITE);
  tx(s, 'Installed Capacity', { fontSize: 14, fontFace: MED, color: WHITE, x: 0.923, y: 4.997, w: 1.616, h: 0.236, wrap: false, fit: 'resize' });
  head(s, '3210+', 3.356, 4.476, 1.127, 0.471, 28, WHITE);
  tx(s, 'Customers Globally', { fontSize: 14, fontFace: MED, color: WHITE, x: 3.356, y: 4.997, w: 1.762, h: 0.236, wrap: false, fit: 'resize' });
  imageBox(s, 8.042, 1.052, 5.292, 4.389);
  chrome(s);
}

/* Slide 7 — We Are Building A Sustainable Future */
function slide07(s) {
  head(s, 'We Are Building A', 0.677, 0.542, 6.493, 0.909, 54);
  head(s, 'Sustainable Future', 0.677, 1.35, 6.711, 0.909, 54);
  body(s, COPY1, 0.677, 2.867, 5.601, 0.869);
  body(s, COPY2, 0.677, 4.065, 5.012, 0.566);
  learnMore(s, 0.68, 5.184);
  box(s, 12.667, 1.35, 0.667, 1.382, GREEN);
  imageBox(s, 7.388, 2.732, 5.266, 2.827);
  chrome(s);
}

/* Slide 8 — Greener Tomorrow For Everyone — progress bars */
function slide08(s) {
  box(s, 0.677, 4.212, 3.916, 0.131, GREY15);
  head(s, 'Greener Tomorrow', 0.677, 0.542, 6.038, 0.808, 48);
  head(s, 'For Everyone', 0.677, 1.203, 4.256, 0.808, 48);
  body(s, COPY5, 0.677, 2.397, 5.657, 0.869, INK);
  box(s, 0.677, 4.212, 3.501, 0.131, GREEN);
  head(s, 'Maintenance', 0.677, 3.865, 1.015, 0.202, 12);
  tx(s, '90%', { fontSize: 12, fontFace: MED, x: 4.2, y: 3.869, w: 0.345, h: 0.202, wrap: false, fit: 'resize' });
  box(s, 0.677, 5.635, 3.916, 0.131, GREY15);
  box(s, 0.677, 5.635, 3.055, 0.131, GREEN);
  head(s, 'Fleet Optimazation', 0.677, 5.288, 1.564, 0.202, 12);
  tx(s, '75%', { fontSize: 12, fontFace: MED, x: 4.2, y: 5.291, w: 0.33, h: 0.202, wrap: false, fit: 'resize' });
  box(s, 0.677, 4.923, 3.916, 0.131, GREY15);
  box(s, 0.677, 4.923, 3.288, 0.131, GREEN);
  head(s, 'Part & Repair', 0.677, 4.576, 1.062, 0.202, 12);
  tx(s, '80%', { fontSize: 12, fontFace: MED, x: 4.2, y: 4.58, w: 0.345, h: 0.202, wrap: false, fit: 'resize' });
  box(s, 10.48, 0.542, 2.201, 2.599, GREEN);
  tx(s, '4', {
    fontSize: 66,
    fontFace: SEMI,
    color: WHITE,
    align: 'center',
    x: 11.304,
    y: 0.877,
    w: 0.554,
    h: 1.111,
    wrap: false,
    fit: 'resize'
  });
  tx(s, 'Years', {
    fontSize: 20,
    fontFace: SEMI,
    color: WHITE,
    align: 'center',
    x: 11.209,
    y: 2.078,
    w: 0.743,
    h: 0.337,
    wrap: false,
    fit: 'resize'
  });
  tx(s, 'Experience', {
    fontSize: 20,
    fontFace: SEMI,
    color: WHITE,
    align: 'center',
    x: 10.837,
    y: 2.469,
    w: 1.488,
    h: 0.337,
    wrap: false,
    fit: 'resize'
  });
  imageBox(s, 7.59, 0.542, 2.865, 5.224);
  imageBox(s, 10.48, 3.167, 2.201, 2.599);
  chrome(s);
}

/* Slide 9 — We Use Solar Power To Simplify Your Life */
function slide09(s) {
  head(s, 'We Use Solar Power To', 0.677, 0.634, 8.346, 0.909, 54);
  head(s, 'Simplify Your Life', 0.677, 1.527, 6.388, 0.909, 54);
  body(s, COPY1, 7.094, 3.161, 5.601, 0.869);
  body(s, COPY2, 7.094, 4.359, 5.012, 0.566);
  learnMore(s, 7.097, 5.478);
  box(s, 12.667, 0, 0.667, 1.543, GREEN);
  imageBox(s, 0, 3.161, 6.345, 2.692);
  chrome(s);
}

/* Slide 10 — Founder Of Culka Company */
function slide10(s) {
  box(s, 8.344, 4.046, 3.9, 1.343, GREEN);
  tx(s, 'Founder Of', { fontSize: 44, fontFace: SEMI, x: 0.642, y: 0.542, w: 4.957, h: 0.741, fit: 'resize' });
  tx(s, 'Culka Company', { fontSize: 44, fontFace: SEMI, x: 0.642, y: 1.222, w: 5.209, h: 0.741, fit: 'resize' });
  body(s, COPY4, 0.677, 2.521, 5.545, 0.869);
  rule(s, 0.667, 3.894, 5.556, 0, GREY25, 1.5);
  tx(s, 'Leadership', { fontSize: 11, fontFace: SEMI, x: 0.679, y: 4.603, w: 1.509, h: 0.185, fit: 'resize' });
  tx(s, 'Management', { fontSize: 11, fontFace: SEMI, x: 0.679, y: 5.089, w: 1.509, h: 0.185, fit: 'resize' });
  tx(s, 'Skills', { fontSize: 11, fontFace: SEMI, x: 0.679, y: 5.575, w: 1.187, h: 0.185, fit: 'resize' });
  box(s, 2.422, 4.645, 2.711, 0.109, GREY15);
  box(s, 2.422, 5.141, 2.711, 0.105, GREY15);
  box(s, 2.422, 5.622, 2.711, 0.109, GREY15);
  box(s, 2.422, 4.645, 2.163, 0.109, GREEN);
  box(s, 2.422, 5.141, 2.324, 0.105, GREEN);
  box(s, 2.422, 5.622, 1.708, 0.109, GREEN);
  tx(s, '87+', { fontSize: 12, fontFace: MED, align: 'right', x: 5.326, y: 4.603, w: 0.525, h: 0.202, fit: 'resize' });
  tx(s, '95+', { fontSize: 12, fontFace: MED, align: 'right', x: 5.326, y: 5.089, w: 0.525, h: 0.202, fit: 'resize' });
  tx(s, '82+', { fontSize: 12, fontFace: MED, align: 'right', x: 5.326, y: 5.575, w: 0.525, h: 0.202, fit: 'resize' });
  tx(s, 'Johan Ramsey', { fontSize: 36, fontFace: SEMI, color: WHITE, x: 8.563, y: 4.173, w: 3.53, h: 0.606, fit: 'resize' });
  tx(s, 'Culka Founder', { fontSize: 16, fontFace: REG, color: WHITE, x: 8.563, y: 4.849, w: 1.509, h: 0.269, fit: 'resize' });
  imageBox(s, 8.344, 0.542, 4.322, 3.208);
  chrome(s);
}

/* Slide 11 — This Is Our Best Team */
function slide11(s) {
  tx(s, [
    { text: 'This Is ', options: { fontSize: 54, fontFace: SEMI } },
    { text: 'Our', options: { fontSize: 54, fontFace: SEMI, bold: true } }
  ], { x: 0.673, y: 0.542, w: 5.509, h: 0.909, fit: 'resize' });
  tx(s, [
    { text: 'Best', options: { fontSize: 54, fontFace: SEMI, bold: true } },
    { text: ' Team', options: { fontSize: 54, fontFace: SEMI } }
  ], { x: 0.673, y: 1.353, w: 5.171, h: 0.909, fit: 'resize' });
  s.addShape('rect', { x: 7.823, y: 2.976, w: 1.879, h: 0.387, fill: { color: GREEN }, line: { color: WHITE, width: 1 } });
  tx(s, 'Jack Bruce', {
    fontSize: 12,
    fontFace: MED,
    color: WHITE,
    align: 'center',
    x: 8.339,
    y: 3.068,
    w: 0.847,
    h: 0.202,
    wrap: false,
    fit: 'resize'
  });
  s.addShape('rect', { x: 10.626, y: 2.976, w: 1.879, h: 0.387, fill: { color: GREEN }, line: { color: WHITE, width: 1 } });
  tx(s, 'Alan Darlan', {
    fontSize: 12,
    fontFace: MED,
    color: WHITE,
    align: 'center',
    x: 11.113,
    y: 3.068,
    w: 0.906,
    h: 0.202,
    wrap: false,
    fit: 'resize'
  });
  tx(s, 'Laoreet suspendisse interdum consectetur liber.', {
    fontSize: 12,
    fontFace: REG,
    align: 'center',
    lineSpacingMultiple: 1.5,
    x: 7.662,
    y: 3.485,
    w: 2.202,
    h: 0.566,
    fit: 'resize'
  });
  tx(s, 'Consectetur libero id faucibus laoreet.', {
    fontSize: 12,
    fontFace: REG,
    align: 'center',
    lineSpacingMultiple: 1.5,
    x: 10.557,
    y: 3.485,
    w: 2.017,
    h: 0.566,
    fit: 'resize'
  });
  s.addShape('rect', { x: 3.823, y: 3.484, w: 1.879, h: 0.387, fill: { color: GREEN }, line: { color: WHITE, width: 1 } });
  tx(s, 'Nia Darci', {
    fontSize: 12,
    fontFace: MED,
    color: WHITE,
    align: 'center',
    x: 4.407,
    y: 3.576,
    w: 0.712,
    h: 0.202,
    wrap: false,
    fit: 'resize'
  });
  tx(s, 'Team Leader', { fontSize: 12, fontFace: MED, lineSpacingMultiple: 1.5, x: 3.823, y: 3.995, w: 1.273, h: 0.268, fit: 'resize' });
  rule(s, 3.823, 4.384, 0.823, 0, NEARBLACK, 1.5);
  body(s, 'Vel turpis nunc eget lorem dolor sed viverra ipsum. Laoreet suspendisse interdum consectetur libero.', 3.823, 4.928, 2.753, 0.869);
  rule(s, 7.662, 4.323, 5.005, 0, NEARBLACK, 1.5);
  body(s, 'Laoreet suspendisse interdum consectetur libero id faucibus nisl tincidunt. Arcu risus quis varius quam quisque id diam vel. ', 7.662, 4.928, 5.005, 0.566);
  imageBox(s, 0.667, 3.484, 2.707, 2.276);
  imageBox(s, 7.662, 0.542, 2.202, 2.276);
  imageBox(s, 10.465, 0.542, 2.202, 2.276);
  chrome(s);
}

/* Slide 12 — Image Snapshot */
function slide12(s) {
  box(s, 9.905, 0, 3.402, 7.5, GREY05);
  tx(s, 'Image Snapshot', { fontSize: 54, fontFace: SEMI, color: BLACK, x: 0.677, y: 0.892, w: 6.267, h: 0.909, fit: 'resize' });
  body(s, 'Vel turpis nunc eget lorem dolor sed viverra ipsum. Laoreet suspendisse interdum consectetur libero id faucibus nisl tincidunt. Arcu risus quis varius quam quisque id diam vel.', 0.677, 2.422, 4.819, 0.88);
  tx(s, [
    { text: 'Laoreet suspendisse interdum consectetur libero.', options: {
      fontSize: 12,
      fontFace: REG,
      bold: true,
      lineSpacingMultiple: 1.5,
      bullet: { characterCode: '2022', indent: 13.5 },
      breakLine: true
    } },
    { text: 'faucibus nisl tincidunt arcu risus quis varius quam.', options: { fontSize: 12, fontFace: REG, bold: true, lineSpacingMultiple: 1.5, bullet: { characterCode: '2022', indent: 13.5 } } }
  ], { x: 0.677, y: 5.029, w: 4.4, h: 0.577, fit: 'resize' });
  badge(s, 0.667, 3.875);
  imageBox(s, 8.721, 0.542, 3.932, 5.415);
  chrome(s);
}

/* Slide 13 — Our Service */
function slide13(s) {
  tx(s, 'Our Service', { fontSize: 60, fontFace: SEMI, x: 0.677, y: 0.786, w: 4.957, h: 1.01, fit: 'resize' });
  body(s, COPY4, 0.677, 2.032, 5.545, 0.869);
  body(s, COPY1, 7.111, 2.032, 5.556, 0.869);
  body(s, COPY2, 7.111, 3.391, 5.012, 0.566);
  learnMore(s, 7.114, 4.51);
  rule(s, 7.111, 3.168, 5.556, 0, GREY35, 1);
  box(s, 12.667, 0, 0.667, 1.543, GREEN);
  imageBox(s, 0.677, 3.391, 5.668, 2.347);
  chrome(s);
}

/* Slide 14 — Our Great Service — two cards */
function slide14(s) {
  tx(s, 'What We Offer', { fontFace: MED, x: 0.677, y: 0.991, w: 1.79, h: 0.303, wrap: false, fit: 'resize' });
  head(s, 'Our Great Service', 0.677, 1.359, 5.196, 0.741, 44);
  body(s, COPY5, 0.677, 2.516, 5.533, 0.869);
  badge(s, 0.667, 3.875);
  body(s, COPY2, 0.677, 4.987, 5.012, 0.566);
  box(s, 7.233, 1.453, 2.562, 4.1, GREY05);
  box(s, 10.094, 1.453, 2.562, 4.1, GREEN);
  tx(s, 'Technical Service', { fontFace: SEMI, align: 'center', x: 7.48, y: 3.883, w: 2.069, h: 0.303, wrap: false, fit: 'resize' });
  tx(s, COPY6, {
    fontSize: 11,
    fontFace: REG,
    align: 'center',
    lineSpacingMultiple: 1.5,
    x: 7.48,
    y: 4.375,
    w: 2.069,
    h: 0.797,
    fit: 'resize'
  });
  tx(s, 'Global Expersite', { fontFace: SEMI, color: WHITE, align: 'center', x: 10.407, y: 3.883, w: 1.937, h: 0.303, wrap: false, fit: 'resize' });
  tx(s, COPY6, {
    fontSize: 11,
    fontFace: REG,
    color: WHITE,
    align: 'center',
    lineSpacingMultiple: 1.5,
    x: 10.344,
    y: 4.375,
    w: 2.069,
    h: 0.797,
    fit: 'resize'
  });
  imageBox(s, 7.725, 1.828, 1.578, 1.578);
  imageBox(s, 10.586, 1.828, 1.578, 1.578);
  chrome(s);
}

/* Slide 15 — Special Corporate Solution */
function slide15(s) {
  head(s, 'Sepecial Corporate', 0.677, 0.634, 5.992, 0.808, 48);
  head(s, 'Solution', 0.677, 1.442, 2.603, 0.808, 48);
  body(s, COPY4, 0.677, 2.57, 5.545, 0.869);
  box(s, 7.552, 1.133, 5.115, 4.853, GREEN);
  box(s, 8.181, 1.468, 2.223, 0.452, WHITE);
  tx(s, 'Enhanced Service', { fontSize: 14, fontFace: MED, x: 8.481, y: 1.576, w: 1.623, h: 0.236, wrap: false, fit: 'resize' });
  body(s, COPY7, 8.181, 2.01, 3.945, 0.566, WHITE);
  box(s, 8.181, 3.005, 2.223, 0.452, WHITE);
  tx(s, 'Technical Service', { fontSize: 14, fontFace: MED, x: 8.499, y: 3.114, w: 1.588, h: 0.236, wrap: false, fit: 'resize' });
  body(s, COPY7, 8.181, 3.548, 3.945, 0.566, WHITE);
  box(s, 8.181, 4.543, 2.223, 0.452, WHITE);
  tx(s, 'Solar Panel Service', { fontSize: 14, fontFace: MED, x: 8.415, y: 4.651, w: 1.755, h: 0.236, wrap: false, fit: 'resize' });
  body(s, COPY7, 8.181, 5.086, 3.945, 0.566, WHITE);
  imageBox(s, 0.677, 3.759, 5.545, 2.228);
  chrome(s);
}

/* Slide 16 — Best Solution For Your Home — three panels */
function slide16(s) {
  head(s, 'Best Solution For', 0.61, 0.542, 6.116, 0.909, 54);
  head(s, 'Your Home', 0.61, 1.327, 4.002, 0.909, 54);
  body(s, COPY3, 7.656, 1.367, 4.934, 0.869);
  box(s, 0.677, 3.232, 3.868, 2.478, GREEN);
  box(s, 4.733, 3.232, 3.868, 2.478, GREY05);
  box(s, 8.789, 3.232, 3.868, 2.478, GREEN);
  tx(s, 'Wind Turbine Service', {
    fontFace: MED,
    color: WHITE,
    bullet: { characterCode: '2022', indent: 22.5 },
    x: 1.198,
    y: 3.578,
    w: 2.826,
    h: 0.303,
    wrap: false,
    fit: 'resize'
  });
  body(s, COPY8, 1.198, 4.056, 2.826, 0.566, WHITE);
  tx(s, [
    { text: 'Dignissim diam quis enim lobort. ', options: {
      fontSize: 12,
      fontFace: REG,
      color: WHITE,
      lineSpacingMultiple: 1.5,
      bullet: { characterCode: '2022', indent: 13.5 },
      breakLine: true
    } },
    { text: 'Tincidunt tortor aliquam nulla.', options: { fontSize: 12, fontFace: REG, color: WHITE, lineSpacingMultiple: 1.5, bullet: { characterCode: '2022', indent: 13.5 } } }
  ], { x: 1.204, y: 4.797, w: 2.702, h: 0.566, fit: 'resize' });
  tx(s, 'Technical Service', {
    fontFace: MED,
    bullet: { characterCode: '2022', indent: 22.5 },
    x: 5.254,
    y: 3.578,
    w: 2.351,
    h: 0.303,
    wrap: false,
    fit: 'resize'
  });
  body(s, COPY8, 5.254, 4.056, 2.826, 0.566);
  tx(s, [
    { text: 'Dignissim diam quis enim lobort. ', options: {
      fontSize: 12,
      fontFace: REG,
      lineSpacingMultiple: 1.5,
      bullet: { characterCode: '2022', indent: 13.5 },
      breakLine: true
    } },
    { text: 'Tincidunt tortor aliquam nulla.', options: { fontSize: 12, fontFace: REG, lineSpacingMultiple: 1.5, bullet: { characterCode: '2022', indent: 13.5 } } }
  ], { x: 5.26, y: 4.797, w: 2.702, h: 0.566, fit: 'resize' });
  tx(s, 'Solar Panel Service', {
    fontFace: MED,
    color: WHITE,
    bullet: { characterCode: '2022', indent: 22.5 },
    x: 9.303,
    y: 3.578,
    w: 2.565,
    h: 0.303,
    wrap: false,
    fit: 'resize'
  });
  body(s, COPY8, 9.303, 4.056, 2.826, 0.566, WHITE);
  tx(s, [
    { text: 'Dignissim diam quis enim lobort. ', options: {
      fontSize: 12,
      fontFace: REG,
      color: WHITE,
      lineSpacingMultiple: 1.5,
      bullet: { characterCode: '2022', indent: 13.5 },
      breakLine: true
    } },
    { text: 'Tincidunt tortor aliquam nulla.', options: { fontSize: 12, fontFace: REG, color: WHITE, lineSpacingMultiple: 1.5, bullet: { characterCode: '2022', indent: 13.5 } } }
  ], { x: 9.31, y: 4.797, w: 2.702, h: 0.566, fit: 'resize' });
  chrome(s);
}

/* Slide 17 — Greener Tomorrow Everyone */
function slide17(s) {
  head(s, 'Greener Tomorrow', 0.677, 0.724, 6.793, 0.909, 54);
  head(s, 'Everyone', 0.687, 1.549, 3.396, 0.909, 54);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eiusmod tempor incididunt ut labore et dolore magna aliqua. Dignissim diam quis enim lobortis. Tincidunt tortor aliquam.', 8, 1.549, 4.657, 0.869);
  imageBox(s, 8, 3.167, 4.667, 2.828);
  imageBox(s, 0.687, 3.167, 7.18, 2.828);
  chrome(s);
}

/* Slide 18 — Wind Energy */
function slide18(s) {
  head(s, 'Wind Energy', 0.677, 1.035, 4.598, 0.909, 54);
  body(s, COPY1, 0.677, 2.402, 5.601, 0.869);
  body(s, COPY2, 0.677, 3.6, 5.012, 0.566);
  learnMore(s, 0.677, 4.806);
  box(s, 8.108, 4.312, 4.545, 1.49, GREEN);
  head(s, '82%', 8.475, 4.54, 1.534, 0.909, 54, WHITE);
  tx(s, 'Energy', { fontFace: MED, color: WHITE, x: 10.422, y: 4.62, w: 0.822, h: 0.303, wrap: false, fit: 'resize' });
  tx(s, 'Laoreet suspendisse interdum consectetur libo.', {
    fontSize: 11,
    fontFace: REG,
    color: WHITE,
    lineSpacingMultiple: 1.5,
    x: 10.417,
    y: 4.975,
    w: 1.87,
    h: 0.519,
    fit: 'resize'
  });
  imageBox(s, 8.111, 0, 4.545, 4.166);
  chrome(s);
}

/* Slide 19 — Solar Energy */
function slide19(s) {
  head(s, 'Solar Energy', 0.677, 1.035, 4.549, 0.909, 54);
  body(s, COPY1, 0.677, 2.402, 5.601, 0.869);
  body(s, COPY2, 0.677, 3.6, 5.012, 0.566);
  learnMore(s, 0.677, 4.806);
  box(s, 8.108, 0.544, 4.545, 1.49, GREEN);
  head(s, '88%', 8.475, 0.772, 1.546, 0.909, 54, WHITE);
  tx(s, 'Energy', { fontFace: MED, color: WHITE, x: 10.422, y: 0.852, w: 0.822, h: 0.303, wrap: false, fit: 'resize' });
  tx(s, 'Laoreet suspendisse interdum consectetur libo.', {
    fontSize: 11,
    fontFace: REG,
    color: WHITE,
    lineSpacingMultiple: 1.5,
    x: 10.417,
    y: 1.207,
    w: 1.87,
    h: 0.519,
    fit: 'resize'
  });
  imageBox(s, 8.111, 2.156, 4.545, 3.511);
  chrome(s);
}

/* Slide 20 — Service Impact — concentric rings */
function slide20(s) {
  box(s, 0.677, 4.183, 3.705, 0.903, GREEN);
  tx(s, '$ 4.000.000', {
    fontSize: 40,
    fontFace: SEMI,
    color: WHITE,
    align: 'center',
    x: 0.858,
    y: 4.298,
    w: 3.343,
    h: 0.673,
    wrap: false,
    fit: 'resize'
  });
  head(s, 'Service Impact', 0.677, 0.883, 5.31, 0.909, 54);
  body(s, 'Vel turpis nunc eget lorem dolor sed viverra ipsum. Laoreet suspendisse interdum consectetur libero id faucibus nisl tincidunt. Arcu risus quis varius quam quisque id diam vel. ', 0.677, 2.192, 4.801, 0.88);
  rule(s, 0.677, 3.639, 5.283, 0, NEARBLACK, 0.5);
  s.addShape('ellipse', { x: 7.969, y: 0.783, w: 4.712, h: 4.712, fill: { color: GREY05 } });
  s.addShape('ellipse', { x: 8.903, y: 1.717, w: 2.844, h: 2.844, fill: { color: GREY25, transparency: 74 } });
  s.addShape('ellipse', { x: 9.437, y: 2.251, w: 1.777, h: 1.777, fill: { color: GREEN } });
  tx(s, '80%', {
    fontSize: 40,
    fontFace: SEMI,
    color: WHITE,
    align: 'center',
    x: 9.741,
    y: 2.803,
    w: 1.169,
    h: 0.673,
    wrap: false,
    fit: 'resize'
  });
  chrome(s);
}

/* Slide 21 — Company Chart — grouped column chart */
function slide21(s) {
  tx(s, 'Company Chart', { fontSize: 54, fontFace: SEMI, color: BLACK, x: 0.667, y: 0.978, w: 6.267, h: 0.909, fit: 'resize' });
  body(s, 'Vel turpis nunc eget lorem dolor sed viverra ipsum. Laoreet suspendisse interdum consectetur libero id faucibus nisl tincidunt. Arcu risus quis varius.', 0.667, 2.612, 3.989, 0.88);
  box(s, 0.667, 4.341, 1.199, 0.489, GREEN);
  box(s, 0.667, 5.109, 1.199, 0.489, GREEN);
  tx(s, '78%', { fontFace: MED, color: WHITE, align: 'center', x: 1.014, y: 4.434, w: 0.503, h: 0.303, wrap: false, fit: 'resize' });
  tx(s, '45%', { fontFace: MED, color: WHITE, align: 'center', x: 1.018, y: 5.202, w: 0.496, h: 0.303, wrap: false, fit: 'resize' });
  tx(s, 'Laoreet suspendisse interdum.', {
    fontSize: 12,
    fontFace: REG,
    bold: true,
    lineSpacingMultiple: 1.5,
    bullet: { characterCode: '2022', indent: 13.5 },
    x: 2.07,
    y: 4.448,
    w: 2.74,
    h: 0.263,
    fit: 'resize'
  });
  tx(s, 'faucibus nisl tincidunt arcu ris.', {
    fontSize: 12,
    fontFace: REG,
    bold: true,
    lineSpacingMultiple: 1.5,
    bullet: { characterCode: '2022', indent: 13.5 },
    x: 2.07,
    y: 5.216,
    w: 2.74,
    h: 0.263,
    fit: 'resize'
  });
  box(s, 8.276, 3.739, 0.645, 1.74, GREY15);
  box(s, 7.605, 3.95, 0.645, 1.53, GREY25);
  box(s, 6.933, 3.364, 0.645, 2.116, GREEN);
  box(s, 11.47, 3.95, 0.645, 1.53, GREY15);
  box(s, 10.798, 3.502, 0.645, 1.977, GREY25);
  box(s, 10.127, 4.243, 0.645, 1.237, GREEN);
  rule(s, 6.444, 3.368, 0, 2.695, BLACK, 1);
  tx(s, '001', {
    fontFace: MED,
    align: 'center',
    bullet: { characterCode: '2022', indent: 22.5 },
    x: 7.408,
    y: 5.663,
    w: 0.761,
    h: 0.303,
    wrap: false,
    fit: 'resize'
  });
  tx(s, '002', {
    fontFace: MED,
    align: 'center',
    bullet: { characterCode: '2022', indent: 22.5 },
    x: 10.598,
    y: 5.663,
    w: 0.794,
    h: 0.303,
    wrap: false,
    fit: 'resize'
  });
  rule(s, 6.046, 5.478, 6.627, 0, BLACK, 1);
  body(s, COPY3, 7.739, 1.522, 4.934, 0.869);
  chrome(s);
}

/* Slide 22 — Data Chart — stacked bar chart */
function slide22(s) {
  rule(s, 2.034, 1.143, 0, 4.669, GREY15, 0.25);
  rule(s, 2.845, 1.143, 0, 4.669, GREY15, 0.25);
  rule(s, 3.655, 1.143, 0, 4.669, GREY15, 0.25);
  rule(s, 4.466, 1.143, 0, 4.669, GREY15, 0.25);
  rule(s, 5.276, 1.143, 0, 4.669, GREY15, 0.25);
  rule(s, 6.087, 1.143, 0, 4.669, GREY15, 0.25);
  rule(s, 6.897, 1.143, 0, 4.669, GREY15, 0.25);
  box(s, 4.995, 1.487, 2.041, 0.483, SILVER);
  box(s, 4.216, 2.666, 1.179, 0.483, SILVER);
  box(s, 4.842, 3.844, 0.826, 0.483, SILVER);
  box(s, 4.747, 5.022, 0.802, 0.483, SILVER);
  box(s, 3.874, 1.487, 1.121, 0.483, GREY15);
  box(s, 3.497, 2.666, 0.72, 0.483, GREY15);
  box(s, 3.061, 3.844, 1.781, 0.483, GREY15);
  box(s, 3.792, 5.022, 0.955, 0.483, GREY15);
  box(s, 2.034, 1.487, 1.84, 0.483, GREEN);
  box(s, 2.034, 2.666, 1.462, 0.483, GREEN);
  box(s, 2.034, 3.844, 1.026, 0.483, GREEN);
  box(s, 2.034, 5.022, 1.757, 0.483, GREEN);
  tx(s, 'Category 01', { fontSize: 9, fontFace: REG, x: 0.686, y: 5.184, w: 1.186, h: 0.151, fit: 'resize' });
  tx(s, 'Category 02', { fontSize: 9, fontFace: REG, x: 0.686, y: 4.006, w: 1.186, h: 0.151, fit: 'resize' });
  tx(s, 'Category 03', { fontSize: 9, fontFace: REG, x: 0.677, y: 2.828, w: 1.205, h: 0.151, fit: 'resize' });
  tx(s, 'Category 04', { fontSize: 9, fontFace: REG, x: 0.677, y: 1.644, w: 1.186, h: 0.151, fit: 'resize' });
  tx(s, 'Data Chart', { fontSize: 54, fontFace: SEMI, color: BLACK, x: 8.215, y: 1.078, w: 3.992, h: 0.909, fit: 'resize' });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed eiusmod tempor incididunt ut labore et dolore magna aliqua. Dignissim diam quis enim lobortis tincidunt tortor.', 8.215, 2.231, 4.442, 0.869);
  badge(s, 8.215, 3.533);
  box(s, 8.22, 4.458, 2.48, 1.355, GREY05);
  tx(s, '82%', { fontSize: 54, fontFace: SEMI, align: 'center', x: 8.452, y: 4.809, w: 2.017, h: 0.909, fit: 'resize' });
  tx(s, 'Data Average', { fontSize: 14, fontFace: MED, align: 'center', x: 8.833, y: 4.619, w: 1.253, h: 0.236, wrap: false, fit: 'resize' });
  chrome(s);
}

/* Slide 23 — Pricing table */
function slide23(s) {
  box(s, 0.927, 0.979, 3.188, 4.694, GREY05);
  box(s, 0.927, 0.979, 3.188, 0.802, GREEN);
  box(s, 5.073, 0.979, 3.188, 4.694, GREY05);
  box(s, 5.073, 0.979, 3.188, 0.802, GREEN);
  box(s, 9.219, 0.977, 3.188, 4.694, GREY05);
  box(s, 9.219, 0.977, 3.188, 0.802, GREEN);
  tx(s, 'Standard', {
    fontSize: 28,
    fontFace: SEMI,
    color: WHITE,
    align: 'center',
    x: 1.674,
    y: 1.148,
    w: 1.694,
    h: 0.471,
    wrap: false,
    fit: 'resize'
  });
  tx(s, 'Popular', {
    fontSize: 28,
    fontFace: SEMI,
    color: WHITE,
    align: 'center',
    x: 5.945,
    y: 1.143,
    w: 1.443,
    h: 0.471,
    wrap: false,
    fit: 'resize'
  });
  tx(s, 'Premium', {
    fontSize: 28,
    fontFace: SEMI,
    color: WHITE,
    align: 'center',
    x: 9.967,
    y: 1.143,
    w: 1.69,
    h: 0.471,
    wrap: false,
    fit: 'resize'
  });
  tx(s, [
    { text: '$471', options: { fontSize: 44, fontFace: SEMI, align: 'center' } },
    { text: '/Monthly', options: { fontSize: 11, fontFace: SEMI, align: 'center' } }
  ], { x: 5.648, y: 2.155, w: 2.037, h: 0.741, wrap: false, fit: 'resize' });
  tx(s, [
    { text: '$821', options: { fontSize: 44, fontFace: SEMI, align: 'center' } },
    { text: '/Monthly', options: { fontSize: 11, fontFace: SEMI, align: 'center' } }
  ], { x: 9.786, y: 2.155, w: 2.053, h: 0.741, wrap: false, fit: 'resize' });
  tx(s, [
    { text: '$230', options: { fontSize: 44, fontFace: SEMI, align: 'center' } },
    { text: '/Monthly', options: { fontSize: 11, fontFace: SEMI, align: 'center' } }
  ], { x: 1.442, y: 2.155, w: 2.158, h: 0.741, wrap: false, fit: 'resize' });
  tx(s, [
    { text: 'Installation', options: { fontSize: 12, fontFace: REG, align: 'center', lineSpacingMultiple: 1.5, breakLine: true } },
    { text: 'Repair & Replacement', options: { fontSize: 12, fontFace: REG, align: 'center', lineSpacingMultiple: 1.5 } }
  ], { x: 1.669, y: 3.488, w: 1.705, h: 0.566, fit: 'resize' });
  box(s, 1.534, 4.757, 1.974, 0.583, GREY25);
  head(s, 'Get Started', 1.828, 4.897, 1.385, 0.303, 18);
  tx(s, [
    { text: 'Installation', options: { fontSize: 12, fontFace: REG, align: 'center', lineSpacingMultiple: 1.5, breakLine: true } },
    { text: 'Repair & Replacement', options: { fontSize: 12, fontFace: REG, align: 'center', lineSpacingMultiple: 1.5, breakLine: true } },
    { text: 'Monitoring Work', options: { fontSize: 12, fontFace: REG, align: 'center', lineSpacingMultiple: 1.5 } }
  ], { x: 5.814, y: 3.336, w: 1.705, h: 0.869, fit: 'resize' });
  box(s, 5.679, 4.757, 1.974, 0.583, GREY25);
  head(s, 'Get Started', 5.974, 4.897, 1.385, 0.303, 18);
  tx(s, [
    { text: 'Installation', options: { fontSize: 12, fontFace: REG, align: 'center', lineSpacingMultiple: 1.5, breakLine: true } },
    { text: 'Repair & Replacement', options: { fontSize: 12, fontFace: REG, align: 'center', lineSpacingMultiple: 1.5, breakLine: true } },
    { text: 'Monitoring Work', options: { fontSize: 12, fontFace: REG, align: 'center', lineSpacingMultiple: 1.5, breakLine: true } },
    { text: 'Panel Maintenance', options: { fontSize: 12, fontFace: REG, align: 'center', lineSpacingMultiple: 1.5 } }
  ], { x: 9.96, y: 3.165, w: 1.705, h: 1.212, fit: 'resize' });
  box(s, 9.825, 4.757, 1.974, 0.583, GREY25);
  head(s, 'Get Started', 10.12, 4.897, 1.385, 0.303, 18);
  chrome(s);
}

/* Slide 24 — Contact Us */
function slide24(s) {
  head(s, 'Contact Us', 0.677, 0.63, 4.763, 1.111, 66);
  body(s, COPY3, 0.677, 2.027, 4.934, 0.869);
  box(s, 7.383, 3.496, 5.283, 2.578, GREEN);
  tx(s, 'Phone      :', { fontSize: 24, fontFace: BARLOW, bold: true, color: WHITE, x: 7.969, y: 3.858, w: 1.464, h: 0.404, fit: 'resize' });
  tx(s, '(123) 456 - 7890', {
    fontSize: 14,
    fontFace: BARLOW,
    color: WHITE,
    lineSpacingMultiple: 1.5,
    x: 10.004,
    y: 3.934,
    w: 1.611,
    h: 0.304,
    fit: 'resize'
  });
  tx(s, 'E Mail       :', { fontSize: 24, fontFace: BARLOW, bold: true, color: WHITE, x: 7.969, y: 4.511, w: 1.425, h: 0.404, fit: 'resize' });
  tx(s, 'Yourname@mail.com', {
    fontSize: 14,
    fontFace: BARLOW,
    color: WHITE,
    lineSpacingMultiple: 1.5,
    x: 10.004,
    y: 4.587,
    w: 2.077,
    h: 0.304,
    fit: 'resize'
  });
  tx(s, 'Location  :', { fontSize: 24, fontFace: BARLOW, bold: true, color: WHITE, x: 7.969, y: 5.164, w: 1.579, h: 0.404, fit: 'resize' });
  tx(s, 'Amphitrea 21. Mountas New York', { fontSize: 14, fontFace: BARLOW, color: WHITE, x: 10.004, y: 5.24, w: 2.077, h: 0.471, fit: 'resize' });
  body(s, 'Eiusmod tempor incididunt ut labore et dolore magna aliqua. Dignissim diam quis enim lobortis. Tincidunt tortor aliquam nulla.', 7.656, 2.028, 4.934, 0.566);
  imageBox(s, 0.677, 3.496, 6.476, 2.578);
  chrome(s);
}

/* Slide 25 — Thank You! */
function slide25(s) {
  head(s, 'Thank You!', 7.295, 0.542, 5.498, 1.212, 72);
  tx(s, 'Solar & Renewable Energy Presentation', { fontFace: MED, color: INK, x: 7.295, y: 1.74, w: 4.658, h: 0.303, wrap: false, fit: 'resize' });
  body(s, COPY3, 7.295, 2.728, 4.934, 0.869);
  body(s, COPY2, 0.677, 4.155, 5.012, 0.566);
  badge(s, 0.667, 5.28);
  box(s, 7.295, 4.155, 5.372, 1.747, GREEN);
  tx(s, [
    { text: 'Take Care Of Nature ', options: { fontSize: 32, fontFace: MED, color: WHITE, breakLine: true } },
    { text: 'For A Healthier Life', options: { fontSize: 32, fontFace: MED, color: WHITE } }
  ], { x: 7.771, y: 4.49, w: 4.42, h: 1.077, wrap: false, fit: 'resize' });
  imageBox(s, 0.677, 0.542, 5.99, 3.056);
  chrome(s);
}


const SLIDES = [
  slide01,
  slide02,
  slide03,
  slide04,
  slide05,
  slide06,
  slide07,
  slide08,
  slide09,
  slide10,
  slide11,
  slide12,
  slide13,
  slide14,
  slide15,
  slide16,
  slide17,
  slide18,
  slide19,
  slide20,
  slide21,
  slide22,
  slide23,
  slide24,
  slide25,
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'CULKA', width: 13.333, height: 7.5 });
  pptx.layout = 'CULKA';
  pptx.author = 'Culka Company';
  pptx.title = 'Culka Business Presentation Template';

  SLIDES.forEach(fn => {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    fn(s);
  });

  return pptx.writeFile({ fileName: path.join(__dirname, '0ca89487-79d9-4745-928c-5592fee7451e_grok_final.pptx') });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
