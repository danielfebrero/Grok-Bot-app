/**
 * Momiar — Real Estate Presentation Template (30 slides, 10 x 5.625in).
 * Standalone pptxgenjs re-creation of the reference deck.
 *
 * Photographs in the original are replaced by flat colour placeholders that keep
 * the same position, size and average tone; the device mock-ups are rebuilt from
 * native shapes (chassis + screen block).
 *
 * Run: node <thisfile>.js   ->   writes <thisfile>.pptx next to the script.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const C = {
  paper: 'FFFEFD', // page wash used on nearly every slide
  panel: 'EAE9E4', // warm grey blocks behind content
  ink: '262626',
  black: '000000',
  white: 'FFFFFF',
  navy: '152431',
  slate: '8092A2', // accent used for rings and the quote bar
  steel: '7F7F7F', // icon / rule grey
  ghost: 'F2F2F2', // oversized step numerals
  muted: '374151',
  photoDark: '3C3C3C', // stand-in tone for the dark photography
  photoLight: 'D9D7D8', // stand-in tone for the light photography
  device: 'C9C9C9', // mock-up aluminium body
  bezel: '2E2E2E', // mock-up screen bezel
};

/* -------------------------------------------------------------------- fonts */
const F = { head: 'Poppins', semi: 'Poppins SemiBold', body: 'Open Sans', mark: 'Montserrat' };

/* Text insets of the source deck, in points (pptxgenjs converts pt -> EMU).
 * Kept at the exact source EMU values: several headlines fill their box to
 * within 0.002in, so rounding the inset up would push them onto a second line. */
const EMU_PER_PT = 12700;
const INSET = [68575 / EMU_PER_PT, 68575 / EMU_PER_PT, 34275 / EMU_PER_PT, 34275 / EMU_PER_PT];

/* ------------------------------------------------------------------- styles */
const S = {
  body:    { fontFace: F.body, fontSize: 8,  color: C.black, lineSpacingMultiple: 1.5 },
  label:   { fontFace: F.body, fontSize: 11, color: C.black, lineSpacingMultiple: 1.5 },
  kicker:  { fontFace: F.body, fontSize: 11, color: C.ink },
  h1:      { fontFace: F.head, fontSize: 30, color: C.ink },
  h2:      { fontFace: F.head, fontSize: 24, color: C.black },
  hero:    { fontFace: F.semi, fontSize: 50, color: C.ink },
  logo:    { fontFace: F.semi, fontSize: 14, color: C.ink },
  mark:    { fontFace: F.mark, fontSize: 8,  color: C.ink },
  numeral: { fontFace: F.head, fontSize: 50, color: C.ghost },
  stat:    { fontFace: F.head, fontSize: 21, color: C.ink },
};

/* ------------------------------------------------------------------ helpers */
/* Every helper takes a geometry tuple g = [x, y, w, h] in inches. */
const box = (g) => ({ x: g[0], y: g[1], w: g[2], h: g[3] });

function rect(s, g, color, transparency) {
  s.addShape('rect', { ...box(g), fill: { color, transparency: transparency || 0 }, line: { type: 'none' } });
}

function text(s, g, content, opts) {
  s.addText(content, {
    ...box(g), align: 'left', valign: 'top', wrap: true, margin: INSET,
    fontFace: F.body, fontSize: 8, color: C.black, ...opts,
  });
}

/* Raster imagery of the reference deck -> flat tonal block; the tone matches the
 * average of the photo it stands in for. `round` mirrors an ellipse-cropped pic. */
function photo(s, g, kind, opts = {}) {
  s.addShape(opts.round ? 'ellipse' : 'rect', {
    ...box(g), fill: { color: kind === 'light' ? C.photoLight : C.photoDark }, line: { type: 'none' },
    ...(opts.rot ? { rotate: opts.rot } : {}),
  });
}

/* Device mock-ups: the chassis is rebuilt from flat rectangles traced off the
 * silhouette of the original artwork; the matching screen photo is drawn on top
 * by the slide builder, so only the surrounding body is described here.
 * Parts are [x, y, w, h, tone, cornerRadius?] in 0..1 space of the art box. */
const BEZEL = 'bezel'; // dark lid / bezel of the device
const SHELL = 'shell'; // light aluminium base, chin and stand
const CHASSIS_PARTS = {
  monitor: [
    [0.020, 0.019, 0.961, 0.674, BEZEL],
    [0.015, 0.693, 0.966, 0.108, SHELL], // chin below the glass
    [0.397, 0.801, 0.202, 0.178, SHELL], // stand
  ],
  laptop: [
    [0.093, 0.000, 0.814, 0.940, BEZEL], // lid
    [0.016, 0.940, 0.967, 0.060, SHELL], // keyboard deck
  ],
  phone: [
    [0.008, 0.000, 0.982, 1.000, BEZEL, 0.09], // rounded unibody
  ],
};

function chassis(s, g, kind, opts = {}) {
  const [x, y, w, h] = g;
  const deg = opts.rot || 0;
  const rad = (deg * Math.PI) / 180;
  const cx = x + w / 2;
  const cy = y + h / 2;
  CHASSIS_PARTS[kind].forEach(([a, b, aw, ah, tone, round]) => {
    const pw = aw * w;
    const ph = ah * h;
    // Swing each part about the chassis centre so the whole frame turns as one.
    const dx = x + a * w + pw / 2 - cx;
    const dy = y + b * h + ph / 2 - cy;
    s.addShape(round ? 'roundRect' : 'rect', {
      x: cx + dx * Math.cos(rad) - dy * Math.sin(rad) - pw / 2,
      y: cy + dx * Math.sin(rad) + dy * Math.cos(rad) - ph / 2,
      w: pw, h: ph,
      fill: { color: tone === BEZEL ? C.bezel : C.device },
      line: { type: 'none' },
      ...(round ? { rectRadius: round * Math.min(pw, ph) } : {}),
      ...(deg ? { rotate: deg } : {}),
    });
  });
}

/* Thin slate ring (donut). rectRadius sets the donut wall thickness in inches. */
function ring(s, g) {
  s.addShape('donut', { ...box(g), fill: { color: C.slate }, line: { type: 'none' }, rectRadius: 0.031 });
}

/* Short arrow rule that marks the top of most content slides. */
function tick(s, g, flipH) {
  s.addShape('rightArrow', { ...box(g), fill: { color: C.steel }, line: { type: 'none' }, flipH: !!flipH });
}

/* Tiny line-art glyphs that sit inside the rings, drawn from primitives so the
 * deck stays self-contained. Coordinates are 0..1 inside the icon box g. */
function icon(s, kind, g) {
  const [x, y, w, h] = g;
  const hair = { color: C.steel, width: 0.75 };
  const frame = (a, b, aw, ah) =>
    s.addShape('rect', { x: x + a * w, y: y + b * h, w: aw * w, h: ah * h, fill: { type: 'none' }, line: hair });
  const solid = (a, b, aw, ah) =>
    s.addShape('rect', { x: x + a * w, y: y + b * h, w: aw * w, h: ah * h, fill: { color: C.steel }, line: { type: 'none' } });
  const seg = (a1, b1, a2, b2) =>
    s.addShape('line', {
      x: x + Math.min(a1, a2) * w, y: y + Math.min(b1, b2) * h,
      w: Math.abs(a2 - a1) * w, h: Math.abs(b2 - b1) * h,
      line: hair, flipV: b2 < b1,
    });

  if (kind === 'building') {
    frame(0.04, 0.0, 0.44, 1.0); // tower
    frame(0.48, 0.36, 0.48, 0.64); // low wing
    [0.16, 0.36, 0.56, 0.76].forEach((b) => solid(0.13, b, 0.26, 0.07)); // windows
    solid(0.6, 0.62, 0.24, 0.2);
  } else if (kind === 'growth') {
    seg(0.02, 0.86, 0.34, 0.44); // rising trend line
    seg(0.34, 0.44, 0.55, 0.62);
    seg(0.55, 0.62, 0.94, 0.14);
    seg(0.72, 0.14, 0.94, 0.14); // arrow head
    seg(0.94, 0.14, 0.94, 0.36);
    seg(0.02, 0.86, 0.94, 0.86); // baseline
    s.addShape('ellipse', { x: x + 0.6 * w, y: y + 0.58 * h, w: 0.36 * w, h: 0.36 * h, fill: { type: 'none' }, line: hair });
  } else {
    frame(0.02, 0.5, 0.2, 0.5); // three bars
    frame(0.3, 0.34, 0.2, 0.66);
    frame(0.58, 0.56, 0.2, 0.44);
    seg(0.06, 0.26, 0.38, 0.1); // sparkline over the bars
    seg(0.38, 0.1, 0.66, 0.3);
    seg(0.66, 0.3, 0.96, 0.04);
    frame(0.8, 0.42, 0.2, 0.58);
  }
}

/* Slide 29 draws its analytics graphic as vector art (no embedded chart part):
 * five gridlines, an L-shaped axis, a polyline and four pill markers. */
function lineChartArt(s) {
  const PLOT = { x: 1.041, y: 1.514, w: 2.704, h: 1.887 }; // 100 at top, 20 at bottom
  const yFor = (v) => PLOT.y + ((100 - v) / 80) * PLOT.h;
  [100, 80, 60, 40, 20].forEach((v) => {
    rect(s, [PLOT.x, yFor(v), PLOT.w, 0.008], C.ink);
    rect(s, [PLOT.x - 0.088, yFor(v), 0.085, 0.008], C.ink);
    text(s, [PLOT.x - 0.62, yFor(v) - 0.09, 0.48, 0.2], String(v),
      { fontSize: 6, color: C.ink, align: 'right' });
  });
  rect(s, [1.034, 1.415, 0.008, 2.323], C.ink); // y axis
  rect(s, [1.034, 3.73, 2.712, 0.008], C.ink); // x axis

  const SERIES = [[1.049, 3.733], [1.538, 2.469], [2.114, 2.932], [2.692, 2.001], [3.248, 3.406], [3.749, 3.736]];
  for (let i = 0; i < SERIES.length - 1; i++) {
    const [x1, y1] = SERIES[i];
    const [x2, y2] = SERIES[i + 1];
    s.addShape('line', {
      x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
      line: { color: C.ink, width: 1 }, flipV: y2 < y1,
    });
  }
  SERIES.slice(1, 5).forEach(([cx, cy]) => {
    s.addShape('roundRect', {
      x: cx - 0.11, y: cy - 0.19, w: 0.22, h: 0.38,
      fill: { color: C.ink }, line: { type: 'none' }, rectRadius: 0.11,
    });
  });
}

/* --------------------------------------------------------------- the slides */

/* 1. Cover */
function slide01(s) {
  rect(s, [5.737, 0, 4.263, 5.625], C.paper);
  photo(s, [0, 0, 5.737, 5.625], 'dark');
  text(s, [4.954, 1.814, 3.07, 0.909], "Momiar.", { ...S.hero });
  text(s, [4.974, 2.775, 3.732, 0.252], "Real Estate Presentation Template", { ...S.kicker });
  text(s, [4.974, 3.256, 4.326, 0.622], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  text(s, [8.438, 5.052, 1.057, 0.215], "IndoartCreative", { ...S.mark, align: 'right' });
  text(s, [8.353, 0.414, 1.285, 0.311], "Momiar.", { ...S.logo });
}

/* 2. Welcome slide — image + card */
function slide02(s) {
  rect(s, [0, 0, 10, 5.625], C.paper, 68.24);
  photo(s, [1.104, 1.315, 3.134, 3.073], 'dark');
  rect(s, [3.406, 3.333, 3.177, 1.51], C.panel, 65.1);
  text(s, [4.789, 1.576, 3.871, 0.581], [{ text: "Welcome Slide", options: { bold: true } }], { ...S.h1 });
  text(s, [3.713, 3.767, 2.563, 0.644], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et", { ...S.body, color: C.ink });
  text(s, [4.789, 2.37, 4.388, 0.53], "We create impactful, eye-catching campaigns that are tailored to meet your unique needs", { ...S.kicker, fontSize: 14, color: C.muted });
}

/* 3. Welcome slide — side panel */
function slide03(s) {
  rect(s, [0, 0, 10, 5.625], C.paper, 68.24);
  rect(s, [6.869, 0.393, 2.123, 4.582], C.panel, 65.1);
  photo(s, [6.229, 1.781, 2.497, 2.448], 'dark');
  photo(s, [0.975, 3.306, 0.942, 0.923], 'dark', { round: true });
  text(s, [0.975, 1.367, 3.848, 0.581], [{ text: "Welcome Slide", options: { bold: true } }], { ...S.h1 });
  text(s, [0.975, 2.074, 4.275, 0.701], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  text(s, [2.449, 3.601, 2.801, 0.701], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna", { ...S.body });
  text(s, [2.449, 3.173, 2.374, 0.311], [{ text: "Lorem ipsum dolor sit amet", options: { bold: true } }], { ...S.label });
}

/* 4. Welcome message — slate quote bar */
function slide04(s) {
  rect(s, [0, 0, 10, 5.625], C.paper, 68.24);
  rect(s, [6.516, 0.641, 3.174, 2.69], C.panel, 65.1);
  photo(s, [1.047, 3.732, 0.751, 0.736], 'dark', { round: true });
  photo(s, [7.193, 2.876, 2.497, 2.448], 'dark');
  rect(s, [0.947, 3.588, 3.679, 0.976], C.slate);
  text(s, [1.93, 3.754, 2.563, 0.644], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et", { ...S.body, color: C.white });
  text(s, [0.893, 1.04, 3.848, 1.085], [{ text: "Welcome Message", options: { bold: true } }], { ...S.h1 });
  text(s, [0.893, 2.385, 4.275, 0.701], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
}

/* 5. Welcome message — founder */
function slide05(s) {
  rect(s, [0, 0, 10, 5.625], C.paper, 68.24);
  rect(s, [1.135, 2.65, 2.817, 1.51], C.panel, 65.1);
  photo(s, [0.908, 1.138, 2.497, 2.448], 'dark');
  text(s, [4.737, 1.138, 3.848, 1.085], [{ text: "Welcome Message", options: { bold: true } }], { ...S.h1 });
  text(s, [4.737, 3.236, 4.275, 0.701], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  text(s, [4.737, 2.362, 2.797, 0.344], [{ text: "John Marteen", options: { bold: true } }], { ...S.body, fontSize: 12 });
  text(s, [4.737, 2.713, 2.797, 0.311], "Founder Momiar", { ...S.label });
  text(s, [4.737, 4.086, 4.275, 0.701], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
}

/* 6. Welcome message — tall photo */
function slide06(s) {
  rect(s, [-0.012, 0, 10, 5.625], C.paper, 68.24);
  rect(s, [6.516, 0.781, 2.953, 2.55], C.panel, 65.1);
  photo(s, [6.384, 1.341, 2.497, 4.284], 'dark');
  text(s, [0.987, 2.462, 2.704, 1.085], [{ text: "Welcome", options: { breakLine: true } }, { text: "Message", options: { bold: true } }], { ...S.h1 });
  text(s, [0.987, 3.677, 3.451, 0.909], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  text(s, [4.721, 4.852, 1.379, 0.346], [{ text: "Pitch Deck", options: { bold: true } }], { ...S.body, fontSize: 12 });
  rect(s, [0.987, -0.583, 0.06, 2.917], C.slate);
}

/* 7. Diversity & inclusion statement */
function slide07(s) {
  rect(s, [0, 0, 10, 5.625], C.paper, 68.24);
  rect(s, [1.95, 2.784, 6.07, 1.865], C.panel, 65.1);
  photo(s, [0.701, 3.757, 2.497, 1.868], 'dark');
  photo(s, [6.466, 0, 3.534, 3.25], 'dark');
  text(s, [0.701, 0.748, 4.753, 1.287], [{ text: "Unleashing the Power of Diversity and Inclusion for " }, { text: "Business Success", options: { bold: true } }], { ...S.h2 });
  text(s, [4.629, 4.184, 4.275, 0.701], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
}

/* 8. About us — navy */
function slide08(s) {
  rect(s, [0, 0, 10, 5.625], C.navy);
  photo(s, [0.618, 0, 3.322, 5.625], 'dark');
  text(s, [0.927, 0.831, 2.704, 1.085], [{ text: "About Us", options: { breakLine: true } }, { text: "Momiar.", options: { bold: true } }], { ...S.h1, color: C.white });
  text(s, [5.13, 2.637, 2.192, 1.303], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body, color: C.white });
  text(s, [5.13, 4.17, 2.192, 0.701], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut", { ...S.body, color: C.white });
  text(s, [5.13, 2.097, 2.572, 0.311], [{ text: "Lorem ipsum dolor sit amet,", options: { bold: true } }], { ...S.label, color: C.white });
  rect(s, [8.891, -0.82, 0.037, 2.917], C.white);
}

/* 9. About us — banner */
function slide09(s) {
  rect(s, [0.013, 0, 10, 5.625], C.paper, 68.24);
  rect(s, [3.93, 0, 6.07, 1.865], C.panel, 65.1);
  photo(s, [0.618, 4.125, 3.322, 1.5], 'dark');
  text(s, [0.618, 0.867, 2.09, 1.085], [{ text: "About", options: { breakLine: true } }, { text: "Us", options: { bold: true } }], { ...S.h1 });
  text(s, [0.618, 2.573, 3.674, 0.48], "PLACEHOLDER", { ...S.body, fontSize: 12, lineSpacingMultiple: null });
  text(s, [5.013, 2.702, 4.275, 0.701], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  text(s, [5.015, 3.552, 4.275, 0.701], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
}

/* 10. About us — big heading */
function slide10(s) {
  rect(s, [0.013, 0, 10, 5.625], C.paper, 68.24);
  rect(s, [3.93, 0, 6.07, 1.865], C.panel, 65.1);
  rect(s, [-0.008, 3.075, 6.272, 2.55], C.panel, 85.1);
  photo(s, [6.264, 0.829, 3.322, 1.5], 'dark');
  text(s, [1.079, 0.833, 3.499, 0.909], "About Us.", { ...S.hero });
  text(s, [1.079, 1.9, 3.732, 0.429], "PLACEHOLDER", { ...S.label, lineSpacingMultiple: null });
  text(s, [1.079, 3.714, 1.932, 1.212], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  text(s, [3.846, 4.282, 4.175, 0.644], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
}

/* 11. Break slide */
function slide11(s) {
  photo(s, [0, 0, 10, 5.625], 'dark');
  rect(s, [0, 0, 10, 5.625], C.paper, 13.33);
  text(s, [1.715, 1.782, 4.238, 0.909], "Break Slide.", { ...S.hero });
  text(s, [1.715, 2.744, 3.732, 0.252], "Real Estate Presentation Template", { ...S.kicker });
  text(s, [8.421, 5.073, 1.057, 0.215], "IndoartCreative", { ...S.mark, align: 'right' });
  text(s, [0.613, 0.373, 0.947, 0.303], "Momiar.", { ...S.logo });
  text(s, [1.715, 3.218, 4.175, 0.644], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
}

/* 12. Full-service agency + 3 features */
function slide12(s) {
  rect(s, [0, 0, 10, 5.625], C.paper, 68.24);
  rect(s, [0, 0, 4.155, 1.378], C.panel, 65.1);
  rect(s, [3.728, 1.952, 6.272, 3.673], C.panel, 85.1);
  photo(s, [6.203, 0.281, 3.391, 2.25], 'dark');
  text(s, [0.916, 2.387, 4.061, 0.644], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  text(s, [0.916, 0.713, 4.371, 1.287], "We're a full-service creative agency that delivers results.", { ...S.h2 });
  text(s, [0.916, 3.948, 1.932, 1.212], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  icon(s, 'growth', [7.118, 3.819, 0.229, 0.229]);
  icon(s, 'bars', [8.689, 3.819, 0.23, 0.23]);
  icon(s, 'building', [5.536, 3.819, 0.196, 0.234]);
  ring(s, [5.389, 3.712, 0.49, 0.49]);
  ring(s, [6.987, 3.711, 0.49, 0.49]);
  ring(s, [8.564, 3.711, 0.49, 0.49]);
  text(s, [4.877, 4.514, 1.515, 0.644], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed", { ...S.body, align: 'center' });
  text(s, [5.033, 4.298, 1.202, 0.252], [{ text: "Subtitle Text", options: { bold: true } }], { ...S.kicker, align: 'center' });
  text(s, [6.478, 4.514, 1.515, 0.644], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed", { ...S.body, align: 'center' });
  text(s, [6.634, 4.298, 1.202, 0.252], [{ text: "Subtitle Text", options: { bold: true } }], { ...S.kicker, align: 'center' });
  text(s, [8.079, 4.514, 1.515, 0.644], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed", { ...S.body, align: 'center' });
  text(s, [8.235, 4.298, 1.202, 0.252], [{ text: "Subtitle Text", options: { bold: true } }], { ...S.kicker, align: 'center' });
}

/* 13. About our service */
function slide13(s) {
  rect(s, [0, 0, 10, 5.625], C.paper, 68.24);
  photo(s, [1.042, 2.725, 2.168, 2.25], 'dark');
  text(s, [4.013, 1.41, 4.549, 0.644], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  text(s, [0.959, 0.968, 2.704, 1.085], [{ text: "About Our", options: { breakLine: true } }, { text: "Service.", options: { bold: true } }], { ...S.h1 });
  rect(s, [2.922, 3.109, 6.07, 1.865], C.panel, 65.1);
  text(s, [4.013, 3.85, 4.175, 0.644], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  text(s, [4.013, 3.514, 1.612, 0.252], [{ text: "Service is priority", options: { bold: true } }], { ...S.kicker });
  tick(s, [4.783, 0.517, 0.477, 0.037]);
}

/* 14. About our service — 3 features */
function slide14(s) {
  rect(s, [0, 0, 10, 5.625], C.paper, 68.24);
  photo(s, [0, 0, 1.821, 5.625], 'dark');
  rect(s, [6.399, 1.409, 2.686, 3.538], C.panel, 65.1);
  text(s, [2.659, 2.998, 1.757, 1.401], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  text(s, [2.659, 1.409, 2.704, 1.085], [{ text: "About Our", options: { breakLine: true } }, { text: "Service.", options: { bold: true } }], { ...S.h1 });
  icon(s, 'growth', [6.747, 3.836, 0.229, 0.229]);
  icon(s, 'bars', [6.718, 2.078, 0.23, 0.23]);
  icon(s, 'building', [6.763, 2.936, 0.196, 0.234]);
  ring(s, [6.616, 2.829, 0.49, 0.49]);
  ring(s, [6.616, 3.728, 0.49, 0.49]);
  ring(s, [6.593, 1.971, 0.49, 0.49]);
  text(s, [7.221, 2.168, 1.863, 0.454], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed", { ...S.body });
  text(s, [7.218, 1.952, 1.202, 0.252], [{ text: "Subtitle Text", options: { bold: true } }], { ...S.kicker });
  text(s, [7.221, 3.049, 1.863, 0.454], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed", { ...S.body });
  text(s, [7.218, 2.833, 1.202, 0.252], [{ text: "Subtitle Text", options: { bold: true } }], { ...S.kicker });
  text(s, [7.221, 3.944, 1.922, 0.454], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed", { ...S.body });
  text(s, [7.218, 3.728, 1.202, 0.252], [{ text: "Subtitle Text", options: { bold: true } }], { ...S.kicker });
  tick(s, [4.783, 0.517, 0.477, 0.037]);
}

/* 15. Creative team — 3 portraits */
function slide15(s) {
  rect(s, [0, 0, 10, 5.625], C.paper, 68.24);
  rect(s, [3.792, 0, 6.208, 3.673], C.panel, 85.1);
  photo(s, [4.667, 1.11, 1.281, 2.744], 'dark');
  photo(s, [6.271, 1.11, 1.281, 2.744], 'dark');
  photo(s, [7.86, 1.11, 1.281, 2.744], 'dark');
  text(s, [0.78, 1.11, 2.704, 1.085], [{ text: "Creative ", options: { breakLine: true } }, { text: "Team.", options: { bold: true } }], { ...S.h1 });
  text(s, [4.592, 4.327, 4.549, 0.644], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  text(s, [0.807, 3.721, 1.932, 1.212], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  text(s, [0.78, 2.527, 1.304, 0.341], [{ text: "John Marteen", options: { bold: true } }], { ...S.label });
  text(s, [0.78, 2.788, 1.304, 0.313], "Leader ", { ...S.label });
  text(s, [0.78, 3.049, 1.304, 0.313], "Creative Team", { ...S.label });
  tick(s, [4.783, 0.517, 0.477, 0.037]);
}

/* 16. Leader team */
function slide16(s) {
  rect(s, [0, 0, 10, 5.625], C.paper, 68.24);
  rect(s, [5.768, 0, 4.232, 3.673], C.panel, 85.1);
  photo(s, [4.464, 2.774, 3, 2.851], 'dark');
  text(s, [0.78, 0.86, 2.704, 1.085], [{ text: "Leader ", options: { breakLine: true } }, { text: "Team.", options: { bold: true } }], { ...S.h1 });
  text(s, [0.78, 3.225, 2.871, 0.833], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  text(s, [0.78, 4.094, 2.871, 0.454], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut.", { ...S.body });
  text(s, [0.78, 2.774, 1.304, 0.313], "Creative Leader", { ...S.label });
  text(s, [4.464, 0.86, 1.845, 0.454], [{ text: "John Marteen", options: { bold: true } }], { ...S.body, fontSize: 15 });
  text(s, [4.464, 1.817, 3, 0.454], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut.", { ...S.body });
  text(s, [4.464, 1.259, 1.304, 0.313], "Leader", { ...S.label });
  tick(s, [8.402, 0.517, 0.477, 0.037], true);
}

/* 17. Our team — photo grid */
function slide17(s) {
  rect(s, [0, 0, 10, 5.625], C.paper, 68.24);
  rect(s, [0.006, 0, 4.232, 3.673], C.panel, 85.1);
  photo(s, [2.859, 2.881, 1.885, 1.792], 'dark');
  photo(s, [0.974, 1.089, 1.885, 1.792], 'dark');
  photo(s, [2.859, 1.089, 1.885, 1.792], 'light');
  photo(s, [0.974, 2.881, 1.885, 1.792], 'light');
  text(s, [5.589, 0.882, 1.78, 1.085], [{ text: "Our ", options: { breakLine: true } }, { text: "Team.", options: { bold: true } }], { ...S.h1 });
  text(s, [5.589, 2.146, 3.744, 0.644], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  icon(s, 'bars', [5.719, 3.306, 0.23, 0.23]);
  icon(s, 'building', [5.764, 4.165, 0.196, 0.234]);
  ring(s, [5.617, 4.058, 0.49, 0.49]);
  ring(s, [5.594, 3.199, 0.49, 0.49]);
  text(s, [6.222, 3.396, 1.863, 0.454], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed", { ...S.body });
  text(s, [6.219, 3.18, 1.202, 0.252], [{ text: "Subtitle Text", options: { bold: true } }], { ...S.kicker });
  text(s, [6.222, 4.278, 1.863, 0.454], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed", { ...S.body });
  text(s, [6.219, 4.062, 1.202, 0.252], [{ text: "Subtitle Text", options: { bold: true } }], { ...S.kicker });
  tick(s, [0.498, 0.517, 0.477, 0.037]);
}

/* 18. Product strategy — single photo */
function slide18(s) {
  rect(s, [0, 0, 10, 5.625], C.paper, 68.24);
  rect(s, [0, 1.952, 4.232, 3.673], C.panel, 85.1);
  photo(s, [1.12, 1.065, 3.689, 3.506], 'dark');
  text(s, [5.687, 1.065, 2.767, 1.085], [{ text: "Product ", options: { breakLine: true } }, { text: "Strategy.", options: { bold: true } }], { ...S.h1 });
  text(s, [5.687, 3.738, 3.435, 0.833], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  text(s, [5.687, 2.662, 1.304, 0.313], [{ text: "Strategy", options: { bold: true } }], { ...S.label });
  text(s, [5.687, 2.923, 1.304, 0.313], "Producr ", { ...S.label });
  text(s, [5.687, 3.184, 1.304, 0.313], "Creative Product", { ...S.label });
  tick(s, [5.211, 0.517, 0.477, 0.037]);
}

/* 19. Product strategy — 3 feature cards */
function slide19(s) {
  rect(s, [0, 0, 10, 5.625], C.paper, 68.24);
  rect(s, [6.864, -0.004, 3.136, 3.117], C.panel, 85.1);
  rect(s, [0, 1.952, 4.232, 3.673], C.panel, 85.1);
  rect(s, [5.938, 3.705, 2.686, 1.043], C.panel, 65.1);
  rect(s, [5.938, 2.446, 2.686, 1.043], C.panel, 65.1);
  text(s, [1.187, 1.065, 2.767, 1.085], [{ text: "Product ", options: { breakLine: true } }, { text: "Strategy.", options: { bold: true } }], { ...S.h1 });
  text(s, [1.187, 3.916, 3.435, 0.833], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  text(s, [1.187, 3.549, 1.304, 0.313], [{ text: "Strategy", options: { bold: true } }], { ...S.label });
  rect(s, [5.938, 1.182, 2.686, 1.043], C.panel, 65.1);
  icon(s, 'growth', [6.287, 4.033, 0.229, 0.229]);
  icon(s, 'bars', [6.258, 1.523, 0.23, 0.23]);
  icon(s, 'building', [6.303, 2.744, 0.196, 0.234]);
  ring(s, [6.156, 2.636, 0.49, 0.49]);
  ring(s, [6.156, 3.925, 0.49, 0.49]);
  ring(s, [6.133, 1.415, 0.49, 0.49]);
  text(s, [6.761, 1.613, 1.863, 0.454], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed", { ...S.body });
  text(s, [6.757, 1.397, 1.202, 0.252], [{ text: "Subtitle Text", options: { bold: true } }], { ...S.kicker });
  text(s, [6.761, 2.857, 1.863, 0.454], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed", { ...S.body });
  text(s, [6.757, 2.64, 1.202, 0.252], [{ text: "Subtitle Text", options: { bold: true } }], { ...S.kicker });
  text(s, [6.761, 4.142, 1.922, 0.454], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed", { ...S.body });
  text(s, [6.757, 3.925, 1.202, 0.252], [{ text: "Subtitle Text", options: { bold: true } }], { ...S.kicker });
  tick(s, [0.711, 0.517, 0.477, 0.037]);
}

/* 20. Product strategy — 2 feature cards */
function slide20(s) {
  rect(s, [0, 0, 10, 5.625], C.paper, 68.24);
  photo(s, [6.672, 1.05, 1.856, 3.506], 'dark');
  rect(s, [0, 1.952, 4.232, 3.673], C.panel, 85.1);
  text(s, [0.553, 1.107, 2.767, 1.085], [{ text: "Product ", options: { breakLine: true } }, { text: "Strategy.", options: { bold: true } }], { ...S.h1 });
  text(s, [0.553, 2.425, 4.422, 0.644], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  rect(s, [3.433, 3.497, 2.686, 1.043], C.panel, 65.1);
  rect(s, [0.553, 3.5, 2.686, 1.043], C.panel, 65.1);
  icon(s, 'bars', [0.873, 3.841, 0.23, 0.23]);
  icon(s, 'building', [3.798, 3.795, 0.196, 0.234]);
  ring(s, [3.651, 3.688, 0.49, 0.49]);
  ring(s, [0.747, 3.733, 0.49, 0.49]);
  text(s, [1.376, 3.931, 1.863, 0.454], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed", { ...S.body });
  text(s, [1.372, 3.715, 1.202, 0.252], [{ text: "Subtitle Text", options: { bold: true } }], { ...S.kicker });
  text(s, [4.256, 3.908, 1.863, 0.454], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed", { ...S.body });
  text(s, [4.252, 3.692, 1.202, 0.252], [{ text: "Subtitle Text", options: { bold: true } }], { ...S.kicker });
  tick(s, [0.617, 0.517, 0.477, 0.037]);
}

/* 21. Our portofolio — 2 banners */
function slide21(s) {
  rect(s, [0, 0, 10, 5.625], C.paper, 68.24);
  rect(s, [0, 1.952, 4.232, 3.673], C.panel, 85.1);
  photo(s, [0.922, 0.969, 4.652, 1.54], 'dark');
  photo(s, [0.922, 3.125, 4.652, 1.54], 'dark');
  text(s, [5.327, 1.718, 3.578, 1.085], [{ text: "Our ", options: { breakLine: true } }, { text: "Portofolio.", options: { bold: true } }], { ...S.h1, align: 'right' });
  text(s, [6.791, 3.181, 2.114, 1.212], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body, align: 'right' });
  tick(s, [6.553, 0.517, 0.477, 0.037]);
}

/* 22. Creative portofolio — 4 tiles */
function slide22(s) {
  rect(s, [0, 0, 10, 5.625], C.paper, 68.24);
  rect(s, [5.768, 0, 4.232, 3.673], C.panel, 85.1);
  photo(s, [5.156, 2.823, 1.762, 1.54], 'dark');
  photo(s, [7.042, 2.823, 1.762, 1.54], 'dark');
  photo(s, [7.042, 1.161, 1.762, 1.54], 'dark');
  photo(s, [5.156, 1.161, 1.762, 1.54], 'dark');
  text(s, [0.829, 1.229, 2.767, 1.085], [{ text: "Creative ", options: { breakLine: true } }, { text: "Portofolio.", options: { bold: true } }], { ...S.h1 });
  text(s, [0.829, 2.597, 3.435, 0.833], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  text(s, [0.829, 3.713, 1.671, 0.341], [{ text: "IndoartCreative", options: { bold: true } }], { ...S.label });
  tick(s, [0.449, 0.517, 0.477, 0.037]);
}

/* 23. Our portofolio — hero + 2 tiles */
function slide23(s) {
  rect(s, [0, 0, 10, 5.625], C.paper, 68.24);
  rect(s, [0, 1.952, 4.232, 3.673], C.panel, 85.1);
  photo(s, [0.703, 0.578, 8.469, 2.17], 'dark');
  photo(s, [4.09, 2.955, 2.426, 2.12], 'dark');
  photo(s, [6.746, 2.955, 2.426, 2.12], 'dark');
  text(s, [0.703, 2.955, 3.578, 1.085], [{ text: "Our ", options: { breakLine: true } }, { text: "Portofolio.", options: { bold: true } }], { ...S.h1 });
  text(s, [0.703, 4.246, 3.156, 0.833], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
}

/* 24. Our best works — collage */
function slide24(s) {
  rect(s, [0, 0, 10, 5.625], C.paper, 68.24);
  rect(s, [5.768, 0, 4.232, 3.673], C.panel, 85.1);
  photo(s, [0.199, 3.262, 2.426, 2.12], 'dark');
  photo(s, [0.199, 0.24, 4.77, 2.17], 'dark');
  photo(s, [1.931, 1.846, 3.883, 1.93], 'light');
  text(s, [6.771, 0.713, 2.092, 1.085], [{ text: "Our Best ", options: { breakLine: true } }, { text: "Works.", options: { bold: true } }], { ...S.h1, align: 'right' });
  text(s, [7.133, 1.998, 1.731, 0.813], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt labore et dolore magna aliqua.", { ...S.body, align: 'right' });
  text(s, [6.569, 4.599, 1.863, 0.434], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed", { ...S.body, align: 'right' });
  text(s, [7.226, 4.383, 1.202, 0.252], [{ text: "Subtitle Text", options: { bold: true } }], { ...S.kicker, align: 'right' });
  text(s, [3.446, 4.599, 1.863, 0.434], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed", { ...S.body, align: 'right' });
  text(s, [4.103, 4.383, 1.202, 0.252], [{ text: "Subtitle Text", options: { bold: true } }], { ...S.kicker, align: 'right' });
  text(s, [8.319, 4.007, 0.545, 0.909], "1", { ...S.numeral, align: 'center' });
  text(s, [5.249, 4.007, 0.545, 0.909], "2", { ...S.numeral, align: 'center' });
  tick(s, [6.211, 0.517, 0.477, 0.037]);
}

/* 25. Product strategy — desktop mockup */
function slide25(s) {
  rect(s, [0, 0, 10, 5.625], C.paper, 68.24);
  rect(s, [0, 0, 5.806, 4.027], C.panel, 85.1);
  chassis(s, [1.203, 1.171, 3.091, 2.567], 'monitor');
  photo(s, [1.32, 1.289, 2.844, 1.586], 'dark');
  text(s, [4.591, 4.807, 1.863, 0.454], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed", { ...S.body });
  text(s, [4.604, 4.59, 1.202, 0.252], [{ text: "Subtitle Text", options: { bold: true } }], { ...S.kicker });
  text(s, [1.468, 4.807, 1.863, 0.454], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed", { ...S.body });
  text(s, [1.477, 4.59, 1.202, 0.252], [{ text: "Subtitle Text", options: { bold: true } }], { ...S.kicker });
  text(s, [1.081, 4.214, 0.545, 0.909], "1", { ...S.numeral });
  text(s, [4.021, 4.214, 0.545, 0.909], "2", { ...S.numeral });
  text(s, [7.026, 1.171, 2.188, 1.085], [{ text: "Product ", options: { breakLine: true } }, { text: "Strategy.", options: { bold: true } }], { ...S.h1 });
  text(s, [7.026, 2.489, 1.855, 1.401], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  tick(s, [4.783, 0.517, 0.477, 0.037]);
}

/* 26. Product strategy — laptop + phone mockup */
function slide26(s) {
  rect(s, [0, 0, 5.806, 4.027], C.panel, 85.1);
  chassis(s, [0.376, 1.142, 4.56, 2.64], 'laptop');
  chassis(s, [4.054, 1.456, 1.311, 2.325], 'phone');
  photo(s, [0.859, 1.21, 3.562, 2.32], 'dark');
  photo(s, [4.109, 1.483, 1.18, 2.242], 'light');
  text(s, [4.591, 4.807, 1.863, 0.454], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed", { ...S.body });
  text(s, [4.604, 4.59, 1.202, 0.252], [{ text: "Subtitle Text", options: { bold: true } }], { ...S.kicker });
  text(s, [1.468, 4.807, 1.863, 0.454], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed", { ...S.body });
  text(s, [1.477, 4.59, 1.202, 0.252], [{ text: "Subtitle Text", options: { bold: true } }], { ...S.kicker });
  text(s, [1.081, 4.214, 0.545, 0.909], "1", { ...S.numeral });
  text(s, [4.021, 4.214, 0.545, 0.909], "2", { ...S.numeral });
  text(s, [6.704, 0.942, 2.188, 1.085], [{ text: "Product ", options: { breakLine: true } }, { text: "Strategy.", options: { bold: true } }], { ...S.h1 });
  text(s, [6.704, 2.26, 2.629, 0.833], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  text(s, [7.47, 4.807, 1.863, 0.454], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed", { ...S.body });
  text(s, [7.483, 4.59, 1.202, 0.252], [{ text: "Subtitle Text", options: { bold: true } }], { ...S.kicker });
  text(s, [6.9, 4.214, 0.545, 0.909], "3", { ...S.numeral });
}

/* 27. Mockup — bleed left */
function slide27(s) {
  rect(s, [0, 0, 10, 5.625], C.paper, 68.24);
  chassis(s, [-2.106, 0.966, 5.707, 3.303], 'laptop');
  photo(s, [-1.479, 1.072, 4.469, 2.854], 'dark');
  rect(s, [4.194, 0, 5.806, 4.75], C.panel, 85.1);
  text(s, [4.828, 1.083, 2.099, 0.581], [{ text: "Mockup" }, { text: ".", options: { bold: true } }], { ...S.h1 });
  text(s, [4.828, 1.789, 4.203, 0.454], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ", { ...S.body });
  icon(s, 'bars', [4.953, 2.784, 0.23, 0.23]);
  icon(s, 'building', [4.998, 3.781, 0.196, 0.234]);
  ring(s, [4.851, 3.674, 0.49, 0.49]);
  ring(s, [4.828, 2.677, 0.49, 0.49]);
  text(s, [5.456, 2.874, 1.863, 0.454], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed", { ...S.body });
  text(s, [5.453, 2.658, 1.202, 0.252], [{ text: "Subtitle Text", options: { bold: true } }], { ...S.kicker });
  text(s, [5.456, 3.894, 1.863, 0.454], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed", { ...S.body });
  text(s, [5.453, 3.678, 1.202, 0.252], [{ text: "Subtitle Text", options: { bold: true } }], { ...S.kicker });
  tick(s, [4.783, 0.517, 0.477, 0.037]);
}

/* 28. Mockup strategy — tilted phones */
function slide28(s) {
  rect(s, [0, 0, 10, 5.625], C.paper, 68.24);
  rect(s, [4.194, 0, 5.806, 4.75], C.panel, 85.1);
  chassis(s, [-1.054, -0.192, 2.109, 3.74], 'phone', { rot: 329.25 });
  chassis(s, [1.677, -0.192, 2.109, 3.74], 'phone', { rot: 329.25 });
  chassis(s, [3.768, 3.152, 2.109, 3.74], 'phone', { rot: 329.25 });
  chassis(s, [0.963, 3.237, 2.109, 3.74], 'phone', { rot: 329.25 });
  photo(s, [1.741, -0.16, 1.933, 3.654], 'dark', { rot: 328.57 });
  photo(s, [1.059, 3.279, 1.933, 3.654], 'dark', { rot: 328.57 });
  photo(s, [3.856, 3.194, 1.933, 3.654], 'dark', { rot: 328.57 });
  text(s, [6.195, 0.854, 2.188, 1.085], [{ text: "Mockup ", options: { breakLine: true } }, { text: "Strategy.", options: { bold: true } }], { ...S.h1 });
  text(s, [6.195, 2.172, 2.574, 1.022], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  photo(s, [-0.947, -0.17, 1.933, 3.654], 'dark', { rot: 328.57 });
}

/* 29. Interested our chart */
function slide29(s) {
  rect(s, [0, 0, 10, 5.625], C.paper, 68.24);
  rect(s, [0.031, 0, 4.493, 4.75], C.panel, 85.1);
  lineChartArt(s);
  text(s, [5.169, 1.093, 4.659, 1.085], [{ text: "Interested our chart ", options: { breakLine: true } }, { text: "Strategy.", options: { bold: true } }], { ...S.h1 });
  text(s, [5.169, 3.708, 4.357, 0.644], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
  text(s, [5.169, 2.539, 0.831, 0.429], "123K", { ...S.stat });
  text(s, [5.169, 2.917, 0.831, 0.252], "Client", { ...S.kicker });
  text(s, [6.471, 2.539, 1.016, 0.429], "200K", { ...S.stat });
  text(s, [6.471, 2.917, 1.016, 0.252], "Testimonial", { ...S.kicker });
  tick(s, [4.783, 0.517, 0.477, 0.037]);
}

/* 30. Thanks slide */
function slide30(s) {
  photo(s, [0, 0, 10, 5.625], 'dark');
  rect(s, [0, 0, 10, 5.625], C.paper, 13.33);
  text(s, [1.56, 1.782, 4.78, 0.909], "Thanks Slide.", { ...S.hero });
  text(s, [1.56, 2.744, 2.65, 0.252], "Real Estate Presentation Template", { ...S.kicker });
  text(s, [8.421, 5.073, 1.057, 0.215], "IndoartCreative", { ...S.mark, align: 'right' });
  text(s, [0.613, 0.373, 0.947, 0.303], "Momiar.", { ...S.logo });
  text(s, [1.56, 3.218, 4.175, 0.644], "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut", { ...S.body });
}

const SLIDES = [
  slide01, slide02, slide03, slide04, slide05,
  slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15,
  slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25,
  slide26, slide27, slide28, slide29, slide30,
];

/* ------------------------------------------------------------------- render */
function main() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'MOMIAR', width: 10, height: 5.625 });
  pptx.layout = 'MOMIAR';
  pptx.author = 'IndoartCreative';
  pptx.title = 'Momiar. — Real Estate Presentation Template';

  SLIDES.forEach((build) => build(pptx.addSlide()));

  const out = path.join(__dirname, '0e338a33-166f-4050-9fc1-c5e6979f81b3_grok_final.pptx');
  return pptx.writeFile({ fileName: out }).then(() => console.log('wrote ' + out));
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
