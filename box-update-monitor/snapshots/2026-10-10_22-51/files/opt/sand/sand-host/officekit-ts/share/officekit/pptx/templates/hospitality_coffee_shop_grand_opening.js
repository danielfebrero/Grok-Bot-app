/*
 * "Soffee — Grand Opening" coffee-shop story deck.
 * 9 slides, 11.25" x 19.72" portrait (Instagram-story proportions), rebuilt with pptxgenjs.
 *
 * Every photo in the source file is a flat grey swatch, so each one is redrawn here
 * as a native #CCCCCC shape carrying the same mask (circle / heptagon / arch / ...).
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// --------------------------------------------------------------- deck basics
const DECK = { w: 11.251722, h: 19.717 };

const C = {
  cream: 'F5ECD8', // page background
  green: '1D4C34',
  greenDark: '153926',
  gold: 'D4A94F',
  sand: 'EDDCB7',
  ring: 'E5CB94',
  leaf: '548135',
  photo: 'CCCCCC', // stand-in for the deck's photo placeholders
  panel: 'F2F2F2',
  grey: '7F7F7F',
  ink: '000000',
  inkSoft: '262626',
  body: '595959',
  white: 'FFFFFF',
};

const F = { serif: 'PT Serif', sans: 'Roboto' };

// Shadow presets, in points. These are factories because pptxgenjs rewrites the
// object it is handed (pt -> EMU) and would double-convert a shared literal.
// It also reads an offset/angle of exactly 0 as "unset", hence the 0.01 nudges.
const SHADOW = {
  badge: () => ({ type: 'outer', blur: 42, offset: 0.01, angle: 0.01, color: C.ink, opacity: 0.2 }),
  dot: () => ({ type: 'outer', blur: 5, offset: 2, angle: 90, color: '5E5E5E', opacity: 0.5 }),
  button: () => ({ type: 'outer', blur: 31, offset: 3, angle: 90, color: C.ink, opacity: 0.08 }),
  disc: () => ({ type: 'outer', blur: 33, offset: 0.01, angle: 0.01, color: C.ink, opacity: 0.12 }),
  device: () => ({ type: 'outer', blur: 100, offset: 60, angle: 90, color: '3F3F3F', opacity: 0.31 }),
};

// ------------------------------------------------------------- path library
// Outlines normalised to the unit square. Command forms:
//   [x, y]                    line-to  (move-to when it opens a sub-path)
//   ['m', x, y]               move-to  (explicitly starts a new sub-path)
//   ['c', x1,y1, x2,y2, x,y]  cubic bezier
//   ['z']                     close sub-path
const BLOB = [[0, 0], [1, 0], [1, 1], ['c', 0.77, 0.865, 0.883, 0.584, 0.57, 0.381], ['c', 0.258, 0.178, 0.104, 0.198, 0, 0], ['z']];
const BADGE = [[0.5, 0], ['c', 0.536, 0, 0.572, 0.008, 0.604, 0.025], [0.896, 0.18], ['c', 0.961, 0.214, 1, 0.278, 1, 0.345], [1, 0.655], ['c', 1, 0.724, 0.961, 0.786, 0.896, 0.82], [0.604, 0.975], ['c', 0.539, 1.008, 0.461, 1.008, 0.396, 0.975], [0.104, 0.82], ['c', 0.039, 0.786, 0, 0.722, 0, 0.655], [0, 0.345], ['c', 0, 0.276, 0.039, 0.214, 0.104, 0.18], [0.396, 0.025], ['c', 0.428, 0.008, 0.464, 0, 0.5, 0], ['z']];
const DBOWL = [[0.752, 0], ['c', 0.83, 0, 0.905, 0.008, 0.975, 0.022], [1, 0.028], [1, 0.972], [0.975, 0.978], ['c', 0.905, 0.992, 0.83, 1, 0.752, 1], ['c', 0.337, 1, 0, 0.776, 0, 0.5], ['c', 0, 0.224, 0.337, 0, 0.752, 0], ['z']];
const LEAF = [[0.202, 0.584], ['c', 0.347, 0.769, 0.554, 0.909, 0.789, 1], ['c', 0.746, 0.979, 0.714, 0.955, 0.685, 0.92], ['c', 0.681, 0.92, 0.681, 0.916, 0.681, 0.916], ['c', 0.657, 0.885, 0.643, 0.853, 0.643, 0.818], ['c', 0.634, 0.738, 0.69, 0.654, 0.793, 0.608], ['c', 0.944, 0.542, 1, 0.395, 0.911, 0.28], ['c', 0.822, 0.168, 0.624, 0.126, 0.469, 0.192], ['c', 0.366, 0.238, 0.244, 0.231, 0.15, 0.185], ['c', 0.113, 0.168, 0.08, 0.143, 0.056, 0.112], ['c', 0.056, 0.108, 0.056, 0.108, 0.052, 0.108], ['c', 0.028, 0.073, 0.014, 0.035, 0.014, 0], ['c', 0, 0.196, 0.061, 0.402, 0.202, 0.584], ['z']];
const LEAF_TIP = [[0, 0], ['c', 0, 0.132, 0.03, 0.265, 0.08, 0.397], ['c', 0.26, 0.853, 0.65, 1, 0.96, 0.75], ['c', 0.97, 0.735, 1, 0.706, 1, 0.706], ['c', 0.64, 0.574, 0.3, 0.338, 0, 0], ['z']];
const LEAF_TAIL = [[1, 0], ['c', 0.978, 0, 0.935, 0.009, 0.913, 0.018], ['c', 0.239, 0.18, 0, 0.532, 0.391, 0.811], ['c', 0.5, 0.892, 0.652, 0.955, 0.826, 1], ['c', 0.717, 0.658, 0.783, 0.315, 1, 0], ['z']];
const CHEVRON = [[1, 1], [0.496, 0], [0, 1]];
const PIN = [[1, 0.375], ['c', 1, 0.498, 0.825, 0.694, 0.682, 0.833], ['c', 0.583, 0.93, 0.5, 1, 0.5, 1], ['c', 0.5, 1, 0.417, 0.93, 0.318, 0.833], ['c', 0.175, 0.694, 0, 0.498, 0, 0.375], ['c', 0, 0.168, 0.224, 0, 0.5, 0], ['c', 0.776, 0, 1, 0.168, 1, 0.375], ['z']];
const PHONE = [[0.581, 0.58], ['c', 0.494, 0.669, 0.391, 0.754, 0.352, 0.714], ['c', 0.293, 0.655, 0.257, 0.605, 0.128, 0.709], ['c', 0, 0.812, 0.101, 0.88, 0.156, 0.936], ['c', 0.221, 1, 0.461, 0.938, 0.701, 0.7], ['c', 0.939, 0.462, 1, 0.218, 0.936, 0.154], ['c', 0.88, 0.098, 0.81, 0, 0.709, 0.126], ['c', 0.606, 0.255, 0.656, 0.291, 0.715, 0.35], ['c', 0.754, 0.389, 0.67, 0.493, 0.581, 0.58], ['z']];
const CUP = [[0.773, 0.519], [0.773, 0.768], [0.859, 0.725], [0.893, 0.644], [0.859, 0.562], [0.773, 0.519], ['z'], ['m', 0.187, 0.425], [0.149, 0.439], [0.134, 0.471], [0.135, 0.825], [0.157, 0.854], [0.187, 0.862], [0.224, 0.849], [0.24, 0.816], [0.239, 0.462], [0.217, 0.433], [0.187, 0.425], ['z'], ['m', 0.245, 0], [0.273, 0.018], [0.272, 0.081], [0.373, 0.206], [0.373, 0.257], [0.342, 0.31], [0.501, 0.31], [0.493, 0.274], [0.413, 0.177], [0.396, 0.119], [0.427, 0.048], [0.468, 0.007], [0.497, 0.002], [0.512, 0.018], [0.511, 0.081], [0.587, 0.162], [0.616, 0.218], [0.608, 0.27], [0.581, 0.31], [0.749, 0.318], [0.772, 0.347], [0.773, 0.427], [0.901, 0.471], [0.982, 0.564], [0.999, 0.665], [0.961, 0.759], [0.881, 0.829], [0.771, 0.861], [0.746, 0.923], [0.679, 0.979], [0.587, 1], [0.166, 0.999], [0.076, 0.969], [0.009, 0.89], [0.001, 0.347], [0.043, 0.311], [0.261, 0.31], [0.253, 0.274], [0.178, 0.186], [0.156, 0.13], [0.178, 0.061], [0.245, 0], ['z']];
const HOME = [[0.938, 0.656], [0.875, 0.656], [0.858, 0.743], [0.802, 0.822], [0.743, 0.858], [0.672, 0.875], [0.344, 0.875], [0.27, 0.863], [0.178, 0.802], [0.129, 0.702], [0.125, 0.344], [0.137, 0.27], [0.161, 0.22], [0.208, 0.169], [0.27, 0.137], [0.328, 0.125], [0.672, 0.125], [0.73, 0.137], [0.792, 0.169], [0.846, 0.231], [0.871, 0.298], [0.875, 0.656], [1, 0.656], [0.999, 0.32], [0.984, 0.232], [0.924, 0.118], [0.823, 0.037], [0.683, 0.001], [0.275, 0.007], [0.135, 0.072], [0.037, 0.191], [0.001, 0.32], [0.007, 0.734], [0.059, 0.861], [0.171, 0.959], [0.253, 0.99], [0.344, 1], [0.683, 0.999], [0.823, 0.969], [0.924, 0.895], [0.977, 0.802], [1, 0.656], [0.938, 0.656], ['z']];
const INSTAGRAM = [[0.87, 0.998], [0.12, 0.998], [0.071, 0.984], [0.033, 0.954], [0, 0.879], [0, 0.128], [0.01, 0.078], [0.044, 0.032], [0.12, 0], [0.87, 0], [0.92, 0.01], [0.966, 0.044], [0.998, 0.119], [0.998, 0.879], [0.985, 0.927], [0.955, 0.966], [0.87, 0.998], ['z'], ['m', 0.499, 0.302], [0.425, 0.317], [0.367, 0.357], [0.314, 0.429], [0.302, 0.499], [0.318, 0.574], [0.361, 0.636], [0.424, 0.68], [0.499, 0.696], [0.574, 0.68], [0.637, 0.636], [0.68, 0.574], [0.696, 0.499], [0.685, 0.429], [0.631, 0.357], [0.573, 0.317], [0.499, 0.302], ['z'], ['m', 0.882, 0.163], [0.871, 0.135], [0.844, 0.118], [0.706, 0.125], [0.685, 0.16], [0.693, 0.292], [0.728, 0.314], [0.86, 0.305], [0.881, 0.272], [0.882, 0.163], ['z'], ['m', 0.882, 0.417], [0.79, 0.417], [0.801, 0.493], [0.79, 0.578], [0.759, 0.65], [0.711, 0.711], [0.615, 0.777], [0.499, 0.801], [0.383, 0.777], [0.287, 0.711], [0.239, 0.65], [0.208, 0.578], [0.197, 0.493], [0.209, 0.417], [0.117, 0.417], [0.122, 0.855], [0.163, 0.882], [0.855, 0.876], [0.882, 0.839], [0.882, 0.417], ['z']];
const FACEBOOK = [[0.915, 0.998], [0.07, 0.998], [0.016, 0.975], [0, 0.928], [0, 0.082], [0.029, 0.017], [0.07, 0], [0.915, 0], [0.978, 0.027], [0.998, 0.082], [0.998, 0.934], [0.975, 0.981], [0.915, 0.998], ['z'], ['m', 0.424, 0.188], [0.327, 0.203], [0.271, 0.238], [0.234, 0.281], [0.223, 0.423], [0.153, 0.423], [0.153, 0.541], [0.223, 0.541], [0.223, 0.928], [0.387, 0.928], [0.387, 0.541], [0.482, 0.541], [0.517, 0.423], [0.387, 0.423], [0.39, 0.335], [0.399, 0.318], [0.456, 0.306], [0.517, 0.329], [0.551, 0.223], [0.424, 0.188], ['z']];
const TWITTER = [[0.681, 0], [0.763, 0.02], [0.846, 0.087], [0.975, 0.03], [0.939, 0.106], [0.88, 0.157], [0.998, 0.129], [0.904, 0.258], [0.896, 0.393], [0.857, 0.543], [0.796, 0.687], [0.733, 0.785], [0.592, 0.918], [0.493, 0.97], [0.411, 0.988], [0.199, 0.984], [0, 0.885], [0.115, 0.878], [0.246, 0.827], [0.306, 0.77], [0.246, 0.77], [0.17, 0.713], [0.118, 0.599], [0.159, 0.609], [0.199, 0.599], [0.159, 0.586], [0.107, 0.54], [0.05, 0.448], [0.035, 0.357], [0.13, 0.385], [0.073, 0.305], [0.044, 0.225], [0.042, 0.139], [0.064, 0.066], [0.095, 0.073], [0.136, 0.13], [0.283, 0.245], [0.399, 0.3], [0.493, 0.315], [0.487, 0.224], [0.529, 0.108], [0.583, 0.047], [0.681, 0], ['z']];

// Rectangle whose two top corners are rounded by half its shortest side
// (PowerPoint's `round2SameRect` at adj1 = 50%) — the arch on slide 9.
function archOutline(w, h) {
  const r = Math.min(w, h) / 2;
  const k = 0.5523; // circle-to-bezier handle length
  const rx = r / w;
  const ry = r / h;
  return [
    [0, 1], [0, ry],
    ['c', 0, ry - ry * k, rx - rx * k, 0, rx, 0],
    [1 - rx, 0],
    ['c', 1 - rx + rx * k, 0, 1, ry - ry * k, 1, ry],
    [1, 1], ['z'],
  ];
}

// -------------------------------------------------------------- draw helpers
function outlineToPoints(outline, w, h) {
  const pts = [];
  let open = false;
  for (const cmd of outline) {
    if (cmd[0] === 'z') { pts.push({ close: true }); open = false; } else if (cmd[0] === 'c') {
      pts.push({ x: cmd[5] * w, y: cmd[6] * h, curve: { type: 'cubic', x1: cmd[1] * w, y1: cmd[2] * h, x2: cmd[3] * w, y2: cmd[4] * h } });
    } else if (cmd[0] === 'm') {
      pts.push({ x: cmd[1] * w, y: cmd[2] * h, moveTo: true }); open = true;
    } else {
      pts.push({ x: cmd[0] * w, y: cmd[1] * h, moveTo: !open }); open = true;
    }
  }
  return pts;
}

/** Free-form shape from a unit-square outline. */
function poly(slide, outline, o) {
  slide.addShape('custGeom', Object.assign({ points: outlineToPoints(outline, o.w, o.h) }, o));
}

const shape = (slide, kind, o) => slide.addShape(kind, o);
const text = (slide, runs, o) => slide.addText(runs, Object.assign({ valign: 'top', fontFace: F.serif, color: C.ink }, o));

/** Grey stand-in for one of the deck's photo placeholders. */
const photoRect = (slide, o) => slide.addShape('rect', Object.assign({ fill: { color: C.photo } }, o));
const photoDisc = (slide, o) => slide.addShape('ellipse', Object.assign({ fill: { color: C.photo } }, o));
const photoPoly = (slide, outline, o) => poly(slide, outline, Object.assign({ fill: { color: C.photo } }, o));

// ------------------------------------------------------- recurring furniture
/** Organic corner blob used as the layout background on almost every slide. */
function blob(slide, x, y, w, h, o) {
  poly(slide, BLOB, Object.assign({ x, y, w, h, fill: { color: C.green } }, o));
}

/** Scroll indicator: three shrinking discs plus an up chevron. Bottom centre. */
function scrollDots(slide, accent) {
  const rings = [
    { x: 5.622, y: 17.65, d: 0.739, inset: 0.062 },
    { x: 5.827, y: 18.545, d: 0.329, inset: 0.027 },
    { x: 5.901, y: 19.118, d: 0.183, inset: 0.015 },
  ];
  for (const r of rings) {
    shape(slide, 'ellipse', { x: r.x, y: r.y, w: r.d, h: r.d, fill: { color: C.white }, shadow: SHADOW.dot() });
    shape(slide, 'ellipse', { x: r.x + r.inset, y: r.y + r.inset, w: r.d - 2 * r.inset, h: r.d - 2 * r.inset, fill: { color: accent } });
  }
  poly(slide, CHEVRON, { x: 5.827, y: 17.929, w: 0.329, h: 0.169, fill: { type: 'none' }, line: { color: C.white, width: 5 } });
}

/** Heptagon badge carrying the white coffee-cup mark. */
function cupBadge(slide, x, y, fill) {
  poly(slide, BADGE, { x, y, w: 0.779, h: 0.84, fill: { color: fill }, shadow: SHADOW.badge() });
  poly(slide, CUP, { x: x + 0.229, y: y + 0.234, w: 0.321, h: 0.373, fill: { color: C.white }, flipH: true });
}

/** Green disc carrying the white handset mark, next to the phone number. */
function phoneBadge(slide, x, y, d) {
  shape(slide, 'ellipse', { x, y, w: d, h: d * 0.995, fill: { color: C.green } });
  poly(slide, PHONE, { x: x + d * 0.294, y: y + d * 0.294, w: d * 0.411, h: d * 0.408, fill: { color: C.white } });
}

/** Tilted "GRAND OPENING" / "OPENING PROMO" sticker on a heptagon badge. */
function sticker(slide, o) {
  slide.addText([{ text: o.head }, { text: o.tail, options: { underline: { style: 'sng' } } }], {
    x: o.x, y: o.y, w: o.w, h: o.h, rotate: o.rotate || 0,
    shape: 'custGeom', points: outlineToPoints(BADGE, o.w, o.h),
    fill: { color: o.fill }, shadow: SHADOW.badge(),
    align: 'center', valign: 'middle', fontFace: F.serif, fontSize: o.fontSize, bold: true, color: C.white,
  });
}

/** Green pill button: a bold word followed by a light one. */
function pillButton(slide, o) {
  slide.addText([{ text: o.head, options: { bold: true } }, { text: o.tail }], {
    x: o.x, y: o.y, w: o.w, h: o.h, shape: 'roundRect', rectRadius: Math.min(o.w, o.h) / 2,
    fill: { color: C.green }, shadow: SHADOW.button(),
    align: 'center', valign: 'middle', fontFace: F.serif, fontSize: o.fontSize, color: C.white,
  });
}

/** Small caption flanked by two gold diamonds ("Our Best Coffee's", "Follow Us"). */
function diamondCaption(slide, o) {
  text(slide, o.label, {
    x: o.x, y: o.y, w: o.w, h: o.h, rotate: o.rotate || 0,
    align: 'center', fontSize: o.fontSize, bold: true,
  });
  for (const d of o.diamonds) {
    shape(slide, 'diamond', { x: d[0], y: d[1], w: 0.13, h: 0.13, rotate: o.diamondRotate, fill: { color: C.gold } });
  }
}

/** Price bubble: "IDR" over a struck-through old price over the new price. */
function priceBubble(slide, o) {
  const runs = [
    { text: 'IDR', options: { breakLine: true } },
    { text: o.was, options: { strike: 'sngStrike', color: o.wasColor, breakLine: true } },
    { text: o.now },
  ];
  slide.addText(runs, {
    x: o.x, y: o.y, w: o.d, h: o.d, shape: 'ellipse', fill: { color: o.fill },
    align: 'center', valign: 'middle', fontFace: F.serif, fontSize: o.fontSize, bold: true, color: C.white,
  });
}

const WORDMARK = [
  { text: 'West Cowboy ', options: { italic: true } },
  { text: 'Coffee Shop' },
];
const HEADER_LINE = 'Grand Opening Soffee West Cowboy  Coffe Shop';
const PHONE_NUMBER = '+65 726 6143 8826';
const LOREM_LONG = 'Dev certa volle sopra anime animo qua. Tenta anima la no aveva ch rombo le. Se da distrutta liberarli infantile usignuoli. Ora afa rimorso dai braccia sentito chiedere fu acerbita sussulto. Eri lei dio arresti maniera compita';
const LOREM_MED = 'Dev certa volle sopra anime animo qua. Tenta anima la';
const LOREM_SHORT = 'Dev certa volle sopra anime animo qua. Tenta anima la no aveva ch';

// ------------------------------------------------------------------- slides
function slide1(pres) {
  const s = pres.addSlide();
  s.background = { color: C.cream };
  shape(s, 'ellipse', { x: 6.607, y: 4.692, w: 2.599, h: 2.599, fill: { color: C.green } });
  blob(s, 0, 0, 3.83, 2.911, { flipH: true });
  shape(s, 'ellipse', { x: 2.295, y: 3.667, w: 1.315, h: 1.315, fill: { color: C.green } });

  photoPoly(s, BADGE, { x: 1.672, y: 4.981, w: 7.908, h: 8.524 });

  // Title spine running bottom-to-top along the left edge.
  text(s, HEADER_LINE, { x: -4.442, y: 12.285, w: 10.748, h: 0.639, rotate: 270, align: 'center', fontSize: 32 });
  poly(s, BADGE, { x: 0.816, y: 17.979, w: 0.232, h: 0.25, fill: { color: C.gold }, shadow: SHADOW.badge() });

  text(s, 'Soffee', { x: 2.65, y: 13.66, w: 5.952, h: 2.423, align: 'center', fontSize: 138, bold: true });
  text(s, WORDMARK, { x: 2.986, y: 16.084, w: 5.281, h: 0.639, align: 'center', fontSize: 32 });

  phoneBadge(s, 9.834, 0.958, 0.479);
  text(s, PHONE_NUMBER, { x: 6.362, y: 1.045, w: 3.321, h: 0.505, align: 'right', fontSize: 24 });

  sticker(s, { x: 7.349, y: 11.072, w: 1.833, h: 1.976, rotate: 345, fill: C.green, head: 'GRAND ', tail: 'OPENING', fontSize: 24 });
  scrollDots(s, C.green);
}

function slide2(pres) {
  const s = pres.addSlide();
  s.background = { color: C.cream };
  blob(s, 0, 16.806, 3.83, 2.911, { rotate: 180 });
  shape(s, 'ellipse', { x: 6.017, y: 10.059, w: 0.749, h: 0.749, fill: { color: C.green } });
  shape(s, 'ellipse', { x: 5.251, y: 2.086, w: 0.375, h: 0.375, fill: { color: C.green } });

  photoDisc(s, { x: 0.479, y: 4.538, w: 5.521, h: 5.521 });
  photoDisc(s, { x: 6.849, y: 10.666, w: 3.468, h: 3.468 });

  text(s, [
    { text: 'Welcome, Soffee West ' },
    { text: 'Cowboy Coffeeshop', options: { color: C.green } },
  ], { x: 6.441, y: 5.682, w: 5.147, h: 3.332, fontSize: 48, bold: true, color: C.inkSoft });

  text(s, LOREM_LONG, {
    x: 1.561, y: 14.745, w: 7.933, h: 3.063, align: 'justify', fontFace: F.sans,
    fontSize: 24, color: C.body, lineSpacingMultiple: 1.5,
  });

  pillButton(s, { x: 1.561, y: 13.852, w: 2.615, h: 0.565, head: 'Watch ', tail: 'More', fontSize: 18 });

  cupBadge(s, 9.805, 0.874, C.green);
  text(s, HEADER_LINE, { x: 0.479, y: 0.805, w: 5.405, h: 0.909, align: 'center', fontSize: 24 });
  sticker(s, { x: 1.124, y: 8.706, w: 1.833, h: 1.976, rotate: 345, fill: C.gold, head: 'GRAND ', tail: 'OPENING', fontSize: 24 });
}

function slide3(pres) {
  const s = pres.addSlide();
  s.background = { color: C.cream };
  blob(s, 7.422, 0, 3.83, 2.911);

  photoRect(s, { x: 1.092, y: 7.835, w: 9.068, h: 9.422 });

  // Two intro columns, each headed by a small gold badge.
  const columns = [
    { x: 1.752, title: 'Coffee Variants' },
    { x: 6.102, title: 'Latte Variants' },
  ];
  for (const col of columns) {
    poly(s, BADGE, { x: col.x, y: 4.633, w: 0.426, h: 0.459, fill: { color: C.gold }, shadow: SHADOW.badge() });
    text(s, col.title, { x: col.x + 0.125, y: 4.83, w: 2.826, h: 0.505, fontSize: 24 });
    text(s, LOREM_SHORT, {
      x: col.x + 0.156, y: 5.479, w: 3.221, h: 1.56, align: 'justify', fontFace: F.sans,
      fontSize: 20, color: C.body, lineSpacingMultiple: 1.5,
    });
  }

  sticker(s, { x: 8.563, y: 0.483, w: 1.833, h: 1.976, fill: C.gold, head: 'OPENING ', tail: 'PROMO', fontSize: 24 });

  // Gold quote card overlapping the photo.
  shape(s, 'rect', { x: 2.78, y: 13.744, w: 5.691, h: 4.429, fill: { color: C.gold } });
  text(s, 'West Coffeeshop', { x: 3.225, y: 14.188, w: 4.801, h: 0.699, align: 'center', fontSize: 20, bold: true, color: C.white });
  text(s, 'Dev certa volle sopra anime animo qua. Tenta anima la no aveva ch rombo le. Se da distrutta liberarli infantile', {
    x: 3.225, y: 14.672, w: 4.801, h: 2.064, align: 'center', fontFace: F.sans,
    fontSize: 20, italic: true, color: C.white, lineSpacingMultiple: 1.5,
  });
  text(s, 'BY FOUNDER', { x: 4.268, y: 17.318, w: 2.716, h: 0.37, align: 'center', fontSize: 16, bold: true, color: C.white });
  for (const x of [4.298, 6.72]) {
    poly(s, BADGE, { x, y: 17.456, w: 0.166, h: 0.179, fill: { color: C.white }, shadow: SHADOW.badge() });
  }

  phoneBadge(s, 0.545, 0.508, 0.479);
  text(s, PHONE_NUMBER, { x: 1.144, y: 0.595, w: 3.321, h: 0.505, fontSize: 24 });
}

function slide4(pres) {
  const s = pres.addSlide();
  s.background = { color: C.cream };
  blob(s, 7.965, 16.342, 3.835, 2.915, { rotate: 90 });
  blob(s, -0.46, 0.46, 3.835, 2.915, { rotate: 270 });

  // Three numbered products; odd rows read left-to-right, the middle one mirrors.
  const items = [
    {
      photo: [0.835, 1.754], num: '01', numAt: [3.682, 1.55], align: 'left',
      title: 'COFFEE CUP VARIANTS', titleAt: [5.122, 2.279, 4.301], bodyAt: [5.122, 2.649, 3.706],
      price: { x: 0.355, y: 4.384, was: '29K', now: '20K' },
    },
    {
      photo: [5.72, 6.242], num: '02', numAt: [5.255, 6.227], align: 'right',
      title: 'LATTE CUP VARIANTS', titleAt: [1.487, 7.281, 3.822], bodyAt: [1.876, 7.804, 3.433],
      price: { x: 8.076, y: 8.771, was: '25K', now: '16K' },
    },
    {
      photo: [1.458, 11.105], num: '03', numAt: [4.634, 10.951], align: 'left',
      title: 'EXPRESSO CUP VARIANTS', titleAt: [5.901, 12.105, 4.586], bodyAt: [5.901, 12.672, 3.706],
      price: { x: 0.952, y: 13.549, was: '20K', now: '11K' },
    },
  ];
  for (const it of items) {
    photoDisc(s, { x: it.photo[0], y: it.photo[1], w: 4.162, h: 4.162 });
    text(s, it.num, { x: it.numAt[0], y: it.numAt[1], w: 1.29, h: 1.212, align: 'center', fontSize: 66, bold: true });
    text(s, it.title, { x: it.titleAt[0], y: it.titleAt[1], w: it.titleAt[2], h: 0.505, align: it.align, fontSize: 24 });
    text(s, LOREM_MED, {
      x: it.bodyAt[0], y: it.bodyAt[1], w: it.bodyAt[2], h: 1.851, align: it.align, fontFace: F.sans,
      fontSize: 24, color: C.body, lineSpacingMultiple: 1.5,
    });
    priceBubble(s, { x: it.price.x, y: it.price.y, d: 2.176, fill: C.green, was: it.price.was, now: it.price.now, wasColor: C.gold, fontSize: 32 });
  }

  diamondCaption(s, {
    label: 'Our Best Coffee\u2019s', x: 8.649, y: 1.991, w: 3.437, h: 0.606, rotate: 90, fontSize: 24,
    diamonds: [[10.303, 0.301], [10.303, 4.028]], diamondRotate: 180,
  });
  text(s, WORDMARK, { x: 2.986, y: 16.434, w: 5.281, h: 0.639, align: 'center', fontSize: 32 });
  scrollDots(s, C.green);
}

function slide5(pres) {
  const s = pres.addSlide();
  s.background = { color: C.cream };
  // Concentric hairline rings radiating out from the hero photo.
  const rings = [[1.985, 5.848, 7.281, 7.171], [1.544, 5.413, 8.164, 8.041], [1.15, 5.025, 8.952, 8.817], [0.771, 4.652, 9.711, 9.564]];
  for (const [x, y, w, h] of rings) {
    shape(s, 'ellipse', { x, y, w, h, fill: { type: 'none' }, line: { color: C.ring, width: 2.25 } });
  }
  blob(s, 6.647, 0, 4.604, 3.5);
  blob(s, 0, 15.577, 5.527, 4.202, { rotate: 180 });

  photoDisc(s, { x: 2.443, y: 6.251, w: 6.365, h: 6.365 });

  text(s, 'We Are Have The Best Cozy Place', { x: 0.863, y: 5.363, w: 6.022, h: 1.717, fontSize: 48, bold: true, color: C.inkSoft });
  text(s, 'Let\u2019s Chill & Relax', { x: 4.979, y: 12.737, w: 4.34, h: 1.717, fontSize: 48, bold: true, color: C.inkSoft });

  cupBadge(s, 10.004, 0.354, C.gold);
  text(s, HEADER_LINE, { x: 0.51, y: 0.588, w: 5.405, h: 0.909, align: 'center', fontSize: 24 });
  sticker(s, { x: 6.873, y: 5.867, w: 2.463, h: 2.655, rotate: 345, fill: C.gold, head: 'OPENING ', tail: 'PROMO', fontSize: 32 });

  // Two dine-in offers.
  const offers = [
    { x: 3.518, title: '1 Food & 1 Coffee', pct: '40%' },
    { x: 6.999, title: '2 Food & 3 Coffee', pct: '50%' },
  ];
  for (const o of offers) {
    text(s, o.title, { x: o.x, y: 15.23, w: 3.294, h: 0.505, fontSize: 24, bold: true });
    text(s, [
      { text: 'Dine Only ' }, { text: o.pct, options: { color: C.greenDark } }, { text: ' Off' },
    ], { x: o.x, y: 15.768, w: 3.294, h: 0.438, fontSize: 20 });
  }

  diamondCaption(s, {
    label: 'Our Best Coffee\u2019s', x: 0.93, y: 13.915, w: 3.01, h: 0.538, fontSize: 20,
    diamonds: [[0.733, 14.119], [4.024, 14.119]], diamondRotate: 90,
  });
  scrollDots(s, C.green);
}

function slide6(pres) {
  const s = pres.addSlide();
  s.background = { color: C.cream };

  // Four alternating cards: white disc + dark accent leaf behind the photo,
  // a bright leaf on top, a gold price bubble and a caption.
  const cards = [
    { disc: [5.813, 1.727], tip: [8.605, 1.751], tail: [9.979, 3.517], photo: [6.052, 1.975], leaf: [8.739, 1.856], price: [4.724, 2.805, '29K', '20K'], caption: [6.452, 4.28] },
    { disc: [1.113, 4.56], tip: [3.905, 4.584], tail: [5.279, 6.35], photo: [1.352, 4.808], leaf: [4.04, 4.689], price: [0.522, 4.186, '25K', '16K'], caption: [1.75, 7.229] },
    { disc: [4.906, 9.149], tip: [7.697, 9.173], tail: [9.071, 10.94], photo: [5.145, 9.397], leaf: [7.832, 9.278], price: [4.082, 8.88, '20K', '11K'], caption: [5.586, 11.684] },
    { disc: [1.113, 13.01], tip: [3.905, 13.034], tail: [5.279, 14.8], photo: [1.352, 13.258], leaf: [4.04, 13.139], price: [0.522, 12.328, '25K', '16K'], caption: [1.75, 15.642] },
  ];
  for (const c of cards) {
    poly(s, LEAF_TIP, { x: c.tip[0], y: c.tip[1], w: 0.715, h: 0.485, rotate: 180, fill: { color: C.greenDark } });
    poly(s, LEAF_TAIL, { x: c.tail[0], y: c.tail[1], w: 0.328, h: 0.794, rotate: 180, fill: { color: C.greenDark } });
    shape(s, 'ellipse', { x: c.disc[0], y: c.disc[1], w: 4.25, h: 4.247, fill: { color: C.white }, shadow: SHADOW.disc() });
  }
  for (const c of cards) {
    photoDisc(s, { x: c.photo[0], y: c.photo[1], w: 3.751, h: 3.751 });
  }
  for (const c of cards) {
    poly(s, LEAF, { x: c.leaf[0], y: c.leaf[1], w: 1.522, h: 2.044, rotate: 180, fill: { color: C.leaf } });
    priceBubble(s, { x: c.price[0], y: c.price[1], d: 2.044, fill: C.gold, was: c.price[2], now: c.price[3], wasColor: C.white, fontSize: 28 });
    text(s, 'Dev certa volle sopra anime animo qua. ', {
      x: c.caption[0], y: c.caption[1], w: 2.955, h: 1.055, align: 'center', fontFace: F.sans,
      fontSize: 20, color: C.body, lineSpacingMultiple: 1.5,
    });
  }

  text(s, 'Which Coffee You Like The Most', { x: -0.106, y: 0.884, w: 7.248, h: 1.717, align: 'center', fontSize: 48, bold: true, color: C.inkSoft });
  shape(s, 'line', { x: 2.566, y: 2.773, w: 1.905, h: 0, line: { color: C.green, width: 4 } });

  cupBadge(s, 10.004, 0.354, C.green);
  text(s, HEADER_LINE, { x: 5.378, y: 15.824, w: 5.405, h: 0.909, align: 'center', fontSize: 24 });
  scrollDots(s, C.green);
}

function slide7(pres) {
  const s = pres.addSlide();
  s.background = { color: C.cream };
  blob(s, 6.233, 14.698, 4.717, 5.321, { rotate: 90 });
  blob(s, 0.187, -0.187, 2.917, 3.29, { rotate: 270 });

  photoPoly(s, DBOWL, { x: 0.003, y: 3.82, w: 8.031, h: 12.076, flipH: true });

  text(s, 'Find Yours Best Coffee Philosophy', { x: 4.299, y: 3.59, w: 7.47, h: 1.717, fontSize: 48, bold: true, color: C.inkSoft });
  shape(s, 'line', { x: 4.494, y: 5.618, w: 1.905, h: 0, line: { color: C.green, width: 4 } });
  text(s, 'Follow Us To Get Discout %', { x: 4.494, y: 5.986, w: 4.566, h: 0.505, fontSize: 24 });
  pillButton(s, { x: 4.494, y: 6.77, w: 3.568, h: 0.771, head: 'Follow ', tail: 'Here', fontSize: 28 });

  shape(s, 'ellipse', { x: 4.877, y: 1.701, w: 0.749, h: 0.749, fill: { color: C.green } });
  shape(s, 'ellipse', { x: 5.22, y: 8.637, w: 0.375, h: 0.375, fill: { color: C.green } });

  cupBadge(s, 0.535, 17.764, C.gold);
  text(s, HEADER_LINE, { x: 1.314, y: 17.997, w: 5.405, h: 0.909, align: 'center', fontSize: 24 });
}

function slide8(pres) {
  const s = pres.addSlide();
  s.background = { color: C.cream };
  blob(s, -0.002, 0, 9.497, 7.219, { flipH: true });

  // Tablet mock-up: body, camera dots, screen, home button.
  shape(s, 'roundRect', { x: 1.908, y: 1.248, w: 7.408, h: 10.52, rectRadius: 0.44, fill: { color: C.panel }, shadow: SHADOW.device() });
  shape(s, 'ellipse', { x: 5.545, y: 1.666, w: 0.133, h: 0.133, fill: { color: C.white } });
  shape(s, 'rect', { x: 5.403, y: 1.71, w: 0.062, h: 0.062, fill: { color: C.white } });
  shape(s, 'rect', { x: 2.353, y: 2.128, w: 6.528, h: 8.706, fill: { color: C.white }, shadow: SHADOW.device() });
  shape(s, 'ellipse', { x: 5.341, y: 11.021, w: 0.551, h: 0.56, fill: { color: C.grey } });
  poly(s, HOME, { x: 5.474, y: 11.154, w: 0.285, h: 0.285, fill: { color: C.white } });
  blob(s, 6.774, 16.308, 4.485, 3.409, { rotate: 180, flipH: true });

  photoRect(s, { x: 2.362, y: 2.169, w: 6.528, h: 8.668 });

  // Caption panel sitting over the lower half of the screen.
  shape(s, 'rect', { x: 2.855, y: 9.01, w: 5.535, h: 5.948, fill: { color: C.sand } });
  text(s, 'Watch Now', { x: 2.855, y: 8.458, w: 5.535, h: 0.741, fill: { color: C.gold }, align: 'center', valign: 'middle', fontSize: 32, bold: true });
  text(s, 'BEST PHILOSOPHY', { x: 3.891, y: 9.803, w: 3.462, h: 0.505, align: 'center', fontSize: 24, bold: true });
  text(s, LOREM_LONG, {
    x: 3.317, y: 10.284, w: 4.61, h: 3.579, align: 'center', fontFace: F.sans,
    fontSize: 20, color: C.body, lineSpacingMultiple: 1.5,
  });
  diamondCaption(s, {
    label: 'Follow Us', x: 4.11, y: 14.018, w: 3.01, h: 0.538, fontSize: 20,
    diamonds: [[3.912, 14.223], [7.203, 14.223]], diamondRotate: 90,
  });

  text(s, HEADER_LINE, { x: 1.418, y: 15.85, w: 5.405, h: 0.909, fontSize: 24 });
  poly(s, BADGE, { x: 0.968, y: 15.91, w: 0.232, h: 0.25, rotate: 90, fill: { color: C.gold }, shadow: SHADOW.badge() });

  diamondCaption(s, {
    label: 'Our Best Coffee\u2019s', x: 8.649, y: 1.991, w: 3.437, h: 0.606, rotate: 90, fontSize: 24,
    diamonds: [[10.303, 0.301], [10.303, 4.028]], diamondRotate: 180,
  });
  scrollDots(s, C.green);
}

function slide9(pres) {
  const s = pres.addSlide();
  s.background = { color: C.cream };
  shape(s, 'rect', { x: -0.04, y: 14.133, w: 11.292, h: 5.584, fill: { color: C.green } });

  photoPoly(s, archOutline(8.584, 9.185), { x: 1.334, y: 4.948, w: 8.584, h: 9.185 });
  // Four hairline arches echoing the photo mask.
  const arches = [[1.884, 5.24, 7.485, 7.371], [1.43, 4.793, 8.392, 8.265], [1.025, 4.394, 9.202, 9.063], [0.635, 4.011, 9.982, 9.831]];
  for (const [x, y, w, h] of arches) {
    poly(s, archOutline(w, h), { x, y, w, h, fill: { type: 'none' }, line: { color: C.ring, width: 2.25 } });
  }

  // Map pin, drawn from a teardrop plus two ellipses.
  const pinLine = { color: C.cream, width: 4 };
  poly(s, PIN, { x: 0.692, y: 0.389, w: 0.509, h: 0.678, fill: { color: C.green }, line: pinLine });
  shape(s, 'ellipse', { x: 0.862, y: 0.53, w: 0.17, h: 0.17, fill: { color: C.green }, line: pinLine });
  shape(s, 'ellipse', { x: 0.692, y: 1.005, w: 0.509, h: 0.17, fill: { color: C.green }, line: pinLine });

  const socials = [
    { glyph: INSTAGRAM, at: [0.742, 1.557, 0.278, 0.278], label: 'instagram/soffeeshop.com', labelAt: [1.137, 1.525, 3.519] },
    { glyph: FACEBOOK, at: [0.742, 2.245, 0.29, 0.288], label: 'facebook/soffeeshop.com', labelAt: [1.143, 2.219, 3.396] },
    { glyph: TWITTER, at: [0.742, 3.007, 0.29, 0.237], label: 'twitter/soffeeshop.com', labelAt: [1.143, 2.955, 3.195] },
  ];
  for (const soc of socials) {
    poly(s, soc.glyph, { x: soc.at[0], y: soc.at[1], w: soc.at[2], h: soc.at[3], fill: { color: C.green } });
    text(s, soc.label, {
      x: soc.labelAt[0], y: soc.labelAt[1], w: soc.labelAt[2], h: 0.462,
      fontSize: 18, color: C.green, lineSpacingMultiple: 1.3,
    });
  }

  phoneBadge(s, 10.305, 0.472, 0.624);
  text(s, PHONE_NUMBER, { x: 7.247, y: 0.614, w: 2.919, h: 0.438, align: 'right', fontSize: 20 });
  text(s, HEADER_LINE, { x: 4.899, y: 2.513, w: 5.405, h: 0.909, align: 'right', fontSize: 24 });

  text(s, 'Keep in Touch', { x: 2.54, y: 14.98, w: 6.172, h: 0.909, align: 'center', fontSize: 48, bold: true, color: C.white });
  text(s, '23 North 24th Street, Canada City, CA 71539', {
    x: 1.23, y: 15.889, w: 8.791, h: 0.659, align: 'center', fontFace: F.sans,
    fontSize: 28, color: C.white, lineSpacingMultiple: 1.3,
  });
  scrollDots(s, C.gold);
}

// -------------------------------------------------------------------- build
const pres = new PptxGenJS();
pres.defineLayout({ name: 'STORY', width: DECK.w, height: DECK.h });
pres.layout = 'STORY';

[slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9].forEach(fn => fn(pres));

pres.writeFile({ fileName: path.join(__dirname, '27cefa6f-0d48-4ebd-9cfe-e09468b0c49f_grok_final.pptx') })
  .then(f => console.log('wrote', f));
