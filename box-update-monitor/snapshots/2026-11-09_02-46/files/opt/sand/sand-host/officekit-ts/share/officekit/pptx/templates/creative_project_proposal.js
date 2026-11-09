/*
 * CASELDO — Creative Project Proposal (31 slides, 13.333 x 7.5 in)
 * Standalone pptxgenjs re-creation of the reference deck.
 *
 * The reference deck's decorative line-art (globes, sparkles, flowers, icons)
 * ships as raster PNGs; here every one of them is redrawn with native
 * pptxgenjs shapes / custom geometry, so no binary asset is embedded.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Design tokens
 * ------------------------------------------------------------------ */
const BRAND = 'A2491D';          // terracotta — every stroke and letter
const PAPER = 'F3EFE7';          // warm off-white slide background
const WHITE = 'FFFFFF';

const DISPLAY = 'Raleway Black';     // big headlines
const SUBHEAD = 'Raleway SemiBold';  // section / card headings
const BODY = 'Nunito';               // running text
const BODY_BLACK = 'Nunito Black';   // emphasised numbers

const SLIDE_W = 13.333333;   // 12192000 EMU — 16:9 widescreen
const SLIDE_H = 7.5;

/* Reusable copy from the template. */
const LOREM_LONG =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
  'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.';
const LOREM_MID =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
  'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero.';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa.';
const LOREM_NUNC = 'Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus.';
const LOREM_PELL =
  'Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. ' +
  'Proin pharetra nonummy pede. Mauris et orci.';
const FILL_TEXT = 'Fill your amazing text report here, for your client impress of that';

/* ------------------------------------------------------------------ *
 * Primitive helpers
 * ------------------------------------------------------------------ */

/** Text block. Defaults mirror the deck: Nunito 14pt, brand colour, top-left. */
function text(slide, body, o) {
  slide.addText(body, Object.assign({
    fontFace: BODY, fontSize: 14, color: BRAND,
    align: 'left', valign: 'top', wrap: true, isTextBox: true,
  }, o));
}

/** Turn a list of strings into consecutive paragraphs ('' gives a blank line). */
function paras(lines) {
  return lines.map((t, i) => ({ text: t, options: { breakLine: i < lines.length - 1 } }));
}

/** Solid colour block. */
function block(slide, x, y, w, h, color, o) {
  const opts = Object.assign({ x, y, w, h, fill: { color }, line: { color, width: 0 } }, o);
  slide.addShape(opts.shape || 'rect', opts);
}

/** Outlined shape (no fill). */
function outline(slide, shape, x, y, w, h, o) {
  const opts = Object.assign({ x, y, w, h, fill: { type: 'none' } }, o);
  opts.line = Object.assign({ color: BRAND, width: 1 }, opts.line);
  slide.addShape(shape, opts);
}

/** Straight rule. */
function rule(slide, x, y, w, h, o) {
  const opts = Object.assign({ x, y, w, h }, o);
  opts.line = Object.assign({ color: BRAND, width: 1 }, opts.line);
  slide.addShape('line', opts);
}

/** The bordered banner that heads most content slides. */
function banner(slide, title, o) {
  outline(slide, 'rect', 0.375, 0.372, 12.595, 0.803, o);
  text(slide, title, {
    x: 0.375, y: 0.372, w: 12.595, h: 0.803,
    fontFace: DISPLAY, fontSize: 28, align: 'center', valign: 'middle',
  });
}

/* ------------------------------------------------------------------ *
 * Custom geometry
 *
 * Paths are stored as unit-box command lists (values 0..1) so one path can
 * be stamped at any position and size:  M/L x y | C x1 y1 x2 y2 x y | Z
 * ------------------------------------------------------------------ */

function drawPath(slide, cmds, x, y, w, h, o) {
  const pts = [];
  for (const c of cmds) {
    if (c[0] === 'M') pts.push({ x: c[1] * w, y: c[2] * h, moveTo: true });
    else if (c[0] === 'L') pts.push({ x: c[1] * w, y: c[2] * h });
    else if (c[0] === 'C') pts.push({
      x: c[5] * w, y: c[6] * h,
      curve: { type: 'cubic', x1: c[1] * w, y1: c[2] * h, x2: c[3] * w, y2: c[4] * h },
    });
    else pts.push({ close: true });
  }
  const opts = Object.assign({ x, y, w, h, points: pts, fill: { type: 'none' } }, o);
  opts.line = Object.assign({ color: BRAND, width: 1 }, opts.line);
  slide.addShape('custGeom', opts);
}

/** Vertical mirror of a unit path (y -> 1 - y). */
function mirrorY(cmds) {
  return cmds.map(c => (c[0] === 'Z' ? c : [c[0]].concat(c.slice(1).map((v, i) => (i % 2 ? 1 - v : v)))));
}

/** Four-point sparkle — the deck's signature ornament. */
const SPARK = [['M', .499, 0], ['C', .499, 0, .513, .092, .529, .191],
  ['C', .551, .335, .664, .448, .808, .47], ['C', .907, .486, .999, .5, .999, .5],
  ['C', .999, .5, .907, .514, .808, .53], ['C', .664, .552, .551, .665, .529, .809],
  ['C', .513, .908, .499, 1, .499, 1], ['C', .499, 1, .485, .908, .47, .809],
  ['C', .447, .665, .334, .552, .19, .53], ['C', .091, .514, 0, .5, 0, .5],
  ['C', 0, .5, .091, .486, .19, .47], ['C', .334, .448, .447, .335, .47, .191],
  ['C', .485, .092, .499, 0, .499, 0], ['Z']];

/** The inner sparkle: the same star with two arms clipped back to fine spikes. */
const SPARK_INNER = [['M', .562, .294], ['C', .61, .387, .7, .454, .807, .47],
  ['C', .907, .486, .999, .5, .999, .5], ['C', .999, .5, .907, .514, .807, .53],
  ['C', .663, .552, .55, .665, .528, .809], ['C', .513, .908, .499, 1, .499, 1],
  ['C', .499, 1, .484, .908, .469, .809], ['C', .453, .701, .385, .611, .293, .563],
  ['C', .341, .44, .439, .343, .562, .294], ['Z'],
  ['M', .111, .483], ['C', .119, .495, .126, .508, .133, .521], ['L', 0, .5], ['L', .111, .483], ['Z'],
  ['M', .519, .134], ['C', .506, .127, .494, .12, .481, .112], ['L', .499, 0], ['L', .519, .134], ['Z']];

/** Six-point star. */
const STAR6 = [['M', .499, 0], ['C', .499, 0, .535, .184, .56, .308],
  ['C', .573, .374, .625, .426, .691, .44], ['C', .815, .464, .999, .5, .999, .5],
  ['C', .999, .5, .815, .536, .691, .56], ['C', .625, .573, .573, .626, .56, .692],
  ['C', .535, .816, .499, 1, .499, 1], ['C', .499, 1, .463, .816, .439, .692],
  ['C', .426, .626, .374, .573, .307, .56], ['C', .183, .536, 0, .5, 0, .5],
  ['C', 0, .5, .183, .464, .307, .44], ['C', .374, .426, .426, .374, .439, .308],
  ['C', .463, .184, .499, 0, .499, 0], ['Z']];

/** Four thin diagonal spikes that nest inside the six-point star. */
const STAR_CROSS = [['M', .244, .634], ['L', 0, 1], ['L', .364, .754],
  ['C', .345, .698, .301, .654, .244, .634], ['Z'],
  ['M', .999, 1], ['L', .753, .634], ['C', .697, .654, .653, .698, .633, .754], ['L', .999, 1], ['Z'],
  ['M', .364, .245], ['L', 0, 0], ['L', .244, .365],
  ['C', .301, .346, .345, .302, .364, .245], ['Z'],
  ['M', .753, .365], ['L', .999, 0], ['L', .633, .245],
  ['C', .653, .302, .697, .346, .753, .365], ['Z']];

/** Eight-lobed "flower" outline (the deck's cloud-like motif). */
const FLOWER = [['M', .344, .123], ['C', .357, .054, .422, 0, .5, 0],
  ['C', .577, 0, .643, .054, .656, .123], ['C', .714, .084, .799, .092, .853, .146],
  ['C', .908, .201, .916, .286, .877, .344], ['C', .946, .357, 1, .422, 1, .5],
  ['C', 1, .577, .946, .643, .877, .656], ['C', .916, .714, .908, .799, .853, .853],
  ['C', .799, .908, .714, .916, .656, .877], ['C', .643, .946, .577, 1, .5, 1],
  ['C', .422, 1, .357, .946, .344, .877], ['C', .286, .916, .201, .908, .146, .853],
  ['C', .091, .799, .083, .714, .123, .656], ['C', .054, .643, 0, .577, 0, .5],
  ['C', 0, .422, .054, .357, .123, .344], ['C', .083, .286, .091, .201, .146, .146],
  ['C', .201, .092, .286, .084, .344, .123], ['Z']];

/** Lens-shaped polar cap used by the globe (the bottom copy is its mirror). */
const GLOBE_CAP = [['M', 0, .667], ['C', .112, .263, .295, 0, .5, 0],
  ['C', .705, 0, .887, .263, 1, .667], ['C', .861, .876, .688, 1, .5, 1],
  ['C', .312, 1, .139, .876, 0, .667], ['Z']];

/* ------------------------------------------------------------------ *
 * Icon library — every icon is drawn inside a square/rect of side `s`
 * ------------------------------------------------------------------ */

/** Wire-frame globe: rim, meridian, two parallels and the equator. */
function iconGlobe(slide, x, y, s, o) {
  const ln = Object.assign({ color: BRAND, width: 1 }, o && o.line);
  const capW = 0.8085 * s, capH = 0.3086 * s;
  outline(slide, 'ellipse', x, y, s, s, { line: ln });
  outline(slide, 'ellipse', x + 0.3 * s, y, 0.4 * s, s, { line: ln });
  drawPath(slide, GLOBE_CAP, x + 0.0958 * s, y, capW, capH, { line: ln });
  drawPath(slide, mirrorY(GLOBE_CAP), x + 0.0958 * s, y + 0.6914 * s, capW, capH, { line: ln });
  rule(slide, x, y + 0.5 * s, s, 0, { line: ln });
}

/** Two overlapping four-point sparkles. */
function iconSparkle(slide, x, y, s, o) {
  const ln = Object.assign({ color: BRAND, width: 1 }, o && o.line);
  drawPath(slide, SPARK, x, y, 0.96 * s, 0.96 * s, { line: ln });
  drawPath(slide, SPARK_INNER, x + 0.33 * s, y + 0.33 * s, 0.67 * s, 0.67 * s, { line: ln });
}

function iconFlower(slide, x, y, s, o) {
  drawPath(slide, FLOWER, x, y, s, s, o);
}

function iconStar6(slide, x, y, s, o) {
  drawPath(slide, STAR6, x, y, s, s, o);
}

/** Six-point star with the thin diagonal spikes nested inside it. */
function iconStarBurst(slide, x, y, s, o) {
  drawPath(slide, STAR6, x, y, s, s, o);
  drawPath(slide, STAR_CROSS, x + 0.146 * s, y + 0.146 * s, 0.707 * s, 0.707 * s, o);
}

/** Rotated variants used as loose ornaments on the "deep talk" spread. */
function spin(o, deg) {
  return Object.assign({ rotate: deg }, o);
}

/** Envelope: rounded body, V-flap and two "address" rules. */
const ENVELOPE_FLAP = [['M', 0, .172], ['C', .016, .069, .049, 0, .086, 0],
  ['C', .282, 0, .717, 0, .913, 0], ['C', .95, 0, .983, .069, 1, .172],
  ['L', .549, .959], ['C', .518, 1.013, .481, 1.013, .45, .959], ['L', 0, .172], ['Z']];

const ENVELOPE_BODY = [['M', .999, .144], ['C', .999, .064, .956, 0, .903, 0],
  ['C', .712, 0, .287, 0, .096, 0], ['C', .043, 0, 0, .064, 0, .144],
  ['C', 0, .333, 0, .666, 0, .856], ['C', 0, .935, .043, 1, .096, 1],
  ['C', .287, 1, .712, 1, .903, 1], ['C', .956, 1, .999, .935, .999, .856],
  ['C', .999, .666, .999, .333, .999, .144], ['Z']];

function iconEnvelope(slide, x, y, w, h) {
  const ln = { color: BRAND, width: 1 };
  drawPath(slide, ENVELOPE_BODY, x, y, w, h, { line: ln });
  drawPath(slide, ENVELOPE_FLAP, x + 0.012 * w, y, 0.976 * w, 0.431 * h, { line: ln });
  rule(slide, x + 0.694 * w, y + 0.741 * h, 0.225 * w, 0, { line: ln });
  rule(slide, x + 0.694 * w, y + 0.876 * h, 0.225 * w, 0, { line: ln });
}

/** Map pin on a trapezoid plinth. */
const PIN_DROP = [['M', .969, .397], ['C', 1.025, .307, 1.004, .206, .91, .126],
  ['C', .817, .047, .663, 0, .5, 0], ['C', .336, 0, .183, .047, .089, .126],
  ['C', -.004, .206, -.026, .307, .031, .397],
  ['C', .149, .584, .284, .8, .379, .951], ['C', .398, .98, .446, 1, .5, 1],
  ['C', .554, 1, .602, .98, .621, .951], ['C', .715, .8, .851, .584, .969, .397], ['Z']];
const PIN_PLINTH = [['M', .615, 0], ['L', .55, .4], ['C', .542, .448, .522, .48, .5, .48],
  ['C', .478, .48, .457, .448, .45, .4], ['L', .384, 0], ['L', .122, 0],
  ['C', .086, 0, .056, .059, .05, .139], ['C', .037, .313, .015, .61, .001, .806],
  ['C', -.003, .854, .003, .904, .017, .941], ['C', .031, .979, .052, 1, .073, 1],
  ['C', .267, 1, .733, 1, .927, 1], ['C', .948, 1, .969, .979, .983, .941],
  ['C', .996, .904, 1.002, .854, .999, .806], ['C', .985, .61, .963, .313, .95, .139],
  ['C', .944, .059, .914, 0, .878, 0], ['L', .615, 0], ['Z']];

function iconPin(slide, x, y, w, h) {
  const ln = { color: BRAND, width: 1 };
  drawPath(slide, PIN_DROP, x + 0.29 * w, y, 0.417 * w, 0.755 * h, { line: ln });
  outline(slide, 'ellipse', x + 0.383 * w, y + 0.102 * h, 0.233 * w, 0.249 * h, { line: ln });
  drawPath(slide, PIN_PLINTH, x, y + 0.529 * h, w, 0.471 * h, { line: ln });
  rule(slide, x + 0.033 * w, y + 0.697 * h, 0.41 * w, 0, { line: ln });
  rule(slide, x + 0.557 * w, y + 0.697 * h, 0.41 * w, 0, { line: ln });
}

/** Clock: rim, inner ring and two hands. */
function iconClock(slide, x, y, s) {
  const ln = { color: BRAND, width: 1 };
  outline(slide, 'ellipse', x, y, s, s, { line: ln });
  outline(slide, 'ellipse', x + 0.13 * s, y + 0.13 * s, 0.7375 * s, 0.7375 * s, { line: ln });
  drawPath(slide, [['M', 0, .844], ['C', 0, .93, .086, 1, .192, 1], ['L', .884, .938],
    ['C', .948, .938, 1, .896, 1, .844], ['C', 1, .792, .948, .749, .884, .749],
    ['L', .37, .704], ['L', .309, .094], ['C', .309, .042, .256, 0, .192, 0],
    ['C', .128, 0, .076, .042, .076, .094], ['C', .076, .094, 0, .844, 0, .844], ['Z']],
  x + 0.43 * s, y + 0.218 * s, 0.29 * s, 0.355 * s, { line: ln });
}

/** Telephone handset with two signal arcs. */
const HANDSET = [['M', .72, .329], ['C', .72, .374, .713, .418, .7, .45],
  ['C', .687, .482, .67, .5, .651, .5], ['C', .569, .5, .431, .5, .348, .5],
  ['C', .33, .5, .313, .482, .3, .45], ['C', .287, .418, .28, .374, .28, .329],
  ['C', .28, .262, .28, .2, .28, .2], ['C', .28, .147, .271, .096, .256, .059],
  ['C', .241, .021, .221, 0, .2, 0], ['C', .162, 0, .118, 0, .08, 0],
  ['C', .059, 0, .038, .021, .023, .059], ['C', .008, .096, 0, .147, 0, .2],
  ['L', 0, .315], ['C', 0, .497, .029, .671, .08, .8], ['C', .132, .928, .201, 1, .274, 1],
  ['C', .415, 1, .585, 1, .726, 1], ['C', .799, 1, .868, .928, .92, .8],
  ['C', .971, .671, 1, .497, 1, .315], ['C', 1, .246, 1, .2, 1, .2],
  ['C', 1, .147, .991, .096, .976, .059], ['C', .961, .021, .941, 0, .92, 0],
  ['C', .882, 0, .838, 0, .8, 0], ['C', .779, 0, .758, .021, .743, .059],
  ['C', .728, .096, .72, .147, .72, .2], ['L', .72, .329], ['Z']];
const SIGNAL_ARC = [['M', 0, 1], ['C', .081, .413, .274, 0, .5, 0], ['C', .725, 0, .919, .413, 1, 1]];

function iconPhone(slide, x, y, w, h) {
  const ln = { color: BRAND, width: 1 };
  drawPath(slide, HANDSET, x, y + 0.481 * h, w, 0.519 * h, { line: ln, rotate: 45 });
  drawPath(slide, SIGNAL_ARC, x + 0.34 * w, y + 0.163 * h, 0.249 * w, 0.087 * h, { line: ln, rotate: 45 });
  drawPath(slide, SIGNAL_ARC, x + 0.214 * w, y, 0.498 * w, 0.217 * h, { line: ln, rotate: 45 });
}

/**
 * Speech bubble + avatar. Appears twice on "Work in progress", the second
 * copy mirrored horizontally, so every x is expressed through `at()`.
 */
const CHAT_BUBBLE = [['M', .955, 0], ['L', .15, 0],
  ['C', .11, 0, .072, .022, .044, .061], ['C', .015, .1, 0, .153, 0, .208],
  ['L', 0, .792], ['C', 0, .847, .015, .9, .044, .939], ['C', .072, .978, .11, 1, .15, 1],
  ['L', .692, 1], ['C', .732, 1, .77, .978, .799, .939], ['C', .827, .9, .843, .847, .843, .792],
  ['L', .843, .481], ['C', .843, .365, .876, .254, .936, .173],
  ['C', .953, .15, .97, .127, .986, .105], ['C', .999, .088, 1.003, .061, .996, .038],
  ['C', .99, .015, .973, 0, .955, 0], ['Z']];
const AVATAR_BODY = [['M', .693, 0], ['C', .83, .091, .94, .274, .999, .506],
  ['C', .897, .802, .711, 1, .499, 1], ['C', .288, 1, .102, .802, 0, .506],
  ['C', .059, .274, .169, .091, .305, 0], ['C', .351, .1, .421, .165, .499, .165],
  ['C', .578, .165, .648, .1, .693, 0], ['Z']];

function iconChatPerson(slide, x, y, flip) {
  const ln = { color: BRAND, width: 1.5 };
  /* Mirror helper: maps a left-anchored offset+width to the flipped position. */
  const at = (dx, dw) => (flip ? x + 1.807 - dx - dw : x + dx);
  drawPath(slide, CHAT_BUBBLE, at(0, 1.122), y, 1.122, 0.811, { line: ln, flipH: !!flip });
  [[0.274, 0.489, 0.188], [0.182, 0.581, 0.333], [0.182, 0.581, 0.479], [0.182, 0.441, 0.624]]
    .forEach(([dx, lw, dy]) => rule(slide, at(dx, lw), y + dy, lw, 0, { line: ln }));
  outline(slide, 'ellipse', at(1.075, 0.732), y + 0.079, 0.732, 0.732, { line: ln });
  outline(slide, 'ellipse', at(1.287, 0.31), y + 0.205, 0.31, 0.31, { line: ln });
  drawPath(slide, AVATAR_BODY, at(1.129, 0.625), y + 0.456, 0.625, 0.355, { line: ln });
}

/** Projector on a stand — the "presentations" mark. */
const PJ_TOP = [['M', .081, .889], ['C', .081, .889, .17, .931, .291, .962],
  ['C', .355, .979, .427, .997, .5, .998], ['C', .542, .999, .585, .99, .625, .981],
  ['C', .786, .944, .919, .889, .919, .889], ['C', .963, .889, 1, .714, 1, .498],
  ['C', 1, .283, .963, .108, .919, .108], ['C', .919, .108, .786, .053, .625, .016],
  ['C', .585, .007, .542, -.002, .5, -.001], ['C', .427, 0, .355, .018, .291, .035],
  ['C', .17, .066, .081, .108, .081, .108], ['C', .036, .108, 0, .283, 0, .498],
  ['C', 0, .714, .036, .889, .081, .889], ['Z']];
const PJ_BODY = [['M', 0, .999], ['L', 1, .999], ['L', 1, .034],
  ['C', 1, .022, .999, .011, .998, 0], ['C', .994, 0, .991, .001, .987, .001],
  ['C', .987, .001, .832, .027, .646, .045], ['C', .599, .049, .549, .053, .5, .053],
  ['C', .415, .052, .331, .044, .257, .036], ['C', .116, .021, .013, .001, .013, .001],
  ['C', .009, .001, .005, 0, .001, 0], ['C', 0, .011, 0, .022, 0, .034], ['L', 0, .999], ['Z']];
const PJ_BASE = [['M', 1, .07], ['C', 1, .051, .999, .032, .997, .019],
  ['C', .995, .005, .993, 0, .99, 0], ['C', .892, 0, .107, 0, .01, 0],
  ['C', .007, 0, .005, .005, .003, .019], ['C', .001, .032, 0, .051, 0, .07],
  ['C', 0, .15, 0, .27, 0, .27], ['C', 0, .672, .045, .998, .101, .998],
  ['C', .293, .998, .707, .998, .899, .998], ['C', .955, .998, 1, .672, 1, .27], ['L', 1, .07], ['Z']];
const PJ_PLATE = [['M', 1, 0], ['C', .759, 0, .241, 0, 0, 0], ['L', 0, .321],
  ['C', 0, .694, .032, .997, .071, .997], ['C', .255, .997, .745, .997, .929, .997],
  ['C', .968, .997, 1, .694, 1, .321], ['L', 1, 0], ['Z']];
const PJ_LEG = [['M', .999, 0], ['C', .999, .002, .999, .004, .999, .004],
  ['C', .999, .004, .962, .411, .622, .728], ['C', .512, .831, .373, .927, .196, .998],
  ['C', .193, .88, .117, .774, 0, .704], ['C', .088, .652, .157, .59, .212, .526],
  ['C', .402, .303, .403, .04, .401, 0], ['L', .999, 0], ['Z']];
const PJ_HANDLE = [['M', .063, .893], ['C', .063, .893, .156, .931, .281, .961],
  ['C', .348, .977, .424, .995, .499, .996], ['C', .544, .997, .588, .989, .63, .981],
  ['C', .798, .947, .936, .893, .936, .893], ['C', .971, .893, .999, .715, .999, .497],
  ['C', .999, .278, .971, .101, .936, .101], ['C', .936, .101, .798, .046, .63, .012],
  ['C', .588, .004, .544, -.004, .499, -.003], ['C', .424, -.002, .348, .016, .281, .032],
  ['C', .156, .062, .063, .101, .063, .101], ['C', .028, .101, 0, .278, 0, .497],
  ['C', 0, .715, .028, .893, .063, .893], ['Z']];

function iconProjector(slide, x, y) {
  const ln = { color: BRAND, width: 1.5 };
  drawPath(slide, PJ_TOP, x, y, 1.788, 0.372, { line: ln });
  drawPath(slide, PJ_BODY, x + 0.125, y + 0.331, 1.538, 0.78, { line: ln });
  drawPath(slide, PJ_BASE, x + 0.074, y + 1.11, 1.639, 0.227, { line: ln });
  drawPath(slide, PJ_PLATE, x + 0.238, y + 1.338, 1.311, 0.138, { line: ln });
  drawPath(slide, PJ_LEG, x + 0.95, y + 1.475, 0.16, 0.223, { line: ln });
  drawPath(slide, PJ_LEG, x + 0.677, y + 1.475, 0.16, 0.223, { line: ln, flipH: true });
  drawPath(slide, PJ_HANDLE, x + 0.391, y + 0.106, 1.004, 0.161, { line: ln });
  outline(slide, 'ellipse', x + 0.806, y + 1.613, 0.176, 0.176, { line: ln });
  [[0.469, 0.536, 0.459], [0.791, 0.464, 0.531], [1.113, 0.609, 0.386]]
    .forEach(([dx, dy, bh]) => outline(slide, 'roundRect', x + dx, y + dy, 0.205, bh,
      { rectRadius: 0.05, line: ln }));
}

/** Banknote: rounded frame, dollar sign and four corner arcs. */
const BILL_FRAME = [['M', .999, .056], ['C', .999, .025, .987, 0, .971, 0],
  ['C', .828, 0, .17, 0, .027, 0], ['C', .011, 0, 0, .025, 0, .056],
  ['C', 0, .232, 0, .768, 0, .943], ['C', 0, .975, .011, 1, .027, 1],
  ['C', .17, 1, .828, 1, .971, 1], ['C', .987, 1, .999, .975, .999, .943],
  ['C', .999, .768, .999, .232, .999, .056], ['Z']];
const BILL_CORNERS = [['M', 0, .697], ['C', .116, .697, .211, .833, .211, 1],
  ['L', .027, 1], ['C', .011, 1, 0, .975, 0, .943], ['L', 0, .697], ['Z'],
  ['M', .999, .697], ['L', .999, .943], ['C', .999, .975, .987, 1, .971, 1],
  ['L', .787, 1], ['C', .787, .833, .882, .697, .999, .697], ['Z'],
  ['M', .211, 0], ['C', .211, .167, .116, .303, 0, .303], ['L', 0, .056],
  ['C', 0, .025, .011, 0, .027, 0], ['L', .211, 0], ['Z'],
  ['M', .787, 0], ['L', .971, 0], ['C', .987, 0, .999, .025, .999, .056],
  ['L', .999, .303], ['C', .882, .303, .787, .167, .787, 0], ['Z']];
const DOLLAR_S = [['M', .066, .881], ['C', .154, .952, .312, 1, .492, 1],
  ['C', .768, 1, .992, .888, .992, .75], ['C', .992, .612, .768, .5, .492, .5],
  ['C', .216, .5, -.008, .388, -.008, .25], ['C', -.008, .112, .216, 0, .492, 0],
  ['C', .7, 0, .878, .063, .953, .153]];

function iconMoney(slide, x, y, w, h) {
  const ln = { color: BRAND, width: 1 };
  drawPath(slide, BILL_FRAME, x, y, w, h, { line: ln });
  drawPath(slide, BILL_CORNERS, x, y, w, h, { line: ln });
  outline(slide, 'ellipse', x + 0.355 * w, y + 0.12 * h, 0.288 * w, 0.761 * h, { line: ln });
  drawPath(slide, DOLLAR_S, x + 0.443 * w, y + 0.267 * h, 0.113 * w, 0.465 * h, { line: ln });
  rule(slide, x + 0.5 * w, y + 0.225 * h, 0, 0.042 * h, { line: ln });
  rule(slide, x + 0.5 * w, y + 0.733 * h, 0, 0.042 * h, { line: ln });
  rule(slide, x + 0.706 * w, y + 0.501 * h, 0.061 * w, 0, { line: ln });
  rule(slide, x + 0.232 * w, y + 0.501 * h, 0.061 * w, 0, { line: ln });
}

/** Open cardboard box with four flaps. */
const BOX_LID = [['M', 0, .453], ['C', 0, .453, .452, .023, .476, 0],
  ['C', .477, 0, .478, 0, .479, 0], ['C', .504, .022, 1, .453, 1, .453],
  ['L', .566, .999], ['L', 0, .453], ['Z']];
const BOX_FLAP_L = [['M', 1, .211], ['L', .145, 0],
  ['C', .145, 0, .07, .276, .012, .492], ['C', -.003, .548, -.005, .617, .009, .676],
  ['C', .022, .734, .048, .775, .079, .786], ['C', .256, .846, .561, .949, .702, .997],
  ['C', .748, 1.012, .793, .967, .816, .883], ['C', .88, .651, 1, .211, 1, .211], ['Z']];
const BOX_FLAP_R = [['M', 1, .22], ['L', .171, 0],
  ['C', .171, 0, .081, .296, .013, .522], ['C', -.004, .58, -.005, .649, .012, .708],
  ['C', .028, .767, .061, .809, .098, .821], ['C', .265, .872, .521, .952, .662, .996],
  ['C', .724, 1.015, .785, .962, .813, .865], ['C', .881, .629, 1, .22, 1, .22], ['Z']];
const BOX_BODY = [['M', .664, .303], ['C', .679, .348, .711, .373, .743, .364],
  ['L', 1, .293], ['L', 1, .77], ['C', 1, .824, .975, .87, .941, .881],
  ['C', .833, .916, .609, .987, .572, .999], ['C', .568, 1, .565, 1, .561, .999],
  ['C', .516, .988, .198, .911, .062, .878], ['C', .026, .869, 0, .821, 0, .765],
  ['L', 0, .292], ['L', .369, .384], ['C', .399, .392, .43, .37, .445, .328],
  ['C', .487, .215, .566, 0, .566, 0], ['L', .664, .303], ['Z']];
const BOX_LABEL = [['M', 0, 0], ['L', .998, .131], ['L', .998, 1],
  ['C', .998, 1, .115, .846, .115, .846], ['C', .048, .834, 0, .77, 0, .695], ['L', 0, 0], ['Z']];

function iconBox(slide, x, y) {
  const ln = { color: BRAND, width: 1 };
  drawPath(slide, BOX_LID, x + 0.116, y, 1.19, 0.145, { line: ln });
  drawPath(slide, BOX_FLAP_L, x, y + 0.066, 0.789, 0.375, { line: ln });
  drawPath(slide, BOX_FLAP_R, x + 0.79, y + 0.066, 0.622, 0.36, { line: ln, flipH: true });
  drawPath(slide, BOX_BODY, x + 0.116, y + 0.145, 1.19, 0.766, { line: ln });
  drawPath(slide, BOX_LABEL, x + 0.168, y + 0.575, 0.251, 0.226, { line: ln });
  rule(slide, x + 0.221, y + 0.701, 0.137, 0.019, { line: ln });
  rule(slide, x + 0.221, y + 0.648, 0.137, 0.019, { line: ln });
  rule(slide, x + 0.79, y + 0.145, 0.006, 0.766, { line: ln });
}

/** Person under three rounded stars — the customer-satisfaction mark. */
const STAR5 = [['M', .453, .027], ['C', .463, .01, .481, 0, .5, 0],
  ['C', .519, 0, .537, .01, .547, .027], ['C', .592, .1, .656, .205, .683, .248],
  ['C', .69, .26, .702, .269, .716, .273], ['C', .763, .285, .878, .317, .958, .339],
  ['C', .977, .344, .991, .359, .997, .378], ['C', 1.003, .397, .999, .418, .987, .433],
  ['C', .935, .5, .859, .596, .829, .636], ['C', .82, .647, .815, .662, .816, .676],
  ['C', .819, .727, .826, .851, .83, .938], ['C', .832, .958, .823, .977, .807, .989],
  ['C', .792, 1.001, .772, 1.003, .754, .996], ['C', .677, .964, .566, .919, .52, .901],
  ['C', .507, .895, .493, .895, .48, .901], ['C', .434, .919, .323, .964, .246, .996],
  ['C', .228, 1.003, .208, 1.001, .193, .989], ['C', .177, .977, .168, .958, .17, .938],
  ['C', .174, .851, .181, .727, .184, .676], ['C', .185, .662, .18, .647, .171, .636],
  ['C', .141, .596, .065, .5, .013, .433], ['C', .001, .418, -.003, .397, .003, .378],
  ['C', .009, .359, .023, .344, .042, .339], ['C', .122, .317, .237, .285, .284, .273],
  ['C', .298, .269, .31, .26, .317, .248], ['C', .344, .205, .408, .1, .453, .027], ['Z']];
const PERSON = [['M', .253, 0], ['C', .307, .089, .397, .148, .5, .148],
  ['C', .602, .148, .693, .089, .746, 0], ['L', .776, .012],
  ['C', .934, .081, 1.025, .268, .994, .457], ['C', .973, .581, .95, .717, .932, .822],
  ['C', .915, .925, .836, 1, .744, 1], ['L', .255, 1], ['C', .163, 1, .084, .925, .066, .822],
  ['C', .049, .718, .027, .584, .006, .461], ['C', -.026, .27, .067, .081, .226, .011],
  ['L', .253, 0], ['Z']];

function iconRating(slide, x, y) {
  const ln = { color: BRAND, width: 1.5 };
  drawPath(slide, STAR5, x + 0.598, y, 0.755, 0.722, { line: ln });
  drawPath(slide, STAR5, x, y + 0.468, 0.584, 0.559, { line: ln });
  drawPath(slide, STAR5, x + 1.368, y + 0.468, 0.584, 0.559, { line: ln });
  outline(slide, 'ellipse', x + 0.696, y + 0.895, 0.56, 0.56, { line: ln });
  drawPath(slide, PERSON, x + 0.508, y + 1.333, 0.936, 0.824, { line: ln });
  rule(slide, x + 1.134, y + 1.683, 0.168, 0.013, { line: ln });
}

/** Two overlapping price tags, tilted 30° like the reference. */
const TAG = [['M', .296, .091], ['C', .315, .039, .399, 0, .5, 0],
  ['C', .601, 0, .685, .039, .704, .091], ['L', .833, .091],
  ['C', .877, .091, .92, .101, .951, .118], ['C', .982, .135, 1, .158, 1, .182],
  ['C', 1, .356, 1, .735, 1, .909], ['C', 1, .933, .982, .956, .951, .973],
  ['C', .92, .99, .877, 1, .833, 1], ['C', .648, 1, .352, 1, .167, 1],
  ['C', .122, 1, .08, .99, .049, .973], ['C', .017, .956, 0, .933, 0, .909],
  ['C', 0, .735, 0, .356, 0, .182], ['C', 0, .158, .017, .135, .049, .118],
  ['C', .08, .101, .122, .091, .167, .091], ['L', .296, .091], ['Z']];
const TAG_STRING = [['M', 1, 0], ['C', 1, 0, .338, .543, .024, .8],
  ['C', -.017, .834, -.006, .87, .053, .901], ['C', .112, .932, .213, .954, .331, .961],
  ['C', .645, .979, 1, 1, 1, 1]];
const TAG_RING = [['M', .227, .938], ['C', .09, .847, 0, .689, 0, .51],
  ['C', 0, .229, .224, 0, .5, 0], ['C', .776, 0, 1, .229, 1, .51],
  ['C', 1, .742, .848, .938, .641, 1]];
const TAG_SLOT = [['M', 1, .158], ['C', 1, .116, .992, .076, .978, .047],
  ['C', .964, .017, .944, 0, .924, 0], ['C', .739, 0, .261, 0, .075, 0],
  ['C', .055, 0, .036, .017, .022, .047], ['C', .008, .076, 0, .116, 0, .158],
  ['C', 0, .345, 0, .655, 0, .842], ['C', 0, .884, .008, .924, .022, .954],
  ['C', .036, .984, .055, 1, .075, 1], ['C', .261, 1, .739, 1, .924, 1],
  ['C', .944, 1, .964, .984, .978, .954], ['C', .992, .924, 1, .884, 1, .842],
  ['C', 1, .655, 1, .345, 1, .158], ['Z']];

function iconPriceTag(slide, x, y) {
  const ln = { color: BRAND, width: 2.25 };
  const R = { line: ln, rotate: 30 };
  drawPath(slide, TAG_STRING, x + 0.453, y + 0.45, 0.58, 1.968, R);
  drawPath(slide, TAG, x + 0.919, y + 0.471, 1.583, 2.902, R);
  drawPath(slide, TAG_SLOT, x + 0.868, y + 2.507, 0.791, 0.378, R);
  drawPath(slide, TAG_RING, x + 1.978, y + 0.148, 0.801, 0.784, R);
  drawPath(slide, [['M', 0, 0], ['L', 1, 0]], x + 0.732, y + 2.245, 1.583, 0.021, R);
  drawPath(slide, [['M', 0, 0], ['L', 1, 0]], x + 1.531, y + 1.346, 0.33, 0.021, R);
  drawPath(slide, [['M', 0, 0], ['L', 1, 0]], x + 1.311, y + 1.771, 0.66, 0.021, R);
  outline(slide, 'ellipse', x + 2.163, y + 0.83, 0.226, 0.226, R);
}

/** Smartphone outline with a notch — the "follow us" mock-up. */
const PHONE_SHELL = [['M', 1, .066], ['C', 1, .048, .986, .032, .96, .019],
  ['C', .934, .007, .899, 0, .862, 0], ['C', .631, 0, .37, 0, .138, 0],
  ['C', .102, 0, .067, .007, .041, .019], ['C', .015, .032, 0, .048, 0, .066],
  ['C', 0, .278, 0, .722, 0, .934], ['C', 0, .951, .015, .968, .041, .981],
  ['C', .067, .993, .102, 1, .138, 1], ['C', .37, 1, .631, 1, .862, 1],
  ['C', .899, 1, .934, .993, .96, .981], ['C', .986, .968, 1, .951, 1, .934],
  ['C', 1, .722, 1, .278, 1, .066], ['Z']];
const PHONE_NOTCH = [['M', .999, 0], ['L', 0, 0], ['C', 0, 0, .006, .208, .014, .435],
  ['C', .025, .765, .067, .999, .115, .999], ['C', .292, .999, .707, .999, .884, .999],
  ['C', .932, .999, .974, .765, .985, .435], ['C', .993, .208, .999, 0, .999, 0], ['Z']];

function iconPhoneMockup(slide, x, y, w, h) {
  drawPath(slide, PHONE_SHELL, x, y, w, h, { line: { color: '000000', width: 5.7 } });
  drawPath(slide, PHONE_NOTCH, x + 0.236 * w, y, 0.528 * w, 0.036 * h,
    { fill: { color: '000000' }, line: { type: 'none' } });
}

/**
 * Picture frame. The template ships every one of these placeholders empty, so
 * they contribute nothing to the printed page; the call still records where a
 * photograph belongs and at what size, and no bitmap is embedded.
 */
function photo(slide, x, y, w, h, o) {
  const opts = Object.assign({ x, y, w, h, fill: { type: 'none' }, line: { type: 'none' } }, o);
  slide.addShape(opts.shape || 'rect', opts);
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

/* 1 — Title. */
function slide01(s) {
  photo(s, 4.518, 2.533, 4.298, 4.967, { shape: 'round2SameRect' });
  outline(s, 'rect', 0.375, 0.372, 12.583, 0.803);
  text(s, 'CASELDO PRESENTATION PROJECT PROPOSAL', {
    x: 0.375, y: 0.372, w: 12.583, h: 0.803,
    fontFace: BODY_BLACK, fontSize: 18, align: 'center', valign: 'middle',
  });
  text(s, 'CASELDO', { x: 4.434, y: 1.248, w: 4.464, h: 1.212, fontFace: DISPLAY, fontSize: 66 });
  text(s, 'PRESENTATION TEMPLATE', { x: 0.285, y: 4.824, w: 1.676, h: 0.572 });
  text(s, 'CREATIVE PROJECT PRESENTATION', { x: 10.994, y: 4.824, w: 2.055, h: 0.572, align: 'right' });
  iconGlobe(s, 2.908, 0.563, 0.421);
  iconGlobe(s, 10.005, 0.563, 0.421);
  iconFlower(s, 0.875, 4.26, 0.496);
  iconStarBurst(s, 11.773, 4.26, 0.496);
  text(s, 'CASELDO CREATIVE PROJECT PROPOSAL PRESENTATION 2021',
    { x: 3.498, y: 6.791, w: 6.337, h: 0.337, align: 'center' });
}

/* 2 — Section intro with a boxed title and a sparkle badge. */
function slide02(s) {
  outline(s, 'rect', 0.375, 0.372, 10.324, 0.803);
  text(s, 'PRESENTATION INFORMATION', {
    x: 0.375, y: 0.372, w: 10.324, h: 0.803,
    fontFace: DISPLAY, fontSize: 28, align: 'center', valign: 'middle',
  });
  outline(s, 'rect', 10.961, 0.372, 2.062, 0.803);
  iconSparkle(s, 11.69, 0.472, 0.603);
  text(s, paras(['CASELDO WELCOME INTERFACE', 'PRESENTATION TEMPLATE']), {
    x: 2.165, y: 3.026, w: 9.003, h: 1.447, fontFace: DISPLAY, fontSize: 40, align: 'center',
  });
  text(s, 'you can share experiences with brands you work with',
    { x: 3.498, y: 6.157, w: 6.337, h: 0.337, align: 'center' });
}

/* 3 — Our company. */
function slide03(s) {
  iconFlower(s, 2.256, 2.669, 0.83);
  iconFlower(s, 10.248, 2.669, 0.83);
  text(s, 'OUR COMPANY', { x: 4.252, y: 0.37, w: 4.83, h: 0.774, fontFace: DISPLAY, fontSize: 40, align: 'center' });
  text(s, paras([LOREM_LONG, '', LOREM_PELL]), { x: 2.129, y: 5.196, w: 9.074, h: 1.515, align: 'center' });
  photo(s, 3.469, 1.333, 6.396, 3.502);
}

/* 4 — Introduction; full-height photo panel on the left. */
function slide04(s) {
  text(s, 'INTRODUCTION', { x: 5.458, y: 2.034, w: 5.051, h: 0.774, fontFace: DISPLAY, fontSize: 40 });
  text(s, paras([LOREM_LONG, '', LOREM_PELL]),
    { x: 5.458, y: 2.997, w: 7.41, h: 1.661, fontSize: 12, lineSpacingMultiple: 1.3 });
  iconFlower(s, 5.66, 0.408, 1.243);
  photo(s, 0.465, 0.408, 4.569, 6.684);
}

/* 5 — About us: white card spanning the slide, text left, photo right. */
function slide05(s) {
  block(s, 0.557, 0.631, 12.22, 6.238, WHITE);
  text(s, 'ABOUT US', { x: 0.75, y: 0.934, w: 3.386, h: 0.774, fontFace: DISPLAY, fontSize: 40 });
  text(s, paras([LOREM_LONG, '', LOREM_NUNC, '', LOREM_PELL]), { x: 0.75, y: 2.011, w: 4.667, h: 2.928 });
  photo(s, 5.884, 0.63, 6.892, 6.24);
}

/* 6 — Moodboard. */
function slide06(s) {
  text(s, 'MOODBOARD', { x: 0.557, y: 2.06, w: 4.29, h: 0.774, fontFace: DISPLAY, fontSize: 40 });
  text(s, paras([LOREM_LONG, '', LOREM_NUNC, '', LOREM_PELL]), { x: 0.557, y: 2.983, w: 5.18, h: 2.457 });
  photo(s, 7.981, 0.0, 4.873, 7.5);
  block(s, 6.243, 2.01, 3.825, 3.481, WHITE);
  photo(s, 6.417, 2.168, 3.476, 3.163);
}

/* 7 — Product sales. */
function slide07(s) {
  block(s, 5.983, 2.423, 7.351, 4.293, WHITE);
  photo(s, 0.0, 1.139, 5.983, 5.576);
  text(s, 'PRODUCT SALES', { x: 6.272, y: 1.134, w: 5.316, h: 0.774, fontFace: DISPLAY, fontSize: 40 });
  text(s, 'Best products sold', { x: 6.272, y: 2.68, w: 3.706, h: 0.505, fontFace: SUBHEAD, fontSize: 24 });
  text(s, paras([LOREM_LONG, '', LOREM_NUNC, '', LOREM_PELL]),
    { x: 6.272, y: 3.372, w: 6.501, h: 1.717, fontSize: 12 });
}

/* 8 — Left side / right side, photo column down the middle. */
function slide08(s) {
  block(s, 4.282, 0.0, 4.77, 7.5, WHITE);
  photo(s, 4.557, 1.398, 4.22, 4.704);
  text(s, 'LEFT SIDE', { x: 0.585, y: 2.614, w: 2.467, h: 0.572, fontFace: DISPLAY, fontSize: 28 });
  text(s, LOREM_LONG, { x: 0.585, y: 3.225, w: 3.212, h: 1.661, fontSize: 12, lineSpacingMultiple: 1.3 });
  text(s, 'RIGHT SIDE', { x: 9.949, y: 2.614, w: 2.8, h: 0.572, fontFace: DISPLAY, fontSize: 28, align: 'right' });
  text(s, LOREM_LONG,
    { x: 9.537, y: 3.225, w: 3.212, h: 1.661, fontSize: 12, align: 'right', lineSpacingMultiple: 1.3 });
}

/* 9 — What are beauty: outlined circle left, copy right. */
function slide09(s) {
  photo(s, 0.455, 1.558, 4.161, 4.161, { shape: 'ellipse' });
  text(s, 'WHAT ARE BEAUTY', { x: 5.934, y: 2.164, w: 5.936, h: 0.774, fontFace: DISPLAY, fontSize: 40 });
  text(s,
    'Beauty is commonly described as a feature of objects that makes these objects pleasurable to perceive. ' +
    'Such objects include landscapes, sunsets, humans and works of art. Beauty, together with art and taste, ' +
    'is the main subject of aesthetics, one of the major branches of philosophy. As a positive aesthetic value, ' +
    'it is contrasted with ugliness as its negative counterpart. Along with truth and goodness it is one of the ' +
    'transcendentals, which are often considered the three fundamental concepts of human understanding.',
    { x: 4.862, y: 3.191, w: 8.08, h: 1.921, align: 'center', lineSpacingMultiple: 1.3 });
  outline(s, 'ellipse', 0.455, 1.558, 4.161, 4.161);
}

/* 10 — Purpose of life: mirror of slide 9. */
function slide10(s) {
  photo(s, 8.785, 1.559, 4.161, 4.158, { shape: 'ellipse' });
  text(s, 'PURPOSE OF LIFE',
    { x: 1.459, y: 2.582, w: 5.936, h: 0.774, fontFace: DISPLAY, fontSize: 40, align: 'center' });
  text(s,
    'We are constituted so that simple acts of kindness, such as giving to charity or expressing gratitude, ' +
    'have a positive effect on our long-term moods. The key to the happy life, it seems, is the good life: ' +
    'a life with sustained relationships, challenging work, and connections to community.',
    { x: 0.387, y: 3.61, w: 8.08, h: 1.308, align: 'center', lineSpacingMultiple: 1.3 });
  outline(s, 'ellipse', 8.785, 1.558, 4.161, 4.161);
}

/* 11 — Founder profile. */
function slide11(s) {
  text(s, 'ZAYN BROWN', { x: 0.379, y: 2.107, w: 3.242, h: 0.572, fontFace: SUBHEAD, fontSize: 28 });
  text(s, LOREM_SHORT, { x: 0.379, y: 2.679, w: 3.339, h: 0.873, fontSize: 12, lineSpacingMultiple: 1.3 });
  photo(s, 4.068, 0.368, 5.196, 6.377);
  text(s, 'FOUNDER OF CASELDO',
    { x: 3.096, y: 6.358, w: 7.141, h: 0.774, fontFace: DISPLAY, fontSize: 40, align: 'center' });
  text(s, 'HISTORY', { x: 9.613, y: 3.675, w: 3.242, h: 0.572, fontFace: SUBHEAD, fontSize: 28 });
  text(s, LOREM_SHORT, { x: 9.613, y: 4.247, w: 3.339, h: 0.873, fontSize: 12, lineSpacingMultiple: 1.3 });
  iconSparkle(s, 9.809, 2.801, 0.79);
  iconGlobe(s, 1.632, 1.138, 0.737);
}

/* 12 — Gallery album. */
function slide12(s) {
  outline(s, 'roundRect', 0.919, 0.662, 2.957, 6.177, { rectRadius: 0.358, line: { width: 1.5 } });
  text(s, 'GALLERY ALBUM', { x: 4.334, y: 2.353, w: 5.233, h: 0.774, fontFace: DISPLAY, fontSize: 40 });
  text(s, paras([LOREM_LONG, '', LOREM_PELL]),
    { x: 4.334, y: 3.226, w: 8.08, h: 1.921, lineSpacingMultiple: 1.3 });
  photo(s, 1.017, 0.759, 2.76, 5.983, { shape: 'roundRect', rectRadius: 0.334 });
  iconSparkle(s, 9.149, 1.816, 0.875);
}

/* 13 — Portfolio. */
function slide13(s) {
  iconGlobe(s, 2.029, 0.604, 1.192, { line: { width: 1.5 } });
  text(s, 'PORTOFOLIO', { x: 0.55, y: 1.769, w: 4.151, h: 0.774, fontFace: DISPLAY, fontSize: 40 });
  text(s, paras([LOREM_LONG, '', LOREM_PELL]),
    { x: 4.87, y: 0.622, w: 7.913, h: 1.921, lineSpacingMultiple: 1.3 });
  photo(s, 0.55, 2.771, 12.233, 4.198);
}

/* 14 — Follow us, with a phone mock-up. */
function slide14(s) {
  photo(s, 9.009, 0.46, 3.144, 6.58, { shape: 'roundRect', rectRadius: 0.456 });
  iconPhoneMockup(s, 9.009, 0.46, 3.143, 6.566);
  text(s, 'FOLLOW US', { x: 0.386, y: 0.693, w: 3.817, h: 0.774, fontFace: DISPLAY, fontSize: 40 });
  text(s, paras([LOREM_LONG, '', LOREM_PELL]),
    { x: 0.408, y: 1.632, w: 7.913, h: 1.921, lineSpacingMultiple: 1.3 });
  iconGlobe(s, 0.512, 3.821, 0.354);
  iconEnvelope(s, 1.075, 3.861, 0.408, 0.274);
}

/* 15 — Team: four square portraits above the strapline. */
function slide15(s) {
  [0.514, 3.765, 7.018, 10.262].forEach(x => photo(s, x, 0.592, 2.554, 2.554));
  text(s, 'OUR GREAT TEAM', { x: 0.306, y: 4.136, w: 5.48, h: 0.774, fontFace: DISPLAY, fontSize: 40 });
  text(s, LOREM_LONG, { x: 0.328, y: 5.007, w: 10.054, h: 0.696, lineSpacingMultiple: 1.3 });
  iconGlobe(s, 10.946, 4.136, 1.873, { line: { width: 1.5 } });
}

/* 16 — Report table drawn from ruled lines (the original is not a real table). */
function slide16(s) {
  banner(s, 'PROJECT REPORT ANALYST ');
  outline(s, 'rect', 0.375, 1.652, 12.595, 5.476);
  [3.534, 6.667, 9.806].forEach(x => rule(s, x, 1.652, 0, 5.476));
  [2.435, 4.014, 5.58].forEach(y => rule(s, 0.375, y, 12.595, 0));

  const heads = [['REPORT', 4.329, 1.706], ['ANALYST', 7.345, 1.939], ['PERCENTS', 10.302, 2.172]];
  heads.forEach(([label, x, w]) =>
    text(s, label, { x, y: 1.823, w, h: 0.505, fontFace: DISPLAY, fontSize: 24, align: 'center' }));
  iconSparkle(s, 1.653, 1.744, 0.603);

  const rows = [['2020', '25% OFF', 2.817, 2.793, 2.878],
    ['2021', '75% OFF', 4.444, 4.384, 4.443],
    ['2022', '95% OFF', 6.033, 5.969, 5.993]];
  rows.forEach(([year, pct, yYear, yText, yPct]) => {
    text(s, year, { x: 1.197, y: yYear, w: 1.541, h: 0.707, fontFace: DISPLAY, fontSize: 36 });
    text(s, FILL_TEXT, { x: 3.832, y: yText, w: 2.537, h: 0.808 });
    text(s, FILL_TEXT, { x: 6.971, y: yText === 2.793 ? 2.816 : yText, w: 2.537, h: 0.808 });
    text(s, pct, { x: 10.124, y: yPct, w: 2.537, h: 0.707, fontFace: BODY_BLACK, fontSize: 36, align: 'center' });
  });
}

/* 17 — About the project (bulleted list + flower/sparkle motif). */
function slide17(s) {
  text(s, 'ABOUT THE PROJECT', { x: 0.464, y: 1.549, w: 5.342, h: 0.64, fontFace: DISPLAY, fontSize: 32 });
  const intro = 'PLACEHOLDER' +
    'will do some checks in the product whether it will deserve attention, which must be considered :';
  const bullets = ['Very high quality fashion product', 'Curation of the best',
    'The best content in this text', 'And many more'];
  const outro = 'And most importantly, you have to get your audience to get connection in with your ' +
    'conversation. good luck with your project';
  const body = [{ text: intro, options: { breakLine: true } }, { text: '', options: { breakLine: true } }];
  bullets.forEach(b => body.push({ text: b, options: { bullet: { characterCode: '2022', indent: 22.5 }, breakLine: true } }));
  body.push({ text: '', options: { breakLine: true } }, { text: outro });
  text(s, body, { x: 0.464, y: 2.498, w: 6.449, h: 3.453, lineSpacingMultiple: 1.3 });
  iconFlower(s, 8.258, 1.679, 4.134, { line: { width: 1.5 } });
  iconSparkle(s, 9.034, 2.454, 2.586, { line: { width: 1.5 } });
}

/* 18 — Deep talk with: scattered ornaments around a big globe. */
function slide18(s) {
  text(s, 'CASELDO CREATIVE PROJECT', { x: 0.353, y: 0.481, w: 6.715, h: 0.64, fontFace: DISPLAY, fontSize: 32 });
  iconGlobe(s, 6.938, 1.124, 5.241, { line: { width: 1.5 } });
  /* The scattered marks are printed as hairlines so they read as watermarks. */
  const faint = { line: { color: BRAND, width: 0.25 } };
  iconFlower(s, 0.452, 1.453, 1.231, faint);
  iconSparkle(s, 4.705, 2.318, 0.944, spin(faint, -15));
  iconStarBurst(s, 0.829, 5.5, 1.5, spin(faint, 15));
  iconFlower(s, 4.431, 4.81, 0.746, faint);
  text(s, 'DEEP TALK WITH',
    { x: 1.27, y: 2.965, w: 3.907, h: 0.572, fontFace: SUBHEAD, fontSize: 28, align: 'center' });
  text(s,
    'Believe me, the journey has not been a simple journey of progress. There have been many ups and downs, ' +
    'and it is the choices that I made at each of those times that have helped shape what I have achieved.',
    { x: 1.04, y: 3.528, w: 4.368, h: 1.007, fontSize: 10.5, align: 'center', lineSpacingMultiple: 1.3 });
}

/* 19 — Sales information: donut split into four quadrant arcs with elbow leaders. */
function slide19(s) {
  text(s, 'SALES INFORMATION',
    { x: 4.218, y: 0.627, w: 5.016, h: 0.64, fontFace: DISPLAY, fontSize: 32, align: 'center' });

  const ln = { color: BRAND, width: 1.5 };
  /* Three thin quadrants of a ring plus one solid pie wedge. */
  const RING_BL = [['M', 1.0, 1.0], ['L', 1.0, 0.618],
    ['C', 0.658, 0.618, 0.381, 0.341, 0.381, 0.0], ['L', 0.0, 0.0],
    ['C', 0.0, 0.552, 0.448, 1.0, 1.0, 1.0], ['Z']];
  const RING_TL = [['M', 0.0, 1.0], ['L', 0.381, 1.0],
    ['C', 0.381, 0.658, 0.658, 0.381, 1.0, 0.381], ['L', 1.0, 0.0],
    ['C', 0.447, 0.0, 0.0, 0.448, 0.0, 1.0], ['Z']];
  const RING_BR = [['M', 0.0, 0.618], ['L', 0.0, 1.0],
    ['C', 0.552, 1.0, 1.0, 0.552, 1.0, 0.0], ['L', 0.618, 0.0],
    ['C', 0.618, 0.341, 0.341, 0.618, 0.0, 0.618], ['Z']];
  const WEDGE_TR = [['M', 1.0, 1.0],
    ['C', 1.0, 0.448, 0.552, 0.0, 0.0, 0.0], ['L', 0.0, 1.0], ['Z']];
  drawPath(s, RING_BL, 5.044, 3.951, 1.623, 1.623, { line: ln });
  drawPath(s, RING_TL, 5.044, 2.328, 1.623, 1.623, { line: ln });
  drawPath(s, RING_BR, 6.667, 3.951, 1.623, 1.623, { line: ln });
  drawPath(s, WEDGE_TR, 6.667, 2.213, 1.738, 1.738, { line: ln });

  /* Elbow leader lines: out from the ring, up/down, then on to the caption. */
  const elbow = (x, y, w, h, right, down) =>
    drawPath(s, [['M', right ? 0 : 1, down ? 0 : 1], ['L', 0.5, down ? 0 : 1],
      ['L', 0.5, down ? 1 : 0], ['L', right ? 1 : 0, down ? 1 : 0]],
    x, y, w, h, { line: ln });
  elbow(8.129, 2.379, 1.511, 0.623, true, false);
  elbow(7.594, 5.285, 1.845, 0.467, true, true);
  elbow(3.979, 2.379, 1.356, 0.648, false, false);
  elbow(3.979, 5.102, 1.543, 0.767, false, true);

  const stats = [['45% OFF', 1.272, 1.658, 0.982, 2.379],
    ['45% OFF', 1.272, 5.149, 0.982, 5.87],
    ['95% OFF', 9.728, 1.658, 9.438, 2.379],
    ['45% OFF', 9.728, 5.031, 9.438, 5.751]];
  stats.forEach(([pct, px, py, tx, ty]) => {
    text(s, pct, { x: px, y: py, w: 2.537, h: 0.707, fontFace: DISPLAY, fontSize: 36, align: 'center' });
    text(s, LOREM_SHORT, { x: tx, y: ty, w: 3.117, h: 0.707, fontSize: 12, align: 'center' });
  });
}

/* 20 — Big number with a price-tag illustration. */
function slide20(s) {
  iconPriceTag(s, 5.106, 0.79);
  text(s, '2,000,000',
    { x: 3.304, y: 4.216, w: 6.725, h: 1.582, fontFace: DISPLAY, fontSize: 88, align: 'center' });
  text(s, "in one month we made to sell the product as a result of the team's hard work",
    { x: 3.712, y: 5.702, w: 5.909, h: 0.286, fontSize: 11, align: 'center' });
}

/* 21 — Three-step process: circled icons joined by dashed rules. */
function slide21(s) {
  banner(s, 'WHAT ARE WE PROCESS ON');
  const steps = [
    { cx: 1.534, label: 'WEBSITE', lx: 1.278, lw: 2.137, tx: 0.401, tw: 3.89 },
    { cx: 5.854, label: 'PAYMENT', lx: 5.523, lw: 2.288, tx: 4.654, tw: 4.026 },
    { cx: 10.175, label: 'SHIPPING', lx: 9.864, lw: 2.246, tx: 9.042, tw: 3.89 },
  ];
  steps.forEach(st => {
    outline(s, 'ellipse', st.cx, 1.922, 1.625, 1.625);
    text(s, st.label,
      { x: st.lx, y: 3.798, w: st.lw, h: 0.572, fontFace: SUBHEAD, fontSize: 28, align: 'center' });
    text(s, LOREM_MID, { x: st.tx, y: 4.497, w: st.tw, h: 0.909, fontSize: 12, align: 'center' });
  });
  iconGlobe(s, 1.789, 2.176, 1.113);
  iconMoney(s, 6.083, 2.449, 1.165, 0.569);
  iconBox(s, 10.278, 2.277);
  rule(s, 3.159, 2.734, 2.696, 0, { line: { dashType: 'dash' } });
  rule(s, 7.479, 2.734, 2.696, 0, { line: { dashType: 'dash' } });
}

/* 22 — Work in progress: two illustrated cards. */
function slide22(s) {
  banner(s, 'WORK IN PROGRESS');
  iconSparkle(s, 3.836, 0.443, 0.651);
  iconSparkle(s, 8.846, 0.443, 0.651);
  const cards = [
    { x: 2.409, title: 'DISCUSSION TEAM' },
    { x: 7.315, title: 'PRESENTATIONS' },
  ];
  cards.forEach(c => {
    text(s, c.title, { x: c.x, y: 4.764, w: 3.61, h: 0.505, fontFace: SUBHEAD, fontSize: 24, align: 'center' });
    text(s, LOREM_SHORT, { x: c.x, y: 5.298, w: 3.61, h: 0.707, fontSize: 12, align: 'center' });
  });
  iconChatPerson(s, 3.31, 2.873, false);
  iconChatPerson(s, 3.31, 3.777, true);
  iconProjector(s, 8.226, 2.873);
}

/* 23 — Customer feedback: four quotes flanking circular portraits. */
function slide23(s) {
  banner(s, 'CUSTOMER FEEDBACK');
  iconSparkle(s, 3.559, 0.443, 0.651);
  iconSparkle(s, 9.149, 0.443, 0.651);
  const quote = '“Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa”';
  const cards = [
    { name: 'Jayden Sacramento', nx: 1.838, tx: 1.601, y: 2.24, px: 4.842, py: 2.075 },
    { name: 'Aarav Adelio', nx: 7.088, tx: 6.851, y: 3.016, px: 10.092, py: 2.85 },
    { name: 'Berg Zaynskie', nx: 1.913, tx: 1.676, y: 4.457, px: 4.842, py: 4.291 },
    { name: 'Zayd Pradipta', nx: 7.227, tx: 6.99, y: 5.264, px: 10.092, py: 5.112 },
  ];
  cards.forEach(c => {
    text(s, c.name,
      { x: c.nx, y: c.y, w: 2.852, h: 0.438, fontFace: SUBHEAD, fontSize: 20, align: 'right' });
    text(s, quote, { x: c.tx, y: c.y + 0.464, w: 3.089, h: 0.707, fontSize: 12, align: 'right' });
    photo(s, c.px, c.py, 1.501, 1.501, { shape: 'ellipse' });
  });
}

/* 24 — Step tutorial: globe over a dashed timeline with three markers. */
function slide24(s) {
  banner(s, 'STEP TUTORIAL');
  iconSparkle(s, 4.098, 0.443, 0.651);
  iconSparkle(s, 8.472, 0.443, 0.651);
  iconGlobe(s, 5.188, 1.573, 2.958, { line: { width: 1.5 } });
  rule(s, 1.107, 5.03, 11.594, 0, { line: { width: 1.5, dashType: 'dash' } });
  const steps = [
    { dot: 0.766, x: 0.753, title: 'Create Product', tw: 1.948 },
    { dot: 4.994, x: 4.994, title: 'Publishing Product', tw: 2.239 },
    { dot: 9.22, x: 9.235, title: 'Sell A Product', tw: 1.948 },
  ];
  steps.forEach(st => {
    block(s, st.dot, 4.86, 0.34, 0.34, BRAND, { shape: 'ellipse' });
    text(s, st.title, { x: st.x, y: 5.263, w: st.tw, h: 0.37, fontFace: SUBHEAD, fontSize: 16 });
    text(s, LOREM_MID, { x: st.x, y: 5.633, w: 3.345, h: 0.774, fontSize: 10 });
  });
}

/* 25 — A History Of Brand. */
function slide25(s) {
  iconGlobe(s, 10.616, 1.347, 2.283, { line: { width: 1.5 } });
  text(s, paras(['A History', 'Of Brand']),
    { x: 0.39, y: 1.226, w: 4.53, h: 2.524, fontFace: SUBHEAD, fontSize: 72 });
  text(s,
    "A brand is a name, term, design, symbol or any other feature that identifies one seller's good or service " +
    'as distinct from those of other sellers. Brands are used in business, marketing, and advertising for ' +
    "recognition and, importantly, to create and store value as brand equity for the object identified, to the " +
    "benefit of the brand's customers, its owners and shareholders. Name brands are sometimes distinguished " +
    'from generic or store brands. ',
    { x: 0.39, y: 4.36, w: 12.553, h: 1.308, lineSpacingMultiple: 1.3 });
  rule(s, 5.12, 2.488, 4.967, 0, { line: { width: 1.5 } });
}

/* 26 — Data process: four petal circles with leader lines to captions. */
function slide26(s) {
  banner(s, 'DATA PROCESS');
  iconSparkle(s, 4.098, 0.443, 0.651);
  iconSparkle(s, 8.472, 0.443, 0.651);

  const ln = { color: BRAND, width: 1.5 };
  /* Four circles pinwheeled around a small hub, each bitten where the hub sits. */
  const PETAL_TL = [['M', .815, .916], ['C', .728, .83, .619, .779, .499, .779],
    ['C', .344, .779, .206, .865, .114, 1], ['C', .042, .894, 0, .758, 0, .61],
    ['C', 0, .273, .223, 0, .499, 0], ['C', .775, 0, .999, .273, .999, .61],
    ['C', .999, .701, .983, .788, .953, .866], ['C', .949, .866, .944, .866, .94, .866],
    ['C', .893, .866, .85, .885, .815, .916], ['Z']];
  const PETAL_TR = [['M', 0, .5], ['C', .031, .436, .049, .365, .049, .29],
    ['C', .049, .255, .045, .222, .038, .19], ['C', .137, .074, .289, 0, .46, 0],
    ['C', .758, 0, .999, .224, .999, .5], ['C', .999, .776, .758, 1, .46, 1],
    ['C', .332, 1, .213, .958, .121, .887], ['C', .179, .849, .216, .785, .216, .714],
    ['C', .216, .6, .12, .506, 0, .5], ['Z']];
  const PETAL_BR = [['M', .184, .083], ['C', .27, .169, .379, .22, .499, .22],
    ['C', .654, .22, .792, .134, .884, 0], ['C', .956, .105, .999, .241, .999, .389],
    ['C', .999, .726, .775, .999, .499, .999], ['C', .223, .999, 0, .726, 0, .389],
    ['C', 0, .297, .016, .21, .045, .132], ['C', .049, .133, .054, .133, .058, .133],
    ['C', .105, .133, .148, .114, .184, .083], ['Z']];
  const PETAL_BL = [['M', .999, .499], ['C', .967, .563, .949, .634, .949, .709],
    ['C', .949, .744, .953, .777, .96, .809], ['C', .861, .925, .709, 1, .537, 1],
    ['C', .24, 1, 0, .775, 0, .5], ['C', 0, .224, .24, 0, .537, 0],
    ['C', .666, 0, .784, .042, .877, .112], ['C', .819, .151, .782, .214, .782, .285],
    ['C', .782, .399, .878, .493, .999, .499], ['Z']];
  drawPath(s, PETAL_TL, 5.233, 2.878, 1.523, 1.248, { line: ln });
  drawPath(s, PETAL_TR, 6.686, 3.198, 1.414, 1.523, { line: ln });
  drawPath(s, PETAL_BR, 6.577, 4.446, 1.523, 1.248, { line: ln });
  drawPath(s, PETAL_BL, 5.233, 3.851, 1.414, 1.523, { line: ln });
  outline(s, 'ellipse', 6.507, 4.126, 0.32, 0.32, { line: ln });

  [[4.531, 2.878], [4.533, 5.374], [7.339, 3.198], [7.339, 5.694]]
    .forEach(([x, y]) => rule(s, x, y, 1.462, 0, { line: ln }));

  const items = [
    { title: 'Business Work', tx: 2.262, tw: 1.948, bx: 1.33, y: 2.332, align: 'right' },
    { title: 'Team Progress', tx: 2.262, tw: 1.948, bx: 1.33, y: 4.842, align: 'right' },
    { title: 'Work Life Balance', tx: 8.862, tw: 2.21, bx: 8.862, y: 2.693, align: 'left' },
    { title: 'Work As Team', tx: 8.862, tw: 1.948, bx: 8.862, y: 5.169, align: 'left' },
  ];
  items.forEach(it => {
    text(s, it.title,
      { x: it.tx, y: it.y, w: it.tw, h: 0.37, fontFace: SUBHEAD, fontSize: 16, align: it.align });
    text(s, LOREM_MID,
      { x: it.bx, y: it.y + 0.37, w: 2.91, h: 0.963, fontSize: 10, align: it.align, lineSpacingMultiple: 1.3 });
  });
}

/* 27 — Gallery of catalog: mosaic of image frames. */
function slide27(s) {
  text(s, 'GALLERY OF CATALOG', { x: 0.299, y: 0.287, w: 5.661, h: 0.64, fontFace: DISPLAY, fontSize: 32 });
  text(s,
    "A brand is a name, term, design, symbol or any other feature that identifies one seller's good or service " +
    'as distinct from those of other sellers. Brands are used in business, marketing, and advertising for ' +
    'recognition and, importantly,',
    { x: 6.323, y: 0.825, w: 6.605, h: 0.873, fontSize: 12, lineSpacingMultiple: 1.3 });
  [[0.405, 0.927, 2.818, 6.286], [3.329, 0.927, 2.876, 2.823], [3.329, 3.864, 5.699, 3.349],
    [9.135, 3.864, 3.793, 3.349], [6.323, 1.812, 6.605, 1.938]]
    .forEach(([x, y, w, h]) => photo(s, x, y, w, h));
}

/* 28 — Customer satisfaction. */
function slide28(s) {
  banner(s, 'CUSTOMER SATISFACTION');
  iconSparkle(s, 3.239, 0.443, 0.651);
  iconSparkle(s, 9.443, 0.443, 0.651);
  text(s, paras(['Rate', 'Feedback']),
    { x: 0.39, y: 1.765, w: 5.066, h: 2.524, fontFace: SUBHEAD, fontSize: 72 });
  text(s,
    'Customer satisfaction (often abbreviated as CSAT) is a term frequently used in marketing. It is a measure ' +
    'of how products and services supplied by a company meet or surpass customer expectation. Customer ' +
    'satisfaction is defined as "the number of customers, or percentage of total customers, whose reported ' +
    'experience with a firm, its products, or its services (ratings) exceeds specified satisfaction goals. ' +
    'Customers play an important role and are essential in keeping a product or service relevant; it is, ' +
    'therefore, in the best interest of the business to ensure customer satisfaction and build customer loyalty. ',
    { x: 0.39, y: 5.019, w: 12.553, h: 1.615, lineSpacingMultiple: 1.3 });
  iconRating(s, 10.68, 1.814);
  rule(s, 5.718, 2.971, 4.369, 0, { line: { width: 1.5 } });
}

/* 29 — Our catalog. */
function slide29(s) {
  text(s, 'OUR CATALOG', { x: 0.282, y: 1.342, w: 3.875, h: 0.64, fontFace: DISPLAY, fontSize: 32 });
  text(s, 'Shopping Market Product',
    { x: 0.282, y: 1.894, w: 3.466, h: 0.472, fontSize: 18, lineSpacingMultiple: 1.3 });
  text(s, paras([LOREM_LONG, '', LOREM_NUNC, '', LOREM_PELL]),
    { x: 0.362, y: 2.744, w: 5.3, h: 2.711, fontSize: 12, lineSpacingMultiple: 1.3 });
  iconStarBurst(s, 4.272, 1.478, 0.754);
  photo(s, 6.667, 1.41, 2.786, 3.34);
  photo(s, 10.184, 1.41, 2.786, 3.34);
  const cards = [
    { x: 6.543, title: 'Ready-Made Fabric', tw: 2.91 },
    { x: 10.061, title: 'Made By Hand', tw: 1.948 },
  ];
  cards.forEach(c => {
    text(s, c.title, { x: c.x, y: 5.013, w: c.tw, h: 0.37, fontFace: SUBHEAD, fontSize: 16 });
    text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere,',
      { x: c.x, y: 5.383, w: 2.91, h: 0.745, fontSize: 10, lineSpacingMultiple: 1.3 });
  });
}

/* 30 — Contact us: paper header band over a white footer panel. */
function slide30(s) {
  block(s, 0.0, 3.495, 13.333, 4.005, WHITE);
  iconSparkle(s, 9.012, 0.402, 1.132);
  text(s, 'CASELDO',
    { x: 3.817, y: 0.95, w: 5.7, h: 1.447, fontFace: DISPLAY, fontSize: 80, align: 'center' });
  text(s, 'CASELDO PRESENTATION PROJECT PROPOSAL',
    { x: 4.026, y: 2.174, w: 5.281, h: 0.39, align: 'center', lineSpacingMultiple: 1.3 });
  text(s, 'CONTACT US',
    { x: 4.984, y: 3.755, w: 3.366, h: 0.654, fontFace: DISPLAY, fontSize: 28, align: 'center', lineSpacingMultiple: 1.3 });
  text(s, LOREM_LONG,
    { x: 3.498, y: 4.461, w: 6.337, h: 0.758, fontSize: 11, align: 'center', lineSpacingMultiple: 1.2 });

  iconPin(s, 3.938, 5.651, 0.42, 0.393);
  iconClock(s, 7.235, 5.647, 0.4);
  iconPhone(s, 3.953, 6.595, 0.43, 0.332);
  iconEnvelope(s, 7.242, 6.643, 0.408, 0.274);

  const lines = [
    { x: 4.373, y: 5.609, h: 0.536, body: ['3456 Sacramento States Street', 'CA 78211 United States Of America'] },
    { x: 7.65, y: 5.591, h: 0.536, body: ['Monday – Friday', '07.30 – 16.00'] },
    { x: 4.373, y: 6.604, h: 0.314, body: ['+1 234 567 890'] },
    { x: 7.65, y: 6.604, h: 0.314, body: ['youremail@example.com'] },
  ];
  lines.forEach(l =>
    text(s, paras(l.body), { x: l.x, y: l.y, w: 2.724, h: l.h, fontSize: 11, lineSpacingMultiple: 1.2 }));
}

/* 31 — Closing card. */
function slide31(s) {
  block(s, 0.379, 0.597, 12.576, 6.306, WHITE);
  photo(s, 4.427, 1.015, 4.479, 5.471, { shape: 'round2SameRect' });
  text(s, 'CASELDO',
    { x: 4.427, y: 4.88, w: 4.479, h: 1.111, fontFace: DISPLAY, fontSize: 60, color: WHITE, align: 'center' });
  text(s, 'CASELDO PRESENTATION PROJECT PROPOSAL',
    { x: 4.427, y: 5.937, w: 4.479, h: 0.328, fontSize: 11, color: WHITE, align: 'center', lineSpacingMultiple: 1.3 });
  iconGlobe(s, 5.411, 1.692, 2.511, { line: { color: WHITE, width: 1.5 } });
  text(s, 'MADE BY PROJECT PROPOSAL',
    { x: 0.379, y: 6.616, w: 2.466, h: 0.287, fontSize: 9, lineSpacingMultiple: 1.3 });
  text(s, 'BIG THANK FOR THE PRESENTATION',
    { x: 10.039, y: 6.616, w: 2.916, h: 0.287, fontSize: 9, align: 'right', lineSpacingMultiple: 1.3 });
}

/* ------------------------------------------------------------------ *
 * Assemble
 * ------------------------------------------------------------------ */
const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
  slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
  slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24,
  slide25, slide26, slide27, slide28, slide29, slide30, slide31,
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'CASELDO', width: SLIDE_W, height: SLIDE_H });
  pptx.layout = 'CASELDO';
  pptx.author = 'CASELDO';
  pptx.title = 'CASELDO Presentation Project Proposal';

  pptx.defineSlideMaster({ title: 'PAPER', background: { color: PAPER } });

  BUILDERS.forEach(fn => fn(pptx.addSlide({ masterName: 'PAPER' })));

  return pptx.writeFile({ fileName: path.join(__dirname, '14001439-4310-421c-9b69-6d5e8ea1e1a1_grok_final.pptx') });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
