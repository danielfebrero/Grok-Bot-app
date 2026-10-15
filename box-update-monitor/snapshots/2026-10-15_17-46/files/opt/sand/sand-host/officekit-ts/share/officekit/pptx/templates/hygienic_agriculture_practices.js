/**
 * "Hygienic Agriculture Product" - 25-slide deck rebuilt with pptxgenjs.
 *
 *   node 154dd48c-4153-444f-b91c-e1861f280ed2_grok_final.js
 *   -> 154dd48c-4153-444f-b91c-e1861f280ed2_grok_final.pptx (written next to this file)
 *
 * Everything is native pptxgenjs geometry. Photographs and raster icons from the
 * source deck are stood in for by `photo()` blocks and `iconChip()` outlines.
 * All coordinates are inches on a 13.333 x 7.5in stage.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ================================================================= palette */

// Theme accents plus the tints/shades the deck derives from them.
const C = {
  WHITE:      'FFFFFF',
  BLACK:      '000000',
  INDIGO:     '2C097A',   // accent3 shaded - the dominant heading colour
  VIOLET:     '3A0CA3',   // accent3
  PLUM:       '1D0652',
  ORCHID:     '7944F1',
  LILAC:      'A582F6',
  LILAC_LT:   'D2C1FA',
  BLUE:       '4361EE',   // accent4
  BLUE_DK:    '1334D2',
  NAVY:       '0D238C',
  PERI:       '8EA0F5',
  LEAF:       '2D8B17',   // accent2
  LEAF_DK:    '226811',
  FOREST:     '17460C',
  SPRING:     '68E14C',
  GREEN:      '134D36',   // accent1
  GREEN_DK:   '0E3A29',
  MINT:       'DEF7ED',
  MAGENTA:    '830B52',   // accent5
  MAGENTA_DK: '420529',
  GOLD:       'D0B200',   // accent6
  CHARCOAL:   '222A35',
  GREY_05:    'F9F9F9',
  GREY_10:    'F2F2F2',
  GREY_12:    'E7E6E6',
  GREY_15:    'D6DCE5',
  GREY_20:    'D9D9D9',
  GREY_25:    'BFBFBF',
  GREY_45:    'A6A6A6',
  GREY_50:    '808080',
};

const HEAD_FONT = 'Jost';         // theme major font
const BODY_FONT = 'Lato';         // theme minor font
const KICKER_FONT = 'Google Sans';

/* ============================================================== boilerplate */

// The lorem body copy is assembled from these sentences all over the deck.
const S1 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. ';
const S2 = 'In it ipsum sed ante commodo, id mattis odio posuere. ';
const S3 = 'Nunc eu mauris non arcu blandit cursus. ';
const S4 = 'Quisque convallis augue vit augue posuere suscipit. ';
const S5 = 'Pellentesque at sem magna';
const S6 = 'Vivamus a purus consectetur, molestie purus.';

const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, ';
const LOREM_MED   = S1 + S2 + S3;
const LOREM_LONG  = S1 + S2 + S3 + S4 + S5;
const LOREM_XL    = LOREM_LONG + '. ' + S1 + S6;

/* ================================================================== helpers */

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'WIDE';
pptx.theme = { headFontFace: HEAD_FONT, bodyFontFace: BODY_FONT };
pptx.title = 'Hygienic Agriculture Product';

/** Slide furniture inherited from the slide master: page badge + kicker. */
function chrome(s, num) {
  s.addShape('ellipse', { x: 12.514, y: 6.783, w: 0.531, h: 0.531, fill: { color: C.GREEN_DK } });
  s.addText(String(num), { x: 12.391, y: 6.863, w: 0.776, h: 0.37, align: 'center',
    fontSize: 16, bold: true, color: C.WHITE });
  s.addText('Hygienic Agriculture', { x: 9.951, y: 0.08, w: 3.217, h: 0.37, align: 'right',
    fontSize: 16, color: C.LEAF, fontFace: KICKER_FONT });
}

/** Stand-in block for a photograph in the source deck. */
function photo(s, x, y, w, h) {
  s.addShape('rect', { x, y, w, h, fill: { color: C.GREY_15 } });
  s.addText('[image]', { x, y, w, h, align: 'center', valign: 'middle',
    fontSize: 14, color: C.GREY_50 });
}

/** Stand-in outline for one of the small white line-art raster icons. */
function iconChip(s, x, y, d, color) {
  s.addShape('roundRect', { x, y, w: d, h: d, rectRadius: d * 0.22,
    fill: { type: 'none' }, line: { color: color || C.WHITE, width: 1 } });
}

/**
 * Draw a traced outline from ART, rescaled into the box (x, y, w, h).
 * ART entries are one contour (flat [x0,y0, x1,y1, ...] in unit coords) or a
 * list of contours for shapes with holes / separate islands.
 */
function art(s, name, x, y, w, h, opts) {
  const raw = ART[name];
  const contours = Array.isArray(raw[0]) ? raw : [raw];
  const points = [];
  contours.forEach(flat => {
    for (let i = 0; i < flat.length; i += 2) {
      points.push({ x: flat[i] * w, y: flat[i + 1] * h, moveTo: i === 0 });
    }
    points.push({ close: true });
  });
  s.addShape('custGeom', Object.assign({ x, y, w, h, points }, opts));
}

/** Heading in the deck's display face. */
function title(s, x, y, w, h, text, fontSize, color) {
  s.addText(text, { x, y, w, h, fontSize, bold: true, color: color || C.INDIGO,
    fontFace: HEAD_FONT, lineSpacingMultiple: 0.9 });
}

/** Running body copy (14pt) . */
function body(s, x, y, w, h, text, opts) {
  s.addText(text, Object.assign({ x, y, w, h, fontSize: 14, color: C.BLACK,
    lineSpacingMultiple: 1.3 }, opts));
}

/** Small caption copy (12pt). */
function note(s, x, y, w, h, text, opts) {
  s.addText(text, Object.assign({ x, y, w, h, fontSize: 12, color: C.BLACK,
    lineSpacingMultiple: 1.3 }, opts));
}

/**
 * The deck's standard list item: a bold 16pt label above a caption.
 * o = { x, y, w, label, color, text, align?, gap?, capW?, capH?, capSize? }
 */
function labelled(s, o) {
  s.addText(o.label, { x: o.x, y: o.y, w: o.w, h: 0.37, align: o.align,
    fontSize: 16, bold: true, color: o.color, lineSpacingMultiple: 1.3 });
  s.addText(o.text, { x: o.x, y: o.y + (o.gap || 0.336), w: o.capW || o.w, h: o.capH || 0.6,
    align: o.align, fontSize: o.capSize || 12, color: C.BLACK, lineSpacingMultiple: 1.3 });
}

/** Big number over a small caption (the deck's KPI treatment). */
function stat(s, x, y, value, color, caption) {
  s.addText(value, { x, y, w: 1.385, h: 0.572, fontSize: 28, bold: true, color, fontFace: HEAD_FONT });
  s.addText(caption || 'Lorem ipsum', { x, y: y + 0.532, w: 1.385, h: 0.348,
    fontSize: 12, color, fontFace: HEAD_FONT, lineSpacingMultiple: 1.3 });
}

/**
 * Filled circle carrying a bold label and a small blurb.
 * `dy` is the label's offset below the top of the circle; the blurb follows 0.226in later.
 */
function circleBadge(s, x, y, d, fill, label, text, dy) {
  const cx = x + d / 2;
  s.addShape('ellipse', { x, y, w: d, h: d, fill: { color: fill } });
  s.addText(label, { x: cx - 0.984, y: y + dy, w: 1.968, h: 0.303,
    align: 'center', fontSize: 12, bold: true, color: C.WHITE });
  s.addText(text, { x: cx - 0.882, y: y + dy + 0.226, w: 1.764, h: 0.764,
    align: 'center', fontSize: 10.5, color: C.WHITE, lineSpacingMultiple: 1.3 });
}

/** Stadium-shaped button/banner with a centred bold caption. */
function pill(s, x, y, w, h, fill, text, opts) {
  s.addShape('roundRect', { x, y, w, h, rectRadius: Math.min(w, h) / 2, fill: { color: fill } });
  if (text) {
    s.addText(text, Object.assign({ x, y, w, h, align: 'center', valign: 'middle',
      fontSize: 16, bold: true, color: C.WHITE }, opts));
  }
}

/* ============================================================ traced shapes */

// Outlines lifted from the source deck, normalised into a 0..1 unit box.
const ART = {
  leafHand: [
    [0.9810,0.7202, 0.4882,0.5550, 0.7583,0.1009, 0.1564,0.0092, 0.4502,0.5596, 0.0881,0.8278,
    0.9810,0.7202],
    [0.7251,0.1376, 0.5166,0.3945, 0.7251,0.1376],
    [0.2417,0.2248, 0.4976,0.2294, 0.2417,0.2248],
    [0.1706,0.7248, 0.6730,0.6468, 0.1706,0.7248],
    [0.8626,0.8945, 0.0379,0.7156, 0.8626,0.8945]],
  roundRightPanel: [0,0, 0.7978,0, 0.8617,0.0126, 0.9172,0.0476, 0.9774,0.1333, 1,0.2467, 1,0.7533,
    0.9897,0.8313, 0.9610,0.8990, 0.8907,0.9725, 0.8306,0.9968, 0,1, 0,0],
  molecule: [
    []],
  pieSliceA: [
    [0.9242,0.9744, 0.9470,0.0642, 0.5058,0.0985, 0.0639,0.3335, 0.9242,0.9744],
    [0.8600,0.8550, 0.1767,0.3256, 0.4946,0.1770, 0.8600,0.1175, 0.8600,0.8550]],
  pieSliceB: [
    [0.9775,0.5620, 0.3444,0.0375, 0.1335,0.2517, 0.0504,0.4953, 0.0888,0.7775, 0.2406,0.9860,
    0.9775,0.5620],
    [0.2500,0.9175, 0.1212,0.4961, 0.3438,0.1055, 0.8888,0.5555, 0.2500,0.9175]],
  pieRing: [
    [0.9741,0.4477, 0.7976,0.1430, 0.4425,0.0289, 0.4276,0.4973, 0.0286,0.7866, 0.3038,0.9763,
    0.6522,0.9627, 0.9151,0.7430, 0.9741,0.4477],
    [0.9380,0.5360, 0.8357,0.7889, 0.5869,0.9440, 0.3033,0.9367, 0.0716,0.7787, 0.4673,0.5125,
    0.4683,0.0627, 0.6524,0.1008, 0.8148,0.2110, 0.9380,0.5360]],
  calDigit1: [
    [0.1729,0.3780, 0.9936,0.9671, 1.0166,0.0653, 0.1729,0.3780],
    [0.2386,0.2613, 0.8943,0.1113, 0.8943,0.9093, 0.2386,0.2613]],
  calDigit5: [0.2029,0.9067, 0.8165,0.0541, 0.1324,0.0641, 0.2029,0.9067],
  calBody: [
    [0.8792,0.1690, 0.8035,0.0232, 0.7178,0.1695, 0.1813,0.0354, 0.0203,0.2675, 0.0638,0.9730,
    0.9652,0.9518, 0.8792,0.1690],
    [0.7524,0.0755, 0.7795,0.2722, 0.7524,0.0755],
    [0.1978,0.0755, 0.2249,0.2722, 0.1978,0.0755],
    [0.9481,0.8935, 0.0776,0.9384, 0.1262,0.2035, 0.2241,0.3170, 0.9281,0.2222, 0.9484,0.8935]],
  calInner: [
    [0.9752,0.0357, 0.0352,0.0402, 0.0303,0.9231, 0.1013,0.9924, 0.9510,0.9834, 0.9981,0.8791,
    0.9981,0.0686, 0.9752,0.0357],
    [0.9529,0.8791, 0.9206,0.9267, 0.0847,0.9201, 0.0687,0.8791, 0.0687,0.1019, 0.9529,0.1019,
    0.9529,0.8791]],
  pyrFrontA: [0,0.4679, 0.4460,0, 1,0.4377, 0.5552,1, 0,0.4679],
  pyrRightA: [0.7792,0, 0,0.4352, 0.0202,1, 1,0.4536, 0.7792,0],
  pyrLeftA: [0.9833,0.4188, 1,1, 0,0.4741, 0.1889,0, 0.9833,0.4188],
  pyrFrontB: [0,0.4923, 0.4574,0, 1,0.4563, 0.5521,1, 0,0.4923],
  pyrRightB: [0.7100,0, 0,0.3470, 0.0274,1, 1,0.5217, 0.7100,0],
  pyrLeftB: [1,1, 0.9772,0.3317, 0.2486,0, 0,0.5424, 1,1],
  pyrFrontC: [0,0.4031, 0.4322,0, 1,0.3796, 0.5577,1, 0,0.4031],
  pyrRightC: [0.5690,0, 0,0.2229, 0.0406,1, 1,0.6260, 0.5690,0],
  pyrLeftC: [1,1, 0.9685,0.2108, 0.3762,0, 0,0.6422, 1,1],
  pyrTipRight: [0,0, 0.0911,1, 1,0.8016, 0,0],
  pyrTipLeft: [1,1, 0,0.8133, 0.9188,0, 1,1],
  swotPetal: [0.0723,0, 0.5260,0, 0.6250,0.0145, 0.7593,0.0666, 0.8180,0.1050, 0.9146,0.2028,
    0.9775,0.3232, 0.9942,0.3899, 1,0.5401, 0.9935,0.6147, 0.9745,0.6854, 0.9442,0.7514,
    0.9035,0.8117, 0.7953,0.9113, 0.7298,0.9487, 0.6580,0.9766, 0.5811,0.9940, 0.5000,1,
    0.4189,0.9940, 0.3420,0.9766, 0.2702,0.9487, 0.2047,0.9113, 0.0965,0.8117, 0.0558,0.7514,
    0.0255,0.6854, 0.0065,0.6147, 0,0.5401, 0.0019,0.0513, 0.0271,0.0146, 0.0723,0],
  infoBadge: [
    [0.4978,0.0195, 0,0.5184, 0.4978,0.9978, 0.9978,0.4989, 0.4978,0.0195],
    [0.4221,0.8069, 0.3268,0.5184, 0.5368,0.4013, 0.4221,0.8069]],
  funnelSpout: [1,0, 1,0.8809, 0.9954,0.9186, 0.9634,0.9770, 0.9106,1, 0.0894,1, 0.0611,0.9939,
    0.0172,0.9513, 0,0.8809, 0,0, 1,0],
  dot: [1,0.5000, 0.7953,0.9035, 0.5000,1, 0.2047,0.9035, 0,0.5000, 0.0965,0.2047, 0.5000,0,
    0.9035,0.2047, 1,0.5000],
  leader: [0,0, 1,0, 1,1, 0,1],
  funnelBL: [1,0, 1,1, 0.6180,1, 0,0],
  funnelBR: [1,0, 0.3821,1, 0,1, 0,0],
  funnelTR: [0.9960,0.0697, 0.6325,1, 0,1, 0,0, 0.9631,0, 0.9945,0.0229, 0.9960,0.0697],
  funnelTL: [1,0, 1,1, 0.3676,1, 0.0003,0.0535, 0.0134,0.0110, 0.0370,0, 1,0],
  hubRing: [0.5000,1, 0.6005,0.9907, 0.6706,0.9709, 0.7989,0.9008, 0.8704,0.8360, 0.9140,0.7804,
    0.9616,0.6944, 0.9894,0.6019, 1,0.4749, 0.9841,0.3757, 0.9511,0.2831, 0.8849,0.1812,
    0.8360,0.1296, 0.7593,0.0728, 0.6706,0.0304, 0.5516,0.0026, 0.4735,0, 0.3505,0.0225,
    0.2606,0.0608, 0.1812,0.1151, 0.1138,0.1812, 0.0595,0.2619, 0.0291,0.3294, 0.0053,0.4246,
    0,0.5265, 0.0225,0.6495, 0.0595,0.7394, 0.1138,0.8188, 0.1640,0.8704, 0.2606,0.9405,
    0.3280,0.9709, 0.3981,0.9907, 0.5000,1],
  noteBar: [0,0, 1,0, 0.9983,0.7983, 0.9789,1, 0.0211,1, 0.0062,0.9033, 0,0.6697],
  calIcon: [
    [0,0.8906, 0.0542,0.1227, 0.9850,0.1618, 0.9458,0.9850, 0,0.8906],
    [0.0847,0.8906, 0.9155,0.3918, 0.0847,0.8906]],
  caret: [0.0238,0.6353, 0.1212,0.5301, 0.4199,0.7724, 0.5010,0, 0.5821,0.7724, 0.8630,0.5410,
    0.9988,0.5713, 0.5007,1, 0.0238,0.6353],
  mapBit1: [0.9500,0.4545, 0.3000,0, 0,0.5455, 0.4968,0.8924, 0.8722,0.9966, 0.9500,0.9091,
    0.9500,0.4545],
  mapBit2: [0.4231,0.8929, 0.4231,0.7526, 0.9808,0.2785, 0.6803,0.0480, 0.1731,0.1429, 0,0.5357,
    0.2500,0.7143, 0.3077,1, 0.4231,0.8929],
  mapAlaska: [0.8327,0.9434, 0.8755,0.9371, 0.8599,0.8491, 0.9572,0.8302, 0.9919,0.6433, 0.8833,0.6667,
    0.8379,0.8132, 0.7322,0.8217, 0.6441,0.6583, 0.6498,0.4025, 0.5914,0.3648, 0.5242,0.1888,
    0.4525,0.2125, 0.3502,0.0566, 0.1984,0.0818, 0.0739,0, 0,0.0189, 0.0968,0.2318,
    0.0790,0.3025, 0.2384,0.5691, 0.2386,0.4931, 0.0864,0.1514, 0.0744,0.0515, 0.1362,0.0818,
    0.1510,0.2292, 0.2583,0.3541, 0.3804,0.6027, 0.3944,0.7584, 0.6541,0.9581, 0.7374,0.9406,
    0.8054,1, 0.8327,0.9434],
  mapBit3: [0.8611,0.6571, 0.7778,0, 0.3889,0, 0.5000,0.4000, 0,0.6857, 0.5833,1, 0.8611,0.6571],
  mapBit4: [0.1818,0.2609, 0,0.2609, 0,1, 0.3636,0.9565, 0.6364,0.7826, 0.6858,0.2414, 1,0.1304,
    0.6364,0, 0.1818,0.2609],
  mapBit5: [0.4400,0.0952, 0.6800,0.6667, 0,0.8095, 0.1450,0.9286, 0.9200,1, 1,0.2381, 0.4400,0.0952],
  mapBit6: [0.3793,0.7500, 0.9586,0.6524, 1,0.5500, 0.6681,0.1938, 0.0690,0.1000, 0,0.9000,
    0.0981,0.9328, 0.3793,0.7500],
  mapBit7: [0.5278,0.9459, 0.8889,1, 0.8424,0.8919, 1,0, 0.1667,0.2993, 0.1451,0.4384, 0,0.4865,
    0.4167,0.9459, 0.5278,0.9459],
  mapBit8: [0.9643,0.8182, 1,0.5455, 0.7143,0.0909, 0.1071,0, 0.0614,0.3920, 0.3214,0.4091,
    0.7474,0.9832, 0.9286,1, 0.9643,0.8182],
  mapBit9: [0.8980,0.3043, 0.0408,0.0435, 0,0.4783, 0.3832,0.9259, 0.7143,0.3043, 0.8571,0.9130,
    0.8980,0.3043],
  mapBit10: [0.8090,0.6452, 0.3596,0.0968, 0.1396,0.1755, 0.0225,0.4516, 0.2130,0.1956, 0.3071,0.3866,
    0.4840,0.4756, 0.6791,0.7954, 0.7152,0.9795, 0.9888,0.9355, 0.8090,0.6452],
  mapBit11: [0.0952,0.4000, 0.6267,0.7816, 0.8114,0.8032, 0.9048,0.7000, 0.8480,0.4768, 0.6335,0.3064,
    0.1520,0.2392, 0.0952,0.4000],
  mapBit12: [0.1579,0.4545, 0.3103,0.6815, 0.4931,0.7425, 0.6733,0.6902, 0.8947,0.4545, 0.8434,0.3062,
    0.3794,0.1796, 0.1840,0.2538, 0.1579,0.4545],
  mapGreenland: [
    [0.9438,0.0935, 0.8624,0.1091, 0.8023,0.1610, 0.8278,0.0851, 0.7511,0.1260, 0.7734,0.0917,
    0.6589,0.1039, 0.8395,0.0605, 0.7677,0.0153, 0.4322,0.0286, 0.4496,0.0494, 0.4205,0.0623,
    0.4647,0.1033, 0.3659,0.0582, 0.3798,0.0909, 0.3275,0.0701, 0.3144,0.1176, 0.2991,0.0665,
    0.1942,0.0955, 0.1919,0.1351, 0.0887,0.1784, 0.1279,0.1922, 0.1143,0.2312, 0,0.2779,
    0.1055,0.3134, 0.0248,0.3312, 0.0537,0.3668, 0.2031,0.3838, 0.2691,0.4512, 0.2857,0.5722,
    0.3217,0.5658, 0.3604,0.6159, 0.3043,0.6104, 0.3630,0.6524, 0.3142,0.7318, 0.3176,0.7855,
    0.4038,0.9633, 0.4787,1, 0.5359,0.8435, 0.5253,0.8188, 0.5718,0.7774, 0.6206,0.7812,
    0.6693,0.7062, 0.7668,0.6851, 0.8374,0.6298, 0.7635,0.6286, 0.7789,0.5865, 0.8411,0.6130,
    0.7984,0.5403, 0.8310,0.5440, 0.8371,0.5177, 0.7838,0.5133, 0.8027,0.4872, 0.8568,0.4981,
    0.8416,0.4614, 0.8835,0.4571, 0.8508,0.4234, 0.8824,0.3881, 0.8308,0.3534, 0.8983,0.3308,
    0.8635,0.3206, 0.8741,0.2978, 0.8344,0.3038, 0.8503,0.2533, 0.9078,0.2046, 0.8650,0.1950,
    0.9295,0.1710, 0.8881,0.1633, 1,0.1143, 0.9438,0.0935],
    [0.3430,0.6545, 0.3085,0.6257, 0.2925,0.6440, 0.3430,0.6545]],
  mapCanada: [
    [0.1973,0.8772, 0.1453,0.8426, 0.1973,0.8772],
    [0.1893,0.3633, 0.2903,0.3137, 0.1880,0.2941, 0.1893,0.3633],
    [0.4387,0.4170, 0.4093,0.3097, 0.2498,0.3622, 0.3183,0.4479, 0.4387,0.4170],
    [0.2080,0.2388, 0.2899,0.2024, 0.2080,0.2388],
    [0.2853,0.2353, 0.4007,0.2624, 0.2853,0.2353],
    [0.4107,0.1540, 0.4746,0.1773, 0.4107,0.1540],
    [0.4800,0.2249, 0.4143,0.2266, 0.4800,0.2249],
    [0.4720,0.3754, 0.4884,0.2974, 0.4720,0.3754],
    [0.5520,0.2716, 0.6947,0.2612, 0.5110,0.2119, 0.5520,0.2716],
    [0.5067,0.1228, 0.6307,0.1401, 0.5067,0.1228],
    [0.5667,0.0588, 0.6995,0.0804, 0.5802,0.2200, 0.6847,0.2378, 0.9043,0.0327, 0.5667,0.0588],
    [0.6840,0.5675, 0.6080,0.5692, 0.6840,0.5675],
    [0.8467,0.5104, 0.9026,0.4967, 0.7846,0.3721, 0.5794,0.3426, 0.8080,0.5000, 0.7158,0.5530,
    0.8429,0.6118, 0.8467,0.5104],
    [0.6800,0.3114, 0.7307,0.3304, 0.6800,0.3114],
    [0.9920,0.8997, 0.9636,0.8278, 0.9248,0.8970, 0.9920,0.8997],
    [0.9133,0.9273, 0.7924,0.9118, 0.9650,0.8137, 0.8637,0.6407, 0.7147,0.5969, 0.6909,0.8363,
    0.5240,0.6747, 0.5347,0.5606, 0.6734,0.4796, 0.5706,0.4533, 0.5697,0.3030, 0.5122,0.3114,
    0.5163,0.4823, 0.4890,0.4157, 0.3747,0.4956, 0,0.4221, 0,0.6453, 0.2093,0.8789,
    0.6073,0.8785, 0.6627,1, 0.9133,0.9273],
    [0.2653,0.5087, 0.1880,0.5104, 0.2653,0.5087],
    [0.3280,0.6125, 0.2747,0.6280, 0.3280,0.6125],
    [0.5013,0.8478, 0.4805,0.7811, 0.5013,0.8478]],
  mapUsMex: [
    [0.3750,0.3660, 0.3251,0.2814, 0.3026,0.2959, 0.2664,0.2663, 0.2664,0.0447, 0.1113,0.0016,
    0.0168,0.0685, 0.0736,0.1271, 0.0330,0.1228, 0.0012,0.1478, 0.0210,0.1735, 0.0692,0.1699,
    0.0666,0.1942, 0.0166,0.2360, 0.0256,0.2615, 0.0412,0.2820, 0.0561,0.2715, 0.0601,0.3043,
    0.1079,0.2981, 0.0926,0.3422, 0.0491,0.3814, 0.1343,0.3162, 0.1706,0.2457, 0.1798,0.2545,
    0.1590,0.2916, 0.1963,0.2766, 0.1970,0.2499, 0.2815,0.2749, 0.3084,0.3110, 0.3240,0.2952,
    0.3271,0.3179, 0.3131,0.3196, 0.3283,0.3505, 0.3382,0.3175, 0.3341,0.3557, 0.3575,0.3591,
    0.3425,0.3603, 0.3446,0.3849, 0.3703,0.3866, 0.3750,0.3660],
    [0.0444,0.3832, 0.0315,0.3935, 0.0444,0.3832],
    [0.1554,0.3076, 0.1331,0.3422, 0.1554,0.3076],
    [0.1227,0.9777, 0.1227,0.9966, 0.1227,0.9777],
    [0.9918,0.5292, 0.8446,0.6325, 0.8283,0.5564, 0.8096,0.5739, 0.8072,0.6171, 0.7927,0.6173,
    0.7973,0.5640, 0.8318,0.5464, 0.7892,0.5253, 0.7540,0.5378, 0.7757,0.5137, 0.7243,0.4897,
    0.4498,0.4983, 0.4486,0.5292, 0.4287,0.5069, 0.4387,0.6760, 0.5047,0.7852, 0.6098,0.7955,
    0.6405,0.8381, 0.6658,0.8355, 0.6998,0.8900, 0.6998,0.8574, 0.7242,0.8293, 0.8305,0.8209,
    0.8627,0.9020, 0.8582,0.7971, 0.9144,0.7269, 0.9054,0.6667, 0.9117,0.7016, 0.9143,0.6691,
    0.9486,0.6409, 0.9346,0.6388, 1,0.5704, 0.9918,0.5292]],
  mapBit13: [0.8571,0.9375, 0.5397,0.8218, 0,0.3750, 0.0794,0.1736, 0.2540,0.0972, 0.5714,0.3174,
    0.5714,0.6250, 0.8466,0.8148, 0.8571,0.9375],
  iconPhone: [
    [0.7853,0, 0.0633,0.0633, 0.0633,0.9380, 0.9380,0.9380, 0.7853,0],
    [0.8100,0.7476, 0.2018,0.2425, 0.8100,0.7476]],
  iconMail: [
    [0.1448,0.4766, 0.8751,0.9031, 0.1448,0.4766],
    [0.1246,0.3794, 0.9829,0.9498, 0.0617,0.9861, 0.1246,0.3794]],
  iconGlobe: [0.5000,0, 0,0.5000, 0.5000,1, 1,0.5000, 0.5000,0],
  iconPin: [
    [0.4823,0.0003, 0,0.3998, 0.5098,1.0006, 1.0003,0.3994, 0.4823,0.0003],
    [0.5001,0.5982, 0.5001,0.2101, 0.5001,0.5982]]
};

/* =================================================================== slides */

// Slide 1 - cover
function slide01(s) {
  s.background = { color: C.INDIGO };
  title(s, 7.467, 1.246, 5.174, 3.037, 'Hygienic Agriculture Product', 66, C.WHITE);
  s.addText(LOREM_MED, { x: 7.467, y: 4.283, w: 4.554, h: 0.985,
    fontSize: 14, color: C.WHITE, lineSpacingMultiple: 1.3 });
  pill(s, 7.467, 5.733, 4.554, 0.6, C.GOLD);
  s.addText('Presentation Template', { x: 7.86, y: 5.765, w: 4.026, h: 0.489,
    align: 'center', fontSize: 20, bold: true, color: C.WHITE, lineSpacingMultiple: 1.3 });
}

// Slide 2 - section opener with KPI row
function slide02(s, n) {
  chrome(s, n);
  title(s, 0.752, 0.728, 6.184, 2.446, 'Hygienic agriculture ensures the safety \nof consumers', 44);
  body(s, 0.752, 3.132, 5.625, 1.913, LOREM_XL);
  [['545K', 0.797], ['79%', 2.359], ['1289', 3.744]].forEach(([v, x]) => stat(s, x, 5.328, v, C.LEAF));
  circleBadge(s, 8.752, 1.1, 2.335, C.VIOLET, 'Lorem ipsum', 'Lorem ipsum dolor sit amet, consectetur', 1.132);
  iconChip(s, 9.609, 1.433, 0.645);   // spray-bottle glyph in the source
}

// Slide 3 - hygienic practices in farming
function slide03(s, n) {
  chrome(s, n);
  title(s, 0.76, 0.97, 3.728, 2.127, 'Hygienic Practices in Farming', 48);
  body(s, 7.164, 4.591, 5.625, 1.913, LOREM_XL);
  circleBadge(s, 5.202, 2.256, 2.335, C.VIOLET, 'Lorem ipsum', 'Lorem ipsum dolor sit amet, consectetur', 1.132);
  art(s, 'leafHand', 6.057, 2.587, 0.625, 0.645, { fill: { color: C.WHITE } });
}

// Slide 4 - three practices, mint panel on the left
function slide04(s, n) {
  chrome(s, n);
  art(s, 'roundRightPanel', 0, 0.646, 7.625, 6.25, { fill: { color: C.MINT } });
  title(s, 7.906, 1.026, 4.899, 2.127, 'Hygienic Practices in Harvesting', 48);

  const ROWS = [
    { label: 'Utilization of clean and sanitized', pillText: 'Utilization of clean and sanitized',
      color: C.BLUE_DK, dot: C.NAVY, icon: 'molecule' },
    { label: 'Temperature Control', pillText: 'Temperature control',
      color: C.LEAF, dot: C.FOREST, icon: 'leafHand' },
    { label: 'Implementation of HACCP', pillText: 'Implementation of HACCP',
      color: C.MAGENTA, dot: C.MAGENTA_DK, icon: 'molecule' },
  ];
  ROWS.forEach((r, i) => {
    const y = 3.577 + i * 1.0225;
    pill(s, 6.99, y, 5.16, 0.749, r.color);
    s.addText(r.pillText, { x: 7.144, y: y + 0.129, w: 4.229, h: 0.417, align: 'right',
      fontSize: 16, bold: true, color: C.WHITE, lineSpacingMultiple: 1.3 });
    s.addShape('ellipse', { x: 11.563, y: y - 0.022, w: 0.777, h: 0.777, fill: { color: r.dot } });
    iconChip(s, 11.749, y + 0.164, 0.428);
    labelled(s, { x: 0.76, y: 1.437 + i * 1.653, w: 4.899, label: r.label, color: r.color,
      text: LOREM_MED, gap: 0.394, capW: 5.28, capH: 0.859, capSize: 14 });
  });
}

// Slide 5 - weather conditions + progress bar
function slide05(s, n) {
  chrome(s, n);
  title(s, 0.752, 1.052, 5.625, 2.381,
    'PLACEHOLDER', 36);
  body(s, 7.119, 1.385, 5.471, 1.291, LOREM_LONG);
  stat(s, 1.52, 4.028, '79%', C.GREEN);
  body(s, 2.905, 4.031, 2.736, 0.679, S1.trim().replace(/\.$/, ''));

  pill(s, 0.752, 5.136, 6.287, 1.302, C.GREEN_DK);
  s.addText('Lorem Ipsum', { x: 1.664, y: 5.401, w: 3.977, h: 0.316,
    fontSize: 16, bold: true, color: C.WHITE, fontFace: HEAD_FONT, lineSpacingMultiple: 0.8 });
  pill(s, 1.664, 5.789, 4.462, 0.282, C.GREY_10);   // track
  pill(s, 1.664, 5.789, 3.576, 0.282, C.GOLD);      // fill (79%)
  s.addText('79%', { x: 5.401, y: 5.787, w: 0.787, h: 0.303,
    fontSize: 12, bold: true, color: C.GREEN, fontFace: HEAD_FONT });
}

// Slide 6 - two produce categories on a mint card
function slide06(s, n) {
  chrome(s, n);
  s.addShape('roundRect', { x: 0.603, y: 0.684, w: 12.254, h: 4.41, rectRadius: 0.878,
    fill: { color: C.MINT } });
  const COLS = [
    { x: 1.213, cx: 1.264, cy: 3.944, color: C.LEAF_DK, label: 'Organic Vegetables' },
    { x: 7.035, cx: 9.734, cy: 3.908, color: C.MAGENTA, label: 'Fresh Fruits' },
  ];
  COLS.forEach(c => {
    labelled(s, { x: c.x, y: 1.327, w: 4.899, label: c.label, color: c.color,
      text: S1 + S2 + S3 + S4, gap: 0.394, capW: 5.28, capH: 1.291, capSize: 14 });
    circleBadge(s, c.cx, c.cy, 2.335, c.color, c.label,
      'Lorem ipsum dolor sit amet, consectetur , id mattis odio posuere.', 0.639);
  });
}

// Slide 7 - photo strip
function slide07(s, n) {
  chrome(s, n);
  title(s, 9.356, 1.451, 3.217, 3.737, 'Sorted and packed in hygienic conditions to maintain freshness', 40);
  body(s, 0.76, 5.833, 11.406, 0.679, LOREM_LONG);
}

// Slide 8 - same headline with two donut cards
function slide08(s, n) {
  chrome(s, n);
  title(s, 0.727, 0.952, 3.217, 3.737, 'Sorted and packed in hygienic conditions to maintain freshness', 40);
  body(s, 0.76, 4.934, 2.781, 1.291, S1 + S2);

  [{ y: 1.166, fill: C.VIOLET }, { y: 3.375, fill: C.LEAF }].forEach(card => {
    pill(s, 8.431, card.y, 4.017, 1.598, card.fill);
    const dy = card.y + 0.29;
    s.addShape('pie', { x: 8.826, y: dy, w: 0.95, h: 0.965, angleRange: [169.3, 271.5], fill: { color: C.GREY_12 } });
    s.addShape('pie', { x: 8.826, y: dy, w: 0.95, h: 0.965, angleRange: [271.3, 169.8], fill: { color: C.GOLD } });
    s.addShape('ellipse', { x: 8.987, y: dy + 0.163, w: 0.629, h: 0.639, fill: { color: C.GREY_05 } });
    s.addText('75%', { x: 8.839, y: dy + 0.249, w: 0.95, h: 0.454, align: 'center',
      fontSize: 18, bold: true, color: C.BLACK, lineSpacingMultiple: 1.3 });
    s.addText('Lorem Ipsum', { x: 9.889, y: card.y + 0.216, w: 2.559, h: 0.411,
      fontSize: 14, bold: true, color: C.WHITE, lineSpacingMultiple: 1.5 });
    s.addText(S1.trim(), { x: 9.889, y: card.y + 0.586, w: 2.559, h: 0.669,
      fontSize: 12, color: C.WHITE, lineSpacingMultiple: 1.5 });
  });
}

// Slide 9 - two feature cards with icons and paired stats
function slide09(s, n) {
  chrome(s, n);
  s.addText('PLACEHOLDER',
    { x: 1.947, y: 0.86, w: 9.44, h: 1.227, align: 'center', fontSize: 36, bold: true,
      color: C.INDIGO, fontFace: HEAD_FONT, lineSpacingMultiple: 0.9 });

  const CARDS = [
    { x: 0.721, fill: C.LEAF, radius: 0.572, label: 'Nutritional Information', a: '+96', b: '+120' },
    { x: 4.708, fill: C.VIOLET, radius: 0.471, label: 'Batch Codes', a: '87+', b: '+145' },
  ];
  CARDS.forEach(c => {
    s.addShape('roundRect', { x: c.x, y: 2.4, w: 3.806, h: 4.115, rectRadius: c.radius, fill: { color: c.fill } });
    s.addText(c.label, { x: c.x + 0.277, y: 3.811, w: 3.244, h: 0.38, align: 'center',
      fontSize: 14, bold: true, color: C.WHITE, lineSpacingMultiple: 1.3 });
    s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna.',
      { x: c.x + 0.284, y: 4.359, w: 3.236, h: 1.046, align: 'center', fontSize: 12,
        color: C.WHITE, lineSpacingMultiple: 1.2 });
    [[c.a, 'Statistic One', c.x + 0.284], [c.b, 'Statistic Two', c.x + 1.937]].forEach(([v, cap, x]) => {
      s.addText(v, { x, y: 5.384, w: 1.594, h: 0.502, align: 'center', valign: 'bottom',
        fontSize: 20, bold: true, color: C.WHITE, lineSpacingMultiple: 1.3 });
      s.addText(cap, { x, y: 5.851, w: 1.591, h: 0.341, align: 'center',
        fontSize: 12, color: C.WHITE, lineSpacingMultiple: 1.3 });
    });
  });

  // Line-art pie icon (left card) and calendar icon (right card).
  art(s, 'pieSliceA', 2.365, 2.786, 0.293, 0.391, { fill: { color: C.WHITE } });
  art(s, 'pieSliceB', 2.247, 2.951, 0.391, 0.489, { fill: { color: C.WHITE } });
  art(s, 'pieRing', 2.377, 2.885, 0.685, 0.734, { fill: { color: C.WHITE } });
  art(s, 'calBody', 6.266, 2.849, 0.768, 0.831, { fill: { color: C.WHITE } });
  art(s, 'calInner', 6.325, 3.181, 0.644, 0.436, { fill: { color: C.WHITE } });
  art(s, 'calDigit1', 6.455, 3.242, 0.145, 0.312, { fill: { color: C.WHITE } });
  art(s, 'calDigit5', 6.615, 3.241, 0.229, 0.312, { fill: { color: C.WHITE } });

  stat(s, 8.799, 3.208, '745K', C.BLACK);
  stat(s, 10.106, 3.208, '89%', C.BLACK);
  body(s, 8.799, 4.24, 4.187, 1.904,
    S1 + 'In it ipsum sed ante commodo, id mattis odio il ' + S4.trim() + ' amet, consectetur adipiscing elit. ' + S6);
}

// Slide 10 - labeling KPIs
function slide10(s, n) {
  chrome(s, n);
  title(s, 0.721, 1.289, 7.016, 2.044, 'PLACEHOLDER', 40);
  body(s, 0.721, 3.444, 5.625, 1.598, LOREM_LONG + '. ' + S1);
  [['663K', 0.734], ['47%', 2.296], ['1679', 3.681]].forEach(([v, x]) => stat(s, x, 5.33, v, C.LEAF));
  circleBadge(s, 6.569, 4.19, 2.335, C.BLUE_DK, 'Labeling',
    'Lorem ipsum dolor sit amet, consectetur, id mattis odio posuere.', 0.639);
}

// Slide 11 - three awareness steps
function slide11(s, n) {
  chrome(s, n);
  s.addShape('roundRect', { x: 7.557, y: 2.524, w: 5.271, h: 4.21, rectRadius: 0.527,
    fill: { color: C.MINT } });
  title(s, 4.995, 1.002, 5.968, 1.305, 'Consumer Awareness and Education', 40);
  [['01. Providing Information', C.BLUE_DK], ['02. Educating Consumers', C.LEAF],
   ['03. Encouraging Support', C.GOLD]].forEach(([label, color], i) => {
    const y = 3.042 + i * 1.0225;
    pill(s, 4.17, y, 4.035, 0.749, color);
    s.addText(label, { x: 4.717, y: y + 0.111, w: 3.161, h: 0.417,
      fontSize: 16, bold: true, color: C.WHITE, lineSpacingMultiple: 1.3 });
  });
  body(s, 8.505, 3.103, 3.689, 2.823,
    S1 + S2 + 'Neu mauris non arcu blandit cursus. ' + S4 + S5 + '. ' + S1 + S6);
}

// Slide 12 - website mock-up with a segmented ring
function slide12(s, n) {
  chrome(s, n);
  s.addShape('roundRect', { x: 0.391, y: 0.71, w: 5.725, h: 6.027, rectRadius: 0.62,
    fill: { color: C.MINT } });
  photo(s, 5.412, 1.351, 7.154, 4.802);   // browser mock-up screenshot
  title(s, 0.803, 1.546, 4.309, 1.377, 'Access through our website', 40);
  body(s, 0.891, 3.1, 4.132, 1.904, LOREM_LONG);
  stat(s, 0.939, 5.296, '26K', C.VIOLET);
  stat(s, 2.413, 5.296, '15K', C.VIOLET);

  s.addShape('ellipse', { x: 4.55, y: 4.727, w: 2.066, h: 2.066, fill: { color: C.VIOLET } });
  // Six wedges alternating value / gap around the ring.
  [[183.9, 266.3, C.GREY_10], [36.2, 185.1, C.GOLD], [356.2, 35.9, C.GREY_10],
   [321.2, 356.7, C.LEAF], [291.8, 321.5, C.GREY_10], [265.7, 299.1, C.MAGENTA],
  ].forEach(([a1, a2, color]) => {
    s.addShape('pie', { x: 4.876, y: 5.025, w: 1.401, h: 1.401, angleRange: [a1, a2],
      fill: { color, transparency: color === C.GREY_10 ? 45 : 0 } });
  });
  s.addShape('ellipse', { x: 4.949, y: 5.097, w: 1.256, h: 1.256, fill: { color: C.WHITE } });
  s.addText('25K', { x: 4.958, y: 5.347, w: 1.237, h: 0.572, align: 'center',
    fontSize: 28, bold: true, fontFace: HEAD_FONT });
  s.addText('On value', { x: 4.958, y: 5.782, w: 1.237, h: 0.337, align: 'center', fontSize: 14 });
}

// Slide 13 - mobile app with two rating gauges
function slide13(s, n) {
  chrome(s, n);
  s.addShape('roundRect', { x: 6.45, y: 0.651, w: 6.191, h: 6.33, rectRadius: 0.961,
    fill: { color: C.MINT } });
  photo(s, 5.32, 1.308, 2.668, 5.193);   // phone mock-up screenshot
  title(s, 0.82, 1.438, 3.966, 1.935, 'Access through our Mobile App', 40);
  body(s, 0.82, 3.555, 4.132, 1.904, LOREM_LONG);

  [{ y: 1.958, color: C.LEAF, score: '8.5', sweep: 49.0, w: 2.771 },
   { y: 4.004, color: C.VIOLET, score: '9.5', sweep: 32.2, w: 3.065 },
  ].forEach(g => {
    s.addShape('ellipse', { x: 8.349, y: g.y, w: 1.438, h: 1.438, fill: { color: g.color } });
    s.addShape('ellipse', { x: 8.616, y: g.y + 0.267, w: 0.905, h: 0.905, fill: { color: g.color } });
    const ax = 8.46, ay = g.y + 0.111;
    s.addShape('arc', { x: ax, y: ay, w: 1.216, h: 1.216, angleRange: [128.6, 49.0],
      fill: { color: g.color }, line: { color: C.GREY_45, width: 4, transparency: 75 } });
    s.addShape('arc', { x: ax, y: ay, w: 1.216, h: 1.216, angleRange: [128.6, g.sweep],
      fill: { color: g.color }, line: { color: C.WHITE, width: 4 } });
    s.addText('Rating', { x: 8.533, y: g.y + 0.423, w: 1.07, h: 0.286, align: 'center',
      fontSize: 11, bold: true, color: C.WHITE, fontFace: HEAD_FONT });
    s.addText(g.score, { x: 8.533, y: g.y + 0.587, w: 1.07, h: 0.505, align: 'center',
      fontSize: 24, bold: true, color: C.WHITE, fontFace: HEAD_FONT });
    s.addText(S1, { x: 9.943, y: g.y + 0.423, w: g.w, h: 0.671, fontSize: 12, lineSpacingMultiple: 1.5 });
  });
}

// Slide 14 - pricing table
function slide14(s, n) {
  chrome(s, n);
  title(s, 0.737, 1.989, 4.574, 0.835, 'Table Price', 48);
  const PACKAGE_BLURB = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, ullamco ' +
    'consectetur adipiscing elit. Integer vitae justo ullamcorper, sceler';
  labelled(s, { x: 0.838, y: 3.218, w: 4.774, label: 'Silver Package', color: C.INDIGO,
    text: PACKAGE_BLURB, capH: 0.863 });
  labelled(s, { x: 0.838, y: 4.646, w: 4.774, label: 'Gold Package', color: C.LEAF_DK,
    text: PACKAGE_BLURB, capH: 0.863 });

  [{ x: 5.769, fill: C.INDIGO, name: 'SILVER PACKAGE', price: '56.00' },
   { x: 9.357, fill: C.LEAF_DK, name: 'GOLD PACKAGE', price: '68.00' },
  ].forEach(card => {
    s.addShape('roundRect', { x: card.x, y: 0.83, w: 3.239, h: 5.84, rectRadius: 1.62,
      fill: { color: card.fill } });
    s.addText(card.name, { x: card.x + 0.744, y: 1.769, w: 1.743, h: 0.572, align: 'center',
      fontSize: 14, bold: true, charSpacing: 3, color: C.WHITE, fontFace: HEAD_FONT });
    s.addText([{ text: '$', options: { superscript: true } }, { text: card.price }],
      { x: card.x + 0.449, y: 2.658, w: 2.333, h: 0.774, align: 'center',
        fontSize: 40, color: C.WHITE, fontFace: HEAD_FONT });
    s.addText('Vivamus elementum semper nisi. Aenean vulputate.',
      { x: card.x + 0.348, y: 3.75, w: 2.534, h: 0.985, align: 'center',
        fontSize: 14, color: C.WHITE, lineSpacingMultiple: 1.3 });
    pill(s, card.x + 0.618, 5.373, 1.994, 0.601, C.GOLD, 'BOOK NOW', { fontSize: 18 });
  });
}

// Slide 15 - 3D pyramid, four tiers
function slide15(s, n) {
  chrome(s, n);
  title(s, 0.756, 0.99, 3.819, 2.898, 'Produce  Super\nHygienic Products', 48);
  note(s, 0.819, 4.189, 2.319, 1.646, LOREM_MED.trim());

  // Tiers are drawn bottom-up: left face, right face, top face, then the icon.
  art(s, 'pyrFrontA', 4.077, 4.264, 4.162, 1.731, { fill: { color: C.PERI } });
  art(s, 'pyrRightA', 6.367, 5.012, 2.391, 1.846, { fill: { color: C.BLUE } });
  art(s, 'pyrLeftA', 3.527, 5.065, 2.897, 1.794, { fill: { color: C.BLUE_DK } });
  iconChip(s, 4.726, 5.65, 0.5);
  art(s, 'pyrFrontB', 4.729, 3.362, 2.87, 1.016, { fill: { color: C.ORCHID } });
  art(s, 'pyrRightB', 6.304, 3.817, 1.811, 1.592, { fill: { color: C.VIOLET } });
  art(s, 'pyrLeftB', 4.188, 3.853, 2.175, 1.556, { fill: { color: C.INDIGO } });
  iconChip(s, 5.068, 4.336, 0.5);
  art(s, 'pyrFrontC', 5.399, 2.391, 1.551, 0.667, { fill: { color: C.SPRING } });
  art(s, 'pyrRightC', 6.245, 2.644, 1.221, 1.348, { fill: { color: C.LEAF } });
  art(s, 'pyrLeftC', 4.853, 2.66, 1.451, 1.333, { fill: { color: C.LEAF_DK } });
  iconChip(s, 5.426, 3.084, 0.5);
  art(s, 'pyrTipRight', 6.181, 1.357, 0.645, 1.278, { fill: { color: C.GREEN } });
  art(s, 'pyrTipLeft', 5.519, 1.357, 0.724, 1.274, { fill: { color: C.GREEN_DK } });
  iconChip(s, 5.769, 1.889, 0.408);

  const TIERS = [
    { num: '01', bx: 6.297, by: 1.461, ring: C.GREEN_DK, tx: 7.165, ty: 1.151, color: C.GREEN, w: 4.162 },
    { num: '02', bx: 6.949, by: 2.685, ring: C.LEAF_DK, tx: 7.761, ty: 2.401, color: C.LEAF, w: 4.162 },
    { num: '03', bx: 7.602, by: 3.91, ring: C.INDIGO, tx: 8.397, ty: 3.651, color: C.VIOLET, w: 4.162 },
    { num: '04', bx: 8.255, by: 5.134, ring: C.BLUE_DK, tx: 9.066, ty: 4.902, color: C.BLUE_DK, w: 3.61 },
  ];
  TIERS.forEach(t => {
    s.addShape('ellipse', { x: t.bx, y: t.by, w: 0.485, h: 0.485, fill: { color: C.WHITE } });
    s.addShape('ellipse', { x: t.bx + 0.045, y: t.by + 0.045, w: 0.395, h: 0.395, fill: { color: t.ring } });
    s.addText(t.num, { x: t.bx - 0.082, y: t.by + 0.099, w: 0.649, h: 0.286, align: 'center',
      valign: 'middle', fontSize: 10.5, bold: true, color: C.WHITE });
    labelled(s, { x: t.tx, y: t.ty, w: 4.162, label: 'Lorem Ipsum', color: t.color,
      text: S1 + 'In it ipsum sed ante commodo, id mattis odio', gap: 0.395, capW: t.w, capH: 0.596 });
  });
}

// Slide 16 - SWOT petals
function slide16(s, n) {
  chrome(s, n);
  s.addShape('ellipse', { x: 4.21, y: 1.433, w: 4.759, h: 4.759,
    fill: { color: C.WHITE, transparency: 60 } });
  const QUADS = [
    { letter: 'S', word: 'Strength', fill: C.VIOLET, x: 3.914, y: 0.617, tx: 3.758, ty: 0.852 },
    { letter: 'W', word: 'Weakness', fill: C.BLUE, x: 6.9, y: 0.617, tx: 6.897, ty: 0.852, flipH: true },
    { letter: 'O', word: 'Opportunities', fill: C.GOLD, x: 3.914, y: 3.84, tx: 3.758, ty: 4.227, flipV: true },
    { letter: 'T', word: 'Threats', fill: C.LEAF, x: 6.9, y: 3.84, tx: 6.897, ty: 4.227, flipH: true, flipV: true },
  ];
  QUADS.forEach(q => {
    art(s, 'swotPetal', q.x, q.y, 2.507, 2.725,
      { fill: { color: q.fill }, flipH: q.flipH, flipV: q.flipV });
    s.addText(q.letter, { x: q.tx, y: q.ty, w: 2.669, h: 1.717, align: 'center',
      fontSize: 96, color: C.WHITE, fontFace: HEAD_FONT });
    s.addText(q.word, { x: q.tx, y: q.ty + 1.467, w: 2.669, h: 0.404, align: 'center',
      fontSize: 18, bold: true, color: C.WHITE });
  });
  pill(s, 5.811, 3.279, 1.812, 0.597, C.GREEN, 'SWOT', { fontSize: 20 });

  const NOTES = [
    { x: 0.886, y: 1.735, label: 'Strength', color: C.INDIGO, align: 'right' },
    { x: 0.861, y: 4.832, label: 'Opportunities', color: C.GOLD, align: 'right' },
    { x: 9.753, y: 1.735, label: 'Weakness', color: C.BLUE, align: 'left' },
    { x: 9.728, y: 4.832, label: 'Threats', color: C.LEAF, align: 'left' },
  ];
  NOTES.forEach(t => labelled(s, { x: t.x, y: t.y, w: 2.507, label: t.label, color: t.color,
    text: LOREM_SHORT, align: t.align }));
}

// Slide 17 - weekly agenda grid
function slide17(s, n) {
  chrome(s, n);
  s.addShape('roundRect', { x: 0.982, y: 1.75, w: 11.383, h: 4.889, rectRadius: 0.313,
    fill: { color: C.LILAC } });
  s.addText('Agenda Slide', { x: 3.578, y: 0.739, w: 6.177, h: 0.774, align: 'center',
    fontSize: 48, bold: true, color: C.INDIGO, fontFace: HEAD_FONT, lineSpacingMultiple: 0.9 });

  const COL_X = [1.215, 3.035, 4.854, 6.674, 8.493, 10.313];
  const ROW_Y = [2.511, 3.702, 4.892];
  const CELL_W = 1.819, CELL_H = 1.19;
  const HAIR = { color: C.GREY_20 };

  // Header strip and the day dividers that sit inside it.
  s.addShape('rect', { x: 1.215, y: 1.929, w: 10.917, h: 0.576, fill: { color: C.GREY_10 }, line: HAIR });
  COL_X.slice(1).forEach(x => s.addShape('line', { x, y: 1.939, w: 0, h: 0.591, line: { color: C.GREY_20, width: 1 } }));

  // Body grid - alternating cells carry a 40% white wash.
  ROW_Y.forEach((y, r) => COL_X.forEach((x, c) => {
    const shaded = (r + c) % 2 === 0;
    s.addShape('rect', { x, y, w: CELL_W, h: CELL_H, line: HAIR,
      fill: shaded ? { color: C.GREY_10, transparency: 60 } : { type: 'none' } });
  }));

  ['19 Mon', '20 Tue', '21 Wed', '22 Thu', '23 Fri', '24 Sat'].forEach((d, i) => {
    const [num, day] = d.split(' ');
    s.addText([{ text: num + ' ', options: { fontSize: 24 } }, { text: day, options: { fontSize: 14 } }],
      { x: COL_X[i] + 0.011, y: 1.925, w: CELL_W, h: 0.552, align: 'center', lineSpacingMultiple: 1.2 });
  });

  const RAISE = 'We want to raise your crops and livestock';
  // Plain entries: [column, row, time, copy]
  [[0, 1, '09.00 AM', RAISE], [1, 0, '06.00 AM', RAISE], [1, 1, '10.00 AM', RAISE],
   [1, 2, '02.00 PM', RAISE], [2, 0, '04.00 AM', RAISE], [3, 0, '09.00 AM', RAISE],
   [4, 0, '11.00 AM', RAISE], [5, 0, '07.00 AM', RAISE], [5, 2, '04.00 PM', RAISE],
  ].forEach(([c, r, time, copy]) => {
    const x = COL_X[c] + 0.1, y = ROW_Y[r] + 0.153;
    s.addText(time, { x, y, w: 1.599, h: 0.307, fontSize: 11, lineSpacingMultiple: 1.2 });
    s.addText(copy, { x, y: y + 0.229, w: 1.599, h: 0.656, fontSize: 11, lineSpacingMultiple: 1.3 });
  });

  // Highlighted entries sit on a coloured block covering their cell.
  [[0, 0, C.GREEN, '04.00 AM', RAISE], [3, 1, C.VIOLET, '03.00 PM', 'Farming the future'],
   [5, 1, C.GREEN, '09.00 AM', 'The best acticulture product for you'],
   [1, 2, C.GOLD, '12.00 AM', RAISE],
  ].forEach(([c, r, color, time, copy]) => {
    const x = COL_X[c], y = ROW_Y[r];
    s.addShape('rect', { x, y: y - 0.012, w: CELL_W, h: 1.205, fill: { color } });
    s.addText(time, { x: x + 0.121, y: y + 0.16, w: 1.51, h: 0.3,
      fontSize: 11, color: C.WHITE, lineSpacingMultiple: 1.2 });
    s.addText(copy, { x: x + 0.121, y: y + 0.335, w: 1.51, h: 0.641,
      fontSize: 11, color: C.WHITE, lineSpacingMultiple: 1.3 });
  });

  art(s, 'infoBadge', 1.226, 6.24, 0.242, 0.242, { fill: { color: C.WHITE } });
  s.addText('Lorem ipsum dolor sit amet, consectetuer adipiscing elit.',
    { x: 1.517, y: 6.192, w: 3.786, h: 0.305, fontSize: 10.5, italic: true, lineSpacingMultiple: 1.3 });
}

// Slide 18 - four-stage funnel
function slide18(s, n) {
  chrome(s, n);
  s.addText('Funnel Infographic', { x: 1.925, y: 0.734, w: 9.484, h: 0.842, align: 'center',
    fontSize: 48, bold: true, color: C.INDIGO, fontFace: HEAD_FONT, lineSpacingMultiple: 0.9 });

  art(s, 'funnelSpout', 6.051, 5.966, 1.232, 0.924, { fill: { color: C.CHARCOAL } });
  art(s, 'funnelTL', 4.117, 2.034, 2.55, 1.966, { fill: { color: C.GOLD } });
  art(s, 'funnelTR', 6.667, 2.034, 2.55, 1.966, { fill: { color: C.VIOLET } });
  art(s, 'funnelBL', 5.054, 4.001, 1.613, 1.966, { fill: { color: C.BLUE_DK } });
  art(s, 'funnelBR', 6.667, 4.001, 1.612, 1.966, { fill: { color: C.LEAF } });
  [5.476, 7.316].forEach(x => iconChip(s, x, 2.798, 0.5));
  [5.828, 6.973].forEach(x => iconChip(s, x, 4.595, 0.5));

  // Leader lines: a dot on the outside edge joined to the funnel by a hairline bar.
  const LEGS = [
    { dot: 3.248, dy: 3.053, bar: 3.355, bw: 2.058, by: 3.145, color: C.GOLD },
    { dot: 3.248, dy: 4.988, bar: 3.355, bw: 2.254, by: 5.08, color: C.BLUE },
    { dot: 9.857, dy: 3.02, bar: 7.886, bw: 2.077, by: 3.108, color: C.VIOLET },
    { dot: 9.863, dy: 4.988, bar: 7.716, bw: 2.254, by: 5.08, color: C.LEAF },
  ];
  LEGS.forEach(l => {
    art(s, 'dot', l.dot, l.dy, 0.214, 0.214, { fill: { color: l.color } });
    s.addShape('rect', { x: l.bar, y: l.by, w: l.bw, h: 0.029, fill: { color: l.color } });
  });

  const CAPTION = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer. ';
  [{ x: 0.553, y: 3.429, label: 'One', color: C.GOLD, align: 'right' },
   { x: 0.553, y: 5.289, label: 'Two', color: C.BLUE_DK, align: 'right' },
   { x: 9.607, y: 3.429, label: 'Three', color: C.VIOLET, align: 'left' },
   { x: 9.607, y: 5.289, label: 'Four', color: C.LEAF, align: 'left' },
  ].forEach(t => {
    s.addText(t.label, { x: t.x, y: t.y, w: 3.001, h: 0.37, align: t.align,
      fontSize: 16, bold: true, color: t.color });
    s.addText(CAPTION, { x: t.x, y: t.y + 0.346, w: 3.001, h: 0.673, align: t.align,
      fontSize: 12, lineSpacingMultiple: 1.5 });
  });
}

// Slide 19 - two interlocking circular arrow cycles
function slide19(s, n) {
  chrome(s, n);
  s.addText('Creative Infographic', { x: 2.597, y: 1.046, w: 8.134, h: 0.902, align: 'center',
    fontSize: 48, bold: true, color: C.INDIGO, fontFace: HEAD_FONT, lineSpacingMultiple: 0.9 });

  // Each segment = a 3/4 block arc plus a right-triangle arrow head.
  // The two cycles are mirrored copies, so every piece carries flipH.
  const SEGMENTS = [
    { color: C.LEAF,    arcX: 3.719, arcRot: 270, triX: 3.611, triY: 4.257, triRot: 315, flipV: true, num: '01', nx: 4.191, ny: 5.367, nw: 0.482 },
    { color: C.LEAF_DK, arcX: 3.719, arcRot: 0,   triX: 4.943, triY: 2.902, triRot: 315, num: '02', nx: 4.179, ny: 3.470, nw: 0.498 },
    { color: C.FOREST,  arcX: 3.719, arcRot: 90,  triX: 6.283, triY: 4.234, triRot: 45,  num: '03', nx: 6.055, ny: 3.470, nw: 0.496 },
    { color: C.PLUM,    arcX: 6.391, arcRot: 270, triX: 7.614, triY: 5.588, triRot: 315, num: '04', nx: 6.830, ny: 5.367, nw: 0.500 },
    { color: C.INDIGO,  arcX: 6.391, arcRot: 180, triX: 8.960, triY: 4.254, triRot: 315, flipV: true, num: '05', nx: 8.701, ny: 5.367, nw: 0.496 },
    { color: C.VIOLET,  arcX: 6.391, arcRot: 90,  triX: 7.638, triY: 2.902, triRot: 45,  num: '06', nx: 8.698, ny: 3.470, nw: 0.509 },
  ];
  SEGMENTS.forEach(g => {
    s.addShape('blockArc', { x: g.arcX, y: 3.01, w: 3.233, h: 3.233, rotate: g.arcRot, flipH: true,
      angleRange: [270, 0], arcThicknessRatio: 0.35, fill: { color: g.color } });
    s.addShape('rtTriangle', { x: g.triX, y: g.triY, w: 0.762, h: 0.762, rotate: g.triRot,
      flipH: true, flipV: g.flipV, fill: { color: g.color } });
  });
  SEGMENTS.forEach(g => {
    s.addText(g.num, { x: g.nx, y: g.ny, w: g.nw, h: 0.404, align: 'center', valign: 'middle',
      fontSize: 18, bold: true, color: C.WHITE, fontFace: HEAD_FONT, wrap: false });
  });

  [4.625, 7.328].forEach(x => art(s, 'hubRing', x, 3.94, 1.373, 1.373, { fill: { color: C.LILAC_LT } }));
  s.addText('01', { x: 4.849, y: 4.273, w: 0.836, h: 0.707, align: 'center', valign: 'middle',
    fontSize: 36, bold: true, color: C.WHITE, fontFace: HEAD_FONT });
  s.addText('02', { x: 7.539, y: 4.261, w: 0.9, h: 0.707, align: 'center', valign: 'middle',
    fontSize: 36, bold: true, color: C.WHITE, fontFace: HEAD_FONT });

  [[0.806, 2.678, C.FOREST, 'right'], [0.806, 3.958, C.LEAF_DK, 'right'], [0.806, 5.238, C.LEAF, 'right'],
   [9.937, 2.678, C.VIOLET, 'left'], [9.937, 3.958, C.INDIGO, 'left'], [9.937, 5.238, C.PLUM, 'left'],
  ].forEach(([x, y, color, align]) => labelled(s, { x, y, w: 2.507, label: 'Lorem Ipsum',
    color, text: LOREM_SHORT, align }));
}

// Slide 20 - three-node cycle diagram
function slide20(s, n) {
  chrome(s, n);
  title(s, 0.727, 1.189, 3.254, 1.477, 'Creative Diagram', 48);
  s.addShape('ellipse', { x: 4.156, y: 1.401, w: 4.934, h: 5.056, fill: { color: C.GREY_20, transparency: 75 },
    line: { color: C.GREY_50, width: 1.5, dashType: 'dash', transparency: 60 } });
  [{ x: 4.301, y: 2.653, rot: 30, color: C.MAGENTA },
   { x: 8.602, y: 2.599, rot: 330, color: C.LEAF, flipV: true },
   { x: 6.435, y: 6.363, rot: 90, color: C.VIOLET, flipV: true },
  ].forEach(t => s.addShape('triangle', { x: t.x, y: t.y, w: 0.305, h: 0.18,
      rotate: t.rot, flipV: t.flipV, fill: { color: t.color } }));

  s.addShape('ellipse', { x: 5.374, y: 2.677, w: 2.525, h: 2.588, fill: { color: C.LILAC } });
  s.addText('ZoneType', { x: 5.374, y: 2.677, w: 2.525, h: 2.588, align: 'center', valign: 'middle',
    fontSize: 18, bold: true, color: C.WHITE });

  const NODES = [
    { x: 5.991, y: 0.894, fill: C.MAGENTA, label: 'Type 01', tx: 7.836, ty: 0.428, align: 'left', color: C.MAGENTA },
    { x: 7.91, y: 4.901, fill: C.LEAF, label: 'Type 02', tx: 9.665, ty: 5.115, align: 'left', color: C.GREEN },
    { x: 4.098, y: 4.901, fill: C.VIOLET, label: 'Type 03', tx: 1.162, ty: 5.115, align: 'right', color: C.VIOLET },
  ];
  NODES.forEach(nd => {
    s.addShape('ellipse', { x: nd.x, y: nd.y, w: 1.326, h: 1.359, fill: { color: nd.fill } });
    s.addText(nd.label, { x: nd.x, y: nd.y, w: 1.326, h: 1.359, align: 'center', valign: 'middle',
      fontSize: 12, bold: true, color: C.WHITE });
    labelled(s, { x: nd.tx, y: nd.ty, w: 2.507, label: nd.label, color: nd.color,
      text: LOREM_SHORT, align: nd.align });
  });
}

// Slide 21 - percent-stacked bar chart on a lilac card
function slide21(s, n) {
  chrome(s, n);
  s.addText('Monthly Data Report', { x: 1.155, y: 1.136, w: 11.024, h: 0.842, align: 'center',
    fontSize: 48, bold: true, color: C.INDIGO, fontFace: HEAD_FONT, lineSpacingMultiple: 0.9 });
  s.addShape('roundRect', { x: 1.227, y: 2.418, w: 10.88, h: 3.699, rectRadius: 0.314,
    fill: { color: C.LILAC } });
  s.addText('Market Data Base', { x: 4.618, y: 2.778, w: 4.097, h: 0.472, align: 'center',
    fontSize: 18, fontFace: HEAD_FONT, lineSpacingMultiple: 1.3 });

  const CATS = ['Data D', 'Data C', 'Data B', 'Data A'];
  s.addChart([
    { type: pptx.ChartType.bar,
      data: [{ name: 'Series 1', labels: [CATS], values: [65, 41, 77, 63] }],
      options: { chartColors: [C.MAGENTA, C.LEAF_DK, C.BLUE_DK, C.GOLD] } },
    { type: pptx.ChartType.bar,
      data: [{ name: 'Series 2', labels: [CATS], values: [35, 59, 23, 37] }],
      options: { chartColors: [C.GREY_10] } },
  ], {
    x: 1.61, y: 3.201, w: 10.286, h: 2.392,
    barDir: 'bar', barGrouping: 'percentStacked', barGapWidthPct: 87, barOverlapPct: 100,
    showLegend: false, showTitle: false, valAxisHidden: true,
    catAxisLabelFontSize: 14, catAxisLabelColor: C.BLACK,
    valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
    plotArea: { fill: { color: C.LILAC } }, chartArea: { fill: { color: C.LILAC } },
  });
  ['63%', '77%', '41%', '65%'].forEach((v, i) => {
    s.addText(v, { x: 11.117, y: 3.458 + i * 0.505, w: 0.514, h: 0.303,
      align: 'center', fontSize: 12, wrap: false });
  });

  art(s, 'noteBar', 1.231, 5.659, 10.871, 0.694, { fill: { color: C.INDIGO } });
  s.addText([{ text: 'Note : ', options: { bold: true } },
             { text: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue', options: { italic: true } }],
    { x: 1.959, y: 5.791, w: 9.415, h: 0.385, align: 'center', fontSize: 14, color: C.WHITE, lineSpacingMultiple: 1.3 });
}

// Slide 22 - pie infographic with callouts and four mini donuts
function slide22(s, n) {
  chrome(s, n);
  title(s, 7.296, 1.166, 4.529, 1.434, 'Pie Chart \nInfographic', 48);
  note(s, 7.343, 2.696, 5.047, 0.596,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, veniam minim as tempor incididunt ut labore et dolore magna aliqua. ');

  // Callout leaders around the pie.
  const LEADER = { color: C.GREY_50, width: 0.75 };
  [[1.801, 3.047, 0.7, 0], [5.861, 3.978, 0.512, 0], [3.093, 0.974, 0.721, 0],
   [3.093, 0.973, 0, 1.391], [6.373, 3.982, 0, 1.496], [3.541, 5.412, 0.546, 0], [2.821, 4.693, 0, 1.985],
  ].forEach(([x, y, w, h]) => s.addShape('line', { x, y, w, h, line: LEADER }));

  // Pie wedges (start, end in degrees) around a common centre.
  [[316.5, 71.0, C.LEAF], [315.3, 317.4, C.GREY_10], [218.4, 317.4, C.VIOLET],
   [169.0, 219.0, C.GOLD], [70.7, 169.3, C.MAGENTA],
  ].forEach(([a1, a2, color]) => s.addShape('pie', { x: 2.217, y: 1.613, w: 3.933, h: 3.933,
      angleRange: [a1, a2], fill: { color } }));
  [{ v: '25', x: 3.661, y: 2.008, cx: 3.651, cy: 2.538, cw: 0.826, size: 32 },
   { v: '35', x: 4.878, y: 3.464, cx: 4.878, cy: 3.978, cw: 0.826, size: 28 },
   { v: '25', x: 3.155, y: 4.027, cx: 3.182, cy: 4.528, cw: 0.826, size: 32 },
   { v: '15', x: 2.501, y: 2.751, cx: 2.366, cy: 3.286, cw: 1.075, size: 32 },
  ].forEach(l => {
    s.addText(l.v, { x: l.x, y: l.y, w: 0.826, h: 0.639, align: 'center', fontSize: l.size, color: C.WHITE });
    s.addText('Lorem Ipsum', { x: l.cx, y: l.cy, w: l.cw, h: 0.438, align: 'center', fontSize: 10, color: C.WHITE });
  });

  // Four mini gauges on the right - value arc over a grey donut.
  [{ x: 7.401, y: 3.79, pct: '25%', sweep: 109.4, color: C.VIOLET },
   { x: 9.745, y: 3.79, pct: '35%', sweep: 32.9, color: C.LEAF },
   { x: 7.401, y: 5.15, pct: '25%', sweep: 66.1, color: C.MAGENTA },
   { x: 9.745, y: 5.15, pct: '15%', sweep: 179.4, color: C.GOLD },
  ].forEach(g => {
    s.addShape('donut', { x: g.x, y: g.y, w: 0.808, h: 0.808, arcThicknessRatio: 0.134,
      flipH: true, fill: { color: C.GREY_15 } });
    s.addShape('blockArc', { x: g.x, y: g.y, w: 0.808, h: 0.808, rotate: 90, flipH: true,
      angleRange: [g.sweep, 353.3], arcThicknessRatio: 0.134, fill: { color: g.color } });
    s.addText(g.pct, { x: g.x + 0.135, y: g.y + 0.25, w: 0.637, h: 0.337,
      fontSize: 14, bold: true, color: g.color, fontFace: HEAD_FONT, wrap: false });
    s.addText('Lorem Ipsum', { x: g.x + 0.908, y: g.y + 0.118, w: 1.235, h: 0.572, fontSize: 14 });
  });

  [[0.661, 1.687, 1.556, 1.121], [3.935, 0.642, 2.507, 0.596],
   [0.643, 5.782, 2.507, 0.596], [4.553, 5.672, 1.951, 0.859],
  ].forEach(([x, y, w, h]) => note(s, x, y, w, h, LOREM_SHORT));
}

// Slide 23 - combined column + line chart
function slide23(s, n) {
  chrome(s, n);
  title(s, 0.805, 0.818, 4.695, 1.42, 'Chart \nData Report', 48);
  [{ x: 7.333, v: '720K', color: C.VIOLET, ty: 1.127 },
   { x: 9.77, v: '3,328', color: C.GOLD, ty: 1.062 },
  ].forEach(k => {
    s.addText(k.v, { x: k.x, y: k.ty, w: 2.309, h: 0.505, fontSize: 24, bold: true,
      color: k.color, fontFace: HEAD_FONT });
    body(s, k.x - 0.053, 1.567, 2.309, 0.679, 'Lorem ipsum dolor sit amet, consectetuer.');
  });

  s.addShape('roundRect', { x: 0.76, y: 2.571, w: 11.825, h: 3.954, rectRadius: 0.28,
    fill: { color: C.LILAC } });
  s.addText('Some project title', { x: 1.656, y: 2.993, w: 2.648, h: 0.353, fontSize: 15, fontFace: HEAD_FONT });
  art(s, 'calIcon', 9.77, 3.024, 0.289, 0.234, { fill: { color: C.BLACK } });
  s.addText('December 2024', { x: 10.113, y: 3.014, w: 1.74, h: 0.311, fontSize: 12.5, fontFace: HEAD_FONT });
  art(s, 'caret', 11.71, 3.076, 0.129, 0.157, { fill: { color: C.BLACK }, rotate: 270 });
  art(s, 'caret', 9.446, 3.076, 0.129, 0.157, { fill: { color: C.GREY_45 }, rotate: 90, flipH: true });

  const YEARS = [['2019', '2020', '2021', '2022', '2023', '2024']];
  s.addChart([
    { type: pptx.ChartType.bar, data: [
        { name: 'Series 1', labels: YEARS, values: [2.6, 5, 3.5, 4.5, 2, 4] },
        { name: 'Series 3', labels: YEARS, values: [1.8, 4, 3, 5, 4, 5] },
      ],
      options: { chartColors: [C.GOLD, C.VIOLET], barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 269 } },
    { type: pptx.ChartType.line, data: [
        { name: 'Series 2', labels: YEARS, values: [2.4, 4.7, 4, 2, 3.3, 3.8] },
      ],
      options: { chartColors: [C.MAGENTA], lineSize: 3, lineDataSymbol: 'circle', lineDataSymbolSize: 8 } },
  ], {
    x: 1.8, y: 3.505, w: 9.721, h: 2.375,
    showLegend: false, showTitle: false,
    catAxisLabelFontSize: 20, valAxisLabelFontSize: 20,
    valAxisMaxVal: 6, valAxisMinVal: 0, valAxisMajorUnit: 2,
    valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
    plotArea: { fill: { color: C.LILAC } }, chartArea: { fill: { color: C.LILAC } },
  });
}

// Slide 24 - North America map with two progress bars
function slide24(s, n) {
  chrome(s, n);
  // Base land masses, then the small island fragments that sit on top.
  art(s, 'mapCanada', 0.64, 1.013, 4.467, 3.446, { fill: { color: C.GREY_20 } });
  art(s, 'mapGreenland', 4.088, 0.973, 3.074, 2.295, { fill: { color: C.LEAF } });
  art(s, 'mapUsMex', -0.718, 2.313, 5.099, 3.471, { fill: { color: C.VIOLET } });
  [['mapBit1', 3.212, 5.999, 0.121, 0.066, C.GREY_20], ['mapBit2', 3.247, 5.908, 0.31, 0.169, C.VIOLET],
   ['mapAlaska', 1.854, 5.021, 1.532, 0.947, C.VIOLET], ['mapBit3', 3.088, 5.825, 0.214, 0.209, C.VIOLET],
   ['mapBit4', 3.255, 5.79, 0.066, 0.136, C.GREY_20], ['mapBit5', 3.993, 5.694, 0.149, 0.126, C.VIOLET],
   ['mapBit6', 4.129, 5.712, 0.174, 0.121, C.VIOLET], ['mapBit7', 3.343, 5.969, 0.214, 0.222, C.VIOLET],
   ['mapBit8', 3.416, 6.178, 0.166, 0.131, C.GREY_20], ['mapBit9', 3.57, 6.243, 0.292, 0.136, C.VIOLET],
   ['mapBit10', 3.481, 5.528, 0.529, 0.184, C.VIOLET], ['mapBit11', 3.784, 5.777, 0.126, 0.06, C.VIOLET],
   ['mapBit12', 4.351, 5.772, 0.113, 0.066, C.VIOLET], ['mapBit13', 3.801, 5.42, 0.043, 0.096, C.GREY_20],
  ].forEach(([name, x, y, w, h, color]) => art(s, name, x, y, w, h, { fill: { color } }));

  title(s, 7.326, 1.821, 4.467, 1.447, 'Creative Map Infographic', 48);
  [{ y: 4.208, label: 'Parameter 1', pct: '80%', px: 10.759, w: 4.25, color: C.VIOLET },
   { y: 4.935, label: 'Parameter 2', pct: '70%', px: 10.055, w: 3.545, color: C.LEAF },
  ].forEach(p => {
    s.addText(p.label, { x: 6.913, y: p.y - 0.443, w: 1.705, h: 0.417, fontSize: 14,
      fontFace: HEAD_FONT, lineSpacingMultiple: 1.5 });
    s.addText(p.pct, { x: p.px, y: p.y - 0.458, w: 0.987, h: 0.417, align: 'center', fontSize: 14,
      bold: true, fontFace: HEAD_FONT, lineSpacingMultiple: 1.5 });
    pill(s, 7.331, p.y, 4.917, 0.084, C.GREY_25);
    pill(s, 7.002, p.y, p.w, 0.084, p.color);
  });
  note(s, 6.913, 5.361, 5.742, 0.859,
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, ullamco \nconsectetur adipiscing elit. Integer vitae justo ullamcorper, sceler\nisque mi quis, ornare erat. Lorem ipsum dolor consectetur');
}

// Slide 25 - closing slide with a contact bar
function slide25(s) {
  s.background = { color: C.INDIGO };
  title(s, 0.769, 1.475, 4.53, 1.434, 'Get in Touch With Us!', 48, C.WHITE);
  s.addText(LOREM_LONG, { x: 0.76, y: 3.352, w: 4.165, h: 1.904,
    fontSize: 14, color: C.WHITE, lineSpacingMultiple: 1.3 });

  pill(s, 0.752, 6.102, 11.873, 0.6, C.GOLD);
  [{ icon: 'iconPin', ix: 1.368, iw: 0.183, ih: 0.229, tx: 1.651, tw: 2.451, text: '256 Newland Park, CA' },
   { icon: 'iconPhone', ix: 4.286, iw: 0.231, ih: 0.231, tx: 4.593, tw: 2.235, text: '+99 4453 1234' },
   { icon: 'iconGlobe', ix: 7.012, iw: 0.231, ih: 0.231, tx: 7.32, tw: 2.235, text: 'www.website.com' },
   { icon: 'iconMail', ix: 9.738, iw: 0.188, ih: 0.231, tx: 10.025, tw: 2.235, text: 'mail@company.com' },
  ].forEach(c => {
    art(s, c.icon, c.ix, 6.288, c.iw, c.ih, { fill: { color: C.WHITE } });
    s.addText(c.text, { x: c.tx, y: 6.234, w: c.tw, h: 0.337, fontSize: 14, color: C.WHITE });
  });
}

/* ==================================================================== build */

const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
  slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18,
  slide19, slide20, slide21, slide22, slide23, slide24, slide25,
];

BUILDERS.forEach((build, i) => build(pptx.addSlide(), i + 1));

pptx.writeFile({ fileName: path.join(__dirname, '154dd48c-4153-444f-b91c-e1861f280ed2_grok_final.pptx') })
  .then(f => console.log('wrote ' + f))
  .catch(err => { console.error(err); process.exit(1); });
