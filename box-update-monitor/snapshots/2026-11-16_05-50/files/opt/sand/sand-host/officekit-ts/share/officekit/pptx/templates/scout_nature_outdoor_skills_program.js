/**
 * Standalone pptxgenjs recreation of the "SCOUT. Community" (Nature Edition)
 * deck: 18 slides at 13.333 x 7.5 in.   Run:  node <thisfile>.js
 *
 * The source deck's picture frames are empty placeholders (it ships no embedded
 * raster media), so every visual here is rebuilt from native pptxgenjs shapes:
 * rounded-rectangle cards, ellipses, stars, a pie wedge, and custGeom freeforms
 * traced from the original artwork (see the normalised outlines below).
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

const OUT_NAME = '0973477c-16dc-4cb8-8a0c-6e295d6392b9_grok_final.pptx';

/* ------------------------------------------------------------------ palette */
const C = {
  green: '97BB23',   // accent1
  green2: '7A981C',  // accent2
  green3: '566A14',  // accent3
  olive: '3C4A0E',   // accent4 - darkest green
  yellow: 'FDDB4D',  // accent5
  gold: 'F5C603',
  white: 'FFFFFF',
  ink: '404040',     // body text (tx1 lum 75%)
  black: '000000',
  dark: '1A1A1A',
  grey: 'A6A6A6',
  tree: 'BFBFBF',    // silhouette trees in background
};

const FONT = 'Ubuntu';          // major (heading) theme font
const BODY_FONT = 'Roboto Light'; // minor (body) theme font

/*
 * Shadow presets lifted from the source deck.  These are *factories* because
 * pptxgenjs rewrites the shadow object in place while serialising a shape.
 *
 * The originals shrink the shadow to 70% and top-align it (`sx/sy` + `algn`),
 * which pptxgenjs cannot express, so the offsets below are reduced to land the
 * blur in the same place: offset = origDistance - 0.3 x shapeHeight(pt).
 */
const shadow = (color, opacity, blur, offset, angle) => () =>
  ({ type: 'outer', color, opacity, blur, offset, angle });
const SH_CARD = shadow(C.black, 0.07, 45, 35, 45);    // drifting photo cards
const SH_SOFT = shadow(C.black, 0.07, 34, 15, 45);    // small "+780" chips
const SH_GREEN = shadow(C.green, 0.2, 45, 25, 90);    // green tabs / pills
const SH_TAB = shadow(C.green, 0.2, 45, 29, 90);      // countdown tiles
const SH_OLIVE = shadow(C.olive, 0.2, 45, 44, 90);    // yellow skull badges
const SH_YELLOW = shadow(C.yellow, 0.2, 45, 44, 90);
const SH_BADGE = shadow(C.green, 0.68, 31, 5, 90);    // slide-number bubble
const SH_DOT = shadow(C.black, 0.07, 41, 0, 90);      // pricing "01/02/03" dots
const SH_DOTGREEN = shadow(C.green, 0.2, 45, 21, 90); // contact icon circles

// pptxgenjs always emits an <a:ln>; a fully transparent stroke keeps shapes clean.
const noLine = () => ({ color: C.white, transparency: 100 });

// PowerPoint's default text-box insets, in points: [left, right, bottom, top].
const INSETS = [7.2, 7.2, 3.6, 3.6];

/* ------------------------------------------------ vector artwork paths */
/* Normalised (0..1) outlines traced from the original freeform shapes. */
const TREE = [[[0.2756,0.1554],[0.3121,0.0687],[0.3522,0],[0.3714,0.003],[0.3437,0.0777],[0.3877,0.0752],[0.4287,0.0524],[0.4985,0.0362],[0.5441,0.0354],[0.6099,0.0444],[0.5972,0.0505],[0.5307,0.0428],[0.4925,0.0469],[0.4428,0.0636],[0.4311,0.0878],[0.3566,0.1014],[0.3412,0.108],[0.314,0.1645],[0.3179,0.1897],[0.332,0.1925],[0.4247,0.1806],[0.4354,0.1877],[0.3204,0.2109],[0.3241,0.2747],[0.3379,0.308],[0.3836,0.3147],[0.4217,0.3148],[0.4993,0.2926],[0.5227,0.2583],[0.5419,0.2563],[0.5334,0.2866],[0.5848,0.2805],[0.6441,0.2684],[0.7069,0.2344],[0.7442,0.2341],[0.8145,0.2108],[0.823,0.2199],[0.7414,0.2474],[0.6987,0.2536],[0.6825,0.2684],[0.7321,0.2676],[0.9842,0.2426],[0.9961,0.2426],[1,0.2507],[0.8124,0.2694],[0.8498,0.2886],[0.8461,0.2913],[0.7634,0.2805],[0.6527,0.2876],[0.5734,0.3017],[0.5164,0.3188],[0.5287,0.3296],[0.5632,0.338],[0.6588,0.3264],[0.6513,0.3374],[0.6288,0.3452],[0.542,0.3501],[0.4913,0.3302],[0.4738,0.3289],[0.3759,0.3622],[0.3615,0.4668],[0.3652,0.4884],[0.3794,0.4915],[0.5202,0.4863],[0.5577,0.4806],[0.5675,0.4732],[0.5886,0.4241],[0.6051,0.4214],[0.7187,0.4258],[0.8273,0.3703],[0.8433,0.3695],[0.8474,0.3739],[0.7655,0.4369],[0.6379,0.445],[0.6143,0.4918],[0.4226,0.5392],[0.3886,0.552],[0.4077,0.6942],[0.412,0.8305],[0.4272,0.827],[0.4823,0.7952],[0.4967,0.806],[0.4184,0.8618],[0.4227,0.9596],[0.1927,1],[0.2152,0.8002],[0.2354,0.7078],[0.2332,0.7013],[0.2084,0.6903],[0.1033,0.6609],[0.1182,0.6518],[0.2204,0.669],[0.2275,0.6616],[0.2261,0.6411],[0.2161,0.6286],[0.1639,0.6239],[0.1032,0.56],[0.0074,0.5378],[0.0221,0.5306],[0.0981,0.5394],[0.0649,0.4864],[0.0622,0.4625],[0.0348,0.4467],[0.0524,0.4245],[0.0738,0.4119],[0.0812,0.4194],[0.0567,0.4443],[0.0954,0.4531],[0.1054,0.4955],[0.1586,0.5025],[0.1503,0.5074],[0.116,0.5086],[0.1152,0.5133],[0.1868,0.5909],[0.1969,0.5943],[0.2188,0.5922],[0.2441,0.3538],[0.2248,0.3492],[0.14,0.3467],[0.1295,0.3082],[0.0616,0.299],[0.022,0.2621],[0.0321,0.2564],[0.0499,0.2553],[0.0883,0.2845],[0.1841,0.2886],[0.1756,0.3269],[0.2193,0.3279],[0.2565,0.3224],[0.2662,0.3142],[0.2653,0.2401],[0.2471,0.1918],[0.2069,0.1568],[0.1117,0.1383],[0.0182,0.1299],[0.0017,0.1249],[0.0102,0.1157],[0.0719,0.1199],[0.0819,0.117],[0.0584,0.092],[0.084,0.0564],[0.062,0.0263],[0.0689,0.0191],[0.0852,0.0183],[0.0958,0.0238],[0.1117,0.0554],[0.1588,0.047],[0.1663,0.0498],[0.1593,0.0576],[0.1179,0.0664],[0.1075,0.0742],[0.1066,0.1003],[0.1186,0.115],[0.1531,0.1282],[0.1723,0.1273],[0.2211,0.0896],[0.2322,0.087],[0.2426,0.0965],[0.2099,0.1346],[0.2429,0.1493],[0.2757,0.1553]]];
const TREE_ALT = [[[0.2647,0],[0.2076,0.0339],[0.1898,0.074],[0.1905,0.1094],[0.1728,0.1279],[0.1227,0.0515],[0.1116,0.0559],[0.2048,0.224],[0.1602,0.2277],[0.143,0.1818],[0.1304,0.1773],[0.1371,0.2351],[0.1027,0.2441],[0.0635,0.212],[0.0696,0.1687],[0.0498,0.1939],[0.0161,0.1718],[0.001,0.1752],[0.0382,0.205],[0.0564,0.246],[0.0878,0.2594],[0.054,0.292],[0.1773,0.2484],[0.2613,0.2897],[0.3356,0.3658],[0.3312,0.3889],[0.3055,0.4106],[0.2773,0.3735],[0.2308,0.4053],[0.1891,0.3936],[0.1822,0.41],[0.2309,0.428],[0.2691,0.411],[0.306,0.4412],[0.3588,0.409],[0.5468,0.6294],[0.5313,0.6401],[0.4342,0.5979],[0.4447,0.5723],[0.4149,0.5857],[0.375,0.5519],[0.3502,0.5584],[0.3308,0.5234],[0.3424,0.5686],[0.3679,0.5725],[0.4484,0.6273],[0.4068,0.6476],[0.4062,0.6594],[0.4679,0.6437],[0.5486,0.6777],[0.576,0.6627],[0.5909,0.67],[0.6117,0.6968],[0.5514,0.7194],[0.5523,0.7329],[0.6509,0.7256],[0.876,1],[0.9457,0.8793],[0.8592,0.7887],[0.8489,0.7139],[0.8636,0.694],[0.8405,0.7082],[0.8341,0.5596],[0.8954,0.5274],[0.9018,0.5025],[0.9364,0.4982],[0.9585,0.4635],[0.9237,0.4922],[0.8884,0.4913],[0.88,0.5215],[0.8349,0.5274],[0.8295,0.4623],[0.8573,0.4405],[0.8758,0.4518],[0.9018,0.4481],[0.9066,0.4387],[0.8695,0.4377],[0.8904,0.4232],[0.9583,0.4204],[0.9491,0.4087],[0.9994,0.395],[0.9139,0.4082],[0.9497,0.3775],[0.8972,0.4111],[0.8734,0.4179],[0.8723,0.4022],[0.8569,0.4261],[0.8227,0.4304],[0.8157,0.3775],[0.8469,0.3651],[0.815,0.3662],[0.8213,0.3226],[0.8623,0.2899],[0.8942,0.2886],[0.864,0.2842],[0.822,0.3064],[0.8243,0.2649],[0.8035,0.3479],[0.7857,0.3368],[0.7917,0.3114],[0.7755,0.3329],[0.7609,0.3263],[0.7579,0.3045],[0.7739,0.2915],[0.759,0.2945],[0.7465,0.276],[0.751,0.3274],[0.7292,0.3316],[0.7849,0.3486],[0.8001,0.3861],[0.7921,0.4392],[0.7764,0.4394],[0.7787,0.419],[0.7347,0.4048],[0.7455,0.4245],[0.7639,0.4294],[0.7667,0.45],[0.795,0.4538],[0.7881,0.581],[0.76,0.5389],[0.7718,0.5332],[0.7573,0.5294],[0.7441,0.5021],[0.7488,0.4848],[0.7382,0.5034],[0.7554,0.5529],[0.7307,0.552],[0.7567,0.5639],[0.7732,0.598],[0.7893,0.6039],[0.7885,0.622],[0.7568,0.6177],[0.792,0.6393],[0.7859,0.7243],[0.5895,0.531],[0.6407,0.3944],[0.6122,0.3442],[0.6628,0.2899],[0.647,0.2039],[0.6349,0.2081],[0.6325,0.2974],[0.5725,0.3439],[0.6058,0.4054],[0.5348,0.4809],[0.5061,0.4657],[0.4251,0.3678],[0.4413,0.3021],[0.4897,0.2957],[0.5246,0.2592],[0.5225,0.2316],[0.4891,0.2771],[0.452,0.2775],[0.4872,0.1994],[0.5311,0.1523],[0.5773,0.1313],[0.5438,0.1243],[0.6138,0.0399],[0.6025,0.0333],[0.4846,0.1714],[0.4795,0.1524],[0.5072,0.0766],[0.4958,0.0717],[0.4836,0.1182],[0.4671,0.1323],[0.4673,0.1856],[0.4508,0.2183],[0.4327,0.2426],[0.4112,0.2127],[0.4043,0.2216],[0.4224,0.2606],[0.406,0.3088],[0.3587,0.3336],[0.2736,0.2544],[0.3059,0.1913],[0.2952,0.189],[0.2548,0.2366],[0.197,0.1556],[0.2205,0.1045],[0.2055,0.0787],[0.2139,0.0456],[0.264,0.0101]]];
const CAL_BODY = [[[0,0.6429],[0.0611,0.8953],[0.2083,1],[0.8336,0.9927],[0.9643,0.8424],[1,0.6429],[1,0],[0,0]], [[0.7083,0.3214],[0.7708,0.4286],[0.7083,0.5357],[0.6458,0.4286],[0.7083,0.3214]], [[0.5,0.3214],[0.5625,0.4286],[0.5,0.5357],[0.4375,0.4286],[0.5,0.3214]], [[0.2917,0.3214],[0.3542,0.4286],[0.2917,0.5357],[0.2292,0.4286],[0.2917,0.3214]]];
const CAL_TOP = [[[0.7917,0.25],[0.75,0.25],[0.7083,0],[0.6667,0.25],[0.3333,0.25],[0.3262,0.0551],[0.2754,0.0098],[0.25,0.25],[0.0611,0.4333],[0,1],[1,1],[0.9389,0.4333],[0.7917,0.25]]];
const ACORN_NUT = [[[0.3098,0],[0.1001,0.2202],[0.0228,0.3825],[0.0003,0.5593],[0.0777,0.8082],[0.0628,0.8731],[0.1061,0.921],[0.1754,0.9101],[0.3381,0.9858],[0.6095,0.9793],[0.7768,0.9012],[1,0.6898],[0.6901,0.3085],[0.3098,0]]];
const ACORN_CAP = [[[0.9861,0.6938],[0.8912,0.471],[0.8793,0.3085],[0.8452,0.2561],[0.9703,0.1909],[0.9703,0.0723],[0.7606,0.1723],[0.7006,0.123],[0.5297,0.1091],[0.2175,0.0003],[0.0947,0.0431],[0.0343,0.1078],[0,0.2261],[0.2022,0.3572],[0.4349,0.5651],[0.6427,0.798],[0.7738,1],[0.8624,0.9814],[0.9621,0.8981],[0.9957,0.8163],[0.9861,0.6938]]];
const SHIELD = [[[0.8416,0.0609],[0.8416,0],[1,0],[1,0.5609],[0.9898,0.6494],[0.9146,0.8064],[0.7796,0.925],[0.6008,0.9911],[0.4034,0.9918],[0.2306,0.9309],[0.0971,0.821],[0.0169,0.6744],[0,0],[0.1584,0],[0.1584,0.0609],[0.4257,0.0609],[0.4257,0],[0.5842,0],[0.5842,0.0609],[0.8416,0.0609]], [[0.8416,0.5609],[0.8416,0.4755],[0.7044,0.5705],[0.5534,0.6315],[0.3876,0.6642],[0.1815,0.6696],[0.2318,0.7467],[0.3486,0.8299],[0.4467,0.8572],[0.5688,0.8548],[0.691,0.8096],[0.7832,0.7286],[0.8416,0.5609]], [[0.4408,0.4561],[0.3396,0.5286],[0.503,0.4996],[0.6649,0.4267],[0.7795,0.3224],[0.8362,0.2],[0.5791,0.2],[0.5373,0.3285],[0.4408,0.4561]], [[0.1584,0.4454],[0.1584,0.2],[0.4185,0.2],[0.3255,0.3606],[0.1584,0.4454]]];
const SKULL = [[[0.534,0.001],[0.3807,0.0131],[0.2408,0.0664],[0.1229,0.1577],[0.0236,0.3198],[0,0.4583],[0.0028,0.7335],[0.0233,0.7782],[0.0833,0.8235],[0.2273,0.8333],[0.2273,1],[0.7727,1],[0.7727,0.8333],[0.9167,0.8235],[0.9767,0.7782],[0.9972,0.7335],[0.9907,0.382],[0.9215,0.2186],[0.7963,0.0918],[0.7171,0.0465],[0.534,0.001]], [[0.3182,0.5417],[0.2539,0.5173],[0.2273,0.4583],[0.2539,0.3994],[0.3182,0.375],[0.3825,0.3994],[0.4091,0.4583],[0.3825,0.5173],[0.3182,0.5417]], [[0.5,0.7499],[0.4246,0.7416],[0.4091,0.7082],[0.4357,0.6284],[0.5,0.5832],[0.5643,0.6284],[0.5909,0.7082],[0.5754,0.7416],[0.5,0.7499]], [[0.6818,0.5417],[0.6175,0.5173],[0.5909,0.4583],[0.6175,0.3994],[0.6818,0.375],[0.7461,0.3994],[0.7727,0.4583],[0.7461,0.5173],[0.6818,0.5417]]];
const QUOTE = [[[0.9002,0],[1,0.2181],[0.8312,0.3499],[0.7792,0.5049],[1,0.5049],[1,1],[0.5382,1],[0.5842,0.3031],[0.6977,0.1376],[0.9002,0]], [[0.362,0],[0.4618,0.2181],[0.293,0.3499],[0.241,0.5049],[0.4618,0.5049],[0.4618,1],[0,1],[0.046,0.3031],[0.1595,0.1376],[0.362,0]]];
const ENVELOPE = [[[0.9981,0],[0.6473,0.4822],[0.5784,0.545],[0.5,0.5659],[0.4216,0.545],[0.3527,0.4822],[0.0019,0],[0,0.7136],[0.0164,0.825],[0.0919,0.951],[0.1664,0.9941],[0.7917,1],[0.9081,0.951],[0.9643,0.8736],[0.9957,0.7713],[0.9981,0]]];
const ENV_FLAP = [[[0.5943,0.9295],[1,0.1941],[0.9195,0.0521],[0.811,0],[0.1603,0.0035],[0.057,0.08],[0,0.1941],[0.4267,0.9603],[0.5255,0.9956],[0.5943,0.9295]]];
const PHONE = [[[0.9583,0.4583],[0.9198,0.4329],[0.8902,0.287],[0.8187,0.1811],[0.7126,0.1097],[0.5594,0.0763],[0.5532,0.0123],[0.7451,0.0329],[0.8777,0.1222],[0.9915,0.3327],[0.9992,0.425],[0.9583,0.4583]], [[0.8331,0.4167],[0.7598,0.2399],[0.6332,0.1718],[0.5532,0.1789],[0.5532,0.2378],[0.7211,0.3235],[0.7619,0.4461],[0.8076,0.455],[0.8331,0.4167]], [[0.9242,0.9234],[0.9999,0.7885],[0.9457,0.6845],[0.7711,0.5833],[0.6023,0.6861],[0.4674,0.6053],[0.3626,0.4897],[0.313,0.3968],[0.3955,0.2977],[0.4136,0.2046],[0.3015,0.0379],[0.1873,0.0021],[0.073,0.0776],[0.0063,0.212],[0.0098,0.3618],[0.0698,0.5166],[0.3047,0.7994],[0.6019,0.9768],[0.7909,0.995],[0.9242,0.9234]]];
const GLOBE = [[[0.9755,0.6704],[0.5989,0.5015],[0.5023,0.5651],[0.7016,0.9995],[0.8227,0.8227],[0.9906,0.7346],[0.9756,0.6704]], [[0.538,0.9149],[0.3653,0.7083],[0.4253,0.6961],[0.4303,0.6433],[0.3321,0.625],[0.3321,0.375],[0.6678,0.375],[0.7204,0.4583],[0.7593,0.4331],[0.7543,0.375],[0.8973,0.375],[0.9243,0.5684],[0.983,0.5737],[0.9998,0.5],[0.9605,0.3056],[0.8533,0.1466],[0.6944,0.0394],[0.5,0],[0.3055,0.0394],[0.1466,0.1466],[0.0393,0.3055],[0,0.5],[0.0393,0.6944],[0.2206,0.9145],[0.4999,1],[0.578,0.9766],[0.5379,0.9149]], [[0.2455,0.625],[0.1025,0.625],[0.1025,0.3749],[0.2455,0.3749],[0.2455,0.6249]], [[0.3655,0.2917],[0.5,0.1005],[0.6346,0.2917],[0.3655,0.2917]], [[0.8607,0.2917],[0.7262,0.2917],[0.6082,0.0976],[0.7542,0.1702],[0.8607,0.2917]], [[0.3916,0.0977],[0.2736,0.2917],[0.1392,0.2917],[0.2456,0.1702],[0.3916,0.0977]], [[0.1392,0.7083],[0.2736,0.7083],[0.3916,0.9023],[0.2456,0.8297],[0.1392,0.7083]]];
const BANNER = [[[0.6118,0.1992],[0.6016,0.1213],[0.5774,0.058],[0.5424,0.0155],[0.5,0],[0.4576,0.0155],[0.4226,0.058],[0.3984,0.1213],[0.3882,0.1992],[0.0483,0.1992],[0.0213,0.2146],[0.0038,0.2543],[0,0.2894],[0,1],[0.3824,1],[0.3949,0.9227],[0.4208,0.8603],[0.4569,0.8185],[0.5,0.8033],[0.5431,0.8185],[0.5792,0.8603],[0.6051,0.9227],[0.6176,1],[1,1],[1,0.2894],[0.9962,0.2543],[0.9859,0.2256],[0.9705,0.2063],[0.9517,0.1992],[0.6118,0.1992]]];

/* ---------------------------------------------------------------- helpers */

// Convert a normalised multi-subpath outline into pptxgenjs custGeom points.
function geomPoints(subpaths, w, h) {
  const pts = [];
  subpaths.forEach(sub => {
    sub.forEach(([px, py], i) => {
      pts.push(i === 0 ? { x: px * w, y: py * h, moveTo: true } : { x: px * w, y: py * h });
    });
    pts.push({ close: true });
  });
  return pts;
}

// Freeform shape built from one of the normalised outlines above.
function freeform(slide, subpaths, o) {
  slide.addShape('custGeom', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    points: geomPoints(subpaths, o.w, o.h),
    fill: { color: o.color, transparency: o.transparency || 0 },
    line: noLine(),
    rotate: o.rotate, flipH: o.flipH, flipV: o.flipV,
  });
}

// Background tree silhouette - the deck's recurring decorative motif.
function tree(slide, x, y, w, h, alpha, opts = {}) {
  freeform(slide, opts.alt ? TREE_ALT : TREE, {
    x, y, w, h, color: C.tree, transparency: 100 - alpha,
    rotate: opts.rotate, flipH: opts.flipH,
  });
}

// Rounded rectangle "card".  `adj` is the source deck's corner-radius adjust
// value (fraction of the shorter side x 100000), mapped to `rectRadius` inches.
function card(slide, o) {
  slide.addShape('roundRect', {
    x: o.x, y: o.y, w: o.w, h: o.h,
    rectRadius: (o.adj / 100000) * Math.min(o.w, o.h),
    fill: { color: o.fill },
    line: noLine(),
    rotate: o.rotate,
    shadow: o.shadow === undefined ? SH_CARD() : (o.shadow && o.shadow()),
  });
}

// Plain text box.  `runs` is a string or an array of {text, options} runs.
function text(slide, runs, o) {
  slide.addText(runs, Object.assign({
    fontFace: FONT, fontSize: 12, color: C.ink,
    align: 'left', valign: 'top', margin: INSETS, wrap: true,
  }, o));
}

// Body paragraph: 12pt Roboto Light on 150% leading - the deck's running text.
function body(slide, str, o) {
  text(slide, str, Object.assign({
    fontFace: BODY_FONT, fontSize: 12, color: C.ink, lineSpacingMultiple: 1.5,
  }, o));
}

// Heading block: small green eyebrow line + large multi-colour headline.
// Each headline run is [text, colour?, breakLineAfter?].
function heading(slide, o) {
  text(slide, o.eyebrow, {
    x: o.ex === undefined ? o.x : o.ex,
    y: o.ey === undefined ? o.y - 0.46 : o.ey,
    w: o.ew || 2.865, h: 0.337,
    fontFace: BODY_FONT, fontSize: 14, color: C.green, rotate: o.rotate, align: o.align,
  });
  text(slide, o.runs.map(r => ({
    text: r[0],
    options: { bold: true, fontSize: o.size || 44, color: r[1] || C.ink, fontFace: FONT, breakLine: r[2] },
  })), { x: o.x, y: o.y, w: o.w, h: o.h || 1.582, rotate: o.rotate, align: o.align });
}

// Small tab holding a rotated label - used all over the deck.
function tab(slide, o) {
  card(slide, {
    x: o.x, y: o.y, w: o.w || 0.684, h: o.h || 1.087, adj: 13558,
    fill: o.fill || C.green, rotate: o.rotate === undefined ? 90 : o.rotate,
    shadow: o.shadow || SH_GREEN,
  });
  text(slide, o.label, {
    x: o.lx, y: o.ly, w: o.lw || 1.017, h: o.lh || 0.505, rotate: o.lrot,
    fontSize: o.size || 12, bold: true, color: C.white,
    align: 'center', valign: 'middle',
  });
}

// Pill "Join Now" button: a 90-degree rotated rounded bar + centred caption.
function joinButton(slide, x, y) {
  card(slide, {
    x: x + 0.332, y: y - 0.428, w: 0.352, h: 1.159, adj: 50000,
    fill: C.green, rotate: 90, shadow: SH_GREEN,
  });
  text(slide, 'Join Now', {
    x, y, w: 1.017, h: 0.303, fontSize: 12, bold: true, color: C.white,
    align: 'center', valign: 'middle',
  });
}

// Calendar glyph (top bar + body) drawn next to dates and percentages.
function calendarIcon(slide, x, y, w, color) {
  freeform(slide, CAL_TOP, { x, y, w, h: w * 0.333, color });
  freeform(slide, CAL_BODY, { x, y: y + w * 0.417, w, h: w * 0.585, color });
}

// Yellow circle carrying a white skull glyph (the deck's "badge").
function skullBadge(slide, x, y, d, ringColor = C.gold, sh = SH_OLIVE) {
  slide.addShape('ellipse', {
    x, y, w: d, h: d, fill: { color: ringColor }, line: noLine(), shadow: sh(),
  });
  const s = d * 0.418;
  freeform(slide, SKULL, { x: x + (d - s) / 2, y: y + d * 0.3, w: s, h: s * 1.09, color: C.white });
}

/* Reused copy from the source deck. */
const LOREM_FULL = 'The Nature Edition of the Scout Community focuses on immersive outdoor exploration where young scouts learn essential survival and character-building skills. Through direct interaction with forests, rivers, and wildlife, scouts cultivate teamwork, responsibility, and a deeper respect for the natural environment.';
const LOREM_MID = 'The Nature Edition of the Scout Community focuses on immersive outdoor exploration where young scouts learn essential survival and character-building skills. Through direct interaction with forests, rivers, and wildlife.';
const LOREM_SKILLS = 'The Nature Edition of the Scout Community focuses on immersive outdoor exploration where young scouts learn essential survival and character-building skills. ';
const LOREM_SHORT = 'The Nature Edition of the Scout Community focuses on immersive outdoor exploration where young scouts learn essential survival.';
const LOREM_TINY = 'The Nature Edition of the Scout Community focuses on immersive outdoor exploration.';
const LOREM_CARD = 'The Nature Edition of the Scout Community.';

/* ------------------------------------------------------------ slide chrome */

// Master decoration: green circle badge with the slide number (top-right).
function slideBadge(slide, n) {
  slide.addShape('ellipse', {
    x: 12.767, y: 0.26, w: 0.331, h: 0.331,
    fill: { color: C.green }, line: noLine(),
    shadow: SH_BADGE(),
  });
  text(slide, String(n), {
    x: 12.4, y: 0.209, w: 1.058, h: 0.399,
    fontSize: 10.5, bold: true, color: C.white, align: 'center', valign: 'middle',
  });
}

/* -------------------------------------------------------------- the slides */

// 1 - Cover: oversized "SCOUT. / Community" wordmark, acorn dot, logo lockup.
function slide01(pptx) {
  const s = pptx.addSlide();
  // the source box does not wrap, so PowerPoint centres the word inside it
  text(s, 'SCOUT.', { x: 1.226, y: 1.957, w: 5.985, h: 2.036, fontSize: 115, bold: true, color: C.green, wrap: false, align: 'center' });
  text(s, 'Community', { x: 1.226, y: 3.509, w: 10.108, h: 2.036, fontSize: 115, bold: true, color: C.ink });
  body(s, 'A special scouting journey that teaches valuable life skills through direct experiences\nin the natural world.',
    { x: 1.358, y: 5.456, w: 7.219, h: 0.673, color: C.black });

  // acorn punctuating the wordmark
  freeform(s, ACORN_CAP, { x: 7.293, y: 2.404, w: 0.356, h: 0.356, color: C.yellow });
  freeform(s, ACORN_NUT, { x: 7.19, y: 2.518, w: 0.346, h: 0.345, color: C.yellow });

  // "Logoipsum University" lockup: shield mark + serif wordmark
  freeform(s, SHIELD, { x: 1.359, y: 1.392, w: 0.329, h: 0.375, color: C.olive });
  text(s, 'Logoipsum\nUniversity', {
    x: 1.744, y: 1.36, w: 1.8, h: 0.46, fontFace: 'Times New Roman',
    fontSize: 14, bold: true, color: C.olive, lineSpacingMultiple: 0.95, margin: 0,
  });
}

// 2 - Intro: dark portrait card, headline, CTA row.
function slide02(pptx) {
  const s = pptx.addSlide();
  tree(s, -2.286, 0.604, 3.526, 7.441, 30, { flipH: true });
  card(s, { x: 0.946, y: 0.975, w: 3.86, h: 5.514, adj: 7517, fill: C.olive, rotate: 7.3, shadow: null });

  heading(s, {
    x: 5.773, y: 2.134, ey: 1.674, ew: 4.66, w: 6.772,
    eyebrow: 'Nature as a Teacher',
    runs: [['Outdoor'], [' ', C.olive], ['Experience', C.green2], [' ', C.olive], ['Builds'],
      [' ', C.olive], ['Inner'], [' ', C.olive], ['Strength']],
  });
  body(s, LOREM_FULL, { x: 5.773, y: 3.839, w: 6.505, h: 1.279, color: C.olive });

  joinButton(s, 5.926, 5.636);
  text(s, [
    { text: 'December', options: { color: C.green } },
    { text: ' 31, 2026', options: { color: C.black } },
  ], { x: 7.364, y: 5.585, w: 3.559, h: 0.404, fontSize: 18, bold: true });

  tree(s, 10.849, 3.578, 1.959, 4.133, 15, { flipH: true });
  s.addShape('wedgeEllipseCallout', {
    x: 4.204, y: 4.772, w: 1.099, h: 1.09, fill: { color: C.gold }, line: noLine(), flipH: true,
  });
  text(s, 'Outdoor Adventure', {
    x: 4.171, y: 5.024, w: 1.142, h: 0.505, rotate: 347.78,
    fontSize: 12, bold: true, color: C.white, align: 'center', valign: 'middle',
  });
  slideBadge(s, 2);
}

// 3 - Agenda: 30 / 15 / 36 countdown tiles.
function slide03(pptx) {
  const s = pptx.addSlide();
  const tiles = [
    { x: 3.655, tx: 2.933, value: '30', label: 'Days', fill: C.green },
    { x: 5.948, tx: 5.225, value: '15', label: 'Hours', fill: C.green3 },
    { x: 8.24, tx: 7.518, value: '36', label: 'Minutes', fill: C.green },
  ];
  tiles.forEach(t => {
    card(s, { x: t.x, y: 2.346, w: 1.438, h: 1.71, adj: 13558, fill: t.fill, rotate: 90, shadow: SH_TAB });
    text(s, t.value, { x: t.tx, y: 2.703, w: 2.883, h: 0.774, fontSize: 40, bold: true, color: C.white, align: 'center', valign: 'middle' });
    text(s, t.label, { x: t.tx, y: 3.293, w: 2.883, h: 0.337, fontFace: BODY_FONT, fontSize: 14, color: C.white, align: 'center', valign: 'middle' });
  });

  calendarIcon(s, 6.545, 1.286, 0.243, C.green);
  text(s, [
    { text: 'December ', options: { color: C.green } },
    { text: '31, 2026', options: { color: C.ink } },
  ], { x: 4.45, y: 1.79, w: 4.433, h: 0.438, fontSize: 20, bold: true, align: 'center' });

  text(s, 'Agenda', { x: 2.539, y: 4.509, w: 8.256, h: 0.841, fontSize: 44, bold: true, color: C.ink, align: 'center' });
  body(s, LOREM_FULL, { x: 2.194, y: 5.521, w: 8.945, h: 0.976, color: C.olive, align: 'center' });

  tree(s, -1.157, 0.825, 3.278, 6.916, 30, { rotate: 22.21, flipH: true });
  tree(s, 11.171, 2.385, 2.429, 5.125, 15, { rotate: 340.83, flipH: true });
  slideBadge(s, 3);
}

// 4 - "Nature as a Teacher": two tilted photo cards behind a big bare tree.
function slide04(pptx) {
  const s = pptx.addSlide();
  tree(s, 5.719, 0.913, 6.752, 7.24, 30, { rotate: 341.29, flipH: true, alt: true });
  card(s, { x: 9.399, y: 2.304, w: 2.662, h: 3.749, adj: 5269, fill: C.olive, rotate: 15, shadow: null });

  heading(s, {
    x: 0.969, y: 2.556, w: 5.238, ew: 4.377,
    eyebrow: 'Outdoor Experience Builds Inner Strength',
    runs: [['Nature as a '], ['Teacher', C.green]],
  });
  body(s, LOREM_SKILLS, { x: 0.969, y: 4.448, w: 4.377, h: 0.976, color: C.olive });

  card(s, { x: 6.756, y: 1.227, w: 2.662, h: 3.749, adj: 5269, fill: C.white, rotate: 345 });
  skullBadge(s, 9.311, 4.017, 0.758);
  tab(s, { x: 7.274, y: 0.806, rotate: 75, label: 'Nature Skills', lx: 7.11, ly: 1.107, lrot: 345 });
  slideBadge(s, 4);
}

// 5 - "Scouting with Respect": stacked cards + 25/50/75% strip.
function slide05(pptx) {
  const s = pptx.addSlide();
  card(s, { x: 2.882, y: 0.879, w: 2.662, h: 3.749, adj: 5269, fill: C.olive, rotate: 15, shadow: null });
  tree(s, -1.706, 0.04, 3.278, 6.916, 30, { rotate: 22.21, flipH: true });

  heading(s, {
    x: 6.781, y: 1.285, w: 5.976, h: 2.322, ew: 2.865,
    eyebrow: 'Outdoor Ethics',
    runs: [['Scouting'], [' ', C.black], ['with'], [' ', C.black], ['Respect', C.green],
      [' ', C.black], ['and'], [' ', C.black], ['Responsibility', C.green]],
  });
  body(s, LOREM_MID, { x: 6.775, y: 3.783, w: 5.976, h: 0.976, color: C.olive });

  card(s, { x: 1.098, y: 2.872, w: 2.662, h: 3.749, adj: 5269, fill: C.white, rotate: 345 });
  tab(s, { x: 2.321, y: 1.41, label: 'Nature Skills', lx: 2.154, ly: 1.712 });
  tree(s, 11.654, 1.189, 1.892, 3.991, 15, { rotate: 324.19, flipH: true });

  card(s, { x: 2.943, y: 5.356, w: 9.944, h: 1.198, adj: 6081, fill: C.white });
  [
    { icon: 3.372, px: 3.895, lx: 5.052, pct: '25%', color: C.green },
    { icon: 6.656, px: 7.179, lx: 8.337, pct: '50%', color: C.green2 },
    { icon: 9.97, px: 10.493, lx: 11.651, pct: '75%', color: C.green3 },
  ].forEach(g => {
    calendarIcon(s, g.icon, 5.755, 0.4, g.color);
    text(s, g.pct, { x: g.px, y: 5.635, w: 1.858, h: 0.64, fontSize: 32, bold: true, color: g.color });
    text(s, 'Your Text Here', { x: g.lx, y: 5.736, w: 1.074, h: 0.438, fontSize: 10, bold: true, color: C.grey });
  });
  slideBadge(s, 5);
}

// 6 - Schedule: five floating time cards over a tree.
function slide06(pptx) {
  const s = pptx.addSlide();
  tree(s, 7.464, 1.615, 3.278, 6.916, 30, { flipH: true });

  const slots = [
    { x: 6.924, y: 2.511, rot: 356.57, time: '7:00 PM', tx: 6.945, ty: 2.72, name: 'Group Bonding', nx: 6.966, ny: 3.034 },
    { x: 10.357, y: 2.801, rot: 4.5, time: '9:00 PM', tx: 10.402, ty: 3.011, name: 'Exploration', nx: 10.375, ny: 3.324 },
    { x: 7.147, y: 4.155, rot: 0, time: '4:00 PM', tx: 7.178, ty: 4.364, name: 'Team Hike', nx: 7.178, ny: 4.679 },
    { x: 10.389, y: 4.316, rot: 355.32, time: '7:00 PM', tx: 10.406, ty: 4.526, name: 'Campfire ', nx: 10.435, ny: 4.839 },
    { x: 8.527, y: 1.212, rot: 5.48, w: 2.659, tw: 2.573, dark: true, time: '2:00 Am', tx: 8.586, ty: 1.422, name: 'Scout Camp', nx: 8.553, ny: 1.735 },
  ];
  slots.forEach(t => {
    card(s, { x: t.x, y: t.y, w: t.w || 1.956, h: 1.137, adj: 37227, fill: t.dark ? C.green : C.white, rotate: t.rot });
    text(s, t.time, {
      x: t.tx, y: t.ty, w: t.tw || 1.893, h: 0.37, rotate: t.rot,
      fontSize: 16, bold: true, color: t.dark ? C.white : C.green, align: 'center',
    });
    text(s, t.name, {
      x: t.nx, y: t.ny, w: t.tw || 1.893, h: 0.449, rotate: t.rot,
      fontSize: 16, bold: true, color: t.dark ? C.white : C.ink, align: 'center', lineSpacingMultiple: 1.5,
    });
  });

  heading(s, {
    x: 1.167, y: 1.354, w: 5.976, ew: 2.865,
    eyebrow: 'Water Collection & Purification',
    runs: [['A '], ['Vital', C.green], [' Wilderness '], ['Lesson', C.green]],
  });
  text(s, 'Environmental Education Through Scouting', { x: 1.167, y: 3.349, w: 4.111, h: 0.707, fontSize: 18, bold: true, color: C.ink });
  body(s, LOREM_TINY, { x: 1.167, y: 4.221, w: 4.66, h: 0.673 });

  [['December 31, 2026', 5.292, 5.39], ['January 01, 2027', 5.928, 6.025]].forEach(([label, ty, iy]) => {
    calendarIcon(s, 1.267, iy, 0.243, C.green);
    text(s, label, { x: 1.828, y: ty, w: 3.054, h: 0.438, fontSize: 20, bold: true, color: C.ink });
  });
  slideBadge(s, 6);
}

// 7 - "Strong Connections": stat pair over a dark portrait card.
function slide07(pptx) {
  const s = pptx.addSlide();
  tree(s, 3.209, 3.412, 3.278, 6.916, 15, { rotate: 19.16 });
  card(s, { x: 0.926, y: 2.636, w: 3.191, h: 4.494, adj: 5269, fill: C.olive, rotate: 354.75, shadow: null });

  heading(s, {
    x: 7.799, y: 2.415, w: 4.487, h: 2.322, ew: 2.865,
    eyebrow: ' Knot Skills',
    runs: [['Strong ', C.black], ['Connections', C.green], [' for Survival', C.black]],
  });
  body(s, 'The Nature Edition of the Scout Community focuses on immersive outdoor exploration where young scouts learn..',
    { x: 7.799, y: 4.899, w: 3.707, h: 0.976, color: C.olive });

  card(s, { x: 3.573, y: -0.655, w: 3.191, h: 4.494, adj: 5269, fill: C.white, rotate: 11.2 });
  text(s, '12M', { x: 0.988, y: 1.432, w: 1.858, h: 0.841, rotate: 355.3, fontSize: 44, bold: true, color: C.green, align: 'center' });
  text(s, 'Learning Path', { x: 0.857, y: 2.123, w: 2.231, h: 0.337, rotate: 355.3, fontSize: 14, bold: true, color: C.green3, align: 'center' });

  tab(s, { x: 3.829, y: 5.234, label: '+543', lx: 3.662, ly: 5.535, size: 24 });
  text(s, '70%', { x: 5.198, y: 5.381, w: 1.858, h: 0.572, fontSize: 28, bold: true, color: C.green });
  text(s, 'Wild Navigation', { x: 5.198, y: 5.871, w: 2.231, h: 0.303, fontSize: 12, bold: true, color: C.green3 });

  tree(s, 11.174, -0.45, 3.278, 6.916, 15, { rotate: 340.84, flipH: true });
  skullBadge(s, 2.904, 2.867, 0.869);
  slideBadge(s, 7);
}

// 8 - "Wilderness First Aid": black portrait card, small white card behind.
function slide08(pptx) {
  const s = pptx.addSlide();
  card(s, { x: 6.749, y: 1.378, w: 3.421, h: 4.818, adj: 5269, fill: C.black, rotate: 354.75, shadow: null });
  tree(s, 10.296, 0.37, 3.526, 7.439, 30, { flipH: true });

  heading(s, {
    x: 1.18, y: 2.43, w: 4.451, ew: 2.865,
    eyebrow: 'Nature Observation',
    runs: [['Wilderness', C.green], [' First Aid']],
  });
  body(s, LOREM_MID, { x: 1.18, y: 4.335, w: 4.123, h: 1.582 });

  tree(s, -0.515, 0.37, 3.526, 7.439, 15);
  card(s, { x: 10.1, y: 2.977, w: 1.928, h: 2.715, adj: 5269, fill: C.white, rotate: 10.42 });
  tab(s, { x: 6.549, y: 4.829, rotate: 84.64, label: 'Survival Training', lx: 6.384, ly: 5.13, lrot: 354.64 });
  skullBadge(s, 9.313, 0.929, 0.879);
  slideBadge(s, 8);
}

// 9 - Two rated course cards + "Discover What Nature Provides".
function slide09(pptx) {
  const s = pptx.addSlide();
  tree(s, -0.236, 0.319, 3.526, 7.439, 30, { rotate: 354.44 });

  const cards = [
    { cx: 0.667, rot: 355.68, fill: C.white, tx: 3.502, title: 'Eco Education', star: 3.603, sw: 3.05, bw: 2.563 },
    { cx: 6.96, rot: 5.17, fill: C.olive, tx: 9.788, title: 'Forest Exploration', star: 9.889, sw: 3.099, bw: 2.356 },
  ];
  cards.forEach(c => {
    card(s, { x: c.cx, y: 1.115, w: 2.521, h: 3.551, adj: 7385, fill: c.fill, rotate: c.rot });
    text(s, c.title, { x: c.tx, y: 1.742, w: 3.688, h: 0.505, fontSize: 24, bold: true, color: C.black });
    text(s, 'Outdoor Exploration Skills', { x: c.tx, y: 2.25, w: c.sw, h: 0.337, fontFace: BODY_FONT, fontSize: 14, bold: true, color: C.green });
    for (let i = 0; i < 5; i++) {
      s.addShape('star5', { x: c.star + i * 0.2438, y: 2.802, w: 0.206, h: 0.206, fill: { color: C.olive }, line: noLine() });
    }
    body(s, 'Environmental Education Through Scouting', { x: c.tx, y: 3.359, w: c.bw, h: 0.679 });
  });

  heading(s, {
    x: 1.592, y: 5.632, w: 10.15, h: 0.841, ex: 5.234, ey: 5.293, ew: 2.865, align: 'center',
    eyebrow: 'Edible Plants',
    runs: [['Discover What '], ['Nature', C.green], [' Provides']],
  });
  tree(s, 11.589, 2.788, 2.484, 5.241, 15, { rotate: 12.03 });
  slideBadge(s, 9);
}

// 10 - "Living Fully in the Outdoors": tilted photo cards + two stat cards.
function slide10(pptx) {
  const s = pptx.addSlide();
  tree(s, 9.614, 0.781, 3.331, 7.029, 30, { rotate: 9.13, flipH: true });
  card(s, { x: 9.836, y: 2.304, w: 3.086, h: 1.446, adj: 9872, fill: C.white, rotate: 15 });
  card(s, { x: 6.917, y: 0.394, w: 2.958, h: 4.166, adj: 7385, fill: C.olive, rotate: 97.69 });
  card(s, { x: 8.622, y: 5.312, w: 3.086, h: 1.176, adj: 9872, fill: C.green, rotate: 351.59 });

  text(s, '10:30 PM', { x: 8.836, y: 5.598, w: 1.893, h: 0.438, rotate: 351.59, fontSize: 20, bold: true, color: C.white });
  text(s, 'Teamwork In Wilderness', { x: 8.882, y: 5.911, w: 3.087, h: 0.337, rotate: 351.59, fontSize: 14, bold: true, color: C.white });
  text(s, '70%', { x: 10.837, y: 2.48, w: 1.858, h: 0.841, rotate: 15, fontSize: 44, bold: true, color: C.green3 });
  text(s, 'Adventure Program', { x: 10.703, y: 3.258, w: 2.231, h: 0.337, rotate: 15, fontSize: 14, bold: true, color: C.green });

  heading(s, {
    x: 0.847, y: 2.43, w: 4.451, ey: 1.971, ew: 2.865,
    eyebrow: 'Camping in the Wild',
    runs: [['Living', C.green], [' Fully in the '], ['Outdoors', C.green]],
  });
  body(s, LOREM_MID, { x: 0.847, y: 4.335, w: 4.123, h: 1.582 });

  tree(s, -0.292, 2.325, 2.655, 5.602, 15, { rotate: 352.91 });
  card(s, { x: 5.691, y: 3.163, w: 2.521, h: 3.551, adj: 7385, fill: C.white, rotate: 86.39 });
  slideBadge(s, 10);
}

// 11 - "Following the Signs of Nature".
function slide11(pptx) {
  const s = pptx.addSlide();
  card(s, { x: 6.088, y: 1.14, w: 3.759, h: 5.295, adj: 5269, fill: C.olive, rotate: 346.53 });
  tree(s, 10.43, 1.018, 3.331, 7.029, 15, { rotate: 9.13, flipH: true });
  card(s, { x: 8.858, y: -1.007, w: 2.465, h: 3.472, adj: 5269, fill: C.white, rotate: 3.43 });

  heading(s, {
    x: 0.82, y: 2.96, w: 4.772, ey: 2.501, ew: 2.865,
    eyebrow: 'Tracking & Awareness',
    runs: [['Following', C.green], [' the Signs of '], ['Nature', C.green]],
  });
  body(s, LOREM_MID, { x: 0.82, y: 4.802, w: 4.367, h: 1.279 });

  card(s, { x: 9.419, y: 4.544, w: 2.9, h: 4.084, adj: 5269, fill: C.white, rotate: 11.77 });
  text(s, '70%', { x: 10.656, y: 2.921, w: 1.858, h: 0.841, fontSize: 44, bold: true, color: C.green3 });
  text(s, 'Navigation', { x: 10.656, y: 3.667, w: 2.231, h: 0.337, fontSize: 14, bold: true, color: C.green });

  tree(s, -1.859, -1.406, 3.331, 7.029, 15, { rotate: 41.86 });
  tab(s, { x: 11.013, y: 0.688, rotate: 93.5, label: 'Nature Edition', lx: 10.846, ly: 0.99, lrot: 3.5 });
  skullBadge(s, 9.485, 3.955, 1.033, C.yellow, SH_YELLOW);
  slideBadge(s, 11);
}

// 12 - "Adventures Beneath the Moon": two labelled portrait cards.
function slide12(pptx) {
  const s = pptx.addSlide();
  tree(s, 4.108, 2.951, 2.335, 4.927, 15, { rotate: 7.4 });
  tree(s, -0.201, -0.5, 4.932, 8.31, 30);
  card(s, { x: 0.978, y: 2.203, w: 3.077, h: 4.334, adj: 5088, fill: C.olive, rotate: 355.42 });
  text(s, 'Scout Community', { x: 1.112, y: 5.846, w: 3.077, h: 0.404, rotate: 355.42, fontSize: 18, bold: true, color: C.white, align: 'center' });

  heading(s, {
    x: 7.963, y: 1.734, w: 3.972, h: 2.322, ey: 1.275, ew: 3.61,
    eyebrow: 'Night Hiking & Stargazing',
    runs: [['Adventures ', C.black], ['Beneath', C.green, true], ['the ', C.black], ['Moon', C.green]],
  });
  body(s, LOREM_SHORT, { x: 7.963, y: 4.259, w: 4.123, h: 0.976 });

  tab(s, { x: 8.274, y: 5.34, label: '+902', lx: 8.108, ly: 5.641, size: 24 });
  tab(s, { x: 9.535, y: 5.34, fill: C.green3, label: '70%', lx: 9.368, ly: 5.641, size: 24 });

  card(s, { x: 3.575, y: 0.838, w: 3.077, h: 4.334, adj: 5088, fill: C.white, rotate: 7.4 });
  text(s, 'Leadership Through Nature', { x: 3.419, y: 4.345, w: 2.957, h: 0.707, rotate: 7.4, fontSize: 18, bold: true, color: C.olive, align: 'center' });
  slideBadge(s, 12);
}

// 13 - "Brotherhood in the Forest": dark hill arc, quote card, stat pair.
function slide13(pptx) {
  const s = pptx.addSlide();
  tree(s, 6.064, 0.612, 2.271, 4.793, 15);
  s.addShape('pie', {
    x: -0.658, y: 5.646, w: 14.649, h: 4.054, flipV: true,
    angleRange: [0, 180], fill: { color: C.olive }, line: noLine(),
  });
  tree(s, -0.191, 1.527, 2.582, 5.447, 30, { rotate: 355.56 });
  tree(s, 10.396, 0.193, 3.602, 7.602, 30, { rotate: 4.26, flipH: true });

  heading(s, {
    x: 0.844, y: 1.468, w: 5.376, ey: 1.008, ew: 2.865,
    eyebrow: 'Group Bonding Outdoors',
    runs: [['Brotherhood', C.green], [' in the Forest']],
  });

  freeform(s, QUOTE, { x: 9.631, y: 1.254, w: 0.179, h: 0.155, color: C.green });
  body(s, [
    { text: '\u201cIn every walk with nature one receives far more than he seeks.\u201d ' },
    { text: '\u2013 John Muir', options: { bold: true, italic: true } },
  ], { x: 9.533, y: 1.496, w: 3.171, h: 0.673 });

  card(s, { x: 9.286, y: 3.022, w: 3.545, h: 1.209, adj: 9872, fill: C.white });
  text(s, 'Forest Adventure Learning Series', { x: 9.52, y: 3.186, w: 3.311, h: 0.774, fontSize: 20, bold: true, color: C.olive });

  [['12M', C.green, 0.844], ['21K', C.olive, 2.242]].forEach(([value, color, x]) => {
    text(s, value, { x, y: 3.676, w: 1.858, h: 0.64, fontSize: 32, bold: true, color });
    text(s, 'Wilderness', { x, y: 4.32, w: 2.231, h: 0.303, fontSize: 12, bold: true, color: C.green });
  });
  slideBadge(s, 13);
}

// 14 - "Learning Through Fun and Play": collage of tilted cards.
function slide14(pptx) {
  const s = pptx.addSlide();
  tree(s, 9.0, -1.971, 4.635, 9.781, 30, { flipH: true });
  card(s, { x: 9.346, y: 0.526, w: 3.759, h: 5.295, adj: 5269, fill: C.olive, rotate: 278.51 });

  heading(s, {
    x: 0.744, y: 1.398, w: 5.376, ey: 0.938, ew: 2.865,
    eyebrow: 'Nature Scouting Games',
    runs: [['Learning Through '], ['Fun', C.green], [' and '], ['Play', C.green]],
  });
  body(s, LOREM_TINY, { x: 0.75, y: 3.07, w: 4.123, h: 0.673 });
  tab(s, { x: 1.061, y: 4.658, label: '+543', lx: 0.895, ly: 4.959, size: 24 });

  tree(s, -0.515, 0.37, 3.526, 7.439, 15);
  card(s, { x: 6.427, y: 2.492, w: 3.03, h: 4.267, adj: 5269, fill: C.green, rotate: 284.38 });
  card(s, { x: 3.48, y: 3.728, w: 2.539, h: 3.575, adj: 5269, fill: C.white, rotate: 276.01 });

  tab(s, { x: 6.298, y: 5.075, h: 1.639, rotate: 107.6, label: 'Real-World Outdoor Training', lx: 5.787, ly: 5.652, lw: 1.699, lrot: 17.6 });
  card(s, { x: 9.621, y: 4.739, w: 2.375, h: 0.992, adj: 12688, fill: C.white, rotate: 7.91, shadow: SH_SOFT });
  text(s, '+780', { x: 9.88, y: 4.84, w: 1.858, h: 0.841, rotate: 7.97, fontSize: 44, bold: true, color: C.black, align: 'center' });
  skullBadge(s, 5.783, 3.116, 0.675, C.yellow, SH_YELLOW);
  slideBadge(s, 14);
}

// 15 - "Cultural Nature Knowledge": rotated headline group + stat pair.
function slide15(pptx) {
  const s = pptx.addSlide();
  card(s, { x: 3.775, y: 1.275, w: 3.421, h: 4.818, adj: 5269, fill: C.white, rotate: 354.18 });
  tree(s, 2.29, 4.182, 1.961, 4.227, 30, { rotate: 12.39 });
  card(s, { x: -0.213, y: 4.552, w: 2.634, h: 3.709, adj: 5269, fill: C.olive, rotate: 353.92 });
  card(s, { x: -0.017, y: -0.472, w: 3.421, h: 4.818, adj: 5269, fill: C.olive, rotate: 354.18 });

  heading(s, {
    x: 7.788, y: 1.545, w: 5.222, ex: 7.69, ey: 1.152, ew: 3.892, rotate: 354.67,
    eyebrow: 'Honoring Indigenous Outdoor Wisdom',
    runs: [['Cultural '], ['Nature', C.green], [' Knowledge']],
  });
  body(s, LOREM_SHORT, { x: 7.923, y: 3.261, w: 4.123, h: 0.976, rotate: 354.67 });

  [
    { pct: '30%', label: 'Navigation', px: 8.123, py: 5.019, lx: 8.158, ly: 5.489 },
    { pct: '70%', label: 'Survival ', px: 10.009, py: 4.788, lx: 10.043, ly: 5.258 },
  ].forEach(g => {
    text(s, g.pct, { x: g.px, y: g.py, w: 1.858, h: 0.572, rotate: 354.22, fontSize: 28, bold: true, color: C.olive });
    text(s, g.label, { x: g.lx, y: g.ly, w: 2.231, h: 0.303, rotate: 354.22, fontSize: 12, bold: true, color: C.green });
  });

  tree(s, 12.238, 3.661, 2.466, 5.315, 15, { rotate: 325.04, flipH: true });
  tab(s, { x: 3.273, y: 3.514, rotate: 84.09, label: 'Learning Path', lx: 3.107, ly: 3.815, lrot: 354.09 });
  card(s, { x: 1.548, y: 5.636, w: 2.375, h: 0.992, adj: 12688, fill: C.white, rotate: 353.97 });
  text(s, '+780', { x: 1.806, y: 5.696, w: 1.858, h: 0.841, rotate: 354.03, fontSize: 44, bold: true, color: C.black, align: 'center' });
  tab(s, { x: 6.689, y: 5.162, rotate: 83.82, label: '+543', lx: 6.524, ly: 5.464, lrot: 353.82, size: 24 });
  slideBadge(s, 15);
}

// 16 - Pricing: three tilted plan cards over "Live Peacefully with the Wild".
function slide16(pptx) {
  const s = pptx.addSlide();
  tree(s, -0.24, 1.345, 3.056, 6.448, 15, { rotate: 355.03 });
  tree(s, 9.785, 0.373, 3.585, 7.564, 30, { rotate: 6.4, flipH: true });

  const plans = [
    { rot: 355.02, accent: C.green, num: '01', name: 'Basic', price: '$9.99/month',
      cx: 1.687, cy: 1.004, bx: 1.713, by: 0.855, ox: 2.718, oy: 0.932, nx: 2.558, ny: 0.957,
      tx: 1.674, ty: 1.354, px: 1.78, py: 2.461, dx: 1.936, dy: 2.938, dh: 0.673 },
    { rot: 3.47, accent: C.green2, num: '02', name: 'Premium', price: '$14.99/month',
      cx: 5.322, cy: 1.037, bx: 5.498, by: 0.884, ox: 6.563, oy: 0.96, nx: 6.403, ny: 0.985,
      tx: 5.45, ty: 1.383, px: 5.391, py: 2.493, dx: 5.454, dy: 2.974, dh: 0.667 },
    { rot: 356.39, accent: C.yellow, num: '03', name: 'Ultimate', price: '$19.99/month',
      cx: 8.978, cy: 1.034, bx: 9.028, by: 0.882, ox: 10.042, oy: 0.958, nx: 9.883, ny: 0.984,
      tx: 8.988, ty: 1.382, px: 9.056, py: 2.491, dx: 9.197, dy: 2.972, dh: 0.667 },
  ];
  plans.forEach(p => {
    card(s, { x: p.cx, y: p.cy, w: 2.68, h: 3.042, adj: 6835, fill: C.white, rotate: p.rot });
    freeform(s, BANNER, { x: p.bx, y: p.by, w: 2.428, h: 1.3, color: p.accent, rotate: p.rot });
    s.addShape('ellipse', {
      x: p.ox, y: p.oy, w: 0.348, h: 0.348, rotate: p.rot,
      fill: { color: C.dark }, line: noLine(), shadow: SH_DOT(),
    });
    text(s, p.num, { x: p.nx, y: p.ny, w: 0.667, h: 0.303, rotate: p.rot, fontSize: 12, bold: true, color: p.accent, align: 'center' });
    text(s, p.name, { x: p.tx, y: p.ty, w: 2.515, h: 0.438, rotate: p.rot, fontSize: 20, bold: true, color: C.white, align: 'center', valign: 'middle' });
    text(s, p.price, { x: p.px, y: p.py, w: 2.515, h: 0.37, rotate: p.rot, fontSize: 16, bold: true, color: p.accent, align: 'center', valign: 'middle' });
    body(s, LOREM_CARD, { x: p.dx, y: p.dy, w: 2.312, h: p.dh, rotate: p.rot, align: 'center' });
  });

  heading(s, {
    x: 1.545, y: 5.055, w: 5.376, ey: 4.595, ew: 2.865,
    eyebrow: 'Wildlife Safety',
    runs: [['Live '], ['Peacefully', C.green], [' with the Wild']],
  });
  text(s, '+670', { x: 9.384, y: 5.243, w: 1.858, h: 0.64, fontSize: 32, bold: true, color: C.olive, align: 'center' });
  joinButton(s, 9.83, 5.932);
  slideBadge(s, 16);
}

// 17 - "Growing Strong Through Nature": card stack + 50/70% strip.
function slide17(pptx) {
  const s = pptx.addSlide();
  tree(s, 7.927, 0.152, 3.521, 8.131, 30, { rotate: 333.48 });
  card(s, { x: 8.762, y: 0.939, w: 3.985, h: 5.613, adj: 5269, fill: C.white, rotate: 9.51 });
  tree(s, -0.082, 0.222, 3.526, 7.439, 15);

  heading(s, {
    x: 0.932, y: 1.811, w: 5.096, h: 1.447, ey: 1.351, ew: 3.136, size: 40,
    eyebrow: ' Reflections in Nature',
    runs: [['Growing', C.green], [' ', C.black], ['Strong'], [' ', C.black], ['Through'], [' ', C.black], ['Nature', C.green]],
  });
  body(s, LOREM_SHORT, { x: 0.932, y: 3.45, w: 4.123, h: 0.982, color: C.olive });

  card(s, { x: 6.185, y: 1.844, w: 3.394, h: 4.779, adj: 5269, fill: C.olive, rotate: 351.37 });
  tab(s, { x: 8.589, y: 0.574, h: 1.325, rotate: 102.51, label: 'Outdoor Exploration', lx: 8.313, ly: 0.994, lw: 1.231, lrot: 12.51 });

  card(s, { x: 0.524, y: 5.08, w: 6.704, h: 1.525, adj: 6081, fill: C.white });
  [
    { icon: 0.953, tx: 1.622, pct: '50%', label: 'Environmental ', note: 2.78 },
    { icon: 4.237, tx: 4.907, pct: '70%', label: 'Campfire ', note: null },
  ].forEach(g => {
    calendarIcon(s, g.icon, 5.643, 0.4, C.green);
    text(s, g.label, { x: g.tx, y: 5.407, w: 2.231, h: 0.303, fontSize: 12, bold: true, color: C.green });
    text(s, g.pct, { x: g.tx, y: 5.731, w: 1.858, h: 0.64, fontSize: 32, bold: true, color: C.green });
    if (g.note !== null) text(s, 'Your Text Here', { x: g.note, y: 5.832, w: 1.074, h: 0.438, fontSize: 10, bold: true, color: C.grey });
  });
  text(s, 'Your Text Here', { x: 6.064, y: 5.832, w: 1.074, h: 0.438, fontSize: 10, bold: true, color: C.grey });
  slideBadge(s, 17);
}

// 18 - Thank you + contact bar.
function slide18(pptx) {
  const s = pptx.addSlide();
  tree(s, 3.631, 1.935, 2.863, 6.041, 15);
  tree(s, 6.68, 1.293, 3.088, 6.517, 15, { flipH: true });
  tree(s, 10.934, 0.37, 3.526, 7.439, 30, { rotate: 4.98 });
  tree(s, -0.515, 0.37, 3.526, 7.439, 30, { rotate: 355.93 });

  text(s, 'Let the Wild Shape You', { x: 5.099, y: 1.831, w: 3.136, h: 0.337, fontFace: BODY_FONT, fontSize: 14, color: C.green, align: 'center' });
  text(s, [
    { text: 'Thank', options: { color: C.ink } },
    { text: ' ', options: { color: C.black } },
    { text: 'You', options: { color: C.green } },
    { text: '!', options: { color: C.ink } },
  ], { x: 1.283, y: 2.276, w: 10.767, h: 1.717, fontSize: 96, bold: true, align: 'center' });

  card(s, { x: 1.8, y: 4.649, w: 9.733, h: 1.02, adj: 14767, fill: C.white });
  const contacts = [
    { x: 2.129, tx: 2.665, label: 'loremipsum@yourmail.com', color: C.green },
    { x: 5.71, tx: 6.246, label: '+030 3456 7890', color: C.green2 },
    { x: 8.671, tx: 9.207, label: 'www.yourwebsite.com', color: C.green3 },
  ];
  contacts.forEach(c => {
    s.addShape('ellipse', {
      x: c.x, y: 4.959, w: 0.413, h: 0.413, fill: { color: c.color }, line: noLine(), shadow: SH_DOTGREEN(),
    });
    body(s, c.label, { x: c.tx, y: 5.008, w: 2.922, h: 0.303, color: C.olive, lineSpacingMultiple: 1, valign: 'middle' });
  });
  freeform(s, ENVELOPE, { x: 2.25, y: 5.123, w: 0.171, h: 0.098, color: C.white });
  freeform(s, ENV_FLAP, { x: 2.256, y: 5.097, w: 0.16, h: 0.07, color: C.white });
  freeform(s, PHONE, { x: 5.831, y: 5.074, w: 0.171, h: 0.171, color: C.white });
  freeform(s, GLOBE, { x: 8.792, y: 5.074, w: 0.171, h: 0.171, color: C.white });
  slideBadge(s, 18);
}

/* ------------------------------------------------------------------- main */

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK', width: 13.333, height: 7.5 });
  pptx.layout = 'DECK';
  pptx.theme = { headFontFace: FONT, bodyFontFace: BODY_FONT };
  pptx.title = 'SCOUT. Community - Nature Edition';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
    slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18]
    .forEach(fn => fn(pptx));

  return pptx.writeFile({ fileName: path.join(__dirname, OUT_NAME) });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
