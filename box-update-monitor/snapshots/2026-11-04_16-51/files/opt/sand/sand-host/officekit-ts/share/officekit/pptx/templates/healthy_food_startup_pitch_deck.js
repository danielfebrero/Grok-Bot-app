// Recreation of the "Ifood" healthy-food pitch deck (48 slides, 26.667 x 15 in)
// with pptxgenjs. Photographs in the original are replaced by flat placeholder
// shapes cut to the same organic frames the template uses.
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette ---
const C = {
  navy:    '1A1C4D',   // headline / brand colour
  grey:    '818181',   // body copy
  sand:    'EBE5DD',   // decorative blobs
  mist:    'F8F6F7',   // pale wash blobs
  pale:    'E6E6E6',
  mapGrey: 'D0CECF',   // world-map landmass
  rule:    '282828',   // hairline rules
  black:   '000000',
  white:   'FFFFFF',
  photo:   'C4D4DC',   // stand-in for the template photography
  icon:    '5F7D95',   // filled icon sheets
};

// ------------------------------------------------------------ text styles ---
// [size pt, colour, typeface, line-spacing %, letter-spacing pt, align]
const S = {
  label:           { fontSize:22, color:'818181', fontFace:'Montserrat Light' },
  body:            { fontSize:32, color:'818181', fontFace:'Montserrat Light', lineSpacingMultiple:1.2 },
  bodyBold:        { fontSize:32, color:'818181', fontFace:'Montserrat Bold', lineSpacingMultiple:1.2 },
  h1:              { fontSize:100, color:'1A1C4D', fontFace:'Lora Regular', lineSpacingMultiple:0.8, charSpacing:-5 },
  sub:             { fontSize:48, color:'818181', fontFace:'Montserrat Regular' },
  h0:              { fontSize:180, color:'1A1C4D', fontFace:'Lora Regular', lineSpacingMultiple:0.8, charSpacing:-9 },
  h0m:             { fontSize:180, color:'1A1C4D', fontFace:'Montserrat Regular', lineSpacingMultiple:0.8, charSpacing:-9 },
  stat:            { fontSize:100, color:'1A1C4D', fontFace:'Montserrat SemiBold', charSpacing:-5 },
  labelBold:       { fontSize:22, color:'818181', fontFace:'Montserrat Bold' },
  labelBoldIt:     { fontSize:22, color:'818181', fontFace:'Montserrat Bold', italic:true },
  labelIt:         { fontSize:22, color:'818181', fontFace:'Montserrat Regular', italic:true },
  labelBoldOnDark: { fontSize:22, color:'F8F6F7', fontFace:'Montserrat Bold' },
  bodyOnDark:      { fontSize:32, color:'F8F6F7', fontFace:'Montserrat Bold', align:'center', lineSpacingMultiple:1.2 },
  labelOnDark:     { fontSize:22, color:'F8F6F7', fontFace:'Montserrat Light', align:'center' },
  display:         { fontSize:300, color:'1A1C4D', fontFace:'Lora Regular', lineSpacingMultiple:0.8, charSpacing:-15 },
  displayBlack:    { fontSize:300, color:'000000', fontFace:'Lora Regular', lineSpacingMultiple:0.8, charSpacing:-15 },
  h1Right:         { fontSize:100, color:'1A1C4D', fontFace:'Lora Regular', align:'right', lineSpacingMultiple:0.8, charSpacing:-5 },
  h1Italic:        { fontSize:100, color:'1A1C4D', fontFace:'Lora Regular', italic:true, lineSpacingMultiple:0.8, charSpacing:-5 },
  tiny:            { fontSize:20, color:'C0C0C0', fontFace:'Montserrat Light', charSpacing:4 },
};

// ------------------------------------------------- repeated copy ---
const T = {
  ACT:     'A top priority in 2021 will be to continue our work across the four pillars of the\u00a0ACT-Accelerator, to achieve equitable access to safe and effective vaccines',
  LACUS:   'Lacus luctus accumsan tortor posuere. Ultricies mi eget mauris pharetra et ultrices neque.',
  LACUS2:  'Lacus luctus accumsan tortor posuere. Ultricies mi eget mauris pharetra et ultrices neque enaque.',
  MALES:   'Malesuada bibendum arcu vitae elementum. Leo vel fringilla est ullamcorper eget nulla facilisi etiam. Commodo elit at imperdiet dui accumsan sit amet nulla facilisi.',
  MALES2:  'Malesuada bibendum arcu vitae elementum. Leo vel fringilla est ullamcorper eget nulla facilisi etiam.',
  LOREM:   'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.',
  LOREM2:  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor. Lacus luctus accumsan tortor posuere. Ultricies mi eget mauris pharetra et ultrices neque.',
  CRAS:    'Cras pulvinar mattis nunc sed blandit libero volutpat sed cras. Purus semper eget duis.',
  NUNC:    'Nunc faucibus a palette pellentesque sit amet porttitor eget. Risus pretium risus.',
  HEART:   'Heart disease, Stroke, and Other Cardiovascular Diseases cause 1 in 3 Deaths in the US',
  SUGAR:   'Higher Sugar Content that Kills Us Slowly',
  SOLVE:   'If there’s a problem, there’s always a solution',
  DIET:    'Healthy diet and food can make your life is good',
};

// ----------------------------------------------------------- blob shapes ---
// Every organic shape in the template is one of eight closed bezier outlines,
// stored as unit-square path data: [x,y] = point, [x1,y1,x2,y2,x,y] = cubic,
// 'z' = close.  scalePath() maps them onto an actual box.
const PATHS = {
  blob10:   [[0.0512,0.9529],[0.1539,1.1615,0.441,0.6034,0.7098,0.7182],[1.1874,0.9222,1.0701,-0.1013,0.3968,0.0082],[0.0196,0.0696,-0.0715,0.7038,0.0512,0.9529],'z'],
  blob8:    [[0.0004,1],[1,1],[1,1,0.9793,0.3915,0.9143,0.2913],[0.7292,0.0061,0.4295,0.0998,0.2122,0.0064],[-0.0188,-0.0929,0.0004,1,0.0004,1],'z'],
  blob12:   [[0.0023,0.9841],[0.0023,0.9841,0.3553,0.8766,0.6266,0.983],[0.8262,1.0612,1.014,0.8762,0.9992,0.3021],[0.9852,-0.2386,0.6583,0.0994,0.2599,0.1589],[0.0827,0.1853,0,0.4432,0,0.4432],[0.0023,0.9841],'z'],
  blob14:   [[0.0796,0.9988],[0.2724,0.9387,0.3128,0.5229,0.4952,0.6555],[0.6645,0.7786,0.7771,0.9517,0.8826,0.8842],[1.1663,0.7026,0.8778,0.22,0.6283,0.0118],[0.5236,-0.0757,0.439,0.3555,0.2915,0.2792],[-0.0567,0.0993,-0.0449,1.0376,0.0796,0.9988],'z'],
  blob18:   [[0.0165,1],[1,1],[1,1,0.9101,0.6424,0.8303,0.5483],[0.7522,0.4562,0.6772,0.3955,0.5212,0.4355],[0.4111,0.4637,0.5954,0,0.3169,0],[0.2301,0,0.0656,0.1031,0.0245,0.2396],[-0.0255,0.4063,0.0165,1,0.0165,1],'z'],
  blob22:   [[1,0.966],[1,0],[0.4706,0.0179,0.5395,0.388,0.3593,0.5389],[0.0642,0.786,-0.0704,0.8979,0.0364,0.9691],[0.1432,1.0403,1,0.966,1,0.966],'z'],
  blobHalf: [[0,0],[0,1],[0.9413,1],[0.9752,0.671,0.9782,0.321,1,0],[0,0],'z'],
  leaf:     [[0.8788,0.2042],[1.0247,0.4266,1.0555,0.7195,0.8785,0.8935],[0.7085,1.0607,0.4458,1.0126,0.2425,0.8761],[0.1047,0.7835,-0.0073,0.6514,0.0003,0.4895],[0.0063,0.3636,0.0874,0.2581,0.1848,0.1716],[0.3416,0.0324,0.5417,-0.0543,0.7154,0.0387],[0.7827,0.0747,0.8345,0.1366,0.8788,0.2042],'z'],
};

// Photo frames: the template masks its photography with these outlines, so the
// placeholder rectangles get the very same silhouettes.
const FRAMES = {
  frameA:    [[0.5641,0],[0.6156,0.0006,0.6665,0.0125,0.7154,0.0386],[0.7827,0.0747,0.8345,0.1366,0.8788,0.2041],[1.0247,0.4266,1.0555,0.7195,0.8785,0.8935],[0.7085,1.0607,0.4458,1.0127,0.2425,0.8761],[0.1047,0.7835,-0.0073,0.6514,0.0004,0.4895],[0.0063,0.3636,0.0875,0.2581,0.1848,0.1716],[0.2975,0.0716,0.4326,-0.0014,0.5641,0],'z'],
  frameB:    [[0.3927,0],[0.4181,-0.0003,0.4451,0.001,0.4738,0.004],[1.0131,0.0613,1,0.3575,1,0.3575],[1,0.9931],[0,1],[0,0.3633],[0,0.3633,0.0127,0.0047,0.3927,0],'z'],
  frameC:    [[0.3934,0],[0.6329,-0.0003,0.8967,0.0931,0.9696,0.3266],[1.1779,0.9931,0.2428,1.2355,0.3575,0.7285],[0.4221,0.4432,-0.1433,0.2209,0.0349,0.0919],[0.1147,0.0342,0.2497,0.0002,0.3934,0],'z'],
  frameD:    [[1,0.3567],[1,0.3572,1,0.3575,1,0.3575],[1,0.9931],[0,1],[0,0.3633],[0,0.3633,0.0144,-0.0447,0.4738,0.004],[0.9963,0.0595,1.0003,0.3392,1,0.3567],'z'],
  frameFull: [[0,0],[1,0],[0.9726,0.329,0.9701,0.679,0.9525,1],[0,1],'z'],
  frameWide: [[0,0],[1,0],[1,0.5534],[0.9755,0.5635],[0.7245,0.6778,0.5465,1.093,0.3366,0.981],[0.2246,0.9213,0.1095,0.8475,0,0.7651],'z'],
  rect:      [[0,0],[1,0],[1,1],[0,1],'z'],
};

// ---------------------------------------------------------------- helpers ---
// Turn unit-square path data into pptxgenjs custGeom points inside w x h.
function scalePath(pts, w, h) {
  return pts.map(p => {
    if (p === 'z') return { close: true };
    if (p.length === 2) return { x: p[0] * w, y: p[1] * h };
    return { x: p[4] * w, y: p[5] * h,
             curve: { type: 'cubic', x1: p[0] * w, y1: p[1] * h, x2: p[2] * w, y2: p[3] * h } };
  });
}

const NO_LINE = { color: C.white, transparency: 100 };   // shapes in this deck are unstroked

function custom(s, pathData, x, y, w, h, fill, rot, line) {
  const o = { x, y, w, h, points: scalePath(pathData, w, h), line: line || NO_LINE };
  if (fill) o.fill = { color: fill };
  if (rot) o.rotate = rot;
  s.addShape('custGeom', o);
}

const blob = (s, name, x, y, w, h, fill, rot) => custom(s, PATHS[name], x, y, w, h, fill, rot);

// Picture placeholders. The template ships them empty, so they reserve their
// silhouette without painting anything; pass a colour for the one frame that
// actually carries artwork.
const photoFrame = (s, frame, x, y, w, h, rot, fill) => custom(s, FRAMES[frame], x, y, w, h, fill, rot);
const rule = (s, x, y, w, h, fill, rot) =>
  s.addShape('rect', { x, y, w, h, fill: { color: fill }, line: NO_LINE, rotate: rot || 0 });

// Every text box in the deck is middle-anchored, inset 4pt, and grows to fit.
function txt(s, text, x, y, w, h, style) {
  s.addText(text, { x, y, w, h, valign: 'middle', margin: 4, fit: 'resize', ...style });
}

// The corner ribbon of three overlapping leaves used on many slides.
function leafCluster(s, x, y, fill) {
  blob(s, 'blob14', x, y, 9.163, 2.989, fill, 93.03);
  blob(s, 'leaf', x + 1.847, y + 1.207, 1.378, 1.272, fill);
  blob(s, 'leaf', x + 0.94, y + 3.968, 1.596, 1.473, fill);
}

// land polygons: [x, y, w, h, [normalized outline pts]]
const MAP_LAND = [
  [15.635,7.786,3.177,3.546,'D0CECF',[0.572,0.069,0.588,0.071,0.598,0.083,0.654,0.095,0.752,0.092,0.762,0.115,0.752,0.142,0.724,0.112,0.724,0.12,0.789,0.222,0.793,0.257,0.816,0.279,0.829,0.31,0.882,0.349,0.886,0.372,0.904,0.383,0.996,0.361,1,0.368,0.979,0.423,0.938,0.478,0.859,0.533,0.826,0.582,0.827,0.627,0.846,0.659,0.84,0.695,0.845,0.716,0.828,0.742,0.795,0.757,0.756,0.792,0.77,0.823,0.766,0.849,0.729,0.864,0.728,0.896,0.668,0.969,0.624,0.99,0.599,0.986,0.545,1,0.523,0.993,0.514,0.97,0.52,0.958,0.468,0.873,0.46,0.818,0.426,0.766,0.421,0.74,0.454,0.675,0.43,0.591,0.391,0.55,0.383,0.53,0.394,0.512,0.395,0.471,0.346,0.462,0.305,0.435,0.223,0.46,0.192,0.453,0.131,0.46,0.089,0.433,0.057,0.394,0.009,0.353,0.014,0.35,0.001,0.323,0.022,0.281,0.015,0.221,0.061,0.15,0.109,0.124,0.123,0.081,0.167,0.045,0.169,0.027,0.22,0.036,0.293,0.01,0.338,0.011,0.395,0,0.419,0.006,0.411,0.018,0.418,0.028,0.415,0.056,0.477,0.077,0.501,0.094,0.537,0.102,0.546,0.074,0.557,0.066,0.572,0.069]],
  [12.664,9.061,2.172,3.5,'EBE5DD',[0.084,0.109,0.069,0.067,0.132,0.02,0.211,0,0.216,0.044,0.218,0.017,0.251,0.014,0.237,0.005,0.29,0.025,0.409,0.022,0.398,0.027,0.518,0.083,0.65,0.112,0.677,0.147,0.643,0.175,0.667,0.19,0.689,0.196,0.725,0.173,0.787,0.188,0.792,0.202,0.89,0.203,0.993,0.238,1,0.255,0.991,0.283,0.913,0.34,0.907,0.398,0.861,0.463,0.711,0.511,0.695,0.567,0.592,0.649,0.495,0.645,0.531,0.679,0.507,0.706,0.417,0.717,0.399,0.753,0.351,0.75,0.384,0.773,0.357,0.774,0.347,0.81,0.307,0.825,0.338,0.859,0.263,0.924,0.297,0.967,0.354,0.989,0.269,1,0.153,0.952,0.165,0.933,0.133,0.873,0.151,0.853,0.14,0.832,0.167,0.836,0.192,0.767,0.168,0.762,0.164,0.741,0.178,0.725,0.173,0.693,0.211,0.64,0.239,0.422,0.109,0.363,0,0.233,0.032,0.202,0.012,0.199,0.013,0.17,0.084,0.109]],
  [15.975,4.648,1.914,3.276,'D0CECF',[0.94,0.832,0.91,0.874,0.937,0.897,0.786,0.912,0.832,0.948,0.793,0.945,0.797,0.97,0.764,0.961,0.772,0.945,0.714,0.912,0.705,0.882,0.565,0.816,0.537,0.825,0.54,0.843,0.684,0.916,0.64,0.913,0.653,0.931,0.618,0.948,0.621,0.918,0.46,0.84,0.318,0.863,0.316,0.889,0.262,0.904,0.239,0.946,0.178,0.973,0.108,0.981,0.074,0.963,0.033,0.965,0.017,0.935,0.033,0.861,0.211,0.857,0.222,0.809,0.133,0.771,0.205,0.766,0.204,0.747,0.238,0.753,0.345,0.711,0.384,0.677,0.457,0.666,0.437,0.62,0.478,0.585,0.504,0.614,0.474,0.64,0.51,0.663,0.663,0.645,0.808,0.655,0.824,0.722,0.796,0.742,0.796,0.771,0.909,0.771,0.928,0.814,0.959,0.819,0.94,0.832]],
  [21.704,9.736,3.081,2.241,'D0CECF',[0.595,0.481,0.502,0.344,0.445,0.172,0.408,0.322,0.336,0.263,0.348,0.2,0.282,0.184,0.291,0.201,0.257,0.214,0.242,0.266,0.222,0.268,0.206,0.241,0.178,0.255,0.166,0.295,0.155,0.29,0.149,0.314,0.137,0.304,0.122,0.353,0.003,0.416,0.001,0.52,0.036,0.653,0.024,0.686,0.059,0.711,0.156,0.682,0.169,0.659,0.24,0.632,0.276,0.631,0.312,0.651,0.336,0.708,0.37,0.656,0.356,0.713,0.374,0.693,0.411,0.778,0.456,0.8,0.477,0.784,0.503,0.806,0.557,0.754,0.598,0.632,0.608,0.557,0.595,0.481]],
  [8.643,4.288,7.216,5.021,'D0CECF',[0.709,0.495,0.692,0.511,0.716,0.502,0.722,0.527,0.717,0.535,0.693,0.552,0.654,0.553,0.636,0.571,0.641,0.577,0.667,0.566,0.659,0.58,0.666,0.597,0.687,0.61,0.66,0.632,0.657,0.618,0.67,0.61,0.649,0.613,0.628,0.635,0.631,0.653,0.604,0.663,0.593,0.701,0.589,0.686,0.591,0.725,0.555,0.775,0.562,0.825,0.548,0.807,0.544,0.782,0.52,0.773,0.504,0.773,0.499,0.788,0.488,0.78,0.473,0.781,0.452,0.805,0.454,0.868,0.472,0.895,0.491,0.892,0.505,0.863,0.521,0.864,0.509,0.918,0.542,0.92,0.543,0.969,0.557,0.984,0.571,0.977,0.584,0.988,0.578,0.995,0.571,0.984,0.562,1,0.529,0.974,0.517,0.942,0.49,0.934,0.469,0.914,0.461,0.921,0.451,0.917,0.404,0.884,0.397,0.85,0.357,0.789,0.35,0.767,0.34,0.76,0.373,0.851,0.358,0.832,0.357,0.818,0.341,0.804,0.346,0.801,0.324,0.746,0.301,0.727,0.278,0.668,0.282,0.598,0.277,0.581,0.287,0.584,0.289,0.573,0.26,0.543,0.238,0.494,0.216,0.473,0.218,0.459,0.21,0.462,0.211,0.451,0.201,0.446,0.197,0.452,0.174,0.43,0.146,0.426,0.129,0.411,0.12,0.429,0.102,0.439,0.11,0.412,0.069,0.472,0.025,0.505,0.018,0.503,0.077,0.447,0.058,0.444,0.057,0.45,0.046,0.441,0.04,0.449,0.037,0.43,0.031,0.435,0.016,0.42,0.014,0.395,0.041,0.379,0.044,0.361,0.031,0.366,0.007,0.362,0,0.346,0.014,0.334,0.04,0.34,0.007,0.304,0.011,0.295,0.023,0.295,0.07,0.254,0.075,0.263,0.158,0.274,0.196,0.294,0.213,0.281,0.247,0.278,0.255,0.268,0.266,0.283,0.275,0.274,0.281,0.283,0.291,0.279,0.339,0.294,0.337,0.311,0.371,0.309,0.383,0.322,0.393,0.299,0.378,0.3,0.389,0.29,0.425,0.311,0.441,0.313,0.454,0.301,0.461,0.318,0.475,0.301,0.457,0.271,0.467,0.243,0.479,0.252,0.487,0.271,0.482,0.279,0.493,0.284,0.496,0.3,0.505,0.286,0.512,0.294,0.516,0.318,0.532,0.293,0.532,0.278,0.557,0.291,0.55,0.306,0.556,0.317,0.547,0.331,0.533,0.336,0.523,0.33,0.514,0.348,0.5,0.34,0.519,0.351,0.517,0.362,0.497,0.375,0.482,0.37,0.495,0.38,0.473,0.411,0.469,0.443,0.479,0.448,0.484,0.469,0.492,0.464,0.52,0.486,0.545,0.49,0.552,0.524,0.566,0.541,0.574,0.529,0.566,0.5,0.584,0.478,0.571,0.439,0.58,0.426,0.576,0.392,0.605,0.391,0.629,0.412,0.645,0.448,0.665,0.422,0.683,0.457,0.681,0.472,0.709,0.495]],
  [17.348,4.947,7.93,4.602,'D0CECF',[0.496,0,0.467,0.029,0.455,0.034,0.435,0.026,0.415,0.038,0.387,0.078,0.361,0.081,0.358,0.108,0.347,0.104,0.327,0.112,0.323,0.098,0.312,0.119,0.318,0.185,0.308,0.142,0.311,0.104,0.301,0.098,0.29,0.103,0.275,0.133,0.285,0.181,0.241,0.156,0.237,0.181,0.228,0.172,0.208,0.184,0.185,0.179,0.156,0.206,0.148,0.176,0.14,0.218,0.127,0.218,0.117,0.24,0.1,0.238,0.1,0.252,0.09,0.247,0.086,0.218,0.092,0.213,0.108,0.222,0.121,0.218,0.126,0.203,0.107,0.183,0.063,0.158,0.052,0.176,0.068,0.271,0.05,0.305,0.052,0.322,0.022,0.33,0.027,0.356,0.016,0.35,0.012,0.354,0,0.395,0.002,0.401,0.018,0.398,0.024,0.41,0.019,0.484,0.041,0.481,0.053,0.518,0.067,0.505,0.083,0.532,0.098,0.521,0.088,0.514,0.109,0.498,0.099,0.521,0.129,0.566,0.107,0.574,0.082,0.562,0.056,0.572,0.038,0.59,0.037,0.606,0.048,0.626,0.06,0.632,0.067,0.624,0.077,0.633,0.096,0.624,0.085,0.686,0.088,0.72,0.134,0.844,0.139,0.883,0.145,0.888,0.191,0.858,0.208,0.841,0.234,0.792,0.217,0.764,0.214,0.743,0.202,0.762,0.185,0.764,0.185,0.745,0.18,0.751,0.167,0.717,0.171,0.697,0.2,0.739,0.214,0.733,0.221,0.748,0.247,0.755,0.273,0.752,0.282,0.772,0.296,0.779,0.288,0.786,0.294,0.8,0.304,0.799,0.31,0.786,0.31,0.827,0.338,0.932,0.352,0.908,0.354,0.856,0.391,0.808,0.394,0.793,0.409,0.79,0.418,0.777,0.436,0.824,0.439,0.854,0.452,0.838,0.457,0.846,0.461,0.927,0.478,0.982,0.492,1,0.488,0.955,0.474,0.939,0.465,0.911,0.47,0.877,0.499,0.911,0.501,0.927,0.526,0.888,0.522,0.856,0.504,0.822,0.506,0.806,0.518,0.79,0.529,0.803,0.552,0.778,0.564,0.778,0.598,0.713,0.598,0.685,0.585,0.643,0.604,0.618,0.592,0.61,0.585,0.619,0.575,0.602,0.594,0.577,0.6,0.582,0.596,0.597,0.613,0.586,0.62,0.59,0.618,0.61,0.628,0.617,0.626,0.65,0.642,0.64,0.632,0.587,0.654,0.55,0.668,0.554,0.68,0.542,0.706,0.48,0.711,0.413,0.706,0.403,0.691,0.406,0.682,0.391,0.713,0.337,0.722,0.329,0.761,0.321,0.787,0.331,0.802,0.296,0.821,0.287,0.829,0.304,0.847,0.28,0.84,0.304,0.801,0.362,0.807,0.439,0.854,0.313,0.876,0.32,0.939,0.272,0.93,0.252,0.947,0.218,0.98,0.249,1,0.221,0.99,0.204,0.918,0.154,0.889,0.15,0.885,0.171,0.871,0.156,0.838,0.161,0.824,0.155,0.81,0.134,0.783,0.131,0.766,0.11,0.739,0.099,0.713,0.103,0.704,0.126,0.694,0.131,0.658,0.117,0.653,0.133,0.641,0.116,0.64,0.087,0.614,0.078,0.606,0.092,0.591,0.095,0.547,0.07,0.544,0.062,0.514,0.085,0.553,0.041,0.533,0.02,0.511,0.02,0.496,0]],
  [18.821,5.005,0.817,0.604,'D0CECF',[0.92,0,0.819,0.041,0.721,0.114,0.62,0.154,0.56,0.121,0.532,0.13,0.426,0.188,0.315,0.284,0.224,0.307,0.23,0.378,0.202,0.445,0.126,0.541,0.156,0.6,0.079,0.643,0.06,0.716,0,0.824,0.036,0.878,0.096,0.878,0.116,0.901,0.126,0.973,0.243,0.997,0.346,0.992,0.25,0.882,0.23,0.803,0.271,0.629,0.363,0.49,0.585,0.323,0.61,0.276,0.959,0.137,1,0.067,0.984,0.036,0.92,0]],
  [22.523,7.562,0.521,0.426,'D0CECF',[0.913,0,0.911,0.026,0.928,0.039,0.903,0.054,0.888,0.027,0.85,0.055,0.825,0.113,0.815,0.169,0.828,0.215,0.796,0.347,0.728,0.461,0.6,0.569,0.564,0.579,0.549,0.56,0.564,0.518,0.454,0.702,0.411,0.732,0.216,0.732,0.148,0.76,0.008,0.899,0.007,0.92,0.108,0.916,0.207,0.871,0.276,0.859,0.307,0.829,0.341,0.833,0.401,0.863,0.388,0.954,0.435,0.997,0.489,0.918,0.542,0.894,0.511,0.832,0.53,0.801,0.575,0.85,0.621,0.867,0.694,0.816,0.73,0.84,0.775,0.765,0.815,0.748,0.812,0.81,0.843,0.803,0.888,0.732,0.872,0.66,0.906,0.552,0.909,0.42,0.928,0.38,0.957,0.373,1,0.23,0.957,0.096,0.95,0.019,0.913,0]],
];
const MAP_ISLES = [
  [18.951,4.506,0.198,0.134],[20.656,4.558,0.425,0.274],[18.482,4.607,0.331,0.098],[21.051,4.756,0.283,0.162],[22.814,5.078,0.364,0.139],[23.028,6.777,0.132,0.514],[22.945,7.328,0.267,0.228],[22.002,8.747,0.199,0.284],[22.121,9.155,0.206,0.198]
];

// The two dotted leader lines that link the map pins to their regions.
const MAP_LEADERS = [
  [13.361, 8.715, 1.425, 1.401, [0.0251, 1, -0.0973, 0.2713, 0.2277, -0.0591, 1, 0.0087]],
  [20.288, 9.456, 1.413, 1.449, [0.0173, 1, -0.0819, 0.1971, 0.2456, -0.1223, 1, 0.0418]],
];

function worldMap(s) {
  MAP_LAND.forEach(([x, y, w, h, fill, pts]) => {
    const poly = [];
    for (let i = 0; i < pts.length; i += 2) poly.push([pts[i], pts[i + 1]]);
    poly.push('z');
    custom(s, poly, x, y, w, h, fill);
  });
  MAP_ISLES.forEach(([x, y, w, h]) =>
    s.addShape('ellipse', { x, y, w, h, fill: { color: C.mapGrey }, line: NO_LINE }));
  MAP_LEADERS.forEach(([x, y, w, h, [sx, sy, c1x, c1y, c2x, c2y, ex, ey]]) =>
    custom(s, [[sx, sy], [c1x, c1y, c2x, c2y, ex, ey]], x, y, w, h, null, 0,
      { color: C.navy, width: 2, dashType: 'sysDot' }));
}

// ------------------------------------------------ slide builders ---
function slide01(p) {
  const s = p.addSlide();
  photoFrame(s, 'frameA', 3.678,2.438,10.966,10.124);
  blob(s, 'blob10', -3.794,13.04,14.089,3.059, C.sand);
  blob(s, 'blob18', 20.423,-0.467,9.299,7.423, C.sand, 180);
  txt(s, 'Presentation', 1.286,1.108,2.064,0.486, {...S.label, wrap:false});
  txt(s, 'January 2021', 23.323,1.108,2.058,0.486, {...S.label, wrap:false});
  txt(s, 'Pitch Deck', 20.682,1.108,1.763,0.486, {...S.label, wrap:false});
  txt(s, 'Ifood', 11.861,3.965,9.196,5.444, {...S.display, wrap:false});
  txt(s, 'Healthy Food Pitch Deck Presentation Template', 11.861,8.548,9.368,1.722, S.sub);
}

function slide02(p) {
  const s = p.addSlide();
  blob(s, 'blob8', 15.22,11.551,12.219,3.449, C.sand);
  txt(s, T.SUGAR, 1.286,2.396,15.992,8.417, S.h0);
  blob(s, 'blob14', -1.858,-2.527,7.664,2.228, C.sand, 314.4);
  blob(s, 'blob14', 7.504,15.181,7.664,2.228, C.sand, 314.4);
  blob(s, 'blob14', 23.605,9.779,7.664,2.228, C.sand, 11.62);
  txt(s, 'A top priority in 2021 will be to continue our work across the four pillars of the\u00a0ACT-Accelerator, to achieve equitable access to safe and effective vaccines. Lorem Ipsum Dolor.', 17.49,3.127,7.89,3.903, S.body);
  txt(s, 'Problem', 1.286,1.108,1.435,0.486, {...S.label, wrap:false});
  rule(s, 14.963,-10.314,0.027,23.33, C.rule, 90);
}

function slide03(p) {
  const s = p.addSlide();
  photoFrame(s, 'frameB', 13.227,3.771,12.153,7.458);
  txt(s, T.HEART, 1.284,3.409,15.088,6.156, S.h1);
  txt(s, 'Problem', 1.286,1.108,1.435,0.486, {...S.label, wrap:false});
  rule(s, 14.963,-10.314,0.027,23.33, C.rule, 90);
  txt(s, T.ACT, 1.286,10.042,6.553,1.611, S.label);
  blob(s, 'blob10', 22.702,11.94,6.432,1.22, C.sand, 347.71);
  blob(s, 'blob22', 23.45,-1.529,1.972,5.029, C.sand, 270);
  blob(s, 'blob10', 21.426,13.922,6.432,1.22, C.sand, 347.71);
  blob(s, 'blob10', 15.795,13.922,6.432,1.22, C.sand, 347.71);
}

function slide04(p) {
  const s = p.addSlide();
  photoFrame(s, 'frameC', 7.727,1.346,11.221,12.306);
  txt(s, T.ACT, 3.901,8.623,11.564,4.139, S.sub);
  txt(s, T.SUGAR, 14.437,2.711,10.158,4.733, S.h1Right);
  blob(s, 'blob8', 7.79,13.892,12.219,3.449, C.sand);
  blob(s, 'blob14', 22.156,10.295,9.163,2.989, C.sand, 93.03);
  blob(s, 'leaf', 25.381,11.057,0.482,0.445, C.sand);
  blob(s, 'leaf', 24.003,11.502,1.378,1.272, C.sand);
  blob(s, 'leaf', 23.096,14.263,1.596,1.473, C.sand);
  blob(s, 'blob22', 0,-0.249,2.709,5.796, C.sand, 180);
}

function slide05(p) {
  const s = p.addSlide();
  photoFrame(s, 'frameB', 7.363,2.558,12.153,7.458);
  txt(s, T.HEART, 1.286,6.033,15.088,6.156, S.h1);
  txt(s, T.ACT, 15.842,10.888,9.539,2.603, S.body);
  txt(s, 'Problem', 1.286,1.108,1.435,0.486, {...S.label, wrap:false});
  rule(s, 14.963,-10.314,0.027,23.33, C.rule, 90);
  blob(s, 'blob12', -0.658,14.092,5.323,1.259, C.sand);
  blob(s, 'blob8', -0.475,13.039,11.2,0.852, C.sand);
  blob(s, 'blob22', 23.869,-1.869,3.023,6.467, C.sand);
}

function slide06(p) {
  const s = p.addSlide();
  photoFrame(s, 'rect', 0,0,26.667,15);
  blob(s, 'blob8', 15.699,0,9.681,8.266, C.sand, 180);
  txt(s, T.ACT, 17.689,1.443,6.146,3.903, S.body);
}

function slide07(p) {
  const s = p.addSlide();
  blob(s, 'blob12', -2.592,9.873,5.323,1.259, C.sand);
  blob(s, 'blob8', -2.408,8.821,11.2,0.852, C.sand);
  photoFrame(s, 'frameB', 1.276,6.061,9.123,7.831);
  photoFrame(s, 'frameB', 10.629,6.061,7.434,7.831);
  txt(s, 'Problem', 1.286,2.632,6.077,1.889, S.h1);
  txt(s, T.ACT, 16.257,2.275,9.123,2.603, S.body);
  txt(s, 'Problem', 1.286,1.108,1.435,0.486, {...S.label, wrap:false});
  rule(s, 14.963,-10.314,0.027,23.33, C.rule, 90);
  blob(s, 'blob10', 23.568,12.158,6.432,1.22, C.sand, 347.71);
  blob(s, 'blob10', 22.293,14.14,6.432,1.22, C.sand, 347.71);
  blob(s, 'blob10', 16.661,14.14,6.432,1.22, C.sand, 347.71);
}

function slide08(p) {
  const s = p.addSlide();
  txt(s, 'Problem', 1.286,1.108,1.435,0.486, {...S.label, wrap:false});
  rule(s, 14.963,-10.314,0.027,23.33, C.rule, 90);
  txt(s, 'Global Health Issues 2021', 14.225,2.107,11.155,3.311, S.h1);
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Lacinia quis vel eros donec. Varius sit amet mattis vulputate. ', 14.225,12.28,10.249,1.236, S.label);
  txt(s, 'Covid-19 Omnicron', 14.225,6.161,11.527,0.917, S.sub);
  txt(s, 'Heart Disease and Stroke', 14.225,7.355,11.527,0.917, S.sub);
  txt(s, 'High in Obesity', 14.225,8.55,11.527,0.917, S.sub);
  txt(s, 'Secondhand Smoke Exposure', 14.225,9.744,12.047,0.917, S.sub);
  txt(s, 'Diabetes and Cholesterol', 14.225,10.939,11.527,0.917, S.sub);
  blob(s, 'blob10', -2.257,12.089,10.334,2.244, C.sand, 225);
  photoFrame(s, 'frameA', 1.284,2.711,10.966,10.124);
}

function slide09(p) {
  const s = p.addSlide();
  txt(s, 'Solution', 1.286,1.108,1.381,0.486, {...S.label, wrap:false});
  rule(s, 14.963,-10.314,0.027,23.33, C.rule, 90);
  blob(s, 'blob8', 0.375,11.551,12.219,3.449, C.sand);
  blob(s, 'blob14', -5.776,12.248,7.664,2.228, C.sand, 314.4);
  blob(s, 'blob14', 10.884,13.744,7.664,2.228, C.sand, 64.08);
  txt(s, 'On problem, there’s always about solution. After the problem has explained well, show how your startup can give a simple solution and tell investor how unique you are', 1.284,2.271,23.188,7.578, S.h1);
  txt(s, 'Nulla pharetra diam sit amet nisl. Pretium vulputate sapien nec sagittis aliquam malesuada. Diam ut venenatis tellus in metus vulputate eu.', 16.372,10.25,9.011,2.603, S.body);
}

function slide10(p) {
  const s = p.addSlide();
  photoFrame(s, 'frameD', 16.957,4.086,12.153,7.458, 270);
  txt(s, T.DIET, 1.284,3.131,15.315,3.311, S.h1);
  blob(s, 'blob12', -0.658,14.092,5.323,1.259, C.sand);
  blob(s, 'blob8', -0.475,13.039,11.2,0.852, C.sand);
  txt(s, T.LACUS, 12.403,8.209,5.767,2.603, S.body);
  txt(s, T.LACUS, 2.471,8.209,7.534,1.953, S.body);
  txt(s, '01', 1.284,8.209,0.912,0.653, S.bodyBold);
  txt(s, '02', 11.267,8.209,0.912,0.653, S.bodyBold);
  txt(s, 'Solution', 1.286,1.108,1.381,0.486, {...S.label, wrap:false});
  rule(s, 14.963,-10.314,0.027,23.33, C.rule, 90);
}

function slide11(p) {
  const s = p.addSlide();
  photoFrame(s, 'frameA', 7.362,1.108,11.167,10.309);
  txt(s, 'Solution', 1.286,1.108,1.381,0.486, {...S.label, wrap:false});
  txt(s, T.DIET, 1.284,2.831,9.704,4.733, S.h1);
  txt(s, T.LACUS, 19.516,9.908,5.789,2.603, S.body);
  txt(s, '01', 19.58,8.977,0.912,0.653, S.bodyBold);
  blob(s, 'blob12', -0.658,14.092,5.323,1.259, C.sand);
  blob(s, 'blob8', -0.475,13.039,11.2,0.852, C.sand);
  blob(s, 'blob22', 23.869,-1.869,3.023,6.467, C.sand);
}

function slide12(p) {
  const s = p.addSlide();
  photoFrame(s, 'frameC', 6.45,1.086,11.221,12.306, 352.86);
  photoFrame(s, 'frameA', -0.848,5.913,11.167,10.309);
  txt(s, T.SOLVE, 14.824,5.133,10.557,4.733, S.h1);
  blob(s, 'blob12', -0.658,1.005,5.323,1.259, C.sand);
  blob(s, 'blob8', -0.658,0,11.2,0.852, C.sand, 180);
  blob(s, 'blob22', 24.027,10.168,3.023,6.467, C.sand);
  txt(s, T.LACUS, 18.397,11.289,5.789,2.603, S.body);
  txt(s, '02', 18.461,10.358,0.912,0.653, S.bodyBold);
}

function slide13(p) {
  const s = p.addSlide();
  txt(s, T.SOLVE, 1.286,2.236,14.499,3.311, S.h1);
  txt(s, 'Solution', 1.286,1.108,1.381,0.486, {...S.label, wrap:false});
  rule(s, 14.963,-10.314,0.027,23.33, C.rule, 90);
  txt(s, T.CRAS, 7.363,8.431,5.789,2.603, S.body);
  txt(s, '02', 7.363,7.5,0.912,0.653, S.bodyBold);
  txt(s, T.LACUS, 1.284,8.431,5.789,2.603, S.body);
  txt(s, '01', 1.284,7.5,0.912,0.653, S.bodyBold);
  txt(s, T.NUNC, 13.439,8.431,5.789,2.603, S.body);
  txt(s, '03', 13.439,7.5,0.912,0.653, S.bodyBold);
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit,\u00a0', 19.516,8.433,5.789,1.953, S.body);
  txt(s, '04', 19.516,7.5,0.912,0.653, S.bodyBold);
  blob(s, 'blob12', -0.658,14.092,5.323,1.259, C.sand);
  blob(s, 'blob8', -0.475,13.039,11.2,0.852, C.sand);
  blob(s, 'leaf', 23.544,12.375,0.482,0.445, C.sand);
  blob(s, 'leaf', 22.166,12.82,1.378,1.272, C.sand);
  blob(s, 'leaf', 23.096,14.263,1.596,1.473, C.sand);
  blob(s, 'leaf', 24.38,11.376,4.305,3.974, C.sand);
}

function slide14(p) {
  const s = p.addSlide();
  txt(s, T.SOLVE, 1.286,2.236,14.499,3.311, S.h1);
  txt(s, 'Solution', 1.286,1.108,1.381,0.486, {...S.label, wrap:false});
  rule(s, 14.963,-10.314,0.027,23.33, C.rule, 90);
  txt(s, T.LACUS, 1.284,8.431,7.499,1.953, S.body);
  txt(s, 'Solution 1', 1.284,7.5,3.028,0.653, S.bodyBold);
  blob(s, 'blob12', -0.658,14.092,5.323,1.259, C.sand);
  blob(s, 'blob8', -0.475,13.039,11.2,0.852, C.sand);
  txt(s, T.NUNC, 9.388,8.431,7.499,1.953, S.body);
  txt(s, 'Solution 2', 9.388,7.5,2.932,0.653, S.bodyBold);
  txt(s, T.CRAS, 17.49,8.431,7.499,1.953, S.body);
  txt(s, 'Solution 3', 17.49,7.5,2.932,0.653, S.bodyBold);
  blob(s, 'blob14', 22.156,10.295,9.163,2.989, C.sand, 93.03);
  blob(s, 'leaf', 25.381,11.057,0.482,0.445, C.sand);
  blob(s, 'leaf', 24.003,11.502,1.378,1.272, C.sand);
  blob(s, 'leaf', 23.096,14.263,1.596,1.473, C.sand);
}

function slide15(p) {
  const s = p.addSlide();
  blob(s, 'blob10', 23.568,12.158,6.432,1.22, C.sand, 347.71);
  blob(s, 'blob10', 22.293,14.14,6.432,1.22, C.sand, 347.71);
  blob(s, 'blob10', 16.661,14.14,6.432,1.22, C.sand, 347.71);
  txt(s, T.LACUS, 16.687,7.909,7.47,1.953, S.body);
  txt(s, T.MALES, 16.687,10.26,7.47,1.611, S.label);
  txt(s, 'Solutions for Healthy Eating', 1.286,1.478,15.992,5.861, S.h0);
  txt(s, 'Solution', 1.286,1.108,1.381,0.486, {...S.label, wrap:false});
  rule(s, 14.963,-10.314,0.027,23.33, C.rule, 90);
  photoFrame(s, 'frameB', 1.242,7.5,13.967,11.988);
}

function slide16(p) {
  const s = p.addSlide();
  txt(s, 'About Market', 1.286,1.478,15.781,3.306, S.h0);
  txt(s, '1.6B', 11.716,6.241,5.83,3.153, S.h0m);
  txt(s, 'Total Market', 11.714,5.817,2.903,0.653, S.body);
  txt(s, T.MALES, 11.716,10.043,6.345,1.986, S.label);
  txt(s, '52%', 18.35,5.789,2.903,1.806, S.stat);
  txt(s, 'Asia Pacific', 21.253,6.241,4.128,0.903, S.sub);
  txt(s, 'American', 21.253,8.472,3.638,0.903, S.sub);
  txt(s, 'European', 21.253,10.703,3.638,0.903, S.sub);
  txt(s, '33%', 18.35,8.02,2.903,1.806, S.stat);
  txt(s, '15%', 18.35,10.251,2.903,1.806, S.stat);
  blob(s, 'blob10', -2.64,12.472,10.334,2.244, C.sand, 225);
  blob(s, 'blob18', 19.304,-2.64,9.299,7.423, C.sand, 180);
  photoFrame(s, 'frameA', 1.286,5.482,8.508,7.854);
}

function slide17(p) {
  const s = p.addSlide();
  s.background = { color: C.sand };
  blob(s, 'blob14', 0.889,9.885,35.553,6.634, C.mist, 155.74);
  blob(s, 'blob14', -16.743,-0.548,35.553,6.634, C.mist, 155.74);
  rule(s, 1.286,5.808,11.31,0.042, C.black, 180);
  rule(s, 1.286,9.676,11.31,0.042, C.black, 180);
  txt(s, '500.987', 1.286,6.523,11.31,3.153, S.h0m);
  txt(s, 'Total Available Market', 1.286,6.085,5.31,0.653, S.body);
  txt(s, '374.562', 1.286,10.391,11.31,3.153, S.h0m);
  txt(s, 'Market Potential', 1.286,9.954,5.31,0.653, S.body);
  txt(s, 'Market', 1.286,1.108,1.19,0.486, {...S.label, wrap:false});
  rule(s, 14.963,-10.314,0.027,23.33, C.rule, 90);
  txt(s, T.MALES, 1.284,2.623,13.972,1.953, S.body);
  photoFrame(s, 'frameA', 14.849,4.784,9.624,8.885);
}

function slide18(p) {
  const s = p.addSlide();
  blob(s, 'blob8', 0.375,11.551,12.219,3.449, C.sand);
  txt(s, 'Market', 1.286,1.108,1.19,0.486, {...S.label, wrap:false});
  rule(s, 14.963,-10.314,0.027,23.33, C.rule, 90);
  blob(s, 'blob14', -5.776,12.248,7.664,2.228, C.sand, 314.4);
  blob(s, 'blob14', 10.884,13.744,7.664,2.228, C.sand, 64.08);
  txt(s, 'Marketing Strategy', 1.286,2.354,14.499,1.889, S.h1);
  blob(s, 'leaf', 17.49,3.156,0.527,0.486, C.sand);
  txt(s, 'Strategy 01', 18.295,2.941,4.051,0.917, S.sub);
  txt(s, T.LOREM2, 18.295,3.996,6.179,1.986, S.label);
  blob(s, 'leaf', 17.49,7.111,0.527,0.486, C.sand);
  txt(s, 'Strategy 02', 18.295,6.895,4.051,0.917, S.sub);
  txt(s, T.LOREM2, 18.295,7.951,6.179,1.986, S.label);
  blob(s, 'leaf', 17.49,11.065,0.527,0.486, C.sand);
  txt(s, 'Strategy 03', 18.295,10.85,4.051,0.917, S.sub);
  txt(s, T.LOREM2, 18.295,11.905,6.179,1.986, S.label);
  photoFrame(s, 'frameB', 2.527,5.232,7.732,8.145);
}

function slide19(p) {
  const s = p.addSlide();
  txt(s, 'Our Products', 1.286,2.911,8.62,1.889, S.h1);
  blob(s, 'blob12', -0.658,1.005,5.323,1.259, C.sand);
  blob(s, 'blob8', -0.658,0,11.2,0.852, C.sand, 180);
  blob(s, 'blob22', 24.027,10.168,3.023,6.467, C.sand);
  txt(s, T.LACUS, 1.286,8.94,7.47,1.953, S.body);
  txt(s, T.MALES, 1.286,11.291,7.47,1.611, S.label);
  photoFrame(s, 'frameA', 16.501,6.766,7.536,6.904);
  photoFrame(s, 'frameA', 18.277,-0.839,7.536,6.904, 317.32);
  photoFrame(s, 'frameA', 10.876,1.325,7.536,6.904);
}

function slide20(p) {
  const s = p.addSlide();
  photoFrame(s, 'frameB', 1.834,1.108,9.367,14.09);
  txt(s, 'Vegan Curry Chicken', 12.321,2.722,12.868,1.889, S.h1);
  blob(s, 'blob12', -0.658,1.005,5.323,1.259, C.sand);
  blob(s, 'blob8', -0.658,0,11.2,0.852, C.sand, 180);
  blob(s, 'blob10', 23.568,12.158,6.432,1.22, C.sand, 347.71);
  blob(s, 'blob10', 22.293,14.14,6.432,1.22, C.sand, 347.71);
  blob(s, 'blob10', 16.661,14.14,6.432,1.22, C.sand, 347.71);
  txt(s, T.LACUS, 12.321,5.149,11.035,1.303, S.body);
  txt(s, T.MALES, 12.321,6.729,7.47,1.611, S.label);
  txt(s, 'Carbs', 12.321,8.94,3.308,0.653, S.bodyBold);
  txt(s, '192 g', 12.321,9.592,3.308,1.889, S.h1);
  txt(s, 'Protein', 17.197,8.94,3.308,0.653, S.bodyBold);
  txt(s, '204 g', 17.197,9.592,3.508,1.889, S.h1);
  txt(s, 'Fat', 22.073,8.94,3.308,0.653, S.bodyBold);
  txt(s, '51 g', 22.073,9.592,3.308,1.889, S.h1);
}

function slide21(p) {
  const s = p.addSlide();
  txt(s, 'Products', 1.286,1.108,1.491,0.486, {...S.label, wrap:false});
  rule(s, 14.963,-10.314,0.027,23.33, C.rule, 90);
  txt(s, 'A healthy diet is\u00a0essential for good health and nutrition.', 1.286,2.23,8.438,6.156, S.h1);
  txt(s, 'Malesuada bibendum arcu vitae elementum. Leo vel fringilla est ullamcorper eget nulla facilisi etiam. Commodo elit at imperdiet dui accumsan sit amet nulla.', 1.284,8.623,8.443,3.253, S.body);
  blob(s, 'blob10', -3.009,12.872,10.334,2.244, C.sand, 225);
  txt(s, 'Orange Carpaccio', 12.058,8.721,4.305,0.653, {...S.bodyBold, wrap:false});
  txt(s, 'Salad-Salad Club', 19.041,8.721,4.04,0.653, {...S.bodyBold, wrap:false});
  txt(s, T.MALES2, 12.058,9.513,6.34,1.236, S.label);
  txt(s, T.LACUS2, 19.041,9.513,6.34,1.236, S.label);
  txt(s, 'Total Calories : 1234 kcal', 12.058,11.389,4.314,0.486, S.labelIt);
  txt(s, 'Total Calories : 1025 kcal', 19.041,11.389,4.314,0.486, S.labelIt);
  photoFrame(s, 'frameB', 12.058,4.233,6.34,3.89);
  photoFrame(s, 'frameB', 19.041,4.232,6.34,3.89);
}

function slide22(p) {
  const s = p.addSlide();
  blob(s, 'blob12', -4.366,0.852,5.323,1.259, C.sand);
  blob(s, 'blob8', -0.658,0,11.2,0.852, C.sand, 180);
  blob(s, 'blob10', 23.568,12.158,6.432,1.22, C.sand, 347.71);
  blob(s, 'blob10', 22.293,14.14,6.432,1.22, C.sand, 347.71);
  blob(s, 'blob10', 16.661,14.14,6.432,1.22, C.sand, 347.71);
  txt(s, 'Team', 1.286,1.108,0.961,0.486, {...S.label, wrap:false});
  rule(s, 14.963,-10.314,0.027,23.33, C.rule, 90);
  txt(s, 'Ernest Floyd', 10.405,9.567,3.027,0.653, {...S.bodyBold, wrap:false});
  txt(s, 'Aleksander Kosarov', 17.389,9.567,4.749,0.653, {...S.bodyBold, wrap:false});
  txt(s, T.LACUS2, 10.405,10.359,6.34,1.236, S.label);
  txt(s, 'Commodo viverra maecenas accumsan lacus vel facilisis volutpat est. Proin fermentum leo vel orci porta non pulvinar neque.', 17.389,10.359,6.34,1.611, S.label);
  txt(s, 'Founder & CEO', 10.405,12.235,4.314,0.486, S.labelBoldIt);
  txt(s, 'Chef Project Manager', 17.389,12.235,4.314,0.486, S.labelBoldIt);
  txt(s, 'Dima Gregory', 2.938,9.636,3.369,0.653, {...S.bodyBold, wrap:false});
  txt(s, T.MALES2, 2.938,10.428,6.34,1.236, S.label);
  txt(s, 'Executive Creative Director', 2.938,12.304,4.749,0.486, S.labelBoldIt);
  txt(s, 'Our Team', 1.286,2.311,8.62,1.889, S.h1);
  photoFrame(s, 'frameB', 2.938,5.078,6.34,3.89);
  photoFrame(s, 'frameB', 10.405,5.031,6.34,3.89);
  photoFrame(s, 'frameB', 17.42,5.078,6.34,3.89);
}

function slide23(p) {
  const s = p.addSlide();
  txt(s, 'Our Team', 1.286,3.654,5.997,1.889, {...S.h1, wrap:false});
  txt(s, 'A wonderful serenity has taken possession of my entire soul.', 1.284,8.028,7.571,1.303, S.bodyBold);
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Aliquet nibh praesent tristique. Aliquet nibh praesent tristique.', 1.306,9.735,7.746,1.611, S.label);
  blob(s, 'blob8', 17.183,0,11.2,0.852, C.sand, 180);
  txt(s, 'Team', 1.286,1.108,0.961,0.486, {...S.label, wrap:false});
  rule(s, 14.963,-10.314,0.027,23.33, C.rule, 90);
  txt(s, 'Dima Gregory', 19.516,2.353,3.663,0.653, S.bodyBold);
  txt(s, 'Founder & CEO', 19.516,3.02,4.285,0.486, S.label);
  txt(s, T.LOREM, 19.516,4.3,4.906,1.236, S.label);
  txt(s, 'Ernest Floyd', 19.516,6.41,3.663,0.653, S.bodyBold);
  txt(s, 'Executive Creative Director', 19.516,7.076,4.285,0.486, S.label);
  txt(s, T.LOREM, 19.516,8.357,4.906,1.236, S.label);
  txt(s, 'Bhami Loque', 19.516,10.466,3.663,0.653, S.bodyBold);
  txt(s, 'Design Creative Lead', 19.516,11.133,4.285,0.486, S.label);
  txt(s, T.LOREM, 19.516,12.413,4.906,1.236, S.label);
  blob(s, 'blob10', 5.259,13.083,6.432,1.22, C.sand, 347.71);
  blob(s, 'blob10', 3.983,15.065,6.432,1.22, C.sand, 347.71);
  blob(s, 'blob10', -1.648,15.065,6.432,1.22, C.sand, 347.71);
  photoFrame(s, 'frameA', 14.379,2.111,3.886,3.668);
  photoFrame(s, 'frameA', 14.335,6.167,3.886,3.668);
  photoFrame(s, 'frameA', 14.379,10.224,3.886,3.668);
  // the only bitmap in the template: a flat pale-blue swatch [image]
  photoFrame(s, 'frameA', 14.335, 10.224, 3.973, 3.668, 0, C.photo);
}

function slide24(p) {
  const s = p.addSlide();
  s.background = { color: C.sand };
  photoFrame(s, 'rect', -2.099,3.599,16.685,21.753);
  blob(s, 'blob14', -16.743,-0.548,35.553,6.634, C.mist, 155.74);
  blob(s, 'blob14', 0.889,9.885,35.553,6.634, C.mist, 155.74);
  txt(s, 'Team', 1.286,1.108,0.961,0.486, {...S.label, wrap:false});
  rule(s, 14.963,-10.314,0.027,23.33, C.rule, 90);
  txt(s, 'Ernest Floyd', 12.807,2.999,7.389,1.889, {...S.h1, wrap:false});
  txt(s, 'Founder & CEO', 12.807,4.888,4.285,0.653, S.body);
  txt(s, 'I am alone, and feel the charm of existence in this spot, which was created for the bliss of souls like mine. I am so happy, my dear friend, so absorbed in the exquisite sense.', 12.807,6.212,9.916,2.603, S.body);
}

function slide25(p) {
  const s = p.addSlide();
  txt(s, 'Testimonial', 1.286,1.108,1.868,0.486, {...S.label, wrap:false});
  rule(s, 14.963,-10.314,0.027,23.33, C.rule, 90);
  txt(s, '“Since I take care of my health, Ifood is super useful. The best part is you don’t need to search for any additional information about nutrition, its all in the app.”', 1.284,3.343,21.812,5.531, S.h1Italic);
  blob(s, 'blob12', -0.658,14.092,5.323,1.259, C.sand);
  blob(s, 'blob8', -0.475,13.039,11.2,0.852, C.sand);
  blob(s, 'blob14', 22.156,10.295,9.163,2.989, C.sand, 93.03);
  blob(s, 'leaf', 24.003,11.502,1.378,1.272, C.sand);
  blob(s, 'leaf', 23.096,14.263,1.596,1.473, C.sand);
  txt(s, [{text:'Marcus Spoiler'}, {text:' - Product Manager at Startup', options:S.body}], 3.875,11.175,10.497,0.653, S.bodyBold);
  photoFrame(s, 'frameA', 1.284,10.498,2.175,2.008);
}

function slide26(p) {
  const s = p.addSlide();
  txt(s, 'Marketing', 1.286,1.108,1.684,0.486, {...S.label, wrap:false});
  rule(s, 14.963,-10.314,0.027,23.33, C.rule, 90);
  blob(s, 'blob8', -2.5,12.309,7.626,0.852, C.sand);
  blob(s, 'blob8', 5.337,10.98,7.89,0.852, C.sand);
  blob(s, 'blob8', 13.439,9.774,7.89,0.852, C.sand);
  blob(s, 'blob8', 21.541,8.922,7.89,0.852, C.sand);
  txt(s, 'Market To-Go', 1.286,2.619,8.443,1.889, {...S.h1, wrap:false});
  txt(s, T.LACUS, 1.284,9.847,3.373,1.986, S.label);
  txt(s, 'Step 1', 1.284,8.916,3.028,0.653, S.bodyBold);
  txt(s, T.LACUS, 5.337,9.327,5.293,1.236, S.label);
  txt(s, 'Step 2', 5.337,8.536,3.028,0.653, S.bodyBold);
  txt(s, T.LACUS, 13.439,8.121,5.293,1.236, S.label);
  txt(s, 'Step 3', 13.439,7.33,3.028,0.653, S.bodyBold);
  txt(s, T.LACUS, 21.541,7.033,3.839,1.611, S.label);
  txt(s, 'Step 4', 21.541,6.241,3.028,0.653, S.bodyBold);
  blob(s, 'blob14', 21.897,-0.355,9.163,2.989, C.pale, 93.03);
  blob(s, 'leaf', 23.744,0.851,1.378,1.272, C.pale);
  blob(s, 'leaf', 22.837,3.613,1.596,1.473, C.pale);
  blob(s, 'blob22', 22.865,11.117,3.023,6.467, C.pale, 71.32);
}

function slide27(p) {
  const s = p.addSlide();
  blob(s, 'blob12', -0.658, 1.005, 5.323, 1.259, C.sand);
  worldMap(s);
  blob(s, 'blob8', -0.658, 0, 11.2, 0.852, C.sand, 180);
  blob(s, 'blob22', 24.027, 10.168, 3.023, 6.467, C.sand);
  txt(s, 'Our Milestones', 1.286,2.726,9.079,1.889, {...S.h1, wrap:false});
  txt(s, 'Morbi tristique senectus et netus et. Turpis egestas maecenas pharetra convallis posuere morbi leo. Purus ut faucibus pulvinar elementum integer enim neque. Maecenas accumsan lacus vel facilisis volutpat.', 1.284,9.16,7.893,4.553, S.body);
  // two "Gross Profit / $123B" pins dropped on the map
  [[14.845, 7.045], [18.756, 11.167]].forEach(([x, y]) => {
    blob(s, 'leaf', x, y, 2.703, 2.724, C.navy);
    txt(s, 'Gross Profit', x + 0.265, y + 0.876, 2.172, 0.486, S.labelOnDark);
    txt(s, '$123B', x + 0.587, y + 1.305, 1.516, 0.653, S.bodyOnDark);
  });
}

function slide28(p) {
  const s = p.addSlide();
  blob(s, 'blobHalf', 0,0,12.25,15, C.sand);
  blob(s, 'blob14', -16.743,-0.548,35.553,6.634, C.mist, 155.74);
  txt(s, 'Financial Needs', 1.286,8.069,9.916,5.861, S.h0);
  txt(s, 'Offline-Site', 20.016,7.908,4.635,0.917, S.sub);
  txt(s, 'Online Shop', 20.016,8.971,4.592,0.917, S.sub);
  txt(s, 'Marketing', 20.016,10.019,4.592,0.917, S.sub);
  txt(s, T.LOREM, 13.439,7.908,5.458,2.603, S.body);
  txt(s, '$1450M', 14.255,4.064,9.285,3.153, S.h0m);
  blob(s, 'blob10', 23.568,12.158,6.432,1.22, C.sand, 347.71);
  blob(s, 'blob10', 22.293,14.14,6.432,1.22, C.sand, 347.71);
  blob(s, 'blob10', 16.661,14.14,6.432,1.22, C.sand, 347.71);
  rule(s, 19.012,1.788,0.028,11.163, C.rule, 90);
}

function slide29(p) {
  const s = p.addSlide();
  blob(s, 'blob14', 2.377,10.735,35.553,6.634, C.sand, 155.74);
  txt(s, 'It is health that is real wealth and not pieces of gold and silver.', 1.284,9.154,18.918,3.311, S.h1);
  txt(s, [{text:'Marcus Spoiler'}, {text:' - Product Manager at Startup', options:S.body}], 1.284,13.06,10.497,0.653, S.bodyBold);
  photoFrame(s, 'frameWide', 0,0,26.667,7.792);
}

function slide30(p) {
  const s = p.addSlide();
  photoFrame(s, 'frameFull', 0,0,15.135,15);
  blob(s, 'blob14', -5.776,12.248,7.664,2.228, C.sand, 314.4);
  blob(s, 'blobHalf', 14.416,0,12.25,15, C.sand, 180);
  txt(s, 'Thank You', 16.527,3.252,7.22,5.861, S.h0);
  txt(s, [{text:'Funding Claster 22843, '}, {text:'', breakLine:true}, {text:'Birmingham, United Kingdom 123456'}], 16.527,10.088,5.856,0.861, S.label);
  txt(s, 'For Further Information ', 16.527,9.567,5.44,0.486, S.labelBold);
  txt(s, 'mail@ifood.com', 16.527,11.262,5.312,0.486, S.label);
  txt(s, 'Ifood © 2022', 1.284,1.269,4.453,0.486, S.labelBoldOnDark);
  blob(s, 'blob8', 0.375,11.551,12.219,3.449, C.sand);
  blob(s, 'blob14', 22.156,10.295,9.163,2.989, C.mist, 93.03);
  blob(s, 'leaf', 24.003,11.502,1.378,1.272, C.mist);
  blob(s, 'leaf', 23.096,14.263,1.596,1.473, C.mist);
  blob(s, 'blob18', 19.304,-2.64,9.299,7.423, C.mist, 180);
}

function slide31(p) {
  const s = p.addSlide();
  txt(s, 'Icons', 1.284,4.778,14.098,5.444, S.displayBlack);
  txt(s, 'Made on 2021', 1.284,13.282,2.862,0.444, {...S.tiny, wrap:false});
}

// ------------------------------------------------- icon library sheets ---
// Slides 32-48 are contact sheets of pictogram icons. Each sheet is rebuilt on
// its original grid with a placeholder glyph in the sheet's own weight:
// "solid" sheets get a blocky filled pictogram, "outline" sheets get the
// circled line-icon motif the template uses.
// rows entries are [centre y, count, first centre x, x pitch].
const ICON_SHEETS = [
  { slide:32, size:1.31, style:'outline', rows:[[2.821,10,3.327,2.221], [5.207,10,3.316,2.226], [7.505,10,3.327,2.221], [9.793,12,3.327,1.817], [12.179,10,3.327,2.221]] },
  { slide:33, size:1.28, style:'outline', rows:[[2.363,10,2.111,2.496], [4.963,10,2.111,2.496], [7.382,10,2.111,2.496], [10.06,10,2.111,2.496], [12.638,12,2.092,2.044]] },
  { slide:34, size:0.67, style:'solid', rows:[[4.106,2,4.725,0.91], [5.356,22,4.606,0.831], [6.197,22,4.606,0.831], [7.019,22,4.606,0.831], [7.838,22,4.612,0.831], [8.921,22,4.607,0.831], [9.754,22,4.609,0.831], [10.585,22,4.607,0.831], [11.415,22,4.611,0.831]] },
  { slide:35, size:1.08, style:'solid', rows:[[3.387,16,3.199,1.351], [5.032,17,3.199,1.267], [6.677,17,3.199,1.267], [8.323,17,3.199,1.267], [9.968,17,3.199,1.267], [11.613,16,3.199,1.351]] },
  { slide:36, size:1.02, style:'solid', rows:[[4.296,17,3.486,1.231], [5.476,17,3.486,1.231], [6.652,16,3.49,1.231], [8.289,17,3.481,1.231], [9.518,17,3.486,1.231], [10.708,16,3.491,1.231]] },
  { slide:37, size:0.98, style:'solid', rows:[[4.296,17,3.699,1.208], [5.503,17,3.698,1.208], [6.712,16,3.697,1.208], [8.288,17,3.7,1.208], [9.495,17,3.698,1.208], [10.214,2,13.327,0], [10.705,16,3.699,1.208]] },
  { slide:38, size:0.97, style:'solid', rows:[[4.347,17,3.66,1.209], [5.534,20,3.662,1.018], [6.724,16,3.662,1.203], [8.329,17,3.66,1.209], [9.519,17,3.661,1.209], [10.708,16,3.661,1.202]] },
  { slide:39, size:1, style:'solid', rows:[[4.311,17,3.882,1.182], [5.478,17,3.876,1.182], [6.629,16,3.875,1.179], [8.328,17,3.872,1.181], [9.495,17,3.872,1.182], [10.644,16,3.872,1.18]] },
  { slide:40, size:0.98, style:'solid', rows:[[4.292,17,3.487,1.231], [5.535,17,3.488,1.231], [6.686,16,3.488,1.23], [8.301,17,3.483,1.231], [9.516,17,3.488,1.231], [10.715,16,3.481,1.231]] },
  { slide:41, size:0.98, style:'solid', rows:[[4.297,17,3.482,1.232], [5.53,17,3.482,1.232], [6.747,16,3.482,1.232], [8.229,17,3.481,1.232], [9.432,17,3.481,1.232], [10.698,16,3.48,1.232]] },
  { slide:42, size:0.99, style:'solid', rows:[[4.303,17,3.554,1.222], [5.529,17,3.557,1.222], [6.736,16,3.557,1.222], [8.266,17,3.554,1.222], [9.486,17,3.557,1.222], [10.704,16,3.557,1.222]] },
  { slide:43, size:1, style:'solid', rows:[[4.256,14,5.231,1.247], [5.446,14,5.231,1.247], [6.829,12,5.055,1.246], [8.274,14,5.231,1.247], [9.507,15,5.231,1.158], [10.745,12,5.112,1.247]] },
  { slide:44, size:0.88, style:'solid', rows:[[3.458,17,4.221,1.139], [4.526,17,4.226,1.139], [5.618,16,4.261,1.136], [6.674,15,4.32,1.088], [8.236,17,4.237,1.139], [9.321,17,4.215,1.14], [10.481,16,4.264,1.137], [11.579,15,4.332,1.088]] },
  { slide:45, size:0.96, style:'solid', rows:[[4.322,17,3.544,1.224], [5.466,17,3.543,1.224], [6.707,16,3.545,1.222], [8.293,17,3.545,1.224], [9.507,17,3.546,1.224], [10.676,16,3.544,1.222]] },
  { slide:46, size:1.29, style:'outline', rows:[[4.599,10,5.146,1.819], [6.407,10,5.114,1.827], [8.217,10,4.973,1.841], [10.023,10,5.003,1.851], [11.83,10,4.856,1.884]] },
  { slide:47, size:1.38, style:'outline', rows:[[3.792,10,5.043,1.878], [5.646,10,5.043,1.842], [7.501,10,4.798,1.869], [9.355,10,4.963,1.865], [11.21,10,5.043,1.856]] },
  { slide:48, size:1.11, style:'outline', rows:[[3.39,10,4.912,1.872], [5.444,10,4.912,1.872], [7.501,10,4.714,1.893], [9.555,10,4.912,1.872], [11.61,10,4.912,1.872]] },
];

function iconGlyph(s, cx, cy, size, style) {
  if (style === 'solid') {
    const d = size * 0.77, t = size * 0.13;
    s.addShape('roundRect', { x: cx - d / 2, y: cy - d / 2, w: d, h: d, rectRadius: d * 0.22,
      line: { color: C.icon, width: t * 72 } });
  } else {
    const d = size * 0.92, t = size * 0.05, dot = size * 0.4;
    s.addShape('ellipse', { x: cx - d / 2, y: cy - d / 2, w: d, h: d,
      line: { color: C.black, width: t * 72 } });
    s.addShape('ellipse', { x: cx - dot / 2, y: cy - dot / 2, w: dot, h: dot,
      fill: { color: C.black }, line: NO_LINE });
  }
}

function iconSheet(s, sheet) {
  sheet.rows.forEach(([cy, count, x0, dx]) => {
    for (let i = 0; i < count; i++) iconGlyph(s, x0 + i * dx, cy, sheet.size, sheet.style);
  });
}

// ------------------------------------------------------------------ main ---
function build() {
  const p = new PptxGenJS();
  p.defineLayout({ name: 'DECK', width: 26.6667, height: 15 });
  p.layout = 'DECK';
  p.author = 'Ifood';
  p.title = 'Ifood - Healthy Food Pitch Deck';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
   slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
   slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24,
   slide25, slide26, slide27, slide28, slide29, slide30, slide31].forEach(fn => fn(p));

  ICON_SHEETS.forEach(sheet => iconSheet(p.addSlide(), sheet));

  return p.writeFile({ fileName: path.join(__dirname, '02daca65-a01e-43de-809b-175618379a28_grok_final.pptx') });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
