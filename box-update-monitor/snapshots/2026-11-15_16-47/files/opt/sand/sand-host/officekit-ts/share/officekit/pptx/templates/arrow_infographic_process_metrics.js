/*
 * Arrow Infographic - a 15-slide deck rebuilt with pptxgenjs.
 * Slide size 10 x 5.625 in (16:9); every measurement below is in inches.
 * The source deck's small raster glyph icons are stood in for by colour blocks.
 */
'use strict';
const path = require('path');
const PptxGenJS = require('pptxgenjs');

const FONT = 'Red Hat Display';

// theme accents, their darker press-in tints, and the neutral text greys
const C = {
  navy    : '00296B',
  blue    : '003F88',
  ocean   : '00509D',
  gold    : 'FDC500',
  navyDk  : '001E50',
  blueDk  : '002F66',
  oceanDk : '003C75',
  goldDk  : 'BD9300',
  navyDp  : '001435',
  blueDp  : '001F44',
  oceanDp : '00284E',
  goldDp  : '7E6200',
  white   : 'FFFFFF',
  ink     : '3F3F3F',
  grey    : '7F7F7F',
  inkDk   : '262626',
  black   : '191919',
  hair    : 'F2F2F2',
  teal    : '0594A4',
  steel   : 'D5DBE5',
  plum    : '5F0F40',
  brick   : 'B80000',
  amber   : 'FF9800',
  sage    : '5F8670',
};

// text-box insets used throughout the source deck (0.075 x 0.037 in, in points)
const INSET = [5.4, 5.4, 2.7, 2.7];

// pptxgenjs surfaces preset-shape adjust handles only via `rectRadius` /
// `angleRange`; these translate the deck's raw OOXML adjust values.
const adj = (v, w, h) => (v / 100000) * Math.min(w, h);
const adj2 = (a, b) => [a / 60000, b / 60000];

// pptxgenjs fills are flat, so the deck's few two-stop gradients are
// flattened to the average of their stops.
const blend = (a, b, t) => [0, 2, 4].map(i =>
  Math.round(parseInt(a.substr(i, 2), 16) * t + parseInt(b.substr(i, 2), 16) * (1 - t))
    .toString(16).padStart(2, '0')).join('').toUpperCase();

// the placeholder copy that repeats across the deck
const L = {
  tempor   : "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor",
  dot      : "Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
  serenity : "A wonderful serenity has taken possession of my entire",
  incid    : "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt",
  comma    : "Lorem ipsum dolor sit amet, consectetur adipiscing elit, ",
  eiusmod  : "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod",
  sit      : "Lorem ipsum dolor sit",
  labore   : "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore.",
  labore2  : "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore",
  amet     : "Lorem ipsum dolor sit amet, ",
  subtitle : "Subtitle Text Goes Here",
};

// the drop shadows of the deck, in order of first appearance
const SH = {
  soft  : { type:'outer', blur:63, offset:43, angle:45, color:'000000', opacity:0.1 },
  lift  : { type:'outer', blur:35, offset:21, angle:46, color:'000000', opacity:0.15 },
  wide  : { type:'outer', blur:40, offset:10, angle:90, color:'000000', opacity:0.15 },
  broad : { type:'outer', blur:51, offset:22, angle:90, color:'221B43', opacity:0.11 },
  drop  : { type:'outer', blur:35, offset:21, angle:46, color:'000000', opacity:0.06 },
  faint : { type:'outer', blur:39, offset:29, angle:45, color:'000000', opacity:0.15 },
  tint  : { type:'outer', blur:14.6, offset:4.5, angle:67, color:'000000', opacity:0.12 },
  close : { type:'outer', blur:80, offset:30, angle:35, color:'000000', opacity:0.09 },
  card  : { type:'outer', blur:26, offset:12, angle:135, color:C.ocean, opacity:0.29 },
};

// Text presets - every paragraph/run formatting combination the deck reuses.
// Names read as <align><pt>[b=bold][n=no-wrap][m/e=middle/bottom][z=no inset][xNN=line spacing];
// colour stays at the call site because it changes from column to column.
const T = {
  c30b     : { align:'center', fontSize:30, bold:true },
  c11bmx13 : { align:'center', valign:'middle', lineSpacingMultiple:1.3, fontSize:11, bold:true },
  c9x13    : { align:'center', lineSpacingMultiple:1.3, fontSize:9 },
  c21bn    : { align:'center', wrap:false, fontSize:21, bold:true },
  c11      : { align:'center', fontSize:11 },
  c24bn    : { align:'center', wrap:false, fontSize:24, bold:true },
  c9x15    : { align:'center', lineSpacingMultiple:1.5, fontSize:9 },
  c14bmx13 : { align:'center', valign:'middle', lineSpacingMultiple:1.3, fontSize:14, bold:true },
  c11x15   : { align:'center', lineSpacingMultiple:1.5, fontSize:11 },
  c21nx15  : { align:'center', lineSpacingMultiple:1.5, wrap:false, fontSize:21 },
  c18bnx13 : { align:'center', lineSpacingMultiple:1.3, wrap:false, fontSize:18, bold:true },
  c27bnx9  : { align:'center', lineSpacingMultiple:0.9, wrap:false, fontSize:27, bold:true },
  c12b     : { align:'center', fontSize:12, bold:true },
  c24bnx13 : { align:'center', lineSpacingMultiple:1.3, wrap:false, fontSize:24, bold:true },
  l9x13    : { lineSpacingMultiple:1.3, fontSize:9 },
  c27bx13  : { align:'center', lineSpacingMultiple:1.3, fontSize:27, bold:true },
  c27bnx13 : { align:'center', lineSpacingMultiple:1.3, wrap:false, fontSize:27, bold:true },
  r9x13    : { align:'right', lineSpacingMultiple:1.3, fontSize:9 },
  l11      : { fontSize:11 },
  l30b     : { fontSize:30, bold:true },
  l11bnx15 : { lineSpacingMultiple:1.5, wrap:false, fontSize:11, bold:true },
  l18bnx15 : { lineSpacingMultiple:1.5, wrap:false, fontSize:18, bold:true },
  l9x15    : { lineSpacingMultiple:1.5, fontSize:9 },
  c18bnmz  : { align:'center', valign:'middle', margin:0, wrap:false, fontSize:18, bold:true },
  c15bnex13: { align:'center', valign:'bottom', lineSpacingMultiple:1.3, wrap:false, fontSize:15, bold:true },
  c15nex13 : { align:'center', valign:'bottom', lineSpacingMultiple:1.3, wrap:false, fontSize:15 },
  c11bnx15 : { align:'center', lineSpacingMultiple:1.5, wrap:false, fontSize:11, bold:true },
  l12b     : { fontSize:12, bold:true },
  l27bn    : { wrap:false, fontSize:27, bold:true },
  c18bnx15 : { align:'center', lineSpacingMultiple:1.5, wrap:false, fontSize:18, bold:true },
  c8x15    : { align:'center', lineSpacingMultiple:1.5, fontSize:8 },
  c14x9    : { align:'center', lineSpacingMultiple:0.9 },
  c9x11    : { align:'center', lineSpacingMultiple:1.1, fontSize:9 },
  c12bn    : { align:'center', wrap:false, fontSize:12, bold:true },
  c30bn    : { align:'center', wrap:false, fontSize:30, bold:true },
  c14      : { align:'center', fontSize:14 },
  r12b     : { align:'right', fontSize:12, bold:true },
  c11nx15  : { align:'center', lineSpacingMultiple:1.5, wrap:false, fontSize:11 },
  c7x15    : { align:'center', lineSpacingMultiple:1.5, fontSize:7 },
  c18bx13  : { align:'center', lineSpacingMultiple:1.3, fontSize:18, bold:true },
  r12bx13  : { align:'right', lineSpacingMultiple:1.3, fontSize:12, bold:true },
  r8x15    : { align:'right', lineSpacingMultiple:1.5, fontSize:8 },
  l12bn    : { wrap:false, fontSize:12, bold:true },
  l27bnx9  : { lineSpacingMultiple:0.9, wrap:false, fontSize:27, bold:true },
  r12bn    : { align:'right', wrap:false, fontSize:12, bold:true },
  r27bnx9  : { align:'right', lineSpacingMultiple:0.9, wrap:false, fontSize:27, bold:true },
  l14n     : { wrap:false, fontSize:14 },
  c27bm    : { align:'center', valign:'middle', fontSize:27, bold:true },
  c11bx8   : { align:'center', lineSpacingMultiple:0.8, fontSize:11, bold:true },
};

// Outlines that have no preset equivalent. Each entry is a path expressed in
// fractions of the shape's own box: M/L = move/line to, C = cubic bezier, Z = close.
const G = {
  arrowTail: [['M', 0.477, 0], ['L', 0, 0.442], ['L', 0.602, 1], ['L', 0.934, 0.692], ['C', 0.934, 0.692, 1.062, 0.541, 0.962, 0.449], ['Z']],
  arrowBody: [['M', 0.902, 0], ['C', 0.902, 0, 0.968, 0.022, 0.966, 0.208], ['C', 0.964, 0.411, 1, 0.569, 1, 0.61], ['C', 1, 0.65, 0.961, 0.691, 0.884, 0.739], ['L', 0.637, 0.894], ['L', 0.71, 0.94], ['C', 0.748, 0.964, 0.735, 1.001, 0.655, 0.998], ['C', 0.574, 0.994, 0.165, 1, 0.1, 1], ['C', 0.036, 1, 0, 0.992, 0, 0.937], ['C', 0, 0.882, 0.009, 0.619, 0.009, 0.576], ['C', 0.009, 0.533, 0.068, 0.541, 0.098, 0.56], ['L', 0.161, 0.6], ['L', 0.538, 0.362], ['L', 0.851, 0.166], ['C', 0.924, 0.12, 0.931, 0.021, 0.902, 0], ['Z']],
  ribbonFold: [['M', 0.488, 0.001], ['C', 0.576, -0.013, 0.668, 0.148, 0.668, 0.148], ['L', 1, 0.897], ['L', 0.954, 1], ['L', 0, 1], ['L', 0.405, 0.086], ['C', 0.43, 0.03, 0.459, 0.006, 0.488, 0.001], ['Z']],
  shieldUp: [['M', 0.419, 0.011], ['L', 0.037, 0.142], ['C', 0.014, 0.15, 0.001, 0.161, 0.001, 0.173], ['L', 0.001, 0.958], ['C', 0, 0.981, 0.053, 1, 0.118, 1], ['C', 0.118, 1, 0.118, 1, 0.118, 1], ['L', 0.883, 1], ['C', 0.948, 1, 1, 0.981, 1, 0.958], ['L', 1, 0.173], ['C', 1, 0.161, 0.987, 0.15, 0.964, 0.142], ['L', 0.582, 0.011], ['C', 0.536, -0.004, 0.465, -0.004, 0.419, 0.011], ['Z']],
  starBadge: [['M', 1, 0.328], ['L', 0.774, 1], ['L', 0.1, 0.868], ['L', 0.224, 0.793], ['L', 0, 0.391], ['L', 0.651, 0], ['L', 0.875, 0.402], ['L', 1, 0.328], ['Z']],
  stepTread: [['M', 0.599, 0.998], ['L', -0.001, 0.998], ['L', -0.001, 0.408], ['C', -0.001, 0.182, 0.05, -0.002, 0.112, -0.002], ['L', 1, -0.002], ['L', 0.988, 0.09], ['C', 0.918, 0.644, 0.766, 0.999, 0.599, 0.998], ['Z']],
  stepRiser: [['M', 0.628, 0], ['C', 0.639, 0, 0.651, 0.004, 0.66, 0.011], ['C', 0.66, 0.011, 0.66, 0.011, 0.66, 0.011], ['L', 0.987, 0.285], ['C', 1.005, 0.299, 1.004, 0.323, 0.986, 0.337], ['C', 0.978, 0.344, 0.967, 0.347, 0.955, 0.347], ['L', 0.883, 0.347], ['L', 0.728, 0.621], ['C', 0.596, 0.852, 0.313, 1, 0.001, 1], ['L', 0, 1], ['L', 0.372, 0.347], ['L', 0.3, 0.347], ['C', 0.276, 0.347, 0.255, 0.331, 0.255, 0.31], ['C', 0.255, 0.301, 0.26, 0.292, 0.268, 0.285], ['L', 0.596, 0.011], ['C', 0.605, 0.004, 0.616, 0, 0.628, 0], ['Z']],
  waveUpA: [['M', 0.91, 1], ['L', 1, 0.821], ['L', 0.969, 0.821], ['L', 0.961, 0.68], ['C', 0.916, 0.293, 0.721, 0.001, 0.487, 0], ['C', 0.219, -0.001, 0.002, 0.378, 0, 0.847], ['L', 0.113, 0.849], ['C', 0.114, 0.488, 0.281, 0.198, 0.486, 0.199], ['C', 0.666, 0.2, 0.816, 0.424, 0.85, 0.721], ['L', 0.856, 0.821], ['L', 0.821, 0.821], ['L', 0.91, 1], ['Z']],
  waveDnA: [['M', 0.91, 0], ['L', 1, 0.179], ['L', 0.969, 0.179], ['L', 0.961, 0.32], ['C', 0.916, 0.707, 0.721, 0.999, 0.487, 1], ['C', 0.219, 1.001, 0.002, 0.622, 0, 0.153], ['L', 0.113, 0.151], ['C', 0.114, 0.511, 0.281, 0.802, 0.486, 0.801], ['C', 0.666, 0.8, 0.816, 0.576, 0.85, 0.279], ['L', 0.856, 0.179], ['L', 0.821, 0.179], ['L', 0.91, 0], ['Z']],
  waveUpB: [['M', 0.91, 1], ['L', 1, 0.821], ['L', 0.969, 0.821], ['L', 0.961, 0.68], ['C', 0.916, 0.293, 0.721, 0.001, 0.487, 0], ['C', 0.219, -0.001, 0.002, 0.378, 0, 0.847], ['L', 0.113, 0.849], ['C', 0.114, 0.488, 0.281, 0.198, 0.486, 0.199], ['C', 0.666, 0.2, 0.816, 0.424, 0.85, 0.721], ['L', 0.856, 0.821], ['L', 0.821, 0.821], ['L', 0.91, 1], ['Z']],
  waveDnB: [['M', 0.91, 0], ['L', 1, 0.179], ['L', 0.969, 0.179], ['L', 0.961, 0.32], ['C', 0.916, 0.707, 0.721, 0.999, 0.487, 1], ['C', 0.219, 1.001, 0.002, 0.622, 0, 0.153], ['L', 0.113, 0.151], ['C', 0.114, 0.512, 0.281, 0.802, 0.486, 0.801], ['C', 0.666, 0.8, 0.816, 0.576, 0.85, 0.279], ['L', 0.856, 0.179], ['L', 0.821, 0.179], ['L', 0.91, 0], ['Z']],
  discLowerRight: [['M', 0.845, 0.765], ['L', 0.841, 0.525], ['C', 0.788, 0.547, 0.733, 0.554, 0.677, 0.545], ['C', 0.575, 0.529, 0.484, 0.462, 0.42, 0.357], ['C', 0.378, 0.286, 0.351, 0.204, 0.342, 0.117], ['L', 0.178, 0], ['L', 0.178, 0], ['L', 0, 0.097], ['C', 0.003, 0.168, 0.012, 0.239, 0.027, 0.308], ['C', 0.053, 0.426, 0.096, 0.534, 0.154, 0.63], ['C', 0.212, 0.727, 0.283, 0.807, 0.364, 0.867], ['C', 0.448, 0.93, 0.54, 0.97, 0.637, 0.985], ['C', 0.734, 1, 0.83, 0.99, 0.923, 0.955], ['C', 0.95, 0.945, 0.975, 0.933, 1, 0.919], ['L', 0.845, 0.765]],
  discUpperLeft: [['M', 0.172, 0.864], ['L', 0.327, 1], ['C', 0.344, 0.865, 0.391, 0.745, 0.459, 0.659], ['C', 0.536, 0.561, 0.633, 0.517, 0.73, 0.536], ['C', 0.763, 0.542, 0.794, 0.555, 0.824, 0.574], ['L', 0.976, 0.425], ['L', 0.976, 0.425], ['L', 1, 0.13], ['C', 0.928, 0.071, 0.85, 0.033, 0.769, 0.018], ['C', 0.676, 0, 0.584, 0.012, 0.495, 0.053], ['C', 0.408, 0.093, 0.329, 0.159, 0.258, 0.249], ['C', 0.188, 0.338, 0.129, 0.447, 0.085, 0.572], ['C', 0.04, 0.698, 0.012, 0.835, 0, 0.98], ['L', 0.172, 0.864]],
  discRight: [['M', 0.363, 0.164], ['L', 0.042, 0.245], ['C', 0.074, 0.257, 0.104, 0.271, 0.132, 0.285], ['C', 0.253, 0.349, 0.311, 0.427, 0.299, 0.507], ['C', 0.298, 0.513, 0.297, 0.519, 0.295, 0.525], ['C', 0.27, 0.612, 0.166, 0.689, 0.001, 0.743], ['C', 0, 0.743, 0, 0.743, 0, 0.743], ['L', 0.009, 0.901], ['L', 0.009, 0.901], ['L', 0.322, 1], ['C', 0.358, 0.99, 0.394, 0.98, 0.428, 0.968], ['C', 0.579, 0.919, 0.703, 0.859, 0.798, 0.79], ['C', 0.896, 0.719, 0.958, 0.641, 0.981, 0.559], ['C', 0.985, 0.548, 0.987, 0.536, 0.989, 0.525], ['C', 1, 0.454, 0.981, 0.384, 0.934, 0.317], ['C', 0.881, 0.24, 0.794, 0.17, 0.675, 0.107], ['C', 0.599, 0.067, 0.511, 0.031, 0.413, 0], ['L', 0.363, 0.164]],
  pinBanner: [['M', 0.349, 0], ['L', 0.941, 0], ['L', 0.784, 0.177], ['L', 0.784, 0.596], ['L', 0.939, 0.596], ['C', 0.973, 0.596, 1, 0.616, 1, 0.639], ['C', 1, 0.65, 0.994, 0.661, 0.983, 0.669], ['L', 0.543, 0.987], ['C', 0.52, 1.004, 0.482, 1.004, 0.458, 0.988], ['C', 0.457, 0.988, 0.457, 0.987, 0.457, 0.987], ['L', 0.017, 0.669], ['C', -0.006, 0.652, -0.006, 0.625, 0.018, 0.609], ['C', 0.03, 0.601, 0.045, 0.596, 0.061, 0.596], ['L', 0.192, 0.596], ['L', 0.192, 0.111], ['C', 0.192, 0.05, 0.262, 0, 0.349, 0], ['Z']],
  pinCap: [['M', 0.498, -0.001], ['L', 0.498, -0.001], ['C', 0.775, -0.001, 0.998, 0.122, 0.998, 0.275], ['L', 0.998, 0.723], ['C', 0.998, 0.875, 0.775, 0.999, 0.498, 0.999], ['L', -0.002, 0.999], ['L', -0.002, 0.274], ['C', 0, 0.122, 0.223, -0.001, 0.498, -0.001], ['Z']],
  tagDown: [['M', 0.98, 0.43], ['L', 0.527, 0.985], ['C', 0.511, 1.005, 0.49, 1.005, 0.474, 0.985], ['L', 0.02, 0.43], ['C', 0.007, 0.415, 0, 0.39, -0.001, 0.362], ['L', -0.001, 0], ['L', 1, 0], ['L', 1, 0.362], ['C', 1, 0.389, 0.993, 0.415, 0.98, 0.43], ['Z']],
  slantL: [['M', 0.6, 1], ['L', 0, 1], ['L', 0.611, 0], ['L', 1, 0], ['L', 0.6, 1], ['Z']],
  tagDown2: [['M', 0.98, 0.43], ['L', 0.526, 0.985], ['C', 0.51, 1.005, 0.489, 1.005, 0.473, 0.985], ['L', 0.02, 0.43], ['C', 0.007, 0.415, -0.001, 0.39, -0.001, 0.362], ['L', -0.001, 0], ['L', 1, 0], ['L', 1, 0.362], ['C', 1, 0.389, 0.993, 0.415, 0.98, 0.43], ['Z']],
  slantWide: [['M', 0.905, 1], ['L', 0, 1], ['L', 0.413, 0], ['L', 1, 0], ['L', 0.905, 1], ['Z']],
  slantR: [['M', 1, 1], ['L', 0.095, 1], ['L', 0, 0], ['L', 0.587, 0], ['L', 1, 1], ['Z']],
  slantNarrow: [['M', 1, 1], ['L', 0.4, 1], ['L', 0, 0], ['L', 0.389, 0], ['L', 1, 1], ['Z']],
  chevUpNotch: [['M', 0.973, 0.999], ['L', 0.752, 0.999], ['C', 0.746, 0.999, 0.739, 0.994, 0.734, 0.985], ['L', 0.526, 0.582], ['C', 0.512, 0.553, 0.489, 0.553, 0.475, 0.582], ['L', 0.267, 0.985], ['C', 0.262, 0.994, 0.255, 0.999, 0.248, 0.999], ['L', 0.027, 0.999], ['C', 0.003, 0.999, -0.009, 0.933, 0.01, 0.897], ['L', 0.453, 0.039], ['C', 0.48, -0.014, 0.522, -0.014, 0.549, 0.039], ['L', 0.992, 0.897], ['C', 1.01, 0.933, 0.998, 0.999, 0.973, 0.999], ['Z']],
  arrowGlyph: [['M', 0.001, 0.973], ['L', 0.001, 0.752], ['C', 0.001, 0.745, 0.006, 0.739, 0.016, 0.734], ['L', 0.419, 0.526], ['C', 0.447, 0.511, 0.447, 0.488, 0.419, 0.474], ['L', 0.016, 0.266], ['C', 0.006, 0.261, 0.001, 0.255, 0.001, 0.248], ['L', 0.001, 0.027], ['C', 0.001, 0.002, 0.067, -0.01, 0.103, 0.009], ['L', 0.961, 0.452], ['C', 1.014, 0.479, 1.014, 0.521, 0.961, 0.548], ['L', 0.103, 0.991], ['C', 0.067, 1.009, 0.001, 0.998, 0.001, 0.973], ['Z']],
};

// ---------------------------------------------------------------- helpers
// Every shape and text box is placed with a box literal [x, y, w, h] in inches,
// optionally followed by a rotation in degrees.  Note that pptxgenjs rewrites `shadow` in place while serialising, so the shared SH
// `shadow` in place while serialising, so a shared SH preset must be copied.
function opts(box, o) {
  o = Object.assign({ x: box[0], y: box[1], w: box[2], h: box[3], rotate: box[4] || 0 }, o);
  if (typeof o.fill === 'string') o.fill = { color: o.fill };
  if (o.shadow) o.shadow = Object.assign({}, o.shadow);
  return o;
}

const sh = (s, type, box, o) => s.addShape(type, opts(box, o));

function tx(s, text, box, style, o) {
  s.addText(text, Object.assign({ fontFace: FONT, valign: 'top', margin: INSET },
    style, opts(box, o)));
}

// a G[] outline scaled into its own box
function poly(s, pts, box, o) {
  o = opts(box, o);
  o.points = pts.map(p =>
    p[0] === 'Z' ? { close: true }
    : p[0] === 'M' ? { x: p[1] * o.w, y: p[2] * o.h, moveTo: true }
    : p[0] === 'L' ? { x: p[1] * o.w, y: p[2] * o.h }
    : { x: p[5] * o.w, y: p[6] * o.h,
        curve: { type: 'cubic', x1: p[1] * o.w, y1: p[2] * o.h, x2: p[3] * o.w, y2: p[4] * o.h } });
  s.addShape('custGeom', o);
}

// stand-in for one of the deck's raster glyphs: a rounded block in the
// glyph's own colour, occupying the same box.
function icon(s, box, color) {
  const pad = 0.09 * Math.min(box[2], box[3]);
  sh(s, 'roundRect', [box[0] + pad, box[1] + pad, box[2] - 2 * pad, box[3] - 2 * pad, box[4]],
    { fill: color, rectRadius: Math.min(box[2], box[3]) * 0.2 });
}

// the slide master puts a web address and the page number on every slide
function footer(s, n) {
  tx(s, 'www.website.com', [0.5, 5.062, 2.135, 0.202], { fontSize: 8, color: C.navy });
  tx(s, String(n), [9.133, 5.075, 0.366, 0.202],
    { fontSize: 8, bold: true, align: 'center', valign: 'middle', color: C.navy });
}

// ------------------------------------------------------------- the slides

// 1. zig-zag arrow band with four step labels and four KPI columns above
function slide1(s, n) {
  tx(s, "Arrow Infographic", [2.6, 0.562, 4.799, 0.581], T.c30b, { color:C.ink });
  sh(s, 'rightArrow', [6.594, 3.743, 2.141, 0.968, 225], { flipH:true, fill:C.gold, shadow:SH.soft, angleRange:adj2(62489, 50000) });
  sh(s, 'rightArrow', [1.266, 3.263, 2.141, 0.968, 45], { flipH:true, fill:C.navy, shadow:SH.soft, angleRange:adj2(62489, 50000) });
  sh(s, 'roundRect', [4.563, 3.639, 0.782, 0.815, 45], { flipH:true, fill:C.blueDk, rectRadius:adj(16667, 0.782, 0.815) });
  sh(s, 'roundRect', [2.943, 3.639, 0.782, 0.815, 45], { flipH:true, fill:C.navyDk, rectRadius:adj(16667, 0.782, 0.815) });
  sh(s, 'round2DiagRect', [2.279, 3.037, 0.61, 1.926, 135], { flipH:true, fill:C.navy, angleRange:adj2(34656, 0) });
  sh(s, 'round2DiagRect', [3.896, 3.037, 0.61, 1.926, 135], { flipH:true, fill:C.blue, angleRange:adj2(34656, 0) });
  sh(s, 'roundRect', [6.18, 3.639, 0.782, 0.815, 45], { flipH:true, fill:C.oceanDk, rectRadius:adj(16667, 0.782, 0.815) });
  sh(s, 'round2DiagRect', [5.515, 3.037, 0.61, 1.926, 135], { flipH:true, fill:C.ocean, angleRange:adj2(34656, 0) });
  sh(s, 'round2DiagRect', [7.133, 3.037, 0.61, 1.926, 135], { flipH:true, fill:C.gold, angleRange:adj2(34656, 0) });
  tx(s, "Step One", [2.241, 3.782, 0.887, 0.547, 45], T.c11bmx13, { color:C.white });
  tx(s, "Step Two", [3.861, 3.782, 0.887, 0.547, 45], T.c11bmx13, { color:C.white });
  tx(s, "Step Three", [5.367, 3.782, 1.131, 0.547, 45], T.c11bmx13, { color:C.white });
  tx(s, "Step Four", [7.094, 3.782, 0.887, 0.547, 45], T.c11bmx13, { color:C.white });
  tx(s, L.comma, [1.237, 2.344, 1.801, 0.454], T.c9x13, { color:C.grey });
  tx(s, "+52,000", [1.237, 1.906, 1.801, 0.429], T.c21bn, { color:C.navy });
  tx(s, L.comma, [3.145, 2.344, 1.801, 0.454], T.c9x13, { color:C.grey });
  tx(s, "+32,000", [3.145, 1.906, 1.801, 0.429], T.c21bn, { color:C.blue });
  tx(s, L.comma, [6.962, 2.344, 1.801, 0.454], T.c9x13, { color:C.grey });
  tx(s, "+12,000", [6.962, 1.906, 1.801, 0.429], T.c21bn, { color:C.gold });
  tx(s, L.comma, [5.053, 2.344, 1.801, 0.454], T.c9x13, { color:C.grey });
  tx(s, "+22,000", [5.053, 1.906, 1.801, 0.429], T.c21bn, { color:C.ocean });
  icon(s, [6.841, 3.363, 0.351, 0.351, 45], C.white);
  icon(s, [1.983, 3.39, 0.351, 0.351, 45], C.white);
  icon(s, [5.2, 3.39, 0.351, 0.351, 45], C.white);
  icon(s, [3.571, 3.395, 0.351, 0.351, 45], C.white);
  tx(s, L.subtitle, [2.681, 1.139, 4.638, 0.252], T.c11, { color:C.grey });
  footer(s, n);
}

// 2. four bent-arrow + circle-icon units, step title above, KPI block below
function slide2(s, n) {
  tx(s, "Arrow Infographic", [2.6, 0.576, 4.799, 0.581], T.c30b, { color:C.ink });
  tx(s, "+90M", [1.073, 3.603, 1.211, 0.48], T.c24bn, { color:C.navy });
  tx(s, "2500+ employee", [0.933, 3.988, 1.491, 0.303], T.c9x15, { color:C.grey });
  tx(s, L.serenity, [0.608, 4.247, 2.141, 0.53], T.c9x15, { color:C.grey });
  tx(s, "+80M", [3.29, 3.603, 1.211, 0.48], T.c24bn, { color:C.blue });
  tx(s, "2500+ employee", [3.15, 3.988, 1.491, 0.303], T.c9x15, { color:C.grey });
  tx(s, L.serenity, [2.825, 4.247, 2.141, 0.53], T.c9x15, { color:C.grey });
  tx(s, "+70M", [5.508, 3.603, 1.211, 0.48], T.c24bn, { color:C.ocean });
  tx(s, "2500+ employee", [5.367, 3.988, 1.491, 0.303], T.c9x15, { color:C.grey });
  tx(s, L.serenity, [5.043, 4.247, 2.141, 0.53], T.c9x15, { color:C.grey });
  tx(s, "+60M", [7.725, 3.603, 1.211, 0.48], T.c24bn, { color:C.gold });
  tx(s, "2500+ employee", [7.585, 3.988, 1.491, 0.303], T.c9x15, { color:C.grey });
  tx(s, L.serenity, [7.26, 4.247, 2.141, 0.53], T.c9x15, { color:C.grey });
  poly(s, G.arrowTail, [0.78, 2.451, 0.864, 0.933, 270], { flipH:true, fill:{ color:blend(C.navyDk, C.navyDp, 0.62) } });
  poly(s, G.arrowBody, [1.432, 2.276, 0.876, 1.394, 270], { flipH:true, fill:C.navy, shadow:SH.lift });
  sh(s, 'ellipse', [1.427, 2.078, 0.502, 0.502], { fill:C.navy, shadow:SH.wide });
  icon(s, [1.524, 2.175, 0.307, 0.307], C.white);
  tx(s, "Step One", [1.073, 1.473, 1.211, 0.547], T.c14bmx13, { color:C.navy });
  poly(s, G.arrowTail, [2.997, 2.512, 0.864, 0.933, 270], { fill:{ color:blend(C.blueDk, C.blueDp, 0.62) } });
  poly(s, G.arrowBody, [3.649, 2.227, 0.876, 1.394, 270], { fill:C.blue, shadow:SH.lift });
  sh(s, 'ellipse', [3.641, 2.078, 0.502, 0.502], { fill:C.blue, shadow:SH.wide });
  icon(s, [3.739, 2.179, 0.307, 0.307], C.white);
  tx(s, "Step Two", [3.29, 1.473, 1.211, 0.547], T.c14bmx13, { color:C.blue });
  poly(s, G.arrowTail, [5.215, 2.451, 0.864, 0.933, 270], { flipH:true, fill:{ color:blend(C.oceanDk, C.oceanDp, 0.62) } });
  poly(s, G.arrowBody, [5.867, 2.276, 0.876, 1.394, 270], { flipH:true, fill:C.ocean, shadow:SH.lift });
  sh(s, 'ellipse', [5.862, 2.078, 0.502, 0.502], { fill:C.ocean, shadow:SH.wide });
  icon(s, [5.959, 2.179, 0.307, 0.307], C.white);
  tx(s, "Step Three", [5.341, 1.473, 1.543, 0.547], T.c14bmx13, { color:C.ocean });
  poly(s, G.arrowTail, [7.432, 2.512, 0.864, 0.933, 270], { fill:{ color:blend(C.goldDk, C.goldDp, 0.62) } });
  poly(s, G.arrowBody, [8.084, 2.227, 0.876, 1.394, 270], { fill:C.gold, shadow:SH.lift });
  sh(s, 'ellipse', [8.079, 2.078, 0.502, 0.502], { fill:C.gold, shadow:SH.wide });
  icon(s, [8.177, 2.179, 0.307, 0.307], C.white);
  tx(s, "Step Four", [7.725, 1.473, 1.211, 0.547], T.c14bmx13, { color:C.gold });
  tx(s, L.subtitle, [2.681, 1.139, 4.638, 0.252], T.c11, { color:C.grey });
  footer(s, n);
}

// 3. centre ribbon flowing between two white "2024" cards
function slide3(s, n) {
  sh(s, 'roundRect', [7.361, 2.772, 1.928, 1.92], { fill:C.white, shadow:SH.broad, rectRadius:adj(5529, 1.928, 1.92) });
  sh(s, 'roundRect', [0.717, 2.773, 1.928, 1.92], { fill:C.white, shadow:SH.broad, rectRadius:adj(5529, 1.928, 1.92) });
  poly(s, G.arrowBody, [5.331, 1.821, 1.445, 2.299, 270], { fill:C.gold, shadow:SH.lift });
  poly(s, G.ribbonFold, [4.312, 2.348, 1.425, 0.632], { fill:C.goldDk });
  poly(s, G.arrowBody, [3.23, 1.821, 1.445, 2.299, 90], { fill:C.navy, shadow:SH.lift });
  poly(s, G.ribbonFold, [4.269, 2.96, 1.425, 0.632, 180], { fill:C.navyDp });
  tx(s, L.subtitle, [2.681, 1.259, 4.638, 0.252], T.c11, { color:C.grey });
  tx(s, "Arrow Infographic", [2.6, 0.695, 4.799, 0.581], T.c30b, { color:C.ink });
  tx(s, L.dot, [7.365, 3.959, 1.919, 0.53], T.c9x15, { color:C.grey });
  tx(s, "Your text here", [7.365, 3.682, 1.919, 0.341], T.c11x15, { color:C.ink });
  tx(s, "2024", [7.365, 3.201, 1.919, 0.606], T.c21nx15, { color:C.ink });
  sh(s, 'downArrow', [8.183, 2.975, 0.282, 0.308], { fill:C.gold });
  tx(s, L.dot, [0.721, 3.96, 1.919, 0.53], T.c9x15, { color:C.grey });
  tx(s, "Your text here", [0.721, 3.683, 1.919, 0.341], T.c11x15, { color:C.ink });
  tx(s, "2024", [0.721, 3.202, 1.919, 0.606], T.c21nx15, { color:C.ink });
  sh(s, 'downArrow', [1.54, 2.976, 0.282, 0.308, 180], { fill:C.blue });
  tx(s, "680+", [2.807, 2.444, 1.046, 0.441], T.c18bnx13, { color:C.white });
  tx(s, "-550", [6.205, 2.997, 1.046, 0.441], T.c18bnx13, { color:C.white });
  tx(s, "$35M", [0.747, 1.835, 1.865, 0.485], T.c27bnx9, { color:C.blue });
  tx(s, "Lorem Ipsum", [0.747, 2.267, 1.865, 0.278], T.c12b, { color:C.inkDk });
  tx(s, "$35M", [7.392, 1.835, 1.865, 0.485], T.c27bnx9, { color:C.gold });
  tx(s, "Lorem Ipsum", [7.392, 2.267, 1.865, 0.278], T.c12b, { color:C.inkDk });
  tx(s, L.labore, [3.232, 4.224, 3.542, 0.53], T.c9x15, { color:C.grey });
  footer(s, n);
}

// 4. four horizontal arrow bars pointing at a grey hub circle
function slide4(s, n) {
  poly(s, G.shieldUp, [2.112, 0.714, 1.336, 3.336, 90], { fill:C.navy, shadow:SH.drop });
  poly(s, G.starBadge, [3.394, 1.818, 1.213, 1.169, 300], { fill:C.navyDk, shadow:SH.soft });
  tx(s, "01", [3.693, 2.13, 0.614, 0.563], T.c24bnx13, { color:C.white, flipH:true });
  tx(s, L.eiusmod, [2.34, 1.83, 1.101, 1.044], T.l9x13, { color:C.white });
  tx(s, [{ text:"6" }, { text:"9%" }], [1.203, 2.044, 1.101, 0.624], T.c27bx13, { color:C.white });
  poly(s, G.shieldUp, [2.112, 2.349, 1.336, 3.336, 90], { fill:C.blue, shadow:SH.drop });
  poly(s, G.starBadge, [3.394, 3.453, 1.213, 1.169, 300], { fill:C.blueDk, shadow:SH.soft });
  tx(s, "04", [3.693, 3.765, 0.614, 0.563], T.c24bnx13, { color:C.white, flipH:true });
  tx(s, L.eiusmod, [2.34, 3.465, 1.101, 1.044], T.l9x13, { color:C.white });
  tx(s, "84%", [1.203, 3.679, 1.101, 0.624], T.c27bnx13, { color:C.white });
  poly(s, G.shieldUp, [6.552, 0.714, 1.336, 3.336, 270], { flipH:true, fill:C.ocean, shadow:SH.drop });
  poly(s, G.starBadge, [5.393, 1.818, 1.213, 1.169, 60], { flipH:true, fill:C.oceanDk, shadow:SH.soft });
  tx(s, "02", [5.693, 2.13, 0.614, 0.563], T.c24bnx13, { color:C.white });
  tx(s, L.eiusmod, [6.559, 1.83, 1.101, 1.044], T.r9x13, { color:C.white, flipH:true });
  tx(s, "58%", [7.696, 2.044, 1.101, 0.624], T.c27bnx13, { color:C.white, flipH:true });
  poly(s, G.shieldUp, [6.552, 2.349, 1.336, 3.336, 270], { flipH:true, fill:C.gold, shadow:SH.drop });
  poly(s, G.starBadge, [5.393, 3.453, 1.213, 1.169, 60], { flipH:true, fill:C.goldDk, shadow:SH.soft });
  tx(s, "03", [5.693, 3.765, 0.614, 0.563], T.c24bnx13, { color:C.white });
  tx(s, L.eiusmod, [6.559, 3.465, 1.101, 1.044], T.r9x13, { color:C.white, flipH:true });
  tx(s, [{ text:"97" }, { text:"%" }], [7.696, 3.679, 1.101, 0.624], T.c27bx13, { color:C.white, flipH:true });
  tx(s, L.subtitle, [2.681, 1.133, 4.638, 0.252], T.c11, { color:C.grey });
  tx(s, "Arrow Infographic", [2.6, 0.57, 4.799, 0.581], T.c30b, { color:C.ink });
  sh(s, 'ellipse', [4.404, 2.654, 1.193, 1.193], { fill:C.hair });
  footer(s, n);
}

// 5. ascending staircase of arrows, each with value + caption
function slide5(s, n) {
  poly(s, G.stepTread, [6.637, 1.736, 2.106, 0.576], { fill:C.teal });
  poly(s, G.stepTread, [6.637, 1.736, 2.106, 0.576], { fill:C.teal });
  poly(s, G.stepTread, [6.637, 1.736, 2.106, 0.576], { fill:C.navyDk });
  poly(s, G.stepRiser, [7.899, 0.93, 1.126, 1.381], { fill:C.navy });
  poly(s, G.stepTread, [4.718, 2.636, 2.106, 0.576], { fill:C.teal });
  poly(s, G.stepTread, [4.718, 2.636, 2.106, 0.576], { fill:C.teal });
  poly(s, G.stepTread, [4.718, 2.636, 2.106, 0.576], { fill:C.blueDk });
  poly(s, G.stepRiser, [5.98, 1.83, 1.126, 1.381], { fill:C.blue });
  poly(s, G.stepTread, [2.798, 3.538, 2.106, 0.576], { fill:C.teal });
  poly(s, G.stepTread, [2.798, 3.538, 2.106, 0.576], { fill:C.teal });
  poly(s, G.stepTread, [2.798, 3.538, 2.106, 0.576], { fill:C.oceanDk });
  poly(s, G.stepRiser, [4.06, 2.732, 1.126, 1.381], { fill:C.ocean });
  poly(s, G.stepTread, [0.879, 4.438, 2.106, 0.576], { fill:C.teal });
  poly(s, G.stepTread, [0.879, 4.438, 2.106, 0.576], { fill:C.teal });
  poly(s, G.stepTread, [0.879, 4.438, 2.106, 0.576], { fill:C.goldDk });
  poly(s, G.stepRiser, [2.141, 3.633, 1.126, 1.381], { fill:C.gold });
  tx(s, L.subtitle, [0.512, 1.145, 4.638, 0.252], T.l11, { color:C.grey });
  tx(s, "Arrow Infographic", [0.512, 0.582, 4.799, 0.581], T.l30b, { color:C.ink });
  tx(s, "Product", [1.093, 4.536, 1.279, 0.316], T.l11bnx15, { color:C.white });
  tx(s, "$1,300", [0.879, 3.313, 1.05, 0.486], T.l18bnx15, { color:C.gold });
  tx(s, [{ text:"Lorem ipsum dolor", options:{ breakLine:true } }, { text:"sit amet, consectetur" }], [0.879, 3.782, 1.395, 0.508], T.l9x15, { color:C.grey });
  tx(s, "Sales", [7.16, 1.838, 1.279, 0.316], T.l11bnx15, { color:C.white });
  tx(s, "$5,500", [6.771, 0.625, 1.096, 0.486], T.l18bnx15, { color:C.navy });
  tx(s, [{ text:"Lorem ipsum dolor", options:{ breakLine:true } }, { text:"sit amet, consectetur" }], [6.771, 1.094, 1.545, 0.508], T.l9x15, { color:C.grey });
  tx(s, "Income", [5.176, 2.738, 1.279, 0.316], T.l11bnx15, { color:C.white });
  tx(s, "$4,500", [4.774, 1.502, 1.026, 0.482], T.l18bnx15, { color:C.blue });
  tx(s, [{ text:"Lorem ipsum dolor", options:{ breakLine:true } }, { text:"sit amet, consectetur" }], [4.774, 1.971, 1.395, 0.508], T.l9x15, { color:C.grey });
  tx(s, "Market", [3.305, 3.629, 1.279, 0.316], T.l11bnx15, { color:C.white });
  tx(s, "$2,200", [2.831, 2.427, 1.05, 0.486], T.l18bnx15, { color:C.ocean });
  tx(s, [{ text:"Lorem ipsum dolor", options:{ breakLine:true } }, { text:"sit amet, consectetur" }], [2.831, 2.896, 1.395, 0.508], T.l9x15, { color:C.grey });
  sh(s, 'roundRect', [2.636, 4.382, 0.631, 0.631], { fill:C.white, shadow:SH.faint, rectRadius:adj(40319, 0.631, 0.631) });
  sh(s, 'roundRect', [4.541, 3.483, 0.631, 0.631], { fill:C.white, shadow:SH.faint, rectRadius:adj(40319, 0.631, 0.631) });
  sh(s, 'roundRect', [6.447, 2.583, 0.631, 0.631], { fill:C.white, shadow:SH.faint, rectRadius:adj(40319, 0.631, 0.631) });
  sh(s, 'roundRect', [8.353, 1.683, 0.631, 0.631], { fill:C.white, shadow:SH.faint, rectRadius:adj(40319, 0.631, 0.631) });
  icon(s, [8.441, 1.771, 0.455, 0.455], C.brick);
  icon(s, [4.629, 3.571, 0.455, 0.455], C.sage);
  icon(s, [2.724, 4.47, 0.455, 0.455], C.plum);
  icon(s, [6.535, 2.671, 0.455, 0.455], C.amber);
  tx(s, L.dot, [3.354, 4.407, 1.919, 0.53], T.l9x15, { color:C.grey });
  tx(s, L.dot, [5.24, 3.517, 1.919, 0.53], T.l9x15, { color:C.grey });
  tx(s, L.dot, [7.126, 2.627, 1.919, 0.53], T.l9x15, { color:C.grey });
  footer(s, n);
}

// 6. four revenue cards on a stem above chevrons on a horizon line
function slide6(s, n) {
  tx(s, "Arrow Infographic", [2.6, 0.569, 4.799, 0.581], T.c30b, { color:C.ink });
  sh(s, 'line', [0, 4.182, 10.001, 0], { line:{ color:C.hair, width:7 } });
  sh(s, 'chevron', [1.803, 3.42, 1.383, 1.494], { fill:C.navy, shadow:SH.wide, rectRadius:adj(18657, 1.383, 1.494) });
  sh(s, 'ellipse', [2.192, 2.944, 0.672, 0.672], { fill:C.white, line:{ color:C.hair, width:0.75 }, shadow:SH.tint });
  tx(s, "01", [2.318, 3.113, 0.388, 0.345], T.c18bnmz, { color:C.navy });
  sh(s, 'line', [2.528, 2.659, 0, 0.284, 180], { fill:'0D67FF', line:{ color:C.inkDk, width:2, dashType:'dash', endArrowType:'oval' } });
  sh(s, 'roundRect', [1.842, 1.48, 1.372, 1.075], { fill:C.navy, shadow:SH.close, rectRadius:adj(5562, 1.372, 1.075) });
  tx(s, "40,214", [1.896, 1.826, 1.263, 0.404], T.c15bnex13, { color:C.white });
  tx(s, "Your Value", [1.835, 2.18, 1.386, 0.273], T.c9x13, { color:C.white });
  tx(s, "Revenue", [1.895, 1.538, 1.263, 0.404], T.c15nex13, { color:C.white });
  sh(s, 'chevron', [3.477, 3.409, 1.383, 1.494], { fill:C.blue, shadow:SH.wide, rectRadius:adj(18657, 1.383, 1.494) });
  sh(s, 'ellipse', [3.841, 2.955, 0.672, 0.672], { fill:C.white, line:{ color:C.hair, width:0.75 }, shadow:SH.tint });
  tx(s, "02", [3.982, 3.118, 0.388, 0.345], T.c18bnmz, { color:C.blue });
  sh(s, 'line', [4.177, 2.659, 0, 0.296, 180], { fill:'0D67FF', line:{ color:C.inkDk, width:2, dashType:'dash', endArrowType:'oval' } });
  sh(s, 'roundRect', [3.49, 1.48, 1.372, 1.075], { fill:C.blue, shadow:SH.close, rectRadius:adj(5562, 1.372, 1.075) });
  tx(s, "40,214", [3.545, 1.826, 1.263, 0.404], T.c15bnex13, { color:C.white });
  tx(s, "Your Value", [3.483, 2.18, 1.386, 0.273], T.c9x13, { color:C.white });
  tx(s, "Revenue", [3.543, 1.538, 1.263, 0.404], T.c15nex13, { color:C.white });
  sh(s, 'chevron', [5.125, 3.42, 1.383, 1.494], { fill:C.ocean, shadow:SH.wide, rectRadius:adj(18657, 1.383, 1.494) });
  sh(s, 'ellipse', [5.481, 2.944, 0.672, 0.672], { fill:C.white, line:{ color:C.hair, width:0.75 }, shadow:SH.tint });
  tx(s, "03", [5.623, 3.107, 0.388, 0.345], T.c18bnmz, { color:C.ocean });
  sh(s, 'line', [5.841, 2.659, 0, 0.284, 180], { fill:'0D67FF', line:{ color:C.inkDk, width:2, dashType:'dash', endArrowType:'oval' } });
  sh(s, 'roundRect', [5.155, 1.48, 1.372, 1.075], { fill:C.ocean, shadow:SH.close, rectRadius:adj(5562, 1.372, 1.075) });
  tx(s, "40,214", [5.209, 1.826, 1.263, 0.404], T.c15bnex13, { color:C.white });
  tx(s, "Your Value", [5.148, 2.18, 1.386, 0.273], T.c9x13, { color:C.white });
  tx(s, "Revenue", [5.208, 1.538, 1.263, 0.404], T.c15nex13, { color:C.white });
  sh(s, 'chevron', [6.793, 3.409, 1.383, 1.494], { fill:C.gold, shadow:SH.wide, rectRadius:adj(18657, 1.383, 1.494) });
  sh(s, 'ellipse', [7.149, 2.955, 0.672, 0.672], { fill:C.white, line:{ color:C.hair, width:0.75 }, shadow:SH.tint });
  tx(s, "04", [7.29, 3.118, 0.388, 0.345], T.c18bnmz, { color:C.gold });
  sh(s, 'line', [7.483, 2.659, 0, 0.296, 180], { fill:'0D67FF', line:{ color:C.inkDk, width:2, dashType:'dash', endArrowType:'oval' } });
  sh(s, 'roundRect', [6.797, 1.48, 1.372, 1.075], { fill:C.gold, shadow:SH.close, rectRadius:adj(5562, 1.372, 1.075) });
  tx(s, "40,214", [6.851, 1.826, 1.263, 0.404], T.c15bnex13, { color:C.white });
  tx(s, "Your Value", [6.79, 2.18, 1.386, 0.273], T.c9x13, { color:C.white });
  tx(s, "Revenue", [6.85, 1.538, 1.263, 0.404], T.c15nex13, { color:C.white });
  tx(s, "Product", [1.803, 4.499, 1.279, 0.316], T.c11bnx15, { color:C.white });
  tx(s, "Sales", [6.834, 4.499, 1.279, 0.316], T.c11bnx15, { color:C.white });
  tx(s, "Income", [5.151, 4.499, 1.279, 0.316], T.c11bnx15, { color:C.white });
  tx(s, "Market", [3.477, 4.499, 1.279, 0.316], T.c11bnx15, { color:C.white });
  icon(s, [7.29, 3.911, 0.455, 0.455], C.white);
  icon(s, [3.941, 3.911, 0.455, 0.455], C.white);
  icon(s, [2.285, 3.911, 0.455, 0.455], C.white);
  icon(s, [5.642, 3.911, 0.455, 0.455], C.white);
  footer(s, n);
}

// 7. four full-width arrows of decreasing length, icon + copy at the tips
function slide7(s, n) {
  sh(s, 'rightArrow', [0, 1.62, 5.345, 1.298], { fill:C.blue, shadow:SH.wide });
  sh(s, 'rightArrow', [0, 0.548, 6.124, 1.298], { fill:C.navy, shadow:SH.wide });
  sh(s, 'rightArrow', [0, 2.692, 4.566, 1.298], { fill:C.ocean, shadow:SH.wide });
  sh(s, 'rightArrow', [0, 3.765, 3.687, 1.298], { fill:C.gold, shadow:SH.wide });
  tx(s, L.tempor, [6.921, 1.248, 2.722, 0.455], T.l9x13, { color:C.grey });
  tx(s, "Value One", [6.921, 0.955, 2.576, 0.278], T.l12b, { color:C.navy });
  sh(s, 'ellipse', [6.275, 0.955, 0.502, 0.502], { fill:C.navy, shadow:SH.wide });
  icon(s, [6.372, 1.052, 0.307, 0.307], C.white);
  sh(s, 'ellipse', [5.485, 2.029, 0.502, 0.502], { fill:C.blue, shadow:SH.wide });
  tx(s, L.tempor, [6.131, 2.322, 2.722, 0.453], T.l9x13, { color:C.grey });
  tx(s, "Value Two", [6.131, 2.029, 2.576, 0.278], T.l12b, { color:C.blue });
  icon(s, [5.583, 2.131, 0.307, 0.307], C.white);
  sh(s, 'ellipse', [4.668, 3.101, 0.502, 0.502], { fill:C.ocean, shadow:SH.wide });
  tx(s, L.tempor, [5.314, 3.393, 2.722, 0.453], T.l9x13, { color:C.grey });
  tx(s, "Value Three", [5.314, 3.101, 2.576, 0.278], T.l12b, { color:C.ocean });
  icon(s, [4.765, 3.202, 0.307, 0.307], C.white);
  sh(s, 'ellipse', [3.777, 4.172, 0.502, 0.502], { fill:C.gold, shadow:SH.wide });
  tx(s, L.tempor, [4.423, 4.465, 2.722, 0.453], T.l9x13, { color:C.grey });
  tx(s, "Value Four", [4.423, 4.172, 2.576, 0.278], T.l12b, { color:C.gold });
  icon(s, [3.874, 4.274, 0.307, 0.307], C.white);
  tx(s, "+80K", [3.53, 2.005, 1.306, 0.53], T.l27bn, { color:C.white });
  tx(s, "+90K", [4.298, 0.93, 1.279, 0.53], T.l27bn, { color:C.white });
  tx(s, "+70K", [2.766, 3.089, 1.325, 0.53], T.l27bn, { color:C.white });
  tx(s, "+60K", [1.847, 4.157, 1.279, 0.53], T.l27bn, { color:C.white });
  tx(s, "Product", [0.505, 0.917, 1.279, 0.486], T.l18bnx15, { color:C.white });
  tx(s, "Income", [0.505, 3.062, 1.279, 0.486], T.l18bnx15, { color:C.white });
  tx(s, "Market", [0.505, 1.989, 1.279, 0.486], T.l18bnx15, { color:C.white });
  tx(s, "Sales", [0.505, 4.134, 1.279, 0.486], T.l18bnx15, { color:C.white });
  footer(s, n);
}

// 8. S-curve of half-circle arrows linking four icon discs
function slide8(s, n) {
  poly(s, G.waveUpA, [6.429, 1.831, 1.837, 1.044], { fill:C.gold });
  poly(s, G.waveDnA, [4.864, 2.542, 1.837, 1.044], { fill:C.ocean });
  poly(s, G.waveUpB, [3.299, 1.824, 1.83, 1.044], { fill:C.blue });
  poly(s, G.waveDnB, [1.734, 2.527, 1.837, 1.051], { fill:C.navy });
  sh(s, 'ellipse', [2.164, 2.232, 0.938, 0.938], { fill:C.navy, shadow:SH.wide });
  sh(s, 'ellipse', [6.852, 2.232, 0.938, 0.938], { fill:C.gold, shadow:SH.wide });
  sh(s, 'ellipse', [3.715, 2.232, 0.945, 0.938], { fill:C.blue, shadow:SH.wide });
  sh(s, 'ellipse', [5.31, 2.232, 0.945, 0.938], { fill:C.ocean, shadow:SH.wide });
  tx(s, L.subtitle, [2.681, 1.133, 4.638, 0.252], T.c11, { color:C.grey });
  tx(s, "Arrow Infographic", [2.6, 0.57, 4.799, 0.581], T.c30b, { color:C.ink });
  tx(s, "Product", [1.994, 4.003, 1.279, 0.316], T.c11bnx15, { color:C.ink });
  tx(s, "$1,300", [2.01, 3.602, 1.247, 0.486], T.c18bnx15, { color:C.navy });
  tx(s, [{ text:"Lorem ipsum dolor", options:{ breakLine:true } }, { text:"sit amet, consectetur adipiscing elit" }], [1.968, 4.316, 1.331, 0.625], T.c8x15, { color:C.ink });
  tx(s, "Sales", [3.536, 3.602, 1.279, 0.316], T.c11bnx15, { color:C.ink });
  tx(s, "$2,300", [3.552, 3.201, 1.247, 0.486], T.c18bnx15, { color:C.blue });
  tx(s, [{ text:"Lorem ipsum dolor", options:{ breakLine:true } }, { text:"sit amet, consectetur adipiscing elit" }], [3.51, 3.915, 1.331, 0.625], T.c8x15, { color:C.ink });
  tx(s, "Income", [5.114, 3.987, 1.279, 0.316], T.c11bnx15, { color:C.ink });
  tx(s, "$3,300", [5.13, 3.586, 1.247, 0.486], T.c18bnx15, { color:C.ocean });
  tx(s, [{ text:"Lorem ipsum dolor", options:{ breakLine:true } }, { text:"sit amet, consectetur adipiscing elit" }], [5.088, 4.299, 1.331, 0.625], T.c8x15, { color:C.ink });
  tx(s, "Market", [6.668, 3.586, 1.279, 0.316], T.c11bnx15, { color:C.ink });
  tx(s, "$4,300", [6.684, 3.185, 1.247, 0.486], T.c18bnx15, { color:C.gold });
  tx(s, [{ text:"Lorem ipsum dolor", options:{ breakLine:true } }, { text:"sit amet, consectetur adipiscing elit" }], [6.642, 3.899, 1.331, 0.625], T.c8x15, { color:C.ink });
  icon(s, [7.089, 2.471, 0.455, 0.455], C.white);
  icon(s, [3.965, 2.471, 0.455, 0.455], C.white);
  icon(s, [2.405, 2.471, 0.455, 0.455], C.white);
  icon(s, [5.555, 2.471, 0.455, 0.455], C.white);
  footer(s, n);
}

// 9. four percentage discs in a pinwheel with side captions
function slide9(s, n) {
  sh(s, 'triangle', [6.99, 2.481, 0.531, 0.458, 310.154], { fill:C.gold });
  sh(s, 'arc', [6.218, 2.709, 1.5, 1.5, 180], { flipH:true, line:{ color:C.gold, width:21.25 }, angleRange:[145.51, 51.61] });
  sh(s, 'arc', [4.906, 1.984, 1.5, 1.5], { flipH:true, line:{ color:C.ocean, width:21.25 }, angleRange:[151.34, 33.26] });
  sh(s, 'arc', [2.283, 1.984, 1.5, 1.5], { flipH:true, line:{ color:C.navy, width:21.25 }, angleRange:[149.72, 56.5] });
  sh(s, 'arc', [3.596, 2.709, 1.5, 1.5, 180], { flipH:true, line:{ color:C.blue, width:21.25 }, angleRange:[150.8, 30.97] });
  tx(s, L.subtitle, [2.681, 1.133, 4.638, 0.252], T.c11, { color:C.grey });
  tx(s, "Arrow Infographic", [2.6, 0.57, 4.799, 0.581], T.c30b, { color:C.ink });
  sh(s, 'ellipse', [2.462, 2.157, 1.164, 1.164], { fill:C.navy, shadow:SH.card });
  tx(s, [{ text:"24", options:{ fontSize:30 } }, { text:"%", options:{ fontSize:21 } }], [2.551, 2.384, 0.986, 0.53], T.c14x9, { color:C.white });
  tx(s, L.sit, [2.551, 2.754, 0.986, 0.409], T.c9x11, { color:C.white });
  sh(s, 'ellipse', [3.77, 2.881, 1.164, 1.164], { fill:C.blue, shadow:SH.card });
  tx(s, [{ text:"33", options:{ fontSize:30 } }, { text:"%", options:{ fontSize:21 } }], [3.859, 3.108, 0.986, 0.53], T.c14x9, { color:C.white });
  tx(s, L.sit, [3.859, 3.478, 0.986, 0.409], T.c9x11, { color:C.white });
  sh(s, 'ellipse', [5.078, 2.157, 1.164, 1.164], { fill:C.ocean, shadow:SH.card });
  tx(s, [{ text:"44", options:{ fontSize:30 } }, { text:"%", options:{ fontSize:21 } }], [5.167, 2.384, 0.986, 0.53], T.c14x9, { color:C.white });
  tx(s, L.sit, [5.167, 2.754, 0.986, 0.409], T.c9x11, { color:C.white });
  sh(s, 'ellipse', [6.406, 2.876, 1.164, 1.164], { fill:C.gold, shadow:SH.card });
  tx(s, [{ text:"54", options:{ fontSize:30 } }, { text:"%", options:{ fontSize:21 } }], [6.495, 3.104, 0.986, 0.53], T.c14x9, { color:C.white });
  tx(s, L.sit, [6.495, 3.474, 0.986, 0.409], T.c9x11, { color:C.white });
  tx(s, L.dot, [0.764, 2.414, 1.389, 0.65], T.c9x13, { color:C.ink });
  tx(s, "Sales", [0.55, 2.122, 1.819, 0.278], T.c12bn, { color:C.navy });
  tx(s, L.dot, [2.252, 4.137, 1.389, 0.65], T.c9x13, { color:C.ink });
  tx(s, "Product ", [2.037, 3.845, 1.819, 0.278], T.c12bn, { color:C.blue });
  tx(s, L.dot, [4.984, 4.133, 1.389, 0.65], T.c9x13, { color:C.ink });
  tx(s, "Income", [4.769, 3.842, 1.819, 0.278], T.c12bn, { color:C.ocean });
  tx(s, L.dot, [7.696, 2.414, 1.389, 0.65], T.c9x13, { color:C.ink });
  tx(s, "Market", [7.481, 2.122, 1.819, 0.278], T.c12bn, { color:C.gold });
  sh(s, 'ellipse', [2.779, 1.566, 0.502, 0.502], { fill:C.navy, shadow:SH.wide });
  icon(s, [2.876, 1.663, 0.307, 0.307], C.white);
  sh(s, 'ellipse', [4.081, 4.128, 0.502, 0.502], { fill:C.blue, shadow:SH.wide });
  icon(s, [4.179, 4.23, 0.307, 0.307], C.white);
  sh(s, 'ellipse', [6.777, 4.132, 0.502, 0.502], { fill:C.gold, shadow:SH.wide });
  icon(s, [6.874, 4.23, 0.307, 0.307], C.white);
  sh(s, 'ellipse', [5.421, 1.567, 0.502, 0.502], { fill:C.ocean, shadow:SH.wide });
  icon(s, [5.518, 1.669, 0.307, 0.307], C.white);
  footer(s, n);
}

// 10. circular three-arrow ring around a central 2024
function slide10(s, n) {
  poly(s, G.discLowerRight, [3.122, 2.977, 2.692, 2.059], { fill:C.gold, shadow:SH.wide });
  poly(s, G.discUpperLeft, [3.198, 1.29, 2.803, 1.73], { fill:C.navy, shadow:SH.wide });
  poly(s, G.discRight, [5.656, 1.66, 1.319, 3.178], { fill:C.blue, shadow:SH.wide });
  sh(s, 'chevron', [5.147, 3.879, 1.232, 1.183, 135.282], { fill:C.blueDk, shadow:SH.wide, rectRadius:adj(57006, 1.232, 1.183) });
  sh(s, 'chevron', [3.001, 2.521, 1.232, 1.183, 265.668], { fill:C.goldDk, shadow:SH.wide, rectRadius:adj(57006, 1.232, 1.183) });
  sh(s, 'chevron', [5.269, 1.285, 1.232, 1.183, 6.132], { fill:C.navyDk, shadow:SH.wide, rectRadius:adj(57006, 1.232, 1.183) });
  tx(s, "2024", [4.509, 2.837, 1.16, 0.581], T.c30bn, { color:C.black });
  tx(s, L.amet, [3.627, 1.662, 3.034, 2.716, 331.454], T.c14, { color:C.white });
  tx(s, L.amet, [3.649, 1.943, 3.034, 2.716, 96.332], T.c14, { color:C.white });
  tx(s, L.amet, [3.433, 1.823, 3.034, 2.716, 209.33], T.c14, { color:C.white });
  tx(s, "Arrow Infographic", [0.515, 0.57, 4.799, 0.581], T.l30b, { color:C.ink });
  tx(s, L.tempor, [7.009, 1.481, 1.877, 0.65], T.l9x13, { color:C.grey });
  tx(s, "Value One", [7.009, 1.188, 2.576, 0.278], T.l12b, { color:C.navy });
  sh(s, 'ellipse', [6.363, 1.188, 0.502, 0.502], { fill:C.navy, shadow:SH.wide });
  icon(s, [6.461, 1.285, 0.307, 0.307], C.white);
  sh(s, 'ellipse', [6.899, 4.01, 0.502, 0.502], { fill:C.blue, shadow:SH.wide });
  tx(s, "Lorem ipsum dolor sit amet, consectetur adipiscing elit, deiusmod tempor", [7.545, 4.303, 1.877, 0.65], T.l9x13, { color:C.grey });
  tx(s, "Value Two", [7.545, 4.01, 2.576, 0.278], T.l12b, { color:C.blue });
  icon(s, [6.997, 4.111, 0.307, 0.307], C.white);
  sh(s, 'ellipse', [2.598, 2.516, 0.502, 0.502], { fill:C.gold, shadow:SH.wide });
  tx(s, L.tempor, [0.586, 2.809, 1.892, 0.65], T.r9x13, { color:C.grey });
  tx(s, "Value Three", [-0.098, 2.516, 2.576, 0.278], T.r12b, { color:C.gold });
  icon(s, [2.695, 2.618, 0.307, 0.307], C.white);
  tx(s, "$35M", [0.97, 2.046, 1.865, 0.485], T.c27bnx9, { color:C.gold });
  tx(s, "$35M", [6.614, 0.707, 1.865, 0.485], T.c27bnx9, { color:C.navy });
  tx(s, "$35M", [7.171, 3.546, 1.865, 0.485], T.c27bnx9, { color:C.blue });
  tx(s, L.subtitle, [0.512, 1.145, 4.638, 0.252], T.l11, { color:C.grey });
  footer(s, n);
}

// 11. four arrow markers on a horizon line, values underneath
function slide11(s, n) {
  sh(s, 'line', [0, 2.932, 10.001, 0], { line:{ color:C.hair, width:7 } });
  poly(s, G.pinBanner, [1.268, 2.24, 0.968, 1.372, 270], { fill:C.navy });
  poly(s, G.pinCap, [1.19, 2.224, 0.304, 0.55, 270], { fill:C.navyDk });
  poly(s, G.pinBanner, [3.495, 2.24, 0.968, 1.372, 270], { fill:C.blue });
  poly(s, G.pinCap, [3.418, 2.224, 0.304, 0.55, 270], { fill:C.blueDk });
  poly(s, G.pinBanner, [5.723, 2.24, 0.968, 1.372, 270], { fill:C.ocean });
  poly(s, G.pinCap, [5.645, 2.224, 0.304, 0.55, 270], { fill:C.oceanDk });
  poly(s, G.pinBanner, [7.95, 2.24, 0.968, 1.372, 270], { fill:C.gold });
  poly(s, G.pinCap, [7.873, 2.224, 0.304, 0.55, 270], { fill:C.goldDk });
  tx(s, L.subtitle, [2.681, 1.133, 4.638, 0.252], T.c11, { color:C.grey });
  tx(s, "Arrow Infographic", [2.6, 0.57, 4.799, 0.581], T.c30b, { color:C.ink });
  tx(s, "Revenue", [3.25, 3.957, 1.279, 0.316], T.c11nx15, { color:C.ink });
  tx(s, "$2,300", [3.266, 3.556, 1.247, 0.486], T.c18bnx15, { color:C.blue });
  tx(s, [{ text:"Lorem ipsum dolor", options:{ breakLine:true } }, { text:"sit amet, consectetur adipiscing elit" }], [3.266, 4.269, 1.247, 0.571], T.c7x15, { color:C.ink });
  tx(s, "Revenue", [5.522, 3.957, 1.279, 0.316], T.c11nx15, { color:C.ink });
  tx(s, "$3,400", [5.538, 3.556, 1.247, 0.486], T.c18bnx15, { color:C.ocean });
  tx(s, [{ text:"Lorem ipsum dolor", options:{ breakLine:true } }, { text:"sit amet, consectetur adipiscing elit" }], [5.538, 4.269, 1.247, 0.571], T.c7x15, { color:C.ink });
  tx(s, "Revenue", [0.978, 3.957, 1.279, 0.316], T.c11nx15, { color:C.ink });
  tx(s, "$1,200", [1.146, 3.556, 0.942, 0.482], T.c18bnx15, { color:C.navy });
  tx(s, [{ text:"Lorem ipsum dolor", options:{ breakLine:true } }, { text:"sit amet, consectetur adipiscing elit" }], [0.994, 4.269, 1.247, 0.571], T.c7x15, { color:C.ink });
  tx(s, "Revenue", [7.795, 3.957, 1.279, 0.316], T.c11nx15, { color:C.ink });
  tx(s, "$4,500", [7.921, 3.556, 1.026, 0.482], T.c18bnx15, { color:C.gold });
  tx(s, [{ text:"Lorem ipsum dolor", options:{ breakLine:true } }, { text:"sit amet, consectetur adipiscing elit" }], [7.811, 4.269, 1.247, 0.571], T.c7x15, { color:C.ink });
  tx(s, "Product", [1.304, 2.746, 1.279, 0.315], T.l11bnx15, { color:C.white });
  tx(s, "Income", [3.504, 2.746, 1.279, 0.315], T.l11bnx15, { color:C.white });
  tx(s, "Market", [5.797, 2.752, 1.279, 0.315], T.l11bnx15, { color:C.white });
  tx(s, "Sales ", [8.025, 2.741, 1.279, 0.315], T.l11bnx15, { color:C.white });
  sh(s, 'ellipse', [1.103, 2.018, 0.502, 0.502], { fill:C.navy, shadow:SH.wide });
  icon(s, [1.201, 2.116, 0.307, 0.307], C.white);
  sh(s, 'ellipse', [3.313, 2.017, 0.502, 0.502], { fill:C.blue, shadow:SH.wide });
  icon(s, [3.41, 2.119, 0.307, 0.307], C.white);
  sh(s, 'ellipse', [7.772, 2.018, 0.502, 0.502], { fill:C.gold, shadow:SH.wide });
  icon(s, [7.87, 2.115, 0.307, 0.307], C.white);
  sh(s, 'ellipse', [5.546, 2.02, 0.502, 0.502], { fill:C.ocean, shadow:SH.wide });
  icon(s, [5.644, 2.121, 0.307, 0.307], C.white);
  footer(s, n);
}

// 12. four numbered tags feeding labelled value bars
function slide12(s, n) {
  poly(s, G.tagDown, [3.837, 0.629, 0.815, 0.454, 90], { fill:C.navy });
  sh(s, 'rect', [7.235, -0.954, 0.529, 5, 90], { fill:C.navy });
  poly(s, G.slantL, [4.055, 0.867, 1.361, 0.529, 90], { fill:C.navyDk });
  poly(s, G.tagDown2, [3.837, 1.934, 0.815, 0.454, 90], { fill:C.blue });
  sh(s, 'rect', [7.235, -0.11, 0.529, 5, 90], { fill:C.blue });
  poly(s, G.slantWide, [4.284, 1.938, 0.902, 0.529, 90], { fill:C.blueDk });
  poly(s, G.tagDown, [3.837, 3.235, 0.815, 0.454, 90], { fill:C.gold });
  sh(s, 'rect', [7.235, 0.734, 0.529, 5, 90], { fill:C.gold });
  poly(s, G.slantR, [4.284, 3.156, 0.902, 0.529, 90], { fill:C.goldDk });
  poly(s, G.tagDown2, [3.837, 4.54, 0.815, 0.454, 90], { fill:C.ocean });
  sh(s, 'rect', [7.235, 1.578, 0.53, 5, 90], { fill:C.ocean });
  poly(s, G.slantNarrow, [4.055, 4.229, 1.361, 0.529, 90], { fill:C.oceanDk });
  tx(s, "Product", [5.161, 1.273, 1.279, 0.486], T.l18bnx15, { color:C.white });
  tx(s, "Income", [5.161, 2.093, 1.279, 0.486], T.l18bnx15, { color:C.white });
  tx(s, "Market", [5.161, 2.941, 1.279, 0.486], T.l18bnx15, { color:C.white });
  tx(s, "Sales ", [5.161, 3.795, 1.279, 0.486], T.l18bnx15, { color:C.white });
  tx(s, "01", [3.666, 0.615, 1.26, 0.441], T.c18bnx13, { color:C.white });
  tx(s, "02", [3.639, 1.914, 1.26, 0.441], T.c18bnx13, { color:C.white });
  tx(s, [{ text:"0" }, { text:"3" }], [3.639, 3.215, 1.26, 0.441], T.c18bx13, { color:C.white });
  tx(s, "04", [3.635, 4.516, 1.26, 0.441], T.c18bnx13, { color:C.white });
  tx(s, "Your Text Here", [1.517, 0.594, 2.38, 0.32], T.r12bx13, { color:C.navy });
  tx(s, L.labore, [0.99, 0.907, 2.907, 0.473], T.r8x15, { color:C.grey });
  tx(s, "Your Text Here", [1.517, 1.786, 2.38, 0.32], T.r12bx13, { color:C.blue });
  tx(s, L.labore, [0.99, 2.099, 2.907, 0.473], T.r8x15, { color:C.grey });
  tx(s, "Your Text Here", [1.517, 2.979, 2.38, 0.32], T.r12bx13, { color:C.ocean });
  tx(s, L.labore, [0.99, 3.292, 2.907, 0.473], T.r8x15, { color:C.grey });
  tx(s, "Your Text Here", [1.517, 4.171, 2.38, 0.32], T.r12bx13, { color:C.gold });
  tx(s, L.labore, [0.99, 4.484, 2.907, 0.473], T.r8x15, { color:C.grey });
  tx(s, "$15.125,00", [7.149, 1.306, 2.664, 0.53], T.l27bn, { color:C.white });
  tx(s, "$25.320,00", [7.149, 2.136, 2.664, 0.53], T.l27bn, { color:C.white });
  tx(s, "$32.040,00", [7.149, 2.978, 2.664, 0.53], T.l27bn, { color:C.white });
  tx(s, "$41.721,00", [7.149, 3.82, 2.664, 0.53], T.l27bn, { color:C.white });
  sh(s, 'rightArrow', [6.601, 1.444, 0.282, 0.245], { fill:{ color:C.white, transparency:36 } });
  sh(s, 'rightArrow', [6.601, 2.275, 0.282, 0.245], { fill:{ color:C.white, transparency:36 } });
  sh(s, 'rightArrow', [6.601, 3.123, 0.282, 0.245], { fill:{ color:C.white, transparency:36 } });
  sh(s, 'rightArrow', [6.601, 3.97, 0.282, 0.245], { fill:{ color:C.white, transparency:36 } });
  footer(s, n);
}

// 13. four elbow arrows converging on a central icon disc
function slide13(s, n) {
  poly(s, G.stepTread, [2.704, 2.297, 1.648, 0.451, 90], { fill:C.navyDk });
  poly(s, G.stepRiser, [3.404, 2.586, 0.881, 1.08, 90], { fill:C.navy });
  poly(s, G.stepTread, [5.648, 3.749, 1.648, 0.451, 270], { fill:C.oceanDk });
  poly(s, G.stepRiser, [5.716, 2.83, 0.881, 1.08, 270], { fill:C.ocean });
  poly(s, G.stepTread, [3.559, 4.604, 1.648, 0.451], { fill:C.goldDk });
  poly(s, G.stepRiser, [4.547, 3.974, 0.881, 1.08], { fill:C.gold });
  poly(s, G.stepTread, [5.011, 1.661, 1.648, 0.451, 180], { fill:C.blueDk });
  poly(s, G.stepRiser, [4.791, 1.662, 0.881, 1.08, 180], { fill:C.blue });
  tx(s, L.subtitle, [2.681, 1.133, 4.638, 0.252], T.c11, { color:C.grey });
  tx(s, "Arrow Infographic", [2.6, 0.57, 4.799, 0.581], T.c30b, { color:C.ink });
  tx(s, L.tempor, [6.88, 2.434, 1.877, 0.65], T.l9x13, { color:C.grey });
  tx(s, "Cost", [6.88, 2.141, 2.576, 0.278], T.l12bn, { color:C.blue });
  tx(s, "$45M", [6.88, 1.661, 1.865, 0.485], T.l27bnx9, { color:C.blue });
  tx(s, L.tempor, [6.88, 4.18, 1.877, 0.65], T.l9x13, { color:C.grey });
  tx(s, "Product", [6.88, 3.887, 2.576, 0.278], T.l12bn, { color:C.ocean });
  tx(s, "$35M", [6.88, 3.406, 1.865, 0.485], T.l27bnx9, { color:C.ocean });
  tx(s, L.tempor, [1.15, 2.434, 1.877, 0.65], T.r9x13, { color:C.grey });
  tx(s, "Market ", [0.451, 2.141, 2.576, 0.278], T.r12bn, { color:C.navy });
  tx(s, "$55M", [1.162, 1.661, 1.865, 0.485], T.r27bnx9, { color:C.navy });
  tx(s, L.tempor, [1.15, 4.18, 1.877, 0.65], T.r9x13, { color:C.grey });
  tx(s, "Income", [0.451, 3.887, 2.576, 0.278], T.r12bn, { color:C.gold });
  tx(s, "$25M", [1.162, 3.406, 1.865, 0.485], T.r27bnx9, { color:C.gold });
  sh(s, 'ellipse', [4.558, 2.935, 0.881, 0.881], { fill:C.steel, shadow:SH.wide });
  icon(s, [4.729, 3.113, 0.539, 0.539], C.white);
  footer(s, n);
}

// 14. four tall arrow panels, icon top and bottom
function slide14(s, n) {
  tx(s, String(n), [7.593, 5.174, 0.444, 0.299], T.l14n, { color:'000000' });
  poly(s, G.shieldUp, [1.935, 1.343, 1.336, 3.72], { fill:C.navy, shadow:SH.lift });
  poly(s, G.chevUpNotch, [2.053, 1.438, 1.101, 0.534], { fill:C.white });
  poly(s, G.shieldUp, [3.545, 1.343, 1.336, 3.72, 180], { fill:C.blue, shadow:SH.lift });
  poly(s, G.chevUpNotch, [3.662, 4.433, 1.101, 0.534, 180], { fill:C.white });
  poly(s, G.shieldUp, [5.092, 1.343, 1.336, 3.72], { fill:C.ocean, shadow:SH.lift });
  poly(s, G.chevUpNotch, [5.209, 1.438, 1.101, 0.534], { fill:C.white });
  poly(s, G.shieldUp, [6.701, 1.343, 1.336, 3.72, 180], { fill:C.gold, shadow:SH.lift });
  poly(s, G.chevUpNotch, [6.818, 4.433, 1.101, 0.534, 180], { fill:C.white });
  tx(s, "Arrow Infographic", [2.6, 0.57, 4.799, 0.581], T.c30b, { color:C.ink });
  sh(s, 'ellipse', [2.209, 4.033, 0.799, 0.799], { fill:C.navy, shadow:SH.wide });
  icon(s, [2.364, 4.188, 0.489, 0.489], C.white);
  sh(s, 'ellipse', [3.815, 1.572, 0.799, 0.799], { fill:C.blue, shadow:SH.wide });
  icon(s, [3.97, 1.734, 0.489, 0.489], C.white);
  sh(s, 'ellipse', [6.969, 1.579, 0.799, 0.799], { fill:C.gold, shadow:SH.wide });
  icon(s, [7.124, 1.734, 0.489, 0.489], C.white);
  sh(s, 'ellipse', [5.36, 4.033, 0.799, 0.799], { fill:C.ocean, shadow:SH.wide });
  icon(s, [5.515, 4.195, 0.489, 0.489], C.white);
  tx(s, "+90M", [1.986, 2.287, 1.211, 0.48], T.c24bn, { color:C.white });
  tx(s, "2500+ employee", [1.846, 2.672, 1.491, 0.303], T.c9x15, { color:C.white });
  tx(s, L.serenity, [1.986, 2.93, 1.211, 0.984], T.c9x15, { color:C.white });
  tx(s, "+80M", [3.611, 2.618, 1.211, 0.48], T.c24bn, { color:C.white });
  tx(s, "2500+ employee", [3.47, 3.003, 1.491, 0.303], T.c9x15, { color:C.white });
  tx(s, L.serenity, [3.628, 3.261, 1.176, 0.984], T.c9x15, { color:C.white });
  tx(s, "+70M", [5.148, 2.294, 1.211, 0.48], T.c24bn, { color:C.white });
  tx(s, "2500+ employee", [5.008, 2.679, 1.491, 0.303], T.c9x15, { color:C.white });
  tx(s, L.serenity, [5.148, 2.937, 1.211, 0.984], T.c9x15, { color:C.white });
  tx(s, "+60M", [6.763, 2.618, 1.211, 0.48], T.c24bn, { color:C.white });
  tx(s, "2500+ employee", [6.623, 3.003, 1.491, 0.303], T.c9x15, { color:C.white });
  tx(s, L.serenity, [6.763, 3.261, 1.211, 0.984], T.c9x15, { color:C.white });
  footer(s, n);
}

// 15. four chevron-in-pill rows with percentage stats below
function slide15(s, n) {
  tx(s, L.subtitle, [2.681, 1.133, 4.638, 0.252], T.c11, { color:C.grey });
  tx(s, "Arrow Infographic", [2.6, 0.57, 4.799, 0.581], T.c30b, { color:C.ink });
  sh(s, 'roundRect', [1.071, 1.67, 0.861, 1.68, 270], { fill:'AECCFF', rectRadius:adj(28868, 0.861, 1.68) });
  poly(s, G.arrowGlyph, [1.954, 1.714, 0.776, 1.6], { fill:C.navy, shadow:SH.lift });
  sh(s, 'ellipse', [0.732, 2.249, 0.522, 0.522], { fill:C.navy, shadow:SH.wide });
  icon(s, [0.833, 2.35, 0.319, 0.319], C.white);
  tx(s, "Product ", [1.316, 2.371, 0.803, 0.278], T.l12bn, { color:C.navy });
  sh(s, 'roundRect', [3.324, 1.687, 0.861, 1.645, 270], { fill:'B4D5FF', rectRadius:adj(28868, 0.861, 1.645) });
  poly(s, G.arrowGlyph, [4.19, 1.714, 0.776, 1.6], { fill:C.blue, shadow:SH.lift });
  sh(s, 'ellipse', [3, 2.249, 0.522, 0.522], { fill:C.blue, shadow:SH.wide });
  icon(s, [3.101, 2.355, 0.319, 0.319], C.white);
  tx(s, "Market", [3.556, 2.371, 0.803, 0.278], T.l12bn, { color:C.navy });
  sh(s, 'roundRect', [5.556, 1.682, 0.861, 1.655, 270], { fill:'B8DCFE', rectRadius:adj(28868, 0.861, 1.655) });
  poly(s, G.arrowGlyph, [6.425, 1.714, 0.776, 1.6], { fill:C.ocean, shadow:SH.lift });
  sh(s, 'ellipse', [5.229, 2.249, 0.522, 0.522], { fill:C.ocean, shadow:SH.wide });
  icon(s, [5.33, 2.355, 0.319, 0.319], C.white);
  tx(s, "Target", [5.756, 2.371, 0.803, 0.278], T.l12bn, { color:C.navy });
  sh(s, 'roundRect', [7.747, 1.687, 0.861, 1.646, 270], { fill:'FFF3CB', rectRadius:adj(28868, 0.861, 1.646) });
  poly(s, G.arrowGlyph, [8.612, 1.714, 0.776, 1.6], { fill:C.gold, shadow:SH.lift });
  sh(s, 'ellipse', [7.424, 2.249, 0.522, 0.522], { fill:C.gold, shadow:SH.wide });
  icon(s, [7.525, 2.35, 0.319, 0.319], C.white);
  tx(s, "Income", [8.015, 2.371, 0.803, 0.278], T.l12bn, { color:C.gold });
  tx(s, [{ text:"50" }, { text:"%", options:{ superscript:true } }], [0.516, 3.427, 1.97, 0.53], T.c27bm, { color:C.navy });
  tx(s, "Step One", [0.516, 3.964, 1.97, 0.217], T.c11bx8, { color:C.ink });
  tx(s, L.incid, [0.514, 4.152, 1.975, 0.757], T.c9x15, { color:C.grey });
  tx(s, [{ text:"60" }, { text:"%", options:{ superscript:true } }], [2.796, 3.427, 1.97, 0.53], T.c27bm, { color:C.navy });
  tx(s, "Step Two", [2.796, 3.964, 1.97, 0.217], T.c11bx8, { color:C.ink });
  tx(s, L.incid, [2.794, 4.152, 1.975, 0.757], T.c9x15, { color:C.grey });
  tx(s, [{ text:"70" }, { text:"%", options:{ superscript:true } }], [5.002, 3.427, 1.97, 0.53], T.c27bm, { color:C.navy });
  tx(s, "Step Three", [5.002, 3.964, 1.97, 0.217], T.c11bx8, { color:C.ink });
  tx(s, L.incid, [5, 4.152, 1.975, 0.757], T.c9x15, { color:C.grey });
  tx(s, [{ text:"80" }, { text:"%", options:{ superscript:true } }], [7.208, 3.427, 1.97, 0.53], T.c27bm, { color:C.gold });
  tx(s, "Step Four", [7.208, 3.964, 1.97, 0.217], T.c11bx8, { color:C.ink });
  tx(s, L.incid, [7.206, 4.152, 1.975, 0.757], T.c9x15, { color:C.grey });
  footer(s, n);
}

// -------------------------------------------------------------------- main
const pres = new PptxGenJS();
pres.layout = 'LAYOUT_16x9';   // 10 x 5.625 in

[slide1,
 slide2,
 slide3,
 slide4,
 slide5,
 slide6,
 slide7,
 slide8,
 slide9,
 slide10,
 slide11,
 slide12,
 slide13,
 slide14,
 slide15]
  .forEach((build, i) => build(pres.addSlide(), i + 1));

pres.writeFile({ fileName: path.join(__dirname, '119c3a57-4f30-485a-90e4-c747e98bc97a_grok_final.pptx') })
  .then(f => console.log('wrote ' + f));
