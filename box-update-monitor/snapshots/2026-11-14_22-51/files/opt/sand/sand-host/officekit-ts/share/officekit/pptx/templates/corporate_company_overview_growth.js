/*
 * Recreation of "0b088c09-3a2a-4031-9f2e-a2d62009aa44.pptx" using pptxgenjs.
 *
 *     node 0b088c09-3a2a-4031-9f2e-a2d62009aa44_grok_final.js
 *
 * 20 slides at 26.66in x 15in.  The source deck has empty picture
 * placeholders and no embedded photography, so there is nothing to swap out
 * for an image placeholder - every visual below is a native pptxgenjs shape,
 * a custom-geometry path, or text.
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* --------------------------------------------------------------- palette */

const BLUE = '004BC0'; // theme accent1
const NAVY = '192229'; // theme accent2
const LIME = 'ADEE68'; // theme accent3
const MIST = 'E5E9EF'; // theme accent4
const DEEP = '003890'; // accent1 @ 75% lum
const MIDNIGHT = '002660'; // accent1 @ 50% lum
const OLIVE = '567734'; // accent3 @ 50% lum
const INK = '262626';
const CHARCOAL = '0D0D0D';
const SLATE = '131A1F';
const GRAY = '808080';
const MUTED = 'A6A6A6';
const SILVER = 'D9D9D9';
const WHITE = 'FFFFFF';
const BLACK = '000000';
const AMBER = 'FFC000';
const HAIR_DARK = '8F98A6'; // agenda rule, every other row
const HAIR_LITE = 'DADDE1';
const HATCH_BLUE = '002259'; // flat stand-in for the wdDnDiag accent1/dark pattern
const HATCH_LIME = '8FE731';
const HATCH_OLIVE = '234C4A';

const F = 'Poppins';
const F_SB = 'Poppins SemiBold';
const F_L = 'Poppins Light';
const F_EB = 'Poppins ExtraBold';
const F_EL = 'Poppins ExtraLight';

const W = 26.66; // slide width, inches
const H = 15.0; // slide height, inches

const SOFT = { type: 'outer', angle: 90, blur: 62, offset: 26, color: BLACK, opacity: 0.22, rotateWithShape: false };
const DROP = { type: 'outer', angle: 45, blur: 25, offset: 9, color: BLACK, opacity: 0.25, rotateWithShape: false };
const LIFT = { type: 'outer', angle: 45, blur: 60, offset: 30, color: BLACK, opacity: 0.25, rotateWithShape: false };

/* ------------------------------------------------------- shared copy text */

const SUBTITLE = 'Here goes your shorter sub title in second line.';
const CARD_TITLE = 'Trusted expertise in financial planning';
const IDEA_TITLE = 'iDea number four';
const DUMMY = '\u00A0is simply dummy text of the printing and typesetting industry';

const LOREM_LONG =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et ' +
  'dolore magna aliqua. Ut enim ad minim veniam, quis nostrudLorem ipsum dolor sit amet, consectetur ' +
  'adipiscing elit. Sed do eiusmod tempor ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor';
const LOREM_WIDE =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et ' +
  'dolore magna aliqua. Ut enim ad minim veniam, quis nostrudLorem ipsum dolor sit amet, consectetur ' +
  'adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim ' +
  'veniam, quis nostrud';
const LOREM_PANEL =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et ' +
  'dolore magna aliqua. Ut enim ad minim veniam, quis nostrudLorem ipsum dolor sit amet, consectetur ' +
  'adipiscing elit. Sed';
const LOREM_CARD =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et ' +
  'dolore magna aliqua. U Lorem ipsum';
const LOREM_SERVICE = 'Lorem ipsum dolor sit amet, onsectetur adipiscing elit. Sed do eiusmod tempor incididunt';
const LOREM_QUOTE =
  'Lorem ipsum dolor sit amet, onsectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et ' +
  'dolore magna aliqua. U Lorem ipsum tempor incididunt ut labore et dolore magna aliqua. U Lorem ipsum';

/* ------------------------------------------------- procedural path pieces */

const rnd = (v) => Math.round(v * 1000) / 1000;

// Quarter-circle band anchored on the bottom-right corner of the shape box.
// The deck's signature "rainbow" corner decorations are all built from these.
function quarterBand(rOuter, rInner) {
  const N = 20;
  const pts = [];
  for (let i = 0; i <= N; i++) {
    const a = (Math.PI / 2) * (i / N);
    pts.push({ x: rnd(1 - rOuter * Math.cos(a)), y: rnd(1 - rOuter * Math.sin(a)), moveTo: i === 0 });
  }
  if (rInner <= 0) {
    pts.push({ x: 1, y: 1 });
  } else {
    for (let i = N; i >= 0; i--) {
      const a = (Math.PI / 2) * (i / N);
      pts.push({ x: rnd(1 - rInner * Math.cos(a)), y: rnd(1 - rInner * Math.sin(a)) });
    }
  }
  pts.push({ close: true });
  return pts;
}

// Open sine curve used for the small hand-drawn squiggle motif.
function sineCurve(cycles, steps) {
  const pts = [];
  for (let i = 0; i <= steps; i++) {
    const t = i / steps;
    pts.push({ x: rnd(t), y: rnd(0.5 - 0.44 * Math.sin(t * cycles * 2 * Math.PI)), moveTo: i === 0 });
  }
  return pts;
}

// Puffy cloud bank: the silhouette is the union of half-circle mounds given as
// [centre, height, radius] triples (all normalised), sampled left to right so
// the outline never doubles back on itself.
function cloudBand(mounds) {
  const pts = [{ x: 0, y: 1, moveTo: true }];
  for (let i = 0; i <= 160; i++) {
    const x = i / 160;
    let y = 1;
    mounds.forEach(([cx, rise, radius]) => {
      const t = (x - cx) / radius;
      if (Math.abs(t) < 1) y = Math.min(y, 1 - rise * Math.sqrt(1 - t * t));
    });
    pts.push({ x: rnd(x), y: rnd(y) });
  }
  pts.push({ x: 1, y: 1 }, { close: true });
  return pts;
}

const RING_BANDS = [quarterBand(1, 0.675), quarterBand(0.575, 0.385), quarterBand(0.283, 0.185)];
const RING_SOLID = quarterBand(1, 0.67);
const RING_DISC = quarterBand(1, 0);
const SINE = sineCurve(4, 56);
const CLOUD = cloudBand([
  [0.00, 0.66, 0.075], [0.115, 0.94, 0.085], [0.28, 0.42, 0.085], [0.375, 0.50, 0.07],
  [0.47, 0.30, 0.07], [0.59, 0.92, 0.075], [0.70, 0.55, 0.075], [0.77, 0.62, 0.07],
  [0.865, 0.90, 0.085], [0.98, 0.55, 0.07],
]);

const HERO_PANEL = [
  {x:0,y:0,moveTo:true},{x:0.681,y:0},{x:0.626,y:0.057},{x:0.55,y:0.159},{x:0.482,y:0.279},{x:0.425,y:0.414},
  {x:0.378,y:0.564},{x:0.343,y:0.725},{x:0.322,y:0.896},{x:0.315,y:1},{x:0,y:1},{close:true}
];
const ARROW_NE = [
  {x:0.618,y:0,moveTo:true},{x:0.974,y:0.459},{x:1,y:0.5},{x:0.974,y:0.541},{x:0.618,y:1},{x:0.555,y:0.919},
  {x:0.84,y:0.554},{x:0,y:0.554},{x:0,y:0.446},{x:0.84,y:0.446},{x:0.555,y:0.081},{close:true}
];
const CARD_TILT = [
  {x:0.07,y:0,moveTo:true},{x:0.93,y:0},{x:0.946,y:0.007},{x:0.967,y:0.044},{x:0.984,y:0.106},{x:0.996,y:0.187},
  {x:1,y:0.282},{x:0.996,y:0.741},{x:0.976,y:0.848},{x:0.948,y:0.894},{x:0.058,y:1},{x:0.046,y:0.995},
  {x:0.03,y:0.972},{x:0.017,y:0.933},{x:0.001,y:0.816},{x:0,y:0.249},{x:0.007,y:0.158},{x:0.026,y:0.062},
  {x:0.047,y:0.016},{x:0.07,y:0},{close:true}
];
const HOOK = [
  {x:0.013,y:0,moveTo:true},{x:0.281,y:0},{x:0.266,y:0.07},{x:0.261,y:0.151},{x:0.264,y:0.216},{x:0.29,y:0.338},
  {x:0.311,y:0.395},{x:0.369,y:0.499},{x:0.446,y:0.585},{x:0.49,y:0.62},{x:0.588,y:0.675},{x:0.642,y:0.692},
  {x:0.755,y:0.707},{x:0.813,y:0.703},{x:0.882,y:0.688},{x:0.948,y:0.663},{x:1,y:0.631},{x:1,y:0.953},
  {x:0.934,y:0.976},{x:0.838,y:0.995},{x:0.755,y:1},{x:0.667,y:0.994},{x:0.582,y:0.978},{x:0.5,y:0.951},
  {x:0.423,y:0.914},{x:0.35,y:0.868},{x:0.283,y:0.814},{x:0.221,y:0.751},{x:0.166,y:0.682},{x:0.118,y:0.606},
  {x:0.077,y:0.524},{x:0.044,y:0.437},{x:0.02,y:0.346},{x:0.005,y:0.25},{x:0,y:0.151},{x:0.004,y:0.064},
  {close:true}
];
const CARD_SERVICE = [
  {x:0.167,y:1,moveTo:true},{x:0.852,y:0.999},{x:0.922,y:0.98},{x:0.974,y:0.941},{x:0.996,y:0.903},{x:1,y:0.874},
  {x:1,y:0.193},{x:0.993,y:0.167},{x:0.974,y:0.142},{x:0.918,y:0.114},{x:0.254,y:0.001},{x:0.18,y:0.005},
  {x:0.131,y:0.018},{x:0.088,y:0.039},{x:0.036,y:0.082},{x:0.006,y:0.137},{x:0,y:0.178},{x:0.001,y:0.889},
  {x:0.017,y:0.929},{x:0.063,y:0.972},{x:0.111,y:0.993},{x:0.167,y:1},{close:true}
];
const CHECK_MARK = [
  {x:0.945,y:0,moveTo:true},{x:1,y:0.075},{x:0.318,y:1},{x:0,y:0.571},{x:0.055,y:0.497},{x:0.318,y:0.86},
  {close:true}
];
const HOOK_WIDE = [
  {x:0.511,y:0,moveTo:true},{x:0.599,y:0.013},{x:0.683,y:0.05},{x:0.76,y:0.108},{x:0.83,y:0.187},
  {x:0.891,y:0.283},{x:0.941,y:0.394},{x:0.981,y:0.519},{x:1,y:0.621},{x:0.752,y:0.716},{x:0.714,y:0.594},
  {x:0.67,y:0.518},{x:0.635,y:0.479},{x:0.597,y:0.45},{x:0.555,y:0.431},{x:0.511,y:0.425},{x:0.452,y:0.436},
  {x:0.374,y:0.491},{x:0.33,y:0.55},{x:0.295,y:0.622},{x:0.262,y:0.753},{x:0.255,y:0.85},{x:0.259,y:0.905},
  {x:0.009,y:1},{x:0,y:0.85},{x:0.014,y:0.655},{x:0.052,y:0.476},{x:0.112,y:0.319},{x:0.191,y:0.187},
  {x:0.286,y:0.086},{x:0.394,y:0.022},{x:0.451,y:0.006},{x:0.511,y:0},{close:true}
];
const CARD_QUOTE = [
  {x:0.099,y:0,moveTo:true},{x:0.901,y:0},{x:0.935,y:0.008},{x:0.963,y:0.03},{x:0.99,y:0.077},{x:1,y:0.138},
  {x:1,y:0.836},{x:0.992,y:0.872},{x:0.97,y:0.905},{x:0.953,y:0.918},{x:0.932,y:0.922},{x:0.129,y:1},
  {x:0.085,y:0.99},{x:0.048,y:0.961},{x:0.013,y:0.9},{x:0.001,y:0.842},{x:0.001,y:0.122},{x:0.01,y:0.077},
  {x:0.037,y:0.03},{x:0.065,y:0.008},{x:0.099,y:0},{close:true}
];
const BULB_COIL = [
  {x:0.838,y:0.957,moveTo:true},{x:0.066,y:0.437},{x:0.025,y:0.346},{x:0.014,y:0.228},{x:0.03,y:0.131},
  {x:0.076,y:0.05},{x:0.112,y:0.03},{x:0.149,y:0.035},{x:0.91,y:0.54},{x:0.934,y:0.563},{x:0.961,y:0.613},
  {x:0.983,y:0.699},{x:0.986,y:0.772},{x:0.976,y:0.846},{x:0.954,y:0.907},{x:0.912,y:0.959},{x:0.876,y:0.971},
  {x:0.838,y:0.957},{close:true}
];
const BULB_BODY = [
  {x:0.619,y:0.415,moveTo:true},{x:0.61,y:0.342},{x:0.569,y:0.248},{x:0.447,y:0.133},{x:0.145,y:0.019},
  {x:0.081,y:0.03},{x:0.032,y:0.079},{x:0.015,y:0.167},{x:0.06,y:0.255},{x:0.344,y:0.374},{x:0.366,y:0.402},
  {x:0.368,y:0.415},{x:0.066,y:0.415},{x:0.066,y:0.719},{x:0.083,y:0.719},{x:0.159,y:0.793},{x:0.181,y:0.861},
  {x:0.225,y:0.914},{x:0.269,y:0.939},{x:0.328,y:0.948},{x:0.328,y:1},{x:0.738,y:1},{x:0.738,y:0.948},
  {x:0.838,y:0.914},{x:0.882,y:0.861},{x:0.904,y:0.793},{x:0.98,y:0.719},{x:1,y:0.719},{x:1,y:0.415},
  {x:0.619,y:0.415},{close:true}
];
const BULB_RING = [
  {x:0.917,y:0,moveTo:true},{x:0.083,y:0},{x:0.065,y:0.012},{x:0.04,y:0.072},{x:0.019,y:0.177},{x:0,y:0.5},
  {x:0.019,y:0.802},{x:0.048,y:0.945},{x:0.074,y:0.996},{x:0.917,y:1},{x:0.937,y:0.985},{x:0.963,y:0.916},
  {x:0.983,y:0.802},{x:0.998,y:0.606},{x:1,y:0.437},{x:0.992,y:0.27},{x:0.971,y:0.102},{x:0.947,y:0.027},
  {x:0.917,y:0},{close:true}
];
const BULB_RING_TALL = [
  {x:0.917,y:0,moveTo:true},{x:0.083,y:0},{x:0.048,y:0.026},{x:0.014,y:0.123},{x:0.001,y:0.243},
  {x:0.005,y:0.366},{x:0.026,y:0.469},{x:0.048,y:0.525},{x:0.074,y:0.554},{x:0.231,y:0.556},{x:0.239,y:0.708},
  {x:0.26,y:0.836},{x:0.292,y:0.932},{x:0.333,y:0.989},{x:0.364,y:1},{x:0.465,y:1},{x:0.636,y:1},
  {x:0.667,y:0.989},{x:0.708,y:0.932},{x:0.74,y:0.836},{x:0.761,y:0.708},{x:0.769,y:0.556},{x:0.917,y:0.556},
  {x:0.937,y:0.548},{x:0.955,y:0.525},{x:0.992,y:0.394},{x:0.998,y:0.21},{x:0.971,y:0.057},{x:0.947,y:0.015},
  {x:0.917,y:0},{close:true}
];
const BULB_CAP = [
  {x:0.94,y:1,moveTo:true},{x:0.972,y:0.825},{x:0.973,y:0.7},{x:0.941,y:0.548},{x:0.912,y:0.484},
  {x:0.876,y:0.434},{x:0.833,y:0.398},{x:0.259,y:0.051},{x:0.19,y:0.049},{x:0.124,y:0.087},{x:0.069,y:0.16},
  {x:0.041,y:0.227},{x:0,y:0.438},{x:0.94,y:1},{close:true}
];
const CHEVRON = [
  {x:0.433,y:0.24,moveTo:true},{x:0.688,y:0.5},{x:0.433,y:0.76},{x:0.375,y:0.702},{x:0.579,y:0.5},
  {x:0.375,y:0.298},{close:true}
];
const RIBBON_TIP = [
  {x:0.181,y:0,moveTo:true},{x:0.257,y:0.009},{x:0.321,y:0.032},{x:0.937,y:0.407},{x:0.984,y:0.451},{x:1,y:0.5},
  {x:0.984,y:0.549},{x:0.946,y:0.587},{x:0.333,y:0.962},{x:0.262,y:0.99},{x:0.182,y:1},{x:0.072,y:0.977},
  {x:0,y:0.923},{x:0,y:0.074},{x:0.077,y:0.018},{x:0.181,y:0},{close:true}
];
const RIBBON_FOLD = [
  {x:1,y:1,moveTo:true},{x:0.371,y:0},{x:0,y:0},{x:0.381,y:0.605},{x:0.475,y:0.727},{x:0.569,y:0.814},
  {x:0.706,y:0.906},{x:0.815,y:0.957},{x:0.966,y:0.998},{close:true}
];
const RIBBON_ARROW_L = [
  {x:0.776,y:0.039,moveTo:true},{x:0.982,y:0.413},{x:1,y:0.507},{x:0.982,y:0.587},{x:0.757,y:0.987},
  {x:0.713,y:0.996},{x:0.675,y:0.955},{x:0.658,y:0.763},{x:0.624,y:0.713},{x:0.082,y:0.711},{x:0,y:0.677},
  {x:0.285,y:0.285},{x:0.626,y:0.286},{x:0.657,y:0.238},{x:0.66,y:0.104},{x:0.681,y:0.032},{x:0.729,y:0},
  {x:0.776,y:0.039},{close:true}
];
const RIBBON_ARROW_R = [
  {x:0.984,y:0.407,moveTo:true},{x:1,y:0.493},{x:0.988,y:0.581},{x:0.824,y:0.967},{x:0.787,y:1},
  {x:0.745,y:0.945},{x:0.735,y:0.77},{x:0.711,y:0.715},{x:0.286,y:0.711},{x:0.206,y:0.663},{x:0.127,y:0.544},
  {x:0,y:0.225},{x:0.045,y:0.282},{x:0.141,y:0.315},{x:0.419,y:0.286},{x:0.707,y:0.287},{x:0.734,y:0.238},
  {x:0.745,y:0.05},{x:0.789,y:0},{x:0.823,y:0.032},{x:0.984,y:0.407},{close:true}
];
const PLANE = [
  {x:0.159,y:0.547,moveTo:true},{x:0.106,y:0.847},{x:0.702,y:0.547},{close:true},{x:0.106,y:0.153,moveTo:true},
  {x:0.159,y:0.453},{x:0.702,y:0.453},{close:true},{x:0,y:0,moveTo:true},{x:1,y:0.5},{x:0,y:1},{x:0.091,y:0.5},
  {x:0.014,y:0.088},{close:true}
];
const HALO = [
  {x:0.5,y:0,moveTo:true},{x:0.615,y:0.015},{x:0.669,y:0.032},{x:0.768,y:0.086},{x:0.813,y:0.122},
  {x:0.89,y:0.207},{x:0.949,y:0.31},{x:0.971,y:0.367},{x:0.997,y:0.489},{x:1,y:0.553},{x:0.996,y:0.626},
  {x:0.974,y:0.729},{x:0.903,y:0.881},{x:0.841,y:0.958},{x:0.794,y:1},{x:0.794,y:0.938},{x:0.871,y:0.847},
  {x:0.929,y:0.724},{x:0.953,y:0.612},{x:0.953,y:0.494},{x:0.929,y:0.383},{x:0.856,y:0.238},{x:0.745,y:0.128},
  {x:0.654,y:0.078},{x:0.553,y:0.052},{x:0.5,y:0.049},{x:0.447,y:0.052},{x:0.346,y:0.078},{x:0.299,y:0.1},
  {x:0.215,y:0.16},{x:0.144,y:0.238},{x:0.09,y:0.332},{x:0.047,y:0.494},{x:0.051,y:0.641},{x:0.071,y:0.724},
  {x:0.102,y:0.8},{x:0.144,y:0.869},{x:0.202,y:0.935},{x:0.202,y:0.997},{x:0.097,y:0.881},{x:0.026,y:0.729},
  {x:0.001,y:0.59},{x:0.013,y:0.426},{x:0.051,y:0.31},{x:0.11,y:0.207},{x:0.187,y:0.122},{x:0.28,y:0.056},
  {x:0.385,y:0.015},{x:0.5,y:0},{close:true}
];
const ROCKET_FIN = [
  {x:0.928,y:0.689,moveTo:true},{x:0.819,y:0.592},{x:0.681,y:0.422},{x:0.601,y:0.258},{x:0.5,y:0},
  {x:0.337,y:0.388},{x:0.301,y:0.449},{x:0.218,y:0.551},{x:0.063,y:0.701},{x:0.018,y:0.799},{x:0.011,y:0.883},
  {x:0.191,y:0.889},{x:0.34,y:0.91},{x:0.46,y:0.954},{x:0.5,y:1},{x:0.54,y:0.954},{x:0.66,y:0.91},
  {x:0.809,y:0.889},{x:0.989,y:0.883},{x:0.976,y:0.78},{x:0.928,y:0.689},{close:true}
];
const ROCKET_BODY = [
  {x:0.509,y:0.002,moveTo:true},{x:0.346,y:0.029},{x:0.183,y:0.089},{x:0.065,y:0.166},{x:0,y:0.235},
  {x:0.058,y:1},{x:0.942,y:1},{x:1,y:0.235},{x:0.935,y:0.166},{x:0.816,y:0.089},{x:0.653,y:0.029},
  {x:0.509,y:0.002},{close:true}
];
const ROCKET_NOSE = [
  {x:1,y:1,moveTo:true},{x:0.935,y:0.704},{x:0.897,y:0.581},{x:0.817,y:0.379},{x:0.733,y:0.229},
  {x:0.653,y:0.124},{x:0.583,y:0.056},{x:0.51,y:0.008},{x:0.51,y:0},{x:0.49,y:0},{x:0.49,y:0.008},
  {x:0.383,y:0.086},{x:0.308,y:0.171},{x:0.183,y:0.379},{x:0.103,y:0.581},{x:0.065,y:0.704},{x:0,y:1},{x:1,y:1},
  {close:true}
];
const ROCKET_FLAME = [
  {x:0.505,y:0,moveTo:true},{x:0.33,y:0.025},{x:0.184,y:0.062},{x:0.068,y:0.123},{x:0.032,y:0.179},
  {x:0.041,y:0.255},{x:0.096,y:0.439},{x:0.224,y:0.733},{x:0.352,y:0.915},{x:0.449,y:0.986},{x:0.5,y:1},
  {x:0.556,y:0.986},{x:0.606,y:0.957},{x:0.653,y:0.915},{x:0.741,y:0.801},{x:0.882,y:0.511},{x:0.973,y:0.212},
  {x:0.959,y:0.15},{x:0.863,y:0.08},{x:0.722,y:0.035},{x:0.505,y:0},{close:true}
];
const SWOOSH_LONG = [
  {x:0.28,y:1,moveTo:true},{x:0.29,y:0.926},{x:0.332,y:0.783},{x:0.401,y:0.644},{x:0.494,y:0.511},
  {x:0.55,y:0.447},{x:0.679,y:0.325},{x:0.83,y:0.21},{x:1,y:0.103},{x:0.796,y:0},{x:0.607,y:0.119},
  {x:0.44,y:0.247},{x:0.298,y:0.384},{x:0.236,y:0.455},{x:0.132,y:0.603},{x:0.056,y:0.758},{x:0.011,y:0.918},
  {x:0,y:1},{x:0.28,y:1},{close:true}
];
const SWOOSH_SHORT = [
  {x:0.103,y:1,moveTo:true},{x:0.21,y:0.829},{x:0.326,y:0.678},{x:0.449,y:0.55},{x:0.513,y:0.494},
  {x:0.646,y:0.402},{x:0.784,y:0.335},{x:0.927,y:0.295},{x:1,y:0.286},{x:1,y:0},{x:0.918,y:0.01},
  {x:0.838,y:0.028},{x:0.759,y:0.053},{x:0.605,y:0.128},{x:0.457,y:0.232},{x:0.385,y:0.294},{x:0.248,y:0.438},
  {x:0.119,y:0.607},{x:0,y:0.799},{x:0.103,y:1},{close:true}
];
const CLOUD_BODY = [
  {x:0.5,y:0.091,moveTo:true},{x:0.428,y:0.111},{x:0.375,y:0.159},{x:0.323,y:0.272},{x:0.313,y:0.409},
  {x:0.267,y:0.41},{x:0.234,y:0.428},{x:0.2,y:0.478},{x:0.186,y:0.545},{x:0.129,y:0.566},{x:0.092,y:0.609},
  {x:0.064,y:0.708},{x:0.081,y:0.824},{x:0.127,y:0.888},{x:0.18,y:0.909},{x:0.848,y:0.902},{x:0.915,y:0.833},
  {x:0.938,y:0.727},{x:0.921,y:0.635},{x:0.875,y:0.568},{x:0.852,y:0.425},{x:0.779,y:0.335},{x:0.688,y:0.318},
  {x:0.654,y:0.207},{x:0.599,y:0.131},{x:0.552,y:0.101},{x:0.5,y:0.091},{close:true},{x:0.5,y:0,moveTo:true},
  {x:0.574,y:0.016},{x:0.631,y:0.055},{x:0.693,y:0.138},{x:0.73,y:0.23},{x:0.786,y:0.244},{x:0.836,y:0.279},
  {x:0.911,y:0.4},{x:0.934,y:0.526},{x:0.992,y:0.648},{x:0.998,y:0.774},{x:0.953,y:0.908},{x:0.903,y:0.967},
  {x:0.856,y:0.993},{x:0.166,y:0.998},{x:0.079,y:0.95},{x:0.018,y:0.846},{x:0.002,y:0.689},{x:0.043,y:0.555},
  {x:0.082,y:0.506},{x:0.137,y:0.472},{x:0.191,y:0.363},{x:0.256,y:0.324},{x:0.288,y:0.182},{x:0.346,y:0.081},
  {x:0.423,y:0.017},{x:0.5,y:0},{close:true}
];
const CLOUD_ARROW_DOWN = [
  {x:0.469,y:0.318,moveTo:true},{x:0.531,y:0.318},{x:0.531,y:0.616},{x:0.602,y:0.511},{x:0.648,y:0.58},
  {x:0.5,y:0.79},{x:0.352,y:0.58},{x:0.398,y:0.511},{x:0.469,y:0.616},{close:true}
];
const CLOUD_ARROW_UP = [
  {x:0.5,y:0.347,moveTo:true},{x:0.648,y:0.557},{x:0.602,y:0.625},{x:0.531,y:0.52},{x:0.531,y:0.818},
  {x:0.469,y:0.818},{x:0.469,y:0.52},{x:0.398,y:0.625},{x:0.352,y:0.557},{x:0.477,y:0.375},{close:true}
];
const ICON_CLOUD_DOWN = CLOUD_ARROW_DOWN.concat(CLOUD_BODY);
const ICON_CLOUD_UP = CLOUD_ARROW_UP.concat(CLOUD_BODY);
const ICON_SCALES = [
  {x:0.702,y:0.619,moveTo:true},{x:0.742,y:0.691},{x:0.767,y:0.706},{x:0.808,y:0.714},{x:0.849,y:0.706},
  {x:0.874,y:0.691},{x:0.913,y:0.619},{close:true},{x:0.087,y:0.619,moveTo:true},{x:0.126,y:0.691},
  {x:0.151,y:0.706},{x:0.192,y:0.714},{x:0.233,y:0.706},{x:0.258,y:0.691},{x:0.298,y:0.619},{close:true},
  {x:0.808,y:0.333,moveTo:true},{x:0.721,y:0.524},{x:0.894,y:0.524},{close:true},{x:0.192,y:0.333,moveTo:true},
  {x:0.106,y:0.524},{x:0.279,y:0.524},{close:true},{x:0.5,y:0.095,moveTo:true},{x:0.471,y:0.107},
  {x:0.462,y:0.143},{x:0.471,y:0.178},{x:0.5,y:0.19},{x:0.529,y:0.178},{x:0.538,y:0.143},{x:0.529,y:0.107},
  {x:0.5,y:0.095},{close:true},{x:0.5,y:0,moveTo:true},{x:0.532,y:0.004},{x:0.567,y:0.024},{x:0.608,y:0.095},
  {x:0.885,y:0.095},{x:0.885,y:0.19},{x:0.832,y:0.19},{x:1,y:0.56},{x:0.995,y:0.625},{x:0.977,y:0.686},
  {x:0.944,y:0.74},{x:0.91,y:0.774},{x:0.862,y:0.801},{x:0.819,y:0.81},{x:0.775,y:0.807},{x:0.725,y:0.787},
  {x:0.672,y:0.74},{x:0.639,y:0.686},{x:0.62,y:0.625},{x:0.615,y:0.56},{x:0.784,y:0.19},{x:0.608,y:0.19},
  {x:0.581,y:0.243},{x:0.538,y:0.277},{x:0.538,y:0.905},{x:0.692,y:0.905},{x:0.692,y:1},{x:0.308,y:1},
  {x:0.308,y:0.905},{x:0.462,y:0.905},{x:0.462,y:0.277},{x:0.419,y:0.243},{x:0.392,y:0.19},{x:0.216,y:0.19},
  {x:0.385,y:0.56},{x:0.371,y:0.663},{x:0.32,y:0.75},{x:0.275,y:0.787},{x:0.235,y:0.804},{x:0.181,y:0.81},
  {x:0.138,y:0.801},{x:0.099,y:0.781},{x:0.056,y:0.74},{x:0.023,y:0.686},{x:0.005,y:0.625},{x:0,y:0.56},
  {x:0.168,y:0.19},{x:0.115,y:0.19},{x:0.115,y:0.095},{x:0.392,y:0.095},{x:0.433,y:0.024},{x:0.468,y:0.004},
  {x:0.5,y:0},{close:true}
];
const ICON_DOWNLOAD = [
  {x:0,y:0.917,moveTo:true},{x:1,y:0.917},{x:1,y:1},{x:0,y:1},{close:true},{x:0.444,y:0,moveTo:true},
  {x:0.556,y:0},{x:0.556,y:0.69},{x:0.847,y:0.469},{x:0.931,y:0.531},{x:0.542,y:0.823},{x:0.5,y:0.849},
  {x:0.458,y:0.823},{x:0.069,y:0.531},{x:0.153,y:0.469},{x:0.444,y:0.69},{close:true}
];
const TICK = [
  {x:0.933,y:0.163,moveTo:true},{x:0.99,y:0.221},{x:0.5,y:0.707},{x:0.279,y:0.49},{x:0.337,y:0.433},
  {x:0.5,y:0.599},{close:true}
];
const ARCH = [
  {x:0.504,y:0,moveTo:true},{x:0.651,y:0.02},{x:0.785,y:0.074},{x:0.888,y:0.154},{x:0.959,y:0.253},
  {x:0.989,y:0.336},{x:1,y:0.406},{x:1,y:1},{x:0.873,y:1},{x:0.873,y:0.419},{x:0.862,y:0.35},{x:0.836,y:0.284},
  {x:0.776,y:0.208},{x:0.713,y:0.162},{x:0.612,y:0.12},{x:0.503,y:0.105},{x:0.394,y:0.118},{x:0.316,y:0.145},
  {x:0.227,y:0.203},{x:0.165,y:0.278},{x:0.137,y:0.343},{x:0.128,y:0.389},{x:0.127,y:1},{x:0,y:1},
  {x:0.001,y:0.384},{x:0.011,y:0.326},{x:0.043,y:0.244},{x:0.117,y:0.147},{x:0.222,y:0.069},{x:0.357,y:0.017},
  {x:0.504,y:0},{close:true}
];
const FUNNEL_TOP = [
  {x:0.334,y:0,moveTo:true},{x:0.356,y:0.056},{x:0.399,y:0.119},{x:0.45,y:0.16},{x:0.492,y:0.174},
  {x:0.542,y:0.17},{x:0.596,y:0.139},{x:0.634,y:0.096},{x:0.668,y:0.032},{x:1,y:0.617},{x:0.981,y:0.654},
  {x:0.926,y:0.741},{x:0.866,y:0.816},{x:0.802,y:0.88},{x:0.733,y:0.931},{x:0.661,y:0.969},{x:0.585,y:0.992},
  {x:0.468,y:0.998},{x:0.353,y:0.969},{x:0.281,y:0.931},{x:0.212,y:0.88},{x:0.148,y:0.816},{x:0.088,y:0.741},
  {x:0,y:0.589},{close:true}
];
const FUNNEL_MID = [
  {x:0.195,y:0,moveTo:true},{x:0.262,y:0.147},{x:0.342,y:0.255},{x:0.385,y:0.292},{x:0.43,y:0.318},
  {x:0.502,y:0.333},{x:0.574,y:0.318},{x:0.619,y:0.292},{x:0.662,y:0.255},{x:0.742,y:0.147},{x:0.805,y:0.011},
  {x:1,y:0.482},{x:0.988,y:0.513},{x:0.932,y:0.635},{x:0.871,y:0.741},{x:0.77,y:0.869},{x:0.697,y:0.932},
  {x:0.621,y:0.975},{x:0.542,y:0.997},{x:0.502,y:1},{x:0.462,y:0.997},{x:0.383,y:0.975},{x:0.307,y:0.932},
  {x:0.234,y:0.869},{x:0.166,y:0.788},{x:0.102,y:0.69},{x:0,y:0.472},{close:true}
];
const FUNNEL_BASE = [
  {x:0.133,y:0,moveTo:true},{x:0.207,y:0.188},{x:0.254,y:0.274},{x:0.304,y:0.345},{x:0.358,y:0.4},
  {x:0.414,y:0.437},{x:0.501,y:0.459},{x:0.589,y:0.437},{x:0.645,y:0.4},{x:0.698,y:0.345},{x:0.749,y:0.274},
  {x:0.796,y:0.188},{x:0.867,y:0.008},{x:1,y:0.391},{x:0.992,y:0.415},{x:0.935,y:0.561},{x:0.873,y:0.689},
  {x:0.771,y:0.843},{x:0.698,y:0.918},{x:0.622,y:0.97},{x:0.542,y:0.997},{x:0.501,y:1},{x:0.46,y:0.997},
  {x:0.381,y:0.97},{x:0.304,y:0.918},{x:0.231,y:0.843},{x:0.162,y:0.746},{x:0.098,y:0.628},{x:0,y:0.382},
  {close:true}
];

/* ------------------------------------------------------------- primitives */

// Custom-geometry shape.  Path points are stored normalised to 0..1 so the
// same outline can be reused at any size; scale them into inches here.
function poly(s, pts, x, y, w, h, o) {
  const scaled = pts.map((p) => {
    if (p.close) return p;
    const q = { x: p.x * w, y: p.y * h };
    if (p.moveTo) q.moveTo = true;
    if (p.curve) {
      q.curve = { type: p.curve.type, x1: p.curve.x1 * w, y1: p.curve.y1 * h };
      if (p.curve.x2 !== undefined) { q.curve.x2 = p.curve.x2 * w; q.curve.y2 = p.curve.y2 * h; }
    }
    return q;
  });
  s.addShape('custGeom', Object.assign({ points: scaled, x, y, w, h }, o));
}

// Rounded rectangle whose corner radius is given as the OOXML "adj" fraction.
function round(s, x, y, w, h, adj, o) {
  s.addShape('roundRect', Object.assign({ x, y, w, h, rectRadius: adj * Math.min(w, h) }, o));
}

function txt(s, body, o) {
  s.addText(body, Object.assign({ fontFace: F, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6] }, o));
}

// PowerPoint stores the deck's big word-marks as 270deg-rotated vertical text
// boxes; visually they are ordinary horizontal text, so swap w/h about centre.
function unrotate(x, y, w, h) {
  return { x: x + w / 2 - h / 2, y: y + h / 2 - w / 2, w: h, h: w };
}

/* -------------------------------------------------------- shared ornaments */

// Three concentric quarter rings; `q` picks the corner they curl into.
function rings(s, x, y, w, h, fill, transparency, flipH) {
  RING_BANDS.forEach((band) => poly(s, band, x, y, w, h, { fill: { color: fill, transparency }, flipH }));
}

function blob(s, x, y, w, h, fill, transparency, flipH) {
  poly(s, RING_SOLID, x, y, w, h, { fill: { color: fill, transparency }, flipH });
}

// The two-tone squiggle mark: `x`,`y` is the top (first) wave's corner.
function squiggle(s, x, y, top, bottom) {
  poly(s, SINE, x, y, 2.672, 0.38, { line: { color: top, width: 3 } });
  poly(s, SINE, x - 0.648, y + 0.326, 2.672, 0.38, { line: { color: bottom, width: 3 } });
}

// Master-slide furniture: hairline, site URL and the page number.
function chrome(s, n) {
  s.addShape('line', { x: 0.429, y: 13.99, w: 25.653, h: 0, line: { color: BLACK, transparency: 82, width: 1 } });
  txt(s, 'www.yoursitename.com', {
    x: 0.429, y: 14.279, w: 7.306, h: 0.404, fontSize: 18, color: BLACK, transparency: 48,
  });
  txt(s, String(n), {
    x: 25.184, y: 14.122, w: 1.051, h: 0.799, fontSize: 24, bold: true, fontFace: F_SB,
    color: GRAY, align: 'right', valign: 'middle',
  });
}

// Section head used on most content slides: bold headline + muted sub line.
function heading(s, x, y, title, sub, titleW) {
  txt(s, title, { x, y, w: titleW || 13.793, h: 1.01, fontSize: 54, bold: true, color: BLACK });
  if (sub) txt(s, sub, { x, y: y + 1.035, w: 9.572, h: 0.505, fontSize: 24, color: SLATE, transparency: 40 });
}

// Round white badge with the north-east arrow glyph (appears on every card).
function arrowBadge(s, x, y, d, arrowColor) {
  s.addShape('ellipse', { x, y, w: d, h: d, fill: WHITE, shadow: DROP });
  poly(s, ARROW_NE, x + d * 0.367, y + d * 0.39, d * 0.309, d * 0.24, { fill: arrowColor, rotate: 346.19 });
}

// Left-aligned / right-aligned "iDea number four" caption pair.
function ideaLabel(s, x, y, title, align) {
  txt(s, title, { x, y, w: 5.018, h: 0.505, fontSize: 24, fontFace: F_SB, align, lineSpacingMultiple: 1 });
  txt(s, [{ text: 'Lorem Ipsum', options: { bold: true } }, { text: DUMMY }], {
    x: align === 'right' ? x + 0.77 : x + 0.788, y: y + 0.621, w: 4.248, h: 1.111,
    fontSize: 20, color: MUTED, align, lineSpacingMultiple: 1,
  });
}

/* ============================================================== slide 01 */
/* Title / cover: full-bleed blue panel with the "Welcome." word-mark.      */

function slide01(s) {
  poly(s, HERO_PANEL, 0, 0, W, H, { fill: BLUE });
  poly(s, RING_DISC, 20.658, 9.792, 6.002, 5.208, { fill: { color: NAVY, transparency: 69 } });
  rings(s, 0.009, 4.644, 10.896, 10.357, MIDNIGHT, 76, true);
  blob(s, -0.062, 9.052, 6.458, 5.928, LIME, 0, true);
  squiggle(s, 2.061, 2.174, WHITE, WHITE);
  txt(s, [
    { text: 'Welcome', options: { color: WHITE } },
    { text: '.', options: { color: LIME } },
  ], Object.assign(unrotate(5.159, -0.502, 2.137, 9.628), { fontSize: 115, bold: true, fontFace: F_SB }));
  txt(s, 'To presentation', Object.assign(unrotate(5.35, 1.183, 0.74, 8.615), {
    fontSize: 32, fontFace: F_L, color: WHITE,
  }));
  chrome(s, 1);
}

/* ============================================================== slide 02 */
/* Agenda: twelve ruled rows with page numbers next to a vertical wordmark. */

const AGENDA = [
  'About company', 'Our vision', 'Our mission', 'Why choose us',
  'Our speciality features', 'Our services', 'Our unique projects', 'Our team members',
  'Client testimonials', 'Annual financial growth', 'Our unique process', 'Key metrices',
];

function slide02(s) {
  rings(s, 15.764, 4.643, 10.896, 10.357, MIST, 76);
  blob(s, 20.314, 9.051, 6.458, 5.928, LIME, 66);

  const rowY = [1.383, 2.369, 3.383, 4.368, 5.341, 6.327, 7.293, 8.279, 9.231, 10.216, 11.189, 12.175];
  AGENDA.forEach((label, i) => {
    const y = rowY[i];
    txt(s, label, { x: 14.202, y, w: 6.982, h: 0.505, fontSize: 24, color: CHARCOAL });
    txt(s, String((i + 1) * 10), {
      x: 24.14, y, w: 1.482, h: 0.505, fontSize: 24, bold: true, fontFace: F_SB, color: CHARCOAL, align: 'right',
    });
    s.addShape('line', {
      x: 13.767, y: y + 0.653, w: 11.854, h: 0, line: { color: i % 2 ? HAIR_LITE : HAIR_DARK, width: 1 },
    });
  });

  s.addShape('line', { x: 12.277, y: 8.5, w: 0, h: 5.524, line: { color: BLUE, transparency: 66, width: 1 } });
  txt(s, [
    { text: 'Agenda', options: { color: NAVY } },
    { text: '.', options: { color: LIME } },
  ], { x: 10.954, y: 1.36, w: 2.996, h: 11.616, fontSize: 166, bold: true, fontFace: F_SB, vert: 'vert' });
  squiggle(s, 1.034, 12.299, LIME, BLUE);
  chrome(s, 2);
}

/* ============================================================== slide 03 */
/* About the growth: copy block plus two KPI tiles (donut + mini bar chart).*/

function slide03(s) {
  rings(s, 0.009, 4.644, 10.896, 10.357, MIST, 36, true);
  blob(s, -0.103, 9.052, 6.458, 5.928, LIME, 0, true);

  txt(s, 'About the growth of our company.', { x: 14.361, y: 2.082, w: 9.814, h: 1.919, fontSize: 54, bold: true });
  txt(s, SUBTITLE, { x: 14.331, y: 4.148, w: 7.201, h: 0.438, fontSize: 20, bold: true, fontFace: F_SB, color: SLATE });
  txt(s, LOREM_LONG, {
    x: 14.331, y: 5.159, w: 9.265, h: 2.121, fontSize: 20, color: GRAY, transparency: 53, align: 'justify',
  });

  // Retention tile: 90% progress donut drawn as two pies over a filled circle.
  round(s, 14.364, 7.5, 4.368, 5.418, 0.20333, { fill: BLUE });
  s.addShape('pie', { x: 15.188, y: 8.375, w: 2.715, h: 2.715, fill: DEEP, angleRange: [123, 51] });
  s.addShape('pie', { x: 15.186, y: 8.375, w: 2.715, h: 2.715, fill: WHITE, angleRange: [123, 353] });
  s.addShape('ellipse', { x: 15.345, y: 8.538, w: 2.397, h: 2.397, fill: BLUE });
  txt(s, '90%', { x: 15.526, y: 9.249, w: 2.035, h: 0.909, fontSize: 48, fontFace: F_L, color: WHITE, align: 'center' });
  txt(s, 'Retention rate', {
    x: 14.555, y: 11.572, w: 3.986, h: 0.505, fontSize: 24, bold: true, fontFace: F_SB, color: WHITE, align: 'center',
  });

  // Weekly growth tile: six olive bars, the last one hatched.
  round(s, 19.217, 7.521, 4.368, 5.418, 0.20333, { fill: LIME });
  txt(s, 'Weekly growth', {
    x: 19.418, y: 11.593, w: 3.986, h: 0.505, fontSize: 24, bold: true, fontFace: F_SB, color: OLIVE, align: 'center',
  });
  const bars = [
    [19.797, 9.488, 1.602, OLIVE], [20.348, 9.017, 2.074, OLIVE], [20.899, 8.789, 2.302, OLIVE],
    [21.451, 9.12, 1.97, OLIVE], [22.038, 8.706, 2.385, OLIVE], [22.575, 8.187, 2.903, HATCH_OLIVE],
  ];
  bars.forEach(([x, y, h, c]) => round(s, x, y, 0.326, h, 0.31404, { fill: c }));

  squiggle(s, 1.11, 0.88, LIME, BLUE);
  chrome(s, 3);
}

/* ============================================================== slide 04 */
/* Vision / mission: two stacked copy blocks split by a rule.               */

function slide04(s) {
  rings(s, 0.009, 4.644, 10.896, 10.357, MIST, 36, true);
  blob(s, -0.103, 9.052, 6.458, 5.928, LIME, 0, true);

  // Loose pill shapes that stand in for the empty photo collage on the left.
  round(s, 2.947, 10.21, 3.232, 1.02, 0.41667, { fill: { color: LIME, transparency: 36 } });
  round(s, 10.542, 3.371, 2.167, 1.02, 0.41667, { fill: { color: BLUE, transparency: 45 } });
  round(s, 2.947, 3.352, 1.712, 1.02, 0.41667, { fill: { color: NAVY, transparency: 43 } });

  [['Our vision', 2.637, 14.585], ['Our mission', 8.829, 14.631]].forEach(([title, y, x]) => {
    txt(s, title, { x, y, w: 9.814, h: 1.01, fontSize: 54, bold: true });
    txt(s, LOREM_LONG, {
      x, y: y + 1.413, w: 9.265, h: 2.121, fontSize: 20, color: GRAY, transparency: 53, align: 'justify',
    });
  });
  s.addShape('line', { x: 14.652, y: 7.5, w: 10.169, h: 0, line: { color: BLACK, transparency: 61, width: 1 } });
  squiggle(s, 22.289, 2.76, LIME, BLUE);
  chrome(s, 4);
}

/* ============================================================== slide 05 */
/* Partners: two paragraphs of copy plus two arrow bullet rows.            */

function slide05(s) {
  rings(s, 15.764, 4.643, 10.896, 10.357, MIST, 36);
  blob(s, 20.314, 9.051, 6.458, 5.928, LIME);

  txt(s, 'About Our partners in progress.', { x: 2.694, y: 2.319, w: 9.814, h: 1.919, fontSize: 54, bold: true });
  txt(s, SUBTITLE, { x: 2.664, y: 4.385, w: 7.201, h: 0.438, fontSize: 20, bold: true, fontFace: F_SB, color: SLATE });
  txt(s, [{ text: LOREM_LONG, options: { breakLine: true } }, { text: LOREM_LONG }], {
    x: 2.664, y: 5.396, w: 10.086, h: 4.477, fontSize: 20, color: GRAY, transparency: 53, align: 'justify',
  });

  arrowRow(s, 2.694, 10.413, BLUE, WHITE, 'Innovating through design and technology.', 8.198);
  arrowRow(s, 2.694, 11.482, LIME, OLIVE, 'Turning ideas into intelligent solutions.', 7.101);
  squiggle(s, 23.503, 3.019, LIME, BLUE);
  chrome(s, 5);
}

// Rounded square icon + arrow glyph + caption; shared by slides 5 and 6.
function arrowRow(s, x, y, tile, arrow, label, labelW) {
  round(s, x, y, 0.76, 0.603, 0.37671, { fill: tile });
  poly(s, ARROW_NE, x + 0.241, y + 0.206, 0.279, 0.216, { fill: arrow, rotate: 325.25 });
  txt(s, label, { x: x + 1.063, y: y + 0.049, w: labelW, h: 0.505, fontSize: 24 });
}

/* ============================================================== slide 06 */
/* Company details: one copy block plus three arrow bullet rows.           */

function slide06(s) {
  rings(s, 15.764, 4.643, 10.896, 10.357, MIST, 36);
  blob(s, 20.314, 9.051, 6.458, 5.928, LIME);

  txt(s, 'About the details of our company.', { x: 2.736, y: 2.319, w: 9.814, h: 1.919, fontSize: 54, bold: true });
  txt(s, SUBTITLE, { x: 2.706, y: 4.385, w: 7.201, h: 0.438, fontSize: 20, bold: true, fontFace: F_SB, color: SLATE });
  txt(s, LOREM_LONG, {
    x: 2.706, y: 5.396, w: 10.086, h: 2.794, fontSize: 20, color: GRAY, transparency: 53, align: 'justify',
  });

  arrowRow(s, 2.736, 8.427, NAVY, WHITE, 'Innovating through design and technology.', 8.198);
  arrowRow(s, 2.736, 9.576, BLUE, WHITE, 'Turning ideas into intelligent solutions.', 7.101);
  arrowRow(s, 2.736, 10.726, BLUE, WHITE, 'Turning ideas into intelligent solutions.', 7.101);
  chrome(s, 6);
}

/* ============================================================ slides 7-9 */
/* "ABOUT US" / "WHY US?" tilted cards.                                    */

// One tilted card: slanted body, white arrow badge, title and body copy.
function tiltCard(s, x, y, fill, title) {
  const w = 11.583;
  const h = 2.897;
  poly(s, CARD_TILT, x, y, w, h, { fill });
  s.addShape('ellipse', { x: x + 9.91, y: y + 1.988, w: 0.908, h: 0.908, fill: WHITE, shadow: DROP });
  poly(s, ARROW_NE, x + 10.22, y + 2.349, 0.297, 0.23, { fill: BLUE, rotate: 323 });
  txt(s, title, { x: x + 0.875, y: y + 0.449, w: 10.574, h: 0.505, fontSize: 24, bold: true, fontFace: F_SB, color: WHITE });
  txt(s, LOREM_CARD, { x: x + 0.875, y: y + 1.135, w: 8.655, h: 1.01, fontSize: 18, color: WHITE, transparency: 25 });
}

// Faint oversized section word behind the cards.
function ghostWord(s, word, x, y, w, h) {
  txt(s, word, Object.assign(unrotate(x, y, w, h), {
    fontSize: 344, bold: true, fontFace: F_EB, color: NAVY, transparency: 97,
  }));
}

function slide07(s) {
  ghostWord(s, 'ABOUT US', 10.809, -9.575, 5.991, 25.182);
  tiltCard(s, 3.263, 3.488, INK, CARD_TITLE);
  rings(s, 14.21, 3.167, 12.45, 11.833, DEEP, 0);
  tiltCard(s, 1.913, 7.113, BLUE, CARD_TITLE);
  poly(s, HOOK, 12.277, 7.105, 1.21, 1.076, { fill: LIME });
  blob(s, 20.314, 9.051, 6.458, 5.928, LIME);
  squiggle(s, 22.303, 2.211, LIME, BLUE);
  heading(s, 1.854, 11.515, 'Our core values of business.', SUBTITLE);
  chrome(s, 7);
}

function slide08(s) {
  rings(s, 15.764, 4.643, 10.896, 10.357, MIST, 36);
  ghostWord(s, 'ABOUT US', 10.809, -9.575, 5.991, 25.182);
  blob(s, 20.314, 9.051, 6.458, 5.928, LIME);
  squiggle(s, 22.303, 2.003, LIME, BLUE);

  tiltCard(s, 1.33, 3.364, INK, CARD_TITLE);
  tiltCard(s, 13.392, 3.368, BLUE, CARD_TITLE);
  poly(s, HOOK, 23.755, 3.36, 1.21, 1.076, { fill: LIME });
  tiltCard(s, 1.372, 6.889, BLUE, CARD_TITLE);
  tiltCard(s, 13.392, 6.889, BLUE, CARD_TITLE);

  txt(s, 'Our core objectives.', { x: 1.33, y: 10.814, w: 13.793, h: 1.01, fontSize: 54, bold: true });
  txt(s, LOREM_WIDE, {
    x: 1.372, y: 12.023, w: 13.566, h: 2.121, fontSize: 20, color: GRAY, transparency: 53, align: 'justify',
  });
  chrome(s, 8);
}

function slide09(s) {
  rings(s, -0.035, 4.934, 10.896, 10.357, MIST, 41, true);
  ghostWord(s, 'WHY US?', 9.953, -8.719, 5.991, 23.47);
  poly(s, RING_SOLID, -0.057, 9.335, 6.431, 5.665, { fill: LIME, flipH: true });

  tiltCard(s, 2.827, 3.384, INK, CARD_TITLE);
  tiltCard(s, 1.765, 7.009, BLUE, 'Proven expertise backed by data-driven results.');
  poly(s, HOOK, 12.138, 6.993, 1.21, 1.076, { fill: LIME });
  tiltCard(s, 2.827, 10.647, BLUE, 'End-to-end support from strategy to execution.');

  squiggle(s, 1.728, 1.571, LIME, BLUE);
  chrome(s, 9);
}

/* =========================================================== slides 10-11 */
/* "SERVICES" pointer cards.                                                */

// Pointer-tab card used for services / team / community panels.
function pointerCard(s, x, y, w, h, fill) {
  poly(s, CARD_SERVICE, x, y, w, h, { fill, flipV: true });
}

// Service card: pointer body, check glyph, two-line title and body copy.
function serviceCard(s, x, y, fill, splitTitle, badgeArrow) {
  pointerCard(s, x, y, 5.563, 7.333, fill);
  poly(s, CHECK_MARK, x + 0.786, y + 0.758, 0.777, 0.569, { fill: WHITE });
  const title = splitTitle
    ? [{ text: 'Wealth' }, { text: ' ', options: { fontFace: F_L } }, { text: 'Manage\nment' }]
    : [{ text: 'Wealth Manage\nment' }];
  txt(s, title, { x: x + 0.573, y: y + 1.917, w: 4.742, h: 1.043, fontSize: 28, fontFace: F_SB, color: WHITE });
  txt(s, [{ text: LOREM_SERVICE, options: { breakLine: true } }, { text: LOREM_SERVICE }], {
    x: x + 0.554, y: y + 3.323, w: 4.448, h: 2.221, fontSize: 18, color: WHITE, transparency: 45,
  });
  arrowBadge(s, x + 3.561, y + 6.146, 1.188, badgeArrow);
}

function slide10(s) {
  ghostWord(s, 'SERVICES', 9.953, -8.719, 5.991, 23.47);
  serviceCard(s, 1.944, 2.519, NAVY, false, BLUE);
  serviceCard(s, 8.385, 3.674, BLUE, true, NAVY);
  rings(s, 15.764, 4.643, 10.896, 10.357, MIST, 36);
  blob(s, 20.314, 9.051, 6.458, 5.928, LIME);

  // Right-hand "Our Community & Reach" panel.
  pointerCard(s, 16.273, 1.929, 8.089, 10.663, DEEP);
  txt(s, 'Our Community \n& Reach', { x: 17.319, y: 3.304, w: 5.998, h: 2.827, fontSize: 54, fontFace: F_SB, color: WHITE });
  txt(s, 'Here goes your shorter sub title in', {
    x: 17.278, y: 6.76, w: 7.201, h: 0.438, fontSize: 20, bold: true, fontFace: F_SB, color: WHITE,
  });
  txt(s, LOREM_PANEL, {
    x: 17.278, y: 7.813, w: 6.41, h: 2.457, fontSize: 20, color: WHITE, transparency: 55, align: 'justify',
  });
  poly(s, HOOK, 12.43, 3.659, 1.519, 1.351, { fill: LIME });
  squiggle(s, 11.464, 2.382, LIME, BLUE);

  txt(s, [
    { text: 'What', options: { bold: true } },
    { text: ' ', options: { fontFace: F_EL } },
    { text: 'We Bring to the Table', options: { bold: true } },
  ], { x: 1.854, y: 11.869, w: 13.793, h: 1.01, fontSize: 54 });
  txt(s, SUBTITLE, { x: 1.854, y: 12.966, w: 9.572, h: 0.505, fontSize: 24, color: SLATE, transparency: 40 });
  chrome(s, 10);
}

function slide11(s) {
  rings(s, 15.764, 4.643, 10.896, 10.357, MIST, 36);
  ghostWord(s, 'SERVICES', 9.953, -8.719, 5.991, 23.47);
  blob(s, 20.293, 9.031, 6.458, 5.928, LIME);

  serviceCard(s, 1.289, 2.619, NAVY, true, BLUE);
  serviceCard(s, 13.648, 2.619, BLUE, true, NAVY);
  serviceCard(s, 19.693, 3.661, BLUE, true, NAVY);
  serviceCard(s, 7.453, 3.661, BLUE, false, NAVY);
  poly(s, HOOK, 11.492, 3.657, 1.519, 1.351, { fill: LIME });
  squiggle(s, 14.558, 12.099, LIME, BLUE);

  txt(s, [
    { text: 'What', options: { bold: true } },
    { text: ' ', options: { fontFace: F_EL } },
    { text: 'We Bring to the Table', options: { bold: true } },
  ], { x: 1.89, y: 11.891, w: 13.793, h: 1.01, fontSize: 54 });
  txt(s, SUBTITLE, { x: 1.89, y: 12.925, w: 9.572, h: 0.505, fontSize: 24, color: SLATE, transparency: 40 });
  chrome(s, 11);
}

/* ============================================================== slide 12 */
/* Team members: four staggered pointer cards with name / job captions.    */

function slide12(s) {
  rings(s, 15.764, 4.643, 10.896, 10.357, MIST, 36);
  ghostWord(s, 'MEMBERS', 10.63, -9.417, 5.991, 24.825);
  blob(s, 20.293, 9.031, 6.458, 5.928, LIME);

  const people = [
    { x: 0.935, y: 4.364, fill: NAVY, hook: false },
    { x: 7.343, y: 2.973, fill: BLUE, hook: true },
    { x: 13.752, y: 4.395, fill: BLUE, hook: false },
    { x: 20.161, y: 3.765, fill: BLUE, hook: false },
  ];
  people.forEach(({ x, y, fill, hook }) => {
    pointerCard(s, x, y, 5.563, 7.333, fill);
    if (hook) poly(s, HOOK_WIDE, x + 3.289, y + 5.896, 1.71, 1.027, { fill: LIME });
    arrowBadge(s, x + 3.56, y + 6.147, 1.188, BLUE);
    txt(s, 'Jonathan Doe', {
      x: x + 0.082, y: y + 7.569, w: 4.625, h: 0.505, fontSize: 24, bold: true, fontFace: F_SB,
    });
    txt(s, 'Job title goes here', {
      x: x + 0.082, y: y + 8.096, w: 4.293, h: 0.404, fontSize: 18, italic: true, color: BLACK, transparency: 47,
    });
  });
  squiggle(s, 8.104, 12.02, LIME, BLUE);
  chrome(s, 12);
}

/* ============================================================== slide 13 */
/* Worker of the month: big split headline plus two progress sliders.      */

function slide13(s) {
  ghostWord(s, 'MEMBERS', 10.63, -9.417, 5.991, 24.825);
  rings(s, 0.002, 3.167, 12.45, 11.833, DEEP, 0, true);
  poly(s, RING_SOLID, 0, 9.075, 6.458, 5.928, { fill: LIME, flipH: true });

  txt(s, 'And your subtitle goes here', { x: 13.222, y: 2.597, w: 8.938, h: 0.505, fontSize: 24, fontFace: F_SB });
  txt(s, [
    { text: 'Best Worker ', options: { fontFace: F_SB } },
    { text: 'of the month', options: { fontFace: F_L } },
  ], { x: 13.16, y: 3.366, w: 13.45, h: 3.972, fontSize: 115 });
  txt(s, [
    { text: 'Jonathan Doe', options: { bold: true } },
    { text: '.', options: { fontSize: 72, fontFace: F_SB } },
  ], { x: 13.201, y: 7.655, w: 10.092, h: 1.313, fontSize: 66 });

  [[9.581, 10.264, 7.508, 20.265, '80%'], [11.524, 12.207, 8.421, 21.638, '90%']].forEach(
    ([labelY, trackY, fillW, knobX, pct]) => {
      txt(s, 'Training performance', {
        x: 13.26, y: labelY, w: 5.926, h: 0.438, fontSize: 20, color: BLACK, transparency: 26,
      });
      round(s, 13.28, trackY, 11.5, 0.438, 0.5, { fill: { color: SILVER, transparency: 83 } });
      round(s, 13.47, trackY + 0.11, fillW, 0.15, 0.5, { fill: BLUE });
      round(s, knobX, trackY - 0.076, 1.104, 0.604, 0.5, { fill: LIME });
      round(s, knobX + 0.263, trackY + 0.166, 0.226, 0.124, 0.5, { fill: BLUE });
      txt(s, pct, { x: 23.4, y: labelY - 0.058, w: 1.341, h: 0.707, fontSize: 36, fontFace: F_SB, align: 'right' });
    },
  );
  squiggle(s, 8.753, 7.802, LIME, BLUE);
  chrome(s, 13);
}

/* ============================================================== slide 14 */
/* Testimonials: two quote cards with 5-star ratings.                       */

function quoteCard(s, x, y, fill, dimStar) {
  poly(s, CARD_QUOTE, x, y, 9.932, 7.117, { fill });
  for (let i = 0; i < 5; i++) {
    s.addShape('star5', {
      x: x + 4.195 + i * 0.316, y: y + 0.895, w: 0.28, h: 0.28,
      fill: i < 4 ? AMBER : { color: dimStar, transparency: 24 },
    });
  }
  txt(s, LOREM_QUOTE, {
    x: x + 1.14, y: y + 1.878, w: 7.653, h: 2.524, fontSize: 24, italic: true, fontFace: F_L,
    color: WHITE, transparency: 25, align: 'center',
  });
  txt(s, 'Jonathan Doe', {
    x: x + 2.098, y: y + 5.43, w: 3.153, h: 0.505, fontSize: 24, bold: true, color: WHITE,
  });
  txt(s, 'Job title goes here', {
    x: x + 2.098, y: y + 5.892, w: 2.927, h: 0.37, fontSize: 16, italic: true, color: WHITE, transparency: 26,
  });
  arrowBadge(s, x + 7.974, y + 5.929, 1.188, BLUE);
}

function slide14(s) {
  rings(s, 15.764, 4.643, 10.896, 10.357, MIST, 36);
  blob(s, 20.314, 9.051, 6.458, 5.928, LIME);
  ghostWord(s, 'CLIENTS', 10.63, -9.271, 5.991, 24.825);

  quoteCard(s, 2.083, 2.981, NAVY, 'BFBFBF');
  quoteCard(s, 14.376, 3.71, BLUE, '154A85');
  poly(s, HOOK, 22.789, 3.71, 1.519, 1.351, { fill: LIME });

  squiggle(s, 21.506, 2.021, LIME, BLUE);
  heading(s, 2.104, 11.68, 'What Our Clients Say ', SUBTITLE, 16.271);
  chrome(s, 14);
}

/* ============================================================== slide 15 */
/* Lightbulb infographic: stacked coil segments with three callouts.       */

function slide15(s) {
  rings(s, 15.764, 4.643, 10.896, 10.357, MIST, 36);
  blob(s, 20.273, 9.051, 6.458, 5.928, LIME);

  poly(s, BULB_COIL, 7.068, 7.509, 3.779, 1.907, { fill: BLUE });
  poly(s, BULB_COIL, 7.068, 6.363, 3.779, 1.907, { fill: DEEP });
  poly(s, BULB_COIL, 7.247, 5.213, 3.779, 1.911, { fill: LIME });
  poly(s, BULB_BODY, 7.068, 8.703, 3.556, 3.162, { fill: INK });
  poly(s, BULB_RING, 8.245, 11.992, 1.426, 0.236, { fill: INK });
  poly(s, BULB_RING, 8.245, 12.355, 1.426, 0.236, { fill: INK });
  poly(s, BULB_RING_TALL, 8.245, 12.718, 1.426, 0.424, { fill: INK });
  poly(s, BULB_CAP, 8.258, 4.5, 2.756, 1.496, { fill: CHARCOAL });

  // Leader lines: a shallow arc from the bulb out to each caption dot.
  const leads = [
    { arc: [10.624, 5.137, 4.468, 4.396], dot: [13.32, 9.336, LIME] },
    { arc: [2.977, 5.384, 4.468, 4.396], dot: [5.672, 9.583, NAVY] },
    { arc: [10.624, 7.495, 4.468, 4.396], dot: [13.32, 11.694, BLUE] },
  ];
  leads.forEach(({ arc, dot }) => {
    s.addShape('arc', {
      x: arc[0], y: arc[1], w: arc[2], h: arc[3], flipV: true,
      angleRange: [286, 334], line: { color: 'BFBFBF', width: 1 },
    });
    s.addShape('ellipse', { x: dot[0], y: dot[1], w: 0.235, h: 0.235, fill: dot[2] });
  });

  ideaLabel(s, 12.779, 4.995, IDEA_TITLE, 'left');
  ideaLabel(s, 12.779, 7.342, IDEA_TITLE, 'left');
  ideaLabel(s, 0.413, 5.246, IDEA_TITLE, 'right');

  heading(s, 1.161, 1.494, 'Our inspiring business solutions', SUBTITLE);
  squiggle(s, 22.475, 1.941, LIME, BLUE);
  chrome(s, 15);
}

/* ============================================================== slide 16 */
/* Four-step timeline: pointer cards with ribbon arrows over a rail.       */

function stepCard(s, x, y, fill, tipFill, ribbonFill, ribbonTip, title, bodyColor, bodyAlpha) {
  pointerCard(s, x, y, 5.563, 7.333, fill);
  poly(s, RIBBON_TIP, x + 2.444, y + 6.441, 0.676, 1.108, { fill: tipFill, rotate: 270, flipH: true, flipV: true, shadow: DROP });
  poly(s, RIBBON_FOLD, x + 3.406, y + 0.867, 0.683, 0.513, { fill: ribbonTip, rotate: 306.6, shadow: DROP });
  poly(s, RIBBON_ARROW_L, x + 1.544, y - 0.212, 1.946, 1.207, { fill: ribbonTip, rotate: 46.1, flipH: true });
  poly(s, RIBBON_ARROW_R, x + 1.483, y - 0.081, 2.502, 1.207, { fill: ribbonFill, rotate: 53.4, flipH: true, shadow: DROP });
  poly(s, PLANE, x + 0.859, y + 0.699, 0.537, 0.439, { fill: WHITE });
  txt(s, title, {
    x: x + 0.421, y: y + 2.848, w: 4.721, h: 0.505, fontSize: 24, bold: true, fontFace: F_SB, color: WHITE, align: 'center',
  });
  txt(s, [{ text: 'Lorem Ipsum', options: { bold: true } }, { text: DUMMY }], {
    x: x + 0.657, y: y + 3.74, w: 4.248, h: 1.111, fontSize: 20, color: bodyColor, transparency: bodyAlpha, align: 'center',
  });
}

function slide16(s) {
  rings(s, 15.764, 4.643, 10.896, 10.357, MIST, 36);
  blob(s, 20.273, 9.051, 6.458, 5.928, LIME);

  // Timeline rail: three stacked hairlines with a chevron node per step.
  round(s, 3.608, 12.479, 19.481, 0.083, 0.5, { fill: SILVER });
  round(s, 3.608, 12.146, 19.481, 0.104, 0.5, { fill: SILVER });
  round(s, 3.608, 12.333, 19.564, 0.062, 0.5, { fill: BLUE });

  const stepX = [1.162, 7.468, 13.747, 20.015];
  const nodeX = [3.608, 9.888, 16.167, 22.447];
  nodeX.forEach((x, i) => {
    const y = i === 0 ? 12.027 : 12.007;
    s.addShape('ellipse', { x, y, w: 0.723, h: 0.723, fill: WHITE });
    s.addShape('ellipse', { x: x - 0.023, y: y - 0.022, w: 0.746, h: 0.746, line: { color: BLUE, width: 2 } });
    poly(s, CHEVRON, x - 0.023, y - 0.022, 0.746, 0.746, { fill: BLUE });
  });

  const titles = ['Step one title here', 'Step two title here', 'Step three title here', 'Step four title here'];
  stepX.forEach((x, i) => {
    const first = i === 0;
    stepCard(
      s, x, first ? 4.253 : 4.274, first ? NAVY : BLUE,
      first ? BLUE : LIME, first ? DEEP : LIME, first ? MIDNIGHT : OLIVE,
      titles[i], first ? MUTED : WHITE, first ? 20 : 44,
    );
  });

  heading(s, 1.161, 1.181, 'Our Path to Performance', SUBTITLE);
  squiggle(s, 22.223, 1.754, LIME, BLUE);
  chrome(s, 16);
}

/* ============================================================== slide 17 */
/* Weekly analysis: seven rounded bars, each labelled with a 20% chip.     */

const WEEK = [
  { day: 'Sun', y: 8.924, h: 2.687 },
  { day: 'Mon', y: 7.374, h: 4.238 },
  { day: 'Tue', y: 5.233, h: 6.378 },
  { day: 'Wed', y: 6.631, h: 4.98 },
  { day: 'Thu', y: 4.447, h: 7.164 },
  { day: 'Fri', y: 3.07, h: 8.541 },
  { day: 'Sat', y: 8.771, h: 2.84 },
];

function slide17(s) {
  rings(s, 15.764, 4.643, 10.896, 10.357, MIST, 49);
  blob(s, 20.273, 9.051, 6.458, 5.928, 'EEF1F5');

  WEEK.forEach((d, i) => {
    const x = 12.551 + i * 1.842;
    const peak = d.day === 'Fri';
    round(s, x, d.y, 1.485, d.h, 0.31373, { fill: peak ? HATCH_LIME : HATCH_BLUE, shadow: peak ? undefined : DROP });
    // Frosted percentage chip pinned near the top of each bar.
    round(s, x + 0.349, d.y + 0.568, 1.311, 0.917, 0.26191, {
      fill: { color: peak ? OLIVE : WHITE, transparency: peak ? 72 : 68 },
      line: { color: 'D2DEF4', width: 2 }, shadow: DROP,
    });
    txt(s, '20%', {
      x: x + 0.539, y: d.y + 0.793, w: 1.114, h: 0.529, fontSize: 24, fontFace: F_SB,
      color: WHITE, align: 'center', lineSpacingMultiple: 1,
    });
    txt(s, d.day, {
      x: 12.638 + i * 1.842, y: 12.179, w: 1.311, h: 0.529, fontSize: 24, color: BLACK,
      transparency: 56, align: 'center',
    });
  });

  txt(s, 'Weekly Analysis and growth Chart', { x: 1.532, y: 2.722, w: 9.814, h: 1.919, fontSize: 54, bold: true });
  txt(s, SUBTITLE, { x: 1.502, y: 4.787, w: 7.201, h: 0.438, fontSize: 20, bold: true, fontFace: F_SB, color: SLATE });
  txt(s, LOREM_LONG, {
    x: 1.502, y: 5.798, w: 9.265, h: 2.121, fontSize: 20, color: GRAY, transparency: 53, align: 'justify',
  });

  round(s, 1.502, 8.51, 3.214, 3.269, 0.23943, { fill: DEEP });
  txt(s, '90%', { x: 1.827, y: 8.88, w: 2.888, h: 1.212, fontSize: 66, fontFace: F_L, color: WHITE });
  txt(s, 'Retention rate', {
    x: 1.877, y: 10.069, w: 2.905, h: 0.438, fontSize: 20, fontFace: F_SB, color: WHITE, transparency: 33,
  });
  poly(s, ARROW_NE, 3.825, 11.035, 0.396, 0.307, { fill: WHITE, rotate: 320.74, flipV: true });
  chrome(s, 17);
}

/* ============================================================== slide 18 */
/* Rocket infographic: six callouts orbiting a launching rocket.           */

function slide18(s) {
  rings(s, 15.764, 4.643, 10.896, 10.357, MIST, 80);
  blob(s, 20.314, 9.051, 6.458, 5.928, LIME);

  // Moons, bottom-left and top-right, with a few crater dots each.
  const moons = [
    { x: 0.41, y: 9.433, d: 6.816, craters: [[2.103, 10.101, 1.597], [3.004, 12.435, 0.911], [1.421, 13.669, 0.911], [4.36, 10.988, 0.52], [1.876, 11.938, 0.52], [1.254, 11.003, 0.52]] },
    { x: 21.648, y: 2.935, d: 1.874, craters: [[22.113, 3.119, 0.439], [22.361, 3.76, 0.25], [21.925, 4.1, 0.25], [22.734, 3.363, 0.143], [22.051, 3.624, 0.143], [21.88, 3.367, 0.143]] },
  ];
  moons.forEach((m) => {
    s.addShape('ellipse', { x: m.x, y: m.y, w: m.d, h: m.d, fill: 'F4F4F4' });
    m.craters.forEach(([cx, cy, cd]) => s.addShape('ellipse', { x: cx, y: cy, w: cd, h: cd, fill: 'EDEDED' }));
  });

  // Cloud bank across the bottom.  The original is one silhouette filled with a
  // vertical blue-to-haze gradient; approximate it with the puffy silhouette in
  // the top colour plus lighter full-width bands below the last white gap.
  poly(s, CLOUD, -1.143, 11.0, 29.047, 4.0, { fill: { color: '4E81D0' } });
  [['8DACDD', 13.95], ['9AB6E0', 14.35], ['A5BDE2', 14.7]].forEach(([c, y]) => {
    s.addShape('rect', { x: -0.5, y, w: W + 1, h: H - y, fill: c, line: { color: c, width: 0.5 } });
  });

  poly(s, HALO, 9.502, 3.939, 7.5, 6.778, { fill: { color: SILVER, transparency: 51 } });
  poly(s, ROCKET_FIN, 11.871, 6.489, 2.717, 2.106, { fill: BLUE, shadow: DROP });
  poly(s, ROCKET_FIN, 10.713, 7.384, 5.038, 3.905, { fill: BLUE, shadow: DROP });
  poly(s, ROCKET_BODY, 12.644, 5.196, 1.176, 6.041, { fill: 'F2F6FE' });
  poly(s, ROCKET_NOSE, 12.686, 5.196, 1.052, 0.966, { fill: DEEP });
  poly(s, ROCKET_FLAME, 12.889, 8.542, 0.695, 2.489, { fill: BLUE, rotate: 1.3 });
  poly(s, ROCKET_BODY, 12.212, 8.063, 0.515, 3.103, { fill: 'E4EBF8' });
  poly(s, ROCKET_BODY, 13.737, 8.063, 0.515, 3.103, { fill: 'E4EBF8' });
  s.addShape('ellipse', { x: 14.717, y: 6.415, w: 1.014, h: 0.697, fill: 'CCDBF2', rotate: 348.9 });
  s.addShape('ellipse', { x: 11.812, y: 8.457, w: 1.358, h: 0.697, fill: 'CCDBF2', rotate: 328.8 });
  round(s, 12.835, 11.173, 0.794, 0.247, 0.5, { fill: LIME });
  round(s, 13.115, 11.4, 0.23, 3.6, 0.5, { fill: { color: 'A8BEEC' } });

  // Six numbered nodes; each is a ring badge on the end of a hairline arc.
  const nodes = [
    { x: 15.07, y: 4.534, fill: BLUE, arc: [15.23, 3.793, false, true], label: [17.197, 3.472, 'left'] },
    { x: 10.555, y: 4.534, fill: BLUE, arc: [6.65, 3.793, false, false], label: [3.985, 3.472, 'right'] },
    { x: 9.21, y: 7.328, fill: BLUE, line: [7.751, 7.68, false], label: [2.236, 7.432, 'right'] },
    { x: 16.408, y: 7.328, fill: BLUE, line: [17.24, 7.68, true], label: [19.139, 7.432, 'left'] },
    { x: 14.845, y: 10.289, fill: LIME, arc: [15.005, 7.411, true, true], label: [17.129, 11.523, 'left'] },
    { x: 10.872, y: 10.31, fill: BLUE, arc: [7.022, 7.432, true, false], label: [4.334, 11.523, 'right'] },
  ];
  nodes.forEach((n) => {
    if (n.arc) {
      s.addShape('arc', {
        x: n.arc[0], y: n.arc[1], w: 4.468, h: 4.396, flipV: n.arc[2], flipH: n.arc[3],
        angleRange: [286, 322], line: { color: BLUE, width: 1 },
      });
    }
    if (n.line) {
      s.addShape('line', { x: n.line[0], y: n.line[1], w: 1.493, h: 0, line: { color: BLUE, width: 1, dashType: 'sysDash' } });
    }
    s.addShape('ellipse', { x: n.x, y: n.y, w: 0.777, h: 0.777, fill: n.fill });
    s.addShape('ellipse', { x: n.x + 0.16, y: n.y + 0.15, w: 0.457, h: 0.457, fill: WHITE, shadow: DROP });
    poly(s, CHEVRON, n.x + 0.197, n.y + 0.188, 0.382, 0.382, { fill: BLUE });
    ideaLabel(s, n.label[0], n.label[1], IDEA_TITLE, n.label[2]);
  });

  heading(s, 1.161, 0.952, 'Launching Ideas into Impact', SUBTITLE);
  squiggle(s, 22.475, 1.399, LIME, BLUE);
  chrome(s, 18);
}

/* ============================================================== slide 19 */
/* Strategic pillars: four icon circles feeding into a layered arch.       */

function slide19(s) {
  poly(s, SWOOSH_LONG, 15.491, 9.191, 1.055, 2.029, { fill: LIME, flipH: true });
  poly(s, SWOOSH_LONG, 10.166, 9.191, 1.055, 2.029, { fill: BLUE });
  poly(s, SWOOSH_SHORT, 13.476, 8.167, 2.014, 1.037, { fill: LIME, flipH: true });
  poly(s, SWOOSH_SHORT, 11.206, 8.166, 2.014, 1.037, { fill: BLUE });

  // Icon nodes: coloured ring, white core, dotted leader to the caption dot.
  const pillars = [
    { x: 17.88, y: 8.427, ring: LIME, lead: [16.353, 9.668, 1.179, 0.356, 'E5E9EF', true], label: [20.791, 8.715, 'left'] },
    { x: 6.534, y: 8.427, ring: BLUE, lead: [9.179, 9.668, 1.179, 0.356, 'BFBFBF', true], label: [0.918, 8.715, 'right'] },
    { x: 14.436, y: 4.714, ring: BLUE, lead: [14.505, 7.109, 0.531, 1.339, 'BFBFBF', false], label: [17.249, 4.877, 'left'] },
    { x: 10.057, y: 4.683, ring: BLUE, lead: [11.661, 7.108, 0.531, 1.339, 'BFBFBF', false], label: [4.489, 4.877, 'right'] },
  ];
  pillars.forEach((p) => {
    const [lx, ly, lw, lh, lc] = p.lead;
    s.addShape('line', { x: lx, y: ly, w: lw, h: lh, flipV: true, line: { color: lc, width: 2.25, dashType: 'dash' } });
    s.addShape('ellipse', { x: p.x, y: p.y, w: 2.299, h: 2.299, fill: p.ring });
    s.addShape('ellipse', { x: p.x + 0.206, y: p.y + 0.207, w: 1.886, h: 1.886, fill: WHITE, shadow: LIFT });
    ideaLabel(s, p.label[0], p.label[1], IDEA_TITLE, p.label[2]);
  });

  // Icon glyphs sitting inside the four white cores.
  poly(s, ICON_SCALES, 15.304, 5.646, 0.537, 0.55, { fill: BLACK });        // scales, top right
  poly(s, ICON_CLOUD_DOWN, 10.875, 5.595, 0.661, 0.56, { fill: BLACK });    // cloud download, top left
  poly(s, ICON_DOWNLOAD, 18.83, 9.353, 0.372, 0.496, { fill: BLACK });      // download, right
  poly(s, ICON_CLOUD_UP, 7.367, 9.325, 0.661, 0.56, { fill: BLACK });       // cloud upload, left

  // Centre teardrop badge with the tick glyph.
  s.addShape('teardrop', { x: 12.575, y: 10.52, w: 1.562, h: 1.562, fill: BLUE, rotate: 21.3, shadow: DROP });
  s.addShape('ellipse', { x: 12.77, y: 10.726, w: 1.156, h: 1.156, fill: WHITE, shadow: DROP });
  s.addShape('ellipse', { x: 12.908, y: 10.868, w: 0.876, h: 0.876, line: { color: BLACK, width: 2 } });
  poly(s, TICK, 12.908, 10.868, 0.876, 0.876, { fill: BLACK });

  // Layered arches anchored below the slide edge.
  poly(s, ARCH, 10.751, 8.765, 5.255, 6.36, { fill: LIME });
  s.addShape('blockArc', { x: 13.602, y: 12.539, w: 5.242, h: 5.242, fill: BLUE, angleRange: [180, 1.3], arcThicknessRatio: 0.26 });
  s.addShape('blockArc', { x: 14.63, y: 13.532, w: 3.187, h: 3.187, fill: NAVY, angleRange: [180, 1.3], arcThicknessRatio: 0.26 });

  heading(s, 1.161, 1.735, 'Strategic Pillars of Success', SUBTITLE);
  squiggle(s, 22.475, 2.006, LIME, BLUE);
  chrome(s, 19);
}

/* ============================================================== slide 20 */
/* Foundation pyramid: three stacked funnel layers with three callouts.    */

function slide20(s) {
  rings(s, 15.764, 4.643, 10.896, 10.357, MIST, 71);
  blob(s, 20.273, 9.051, 6.458, 5.928, LIME);

  poly(s, FUNNEL_BASE, 6.451, 8.777, 10.618, 3.689, { fill: BLUE, shadow: SOFT });
  poly(s, FUNNEL_MID, 8.119, 7.098, 7.284, 3.014, { fill: NAVY });
  poly(s, FUNNEL_TOP, 9.907, 5.484, 3.735, 2.117, { fill: LIME });
  s.addShape('pie', { x: 11.217, y: 4.369, w: 1.115, h: 1.115, fill: '8496B0', angleRange: [42, 138], shadow: DROP });

  // Three white chevron plates sitting on the layer fronts.
  [[9.233, BLUE], [6.722, NAVY], [4.294, LIME]].forEach(([y, c], i) => {
    poly(s, RIBBON_TIP, 11.209, y, 1.131, 1.852, { fill: WHITE, rotate: 270, shadow: DROP });
    poly(s, RIBBON_TIP, 11.437, y + 0.288, 0.789, 1.293, { fill: c, rotate: 270 });
  });

  // Callout nodes anchored to each layer.
  const marks = [
    { x: 8.478, y: 7.373, ring: NAVY, arc: [4.573, 6.632, false], label: [1.878, 6.431, 'right'], title: 'Process & Performance' },
    { x: 15.915, y: 9.084, ring: BLUE, arc: [16.075, 8.343, true], label: [18.227, 8.15, 'left'], title: 'People & Culture' },
    { x: 12.594, y: 5.763, ring: LIME, arc: [12.754, 5.022, true], label: [14.978, 4.755, 'left'], title: 'Innovation & Leadership' },
  ];
  marks.forEach((m) => {
    s.addShape('arc', {
      x: m.arc[0], y: m.arc[1], w: 4.468, h: 4.396, flipH: m.arc[2],
      angleRange: [286, 322], line: { color: BLUE, width: 1 },
    });
    s.addShape('ellipse', { x: m.x, y: m.y, w: 0.777, h: 0.777, fill: m.ring });
    s.addShape('ellipse', { x: m.x + 0.17, y: m.y + 0.15, w: 0.457, h: 0.457, fill: WHITE, shadow: DROP });
    poly(s, CHEVRON, m.x + 0.208, m.y + 0.188, 0.382, 0.382, { fill: BLUE });
    ideaLabel(s, m.label[0], m.label[1], m.title, m.label[2]);
  });

  heading(s, 1.161, 1.494, 'Foundation for Lasting Growth', SUBTITLE);
  squiggle(s, 22.475, 1.941, LIME, BLUE);
  chrome(s, 20);
}

/* ------------------------------------------------------------------ main */

const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'ICY', width: W, height: H });
  pptx.layout = 'ICY';
  pptx.theme = { headFontFace: F, bodyFontFace: F };
  pptx.title = 'ICY Theme deck';

  BUILDERS.forEach((builder) => {
    const slide = pptx.addSlide();
    slide.background = { color: WHITE };
    builder(slide);
  });

  return pptx.writeFile({
    fileName: path.join(__dirname, '0b088c09-3a2a-4031-9f2e-a2d62009aa44_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => {
  console.error(e);
  process.exit(1);
});
