#!/usr/bin/env node
/**
 * "Paws." - Animal Care Presentation (35 slides, 13.333 x 7.5 in)
 * Rebuilt with pptxgenjs.  Photographs in the original deck are flat
 * placeholder blocks, so they are redrawn here as PHOTO-coloured shapes.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const TEAL = '2EC4B6', MINT = 'CBF3F0', SAND = 'FFBF69', ORANGE = 'FF9F1C';
const AMBER = 'FCA311', GOLD = 'FFC000', HONEY = 'FFBD3B', SUN = 'FFCA3A';
const SILVER = 'D4D7E0', CLOUD = 'D0CECE', FOG = 'E7E6E6', PHOTO = 'C4D4DC';
const WHITE = 'FFFFFF', SNOW = 'F2F2F2', IVORY = 'F3F9E9';
const INK = '3B3838', DARK = '262626', SLATE = '595959', GREY = '808080';
const COAL = '404040', CHARCOAL = '333333', PITCH = '1C1C1C', STONE = '767171';

/* -------------------------------------------------------------------- fonts */
const PL = 'Poppins Light', PR = 'Poppins', PM = 'Poppins Medium';
const PS = 'Poppins SemiBold';
const OL = 'Oxygen Light', OR = 'Oxygen', OB = 'Oxygen bold';

/* ------------------------------------------------------------ text presets */
const st = (face, size, color, lineSpacingMultiple) =>
  lineSpacingMultiple ? { fontFace: face, fontSize: size, color, lineSpacingMultiple }
                      : { fontFace: face, fontSize: size, color };

const BODY   = st(PL, 11, GREY, 1.5);   // running paragraph copy
const BODY_W = st(PL, 11, WHITE, 1.5);  // ...on a coloured panel
const NOTE   = st(PL, 11, GREY);        // one-line caption
const SMALL  = st(PL, 12, GREY);
const H2     = st(PM, 22, SLATE);       // slide headline
const H2_D   = st(PM, 22, DARK);
const LEAD   = st(PM, 12, INK);         // sub-heading above body copy
const LEAD_D = st(PM, 12, DARK);
const LEAD_W = st(PM, 12, WHITE);

/* ------------------------------------------------------ geometry primitives */
const KAPPA = 0.5523;      // cubic-bezier circle constant
const CARD_R = 0.132;      // corner radius shared by the deck's cards & photos

const paint = (fill) => Array.isArray(fill)
  ? { color: fill[0], transparency: fill[1] } : { color: fill };
const NO_LINE = { type: 'none' };

function shape(s, kind, x, y, w, h, fill, extra) {
  s.addShape(kind, Object.assign({ x, y, w, h, fill: paint(fill), line: NO_LINE }, extra));
}
const rect  = (s, x, y, w, h, fill, extra) => shape(s, 'rect', x, y, w, h, fill, extra);
const oval  = (s, x, y, w, h, fill, extra) => shape(s, 'ellipse', x, y, w, h, fill, extra);
const rrect = (s, x, y, w, h, fill, r, extra) =>
  shape(s, 'roundRect', x, y, w, h, fill, Object.assign({ rectRadius: r }, extra));
const pill = (s, x, y, w, h, fill, extra) =>
  rrect(s, x, y, w, h, fill, Math.min(w, h) / 2, extra);

/** Free-form outline. Points: ['M',x,y] move, [x,y] line, [x1,y1,x2,y2,x,y] cubic, [] close. */
function poly(s, x, y, w, h, fill, pts, dropShadow) {
  const points = pts.map((p) => {
    if (p.length === 0) return { close: true };
    if (p[0] === 'M') return { x: p[1], y: p[2], moveTo: true };
    if (p.length === 6) return { x: p[4], y: p[5], curve: { type: 'cubic', x1: p[0], y1: p[1], x2: p[2], y2: p[3] } };
    return { x: p[0], y: p[1] };
  });
  s.addShape('custGeom', {
    x, y, w, h, fill: paint(fill), line: NO_LINE, points, shadow: dropShadow,
  });
}

/** Rectangle with only the named corners rounded, e.g. softRect(s, .., 'tl bl'). */
function softRect(s, x, y, w, h, fill, corners, r) {
  const rad = r || CARD_R;
  const c = (k) => (corners.indexOf(k) >= 0 ? rad : 0);
  const [tl, tr, br, bl] = ['tl', 'tr', 'br', 'bl'].map(c);
  const k = KAPPA;
  poly(s, x, y, w, h, fill, [
    ['M', tl, 0], [w - tr, 0], [w - tr + tr * k, 0, w, tr - tr * k, w, tr],
    [w, h - br], [w, h - br + br * k, w - br + br * k, h, w - br, h],
    [bl, h], [bl - bl * k, h, 0, h - bl + bl * k, 0, h - bl],
    [0, tl], [0, tl - tl * k, tl - tl * k, 0, tl, 0], [],
  ]);
}

/** Stand-in for a photograph in the source deck (a flat grey-blue block). */
const photo = (s, x, y, w, h, corners, r) => corners
  ? softRect(s, x, y, w, h, PHOTO, corners, r)
  : rect(s, x, y, w, h, PHOTO);

const hline = (s, x, y, w, color, width) =>
  s.addShape('line', { x, y, w, h: 0, line: { color, width: width || 1 } });
const vline = (s, x, y, h, color, width) =>
  s.addShape('line', { x, y, w: 0, h, line: { color, width: width || 1 } });
const pieSlice = (s, x, y, w, h, fill, from, to) =>
  s.addShape('pie', { x, y, w, h, fill: paint(fill), line: NO_LINE, angleRange: [from, to] });
const ringArc = (s, x, y, w, h, from, to, color, width, endCap) =>
  s.addShape('arc', {
    x, y, w, h, angleRange: [from, to],
    line: { color, width: width || 1, endArrowType: endCap },
  });

/** Soft drop shadow, as used on the deck's cards and floating pills. */
const shadow = (opacity, blur, offset) =>
  ({ type: 'outer', color: '000000', opacity, blur, offset, angle: 90 });

function text(s, x, y, w, h, body, style, extra) {
  s.addText(body, Object.assign({ x, y, w, h, margin: 0, valign: 'top' }, style, extra));
}

/* -------------------------------------------------------- recurring motifs */
/** The teal / sand / silver dot trio that heads most slides. */
function dots(s, x, y) {
  [TEAL, SAND, SILVER].forEach((color, i) => oval(s, x + i * 0.2535, y, 0.099, 0.099, color));
}

/** The 2x2 dot cluster bleeding off the right edge of every layout. */
function edgeDots(s) {
  [[12.91, 3.627, SILVER], [13.087, 3.627, TEAL],
   [12.91, 3.804, ORANGE], [13.087, 3.804, SILVER]]
    .forEach(([x, y, color]) => oval(s, x, y, 0.07, 0.07, color));
}

/** The white rounded card every slide sits on. */
function card(s) {
  rrect(s, 0.6, 0.6, 12.133, 6.3, WHITE, CARD_R, { shadow: shadow(0.1, 15, 5) });
}

/** Page number, bottom-left off the card (the deck keeps it mid-height). */
function pageNo(s, n) {
  text(s, 0.132, 3.648, 0.332, 0.204, n < 10 ? '0' + n : String(n),
    st(PS, 10, COAL), { align: 'center', valign: 'middle' });
}

/* ------------------------------------------------------------- backdrops */
/* One entry per slide layout in the source deck: the tinted bleed shapes
 * behind the card, the card itself and the edge dots. */
const BACKDROPS = {
  2(s) {
    rect(s, 6.667, 0, 6.667, 7.5, [TEAL, 90]);
    rect(s, 0, 0, 6.667, 7.5, [SAND, 90]);
    edgeDots(s);
    card(s);
  },
  3(s) {
    rect(s, 5.467, 0, 2.4, 2.4, [ORANGE, 90]);
    card(s);
    edgeDots(s);
  },
  4(s) {
    rect(s, 0, 4.766, 4.267, 2.734, [TEAL, 90]);
    rect(s, 3.667, 6.9, 1.2, 0.6, SAND);
    card(s);
    softRect(s, 0.6, 4.766, 3.667, 2.134, TEAL, 'bl');
    rect(s, 3.667, 6.3, 1.2, 0.6, SAND);
    edgeDots(s);
    hline(s, 9.667, 2.892, 2.465, SILVER);
    hline(s, 9.667, 4.595, 2.465, SILVER);
  },
  5(s) {
    rect(s, 0, 2.835, 13.333, 1.829, [SILVER, 80]);
    card(s);
    rect(s, 4.644, 0.601, 4.047, 6.299, TEAL);
    hline(s, 1.2, 2.835, 10.933, SAND);
    hline(s, 1.2, 4.652, 10.933, SAND);
    hline(s, 4.644, 2.835, 4.044, WHITE);
    hline(s, 4.644, 4.652, 4.044, WHITE);
    edgeDots(s);
  },
  6(s) {
    card(s);
    edgeDots(s);
    rect(s, 1.2, 4.766, 3.75, 2.734, [TEAL, 90]);
    rect(s, 12.133, 1.2, 1.2, 3.75, [SAND, 90]);
    card(s);
    rect(s, 12.133, 1.2, 0.6, 3.75, SAND);
    rect(s, 0.6, 1.2, 0.6, 3.75, TEAL);
  },
  7(s) {
    card(s);
    edgeDots(s);
    rect(s, 4.867, 0, 0.6, 1.207, [TEAL, 90]);
    rect(s, 6.667, 5.7, 1.2, 1.2, SAND);
    vline(s, 6.667, 1.2, 3.9, CLOUD);
    rect(s, 4.867, 0.6, 0.6, 1.207, TEAL);
  },
  8(s) {
    card(s);
    edgeDots(s);
    rect(s, 12.733, 2.85, 0.6, 1.8, [TEAL, 90]);
    card(s);
    rect(s, 12.133, 2.85, 0.6, 1.8, TEAL);
    rect(s, 1.2, 5.7, 5.467, 0.6, SAND);
    rect(s, 5.467, 5.7, 1.2, 0.6, TEAL);
  },
  9(s) {
    rect(s, 7.867, 0.005, 4.267, 1.2, [TEAL, 90], { flipH: true });
    rect(s, 5.467, 0.005, 2.4, 1.2, [SAND, 90]);
    card(s);
    edgeDots(s);
    rect(s, 5.467, 0.6, 6.667, 1.2, TEAL, { flipH: true });
    rect(s, 5.467, 0.6, 2.4, 1.2, SAND);
  },
  10(s) {
    card(s);
    edgeDots(s);
    vline(s, 1.2, 1.2, 5.1, CLOUD);
    hline(s, 7.267, 4.678, 4.783, SILVER);
  },
  11(s) {
    rect(s, 0, 1.2, 0.6, 2.55, [SILVER, 80]);
    rect(s, 0, 3.75, 0.6, 0.6, [TEAL, 90]);
    card(s);
    edgeDots(s);
    rect(s, 0.6, 3.75, 6.067, 0.6, TEAL);
  },
  12(s) {
    rect(s, 0, 4.9, 9.133, 2.6, [SILVER, 80]);
    rect(s, 9.133, 4.9, 4.2, 2.6, [TEAL, 90]);
    card(s);
    edgeDots(s);
    poly(s, 9.133, 4.9, 3.6, 2, TEAL, [['M', 0, 0], [3.6, 0], [3.6, 0.499], [3.6, 0.499], [3.6, 1.868], [3.6, 1.941, 3.541, 2, 3.468, 2], [0, 2], [0, 0.499], [0, 0.499], []]);
    hline(s, 1.2, 4.9, 7.933, SAND);
  },
  13(s) {
    rect(s, 6.667, 0, 6.667, 7.5, [SILVER, 80]);
    rect(s, 6.667, 6.911, 1.2, 0.583, [TEAL, 90]);
    card(s);
    edgeDots(s);
    softRect(s, 0.6, 5.7, 6.067, 1.2, SAND, 'bl');
  },
  14(s) {
    card(s);
    edgeDots(s);
    rect(s, 8.833, 0, 3.3, 1.8, [TEAL, 90]);
    rect(s, 5.533, 5.7, 3.3, 1.8, ORANGE);
    card(s);
    rect(s, 8.833, 0.6, 3.3, 1.2, TEAL);
    rect(s, 5.533, 5.7, 3.3, 1.2, SAND);
  },
  15(s) {
    card(s);
    edgeDots(s);
    rect(s, 0, 3.75, 4.8, 3.75, [SAND, 90]);
    softRect(s, 0.6, 3.75, 4.2, 3.15, SAND, 'bl');
  },
  16(s) {
    rect(s, 9.067, 0, 4.267, 3, [TEAL, 90]);
    card(s);
    edgeDots(s);
    poly(s, 9.067, 0.6, 3.667, 2.4, TEAL, [['M', 0, 0], [0.6, 0], [1.06, 0], [3.535, 0], [3.608, 0, 3.667, 0.059, 3.667, 0.132], [3.667, 2.4], [1.06, 2.4], [0.6, 2.4], [0, 2.4], []]);
  },
  17(s) {
    rect(s, 12.733, 2.775, 0.6, 1.95, [SILVER, 80]);
    card(s);
    edgeDots(s);
  },
  18(s) {
    card(s);
    edgeDots(s);
    vline(s, 7.252, 1.2, 2.35, CLOUD);
    text(s, 5.792, 4.974, 2.35, 0.303, [{ text: '400+ ', options: { fontFace: OB } }, { text: 'Experience' }], st(OR, 12, WHITE), { align: 'center', rotate: 90, margin: [0, 0, 3.6, 3.6] });
    rect(s, 6.667, 3.95, 0.6, 2.35, SAND);
  },
  19(s) {
    card(s);
    edgeDots(s);
    rect(s, 6.667, 1.2, 0.6, 2.35, TEAL);
    rect(s, 4.133, 3.95, 0.6, 2.35, SAND);
  },
  20(s) {
    rect(s, 0.005, 5.601, 0.592, 0.599, [TEAL, 90], { flipH: true });
    rect(s, -0.003, 1.199, 0.6, 4.401, [SILVER, 80], { flipH: true });
    card(s);
    edgeDots(s);
    rect(s, 0.597, 5.601, 4.273, 0.599, TEAL);
  },
  21(s) {
    card(s);
    edgeDots(s);
    rect(s, 7.867, 4.873, 0.4, 1.433, ORANGE);
    rect(s, 7.867, 1.2, 0.4, 1.433, TEAL);
    vline(s, 5.067, 3.033, 1.433, CLOUD);
  },
  22(s) {
    rect(s, 1.2, 6.9, 3.628, 0.6, [SAND, 90]);
    card(s);
    edgeDots(s);
  },
  23(s) {
    card(s);
    edgeDots(s);
    rect(s, 1.2, 0, 3.628, 0.6, [TEAL, 90]);
    rect(s, 5.433, 6.3, 3.033, 0.6, [SAND, 30]);
  },
  24(s) {
    rect(s, 5.467, 6.9, 2.4, 0.6, [TEAL, 90], { flipV: true });
    card(s);
    edgeDots(s);
  },
  25(s) {
    rect(s, 0, 1.2, 13.333, 5.1, [SILVER, 90]);
    card(s);
    edgeDots(s);
    rect(s, 7.267, 1.2, 0.6, 5.1, TEAL, { flipH: true });
  },
  26(s) {
    card(s);
    edgeDots(s);
    rrect(s, 4.034, 0, 2.632, 7.5, [SILVER, 80], 0);
    rect(s, 4.034, 1.565, 2.632, 4.37, TEAL, { shadow: shadow(0.1, 10, 5) });
    rect(s, 1.2, 1.565, 2.632, 4.37, ORANGE, { shadow: shadow(0.1, 10, 5) });
  },
  27(s) {
    card(s);
    edgeDots(s);
    rect(s, 6.766, 2.1, 2.585, 3.3, WHITE, { shadow: shadow(0.1, 10, 5) });
    rect(s, 9.549, 2.1, 2.585, 3.3, WHITE, { shadow: shadow(0.1, 10, 5) });
    rect(s, 3.983, 5.1, 2.585, 0.8, SAND, { shadow: shadow(0.1, 10, 5) });
    rect(s, 3.983, 1.6, 2.585, 3.3, ORANGE, { shadow: shadow(0.1, 10, 5) });
  },
  28(s) {
    card(s);
    edgeDots(s);
    oval(s, 4.117, 1.2, 5.1, 5.1, [SILVER, 90]);
    oval(s, 4.901, 1.985, 3.531, 3.531, [SAND, 50]);
    poly(s, 5.415, 1.2, 2.503, 5.1, CHARCOAL, [['M', 0.319, 0], [2.184, 0], [2.352, 0, 2.489, 0.137, 2.489, 0.305], [2.489, 1.307], [2.489, 1.307], [2.497, 1.307, 2.503, 1.313, 2.503, 1.32], [2.503, 1.898], [2.503, 1.905, 2.497, 1.911, 2.489, 1.911], [2.489, 1.911], [2.489, 4.795], [2.489, 4.963, 2.352, 5.1, 2.184, 5.1], [0.319, 5.1], [0.151, 5.1, 0.014, 4.963, 0.014, 4.795], [0.014, 2.042], [0.014, 2.042], [0.006, 2.042, 0, 2.035, 0, 2.028], [0, 1.677], [0, 1.67, 0.006, 1.664, 0.014, 1.664], [0.014, 1.664], [0.014, 1.55], [0.014, 1.55], [0.006, 1.55, 0, 1.544, 0, 1.536], [0, 1.186], [0, 1.178, 0.006, 1.172, 0.014, 1.172], [0.014, 1.172], [0.014, 0.989], [0.012, 0.989], [0.005, 0.989, 0, 0.984, 0, 0.977], [0, 0.814], [0, 0.808, 0.005, 0.802, 0.012, 0.802], [0.014, 0.802], [0.014, 0.305], [0.014, 0.137, 0.151, 0, 0.319, 0], []]);
    softRect(s, 6.499, 1.334, 0.336, 0.024, STONE, 'tl tr br bl', 0.012);
    oval(s, 6.897, 1.314, 0.065, 0.064, STONE);
  },
  29(s) {
    rect(s, 0, 0, 6.667, 7.5, [TEAL, 90]);
    card(s);
    edgeDots(s);
    poly(s, 1.2, 1.2, 3.663, 5.093, CHARCOAL, [['M', 0.241, 0.008], [3.143, 0.008], [3.143, 0.008], [3.143, 0.004, 3.147, 0, 3.151, 0], [3.388, 0], [3.392, 0, 3.396, 0.004, 3.396, 0.008], [3.396, 0.008], [3.414, 0.008], [3.547, 0.008, 3.655, 0.116, 3.655, 0.249], [3.655, 0.411], [3.655, 0.411], [3.659, 0.411, 3.663, 0.415, 3.663, 0.419], [3.663, 0.613], [3.663, 0.617, 3.659, 0.621, 3.655, 0.621], [3.655, 0.621], [3.655, 0.661], [3.655, 0.661], [3.659, 0.661, 3.663, 0.664, 3.663, 0.669], [3.663, 0.862], [3.663, 0.866, 3.659, 0.87, 3.655, 0.87], [3.655, 0.87], [3.655, 4.853], [3.655, 4.985, 3.547, 5.093, 3.414, 5.093], [0.241, 5.093], [0.108, 5.093, 0, 4.985, 0, 4.853], [0, 0.249], [0, 0.116, 0.108, 0.008, 0.241, 0.008], []], shadow(0.2, 10, 10));
    vline(s, 6.667, 1.2, 5.093, CLOUD);
  },
  30(s) {
    rect(s, 6.667, 0, 6.667, 7.5, [SAND, 90]);
    card(s);
    edgeDots(s);
    vline(s, 7.416, 2.212, 1.338, WHITE, 0.5);
    poly(s, 7.178, 1.2, 4.444, 3.08, DARK, [['M', 0.15, 0], [4.294, 0], [4.376, 0, 4.444, 0.067, 4.444, 0.15], [4.444, 2.93], [4.444, 3.013, 4.376, 3.08, 4.294, 3.08], [0.15, 3.08], [0.068, 3.08, 0, 3.013, 0, 2.93], [0, 0.15], [0, 0.067, 0.068, 0, 0.15, 0], [], ['M', 0.141, 0.185], [4.303, 0.185], [4.303, 2.746], [0.141, 2.746], [0.141, 0.185], []]);
    softRect(s, 7.178, 1.2, 4.444, 2.909, CHARCOAL, 'tl tr br bl', 0.081);
    poly(s, 6.667, 4.184, 5.467, 0.16, CHARCOAL, [['M', 0, 0], [5.467, 0], [5.467, 0.099], [5.422, 0.127, 5.333, 0.16, 5.279, 0.16], [2.733, 0.16], [0.188, 0.16], [0.134, 0.16, 0.045, 0.127, 0, 0.099], [0, 0], []]);
    poly(s, 6.667, 4.283, 5.467, 0.061, DARK, [['M', 0, 0], [2.733, 0], [5.467, 0], [5.422, 0.028, 5.333, 0.061, 5.279, 0.061], [2.733, 0.061], [0.188, 0.061], [0.134, 0.061, 0.045, 0.028, 0, 0], []]);
    oval(s, 9.367, 1.26, 0.066, 0.066, PITCH);
    oval(s, 9.384, 1.277, 0.033, 0.033, PITCH);
    softRect(s, 9.018, 4.184, 0.764, 0.062, DARK, 'br bl', 0.062);
  },
  31(s) {
    rect(s, 8.464, 6.9, 3.676, 0.6, [SAND, 90]);
    card(s);
    edgeDots(s);
    rect(s, 8.464, 6.3, 3.676, 0.6, SAND);
  },
  32(s) {
    rect(s, 5.533, 0, 3.3, 0.6, [TEAL, 90]);
    rect(s, 8.833, 6.9, 3.306, 0.6, [SAND, 90]);
    card(s);
    edgeDots(s);
  },
  33(s) {
    card(s);
    edgeDots(s);
  },
  34(s) {
    rect(s, 0, 5.7, 6.667, 1.8, [TEAL, 90]);
    card(s);
    edgeDots(s);
    softRect(s, 0.6, 5.7, 6.067, 1.2, TEAL, 'bl', 0.102);
  },
  35(s) {
    rect(s, 1.2, 1.2, 5.467, 6.3, [ORANGE, 90]);
    rect(s, 6.667, 0, 5.467, 1.2, [TEAL, 90]);
    card(s);
    edgeDots(s);
    rect(s, 6.667, 0.6, 5.467, 0.6, [TEAL, 40]);
    rect(s, 1.2, 1.2, 5.467, 5.7, ORANGE);
  },
};

/* ---------------------------------------------------------------- slides */

function slide01(s) {
  BACKDROPS[2](s);
  photo(s, 0.6, 0.6, 12.133, 6.3, 'tl tr br bl');
  softRect(s, 0.6, 0.6, 6.067, 6.3, [WHITE, 10], 'tl bl');
  softRect(s, 6.667, 0.6, 6.067, 6.3, [TEAL, 30], 'tr br');
  text(s, 1.147, 2.577, 3.392, 1.212, 'Paws.', st(PS, 72, SAND));
  text(s, 1.2, 3.735, 3.338, 0.366, 'Animal Care Presentation', st(PR, 16, DARK), { lineSpacingMultiple: 1.5 });
  text(s, 1.2, 2.2, 3.753, 0.303, 'Business Template', st(PL, 12, GREY), { lineSpacingMultiple: 1.5 });
  oval(s, 1.2, 4.862, 0.095, 0.095, SILVER);
  oval(s, 1.443, 4.862, 0.095, 0.095, SAND);
  oval(s, 1.686, 4.862, 0.095, 0.095, SILVER);
  oval(s, 1.929, 4.862, 0.095, 0.095, ORANGE);
  oval(s, 2.172, 4.862, 0.095, 0.095, SILVER);
  oval(s, 1.2, 4.862, 0.095, 0.095, SILVER);
  oval(s, 1.2, 5.105, 0.095, 0.095, TEAL);
  oval(s, 1.443, 5.105, 0.095, 0.095, MINT);
  oval(s, 1.686, 5.105, 0.095, 0.095, SILVER);
  oval(s, 1.929, 5.105, 0.095, 0.095, SAND);
  oval(s, 2.172, 5.105, 0.095, 0.095, SILVER);
  text(s, 1.2, 4.171, 3.338, 0.185, 'Nunc mioasjaa sapiennasa', NOTE);
  rrect(s, 5.664, 3.364, 2.006, 0.772, [SAND, 50], 0.099, { rotate: 180 });
  rrect(s, 5.862, 3.553, 1.608, 0.393, WHITE, 0.072, { rotate: 180, shadow: shadow(0.2, 4, 3) });
  text(s, 6.065, 3.598, 1.203, 0.303, [{ text: 'Pet', options: { color: TEAL } }, { text: ' ', options: { color: COAL } }, { text: 'Shop' }], st(PR, 12, ORANGE), { align: 'center', margin: [0, 0, 3.6, 3.6] });
  pageNo(s, 1);
}

function slide02(s) {
  BACKDROPS[3](s);
  photo(s, 6.667, 1.2, 5.467, 5.1);
  rect(s, 5.467, 0.6, 2.4, 3, SAND);
  rect(s, 5.467, 2.4, 1.2, 1.2, TEAL);
  text(s, 1.2, 3.265, 3.667, 2.473, 'Nunc mioasjaa sapien, venenatis quis erat sed, sodales pharetra justo. Donec adat iaculis diam. Pellentesque in risus namsa asut magna Phasellus metus arcu\n\nEtiam accumsan, ante eu luctus tin nulla ante iaculis tortor, eanamt auctor purus nunc eu metus. viverra at risus quis', BODY);
  text(s, 1.2, 2.304, 3.667, 0.74, 'The purpose \nof lives is be happy.', H2_D);
  dots(s, 1.2, 1.762);
  text(s, 6.48, 2.307, 1.859, 0.185, 'Nunc mioasjaa veenatis', st(PL, 11, WHITE), { rotate: 270 });
  text(s, 6.202, 2.299, 1.859, 0.202, 'Option One', LEAD_W, { rotate: 270 });
  poly(s, 5.855, 2.814, 0.422, 0.373, WHITE, [['M', 0.36, 0.054], [0.32, 0.019, 0.267, 0, 0.211, 0], [0.155, 0, 0.102, 0.019, 0.062, 0.054], [0.022, 0.089, 0, 0.136, 0, 0.186], [0, 0.227, 0.015, 0.267, 0.043, 0.299], [0.037, 0.312, 0.029, 0.322, 0.018, 0.329], [0.012, 0.333, 0.009, 0.34, 0.011, 0.346], [0.012, 0.353, 0.017, 0.357, 0.024, 0.358], [0.031, 0.359, 0.037, 0.36, 0.044, 0.36], [0.065, 0.36, 0.085, 0.355, 0.102, 0.346], [0.135, 0.364, 0.172, 0.373, 0.211, 0.373], [0.267, 0.373, 0.32, 0.354, 0.36, 0.319], [0.4, 0.284, 0.422, 0.237, 0.422, 0.186], [0.422, 0.136, 0.4, 0.089, 0.36, 0.054], [], ['M', 0.211, 0.352], [0.174, 0.352, 0.138, 0.343, 0.107, 0.325], [0.105, 0.324, 0.104, 0.324, 0.102, 0.324], [0.1, 0.324, 0.098, 0.324, 0.097, 0.325], [0.081, 0.334, 0.064, 0.339, 0.044, 0.339], [0.042, 0.339, 0.041, 0.339, 0.039, 0.339], [0.05, 0.329, 0.059, 0.316, 0.065, 0.301], [0.066, 0.297, 0.065, 0.293, 0.063, 0.29], [0.035, 0.26, 0.021, 0.225, 0.021, 0.186], [0.021, 0.095, 0.106, 0.021, 0.211, 0.021], [0.316, 0.021, 0.401, 0.095, 0.401, 0.186], [0.401, 0.278, 0.316, 0.352, 0.211, 0.352], []]);
  poly(s, 5.968, 2.921, 0.199, 0.179, WHITE, [['M', 0.138, 0], [0.124, 0, 0.11, 0.005, 0.1, 0.014], [0.089, 0.005, 0.075, 0, 0.061, 0], [0.045, 0, 0.029, 0.006, 0.018, 0.018], [0.006, 0.03, 0, 0.045, 0, 0.061], [0, 0.078, 0.006, 0.093, 0.018, 0.105], [0.087, 0.174], [0.09, 0.177, 0.095, 0.179, 0.1, 0.179], [0.104, 0.179, 0.109, 0.177, 0.112, 0.174], [0.181, 0.105], [0.193, 0.093, 0.199, 0.078, 0.199, 0.061], [0.199, 0.045, 0.193, 0.03, 0.181, 0.018], [0.17, 0.006, 0.154, 0, 0.138, 0], [], ['M', 0.166, 0.09], [0.1, 0.157], [0.033, 0.09], [0.025, 0.082, 0.021, 0.072, 0.021, 0.061], [0.021, 0.051, 0.025, 0.04, 0.033, 0.033], [0.04, 0.025, 0.05, 0.021, 0.061, 0.021], [0.072, 0.021, 0.082, 0.025, 0.089, 0.033], [0.092, 0.036], [0.094, 0.038, 0.097, 0.039, 0.1, 0.039], [0.102, 0.039, 0.105, 0.038, 0.107, 0.036], [0.11, 0.033], [0.117, 0.025, 0.127, 0.021, 0.138, 0.021], [0.149, 0.021, 0.159, 0.025, 0.166, 0.033], [0.174, 0.04, 0.178, 0.051, 0.178, 0.061], [0.178, 0.072, 0.174, 0.082, 0.166, 0.09], []]);
  pageNo(s, 2);
}

function slide03(s) {
  BACKDROPS[4](s);
  photo(s, 5.467, 1.2, 3.6, 5.1);
  photo(s, 1.2, 3.233, 3.667, 3.067);
  text(s, 1.2, 2.331, 3.067, 0.529, 'Nunc mioasjaa sapien, venenatis quis erat sed, sodales pharetra.', BODY);
  text(s, 1.2, 1.741, 3.067, 0.37, 'About Us', H2_D);
  dots(s, 1.2, 1.2);
  text(s, 10.267, 5.257, 1.865, 0.555, 'Nunc mioasjaa veenatis masnan naub', BODY);
  text(s, 9.665, 5.052, 0.582, 0.808, 'T', st(PS, 48, TEAL));
  text(s, 10.267, 5.045, 1.865, 0.202, 'Option One', st(PM, 12, SLATE));
  text(s, 10.267, 3.553, 1.865, 0.555, 'Nunc mioasjaa veenatis masnan naub', BODY);
  text(s, 9.665, 3.348, 0.582, 0.808, 'E', st(PS, 48, ORANGE));
  text(s, 10.267, 3.341, 1.865, 0.202, 'Option One', st(PM, 12, SLATE));
  text(s, 10.267, 1.849, 1.865, 0.555, 'Nunc mioasjaa veenatis masnan naub', BODY);
  text(s, 9.665, 1.645, 0.582, 0.808, 'P', st(PS, 48, SLATE));
  text(s, 10.267, 1.638, 1.865, 0.202, 'Option One', st(PM, 12, SLATE));
  pageNo(s, 3);
}

function slide04(s) {
  BACKDROPS[5](s);
  text(s, 2.081, 5.345, 1.859, 0.555, 'Nunc mioasjaa veenatis masnan naub', BODY);
  text(s, 1.2, 5.092, 0.54, 0.808, 'C', st(PS, 48, SAND), { align: 'center' });
  text(s, 2.081, 5.083, 1.859, 0.202, 'Option One', LEAD);
  text(s, 2.081, 3.603, 1.859, 0.555, 'Nunc mioasjaa veenatis masnan naub', BODY);
  text(s, 1.2, 3.351, 0.54, 0.808, 'B', st(PS, 48, SAND), { align: 'center' });
  text(s, 2.081, 3.342, 1.859, 0.202, 'Option One', LEAD);
  text(s, 2.081, 1.861, 1.859, 0.555, 'Nunc mioasjaa veenatis masnan naub', BODY);
  text(s, 1.2, 1.609, 0.54, 0.808, 'A', st(PS, 48, SLATE), { align: 'center' });
  text(s, 2.081, 1.6, 1.859, 0.202, 'Option One', LEAD);
  text(s, 6.178, 5.345, 1.859, 0.555, 'Nunc mioasjaa veenatis masnan naub', BODY_W);
  text(s, 5.296, 5.092, 0.54, 0.808, 'F', st(PS, 48, WHITE), { align: 'center' });
  text(s, 6.178, 5.083, 1.859, 0.202, 'Option One', LEAD_W);
  text(s, 6.178, 3.603, 1.859, 0.555, 'Nunc mioasjaa veenatis masnan naub', BODY_W);
  text(s, 5.296, 3.351, 0.54, 0.808, 'E', st(PS, 48, WHITE), { align: 'center' });
  text(s, 6.178, 3.342, 1.859, 0.202, 'Option One', LEAD_W);
  text(s, 6.178, 1.861, 1.859, 0.555, 'Nunc mioasjaa veenatis masnan naub', BODY_W);
  text(s, 5.296, 1.609, 0.54, 0.808, 'D', st(PS, 48, WHITE), { align: 'center' });
  text(s, 6.178, 1.6, 1.859, 0.202, 'Option One', LEAD_W);
  text(s, 10.274, 5.345, 1.859, 0.555, 'Nunc mioasjaa veenatis masnan naub', BODY);
  text(s, 9.393, 5.092, 0.54, 0.808, 'I', st(PS, 48, SLATE), { align: 'center' });
  text(s, 10.274, 5.083, 1.859, 0.202, 'Option One', LEAD);
  text(s, 10.274, 3.603, 1.859, 0.555, 'Nunc mioasjaa veenatis masnan naub', BODY);
  text(s, 9.393, 3.351, 0.54, 0.808, 'H', st(PS, 48, SLATE), { align: 'center' });
  text(s, 10.274, 3.342, 1.859, 0.202, 'Option One', LEAD);
  text(s, 10.274, 1.861, 1.859, 0.555, 'Nunc mioasjaa veenatis masnan naub', BODY);
  text(s, 9.393, 1.609, 0.54, 0.808, 'G', st(PS, 48, SAND), { align: 'center' });
  text(s, 10.274, 1.6, 1.859, 0.202, 'Option One', LEAD);
  hline(s, 1.2, 2.835, 10.933, SAND);
  hline(s, 1.2, 4.652, 10.933, SAND);
  pageNo(s, 4);
}

function slide05(s) {
  BACKDROPS[6](s);
  rect(s, 4.95, 6.9, 0.6, 0.6, [SAND, 90], { flipV: true });
  text(s, 6.15, 3.68, 4.783, 1.085, 'Nunc mioasjaa sapien, venenatis quis erat sed, sodales pharetra justo. Donec adat iaculis diam. Pellentesque in risus namsa asut magna viverra consequat ac in augue. Aliquam arcu est etiam accumsan, ante eu', BODY);
  text(s, 6.15, 2.343, 4.783, 1.111, 'Life’s tragedy is \nthat we get old too soon and wise too late.', H2);
  dots(s, 6.15, 1.801);
  rect(s, 6.754, 5.55, 3.165, 0.265, IVORY);
  rect(s, 6.149, 5.55, 2.62, 0.265, SAND);
  rect(s, 6.149, 5.55, 1.921, 0.265, TEAL);
  text(s, 10.196, 5.583, 0.737, 0.202, '2500+', st(PM, 12, SLATE), { align: 'right' });
  text(s, 6.149, 5.218, 1.974, 0.202, 'Treatment Process', SMALL);
  text(s, 7.319, 5.945, 0.751, 0.202, '+47', st(PR, 12, TEAL), { align: 'right' });
  text(s, 8.268, 5.945, 0.501, 0.202, '+21', st(PR, 12, SAND), { align: 'right' });
  photo(s, 1.2, 1.2, 3.75, 3.75);
  photo(s, 1.2, 5.55, 3.75, 1.349);
  pageNo(s, 5);
}

function slide06(s) {
  BACKDROPS[7](s);
  photo(s, 0.6, 0.6, 4.267, 1.2, 'tl');
  photo(s, 7.867, 1.2, 4.867, 5.7, 'br', 0.101);
  text(s, 1.2, 5.053, 1.849, 0.185, 'Mauris ullamcorper', NOTE);
  text(s, 1.2, 4.783, 1.849, 0.202, 'Briefcase Set', LEAD_D);
  text(s, 3.618, 5.053, 1.849, 0.185, 'Sed pharetra eu tellus', NOTE);
  text(s, 3.618, 4.783, 1.849, 0.202, 'Project Launch', LEAD_D);
  text(s, 5.847, 2.482, 2.867, 0.303, 'Completed Project Experience', st(PR, 12, INK), { rotate: 90, margin: [0, 0, 3.6, 3.6] });
  oval(s, 7.07, 4.303, 0.393, 0.393, INK, { rotate: 90 });
  text(s, 7.071, 4.349, 0.393, 0.303, [{ text: '5' }, { text: '+', options: { fontSize: 8 } }], st(PR, 12, SNOW), { align: 'center', rotate: 90, margin: [0, 0, 3.6, 3.6] });
  text(s, 1.2, 3.381, 4.267, 0.807, 'Nunc mioasjaa sapien, venenatis quis erat sed, sodales pharetra. mioasjaa sapien, venenatis quis erat sed, sodales pharetra justo', BODY);
  text(s, 1.2, 2.791, 4.267, 0.37, 'We care about you', H2_D);
  dots(s, 1.2, 2.25);
  text(s, 1.195, 5.463, 1.344, 0.673, [{ text: '40' }, { text: '%', options: { superscript: true } }], st(PS, 40, GREY));
  text(s, 3.618, 5.463, 1.344, 0.673, [{ text: '70' }, { text: '%', options: { superscript: true } }], st(PS, 40, TEAL));
  hline(s, 1.2, 4.642, 1.249, SILVER);
  hline(s, 3.618, 4.642, 1.249, TEAL);
  pageNo(s, 6);
}

function slide07(s) {
  BACKDROPS[8](s);
  photo(s, 1.2, 1.2, 5.467, 4.5);
  photo(s, 7.267, 2.85, 4.867, 1.8);
  text(s, 2.55, 5.846, 2.917, 0.303, 'Completed Project Experience', st(PR, 12, INK), { margin: [0, 0, 3.6, 3.6] });
  text(s, 1.743, 5.846, 0.74, 0.303, '4000+', LEAD_D, { margin: [0, 0, 3.6, 3.6] });
  text(s, 5.697, 5.846, 0.74, 0.303, 'A+', st(PR, 12, WHITE), { align: 'center', margin: [0, 0, 3.6, 3.6] });
  text(s, 7.267, 5.19, 4.878, 0.807, 'Nunc mioasjaa sapien, venenatis quis erat sed, sodales pharetra justo. Donec adat iaculis diam. Pellentesque in risus namsa asut magna viverra consequa', BODY);
  text(s, 12.144, 3.599, 0.589, 0.303, 'A+', st(PR, 12, WHITE), { align: 'center', margin: [0, 0, 3.6, 3.6] });
  text(s, 7.267, 1.742, 4.783, 0.74, 'The unexamined \nlife is not worth living.', H2);
  dots(s, 7.267, 1.2);
  pageNo(s, 7);
}

function slide08(s) {
  BACKDROPS[9](s);
  text(s, 1.789, 1.676, 3.078, 0.807, 'Nunc mioasjaa sapien, venenatis naub \nerat sed, sodales pharetra justo conec adat iaculis diam. ', BODY);
  text(s, 1.2, 1.225, 0.589, 0.37, '01.', H2_D);
  text(s, 1.789, 1.362, 2.044, 0.202, 'Briefcase Set', LEAD_D);
  text(s, 1.789, 5.538, 3.078, 0.807, 'Nunc mioasjaa sapien, venenatis naub \nerat sed, sodales pharetra justo conec adat iaculis diam. ', BODY);
  text(s, 1.2, 5.087, 0.589, 0.37, '03.', H2_D);
  text(s, 1.789, 5.225, 2.044, 0.202, 'Briefcase Set', LEAD_D);
  text(s, 1.789, 3.607, 3.078, 0.807, 'Nunc mioasjaa sapien, venenatis naub \nerat sed, sodales pharetra justo conec adat iaculis diam. ', BODY);
  text(s, 1.2, 3.156, 0.589, 0.37, '02.', st(PM, 22, ORANGE));
  text(s, 1.789, 3.293, 2.044, 0.202, 'Briefcase Set', st(PM, 12, ORANGE));
  text(s, 6.22, 0.998, 0.893, 0.404, '5000%', LEAD_W, { align: 'center', margin: [0, 0, 3.6, 3.6] });
  text(s, 5.467, 5.549, 2.989, 0.74, 'Turn your wounds \ninto wisdom', H2_D);
  text(s, 8.901, 5.792, 1.393, 0.529, 'sapien venenatis naub aama', BODY, { align: 'center' });
  vline(s, 10.517, 5.549, 0.74, SILVER);
  text(s, 9.037, 5.458, 1.122, 0.265, '70+', st(OB, 12, DARK), { align: 'center', lineSpacingMultiple: 1.5 });
  text(s, 10.74, 5.792, 1.393, 0.529, 'sapien venenatis naub aama', BODY, { align: 'center' });
  text(s, 10.876, 5.458, 1.122, 0.265, '45+', st(OB, 12, DARK), { align: 'center', lineSpacingMultiple: 1.5 });
  text(s, 8.62, 1.166, 1.487, 0.236, 'Great Choices', st(PM, 14, WHITE));
  text(s, 8.62, 0.967, 1.487, 0.177, 'Aenean eget', st(PL, 10.5, WHITE));
  text(s, 10.739, 1.166, 0.806, 0.236, 'SS+', st(PM, 14, WHITE), { align: 'right' });
  photo(s, 5.467, 1.8, 6.667, 3.3);
  pageNo(s, 8);
}

function slide09(s) {
  BACKDROPS[10](s);
  photo(s, 2.4, 1.2, 4.267, 5.1);
  text(s, 0.376, 4.612, 2.847, 0.529, 'Nunc mioasjaa sapien, venenatis quis erat sed, sodales', BODY, { rotate: 270 });
  pill(s, 0.996, 1.808, 1.608, 0.393, TEAL, { rotate: 90 });
  text(s, 1.412, 1.853, 0.775, 0.303, '25.000+', st(OB, 12, SNOW), { align: 'center', rotate: 270, margin: [0, 0, 3.6, 3.6] });
  rect(s, 2.4, 3.75, 4.267, 2.55, [ORANGE, 30]);
  text(s, 7.267, 3.614, 2.25, 0.529, 'Nunc mioasjaa sapien erat sed, sodales phare.', BODY);
  text(s, 7.267, 3.318, 2.25, 0.236, 'Concept', st(PM, 14, ORANGE));
  text(s, 9.8, 3.614, 2.25, 0.529, 'Erat sed, sodales pharetra justo conec.', BODY);
  text(s, 9.8, 3.318, 2.25, 0.236, 'Treats', st(PM, 14, INK));
  text(s, 7.267, 2.107, 4.783, 0.74, 'The unexamined life is not worth living.', H2);
  dots(s, 7.267, 1.565);
  text(s, 7.267, 4.925, 2.25, 1.01, '47+', st(PS, 60, SLATE));
  text(s, 9.8, 5.313, 2.25, 0.529, 'Erat sed, sodales pharetra justo conec.', BODY);
  text(s, 9.8, 5.018, 2.25, 0.236, 'Points', st(PM, 14, INK));
  pageNo(s, 9);
}

function slide10(s) {
  BACKDROPS[11](s);
  photo(s, 0.6, 1.2, 6.067, 2.55);
  photo(s, 7.267, 4.99, 4.781, 1.907);
  text(s, 1.2, 3.896, 2.917, 0.303, 'Completed Project Experience', st(PR, 12, WHITE), { margin: [0, 0, 3.6, 3.6] });
  text(s, 7.267, 1.742, 4.783, 0.74, 'In order to write \nabout life first you must live it', H2);
  dots(s, 7.267, 1.2);
  oval(s, 7.267, 3.08, 0.746, 0.746, [SAND, 70]);
  oval(s, 7.397, 3.21, 0.486, 0.486, SAND);
  text(s, 8.322, 3.613, 3.728, 0.807, 'Nunc mioasjaa sapien erata sed, \nsodales pharetra enean eget nequena sagittis, mattis libero quis.', BODY);
  text(s, 7.334, 3.318, 0.589, 0.269, '01', st(PS, 16, WHITE), { align: 'center' });
  text(s, 8.322, 3.329, 2.044, 0.236, 'Description Text', st(PM, 14, DARK));
  oval(s, 1.197, 4.956, 0.746, 0.746, [TEAL, 70]);
  oval(s, 1.327, 5.086, 0.486, 0.486, TEAL);
  text(s, 2.253, 5.49, 3.728, 0.807, 'Nunc mioasjaa sapien erata sed, \nsodales pharetra enean eget nequena sagittis, mattis libero quis.', BODY);
  text(s, 1.265, 5.195, 0.589, 0.269, '01', st(PS, 16, WHITE), { align: 'center' });
  text(s, 2.253, 5.206, 2.044, 0.236, 'Description Text', st(PM, 14, DARK));
  rect(s, 9.657, 4.987, 2.39, 0.573, SAND);
  text(s, 9.99, 5.105, 1.724, 0.303, '24 Days', st(PR, 12, WHITE), { align: 'center', margin: [0, 0, 3.6, 3.6] });
  pageNo(s, 10);
}

function slide11(s) {
  BACKDROPS[12](s);
  photo(s, 9.133, 0.6, 3.6, 4.3, 'tr');
  text(s, 1.2, 2.678, 7.333, 0.529, 'Mauris ullamcorper, ex ac semper finibus, diam nulla mollis ex, non pharetra ex nisl sed mi. In fermentum, leo non efficitur tempor, mi nulla fringilla libero', BODY);
  text(s, 1.789, 3.959, 2.478, 0.529, 'Nunc mioasjaa sapien, venenats naub erat sedanj', BODY);
  text(s, 1.2, 3.509, 0.589, 0.37, '01.', st(PM, 22, SAND));
  text(s, 1.789, 3.637, 2.044, 0.202, 'Briefcase Set', st(PM, 12, SAND));
  text(s, 6.056, 3.959, 2.478, 0.529, 'Nunc mioasjaa sapien, venenats naub erat sedanj', BODY);
  text(s, 5.467, 3.509, 0.589, 0.37, '02.', H2_D);
  text(s, 6.056, 3.637, 2.044, 0.202, 'Briefcase Set', LEAD_D);
  text(s, 1.789, 5.769, 2.478, 0.529, 'Nunc mioasjaa sapien, venenats naub erat sedanj', BODY);
  text(s, 1.2, 5.318, 0.589, 0.37, '03.', H2_D);
  text(s, 1.789, 5.446, 2.044, 0.202, 'Briefcase Set', LEAD_D);
  text(s, 6.056, 5.769, 2.478, 0.529, 'Nunc mioasjaa sapien, venenats naub erat sedanj', BODY);
  text(s, 5.467, 5.318, 0.589, 0.37, '04.', st(PM, 22, TEAL));
  text(s, 6.056, 5.446, 2.044, 0.202, 'Briefcase Set', st(PM, 12, TEAL));
  text(s, 1.2, 1.742, 7.333, 0.74, 'Not how long but how well \nyou have lived is the main thing.', H2);
  dots(s, 1.2, 1.2);
  pageNo(s, 11);
}

function slide12(s) {
  BACKDROPS[13](s);
  photo(s, 6.661, 0.6, 6.067, 6.3, 'tr br');
  softRect(s, 6.667, 0.6, 6.067, 6.3, [WHITE, 10], 'tr br');
  rect(s, 6.667, 5.7, 1.2, 1.2, TEAL);
  text(s, 8.456, 2.168, 3.678, 0.185, 'Nulla non tortor at augue tincidunt elementum. ', NOTE);
  text(s, 7.867, 1.958, 0.487, 0.303, '01.', st(OB, 12, INK));
  text(s, 8.456, 1.88, 2.044, 0.202, 'Option Description', LEAD);
  text(s, 8.456, 4.856, 3.678, 0.185, 'Mauris ullamcorper, ex ac semper finibus', NOTE);
  text(s, 7.867, 4.645, 0.487, 0.303, '03.', st(OB, 12, INK));
  text(s, 8.456, 4.568, 2.044, 0.202, 'Description Text', LEAD);
  text(s, 8.456, 3.38, 3.678, 0.37, 'Donec sollicitudin sollicitudin risus eu diam nulla naums sit amet ante tellus.', NOTE);
  text(s, 7.867, 3.169, 0.487, 0.303, '02.', st(OB, 12, TEAL));
  text(s, 8.456, 3.092, 2.044, 0.202, 'Option Description Text', st(PM, 12, TEAL));
  text(s, 1.2, 2.884, 4.825, 2.195, 'Nulla non tortor at augue tincidunt elementum. lorema\nDonec sollicitudin sollicitudin risus eu ultrices. Aliquam eget erat vitae nisl dignissim semper et vitae diam. Aliquam rhoncus fringilla tempor. \n\nneque est efficitur velit, in tincidunt urna risus sit \namet libero. Cras vel fermentum metus. Integer luctus convallis massa, ut rutrum lectus', BODY);
  text(s, 1.242, 1.91, 4.783, 0.74, 'In order to write about life \nfirst you must live it', H2);
  dots(s, 1.242, 1.368);
  text(s, 1.2, 6.368, 2.439, 0.185, 'mattis libero rhoncusana', st(PL, 11, WHITE));
  text(s, 1.2, 6.047, 2.511, 0.236, 'Option A', st(PM, 14, WHITE));
  pageNo(s, 12);
}

function slide13(s) {
  BACKDROPS[14](s);
  photo(s, 5.533, 1.8, 3.3, 3.9);
  text(s, 1.2, 2.776, 3.733, 2.195, 'Nulla non tortor at augue tincidunt elementum. Donec sollicitudin sollicitudin risus eu ultrices. Aliquam eget erat vitae nisl dignissim semper et vitae diam. Aliquam rhoncus\n\nneque est efficitur velit, in tincidunt urna risus sit amet libero. Cras vel fermentum metus. Integer luctus convallis massa, ut', BODY);
  text(s, 9.686, 2.903, 2.447, 0.807, 'Erat sed, sodales phar anunjusto conec adat iaculis diam. Mauris', BODY);
  text(s, 9.433, 2.533, 2.1, 0.202, '500 -  Target Pets', LEAD, { bullet: { indent: 18.25 } });
  text(s, 9.62, 4.471, 2.513, 0.807, 'Nulla non tortor at augue tincidunt elementum. Donec sollicitudin', BODY);
  text(s, 9.433, 4.186, 2.1, 0.202, 'Release', st(PM, 12, TEAL), { bullet: { indent: 13.5 } });
  text(s, 6.133, 6, 2.1, 0.497, 'Nulla non tortor at augunmehae tincidunt elementum', st(OL, 10.5, WHITE), { lineSpacingMultiple: 1.5 });
  text(s, 1.242, 1.742, 3.692, 0.74, 'The healthiest \nresponse to life is joy.', H2);
  dots(s, 1.242, 1.2);
  oval(s, 1.2, 5.393, 0.607, 0.607, [TEAL, 50]);
  text(s, 2.007, 5.549, 2.699, 0.303, 'Good Treatment', st(PR, 12, INK), { margin: [0, 0, 3.6, 3.6] });
  oval(s, 1.307, 5.5, 0.393, 0.393, TEAL);
  text(s, 1.307, 5.544, 0.393, 0.303, 'S+', st(PR, 12, WHITE), { align: 'center', margin: [0, 0, 3.6, 3.6] });
  pageNo(s, 13);
}

function slide14(s) {
  BACKDROPS[15](s);
  text(s, 1.2, 2.653, 3.6, 0.529, 'Nulla non tortor at augue tincidunt elementum. Donec sollicitudin sollicitudin.', BODY);
  text(s, 9.665, 4.561, 1.307, 0.269, '70+ Sub', st(PS, 16, SLATE));
  text(s, 9.665, 4.861, 1.307, 0.185, 'Anhte eu luctusa', NOTE);
  oval(s, 8.989, 4.564, 0.5, 0.5, WHITE, { shadow: shadow(0.1, 5, 3) });
  text(s, 9.656, 5.534, 2.477, 0.555, 'Nunc mioasa veenatis loramsna  masnan naubanan', BODY);
  text(s, 9.656, 5.273, 2.16, 0.202, 'Option Two', st(PM, 12, SLATE));
  text(s, 6.209, 4.561, 1.307, 0.269, '200+ Sub', st(PS, 16, SLATE));
  text(s, 6.209, 4.861, 1.307, 0.185, 'Anhte eu luctusa', NOTE);
  oval(s, 5.533, 4.564, 0.5, 0.5, WHITE, { shadow: shadow(0.1, 5, 3) });
  text(s, 6.2, 5.534, 2.477, 0.555, 'Nunc mioasa veenatis loramsna  masnan naubanan', BODY);
  text(s, 6.2, 5.273, 2.16, 0.202, 'Option One', st(PM, 12, SLATE));
  poly(s, 5.682, 4.715, 0.203, 0.198, TEAL, [['M', 0.192, 0.075], [0.187, 0.071, 0.18, 0.069, 0.174, 0.069], [0.167, 0.069], [0.144, 0.069], [0.13, 0.069], [0.13, 0.037], [0.13, 0.024, 0.126, 0.014, 0.119, 0.009], [0.107, 0, 0.09, 0.005, 0.089, 0.005], [0.087, 0.006, 0.086, 0.008, 0.086, 0.01], [0.086, 0.045], [0.086, 0.057, 0.08, 0.066, 0.07, 0.074], [0.062, 0.08, 0.053, 0.082, 0.052, 0.083], [0.051, 0.083], [0.049, 0.08, 0.046, 0.078, 0.042, 0.078], [0.01, 0.078], [0.005, 0.078, 0, 0.082, 0, 0.088], [0, 0.183], [0, 0.188, 0.005, 0.193, 0.01, 0.193], [0.043, 0.193], [0.046, 0.193, 0.049, 0.191, 0.05, 0.189], [0.056, 0.195, 0.063, 0.198, 0.071, 0.198], [0.098, 0.198], [0.101, 0.198], [0.157, 0.198], [0.176, 0.198, 0.188, 0.188, 0.19, 0.171], [0.201, 0.102], [0.203, 0.092, 0.199, 0.082, 0.192, 0.075], [], ['M', 0.043, 0.183], [0.043, 0.183, 0.043, 0.183, 0.042, 0.183], [0.01, 0.183], [0.01, 0.183, 0.01, 0.183, 0.01, 0.183], [0.01, 0.088], [0.01, 0.088, 0.01, 0.088, 0.01, 0.088], [0.043, 0.088], [0.043, 0.088, 0.043, 0.088, 0.043, 0.088], [0.043, 0.183], [0.043, 0.183], [], ['M', 0.192, 0.1], [0.18, 0.169], [0.18, 0.169, 0.18, 0.169, 0.18, 0.169], [0.179, 0.182, 0.171, 0.188, 0.157, 0.188], [0.101, 0.188], [0.098, 0.188], [0.071, 0.188], [0.062, 0.188, 0.054, 0.182, 0.053, 0.173], [0.053, 0.173, 0.053, 0.172, 0.053, 0.172], [0.053, 0.093], [0.055, 0.092], [0.055, 0.092, 0.055, 0.092, 0.055, 0.092], [0.056, 0.092, 0.065, 0.089, 0.075, 0.082], [0.089, 0.073, 0.096, 0.06, 0.096, 0.045], [0.096, 0.014], [0.1, 0.013, 0.108, 0.013, 0.113, 0.017], [0.118, 0.02, 0.12, 0.027, 0.12, 0.037], [0.12, 0.073], [0.12, 0.076, 0.123, 0.078, 0.125, 0.078], [0.144, 0.078], [0.167, 0.078], [0.174, 0.078], [0.178, 0.078, 0.182, 0.08, 0.185, 0.083], [0.19, 0.087, 0.193, 0.094, 0.192, 0.1], []]);
  poly(s, 9.146, 4.729, 0.186, 0.202, SAND, [['M', 0.186, 0.048], [0.186, 0.047, 0.186, 0.046, 0.185, 0.045], [0.185, 0.044, 0.183, 0.043, 0.182, 0.043], [0.095, 0.001], [0.094, 0, 0.092, 0, 0.091, 0.001], [0.003, 0.043], [0.001, 0.044, 0, 0.046, 0, 0.048], [0, 0.048], [0, 0.048, 0, 0.048, 0, 0.048], [0, 0.154], [0, 0.156, 0.001, 0.158, 0.003, 0.159], [0.091, 0.201], [0.091, 0.201, 0.091, 0.201, 0.091, 0.201], [0.091, 0.202, 0.091, 0.202, 0.091, 0.202], [0.091, 0.202, 0.091, 0.202, 0.091, 0.202], [0.091, 0.202, 0.092, 0.202, 0.092, 0.202], [0.092, 0.202, 0.092, 0.202, 0.092, 0.202], [0.092, 0.202, 0.092, 0.202, 0.092, 0.202], [0.092, 0.202, 0.092, 0.202, 0.092, 0.202], [0.092, 0.202, 0.093, 0.202, 0.093, 0.202], [0.093, 0.202, 0.093, 0.202, 0.093, 0.202], [0.093, 0.202, 0.094, 0.202, 0.094, 0.202], [0.094, 0.202, 0.094, 0.202, 0.094, 0.202], [0.094, 0.202, 0.094, 0.202, 0.094, 0.202], [0.094, 0.202, 0.094, 0.202, 0.094, 0.202], [0.095, 0.202, 0.095, 0.202, 0.095, 0.202], [0.095, 0.202, 0.095, 0.202, 0.095, 0.201], [0.095, 0.201, 0.095, 0.201, 0.095, 0.201], [0.183, 0.159], [0.185, 0.158, 0.186, 0.156, 0.186, 0.154], [0.186, 0.048], [0.186, 0.048, 0.186, 0.048, 0.186, 0.048], [], ['M', 0.093, 0.011], [0.169, 0.048], [0.141, 0.061], [0.065, 0.025], [0.093, 0.011], [], ['M', 0.093, 0.085], [0.017, 0.048], [0.053, 0.03], [0.129, 0.067], [0.093, 0.085], [], ['M', 0.01, 0.056], [0.088, 0.094], [0.088, 0.189], [0.01, 0.151], [0.01, 0.056], [], ['M', 0.098, 0.189], [0.098, 0.094], [0.134, 0.076], [0.134, 0.101], [0.134, 0.104, 0.137, 0.106, 0.139, 0.106], [0.142, 0.106, 0.145, 0.104, 0.145, 0.101], [0.145, 0.071], [0.176, 0.056], [0.176, 0.151], [0.098, 0.189], []]);
  text(s, 1.2, 1.742, 3.558, 0.74, 'The greatest \npleasure of life is love.', H2);
  dots(s, 1.2, 1.2);
  photo(s, 1.2, 3.75, 3.6, 2.55);
  photo(s, 5.533, 1.2, 6.6, 2.55);
  pageNo(s, 14);
}

function slide15(s) {
  BACKDROPS[16](s);
  text(s, 1.2, 2.822, 3.067, 1.918, 'Nunc mioasjaa sapien, venenatis quis erat sed, sodales pharetra justo. Donec adat iaculis diam. Pellentesque in risus namsa asut \n\nEtiam accumsan, ante eu luctus nulla ante iaculis tortor, ', BODY);
  text(s, 1.889, 5.339, 2.378, 0.303, 'Completed Experience', LEAD, { margin: [0, 0, 3.6, 3.6] });
  oval(s, 1.2, 5.307, 0.393, 0.393, INK);
  text(s, 1.2, 5.351, 0.393, 0.303, [{ text: '17' }, { text: '+', options: { fontSize: 8 } }], st(PR, 12, SNOW), { align: 'center', margin: [0, 0, 3.6, 3.6] });
  text(s, 9.067, 3.469, 3.067, 0.529, 'Nunc mioasjaa sapien \nsodales pharetra justo conec adat', BODY);
  text(s, 9.067, 4.175, 3.067, 0.202, 'Concept', LEAD);
  text(s, 9.067, 4.817, 3.067, 0.529, 'Maecenas ligula orci, feugiat suscipit purus sed, suscipit lacinia', BODY);
  text(s, 9.067, 5.523, 3.067, 0.202, 'Development', LEAD);
  text(s, 1.242, 1.742, 3.025, 0.74, 'Every moment is beginning', H2);
  dots(s, 1.242, 1.2);
  photo(s, 4.867, 1.2, 3.6, 5.1);
  photo(s, 9.067, 1.2, 3.067, 1.8);
  pageNo(s, 15);
}

function slide16(s) {
  BACKDROPS[17](s);
  photo(s, 4.133, 2.77, 2.733, 1.95);
  photo(s, 7.067, 2.77, 2.733, 3.53);
  photo(s, 1.2, 2.77, 2.733, 1.95);
  photo(s, 9.994, 2.77, 2.733, 1.95);
  rect(s, 4.133, 2.77, 2.733, 1.95, [TEAL, 30]);
  rect(s, 7.067, 2.78, 2.733, 1.94, [SAND, 30]);
  text(s, 1.453, 5.388, 1.847, 0.529, 'Erat sed, sodales nuam anunjusto cone.', BODY);
  text(s, 1.2, 5.129, 2.1, 0.202, 'Build', LEAD, { bullet: { indent: 18.25 } });
  text(s, 4.386, 5.388, 1.847, 0.529, 'Erat sed, sodales nuam anunjusto cone.', BODY);
  text(s, 4.133, 5.129, 2.1, 0.202, 'Build', st(PM, 12, TEAL), { bullet: { indent: 18.25 } });
  text(s, 10.253, 5.388, 1.847, 0.529, 'Erat sed, sodales nuam anunjusto cone.', BODY);
  text(s, 10, 5.129, 2.1, 0.202, 'Build', LEAD, { bullet: { indent: 18.25 } });
  text(s, 1.242, 2.176, 4.667, 0.177, 'Nulla non torto augue tincidunt elementum. ', st(OL, 10.5, GREY));
  text(s, 1.242, 1.742, 5.625, 0.37, 'Don’t cry because it’s not over', H2);
  dots(s, 1.242, 1.2);
  text(s, 7.383, 2.001, 2.1, 0.529, 'Erat sed, sodales nuam anunjusto cone.', BODY, { align: 'center' });
  text(s, 7.383, 1.742, 2.1, 0.202, 'Build', st(PM, 12, SAND), { align: 'center' });
  pageNo(s, 16);
}

function slide17(s) {
  BACKDROPS[18](s);
  photo(s, 4.133, 1.2, 2.533, 2.35);
  photo(s, 1.2, 1.2, 2.533, 2.35);
  photo(s, 4.133, 3.95, 2.533, 2.35);
  rect(s, 4.133, 1.2, 2.533, 2.35, [TEAL, 30]);
  text(s, 5.792, 2.224, 2.35, 0.303, 'Project Experience', st(OR, 12, INK), { align: 'center', rotate: 90, margin: [0, 0, 3.6, 3.6] });
  text(s, 5.792, 4.974, 2.35, 0.303, [{ text: '400+ ', options: { fontFace: OB } }, { text: 'Experience' }], st(OR, 12, WHITE), { align: 'center', rotate: 90, margin: [0, 0, 3.6, 3.6] });
  text(s, 8.467, 3.918, 3.667, 1.918, 'Nulla non tortor at augue tincidunt elementum. Donec sollicitudin sollicitudin risus eu ultrices. Aliquam eget erat vitae nisl dignissim semper et vitae diam. Aliquam namus\n\nneque est efficitur velit, in tincidunt urna risus sit amet libero. Cras vel ferment muna', BODY);
  text(s, 8.467, 2.207, 3.692, 1.111, 'When you cease \nto dream you cease to live.', H2);
  dots(s, 8.467, 1.665);
  text(s, 1.2, 4.163, 2.533, 1.01, '14+', st(PS, 60, TEAL));
  text(s, 1.2, 5.558, 2.533, 0.529, 'Erat sed, sodales pharetra justo conec lorana', BODY);
  text(s, 1.2, 5.263, 2.533, 0.236, 'Points', st(PM, 14, INK));
  pageNo(s, 17);
}

function slide18(s) {
  BACKDROPS[19](s);
  photo(s, 7.267, 1.2, 4.867, 2.35);
  photo(s, 4.733, 3.95, 4.467, 2.35);
  text(s, 5.792, 2.224, 2.35, 0.303, '400+ Experience', LEAD_W, { align: 'center', rotate: 90, margin: [0, 0, 3.6, 3.6] });
  text(s, 1.2, 2.375, 4.267, 0.807, 'Aliquam consequat id velit vel hendrerit. \nSed sed metus eleifend, imperdiet magna non, iaculis neque. Sed facilisis sapien acan.', BODY);
  text(s, 1.387, 4.32, 2.146, 0.807, 'Mi nulla fringilla \nlibero nec pulvinar arcu mi et nunc', BODY);
  text(s, 1.2, 4.018, 1.733, 0.202, 'Option', LEAD_D, { bullet: { indent: 13.5 } });
  text(s, 1.387, 5.803, 2.146, 0.529, 'Aliquam conse velit velahu hendrerina', BODY);
  text(s, 1.2, 5.501, 1.733, 0.202, 'Option Two', LEAD_D, { bullet: { indent: 13.5 } });
  text(s, 9.8, 4.797, 1.733, 0.807, 'Nulla fringi libero \nnecan pulvinar arcu minet nuna.', BODY, { align: 'center' });
  text(s, 9.8, 4.495, 1.733, 0.202, 'Aa', LEAD_D, { align: 'center' });
  text(s, 1.2, 1.742, 4.361, 0.37, 'Every moment is beginning', H2);
  dots(s, 1.2, 1.2);
  text(s, 3.258, 4.974, 2.35, 0.303, '200+ Experience', LEAD_W, { align: 'center', rotate: 90, margin: [0, 0, 3.6, 3.6] });
  pageNo(s, 18);
}

function slide19(s) {
  BACKDROPS[20](s);
  photo(s, 9.067, 1.2, 3.067, 1.559);
  photo(s, 9.067, 4.043, 3.067, 1.559);
  photo(s, 0.597, 1.2, 7.87, 4.401);
  rect(s, 9.067, 4.043, 3.067, 1.559, [ORANGE, 30]);
  text(s, 9.067, 3.188, 3, 0.185, 'Curabitur sit amet ante tellus.', NOTE, { align: 'center' });
  text(s, 9.067, 2.922, 3, 0.202, 'Option One', LEAD_D, { align: 'center' });
  text(s, 9.067, 6.031, 3, 0.185, 'Nunc mioasjaa sapi venenatis ', NOTE, { align: 'center' });
  text(s, 9.067, 5.765, 3, 0.202, 'Option Two', LEAD_D, { align: 'center' });
  rect(s, 0.597, 1.199, 4.27, 4.402, [WHITE, 10]);
  text(s, 1.2, 3.885, 3.063, 1.085, 'Nunc mioasjaa sapien, venenatis quis erat sed, sodales pharetra justo. Donec adat iaculis diam eget neque sagittis, mattis libero quis', BODY);
  pill(s, 5.733, 5.707, 1.867, 0.393, INK, { rotate: 180 });
  text(s, 5.956, 5.797, 1.457, 0.202, [{ text: 'Points    ' }, { text: '|', options: { fontFace: PL } }, { text: '   ' }, { text: '25.5+', options: { color: SNOW } }], st(PM, 12, TEAL));
  text(s, 1.2, 5.809, 2.467, 0.202, 'Aliquam consequat id velit', st(PR, 12, WHITE));
  text(s, 1.2, 2.308, 3.063, 1.111, 'You never learn much from hearing yourself speak', H2);
  dots(s, 1.2, 1.766);
  pageNo(s, 19);
}

function slide20(s) {
  BACKDROPS[21](s);
  photo(s, 5.467, 1.2, 2.4, 1.433);
  photo(s, 5.467, 3.037, 2.4, 1.433);
  photo(s, 5.467, 4.867, 2.4, 1.433);
  text(s, 9.656, 1.887, 2.478, 0.807, 'Cras in orci quis velit dapibus mollis. Pellentesque dictum a nisl vel fermentum. ', BODY);
  text(s, 9.067, 1.437, 0.589, 0.37, '01.', st(PM, 22, TEAL));
  text(s, 9.656, 1.563, 2.264, 0.202, 'Option One', LEAD_D);
  text(s, 9.656, 3.727, 2.478, 0.807, 'Aenean eget neque sagittis, mattis libero quis, rhoncusansa augue sit amet ante', BODY);
  text(s, 9.067, 3.276, 0.589, 0.37, '02.', st(PM, 22, INK));
  text(s, 9.656, 3.403, 2.264, 0.202, 'Option One', LEAD);
  text(s, 9.656, 5.567, 2.478, 0.807, 'Nunc mioasjaa sapi ventis masnan naub erat sed, sodale samet ante tellus.', BODY);
  text(s, 9.067, 5.116, 0.589, 0.37, '03.', st(PM, 22, ORANGE));
  text(s, 9.656, 5.242, 2.264, 0.202, 'Option One', LEAD_D);
  text(s, 7.464, 5.415, 1.206, 0.303, '1500+', st(PR, 12, WHITE), { align: 'center', rotate: 90, margin: [0, 0, 3.6, 3.6] });
  text(s, 7.464, 1.742, 1.206, 0.303, '1500+', st(PR, 12, WHITE), { align: 'center', rotate: 90, margin: [0, 0, 3.6, 3.6] });
  text(s, 1.2, 2.393, 3.272, 1.111, 'When you cease \nto dream you cease to live.', H2);
  text(s, 1.2, 3.714, 3.267, 1.085, 'Nulla non tortor at augue tincidunt elementum. Donec sollicitudin sollicitudin risus eu ultrices. Aliquam eget erat vitae nisl dignissim', BODY);
  rrect(s, 1.2, 5.314, 1.383, 0.393, INK, 0, { rotate: 180 });
  text(s, 1.504, 5.359, 0.775, 0.303, '25.000+', st(OB, 12, SNOW), { align: 'center', margin: [0, 0, 3.6, 3.6] });
  hline(s, 1.2, 5.316, 3.067, GREY);
  text(s, 2.822, 5.359, 1.206, 0.303, 'Rhoncus augue', st(OL, 12, GREY), { align: 'center', margin: [0, 0, 3.6, 3.6] });
  dots(s, 1.2, 1.793);
  text(s, 4.664, 3.599, 1.206, 0.303, '1500+', st(PR, 12, GREY), { align: 'center', rotate: 90, margin: [0, 0, 3.6, 3.6] });
  pageNo(s, 20);
}

function slide21(s) {
  BACKDROPS[22](s);
  photo(s, 5.43, 1.2, 3.059, 2.35);
  photo(s, 1.2, 3.95, 3.628, 2.95);
  rect(s, 1.2, 6.3, 3.628, 0.6, [SAND, 30]);
  text(s, 9.067, 3.339, 3.067, 1.822, 'Nulla non tortor at augue tincidunt elementum. Donec sollicitudin sollicitudin risus eu ultrices. Aliquam eget erat vitae nisl dignissim semper et vitae diam. Aliquam namus \n\nneque est efficitur velit, in tincidunt urna risus sit amet libero vel ferment muna', st(OL, 10.5, GREY), { lineSpacingMultiple: 1.5 });
  rrect(s, 9.067, 5.735, 1.383, 0.265, INK, 0, { rotate: 180 });
  text(s, 9.371, 5.753, 0.775, 0.202, '25.000+', st(OB, 12, SNOW), { align: 'center' });
  text(s, 10.642, 5.753, 1.3, 0.202, [{ text: 'Rhoncus' }, { text: ' augue', options: { fontFace: OL } }], st(PM, 12, GREY), { align: 'center' });
  text(s, 9.067, 2.042, 3.067, 1.111, 'Some infinities are bigger than other infinities.', H2);
  dots(s, 9.067, 1.5);
  rect(s, 5.43, 5.737, 2.117, 0.265, [SILVER, 50]);
  text(s, 5.432, 5.395, 1.763, 0.202, 'Release Process', SMALL);
  text(s, 7.547, 5.769, 0.557, 0.202, '35%', st(PM, 12, SUN), { align: 'right', bullet: { indent: 13.5 } });
  rect(s, 5.43, 5.737, 0.886, 0.265, SAND);
  rect(s, 5.428, 4.662, 2.485, 0.265, [SILVER, 50]);
  text(s, 5.429, 4.32, 2.118, 0.202, 'Development Process', SMALL);
  text(s, 7.91, 4.694, 0.557, 0.202, '85%', st(PM, 12, TEAL), { align: 'right', bullet: { indent: 13.5 } });
  rect(s, 5.428, 4.662, 1.882, 0.265, TEAL);
  text(s, 1.2, 2.217, 3.628, 0.807, 'Cras in orci quis velit dapibus mollis. Pellentesque dictum a nisl vel mentum. Aenean eget neque sagittis, mattis', BODY);
  text(s, 1.2, 1.894, 2.507, 0.202, 'Bar Ratio', LEAD_D);
  pageNo(s, 21);
}

function slide22(s) {
  BACKDROPS[23](s);
  photo(s, 1.2, 0.6, 3.628, 3.75);
  photo(s, 5.433, 4.35, 3.033, 1.95);
  photo(s, 9.1, 4.35, 3.033, 1.95);
  text(s, 1.2, 5.273, 3.628, 0.807, 'Aliquam consequat id velit vel hendrerit. \nSed sed metus eleifend, imperdiet magna non, iaculis neque. Sed facilisis sapien', BODY);
  text(s, 1.2, 4.903, 2.467, 0.236, 'One', st(PM, 14, DARK));
  text(s, 5.433, 3.167, 3.067, 0.529, 'Nunc mioasjaa sapien, venenatis quis erat sed, sodales phare', BODY);
  hline(s, 9.1, 2.775, 3.033, GREY);
  text(s, 9.1, 3.144, 3.033, 0.529, 'Nunc mioasjaa sapien sodales pharetra justo conec adat . Curabitur sit amet', BODY);
  text(s, 9.1, 3.809, 3.033, 0.202, '56 – 89%', LEAD);
  text(s, 9.1, 1.24, 3.033, 0.807, 'Pellentesque dictum a nisl vel ferment  Aenean eget neque sagittis, mattis libero quis, rhoncus augue', BODY);
  text(s, 9.1, 2.167, 3.033, 0.202, '500+', LEAD);
  text(s, 5.433, 1.785, 3.067, 1.111, 'Time spent with \npet is never \nwasted.', H2);
  dots(s, 5.433, 1.243);
  pageNo(s, 22);
}

function slide23(s) {
  BACKDROPS[24](s);
  photo(s, 5.467, 5.1, 2.4, 1.8);
  photo(s, 0.6, 0.6, 4.867, 4.5, 'tl');
  rect(s, 5.467, 5.1, 2.4, 1.8, [TEAL, 30]);
  rect(s, 4.267, 3.9, 1.2, 1.2, [ORANGE, 30]);
  text(s, 1.206, 5.872, 1.625, 0.529, 'Venenatis masnan naub erat sed', BODY);
  text(s, 1.206, 5.599, 1.625, 0.202, 'Option One', LEAD);
  text(s, 3.242, 5.872, 1.625, 0.529, 'Nunc mioasjaa sapi venenatis', BODY);
  text(s, 3.242, 5.599, 1.625, 0.202, 'Option One', LEAD);
  text(s, 6.667, 3.083, 4.867, 1.085, 'Nunc mioasjaa sapien, venenatis quis erat sed, sodales pharetra justo. Donec adat iaculis diam. Pellentesque in risus namsa asut Etiam accumsan, ante eu luctus nulla ante iaculis tortorsit amet ante tellus', BODY);
  text(s, 6.667, 1.8, 4.867, 0.74, 'Live in the sunshine, \nswim the sea, drink the wild air.', H2);
  dots(s, 6.667, 1.2);
  pageNo(s, 23);
}

function slide24(s) {
  BACKDROPS[25](s);
  photo(s, 7.867, 1.2, 4.867, 5.1);
  photo(s, 0.6, 1.2, 1.2, 5.1);
  rect(s, 7.867, 1.2, 0.6, 5.1, [TEAL, 30], { flipH: true });
  text(s, 5.973, 3.243, 3.188, 0.303, 'Completed Project Experience', st(OR, 12, WHITE), { rotate: 90, margin: [0, 0, 3.6, 3.6] });
  text(s, 6.573, 3.243, 3.188, 0.303, '4000+', st(OB, 12, WHITE), { rotate: 90, margin: [0, 0, 3.6, 3.6] });
  text(s, 2.4, 1.742, 4.267, 0.74, 'When you cease \nto dream you cease to live.', st(OB, 22, DARK));
  text(s, 2.4, 2.68, 4.267, 1.085, 'Nulla non tortor at augue tincidunt elementum. Donec sollicitudin sollicitudin risus eu ultrices. Aliquam eget erat vitae nisl dignissim semper et vitae diam. Aliquam namus neque est efficitur', BODY);
  text(s, 2.4, 4.488, 1.933, 0.807, 'Sed pharetra eu tellus eget fringilla. Mauris ex nisi, convallis id', BODY);
  text(s, 4.733, 4.488, 1.933, 0.807, 'Nam nec ex congue \nrisus pretium vehiculana hendrerit et', BODY);
  rect(s, 1.2, 1.2, 0.6, 5.1, [SAND, 30], { flipH: true });
  text(s, -0.094, 3.243, 3.188, 0.303, 'Project Experience', st(OR, 12, WHITE), { rotate: 90, margin: [0, 0, 3.6, 3.6] });
  dots(s, 2.395, 1.2);
  text(s, 2.4, 4.196, 1.708, 0.236, 'First Point', st(PS, 12, SLATE), { lineSpacingMultiple: 1.167 });
  text(s, 4.733, 4.196, 1.708, 0.236, 'Second Point', st(PS, 12, SLATE), { lineSpacingMultiple: 1.167 });
  text(s, 2.394, 5.545, 0.944, 0.303, '250+', st(PM, 18, SAND));
  text(s, 4.733, 5.545, 0.944, 0.303, '400+', st(PM, 18, TEAL));
  pageNo(s, 24);
}

function slide25(s) {
  BACKDROPS[26](s);
  text(s, 4.283, 1.944, 1.985, 0.606, [{ text: '$', options: { superscript: true } }, { text: '40.00' }], st(PS, 36, WHITE));
  text(s, 4.283, 2.689, 1.985, 0.833, 'Tincidunt, pulvinar \nnibh in aliquam estlandit eratanma..', BODY_W);
  hline(s, 4.283, 3.75, 1.985, WHITE);
  hline(s, 1.495, 3.75, 1.985, WHITE);
  text(s, 4.283, 4.035, 1.985, 1.111, [{ text: 'Mauris eu nisl sed ', options: { bullet: { indent: 13.5 }, breakLine: true } }, { text: 'diam imperdiet iaculis ', options: { fontFace: PM, bullet: { indent: 13.5 }, breakLine: true } }, { text: 'Nam dimentum ', options: { bullet: { indent: 13.5 } } }], st(PL, 11, WHITE), { lineSpacingMultiple: 2 });
  text(s, 1.5, 1.944, 1.985, 0.606, [{ text: '$', options: { superscript: true } }, { text: '70.00' }], st(PS, 36, WHITE));
  text(s, 1.5, 2.689, 1.985, 0.833, 'Tincidunt, pulvinar \nnibh in aliquam estlandit eratanma..', BODY_W);
  text(s, 1.5, 4.035, 1.985, 1.481, [{ text: 'Aenean neque ', options: { fontFace: PL, bullet: { indent: 13.5 }, breakLine: true } }, { text: 'sagittis, mattis libero ', options: { bullet: { indent: 13.5 }, breakLine: true } }, { text: 'Rhoncus augue. ', options: { bullet: { indent: 13.5 }, breakLine: true } }, { text: 'Curabitur sit', options: { fontFace: PL, bullet: { indent: 13.5 } } }], st(PM, 11, WHITE), { lineSpacingMultiple: 2 });
  text(s, 7.867, 2.486, 3.667, 0.74, 'Price package\nTo choose for you', H2);
  dots(s, 7.867, 1.944);
  text(s, 7.867, 3.462, 3.667, 0.807, 'Nunc mioasjaa sapien, venenatis quis erat sed, sodales pharetra justo. Donec adat iaculis diam. Pellentesque in risus namsa', BODY);
  oval(s, 7.867, 4.692, 0.875, 0.875, [SILVER, 80]);
  oval(s, 8.004, 4.83, 0.6, 0.6, SAND);
  text(s, 7.999, 4.995, 0.589, 0.269, '4+', st(PS, 16, WHITE), { align: 'center' });
  text(s, 9.075, 4.996, 2.458, 0.529, 'Pellentesque dictumnan nisl vel fermenmante tellus.', BODY);
  text(s, 9.075, 4.716, 2.458, 0.236, 'Price', st(PM, 14, DARK));
  pageNo(s, 25);
}

function slide26(s) {
  BACKDROPS[27](s);
  text(s, 4.283, 1.968, 1.985, 0.539, [{ text: '$', options: { superscript: true } }, { text: '40.00' }], st(PM, 32, WHITE));
  text(s, 4.283, 2.581, 1.985, 0.833, [{ text: 'Mauris eu nisl sed ', options: { breakLine: true } }, { text: 'diam imperdiet iaculisan vitae ut metu' }], BODY_W);
  text(s, 4.283, 3.809, 1.985, 0.236, '500 items', st(PR, 14, WHITE));
  text(s, 4.283, 4.058, 1.985, 0.518, 'Aenean eget neque sagittis, mattis libero quis', st(PL, 11, WHITE), { lineSpacingMultiple: 1.4 });
  hline(s, 4.283, 3.621, 1.985, HONEY, 0.75);
  text(s, 4.283, 5.5, 1.487, 0.236, 'Great Choice', st(PM, 14, WHITE));
  text(s, 4.283, 5.301, 1.487, 0.185, 'Aenean eget', st(PL, 11, WHITE));
  oval(s, 5.911, 5.322, 0.356, 0.356, WHITE);
  text(s, 5.911, 5.4, 0.356, 0.202, 'A+', st(PM, 12, SLATE), { align: 'center' });
  text(s, 7.066, 2.468, 1.985, 0.539, [{ text: '$', options: { superscript: true } }, { text: '30.00' }], st(PM, 32, TEAL));
  text(s, 7.066, 3.081, 1.985, 0.833, [{ text: 'Mauris eu nisl sed ', options: { breakLine: true } }, { text: 'diam imperdiet iaculisan vitae ut metu' }], BODY);
  text(s, 7.066, 4.309, 1.985, 0.236, '300 items', st(PR, 14, SLATE));
  text(s, 7.066, 4.558, 1.985, 0.518, 'Aenean eget nque sagitti mattis libero quis', st(PL, 11, GREY), { lineSpacingMultiple: 1.4 });
  hline(s, 7.066, 4.121, 1.985, GREY, 0.75);
  text(s, 9.853, 2.468, 1.985, 0.539, [{ text: '$', options: { superscript: true } }, { text: '20.00' }], st(PM, 32, TEAL));
  text(s, 9.853, 3.081, 1.985, 0.833, [{ text: 'Mauris eu nisl sed ', options: { breakLine: true } }, { text: 'diam imperdiet iaculisan vitae ut metu' }], BODY);
  text(s, 9.853, 4.309, 1.985, 0.236, '200 items', st(PR, 14, SLATE));
  text(s, 9.853, 4.558, 1.985, 0.518, 'Aenean eget nque sagitti mattis libero quis', st(PL, 11, GREY), { lineSpacingMultiple: 1.4 });
  hline(s, 9.853, 4.121, 1.985, GREY, 0.75);
  text(s, 1.2, 2.892, 2.183, 0.37, 'Packages', H2);
  text(s, 1.2, 3.492, 2.183, 0.833, 'Nam tempus leo enim. \nvel quam tincidunt, pulvinar nibh aliquam est', st(PL, 11, GREY), { lineSpacingMultiple: 1.5 });
  oval(s, 1.2, 4.559, 0.607, 0.607, [TEAL, 50]);
  text(s, 2.007, 4.715, 0.978, 0.303, 'Good Price', st(PR, 12, INK), { margin: [0, 0, 3.6, 3.6] });
  oval(s, 1.307, 4.666, 0.393, 0.393, TEAL);
  text(s, 1.307, 4.709, 0.393, 0.303, '30', st(PR, 12, SNOW), { align: 'center', margin: [0, 0, 3.6, 3.6] });
  dots(s, 1.2, 2.334);
  pageNo(s, 26);
}

function slide27(s) {
  BACKDROPS[28](s);
  poly(s, 5.503, 1.273, 2.328, 4.953, PHOTO, [['M', 0.232, 0], [0.551, 0], [0.561, 0, 0.569, 0.008, 0.569, 0.018], [0.569, 0.054], [0.569, 0.119, 0.622, 0.172, 0.687, 0.172], [1.664, 0.172], [1.729, 0.172, 1.782, 0.119, 1.782, 0.054], [1.782, 0.018], [1.782, 0.008, 1.79, 0, 1.8, 0], [2.118, 0], [2.118, 0], [2.14, 0.004], [2.248, 0.025, 2.328, 0.121, 2.328, 0.232], [2.328, 4.721], [2.328, 4.848, 2.223, 4.953, 2.096, 4.953], [0.232, 4.953], [0.105, 4.953, 0, 4.848, 0, 4.721], [0, 0.232], [0, 0.105, 0.105, 0, 0.232, 0], []]);
  text(s, 1.2, 1.8, 3.067, 1.481, 'Never let the fear \nof striking out keep you from playing\nthe game.', H2);
  text(s, 1.2, 3.541, 3.067, 0.807, 'Nulla non tortor at augue tincidunt elementum. Donec sollicitudin sollicitudin risus eu', BODY);
  text(s, 9.656, 2.251, 2.478, 1.085, 'Nunc mioasjaa sapien nenatis naub erat sed, sodales pharetra justo conec adat iaculis diam. Aenean eget.', BODY);
  text(s, 9.067, 1.825, 0.589, 0.37, '01.', H2_D);
  text(s, 9.656, 1.931, 2.044, 0.202, 'Briefcase Set', LEAD_D);
  text(s, 9.656, 4.182, 2.478, 1.64, 'Nunc mioasjaa sapien, veenatis naub erat sed, sodales pharetra justo conec adat. \n\nNam nec ex congue risusloram pretium vehicula Hendra.', BODY);
  text(s, 9.067, 3.756, 0.589, 0.37, '02.', st(PM, 22, INK));
  text(s, 9.656, 3.873, 2.044, 0.202, 'Briefcase Set', LEAD);
  dots(s, 1.2, 1.2);
  oval(s, 1.206, 4.826, 0.6, 0.6, SLATE);
  text(s, 2.077, 5.293, 1.901, 0.529, 'Pellentesque dictumnan nisl vel fermenm. ', BODY);
  text(s, 1.2, 4.991, 0.589, 0.269, '4+', st(PS, 16, WHITE), { align: 'center' });
  text(s, 2.077, 5.013, 1.901, 0.202, 'Mockup', LEAD_D);
  oval(s, 7.316, 4.412, 1.203, 1.203, [TEAL, 50], { shadow: shadow(0.1, 10, 3) });
  oval(s, 7.493, 4.589, 0.85, 0.85, WHITE);
  poly(s, 7.754, 4.868, 0.34, 0.322, SLATE, [['M', 0.316, 0.116], [0.31, 0.116], [0.302, 0.095, 0.288, 0.075, 0.269, 0.06], [0.269, 0.06, 0.269, 0.06, 0.269, 0.059], [0.266, 0.05, 0.268, 0.039, 0.275, 0.028], [0.279, 0.022, 0.279, 0.015, 0.275, 0.009], [0.271, 0.003, 0.265, 0, 0.258, 0.001], [0.232, 0.004, 0.216, 0.016, 0.207, 0.027], [0.199, 0.025, 0.19, 0.023, 0.181, 0.022], [0.176, 0.021, 0.172, 0.025, 0.171, 0.029], [0.171, 0.034, 0.174, 0.038, 0.179, 0.039], [0.189, 0.04, 0.198, 0.042, 0.208, 0.045], [0.211, 0.046, 0.216, 0.045, 0.218, 0.042], [0.224, 0.033, 0.236, 0.021, 0.26, 0.018], [0.26, 0.018, 0.26, 0.018, 0.261, 0.018], [0.261, 0.019, 0.261, 0.019, 0.261, 0.019], [0.251, 0.034, 0.249, 0.05, 0.253, 0.064], [0.253, 0.064, 0.253, 0.064, 0.253, 0.064], [0.253, 0.067, 0.254, 0.069, 0.256, 0.071], [0.275, 0.086, 0.289, 0.106, 0.296, 0.127], [0.297, 0.131, 0.3, 0.134, 0.304, 0.134], [0.304, 0.134, 0.304, 0.134, 0.304, 0.134], [0.316, 0.134], [0.32, 0.134, 0.323, 0.137, 0.323, 0.141], [0.323, 0.18], [0.323, 0.184, 0.32, 0.187, 0.316, 0.187], [0.305, 0.187], [0.301, 0.187, 0.298, 0.189, 0.296, 0.193], [0.29, 0.213, 0.278, 0.232, 0.261, 0.247], [0.261, 0.247, 0.261, 0.247, 0.261, 0.247], [0.261, 0.247], [0.26, 0.248, 0.259, 0.249, 0.259, 0.249], [0.259, 0.249, 0.259, 0.249, 0.259, 0.249], [0.24, 0.266, 0.238, 0.286, 0.238, 0.291], [0.238, 0.292, 0.238, 0.292, 0.238, 0.292], [0.238, 0.304], [0.186, 0.304], [0.186, 0.292], [0.186, 0.29, 0.185, 0.287, 0.183, 0.286], [0.181, 0.284, 0.179, 0.283, 0.176, 0.284], [0.171, 0.284, 0.165, 0.285, 0.159, 0.285], [0.152, 0.285, 0.144, 0.284, 0.137, 0.283], [0.135, 0.283, 0.132, 0.284, 0.13, 0.285], [0.128, 0.287, 0.127, 0.289, 0.127, 0.292], [0.127, 0.305], [0.075, 0.305], [0.075, 0.301], [0.075, 0.301, 0.075, 0.3, 0.075, 0.3], [0.075, 0.291, 0.072, 0.267, 0.053, 0.248], [0.044, 0.239, 0.036, 0.226, 0.03, 0.212], [0.03, 0.212, 0.03, 0.212, 0.03, 0.211], [0.026, 0.204, 0.022, 0.195, 0.02, 0.187], [0.02, 0.187, 0.02, 0.187, 0.02, 0.187], [0.02, 0.187, 0.02, 0.187, 0.02, 0.187], [0.018, 0.178, 0.017, 0.17, 0.017, 0.161], [0.017, 0.155, 0.018, 0.149, 0.019, 0.143], [0.02, 0.139, 0.016, 0.134, 0.012, 0.133], [0.007, 0.133, 0.002, 0.136, 0.002, 0.141], [0.001, 0.147, 0, 0.154, 0, 0.161], [0, 0.171, 0.001, 0.181, 0.003, 0.19], [0.003, 0.191, 0.003, 0.191, 0.004, 0.191], [0.004, 0.191, 0.004, 0.191, 0.004, 0.191], [0.006, 0.201, 0.009, 0.21, 0.014, 0.219], [0.022, 0.236, 0.031, 0.249, 0.041, 0.26], [0.055, 0.275, 0.058, 0.293, 0.058, 0.301], [0.058, 0.305], [0.058, 0.305, 0.058, 0.305, 0.058, 0.305], [0.058, 0.306, 0.058, 0.306, 0.058, 0.306], [0.059, 0.315, 0.066, 0.322, 0.075, 0.322], [0.127, 0.322], [0.137, 0.322, 0.145, 0.314, 0.145, 0.305], [0.145, 0.301], [0.153, 0.302, 0.161, 0.302, 0.169, 0.302], [0.169, 0.305], [0.169, 0.314, 0.176, 0.322, 0.186, 0.322], [0.238, 0.322], [0.248, 0.322, 0.255, 0.314, 0.255, 0.305], [0.255, 0.292], [0.255, 0.289, 0.256, 0.275, 0.27, 0.262], [0.271, 0.262, 0.271, 0.261, 0.272, 0.261], [0.272, 0.261, 0.272, 0.261, 0.272, 0.26], [0.272, 0.26, 0.272, 0.26, 0.273, 0.26], [0.29, 0.244, 0.303, 0.225, 0.311, 0.204], [0.316, 0.204], [0.329, 0.204, 0.34, 0.193, 0.34, 0.18], [0.34, 0.141], [0.34, 0.127, 0.329, 0.116, 0.316, 0.116], []]);
  poly(s, 7.743, 4.836, 0.17, 0.169, SLATE, [['M', 0.085, 0.169], [0.132, 0.169, 0.17, 0.131, 0.17, 0.085], [0.17, 0.038, 0.132, 0, 0.085, 0], [0.038, 0, 0, 0.038, 0, 0.085], [0, 0.131, 0.038, 0.169, 0.085, 0.169], [], ['M', 0.085, 0.017], [0.122, 0.017, 0.153, 0.047, 0.153, 0.085], [0.153, 0.122, 0.122, 0.152, 0.085, 0.152], [0.048, 0.152, 0.017, 0.122, 0.017, 0.085], [0.017, 0.047, 0.048, 0.017, 0.085, 0.017], []]);
  pageNo(s, 27);
}

function slide28(s) {
  BACKDROPS[29](s);
  rrect(s, 5.418, 1.402, 0.797, 0.393, ORANGE, 0.132, { rotate: 270 });
  text(s, 5.618, 1.447, 0.393, 0.303, '450+', st(OR, 12, SNOW), { align: 'center', rotate: 270, margin: [0, 0, 3.6, 3.6] });
  text(s, 4.393, 4.87, 2.543, 0.303, 'Project Experience', st(OR, 12, INK), { rotate: 270, margin: [0, 0, 3.6, 3.6] });
  text(s, 4.675, 4.929, 2.543, 0.185, 'erat sed, sodales pharetra justo', NOTE, { rotate: 270 });
  hline(s, 5.513, 2.421, 0.549, GREY);
  text(s, 5.618, 2.934, 0.393, 0.303, '120+', st(OR, 12, INK), { align: 'center', rotate: 270, margin: [0, 0, 3.6, 3.6] });
  text(s, 7.867, 3.453, 4.267, 1.085, 'Nunc mioasjaa sapien, venenatis quis erat sed, \nsodales pharetra justo. Donec adat iaculis dia in risus namsa asutan Etiam accumsan, ante eu luctus nulla ante iaculis tortorsit amet ante.', BODY);
  text(s, 7.867, 1.8, 4.267, 1.111, 'Live in the sunshine, \nswim the sea, drink the wild air.', H2);
  rect(s, 7.867, 5.435, 4.263, 0.265, [SILVER, 50]);
  rect(s, 7.867, 5.435, 2.867, 0.265, ORANGE);
  text(s, 7.869, 5.102, 1.556, 0.185, 'Release Process', NOTE);
  text(s, 10.671, 5.976, 1.462, 0.202, [{ text: '35 ' }, { text: '- ', options: { color: DARK } }, { text: '25%', options: { color: ORANGE } }], st(PM, 12, TEAL), { align: 'right' });
  rect(s, 7.867, 5.435, 1.733, 0.265, TEAL);
  dots(s, 7.867, 1.2);
  photo(s, 1.316, 1.32, 3.423, 4.86, 'tl tr br bl', 0.159);
  pageNo(s, 28);
}

function slide29(s) {
  BACKDROPS[30](s);
  text(s, 1.2, 2.795, 4.266, 0.807, 'Etiam ultrices dolor et enim tempus egestas. Mauris ullamcorper, ex ac semper finibus, diam nulla mollis ex, non pharetra ex nisl sed', BODY);
  text(s, 1.2, 1.8, 4.266, 0.74, 'You never learn much from hearing yourself speak.', H2);
  pill(s, 1.202, 4.566, 4.263, 0.265, SILVER);
  pill(s, 1.2, 5.667, 4.263, 0.265, SILVER);
  pill(s, 1.202, 4.566, 3.408, 0.265, ORANGE);
  pill(s, 1.2, 5.667, 2.129, 0.265, TEAL);
  text(s, 1.202, 4.258, 1.947, 0.202, 'Development Process', SMALL);
  text(s, 4.004, 4.236, 1.462, 0.202, '80%', LEAD_D, { align: 'right' });
  text(s, 1.202, 5.359, 1.947, 0.202, 'Release Process', SMALL);
  text(s, 4.004, 5.337, 1.462, 0.202, '55%', LEAD_D, { align: 'right' });
  rrect(s, 8.734, 4.601, 1.333, 1.331, WHITE, 0.063, { shadow: shadow(0.1, 10, 5) });
  text(s, 8.942, 5.454, 0.92, 0.185, 'Mioasjaan', NOTE, { align: 'center' });
  text(s, 9.037, 4.918, 0.735, 0.539, [{ text: '72' }, { text: '+', options: { fontSize: 12 } }], st(PM, 32, TEAL), { align: 'center' });
  text(s, 10.489, 5.439, 0.92, 0.185, 'Lorem an', NOTE, { align: 'center' });
  text(s, 10.583, 4.903, 0.735, 0.539, '25', st(PM, 32, SLATE), { align: 'center' });
  text(s, 7.387, 5.454, 0.92, 0.185, 'Mioasjaa', NOTE, { align: 'center' });
  text(s, 7.481, 4.918, 0.735, 0.539, '14', st(PM, 32, SLATE), { align: 'center' });
  dots(s, 1.2, 1.2);
  photo(s, 7.306, 1.39, 4.156, 2.556, 'tl tr br bl', 0.043);
  pageNo(s, 29);
}

function slide30(s) {
  BACKDROPS[31](s);
  s.addTable([
    [{ text: 'Basic', options: { fontFace: PM, fontSize: 14, color: WHITE, fill: TEAL, border: [{ type: 'none' }, { type: 'none' }, { type: 'none' }, { type: 'none' }] } }, { text: 'Advance', options: { fontFace: PM, fontSize: 14, color: WHITE, fill: SAND, border: [{ type: 'none' }, { type: 'none' }, { type: 'none' }, { type: 'none' }] } }],
    [{ text: 'Free Space', options: { fontFace: PL, fontSize: 11, color: GREY, border: [{ type: 'none' }, { type: 'none' }, { pt: 0.25, color: GREY }, { type: 'none' }] } }, { text: 'Free Space', options: { fontFace: PL, fontSize: 11, color: GREY, border: [{ type: 'none' }, { type: 'none' }, { pt: 0.25, color: GREY }, { type: 'none' }] } }],
    [{ text: '25+ Template', options: { fontFace: PL, fontSize: 11, color: GREY, border: [{ pt: 0.25, color: GREY }, { type: 'none' }, { pt: 0.25, color: GREY }, { type: 'none' }] } }, { text: '25+ Template', options: { fontFace: PL, fontSize: 11, color: GREY, border: [{ pt: 0.25, color: GREY }, { type: 'none' }, { pt: 0.25, color: GREY }, { type: 'none' }] } }],
    [{ text: '20TB Storage', options: { fontFace: PL, fontSize: 11, color: GREY, border: [{ pt: 0.25, color: GREY }, { type: 'none' }, { pt: 0.25, color: GREY }, { type: 'none' }] } }, { text: '20TB Storage', options: { fontFace: PL, fontSize: 11, color: GREY, border: [{ pt: 0.25, color: GREY }, { type: 'none' }, { pt: 0.25, color: GREY }, { type: 'none' }] } }],
    [{ text: 'Free Share', options: { fontFace: PL, fontSize: 11, color: GREY, border: [{ pt: 0.25, color: GREY }, { type: 'none' }, { pt: 2.25, color: TEAL }, { type: 'none' }] } }, { text: 'Free Share', options: { fontFace: PL, fontSize: 11, color: GREY, border: [{ pt: 0.25, color: GREY }, { type: 'none' }, { pt: 2.25, color: SAND }, { type: 'none' }] } }],
    [{ text: 'Free', options: { fontFace: PM, fontSize: 14, color: DARK, border: [{ pt: 2.25, color: TEAL }, { type: 'none' }, { type: 'none' }, { type: 'none' }] } }, { text: '$20', options: { fontFace: PM, fontSize: 14, color: DARK, border: [{ pt: 2.25, color: SAND }, { type: 'none' }, { type: 'none' }, { type: 'none' }] } }],
  ], { x: 1.2, y: 1.2, colW: [3.333, 3.333], rowH: 0.85, align: 'center', valign: 'middle', margin: 0 });
  text(s, 8.464, 2.671, 3.676, 0.251, 'Nunc mioasjaa sapiennai maan.', BODY);
  text(s, 8.464, 3.766, 3.676, 0.529, 'Nunc mioasjaa sapien erat sed, sodales rictuma pharetra justo conec adat iaculis diam ', BODY);
  text(s, 8.479, 3.444, 3.067, 0.236, 'Concept', st(PM, 14, ORANGE));
  text(s, 8.479, 5.15, 3.583, 0.529, 'Erat sed, sodales pharetra justo conec adatan iaculis diam. Mauris ullamcorper.', BODY);
  text(s, 8.479, 4.829, 2.989, 0.236, 'Release', st(PM, 14, TEAL));
  text(s, 8.464, 1.8, 3.67, 0.74, 'The purpose \nof lives is be happy.', H2);
  dots(s, 8.464, 1.2);
  pageNo(s, 30);
}

function slide31(s) {
  BACKDROPS[32](s);
  s.addTable([
    [{ text: 'Advance', options: { fontFace: PM, fontSize: 14, color: WHITE, fill: ORANGE, border: [{ type: 'none' }, { type: 'none' }, { type: 'none' }, { type: 'none' }] } }],
    [{ text: 'Free Space', options: { fontFace: PL, fontSize: 11, color: GREY, border: [{ type: 'none' }, { pt: 0.5, color: CLOUD }, { pt: 0.5, color: CLOUD }, { pt: 0.5, color: CLOUD }] } }],
    [{ text: '25+ Template', options: { fontFace: PL, fontSize: 11, color: GREY, border: [{ pt: 0.5, color: CLOUD }, { pt: 0.5, color: CLOUD }, { pt: 0.5, color: CLOUD }, { pt: 0.5, color: CLOUD }] } }],
  ], { x: 8.833, y: 3.75, colW: [3.3], rowH: 0.85, align: 'center', valign: 'middle', margin: 0 });
  s.addTable([
    [{ text: 'Basic', options: { fontFace: PM, fontSize: 14, color: WHITE, fill: TEAL, border: [{ type: 'none' }, { type: 'none' }, { type: 'none' }, { type: 'none' }] } }],
    [{ text: 'Free Space', options: { fontFace: PL, fontSize: 11, color: GREY, border: [{ type: 'none' }, { pt: 0.5, color: CLOUD }, { pt: 0.5, color: CLOUD }, { pt: 0.5, color: CLOUD }] } }],
    [{ text: '25+ Template', options: { fontFace: PL, fontSize: 11, color: GREY, border: [{ pt: 0.5, color: CLOUD }, { pt: 0.5, color: CLOUD }, { pt: 0.5, color: CLOUD }, { pt: 0.5, color: CLOUD }] } }],
  ], { x: 5.533, y: 1.2, colW: [3.3], rowH: 0.85, align: 'center', valign: 'middle', margin: 0 });
  text(s, 6.133, 4.962, 2.1, 0.529, 'Nunc mioasjaa sapi veneis masnan naubana', BODY, { align: 'center' });
  text(s, 6.133, 4.718, 2.1, 0.202, 'Table Two', LEAD, { align: 'center' });
  text(s, 9.537, 2.412, 1.892, 0.529, 'Nunc ansmsapi venenati masnan naubana', BODY, { align: 'center' });
  text(s, 9.537, 2.168, 1.892, 0.202, 'Table One', LEAD, { align: 'center' });
  text(s, 1.2, 2.433, 3.733, 0.74, 'The healthiest response to life is joy.', H2);
  text(s, 1.2, 3.472, 3.733, 2.195, 'Nulla non tortor at augue tincidunt elementum. Donec sollicitudin sollicitudin risus eu ultrices. Aliquam eget erat vitae nisl dignissim semper et vitae diam. Aliquam rhoncus\n\nneque est efficitur velit, in tincidunt urna risus sit amet libero. Cras vel fermentum metus. Integer luctus convallis massa,amnu', BODY);
  dots(s, 1.2, 1.833);
  pageNo(s, 31);
}

function slide32(s) {
  BACKDROPS[33](s);
  text(s, 11.553, 1.307, 0.58, 0.202, '100', LEAD, { align: 'right' });
  hline(s, 5.467, 3.427, 6.667, CLOUD);
  hline(s, 5.467, 2.315, 6.667, CLOUD);
  hline(s, 5.467, 1.204, 6.667, CLOUD);
  hline(s, 5.467, 4.539, 6.667, CLOUD);
  text(s, 11.553, 2.419, 0.58, 0.202, '75', SMALL, { align: 'right' });
  text(s, 11.553, 3.53, 0.58, 0.202, '50', SMALL, { align: 'right' });
  text(s, 11.553, 4.642, 0.58, 0.202, '25', SMALL, { align: 'right' });
  text(s, 11.553, 5.75, 0.58, 0.202, '0', SMALL, { align: 'right' });
  pill(s, 6.197, 4.479, 0.491, 1.178, MINT);
  rect(s, 6.197, 5.067, 0.491, 0.586, TEAL);
  pill(s, 7.654, 2.904, 0.491, 2.753, MINT);
  rect(s, 7.654, 4.156, 0.491, 1.498, TEAL);
  pill(s, 9.112, 1.578, 0.491, 4.08, SAND);
  rect(s, 9.112, 2.804, 0.491, 2.849, ORANGE);
  pill(s, 10.57, 2.557, 0.491, 3.101, SAND);
  rect(s, 10.57, 4.413, 0.491, 1.24, ORANGE);
  hline(s, 5.467, 5.651, 6.667, INK);
  text(s, 5.819, 5.83, 1.246, 0.202, '27%', LEAD, { align: 'center' });
  text(s, 5.705, 6.076, 1.474, 0.202, 'Cran orciquis', SMALL, { align: 'center' });
  text(s, 7.277, 5.83, 1.246, 0.202, '55%', LEAD, { align: 'center' });
  text(s, 7.163, 6.076, 1.474, 0.202, 'Cran orciquis', SMALL, { align: 'center' });
  text(s, 8.735, 5.83, 1.246, 0.202, '85%', st(PM, 12, GOLD), { align: 'center' });
  text(s, 8.62, 6.076, 1.474, 0.202, 'Cran orciquis', SMALL, { align: 'center' });
  text(s, 10.193, 5.83, 1.246, 0.202, '70%', LEAD, { align: 'center' });
  text(s, 10.078, 6.076, 1.474, 0.202, 'Cran orciquis', SMALL, { align: 'center' });
  text(s, 1.2, 2.677, 3.667, 0.529, 'Nunc mioasjaa sapien, venenatis quis erat sed, sodales', BODY);
  text(s, 1.2, 1.8, 3.667, 0.74, 'The purpose of lives is be happy.', st(OB, 22, DARK));
  text(s, 1.387, 4.217, 3.479, 0.529, 'Nunc mioasjaa sapien erat sed, sodales rictuma pharetra justo', BODY);
  text(s, 1.2, 3.847, 3.667, 0.202, 'Concept', st(PM, 12, ORANGE), { bullet: { indent: 13.5 } });
  text(s, 1.387, 5.543, 3.479, 0.529, 'Erat sed, sodales pharetra justo conec adatan iaculis diam. ', BODY);
  text(s, 1.2, 5.173, 3.667, 0.202, 'Release', st(PM, 12, TEAL), { bullet: { indent: 13.5 } });
  dots(s, 1.2, 1.2);
  pageNo(s, 32);
}

function slide33(s) {
  BACKDROPS[33](s);
  oval(s, 1.2, 1.2, 5.089, 5.1, TEAL);
  pieSlice(s, 1.2, 1.211, 5.089, 5.089, SAND, 347.471, 270);
  pieSlice(s, 1.2, 1.2, 5.089, 5.089, FOG, 154.849, 270);
  oval(s, 1.669, 1.68, 4.151, 4.151, WHITE, { shadow: shadow(0.2, 10, 5) });
  text(s, 1.935, 2.829, 3.618, 0.899, 'Maecenas ligula orci, feugiat \nsuscipit purus sed, suscipit lacinia tortor suspendisse commodo', st(PL, 11, SLATE), { align: 'center', margin: [7.2, 7.2, 3.6, 3.6], lineSpacingMultiple: 1.5 });
  text(s, 2.753, 2.493, 1.982, 0.337, 'Information', st(PM, 14, SLATE), { align: 'center', margin: [7.2, 7.2, 3.6, 3.6] });
  text(s, 2.151, 3.982, 3.186, 0.64, [{ text: '700' }, { text: ' ', options: { color: AMBER } }, { text: '|', options: { fontFace: PL, color: GREY } }, { text: ' ', options: { color: WHITE } }, { text: '420', options: { fontFace: PL, color: TEAL } }], st(PM, 32, SAND), { align: 'center', margin: [7.2, 7.2, 3.6, 3.6] });
  text(s, 2.867, 4.865, 1.755, 0.303, 'Chart Ratio', st(OB, 12, WHITE), { align: 'center', margin: [7.2, 7.2, 3.6, 3.6] });
  ringArc(s, 1.923, 1.934, 3.643, 3.643, 270, 154.164, SLATE, 0.75, 'oval');
  text(s, 2.753, 4.604, 1.982, 0.286, 'Maecenas ligorci', st(PL, 11, SLATE), { align: 'center', margin: [7.2, 7.2, 3.6, 3.6] });
  text(s, 7.267, 3.391, 4.267, 1.085, 'Nunc mioasjaa sapien, venenatis quis erat sed\nsodales pharetra justo. Donec adat iaculis diamansan Pellentesque in risus namsa asut magna viverrana consequat ac in augueante tellusans', BODY);
  text(s, 7.267, 2.109, 4.267, 0.74, 'Every strike brings me closer to the next home run', H2);
  text(s, 7.267, 5.436, 1.78, 0.555, 'Cras in orci quis velit dapibus mollis', BODY);
  text(s, 7.267, 4.727, 1.78, 0.709, '24+', st(PS, 36, TEAL), { lineSpacingMultiple: 1.167 });
  text(s, 9.753, 5.436, 1.78, 0.555, 'Cras in orci quis velit dapibus mollis', BODY);
  text(s, 9.753, 4.727, 1.78, 0.709, '50+', st(PS, 36, SAND), { lineSpacingMultiple: 1.167 });
  dots(s, 7.267, 1.509);
  pageNo(s, 33);
}

function slide34(s) {
  BACKDROPS[34](s);
  text(s, 7.867, 1.8, 3.67, 1.111, 'Dogs do speak, but only to those who know how to listen.', H2);
  text(s, 7.867, 3.304, 3.67, 0.807, 'Nulla non tortor at augue tincidunt elementum. Donec sollicitudin sollicitudin risus eu ultrices. Aliquam eget erat vitae nisl.', BODY);
  text(s, 8.456, 5.203, 3.163, 0.807, 'Maecenas ligula orci, feugiat ansa\nsuscipit purus sed, suscipit lacinia tortor commodo nisl a augue', BODY);
  text(s, 7.867, 4.753, 0.589, 0.37, 'Aa.', st(PM, 22, INK));
  text(s, 8.456, 4.879, 3.678, 0.202, 'Option One', LEAD_D);
  dots(s, 7.867, 1.2);
  text(s, 1.837, 6.316, 3.593, 0.185, 'Pellentesque dictumnan nisl vel fermenm. ', st(PL, 11, WHITE), { align: 'center' });
  text(s, 2.683, 6.036, 1.901, 0.202, 'Mockup', LEAD_W, { align: 'center' });
  photo(s, 0.6, 0.6, 6.067, 5.1, 'tl');
  pageNo(s, 34);
}

function slide35(s) {
  BACKDROPS[35](s);
  photo(s, 1.2, 1.2, 10.933, 5.1);
  rect(s, 1.2, 1.2, 5.467, 5.7, [ORANGE, 40]);
  text(s, 2.4, 3.668, 3.067, 0.807, 'Nulla non tortor at augue tincidunt elementum. Donec sollicitudinn masam sollicitudin namsa', BODY_W);
  text(s, 2.4, 2.998, 3.067, 0.37, 'Closing Message', st(PM, 22, WHITE));
  text(s, 2.4, 4.831, 3.067, 0.185, 'conec adat . Curabitur sit amet', st(PL, 11, WHITE));
  text(s, 2.4, 5.075, 3.067, 0.236, 'Thank You', st(PM, 14, WHITE));
  rrect(s, 2.4, 2.258, 1.608, 0.393, WHITE, 0.072, { rotate: 180, shadow: shadow(0.2, 4, 3) });
  text(s, 2.603, 2.303, 1.203, 0.303, [{ text: 'Pet', options: { color: TEAL } }, { text: ' ', options: { color: COAL } }, { text: 'Shop' }], st(PR, 12, ORANGE), { align: 'center', margin: [0, 0, 3.6, 3.6] });
  pageNo(s, 35);
}

/* ----------------------------------------------------------------- build */
const SLIDES = [
  slide01, slide02, slide03, slide04, slide05,
  slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15,
  slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25,
  slide26, slide27, slide28, slide29, slide30,
  slide31, slide32, slide33, slide34, slide35,
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE';
  pptx.title = 'Paws. - Animal Care Presentation';
  SLIDES.forEach((draw) => draw(pptx.addSlide()));
  return pptx.writeFile({
    fileName: path.join(__dirname, '013f0be6-ec13-4eee-8b8a-f88b106a40c1_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote ' + f)).catch((e) => { console.error(e); process.exit(1); });
