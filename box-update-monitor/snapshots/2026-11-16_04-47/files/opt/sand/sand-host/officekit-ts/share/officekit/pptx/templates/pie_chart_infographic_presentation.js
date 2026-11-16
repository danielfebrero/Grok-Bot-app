#!/usr/bin/env node
/**
 * Pie Chart Infographic Presentation - 31 slides, 13.333 x 7.5 in.
 * Rebuilt with pptxgenjs only; run `node <this file>` to write the .pptx next to it.
 * Raster artwork from the original template is replaced by flat colour placeholders.
 */
'use strict';
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- theme
const C = {
  navy:  '537188',
  gold:  'CBB279',
  sand:  'E1D4BB',
  mist:  'EEEEEE',
  taupe: '967E76',
  clay:  'D7C0AE',
  ink:   '2D3847',
  white: 'FFFFFF',
  paper: 'FDFDFD',
  body:  '808080',
  dark:  '404040',
  slate: '6E85A3',
  haze:  'F2F2F2',
};
const F = { bold: 'Montserrat SemiBold', light: 'Montserrat Light' };
const SH = {
  card:  { type: 'outer', color: '000000', opacity: 0.1, blur: 80, offset: 30, angle: 35 },
  deep:  { type: 'outer', color: '000000', opacity: 0.32, blur: 80, offset: 30, angle: 35 },
  soft:  { type: 'outer', color: '000000', opacity: 0.1, blur: 26, offset: 27, angle: 90 },
  lift:  { type: 'outer', color: '0D0D0D', opacity: 0.14, blur: 45, offset: 31, angle: 45 },
};
const tint = (color) => ({ type: 'outer', color, opacity: 0.1, blur: 30, offset: 35, angle: 135 });

// ---------------------------------------------------------------- primitives
const sh = (s, type, x, y, w, h, o = {}) => s.addShape(type, Object.assign({ x, y, w, h }, o));

const tx = (s, text, x, y, w, h, o = {}) => s.addText(text, Object.assign(
  { x, y, w, h, fontFace: F.light, fontSize: 11, color: C.body, valign: 'top', margin: [3.6, 7.2, 3.6, 7.2] }, o));

// Paths are stored as compact commands: 'M x,y', 'L x,y', 'C x1,y1 x2,y2 x,y', 'Z' in a 0..1 box.
function pathPoints(d, w, h) {
  const out = [];
  for (const tok of d.match(/[MLCZ][^MLCZ]*/g) || []) {
    if (tok[0] === 'Z') { out.push({ close: true }); continue; }
    const n = tok.slice(1).trim().split(/[ ,]+/).map(Number);
    if (tok[0] === 'M') out.push({ x: n[0] * w, y: n[1] * h, moveTo: true });
    else if (tok[0] === 'L') out.push({ x: n[0] * w, y: n[1] * h });
    else out.push({ x: n[4] * w, y: n[5] * h, curve: { type: 'cubic', x1: n[0] * w, y1: n[1] * h, x2: n[2] * w, y2: n[3] * h } });
  }
  return out;
}

const poly = (s, d, x, y, w, h, o = {}) =>
  s.addShape('custGeom', Object.assign({ x, y, w, h, points: pathPoints(d, w, h) }, o));

// Flat-colour stand-in for a bitmap image in the original deck.
const imgBox = (s, x, y, w, h) => {
  sh(s, 'roundRect', x, y, w, h, { fill: C.haze, rectRadius: 0.08 });
  tx(s, '[image]', x, y + h / 2 - 0.14, w, 0.28, { fontSize: 8, align: 'center', color: 'A6A6A6' });
};

const icon = (s, name, x, y, w, h, color) =>
  s.addShape('custGeom', { x, y, w, h, fill: color, points: pathPoints(ICONS[name], w / 100, h / 100) });

// Every slide carries the same two-tone headline.
function title(s, x, y, w, h, o = {}) {
  tx(s, [{ text: 'Pie Chart ', options: { color: C.navy, bold: true } },
         { text: o.tail || 'Infographic', options: { color: C.dark } }],
     x, y, w, h, { fontFace: F.bold, fontSize: o.fontSize || 40, align: o.align });
}

// White pill holding a progress ring, a percentage, a heading and optional body copy.
function statCard(s, x, y, w, o) {
  sh(s, 'roundRect', x, y, w, 1.15, { fill: C.white, rectRadius: 0.04, shadow: SH.card });
  sh(s, 'donut', x + 0.34, y + 0.17, 0.8, 0.8, { fill: C.haze, rectRadius: 0.1 });
  sh(s, 'blockArc', x + 0.34, y + 0.17, 0.8, 0.8,
     { fill: o.ring, angleRange: [50.4, 1.5], arcThicknessRatio: 0.26 });
  tx(s, o.pct, x + 0.45, y + 0.42, 0.59, 0.3, { fontFace: F.bold, fontSize: 12, bold: true, align: 'center', color: '000000' });
  tx(s, o.head, x + 1.18, y + 0.16, 1.73, 0.34, { fontFace: F.bold, fontSize: 14, bold: true, color: o.headColor });
  if (o.body) tx(s, o.body, x + 1.18, y + 0.41, o.bodyW || 1.73, 0.53,
     { fontFace: o.bodyBold ? F.bold : F.light, lineSpacingMultiple: 1.2 });
}

// Row of colour swatch + caption pairs used under the chart graphics.
function legend(s, x, y, step, items, o = {}) {
  items.forEach(([color, label], i) => {
    sh(s, 'rect', x + i * step, y, o.sw || 0.1, o.sw || 0.1, { fill: color });
    tx(s, label, x + i * step + (o.tx || 0.11), y + (o.ty || -0.09), o.tw || 0.9, 0.35, { fontSize: o.fs || 11 });
  });
}

// The isometric 3-D pie illustration: eight facets [x, y, w, h, path] in a 0..1 box.
const ISO_FACETS = [
  [0, 0.37, 1, 0.63, 'M0,0.42 C0,0.74 0.22,1 0.5,1 C0.78,1 1,0.74 1,0.42 L1,0 L0,0 L0,0.4'],
  [0, 0.01, 1, 0.72, 'M0,0.5 C0,0.78 0.22,1 0.5,1 C0.78,1 1,0.78 1,0.5 C1,0.22 0.78,0 0.5,0 C0.22,0 0,0.22 0,0.5 Z'],
  [0, 0.37, 0.17, 0.54, 'M1,0.51 L1,1 C1,1 0.07,0.84 0,0.52 L0,0 C0,0 -0.09,0.28 1,0.51 Z'],
  [0, 0.35, 0.5, 0.29, 'M1,0.06 L0.35,1 C0.35,1 -0.03,0.64 0,0 L1,0.06 Z'],
  [0, 0.03, 0.5, 0.34, 'M1,1 L0.62,0 C0.62,0 0.03,0.23 0,0.95 L1,1 Z'],
  [0.31, 0, 0.58, 0.37, 'M0.32,1 L1,0.39 C1,0.39 0.64,-0.22 0,0.09 L0.32,1 Z'],
  [0.5, 0.14, 0.5, 0.58, 'M0.78,0 C0.78,0 1.1,0.23 0.97,0.55 C0.97,0.55 0.85,0.9 0.25,1 L0,0.39 L0.78,0 Z'],
  [0.63, 0.39, 0.37, 0.6, 'M1,0 L1,0.43 C1,0.43 1.02,0.85 0.02,1 L0,1 L0,0.55 C0,0.55 0.92,0.46 1,0 Z'],
];
function isoChart(s, x, y, w, h, colors, flipH) {
  ISO_FACETS.forEach(([bx, by, bw, bh, d], i) => {
    poly(s, d, x + (flipH ? 1 - bx - bw : bx) * w, y + by * h, bw * w, bh * h,
      { fill: colors[i], flipH: !!flipH });
  });
}

// ---------------------------------------------------------------- vector glyphs (0..100 box)
const ICONS = {
  briefcase: 'M25,20 L25,5 C25,2 27,0 30,0 L70,0 C73,0 75,2 75,5 L75,20 L95,20 C98,20 100,22 100,25 L100,95 C100,98 98,100 95,100 L5,100 C2,100 0,98 0,95 L0,25 C0,22 2,20 5,20 L25,20 Z M65,30 L35,30 L35,90 L65,90 L65,30 Z M25,30 L10,30 L10,90 L25,90 L25,30 Z M75,30 L75,90 L90,90 L90,30 L75,30 Z M35,10 L35,20 L65,20 L65,10 L35,10 Z',
  boxDown: 'M10,0 L90,0 L100,22 L100,94 C100,98 98,100 95,100 L5,100 C2,100 0,98 0,94 L0,22 L10,0 Z M90,33 L10,33 L10,89 L90,89 L90,33 Z M89,22 L84,11 L16,11 L11,22 L89,22 Z M55,61 L70,61 L50,83 L30,61 L45,61 L45,39 L55,39 L55,61 Z',
  archiveBox: 'M5,39 L0,39 L0,6 C0,2 2,0 5,0 L95,0 C98,0 100,2 100,5 L100,6 L100,39 L95,39 L95,94 C95,98 93,100 90,100 L10,100 C7,100 5,98 5,94 L5,39 Z M85,39 L15,39 L15,89 L85,89 L85,39 Z M10,11 L10,28 L90,28 L90,11 L10,11 Z M35,50 L65,50 L65,61 L35,61 L35,50 Z',
  drawer: 'M0,5 C0,2 2,0 6,0 L94,0 C98,0 100,2 100,5 L100,95 C100,98 98,100 94,100 L6,100 C2,100 0,98 0,95 L0,5 Z M89,45 L89,10 L11,10 L11,45 L89,45 Z M89,55 L11,55 L11,90 L89,90 L89,55 Z M33,20 L67,20 L67,30 L33,30 L33,20 Z M33,65 L67,65 L67,75 L33,75 L33,65 Z',
  globe: 'M50,100 C22,100 0,78 0,50 C0,22 22,0 50,0 C78,0 100,22 100,50 C100,78 78,100 50,100 Z M39,88 C34,78 31,67 30,55 L10,55 C12,71 23,84 39,88 Z M40,55 C41,67 44,79 50,89 C56,78 59,67 60,55 L40,55 Z M90,55 L70,55 C69,67 66,78 61,88 C77,84 88,71 90,55 Z M10,45 L30,45 C31,33 34,22 39,12 C23,16 12,29 10,45 Z M40,45 L60,45 C59,33 56,22 50,11 C44,22 41,33 40,45 Z M61,12 C66,22 69,33 70,45 L90,45 C88,29 77,16 61,12 Z',
  badge: 'M91,10 L91,80 C91,80 90,81 90,82 L50,100 L10,82 C10,81 9,80 9,80 L9,10 L0,10 L0,0 L100,0 L100,10 L91,10 Z M18,10 L18,75 L50,90 L82,75 L82,10 L18,10 Z M32,29 L68,29 L68,38 L32,38 L32,29 Z M32,48 L68,48 L68,57 L32,57 L32,48 Z',
  inbox: 'M5,0 L95,0 C98,0 100,2 100,6 L100,94 C100,98 98,100 95,100 L5,100 C2,100 0,98 0,94 L0,6 C0,2 2,0 5,0 Z M90,61 L73,61 C67,75 53,82 40,75 C34,73 30,68 27,61 L10,61 L10,89 L90,89 L90,61 Z M90,50 L90,11 L10,11 L10,50 L35,50 C35,59 42,67 50,67 C58,67 65,59 65,50 L90,50 Z',
  pieChart: 'M42,0 C43,0 45,0 47,0 C76,0 100,24 100,53 C100,55 100,57 100,58 L93,58 C90,82 71,100 47,100 C21,100 0,79 0,53 C0,29 18,10 42,7 L42,0 Z M42,58 L42,17 C21,19 7,38 10,58 C12,77 28,91 47,91 C65,91 81,77 83,58 L42,58 Z M90,49 C88,28 72,12 51,10 L51,49 L90,49 Z',
  screenPlay: 'M55,75 L55,90 L80,90 L80,100 L20,100 L20,90 L45,90 L45,75 L10,75 C7,75 5,73 5,70 L5,10 L0,10 L0,0 L100,0 L100,10 L95,10 L95,70 C95,73 93,75 90,75 L55,75 Z M15,65 L85,65 L85,10 L15,10 L15,65 Z M40,20 L65,38 L40,55 L40,20 Z',
  screenChart: 'M55,90 L55,100 L45,100 L45,90 L5,90 C2,90 0,88 0,85 L0,15 L100,15 L100,85 C100,88 98,90 95,90 L55,90 Z M10,80 L90,80 L90,25 L10,25 L10,80 Z M55,35 L80,35 L80,45 L55,45 L55,35 Z M55,55 L80,55 L80,65 L55,65 L55,55 Z M35,35 L35,50 L50,50 C50,58 43,65 35,65 C27,65 20,58 20,50 C20,42 27,35 35,35 Z M0,0 L100,0 L100,10 L0,10 L0,0 Z',
  layers: 'M92,67 L99,70 C100,71 100,72 100,73 L99,74 L53,99 C51,100 49,100 47,99 L1,74 C0,74 0,72 0,71 C1,71 1,70 1,70 L8,67 L50,90 L92,67 Z M92,44 L99,48 C100,49 100,50 100,51 C99,51 99,52 99,52 L50,78 L1,52 C0,51 0,50 0,49 L1,48 L8,44 L50,67 L92,44 Z M53,1 L99,26 C100,26 100,28 100,29 C99,29 99,30 99,30 L50,56 L1,30 C0,29 0,28 0,27 L1,26 L47,1 C49,0 51,0 53,1 Z M50,10 L18,28 L50,45 L82,28 L50,10 Z',
  chartTile: 'M5,0 L95,0 C98,0 100,2 100,6 L100,94 C100,98 98,100 95,100 L5,100 C2,100 0,98 0,94 L0,6 C0,2 2,0 5,0 Z M10,11 L10,89 L90,89 L90,11 L10,11 Z M74,56 C72,71 58,80 45,77 C33,74 25,63 25,50 C25,37 33,25 45,23 L45,56 L74,56 Z M74,44 L55,44 L55,23 C65,25 72,34 74,44 Z',
  binoculars: 'M59,82 C51,61 53,29 62,11 C72,-6 87,-3 95,18 C103,39 101,71 92,89 C88,96 82,100 77,100 L23,100 C10,100 0,78 0,50 C0,22 10,0 23,0 C36,0 46,22 46,50 C46,62 44,73 41,82 L59,82 Z M23,82 C31,82 37,68 37,50 C37,32 31,18 23,18 C15,18 8,32 8,50 C8,68 15,82 23,82 Z M77,82 C85,82 92,68 92,50 C92,32 85,18 77,18 C69,18 63,32 63,50 C63,68 69,82 77,82 Z',
  cursorSend: 'M2,35 C-1,34 -1,33 2,32 L96,0 C99,-1 101,1 100,3 L73,98 C72,101 70,101 69,98 L47,53 L2,35 Z M26,34 L54,45 L69,75 L87,14 L26,34 Z',
  browser: 'M5,0 L95,0 C98,0 100,2 100,6 L100,94 C100,98 98,100 95,100 L5,100 C2,100 0,98 0,94 L0,6 C0,2 2,0 5,0 Z M90,44 L10,44 L10,89 L90,89 L90,44 Z M90,33 L90,11 L10,11 L10,33 L90,33 Z M65,17 L85,17 L85,28 L65,28 L65,17 Z',
  chatPair: 'M59,100 L46,83 L23,83 C20,83 18,81 18,78 L18,25 C18,22 20,20 23,20 L95,20 C98,20 100,22 100,25 L100,78 C100,81 98,83 95,83 L72,83 L59,100 Z M67,73 L91,73 L91,30 L27,30 L27,73 L51,73 L59,84 L67,73 Z M5,0 L82,0 L82,10 L9,10 L9,63 L0,63 L0,5 C0,2 2,0 5,0 Z',
  chatAlert: 'M22,82 L0,100 L0,5 C0,2 2,0 5,0 L95,0 C98,0 100,2 100,5 L100,77 C100,80 98,82 95,82 L22,82 Z M10,79 L19,72 L90,72 L90,10 L10,10 L10,79 Z M45,51 L55,51 L55,62 L45,62 L45,51 Z M45,21 L55,21 L55,46 L45,46 L45,21 Z',
  cropMark: 'M65,75 L65,85 L20,85 C17,85 15,83 15,80 L15,25 L0,25 L0,15 L15,15 L15,0 L25,0 L25,75 L65,75 Z M75,100 L75,25 L35,25 L35,15 L80,15 C83,15 85,17 85,20 L85,75 L100,75 L100,85 L85,85 L85,100 L75,100 Z',
  timer: 'M53,10 C31,8 12,25 10,47 C8,69 25,88 47,90 C69,92 88,75 90,53 C90,51 90,49 90,47 C90,42 88,38 87,34 L94,27 C98,34 100,42 100,50 C100,78 78,100 50,100 C22,100 0,78 0,50 C0,22 22,0 50,0 C58,0 66,2 73,6 L66,13 C62,12 58,10 53,10 Z M92,1 L99,8 L54,54 L46,54 L46,46 L92,1 Z',
  rulerPencil: 'M24,61 L13,72 L28,87 L87,28 L72,13 L61,24 L69,31 L61,39 L54,31 L46,39 L54,46 L46,54 L39,46 L31,54 L39,61 L31,69 L24,61 Z M76,2 L98,24 C101,26 101,29 98,31 L31,98 C29,101 26,101 24,98 L2,76 C-1,74 -1,71 2,69 L69,2 C71,-1 74,-1 76,2 Z M61,84 L69,76 L80,88 L88,88 L88,80 L76,69 L84,61 L97,75 L97,97 L75,97 L61,84 Z M16,39 L2,24 C-1,22 -1,19 2,16 L2,16 L16,2 C19,-1 22,-1 24,2 L39,16 L31,24 L20,13 L13,20 L24,31 L16,39 Z',
  cropFrame: 'M32,75 L65,75 L65,85 L20,85 C17,85 15,83 15,80 L15,25 L0,25 L0,15 L15,15 L15,0 L25,0 L25,68 L68,25 L35,25 L35,15 L78,15 L91,2 L98,9 L85,22 L85,75 L100,75 L100,85 L85,85 L85,100 L75,100 L75,32 L32,75 Z',
  moveArrows: 'M45,45 L45,19 L36,28 L29,21 L50,0 L71,21 L64,28 L55,19 L55,45 L81,45 L72,36 L79,29 L100,50 L79,71 L72,64 L81,55 L55,55 L55,81 L64,72 L71,79 L50,100 L29,79 L36,72 L45,81 L45,55 L19,55 L28,64 L21,71 L0,50 L21,29 L28,36 L19,45 Z',
  scissors: 'M50,60 L41,68 C48,78 44,91 34,97 C23,103 9,100 3,90 C-3,80 0,67 11,61 C18,57 26,57 33,61 L42,53 L19,31 L27,23 L50,45 L73,23 L81,31 L58,53 L66,61 C77,55 91,58 97,68 C103,78 100,91 89,97 C78,103 65,100 59,90 C55,83 55,75 59,68 L50,60 Z M22,89 C28,89 33,85 33,79 C33,73 28,68 22,68 C16,68 11,73 11,79 C11,85 16,89 22,89 Z M78,89 C84,89 89,85 89,79 C89,73 84,68 78,68 C72,68 67,73 67,79 C67,85 72,89 78,89 Z M89,53 L89,11 L11,11 L11,53 L0,53 L0,5 C0,2 2,0 6,0 L94,0 C98,0 100,2 100,5 L100,53 L89,53 Z',
  barTile: 'M10,28 L90,28 L90,11 L10,11 L10,28 Z M60,89 L60,39 L40,39 L40,89 L60,89 Z M70,89 L90,89 L90,39 L70,39 L70,89 Z M30,89 L30,39 L10,39 L10,89 L30,89 Z M5,0 L95,0 C98,0 100,2 100,6 L100,94 C100,98 98,100 95,100 L5,100 C2,100 0,98 0,94 L0,6 C0,2 2,0 5,0 Z',
  terminal: 'M5,0 L95,0 C98,0 100,2 100,6 L100,94 C100,98 98,100 95,100 L5,100 C2,100 0,98 0,94 L0,6 C0,2 2,0 5,0 Z M10,11 L10,89 L90,89 L90,11 L10,11 Z M50,67 L80,67 L80,78 L50,78 L50,67 Z M33,50 L19,34 L26,26 L47,50 L26,74 L19,66 L33,50 Z',
  photoTile: 'M52,59 L77,96 C78,97 78,99 76,99 C76,100 76,100 75,100 L25,100 C24,100 23,99 23,97 C23,97 23,96 23,96 L48,59 C49,57 50,57 52,58 C52,58 52,58 52,59 Z M50,74 L40,89 L60,89 L50,74 Z M80,89 L80,78 L90,78 L90,11 L10,11 L10,78 L20,78 L20,89 L5,89 C2,89 0,86 0,83 L0,83 L0,6 C0,2 2,0 5,0 L95,0 C98,0 100,2 100,6 L100,83 C100,86 98,89 95,89 L80,89 Z',
  castScreen: 'M5,0 L95,0 C98,0 100,2 100,6 L100,94 C100,98 98,100 95,100 L65,100 C65,96 65,93 64,89 L90,89 L90,11 L10,11 L10,29 C7,28 3,28 0,28 L0,6 C0,2 2,0 5,0 Z M55,100 L45,100 C45,72 25,50 0,50 L0,39 C30,39 55,66 55,100 Z M35,100 L25,100 C25,85 14,72 0,72 L0,61 C19,61 35,79 35,100 Z M15,100 L0,100 L0,83 C8,83 15,91 15,100 Z',
  docList: 'M45,89 L45,33 L10,33 L10,89 L45,89 Z M45,22 L45,6 C45,2 47,0 50,0 L95,0 C98,0 100,2 100,6 L100,94 C100,98 98,100 95,100 L5,100 C2,100 0,98 0,94 L0,28 C0,25 2,22 5,22 L45,22 Z M55,11 L55,89 L90,89 L90,11 L55,11 Z M15,72 L40,72 L40,83 L15,83 L15,72 Z M60,72 L85,72 L85,83 L60,83 L60,72 Z M60,56 L85,56 L85,67 L60,67 L60,56 Z M60,39 L85,39 L85,50 L60,50 L60,39 Z M15,56 L40,56 L40,67 L15,67 L15,56 Z',
  laptop: 'M14,11 L14,72 L86,72 L86,11 L14,11 Z M5,6 C5,3 7,0 9,0 L91,0 C93,0 95,2 95,6 L95,83 L5,83 L5,6 Z M0,89 L100,89 L100,100 L0,100 L0,89 Z',
  floppy: 'M83,89 L89,89 L89,21 L79,11 L72,11 L72,33 L22,33 L22,11 L11,11 L11,89 L17,89 L17,50 L83,50 L83,89 Z M6,0 L83,0 L98,15 C99,16 100,17 100,19 L100,94 C100,98 98,100 94,100 L6,100 C2,100 0,98 0,94 L0,6 C0,2 2,0 6,0 Z M28,61 L28,89 L72,89 L72,61 L28,61 Z',
  photoTile2: 'M52,59 L77,96 C78,97 78,99 76,99 C76,100 76,100 75,100 L25,100 C24,100 23,99 23,97 C23,97 23,96 23,96 L48,59 C49,57 50,57 52,58 C52,58 52,58 52,59 Z M80,89 L80,78 L90,78 L90,11 L10,11 L10,78 L20,78 L20,89 L5,89 C2,89 0,86 0,83 L0,83 L0,6 C0,2 2,0 5,0 L95,0 C98,0 100,2 100,6 L100,83 C100,86 98,89 95,89 L80,89 Z',
  keyboard: 'M10,11 L10,89 L90,89 L90,11 L10,11 Z M5,0 L95,0 C98,0 100,2 100,6 L100,94 C100,98 98,100 95,100 L5,100 C2,100 0,98 0,94 L0,6 C0,2 2,0 5,0 Z M20,22 L30,22 L30,33 L20,33 L20,22 Z M20,44 L30,44 L30,56 L20,56 L20,44 Z M20,67 L80,67 L80,78 L20,78 L20,67 Z M45,44 L55,44 L55,56 L45,56 L45,44 Z M45,22 L55,22 L55,33 L45,33 L45,22 Z M70,22 L80,22 L80,33 L70,33 L70,22 Z M70,44 L80,44 L80,56 L70,56 L70,44 Z',
  sheet: 'M12,10 L12,90 L88,90 L88,29 L64,10 L12,10 Z M6,0 L69,0 L98,24 C99,24 100,26 100,27 L100,95 C100,98 97,100 94,100 L6,100 C3,100 0,98 0,95 L0,5 C0,2 3,0 6,0 Z M56,40 L56,80 L44,80 L44,50 L25,50 L25,40 L56,40 Z M25,55 L38,55 L38,65 L25,65 L25,55 Z M62,55 L75,55 L75,65 L62,65 L62,55 Z M62,40 L75,40 L75,50 L62,50 L62,40 Z M25,70 L38,70 L38,80 L25,80 L25,70 Z M62,70 L75,70 L75,80 L62,80 L62,70 Z',
  floppy2: 'M83,89 L89,89 L89,21 L79,11 L72,11 L72,33 L22,33 L22,11 L11,11 L11,89 L17,89 L17,50 L83,50 L83,89 Z M6,0 L83,0 L98,15 C99,16 100,17 100,19 L100,94 C100,98 98,100 94,100 L6,100 C2,100 0,98 0,94 L0,6 C0,2 2,0 6,0 Z M28,61 L28,89 L72,89 L72,61 L28,61 Z',
  target: 'M0,55 L10,55 C13,77 33,92 55,90 C73,87 88,73 90,55 L100,55 C97,80 76,100 50,100 C24,100 3,80 0,55 Z M0,45 C3,20 24,0 50,0 C76,0 97,20 100,45 L90,45 C87,23 67,8 45,10 C27,13 12,27 10,45 L0,45 Z M50,60 C44,60 40,56 40,50 C40,44 44,40 50,40 C56,40 60,44 60,50 C60,56 56,60 50,60 Z',
  bookMark: 'M0,82 L0,15 C0,7 7,0 17,0 L94,0 C98,0 100,2 100,5 L100,95 C100,98 98,100 94,100 L19,100 C9,100 0,92 0,82 Z M89,90 L89,75 L19,75 C15,75 11,78 11,82 C11,87 15,90 19,90 L89,90 Z M39,10 L17,10 C14,10 11,12 11,15 L11,67 C14,66 17,65 19,65 L89,65 L89,10 L78,10 L78,50 L58,40 L39,50 L39,10 Z',
  contactBook: 'M0,0 L76,0 C81,0 86,4 86,10 L86,90 C86,96 81,100 76,100 L0,100 L0,0 Z M19,10 L10,10 L10,90 L19,90 L19,10 Z M29,90 L76,90 L76,10 L29,10 L29,90 Z M38,70 C38,62 44,55 52,55 C60,55 67,62 67,70 L38,70 Z M52,50 C47,50 43,46 43,40 C43,34 47,30 52,30 C58,30 62,34 62,40 C62,46 58,50 52,50 Z M90,20 L100,20 L100,40 L90,40 L90,20 Z M90,50 L100,50 L100,70 L90,70 L90,50 Z',
  copyPages: 'M22,20 L22,5 C22,2 25,0 28,0 L94,0 C98,0 100,2 100,5 L100,75 C100,78 98,80 94,80 L78,80 L78,95 C78,98 75,100 72,100 L6,100 C3,100 0,98 0,95 L0,95 L0,25 C0,22 3,20 6,20 L22,20 Z M11,30 L11,90 L67,90 L67,30 L11,30 Z M33,20 L78,20 L78,70 L89,70 L89,10 L33,10 L33,20 Z M22,45 L56,45 L56,55 L22,55 L22,45 Z M22,65 L56,65 L56,75 L22,75 L22,65 Z',
  clipboard: 'M78,9 L94,9 C98,9 100,11 100,14 L100,95 C100,98 98,100 94,100 L6,100 C2,100 0,98 0,95 L0,14 C0,11 2,9 6,9 L22,9 L22,0 L33,0 L33,9 L67,9 L67,0 L78,0 L78,9 Z M78,18 L78,27 L67,27 L67,18 L33,18 L33,27 L22,27 L22,18 L11,18 L11,91 L89,91 L89,18 L78,18 Z M22,36 L78,36 L78,45 L22,45 L22,36 Z M22,55 L78,55 L78,64 L22,64 L22,55 Z',
  book: 'M100,10 L22,10 C16,10 11,14 11,20 C11,26 16,30 22,30 L100,30 L100,95 C100,98 98,100 94,100 L22,100 C10,100 0,91 0,80 L0,20 C0,9 10,0 22,0 L94,0 C98,0 100,2 100,5 L100,10 Z M11,80 C11,86 16,90 22,90 L89,90 L89,40 L22,40 C18,40 14,39 11,37 L11,80 Z M94,25 L22,25 C19,25 17,23 17,20 C17,17 19,15 22,15 L94,15 L94,25 Z',
  shredder: 'M20,50 L80,50 L80,30 L60,30 L60,10 L20,10 L20,50 Z M10,50 L10,5 C10,2 12,0 15,0 L65,0 L90,25 L90,50 L100,50 L100,60 L0,60 L0,50 L10,50 Z M5,70 L15,70 L15,100 L5,100 L5,70 Z M85,70 L95,70 L95,100 L85,100 L85,70 Z M65,70 L75,70 L75,100 L65,100 L65,70 Z M45,70 L55,70 L55,100 L45,100 L45,70 Z M25,70 L35,70 L35,100 L25,100 L25,70 Z',
  folder: 'M52,11 L95,11 C98,11 100,14 100,17 L100,94 C100,98 98,100 95,100 L5,100 C2,100 0,98 0,94 L0,6 C0,2 2,0 5,0 L42,0 L52,11 Z M90,44 L10,44 L10,89 L90,89 L90,44 Z M90,33 L90,22 L48,22 L38,11 L10,11 L10,33 L90,33 Z',
  folderPie: 'M52,11 L95,11 C98,11 100,14 100,17 L100,94 C100,98 98,100 95,100 L5,100 C2,100 0,98 0,94 L0,6 C0,2 2,0 5,0 L42,0 L52,11 Z M10,11 L10,89 L90,89 L90,22 L48,22 L38,11 L10,11 Z M50,33 L50,56 L70,56 C70,68 61,78 50,78 C39,78 30,68 30,56 C30,43 39,33 50,33 Z',
};

// ---------------------------------------------------------------- background art
// Organic 'wash' shapes from the template masters, normalised to a 0..1 box.
const WASHES = {
  w1: 'M0.49,0.07 C0.35,0.13 0.32,-0.03 0.2,0.01 C0,0.07 0.15,0.38 0.1,0.52 C0.09,0.56 0.03,0.63 0.02,0.67 C-0.01,0.74 -0.03,0.91 0.12,0.98 C0.2,1.02 0.28,1.01 0.28,0.97 C0.28,0.94 0.24,0.92 0.27,0.88 C0.32,0.81 0.44,0.79 0.52,0.78 C0.62,0.78 0.88,0.85 0.97,0.64 C1.09,0.36 0.91,-0.09 0.49,0.07 Z',
  w2: 'M0.01,0.32 C-0.03,0.53 0.04,0.71 0.14,0.88 C0.17,0.95 0.2,1.01 0.28,1 C0.34,0.99 0.39,0.93 0.42,0.87 C0.45,0.8 0.46,0.75 0.5,0.69 C0.56,0.61 0.63,0.61 0.71,0.57 C0.82,0.53 0.97,0.48 1,0.34 C1.02,0.22 0.98,0.1 0.88,0.06 C0.7,-0.01 0.48,-0.02 0.31,0.03 C0.16,0.06 0,0.17 0.01,0.36',
};

// Per-layout decoration: soft blobs (ellipses) and hand-drawn doodles (rounded boxes).
function deco(s, items) {
  for (const d of items) {
    if (d[0] === 'blob') sh(s, 'ellipse', d[1], d[2], d[3], d[4], { fill: d[5] });
    else if (d[0] === 'wash') poly(s, WASHES[d[5]], d[1], d[2], d[3], d[4], { fill: d[6] });
    else sh(s, 'roundRect', d[1], d[2], d[3], d[4], { fill: d[5], rectRadius: Math.min(d[3], d[4]) / 2.5 });
  }
}
const DECO = {
  'Cover': [
    ['blob', 10.44, 4.74, 2.89, 2.76, 'F7F9FA'],
    ['blob', 11.61, 5.94, 1.61, 1.45, 'E8E8E8'],
    ['blob', 0.91, 1.9, 2.04, 1.48, 'FAF9F6'],
    ['blob', 0, 0, 2.29, 3.51, 'F8F8F8'],
    ['blob', 11.55, 0.5, 1.95, 1.42, 'F0F0F0'],
    ['blob', -1.12, 6.12, 2.48, 1.1, 'EFEFEF'],
    ['blob', 2.45, 0.41, 0.57, 0.6, 'C2C2C2'],
    ['blob', 1.55, 0.41, 0.64, 0.64, 'E8E8E8'],
    ['blob', 8.33, 6.48, 0.13, 0.36, 'C8C8C8'],
    ['doodle', 11.53, 1.08, 0.18, 0.14, '585858'],
  ],
  'Option 1': [
    ['blob', 2.11, 5.55, 1.76, 1.28, 'F3F1F1'],
    ['blob', 0, 4.88, 3.52, 2.62, 'F9F9F9'],
    ['blob', -0.74, 0.13, 1.23, 1.24, 'E2E2E2'],
    ['blob', 10.77, 0, 2.59, 4.09, 'FBFAF7'],
    ['blob', 9.62, -0.36, 1.61, 1.45, 'E8E8E8'],
    ['blob', 1.99, 6.21, 0.8, 0.92, 'C9C9C9'],
  ],
  'Option 20': [
    ['blob', 2.03, 5.56, 5.52, 1.94, 'F5F5F5'],
    ['blob', 8.97, 0, 4.36, 3.92, 'F5F7F8'],
    ['blob', 1.13, 0.3, 0.64, 0.64, 'E8E8E8'],
  ],
  'Slide 5': [
    ['blob', 10.44, 4.74, 2.89, 2.76, 'F7F9FA'],
    ['blob', 0.91, 1.9, 2.04, 1.48, 'FAF9F6'],
    ['blob', 0, 0, 2.29, 3.51, 'F8F8F8'],
    ['blob', 11.33, 6.24, 1.61, 1.45, 'E8E8E8'],
    ['blob', 11.63, -0.39, 1.95, 1.42, 'F0F0F0'],
    ['blob', 2.58, 0.39, 0.64, 0.64, 'E8E8E8'],
  ],
  'Option 6': [
    ['blob', -0.36, 6.36, 1.84, 0.88, 'EBEBEB'],
    ['blob', 11.51, 0.22, 1.5, 1.35, 'E8E8E8'],
    ['blob', 3.82, 1.5, 0.23, 0.35, 'D5D5D5'],
    ['wash', -1.41, 2.17, 6.77, 5.87, 'w1', { color: C.mist, transparency: 50 }],
    ['wash', 11.45, -0.29, 3.53, 3.02, 'w2', { color: 'F5F0E4', transparency: 35 }],
  ],
  'Option 8': [
    ['blob', 0, 3.44, 2.93, 4.36, 'F8F8F8'],
    ['blob', 8.31, 1.11, 4.88, 2.05, 'FAF9F6'],
    ['blob', 7.39, 0, 3.81, 1.61, 'F7F7F7'],
    ['blob', 0.83, 0.53, 1.27, 1.43, 'EFEFEF'],
    ['blob', 12.44, 5.95, 1.5, 1.35, 'E8E8E8'],
    ['wash', 9.43, 5.31, 3.53, 3.02, 'w2', { color: 'F5F0E4', transparency: 35 }],
  ],
  'Option 9': [
    ['blob', 2.03, 5.56, 5.52, 1.94, 'F5F5F5'],
    ['blob', 8.97, 0, 4.36, 3.92, 'F5F7F8'],
    ['blob', -0.1, 0.7, 0.64, 0.64, 'E8E8E8'],
  ],
  'Option 10': [
    ['blob', 2.11, 5.55, 1.76, 1.28, 'F3F1F1'],
    ['blob', 0, 4.88, 3.52, 2.62, 'F9F9F9'],
    ['blob', 10.77, 0, 2.59, 4.09, 'FBFAF7'],
    ['blob', 0.95, -0.41, 1.61, 1.45, 'E8E8E8'],
    ['blob', 11.85, 6.83, 1.23, 1.24, 'E2E2E2'],
    ['blob', 12.05, 0.31, 0.8, 0.92, 'C9C9C9'],
  ],
  'Option 11': [
    ['blob', -0.36, 6.36, 1.84, 0.88, 'EBEBEB'],
    ['blob', 9.93, 1.06, 2.5, 3.72, 'FAF9F9'],
    ['blob', 7.81, 0, 4.14, 1.38, 'F5F7F8'],
    ['blob', 12.27, 5.86, 1.5, 1.35, 'E8E8E8'],
    ['wash', -1.41, 2.17, 6.77, 5.87, 'w1', { color: C.mist, transparency: 50 }],
  ],
  'Option 12': [
    ['blob', 0, 2.2, 2.93, 4.36, 'F8F8F8'],
    ['blob', 7.08, 1.11, 4.88, 2.05, 'FAF9F6'],
    ['blob', 6.15, 0, 3.81, 1.61, 'F7F7F7'],
    ['blob', 6.68, 0.54, 0.96, 0.87, 'F6F6F6'],
    ['blob', 12.58, 6.17, 1.24, 1.11, 'E8E8E8'],
    ['blob', 0.33, 6.61, 1.04, 0.58, 'D8D8D8'],
    ['blob', 11.96, -0.45, 1.27, 1.43, 'EFEFEF'],
  ],
  'Option 14': [
    ['blob', 10.44, 4.74, 2.89, 2.76, 'F7F9FA'],
    ['blob', 9.64, 6.56, 1.61, 1.45, 'E8E8E8'],
    ['blob', 0.91, 1.9, 2.04, 1.48, 'FAF9F6'],
    ['blob', 0, 0, 2.29, 3.51, 'F8F8F8'],
    ['blob', 12.06, -0.33, 1.95, 1.42, 'F0F0F0'],
    ['blob', 0.52, 6.36, 2.48, 1.1, 'EFEFEF'],
    ['blob', 2.11, 0.33, 0.57, 0.6, 'C2C2C2'],
    ['blob', 1.44, 0.33, 0.64, 0.64, 'E8E8E8'],
    ['doodle', 12.05, 0.26, 0.18, 0.14, '585858'],
  ],
  'Option 15': [
    ['blob', 0, 4.05, 2.56, 3.8, 'F8F7F6'],
    ['blob', 1.45, 2, 2.9, 1.8, 'F9F8F8'],
    ['blob', 5.95, 0, 5.99, 2.11, 'F5F5F6'],
  ],
  'Option 16': [
    ['blob', 10.44, 4.74, 2.89, 2.76, 'F7F9FA'],
    ['blob', 0, 0, 2.9, 4.44, 'F8F8F8'],
    ['blob', 11.33, 6.24, 1.61, 1.45, 'E8E8E8'],
    ['blob', 6.3, -0.21, 1.95, 1.42, 'F0F0F0'],
  ],
  'Option 21': [
    ['blob', 7.17, 5.69, 5.52, 1.94, 'F5F5F5'],
    ['blob', 8.97, 0, 4.36, 3.92, 'F5F7F8'],
    ['blob', 1.61, -0.56, 1.5, 1.35, 'E8E8E8'],
    ['wash', -0.98, 0.13, 3.53, 3.02, 'w2', { color: 'F5F0E4', transparency: 35 }],
  ],
  'Option 22': [
    ['blob', 7.26, 3.78, 7.93, 4.91, 'FAF9F6'],
    ['blob', -0.02, -0.01, 3.73, 3.35, 'F3F4F6'],
    ['doodle', 3.63, 0.28, 1.55, 0.96, 'CDD8E0'],
    ['blob', 12.09, 2.52, 0.99, 0.82, 'E4E4E4'],
    ['blob', 0.71, 2.5, 0.22, 0.72, 'CDCDCD'],
  ],
  'Option 23': [
    ['blob', 3.31, 5.71, 5.52, 1.94, 'F5F5F5'],
    ['blob', 11.98, 1.18, 1.5, 1.35, 'E8E8E8'],
    ['blob', -0.01, 1.03, 1.62, 1.82, 'EFEFEF'],
    ['wash', 9.08, -0.85, 3.53, 3.02, 'w2', { color: 'F5F0E4', transparency: 35 }],
  ],
  'Option 24': [
    ['blob', 7.5, 5.69, 5.52, 1.94, 'F5F5F5'],
    ['blob', 8.97, 0, 4.36, 3.92, 'F5F7F8'],
    ['blob', 11.52, 5.98, 1.5, 1.35, 'E8E8E8'],
    ['wash', -0.73, -0.97, 3.53, 3.02, 'w2', { color: 'F5F0E4', transparency: 35 }],
  ],
  'Option 26': [
    ['blob', 10.44, 4.74, 2.89, 2.76, 'F7F9FA'],
    ['blob', 0.91, 1.9, 2.04, 1.48, 'FAF9F6'],
    ['blob', 0, 0, 2.29, 3.51, 'F8F8F8'],
    ['blob', 11.65, 6.48, 1.61, 1.45, 'E8E8E8'],
    ['blob', 11.32, 0, 1.95, 1.42, 'F0F0F0'],
    ['blob', 2.58, 0.39, 0.64, 0.64, 'E8E8E8'],
  ],
  'Option 27': [
    ['blob', -0.36, 6.36, 1.84, 0.88, 'EBEBEB'],
    ['blob', 11.51, 0.22, 1.5, 1.35, 'E8E8E8'],
    ['blob', 0.62, 0.92, 0.23, 0.35, 'D5D5D5'],
    ['wash', -1.41, 2.17, 6.77, 5.87, 'w1', { color: C.mist, transparency: 50 }],
    ['wash', 11.45, -0.29, 3.53, 3.02, 'w2', { color: 'F5F0E4', transparency: 35 }],
  ],
  'Option 28': [
    ['blob', 3.7, -0.01, 5.1, 1.7, 'F5F6F8'],
    ['blob', 11.64, 0, 1.83, 1.2, 'F7F8F9'],
    ['blob', 0, 3.62, 6.1, 3.92, 'F3F2F1'],
    ['blob', -0.43, 1.01, 1.5, 1.35, 'E8E8E8'],
  ],
  'Option 29': [
    ['blob', 0, 0, 1.66, 2.55, 'F4F2F2'],
    ['blob', 10.89, 5.16, 2.45, 2.34, 'F8F5F0'],
    ['blob', 0.69, 0.42, 1.33, 0.85, 'DBDBDB'],
    ['blob', 11.44, 6.38, 1.19, 0.3, 'D0D0D0'],
    ['blob', 10.94, 1.35, 0.81, 0.63, 'CECECE'],
  ],
  'Option 30': [
    ['blob', 0, 4.35, 4.96, 3.18, 'F8F8F7'],
    ['blob', 1.7, -0.01, 7.25, 2.42, 'F5F6F8'],
    ['blob', 2.91, 6.07, 1.59, 0.7, 'EFEFEF'],
    ['blob', 11.95, 4.05, 0.89, 0.91, 'DFDFDF'],
    ['blob', 5.96, 1.21, 0.55, 0.28, 'CFCFCF'],
  ],
  'Option 31': [
    ['blob', -0.06, 2.57, 2.93, 4.36, 'F8F8F8'],
    ['blob', 8.26, 1.11, 4.88, 2.05, 'FAF9F6'],
    ['blob', 7.34, 0, 3.81, 1.61, 'F7F7F7'],
    ['blob', 12.1, 5.47, 1.5, 1.35, 'E8E8E8'],
    ['blob', 0.69, 0.19, 1.27, 1.43, 'EFEFEF'],
    ['wash', 9.38, 5.31, 3.53, 3.02, 'w2', { color: 'F5F0E4', transparency: 35 }],
  ],
  'Option 18': [
    ['blob', 0, 4.24, 4.29, 3.26, 'FAF9F7'],
    ['blob', -0.21, 4.3, 1.24, 1.11, 'E8E8E8'],
    ['blob', 7.37, 7.15, 2.33, 1.11, 'EBEBEB'],
    ['blob', 9.74, 0.93, 1.27, 1.43, 'EFEFEF'],
    ['wash', 7.48, -1.41, 6.77, 5.87, 'w1', { color: C.mist, transparency: 59 }],
  ],
};

// ---------------------------------------------------------------- repeated body copy
const L1 = 'Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s';
const L2 = 'Lorem Ipsum is simply dummy text of the printing and typesetting industry. ';
const L3 = 'Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry\'s standard';
const L4 = 'Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has';
const L5 = 'Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry\'s standard dummy text ever since the ';
const L6 = 'Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s, when ';
const L7 = 'Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry\'s standard dummy text ever since the 1500s, when an';
const L8 = 'an and typesetting industry. Lorem Ipsum has been the industry\'s standard dummy';
const L9 = 'PLACEHOLDER';
const L10 = 'Lorem Ipsum is simply dummy text of the printing and';
const L11 = 'Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry\'s ';

// ---------------------------------------------------------------- slides
function slide01(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Cover']);
  title(s, 1.31, 2.79, 10.71, 1.92, { fontSize: 54, align: 'center', tail: 'Infographic Presentation' });
  pageNum(s, 1);
}

function slide02(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 1']);
  sh(s, 'ellipse', 5.47, 2.56, 2.4, 2.4, { fill: C.gold, shadow: SH.soft });
  sh(s, 'ellipse', 5.62, 2.72, 2.09, 2.09, { fill: C.sand, shadow: tint(C.navy) });
  sh(s, 'ellipse', 5.8, 2.89, 1.74, 1.74, { fill: C.mist, shadow: tint(C.navy) });
  sh(s, 'ellipse', 5.97, 3.07, 1.4, 1.4, { fill: C.white, shadow: SH.soft });
  sh(s, 'line', 7.51, 2.23, 0.61, 0.73, { line: { color: C.gold, width: 1, dashType: 'dash' }, flipH: true });
  sh(s, 'line', 8.12, 2.22, 0.78, 0, { line: { color: C.gold, width: 1, dashType: 'dash' }, flipH: true });
  sh(s, 'ellipse', 8.82, 2.12, 0.21, 0.21, { fill: C.gold, flipH: true });
  sh(s, 'line', 7.59, 3.89, 1.23, 0.02, { line: { color: C.sand, width: 1, dashType: 'dash' }, flipH: true });
  sh(s, 'ellipse', 8.82, 3.79, 0.21, 0.21, { fill: C.sand, flipH: true });
  sh(s, 'line', 5.51, 4.3, 0.59, 0.61, { line: { color: C.mist, width: 1, dashType: 'dash' }, flipV: true });
  sh(s, 'line', 4.6, 4.91, 0.91, 0, { line: { color: C.mist, width: 1, dashType: 'dash' }, flipV: true });
  sh(s, 'ellipse', 4.4, 4.84, 0.23, 0.24, { fill: C.mist, flipV: true });
  title(s, 0.88, 1.16, 4.34, 1.45);
  statCard(s, 0.88, 4.38, 3.16, { ring: C.mist, pct: '53%', head: 'Your Text Here', headColor: C.mist, body: 'Lorem ipsum dolor sit amet, ', bodyW: 1.73 });
  statCard(s, 9.29, 1.57, 3.16, { ring: C.gold, pct: '53%', head: 'Your Text Here', headColor: C.gold, body: 'Lorem ipsum dolor sit amet, ', bodyW: 1.73, bodyBold: true });
  statCard(s, 9.29, 3.33, 3.16, { ring: C.sand, pct: '53%', head: 'Your Text Here', headColor: C.sand, body: 'Lorem ipsum dolor sit amet, ', bodyW: 1.73, bodyBold: true });
  tx(s, L1, 0.85, 2.96, 4.45, 0.9, { lineSpacingMultiple: 1.5 });
  tx(s, L1, 8.01, 5.44, 4.45, 0.9, { lineSpacingMultiple: 1.5 });
  icon(s, 'briefcase', 6.46, 3.55, 0.42, 0.42, C.navy);
  pageNum(s, 2);
}

function slide03(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 20']);
  sh(s, 'pie', 5.55, 1.89, 3.79, 3.79, { fill: C.sand, angleRange: [221.83, 112.05] });
  sh(s, 'pie', 5.29, 1.62, 4.32, 4.32, { fill: C.mist, angleRange: [352.63, 78.31] });
  sh(s, 'pie', 5, 1.33, 4.9, 4.9, { fill: C.navy, angleRange: [77.99, 250.35] });
  sh(s, 'ellipse', 6.78, 3.12, 1.33, 1.33, { fill: C.white, shadow: SH.card });
  sh(s, 'roundRect', 4.72, 3.51, 1.27, 0.58, { fill: C.white, rectRadius: 0.29, shadow: SH.deep });
  sh(s, 'triangle', 5.25, 4.01, 0.23, 0.2, { fill: C.white, flipV: true, shadow: SH.card });
  tx(s, [{ text: '48', options: { fontSize: 20, fontFace: F.bold, bold: true, color: C.gold } }, { text: '%', options: { fontSize: 20, fontFace: F.bold, bold: true, color: C.gold, baseline: 40 } }], 4.88, 3.59, 0.94, 0.44, { fontSize: 20, fontFace: F.bold, bold: true, color: C.gold, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  sh(s, 'roundRect', 7.92, 2.18, 1.27, 0.58, { fill: C.white, rectRadius: 0.29, shadow: SH.deep });
  sh(s, 'triangle', 8.46, 2.68, 0.23, 0.2, { fill: C.white, flipV: true, shadow: SH.card });
  tx(s, [{ text: '27', options: { fontSize: 20, fontFace: F.bold, bold: true, color: C.sand } }, { text: '%', options: { fontSize: 20, fontFace: F.bold, bold: true, color: C.sand, baseline: 40 } }], 8.09, 2.24, 0.94, 0.44, { fontSize: 20, fontFace: F.bold, bold: true, color: C.sand, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  sh(s, 'roundRect', 8.32, 4.55, 1.27, 0.58, { fill: C.white, rectRadius: 0.29, shadow: SH.deep });
  sh(s, 'triangle', 8.86, 5.04, 0.23, 0.2, { fill: C.white, flipV: true, shadow: SH.card });
  tx(s, [{ text: '25', options: { fontSize: 20, fontFace: F.bold, bold: true, color: C.mist } }, { text: '%', options: { fontSize: 20, fontFace: F.bold, bold: true, color: C.mist, baseline: 40 } }], 8.48, 4.61, 0.94, 0.44, { fontSize: 20, fontFace: F.bold, bold: true, color: C.mist, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  sh(s, 'roundRect', 10.2, 2.32, 2.26, 1.28, { fill: { color: 'D9D9D9', transparency: 82 }, rectRadius: 0.16, shadow: SH.card });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur.', 10.39, 2.88, 2.07, 0.53, { lineSpacingMultiple: 1.2 });
  tx(s, 'Networking', 10.33, 2.4, 2.12, 0.56, { fontSize: 18, fontFace: F.bold, bold: true, color: C.sand, valign: 'middle', lineSpacingMultiple: 1.2 });
  sh(s, 'roundRect', 10.2, 0.86, 2.26, 1.28, { fill: { color: 'D9D9D9', transparency: 82 }, rectRadius: 0.18, shadow: SH.card });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur.', 10.39, 1.42, 2.07, 0.53, { lineSpacingMultiple: 1.2 });
  tx(s, 'Database', 10.33, 0.84, 1.78, 0.74, { fontSize: 18, fontFace: F.bold, bold: true, color: C.navy, valign: 'middle', lineSpacingMultiple: 1.2 });
  sh(s, 'roundRect', 10.19, 5.38, 2.26, 1.28, { fill: { color: 'D9D9D9', transparency: 82 }, rectRadius: 0.16, shadow: SH.card });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur.', 10.38, 5.95, 2.07, 0.53, { lineSpacingMultiple: 1.2 });
  tx(s, 'Programming', 10.33, 5.44, 2.12, 0.56, { fontSize: 18, fontFace: F.bold, bold: true, color: 'B2B2B2', valign: 'middle', lineSpacingMultiple: 1.2 });
  sh(s, 'roundRect', 10.2, 3.91, 2.26, 1.28, { fill: C.mist, rectRadius: 0.18, shadow: SH.deep });
  tx(s, 'Q1 Total Sales', 10.36, 4.18, 1.95, 0.33, { fontSize: 16, color: C.white, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  tx(s, '230K', 10.46, 4.56, 1.74, 0.33, { fontSize: 28, fontFace: F.bold, bold: true, color: C.white, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  title(s, 0.88, 1.58, 3.63, 1.45);
  tx(s, L7, 0.87, 3.39, 3.68, 1.18, { lineSpacingMultiple: 1.5 });
  tx(s, L7, 0.87, 4.73, 3.68, 1.18, { lineSpacingMultiple: 1.5 });
  icon(s, 'boxDown', 7.24, 3.61, 0.42, 0.38, C.navy);
  pageNum(s, 3);
}

function slide04(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Slide 5']);
  sh(s, 'pie', 8.51, 1.36, 2.98, 2.98, { fill: C.taupe, line: { color: C.white, width: 1 }, angleRange: [229.88, 89.55], flipH: true });
  sh(s, 'pie', 8.36, 1.21, 3.28, 3.28, { fill: 'F1F1F1', line: { color: C.white, width: 1 }, angleRange: [229.88, 29.56], flipH: true });
  sh(s, 'pie', 8.19, 1.04, 3.6, 3.6, { fill: C.sand, line: { color: C.white, width: 1 }, angleRange: [229.88, 359.98], flipH: true });
  sh(s, 'pie', 8.01, 0.86, 3.96, 3.96, { fill: C.navy, line: { color: C.white, width: 1 }, angleRange: [89.55, 270], flipH: true });
  tx(s, '10%', 8.47, 2.95, 0.73, 0.44, { fontSize: 20, fontFace: F.bold, bold: true, color: C.white, align: 'center', wrap: false });
  tx(s, '20%', 9.02, 3.46, 0.91, 0.51, { fontSize: 24, fontFace: F.bold, bold: true, color: C.white, align: 'center', wrap: false });
  tx(s, '25%', 8.65, 1.76, 1.14, 0.64, { fontSize: 32, fontFace: F.bold, bold: true, color: C.white, align: 'center', wrap: false });
  tx(s, '50%', 10.08, 3.44, 1.3, 0.71, { fontSize: 36, fontFace: F.bold, bold: true, color: C.white, align: 'center', wrap: false });
  sh(s, 'ellipse', 9.4, 2.26, 1.16, 1.16, { fill: C.white, shadow: SH.deep });
  sh(s, 'roundRect', 0.91, 5.15, 11.07, 1.36, { fill: C.white, rectRadius: 0.14, shadow: SH.lift });
  sh(s, 'rect', 1.41, 5.62, 0.2, 0.17, { fill: C.taupe });
  tx(s, 'Insert Data 1 Detail Here', 1.61, 5.47, 2.14, 0.73, { fontSize: 14, color: '595959', lineSpacingMultiple: 1.4 });
  legend(s, 4, 5.63, 2.73, [[C.sand, 'Insert Data 2 Detail Here'], [C.navy, 'Insert Data 3 Detail Here'], [C.mist, 'Insert Data 4 Detail Here']], { sw: 0.17, tw: 2.14, ty: -0.15 });
  title(s, 0.91, 1.43, 3.63, 1.45);
  tx(s, [{ text: L6, options: {  } }, { text: L8, options: {  } }], 0.91, 3.14, 6.87, 0.9, { lineSpacingMultiple: 1.5 });
  tx(s, [{ text: L6, options: {  } }, { text: 'an', options: {  } }], 0.91, 4.12, 6.87, 0.63, { lineSpacingMultiple: 1.5 });
  icon(s, 'archiveBox', 9.77, 2.69, 0.42, 0.38, C.navy);
  pageNum(s, 4);
}

function slide05(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 6']);
  tx(s, 'A wonderful serenity has taken possession of my entire soul.', 9.98, 5.78, 2.72, 0.59, { color: C.slate, lineSpacingMultiple: 1.4 });
  sh(s, 'pie', 4.47, 1.44, 4.47, 4.47, { fill: C.gold, angleRange: [174.25, 253.8] });
  sh(s, 'pie', 4.9, 1.6, 3.97, 3.97, { fill: C.mist, angleRange: [254.18, 291.25] });
  sh(s, 'pie', 4.61, 1.72, 4.28, 4.28, { fill: C.navy, angleRange: [91.75, 174.02] });
  sh(s, 'pie', 4.24, 0.69, 5.56, 6.12, { fill: C.taupe, angleRange: [292.17, 21.66] });
  sh(s, 'pie', 4.75, 1.51, 4.34, 4.78, { fill: C.sand, angleRange: [21.67, 91.5] });
  tx(s, '25%', 5.4, 4.34, 0.91, 0.51, { fontSize: 24, fontFace: F.bold, bold: true, color: C.white, align: 'center', wrap: false });
  tx(s, '50%', 7.87, 2.64, 1.3, 0.71, { fontSize: 36, fontFace: F.bold, bold: true, color: C.white, align: 'center', wrap: false });
  tx(s, '25%', 5.12, 2.65, 1.02, 0.57, { fontSize: 28, fontFace: F.bold, bold: true, color: C.white, align: 'center', wrap: false });
  tx(s, '50%', 7.25, 4.65, 1.06, 0.57, { fontSize: 28, fontFace: F.bold, bold: true, color: C.white, align: 'center', wrap: false });
  tx(s, '10%', 6.63, 1.9, 0.63, 0.37, { fontSize: 16, fontFace: F.bold, bold: true, color: C.white, align: 'center', wrap: false });
  sh(s, 'roundRect', 8.19, 1.26, 3.24, 0.94, { fill: C.white, rectRadius: 0.47, shadow: SH.card });
  sh(s, 'triangle', 9.05, 2.06, 0.37, 0.32, { fill: C.white, flipV: true, shadow: SH.card });
  tx(s, 'Highest Demand', 7.77, 1.46, 4.08, 0.54, { fontSize: 22, fontFace: F.bold, bold: true, color: C.taupe, baseline: 40, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  tx(s, 'Chart Data Slide', 9.98, 5.2, 2.31, 0.4, { fontSize: 18, fontFace: F.bold, bold: true, color: C.dark });
  tx(s, '112M', 9.98, 4.33, 2.15, 0.84, { fontSize: 44, fontFace: F.bold, bold: true, color: C.gold });
  title(s, 0.91, 1.56, 3.63, 1.45);
  tx(s, L1, 0.91, 3.27, 3.4, 1.18, { lineSpacingMultiple: 1.5 });
  tx(s, L1, 0.91, 4.5, 3.4, 1.18, { lineSpacingMultiple: 1.5 });
  pageNum(s, 5);
}

function slide06(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 8']);
  poly(s, 'M1,1 C0.99,0.44 0.55,0 0,0 L0,1 L1,1', 1.07, 3.74, 2.29, 2.24, { fill: C.sand, line: { color: C.white, width: 3.5 }, rotate: 180 });
  poly(s, 'M1,0.01 C1,0.01 1,0 1,0 L0,0 L0,1 L0.01,1 C0.55,1 0.99,0.56 1,0.03 L1,0.01', 1.43, 1.77, 1.93, 1.96, { fill: C.mist, line: { color: C.white, width: 3.5 }, rotate: 180 });
  poly(s, 'M0,1 C0.01,0.44 0.45,0 1,0 L1,1 L0,1', 3.36, 3.65, 2.06, 2.11, { fill: C.gold, line: { color: C.white, width: 3.5 }, rotate: 180 });
  poly(s, 'M0,0.01 C0,0.01 0,0 0,0 L1,0 L1,1 L0.99,1 C0.45,1 0.01,0.56 0,0.03 L0,0.01', 3.35, 1.32, 2.31, 2.35, { fill: C.navy, line: { color: C.white, width: 3.5 }, rotate: 180 });
  poly(s, 'M0.93,0.34 C0.87,0.13 0.68,0 0.49,0.01 C0.45,0.01 0.41,0.01 0.37,0.03 C0.13,0.1 0,0.36 0.07,0.61 C0.07,0.62 0.07,0.62 0.07,0.62 C0.08,0.63 0.08,0.64 0.08,0.64 C0.08,0.64 0.08,0.65 0.08,0.65 L0.08,0.65 C0.17,0.88 0.4,1 0.63,0.93 C0.87,0.85 1,0.59 0.93,0.34 Z', 2.61, 2.98, 1.5, 1.42, { fill: C.white, rotate: 180 });
  title(s, 6.39, 1.05, 3.63, 1.45);
  sh(s, 'roundRect', 6.46, 2.9, 2.71, 1.64, { fill: { color: 'D9D9D9', transparency: 82 }, rectRadius: 0.23, shadow: SH.card });
  tx(s, '10%', 6.66, 3.09, 1.03, 0.57, { fontSize: 28, fontFace: F.bold, bold: true, color: 'CA3AB0' });
  tx(s, 'Your text here', 7.51, 3.27, 1.69, 0.34, { fontSize: 14, fontFace: F.bold, color: 'CA3AB0' });
  tx(s, 'Lorem ipsum dolor sit amet, ipsum consectetuer', 6.66, 3.65, 2.54, 0.56, { fontFace: F.bold, lineSpacingMultiple: 1.3 });
  sh(s, 'roundRect', 9.52, 2.9, 2.71, 1.64, { fill: C.navy, rectRadius: 0.23, shadow: SH.card });
  tx(s, '30%', 9.72, 3.09, 1.03, 0.57, { fontSize: 28, fontFace: F.bold, bold: true, color: C.white });
  tx(s, 'Your text here', 10.57, 3.27, 1.69, 0.34, { fontSize: 14, fontFace: F.bold, color: C.white });
  tx(s, 'Lorem ipsum dolor sit amet, ipsum consectetuer', 9.72, 3.65, 2.54, 0.56, { fontFace: F.bold, color: C.white, lineSpacingMultiple: 1.3 });
  sh(s, 'roundRect', 6.46, 4.82, 2.71, 1.64, { fill: C.sand, rectRadius: 0.23, shadow: SH.card });
  tx(s, '35%', 6.66, 5, 1.03, 0.57, { fontSize: 28, fontFace: F.bold, bold: true, color: C.white });
  tx(s, 'Your text here', 7.51, 5.19, 1.69, 0.34, { fontSize: 14, fontFace: F.bold, color: C.white });
  tx(s, 'Lorem ipsum dolor sit amet, ipsum consectetuer', 6.66, 5.56, 2.54, 0.56, { fontFace: F.bold, color: C.white, lineSpacingMultiple: 1.3 });
  sh(s, 'roundRect', 9.52, 4.82, 2.71, 1.64, { fill: { color: 'D9D9D9', transparency: 82 }, rectRadius: 0.23, shadow: SH.card });
  tx(s, '25%', 9.72, 5, 1.03, 0.57, { fontSize: 28, fontFace: F.bold, bold: true, color: C.gold });
  tx(s, 'Your text here', 10.57, 5.19, 1.78, 0.34, { fontSize: 14, fontFace: F.bold, color: C.gold });
  tx(s, 'Lorem ipsum dolor sit amet, ipsum consectetuer', 9.72, 5.56, 2.54, 0.56, { fontFace: F.bold, lineSpacingMultiple: 1.3 });
  icon(s, 'drawer', 2.24, 2.75, 0.38, 0.42, C.white);
  icon(s, 'globe', 4.04, 4.4, 0.42, 0.42, C.white);
  icon(s, 'badge', 2.14, 4.53, 0.46, 0.44, C.white);
  icon(s, 'inbox', 4.18, 2.38, 0.42, 0.38, C.white);
  icon(s, 'pieChart', 3.14, 3.51, 0.45, 0.45, '000000');
  pageNum(s, 6);
}

function slide07(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 9']);
  isoChart(s, 4.53, 3.51, 4.47, 2.53, ['1F2A34', '516F83', '776846', 'C9AF78', 'EBEBEB', 'EBEBEB', '527085', '1E2933']);
  sh(s, 'roundRect', 6.66, 3.09, 2.02, 0.58, { fill: C.white, rectRadius: 0.29, shadow: SH.deep });
  sh(s, 'triangle', 7.19, 3.58, 0.23, 0.2, { fill: C.white, flipV: true, shadow: SH.card });
  tx(s, 'Highest Peak', 6.39, 3.21, 2.53, 0.34, { fontSize: 22, fontFace: F.bold, bold: true, color: C.mist, baseline: 40, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  title(s, 3.2, 0.93, 6.92, 0.77, { align: 'center' });
  tx(s, [{ text: L6, options: {  } }, { text: L8, options: {  } }], 1.31, 1.96, 10.69, 0.63, { align: 'center', lineSpacingMultiple: 1.5 });
  sh(s, 'bentConnector3', 3.72, 3.68, 1.39, 0.32, { line: { color: C.mist, width: 1, dashType: 'dash' } });
  sh(s, 'bentConnector3', 8.39, 4.89, 1.22, 0.57);
  sh(s, 'bentConnector3', 3.72, 5.16, 1.52, 0.68, { line: { color: C.gold, width: 1, dashType: 'dash' }, rotate: 180, flipV: true });
  sh(s, 'roundRect', 0.78, 3.06, 2.71, 1.08, { fill: C.white, rectRadius: 0.15, shadow: SH.soft });
  tx(s, '55%', 0.96, 3.19, 1.07, 0.51, { fontSize: 24, fontFace: F.bold, bold: true, color: 'CA3AB0' });
  tx(s, 'Title Here', 1.88, 3.29, 1.42, 0.37, { fontSize: 16, fontFace: F.bold, color: 'CA3AB0' });
  tx(s, 'Lorem ipsum dolor sit amet.', 0.96, 3.64, 2.58, 0.32, { lineSpacingMultiple: 1.3 });
  sh(s, 'roundRect', 0.78, 5.3, 2.71, 1.08, { fill: C.white, rectRadius: 0.15, shadow: SH.soft });
  tx(s, '18%', 0.96, 5.43, 1.07, 0.51, { fontSize: 24, fontFace: F.bold, bold: true, color: C.gold });
  tx(s, 'Title Here', 1.88, 5.53, 1.42, 0.37, { fontSize: 16, fontFace: F.bold, color: C.gold });
  tx(s, 'Lorem ipsum dolor sit amet.', 0.96, 5.88, 2.58, 0.32, { lineSpacingMultiple: 1.3 });
  sh(s, 'roundRect', 9.84, 4.92, 2.71, 1.08, { fill: C.white, rectRadius: 0.15, shadow: SH.soft });
  tx(s, '27%', 10.03, 5.05, 1.07, 0.51, { fontSize: 24, fontFace: F.bold, bold: true, color: C.navy });
  tx(s, 'Title Here', 10.94, 5.15, 1.42, 0.37, { fontSize: 16, fontFace: F.bold, color: C.navy });
  tx(s, 'Lorem ipsum dolor sit amet.', 10.02, 5.5, 2.58, 0.32, { lineSpacingMultiple: 1.3 });
  icon(s, 'screenPlay', 7.61, 4.6, 0.31, 0.31, C.white);
  icon(s, 'screenChart', 6.44, 3.76, 0.31, 0.31, C.white);
  icon(s, 'layers', 5.25, 4.53, 0.29, 0.32, C.white);
  pageNum(s, 7);
}

function slide08(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 10']);
  sh(s, 'bentConnector3', 7.05, 1.4, 2.4, 0.59, { line: { color: C.gold, width: 1, dashType: 'dash' } });
  sh(s, 'bentConnector3', 8.26, 3.29, 1.2, 0.86);
  sh(s, 'bentConnector3', 7.39, 5.71, 2.07, 0.7, { line: { color: C.mist, width: 1, dashType: 'dash' } });
  tx(s, 'Networking', 9.73, 1.44, 2.21, 0.4, { fontSize: 18, fontFace: F.bold, bold: true, color: C.gold });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.', 9.73, 1.88, 2.32, 0.79, { fontFace: F.bold, color: '595959', lineSpacingMultiple: 1.3 });
  tx(s, 'Database', 9.74, 5.62, 2.21, 0.4, { fontSize: 18, fontFace: F.bold, bold: true, color: C.mist });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.', 9.74, 6.06, 2.32, 0.79, { fontFace: F.bold, color: '595959', lineSpacingMultiple: 1.3 });
  tx(s, 'Programming', 9.74, 3.51, 2.21, 0.4, { fontSize: 18, fontFace: F.bold, bold: true, color: C.navy });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.', 9.74, 3.94, 2.32, 0.79, { fontFace: F.bold, color: '595959', lineSpacingMultiple: 1.3 });
  title(s, 0.91, 1.74, 3.63, 1.45);
  tx(s, [{ text: L6, options: {  } }, { text: 'an and typesetting industry. ', options: {  } }], 0.91, 3.53, 4.17, 1.18, { lineSpacingMultiple: 1.5 });
  tx(s, 'Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry\'s standard dummy text ever', 0.91, 4.86, 4.17, 0.9, { lineSpacingMultiple: 1.5 });
  sh(s, 'pie', 2.87, 0.81, 5.59, 5.59, { fill: C.gold, angleRange: [180.04, 229.72], rotate: 90, shadow: SH.lift });
  sh(s, 'pie', 2.88, 0.92, 5.74, 5.67, { fill: C.navy, angleRange: [229.61, 299.1], rotate: 90, shadow: SH.lift });
  sh(s, 'pie', 2.89, 1.09, 5.59, 5.59, { fill: C.mist, angleRange: [299.01, 359.83], rotate: 90, shadow: SH.lift });
  tx(s, [{ text: '25', options: { fontSize: 24, fontFace: F.bold, bold: true, color: C.white } }, { text: '%', options: { fontSize: 24, fontFace: F.bold, bold: true, color: C.white, baseline: 40 } }], 6.74, 3.44, 1.3, 0.62, { fontSize: 24, fontFace: F.bold, bold: true, color: C.white, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  tx(s, [{ text: '35', options: { fontSize: 24, fontFace: F.bold, bold: true, color: C.white } }, { text: '%', options: { fontSize: 24, fontFace: F.bold, bold: true, color: C.white, baseline: 40 } }], 5.75, 1.85, 1.3, 0.62, { fontSize: 24, fontFace: F.bold, bold: true, color: C.white, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  tx(s, [{ text: '45', options: { fontSize: 24, fontFace: F.bold, bold: true, color: C.white } }, { text: '%', options: { fontSize: 24, fontFace: F.bold, bold: true, color: C.white, baseline: 40 } }], 5.8, 5.34, 1.3, 0.62, { fontSize: 24, fontFace: F.bold, bold: true, color: C.white, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  icon(s, 'chartTile', 6.09, 1.49, 0.42, 0.38, C.white);
  icon(s, 'binoculars', 7.1, 3.29, 0.5, 0.23, C.white);
  icon(s, 'cursorSend', 6.14, 4.97, 0.42, 0.42, C.white);
  pageNum(s, 8);
}

function slide09(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 11']);
  sh(s, 'arc', 8.3, 2.08, 3.96, 3.96, { line: { color: 'D0CECE', width: 25, transparency: 58 }, angleRange: [352.15, 90.3] });
  sh(s, 'arc', 8.3, 2.08, 3.96, 3.96, { line: { color: C.taupe, width: 25 }, angleRange: [288.62, 22.43] });
  sh(s, 'arc', 8.3, 2.08, 3.96, 3.96, { line: { color: C.sand, width: 25 }, angleRange: [142.2, 306.8] });
  sh(s, 'ellipse', 8.78, 2.58, 3.02, 3.02, { fill: C.white, shadow: SH.card });
  sh(s, 'arc', 8.3, 2.08, 3.96, 3.96, { line: { color: C.navy, width: 25 }, angleRange: [85.84, 154.19] });
  sh(s, 'roundRect', 7.45, 1.51, 2.48, 0.72, { fill: C.white, rectRadius: 0.36, shadow: SH.deep });
  sh(s, 'triangle', 9.17, 2.13, 0.28, 0.25, { fill: C.white, flipV: true, shadow: SH.card });
  tx(s, 'Highest Sales', 7.12, 1.67, 3.11, 0.41, { fontSize: 22, fontFace: F.bold, bold: true, color: 'F23D92', baseline: 40, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  sh(s, 'ellipse', 0.94, 2.35, 1.14, 1.14, { fill: C.white, shadow: SH.soft });
  tx(s, [{ text: '15', options: { fontSize: 28, fontFace: F.bold, bold: true, color: C.navy } }, { text: '%', options: { fontSize: 28, fontFace: F.bold, bold: true, color: C.navy, baseline: 40 } }], 1.07, 2.68, 0.95, 0.46, { fontSize: 28, fontFace: F.bold, bold: true, color: C.navy, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.', 2.42, 2.78, 3.79, 0.53, { lineSpacingMultiple: 1.2 });
  tx(s, 'Your Text Here', 2.42, 2.43, 3.23, 0.37, { fontSize: 16, fontFace: F.bold, bold: true, color: C.navy });
  sh(s, 'ellipse', 0.94, 3.79, 1.14, 1.14, { fill: C.white, shadow: SH.soft });
  tx(s, [{ text: '15', options: { fontSize: 28, fontFace: F.bold, bold: true, color: C.sand } }, { text: '%', options: { fontSize: 28, fontFace: F.bold, bold: true, color: C.sand, baseline: 40 } }], 1.07, 4.12, 0.95, 0.46, { fontSize: 28, fontFace: F.bold, bold: true, color: C.sand, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.', 2.42, 4.22, 3.79, 0.53, { lineSpacingMultiple: 1.2 });
  tx(s, 'Your Text Here', 2.42, 3.87, 3.23, 0.37, { fontSize: 16, fontFace: F.bold, bold: true, color: C.sand });
  sh(s, 'ellipse', 0.94, 5.28, 1.14, 1.14, { fill: C.white, shadow: SH.soft });
  tx(s, [{ text: '15', options: { fontSize: 28, fontFace: F.bold, bold: true, color: C.taupe } }, { text: '%', options: { fontSize: 28, fontFace: F.bold, bold: true, color: C.taupe, baseline: 40 } }], 1.07, 5.6, 0.95, 0.46, { fontSize: 28, fontFace: F.bold, bold: true, color: C.taupe, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.', 2.42, 5.71, 3.79, 0.53, { lineSpacingMultiple: 1.2 });
  tx(s, 'Your Text Here', 2.42, 5.36, 3.23, 0.37, { fontSize: 16, fontFace: F.bold, bold: true, color: C.taupe });
  title(s, 0.99, 0.62, 6.26, 1.45);
  icon(s, 'browser', 9.9, 3.72, 0.78, 0.7, C.gold);
  pageNum(s, 9);
}

function slide10(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 12']);
  sh(s, 'roundRect', 6.94, 1.51, 2.46, 2.3, { fill: C.white, rectRadius: 0.23, shadow: SH.card });
  tx(s, '75%', 7.77, 2.24, 0.78, 0.43, { fontSize: 16, fontFace: F.bold, bold: true, color: C.ink, align: 'center', valign: 'middle', lineSpacingMultiple: 1.3 });
  tx(s, 'Performance One', 7.12, 3.2, 2.09, 0.38, { fontSize: 14, color: '595959', align: 'center', lineSpacingMultiple: 1.3 });
  sh(s, 'roundRect', 9.68, 1.51, 2.46, 2.3, { fill: C.white, rectRadius: 0.2, shadow: SH.card });
  tx(s, '50%', 10.52, 2.24, 0.78, 0.43, { fontSize: 16, fontFace: F.bold, bold: true, color: C.sand, align: 'center', valign: 'middle', lineSpacingMultiple: 1.3 });
  tx(s, 'Performance Two', 9.87, 3.2, 2.09, 0.38, { fontSize: 14, color: '595959', align: 'center', lineSpacingMultiple: 1.3 });
  sh(s, 'roundRect', 0.98, 1.32, 5.06, 2.83, { fill: C.white, rectRadius: 0.06, shadow: SH.card });
  tx(s, '2023 Sales Record', 1.19, 1.54, 3.83, 0.39, { fontSize: 14, fontFace: F.bold, bold: true, color: '262626', lineSpacingMultiple: 1.3 });
  tx(s, '75%', 2.06, 2.46, 0.78, 0.43, { fontSize: 16, fontFace: F.bold, bold: true, color: C.white, align: 'center', valign: 'middle', lineSpacingMultiple: 1.3 });
  tx(s, '50%', 3.37, 2.52, 0.78, 0.43, { fontSize: 16, fontFace: F.bold, bold: true, color: C.white, align: 'center', valign: 'middle', lineSpacingMultiple: 1.3 });
  title(s, 0.98, 4.68, 3.63, 1.45);
  tx(s, L4, 6.94, 5, 2.46, 1.18, { lineSpacingMultiple: 1.5 });
  tx(s, 'Your Text Here', 6.94, 4.66, 2.27, 0.37, { fontSize: 16, fontFace: F.bold, bold: true, color: C.navy });
  tx(s, L4, 9.68, 5, 2.46, 1.18, { lineSpacingMultiple: 1.5 });
  tx(s, 'Your Text Here', 9.68, 4.66, 2.27, 0.37, { fontSize: 16, fontFace: F.bold, bold: true, color: C.sand });
  sh(s, 'blockArc', 7.58, 1.86, 1.17, 1.18, { fill: 'DBE3E9', angleRange: [278.62, 36.16], arcThicknessRatio: 0.34 });
  sh(s, 'blockArc', 7.58, 1.86, 1.17, 1.18, { fill: C.navy, angleRange: [37.63, 276.86], arcThicknessRatio: 0.35 });
  sh(s, 'blockArc', 10.32, 1.86, 1.17, 1.18, { fill: 'F9F6F1', angleRange: [278.62, 36.16], arcThicknessRatio: 0.34 });
  sh(s, 'blockArc', 10.32, 1.86, 1.17, 1.18, { fill: C.sand, line: { color: C.white, width: 1 }, angleRange: [37.63, 276.86], arcThicknessRatio: 0.35 });
  isoChart(s, 1.56, 2.12, 2.8, 1.58, ['445C6E', '527085', '5A544A', 'DDD1B9', '527085', '527085', '527085', '445C6E'], true);
  tx(s, '1 st Half Year', 5.05, 1.94, 0.9, 0.62, { lineSpacingMultiple: 1.5 });
  sh(s, 'rect', 4.74, 2.12, 0.1, 0.1, { fill: C.navy });
  tx(s, '2 st Half Year', 5.05, 2.87, 0.9, 0.62, { lineSpacingMultiple: 1.5 });
  sh(s, 'rect', 4.74, 3.05, 0.1, 0.1, { fill: 'DDD1B9' });
  pageNum(s, 10);
}

function slide11(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 14']);
  sh(s, 'ellipse', 5.61, 1.94, 3.63, 3.63, { line: { color: C.haze, width: 25 } });
  sh(s, 'pie', 5.15, 1.47, 4.55, 4.55, { fill: C.mist, angleRange: [55.22, 194.34], rotate: 180 });
  sh(s, 'pie', 5.15, 1.47, 4.55, 4.55, { fill: C.gold, angleRange: [194.19, 328.2], rotate: 180 });
  sh(s, 'ellipse', 6.71, 3.04, 1.42, 1.42, { fill: C.white });
  imgBox(s, 7.04, 3.37, 0.76, 0.76);
  sh(s, 'roundRect', 7.79, 1.94, 1.35, 0.62, { fill: C.white, rectRadius: 0.31, shadow: SH.deep });
  sh(s, 'triangle', 8.36, 2.47, 0.25, 0.21, { fill: C.white, flipV: true, shadow: SH.card });
  tx(s, [{ text: '48', options: { fontSize: 20, fontFace: F.bold, bold: true, color: 'CA3AB0' } }, { text: '%', options: { fontSize: 20, fontFace: F.bold, bold: true, color: 'CA3AB0', baseline: 40 } }], 7.79, 2.02, 1.35, 0.47, { fontSize: 20, fontFace: F.bold, bold: true, color: 'CA3AB0', align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  sh(s, 'roundRect', 6.45, 4.84, 1.35, 0.62, { fill: C.white, rectRadius: 0.31, shadow: SH.deep });
  sh(s, 'triangle', 7.02, 5.37, 0.25, 0.21, { fill: C.white, flipV: true, shadow: SH.card });
  tx(s, [{ text: '34', options: { fontSize: 20, fontFace: F.bold, bold: true, color: 'F71746' } }, { text: '%', options: { fontSize: 20, fontFace: F.bold, bold: true, color: 'F71746', baseline: 40 } }], 6.45, 4.92, 1.35, 0.47, { fontSize: 20, fontFace: F.bold, bold: true, color: 'F71746', align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  sh(s, 'roundRect', 10.1, 1.21, 2.26, 2.13, { fill: C.mist, rectRadius: 0.29, shadow: SH.card });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur', 10.36, 2.55, 1.87, 0.53, { fontFace: F.bold, color: C.white, lineSpacingMultiple: 1.2 });
  tx(s, '2022', 10.36, 1.38, 1.62, 0.64, { fontSize: 32, fontFace: F.bold, bold: true, color: C.white });
  tx(s, 'Your Text Here', 10.36, 1.99, 2.1, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: C.white });
  sh(s, 'roundRect', 10.1, 4.16, 2.26, 2.13, { fill: { color: 'D9D9D9', transparency: 82 }, rectRadius: 0.29, shadow: SH.card });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur', 10.36, 5.49, 1.87, 0.53, { fontFace: F.bold, color: C.slate, lineSpacingMultiple: 1.2 });
  tx(s, '2023', 10.36, 4.33, 1.62, 0.64, { fontSize: 32, fontFace: F.bold, bold: true, color: C.gold });
  tx(s, 'Your Text Here', 10.36, 4.93, 2.1, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: C.ink });
  title(s, 0.91, 1.68, 3.63, 1.45);
  tx(s, L1, 0.91, 3.39, 3.63, 1.18, { lineSpacingMultiple: 1.5 });
  tx(s, L1, 0.91, 4.63, 3.63, 1.18, { lineSpacingMultiple: 1.5 });
  icon(s, 'chatPair', 8.06, 4.77, 0.46, 0.43, C.white);
  icon(s, 'chatAlert', 8.46, 3.04, 0.42, 0.41, C.white);
  pageNum(s, 11);
}

function slide12(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 15']);
  sh(s, 'rect', 1.76, 2.88, 0.8, 0.04, { fill: C.white });
  sh(s, 'roundRect', 1.27, 0.56, 3.74, 3.94, { fill: C.white, rectRadius: 0.12, rotate: 270, flipH: true, shadow: SH.soft });
  tx(s, '120K', 1.31, 3.62, 1.69, 0.44, { fontSize: 20, fontFace: F.bold, bold: true, color: C.navy });
  tx(s, 'Total Data', 1.31, 3.31, 1.69, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: '595959' });
  tx(s, 'Database', 1.69, 0.85, 2.94, 0.4, { fontSize: 18, fontFace: F.bold, bold: true, color: C.navy, align: 'center' });
  sh(s, 'arc', 1.9, 1.54, 2.52, 2.52, { line: { color: C.navy, width: 17 }, angleRange: [180.92, 66.85] });
  sh(s, 'arc', 2.13, 1.77, 2.06, 2.06, { line: { color: C.gold, width: 17 }, angleRange: [180.57, 24.24] });
  sh(s, 'arc', 2.36, 1.99, 1.61, 1.61, { line: { color: C.sand, width: 17 }, angleRange: [180.7, 82.78] });
  sh(s, 'teardrop', 1.07, 1.79, 0.79, 0.79, { fill: C.navy, line: { color: C.white, width: 4 }, rotate: 135, flipH: true, shadow: SH.deep });
  tx(s, [{ text: '65', options: { fontSize: 20, fontFace: F.bold, bold: true, color: C.white } }, { text: '%', options: { fontSize: 20, fontFace: F.bold, bold: true, color: C.white, baseline: 40 } }], 1.03, 2.01, 0.94, 0.44, { fontSize: 20, fontFace: F.bold, bold: true, color: C.white, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  sh(s, 'roundRect', 2.71, 1.66, 0.53, 0.24, { fill: C.white, rectRadius: 0.12, shadow: SH.soft });
  sh(s, 'triangle', 2.94, 1.9, 0.06, 0.09, { fill: C.white, flipV: true, shadow: SH.card });
  tx(s, [{ text: '67', options: { fontSize: 9, bold: true, color: 'F23D92' } }, { text: '%', options: { fontSize: 9, bold: true, color: 'F23D92', baseline: 40 } }], 2.66, 1.69, 0.61, 0.28, { fontSize: 9, bold: true, color: 'F23D92', align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  sh(s, 'roundRect', 3.92, 2.66, 0.53, 0.24, { fill: C.white, rectRadius: 0.12, shadow: SH.soft });
  sh(s, 'triangle', 4.15, 2.91, 0.06, 0.09, { fill: C.white, flipV: true, shadow: SH.card });
  tx(s, [{ text: '55', options: { fontSize: 9, bold: true, color: 'F71746' } }, { text: '%', options: { fontSize: 9, bold: true, color: 'F71746', baseline: 40 } }], 3.87, 2.69, 0.61, 0.28, { fontSize: 9, bold: true, color: 'F71746', align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  sh(s, 'ellipse', 2.7, 2.28, 0.98, 0.98, { fill: C.white, flipV: true, shadow: SH.soft });
  sh(s, 'roundRect', 8.53, 2.78, 3.74, 3.94, { fill: C.white, rectRadius: 0.12, rotate: 270, flipH: true, shadow: SH.soft });
  sh(s, 'arc', 9.11, 3.76, 2.52, 2.52, { line: { color: C.clay, width: 17 }, angleRange: [180.92, 52.31] });
  sh(s, 'arc', 9.34, 3.99, 2.06, 2.06, { line: { color: C.taupe, width: 17 }, angleRange: [180.57, 100.11] });
  sh(s, 'arc', 9.58, 4.21, 1.58, 1.59, { line: { color: C.mist, width: 17 }, angleRange: [180.7, 351.94] });
  sh(s, 'rect', 8.97, 5.11, 0.8, 0.04, { fill: C.white });
  tx(s, 'Networking', 8.97, 3.07, 2.94, 0.4, { fontSize: 18, fontFace: F.bold, bold: true, color: '69208E', align: 'center' });
  sh(s, 'roundRect', 10.36, 3.94, 0.53, 0.24, { fill: C.white, rectRadius: 0.12, shadow: SH.soft });
  sh(s, 'triangle', 10.6, 4.18, 0.06, 0.09, { fill: C.white, flipV: true, shadow: SH.card });
  tx(s, [{ text: '48', options: { fontSize: 9, bold: true, color: 'CA3AB0' } }, { text: '%', options: { fontSize: 9, bold: true, color: 'CA3AB0', baseline: 40 } }], 10.29, 3.97, 0.61, 0.28, { fontSize: 9, bold: true, color: 'CA3AB0', align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  sh(s, 'teardrop', 11.45, 4.17, 0.79, 0.79, { fill: C.taupe, line: { color: C.white, width: 4 }, rotate: 225, shadow: SH.deep });
  tx(s, [{ text: '70', options: { fontSize: 20, fontFace: F.bold, bold: true, color: C.white } }, { text: '%', options: { fontSize: 20, fontFace: F.bold, bold: true, color: C.white, baseline: 40 } }], 11.35, 4.4, 0.94, 0.44, { fontSize: 20, fontFace: F.bold, bold: true, color: C.white, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  sh(s, 'roundRect', 9.07, 3.91, 0.53, 0.24, { fill: C.white, rectRadius: 0.12, shadow: SH.soft });
  sh(s, 'triangle', 9.3, 4.15, 0.06, 0.09, { fill: C.white, flipV: true, shadow: SH.card });
  tx(s, [{ text: '57', options: { fontSize: 9, bold: true, color: C.taupe } }, { text: '%', options: { fontSize: 9, bold: true, color: C.taupe, baseline: 40 } }], 8.99, 3.94, 0.61, 0.28, { fontSize: 9, bold: true, color: C.taupe, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  tx(s, '177K', 8.57, 5.84, 1.69, 0.44, { fontSize: 20, fontFace: F.bold, bold: true, color: '69208E' });
  tx(s, 'Total Data', 8.57, 5.53, 1.69, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: '595959' });
  sh(s, 'ellipse', 9.91, 4.56, 0.98, 0.98, { fill: C.white, flipV: true, shadow: SH.soft });
  title(s, 0.9, 4.99, 6.78, 0.77);
  sh(s, 'roundRect', 6.91, 1.66, 3, 0.23, { fill: C.haze, rectRadius: 0.12 });
  sh(s, 'roundRect', 6.91, 1.66, 1.86, 0.23, { fill: C.navy, rectRadius: 0.12 });
  tx(s, '65%', 9.15, 1.63, 0.7, 0.29, { bold: true, color: '595959', align: 'right' });
  tx(s, 'Database', 7, 1.64, 1.03, 0.27, { fontSize: 10, fontFace: F.bold, bold: true, color: C.white });
  sh(s, 'roundRect', 6.91, 2.12, 3, 0.23, { fill: C.haze, rectRadius: 0.12 });
  sh(s, 'roundRect', 6.91, 2.11, 2.04, 0.23, { fill: C.taupe, rectRadius: 0.12 });
  tx(s, '70%', 9.15, 2.09, 0.7, 0.29, { bold: true, color: '595959', align: 'right' });
  tx(s, 'Networking', 7, 2.1, 1.03, 0.44, { fontSize: 10, fontFace: F.bold, bold: true, color: C.white });
  tx(s, 'Data Accumulation', 6.83, 1.04, 3, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: '595959' });
  tx(s, 'Lorem Ipsum is simply dummy text of the printing ', 10.7, 1.53, 2.01, 0.9, { lineSpacingMultiple: 1.5 });
  tx(s, L1, 0.9, 5.92, 6.17, 0.9, { lineSpacingMultiple: 1.5 });
  icon(s, 'cropMark', 10.27, 4.91, 0.29, 0.29, '000000');
  icon(s, 'timer', 3.04, 2.61, 0.29, 0.29, '000000');
  pageNum(s, 12);
}

function slide13(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 16']);
  sh(s, 'pie', 8.29, 1.86, 3.79, 3.79, { fill: 'E7DCC8', angleRange: [221.83, 112.05] });
  sh(s, 'pie', 8.03, 1.59, 4.32, 4.32, { fill: 'F1F1F1', angleRange: [352.63, 78.31] });
  sh(s, 'pie', 7.73, 1.3, 4.9, 4.9, { fill: 'D5C194', angleRange: [77.99, 250.35] });
  sh(s, 'ellipse', 9.52, 3.08, 1.33, 1.33, { fill: C.white });
  title(s, 0.91, 0.99, 3.63, 1.45);
  tx(s, 'Subtitle Here', 2.27, 2.99, 1.74, 0.34, { fontSize: 14, fontFace: F.bold, color: '262626' });
  tx(s, L2, 2.27, 3.29, 3.84, 0.63, { lineSpacingMultiple: 1.5 });
  tx(s, 'Subtitle Here', 2.27, 4.16, 1.74, 0.34, { fontSize: 14, fontFace: F.bold, color: '262626' });
  tx(s, L2, 2.27, 4.46, 3.84, 0.63, { lineSpacingMultiple: 1.5 });
  tx(s, 'Subtitle Here', 2.27, 5.54, 1.74, 0.34, { fontSize: 14, fontFace: F.bold, color: '262626' });
  tx(s, L2, 2.27, 5.84, 3.84, 0.63, { lineSpacingMultiple: 1.5 });
  sh(s, 'ellipse', 1, 2.86, 0.97, 0.97, { fill: C.white, shadow: SH.soft });
  tx(s, [{ text: '15', options: { fontSize: 24, fontFace: F.bold, bold: true, color: C.gold } }, { text: '%', options: { fontSize: 24, fontFace: F.bold, bold: true, color: C.gold, baseline: 40 } }], 1.04, 3.04, 0.95, 0.46, { fontSize: 24, fontFace: F.bold, bold: true, color: C.gold, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  sh(s, 'ellipse', 1, 4.16, 0.97, 0.97, { fill: C.white, shadow: SH.soft });
  tx(s, [{ text: '15', options: { fontSize: 24, fontFace: F.bold, bold: true, color: C.sand } }, { text: '%', options: { fontSize: 24, fontFace: F.bold, bold: true, color: C.sand, baseline: 40 } }], 1.04, 4.35, 0.95, 0.46, { fontSize: 24, fontFace: F.bold, bold: true, color: C.sand, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  sh(s, 'ellipse', 1, 5.54, 0.97, 0.97, { fill: C.white, shadow: SH.soft });
  tx(s, [{ text: '15', options: { fontSize: 24, fontFace: F.bold, bold: true, color: C.mist } }, { text: '%', options: { fontSize: 24, fontFace: F.bold, bold: true, color: C.mist, baseline: 40 } }], 1.04, 5.73, 0.95, 0.46, { fontSize: 24, fontFace: F.bold, bold: true, color: C.mist, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  icon(s, 'rulerPencil', 9.99, 3.55, 0.4, 0.4, '000000');
  pageNum(s, 13);
}

function slide14(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 20']);
  sh(s, 'ellipse', -2.08, 1.04, 5.42, 5.42, { line: { color: C.haze, width: 40 }, rotate: 90 });
  sh(s, 'pie', -2.34, 0.43, 6.64, 6.64, { fill: C.taupe, angleRange: [201.45, 229.72], rotate: 90, shadow: SH.lift });
  sh(s, 'pie', -3.17, -0.4, 8.29, 8.29, { fill: C.sand, angleRange: [229.52, 269.04], rotate: 90, shadow: SH.lift });
  sh(s, 'pie', -2.34, 0.43, 6.64, 6.64, { fill: C.navy, angleRange: [269.03, 299.23], rotate: 90, shadow: SH.lift });
  sh(s, 'pie', -2.53, 0.24, 7.02, 7.02, { fill: C.mist, angleRange: [299.01, 329.1], rotate: 90, shadow: SH.lift });
  title(s, 6.39, 1.26, 3.63, 1.45);
  tx(s, L4, 6.39, 3.69, 2.82, 1.18, { lineSpacingMultiple: 1.5 });
  tx(s, 'Your Text Here', 6.39, 3.35, 2.27, 0.37, { fontSize: 16, fontFace: F.bold, bold: true, color: C.navy });
  tx(s, L4, 9.55, 3.69, 2.82, 1.18, { lineSpacingMultiple: 1.5 });
  tx(s, 'Your Text Here', 9.55, 3.35, 2.27, 0.37, { fontSize: 16, fontFace: F.bold, bold: true, color: C.sand });
  tx(s, L4, 6.39, 5.34, 2.82, 1.18, { lineSpacingMultiple: 1.5 });
  tx(s, 'Your Text Here', 6.39, 4.99, 2.27, 0.37, { fontSize: 16, fontFace: F.bold, bold: true, color: C.taupe });
  tx(s, L4, 9.55, 5.34, 2.82, 1.18, { lineSpacingMultiple: 1.5 });
  tx(s, 'Your Text Here', 9.55, 4.99, 2.27, 0.37, { fontSize: 16, fontFace: F.bold, bold: true, color: C.mist });
  icon(s, 'cropFrame', 2.19, 1.57, 0.42, 0.42, C.white);
  icon(s, 'moveArrows', 2.81, 5.51, 0.42, 0.42, C.white);
  icon(s, 'scissors', 3.78, 2.42, 0.38, 0.4, C.white);
  icon(s, 'barTile', 3.3, 4.14, 0.42, 0.38, C.white);
  pageNum(s, 14);
}

function slide15(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 21']);
  sh(s, 'ellipse', 4.53, 1.61, 4.28, 4.28, { fill: C.white, shadow: SH.lift });
  sh(s, 'ellipse', 5.25, 2.33, 2.83, 2.83, { line: { color: '8A9AA8', width: 16.5, transparency: 90 } });
  sh(s, 'arc', 5.25, 2.33, 2.83, 2.83, { line: { color: 'E7DCC8', width: 12 }, angleRange: [270, 186.84], shadow: tint(C.sand) });
  sh(s, 'ellipse', 5.6, 2.67, 2.14, 2.14, { line: { color: '8A9AA8', width: 16.5, transparency: 90 } });
  sh(s, 'arc', 5.6, 2.67, 2.14, 2.14, { line: { color: 'D6C294', width: 12 }, angleRange: [270, 163.66], shadow: tint(C.gold) });
  sh(s, 'ellipse', 5.94, 3.02, 1.45, 1.45, { line: { color: '8A9AA8', width: 16.5, transparency: 90 } });
  sh(s, 'arc', 5.94, 3.02, 1.45, 1.45, { line: { color: '738EA2', width: 12 }, angleRange: [270, 91.1], shadow: tint(C.navy) });
  sh(s, 'ellipse', 4.91, 1.98, 3.52, 3.52, { line: { color: '8A9AA8', width: 16.5, transparency: 90 } });
  sh(s, 'arc', 4.91, 1.98, 3.52, 3.52, { line: { color: C.haze, width: 12 }, angleRange: [270, 37.2], shadow: tint(C.mist) });
  sh(s, 'ellipse', 6.22, 3.29, 0.9, 0.9, { fill: C.taupe, shadow: SH.lift });
  statCard(s, 9.19, 1.04, 3.16, { ring: C.gold, pct: '53%', head: 'Your Text Here', headColor: C.gold, body: 'Lorem ipsum dolor sit amet, ', bodyW: 1.73, bodyBold: true });
  statCard(s, 9.19, 2.47, 3.16, { ring: C.navy, pct: '53%', head: 'Your Text Here', headColor: C.navy, body: 'Lorem ipsum dolor sit amet, ', bodyW: 1.73, bodyBold: true });
  statCard(s, 9.19, 3.91, 3.16, { ring: C.sand, pct: '53%', head: 'Your Text Here', headColor: C.sand, body: 'Lorem ipsum dolor sit amet, ', bodyW: 1.73, bodyBold: true });
  statCard(s, 9.19, 5.31, 3.16, { ring: C.mist, pct: '53%', head: 'Your Text Here', headColor: C.mist, body: 'Lorem ipsum dolor sit amet, ', bodyW: 1.73, bodyBold: true });
  title(s, 0.91, 1.66, 3.63, 1.45);
  tx(s, L1, 0.93, 3.58, 3.32, 1.46, { lineSpacingMultiple: 1.5 });
  tx(s, 'Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry\'s', 0.93, 4.93, 3.32, 0.9, { lineSpacingMultiple: 1.5 });
  icon(s, 'terminal', 6.53, 3.65, 0.27, 0.25, C.white);
  pageNum(s, 15);
}

function slide16(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 22']);
  tx(s, '44%', 1.05, 3.23, 1.3, 0.51, { fontSize: 24, fontFace: F.bold, bold: true, color: C.navy });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.', 1.05, 4.03, 2.63, 0.53, { color: C.slate, lineSpacingMultiple: 1.2 });
  tx(s, 'Your Text Here', 1.05, 3.69, 2.1, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: C.dark });
  tx(s, '44%', 3.5, 3.23, 1.3, 0.51, { fontSize: 24, fontFace: F.bold, bold: true, color: C.gold });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.', 3.5, 4.03, 2.63, 0.53, { color: C.slate, lineSpacingMultiple: 1.2 });
  tx(s, 'Your Text Here', 3.5, 3.69, 2.1, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: C.dark });
  tx(s, '44%', 1.05, 4.78, 1.3, 0.51, { fontSize: 24, fontFace: F.bold, bold: true, color: C.sand });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.', 1.05, 5.58, 2.63, 0.53, { color: C.slate, lineSpacingMultiple: 1.2 });
  tx(s, 'Your Text Here', 1.05, 5.24, 2.1, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: C.dark });
  tx(s, '44%', 3.5, 4.78, 1.3, 0.51, { fontSize: 24, fontFace: F.bold, bold: true, color: C.mist });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.', 3.5, 5.58, 2.63, 0.53, { color: C.slate, lineSpacingMultiple: 1.2 });
  tx(s, 'Your Text Here', 3.5, 5.24, 2.1, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: C.dark });
  title(s, 1.05, 1.06, 3.63, 1.45);
  tx(s, L1, 7.47, 1.2, 4.81, 0.9, { lineSpacingMultiple: 1.5 });
  isoChart(s, 7.52, 2.78, 4.71, 2.66, ['A29374', 'D8C7A0', '213C66', '718CA0', C.mist, 'E4DBC8', 'D8C7A0', 'A29374'], true);
  legend(s, 7.46, 5.87, 1.37, [['D8C7A0', 'Chart 1'], ['718CA0', 'Chart 2'], [C.mist, 'Chart 3'], ['E4DBC8', 'Chart 4']], { sw: 0.22, tw: 0.9, ty: -0.06 });
  pageNum(s, 16);
}

function slide17(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 23']);
  sh(s, 'ellipse', 5.15, 2.87, 3.03, 3.03, { line: { color: C.haze, width: 25 } });
  sh(s, 'pie', 4.76, 2.47, 3.8, 3.8, { fill: C.gold, angleRange: [90.42, 135.36], rotate: 224.9 });
  sh(s, 'pie', 4.66, 2.37, 4.01, 4.01, { fill: C.navy, angleRange: [90.23, 135.36], rotate: 270, shadow: SH.soft });
  sh(s, 'pie', 4.57, 2.29, 4.2, 4.2, { fill: C.taupe, angleRange: [5.96, 135.36], rotate: 180, shadow: SH.soft });
  sh(s, 'pie', 4.76, 2.47, 3.8, 3.8, { fill: C.mist, angleRange: [90.23, 135.36], rotate: 315 });
  sh(s, 'ellipse', 5.97, 3.68, 1.4, 1.4, { fill: C.white, shadow: SH.card });
  sh(s, 'roundRect', 9.51, 3.23, 0.51, 0.51, { fill: C.gold, shadow: tint(C.gold) });
  tx(s, 'Your Text Here', 10.18, 3.1, 2.09, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: C.gold });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur.', 10.18, 3.47, 2.09, 0.57, { fontSize: 12, color: C.slate, lineSpacingMultiple: 1.2 });
  sh(s, 'roundRect', 9.52, 5.02, 0.5, 0.5, { fill: C.navy, shadow: tint(C.mist) });
  tx(s, 'Your Text Here', 10.18, 4.93, 2.09, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: C.navy });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur', 10.18, 5.25, 2.03, 0.57, { fontSize: 12, color: C.slate, lineSpacingMultiple: 1.2 });
  sh(s, 'roundRect', 1.13, 3.23, 0.51, 0.51, { fill: C.taupe, shadow: tint(C.gold) });
  tx(s, 'Your Text Here', 1.79, 3.1, 2.09, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: C.taupe });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur.', 1.79, 3.47, 2.09, 0.57, { fontSize: 12, color: C.slate, lineSpacingMultiple: 1.2 });
  sh(s, 'roundRect', 1.13, 5.02, 0.5, 0.5, { fill: C.mist, shadow: tint(C.mist) });
  tx(s, 'Your Text Here', 1.79, 4.93, 2.09, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: C.mist });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur', 1.79, 5.25, 2.03, 0.57, { fontSize: 12, color: C.slate, lineSpacingMultiple: 1.2 });
  title(s, 3.2, 1.01, 6.92, 0.77, { align: 'center' });
  icon(s, 'photoTile', 9.64, 3.38, 0.25, 0.23, C.white);
  icon(s, 'castScreen', 6.46, 4.19, 0.42, 0.38, '000000');
  icon(s, 'docList', 1.24, 5.18, 0.25, 0.23, C.white);
  icon(s, 'laptop', 9.63, 5.15, 0.28, 0.23, C.white);
  icon(s, 'floppy', 1.27, 3.37, 0.23, 0.23, C.white);
  pageNum(s, 17);
}

function slide18(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 24']);
  tx(s, '17k', 1.41, 4.82, 0.97, 0.51, { fontSize: 24, fontFace: F.bold, bold: true, color: 'F7A900' });
  tx(s, 'Lorem ipsum dolor sit amet', 1.41, 5.25, 1.55, 0.53, { color: C.slate, lineSpacingMultiple: 1.2 });
  tx(s, '31%', 1.41, 6.06, 1.26, 0.41, { fontSize: 16, fontFace: F.bold, bold: true, color: C.dark, lineSpacingMultiple: 1.2 });
  sh(s, 'rect', 1.49, 6.47, 1.31, 0.07, { fill: { color: C.haze, transparency: 12 } });
  sh(s, 'rect', 1.49, 6.47, 0.96, 0.07, { fill: C.sand });
  tx(s, '17k', 2.96, 4.82, 0.97, 0.51, { fontSize: 24, fontFace: F.bold, bold: true, color: C.taupe });
  tx(s, 'Lorem ipsum dolor sit amet', 2.96, 5.25, 1.55, 0.53, { color: C.slate, lineSpacingMultiple: 1.2 });
  tx(s, '31%', 2.96, 6.06, 1.26, 0.41, { fontSize: 16, fontFace: F.bold, bold: true, color: C.dark, lineSpacingMultiple: 1.2 });
  sh(s, 'rect', 3.04, 6.47, 1.31, 0.07, { fill: { color: C.haze, transparency: 12 } });
  sh(s, 'rect', 3.04, 6.47, 0.96, 0.07, { fill: C.taupe });
  tx(s, '17k', 4.51, 4.82, 0.97, 0.51, { fontSize: 24, fontFace: F.bold, bold: true, color: C.navy });
  tx(s, 'Lorem ipsum dolor sit amet', 4.51, 5.25, 1.55, 0.53, { color: C.slate, lineSpacingMultiple: 1.2 });
  tx(s, '31%', 4.51, 6.06, 1.26, 0.41, { fontSize: 16, fontFace: F.bold, bold: true, color: C.dark, lineSpacingMultiple: 1.2 });
  sh(s, 'rect', 4.59, 6.47, 1.31, 0.07, { fill: { color: C.haze, transparency: 12 } });
  sh(s, 'rect', 4.59, 6.47, 0.96, 0.07, { fill: C.navy });
  title(s, 7.22, 1.72, 3.63, 1.45);
  tx(s, [{ text: L5, options: {  } }, { text: '1500s dummy text of the printing and typesetting industry. Lorem Ipsum ', options: {  } }, { text: 'hing', options: {  } }, { text: ' and typesetting has ', options: {  } }], 7.22, 3.63, 5.35, 1.18, { lineSpacingMultiple: 1.5 });
  tx(s, [{ text: L5, options: {  } }, { text: '1500s ', options: {  } }, { text: 'ing', options: {  } }, { text: ' and typesetting', options: {  } }], 7.22, 4.87, 5.35, 0.9, { lineSpacingMultiple: 1.5 });
  sh(s, 'pie', 1.89, 1.12, 2.74, 2.73, { fill: C.sand, angleRange: [277.47, 107.41] });
  sh(s, 'pie', 1.89, 1.12, 2.74, 2.73, { fill: C.navy, angleRange: [124.69, 283.13] });
  sh(s, 'pie', 1.89, 1.12, 2.74, 2.73, { fill: C.navy, angleRange: [205.33, 249.51] });
  sh(s, 'pie', 1.89, 1.12, 2.74, 2.73, { fill: C.taupe, angleRange: [103.47, 215.77] });
  legend(s, 1.38, 4.17, 1.37, [[C.navy, 'Data 1'], [C.taupe, 'Data 2'], [C.sand, 'Data 3']], { sw: 0.17, tw: 0.9, ty: -0.09 });
  pageNum(s, 18);
}

function slide19(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 20']);
  sh(s, 'roundRect', 8.8, 2.23, 3.22, 2.45, { fill: C.white, rectRadius: 0.09, shadow: SH.card });
  pieChart(s, 4.86, 2.23, 3.54, 2.45);
  sh(s, 'roundRect', 5.06, 2.23, 3.22, 2.45, { fill: C.white, rectRadius: 0.09, shadow: SH.card });
  sh(s, 'roundRect', 1.32, 2.23, 3.22, 2.45, { fill: C.white, rectRadius: 0.09, shadow: SH.card });
  sh(s, 'pie', 5.86, 2.69, 1.55, 1.55, { fill: C.white, rotate: 225, flipH: true, shadow: SH.card });
  sh(s, 'ellipse', 5.98, 2.83, 1.29, 1.29, { fill: C.white, flipH: true, shadow: SH.lift });
  tx(s, [{ text: '43', options: { fontSize: 28, fontFace: F.bold, bold: true, color: 'F7A900' } }, { text: '%', options: { fontSize: 14, fontFace: F.bold, bold: true, color: 'F7A900' } }], 6.19, 3.41, 0.86, 0.57, { fontSize: 28, fontFace: F.bold, bold: true, color: 'F7A900', align: 'center', wrap: false });
  tx(s, 'Market Goal', 6.03, 2.99, 1.2, 0.51, { fontSize: 12, color: '262626', align: 'center' });
  sh(s, 'line', 5.98, 2.89, 0, 0.04, { line: { color: '8A9AA8', width: 1.5, transparency: 80 }, rotate: 307.7, flipH: true });
  sh(s, 'line', 5.84, 3.09, 0, 0.04, { line: { color: '8A9AA8', width: 1.5, transparency: 80 }, rotate: 293.9, flipH: true });
  sh(s, 'line', 6.16, 2.74, 0, 0.04, { line: { color: '8A9AA8', width: 1.5, transparency: 80 }, rotate: 322.2, flipH: true });
  sh(s, 'line', 6.38, 2.64, 0, 0.04, { line: { color: '8A9AA8', width: 1.5, transparency: 80 }, rotate: 340.4, flipH: true });
  sh(s, 'line', 7.28, 2.89, 0, 0.04, { line: { color: '8A9AA8', width: 1.5, transparency: 80 }, rotate: 52.3 });
  sh(s, 'line', 7.41, 3.09, 0, 0.04, { line: { color: '8A9AA8', width: 1.5, transparency: 80 }, rotate: 66.1 });
  sh(s, 'line', 7.1, 2.74, 0, 0.04, { line: { color: '8A9AA8', width: 1.5, transparency: 80 }, rotate: 37.8 });
  sh(s, 'line', 6.88, 2.64, 0, 0.04, { line: { color: '8A9AA8', width: 1.5, transparency: 80 }, rotate: 19.6 });
  sh(s, 'line', 5.77, 3.57, 0, 0.04, { line: { color: '8A9AA8', width: 1.5, transparency: 80 }, rotate: 86.9, flipH: true, flipV: true });
  sh(s, 'line', 5.98, 4.01, 0, 0.04, { line: { color: '8A9AA8', width: 1.5, transparency: 80 }, rotate: 52.3, flipH: true, flipV: true });
  sh(s, 'line', 5.84, 3.81, 0, 0.04, { line: { color: '8A9AA8', width: 1.5, transparency: 80 }, rotate: 66.1, flipH: true, flipV: true });
  sh(s, 'line', 6.16, 4.17, 0, 0.04, { line: { color: '8A9AA8', width: 1.5, transparency: 80 }, rotate: 37.8, flipH: true, flipV: true });
  sh(s, 'line', 6.38, 4.26, 0, 0.04, { line: { color: '8A9AA8', width: 1.5, transparency: 80 }, rotate: 19.6, flipH: true, flipV: true });
  sh(s, 'line', 7.49, 3.57, 0, 0.04, { line: { color: '8A9AA8', width: 1.5, transparency: 80 }, rotate: 273.1, flipV: true });
  sh(s, 'line', 7.28, 4.01, 0, 0.04, { line: { color: '8A9AA8', width: 1.5, transparency: 80 }, rotate: 307.7, flipV: true });
  sh(s, 'line', 7.41, 3.81, 0, 0.04, { line: { color: '8A9AA8', width: 1.5, transparency: 80 }, rotate: 293.9, flipV: true });
  sh(s, 'line', 7.1, 4.17, 0, 0.04, { line: { color: '8A9AA8', width: 1.5, transparency: 80 }, rotate: 322.2, flipV: true });
  sh(s, 'line', 6.88, 4.26, 0, 0.04, { line: { color: '8A9AA8', width: 1.5, transparency: 80 }, rotate: 340.4, flipV: true });
  sh(s, 'line', 6.63, 2.6, 0, 0.04, { line: { color: '8A9AA8', width: 1.5, transparency: 80 }, flipH: true });
  sh(s, 'line', 6.63, 4.29, 0, 0.04, { line: { color: '8A9AA8', width: 1.5, transparency: 80 }, flipH: true });
  sh(s, 'line', 5.75, 3.37, 0.04, 0, { line: { color: '8A9AA8', width: 1.5, transparency: 80 } });
  sh(s, 'line', 7.44, 3.37, 0.04, 0, { line: { color: '8A9AA8', width: 1.5, transparency: 80 } });
  sh(s, 'arc', 5.86, 2.69, 1.55, 1.55, { line: { color: 'F5F5F5', width: 4 }, angleRange: [78.63, 240.05], rotate: 150, flipH: true });
  sh(s, 'ellipse', 6.62, 2.68, 0.02, 0.02, { fill: C.white, flipH: true });
  sh(s, 'blockArc', 1.89, 2.61, 2.12, 2.12, { fill: 'AB9891', angleRange: [295.49, 359.6], arcThicknessRatio: 0.32 });
  sh(s, 'blockArc', 1.89, 2.61, 2.12, 2.12, { fill: 'D5C194', angleRange: [316.62, 359.6], arcThicknessRatio: 0.32, flipH: true });
  sh(s, 'blockArc', 1.89, 2.61, 2.12, 2.12, { fill: 'E7DCC8', angleRange: [244.24, 316.9], arcThicknessRatio: 0.32, flipH: true });
  sh(s, 'pie', 2.52, 3.23, 0.87, 0.87, { fill: C.haze, angleRange: [89.04, 270], rotate: 90, shadow: SH.card });
  tx(s, '70', 3.66, 3.63, 0.35, 0.52, { fontSize: 10, color: C.dark, align: 'center', lineSpacingMultiple: 1.3 });
  tx(s, '0', 1.91, 3.63, 0.29, 0.3, { fontSize: 10, color: C.dark, align: 'center', lineSpacingMultiple: 1.3 });
  tx(s, '40,214', 2.52, 3.34, 0.87, 0.29, { fontFace: F.bold, bold: true, color: C.dark, align: 'center', valign: 'middle' });
  sh(s, 'rect', 2.09, 4.21, 0.1, 0.1, { fill: C.gold });
  tx(s, 'Jul', 2.2, 4.12, 0.47, 0.28, { fontSize: 10.5, color: '000000' });
  sh(s, 'rect', 2.6, 4.21, 0.1, 0.1, { fill: C.sand });
  tx(s, 'Aug', 2.71, 4.12, 0.61, 0.28, { fontSize: 10.5, color: '000000' });
  sh(s, 'rect', 3.19, 4.2, 0.1, 0.1, { fill: C.taupe });
  tx(s, 'Sep', 3.3, 4.11, 0.61, 0.28, { fontSize: 10.5, color: '000000' });
  title(s, 3.2, 1.01, 6.92, 0.77, { align: 'center' });
  tx(s, 'Data Chart One', 1.74, 5.2, 2.43, 0.37, { fontSize: 16, fontFace: F.bold, bold: true, color: C.dark, align: 'center' });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin convallis,.', 1.66, 5.71, 2.57, 0.75, { color: C.slate, align: 'center', lineSpacingMultiple: 1.2 });
  tx(s, 'Data Chart Two', 5.41, 5.2, 2.43, 0.37, { fontSize: 16, fontFace: F.bold, bold: true, color: C.dark, align: 'center' });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin convallis,.', 5.33, 5.71, 2.57, 0.75, { color: C.slate, align: 'center', lineSpacingMultiple: 1.2 });
  tx(s, 'Data Chart Three', 9.19, 5.2, 2.43, 0.37, { fontSize: 16, fontFace: F.bold, bold: true, color: C.dark, align: 'center' });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin convallis,.', 9.12, 5.71, 2.57, 0.75, { color: C.slate, align: 'center', lineSpacingMultiple: 1.2 });
  sh(s, 'pie', 9.76, 2.62, 1.29, 1.29, { fill: 'B2A19B', angleRange: [277.47, 107.41] });
  sh(s, 'pie', 9.76, 2.62, 1.29, 1.29, { fill: 'E3D2C5', angleRange: [124.69, 283.13] });
  sh(s, 'pie', 9.76, 2.62, 1.29, 1.29, { fill: 'E3D2C5', angleRange: [205.33, 249.51] });
  sh(s, 'pie', 9.76, 2.62, 1.29, 1.29, { fill: 'EBE2D2', angleRange: [103.47, 215.77] });
  legend(s, 9.02, 4.08, 1.04, [['EBE2D2', 'Data 1'], ['E3D2C5', 'Data 2'], ['B2A19B', 'Data 3']], { sw: 0.1, tw: 0.9, ty: -0.12 });
  pageNum(s, 19);
}

function slide20(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 24']);
  sh(s, 'ellipse', 8.58, 2.88, 2.64, 2.64, { line: { color: C.haze, width: 10 } });
  sh(s, 'ellipse', 7.67, 4.09, 2.02, 2.02, { fill: 'F6F6F6', shadow: SH.soft });
  sh(s, 'pie', 7.83, 4.25, 1.72, 1.72, { fill: 'D5CBC8', shadow: tint(C.mist) });
  sh(s, 'pie', 7.83, 4.25, 1.72, 1.72, { fill: C.taupe, angleRange: [179.96, 0.13], shadow: tint(C.mist) });
  sh(s, 'ellipse', 8.1, 4.52, 1.17, 1.17, { fill: C.white });
  tx(s, [{ text: '62', options: { fontSize: 18, fontFace: F.bold, bold: true, color: 'FF593F' } }, { text: '%', options: { fontSize: 18, fontFace: F.bold, bold: true, color: 'FF593F', baseline: 600 } }], 8.22, 4.9, 0.97, 0.4, { fontSize: 18, fontFace: F.bold, bold: true, color: 'FF593F', align: 'center', valign: 'middle' });
  sh(s, 'ellipse', 10.11, 4.09, 2.02, 2.02, { fill: 'F6F6F6', shadow: SH.soft });
  sh(s, 'pie', 10.27, 4.25, 1.72, 1.72, { fill: 'F8F8F8', shadow: tint(C.mist) });
  sh(s, 'pie', 10.27, 4.25, 1.72, 1.72, { fill: C.mist, angleRange: [268.96, 118.7], shadow: tint(C.clay) });
  sh(s, 'ellipse', 10.54, 4.52, 1.17, 1.17, { fill: C.white });
  tx(s, [{ text: '78', options: { fontSize: 18, fontFace: F.bold, bold: true, color: C.mist } }, { text: '%', options: { fontSize: 18, fontFace: F.bold, bold: true, color: C.mist, baseline: 600 } }], 10.66, 4.97, 0.97, 0.4, { fontSize: 18, fontFace: F.bold, bold: true, color: C.mist, align: 'center', valign: 'middle' });
  sh(s, 'ellipse', 8.45, 1.12, 2.91, 2.91, { fill: 'F6F6F6', shadow: SH.soft });
  sh(s, 'pie', 8.67, 1.34, 2.47, 2.47, { fill: 'EEE7D9', shadow: tint(C.gold) });
  sh(s, 'pie', 8.67, 1.34, 2.47, 2.47, { fill: C.sand, angleRange: [268.96, 24.36], shadow: tint(C.gold) });
  sh(s, 'ellipse', 9.06, 1.74, 1.68, 1.68, { fill: C.white });
  tx(s, [{ text: '23', options: { fontSize: 32, fontFace: F.bold, bold: true, color: C.sand } }, { text: '%', options: { fontSize: 32, fontFace: F.bold, bold: true, color: C.sand, baseline: 600 } }], 9.4, 2.3, 1, 0.64, { fontSize: 32, fontFace: F.bold, bold: true, color: C.sand, align: 'center', wrap: false });
  title(s, 1.05, 1.09, 3.63, 1.45);
  sh(s, 'roundRect', 1.13, 3.1, 0.51, 0.51, { fill: C.taupe, shadow: tint(C.gold) });
  tx(s, 'Your Text Here', 1.79, 3.04, 2.09, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: C.taupe });
  tx(s, L2, 1.79, 3.35, 3.87, 0.68, { fontSize: 12, lineSpacingMultiple: 1.5 });
  sh(s, 'roundRect', 1.13, 4.3, 0.51, 0.51, { fill: C.sand, shadow: tint(C.gold) });
  tx(s, 'Your Text Here', 1.79, 4.24, 2.09, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: C.sand });
  tx(s, L2, 1.79, 4.54, 3.87, 0.68, { fontSize: 12, lineSpacingMultiple: 1.5 });
  sh(s, 'roundRect', 1.13, 5.49, 0.51, 0.51, { fill: C.mist, shadow: tint(C.gold) });
  tx(s, 'Your Text Here', 1.79, 5.43, 2.09, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: C.mist });
  tx(s, L2, 1.79, 5.73, 3.87, 0.68, { fontSize: 12, lineSpacingMultiple: 1.5 });
  tx(s, '46,921+', 4.28, 3.04, 1.39, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: 'FF593F' });
  tx(s, '46,921+', 4.28, 4.24, 1.39, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: C.sand });
  tx(s, '46,921+', 4.28, 5.39, 1.39, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: C.mist });
  icon(s, 'photoTile2', 1.24, 3.22, 0.3, 0.27, C.white);
  icon(s, 'keyboard', 1.22, 5.61, 0.3, 0.27, C.white);
  icon(s, 'sheet', 1.27, 4.4, 0.24, 0.3, C.white);
  pageNum(s, 20);
}

function slide21(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 26']);
  title(s, 1.05, 1.46, 3.63, 1.45);
  statCard(s, 1.05, 4.02, 4.96, { ring: C.taupe, pct: '53%', head: 'Your Text Here', headColor: C.taupe });
  statCard(s, 1.05, 5.46, 4.96, { ring: C.navy, pct: '53%', head: 'Your Text Here', headColor: C.navy, body: L2, bodyW: 3.6 });
  tx(s, 'Lorem Ipsum is simply dummy text of the printing and typesetting industry.', 2.23, 5.87, 3.49, 0.62, { lineSpacingMultiple: 1.5 });
  tx(s, '56%', 9.06, 5.39, 0.57, 0.34, { fontSize: 12, fontFace: F.bold, bold: true, color: C.dark, lineSpacingMultiple: 1.3 });
  sh(s, 'roundRect', 7.72, 5.75, 1.9, 0.08, { fill: { color: 'D9D9D9', transparency: 80 }, rectRadius: 0.04 });
  sh(s, 'roundRect', 7.72, 5.75, 0.9, 0.08, { fill: C.navy, rectRadius: 0.04 });
  tx(s, 'Data One', 7.61, 5.38, 1.52, 0.33, { fontSize: 12, color: C.dark, lineSpacingMultiple: 1.2 });
  tx(s, '62%', 9.06, 5.88, 0.57, 0.34, { fontSize: 12, fontFace: F.bold, bold: true, color: C.dark, lineSpacingMultiple: 1.3 });
  tx(s, 'Data Two', 7.61, 5.87, 1.44, 0.33, { fontSize: 12, color: C.dark, lineSpacingMultiple: 1.2 });
  sh(s, 'roundRect', 7.72, 6.24, 1.9, 0.08, { fill: { color: 'D9D9D9', transparency: 80 }, rectRadius: 0.04 });
  sh(s, 'roundRect', 7.72, 6.24, 1.08, 0.08, { fill: C.taupe, rectRadius: 0.04 });
  tx(s, '71%', 11.71, 5.39, 0.57, 0.34, { fontSize: 12, fontFace: F.bold, bold: true, color: C.dark, lineSpacingMultiple: 1.3 });
  sh(s, 'roundRect', 10.38, 5.75, 1.9, 0.08, { fill: { color: 'D9D9D9', transparency: 80 }, rectRadius: 0.04 });
  sh(s, 'roundRect', 10.38, 5.75, 0.99, 0.08, { fill: C.mist, rectRadius: 0.04 });
  tx(s, 'Data Three', 10.27, 5.38, 1.32, 0.33, { fontSize: 12, color: C.dark, lineSpacingMultiple: 1.2 });
  tx(s, '47%', 11.71, 5.88, 0.57, 0.34, { fontSize: 12, fontFace: F.bold, bold: true, color: C.dark, lineSpacingMultiple: 1.3 });
  sh(s, 'roundRect', 10.38, 6.24, 1.9, 0.08, { fill: { color: 'D9D9D9', transparency: 80 }, rectRadius: 0.04 });
  sh(s, 'roundRect', 10.38, 6.24, 1.15, 0.08, { fill: C.sand, rectRadius: 0.04 });
  tx(s, 'Data Four', 10.27, 5.87, 1.32, 0.33, { fontSize: 12, color: C.dark, lineSpacingMultiple: 1.2 });
  sh(s, 'roundRect', 7.72, 4.44, 0.51, 0.51, { fill: C.navy, shadow: tint(C.gold) });
  tx(s, 'Timing And Progress', 8.3, 4.54, 2.93, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: '000000' });
  icon(s, 'floppy2', 7.84, 4.57, 0.26, 0.26, C.white);
  isoChart(s, 8.49, 1.39, 3.03, 1.71, ['283743', '527085', '283743', '527085', 'DDD1B9', 'EBEBEB', '937C74', '79655F'], true);
  legend(s, 7.74, 3.4, 1.18, [['937C74', 'Chart 1'], ['527085', 'Chart 2'], ['DDD1B9', 'Chart 3'], ['EBEBEB', 'Chart 4']], { sw: 0.16, tw: 0.9, ty: -0.09 });
  pageNum(s, 21);
}

function slide22(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 27']);
  sh(s, 'ellipse', 8.14, 2.28, 3.76, 3.76, { line: { color: C.haze, width: 25 } });
  sh(s, 'pie', 7.42, 1.56, 5.21, 5.21, { fill: C.navy, angleRange: [330.51, 135.36], rotate: 180, shadow: SH.soft });
  sh(s, 'ellipse', 9.16, 3.29, 1.73, 1.73, { fill: C.white, shadow: SH.card });
  tx(s, [{ text: '56', options: { fontSize: 40, fontFace: F.bold, bold: true, color: C.white } }, { text: '%', options: { fontSize: 40, fontFace: F.bold, bold: true, color: C.white, baseline: 600 } }], 8.21, 2.64, 1.76, 0.77, { fontSize: 40, fontFace: F.bold, bold: true, color: C.white, align: 'center' });
  title(s, 1.05, 1.22, 3.63, 1.45);
  tx(s, L3, 1.05, 5.65, 5.21, 0.9, { lineSpacingMultiple: 1.5 });
  tx(s, 'Chart Data Slide', 1.05, 5.07, 2.31, 0.4, { fontSize: 18, fontFace: F.bold, bold: true, color: C.dark });
  tx(s, '112M', 1.05, 4.2, 2.15, 0.84, { fontSize: 44, fontFace: F.bold, bold: true, color: C.taupe });
  tx(s, [{ text: L5, options: {  } }, { text: '1500s ', options: {  } }, { text: 'ing', options: {  } }, { text: ' and typesetting', options: {  } }], 1.05, 2.94, 5.35, 0.9, { lineSpacingMultiple: 1.5 });
  icon(s, 'target', 9.71, 3.84, 0.62, 0.62, '000000');
  pageNum(s, 22);
}

function slide23(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 28']);
  title(s, 3.2, 1.01, 6.92, 0.77, { align: 'center' });
  sh(s, 'roundRect', 9.51, 3.23, 0.51, 0.51, { fill: C.gold, shadow: tint(C.gold) });
  tx(s, 'Your Text Here', 10.18, 3.15, 2.09, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: C.gold });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur.', 10.18, 3.47, 2.09, 0.57, { fontSize: 12, color: C.slate, lineSpacingMultiple: 1.2 });
  sh(s, 'roundRect', 9.52, 5.02, 0.5, 0.5, { fill: C.navy, shadow: tint(C.mist) });
  tx(s, 'Your Text Here', 10.18, 4.97, 2.09, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: C.navy });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur', 10.18, 5.29, 2.03, 0.57, { fontSize: 12, color: C.slate, lineSpacingMultiple: 1.2 });
  sh(s, 'roundRect', 1.13, 3.23, 0.51, 0.51, { fill: C.taupe, shadow: tint(C.gold) });
  tx(s, 'Your Text Here', 1.79, 3.15, 2.09, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: C.taupe });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur.', 1.79, 3.47, 2.09, 0.57, { fontSize: 12, color: C.slate, lineSpacingMultiple: 1.2 });
  sh(s, 'roundRect', 1.13, 5.02, 0.5, 0.5, { fill: 'B2B2B2', shadow: tint(C.mist) });
  tx(s, 'Your Text Here', 1.79, 4.97, 2.09, 0.34, { fontSize: 14, fontFace: F.bold, bold: true, color: 'B2B2B2' });
  tx(s, 'Lorem ipsum dolor sit amet, consectetur', 1.79, 5.29, 2.03, 0.57, { fontSize: 12, color: C.slate, lineSpacingMultiple: 1.2 });
  icon(s, 'bookMark', 9.66, 5.14, 0.22, 0.24, C.white);
  icon(s, 'contactBook', 1.27, 5.16, 0.26, 0.24, C.white);
  icon(s, 'copyPages', 1.27, 3.36, 0.22, 0.24, C.white);
  icon(s, 'clipboard', 9.66, 3.34, 0.22, 0.27, C.white);
  sh(s, 'pie', 5.07, 2.74, 3.4, 3.4, { fill: C.navy, angleRange: [280.29, 107.41] });
  sh(s, 'pie', 5.04, 2.68, 3.4, 3.4, { fill: C.gold, angleRange: [249.16, 277.75] });
  sh(s, 'pie', 4.83, 2.45, 3.4, 3.4, { fill: C.sand, angleRange: [215.51, 249.51] });
  sh(s, 'pie', 4.91, 2.63, 3.4, 3.4, { fill: C.mist, angleRange: [110.68, 215.77] });
  tx(s, '82%', 7.56, 4.38, 0.61, 0.32, { fontSize: 12, bold: true, color: C.white, lineSpacingMultiple: 1.2 });
  tx(s, '12%', 6.35, 2.74, 0.61, 0.32, { fontSize: 12, bold: true, color: C.white, lineSpacingMultiple: 1.2 });
  tx(s, '14%', 5.56, 2.98, 0.61, 0.32, { fontSize: 12, bold: true, color: C.white, lineSpacingMultiple: 1.2 });
  tx(s, '32%', 5.34, 4.38, 0.61, 0.32, { fontSize: 12, bold: true, color: C.white, lineSpacingMultiple: 1.2 });
  legend(s, 4.88, 6.36, 1.01, [[C.sand, 'Chart 1'], [C.gold, 'Chart 2'], [C.mist, 'Chart 3'], [C.navy, 'Chart 4']], { sw: 0.1, tw: 0.9, ty: -0.12 });
  pageNum(s, 23);
}

function slide24(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 29']);
  sh(s, 'pie', 0.3, 0.84, 6.2, 6.2, { fill: C.navy, angleRange: [216.39, 253.8] });
  sh(s, 'pie', 1.15, 1.55, 4.66, 4.66, { fill: C.mist, angleRange: [254.18, 291.25] });
  sh(s, 'pie', 1.64, 2.31, 3.5, 3.5, { fill: C.gold, angleRange: [91.75, 215.82] });
  sh(s, 'pie', 0.74, 0.92, 5.56, 6.12, { fill: C.taupe, angleRange: [292.29, 21.66] });
  sh(s, 'pie', 1.17, 1.51, 4.66, 5.12, { fill: C.sand, angleRange: [21.67, 91.5] });
  tx(s, '15%', 1.79, 1.89, 0.84, 0.51, { fontSize: 24, fontFace: F.bold, bold: true, color: C.white, align: 'center', wrap: false });
  tx(s, '10%', 3.2, 1.9, 0.73, 0.44, { fontSize: 20, fontFace: F.bold, bold: true, color: C.white, align: 'center', wrap: false });
  tx(s, '30%', 4.57, 2.99, 0.92, 0.51, { fontSize: 24, fontFace: F.bold, bold: true, color: C.white, align: 'center', wrap: false });
  tx(s, '20%', 3.91, 5.09, 0.91, 0.51, { fontSize: 24, fontFace: F.bold, bold: true, color: C.white, align: 'center', wrap: false });
  tx(s, '25%', 2.19, 4.09, 0.79, 0.44, { fontSize: 20, fontFace: F.bold, bold: true, color: C.white, align: 'center', wrap: false });
  title(s, 7.32, 1.22, 3.63, 1.45);
  tx(s, L3, 7.32, 5.65, 5.21, 0.9, { lineSpacingMultiple: 1.5 });
  tx(s, 'Chart Data Slide', 7.32, 5.07, 2.31, 0.4, { fontSize: 18, fontFace: F.bold, bold: true, color: C.dark });
  tx(s, '112M', 7.31, 4.2, 2.15, 0.84, { fontSize: 44, fontFace: F.bold, bold: true, color: C.taupe });
  tx(s, [{ text: L5, options: {  } }, { text: '1500s ', options: {  } }, { text: 'ing', options: {  } }, { text: ' and typesetting', options: {  } }], 7.31, 2.94, 5.35, 0.9, { lineSpacingMultiple: 1.5 });
  pageNum(s, 24);
}

function slide25(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 30']);
  sh(s, 'bentConnector3', 7.75, 5, 1.35, 0.62);
  sh(s, 'bentConnector3', 6.44, 2.46, 2.66, 0.67, { line: { color: 'B78D6D', width: 1, dashType: 'dash' }, flipV: true });
  sh(s, 'bentConnector3', 4.23, 3.57, 1.5, 0.63, { line: { color: 'C4A287', width: 1, dashType: 'dash' }, rotate: 180, flipV: true });
  sh(s, 'bentConnector3', 4.23, 4.86, 1.33, 0.77, { line: { color: 'D1B6A1', width: 1, dashType: 'dash' }, rotate: 180, flipV: true });
  statCard(s, 9.5, 5.05, 3.16, { ring: C.navy, pct: '53%', head: 'Your Text Here', headColor: C.navy, body: 'Lorem ipsum dolor sit amet, ', bodyW: 1.73, bodyBold: true });
  statCard(s, 9.5, 1.89, 3.16, { ring: 'B78D6D', pct: '53%', head: 'Your Text Here', headColor: C.ink, body: 'Lorem ipsum dolor sit amet, ', bodyW: 1.73, bodyBold: true });
  statCard(s, 0.68, 3.57, 3.16, { ring: 'B78D6D', pct: '53%', head: 'Your Text Here', headColor: C.ink, body: 'Lorem ipsum dolor sit amet, ', bodyW: 1.73, bodyBold: true });
  statCard(s, 0.68, 5.05, 3.16, { ring: 'B78D6D', pct: '53%', head: 'Your Text Here', headColor: C.ink, body: 'Lorem ipsum dolor sit amet, ', bodyW: 1.73, bodyBold: true });
  title(s, 0.68, 1.21, 3.63, 1.45);
  sh(s, 'pie', 5.27, 2.95, 2.81, 2.81, { fill: C.navy, angleRange: [272.67, 101.39] });
  sh(s, 'pie', 5.27, 2.95, 2.81, 2.81, { fill: 'B78D6D', angleRange: [236.98, 270] });
  sh(s, 'pie', 5.27, 2.94, 2.81, 2.81, { fill: 'C4A287', angleRange: [192.68, 233.68] });
  sh(s, 'pie', 5.27, 2.94, 2.81, 2.81, { fill: 'D1B6A1', angleRange: [102.78, 189.85] });
  sh(s, 'ellipse', 5.9, 3.56, 1.57, 1.57, { fill: C.white });
  legend(s, 4.98, 6.36, 0.92, [[C.navy, 'Data 1'], ['D1B6A1', 'Data 2'], ['B78D6D', 'Data 3'], ['C4A287', 'Data 4']], { sw: 0.1, tw: 0.9, ty: -0.12 });
  pageNum(s, 25);
}

function slide26(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 31']);
  sh(s, 'bentConnector3', 6.41, 3.75, 2.08, 1.51, { line: { color: C.mist, width: 1, dashType: 'dash' }, rotate: 270, flipV: true });
  sh(s, 'bentConnector3', 7.59, 2.46, 1.44, 0.56, { rotate: 270, flipV: true });
  sh(s, 'bentConnector3', 10.14, 2.47, 1.44, 0.54, { line: { color: C.mist, width: 1, dashType: 'dash' }, rotate: 90, flipH: true, flipV: true });
  sh(s, 'bentConnector3', 10.61, 3.71, 1.93, 1.43, { line: { color: C.sand, width: 1, dashType: 'dash' }, rotate: 90, flipH: true, flipV: true });
  tx(s, 'Chart Data', 6.21, 2.9, 1.33, 0.3, { fontSize: 12, fontFace: F.bold, bold: true, color: C.dark });
  tx(s, '35%', 6.21, 2.28, 1.33, 0.64, { fontSize: 32, fontFace: F.bold, bold: true, color: C.mist });
  tx(s, 'Chart Data', 7.53, 1.51, 1.33, 0.3, { fontSize: 12, fontFace: F.bold, bold: true, color: C.dark });
  tx(s, '35%', 7.53, 0.9, 1.33, 0.64, { fontSize: 32, fontFace: F.bold, bold: true, color: C.navy });
  tx(s, 'Chart Data', 10.47, 1.51, 1.33, 0.3, { fontSize: 12, fontFace: F.bold, bold: true, color: C.dark });
  tx(s, '35%', 10.47, 0.9, 1.33, 0.64, { fontSize: 32, fontFace: F.bold, bold: true, color: C.gold });
  tx(s, 'Chart Data', 11.71, 2.9, 1.33, 0.3, { fontSize: 12, fontFace: F.bold, bold: true, color: C.dark });
  tx(s, '35%', 11.71, 2.28, 1.33, 0.64, { fontSize: 32, fontFace: F.bold, bold: true, color: C.sand });
  title(s, 0.82, 1.72, 3.63, 1.45);
  tx(s, [{ text: L5, options: {  } }, { text: 'PLACEHOLDER', options: {  } }], 0.82, 3.63, 4.43, 1.46, { lineSpacingMultiple: 1.5 });
  tx(s, L1, 0.82, 4.87, 4.43, 0.9, { lineSpacingMultiple: 1.5 });
  sh(s, 'pie', 7.52, 2.4, 3.79, 3.8, { fill: C.sand, angleRange: [352.12, 59.33] });
  sh(s, 'pie', 7.52, 2.4, 3.79, 3.8, { fill: C.gold, angleRange: [289.85, 352.6] });
  sh(s, 'pie', 7.52, 2.39, 3.79, 3.8, { fill: C.navy, angleRange: [153.28, 290.53] });
  sh(s, 'pie', 7.52, 2.39, 3.79, 3.8, { fill: C.mist, angleRange: [58.47, 155.02] });
  sh(s, 'ellipse', 8.32, 3.16, 2.25, 2.26, { fill: C.white });
  pageNum(s, 26);
}

function slide27(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 28']);
  sh(s, 'ellipse', 3.92, 1.03, 5.5, 5.5, { fill: C.white, shadow: SH.soft });
  sh(s, 'pie', 4.12, 1.21, 5.1, 5.1, { fill: C.navy, angleRange: [272.67, 56.69] });
  sh(s, 'pie', 4.12, 1.21, 5.1, 5.1, { fill: C.mist, angleRange: [188.54, 270] });
  sh(s, 'pie', 4.12, 1.19, 5.1, 5.1, { fill: C.sand, angleRange: [156.89, 186.55] });
  sh(s, 'pie', 4.12, 1.19, 5.1, 5.1, { fill: C.gold, angleRange: [58.47, 155.02] });
  sh(s, 'ellipse', 5.32, 2.37, 2.76, 2.76, { fill: C.white });
  sh(s, 'wedgeEllipseCallout', 2.78, 0.98, 2.01, 2.01, { fill: '3E5566' });
  tx(s, '120', 2.97, 1.51, 1.62, 0.71, { fontSize: 36, fontFace: F.bold, bold: true, color: C.white, align: 'center' });
  tx(s, 'Sales !', 2.97, 2.09, 1.62, 0.37, { fontSize: 16, color: C.white, align: 'center' });
  sh(s, 'roundRect', 2.49, 1.03, 0.64, 0.64, { fill: C.white, rectRadius: 0.26, shadow: SH.lift });
  tx(s, 'Lorem ipsum dolor sit amet. Qui sint neque', 5.79, 3.21, 1.76, 1.12, { fontSize: 14, color: '262626', align: 'center', lineSpacingMultiple: 1.5 });
  title(s, 0.82, 4.78, 3.63, 1.45);
  tx(s, L3, 9.74, 1.35, 2.77, 1.46, { lineSpacingMultiple: 1.5 });
  tx(s, 'Subtitle Here', 9.74, 4.78, 1.74, 0.34, { fontSize: 14, fontFace: F.bold, color: '262626' });
  tx(s, L3, 9.74, 5.09, 2.77, 1.46, { lineSpacingMultiple: 1.5 });
  icon(s, 'book', 2.69, 1.22, 0.24, 0.26, C.gold);
  pageNum(s, 27);
}

function slide28(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 8']);
  sh(s, 'ellipse', 1.81, 3.26, 2.95, 2.95, { fill: C.white, shadow: SH.card });
  sh(s, 'ellipse', 1.05, 1.18, 1.92, 1.92, { fill: C.white, shadow: SH.card });
  sh(s, 'ellipse', 3.59, 1.18, 1.92, 1.92, { fill: C.white, shadow: SH.card });
  title(s, 6.67, 1.14, 3.63, 1.45);
  tx(s, 'Subtitle Here', 6.67, 2.96, 1.74, 0.34, { fontSize: 14, fontFace: F.bold, color: '262626' });
  tx(s, L3, 6.67, 3.27, 5.63, 0.63, { lineSpacingMultiple: 1.5 });
  tx(s, 'Subtitle Here', 6.67, 4.2, 1.74, 0.34, { fontSize: 14, fontFace: F.bold, color: '262626' });
  tx(s, L3, 6.67, 4.5, 5.63, 0.63, { lineSpacingMultiple: 1.5 });
  tx(s, 'Subtitle Here', 6.67, 5.43, 1.74, 0.34, { fontSize: 14, fontFace: F.bold, color: '262626' });
  tx(s, L3, 6.67, 5.74, 5.63, 0.63, { lineSpacingMultiple: 1.5 });
  sh(s, 'pie', 3.72, 1.3, 1.68, 1.68, { fill: C.sand, angleRange: [269.93, 58.46] });
  sh(s, 'pie', 3.72, 1.3, 1.68, 1.68, { fill: 'CFCFCF', angleRange: [186.18, 270] });
  sh(s, 'pie', 3.72, 1.29, 1.68, 1.68, { fill: 'B0B0B0', angleRange: [154.05, 186.55] });
  sh(s, 'pie', 3.72, 1.29, 1.68, 1.68, { fill: '8A8A8A', angleRange: [58.47, 155.02] });
  sh(s, 'ellipse', 4.1, 1.66, 0.94, 0.94, { fill: C.white });
  tx(s, 'chart  3', 4.23, 1.91, 0.76, 0.39, { fontSize: 14, fontFace: F.bold, color: '262626', wrap: false });
  tx(s, '40%', 4.95, 2.07, 0.58, 0.27, { fontSize: 10, fontFace: F.bold, color: C.white });
  tx(s, '20%', 3.88, 1.49, 0.58, 0.27, { fontSize: 10, fontFace: F.bold, color: C.white });
  tx(s, '10%', 3.66, 2.07, 0.58, 0.27, { fontSize: 10, fontFace: F.bold, color: C.white });
  tx(s, [{ text: '3', options: { fontSize: 10, fontFace: F.bold, color: C.white } }, { text: '0%', options: { fontSize: 10, fontFace: F.bold, color: C.white } }], 4.17, 2.61, 0.58, 0.27, { fontSize: 10, fontFace: F.bold, color: C.white });
  sh(s, 'pie', 1.17, 1.3, 1.68, 1.68, { fill: C.gold, angleRange: [269.93, 58.46] });
  sh(s, 'pie', 1.17, 1.3, 1.68, 1.68, { fill: 'CFCFCF', angleRange: [186.18, 270] });
  sh(s, 'pie', 1.17, 1.29, 1.68, 1.68, { fill: 'B0B0B0', angleRange: [154.05, 186.55] });
  sh(s, 'pie', 1.17, 1.29, 1.68, 1.68, { fill: '8A8A8A', angleRange: [58.47, 155.02] });
  sh(s, 'ellipse', 1.53, 1.65, 0.97, 0.97, { fill: C.white });
  tx(s, 'chart  1', 1.66, 1.88, 0.72, 0.39, { fontSize: 14, fontFace: F.bold, color: '262626', wrap: false });
  tx(s, '40%', 2.4, 2.07, 0.58, 0.27, { fontSize: 10, fontFace: F.bold, color: C.white });
  tx(s, '20%', 1.33, 1.49, 0.58, 0.27, { fontSize: 10, fontFace: F.bold, color: C.white });
  tx(s, '10%', 1.15, 2.07, 0.58, 0.27, { fontSize: 10, fontFace: F.bold, color: C.white });
  tx(s, [{ text: '3', options: { fontSize: 10, fontFace: F.bold, color: C.white } }, { text: '0%', options: { fontSize: 10, fontFace: F.bold, color: C.white } }], 1.62, 2.61, 0.58, 0.27, { fontSize: 10, fontFace: F.bold, color: C.white });
  sh(s, 'pie', 2.01, 3.44, 2.61, 2.61, { fill: C.navy, angleRange: [269.93, 58.46] });
  sh(s, 'pie', 2.01, 3.44, 2.61, 2.61, { fill: 'CFCFCF', angleRange: [186.18, 270] });
  sh(s, 'pie', 2.01, 3.43, 2.61, 2.61, { fill: 'B0B0B0', angleRange: [154.05, 186.55] });
  sh(s, 'pie', 2.01, 3.43, 2.61, 2.61, { fill: '8A8A8A', angleRange: [58.47, 155.02] });
  sh(s, 'ellipse', 2.58, 3.98, 1.51, 1.51, { fill: C.white });
  tx(s, 'chart  2', 2.99, 4.46, 0.76, 0.39, { fontSize: 14, fontFace: F.bold, color: '262626', wrap: false });
  tx(s, '40%', 4.04, 4.51, 0.75, 0.34, { fontSize: 14, fontFace: F.bold, color: C.white });
  tx(s, '20%', 2.35, 3.8, 0.75, 0.34, { fontSize: 14, fontFace: F.bold, color: C.white });
  tx(s, '10%', 2.03, 4.68, 0.75, 0.34, { fontSize: 14, fontFace: F.bold, color: C.white });
  tx(s, '30%', 2.72, 5.52, 0.75, 0.34, { fontSize: 14, fontFace: F.bold, color: C.white });
  pageNum(s, 28);
}

function slide29(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 18']);
  tx(s, '2023', 5.57, 3.86, 1.93, 0.64, { fontSize: 32, fontFace: F.bold, bold: true, color: '262626', align: 'center' });
  sh(s, 'bentConnector3', 7.79, 4.69, 1.8, 0.98);
  sh(s, 'bentConnector3', 6.26, 2.69, 3.33, 0.85, { line: { color: C.mist, width: 1, dashType: 'dash' } });
  sh(s, 'bentConnector3', 3.48, 5, 1.75, 0.87, { line: { color: C.gold, width: 1, dashType: 'dash' }, rotate: 180, flipV: true });
  sh(s, 'bentConnector3', 3.48, 3.08, 1.44, 0.67, { line: { color: C.sand, width: 1, dashType: 'dash' }, rotate: 180, flipV: true });
  tx(s, L9, 9.86, 5.79, 2.75, 0.9, { lineSpacingMultiple: 1.5 });
  tx(s, 'Chart Data Slide', 9.86, 5.54, 2.31, 0.3, { fontSize: 12, fontFace: F.bold, bold: true, color: C.dark });
  tx(s, '112M', 9.86, 4.94, 2.15, 0.64, { fontSize: 32, fontFace: F.bold, bold: true, color: C.navy });
  tx(s, L9, 9.86, 3.65, 2.75, 0.9, { lineSpacingMultiple: 1.5 });
  tx(s, 'Chart Data Slide', 9.86, 3.41, 2.31, 0.3, { fontSize: 12, fontFace: F.bold, bold: true, color: C.dark });
  tx(s, '112M', 9.86, 2.81, 2.15, 0.64, { fontSize: 32, fontFace: F.bold, bold: true, color: C.mist });
  tx(s, L10, 0.73, 5.79, 2.55, 0.9, { align: 'right', lineSpacingMultiple: 1.5 });
  tx(s, 'Chart Data Slide', 0.97, 5.54, 2.31, 0.3, { fontSize: 12, fontFace: F.bold, bold: true, color: C.dark, align: 'right' });
  tx(s, '112M', 1.13, 4.94, 2.15, 0.64, { fontSize: 32, fontFace: F.bold, bold: true, color: C.gold, align: 'right' });
  tx(s, L10, 0.73, 3.65, 2.55, 0.9, { align: 'right', lineSpacingMultiple: 1.5 });
  tx(s, 'Chart Data Slide', 0.97, 3.41, 2.31, 0.3, { fontSize: 12, fontFace: F.bold, bold: true, color: C.dark, align: 'right' });
  tx(s, '112M', 1.13, 2.81, 2.15, 0.64, { fontSize: 32, fontFace: F.bold, bold: true, color: C.sand, align: 'right' });
  title(s, 3.2, 1.01, 6.92, 0.77, { align: 'center' });
  sh(s, 'pie', 4.86, 2.43, 3.35, 3.35, { fill: C.navy, angleRange: [272.67, 98] });
  sh(s, 'pie', 4.86, 2.37, 3.35, 3.35, { fill: C.mist, angleRange: [234.79, 270] });
  sh(s, 'pie', 4.71, 2.27, 3.35, 3.35, { fill: C.sand, angleRange: [197.52, 229.37] });
  sh(s, 'pie', 4.86, 2.47, 3.35, 3.35, { fill: C.gold, angleRange: [103.39, 198.74] });
  sh(s, 'ellipse', 5.21, 2.75, 2.69, 2.69, { fill: C.white });
  legend(s, 4.74, 6.36, 0.99, [[C.navy, 'May '], [C.gold, 'Juni '], [C.sand, 'July '], [C.mist, 'Agustus ']], { sw: 0.1, tw: 0.9, ty: -0.12 });
  pageNum(s, 29);
}

function slide30(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Option 29']);
  sh(s, 'arc', 6.72, 1.71, 4.08, 4.08, { line: { color: C.navy, width: 25 }, angleRange: [270.61, 90.3] });
  sh(s, 'arc', 6.72, 1.71, 4.08, 4.08, { line: { color: C.sand, width: 25 }, angleRange: [90.27, 270.03] });
  sh(s, 'ellipse', 5.17, 1.29, 2.11, 2.11, { fill: C.white, shadow: SH.lift });
  sh(s, 'ellipse', 10.24, 4.1, 2.11, 2.11, { fill: C.white, shadow: SH.lift });
  title(s, 0.82, 1.46, 3.63, 1.45);
  tx(s, 'Subtitle Here', 0.82, 3.32, 1.74, 0.34, { fontSize: 14, fontFace: F.bold, color: '262626' });
  tx(s, [{ text: L11, options: {  } }, { text: 'standard simply dummy text of the ', options: {  } }], 0.82, 3.62, 4.08, 0.9, { lineSpacingMultiple: 1.5 });
  tx(s, 'Subtitle Here', 0.82, 4.83, 1.74, 0.34, { fontSize: 14, fontFace: F.bold, color: '262626' });
  tx(s, [{ text: L11, options: {  } }, { text: 'standard simply dummy text of the ', options: {  } }], 0.82, 5.13, 4.08, 0.9, { lineSpacingMultiple: 1.5 });
  icon(s, 'shredder', 5.84, 1.95, 0.78, 0.78, C.sand);
  icon(s, 'folder', 10.9, 4.81, 0.78, 0.7, C.navy);
  tx(s, '2023', 7.6, 3.32, 2.32, 0.91, { fontSize: 48, fontFace: F.bold, bold: true, color: '262626', align: 'center' });
  pageNum(s, 30);
}

function slide31(p) {
  const s = p.addSlide();
  s.background = { color: C.paper };
  deco(s, DECO['Cover']);
  sh(s, 'ellipse', 1.18, 1.35, 4.92, 4.92, { fill: C.mist });
  sh(s, 'ellipse', 1.5, 1.67, 4.29, 4.29, { fill: C.sand });
  sh(s, 'ellipse', 1.86, 2.03, 3.56, 3.56, { fill: C.gold });
  sh(s, 'ellipse', 2.21, 2.4, 2.87, 2.87, { fill: C.white });
  sh(s, 'roundRect', 4.47, 4.72, 2.61, 1.19, { fill: C.navy });
  tx(s, '3,456', 4.47, 4.87, 2.6, 0.7, { fontSize: 36, fontFace: F.bold, bold: true, color: C.white, align: 'center', valign: 'middle', lineSpacingMultiple: 0.9 });
  tx(s, 'Production', 4.46, 5.44, 2.61, 0.34, { fontSize: 14, color: C.white, align: 'center' });
  title(s, 7.49, 1.17, 3.63, 1.45);
  sh(s, 'donut', 7.6, 4.45, 0.8, 0.8, { fill: C.haze, rectRadius: 0.11 });
  sh(s, 'blockArc', 7.6, 4.45, 0.8, 0.8, { fill: C.gold, angleRange: [50.42, 1.51], arcThicknessRatio: 0.26 });
  tx(s, '53%', 7.71, 4.7, 0.59, 0.3, { fontSize: 12, fontFace: F.bold, bold: true, color: '000000', align: 'center' });
  tx(s, 'Your Text Here', 8.52, 4.58, 1.19, 0.57, { fontSize: 14, fontFace: F.bold, color: C.ink });
  sh(s, 'donut', 10.21, 4.45, 0.8, 0.8, { fill: C.haze, rectRadius: 0.11 });
  sh(s, 'blockArc', 10.21, 4.45, 0.8, 0.8, { fill: C.navy, angleRange: [50.42, 1.51], arcThicknessRatio: 0.26 });
  tx(s, '53%', 10.32, 4.7, 0.59, 0.3, { fontSize: 12, fontFace: F.bold, bold: true, color: '000000', align: 'center' });
  tx(s, 'Your Text Here', 11.13, 4.58, 1.19, 0.57, { fontSize: 14, fontFace: F.bold, color: C.ink });
  sh(s, 'donut', 7.6, 5.53, 0.8, 0.8, { fill: C.haze, rectRadius: 0.11 });
  sh(s, 'blockArc', 7.6, 5.53, 0.8, 0.8, { fill: C.sand, angleRange: [50.42, 1.51], arcThicknessRatio: 0.26 });
  tx(s, '53%', 7.71, 5.78, 0.59, 0.3, { fontSize: 12, fontFace: F.bold, bold: true, color: '000000', align: 'center' });
  tx(s, 'Your Text Here', 8.52, 5.67, 1.19, 0.57, { fontSize: 14, fontFace: F.bold, color: C.ink });
  sh(s, 'donut', 10.21, 5.53, 0.8, 0.8, { fill: C.haze, rectRadius: 0.11 });
  sh(s, 'blockArc', 10.21, 5.53, 0.8, 0.8, { fill: C.mist, angleRange: [50.42, 1.51], arcThicknessRatio: 0.26 });
  tx(s, '53%', 10.32, 5.78, 0.59, 0.3, { fontSize: 12, fontFace: F.bold, bold: true, color: '000000', align: 'center' });
  tx(s, 'Your Text Here', 11.13, 5.67, 1.19, 0.57, { fontSize: 14, fontFace: F.bold, color: C.ink });
  tx(s, [{ text: L5, options: {  } }, { text: '1500s the printing and typesetting', options: {  } }], 7.49, 2.96, 4.92, 1.18, { lineSpacingMultiple: 1.5 });
  icon(s, 'folderPie', 3.28, 3.42, 0.73, 0.66, C.navy);
  pageNum(s, 31);
}

// Slide 19 embeds a real pie chart; recreated with the native chart API.
function pieChart(s, x, y, w, h) {
  s.addChart('pie', [{ name: 'Sales', labels: ['Data 1', 'Data 2', 'Data 3'], values: [8.2, 3.2, 5.4] }],
    { x, y, w, h, chartColors: [C.navy, C.gold, C.sand], showLegend: true, legendPos: 'b',
      legendFontFace: F.light, legendFontSize: 12, legendColor: '595959', dataBorder: { pt: 0, color: C.white } });
}

// Master furniture: page number flanked by two dots, bottom centre.
function pageNum(s, n) {
  sh(s, 'ellipse', 6.3, 7.05, 0.09, 0.09, { fill: 'DDE3E9' });
  sh(s, 'ellipse', 6.94, 7.05, 0.09, 0.09, { fill: 'DDE3E9' });
  tx(s, String(n), 6.36, 6.91, 0.62, 0.37, { fontFace: F.bold, fontSize: 16, align: 'center', color: C.navy });
}

// ---------------------------------------------------------------- build
function build() {
  const p = new PptxGenJS();
  p.defineLayout({ name: 'WIDE', width: 13.3333, height: 7.5 });
  p.layout = 'WIDE';
  p.author = 'pptxgenjs';
  p.title = 'Pie Chart Infographic Presentation';
  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30, slide31]
    .forEach((fn) => fn(p));
  return p.writeFile({ fileName: path.join(__dirname, '0d1463f1-e119-4994-95ad-67d1d0deea57_grok_final.pptx') });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
