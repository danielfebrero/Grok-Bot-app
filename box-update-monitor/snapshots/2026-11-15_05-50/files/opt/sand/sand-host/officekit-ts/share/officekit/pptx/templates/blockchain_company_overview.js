/**
 * "Blockchain Technology" — 27-slide deck rebuilt with pptxgenjs.
 * Run: node 13c71d4f-34cf-4656-b442-2c88a28aa6f0_grok_final.js
 */
'use strict';
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const C = {
  lime: 'B5E72D', limeMid: '8DB915', limeDark: '5E7B0E',
  white: 'F7F9FF', pure: 'FFFFFF', snow: 'FCFFFF',
  ink: '012525', deep: '0B2525', moss: '1C3425', olive: '3D5728',
  ph: 'FFFFFF',
  // pptxgenjs has no gradient fill, so the deck's two-stop arrow gradients
  // are flattened to their midpoints.
  gInkOlive: '1F3E26', gInkMoss: '0E2C25', gOliveMoss: '2C4526'
};
const F = { med: 'Space Grotesk Medium', reg: 'Space Grotesk', light: 'Space Grotesk Light' };

const W = 10, H = 5.625;               // slide size (inches)
const fill = (color, opacity) => (opacity === undefined ? { color } : { color, transparency: 100 - opacity });
const stroke = (color, opacity, width) => ({ color, transparency: 100 - (opacity === undefined ? 100 : opacity), width: width || 0.75 });
const dashed = (opacity) => ({ color: C.pure, transparency: 100 - opacity, width: 0.75, dashType: 'dash' });
const NOLINE = { type: 'none' };

/* ------------------------------------------------------------------ helpers */

// Corner gradient of the two masters, approximated with 45-degree bands.
// Default runs olive (top-left) -> deep teal (bottom-right); `flip` reverses it
// for the "Side with logo C" layout used from slide 16 onwards.
function background(s, flip) {
  const mix = (a, b, t) => [0, 2, 4].map(i =>
    Math.round(parseInt(a.substr(i, 2), 16) * (1 - t) + parseInt(b.substr(i, 2), 16) * t)
      .toString(16).padStart(2, '0')).join('').toUpperCase();
  const BANDS = 26, span = (W + H) * Math.SQRT1_2, band = span / BANDS;
  for (let i = 0; i < BANDS; i++) {
    const d = (i + 0.5) * band;                       // distance along the diagonal
    const t = flip ? d / span : 1 - d / span;
    s.addShape('rect', {
      x: Math.SQRT1_2 * d + (W - H) / 4 - band / 2 - 0.01,
      y: Math.SQRT1_2 * d - (W - H) / 4 - H * 2.5,
      w: band + 0.02, h: H * 5, rotate: 45,
      fill: fill(mix(C.deep, C.olive, t)), line: NOLINE
    });
  }
}
const bgFlip = s => background(s, true);

// "Blockchain." wordmark + accent tick + slide number (from the shared layouts)
function footer(s, num) {
  s.addText([{ text: 'Blockchain', options: { color: C.white } }, { text: '.', options: { color: C.lime } }],
    { x: 0.633, y: 5.062, w: 0.9, h: 0.23, fontFace: F.reg, fontSize: 9, margin: 1.5, valign: 'middle' });
  s.addShape('line', { x: 8.99, y: 5.091, w: 0, h: 0.203, line: { color: C.lime, width: 0.75 } });
  s.addText(String(num), { x: 9.048, y: 5.071, w: 0.402, h: 0.24, fontFace: F.light, fontSize: 11, color: C.white, align: 'center', valign: 'middle', margin: 0 });
}

// Pointy-top hexagon (the deck's signature shape).
const HEX_PTS = [[0.5, 0], [1, 0.25], [1, 0.75], [0.5, 1], [0, 0.75], [0, 0.25]];
function hex(s, x, y, w, h, opts) {
  s.addShape('custGeom', Object.assign({
    x, y, w, h, points: HEX_PTS.map(p => ({ x: p[0] * w, y: p[1] * h })).concat([{ close: true }])
  }, opts));
}
const HEX_RATIO = 1.1549;              // h / w for every hexagon in the deck

// Decorative background clusters: [x, y, w, outline%, fill%, limeTinted]
function hexCluster(s, list) {
  list.forEach(([x, y, w, ol, fl, limy]) => {
    const h = w * HEX_RATIO, ix = x + w * 0.0475, iy = y + h * 0.0475;
    hex(s, x, y, w, h, { fill: NOLINE, line: stroke(limy ? C.lime : C.pure, ol) });
    hex(s, ix, iy, w * 0.905, h * 0.905, { fill: fill(limy ? C.lime : C.snow, fl), line: NOLINE });
  });
}
const CLUSTER_RIGHT = [   // slides 1 & 12
  [5.631, 3.559, 2.002, 9, 3], [7.294, 3.178, 0.749, 24, 7], [7.724, 3.922, 0.749, 20, 5],
  [8.151, 2.516, 1.514, 36, 11, 1], [7.327, 1.131, 1.514, 31, 11], [7.650, -1.391, 2.507, 7, 2],
  [5.732, -0.665, 1.032, 20, 5], [9.075, 1.425, 0.749, 20, 5], [6.781, 1.027, 0.405, 33, 14, 1],
  [4.963, -0.294, 0.405, 20, 5], [8.744, 4.821, 0.328, 10, 2]
];
const CLUSTER_LEFT = [    // slide 11
  [-0.400, 3.278, 2.002, 9, 3], [0.627, 2.532, 0.749, 24, 7], [1.724, 4.349, 0.749, 20, 5],
  [3.265, 2.725, 1.514, 36, 11, 1], [1.473, 1.131, 1.514, 31, 11], [0.108, -1.391, 2.507, 7, 2],
  [2.698, -0.665, 1.032, 20, 5], [1.273, 5.102, 0.749, 20, 5], [4.281, 0.512, 0.405, 33, 14, 1],
  [5.650, -0.138, 0.405, 20, 5], [3.608, 4.852, 0.328, 10, 2]
];
const CLUSTER_CONTACT = [ // slide 27
  [0.108, -1.391, 2.507, 7, 2], [-0.400, 3.278, 2.002, 9, 3], [-0.380, 1.627, 0.749, 24, 7],
  [1.724, 4.349, 0.749, 20, 5], [2.157, 4.885, 1.514, 36, 11, 1], [2.682, 0.183, 0.687, 20, 5],
  [1.273, 5.102, 0.749, 20, 5], [4.157, 0.851, 0.405, 33, 14, 1], [5.650, -0.138, 0.405, 20, 5],
  [3.608, 4.852, 0.328, 10, 2]
];
const CLUSTER_PYRAMID = [ // slide 25
  [8.064, 0.720, 2.002, 9, 3], [7.644, 0.365, 0.749, 24, 7], [9.082, 2.904, 0.749, 20, 5],
  [6.985, -0.161, 0.405, 33, 14, 1], [7.878, 5.139, 0.749, 20, 5]
];

// Isometric green cube: top diamond + two side faces.
function cube(s, x, y, w, h) {
  const poly = (pts, f) => s.addShape('custGeom', {
    x, y, w, h, fill: f, line: NOLINE,
    points: pts.map(p => ({ x: p[0] * w, y: p[1] * h })).concat([{ close: true }])
  });
  poly([[0.5, 0], [1, 0.218], [0.5, 0.436], [0, 0.218]], fill(C.lime));
  poly([[0, 0.218], [0.5, 0.436], [0.5, 1], [0, 0.782]], fill(C.limeMid, 60));
  poly([[1, 0.218], [0.5, 0.436], [0.5, 1], [1, 0.782]], fill(C.limeDark, 40));
}
// The cube always sits inside a dark hexagon of width `hw`.
function hexCube(s, x, y, hw, hexFill) {
  const hh = hw * HEX_RATIO;
  hex(s, x, y, hw, hh, { fill: fill(hexFill || C.ink), line: NOLINE });
  cube(s, x + hw * 0.162, y + hh * 0.147, hw * 0.735, hh * 0.705);
}

// Pointed banner / chevron used for the big translucent arrows.
function chevron(s, x, y, w, h, tip, opts) {
  s.addShape('custGeom', Object.assign({
    x, y, w, h,
    points: [{ x: 0, y: 0 }, { x: w - tip, y: 0 }, { x: w, y: h / 2 }, { x: w - tip, y: h }, { x: 0, y: h }, { close: true }]
  }, opts));
}

// Rounded rectangle with rounded rectangular windows cut out of it (even-odd fill).
// `holes` are [x, y, w, h] in local coordinates.
function panelWithHoles(s, x, y, w, h, r, holes, opts) {
  const rr = (X, Y, W2, H2, R) => [
    { x: X + R, y: Y, moveTo: true }, { x: X + W2 - R, y: Y },
    { x: X + W2, y: Y + R, curve: { type: 'quadratic', x1: X + W2, y1: Y } },
    { x: X + W2, y: Y + H2 - R },
    { x: X + W2 - R, y: Y + H2, curve: { type: 'quadratic', x1: X + W2, y1: Y + H2 } },
    { x: X + R, y: Y + H2 },
    { x: X, y: Y + H2 - R, curve: { type: 'quadratic', x1: X, y1: Y + H2 } },
    { x: X, y: Y + R },
    { x: X + R, y: Y, curve: { type: 'quadratic', x1: X, y1: Y } }, { close: true }
  ];
  let pts = rr(0, 0, w, h, r);
  holes.forEach(([hx, hy, hw, hh, hr]) => { pts = pts.concat(rr(hx, hy, hw, hh, hr === undefined ? r * 0.6 : hr)); });
  s.addShape('custGeom', Object.assign({ x, y, w, h, points: pts }, opts));
}

/* -------------------------------------------------------------- icon glyphs */
// Compact stand-ins for the deck's line-art icons, drawn from plain polygons.
const GLYPHS = {
  shield: [[0.5, 0], [1, 0.18], [1, 0.6], [0.5, 1], [0, 0.6], [0, 0.18]],
  check: [[0.42, 1], [0, 0.6], [0.15, 0.44], [0.42, 0.7], [0.87, 0.12], [1, 0.28]],
  doc: [[0.1, 0], [0.75, 0], [0.95, 0.2], [0.95, 1], [0.1, 1]],
  bank: [[0.5, 0], [1, 0.3], [1, 0.42], [0, 0.42], [0, 0.3]],
  bolt: [[0.55, 0], [0.15, 0.55], [0.45, 0.55], [0.3, 1], [0.85, 0.42], [0.5, 0.42]],
  chip: [[0.15, 0.15], [0.85, 0.15], [0.85, 0.85], [0.15, 0.85]],
  clock: [[0.5, 0], [1, 0.35], [0.8, 1], [0.2, 1], [0, 0.35]],
  arrows: [[0, 0.35], [0.7, 0.35], [0.7, 0.15], [1, 0.5], [0.7, 0.85], [0.7, 0.65], [0, 0.65]]
};
function glyph(s, kind, x, y, w, h, color) {
  if (kind === 'lock') {
    s.addShape('roundRect', { x, y: y + h * 0.42, w, h: h * 0.58, rectRadius: Math.min(w, h) * 0.15, fill: fill(color), line: NOLINE });
    s.addShape('arc', { x: x + w * 0.18, y: y, w: w * 0.64, h: h * 0.62, angleRange: [180, 360], line: { color, width: Math.max(0.75, w * 8) } });
    return;
  }
  if (kind === 'search') {
    s.addShape('ellipse', { x, y, w: w * 0.78, h: h * 0.78, fill: NOLINE, line: { color, width: Math.max(0.75, w * 8) } });
    s.addShape('line', { x: x + w * 0.66, y: y + h * 0.66, w: w * 0.34, h: h * 0.34, line: { color, width: Math.max(1, w * 9) } });
    return;
  }
  if (kind === 'person') {
    s.addShape('ellipse', { x: x + w * 0.28, y, w: w * 0.44, h: h * 0.44, fill: fill(color), line: NOLINE });
    s.addShape('custGeom', {
      x, y: y + h * 0.5, w, h: h * 0.5, fill: fill(color), line: NOLINE,
      points: [{ x: 0, y: h * 0.5 }, { x: w * 0.1, y: h * 0.1 }, { x: w * 0.9, y: h * 0.1 }, { x: w, y: h * 0.5 }, { close: true }]
    });
    return;
  }
  if (kind === 'coin') {
    s.addShape('ellipse', { x, y, w, h, fill: fill(color), line: NOLINE });
    return;
  }
  if (kind === 'pin') {
    s.addShape('custGeom', {
      x, y, w, h, fill: fill(color), line: NOLINE,
      points: [{ x: w * 0.5, y: h }, { x: 0, y: h * 0.42 }, { x: w * 0.15, y: h * 0.1 },
      { x: w * 0.85, y: h * 0.1 }, { x: w, y: h * 0.42 }, { close: true }]
    });
    return;
  }
  const pts = GLYPHS[kind] || GLYPHS.chip;
  s.addShape('custGeom', {
    x, y, w, h, fill: fill(color), line: NOLINE,
    points: pts.map(p => ({ x: p[0] * w, y: p[1] * h })).concat([{ close: true }])
  });
}
// glyph centred inside a hexagon badge
function hexGlyph(s, kind, x, y, hw, badge, ink, gs) {
  const hh = hw * HEX_RATIO, g = gs || hw * 0.42;
  hex(s, x, y, hw, hh, { fill: fill(badge), line: NOLINE });
  glyph(s, kind, x + (hw - g) / 2, y + (hh - g) / 2, g, g, ink);
}

/* ------------------------------------------------------------------- text */
// `rich('Types of *Blockchains*')` — text between asterisks is lime accented.
function rich(str, base, accent) {
  return str.split('*').map((t, i) => ({ text: t, options: { color: i % 2 ? (accent || C.lime) : (base || C.white) } }))
    .filter(r => r.text.length);
}
function text(s, runs, o) {
  s.addText(runs, Object.assign({
    fontFace: F.light, fontSize: 9, color: C.white, margin: 1.5,
    lineSpacingMultiple: 1, valign: 'top'
  }, o));
}
// Standard slide heading (17pt medium, baseline-aligned at y=0.602)
function heading(s, str, o) {
  text(s, rich(str), Object.assign({ x: 0.637, y: 0.602, w: 8.7, h: 0.344, fontFace: F.med, fontSize: 17, valign: 'bottom' }, o));
}
// Standard intro paragraph under the heading
function intro(s, str, o) {
  text(s, rich(str), Object.assign({ x: 0.637, y: 1.077, w: 8.668, h: 0.5 }, o));
}
// Bulleted list using the deck's en-dash bullet
function bullets(s, items, o) {
  text(s, items.map(t => ({ text: t, options: { breakLine: true } })),
    Object.assign({ bullet: { characterCode: '002D', indent: 7 } }, o));
}
// Stack of plain paragraphs (no bullet)
function lines(s, items, o) {
  text(s, items.map(t => ({ text: t, options: { breakLine: true } })), o);
}

/* ==================================================================== SLIDES */

// 1 — cover
function slide01(s) {
  background(s);
  hexCluster(s, CLUSTER_RIGHT);
  s.addShape('rect', { x: 2.107, y: 3.504, w: 1.146, h: 0.026, fill: fill(C.lime, 25), line: NOLINE });
  // top navigation bar
  glyph(s, 'chip', 0.66, 0.657, 0.105, 0.093, C.pure);
  text(s, [{ text: 'Blockchain', options: { color: C.white } }, { text: '.', options: { color: C.lime } }],
    { x: 0.816, y: 0.545, w: 1.4, h: 0.3, fontFace: F.med, fontSize: 15 });
  [['About Us', 2.675], ['Service', 3.581], ['News', 4.487], ['Blockchain', 5.393], ['Contact Us', 6.299]]
    .forEach(([label, x], i) => text(s, label, { x, y: 0.605, w: 0.865, h: 0.24, fontSize: 10, align: 'center', color: i === 3 ? C.lime : C.white }));
  // headline block
  text(s, rich('*BLOCKCHAIN* TECHNOLOGY'), { x: 0.633, y: 1.7, w: 4.6, h: 1.4, fontFace: F.med, fontSize: 38 });
  text(s, '27 Unique Slides with Valuable Content', { x: 0.677, y: 3.24, w: 2.9, h: 0.24, fontSize: 10 });
  // call-to-action button
  s.addShape('roundRect', { x: 0.667, y: 4.045, w: 2.017, h: 0.46, rectRadius: 0.109, fill: fill(C.lime), line: NOLINE });
  s.addShape('roundRect', { x: 2.188, y: 4.045, w: 0.495, h: 0.46, rectRadius: 0.109, fill: fill(C.snow, 25), line: NOLINE });
  text(s, 'More Information', { x: 0.873, y: 4.155, w: 1.4, h: 0.24, fontFace: F.reg, fontSize: 10, color: C.ink });
  s.addShape('triangle', { x: 2.365, y: 4.2, w: 0.11, h: 0.15, rotate: 90, fill: fill(C.ink), line: NOLINE });
}

// 2 — What is a Blockchain?
function slide02(s) {
  background(s); footer(s, 2);
  heading(s, 'What is a *Blockchain*?', { y: 1.252, w: 4 });
  lines(s, [
    'Blockchain is a digital ledger technology that allows secure and transparent transactions. It is a decentralized system, meaning there is no central authority controlling it. Instead, all participants in the network have access to the same information and can validate transactions.',
    '',
    'The blockchain is essentially a database that stores information in a series of blocks, each block containing a record of multiple transactions. Each block is linked to the previous one, forming a chain of blocks, hence the name blockchain. This chain of blocks creates an unalterable and transparent record of all transactions, making it ideal for use in financial transactions, supply chain management, and other applications where transparency and security are critical.',
    '',
    'One of the key features of blockchain technology is its immutability, which means that once a transaction is recorded on the blockchain, it cannot be altered or deleted. This feature ensures the integrity of the data, making it highly secure.'
  ], { x: 0.637, y: 1.727, w: 5.026, h: 2.617 });
}

// 3 — About Us (team grid)
const TEAM = [
  ['John\nSmith', 'CEO'], ['Sarah\nJohnson', 'CTO'], ['David\nBrown', 'COO'],
  ['Elizabeth Davis', 'Head of Business Development'], ['Thomas\nWilson', 'Lead Engineer'], ['Lisa\nMiller', 'Head of Marketing'],
  ['Michael Anderson', 'Head of Product'], ['Jennifer Lee', 'Chief Financial Officer'], ['Christopher Adams', 'Head of Security'],
  ['Karen Thompson', 'Head of Operations'], ['Richard Wright', 'Lead Blockchain Developer'], ['Amanda James', 'Senior Business Analyst']
];
function slide03(s) {
  background(s); footer(s, 3);
  heading(s, 'About *Us*');
  intro(s, 'We are a team of blockchain experts dedicated to revolutionizing the way businesses operate through the power of decentralized technology. With years of experience in the industry and a passion for innovation, we are committed to creating secure and efficient blockchain solutions for a better tomorrow.');
  TEAM.forEach(([name, role], i) => {
    const x = 0.673 + (i % 6) * 1.4815, y = 1.976 + Math.floor(i / 6) * 1.495;
    s.addShape('roundRect', { x, y, w: 1.223, h: 1.222, rectRadius: 0.104, fill: fill(C.white, 10), line: NOLINE });
    s.addShape('ellipse', { x: x + 0.125, y: y + 0.125, w: 0.344, h: 0.344, fill: fill(C.ph, 5), line: NOLINE });   // [image] avatar
    lines(s, name.split('\n'), { x: x + 0.109, y: y + 0.512, w: 1.05, h: 0.35, fontFace: F.reg, color: C.lime });
    text(s, role, { x: x + 0.103, y: y + 0.878, w: 1.097, h: 0.3, fontSize: 8 });
  });
}

// 4 — testimonial + counters
function slide04(s) {
  background(s); footer(s, 4);
  // quote card with a tail on its upper-left corner
  s.addShape('custGeom', {
    x: 4.347, y: 0.612, w: 4.826, h: 2.801, fill: fill(C.white, 10), line: NOLINE,
    points: [{ x: 0.30, y: 0 }, { x: 4.826, y: 0 }, { x: 4.826, y: 2.801 }, { x: 0.30, y: 2.801 },
    { x: 0.30, y: 0.95 }, { x: 0, y: 0.36 }, { x: 0.30, y: 0.36 }, { close: true }]
  });
  text(s, 'Working on the blockchain platform project was an incredible experience. The team was knowledgeable, supportive, and dedicated to delivering high-quality results. I was constantly challenged to think outside the box and push the boundaries of what was possible. The end result was a cutting-edge platform that exceeded all expectations. I am proud to be a part of such a talented and innovative company.',
    { x: 4.999, y: 0.927, w: 3.844, h: 1.102 });
  text(s, '\u201D', { x: 8.6, y: 0.55, w: 0.5, h: 0.5, fontFace: F.med, fontSize: 30, color: C.ink });
  text(s, 'Michael Johnson', { x: 4.999, y: 2.293, w: 1.4, h: 0.24, color: C.lime });
  lines(s, ['CEO', 'Decentral Solutions, Inc.'], { x: 4.989, y: 2.505, w: 1.9, h: 0.35 });
  glyph(s, 'pin', 5.019, 2.984, 0.09, 0.123, C.lime);
  text(s, 'San Francisco, California, USA', { x: 5.153, y: 2.931, w: 2.4, h: 0.24 });
  s.addShape('custGeom', {   // signature squiggle
    x: 7.463, y: 2.277, w: 1.344, h: 0.653, fill: NOLINE, line: stroke(C.white, 55, 1),
    points: [{ x: 0.05, y: 0.60 }, { x: 0.28, y: 0.08, curve: { type: 'cubic', x1: 0.10, y1: 0.15, x2: 0.16, y2: 0.05 } },
    { x: 0.42, y: 0.62, curve: { type: 'cubic', x1: 0.36, y1: 0.20, x2: 0.44, y2: 0.45 } },
    { x: 0.70, y: 0.12, curve: { type: 'cubic', x1: 0.50, y1: 0.68, x2: 0.58, y2: 0.10 } },
    { x: 1.00, y: 0.55, curve: { type: 'cubic', x1: 0.82, y1: 0.14, x2: 0.90, y2: 0.58 } }]
  });
  // rating pill with five stars
  s.addShape('roundRect', { x: 3.283, y: 2.207, w: 1.591, h: 0.423, rectRadius: 0.21, fill: fill(C.ink), line: NOLINE });
  for (let i = 0; i < 5; i++) s.addShape('star5', { x: 3.50 + i * 0.235, y: 2.32, w: 0.2, h: 0.2, fill: fill(C.lime), line: NOLINE });
  s.addShape('custGeom', {   // mouse cursor
    x: 3.729, y: 2.584, w: 0.203, h: 0.203, fill: fill(C.white), line: NOLINE,
    points: [{ x: 0, y: 0 }, { x: 0.203, y: 0.13 }, { x: 0.11, y: 0.145 }, { x: 0.155, y: 0.203 }, { x: 0.10, y: 0.203 }, { x: 0.06, y: 0.15 }, { close: true }]
  });
  // "75 satisfied customers" avatars
  [4.681, 4.956, 5.231].forEach(x => s.addShape('ellipse', { x, y: 3.661, w: 0.344, h: 0.344, fill: fill(C.ph, 5), line: { color: C.moss, width: 2 } }));
  s.addShape('ellipse', { x: 5.506, y: 3.661, w: 0.344, h: 0.344, fill: fill(C.olive), line: { color: C.deep, width: 2 } });
  text(s, '75', { x: 5.562, y: 3.735, w: 0.235, h: 0.2, fontSize: 8, align: 'center' });
  text(s, '75 satisfied customers', { x: 5.903, y: 3.732, w: 1.4, h: 0.24, fontSize: 8 });
  [['27', 'professional hard-\nskilled employees', 4.701], ['85', 'successfully\ncompleted projects', 6.409], ['150+', 'customisation\ntasks completed', 8.118]]
    .forEach(([n, label, x]) => {
      text(s, n, { x, y: 4.155, w: 0.9, h: 0.38, fontFace: F.med, fontSize: 19, color: C.lime });
      lines(s, label.split('\n'), { x, y: 4.492, w: 1.35, h: 0.35 });
    });
}

// 5 — pricing tables
const PLANS = [
  ['Base Plan', '25', 2], ['Popular Plan', '35', 4], ['Gold Plan', '48', 5], ['Premium Plan', '50', 9]
];
const PLAN_FEATURES = ['Blockchain platform development', 'Smart contract development', 'Decentralized app development',
  'Custom blockchain solutions', 'Technical support', 'Maintenance and upgrades', 'Security audits',
  'Integration (existing systems)', 'Consulting services'];
function slide05(s) {
  background(s); footer(s, 5);
  heading(s, '*Pricing* Tables');
  intro(s, 'Discover our flexible pricing options and choose the plan that fits your business needs. Our plans are designed to accommodate businesses of all sizes, from startups to enterprise-level organizations. All plans come with 24/7 support and access to our full suite of blockchain solutions. Let us help you take your business to the next level with our innovative technology.');
  PLANS.forEach(([name, price, included], i) => {
    const x = 0.664 + i * 2.225;
    s.addShape('roundRect', { x, y: 1.860, w: 1.965, h: 2.953, rectRadius: 0.159, fill: fill(C.white, 10), line: NOLINE });
    text(s, name, { x: x + 0.210, y: 2.020, w: 1.31, h: 0.24, fontFace: F.med, color: C.ink });
    text(s, price, { x: x + 0.212, y: 2.230, w: 0.62, h: 0.44, fontFace: F.med, fontSize: 23, color: C.lime });
    text(s, '$/ h', { x: x + 0.677, y: 2.288, w: 0.38, h: 0.2, fontSize: 8 });
    PLAN_FEATURES.forEach((feat, k) => text(s, feat, {
      x: x + 0.208, y: 2.740 + k * 0.1435, w: 1.70, h: 0.15, fontSize: 7, strike: k >= included
    }));
    s.addShape('roundRect', { x: x + 0.205, y: 4.230, w: 1.555, h: 0.387, rectRadius: 0.078, fill: fill(C.lime), line: NOLINE });
    s.addShape('roundRect', { x: x + 1.369, y: 4.230, w: 0.391, h: 0.387, rectRadius: 0.078, fill: fill(C.pure, 25), line: NOLINE });
    text(s, 'More Inform', { x: x + 0.366, y: 4.310, w: 0.95, h: 0.2, fontFace: F.reg, fontSize: 8, color: C.ink });
    s.addShape('triangle', { x: x + 1.505, y: 4.36, w: 0.09, h: 0.125, rotate: 90, fill: fill(C.deep), line: NOLINE });
  });
}

// 6 — customer principles + client logo strip
const PRINCIPLES = [
  ['Transparency', '01', 'We believe in being transparent in every aspect of our work, from project timelines to budgeting and everything in between. This allows us to build trust with our clients and ensure that expectations are aligned from the start.'],
  ['Collaboration', '02', 'Collaboration is key to delivering successful projects. We work closely with our clients to understand their needs and goals, and we involve them in every step of the process to ensure that the final product meets their expectations.'],
  ['Trust', '03', 'Trust is the foundation of any successful business relationship. We take pride in our ability to deliver results, and we always go the extra mile to ensure that our clients are satisfied with the services we provide. By building trust with our clients, we establish long-lasting partnerships that drive results for both parties.']
];
const LOGOS = ['GRAPHIC\nDESIGN', 'GRAPHIC\nDESIGN INSTITUTE', 'AWARD\nWINNER', 'CENTRE DESIGN\nNAME', 'CENTRE\nDESIGN NAME', 'FINE\nARTS', 'GRAPHIC\nDESIGN INSTITUTE'];
function slide06(s) {
  background(s); footer(s, 6);
  heading(s, 'Our *Customers*');
  intro(s, 'At our company, we value our customers and their success above all else. Our approach to customer service is based on three key principles: *transparency*, *collaboration*, and t*rust*. We believe in open communication and working closely with our clients to understand their needs and deliver solutions that drive results. Take a look at some of the businesses we\u2019ve had the privilege of serving.');
  PRINCIPLES.forEach(([name, num, copy], i) => {
    const x = 0.653 + i * 3.076;
    text(s, num, { x: x + 1.447, y: 1.805, w: 1.322, h: 0.99, fontFace: F.med, fontSize: 56, align: 'right' });
    text(s, name, { x, y: 1.975, w: 1.95, h: 0.38, fontFace: F.med, fontSize: 19, color: C.lime });
    text(s, copy, { x: x - 0.003, y: 2.389, w: 2.649, h: 1.25 });
  });
  s.addShape('roundRect', { x: 0.641, y: 4.096, w: 8.668, h: 0.799, rectRadius: 0.154, fill: fill(C.white, 10), line: NOLINE });
  LOGOS.forEach((label, i) => {            // [image] client logos
    const x = 1.10 + i * 1.20;
    s.addShape('ellipse', { x, y: 4.34, w: 0.3, h: 0.3, fill: fill(C.ink), line: NOLINE });
    lines(s, label.split('\n'), { x: x + 0.36, y: 4.33, w: 0.85, h: 0.33, fontFace: F.med, fontSize: 6, color: C.ink });
  });
}

// 7 — six numbered benefits in two columns
const BENEFITS = [
  ['01/', 'Increased security', 'Blockchain provides a secure and tamper-proof ledger for recording transactions.'],
  ['02/', 'Improved transparency', 'Blockchain offers a decentralized and transparent system for tracking information and transactions.'],
  ['03/', 'Cost savings', 'Blockchain eliminates the need for intermediaries, reducing costs and increasing efficiency.'],
  ['04/', 'Streamlined processes', 'Blockchain automates many manual processes, saving time and reducing errors.'],
  ['05/', 'Improved accountability', 'Blockchain provides an auditable trail of all transactions, increasing accountability and reducing fraud.'],
  ['06/', 'Better data management', 'Blockchain allows for secure and efficient management of sensitive data.']
];
function slide07(s) {
  background(s); footer(s, 7);
  heading(s, 'Unlocking the Power of *Blockchain Technology*');
  intro(s, 'The benefits of using blockchain technology are vast and far-reaching. From increased security and transparency to cost savings and streamlined processes, businesses of all industries are turning to blockchain to gain a competitive advantage. Here are just a few of the ways blockchain can transform the way you do business:');
  BENEFITS.forEach(([num, title, copy], i) => {
    const x = 0.699 + Math.floor(i / 3) * 4.411, y = 1.877 + (i % 3) * 0.7615;
    text(s, num, { x, y, w: 0.338, h: 0.24, fontFace: F.med, fontSize: 11, align: 'right' });
    text(s, title, { x: x + 0.386, y, w: 2.5, h: 0.24, fontFace: F.med, fontSize: 11, color: C.lime });
    text(s, copy, { x: x + 0.381, y: y + 0.248, w: 3.529, h: 0.35 });
  });
  text(s, 'In conclusion, the adoption of blockchain technology offers countless benefits for businesses looking to stay ahead of the curve. From increased security and transparency to cost savings and new business opportunities, the potential of blockchain technology is limitless. By leveraging its power, businesses can improve their processes, build trust, and drive innovation and growth in the years to come.',
    { x: 0.637, y: 4.334, w: 8.445, h: 0.5 });
}

// Shared left-hand panel of slides 8 & 10: rounded photo frame + three KPI cards
const KPIS = [['6M', 'project implementation time'], ['27', 'professional hard-skilled employees'], ['150+', 'customisation tasks completed']];
function statPanel(s, customizationLabel) {
  // one translucent panel whose three KPI tiles are cut out, so they read darker
  panelWithHoles(s, 0.662, 0.649, 3.833, 4.297, 0.20,
    [[0.114, 3.032, 1.171, 1.113], [1.362, 3.032, 1.171, 1.113], [2.609, 3.032, 1.171, 1.113]],
    { fill: fill(C.white, 10), line: NOLINE });
  s.addShape('roundRect', { x: 0.766, y: 0.757, w: 3.625, h: 2.827, rectRadius: 0.12, fill: fill(C.ph, 5), line: NOLINE });  // [image]
  [0.890, 1.148, 1.406].forEach(x => s.addShape('ellipse', { x, y: 0.914, w: 0.344, h: 0.344, fill: fill(C.ph, 5), line: { color: C.moss, width: 2 } }));
  s.addShape('ellipse', { x: 1.664, y: 0.914, w: 0.344, h: 0.344, fill: fill(C.moss), line: { color: C.deep, width: 2 } });
  text(s, '25', { x: 1.664, y: 0.988, w: 0.344, h: 0.2, fontSize: 8, align: 'center' });
  text(s, '25 hard-skilled employees', { x: 2.115, y: 0.915, w: 1.337, h: 0.32, fontSize: 8 });
  KPIS.forEach(([n, label], i) => {
    const x = 0.853 + i * 1.2475;
    text(s, n, { x, y: 3.780, w: 0.816, h: 0.38, fontFace: F.med, fontSize: 19, color: C.lime });
    text(s, i === 2 ? customizationLabel : label, { x, y: 4.118, w: 1.02, h: 0.52 });
  });
}

// 8 — product page
function slide08(s) {
  background(s); footer(s, 8);
  statPanel(s, 'customisation tasks completed');
  heading(s, '*Product* Page', { x: 4.907, w: 3 });
  intro(s, 'Our blockchain platform development project is designed to deliver secure and efficient solutions for businesses of all sizes. By leveraging the power of decentralized technology, we aim to revolutionize the way businesses operate, improve transparency, and streamline processes. Here are just a few of the many benefits of our platform:',
    { x: 4.907, w: 4.364, h: 0.799 });
  bullets(s, ['Increased security', 'Improved transparency', 'Cost savings', 'Streamlined processes', 'Improved accountability'],
    { x: 4.907, y: 2.089, w: 2.205, h: 0.94, lineSpacingMultiple: 1.2 });
  bullets(s, ['Better data management', 'Increased trust', 'New business opportunities', 'Customizable solutions', 'Scalability'],
    { x: 7.232, y: 2.089, w: 1.97, h: 0.94, lineSpacingMultiple: 1.2 });
  // inline testimonial strip
  s.addShape('roundRect', { x: 4.907, y: 3.244, w: 4.364, h: 0.727, rectRadius: 0.08, fill: fill(C.white, 10), line: NOLINE });
  text(s, 'The team delivered on our blockchain project. Their expertise and efficient communication made the process successful. We recommend them.',
    { x: 5.06, y: 3.35, w: 3.55, h: 0.32, fontSize: 7 });
  text(s, [{ text: 'Elizabeth Davis', options: { color: C.lime } }, { text: ', Head of Business Development', options: { color: C.white } }],
    { x: 5.06, y: 3.70, w: 3.55, h: 0.2, fontSize: 7 });
  text(s, '\u201D', { x: 8.72, y: 3.30, w: 0.4, h: 0.4, fontFace: F.med, fontSize: 22, color: C.ink });
  text(s, 'In conclusion, our blockchain platform development project is poised to deliver significant benefits for businesses looking to stay ahead of the curve. By leveraging the power of decentralized technology, businesses can improve their processes, build trust, and drive innovation and growth in the years to come.',
    { x: 4.907, y: 4.114, w: 4.364, h: 0.8 });
}

// 9 — three project cards
const PROJECTS = [
  ['ChainFlow', 'A decentralized platform for secure and efficient supply chain management.', '160h', '150+'],
  ['TrustNet', 'A decentralized platform for secure and transparent voting systems.', '95h', '83'],
  ['GreenChain', 'A decentralized platform for tracking and verifying sustainable and eco-friendly products.', '250h', '540']
];
function slide09(s) {
  background(s); footer(s, 9);
  heading(s, '*Our Projects* Over the Past Year');
  intro(s, 'The benefits of using blockchain technology are vast and far-reaching. From increased security and transparency to cost savings and streamlined processes, businesses of all industries are turning to blockchain to gain a competitive advantage. Here are just a few of the ways blockchain can transform the way you do business:');
  PROJECTS.forEach(([name, copy, hours, tasks], i) => {
    const x = 0.664 + i * 2.9765;
    panelWithHoles(s, x, 1.860, 2.694, 2.954, 0.14,
      [[0.103, 2.070, 1.185, 0.750, 0.08], [1.406, 2.070, 1.185, 0.750, 0.08]],
      { fill: fill(C.white, 10), line: NOLINE });
    s.addShape('roundRect', { x: x + 0.103, y: 1.955, w: 2.488, h: 1.070, rectRadius: 0.09, fill: fill(C.ph, 5), line: NOLINE }); // [image]
    for (let k = 0; k < 5; k++) s.addShape('ellipse', { x: x + 0.197 + k * 0.217, y: 2.048, w: 0.260, h: 0.260, fill: fill(C.ph, 5), line: { color: C.ink, width: 2 } });
    text(s, name, { x: x + 0.147, y: 3.175, w: 1.6, h: 0.24, fontFace: F.med, fontSize: 10, color: C.lime });
    text(s, copy, { x: x + 0.144, y: 3.435, w: 2.35, h: 0.32, fontSize: 8 });
    [[hours, 'project implementation time', 'coin', 0.211], [tasks, 'customisation tasks completed', 'check', 1.508]].forEach(([v, label, ic, dx]) => {
      glyph(s, ic, x + dx + 0.05, 4.045, 0.125, 0.125, C.white);
      text(s, v, { x: x + dx + 0.24, y: 3.978, w: 0.9, h: 0.24, fontFace: F.med, fontSize: 11, color: C.lime });
      text(s, label, { x: x + dx + 0.21, y: 4.175, w: 0.90, h: 0.45, fontSize: 8 });
    });
  });
}

// 10 — stat panel + six mini product cards
const MINI = [
  ['ChainFlow', 'A decentralized platform for secure and efficient supply chain management.'],
  ['AssetChain', 'A blockchain-based platform for secure and efficient management of physical assets.'],
  ['BlockVault', 'A secure and decentralized platform for managing and tracking digital assets.'],
  ['SafeTrade', 'A decentralized platform for secure and efficient cross-border trade.'],
  ['TrustNet', 'A decentralized platform for secure and transparent voting systems.'],
  ['CarbonChain', 'A blockchain-based platform for tracking and reducing carbon emissions.']
];
function slide10(s) {
  background(s); footer(s, 10);
  statPanel(s, 'customization tasks completed');
  MINI.forEach(([name, copy], i) => {
    const x = 4.754 + (i % 2) * 2.417, y = 0.646 + Math.floor(i / 2) * 1.522;
    s.addShape('roundRect', { x, y, w: 2.158, h: 1.261, rectRadius: 0.16, fill: fill(C.white, 10), line: NOLINE });
    s.addShape('roundRect', { x: x + 0.108, y: y + 0.104, w: 1.053, h: 1.053, rectRadius: 0.085, fill: fill(C.ph, 5), line: NOLINE }); // [image]
    text(s, name, { x: x + 1.260, y: y + 0.075, w: 0.9, h: 0.24, fontFace: F.med, color: C.lime });
    text(s, copy, { x: x + 1.262, y: y + 0.29, w: 0.90, h: 0.85, fontSize: 7 });
  });
}

// Shared body for the "Future" / "Applications" hexagon-background slides
function hexTextSlide(s, num, cluster, title, body, listTitle, colA, colB, geo) {
  background(s); footer(s, num);
  hexCluster(s, cluster);
  heading(s, title, Object.assign({ x: geo.x, y: geo.titleY, w: geo.titleW, h: 0.646 }, {}));
  text(s, body, { x: geo.x, y: geo.bodyY, w: geo.bodyW, h: 1.4 });
  text(s, listTitle, { x: geo.x, y: geo.listY, w: 2.4, h: 0.24, fontFace: F.med, fontSize: 10, color: C.lime });
  bullets(s, colA, { x: geo.x, y: geo.listY + 0.302, w: 1.841, h: 0.648 });
  bullets(s, colB, { x: geo.x + geo.colGap, y: geo.listY + 0.302, w: 1.841, h: 0.648 });
}

// 11 — the future of blockchain
function slide11(s) {
  hexTextSlide(s, 11, CLUSTER_LEFT, 'The *Future* of Blockchain Technology',
    'As the world becomes increasingly digital, the potential for blockchain technology continues to grow. From increased security and transparency to efficient data management and decentralized systems, the benefits of implementing blockchain solutions are endless. With industries such as finance, supply chain, and healthcare already experiencing the positive impact of blockchain, the future looks bright for this game-changing technology.',
    'Brief List of Benefits:',
    ['Increased security', 'Improved transparency', 'Efficient data management', 'Decentralized systems'],
    ['Reduced costs', 'Improved accountability', 'Enhanced privacy', 'Streamlined processes'],
    { x: 5.376, titleY: 1.143, titleW: 3.7, bodyY: 1.920, bodyW: 3.632, listY: 3.441, colGap: 2.167 });
}

// 12 — diverse applications
function slide12(s) {
  hexTextSlide(s, 12, CLUSTER_RIGHT, 'The *Diverse Applications* of Blockchain Technology',
    'Blockchain technology has the potential to revolutionize a wide range of industries and fields. From finance and supply chain management to healthcare and voting systems, the implementation of blockchain solutions can bring increased security, transparency, and efficiency to various aspects of our daily lives. With the number of potential applications constantly growing, the future looks bright for blockchain technology.',
    'Brief List of Applications:',
    ['Finance', 'Supply chain management', 'Healthcare', 'Voting systems', 'Identity management', 'Real estate', 'Gaming'],
    ['Charity and non-profits', 'Digital advertising', 'Intellectual property management', 'Government services', 'Retail and e-commerce', 'Music industry'],
    { x: 0.626, titleY: 0.937, titleW: 3.5, bodyY: 1.691, bodyW: 4.243, listY: 3.050, colGap: 2.104 });
}

// 13 — how a blockchain works (7-step flow)
const FLOW = [
  ['1', 'A transaction is requested', 3.063, 3.154, 2.258, 2.631],
  ['2', 'A block that represents the transaction is created', 4.971, 3.154, 4.618, 2.632],
  ['3', 'The block is sent to every node in thee network', 7.212, 3.154, 6.801, 2.632],
  ['4', '', 0, 0, 8.999, 3.350],
  ['5', 'Nodes validate the transaction', 7.410, 4.519, 6.714, 4.016],
  ['6', 'Nodes receive a reward for the proof of work', 5.091, 4.519, 4.540, 4.013],
  ['7', 'A bock is added, o the existing blockchain', 2.916, 4.519, 2.380, 4.016],
  ['', 'The transaction is complete', 0.912, 4.519, 0, 0]
];
function slide13(s) {
  background(s); footer(s, 13);
  heading(s, 'How Does a *Blockchain Work*?');
  intro(s, 'A blockchain works by using complex algorithms to validate and record transactions in a secure and transparent way. When a transaction is made, it is verified by multiple participants in the network and added to a block. This block is then added to the existing chain of blocks, creating an unalterable and transparent record of all transactions. Each participant in the network has a copy of the blockchain, making it nearly impossible for any individual to manipulate the data. This decentralized and secure system has many applications, including financial transactions, supply chain management, and digital identity verification.',
    { h: 0.799 });
  // three gradient arrows bleeding off the left edge
  [[0.002, 2.177, 1.829], [0.008, 2.584, 2.200], [0.002, 2.991, 1.829]].forEach(([x, y, w]) =>
    chevron(s, x, y, w, 0.339, 0.152, { fill: fill(C.gInkOlive), line: NOLINE }));
  // connector arrows between the steps
  const arrow = (x, y, w, flip) => s.addShape('rightArrow', { x, y, w, h: 0.130, fill: fill(C.olive), line: NOLINE, flipH: !!flip });
  arrow(2.422, 2.700, 0.755); arrow(4.104, 2.699, 1.306); arrow(6.277, 2.700, 1.315);
  arrow(6.198, 4.084, 1.309, 1); arrow(4.021, 4.084, 1.308, 1); arrow(1.872, 4.084, 1.284, 1);
  s.addShape('custGeom', {  // the wrap-around elbow on the right
    x: 8.375, y: 2.738, w: 0.785, h: 1.476, fill: NOLINE, line: { color: C.olive, width: 1.5 },
    points: [{ x: 0, y: 0 }, { x: 0.60, y: 0 }, { x: 0.60, y: 1.476 }, { x: 0, y: 1.476 }]
  });
  // node illustrations
  const laptop = (x, y, w) => {  // screen + base with a small cube on it
    const h = w * 1.44;
    s.addShape('roundRect', { x, y, w, h: h * 0.72, rectRadius: 0.03, fill: fill(C.olive), line: NOLINE });
    s.addShape('custGeom', {
      x: x - w * 0.10, y: y + h * 0.76, w: w * 1.20, h: h * 0.19, fill: fill(C.olive), line: NOLINE,
      points: [{ x: w * 0.10, y: 0 }, { x: w * 1.10, y: 0 }, { x: w * 1.20, y: h * 0.19 }, { x: 0, y: h * 0.19 }, { close: true }]
    });
  };
  laptop(3.458, 2.439, 0.355);
  s.addShape('ellipse', { x: 3.502, y: 2.488, w: 0.267, h: 0.267, fill: fill(C.lime), line: NOLINE });
  text(s, 'B', { x: 3.502, y: 2.51, w: 0.267, h: 0.22, fontFace: F.med, fontSize: 12, color: C.ink, align: 'center' });
  // package / server that holds a cube
  s.addShape('roundRect', { x: 5.604, y: 2.549, w: 0.411, h: 0.472, rectRadius: 0.05, fill: fill(C.olive), line: NOLINE });
  s.addShape('roundRect', { x: 5.707, y: 2.488, w: 0.205, h: 0.123, rectRadius: 0.03, fill: fill(C.moss), line: NOLINE });
  cube(s, 5.717, 2.700, 0.185, 0.232);
  // network hubs (hex cube surrounded by satellite dots)
  const hub = (cx, cy, n, r) => {
    for (let i = 0; i < n; i++) {
      const a = -Math.PI / 2 + i * 2 * Math.PI / n;
      s.addShape('ellipse', { x: cx + r * Math.cos(a) - 0.031, y: cy + r * Math.sin(a) * 1.05 - 0.031, w: 0.062, h: 0.062, fill: fill(C.moss), line: NOLINE });
    }
  };
  [[7.983, 2.755], [7.983, 4.146], [3.646, 4.146]].forEach(([cx, cy]) => {
    s.addShape('roundRect', { x: cx - 0.221, y: cy - 0.252, w: 0.442, h: 0.503, rectRadius: 0.09, fill: fill(C.olive), line: NOLINE });
    hub(cx, cy, 6, 0.235); cube(s, cx - 0.0925, cy - 0.108, 0.185, 0.232);
  });
  s.addShape('roundRect', { x: 5.558, y: 3.894, w: 0.503, h: 0.503, rectRadius: 0.10, fill: fill(C.olive), line: NOLINE });
  hub(5.810, 4.146, 8, 0.267);
  s.addShape('ellipse', { x: 5.676, y: 4.012, w: 0.267, h: 0.267, fill: fill(C.lime), line: NOLINE });
  text(s, 'B', { x: 5.676, y: 4.034, w: 0.267, h: 0.22, fontFace: F.med, fontSize: 12, color: C.ink, align: 'center' });
  glyph(s, 'chip', 3.770, 3.916, 0.184, 0.184, C.lime);      // "+" badge on step 6
  glyph(s, 'shield', 8.087, 3.863, 0.225, 0.243, C.lime);    // reward badge on step 5
  // wallet with a coin, step 7
  s.addShape('roundRect', { x: 1.217, y: 4.022, w: 0.514, h: 0.390, rectRadius: 0.06, fill: fill(C.olive), line: NOLINE });
  s.addShape('roundRect', { x: 1.217, y: 3.961, w: 0.472, h: 0.123, rectRadius: 0.03, fill: fill(C.ink), line: NOLINE });
  s.addShape('ellipse', { x: 1.320, y: 3.879, w: 0.267, h: 0.205, fill: fill(C.lime), line: NOLINE });
  s.addShape('ellipse', { x: 1.628, y: 4.187, w: 0.123, h: 0.123, fill: fill(C.moss), line: NOLINE });
  // step numbers + captions
  FLOW.forEach(([num, caption, cx, cy, nx, ny]) => {
    if (caption) text(s, caption, { x: cx, y: cy, w: cx > 7 ? 1.57 : (caption.length > 30 ? 1.44 : 1.144), h: 0.295, fontSize: 8, align: 'center', valign: 'middle' });
    if (num) {
      s.addShape('ellipse', { x: nx, y: ny, w: 0.267, h: 0.267, fill: fill(C.moss), line: NOLINE });
      text(s, num, { x: nx, y: ny + 0.055, w: 0.267, h: 0.16, fontFace: F.reg, fontSize: 7, align: 'center', margin: 1.2 });
    }
  });
}

// Small server/monitor node used across slides 14 & 16
function serverNode(s, x, y, w, dim) {
  const body = dim ? C.olive : C.olive, shade = dim ? C.olive : C.moss, h = w * 1.44;
  s.addShape('roundRect', { x, y, w, h: h * 0.72, rectRadius: 0.02, fill: fill(body), line: NOLINE });
  s.addShape('custGeom', {
    x: x - w * 0.10, y: y + h * 0.755, w: w * 1.20, h: h * 0.19, fill: fill(body), line: NOLINE,
    points: [{ x: w * 0.12, y: 0 }, { x: w * 1.08, y: 0 }, { x: w * 1.20, y: h * 0.19 }, { x: 0, y: h * 0.19 }, { close: true }]
  });
  s.addShape('rect', { x: x - w * 0.17, y: y + h * 0.945, w: w * 1.34, h: h * 0.055, fill: fill(shade), line: NOLINE });
  cube(s, x, y + h * 0.09, w, h * 0.62);
}

// 14 — centralized vs distributed networks
function slide14(s) {
  background(s); footer(s, 14);
  text(s, rich('*Centralized* and* Distributed *Networks'), { x: 6.862, y: 0.526, w: 2.484, h: 0.24, align: 'right' });
  const dash = (x, y, w, h, flip) => s.addShape('line', { x, y, w, h, line: dashed(40), flipH: !!flip });
  // left diagram: hub-and-spoke
  const LEFT = [[1.524, 0.904], [2.273, 1.282], [0.775, 1.282], [2.275, 2.160], [0.777, 2.160], [1.524, 2.557]];
  dash(1.579, 1.052, 0, 0.608); dash(1.579, 2.010, 0, 0.608);
  dash(1.756, 1.945, 0.469, 0.273); dash(0.923, 1.461, 0.469, 0.273);
  dash(1.813, 1.460, 0.416, 0.246, 1); dash(0.954, 1.965, 0.416, 0.246, 1);
  LEFT.forEach(([x, y]) => serverNode(s, x, y, 0.109));
  serverNode(s, 1.489, 1.643, 0.181);
  // right diagram: fully-meshed hexagon
  const RIGHT = [[6.251, 0.893], [7.009, 1.268], [5.507, 1.268], [7.009, 2.138], [5.507, 2.138], [6.251, 2.532]];
  s.addShape('custGeom', { x: 5.561, y: 0.965, w: 1.503, h: 1.736, fill: NOLINE, line: dashed(40), points: HEX_PTS.map(p => ({ x: p[0] * 1.503, y: p[1] * 1.736 })).concat([{ close: true }]) });
  dash(6.307, 1.099, 0, 1.455);
  dash(5.673, 1.466, 1.263, 0.741, 1); dash(5.668, 1.462, 1.287, 0.744);
  dash(5.566, 2.262, 1.489, 0); dash(5.562, 0.968, 1.505, 0.741);
  dash(5.567, 1.401, 1.482, 1.300);
  RIGHT.forEach(([x, y]) => serverNode(s, x, y, 0.109));
  [['Centralized networks', 0.648, 'Distributed networks are computer networks where each device operates independently, and there is no central authority controlling the communication flow. In a distributed network, each node can communicate directly with other nodes. This type of network is often used in peer-to-peer applications. Distributed networks offer several advantages over centralized networks, such as resilience and fault tolerance, security, and scalability. However, they can be more complex to set up and manage than centralized networks.', 3.734],
  ['Distributed networks', 5.365, 'Centralized networks are computer networks where all devices are connected to a central server or a group of servers. The central server controls the communication flow, and all data goes through it. This type of network is widely used in organizations where there is a need for centralized control, such as in a corporate environment. The advantages of a centralized network include simplicity in management and maintenance, better bandwidth usage and data transfer speeds. However, they can be less resilient and more vulnerable to security breaches and cyberattacks.', 3.808]]
    .forEach(([title, x, copy, w]) => {
      text(s, title, { x, y: 2.979, w: 1.665, h: 0.24, fontFace: F.med, fontSize: 10, color: C.lime, valign: 'middle' });
      text(s, copy, { x: x - 0.014, y: 3.266, w, h: 1.557 });
    });
}

// 15 — decentralization hex diagram (6 labelled badges + concentric rings)
const DECENTRAL = [
  ['Permissionless', 'lock', 2.388, 0.872, 2.165, 0.622, 0.935, 'ctr'],
  ['Decentralization', 'chip', 3.748, 1.656, 3.452, 2.272, 1.049, 'ctr'],
  ['Trustless', 'shield', 3.748, 3.221, 3.675, 3.822, 0.603, 'ctr'],
  ['Transparent', 'search', 2.388, 4.008, 2.199, 4.588, 0.792, 'ctr'],
  ['Censorship resistant', 'doc', 1.027, 3.221, 0.770, 3.822, 0.984, 'ctr'],
  ['Programmable', 'chip', 1.026, 1.656, 0.802, 2.272, 0.920, 'ctr']
];
function slide15(s) {
  background(s); footer(s, 15);
  // seven nested hexagon outlines fading outwards
  [[1.250, 1.125, 2.747, 7], [1.148, 1.010, 2.946, 7], [1.067, 0.921, 3.103, 4], [0.994, 0.837, 3.252, 4],
  [0.927, 0.755, 3.393, 4], [0.860, 0.677, 3.528, 2], [0.791, 0.595, 3.665, 2]]
    .forEach(([x, y, w, op]) => hex(s, x, y, w, w * HEX_RATIO, { fill: NOLINE, line: stroke(C.pure, op) }));
  // inner star of dashed connectors
  const dash = (x, y, w, h, flip) => s.addShape('line', { x, y, w, h, line: dashed(40), flipH: !!flip });
  dash(1.236, 1.905, 2.820, 1.595); dash(1.179, 1.899, 2.834, 1.662, 1);
  dash(1.226, 1.906, 2.792, 1.615, 1); dash(1.223, 1.895, 2.805, 1.628);
  s.addShape('custGeom', { x: 2.152, y: 1.101, w: 0.476, h: 3.223, fill: NOLINE, line: dashed(40), points: [{ x: 0.471, y: 0 }, { x: 0, y: 0.806 }, { x: 0.005, y: 2.417 }, { x: 0.476, y: 3.223 }] });
  s.addShape('custGeom', { x: 2.629, y: 1.103, w: 0.482, h: 3.219, fill: NOLINE, line: dashed(40), points: [{ x: 0, y: 0 }, { x: 0.482, y: 0.805 }, { x: 0.482, y: 2.414 }, { x: 0, y: 3.219 }] });
  // small lime nodes along the connectors
  [[2.102, 1.857], [3.060, 1.857], [3.060, 2.642], [2.102, 2.642], [3.061, 3.454], [2.103, 3.454], [2.580, 3.053], [2.580, 2.251]]
    .forEach(([x, y]) => hex(s, x, y, 0.105, 0.121, { fill: fill(C.lime), line: NOLINE }));
  DECENTRAL.forEach(([label, ic, hx, hy, tx, ty, tw]) => {
    hexGlyph(s, ic, hx, hy, 0.471, C.olive, C.white, 0.20);
    text(s, label, { x: tx, y: ty, w: tw, h: label.length > 15 ? 0.345 : 0.24, align: 'center', valign: 'middle' });
  });
  heading(s, '*Decentralization* with Blockchain: Permissionless, Decentralization, Programmable, Censorship Resistant, Transparent, Trustless',
    { x: 5.104, y: 1.017, w: 3.979, h: 1.250 });
  text(s, 'Blockchain technology has revolutionized decentralization, offering features like permissionless decentralization, programmability, censorship resistance, and transparency. In a permissionless blockchain, anyone can participate in the network without the need for approval from a central authority. Smart contracts enable programmability, allowing developers to create custom applications on top of the blockchain. Censorship resistance is a crucial feature, ensuring that no single entity can manipulate or shut down the network. The transparency of the blockchain is another critical feature, enabling users to view all transactions and verify the integrity of the system. These features make blockchain a powerful tool for creating decentralized applications that operate securely and transparently.',
    { x: 5.104, y: 2.398, w: 4.140, h: 2.011 });
}

// A hexagonal "network of six nodes" cluster; `node` draws each vertex.
function hexNetwork(s, x, y, node) {
  hex(s, x, y, 1.068, 1.233, { fill: NOLINE, line: dashed(40) });
  s.addShape('line', { x: x + 0.001, y: y + 0.310, w: 1.066, h: 0.616, line: dashed(40) });
  s.addShape('line', { x, y: y + 0.309, w: 1.074, h: 0.611, line: dashed(40), flipH: true });
  s.addShape('line', { x: x + 0.535, y: y + 0.031, w: 0, h: 1.201, line: dashed(40) });
  [[0.428, 0.498], [0.428, -0.082], [0.428, 1.072], [0.928, 0.780], [-0.069, 0.780], [-0.069, 0.204], [0.927, 0.202]]
    .forEach(([dx, dy], i) => { if (i === 0) return; node(x + dx, y + dy); });
  node(x + 0.428, y + 0.498);
}

// 16 — four blockchain types, each drawn as a small hexagonal network
const TYPES = [
  ['Public blockchain', 'Public blockchains are decentralized and open to anyone who wants to participate, making them suitable for cryptocurrencies.', null, 0.647, 0.658],
  ['Private blockchain', 'Private blockchains are owned and controlled by a single entity or organization, restricting access to authorized participants.', 'lock', 2.917, 2.953],
  ['Consortium blockchain', 'Consortium blockchains are a hybrid of public and private blockchains, controlled by a group of organizations that work together to maintain and validate the network.', 'person', 5.184, 5.210],
  ['Hybrid blockchain', 'Hybrid blockchains combine both public and private blockchains, allowing for flexibility and transparency depending on the network\u2019s requirements.', 'lock', 7.454, 7.480]
];
function slide16(s) {
  bgFlip(s); footer(s, 16);
  heading(s, 'Types of Blockchains: *Public*, *Private*, *Consortium*, *Hybrid*');
  intro(s, 'Blockchains come in different types, each with its own unique characteristics. Public blockchains are decentralized and open to anyone who wants to participate, making them suitable for cryptocurrencies. Private blockchains, on the other hand, are owned and controlled by a single entity or organization, restricting access to authorized participants. Consortium blockchains are a hybrid of public and private blockchains, controlled by a group of organizations that work together to maintain and validate the network. Lastly, hybrid blockchains combine both public and private blockchains, allowing for flexibility and transparency depending on the network\u2019s requirements. Each type of blockchain has its own strengths and weaknesses, making it important to choose the right one for a particular use case.',
    { h: 0.951 });
  TYPES.forEach(([title, copy, badge, tx, nx]) => {
    hexNetwork(s, nx, 2.511, (px, py) => {
      hexCube(s, px - 0.106, py - 0.122, 0.212);
      if (badge) {
        s.addShape('ellipse', { x: px - 0.078, y: py - 0.128, w: 0.156, h: 0.156, fill: fill(C.ink), line: NOLINE });
        glyph(s, badge, px - 0.032, py - 0.088, 0.064, 0.076, C.pure);
      }
    });
    text(s, title, { x: tx, y: 3.945, w: 1.841, h: 0.24, fontFace: F.med, color: C.lime });
    text(s, copy, { x: tx, y: 4.141, w: 1.841, h: 0.8, fontFace: F.reg, fontSize: 8 });
  });
}

// 17 — Venn diagram of blockchain permission models
function slide17(s) {
  bgFlip(s); footer(s, 17);
  s.addShape('ellipse', { x: 0.665, y: 0.625, w: 3.147, h: 3.147, fill: fill(C.white, 10), line: NOLINE });
  s.addShape('ellipse', { x: 1.893, y: 1.853, w: 3.147, h: 3.147, fill: fill(C.white, 10), line: NOLINE });
  s.addShape('custGeom', {   // lime lens where the two circles overlap
    x: 1.890, y: 1.850, w: 1.920, h: 1.923, fill: fill(C.lime), line: NOLINE,
    points: [{ x: 1.920, y: 0 },
    { x: 0, y: 1.923, curve: { type: 'cubic', x1: 0.86, y1: 0.10, x2: 0.10, y2: 0.86 } },
    { x: 1.920, y: 0, curve: { type: 'cubic', x1: 0.10, y1: 1.06, x2: 0.86, y2: 1.83 } }, { close: true }]
  });
  [[3.471, 0.693, 180], [0.616, 4.267, 0]].forEach(([x, y, r]) =>
    chevron(s, x, y, 1.376, 0.339, 0.152, { rotate: r, fill: fill(C.gInkOlive), line: NOLINE }));
  text(s, 'Permissionless', { x: 3.659, y: 0.745, w: 1.031, h: 0.24 });
  text(s, 'Permissioned', { x: 0.782, y: 4.321, w: 1.031, h: 0.24, align: 'right' });
  glyph(s, 'shield', 1.707, 1.316, 0.283, 0.318, C.white);
  glyph(s, 'lock', 4.071, 2.852, 0.283, 0.318, C.white);
  glyph(s, 'person', 3.325, 3.844, 0.283, 0.318, C.white);
  glyph(s, 'lock', 2.712, 2.319, 0.283, 0.318, C.ink);
  lines(s, ['Public blockchain', 'No central authority'], { x: 1.189, y: 1.674, w: 1.320, h: 0.269, fontSize: 7, align: 'center', valign: 'middle' });
  lines(s, ['Private blockchain', 'Controlled by one authority'], { x: 3.552, y: 3.240, w: 1.320, h: 0.269, fontSize: 7, align: 'center', valign: 'middle' });
  lines(s, ['Consortium blockchain', 'Controlled by a group'], { x: 2.806, y: 4.216, w: 1.320, h: 0.269, fontSize: 7, align: 'center', valign: 'middle' });
  lines(s, ['Hybrid blockchain', 'Controlled by one authority with some permissionless processes'],
    { x: 2.194, y: 2.695, w: 1.320, h: 0.496, fontSize: 7, align: 'center', valign: 'middle', color: C.ink });
  heading(s, '*Types of Blockchains*: Public, Private, Consortium, Hybrid', { x: 5.818, y: 1.270, w: 3.44, h: 0.646 });
  text(s, 'Blockchains come in different types, each with its own unique characteristics. Public blockchains are decentralized and open to anyone who wants to participate, making them suitable for cryptocurrencies. Private blockchains, on the other hand, are owned and controlled by a single entity or organization, restricting access to authorized participants. Consortium blockchains are a hybrid of public and private blockchains, controlled by a group of organizations that work together to maintain and validate the network. Lastly, hybrid blockchains combine both public and private blockchains, allowing for flexibility and transparency depending on the network\u2019s requirements. Each type of blockchain has its own strengths and weaknesses, making it important to choose the right one for a particular use case.',
    { x: 5.818, y: 2.048, w: 3.588, h: 2.464 });
}

// 18 — opportunities vs challenges around a central network
const OPPS = [
  ['Transparent', 'search', 2.402, 1.989, 2.051, 1.472, 0.963],
  ['Reduced transaction time', 'clock', 1.897, 2.613, 1.070, 2.020, 1.216],
  ['Secure', 'shield', 2.402, 3.237, 0.835, 2.720, 0.963],
  ['Cost efficient', 'coin', 3.110, 3.861, 1.070, 3.344, 1.216],
  ['Irreversible transaction', 'lock', 3.110, 1.365, 1.778, 3.892, 1.216]
];
const CHALL = [
  ['Interoperability', 'arrows', 6.514, 1.365, 6.985, 1.472, 1.256],
  ['Scalability', 'chip', 7.230, 1.989, 7.693, 2.096, 1.216],
  ['Storage', 'doc', 7.738, 2.613, 8.201, 2.709, 1.216],
  ['Social acceptance', 'person', 7.230, 3.237, 7.693, 3.332, 1.216],
  ['Requires standardisation', 'bank', 6.514, 3.861, 6.966, 3.887, 1.216]
];
function slide18(s) {
  bgFlip(s); footer(s, 18);
  // vertical arrow ribbons top & bottom
  [[4.090, 0.646, 1.798, 90], [5.150, 0.307, 1.111, 90], [3.716, 0.307, 1.111, 90],
  [4.136, 4.513, 1.708, -90], [3.762, 4.852, 1.023, -90], [5.195, 4.852, 1.023, -90]]
    .forEach(([x, y, w, r]) => chevron(s, x, y, w, 0.518, 0.264, { rotate: r, fill: fill(C.gInkMoss), line: NOLINE }));
  // translucent chevrons + nested outlines forming the central "network" shell
  chevron(s, 3.057, 1.123, 1.460, 3.379, 0.365, { fill: fill(C.pure, 7), line: NOLINE });
  chevron(s, 5.460, 1.123, 1.460, 3.379, 0.365, { fill: fill(C.pure, 7), line: NOLINE, flipH: true });
  [[3.198, 1.277, 1.332, 3.074, 7], [3.290, 1.385, 1.242, 2.860, 5], [3.378, 1.485, 1.155, 2.660, 3],
  [3.469, 1.592, 1.067, 2.449, 3], [3.563, 1.700, 0.974, 2.231, 2], [3.656, 1.805, 0.880, 2.021, 2]]
    .forEach(([x, y, w, h, op]) => {
      const pts = [{ x: w, y: 0 }, { x: 0, y: h * 0.25 }, { x: 0, y: h * 0.75 }, { x: w, y: h }];
      s.addShape('custGeom', { x, y, w, h, fill: NOLINE, line: stroke(C.pure, op), points: pts });
      s.addShape('custGeom', { x: 9.994 - x - w, y, w, h, fill: NOLINE, line: stroke(C.pure, op), points: pts, flipH: true });
    });
  hexCube(s, 4.640, 2.780, 0.700);
  text(s, 'Blockchain Network ', { x: 4.310, y: 2.250, w: 1.353, h: 0.49, fontFace: F.reg, fontSize: 13, align: 'center', valign: 'bottom' });
  OPPS.concat(CHALL).forEach(([label, ic, hx, hy, tx, ty, tw], i) => {
    const left = i < OPPS.length;
    hexGlyph(s, ic, hx, hy, 0.353, left ? C.olive : C.ink, C.pure, 0.155);
    text(s, label, { x: tx, y: ty, w: tw, h: label.length > 18 ? 0.345 : 0.24, align: left ? 'right' : 'left', valign: 'middle' });
  });
  // side ribbons
  chevron(s, 0.003, 1.162, 1.016, 3.283, 0.445, { fill: fill(C.gInkOlive), line: NOLINE });
  chevron(s, 8.977, 1.162, 1.016, 3.283, 0.445, { fill: fill(C.gInkOlive), line: NOLINE, flipH: true });
  text(s, 'Opportunities', { x: -0.368, y: 2.701, w: 1.752, h: 0.24, rotate: -90, fontSize: 10, color: C.lime, align: 'center', valign: 'bottom' });
  text(s, '\u0421hallenges', { x: 8.605, y: 2.701, w: 1.752, h: 0.24, rotate: -90, fontSize: 10, color: C.lime, align: 'center', valign: 'bottom' });
}

// 19 — Hyperledger-style architecture diagram
const ARCH_ROW2 = [['CA SDK', 3.584], ['Ledger management', 5.021], ['Transaction management', 6.457], ['Client SDK', 7.893]];
const ARCH_BLOCKS = [
  ['MSP', 0.698, 0.940, [['Attribute certificate', 2.366], ['Assign roles', 2.369]], 2.111],
  ['Consensus services', 3.577, 3.839, [['P2P protocol', 5.274], ['Endorsement verification', 5.277]], 4.990],
  ['Smart contract services', 6.455, 6.696, [['Container execution', 8.147], ['Image repository', 8.150]], 7.869]
];
const ARCH_LABELS = [
  ['Register', 2.375, 1.391], ['Issue certificate', 2.373, 1.786], ['Rest API', 5.119, 1.120],
  ['Create channel', 6.585, 2.031], ['Invoke transaction', 2.514, 3.406], ['Endorsed transaction', 4.242, 3.352],
  ['Endorsement response', 5.966, 3.352], ['Join network', 7.690, 3.406],
  ['Deliver block', 5.502, 4.169], ['Deliver block', 3.580, 4.169], ['Gossip protocol', 4.550, 4.630]
];
function slide19(s) {
  background(s); footer(s, 19);
  const rr = (x, y, w, h, c, r, op) => s.addShape('roundRect', { x, y, w, h, rectRadius: r, fill: fill(c, op), line: NOLINE });
  const cell = (t, x, y, w, h, fs) => text(s, t, { x, y, w, h, fontSize: fs || 8, align: 'center', valign: 'middle' });
  const dash = (x, y, w, h, flip) => s.addShape('line', { x, y, w, h, line: dashed(50), flipH: !!flip });
  // client bar
  rr(0.668, 0.627, 8.663, 0.379, C.white, 0.052, 10);
  text(s, 'Client', { x: 4.541, y: 0.700, w: 0.9, h: 0.22, fontSize: 8, color: C.lime, align: 'center' });
  // SDK row
  rr(0.668, 1.416, 1.444, 0.530, C.ink, 0.104);
  rr(3.557, 1.416, 5.772, 0.530, C.white, 0.104, 10);
  ARCH_ROW2.forEach(([label, x]) => { rr(x, 1.440, 1.410, 0.477, C.ink, 0.083, 81); cell(label, x, 1.440, 1.410, 0.477); });
  cell('Certificate authority', 0.668, 1.416, 1.444, 0.530);
  // services row
  rr(0.668, 2.302, 8.661, 0.972, C.white, 0.104, 10);
  ARCH_BLOCKS.forEach(([title, bx, tx, subs, sx]) => {
    rr(bx, 2.329, 2.847, 0.917, C.ink, 0.070);
    cell(title, tx, 2.329, 0.900, 0.917);
    subs.forEach(([label], k) => { rr(sx, 2.356 + k * 0.443, 1.409, 0.421, C.olive, 0.048); cell(label, sx, 2.356 + k * 0.443, 1.409, 0.421); });
  });
  // peers row
  rr(0.664, 3.711, 8.661, 1.206, C.white, 0.052, 10);
  [[0.692, 'Endorsing', 0.721], [6.437, 'Non endorsing', 6.466]].forEach(([bx, title, cx]) => {
    rr(bx, 3.737, 2.863, 0.826, C.ink, 0.070);
    cell(title, bx, 3.780, 2.863, 0.25);
    rr(cx, 4.162, 1.389, 0.375, C.moss, 0.060); cell('Chain-code', cx, 4.162, 1.389, 0.375);
    rr(cx + 1.418, 4.162, 1.389, 0.375, C.moss, 0.060); cell('Chain-code', cx + 1.418, 4.162, 1.389, 0.375);
  });
  rr(4.514, 3.737, 0.972, 0.826, C.ink, 0.070); cell('Orderer', 4.514, 3.737, 0.972, 0.826);
  // dashed flow arrows
  [[2.415, 3.209], [4.134, 3.209], [5.852, 3.308], [7.571, 3.202], [6.444, 1.947], [4.996, 1.007]]
    .forEach(([x, y], i) => dash(x, y, 0, i === 2 ? 0.403 : (i === 4 ? 0.335 : (i === 5 ? 0.386 : 0.471))));
  dash(2.089, 1.785, 1.458, 0); dash(2.141, 1.559, 1.458, 0);
  dash(5.498, 4.148, 0.881, 0); dash(3.612, 4.150, 0.882, 0);
  s.addShape('custGeom', {   // gossip-protocol loop under the peers
    x: 2.126, y: 4.578, w: 5.747, h: 0.219, fill: NOLINE, line: dashed(50),
    points: [{ x: 0, y: 0 }, { x: 0, y: 0.219 }, { x: 5.747, y: 0.219 }, { x: 5.747, y: 0 }]
  });
  ARCH_LABELS.forEach(([t, x, y]) => text(s, t, { x, y, w: 0.919, h: 0.20, fontSize: 7, align: 'center', valign: 'middle' }));
  text(s, 'Architecture diagram of a blockchain', { x: 6.371, y: 5.064, w: 2.398, h: 0.223, color: C.pure, align: 'right', valign: 'bottom' });
}

// 20 — blockchain structure: six nodes in a hexagon, each with a ledger caption
const STRUCT_NODES = [
  [6.183, 1.225, 6.101, 0.976, 'ctr', 5.289, 0.976, 'r', 6.046, 1.029, 1],
  [7.578, 1.225, 7.496, 0.976, 'ctr', 8.294, 0.976, 'l', 8.088, 1.029, 0],
  [8.376, 2.516, 8.292, 3.165, 'ctr', 8.292, 3.560, 'ctr', 8.581, 3.416, 2],
  [7.581, 3.801, 7.496, 4.447, 'ctr', 8.296, 4.447, 'l', 8.090, 4.500, 0],
  [6.183, 3.804, 6.101, 4.447, 'ctr', 5.290, 4.447, 'r', 6.047, 4.500, 1],
  [5.383, 2.516, 5.279, 3.168, 'ctr', 5.279, 3.560, 'ctr', 5.574, 3.416, 2]
];
function slide20(s) {
  bgFlip(s); footer(s, 20);
  heading(s, 'Blockchain *Structure*', { x: 0.649, y: 1.305, w: 3.44 });
  text(s, 'The blockchain structure is the foundation of the decentralized and secure nature of blockchain technology. It is composed of blocks, which are linked together in chronological order to form a chain. Each block contains a set of transactions, a timestamp, and a unique cryptographic hash that connects it to the previous block in the chain. The cryptographic hash ensures that any tampering with a block would break the chain, making it immediately detectable. This structure ensures that data on the blockchain is immutable and transparent, making it a highly secure and trustworthy platform for storing and transmitting information. As the number of transactions on the blockchain increases, so does its security, making it an attractive option for a wide range of industries.',
    { x: 0.649, y: 1.780, w: 3.588, h: 2.302 });
  hex(s, 5.839, 1.319, 2.585, 2.985, { fill: NOLINE, line: dashed(40), rotate: 90 });
  s.addShape('line', { x: 6.383, y: 1.524, w: 1.490, h: 2.579, line: dashed(40), flipH: true });
  s.addShape('line', { x: 6.396, y: 1.520, w: 1.478, h: 2.598, line: dashed(40) });
  s.addShape('line', { x: 5.640, y: 2.814, w: 2.907, h: 0, line: dashed(40) });
  [[6.401, false], [5.645, true]].forEach(([x, flip]) => s.addShape('custGeom', {
    x, y: flip ? 1.543 : 1.536, w: 2.232, h: 2.545, fill: NOLINE, line: dashed(40), flipH: flip,
    points: [{ x: 0, y: 0 }, { x: 0, y: 2.545 }, { x: 2.232, y: 1.298 }, { close: true }]
  }));
  text(s, 'Consensus', { x: 6.564, y: 2.696, w: 1.144, h: 0.231, fontFace: F.reg, fontSize: 11, color: C.lime, align: 'center', valign: 'middle' });
  STRUCT_NODES.forEach(([hx, hy, nx, ny, na, lx, ly, la, cx, cdir]) => {
    hexCube(s, hx, hy, 0.512);
    text(s, 'Node', { x: nx, y: ny, w: 0.676, h: 0.194, align: na, valign: 'middle' });
    text(s, 'Ledger', { x: lx, y: ly, w: 0.676, h: 0.194, align: la, valign: 'middle' });
    // chevron pointing from the node toward its ledger caption
    s.addShape('custGeom', {
      x: cx, y: cdir === 2 ? 3.416 : (ly < 2 ? 1.029 : 4.500), w: 0.103, h: 0.087,
      rotate: cdir === 2 ? 90 : 0, flipH: cdir === 1, fill: fill(C.lime), line: NOLINE,
      points: [{ x: 0, y: 0 }, { x: 0.052, y: 0.0435 }, { x: 0, y: 0.087 }, { x: 0.024, y: 0.087 },
      { x: 0.076, y: 0.0435 }, { x: 0.024, y: 0 }, { close: true }]
    });
  });
}

// 21 — three linked blocks, each listing its fields
const BLOCK_FIELDS = ['Header', 'Previous block address', 'Timestamp', 'Nonce', 'Merkel root'];
function slide21(s) {
  bgFlip(s); footer(s, 21);
  heading(s, 'Blockchain *Structure* Scheme', { y: 0.498 });
  intro(s, 'The blockchain structure is the foundation of the decentralized and secure nature of blockchain technology. It is composed of blocks, which are linked together in chronological order to form a chain. Each block contains a set of transactions, a timestamp, and a unique cryptographic hash that connects it to the previous block in the chain. The cryptographic hash ensures that any tampering with a block would break the chain, making it immediately detectable.',
    { y: 0.973, h: 0.648 });
  // faint chain of tiny cubes running behind the blocks
  for (let i = 0; i < 18; i++) {
    const x = (i < 9 ? 2.637 : 5.346 - 9 * 0.2115) + i * 0.2115;
    cube(s, x, 2.309, 0.348, 0.248);
    s.addShape('rect', { x, y: 2.309, w: 0.348, h: 0.248, fill: fill(C.moss, 88), line: NOLINE });
  }
  ['Block n-1', 'Block n', 'Block n+1'].forEach((label, i) => {
    const x = 1.325 + i * 2.7075;
    text(s, label, { x: x + 0.516, y: 1.797, w: 0.910, h: 0.194, align: 'center', valign: 'middle' });
    hexCube(s, x + 0.512, 1.889, 0.923);
    chevron(s, x + 0.717, 2.671, 0.509, 0.518, 0.142, { rotate: 90, fill: fill(C.gInkOlive), line: NOLINE });
    BLOCK_FIELDS.forEach((f, k) => {
      s.addShape('roundRect', { x, y: 3.218 + k * 0.365, w: 1.948, h: 0.312, rectRadius: 0.053, fill: fill(C.ink), line: NOLINE });
      text(s, f, { x: x + 0.181, y: 3.218 + k * 0.365, w: 1.586, h: 0.312, fontSize: 8, align: 'center', valign: 'middle' });
    });
  });
}

// 22 — five process steps along a horizontal track
const STEPS = [
  ['01', 'P2P Network', 'Someone in the Peer to Peer network request a transaction', 'chip'],
  ['02', 'Communication', 'The requested transaction is broadcast to the P2P network consisting of computers, known as nodes.', 'arrows'],
  ['03', 'Validation', 'The network of nodes validated the transaction and the users status using algorithms. A verified transaction can involve cryptocurrency, contracts, records or other information.', 'shield'],
  ['04', 'Verification', 'Once verified, the transaction is combined with other transactions to create a new block of data for the ledger.', 'search'],
  ['05', 'Conformation', 'The new block is then added to the existing blockchain, in a way that is permanent and unalterable.\nThe transaction is complete.', 'chip']
];
function slide22(s) {
  bgFlip(s); footer(s, 22);
  heading(s, 'Blockchain *Technology Process* Steps', { y: 0.633 });
  intro(s, 'Blockchain technology involves a set of process steps that ensure secure and transparent transactions. First, a transaction is initiated and broadcasted to the network. The transaction is then validated and confirmed by nodes in the network using consensus mechanisms such as Proof of Work or Proof of Stake. Once the transaction is validated, it is added to a block, which is linked to the previous block in the chain through a cryptographic hash. The block is then broadcasted to the network, and once confirmed by the nodes, it is added to the blockchain. The transaction is now immutable, transparent, and secure, ensuring that it cannot be tampered with or deleted. This process provides a highly secure and trustworthy platform for transactions and data storage.',
    { y: 1.108, h: 0.951 });
  // thin shaft with a wide open arrowhead at the right
  s.addShape('custGeom', {
    x: 0.673, y: 2.565, w: 8.666, h: 0.700, fill: fill(C.snow, 10), line: NOLINE,
    points: [{ x: 0.03, y: 0.29 }, { x: 8.10, y: 0.29 }, { x: 7.95, y: 0 }, { x: 8.30, y: 0 },
    { x: 8.666, y: 0.35 }, { x: 8.30, y: 0.700 }, { x: 7.95, y: 0.700 }, { x: 8.10, y: 0.41 },
    { x: 0.03, y: 0.41 }, { close: true }]
  });
  s.addShape('ellipse', { x: 0.673, y: 2.825, w: 0.16, h: 0.18, fill: fill(C.snow, 10), line: NOLINE });
  STEPS.forEach(([num, title, copy, ic], i) => {
    const x = 0.821 + i * 1.7665;
    text(s, num, { x: x + 0.128, y: 2.416, w: 0.6, h: 0.344, fontFace: F.med, fontSize: 17, valign: 'bottom' });
    s.addShape('line', { x: x + 0.063, y: 2.467, w: 0, h: 0.957, line: dashed(50) });
    glyph(s, ic, x + 0.142, 3.216, 0.21, 0.21, C.pure);
    text(s, title, { x: x + 0.010, y: 3.459, w: 1.3, h: 0.214, fontFace: F.reg, color: C.lime, valign: 'bottom' });
    lines(s, copy.split('\n'), { x, y: 3.686, w: i === 2 ? 1.554 : 1.389, h: 1.15, fontSize: 8 });
  });
}

// 23 — five hexagon "disadvantage" badges with alternating callouts
const DISADV = [
  ['Scalability', 'chip', 'Public blockchains can be slow and have limited scalability due to their decentralized nature. This can result in slower transaction times and higher fees.', 1, 0.981, 1.884],
  ['Privacy', 'lock', 'Public blockchains are transparent, meaning that all transactions and data are visible to anyone on the network. This lack of privacy may not be suitable for some use cases, such as confidential financial transactions.', 0, 2.421, 2.082],
  ['Governance', 'bank', 'Public blockchains are governed by a decentralized network of participants, making it difficult to implement changes or updates to the network.', 1, 4.095, 1.778],
  ['Energy consumption', 'bolt', 'Public blockchains rely on a proof-of-work consensus mechanism, which requires significant computational power and energy consumption.', 0, 5.664, 1.728],
  ['Regulation', 'doc', 'As public blockchains are open and decentralized, they may be difficult to regulate and monitor by governments and regulatory bodies.', 1, 7.103, 1.924]
];
function slide23(s) {
  bgFlip(s); footer(s, 23);
  heading(s, '*Public Blockchain *Technology Disadvantages', { y: 0.633 });
  text(s, 'While public blockchain technology offers many advantages, there are also several disadvantages to consider:', { x: 0.659, y: 1.108, w: 7.487, h: 0.24 });
  DISADV.forEach(([label, ic, copy, up, tx, tw], i) => {
    const hx = 1.406 + i * 1.5375;
    // faint half-hexagon shards linking the badges into a chain
    s.addShape('custGeom', {
      x: hx - 0.302, y: 2.623, w: 0.087, h: 0.942, fill: fill(C.pure, 15), line: NOLINE,
      points: [{ x: 0, y: 0 }, { x: 0, y: 0.942 }, { x: 0.086, y: 0.885 }, { x: 0.087, y: 0.047 }, { close: true }]
    });
    hex(s, hx, 2.507, 1.020, 1.177, { fill: fill(C.white, 10), line: NOLINE });
    glyph(s, ic, hx + 0.41, 2.801, 0.20, 0.219, C.lime);
    text(s, label, { x: hx + 0.019, y: 3.143, w: 0.995, h: 0.188, fontSize: 8, align: 'center', valign: 'middle' });
    s.addShape('line', { x: hx + 0.512, y: up ? 2.405 : 3.535, w: 0, h: 0.265, line: dashed(50) });
    text(s, copy, { x: tx, y: up ? 1.678 : 3.840, w: tw, h: 0.8, fontSize: 8, align: 'center', valign: up ? 'bottom' : 'top' });
  });
}

// 24 — seven-stage integration ribbon
const STAGES = [
  ['Wallet', 'An encrypted transaction is created by the transmitting wallet and submitted to the network node. The submission can include a fee.', 0.656, 1.393, 2.216, 0.837],
  ['Network broadcast', 'The pending transactions are broadcasted over a network.', 2.003, 1.286, 2.557, 0.496],
  ['Pool of proposed transactions', 'The transaction is validated by Node into a queue of proposed transactions.', 3.242, 1.286, 2.557, 0.496],
  ['Consensus mechanisms', 'Miners or validators use consensus algorithms, choose transactions from the pool and assemble them into authoritative ledger and blocks.', 4.482, 1.286, 2.216, 0.837],
  ['Ledger of confirmed transactions', 'Once the consensus is reached, miners or validators get access to a block reward for incorporating the authenticated reward.', 5.722, 1.286, 2.330, 0.724],
  ['Network ledger replication', 'Miners or validators share the new block of encrypted transactions to all network nodes for creating trusted records.', 6.961, 1.286, 2.330, 0.724],
  ['Wallet', 'The validated transactions are factored to its store of value by receiving wallet.', 8.201, 1.136, 2.444, 0.610]
];
function slide24(s) {
  bgFlip(s); footer(s, 24);
  heading(s, 'Blockchain Technology *Integration*: Anatomy and Key Considerations', { y: 0.498 });
  intro(s, 'Integrating blockchain technology requires careful consideration of various factors, including the specific use case, regulatory requirements, and the compatibility of existing systems. Key considerations include identifying the right blockchain platform, ensuring interoperability with other systems, and addressing potential security risks. A successful blockchain integration requires a deep understanding of the technology and the ability to design and implement solutions that meet the unique needs of the organization. With the right strategy and approach, blockchain technology can transform businesses and industries by enabling new levels of transparency, security, and efficiency.',
    { y: 0.973, h: 0.799 });
  STAGES.forEach(([title, copy, x, w, ty, th], i) => {
    const tall = i === 3;
    chevron(s, x, 3.671, w, tall ? 0.847 : 0.208, 0.10, { fill: fill(C.white, 10), line: NOLINE });
    text(s, String(i + 1), { x: x + 0.072, y: 3.687, w: 0.303, h: 0.175, fontSize: 8, valign: 'bottom' });
    text(s, title, { x: x + 0.037, y: 3.081, w: 1.07, h: 0.52, fontFace: F.reg, color: C.lime, valign: 'bottom' });
    text(s, copy, { x: x + 0.042, y: ty, w: 1.144, h: th, fontSize: 7, valign: 'bottom' });
  });
  chevron(s, 4.987, 4.886, 0.467, 0.216, 0.113, { rotate: 90, fill: fill(C.gOliveMoss), line: NOLINE });
  s.addShape('ellipse', { x: 5.046, y: 4.037, w: 0.343, h: 0.343, fill: fill(C.lime), line: NOLINE });
  glyph(s, 'chip', 5.111, 4.102, 0.212, 0.212, C.ink);
  text(s, 'Blocks', { x: 5.469, y: 4.093, w: 0.893, h: 0.194, fontFace: F.reg, color: C.lime, valign: 'bottom' });
  hexCube(s, 5.023, 4.560, 0.405);
}

// 25 — value-proposition pyramid
const PYRAMID = [
  ['Advantages:', ['Immutability', 'Traceability', 'Security', 'Privacy', 'Interoperability', 'Reduced costs'], 3.826, 3.993, 1.014],
  ['Data use cases:', ['Supply chain', 'Financial', 'Certificates', 'Legal contracts', 'Ownership', 'Personal data'], 5.330, 3.981, 1.014],
  ['', ['Identity management', 'Permission management', 'Activity logs', 'Profile management', 'Voting'], 6.834, 4.049, 0.875],
  ['Analytics and Automation:', ['Smart business contracts', 'Intelligent legal contracts', 'Oracles', 'Data analytics'], 3.826, 2.203, 1.153],
  ['Crypto-economic models:', ['Tokens and incentive mechanisms', 'Micropayments', 'Investment', 'Micro-lending'], 5.334, 2.203, 1.014],
  ['Decentralised governance:', ['Distributed ownership', 'Democratic decisions', 'Decentralised autonomous Organization DAO'], 3.826, 0.946, 0.736]
];
function slide25(s) {
  bgFlip(s); footer(s, 25);
  hexCluster(s, CLUSTER_PYRAMID);
  heading(s, 'The *Blockchain Value *Proposition Pyramid', { x: 6.357, y: 0.542, w: 2.964, h: 0.646, align: 'right' });
  // the stepped pyramid on the left
  s.addShape('custGeom', { x: 2.377, y: 0.622, w: 0.723, h: 1.445, fill: fill(C.white, 5), line: NOLINE, points: [{ x: 0.723, y: 0 }, { x: 0, y: 1.445 }, { x: 0.723, y: 1.445 }, { close: true }] });
  s.addShape('custGeom', { x: 1.641, y: 2.118, w: 1.459, h: 1.420, fill: fill(C.white, 5), line: NOLINE, points: [{ x: 0.715, y: 0 }, { x: 0, y: 1.420 }, { x: 1.459, y: 1.420 }, { x: 1.459, y: 0 }, { close: true }] });
  s.addShape('custGeom', { x: 0.909, y: 3.592, w: 2.191, h: 1.413, fill: fill(C.white, 5), line: NOLINE, points: [{ x: 0.701, y: 0 }, { x: 0, y: 1.413 }, { x: 2.191, y: 1.413 }, { x: 2.191, y: 0 }, { close: true }] });
  [[0.624, 1.443, 'Advanced', 1.250], [2.117, 1.422, 'Intermediate', 2.734], [3.593, 1.411, 'Base', 4.217]]
    .forEach(([y, h, label, ty]) => {
      s.addShape('rect', { x: 3.112, y, w: 0.348, h, fill: fill(C.white, 10), line: NOLINE });
      text(s, label, { x: 2.616, y: ty, w: 1.341, h: 0.188, rotate: -90, fontFace: F.reg, fontSize: 8, color: C.lime, align: 'center', valign: 'middle' });
    });
  // the two axis arrows
  [[0.032, 3.762, 2.202], [0.917, 1.618, 2.598]].forEach(([x, y, w]) =>
    chevron(s, x, y, w, 0.336, 0.193, { rotate: -63, fill: fill(C.gInkOlive), line: NOLINE }));
  text(s, 'Complexity of implementation', { x: 1.013, y: 2.078, w: 1.989, h: 0.188, rotate: -63, fontFace: F.reg, fontSize: 8, align: 'right', valign: 'middle' });
  text(s, 'Degree of disruption', { x: 0.267, y: 3.865, w: 1.689, h: 0.188, rotate: -63, fontFace: F.reg, fontSize: 8, align: 'right', valign: 'middle' });
  s.addShape('line', { x: 1.656, y: 3.567, w: 7.668, h: 0, line: dashed(40) });
  s.addShape('line', { x: 2.398, y: 2.094, w: 6.927, h: 0, line: dashed(40) });
  text(s, 'Decentralised data infrastructure', { x: 3.811, y: 3.698, w: 2.288, h: 0.188, fontFace: F.reg, fontSize: 8, color: C.lime, valign: 'middle' });
  text(s, 'Memership management', { x: 6.872, y: 3.698, w: 2.153, h: 0.188, fontFace: F.reg, fontSize: 8, color: C.lime, valign: 'middle' });
  PYRAMID.forEach(([title, items, x, y, h]) => {
    const runs = [];
    if (title) runs.push({ text: title, options: { color: C.lime, breakLine: true, bullet: false } });
    items.forEach(t => runs.push({ text: t, options: { breakLine: true, bullet: { characterCode: '002D', indent: 7 } } }));
    text(s, runs, { x, y, w: 1.46, h, fontFace: F.reg, fontSize: 8, valign: y > 3 ? 'middle' : 'top' });
  });
}

// 26 — pentagon advantages diagram
const ADV_NODES = [
  ['Decentralization', 'chip', 6.918, 1.285, 6.564, 1.049],
  ['Security', 'shield', 8.239, 2.238, 7.889, 1.982],
  ['Transparency', 'check', 7.734, 3.786, 7.380, 4.371],
  ['Accessibility', 'check', 6.099, 3.786, 5.731, 4.371],
  ['Immutability', 'lock', 5.602, 2.238, 5.248, 1.982]
];
const ADV_TEXT = [
  ['Decentralization: ', 'Public blockchains are decentralized, meaning that there is no central authority controlling the network. This makes the system more resilient and resistant to attacks.'],
  ['Transparency: ', 'Public blockchains are transparent, meaning that anyone can view the transactions and data on the network. This enhances trust and accountability within the system.'],
  ['Security: ', 'Public blockchains are secured using cryptographic techniques that make it nearly impossible to tamper with the data or transactions on the network.'],
  ['Accessibility: ', 'Public blockchains are open to anyone who wants to participate, making it easy for individuals and organizations to join the network and contribute to its growth and development.'],
  ['Immutability: ', 'Public blockchains are designed to be immutable, meaning that once a transaction is recorded on the blockchain, it cannot be changed or deleted. This enhances the reliability and accuracy of the system.']
];
function slide26(s) {
  bgFlip(s); footer(s, 26);
  heading(s, 'Public Blockchain Technology *Advantages*', { x: 0.640, y: 0.612, w: 3.44, h: 0.646 });
  const runs = [];
  ADV_TEXT.forEach(([lead, body], i) => {
    if (i) runs.push({ text: '', options: { breakLine: true, bullet: false } });
    runs.push({ text: lead, options: { color: C.lime, bullet: { characterCode: '002D', indent: 7 } } });
    runs.push({ text: body, options: { breakLine: true } });
  });
  text(s, runs, { x: 0.657, y: 1.492, w: 4.024, h: 3.223, valign: 'middle' });
  // five nested pentagon outlines
  [[5.418, 1.087, 3.543, 3.370, 10], [5.323, 0.987, 3.733, 3.550, 8], [5.238, 0.897, 3.904, 3.713, 5],
  [5.152, 0.807, 4.075, 3.876, 3], [5.069, 0.720, 4.241, 4.033, 1]]
    .forEach(([x, y, w, h, op]) => s.addShape('pentagon', { x, y, w, h, fill: NOLINE, line: stroke(C.pure, op) }));
  s.addShape('pentagon', { x: 5.813, y: 1.501, w: 2.754, h: 2.620, fill: NOLINE, line: dashed(40) });
  s.addShape('pentagon', { x: 6.334, y: 2.050, w: 1.712, h: 1.628, fill: fill(C.white, 10), line: NOLINE });
  cube(s, 7.078, 2.464, 0.224, 0.259);
  lines(s, ['Public blockchain', 'technology', 'advantages'], { x: 6.526, y: 2.860, w: 1.338, h: 0.496, fontFace: F.reg, align: 'center', valign: 'bottom' });
  ADV_NODES.forEach(([label, ic, px, py, tx, ty]) => {
    s.addShape('pentagon', { x: px, y: py, w: 0.543, h: 0.516, fill: fill(C.olive), line: NOLINE });
    glyph(s, ic, px + 0.196, py + 0.180, 0.155, 0.156, C.white);
    text(s, label, { x: tx, y: ty, w: 1.251, h: 0.194, align: 'center', valign: 'middle' });
  });
}

// 27 — contact information
const OFFICES = [
  ['Paris', '10 Avenue Montaigne, 75008 Paris, France', 'Phone: +33-1-345-6789', 0.664],
  ['Berlin', 'Kurf\u00fcrstendamm 207-208, 10719 Berlin, Germany', 'Phone: +49-30-654-3210', 2.123],
  ['Los Angeles', '789 Sunset Blvd, Los Angeles, CA 90046, USA', 'Phone: +1-555-123-4567', 3.597]
];
function slide27(s) {
  background(s); footer(s, 27);
  hexCluster(s, CLUSTER_CONTACT);
  heading(s, '*Contact* Information', { x: 0.671, y: 1.275, w: 3.44 });
  text(s, 'Include your company name, address, phone number, and email address as basic contact information. Additional details like social media links, feedback forms, and business hours can also be useful. Make sure your contact information is easily accessible for customers to enhance their experience with your company.',
    { x: 0.671, y: 1.750, w: 4.078, h: 0.799 });
  OFFICES.forEach(([city, address, phone, x]) => {
    text(s, city, { x, y: 2.984, w: 0.999, h: 0.206, fontFace: F.reg, fontSize: 10, color: C.lime, valign: 'middle' });
    text(s, address, { x: x - 0.008, y: 3.207, w: 1.304, h: 0.5, fontSize: 8 });
    glyph(s, 'coin', x + 0.013, 3.685, 0.127, 0.127, C.lime);
    text(s, phone, { x: x - 0.008, y: 3.940, w: 1.304, h: 0.28, fontSize: 8 });
  });
  glyph(s, 'pin', 7.739, 2.596, 0.160, 0.219, C.lime);
}

/* ==================================================================== build */
const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
  slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18,
  slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27];

const pptx = new PptxGenJS();
pptx.author = 'Blockchain';
pptx.title = 'Blockchain Technology';
pptx.defineLayout({ name: 'DECK', width: W, height: H });
pptx.layout = 'DECK';
BUILDERS.forEach(build => build(pptx.addSlide()));
pptx.writeFile({ fileName: path.join(__dirname, '13c71d4f-34cf-4656-b442-2c88a28aa6f0_grok_final.pptx') })
  .then(f => console.log('wrote', f));
