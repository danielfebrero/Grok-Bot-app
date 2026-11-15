/**
 * "47 Info dark" isometric infographic deck — 18 slides, 13.333 x 7.5 in.
 *
 * The reference deck is flat-shaded isometric vector artwork. It is rebuilt
 * here from a small library of isometric primitives (pucks, boxes, slabs,
 * jigsaw pieces, screens, figures) driven by per-slide data tables, so the
 * design is visible directly in the source.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const BG = '242E3D';       // slide background (theme colour lt2/bg2)
const GY50 = '7D92B2';     // bg2 lumMod 50% — desk / device tops
const GY75 = '4A5F7E';     // bg2 lumMod 75% — mid greys
const GY90 = '334257';     // bg2 lumMod 90% — ground shadows, dark faces
const WHITE = 'FFFFFF';
const TITLE_BLUE = '0070C0';   // headline colour from slide layout "1_Title and Content"
const TITLE_BLUE6 = '3A67C9';  // headline colour on the centred "5 OPTIONS" slide

const F_HEAD = 'Open Sans Extrabold';
const F_BODY = 'Open Sans';

const SKIN = 'FDC9B1';
const SKIN_D = 'FDBB9D';
const HAIR = '231F20';
const SUIT = '3E454E';
const SUIT_D = '333B42';

const LOREM_FULL =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
  'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud ' +
  'exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure ' +
  'dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ' +
  'Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt ' +
  'mollit anim id est laborum.';
const LOREM_SHORT = LOREM_FULL.slice(0, LOREM_FULL.indexOf('nulla pariatur.') + 15);
const LOREM_TINY =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
  'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud ' +
  'exercitation ullamco laboris';

/* ------------------------------------------------------- shape primitives */

const NOLINE = { type: 'none' };

/** Darker variant of a hex colour, used for the shaded side of a solid. */
function shade(hex, f) {
  const k = f === undefined ? 0.82 : f;
  return [0, 2, 4]
    .map((i) => Math.round(parseInt(hex.substr(i, 2), 16) * k).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
}

function rect(s, x, y, w, h, color, opts) {
  s.addShape('rect', Object.assign({ x, y, w, h, fill: { color }, line: NOLINE }, opts));
}

function rrect(s, x, y, w, h, color, radius, opts) {
  s.addShape('roundRect', Object.assign(
    { x, y, w, h, fill: { color }, line: NOLINE, rectRadius: radius }, opts));
}

function ell(s, x, y, w, h, color, opts) {
  s.addShape('ellipse', Object.assign({ x, y, w, h, fill: { color }, line: NOLINE }, opts));
}

/** Unfilled ring (used for handles, mug ears, magnifier glass rims). */
function ring(s, x, y, w, h, color, pt) {
  s.addShape('ellipse', { x, y, w, h, fill: { color: BG }, line: { color, width: pt } });
}

/** Filled polygon from absolute [x, y] points in inches. */
function poly(s, pts, color, opts) {
  const xs = pts.map((p) => p[0]);
  const ys = pts.map((p) => p[1]);
  const x = Math.min.apply(null, xs);
  const y = Math.min.apply(null, ys);
  const w = Math.max(Math.max.apply(null, xs) - x, 0.005);
  const h = Math.max(Math.max.apply(null, ys) - y, 0.005);
  const points = pts.map((p) => ({ x: +(p[0] - x).toFixed(4), y: +(p[1] - y).toFixed(4) }));
  points.push({ close: true });
  s.addShape('custGeom', Object.assign(
    { x, y, w, h, points, fill: { color }, line: NOLINE }, opts));
}

const TEXT_INSET = [7.2, 7.2, 3.6, 3.6]; // pt — matches the 0.1in placeholder insets

function txt(s, text, opts) {
  s.addText(text, Object.assign(
    { fontFace: F_BODY, color: WHITE, align: 'left', valign: 'top', margin: 0 }, opts));
}

/** Text drawn in a box with the deck's standard placeholder insets. */
function txtBox(s, text, opts) {
  txt(s, text, Object.assign({ margin: TEXT_INSET }, opts));
}

/* -------------------------------------------------- isometric primitives */

const ISO = 0.58;    // box rhombus height : width
const ISOE = 0.578;  // cylinder ellipse height : width

/** Soft ground shadow ellipse under an isometric object. */
function shadow(s, x, y, w) {
  ell(s, x, y, w, w * 0.50, GY90);
}

/**
 * Isometric cylinder / puck: side wall, bottom ellipse and top ellipse.
 * `label` prints the step number either wrapped around the wall (`labelOn:
 * 'side'`) or flat on the top face, matching the two styles in the deck.
 */
function isoDisc(s, o) {
  const eh = o.w * ISOE;
  const d = o.depth;
  if (o.shadow) shadow(s, o.x + o.w * 0.036, o.y + d + eh * 0.30, o.w * 1.04);
  ell(s, o.x, o.y + d, o.w, eh, o.side);
  rect(s, o.x, o.y + eh / 2, o.w, d, o.side);
  ell(s, o.x, o.y, o.w, eh, o.top);
  if (!o.label) return;
  const onSide = o.labelOn === 'side';
  txt(s, o.label, {
    x: o.x + o.w * (onSide ? 0.18 : 0.22),
    y: o.y + (onSide ? eh : eh * 0.22),
    w: o.w * (onSide ? 0.64 : 0.56),
    h: onSide ? d * 0.58 : eh * 0.56,
    fontFace: F_HEAD, bold: true, fontSize: o.labelSize || 14,
    color: o.labelColor, align: 'center', valign: 'middle',
    rotate: o.labelRotate === undefined ? 0 : o.labelRotate,
  });
}

/** Isometric box / cube: top rhombus over left and right side faces. */
function isoBox(s, o) {
  const r = o.ratio === undefined ? ISO : o.ratio;
  const rh = o.w * r;
  const d = o.depth;
  const x = o.x;
  const y = o.y;
  const cx = x + o.w / 2;
  if (o.shadow) shadow(s, cx, y + rh + d, o.w * 1.1);
  poly(s, [[x, y + rh / 2], [cx, y + rh], [cx, y + rh + d], [x, y + rh / 2 + d]], o.left);
  poly(s, [[cx, y + rh], [x + o.w, y + rh / 2], [x + o.w, y + rh / 2 + d],
    [cx, y + rh + d]], o.right);
  poly(s, [[cx, y], [x + o.w, y + rh / 2], [cx, y + rh], [x, y + rh / 2]], o.top);
  if (o.label) {
    txt(s, o.label, {
      x: cx - o.w * 0.22, y: y + rh * 0.20, w: o.w * 0.44, h: rh * 0.55,
      fontFace: F_HEAD, bold: true, fontSize: o.labelSize || 20,
      color: o.labelColor, align: 'center', valign: 'middle',
      rotate: o.labelRotate === undefined ? 330 : o.labelRotate,
    });
  }
}

/** Flat isometric rhombus (a screen, a mat, a card face lying down). */
function isoPlane(s, x, y, w, color, ratio) {
  const rh = w * (ratio === undefined ? ISO : ratio);
  poly(s, [[x + w / 2, y], [x + w, y + rh / 2], [x + w / 2, y + rh], [x, y + rh / 2]], color);
}

/**
 * Isometric jigsaw piece. Each edge is generated in unit-square space as a
 * semicircular tab (or notch) at its midpoint; the outline is then projected
 * into the isometric rhombus and the two lower edges are extruded downwards.
 */
function isoPuzzle(s, o) {
  const w = o.w;
  const rh = w * ISO;
  const d = o.depth;
  const project = (p) => [o.x + w / 2 + (p[0] - p[1]) * w / 2, o.y + (p[0] + p[1]) * rh / 2];
  const R = 0.15;   // tab radius as a fraction of the edge length
  /** Edge a->b with a semicircular bump of radius R along direction n. */
  function edge(a, b, n) {
    const dir = [b[0] - a[0], b[1] - a[1]];
    const mid = [a[0] + dir[0] * 0.5, a[1] + dir[1] * 0.5];
    const pts = [a];
    for (let i = 0; i <= 10; i += 1) {
      const t = (Math.PI * i) / 10;
      pts.push([
        mid[0] - Math.cos(t) * dir[0] * R + n[0] * Math.sin(t) * R,
        mid[1] - Math.cos(t) * dir[1] * R + n[1] * Math.sin(t) * R,
      ]);
    }
    return pts;
  }
  const NW = [0, 0];
  const NE = [1, 0];
  const SE = [1, 1];
  const SW = [0, 1];
  const eNE = edge(NW, NE, [0, -1]).map(project);  // tab out of the upper-right edge
  const eSE = edge(NE, SE, [1, 0]).map(project);   // tab out of the lower-right edge
  const eSW = edge(SE, SW, [0, -1]).map(project);  // notch cut into the lower-left edge
  const eNW = edge(SW, NW, [1, 0]).map(project);   // notch cut into the upper-left edge
  const outline = eNE.concat(eSE, eSW, eNW);
  const band = (pts, color) => poly(s,
    pts.concat(pts.slice().reverse().map((q) => [q[0], q[1] + d])), color);
  poly(s, outline.map((q) => [q[0], q[1] + d]), o.left);   // extruded under-body
  band(eSE, o.right);
  band(eSW, o.left);
  poly(s, outline, o.top);
}

/** Isometric screen standing up: a parallelogram leaning back to the right. */
function isoScreen(s, x, y, w, h, color) {
  const rise = w * ISO;
  poly(s, [[x, y], [x + w, y + rise], [x + w, y + rise + h], [x, y + h]], color);
}

/**
 * Isometric person seen three-quarters from behind, as everywhere in this deck:
 * hair cap, head, tapered torso with a shaded far arm, two legs and two shoes.
 * `h` is the total height; the body is about 0.30 h wide.
 */
function person(s, o) {
  const h = o.h;
  const w = h * 0.30;
  const x = o.x;
  const y = o.y;
  const skin = o.skin || SKIN;
  const hair = o.hair || HAIR;
  const shirt = o.shirt || SUIT;
  const pants = o.pants || SUIT_D;
  const shoe = o.shoe || '231F20';
  const shirtD = shade(shirt);
  const pantsD = shade(pants);
  // legs, offset slightly in depth so the far leg reads as behind
  rrect(s, x + w * 0.18, y + h * 0.53, w * 0.26, h * 0.40, pantsD, w * 0.12);
  rrect(s, x + w * 0.52, y + h * 0.55, w * 0.28, h * 0.39, pants, w * 0.12);
  ell(s, x + w * 0.10, y + h * 0.90, w * 0.40, h * 0.055, shade(shoe));
  ell(s, x + w * 0.48, y + h * 0.92, w * 0.42, h * 0.055, shoe);
  // torso: shoulders down to a narrower waist
  poly(s, [[x + w * 0.20, y + h * 0.26], [x + w * 0.80, y + h * 0.26],
    [x + w * 0.86, y + h * 0.40], [x + w * 0.80, y + h * 0.60],
    [x + w * 0.20, y + h * 0.60], [x + w * 0.14, y + h * 0.40]], shirt);
  rrect(s, x + w * 0.20, y + h * 0.23, w * 0.60, h * 0.12, shirt, w * 0.16);
  if (o.armsUp) {
    poly(s, [[x + w * 0.18, y + h * 0.28], [x - w * 0.10, y + h * 0.01],
      [x + w * 0.04, y - h * 0.04], [x + w * 0.32, y + h * 0.25]], shirtD);
    poly(s, [[x + w * 0.82, y + h * 0.28], [x + w * 1.10, y + h * 0.01],
      [x + w * 0.96, y - h * 0.04], [x + w * 0.68, y + h * 0.25]], shirt);
    ell(s, x - w * 0.14, y - h * 0.07, w * 0.20, h * 0.06, skin);
    ell(s, x + w * 0.94, y - h * 0.07, w * 0.20, h * 0.06, skin);
  } else {
    rrect(s, x + w * 0.02, y + h * 0.28, w * 0.16, h * 0.26, shirtD, w * 0.09);
    rrect(s, x + w * 0.82, y + h * 0.28, w * 0.16, h * 0.26, shirt, w * 0.09);
    ell(s, x, y + h * 0.51, w * 0.19, h * 0.06, skin);
    ell(s, x + w * 0.81, y + h * 0.51, w * 0.19, h * 0.06, skin);
  }
  // neck, head and hair cap
  rect(s, x + w * 0.40, y + h * 0.19, w * 0.20, h * 0.06, skin);
  ell(s, x + w * 0.28, y + h * 0.05, w * 0.44, h * 0.17, skin);
  poly(s, [[x + w * 0.27, y + h * 0.14], [x + w * 0.27, y + h * 0.07],
    [x + w * 0.36, y + h * 0.02], [x + w * 0.64, y + h * 0.02],
    [x + w * 0.73, y + h * 0.07], [x + w * 0.73, y + h * 0.12],
    [x + w * 0.62, y + h * 0.09], [x + w * 0.36, y + h * 0.09]], hair);
}

/* --------------------------------------------------------- text captions */

/**
 * Left caption block used by 17 of the 18 slides: two 36 pt extrabold blue
 * headline lines over a 12 pt white lorem paragraph.
 */
function caption(s, c) {
  const x = c.x === undefined ? 1.146 : c.x;
  txt(s, c.line1, {
    x, y: c.y1 === undefined ? 2.323 : c.y1, w: 4.141, h: 0.485,
    fontFace: F_HEAD, bold: true, fontSize: c.size1 || 36, color: TITLE_BLUE,
  });
  txt(s, c.line2, {
    x, y: 2.808, w: c.w2 || 4.141, h: 0.485,
    fontFace: F_HEAD, bold: true, fontSize: 36, color: TITLE_BLUE,
  });
  txt(s, c.body === undefined ? LOREM_FULL : c.body, {
    x, y: 3.526, w: c.bodyW || 3.518, h: 2.215,
    fontSize: 12, lineSpacing: 15.5,
  });
}

/* --------------------------------------------------------- shared props */

/**
 * Gold cup: tapered bowl, two ring handles, stem and a two-tier base.
 * Coordinates are fractions of the 1.13 x 1.81 in cup on slide 1, scaled by `w`.
 */
function trophy(s, x, y, w) {
  const u = w / 1.13;
  const px = (v) => x + v * u;
  const py = (v) => y + v * u;
  ring(s, px(-0.02), py(0.32), 0.50 * u, 0.52 * u, 'FF9500', 6 * u);   // left handle
  ring(s, px(0.66), py(0.44), 0.48 * u, 0.52 * u, 'FF9500', 6 * u);    // right handle
  poly(s, [[px(0.08), py(0.31)], [px(1.02), py(0.31)], [px(0.68), py(1.14)],
    [px(0.42), py(1.14)]], 'FFE600');                                  // bowl walls
  ell(s, px(0.42), py(1.02), 0.26 * u, 0.16 * u, 'FFE600');
  ell(s, px(0.08), py(0.00), 0.94 * u, 0.62 * u, 'FFC800');            // rim
  ell(s, px(0.14), py(0.06), 0.82 * u, 0.50 * u, 'FFDD00');
  rect(s, px(0.48), py(1.08), 0.14 * u, 0.42 * u, 'FFDD00');           // stem
  ell(s, px(0.20), py(1.31), 0.69 * u, 0.42 * u, 'FFE600');            // upper base
  rect(s, px(0.20), py(1.50), 0.69 * u, 0.10 * u, 'FF9500');
  ell(s, px(0.20), py(1.49), 0.69 * u, 0.22 * u, 'FF9500');
  ell(s, px(0.14), py(1.36), 0.81 * u, 0.46 * u, 'FFDD00');            // lower base
  rect(s, px(0.14), py(1.58), 0.81 * u, 0.12 * u, 'FF9500');
  ell(s, px(0.14), py(1.57), 0.81 * u, 0.26 * u, 'FF9500');
}

/**
 * Extruded zig-zag arrow rising to the right, as used on slides 7 and 13.
 * `pts` are the centre-line vertices of the face; `t` is the ribbon thickness.
 */
function isoArrow(s, o) {
  poly(s, o.face, o.top);
  poly(s, o.side, o.dark);
  poly(s, o.head, o.top);
}

/** Stack of gold coins, drawn bottom-up so each coin overlaps the one below. */
function coinStack(s, x, y, w, n, step) {
  const eh = w * ISOE;
  for (let i = 0; i < n; i += 1) {
    const cx = x + (i % 2) * w * 0.06;
    const cy = y + (n - 1 - i) * step;
    ell(s, cx, cy + step * 0.8, w, eh, 'FFA942');
    rect(s, cx, cy + eh / 2, w, step * 0.8, 'FFC14A');
    ell(s, cx, cy, w, eh, 'FFF459');
  }
}

/** Dark coffee mug with a ring handle. */
function mug(s, x, y, w) {
  const d = w * 0.85;
  ring(s, x + w * 0.82, y + d * 0.42, w * 0.42, d * 0.62, '414042', 5);
  isoDisc(s, { x, y, w, depth: d, top: '414042', side: '242E3D', shadow: false });
  ell(s, x + w * 0.09, y + w * ISO * 0.16, w * 0.82, w * ISO * 0.7, '735033');
}

/** Round magnifier with a tilted handle. */
function magnifier(s, x, y, w) {
  ring(s, x, y, w, w * 0.62, '0757BD', 6);
  ell(s, x + w * 0.08, y + w * 0.05, w * 0.84, w * 0.5, '84ECFE');
  poly(s, [[x + w * 0.03, y + w * 0.55], [x + w * 0.30, y + w * 0.42],
    [x + w * 0.45, y + w * 0.72], [x + w * 0.16, y + w * 0.86]], '0757BD');
}

/** Pot plant: four tapered leaves fanning out of an isometric pot. */
function plant(s, x, y, w, leafH) {
  [[-0.34, 0.72, '00BDA0'], [-0.13, 0.96, '01BBAE'],
    [0.12, 1.00, '85D1CF'], [0.34, 0.78, '01BBAE']].forEach((l) => {
    const lx = x + w * (0.5 + l[0] * 0.9);
    const tip = y - leafH * l[1];
    poly(s, [[lx, tip], [lx + w * 0.30 * Math.sign(l[0] || 1), tip + leafH * l[1] * 0.45],
      [x + w * 0.5, y + w * 0.10], [lx - w * 0.10 * Math.sign(l[0] || 1),
        tip + leafH * l[1] * 0.40]], l[2]);
  });
  isoDisc(s, { x, y, w, depth: w * 0.85, top: '096CEB', side: '0757BD' });
}

/** Stack of banknotes with an oval portrait printed on top. */
function banknotes(s, x, y, w, n) {
  for (let i = 0; i < n; i += 1) {
    isoBox(s, {
      x: x - i * 0.01, y: y - i * 0.16, w, ratio: 0.5, depth: 0.11,
      top: 'F1F2F2', left: 'E2E8F1', right: 'C3CFE2',
    });
  }
  ell(s, x + w * 0.28, y - (n - 1) * 0.16 + w * 0.16, w * 0.44, w * 0.24, 'FF8A39');
}

/**
 * Grid of isometric keys, used on the laptop slides. (x, y) is the top corner
 * of the grid; `k` is the pitch between neighbouring keys along one axis.
 */
function keyboard(s, x, y, k, cols, rows, color) {
  const kh = k * ISO;
  for (let r = 0; r < rows; r += 1) {
    for (let c = 0; c < cols; c += 1) {
      const kx = x + (c - r) * k * 0.5;
      const ky = y + (c + r) * kh * 0.5;
      poly(s, [[kx, ky], [kx + k * 0.40, ky + kh * 0.40],
        [kx, ky + kh * 0.80], [kx - k * 0.40, ky + kh * 0.40]], color);
    }
  }
}

/** Placeholder standing in for a raster asset in the source deck. */
function imagePlaceholder(s, x, y, w, h, caption_) {
  rect(s, x, y, w, h, '2E3A4C', { line: { color: '46556B', width: 1, dashType: 'dash' } });
  txt(s, caption_ || '[image]', {
    x, y: y + h / 2 - 0.15, w, h: 0.3, align: 'center', color: '8894A6', fontSize: 10,
  });
}

/* =========================================================== slide 1 ==== */
/* 5 STEPS INFOGRAPHIC — a leaning stack of five numbered pucks with a trophy. */

// listed bottom-first, which is also the back-to-front paint order
const STACK5 = [
  { x: 8.15, y: 4.83, top: 'FE2635', side: 'DB212E', rim: 'BF1D28', label: 'STEP 01' },
  { x: 8.60, y: 4.03, top: 'FFF200', side: 'FFD000', rim: 'FFB300', label: 'STEP 02' },
  { x: 8.06, y: 3.17, top: '51F3FE', side: '3ED8F7', rim: '39C6E3', label: 'STEP 03' },
  { x: 8.44, y: 2.40, top: '5271FF', side: '3463C4', rim: '25478C', label: 'STEP 04' },
  { x: 8.00, y: 1.51, top: 'B999F7', side: '7A4DDC', rim: '5C3AA6', label: 'STEP 05' },
];

/** Brown step ladder leaning against something. */
function ladder(s, x, y, h, w, rungs, color, colorD) {
  const lean = h * 0.22;
  poly(s, [[x, y + h], [x + lean, y], [x + lean + w * 0.14, y], [x + w * 0.14, y + h]], colorD);
  poly(s, [[x + w * 0.86, y + h], [x + lean + w, y], [x + lean + w * 0.86, y],
    [x + w, y + h]], color);
  for (let i = 0; i < rungs; i += 1) {
    const t = (i + 0.6) / (rungs + 0.2);
    rect(s, x + lean * (1 - t) + w * 0.10, y + h * t, w * 0.80, h * 0.035, color);
  }
}

function slide01(s) {
  STACK5.forEach((d, i) => {
    isoDisc(s, {
      x: d.x, y: d.y, w: 2.12, depth: 0.66, top: d.top, side: d.side,
      label: d.label, labelColor: d.rim, labelSize: 20, labelRotate: 358, labelOn: 'side',
    });
    if (i === 0) {
      // surveyor and his ladder standing on step 01
      ladder(s, 8.01, 5.83, 0.87, 0.34, 5, '5A3A22', '3C2415');
      person(s, { x: 8.50, y: 5.51, h: 0.91, shirt: '0757BD', pants: '064AA1' });
    }
    if (i === 3) {
      // climber hauling himself up the right-hand side of step 04
      person(s, { x: 10.19, y: 3.42, h: 1.20, shirt: '778498', pants: '637084',
        armsUp: true });
    }
  });
  trophy(s, 8.56, 0.62, 1.13);
  person(s, { x: 8.18, y: 1.45, h: 1.21, shirt: SUIT, pants: SUIT_D, armsUp: true });
  caption(s, { line1: '5 STEPS', line2: 'INFOGRAPHIC', w2: 4.757, bodyW: 4.757 });
}

/* =========================================================== slide 2 ==== */
/* BUSINESS PARTNERSHIP — four people each carrying a jigsaw piece. */

const PUZZLES2 = [
  { piece: { x: 6.32, y: 0.47, top: 'FFF200', left: 'FFBB00', right: 'FFCD00' },
    who: { x: 5.18, y: 0.05, h: 2.08, shirt: '253545', pants: '1B2733', shoe: '8B5E3C' } },
  { piece: { x: 10.36, y: 1.22, top: '8459E0', left: '6E3FD7', right: '7447D9' },
    who: { x: 11.74, y: 0.46, h: 2.57, shirt: 'F53C33', pants: '0757BD', hair: '8B5E3C' } },
  { piece: { x: 4.88, y: 3.18, top: '77CDFF', left: '6CB5FF', right: '73C4FF' },
    who: { x: 4.93, y: 3.82, h: 2.43, shirt: '414042', pants: 'F53C33', hair: 'D1785A' } },
  { piece: { x: 8.68, y: 4.08, top: 'FE2635', left: 'DB212E', right: 'ED2531' },
    who: { x: 10.41, y: 5.03, h: 2.61, shirt: 'C4C3EF', pants: '3A3079' } },
];

function slide02(s) {
  PUZZLES2.forEach((g) => {
    isoPuzzle(s, Object.assign({ w: 2.50, depth: 0.24 }, g.piece));
    person(s, g.who);
  });
  caption(s, { line1: 'BUSINESS', line2: 'PARTNERSHIP' });
}

/* =========================================================== slide 3 ==== */
/* 4 STEPS INFOGRAPHIC — four big isometric cubes of increasing height. */

const CUBES3 = [
  { x: 6.16, y: 2.86, depth: 1.37, top: '77CDFF', left: '73C4FF', right: '6CB5FF', label: '02' },
  { x: 7.77, y: 1.52, depth: 1.77, top: 'FFF200', left: 'FFCD00', right: 'FFC400', label: '03' },
  { x: 9.27, y: 1.98, depth: 2.18, top: '8459E0', left: '7447D9', right: '6E3FD7', label: '04' },
  { x: 7.65, y: 4.14, depth: 0.96, top: 'FE2635', left: 'FA2635', right: 'DE212E', label: '01' },
];

function slide03(s) {
  CUBES3.forEach((c) => {
    isoBox(s, {
      x: c.x, y: c.y, w: 3.11, depth: c.depth, top: c.top, left: c.left, right: c.right,
      label: c.label, labelColor: c.left, labelSize: 44, labelRotate: 345,
    });
  });
  ladder(s, 11.46, 3.21, 2.41, 0.50, 9, '2C2829', HAIR);
  person(s, { x: 11.49, y: 3.77, h: 0.89, shirt: '4DC7FF', pants: 'DEE9F9' });
  person(s, { x: 9.65, y: 4.95, h: 0.90, shirt: '0757BD', pants: '064AA1' });
  person(s, { x: 10.48, y: 1.64, h: 1.28, shirt: SUIT, pants: SUIT_D, armsUp: true });
  person(s, { x: 9.38, y: 1.45, h: 1.04, shirt: '3A393B', pants: '2E2D2E' });
  person(s, { x: 9.06, y: 1.36, h: 1.04, shirt: '0757BD', pants: '053F8A' });
  person(s, { x: 7.47, y: 2.87, h: 1.04, shirt: '3E454E', pants: '343B42' });
  caption(s, { line1: '4 STEPS', line2: 'INFOGRAPHIC', w2: 4.757, bodyW: 4.757 });
}

/* =========================================================== slide 4 ==== */
/* 4 STEPS INFOGRAPHIC — ascending coloured slabs with a white staircase. */

const SLABS4 = [
  { x: 5.46, y: 4.36, top: '4D6AEA', left: '3060C6', right: '064BA2', label: '01' },
  { x: 6.39, y: 3.42, top: '7DDCFE', left: '77CDFF', right: '66A6FF', label: '02' },
  { x: 7.33, y: 2.49, top: 'FFDA00', left: 'FFCF00', right: 'FFB700', label: '03' },
  { x: 8.27, y: 1.55, top: 'F64336', left: 'FF493D', right: 'E32D1F', label: '04' },
];

function slide04(s) {
  SLABS4.forEach((b, i) => {
    isoBox(s, {
      x: b.x, y: b.y, w: 4.68, ratio: 0.5, depth: 0.23,
      top: b.top, left: b.left, right: b.right,
      label: b.label, labelColor: b.right, labelSize: 26, labelRotate: 331,
    });
    // white stair tread running up the centre of each slab
    isoBox(s, {
      x: 6.39 + i * 0.94, y: 5.53 - i * 0.94, w: 1.41, ratio: 0.5, depth: 0.24,
      top: 'EEF5FF', left: 'E2EEFF', right: 'E2EEFF',
    });
  });
  // arrow head at the top of the staircase
  poly(s, [[10.44, 2.57], [9.21, 3.19], [9.68, 3.42], [10.91, 2.80]], 'EEF5FF');
  poly(s, [[11.00, 2.41], [10.68, 1.68], [10.35, 2.09], [10.44, 2.13],
    [10.44, 2.57], [10.91, 2.80], [10.91, 2.36]], 'E2EEFF');
  person(s, { x: 9.37, y: 2.01, h: 1.23, shirt: SUIT, pants: SUIT_D, armsUp: true });
  person(s, { x: 8.99, y: 3.16, h: 1.04, shirt: '3A393B', pants: '2E2D2E' });
  person(s, { x: 8.67, y: 3.08, h: 1.04, shirt: '0757BD', pants: '053F8A' });
  person(s, { x: 7.78, y: 3.94, h: 1.16, shirt: '3E454E', pants: '343B42' });
  person(s, { x: 6.45, y: 5.64, h: 0.91, shirt: '84ECFE', pants: 'FF493D' });
  caption(s, { x: 0.66, line1: '4 STEPS', line2: 'INFOGRAPHIC', w2: 4.757, bodyW: 4.757 });
}

/* =========================================================== slide 5 ==== */
/* BUSINESS PARTNERSHIP — two businessmen shaking hands across two pieces. */

function slide05(s) {
  isoPuzzle(s, {
    x: 6.38, y: 3.28, w: 3.82, depth: 0.34,
    top: '81E5FE', left: '73C4FF', right: '6CB5FF',
  });
  isoPuzzle(s, {
    x: 8.29, y: 4.38, w: 3.82, depth: 0.34,
    top: 'FF4538', left: 'E03E31', right: 'DB212E',
  });
  person(s, { x: 7.93, y: 0.62, h: 4.45, shirt: '0757BD', pants: '053F8A', shoe: '2E2D2E' });
  person(s, { x: 9.40, y: 1.24, h: 4.43, shirt: '3A393B', pants: '2E2D2E' });
  // the two outstretched hands meeting in the middle
  ell(s, 8.86, 2.86, 0.46, 0.22, SKIN);
  ell(s, 9.18, 2.98, 0.46, 0.22, SKIN_D);
  rect(s, 8.62, 1.66, 0.13, 0.60, 'DB212E');   // red tie on the blue suit
  rrect(s, 9.94, 3.72, 0.81, 0.62, '8B5E3C', 0.05); // briefcase
  rect(s, 9.94, 3.94, 0.81, 0.08, '6E4529');
  rect(s, 10.24, 3.66, 0.20, 0.10, '6E4529');
  caption(s, { line1: 'BUSINESS', line2: 'PARTNERSHIP', bodyW: 4.499 });
}

/* =========================================================== slide 6 ==== */
/* 5 OPTIONS INFOGRAPHIC — a centred headline over a row of five pucks. */

const OPTIONS6 = [
  { x: 1.90, y: 4.27, top: 'FE3341', side: 'ED303D', rim: 'DE2C37', label: '01' },
  { x: 3.91, y: 4.25, top: 'FFF200', side: 'FFD000', rim: 'FFB300', label: '02' },
  { x: 5.93, y: 4.32, top: '51F3FE', side: '3ED8F7', rim: '39C6E3', label: '03' },
  { x: 7.89, y: 4.32, top: '5271FF', side: '3463C4', rim: '25478C', label: '04' },
  { x: 9.70, y: 5.13, top: 'A074F7', side: '7A4DDC', rim: '5C3AA6', label: '05' },
];

function slide06(s) {
  txt(s, '5 OPTIONS', {
    x: 5.394, y: 0.380, w: 3.003, h: 0.823,
    fontFace: F_HEAD, bold: true, fontSize: 36, color: TITLE_BLUE6,
  });
  txt(s, 'INFOGRAPHIC', {
    x: 4.986, y: 0.852, w: 3.811, h: 0.821,
    fontFace: F_HEAD, bold: true, fontSize: 36, color: TITLE_BLUE6,
  });
  txt(s, LOREM_TINY, {
    x: 4.045, y: 1.224, w: 5.395, h: 1.175,
    fontSize: 12, align: 'center', valign: 'middle', lineSpacing: 15.5,
  });
  OPTIONS6.forEach((d) => {
    isoDisc(s, {
      x: d.x, y: d.y, w: 1.51, depth: 0.62, top: d.top, side: d.side, shadow: true,
      label: d.label, labelColor: d.rim, labelSize: 24, labelRotate: 330,
    });
  });
  person(s, { x: 2.44, y: 3.63, h: 1.29, shirt: '3061C2', pants: '064FAB', hair: '414042' });
  person(s, { x: 4.27, y: 4.54, h: 1.27, shirt: '84ECFE', pants: 'FF493D' });
  person(s, { x: 6.39, y: 3.46, h: 1.53, shirt: '3061C2', pants: 'FE2635', armsUp: true });
  person(s, { x: 8.25, y: 3.67, h: 1.19, shirt: '717E91', pants: '647286' });
  person(s, { x: 8.62, y: 3.77, h: 1.18, shirt: '3A393B', pants: '2E2D2E' });
  person(s, { x: 9.38, y: 6.12, h: 1.33, shirt: 'E6E7E8', pants: '2A292B' });
}

/* =========================================================== slide 7 ==== */
/* INVESTMENT SOLUTIONS — rising arrow, bar chart, banknotes, coins, people. */

const BARS7 = [
  { x: 7.33, y: 4.45, depth: 1.30, top: '83E9FB', left: '77CDFF', right: '66A6FF' },
  { x: 8.07, y: 3.40, depth: 1.94, top: '0763D9', left: '075CC9', right: '0757BD' },
  { x: 8.80, y: 2.06, depth: 2.86, top: 'CF09F2', left: 'AA09DA', right: '8408C2' },
];

function slide07(s) {
  // big orange growth arrow behind the chart
  poly(s, [[9.22, 1.55], [10.04, 0.72], [9.89, 1.81], [9.67, 1.72], [8.22, 4.00],
    [7.59, 3.56], [6.60, 6.06], [6.35, 6.03], [7.54, 3.02], [8.25, 3.51],
    [9.44, 1.63]], 'FF3419');
  poly(s, [[6.35, 6.04], [6.17, 5.94], [7.36, 2.92], [7.53, 3.02]], 'FF5627');
  poly(s, [[8.25, 3.51], [8.07, 3.41], [9.25, 1.53], [9.43, 1.63]], 'FF5627');
  poly(s, [[9.22, 1.55], [9.04, 1.45], [9.87, 0.62], [10.04, 0.72]], 'FF5627');
  BARS7.forEach((b) => isoBox(s, Object.assign({ w: 1.48 }, b)));
  banknotes(s, 8.36, 3.08, 3.43, 2);
  coinStack(s, 6.66, 4.86, 1.45, 3, 0.15);
  person(s, { x: 7.71, y: 1.57, h: 3.38, shirt: '464C56', pants: '394149' });
  person(s, { x: 11.16, y: 3.90, h: 3.03, shirt: 'FE2635', pants: 'FE2635', hair: HAIR });
  magnifier(s, 9.86, 5.48, 1.18);
  isoDisc(s, { x: 8.98, y: 6.15, w: 0.82, depth: 0.25, top: '0757BD', side: '053D85' });
  caption(s, { line1: 'INVESTMENT', line2: 'SOLUTIONS', bodyW: 4.499 });
}

/* =========================================================== slide 8 ==== */
/* MARKETING SERVICES — laptop, megaphone, plant, phone, notebook, mug. */

function slide08(s) {
  // laptop screen: dark bezel with a pink UI mockup inside
  poly(s, [[8.57, 0.68], [11.94, 1.47], [11.94, 5.05], [8.57, 4.26]], '414042');
  poly(s, [[8.73, 0.93], [11.71, 1.62], [11.71, 4.74], [8.73, 4.05]], 'FDC4C7');
  poly(s, [[9.07, 2.92], [11.41, 3.46], [11.41, 4.75], [9.07, 4.21]], 'FDDFE0');
  poly(s, [[8.95, 1.31], [10.16, 1.59], [10.16, 1.75], [8.95, 1.47]], '0757BD');
  [0, 1, 2, 3].forEach((i) => {
    poly(s, [[8.95, 1.66 + i * 0.20], [10.10, 1.92 + i * 0.20],
      [10.10, 2.02 + i * 0.20], [8.95, 1.76 + i * 0.20]], '8691A6');
  });
  // mini bar chart in the corner of the screen
  [['0757BD', 0.60], ['FE2635', 0.38], ['84ECFE', 0.50]].forEach((b, i) => {
    poly(s, [[10.55 + i * 0.24, 2.13 - b[1]], [10.72 + i * 0.24, 2.17 - b[1]],
      [10.72 + i * 0.24, 2.17], [10.55 + i * 0.24, 2.13]], b[0]);
  });
  // plant behind the screen
  plant(s, 7.40, 2.14, 1.09, 1.35);
  // laptop base with keyboard and trackpad
  isoBox(s, { x: 6.23, y: 3.04, w: 5.62, ratio: 0.58, depth: 0.20,
    top: '414042', left: '2E3644', right: '242E3D' });
  keyboard(s, 9.10, 3.80, 0.42, 11, 5, '58595B');
  isoPlane(s, 7.05, 4.10, 1.90, '58595B');
  // megaphone bursting out of the screen
  poly(s, [[8.90, 2.66], [9.96, 3.05], [9.96, 3.73], [8.90, 3.34]], GY75);
  poly(s, [[9.30, 2.64], [10.07, 2.94], [10.07, 3.52], [9.30, 3.22]], 'E92C43');
  poly(s, [[7.90, 2.96], [9.43, 3.52], [9.43, 4.60], [7.90, 4.04]], 'E92C43');
  poly(s, [[7.90, 3.09], [8.87, 3.44], [8.87, 4.60], [7.90, 4.25]], 'FF455B');
  ell(s, 7.47, 2.97, 0.86, 1.97, 'C22538');
  ell(s, 7.59, 3.16, 0.62, 1.70, '941D2A');
  poly(s, [[8.15, 3.41], [8.83, 3.66], [8.83, 4.02], [8.15, 3.77]], 'E92C43');
  rect(s, 9.45, 3.50, 0.25, 1.00, 'E92C43');
  ell(s, 9.44, 3.42, 0.28, 0.16, 'E92C43');
  ell(s, 8.13, 3.77, 0.20, 0.26, '231F20');
  // phone with a video thumbnail and a "like" bubble above it
  isoBox(s, { x: 5.10, y: 4.99, w: 1.85, ratio: 0.58, depth: 0.09,
    top: '3F3D3F', left: '231F20', right: '6D6E71' });
  isoPlane(s, 5.16, 5.02, 1.73, GY50);
  poly(s, [[5.54, 5.22], [6.56, 5.51], [6.10, 5.81], [5.08, 5.52]], GY90);
  poly(s, [[5.66, 5.43], [6.14, 5.55], [5.82, 5.72]], 'FE2635');
  [0, 1, 2, 3].forEach((i) => {
    poly(s, [[6.17 + i * 0.125, 5.13 + i * 0.065], [6.23 + i * 0.125, 5.16 + i * 0.065],
      [6.29 + i * 0.125, 5.20 + i * 0.065], [6.23 + i * 0.125, 5.17 + i * 0.065]], GY90);
  });
  rrect(s, 5.65, 4.21, 0.64, 0.55, 'E92C43', 0.12);
  poly(s, [[5.86, 4.70], [6.08, 4.70], [5.92, 4.99]], 'CC273B');
  s.addShape('heart', { x: 5.72, y: 4.34, w: 0.20, h: 0.18,
    fill: { color: WHITE }, line: NOLINE });
  txt(s, '5', { x: 5.92, y: 4.31, w: 0.30, h: 0.28, align: 'center',
    fontSize: 11, bold: true });
  // notebook with a play badge, mug and magnifier
  isoBox(s, { x: 6.80, y: 5.66, w: 2.13, ratio: 0.5, depth: 0.16,
    top: '76CAFF', left: '66A6FF', right: '4280D6' });
  ell(s, 7.55, 5.90, 0.62, 0.38, WHITE);
  poly(s, [[7.74, 5.97], [7.98, 6.09], [7.74, 6.21]], '276BCC');
  mug(s, 11.63, 4.20, 1.03);
  magnifier(s, 9.93, 5.68, 1.20);
  caption(s, { line1: 'MARKETING', line2: 'SERVICES', body: LOREM_SHORT });
}

/* =========================================================== slide 9 ==== */
/* BUSINESS PROCESS — a man on an office chair at a desk with a monitor. */

function slide09(s) {
  // desk: legs, top rhombus and front edge band
  [[6.98, 2.83], [8.28, 2.09], [10.35, 4.80], [11.65, 4.03]].forEach((l) => {
    rect(s, l[0], l[1], 0.15, 2.18, GY75);
  });
  isoPlane(s, 6.28, 1.63, 6.24, GY50);
  poly(s, [[6.28, 3.44], [9.40, 5.24], [12.52, 3.44], [12.52, 3.62],
    [9.40, 5.42], [6.28, 3.62]], GY75);
  // monitor: casing, screen and dashboard content
  poly(s, [[8.97, 0.61], [11.19, 1.13], [11.19, 3.39], [8.97, 2.87]], GY75);
  poly(s, [[9.06, 0.80], [11.09, 1.27], [11.09, 3.19], [9.06, 2.72]], '7DDBFE');
  poly(s, [[9.06, 0.80], [11.09, 1.27], [11.09, 1.74], [9.06, 1.27]], '3665C6');
  ell(s, 10.72, 1.36, 0.22, 0.22, WHITE);
  [0, 1, 2, 3].forEach((i) => {
    poly(s, [[9.20, 1.88 + i * 0.24], [10.30, 2.14 + i * 0.24],
      [10.30, 2.24 + i * 0.24], [9.20, 1.98 + i * 0.24]], WHITE);
  });
  s.addShape('pie', { x: 10.44, y: 2.02, w: 0.62, h: 0.60, angleRange: [0, 250],
    fill: { color: '00B300' }, line: NOLINE });
  s.addShape('pie', { x: 10.44, y: 2.02, w: 0.62, h: 0.60, angleRange: [250, 360],
    fill: { color: 'C525ED' }, line: NOLINE });
  [0.10, 0.20, 0.30].forEach((hh, i) => {
    rect(s, 10.52 + i * 0.14, 2.98 - hh, 0.09, hh, WHITE);
  });
  // monitor stand
  rect(s, 9.94, 3.19, 0.36, 0.30, GY75);
  isoPlane(s, 9.58, 3.24, 1.10, GY75);
  // keyboard, mouse mat, phone, pen cup and books on the desk
  isoPlane(s, 8.16, 2.99, 2.01, GY75);
  keyboard(s, 9.18, 3.10, 0.22, 10, 4, GY90);
  isoPlane(s, 10.14, 4.33, 0.74, '6E68E5');
  isoPlane(s, 11.14, 4.18, 0.39, 'A3BDED');
  isoDisc(s, { x: 10.66, y: 3.66, w: 0.43, depth: 0.42, top: GY75, side: GY90 });
  [['FE2635', 0], ['FFD200', 1], ['00CEE8', 2], ['00B300', 3]].forEach((c) => {
    rect(s, 10.72 + c[1] * 0.075, 3.51 - c[1] * 0.02, 0.05, 0.26, c[0]);
  });
  isoBox(s, { x: 7.82, y: 1.94, w: 0.99, ratio: 0.5, depth: 0.11,
    top: '345CB4', left: '2F53A3', right: '667EEA' });
  isoBox(s, { x: 7.91, y: 1.71, w: 0.99, ratio: 0.5, depth: 0.11,
    top: 'F04C45', left: 'DF353A', right: 'FF8A46' });
  // the seated worker, drawn between the chair back and the chair seat
  person(s, { x: 7.31, y: 2.30, h: 2.90, shirt: '076BBD', pants: '0765B3' });
  // office chair: back rest, seat, column and star base
  poly(s, [[7.01, 3.15], [7.68, 2.94], [8.05, 5.20], [7.34, 5.44]], 'D4202C');
  poly(s, [[7.21, 2.99], [7.85, 2.80], [8.22, 5.05], [7.55, 5.28]], 'FE2635');
  isoBox(s, { x: 6.94, y: 4.28, w: 2.07, ratio: 0.5, depth: 0.13,
    top: 'FE2635', left: 'D4202C', right: 'B41A24' });
  rect(s, 7.92, 5.05, 0.11, 0.72, 'DBE2EE');
  for (let i = 0; i < 5; i += 1) {
    const ang = (i / 5) * Math.PI * 2 + 0.4;
    poly(s, [[7.92, 5.72], [7.97 + Math.cos(ang) * 0.62, 5.76 + Math.sin(ang) * 0.33],
      [7.97 + Math.cos(ang) * 0.62, 5.84 + Math.sin(ang) * 0.33], [7.92, 5.80]], 'DBE2EE');
    ell(s, 7.90 + Math.cos(ang) * 0.62, 5.78 + Math.sin(ang) * 0.33, 0.16, 0.13, GY90);
  }
  caption(s, { line1: 'BUSINESS', line2: 'PROCESS', body: LOREM_SHORT });
}

/* ========================================================== slide 10 ==== */
/* RECRUITMENT SERVICES — CV cards floating around a phone. */

const CARDS10 = [
  { x: 7.77, y: 1.28, av: '0757BD', stars: 4, bg: '2E3A4C' },
  { x: 10.00, y: 2.58, av: 'FE2635', stars: 3, bg: '2E3A4C' },
];

/** One floating CV card: avatar, star rating and three text rules. */
function cvCard(s, c) {
  poly(s, [[c.x, c.y], [c.x + 1.48, c.y + 0.34], [c.x + 1.48, c.y + 3.07],
    [c.x, c.y + 2.73]], c.bg);
  ell(s, c.x + 0.30, c.y + 0.58, 0.88, 0.86, 'B1D6FC');
  ell(s, c.x + 0.48, c.y + 0.74, 0.52, 0.48, c.av);
  ell(s, c.x + 0.42, c.y + 1.16, 0.64, 0.30, c.av);
  for (let i = 0; i < 5; i += 1) {
    s.addShape('star5', {
      x: c.x + 0.18 + i * 0.24, y: c.y + 1.58 + i * 0.05, w: 0.21, h: 0.21,
      fill: { color: i < c.stars ? 'FFD145' : '5A6577' }, line: NOLINE,
    });
  }
  [0, 1, 2].forEach((i) => {
    poly(s, [[c.x + 0.16, c.y + 2.00 + i * 0.26], [c.x + 1.26, c.y + 2.25 + i * 0.26],
      [c.x + 1.26, c.y + 2.35 + i * 0.26], [c.x + 0.16, c.y + 2.10 + i * 0.26]], '4A5568');
  });
}

function slide10(s) {
  // the two flanking CV cards, painted before the phone so it overlaps them
  CARDS10.forEach((c) => cvCard(s, c));
  // phone standing in the middle
  poly(s, [[8.78, 1.01], [10.58, 1.42], [10.58, 5.45], [8.78, 5.04]], '414042');
  poly(s, [[8.81, 1.05], [10.39, 1.41], [10.39, 5.41], [8.81, 5.05]], '000000');
  poly(s, [[8.90, 1.25], [10.31, 1.57], [10.31, 5.15], [8.90, 4.83]], '78D1FE');
  cvCard(s, { x: 8.85, y: 1.56, av: '414042', stars: 5, bg: '3B4757' });
  // recruiter reaching for the highlighted card
  person(s, { x: 7.23, y: 3.54, h: 3.26, shirt: '414042', pants: '58595B' });
  // rolled-up poster lying under the phone
  isoDisc(s, { x: 8.31, y: 4.72, w: 0.72, depth: 0.42, top: '7BD6FE', side: '50C1F2' });
  poly(s, [[8.67, 4.72], [9.92, 5.13], [9.92, 5.55], [8.67, 5.14]], '4A6E96');
  isoDisc(s, { x: 9.56, y: 5.13, w: 0.49, depth: 0.42, top: '7BD6FE', side: '50C1F2' });
  // small isometric column chart
  [['FF3419', 6.88, 3.34, 0.55], ['2E51B8', 6.60, 3.64, 0.42], ['70E4FE', 6.39, 3.89, 0.28]]
    .forEach((b) => {
      isoDisc(s, { x: b[1], y: b[2], w: 0.42, depth: b[3], top: b[0], side: shade(b[0]) });
    });
  // pie chart of five equal native wedges
  [['FF3419', 200], ['FFE630', 272], ['7C56CF', 344], ['70E4FE', 56], ['3B66E8', 128]]
    .forEach((wedge) => {
      s.addShape('pie', {
        x: 10.27, y: 5.78, w: 1.16, h: 0.70, angleRange: [wedge[1], (wedge[1] + 72) % 360],
        fill: { color: wedge[0] }, line: NOLINE,
      });
    });
  caption(s, { line1: 'RECRUITMENT', line2: 'SERVICES', body: LOREM_SHORT });
}

/* ========================================================== slide 11 ==== */
/* ONLINE PAYMENT — laptop with a credit card sliding into the screen. */

function slide11(s) {
  // laptop lid: dark bezel with a big blue screen
  poly(s, [[7.77, 0.52], [11.61, 1.41], [11.61, 5.40], [7.77, 4.51]], GY75);
  poly(s, [[7.95, 0.79], [11.44, 1.60], [11.44, 5.14], [7.95, 4.33]], '66A6FF');
  // laptop base: top rhombus, front edge, keyboard well and trackpad
  isoPlane(s, 5.07, 3.23, 6.48, GY75);
  poly(s, [[5.07, 5.11], [8.31, 6.98], [11.55, 5.11], [11.55, 5.32],
    [8.31, 7.19], [5.07, 5.32]], GY90);
  isoPlane(s, 6.21, 3.44, 4.98, GY50);
  keyboard(s, 8.10, 3.68, 0.48, 12, 5, GY75);
  isoPlane(s, 6.23, 4.86, 2.51, GY50);
  isoPlane(s, 8.70, 3.70, 1.98, GY50);
  // credit card sliding into a magenta slot on the screen
  poly(s, [[7.72, 2.62], [10.36, 2.62], [10.36, 2.78], [7.72, 2.78]], 'CC2B5E');
  isoBox(s, { x: 7.75, y: 2.85, w: 2.58, depth: 0.08,
    top: '075DC9', left: '053D85', right: '0651A8' });
  rect(s, 8.22, 3.34, 0.32, 0.22, 'FCA954');
  txt(s, 'CREDIT CARD', { x: 8.60, y: 3.06, w: 1.5, h: 0.3, fontSize: 8,
    color: 'CFE0FF', rotate: 340 });
  txt(s, '1234    4567    8910', { x: 8.20, y: 3.62, w: 1.7, h: 0.25, fontSize: 7,
    color: 'CFE0FF', rotate: 340 });
  // props on the desk around the laptop
  plant(s, 6.57, 2.85, 0.85, 1.20);
  banknotes(s, 10.20, 3.77, 2.55, 2);
  mug(s, 11.19, 4.78, 1.03);
  coinStack(s, 10.41, 5.88, 0.84, 3, 0.11);
  poly(s, [[6.16, 6.60], [7.18, 6.02], [7.30, 6.18], [6.28, 6.72]], '83A5D8');
  poly(s, [[7.18, 6.02], [7.30, 6.18], [7.44, 5.98]], 'EADDD3');
  caption(s, { line1: 'ONLINE', line2: 'PAYMENT', body: LOREM_SHORT });
}

/* ========================================================== slide 12 ==== */
/* ONLINE SHOPPING — tablet "floor" with a clothes rail, trolley and boxes. */

function slide12(s) {
  // tablet lying flat: grey frame, dark bezel, blue screen
  isoBox(s, { x: 6.37, y: 3.18, w: 5.59, ratio: 0.58, depth: 0.19,
    top: '231F20', left: '3E3C3E', right: '6D6E71' });
  isoPlane(s, 6.55, 3.28, 5.23, '77CCFF');
  // clothes rail with three garments
  rect(s, 6.51, 0.60, 0.06, 2.90, 'D1DEF0');
  rect(s, 8.33, 1.02, 0.06, 2.40, 'D1DEF0');
  poly(s, [[6.46, 0.56], [8.39, 1.00], [8.39, 1.12], [6.46, 0.68]], 'D1DEF0');
  poly(s, [[6.30, 1.46], [6.98, 1.62], [6.86, 3.20], [6.18, 3.04]], '0757BD');
  poly(s, [[6.78, 1.24], [7.44, 1.39], [7.28, 2.86], [6.62, 2.71]], 'B1D6FC');
  poly(s, [[7.14, 0.96], [7.82, 1.12], [7.66, 2.66], [6.98, 2.50]], 'FE2635');
  // shopping trolley: basket, base tray, wheels
  poly(s, [[7.84, 3.46], [8.90, 3.72], [8.90, 4.82], [7.84, 4.56]], '8FA6BD');
  poly(s, [[8.46, 3.81], [9.52, 4.07], [9.52, 5.17], [8.46, 4.91]], 'B0CCE8');
  poly(s, [[7.85, 4.21], [9.50, 4.63], [9.50, 5.16], [7.85, 4.74]], 'EEF1F7');
  rect(s, 8.60, 4.40, 0.24, 0.60, 'FE2635');
  rect(s, 8.92, 4.48, 0.24, 0.55, 'FFCF60');
  ell(s, 7.90, 5.02, 0.30, 0.30, '231F20');
  ell(s, 9.16, 5.42, 0.30, 0.30, '231F20');
  // stool, cardboard box, football and shopping bag
  isoBox(s, { x: 6.68, y: 3.76, w: 1.10, depth: 0.32,
    top: '0863D6', left: '0753B3', right: '144794' });
  isoBox(s, { x: 8.56, y: 5.14, w: 2.08, depth: 0.72,
    top: 'FFE985', left: 'FFCF60', right: 'FFB53D' });
  rect(s, 9.20, 6.34, 0.74, 0.28, 'F1F2F2');
  for (let i = 0; i < 9; i += 1) rect(s, 9.24 + i * 0.078, 6.38, 0.03, 0.20, '231F20');
  ell(s, 10.23, 5.79, 0.60, 0.60, 'F1F2F2');
  s.addShape('pentagon', { x: 10.38, y: 5.94, w: 0.30, h: 0.28,
    fill: { color: '231F20' }, line: NOLINE });
  poly(s, [[10.35, 4.60], [11.72, 4.90], [11.62, 6.15], [10.44, 6.15]], '0757BD');
  poly(s, [[10.35, 4.60], [11.72, 4.90], [11.72, 5.04], [10.35, 4.74]], '074DA8');
  ring(s, 10.72, 4.32, 0.58, 0.46, '074DA8', 4);
  person(s, { x: 9.18, y: 1.90, h: 3.14, shirt: 'FE2635', pants: 'FE2635', hair: HAIR });
  caption(s, { line1: 'ONLINE', line2: 'SHOPPING' });
}

/* ========================================================== slide 13 ==== */
/* THE ROAD TO SUCCESS RUNS UPHILL — stairs, red arrow, trophy, bar chart. */

function slide13(s) {
  for (let i = 0; i < 6; i += 1) {
    isoBox(s, {
      x: 6.26 + i * 0.51, y: 5.39 - i * 0.51, w: 2.17, depth: 0.34,
      top: GY50, left: GY75, right: GY90,
    });
  }
  isoBox(s, { x: 8.80, y: 2.32, w: 3.06, depth: 0.34,
    top: GY50, left: GY75, right: GY90 });
  // red arrow sweeping up over the stairs
  poly(s, [[8.86, 1.74], [9.57, 1.03], [9.44, 1.96], [9.25, 1.89], [8.01, 3.86],
    [7.46, 3.48], [6.61, 5.63], [6.39, 5.61], [7.42, 3.00], [8.02, 3.43],
    [9.05, 1.81]], 'FF3419');
  poly(s, [[6.40, 5.61], [6.25, 5.52], [7.27, 2.92], [7.42, 3.00]], 'FF5627');
  trophy(s, 9.58, 0.91, 1.63);
  person(s, { x: 6.61, y: 3.96, h: 2.74, shirt: '0A8BD9', pants: '0A80C9' });
  [
    { x: 10.74, y: 4.19, w: 0.71, d: 1.30, top: 'CF09F2', side: 'AA09DA' },
    { x: 9.94, y: 4.98, w: 0.72, d: 0.98, top: '0A8BD9', side: '0A80C9' },
    { x: 10.32, y: 5.03, w: 0.72, d: 0.70, top: '3B66E8', side: '2E51B8' },
    { x: 9.60, y: 5.74, w: 0.72, d: 0.42, top: '83EEFF', side: '70E4FE' },
  ].forEach((b) => {
    isoDisc(s, { x: b.x, y: b.y, w: b.w, depth: b.d, top: b.top, side: b.side,
      shadow: false });
  });
  caption(s, {
    line1: 'THE ROAD TO SUCCESS', line2: 'RUNS UPHILL', size1: 22, y1: 2.473,
  });
}

/* ========================================================== slide 14 ==== */
/* SOCIAL MEDIA MARKETING — a magnet on a phone pulling in reaction icons. */

const REACTIONS = [
  { x: 4.71, y: 5.74, w: 0.75, c: '0760CF', d: '064BA2' },
  { x: 5.15, y: 6.48, w: 0.58, c: 'F9362A', d: 'C90B22' },
  { x: 5.55, y: 5.38, w: 0.67, c: 'FFE400', d: 'FFB700' },
  { x: 5.62, y: 5.99, w: 0.43, c: 'AC09DC', d: '7A06B3' },
  { x: 6.12, y: 6.11, w: 0.69, c: 'FFA538', d: 'FF7E38' },
  { x: 6.21, y: 4.68, w: 0.76, c: 'FFA538', d: 'FF7E38' },
  { x: 6.44, y: 5.48, w: 0.62, c: '00CEE8', d: '00B6CC' },
  { x: 6.94, y: 6.44, w: 0.70, c: 'BD114D', d: '9E0E41' },
  { x: 7.13, y: 4.39, w: 0.80, c: 'FFE400', d: 'FFB700' },
  { x: 7.13, y: 5.80, w: 0.75, c: '0760CF', d: '064BA2' },
  { x: 7.20, y: 5.10, w: 0.76, c: 'F9362A', d: 'C90B22' },
  { x: 7.86, y: 6.30, w: 0.82, c: '00CEE8', d: '00B6CC' },
  { x: 8.04, y: 5.52, w: 0.79, c: 'AC09DC', d: '7A06B3' },
  { x: 8.95, y: 6.44, w: 0.55, c: 'F9362A', d: 'C90B22' },
  { x: 9.17, y: 5.69, w: 0.56, c: 'BD114D', d: '9E0E41' },
  { x: 11.08, y: 1.66, w: 0.68, c: '00CEE8', d: '00B6CC' },
];

function slide14(s) {
  // phone lying flat
  isoBox(s, { x: 6.75, y: 2.34, w: 4.71, ratio: 0.58, depth: 0.18,
    top: '231F20', left: '3E3C3E', right: '6D6E71' });
  isoPlane(s, 6.89, 2.43, 4.41, '77CCFF');
  // horseshoe magnet with white poles
  poly(s, [[8.64, 3.15], [9.22, 2.92], [10.29, 3.35], [9.72, 3.58],
    [9.42, 3.46], [9.02, 3.62]], '0757BD');
  poly(s, [[8.38, 2.92], [8.98, 3.16], [9.14, 3.74], [8.60, 3.98],
    [8.44, 3.40], [8.64, 3.32]], '0761D6');
  poly(s, [[9.48, 3.41], [10.22, 3.10], [10.30, 3.70], [9.60, 4.02]], 'FE2635');
  poly(s, [[9.22, 3.71], [9.79, 3.48], [9.90, 4.10], [9.36, 4.32]], 'DE212F');
  rect(s, 8.19, 2.92, 0.34, 0.24, 'F2F4F8');
  rect(s, 9.66, 3.98, 0.34, 0.24, 'F2F4F8');
  // lightning sparks below the magnet
  poly(s, [[7.56, 3.91], [8.15, 3.99], [7.90, 4.10], [8.33, 4.28],
    [7.72, 4.19], [7.98, 4.06]], WHITE);
  poly(s, [[8.06, 4.20], [8.65, 4.28], [8.40, 4.39], [8.83, 4.57],
    [8.22, 4.48], [8.48, 4.35]], WHITE);
  // reaction pills scattered on the ground
  REACTIONS.forEach((r) => {
    const h = r.w * 0.62;
    ell(s, r.x, r.y + h * 0.30, r.w, h, r.d);
    ell(s, r.x, r.y, r.w, h, r.c);
    ell(s, r.x + r.w * 0.28, r.y + h * 0.26, r.w * 0.44, h * 0.46, WHITE);
  });
  // bean bag with a seated laptop user
  ell(s, 6.29, 1.00, 1.97, 2.16, 'ACCBFD');
  ell(s, 6.29, 1.52, 1.97, 1.64, '90BAFD');
  isoPlane(s, 6.35, 2.30, 1.91, '79ACFD');
  person(s, { x: 6.79, y: 1.09, h: 2.28, shirt: '0761D6', pants: '414042' });
  poly(s, [[6.86, 2.28], [7.55, 2.42], [7.42, 2.60], [6.76, 2.46]], 'EEF1F7');
  // two standing people
  person(s, { x: 11.06, y: 0.87, h: 2.32, shirt: 'F53C33', pants: 'F53C33',
    hair: '8B5E3C' });
  person(s, { x: 10.49, y: 4.00, h: 2.46, shirt: 'FE2635', pants: '0757BD' });
  caption(s, { line1: 'SOCIAL MEDIA', line2: 'MARKETING' });
}

/* ========================================================== slide 15 ==== */
/* TIME MANAGEMENT — a giant wall clock, gears, plant, mug, books. */

const GEARS15 = [
  { x: 10.38, y: 2.21, w: 1.64 },
  { x: 9.87, y: 1.17, w: 0.96 },
  { x: 8.83, y: 1.23, w: 1.15 },
];

function slide15(s) {
  GEARS15.forEach((g) => {
    s.addShape('gear9', { x: g.x, y: g.y + g.w * 0.10, w: g.w, h: g.w * 1.22,
      fill: { color: '0A8BD9' }, line: NOLINE });
    s.addShape('gear9', { x: g.x, y: g.y, w: g.w, h: g.w * 1.22,
      fill: { color: '0AA4FF' }, line: NOLINE });
    ell(s, g.x + g.w * 0.34, g.y + g.w * 0.42, g.w * 0.32, g.w * 0.38, BG);
  });
  // clock: two rim discs offset to fake thickness, then face and markers
  ell(s, 8.01, 1.74, 2.91, 3.91, '31CEE8');
  ell(s, 7.58, 1.79, 2.88, 4.21, '59E6FE');
  ell(s, 7.78, 2.08, 2.48, 3.63, GY75);
  ell(s, 7.92, 2.22, 2.20, 3.35, GY50);
  for (let i = 0; i < 12; i += 1) {
    const a = (i / 12) * Math.PI * 2;
    ell(s, 9.02 + Math.cos(a) * 0.92 - 0.06, 3.90 + Math.sin(a) * 1.42 - 0.07,
      0.12, 0.14, GY90);
  }
  poly(s, [[8.96, 3.94], [9.14, 3.80], [8.98, 2.42], [8.88, 2.44]], 'FE2635');
  poly(s, [[8.96, 3.94], [9.02, 4.06], [9.94, 4.36], [9.98, 4.24]], '065381');
  ell(s, 8.94, 3.86, 0.14, 0.16, '065381');
  plant(s, 6.57, 2.63, 1.44, 2.00);
  mug(s, 10.80, 4.30, 1.13);
  // books, laptop and pencil at the base
  isoBox(s, { x: 9.57, y: 5.37, w: 1.94, ratio: 0.5, depth: 0.30,
    top: GY50, left: GY75, right: GY90 });
  isoBox(s, { x: 9.65, y: 5.13, w: 1.80, ratio: 0.5, depth: 0.14,
    top: 'B8C4D9', left: 'CAD7EE', right: '95A5C0' });
  poly(s, [[10.57, 6.74], [11.60, 6.34], [11.72, 6.46], [10.70, 6.88]], '83A5D8');
  poly(s, [[11.60, 6.34], [11.72, 6.46], [11.86, 6.30]], 'EADDD3');
  person(s, { x: 9.65, y: 5.17, h: 1.91, shirt: 'FE2635', pants: '0AA4FF' });
  person(s, { x: 7.05, y: 4.47, h: 2.12, shirt: '414042', pants: '58595B' });
  caption(s, { line1: 'TIME', line2: 'MANAGEMENT' });
}

/* ========================================================== slide 16 ==== */
/* CONTACTLESS PAYMENT — phone with a PAY NOW button, receipt, card, coins. */

function slide16(s) {
  // phone body: three offset rounded rects give the bevelled edge
  rrect(s, 8.63, 0.65, 2.10, 5.44, '6D6E71', 0.28);
  rrect(s, 8.52, 0.72, 2.02, 5.45, '414042', 0.28);
  rrect(s, 8.56, 0.78, 1.94, 5.34, '000000', 0.26);
  // magenta-to-orange screen, faked with three bands
  rrect(s, 8.64, 0.89, 1.78, 5.11, 'E8483F', 0.22);
  rect(s, 8.64, 2.60, 1.78, 1.60, 'DC395F');
  poly(s, [[8.64, 4.20], [10.42, 4.20], [10.42, 5.80], [8.64, 5.80]], 'D0308C');
  ell(s, 8.64, 5.30, 1.78, 1.00, 'D0308C');
  ell(s, 9.32, 0.95, 0.42, 0.10, '2E2E2E');
  // contactless arcs and the call-to-action button
  [1.00, 0.72, 0.46].forEach((r, i) => {
    s.addShape('arc', {
      x: 9.55 - r / 2, y: 2.06 - r * 0.45 + i * 0.13, w: r, h: r * 0.9,
      angleRange: [200, 340], line: { color: '0A8BD9', width: 5 },
    });
  });
  ell(s, 9.46, 2.36, 0.18, 0.18, '0A8BD9');
  rrect(s, 8.81, 3.44, 1.41, 0.62, '0A8BD9', 0.16);
  txt(s, 'PAY NOW', { x: 8.81, y: 3.60, w: 1.41, h: 0.34, align: 'center',
    fontSize: 12, bold: true });
  // paper receipt curling out from behind the phone
  poly(s, [[8.21, 3.59], [10.09, 3.98], [9.15, 5.69], [7.30, 5.30]], 'D1D3D4');
  poly(s, [[7.14, 4.58], [9.15, 4.99], [8.31, 6.85], [6.32, 6.44]], 'B8BCC4');
  [0, 1, 2, 3].forEach((i) => {
    poly(s, [[8.55 - i * 0.19, 4.16 + i * 0.30], [9.75 - i * 0.19, 4.42 + i * 0.30],
      [9.72 - i * 0.19, 4.52 + i * 0.30], [8.52 - i * 0.19, 4.26 + i * 0.30]], '9AA3B4');
  });
  [0, 1, 2, 3].forEach((i) => {
    poly(s, [[7.50 - i * 0.19, 5.10 + i * 0.30], [8.70 - i * 0.19, 5.36 + i * 0.30],
      [8.67 - i * 0.19, 5.46 + i * 0.30], [7.47 - i * 0.19, 5.20 + i * 0.30]], '9AA3B4');
  });
  // blue paper bag tucked behind the receipt
  poly(s, [[7.50, 3.43], [8.69, 3.69], [8.60, 5.20], [7.60, 5.20]], '0757BD');
  ring(s, 7.84, 3.22, 0.56, 0.44, '074DA8', 4);
  coinStack(s, 10.19, 4.80, 1.02, 6, 0.10);
  // bank card lying flat in the corner
  isoBox(s, { x: 9.88, y: 5.69, w: 2.34, ratio: 0.5, depth: 0.12,
    top: '0A8BD9', left: '0A7CC2', right: '065381' });
  rect(s, 10.30, 6.16, 0.32, 0.22, 'FFD000');
  txt(s, 'CREDIT CARD', { x: 10.72, y: 5.96, w: 1.4, h: 0.3, fontSize: 8,
    color: 'CFEBFF', rotate: 340 });
  caption(s, { line1: 'CONTACTLESS', line2: 'PAYMENT', bodyW: 4.743 });
}

/* ========================================================== slide 17 ==== */
/* HAPPY VALENTINE'S DAY — heart, couple, champagne, teddy, gifts, cupid. */

function slide17(s) {
  s.addShape('heart', { x: 8.39, y: 0.70, w: 3.45, h: 3.35,
    fill: { color: 'DB212E' }, line: NOLINE });
  s.addShape('heart', { x: 8.55, y: 0.60, w: 3.05, h: 3.00,
    fill: { color: 'FE2635' }, line: NOLINE });
  // heart-shaped chocolate box with its lid tipped up
  s.addShape('heart', { x: 7.38, y: 2.30, w: 1.66, h: 1.24,
    fill: { color: 'FFBAC2' }, line: NOLINE });
  s.addShape('heart', { x: 7.49, y: 2.62, w: 1.44, h: 1.10,
    fill: { color: 'FF9A9E' }, line: NOLINE });
  poly(s, [[7.60, 2.34], [8.83, 2.34], [8.75, 2.68], [7.68, 2.68]], 'FE2635');
  // champagne bottle
  poly(s, [[8.19, 3.10], [8.88, 3.10], [8.88, 4.35], [8.19, 4.35]], 'F2F4F8');
  poly(s, [[8.36, 2.09], [8.71, 2.09], [8.78, 3.14], [8.29, 3.14]], '4991A0');
  rect(s, 8.19, 3.30, 0.69, 0.55, '4991A0');
  rect(s, 8.40, 2.09, 0.27, 0.30, '3B7A88');
  // two glasses
  [7.70, 8.79].forEach((gx) => {
    poly(s, [[gx, 2.83], [gx + 0.61, 2.83], [gx + 0.42, 3.52], [gx + 0.19, 3.52]], 'DBE2EE');
    poly(s, [[gx + 0.04, 2.90], [gx + 0.57, 2.90], [gx + 0.40, 3.30],
      [gx + 0.21, 3.30]], 'FFD1B5');
    rect(s, gx + 0.26, 3.52, 0.08, 0.52, 'DBE2EE');
    ell(s, gx + 0.12, 4.02, 0.37, 0.16, 'DBE2EE');
  });
  // the couple
  person(s, { x: 8.97, y: 3.11, h: 2.71, shirt: '77CDFF', pants: '5FA8E8', hair: HAIR });
  person(s, { x: 6.69, y: 4.61, h: 2.41, shirt: '484E59', pants: '3C404A' });
  // bouquet held between them
  ell(s, 8.08, 4.24, 0.53, 0.44, 'FE2635');
  ell(s, 8.20, 4.34, 0.28, 0.24, 'FFA340');
  ell(s, 8.14, 4.46, 0.22, 0.20, '77CDFF');
  // teddy bear
  ell(s, 10.24, 3.66, 0.38, 0.38, 'FCB69F');
  ell(s, 11.02, 3.68, 0.38, 0.38, 'FCB69F');
  ell(s, 10.04, 4.34, 1.56, 1.71, 'FDD4BE');
  ell(s, 10.29, 3.74, 1.06, 1.00, 'FDCAB1');
  ell(s, 10.58, 4.02, 0.46, 0.34, 'FFECD2');
  poly(s, [[10.58, 4.42], [10.98, 4.42], [10.78, 4.76]], 'FE2635');
  // gift box with ribbon and a small ring box
  isoBox(s, { x: 11.11, y: 3.95, w: 1.45, depth: 0.84,
    top: 'FFE985', left: 'FFCF60', right: 'FFB53D' });
  poly(s, [[11.84, 3.95], [12.20, 4.16], [11.84, 4.37], [11.48, 4.16]], '0757BD');
  rect(s, 11.72, 4.32, 0.26, 0.90, '235BB8');
  isoBox(s, { x: 11.27, y: 5.07, w: 0.92, depth: 0.40,
    top: 'FE2635', left: 'DB212E', right: 'C21C27' });
  ring(s, 11.58, 4.92, 0.22, 0.24, 'DBE2EE', 3);
  // cupid perched on the heart
  person(s, { x: 10.19, y: 1.13, h: 1.39, shirt: SKIN, pants: SKIN, hair: 'F7941E' });
  poly(s, [[9.98, 1.24], [10.24, 1.60], [9.94, 1.90]], 'FDFBFB');
  caption(s, { line1: 'HAPPY', line2: 'VALENTINE\u2019S DAY', w2: 4.757, bodyW: 4.461 });
}

/* ========================================================== slide 18 ==== */
/* 3 STEPS INFOGRAPHIC — hub disc, dotted orbit and three numbered pucks. */

const NODES18 = [
  { x: 6.49, y: 4.97, top: 'FE3341', side: 'ED303D', rim: 'DB2D2D', label: '01',
    sx: 6.07, sy: 5.08, sw: 2.01, mx: 7.03, my: 4.63, mc: 'F73131' },
  { x: 9.91, y: 4.36, top: 'FFF200', side: 'FFD000', rim: 'FFB300', label: '02',
    sx: 9.49, sy: 4.58, sw: 2.06, mx: 8.95, my: 4.26, mc: 'FFF200' },
  { x: 11.00, y: 2.35, top: '51F3FE', side: '3ED8F7', rim: '39C6E3', label: '03',
    sx: 10.38, sy: 2.60, sw: 2.21, mx: 9.58, my: 3.15, mc: '89F7FE' },
];

function slide18(s) {
  // dotted orbit ellipse, drawn as individual dots
  for (let i = 0; i < 46; i += 1) {
    const t = (i / 46) * Math.PI * 2;
    ell(s, 7.66 + Math.cos(t) * 2.19 - 0.05, 3.52 + Math.sin(t) * 1.25 - 0.05,
      0.10, 0.10, 'BCBEC0');
  }
  // central hub: extruded disc with a skewed two-line label
  ell(s, 6.06, 2.51, 3.26, 2.15, GY90);
  rect(s, 6.06, 3.08, 3.26, 0.51, GY90);
  ell(s, 6.06, 2.00, 3.26, 2.15, GY75);
  txt(s, 'ISOMETRIC\nINFOGRAPHIC', {
    x: 6.68, y: 2.55, w: 2.05, h: 1.00, align: 'center', valign: 'middle',
    fontFace: F_HEAD, bold: true, fontSize: 19, color: GY90, rotate: 20,
    lineSpacing: 20,
  });
  NODES18.forEach((n) => {
    ell(s, n.sx, n.sy, n.sw, n.sw * 0.63, GY90);
    isoDisc(s, {
      x: n.x, y: n.y, w: 1.21, depth: 0.42, top: n.top, side: n.side,
      label: n.label, labelColor: n.rim, labelSize: 20, labelRotate: 330,
    });
    // ring marker where the node connects to the orbit
    ell(s, n.mx, n.my, 0.63, 0.37, n.mc);
    ell(s, n.mx + 0.13, n.my + 0.08, 0.37, 0.21, BG);
    ell(s, n.mx + 0.22, n.my + 0.13, 0.19, 0.11, n.mc);
  });
  caption(s, { line1: '3 STEPS', line2: 'INFOGRAPHIC', bodyW: 3.94 });
}

/* ------------------------------------------------------------------ main */

const BUILDERS = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
  slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18,
];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'W16x9', width: 13.333, height: 7.5 });
  pptx.layout = 'W16x9';
  pptx.theme = { headFontFace: F_HEAD, bodyFontFace: F_BODY };
  pptx.title = 'Isometric Infographic';

  BUILDERS.forEach((fn) => {
    const slide = pptx.addSlide();
    slide.background = { color: BG };
    fn(slide);
  });
  return pptx;
}

build().writeFile({
  fileName: path.join(__dirname, '0ab87097-f1a2-481e-9d7a-6e99362b5618_grok_final.pptx'),
});
