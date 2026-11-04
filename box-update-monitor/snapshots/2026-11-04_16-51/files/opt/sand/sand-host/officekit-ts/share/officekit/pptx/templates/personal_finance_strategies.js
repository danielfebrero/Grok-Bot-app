/**
 * "Financialy" finance presentation — rebuilt with pptxgenjs.
 *
 * Run: node 023bcb8a-2053-4f3f-9f81-078d3981c1e9_grok_final.js
 * Writes 023bcb8a-2053-4f3f-9f81-078d3981c1e9_grok_final.pptx next to this file.
 *
 * Raster artwork of the original deck is replaced by flat grey "[image]" placeholders.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const FONT = { head: 'Work Sans', body: 'Work Sans Light', icon: 'DejaVu Sans' };

const C = {
  green: '8DEA68', greenD: '55DF1F', greenL: 'D1F7C3', greenXL: 'E8FBE1', greenDk: '399415',
  teal: '18CBD4', tealD: '12989F', tealDk: '0C666A',
  blue: '1295BE', blueD: '0D708F', blueDk: '094A5F',
  navy: '2653AE', navyD: '1D3E82', navyDk: '132A57', navyAlt: '345CB7',
  purple: '6B23D5', purpleL: 'E1D2F8',
  ink: '1D2637', panel: '273248',
  g40: '404040', g59: '595959', g80: '808080',
  white: 'FFFFFF', w65: 'A6A6A6', w95: 'F2F2F2',
  card: 'F6F7FB', grid: 'E7E6E6',
  img: '7F7F7F', imgTxt: 'DDDDDD',
};

/** soft drop shadow, straight down (the deck only uses dir=90deg) */
const shadow = (blur, offset = 0, opacity = 0.15) =>
  ({ type: 'outer', color: '000000', blur: blur, offset: offset, angle: 90, opacity: opacity, rotateWithShape: false });

const tintShadow = (color, blur, offset, opacity) =>
  ({ type: 'outer', color: color, blur: blur, offset: offset, angle: 90, opacity: opacity, rotateWithShape: false });

/* ---------------------------------------------------------------- helpers */

const NOLINE = { type: 'none' };

function rect(s, o) { s.addShape('rect', Object.assign({ line: NOLINE }, o)); }
function ellipse(s, o) { s.addShape('ellipse', Object.assign({ line: NOLINE }, o)); }
function roundRect(s, o) { s.addShape('roundRect', Object.assign({ line: NOLINE }, o)); }

/** PowerPoint's default text-box insets (points): left, right, bottom, top */
const INSET = [7.2, 7.2, 3.6, 3.6];

/** headline / body text shortcuts (everything in the deck is left+top aligned unless told otherwise) */
function tx(s, text, o) {
  s.addText(text, Object.assign({ fontFace: FONT.head, valign: 'top', margin: INSET, wrap: true }, o));
}
function tbody(s, text, o) { tx(s, text, Object.assign({ fontFace: FONT.body }, o)); }

/** polygon / bezier outline; pts are [x,y] pairs in inches relative to (x,y) */
function poly(s, x, y, pts, o) {
  const points = pts.map((p, i) => (i === 0 ? { x: p[0], y: p[1], moveTo: true } : { x: p[0], y: p[1] }));
  points.push({ close: true });
  s.addShape('custGeom', Object.assign({ x: x, y: y, w: 0, h: 0, line: NOLINE, points: points }, o));
}

/** polygon described with fractions of a w x h box */
function polyBox(s, box, fracs, o) {
  poly(s, box.x, box.y, fracs.map(f => [f[0] * box.w, f[1] * box.h]), Object.assign({ w: box.w, h: box.h }, o));
}

/** circular / elliptical pie wedge (used for the donut charts and the 80% gauges) */
function wedge(s, cx, cy, rx, ry, startDeg, sweepDeg, o) {
  const rad = d => (d * Math.PI) / 180;
  const x0 = rx + rx * Math.cos(rad(startDeg));
  const y0 = ry + ry * Math.sin(rad(startDeg));
  s.addShape('custGeom', Object.assign({
    x: cx - rx, y: cy - ry, w: 2 * rx, h: 2 * ry, line: NOLINE,
    points: [
      { x: rx, y: ry, moveTo: true },
      { x: x0, y: y0 },
      { x: 0, y: 0, curve: { type: 'arc', hR: ry, wR: rx, stAng: startDeg, swAng: sweepDeg } },
      { close: true },
    ],
  }, o));
}

/** hex string blend, t=0 -> a, t=1 -> b */
function mix(a, b, t) {
  const ch = i => Math.round(parseInt(a.substr(i, 2), 16) * (1 - t) + parseInt(b.substr(i, 2), 16) * t);
  return [ch(0), ch(2), ch(4)].map(v => v.toString(16).padStart(2, '0').toUpperCase()).join('');
}

/**
 * pptxgenjs has no gradient fill, so left-to-right gradients are painted as
 * vertical slices. `edges(u)` gives [topFrac, bottomFrac] at horizontal fraction u.
 */
function gradientBand(s, box, from, to, edges, steps) {
  steps = steps || 20;
  const over = 0.6 / steps; // slight overlap so slices don't show hairlines
  for (let i = 0; i < steps; i++) {
    const u0 = i / steps, u1 = Math.min(1, (i + 1) / steps + over);
    const e0 = edges(u0), e1 = edges(u1);
    polyBox(s, box, [[u0, e0[0]], [u1, e1[0]], [u1, e1[1]], [u0, e0[1]]],
      { fill: { color: mix(from, to, (u0 + u1) / 2) } });
  }
}

/** grey stand-in for a photo */
function imageBox(s, x, y, w, h) {
  rect(s, { x: x, y: y, w: w, h: h, fill: { color: C.img } });
  imageLabel(s, x, y, w, h);
}
function imageLabel(s, x, y, w, h) {
  tx(s, '[image]', {
    x: x, y: y + h / 2 - 0.16, w: w, h: 0.32, align: 'center', valign: 'middle',
    fontSize: 11, color: C.imgTxt,
  });
}

/* ------------------------------------------------------- recurring motifs */

/**
 * "@Financialy" strip inherited from the slide master (top-left of every slide).
 * `tone` is the disc colour: green on light slides, white on the green master.
 */
function socialHeader(s, tone) {
  const glyphColor = tone === C.white ? C.green : C.white;
  ['f', '\u25B6', '\u25CE', '\u2726'].forEach((m, i) => {
    const x = 0.509 + i * 0.2475;
    ellipse(s, { x: x, y: 0.471, w: 0.164, h: 0.164, fill: { color: tone } });
    s.addText(m, {
      x: x, y: 0.471, w: 0.164, h: 0.164, align: 'center', valign: 'middle', margin: 0,
      fontSize: 7, color: glyphColor, fontFace: FONT.icon,
    });
  });
  tx(s, '@Financialy', { x: 1.428, y: 0.395, w: 1.488, h: 0.286, fontSize: 11, color: tone === C.white ? C.white : C.w65 });
}

/** translucent disc + glyph — the little round icon badges used all over the deck */
function iconBadge(s, x, y, size, color, glyph, opt) {
  opt = opt || {};
  ellipse(s, {
    x: x, y: y, w: size, h: size,
    fill: { color: color, transparency: 80 },
    line: { color: color, transparency: 75, width: 1 },
  });
  const inner = size * 0.615;
  if (opt.filled) ellipse(s, { x: x + (size - inner) / 2, y: y + (size - inner) / 2, w: inner, h: inner, fill: { color: color } });
  s.addText(glyph, {
    x: x, y: y, w: size, h: size, align: 'center', valign: 'middle', margin: 0, fontFace: FONT.icon,
    fontSize: opt.fontSize || Math.round(size * 26), color: opt.glyphColor || (opt.filled ? C.white : color),
  });
}

/** round previous / next button (arrow inside a translucent disc) */
function navButton(s, x, y, size, color, dir) {
  ellipse(s, {
    x: x, y: y, w: size, h: size,
    fill: { color: color, transparency: 80 },
    line: { color: color, transparency: 75, width: 1 },
  });
  const aw = size * 0.52, ah = size * 0.40;
  s.addShape(dir === 'left' ? 'leftArrow' : 'rightArrow', {
    x: x + (size - aw) / 2, y: y + (size - ah) / 2, w: aw, h: ah,
    fill: { color: color }, line: NOLINE,
  });
}

/** white disc with a coloured chevron (slide 1 hero button, slide 7 month arrows) */
function chevronButton(s, x, y, size, ringColor, markColor, dir) {
  ellipse(s, {
    x: x, y: y, w: size, h: size,
    fill: { color: ringColor, transparency: 80 },
    line: { color: ringColor, transparency: 75, width: 1 },
  });
  const d = size * 0.62;
  ellipse(s, { x: x + (size - d) / 2, y: y + (size - d) / 2, w: d, h: d, fill: { color: ringColor } });
  const s2 = size * 0.30, cx = x + size / 2 - s2 / 2, cy = y + size / 2 - s2 / 2;
  const f = dir === 'left'
    ? [[0.75, 0.05], [0.95, 0.22], [0.62, 0.5], [0.95, 0.78], [0.75, 0.95], [0.25, 0.5]]
    : [[0.25, 0.05], [0.75, 0.5], [0.25, 0.95], [0.05, 0.78], [0.38, 0.5], [0.05, 0.22]];
  polyBox(s, { x: cx, y: cy, w: s2, h: s2 }, f, { fill: { color: markColor } });
}

/** the "80%" ring + caption block that shows up on slides 4 and 9 */
function gaugeBlock(s, x, y, textColor) {
  ellipse(s, { x: x, y: y, w: 0.944, h: 0.944, fill: { color: C.purpleL, transparency: 70 } });
  wedge(s, x + 0.472, y + 0.472, 0.472, 0.472, 210, 262, { fill: { color: C.greenL } });
  ellipse(s, { x: x + 0.085, y: y + 0.085, w: 0.774, h: 0.774, fill: { color: C.white } });
  s.addText([
    { text: '80', options: { fontSize: 18, bold: true } },
    { text: '%', options: { fontSize: 9, bold: true } },
  ], { x: x + 0.055, y: y + 0.24, w: 0.855, h: 0.404, align: 'center', valign: 'middle', margin: 0, fontFace: FONT.head, color: C.g40 });
  tx(s, 'Safe and Secure', { x: x + 1.21, y: y + 0.004, w: 2.666, h: 0.34, fontSize: 12, bold: true, color: textColor, lineSpacingMultiple: 1.3 });
  tbody(s, 'A wonderful serenity has taken possession of my entire soul.',
    { x: x + 1.21, y: y + 0.338, w: 2.854, h: 0.602, fontSize: 12, color: textColor, lineSpacingMultiple: 1.3 });
}

/** white card + heading + sub-line, used on slides 5 and 10 */
function contentCard(s, x, y, title, sub, subColor, accent, glyph) {
  rect(s, { x: x, y: y, w: 4.612, h: 1.141, fill: { color: C.white }, shadow: shadow(66, 17, 0.15) });
  tx(s, title, { x: x + 0.246, y: y + 0.231, w: 3.279, h: 0.37, fontSize: 16, color: C.ink });
  tbody(s, sub, { x: x + 0.246, y: y + 0.606, w: 3.279, h: 0.303, fontSize: 12, color: subColor });
  iconBadge(s, x + 3.722, y + 0.306, 0.528, accent, glyph, { filled: true, fontSize: 12 });
}

/** faint dotted background pattern (slides 12 and 14) */
function dotGrid(s, x, y, w, h, cols, rows, dot, color, transparency) {
  const sx = w / cols, sy = h / rows;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      ellipse(s, {
        x: x + c * sx, y: y + r * sy, w: dot, h: dot,
        fill: { color: color, transparency: transparency },
      });
    }
  }
}

/** polyline drawn as a chain of line shapes with round dot markers */
function polyline(s, pts, color, width, marker) {
  for (let i = 0; i < pts.length - 1; i++) {
    const a = pts[i], b = pts[i + 1];
    s.addShape('line', {
      x: Math.min(a[0], b[0]), y: Math.min(a[1], b[1]),
      w: Math.abs(b[0] - a[0]), h: Math.abs(b[1] - a[1]),
      flipH: b[0] < a[0], flipV: b[1] < a[1],
      line: { color: color, width: width },
    });
  }
  if (marker) {
    pts.forEach(p => ellipse(s, {
      x: p[0] - marker / 2, y: p[1] - marker / 2, w: marker, h: marker,
      fill: { color: C.white }, line: { color: color, width: 2.25 },
    }));
  }
}

/* ------------------------------------------------------------ slide 1..15 */

function slide01(pptx) {
  const s = pptx.addSlide();
  socialHeader(s, C.green);
  imageBox(s, 1.116, 4.592, 5.125, 2.908);
  imageBox(s, 8.208, 0, 5.125, 7.5);
  gradientBand(s, { x: 5.012, y: 3.536, w: 4.515, h: 1.758 }, C.green, C.greenD, () => [0, 1], 24);
  tx(s, 'Finance Presentation ', { x: 5.524, y: 3.961, w: 2.903, h: 1.043, fontSize: 28, color: C.white });
  chevronButton(s, 8.382, 4.293, 0.53, C.white, C.green, 'right');
  tx(s, 'Financialy', { x: 1.309, y: 1.795, w: 6.101, h: 1.447, fontSize: 80, color: C.green });
}

function slide02(pptx) {
  const s = pptx.addSlide();
  socialHeader(s, C.green);
  // corner photo, clipped to the lower-right triangle of its frame
  polyBox(s, { x: 5.833, y: 0, w: 7.5, h: 7.5 }, [[1, 0], [1, 1], [0, 1]], { fill: { color: C.img } });
  imageLabel(s, 10.2, 4.6, 2.2, 0.4);

  tx(s, 'Personal Finance Strategies', { x: 1.62, y: 1.83, w: 5.102, h: 2.524, fontSize: 48, color: C.ink });
  tbody(s, "Winged waters beginning saw which. First signs Can't may great seed image. Two, sixth, darkness. ",
    { x: 1.62, y: 4.513, w: 3.957, h: 0.992, fontSize: 14, color: C.g59, lineSpacingMultiple: 1.3 });

  const cards = [
    { y: 1.128, label: 'Creativity', color: C.green, glyph: '\u21A9', ty: 1.716, badge: 0.817, body: 2.12 },
    { y: 4.149, label: 'Skills', color: C.blue, glyph: null, ty: 4.791, badge: 3.82, body: 5.205 },
  ];
  cards.forEach(c => {
    rect(s, { x: 6.47, y: c.y, w: 4.008, h: 2.408, fill: { color: C.white }, shadow: shadow(100, 0, 0.12) });
    tx(s, c.label, { x: 7.109, y: c.ty, w: 2.73, h: 0.406, fontSize: 16, bold: true, align: 'center', color: c.color });
    tbody(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy.',
      { x: 6.921, y: c.body, w: 3.106, h: 0.812, fontSize: 12, align: 'center', color: C.g40, lineSpacingMultiple: 1.2 });
    rect(s, { x: 8.152, y: c.badge, w: 0.643, h: 0.646, fill: { color: c.color }, shadow: shadow(25, 10, 0.15) });
    if (c.glyph) {
      tx(s, c.glyph, { x: 8.152, y: c.badge, w: 0.643, h: 0.646, align: 'center', valign: 'middle', fontSize: 20, color: C.white, fontFace: FONT.icon });
    } else {
      [[0.20, 0.30, 0.16], [0.42, 0.42, 0.24], [0.64, 0.22, 0.34]].forEach(b =>
        rect(s, { x: 8.152 + b[0] * 0.643, y: c.badge + 0.646 - 0.14 - b[2] * 0.646, w: 0.643 * 0.17, h: b[2] * 0.646, fill: { color: C.white } }));
    }
  });
}

function slide03(pptx) {
  const s = pptx.addSlide();
  socialHeader(s, C.green);
  rect(s, { x: 0, y: 0, w: 13.333, h: 4.278, fill: { color: C.green } });
  tx(s, 'Discussion and analyze', { x: 1.35, y: 0.617, w: 10.633, h: 0.909, fontSize: 48, align: 'center', color: C.white });

  [{ x: 1.213, pct: '60%' }, { x: 4.838, pct: '45%' }, { x: 8.463, pct: '50%' }].forEach(col => {
    rect(s, { x: col.x, y: 2.069, w: 3.331, h: 4.551, fill: { color: C.white }, shadow: shadow(100, 0, 0.12) });
    imageBox(s, col.x + 0.064, 2.127, 3.203, 1.936);
    tx(s, col.pct, { x: col.x + 0.49, y: 4.521, w: 2.351, h: 0.64, fontSize: 32, bold: true, align: 'center', color: C.ink });
    tx(s, 'Your text goes here', { x: col.x + 0.267, y: 5.161, w: 2.796, h: 0.303, fontSize: 12, align: 'center', color: C.green });
    tbody(s, 'A wonderful serenity has taken possession of my entire soul',
      { x: col.x + 0.267, y: 5.534, w: 2.796, h: 0.57, fontSize: 12, align: 'center', color: C.g40, lineSpacingMultiple: 1.2 });
  });
}

/** flat mock-up of the laptop artwork (white shell + grey base) */
function laptopShell(s, x, y, w, h) {
  const bodyH = h * 0.85;
  roundRect(s, { x: x + w * 0.07, y: y, w: w * 0.86, h: bodyH, rectRadius: 0.06, fill: { color: C.white }, line: { color: C.grid, width: 1 } });
  rect(s, { x: x, y: y + bodyH, w: w, h: h * 0.06, fill: { color: C.w95 } });
  polyBox(s, { x: x, y: y + bodyH + h * 0.06, w: w, h: h * 0.09 },
    [[0.02, 0], [0.98, 0], [0.93, 1], [0.07, 1]], { fill: { color: C.grid } });
}

function slide04(pptx) {
  const s = pptx.addSlide();
  laptopShell(s, 5.917, 3.056, 6.578, 3.399);
  imageBox(s, 8.163, 0.942, 3.163, 1.851);
  imageBox(s, 6.842, 3.269, 4.727, 2.779);
  rect(s, { x: 0, y: 0, w: 6.5, h: 7.5, fill: { color: C.green } });
  tx(s, 'Laptop Mockup', { x: 1.113, y: 1.366, w: 4.234, h: 1.717, fontSize: 48, color: C.white });
  tbody(s, "Winged waters beginning saw which. First signs Can't may great seed image. Two, sixth, darkness. ",
    { x: 1.113, y: 3.388, w: 3.457, h: 0.992, fontSize: 14, color: C.white, lineSpacingMultiple: 1.3 });
  gaugeBlock(s, 1.113, 4.832, C.white);
  imageBox(s, 4.981, 1.972, 4.433, 2.594);
}

function slide05(pptx) {
  const s = pptx.addSlide();
  socialHeader(s, C.green);
  rect(s, { x: 7.583, y: 0, w: 5.75, h: 7.5, fill: { color: C.green } });
  laptopShell(s, 7.441, 1.702, 9.098, 4.701);
  tx(s, 'The Mockup Slide', { x: 1.155, y: 1.533, w: 4.871, h: 1.717, fontSize: 48, color: C.ink });

  contentCard(s, 1.31, 3.554, 'Amazing content One', 'Winged waters beginning.', C.green, C.green, '\u2699');
  contentCard(s, 1.31, 4.859, 'Amazing content Two', 'Winged waters beginning.', C.teal, C.teal, '\u2708');

  // tablet
  roundRect(s, { x: 6.835, y: 2.639, w: 2.621, h: 4.049, rectRadius: 0.12, fill: { color: C.white }, line: { color: C.grid, width: 1.5 } });
  [3.272, 3.698, 4.088].forEach(y => roundRect(s, { x: 6.813, y: y, w: 0.135, h: 0.211, rectRadius: 0.5, fill: { color: C.g40 } }));
  roundRect(s, { x: 9.346, y: 3.337, w: 0.135, h: 0.372, rectRadius: 0.5, fill: { color: C.g40 } });
  roundRect(s, { x: 7.911, y: 2.88, w: 0.469, h: 0.035, rectRadius: 0.5, fill: { color: '000000', transparency: 80 } });
  ellipse(s, { x: 7.752, y: 2.875, w: 0.044, h: 0.044, fill: { color: '000000', transparency: 80 } });
  ellipse(s, { x: 8.123, y: 2.737, w: 0.044, h: 0.044, fill: { color: '000000', transparency: 80 } });
  ellipse(s, { x: 7.992, y: 6.271, w: 0.306, h: 0.31, fill: { type: 'none' }, line: { color: C.grid, width: 1.25 } });
  imageBox(s, 6.947, 3.022, 2.397, 3.142);
  imageBox(s, 8.714, 2.0, 4.619, 3.865);
}

function slide06(pptx) {
  const s = pptx.addSlide();
  s.background = { color: C.green };
  socialHeader(s, C.white);
  rect(s, { x: 0.557, y: 0.902, w: 12.219, h: 5.697, fill: { color: C.white }, shadow: shadow(100, 0, 0.12) });

  const cols = [1.64, 3.748, 5.855, 7.962, 10.069];
  const rows = [
    {
      y: 1.744, title: 'Schedule Number One', month: 'January 2024', highlight: 3,
      cells: [['Departure', '08 AM'], ['Departure', '10 AM'], ['Arrival', '04 PM'], ['Arrival', '04 PM'], ['Departure', '02 PM']],
    },
    {
      y: 4.57, title: 'Schedule Number Two', month: 'February 2024', highlight: -1,
      cells: [['Departure', '08 AM'], ['Departure', '10 AM'], ['Arrival', '04 PM'], ['Arrival', '04 PM'], ['Departure', '02 PM']],
    },
  ];

  rows.forEach(row => {
    const hy = row.y - 0.62;
    roundRect(s, { x: 1.004, y: hy - 0.172 + 0.344, w: 0.394, h: 0.05, rectRadius: 0.5, fill: { color: C.green }, shadow: tintShadow(C.green, 8, 3, 0.25) });
    tx(s, row.title, { x: 1.64, y: hy, w: 2.992, h: 0.362, fontSize: 14, color: C.ink });
    tx(s, row.month, { x: 9.518, y: hy, w: 2.175, h: 0.362, fontSize: 14, align: 'right', color: C.ink });

    row.cells.forEach((cell, i) => {
      const hot = i === row.highlight;
      const fg = hot ? C.white : C.g59;
      rect(s, {
        x: cols[i], y: row.y, w: 1.624, h: 1.722,
        fill: hot ? { color: C.teal } : { color: C.white, transparency: 100 },
        line: hot ? NOLINE : { color: C.w65, width: 1, dashType: 'dash' },
        shadow: hot ? shadow(54, 12, 0.25) : undefined,
      });
      tbody(s, cell[0], { x: cols[i] + 0.148, y: row.y + 0.154, w: 1.329, h: 0.306, fontSize: 11, color: fg, lineSpacingMultiple: 1.2 });
      tx(s, cell[1], { x: cols[i] + 0.148, y: row.y + 0.35, w: 1.329, h: 0.548, fontSize: 24, color: hot ? C.white : C.green, lineSpacingMultiple: 1.2 });
      tbody(s, 'Thu 22 Feb', { x: cols[i] + 0.148, y: row.y + 1.263, w: 1.329, h: 0.306, fontSize: 11, color: fg, lineSpacingMultiple: 1.2 });
    });
  });

  s.addShape('line', { x: 1.633, y: 3.75, w: 10.068, h: 0, line: { color: C.w65, transparency: 70, width: 1 } });
  [2.188, 5.013].forEach(y => {
    navButton(s, 12.042, y, 0.363, C.green, 'left');
    navButton(s, 12.042, y + 0.472, 0.363, C.green, 'right');
  });
}

function slide07(pptx) {
  const s = pptx.addSlide();
  rect(s, { x: 0, y: 0, w: 5.75, h: 7.5, fill: { color: C.green } });
  rect(s, { x: 5.347, y: 0.944, w: 6.832, h: 5.612, fill: { color: C.white }, shadow: shadow(100, 0, 0.12) });
  rect(s, { x: 5.347, y: 0.948, w: 6.832, h: 0.684, fill: { color: C.teal } });
  tx(s, 'August 2024', { x: 6.867, y: 1.033, w: 3.79, h: 0.505, fontSize: 24, align: 'center', color: C.white });
  chevronButton(s, 5.759, 1.077, 0.417, C.white, C.teal, 'left');
  chevronButton(s, 11.35, 1.077, 0.417, C.white, C.teal, 'right');

  const colX = [5.658, 6.594, 7.53, 8.466, 9.402, 10.338, 11.269];
  ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].forEach((d, i) => {
    roundRect(s, { x: 5.537 + i * 0.936, y: 1.852, w: 0.841, h: 0.358, rectRadius: 0.05, fill: { color: C.white } });
    tx(s, d, { x: 5.537 + i * 0.936, y: 1.852, w: 0.841, h: 0.358, align: 'center', valign: 'middle', fontSize: 12, color: C.green });
  });

  const rowY = [2.311, 3.012, 3.713, 4.43, 5.094, 5.761];
  const weeks = [
    ['29', '30', '31', '1', '2', '3', '4'],
    ['5', '6', '7', '8', '9', '10', '11'],
    ['12', '13', '14', '15', '16', '17', '18'],
    ['19', '20', '21', '22', '23', '24', '25'],
    ['26', '27', '28', '29', '30', '31', '1'],
    ['2', '3', '4', '5', '6', '7', '8'],
  ];
  const muted = (r, c) => (r === 0 && c < 3) || (r === 4 && c === 6) || r === 5;

  // highlight pill behind 14 - 16
  roundRect(s, {
    x: 7.461, y: 3.662, w: 2.604, h: 0.701, rectRadius: 0.5,
    fill: { color: C.w65, transparency: 85 }, line: { color: C.w65, transparency: 75, width: 1 },
  });

  weeks.forEach((week, r) => week.forEach((day, c) => {
    const picked = r === 2 && (c === 2 || c === 4);
    if (picked) {
      ellipse(s, { x: colX[c], y: rowY[r], w: 0.6, h: 0.6, fill: { color: C.teal }, shadow: tintShadow(C.green, 32, 3, 0.4) });
    }
    tbody(s, day, {
      x: colX[c], y: rowY[r], w: 0.6, h: 0.6, align: 'center', valign: 'middle', fontSize: 12,
      color: picked ? C.white : (muted(r, c) ? C.w65 : C.g59),
    });
  }));

  tx(s, '2024\nCalendar', { x: 1.155, y: 1.778, w: 3.781, h: 1.717, fontSize: 48, color: C.white });

  rect(s, { x: 1.192, y: 3.994, w: 3.468, h: 1.717, fill: { color: C.card }, line: { color: C.w65, transparency: 75, width: 1 } });
  s.addShape('line', { x: 1.643, y: 4.862, w: 3.017, h: 0, line: { color: C.w65, transparency: 75, width: 1 } });
  s.addShape('line', { x: 1.429, y: 4.596, w: 0, h: 0.612, line: { color: C.green, width: 1.5, dashType: 'sysDot' } });
  [['Tue, 14 Aug', 4.139], ['Tue, 16 Aug', 5.008]].forEach(([label, y]) => {
    tx(s, label, { x: 1.603, y: y, w: 2.789, h: 0.37, fontSize: 16, color: C.g59 });
    tx(s, 'San Francisco Intl. Departement', { x: 1.603, y: y + 0.342, w: 2.789, h: 0.286, fontSize: 11, color: C.green });
    ellipse(s, { x: 1.374, y: y + 0.287, w: 0.11, h: 0.11, fill: { color: C.card }, line: { color: C.green, width: 1.5 } });
  });
}

function slide08(pptx) {
  const s = pptx.addSlide();
  socialHeader(s, C.green);
  rect(s, { x: 5.524, y: 0, w: 7.254, h: 7.5, fill: { color: C.green } });

  // nested hexagons (radar background)
  const hexPts = [[0.5, 0], [1, 0.25], [1, 0.75], [0.5, 1], [0, 0.75], [0, 0.25]];
  [[7.506, 2.042, 3.062, 3.427, 46], [7.85, 2.425, 2.376, 2.659, 33],
   [8.19, 2.807, 1.695, 1.896, 33], [8.527, 3.183, 1.022, 1.143, 33]].forEach(h => {
    polyBox(s, { x: h[0], y: h[1], w: h[2], h: h[3] }, hexPts,
      { fill: { color: C.white }, shadow: shadow(h[4], 0, 0.15) });
  });

  // radar plot
  const radar = [[7.879, 3.638], [8.774, 2.532], [10.111, 3.06], [9.87, 4.713], [8.539, 4.424], [7.879, 3.638]];
  polyline(s, radar, C.green, 2, 0);
  radar.slice(0, 5).forEach(p => rect(s, { x: p[0] - 0.031, y: p[1] - 0.031, w: 0.062, h: 0.062, fill: { color: C.teal }, line: { color: C.green, width: 1 } }));

  s.addShape('wedgeRectCallout', { x: 9.911, y: 2.649, w: 0.543, h: 0.31, fill: { color: C.g59 }, line: NOLINE });
  tx(s, '62%', { x: 9.911, y: 2.649, w: 0.543, h: 0.31, align: 'center', valign: 'middle', fontSize: 11, color: C.white, margin: 0 });

  [[7.822, 1.62, 2.431, 'center'], [7.822, 5.577, 2.431, 'center'],
   [10.641, 2.743, 1.538, 'left'], [10.641, 4.535, 1.538, 'left'],
   [5.897, 2.743, 1.538, 'right'], [5.897, 4.535, 1.538, 'right']].forEach(p => {
    tbody(s, 'Value content', { x: p[0], y: p[1], w: p[2], h: 0.303, fontSize: 12, align: p[3], color: C.white });
  });

  tx(s, 'The Custom Chart', { x: 1.155, y: 1.61, w: 4.727, h: 1.717, fontSize: 48, color: C.ink });
  rect(s, { x: 1.263, y: 3.826, w: 4.405, h: 2.064, fill: { color: C.white }, shadow: shadow(68, 17, 0.15) });

  const bars = [0.539, 0.788, 0.685, 1.097, 1.338, 0.943, 0.501];
  bars.forEach((bh, i) => roundRect(s, {
    x: 3.692 + i * 0.2637, y: 5.687 - bh, w: 0.124, h: bh, rectRadius: 0.5,
    fill: { color: C.green, transparency: i === 4 ? 0 : 75 },
  }));
  tbody(s, 'Week 5', { x: 4.284, y: 4.028, w: 1.057, h: 0.286, fontSize: 11, align: 'center', color: C.g59 });
  tx(s, '$ 3,220', { x: 1.507, y: 4.828, w: 1.916, h: 0.572, fontSize: 28, color: C.ink });
  tbody(s, 'Total flights', { x: 1.507, y: 5.375, w: 1.916, h: 0.286, fontSize: 11, color: C.g59 });
  iconBadge(s, 1.605, 4.137, 0.526, C.green, '\u2261', { filled: true, fontSize: 12 });

  navButton(s, 11.173, 5.721, 0.424, C.white, 'left');
  navButton(s, 11.755, 5.721, 0.424, C.white, 'right');
}

function slide09(pptx) {
  const s = pptx.addSlide();

  // dashed measuring grid
  const vLines = [[8.848, 1.699, 5.7], [9.515, 1.699, 5.7], [10.182, 1.699, 5.7],
                  [8.182, 2.366, 5.7], [10.849, 2.366, 5.7], [7.515, 3.7, 5.7], [11.516, 3.7, 5.7]];
  vLines.forEach(l => s.addShape('line', { x: l[0], y: l[1], w: 0, h: l[2] - l[1], line: { color: C.w65, transparency: 60, width: 1, dashType: 'dash' } }));
  const hLines = [[7.515, 11.516, 5.033], [7.515, 11.516, 4.366], [7.515, 11.516, 3.7],
                  [8.182, 10.849, 3.033], [8.182, 10.849, 2.366], [8.848, 10.182, 1.699]];
  hLines.forEach(l => s.addShape('line', { x: l[0], y: l[2], w: l[1] - l[0], h: 0, line: { color: C.w65, transparency: 60, width: 1, dashType: 'dash' } }));
  s.addShape('line', { x: 6.307, y: 5.7, w: 6.416, h: 0, line: { color: C.w65, transparency: 60, width: 1 } });

  polyline(s, [[6.357, 5.68], [7.717, 4.759], [8.873, 1.892], [10.158, 1.889], [11.431, 4.806], [12.68, 5.691]], C.green, 3, 0.09);
  polyline(s, [[6.367, 5.66], [7.83, 5.132], [9.085, 3.179], [9.504, 3.064]], C.teal, 3, 0.09);
  polyline(s, [[9.522, 3.731], [9.966, 3.794], [11.442, 5.446], [12.669, 5.693]], C.blue, 3, 0.09);
  polyline(s, [[6.367, 5.68], [7.945, 5.309], [9.148, 4.415], [9.535, 4.38]], C.navy, 3, 0.09);

  rect(s, { x: 0, y: 0, w: 5.75, h: 7.5, fill: { color: C.green } });

  s.addShape('wedgeRectCallout', { x: 9.776, y: 2.365, w: 2.303, h: 1.141, fill: { color: C.white }, line: NOLINE, shadow: shadow(66, 17, 0.15) });
  tx(s, '$ 3,220', { x: 9.974, y: 2.527, w: 1.916, h: 0.572, fontSize: 28, color: C.ink });
  tbody(s, 'Total flights', { x: 9.974, y: 3.059, w: 1.916, h: 0.286, fontSize: 11, color: C.g59 });

  [7.367, 9.05, 10.733].forEach(x => tbody(s, 'Data content', {
    x: x, y: 5.94, w: 1.33, h: 0.286, fontSize: 11, color: C.g59, bullet: { characterCode: '2022', indent: 14 },
  }));

  tx(s, 'The Chart Custom', { x: 1.113, y: 1.366, w: 4.727, h: 1.717, fontSize: 48, color: C.white });
  tbody(s, "Winged waters beginning saw which. First signs Can't may great seed image. Two, sixth, darkness. ",
    { x: 1.113, y: 3.388, w: 3.457, h: 0.992, fontSize: 14, color: C.white, lineSpacingMultiple: 1.3 });
  gaugeBlock(s, 1.113, 4.832, C.white);
}

function slide10(pptx) {
  const s = pptx.addSlide();
  socialHeader(s, C.green);

  // column chart: baseline 6.333, 45 units == 4.847 in
  const base = 6.333, unit = 4.847 / 45;
  for (let v = 45; v >= 0; v -= 5) {
    const y = base - v * unit;
    s.addShape('line', { x: 1.336, y: y, w: 5.39, h: 0, line: { color: C.grid, width: 0.75, dashType: 'lgDash' } });
    tbody(s, String(v), { x: 0.788, y: y - 0.21, w: 0.676, h: 0.421, fontSize: 16, align: 'center', color: C.g80 });
  }
  const series = [
    [1.505, 9.9, C.teal], [1.83, 19.8, C.green], [2.165, 36.8, C.teal], [2.495, 22.2, C.green],
    [2.835, 33.0, C.teal], [3.168, 34.9, C.green], [3.503, 29.8, C.teal], [3.834, 27.5, C.green],
    [4.164, 23.0, C.teal], [4.497, 41.5, C.green], [4.83, 19.8, C.teal], [5.157, 27.0, C.green],
    [5.501, 24.8, C.teal], [5.828, 27.9, C.green], [6.162, 39.7, C.teal], [6.495, 17.9, C.green],
  ];
  series.forEach(b => rect(s, { x: b[0], y: base - b[1] * unit, w: 0.155, h: b[1] * unit, fill: { color: b[2] } }));

  rect(s, { x: 7.181, y: 0, w: 5.365, h: 7.5, fill: { color: C.green } });
  s.addText([
    { text: 'The ' }, { text: 'Financialy' }, { text: ' Costume Chart' },
  ], { x: 7.5, y: 1.735, w: 4.516, h: 1.447, fontSize: 40, color: C.white, fontFace: FONT.head, valign: 'top', margin: 0 });

  contentCard(s, 7.567, 3.456, 'Amazing content One', 'Winged waters beginning.', C.green, C.green, '\u2699');
  contentCard(s, 7.567, 4.762, 'Amazing content Two', 'Winged waters beginning.', C.teal, C.teal, '\u2708');

  rect(s, { x: 5.011, y: 4.177, w: 1.881, h: 1.881, fill: { color: C.panel }, shadow: shadow(66, 17, 0.15) });
  tx(s, '35K', { x: 5.091, y: 4.615, w: 1.722, h: 0.774, fontSize: 40, align: 'center', color: C.white });
  tx(s, 'Total data', { x: 5.091, y: 5.283, w: 1.722, h: 0.337, fontSize: 14, align: 'center', color: C.white });
}

function slide11(pptx) {
  const s = pptx.addSlide();
  socialHeader(s, C.green);
  rect(s, { x: 0, y: 0, w: 13.333, h: 3.333, fill: { color: C.green } });
  s.addText([
    { text: 'The ' }, { text: 'Financialy' }, { text: ' Chart' },
  ], { x: 1.35, y: 0.617, w: 10.633, h: 0.909, fontSize: 48, align: 'center', color: C.white, fontFace: FONT.head, valign: 'top', margin: 0 });

  const donutSlices = [
    { start: 342.24, sweep: 107.33, color: C.teal },
    { start: 269.37, sweep: 73.0, color: C.green },
    { start: 89.69, sweep: 144.44, color: C.blue },
    { start: 234.11, sweep: 35.18, color: C.navyAlt },
  ];
  const legend = [['Data 1', C.green], ['Data 2 ', C.teal], ['Data 3', C.blue], ['Data 4', C.navy]];

  [1.445, 9.179].forEach(cardX => {
    rect(s, { x: cardX, y: 2.163, w: 3.502, h: 4.127, fill: { color: C.white }, shadow: shadow(88, 21, 0.15) });
    tx(s, 'Amazing content', { x: cardX + 0.289, y: 2.374, w: 2.925, h: 0.37, fontSize: 16, align: 'center', color: C.ink });
    tbody(s, '2023-2024', { x: cardX + 0.289, y: 2.719, w: 2.925, h: 0.286, fontSize: 11, align: 'center', color: C.green });
    donutSlices.forEach(sl => wedge(s, cardX + 1.762, 4.448, 1.194, 1.213, sl.start, sl.sweep, { fill: { color: sl.color } }));
    ellipse(s, { x: cardX + 0.923, y: 3.608, w: 1.679, h: 1.679, fill: { color: C.panel } });
    tx(s, '78%', { x: cardX + 0.923, y: 3.608, w: 1.679, h: 1.679, align: 'center', valign: 'middle', fontSize: 18, color: C.white, fontFace: 'Montserrat' });
    legend.forEach((lg, i) => {
      const lx = cardX + 0.314 + i * 0.7935;
      rect(s, { x: lx, y: 5.878, w: 0.11, h: 0.1, fill: { color: lg[1] } });
      tbody(s, lg[0], { x: lx + 0.09, y: 5.781, w: 0.758, h: 0.294, fontSize: 10, color: C.g40, lineSpacingMultiple: 1.3 });
    });
  });

  // centre card: stacked column chart
  rect(s, { x: 4.947, y: 1.829, w: 4.231, h: 4.796, fill: { color: C.white }, shadow: shadow(88, 21, 0.15) });
  tx(s, 'Amazing main content', { x: 5.236, y: 2.116, w: 3.654, h: 0.404, fontSize: 18, align: 'center', color: C.ink });
  tbody(s, '2023-2024', { x: 5.38, y: 2.508, w: 3.366, h: 0.286, fontSize: 11, align: 'center', color: C.green });
  const stacks = [
    { x: 5.55, seg: [[3.209, 0.574, C.blue], [3.781, 0.693, C.teal], [4.473, 1.229, C.green]] },
    { x: 6.431, seg: [[3.138, 0.593, C.blue], [3.731, 1.266, C.teal], [4.997, 0.706, C.green]] },
    { x: 7.312, seg: [[3.325, 0.858, C.blue], [4.183, 0.522, C.teal], [4.705, 0.997, C.green]] },
    { x: 8.192, seg: [[3.414, 0.858, C.blue], [4.272, 0.636, C.teal], [4.908, 0.794, C.green]] },
  ];
  stacks.forEach((st, i) => {
    st.seg.forEach(sg => rect(s, { x: st.x, y: sg[0], w: 0.345, h: sg[1], fill: { color: sg[2] } }));
    tbody(s, 'Category ' + (i + 1), { x: st.x - 0.243, y: 5.74, w: 0.87, h: 0.32, fontSize: 10, align: 'center', color: C.g59 });
  });
  s.addShape('line', { x: 5.359, y: 5.702, w: 3.366, h: 0, line: { color: C.w95, width: 0.75 } });
}

/** isometric platform slab (rounded parallelogram seen from above) */
function slab(s, x, y, w, h, color, transparency) {
  polyBox(s, { x: x, y: y, w: w, h: h },
    [[0.017, 0.57], [0.56, 0.024], [0.983, 0.43], [0.44, 0.978]],
    { fill: { color: color, transparency: transparency || 0 } });
}

function slide12(pptx) {
  const s = pptx.addSlide();
  rect(s, { x: 0, y: 0, w: 5.635, h: 7.5, fill: { color: C.green } });
  dotGrid(s, 6.423, 1.035, 3.51, 3.12, 13, 12, 0.075, C.green, 92);

  slab(s, 5.87, 1.956, 5.798, 3.425, C.greenXL);
  slab(s, 5.895, 2.116, 5.748, 3.425, C.greenL);
  slab(s, 5.87, 2.68, 5.798, 3.425, C.green, 91);

  // Simplified isometric "airport lounge": three rows of seats on the same
  // projection as the platform, plus a handful of travellers and luggage.
  const ALONG = [0.50, 0.29];   // step to the next seat within a row
  const BACK = [0.46, -0.75];   // step from one row to the row behind it
  for (let row = 0; row < 3; row++) {
    for (let i = 0; i < 3; i++) {
      const x = 7.59 + i * ALONG[0] + row * BACK[0];
      const y = 3.77 + i * ALONG[1] + row * BACK[1];
      polyBox(s, { x: x, y: y, w: 0.57, h: 0.33 }, [[0, 0.5], [0.5, 0], [1, 0.5], [0.5, 1]], { fill: { color: 'FFAE26' } });
      polyBox(s, { x: x + 0.28, y: y - 0.45, w: 0.36, h: 0.50 },
        [[0, 0.62], [0.62, 0], [1, 0.22], [0.38, 0.86]], { fill: { color: 'EF8E00' } });
      rect(s, { x: x + 0.25, y: y + 0.26, w: 0.05, h: 0.30, fill: { color: '5E5966' } });
    }
  }
  const people = [
    [7.55, 3.62, '32293A'], [8.15, 3.15, 'D3293D'], [8.72, 3.42, 'D3293D'],
    [9.10, 2.48, '75AA00'], [9.62, 2.64, 'FFAE26'],
  ];
  people.forEach(p => {
    ellipse(s, { x: p[0], y: p[1] - 0.52, w: 0.24, h: 0.26, fill: { color: '32293A' } });
    polyBox(s, { x: p[0] - 0.06, y: p[1] - 0.30, w: 0.36, h: 0.44 },
      [[0.34, 0], [0.80, 0.12], [0.76, 1], [0.20, 1], [0.14, 0.12]], { fill: { color: p[2] } });
    polyBox(s, { x: p[0] - 0.14, y: p[1] + 0.10, w: 0.56, h: 0.50 },
      [[0.40, 0], [0.74, 0.06], [0.30, 0.52], [0.06, 0.94], [0.0, 0.72], [0.22, 0.26]], { fill: { color: '32293A' } });
  });
  roundRect(s, { x: 9.42, y: 3.55, w: 0.42, h: 0.60, rectRadius: 0.12, fill: { color: 'D3293D' } });
  roundRect(s, { x: 9.83, y: 3.28, w: 0.36, h: 0.50, rectRadius: 0.12, fill: { color: '75AA00' } });

  slab(s, 9.425, 4.446, 1.479, 0.935, C.greenXL);
  slab(s, 9.837, 4.903, 1.479, 0.935, C.greenXL);

  tx(s, '12,000+', { x: 0.853, y: 2.17, w: 4.087, h: 1.212, fontSize: 66, bold: true, color: C.white });
  tbody(s, '300+ more waters beginning saw which first signs at projects.',
    { x: 0.853, y: 3.344, w: 3.693, h: 0.722, fontSize: 16, color: C.white, lineSpacingMultiple: 1.2 });

  [4.419, 4.89, 5.36].forEach(y => {
    tbody(s, 'Amazing content here', { x: 1.188, y: y, w: 2.702, h: 0.362, fontSize: 14, color: C.white, lineSpacingMultiple: 1.2 });
    ellipse(s, { x: 0.906, y: y + 0.082, w: 0.201, h: 0.201, fill: { color: C.white } });
    s.addShape('rightArrow', { x: 0.958, y: y + 0.135, w: 0.098, h: 0.095, fill: { color: C.green }, line: NOLINE });
  });
}

function slide13(pptx) {
  const s = pptx.addSlide();
  rect(s, { x: 0, y: 0, w: 13.333, h: 2.238, fill: { color: C.green } });
  tx(s, 'Funnel Infographic', { x: 1.35, y: 0.665, w: 10.633, h: 0.909, fontSize: 48, align: 'center', color: C.white });

  // 4 stacked inverted trapezoids. Each sits on a slightly wider dark "rim" that
  // is only visible where the tapered band above it stops short, so all rims are
  // painted first and the bands then cover their middles.
  const bands = [
    { x: 4.358, y: 2.751, w: 4.617, h: 0.928, inset: 0.1397, rim: 0.173, rimInset: 0.0000, color: C.green, dark: C.greenDk, grad: C.greenD },
    { x: 4.777, y: 3.679, w: 3.782, h: 0.928, inset: 0.1695, rim: 0.169, rimInset: 0.0305, color: C.teal, dark: C.tealDk, grad: C.tealD },
    { x: 5.193, y: 4.606, w: 2.951, h: 0.932, inset: 0.2171, rim: 0.166, rimInset: 0.0395, color: C.blue, dark: C.blueDk, grad: C.blueD },
    { x: 5.608, y: 5.538, w: 2.118, h: 0.928, inset: 0.3033, rim: 0.170, rimInset: 0.0537, color: C.navy, dark: C.navyDk, grad: C.navyD },
  ];
  bands.forEach((b, i) => {
    if (i === 0) { // two dark spikes poking above the widest corners
      polyBox(s, { x: b.x, y: b.y - b.rim, w: 0.247, h: b.rim }, [[0, 1], [0.52, 0], [1, 1]], { fill: { color: b.dark } });
      polyBox(s, { x: b.x + b.w - 0.246, y: b.y - b.rim, w: 0.246, h: b.rim }, [[0, 1], [0.47, 0], [1, 1]], { fill: { color: b.dark } });
    } else {
      polyBox(s, { x: b.x, y: b.y - b.rim, w: b.w, h: b.rim },
        [[b.rimInset, 0], [1 - b.rimInset, 0], [1, 1], [0, 1]], { fill: { color: bands[i - 1].dark } });
    }
  });
  bands.forEach(b => {
    // gradient slices, each one clipped to the tapering silhouette
    gradientBand(s, { x: b.x, y: b.y, w: b.w, h: b.h }, b.color, b.grad, u => {
      if (u < b.inset) return [0, u / b.inset];
      if (u > 1 - b.inset) return [0, (1 - u) / b.inset];
      return [0, 1];
    }, 26);
  });

  [[8.463, 3.129], [7.628, 5.006], [4.458, 4.105], [5.292, 5.982]].forEach((p, i) => {
    roundRect(s, {
      x: p[0], y: p[1], w: 0.858, h: 0.056, rectRadius: 0.5, rotate: -54.53,
      fill: { color: i % 2 === 0 ? C.green : C.blue },
      shadow: tintShadow(i % 2 === 0 ? C.green : C.blue, 8, 3, 0.25),
    });
  });

  const blocks = [
    { x: 9.299, y: 2.649, title: 'Content title one', align: 'left' },
    { x: 8.493, y: 4.526, title: 'Content title three', align: 'left' },
    { x: 0.881, y: 3.625, title: 'Content title two', align: 'right' },
    { x: 1.716, y: 5.505, title: 'Content title four', align: 'right' },
  ];
  blocks.forEach(b => {
    tx(s, b.title, { x: b.x, y: b.y, w: 3.588, h: 0.37, fontSize: 16, align: b.align, color: C.ink });
    tbody(s, "Winged waters beginning saw which. First signs Can't may great seed image.",
      { x: b.x, y: b.y + 0.373, w: 3.588, h: 0.602, fontSize: 12, align: b.align, color: C.g40, lineSpacingMultiple: 1.3 });
  });

  [['\u2708', 2.966], ['\u263B', 3.9], ['\u2709', 4.818], ['\u2699', 5.736]].forEach(ic => {
    iconBadge(s, 6.404, ic[1], 0.526, C.white, ic[0], { fontSize: 13, glyphColor: C.white });
  });
}

/**
 * One stage of the 3-D cone funnel: an elliptical cap sitting on a tapered bowl.
 * Both carry a left-to-right gradient (bright <-> 50% luminance) which is painted
 * as vertical slices. The bowl wall runs straight in by ~0.36" per side down to
 * 80% height, then rounds off on a quarter ellipse to a flat bottom.
 */
function coneStage(s, x, y, w, h, capH, color, dark) {
  const capMid = capH / 2 / h;      // cap centre line, as a fraction of h
  const wall = 0.36 / w;            // horizontal travel of the straight wall
  const flat = 0.444;               // half width of the flat bottom
  const semi = u => Math.sqrt(Math.max(0, 1 - Math.pow(2 * u - 1, 2)));
  const bowlBottom = u => {
    const uu = u > 0.5 ? 1 - u : u;
    if (uu <= wall) return capMid + (0.80 - capMid) * (uu / wall);
    if (uu >= flat) return 1;
    return 0.80 + 0.20 * Math.sqrt(1 - Math.pow((flat - uu) / (flat - wall), 2));
  };
  gradientBand(s, { x: x, y: y, w: w, h: h }, color, dark,
    u => [capMid * (1 + semi(u)), bowlBottom(u)], 40);
  gradientBand(s, { x: x, y: y, w: w, h: capH }, dark, color,
    u => [0.5 - semi(u) / 2, 0.5 + semi(u) / 2], 40);
}

function slide14(pptx) {
  const s = pptx.addSlide();
  socialHeader(s, C.green);
  rect(s, { x: 5.73, y: 0, w: 6.619, h: 7.5, fill: { color: C.green } });
  dotGrid(s, 1.226, 3.28, 3.51, 3.12, 13, 12, 0.075, C.green, 86);

  coneStage(s, 1.155, 1.720, 4.132, 1.585, 0.859, C.green, C.greenDk);
  coneStage(s, 1.977, 2.798, 3.309, 1.448, 0.688, C.teal, C.tealDk);
  coneStage(s, 1.997, 3.918, 2.448, 1.271, 0.510, C.blue, C.blueDk);
  coneStage(s, 2.837, 5.029, 1.590, 1.092, 0.332, C.navy, C.navyDk);

  [['\u267B', 3.04, 2.776], ['\u2708', 3.452, 3.691], ['\u260E', 3.04, 4.632], ['\u2691', 3.452, 5.566]].forEach(ic => {
    ellipse(s, { x: ic[1], y: ic[2], w: 0.361, h: 0.361, fill: { color: C.white } });
    tx(s, ic[0], { x: ic[1], y: ic[2], w: 0.361, h: 0.361, align: 'center', valign: 'middle', fontSize: 11, color: C.blue, fontFace: FONT.icon });
  });

  tx(s, 'Funnel Diagram Infographic', { x: 6.122, y: 1.183, w: 5.945, h: 1.717, fontSize: 48, color: C.white });
  [[6.122, 3.234, 'Amazing Projects'], [6.122, 4.937, 'Cool Stuffs'],
   [9.203, 3.234, 'Amazing Projects'], [9.203, 4.937, 'Cool Stuffs']].forEach(b => {
    tx(s, b[2], { x: b[0], y: b[1], w: 2.726, h: 0.436, fontSize: 18, color: C.white, lineSpacingMultiple: 1.2 });
    tbody(s, "Winged waters beginning saw which. First signs Can't may great seed image. Two, sixth, darkness. ",
      { x: b[0], y: b[1] + 0.575, w: 2.9, h: 0.865, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3 });
  });
}

/** rounded diamond (an isometric square seen from above) */
function roundDiamond(s, x, y, w, h, color) {
  const P = (a, b) => ({ x: a * w, y: b * h });
  const cub = (a, b, x1, y1, x2, y2) =>
    ({ x: a * w, y: b * h, curve: { type: 'cubic', x1: x1 * w, y1: y1 * h, x2: x2 * w, y2: y2 * h } });
  s.addShape('custGeom', {
    x: x, y: y, w: w, h: h, line: NOLINE, fill: { color: color },
    points: [
      Object.assign({ moveTo: true }, P(0.9547, 0.4169)),
      cub(0.9547, 0.5831, 1.0, 0.4625, 1.0, 0.5375),
      P(0.5849, 0.9544),
      cub(0.4189, 0.9544, 0.5396, 1.0, 0.4660, 1.0),
      P(0.0453, 0.5831),
      cub(0.0453, 0.4169, 0.0, 0.5375, 0.0, 0.4625),
      P(0.4170, 0.0456),
      cub(0.5811, 0.0456, 0.4623, 0.0, 0.5358, 0.0),
      { close: true },
    ],
  });
}

/**
 * Isometric process tile: the top face plus the extruded side wall.
 * A lighter copy offset by a hair shows as the highlight line along the front edges.
 */
function isoTile(s, x, y, w, h, color) {
  const face = h * 0.861;
  roundDiamond(s, x, y + 0.225, w, face, color);
  roundDiamond(s, x, y + 0.034, w, face, mix(color, C.white, 0.6));
  roundDiamond(s, x, y, w, face, color);
}

function slide15(pptx) {
  const s = pptx.addSlide();
  socialHeader(s, C.green);
  polyBox(s, { x: 1.576, y: 0, w: 11.758, h: 7.5 }, [[0.638, 0], [1, 0], [1, 1], [0, 1]], { fill: { color: C.img } });
  imageLabel(s, 9.6, 6.0, 2.2, 0.4);

  tx(s, '15', {
    x: 10.928, y: 0.081, w: 2.039, h: 1.01, align: 'right', valign: 'middle',
    fontSize: 54, color: C.w65, transparency: 85,
  });

  const steps = [
    { n: '01', tile: [1.927, 5.062], color: C.green, num: [2.188, 5.317], tx: [4.882, 5.82], w: 6.957, title: 'Content title five', body: "Winged waters beginning saw which. First signs Can't may great seed image." },
    { n: '02', tile: [2.945, 4.029], color: C.teal, num: [3.269, 4.284], tx: [5.926, 4.783], w: 6.042, title: 'Content title four', body: "Winged waters beginning saw which. First signs Can't may great." },
    { n: '03', tile: [3.976, 2.997], color: C.blue, num: [4.29, 3.251], tx: [6.806, 3.745], w: 5.29, title: 'Content title three', body: "Winged waters beginning saw which. First signs Can't may." },
    { n: '04', tile: [4.995, 1.964], color: C.navy, num: [5.299, 2.252], tx: [7.72, 2.707], w: 4.377, title: 'Content title two', body: 'Winged waters beginning saw which. First signs.' },
    { n: '05', tile: [6.047, 0.931], color: C.purple, num: [6.375, 1.123], tx: [8.674, 1.669], w: 3.588, title: 'Content title one', body: 'Winged waters beginning saw which. ' },
  ];
  steps.forEach(st => isoTile(s, st.tile[0], st.tile[1], 2.405, 1.611, st.color));
  steps.forEach(st => {
    tx(s, st.n, { x: st.num[0], y: st.num[1], w: 1.785, h: 0.909, fontSize: 48, align: 'center', color: C.white });
    tx(s, st.title, { x: st.tx[0], y: st.tx[1], w: 3.588, h: 0.37, fontSize: 16, color: C.white });
    tbody(s, st.body, { x: st.tx[0], y: st.tx[1] + 0.373, w: st.w, h: 0.34, fontSize: 12, color: C.white, lineSpacingMultiple: 1.3 });
  });

  s.addText('Financialy  Process', { x: 1.155, y: 1.284, w: 3.728, h: 1.717, fontSize: 48, color: C.ink, fontFace: FONT.head, valign: 'top', margin: 0 });

  navButton(s, 1.27, 3.404, 0.363, C.green, 'left');
  navButton(s, 1.27, 3.963, 0.363, C.green, 'right');

  s.addShape('rtTriangle', { x: 12.472, y: 6.638, w: 0.862, h: 0.862, flipH: true, fill: { color: C.w65, transparency: 78 }, line: NOLINE });
  s.addShape('rtTriangle', { x: 12.472, y: 6.638, w: 0.428, h: 0.428, flipH: true, fill: { color: 'BFBFBF', transparency: 80 }, line: NOLINE });
}

/* -------------------------------------------------------------- build/run */

function build() {
  const pptx = new PptxGenJS();
  pptx.author = 'Financialy';
  pptx.title = 'Finance Presentation';
  pptx.defineLayout({ name: 'WIDE_16x9', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE_16x9';
  pptx.theme = { headFontFace: FONT.head, bodyFontFace: FONT.body };

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
   slide09, slide10, slide11, slide12, slide13, slide14, slide15].forEach(fn => fn(pptx));

  return pptx.writeFile({
    fileName: path.join(__dirname, '023bcb8a-2053-4f3f-9f81-078d3981c1e9_grok_final.pptx'),
  });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
