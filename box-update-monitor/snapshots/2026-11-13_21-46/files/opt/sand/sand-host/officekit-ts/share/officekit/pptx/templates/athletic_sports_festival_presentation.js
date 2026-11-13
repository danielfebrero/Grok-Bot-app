/**
 * "ATHLETIC – Sport Presentation Template" (25 slides) rebuilt with pptxgenjs.
 * Run: node 0a7ede93-5c0b-4ab8-9aa8-bb276aa178be_grok_final.js
 */
const path = require('path');
const pptxgen = require('pptxgenjs');

const W = 10, H = 5.625;

const C = {
  white: 'FFFFFF', black: '000000', ink: '2D2D2C',
  grey: '595959', greyText: '3F3F3F', pageNum: '7F7F7F',
  silver: 'BFBFBF', lightGrey: 'D8D8D8', paleGrey: 'F2F2F2',
  maroon: '590202', red: 'A60303', orange: 'F24405',
  amber: 'F27405', gold: 'F29F05', sand: 'FBC664',
  panelTop: '520101', panelBottom: '010000',
};

const FONT_H = 'Sora';              // display / headings
const FONT_B = 'Plus Jakarta Sans'; // body copy

// Google-Slides default text insets: 0.075" sides, 0.0375" top/bottom (in points).
const INS = [2.7, 5.4, 2.7, 5.4];

// ------------------------------------------------------------------ helpers

function hex2rgb(h) { return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)]; }
function mix(a, b, t) {
  return a.map((c, i) => Math.round(c + (b[i] - c) * t).toString(16).padStart(2, '0').toUpperCase()).join('');
}

/** Text box with the deck's default insets. */
function text(s, body, o) {
  s.addText(body, Object.assign({ fontFace: FONT_B, fontSize: 11, color: C.black, margin: INS, valign: 'top' }, o));
}

function rect(s, o) { s.addShape('rect', Object.assign({ line: { type: 'none' } }, o)); }

/** The deck's recurring dark panel: maroon at the top fading to near-black. */
function darkPanel(s, x, y, w, h) { vGrad(s, x, y, w, h, C.panelTop, C.panelBottom); }

/** Vertical colour ramp, banded (pptxgenjs writes solid fills only). */
function vGrad(s, x, y, w, h, top, bottom, steps) {
  steps = steps || 30;
  const a = hex2rgb(top), b = hex2rgb(bottom);
  for (let i = 0; i < steps; i++) {
    rect(s, { x: x, y: y + (h * i) / steps, w: w, h: h / steps + 0.006, fill: { color: mix(a, b, i / (steps - 1)) } });
  }
}

/** 135-degree ramp (dark top-right -> light bottom-left) as rotated bands. */
function diagGrad(s, from, to, steps) {
  steps = steps || 40;
  const a = hex2rgb(from), b = hex2rgb(to);
  const span = (W + H) / Math.SQRT2, len = span + 1;
  for (let i = 0; i < steps; i++) {
    const t = (i + 0.5) / steps;
    const d = (t - 0.5) * span;
    rect(s, {
      x: W / 2 - d / Math.SQRT2 - len / 2, y: H / 2 + d / Math.SQRT2 - span / steps,
      w: len, h: (span / steps) * 2 + 0.02, rotate: 45,
      fill: { color: mix(a, b, i / (steps - 1)) },
    });
  }
}

/** Elliptical glow radiating from a corner (slide 1 background). */
function cornerGlow(s, cx, cy, rx, ry, inner, outer, steps) {
  steps = steps || 22;
  const a = hex2rgb(inner), b = hex2rgb(outer);
  for (let i = steps - 1; i >= 0; i--) {
    const f = (i + 1) / steps;
    s.addShape('ellipse', {
      x: cx - rx * f, y: cy - ry * f, w: 2 * rx * f, h: 2 * ry * f,
      fill: { color: mix(a, b, i / (steps - 1)) }, line: { type: 'none' },
    });
  }
}

// The deck's motif: nested quarter rings hugging a corner point (cx, cy).
const QUAD = { br: [180, 270], bl: [270, 360], tl: [0, 90], tr: [90, 180] };

/** One quarter ring, split into colour segments to fake its gradient. */
function ring(s, corner, cx, cy, R, t, cols) {
  const [a0, a1] = QUAD[corner], n = cols.length;
  const flip = corner === 'br' || corner === 'tl';
  for (let i = 0; i < n; i++) {
    const k = flip ? n - 1 - i : i;
    s.addShape('blockArc', {
      x: cx - R, y: cy - R, w: 2 * R, h: 2 * R,
      angleRange: [a0 + ((a1 - a0) * k) / n, a0 + ((a1 - a0) * (k + 1)) / n - 0.01],
      arcThicknessRatio: t / R, fill: { color: cols[i] }, line: { type: 'none' },
    });
  }
}

/** Concentric ring family sharing a corner, thickness and colour ramp. */
function rings(s, corner, cx, cy, radii, t, cols) {
  radii.forEach((R) => ring(s, corner, cx, cy, R, t, cols));
}

/** Same family as thin sand-coloured outlines (slides 12 and 18). */
function ringsOutline(s, corner, cx, cy, radii, t) {
  const [a0, a1] = QUAD[corner];
  radii.forEach((R) => s.addShape('blockArc', {
    x: cx - R, y: cy - R, w: 2 * R, h: 2 * R, angleRange: [a0, a1 - 0.01],
    arcThicknessRatio: t / R, fill: { type: 'none' }, line: { color: C.sand, width: 0.75 },
  }));
}

/** Vector art tables: [fill, transparency?, ...flat x1,y1,x2,y2,... polygons]. */
function art(s, shapes) {
  shapes.forEach((sh) => {
    const fill = { color: sh[0] };
    let i = 1;
    if (typeof sh[1] === 'number') { fill.transparency = sh[1]; i = 2; }
    const pts = [];
    for (; i < sh.length; i++) {
      const f = sh[i];
      for (let k = 0; k < f.length; k += 2) {
        pts.push(k === 0 ? { x: f[0], y: f[1], moveTo: true } : { x: f[k], y: f[k + 1] });
      }
      pts.push({ close: true });
    }
    s.addShape('custGeom', { x: 0, y: 0, w: W, h: H, points: pts, fill: fill, line: { type: 'none' } });
  });
}

/** Placeholder standing in for a photograph in the reference deck. */
function photo(s, x, y, w, h, o) {
  o = o || {};
  s.addShape(o.shape || 'rect', Object.assign(
    { x: x, y: y, w: w, h: h, fill: { color: o.fill || C.lightGrey }, line: { type: 'none' } },
    o.rectRadius ? { rectRadius: o.rectRadius } : {}));
  s.addText('[image]', {
    x: x, y: y + h / 2 - 0.16, w: w, h: 0.32, align: 'center', valign: 'middle',
    fontFace: FONT_B, fontSize: 9, color: o.label || '9E9E9E', margin: 0,
  });
}

function footer(s, color) {
  s.addText([{ text: 'ATHLETIC SPORT ', options: { bold: true } },
             { text: '– presentation template', options: { bold: false } }],
    { x: 0.794, y: 5.235, w: 2.524, h: 0.202, fontFace: FONT_H, fontSize: 8,
      color: color || C.greyText, margin: INS, valign: 'top' });
}

function pageNum(s, n, color) {
  s.addText(String(n), { x: 8.712, y: 5.222, w: 0.422, h: 0.215, align: 'right',
    fontFace: FONT_H, fontSize: 8, color: color || C.pageNum, margin: INS, valign: 'top' });
}

/** Text rotated 90 degrees CCW; w/h are the unrotated box, as PPT stores them. */
function vText(s, body, o) {
  text(s, body, Object.assign({ rotate: 270 }, o));
}

/** Eyebrow label above a large Sora headline — used on most content slides. */
function titleBlock(s, x, y, eyebrow, title, o) {
  o = o || {};
  text(s, eyebrow, { x: x + 0.052, y: y, w: 3.285, h: 0.227, fontSize: 9, color: o.color || C.black });
  text(s, title, { x: x, y: y + 0.291, w: o.w || 3.794, h: o.h || 1.169, fontFace: FONT_H,
    fontSize: o.size || 36, color: o.color || C.black, lineSpacingMultiple: 0.9 });
}

/**
 * Number / icon badge: a square with only its top-right corner rounded,
 * filled with the amber-to-orange ramp (`rectRadius` is in inches here).
 */
function badge(s, x, y, size, label) {
  s.addShape('round1Rect', { x: x, y: y, w: size, h: size, rectRadius: size * 0.41,
    fill: { color: mix(hex2rgb(C.amber), hex2rgb(C.orange), 0.5) }, line: { type: 'none' } });
  if (label) {
    s.addText(label, { x: x, y: y, w: size, h: size, align: 'center', valign: 'middle',
      fontFace: FONT_B, fontSize: 14, color: C.white, margin: 0 });
  }
}

/**
 * Arrow-bullet list (the reference uses a Noto "▶" glyph).
 * pptxgenjs re-emits paragraph properties for every run, so the bullet has to
 * be repeated on each run of a paragraph or the last run would clear it.
 */
function arrowBullets(s, x, y, w, h, paras, o) {
  o = o || {};
  const bullet = { characterCode: '25B8', indent: o.indent || 10 };
  const body = [];
  paras.forEach((runs) => {
    runs.forEach((r, i) => body.push({
      text: r.t, options: { bold: !!r.b, bullet: bullet, breakLine: i === runs.length - 1 },
    }));
  });
  text(s, body, { x: x, y: y, w: w, h: h, fontSize: o.fontSize || 11, color: C.black,
    align: 'left', lineSpacingMultiple: 1.3 });
}
const WORLD_MAP = [
  ['D4D0C6', [6.28,3.02,5.99,3.03,5.75,2.62,5.74,2.29,5.35,2.22,5.26,2.32,4.97,2.18,4.95,2.04,4.45,2.11,4.08,2.65,4.09,2.98,4.33,3.23,4.95,3.28,5.08,3.82,5.02,4.06,5.27,4.69,5.53,4.63,5.69,4.44,5.76,4.13,5.96,3.94,5.91,3.59,6.28,3.02]],
  ['FCEDCF', [6.23,3.84,6.18,3.93,6.06,4.03,6.08,4.13,6.04,4.23,6.06,4.29,6.1,4.33,6.17,4.29,6.25,3.98,6.28,3.96,6.23,3.84]],
  ['F7CF83', [3.22,3.71,2.71,3.46,2.67,3.3,2.34,3.1,2.03,3.02,1.81,3.14,1.71,3.65,2.07,4.18,1.9,5.35,2.03,5.59,2.23,5.61,2.11,5.42,2.24,4.98,2.47,4.88,2.46,4.73,2.59,4.73,2.78,4.38,3.01,4.24,3.22,3.71]],
  ['F9DDA8', [10.79,0.62,8.1,0.23,8.34,0.03,8,-0.09,7.03,0.31,6.99,0.65,6.99,0.25,6.7,0.46,6.53,1.5,6.14,1.57,6.4,2.02,6.23,1.83,5.49,1.93,5.8,2.05,6.05,2.94,6.56,2.62,6.21,2.3,7,2.62,7.15,3.12,7.61,2.59,8.08,3.6,7.89,2.94,8.17,3.02,8.1,2.68,8.58,2.39,8.48,1.96,8.85,2.11,9.24,1.39,9.05,1.24,10.02,0.86,9.76,1.43,9.97,1,10.79,0.62]],
  ['FDF1DB', [8.49,3.26,8.5,3.21,8.47,3.17,8.44,3.17,8.24,3.35,8.2,3.36,8.2,3.44,8.23,3.5,8.29,3.52,8.42,3.54,8.51,3.36,8.48,3.3,8.49,3.26]],
  ['FAE3B7', [9.17,3.48,9.12,3.47,9.04,3.53,9.01,3.49,9,3.44,8.96,3.43,8.93,3.44,8.95,3.47,9,3.49,8.97,3.51,8.98,3.53,9,3.53,9.12,3.59,9.15,3.62,9.17,3.69,9.23,3.72,9.24,3.5,9.17,3.48]],
  ['FBF4E3', [9.25,1.87,9.21,1.9,9.2,1.98,9.17,2.03,9.1,2.07,9.07,2.11,8.97,2.12,8.91,2.16,8.97,2.15,8.97,2.21,9.03,2.18,9.01,2.15,9.02,2.13,9.05,2.14,9.07,2.18,9.09,2.12,9.13,2.13,9.22,2.11,9.27,1.94,9.25,1.87]],
  ['D8D8D8', [7.7,-0.25,7.74,-0.21,7.84,-0.19,7.86,-0.2,7.88,-0.25,7.87,-0.28,7.83,-0.27,7.81,-0.28,7.79,-0.31,7.8,-0.35,7.75,-0.38,7.68,-0.36,7.66,-0.32,7.63,-0.31,7.69,-0.28,7.7,-0.25]],
  ['FBE6BF', [9.15,0.1,9.32,0.09,9.35,0.08,9.36,0.06,9.27,0.02,9.23,0.05,9.18,0.02,9.14,0.02,9.12,0.04,9.12,0.06,9.12,0.08,9.15,0.1]],
  ['D8D8D8', [8.05,-0.16,8.07,-0.16,8.07,-0.17,8.04,-0.19,7.95,-0.22,7.9,-0.16,7.89,-0.12,7.93,-0.14,8.05,-0.16]],
  ['FCECCE', [9.27,1.36,9.28,1.46,9.28,1.65,9.3,1.63,9.29,1.57,9.3,1.52,9.34,1.51,9.31,1.41,9.31,1.32,9.27,1.36]],
  ['FCF3DF', [9.36,1.76,9.31,1.74,9.27,1.7,9.27,1.72,9.26,1.79,9.28,1.82,9.31,1.83,9.33,1.8,9.37,1.79,9.37,1.78,9.36,1.76]],
  ['FCF0D8', [8.21,3.66,8.13,3.63,8.07,3.64,8.16,3.67,8.34,3.7,8.28,3.66,8.21,3.66]],
  ['FDF6E8', [8.6,3.48,8.57,3.46,8.54,3.44,8.54,3.45,8.52,3.52,8.54,3.54,8.54,3.59,8.55,3.6,8.56,3.59,8.56,3.5,8.59,3.5,8.6,3.58,8.63,3.55,8.63,3.53,8.6,3.48]],
  ['FBE9C7', [8.57,2.88,8.58,2.89,8.58,2.91,8.6,2.91,8.59,2.85,8.62,2.82,8.62,2.81,8.61,2.77,8.59,2.76,8.57,2.77,8.57,2.83,8.55,2.86,8.57,2.88]],
  ['FCEDD2', [8.69,3.16,8.7,3.19,8.73,3.19,8.73,3.19,8.72,3.17,8.73,3.15,8.76,3.15,8.75,3.08,8.7,3.12,8.69,3.16]],
  ['F9D89B', [8.88,2.2,8.9,2.21,8.89,2.26,8.9,2.27,8.91,2.27,8.92,2.25,8.93,2.23,8.93,2.2,8.93,2.19,8.91,2.18,8.91,2.17,8.88,2.19,8.87,2.2,8.88,2.2]],
  ['FBE4BA', [9.55,0.09,9.48,0.07,9.41,0.07,9.48,0.1,9.5,0.1,9.55,0.09]],
  ['FADBA2', [2.92,1.55,2.58,0.99,2.28,0.86,2.09,0.87,2.06,1.4,1.58,1.03,1.73,0.72,2.03,0.81,1.89,0.69,2.03,0.44,2.28,0.58,2.11,0.78,2.56,0.85,2.47,0.67,2.64,0.62,2.12,0.19,1.76,0.2,1.81,0.53,1.59,0.32,1.7,0.17,1.55,0.17,1.55,0.52,1.35,0.55,1.2,0.22,0.67,0.13,0.55,0.3,0.96,0.4,0.23,0.48,-0.46,0.32,-0.8,0.51,-0.84,0.72,-0.64,0.79,-0.76,0.99,-0.52,1.11,-0.67,1.27,-0.17,0.98,0.26,1.12,0.98,2.54,2.01,3.18,1.95,2.86,1.77,2.83,1.84,2.63,1.55,2.73,1.5,2.43,1.82,2.34,2.05,2.53,2.26,1.92,2.68,1.72,2.55,1.5,2.92,1.55]],
  ['FCECCF', [3.94,0.43,3.83,0.22,4.04,0.14,3.96,-0.01,4.08,-0.06,3.96,-0.12,4.08,-0.28,3.99,-0.28,4.28,-0.4,3.94,-0.39,3.83,-0.44,3.94,-0.51,3.43,-0.58,3.19,-0.54,3.2,-0.43,2.71,-0.44,2.46,-0.31,2.51,-0.21,2.27,-0.13,2.42,0.02,2.78,0.1,2.84,0.33,3.01,0.44,2.91,0.68,3.19,1,3.36,0.72,3.94,0.43]],
  ['FEFEFE', [1.57,-0.22,1.67,-0.13,1.8,-0.16,1.74,-0.01,2.04,0.02,2.12,-0.01,2.09,-0.05,2.22,-0.22,2.67,-0.47,2.37,-0.54,2,-0.51,1.66,-0.42,1.83,-0.34,1.82,-0.28,1.58,-0.39,1.5,-0.3,1.57,-0.22]],
  ['FDF4E4', [1.6,0.01,1.64,0.04,1.65,0.11,1.72,0.13,2.02,0.14,2.06,0.1,2.06,0.07,2.01,0.05,1.77,0.06,1.67,-0.02,1.5,-0.05,1.53,0.01,1.6,0.01]],
  ['FCE9C7', [0.81,0.08,0.83,0.09,0.93,0.1,0.92,0.13,0.97,0.15,1.19,0.1,1.22,0.05,1.19,0.02,1.14,0.03,1.1,-0.03,1.05,-0,1.07,0.06,1.04,0.06,0.91,-0.01,0.86,0.01,0.81,0.08]],
  ['F7E9CA', [4.18,0.7,4.11,0.64,4.09,0.66,4,0.68,3.98,0.67,3.94,0.71,3.93,0.66,3.89,0.65,3.85,0.7,3.91,0.7,3.91,0.72,3.86,0.74,3.91,0.75,3.92,0.77,3.9,0.79,4.01,0.82,4.15,0.76,4.19,0.72,4.18,0.7]],
  ['FAE4B9', [1.47,0.17,1.36,0.17,1.34,0.21,1.34,0.22,1.37,0.24,1.33,0.23,1.32,0.24,1.32,0.28,1.38,0.31,1.42,0.36,1.5,0.32,1.52,0.3,1.52,0.27,1.49,0.23,1.5,0.18,1.49,0.17,1.47,0.17]],
  ['FBE3B8', [2.22,2.67,2,2.58,1.94,2.58,1.89,2.63,1.9,2.64,1.95,2.62,2.07,2.65,2.12,2.68,2.13,2.72,2.22,2.71,2.24,2.7,2.22,2.67]],
  ['D8D8D8', [1.22,-0.2,1.23,-0.18,1.23,-0.14,1.24,-0.13,1.32,-0.14,1.38,-0.1,1.43,-0.1,1.4,-0.18,1.33,-0.2,1.29,-0.23,1.21,-0.22,1.22,-0.2]],
  ['FAE2B4', [1.39,-0.02,1.34,-0,1.32,0.05,1.37,0.06,1.37,0.1,1.38,0.11,1.46,0.1,1.48,0.07,1.48,0.03,1.48,-0,1.44,-0.03,1.39,-0.02]],
  ['FCF7ED', [0.67,0.04,0.75,0.03,0.8,-0.02,0.82,0.01,0.85,-0,0.87,-0.07,0.85,-0.08,0.83,-0.07,0.78,-0.07,0.65,0,0.66,0.03,0.67,0.04]],
  ['FCF0D9', [2.41,2.74,2.36,2.71,2.28,2.7,2.27,2.71,2.28,2.76,2.28,2.76,2.27,2.74,2.25,2.74,2.23,2.76,2.24,2.77,2.31,2.78,2.33,2.79,2.35,2.77,2.42,2.77,2.43,2.75,2.41,2.74]],
  ['D8D8D8', [0.97,-0.08,1,-0.07,1.04,-0.08,1.07,-0.13,1.1,-0.14,1.1,-0.16,1.08,-0.17,1.05,-0.18,0.99,-0.18,0.97,-0.16,0.97,-0.14,1.06,-0.14,1.06,-0.13,0.96,-0.11,0.96,-0.09,0.97,-0.08]],
  ['FBE5BC', [1.51,0.09,1.52,0.11,1.57,0.13,1.59,0.13,1.6,0.12,1.61,0.11,1.61,0.08,1.6,0.07,1.57,0.06,1.55,0.06,1.53,0.06,1.51,0.08,1.51,0.09]],
  ['FCE7C2', [9.54,4.2,9.41,4.07,9.28,3.78,9.21,4.03,9.06,3.94,9.1,3.83,8.95,3.8,8.86,3.93,8.75,3.91,8.57,4.1,8.34,4.2,8.39,4.63,8.48,4.67,8.92,4.54,9.06,4.67,9.13,4.59,9.1,4.68,9.14,4.64,9.22,4.78,9.42,4.82,9.52,4.77,9.63,4.5,9.64,4.37,9.54,4.2]],
  ['FCE8C5', [9.28,3.73,9.3,3.73,9.31,3.7,9.36,3.67,9.47,3.76,9.55,3.77,9.47,3.71,9.43,3.65,9.46,3.64,9.45,3.62,9.33,3.53,9.24,3.5,9.23,3.72,9.28,3.73]],
  ['D8D8D8', [10.29,4.92,10.28,4.9,10.26,4.9,10.21,4.98,10.11,5.07,10.08,5.1,10.08,5.12,10.15,5.15,10.18,5.13,10.23,5.05,10.27,5.02,10.32,4.95,10.29,4.92]],
  ['D8D8D8', [10.45,4.77,10.43,4.79,10.41,4.79,10.37,4.75,10.34,4.75,10.34,4.8,10.31,4.84,10.35,4.87,10.35,4.9,10.34,4.92,10.35,4.93,10.37,4.92,10.42,4.84,10.46,4.81,10.45,4.77]],
  ['FBE5BB', [9.45,4.91,9.41,4.92,9.37,4.9,9.36,4.91,9.36,4.93,9.38,4.98,9.41,5.02,9.43,5.01,9.47,4.97,9.47,4.91,9.46,4.91,9.45,4.91]],
  ['FEFAF3', [9.6,3.58,9.61,3.57,9.59,3.56,9.57,3.59,9.55,3.6,9.48,3.6,9.5,3.62,9.52,3.62,9.54,3.62,9.57,3.61,9.6,3.58]],
  ['FBE5BE', [6.7,0.46,6.16,0.62,6.08,0.52,6.08,0.66,5.86,0.8,5.71,0.62,5.91,0.66,5.94,0.56,5.55,0.37,5.25,0.47,4.82,0.91,4.82,1.07,4.98,1.03,5.09,1.23,5.24,1.03,5.21,0.86,5.4,0.67,5.47,0.72,5.34,0.96,5.59,1,5.41,1.05,5.27,1.3,5,1.32,4.93,1.16,4.92,1.33,4.49,1.58,4.6,1.77,4.36,1.8,4.36,2.03,4.56,2.05,4.74,1.78,4.93,1.74,5.16,1.92,5.06,1.67,5.36,2.04,5.36,1.9,5.52,1.91,5.64,1.63,5.74,1.71,5.91,1.61,6,1.82,6.21,1.84,6.17,1.49,6.53,1.5,6.45,1.22,6.7,0.46]],
  ['D8D8D8', [5.04,-0.15,5.11,-0.1,5.09,-0.08,5.1,-0.07,5.18,-0.02,5.27,-0.16,5.33,-0.18,5.18,-0.28,5.16,-0.27,5.15,-0.22,5.12,-0.26,5,-0.25,5,-0.22,5.04,-0.15]],
  ['FCEDD2', [6.84,-0.04,6.52,0.06,6.32,0.32,6.38,0.33,6.4,0.38,6.49,0.38,6.44,0.31,6.46,0.21,6.62,0.08,6.86,-0,6.87,-0.04,6.84,-0.04]],
  ['FCF5E8', [4.48,1.25,4.52,1.25,4.55,1.31,4.54,1.34,4.5,1.35,4.51,1.38,4.47,1.41,4.55,1.41,4.54,1.45,4.65,1.44,4.65,1.4,4.68,1.38,4.63,1.35,4.58,1.25,4.54,1.2,4.56,1.14,4.54,1.13,4.5,1.13,4.52,1.08,4.48,1.09,4.45,1.13,4.48,1.25]],
  ['D8D8D8', [5.27,-0.26,5.32,-0.25,5.3,-0.25,5.3,-0.23,5.44,-0.22,5.52,-0.28,5.48,-0.3,5.38,-0.31,5.36,-0.28,5.29,-0.31,5.24,-0.29,5.25,-0.26,5.27,-0.26]],
  ['FBEBCC', [4.4,1.39,4.43,1.37,4.42,1.31,4.45,1.29,4.42,1.25,4.37,1.26,4.36,1.29,4.35,1.3,4.32,1.29,4.32,1.31,4.32,1.33,4.34,1.35,4.32,1.39,4.32,1.42,4.34,1.42,4.4,1.39]],
  ['D8D8D8', [5.43,-0.11,5.37,-0.14,5.35,-0.14,5.33,-0.13,5.34,-0.12,5.34,-0.11,5.32,-0.09,5.4,-0.08,5.42,-0.09,5.43,-0.11]],
];

const STADIUM = [
  ['A60303', [5,3.34,6.65,4.49,6.47,4.63,6.21,4.72,5,4.73,5,3.34]],
  ['590202', [6.64,2.2,6.83,2.69,6.83,4.09,6.65,4.49,5,3.34,6.64,2.2]],
  ['A60303', [5,1.97,6.21,1.97,6.47,2.07,6.64,2.2,5,3.34,5,1.97]],
  ['590202', [5,3.34,5,4.73,3.79,4.72,3.53,4.63,3.35,4.49,5,3.34]],
  ['A60303', [3.36,2.2,5,3.34,3.35,4.49,3.17,4.09,3.17,2.69,3.36,2.2]],
  ['590202', [3.89,1.97,5,1.97,5,3.34,3.36,2.2,3.61,2.02,3.89,1.97]],
  ['7C0202', [5,2.51,5.98,2.51,6.12,2.56,5,3.34]],
  ['420101', [6.12,2.56,6.21,2.74,6.21,3.95,6.13,4.13,5,3.34]],
  ['7C0202', [5,3.34,6.13,4.13,5.99,4.18,5,4.18]],
  ['420101', [5,3.34,5,4.18,4.01,4.18,3.87,4.13]],
  ['7C0202', [3.88,2.56,5,3.34,3.87,4.13,3.79,3.95,3.79,2.74,3.85,2.58]],
  ['420101', [4.02,2.51,5,2.51,5,3.34,3.88,2.56,4.02,2.51]],
  ['00A200', [3.97,2.67,4.26,2.67,4.26,4.02,3.97,4.02]],
  ['008000', [4.26,2.67,4.56,2.67,4.56,4.02,4.26,4.02]],
  ['00A200', [4.56,2.67,4.85,2.67,4.85,4.02,4.56,4.02]],
  ['008000', [4.85,2.67,5.15,2.67,5.15,4.02,4.85,4.02]],
  ['00A200', [5.15,2.67,5.44,2.67,5.44,4.02,5.15,4.02]],
  ['008000', [5.44,2.67,5.74,2.67,5.74,4.02,5.44,4.02]],
  ['00A200', [5.74,2.67,6.03,2.67,6.03,4.02,5.74,4.02]],
  ['FFFFFF', [5.87,3.86,4.13,3.86,4.13,2.84,5.87,2.84,5.87,3.86], [4.14,3.84,5.86,3.84,5.86,2.85,4.14,2.85,4.14,3.84]],
  ['FFFFFF', [4.37,3.58,4.13,3.57,4.13,3.12,4.38,3.12,4.37,3.58], [4.14,3.56,4.36,3.56,4.36,3.14,4.14,3.14,4.14,3.56]],
  ['FFFFFF', [4.23,3.46,4.13,3.45,4.13,3.24,4.24,3.24,4.24,3.45,4.23,3.46], [4.14,3.44,4.22,3.44,4.22,3.25,4.14,3.25,4.14,3.44]],
  ['FFFFFF', [4.13,3.46,4.07,3.46,4.04,3.43,4.04,3.27,4.07,3.24,4.14,3.24,4.14,3.45,4.13,3.46], [4.08,3.25,4.06,3.27,4.06,3.42,4.08,3.44,4.13,3.44,4.13,3.25,4.08,3.25]],
  ['FFFFFF', [4.3,3.34,4.31,3.34,4.31,3.35,4.31,3.35,4.3,3.36,4.29,3.36,4.29,3.35,4.29,3.34,4.3,3.34], [4.3,3.32,4.29,3.32,4.28,3.33,4.27,3.34,4.27,3.35,4.27,3.35,4.28,3.37,4.29,3.37,4.3,3.37,4.31,3.37,4.32,3.37,4.32,3.35,4.32,3.35,4.32,3.34,4.32,3.33,4.31,3.32,4.3,3.32]],
  ['FFFFFF', [4.37,3.43,4.36,3.42,4.36,3.27,4.37,3.27,4.39,3.27,4.41,3.3,4.43,3.35,4.41,3.4,4.39,3.42,4.37,3.43], [4.38,3.29,4.38,3.41,4.4,3.38,4.41,3.35,4.4,3.31,4.38,3.29]],
  ['FFFFFF', [4.17,3.86,4.13,3.86,4.13,3.86,4.13,3.85,4.13,3.81,4.13,3.81,4.13,3.81,4.15,3.81,4.16,3.81,4.17,3.82,4.18,3.83,4.18,3.84,4.18,3.85,4.18,3.86,4.17,3.86], [4.14,3.84,4.16,3.84,4.16,3.83,4.14,3.82,4.14,3.84]],
  ['FFFFFF', [5.87,3.58,5.62,3.57,5.63,3.12,5.87,3.12,5.87,3.58], [5.64,3.56,5.86,3.56,5.86,3.14,5.64,3.14,5.64,3.56]],
  ['FFFFFF', [5.87,3.46,5.76,3.45,5.77,3.24,5.87,3.24,5.88,3.45,5.87,3.46], [5.78,3.44,5.86,3.44,5.86,3.25,5.78,3.25,5.78,3.44]],
  ['FFFFFF', [5.92,3.46,5.86,3.45,5.86,3.25,5.87,3.24,5.93,3.24,5.96,3.27,5.96,3.43,5.94,3.45,5.92,3.46], [5.88,3.44,5.93,3.44,5.94,3.42,5.95,3.28,5.93,3.26,5.88,3.25,5.88,3.44]],
  ['FFFFFF', [5.7,3.34,5.71,3.34,5.71,3.35,5.71,3.35,5.7,3.36,5.7,3.36,5.69,3.35,5.7,3.34,5.7,3.34], [5.7,3.32,5.7,3.32,5.68,3.33,5.68,3.34,5.68,3.35,5.68,3.35,5.68,3.37,5.7,3.37,5.7,3.37,5.71,3.37,5.72,3.37,5.73,3.35,5.73,3.35,5.73,3.34,5.72,3.33,5.71,3.32,5.7,3.32]],
  ['FFFFFF', [5.63,3.43,5.61,3.42,5.59,3.4,5.57,3.35,5.59,3.3,5.61,3.27,5.63,3.27,5.64,3.27,5.64,3.42,5.63,3.43], [5.62,3.29,5.6,3.31,5.59,3.35,5.6,3.38,5.62,3.41,5.62,3.29]],
  ['FFFFFF', [5.87,3.86,5.83,3.86,5.82,3.86,5.82,3.85,5.82,3.84,5.83,3.83,5.83,3.82,5.84,3.81,5.85,3.81,5.87,3.81,5.87,3.81,5.87,3.81,5.87,3.85,5.87,3.86,5.87,3.86], [5.84,3.84,5.86,3.84,5.86,3.82,5.84,3.83,5.84,3.84]],
  ['FFFFFF', [5.87,2.89,5.85,2.89,5.84,2.89,5.83,2.88,5.83,2.87,5.82,2.86,5.82,2.84,5.82,2.84,5.83,2.84,5.87,2.84,5.87,2.84,5.87,2.84,5.87,2.88,5.87,2.89,5.87,2.89], [5.84,2.85,5.84,2.87,5.86,2.88,5.86,2.85,5.84,2.85]],
  ['FFFFFF', [5,3.86,5,2.84,5,3.86]],
  ['FFFFFF', [5,3.44,4.97,3.44,4.92,3.41,4.89,3.36,4.89,3.33,4.89,3.3,4.92,3.25,4.97,3.22,5,3.22,5.03,3.22,5.08,3.25,5.11,3.3,5.11,3.33,5.11,3.36,5.08,3.41,5.03,3.44,5,3.44], [5,3.23,4.98,3.24,4.93,3.26,4.91,3.31,4.9,3.33,4.91,3.36,4.93,3.4,4.98,3.43,5,3.43,5.03,3.43,5.07,3.4,5.1,3.36,5.1,3.33,5.1,3.31,5.07,3.26,5.03,3.24,5,3.23]],
  ['FFFFFF', 80, [4.62,3.54,5.2,3.48,5.69,3.3,6.03,3.05,6.03,2.67,3.97,2.67,3.97,3.46,4.62,3.54]],
  ['000000', 80, [5.74,2.67,3.97,2.67,3.97,4.02,6.03,4.02,6.03,2.67,5.74,2.67], [6,3.98,4.01,3.98,4.01,2.72,5.99,2.72,5.99,3.98]],
];

const VENUE_ICONS_ART = [
  ['FFFFFF', [2.71,2.49,2.74,2.49,2.71,2.49], [2.69,2.46,2.75,2.47,2.7,2.46], [2.73,2.43,2.71,2.43,2.73,2.45], [2.74,2.37,2.74,2.42,2.76,2.4,2.74,2.37], [2.69,2.37,2.68,2.41,2.71,2.39], [2.78,2.36,2.77,2.39,2.78,2.36], [2.71,2.35,2.72,2.38,2.73,2.35], [2.75,2.32,2.76,2.37,2.77,2.33], [2.69,2.32,2.65,2.35,2.67,2.37,2.7,2.34], [2.74,2.3,2.77,2.32,2.74,2.3], [2.72,2.3,2.71,2.34,2.74,2.32,2.72,2.3], [2.72,2.29,2.79,2.34,2.74,2.43,2.77,2.52,2.75,2.52,2.76,2.48,2.67,2.48,2.67,2.51,2.73,2.52,2.66,2.52,2.7,2.43,2.65,2.34,2.72,2.29]],
  ['FFFFFF', [2.69,4.3,2.73,4.34,2.73,4.3], [2.74,4.24,2.74,4.37,2.75,4.24], [2.74,4.16,2.74,4.23,2.75,4.16], [2.73,4.15,2.83,4.19,2.78,4.22,2.82,4.19,2.76,4.16,2.76,4.38,2.6,4.38,2.61,4.29,2.61,4.37,2.73,4.37,2.64,4.29,2.73,4.29,2.73,4.15]],
  ['FFFFFF', [7.2,2.46,7.2,2.51,7.24,2.48,7.28,2.51,7.28,2.46], [7.24,2.38,7.27,2.38,7.25,2.42,7.24,2.38], [7.37,2.34,7.37,2.47,7.29,2.47,7.33,2.42,7.33,2.46,7.36,2.46,7.36,2.41,7.32,2.41,7.36,2.4,7.37,2.34], [7.22,2.34,7.31,2.36,7.28,2.52,7.19,2.52,7.2,2.44,7.28,2.45,7.28,2.38,7.3,2.4,7.29,2.35,7.24,2.37,7.21,2.35,7.18,2.37,7.2,2.42,7.17,2.41,7.17,2.36,7.22,2.34], [7.33,2.34,7.35,2.34,7.34,2.38,7.33,2.34], [7.31,2.29,7.38,2.3,7.4,2.36,7.37,2.3,7.32,2.32,7.28,2.3,7.25,2.33,7.31,2.29]],
  ['FFFFFF', [7.3,3.3,7.3,3.32,7.3,3.3], [7.3,3.29,7.3,3.33,7.31,3.29], [7.25,3.29,7.25,3.33,7.26,3.29], [7.3,3.28,7.32,3.29,7.31,3.33,7.28,3.32,7.3,3.28], [7.25,3.28,7.27,3.34,7.24,3.33,7.25,3.28], [7.26,3.23,7.3,3.23,7.26,3.23], [7.23,3.23,7.2,3.3,7.23,3.27,7.23,3.36,7.33,3.36,7.36,3.24,7.31,3.23,7.28,3.26,7.23,3.23], [7.23,3.22,7.37,3.24,7.34,3.45,7.23,3.45,7.23,3.41,7.28,3.44,7.28,3.39,7.29,3.44,7.33,3.44,7.33,3.36,7.23,3.38,7.19,3.25,7.23,3.22]],
  ['FFFFFF', [7.24,4.27,7.27,4.31,7.24,4.34,7.21,4.31,7.25,4.32,7.24,4.27], [7.25,4.19,7.27,4.2,7.25,4.19], [7.25,4.16,7.28,4.18,7.25,4.16], [7.37,4.15,7.4,4.18,7.33,4.25,7.39,4.18,7.37,4.16,7.18,4.28,7.19,4.35,7.26,4.37,7.3,4.32,7.29,4.25,7.31,4.31,7.29,4.36,7.17,4.38,7.18,4.26,7.27,4.2,7.29,4.22,7.3,4.18], [7.29,4.15,7.29,4.18,7.29,4.15]],
  ['FFFFFF', [2.63,3.39,2.69,3.43,2.67,3.4], [2.7,3.36,2.68,3.4,2.72,3.44,2.76,3.4,2.74,3.36], [2.79,3.31,2.74,3.36,2.77,3.39,2.81,3.38,2.82,3.35], [2.64,3.31,2.61,3.35,2.62,3.38,2.67,3.39,2.68,3.32], [2.72,3.3,2.69,3.32,2.7,3.35,2.74,3.35,2.75,3.32], [2.63,3.27,2.61,3.33,2.64,3.31], [2.8,3.27,2.82,3.33,2.8,3.27], [2.68,3.23,2.64,3.26,2.65,3.31,2.68,3.31,2.71,3.26], [2.71,3.23,2.69,3.23,2.73,3.25,2.72,3.29,2.79,3.31,2.8,3.26,2.71,3.23], [2.71,3.22,2.8,3.25,2.83,3.33,2.8,3.41,2.77,3.4,2.75,3.43,2.77,3.43,2.72,3.45,2.64,3.42,2.6,3.34,2.63,3.26,2.71,3.22]],
];

const CLIMBER = [
  ['A60303', [2.49,1.21,2.49,1.61,2.34,1.79,2.24,2.23,2.58,2.38,2.86,2.36,2.85,2.23,2.76,2.19,2.76,2.07,2.8,2.03,2.5,1.09,2.49,1.21], [3.31,3.07,3.3,3.14,3.34,3.22,3.29,3.31,3.35,3.47,3.31,3.54,3.33,3.92,3.3,4.02,3.25,3.83,3.26,3.33,3.19,3.22,3.23,3.15,3.13,3.22,3.16,3.38,3.08,3.44,3,3.42,2.94,3.26,2.83,3.21,2.76,3.25,2.42,3.26,2.48,3.56,2.42,3.82,2.47,4.23,2.39,4.26,2.29,4.21,2.17,4.21,2.3,4.34,2.47,4.94,2.61,5.23,2.62,5.42,2.51,5.62,0,5.62,0,0,2.66,0,2.59,0.14,2.6,0.32,2.51,0.53,2.57,0.75,2.51,1.05,2.82,2.01,2.95,2,3.02,2.13,3.14,2.12,3.15,2.03,3.06,1.87,3.07,1.74,3.12,1.63,3.22,1.59,3.31,1.63,3.36,1.71,3.38,1.86,3.33,1.95,3.43,1.97,3.53,2.11,3.5,2.58,3.44,2.81,3.57,2.85,3.54,3.07,3.44,3.22,3.35,3.19,3.34,3.11,3.41,2.9,3.31,3.07], [3.29,3.1,3.27,3.15,3.29,3.1], [3.3,3.41,3.3,3.5,3.33,3.46,3.3,3.41], [3.29,3.18,3.29,3.23,3.32,3.22,3.29,3.18], [3.24,3.19,3.22,3.23,3.24,3.19], [3.1,3.26,3.09,3.33,3.12,3.39,3.15,3.36,3.1,3.26], [3.06,3.27,3.07,3.39,3.06,3.27], [2.99,3.29,3.03,3.38,3.03,3.29,2.99,3.29], [2.89,2.35,2.91,2.33,2.88,2.27,2.89,2.35], [2.93,2.32,2.97,2.28,2.89,2.24,2.93,2.32], [2.99,2.13,2.9,2.13,2.85,2.02,2.95,2.03,2.99,2.13], [2.16,2.48,2.09,2.69,1.92,2.92,1.88,3.13,1.77,3.27,1.79,3.48,1.93,3.65,1.98,3.82,2.16,2.48], [2,3.87,2.08,4.04,2.25,3.96,2.28,3.88,2.21,3.34,2.24,3.08,2.33,3,2.59,2.91,2.61,2.83,2.68,2.79,2.47,2.63,2.32,2.59,2.18,2.5,2,3.87], [2.3,1.84,2.22,2.23,2.19,2.22,2.16,2.2,2.17,1.99,2.21,1.92,2.3,1.84]],
];

const CONTACT_ART = [
  ['C42204', [7.01,1.94,7.08,1.97,7.11,2.04,7.08,2.11,7.01,2.14,6.93,2.11,6.9,2.04,6.93,1.97,7.01,1.94], [7.01,1.89,6.9,1.93,6.85,2.04,6.9,2.15,7.01,2.2,7.11,2.15,7.16,2.04,7.11,1.93,7.01,1.89], [7.01,1.76,7.08,1.77,7.2,1.84,7.27,1.97,7.28,2.04,7.21,2.31,7.01,2.65,6.8,2.31,6.73,2.04,6.74,1.97,6.81,1.84,6.93,1.77,7.01,1.76]],
  ['F29F05', [3.55,4.44,3.52,4.45,3.54,4.42,3.51,4.43,3.48,4.42,3.46,4.42,3.43,4.44,3.42,4.47,3.42,4.47,3.37,4.46,3.32,4.42,3.31,4.46,3.34,4.49,3.31,4.48,3.31,4.5,3.33,4.52,3.36,4.53,3.33,4.53,3.34,4.55,3.38,4.57,3.34,4.59,3.3,4.59,3.34,4.6,3.38,4.61,3.42,4.6,3.46,4.59,3.5,4.53,3.52,4.47,3.55,4.44]],
  ['F29F05', [1.56,4.41,1.39,4.41,1.37,4.42,1.36,4.43,1.36,4.6,1.37,4.62,1.39,4.63,1.47,4.63,1.47,4.55,1.45,4.55,1.45,4.51,1.47,4.51,1.48,4.47,1.5,4.46,1.52,4.45,1.54,4.45,1.54,4.49,1.51,4.5,1.51,4.51,1.54,4.51,1.54,4.55,1.51,4.55,1.51,4.63,1.56,4.63,1.58,4.62,1.58,4.6,1.58,4.43,1.58,4.42,1.56,4.41]],
];
// ------------------------------------------------------------------- slides

function slide01(s) {                                   // title / cover
  s.background = { color: '0C0C0C' };
  cornerGlow(s, 0, H, 8.6, 4.9, '840101', '0C0C0C', 24);
  rings(s, 'br', 6.243, 5.643, [2.632, 1.88, 1.127], 0.377, ['DC8204', 'C76704', 'B04B03', '9A3304', '841E05', '700F05']);
  rings(s, 'tl', 6.243, 2.636, [2.632, 1.88, 1.127], 0.377, ['671006', '781F06', '8F3405', 'A84B03', 'C36704', 'DB8304']);
  rings(s, 'br', 10, 2.638, [2.632, 1.88, 1.127], 0.377, ['DA8405', 'C06805', 'A74B04', '8F3305', '7A1F06', '671106']);
  rings(s, 'bl', 0, 6.079, [2.632, 1.88, 1.127], 0.377, ['D16B03', 'CD6503', 'BB4B03', 'AB3103', '9C1C03', '930D02']);
  text(s, 'ATHLETIC', { x: 0.866, y: 0.704, w: 5.741, h: 1.363, fontFace: FONT_H, fontSize: 77, color: C.white });
  text(s, 'SPORT PRESENTATION TEMPLATE', { x: 0.866, y: 1.868, w: 5.741, h: 0.395, fontFace: FONT_H, fontSize: 15, color: C.amber, lineSpacingMultiple: 1.3 });
  text(s, [{ text: 'Presented by:', options: { breakLine: true } }, { text: ' John Doe', options: { color: C.amber } }],
    { x: 0.866, y: 2.806, w: 1.321, h: 0.429, fontFace: FONT_H, fontSize: 11, color: C.white });
  text(s, [{ text: 'Date:', options: { breakLine: true } }, { text: '03 August 2025', options: { color: C.amber } }],
    { x: 3.076, y: 2.806, w: 1.321, h: 0.606, fontFace: FONT_H, fontSize: 11, color: C.white });
}

function slide02(s) {                                   // introduction
  darkPanel(s, 0, 0, 3.405, H);
  rings(s, 'tl', 0.002, 0.005, [2.632, 1.88, 1.127], 0.377, ['A95303', 'B95603', 'C95404', 'D65003', 'E24A03', 'E54904']);
  rings(s, 'tr', 3.41, 2.636, [2.632, 1.88, 1.127], 0.377, ['2C0000', '330300', '420900', '530F01', '631601', '6C1901']);
  rings(s, 'bl', 3.41, 5.637, [2.632, 1.88, 1.127], 0.377, ['F25813', 'F37B2F', 'F59846', 'F5AD5B', 'F6BB69', 'F7C271']);
  text(s, 'INTRODUCTION', { x: 4.085, y: 0.765, w: 1.281, h: 0.256, fontSize: 9, lineSpacingMultiple: 1.3 });
  text(s, [{ text: 'Welcome to the ' }, { text: 'Athletic', options: { bold: true, color: C.maroon } }],
    { x: 4.085, y: 1.238, w: 5.218, h: 1.434, fontFace: FONT_H, fontSize: 50, valign: 'middle', lineSpacingMultiple: 0.8 });
  text(s, 'Athletic Festival is a global celebration of sports, bringing together the best athletes from around the world. Join us as we witness record-breaking performances and unforgettable moments.',
    { x: 5.813, y: 2.889, w: 3.321, h: 1.198, lineSpacingMultiple: 1.3 });
  footer(s, C.white);
  pageNum(s, 2);
}

const TOC = [
  ['01', 'LOREM IPSUM DOLOR'], ['05', 'CONSECTETUR ADIPISCING'], ['07', 'SED DO EIUSMOD'],
  ['10', 'TEMPOR INCIDIDUNT UT'], ['13', 'LABORE ET DOLORE'],
];

function slide03(s) {                                   // table of content
  darkPanel(s, 0, 0, 3.405, H);
  rings(s, 'br', 2.603, 2.624, [2.632, 1.88], 0.377, ['CE6D03', 'C25803', 'B94A03', 'B64203', 'B74403', 'BB4A03']);
  rings(s, 'tr', 1.478, 2.625, [1.13], 0.377, ['CE6D03', 'C25803', 'B94A03', 'B64203', 'B74403', 'BB4A03']);
  rings(s, 'br', 1.478, 6.387, [2.632, 1.88, 1.127], 0.377, ['A74503', 'A74503', 'A24003', '9D3A03', '943102', '8E2C02']);
  vText(s, 'TABLE OF CONTENT', { x: 0.627, y: 2.328, w: 3.794, h: 1.306, fontFace: FONT_H,
    fontSize: 41, color: C.white, lineSpacingMultiple: 0.9 });
  TOC.forEach(([num, label], i) => {
    const y = 1.123 + i * 0.7675;
    text(s, label, { x: 4.315 + i * 0.009, y: y, w: 3.368, h: 0.311, fontSize: 18, margin: [0.57, 0, 0, 0] });
    text(s, num, { x: 8.701 + i * 0.009, y: y, w: 0.683, h: 0.311, fontSize: 18, margin: [0.57, 0, 0, 0] });
    if (i < TOC.length - 1) s.addShape('line', { x: 4.324, y: y + 0.412, w: 4.8, h: 0, line: { color: C.ink, width: 0.82 } });
  });
  pageNum(s, 3);
}

function slide04(s) {                                   // celebrating sport event
  rings(s, 'br', 2.603, 5.626, [2.632, 1.88, 1.127], 0.377, ['F19A13', 'F4952F', 'F59446', 'F5955C', 'F69868', 'F79B71']);
  rings(s, 'bl', 2.605, 6.002, [2.632, 1.88, 1.127], 0.377, ['F39725', 'F49530', 'F59446', 'F6955B', 'F69869', 'F79B71']);
  text(s, 'Celebrating Sport Event', { x: 0.724, y: 0.793, w: 3.761, h: 1.439, fontSize: 41 });
  [[0.871, 1.212, 0.902, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit'],
   [2.575, 2.915, 0.690, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna'],
   [4.066, 4.407, 0.690, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna'],
  ].forEach(([yTitle, yBody, h, body], i) => {
    const x = i === 2 ? 5.903 : 5.846;
    text(s, 'Insert Title Here', { x: x, y: yTitle, w: 2.813, h: 0.252, bold: true });
    text(s, body, { x: x, y: yBody, w: 3.288, h: h, lineSpacingMultiple: 1.2 });
  });
  footer(s);
  pageNum(s, 4);
}

function slide05(s) {                                   // one slide paragraph
  rings(s, 'tr', 10, 0.005, [2.825, 2.017, 1.21], 0.404, ['F25813', 'F47C2F', 'F59A48', 'F6AE5B', 'F6BC6A', 'F7C272']);
  text(s, 'One slide paragraph', { x: 0.866, y: 1.134, w: 5.999, h: 0.533, fontFace: FONT_H, fontSize: 30, lineSpacingMultiple: 0.9 });
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin pharetra nonummy pede. Mauris et orci. Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. Nunc viverra',
    { x: 0.866, y: 2.057, w: 8.268, h: 1.235, fontSize: 9, align: 'justify', lineSpacingMultiple: 1.3 });
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. Nunc viverra imperdiet enim. Fusce est. Vivamus a',
    { x: 0.866, y: 3.575, w: 4.134, h: 0.841, fontSize: 9, lineSpacingMultiple: 1.3 });
  arrowBullets(s, 5.448, 3.575, 3.686, 0.841, [
    [{ t: 'Sed ut perspiciatis ', b: true }, { t: 'unde omnis iste Error sit voluptatem accusan dolor sit amet, consectetuer adipiscing ' }],
    [{ t: 'Dolor emque laudantium, totam rem Eaque ipsa quae ab illo Nunc viverra imperdiet enim. Fusce est' }],
  ], { fontSize: 9, indent: 9 });
  footer(s);
  pageNum(s, 5);
}

function slide06(s) {                                   // history and legacy
  photo(s, -0.016, 0, 5.008, 2.813);
  photo(s, -0.016, 2.812, 5.008, 2.813);
  rings(s, 'tr', 6.129, 2.99, [2.632, 1.88], 0.377, ['DAD9D9', 'DED4D0', 'DFCAC2', 'E9C3B5', 'F2BAA6', 'F6AA8E']);
  rings(s, 'br', 5.004, 2.99, [1.13], 0.377, ['DAD9D9', 'DED4D0', 'DFCAC2', 'E9C3B5', 'F2BAA6', 'F6AA8E']);
  rings(s, 'tr', 5.01, -0.773, [2.638, 1.884, 1.127], 0.378, ['E0C8BF', 'E0C8BF', 'E1C0B5', 'E2B5A5', 'E3A893', 'E4A188']);
  titleBlock(s, 5.485, 1.05, 'Building on Rich Tradition', 'History And Legacy');
  text(s, [
    { text: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.', options: { breakLine: true } },
    { text: ' ', options: { breakLine: true } },
    { text: 'Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus. Pellentesque habitant morbi tristique senectus et netus et' },
  ], { x: 5.537, y: 2.701, w: 3.58, h: 1.329, fontSize: 9, lineSpacingMultiple: 1.2 });
  pageNum(s, 6);
}

const FEATURED_BODY = [{ t: 'Athletes from over ' }, { t: '150 nations ', b: true },
  { t: "will compete, representing diverse cultures and showcasing global talent. This year's festival is set to be the most inclusive yet." }];
const FEATURED_BULLETS = [
  [{ t: 'Sed ut perspiciatis ', b: true }, { t: 'unde omnis iste Error sit voluptatem accusan dolor sit ' }],
  [{ t: 'Dolor emque laudantium, totam rem Eaque ipsa quae ab illo Nunc viverra' }],
];

function slide07(s) {                                   // featured sport (marathon)
  titleBlock(s, 0.866, 1.136, 'WRITE SOMETHING HERE', 'Featured Sport');
  text(s, FEATURED_BODY.map((r) => ({ text: r.t, options: { bold: !!r.b } })),
    { x: 0.918, y: 2.906, w: 2.957, h: 0.784, fontSize: 9, lineSpacingMultiple: 1.2 });
  arrowBullets(s, 0.918, 3.728, 3.035, 0.968, FEATURED_BULLETS);
  rings(s, 'tr', 6.377, 0.005, [2.632, 1.88, 1.127], 0.377, ['E98355', 'EA9B68', 'EBAD76', 'EBB37C', 'EBB37C', 'EBB37C']);
  photo(s, 4.876, 0, 2.56, 3.5);
  photo(s, 7.445, 0, 2.56, 3.5);
  darkPanel(s, 4.876, 3.505, 5.128, 2.114);
  rect(s, { x: 4.563, y: 3.275, w: 2.546, h: 0.674, fill: { color: C.white } });
  badge(s, 4.387, 3.097, 0.752);
  s.addText('\u2691', { x: 4.475, y: 3.23, w: 0.533, h: 0.533, align: 'center', valign: 'middle', fontSize: 22, color: C.white, margin: 0 });
  text(s, 'Marathon', { x: 5.321, y: 3.441, w: 1.986, h: 0.35, fontFace: FONT_H, fontSize: 18, lineSpacingMultiple: 0.9 });
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet',
    { x: 5.32, y: 4.115, w: 4.08, h: 0.602, fontSize: 9, color: C.white, lineSpacingMultiple: 1.2 });
  footer(s);
  pageNum(s, 7);
}

const SPORTS = [['Basketball', '\u25CE'], ['Archery', '\u25C9'], ['Thriatlon', '\u2248'], ['Football', '\u2B24']];

function slide08(s) {                                   // featured sport list
  darkPanel(s, 6.595, 0, 3.405, H);
  rings(s, 'tr', 9.998, 0.005, [2.632, 1.88, 1.127], 0.377, ['E74803', 'D35003', 'BC5A0C', 'AF5301', 'A65402', 'A05303']);
  rings(s, 'tl', 6.59, 2.636, [2.632, 1.88, 1.127], 0.377, ['6F1600', '784237', '603B37', '523837', '503837', '4E3837']);
  rings(s, 'br', 6.59, 5.637, [2.632, 1.88, 1.127], 0.377, ['F15813', 'F37C2F', 'F49846', 'F3AC5A', 'F1B564', 'EDB668']);
  SPORTS.forEach(([label, icon], i) => {
    const y = 1.093 + i * 1.0045;
    rect(s, { x: 6.004, y: y, w: 2.744, h: 0.674, fill: { color: C.white } });
    badge(s, 5.828, y - 0.178, 0.752);
    s.addText(icon, { x: 5.999, y: y - 0.012, w: 0.405, h: 0.405, align: 'center', valign: 'middle', fontSize: 17, color: C.white, margin: 0 });
    text(s, label, { x: 6.762, y: y + 0.166, w: 1.986, h: 0.35, fontFace: FONT_H, fontSize: 18, lineSpacingMultiple: 0.9 });
  });
  titleBlock(s, 0.866, 1.136, 'WRITE SOMETHING HERE', 'Featured Sport');
  text(s, FEATURED_BODY.map((r) => ({ text: r.t, options: { bold: !!r.b } })),
    { x: 0.918, y: 2.906, w: 3.58, h: 0.602, fontSize: 9, lineSpacingMultiple: 1.2 });
  arrowBullets(s, 0.918, 3.728, 3.035, 0.968, FEATURED_BULLETS);
  footer(s);
  pageNum(s, 8);
}

function slide09(s) {                                   // the venue
  photo(s, 0, 0, 10, 3.5);
  rings(s, 'br', 10, 3.524, [2.632, 1.88, 1.127], 0.377, ['F05510', 'EC7528', 'EA8E3C', 'E8A04D', 'E7AB5A', 'E6B160']);
  rings(s, 'tl', 0.86, 0, [1.883, 1.13], 0.377, ['F05612', 'EC7628', 'EA8E3C', 'E89F4D', 'E7AA59', 'E6B05E']);
  titleBlock(s, 0.866, 4.007, 'WRITE SOMETHING HERE', 'The Venue', { h: 0.624 });
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus. Pellentesque habitant morbi',
    { x: 4.427, y: 3.978, w: 4.69, h: 0.784, fontSize: 9, lineSpacingMultiple: 1.2 });
  footer(s);
  pageNum(s, 9);
}

const HIGHLIGHTS = [
  ['01', 'THRILLS OF THE TRACK', 0.803],
  ['02', 'PRECISION AND FOCUS', 2.271],
  ['03', 'GRACE AND STRENGTH', 3.74],
];

function slide10(s) {                                   // write your title here
  photo(s, 2.406, 0.417, 4.172, 5.208);
  rings(s, 'tr', 5.043, 2.99, [2.632, 1.88], 0.377, ['F29305', 'F28205', 'F27105', 'F26205', 'F25605', 'F24B05']);
  rings(s, 'br', 3.918, 2.99, [1.13], 0.377, ['F29305', 'F28205', 'F27105', 'F26205', 'F25605', 'F24B05']);
  rings(s, 'tr', 3.924, -0.773, [2.638, 1.884, 1.127], 0.378, ['F27105', 'F27105', 'F26605', 'F25D05', 'F25305', 'F24E05']);
  vText(s, [{ text: 'Write your', options: { breakLine: true } }, { text: 'Title Here' }],
    { x: -0.605, y: 2.045, w: 4.314, h: 1.169, fontFace: FONT_H, fontSize: 36, lineSpacingMultiple: 0.9 });
  HIGHLIGHTS.forEach(([num, title, y]) => {
    badge(s, 6.346, y + 0.03, 0.608, num);
    text(s, title, { x: 7.069, y: y, w: 2.286, h: 0.278, fontFace: FONT_H, fontSize: 12, bold: true });
    text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue',
      { x: 7.069, y: y + 0.34, w: 1.933, h: 0.608, fontSize: 9, lineSpacingMultiple: 1.2 });
  });
  footer(s);
  pageNum(s, 10);
}

function slide11(s) {                                   // opening ceremony
  photo(s, 0, 2.365, 10, 3.267);
  darkPanel(s, 0, 0, 10, 2.375);
  rings(s, 'br', 6.204, 6.759, [2.635, 1.883], 0.377, ['F26905', 'F26B05', 'F27205', 'F27E05', 'F28C05', 'F29605']);
  rings(s, 'bl', 6.204, 5.634, [1.127], 0.377, ['F26905', 'F26B05', 'F27205', 'F27E05', 'F28C05', 'F29605']);
  rings(s, 'br', 9.978, 5.634, [2.635, 1.883, 1.13], 0.377, ['F24D05', 'F26205', 'F27505', 'F28505', 'F28E05', 'F29405']);
  text(s, 'Opening Ceremony', { x: 0.866, y: 0.603, w: 3.794, h: 1.169, fontFace: FONT_H, fontSize: 36, color: C.white, lineSpacingMultiple: 0.9 });
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus. ',
    { x: 4.174, y: 0.795, w: 4.96, h: 0.784, fontSize: 9, color: C.white, lineSpacingMultiple: 1.2 });
  text(s, '50+', { x: 7.475, y: 2.612, w: 2.174, h: 1.173, fontFace: FONT_H, fontSize: 72, color: C.white, align: 'center', lineSpacingMultiple: 0.9 });
  text(s, 'PERFORMANCE ARTIST', { x: 7.649, y: 3.557, w: 1.599, h: 0.478, color: C.white, lineSpacingMultiple: 1.2 });
}

function slide12(s) {                                   // participating nation
  art(s, WORLD_MAP);
  ringsOutline(s, 'tr', 9.071, 0.005, [2.632, 1.88, 1.127], 0.377);
  ringsOutline(s, 'tl', 5.663, 2.636, [2.632, 1.88, 1.127], 0.377);
  ringsOutline(s, 'br', 5.663, 5.637, [2.632, 1.88, 1.127], 0.377);
  ringsOutline(s, 'tl', 0.024, 0.001, [2.632, 1.88, 1.127], 0.377);
  text(s, [{ text: 'Participating', options: { breakLine: true } }, { text: 'Nation' }],
    { x: 0.866, y: 3.578, w: 3.794, h: 1.169, fontFace: FONT_H, fontSize: 36, lineSpacingMultiple: 0.9 });
  [4.772, 6.385, 7.998].forEach((x) => {
    text(s, ['Lorem ipsum', 'consectetuer ', 'Maecenas ', 'pulvinar ', 'ultricies,'].map((t) => ({
      text: t, options: { breakLine: true, bullet: { characterCode: '2022', indent: 9 } },
    })), { x: x, y: 3.68, w: 1.402, h: 0.966, fontSize: 9, align: 'left', lineSpacingMultiple: 1.2 });
  });
  footer(s);
  pageNum(s, 12);
}

const TICKETS = [
  { x: 4.273, tx: 4.523, bx: 4.549, plan: 'BASIC', price: '$79,00', fg: C.black, bullet: C.amber,
    shape: 'roundRect', fill: { color: C.white }, radius: 0 },
  { x: 7.044, tx: 7.293, bx: 7.32, plan: 'PREMIUM', price: '$89,00', fg: C.white, bullet: C.black,
    shape: 'round1Rect', fill: { color: C.amber }, radius: 0.287 },
];
const TICKET_ROWS = ['FIRST DISCUSSION', 'SECOND DISCUSSION ', 'THIRD DISCUSSION', 'SECOND DISCUSSION '];

function slide13(s) {                                   // ticketing
  rings(s, 'tr', 4.987, 0.005, [2.632, 1.88, 1.127], 0.377, ['F25813', 'F47C2F', 'F59A48', 'F5AF5C', 'F3B867', 'F4BA69']);
  darkPanel(s, 5, -0.073, 5, 5.715);
  TICKETS.forEach((t) => {
    s.addShape(t.shape, Object.assign({ x: t.x, y: 1.444, w: 2.521, h: 3.125, fill: t.fill, line: { type: 'none' } },
      t.radius ? { rectRadius: t.radius } : {}));
    text(s, t.plan, { x: t.tx, y: 1.793, w: 1.7, h: 0.303, fontFace: FONT_H, fontSize: 14, color: t.fg });
    text(s, t.price, { x: t.tx, y: 2.051, w: 1.881, h: 0.608, fontFace: FONT_H, fontSize: 30, color: t.fg, lineSpacingMultiple: 1.1 });
    s.addShape('line', { x: t.x, y: 2.772, w: 2.521, h: 0, line: { color: C.white, width: 0.75, transparency: 76 } });
    text(s, TICKET_ROWS.map((r) => ({ text: r, options: { breakLine: true, bullet: { characterCode: '2714', indent: 17 } } })),
      { x: t.bx, y: 2.959, w: 2.071, h: 1.304, color: t.fg, align: 'left', lineSpacingMultiple: 1.8 });
  });
  text(s, 'Ticketing', { x: 0.772, y: 3.007, w: 3.232, h: 0.624, fontFace: FONT_H, fontSize: 36, lineSpacingMultiple: 0.9 });
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit',
    { x: 0.839, y: 3.748, w: 2.896, h: 0.841, fontSize: 9, lineSpacingMultiple: 1.3 });
  footer(s);
  pageNum(s, 13, C.orange);
}

function slide14(s) {                                   // pull quote
  photo(s, 0, 0, 5, 5.625, { fill: 'D9D9D9' });
  darkPanel(s, 5, -0.073, 5, 5.715);
  rings(s, 'br', 10, 5.625, [2.635, 1.883, 1.13], 0.377, ['653E01', '452301', '301000', '240600', '1F0200', '1E0000']);
  rings(s, 'br', 4.994, 3.766, [1.13], 0.378, ['F28F05', 'F28B05', 'F27F05', 'F27305', 'F26005', 'F24C05']);
  rings(s, 'tr', 5, 0, [2.638, 1.884, 1.127], 0.378, ['F29505', 'F28005', 'F26C05', 'F25E05', 'F25305', 'F24D05']);
  text(s, [{ text: 'He who has a ' }, { text: 'why to live ', options: { color: C.gold } }, { text: 'can bear almost any how.' }],
    { x: 5.401, y: 1.227, w: 3.733, h: 2.297, fontFace: FONT_H, fontSize: 33, color: C.white });
  text(s, 'John Doe', { x: 5.401, y: 4.034, w: 1.483, h: 0.303, fontFace: FONT_H, fontSize: 14, color: C.white });
  text(s, 'HEAT – Sport Presentation Template', { x: 0.778, y: 5.104, w: 3.233, h: 0.227, fontFace: FONT_H, fontSize: 9, color: C.greyText });
}

const PLAYER_STATS = [
  ['HEIGHT', '190 CM', 5.379, 2.644, 1.131], ['WEIGHT', '60 KG', 6.798, 2.644, 1.131],
  ['AGE', '28', 8.198, 2.644, 1.131], ['REGION', 'SHEFFIELD, UK', 5.379, 3.352, 1.887],
  ['POSITION', 'FW', 8.198, 3.352, 1.131],
];

function slide15(s) {                                   // player profile
  photo(s, 0.87, 0.393, 4.13, 5.232);
  rings(s, 'tr', 10, 0.005, [2.825, 2.017, 1.21], 0.404, ['F25813', 'F47C2F', 'F59A48', 'F6AE5C', 'F6BC6A', 'F7C272']);
  rings(s, 'tr', 2.673, 2.99, [2.632, 1.88], 0.377, ['F29305', 'F28305', 'F27105', 'F26305', 'F25605', 'F24C05']);
  rings(s, 'br', 1.548, 2.99, [1.13], 0.377, ['F29305', 'F28305', 'F27105', 'F26305', 'F25605', 'F24C05']);
  rings(s, 'tr', 1.554, -0.777, [2.638, 1.884, 1.127], 0.378, ['F26405', 'F26405', 'F26005', 'F25A05', 'F25305', 'F24D05']);
  text(s, [{ text: 'BAYU', options: { breakLine: true } }, { text: 'ARRUMI' }],
    { x: 5.178, y: 0.963, w: 3.525, h: 1.187, fontFace: FONT_H, fontSize: 41, bold: true, color: C.orange, lineSpacingMultiple: 0.8 });
  PLAYER_STATS.forEach(([label, value, x, y, w]) => {
    s.addShape('triangle', { x: x - 0.118, y: y + 0.071, w: 0.107, h: 0.081, rotate: 90, fill: { color: C.red }, line: { type: 'none' } });
    text(s, [{ text: label, options: { fontSize: 9, breakLine: true } }, { text: value, options: { fontSize: 15, bold: true } }],
      { x: x, y: y, w: w, h: 0.48, fontFace: FONT_H });
  });
  [['200+', 'GOAL SCORED', 5.261], ['86%', 'SHOT ACCURACY', 7.609]].forEach(([big, small, x]) => {
    text(s, big, { x: x, y: 4.026, w: 1.525, h: 0.631, fontFace: FONT_H, fontSize: 33, bold: true, color: C.orange });
    text(s, small, { x: x, y: 4.572, w: 1.419, h: 0.227, fontFace: FONT_H, fontSize: 9 });
  });
  pageNum(s, 15);
}

const TEAM = [['JOHN DOE', 1.235, 1.619, 1.645], ['JOHN DOE', 3.909, 4.293, 1.645], ['JANE DOE', 6.626, 6.967, 1.684]];

function slide16(s) {                                   // meet the team
  darkPanel(s, 0, 2.828, 10, 2.797);
  rings(s, 'br', 2.635, 3.617, [2.635, 1.883, 1.13], 0.377, ['E97C4B', 'E98A4B', 'E9964B', 'E99B4B', 'E99F4B', 'E9A14B']);
  rings(s, 'tl', 7.364, 0, [2.638, 1.884, 1.127], 0.378, ['F24E05', 'F26205', 'F27605', 'F28305', 'F28A05', 'F28D05']);
  text(s, 'Meet the Team', { x: 2.552, y: 0.78, w: 4.896, h: 0.682, fontFace: FONT_H, fontSize: 36, align: 'center' });
  TEAM.forEach(([name, px, tx, py]) => {
    photo(s, px, py, 2.182, 2.603, { shape: 'round1Rect', rectRadius: 0.167 });
    text(s, name, { x: tx, y: 4.509, w: 1.414, h: 0.285, fontFace: FONT_H, fontSize: 15, bold: true,
      color: C.white, align: 'center', lineSpacingMultiple: 0.8 });
    text(s, 'Job Position', { x: tx, y: 4.748, w: 1.414, h: 0.266, fontSize: 11, color: C.silver,
      align: 'center', lineSpacingMultiple: 1.2 });
  });
}

const STANDINGS_COLS = [
  { key: '#', x: 1.028, w: 0.631, dx: 1.028, dw: 0.631, align: 'center' },
  { key: 'PLAYER NAME', x: 1.941, w: 2.253, dx: 1.941, dw: 1.694, align: 'left' },
  { key: 'PLAY', x: 4.179, w: 1.053, dx: 4.239, dw: 0.862, align: 'center' },
  { key: 'WIN', x: 5.073, w: 1.053, dx: 5.16, dw: 0.862, align: 'center' },
  { key: 'LOSE', x: 6.031, w: 1.053, dx: 6.136, dw: 0.862, align: 'center' },
  { key: 'DRAW', x: 6.962, w: 1.053, dx: 7.058, dw: 0.862, align: 'center' },
  { key: 'PTS', x: 7.856, w: 1.053, dx: 7.952, dw: 0.862, align: 'center' },
];
const STANDINGS = [
  ['1', 'Steven Johnson', '15', '10', '1', '4', '90'],
  ['2', 'Patrick Wilson', '15', '10', '2', '3', '88'],
  ['3', 'Eduardo Claim', '15', '9', '2', '4', '84'],
  ['4', 'Joe Carlos', '15', '9', '3', '3', '82'],
  ['5', 'Esteban Perez', '15', '7', '3', '5', '70'],
  ['6', 'Addison Alves', '15', '6', '3', '6', '65'],
];

function slide17(s) {                                   // table standings
  rings(s, 'bl', 7.175, 5.629, [2.825, 2.017, 1.21], 0.404, ['F25713', 'F47A2D', 'F49F50', 'F4B26D', 'F5BD78', 'F4C185']);
  rings(s, 'tr', 2.825, 0.005, [2.825, 2.017, 1.21], 0.404, ['F25813', 'F47C2F', 'F59A47', 'EFA556', 'EFAD5E', 'F0AF60']);
  text(s, 'Table standings', { x: 1.844, y: 0.78, w: 6.312, h: 0.682, fontFace: FONT_H, fontSize: 36, align: 'center' });
  s.addShape('roundRect', { x: 0.958, y: 1.659, w: 8.083, h: 0.552, rectRadius: 0.058, fill: { color: C.maroon }, line: { type: 'none' } });
  STANDINGS_COLS.forEach((c) => {
    const head = c.key === '#'
      ? { x: 1.028, y: 1.771, w: 0.631, h: 0.328, fontSize: 15 }
      : { x: c.x, y: 1.821, w: c.w, h: 0.227, fontSize: 9 };
    text(s, c.key, Object.assign({ color: C.white, align: c.align, valign: 'middle' }, head));
  });
  STANDINGS.forEach((row, r) => {
    const y = 2.224 + r * 0.441;
    s.addShape('roundRect', { x: 0.958, y: y, w: 8.083, h: 0.428, rectRadius: 0.058, fill: { color: C.paleGrey }, line: { type: 'none' } });
    STANDINGS_COLS.forEach((c, i) => {
      text(s, row[i], { x: c.dx, y: y + 0.088, w: c.dw, h: 0.252, fontSize: 11, color: C.greyText, align: c.align, valign: 'middle' });
    });
  });
  footer(s);
  pageNum(s, 17);
}

function slide18(s) {                                   // process
  art(s, CLIMBER);
  ringsOutline(s, 'tl', 0.024, 0.001, [2.632, 1.88, 1.127], 0.377);
  vText(s, 'PROCESS', { x: -1.159, y: 2.519, w: 4.135, h: 0.587, fontFace: FONT_H,
    fontSize: 36, color: C.white, align: 'center', lineSpacingMultiple: 0.9 });
  ['01', '02', '03', '04'].forEach((num, i) => {
    const y = 0.922 + i * 1.011;
    text(s, num, { x: 4.348, y: y, w: 1.052, h: 0.833, fontFace: FONT_H, fontSize: 45 });
    text(s, 'Insert Title Here', { x: 5.524, y: y + 0.102, w: 2.286, h: 0.252, fontSize: 11 });
    text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna',
      { x: 5.524, y: y + 0.316, w: 3.69, h: 0.608, fontSize: 9, lineSpacingMultiple: 1.2 });
  });
  footer(s);
  pageNum(s, 18);
}

const GAUGES = [
  { x: 4.1, value: '74B', title: 'Build With Ease', end: 67.7 },
  { x: 6.752, value: '52B', title: 'Daily Journey', end: 138.5 },
];

function slide19(s) {                                   // our value year by year
  rings(s, 'bl', 3.466, 5.625, [2.825, 2.017, 1.21], 0.404, ['F15612', 'EB752B', 'E78035', 'E7863A', 'ED9D4F', 'F7C170']);
  rings(s, 'tr', 10, 0.005, [2.825, 2.017, 1.21], 0.404, ['F25813', 'F1772C', 'EF8537', 'EE9648', 'F2B969', 'F6C272']);
  text(s, 'Our value year by year', { x: 0.834, y: 1.519, w: 2.951, h: 1.715, fontFace: FONT_H, fontSize: 36, lineSpacingMultiple: 0.9 });
  text(s, 'Lorem ipsum dolor sit amet, adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, ',
    { x: 0.834, y: 3.368, w: 2.208, h: 0.602, fontSize: 9, lineSpacingMultiple: 1.2 });
  GAUGES.forEach((g) => {
    rect(s, { x: g.x, y: 0.948, w: 2.346, h: 3.728, fill: { color: C.white } });
    s.addShape('line', { x: g.x, y: 1.72, w: 2.346, h: 0, line: { color: C.grey, width: 2.5, transparency: 92 } });
    text(s, g.title, { x: g.x + 0.218, y: 1.219, w: 1.911, h: 0.328, fontFace: FONT_H, fontSize: 15, align: 'center' });
    s.addShape('ellipse', { x: g.x + 0.425, y: 2.176, w: 1.497, h: 1.497, fill: { type: 'none' }, line: { color: C.lightGrey, width: 16, transparency: 20 } });
    s.addShape('arc', { x: g.x + 0.425, y: 2.176, w: 1.497, h: 1.497, angleRange: [270, g.end], fill: { type: 'none' }, line: { color: C.orange, width: 10.5 } });
    text(s, g.value, { x: g.x + 0.536, y: 2.542, w: 1.274, h: 0.682, fontFace: FONT_H, fontSize: 36, align: 'center' });
    text(s, '/year', { x: g.x + 0.942, y: 3.043, w: 0.463, h: 0.258, fontSize: 9, align: 'center', lineSpacingMultiple: 1.3 });
    text(s, 'What has fingers but cannot use them? ', { x: g.x + 0.218, y: 3.913, w: 1.911, h: 0.429, align: 'center' });
  });
  footer(s);
  pageNum(s, 19);
}

// Icon discs: [x, y, disc colour] — the white pictogram inside each comes
// from VENUE_ICONS_ART.
const VENUE_DISCS = [
  [2.497, 2.183, C.maroon], [2.497, 3.112, C.red], [2.497, 4.042, '420101'],
  [7.062, 2.183, C.red], [7.062, 3.112, '420101'], [7.062, 4.042, C.red],
];
const VENUE_LABELS = [
  [3.563, 2.202, 'right'], [5.136, 2.202, 'left'], [3.563, 4.386, 'right'], [5.136, 4.386, 'left'],
];

function slide20(s) {                                   // tournament venue
  art(s, STADIUM);
  s.addShape('roundRect', { x: 3.677, y: 2.418, w: 2.645, h: 1.861, rectRadius: 0.17,
    fill: { color: C.white, transparency: 85 }, line: { type: 'none' } });
  text(s, 'WRITE SOMETHING HERE', { x: 3.358, y: 0.749, w: 3.285, h: 0.215, fontSize: 8, align: 'center' });
  text(s, 'Tournament Venue', { x: 0.717, y: 0.886, w: 8.565, h: 0.682, fontFace: FONT_H, fontSize: 36, align: 'center' });
  [[0.572, 'right'], [7.57, 'left']].forEach(([x, align]) => {
    [2.167, 3.096, 4.026].forEach((y) => {
      text(s, 'Far away, behind the word mountains, far from',
        { x: x, y: y, w: 1.926, h: 0.5, fontSize: 9, color: C.grey, align: align, lineSpacingMultiple: 1.5 });
    });
  });
  VENUE_DISCS.forEach(([x, y, fill]) => {
    s.addShape('ellipse', { x: x, y: y, w: 0.441, h: 0.441, fill: { color: fill }, line: { type: 'none' } });
  });
  art(s, VENUE_ICONS_ART);
  VENUE_LABELS.forEach(([x, y, align]) => {
    text(s, 'WEST VALUE 03', { x: x, y: y, w: 1.301, h: 0.227, fontFace: FONT_H, fontSize: 9, bold: true,
      color: C.white, align: align, valign: 'middle' });
  });
  [[5.91, 270], [2.767, 90]].forEach(([x, rot]) => {
    text(s, 'WEST VALUE 03', { x: x, y: 3.235, w: 1.301, h: 0.227, rotate: rot, fontFace: FONT_H, fontSize: 9,
      bold: true, color: C.white, align: 'center', valign: 'middle' });
  });
  footer(s);
  pageNum(s, 20);
}

const MATCHES = [
  { day: '22', cardX: 0.974, cardY: 0.75, photoX: 5.0, photoY: 0.75 },
  { day: '23', cardX: 5.0, cardY: 2.24, photoX: 0.974, photoY: 2.24 },
  { day: '25', cardX: 0.974, cardY: 3.73, photoX: 5.0, photoY: 3.729 },
];

function slide21(s) {                                   // match schedule
  rings(s, 'bl', 7.365, 5.637, [2.635, 1.883], 0.377, ['F8F2E6', 'F8EDD7', 'F7E7C8', 'F5DEB4', 'F4E0B9', 'F4E2BF']);
  rings(s, 'br', 7.365, 4.512, [1.127], 0.377, ['F8F2E6', 'F8EDD7', 'F7E7C8', 'F5DEB4', 'F4E0B9', 'F4E2BF']);
  rings(s, 'bl', 3.598, 4.512, [2.635, 1.883, 1.13], 0.377, ['EED9B4', 'EED9B4', 'EDD8B0', 'EDD8B0', 'ECC378', 'EABF6E']);
  MATCHES.forEach((m) => {
    photo(s, m.photoX, m.photoY, 4.026, 1.24);
    rect(s, { x: m.cardX, y: m.cardY, w: 4.026, h: 1.24, fill: { color: C.white } });
    text(s, 'AUGUST', { x: m.cardX - 0.044, y: m.cardY + 0.169, w: 1.271, h: 0.278, fontFace: FONT_H,
      fontSize: 12, bold: true, align: 'center' });
    text(s, m.day, { x: m.cardX - 0.025, y: m.cardY + 0.352, w: 1.233, h: 0.833, fontFace: FONT_H,
      fontSize: 45, bold: true, color: C.red, align: 'center' });
    s.addShape('line', { x: m.cardX + 1.138, y: m.cardY + 0.14, w: 0, h: 0.961, line: { color: 'A5A5A5', width: 1, transparency: 65 } });
    text(s, 'GREEN FC VS RED FC', { x: m.cardX + 1.293, y: m.cardY + 0.202, w: 2.841, h: 0.328,
      fontFace: FONT_H, fontSize: 15, bold: true, color: C.red });
    ['Content Number One', 'Content Number Two'].forEach((t, i) => {
      const y = m.cardY + 0.59 + i * 0.228;
      s.addShape('roundRect', { x: m.cardX + 1.382, y: y + 0.048, w: 0.119, h: 0.119, rectRadius: 0.03,
        fill: { color: C.amber }, line: { type: 'none' } });
      text(s, t, { x: m.cardX + 1.511, y: y, w: 1.653, h: 0.215, fontSize: 8, color: C.grey, valign: 'middle' });
    });
  });
  footer(s);
  pageNum(s, 21);
}

function slide22(s) {                                   // gallery
  photo(s, 0.866, 0.893, 2.556, 1.853);
  photo(s, 3.592, 0.893, 2.048, 1.853);
  photo(s, 5.811, 0.893, 3.323, 1.853);
  photo(s, 0.866, 2.879, 5.118, 1.862);
  s.addShape('round1Rect', { x: 6.148, y: 2.879, w: 2.986, h: 1.883, rectRadius: 0.369,
    fill: { color: C.orange }, line: { type: 'none' } });
  text(s, 'Our gallery', { x: 6.415, y: 3.152, w: 2.298, h: 0.379, fontFace: FONT_H, fontSize: 18, color: C.white });
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna',
    { x: 6.415, y: 3.631, w: 2.298, h: 0.849, fontSize: 9, color: C.white, lineSpacingMultiple: 1.3 });
  footer(s);
  pageNum(s, 22);
}

function slide23(s) {                                   // mockup slide
  rings(s, 'bl', 3.658, 3.236, [2.825, 2.017, 1.21], 0.404, ['7B6F5A', '7B6F5A', '7B6F5A', 'AC9975', 'AB966E', '9D917A']);
  rings(s, 'tr', 3.658, 0.005, [2.825, 2.017, 1.21], 0.404, ['FEFEFD', 'FDF9EF', 'FCF0DA', 'FBE6C0', 'FADEAA', 'F9D99D']);
  // Tablet mock-up: dark bezel with a screen placeholder, plus its stylus.
  s.addShape('roundRect', { x: 3.66, y: 1.43, w: 7.24, h: 4.5, rectRadius: 0.05, fill: { color: '1A1A1A' }, line: { type: 'none' } });
  s.addShape('roundRect', { x: 4.6, y: 1.14, w: 3.4, h: 0.1, rectRadius: 0.5, fill: { color: 'E8E8E8' }, line: { type: 'none' }, rotate: 353 });
  photo(s, 3.891, 1.664, 6.781, 4.2, { shape: 'roundRect', rectRadius: 0.014 });
  text(s, [{ text: 'Mockup', options: { breakLine: true } }, { text: 'slide' }],
    { x: 0.518, y: 2.593, w: 2.363, h: 1.287, fontFace: FONT_H, fontSize: 36, lineSpacingMultiple: 1.0 });
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce',
    { x: 0.494, y: 4.047, w: 2.492, h: 0.644, fontSize: 9, lineSpacingMultiple: 1.3 });
  footer(s);
  pageNum(s, 23);
}

function slide24(s) {                                   // get in touch
  photo(s, 4.37, -0.009, 5.63, 5.63);
  rings(s, 'tl', 1.896, 2.99, [2.632, 1.88], 0.377, ['EAE7E1', 'EAE7E1', 'EBE7E0', 'EBE4D6', 'EEE2CA', 'EED9B2']);
  rings(s, 'bl', 3.021, 2.99, [1.13], 0.377, ['EAE7E1', 'EAE7E1', 'EBE7E0', 'EBE4D6', 'EEE2CA', 'EED9B2']);
  rings(s, 'tl', 3.021, -0.777, [2.638, 1.884, 1.127], 0.378, ['F0DBB5', 'F0DBB5', 'F2E5CE', 'F3EDE2', 'FEFBF6', 'FEFCFA']);
  rings(s, 'br', 10, 5.625, [2.632, 1.88, 1.127], 0.377, ['F05611', 'EC7528', 'EA8E3D', 'E8A04D', 'E6AB59', 'E6B161']);
  art(s, CONTACT_ART);
  text(s, 'Get in touch', { x: 0.788, y: 0.9, w: 1.79, h: 1.287, fontFace: FONT_H, fontSize: 36, lineSpacingMultiple: 1.0 });
  s.addShape('round1Rect', { x: 0.866, y: 2.652, w: 5.147, h: 2.327, rectRadius: 0.396,
    fill: { color: C.white }, line: { type: 'none' } });
  text(s, 'Overview', { x: 1.301, y: 2.965, w: 2.187, h: 0.221, bold: true, valign: 'bottom', lineSpacingMultiple: 0.8 });
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet ',
    { x: 1.301, y: 3.249, w: 3.793, h: 0.53, fontSize: 9 });
  s.addShape('line', { x: 1.284, y: 3.856, w: 4.312, h: 0, line: { color: C.white, width: 0.75 } });
  text(s, 'Social Media', { x: 1.301, y: 4.044, w: 2.187, h: 0.221, bold: true, valign: 'bottom', lineSpacingMultiple: 0.8 });
  [1.658, 3.733].forEach((x) => text(s, 'Ut wisi enim ad', { x: x, y: 4.406, w: 1.174, h: 0.227, fontSize: 9 }));
  s.addText('f', { x: 1.361, y: 4.395, w: 0.224, h: 0.24, align: 'center', valign: 'middle',
    fontFace: FONT_H, fontSize: 11, bold: true, color: C.white, margin: 0 });
  footer(s);
  pageNum(s, 24);
}

function slide25(s) {                                   // thank you
  diagGrad(s, '010000', C.panelTop);
  rings(s, 'tl', 6.243, 2.636, [2.632, 1.88, 1.127], 0.377, ['932C02', 'AB4103', 'C15A03', 'D26F04', 'DE8104', 'E58B03']);
  rings(s, 'br', 10, 2.638, [2.632, 1.88, 1.127], 0.377, ['E58C04', 'CC6A04', 'B55002', 'A23D02', '933102', '8B2C02']);
  rings(s, 'tl', 0.002, 0.003, [2.632, 1.88, 1.127], 0.377, ['9B2B03', 'B14103', 'C55A04', 'D46F03', 'E08203', 'E68B03']);
  text(s, 'THANK YOU', { x: 0.769, y: 2.867, w: 4.231, h: 2.174, fontFace: FONT_H, fontSize: 77,
    color: C.white, lineSpacingMultiple: 0.8 });
  text(s, [{ text: 'Presented by:', options: { breakLine: true } }, { text: ' John Doe' }],
    { x: 3.842, y: 0.885, w: 1.321, h: 0.429, fontFace: FONT_H, fontSize: 11, color: C.white });
  text(s, [{ text: 'Date:', options: { breakLine: true } }, { text: '03 August 2025' }],
    { x: 6.158, y: 0.885, w: 1.321, h: 0.606, fontFace: FONT_H, fontSize: 11, color: C.white });
}

// --------------------------------------------------------------------- main

const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
  slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25];

function build() {
  const pres = new pptxgen();
  pres.defineLayout({ name: 'ATHLETIC', width: W, height: H });
  pres.layout = 'ATHLETIC';
  pres.author = 'ATHLETIC';
  pres.title = 'ATHLETIC – Sport Presentation Template';
  BUILDERS.forEach((fn) => fn(pres.addSlide()));
  return pres.writeFile({ fileName: path.join(__dirname, '0a7ede93-5c0b-4ab8-9aa8-bb276aa178be_grok_final.pptx') });
}

build().then((f) => console.log('wrote', f));
