#!/usr/bin/env node
/*
 * "Futuristic Presentation Template" - 25 slides, 13.333 x 7.5 in (16:9).
 * Rebuilt from scratch with pptxgenjs only (no external assets).
 *
 * Photographs in the original deck are replaced by "[image]" placeholder
 * plates that keep the original position, size and clipped outline.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ================================================================= palette */

const C = {
  cyan:   '22EDEA',  // accent1 - headline highlight
  green:  '66FF33',  // accent5 - gradient partner of cyan
  mint:   '44F68E',  // visual mid-point of the cyan -> green gradient
  purple: 'A62EE1',  // accent3 - panel body
  lilac:  'CA82ED',  // accent3 lightened - panel highlight
  orchid: 'B858E7',  // mid point of the lilac -> purple button gradient
  plum:   '801AB2',  // darkest stop of the pie wedges
  white:  'FFFFFF',
  grey:   'D9D9D9',
  ltGrey: 'F2F2F2',
  dimGrey:'9240C0',  // 15%-opaque grey seen over the purple background
  yellow: 'FAFF00',
};

const FONT = 'Saira';
const SLIDE_W = 13.3333;
const SLIDE_H = 7.5;

/* ============================================== background (master slide) */

// The deck background is a 70-degree ramp from bright purple (top-left) to
// near-black (bottom-right). pptxgenjs has no gradient fill, so the ramp is
// painted as a stack of rotated half-plane bands sampled from these stops.
const BG_ANGLE = 70;
const BG_STOPS = [
  [0.000, '801AB2'], [0.111, '801AB2'], [0.222, '7718A6'], [0.333, '66158E'],
  [0.444, '561277'], [0.556, '450E60'], [0.667, '340B48'], [0.778, '230731'],
  [0.889, '13051B'], [1.000, '09040B'],
];

function rampColor(t) {
  for (let i = 1; i < BG_STOPS.length; i++) {
    if (t <= BG_STOPS[i][0] || i === BG_STOPS.length - 1) {
      const [p0, c0] = BG_STOPS[i - 1];
      const [p1, c1] = BG_STOPS[i];
      const f = (t - p0) / (p1 - p0);
      let out = '';
      for (let k = 0; k < 3; k++) {
        const a = parseInt(c0.substr(k * 2, 2), 16);
        const b = parseInt(c1.substr(k * 2, 2), 16);
        out += Math.round(a + (b - a) * f).toString(16).padStart(2, '0');
      }
      return out.toUpperCase();
    }
  }
}

function backgroundBands() {
  const rad = (BG_ANGLE * Math.PI) / 180;
  const uMax = SLIDE_W * Math.cos(rad) + SLIDE_H * Math.sin(rad);
  const bands = 56;
  const step = uMax / bands;
  const objects = [];
  for (let i = 1; i < bands; i++) {
    const u = i * step;           // this band covers the half plane beyond `u`
    const len = uMax * 2;
    objects.push({ rect: {
      x: (u + len / 2) * Math.cos(rad) - 17,
      y: (u + len / 2) * Math.sin(rad) - len / 2,
      w: 34,
      h: len,
      fill: { color: rampColor((i + 0.5) / bands) },
      rotate: BG_ANGLE - 90,
    } });
  }
  return objects;
}

/* =========================================================== text helpers */

// Default text options: Saira, white, top aligned, PowerPoint default insets.
const txt = (o) => Object.assign({
  fontFace: FONT, fontSize: 12, color: C.white, valign: 'top',
  margin: [7.2, 7.2, 3.6, 3.6],
}, o);

/** Two-tone headline: cyan lead words then white remainder. */
function heading(s, cyanPart, whitePart, o) {
  const runs = [{ text: cyanPart, options: { color: C.cyan } }];
  if (whitePart) runs.push({ text: whitePart, options: { color: C.white } });
  s.addText(runs, txt(Object.assign({ fontSize: 40, bold: true }, o)));
}

/** The recurring cyan -> green "double dash" rule (two 0.55in segments). */
function dashRule(s, x, y) {
  s.addShape('rect', { x, y, w: 0.551, h: 0.05, fill: { color: C.cyan, transparency: 55 } });
  s.addShape('rect', { x: x + 0.493, y, w: 0.551, h: 0.05, fill: { color: C.green, transparency: 55 } });
}

/* ======================================================= panel geometries */

// corner-cut presets shared by many plates
const WIDE_CUT = [0.868, 0, 0.868, 0];  // full-bleed content frame
const KPI_CUT  = [0.084, 0.2, 0.084, 0.2];  // 0.68in tall pill
const BTN_CUT  = [0, 0.062, 0, 0.062];      // 0.37in tall button

/** Rectangle with independently sized corner cuts [tl, tr, br, bl] (inches). */
function cutPts(w, h, cuts) {
  const [tl, tr, br, bl] = cuts;
  return [
    { x: tl, y: 0 }, { x: w - tr, y: 0 }, { x: w, y: tr }, { x: w, y: h - br },
    { x: w - br, y: h }, { x: bl, y: h }, { x: 0, y: h - bl }, { x: 0, y: tl },
    { close: true },
  ];
}

/**
 * Frosted "HUD" plate: translucent purple body with a bright cyan hairline.
 * `cuts` are the four diagonal corner snips in inches, clockwise from top-left.
 * opts: { alpha - fill opacity %, lw - hairline pt, fill, line }
 */
function panel(s, x, y, w, h, cuts, opts) {
  const o = opts || {};
  s.addShape('custGeom', {
    x, y, w, h,
    points: cutPts(w, h, cuts),
    fill: { color: o.fill || C.lilac, transparency: 100 - (o.alpha === undefined ? 20 : o.alpha) },
    line: { color: o.line || C.cyan, width: o.lw === undefined ? 1.5 : o.lw },
  });
}

/** Solid pill/button plate used for "More Here" and the KPI badges. */
function chip(s, x, y, w, h, cuts, o) {
  o = o || {};
  panel(s, x, y, w, h, cuts, { alpha: 100, fill: C.orchid, lw: o.lw === undefined ? 1.5 : o.lw });
  if (o.text) {
    s.addText(o.text, txt({ x, y, w, h, align: 'center', valign: 'middle',
      bold: true, color: C.cyan, fontSize: o.size || 16 }));
  }
}

/** Small chip that hosts a line-art icon (0.6 x 0.51 in throughout the deck). */
function iconChip(s, x, y) {
  panel(s, x, y, 0.601, 0.509, [0.085, 0, 0.085, 0], { alpha: 50 });
}

/** Placeholder standing in for a photograph in the reference deck. */
function imagePlate(s, x, y, w, h, cuts, o) {
  o = o || {};
  s.addShape('custGeom', {
    x, y, w, h,
    points: cutPts(w, h, cuts || [0, 0, 0, 0]),
    fill: { color: o.fill || C.lilac, transparency: o.transparency === undefined ? 88 : o.transparency },
    line: { color: o.line || '7FD7E8', width: 0.5, transparency: 45 },
    rotate: o.rotate, flipH: o.flipH,
  });
  s.addText(o.label || '[image]', txt({ x, y, w, h, align: 'center', valign: 'middle',
    fontSize: o.size || 11, color: 'C9A9E4' }));
}

/* ================================================================== icons */

/* Line-art glyphs rebuilt from native shapes (originals were freeform art). */

function iconUsb(s, x, y, w, h, col) {
  const c = col || C.mint, cx = x + w / 2;
  s.addShape('upArrow', { x: cx - w * 0.13, y, w: w * 0.26, h: h * 0.62, fill: { color: c } });
  s.addShape('rect', { x: cx - w * 0.05, y: y + h * 0.3, w: w * 0.1, h: h * 0.5, fill: { color: c } });
  s.addShape('ellipse', { x: cx - w * 0.15, y: y + h * 0.74, w: w * 0.3, h: h * 0.26, fill: { color: c } });
  s.addShape('line', { x: x + w * 0.2, y: y + h * 0.42, w: w * 0.3, h: h * 0.3, line: { color: c, width: 2 } });
  s.addShape('line', { x: cx, y: y + h * 0.42, w: w * 0.28, h: h * 0.3, line: { color: c, width: 2 }, flipH: true });
  s.addShape('ellipse', { x: x + w * 0.09, y: y + h * 0.3, w: w * 0.19, h: h * 0.17, fill: { color: c } });
  s.addShape('rect', { x: x + w * 0.7, y: y + h * 0.3, w: w * 0.2, h: h * 0.17, fill: { color: c } });
}

function iconCloud(s, x, y, w, h, col) {
  const c = col || C.mint;
  s.addShape('cloud', { x, y: y + h * 0.04, w, h: h * 0.66,
    fill: { type: 'none' }, line: { color: c, width: 2 } });
  s.addShape('downArrow', { x: x + w * 0.31, y: y + h * 0.24, w: w * 0.38, h: h * 0.7, fill: { color: c } });
}

function iconShare(s, x, y, w, h, col) {
  const c = col || C.mint;
  s.addShape('cloud', { x: x + w * 0.42, y, w: w * 0.58, h: h * 0.42, fill: { color: c } });
  s.addShape('rect', { x, y: y + h * 0.42, w: w * 0.6, h: h * 0.38,
    fill: { type: 'none' }, line: { color: c, width: 2 } });
  s.addShape('rect', { x: x + w * 0.22, y: y + h * 0.8, w: w * 0.16, h: h * 0.12, fill: { color: c } });
  s.addShape('bentUpArrow', { x: x + w * 0.16, y: y + h * 0.06, w: w * 0.26, h: h * 0.26, fill: { color: c } });
  s.addShape('bentUpArrow', { x: x + w * 0.62, y: y + h * 0.38, w: w * 0.26, h: h * 0.26,
    fill: { color: c }, rotate: 180 });
}

function iconLaptop(s, x, y, w, h, col) {
  const c = col || C.mint;
  s.addShape('rect', { x: x + w * 0.16, y, w: w * 0.68, h: h * 0.74,
    fill: { type: 'none' }, line: { color: c, width: 2.5 } });
  s.addShape('roundRect', { x, y: y + h * 0.78, w, h: h * 0.16, rectRadius: h * 0.08, fill: { color: c } });
  s.addShape('ellipse', { x: x + w * 0.36, y: y + h * 0.16, w: w * 0.28, h: h * 0.42,
    fill: { type: 'none' }, line: { color: c, width: 1.5 } });
  s.addShape('line', { x: x + w * 0.36, y: y + h * 0.37, w: w * 0.28, h: 0, line: { color: c, width: 1.5 } });
  s.addShape('line', { x: x + w * 0.5, y: y + h * 0.16, w: 0, h: h * 0.42, line: { color: c, width: 1.5 } });
}

function iconMoney(s, x, y, w, h, col) {
  const c = col || C.mint;
  s.addShape('rect', { x: x + w * 0.14, y, w: w * 0.72, h: h * 0.76,
    fill: { type: 'none' }, line: { color: c, width: 2.5 } });
  s.addShape('roundRect', { x, y: y + h * 0.8, w, h: h * 0.16, rectRadius: h * 0.08, fill: { color: c } });
  s.addShape('rect', { x: x + w * 0.28, y: y + h * 0.16, w: w * 0.44, h: h * 0.44,
    fill: { type: 'none' }, line: { color: c, width: 1.5 } });
  s.addShape('ellipse', { x: x + w * 0.42, y: y + h * 0.28, w: w * 0.16, h: h * 0.2, fill: { color: c } });
}

function iconWifi(s, x, y, w, h) {
  // three quarter-arcs plus a dot, drawn bottom-right anchored
  [0.95, 0.68, 0.42].forEach((k, i) => {
    s.addShape('arc', { x: x + w - w * k, y: y + h - h * k, w: w * k * 2, h: h * k * 2,
      angleRange: [180, 270], line: { color: [C.green, C.mint, C.cyan][i], width: 4 } });
  });
  s.addShape('ellipse', { x: x + w * 0.78, y: y + h * 0.78, w: w * 0.22, h: h * 0.22, fill: { color: C.cyan } });
}

function iconBars(s, x, y, w, h) {
  [[0.0, 0.28], [0.26, 0.52], [0.52, 0.76], [0.78, 1.0]].forEach(([bx, bh], i) => {
    s.addShape('rect', { x: x + w * bx, y: y + h * (1 - bh), w: w * 0.22, h: h * bh,
      fill: { color: [C.cyan, C.mint, C.mint, C.green][i] } });
  });
}

function iconApple(s, x, y, w, h) {
  s.addShape('ellipse', { x: x + w * 0.04, y: y + h * 0.24, w: w * 0.56, h: h * 0.76, fill: { color: C.mint } });
  s.addShape('ellipse', { x: x + w * 0.4, y: y + h * 0.24, w: w * 0.56, h: h * 0.76, fill: { color: C.mint } });
  s.addShape('rect', { x: x + w * 0.2, y: y + h * 0.3, w: w * 0.6, h: h * 0.4, fill: { color: C.mint } });
  s.addShape('teardrop', { x: x + w * 0.48, y, w: w * 0.3, h: h * 0.3, fill: { color: C.yellow }, rotate: 300 });
}

function iconPlay(s, x, y, w, h) {
  s.addShape('triangle', { x: x - w * 0.15, y: y + h * 0.15, w: h * 0.7, h: w * 1.3,
    fill: { color: C.green }, rotate: 90 });
  s.addShape('triangle', { x: x - w * 0.15, y: y + h * 0.5, w: h * 0.36, h: w * 1.3,
    fill: { color: C.cyan }, rotate: 90 });
}

/** Standing person pictogram used for the "Total User" figures. */
function iconPerson(s, x, y, w, h, skirt) {
  s.addShape('ellipse', { x: x + w * 0.28, y, w: w * 0.44, h: h * 0.28, fill: { color: C.mint } });
  s.addShape(skirt ? 'trapezoid' : 'roundRect',
    { x, y: y + h * 0.31, w, h: h * 0.42, rectRadius: w * 0.2, fill: { color: C.mint } });
  s.addShape('rect', { x: x + w * 0.2, y: y + h * 0.7, w: w * 0.22, h: h * 0.3, fill: { color: C.mint } });
  s.addShape('rect', { x: x + w * 0.58, y: y + h * 0.7, w: w * 0.22, h: h * 0.3, fill: { color: C.mint } });
}

/** Check-mark tick used by the pricing feature lists. */
function tickBox(s, x, y, size) {
  s.addShape('rect', { x, y, w: size, h: size, fill: { type: 'none' }, line: { color: C.white, width: 1.25 } });
  s.addText('\u2713', txt({ x: x - 0.02, y: y - 0.04, w: size + 0.04, h: size + 0.08,
    align: 'center', valign: 'middle', fontSize: 10, bold: true, color: C.white }));
}

/** Row of five rating stars, `filled` of them mint and the rest grey. */
function starRow(s, x, y, filled) {
  for (let i = 0; i < 5; i++) {
    s.addShape('star5', { x: x + i * 0.2465, y, w: 0.18, h: 0.18,
      fill: { color: i < filled ? C.mint : 'D8D8D8' } });
  }
}

/* ====================================================== composite widgets */

/** "1.4333 Million / SUBTITLE HERE / body" statistic block. */
function statBlock(s, x, y, o) {
  s.addText(o.value, txt({ x, y, w: o.w || 2.95, h: 0.44,
    fontSize: o.valueSize || 20, color: o.valueColor || C.cyan, bold: !!o.valueBold }));
  s.addText(o.sub, txt({ x, y: y + 0.415, w: o.w || 2.95, h: 0.27, fontSize: 10, bold: true }));
  if (o.body) {
    s.addText(o.body, txt({ x, y: y + 0.62, w: o.bodyW || 2.32, h: o.bodyH || 0.58,
      fontSize: 10, lineSpacingMultiple: 1.5 }));
  }
}

/** "2024 / 2500+ employee / A wonderful serenity..." year block. */
function yearBlock(s, x, y, o) {
  o = o || {};
  s.addText(o.year || '2024', txt({ x, y, w: 1.61, h: o.big ? 0.64 : 0.57,
    fontSize: o.big ? 32 : 28, bold: true }));
  s.addText('2500+ employee', txt({ x, y: y + (o.big ? 0.51 : 0.52), w: 1.99, h: o.big ? 0.4 : 0.38,
    fontSize: o.big ? 12 : 11, color: o.big ? C.cyan : C.mint, lineSpacingMultiple: 1.5 }));
  s.addText(o.body || 'A wonderful serenity has taken possession of my',
    txt({ x, y: y + (o.big ? 0.86 : 0.86), w: o.big ? 2.12 : 2.0, h: o.big ? 0.71 : 0.66,
      fontSize: o.big ? 12 : 11, lineSpacingMultiple: 1.5 }));
}

/** Label + percentage + rounded progress track. */
function progressRow(s, x, y, o) {
  s.addText(o.label, txt({ x, y, w: 1.71, h: 0.45, fontSize: 14, lineSpacingMultiple: 1.5 }));
  s.addText(o.pct + '%', txt({ x: x + 1.44, y, w: 0.99, h: 0.45, fontSize: 14, bold: true,
    align: 'right', color: o.pctColor || C.cyan, lineSpacingMultiple: 1.5 }));
  s.addShape('roundRect', { x: x + 0.42, y: y + 0.46, w: 2.04, h: 0.08, rectRadius: 0.04,
    fill: { color: C.ltGrey } });
  s.addShape('roundRect', { x: x + 0.09, y: y + 0.46, w: 2.37 * (o.pct / 100), h: 0.08,
    rectRadius: 0.04, fill: { color: C.mint } });
}

/** "Your text 01 ............ +153" dotted data row. */
function dataRow(s, x, y, w, label, value) {
  s.addText(label, txt({ x, y, w: w * 0.5, h: 0.36, lineSpacingMultiple: 1.3 }));
  s.addText(value, txt({ x: x + w * 0.5, y, w: w * 0.5, h: 0.36, align: 'right',
    bold: true, color: C.cyan, lineSpacingMultiple: 1.3 }));
  s.addShape('line', { x, y: y + 0.18, w, h: 0,
    line: { color: C.white, width: 0.75, dashType: 'dash', transparency: 51 } });
}

/** "SUBTITLE HERE" + one-line body, optionally with a leading icon chip. */
function subtitleItem(s, x, y, o) {
  s.addText('SUBTITLE HERE', txt({ x, y, w: 2.95, h: 0.36, fontSize: 11, bold: true,
    color: o.subColor || C.white, lineSpacingMultiple: 1.5 }));
  s.addText(o.body, txt({ x, y: y + 0.34, w: o.w || 4.73, h: o.h || 0.36, fontSize: 11,
    lineSpacingMultiple: 1.5 }));
}

/** Four concentric arcs + Q1 half-disc: the "3,250" radial gauge. */
function radialGauge(s, x, y) {
  // x,y = top-left of the 2.01 x 1.66 gauge cluster
  const rings = [
    { d: 0.95, off: 0.53, a: [0.51, 89.5],  col: C.cyan },
    { d: 1.18, off: 0.41, a: [0.82, 112.5], col: C.mint },
    { d: 1.42, off: 0.29, a: [1.06, 138.9], col: C.mint },
    { d: 1.66, off: 0.18, a: [0.64, 159.6], col: C.green },
  ];
  rings.forEach((r) => {
    s.addShape('arc', { x: x + r.off, y: y + (1.66 - r.d) + 0.15, w: r.d, h: r.d,
      angleRange: r.a, rotate: 180, line: { color: r.col, width: 5 } });
  });
  s.addShape('pie', { x: x + 0.65, y: y + 0.5, w: 0.71, h: 0.72,
    angleRange: [180, 360], fill: { color: C.mint } });
  s.addShape('line', { x, y: y + 0.86, w: 2.01, h: 0,
    line: { color: C.white, width: 1, transparency: 55 } });
  s.addText('3,250', txt({ x: x + 1.06, y: y - 0.44, w: 0.9, h: 0.34, fontSize: 14, bold: true }));
  s.addShape('line', { x: x + 1.04, y: y - 0.26, w: 0, h: 0.64, line: { color: C.cyan, width: 1 } });
  s.addText('Q1', txt({ x: x + 0.68, y: y + 0.59, w: 0.66, h: 0.3, align: 'center' }));
}

/** Two person pictograms with "+86,214 / Total User" captions. */
function userStats(s, x, y) {
  [0, 0.85].forEach((dy, i) => {
    iconPerson(s, x + 0.02, y + dy + 0.08, 0.21, 0.44, i === 1);
    s.addText('+86,214', txt({ x: x + 0.29, y: y + dy - 0.2, w: 1.66, h: 0.54,
      fontSize: 20, valign: 'bottom', lineSpacingMultiple: 1.3 }));
    s.addText('Total User', txt({ x: x + 0.29, y: y + dy + 0.27, w: 1.66, h: 0.36,
      lineSpacingMultiple: 1.3 }));
  });
}

/** Doughnut KPI: ring chart, big percentage, "View More" and body copy. */
function donutKpi(pres, s, x, y, o) {
  s.addChart(pres.ChartType.doughnut,
    [{ name: 'Sales', labels: ['rest', 'value'], values: [18, 82] }],
    { x, y, w: 1.71, h: 1.77, holeSize: 75, firstSliceAng: 0, showLegend: false,
      showValue: false, chartColors: [C.dimGrey, C.mint], dataBorder: { pt: 0, color: C.white } });
  s.addText(o.pct, txt({ x: x + 0.39, y: y + 0.57, w: 0.95, h: 0.63, align: 'center',
    fontSize: 24, bold: true, lineSpacingMultiple: 1.3 }));
  if (o.withText) {
    s.addText('View More', txt({ x: x + 1.74, y: y + 0.16, w: 1.57, h: 0.45,
      fontSize: 14, bold: true, color: C.cyan, lineSpacingMultiple: 1.5 }));
    s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit.',
      txt({ x: x + 1.74, y: y + 0.52, w: o.bodyW || 1.96, h: 1.01, lineSpacingMultiple: 1.5 }));
  }
}

/** Store badge: glyph + "Available on / Epp Store" + rounded action button. */
function storeRow(s, x, y) {
  iconApple(s, x + 0.2, y + 0.22, 0.22, 0.21);
  s.addText('Available on', txt({ x: x + 0.41, y, w: 1.47, h: 0.38, bold: true,
    color: C.cyan, lineSpacingMultiple: 1.5 }));
  s.addText('Epp Store', txt({ x: x + 0.41, y: y + 0.2, w: 1.47, h: 0.38, lineSpacingMultiple: 1.5 }));
  iconPlay(s, x + 2.09, y + 0.21, 0.22, 0.25);
  s.addText('Get it on', txt({ x: x + 2.31, y: y + 0.01, w: 0.88, h: 0.38, bold: true,
    color: C.cyan, lineSpacingMultiple: 1.5 }));
  s.addText('Pay Store', txt({ x: x + 2.31, y: y + 0.2, w: 1.08, h: 0.38, lineSpacingMultiple: 1.5 }));
  chip(s, x, y + 0.7, 1.475, 0.37, BTN_CUT, { text: 'More Here', size: 10.5, lw: 1 });
  chip(s, x + 1.881, y + 0.7, 1.475, 0.37, BTN_CUT, { text: 'Download', size: 10.5, lw: 1 });
}

/** "Total Data / 155M / Lorem ipsum." callout card. */
function totalDataCard(s, x, y, value, pointer) {
  panel(s, x, y, 1.569, 1.47, [0.245, 0, 0.245, 0]);
  s.addText('Total Data', txt({ x: x + 0.15, y: y + 0.11, w: 2.02, h: 0.34, fontSize: 14, bold: true }));
  s.addText(value, txt({ x: x + 0.15, y: y + 0.49, w: 2.02, h: 0.64, fontSize: 32, bold: true, color: C.cyan }));
  s.addText('Lorem ipsum.', txt({ x: x + 0.15, y: y + 1.07, w: 3.35, h: 0.33, fontSize: 11 }));
  if (pointer) {
    s.addShape('triangle', { x: x + 0.15, y: y + 1.47, w: 0.22, h: 0.34, rotate: 180,
      fill: { color: C.lilac, transparency: 80 }, line: { color: C.cyan, width: 1.5 } });
  }
}

/* ================================================================= slides */

const LOREM = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam. Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud ';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.';
const LOREM_TINY = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.';

/* --- 1 / 25: cover + closing share the same layout --------------------- */

function coverSlide(s, titleCyan, titleWhite, plateW, titleW) {
  panel(s, 1.249, 2.245, plateW, 4.505, [1.22, 0.36, 1.22, 0.36]);
  panel(s, 0.667, 0.75, 5.418, 3.79, [1.025, 0, 1.025, 0], { alpha: 80 });
  dashRule(s, 1.492, 1.489);
  s.addText([{ text: titleCyan, options: { color: C.cyan } },
             { text: titleWhite, options: { color: C.white } }],
    txt({ x: 1.304, y: 1.61, w: titleW || 4.676, h: 2.524, fontSize: 48, bold: true }));
  s.addText([{ text: 'www.website.com', options: { breakLine: true } },
             { text: '+123 6678 8890', options: { breakLine: true } },
             { text: 'Address Street, Address City, Address Country' }],
    txt({ x: 1.809, y: 5.396, w: 4.098, h: 0.911, fontSize: 11, lineSpacingMultiple: 1.5 }));
  s.addText('Template for your business',
    txt({ x: 7.144, y: 5.261, w: 4.941, h: 0.37, fontSize: 16, bold: true, color: C.cyan }));
  s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Aenean tincidunt imperdiet congue. Nullam vitae tempus sem.',
    txt({ x: 7.144, y: 5.664, w: 4.781, h: 0.633, fontSize: 11, lineSpacingMultiple: 1.5 }));
}

function slide01(s) { coverSlide(s, 'Futuristic ', 'Presentation Template', 11.418); }
function slide25(s) { coverSlide(s, 'Thank you ', 'for your attention', 10.731, 4.24); }

/* --- 2: intro with two stat cards -------------------------------------- */

function slide02(s) {
  panel(s, 0.667, 3.932, 6.0, 2.568, [0, 0.428, 0, 0.428]);
  panel(s, 0.667, 1.0, 5.786, 2.734, [0.456, 0, 0.456, 0]);
  imagePlate(s, 6.667, 1.0, 6.0, 2.568, [0, 0.428, 0, 0.428]);
  imagePlate(s, 6.88, 3.766, 5.786, 2.734, [0.456, 0, 0.456, 0]);
  heading(s, 'Your title ', 'here', { x: 1.359, y: 1.419, w: 4.403, h: 0.707, fontSize: 36 });
  s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam. Lorem ipsum dolor sit amet, consectetur adipiscing elit.',
    txt({ x: 1.359, y: 2.126, w: 4.403, h: 1.189, fontSize: 11, lineSpacingMultiple: 1.5 }));
  iconWifi(s, 1.22, 4.6, 0.35, 0.35);
  iconBars(s, 3.67, 4.57, 0.37, 0.37);
  yearBlock(s, 1.664, 4.513, { year: '2024' });
  yearBlock(s, 4.132, 4.513, { year: '2025' });
  dashRule(s, 5.452, 6.311);
}

/* --- 3: two column text ------------------------------------------------ */

function slide03(s) {
  imagePlate(s, 0.667, 0.75, 6.0, 6.0, [0, 1.0, 0, 1.0]);
  heading(s, 'Your title ', 'here', { x: 7.19, y: 1.021, w: 5.25, h: 0.769 });
  s.addText(LOREM, txt({ x: 7.207, y: 2.0, w: 5.234, h: 1.594, lineSpacingMultiple: 1.5 }));
  s.addText('Title Here', txt({ x: 7.207, y: 4.03, w: 2.07, h: 0.52, fontSize: 18, bold: true,
    color: C.cyan, lineSpacingMultiple: 1.5 }));
  s.addText(LOREM, txt({ x: 7.207, y: 4.99, w: 5.234, h: 1.594, lineSpacingMultiple: 1.5 }));
  dashRule(s, 8.428, 4.317);
}

/* --- 4: "Design Thinking" with two side rails -------------------------- */

function slide04(s) {
  panel(s, 0.667, 0.75, 3.0, 6.0, [0.5, 0, 0.5, 0]);
  panel(s, 9.667, 0.75, 3.0, 6.0, [0, 0.5, 0, 0.5]);
  imagePlate(s, 4.104, 0.75, 5.125, 3.0, [0.5, 0, 0.5, 0]);

  heading(s, 'Your title ', 'here', { x: 0.933, y: 1.4, w: 2.62, h: 0.5, fontSize: 24 });
  s.addText('Lorem ipsum dolor sit amet.', txt({ x: 0.933, y: 1.84, w: 2.45, h: 0.38,
    lineSpacingMultiple: 1.5 }));
  iconLaptop(s, 1.02, 2.51, 0.96, 0.58);
  s.addText('Title Here', txt({ x: 2.04, y: 2.39, w: 1.39, h: 0.42, fontSize: 14, bold: true,
    lineSpacingMultiple: 1.5 }));
  s.addText('Lorem Ipsum', txt({ x: 2.04, y: 2.73, w: 1.39, h: 0.42, fontSize: 14,
    lineSpacingMultiple: 1.5 }));
  s.addText(LOREM_TINY, txt({ x: 0.933, y: 3.34, w: 2.92, h: 0.71, bold: true,
    lineSpacingMultiple: 1.5 }));
  ['01', '02', '03', '04'].forEach((n, i) => {
    dataRow(s, 0.933, 4.16 + i * 0.563, 2.47, 'Your text ' + n, ['+153', '+345', '+111', '+336'][i]);
  });

  heading(s, 'Design Thinking', '', { x: 4.62, y: 4.07, w: 4.09, h: 0.71, fontSize: 36 });
  s.addText(LOREM, txt({ x: 4.104, y: 4.84, w: 5.12, h: 1.594, lineSpacingMultiple: 1.5 }));

  statBlock(s, 10.04, 1.39, { value: '2024', valueColor: C.white,
    sub: 'SUBTITLE HERE', body: LOREM_SHORT, bodyH: 1.09 });
  statBlock(s, 10.04, 3.33, { value: '1.4333 Million', valueColor: C.white,
    sub: 'SUBTITLE HERE', body: LOREM_TINY });
  statBlock(s, 10.04, 4.78, { value: '1.4333 Million', valueColor: C.white,
    sub: 'SUBTITLE HERE', body: LOREM_TINY, bodyW: 2.41 });
  dashRule(s, 11.466, 6.531);
}

/* --- 5: radial gauge + progress list ----------------------------------- */

function slide05(s) {
  panel(s, 0.667, 1.021, 3.167, 5.556, [0.528, 0, 0.528, 0]);
  imagePlate(s, 9.859, 1.02, 4.161, 5.56, [0, 0.694, 0, 0.694]);

  heading(s, 'Your title ', 'here', { x: 4.25, y: 1.021, w: 5.25, h: 0.769 });
  s.addText(LOREM, txt({ x: 4.267, y: 2.0, w: 5.234, h: 1.594, lineSpacingMultiple: 1.5 }));
  s.addText(LOREM, txt({ x: 4.267, y: 4.99, w: 5.234, h: 1.594, lineSpacingMultiple: 1.5 }));
  statBlock(s, 4.267, 3.951, { value: '1.4333 Million', sub: 'SUBTITLE HERE' });
  dashRule(s, 6.281, 4.129);

  radialGauge(s, 1.19, 2.15);
  s.addText([{ text: 'A wonderful serenity', options: { breakLine: true } },
             { text: 'has taken text' }],
    txt({ x: 1.19, y: 3.17, w: 2.12, h: 0.71, lineSpacingMultiple: 1.5,
      bullet: { characterCode: '2714', indent: 13.5 } }));

  s.addText('Data and Progress', txt({ x: 0.92, y: 4.41, w: 1.99, h: 0.41, fontSize: 14,
    bold: true, lineSpacingMultiple: 1.3 }));
  ['22', '44', '66'].forEach((pct, i) => {
    const y = 4.87 + i * 0.33;
    s.addText('Your Data 0' + (i + 1), txt({ x: 0.94, y, w: 1.56, h: 0.36, lineSpacingMultiple: 1.3 }));
    s.addText(pct + '%', txt({ x: 2.7, y, w: 0.88, h: 0.36, align: 'right', bold: true,
      color: C.cyan, lineSpacingMultiple: 1.3 }));
    s.addShape('line', { x: 2.03, y: y + 0.18, w: 0.97, h: 0,
      line: { color: 'A5A5A5', width: 0.75, dashType: 'dash', transparency: 51 } });
  });
}

/* --- 6: three stat columns + SEO donut --------------------------------- */

function slide06(s) {
  heading(s, 'Your title ', 'here', { x: 4.042, y: 0.771, w: 5.25, h: 0.769, align: 'center' });
  s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam. Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad',
    txt({ x: 1.046, y: 1.542, w: 11.24, h: 0.68, align: 'center', lineSpacingMultiple: 1.5 }));

  imagePlate(s, 1.048, 2.62, 3.0, 2.06, [0.343, 0, 0.343, 0]);
  imagePlate(s, 5.167, 2.62, 3.0, 2.06, [0.343, 0, 0.343, 0]);
  panel(s, 9.286, 2.628, 3.006, 2.052, [0.342, 0, 0.342, 0]);

  [1.046, 5.166, 9.287].forEach((x) => {
    statBlock(s, x, 4.935, { value: '1.4333 Million', sub: 'SUBTITLE HERE',
      body: LOREM_SHORT, bodyW: 2.99, bodyH: 0.84 });
  });
  dashRule(s, 5.336, 4.474);

  // SEO donut gauge
  s.addText('SEO', txt({ x: 9.45, y: 2.8, w: 1.75, h: 0.34, valign: 'bottom', bold: true, color: C.cyan }));
  s.addText('Lorem ipsum', txt({ x: 9.45, y: 3.03, w: 0.97, h: 0.27, fontSize: 8 }));
  s.addText('12 Sept 2024', txt({ x: 9.45, y: 4.18, w: 0.82, h: 0.4, fontSize: 8 }));
  s.addText('(34%)', txt({ x: 10.9, y: 2.8, w: 1.31, h: 0.32, fontSize: 11, bold: true,
    align: 'right', color: C.cyan, valign: 'middle' }));
  s.addShape('arc', { x: 10.23, y: 3.18, w: 1.12, h: 1.1, angleRange: [127.6, 60.0],
    line: { color: 'D0CECE', width: 10, transparency: 75 } });
    s.addShape('ellipse', { x: 10.38, y: 3.33, w: 0.81, h: 0.8,
    fill: { color: C.lilac, transparency: 80 }, line: { color: C.cyan, width: 1.5 } });
  s.addText('2,678', txt({ x: 10.33, y: 3.58, w: 0.92, h: 0.29, align: 'center', valign: 'middle',
    fontSize: 11, bold: true, color: C.cyan }));
  s.addShape('arc', { x: 10.23, y: 3.18, w: 1.12, h: 1.1, angleRange: [127.2, 294.3],
    line: { color: C.mint, width: 10 } });
  s.addShape('line', { x: 11.02, y: 2.96, w: 0.54, h: 0.27, line: { color: C.mint, width: 1 } });
  s.addShape('ellipse', { x: 11.54, y: 2.94, w: 0.05, h: 0.05, fill: { color: C.mint } });
  s.addText('Your text here', txt({ x: 11.38, y: 4.18, w: 0.82, h: 0.4, fontSize: 8 }));
  s.addShape('line', { x: 11.35, y: 4.14, w: 0, h: 0.16, line: { color: C.white, width: 0.75 } });
}

/* --- 7: photo left, icon list right ------------------------------------ */

function slide07(s) {
  imagePlate(s, 0.667, 0.75, 6.0, 3.0, [0.5, 0, 0.5, 0]);
  panel(s, 6.667, 3.75, 6.0, 3.0, [0.5, 0, 0.5, 0]);
  heading(s, 'Your title ', 'here', { x: 7.052, y: 0.771, w: 5.47, h: 0.769 });
  s.addText(LOREM_SHORT, txt({ x: 7.052, y: 1.7, w: 5.47, h: 0.68, lineSpacingMultiple: 1.5 }));
  s.addText(LOREM_SHORT, txt({ x: 7.052, y: 2.55, w: 5.47, h: 0.68, lineSpacingMultiple: 1.5 }));

  const bodies = [
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do',
    'Lorem ipsum dolor sit amet, consectetur adipiscing eli. Sed do',
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do',
  ];
  const icons = [iconUsb, iconCloud, iconShare];
  bodies.forEach((body, i) => {
    const y = 4.21 + i * 0.803;
    iconChip(s, 7.023, y);
    icons[i](s, 7.15, y + 0.08, 0.35, 0.35);
    subtitleItem(s, 7.79, y - 0.1, { body });
  });

  statBlock(s, 0.68, 4.117, { value: '1.4333 Million', sub: 'SUBTITLE HERE' });
  dashRule(s, 2.664, 4.297);
  s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam. Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad',
    txt({ x: 0.68, y: 4.82, w: 4.54, h: 1.89, lineSpacingMultiple: 1.5 }));
}

/* --- 8: pie chart + parameters ----------------------------------------- */

function slide08(s) {
  panel(s, 0.667, 0.75, 12.0, 3.0, [0.5, 0, 0.5, 0]);
  heading(s, 'Your title ', 'here', { x: 4.62, y: 1.14, w: 4.09, h: 0.71, fontSize: 36, align: 'center' });
  s.addText('Subtitle here', txt({ x: 5.36, y: 1.88, w: 2.62, h: 0.44, fontSize: 20, bold: true,
    color: C.cyan, align: 'center' }));
  s.addText(LOREM, txt({ x: 1.06, y: 2.35, w: 11.22, h: 0.98, align: 'center',
    lineSpacingMultiple: 1.5 }));
  s.addShape('rect', { x: 5.14, y: 2.07, w: 0.551, h: 0.05, fill: { color: C.cyan, transparency: 55 } });
  s.addShape('rect', { x: 7.65, y: 2.08, w: 0.551, h: 0.05, fill: { color: C.green, transparency: 55 } });

  imagePlate(s, 0.667, 3.918, 6.0, 2.832, [0.472, 0, 0.472, 0]);
  panel(s, 6.828, 3.918, 5.839, 2.832, [0, 0.472, 0, 0.472]);
  s.addText('Paid Ads', txt({ x: 6.96, y: 4.05, w: 1.75, h: 0.34, valign: 'bottom', bold: true, color: C.cyan }));
  s.addText('12 Sept 2024', txt({ x: 6.96, y: 4.31, w: 1.75, h: 0.31, fontSize: 10 }));

  // pie: four wedges, angles taken from the reference
  s.addShape('ellipse', { x: 7.55, y: 4.64, w: 1.74, h: 1.74,
    fill: { color: C.mint }, line: { color: C.white, width: 1 } });
  [[101.8, 270.0, C.plum], [227.2, 298.9, C.purple], [294.8, 102.4, C.lilac],
   [0.4, 102.4, 'DBABF3']].forEach(([a, b, col]) => {
    s.addShape('pie', { x: 7.58, y: 4.68, w: 1.67, h: 1.67, angleRange: [a, b], fill: { color: col } });
  });
  [['35%', 7.72, 5.52], ['20%', 8.0, 4.91], ['10%', 8.51, 5.12], ['35%', 8.39, 5.73]]
    .forEach(([t, x, y]) => s.addText(t, txt({ x, y, w: 0.62, h: 0.31, align: 'center',
      valign: 'middle', fontSize: 11 })));

  ['80', '40', '67'].forEach((pct, i) => {
    progressRow(s, 9.59, 4.33 + i * 0.727, { label: 'Parameter ' + (i + 1), pct: +pct });
  });
}

/* --- 9: KPI card with up/down arrows ----------------------------------- */

function slide09(s) {
  panel(s, 0.667, 0.75, 6.0, 2.812, [0.345, 0.826, 0.345, 0.826]);
  imagePlate(s, 0.667, 3.75, 12.0, 3.0, [0.833, 0.271, 0.833, 0.271]);

  s.addText('Your Title Here', txt({ x: 1.26, y: 0.99, w: 1.73, h: 0.39, fontSize: 13.5,
    bold: true, align: 'center', lineSpacingMultiple: 1.15 }));
  [['79%', 1.56], ['37%', 2.35]].forEach(([label, y], i) => {
    chip(s, 1.258, y, 1.85, 0.682, KPI_CUT);
    s.addText(label, txt({ x: 1.1, y: y + 0.05, w: 1.73, h: 0.6, fontSize: 24, bold: true,
      align: 'center', lineSpacingMultiple: 1.15 }));
    s.addShape('triangle', { x: 2.63, y: y + 0.2, w: 0.24, h: 0.28,
      rotate: i === 0 ? 0 : 180, fill: { color: C.mint } });
  });

  s.addShape('triangle', { x: 3.66, y: 1.2, w: 0.18, h: 0.15, rotate: 90, fill: { color: C.mint } });
  s.addText('$12,809', txt({ x: 3.88, y: 1.02, w: 2.29, h: 0.6, fontSize: 24, bold: true,
    lineSpacingMultiple: 1.15 }));
  s.addText('Last Transaction', txt({ x: 3.91, y: 1.53, w: 1.82, h: 0.39, fontSize: 13.5,
    lineSpacingMultiple: 1.15 }));
  s.addText('USD', txt({ x: 5.64, y: 1.3, w: 0.69, h: 0.31, fontSize: 9, lineSpacingMultiple: 1.15 }));
  statBlock(s, 3.88, 2.18, { value: '9.457 Million', sub: 'SUBTITLE HERE',
    body: '556 Value Here', w: 2.48, bodyW: 2.48, bodyH: 0.33 });
  chip(s, 2.329, 3.315, 2.675, 0.682, KPI_CUT, { text: 'More Here', size: 18 });

  heading(s, 'Your title ', 'here', { x: 7.052, y: 0.914, w: 5.47, h: 0.769 });
  s.addText(LOREM_SHORT, txt({ x: 7.052, y: 1.85, w: 5.47, h: 0.68, lineSpacingMultiple: 1.5 }));
  s.addText(LOREM_SHORT, txt({ x: 7.052, y: 2.69, w: 5.47, h: 0.68, lineSpacingMultiple: 1.5 }));
  dashRule(s, 11.05, 1.28);
}

/* --- 10: single full-width quote panel --------------------------------- */

function slide10(s) {
  panel(s, 1.368, 2.017, 10.598, 3.466, [0.578, 0.293, 0.578, 0.293]);
  chip(s, 2.226, 1.676, 0.981, 0.682, [0.2, 0.075, 0.2, 0.075]);
  iconShare(s, 2.5, 1.79, 0.44, 0.45);
  heading(s, 'Your title ', 'here', { x: 3.93, y: 2.31, w: 5.47, h: 0.71, fontSize: 36, align: 'center' });
  s.addText(LOREM + 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam. Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud ',
    txt({ x: 1.81, y: 2.96, w: 9.72, h: 2.2, align: 'center', lineSpacingMultiple: 1.5 }));
  dashRule(s, 6.12, 5.109);
  chip(s, 8.381, 5.142, 2.675, 0.682, KPI_CUT, { text: 'More Here' });
}

/* --- 11: four price cards ---------------------------------------------- */

const PRICE_TIERS = [
  { name: 'Bronze',   price: '$105', icon: iconShare,  faded: false },
  { name: 'Silver',   price: '$205', icon: iconCloud,  faded: false },
  { name: 'Gold',     price: '$305', icon: iconUsb,    faded: true  },
  { name: 'Platinum', price: '$405', icon: iconMoney,  faded: false },
];

function slide11(s) {
  heading(s, 'Your title ', 'here', { x: 3.93, y: 0.76, w: 5.47, h: 0.769, align: 'center' });
  s.addShape('rect', { x: 4.08, y: 1.12, w: 0.551, h: 0.05, fill: { color: C.cyan, transparency: 55 } });
  s.addShape('rect', { x: 8.7, y: 1.12, w: 0.551, h: 0.05, fill: { color: C.green, transparency: 55 } });

  PRICE_TIERS.forEach((tier, i) => {
    const x = 0.667 + i * 3.113;
    panel(s, x, 2.222, 2.661, 3.621, [0.443, 0, 0.443, 0], { alpha: tier.faded ? 80 : 20 });
    chip(s, x + 1.78, 1.882, 0.881, 0.682, [0.2, 0, 0.2, 0]);
    tier.icon(s, x + 2.0, 1.99, 0.44, 0.45);
    s.addText(tier.name, txt({ x: x + 0.58, y: 2.67, w: 1.49, h: 0.44, fontSize: 20, align: 'center' }));
    s.addText(tier.price, txt({ x: x + 0.34, y: 3.2, w: 1.97, h: 0.77, fontSize: 44,
      color: C.cyan, align: 'center', valign: 'middle', lineSpacingMultiple: 0.9 }));
    s.addText('Monthly', txt({ x: x + 0.58, y: 3.9, w: 1.49, h: 0.45, fontSize: 16,
      align: 'center', lineSpacingMultiple: 1.3 }));
    tickBox(s, x + 0.32, 4.53, 0.23);
    s.addText('Lorem ipsum dolor sit amet, consectetur', txt({ x: x + 0.6, y: 4.47, w: 1.81, h: 0.63,
      lineSpacingMultiple: 1.3 }));
    tickBox(s, x + 0.32, 5.2, 0.23);
    s.addText('Adipiscing elit labore', txt({ x: x + 0.6, y: 5.14, w: 2.24, h: 0.36,
      lineSpacingMultiple: 1.3 }));
  });
  s.addText(LOREM_SHORT, txt({ x: 1.55, y: 6.15, w: 10.24, h: 0.38, align: 'center',
    lineSpacingMultiple: 1.5 }));
}

/* --- 12: profile card --------------------------------------------------- */

function slide12(s) {
  imagePlate(s, 6.069, 0.75, 7.851, 6.0, [0.448, 1.762, 0.448, 1.762]);
  panel(s, 1.208, 1.594, 4.308, 1.014, [0, 0.274, 0, 0.274], { alpha: 80 });
  panel(s, 0.667, 2.728, 6.979, 3.009, [0.814, 0.238, 0.814, 0.238]);
  heading(s, 'Your title ', 'here', { x: 1.42, y: 1.71, w: 4.09, h: 0.769 });
  s.addText('Franky Cyborg', txt({ x: 2.99, y: 2.97, w: 2.34, h: 0.44, fontSize: 20,
    color: C.cyan, align: 'center' }));
  s.addShape('rect', { x: 2.44, y: 3.16, w: 0.551, h: 0.05, fill: { color: C.cyan, transparency: 55 } });
  s.addShape('rect', { x: 5.28, y: 3.16, w: 0.551, h: 0.05, fill: { color: C.green, transparency: 55 } });

  ['175cm', '65kg', '25th', '65cm'].forEach((v, i) => {
    const x = 1.14 + i * 1.58;
    s.addShape('rect', { x, y: 3.68, w: 0.12, h: 0.12, fill: { color: C.mint } });
    s.addText(v, txt({ x: x + 0.12, y: 3.58, w: 1.69, h: 0.4, fontSize: 18, bold: true }));
    s.addText('Your Data Here', txt({ x: x + 0.12, y: 3.85, w: 1.88, h: 0.34, fontSize: 10.5,
      bold: true, lineSpacingMultiple: 1.3 }));
  });

  progressRow(s, 1.13, 4.37, { label: 'Your Value', pct: 67, pctColor: C.white });
  s.addText('Lorem ipsum dolor sit amet.', txt({ x: 1.13, y: 4.98, w: 2.43, h: 0.4,
    lineSpacingMultiple: 1.5 }));
  starRow(s, 4.08, 4.57, 4);
  s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.',
    txt({ x: 3.99, y: 4.77, w: 3.29, h: 0.63, lineSpacingMultiple: 1.3 }));
}

/* --- 13: stacked bar chart + store badges ------------------------------ */

function slide13(pres, s) {
  panel(s, 6.667, 1.758, 7.375, 3.985, [0, 0.664, 0, 0.664]);
  heading(s, 'Your title ', 'here', { x: 0.667, y: 1.757, w: 5.25, h: 0.769 });
  s.addText(LOREM, txt({ x: 0.685, y: 2.735, w: 5.234, h: 1.594, lineSpacingMultiple: 1.5 }));
  s.addText('Title Here', txt({ x: 7.362, y: 1.959, w: 2.07, h: 0.425, fontSize: 14,
    bold: true, color: C.cyan, lineSpacingMultiple: 1.5 }));

  const cats = ['Point', 'Point', 'Point', 'Point'];
  s.addChart(pres.ChartType.bar, [
    { name: 'Series 1', labels: cats, values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels: cats, values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels: cats, values: [2.0, 2.0, 3.0, 5.0] },
  ], { x: 7.301, y: 2.062, w: 3.027, h: 2.711, barDir: 'col', barGrouping: 'stacked',
       barGapWidthPct: 150, chartColors: [C.cyan, C.green, C.grey],
       showLegend: false, showValue: false, valAxisHidden: true,
       catAxisLabelColor: C.white, catAxisLabelFontFace: FONT, catAxisLabelFontSize: 10,
       catAxisLineColor: C.white, valGridLine: { style: 'none' },
       plotArea: { fill: { type: 'none' } } });
  s.addShape('line', { x: 7.46, y: 2.51, w: 0, h: 1.85, line: { color: C.cyan, width: 0.75 } });

  ['', '', ''].forEach((_, i) => {
    s.addShape('rect', { x: 10.48, y: 2.8 + i * 0.657, w: 0.26, h: 0.26,
      fill: { color: [C.cyan, C.green, C.grey][i] } });
    s.addText('Lorem ipsum dolor sit amet.', txt({ x: 10.8, y: 2.67 + i * 0.62, w: 1.85, h: 0.58,
      fontSize: 10, lineSpacingMultiple: 1.5 }));
  });
  dashRule(s, 10.48, 2.47);

  s.addText('2024 Statistic', txt({ x: 7.42, y: 4.79, w: 1.77, h: 0.41, fontSize: 14,
    color: C.cyan, lineSpacingMultiple: 1.3 }));
  s.addText('12m/Month', txt({ x: 7.42, y: 5.11, w: 1.77, h: 0.32, fontSize: 10,
    lineSpacingMultiple: 1.3 }));
  s.addText('+$33,999', txt({ x: 11.15, y: 4.7, w: 1.51, h: 0.49, fontSize: 18, bold: true,
    align: 'right', valign: 'bottom', color: C.cyan, lineSpacingMultiple: 1.3 }));
  s.addText('Business dashboard report', txt({ x: 10.33, y: 5.13, w: 2.34, h: 0.36,
    align: 'right', lineSpacingMultiple: 1.3 }));
  storeRow(s, 0.79, 4.67);
}

/* --- 14: team cards ----------------------------------------------------- */

const TEAM = [
  { name: 'Myle Cyrez',   x: 0.667 },
  { name: 'Jorge Brodie', x: 4.857 },
  { name: 'Catrine Mola', x: 9.057 },
];

function slide14(s) {
  heading(s, 'Meet our best team', '', { x: 3.88, y: 0.771, w: 5.57, h: 0.769, align: 'center' });
  TEAM.forEach((m, i) => {
    imagePlate(s, m.x, 2.091, 3.604, 3.985, [0, 0.601, 0, 0.601]);
    panel(s, m.x + 0.333, 4.153, 3.024, 1.553, [0, 0.259, 0, 0.259]);
    s.addText(m.name, txt({ x: m.x + 0.52, y: 4.34, w: 2.34, h: 0.37, fontSize: 16, color: C.cyan }));
    starRow(s, m.x + 0.61, 4.75, i === 2 ? 4 : 5);
    s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed',
      txt({ x: m.x + 0.52, y: 4.95, w: 2.7, h: 0.63, lineSpacingMultiple: 1.3 }));
    chip(s, m.x + 1.361, 5.891, 1.475, 0.37, BTN_CUT, { text: 'More Here', size: 10.5, lw: 1 });
  });
  dashRule(s, 2.55, 4.5);
}

/* --- 15: numbered list + user panel ------------------------------------ */

function slide15(s) {
  imagePlate(s, 6.667, 0.75, 6.0, 6.0, [0, 1.0, 0, 1.0]);
  heading(s, 'Your title ', 'here', { x: 0.82, y: 1.67, w: 5.25, h: 0.769 });
  ['01', '02', '03'].forEach((n, i) => {
    const y = 2.59 + i * 1.275;
    s.addText(n, txt({ x: 0.82, y: y + 0.04, w: 0.47, h: 0.34, fontSize: 14, bold: true }));
    s.addShape('rect', { x: 1.29, y: y + 0.18, w: 0.551, h: 0.05,
      fill: { color: i === 1 ? C.green : C.cyan, transparency: 55 } });
    s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et.',
      txt({ x: 1.8, y, w: 3.4, h: 0.98, lineSpacingMultiple: 1.5 }));
  });
  panel(s, 5.368, 2.626, 2.598, 3.121, [0.703, 0.206, 0.703, 0.206], { alpha: 80 });
  s.addText('Your text here', txt({ x: 5.5, y: 3.11, w: 2.34, h: 0.37, fontSize: 16,
    color: C.cyan, align: 'center' }));
  userStats(s, 5.96, 3.66);
}

/* --- 16: map dashboard --------------------------------------------------- */

function slide16(s) {
  panel(s, 1.368, 0.728, 10.598, 3.466, [0.578, 0, 0.578, 0]);
  imagePlate(s, 7.93, 1.14, 3.64, 2.68, [0, 0, 0, 0], { label: '[map image]' });

  radialGauge(s, 1.77, 1.8);
  s.addText([{ text: 'A wonderful serenity', options: { breakLine: true } },
             { text: 'has taken text', options: { breakLine: true } },
             { text: 'Lorem Ipsum' }],
    txt({ x: 1.77, y: 2.81, w: 2.12, h: 1.01, lineSpacingMultiple: 1.5,
      bullet: { characterCode: '2714', indent: 13.5 } }));

  [C.mint, C.cyan, C.green, C.grey].forEach((col, i) => {
    s.addShape('rect', { x: 4.28, y: 1.47 + i * 0.634, w: 0.26, h: 0.26, fill: { color: col } });
    s.addText('Lorem ipsum dolor sit amet.', txt({ x: 4.59, y: 1.34 + i * 0.61, w: 1.85, h: 0.58,
      fontSize: 10, lineSpacingMultiple: 1.5 }));
  });

  userStats(s, 6.67, 1.29);
  s.addShape('bentConnector3', { x: 8.22, y: 1.56, w: 2.29, h: 1.07, line: { color: C.mint, width: 1 } });
  s.addShape('bentConnector3', { x: 8.22, y: 2.48, w: 0.87, h: 0.84, line: { color: C.mint, width: 1 } });
  s.addShape('irregularSeal1', { x: 9.07, y: 3.19, w: 0.15, h: 0.15, fill: { color: C.mint } });
  s.addShape('irregularSeal1', { x: 10.48, y: 2.48, w: 0.15, h: 0.15, fill: { color: C.mint } });

  heading(s, 'Your title here', '', { x: 3.88, y: 4.6, w: 5.57, h: 0.769, align: 'center' });
  s.addText(LOREM, txt({ x: 1.367, y: 5.37, w: 10.6, h: 0.98, align: 'center',
    lineSpacingMultiple: 1.5 }));
}

/* --- 17: income pie + overlapping panels ------------------------------- */

function slide17(s) {
  heading(s, 'Your title ', 'here', { x: 0.667, y: 1.93, w: 5.2, h: 0.769 });
  s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.\u00a0 consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua',
    txt({ x: 0.667, y: 2.7, w: 6.01, h: 1.594, lineSpacingMultiple: 1.5 }));
  yearBlock(s, 0.667, 4.4, { big: true, body: 'A wonderful serenity has taken possession' });

  s.addText('2.8k', txt({ x: 2.79, y: 5.02, w: 1.39, h: 0.4, fontSize: 18, bold: true, align: 'center' }));
  s.addText([{ text: 'Projects', options: { breakLine: true } }, { text: 'total' }],
    txt({ x: 2.79, y: 5.33, w: 1.39, h: 0.5, align: 'center' }));
  s.addText('Your text here', txt({ x: 4.18, y: 4.89, w: 2.56, h: 0.45, fontSize: 14, bold: true,
    color: C.cyan, lineSpacingMultiple: 1.5 }));
  s.addText(LOREM_TINY, txt({ x: 4.18, y: 5.26, w: 2.56, h: 0.71, lineSpacingMultiple: 1.5 }));

  panel(s, 7.257, 2.176, 5.418, 3.79, [1.025, 0.3, 1.025, 0.3]);
  panel(s, 6.675, 1.458, 5.418, 3.79, [1.025, 0.3, 1.025, 0.3], { alpha: 80 });

  s.addText('43%', txt({ x: 7.49, y: 1.84, w: 1.42, h: 0.71, fontSize: 36, bold: true }));
  s.addText('Income', txt({ x: 7.49, y: 2.42, w: 1.42, h: 0.37, fontSize: 16 }));
  s.addShape('flowChartConnector', { x: 7.45, y: 3.19, w: 1.22, h: 1.22,
    fill: { color: 'A5A5A5', transparency: 80 } });
  s.addShape('pie', { x: 7.45, y: 3.19, w: 1.22, h: 1.22, angleRange: [269.4, 170.8],
    fill: { color: C.lilac } });
  s.addShape('arc', { x: 7.36, y: 3.11, w: 1.39, h: 1.39, angleRange: [270, 168.2],
    line: { color: C.mint, width: 2.5 } });
  s.addText('Your Text \u2013 22,70%', txt({ x: 6.95, y: 4.61, w: 2.22, h: 0.4, align: 'center',
    lineSpacingMultiple: 1.5 }));

  s.addText('23/06/2024', txt({ x: 9.23, y: 1.93, w: 4.0, h: 0.4, fontSize: 18, bold: true, color: C.cyan }));
  s.addText(LOREM_TINY, txt({ x: 9.23, y: 2.3, w: 2.72, h: 0.81, fontSize: 14, lineSpacingMultiple: 1.5 }));
  s.addText('Your text here', txt({ x: 9.23, y: 3.18, w: 2.72, h: 0.45, fontSize: 14, bold: true,
    color: C.cyan, lineSpacingMultiple: 1.5 }));
  s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et',
    txt({ x: 9.23, y: 3.63, w: 2.43, h: 1.28, lineSpacingMultiple: 1.5 }));
}

/* --- 18: two donut KPIs over a laptop shot ----------------------------- */

function slide18(pres, s) {
  panel(s, 1.042, 0.903, 8.46, 2.641, [0.952, 0, 0.952, 0], { alpha: 45 });
  imagePlate(s, 4.75, 0.92, 8.57, 5.57, [0.5, 0.5, 0.5, 0.5], { label: '[laptop image]' });
  donutKpi(pres, s, 1.538, 1.446, { pct: '80%', withText: true, bodyW: 1.96 });
  donutKpi(pres, s, 5.254, 1.446, { pct: '85%', withText: true, bodyW: 1.84 });

  heading(s, 'Your title ', 'here', { x: 1.06, y: 4.01, w: 4.69, h: 0.769 });
  s.addText(LOREM_SHORT, txt({ x: 1.06, y: 4.92, w: 5.77, h: 0.68, lineSpacingMultiple: 1.5 }));
  s.addText('2024 Statistic', txt({ x: 1.05, y: 5.85, w: 1.77, h: 0.41, fontSize: 14,
    color: C.cyan, lineSpacingMultiple: 1.3 }));
  s.addText('12m/Month', txt({ x: 1.05, y: 6.17, w: 1.77, h: 0.32, fontSize: 10,
    lineSpacingMultiple: 1.3 }));
  s.addText('+$33,999', txt({ x: 2.64, y: 5.78, w: 2.03, h: 0.71, fontSize: 28, bold: true,
    valign: 'bottom', color: C.cyan, lineSpacingMultiple: 1.3 }));
  s.addText('Business dashboard report', txt({ x: 4.84, y: 5.82, w: 2.23, h: 0.63,
    lineSpacingMultiple: 1.3 }));
}

/* --- 19: phone mockup ---------------------------------------------------- */

function slide19(s) {
  panel(s, 1.31, 1.109, 8.776, 5.253, [1.093, 0, 1.093, 0], { alpha: 45 });
  imagePlate(s, 9.584, 1.028, 2.525, 5.443, [0.24, 0.24, 0.24, 0.24], { label: '[phone image]' });
  heading(s, 'Your title ', 'here', { x: 1.98, y: 1.81, w: 5.25, h: 0.769 });
  s.addText(LOREM, txt({ x: 1.98, y: 2.79, w: 7.15, h: 1.29, lineSpacingMultiple: 1.5 }));
  storeRow(s, 1.98, 4.73);
  yearBlock(s, 5.65, 4.4, { big: true, body: 'A wonderful serenity has taken possession' });
  userStats(s, 7.69, 4.24);
}

/* --- 20: labelled bar chart + icon list -------------------------------- */

function slide20(pres, s) {
  panel(s, 0.667, 0.75, 12.0, 6.0, WIDE_CUT);
  s.addChart(pres.ChartType.bar,
    [{ name: 'Series 1', labels: ['Data 01', 'Data 02', 'Data 03', 'Data 04'],
       values: [57, 44, 72, 87] }],
    { x: 1.19, y: 1.457, w: 4.776, h: 5.004, barDir: 'col', barGapWidthPct: 80,
      chartColors: [C.mint], showLegend: false, showValue: true,
      dataLabelPosition: 'outEnd', dataLabelColor: C.white, dataLabelFontFace: FONT,
      dataLabelFontSize: 16, dataLabelFontBold: true,
      valAxisMaxVal: 100, valAxisMajorUnit: 20, valAxisLineShow: false,
      valAxisLabelColor: C.white, valAxisLabelFontFace: FONT, valAxisLabelFontSize: 12,
      catAxisLabelColor: C.white, catAxisLabelFontFace: FONT, catAxisLabelFontSize: 12,
      catAxisLineColor: C.white,
      valGridLine: { style: 'solid', size: 0.75, color: 'A98FB8' },
      plotArea: { fill: { type: 'none' } } });

  heading(s, 'Bar chart ', 'data', { x: 7.1, y: 1.25, w: 4.6, h: 0.769 });
  s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis.',
    txt({ x: 7.1, y: 2.1, w: 4.78, h: 0.87, lineSpacingMultiple: 1.3 }));

  const icons = [iconUsb, iconCloud, iconShare];
  const bodies = [
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do',
    'Lorem ipsum dolor sit amet, consectetur adipiscing eli. Sed do',
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do',
  ];
  bodies.forEach((body, i) => {
    const y = 3.28 + i * 0.99;
    iconChip(s, 7.191, y);
    icons[i](s, 7.32, y + 0.08, 0.35, 0.35);
    subtitleItem(s, 7.95, y - 0.1, { body, subColor: C.cyan, w: 3.9, h: 0.63 });
  });
}

/* --- 21: grouped bars + two donuts + 2x2 icon grid --------------------- */

function slide21(pres, s) {
  panel(s, 0.667, 0.75, 12.0, 6.0, WIDE_CUT);
  s.addChart(pres.ChartType.bar, [
    { name: 'Income',  labels: ['Category 1'], values: [46] },
    { name: 'Outcome', labels: ['Category 1'], values: [77] },
    { name: 'Supply',  labels: ['Category 1'], values: [64] },
  ], { x: 1.25, y: 2.882, w: 5.417, h: 3.579, barDir: 'col', barGapWidthPct: 80,
       barOverlapPct: -30, chartColors: [C.mint, C.mint, C.purple],
       showLegend: true, legendPos: 'b', legendColor: C.white, legendFontFace: FONT,
       legendFontSize: 12, showValue: true, dataLabelPosition: 'outEnd',
       dataLabelColor: C.white, dataLabelFontFace: FONT, dataLabelFontSize: 16,
       dataLabelFontBold: true, valAxisHidden: true, catAxisHidden: true,
       valAxisMaxVal: 100, valGridLine: { style: 'solid', size: 0.75, color: 'F2F2F2' },
       plotArea: { fill: { type: 'none' } } });

  heading(s, 'Bar chart ', 'data', { x: 1.33, y: 1.3, w: 5.51, h: 0.769 });
  s.addText(LOREM_SHORT, txt({ x: 1.33, y: 2.16, w: 5.51, h: 0.61, lineSpacingMultiple: 1.3 }));

  donutKpi(pres, s, 6.845, 1.446, { pct: '80%', withText: true });
  donutKpi(pres, s, 10.552, 1.446, { pct: '85%' });

  const icons = [iconUsb, iconShare, iconCloud, iconMoney];
  [[7.1, 3.9], [9.75, 3.9], [7.1, 5.19], [9.75, 5.19]].forEach(([x, y], i) => {
    iconChip(s, x, y);
    icons[i](s, x + 0.13, y + 0.08, 0.35, 0.35);
    s.addText('SUBTITLE HERE', txt({ x: x + 0.77 + (i % 2 ? 0.14 : 0), y: y - 0.11, w: 2.95, h: 0.36,
      fontSize: 11, bold: true, color: C.cyan, lineSpacingMultiple: 1.5 }));
    s.addText('Lorem ipsum dolor sit amet, consectetur', txt({ x: x + 0.77, y: y + 0.23,
      w: 1.71, h: 0.63, fontSize: 11, lineSpacingMultiple: 1.5 }));
  });
}

/* --- 22: bar chart with floating icon chips ---------------------------- */

function slide22(pres, s) {
  panel(s, 0.667, 0.75, 12.0, 6.0, WIDE_CUT);
  s.addChart(pres.ChartType.bar,
    [{ name: 'Series 1', labels: ['Best Data One', 'Best Data Two', 'Best Data Three', 'Best Data Four'],
       values: [60, 45, 75, 30] }],
    { x: 1.224, y: 2.002, w: 6.548, h: 4.468, barDir: 'col', barGapWidthPct: 70,
      chartColors: [C.mint], showLegend: false, showValue: true, dataLabelPosition: 'ctr',
      dataLabelColor: C.white, dataLabelFontFace: FONT, dataLabelFontSize: 16,
      dataLabelFontBold: true, valAxisHidden: true, catAxisHidden: true,
      valAxisMaxVal: 100, valAxisMajorUnit: 20,
      valGridLine: { style: 'solid', size: 0.75, color: 'C9A5DC' },
      plotArea: { fill: { type: 'none' } } });

  heading(s, 'Bar chart ', 'data', { x: 1.33, y: 1.3, w: 5.51, h: 0.769 });

  const chips = [[1.85, 3.18, iconMoney], [3.41, 3.85, iconShare], [4.98, 2.57, iconCloud]];
  chips.forEach(([x, y, icon]) => {
    iconChip(s, x, y);
    icon(s, x + 0.13, y + 0.08, 0.35, 0.35);
  });

  totalDataCard(s, 6.4, 3.23, '155M', true);
  s.addText(LOREM_TINY, txt({ x: 8.44, y: 2.0, w: 3.85, h: 0.71, bold: true, lineSpacingMultiple: 1.5 }));
  ['01', '02', '03', '04'].forEach((n, i) => {
    dataRow(s, 8.44, 2.82 + i * 0.563, 3.27, 'Your text ' + n, ['+153', '+345', '+111', '+336'][i]);
  });
  s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna',
    txt({ x: 8.44, y: 5.37, w: 3.48, h: 0.87, lineSpacingMultiple: 1.3 }));
}

/* --- 23: full width column chart --------------------------------------- */

function slide23(pres, s) {
  panel(s, 0.667, 0.75, 12.0, 6.0, WIDE_CUT);
  s.addChart(pres.ChartType.bar,
    [{ name: 'Series 1', labels: ['Data 01', 'Data 02', 'Data 03', 'Data 04', 'Data 05', 'Data 07'],
       values: [53, 35, 45, 60, 75, 85] }],
    { x: 0.914, y: 2.1, w: 11.241, h: 4.478, barDir: 'col', barGapWidthPct: 0,
      barOverlapPct: -27, chartColors: [C.mint], showLegend: false, showValue: false,
      valAxisMaxVal: 100, valAxisMajorUnit: 10, valAxisLineShow: false,
      valAxisLabelColor: C.white, valAxisLabelFontFace: FONT, valAxisLabelFontSize: 12,
      catAxisLabelColor: C.white, catAxisLabelFontFace: FONT, catAxisLabelFontSize: 12,
      catAxisLineColor: C.white,
      valGridLine: { style: 'solid', size: 0.75, color: 'F2F2F2' },
      plotArea: { fill: { type: 'none' } } });

  heading(s, 'Bar chart ', 'data', { x: 3.91, y: 0.92, w: 5.51, h: 0.769, align: 'center' });
  s.addText(LOREM_SHORT, txt({ x: 3.56, y: 1.56, w: 6.22, h: 0.61, align: 'center',
    lineSpacingMultiple: 1.3 }));
  [['35%', 1.65, 4.84], ['25%', 3.44, 5.19], ['30%', 5.23, 5.01],
   ['40%', 7.0, 4.69], ['60%', 8.72, 4.37], ['70%', 10.45, 4.05]]
    .forEach(([t, x, y]) => s.addText(t, txt({ x, y, w: 1.4, h: 0.64, align: 'center',
      fontSize: 32, bold: true })));
}

/* --- 24: two-series chart + two data cards ------------------------------ */

function slide24(pres, s) {
  panel(s, 0.667, 0.75, 12.0, 6.0, WIDE_CUT);
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May ', 'Jun'];
  s.addChart(pres.ChartType.bar, [
    { name: 'Series 1', labels: months, values: [2.0, 2.5, 3.5, 4.0, 4.6, 5.0] },
    { name: 'Series 2', labels: months, values: [1.3, 2.7, 3.0, 4.3, 4.0, 5.5] },
  ], { x: 1.081, y: 3.083, w: 7.009, h: 3.356, barDir: 'col', barGapWidthPct: 80,
       barOverlapPct: -27, chartColors: ['662A83', C.mint], showLegend: false,
       showValue: false, valAxisHidden: true,
       catAxisLabelColor: C.white, catAxisLabelFontFace: FONT, catAxisLabelFontSize: 12,
       catAxisLineColor: C.white, valGridLine: { style: 'none' },
       plotArea: { fill: { type: 'none' } } });

  heading(s, 'Your title ', 'here', { x: 1.08, y: 1.33, w: 5.25, h: 0.769 });
  s.addText(LOREM, txt({ x: 1.08, y: 2.13, w: 7.15, h: 1.29, lineSpacingMultiple: 1.5 }));

  [['155M', 2.47], ['175K', 4.47]].forEach(([value, y]) => {
    totalDataCard(s, 8.71, y, value, false);
    subtitleItem(s, 10.53, y + 0.21, { body: 'Lorem ipsum dolor sit amet, consectetur',
      subColor: C.cyan, w: 1.71, h: 0.63 });
  });
}

/* ================================================================== build */

function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'W16x9', width: SLIDE_W, height: SLIDE_H });
  pres.layout = 'W16x9';
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pres.title = 'Futuristic Presentation Template';

  pres.defineSlideMaster({
    title: 'FUTURISTIC',
    background: { color: rampColor(0.5 / 56) },
    objects: backgroundBands(),
  });

  // Footer strip repeated on every slide by the original slide master.
  const footer = (s, n) => {
    s.addText('www.website.com', txt({ x: 0.667, y: 6.749, w: 2.847, h: 0.269,
      fontSize: 10, color: C.cyan, charSpacing: 3, valign: 'middle' }));
    s.addText(String(n), txt({ x: 12.17, y: 6.766, w: 0.488, h: 0.269, fontSize: 10,
      bold: true, color: C.cyan, align: 'center', charSpacing: 2, valign: 'middle' }));
  };

  const builders = [
    slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
    slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
    slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24,
    slide25,
  ];

  builders.forEach((fn, i) => {
    const s = pres.addSlide({ masterName: 'FUTURISTIC' });
    if (fn.length === 2) fn(pres, s); else fn(s);
    footer(s, i + 1);
  });

  return pres.writeFile({
    fileName: path.join(__dirname, '116168bd-777d-444d-8f3b-cda0c3f2738c_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
