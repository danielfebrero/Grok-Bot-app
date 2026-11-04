/**
 * "Venn Diagram - Infographic Presentation Template" rebuilt with pptxgenjs.
 * 16 slides, 13.333 x 7.5 in (16:9).  Run: node <thisfile>.js
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette --
const C = {
  purple: '4A24EE',   // theme accent1
  pink: 'F88DCB',     // theme accent2
  green: '10DD93',    // theme accent3
  yellow: 'F9D722',   // theme accent4
  blue: '00B0F0',     // theme accent5
  purpleLt: '927CF5', // accent1 lum 60/40
  pinkLt: 'FBBBE0',
  greenLt: '66F4C1',
  white: 'FFFFFF',
  black: '000000',
  dark: '262626',     // tx1 lum 85/15  - headlines
  gray: '808080',     // tx1 lum 50/50  - body copy
  gray65: '595959',   // tx1 lum 65/35
  gray7F: '7F7F7F',
  rule: 'BFBFBF',     // bg1 lum 75     - dashed leader lines
  ring: 'E7E6E6',     // hairline around white "hub" circles
};

const FONT_HEAD = 'Plus Jakarta Sans ExtraBold';
const FONT_BODY = 'Plus Jakarta Sans';

// Drop shadows used throughout the deck (blur/offset in points, angle in deg).
const SH_BLOB = { type: 'outer', color: C.black, opacity: 0.23, blur: 40, offset: 18, angle: 45 };
const SH_HUB = { type: 'outer', color: C.black, opacity: 0.12, blur: 14.6, offset: 4.5, angle: 67 };
const SH_CARD = { type: 'outer', color: C.black, opacity: 0.2, blur: 35, offset: 10, angle: 90 };
const SH_BADGE = { type: 'outer', color: C.black, opacity: 0.15, blur: 40, offset: 10, angle: 90 };

const NOLINE = { type: 'none' };

// ------------------------------------------------------------ text helper --
/** Body/caption text.  `o` overrides any default. */
function txt(s, text, o) {
  s.addText(text, Object.assign({
    fontFace: FONT_BODY, fontSize: 12, color: C.gray,
    align: 'left', valign: 'top',
  }, o));
}

/** Slide headline ("Venn Diagram") + grey subtitle underneath. */
function heading(s, o) {
  const y = o.y, w = o.w || 5.102;
  txt(s, o.title || 'Venn Diagram', {
    x: o.x, y: y, w: w, h: o.titleH || 0.767,
    fontFace: FONT_HEAD, fontSize: o.size || 44, bold: o.bold !== false,
    color: C.dark, align: o.align || 'left', lineSpacingMultiple: 0.9,
  });
  txt(s, o.sub || 'Subtitle Text Goes Here', {
    x: o.x, y: o.subY !== undefined ? o.subY : y + 0.772, w: w, h: 0.337,
    fontSize: o.subSize || 14, color: o.subColor || C.gray, align: o.align || 'left',
  });
}

// ------------------------------------------------------------ shape atoms --
function blob(s, shape, x, y, w, h, color, o) {
  s.addShape(shape, Object.assign({
    x: x, y: y, w: w, h: h, line: NOLINE,
    fill: { color: color, transparency: 15 }, shadow: SH_BLOB,
  }, o || {}));
}

function circle(s, x, y, d, color, o) {
  blob(s, 'ellipse', x, y, d, d, color, o);
}

/** Small white "hub" disc with a hairline ring (centre of a venn). */
function hub(s, x, y, d, o) {
  s.addShape('ellipse', Object.assign({
    x: x, y: y, w: d, h: d, fill: { color: C.white },
    line: { color: C.ring, width: 0.75 }, shadow: SH_HUB,
  }, o || {}));
}

/** Coloured bullet disc with a small white chevron; dir 'r' (default) or 'l'. */
function arrowDot(s, x, y, color, dir) {
  const D = 0.277, cx = x + D / 2, cy = y + D / 2, sgn = dir === 'l' ? -1 : 1;
  s.addShape('ellipse', { x: x, y: y, w: D, h: D, fill: { color: color }, line: NOLINE });
  [-1, 1].forEach(side => {                    // two strokes meeting at the tip
    const len = 0.083, thick = 0.017;
    s.addShape('rect', {
      x: cx + sgn * 0.004 - len / 2, y: cy + side * 0.0275 - thick / 2, w: len, h: thick,
      fill: { color: C.white }, line: NOLINE, rotate: -side * sgn * 41,
    });
  });
}

// ------------------------------------------------------------------ icons --
// Flat pictogram placeholders built from native shapes.  `hole` is the colour
// showing through cut-outs (i.e. whatever the icon sits on).
function icon(s, kind, x, y, w, h, color, hole) {
  const bg = hole || C.white;
  const box = (sx, sy, sw, sh, c, rad) => s.addShape(rad ? 'roundRect' : 'rect', {
    x: x + sx * w, y: y + sy * h, w: sw * w, h: sh * h,
    fill: { color: c }, line: NOLINE,
    rectRadius: rad ? Math.min(sw * w, sh * h) * rad : undefined,
  });
  if (kind === 'box') {                       // archive box: lid + body + slot
    box(0, 0, 1, 0.26, color, 0.35);
    box(0.06, 0.3, 0.88, 0.7, color, 0.16);
    box(0.3, 0.42, 0.4, 0.14, bg);
  } else if (kind === 'clipboard') {          // board whose top edge rises to a tab
    roundedPoly(s, x, y, w, h,
      [[0, 0.21], [0.27, 0.21], [0.36, 0], [0.64, 0], [0.73, 0.21], [1, 0.21],
       [1, 1], [0, 1]], 0.03 * h, { fill: { color: color } });
  } else if (kind === 'briefcase') {          // case + handle + side seams
    box(0.16, 0.2, 0.68, 0.8, color);
    box(0, 0.28, 0.12, 0.72, color, 0.3);
    box(0.88, 0.28, 0.12, 0.72, color, 0.3);
    box(0.34, 0.06, 0.32, 0.16, color, 0.35);
    box(0.4, 0.14, 0.2, 0.14, bg);
    box(0.44, 0.44, 0.12, 0.12, bg);
  } else if (kind === 'calendar') {           // frame + two hangers + window
    box(0.06, 0.14, 0.88, 0.86, color, 0.16);
    box(0.26, 0, 0.09, 0.26, color);
    box(0.65, 0, 0.09, 0.26, color);
    box(0.18, 0.42, 0.64, 0.42, bg);
  } else if (kind === 'envelope') {           // body with a folded-flap notch
    box(0, 0.1, 1, 0.9, color, 0.14);
    box(0, 0, 1, 0.16, color);
    s.addShape('triangle', { x: x + 0.05 * w, y: y + 0.05 * h, w: 0.9 * w, h: 0.52 * h,
      fill: { color: bg }, line: NOLINE, flipV: true });
  }
}

// ------------------------------------------------- circular-arc freeforms --
/** Cubic-bezier samples of the arc c=[cx,cy] r, sweeping a0 -> a1 (radians). */
function arcPts(c, r, a0, a1) {
  const n = Math.max(1, Math.ceil(Math.abs(a1 - a0) / (Math.PI / 2)));
  const step = (a1 - a0) / n, k = (4 / 3) * Math.tan(step / 4), out = [];
  for (let i = 0; i < n; i++) {
    const s0 = a0 + i * step, s1 = s0 + step;
    const px = c[0] + r * Math.cos(s0), py = c[1] + r * Math.sin(s0);
    const qx = c[0] + r * Math.cos(s1), qy = c[1] + r * Math.sin(s1);
    out.push({ x: qx, y: qy, curve: { type: 'cubic',
      x1: px - k * r * Math.sin(s0), y1: py + k * r * Math.cos(s0),
      x2: qx + k * r * Math.sin(s1), y2: qy - k * r * Math.cos(s1) } });
  }
  return out;
}

/** Closed freeform bounded by a chain of arcs: [{c,r,a0,a1}, ...]. */
function arcShape(s, arcs, opts) {
  const pts = [];
  arcs.forEach(a => pts.push(...arcPts(a.c, a.r, a.a0, a.a1)));
  const xs = [], ys = [];
  arcs.forEach(a => {                                   // sample for the bbox
    for (let i = 0; i <= 24; i++) {
      const t = a.a0 + (a.a1 - a.a0) * i / 24;
      xs.push(a.c[0] + a.r * Math.cos(t));
      ys.push(a.c[1] + a.r * Math.sin(t));
    }
  });
  const x0 = Math.min(...xs), y0 = Math.min(...ys);
  const w = Math.max(...xs) - x0, h = Math.max(...ys) - y0;
  const first = arcs[0];
  const start = { x: first.c[0] + first.r * Math.cos(first.a0) - x0,
                  y: first.c[1] + first.r * Math.sin(first.a0) - y0, moveTo: true };
  const rel = pts.map(p => ({ x: p.x - x0, y: p.y - y0, curve: {
    type: 'cubic', x1: p.curve.x1 - x0, y1: p.curve.y1 - y0,
    x2: p.curve.x2 - x0, y2: p.curve.y2 - y0 } }));
  s.addShape('custGeom', Object.assign({
    x: x0, y: y0, w: w, h: h, line: NOLINE,
    points: [start, ...rel, { close: true }],
  }, opts));
}

/**
 * Closed polygon with eased corners.  `verts` are fractions of the w x h box;
 * `r` is the corner tangent length (one value, or one per vertex).
 */
function roundedPoly(s, x, y, w, h, verts, r, opts) {
  const P = verts.map(v => [v[0] * w, v[1] * h]), n = P.length, pts = [];
  const towards = (a, b, dist) => {
    const t = Math.min(0.5, dist / Math.hypot(b[0] - a[0], b[1] - a[1]));
    return [a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t];
  };
  for (let i = 0; i < n; i++) {
    const ri = Array.isArray(r) ? r[i] : r, cur = P[i];
    const a = towards(cur, P[(i + n - 1) % n], ri), b = towards(cur, P[(i + 1) % n], ri);
    pts.push(i === 0 ? { x: a[0], y: a[1], moveTo: true } : { x: a[0], y: a[1] });
    pts.push({ x: b[0], y: b[1], curve: { type: 'quadratic', x1: cur[0], y1: cur[1] } });
  }
  pts.push({ close: true });
  s.addShape('custGeom', Object.assign({ x: x, y: y, w: w, h: h, points: pts, line: NOLINE }, opts));
}

/**
 * Soft-cornered hexagon, points at top and bottom of the w x h box.  The two
 * apexes sit slightly outside it so that, once eased, the outline fills the box.
 */
function hexagon(s, x, y, w, h, color) {
  const f = 0.2426;   // where the vertical side starts, as a fraction of h
  const ea = 0.008;   // apex overshoot, cancelled by the corner easing
  roundedPoly(s, x, y, w, h,
    [[0.5, -ea], [1, f], [1, 1 - f], [0.5, 1 + ea], [0, 1 - f], [0, f]],
    [0.032, 0.053, 0.053, 0.032, 0.053, 0.053].map(v => v * h),
    { fill: { color: color, transparency: 15 }, shadow: SH_BLOB });
}

/** Half-angle subtended at `a` by the two intersections of equal circles a, b. */
function halfChord(a, b, r) {
  return Math.acos(Math.hypot(b[0] - a[0], b[1] - a[1]) / 2 / r);
}

/** Lens = intersection of two equal circles centred at a and b. */
function lens(s, a, b, r, opts) {
  const th = Math.atan2(b[1] - a[1], b[0] - a[0]), al = halfChord(a, b, r);
  arcShape(s, [
    { c: a, r: r, a0: th + al, a1: th - al },
    { c: b, r: r, a0: th + Math.PI + al, a1: th + Math.PI - al },
  ], opts);
}

/** Crescent = circle a with the overlapping part of circle b bitten out. */
function crescent(s, a, b, r, opts) {
  const th = Math.atan2(b[1] - a[1], b[0] - a[0]), al = halfChord(a, b, r);
  arcShape(s, [
    { c: a, r: r, a0: th - al, a1: th + al - 2 * Math.PI },
    { c: b, r: r, a0: th + Math.PI - al, a1: th + Math.PI + al },
  ], opts);
}

// ============================================================ slide 1 ======
function slide01(pres) {
  const s = pres.addSlide();
  circle(s, 6.222, 1.636, 0.87, C.purple);
  circle(s, 6.479, 2.079, 0.87, C.green);
  circle(s, 5.985, 2.079, 0.87, C.pink);
  txt(s, 'VENN DIAGRAM', { x: 2.533, y: 3.372, w: 8.268, h: 1.313,
    fontFace: FONT_HEAD, fontSize: 72, color: C.dark, align: 'center' });
  txt(s, 'Infographic Presentation Template', { x: 2.533, y: 4.671, w: 8.268, h: 0.438,
    fontSize: 20, color: C.dark, align: 'center' });
}

// ============================================================ slide 2 ======
// Three teardrops in a row + three bullet rows underneath.
function slide02(pres) {
  const s = pres.addSlide();
  heading(s, { x: 4.125, y: 0.613, align: 'center' });

  const drops = [
    { x: 1.938, tx: 2.551, ty: 2.838, color: C.purple, value: '150+' },
    { x: 4.996, tx: 6.025, ty: 2.832, color: C.pink, value: '300+' },
    { x: 8.053, tx: 9.24, ty: 2.838, color: C.green, value: '450+' },
  ];
  drops.forEach(d => {
    blob(s, 'teardrop', d.x, 2.073, 3.373, 3.369, d.color, { rotate: 135, flipH: true });
    txt(s, d.value, { x: d.tx, y: d.ty, w: 2.088, h: 0.841,
      fontFace: FONT_HEAD, fontSize: 44, color: C.white });
    txt(s, 'Value Here', { x: d.tx, y: d.ty + 0.836, w: 2.088, h: 0.404,
      fontSize: 20, color: C.white, lineSpacingMultiple: 0.9 });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur', {
      x: d.tx, y: d.ty + 1.234, w: 2.088, h: 0.604, color: C.white, lineSpacingMultiple: 1.3 });
  });

  const rows = [
    { x: 2.239, color: C.purple, label: 'Data Number One' },
    { x: 5.474, color: C.pink, label: 'Data Number Two' },
    { x: 8.71, color: C.green, label: 'Data Number Three' },
  ];
  rows.forEach(r => {
    arrowDot(s, r.x, 5.869, r.color);
    txt(s, r.label, { x: r.x + 0.388, y: 5.794, w: 2.385, h: 0.37,
      fontSize: 16, bold: true, color: r.color });
    txt(s, 'Lorem ipsum dolor sit amet', { x: r.x + 0.388, y: 6.172, w: 2.385, h: 0.341,
      lineSpacingMultiple: 1.3 });
  });
}

// ============================================================ slide 3 ======
// Two overlapping circles with a "$1M profit" overlap label, side captions.
function slide03(pres) {
  const s = pres.addSlide();
  heading(s, { x: 4.125, y: 0.613, align: 'center', bold: false, subY: 1.395 });

  circle(s, 3.615, 2.478, 3.713, C.purple);
  circle(s, 5.877, 2.478, 3.713, C.pink);
  txt(s, '35%', { x: 3.986, y: 3.695, w: 2.088, h: 0.841,
    fontFace: FONT_HEAD, fontSize: 44, color: C.white });
  txt(s, 'Product A', { x: 3.986, y: 4.57, w: 2.088, h: 0.404,
    fontSize: 20, color: C.white, lineSpacingMultiple: 0.9 });
  txt(s, '65%', { x: 6.768, y: 3.695, w: 2.448, h: 0.841,
    fontFace: FONT_HEAD, fontSize: 44, color: C.white, align: 'right' });
  txt(s, 'Product A', { x: 6.768, y: 4.57, w: 2.462, h: 0.404,
    fontSize: 20, color: C.white, align: 'right', lineSpacingMultiple: 0.9 });
  txt(s, '$1M', { x: 5.935, y: 3.838, w: 1.467, h: 0.643,
    fontFace: FONT_HEAD, fontSize: 28, bold: true, color: C.white, align: 'center' });
  txt(s, 'Profit', { x: 5.935, y: 4.495, w: 1.467, h: 0.337,
    fontSize: 14, color: C.white, align: 'center' });

  sideCaption(s, { x: 0.781, cx: 0.786, ix: 0.978, kind: 'box', color: C.purple,
    align: 'left', label: 'Data One', amount: '$45.000.000' });
  sideCaption(s, { x: 10.122, cx: 11.883, ix: 12.103, kind: 'clipboard', color: C.pink,
    align: 'right', label: 'Data Two', amount: '$85.000.000' });
}

/** Icon disc + label + paragraph + money figure (slide 3). */
function sideCaption(s, o) {
  hub(s, o.cx, 2.903, 0.669);
  const iw = o.kind === 'clipboard' ? 0.229 : 0.286;
  const ih = o.kind === 'clipboard' ? 0.317 : 0.254;
  icon(s, o.kind, o.ix, o.kind === 'clipboard' ? 3.079 : 3.111, iw, ih, o.color);
  txt(s, o.label, { x: o.x, y: 3.893, w: 2.425, h: 0.37, fontSize: 16, color: o.color, align: o.align });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', {
    x: o.x, y: 4.282, w: 2.425, h: 0.859, align: o.align, lineSpacingMultiple: 1.3 });
  txt(s, o.amount, { x: o.x, y: 5.362, w: 2.425, h: 0.404, fontSize: 18, color: o.color, align: o.align });
}

// ============================================================ slide 4 ======
// Four-lobe circle cluster on the right, numbered bullet list on the left.
function slide04(pres) {
  const s = pres.addSlide();
  heading(s, { x: 1.304, y: 1.001, size: 48, bold: false, subY: 1.768, subSize: 16 });

  const rows = [
    { y: 2.901, color: C.purple, label: 'Data Number One' },
    { y: 3.879, color: C.pink, label: 'Data Number Two' },
    { y: 4.858, color: C.green, label: 'Data Number Three' },
    { y: 5.836, color: C.yellow, label: 'Data Number Four' },
  ];
  rows.forEach(r => {
    arrowDot(s, 1.466, r.y + 0.075, r.color);
    txt(s, r.label, { x: 1.854, y: r.y, w: 3.499, h: 0.37, fontSize: 16, bold: true, color: r.color });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur', {
      x: 1.854, y: r.y + 0.378, w: 3.499, h: 0.341, lineSpacingMultiple: 1.3 });
  });

  const lobes = [
    { x: 6.737, y: 2.551, color: C.yellow, value: '96M', tx: 7.076, ty: 3.396, tw: 1.816, align: 'left' },
    { x: 8.197, y: 3.921, color: C.green, value: '28M', tx: 8.529, ty: 5.279, tw: 2.256, align: 'center' },
    { x: 9.657, y: 2.551, color: C.pink, value: '32M', tx: 10.375, ty: 3.396, tw: 1.862, align: 'right' },
    { x: 8.197, y: 1.001, color: C.purple, value: '60M', tx: 8.486, ty: 1.502, tw: 2.256, align: 'center' },
  ];
  lobes.forEach(l => {
    circle(s, l.x, l.y, 2.92, l.color);
    txt(s, l.value, { x: l.tx, y: l.ty, w: l.tw, h: l.align === 'center' ? 0.817 : 0.707,
      fontFace: FONT_HEAD, fontSize: 36, bold: true, color: C.white, align: l.align });
    txt(s, 'Your Text', { x: l.tx, y: l.ty + 0.767, w: l.tw, h: 0.389,
      fontSize: 14, color: C.white, align: l.align });
  });
  hub(s, 8.848, 3.135, 1.571);
  icon(s, 'clipboard', 9.365, 3.549, 0.537, 0.745, C.black);
}

// ============================================================ slide 5 ======
// Four overlapping hexagons in a staircase + footnote paragraph.
function slide05(pres) {
  const s = pres.addSlide();
  heading(s, { x: 4.116, y: 0.624, align: 'center', subY: 1.396 });

  const hexes = [
    { x: 1.471, y: 2.514, tx: 1.825, ty: 3.298, color: C.purple, value: '45%' },
    { x: 4.003, y: 2.232, tx: 4.539, ty: 3.017, color: C.pink, value: '30%' },
    { x: 6.535, y: 1.883, tx: 7.07, ty: 2.654, color: C.green, value: '20%' },
    { x: 9.067, y: 2.582, tx: 9.552, ty: 3.367, color: C.yellow, value: '50%' },
  ];
  hexes.forEach(h => {
    hexagon(s, h.x, h.y, 2.796, 3.139, h.color);
    txt(s, h.value, { x: h.tx, y: h.ty, w: 2.088, h: 0.841,
      fontFace: FONT_HEAD, fontSize: 44, color: C.white });
    txt(s, 'Value Here', { x: h.tx, y: h.ty + 0.837, w: 2.088, h: 0.404,
      fontSize: 20, color: C.white, lineSpacingMultiple: 0.9 });
    txt(s, 'Lorem ipsum dolor sit', { x: h.tx, y: h.ty + 1.235, w: 2.088, h: 0.334,
      color: C.white, lineSpacingMultiple: 1.3 });
  });

  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
        'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud ' +
        'exercitation ullamco laboris nisi ut aliquip',
    { x: 2.492, y: 5.976, w: 8.35, h: 0.667, color: C.gray65, align: 'center', lineSpacingMultiple: 1.5 });
}

// ============================================================ slide 6 ======
// Three-circle venn with a white hub and dashed leader lines to captions.
function slide06(pres) {
  const s = pres.addSlide();
  heading(s, { x: 0.713, y: 0.684 });

  [[2.759, 4.7], [6.889, 2.16], [8.833, 4.846]].forEach(p => {
    s.addShape('line', { x: p[0], y: p[1], w: 1.639, h: 0,
      line: { color: C.rule, width: 1, dashType: 'dash', endArrowType: 'oval' } });
  });

  // Green is a whole circle; purple and pink are bitten back by it, so green
  // reads on top wherever it overlaps them.
  const R = 1.847, P = [6.481, 3.773], G = [5.863, 5.122], K = [7.445, 5.129];
  const lobes = [
    { c: K, cut: true, color: C.pink,
      vx: 7.445, vy: 5.027, vw: 1.622, lx: 7.359, lw: 1.72, align: 'right', pct: 20 },
    { c: G, cut: false, color: C.green,
      vx: 4.264, vy: 4.7, vw: 1.548, lx: 4.178, lw: 1.72, align: 'left', pct: 18 },
    { c: P, cut: true, color: C.purple,
      vx: 5.747, vy: 2.174, vw: 1.548, lx: 5.661, lw: 1.72, align: 'center', pct: 20 },
  ];
  lobes.forEach(l => {
    const style = { fill: { color: l.color, transparency: 15 }, shadow: SH_BLOB };
    if (l.cut) crescent(s, l.c, G, R, style);
    else s.addShape('ellipse', Object.assign({
      x: l.c[0] - R, y: l.c[1] - R, w: 2 * R, h: 2 * R, line: NOLINE }, style));
    s.addText([
      { text: '70', options: { fontSize: 32 } },
      { text: '%', options: { fontSize: l.pct } },
    ], { x: l.vx, y: l.vy, w: l.vw, h: 0.64, fontFace: FONT_HEAD, bold: true,
      color: C.white, align: l.align, valign: 'top', margin: 0 });
    txt(s, 'Monthly Profit', { x: l.lx, y: l.vy + 0.613, w: l.lw, h: 0.365,
      fontSize: 14, color: C.white, align: l.align });
  });
  hub(s, 5.697, 3.708, 1.916);
  icon(s, 'clipboard', 6.416, 4.334, 0.478, 0.663, C.black);

  valueBlock(s, { x: 8.903, y: 1.43, w: 2.695, color: C.purple,
    value: '$55,9M', label: 'Value One', bodyH: 0.596 });
  valueBlock(s, { x: 10.001, y: 5.072, w: 2.695, color: C.pink,
    value: '$65,9M', label: 'Value Two', bodyH: 0.596 });
  valueBlock(s, { x: 0.624, y: 4.92, w: 2.718, color: C.green, align: 'right',
    value: '$55,9M', label: 'Value Three', bodyH: 0.596 });
}

/** Big money figure + coloured label + grey paragraph (slides 6, 9, 13). */
function valueBlock(s, o) {
  const align = o.align || 'left';
  txt(s, o.value, { x: o.x, y: o.y, w: o.w, h: 0.505, fontSize: 24, color: o.color, align: align });
  txt(s, o.label, { x: o.x, y: o.y + 0.471, w: o.w, h: 0.37,
    fontSize: 16, color: o.labelColor || o.color, align: align });
  txt(s, o.body || 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed', {
    x: o.x, y: o.y + 0.813, w: o.w, h: o.bodyH || 0.334, align: align, lineSpacingMultiple: 1.3 });
}

// ============================================================ slide 7 ======
// Nested circles (455k / 298k / 356k) plus a bullet list on the right.
function slide07(pres) {
  const s = pres.addSlide();
  blob(s, 'ellipse', 0.971, 1.025, 5.283, 5.396, C.purple, { shadow: SH_CARD });
  blob(s, 'ellipse', 1.875, 2.916, 3.477, 3.551, C.pink, { shadow: SH_CARD });
  blob(s, 'ellipse', 2.578, 4.361, 2.07, 2.114, C.green, { shadow: SH_CARD });

  const rings = [
    { v: '455k', vy: 1.331, vh: 0.841, vs: 44, l: 'Product A', ly: 2.206, lh: 0.404, ls: 20, x: 2.569 },
    { v: '298k', vy: 3.149, vh: 0.774, vs: 40, l: 'Product B', ly: 3.923, lh: 0.374, ls: 18, x: 2.569 },
    { v: '356k', vy: 4.872, vh: 0.707, vs: 36, l: 'Product B', ly: 5.573, lh: 0.343, ls: 16, x: 2.578 },
  ];
  rings.forEach(r => {
    txt(s, r.v, { x: r.x, y: r.vy, w: 2.088, h: r.vh,
      fontFace: FONT_HEAD, fontSize: r.vs, color: C.white, align: 'center' });
    txt(s, r.l, { x: r.x, y: r.ly, w: 2.088, h: r.lh,
      fontSize: r.ls, color: C.white, align: 'center', lineSpacingMultiple: 0.9 });
  });

  heading(s, { x: 7.26, y: 1.325 });
  const rows = [
    { y: 3.197, color: C.purple, label: 'Data Number One' },
    { y: 4.175, color: C.pink, label: 'Data Number Two' },
    { y: 5.154, color: C.green, label: 'Data Number Three' },
  ];
  rows.forEach(r => {
    arrowDot(s, 7.336, r.y + 0.075, r.color);
    txt(s, r.label, { x: 7.724, y: r.y, w: 3.499, h: 0.37, fontSize: 16, bold: true, color: r.color });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur', {
      x: 7.724, y: r.y + 0.378, w: 3.499, h: 0.341, lineSpacingMultiple: 1.3 });
  });
}

// ============================================================ slide 8 ======
// Classic three-circle venn (numbered overlaps) + three captions.
function slide08(pres) {
  const s = pres.addSlide();
  const R = 1.645;
  const P = [10.115, 3.104], G = [9.256, 4.674], K = [10.92, 4.674];

  [[P, C.purple], [G, C.green], [K, C.pink]].forEach(([c, col]) => {
    s.addShape('ellipse', { x: c[0] - R, y: c[1] - R, w: 2 * R, h: 2 * R,
      fill: { color: col }, line: NOLINE, shadow: SH_BLOB });
  });
  lens(s, P, G, R, { fill: { color: C.purpleLt } });
  lens(s, P, K, R, { fill: { color: C.pinkLt } });
  lens(s, G, K, R, { fill: { color: C.greenLt } });
  tripleCore(s, P, G, K, R);

  [['1', 8.862, 3.261], ['2', 10.997, 3.308], ['3', 9.915, 4.95]].forEach(([n, x, y]) => {
    txt(s, n, { x: x, y: y, w: 0.312, h: 0.64,
      fontFace: FONT_HEAD, fontSize: 32, bold: true, color: C.white, align: 'center', valign: 'middle' });
  });
  hub(s, 9.626, 1.761, 0.929); icon(s, 'clipboard', 9.932, 2.002, 0.318, 0.441, C.purple);
  hub(s, 11.269, 4.583, 0.929); icon(s, 'calendar', 11.535, 4.849, 0.396, 0.396, C.pink);
  hub(s, 7.988, 4.583, 0.929); icon(s, 'box', 8.254, 4.871, 0.396, 0.352, C.green);

  heading(s, { x: 0.909, y: 1.229 });
  const cols = [
    { x: 0.76, cx: 0.765, ix: 0.957, iy: 3.611, iw: 0.286, ih: 0.254, kind: 'box',
      color: C.purple, label: 'Data One', amount: '$45.000.000' },
    { x: 2.978, cx: 2.978, ix: 3.198, iy: 3.579, iw: 0.229, ih: 0.317, kind: 'clipboard',
      color: C.pink, label: 'Data Two', amount: '$85.000.000' },
    { x: 5.152, cx: 5.157, ix: 5.332, iy: 3.595, iw: 0.32, ih: 0.286, kind: 'briefcase',
      color: C.green, label: 'Data Three', amount: '$75.000.000' },
  ];
  cols.forEach(c => {
    s.addShape('ellipse', { x: c.cx, y: 3.403, w: 0.669, h: 0.669,
      fill: { color: c.color }, line: NOLINE, shadow: SH_HUB });
    icon(s, c.kind, c.ix, c.iy, c.iw, c.ih, C.white, c.color);
    txt(s, c.label, { x: c.x, y: 4.393, w: 1.985, h: 0.37, fontSize: 16, color: c.color });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do', {
      x: c.x, y: 4.782, w: 1.985, h: 0.859, lineSpacingMultiple: 1.3 });
    txt(s, c.amount, { x: c.x, y: 5.862, w: 1.985, h: 0.404, fontSize: 18, color: c.color });
  });
}

/** White curved triangle where all three venn circles overlap. */
function tripleCore(s, A, B, D, r) {
  const meet = (p, q, other) => {                 // intersection nearest `other`
    const th = Math.atan2(q[1] - p[1], q[0] - p[0]);
    const al = Math.acos(Math.hypot(q[0] - p[0], q[1] - p[1]) / 2 / r);
    return [th + al, th - al]
      .map(a => [p[0] + r * Math.cos(a), p[1] + r * Math.sin(a)])
      .sort((u, v) => Math.hypot(u[0] - other[0], u[1] - other[1]) -
                      Math.hypot(v[0] - other[0], v[1] - other[1]))[0];
  };
  const ang = (c, p) => Math.atan2(p[1] - c[1], p[0] - c[0]);
  const norm = (a0, a1) => {                      // pick the short way round
    let d = a1 - a0;
    while (d > Math.PI) d -= 2 * Math.PI;
    while (d < -Math.PI) d += 2 * Math.PI;
    return a0 + d;
  };
  const ab = meet(A, B, D), ad = meet(A, D, B), bd = meet(B, D, A);
  arcShape(s, [
    { c: A, r: r, a0: ang(A, ab), a1: norm(ang(A, ab), ang(A, ad)) },
    { c: D, r: r, a0: ang(D, ad), a1: norm(ang(D, ad), ang(D, bd)) },
    { c: B, r: r, a0: ang(B, bd), a1: norm(ang(B, bd), ang(B, ab)) },
  ], { fill: { color: C.white } });
}

// ============================================================ slide 9 ======
// Three staggered rounded cards with icon badges + captions around them.
function slide09(pres) {
  const s = pres.addSlide();
  valueBlock(s, { x: 0.711, y: 4.871, w: 2.526, color: C.pink, align: 'right',
    value: '$560.000', label: 'Value Two', bodyH: 0.859, body: LONG_BODY });
  valueBlock(s, { x: 4.774, y: 1.186, w: 2.526, color: C.purple,
    value: '$560.000', label: 'Value One', bodyH: 1.121, body: LONG_BODY });
  valueBlock(s, { x: 9.823, y: 3.75, w: 2.384, color: C.green,
    value: '$560.000', label: 'Value Three', bodyH: 1.121, body: LONG_BODY });
  heading(s, { x: 7.511, y: 1.192, align: 'right' });

  const cards = [
    { x: 0.924, y: 1.964, color: C.purple, label: 'Data 01', kind: 'clipboard',
      ix: 3.603, iy: 2.378, iw: 0.247, ih: 0.343, cx: 3.365, cy: 2.191, lx: 1.123, ly: 3.549 },
    { x: 3.522, y: 3.258, color: C.pink, label: 'Data 02', kind: 'calendar',
      ix: 6.151, iy: 3.718, iw: 0.308, ih: 0.308, cx: 5.944, cy: 3.511, lx: 3.72, ly: 4.871 },
    { x: 6.119, y: 4.552, color: C.green, label: 'Data 03', kind: 'box',
      ix: 8.742, iy: 5.032, iw: 0.308, ih: 0.274, cx: 8.535, cy: 4.808, lx: 6.305, ly: 6.131 },
  ];
  cards.forEach(c => {
    blob(s, 'roundRect', c.x, c.y, 3.419, 2.31, c.color, { rectRadius: 0.308, shadow: SH_CARD });
    hub(s, c.cx, c.cy, 0.723);
    icon(s, c.kind, c.ix, c.iy, c.iw, c.ih, c.color);
    txt(s, c.label, { x: c.lx, y: c.ly, w: 1.916, h: 0.525,
      fontFace: FONT_HEAD, fontSize: 28, bold: true, color: C.white, lineSpacingMultiple: 0.9 });
  });
}
const LONG_BODY = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor et';

// ============================================================ slide 10 =====
// Three petal ellipses arranged as a pinwheel, four bullet rows around it.
function slide10(pres) {
  const s = pres.addSlide();
  heading(s, { x: 4.116, y: 0.756, align: 'center' });

  const petals = [
    { x: 4.72, y: 3.216, h: 2.158, rot: 90, flipH: true, color: C.purple, kind: 'box',
      ix: 6.524, iy: 2.737, iw: 0.286, ih: 0.254, tx: 6.072, ty: 3.218, align: 'center' },
    { x: 3.624, y: 4.647, h: 2.148, rot: 165, flipH: false, color: C.green, kind: 'calendar',
      ix: 4.351, iy: 5.449, iw: 0.286, ih: 0.286, tx: 4.255, ty: 5.921, align: 'left' },
    { x: 5.799, y: 4.636, h: 2.158, rot: 195, flipH: true, color: C.pink, kind: 'clipboard',
      ix: 8.894, iy: 5.42, iw: 0.229, ih: 0.317, tx: 8.028, ty: 5.92, align: 'right' },
  ];
  petals.forEach(p => {
    blob(s, 'ellipse', p.x, p.y, 3.964, p.h, p.color, { rotate: p.rot, flipH: p.flipH });
    icon(s, p.kind, p.ix, p.iy, p.iw, p.ih, C.white, p.color);
    txt(s, 'Subtitle', { x: p.tx, y: p.ty, w: 1.189, h: 0.374,
      fontFace: FONT_HEAD, fontSize: 18, color: C.white, align: p.align, lineSpacingMultiple: 0.9 });
    txt(s, 'Your Value', { x: p.tx, y: p.ty + 0.35, w: 1.189, h: 0.283,
      color: C.white, align: p.align, lineSpacingMultiple: 0.9 });
  });
  blob(s, 'ellipse', 5.876, 4.599, 1.671, 1.695, C.white);
  icon(s, 'briefcase', 6.524, 4.994, 0.32, 0.286, C.black);
  txt(s, 'Subtitle', { x: 6.073, y: 5.422, w: 1.189, h: 0.268,
    fontFace: FONT_HEAD, fontSize: 11, color: C.black, align: 'center', lineSpacingMultiple: 0.9 });
  txt(s, 'Your Value', { x: 6.072, y: 5.644, w: 1.189, h: 0.237,
    fontSize: 9, color: C.black, align: 'center', lineSpacingMultiple: 0.9 });

  const rows = [
    { dot: 4.917, dir: 'r', tx: 1.305, tw: 3.507, y: 2.69, align: 'right', color: C.green, label: 'Data Number One' },
    { dot: 8.133, dir: 'l', tx: 8.521, tw: 3.507, y: 2.69, align: 'left', color: C.purple, label: 'Data Number Two' },
    { dot: 8.292, dir: 'l', tx: 8.68, tw: 3.412, y: 3.79, align: 'left', color: C.pink, label: 'Data Number Three' },
    { dot: 4.854, dir: 'r', tx: 1.242, tw: 3.495, y: 3.812, align: 'right', color: C.black, label: 'Data Number Four' },
  ];
  rows.forEach(r => {
    arrowDot(s, r.dot, r.y + 0.075, r.color, r.dir);
    txt(s, r.label, { x: r.tx, y: r.y, w: r.tw, h: 0.37,
      fontSize: 16, bold: true, color: r.color, align: r.align });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur', {
      x: r.tx, y: r.y + 0.378, w: r.tw, h: 0.334, align: r.align, lineSpacingMultiple: 1.3 });
  });
}

// ============================================================ slide 11 =====
// Diagonal chain of three circles with bullets, copy block on the left.
function slide11(pres) {
  const s = pres.addSlide();
  const rows = [
    { dot: 8.838, y: 1.155, tw: 2.48, color: C.purple, label: 'Data Number One' },
    { dot: 9.76, y: 3.244, tw: 2.413, color: C.yellow, label: 'Data Number Two' },
    { dot: 8.686, y: 5.462, tw: 2.413, color: C.green, label: 'Data Number Three' },
  ];
  rows.forEach(r => {
    arrowDot(s, r.dot, r.y + 0.075, r.color);
    txt(s, r.label, { x: r.dot + 0.388, y: r.y, w: r.tw, h: 0.37,
      fontSize: 16, bold: true, color: r.color });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur', {
      x: r.dot + 0.388, y: r.y + 0.378, w: r.tw, h: 0.596, lineSpacingMultiple: 1.3 });
  });

  const bubbles = [
    { x: 5.875, y: 0.856, color: C.purple, value: '67K', tx: 6.196, ty: 1.636 },
    { x: 6.992, y: 2.467, color: C.yellow, value: '29K', tx: 7.314, ty: 3.248 },
    { x: 5.875, y: 4.164, color: C.green, value: '85K', tx: 6.196, ty: 5.003 },
  ];
  bubbles.forEach(b => {
    circle(s, b.x, b.y, 2.48, b.color);
    txt(s, b.value, { x: b.tx, y: b.ty, w: 1.838, h: 0.919,
      fontFace: FONT_HEAD, fontSize: 54, bold: true, color: C.white,
      align: 'center', lineSpacingMultiple: 0.9 });
  });

  heading(s, { x: 0.773, y: 2.465, w: 4.619, subY: 3.237, subColor: C.gray65 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
        'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud ' +
        'exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ',
    { x: 0.773, y: 3.844, w: 4.619, h: 1.616, color: C.gray7F,
      lineSpacingMultiple: 1.5, margin: [7.2, 7.2, 3.6, 3.6] });
}

// ============================================================ slide 12 =====
// Five-petal flower venn with a white hub, five bullet rows on the right.
function slide12(pres) {
  const s = pres.addSlide();
  const petals = [
    { x: 0.76, y: 2.281, color: C.blue, tx: 0.98, ty: 3.076, align: 'left' },
    { x: 2.747, y: 0.895, color: C.purple, tx: 3.243, ty: 1.334, align: 'left' },
    { x: 4.74, y: 2.281, color: C.pink, tx: 5.538, ty: 3.076, align: 'right' },
    { x: 1.598, y: 4.519, color: C.yellow, tx: 2.002, ty: 5.522, align: 'center' },
    { x: 3.952, y: 4.519, color: C.green, tx: 4.497, ty: 5.522, align: 'center' },
  ];
  petals.forEach(p => {
    circle(s, p.x, p.y, 2.181, p.color);
    txt(s, 'Subtitle', { x: p.tx, y: p.ty, w: 1.189, h: 0.374,
      fontFace: FONT_HEAD, fontSize: 18, color: C.white, align: p.align, lineSpacingMultiple: 0.9 });
    txt(s, 'Your Value', { x: p.tx, y: p.ty + 0.309, w: 1.189, h: 0.283,
      color: C.white, align: p.align, lineSpacingMultiple: 0.9 });
  });
  s.addShape('ellipse', { x: 2.191, y: 2.249, w: 3.293, h: 3.293,
    fill: { color: C.white, transparency: 15 }, line: NOLINE });
  txt(s, 'Venn Diagram', { x: 2.77, y: 3.36, w: 2.136, h: 0.949,
    fontFace: FONT_HEAD, fontSize: 28, bold: true, color: C.dark,
    align: 'center', lineSpacingMultiple: 0.9 });
  icon(s, 'box', 3.698, 2.499, 0.286, 0.254, C.purple);
  icon(s, 'clipboard', 4.466, 4.821, 0.229, 0.317, C.green);
  icon(s, 'briefcase', 4.931, 3.36, 0.32, 0.286, C.pink);
  icon(s, 'envelope', 2.447, 3.411, 0.297, 0.184, C.blue);
  icon(s, 'calendar', 2.942, 4.819, 0.286, 0.286, C.yellow);

  const rows = [
    { y: 1.444, color: C.purple, label: 'Data Number One' },
    { y: 2.442, color: C.pink, label: 'Data Number Two' },
    { y: 3.441, color: C.green, label: 'Data Number Three' },
    { y: 4.44, color: C.yellow, label: 'Data Number Four' },
    { y: 5.45, color: C.blue, label: 'Data Number Four' },
  ];
  rows.forEach(r => {
    arrowDot(s, 7.931, r.y + 0.074, r.color);
    txt(s, r.label, { x: 8.319, y: r.y, w: 3.499, h: 0.37, fontSize: 16, bold: true, color: r.color });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur', {
      x: 8.319, y: r.y + 0.377, w: 3.499, h: 0.341, lineSpacingMultiple: 1.3 });
  });
}

// ============================================================ slide 13 =====
// Bubble chart: one big circle plus four satellites, captions all round.
function slide13(pres) {
  const s = pres.addSlide();
  blob(s, 'ellipse', 4.806, 2.091, 4.016, 4.016, C.purple);
  txt(s, '90%', { x: 5.702, y: 3.622, w: 2.452, h: 0.909,
    fontFace: FONT_HEAD, fontSize: 48, bold: true, color: C.white, align: 'center' });

  const caps = [
    { x: 4.286, y: 0.579, align: 'right', color: C.purple, value: '$270.000', label: 'Value One' },
    { x: 0.76, y: 2.218, align: 'right', color: C.yellow, value: '$562.000', label: 'Value Two' },
    { x: 2.361, y: 5.698, align: 'right', color: C.blue, value: '$197.000', label: 'Value Five' },
    { x: 9.854, y: 5.392, align: 'left', color: C.green, value: '$615.000', label: 'Value Four' },
    { x: 9.55, y: 1.662, align: 'left', color: C.pink, value: '$287.000', label: 'Value Three' },
  ];
  caps.forEach(c => valueBlock(s, Object.assign({ w: 2.236, labelColor: C.black,
    body: 'Lorem ipsum dolor sit' }, c)));

  const bubbles = [
    { x: 3.324, y: 2.235, d: 2.546, color: C.yellow, v: '75%', tx: 3.371, ty: 3.087, tw: 2.452, th: 0.841, ts: 44 },
    { x: 6.978, y: 1.307, d: 2.204, color: C.pink, v: '65%', tx: 6.854, ty: 1.964, tw: 2.452, th: 0.774, ts: 40 },
    { x: 7.595, y: 4.296, d: 1.954, color: C.green, v: '30%', tx: 7.739, ty: 4.919, tw: 1.667, th: 0.707, ts: 36 },
    { x: 5.17, y: 5.207, d: 1.684, color: C.blue, v: '15%', tx: 5.398, ty: 5.743, tw: 1.228, th: 0.64, ts: 32 },
  ];
  bubbles.forEach(b => {
    circle(s, b.x, b.y, b.d, b.color);
    txt(s, b.v, { x: b.tx, y: b.ty, w: b.tw, h: b.th,
      fontFace: FONT_HEAD, fontSize: b.ts, bold: true, color: C.white, align: 'center' });
  });
}

// ============================================================ slide 14 =====
// Four-circle venn with a "Total Income" hub, three captions on the left.
function slide14(pres) {
  const s = pres.addSlide();
  const quads = [
    { x: 8.233, y: 2.558, color: C.green, tx: 10.547, ty: 5.046, v: '5,3M', l: 'Income 03' },
    { x: 8.233, y: 0.705, color: C.pink, tx: 10.563, ty: 1.535, v: '4,2M', l: 'Income 02' },
    { x: 6.381, y: 0.705, color: C.purple, tx: 6.909, ty: 1.573, v: '3,1M', l: 'Income 01' },
    { x: 6.381, y: 2.558, color: C.yellow, tx: 6.844, ty: 5.046, v: '2,8M', l: 'Income 04' },
  ];
  quads.forEach(q => circle(s, q.x, q.y, 4.244, q.color));
  quads.forEach(q => {
    txt(s, q.v, { x: q.tx, y: q.ty, w: 1.386, h: 0.64,
      fontFace: FONT_HEAD, fontSize: 32, bold: true, color: C.white, align: 'center' });
    txt(s, q.l, { x: q.tx, y: q.ty + 0.623, w: 1.386, h: 0.303, color: C.white, align: 'center' });
  });
  hub(s, 7.854, 2.179, 3.15);
  txt(s, '12,3M', { x: 8.594, y: 3.685, w: 1.671, h: 0.64,
    fontFace: FONT_HEAD, fontSize: 32, bold: true, color: C.black, align: 'center' });
  txt(s, 'Total Income', { x: 8.594, y: 4.34, w: 1.671, h: 0.37,
    fontFace: FONT_HEAD, fontSize: 16, bold: true, color: C.black, align: 'center' });
  icon(s, 'box', 9.05, 2.796, 0.758, 0.673, C.black);

  const rows = [
    { y: 2.706, color: C.purple, label: 'Data Income 01', kind: 'clipboard', ix: 1.086, iy: 2.879, iw: 0.229, ih: 0.317 },
    { y: 3.958, color: C.pink, label: 'Data Income 02', kind: 'calendar', ix: 1.057, iy: 4.15, iw: 0.286, ih: 0.286 },
    { y: 5.21, color: C.green, label: 'Data Income 03', kind: 'box', ix: 1.057, iy: 5.418, iw: 0.286, ih: 0.254 },
  ];
  rows.forEach(r => {
    s.addShape('ellipse', { x: 0.866, y: r.y, w: 0.669, h: 0.669,
      fill: { color: r.color }, line: NOLINE, shadow: SH_BADGE });
    icon(s, r.kind, r.ix, r.iy, r.iw, r.ih, C.white, r.color);
    txt(s, r.label, { x: 1.728, y: r.y, w: 3.544, h: 0.37, fontSize: 16, bold: true, color: r.color });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', {
      x: 1.728, y: r.y + 0.39, w: 3.544, h: 0.604, lineSpacingMultiple: 1.3 });
  });
  heading(s, { x: 0.751, y: 0.768 });
}

// ============================================================ slide 15 =====
// Four teardrops meeting at the centre, two captions per side.
function slide15(pres) {
  const s = pres.addSlide();
  const drops = [
    { x: 4.379, y: 2.228, rot: 0, flipH: true, color: C.purple, v: '3,7M', vx: 4.258, vy: 3.2,
      l: 'Income 01', lx: 4.854, ly: 2.852, lw: 1.274 },
    { x: 6.277, y: 2.228, rot: 0, flipH: false, color: C.pink, v: '3,1M', vx: 6.667, vy: 3.2,
      l: 'Income 02', lx: 7.268, ly: 2.863, lw: 1.32 },
    { x: 6.277, y: 4.101, rot: 180, flipH: true, color: C.green, v: '3,5M', vx: 6.667, vy: 5.644,
      l: 'Income 04', lx: 7.262, ly: 5.331, lw: 1.32 },
    { x: 4.379, y: 4.101, rot: 180, flipH: false, color: C.yellow, v: '3,4M', vx: 4.258, vy: 5.644,
      l: 'Income 03', lx: 4.848, ly: 5.32, lw: 1.274 },
  ];
  drops.forEach(d => {
    blob(s, 'teardrop', d.x, d.y, 2.677, 2.674, d.color, { rotate: d.rot, flipH: d.flipH });
    txt(s, d.l, { x: d.lx, y: d.ly, w: d.lw, h: 0.337, fontSize: 14, color: C.white, align: 'center' });
    txt(s, d.v, { x: d.vx, y: d.vy, w: 2.452, h: 0.64,
      fontFace: FONT_HEAD, fontSize: 32, bold: true, color: C.white, align: 'center' });
  });

  const caps = [
    { bx: 3.223, by: 2.498, kind: 'envelope', ix: 3.409, iy: 2.74, iw: 0.297, ih: 0.184,
      color: C.purple, label: 'Value One', lx: 2.109, ly: 3.263, tx: 1.507, ty: 3.634, align: 'right' },
    { bx: 3.223, by: 4.723, kind: 'clipboard', ix: 3.443, iy: 4.899, iw: 0.229, ih: 0.317,
      color: C.yellow, label: 'Value Four', lx: 2.109, ly: 5.475, tx: 1.507, ty: 5.846, align: 'right' },
    { bx: 9.441, by: 2.498, kind: 'calendar', ix: 9.633, iy: 2.689, iw: 0.286, ih: 0.286,
      color: C.pink, label: 'Value Two', lx: 9.374, ly: 3.257, tx: 9.374, ty: 3.627, align: 'left' },
    { bx: 9.441, by: 4.723, kind: 'briefcase', ix: 9.616, iy: 4.914, iw: 0.32, ih: 0.286,
      color: C.green, label: 'Value Three', lx: 9.374, ly: 5.469, tx: 9.374, ty: 5.839, align: 'left' },
  ];
  caps.forEach(c => {
    s.addShape('ellipse', { x: c.bx, y: c.by, w: 0.669, h: 0.669,
      fill: { color: c.color }, line: NOLINE, shadow: SH_BADGE });
    icon(s, c.kind, c.ix, c.iy, c.iw, c.ih, C.white, c.color);
    txt(s, c.label, { x: c.lx, y: c.ly, w: 1.85, h: 0.37,
      fontFace: FONT_HEAD, fontSize: 16, bold: true, color: c.color, align: c.align });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, ', {
      x: c.tx, y: c.ty, w: 2.452, h: 0.604, align: c.align, lineSpacingMultiple: 1.3 });
  });
  heading(s, { x: 4.5, y: 0.527, w: 4.619, align: 'center', subY: 1.299, subColor: C.gray65 });
}

// ============================================================ slide 16 =====
// Four interlocking hexagons with vertical "Your Text" tabs and captions.
function slide16(pres) {
  const s = pres.addSlide();
  const hexes = [
    { x: 1.138, color: C.purple, kind: 'envelope', ix: 2.588, iy: 2.871, iw: 0.394, ih: 0.243,
      tx: 1.741, v: '298k', l: 'Product A', cx: 1.648 },
    { x: 3.731, color: C.pink, kind: 'clipboard', ix: 5.226, iy: 2.694, iw: 0.303, ih: 0.421,
      tx: 4.334, v: '446k', l: 'Product B', cx: 4.236, tab: 3.215 },
    { x: 6.324, color: C.green, kind: 'briefcase', ix: 7.759, iy: 2.736, iw: 0.424, ih: 0.379,
      tx: 6.927, v: '873k', l: 'Product C', cx: 6.824, tab: 5.826 },
    { x: 8.917, color: C.yellow, kind: 'box', ix: 10.375, iy: 2.778, iw: 0.379, ih: 0.337,
      tx: 9.52, v: '287k', l: 'Product D', cx: 9.413, tab: 8.437 },
  ];
  hexes.forEach(h => {
    hexagon(s, h.x, 1.975, 3.296, 3.701, h.color);
    icon(s, h.kind, h.ix, h.iy, h.iw, h.ih, C.white, h.color);
    txt(s, h.v, { x: h.tx, y: 3.48, w: 2.088, h: 0.774,
      fontFace: FONT_HEAD, fontSize: 40, color: C.white, align: 'center' });
    txt(s, h.l, { x: h.tx, y: 4.255, w: 2.088, h: 0.374,
      fontSize: 18, color: C.white, align: 'center', lineSpacingMultiple: 0.9 });
    if (h.tab !== undefined) {
      txt(s, 'Your Text', { x: h.tab, y: 3.576, w: 1.685, h: 0.419, rotate: 270,
        fontFace: FONT_HEAD, fontSize: 16, bold: true, color: C.white,
        align: 'center', valign: 'bottom' });
    }
    txt(s, 'Your Text', { x: h.cx, y: 5.868, w: 2.275, h: 0.37,
      fontSize: 16, bold: true, color: h.color, align: 'center' });
    txt(s, 'Lorem ipsum dolor sit amet, consectetur', {
      x: h.cx, y: 6.246, w: 2.275, h: 0.596, align: 'center', lineSpacingMultiple: 1.3 });
  });
  heading(s, { x: 4.125, y: 0.613, align: 'center' });
}

// ------------------------------------------------------------------ build --
function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'W16x9', width: 13.333, height: 7.5 });
  pres.layout = 'W16x9';
  pres.author = 'pptxgenjs';
  pres.title = 'Venn Diagram - Infographic Presentation Template';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
   slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16]
    .forEach(fn => fn(pres));

  const out = path.join(__dirname, '041d34f4-5952-4f16-a386-8c255f3543b1_grok_final.pptx');
  return pres.writeFile({ fileName: out }).then(() => console.log('wrote ' + out));
}

build();
