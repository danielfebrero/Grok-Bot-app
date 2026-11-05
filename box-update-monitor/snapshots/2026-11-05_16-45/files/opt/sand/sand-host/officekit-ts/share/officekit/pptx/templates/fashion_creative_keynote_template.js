/**
 * "Perm. Creative" — 28-slide keynote template, rebuilt with pptxgenjs.
 *
 * The source deck carries no embedded media: every photo position is an empty
 * picture frame that renders as bare backdrop. Those slots are reproduced here as
 * backdrop-toned rectangles (`photoSlot`) drawn before the coloured panels, so no
 * binary image data is needed anywhere in this script.
 *
 *   node 14ab0e18-5e3a-49ca-85ac-606b6d495c75_grok_final.js
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const C = {
  bg: 'D8EBE4',      // page background mint
  yellow: 'FED049',
  teal: '007580',
  amber: 'FFC000',   // theme accent4
  orange: 'ED7D31',  // theme accent2
  gray: 'A5A5A5',    // theme accent3
  white: 'FFFFFF',
  ink: '222A35',
  black: '161616',
  navy: '31538F',
};

const F = {
  reg: 'Roboto',
  med: 'Roboto Medium',
  black: 'Roboto Black',
  light: 'Roboto Light',
  ssp: 'Source Sans Pro',
  cormorant: 'Cormorant Garamond',
  poppins: 'Poppins',
  raleway: 'Raleway',
};

/* --------------------------------------------------------- shared copy text */

const LOREP_LONG =
  'Lorep  ipsum duis aute irure dolor in kauselih oilusioisduil reprehenderitisi ' +
  'voluptates esse cill inure dolorlorudi siot amet. Duis aute irusitaseiad dolorin reprehe karisla mari';
const LOREP_MED =
  'Lorep  ipsum duis aute irure dolor in kauselih oilue eprehe deriti vols esse cill inure ' +
  'dolorlaboru sit amt. Duis autelo irusitakus reprehenri Voluptate lorem kuisais.';
const LOREP_SHORT =
  'Lorep  ipsum duis aute irure dolor in kauselih oilusioisduil reprehenderitisi voluptates';
const LOREP_TINY = 'Lorep  ipsum duis aute irure dolor in kauselih';
const LOREP_CARD =
  'Lorep  ipsum duis aute irure dolor in kauselih oilusioisduil reprehenderitisi ' +
  'voluptates esse cill inure dolorlorudi siot amet. Duis aute';
const LEVERAGE =
  'PLACEHOLDER' +
  'overviews. Iterative having dan approaches to have corporate foster coaborative thinking to ' +
  'further the value iterative altern proposition. ';
const SERENITY =
  'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring';
const IPSUM_SHORT =
  'Lorem Ipsum\u00a0is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been';
const IPSUM_LONG =
  "Lorem Ipsum\u00a0is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been " +
  "the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley. ";
const INTERACTIVE =
  'Interactively coordinate proactive e-commerce via process centric "outside the box" thinking. ';
const ONSEQUAE =
  'PLACEHOLDER' +
  'Tatio. Totatem Rem Isquibu';
const STEP_BODY =
  'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo libero';
const SWOT_BODY =
  'Lorem ipsum dolor sit amet, conse ctetur dolor sit amet  ctetur dolor sit ipsum dolor sit amet, ' +
  'conse ctetur dolor sit amet  ctetur';
const TAGLINE = 'with a clean & modern Keynote Presentation.';

/* ------------------------------------------------------------- tiny helpers */

/** Filled rectangle (no outline). */
function rect(s, x, y, w, h, color, extra) {
  s.addShape('rect', Object.assign({ x, y, w, h, fill: { color }, line: { type: 'none' } }, extra));
}

/** Straight rule. `pt` is stroke weight in points. */
function rule(s, x, y, w, pt, color, dash) {
  s.addShape('line', { x, y, w, h: 0, line: { color: color || C.teal, width: pt, dashType: dash || 'solid' } });
}

/** Filled polygon; `pts` are [x, y] pairs in inches relative to the shape's origin. */
function poly(s, x, y, w, h, color, pts) {
  s.addShape('custGeom', {
    x, y, w, h, fill: { color }, line: { type: 'none' },
    points: pts.map(function (p) { return { x: p[0], y: p[1] }; }).concat([{ close: true }]),
  });
}

/** Left-pointing block arrow; `head` is the head length in inches. */
function leftArrow(s, x, y, w, h, color, head) {
  const t = h * 0.505;                       // shaft thickness, as in the source deck
  const top = (h - t) / 2;
  poly(s, x, y, w, h, color,
       [[head, 0], [0, h / 2], [head, h], [head, top + t], [w, top + t], [w, top], [head, top]]);
}

/** Right-pointing chevron with a notched tail; `notch` is the point depth in inches. */
function chevron(s, x, y, w, h, color, notch) {
  poly(s, x, y, w, h, color,
       [[0, 0], [w - notch, 0], [w, h / 2], [w - notch, h], [0, h], [notch, h / 2]]);
}

/** Text box. `runs` is a string or an array of pptxgenjs text objects. */
function txt(s, runs, o) {
  s.addText(runs, Object.assign({ fontFace: F.reg, color: C.teal, align: 'left', valign: 'top' }, o));
}

/** One run of a multi-run paragraph block. */
function run(text, o) {
  return { text, options: Object.assign({ fontFace: F.reg, color: C.teal }, o) };
}

/** Empty picture frame from the source deck; call before the coloured panels. */
function photoSlot(s, x, y, w, h) {
  rect(s, x, y, w, h, C.bg);
}

/** Rotated "PERM / CREATIVE" tab that sits on the right edge of most slides. */
function sideTab(s, color, lines, x, y, w) {
  txt(s, [
    run(lines ? lines[0] : 'PERM', { fontFace: F.med, color: color, breakLine: true }),
    run(lines ? lines[1] : 'CREATIVE', { fontFace: F.med, color: color }),
  ], { x: x === undefined ? 11.567 : x, y: y === undefined ? 1.256 : y, w: w === undefined ? 1.783 : w,
       h: 0.438, fontSize: 10, rotate: 90 });
}

/**
 * "Perm. / Creative / with a clean & modern Keynote Presentation." lock-up.
 * (x, y) is the top-left of the "Perm." line; the rest is offset from it.
 */
function brandBlock(s, x, y) {
  txt(s, 'Perm.', { x, y, w: 1.494, h: 0.535, fontSize: 28, fontFace: F.black,
                    lineSpacingMultiple: 0.8, margin: 0 });
  txt(s, 'Creative', { x: x - 0.054, y: y + 0.018, w: 3.393, h: 1.107, fontSize: 28,
                       valign: 'middle', lineSpacingMultiple: 1, margin: 4 });
  txt(s, TAGLINE, { x: x - 0.054, y: y + 0.831, w: 2.877, h: 1.107, fontSize: 14, margin: 4 });
}

/* ------------------------------------------------------------------- icons */
/*
 * Each icon is a list of primitives laid out in a 0..1 box:
 *   s     pptxgenjs shape name          r    rectRadius, as a fraction of the icon width
 *   x y w h   position/size, 0..1       rot  rotation in degrees
 *   o     outline only (no fill)        cut  paint in the tile colour (knock-out detail)
 *   t     glyph drawn as text instead of a shape
 */
const ICONS = {
  bars:      [{ s: 'rect', x: 0.00, y: 0.39, w: 0.26, h: 0.61 },
              { s: 'rect', x: 0.37, y: 0.16, w: 0.26, h: 0.84 },
              { s: 'rect', x: 0.74, y: 0.00, w: 0.26, h: 1.00 }],
  people:    [{ s: 'ellipse', x: 0.00, y: 0.10, w: 0.26, h: 0.33 },
              { s: 'ellipse', x: 0.37, y: 0.00, w: 0.28, h: 0.36 },
              { s: 'ellipse', x: 0.74, y: 0.10, w: 0.26, h: 0.33 },
              { s: 'round2SameRect', x: 0.00, y: 0.50, w: 0.28, h: 0.50, r: 0.12 },
              { s: 'round2SameRect', x: 0.35, y: 0.42, w: 0.31, h: 0.58, r: 0.14 },
              { s: 'round2SameRect', x: 0.72, y: 0.50, w: 0.28, h: 0.50, r: 0.12 }],
  list:      [{ s: 'rect', x: 0.00, y: 0.00, w: 0.20, h: 0.22 },
              { s: 'rect', x: 0.00, y: 0.39, w: 0.20, h: 0.22 },
              { s: 'rect', x: 0.00, y: 0.78, w: 0.20, h: 0.22 },
              { s: 'rect', x: 0.32, y: 0.02, w: 0.68, h: 0.18 },
              { s: 'rect', x: 0.32, y: 0.41, w: 0.68, h: 0.18 },
              { s: 'rect', x: 0.32, y: 0.80, w: 0.68, h: 0.18 }],
  moneybag:  [{ s: 'triangle', x: 0.30, y: 0.00, w: 0.40, h: 0.24, o: true, rot: 180 },
              { s: 'roundRect', x: 0.08, y: 0.22, w: 0.84, h: 0.78, r: 0.30, o: true },
              { t: '$', x: 0.30, y: 0.34, w: 0.40, h: 0.46 }],
  key:       [{ s: 'donut', x: 0.44, y: 0.00, w: 0.56, h: 0.56 },
              { s: 'rect', x: 0.02, y: 0.60, w: 0.62, h: 0.15, rot: 315 },
              { s: 'rect', x: 0.14, y: 0.78, w: 0.16, h: 0.14 },
              { s: 'rect', x: 0.31, y: 0.61, w: 0.16, h: 0.14 }],
  woman:     [{ s: 'round2SameRect', x: 0.16, y: 0.00, w: 0.68, h: 0.58, r: 0.34 },
              { s: 'rect', x: 0.40, y: 0.50, w: 0.20, h: 0.18 },
              { s: 'trapezoid', x: 0.02, y: 0.62, w: 0.96, h: 0.38 },
              { s: 'triangle', x: 0.32, y: 0.62, w: 0.36, h: 0.34, cut: true, rot: 180 },
              { s: 'triangle', x: 0.45, y: 0.62, w: 0.10, h: 0.36, rot: 180 }],
  cart:      [{ s: 'rect', x: 0.00, y: 0.00, w: 1.00, h: 0.60 },
              { s: 'rect', x: 0.00, y: 0.64, w: 1.00, h: 0.14 },
              { s: 'rect', x: 0.40, y: 0.78, w: 0.20, h: 0.12 },
              { s: 'rect', x: 0.22, y: 0.90, w: 0.56, h: 0.10 },
              { s: 'trapezoid', x: 0.16, y: 0.16, w: 0.50, h: 0.22, cut: true, o: true, rot: 180 },
              { s: 'rect', x: 0.13, y: 0.10, w: 0.04, h: 0.30, cut: true, rot: 10 },
              { s: 'rect', x: 0.20, y: 0.36, w: 0.36, h: 0.04, cut: true },
              { s: 'ellipse', x: 0.22, y: 0.42, w: 0.10, h: 0.10, cut: true },
              { s: 'ellipse', x: 0.44, y: 0.42, w: 0.10, h: 0.10, cut: true },
              { s: 'triangle', x: 0.70, y: 0.26, w: 0.15, h: 0.22, cut: true, rot: 200 }],
  shopbag:   [{ s: 'donut', x: 0.26, y: 0.00, w: 0.48, h: 0.48 },
              { s: 'round2SameRect', x: 0.00, y: 0.24, w: 1.00, h: 0.76, r: 0.16, rot: 180 },
              { s: 'ellipse', x: 0.22, y: 0.40, w: 0.52, h: 0.48, cut: true },
              { t: '$', x: 0.30, y: 0.44, w: 0.36, h: 0.40 }],
  phone:     [{ s: 'roundRect', x: 0.20, y: 0.00, w: 0.60, h: 1.00, r: 0.16, o: true },
              { s: 'rect', x: 0.40, y: 0.86, w: 0.20, h: 0.05 }],
  monitor:   [{ s: 'roundRect', x: 0.00, y: 0.00, w: 1.00, h: 0.68, r: 0.08, o: true },
              { s: 'rect', x: 0.42, y: 0.70, w: 0.16, h: 0.16 },
              { s: 'rect', x: 0.22, y: 0.88, w: 0.56, h: 0.10 }],
  bell:      [{ s: 'ellipse', x: 0.40, y: 0.00, w: 0.20, h: 0.16 },
              { s: 'round2SameRect', x: 0.10, y: 0.06, w: 0.80, h: 0.64, r: 0.40 },
              { s: 'roundRect', x: 0.00, y: 0.60, w: 1.00, h: 0.20, r: 0.08 },
              { s: 'ellipse', x: 0.32, y: 0.72, w: 0.36, h: 0.28 }],
  chat:      [{ s: 'roundRect', x: 0.00, y: 0.00, w: 1.00, h: 0.72, r: 0.1 },
              { s: 'triangle', x: 0.06, y: 0.60, w: 0.26, h: 0.40, rot: 205 },
              { s: 'ellipse', x: 0.16, y: 0.30, w: 0.15, h: 0.16, cut: true },
              { s: 'ellipse', x: 0.42, y: 0.30, w: 0.15, h: 0.16, cut: true },
              { s: 'ellipse', x: 0.68, y: 0.30, w: 0.15, h: 0.16, cut: true }],
  globe:     [{ s: 'ellipse', x: 0.00, y: 0.00, w: 1.00, h: 1.00 },
              { s: 'ellipse', x: 0.26, y: -0.02, w: 0.48, h: 1.04, o: true, cut: true },
              { s: 'rect', x: 0.00, y: 0.20, w: 1.00, h: 0.05, cut: true },
              { s: 'rect', x: 0.00, y: 0.47, w: 1.00, h: 0.05, cut: true },
              { s: 'rect', x: 0.00, y: 0.74, w: 1.00, h: 0.05, cut: true },
              { s: 'rect', x: 0.47, y: 0.00, w: 0.05, h: 1.00, cut: true }],
  magnifier: [{ s: 'ellipse', x: 0.00, y: 0.00, w: 0.76, h: 0.76, o: true },
              { s: 'rect', x: 0.62, y: 0.66, w: 0.38, h: 0.13, rot: 45 }],
  star:      [{ s: 'star5', x: 0.00, y: 0.00, w: 1.00, h: 1.00, o: true }],
  plane:     [{ s: 'triangle', x: 0.00, y: 0.05, w: 1.00, h: 0.90, o: true, rot: 105 }],
  crew:      [{ s: 'ellipse', x: 0.02, y: 0.00, w: 0.24, h: 0.24 },
              { s: 'rect', x: 0.05, y: 0.30, w: 0.18, h: 0.70 },
              { s: 'ellipse', x: 0.38, y: 0.00, w: 0.24, h: 0.24 },
              { s: 'rect', x: 0.41, y: 0.30, w: 0.18, h: 0.70 },
              { s: 'ellipse', x: 0.74, y: 0.00, w: 0.24, h: 0.24 },
              { s: 'rect', x: 0.77, y: 0.30, w: 0.18, h: 0.70 }],
};

/** Draw `name` inside the box (x, y, w, h); `tile` is the colour behind it. */
function icon(s, name, x, y, w, h, color, tile) {
  ICONS[name].forEach(function (p) {
    const c = p.cut ? (tile || C.bg) : color;
    const box = { x: x + p.x * w, y: y + p.y * h, w: p.w * w, h: p.h * h };
    if (p.t) {
      txt(s, p.t, Object.assign({}, box, { color: c, fontSize: Math.round(box.h * 72),
                                           bold: true, align: 'center', valign: 'middle', margin: 0 }));
      return;
    }
    s.addShape(p.s, Object.assign(box, {
      fill: p.o ? { type: 'none' } : { color: c },
      line: p.o ? { color: c, width: 1.5 } : { type: 'none' },
      rotate: p.rot || 0,
      rectRadius: p.r ? p.r * w : undefined,
    }));
  });
}

/* ----------------------------------------------------------- slide builders */

const build = [];

// 1 — Welcome cover
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  photoSlot(s, 0, 0, 7.143, 7.5);
  rect(s, 5.042, 1.035, 7.472, 5.431, C.yellow);
  txt(s, 'Welcome', { x: 6.16, y: 2.806, w: 5.738, h: 1.582, fontSize: 88, bold: true, italic: true });
  s.addShape('line', { x: 10.762, y: 1.02, w: 0, h: 2.047, line: { color: C.teal, width: 3 } });
  brandBlock(s, 6.36, 4.371);
});

// 2 — CREATIVE divider
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  photoSlot(s, 8.625, 0, 4.708, 7.5);
  rect(s, 3.339, 0, 5.286, 7.5, C.yellow);
  sideTab(s, C.bg);
  brandBlock(s, 0.656, 0.588);
  rect(s, 0.898, 4.629, 7.398, 2.0, C.teal);
  txt(s, 'CREATIVE', { x: 1.819, y: 4.905, w: 5.71, h: 1.447, fontSize: 80, fontFace: F.med,
                       color: C.yellow, align: 'center' });
});

// 3 — "#1" chapter opener
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  photoSlot(s, 2.412, 0, 5.251, 7.5);
  rect(s, 0, 0, 2.412, 7.5, C.yellow);
  txt(s, LEVERAGE, { x: 8.143, y: 3.959, w: 3.395, h: 2.454, fontSize: 10.5, lineSpacingMultiple: 1.5 });
  txt(s, '#1', { x: 0.444, y: 5.807, w: 1.968, h: 1.212, fontSize: 66, bold: true, fontFace: F.med });
  sideTab(s, C.bg, ['PERM', 'FASHION'], 11.686, 1.137, 1.545);
  sideTab(s, C.teal);
  txt(s, 'Perm.', { x: 0.601, y: 0.588, w: 1.494, h: 0.535, fontSize: 28, fontFace: F.black,
                    lineSpacingMultiple: 0.8, margin: 0 });
  txt(s, 'Creative', { x: 0.547, y: 0.726, w: 3.393, h: 0.782, fontSize: 28, fontFace: F.med,
                       valign: 'middle', margin: 4 });
  brandBlock(s, 8.309, 2.018);
});

// 4 — Framed CREATIVE title
build.push(function (s) {
  s.addShape('rect', { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: C.bg }, line: { color: C.navy, width: 1 } });
  rect(s, 0, 5.731, 1.635, 1.769, C.yellow);
  s.addShape('rect', { x: 0.602, y: 0.601, w: 12.089, h: 6.255, fill: { type: 'none' },
                       line: { color: C.teal, width: 1 } });
  txt(s, 'CREATIVE', { x: 2.193, y: 2.731, w: 8.948, h: 1.717, fontSize: 96, fontFace: F.raleway,
                       align: 'center' });
  txt(s, [run('Modern Presentation Template', { fontFace: F.light }), run('.', { fontFace: F.cormorant })],
      { x: 4.488, y: 5.691, w: 4.358, h: 0.422, fontSize: 14, align: 'center', lineSpacingMultiple: 1.5 });
  rule(s, 6.387, 5.595, 0.559, 0.75);
  txt(s, [run('2019', { fontFace: F.light }), run('                     ', { color: C.white, fontFace: F.cormorant }),
          run('Since', { fontFace: F.light })],
      { x: 10.425, y: 1.252, w: 1.508, h: 0.269, fontSize: 10 });
  rule(s, 10.877, 1.403, 0.45, 0.75);
  txt(s, '01/ 02', { x: 7.137, y: 1.054, w: 0.74, h: 0.738, fontSize: 11, fontFace: F.light,
                     align: 'center', lineSpacingMultiple: 2.5 });
  rule(s, 6.151, 1.405, 0.887, 0.75);
  txt(s, '2021 Design For You', { x: 1.435, y: 1.168, w: 1.355, h: 0.252, fontSize: 9,
                                  fontFace: F.light, align: 'center' });
});

// 5 — Our Styles
build.push(function (s) {
  s.addShape('rect', { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: C.bg }, line: { color: C.navy, width: 1 } });
  photoSlot(s, 0.602, 0, 5.856, 3.826);
  rect(s, 9.938, 0, 3.395, 7.5, C.yellow);
  txt(s, [run('MULTI FASHION', { breakLine: true }), run('PRESENTATION')],
      { x: 2.088, y: 5.865, w: 3.101, h: 0.572, fontSize: 14, align: 'center' });
  txt(s, 'CREATIVE', { x: 1.084, y: 4.319, w: 5.583, h: 1.212, fontSize: 66, bold: true,
                       fontFace: F.med, align: 'center' });
  txt(s, LEVERAGE, { x: 7.861, y: 3.723, w: 3.395, h: 2.454, fontSize: 10.5, lineSpacingMultiple: 1.5 });
  txt(s, [run('Our', { bold: true, fontFace: F.med, breakLine: true }), run('Styles', { fontFace: F.med })],
      { x: 7.861, y: 1.393, w: 3.773, h: 1.919, fontSize: 54 });
  sideTab(s, C.teal);
});

// 6 — Founder
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  photoSlot(s, 0.872, 3.582, 8.649, 3.918);
  txt(s, 'FOUNDER', { x: 0.706, y: 1.542, w: 5.812, h: 1.01, fontSize: 54, fontFace: F.med });
  txt(s, [run('Lorep  ipsum duis aute irure dolor in ', { breakLine: true }),
          run('kauselih oilue eprehederiti vols esse cill inure dolorlaboru sit amet. Duis autelo ', { breakLine: true }),
          run('irusitakus reprehenderi Voluptate ', { breakLine: true }),
          run('lorem kuisais.')],
      { x: 9.894, y: 4.184, w: 2.568, h: 1.662, fontSize: 10.5, lineSpacingMultiple: 1.5 });
  txt(s, [run('Lorep  ipsu', { bold: true, italic: true }), run('.')],
      { x: 9.894, y: 3.472, w: 2.797, h: 0.64, fontSize: 32 });
  txt(s, [run('Modern Presentation Template'), run('.', { color: C.black })],
      { x: 0.79, y: 2.552, w: 4.358, h: 0.422, fontSize: 14, lineSpacingMultiple: 1.5 });
  txt(s, '01/ 02', { x: 11.811, y: 1.636, w: 0.738, h: 0.816, fontSize: 12, align: 'center',
                     lineSpacingMultiple: 2.5 });
  rule(s, 10.745, 2.006, 0.887, 0.75);
  txt(s, [run('\u201CIt is not from the benevolence of the butcher, the brewer, or the baker that we ' +
              'expect our dinner, but from their regard to their own interest.'), run('\u201D', { color: C.black })],
      { x: 7.661, y: 1.35, w: 2.905, h: 1.132, fontSize: 10.5, lineSpacingMultiple: 1.5 });
});

// 7 — Perm. Gallery
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  photoSlot(s, 5.368, 1.496, 2.583, 4.508);
  photoSlot(s, 9.104, 1.488, 2.568, 4.531);
  rect(s, 0, 0, 4.194, 7.5, C.yellow);
  txt(s, [run('Perm.', { bold: true, fontFace: F.med, breakLine: true }), run('Gallery', { fontFace: F.med })],
      { x: 0.512, y: 0.601, w: 3.58, h: 1.919, fontSize: 54 });
  sideTab(s, C.teal);
  txt(s, LOREP_MED, { x: 0.626, y: 3.041, w: 2.245, h: 1.662, fontSize: 10.5, lineSpacingMultiple: 1.5 });
});

// 8 — Two-column story
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  photoSlot(s, 0, 0.602, 5.073, 3.727);
  photoSlot(s, 5.688, 2.514, 4.842, 3.5);
  rect(s, 3.156, 0, 1.364, 2.771, C.yellow);
  rect(s, 8.202, 3.917, 5.131, 3.6, C.yellow);
  [4.438, 5.766].forEach(function (y) {
    txt(s, [run(LOREP_LONG), run('.', { fontSize: 8 })],
        { x: 0.547, y, w: 3.446, h: 1.132, fontSize: 10.5, lineSpacingMultiple: 1.5 });
  });
  txt(s, [run('Text ', { fontFace: F.cormorant }), run('Tittle')],
      { x: 10.53, y: 5.047, w: 1.168, h: 0.418, fontSize: 14, lineSpacingMultiple: 1.5 });
  txt(s, [run('Lorep  ipsum duis aute ', { breakLine: true }), run('likauselih oilue eprehe')],
      { x: 10.53, y: 5.411, w: 1.804, h: 0.604, fontSize: 10.5, lineSpacingMultiple: 1.5 });
  brandBlock(s, 8.179, 0.655);
  sideTab(s, C.teal);
});

// 9 — Gallery with side notes
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  photoSlot(s, 5.064, 3.405, 4.634, 2.872);
  photoSlot(s, 8.72, 0, 4.634, 3.019);
  rect(s, 5.606, 4.826, 1.771, 2.674, C.yellow);
  [0.651, 1.812].forEach(function (y) {
    txt(s, LOREP_LONG + '.', { x: 4.944, y, w: 3.444, h: 1.132, fontSize: 10.5, lineSpacingMultiple: 1.5 });
  });
  txt(s, '01/ 02', { x: 11.543, y: 3.625, w: 0.738, h: 0.818, fontSize: 12, fontFace: F.poppins,
                     align: 'center', lineSpacingMultiple: 2.5 });
  rule(s, 10.556, 3.997, 0.887, 0.75);
  txt(s, LOREP_MED, { x: 10.446, y: 4.252, w: 2.245, h: 1.662, fontSize: 10.5, lineSpacingMultiple: 1.5 });
  brandBlock(s, 0.606, 5.316);
  sideTab(s, C.amber);
});

// 10 — 90% stats
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  photoSlot(s, 5.078, 0.812, 3.66, 3.243);
  photoSlot(s, 9.028, 3.837, 4.306, 3.024);
  rect(s, 4.789, 0, 4.239, 7.5, C.yellow);
  brandBlock(s, 0.656, 0.601);
  [4.438, 5.766].forEach(function (y) {
    txt(s, [run(LOREP_LONG), run('.', { fontSize: 8 })],
        { x: 0.547, y, w: 3.446, h: 1.132, fontSize: 10.5, lineSpacingMultiple: 1.5 });
  });
  [[4.91, 4.423, 6.3, 4.438], [4.841, 5.766, 6.232, 5.782]].forEach(function (p) {
    txt(s, '90%', { x: p[0], y: p[1], w: 1.391, h: 0.841, fontSize: 44 });
    txt(s, SERENITY + ' which I enjoy with my whole heart. ',
        { x: p[2], y: p[3], w: 2.62, h: 1.132, fontSize: 10.5, lineSpacingMultiple: 1.5 });
  });
  sideTab(s, C.teal);
});

// 11 — Our team
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  const cols = [
    { photo: 1.344, cap: 1.448, note: 1.453, rule: 1.569, num: 1.261, ruleW: 1.724 },
    { photo: 3.951, cap: 4.144, note: 4.149, rule: 4.266, num: 3.879, ruleW: 1.726 },
    { photo: 6.559, cap: 6.844, note: 6.849, rule: 6.965, num: 6.467, ruleW: 1.724 },
    { photo: 9.167, cap: 9.542, note: 9.547, rule: 9.663, num: 9.542, ruleW: 1.724 },
  ];
  cols.forEach(function (c, i) {
    photoSlot(s, c.photo, 2.394, 2.212, 3.266);
    txt(s, 'California, 03/14/1990', { x: c.note, y: 6.302, w: 1.983, h: 0.283, fontSize: 8,
                                       italic: true, lineSpacingMultiple: 1.5 });
    txt(s, 'Your Name', { x: c.cap, y: 5.842, w: 1.582, h: 0.405, fontSize: 18 });
    rule(s, c.rule, 6.684, c.ruleW, 0.75);
    if (i < 3) txt(s, '0' + (i + 1), { x: c.num, y: 2.108, w: 0.55, h: 0.286, fontSize: 11 });
  });
  txt(s, 'OUR TEAM', { x: 1.145, y: 0.92, w: 4.09, h: 0.84, fontSize: 44, fontFace: F.med });
  txt(s, LOREP_SHORT.replace(' voluptates', ' voluptates esse cill inure dolorlorudi siot amet.'),
      { x: 5.955, y: 1.062, w: 5.432, h: 0.602, fontSize: 10.5, lineSpacingMultiple: 1.5 });
  txt(s, '04', { x: 8.902, y: 1.962, w: 0.74, h: 0.74, fontSize: 11, align: 'center',
                 lineSpacingMultiple: 2.5 });
  sideTab(s, C.teal);
});

// 12 — Model grid
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  const cards = [
    { card: 1.15, ph: 1.325, name: 1.328, nameW: 2.115, body: 1.328, label: 'Model Name 001' },
    { card: 5.064, ph: 5.24, name: 5.239, nameW: 3.007, body: 5.204, label: 'Model Name 002' },
    { card: 9.005, ph: 9.153, name: 8.975, nameW: 3.208, body: 9.005, label: 'Model Name 003' },
  ];
  cards.forEach(function (c) {
    photoSlot(s, c.ph, 2.589, 2.567, 2.459);
    rect(s, c.card, 2.45, 2.567, 2.459, C.yellow);
    txt(s, c.label, { x: c.name, y: 5.347, w: c.nameW, h: 0.404, fontSize: 18 });
    txt(s, LOREP_TINY, { x: c.body, y: 5.903, w: 2.723, h: 0.505, fontSize: 12, align: 'justify' });
  });
  txt(s, 'ALL NAME OF CREATIVE', { x: 3.281, y: 1.092, w: 7.039, h: 0.841, fontSize: 44 });
  sideTab(s, C.teal);
});

// 13 — Three description cards
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  const cards = [
    { x: 1.25, bg: C.yellow, fg: C.teal, title: 'First Description', tw: 1.857, tx: 1.969, bx: 1.969,
      by: 2.663, ico: 'bars', ix: 2.104, iy: 1.746, iw: 0.347, ih: 0.328, ic: C.teal },
    { x: 4.972, bg: C.teal, fg: C.yellow, title: 'Second Description', tw: 2.123, tx: 5.541, bx: 5.541,
      by: 2.668, ico: 'people', ix: 5.665, iy: 1.726, iw: 0.442, ih: 0.367, ic: C.yellow },
    { x: 8.694, bg: C.yellow, fg: C.teal, title: 'Third Description', tw: 2.274, tx: 9.263, bx: 9.263,
      by: 2.671, ico: 'list', ix: 9.378, iy: 1.797, iw: 0.347, ih: 0.299, ic: C.teal },
  ];
  cards.forEach(function (c) {
    photoSlot(s, c.x, 3.917, 3.722, 2.75);
    rect(s, c.x, 1.167, 3.722, 2.75, c.bg);
    icon(s, c.ico, c.ix, c.iy, c.iw, c.ih, c.ic, c.bg);
    txt(s, c.title, { x: c.tx, y: 2.322, w: c.tw, h: 0.337, fontSize: 14, color: c.fg });
    txt(s, LOREP_SHORT, { x: c.bx, y: c.by, w: 2.585, h: 0.631, fontSize: 10.5, color: c.fg,
                          align: 'justify' });
  });
});

// 14 — more models.
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  photoSlot(s, 0, 2.929, 9.153, 3.326);
  rect(s, 9.402, 2.925, 3.289, 3.333, C.amber);
  txt(s, 'A wonderful serenity has taken possession of my entire soul, wonderful serenity',
      { x: 9.776, y: 4.259, w: 2.54, h: 1.01, fontSize: 12, align: 'center', lineSpacingMultiple: 1.5 });
  icon(s, 'moneybag', 10.872, 3.572, 0.349, 0.465, C.teal, C.amber);
  txt(s, 'The Result One', { x: 5.55, y: 2.178, w: 1.671, h: 0.337, fontSize: 14 });
  txt(s, 'The Result Two', { x: 9.393, y: 2.178, w: 1.906, h: 0.337, fontSize: 14 });
  txt(s, 'more models.', { x: 1.35, y: 1.242, w: 3.663, h: 0.774, fontSize: 40, fontFace: F.med });
  rule(s, 10.502, 5.623, 1.188, 0.75);
  txt(s, LOREP_SHORT, { x: 5.55, y: 1.367, w: 2.585, h: 0.631, fontSize: 10.5, align: 'justify' });
  txt(s, LOREP_SHORT, { x: 9.393, y: 1.313, w: 2.585, h: 0.631, fontSize: 10.5, align: 'justify' });
});

// 15 — Two descriptions beside a tall photo
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  photoSlot(s, 0.602, 0.601, 4.363, 6.26);
  photoSlot(s, 5.668, 4.024, 3.056, 2.837);
  photoSlot(s, 9.637, 4.024, 3.054, 2.837);
  txt(s, 'Description 1', { x: 5.733, y: 2.675, w: 2.049, h: 0.303, fontSize: 12 });
  txt(s, 'Description 2', { x: 9.636, y: 2.675, w: 2.049, h: 0.303, fontSize: 12 });
  brandBlock(s, 5.722, 0.867);
  txt(s, LOREP_SHORT, { x: 5.669, y: 3.154, w: 2.585, h: 0.631, fontSize: 10.5, align: 'justify' });
  txt(s, LOREP_SHORT, { x: 9.636, y: 3.119, w: 2.585, h: 0.631, fontSize: 10.5, align: 'justify' });
  sideTab(s, C.teal);
});

// 16 — good instruments
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  [[2.748, 1.366], [4.997, 3.75], [7.245, 1.366], [9.491, 3.75]].forEach(function (p) {
    photoSlot(s, p[0], p[1], 2.248, 2.384);
  });
  const softShadow = { type: 'outer', blur: 41, offset: 18, angle: 90, color: '000000', opacity: 0.09 };
  [[1.671, '01', 1.842], [2.738, '02', 2.918]].forEach(function (p) {
    s.addShape('ellipse', { x: 1.603, y: p[0], w: 0.693, h: 0.693, fill: { color: C.yellow },
                            line: { type: 'none' }, shadow: softShadow });
    txt(s, p[1], { x: 1.707, y: p[2], w: 0.52, h: 0.215, fontSize: 16 });
  });
  txt(s, 'good instruments', { x: 1.594, y: 4.051, w: 2.978, h: 1.313, fontSize: 36 });
  txt(s, 'Lorep  ipsum duis aute irure dolor in kauselih oilusioisduil reprehenderitisi',
      { x: 1.599, y: 5.553, w: 2.973, h: 0.438, fontSize: 10, align: 'justify' });
  [[ 'Description 1', 5.33, 1.79, 5.33, 2.182], ['Description 2', 7.509, 4.404, 7.509, 4.825],
   ['Description 3', 9.935, 1.79, 9.935, 2.182]].forEach(function (d) {
    txt(s, d[0], { x: d[1], y: d[2], w: 1.312, h: 0.303, fontSize: 12 });
    txt(s, LOREP_CARD, { x: d[3], y: d[4], w: 1.581, h: 1.279, fontSize: 10 });
  });
  sideTab(s, C.teal);
});

// 17 — good products.
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  photoSlot(s, 0, 0.602, 3.358, 3.148);
  photoSlot(s, 3.692, 3.748, 3.366, 3.111);
  photoSlot(s, 9.953, 2.71, 3.381, 4.15);
  txt(s, '\u201CLorep  ipsum duis aute irure dolor in kauselih\u201D',
      { x: 9.956, y: 1.483, w: 2.713, h: 0.774, fontSize: 14, italic: true, lineSpacingMultiple: 1.5 });
  txt(s, 'good products.', { x: 4.042, y: 1.609, w: 4.383, h: 0.841, fontSize: 44, italic: true,
                             fontFace: F.med });
  txt(s, [run(LOREP_SHORT + ' esse ', { italic: true }), run('cill', { italic: true, fontSize: 10.5 })],
      { x: 4.042, y: 2.773, w: 5.227, h: 0.505, fontSize: 12, align: 'justify' });
  [0.602, 7.254].forEach(function (x, i) {
    txt(s, 'Lorep  ipsum duis aute irure dolor in kauselih oilusioisduilrepre hend eritisi',
        { x, y: 4.672, w: 2.014, h: 0.909, fontSize: 12, align: 'justify' });
    rect(s, i === 0 ? 0.775 : 7.407, i === 0 ? 5.956 : 5.895, 0.63, 0.061, C.teal);
  });
});

// 18 — Services
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  photoSlot(s, 8.99, 0, 4.344, 7.5);
  rect(s, 4.761, 0, 4.239, 7.5, C.yellow);
  [[1.046, 1.419], [2.525, 2.898], [4.044, 4.418], [5.496, 5.823]].forEach(function (p) {
    txt(s, 'Your Tittle Here', { x: 6.335, y: p[0], w: 1.773, h: 0.278, fontSize: 10.5, bold: true });
    txt(s, SERENITY, { x: 6.335, y: p[1], w: 2.252, h: 0.808, fontSize: 10.5 });
  });
  txt(s, 'SERVICES', { x: 0.769, y: 2.479, w: 3.446, h: 0.774, fontSize: 40, bold: true });
  txt(s, LOREP_LONG + '.', { x: 0.769, y: 3.632, w: 3.446, h: 1.132, fontSize: 10.5, lineSpacingMultiple: 1.5 });
  txt(s, [run(LOREP_LONG), run('.', { fontSize: 8 })],
      { x: 0.769, y: 4.846, w: 3.446, h: 1.132, fontSize: 10.5, lineSpacingMultiple: 1.5 });
  sideTab(s, C.amber);
});

// 19 — SWOT circles
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  const swot = [
    { x: 1.603, tag: 1.749, letter: 'S', disc: C.yellow, ink: C.teal, quote: 2.028 },
    { x: 4.157, tag: 4.303, letter: 'W', disc: C.teal, ink: C.yellow, quote: 4.582 },
    { x: 6.711, tag: 6.857, letter: 'O', disc: C.yellow, ink: C.teal, quote: 7.136 },
    { x: 9.266, tag: 9.411, letter: 'T', disc: C.teal, ink: C.yellow, quote: 9.69 },
  ];
  const tagShadow = { type: 'outer', blur: 4, offset: 3, angle: 45, color: '000000', opacity: 0.35 };
  swot.forEach(function (c) {
    s.addShape('ellipse', { x: c.x, y: 2.094, w: 2.846, h: 2.846, fill: { color: c.disc },
                            line: { type: 'none' } });
  });
  swot.forEach(function (c) {
    s.addShape('ellipse', { x: c.tag, y: 2.05, w: 0.814, h: 0.814, fill: { color: c.disc },
                            line: { type: 'none' }, shadow: tagShadow });
    txt(s, c.letter, { x: c.tag, y: 2.05, w: 0.814, h: 0.814, fontSize: 36, bold: true,
                       color: c.ink, align: 'center', valign: 'middle' });
    txt(s, [run('\u201CLorem sit amet, ', { bold: true, italic: true, color: c.ink, breakLine: true }),
            run('conse ctetur dolor sit amet  ctetur dolor sit ipsum dolor sit\u201D',
                { bold: true, italic: true, color: c.ink })],
        { x: c.quote, y: 2.907, w: 1.996, h: 1.313, fontSize: 12, align: 'center',
          lineSpacingMultiple: 1.5 });
  });
  [2.156, 7.264].forEach(function (x) {
    txt(s, SWOT_BODY, { x, y: 5.466, w: 4.257, h: 0.978, fontSize: 12, lineSpacingMultiple: 1.5 });
  });
  txt(s, 'our main info graphics.', { x: 3.394, y: 0.811, w: 7.088, h: 0.909, fontSize: 48,
                                      bold: true, italic: true });
  sideTab(s, C.teal);
});

// 20 — Diamond of four icon tiles
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  s.addShape('diamond', { x: 7.116, y: 1.828, w: 4.616, h: 4.616, fill: { color: C.yellow },
                          line: { type: 'none' } });
  [[7.555, 2.267], [9.493, 2.267], [7.555, 4.206], [9.493, 4.206]].forEach(function (p) {
    s.addShape('roundRect', { x: p[0], y: p[1], w: 1.8, h: 1.8, rectRadius: 0.29,
                              fill: { color: C.teal }, line: { type: 'none' } });
  });
  icon(s, 'key', 8.22, 2.863, 0.554, 0.554, C.yellow, C.teal);
  icon(s, 'woman', 10.11, 2.833, 0.543, 0.634, C.yellow, C.teal);
  icon(s, 'cart', 8.162, 4.85, 0.585, 0.585, C.yellow, C.teal);
  icon(s, 'shopbag', 10.145, 4.817, 0.443, 0.577, C.yellow, C.teal);
  [{ y: 2.752, cx: 1.663, tx: 2.518, letter: 'A', body: ONSEQUAE + ' Eatecum Cum' },
   { y: 3.886, cx: 1.663, tx: 2.518, letter: 'B', body: ONSEQUAE },
   { y: 5.05, cx: 1.648, tx: 2.504, letter: 'C', body: ONSEQUAE }].forEach(function (r) {
    s.addShape('ellipse', { x: r.cx, y: r.y, w: 0.591, h: 0.591, fill: { color: C.teal },
                            line: { type: 'none' } });
    txt(s, r.letter, { x: r.cx, y: r.y, w: 0.591, h: 0.591, fontSize: 18, color: C.yellow,
                       fontFace: F.ssp, align: 'center', valign: 'middle' });
    txt(s, r.body, { x: r.tx, y: r.y, w: 3.846, h: 0.766, fontSize: 10.5, italic: true,
                     lineSpacingMultiple: 1.5, margin: 0 });
  });
  txt(s, 'our main info graphics.', { x: 3.394, y: 0.811, w: 7.088, h: 0.909, fontSize: 48,
                                      bold: true, italic: true });
  sideTab(s, C.teal);
});

// 21 — Rotated diamond + progress bars
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  s.addShape('roundRect', { x: 1.832, y: 2.736, w: 3.046, h: 3.046, rectRadius: 0.508, rotate: 225,
                            fill: { color: C.teal }, line: { type: 'none' } });
  txt(s, INTERACTIVE + 'Completely pursue scalable charity event to bes success. ',
      { x: 2.184, y: 3.668, w: 2.342, h: 1.325, fontSize: 10.5, bold: true, italic: true,
        color: C.yellow, align: 'center', lineSpacingMultiple: 1.5, margin: 0 });
  const bars = [
    { y: 2.639, x: 6.588, w: 2.613, color: C.orange, pct: '30%', px: 9.448, pw: 0.426 },
    { y: 3.255, x: 6.581, w: 2.613, color: C.gray, pct: '30%', px: 9.442, pw: 0.426 },
    { y: 3.958, x: 6.581, w: 3.009, color: C.amber, pct: '45%', px: 9.833, pw: 0.446 },
  ];
  bars.forEach(function (b) {
    txt(s, 'Main Text Here', { x: b.x, y: b.y, w: 1.94, h: 0.185, fontSize: 11, bold: true,
                               fontFace: F.ssp, margin: 0 });
    rule(s, b.x, b.y + 0.289, b.w, 6, b.color);
    txt(s, b.pct, { x: b.px, y: b.y + 0.205, w: b.pw, h: 0.177, fontSize: 10.5, color: C.ink,
                    fontFace: F.ssp, margin: 0 });
  });
  txt(s, INTERACTIVE + 'Completely pursue scalable charity event to bes success. Collaboratively ' +
         'administrate empowered markets. coordinate proactive e-commto bes success. Collaboratively administrate',
      { x: 6.5, y: 4.727, w: 5.803, h: 1.111, fontSize: 11, align: 'justify',
        lineSpacingMultiple: 1.5, margin: 0 });
  txt(s, 'our main info graphics.', { x: 3.394, y: 0.811, w: 7.088, h: 0.909, fontSize: 48,
                                      bold: true, italic: true });
  sideTab(s, C.teal);
});

// 22 — About Our Service
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  photoSlot(s, 1.152, 0.847, 5.153, 2.689);
  photoSlot(s, 1.143, 3.871, 5.153, 2.684);
  rect(s, 0, 0, 3.54, 7.5, C.yellow);
  txt(s, [run('About Our', { bold: true, breakLine: true }), run('Service.', { bold: true })],
      { x: 7.508, y: 1.159, w: 5.405, h: 1.178, fontSize: 32 });
  txt(s, 'Fashion 2022 Presentation Template', { x: 7.536, y: 2.41, w: 3.826, h: 0.269, fontSize: 10 });
  txt(s, IPSUM_LONG + 'Lorem Ipsum\u00a0is simply dummy text of the printing and typesetting industry. ',
      { x: 7.508, y: 3.033, w: 4.778, h: 1.01, fontSize: 9, align: 'justify', lineSpacingMultiple: 1.5 });
  [{ cy: 4.583, ico: 'phone', ix: 7.879, iy: 4.744, iw: 0.182, ih: 0.314, tx: 8.512, ty: 4.611 },
   { cy: 5.625, ico: 'monitor', ix: 7.82, iy: 5.812, iw: 0.3, ih: 0.262, tx: 8.492, ty: 5.667 }
  ].forEach(function (r) {
    s.addShape('ellipse', { x: 7.655, y: r.cy, w: 0.63, h: 0.63, fill: { color: C.yellow },
                            line: { type: 'none' } });
    icon(s, r.ico, r.ix, r.iy, r.iw, r.ih, C.teal);
    txt(s, IPSUM_SHORT, { x: r.tx, y: r.ty, w: 3.571, h: 0.53, fontSize: 9, align: 'justify',
                          lineSpacingMultiple: 1.5 });
  });
  sideTab(s, C.teal);
});

// 23 — Meet The Team
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  photoSlot(s, 7.538, 0.852, 5.153, 2.684);
  photoSlot(s, 7.538, 3.871, 5.153, 2.684);
  rect(s, 10.569, 0, 2.764, 7.5, C.yellow);
  txt(s, 'Meet The Team.', { x: 1.178, y: 1.673, w: 4.878, h: 0.64, fontSize: 32, bold: true });
  txt(s, 'Fashion 2022 Presentation Template', { x: 1.206, y: 2.382, w: 3.826, h: 0.269, fontSize: 10 });
  txt(s, IPSUM_LONG + 'Lorem Ipsum\u00a0is simply dummy text of the printing and typesetting',
      { x: 1.178, y: 3.159, w: 4.669, h: 1.01, fontSize: 9, align: 'justify', lineSpacingMultiple: 1.5 });
  txt(s, IPSUM_LONG, { x: 1.178, y: 4.282, w: 4.669, h: 0.783, fontSize: 9, align: 'justify',
                       lineSpacingMultiple: 1.5 });
  txt(s, 'Fashion New', { x: 1.206, y: 5.595, w: 1.569, h: 0.269, fontSize: 10, bold: true });
  rule(s, 2.741, 5.738, 1.575, 1);
});

// 24 — Bullet list + interleaved arrows
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  // drawn in the source deck's z-order: both yellow arrows first, teal over them
  [{ x: 7.335, y: 3.234, w: 6.377, h: 1.635, head: 1.681, color: C.yellow },
   { x: 7.165, y: 5.089, w: 6.558, h: 1.622, head: 1.462, color: C.yellow },
   { x: 10.028, y: 3.947, w: 3.305, h: 1.826, head: 0.993, color: C.teal },
   { x: 9.302, y: 2.080, w: 4.031, h: 1.827, head: 0.986, color: C.teal }].forEach(function (a) {
    leftArrow(s, a.x, a.y, a.w, a.h, a.color, a.head);
  });
  [{ y: 2.248, h: 0.766, text: INTERACTIVE + 'Completely pursue scalable charity event to bes success. ' },
   { y: 3.442, h: 0.501, text: INTERACTIVE },
   { y: 4.636, h: 0.766, text: INTERACTIVE + 'Completely pursue scalable charity event to bes success.' },
   { y: 5.83, h: 0.501, text: INTERACTIVE }].forEach(function (b) {
    txt(s, b.text, { x: 1.749, y: b.y, w: 5.019, h: b.h, fontSize: 10.5, lineSpacingMultiple: 1.5,
                     margin: 0, bullet: { characterCode: '2022', indent: 13.5 } });
  });
  txt(s, 'our main info graphics.', { x: 3.394, y: 0.811, w: 7.204, h: 0.909, fontSize: 48,
                                      bold: true, italic: true });
  sideTab(s, C.teal);
});

// 25 — Three labelled blobs
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  const blobs = [
    { bx: 2.381, by: 1.754, color: C.yellow, px: 1.891, py: 1.644, ico: 'bell', ix: 3.163, iy: 2.419,
      iw: 0.542, ih: 0.593, ic: C.teal, tx: 2.04, lx: 2.639 },
    { bx: 5.773, by: 1.821, color: C.teal, px: 5.283, py: 3.193, ico: 'chat', ix: 6.589, iy: 2.556,
      iw: 0.475, ih: 0.44, ic: C.yellow, tx: 5.505, lx: 5.896 },
    { bx: 8.922, by: 1.782, color: C.yellow, px: 10.118, py: 1.642, ico: 'globe', ix: 9.711, iy: 2.45,
      iw: 0.529, ih: 0.53, ic: C.teal, tx: 8.97, lx: 9.355 },
  ];
  blobs.forEach(function (b) {
    s.addShape('flowChartTerminator', { x: b.bx, y: b.by, w: 2.106, h: 1.737, fill: { color: b.color },
                                        line: { type: 'none' } });
    icon(s, b.ico, b.ix, b.iy, b.iw, b.ih, b.ic, b.color);
    s.addShape('roundRect', { x: b.px, y: b.py, w: 1.401, h: 0.557, rectRadius: 0.2785,
                              fill: { color: C.white }, line: { color: b.color, width: 3 } });
    txt(s, 'Lorem Ipsum\u00a0', { x: b.px, y: b.py, w: 1.401, h: 0.557, fontSize: 12, align: 'center',
                                  valign: 'middle', lineSpacingMultiple: 0.9 });
    rule(s, b.lx, 4.207, 1.575, 1);
    txt(s, IPSUM_SHORT.replace('has been', "has been the industry's standard dummy"),
        { x: b.tx, y: 4.788, w: 2.596, h: 1.178, fontSize: 12, align: 'center',
          lineSpacingMultiple: 1.5, margin: 0 });
  });
});

// 26 — Chevron timeline
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  for (let i = 0; i < 11; i++) {
    chevron(s, 1.268 + i * 0.6887, 4.002, 0.828, 0.756, i % 2 === 0 ? C.teal : C.yellow, 0.301);
  }
  icon(s, 'magnifier', 1.634, 4.272, 0.218, 0.218, C.yellow);
  icon(s, 'star', 3.703, 4.272, 0.229, 0.222, C.teal);
  icon(s, 'plane', 6.463, 4.271, 0.219, 0.219, C.teal);
  icon(s, 'crew', 8.529, 4.263, 0.219, 0.219, C.yellow);
  // callouts: [labelX, labelY, connector vertical X, connector Y-top, elbow Y, elbow x-from/to]
  const steps = [
    { label: 'Step One', lw: 1.578, x: 1.938, y: 5.281, vx: 1.813, vy: 4.758, ey: 5.658, ex: 1.527 },
    { label: 'Step Two', lw: 1.585, x: 4.086, y: 2.718, vx: 3.628, vy: 3.102, ey: 3.102, ex: 3.914 },
    { label: 'Step Three', lw: 1.772, x: 6.794, y: 5.281, vx: 6.669, vy: 4.758, ey: 5.658, ex: 6.383 },
    { label: 'Step Four', lw: 1.617, x: 8.972, y: 2.72, vx: 8.514, vy: 3.104, ey: 3.104, ex: 8.8 },
  ];
  steps.forEach(function (st) {
    s.addShape('line', { x: st.vx, y: st.vy, w: 0, h: 0.9,
                         line: { color: C.teal, width: 0.75, dashType: 'dot' } });
    rule(s, Math.min(st.vx, st.ex), st.ey, Math.abs(st.ex - st.vx), 0.75, C.teal, 'dot');
    txt(s, st.label, { x: st.x, y: st.y, w: st.lw, h: 0.337, fontSize: 14, bold: true });
    txt(s, STEP_BODY, { x: st.x, y: st.y + 0.4, w: 3.38, h: 0.53, fontSize: 9, lineSpacingMultiple: 1.5 });
  });
  txt(s, 'Infographic Slide.', { x: 3.392, y: 0.81, w: 6.555, h: 0.64, fontSize: 32, bold: true,
                                 align: 'center' });
  txt(s, 'Presentation Template', { x: 4.769, y: 1.558, w: 3.826, h: 0.269, fontSize: 10,
                                    align: 'center' });
});

// 27 — Get Connected
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  photoSlot(s, 0.602, 0.592, 4.439, 6.266);
  rect(s, 5.917, 3.329, 6.774, 3.531, C.yellow);
  txt(s, [run('Get Connected', { bold: true, breakLine: true }), run('With Us Here.', { bold: true })],
      { x: 6.943, y: 1.248, w: 5.64, h: 1.178, fontSize: 32 });
  txt(s, 'Presentation Template', { x: 6.971, y: 2.509, w: 3.826, h: 0.269, fontSize: 10 });
  const contacts = [
    { x: 6.943, y: 3.819, head: 'Our Address', rows: [['Lorem Ipsum\u00a0is simply ', 4.209, 2.023],
                                                      ['Lorem Ipsum\u00a0is simply', 4.569, 2.104]] },
    { x: 9.604, y: 3.819, head: 'Office House', rows: [['Lorem Ipsum\u00a0is simply', 4.209, 2.104],
                                                       ['Lorem Ipsum\u00a0Dolor', 4.569, 2.104]] },
    { x: 6.943, y: 5.261, head: 'Get It Touch', rows: [['www.example.com', 5.647, 1.917]] },
    { x: 9.604, y: 5.265, head: 'Follow Us', rows: [['@example', 5.654, 1.833], ['@example', 6.03, 1.833]] },
  ];
  contacts.forEach(function (c) {
    txt(s, c.head, { x: c.x, y: c.y, w: 1.99, h: 0.303, fontSize: 12, bold: true });
    c.rows.forEach(function (r) {
      txt(s, r[0], { x: c.x, y: r[1], w: r[2], h: 0.307, fontSize: 9, lineSpacingMultiple: 1.5 });
    });
  });
  sideTab(s, C.teal);
});

// 28 — Thanks
build.push(function (s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  photoSlot(s, 4.853, 0, 2.86, 7.5);
  rect(s, 0, 0, 4.853, 7.5, C.yellow);
  rect(s, 2.009, 4.394, 5.689, 2.0, C.teal);
  txt(s, 'THANKS', { x: 2.324, y: 4.66, w: 4.921, h: 1.447, fontSize: 80, fontFace: F.med,
                     color: C.yellow, align: 'center' });
  sideTab(s, C.teal);
});

/* -------------------------------------------------------------------- main */

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'CUSTOM_16x9', width: 13.333, height: 7.5 });
pptx.layout = 'CUSTOM_16x9';
pptx.title = 'Perm. Creative';

build.forEach(function (fn) { fn(pptx.addSlide()); });

pptx.writeFile({ fileName: path.join(__dirname, '14ab0e18-5e3a-49ca-85ac-606b6d495c75_grok_final.pptx') })
  .then(function (f) { console.log('wrote ' + f); });
