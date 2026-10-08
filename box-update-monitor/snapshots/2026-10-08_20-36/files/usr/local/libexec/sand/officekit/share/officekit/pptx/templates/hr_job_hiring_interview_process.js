/**
 * Recreation of "Job Interview" template deck (27 slides, 13.333 x 7.5 in)
 * with pptxgenjs only.  Raster/vector artwork from the source deck is
 * represented by labelled placeholder blocks.
 *
 *   node 1327501b-6de9-4f96-9f3b-729931362550_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * palette / typography                                                *
 * ------------------------------------------------------------------ */

const C = {
  dark: '3E5050',   // accent1 - deep slate green
  green: '61AE84',  // accent2 - brand green
  cream: 'F2F1EF',  // accent3 - page background
  white: 'FFFFFF',
  black: '000000',
  gray: '7F7F7F',
  mint: '9FCEB4',   // light green line art
  pale: 'DEEEE6',   // very light green
  sage: 'BEDECD',
  silver: 'D0CECE',
  smoke: 'E7E6E6',  // lt2
  forest: '2D5941',
  ink: '2E3C3C',
};

const F = {
  p: 'Poppins',
  pl: 'Poppins Light',
  pm: 'Poppins Medium',
  ps: 'Poppins SemiBold',
  i: 'Inter',
  il: 'Inter Light',
  im: 'Inter Medium',
  is: 'Inter SemiBold',
  a: 'Arimo',
  r: 'Roboto',
  rm: 'Roboto Medium',
};

// outer shadows reused across the deck (blur/offset in points)
const SH = {
  card: { type: 'outer', color: C.black, opacity: 0.07, blur: 8, offset: 3, angle: 100 },
  soft: { type: 'outer', color: C.black, opacity: 0.15, blur: 3, offset: 1, angle: 90 },
  chip: { type: 'outer', color: C.black, opacity: 0.15, blur: 5, offset: 2, angle: 45 },
  lift: { type: 'outer', color: C.black, opacity: 0.1, blur: 6, offset: 3, angle: 45 },
  drop: { type: 'outer', color: C.black, opacity: 0.25, blur: 35, offset: 24, angle: 50 },
  wide: { type: 'outer', color: C.black, opacity: 0.26, blur: 42, offset: 21, angle: 30 },
  down: { type: 'outer', color: C.black, opacity: 0.3, blur: 20, offset: 12, angle: 90 },
};

/* ------------------------------------------------------------------ *
 * small drawing helpers                                               *
 * ------------------------------------------------------------------ */

// rounded rectangle whose corner radius is a fraction of its short side
const rr = (o, frac) =>
  Object.assign({ rectRadius: frac * Math.min(o.w, o.h) }, o);

// fully rounded ("pill") rectangle
const pill = (o) => rr(o, 0.5);

// straight line segment between two absolute points
function seg(s, x1, y1, x2, y2, color, width, dashType) {
  s.addShape('line', {
    x: Math.min(x1, x2), y: Math.min(y1, y2),
    w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
    flipV: (x2 - x1) * (y2 - y1) < 0,
    line: { color, width, dashType: dashType || 'solid' },
  });
}

// free-form outline / polygon from a compact point table.
// each entry: [x,y,1] moveTo | [x,y] lineTo | [x,y,cx1,cy1,cx2,cy2] cubic | [] close
function poly(s, pts, o) {
  const P = pts.map((p) => {
    if (p.length === 0) return { close: true };
    const x = p[0] * o.w, y = p[1] * o.h;
    if (p.length === 6) {
      return { x, y, curve: { type: 'cubic', x1: p[2] * o.w, y1: p[3] * o.h, x2: p[4] * o.w, y2: p[5] * o.h } };
    }
    return p.length === 3 ? { x, y, moveTo: true } : { x, y };
  });
  s.addShape('custGeom', Object.assign({ points: P }, o));
}

/* --- reusable free-form paths (unit space, scaled by poly()) -------- */

// soft "guitar-pick" blob used behind the portrait photos
const BLOB = [[0.525, 0.999, 1], [0.008, 0.237, 0.369, 1.028, -0.063, 0.467],
[0.878, 0.106, 0.079, 0.006, 0.569, -0.091], [0.525, 0.999, 1.209, 0.318, 0.789, 0.95], []];

// thin sine squiggle used as a decorative accent
const SQUIGGLE = [[0, 0.866, 1], [0.201, 0.167, 0.063, 0.505, 0.126, 0.145],
[0.452, 0.999, 0.276, 0.189, 0.383, 1.027], [0.615, 0, 0.521, 0.972, 0.544, 0.017],
[0.879, 0.899, 0.686, -0.016, 0.816, 0.861], [0.992, 0.233, 0.941, 0.938, 1.027, 0.689]];

// chunky upward arrow (brand mark)
const ARROW_UP = [[1, 1, 1], [0.586, 0.324, 1, 1, 0.586, 0.899], [0.828, 0.324], [0.414, 0],
[0, 0.324, 0.414, -0.004, 0, 0.324], [0.224, 0.324], [1, 1, 0.224, 0.326, 0.288, 0.873], []];

// top-left corner wash on slide 17
const CORNER = [[0, 0, 1], [0.986, 0], [0.982, 0.011], [0.975, 0.568, 0.939, 0.129, 0.926, 0.306],
[0.98, 0.594, 0.976, 0.577, 0.978, 0.585], [0, 0.513, 1.167, 1.553, 0, 0.513], [0, 0], []];

// wide swoosh backdrop on slide 20
const SWOOSH = [[0.967, 0, 1], [0.981, 0.017, 0.972, 0.001, 0.977, 0.006],
[1, 0.11, 0.987, 0.034, 0.994, 0.066], [1, 0.112], [1, 1], [0, 1], [0, 0.53], [0.001, 0.522],
[0.089, 0.318, 0.022, 0.41, 0.049, 0.329], [0.389, 0.704, 0.16, 0.299, 0.314, 0.74],
[0.541, 0.101, 0.464, 0.668, 0.48, 0.153], [0.754, 0.392, 0.601, 0.049, 0.681, 0.406],
[0.967, 0, 0.82, 0.379, 0.918, -0.009], []];

const squiggle = (s, x, y, w, h, color) =>
  poly(s, SQUIGGLE, { x, y, w, h, fill: { type: 'none' }, line: { color: color || C.mint, width: 3 } });

const blob = (s, o) => poly(s, BLOB, o);

const arrowUp = (s, x, y, w, h, color, shadow) =>
  poly(s, ARROW_UP, { x, y, w, h, fill: { color }, shadow });

/* ------------------------------------------------------------------ *
 * icon library - every icon is drawn from native shapes               *
 * ------------------------------------------------------------------ */

const ICON = {
  magnifier(s, x, y, z, color, wt) {
    const w = wt || Math.max(0.75, z * 9);
    s.addShape('ellipse', { x: x + 0.26 * z, y, w: 0.74 * z, h: 0.74 * z, fill: { type: 'none' }, line: { color, width: w } });
    seg(s, x + 0.02 * z, y + 0.98 * z, x + 0.32 * z, y + 0.68 * z, color, w);
  },
  // three stacked chevrons; each sits on a slightly larger plate in the
  // background colour so the layers stay visually separated when small
  layers(s, x, y, z, color, hole) {
    [0.58, 0.29, 0].forEach((dy) => {
      s.addShape('diamond', { x: x - 0.05 * z, y: y + (dy - 0.05) * z, w: 1.1 * z, h: 0.52 * z, fill: { color: hole || C.white } });
      s.addShape('diamond', { x, y: y + dy * z, w: z, h: 0.42 * z, fill: { color } });
    });
  },
  cloud(s, x, y, z, color) {
    [[0, 0.4, 0.42], [0.24, 0.16, 0.5], [0.6, 0.36, 0.4]].forEach(([dx, dy, d]) =>
      s.addShape('ellipse', { x: x + dx * z, y: y + dy * z, w: d * z, h: d * z, fill: { color } }));
    s.addShape('rect', { x: x + 0.1 * z, y: y + 0.52 * z, w: 0.8 * z, h: 0.24 * z, fill: { color } });
    s.addShape('downArrow', { x: x + 0.4 * z, y: y + 0.44 * z, w: 0.2 * z, h: 0.44 * z, fill: { color } });
  },
  briefcase(s, x, y, z, color) {
    s.addShape('blockArc', { x: x + 0.28 * z, y: y + 0.04 * z, w: 0.44 * z, h: 0.44 * z, angleRange: [180, 360], arcThicknessRatio: 0.34, fill: { color } });
    s.addShape(...['roundRect', rr({ x, y: y + 0.26 * z, w: z, h: 0.62 * z, fill: { color } }, 0.18)]);
  },
  inbox(s, x, y, z, color, hole) {
    s.addShape(...['roundRect', rr({ x, y: y + 0.18 * z, w: z, h: 0.64 * z, fill: { color } }, 0.16)]);
    s.addShape('trapezoid', { x: x + 0.14 * z, y: y + 0.26 * z, w: 0.72 * z, h: 0.3 * z, rotate: 180, fill: { color: hole || C.white } });
  },
  gem(s, x, y, z, color) {
    poly(s, [[0.16, 0.32, 1], [0.3, 0.06], [0.7, 0.06], [0.84, 0.32], [0.5, 0.94], []],
      { x, y, w: z, h: z, fill: { color } });
    seg(s, x + 0.16 * z, y + 0.32 * z, x + 0.84 * z, y + 0.32 * z, C.white, 0.75);
  },
  globe(s, x, y, z, color) {
    s.addShape('ellipse', { x, y, w: z, h: z, fill: { type: 'none' }, line: { color, width: 1.25 } });
    s.addShape('ellipse', { x: x + 0.3 * z, y, w: 0.4 * z, h: z, fill: { type: 'none' }, line: { color, width: 1 } });
    seg(s, x, y + 0.5 * z, x + z, y + 0.5 * z, color, 1);
  },
  lamp(s, x, y, z, color) {
    s.addShape('ellipse', { x: x + 0.18 * z, y, w: 0.64 * z, h: 0.64 * z, fill: { color } });
    s.addShape('rect', { x: x + 0.36 * z, y: y + 0.62 * z, w: 0.28 * z, h: 0.2 * z, fill: { color } });
    seg(s, x + 0.34 * z, y + 0.92 * z, x + 0.66 * z, y + 0.92 * z, color, 1.25);
  },
  gear(s, x, y, z, color, hole) {
    s.addShape('gear6', { x: x - 0.06 * z, y: y + 0.1 * z, w: 1.12 * z, h: 0.8 * z, fill: { color } });
    s.addShape('ellipse', { x: x + 0.36 * z, y: y + 0.36 * z, w: 0.28 * z, h: 0.28 * z, fill: { color: hole || C.white } });
  },
  scissors(s, x, y, z, color) {
    seg(s, x + 0.14 * z, y + 0.04 * z, x + 0.8 * z, y + 0.66 * z, color, 1.25);
    seg(s, x + 0.86 * z, y + 0.04 * z, x + 0.2 * z, y + 0.66 * z, color, 1.25);
    [0.02, 0.64].forEach((dx) =>
      s.addShape('ellipse', { x: x + dx * z, y: y + 0.62 * z, w: 0.34 * z, h: 0.34 * z, fill: { type: 'none' }, line: { color, width: 1.25 } }));
  },
  lock(s, x, y, z, color) {
    s.addShape('blockArc', { x: x + 0.24 * z, y: y + 0.06 * z, w: 0.52 * z, h: 0.6 * z, angleRange: [180, 360], arcThicknessRatio: 0.3, fill: { color } });
    s.addShape(...['roundRect', rr({ x: x + 0.1 * z, y: y + 0.42 * z, w: 0.8 * z, h: 0.52 * z, fill: { color } }, 0.2)]);
  },
  pin(s, x, y, z, color) {
    s.addShape('teardrop', { x: x + 0.14 * z, y: y + 0.05 * z, w: 0.72 * z, h: 0.72 * z, rotate: 135, fill: { color } });
    s.addShape('ellipse', { x: x + 0.36 * z, y: y + 0.22 * z, w: 0.28 * z, h: 0.28 * z, fill: { color: C.white } });
  },
  mail(s, x, y, z, color, alt) {
    s.addShape('rect', { x, y: y + 0.2 * z, w: z, h: 0.6 * z, fill: { color } });
    s.addShape('triangle', { x: x + 0.06 * z, y: y + 0.24 * z, w: 0.88 * z, h: 0.36 * z, rotate: 180, fill: { color: alt || C.white } });
  },
  // handset: a thick "C" (block arc) rotated so the ends point up-left/down-right
  phone(s, x, y, z, color) {
    s.addShape('blockArc', { x, y, w: z, h: z, rotate: 225, angleRange: [200, 340], arcThicknessRatio: 0.62, fill: { color } });
  },
  wallet(s, x, y, z, color, hole) {
    s.addShape(...['roundRect', rr({ x, y: y + 0.18 * z, w: z, h: 0.64 * z, fill: { color } }, 0.14)]);
    s.addShape(...['roundRect', rr({ x: x + 0.56 * z, y: y + 0.36 * z, w: 0.48 * z, h: 0.28 * z, fill: { color: hole || C.white } }, 0.18)]);
    s.addShape('ellipse', { x: x + 0.72 * z, y: y + 0.44 * z, w: 0.12 * z, h: 0.12 * z, fill: { color } });
  },
  arrowDown(s, x, y, z, color) {
    s.addShape('ellipse', { x, y, w: z, h: z, fill: { type: 'none' }, line: { color, width: 1 } });
    seg(s, x + 0.5 * z, y + 0.24 * z, x + 0.5 * z, y + 0.66 * z, color, 1);
    s.addShape('triangle', { x: x + 0.34 * z, y: y + 0.55 * z, w: 0.32 * z, h: 0.24 * z, rotate: 180, fill: { color } });
  },
};

// small circular badge with an icon inside
function iconBadge(s, x, y, d, o) {
  const opt = { x, y, w: d, h: d, fill: o.bg ? { color: o.bg } : { type: 'none' } };
  if (o.ring) opt.line = { color: o.ring, width: o.ringW || 1 };
  if (o.shadow) opt.shadow = o.shadow;
  s.addShape(o.square ? 'roundRect' : 'ellipse', o.square ? rr(opt, 0.28) : opt);
  const z = d * (o.scale || 0.46);
  // 5th arg = colour showing through cut-outs (matches whatever is behind the icon)
  ICON[o.icon](s, x + (d - z) / 2, y + (d - z) / 2, z, o.fg, o.hole || o.bg);
}

// circle + chevron badge used in front of "Let's Get's Stated" / "About Us"
function chevronBadge(s, x, y, d, fill) {
  s.addShape('ellipse', { x, y, w: d, h: d, fill: { color: fill }, shadow: SH.soft });
  s.addShape('chevron', { x: x + 0.4 * d, y: y + 0.31 * d, w: 0.28 * d, h: 0.37 * d, fill: { color: C.white } });
}

/* ------------------------------------------------------------------ *
 * text helpers                                                        *
 * ------------------------------------------------------------------ */

// a text run:  run('Hello ', {b:1})   b=SemiBold face, l=Light face
const run = (text, o) => ({ text, options: o || {} });
const br = (o) => Object.assign({ breakLine: true }, o || {});

function text(s, runs, o) {
  s.addText(runs, Object.assign({ valign: 'top', isTextBox: true }, o));
}

// two-tone headline: light/regular first half + semibold second half
function heading(s, o, parts) {
  text(s, parts.map((p) => run(p[0], {
    fontFace: p[2] || F.ps, fontSize: o.size || 36, color: p[1], breakLine: !!p[3],
  })), {
    x: o.x, y: o.y, w: o.w, h: o.h, align: o.align || 'left',
    lineSpacingMultiple: o.ls || 1.0,
  });
}

// grey body copy
function body(s, o, str, opts) {
  text(s, [run(str, {
    fontFace: (opts && opts.face) || F.i,
    fontSize: (opts && opts.size) || 10,
    color: (opts && opts.color) || C.gray,
  })], {
    x: o.x, y: o.y, w: o.w, h: o.h,
    align: (opts && opts.align) || 'left',
    lineSpacingMultiple: (opts && opts.ls) || 1.5,
  });
}

// bold caption above body copy
function caption(s, o, str, opts) {
  text(s, [run(str, {
    fontFace: (opts && opts.face) || F.is,
    fontSize: (opts && opts.size) || 12,
    color: (opts && opts.color) || C.black,
  })], { x: o.x, y: o.y, w: o.w, h: o.h, align: (opts && opts.align) || 'left', lineSpacingMultiple: (opts && opts.ls) || 1 });
}

// pill button with a centred label
function button(s, o) {
  s.addShape(...['roundRect', pill({
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: o.fill ? { color: o.fill } : { type: 'none' },
    line: o.line ? { color: o.line, width: 1 } : undefined,
    shadow: o.shadow,
  })]);
  text(s, [run(o.label, { fontFace: o.face || F.i, fontSize: o.size || 12, color: o.color || C.black })],
    { x: o.x, y: o.y, w: o.w, h: o.h, align: 'center', valign: 'middle' });
}

// stand-in for the large drawn illustrations of the source deck
// (slides 23-25); a soft neutral block with a short caption
function artwork(s, x, y, w, h, label, captionAtBottom) {
  s.addShape(...['roundRect', rr({ x, y, w, h, fill: { color: 'EAEAE8' } }, 0.03)]);
  text(s, [run(label, { fontFace: F.i, fontSize: 10, color: 'A8A8A4' })],
    { x, y: captionAtBottom ? y + h - 0.42 : y + h / 2 - 0.2, w, h: 0.4, align: 'center', valign: 'middle' });
}

/* ------------------------------------------------------------------ *
 * shared copy                                                         *
 * ------------------------------------------------------------------ */

const T = {
  designers: 'Designers generally have nothing to do with creating content for their projects. Even so,  site can be incomplete Designers generally have nothing to do with creating content for.',
  lorem: "Lorem ipsum\u00a0is simply dummy text of the printing and typesetting industry. Lorem ipsum has been the industry's standard dummy text ever since the",
  loremLong: "Lorem ipsum\u00a0is simply dummy text of the printing and typesetting industry. Lorem ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley",
  loremShort: 'Lorem ipsum\u00a0is simply dummy text of the printing and typesetting',
  mistaken: 'PLACEHOLDER',
  mistakenOf: 'But I must explain to you how all this mistaken idea of',
  mistakenAnd: 'But I must explain to you how all this mistaken idea of and',
  variations: 'There are many variations Lorem available, but majority have suffered alteration in some',
  constituter: 'Constituter Adipescent Elastoses Daimio',
  applicant: 'Applicant Interview',
  execSmarmy: 'Write Your Executive Smarmy off',
};

/* ------------------------------------------------------------------ *
 * per-slide builders                                                  *
 * ------------------------------------------------------------------ */

// 1 - title -----------------------------------------------------------
function slide01(s) {
  ICON.magnifier(s, 1.537, 2.612, 3.6, 'E4EDE7', 26);
  [[3.94, 0.758, 0.158], [3.94, 0.842, 0.267], [3.94, 0.925, 0.376]]
    .forEach(([x, y, w]) => seg(s, x, y, x + w, y, C.dark, 2.25));

  heading(s, { x: 3.766, y: 2.445, w: 7.403, h: 2.175, size: 88, ls: 0.7 },
    [['Job', C.black, F.ps, true], ['Interview', C.dark, F.ps]]);

  s.addShape('ellipse', { x: 11.001, y: 5.87, w: 0.286, h: 0.286, fill: { color: C.green, transparency: 32 } });
  s.addShape('ellipse', { x: 11.692, y: 6.365, w: 0.079, h: 0.079, fill: { color: C.green, transparency: 32 } });
  squiggle(s, 10.637, 1.183, 0.887, 0.334);
  squiggle(s, 0.627, 6.748, 0.555, 0.209);
  ICON.arrowDown(s, 9.802, 3.428, 0.345, C.dark);

  s.addShape(...['roundRect', pill({
    x: 3.94, y: 4.637, w: 4.243, h: 0.516,
    fill: { color: C.white }, line: { color: C.green, width: 1 }, shadow: SH.wide,
  })]);
  text(s, [run('Job Hiring Presentation Template', { fontFace: F.i, fontSize: 11, color: C.gray })],
    { x: 4.084, y: 4.733, w: 4.71, h: 0.31, lineSpacingMultiple: 1.45 });
  s.addShape('ellipse', { x: 7.7, y: 4.684, w: 0.421, h: 0.421, fill: { color: C.green } });
  ICON.magnifier(s, 7.833, 4.828, 0.156, C.white, 1.25);

  text(s, [run('www.example.com', { fontFace: F.i, fontSize: 12, color: C.dark })],
    { x: 3.94, y: 6.543, w: 3.446, h: 0.31, lineSpacingMultiple: 1.33 });
}

// 2 - welcome ---------------------------------------------------------
function slide02(s) {
  s.addShape('rect', { x: 0, y: 0, w: 13.333, h: 0.175, fill: { color: C.dark } });
  s.addShape('rect', { x: 0.005, y: 7.296, w: 13.333, h: 0.209, fill: { color: C.dark } });

  blob(s, { x: 1.432, y: 2.355, w: 4.397, h: 4.149, fill: { color: C.white, transparency: 46 } });
  blob(s, { x: 1.585, y: 2.477, w: 4.091, h: 3.861, fill: { color: C.green } });

  seg(s, 3.62, 1.792, 7.513, 1.792, C.dark, 1.5);
  seg(s, 3.631, 1.792, 3.631, 3.75, C.dark, 1.5);
  squiggle(s, 1.097, 6.192, 0.555, 0.209);
  s.addShape('ellipse', { x: 5.829, y: 5.513, w: 0.286, h: 0.286, fill: { color: C.green, transparency: 36 } });
  s.addShape('ellipse', { x: 6.52, y: 5.808, w: 0.079, h: 0.079, fill: { color: C.green, transparency: 36 } });
  s.addShape('ellipse', { x: 1.614, y: 1.53, w: 0.079, h: 0.079, fill: { color: C.green, transparency: 25 } });
  s.addShape('ellipse', { x: 0.796, y: 2.289, w: 0.144, h: 0.144, fill: { color: C.green, transparency: 25 } });
  ICON.magnifier(s, 6.801, 3.264, 0.91, 'E9F2EC', 6);

  heading(s, { x: 7.838, y: 1.174, w: 4.741, h: 1.192, ls: 0.9 },
    [['Welcome', C.green, F.p, true], ['Our Institute Ltd', C.black, F.ps]]);
  text(s, [run('Founder Her Studio', { fontFace: F.pl, fontSize: 10, color: C.black })],
    { x: 7.932, y: 2.28, w: 2.393, h: 0.269 });
  text(s, [
    run('Write your\u00a0', { fontFace: F.pl, fontSize: 20, color: C.black }),
    run('executive', br({ fontFace: F.ps, fontSize: 20, color: C.black })),
    run('Summary ', { fontFace: F.ps, fontSize: 20, color: C.black }),
    run('of where you want to take', { fontFace: F.pl, fontSize: 20, color: C.green }),
  ], { x: 7.955, y: 3.142, w: 4.068, h: 1.111, lineSpacingMultiple: 1.0 });
  body(s, { x: 7.965, y: 4.31, w: 3.849, h: 1.027 }, T.designers, { color: C.black });
  button(s, { x: 8.062, y: 5.692, w: 1.338, h: 0.439, label: 'See More\u2026.', fill: C.green, color: C.white, face: F.p, size: 11, shadow: SH.drop });
}

// 3 - about us --------------------------------------------------------
function slide03(s) {
  s.addShape('rect', { x: 0, y: 3.888, w: 13.333, h: 3.571, fill: { color: C.dark } });
  heading(s, { x: 0.981, y: 1.023, w: 4.741, h: 0.663, ls: 1.11 },
    [['About', C.green, F.pl], [' Us', C.black, F.ps]]);

  s.addShape('rect', { x: 1.112, y: 2.301, w: 11.11, h: 3.175, fill: { color: C.white } });
  s.addShape('rect', { x: 1.333, y: 2.507, w: 5.333, h: 2.746, fill: { color: C.cream } });
  s.addShape('rect', { x: 6.888, y: 2.507, w: 5.107, h: 2.746, fill: { color: C.green } });

  iconBadge(s, 1.894, 2.984, 0.638, { bg: C.green, fg: C.white, icon: 'layers', shadow: SH.lift, scale: 0.42 });
  iconBadge(s, 7.454, 2.984, 0.638, { bg: C.dark, fg: C.white, icon: 'cloud', shadow: SH.lift, scale: 0.55 });

  [[1.807, C.black, C.gray], [7.384, C.white, C.white]].forEach(([x, tc, bc]) => {
    caption(s, { x, y: 3.808, w: 2.479, h: 0.343 }, T.applicant, { face: F.ps, size: 14, color: tc, ls: 1.21 });
    body(s, { x, y: 4.151, w: 4.365, h: 0.774 }, T.lorem, { color: bc });
  });
  body(s, { x: 2.138, y: 5.91, w: 9.058, h: 0.522 }, T.designers, { color: C.white, align: 'center' });
}

// 4 - importance ------------------------------------------------------
function slide04(s) {
  s.addShape('rect', { x: 0, y: 4.069, w: 4.208, h: 3.431, fill: { color: C.green, transparency: 81 } });
  heading(s, { x: 0.949, y: 0.99, w: 5.337, h: 1.313, ls: 1.0 },
    [['Importance', C.green, F.p, true], ['Effective Heiring ', C.black, F.ps]]);

  s.addShape(...['roundRect', rr({ x: 5.333, y: 2.43, w: 6.91, h: 1.575, fill: { color: C.green } }, 0.1667)]);
  seg(s, 8.789, 2.628, 8.789, 3.809, C.white, 0.75);
  [[5.704, 6.674, 1.921, '32%'], [9.141, 10.11, 2.287, '68%']].forEach(([xa, xb, bw, val]) => {
    text(s, [run('Vs Miss Traffic', { fontFace: F.i, fontSize: 16, color: C.white })],
      { x: xa, y: 2.916, w: 1.176, h: 0.64 });
    text(s, [run(val, { fontFace: F.is, fontSize: 54, bold: true, color: C.white })],
      { x: xb, y: 2.72, w: bw, h: 1.01 });
  });

  caption(s, { x: 7.823, y: 4.35, w: 2.479, h: 0.343 }, 'Subtitle Goes Here', { face: F.ps, size: 14, ls: 1.21 });
  body(s, { x: 7.823, y: 4.657, w: 4.208, h: 1.027 }, T.loremLong);
  button(s, { x: 7.935, y: 6.024, w: 1.168, h: 0.361, label: 'More', line: C.mint, fill: C.white, shadow: SH.drop });
}

// 5 - job analysis ----------------------------------------------------
function slide05(s) {
  s.addShape('rect', { x: 0, y: 0, w: 3.683, h: 7.5, fill: { color: C.sage, transparency: 53 } });
  heading(s, { x: 0.963, y: 1.02, w: 6.0, h: 1.313, ls: 1.0 },
    [['Job Analysis ', C.green, F.p], ['and Job Description', C.black, F.ps]]);

  s.addShape(...['roundRect', pill({ x: 8.68, y: 3.043, w: 0.777, h: 3.346, fill: { color: C.green, transparency: 72 } })]);

  // day columns: [x, stem-top, dot-y, label, highlighted]
  const days = [
    [7.388, 4.407, 4.168, 'S', 0], [8.142, 3.936, 3.655, 'S', 0], [8.895, 3.22, null, 'M', 1],
    [9.643, 4.043, 3.752, 'T', 0], [10.402, 3.712, 3.289, 'W', 0],
    [11.15, 4.043, 3.717, 'T', 0], [11.909, 4.407, 3.971, 'F', 0],
  ];
  days.forEach(([x, top, dot, label, hi]) => {
    const cx = x + 0.156;
    seg(s, cx, top, cx, 5.676, hi ? '7F7F7F' : C.smoke, hi ? 0.75 : 1.25);
    if (dot !== null) s.addShape('ellipse', { x: cx - 0.072, y: dot, w: 0.144, h: 0.144, fill: { color: C.mint } });
    s.addShape('ellipse', { x, y: 5.915, w: 0.312, h: 0.312, fill: { color: hi ? C.forest : C.white } });
    text(s, [run(label, { fontFace: F.i, fontSize: 12, color: hi ? C.white : C.black })],
      { x, y: 5.915, w: 0.312, h: 0.312, align: 'center', valign: 'middle' });
  });

  s.addShape(...['roundRect', pill({ x: 8.68, y: 2.937, w: 0.777, h: 0.277, fill: { color: C.dark } })]);
  text(s, [run('+17.21%', { fontFace: F.i, fontSize: 10, color: C.white })],
    { x: 8.497, y: 2.954, w: 1.144, h: 0.269, align: 'center' });

  s.addShape('rect', { x: 1.112, y: 4.698, w: 2.571, h: 1.691, fill: { color: C.green } });
  text(s, [
    run('Executive', br({ fontFace: F.ps, fontSize: 16, color: C.white })),
    run('Summary ', { fontFace: F.ps, fontSize: 16, color: C.white }),
    run('where you want', { fontFace: F.pl, fontSize: 16, color: C.white }),
  ], { x: 1.418, y: 5.066, w: 2.376, h: 1.051, lineSpacingMultiple: 1.44 });
}

// 6 - making the job offer --------------------------------------------
function slide06(s) {
  s.addShape('rect', { x: 0.006, y: 1.099, w: 3.58, h: 6.401, fill: { color: C.mint } });
  heading(s, { x: 4.035, y: 1.741, w: 4.206, h: 1.313, ls: 1.0 },
    [['Making', C.black, F.ps, true], ['The Job offer', C.black, F.ps]]);

  // profile card
  s.addShape('rect', { x: 1.109, y: 3.51, w: 2.478, h: 2.89, fill: { color: C.dark } });
  s.addShape(...['roundRect', rr({ x: 1.404, y: 3.998, w: 0.633, h: 0.633, fill: { color: C.white } }, 0.1667)]);
  ICON.wallet(s, 1.563, 4.177, 0.317, C.green);
  text(s, [run('Minda Jack', { fontFace: F.im, fontSize: 16, color: C.white })],
    { x: 1.286, y: 5.061, w: 1.917, h: 0.33, lineSpacingMultiple: 1.0 });
  text(s, [run('CEO In Company', { fontFace: F.il, fontSize: 7, color: C.white })],
    { x: 1.289, y: 5.281, w: 1.405, h: 0.219 });
  s.addShape(...['roundRect', pill({ x: 1.404, y: 5.644, w: 1.799, h: 0.113, fill: { color: C.white } })]);
  s.addShape(...['roundRect', pill({ x: 1.404, y: 5.647, w: 1.279, h: 0.11, fill: { color: C.green } })]);
  text(s, [run('0%', { fontFace: F.i, fontSize: 8, color: C.white })], { x: 1.31, y: 5.826, w: 0.495, h: 0.236 });
  text(s, [run('100%', { fontFace: F.i, fontSize: 8, color: C.white })], { x: 2.656, y: 5.826, w: 0.62, h: 0.236, align: 'right' });

  // three feature rows
  [['briefcase', 1.149, 1.719], ['inbox', 3.095, 3.647], ['magnifier', 5.062, 5.583]]
    .forEach(([icon, iy, ty]) => {
      ICON[icon](s, 9.28, iy, 0.42, C.dark);
      caption(s, { x: 9.133, y: ty, w: 2.479, h: 0.343 }, T.applicant, { face: F.ps, size: 14, ls: 1.21 });
      body(s, { x: 9.133, y: ty + 0.318, w: 3.383, h: 0.522 }, T.loremShort);
    });
}

// 7 - prepared for the interview --------------------------------------
function slide07(s) {
  blob(s, { x: 1.261, y: 1.191, w: 5.302, h: 5.003, fill: { color: C.white } });
  blob(s, { x: 1.609, y: 1.497, w: 4.654, h: 4.392, fill: { color: C.green } });

  s.addShape(...['roundRect', pill({ x: 2.277, y: 5.466, w: 3.361, h: 0.933, fill: { color: C.dark } })]);
  text(s, [run('James Williams', { fontFace: F.rm, fontSize: 20, color: C.white })],
    { x: 2.231, y: 5.661, w: 3.453, h: 0.438, align: 'center' });
  text(s, [run('CEO, INN company', { fontFace: F.i, fontSize: 9, color: 'F2F2F2' })],
    { x: 2.499, y: 6.0, w: 2.916, h: 0.252, align: 'center' });

  heading(s, { x: 7.26, y: 1.138, w: 6.0, h: 1.313, ls: 1.0 },
    [['Prepared', C.green, F.pl], [' for the interview', C.black, F.ps]]);
  arrowUp(s, 7.371, 2.494, 0.449, 0.573, C.green);
  caption(s, { x: 8.095, y: 2.547, w: 2.479, h: 0.573 }, T.execSmarmy, { face: F.ps, size: 14, ls: 1.21 });
  body(s, { x: 8.123, y: 3.123, w: 3.909, h: 0.976 }, T.loremLong, { ls: 1.4 });

  iconBadge(s, 8.238, 4.718, 0.528, { bg: C.white, fg: C.black, icon: 'layers', shadow: SH.lift, scale: 0.42 });
  caption(s, { x: 9.055, y: 4.641, w: 1.518, h: 0.303 }, 'Background ');
  body(s, { x: 9.055, y: 4.893, w: 2.628, h: 0.48 }, T.mistakenOf, { size: 9 });
  iconBadge(s, 8.238, 5.841, 0.528, { bg: C.green, fg: C.white, icon: 'lamp', shadow: SH.lift, scale: 0.45 });
  caption(s, { x: 9.055, y: 5.765, w: 1.518, h: 0.303 }, 'Capabilities');
  body(s, { x: 9.055, y: 6.017, w: 2.628, h: 0.48 }, T.mistakenAnd, { size: 9 });
}

// 8 - onboarding ------------------------------------------------------
function slide08(s) {
  blob(s, { x: 4.548, y: 2.401, w: 4.334, h: 4.089, fill: { color: C.white, transparency: 56 } });
  blob(s, { x: 4.763, y: 2.489, w: 3.968, h: 3.744, fill: { color: C.green } });
  s.addShape('ellipse', { x: 5.662, y: 3.081, w: 2.249, h: 2.249, fill: { color: C.white, transparency: 83 } });

  heading(s, { x: 2.354, y: 1.01, w: 8.624, h: 0.663, ls: 1.11, align: 'center' },
    [['Onboarding', C.green, F.pl], [' And Orientation', C.black, F.ps]]);

  // ring + dark badge, icon, and the caption block that belongs to it
  const nodes = [
    [4.267, 2.857, 'briefcase', 0.989, 2.659, 1.204, 3.009, 'right'],
    [8.374, 2.834, 'layers', 9.927, 2.659, 9.927, 3.009, 'left'],
    [5.036, 5.164, 'gem', 0.989, 5.031, 1.204, 5.381, 'right'],
    [7.753, 5.164, 'globe', 9.927, 5.024, 9.927, 5.375, 'left'],
  ];
  nodes.forEach(([bx, by, icon, tx, ty, cx, cy, al]) => {
    s.addShape('ellipse', { x: bx, y: by, w: 0.777, h: 0.777, fill: { color: C.cream } });
    iconBadge(s, bx + 0.059, by + 0.059, 0.658, { bg: C.dark, fg: C.white, icon, scale: 0.47 });
    caption(s, { x: tx, y: ty, w: 2.633, h: 0.37 }, 'Subtitle Here', { size: 16, align: al });
    body(s, { x: cx, y: cy, w: 2.418, h: 0.74 }, T.mistaken, { ls: 1.4, align: al });
  });
}

// 9 - shrining and shortlisting ---------------------------------------
function slide09(s) {
  s.addShape('rect', { x: 0, y: 0, w: 13.333, h: 3.826, fill: { color: C.dark } });
  heading(s, { x: 2.354, y: 1.009, w: 8.624, h: 0.663, ls: 1.11, align: 'center' },
    [['Shrining', C.white, F.pl], [' and Shortlisting ', C.white, F.ps]]);

  seg(s, 4.441, 3.826, 5.002, 3.826, C.green, 4, 'sysDot');
  seg(s, 8.395, 3.826, 8.956, 3.826, C.green, 4, 'sysDot');

  [[1.734, 2.483, 1.282, 1.389, '01', 'Project'],
  [5.66, 6.409, 5.243, 5.35, '02', 'Mission'],
  [9.586, 10.335, 9.303, 9.411, '03', 'Company']].forEach(([hx, nx, tx, bx, num, title]) => {
    s.addShape('hexagon', { x: hx, y: 2.737, w: 2.543, h: 2.193, fill: { color: C.white } });
    s.addShape('flowChartConnector', { x: nx, y: 2.383, w: 0.386, h: 0.386, fill: { color: C.green }, shadow: SH.soft });
    text(s, [run(num, { fontFace: F.a, fontSize: 10, color: C.white })],
      { x: nx, y: 2.383, w: 0.386, h: 0.386, align: 'center', valign: 'middle', margin: 0 });
    caption(s, { x: tx, y: 5.44, w: 2.633, h: 0.37 }, title, { size: 16, align: 'center' });
    body(s, { x: bx, y: 5.771, w: 2.525, h: 0.707 }, T.mistaken, { size: 9, align: 'center' });
  });
}

// 10 - interview as a service -----------------------------------------
function slide10(s) {
  heading(s, { x: 0.962, y: 0.981, w: 5.892, h: 0.663, ls: 1.11 },
    [['Interview', C.green, F.p], [' As a service', C.black, F.ps]]);

  const cards = [
    [1.107, C.white, C.black, C.pale, 'magnifier', SH.card],
    [4.987, C.green, C.white, C.white, 'briefcase', SH.wide],
    [8.856, C.white, C.black, C.pale, 'lock', SH.card],
  ];
  cards.forEach(([x, bg, fg, ring, icon, shadow]) => {
    s.addShape(...['roundRect', rr({ x, y: 2.472, w: 3.36, h: 3.929, fill: { color: bg }, shadow }, 0.0698)]);
    s.addShape('ellipse', { x: x + 0.58, y: 3.009, w: 0.817, h: 0.817, fill: { color: ring } });
    iconBadge(s, x + 0.725, 3.154, 0.527, { bg: C.green, fg: C.white, icon, shadow: SH.lift, scale: 0.44 });
    caption(s, { x: x + 0.483, y: 3.959, w: 2.762, h: 0.583 }, T.constituter, { color: fg, ls: 1.5 });
    body(s, { x: x + 0.483, y: 4.582, w: 2.616, h: 0.707 }, T.variations, { size: 9, color: fg });
    text(s, [run('24H', { fontFace: F.ps, fontSize: 32, color: fg })],
      { x: x + 0.476, y: 5.407, w: 2.479, h: 0.64 });
  });
}

// 11 - making the job process -----------------------------------------
function slide11(s) {
  heading(s, { x: 0.961, y: 0.988, w: 7.129, h: 0.663, ls: 1.11 },
    [['Making The ', C.green, F.pl], ['Job Process', C.black, F.ps]]);

  text(s, [
    run('Write Your ', { fontFace: F.il, fontSize: 18, color: C.black }),
    run('Executive ', { fontFace: F.is, fontSize: 18, color: C.black }),
    run('Smarmy off where you want to take your ', { fontFace: F.il, fontSize: 18, color: C.black }),
    run('Business', { fontFace: F.is, fontSize: 18, color: C.black, underline: { style: 'sng' } }),
  ], { x: 0.886, y: 4.158, w: 3.628, h: 0.942, rotate: 270, lineSpacingMultiple: 1.11 });
  seg(s, 0, 6.288, 1.833, 6.288, C.dark, 3.5);

  s.addShape(...['roundRect', rr({ x: 4.381, y: 2.234, w: 7.85, h: 4.167, fill: { color: C.white }, shadow: SH.card }, 0.0729)]);
  [[2.859, 'gem'], [4.733, 'layers']].forEach(([y, icon]) => {
    s.addShape('ellipse', { x: 7.371, y, w: 0.934, h: 0.934, fill: { color: C.green, transparency: 55 } });
    iconBadge(s, 7.454, y + 0.094, 0.768, { bg: C.green, fg: C.white, icon, shadow: SH.soft, scale: 0.46 });
    text(s, [run('Earning', { fontFace: F.im, fontSize: 14, color: C.black })],
      { x: 9.064, y: y - 0.273, w: 1.077, h: 0.33, lineSpacingMultiple: 1.14 });
    text(s, [run('$20.5K', { fontFace: F.il, fontSize: 40, color: C.black })],
      { x: 9.048, y: y + 0.027, w: 2.825, h: 0.774 });
    s.addShape(...['roundRect', pill({ x: 9.171, y: y + 0.867, w: 2.479, h: 0.399, fill: { color: C.green } })]);
    text(s, [run('+12% since last month', { fontFace: F.i, fontSize: 12, color: C.white })],
      { x: 9.213, y: y + 0.926, w: 2.331, h: 0.303, align: 'center' });
  });
}

// 12 - evaluating culture fit -----------------------------------------
function slide12(s) {
  s.addShape('rect', { x: 0, y: 0, w: 13.25, h: 3.43, fill: { color: C.dark } });
  heading(s, { x: 2.354, y: 0.812, w: 8.624, h: 0.663, ls: 1.11, align: 'center' },
    [['Evaluating', C.white, F.p], [' culture fit', C.white, F.ps]]);

  [[2.16, 1.282, 'gem', 'Project'], [6.035, 5.449, 'layers', 'Mission'], [9.91, 9.303, 'cloud', 'Company']]
    .forEach(([cx, tx, icon, title]) => {
      s.addShape('ellipse', { x: cx, y: 3.96, w: 1.263, h: 1.263, fill: { color: C.white, transparency: 55 } });
      iconBadge(s, cx + 0.112, 4.083, 1.039, { bg: C.green, fg: C.white, icon, scale: 0.48 });
      caption(s, { x: tx, y: 5.458, w: 2.633, h: 0.37 }, title, { size: 16, align: 'center' });
      body(s, { x: tx, y: 5.832, w: 2.633, h: 0.713 }, T.mistaken, { face: F.a, align: 'center' });
    });
}

// 13 - quote ----------------------------------------------------------
function slide13(s) {
  s.addShape('rect', { x: 1.101, y: 1.099, w: 11.132, h: 5.302, fill: { color: C.dark } });
  ICON.magnifier(s, 2.038, 1.945, 2.85, '4E6560', 18);

  text(s, [
    run('\u201cNow there is a way to people who already ', { fontFace: F.p, fontSize: 32, color: C.white }),
    run('looking something, you offer Our relevance', { fontFace: F.p, fontSize: 32, color: C.green }),
  ], { x: 3.488, y: 2.821, w: 7.512, h: 1.616, lineSpacingMultiple: 1.125 });

  button(s, { x: 3.636, y: 4.667, w: 1.543, h: 0.508, label: 'See More', line: C.green, fill: C.white, size: 14, color: C.gray, shadow: SH.drop });
  text(s, [run('INTERVIEW', { fontFace: F.is, fontSize: 88, bold: true, color: C.forest })],
    { x: 5.87, y: 5.174, w: 8.505, h: 1.582 });
}

// 14 - reference background verification -------------------------------
function slide14(s) {
  s.addShape('rect', { x: 0, y: 4.236, w: 13.333, h: 3.264, fill: { color: C.dark } });
  chevronBadge(s, 3.863, 1.076, 0.303, C.green);
  text(s, [run("Let's Get\u2019s Stated ", { fontFace: F.r, fontSize: 12, color: C.gray })],
    { x: 4.204, y: 1.076, w: 1.695, h: 0.303 });
  heading(s, { x: 3.734, y: 1.481, w: 7.643, h: 1.313, ls: 1.0 },
    [['Reference ', C.green, F.p], ['background verification ', C.black, F.ps]]);

  [[3.828, 'magnifier'], [6.951, 'inbox'], [10.074, 'briefcase']].forEach(([x, icon]) => {
    iconBadge(s, x, 3.761, 1.054, { bg: C.white, fg: C.black, icon, scale: 0.4 });
    caption(s, { x: x - 0.094, y: 5.076, w: 2.761, h: 0.566 }, T.constituter, { color: C.white, ls: 1.3 });
    body(s, { x: x - 0.094, y: 5.682, w: 2.616, h: 0.707 }, T.variations, { size: 9, color: C.white });
  });
}

// 15 - question & technique -------------------------------------------
function slide15(s) {
  heading(s, { x: 2.325, y: 1.007, w: 8.624, h: 0.663, ls: 1.11, align: 'center' },
    [['Question', C.green, F.pl], [' & Technique ', C.black, F.ps]]);
  seg(s, 6.635, 2.912, 6.635, 6.401, C.silver, 0.75);

  text(s, [run('01 ', { fontFace: F.pl, fontSize: 48, color: C.black })], { x: 3.047, y: 2.65, w: 2.75, h: 0.909, align: 'right' });
  text(s, [run('02 ', { fontFace: F.pl, fontSize: 48, color: C.black })], { x: 7.578, y: 2.65, w: 2.75, h: 0.909 });
  button(s, { x: 3.817, y: 2.912, w: 1.014, h: 0.313, label: 'More', fill: C.white, line: C.mint, size: 10, shadow: SH.drop });
  button(s, { x: 8.867, y: 2.912, w: 1.014, h: 0.313, label: 'More', fill: C.white, line: C.mint, size: 10, shadow: SH.drop });

  // left column: track + green bar growing from the right, knob at far right
  const left = [[3.874, 1.996, '90%', 'Question #01', 'gear'],
  [4.842, 2.527, '70%', 'Question #02', 'layers'],
  [5.831, 2.243, '78%', 'Question #01', 'scissors']];
  left.forEach(([y, bx, pct, label, icon]) => {
    s.addShape(...['roundRect', pill({ x: 1.17, y, w: 4.58, h: 0.526, fill: { color: C.white }, shadow: SH.soft })]);
    text(s, [run(pct, { fontFace: F.i, fontSize: 10, color: '44546A' })],
      { x: 1.17, y, w: 4.58, h: 0.526, valign: 'middle', margin: [14.4, 0, 0, 0] });
    s.addShape(...['roundRect', pill({ x: bx, y, w: 5.749 - bx, h: 0.526, fill: { color: C.green }, shadow: SH.soft })]);
    text(s, [run(label, { fontFace: F.i, fontSize: 10, color: C.white })],
      { x: bx, y, w: 5.749 - bx, h: 0.526, align: 'right', valign: 'middle', margin: [0, 43.2, 0, 0] });
    s.addShape('flowChartConnector', { x: 5.224, y, w: 0.526, h: 0.526, fill: { color: C.white }, shadow: SH.soft });
    ICON[icon](s, 5.373, y + 0.149, 0.228, C.black);
  });

  // right column: mirrored - dark bar grows from the left
  const right = [[3.874, 2.75, '50%', 'Technique #01', 'gear'],
  [4.842, 3.812, '90%', 'Technique #02', 'layers'],
  [5.831, 3.493, '78%', 'Technique #03', 'scissors']];
  right.forEach(([y, bw, pct, label, icon]) => {
    s.addShape(...['roundRect', pill({ x: 7.651, y, w: 4.58, h: 0.526, fill: { color: C.white }, shadow: SH.soft })]);
    text(s, [run(pct, { fontFace: F.i, fontSize: 10, bold: true, color: '44546A' })],
      { x: 7.651, y, w: 4.58, h: 0.526, align: 'right', valign: 'middle', margin: [0, 14.4, 0, 0] });
    s.addShape(...['roundRect', pill({ x: 7.628, y, w: bw, h: 0.526, fill: { color: C.dark }, shadow: SH.soft })]);
    text(s, [run(label, { fontFace: F.i, fontSize: 10, bold: true, color: C.white })],
      { x: 7.628, y, w: bw, h: 0.526, valign: 'middle', margin: [43.2, 14.4, 0, 0] });
    s.addShape('flowChartConnector', { x: 7.628, y, w: 0.526, h: 0.526, fill: { color: C.white }, shadow: SH.soft });
    ICON[icon](s, 7.776, y + 0.149, 0.228, C.black);
  });
}

// 16 - preparing for the interview (green panel) -----------------------
function slide16(s) {
  s.addShape('rect', { x: 0, y: 0, w: 4.917, h: 7.5, fill: { color: C.green } });

  heading(s, { x: 7.331, y: 1.099, w: 6.0, h: 1.313, ls: 1.0 },
    [['Preparing for ', C.black, F.ps], ['the interview', C.green, F.p]]);
  arrowUp(s, 7.442, 2.548, 0.449, 0.573, C.green);
  caption(s, { x: 8.291, y: 2.601, w: 2.479, h: 0.573 }, T.execSmarmy, { face: F.ps, size: 14, ls: 1.21 });
  text(s, [
    run('Lorem ipsum\u00a0', { fontFace: F.i, fontSize: 10, bold: true, color: C.gray }),
    run('is simply dummy text of ', { fontFace: F.i, fontSize: 10, color: C.gray }),
    run('the printing and typesetting ', { fontFace: F.i, fontSize: 10, bold: true, color: C.gray }),
    run("industry. Lorem ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer", { fontFace: F.i, fontSize: 10, color: C.gray }),
  ], { x: 8.319, y: 3.206, w: 3.654, h: 0.926, lineSpacingMultiple: 1.3 });

  iconBadge(s, 8.434, 4.533, 0.528, { bg: C.white, fg: C.dark, icon: 'layers', shadow: SH.lift, scale: 0.42 });
  caption(s, { x: 9.251, y: 4.457, w: 1.518, h: 0.303 }, 'Background ');
  body(s, { x: 9.251, y: 4.737, w: 2.866, h: 0.449 }, T.mistakenOf, { size: 9, ls: 1.3 });
  iconBadge(s, 8.434, 5.583, 0.528, { bg: C.green, fg: C.white, icon: 'lamp', shadow: SH.lift, scale: 0.45 });
  caption(s, { x: 9.251, y: 5.507, w: 1.518, h: 0.303 }, 'Capabilities');
  body(s, { x: 9.251, y: 5.787, w: 2.866, h: 0.449 }, T.mistakenAnd, { size: 9, ls: 1.3 });
}

// 17 - preparing for the interview (circle) ----------------------------
function slide17(s) {
  poly(s, CORNER, { x: 0, y: 0, w: 6.559, h: 4.679, fill: { color: C.pale } });
  s.addShape('ellipse', { x: 1.569, y: 1.625, w: 4.543, h: 4.543, fill: { color: C.cream } });
  s.addShape('ellipse', { x: 4.097, y: 4.555, w: 1.754, h: 1.754, fill: { color: C.white, transparency: 70 } });
  s.addShape('ellipse', { x: 4.253, y: 4.725, w: 1.442, h: 1.442, fill: { color: C.green } });
  text(s, [run('35%', { fontFace: F.pm, fontSize: 24, color: C.white })],
    { x: 4.253, y: 4.725, w: 1.442, h: 1.442, align: 'center', valign: 'middle' });
  squiggle(s, 0.679, 6.231, 0.887, 0.334);

  chevronBadge(s, 7.484, 1.489, 0.324, C.dark);
  text(s, [run('About Us', { fontFace: F.r, fontSize: 14, color: C.gray })],
    { x: 7.875, y: 1.499, w: 2.382, h: 0.337 });
  heading(s, { x: 7.372, y: 1.919, w: 6.0, h: 1.313, ls: 1.0 },
    [['Preparing for ', C.black, F.ps], ['the interview', C.green, F.p]]);
  arrowUp(s, 7.484, 3.303, 0.449, 0.573, C.green);
  caption(s, { x: 8.208, y: 3.355, w: 2.479, h: 0.573 }, T.execSmarmy, { face: F.ps, size: 14, ls: 1.21 });
  body(s, { x: 7.381, y: 4.057, w: 4.719, h: 0.976 },
    'Designers generally have nothing to do with creating content for their projects. Even so,  site can be incomplete Designers generally have nothing to do with creating content for them. Designers generally have nothing to do with', { ls: 1.4 });
  body(s, { x: 7.381, y: 5.094, w: 4.719, h: 0.505 },
    'Designers generally have nothing to do with creating content for their projects. Even so,  site can', { ls: 1.4, color: C.black });
  button(s, { x: 7.478, y: 5.942, w: 1.506, h: 0.458, label: 'See More', fill: C.white, line: C.mint, size: 14, shadow: SH.drop });
}

// 18 - table ----------------------------------------------------------
function slide18(s) {
  s.addShape('rect', { x: 1.099, y: 1.099, w: 11.132, h: 5.302, fill: { color: C.green } });
  heading(s, { x: 3.102, y: 1.475, w: 7.129, h: 0.663, ls: 1.11 },
    [['The Interview ', C.white, F.pl], ['Table Format', C.white, F.ps]]);

  s.addShape('rect', { x: 1.94, y: 2.476, w: 9.522, h: 0.632, fill: { color: C.dark } });
  s.addShape('rect', { x: 1.926, y: 2.475, w: 9.522, h: 3.395, fill: { type: 'none' }, line: { color: C.white, width: 1, transparency: 7 } });

  // vertical column rules + horizontal row rules
  [[4.069, 2.476, 5.981], [5.867, 2.476, 5.869], [7.723, 3.108, 4.284], [9.579, 2.475, 3.114],
  [7.723, 4.799, 5.87]].forEach(([x, y0, y1]) => seg(s, x, y0, x, y1, C.white, 0.75, 'solid'));
  [3.706, 4.285, 4.799, 5.366].forEach((y) => seg(s, 1.91, y, 11.424, y, C.white, 0.75));

  // header row
  text(s, [run('Tags', { fontFace: F.ps, fontSize: 14, color: C.white })], { x: 2.213, y: 2.649, w: 1.685, h: 0.337 });
  text(s, [run('90:14-10:20 AM', { fontFace: F.pm, fontSize: 12, color: C.white })], { x: 3.926, y: 2.676, w: 2.078, h: 0.303, align: 'center' });
  text(s, [run('Break', { fontFace: F.pm, fontSize: 18, color: C.white })], { x: 6.71, y: 2.603, w: 2.078, h: 0.404, align: 'center' });
  text(s, [run('90:14-10:20 AM', { fontFace: F.pm, fontSize: 12, color: C.white })], { x: 9.502, y: 2.676, w: 2.078, h: 0.303, align: 'center' });

  // body rows: label + 4 numeric columns (null = covered by a pill)
  const rows = [
    ['Income Tax', 3.327, ['00.00', '00.00', '00.00', '00.00']],
    ['Costing', 3.858, ['10.00', '00.00', '00.00', '00.00']],
    ['Agency', 4.397, ['20.00', null, null, '00.00']],
    ['Product Edit', 4.938, ['30.00', '00.00', '00.00', '00.00']],
    ['Sales', 5.44, ['40.00', '00.00', null, null]],
  ];
  const colX = [4.418, 6.293, 8.169, 10.006];
  rows.forEach(([label, y, vals]) => {
    text(s, [run(label, { fontFace: F.pm, fontSize: 12, color: C.white })], { x: 2.213, y: y === 4.397 ? 4.396 : y, w: 1.685, h: 0.303 });
    vals.forEach((v, i) => {
      if (v === null) return;
      text(s, [run(v, { fontFace: F.p, fontSize: 12, color: C.white })], { x: colX[i], y, w: 1.07, h: 0.303, align: 'center' });
    });
  });

  // merged "pill" cells
  [[6.382, 4.347, 2.783, 0.415, 'User permissions', 6.601, 4.408, 2.188],
  [7.938, 5.419, 3.306, 0.415, 'A/B testing & feature flags', 8.126, 5.48, 2.983]]
    .forEach(([x, y, w, h, label, tx, ty, tw]) => {
      s.addShape(...['roundRect', pill({ x, y, w, h, fill: { color: C.dark } })]);
      text(s, [run(label, { fontFace: F.pm, fontSize: 12, color: C.white })],
        { x: tx, y: ty, w: tw, h: 0.4, align: 'center' });
    });
}

// 19 - meet our team --------------------------------------------------
function slide19(s) {
  heading(s, { x: 0.991, y: 0.804, w: 4.123, h: 1.313, ls: 1.0 },
    [['Meet', C.black, F.ps, true], ['Our ', C.black, F.ps], ['Team', C.green, F.p]]);
  text(s, [
    run('Write Your ', { fontFace: F.il, fontSize: 18, color: C.black }),
    run('Executive ', { fontFace: F.is, fontSize: 18, color: C.black }),
    run('Smarmy off where you want to take your ', { fontFace: F.il, fontSize: 18, color: C.black }),
    run('Business', { fontFace: F.is, fontSize: 18, color: C.black, underline: { style: 'sng' } }),
  ], { x: -0.315, y: 4.382, w: 3.628, h: 0.942, rotate: 270, lineSpacingMultiple: 1.11 });

  const team = [
    [3.27, 3.128, C.white, C.black, C.dark, 'Robert Antonin', 'UX Designer', 3.217, 3.609, 6.067],
    [6.575, 6.426, C.green, C.white, C.white, 'Kevin Pollard', 'UI Designer', 6.523, 6.914, 9.277],
    [9.724, 9.575, C.white, C.black, C.dark, 'Christopher Zain', 'WEB Designer', 9.672, 10.064, null],
  ];
  team.forEach(([cx, ex, bg, nameC, roleC, name, role, nx, rx, sx]) => {
    s.addShape(...['roundRect', pill({ x: cx, y: 2.483, w: 2.526, h: 4.109, fill: { color: bg }, shadow: SH.card })]);
    s.addShape('ellipse', { x: ex, y: 2.326, w: 2.824, h: 2.824, fill: { color: C.cream } });
    text(s, [run(name, { fontFace: F.pm, fontSize: 18, color: nameC })], { x: nx, y: 5.462, w: 2.63, h: 0.404, align: 'center' });
    text(s, [run(role, { fontFace: F.i, fontSize: 12, color: roleC })], { x: rx, y: 5.76, w: 1.847, h: 0.303, align: 'center' });
    // three small social-media chips between the cards
    if (sx !== null) {
      [4.034, 4.435, 4.838].forEach((y) =>
        s.addShape('ellipse', { x: sx, y, w: 0.214, h: 0.214, fill: { color: C.smoke } }));
    }
  });
}

// 20 - continues improvement -------------------------------------------
function slide20(s) {
  poly(s, SWOOSH, { x: 0, y: 1.411, w: 13.333, h: 6.105, fill: { color: 'EAF2ED', transparency: 55 } });
  s.addShape('rect', { x: 3.16, y: 2.722, w: 10.174, h: 4.778, fill: { color: C.green } });
  heading(s, { x: 1.035, y: 0.999, w: 6.878, h: 1.313, ls: 1.0 },
    [['Continues improvement ', C.green, F.p], ['And Best Practices', C.black, F.ps]]);

  arrowUp(s, 5.194, 3.261, 0.613, 0.782, C.white);
  caption(s, { x: 5.977, y: 3.121, w: 2.368, h: 0.572 }, T.execSmarmy, { face: F.ps, size: 14, color: C.white, ls: 1.0 });
  text(s, [
    run('Lorem ipsum\u00a0', { fontFace: F.i, fontSize: 10, bold: true, color: C.white }),
    run('is simply dummy text of ', { fontFace: F.i, fontSize: 10, color: C.white }),
    run('the printing and typesetting ', { fontFace: F.i, fontSize: 10, bold: true, color: C.white }),
    run("industry. Lorem ipsum has been the industry's standard dummy text ever since the", { fontFace: F.i, fontSize: 10, color: C.white }),
  ], { x: 6.006, y: 3.64, w: 6.314, h: 0.505, lineSpacingMultiple: 1.4 });

  const cells = [
    [5.133, 4.628, 6.005, 4.597, 'lamp', 'Capabilities', T.mistakenAnd],
    [9.17, 4.63, 10.043, 4.597, 'layers', 'Background ', T.mistakenOf],
    [5.133, 5.751, 6.005, 5.721, 'gear', 'Capabilities', T.mistakenAnd],
    [9.17, 5.753, 10.043, 5.721, 'cloud', 'Background ', T.mistakenOf],
  ];
  cells.forEach(([ix, iy, tx, ty, icon, title, copy]) => {
    iconBadge(s, ix, iy, 0.638, { ring: C.white, ringW: 1, fg: C.white, hole: C.green, icon, scale: 0.45 });
    caption(s, { x: tx, y: ty, w: 1.518, h: 0.303 }, title, { color: C.white });
    body(s, { x: tx, y: ty + 0.252, w: 2.29, h: 0.465 }, copy, { size: 9, color: C.white, ls: 1.4 });
  });
}

// 21 - price table ----------------------------------------------------
function slide21(s) {
  heading(s, { x: 2.381, y: 1.0, w: 8.57, h: 0.663, ls: 1.11, align: 'center' },
    [['Price Table for ', C.green, F.p], ['new member', C.black, F.ps]]);

  const plans = [
    { x: 1.094, name: 'Basic ', bg: C.white, fg: C.black, dot: C.pale, btn: C.dark, shadow: SH.card },
    { x: 5.049, name: 'Standard ', bg: C.green, fg: C.white, dot: C.ink, btn: C.white, shadow: SH.wide },
    { x: 8.877, name: 'Premium ', bg: C.white, fg: C.black, dot: C.pale, btn: C.dark, shadow: SH.lift },
  ];
  plans.forEach((p) => {
    s.addShape(...['roundRect', rr({ x: p.x, y: 2.319, w: 3.354, h: 4.081, fill: { color: p.bg }, shadow: p.shadow }, 0.1192)]);
    text(s, [run(p.name, { fontFace: F.p, fontSize: 12, color: p.fg })], { x: p.x + 0.27, y: 2.778, w: 1.427, h: 0.303 });
    text(s, [
      run('$9.50', { fontFace: F.pm, fontSize: 40, color: p.fg }),
      run('/ ', { fontFace: F.pm, fontSize: 20, color: p.fg }),
      run('Month', { fontFace: F.pm, fontSize: 12, color: p.fg }),
    ], { x: p.x + 0.214, y: 3.19, w: 2.662, h: 0.774 });
    [4.303, 4.709, 5.131].forEach((y) =>
      s.addShape('ellipse', { x: p.x + 0.369, y, w: 0.243, h: 0.243, fill: { color: p.dot, transparency: p.dot === C.ink ? 45 : 0 } }));
    text(s, ['   There are many ', '   Injected humor, ', '   Believable'].map((t, i) => run(t, {
      fontFace: F.i, fontSize: 12, color: p.fg, breakLine: i < 2,
      bullet: { characterCode: '2714', indent: 13.5 },
    })), { x: p.x + 0.369, y: 4.243, w: 2.178, h: 1.192, lineSpacingMultiple: 2.2 });
    button(s, { x: p.x + 0.414, y: 5.727, w: 2.381, h: 0.383, label: 'Choose plan', line: p.btn, size: 14, face: F.im, color: p.fg });
  });
}

// 22 - S.W.O.T --------------------------------------------------------
function slide22(s) {
  heading(s, { x: 2.381, y: 1.0, w: 8.57, h: 0.663, ls: 1.11, align: 'center' },
    [['S.W.O.T ', C.green, F.p], ['Analysis', C.black, F.ps]]);

  // large rotated "X" behind the four quadrant badges (arms are 36.5% of the side)
  const t = (1 - 0.365) / 2;
  poly(s, [[t, 0, 1], [1 - t, 0], [1 - t, t], [1, t], [1, 1 - t], [1 - t, 1 - t],
  [1 - t, 1], [t, 1], [t, 1 - t], [0, 1 - t], [0, t], [t, t], []],
    { x: 4.904, y: 2.549, w: 3.912, h: 3.912, rotate: 315, fill: { color: '3F6A55' } });

  const quads = [
    [4.828, 2.489, 'magnifier', 'Strength', 1.88, 2.642, 1.269, 2.998, 'right'],
    [7.723, 2.489, 'inbox', 'Opportunities', 9.444, 2.642, 9.444, 2.998, 'left'],
    [4.869, 5.347, 'layers', 'Weakness', 2.284, 5.235, 1.269, 5.591, 'right'],
    [7.723, 5.347, 'briefcase', 'Threats', 9.444, 5.235, 9.444, 5.591, 'left'],
  ];
  quads.forEach(([bx, by, icon, title, tx, ty, cx, cy, al]) => {
    s.addShape('ellipse', { x: bx, y: by, w: 1.161, h: 1.161, fill: { color: C.cream } });
    iconBadge(s, bx + 0.072, by + 0.072, 1.018, { bg: C.green, fg: C.white, icon, scale: 0.42 });
    caption(s, { x: tx, y: ty, w: al === 'right' ? 2.173 : 2.173, h: 0.353 }, title, { face: F.ps, size: 16, align: al, ls: 1.125 });
    text(s, [
      run('Designers', { fontFace: F.a, fontSize: 10, bold: true, color: C.gray }),
      run(' generally have nothing to do with creating content for their projects. Even so,  site can be', { fontFace: F.a, fontSize: 10, color: C.gray }),
    ], { x: cx, y: cy, w: 2.784, h: 0.74, align: al, lineSpacingMultiple: 1.4 });
  });
}

// 23 - interview infographics (interview scene) ------------------------
function slide23(s) {
  seg(s, 5.562, 6.375, 12.307, 6.375, C.dark, 1.5);
  artwork(s, 5.922, 1.681, 6.303, 4.694, '[illustration - interview scene]', true);

  heading(s, { x: 0.959, y: 1.028, w: 5.178, h: 1.126, ls: 1.0 },
    [['Interview', C.green, F.p], [' ', C.dark, F.p], ['Infographics', C.black, F.ps]]);
  text(s, [
    run('Write your\u00a0', { fontFace: F.pl, fontSize: 20, color: C.black }),
    run('executive', br({ fontFace: F.ps, fontSize: 20, color: C.black })),
    run('Summary ', { fontFace: F.ps, fontSize: 20, color: C.black }),
    run('of where you want to take', { fontFace: F.pl, fontSize: 20, color: C.silver }),
  ], { x: 0.978, y: 3.74, w: 4.068, h: 1.111, lineSpacingMultiple: 1.0 });

  [[1.109, C.green, C.white, C.white, C.black, '70%', SH.wide],
  [3.329, C.white, C.black, C.green, C.white, '30%', SH.lift]]
    .forEach(([x, bg, fg, btnBg, btnFg, pct, shadow]) => {
      s.addShape(...['roundRect', rr({ x, y: 5.119, w: 1.907, h: 1.305, fill: { color: bg }, shadow }, 0.1667)]);
      text(s, [run('Up To', { fontFace: F.i, fontSize: 16, color: fg })],
        { x: x + 0.101, y: 5.357, w: 0.643, h: 0.55, align: 'center', lineSpacingMultiple: 1.0 });
      text(s, [run(pct, { fontFace: F.is, fontSize: 28, color: fg })],
        { x: x + 0.389, y: 5.303, w: 1.569, h: 0.572, align: 'center' });
      s.addShape(...['roundRect', pill({ x: x + 0.275, y: 5.895, w: 1.351, h: 0.343, fill: { color: btnBg } })]);
      text(s, [run('Claim', { fontFace: F.is, fontSize: 11, color: btnFg }), run(' Here', { fontFace: F.i, fontSize: 11, color: btnFg })],
        { x: x + 0.401, y: 5.933, w: 1.058, h: 0.286, align: 'center' });
    });
}

// 24 - interview infographics (desk scene) ------------------------------
function slide24(s) {
  s.addShape('ellipse', { x: 4.982, y: 5.066, w: 7.44, h: 1.573, fill: { color: C.mint, transparency: 76 } });
  artwork(s, 5.5, 1.115, 6.7, 4.4, '[illustration - desk interview]', true);

  heading(s, { x: 0.959, y: 1.064, w: 3.746, h: 1.126, ls: 1.0 },
    [['Interview', C.green, F.p], [' ', C.dark, F.p], ['Infographics', C.black, F.ps]]);
  seg(s, 1.703, 3.289, 5.416, 3.289, C.dark, 1.25);
  seg(s, 1.703, 3.289, 1.703, 4.458, C.dark, 1.25);
  text(s, [run('89%', { fontFace: F.i, fontSize: 54, color: C.black })], { x: 0.98, y: 4.59, w: 2.928, h: 1.01 });
  body(s, { x: 0.98, y: 5.526, w: 2.928, h: 0.996 },
    'Constituter Adipescent Elastoses Daimio Nonarmy Nib Eui mod', { size: 14, color: C.black, ls: 1.4 });
}

// 25 - country maps ---------------------------------------------------
function slide25(s) {
  heading(s, { x: 2.833, y: 1.053, w: 7.667, h: 0.707, ls: 1.0, align: 'center' },
    [['Different', C.green, F.p], [' ', C.dark, F.p], ['Country Map', C.black, F.ps]]);
  seg(s, 6.667, 2.347, 6.667, 6.401, C.dark, 0.75);

  artwork(s, 1.617, 2.443, 3.477, 2.711, '[map - China]', true);
  artwork(s, 8.114, 2.112, 2.909, 2.936, '[map - North America]', true);

  // value pins
  [[2.62, 3.093, C.dark, '90%'], [3.751, 4.011, C.green, '68%'],
  [8.8, 2.35, C.dark, '90%'], [9.4, 3.15, C.gray, '60%'], [10.0, 3.75, C.green, '23%']]
    .forEach(([x, y, color, label]) => {
      s.addShape('blockArc', { x, y, w: 0.456, h: 0.456, angleRange: [0, 180], arcThicknessRatio: 0.35, fill: { color }, shadow: SH.down });
      s.addShape('ellipse', { x: x + 0.021, y: y + 0.023, w: 0.412, h: 0.412, fill: { color: C.white }, shadow: SH.down });
      text(s, [run(label, { fontFace: F.is, fontSize: 7, color: C.black })],
        { x: x + 0.021, y: y + 0.023, w: 0.412, h: 0.412, align: 'center', valign: 'middle', margin: 0 });
    });

  [[1.085, C.green, C.white, C.white, 1.408, 2.209, SH.wide],
  [7.407, C.white, C.black, C.gray, 7.73, 8.531, SH.lift]]
    .forEach(([x, bg, titleC, bodyC, bx, tx, shadow]) => {
      s.addShape(...['roundRect', pill({ x, y: 5.19, w: 4.83, h: 1.211, fill: { color: bg }, shadow })]);
      caption(s, { x: tx, y: 5.403, w: 2.633, h: 0.37 }, 'Description Title Here', { size: 16, color: titleC, align: 'center' });
      body(s, { x: bx, y: 5.725, w: 4.235, h: 0.505 },
        'PLACEHOLDER',
        { face: F.a, color: bodyC, align: 'center' });
    });
}

// 26 - get in touch ---------------------------------------------------
function slide26(s) {
  s.addShape('rect', { x: 0, y: 0, w: 6.458, h: 7.5, fill: { color: C.pale, transparency: 34 } });
  chevronBadge(s, 1.136, 1.149, 0.263, C.green);
  text(s, [run("Let's Get\u2019s Stated ", { fontFace: F.r, fontSize: 12, color: C.gray })],
    { x: 1.477, y: 1.149, w: 1.695, h: 0.303 });
  heading(s, { x: 0.972, y: 1.567, w: 4.377, h: 1.101, ls: 0.97 },
    [['Interview', C.green, F.pl, true], ['Get in Touch', C.black, F.ps]]);
  text(s, [
    run('Write your\u00a0', { fontFace: F.pl, fontSize: 18, color: C.black }),
    run('executive', br({ fontFace: F.ps, fontSize: 18, color: C.black })),
    run('Summary ', { fontFace: F.ps, fontSize: 18, color: C.black }),
    run('of where you want to take', { fontFace: F.pl, fontSize: 18, color: C.black }),
  ], { x: 1.024, y: 4.502, w: 3.3, h: 1.071, lineSpacingMultiple: 1.1 });
  button(s, { x: 1.112, y: 5.948, w: 1.338, h: 0.439, label: 'See More\u2026.', fill: C.green, color: C.white, face: F.p, size: 11, shadow: SH.drop });

  // highlighted middle contact row
  s.addShape(...['roundRect', pill({ x: 8.623, y: 3.937, w: 3.596, h: 1.292, fill: { type: 'none' }, line: { color: C.green, width: 1 } })]);
  const rows = [
    [2.779, 'pin', C.white, C.dark, ['Example Street Address,', 'City 038404, State'], SH.lift],
    [4.24, 'mail', C.green, C.white, ['lorem@example.com', 'www.example.com'], SH.wide],
    [5.708, 'phone', C.white, C.dark, ['(+123) 456 7890', '(+01) 092 029 029 029'], SH.lift],
  ];
  rows.forEach(([y, icon, bg, fg, lines, shadow]) => {
    iconBadge(s, 8.901, y, 0.679, { bg, fg, icon, square: true, shadow, scale: 0.45 });
    text(s, lines.map((t, i) => run(t, { fontFace: F.a, fontSize: 12, color: C.gray, breakLine: i === 0 })),
      { x: 9.938, y: y - 0.034, w: 2.356, h: 0.678, lineSpacingMultiple: 1.83 });
  });
}

// 27 - thank you ------------------------------------------------------
function slide27(s) {
  s.background = { color: C.white };
  s.addShape('rect', { x: 1.099, y: 1.099, w: 11.132, h: 5.302, fill: { color: C.cream } });
  squiggle(s, 10.332, 1.743, 0.734, 0.277);
  squiggle(s, 2.096, 5.54, 0.555, 0.209);
  ICON.arrowDown(s, 9.99, 3.012, 0.345, C.green);

  text(s, [run('Thank you', { fontFace: F.ps, fontSize: 88, color: C.dark })],
    { x: 3.38, y: 2.821, w: 7.403, h: 1.582 });

  s.addShape(...['roundRect', pill({
    x: 3.719, y: 4.149, w: 3.696, h: 0.428,
    fill: { color: C.white }, line: { color: C.green, width: 1 }, shadow: SH.wide,
  })]);
  text(s, [run('Presentation Template', { fontFace: F.i, fontSize: 12, color: C.gray })],
    { x: 3.862, y: 4.208, w: 3.543, h: 0.31, lineSpacingMultiple: 1.33 });
  s.addShape('ellipse', { x: 7.229, y: 4.177, w: 0.371, h: 0.371, fill: { color: C.dark } });
  ICON.magnifier(s, 7.336, 4.295, 0.157, C.white, 1.25);
}

/* ------------------------------------------------------------------ *
 * build                                                               *
 * ------------------------------------------------------------------ */

const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07,
  slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
  slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24, slide25,
  slide26, slide27];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE';
  pptx.author = 'pptxgenjs';
  pptx.title = 'Job Interview';

  BUILDERS.forEach((fn) => {
    const s = pptx.addSlide();
    s.background = { color: C.cream };
    fn(s);
  });

  return pptx.writeFile({
    fileName: path.join(__dirname, '1327501b-6de9-4f96-9f3b-729931362550_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
